"""
飞书多维表格轮询 Worker

使用方法：
   1. 复制 feishu_config.example.json 为 feishu_config.local.json，填入飞书应用凭证
   2. 运行：python feishu_bitable_worker.py

说明：
- 定时轮询飞书多维表格，发现"待执行"的行就自动运行 CRM 自动化
- 执行完成后回填推广链接、二维码图片、执行状态
- 需要飞书应用具有 bitable:app 和 drive:drive 权限
- 需要将飞书应用添加为多维表格的协作者（可编辑）
"""

import os
import sys
import json
import time
import requests
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright

from crm_create_channel import ensure_login, run_single

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "feishu_config.local.json")
FEISHU_BASE = "https://open.feishu.cn/open-apis"


class FeishuClient:
    """飞书 API 客户端，封装 token 管理和常用接口"""

    def __init__(self, app_id: str, app_secret: str):
        self.app_id = app_id
        self.app_secret = app_secret
        self._token = None
        self._token_expires_at = None

    def _get_tenant_token(self) -> str:
        now = datetime.now()
        if self._token and self._token_expires_at and now < self._token_expires_at:
            return self._token

        resp = requests.post(
            f"{FEISHU_BASE}/auth/v3/tenant_access_token/internal",
            json={"app_id": self.app_id, "app_secret": self.app_secret}
        )
        data = resp.json()
        if data.get("code") != 0:
            raise RuntimeError(f"获取 token 失败：{data}")

        self._token = data["tenant_access_token"]
        self._token_expires_at = now + timedelta(seconds=data.get("expire", 7200) - 300)
        return self._token

    def _headers(self) -> dict:
        return {"Authorization": f"Bearer {self._get_tenant_token()}"}

    def list_records(self, app_token: str, table_id: str) -> list:
        """读取多维表格记录"""
        url = f"{FEISHU_BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records"
        params = {"page_size": 100}
        all_records = []
        page_token = None

        while True:
            if page_token:
                params["page_token"] = page_token
            resp = requests.get(url, headers=self._headers(), params=params)
            data = resp.json()
            if data.get("code") != 0:
                raise RuntimeError(f"读取记录失败：{data}")
            items = data.get("data", {}).get("items", [])
            all_records.extend(items)
            if not data.get("data", {}).get("has_more"):
                break
            page_token = data["data"].get("page_token")
        return all_records

    def update_record(self, app_token: str, table_id: str, record_id: str, fields: dict):
        """更新单条记录"""
        url = f"{FEISHU_BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records/{record_id}"
        resp = requests.put(url, headers=self._headers(), json={"fields": fields})
        data = resp.json()
        if data.get("code") != 0:
            raise RuntimeError(f"更新记录失败：{data}")
        return data

    def upload_media(self, file_path: str, file_name: str, app_token: str) -> str:
        """上传文件到飞书 Drive，返回 file_token"""
        url = f"{FEISHU_BASE}/drive/v1/medias/upload_all"
        file_size = os.path.getsize(file_path)
        with open(file_path, "rb") as f:
            resp = requests.post(
                url,
                headers=self._headers(),
                data={
                    "file_name": file_name,
                    "parent_type": "bitable_image",
                    "parent_node": app_token,
                    "size": str(file_size),
                },
                files={"file": (file_name, f, "image/jpeg")}
            )
        data = resp.json()
        if data.get("code") != 0:
            raise RuntimeError(f"上传文件失败：{data}")
        return data["data"]["file_token"]


def load_config() -> dict:
    if not os.path.exists(CONFIG_FILE):
        print(f"❌ 配置文件不存在：{CONFIG_FILE}")
        print(f"   请复制 feishu_config.example.json 为 feishu_config.local.json 并填入真实值")
        sys.exit(1)
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def process_record(client, config, record, browser, context, page):
    """处理单条待执行记录"""
    app_token = config["app_token"]
    table_id = config["table_id"]
    fields = config["field_names"]
    record_id = record["record_id"]
    record_fields = record.get("fields", {})
    channel_name = record_fields.get(fields["channel_name"], "")

    if isinstance(channel_name, list):
        channel_name = "".join(seg.get("text", "") for seg in channel_name)
    channel_name = str(channel_name).strip()
    if not channel_name:
        return

    print(f"\n{'='*50}")
    print(f">>> 开始处理：{channel_name}")
    print(f"{'='*50}")

    client.update_record(app_token, table_id, record_id, {
        fields["status"]: "执行中"
    })

    try:
        promo_link, qr_file = run_single(
            channel_name, browser=browser, context=context, page=page
        )
        update_fields = {fields["status"]: "成功"}
        if promo_link:
            update_fields[fields["promo_link"]] = promo_link
        if qr_file and os.path.exists(qr_file):
            file_name = os.path.basename(qr_file)
            file_token = client.upload_media(qr_file, file_name, app_token)
            update_fields[fields["qr_image"]] = [{"file_token": file_token}]

        client.update_record(app_token, table_id, record_id, update_fields)
        print(f">>> ✓ 处理完成：{channel_name}")
    except Exception as e:
        print(f">>> ❌ 处理失败：{channel_name} - {e}")
        try:
            client.update_record(app_token, table_id, record_id, {
                fields["status"]: "失败"
            })
        except Exception:
            pass


def main():
    config = load_config()
    client = FeishuClient(config["app_id"], config["app_secret"])
    poll_interval = config.get("poll_interval_seconds", 30)
    fields = config["field_names"]

    print("=" * 50)
    print("  飞书多维表格 → CRM 自动化 Worker")
    print("=" * 50)
    print(f">>> 轮询间隔：{poll_interval} 秒")

    pw = sync_playwright().start()
    browser = pw.chromium.launch(headless=False)
    context, page = ensure_login(browser)
    print(">>> 浏览器已就绪，开始轮询...\n")

    try:
        while True:
            try:
                records = client.list_records(config["app_token"], config["table_id"])
                pending = [
                    r for r in records
                    if r.get("fields", {}).get(fields["channel_name"])
                    and (not r.get("fields", {}).get(fields["status"])
                         or r.get("fields", {}).get(fields["status"]) == "待执行")
                ]
                if pending:
                    print(f">>> 发现 {len(pending)} 条待执行任务")
                    for record in pending:
                        process_record(client, config, record, browser, context, page)
                else:
                    now = datetime.now().strftime("%H:%M:%S")
                    print(f"[{now}] 无待执行任务，{poll_interval}秒后重新检查...", end="\r")
            except KeyboardInterrupt:
                raise
            except Exception as e:
                print(f"\n>>> ⚠️ 轮询出错：{e}，{poll_interval}秒后重试...")
            time.sleep(poll_interval)
    except KeyboardInterrupt:
        print("\n\n>>> Worker 已停止")
    finally:
        browser.close()
        pw.stop()


if __name__ == "__main__":
    main()
