"""
测试宽表格插入（10列，与 publish_to_feishu.py 一致）
"""
import requests
import json
import time

# 配置
APP_ID = "cli_aa8f28597a3adbc3"
APP_SECRET = "1bKmo7AfSheFJB4F9jufryHcCZlNer2K"
DOC_ID = "ZB67dzedbom0yMxbVcfc2AsKnIg"
FEISHU_BASE = "https://open.feishu.cn/open-apis"


def get_token():
    resp = requests.post(
        f"{FEISHU_BASE}/auth/v3/tenant_access_token/internal",
        json={"app_id": APP_ID, "app_secret": APP_SECRET}
    )
    return resp.json()["tenant_access_token"]


def get_root_block_count(token, doc_id):
    headers = {"Authorization": f"Bearer {token}"}
    all_items = []
    page_token = None

    while True:
        params = {"page_size": 500}
        if page_token:
            params["page_token"] = page_token

        resp = requests.get(
            f"{FEISHU_BASE}/docx/v1/documents/{doc_id}/blocks",
            headers=headers,
            params=params
        )
        resp.raise_for_status()
        data = resp.json()

        if data.get("code") != 0:
            raise Exception(f"获取块失败: {data}")

        all_items.extend(data["data"].get("items", []))

        if not data["data"].get("has_more"):
            break
        page_token = data["data"].get("page_token")

    root_blocks = [b for b in all_items if b.get("parent_id") == doc_id]
    return len(root_blocks)


def insert_table(token, doc_id, headers_row, data_rows, index):
    api_headers = {"Authorization": f"Bearer {token}"}
    n_cols = len(headers_row)
    n_rows = len(data_rows) + 1

    # 创建表格
    table_def = {
        "block_type": 31,
        "table": {
            "property": {
                "row_size": n_rows,
                "column_size": n_cols,
                "header_row": True,
            }
        }
    }

    print(f"[创建表格] index={index}, {n_rows}行 x {n_cols}列")
    print(f"  请求: {json.dumps(table_def, ensure_ascii=False)}")

    resp = requests.post(
        f"{FEISHU_BASE}/docx/v1/documents/{doc_id}/blocks/{doc_id}/children",
        headers=api_headers,
        json={"children": [table_def], "index": index}
    )

    print(f"  响应状态: {resp.status_code}")
    print(f"  响应内容: {resp.json()}")

    if resp.status_code != 200:
        return None

    data = resp.json()
    if data.get("code") != 0:
        return None

    children = data["data"].get("children", [])
    if not children:
        return None

    table_block_id = children[0]["block_id"]
    time.sleep(0.3)

    # 获取单元格
    resp = requests.get(
        f"{FEISHU_BASE}/docx/v1/documents/{doc_id}/blocks/{table_block_id}/children",
        headers=api_headers,
        params={"page_size": 500}
    )
    resp.raise_for_status()
    cell_blocks = resp.json()["data"].get("items", [])

    # 填充内容
    all_data = list(headers_row) + [v for row in data_rows for v in row]
    update_requests = []

    for idx, cell in enumerate(cell_blocks):
        if idx >= len(all_data):
            break
        cell_children = cell.get("children", [])
        if not cell_children:
            continue

        text_block_id = cell_children[0]
        content = str(all_data[idx]) if all_data[idx] is not None else ""
        is_header = idx < n_cols

        update_requests.append({
            "block_id": text_block_id,
            "update_text_elements": {
                "elements": [{
                    "text_run": {
                        "content": content,
                        "text_element_style": {"bold": True} if is_header else {}
                    }
                }]
            }
        })

    # 批量更新
    if update_requests:
        chunk_size = 50
        for i in range(0, len(update_requests), chunk_size):
            chunk = update_requests[i:i+chunk_size]
            resp = requests.patch(
                f"{FEISHU_BASE}/docx/v1/documents/{doc_id}/blocks/batch_update",
                headers=api_headers,
                json={"requests": chunk}
            )
            resp.raise_for_status()
            time.sleep(0.2)

    print(f"[成功] 表格插入完成")
    return table_block_id


def main():
    print("=" * 60)
    print("测试宽表格插入（10列）")
    print("=" * 60)

    token = get_token()
    print(f"\n[1] token 获取成功")

    block_count = get_root_block_count(token, DOC_ID)
    print(f"\n[2] 当前根级块数量: {block_count}")

    # 测试10列表格（与 publish_to_feishu.py 第一个表格一致）
    headers = ["供应商", "MTD目标", "实际例子", "例子gap", "例子达成率",
               "约课目标", "实际约课", "约课gap", "约课达成率", "约课成本"]
    rows = [
        ["钟嘉欣图片", "500", "429", "-71", "85.8%", "200", "167", "-33", "83.5%", "150.5"],
        ["周家蔚", "100", "1", "-99", "1.0%", "50", "0", "-50", "0.0%", "-"],
        ["钟嘉欣视频1", "300", "180", "-120", "60.0%", "120", "70", "-50", "58.3%", "200.0"],
        ["其他汇总", "200", "30", "-170", "15.0%", "80", "13", "-67", "16.3%", "180.0"],
        ["KOL-汇总", "1100", "640", "-460", "58.2%", "450", "250", "-200", "55.6%", "165.0"],
    ]

    print(f"\n[3] 插入10列x6行表格（index={block_count}）")
    table_id = insert_table(token, DOC_ID, headers, rows, block_count)

    if table_id:
        print(f"\n{'='*60}")
        print(f"[SUCCESS]")
        print(f"文档链接: https://feishu.cn/docx/{DOC_ID}")
        print(f"{'='*60}")
    else:
        print(f"\n{'='*60}")
        print(f"[FAILED]")
        print(f"{'='*60}")


if __name__ == "__main__":
    main()
