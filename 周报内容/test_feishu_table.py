"""
测试飞书原生表格插入（修复 1770001 错误）

核心修复：动态获取文档块数量，用作表格插入的索引
"""
import requests
import json
import time

# 配置
APP_ID = "cli_aa8f28597a3adbc3"
APP_SECRET = "1bKmo7AfSheFJB4F9jufryHcCZlNer2K"
DOC_ID = "ZB67dzedbom0yMxbVcfc2AsKnIg"  # 测试文档ID
FEISHU_BASE = "https://open.feishu.cn/open-apis"


def get_token():
    """获取 access token"""
    resp = requests.post(
        f"{FEISHU_BASE}/auth/v3/tenant_access_token/internal",
        json={"app_id": APP_ID, "app_secret": APP_SECRET}
    )
    return resp.json()["tenant_access_token"]


def get_root_block_count(token, doc_id):
    """获取文档根级块数量（关键函数）"""
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

    # 统计根级块（parent_id == doc_id）
    root_blocks = [b for b in all_items if b.get("parent_id") == doc_id]
    return len(root_blocks)


def insert_text_block(token, doc_id, text, index):
    """插入文本块"""
    headers = {"Authorization": f"Bearer {token}"}
    block = {
        "block_type": 2,  # text paragraph
        "text": {
            "elements": [{"text_run": {"content": text}}],
            "style": {}
        }
    }

    resp = requests.post(
        f"{FEISHU_BASE}/docx/v1/documents/{doc_id}/blocks/{doc_id}/children",
        headers=headers,
        json={"children": [block], "index": index}
    )
    resp.raise_for_status()
    return resp.json()


def insert_table(token, doc_id, headers_row, data_rows, index):
    """插入原生表格（使用动态索引）"""
    api_headers = {"Authorization": f"Bearer {token}"}
    n_cols = len(headers_row)
    n_rows = len(data_rows) + 1  # 含表头

    # 1. 创建表格块
    table_def = {
        "block_type": 31,  # table
        "table": {
            "property": {
                "row_size": n_rows,
                "column_size": n_cols,
                "header_row": True,
            }
        }
    }

    print(f"  [创建表格] index={index}, {n_rows}行 x {n_cols}列")

    resp = requests.post(
        f"{FEISHU_BASE}/docx/v1/documents/{doc_id}/blocks/{doc_id}/children",
        headers=api_headers,
        json={"children": [table_def], "index": index}
    )

    if resp.status_code != 200:
        print(f"  ❌ 失败: {resp.status_code}")
        print(f"  响应: {resp.json()}")
        return None

    data = resp.json()
    if data.get("code") != 0:
        print(f"  ❌ API 错误: {data}")
        return None

    print(f"  ✅ 表格创建成功")

    children = data["data"].get("children", [])
    if not children:
        return None

    table_block_id = children[0]["block_id"]
    time.sleep(0.3)

    # 2. 获取表格单元格
    resp = requests.get(
        f"{FEISHU_BASE}/docx/v1/documents/{doc_id}/blocks/{table_block_id}/children",
        headers=api_headers,
        params={"page_size": 500}
    )
    resp.raise_for_status()
    cell_blocks = resp.json()["data"].get("items", [])

    # 3. 填充单元格内容
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

    # 4. 批量更新单元格
    if update_requests:
        # 分批更新（每批50个）
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

    print(f"  ✅ 表格内容填充完成")
    return table_block_id


def main():
    print("=" * 60)
    print("飞书原生表格插入测试")
    print("=" * 60)

    # 1. 获取 token
    print("\n[1] 获取 token...")
    token = get_token()
    print(f"  [OK] token 获取成功")

    # 2. 获取当前文档块数量
    print(f"\n[2] 获取文档根级块数量...")
    block_count = get_root_block_count(token, DOC_ID)
    print(f"  当前根级块数量: {block_count}")

    # 3. 插入一段文本
    print(f"\n[3] 插入测试文本（index={block_count}）...")
    insert_text_block(token, DOC_ID, "测试表格插入（修复 1770001 错误）", block_count)
    print(f"  [OK] 文本插入成功")
    time.sleep(0.5)

    # 4. 重新获取块数量
    block_count = get_root_block_count(token, DOC_ID)
    print(f"  更新后根级块数量: {block_count}")

    # 5. 插入表格（使用动态索引）
    print(f"\n[4] 插入表格（index={block_count}）...")
    headers = ["供应商", "例子数", "约课数", "转化率", "ROI"]
    rows = [
        ["钟嘉欣", "429", "167", "7.93%", "0.49"],
        ["周家蔚", "1", "0", "0.0%", "0.00"],
        ["KOLHK汇总", "640", "250", "10.4%", "0.55"],
    ]

    table_id = insert_table(token, DOC_ID, headers, rows, block_count)

    if table_id:
        print(f"\n{'='*60}")
        print(f"[SUCCESS] 测试成功！")
        print(f"文档链接: https://feishu.cn/docx/{DOC_ID}")
        print(f"{'='*60}")
    else:
        print(f"\n{'='*60}")
        print(f"[FAILED] 测试失败")
        print(f"{'='*60}")


if __name__ == "__main__":
    main()
