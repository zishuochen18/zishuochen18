import sys, json, asyncio
sys.path.insert(0, 'smartbi-data-cli-internal-20260526/smartbi-data-cli-internal-20260526/scripts')
from smartbi_browser_export import probe_simple_report_with_browser

async def debug_probe():
    result = await probe_simple_report_with_browser(
        username='63055',
        password='czs63055',
        report_id='I2c928087018de4d7e4d7f139018de8a92ab24ad1',
        max_rows=100
    )
    params = result.get('probe', {}).get('simple_report', {}).get('params', [])
    print('\n=== 销售明细报表参数列表 ===\n')
    for p in params:
        alias = p.get('alias', '(空)')
        name = p.get('name', '(空)')
        print(f"Alias: {alias:30} | Name: {name}")

asyncio.run(debug_probe())
