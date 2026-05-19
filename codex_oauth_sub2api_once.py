# -*- coding: utf-8 -*-
r"""单步测试 Codex OAuth 并导出 Sub2API 账号。

用途：绕开注册、支付、取消订阅等主流程，只验证：
1. 打开 Codex OAuth 授权页
2. 邮箱登录 / 一次性验证码登录
3. 获取 OAuth code 并换 token
4. 写入 Sub2API 格式 JSON

验证码通过文件注入，便于在外部控制台测试：
    Set-Content .\oauth_code.txt 123456
"""
import argparse
import asyncio
import os
from datetime import datetime

import gpt_free_core as bot_module
from gpt_free_core import ChatGPTBot, DEFAULT_OAUTH_PROXY, Sub2APIExporter, normalize_proxy_url


async def main() -> None:
    parser = argparse.ArgumentParser(description="单步测试 Codex OAuth -> Sub2API 导出")
    parser.add_argument("--email", required=True, help="OpenAI/ChatGPT 账号邮箱")
    parser.add_argument("--code-file", required=True, help="等待写入一次性验证码的文件路径")
    parser.add_argument("--output", default="", help="Sub2API 输出文件；指定时追加写入该文件，未指定时创建带当前时间后缀的新 JSON 文件")
    parser.add_argument("--headless", action="store_true", help="无头模式")
    parser.add_argument("--slow-mo", type=int, default=100, help="Playwright 操作延迟(ms)")
    parser.add_argument("--proxy", default=DEFAULT_OAUTH_PROXY, help="浏览器和 token exchange 共用代理；默认 http://127.0.0.1:7897")
    parser.add_argument("--debug", action="store_true", help="出错时保持浏览器打开")
    args = parser.parse_args()

    code_file = os.path.abspath(args.code_file)
    proxy_url = normalize_proxy_url(args.proxy)
    if args.output:
        Sub2APIExporter.FILE = os.path.abspath(args.output)
    else:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        Sub2APIExporter.FILE = os.path.abspath(f"sub2api-accounts-{timestamp}.json")
    bot_module.DEBUG_MODE = args.debug

    print("=" * 60)
    print("Codex OAuth -> Sub2API 单步测试")
    print(f"邮箱: {args.email}")
    print(f"验证码文件: {code_file}")
    print(f"输出文件: {Sub2APIExporter.FILE}")
    print(f"代理: {proxy_url or '未使用'}")
    print("收到验证码后，在另一个终端写入：")
    print(f"  Set-Content -Path \"{code_file}\" -Value <验证码>")
    print("=" * 60)

    async with ChatGPTBot(headless=args.headless, slow_mo=args.slow_mo, proxy_url=proxy_url) as bot:
        bot.oauth_code_file = code_file
        token_data = await bot.export_codex_sub2api_token(args.email)
        output_path = Sub2APIExporter.export_account(token_data, notes="Codex OAuth 单步测试导出")

    print("\n" + "=" * 60)
    print("Codex OAuth 导出完成")
    print(f"邮箱: {args.email}")
    print(f"Sub2API文件: {output_path}")
    print("=" * 60)


if __name__ == "__main__":
    asyncio.run(main())
