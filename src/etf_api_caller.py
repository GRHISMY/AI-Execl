#!/usr/bin/env python3
"""
ETF API调用脚本 - Mac版本
支持通过lixinger API批量获取ETF收盘价数据
"""

import argparse
import json
import sys
import os
from pathlib import Path

# 添加当前目录到Python路径
sys.path.insert(0, str(Path(__file__).parent))

from config_manager import ConfigManager
from api_client import LixingerAPIClient
from utils import (
    parse_etf_codes, create_output_json, create_output_csv,
    setup_logging, get_system_info
)


def main():
    """主入口函数"""
    parser = argparse.ArgumentParser(description='ETF价格API调用工具')
    parser.add_argument('--codes', required=True, help='ETF代码列表，逗号分隔')
    parser.add_argument('--token', help='API token')
    parser.add_argument('--config', help='配置文件路径')
    parser.add_argument('--output', default='json', choices=['json', 'csv'], help='输出格式')
    parser.add_argument('--test', action='store_true', help='测试API连接')
    parser.add_argument('--verbose', action='store_true', help='详细输出')

    args = parser.parse_args()

    # 设置日志
    log_level = "DEBUG" if args.verbose else "INFO"
    logger = setup_logging(log_level)

    try:
        # 初始化配置管理器
        config_manager = ConfigManager(args.config)

        # 获取API token
        api_token = args.token or config_manager.get('api.token')
        if not api_token:
            error_msg = "未提供API token，请使用 --token 参数或配置文件"
            print(create_output_json(error_msg, "error"))
            sys.exit(1)

        # 解析ETF代码
        etf_codes = parse_etf_codes(args.codes)

        if not etf_codes:
            error_msg = "未提供有效的ETF代码"
            print(create_output_json(error_msg, "error"))
            sys.exit(1)

        logger.info(f"处理 {len(etf_codes)} 只ETF: {etf_codes}")

        # 初始化API客户端
        api_client = LixingerAPIClient(
            token=api_token,
            base_url=config_manager.get('api.base_url'),
            timeout=config_manager.get('api.timeout'),
            max_retries=config_manager.get('api.max_retries'),
            rate_limit=config_manager.get('api.rate_limit')
        )

        # 测试连接
        if args.test:
            if api_client.test_connection():
                result = {"message": "API连接测试成功", "system_info": get_system_info()}
                print(create_output_json(result))
            else:
                error_msg = "API连接测试失败"
                print(create_output_json(error_msg, "error"))
                sys.exit(1)
            return

        # 批量获取价格数据
        logger.info("开始获取ETF价格数据...")
        results = api_client.get_batch_latest_prices(etf_codes)

        # 输出结果
        if args.output == 'csv':
            print(create_output_csv(results))
        else:
            print(create_output_json(results))

        # 统计结果
        success_count = sum(1 for r in results.values() if r['status'] == 'success')
        logger.info(f"处理完成: 成功 {success_count}/{len(etf_codes)}")

    except Exception as e:
        logger.error(f"程序执行失败: {e}")
        error_msg = f"程序执行失败: {str(e)}"
        print(create_output_json(error_msg, "error"))
        sys.exit(1)


if __name__ == "__main__":
    main()
