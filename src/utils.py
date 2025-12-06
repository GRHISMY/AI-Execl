"""
工具函数模块
提供日期处理、数据验证等通用功能
"""

import re
import json
import logging
import platform
from datetime import datetime, timedelta
from typing import List, Optional, Dict
from pathlib import Path


def setup_logging(level: str = "INFO") -> logging.Logger:
    """设置日志"""
    logger = logging.getLogger('etf_tool')
    logger.setLevel(getattr(logging, level.upper()))
    
    if not logger.handlers:
        handler = logging.StreamHandler()
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    
    return logger


def validate_etf_code(code: str) -> bool:
    """验证ETF代码格式"""
    if not code or not isinstance(code, str):
        return False
    
    code = code.strip()
    
    # A股ETF代码格式: 6位数字
    if re.match(r'^\d{6}$', code):
        # 简单验证ETF代码范围
        code_int = int(code)
        # 深交所ETF: 159xxx, 上交所ETF: 510xxx, 515xxx等
        if (159000 <= code_int <= 159999 or 
            510000 <= code_int <= 519999 or 
            560000 <= code_int <= 569999):
            return True
    
    return False


def parse_etf_codes(codes_str: str) -> List[str]:
    """解析ETF代码字符串"""
    if not codes_str:
        return []
    
    # 支持逗号、分号、空格分隔
    codes = re.split(r'[,;,\s]+', codes_str.strip())
    
    # 验证并过滤有效代码
    valid_codes = []
    for code in codes:
        code = code.strip()
        if validate_etf_code(code):
            valid_codes.append(code)
    
    # 去重
    return list(set(valid_codes))


def get_trading_date(offset_days: int = 0) -> str:
    """获取交易日期"""
    target_date = datetime.now() + timedelta(days=offset_days)
    
    # 简单处理：如果是周末，调整到最近的工作日
    weekday = target_date.weekday()
    
    if weekday == 5:  # 周六
        target_date = target_date - timedelta(days=1)
    elif weekday == 6:  # 周日
        target_date = target_date - timedelta(days=2)
    
    return target_date.strftime('%Y-%m-%d')


def format_price(price: Optional[float], decimals: int = 3) -> str:
    """格式化价格显示"""
    if price is None:
        return "N/A"
    
    try:
        return f"{float(price):.{decimals}f}"
    except (ValueError, TypeError):
        return "N/A"


def get_system_info() -> Dict[str, str]:
    """获取系统信息"""
    return {
        'platform': platform.platform(),
        'machine': platform.machine(),
        'python_version': platform.python_version(),
        'is_apple_silicon': platform.machine() == 'arm64'
    }


def find_executable_path() -> Optional[str]:
    """查找可执行文件路径"""
    # 可能的路径列表
    possible_paths = [
        # 同级目录
        './etf_api_caller',
        # 相对路径
        '../dist/etf_api_caller',
        './dist/etf_api_caller',
        # 用户目录
        str(Path.home() / 'etf_api_caller'),
        # 应用程序目录
        '/Applications/ETF-Tool/etf_api_caller'
    ]
    
    # 根据架构选择对应的可执行文件
    system_info = get_system_info()
    if system_info['is_apple_silicon']:
        arm64_paths = [path + '_arm64' for path in possible_paths]
        possible_paths = arm64_paths + possible_paths
    
    for path in possible_paths:
        if Path(path).exists() and Path(path).is_file():
            return str(Path(path).resolve())
    
    return None


def create_output_json(data: Dict, status: str = "success") -> str:
    """创建标准化的JSON输出"""
    output = {
        'status': status,
        'timestamp': datetime.now().isoformat(),
        'data': data
    }
    
    if status == "error":
        output['error'] = data
        output['data'] = {}
    
    return json.dumps(output, ensure_ascii=False, indent=2)


def create_output_csv(data: Dict) -> str:
    """创建CSV格式输出"""
    lines = ['ETF代码,收盘价,状态,更新时间']
    
    for code, info in data.items():
        price = format_price(info.get('price'))
        status = info.get('status', 'unknown')
        update_time = info.get('update_time', '')
        
        lines.append(f"{code},{price},{status},{update_time}")
    
    return '\n'.join(lines)


def safe_file_write(file_path: str, content: str, encoding: str = 'utf-8') -> bool:
    """安全写入文件"""
    try:
        path = Path(file_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        
        with open(path, 'w', encoding=encoding) as f:
            f.write(content)
        
        return True
    except Exception as e:
        print(f"写入文件失败 {file_path}: {e}")
        return False
