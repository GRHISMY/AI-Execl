"""
配置管理模块
负责API token和配置参数的管理
"""

import json
import os
from pathlib import Path
from typing import Dict, Optional


class ConfigManager:
    """配置管理类"""
    
    def __init__(self, config_path: Optional[str] = None):
        """初始化配置管理器"""
        if config_path:
            self.config_path = Path(config_path)
        else:
            # 默认配置文件路径
            self.config_path = Path.home() / '.etf_config.json'
        
        self.config = {}
        self.load_config()
    
    def load_config(self) -> bool:
        """加载配置文件"""
        try:
            if self.config_path.exists():
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
                return True
            else:
                self.config = self.get_default_config()
                return self.save_config()
        except Exception as e:
            print(f"加载配置文件失败: {e}")
            return False
    
    def save_config(self) -> bool:
        """保存配置文件"""
        try:
            # 确保目录存在
            self.config_path.parent.mkdir(parents=True, exist_ok=True)
            
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
            
            # 设置文件权限（仅用户可读写）
            os.chmod(self.config_path, 0o600)
            return True
        except Exception as e:
            print(f"保存配置文件失败: {e}")
            return False
    
    def get_default_config(self) -> Dict:
        """获取默认配置"""
        return {
            "api": {
                "base_url": "https://open.lixinger.com",
                "token": "",
                "timeout": 30,
                "max_retries": 3,
                "rate_limit": 0.5  # 每秒最多2次请求
            },
            "app": {
                "log_level": "INFO",
                "cache_enabled": True,
                "batch_size": 20
            }
        }
    
    def get(self, key: str, default=None):
        """获取配置值"""
        keys = key.split('.')
        value = self.config
        
        try:
            for k in keys:
                value = value[k]
            return value
        except (KeyError, TypeError):
            return default
    
    def set(self, key: str, value) -> bool:
        """设置配置值"""
        keys = key.split('.')
        config = self.config
        
        try:
            for k in keys[:-1]:
                if k not in config:
                    config[k] = {}
                config = config[k]
            
            config[keys[-1]] = value
            return self.save_config()
        except Exception as e:
            print(f"设置配置失败: {e}")
            return False
    
    def validate_config(self) -> bool:
        """验证配置有效性"""
        # 检查API token
        token = self.get('api.token')
        if not token or len(token.strip()) == 0:
            return False
        
        # 检查其他必要配置
        if not self.get('api.base_url'):
            return False
        
        return True
