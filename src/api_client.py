"""
lixinger API客户端封装
提供ETF数据获取的完整API封装
"""

import requests
import json
import time
from datetime import datetime, timedelta
from typing import List, Dict, Optional
import threading


class RateLimiter:
    """API频率限制控制器"""
    
    def __init__(self, rate_limit: float = 0.5):
        """
        初始化频率限制器
        rate_limit: 请求间隔时间（秒）
        """
        self.rate_limit = rate_limit
        self.last_request_time = 0
        self.lock = threading.Lock()
    
    def wait_if_needed(self):
        """等待以满足频率限制"""
        with self.lock:
            current_time = time.time()
            time_since_last = current_time - self.last_request_time
            
            if time_since_last < self.rate_limit:
                sleep_time = self.rate_limit - time_since_last
                time.sleep(sleep_time)
            
            self.last_request_time = time.time()


class LixingerAPIClient:
    """lixinger API客户端"""
    
    def __init__(self, token: str, base_url: str = "https://open.lixinger.com", 
                 timeout: int = 30, max_retries: int = 3, rate_limit: float = 0.5):
        """初始化API客户端"""
        self.token = token
        self.base_url = base_url.rstrip('/')
        self.timeout = timeout
        self.max_retries = max_retries
        self.rate_limiter = RateLimiter(rate_limit)
        
        # 创建请求会话
        self.session = requests.Session()
        self.session.headers.update({
            'Content-Type': 'application/json',
            'User-Agent': 'ETF-Excel-Tool/1.0 (macOS)'
        })
    
    def _make_request(self, endpoint: str, payload: Dict) -> Dict:
        """执行API请求"""
        # 应用频率限制
        self.rate_limiter.wait_if_needed()
        
        url = f"{self.base_url}{endpoint}"
        payload['token'] = self.token
        
        for attempt in range(self.max_retries):
            try:
                response = self.session.post(url, json=payload, timeout=self.timeout)
                
                if response.status_code == 200:
                    data = response.json()
                    if data.get('code') == 1:
                        return data
                    else:
                        raise Exception(f"API返回错误: {data.get('msg', '未知错误')}")
                else:
                    response.raise_for_status()
                    
            except requests.RequestException as e:
                if attempt == self.max_retries - 1:
                    raise Exception(f"请求失败: {str(e)}")
                
                # 指数退避重试
                wait_time = (2 ** attempt) * self.rate_limiter.rate_limit
                time.sleep(wait_time)
        
        raise Exception("请求重试次数达到上限")
    
    def get_etf_kline(self, stock_code: str, start_date: Optional[str] = None, 
                      end_date: Optional[str] = None) -> List[Dict]:
        """获取ETF K线数据"""
        
        if not end_date:
            end_date = datetime.now().strftime('%Y-%m-%d')
        
        if not start_date:
            # 默认获取最近5个交易日的数据
            start_date = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
        
        payload = {
            'stockCode': stock_code,
            'startDate': start_date,
            'endDate': end_date
        }
        
        try:
            response = self._make_request('/api/cn/fund/kline', payload)
            return response.get('data', [])
        except Exception as e:
            raise Exception(f"获取{stock_code}的K线数据失败: {str(e)}")
    
    def get_latest_price(self, stock_code: str) -> Optional[float]:
        """获取最新收盘价"""
        try:
            kline_data = self.get_etf_kline(stock_code)
            if kline_data:
                # 获取最新的收盘价
                latest = kline_data[-1]
                return float(latest.get('close', 0))
            return None
        except Exception as e:
            raise Exception(f"获取{stock_code}最新价格失败: {str(e)}")
    
    def get_batch_latest_prices(self, stock_codes: List[str]) -> Dict[str, Dict]:
        """批量获取最新收盘价"""
        results = {}
        
        for code in stock_codes:
            try:
                price = self.get_latest_price(code)
                results[code] = {
                    'price': price,
                    'status': 'success',
                    'update_time': datetime.now().isoformat()
                }
            except Exception as e:
                results[code] = {
                    'price': None,
                    'status': 'error',
                    'error': str(e),
                    'update_time': datetime.now().isoformat()
                }
        
        return results
    
    def test_connection(self) -> bool:
        """测试API连接"""
        try:
            # 使用一个常见的ETF代码测试
            self.get_etf_kline('159915')
            return True
        except Exception:
            return False
