# ETF价格获取工具 (Mac版)

基于lixinger API的Excel ETF收盘价获取工具，专为Mac环境优化，支持Intel和Apple Silicon架构。

## 功能特性

- ✅ 支持50+只ETF批量价格获取
- ✅ lixinger API集成，数据准确可靠
- ✅ Mac双架构支持（Intel/Apple Silicon）
- ✅ 用户零安装（PyInstaller打包）
- ✅ Excel VBA集成，操作简单
- ✅ 自动频率限制控制
- ✅ 完整错误处理机制

## 系统要求

- macOS 10.14+ (推荐 macOS 12+)
- Microsoft Excel for Mac 2019/2021/365
- lixinger API Token

## 快速开始

### 1. 获取API Token
1. 访问 [lixinger开放平台](https://open.lixinger.com/)
2. 注册账户并获取API Token

### 2. 下载和安装
1. 下载项目文件
2. 根据您的Mac架构选择对应的可执行文件：
   - Intel Mac: `etf_api_caller_intel`
   - Apple Silicon Mac: `etf_api_caller_arm64`

### 3. 配置和使用
1. 打开 `ETF_Price_Template.xlsx`
2. 点击"设置"按钮，输入API Token
3. 在A列输入ETF代码（如：159915, 159919）
4. 点击"刷新数据"获取价格信息

## 文件结构

```
ETF-Excel-Tool/
├── src/                           # Python源码
│   ├── etf_api_caller.py         # 主API调用脚本
│   ├── config_manager.py         # 配置管理
│   ├── api_client.py             # API客户端
│   └── utils.py                  # 工具函数
├── vba_modules/                   # VBA模块
│   ├── JsonConverter.bas         # JSON解析器
│   ├── Module_API.bas            # API接口
│   ├── Module_Config.bas         # 配置管理
│   ├── Module_Refresh.bas        # 数据刷新
│   └── ThisWorkbook.cls          # 事件处理
├── config/                        # 配置文件
│   ├── api_config.json           # API配置
│   └── settings.json             # 应用设置
├── build/                         # 构建脚本
│   ├── etf_caller.spec           # PyInstaller配置
│   ├── build_intel.sh            # Intel构建
│   └── build_arm64.sh            # ARM64构建
└── examples/                      # 示例文件
    └── ETF_Price_Template.xlsx    # Excel模板
```

## 开发和构建

### 环境准备
```bash
# 安装Python依赖
pip install -r requirements.txt

# 或手动安装
pip install requests pyinstaller
```

### 构建可执行文件

#### Intel Mac
```bash
cd build
chmod +x build_intel.sh
./build_intel.sh
```

#### Apple Silicon Mac
```bash
cd build
chmod +x build_arm64.sh
./build_arm64.sh
```

### 测试
```bash
# 测试Python脚本
python src/etf_api_caller.py --test --codes "159915"

# 测试可执行文件
./etf_api_caller_arm64 --test --codes "159915"
```

## 支持的ETF代码格式

- 深交所ETF: 159xxx (如: 159915, 159919)
- 上交所ETF: 510xxx, 515xxx (如: 510300, 515000)
- 其他ETF: 560xxx-569xxx

## API限制

- 请求频率: 每秒最多2次
- 自动重试: 最多3次
- 超时时间: 30秒

## 故障排除

### 常见问题

1. **"未找到API可执行文件"**
   - 确保可执行文件在正确位置
   - 检查文件权限: `chmod +x etf_api_caller_*`

2. **"需要配置API Token"**
   - 点击"设置"按钮输入有效的lixinger API Token

3. **macOS安全警告**
   - 系统偏好设置 > 安全性与隐私 > 通用 > 允许应用

4. **API调用失败**
   - 检查网络连接
   - 验证API Token有效性
   - 确认ETF代码格式正确

### 日志和调试
- VBA错误信息会显示在Excel的即时窗口
- Python错误信息通过JSON响应返回
- 可使用 `--verbose` 参数获取详细日志

## 许可证

本项目采用MIT许可证，详见LICENSE文件。

## 联系方式

如有问题或建议，请提交Issue或Pull Request。

---

**注意**: 本工具仅供学习和研究使用，投资决策请谨慎考虑风险。
