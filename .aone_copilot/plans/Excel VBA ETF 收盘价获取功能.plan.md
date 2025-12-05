### Excel VBA ETF 收盘价获取功能 ###
创建一个包含 VBA 宏的 Excel 文件，能够根据用户在第一列输入的 ETF 代码，自动调用理想财经 API 获取最新收盘价并填充到第二列，在打开文件时自动刷新数据。


## 实施计划

### 一、Excel 文件结构设计

1. **创建 Excel 工作簿**
   - 在 AI-Execl 目录下创建 `ETF_Price_Tracker.xlsm`（启用宏的 Excel 文件）
   - 设置工作表名称为 "ETF价格"

2. **设置表格结构**
   - A列：ETF代码（用户手动输入，如：512690）
   - B列：最新收盘价（通过 VBA 自动填充）
   - C列：数据日期（可选，显示收盘价对应的日期）
   - 第1行作为表头：ETF代码 | 最新收盘价 | 数据日期

### 二、VBA 模块开发

#### 模块1：配置模块 (Module_Config)

1. **创建全局配置**
   - API 基础 URL：`https://open.lixinger.com/api/cn/fund/candlestick`
   - API Token：需要用户配置（可以存储在隐藏工作表或配置单元格中）
   - 请求超时时间：30秒
   - API 频率限制控制

#### 模块2：API 调用模块 (Module_API)

1. **实现 HTTP 请求函数**
   - 使用 `MSXML2.XMLHTTP` 或 `WinHttp.WinHttpRequest` 对象
   - 构造 POST 请求，发送 JSON 格式的数据
   - 参数包含：token, stockCode, startDate, endDate

2. **实现 JSON 解析函数**
   - 使用 VBA-JSON 库或自定义解析函数
   - 解析 API 返回的 JSON 数据结构
   - 提取 data 数组中最新一条记录的 close 字段（收盘价）

3. **实现获取最新收盘价函数**
   - 函数名：`GetLatestClosePrice(etfCode As String) As Variant`
   - 调用 API 获取最近5个交易日的数据（避免节假日问题）
   - 从返回的数据中选择最新日期的收盘价
   - 添加错误处理机制（网络错误、API 错误等）

4. **实现频率限制控制**
   - 参考 manager.py 中的 `_wait_for_api_rate_limit` 逻辑
   - 在请求之间添加延迟，避免触发 API 频率限制

#### 模块3：数据刷新模块 (Module_Refresh)

1. **实现批量刷新函数**
   - 函数名：`RefreshAllPrices()`
   - 遍历 A列所有包含 ETF 代码的单元格（从第2行开始）
   - 跳过空单元格
   - 为每个 ETF 代码调用 API 获取收盘价
   - 将结果写入对应的 B列单元格
   - 将数据日期写入对应的 C列单元格
   - 显示进度提示（可选）

2. **添加手动刷新按钮**
   - 在工作表中添加一个按钮控件
   - 绑定到 `RefreshAllPrices()` 函数
   - 按钮文字：刷新价格

#### 模块4：自动刷新模块 (ThisWorkbook)

1. **实现自动刷新事件**
   - 在 `ThisWorkbook` 代码窗口中添加 `Workbook_Open` 事件
   - 打开文件时自动调用 `RefreshAllPrices()` 函数
   - 添加用户提示（可选）："正在刷新ETF价格数据..."

2. **添加保护机制**
   - 检查是否有网络连接
   - 检查 Token 是否配置
   - 如果配置不完整，提示用户配置

### 三、辅助功能

1. **配置界面**
   - 创建一个配置工作表（可隐藏）
   - 提供 Token 输入单元格
   - 提供配置说明文档链接

2. **错误处理**
   - 网络连接失败处理
   - API 返回错误处理（如：ETF 代码不存在）
   - 显示友好的错误信息
   - 对于失败的单元格，显示错误提示（如：#N/A 或 "错误"）

3. **数据验证**
   - 对 A列添加数据验证
   - 限制只能输入6位数字（ETF 代码格式）
   - 添加输入提示

### 四、依赖项处理

1. **JSON 解析库**
   - 需要导入 VBA-JSON 库（Tim Hall 的开源库）
   - 或实现简单的 JSON 解析函数（针对理想财经 API 的特定格式）

2. **引用设置**
   - 启用 Microsoft Scripting Runtime（用于字典对象）
   - 启用 Microsoft XML, v6.0（用于 HTTP 请求）

### 五、测试和优化

1. **测试场景**
   - 测试单个 ETF 代码
   - 测试多个 ETF 代码批量刷新
   - 测试错误情况（无效代码、网络断开等）
   - 测试打开文件自动刷新

2. **性能优化**
   - 使用 Application.ScreenUpdating = False 提高更新速度
   - 批量更新单元格而非逐个更新
   - 合理设置 API 调用间隔

3. **用户体验优化**
   - 添加进度条或状态栏提示
   - 冻结首行（表头）
   - 设置合适的列宽
   - 添加条件格式（如：价格上涨显示绿色，下跌显示红色）

### 六、文档和说明

1. **使用说明**
   - 在 README.md 中添加使用指南
   - 说明如何获取理想财经 API Token
   - 说明如何配置 Token
   - 说明如何使用 Excel 文件

2. **注意事项**
   - 提醒用户启用宏
   - 提醒用户配置 Token
   - 提醒用户注意 API 使用频率限制


updateAtTime: 2025/12/5 23:41:35

planId: e5320253-0763-4cbf-b599-95e6a16e57a5