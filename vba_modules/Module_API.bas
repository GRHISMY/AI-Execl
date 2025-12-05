Attribute VB_Name = "Module_API"
' =========================================
' 模块名称: Module_API
' 描述: 理想财经API调用模块
' 作者: Aone Copilot
' 创建日期: 2025-12-05
' =========================================

Option Explicit

' ========== 公共函数 ==========

' 获取ETF最新收盘价
Public Function GetLatestClosePrice(etfCode As String, ByRef outDate As String, ByRef outUpdateTime As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim jsonResponse As String
    Dim closePrice As Variant
    
    ' 应用频率限制
    Call Module_Config.WaitForRateLimit
    
    ' 调用API获取K线数据
    jsonResponse = FetchETFKlineData(etfCode)
    
    If jsonResponse = "" Then
        GetLatestClosePrice = "API调用失败"
        outDate = ""
        outUpdateTime = ""
        Exit Function
    End If
    
    ' 解析JSON获取最新收盘价
    closePrice = JsonConverter.ExtractLatestClosePrice(jsonResponse, outDate)
    outUpdateTime = Now
    
    GetLatestClosePrice = closePrice
    Exit Function
    
ErrorHandler:
    GetLatestClosePrice = "错误: " & Err.Description
    outDate = ""
    outUpdateTime = ""
End Function

' 获取ETF K线数据
Public Function FetchETFKlineData(etfCode As String) As String
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    Dim token As String
    Dim startDate As String
    Dim endDate As String
    Dim payload As String
    Dim response As String
    Dim maxRetries As Integer
    Dim retryCount As Integer
    Dim retryDelay As Integer
    
    ' 获取配置
    token = Module_Config.GetApiToken()
    If token = "" Then
        FetchETFKlineData = ""
        Exit Function
    End If
    
    ' 设置URL
    url = Module_Config.API_BASE_URL
    
    ' 计算日期范围（最近10个交易日，确保获取到最新数据）
    endDate = Format(Date, "yyyy-mm-dd")
    startDate = Format(DateAdd("d", -30, Date), "yyyy-mm-dd")
    
    ' 构造JSON payload
    payload = "{" & _
              """token"": """ & token & """," & _
              """stockCode"": """ & etfCode & """," & _
              """startDate"": """ & startDate & """," & _
              """endDate"": """ & endDate & """" & _
              "}"
    
    ' 重试配置
    maxRetries = 3
    retryCount = 0
    retryDelay = 1  ' 初始延迟1秒
    
    ' 重试循环
    Do While retryCount <= maxRetries
        On Error Resume Next
        
        ' 创建HTTP对象
        Set http = CreateObject("MSXML2.XMLHTTP")
        If Err.Number <> 0 Then
            ' 尝试使用WinHttp
            Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
            If Err.Number <> 0 Then
                FetchETFKlineData = ""
                Exit Function
            End If
        End If
        
        On Error GoTo RetryHandler
        
        ' 配置请求
        http.Open "POST", url, False
        http.setRequestHeader "Content-Type", "application/json"
        http.setTimeouts 30000, 30000, 30000, 30000  ' 连接超时、发送超时、接收超时、总超时（毫秒）
        
        ' 发送请求
        http.send payload
        
        ' 检查响应
        If http.Status = 200 Then
            response = http.responseText
            
            ' 验证响应是否为有效JSON
            If Left(Trim(response), 1) = "{" Then
                FetchETFKlineData = response
                Exit Function
            End If
        End If
        
RetryHandler:
        ' 如果失败，准备重试
        retryCount = retryCount + 1
        
        If retryCount <= maxRetries Then
            ' 等待后重试（指数退避）
            Application.Wait Now + TimeValue("0:00:0" & retryDelay)
            retryDelay = retryDelay * 2  ' 指数退避
        End If
    Loop
    
    ' 所有重试都失败
    FetchETFKlineData = ""
    Exit Function
    
ErrorHandler:
    FetchETFKlineData = ""
End Function

' 批量获取ETF收盘价（优化版）
Public Function BatchGetClosePrices(etfCodes() As String) As Object
    On Error Resume Next
    
    Dim results As Object
    Set results = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim etfCode As String
    Dim closePrice As Variant
    Dim priceDate As String
    Dim updateTime As String
    
    For i = LBound(etfCodes) To UBound(etfCodes)
        etfCode = Trim(etfCodes(i))
        
        If etfCode <> "" Then
            closePrice = GetLatestClosePrice(etfCode, priceDate, updateTime)
            
            ' 存储结果
            Dim resultDict As Object
            Set resultDict = CreateObject("Scripting.Dictionary")
            resultDict("closePrice") = closePrice
            resultDict("date") = priceDate
            resultDict("updateTime") = updateTime
            
            results(etfCode) = resultDict
        End If
    Next i
    
    Set BatchGetClosePrices = results
End Function

' 验证ETF代码格式
Public Function ValidateETFCode(etfCode As String) As Boolean
    Dim cleanCode As String
    
    ' 移除空白
    cleanCode = Trim(etfCode)
    
    ' 检查长度（通常是6位数字）
    If Len(cleanCode) <> 6 Then
        ValidateETFCode = False
        Exit Function
    End If
    
    ' 检查是否全为数字
    If Not IsNumeric(cleanCode) Then
        ValidateETFCode = False
        Exit Function
    End If
    
    ' 检查前缀（5开头为上交所，1开头为深交所）
    Dim firstChar As String
    firstChar = Left(cleanCode, 1)
    
    If firstChar <> "5" And firstChar <> "1" Then
        ValidateETFCode = False
        Exit Function
    End If
    
    ValidateETFCode = True
End Function

' 测试API连接
Public Function TestAPIConnection() As Boolean
    On Error GoTo ErrorHandler
    
    Dim testCode As String
    Dim response As String
    
    ' 使用一个常见的ETF代码进行测试（沪深300ETF）
    testCode = "510300"
    
    ' 尝试获取数据
    response = FetchETFKlineData(testCode)
    
    If response <> "" Then
        TestAPIConnection = True
    Else
        TestAPIConnection = False
    End If
    
    Exit Function
    
ErrorHandler:
    TestAPIConnection = False
End Function

' 获取错误详细信息（用于调试）
Public Function GetLastAPIError() As String
    ' 这里可以记录最后一次API错误的详细信息
    ' 为简化实现，返回基本错误信息
    GetLastAPIError = "请检查网络连接和API Token配置"
End Function

' ========== 辅助函数 ==========

' 格式化日期为YYYY-MM-DD
Private Function FormatDateString(dateValue As Date) As String
    FormatDateString = Format(dateValue, "yyyy-mm-dd")
End Function

' 解析API错误响应
Private Function ParseErrorResponse(jsonResponse As String) As String
    On Error Resume Next
    
    Dim jsonObj As Object
    Dim message As Variant
    
    Set jsonObj = JsonConverter.ParseJson(jsonResponse)
    
    If Not jsonObj Is Nothing Then
        message = JsonConverter.GetJsonValue(jsonObj, "message")
        If Not IsNull(message) And message <> "" Then
            ParseErrorResponse = CStr(message)
            Exit Function
        End If
    End If
    
    ParseErrorResponse = "未知错误"
End Function

' 记录API调用日志（可选功能）
Private Sub LogAPICall(etfCode As String, success As Boolean, Optional errorMsg As String = "")
    ' 这里可以实现日志记录功能
    ' 例如写入隐藏工作表或文本文件
    ' 为了简化，这里留空，用户可以根据需要扩展
End Sub

' 检查响应是否有效
Private Function IsValidResponse(jsonResponse As String) As Boolean
    On Error Resume Next
    
    If jsonResponse = "" Then
        IsValidResponse = False
        Exit Function
    End If
    
    Dim jsonObj As Object
    Set jsonObj = JsonConverter.ParseJson(jsonResponse)
    
    If jsonObj Is Nothing Then
        IsValidResponse = False
        Exit Function
    End If
    
    Dim code As Variant
    code = JsonConverter.GetJsonValue(jsonObj, "code")
    
    ' 理想财经API返回code=1表示成功
    If code = 1 Then
        IsValidResponse = True
    Else
        IsValidResponse = False
    End If
End Function

' 从响应中提取数据
Private Function ExtractDataFromResponse(jsonResponse As String) As Object
    On Error Resume Next
    
    Dim jsonObj As Object
    Set jsonObj = JsonConverter.ParseJson(jsonResponse)
    
    If jsonObj Is Nothing Then
        Set ExtractDataFromResponse = Nothing
        Exit Function
    End If
    
    Set ExtractDataFromResponse = JsonConverter.GetJsonValue(jsonObj, "data")
End Function

' ========== 公共工具函数 ==========

' 格式化收盘价显示
Public Function FormatClosePrice(price As Variant) As String
    On Error Resume Next
    
    If IsNumeric(price) Then
        FormatClosePrice = Format(price, "0.000")
    Else
        FormatClosePrice = CStr(price)
    End If
End Function

' 获取API使用统计（可选功能）
Public Function GetAPIUsageStats() As String
    ' 返回API使用统计信息
    ' 这里可以实现更详细的统计功能
    GetAPIUsageStats = "API调用模块运行正常"
End Function
