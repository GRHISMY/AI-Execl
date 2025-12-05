Attribute VB_Name = "Module_API"
Option Explicit

' ========== API调用模块 ==========
' 负责调用理想财经API获取ETF数据

' ========== 主要功能函数 ==========
Public Function GetLatestClosePrice(etfCode As String) As Variant
    ' 获取ETF最新收盘价
    ' 参数: etfCode - ETF代码（如：512690）
    ' 返回: 收盘价（数值）或错误信息（字符串）
    
    On Error GoTo ErrorHandler
    
    ' 验证配置
    If Not ValidateConfiguration() Then
        GetLatestClosePrice = ERROR_API
        Exit Function
    End If
    
    ' 验证ETF代码
    If Not IsValidETFCode(etfCode) Then
        GetLatestClosePrice = ERROR_INVALID_CODE
        Exit Function
    End If
    
    ' 应用API频率限制
    WaitForApiRateLimit
    
    ' 构建请求参数
    Dim startDate As String
    Dim endDate As String
    startDate = GetFiveDaysAgoString()
    endDate = GetCurrentDateString()
    
    ' 调用API
    Dim jsonResponse As String
    jsonResponse = CallLixingerApi(etfCode, startDate, endDate)
    
    If Len(jsonResponse) = 0 Then
        GetLatestClosePrice = ERROR_NETWORK
        Exit Function
    End If
    
    ' 解析响应获取收盘价
    Dim closePrice As Variant
    closePrice = GetLatestClosePriceFromJson(jsonResponse)
    
    ' 检查解析结果
    If IsNumeric(closePrice) Then
        GetLatestClosePrice = CDbl(closePrice)
    Else
        GetLatestClosePrice = closePrice ' 返回错误信息
    End If
    
    Exit Function
    
ErrorHandler:
    GetLatestClosePrice = ERROR_API & ": " & Err.Description
End Function

' ========== 获取ETF数据日期 ==========
Public Function GetLatestDataDate(etfCode As String) As Variant
    ' 获取ETF最新数据日期
    ' 参数: etfCode - ETF代码
    ' 返回: 日期字符串或错误信息
    
    On Error GoTo ErrorHandler
    
    ' 验证配置
    If Not ValidateConfiguration() Then
        GetLatestDataDate = ERROR_API
        Exit Function
    End If
    
    ' 验证ETF代码
    If Not IsValidETFCode(etfCode) Then
        GetLatestDataDate = ERROR_INVALID_CODE
        Exit Function
    End If
    
    ' 应用API频率限制
    WaitForApiRateLimit
    
    ' 构建请求参数
    Dim startDate As String
    Dim endDate As String
    startDate = GetFiveDaysAgoString()
    endDate = GetCurrentDateString()
    
    ' 调用API
    Dim jsonResponse As String
    jsonResponse = CallLixingerApi(etfCode, startDate, endDate)
    
    If Len(jsonResponse) = 0 Then
        GetLatestDataDate = ERROR_NETWORK
        Exit Function
    End If
    
    ' 解析响应获取日期
    GetLatestDataDate = GetLatestDateFromJson(jsonResponse)
    
    Exit Function
    
ErrorHandler:
    GetLatestDataDate = ERROR_API & ": " & Err.Description
End Function

' ========== 核心API调用函数 ==========
Private Function CallLixingerApi(etfCode As String, startDate As String, endDate As String) As String
    ' 调用理想财经API
    ' 返回: JSON响应字符串
    
    On Error GoTo ErrorHandler
    
    Dim httpRequest As Object
    Dim requestPayload As String
    Dim responseText As String
    
    ' 创建HTTP请求对象
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    
    ' 构建请求载荷
    requestPayload = BuildRequestPayload(etfCode, startDate, endDate)
    
    ' 配置请求
    With httpRequest
        .Open "POST", API_BASE_URL, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "User-Agent", "Excel VBA ETF Tracker v1.0"
        .setRequestHeader "Accept", "application/json"
        
        ' 设置超时（如果支持）
        On Error Resume Next
        .setTimeouts API_TIMEOUT * 1000, API_TIMEOUT * 1000, API_TIMEOUT * 1000, API_TIMEOUT * 1000
        On Error GoTo ErrorHandler
        
        ' 发送请求
        .send requestPayload
        
        ' 检查响应状态
        If .Status = 200 Then
            responseText = .responseText
        Else
            Debug.Print "API请求失败: HTTP " & .Status & " - " & .statusText
            responseText = ""
        End If
    End With
    
    CallLixingerApi = responseText
    Exit Function
    
ErrorHandler:
    Debug.Print "API调用错误: " & Err.Description
    CallLixingerApi = ""
End Function

' ========== 构建请求载荷 ==========
Private Function BuildRequestPayload(etfCode As String, startDate As String, endDate As String) As String
    ' 构建JSON请求载荷
    Dim payload As String
    
    payload = "{" & _
        """token"": """ & API_TOKEN & """," & _
        """stockCode"": """ & etfCode & """," & _
        """startDate"": """ & startDate & """," & _
        """endDate"": """ & endDate & """" & _
        "}"
    
    BuildRequestPayload = payload
End Function

' ========== ETF代码验证 ==========
Private Function IsValidETFCode(etfCode As String) As Boolean
    ' 验证ETF代码格式
    ' ETF代码通常是6位数字
    
    If Len(etfCode) <> 6 Then
        IsValidETFCode = False
        Exit Function
    End If
    
    If Not IsNumeric(etfCode) Then
        IsValidETFCode = False
        Exit Function
    End If
    
    IsValidETFCode = True
End Function

' ========== 测试API连接 ==========
Public Function TestApiConnection() As Boolean
    ' 测试API连接是否正常
    ' 返回: True表示连接正常，False表示连接失败
    
    On Error GoTo ErrorHandler
    
    ' 使用一个常见的ETF代码进行测试（如：沪深300ETF 510300）
    Dim testCode As String
    testCode = "510300"
    
    Dim result As Variant
    result = GetLatestClosePrice(testCode)
    
    ' 如果返回数值，说明API连接正常
    If IsNumeric(result) Then
        TestApiConnection = True
        MsgBox "API连接测试成功！" & vbCrLf & "测试代码: " & testCode & vbCrLf & "收盘价: " & result, vbInformation, "连接测试"
    Else
        TestApiConnection = False
        MsgBox "API连接测试失败！" & vbCrLf & "错误信息: " & result, vbExclamation, "连接测试"
    End If
    
    Exit Function
    
ErrorHandler:
    TestApiConnection = False
    MsgBox "API连接测试出错！" & vbCrLf & "错误信息: " & Err.Description, vbCritical, "连接测试"
End Function

' ========== 批量获取收盘价（用于优化API调用）==========
Public Function GetMultipleClosePrices(etfCodes As Variant) As Dictionary
    ' 批量获取多个ETF的收盘价
    ' 参数: etfCodes - ETF代码数组
    ' 返回: Dictionary，键为ETF代码，值为收盘价或错误信息
    
    Set GetMultipleClosePrices = New Dictionary
    
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim etfCode As String
    Dim result As Variant
    
    ' 显示进度
    Application.StatusBar = "正在获取ETF数据..."
    
    For i = LBound(etfCodes) To UBound(etfCodes)
        etfCode = CStr(etfCodes(i))
        
        If Len(etfCode) > 0 Then
            ' 显示当前进度
            Application.StatusBar = "正在获取ETF数据... (" & (i - LBound(etfCodes) + 1) & "/" & (UBound(etfCodes) - LBound(etfCodes) + 1) & ") " & etfCode
            
            ' 获取收盘价
            result = GetLatestClosePrice(etfCode)
            GetMultipleClosePrices(etfCode) = result
            
            ' 刷新屏幕显示
            DoEvents
        End If
    Next i
    
    Application.StatusBar = False
    Exit Function
    
ErrorHandler:
    Application.StatusBar = False
    Debug.Print "批量获取数据错误: " & Err.Description
End Function

' ========== API状态检查 ==========
Public Function CheckApiStatus() As String
    ' 检查API服务状态
    ' 返回: 状态描述字符串
    
    On Error GoTo ErrorHandler
    
    Dim httpRequest As Object
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    
    ' 尝试访问API基础URL（简单的连通性测试）
    With httpRequest
        .Open "GET", "https://open.lixinger.com", False
        .setRequestHeader "User-Agent", "Excel VBA ETF Tracker v1.0"
        
        On Error Resume Next
        .setTimeouts 10000, 10000, 10000, 10000
        On Error GoTo ErrorHandler
        
        .send
        
        If .Status = 200 Then
            CheckApiStatus = "API服务正常"
        Else
            CheckApiStatus = "API服务异常: HTTP " & .Status
        End If
    End With
    
    Exit Function
    
ErrorHandler:
    CheckApiStatus = "API服务检查失败: " & Err.Description
End Function

' ========== 获取API使用统计 ==========
Public Function GetApiUsageInfo() As String
    ' 获取API使用统计信息（模拟功能）
    ' 在实际应用中，可以记录API调用次数和频率
    
    Dim info As String
    info = "API配置信息:" & vbCrLf
    info = info & "- 服务地址: " & API_BASE_URL & vbCrLf
    info = info & "- Token: " & Left(API_TOKEN, 8) & "..." & vbCrLf
    info = info & "- 超时时间: " & API_TIMEOUT & "秒" & vbCrLf
    info = info & "- 频率限制: " & API_RATE_LIMIT_DELAY & "秒间隔" & vbCrLf
    info = info & "- 当前时间: " & Now()
    
    GetApiUsageInfo = info
End Function

' ========== 重置API频率限制计时器 ==========
Public Sub ResetApiRateLimit()
    ' 重置API频率限制计时器
    ' 在某些情况下可能需要手动重置
    LastRequestTime = 0
End Sub
