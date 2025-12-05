Attribute VB_Name = "Module_Config"
Option Explicit

' ========== API 配置常量 ==========
Public Const API_BASE_URL As String = "https://open.lixinger.com/api/cn/fund/candlestick"
Public Const API_TOKEN As String = "30cf80b9-ac68-4521-86c2-5847d7ce728e"
Public Const API_TIMEOUT As Integer = 30
Public Const API_RATE_LIMIT_DELAY As Single = 0.5 ' 每次请求间隔0.5秒

' ========== 工作表配置 ==========
Public Const WORKSHEET_NAME As String = "ETF价格"
Public Const ETF_CODE_COLUMN As Integer = 1  ' A列：ETF代码
Public Const PRICE_COLUMN As Integer = 2     ' B列：最新收盘价
Public Const DATE_COLUMN As Integer = 3      ' C列：数据日期
Public Const HEADER_ROW As Integer = 1       ' 第1行为表头

' ========== 表头文本 ==========
Public Const HEADER_ETF_CODE As String = "ETF代码"
Public Const HEADER_PRICE As String = "最新收盘价"
Public Const HEADER_DATE As String = "数据日期"

' ========== 错误信息 ==========
Public Const ERROR_NETWORK As String = "网络错误"
Public Const ERROR_INVALID_CODE As String = "无效代码"
Public Const ERROR_API As String = "API错误"
Public Const ERROR_NO_DATA As String = "无数据"

' ========== 全局变量 ==========
Public LastRequestTime As Double ' 用于控制API调用频率

' ========== 配置验证函数 ==========
Public Function ValidateConfiguration() As Boolean
    ' 验证Token是否配置
    If Len(API_TOKEN) = 0 Then
        MsgBox "错误：未配置API Token，请联系管理员", vbCritical, "配置错误"
        ValidateConfiguration = False
        Exit Function
    End If
    
    ' 验证网络连接
    If Not CheckInternetConnection() Then
        MsgBox "错误：无法连接到互联网，请检查网络连接", vbCritical, "网络错误"
        ValidateConfiguration = False
        Exit Function
    End If
    
    ValidateConfiguration = True
End Function

' ========== 网络连接检查 ==========
Public Function CheckInternetConnection() As Boolean
    On Error GoTo ErrorHandler
    
    ' 使用简单的HTTP请求测试网络连接
    Dim httpRequest As Object
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    
    httpRequest.Open "GET", "https://www.baidu.com", False
    httpRequest.setRequestHeader "User-Agent", "Excel VBA ETF Tracker"
    httpRequest.send
    
    CheckInternetConnection = (httpRequest.Status = 200)
    Exit Function
    
ErrorHandler:
    CheckInternetConnection = False
End Function

' ========== API频率限制控制 ==========
Public Sub WaitForApiRateLimit()
    Dim currentTime As Double
    Dim timeSinceLastRequest As Double
    
    currentTime = Timer
    timeSinceLastRequest = currentTime - LastRequestTime
    
    ' 如果距离上次请求时间不足延迟时间，则等待
    If timeSinceLastRequest < API_RATE_LIMIT_DELAY Then
        Application.Wait DateAdd("s", API_RATE_LIMIT_DELAY - timeSinceLastRequest, Now)
    End If
    
    LastRequestTime = Timer
End Sub

' ========== 工作表初始化 ==========
Public Sub InitializeWorksheet()
    Dim ws As Worksheet
    Dim wsExists As Boolean
    
    ' 检查工作表是否存在
    wsExists = False
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name = WORKSHEET_NAME Then
            wsExists = True
            Set ws = ActiveWorkbook.Worksheets(WORKSHEET_NAME)
            Exit For
        End If
    Next ws
    
    ' 如果工作表不存在，创建它
    If Not wsExists Then
        Set ws = ActiveWorkbook.Worksheets.Add
        ws.Name = WORKSHEET_NAME
    End If
    
    ' 设置表头
    With ws
        .Cells(HEADER_ROW, ETF_CODE_COLUMN).Value = HEADER_ETF_CODE
        .Cells(HEADER_ROW, PRICE_COLUMN).Value = HEADER_PRICE
        .Cells(HEADER_ROW, DATE_COLUMN).Value = HEADER_DATE
        
        ' 格式化表头
        .Range(.Cells(HEADER_ROW, ETF_CODE_COLUMN), .Cells(HEADER_ROW, DATE_COLUMN)).Font.Bold = True
        .Range(.Cells(HEADER_ROW, ETF_CODE_COLUMN), .Cells(HEADER_ROW, DATE_COLUMN)).Interior.Color = RGB(200, 200, 200)
        
        ' 设置列宽
        .Columns(ETF_CODE_COLUMN).ColumnWidth = 12
        .Columns(PRICE_COLUMN).ColumnWidth = 15
        .Columns(DATE_COLUMN).ColumnWidth = 15
        
        ' 冻结首行
        .Range("A2").Select
        ActiveWindow.FreezePanes = True
    End With
    
    ' 激活工作表
    ws.Activate
End Sub

' ========== 获取当前日期字符串 ==========
Public Function GetCurrentDateString() As String
    GetCurrentDateString = Format(Date, "yyyy-mm-dd")
End Function

' ========== 获取5天前日期字符串（用于API默认查询范围）==========
Public Function GetFiveDaysAgoString() As String
    GetFiveDaysAgoString = Format(DateAdd("d", -5, Date), "yyyy-mm-dd")
End Function
