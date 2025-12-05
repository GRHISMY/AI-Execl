Attribute VB_Name = "Module_Config"
' =========================================
' 模块名称: Module_Config
' 描述: ETF价格追踪器配置模块
' 作者: Aone Copilot
' 创建日期: 2025-12-05
' =========================================

Option Explicit

' ========== 全局常量 ==========

' API配置
Public Const API_BASE_URL As String = "https://open.lixinger.com/api/cn/fund/candlestick"
Public Const API_TIMEOUT As Integer = 30  ' 请求超时时间（秒）

' 频率限制配置
Public Const MAX_REQUESTS_PER_MINUTE As Integer = 120  ' 每分钟最大请求数
Public Const REQUEST_DELAY_MS As Integer = 500  ' 请求之间的延迟（毫秒）

' 工作表配置
Public Const SHEET_NAME As String = "ETF价格"
Public Const CONFIG_SHEET_NAME As String = "配置"
Public Const HEADER_ROW As Integer = 1
Public Const DATA_START_ROW As Integer = 2

' 列索引
Public Const COL_ETF_CODE As Integer = 1    ' A列：ETF代码
Public Const COL_CLOSE_PRICE As Integer = 2  ' B列：收盘价
Public Const COL_DATA_DATE As Integer = 3    ' C列：数据日期
Public Const COL_UPDATE_TIME As Integer = 4  ' D列：更新时间

' ========== 全局变量 ==========

' API Token（从配置工作表读取）
Private m_ApiToken As String
Private m_LastRequestTime As Double

' ========== 公共函数 ==========

' 获取API Token
Public Function GetApiToken() As String
    If m_ApiToken = "" Then
        m_ApiToken = ReadConfigValue("ApiToken")
    End If
    
    If m_ApiToken = "" Then
        MsgBox "请先配置理想财经API Token！" & vbCrLf & _
               "请在'" & CONFIG_SHEET_NAME & "'工作表中设置 ApiToken", _
               vbExclamation, "配置错误"
        GetApiToken = ""
        Exit Function
    End If
    
    GetApiToken = m_ApiToken
End Function

' 设置API Token
Public Sub SetApiToken(token As String)
    m_ApiToken = token
    WriteConfigValue "ApiToken", token
End Sub

' 从配置工作表读取配置值
Private Function ReadConfigValue(configKey As String) As String
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' 尝试获取配置工作表
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    
    If ws Is Nothing Then
        ' 如果配置工作表不存在，创建它
        Set ws = CreateConfigSheet()
    End If
    
    ' 查找配置项
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = configKey Then
            ReadConfigValue = ws.Cells(i, 2).Value
            Exit Function
        End If
    Next i
    
    ReadConfigValue = ""
End Function

' 向配置工作表写入配置值
Private Sub WriteConfigValue(configKey As String, configValue As String)
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim found As Boolean
    
    ' 获取或创建配置工作表
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    If ws Is Nothing Then
        Set ws = CreateConfigSheet()
    End If
    
    ' 查找是否已存在该配置项
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    found = False
    
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = configKey Then
            ws.Cells(i, 2).Value = configValue
            found = True
            Exit For
        End If
    Next i
    
    ' 如果不存在，添加新行
    If Not found Then
        lastRow = lastRow + 1
        ws.Cells(lastRow, 1).Value = configKey
        ws.Cells(lastRow, 2).Value = configValue
    End If
End Sub

' 创建配置工作表
Private Function CreateConfigSheet() As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = CONFIG_SHEET_NAME
        
        ' 设置表头
        ws.Cells(1, 1).Value = "配置项"
        ws.Cells(1, 2).Value = "配置值"
        
        ' 格式化表头
        With ws.Range("A1:B1")
            .Font.Bold = True
            .Interior.Color = RGB(200, 200, 200)
        End With
        
        ' 添加默认配置项
        ws.Cells(2, 1).Value = "ApiToken"
        ws.Cells(2, 2).Value = "请在此处输入您的理想财经API Token"
        
        ' 设置列宽
        ws.Columns("A:A").ColumnWidth = 20
        ws.Columns("B:B").ColumnWidth = 50
        
        ' 隐藏工作表（可选）
        ' ws.Visible = xlSheetVeryHidden
    End If
    
    Set CreateConfigSheet = ws
End Function

' API频率限制控制
Public Sub WaitForRateLimit()
    Dim currentTime As Double
    Dim timeSinceLastRequest As Double
    Dim minDelay As Double
    
    ' 最小延迟时间（秒）
    minDelay = REQUEST_DELAY_MS / 1000
    
    currentTime = Timer
    
    If m_LastRequestTime > 0 Then
        timeSinceLastRequest = currentTime - m_LastRequestTime
        
        ' 如果距离上次请求时间太短，等待
        If timeSinceLastRequest < minDelay Then
            Application.Wait Now + TimeValue("0:00:0" & Format(minDelay - timeSinceLastRequest, "0.000"))
        End If
    End If
    
    m_LastRequestTime = Timer
End Sub

' 检查网络连接
Public Function CheckNetworkConnection() As Boolean
    On Error Resume Next
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' 尝试连接到百度
    http.Open "GET", "https://www.baidu.com", False
    http.setTimeOut 3000
    http.send
    
    If Err.Number <> 0 Or http.Status <> 200 Then
        CheckNetworkConnection = False
    Else
        CheckNetworkConnection = True
    End If
    
    Set http = Nothing
End Function

' 验证配置完整性
Public Function ValidateConfig() As Boolean
    Dim token As String
    
    ' 检查Token是否配置
    token = GetApiToken()
    If token = "" Or token = "请在此处输入您的理想财经API Token" Then
        MsgBox "请先配置理想财经API Token！" & vbCrLf & vbCrLf & _
               "步骤：" & vbCrLf & _
               "1. 访问理想财经官网获取API Token" & vbCrLf & _
               "2. 在'" & CONFIG_SHEET_NAME & "'工作表中填入Token" & vbCrLf & _
               "3. 重新打开此工作簿", _
               vbExclamation, "配置错误"
        ValidateConfig = False
        Exit Function
    End If
    
    ' 检查网络连接
    If Not CheckNetworkConnection() Then
        MsgBox "网络连接失败！请检查网络设置。", vbExclamation, "网络错误"
        ValidateConfig = False
        Exit Function
    End If
    
    ValidateConfig = True
End Function
