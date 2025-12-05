VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ========== 自动刷新模块 (ThisWorkbook) ==========
' 负责工作簿事件处理和自动刷新功能

' ========== 工作簿打开事件 ==========
Private Sub Workbook_Open()
    ' 工作簿打开时自动执行的事件
    ' 根据配置决定是否自动刷新ETF价格数据
    
    On Error GoTo ErrorHandler
    
    ' 显示欢迎消息
    Application.StatusBar = "ETF价格追踪器已启动..."
    
    ' 延迟一下，让Excel完全加载
    Application.Wait DateAdd("s", 1, Now)
    
    ' 初始化工作表（如果需要）
    Call InitializeWorksheetIfNeeded
    
    ' 检查是否启用自动刷新
    If ShouldAutoRefresh() Then
        Call PerformAutoRefresh
    Else
        Call ShowWelcomeMessage
    End If
    
    ' 创建刷新按钮（如果不存在）
    Call CreateRefreshButtonIfNeeded
    
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    MsgBox "工作簿初始化时发生错误：" & Err.Description, vbExclamation, "初始化错误"
End Sub

' ========== 工作簿关闭前事件 ==========
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' 工作簿关闭前的清理工作
    
    On Error Resume Next
    
    ' 重置API频率限制计时器
    Call ResetApiRateLimit
    
    ' 清除状态栏
    Application.StatusBar = False
    
    ' 显示退出消息
    Debug.Print "ETF价格追踪器已关闭 - " & Now()
End Sub

' ========== 工作表激活事件 ==========
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    ' 当激活ETF价格工作表时的处理
    
    On Error Resume Next
    
    If Sh.Name = WORKSHEET_NAME Then
        ' 检查数据新鲜度
        Call CheckDataFreshness
    End If
End Sub

' ========== 检查是否应该自动刷新 ==========
Private Function ShouldAutoRefresh() As Boolean
    ' 确定是否应该自动刷新数据
    ' 基于配置、时间、网络连接等因素
    
    ShouldAutoRefresh = False
    
    On Error GoTo ErrorHandler
    
    ' 检查基本配置
    If Not ValidateConfiguration() Then
        Exit Function
    End If
    
    ' 检查是否在工作时间（避免在非交易时间频繁调用API）
    If Not IsMarketHours() Then
        Debug.Print "当前非交易时间，跳过自动刷新"
        Exit Function
    End If
    
    ' 检查上次刷新时间（避免频繁刷新）
    If Not ShouldRefreshBasedOnTime() Then
        Debug.Print "距离上次刷新时间太短，跳过自动刷新"
        Exit Function
    End If
    
    ShouldAutoRefresh = True
    Exit Function
    
ErrorHandler:
    Debug.Print "检查自动刷新条件时发生错误: " & Err.Description
    ShouldAutoRefresh = False
End Function

' ========== 执行自动刷新 ==========
Private Sub PerformAutoRefresh()
    ' 执行自动刷新过程
    
    On Error GoTo ErrorHandler
    
    ' 显示提示信息
    Dim result As VbMsgBoxResult
    result = MsgBox("检测到ETF价格数据，是否立即刷新最新价格？" & vbCrLf & vbCrLf & _
                   "提示：刷新过程可能需要几分钟时间", _
                   vbYesNo + vbQuestion + vbDefaultButton1, "自动刷新")
    
    If result = vbYes Then
        Application.StatusBar = "正在自动刷新ETF价格数据..."
        
        ' 执行刷新
        Call RefreshAllPrices
        
        ' 记录刷新时间
        Call RecordRefreshTime
    Else
        Call ShowManualRefreshTip
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    MsgBox "自动刷新时发生错误：" & Err.Description, vbExclamation, "自动刷新错误"
End Sub

' ========== 显示欢迎消息 ==========
Private Sub ShowWelcomeMessage()
    ' 显示欢迎和使用说明
    
    Dim welcomeMsg As String
    welcomeMsg = "欢迎使用ETF价格追踪器！" & vbCrLf & vbCrLf & _
                "使用说明：" & vbCrLf & _
                "1. 在A列输入ETF代码（6位数字）" & vbCrLf & _
                "2. 点击'刷新价格'按钮获取最新价格" & vbCrLf & _
                "3. 系统会自动填充B列（收盘价）和C列（数据日期）" & vbCrLf & vbCrLf & _
                "API状态：" & CheckApiStatus()
    
    MsgBox welcomeMsg, vbInformation, "ETF价格追踪器"
End Sub

' ========== 显示手动刷新提示 ==========
Private Sub ShowManualRefreshTip()
    ' 显示手动刷新的提示信息
    
    Dim tipMsg As String
    tipMsg = "您可以随时通过以下方式刷新价格：" & vbCrLf & vbCrLf & _
            "• 点击工作表中的'刷新价格'按钮" & vbCrLf & _
            "• 使用快捷键 Ctrl+Shift+R" & vbCrLf & _
            "• 在VBA编辑器中运行 RefreshAllPrices 宏"
    
    MsgBox tipMsg, vbInformation, "使用提示"
End Sub

' ========== 初始化工作表（如果需要）==========
Private Sub InitializeWorksheetIfNeeded()
    ' 如果ETF价格工作表不存在，则创建并初始化
    
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim wsExists As Boolean
    
    wsExists = False
    For Each ws In Me.Worksheets
        If ws.Name = WORKSHEET_NAME Then
            wsExists = True
            Exit For
        End If
    Next ws
    
    If Not wsExists Then
        Call InitializeWorksheet
    End If
End Sub

' ========== 创建刷新按钮（如果需要）==========
Private Sub CreateRefreshButtonIfNeeded()
    ' 如果刷新按钮不存在，则创建
    
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = Me.Worksheets(WORKSHEET_NAME)
    
    If ws Is Nothing Then Exit Sub
    
    ' 检查按钮是否存在
    Dim btnExists As Boolean
    Dim shp As Shape
    
    btnExists = False
    For Each shp In ws.Shapes
        If shp.Name = "RefreshButton" Then
            btnExists = True
            Exit For
        End If
    Next shp
    
    If Not btnExists Then
        Call CreateRefreshButton
    End If
End Sub

' ========== 检查是否在交易时间 ==========
Private Function IsMarketHours() As Boolean
    ' 检查当前是否在股市交易时间
    ' 简化版本：周一到周五，9:00-15:30
    
    Dim currentTime As Date
    Dim dayOfWeek As Integer

    currentTime = Now()
    dayOfWeek = Weekday(currentTime)

    ' 检查是否为工作日（周一到周五）
    If dayOfWeek < 2 Or dayOfWeek > 6 Then
        IsMarketHours = False
        Exit Function
    End If
    
    ' 检查时间范围（9:00-15:30）
    Dim currentHour As Integer
    Dim currentMinute As Integer
    
    currentHour = Hour(currentTime)
    currentMinute = Minute(currentTime)
    
    If currentHour < 9 Or currentHour > 15 Then
        IsMarketHours = False
    ElseIf currentHour = 15 And currentMinute > 30 Then
        IsMarketHours = False
    Else
        IsMarketHours = True
    End If
End Function

' ========== 基于时间检查是否应该刷新 ==========
Private Function ShouldRefreshBasedOnTime() As Boolean
    ' 基于上次刷新时间决定是否应该刷新
    ' 避免过于频繁的自动刷新
    
    ShouldRefreshBasedOnTime = True
    
    On Error Resume Next
    
    ' 获取上次刷新时间（存储在隐藏单元格中）
    Dim ws As Worksheet
    Set ws = Me.Worksheets(WORKSHEET_NAME)
    
    If ws Is Nothing Then Exit Function
    
    Dim lastRefreshTime As Date
    Dim lastRefreshStr As String
    
    ' 使用隐藏单元格存储上次刷新时间
    lastRefreshStr = ws.Range("Z1").Value
    
    If Len(lastRefreshStr) > 0 And IsDate(lastRefreshStr) Then
        lastRefreshTime = CDate(lastRefreshStr)
        
        ' 如果距离上次刷新不到30分钟，跳过自动刷新
        If DateDiff("n", lastRefreshTime, Now()) < 30 Then
            ShouldRefreshBasedOnTime = False
        End If
    End If
End Function

' ========== 记录刷新时间 ==========
Private Sub RecordRefreshTime()
    ' 记录当前刷新时间
    
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = Me.Worksheets(WORKSHEET_NAME)
    
    If Not ws Is Nothing Then
        ws.Range("Z1").Value = Format(Now(), "yyyy-mm-dd hh:mm:ss")
        ws.Range("Z1").EntireColumn.Hidden = True
    End If
End Sub

' ========== 检查数据新鲜度 ==========
Private Sub CheckDataFreshness()
    ' 检查数据是否为最新（在激活工作表时）
    
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = Me.Worksheets(WORKSHEET_NAME)
    
    If ws Is Nothing Then Exit Sub
    
    ' 获取最后一次刷新时间
    Dim lastRefreshStr As String
    lastRefreshStr = ws.Range("Z1").Value
    
    If Len(lastRefreshStr) > 0 And IsDate(lastRefreshStr) Then
        Dim lastRefresh As Date
        lastRefresh = CDate(lastRefreshStr)
        
        ' 如果数据超过1天未更新，显示提示
        If DateDiff("d", lastRefresh, Now()) > 1 Then
            Application.StatusBar = "数据可能已过期，建议刷新价格数据"
            
            ' 5秒后清除状态栏消息
            Application.OnTime Now + TimeValue("00:00:05"), "ClearStatusBar"
        End If
    End If
End Sub

' ========== 清除状态栏消息 ==========
Public Sub ClearStatusBar()
    ' 清除状态栏消息的公共函数
    Application.StatusBar = False
End Sub

' ========== 快捷键绑定 ==========
Private Sub Workbook_WindowActivate(ByVal Wn As Window)
    ' 设置快捷键绑定
    
    On Error Resume Next
    
    ' 绑定Ctrl+Shift+R为刷新快捷键
    Application.OnKey "^+R", "RefreshAllPrices"
    
    ' 绑定Ctrl+Shift+T为API测试快捷键
    Application.OnKey "^+T", "TestApiConnection"
    
    ' 绑定Ctrl+Shift+A为添加ETF代码快捷键
    Application.OnKey "^+A", "AddNewETFCode"
End Sub

' ========== 取消快捷键绑定 ==========
Private Sub Workbook_WindowDeactivate(ByVal Wn As Window)
    ' 取消快捷键绑定
    
    On Error Resume Next
    
    Application.OnKey "^+R"
    Application.OnKey "^+T"
    Application.OnKey "^+A"
End Sub
