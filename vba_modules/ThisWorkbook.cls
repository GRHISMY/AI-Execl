VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
' ThisWorkbook - 工作簿事件处理
' 处理工作簿打开、关闭等事件，初始化环境
'
Option Explicit

' 工作簿打开事件
Private Sub Workbook_Open()
    On Error GoTo ErrorHandler
    
    ' 初始化配置
    Call InitializeWorkbook
    
    ' 检查环境
    Call CheckEnvironment
    
    ' 显示欢迎信息（仅第一次）
    If Not HasBeenInitialized() Then
        Call ShowWelcomeDialog
        Call MarkAsInitialized
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "工作簿打开事件错误: " & Err.Description
End Sub

' 工作簿关闭前事件
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    
    ' 清理状态栏
    Application.StatusBar = False
    
    ' 可以在这里添加其他清理工作
End Sub

' 工作表激活事件
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    
    ' 确保工作表有正确的表头
    Call EnsureWorksheetHeaders(Sh)
End Sub

' 初始化工作簿
Private Sub InitializeWorkbook()
    On Error GoTo ErrorHandler
    
    ' 设置应用程序属性
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ' Mac兼容配置系统，不需要显式加载配置
    
    ' 初始化当前工作表
    Call EnsureWorksheetHeaders(ActiveSheet)
    
    ' 创建按钮（如果不存在）
    Call CreateControlButtons
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "初始化工作簿错误: " & Err.Description
End Sub

' 检查运行环境
Private Sub CheckEnvironment()
    On Error GoTo ErrorHandler
    
    Dim hasExecutable As Boolean
    Dim hasConfig As Boolean
    
    ' 检查可执行文件
    hasExecutable = (Len(Module_API.GetExecutablePath()) > 0)
    
    ' 检查配置
    hasConfig = Module_Config.ValidateConfig()
    
    ' 更新状态栏
    If hasExecutable And hasConfig Then
        Application.StatusBar = "ETF工具就绪"
    ElseIf hasExecutable Then
        Application.StatusBar = "需要配置API Token"
    Else
        Application.StatusBar = "缺少API可执行文件"
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "环境检查错误: " & Err.Description
End Sub

' 确保工作表有正确的表头
Private Sub EnsureWorksheetHeaders(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    ' 检查是否已有表头
    If Len(Trim(ws.Cells(1, 1).Value)) > 0 Then
        Exit Sub
    End If
    
    ' 设置表头
    ws.Cells(1, 1).Value = "ETF代码"
    ws.Cells(1, 2).Value = "收盘价"
    ws.Cells(1, 3).Value = "状态"
    ws.Cells(1, 4).Value = "更新时间"
    
    ' 格式化表头
    With ws.Range("A1:D1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' 设置列宽
    ws.Columns("A:A").ColumnWidth = 12  ' ETF代码
    ws.Columns("B:B").ColumnWidth = 15  ' 收盘价
    ws.Columns("C:C").ColumnWidth = 12  ' 状态
    ws.Columns("D:D").ColumnWidth = 20  ' 更新时间
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "设置表头错误: " & Err.Description
End Sub

' 创建控制按钮
Private Sub CreateControlButtons()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 删除现有按钮（如果存在）
    Call DeleteExistingButtons(ws)
    
    ' 创建刷新按钮
    Dim refreshBtn As Button
    Set refreshBtn = ws.Buttons.Add(ws.Cells(1, 6).Left, ws.Cells(1, 6).Top, 80, 25)
    refreshBtn.Name = "RefreshButton"
    refreshBtn.Caption = "刷新数据"
    refreshBtn.OnAction = "Module_Refresh.RefreshETFPrices"
    
    ' 创建配置按钮
    Dim configBtn As Button
    Set configBtn = ws.Buttons.Add(ws.Cells(1, 7).Left, ws.Cells(1, 7).Top, 80, 25)
    configBtn.Name = "ConfigButton"
    configBtn.Caption = "设置"
    configBtn.OnAction = "Module_Config.ShowConfigDialog"
    
    ' 创建清除按钮
    Dim clearBtn As Button
    Set clearBtn = ws.Buttons.Add(ws.Cells(1, 8).Left, ws.Cells(1, 8).Top, 80, 25)
    clearBtn.Name = "ClearButton"
    clearBtn.Caption = "清除数据"
    clearBtn.OnAction = "Module_Refresh.ClearAllData"
    
    ' 创建示例按钮
    Dim sampleBtn As Button
    Set sampleBtn = ws.Buttons.Add(ws.Cells(1, 9).Left, ws.Cells(1, 9).Top, 80, 25)
    sampleBtn.Name = "SampleButton"
    sampleBtn.Caption = "示例代码"
    sampleBtn.OnAction = "Module_Refresh.AddSampleETFCodes"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "创建按钮错误: " & Err.Description
End Sub

' 删除现有按钮
Private Sub DeleteExistingButtons(ws As Worksheet)
    On Error Resume Next
    
    Dim btn As Button
    
    ' 删除已知的按钮
    ws.Buttons("RefreshButton").Delete
    ws.Buttons("ConfigButton").Delete
    ws.Buttons("ClearButton").Delete
    ws.Buttons("SampleButton").Delete
End Sub

' 显示欢迎对话框
Private Sub ShowWelcomeDialog()
    On Error GoTo ErrorHandler
    
    Dim msg As String
    msg = "欢迎使用ETF价格获取工具！" & vbCrLf & vbCrLf
    msg = msg & "使用方法：" & vbCrLf
    msg = msg & "1. 在A列输入ETF代码（如：159915）" & vbCrLf
    msg = msg & "2. 点击'设置'按钮配置API Token" & vbCrLf
    msg = msg & "3. 点击'刷新数据'获取价格信息" & vbCrLf & vbCrLf
    msg = msg & "注意：需要有效的lixinger API Token才能使用"
    
    MsgBox msg, vbInformation, "ETF价格工具 - 使用指南"
    
    ' 检查是否需要配置
    If Not Module_Config.ValidateConfig() Then
        If MsgBox("是否现在配置API Token？", vbYesNo + vbQuestion, "配置确认") = vbYes Then
            Module_Config.ShowConfigDialog
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "显示欢迎对话框错误: " & Err.Description
End Sub

' 检查是否已初始化
Private Function HasBeenInitialized() As Boolean
    On Error Resume Next
    
    HasBeenInitialized = (Module_Config.GetConfig("app.initialized") = "True")
End Function

' 标记为已初始化
Private Sub MarkAsInitialized()
    On Error Resume Next
    
    Module_Config.SetConfig "app.initialized", "True"
End Sub

' 右键菜单事件
Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    On Error GoTo ErrorHandler
    
    ' 如果点击的是A列（ETF代码列）
    If Target.Column = 1 And Target.Row > 1 Then
        Dim etfCode As String
        etfCode = Trim(Target.Value)
        
        If Len(etfCode) > 0 And Module_API.ValidateETFCodes(etfCode) Then
            ' 显示自定义菜单
            Call ShowCustomContextMenu(etfCode)
            Cancel = True  ' 取消默认右键菜单
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "右键菜单事件错误: " & Err.Description
End Sub

' 显示自定义上下文菜单
Private Sub ShowCustomContextMenu(etfCode As String)
    On Error GoTo ErrorHandler
    
    ' 删除现有的自定义菜单
    On Error Resume Next
    Application.CommandBars("ETF_ContextMenu").Delete
    On Error GoTo ErrorHandler
    
    ' 创建新的上下文菜单
    Dim contextMenu As CommandBar
    Set contextMenu = Application.CommandBars.Add(Name:="ETF_ContextMenu", Position:=msoBarPopup, Temporary:=True)
    
    ' 添加菜单项
    Dim menuItem As CommandBarButton
    
    ' 刷新单个ETF
    Set menuItem = contextMenu.Controls.Add(Type:=msoControlButton)
    menuItem.Caption = "刷新 " & etfCode
    menuItem.OnAction = "RefreshSingleETFFromMenu"
    menuItem.Parameter = etfCode
    
    ' 显示菜单
    contextMenu.ShowPopup
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "显示自定义菜单错误: " & Err.Description
End Sub

' 从菜单刷新单个ETF
Public Sub RefreshSingleETFFromMenu()
    On Error GoTo ErrorHandler
    
    Dim etfCode As String
    etfCode = Application.CommandBars.ActionControl.Parameter
    
    Module_Refresh.RefreshSingleETF etfCode
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "从菜单刷新单个ETF错误: " & Err.Description
End Sub
