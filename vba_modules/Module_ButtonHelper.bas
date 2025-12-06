' Attribute VB_Name = "Module_ButtonHelper"
'
' Module_ButtonHelper - 按钮创建辅助模块
' 专为Mac Excel设计的按钮管理
'
Option Explicit

' 创建Mac兼容配置测试按钮
Sub CreateMacConfigTestButton()
    On Error GoTo ErrorHandler
    
    ' 删除现有按钮（如果存在）
    DeleteButtonSafely "MacConfigTest"
    DeleteButtonSafely "MacTokenConfig"
    DeleteButtonSafely "TestConfig"
    
    ' 创建测试配置系统按钮
    Dim btn1 As Button
    Set btn1 = ActiveSheet.Buttons.Add(10, 10, 140, 30)
    btn1.Name = "MacConfigTest"
    btn1.Text = "测试Mac配置系统"
    btn1.OnAction = "Module_Config_Mac.TestConfigSystem"
    
    ' 创建API Token配置按钮
    Dim btn2 As Button
    Set btn2 = ActiveSheet.Buttons.Add(10, 50, 140, 30)
    btn2.Name = "MacTokenConfig"
    btn2.Text = "设置API Token"
    btn2.OnAction = "Module_Config_Mac.ShowConfigDialog"
    
    ' 创建配置验证按钮
    Dim btn3 As Button
    Set btn3 = ActiveSheet.Buttons.Add(10, 90, 140, 30)
    btn3.Name = "TestConfig"
    btn3.Text = "验证配置"
    btn3.OnAction = "Module_Config_Mac.ValidateConfig"
    
    MsgBox "Mac兼容配置按钮创建成功！" & vbCrLf & vbCrLf & _
           "已创建3个按钮：" & vbCrLf & _
           "• 测试Mac配置系统" & vbCrLf & _
           "• 设置API Token" & vbCrLf & _
           "• 验证配置", vbInformation, "按钮创建完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "创建按钮时出错: " & Err.Description & vbCrLf & _
           "错误号: " & Err.Number, vbExclamation, "按钮创建错误"
End Sub

' 安全删除按钮
Private Sub DeleteButtonSafely(buttonName As String)
    On Error Resume Next
    ActiveSheet.Buttons(buttonName).Delete
    On Error GoTo 0
End Sub

' 直接测试Mac配置功能
Sub QuickTestMacConfig()
    Module_Config_Mac.TestConfigSystem
End Sub

' 直接配置API Token
Sub QuickConfigToken()
    Module_Config_Mac.ShowConfigDialog
End Sub

' 清理所有测试按钮
Sub CleanupTestButtons()
    On Error Resume Next
    DeleteButtonSafely "MacConfigTest"
    DeleteButtonSafely "MacTokenConfig"
    DeleteButtonSafely "TestConfig"
    On Error GoTo 0
    MsgBox "测试按钮已清理", vbInformation, "清理完成"
End Sub

' 检查模块状态
Sub CheckModuleStatus()
    Dim msg As String
    msg = "模块状态检查:" & vbCrLf & vbCrLf
    
    ' 检查Mac配置模块
    msg = msg & "Mac兼容配置模块: 已加载" & vbCrLf
    
    ' 检查系统
    msg = msg & "操作系统: " & IIf(IsMacSystem(), "Mac", "Windows") & vbCrLf
    
    ' 检查配置文件
    Dim configExists As Boolean
    configExists = Module_Config_Mac.ValidateConfig()
    msg = msg & "配置状态: " & IIf(configExists, "已配置", "未配置") & vbCrLf
    
    MsgBox msg, vbInformation, "模块状态"
End Sub

' 检测Mac系统
Private Function IsMacSystem() As Boolean
    #If Mac Then
        IsMacSystem = True
    #Else
        IsMacSystem = False
    #End If
End Function
