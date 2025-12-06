Sub CreateMacConfigTestButton()
    ' 创建Mac兼容配置测试按钮
    
    On Error Resume Next
    
    ' 删除现有按钮
    ActiveSheet.Buttons("MacConfigTest").Delete
    ActiveSheet.Buttons("MacTokenConfig").Delete
    
    On Error GoTo ErrorHandler
    
    ' 创建测试按钮
    Dim btn1 As Button
    Set btn1 = ActiveSheet.Buttons.Add(10, 10, 120, 25)
    btn1.Name = "MacConfigTest"
    btn1.Text = "测试Mac配置"
    btn1.OnAction = "Module_Config_Mac.TestConfigSystem"
    
    ' 创建Token配置按钮
    Dim btn2 As Button
    Set btn2 = ActiveSheet.Buttons.Add(10, 40, 120, 25)
    btn2.Name = "MacTokenConfig"
    btn2.Text = "设置API Token"
    btn2.OnAction = "Module_Config_Mac.ShowConfigDialog"
    
    MsgBox "Mac兼容配置按钮已创建！" & vbCrLf & vbCrLf & "按钮说明:" & vbCrLf & "• 测试Mac配置: 测试配置系统是否正常" & vbCrLf & "• 设置API Token: 配置lixinger API Token", vbInformation, "按钮创建成功"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "创建按钮时出错: " & Err.Description, vbExclamation, "错误"
End Sub

Sub QuickTestMacConfig()
    ' 快速测试Mac配置功能
    Module_Config_Mac.ShowConfigDialog
End Sub
