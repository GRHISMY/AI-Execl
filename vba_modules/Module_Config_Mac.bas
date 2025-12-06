' Attribute VB_Name = "Module_Config_Mac"
'
' Module_Config_Mac - Mac兼容的配置管理模块
' 完全不依赖ActiveX组件，专为Mac Excel设计
'
Option Explicit

' Mac兼容的字典实现
Private Type KeyValue
    Key As String
    Value As String
End Type

Private configDict() As KeyValue
Private configDictSize As Integer

' 检测操作系统
Private Function IsMac() As Boolean
    #If Mac Then
        IsMac = True
    #Else
        IsMac = False
    #End If
End Function

' 初始化配置字典
Private Sub InitConfigDict()
    configDictSize = 0
    ReDim configDict(0 To 99) ' 预分配空间
End Sub

' 向字典添加项目
Private Sub AddToConfigDict(key As String, value As String)
    Dim i As Integer
    
    ' 查找是否已存在，如果存在则更新
    For i = 0 To configDictSize - 1
        If configDict(i).Key = key Then
            configDict(i).value = value
            Exit Sub
        End If
    Next i
    
    ' 添加新项
    If configDictSize >= UBound(configDict) Then
        ReDim Preserve configDict(0 To configDictSize + 99)
    End If
    
    configDict(configDictSize).Key = key
    configDict(configDictSize).value = value
    configDictSize = configDictSize + 1
End Sub

' 从字典获取值
Private Function GetFromConfigDict(key As String) As String
    Dim i As Integer
    For i = 0 To configDictSize - 1
        If configDict(i).Key = key Then
            GetFromConfigDict = configDict(i).value
            Exit Function
        End If
    Next i
    GetFromConfigDict = ""
End Function

' Mac兼容的文件存在检查（修复沙盒问题）
Private Function FileExists(filePath As String) As Boolean
    On Error Resume Next

    ' 方法1: 使用Dir函数
    Dim testVar As String
    testVar = Dir(filePath)
    If testVar <> "" Then
        FileExists = True
        On Error GoTo 0
        Exit Function
    End If

    ' 方法2: 尝试打开文件进行读取测试（适用于沙盒环境）
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Input As fileNum

    ' 如果能打开，说明文件存在
    FileExists = True
    Close fileNum
    On Error GoTo 0
    Exit Function

ErrorHandler:
    FileExists = False
    If fileNum > 0 Then Close fileNum
    On Error GoTo 0
End Function

' 获取配置文件路径
Private Function GetConfigFilePath() As String
    Debug.Print "开始获取配置文件路径..."
    
    Dim configPath As String
    Dim testPaths(0 To 3) As String
    Dim i As Integer
    
    If IsMac() Then
        ' Mac路径
        testPaths(0) = Environ("HOME") & "/Desktop/AIProject/AI-Execl/.api_params.txt"
        testPaths(1) = ThisWorkbook.Path & "/.api_params.txt"
        testPaths(2) = Environ("HOME") & "/.api_params.txt"
        testPaths(3) = ThisWorkbook.Path & "/api_config.txt"
    Else
        ' Windows路径
        testPaths(0) = ThisWorkbook.Path & "\.api_params.txt"
        testPaths(1) = Environ("USERPROFILE") & "\.api_params.txt"
        testPaths(2) = ThisWorkbook.Path & "\api_config.txt"
        testPaths(3) = ThisWorkbook.Path & "\config.txt"
    End If
    
    For i = 0 To UBound(testPaths)
        configPath = testPaths(i)
        Debug.Print "检查路径: " & configPath
        
        If FileExists(configPath) Then
            Debug.Print "找到配置文件: " & configPath
            GetConfigFilePath = configPath
            Exit Function
        End If
    Next i
    
    ' 如果没找到，返回默认路径
    If IsMac() Then
        configPath = ThisWorkbook.Path & "/.api_params.txt"
    Else
        configPath = ThisWorkbook.Path & "\.api_params.txt"
    End If
    
    Debug.Print "使用默认路径: " & configPath
    GetConfigFilePath = configPath
End Function

' Mac兼容的文本文件读取（修复沙盒问题）
Private Function ReadTextFile(filePath As String) As String
    Debug.Print "开始读取文件: " & filePath

    On Error GoTo ErrorHandler

    Dim fileNum As Integer
    Dim fileContent As String
    Dim lineContent As String

    fileNum = FreeFile

    ' 直接尝试打开文件，不依赖FileExists检查
    Open filePath For Input As fileNum

    fileContent = ""
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineContent
        If fileContent = "" Then
            fileContent = lineContent
        Else
            fileContent = fileContent & vbCrLf & lineContent
        End If
    Loop

    Close fileNum

    Debug.Print "成功读取文件，内容长度: " & Len(fileContent)
    ReadTextFile = fileContent
    Exit Function

ErrorHandler:
    Debug.Print "读取文件出错: " & Err.Description & " (错误号: " & Err.Number & ")"
    Debug.Print "这可能是Mac沙盒环境的正常行为，文件可能在不同路径"
    If fileNum > 0 Then Close fileNum
    ReadTextFile = ""
End Function

' Mac兼容的文本文件写入
Private Sub WriteTextFile(filePath As String, content As String)
    On Error GoTo ErrorHandler
    
    Debug.Print "开始写入文件: " & filePath
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Output As fileNum
    Print #fileNum, content
    Close fileNum
    
    Debug.Print "成功写入文件"
    Exit Sub
    
ErrorHandler:
    Debug.Print "写入文件出错: " & Err.Description & " (错误号: " & Err.Number & ")"
    If fileNum > 0 Then Close fileNum
    MsgBox "写入配置文件失败: " & Err.Description, vbCritical, "错误"
End Sub

' 解析简单配置格式 (key=value)
Private Sub ParseSimpleConfig(content As String)
    Dim lines As Variant
    Dim i As Integer
    Dim line As String
    Dim pos As Integer
    Dim key As String
    Dim value As String
    
    ' 使用Split分割行
    If InStr(content, vbCrLf) > 0 Then
        lines = Split(content, vbCrLf)
    ElseIf InStr(content, vbLf) > 0 Then
        lines = Split(content, vbLf)
    Else
        ' 单行内容
        ReDim lines(0 To 0)
        lines(0) = content
    End If
    
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        If Len(line) > 0 And Left(line, 1) <> "#" Then ' 忽略空行和注释
            pos = InStr(line, "=")
            If pos > 0 Then
                key = Trim(Left(line, pos - 1))
                value = Trim(Mid(line, pos + 1))
                ' 移除引号
                If Left(value, 1) = """" And Right(value, 1) = """" Then
                    value = Mid(value, 2, Len(value) - 2)
                End If
                AddToConfigDict key, value
                Debug.Print "解析配置: " & key & " = " & value
            End If
        End If
    Next i
End Sub

' 构建简单配置格式
Private Function BuildSimpleConfig() As String
    Dim result As String
    Dim i As Integer
    
    result = "# API配置文件 - " & Format(Now, "yyyy-mm-dd hh:mm:ss") & vbCrLf
    
    For i = 0 To configDictSize - 1
        result = result & configDict(i).Key & "=" & """" & configDict(i).value & """" & vbCrLf
    Next i
    
    BuildSimpleConfig = result
End Function

' Mac兼容的配置设置
Public Sub SetConfig(key As String, value As String)
    On Error GoTo ErrorHandler
    
    Debug.Print "开始设置配置: " & key & " = " & Left(value, 10) & "..."
    
    Dim configPath As String
    Dim configContent As String
    
    configPath = GetConfigFilePath()
    
    ' 初始化配置字典
    InitConfigDict
    
    ' 读取现有配置
    configContent = ReadTextFile(configPath)
    
    If Len(Trim(configContent)) > 0 Then
        ' 解析现有配置（简单的key=value格式）
        ParseSimpleConfig configContent
    End If
    
    ' 设置新值
    AddToConfigDict key, value
    
    ' 保存配置
    Dim newContent As String
    newContent = BuildSimpleConfig()
    WriteTextFile configPath, newContent
    
    Debug.Print "配置设置成功"
    Exit Sub
    
ErrorHandler:
    Debug.Print "设置配置出错: " & Err.Description
    MsgBox "保存配置失败: " & Err.Description & vbCrLf & "请确保有文件写入权限", vbCritical, "错误"
End Sub

' Mac兼容的配置获取
Public Function GetConfig(key As String) As String
    On Error GoTo ErrorHandler
    
    Debug.Print "开始获取配置: " & key
    
    Dim configPath As String
    Dim configContent As String
    
    configPath = GetConfigFilePath()
    configContent = ReadTextFile(configPath)
    
    If Len(Trim(configContent)) = 0 Then
        Debug.Print "配置文件为空或不存在"
        GetConfig = ""
        Exit Function
    End If
    
    ' 初始化并解析配置
    InitConfigDict
    ParseSimpleConfig configContent
    
    ' 获取值
    Dim result As String
    result = GetFromConfigDict(key)
    
    Debug.Print "获取配置结果: " & key & " = " & Left(result, 10) & "..."
    GetConfig = result
    Exit Function
    
ErrorHandler:
    Debug.Print "获取配置出错: " & Err.Description
    GetConfig = ""
End Function

' Mac兼容的配置对话框
Public Sub ShowConfigDialog()
    Debug.Print "=== Mac兼容配置对话框开始 ==="
    
    On Error GoTo ErrorHandler
    
    Dim apiToken As String
    Dim currentPath As String
    
    ' 显示当前配置
    currentPath = GetConfigFilePath()
    Debug.Print "配置文件路径: " & currentPath
    
    ' 安全地获取API Token
    apiToken = GetConfig("api.token")
    If apiToken = "" Then
        Debug.Print "未找到API Token配置"
    Else
        Debug.Print "当前API Token: " & Left(apiToken, 10) & "..."
    End If
    
    Dim msg As String
    msg = "Mac兼容配置管理" & vbCrLf & vbCrLf
    msg = msg & "当前配置:" & vbCrLf
    msg = msg & "系统: Mac" & vbCrLf
    msg = msg & "配置文件: " & currentPath & vbCrLf
    msg = msg & "API Token: " & IIf(Len(apiToken) > 0, Left(apiToken, 10) & "...", "未设置") & vbCrLf & vbCrLf
    msg = msg & "是否要更新API Token?"
    
    If MsgBox(msg, vbYesNo + vbQuestion, "API配置 - Mac兼容版") = vbYes Then
        Dim newToken As String
        newToken = InputBox("请输入lixinger API Token:" & vbCrLf & "(不使用ActiveX组件，Mac完全兼容)", "API配置", apiToken)
        
        If Len(Trim(newToken)) > 0 Then
            Debug.Print "开始保存新的API Token..."
            SetConfig "api.token", newToken
            
            ' 验证保存是否成功
            Dim savedToken As String
            savedToken = GetConfig("api.token")
            If savedToken = newToken Then
                MsgBox "API Token已成功保存！" & vbCrLf & "文件位置: " & currentPath, vbInformation, "配置更新成功"
                Debug.Print "API Token保存成功"
            Else
                MsgBox "API Token保存可能有问题，请检查配置文件", vbExclamation, "警告"
                Debug.Print "API Token保存验证失败"
            End If
        Else
            Debug.Print "用户取消了Token输入"
        End If
    Else
        Debug.Print "用户选择不更新Token"
    End If
    
    Debug.Print "=== Mac兼容配置对话框结束 ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "ShowConfigDialog 出错: " & Err.Description & " (错误号: " & Err.Number & ")"
    MsgBox "配置对话框出错: " & Err.Description & vbCrLf & vbCrLf & "这是Mac兼容版本，完全不依赖ActiveX组件。" & vbCrLf & "如果仍有问题，请检查文件访问权限。", vbCritical, "Mac兼容配置错误"
End Sub

' 配置验证
Public Function ValidateConfig() As Boolean
    On Error GoTo ErrorHandler
    
    Dim apiToken As String
    apiToken = GetConfig("api.token")
    
    If Len(Trim(apiToken)) = 0 Then
        ValidateConfig = False
        Debug.Print "API Token未配置"
    Else
        ValidateConfig = True
        Debug.Print "配置验证通过"
    End If
    
    Exit Function
    
ErrorHandler:
    ValidateConfig = False
    Debug.Print "配置验证出错: " & Err.Description
End Function

' 测试配置系统（增强版，解决Mac沙盒问题）
Public Sub TestConfigSystem()
    Debug.Print "=== 测试Mac兼容配置系统 ==="

    ' 测试写入
    Dim testKey As String
    Dim testValueWrite As String
    testKey = "test.key"
    testValueWrite = "test_value_" & Format(Now, "hhmmss")

    Debug.Print "准备写入: " & testKey & " = " & testValueWrite
    SetConfig testKey, testValueWrite

    ' 立即测试读取（不重新获取路径）
    Debug.Print "立即测试读取..."
    Dim testValue As String
    testValue = GetConfig(testKey)

    Debug.Print "读取结果: " & testValue

    ' 结果判断和显示
    If testValue <> "" Then
        If testValue = testValueWrite Then
            MsgBox "✅ 配置系统测试完全成功！" & vbCrLf & _
                   "写入值: " & testValueWrite & vbCrLf & _
                   "读取值: " & testValue & vbCrLf & _
                   "Mac沙盒环境工作正常！", vbInformation, "测试成功"
            Debug.Print "测试结果: 完全成功"
        Else
            MsgBox "⚠️ 配置系统部分成功" & vbCrLf & _
                   "写入值: " & testValueWrite & vbCrLf & _
                   "读取值: " & testValue & vbCrLf & _
                   "存在数据不一致", vbExclamation, "部分成功"
            Debug.Print "测试结果: 数据不一致"
        End If
    Else
        ' 尝试诊断问题
        Dim configPath As String
        configPath = GetConfigFilePath()

        MsgBox "❌ 配置系统测试失败！" & vbCrLf & _
               "配置路径: " & configPath & vbCrLf & _
               "这可能是Mac Excel沙盒权限问题" & vbCrLf & _
               "请检查Excel的文件访问权限", vbCritical, "测试失败"
        Debug.Print "测试结果: 读取失败"
    End If
    
    Debug.Print "=== 配置系统测试完成 ==="
End Sub
