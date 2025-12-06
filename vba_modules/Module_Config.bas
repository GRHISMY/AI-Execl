' Attribute VB_Name = "Module_Config"
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

' Mac兼容的文件存在检查（严格权限验证版）
Private Function FileExists(filePath As String) As Boolean
    On Error GoTo ErrorHandler

    Dim fileNum As Integer
    fileNum = FreeFile

    ' 直接尝试打开文件进行读取测试（最可靠的方法）
    Open filePath For Input As fileNum
    Close fileNum

    ' 如果能打开并关闭，说明文件存在且可读
    FileExists = True
    Debug.Print "  ✓ 文件存在且可读: " & filePath
    Exit Function

ErrorHandler:
    FileExists = False
    If fileNum > 0 Then Close fileNum
    Debug.Print "  ✗ 文件不存在或不可读: " & filePath & " (错误: " & Err.Description & ")"
End Function

' 获取配置文件路径（Mac沙盒优化版）
Private Function GetConfigFilePath() As String
    Debug.Print "开始获取配置文件路径（Mac沙盒优化版）..."

    Dim configPath As String
    Dim testPaths(0 To 6) As String
    Dim i As Integer

    If IsMac() Then
        ' Mac沙盒优化路径 - 按优先级排序
        testPaths(0) = ThisWorkbook.Path & "/.api_params.txt"  ' 工作簿同目录（最安全）
        testPaths(1) = Application.Path & "/.api_params.txt"    ' Excel应用目录
        testPaths(2) = Environ("TMPDIR") & ".api_params.txt"    ' 临时目录
        testPaths(3) = Environ("HOME") & "/Documents/.api_params.txt"  ' Documents目录
        testPaths(4) = Environ("HOME") & "/Desktop/.api_params.txt"    ' Desktop根目录
        testPaths(5) = "/tmp/.api_params.txt"                   ' 系统临时目录
        testPaths(6) = ThisWorkbook.Path & "/api_config.txt"    ' 备选文件名
    Else
        ' Windows路径
        testPaths(0) = ThisWorkbook.Path & "\.api_params.txt"
        testPaths(1) = Environ("USERPROFILE") & "\.api_params.txt"
        testPaths(2) = Environ("TEMP") & "\.api_params.txt"
        testPaths(3) = ThisWorkbook.Path & "\api_config.txt"
        testPaths(4) = Environ("APPDATA") & "\.api_params.txt"
        testPaths(5) = "C:\temp\.api_params.txt"
        testPaths(6) = ThisWorkbook.Path & "\config.txt"
    End If

    For i = 0 To UBound(testPaths)
        configPath = testPaths(i)
        Debug.Print "检查路径 " & (i + 1) & ": " & configPath

        If FileExists(configPath) Then
            Debug.Print "✓ 找到配置文件: " & configPath
            GetConfigFilePath = configPath
            Exit Function
        End If

        ' 测试路径可写性
        If TestPathWritable(configPath) Then
            Debug.Print "✓ 路径可写，选择: " & configPath
            GetConfigFilePath = configPath
            Exit Function
        End If
    Next i

    ' 如果所有路径都不可用，使用工作簿目录作为最后选择
    If IsMac() Then
        configPath = ThisWorkbook.Path & "/.api_params.txt"
    Else
        configPath = ThisWorkbook.Path & "\.api_params.txt"
    End If

    Debug.Print "⚠️ 使用默认路径: " & configPath
    GetConfigFilePath = configPath
End Function

' 测试路径是否可写
Private Function TestPathWritable(filePath As String) As Boolean
    On Error GoTo ErrorHandler

    Dim fileNum As Integer
    fileNum = FreeFile

    ' 尝试创建临时文件测试可写性
    Open filePath & ".test" For Output As fileNum
    Print #fileNum, "test"
    Close fileNum

    ' 删除测试文件
    Kill filePath & ".test"

    TestPathWritable = True
    Debug.Print "  → 路径可写"
    Exit Function

ErrorHandler:
    TestPathWritable = False
    If fileNum > 0 Then Close fileNum
    Debug.Print "  → 路径不可写: " & Err.Description
    On Error Resume Next
    Kill filePath & ".test"  ' 清理可能存在的测试文件
    On Error GoTo 0
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

' 配置验证
Public Function ValidateConfig() As Boolean
    On Error GoTo ErrorHandler

    Dim apiToken As String
    apiToken = GetConfig("api.token")

    If Len(Trim(apiToken)) = 0 Then
        ValidateConfig = False
    Else
        ValidateConfig = True
    End If

    Exit Function

ErrorHandler:
    ValidateConfig = False
End Function

' 获取Excel相关配置
Public Function GetExcelConfig(key As String, Optional defaultValue As Variant = "") As Variant
    On Error Resume Next

    Select Case key
        Case "etf_codes_column"
            GetExcelConfig = "A"
        Case "prices_column"
            GetExcelConfig = "B"
        Case "status_column"
            GetExcelConfig = "C"
        Case "update_time_column"
            GetExcelConfig = "D"
        Case "start_row"
            GetExcelConfig = 2
        Case Else
            GetExcelConfig = defaultValue
    End Select
End Function

' 显示配置对话框
Public Sub ShowConfigDialog()
    On Error GoTo ErrorHandler

    Dim apiToken As String
    apiToken = GetConfig("api.token")

    ' 显示输入框
    Dim inputToken As String
    inputToken = InputBox("请输入API Token:" & vbCrLf & vbCrLf & _
                         "当前Token: " & Left(apiToken, 10) & "..." & vbCrLf & _
                         "(如果为空，请输入完整Token)", _
                         "ETF数据获取 - API配置", apiToken)

    ' 用户点击取消
    If inputToken = "" And apiToken = "" Then
        Exit Sub
    End If

    ' 用户输入了新Token
    If inputToken <> "" Then
        SetConfig "api.token", inputToken
        MsgBox "API Token已保存！" & vbCrLf & vbCrLf & _
               "✅ Mac沙盒兼容配置已生效", vbInformation, "配置成功"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "配置对话框出错: " & Err.Description & vbCrLf & vbCrLf & _
           "这是Mac兼容版本，完全不依赖ActiveX组件。" & vbCrLf & _
           "如果仍有问题，请检查文件访问权限。", vbCritical, "Mac兼容配置错误"
End Sub