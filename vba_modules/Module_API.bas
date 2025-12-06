' Attribute VB_Name = "Module_API"
'
' Module_API - API接口模块
' 负责调用Python可执行文件获取ETF数据
'
Option Explicit

' 主要API调用函数
Public Function CallETFAPI(etfCodes As String) As String
    On Error GoTo ErrorHandler
    
    ' 验证ETF代码
    If Len(Trim(etfCodes)) = 0 Then
        CallETFAPI = "{""status"":""error"",""error"":""ETF代码不能为空""}"
        Exit Function
    End If
    
    ' 获取可执行文件路径
    Dim execPath As String
    execPath = GetExecutablePath()
    
    If Len(execPath) = 0 Then
        CallETFAPI = "{""status"":""error"",""error"":""未找到API可执行文件""}"
        Exit Function
    End If
    
    ' 构建命令行
    Dim command As String
    command = """" & execPath & """ --codes """ & etfCodes & """"
    
    ' 添加配置文件路径（如果存在）
    Dim configPath As String
    configPath = GetConfigPath()
    If Len(configPath) > 0 Then
        command = command & " --config """ & configPath & """"
    End If
    
    ' 执行命令
    Dim result As String
    result = ExecuteShellCommand(command)
    
    CallETFAPI = result
    Exit Function
    
ErrorHandler:
    CallETFAPI = "{""status"":""error"",""error"":""API调用失败: " & Err.Description & """}"
    Debug.Print "API调用错误: " & Err.Description
End Function

' 获取可执行文件路径
Public Function GetExecutablePath() As String
    On Error GoTo ErrorHandler
    
    ' 可能的路径列表
    Dim paths As Variant
    Dim i As Integer
    
    ' 检测系统架构
    Dim isAppleSilicon As Boolean
    isAppleSilicon = IsAppleSiliconMac()
    
    ' 根据架构设定搜索路径
    If isAppleSilicon Then
        paths = Array( _
            ThisWorkbook.Path & "/etf_api_caller_arm64", _
            ThisWorkbook.Path & "/../dist/etf_api_caller_arm64", _
            ThisWorkbook.Path & "/dist/etf_api_caller_arm64", _
            ThisWorkbook.Path & "/etf_api_caller", _
            ThisWorkbook.Path & "/../dist/etf_api_caller", _
            ThisWorkbook.Path & "/dist/etf_api_caller" _
        )
    Else
        paths = Array( _
            ThisWorkbook.Path & "/etf_api_caller", _
            ThisWorkbook.Path & "/../dist/etf_api_caller", _
            ThisWorkbook.Path & "/dist/etf_api_caller", _
            ThisWorkbook.Path & "/etf_api_caller_arm64", _
            ThisWorkbook.Path & "/../dist/etf_api_caller_arm64", _
            ThisWorkbook.Path & "/dist/etf_api_caller_arm64" _
        )
    End If
    
    ' 查找存在的可执行文件
    For i = 0 To UBound(paths)
        If Dir(paths(i)) <> "" Then
            GetExecutablePath = paths(i)
            Exit Function
        End If
    Next i
    
    GetExecutablePath = ""
    Exit Function
    
ErrorHandler:
    GetExecutablePath = ""
    Debug.Print "获取可执行文件路径错误: " & Err.Description
End Function

' 检测是否为Apple Silicon Mac
Private Function IsAppleSiliconMac() As Boolean
    On Error GoTo ErrorHandler
    
    ' 通过系统命令检测架构
    Dim result As String
    result = ExecuteShellCommand("uname -m")
    
    IsAppleSiliconMac = (InStr(result, "arm64") > 0)
    Exit Function
    
ErrorHandler:
    IsAppleSiliconMac = False
    Debug.Print "检测Apple Silicon错误: " & Err.Description
End Function

' 获取配置文件路径
Public Function GetConfigPath() As String
    On Error GoTo ErrorHandler
    
    Dim paths As Variant
    Dim i As Integer
    
    paths = Array( _
        ThisWorkbook.Path & "/config/api_config.json", _
        ThisWorkbook.Path & "/../config/api_config.json", _
        Environ("HOME") & "/.etf_config.json" _
    )
    
    For i = 0 To UBound(paths)
        If Dir(paths(i)) <> "" Then
            GetConfigPath = paths(i)
            Exit Function
        End If
    Next i
    
    GetConfigPath = ""
    Exit Function
    
ErrorHandler:
    GetConfigPath = ""
    Debug.Print "获取配置文件路径错误: " & Err.Description
End Function

' 验证ETF代码格式
Public Function ValidateETFCodes(codes As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Len(Trim(codes)) = 0 Then
        ValidateETFCodes = False
        Exit Function
    End If
    
    ' 分割代码
    Dim codeArray As Variant
    codeArray = Split(codes, ",")
    
    Dim i As Integer
    For i = 0 To UBound(codeArray)
        Dim code As String
        code = Trim(codeArray(i))
        
        ' 验证单个代码格式
        If Not ValidateSingleETFCode(code) Then
            ValidateETFCodes = False
            Exit Function
        End If
    Next i
    
    ValidateETFCodes = True
    Exit Function
    
ErrorHandler:
    ValidateETFCodes = False
    Debug.Print "ETF代码验证错误: " & Err.Description
End Function

' 验证单个ETF代码
Private Function ValidateSingleETFCode(code As String) As Boolean
    On Error GoTo ErrorHandler
    
    code = Trim(code)
    
    ' 检查长度
    If Len(code) <> 6 Then
        ValidateSingleETFCode = False
        Exit Function
    End If
    
    ' 检查是否全是数字
    Dim i As Integer
    For i = 1 To Len(code)
        If Not IsNumeric(Mid(code, i, 1)) Then
            ValidateSingleETFCode = False
            Exit Function
        End If
    Next i
    
    ' 检查ETF代码范围
    Dim codeNum As Long
    codeNum = CLng(code)
    
    If (codeNum >= 159000 And codeNum <= 159999) Or _
       (codeNum >= 510000 And codeNum <= 519999) Or _
       (codeNum >= 560000 And codeNum <= 569999) Then
        ValidateSingleETFCode = True
    Else
        ValidateSingleETFCode = False
    End If
    
    Exit Function
    
ErrorHandler:
    ValidateSingleETFCode = False
    Debug.Print "单个ETF代码验证错误: " & Err.Description
End Function

' 执行Shell命令
Public Function ExecuteShellCommand(command As String) As String
    On Error GoTo ErrorHandler
    
    ' 使用AppleScript在Mac上执行命令
    Dim script As String
    script = "do shell script """ & command & """"
    
    Dim result As String
    result = MacScript(script)
    
    ExecuteShellCommand = result
    Exit Function
    
ErrorHandler:
    ExecuteShellCommand = ""
    Debug.Print "Shell命令执行错误: " & Err.Description & " Command: " & command
End Function

' 解析API响应
Public Function ParseAPIResponse(response As String) As Object
    On Error GoTo ErrorHandler
    
    ' 使用JsonConverter解析
    Set ParseAPIResponse = JsonConverter.ParseJSON(response)
    Exit Function
    
ErrorHandler:
    Set ParseAPIResponse = Nothing
    Debug.Print "API响应解析错误: " & Err.Description
End Function

' 处理API错误
Public Sub HandleAPIError(errorMsg As String)
    On Error Resume Next
    
    ' 在状态栏显示错误
    Application.StatusBar = "ETF API错误: " & errorMsg
    
    ' 记录到调试窗口
    Debug.Print "API错误: " & errorMsg & " - " & Now()
    
    ' 可以添加用户提示
    MsgBox "ETF数据获取失败: " & errorMsg, vbExclamation, "API错误"
End Sub

' 测试API连接
Public Function TestAPIConnection() As Boolean
    On Error GoTo ErrorHandler
    
    ' 使用测试ETF代码
    Dim testResult As String
    testResult = CallETFAPI("159915")
    
    ' 解析结果
    Dim jsonResult As Object
    Set jsonResult = ParseAPIResponse(testResult)
    
    If Not jsonResult Is Nothing Then
        If jsonResult.Exists("status") Then
            TestAPIConnection = (jsonResult("status") = "success")
        Else
            TestAPIConnection = False
        End If
    Else
        TestAPIConnection = False
    End If
    
    Exit Function
    
ErrorHandler:
    TestAPIConnection = False
    Debug.Print "API连接测试错误: " & Err.Description
End Function
