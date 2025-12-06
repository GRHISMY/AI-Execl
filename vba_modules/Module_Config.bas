Attribute VB_Name = "Module_Config"
'
' Module_Config - 配置管理模块
' 负责API token和应用配置的管理
'
Option Explicit

' 配置文件路径
Private Const CONFIG_FILE_NAME As String = "api_config.json"
Private configData As Object

' 加载配置
Public Function LoadConfig() As Boolean
    On Error GoTo ErrorHandler
    
    Dim configPath As String
    configPath = GetConfigFilePath()
    
    If Len(configPath) = 0 Or Dir(configPath) = "" Then
        ' 创建默认配置
        Set configData = GetDefaultConfig()
        LoadConfig = SaveConfig()
    Else
        ' 读取现有配置
        Dim fileContent As String
        fileContent = ReadTextFile(configPath)
        
        If Len(fileContent) > 0 Then
            Set configData = JsonConverter.ParseJSON(fileContent)
            LoadConfig = True
        Else
            Set configData = GetDefaultConfig()
            LoadConfig = SaveConfig()
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    Set configData = GetDefaultConfig()
    LoadConfig = False
    Debug.Print "加载配置错误: " & Err.Description
End Function

' 保存配置
Public Function SaveConfig() As Boolean
    On Error GoTo ErrorHandler
    
    If configData Is Nothing Then
        SaveConfig = False
        Exit Function
    End If
    
    Dim configPath As String
    configPath = GetConfigFilePath()
    
    If Len(configPath) = 0 Then
        SaveConfig = False
        Exit Function
    End If
    
    ' 转换为JSON字符串
    Dim jsonContent As String
    jsonContent = JsonConverter.ConvertToJSON(configData)
    
    ' 写入文件
    SaveConfig = WriteTextFile(configPath, jsonContent)
    
    Exit Function
    
ErrorHandler:
    SaveConfig = False
    Debug.Print "保存配置错误: " & Err.Description
End Function

' 获取配置值
Public Function GetConfig(key As String, Optional defaultValue As Variant = "") As Variant
    On Error GoTo ErrorHandler
    
    If configData Is Nothing Then
        LoadConfig
    End If
    
    If configData Is Nothing Then
        GetConfig = defaultValue
        Exit Function
    End If
    
    ' 支持嵌套key，如 "api.token"
    Dim keys As Variant
    keys = Split(key, ".")
    
    Dim currentObj As Object
    Set currentObj = configData
    
    Dim i As Integer
    For i = 0 To UBound(keys)
        If currentObj.Exists(keys(i)) Then
            If i = UBound(keys) Then
                GetConfig = currentObj(keys(i))
            Else
                Set currentObj = currentObj(keys(i))
            End If
        Else
            GetConfig = defaultValue
            Exit Function
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    GetConfig = defaultValue
    Debug.Print "获取配置错误: " & Err.Description
End Function

' 设置配置值
Public Function SetConfig(key As String, value As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    If configData Is Nothing Then
        LoadConfig
    End If
    
    If configData Is Nothing Then
        SetConfig = False
        Exit Function
    End If
    
    ' 支持嵌套key设置
    Dim keys As Variant
    keys = Split(key, ".")
    
    Dim currentObj As Object
    Set currentObj = configData
    
    Dim i As Integer
    For i = 0 To UBound(keys) - 1
        If Not currentObj.Exists(keys(i)) Then
            Set currentObj(keys(i)) = CreateObject("Scripting.Dictionary")
        End If
        Set currentObj = currentObj(keys(i))
    Next i
    
    currentObj(keys(UBound(keys))) = value
    
    SetConfig = SaveConfig()
    
    Exit Function
    
ErrorHandler:
    SetConfig = False
    Debug.Print "设置配置错误: " & Err.Description
End Function

' 验证配置
Public Function ValidateConfig() As Boolean
    On Error GoTo ErrorHandler
    
    If configData Is Nothing Then
        LoadConfig
    End If
    
    If configData Is Nothing Then
        ValidateConfig = False
        Exit Function
    End If
    
    ' 检查必要配置项
    Dim apiToken As String
    apiToken = GetConfig("api.token")
    
    If Len(Trim(apiToken)) = 0 Or apiToken = "YOUR_API_TOKEN_HERE" Then
        ValidateConfig = False
        Exit Function
    End If
    
    ' 检查其他必要配置
    Dim baseUrl As String
    baseUrl = GetConfig("api.base_url")
    
    If Len(Trim(baseUrl)) = 0 Then
        ValidateConfig = False
        Exit Function
    End If
    
    ValidateConfig = True
    Exit Function
    
ErrorHandler:
    ValidateConfig = False
    Debug.Print "配置验证错误: " & Err.Description
End Function

' 重置为默认配置
Public Sub ResetToDefault()
    On Error GoTo ErrorHandler
    
    Set configData = GetDefaultConfig()
    SaveConfig
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "重置配置错误: " & Err.Description
End Sub

' 获取默认配置
Private Function GetDefaultConfig() As Object
    On Error GoTo ErrorHandler
    
    Dim config As Object
    Set config = CreateObject("Scripting.Dictionary")
    
    ' API配置
    Dim apiConfig As Object
    Set apiConfig = CreateObject("Scripting.Dictionary")
    apiConfig("base_url") = "https://open.lixinger.com"
    apiConfig("token") = "YOUR_API_TOKEN_HERE"
    apiConfig("timeout") = 30
    apiConfig("max_retries") = 3
    apiConfig("rate_limit") = 0.5
    
    ' 应用配置
    Dim appConfig As Object
    Set appConfig = CreateObject("Scripting.Dictionary")
    appConfig("log_level") = "INFO"
    appConfig("cache_enabled") = True
    appConfig("batch_size") = 20
    appConfig("executable_path") = ""
    
    config("api") = apiConfig
    config("app") = appConfig
    
    Set GetDefaultConfig = config
    Exit Function
    
ErrorHandler:
    Set GetDefaultConfig = Nothing
    Debug.Print "获取默认配置错误: " & Err.Description
End Function

' 获取配置文件路径
Private Function GetConfigFilePath() As String
    On Error GoTo ErrorHandler
    
    ' 优先使用工作簿同目录
    Dim configPath As String
    configPath = ThisWorkbook.Path & "/config/" & CONFIG_FILE_NAME
    
    ' 如果不存在，尝试其他位置
    If Dir(configPath) = "" Then
        configPath = ThisWorkbook.Path & "/" & CONFIG_FILE_NAME
    End If
    
    GetConfigFilePath = configPath
    Exit Function
    
ErrorHandler:
    GetConfigFilePath = ""
    Debug.Print "获取配置文件路径错误: " & Err.Description
End Function

' 读取文本文件
Private Function ReadTextFile(filePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Input As #fileNum
    ReadTextFile = Input$(LOF(fileNum), fileNum)
    Close #fileNum
    
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    ReadTextFile = ""
    Debug.Print "读取文件错误: " & Err.Description
End Function

' 写入文本文件
Private Function WriteTextFile(filePath As String, content As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Output As #fileNum
    Print #fileNum, content
    Close #fileNum
    
    WriteTextFile = True
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    WriteTextFile = False
    Debug.Print "写入文件错误: " & Err.Description
End Function

' 显示配置对话框
Public Sub ShowConfigDialog()
    On Error GoTo ErrorHandler
    
    Dim apiToken As String
    apiToken = GetConfig("api.token")
    
    ' 如果是默认token，显示为空
    If apiToken = "YOUR_API_TOKEN_HERE" Then
        apiToken = ""
    End If
    
    Dim newToken As String
    newToken = InputBox("请输入lixinger API Token:", "API配置", apiToken)
    
    If Len(Trim(newToken)) > 0 Then
        SetConfig "api.token", newToken
        MsgBox "API Token已保存", vbInformation, "配置更新"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "配置对话框错误: " & Err.Description, vbExclamation, "配置错误"
End Sub

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
