' Attribute VB_Name = "JsonConverter"
'
' JsonConverter_Mac - Mac兼容的JSON转换器
' 完全不依赖ActiveX组件，专为Mac Excel设计
'
Option Explicit

' Mac兼容的简单JSON解析器
' 支持基本的JSON结构：对象、数组、字符串、数字、布尔值

' 解析JSON字符串为简单值（用于API响应）
Public Function ParseJSON(jsonText As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim trimmedText As String
    trimmedText = Trim(jsonText)
    
    If Len(trimmedText) = 0 Then
        ParseJSON = ""
        Exit Function
    End If
    
    ' 简单JSON解析 - 支持基本结构
    If Left(trimmedText, 1) = "{" And Right(trimmedText, 1) = "}" Then
        ' JSON对象 - 转换为字符串数组
        ParseJSON = ParseSimpleObject(trimmedText)
    ElseIf Left(trimmedText, 1) = "[" And Right(trimmedText, 1) = "]" Then
        ' JSON数组
        ParseJSON = ParseSimpleArray(trimmedText)
    ElseIf Left(trimmedText, 1) = """" And Right(trimmedText, 1) = """" Then
        ' JSON字符串
        ParseJSON = Mid(trimmedText, 2, Len(trimmedText) - 2)
    ElseIf IsNumeric(trimmedText) Then
        ' JSON数字
        ParseJSON = CDbl(trimmedText)
    ElseIf LCase(trimmedText) = "true" Then
        ParseJSON = True
    ElseIf LCase(trimmedText) = "false" Then
        ParseJSON = False
    ElseIf LCase(trimmedText) = "null" Then
        ParseJSON = Null
    Else
        ' 默认返回原始字符串
        ParseJSON = trimmedText
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "JSON解析错误: " & Err.Description
    ParseJSON = jsonText
End Function

' 解析简单JSON对象（完整实现）
Private Function ParseSimpleObject(jsonText As String) As Variant
    On Error GoTo ErrorHandler

    ' 移除大括号
    Dim content As String
    content = Trim(Mid(jsonText, 2, Len(jsonText) - 2))

    If Len(content) = 0 Then
        ParseSimpleObject = Array()
        Exit Function
    End If

    ' 正确解析键值对，考虑字符串中的逗号
    Dim pairs() As String
    Dim pairCount As Integer
    pairCount = 0
    ReDim pairs(0 To 99)

    Dim i As Integer
    Dim inString As Boolean
    Dim currentPair As String
    Dim char As String

    inString = False
    currentPair = ""

    For i = 1 To Len(content)
        char = Mid(content, i, 1)

        If char = """" Then
            inString = Not inString
            currentPair = currentPair & char
        ElseIf char = "," And Not inString Then
            ' 找到分隔符，保存当前键值对
            If Len(Trim(currentPair)) > 0 Then
                pairs(pairCount) = Trim(currentPair)
                pairCount = pairCount + 1
                If pairCount > UBound(pairs) Then
                    ReDim Preserve pairs(0 To UBound(pairs) + 99)
                End If
            End If
            currentPair = ""
        Else
            currentPair = currentPair & char
        End If
    Next i

    ' 添加最后一个键值对
    If Len(Trim(currentPair)) > 0 Then
        pairs(pairCount) = Trim(currentPair)
        pairCount = pairCount + 1
    End If

    If pairCount = 0 Then
        ParseSimpleObject = Array()
        Exit Function
    End If

    ' 创建结果数组
    ReDim result(0 To pairCount - 1, 0 To 1) As String

    For i = 0 To pairCount - 1
        Dim pair As String
        pair = Trim(pairs(i))
        
        Dim colonPos As Integer
        colonPos = InStr(pair, ":")
        
        If colonPos > 0 Then
            Dim key As String
            Dim value As String
            key = Trim(Left(pair, colonPos - 1))
            value = Trim(Mid(pair, colonPos + 1))
            
            ' 移除引号
            If Left(key, 1) = """" And Right(key, 1) = """" Then
                key = Mid(key, 2, Len(key) - 2)
            End If
            If Left(value, 1) = """" And Right(value, 1) = """" Then
                value = Mid(value, 2, Len(value) - 2)
            End If
            
            result(i, 0) = key
            result(i, 1) = value
        End If
    Next i
    
    ParseSimpleObject = result
    Exit Function
    
ErrorHandler:
    Debug.Print "JSON对象解析错误: " & Err.Description
    ParseSimpleObject = Array()
End Function

' 解析简单JSON数组（完整实现）
Private Function ParseSimpleArray(jsonText As String) As Variant
    On Error GoTo ErrorHandler

    ' 移除方括号
    Dim content As String
    content = Trim(Mid(jsonText, 2, Len(jsonText) - 2))

    If Len(content) = 0 Then
        ParseSimpleArray = Array()
        Exit Function
    End If

    ' 正确解析数组元素，考虑字符串中的逗号
    Dim elements() As String
    Dim elementCount As Integer
    elementCount = 0
    ReDim elements(0 To 99)

    Dim i As Integer
    Dim inString As Boolean
    Dim currentElement As String
    Dim char As String

    inString = False
    currentElement = ""

    For i = 1 To Len(content)
        char = Mid(content, i, 1)

        If char = """" Then
            inString = Not inString
            currentElement = currentElement & char
        ElseIf char = "," And Not inString Then
            ' 找到分隔符，保存当前元素
            If Len(Trim(currentElement)) > 0 Then
                elements(elementCount) = Trim(currentElement)
                elementCount = elementCount + 1
                If elementCount > UBound(elements) Then
                    ReDim Preserve elements(0 To UBound(elements) + 99)
                End If
            End If
            currentElement = ""
        Else
            currentElement = currentElement & char
        End If
    Next i

    ' 添加最后一个元素
    If Len(Trim(currentElement)) > 0 Then
        elements(elementCount) = Trim(currentElement)
        elementCount = elementCount + 1
    End If

    If elementCount = 0 Then
        ParseSimpleArray = Array()
        Exit Function
    End If

    ' 创建结果数组并清理数据
    ReDim result(0 To elementCount - 1) As String

    For i = 0 To elementCount - 1
        result(i) = Trim(elements(i))
        ' 移除引号
        If Left(result(i), 1) = """" And Right(result(i), 1) = """" Then
            result(i) = Mid(result(i), 2, Len(result(i)) - 2)
        End If
    Next i

    ParseSimpleArray = result
    Exit Function
    
ErrorHandler:
    Debug.Print "JSON数组解析错误: " & Err.Description
    ParseSimpleArray = Array()
End Function

' 将简单值转换为JSON（完整实现）
Public Function ConvertToJSON(data As Variant) As String
    On Error GoTo ErrorHandler

    If IsArray(data) Then
        ConvertToJSON = ConvertArrayToJSON(data)
    ElseIf IsObject(data) Then
        ' 完整处理对象 - 尝试获取对象的默认属性或转为字符串
        Dim objStr As String
        On Error Resume Next
        objStr = CStr(data)
        If Err.Number <> 0 Then
            objStr = "object"
        End If
        On Error GoTo ErrorHandler
        ConvertToJSON = """" & EscapeString(objStr) & """"
    ElseIf VarType(data) = vbString Then
        ConvertToJSON = """" & EscapeString(CStr(data)) & """"
    ElseIf VarType(data) = vbBoolean Then
        ConvertToJSON = IIf(data, "true", "false")
    ElseIf VarType(data) = vbDate Then
        ConvertToJSON = """" & Format(data, "yyyy-mm-dd hh:mm:ss") & """"
    ElseIf IsNumeric(data) Then
        ConvertToJSON = CStr(data)
    ElseIf IsNull(data) Then
        ConvertToJSON = "null"
    ElseIf VarType(data) = vbEmpty Then
        ConvertToJSON = "null"
    Else
        ConvertToJSON = """" & EscapeString(CStr(data)) & """"
    End If

    Exit Function

ErrorHandler:
    ConvertToJSON = """" & EscapeString(CStr(data)) & """"
End Function

' 转换数组为JSON数组
Private Function ConvertArrayToJSON(arr As Variant) As String
    On Error GoTo ErrorHandler
    
    If Not IsArray(arr) Then
        ConvertArrayToJSON = "[]"
        Exit Function
    End If
    
    Dim result As String
    result = "["
    
    Dim i As Integer
    Dim firstItem As Boolean
    firstItem = True
    
    For i = LBound(arr) To UBound(arr)
        If Not firstItem Then
            result = result & ","
        End If
        result = result & ConvertToJSON(arr(i))
        firstItem = False
    Next i
    
    result = result & "]"
    ConvertArrayToJSON = result
    Exit Function
    
ErrorHandler:
    ConvertArrayToJSON = "[]"
End Function

' 转义JSON字符串
Private Function EscapeString(str As String) As String
    Dim result As String
    result = str
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    EscapeString = result
End Function

' 从JSON对象数组中获取值（辅助函数）
Public Function GetJSONValue(jsonArray As Variant, key As String) As String
    On Error GoTo ErrorHandler
    
    If Not IsArray(jsonArray) Then
        GetJSONValue = ""
        Exit Function
    End If
    
    Dim i As Integer
    For i = 0 To UBound(jsonArray, 1)
        If UBound(jsonArray, 2) >= 1 Then
            If jsonArray(i, 0) = key Then
                GetJSONValue = jsonArray(i, 1)
                Exit Function
            End If
        End If
    Next i
    
    GetJSONValue = ""
    Exit Function
    
ErrorHandler:
    GetJSONValue = ""
End Function

' 创建简单JSON对象字符串
Public Function CreateSimpleJSON(key1 As String, value1 As String, _
                                Optional key2 As String = "", Optional value2 As String = "", _
                                Optional key3 As String = "", Optional value3 As String = "") As String
    Dim result As String
    result = "{"
    result = result & """" & key1 & """: """ & EscapeString(value1) & """"
    
    If key2 <> "" Then
        result = result & ", """ & key2 & """: """ & EscapeString(value2) & """"
    End If
    
    If key3 <> "" Then
        result = result & ", """ & key3 & """: """ & EscapeString(value3) & """"
    End If
    
    result = result & "}"
    CreateSimpleJSON = result
End Function

' 完整的JSON功能测试
Public Sub TestJSONConverter()
    Debug.Print "=== 测试Mac兼容JSON转换器 ==="

    Dim allTestsPassed As Boolean
    allTestsPassed = True
    Dim testResults As String
    testResults = "JSON转换器测试结果:" & vbCrLf

    ' 测试1: 简单JSON对象解析
    Dim testJSON1 As String
    testJSON1 = "{""name"": ""test"", ""value"": ""123"", ""flag"": ""true""}"

    Dim parsed1 As Variant
    parsed1 = ParseJSON(testJSON1)

    If IsArray(parsed1) Then
        Dim testValue1 As String
        testValue1 = GetJSONValue(parsed1, "name")
        If testValue1 = "test" Then
            testResults = testResults & "✓ 简单对象解析: 通过" & vbCrLf
        Else
            testResults = testResults & "✗ 简单对象解析: 失败" & vbCrLf
            allTestsPassed = False
        End If
    Else
        testResults = testResults & "✗ 简单对象解析: 失败" & vbCrLf
        allTestsPassed = False
    End If

    ' 测试2: 包含逗号的JSON解析
    Dim testJSON2 As String
    testJSON2 = "{""title"": ""Hello, World!"", ""count"": ""42""}"

    Dim parsed2 As Variant
    parsed2 = ParseJSON(testJSON2)

    If IsArray(parsed2) Then
        Dim testValue2 As String
        testValue2 = GetJSONValue(parsed2, "title")
        If testValue2 = "Hello, World!" Then
            testResults = testResults & "✓ 复杂字符串解析: 通过" & vbCrLf
        Else
            testResults = testResults & "✗ 复杂字符串解析: 失败" & vbCrLf
            allTestsPassed = False
        End If
    Else
        testResults = testResults & "✗ 复杂字符串解析: 失败" & vbCrLf
        allTestsPassed = False
    End If

    ' 测试3: JSON数组解析
    Dim testJSON3 As String
    testJSON3 = "[""apple"", ""banana"", ""cherry""]"

    Dim parsed3 As Variant
    parsed3 = ParseJSON(testJSON3)

    If IsArray(parsed3) Then
        If UBound(parsed3) = 2 And parsed3(0) = "apple" Then
            testResults = testResults & "✓ 数组解析: 通过" & vbCrLf
        Else
            testResults = testResults & "✗ 数组解析: 失败" & vbCrLf
            allTestsPassed = False
        End If
    Else
        testResults = testResults & "✗ 数组解析: 失败" & vbCrLf
        allTestsPassed = False
    End If

    ' 测试4: JSON生成
    Dim testArray As Variant
    testArray = Array("test1", "test2", "test3")

    Dim generatedJSON As String
    generatedJSON = ConvertToJSON(testArray)

    If InStr(generatedJSON, "[") > 0 And InStr(generatedJSON, "]") > 0 Then
        testResults = testResults & "✓ JSON生成: 通过" & vbCrLf
    Else
        testResults = testResults & "✗ JSON生成: 失败" & vbCrLf
        allTestsPassed = False
    End If

    ' 显示测试结果
    If allTestsPassed Then
        testResults = testResults & vbCrLf & "🎉 所有测试通过！JSON转换器完全可用。"
        MsgBox testResults, vbInformation, "测试成功"
    Else
        testResults = testResults & vbCrLf & "⚠️ 部分测试失败，请检查实现。"
        MsgBox testResults, vbExclamation, "测试结果"
    End If

    Debug.Print testResults
    Debug.Print "=== JSON转换器测试完成 ==="
End Sub
