'Attribute VB_Name = "JsonConverter"
'
' JsonConverter - 增强的JSON解析器模块
' 专为Mac Excel环境优化，支持复杂嵌套结构的JSON解析
'
Option Explicit

' 全局变量用于解析状态
Private parseIndex As Long
Private jsonContent As String

' JSON解析主函数
Public Function ParseJSON(jsonString As String) As Object
    On Error GoTo ErrorHandler

    ' 初始化解析状态
    jsonContent = Trim(jsonString)
    parseIndex = 1

    ' 开始解析
    Set ParseJSON = ParseValue()
    Exit Function

ErrorHandler:
    Set ParseJSON = Nothing
    Debug.Print "JSON解析错误: " & Err.Description & " at position " & parseIndex
End Function

' 解析任意JSON值
Private Function ParseValue() As Variant
    On Error GoTo ErrorHandler

    ' 跳过空白字符
    SkipWhitespace

    If parseIndex > Len(jsonContent) Then
        ParseValue = Null
        Exit Function
    End If

    Dim char As String
    char = Mid(jsonContent, parseIndex, 1)

    Select Case char
        Case "{"
            Set ParseValue = ParseObject()
        Case "["
            Set ParseValue = ParseArray()
        Case """"
            ParseValue = ParseString()
        Case "t", "f"
            ParseValue = ParseBoolean()
        Case "n"
            ParseValue = ParseNull()
        Case Else
            If IsNumericChar(char) Then
                ParseValue = ParseNumber()
            Else
                Err.Raise vbObjectError + 1001, "JsonConverter", "无效的JSON字符: " & char
            End If
    End Select

    Exit Function

ErrorHandler:
    ParseValue = Null
    Debug.Print "解析值错误: " & Err.Description
End Function

' 解析JSON对象
Private Function ParseObject() As Object
    On Error GoTo ErrorHandler

    Dim obj As Object
    Set obj = CreateObject("Scripting.Dictionary")

    ' 跳过开始的 {
    parseIndex = parseIndex + 1
    SkipWhitespace

    ' 检查空对象
    If parseIndex <= Len(jsonContent) And Mid(jsonContent, parseIndex, 1) = "}" Then
        parseIndex = parseIndex + 1
        Set ParseObject = obj
        Exit Function
    End If

    ' 解析键值对
    Do
        SkipWhitespace

        ' 解析键
        If parseIndex > Len(jsonContent) Or Mid(jsonContent, parseIndex, 1) <> """" Then
            Err.Raise vbObjectError + 1002, "JsonConverter", "期望字符串键"
        End If

        Dim key As String
        key = ParseString()

        SkipWhitespace

        ' 期望冒号
        If parseIndex > Len(jsonContent) Or Mid(jsonContent, parseIndex, 1) <> ":" Then
            Err.Raise vbObjectError + 1003, "JsonConverter", "期望冒号"
        End If
        parseIndex = parseIndex + 1

        SkipWhitespace

        ' 解析值
        Dim value As Variant
        If IsObject(ParseValue()) Then
            Set obj(key) = ParseValue()
        Else
            obj(key) = ParseValue()
        End If

        SkipWhitespace

        ' 检查是否继续
        If parseIndex > Len(jsonContent) Then Exit Do

        Dim nextChar As String
        nextChar = Mid(jsonContent, parseIndex, 1)

        If nextChar = "}" Then
            parseIndex = parseIndex + 1
            Exit Do
        ElseIf nextChar = "," Then
            parseIndex = parseIndex + 1
        Else
            Err.Raise vbObjectError + 1004, "JsonConverter", "期望逗号或右大括号"
        End If
    Loop

    Set ParseObject = obj
    Exit Function

ErrorHandler:
    Set ParseObject = Nothing
    Debug.Print "解析对象错误: " & Err.Description
End Function

' 解析JSON数组
Private Function ParseArray() As Object
    On Error GoTo ErrorHandler

    Dim arr As Object
    Set arr = CreateObject("Scripting.Dictionary")

    ' 跳过开始的 [
    parseIndex = parseIndex + 1
    SkipWhitespace

    ' 检查空数组
    If parseIndex <= Len(jsonContent) And Mid(jsonContent, parseIndex, 1) = "]" Then
        parseIndex = parseIndex + 1
        Set ParseArray = arr
        Exit Function
    End If

    Dim index As Long
    index = 0

    ' 解析数组元素
    Do
        SkipWhitespace

        ' 解析元素
        Dim element As Variant
        Set element = ParseValue()
        If IsObject(element) Then
            Set arr(CStr(index)) = element
        Else
            arr(CStr(index)) = element
        End If

        index = index + 1
        SkipWhitespace

        ' 检查是否继续
        If parseIndex > Len(jsonContent) Then Exit Do

        Dim nextChar As String
        nextChar = Mid(jsonContent, parseIndex, 1)

        If nextChar = "]" Then
            parseIndex = parseIndex + 1
            Exit Do
        ElseIf nextChar = "," Then
            parseIndex = parseIndex + 1
        Else
            Err.Raise vbObjectError + 1005, "JsonConverter", "期望逗号或右方括号"
        End If
    Loop

    Set ParseArray = arr
    Exit Function

ErrorHandler:
    Set ParseArray = Nothing
    Debug.Print "解析数组错误: " & Err.Description
End Function

' 解析字符串
Private Function ParseString() As String
    On Error GoTo ErrorHandler

    ' 跳过开始引号
    parseIndex = parseIndex + 1

    Dim result As String
    result = ""

    Do While parseIndex <= Len(jsonContent)
        Dim char As String
        char = Mid(jsonContent, parseIndex, 1)

        If char = """" Then
            ' 结束引号
            parseIndex = parseIndex + 1
            ParseString = result
            Exit Function
        ElseIf char = "\" Then
            ' 转义字符
            parseIndex = parseIndex + 1
            If parseIndex <= Len(jsonContent) Then
                Dim escapeChar As String
                escapeChar = Mid(jsonContent, parseIndex, 1)
                Select Case escapeChar
                    Case """"
                        result = result & """"
                    Case "\"
                        result = result & "\"
                    Case "/"
                        result = result & "/"
                    Case "b"
                        result = result & Chr(8)
                    Case "f"
                        result = result & Chr(12)
                    Case "n"
                        result = result & vbLf
                    Case "r"
                        result = result & vbCr
                    Case "t"
                        result = result & vbTab
                    Case Else
                        result = result & escapeChar
                End Select
                parseIndex = parseIndex + 1
            End If
        Else
            result = result & char
            parseIndex = parseIndex + 1
        End If
    Loop

    Err.Raise vbObjectError + 1006, "JsonConverter", "未结束的字符串"

ErrorHandler:
    ParseString = ""
    Debug.Print "解析字符串错误: " & Err.Description
End Function

' 解析数字
Private Function ParseNumber() As Variant
    On Error GoTo ErrorHandler

    Dim startPos As Long
    startPos = parseIndex

    ' 处理负号
    If Mid(jsonContent, parseIndex, 1) = "-" Then
        parseIndex = parseIndex + 1
    End If

    ' 解析整数部分
    If Not IsNumericChar(Mid(jsonContent, parseIndex, 1)) Then
        Err.Raise vbObjectError + 1007, "JsonConverter", "无效的数字格式"
    End If

    Do While parseIndex <= Len(jsonContent) And IsNumericChar(Mid(jsonContent, parseIndex, 1))
        parseIndex = parseIndex + 1
    Loop

    ' 处理小数点
    If parseIndex <= Len(jsonContent) And Mid(jsonContent, parseIndex, 1) = "." Then
        parseIndex = parseIndex + 1
        If Not IsNumericChar(Mid(jsonContent, parseIndex, 1)) Then
            Err.Raise vbObjectError + 1008, "JsonConverter", "小数点后需要数字"
        End If

        Do While parseIndex <= Len(jsonContent) And IsNumericChar(Mid(jsonContent, parseIndex, 1))
            parseIndex = parseIndex + 1
        Loop
    End If

    ' 处理科学计数法
    If parseIndex <= Len(jsonContent) And (Mid(jsonContent, parseIndex, 1) = "e" Or Mid(jsonContent, parseIndex, 1) = "E") Then
        parseIndex = parseIndex + 1
        If parseIndex <= Len(jsonContent) And (Mid(jsonContent, parseIndex, 1) = "+" Or Mid(jsonContent, parseIndex, 1) = "-") Then
            parseIndex = parseIndex + 1
        End If

        If Not IsNumericChar(Mid(jsonContent, parseIndex, 1)) Then
            Err.Raise vbObjectError + 1009, "JsonConverter", "指数需要数字"
        End If

        Do While parseIndex <= Len(jsonContent) And IsNumericChar(Mid(jsonContent, parseIndex, 1))
            parseIndex = parseIndex + 1
        Loop
    End If

    Dim numberStr As String
    numberStr = Mid(jsonContent, startPos, parseIndex - startPos)

    If InStr(numberStr, ".") > 0 Or InStr(LCase(numberStr), "e") > 0 Then
        ParseNumber = CDbl(numberStr)
    Else
        ParseNumber = CLng(numberStr)
    End If

    Exit Function

ErrorHandler:
    ParseNumber = 0
    Debug.Print "解析数字错误: " & Err.Description
End Function

' 解析布尔值
Private Function ParseBoolean() As Boolean
    On Error GoTo ErrorHandler

    If Mid(jsonContent, parseIndex, 4) = "true" Then
        parseIndex = parseIndex + 4
        ParseBoolean = True
    ElseIf Mid(jsonContent, parseIndex, 5) = "false" Then
        parseIndex = parseIndex + 5
        ParseBoolean = False
    Else
        Err.Raise vbObjectError + 1010, "JsonConverter", "无效的布尔值"
    End If

    Exit Function

ErrorHandler:
    ParseBoolean = False
    Debug.Print "解析布尔值错误: " & Err.Description
End Function

' 解析null值
Private Function ParseNull() As Variant
    On Error GoTo ErrorHandler

    If Mid(jsonContent, parseIndex, 4) = "null" Then
        parseIndex = parseIndex + 4
        ParseNull = Null
    Else
        Err.Raise vbObjectError + 1011, "JsonConverter", "无效的null值"
    End If

    Exit Function

ErrorHandler:
    ParseNull = Null
    Debug.Print "解析null值错误: " & Err.Description
End Function

' 跳过空白字符
Private Sub SkipWhitespace()
    Do While parseIndex <= Len(jsonContent)
        Dim char As String
        char = Mid(jsonContent, parseIndex, 1)
        If char = " " Or char = vbTab Or char = vbCr Or char = vbLf Then
            parseIndex = parseIndex + 1
        Else
            Exit Do
        End If
    Loop
End Sub

' 检查是否为数字字符
Private Function IsNumericChar(char As String) As Boolean
    IsNumericChar = (char >= "0" And char <= "9")
End Function

' 将对象转换为JSON字符串
Public Function ConvertToJSON(obj As Variant) As String
    On Error GoTo ErrorHandler

    ConvertToJSON = ConvertValueToJSON(obj)
    Exit Function

ErrorHandler:
    ConvertToJSON = "null"
    Debug.Print "JSON转换错误: " & Err.Description
End Function

' 转换值为JSON格式（增强版）
Private Function ConvertValueToJSON(value As Variant) As String
    On Error GoTo ErrorHandler

    If IsNull(value) Then
        ConvertValueToJSON = "null"
    ElseIf IsEmpty(value) Then
        ConvertValueToJSON = "null"
    ElseIf VarType(value) = vbString Then
        ConvertValueToJSON = """" & EscapeJsonString(CStr(value)) & """"
    ElseIf VarType(value) = vbBoolean Then
        If value Then
            ConvertValueToJSON = "true"
        Else
            ConvertValueToJSON = "false"
        End If
    ElseIf IsNumeric(value) Then
        ConvertValueToJSON = Replace(CStr(value), ",", ".")
    ElseIf IsObject(value) Then
        If value Is Nothing Then
            ConvertValueToJSON = "null"
        Else
            ConvertValueToJSON = ConvertObjectToJSON(value)
        End If
    Else
        ConvertValueToJSON = """" & EscapeJsonString(CStr(value)) & """"
    End If

    Exit Function

ErrorHandler:
    ConvertValueToJSON = "null"
    Debug.Print "值转换JSON错误: " & Err.Description
End Function

' 转换对象为JSON
Private Function ConvertObjectToJSON(ByVal obj As Object) As String
    On Error GoTo ErrorHandler

    If obj Is Nothing Then
        ConvertObjectToJSON = "null"
        Exit Function
    End If

    ' 检查是否是Dictionary（对象）
    If TypeName(obj) = "Dictionary" Then
        Dim json As String
        json = "{"

        Dim key As Variant
        Dim first As Boolean
        first = True

        For Each key In obj.Keys
            If Not first Then json = json & ","
            json = json & """" & EscapeJsonString(CStr(key)) & """:"
            json = json & ConvertValueToJSON(obj(key))
            first = False
        Next key

        json = json & "}"
        ConvertObjectToJSON = json
    Else
        ' 其他对象类型，转换为字符串
        ConvertObjectToJSON = """" & EscapeJsonString(CStr(obj)) & """"
    End If

    Exit Function

ErrorHandler:
    ConvertObjectToJSON = "null"
    Debug.Print "对象转换JSON错误: " & Err.Description
End Function

' JSON字符串转义
Private Function EscapeJsonString(str As String) As String
    On Error GoTo ErrorHandler

    Dim result As String
    Dim i As Long

    result = str
    result = Replace(result, "\", "\\")    ' 反斜杠
    result = Replace(result, """", "\""")  ' 引号
    result = Replace(result, "/", "\/")    ' 斜杠
    result = Replace(result, vbCr, "\r")   ' 回车
    result = Replace(result, vbLf, "\n")   ' 换行
    result = Replace(result, vbTab, "\t")  ' 制表符
    result = Replace(result, Chr(8), "\b") ' 退格
    result = Replace(result, Chr(12), "\f") ' 换页

    EscapeJsonString = result
    Exit Function

ErrorHandler:
    EscapeJsonString = str
    Debug.Print "字符串转义错误: " & Err.Description
End Function

' 获取JSON对象的值（支持嵌套路径）
Public Function GetJSONValue(obj As Object, path As String, Optional defaultValue As Variant = "") As Variant
    On Error GoTo ErrorHandler

    If obj Is Nothing Then
        GetJSONValue = defaultValue
        Exit Function
    End If

    Dim keys As Variant
    keys = Split(path, ".")

    Dim currentObj As Object
    Set currentObj = obj

    Dim i As Long
    For i = 0 To UBound(keys)
        If TypeName(currentObj) <> "Dictionary" Then
            GetJSONValue = defaultValue
            Exit Function
        End If

        If currentObj.Exists(keys(i)) Then
            If i = UBound(keys) Then
                ' 最后一个键，返回值
                If IsObject(currentObj(keys(i))) Then
                    Set GetJSONValue = currentObj(keys(i))
                Else
                    GetJSONValue = currentObj(keys(i))
                End If
            Else
                ' 中间键，继续深入
                If IsObject(currentObj(keys(i))) Then
                    Set currentObj = currentObj(keys(i))
                Else
                    GetJSONValue = defaultValue
                    Exit Function
                End If
            End If
        Else
            GetJSONValue = defaultValue
            Exit Function
        End If
    Next i

    Exit Function

ErrorHandler:
    GetJSONValue = defaultValue
    Debug.Print "获取JSON值错误: " & Err.Description
End Function

' 检查JSON对象是否存在指定路径
Public Function HasJSONPath(obj As Object, path As String) As Boolean
    On Error GoTo ErrorHandler

    If obj Is Nothing Then
        HasJSONPath = False
        Exit Function
    End If

    Dim keys As Variant
    keys = Split(path, ".")

    Dim currentObj As Object
    Set currentObj = obj

    Dim i As Long
    For i = 0 To UBound(keys)
        If TypeName(currentObj) <> "Dictionary" Then
            HasJSONPath = False
            Exit Function
        End If

        If currentObj.Exists(keys(i)) Then
            If i < UBound(keys) Then
                If IsObject(currentObj(keys(i))) Then
                    Set currentObj = currentObj(keys(i))
                Else
                    HasJSONPath = False
                    Exit Function
                End If
            End If
        Else
            HasJSONPath = False
            Exit Function
        End If
    Next i

    HasJSONPath = True
    Exit Function

ErrorHandler:
    HasJSONPath = False
    Debug.Print "检查JSON路径错误: " & Err.Description
End Function
