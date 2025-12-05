Attribute VB_Name = "JsonConverter"
' =========================================
' 模块名称: JsonConverter
' 描述: 简化的JSON解析器（针对理想财经API）
' 作者: Aone Copilot
' 创建日期: 2025-12-05
' =========================================

Option Explicit

' ========== 公共函数 ==========

' 解析JSON字符串为字典对象
Public Function ParseJson(jsonString As String) As Object
    Dim json As Object
    Set json = CreateObject("Scripting.Dictionary")
    
    ' 移除首尾空白
    jsonString = Trim(jsonString)
    
    ' 检查是否为JSON对象
    If Left(jsonString, 1) = "{" And Right(jsonString, 1) = "}" Then
        Set ParseJson = ParseObject(jsonString)
    ElseIf Left(jsonString, 1) = "[" And Right(jsonString, 1) = "]" Then
        Set ParseJson = ParseArray(jsonString)
    Else
        Set ParseJson = Nothing
    End If
End Function

' 从JSON对象中获取值
Public Function GetJsonValue(jsonObj As Object, key As String) As Variant
    On Error Resume Next
    
    If jsonObj Is Nothing Then
        GetJsonValue = Null
        Exit Function
    End If
    
    If jsonObj.Exists(key) Then
        If IsObject(jsonObj(key)) Then
            Set GetJsonValue = jsonObj(key)
        Else
            GetJsonValue = jsonObj(key)
        End If
    Else
        GetJsonValue = Null
    End If
End Function

' 从JSON数组中获取元素
Public Function GetJsonArrayItem(jsonArray As Object, index As Long) As Variant
    On Error Resume Next
    
    If jsonArray Is Nothing Then
        GetJsonArrayItem = Null
        Exit Function
    End If
    
    If index >= 0 And index < jsonArray.Count Then
        If IsObject(jsonArray(index)) Then
            Set GetJsonArrayItem = jsonArray(index)
        Else
            GetJsonArrayItem = jsonArray(index)
        End If
    Else
        GetJsonArrayItem = Null
    End If
End Function

' 获取JSON数组长度
Public Function GetJsonArrayLength(jsonArray As Object) As Long
    If jsonArray Is Nothing Then
        GetJsonArrayLength = 0
    Else
        GetJsonArrayLength = jsonArray.Count
    End If
End Function

' ========== 内部函数 ==========

' 解析JSON对象
Private Function ParseObject(jsonString As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 移除首尾的大括号
    jsonString = Mid(jsonString, 2, Len(jsonString) - 2)
    jsonString = Trim(jsonString)
    
    If Len(jsonString) = 0 Then
        Set ParseObject = dict
        Exit Function
    End If
    
    ' 解析键值对
    Dim pairs() As String
    pairs = SplitJsonPairs(jsonString)
    
    Dim i As Long
    Dim pair As String
    Dim colonPos As Long
    Dim key As String
    Dim value As String
    
    For i = LBound(pairs) To UBound(pairs)
        pair = Trim(pairs(i))
        If Len(pair) > 0 Then
            colonPos = InStr(pair, ":")
            If colonPos > 0 Then
                key = Trim(Mid(pair, 1, colonPos - 1))
                value = Trim(Mid(pair, colonPos + 1))
                
                ' 移除键的引号
                key = RemoveQuotes(key)
                
                ' 解析值
                dict(key) = ParseValue(value)
            End If
        End If
    Next i
    
    Set ParseObject = dict
End Function

' 解析JSON数组
Private Function ParseArray(jsonString As String) As Object
    Dim arr As Object
    Set arr = CreateObject("Scripting.Dictionary")
    
    ' 移除首尾的方括号
    jsonString = Mid(jsonString, 2, Len(jsonString) - 2)
    jsonString = Trim(jsonString)
    
    If Len(jsonString) = 0 Then
        Set ParseArray = arr
        Exit Function
    End If
    
    ' 分割数组元素
    Dim elements() As String
    elements = SplitJsonArray(jsonString)
    
    Dim i As Long
    For i = LBound(elements) To UBound(elements)
        arr(i) = ParseValue(Trim(elements(i)))
    Next i
    
    Set ParseArray = arr
End Function

' 解析JSON值
Private Function ParseValue(value As String) As Variant
    value = Trim(value)
    
    ' 检查是否为对象
    If Left(value, 1) = "{" And Right(value, 1) = "}" Then
        Set ParseValue = ParseObject(value)
        Exit Function
    End If
    
    ' 检查是否为数组
    If Left(value, 1) = "[" And Right(value, 1) = "]" Then
        Set ParseValue = ParseArray(value)
        Exit Function
    End If
    
    ' 检查是否为字符串
    If Left(value, 1) = """" And Right(value, 1) = """" Then
        ParseValue = RemoveQuotes(value)
        Exit Function
    End If
    
    ' 检查是否为null
    If LCase(value) = "null" Then
        ParseValue = Null
        Exit Function
    End If
    
    ' 检查是否为布尔值
    If LCase(value) = "true" Then
        ParseValue = True
        Exit Function
    ElseIf LCase(value) = "false" Then
        ParseValue = False
        Exit Function
    End If
    
    ' 尝试转换为数字
    If IsNumeric(value) Then
        If InStr(value, ".") > 0 Then
            ParseValue = CDbl(value)
        Else
            ParseValue = CLng(value)
        End If
    Else
        ParseValue = value
    End If
End Function

' 分割JSON键值对
Private Function SplitJsonPairs(jsonString As String) As String()
    Dim pairs() As String
    Dim pairCount As Long
    Dim i As Long
    Dim char As String
    Dim inString As Boolean
    Dim braceLevel As Long
    Dim bracketLevel As Long
    Dim currentPair As String
    
    pairCount = 0
    inString = False
    braceLevel = 0
    bracketLevel = 0
    currentPair = ""
    
    ReDim pairs(0 To 100)
    
    For i = 1 To Len(jsonString)
        char = Mid(jsonString, i, 1)
        
        ' 处理字符串
        If char = """" And (i = 1 Or Mid(jsonString, i - 1, 1) <> "\") Then
            inString = Not inString
        End If
        
        If Not inString Then
            ' 处理嵌套结构
            If char = "{" Then
                braceLevel = braceLevel + 1
            ElseIf char = "}" Then
                braceLevel = braceLevel - 1
            ElseIf char = "[" Then
                bracketLevel = bracketLevel + 1
            ElseIf char = "]" Then
                bracketLevel = bracketLevel - 1
            ElseIf char = "," And braceLevel = 0 And bracketLevel = 0 Then
                ' 分隔符
                If Len(currentPair) > 0 Then
                    pairs(pairCount) = currentPair
                    pairCount = pairCount + 1
                    currentPair = ""
                End If
                GoTo NextChar
            End If
        End If
        
        currentPair = currentPair & char
        
NextChar:
    Next i
    
    ' 添加最后一个键值对
    If Len(currentPair) > 0 Then
        pairs(pairCount) = currentPair
        pairCount = pairCount + 1
    End If
    
    ' 调整数组大小
    If pairCount > 0 Then
        ReDim Preserve pairs(0 To pairCount - 1)
    Else
        ReDim pairs(0 To 0)
        pairs(0) = ""
    End If
    
    SplitJsonPairs = pairs
End Function

' 分割JSON数组元素
Private Function SplitJsonArray(jsonString As String) As String()
    Dim elements() As String
    Dim elemCount As Long
    Dim i As Long
    Dim char As String
    Dim inString As Boolean
    Dim braceLevel As Long
    Dim bracketLevel As Long
    Dim currentElem As String
    
    elemCount = 0
    inString = False
    braceLevel = 0
    bracketLevel = 0
    currentElem = ""
    
    ReDim elements(0 To 100)
    
    For i = 1 To Len(jsonString)
        char = Mid(jsonString, i, 1)
        
        ' 处理字符串
        If char = """" And (i = 1 Or Mid(jsonString, i - 1, 1) <> "\") Then
            inString = Not inString
        End If
        
        If Not inString Then
            ' 处理嵌套结构
            If char = "{" Then
                braceLevel = braceLevel + 1
            ElseIf char = "}" Then
                braceLevel = braceLevel - 1
            ElseIf char = "[" Then
                bracketLevel = bracketLevel + 1
            ElseIf char = "]" Then
                bracketLevel = bracketLevel - 1
            ElseIf char = "," And braceLevel = 0 And bracketLevel = 0 Then
                ' 分隔符
                If Len(currentElem) > 0 Then
                    elements(elemCount) = currentElem
                    elemCount = elemCount + 1
                    currentElem = ""
                End If
                GoTo NextChar
            End If
        End If
        
        currentElem = currentElem & char
        
NextChar:
    Next i
    
    ' 添加最后一个元素
    If Len(currentElem) > 0 Then
        elements(elemCount) = currentElem
        elemCount = elemCount + 1
    End If
    
    ' 调整数组大小
    If elemCount > 0 Then
        ReDim Preserve elements(0 To elemCount - 1)
    Else
        ReDim elements(0 To 0)
        elements(0) = ""
    End If
    
    SplitJsonArray = elements
End Function

' 移除字符串首尾的引号
Private Function RemoveQuotes(str As String) As String
    str = Trim(str)
    If Left(str, 1) = """" Then str = Mid(str, 2)
    If Right(str, 1) = """" Then str = Left(str, Len(str) - 1)
    RemoveQuotes = str
End Function

' ========== 便捷函数 ==========

' 从JSON响应中提取最新收盘价和日期
Public Function ExtractLatestClosePrice(jsonResponse As String, ByRef outDate As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim jsonObj As Object
    Dim dataArray As Object
    Dim lastItem As Object
    Dim arrayLength As Long
    Dim closePrice As Variant
    
    ' 解析JSON
    Set jsonObj = ParseJson(jsonResponse)
    If jsonObj Is Nothing Then GoTo ErrorHandler
    
    ' 检查code字段
    Dim code As Variant
    code = GetJsonValue(jsonObj, "code")
    If code <> 1 Then
        ExtractLatestClosePrice = "API返回错误"
        Exit Function
    End If
    
    ' 获取data数组
    Set dataArray = GetJsonValue(jsonObj, "data")
    If dataArray Is Nothing Then GoTo ErrorHandler
    
    arrayLength = GetJsonArrayLength(dataArray)
    If arrayLength = 0 Then
        ExtractLatestClosePrice = "无数据"
        Exit Function
    End If
    
    ' 获取最后一条记录（最新数据）
    Set lastItem = GetJsonArrayItem(dataArray, arrayLength - 1)
    If lastItem Is Nothing Then GoTo ErrorHandler
    
    ' 提取收盘价和日期
    closePrice = GetJsonValue(lastItem, "close")
    outDate = GetJsonValue(lastItem, "date")
    
    If IsNull(closePrice) Or closePrice = "" Then
        ExtractLatestClosePrice = "无收盘价数据"
    Else
        ExtractLatestClosePrice = closePrice
    End If
    
    Exit Function
    
ErrorHandler:
    ExtractLatestClosePrice = "解析错误"
    outDate = ""
End Function
