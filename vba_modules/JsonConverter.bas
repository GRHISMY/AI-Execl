Attribute VB_Name = "JsonConverter"
Option Explicit

' ========== JSON解析模块 ==========
' 专门用于解析理想财经API返回的JSON数据

' ========== 主要解析函数 ==========
Public Function ParseApiResponse(jsonString As String) As Dictionary
    ' 解析API响应，返回包含code和data的字典
    Set ParseApiResponse = New Dictionary
    
    On Error GoTo ErrorHandler
    
    ' 清理JSON字符串
    jsonString = Trim(jsonString)
    If Left(jsonString, 1) <> "{" Or Right(jsonString, 1) <> "}" Then
        ParseApiResponse("error") = "Invalid JSON format"
        Exit Function
    End If
    
    ' 提取code值
    Dim code As Integer
    code = ExtractJsonInteger(jsonString, "code")
    ParseApiResponse("code") = code
    
    ' 如果code不等于1，表示API返回错误
    If code <> 1 Then
        ParseApiResponse("error") = "API returned error code: " & code
        Exit Function
    End If
    
    ' 提取data数组
    Dim dataArray As Collection
    Set dataArray = ExtractDataArray(jsonString)
    Set ParseApiResponse("data") = dataArray
    
    Exit Function
    
ErrorHandler:
    ParseApiResponse("error") = "JSON parsing error: " & Err.Description
End Function

' ========== 提取最新收盘价 ==========
Public Function GetLatestClosePriceFromJson(jsonString As String) As Variant
    ' 从API响应中提取最新收盘价
    On Error GoTo ErrorHandler
    
    Dim response As Dictionary
    Set response = ParseApiResponse(jsonString)
    
    ' 检查是否有错误
    If response.Exists("error") Then
        GetLatestClosePriceFromJson = response("error")
        Exit Function
    End If
    
    ' 获取数据数组
    Dim dataArray As Collection
    Set dataArray = response("data")
    
    If dataArray.Count = 0 Then
        GetLatestClosePriceFromJson = "无数据"
        Exit Function
    End If
    
    ' 获取最新数据（第一条记录通常是最新的）
    Dim latestRecord As Dictionary
    Set latestRecord = dataArray(1)
    
    ' 提取收盘价
    If latestRecord.Exists("close") Then
        GetLatestClosePriceFromJson = latestRecord("close")
    Else
        GetLatestClosePriceFromJson = "无收盘价数据"
    End If
    
    Exit Function
    
ErrorHandler:
    GetLatestClosePriceFromJson = "解析错误: " & Err.Description
End Function

' ========== 提取最新数据日期 ==========
Public Function GetLatestDateFromJson(jsonString As String) As Variant
    ' 从API响应中提取最新数据日期
    On Error GoTo ErrorHandler
    
    Dim response As Dictionary
    Set response = ParseApiResponse(jsonString)
    
    ' 检查是否有错误
    If response.Exists("error") Then
        GetLatestDateFromJson = response("error")
        Exit Function
    End If
    
    ' 获取数据数组
    Dim dataArray As Collection
    Set dataArray = response("data")
    
    If dataArray.Count = 0 Then
        GetLatestDateFromJson = "无数据"
        Exit Function
    End If
    
    ' 获取最新数据
    Dim latestRecord As Dictionary
    Set latestRecord = dataArray(1)
    
    ' 提取日期
    If latestRecord.Exists("date") Then
        GetLatestDateFromJson = latestRecord("date")
    Else
        GetLatestDateFromJson = Format(Date, "yyyy-mm-dd")
    End If
    
    Exit Function
    
ErrorHandler:
    GetLatestDateFromJson = "解析错误: " & Err.Description
End Function

' ========== 辅助函数：提取JSON整数值 ==========
Private Function ExtractJsonInteger(jsonString As String, key As String) As Integer
    Dim keyPattern As String
    Dim startPos As Integer
    Dim endPos As Integer
    Dim valueString As String
    
    keyPattern = """" & key & """"
    startPos = InStr(jsonString, keyPattern)
    
    If startPos = 0 Then
        ExtractJsonInteger = 0
        Exit Function
    End If
    
    ' 找到冒号后的值
    startPos = InStr(startPos, jsonString, ":") + 1
    
    ' 跳过空格
    Do While Mid(jsonString, startPos, 1) = " "
        startPos = startPos + 1
    Loop
    
    ' 找到值的结束位置
    endPos = startPos
    Do While endPos <= Len(jsonString) And IsNumeric(Mid(jsonString, endPos, 1))
        endPos = endPos + 1
    Loop
    endPos = endPos - 1
    
    valueString = Mid(jsonString, startPos, endPos - startPos + 1)
    ExtractJsonInteger = Val(valueString)
End Function

' ========== 辅助函数：提取数据数组 ==========
Private Function ExtractDataArray(jsonString As String) As Collection
    Set ExtractDataArray = New Collection
    
    On Error GoTo ErrorHandler
    
    ' 找到data数组的开始位置
    Dim dataStart As Integer
    Dim dataEnd As Integer
    Dim arrayString As String
    
    dataStart = InStr(jsonString, """data"":")
    If dataStart = 0 Then Exit Function
    
    ' 找到数组开始的[
    dataStart = InStr(dataStart, jsonString, "[")
    If dataStart = 0 Then Exit Function
    
    ' 找到对应的]
    dataEnd = FindMatchingBracket(jsonString, dataStart)
    If dataEnd = 0 Then Exit Function
    
    arrayString = Mid(jsonString, dataStart + 1, dataEnd - dataStart - 1)
    
    ' 解析数组中的每个对象
    ParseArrayObjects arrayString, ExtractDataArray
    
    Exit Function
    
ErrorHandler:
    ' 返回空集合
End Function

' ========== 辅助函数：解析数组中的对象 ==========
Private Sub ParseArrayObjects(arrayString As String, resultCollection As Collection)
    On Error GoTo ErrorHandler
    
    Dim objectStart As Integer
    Dim objectEnd As Integer
    Dim currentPos As Integer
    Dim objectString As String
    Dim recordDict As Dictionary
    
    currentPos = 1
    
    Do While currentPos <= Len(arrayString)
        ' 找到下一个对象的开始
        objectStart = InStr(currentPos, arrayString, "{")
        If objectStart = 0 Then Exit Do
        
        ' 找到对象的结束
        objectEnd = FindMatchingBrace(arrayString, objectStart)
        If objectEnd = 0 Then Exit Do
        
        ' 提取对象字符串
        objectString = Mid(arrayString, objectStart + 1, objectEnd - objectStart - 1)
        
        ' 解析对象
        Set recordDict = ParseJsonObject(objectString)
        resultCollection.Add recordDict
        
        currentPos = objectEnd + 1
    Loop
    
    Exit Sub
    
ErrorHandler:
    ' 静默处理错误
End Sub

' ========== 辅助函数：解析JSON对象 ==========
Private Function ParseJsonObject(objectString As String) As Dictionary
    Set ParseJsonObject = New Dictionary
    
    On Error GoTo ErrorHandler
    
    ' 解析常见字段
    ParseJsonObject("date") = ExtractJsonStringValue(objectString, "date")
    ParseJsonObject("open") = ExtractJsonNumberValue(objectString, "open")
    ParseJsonObject("high") = ExtractJsonNumberValue(objectString, "high")
    ParseJsonObject("low") = ExtractJsonNumberValue(objectString, "low")
    ParseJsonObject("close") = ExtractJsonNumberValue(objectString, "close")
    ParseJsonObject("volume") = ExtractJsonNumberValue(objectString, "volume")
    ParseJsonObject("amount") = ExtractJsonNumberValue(objectString, "amount")
    ParseJsonObject("change") = ExtractJsonNumberValue(objectString, "change")
    
    Exit Function
    
ErrorHandler:
    ' 返回空字典
End Function

' ========== 辅助函数：提取JSON字符串值 ==========
Private Function ExtractJsonStringValue(objectString As String, key As String) As String
    Dim keyPattern As String
    Dim startPos As Integer
    Dim endPos As Integer
    
    keyPattern = """" & key & """"
    startPos = InStr(objectString, keyPattern)
    
    If startPos = 0 Then
        ExtractJsonStringValue = ""
        Exit Function
    End If
    
    ' 找到冒号后的引号
    startPos = InStr(startPos, objectString, ":") + 1
    startPos = InStr(startPos, objectString, """") + 1
    
    ' 找到结束引号
    endPos = InStr(startPos, objectString, """")
    If endPos = 0 Then
        ExtractJsonStringValue = ""
        Exit Function
    End If
    
    ExtractJsonStringValue = Mid(objectString, startPos, endPos - startPos)
End Function

' ========== 辅助函数：提取JSON数值 ==========
Private Function ExtractJsonNumberValue(objectString As String, key As String) As Double
    Dim keyPattern As String
    Dim startPos As Integer
    Dim endPos As Integer
    Dim valueString As String
    Dim char As String
    
    keyPattern = """" & key & """"
    startPos = InStr(objectString, keyPattern)
    
    If startPos = 0 Then
        ExtractJsonNumberValue = 0
        Exit Function
    End If
    
    ' 找到冒号后的值
    startPos = InStr(startPos, objectString, ":") + 1
    
    ' 跳过空格
    Do While startPos <= Len(objectString) And Mid(objectString, startPos, 1) = " "
        startPos = startPos + 1
    Loop
    
    ' 找到数值的结束位置
    endPos = startPos
    Do While endPos <= Len(objectString)
        char = Mid(objectString, endPos, 1)
        If Not (IsNumeric(char) Or char = "." Or char = "-" Or char = "e" Or char = "E" Or char = "+") Then
            Exit Do
        End If
        endPos = endPos + 1
    Loop
    endPos = endPos - 1
    
    If endPos >= startPos Then
        valueString = Mid(objectString, startPos, endPos - startPos + 1)
        ExtractJsonNumberValue = Val(valueString)
    Else
        ExtractJsonNumberValue = 0
    End If
End Function

' ========== 辅助函数：找到匹配的方括号 ==========
Private Function FindMatchingBracket(text As String, startPos As Integer) As Integer
    Dim count As Integer
    Dim i As Integer
    Dim char As String
    
    count = 1
    For i = startPos + 1 To Len(text)
        char = Mid(text, i, 1)
        If char = "[" Then
            count = count + 1
        ElseIf char = "]" Then
            count = count - 1
            If count = 0 Then
                FindMatchingBracket = i
                Exit Function
            End If
        End If
    Next i
    
    FindMatchingBracket = 0
End Function

' ========== 辅助函数：找到匹配的花括号 ==========
Private Function FindMatchingBrace(text As String, startPos As Integer) As Integer
    Dim count As Integer
    Dim i As Integer
    Dim char As String
    
    count = 1
    For i = startPos + 1 To Len(text)
        char = Mid(text, i, 1)
        If char = "{" Then
            count = count + 1
        ElseIf char = "}" Then
            count = count - 1
            If count = 0 Then
                FindMatchingBrace = i
                Exit Function
            End If
        End If
    Next i
    
    FindMatchingBrace = 0
End Function
