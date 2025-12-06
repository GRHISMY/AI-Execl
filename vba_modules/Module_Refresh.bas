' Attribute VB_Name = "Module_Refresh"
'
' Module_Refresh - 数据刷新模块
' 负责ETF价格数据的批量更新和界面交互
'
Option Explicit

' 全局变量
Private isRefreshing As Boolean
Private isCancelled As Boolean

' 主刷新函数
Public Sub RefreshETFPrices()
    On Error GoTo ErrorHandler
    
    If isRefreshing Then
        MsgBox "数据刷新正在进行中，请稍候...", vbInformation, "ETF数据刷新"
        Exit Sub
    End If
    
    isRefreshing = True
    isCancelled = False
    
    ' 设置状态栏
    Application.StatusBar = "正在刷新ETF价格数据..."
    Application.ScreenUpdating = False
    
    ' 获取当前工作表
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 获取ETF代码列表
    Dim etfCodes As String
    etfCodes = GetETFCodesFromSheet(ws)
    
    If Len(etfCodes) = 0 Then
        MsgBox "请在A列输入ETF代码", vbExclamation, "ETF数据刷新"
        GoTo CleanUp
    End If
    
    ' 验证配置
    If Not Module_Config.ValidateConfig() Then
        MsgBox "请先配置API Token", vbExclamation, "配置错误"
        Module_Config.ShowConfigDialog
        GoTo CleanUp
    End If
    
    ' 调用API获取数据
    Dim apiResponse As String
    apiResponse = Module_API.CallETFAPI(etfCodes)
    
    ' 解析响应
    Dim responseData As Object
    Set responseData = Module_API.ParseAPIResponse(apiResponse)
    
    If responseData Is Nothing Then
        ' 显示详细的调试信息
        Dim debugInfo As String
        debugInfo = Module_API.DebugParseAPIResponse(apiResponse)

        MsgBox "API响应解析失败" & vbCrLf & vbCrLf & _
               "调试信息：" & vbCrLf & debugInfo, vbExclamation, "数据刷新错误"
        GoTo CleanUp
    End If

    ' 检查响应状态
    If responseData.Exists("status") Then
        If responseData("status") = "error" Then
            Dim errorMsg As String
            If responseData.Exists("error") Then
                errorMsg = responseData("error")
            Else
                errorMsg = "未知错误"
            End If
            MsgBox "API调用失败: " & errorMsg, vbExclamation, "数据刷新错误"
            GoTo CleanUp
        End If
    End If

    ' 更新工作表数据
    UpdateSheetWithData ws, responseData

    ' 显示完成信息
    Application.StatusBar = "ETF价格数据刷新完成 - " & Format(Now(), "yyyy-mm-dd hh:mm:ss")

    GoTo CleanUp

ErrorHandler:
    MsgBox "数据刷新过程中发生错误: " & Err.Description, vbExclamation, "刷新错误"
    Debug.Print "刷新错误: " & Err.Description

CleanUp:
    Application.ScreenUpdating = True
    isRefreshing = False
    isCancelled = False
End Sub

' 从工作表获取ETF代码
Private Function GetETFCodesFromSheet(ws As Worksheet) As String
    On Error GoTo ErrorHandler

    Dim codesColumn As String
    Dim startRow As Integer

    codesColumn = Module_Config.GetExcelConfig("etf_codes_column", "A")
    startRow = Module_Config.GetExcelConfig("start_row", 2)

    Dim codes As String
    Dim currentRow As Integer
    currentRow = startRow

    ' 循环读取ETF代码
    Do While Len(Trim(ws.Cells(currentRow, codesColumn).Value)) > 0
        Dim code As String
        code = Trim(ws.Cells(currentRow, codesColumn).Value)

        If Module_API.ValidateETFCodes(code) Then
            If Len(codes) > 0 Then codes = codes & ","
            codes = codes & code
        End If

        currentRow = currentRow + 1

        ' 防止无限循环
        If currentRow > startRow + 1000 Then Exit Do
    Loop

    GetETFCodesFromSheet = codes
    Exit Function

ErrorHandler:
    GetETFCodesFromSheet = ""
    Debug.Print "获取ETF代码错误: " & Err.Description
End Function

' 更新工作表数据
Private Sub UpdateSheetWithData(ws As Worksheet, responseData As Object)
    On Error GoTo ErrorHandler

    If Not responseData.Exists("data") Then
        Debug.Print "响应中没有数据字段"
        Exit Sub
    End If

    Dim dataObj As Object
    Set dataObj = responseData("data")

    Dim codesColumn As String
    Dim pricesColumn As String
    Dim statusColumn As String
    Dim timeColumn As String
    Dim startRow As Integer

    codesColumn = Module_Config.GetExcelConfig("etf_codes_column", "A")
    pricesColumn = Module_Config.GetExcelConfig("prices_column", "B")
    statusColumn = Module_Config.GetExcelConfig("status_column", "C")
    timeColumn = Module_Config.GetExcelConfig("update_time_column", "D")
    startRow = Module_Config.GetExcelConfig("start_row", 2)

    Dim currentRow As Integer
    currentRow = startRow

    ' 遍历工作表中的ETF代码
    Do While Len(Trim(ws.Cells(currentRow, codesColumn).Value)) > 0
        Dim etfCode As String
        etfCode = Trim(ws.Cells(currentRow, codesColumn).Value)

        ' 查找对应的价格数据
        If dataObj.Exists(etfCode) Then
            Dim etfData As Object
            Set etfData = dataObj(etfCode)

            ' 更新价格
            If etfData.Exists("price") And Not IsNull(etfData("price")) Then
                ws.Cells(currentRow, pricesColumn).Value = etfData("price")
                ws.Cells(currentRow, pricesColumn).NumberFormat = "0.000"
            Else
                ws.Cells(currentRow, pricesColumn).Value = "N/A"
            End If

            ' 更新状态
            If etfData.Exists("status") Then
                ws.Cells(currentRow, statusColumn).Value = etfData("status")

                ' 根据状态设置颜色
                If etfData("status") = "success" Then
                    ws.Cells(currentRow, statusColumn).Font.Color = RGB(0, 128, 0) ' 绿色
                Else
                    ws.Cells(currentRow, statusColumn).Font.Color = RGB(255, 0, 0) ' 红色
                End If
            End If

            ' 更新时间
            If etfData.Exists("update_time") Then
                ws.Cells(currentRow, timeColumn).Value = FormatUpdateTime(etfData("update_time"))
            End If
        Else
            ' 没有找到数据
            ws.Cells(currentRow, pricesColumn).Value = "未找到"
            ws.Cells(currentRow, statusColumn).Value = "no_data"
            ws.Cells(currentRow, statusColumn).Font.Color = RGB(255, 0, 0)
        End If

        currentRow = currentRow + 1

        ' 防止无限循环
        If currentRow > startRow + 1000 Then Exit Do
    Loop

    Exit Sub

ErrorHandler:
    Debug.Print "更新工作表数据错误: " & Err.Description
End Sub

' 格式化更新时间
Private Function FormatUpdateTime(timeStr As String) As String
    On Error GoTo ErrorHandler

    ' 简单的时间格式化
    If InStr(timeStr, "T") > 0 Then
        ' ISO格式时间
        Dim parts As Variant
        parts = Split(timeStr, "T")
        If UBound(parts) >= 1 Then
            Dim timePart As String
            timePart = parts(1)
            If InStr(timePart, ".") > 0 Then
                timePart = Left(timePart, InStr(timePart, ".") - 1)
            End If
            FormatUpdateTime = parts(0) & " " & timePart
        Else
            FormatUpdateTime = timeStr
        End If
    Else
        FormatUpdateTime = timeStr
    End If

    Exit Function

ErrorHandler:
    FormatUpdateTime = timeStr
End Function

' 清除所有数据
Public Sub ClearAllData()
    On Error GoTo ErrorHandler

    If MsgBox("确定要清除所有ETF数据吗？", vbYesNo + vbQuestion, "清除数据") = vbNo Then
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim pricesColumn As String
    Dim statusColumn As String
    Dim timeColumn As String
    Dim startRow As Integer

    pricesColumn = Module_Config.GetExcelConfig("prices_column", "B")
    statusColumn = Module_Config.GetExcelConfig("status_column", "C")
    timeColumn = Module_Config.GetExcelConfig("update_time_column", "D")
    startRow = Module_Config.GetExcelConfig("start_row", 2)

    ' 清除数据列
    ws.Columns(pricesColumn & ":" & timeColumn).ClearContents

    ' 设置表头
    ws.Cells(1, "A").Value = "ETF代码"
    ws.Cells(1, "B").Value = "收盘价"
    ws.Cells(1, "C").Value = "状态"
    ws.Cells(1, "D").Value = "更新时间"

    ' 格式化表头
    With ws.Range("A1:D1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With

    Application.StatusBar = "数据已清除"

    Exit Sub

ErrorHandler:
    MsgBox "清除数据时发生错误: " & Err.Description, vbExclamation, "清除错误"
End Sub

' 取消刷新
Public Sub CancelRefresh()
    On Error Resume Next

    isCancelled = True
    Application.StatusBar = "正在取消刷新..."
End Sub

' 批量添加示例ETF代码
Public Sub AddSampleETFCodes()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim codesColumn As String
    Dim startRow As Integer

    codesColumn = Module_Config.GetExcelConfig("etf_codes_column", "A")
    startRow = Module_Config.GetExcelConfig("start_row", 2)

    ' 示例ETF代码
    Dim sampleCodes As Variant
    sampleCodes = Array("159915", "159919", "159928", "510300", "510500", "515000")

    Dim i As Integer
    For i = 0 To UBound(sampleCodes)
        ws.Cells(startRow + i, codesColumn).Value = sampleCodes(i)
    Next i

    MsgBox "已添加示例ETF代码", vbInformation, "添加完成"

    Exit Sub

ErrorHandler:
    MsgBox "添加示例代码时发生错误: " & Err.Description, vbExclamation, "添加错误"
End Sub

' 刷新单个ETF
Public Sub RefreshSingleETF(etfCode As String)
    On Error GoTo ErrorHandler

    If Len(Trim(etfCode)) = 0 Then
        Exit Sub
    End If

    ' 验证代码
    If Not Module_API.ValidateETFCodes(etfCode) Then
        MsgBox "无效的ETF代码: " & etfCode, vbExclamation, "代码错误"
        Exit Sub
    End If

    ' 调用API
    Application.StatusBar = "正在获取 " & etfCode & " 的价格数据..."

    Dim apiResponse As String
    apiResponse = Module_API.CallETFAPI(etfCode)

    ' 解析并显示结果
    Dim responseData As Object
    Set responseData = Module_API.ParseAPIResponse(apiResponse)

    If Not responseData Is Nothing Then
        If responseData.Exists("data") Then
            Dim dataObj As Object
            Set dataObj = responseData("data")

            If dataObj.Exists(etfCode) Then
                Dim etfData As Object
                Set etfData = dataObj(etfCode)

                Dim msg As String
                msg = "ETF代码: " & etfCode & vbCrLf

                If etfData.Exists("price") Then
                    msg = msg & "收盘价: " & etfData("price") & vbCrLf
                End If

                If etfData.Exists("status") Then
                    msg = msg & "状态: " & etfData("status") & vbCrLf
                End If

                MsgBox msg, vbInformation, "ETF价格信息"
            End If
        End If
    End If

    Application.StatusBar = ""

    Exit Sub

ErrorHandler:
    Application.StatusBar = ""
    MsgBox "获取ETF价格时发生错误: " & Err.Description, vbExclamation, "获取错误"
End Sub