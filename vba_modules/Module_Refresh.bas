Attribute VB_Name = "Module_Refresh"
Option Explicit

' ========== 数据刷新模块 ==========
' 负责批量刷新ETF价格数据

' ========== 主要刷新函数 ==========
Public Sub RefreshAllPrices()
    ' 批量刷新所有ETF价格
    ' 遍历A列的ETF代码，获取最新收盘价并更新B列和C列
    
    On Error GoTo ErrorHandler
    
    ' 验证配置
    If Not ValidateConfiguration() Then
        MsgBox "配置验证失败，无法刷新数据", vbCritical, "配置错误"
        Exit Sub
    End If
    
    ' 获取或创建工作表
    Dim ws As Worksheet
    Set ws = GetOrCreateWorksheet()
    
    ' 关闭屏幕更新以提高性能
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Dim startTime As Double
    startTime = Timer
    
    ' 显示开始消息
    Application.StatusBar = "开始刷新ETF价格数据..."
    
    ' 获取数据范围
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    
    If lastRow <= HEADER_ROW Then
        Application.StatusBar = False
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        MsgBox "没有找到ETF代码数据", vbInformation, "数据刷新"
        Exit Sub
    End If
    
    ' 统计变量
    Dim totalCount As Long
    Dim successCount As Long
    Dim errorCount As Long
    Dim currentRow As Long
    Dim etfCode As String
    Dim closePrice As Variant
    Dim dataDate As Variant
    
    totalCount = lastRow - HEADER_ROW
    
    ' 遍历每一行数据
    For currentRow = HEADER_ROW + 1 To lastRow
        ' 获取ETF代码
        etfCode = Trim(CStr(ws.Cells(currentRow, ETF_CODE_COLUMN).Value))
        
        ' 跳过空单元格
        If Len(etfCode) = 0 Then
            GoTo NextRow
        End If
        
        ' 显示当前进度
        Dim progressMessage As String
        progressMessage = "正在刷新ETF价格数据... (" & (currentRow - HEADER_ROW) & "/" & totalCount & ") " & etfCode
        Application.StatusBar = progressMessage
        
        ' 获取收盘价和数据日期
        closePrice = GetLatestClosePrice(etfCode)
        dataDate = GetLatestDataDate(etfCode)
        
        ' 更新单元格
        If IsNumeric(closePrice) Then
            ' 成功获取数据
            ws.Cells(currentRow, PRICE_COLUMN).Value = CDbl(closePrice)
            ws.Cells(currentRow, PRICE_COLUMN).NumberFormat = "0.000"
            
            ' 设置价格单元格颜色（绿色表示成功）
            ws.Cells(currentRow, PRICE_COLUMN).Interior.Color = RGB(144, 238, 144)
            
            successCount = successCount + 1
        Else
            ' 获取数据失败
            ws.Cells(currentRow, PRICE_COLUMN).Value = closePrice ' 显示错误信息
            
            ' 设置价格单元格颜色（红色表示错误）
            ws.Cells(currentRow, PRICE_COLUMN).Interior.Color = RGB(255, 182, 193)
            
            errorCount = errorCount + 1
        End If
        
        ' 更新数据日期
        If IsDate(dataDate) Or Len(CStr(dataDate)) > 0 Then
            ws.Cells(currentRow, DATE_COLUMN).Value = dataDate
        Else
            ws.Cells(currentRow, DATE_COLUMN).Value = Format(Date, "yyyy-mm-dd")
        End If
        
        ' 刷新屏幕显示
        DoEvents
        
NextRow:
    Next currentRow
    
    ' 自动调整列宽
    ws.Columns(ETF_CODE_COLUMN).AutoFit
    ws.Columns(PRICE_COLUMN).AutoFit
    ws.Columns(DATE_COLUMN).AutoFit
    
    ' 计算耗时
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime
    
    ' 恢复Excel设置
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ' 显示完成消息
    Dim resultMessage As String
    resultMessage = "数据刷新完成！" & vbCrLf & _
                   "总数: " & totalCount & vbCrLf & _
                   "成功: " & successCount & vbCrLf & _
                   "失败: " & errorCount & vbCrLf & _
                   "耗时: " & Format(elapsedTime, "0.0") & "秒"
    
    MsgBox resultMessage, vbInformation, "刷新完成"
    
    Exit Sub
    
ErrorHandler:
    ' 恢复Excel设置
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "刷新数据时发生错误：" & Err.Description, vbCritical, "刷新错误"
End Sub

' ========== 单个ETF刷新 ==========
Public Sub RefreshSingleETF(targetRow As Long)
    ' 刷新指定行的ETF数据
    ' 参数: targetRow - 目标行号
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = GetOrCreateWorksheet()
    
    ' 验证行号
    If targetRow <= HEADER_ROW Then
        MsgBox "无效的行号", vbExclamation, "参数错误"
        Exit Sub
    End If
    
    ' 获取ETF代码
    Dim etfCode As String
    etfCode = Trim(CStr(ws.Cells(targetRow, ETF_CODE_COLUMN).Value))
    
    If Len(etfCode) = 0 Then
        MsgBox "该行没有ETF代码", vbExclamation, "数据错误"
        Exit Sub
    End If
    
    Application.StatusBar = "正在刷新 " & etfCode & " 的价格数据..."
    
    ' 获取数据
    Dim closePrice As Variant
    Dim dataDate As Variant
    
    closePrice = GetLatestClosePrice(etfCode)
    dataDate = GetLatestDataDate(etfCode)
    
    ' 更新单元格
    If IsNumeric(closePrice) Then
        ws.Cells(targetRow, PRICE_COLUMN).Value = CDbl(closePrice)
        ws.Cells(targetRow, PRICE_COLUMN).NumberFormat = "0.000"
        ws.Cells(targetRow, PRICE_COLUMN).Interior.Color = RGB(144, 238, 144)
        
        MsgBox "刷新成功！" & vbCrLf & "ETF代码: " & etfCode & vbCrLf & "收盘价: " & closePrice, vbInformation, "刷新完成"
    Else
        ws.Cells(targetRow, PRICE_COLUMN).Value = closePrice
        ws.Cells(targetRow, PRICE_COLUMN).Interior.Color = RGB(255, 182, 193)
        
        MsgBox "刷新失败！" & vbCrLf & "ETF代码: " & etfCode & vbCrLf & "错误信息: " & closePrice, vbExclamation, "刷新失败"
    End If
    
    ' 更新日期
    ws.Cells(targetRow, DATE_COLUMN).Value = dataDate
    
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    MsgBox "刷新单个ETF时发生错误：" & Err.Description, vbCritical, "刷新错误"
End Sub

' ========== 清除所有价格数据 ==========
Public Sub ClearAllPrices()
    ' 清除所有价格数据（保留ETF代码）
    
    Dim result As VbMsgBoxResult
    result = MsgBox("确定要清除所有价格数据吗？", vbYesNo + vbQuestion, "确认清除")
    
    If result = vbNo Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = GetOrCreateWorksheet()
    
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    
    If lastRow > HEADER_ROW Then
        ' 清除价格列和日期列的数据
        ws.Range(ws.Cells(HEADER_ROW + 1, PRICE_COLUMN), ws.Cells(lastRow, DATE_COLUMN)).Clear
        
        MsgBox "价格数据已清除", vbInformation, "清除完成"
    Else
        MsgBox "没有数据需要清除", vbInformation, "清除数据"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "清除数据时发生错误：" & Err.Description, vbCritical, "清除错误"
End Sub

' ========== 添加新ETF代码 ==========
Public Sub AddNewETFCode()
    ' 添加新的ETF代码
    
    Dim newCode As String
    newCode = InputBox("请输入ETF代码（6位数字）：", "添加ETF代码")
    
    If Len(newCode) = 0 Then Exit Sub
    
    ' 验证ETF代码格式
    If Len(newCode) <> 6 Or Not IsNumeric(newCode) Then
        MsgBox "ETF代码格式错误，请输入6位数字", vbExclamation, "格式错误"
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = GetOrCreateWorksheet()
    
    ' 检查是否已存在
    If CheckETFCodeExists(ws, newCode) Then
        MsgBox "ETF代码 " & newCode & " 已存在", vbExclamation, "代码重复"
        Exit Sub
    End If
    
    ' 添加到下一个空行
    Dim nextRow As Long
    nextRow = GetLastDataRow(ws) + 1
    
    ws.Cells(nextRow, ETF_CODE_COLUMN).Value = newCode
    
    ' 询问是否立即刷新
    Dim result As VbMsgBoxResult
    result = MsgBox("ETF代码已添加，是否立即刷新价格？", vbYesNo + vbQuestion, "刷新价格")
    
    If result = vbYes Then
        RefreshSingleETF nextRow
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "添加ETF代码时发生错误：" & Err.Description, vbCritical, "添加错误"
End Sub

' ========== 创建刷新按钮 ==========
Public Sub CreateRefreshButton()
    ' 在工作表中创建刷新按钮
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = GetOrCreateWorksheet()
    
    ' 检查按钮是否已存在
    Dim btn As Shape
    For Each btn In ws.Shapes
        If btn.Name = "RefreshButton" Then
            btn.Delete
            Exit For
        End If
    Next btn
    
    ' 创建新按钮
    Set btn = ws.Shapes.AddShape(msoShapeRectangle, 10, 30, 80, 25)
    
    With btn
        .Name = "RefreshButton"
        .TextFrame.Characters.Text = "刷新价格"
        .TextFrame.Characters.Font.Size = 10
        .TextFrame.Characters.Font.Bold = True
        .Fill.ForeColor.RGB = RGB(70, 130, 180)
        .TextFrame.Characters.Font.ColorIndex = 2 ' 白色字体
        .OnAction = "RefreshAllPrices"
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "创建刷新按钮时发生错误：" & Err.Description, vbCritical, "按钮创建错误"
End Sub

' ========== 辅助函数：获取或创建工作表 ==========
Private Function GetOrCreateWorksheet() As Worksheet
    ' 获取或创建ETF价格工作表
    
    Dim ws As Worksheet
    Dim wsExists As Boolean
    
    ' 检查工作表是否存在
    wsExists = False
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name = WORKSHEET_NAME Then
            wsExists = True
            Set GetOrCreateWorksheet = ws
            Exit Function
        End If
    Next ws
    
    ' 如果不存在，则创建并初始化
    If Not wsExists Then
        InitializeWorksheet
        Set GetOrCreateWorksheet = ActiveWorkbook.Worksheets(WORKSHEET_NAME)
    End If
End Function

' ========== 辅助函数：获取最后一行数据 ==========
Private Function GetLastDataRow(ws As Worksheet) As Long
    ' 获取A列最后一行有数据的行号
    
    GetLastDataRow = ws.Cells(ws.Rows.Count, ETF_CODE_COLUMN).End(xlUp).Row
End Function

' ========== 辅助函数：检查ETF代码是否已存在 ==========
Private Function CheckETFCodeExists(ws As Worksheet, etfCode As String) As Boolean
    ' 检查ETF代码是否已在工作表中存在
    
    Dim lastRow As Long
    Dim currentRow As Long
    Dim existingCode As String
    
    lastRow = GetLastDataRow(ws)
    CheckETFCodeExists = False
    
    For currentRow = HEADER_ROW + 1 To lastRow
        existingCode = Trim(CStr(ws.Cells(currentRow, ETF_CODE_COLUMN).Value))
        If existingCode = etfCode Then
            CheckETFCodeExists = True
            Exit Function
        End If
    Next currentRow
End Function

' ========== 导出数据到CSV ==========
Public Sub ExportDataToCSV()
    ' 导出ETF数据到CSV文件
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = GetOrCreateWorksheet()
    
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    
    If lastRow <= HEADER_ROW Then
        MsgBox "没有数据可以导出", vbInformation, "导出数据"
        Exit Sub
    End If
    
    ' 选择保存位置
    Dim fileName As String
    fileName = Application.GetSaveAsFilename( _
        InitialFileName:="ETF_Price_Data_" & Format(Date, "yyyymmdd") & ".csv", _
        FileFilter:="CSV文件 (*.csv), *.csv", _
        Title:="导出ETF数据")
    
    If fileName = "False" Then Exit Sub
    
    ' 导出数据
    Dim exportRange As Range
    Set exportRange = ws.Range(ws.Cells(HEADER_ROW, ETF_CODE_COLUMN), ws.Cells(lastRow, DATE_COLUMN))
    
    Dim tempWb As Workbook
    Set tempWb = Application.Workbooks.Add
    
    exportRange.Copy tempWb.Worksheets(1).Range("A1")
    tempWb.SaveAs fileName, FileFormat:=xlCSV
    tempWb.Close SaveChanges:=False
    
    MsgBox "数据已成功导出到：" & vbCrLf & fileName, vbInformation, "导出完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "导出数据时发生错误：" & Err.Description, vbCritical, "导出错误"
End Sub
