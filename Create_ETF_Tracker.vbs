' ========== Excel ETF价格追踪器自动创建脚本 ==========
' 此脚本将创建一个完整的Excel工作簿，包含所有VBA模块和功能
' 运行此脚本前，请确保vba_modules文件夹中包含所有必要的VBA文件

Option Explicit

Dim xlApp, xlWorkbook, xlWorksheet
Dim fso, scriptPath, vbaModulesPath
Dim objShell

' 创建文件系统对象
Set fso = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' 获取脚本路径
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)
vbaModulesPath = scriptPath & "\vba_modules"

' 检查VBA模块文件夹是否存在
If Not fso.FolderExists(vbaModulesPath) Then
    WScript.Echo "错误：找不到vba_modules文件夹"
    WScript.Echo "请确保VBA模块文件位于: " & vbaModulesPath
    WScript.Quit 1
End If

WScript.Echo "开始创建ETF价格追踪器..."

On Error Resume Next

' 创建Excel应用程序
Set xlApp = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "错误：无法启动Excel应用程序"
    WScript.Echo "请确保已安装Microsoft Excel"
    WScript.Quit 1
End If

On Error GoTo 0

' 设置Excel可见性（可选）
xlApp.Visible = True
xlApp.DisplayAlerts = False

' 创建新工作簿
Set xlWorkbook = xlApp.Workbooks.Add

WScript.Echo "正在导入VBA模块..."

' 导入VBA模块
Call ImportVBAModule("Module_Config.bas", 1)  ' 1 = vbext_ct_StdModule
Call ImportVBAModule("JsonConverter.bas", 1)
Call ImportVBAModule("Module_API.bas", 1)
Call ImportVBAModule("Module_Refresh.bas", 1)

' 导入ThisWorkbook类模块（需要替换现有的）
Call ImportThisWorkbookModule("ThisWorkbook.cls")

WScript.Echo "正在初始化工作表..."

' 获取默认工作表
Set xlWorksheet = xlWorkbook.Worksheets(1)
xlWorksheet.Name = "ETF价格"

' 设置表头
xlWorksheet.Cells(1, 1).Value = "ETF代码"
xlWorksheet.Cells(1, 2).Value = "最新收盘价"
xlWorksheet.Cells(1, 3).Value = "数据日期"

' 格式化表头
With xlWorksheet.Range("A1:C1")
    .Font.Bold = True
    .Interior.Color = RGB(200, 200, 200)
    .HorizontalAlignment = -4108  ' xlCenter
End With

' 设置列宽
xlWorksheet.Columns(1).ColumnWidth = 12  ' ETF代码列
xlWorksheet.Columns(2).ColumnWidth = 15  ' 收盘价列
xlWorksheet.Columns(3).ColumnWidth = 15  ' 日期列

' 冻结首行
xlWorksheet.Range("A2").Select
xlApp.ActiveWindow.FreezePanes = True

' 添加一些示例数据
xlWorksheet.Cells(2, 1).Value = "510300"  ' 沪深300ETF
xlWorksheet.Cells(3, 1).Value = "512690"  ' 白酒ETF
xlWorksheet.Cells(4, 1).Value = "516160"  ' 新能源ETF

WScript.Echo "正在创建刷新按钮..."

' 创建刷新按钮
Call CreateRefreshButton()

' 启用宏（设置信任中心）
xlWorkbook.Application.AutomationSecurity = 1  ' msoAutomationSecurityLow

' 保存工作簿
Dim saveFileName
saveFileName = scriptPath & "\ETF_Price_Tracker.xlsm"

On Error Resume Next
xlWorkbook.SaveAs saveFileName, 52  ' 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
If Err.Number <> 0 Then
    WScript.Echo "警告：无法保存为.xlsm格式，尝试保存为.xls格式"
    Err.Clear
    saveFileName = scriptPath & "\ETF_Price_Tracker.xls"
    xlWorkbook.SaveAs saveFileName, -4143  ' xlWorkbookNormal
    If Err.Number <> 0 Then
        WScript.Echo "错误：保存文件失败 - " & Err.Description
    End If
End If
On Error GoTo 0

WScript.Echo "ETF价格追踪器创建完成！"
WScript.Echo "文件位置: " & saveFileName
WScript.Echo ""
WScript.Echo "使用说明："
WScript.Echo "1. 在A列输入ETF代码（6位数字）"
WScript.Echo "2. 点击'刷新价格'按钮获取最新价格"
WScript.Echo "3. 或使用快捷键 Ctrl+Shift+R 刷新所有价格"
WScript.Echo "4. API Token已预配置，可直接使用"

' 询问是否立即测试
Dim testResult
testResult = MsgBox("是否立即测试API连接？", vbYesNo + vbQuestion, "测试连接")
If testResult = vbYes Then
    Call TestAPIConnection()
End If

' 清理对象
Set xlWorksheet = Nothing
Set xlWorkbook = Nothing
Set xlApp = Nothing
Set fso = Nothing
Set objShell = Nothing

WScript.Echo "脚本执行完成！"

' ========== 子程序：导入VBA模块 ==========
Sub ImportVBAModule(fileName, moduleType)
    Dim moduleFile
    moduleFile = vbaModulesPath & "\" & fileName
    
    If fso.FileExists(moduleFile) Then
        On Error Resume Next
        xlWorkbook.VBProject.VBComponents.Import moduleFile
        If Err.Number = 0 Then
            WScript.Echo "已导入: " & fileName
        Else
            WScript.Echo "导入失败: " & fileName & " - " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Else
        WScript.Echo "警告：找不到文件 " & fileName
    End If
End Sub

' ========== 子程序：导入ThisWorkbook模块 ==========
Sub ImportThisWorkbookModule(fileName)
    Dim moduleFile, thisWB, fileContent
    moduleFile = vbaModulesPath & "\" & fileName
    
    If fso.FileExists(moduleFile) Then
        On Error Resume Next
        
        ' 获取ThisWorkbook对象
        Set thisWB = xlWorkbook.VBProject.VBComponents("ThisWorkbook")
        
        ' 读取文件内容
        Dim textStream, moduleCode
        Set textStream = fso.OpenTextFile(moduleFile, 1)
        moduleCode = textStream.ReadAll
        textStream.Close
        
        ' 跳过CLASS文件的头部信息，只提取代码部分
        Dim lines, i, codeStart
        lines = Split(moduleCode, vbCrLf)
        codeStart = 0
        
        For i = 0 To UBound(lines)
            If InStr(lines(i), "Option Explicit") > 0 Then
                codeStart = i
                Exit For
            End If
        Next
        
        ' 重新构建代码
        Dim actualCode
        actualCode = ""
        For i = codeStart To UBound(lines)
            actualCode = actualCode & lines(i) & vbCrLf
        Next
        
        ' 清除现有代码并添加新代码
        thisWB.CodeModule.DeleteLines 1, thisWB.CodeModule.CountOfLines
        thisWB.CodeModule.AddFromString actualCode
        
        WScript.Echo "已更新: ThisWorkbook"
        
        On Error GoTo 0
    Else
        WScript.Echo "警告：找不到文件 " & fileName
    End If
End Sub

' ========== 子程序：创建刷新按钮 ==========
Sub CreateRefreshButton()
    On Error Resume Next
    
    Dim btn
    ' 删除可能存在的旧按钮
    xlWorksheet.Shapes("RefreshButton").Delete
    
    ' 创建新按钮
    Set btn = xlWorksheet.Shapes.AddShape(1, 10, 30, 80, 25)  ' msoShapeRectangle = 1
    btn.Name = "RefreshButton"
    btn.TextFrame.Characters.Text = "刷新价格"
    btn.TextFrame.Characters.Font.Size = 10
    btn.TextFrame.Characters.Font.Bold = True
    btn.Fill.ForeColor.RGB = RGB(70, 130, 180)
    btn.TextFrame.Characters.Font.ColorIndex = 2  ' 白色
    btn.OnAction = "RefreshAllPrices"
    
    WScript.Echo "已创建刷新按钮"
    On Error GoTo 0
End Sub

' ========== 子程序：测试API连接 ==========
Sub TestAPIConnection()
    On Error Resume Next
    
    WScript.Echo "正在测试API连接..."
    
    ' 运行API测试
    xlApp.Run "TestApiConnection"
    
    On Error GoTo 0
End Sub

Function RGB(r, g, b)
    RGB = r + (g * 256) + (b * 65536)
End Function
