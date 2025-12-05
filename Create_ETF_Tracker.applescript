-- ========== ETF价格追踪器自动创建脚本 (AppleScript版本) ==========
-- 适用于Mac系统的Microsoft Excel
-- 此脚本将创建Excel工作簿并设置基本结构

-- 检查Excel是否已安装
try
	tell application "Finder"
		exists application file id "com.microsoft.Excel"
	end tell
on error
	display alert "错误" message "未找到Microsoft Excel，请先安装Excel for Mac" buttons {"确定"} default button "确定"
	return
end try

-- 显示开始消息
display notification "开始创建ETF价格追踪器..." with title "ETF追踪器"

try
	-- 启动Excel
	tell application "Microsoft Excel"
		activate
		
		-- 创建新工作簿
		set newWorkbook to make new workbook
		
		-- 获取第一个工作表
		set ws to worksheet 1 of newWorkbook
		
		-- 重命名工作表
		set name of ws to "ETF价格"
		
		-- 设置表头
		set value of cell "A1" of ws to "ETF代码"
		set value of cell "B1" of ws to "最新收盘价"
		set value of cell "C1" of ws to "数据日期"
		
		-- 格式化表头
		set headerRange to range "A1:C1" of ws
		
		-- 设置表头格式
		tell headerRange
			set bold of font object to true
			set size of font object to 12
			set color index of interior to 15 -- 浅灰色背景
			set horizontal alignment to center
		end tell
		
		-- 设置列宽
		set column width of column "A:A" of ws to 12
		set column width of column "B:B" of ws to 15
		set column width of column "C:C" of ws to 15
		
		-- 添加示例数据
		set value of cell "A2" of ws to "510300" -- 沪深300ETF
		set value of cell "A3" of ws to "512690" -- 白酒ETF
		set value of cell "A4" of ws to "516160" -- 新能源ETF
		
		-- 添加说明文字
		set value of cell "E1" of ws to "使用说明："
		set value of cell "E2" of ws to "1. 在A列输入ETF代码"
		set value of cell "E3" of ws to "2. 运行VBA宏刷新价格"
		set value of cell "E4" of ws to "3. 需要先导入VBA模块"
		
		-- 格式化说明文字
		set bold of font object of cell "E1" of ws to true
		set color index of font object of range "E1:E4" of ws to 5 -- 蓝色文字
		
		-- 冻结首行
		select cell "A2" of ws
		freeze panes
		
		-- 保存工作簿
		set desktopPath to (path to desktop as string)
		set fileName to desktopPath & "ETF_Price_Tracker.xlsm"
		
		-- 尝试保存为启用宏的格式
		try
			save workbook as newWorkbook filename fileName file format Excel12 format
			display notification "工作簿已保存为: ETF_Price_Tracker.xlsm" with title "保存成功"
		on error
			-- 如果.xlsm格式失败，保存为.xlsx格式
			set fileName to desktopPath & "ETF_Price_Tracker.xlsx"
			save workbook as newWorkbook filename fileName file format open XML workbook format
			display notification "工作簿已保存为: ETF_Price_Tracker.xlsx" with title "保存成功"
		end try
		
	end tell
	
	-- 显示VBA导入说明
	set vbaInstructions to "Excel工作簿已创建完成！

接下来需要手动导入VBA模块：

1. 在Excel中按 Option+F11 打开VBA编辑器

2. 右键点击左侧的VBAProject，选择"导入文件"

3. 依次导入以下文件：
   • Module_Config.bas
   • JsonConverter.bas  
   • Module_API.bas
   • Module_Refresh.bas

4. 双击"ThisWorkbook"，删除现有代码

5. 打开 ThisWorkbook.cls 文件，复制代码内容
   （跳过文件开头的 VERSION 和 BEGIN 部分）

6. 粘贴到ThisWorkbook模块中

7. 保存工作簿（Cmd+S）

VBA模块路径: " & (POSIX path of (path to desktop)) & "AIProject/AI-Execl/vba_modules/"
	
	display alert "创建完成" message vbaInstructions buttons {"打开VBA模块文件夹", "稍后手动操作"} default button "打开VBA模块文件夹"
	
	if button returned of result is "打开VBA模块文件夹" then
		-- 打开VBA模块文件夹
		try
			set vbaFolderPath to (POSIX path of (path to desktop)) & "AIProject/AI-Execl/vba_modules/"
			do shell script "open " & quoted form of vbaFolderPath
		on error
			display alert "提示" message "请手动导航到VBA模块文件夹：
~/Desktop/AIProject/AI-Execl/vba_modules/" buttons {"确定"} default button "确定"
		end try
	end if
	
on error errorMessage
	display alert "创建过程中发生错误" message errorMessage buttons {"确定"} default button "确定"
end try

-- 创建使用脚本
try
	set usageScriptContent to "-- ETF价格刷新快捷脚本
tell application \"Microsoft Excel\"
	activate
	try
		-- 运行刷新所有价格的宏
		run VB macro \"RefreshAllPrices\"
		display notification \"价格刷新完成\" with title \"ETF追踪器\"
	on error
		display alert \"提示\" message \"请确保已导入所有VBA模块\" buttons {\"确定\"} default button \"确定\"
	end try
end tell"
	
	set usageScriptPath to (path to desktop as string) & "刷新ETF价格.applescript"
	
	-- 写入使用脚本
	set fileRef to open for access file usageScriptPath with write permission
	set eof fileRef to 0
	write usageScriptContent to fileRef
	close access fileRef
	
	display notification "已创建快捷刷新脚本" with title "额外功能"
	
on error
	-- 忽略脚本创建错误
end try

-- 显示完成信息
display notification "ETF价格追踪器创建完成！" with title "创建成功"

-- 创建API测试脚本
try
	set testScriptContent to "-- API连接测试脚本
tell application \"Microsoft Excel\"
	activate
	try
		-- 运行API测试宏
		run VB macro \"TestApiConnection\"
	on error
		display alert \"提示\" message \"请确保已导入所有VBA模块\" buttons {\"确定\"} default button \"确定\"
	end try
end tell"
	
	set testScriptPath to (path to desktop as string) & "测试API连接.applescript"
	
	-- 写入测试脚本
	set fileRef to open for access file testScriptPath with write permission
	set eof fileRef to 0
	write testScriptContent to fileRef
	close access fileRef
	
on error
	-- 忽略脚本创建错误
end try

-- 最终提示
set finalMessage to "🎉 ETF价格追踪器创建完成！

已创建的文件：
• ETF_Price_Tracker.xlsm (主工作簿)
• 刷新ETF价格.applescript (快捷刷新)
• 测试API连接.applescript (API测试)

下一步操作：
1. 导入VBA模块（按照提示操作）
2. 在A列输入ETF代码测试
3. 运行刷新脚本获取价格

所有文件已保存到桌面。"

display alert "安装完成" message finalMessage buttons {"确定"} default button "确定"
