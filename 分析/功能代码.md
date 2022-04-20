---
description: 代码及其注释
---

# 功能代码

```
#include<IE.au3>
#include <File.au3>
#include <Excel.au3>
#include <Array.au3>
#include <Date.au3>

Local $aRecords

$hFileLoc = FileOpenDialog("Spreadsheet Data", @ScriptDir, "Excel (*.xls;*.xlsx)")  ;选择一个excel文件
If @error Then Exit

$sDataFile = _TempFile(@TempDir, "~", ".txt")                           ;创建一个零时文件存储数据
ConsoleWrite("CSV file = (" & $sDataFile & ")" & @CRLF)                 ;我们可以在测试时手动查看

;*** 电子表格格式为数字,货币,日期等
$oExcel = _Excel_Open()                 ;打开 Excel 一个新的 Excel 应用程序
if @error Then
    MsgBox(0, "Error", "Failed to Open the Excel file for reading")
    Exit
EndIf





$oWorkbook = _Excel_BookOpen($oExcel, $hFileLoc, True, False)           ;让 Excel 以只读方式打开选定的电子表格

_Excel_BookSaveAs($oWorkbook, $sDataFile, $xlTextWindows, True)         ;另存为制表符分隔的文本文件, 用于 .csv

_Excel_BookClose($oWorkbook, False)                                     ;关闭选定的电子表格
_Excel_Close($oExcel, False)                                            ;关闭excel
If IsObj($oExcel) Then $oExcel.Quit                                     ;确保 Excel 实际关闭

_FileReadToArray($sDataFile, $aRecords, 0, @TAB)                        ;将新创建的制表符分隔文件读入数组
If @error Then
    MsgBox(0, "Error", "Failed to read the Excel data")
    Exit
EndIf

FileDelete($sDataFile)                                                  ;删除临时文件，因为它现在在内存中

;通常会带来额外的行和列，所以用这个循环清理它
For $r = 0 to UBound($aRecords) - 1
    If $aRecords[$r][0] = "" Then                                       ;如果 在这一行上是空的，那么假设我们已经到达数据的末尾
        ReDim $aRecords[$r][10]                                         ;删除多余的行，在这种情况下删除 之后的所有列
        ExitLoop
    EndIf
Next

_ArrayDisplay($aRecords,"Read Values")			;数组，以视图展示



Local $oIE = _IECreate("https://www.wjx.top/vj/wFIwk9z.aspx")	 ;打开指定网页

$tTime = _Date_Time_GetSystemTime()
$aTime = _Date_Time_SystemTimeToArray($tTime)		;获取当前系统时间

For $i = 1 To UBound($aRecords) - 1		;循环从数组里提取数据填写表单
	Local $q1 = _IEGetObjByName($oIE,"q1")
	$q1.value = $aRecords[$i][0]
	Local $q2 = _IEGetObjByName($oIE,"q2")
	$q2.value = $aRecords[$i][1]
	Local $q3 = _IEGetObjByName($oIE,"q3")
	$q3.value = $aRecords[$i][2]
	Local $q4 = _IEGetObjByName($oIE,"q4")
	$q4.value = $aRecords[$i][3]
	Local $q5 = _IEGetObjByName($oIE,"q5")
	$q5.value = $aRecords[$i][4]
	Local $q6 = _IEGetObjByName($oIE,"q6")
	$q6.value = $aRecords[$i][5]
	Local $q7 = _IEGetObjByName($oIE,"q7")
	$q7.value = $aRecords[$i][6]
	Local $q8 = _IEGetObjByName($oIE,"q8")
	$q8.value = $aTime[0]
	$Login = _IEGetObjById($oIE,"submit_button")		;获取按钮,并点击提交
	_IEAction($Login, "click")
	;在某一个窗口访问页面
	_IENavigate($oIE,"https://www.wjx.top/vj/wFIwk9z.aspx")
Next
```
