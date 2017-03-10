#include <IE.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <Excel.au3>

$yyyymm = "201702" ;check YEAR and MONTH
$idEmp = "50160"

Opt("WinTitleMatchMode", 2)


;;;;;;;;;;;;;;  Read data from Excel

; Create application object and open an example workbook
Local $oAppl = _Excel_Open()
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example", "Error creating the Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
Local $oWorkbook = _Excel_BookOpen($oAppl, @ScriptDir & "\timesheet_data.xlsx")
If @error Then
	MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example", "Error opening workbook '" & @ScriptDir & "\timesheet_data.xlsx'." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	_Excel_Close($oWorkbook)
	Exit
EndIf

; Read values
Local $aResult = _Excel_RangeRead($oWorkbook, 1, "A2:H100", 1)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 2", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
;MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 2", "Data successfully read." & @CRLF & "Please click 'OK' to display the formulas of cells A1:C1 of sheet 2.")
;_ArrayDisplay($aResult, "Array:")

Local $vEmpty = True
For $i = UBound($aResult) - 1 To 0 Step -1
	For $j = 0 To UBound($aResult, 2) - 1 Step 1
		If $aResult[$i][$j] <> "" Then
			$vEmpty = False
		EndIf
	Next
	If $vEmpty = True Then _ArrayDelete($aResult, $i)
	$vEmpty = True
Next

;_ArrayDisplay($aResult, "Array:")

_Excel_BookClose($oWorkbook, False)
;WinClose("timesheet_data - Excel")


;;;;;;;;;;;;;;  Submit data to ATPRO


Local $oIE = _IECreate("https://atpro.adastragrp.com/")
_IELoadWait($oIE)

Local $hWnd = WinWait("ATtendance & PROjects", "", 5)
;WinSetState($hWnd, "", @SW_MAXIMIZE)

;Local $oIE = _IE_Example("form")
Local $oSubmit = _IEGetObjByName($oIE, "ctl00_ContentPlaceHolder1_Button2")
_IEAction($oSubmit, "click")
_IELoadWait($oIE)


;~ Local $sMyString = "Set project (light)"
;~ Local $oLinks = _IELinkGetCollection($oIE)
;~ For $oLink In $oLinks
;~     Local $sLinkText = _IEPropertyGet($oLink, "innerText")
;~     If StringInStr($sLinkText, $sMyString) Then
;~         _IEAction($oLink, "click")
;~         ExitLoop
;~     EndIf
;~ Next
_IENavigate($oIE, "https://atpro.adastragrp.com/SetProjectLight.aspx?i=" & $yyyymm & "2" & $idEmp & "9")
_IELoadWait($oIE)

Local $iRows = UBound($aResult, $UBOUND_ROWS)
For $i = 0 To $iRows - 1
	Local $oForm = _IEFormGetObjByName($oIE, "SetProjectLight")
	If $i = 10 Or $i = 20 Or $i = 30 Or $i = 40 Then
		_IEFormSubmit($oForm)
		Local $oIE = _IECreate("https://atpro.adastragrp.com/SetProjectLight.aspx?i=" & $yyyymm & "2" & $idEmp & "9")
		_IELoadWait($oIE)
		Local $oForm = _IEFormGetObjByName($oIE, "SetProjectLight")
		Sleep(3000)
	EndIf
	$iteration = Floor($i / 10)
	$formEntry = $i + 1 - (10 * $iteration)
	;MsgBox(0,"", $i & " " & $formEntry & " " & $iteration)
	$date = $aResult[$i][6]
	$project = $aResult[$i][1]
	$category = $aResult[$i][2]
	$bt = $aResult[$i][3]
	$hour = $aResult[$i][7]
	$remark = $aResult[$i][4]


	Local $oSelect = _IEFormElementGetObjByName($oForm, "Date" & $formEntry)
	_IEAction($oSelect, "focus")
	_IEFormElementOptionSelect($oSelect, $date, 1, "byValue")

	Local $oSelect = _IEFormElementGetObjByName($oForm, "Project" & $formEntry)
	_IEAction($oSelect, "focus")
	_IEFormElementOptionSelect($oSelect, $project, 1, "byValue")

	Local $oSelect = _IEFormElementGetObjByName($oForm, "Category" & $formEntry)
	_IEAction($oSelect, "focus")
	_IEFormElementOptionSelect($oSelect, $category, 1, "byValue")

	Local $oSelect = _IEFormElementGetObjByName($oForm, "BT" & $formEntry)
	_IEAction($oSelect, "focus")
	If $bt = 1 Then
		Send("{SPACE}")
	EndIf

	Local $oSelect = _IEFormElementGetObjByName($oForm, "Hour" & $formEntry)
	_IEAction($oSelect, "focus")
	_IEFormElementSetValue($oSelect, $hour)

	Local $oSelect = _IEFormElementGetObjByName($oForm, "Remark" & $formEntry)
	_IEAction($oSelect, "focus")
	_IEFormElementSetValue($oSelect, $remark)

Next

_IEFormSubmit($oForm)


Exit

Func _Array2DDeleteEmptyRows(ByRef $iArray)
	Local $vEmpty = True
	For $i = UBound($iArray) - 1 To 0 Step -1
		For $j = 0 To UBound($iArray, 2) - 1 Step 1
			If $iArray[$i][$j] <> "" Then
				$vEmpty = False
			EndIf
		Next
		If $vEmpty = True Then _ArrayDelete($iArray, $i)
		$vEmpty = True
	Next
EndFunc   ;==>_Array2DDeleteEmptyRows







