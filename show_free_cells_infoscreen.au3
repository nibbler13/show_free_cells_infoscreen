#pragma compile(ProductVersion, 0.3)
#pragma compile(UPX, true)
#pragma compile(CompanyName, 'ООО Клиника ЛМС')
#pragma compile(FileDescription, Скрипт для отображения свободных для записи ячеек в инфоклинике)
#pragma compile(LegalCopyright, Грашкин Павел Павлович - Нижний Новгород - 31-555)
#pragma compile(ProductName, show_free_cells_infoscreen)

AutoItSetOption("TrayAutoPause", 0)

#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <FontConstants.au3>
#include <WindowsConstants.au3>
#include <WinAPI.au3>
#include <Array.au3>
#include <ColorConstants.au3>
#include <GuiListView.au3>
#include <ListviewConstants.au3>
#include <Debug.au3>
#include <File.au3>
#include <Date.au3>
#include <String.au3>
#include <Math.au3>

#Region ======================== Variables ========================
Local $oMyError = ObjEvent("AutoIt.Error", "HandleComError")

Local $delay = 15000

Local $textColor = 0x2c3d3f
Local $alternateTextColor = 0xffffff
Local $mainBackgroundColor = 0xffffff

Local $childGuiColor = $mainBackgroundColor
Local $childTotalLines = 14
Local $childFirstCellFactorPercentage = 30

Local $freeCellColor = 0xf2f9e0
Local $docNameColor = 0xaa00aa
Local $departmentNameColor = 0x2fbf5f;0x2db55a
Local $borderColor = 0xd6d6d6

Local $bottonLineHeight = 14

;~ Local $dX = @DesktopWidth
;~ Local $dY = @DesktopHeight
Local $dX = 1024
Local $dY = 768


Local $headerHeightPartFromTotalHeight = 9
Local $headerHeight = Round($dY / $headerHeightPartFromTotalHeight)

Local $mainFontSize = Round($headerHeight / 3)
Local $mainFontName = "Franklin Gothic"
Local $mainFontWeight = $FW_BOLD
Local $mainFontQuality = $CLEARTYPE_QUALITY

Local $mainGuiGap = $mainFontSize

Local $headerLabelFontWeight = $FW_SEMIBOLD
Local $headerTextColor = $textColor
Local $headerColor = $mainBackgroundColor

Local $timeLabelHeight = Round($mainFontSize * 1.7)

Local $minutesToShow = 180

Local $showHoursUnderTimeLines = False
Local $showTimeLines = True
Local $showEveryMinuteLines = True

Local $minuteTimeLineColor = 0xeeeeee
Local $halfHourTimeLineColor = 0x999999
Local $hoursColor = 0x666666

Local $showOnlyDepartments = True


Local $childGui = 0
Local $childGuiOld = 0

Local $animationDurationMs = 0
Local $percentageThatFreeCellMustHaveFromInterval = 100


Local $titleText = "Свободные места для записи на ближайшее время"
;~ Local $titleText = "При записи к специалистам в указанное время"  & @CRLF & "предоставляется СКИДКА 5% при наличном расчете"

Local $excludedDepartments[] = ["физиопроцедуры", _
							  "рентген", _
							  "процедурный", _
							  "предрейсовый осмотр", _
							  "стационар", _
							  "анестезиология-реаниматология", _
							  "помощь на дому"]




Local $messageToSend = ""
Local $current_pc_name = @ComputerName
Local $errStr = "===ERROR=== "
ConsoleWrite("Current_pc_name: " & $current_pc_name & @CRLF)

Local $logFilePath = @ScriptDir & "\" & @ScriptName & ".log"
Local $logFile = FileOpen($logFilePath, $FO_OVERWRITE)
ToLog($current_pc_name)
ToLog(@CRLF & "---Check for temp folder and create log---")

If $logFile = -1 Then ToLog($errStr & "Cannot create log file")
#EndRegion ======================== Variables ========================


;================= CURSOR HIDE ===============
;~ _WinAPI_ShowCursor(False)

ShowGui()


Func ShowGui()
	Local $mainGui = GUICreate("ShowFreeCells", $dX, $dY, 0, 0, $WS_POPUPWINDOW) ;, $WS_EX_TOPMOST)
	Local $titleLabel = CreateStandardDesign($mainGui, $titleText, True)

	Local $timeLabel = GuiCtrlCreateLabel("", $mainGuiGap, $dY - $timeLabelHeight - $bottonLineHeight, _
		$dX - $mainGuiGap * 2, $timeLabelHeight, BitOr($SS_CENTER, $SS_CENTERIMAGE))
	GUICtrlSetColor(-1, $textColor)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	UpdateTimeLabel($timeLabel)

	GUISetState(@SW_SHOW)

	Local $titleLabelPosition = ControlGetPos("", "", $titleLabel)
	Local $timeLabelPosition = ControlGetPos("", "", $timeLabel)
	Local $childWidth = $dX - $mainGuiGap * 2
	Local $childHeight = $timeLabelPosition[1] - $titleLabelPosition[1] - $titleLabelPosition[3] - $mainGuiGap
	Local $childX = $mainGuiGap
	Local $childY = $titleLabelPosition[1] + $titleLabelPosition[3]

	Local $firstCellWidth = Round($childWidth * ($childFirstCellFactorPercentage / 100)) ;0.3
	Local $cellHeight = Round($childHeight / $childTotalLines);($mainFontSize / 2))

	Local $childFontSize = Round($cellHeight * 0.45)

	Local $childGuiGap = Round($cellHeight / 10)
	Local $minuteX = Round(($childWidth - $firstCellWidth) / $minutesToShow)

	$firstCellWidth = $childWidth - $minutesToShow * $minuteX

	Local $nothingFoundText = "Уважаемые пациенты," & @CRLF & @CRLF & "к сожалению, на данный момент" & @CRLF & _
		"нет свободных для записи мест" & @CRLF & @CRLF & "Подробную информацию" & @CRLF & _
		"Вы можете получить на регистратуре"

	Local $records = 0

	While True
		If $childGui And $animationDurationMs Then
			GUIDelete($childGui)
		EndIf

		Local $now = _NowCalc()
		Local $xpos = 0
		Local $ypos = 0

		$childGui = CreateChildGui($childWidth, $childHeight, $childX, $childY, $mainGui, $childFontSize, _
			$xpos, $ypos, $firstCellWidth, $minuteX, $childGuiGap, $childFontSize, $now, IsArray($records) ? True : False)

		If Not IsArray($records) Then
			CreateLabel($nothingFoundText, 0, 0, $childWidth, $childHeight, $textColor, $GUI_BKCOLOR_TRANSPARENT, $childGui)
			GUISetState()

			If Not $animationDurationMs Then
				ToLog($childGui & " " & $childGuiOld)
				If $childGuiOld Then GUIDelete($childGuiOld)
				$childGuiOld = $childGui
			EndIf
		EndIf

		For $i = 0 To UBound($records, $UBOUND_ROWS) - 1
			ToLog("----DEPT CYCLE----")
			If $ypos + $childGuiGap * 2 + $cellHeight * 2 > $childHeight Then
				ToLog("====DEPARTMENT HEIGHT TOO MUCH=====")
				FinalizeScreen($xpos, $ypos, $childWidth, $childHeight - $ypos)

				Sleep($delay)

				If $animationDurationMs Then GUIDelete($childGui)
				$childGui = CreateChildGui($childWidth, $childHeight, $childX, $childY, $mainGui, $childFontSize, _
					$xpos, $ypos, $firstCellWidth, $minuteX, $childGuiGap, $childFontSize, $now)
			EndIf

			Local $deptName = $records[$i][0]
			ToLog($deptName)

			If Not $showOnlyDepartments Then
				CreateButton($deptName, $xpos, $ypos, $childWidth, $cellHeight, _
					$departmentNameColor, $departmentNameColor, $alternateTextColor)
				Local $deptNameFontSize = $childFontSize
				If StringLen($deptName) * $childFontSize > $childWidth Then $deptNameFontSize = $childWidth / StringLen($deptName)
				GUICtrlSetFont(-1, $deptNameFontSize, $FW_MEDIUM, -1, "Franklin Gothic Book")

				$ypos += $cellHeight
			Else
				FinalizeScreen($xpos, $ypos, $childWidth, $childHeight - $ypos, True)
			EndIf

			$ypos += $childGuiGap

			Local $elArray = $records[$i][1]
			For $x = 0 To UBound($elArray) - 1
				ToLog("----DOC CYCLE----")
				If $ypos + $cellHeight + $childGuiGap > $childHeight Then
					ToLog("======DOC NAME HEIGHT TOO MUCH======")
					$ypos -= $childGuiGap
					FinalizeScreen($xpos, $ypos, $childWidth, $childHeight - $ypos)

					Sleep($delay)

					If $animationDurationMs Then GUIDelete($childGui)

					$childGui = CreateChildGui($childWidth, $childHeight, $childX, $childY, $mainGui, $childFontSize, _
						$xpos, $ypos, $firstCellWidth, $minuteX, $childGuiGap, $childFontSize, $now)

					If Not $showOnlyDepartments Then
						CreateButton($deptName, $xpos, $ypos, $childWidth, $cellHeight, _
							$departmentNameColor, $departmentNameColor, $alternateTextColor)
						GUICtrlSetFont(-1, -1, $FW_MEDIUM, -1, "Franklin Gothic Book")

						$ypos += $cellHeight + $childGuiGap
					EndIf
				EndIf

				Local $docArray = StringSplit($elArray[$x], "@", $STR_NOCOUNT)
				Local $docName = $docArray[0]
				Local $docNameBackgroundColor = $showOnlyDepartments ? $departmentNameColor : $mainBackgroundColor
				Local $docNameTextColor = $showOnlyDepartments ? $alternateTextColor : $textColor

				CreateButton($docName, $xpos, $ypos, $firstCellWidth - $minuteX, $cellHeight, $docNameBackgroundColor, _
					$docNameBackgroundColor, $docNameTextColor)
				Local $docNameFontSize = $childFontSize
				If StringLen($docName) * $childFontSize > $firstCellWidth Then $docNameFontSize = $firstCellWidth / StringLen($docName)
				GUICtrlSetFont(-1, $docNameFontSize, $FW_MEDIUM)

				;================ LINE BETWEEN DOCTORS NAME =================
				If $x < UBound($elArray) - 1 Or $showOnlyDepartments Then
					GuiCtrlCreateLabel("", $xpos, $ypos + $cellHeight + $childGuiGap, $childWidth, 1)
					GUICtrlSetBkColor(-1, $borderColor)
				EndIf ;===== END

				For $y = 1 To UBound($docArray) - 1
					If Not StringLen($docArray[$y]) Then ExitLoop

					Local $cellInfo = StringSplit($docArray[$y], ";", $STR_NOCOUNT)
					Local $text = $cellInfo[0]
					Local $startTime = $cellInfo[1]


					;===== NEW =====
					Local $interval = $cellInfo[2]

					If StringLeft($text, 1) = "0" Then $text = StringRight($text, 4)

					Local $currentCellX = _DateDiff('n', $now, $startTime) * $minuteX + $firstCellWidth + 1
					Local $currentCellWidth = $minuteX * $interval

					CreateButton($text, $currentCellX, $ypos, $currentCellWidth, $cellHeight)
					If $childFontSize * 4 > $currentCellWidth Then GUICtrlSetFont(-1, $currentCellWidth / 4)
				Next

				$ypos += $cellHeight + $childGuiGap * 2
			Next

			$ypos -= $childGuiGap

			If $i = UBound($records, $UBOUND_ROWS) - 1 Then
				FinalizeScreen($xpos, $ypos, $childWidth, $childHeight - $ypos)
			EndIf

			_ArraySort($records[$i][1])
			ToLog(_ArrayToString($records[$i][1], @CRLF))
		Next

		Local $timer = TimerInit()

		$records = GetData($now)
		_ArraySort($records)

		UpdateTimeLabel($timeLabel)

		Sleep($delay - TimerDiff($timer))
	WEnd
EndFunc


Func FinalizeScreen($x, $y, $width, $height, $onlyTopLine = False)
	If Not $onlyTopLine Then
		GUICtrlCreateLabel("", $x, $y, $width, $height)
		GUICtrlSetBkColor(-1, $childGuiColor)
	EndIf

	GuiCtrlCreateLabel("", $x, $y, $width, 1)
	GUICtrlSetBkColor(-1, $borderColor)



	If Not $animationDurationMs And Not $onlyTopLine Then
		GUISetState()
		If $childGuiOld Then GUIDelete($childGuiOld)
		$childGuiOld = $childGui
	EndIf
EndFunc


Func CreateChildGui($width, $height, $x, $y, $root, $fontSize, ByRef $xpos, ByRef $ypos, _
	$firstCellWidth, $minuteX, $childGuiGap, $childFontSize, $now, $timeLines = True)
	$xpos = 0
	$ypos = 0

	Local $newGui = GUICreate("Child", $width, $height, $x, $y, $WS_CHILD, $WS_EX_TOPMOST, $root)
	GUISetFont($fontSize) ;$mainFontSize * 0.8)
;~ 	Local $color = "0x" & Random(0, 9, 1) & "f" & Random(0, 9, 1) & "f" & Random(0, 9, 1) & "f"
;~ 	ToLog("=============" & $color)
	GUISetBkColor($childGuiColor)

	If $showTimeLines And $timeLines Then
		DrawTimeLines($firstCellWidth + 1, $minuteX, $childGuiGap * 5, $width, $height, $childFontSize, $now)
		$ypos += $childGuiGap * ($showHoursUnderTimeLines ? 5 : 0)
	EndIf

	If $animationDurationMs Then GUISetState()

	Return $newGui
EndFunc


Func CreateButton($text, $x, $y, $width, $height, $bkColor = $freeCellColor, $brColor = $borderColor, $txColor = $textColor)
	GUICtrlCreateLabel("", $x + 1, $y + 1, $width - 1, $height - 2)
	GUICtrlSetBkColor(-1, $brColor)
	Local $id = GUICtrlCreateLabel($text, $x + 2, $y + 2, $width - 3, $height - 4, BitOR($SS_CENTER, $SS_CENTERIMAGE))
	GUICtrlSetColor(-1, $txColor)
	GUICtrlSetBkColor(-1, $bkColor)
	Sleep($animationDurationMs)
	Return $id
EndFunc


Func CreateStandardDesign($gui, $titleText, $trademark = False)
	GUISetBkColor($mainBackgroundColor)
	GUISetFont($mainFontSize, $mainFontWeight, 0, $mainFontName, $gui, $mainFontQuality)
;~ 	Local $fontFactor = StringInStr($titleText, @CRLF) ? 0.8 : 1
	If StringInStr($titleText, @CRLF) Then $mainFontSize *= 0.8
	Local $titleLabel = CreateLabel($titleText, 0, 0, $dX, $headerHeight, $headerTextColor, $headerColor, $gui, $mainFontSize)
	GUICtrlSetFont(-1, $mainFontSize, $headerLabelFontWeight)
	GUICtrlCreatePic(@ScriptDir & "\picBottomLine.jpg", 0, $dY - $bottonLineHeight, $dX, $bottonLineHeight)

	If $trademark Then
		Local $trademarkWidth = $timeLabelHeight * 0.9
		Local $trademarkHeight = $timeLabelHeight
		GUICtrlCreatePic(@ScriptDir & "\picButterfly.jpg", $dX - $trademarkWidth - $mainGuiGap / 2, _
			$dY - $trademarkHeight - 11 - $mainGuiGap / 2, $trademarkWidth, $trademarkHeight)
	EndIf

	Return $titleLabel
EndFunc


Func CreateLabel($text, $x, $y, $width, $height, $textColor, $backgroundColor, $gui, $fontSize = $mainFontSize)
;~ 	ToLog("FOntSIze: " & $fontSize)
	Local $titleLabel = GUICtrlCreateLabel("", $x, $y, $width, $height)
	GUICtrlSetBkColor(-1, $backgroundColor)

	GUISetFont($fontSize)
	Local $label = GUICtrlCreateLabel($text, 0, 0, -1, -1, $SS_CENTER)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetColor(-1, $textColor)

	Local $position = ControlGetPos($gui, "", $label)
	If IsArray($position) Then
		Local $newX = Round($x + ($width - $position[2] ) / 2)
		Local $newY = Round($y + ($height - $position[3]) / 2)
		GUICtrlSetPos($label, $newX, $newY)
	EndIf
	GUISetFont($mainFontSize)

	Return $titleLabel
EndFunc


Func DrawTimeLines($startX, $oneMinuteX, $headerHeight, $totalWidth, $totalHeight, $fontSize, $now)
	ToLog("DrawTimeLines $startX: " & $startX & " $oneMinuteX: " & $oneMinuteX & " totalWidth: " & $totalWidth)

	If $showHoursUnderTimeLines Then
		$headerHeight += 1
	Else
		$headerHeight = 0
	EndIf

	For $i = 0 To $minutesToShow
		Local $tmpDate = _DateAdd('n', $i, $now)
		Local $width = 1
		Local $color = $minuteTimeLineColor

		If StringMid($tmpDate, 15, 2) = "30" Or StringMid($tmpDate, 15, 2) = "00" Then
			$color = $halfHourTimeLineColor
			Local $dateX = $startX + $oneMinuteX * $i - $oneMinuteX
			Local $dateY = 0

			If $dateX + $fontSize * 2 <= $totalWidth And $showHoursUnderTimeLines Then
				Local $text = StringMid($tmpDate, 12, 5)
				GUICtrlCreateLabel($text, $dateX - $fontSize * 1.5, $dateY, $fontSize * 3, $headerHeight, _
					BitOR($SS_CENTER, $SS_CENTERIMAGE))
				GUICtrlSetFont(-1, $fontSize * 0.8, $FW_BOLD)
				GUICtrlSetColor(-1, $hoursColor)
			EndIf
		EndIf

		Local $currentMinute = StringMid($tmpDate, 16, 1)
		If $currentMinute = "0" Or $currentMinute = "5" Or $showEveryMinuteLines Then
			GUICtrlCreateLabel("", $startX + $oneMinuteX * $i - $oneMinuteX, $headerHeight, $width, $totalHeight - $headerHeight)
			GUICtrlSetBkColor(-1, $color)
		EndIf
	Next
EndFunc


Func UpdateTimeLabel($label)
	Local $currentHour = @HOUR >= 10 ? @HOUR : StringRight(@HOUR, 1)
	Local $newText = "Текущее время - " & $currentHour & ":" & @MIN
	GUICtrlSetData($label, $newText)
EndFunc


Func GetData($now)
	ToLog(@CRLF & @CRLF & "======= GETTING DATA =======")

	Local $sqlDoc = "Select D.DName, Ds.DCode, Dep.DepName, Ds.DepNum, Ds.BegHour, Ds.BegMin, Ds.EndHour, Ds.EndMin, Ds.Shinterv " & _
			"From DoctShedule Ds " & _
			"Join Doctor D On D.DCode = Ds.DCode " & _
			"Join Departments Dep On Dep.DepNum = Ds.DepNum " & _
			"Where Ds.WDate = 'today' " & _
			"And Ds.Shinterv Is Not Null"

	Local $doctors = ExecuteSQL($sqlDoc)

	If Not IsArray($doctors) Or UBound($doctors, $UBOUND_ROWS) = 0 Then
		ToLog("Doctors querry return nothing")
		Return
	EndIf

	$doctors = RemoveExcludedDepartments($doctors)
	NormalizeTimesInArray($doctors, 4, 6)

	Local $sqlBusy = "Select Sch.DCode, Ss.DepNum, Sch.BHour, Sch.BMin, Sch.FHour, Sch.FMin " & _
			"From Schedule Sch " & _
			"Join (Select * From DoctShedule Where WDate = 'today' And Shinterv Is Not Null) Ss On Ss.DCode = Sch.DCode " & _
			"Where WorkDate = 'today' " & _
			"And (PCode Is Not Null Or TmStatus > 0) " & _
			"Order By 1, 3 "

	Local $busyCells = ExecuteSQL($sqlBusy)
	NormalizeTimesInArray($busyCells, 2, 4)

	Local $resultArray[0][2]
	Local $maxTimeToShow = _DateAdd('n', $minutesToShow, $now)

	For $i = 0 To UBound($doctors, $UBOUND_ROWS) - 1
		Local $name = $doctors[$i][0]
		If StringInStr($name, "(") Then $name = StringLeft($name, StringInStr($name, "(") - 1)
		If StringRight($name, 1) = " " Then $name = StringLeft($name, StringLen($name) - 1)

		Local $docId = $doctors[$i][1]
		Local $departmentName = $doctors[$i][2]
		Local $departmentId = $doctors[$i][3]
		Local $docIntervalStart = $doctors[$i][4]
		Local $docIntervalEnd = $doctors[$i][5]
		Local $interval = $doctors[$i][6]
		Local $freeTimeIntervals[0]

		;------------------------------------------------------
		;
		; CHECK IF DOC WORKING INTERVAL OUTSIDE CURRENT PERIOD
		;
		;------------------------------------------------------
		;	DOC INTEVAL|
		;			   		*NOW
		;------------------------------------------------------
		If _DateDiff('n', $now, $docIntervalEnd) < 0 Then
;~ 			ToLog("---Skipping, time hours is over")
			ContinueLoop
		EndIf

		;-------------------------------
		;			     |DOC INTERVAL
		;	MAX TIME*
		;-------------------------------
		If (_DateDiff('n', $maxTimeToShow, $docIntervalStart)) >= 0 Then
;~ 			ToLog("---Skipping, time hours is not started yet")
			ContinueLoop
		EndIf

		Local $docBusyArray = ResultFromSearchArrayById($busyCells, $docId, $departmentId, 0, 1)
		If IsArray($docBusyArray) Then _ArraySort($docBusyArray, Default, Default, Default, 2)

		While True
			Local $currentIntervalEnd = _DateAdd('n', $interval, $docIntervalStart)
			Local $currentInterval = $interval

			If IsArray($docBusyArray) Then
				Local $toSkip = False

				;---------------------------------------------
				;
				; CHECK IF FREE CELL INTERSECTS WITH BUSY CELL
				;
				;---------------------------------------------
				For $x = 0 To UBound($docBusyArray, $UBOUND_ROWS) - 1

					Local $busyStart = $docBusyArray[$x][2]
					Local $busyEnd = $docBusyArray[$x][3]

					;-------------------------------
					;			|BUSY CELL|
					;			   *FREE CELL|
					;-------------------------------
					If _DateDiff('n', $busyStart, $docIntervalStart) >= 0 And _
							_DateDiff('n', $busyEnd, $docIntervalStart) < 0 Then
;~ 						ToLog("interval start time inside busy cell")
						$toSkip = True

					;-------------------------------
					;			|BUSY CELL|
					;		|FREE CELL*
					;
					;-------------------------------
					;			|FREE CELL|
					;			   *busy|
					;-------------------------------
					ElseIf  _DateDiff('n', $busyStart, $currentIntervalEnd) > 0 And _
							_DateDiff('n', $busyEnd, $currentIntervalEnd) <= 0 Or _
							_DateDiff('n', $docIntervalStart, $busyStart) >= 0 And _
							_DateDiff('n', $currentIntervalEnd, $busyStart) < 0 Then
						Local $diff = _DateDiff('n', $docIntervalStart, $busyStart)
						If $diff >= $interval * $percentageThatFreeCellMustHaveFromInterval / 100 Then
							$currentIntervalEnd = $busyStart
							$currentInterval = $diff
						Else
;~ 						ToLog("interval end time iside busy cell")
							$toSkip = True
						EndIf
					EndIf

					If $toSkip Then
						$docIntervalStart = $busyEnd
						$toSkip = True
						ExitLoop ; Try next cell
					EndIf
				Next

				If $toSkip Then ContinueLoop ; Try next cell
			EndIf

			;------------------------------------------
			;
			; CHECK IF FREE CELL OUTSIDE CURRENT PERIOD
			;
			;------------------------------------------
			;				  |NOW
			;			 *FREE CELL|
			;------------------------------------------
			If _DateDiff('n', $docIntervalStart, $now) > 0 Then
;~ 				ToLog("---current interval in past, skipping")
				Local $diff = _DateDiff('n', $now, $currentIntervalEnd)
				If $diff >= $interval * $percentageThatFreeCellMustHaveFromInterval / 100 Then
					$docIntervalStart = $now
					$currentInterval = $diff
				Else
					$docIntervalStart = $currentIntervalEnd
					ContinueLoop
				EndIf
			EndIf

			;-------------------------------
			;			MAX TIME|
			;			   |FREE CELL*
			;-------------------------------
			If _DateDiff('n', $currentIntervalEnd, $maxTimeToShow) < 0 Then
;~ 				ToLog("---current interval in future, exiting")
				Local $diff = _DateDiff('n', $docIntervalStart, $maxTimeToShow)
				If $diff >= $interval * $percentageThatFreeCellMustHaveFromInterval / 100 Then
					$currentInterval = $diff
				Else
					ExitLoop
				EndIf
			EndIf

			;----------------------------------------------
			;
			; CHECK IF FREE CELL OUTSIDE DOC WORKING PERIOD
			;
			;----------------------------------------------
			;		DOC INTERVAL|
			;			   	|FREE CELL*
			;----------------------------------------------
			If _DateDiff('n', $currentIntervalEnd, $docIntervalEnd) < 0 Then
;~ 				ToLog("---doc time hours is over, exiting")
				Local $diff = _DateDiff('n', $docIntervalStart, $docIntervalEnd)
				If $diff >= $interval * $percentageThatFreeCellMustHaveFromInterval / 100 Then
					$currentInterval = $diff
				Else
					ExitLoop
				EndIf
			EndIf

			Local $dateTime
			Local $tmpTime
			_DateTimeSplit($docIntervalStart, $dateTime, $tmpTime)

			Local $tmpHour = $tmpTime[1] >= 10 ? $tmpTime[1] : "0" & $tmpTime[1]
			Local $tmpMinute = $tmpTime[2] >= 10 ? $tmpTime[2] : "0" & $tmpTime[2]
			Local $toArray = $tmpHour & ":" & $tmpMinute & ";" & $docIntervalStart & ";" & $currentInterval
			_ArrayAdd($freeTimeIntervals, $toArray)
			$docIntervalStart = $currentIntervalEnd
		WEnd

		If Not UBound($freeTimeIntervals) Then ContinueLoop

		If $showOnlyDepartments Then $name = $departmentName

		_ArrayTranspose($freeTimeIntervals)
		_ArrayColInsert($freeTimeIntervals, 0)

		$freeTimeIntervals[0][0] = $name

		Local $alreadyPresent = _ArraySearch($resultArray, $departmentName)
		If $alreadyPresent < 0 Then
			$alreadyPresent = UBound($resultArray, $UBOUND_ROWS)
			_ArrayAdd($resultArray, $departmentName)
			Local $tmpArray[0]
			$resultArray[$alreadyPresent][1] = $tmpArray
		EndIf

		_ArrayAdd($resultArray[$alreadyPresent][1], _ArrayToString($freeTimeIntervals, "@"))
	Next

	If $showOnlyDepartments Then
		For $i = 0 To UBound($resultArray, $UBOUND_ROWS) - 1
			Local $deptContent = $resultArray[$i][1]
			If UBound($deptContent, $UBOUND_ROWS) < 2 Then ContinueLoop

			Local $deptTotalArray[0]
			Local $deptResultArray[1]
			Local $deptName = $resultArray[$i][0]
			Local $deptContent = $resultArray[$i][1]

			For $x = 0 To UBound($deptContent, $UBOUND_ROWS) - 1
				Local $currentRow = StringSplit($deptContent[$x], "@", $STR_NOCOUNT)
				_ArrayDelete($currentRow, 0)
				_ArrayConcatenate($deptTotalArray, $currentRow)
			Next

			_ArraySort($deptTotalArray)

			Local $prevInterval = $deptTotalArray[0]
			$deptResultArray[0] = $prevInterval
			$prevInterval = StringSplit($prevInterval, ";", $STR_NOCOUNT)
			Local $prevTimeEnd = _DateAdd('n', $prevInterval[2], $prevInterval[1])
			ToLog("prev: " & $prevTimeEnd)

			For $x = 1 To UBound($deptTotalArray) - 1
				Local $curTime = StringSplit($deptTotalArray[$x], ";", $STR_NOCOUNT)
				If _DateDiff('n', $prevTimeEnd, $curTime[1]) < 0 Then ContinueLoop
				_ArrayAdd($deptResultArray, $deptTotalArray[$x])
				$prevTimeEnd = _DateAdd('n', $curTime[2], $curTime[1])
			Next

			_ArrayTranspose($deptResultArray)
			_ArrayColInsert($deptResultArray, 0)
			$deptResultArray[0][0] = $deptName
			Local $tempArray[0]
			$resultArray[$i][1] = $tempArray
			_ArrayAdd($resultArray[$i][1], _ArrayToString($deptResultArray, "@"))
		Next
	EndIf

	If UBound($resultArray, $UBOUND_ROWS) = 0 Then
		ToLog("============ NOTHING FOUND ==========")
		Return
	EndIf

	Return $resultArray
EndFunc


Func ResultFromSearchArrayById($initialArray, $docIdToSearch, $docDepartmentIdToSearch, $columnToSearchDocId, $columnToSearchDocDepartmentId)
	Local $array[0][UBound($initialArray, $UBOUND_COLUMNS)]
	If Not IsArray($initialArray) Then Return $array

	For $i = 0 To UBound($initialArray, $UBOUND_ROWS) - 1
		If $initialArray[$i][$columnToSearchDocId] <> $docIdToSearch Then ContinueLoop
		If $initialArray[$i][$columnToSearchDocDepartmentId] <> $docDepartmentIdToSearch Then ContinueLoop

		Local $tempArray = _ArrayExtract($initialArray, $i, $i, 0, UBound($initialArray, $UBOUND_COLUMNS) - 1)
		_ArrayAdd($array, $tempArray)
	Next

	Return $array
EndFunc


Func RemoveExcludedDepartments($initialArray)
	If Not IsArray($excludedDepartments) Then Return $initialArray

	Local $array[0][UBound($initialArray, $UBOUND_COLUMNS)]

	For $i = 0 To UBound($initialArray, $UBOUND_ROWS) - 1
		If _ArraySearch($excludedDepartments, $initialArray[$i][2]) > -1 Then  ContinueLoop
		Local $tempArray = _ArrayExtract($initialArray, $i, $i, 0, UBound($initialArray, $UBOUND_COLUMNS) - 1)
		_ArrayAdd($array, $tempArray)
	Next

	Return $array
EndFunc


Func NormalizeTimesInArray(ByRef $array, $firstHourColumn, $secondHourColumn)
	If Not IsArray($array) Then Return

	If $secondHourColumn < $firstHourColumn Then
		Local $temp = $firstHourColumn
		$firstHourColumn = $secondHourColumn
		$secondHourColumn = $temp
	EndIf

	For $i = 0 To UBound($array, $UBOUND_ROWS) - 1
		$array[$i][$firstHourColumn] = GetFullDate($array[$i][$firstHourColumn], $array[$i][$firstHourColumn + 1])
		$array[$i][$secondHourColumn] = GetFullDate($array[$i][$secondHourColumn], $array[$i][$secondHourColumn + 1])
	Next

	_ArrayColDelete($array, $secondHourColumn + 1)
	_ArrayColDelete($array, $firstHourColumn + 1)
	_ArraySort($array, Default, Default, Default, 0)
EndFunc


Func GetFullDate($hour, $minute)
	If $hour < 10 Then $hour = "0" & $hour
	If $minute < 10 Then $minute = "0" & $minute
	Local $today = @YEAR & "/" & @MON & "/" & @MDAY
	Return $today & " " & $hour & ":" & $minute & ":00"
EndFunc


Func ExecuteSQL($sql)
;~ 	Local $sqlBD = "DRIVER=Firebird/InterBase(r) driver; UID=; PWD=; DBNAME=;"
	Local $sqlBD = "DRIVER=Firebird/InterBase(r) driver; UID=; PWD=; DBNAME=;"
;~  Local $sqlBD = "DRIVER=Firebird/InterBase(r) driver; UID=; PWD=; DBNAME=;"
	Local $adoConnection = ObjCreate("ADODB.Connection")
	Local $adoRecords = ObjCreate("ADODB.Recordset")

	$adoConnection.Open($sqlBD)
	$adoRecords.CursorType = 2
	$adoRecords.LockType = 3 ;3 - locks 1 - readonly

	Local $result = ""
	$adoRecords.Open($sql, $adoConnection)
	Local $result = $adoRecords.GetRows

	If $adoRecords.EOF = True And $adoRecords.BOF = True Then
		ConsoleWrite("SQL EOF OR BOF" & @CRLF)
		Return ""
	EndIf

	$adoRecords.Close
	$adoRecords = 0
	$adoConnection.Close
	$adoConnection = 0

	Return $result
EndFunc


Func ToLog($message)
	$message &= @CRLF
	$messageToSend &= $message
	ConsoleWrite($message)
	_FileWriteLog($logFile, $message)
EndFunc


#CS Func SendEmail()
; 	If Not $send_email Then
; 		FileClose($logFile)
; 		Return
; 	EndIf
;
; 	ToLog(@CRLF & "---Sending email---")
; 	If _INetSmtpMailCom($server, "Copy MARS data", $login, $to, _
; 			$current_pc_name & ": error(s) occurred", _
; 			$messageToSend, "", "", "", $login, $password) <> 0 Then
;
; 		_INetSmtpMailCom($server_backup, "Copy MARS data", $login_backup, $to_backup, _
; 				$current_pc_name & ": error(s) occurred", _
; 				$messageToSend, "", "", "", $login_backup, $password_backup)
; 	EndIf
;
; 	FileClose($logFile)
; EndFunc   ;==>SendEmail
 #CE


Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, _
		$s_Subject = "", $as_Body = "", $s_AttachFiles = "", $s_CcAddress = "", _
		$s_BccAddress = "", $s_Username = "", $s_Password = "", $IPPort = 25, $ssl = 0)

	Local $objEmail = ObjCreate("CDO.Message")
	Local $i_Error = 0
	Local $i_Error_desciption = ""

	$objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
	$objEmail.To = $s_ToAddress

	If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress
	If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress

	$objEmail.Subject = $s_Subject

	If StringInStr($as_Body, "<") And StringInStr($as_Body, ">") Then
		$objEmail.HTMLBody = $as_Body
	Else
		$objEmail.Textbody = $as_Body & @CRLF
	EndIf

	If $s_AttachFiles <> "" Then
		Local $S_Files2Attach = StringSplit($s_AttachFiles, ";")
		For $x = 1 To $S_Files2Attach[0] - 1
			$S_Files2Attach[$x] = _PathFull($S_Files2Attach[$x])
			If FileExists($S_Files2Attach[$x]) Then
				$objEmail.AddAttachment($S_Files2Attach[$x])
			Else
				$i_Error_desciption = $i_Error_desciption & @LF & 'File not found to attach: ' & $S_Files2Attach[$x]
				SetError(1)
				Return 0
			EndIf
		Next
	EndIf

	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort

	If $s_Username <> "" Then
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
	EndIf

	If $ssl Then
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
	EndIf

	$objEmail.Configuration.Fields.Update
	$objEmail.Send

	If @error Then
		SetError(2)
	EndIf

	Return @error
EndFunc


Func HandleComError()
;~ 	  ConsoleWrite("error.description: " & @TAB & $oMyError.description  & @CRLF & _
;~ 				  "err.windescription:"   & @TAB & $oMyError.windescription & @CRLF & _
;~ 				  "err.number is: "       & @TAB & hex($oMyError.number,8)  & @CRLF & _
;~ 				  "err.lastdllerror is: "   & @TAB & $oMyError.lastdllerror   & @CRLF & _
;~ 				  "err.scriptline is: "   & @TAB & $oMyError.scriptline   & @CRLF & _
;~ 				  "err.source is: "       & @TAB & $oMyError.source       & @CRLF & _
;~ 				  "err.helpfile is: "       & @TAB & $oMyError.helpfile     & @CRLF & _
;~ 				  "err.helpcontext is: " & @TAB & $oMyError.helpcontext & @CRLF)
EndFunc