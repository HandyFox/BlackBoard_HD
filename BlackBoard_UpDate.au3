#cs ----------------------------------------------------------------------------

	AutoIt Version: 3.3.14.5
	Author:         @FoxSoft


	|/\|/\|/\|/\|/\|/\|/\|/\|/\|/\|/\|/\|/\|

	_      _     _   _
	By |_ |_| |_ |  | \ | | |\ |
	_|| | |_ |_ |_/ |_| | \|
	(@CodeFox)

	|/\|/\|/\|/\|/\|/\|/\|/\|/\|/\|/\|/\|/\|

	BLACKBOARD HD PostBeta v1.0
	"A High Definition Blackboard that allows you to easily manage your students grades"

	Features:
	Generates printable reportcard styled tables of averages
	Allows you to enter percentages for Tests and Quizes

	Future Features:
	Generate Graphs......


	And now a endeavor to write BLACKBOARD HD in legible Autoit

#ce ----------------------------------------------------------------------------


#cs

	To Do:
	Add Letter Grades to Report -- Collect Data From New Profile -FIXED
	Make The popup menu on the student list in CREATE PROFILE work - FIXED

	Make a Profile Cleaner Function -- This Could Get Tough
	Finish Help File-DONE for now ;) x)

	Fix:
	Open Profile - FIXED (To My Knowledge - I'll Keep My Eyes Open)
	Bug Found When New Profile Created BlackBoard Has to be restarted to enter grades. blackboard fails to remember new profile on restart - FIXED
	Graph Generation - Rewrite to Autoit - FIXED
	Bug when changing marking periods - FIXED
	_Read_Students_Subjects($Prev_Student_Selection, GUICtrlRead($Current_Subject), $i_Student_Subject_Count, $hGrade_View, $sCurrent_MP)
	_Read_Students_Subjects($Prev_Student_Selection, GUICtrlRead($Current_Subject), ^ ERROR

#ce

;Files to include
#pragma compile(AutoItExecuteAllowed, true)
#include <Inet.au3>
#include <GuiButton.au3>
#include <GUIConstants.au3>
#include <ColorConstants.au3>
#include <GuiEdit.au3>
#include <GUIConstantsEx.au3>
#include <Array.au3>
#include <File.au3>
#include <Date.au3>
#include <GuiImageList.au3>
#include <GuiListView.au3>
#include <WindowsConstants.au3>
#include <ListViewConstants.au3>
#include <GuiMenu.au3>
#include <MsgBoxConstants.au3>
#include "ABC_UDF.au3"
#include 'MPDF_UDF.au3'
#include <GUIToolTip.au3>
#include <GDIPlus.au3>
#include <Zip.au3>
#include <ABC_HTML_UDF.au3>
#include <ScreenCapture.au3>
#include "SmtpMailer.au3"



;++++++++++++++++++++++++++++++++++++
;End of files to include
;++++++++++++++++++++++++++++++++++++

;============================================================================
;Initialize Global Variables
;
;I can't get over my habit of initializing Local variables on the fly
;Am I allowed to blame Just Basic?????
;=============================================================================

Global Enum $e_idOpen = 1000
Global $e_idSave
Global $e_idInfo
Global $hParentWin
Global $aStudent[200]
Global $iStudentCount
Global $aSubject[200]
Global $iSubjectCount
Global $sYearName
Global $sCurrentProfile
Global $sData_Save_Path
Global $sCurrent_MP
Global $Student_Selection
Global $Prev_Student_Selection
Global $iQuiz_Weight
Global $iTest_Weight
Global $sCurrentProfile
Global $BlackBoard_Main_GUI_Handle
Global $BlackBoard_ClipBoard

Global $s_Profile_Pic_Done
Global $hgdraw
Global $hGraphic
Global $hBitmap, $hPath1, $hPen, $h_Profile_Pic_GUI
Global $iScale = 64
Global $hGraph[2]

Global Enum $e_idSet_Number = 1000, $e_idSet_Letter, $e_Report_File_Menu_Print = 2000
Global $sMenuVal
Global $Temp_BH_Dir

Global $StartUpBool = 1
Local $i_Student_Subject_Count
Local $SetFakeClick = 0
Local $aPlugin_Handle[200]
Local $iPlugin

Global $S_Html_Graph_JS
Global $htmlwritedata
Global $Temp_Html_var
;++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;End Initialize Variables
;++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


;================================================
;A Few Startup Routines
;Check for temporary folders etc.
;================================================
$Temp_BH_Dir = @AppDataDir & "\CodeFoxTech\Blackboard_HD"
If DirGetSize($Temp_BH_Dir) = "-1" Then DirCreate($Temp_BH_Dir)
$sCurrentProfile = IniRead(@AppDataDir & "\CodeFoxTech\Blackboard_HD\Recents.ini", "General", "MostRecent", "")

If DirGetSize(@AppDataDir & "\CodeFoxTech\Blackboard_HD\bmp") = "-1" Then DirCreate(@AppDataDir & "\CodeFoxTech\Blackboard_HD\bmp")
FileInstall("make.exe", $Temp_BH_Dir & "\make.exe")
FileInstall("Send.exe", $Temp_BH_Dir & "\Send.exe")

If FileExists($sCurrentProfile) Then
	$sYearName = $sCurrentProfile
	Local $aPaths = StringSplit($sYearName, "\")
	Local $sDir = StringTrimRight($sYearName, StringLen($aPaths[$aPaths[0]]))
	$sData_Save_Path = $sDir & StringTrimRight($aPaths[$aPaths[0]], 3) & "_Data"
Else
	$sCurrentProfile = ""
EndIf

$sCurrent_MP = IniRead($sYearName, "Information", "MarkingPeriod", "Marking Period 1")
$iTest_Weight = IniRead($sYearName, "Information", "TestWeight", "33")
$iQuiz_Weight = IniRead($sYearName, "Information", "QuizWeight", "10")

;============================================
;End Startup
;============================================



;+++++++++++++++++++++++++++++++++++++++++++++++
;GUI Creation
;
;I'll Attempt to keep it legible!!!
;+++++++++++++++++++++++++++++++++++++++++++++++
Local $aProfile_Name = StringSplit($sYearName, "\")
$BlackBoard_Main_GUI_Handle = GUICreate("BlackBoard HD - " & $aProfile_Name[$aProfile_Name[0]], 952, 590, 193, 115, BitOR($WS_MAXIMIZEBOX, $WS_MINIMIZEBOX, $WS_SIZEBOX, $WS_THICKFRAME, $WS_SYSMENU, $WS_CAPTION, $WS_OVERLAPPEDWINDOW, $WS_TILEDWINDOW, $WS_POPUPWINDOW, $WS_GROUP, $WS_TABSTOP, $WS_BORDER, $WS_CLIPSIBLINGS))


$StudentList = GUICtrlCreateListView("Students                ", 1, 97, 249, 473)

$Open_Profile = GUICtrlCreateIcon("open.ico", "", 25, 20)
$Open = GUICtrlCreateLabel("Open Profile", 15, 58, 62, 17)

$Generate_report = GUICtrlCreateIcon("generatereport.ico", "", 115, 20)
$create_report_label = GUICtrlCreateLabel("Generate Report", 91, 58, 79, 17)

$Create_Profile = GUICtrlCreateIcon("profile.ico", "", 209, 20)
$save = GUICtrlCreateLabel("Profile Manager", 185, 58, 100, 17)

$Generate_Graph = GUICtrlCreateIcon("graph2.ico", "", 299, 20)
$Graph_Label = GUICtrlCreateLabel("Generate Graph", 275, 58, 100, 17)

$hGrade_View = GUICtrlCreateListView("Grade                      |Subject              |WeekDay                   |Date                  |Calculation Type              ", 264, 169, 665, 353)
$iListContext = GUICtrlCreateContextMenu($hGrade_View)
$iGrade_Edit = GUICtrlCreateMenuItem("Edit", $iListContext)
$iGrade_Del = GUICtrlCreateMenuItem("Delete", $iListContext)
$iContext_Info = GUICtrlCreateMenu("Extended Info", $iListContext, 1)
Global $iMenuGradeCount = GUICtrlCreateMenuItem("Grade Count =", $iContext_Info)
Global $iMenuDailyWorkCount = GUICtrlCreateMenuItem("Daily Work Count =", $iContext_Info)
Global $iMenuQuizCount = GUICtrlCreateMenuItem("Quiz Count = ", $iContext_Info)
Global $iMenuTestCount = GUICtrlCreateMenuItem("Test Count =", $iContext_Info)
Global $iMenuUnroundedAvg = GUICtrlCreateMenuItem("Unrounded Avg =", $iContext_Info)

$Main_Toolbar_Divide = GUICtrlCreateGraphic(5, 87, 942, 1)

$hAdd_Grade = GUICtrlCreateButton("Add Grade", 725, 537, 114, 25, 0)
;   $hClear_Grade = GUICtrlCreateButton("Delete Grade", 832, 537, 89, 25, 0)
$Current_Subject = GUICtrlCreateCombo("", 408, 132, 110, 25, BitOR($CBS_DROPDOWNLIST, $CBS_AUTOHSCROLL))
$Number_Date = GUICtrlCreateDate("", 680, 132, 110, 21, $WS_TABSTOP)
$Week_Day = GUICtrlCreateCombo("", 543, 132, 110, 21, BitOR($CBS_DROPDOWNLIST, $CBS_AUTOHSCROLL))
$Input_Grade = GUICtrlCreateInput("", 279, 132, 110, 21, BitOR($ES_AUTOHSCROLL, $ES_NUMBER))
$Grade_Info = GUICtrlCreateGroup("Grade Record Info", 256, 97, 681, 65)
$Grade_Weight = GUICtrlCreateCombo("", 807, 130, 110, 21, BitOR($CBS_DROPDOWNLIST, $CBS_AUTOHSCROLL))

$label = GUICtrlCreateLabel("Grade", 304, 113, 33, 17)
$Label3 = GUICtrlCreateLabel("Subject", 440, 111, 40, 17)
$Label4 = GUICtrlCreateLabel("Day of Week", 569, 108, 67, 17)
$Label5 = GUICtrlCreateLabel("Date", 721, 108, 27, 17)
$Label6 = GUICtrlCreateLabel("Grade Type", 830, 109, 93, 17)
$Grade_Info_Group = GUICtrlCreateGroup("Grade Record Info", 256, 97, 681, 65)

$Current_Student_Group = GUICtrlCreateGroup("", 280, 527, 195, 41)
$Current_Student_Average = GUICtrlCreateLabel("", 320, 545, 43, 17)

$Current_MP_Group = GUICtrlCreateGroup("Current Marking Period", 512, 523, 161, 47)
$Marking_Period_Combo = GUICtrlCreateCombo("", 522, 538, 137, 25, BitOR($CBS_DROPDOWNLIST, $CBS_AUTOHSCROLL))


$Main_File_Menu = GUICtrlCreateMenu("&File")
$Main_File_Menu_Open = GUICtrlCreateMenuItem("Open", $Main_File_Menu)
$Main_File_Menu_Generate = GUICtrlCreateMenuItem("Generate Report", $Main_File_Menu)
$Main_File_Menu_PM = GUICtrlCreateMenuItem("Profile Manager", $Main_File_Menu)

$Main_Plugin_Menu = GUICtrlCreateMenu("&Plugins")
#cs
$aPluginList = IniReadSection($Temp_BH_Dir & "\PluginData.ini", "Plugins")
If Not @error Then
	For $iPlugin = 1 To $aPluginList[0][0]
		$hPlug_Handle = GUICtrlCreateMenuItem($aPluginList[$iPlugin][0], $Main_Plugin_Menu)
		$aPlugin_Handle[$iPlugin] = $hPlug_Handle
	Next
EndIf
#ce
$Main_Plugin_Menu_Plugin_Manager = GUICtrlCreateMenuItem("Plugin Manager", $Main_Plugin_Menu)


$Main_Help_Menu = GUICtrlCreateMenu("&Help")
$Main_Help_Menu_Help = GUICtrlCreateMenuItem("Help", $Main_Help_Menu)
$Main_Help_Menu_Help_In_Browser = GUICtrlCreateMenuItem("Launch Help In Browser", $Main_Help_Menu)
$Main_Help_Menu_About = GUICtrlCreateMenuItem("About", $Main_Help_Menu)
$Main_Help_Menu_Send_Bug = GUICtrlCreateMenuItem("Send Bug Report", $Main_Help_Menu)

;Set Opening Data For Gui Controls
_GUICtrlListView_SetColumnWidth ($StudentList, 0, 245)
GUICtrlSetBkColor($Main_Toolbar_Divide, 0x009C9C9C)
GUICtrlSetData($Marking_Period_Combo, "Marking Period 1|Marking Period 2|Marking Period 3|Marking Period 4|Marking Period 5|Marking Period 6|Marking Period 7|Marking Period 8|Marking Period 9|Marking Period 10", $sCurrent_MP)
GUICtrlSetFont($Marking_Period_Combo, 10, 400, 0, "")
GUICtrlSetData($Grade_Weight, "Daily Work|Test|Quiz", "Daily Work")
GUICtrlSetData($Week_Day, "Sunday|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday", _CurrentDay())
GUICtrlSetData($Number_Date, _NowDate())
GUICtrlSetState($hAdd_Grade, $GUI_DEFBUTTON)

GUICtrlSetCursor($Generate_Graph, 0)
GUICtrlSetCursor($Generate_report, 0)
GUICtrlSetCursor($Create_Profile, 0)
GUICtrlSetCursor($Open_Profile, 0)

;GUICtrlSetCursor($StudentList, 0)


GUISetState(@SW_SHOW)

While 1    ;Main Gui Event Loop
	$nMsg = GUIGetMsg()
	If $SetFakeClick = 1 Then
		$nMsg = $Current_Subject
		$SetFakeClick = 0
	EndIf

#cs
	For $i = 1 To $iPlugin
		If String($nMsg) = $aPlugin_Handle[$i] Then
			Run(@ScriptFullPath & ' /AutoIt3ExecuteScript "' & IniRead($Temp_BH_Dir & "\PluginData.ini", "Plugins", $aPluginList[$i][0], "") & '"')
		EndIf
	Next
#ce
	Switch $nMsg

		;Case Gui Closed
		Case $GUI_EVENT_CLOSE
			If $sCurrentProfile > "" Then
				IniWrite(@AppDataDir & "\CodeFoxTech\Blackboard_HD\Recents.ini", "General", "MostRecent", $sCurrentProfile)
				IniWrite($sCurrentProfile, "Information", "MarkingPeriod", $sCurrent_MP)
			EndIf
			GUIDelete($BlackBoard_Main_GUI_Handle)
			ExitLoop ;=> End Case Gui Closed

		Case $Open_Profile, $Main_File_Menu_Open
			$StartUpBool = _Open_Profile()

		Case $Main_Plugin_Menu_Plugin_Manager
			Plugin_Browser()

		Case $Generate_report, $Main_File_Menu_Generate
			_Generate_Report()


		Case $Create_Profile, $Main_File_Menu_PM
			_createProfile()
			_GUICtrlListView_DeleteAllItems ($StudentList)
			$StartUpBool = 1

		Case $Generate_Graph
			_Generate_Graph()


		Case $hGrade_View

		Case $Main_Help_Menu_Help
			_Help()

		Case $Main_Help_Menu_Help_In_Browser
			ShellExecute(@ScriptDir & "\Help\Introduction.html")

		Case $Main_Help_Menu_About
			GUICtrlSetData($iMenuGradeCount, "Hello")
			_About()

		Case $Main_Help_Menu_Send_Bug
			_Send_Bug($sYearName)


		Case $iGrade_Edit
			_Edit_Grade($hGrade_View, $Prev_Student_Selection, GUICtrlRead($Current_Subject), $i_Student_Subject_Count, $sCurrent_MP, $Current_Student_Average)


		Case $iGrade_Del
			If GUICtrlRead(GUICtrlRead($hGrade_View)) > "" Then
				If MsgBox(36, "Please Confirm", "Are you sure you want to delete this grade?") = 6 Then
					GUICtrlDelete(GUICtrlRead($hGrade_View))
					_Save_Grade_List($hGrade_View, $Prev_Student_Selection, GUICtrlRead($Current_Subject), $i_Student_Subject_Count, $sCurrent_MP)
					GUICtrlSetData($Current_Student_Average, _Get_Average($sCurrent_MP, $Prev_Student_Selection, GUICtrlRead($Current_Subject)) & "%")
				EndIf
			EndIf


		Case $hAdd_Grade
			If GUICtrlRead($Input_Grade) > 100 Then
			Else
				_Add_Grade($Current_Subject, $hGrade_View, $Week_Day, $Input_Grade, $Grade_Weight, $Number_Date, $Student_Selection, $sCurrent_MP)
				GUICtrlSetData($Current_Student_Average, _Get_Average($sCurrent_MP, $Prev_Student_Selection, GUICtrlRead($Current_Subject)) & "%")
				GUICtrlSetData($Input_Grade, "")
			EndIf

		Case $Current_Subject
			GUICtrlSetData($Current_Student_Group, $Prev_Student_Selection & "'s " & GUICtrlRead($Current_Subject) & " Average")
			GUICtrlSetData($Input_Grade, "")
			If $Prev_Student_Selection > "" Then
				$i_Student_Subject_Count = IniRead($sYearName, $sCurrent_MP & $Prev_Student_Selection & "GradeCounts", GUICtrlRead($Current_Subject), "")
				_GUICtrlListView_DeleteAllItems ($hGrade_View)
				_Read_Students_Subjects($Prev_Student_Selection, GUICtrlRead($Current_Subject), $i_Student_Subject_Count, $hGrade_View, $sCurrent_MP)
				GUICtrlSetData($Current_Student_Average, _Get_Average($sCurrent_MP, $Prev_Student_Selection, GUICtrlRead($Current_Subject)) & "%")
			EndIf


		Case $Grade_Weight


		Case $Marking_Period_Combo
			$sCurrent_MP = GUICtrlRead($Marking_Period_Combo)
			GUICtrlSetData($Current_Subject, "")
			GUICtrlSetData($Current_Subject, IniRead($sYearName, $Student_Selection, "Subjects", ""))
			GUICtrlSetData($hGrade_View, "")
			GUICtrlSetData($Input_Grade, "")
			_GUICtrlListView_DeleteAllItems ($hGrade_View)
			If $Prev_Student_Selection > "" Then
				$i_Student_Subject_Count = IniRead($sYearName, $sCurrent_MP & $Prev_Student_Selection & "GradeCounts", GUICtrlRead($Current_Subject), "")
				_Read_Students_Subjects($Prev_Student_Selection, GUICtrlRead($Current_Subject), $i_Student_Subject_Count, $hGrade_View, $sCurrent_MP)
			EndIf


	EndSwitch


	;+++++++++++++++++++++++++++++++++++++++++++
	;StartUp Routine
	;
	;This Routine Populates the Student ListView
	;
	;I would have placed this in a separate
	;function but it depends on to many non-global
	;variables
	;+++++++++++++++++++++++++++++++++++++++++++

	If $StartUpBool = 1 Then
		$StartUpBool = 0
		_GUICtrlListView_DeleteAllItems ($StudentList)
		Local $sRead = IniRead($sYearName, "Information", "Students", "")
		$sRead = _ABC_StringTrim ($sRead, "|")

		If StringInStr($sRead, "|") > 0 Then
			Local $sTempArray = StringSplit($sRead, "|")
			$i = 1
			Local $sPath = $sTempArray[$i]
			$hImage = _GUIImageList_Create (32, 32)

			For $i = 1 To $sTempArray[0]
				$sImagePath = TempBMP(IniRead($sYearName, $sTempArray[$i], "ProfilePic", ""))
				$indext = _GUIImageList_AddBitmap ($hImage, $sImagePath)
			Next

			_GUICtrlListView_SetImageList ($StudentList, $hImage, 1)

			For $i = 1 To $sTempArray[0]
				_GUICtrlListView_AddItem ($StudentList, $sTempArray[$i], ($i - 1))
			Next

		Else
			If $sRead > "" Then
				$hImage = _GUIImageList_Create (32, 32)
				$sImagePath = TempBMP(IniRead($sYearName, $sRead, "ProfilePic", ""))
				$indext = _GUIImageList_AddBitmap ($hImage, $sImagePath)

				_GUICtrlListView_SetImageList ($StudentList, $hImage, 1)
				_GUICtrlListView_AddItem ($StudentList, $sRead, 0)

			EndIf
		EndIf
	EndIf

	;========================
	;End StartUp Routine
	;=======================

	;+++++++++++++++++++++++++++++++++++++++++++++
	;The Code Block below reads when the user
	;selects a different student then makes the
	;necessary ajustments
	;++++++++++++++++++++++++++++++++++++++++++++


	$Student_Selection = _GUICtrlListView_GetItemTextString ($StudentList, -1)

	If $Student_Selection <> $Prev_Student_Selection Then
		If $Student_Selection > "" Then
			$BlackBoard_ClipBoard = GUICtrlRead($Current_Subject)
			$Prev_Student_Selection = $Student_Selection
			_GUICtrlListView_DeleteAllItems ($hGrade_View)
			GUICtrlSetData($Current_Student_Group, "")
			GUICtrlSetData($Current_Subject, "")
			GUICtrlSetData($Current_Subject, IniRead($sYearName, $Student_Selection, "Subjects", ""), $BlackBoard_ClipBoard)
			$SetFakeClick = 1
		EndIf
	EndIf

	;=> End Student Selection Code Block

WEnd

Func _Open_Profile()
	Local $sFileOpenDialog = FileOpenDialog("Select BlackBoard HD Profile", @MyDocumentsDir & "\", "BlackBoard HD Files(*.bh)", $FD_FILEMUSTEXIST)
	If $sFileOpenDialog = "" Then Return 0

	Local $s_BH_Genuine_Key = IniRead($sFileOpenDialog, "Information", "BlackBoardHDGenuineStamp", "")
	ConsoleWrite($s_BH_Genuine_Key & @CR)
	If $s_BH_Genuine_Key <> "@&$**@^(" Then
		MsgBox(48, "Error!", "Cannot load profile!" & @CRLF & "BlackBoard HD detected that profile is not genuine")
		Return 0
	EndIf

	$sCurrentProfile = $sFileOpenDialog
	$sYearName = $sFileOpenDialog
	Local $aProfile_Name = StringSplit($sYearName, "\")
	WinSetTitle($BlackBoard_Main_GUI_Handle, "", "BlackBoard - " & $aProfile_Name[$aProfile_Name[0]])
	Return 1

EndFunc   ;==>_Open_Profile

Func _createProfile()

	Local $prevSelection
	Local $aParentWin_Pos = WinGetPos($BlackBoard_Main_GUI_Handle, "")
	ConsoleWrite("++++" & $sYearName & @CR)
	Local $aProfile_Name = StringSplit($sYearName, "\")

	$hEditProfile = GUICreate("Profile Manager - " & $aProfile_Name[$aProfile_Name[0]], 498, 464, $aParentWin_Pos[0] + 100, $aParentWin_Pos[1] + 100, -1, -1, $BlackBoard_Main_GUI_Handle)


	$Group4 = GUICtrlCreateGroup("Letter Grade Values", 223, 8, 265, 197)
	$Ctrl_A_Plus = GUICtrlCreateInput("", 285, 44, 33, 21)
	$Ctrl_A = GUICtrlCreateInput("", 333, 44, 33, 21)
	$Ctrl_A_Minus = GUICtrlCreateInput("", 385, 44, 33, 21)
	$letLabel_1 = GUICtrlCreateLabel("A+", 292, 28, 17, 15)
	$letLabel_1 = GUICtrlCreateLabel("A", 340, 28, 17, 15)
	$letLabel_1 = GUICtrlCreateLabel("A-", 392, 28, 17, 15)

	$Ctrl_B_Plus = GUICtrlCreateInput("", 285, 85, 33, 21)
	$Ctrl_B = GUICtrlCreateInput("", 333, 85, 33, 21)
	$Ctrl_B_Minus = GUICtrlCreateInput("", 385, 85, 33, 21)
	$letLabel_2 = GUICtrlCreateLabel("B+", 292, 69, 17, 15)
	$letLabel_2 = GUICtrlCreateLabel("B", 340, 69, 17, 15)
	$letLabel_2 = GUICtrlCreateLabel("B-", 392, 69, 17, 15)

	$Ctrl_C_Plus = GUICtrlCreateInput("", 285, 126, 33, 21)
	$Ctrl_C = GUICtrlCreateInput("", 333, 126, 33, 21)
	$Ctrl_C_Minus = GUICtrlCreateInput("", 385, 126, 33, 21)
	$letLabel_3 = GUICtrlCreateLabel("C+", 292, 110, 17, 15)
	$letLabel_3 = GUICtrlCreateLabel("C", 340, 110, 17, 15)
	$letLabel_3 = GUICtrlCreateLabel("C-", 392, 110, 17, 15)

	$Ctrl_D = GUICtrlCreateInput("", 285, 167, 33, 21)
	$Ctrl_E = GUICtrlCreateInput("", 333, 167, 33, 21)
	$Ctrl_F = GUICtrlCreateInput("", 385, 167, 33, 21)
	$letLabel_4 = GUICtrlCreateLabel("D", 292, 151, 17, 15)
	$letLabel_4 = GUICtrlCreateLabel("E", 340, 151, 17, 15)
	$letLabel_4 = GUICtrlCreateLabel("F", 392, 151, 17, 15)


	$_A_Plus = IniRead($sYearName, "Information", "A+", "100")
	GUICtrlSetData($Ctrl_A_Plus, $_A_Plus)
	$_A = IniRead($sYearName, "Information", "A", "96")
	GUICtrlSetData($Ctrl_A, $_A)
	$_A_Minus = IniRead($sYearName, "Information", "A-", "94")
	GUICtrlSetData($Ctrl_A_Minus, $_A_Minus)
	$_B_Plus = IniRead($sYearName, "Information", "B+", "92")
	GUICtrlSetData($Ctrl_B_Plus, $_B_Plus)
	$_B = IniRead($sYearName, "Information", "B", "88")
	GUICtrlSetData($Ctrl_B, $_B)
	$_B_Minus = IniRead($sYearName, "Information", "B-", "86")
	GUICtrlSetData($Ctrl_B_Minus, $_B_Minus)
	$_C_Plus = IniRead($sYearName, "Information", "C+", "84")
	GUICtrlSetData($Ctrl_C_Plus, $_C_Plus)
	$_C = IniRead($sYearName, "Information", "C", "80")
	GUICtrlSetData($Ctrl_C, $_C)
	$_C_Minus = IniRead($sYearName, "Information", "C-", "76")
	GUICtrlSetData($Ctrl_C_Minus, $_C_Minus)
	$_D = IniRead($sYearName, "Information", "D", "70")
	GUICtrlSetData($Ctrl_D, $_D)
	$_E = IniRead($sYearName, "Information", "E", "63")
	GUICtrlSetData($Ctrl_E, $_E)
	$_F = IniRead($sYearName, "Information", "F", "0")
	GUICtrlSetData($Ctrl_F, $_F)


	$Group1 = GUICtrlCreateGroup("Subject Manager", 224, 208, 265, 249)
	$addSubject = GUICtrlCreateButton("Add Subject", 392, 236, 89, 25, 0)
	$subjectinput = GUICtrlCreateInput("", 240, 236, 129, 21, BitOR($ES_AUTOHSCROLL, $ES_WANTRETURN))
	$SubjectView = GUICtrlCreateListView("Subjects     ", 232, 272, 249, 175)
	$iListContext = GUICtrlCreateContextMenu($SubjectView)
	$iSubEdit = GUICtrlCreateMenuItem("Edit", $iListContext)
	$iSubDel = GUICtrlCreateMenuItem("Delete", $iListContext)
	$Group2 = GUICtrlCreateGroup("Grade Weights", 8, 88, 201, 113)
	$Test = GUICtrlCreateLabel("Test  %", 40, 120, 39, 17)
	$Test_Input = GUICtrlCreateInput("", 83, 115, 113, 21, BitOR($ES_AUTOHSCROLL, $ES_WANTRETURN, $ES_NUMBER))
	$Quiz = GUICtrlCreateLabel("Quiz %", 42, 153, 36, 17)
	$Quiz_Input = GUICtrlCreateInput("", 80, 152, 113, 21, BitOR($ES_AUTOHSCROLL, $ES_WANTRETURN, $ES_NUMBER))
	$Profile_Student_List = GUICtrlCreateListView("Students", 8, 272, 209, 175)
	$iStudentListContext = GUICtrlCreateContextMenu($Profile_Student_List)
	$iSubEditStudent = GUICtrlCreateMenuItem("Edit", $iStudentListContext)
	$iSubDelStudent = GUICtrlCreateMenuItem("Delete", $iStudentListContext)
	$AddStudent = GUICtrlCreateButton("Add Student", 48, 240, 121, 25, 0)
	$Group3 = GUICtrlCreateGroup("Student Manager", 5, 208, 217, 249)

	$iNewYear = GUICtrlCreateIcon(@ScriptDir & "\newfile.ico", "", 35, 16)
	$label_NewYear = GUICtrlCreateLabel("New Year", 24, 57, 50, 25)
	$Save_Exit = GUICtrlCreateIcon(@ScriptDir & "\save_folder.ico", "", 120, 16)
	$label_Save = GUICtrlCreateLabel("Save && Exit", 110, 57, 70, 25)

	_GUICtrlListView_SetColumnWidth ($SubjectView, 0, 245)
	_GUICtrlListView_SetColumnWidth ($Profile_Student_List, 0, 187)


	GUICtrlSetState($subjectinput, 256)
	GUICtrlSetState($addSubject, $GUI_DEFBUTTON)
	GUICtrlSetData($Test_Input, $iTest_Weight)
	GUICtrlSetData($Quiz_Input, $iQuiz_Weight)

	GUICtrlSetCursor($Save_Exit, 0)
	GUICtrlSetCursor($iNewYear, 0)

	GUISetState(@SW_SHOW)

	If $sYearName > "" Then
		_Add_Students($Profile_Student_List)
		_Add_Subjects($SubjectView)
	EndIf

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				GUIDelete($hEditProfile)
				Return

			Case $iSubDelStudent
				$Selected_Student = GUICtrlRead(GUICtrlRead($Profile_Student_List))
				$Selected_Student = StringTrimRight($Selected_Student, 1)
				If MsgBox(36, "Please Confirm", "Are you sure you want to delete '" & $Selected_Student & "' from your student list?") = 6 Then
					ConsoleWrite($Selected_Student & @CR)
					$_Delete_Temp_Subjects = IniRead($sYearName, $Selected_Student, "Subjects", "")
					$_Delete_ProfilePic = IniRead($sYearName, $Selected_Student, "ProfilePic", "")
					$_Delete_Students = IniRead($sYearName, "Information", "Students", "")
					If StringInStr($_Delete_Students, $Selected_Student & "|") > 0 Then
						$_Delete_Students = StringReplace($_Delete_Students, $Selected_Student & "|", "")
					Else
						$_Delete_Students = StringReplace($_Delete_Students, $Selected_Student, "")
					EndIf
					Local $aDelete_Subjects = StringSplit($_Delete_Temp_Subjects, "|")
					For $i = 1 To 10
						IniDelete($sYearName, "Marking Period " & $i & $Selected_Student & "GradeCounts")
						For $isub = 1 To $aDelete_Subjects[0]
							IniDelete($sYearName, "Marking Period " & $i & $Selected_Student & $aDelete_Subjects[$isub])
						Next
					Next
					$_Delete_Students = _ABC_StringTrim ($_Delete_Students, "|")
					IniWrite($sYearName, "Information", "Students", $_Delete_Students)
					IniDelete($sYearName, $Selected_Student)
					FileDelete($_Delete_ProfilePic)
					$StartUpBool = 1
					_GUICtrlListView_DeleteAllItems ($Profile_Student_List)
					_Add_Students($Profile_Student_List)
				EndIf

			Case $iSubEditStudent
				$Selected_Student = GUICtrlRead(GUICtrlRead($Profile_Student_List))
				$Selected_Student = _ABC_StringTrim ($Selected_Student, "|")
				$sRead = IniRead($sYearName, $Selected_Student, "Subjects", "")
				$sRead_PP = IniRead($sYearName, $Selected_Student, "ProfilePic", "")
				_AddStudent($hEditProfile, $sRead, $sRead_PP, $Selected_Student, "Edit")

			Case $Save_Exit
				If GUICtrlRead($Test_Input) > "" Then
					If GUICtrlRead($Quiz_Input) > "" Then
						IniWrite($sYearName, "Information", "TestWeight", GUICtrlRead($Test_Input))
						IniWrite($sYearName, "Information", "QuizWeight", GUICtrlRead($Quiz_Input))
						$iTest_Weight = GUICtrlRead($Test_Input)
						$iQuiz_Weight = GUICtrlRead($Quiz_Input)
					Else
						_GUICtrlEdit_ShowBalloonTip ($Test_Input, "Error", "This Box Must Be Filled Before Saving", $TTI_INFO)
					EndIf
				Else
					_GUICtrlEdit_ShowBalloonTip ($Quiz_Input, "Error", "This Box Must Be Filled Before Saving", $TTI_INFO)
				EndIf


				IniWrite($sYearName, "Information", "A+", GUICtrlRead($Ctrl_A_Plus))
				IniWrite($sYearName, "Information", "A", GUICtrlRead($Ctrl_A))
				IniWrite($sYearName, "Information", "A-", GUICtrlRead($Ctrl_A_Minus))
				IniWrite($sYearName, "Information", "B+", GUICtrlRead($Ctrl_B_Plus))
				IniWrite($sYearName, "Information", "B", GUICtrlRead($Ctrl_B))
				IniWrite($sYearName, "Information", "B-", GUICtrlRead($Ctrl_B_Minus))
				IniWrite($sYearName, "Information", "C+", GUICtrlRead($Ctrl_C_Plus))
				IniWrite($sYearName, "Information", "C", GUICtrlRead($Ctrl_C))
				IniWrite($sYearName, "Information", "C-", GUICtrlRead($Ctrl_C_Minus))
				IniWrite($sYearName, "Information", "D", GUICtrlRead($Ctrl_D))
				IniWrite($sYearName, "Information", "E", GUICtrlRead($Ctrl_E))
				IniWrite($sYearName, "Information", "F", GUICtrlRead($Ctrl_F))
				LetterValClean()
				GUIDelete($hEditProfile)
				Return




			Case $iNewYear
				$sTempYearName = FileSaveDialog("Save File As", "::{450D8FBA-AD25-11D0-98A8-0800361B1103}", "BlackBoard HD Files (*.bh)", $FD_PATHMUSTEXIST)
				If $sTempYearName > "" Then
					_GUICtrlListView_DeleteAllItems ($SubjectView)
					_GUICtrlListView_DeleteAllItems ($Profile_Student_List)
					If StringRight($sTempYearName, 3) <> ".bh" Then
						$sTempYearName = $sTempYearName & ".bh"
					EndIf
					$sYearName = $sTempYearName
					$sCurrentProfile = $sYearName
					Local $aProfile_Name = StringSplit($sYearName, "\")
					WinSetTitle($hEditProfile, "", "Profile Manager - " & $aProfile_Name[$aProfile_Name[0]])

					If Not FileExists($sYearName) Then
						Local $aPaths = StringSplit($sYearName, "\")
						Local $sDir = StringTrimRight($sYearName, StringLen($aPaths[$aPaths[0]]))
						$sData_Save_Path = $sDir & StringTrimRight($aPaths[$aPaths[0]], 3) & "_Data"

						If DirGetSize($sData_Save_Path) = "-1" Then
							DirCreate($sData_Save_Path)
						EndIf
					EndIf
					Local $aProfile_Name = StringSplit($sYearName, "\")
					WinSetTitle($BlackBoard_Main_GUI_Handle, "", "BlackBoard - " & $aProfile_Name[$aProfile_Name[0]])
					IniWrite($sCurrentProfile, "Information", "BlackBoardHDGenuineStamp", "@&$**@^(")
				EndIf

			Case $addSubject
				If $sYearName > "" Then
					If GUICtrlRead($subjectinput) > "" Then
						GUICtrlCreateListViewItem(GUICtrlRead($subjectinput), $SubjectView)
						GUICtrlSetData($subjectinput, "")
						_Save_ListView_Contents($SubjectView)
						_GUICtrlListView_DeleteAllItems ($SubjectView)
						_Add_Subjects($SubjectView)
					EndIf
				EndIf

			Case $iSubEdit
				$name = GUICtrlRead(GUICtrlRead($SubjectView))
				If $name > "" Then
					$name = StringTrimRight($name, 1)
					$htemphandle = GUICtrlRead($SubjectView)
					Local $sAnswer = InputBox("Rename", $name, $name, "", 250, 130, _WinAPI_GetMousePosX (), _WinAPI_GetMousePosY (), 0, $hEditProfile)

					If $sAnswer <> "" Then
						GUICtrlSetData($htemphandle, $sAnswer)
					EndIf
					_Save_ListView_Contents($SubjectView)
				EndIf

			Case $iSubDel
				If GUICtrlRead(GUICtrlRead($SubjectView)) > "" Then
					GUICtrlDelete(GUICtrlRead($SubjectView))
					_Save_ListView_Contents($SubjectView)
				EndIf


			Case $AddStudent
				If $sYearName <> "" Then
					_AddStudent($hEditProfile)
					_GUICtrlListView_DeleteAllItems ($Profile_Student_List)
					_Add_Students($Profile_Student_List)
				EndIf


		EndSwitch
	WEnd

EndFunc   ;==>_createProfile


Func _AddStudent($hEditProfile, $sTemp_SubjectList = "", $sTemp_ProfilePic = "", $sTemp_StudentName = "", $GuiType = "Add")


	Local $sProfilePic = ""
	$aParentWin_Pos = WinGetPos($hEditProfile, "")
	#cs
		$AddStudent = GUICreate("Add Student", 621, 381, $aParentWin_Pos[0] + 100, $aParentWin_Pos[1] + 100, -1, -1, $hEditProfile)
		$Subject_List = GUICtrlCreateListView("Subjects    ", 272, 16, 337, 353)
		$Pic1 = GUICtrlCreatePic(@ScriptDir & "\setpic.bmp", 96, 16, 64, 64, BitOR($SS_NOTIFY, $WS_GROUP))
		$Student_Name = GUICtrlCreateInput("", 32, 104, 161, 21)
		$Student_Grade = GUICtrlCreateInput("", 40, 168, 65, 21)
		$Date1 = GUICtrlCreateDate("", 40, 224, 113, 25, $WS_TABSTOP)
		$Grade = GUICtrlCreateGroup("Grade", 24, 152, 89, 45)
		$Birthday = GUICtrlCreateGroup("Birthday", 24, 208, 145, 45)
		$name = GUICtrlCreateGroup("Name", 24, 88, 185, 49)
		$Student_SAVE = GUICtrlCreateButton("Save && Exit", 40, 288, 145, 25, 0)
		$Student_SAVE_and_new = GUICtrlCreateButton("Save && New", 40, 331, 145, 25, 0)

	#ce

	$AddStudent = GUICreate($GuiType & " Student", 336, 417, $aParentWin_Pos[0] + 100, $aParentWin_Pos[1] + 100, -1, -1, $hEditProfile)
	$Subject_List = GUICtrlCreateListView("Subjects    ", 28, 184, 281, 177)
	$Pic1 = GUICtrlCreatePic(@ScriptDir & "\setpic.bmp", 136, 30, 64, 64, BitOR($SS_NOTIFY, $WS_GROUP))
	$Student_Name = GUICtrlCreateInput("", 88, 144, 161, 21)
	$name = GUICtrlCreateGroup("Name", 76, 125, 185, 49)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	$Student_SAVE = GUICtrlCreateButton("Save & Exit", 12, 379, 145, 25, 0)
	$Student_SAVE_and_new = GUICtrlCreateButton("Save && New", 180, 379, 145, 25, 0)

	GUISetState(@SW_SHOW)
	_GUICtrlListView_SetExtendedListViewStyle ($Subject_List, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_CHECKBOXES))
	_Add_Subjects($Subject_List)

	If $GuiType = "Edit" Then
		GUICtrlSetState($Student_SAVE_and_new, 128)
		$index_temp = _GUICtrlListView_GetItemCount ($Subject_List)
		Local $sTempArray = StringSplit($sTemp_SubjectList, "|")
		$it_Count = 0
		Do

			$aItem = _GUICtrlListView_GetItem ($Subject_List, $it_Count)

			For $i = 1 To $sTempArray[0]
				If $aItem[3] = $sTempArray[$i] Then _GUICtrlListView_SetItemChecked ($Subject_List, $it_Count)
			Next
			$it_Count = $it_Count + 1
		Until $it_Count >= $index_temp

		GUICtrlSetImage($Pic1, TempBMP($sTemp_ProfilePic))
		GUICtrlSetData($Student_Name, $sTemp_StudentName)
		$sProfilePic = $sTemp_ProfilePic
	EndIf

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				GUIDelete($AddStudent)
				Return

			Case $Pic1
				FileDelete($Temp_BH_Dir & "\finalpic.bmp")
				$sPic = _Get_Profile_Pic($AddStudent)
				If $sPic > "" Then
					ConsoleWrite($sPic & @CR)
					RunWait("cmd.exe /c make.exe " & $sPic, $Temp_BH_Dir, @SW_HIDE)
					$sProfilePic = $Temp_BH_Dir & "\finalpic.bmp"
					GUICtrlSetImage($Pic1, $sProfilePic)
				EndIf



			Case $Student_SAVE_and_new
				If GUICtrlRead($Student_Name) > "" Then
					_AddStudent_Save($Subject_List, $Student_Name, $sProfilePic, $GuiType)
					GUICtrlSetData($Student_Name, "")
					GUICtrlSetImage($Pic1, @ScriptDir & "\setpic.bmp")
					$sProfilePic = ""
					_GUICtrlListView_SetItemChecked ($Subject_List, -1, False)
				EndIf



			Case $Student_SAVE
				If GUICtrlRead($Student_Name) > "" Then
					_AddStudent_Save($Subject_List, $Student_Name, $sProfilePic, $GuiType)
					GUIDelete($AddStudent)
					$sProfilePic = ""
					Return
				EndIf


		EndSwitch
	WEnd


EndFunc   ;==>_AddStudent


Func _AddStudent_Save($Subject_List, $Student_Name, $sProfilePic, $Style = "Add")

	$sRead = IniRead($sYearName, "Information", "Students", "")
	If GUICtrlRead($Student_Name) > "" Then
		If StringInStr($sRead, GUICtrlRead($Student_Name)) > 0 And $Style = "Add" Then
			MsgBox($MB_SYSTEMMODAL, "", "The Name you entered appears to already be entered!")
		Else
			Local $it_Count = 0
			Local $sTempSubjects
			$index_temp = _GUICtrlListView_GetItemCount ($Subject_List)

			If $Style = "Add" Then IniWrite($sYearName, "Information", "Students", $sRead & "|" & GUICtrlRead($Student_Name))
			#cs
				$sTemp_GradeList = IniRead($sYearName, "Information", "Grades", "")
				If StringInStr($sTemp_GradeList, GUICtrlRead($Student_Grade)) > 0 Then
				Else
				$sTemp_GradeList = $sTemp_GradeList & "|" & GUICtrlRead($Student_Grade)
				$sTemp_GradeList = _ABC_StringTrim ($sTemp_GradeList, "|")
				IniWrite($sYearName, "Information", "Grades", $sTemp_GradeList)
				EndIf
			#ce
			If $sProfilePic = "" Then
				FileDelete($Temp_BH_Dir & "\finalpic.bmp")
				RunWait("cmd.exe /c make.exe " & GUICtrlRead($Student_Name), $Temp_BH_Dir, @SW_HIDE)
				; FileCopy($Temp_BH_Dir & "\finalpic.bmp", $sData_Save_Path & "\" & GUICtrlRead($Student_Name) & ".db", 1)
				$sProfilePic = File2Binary($Temp_BH_Dir & "\finalpic.bmp")
			Else
				$sProfilePic = File2Binary($sProfilePic)
			EndIf
			IniWrite($sYearName, GUICtrlRead($Student_Name), "ProfilePic", $sProfilePic)

			Do
				$sItem = _GUICtrlListView_GetItemChecked ($Subject_List, $it_Count)
				$aItem = _GUICtrlListView_GetItem ($Subject_List, $it_Count)
				$it_Count = $it_Count + 1
				If $sItem = "True" Then
					If $aItem[3] <> "" Then
						$sTempSubjects = $sTempSubjects & "|" & $aItem[3]
					EndIf
				EndIf
			Until $it_Count >= $index_temp
			$sTempSubjects = _ABC_StringTrim ($sTempSubjects, "|")
			IniWrite($sYearName, GUICtrlRead($Student_Name), "Subjects", $sTempSubjects)
		EndIf
	EndIf
EndFunc   ;==>_AddStudent_Save


Func _Add_Students($Profile_Student_List)

	Local $sRead = IniRead($sYearName, "Information", "Students", "")
	$sRead = _ABC_StringTrim ($sRead, "|")

	If StringInStr($sRead, "|") > 0 Then
		Local $sTempArray = StringSplit($sRead, "|")
		$i = 1
		Local $sPath = $sTempArray[$i]

		For $i = 1 To $sTempArray[0]
			;$sTempArray[$i]
			GUICtrlCreateListViewItem($sTempArray[$i], $Profile_Student_List)

		Next
	Else
		GUICtrlCreateListViewItem($sRead, $Profile_Student_List)
	EndIf
EndFunc   ;==>_Add_Students


Func _Save_ListView_Contents($sHandle)


	Local $it_Count = 0
	Local $sWrite_Subjects = ""
	$index_temp = _GUICtrlListView_GetItemCount ($sHandle)
	Do
		$aItem = _GUICtrlListView_GetItem ($sHandle, $it_Count)
		$it_Count = $it_Count + 1
		ConsoleWrite($aItem[3] & @CR)

		If $aItem[3] <> "" Then
			$sWrite_Subjects = $sWrite_Subjects & "|" & $aItem[3]
		EndIf
	Until $it_Count >= $index_temp
	ConsoleWrite($sWrite_Subjects & @CR)
	ConsoleWrite($sYearName & @CR)
	IniWrite($sYearName, "Information", "Subjects", $sWrite_Subjects)
EndFunc   ;==>_Save_ListView_Contents

Func _Add_Subjects($hHandle)
	Local $sRead = IniRead($sYearName, "Information", "Subjects", "")
	$sRead = _ABC_StringTrim ($sRead, "|")
	If StringInStr($sRead, "|") > 0 Then
		Local $sTempArray = StringSplit($sRead, "|")

		For $i = 1 To $sTempArray[0]
			ConsoleWrite("|" & $sTempArray[$i] & "|" & @CR)
			GUICtrlCreateListViewItem($sTempArray[$i], $hHandle)
		Next
	Else
		If $sRead > "" Then
			GUICtrlCreateListViewItem($sRead, $hHandle)
		EndIf
	EndIf
EndFunc   ;==>_Add_Subjects

Func _Add_Grade($Current_Subject, $hGrade_View, $Week_Day, $Input_Grade, $Grade_Weight, $Number_Date, $sCurrent_Student, $sCurrent_MP)
	Local $iLoopCounter
	If $sCurrent_MP = "" Then Return
	If $sCurrent_Student = "" Then Return
	Local $Subject = GUICtrlRead($Current_Subject)
	If $Subject = "" Then Return
	Local $sWeekDay = GUICtrlRead($Week_Day)
	If $sWeekDay = "" Then Return
	Local $sDate = GUICtrlRead($Number_Date)
	If $sDate = "" Then Return
	Local $sGrade = GUICtrlRead($Input_Grade)
	If StringLen($sGrade) < 2 Then Return
	If Int($sGrade) > 100 Then Return
	Local $sGradeType = GUICtrlRead($Grade_Weight)
	If $sGradeType = "" Then Return

	$i_Student_Subjects = IniRead($sYearName, $sCurrent_MP & $sCurrent_Student & "GradeCounts", $Subject, "")
	$i_Student_Subjects = $i_Student_Subjects + 1
	IniWrite($sYearName, $sCurrent_MP & $sCurrent_Student & "GradeCounts", $Subject, $i_Student_Subjects)
	IniWrite($sYearName, $sCurrent_MP & $sCurrent_Student & $Subject, $i_Student_Subjects, $sGrade & "|" & $Subject & "|" & $sWeekDay & "|" & $sDate & "|" & $sGradeType)

	_GUICtrlListView_DeleteAllItems ($hGrade_View)

	_Read_Students_Subjects($sCurrent_Student, $Subject, $i_Student_Subjects, $hGrade_View, $sCurrent_MP)




EndFunc   ;==>_Add_Grade

Func _Read_Students_Subjects($sStudent, $sSubject, $i_Student_Subject_Count, $handle, $sCurrent_MP)

	For $iLoopCounter = 1 To $i_Student_Subject_Count
		$sReadData = IniRead($sYearName, $sCurrent_MP & $sStudent & $sSubject, $iLoopCounter, "")
		If $sReadData > "" Then GUICtrlCreateListViewItem($sReadData, $handle)
	Next
EndFunc   ;==>_Read_Students_Subjects


Func _Get_Average($sCurrent_MP, $sCurrent_Student, $Subject)
	$iTempSubjectCount = IniRead($sYearName, $sCurrent_MP & $sCurrent_Student & "GradeCounts", $Subject, "")
	Local $ir
	Local $sTempRet
	Local $aRetVal[($iTempSubjectCount + 1) ]
	Local $iDW_Total
	Local $iTest_Total
	Local $iQuiz_Total
	Local $iDiv_Total
	Local $iDW_Count
	Local $iTest_Count
	Local $iQuiz_Count
	Local $iRet_Avg

	Local $iTest_Percent = $iTest_Weight / 100
	Local $iQuiz_Percent = $iQuiz_Weight / 100

	For $ir = 1 To $iTempSubjectCount
		ConsoleWrite($ir & @CR)
		$sTempRet = IniRead($sYearName, $sCurrent_MP & $sCurrent_Student & $Subject, $ir, "")
		Local $aSplitRet = StringSplit($sTempRet, "|")

		Select
			Case $aSplitRet[5] = "Daily Work"
				$iDW_Total = $iDW_Total + Int($aSplitRet[1])
				$iDW_Count = $iDW_Count + 1

			Case $aSplitRet[5] = "Test"
				$iTest_Total = $iTest_Total + Int($aSplitRet[1])
				$iTest_Count = $iTest_Count + 1

			Case $aSplitRet[5] = "Quiz"
				$iQuiz_Total = $iQuiz_Total + Int($aSplitRet[1])
				$iQuiz_Count = $iQuiz_Count + 1

		EndSelect


	Next



	If $iDW_Total > 1 Then
		Local $iDW_Avg = $iDW_Total / $iDW_Count
	EndIf

	If $iTest_Total > 1 Then
		Local $iTest_Avg = $iTest_Total / $iTest_Count
	EndIf

	If $iQuiz_Total > 1 Then
		Local $iQuiz_Avg = $iQuiz_Total / $iQuiz_Count
	EndIf


	Select

		Case $iDW_Total > 0 And $iTest_Total > 0 And $iQuiz_Total > 0
			$iDW_Percent = (100 - ($iTest_Weight + $iQuiz_Weight)) / 100
			$iRet_Avg = ($iDW_Avg * $iDW_Percent) + ($iTest_Avg * $iTest_Percent) + ($iQuiz_Avg * $iQuiz_Percent)

		Case $iDW_Total > 0 And $iTest_Total > 0
			$iDW_Percent = (100 - $iTest_Weight) / 100
			$iRet_Avg = ($iDW_Avg * $iDW_Percent) + ($iTest_Avg * $iTest_Percent)

		Case $iDW_Total > 0 And $iQuiz_Total > 0
			$iDW_Percent = (100 - $iQuiz_Weight) / 100
			$iRet_Avg = ($iDW_Avg * $iDW_Percent) + ($iQuiz_Avg * $iQuiz_Percent)

		Case $iQuiz_Total > 0 And $iTest_Total > 0
			$iMagic_Number = 100 / ($iTest_Weight + $iQuiz_Weight)
			$iTest_Percent = ($iTest_Weight * $iMagic_Number) / 100
			$iQuiz_Percent = ($iQuiz_Weight * $iMagic_Number) / 100
			$iRet_Avg = ($iQuiz_Avg * $iQuiz_Percent) + ($iTest_Avg * $iTest_Percent)

		Case $iDW_Total > 0
			$iRet_Avg = $iDW_Avg

		Case $iTest_Total > 0
			$iRet_Avg = $iTest_Avg

		Case $iQuiz_Total > 0
			$iRet_Avg = $iQuiz_Avg

	EndSelect

	GUICtrlSetData($iMenuGradeCount, "Grade Count = " & $iTempSubjectCount)
	GUICtrlSetData($iMenuDailyWorkCount, "Daily Work Count = " & $iDW_Count)
	GUICtrlSetData($iMenuQuizCount, "Quiz Count = " & $iQuiz_Count)
	GUICtrlSetData($iMenuTestCount, "Test Count = " & $iTest_Count)
	GUICtrlSetData($iMenuUnroundedAvg, "Unrounded Avg = " & $iRet_Avg)

	Return Round($iRet_Avg)

EndFunc   ;==>_Get_Average


Func _Edit_Grade($Grade_View_Handle, $sTemp_Student, $sTemp_Subject, $i_Student_Subject_Count, $sTemp_MP, $Current_Student_Average)


	Local $sTemp_Info = GUICtrlRead(GUICtrlRead($Grade_View_Handle))
	If $sTemp_Info = "" Then Return
	$sTemp_Info = _ABC_StringTrim ($sTemp_Info, "|")
	Local $aSplitRet = StringSplit($sTemp_Info, "|")


	$aParentWin_Pos = WinGetPos($BlackBoard_Main_GUI_Handle, "")

	$hEdit_Grade = GUICreate("Edit Grade Info", 392, 101, $aParentWin_Pos[0] + 100, $aParentWin_Pos[1] + 100, BitOR($WS_SYSMENU, $WS_CAPTION, $WS_POPUPWINDOW, $WS_BORDER, $WS_CLIPSIBLINGS), -1, $BlackBoard_Main_GUI_Handle)
	$save_grade_edit = GUICtrlCreateButton("Save", 144, 64, 105, 25, 0)
	$hcancelgrade_edit = GUICtrlCreateButton("Cancel", 272, 64, 105, 25, 0)
	$_edit_grade = GUICtrlCreateInput($aSplitRet[1], 8, 24, 65, 21)
	$_edit_subject = GUICtrlCreateInput($aSplitRet[2], 82, 24, 65, 21)
	$_edit_weekday = GUICtrlCreateInput($aSplitRet[3], 155, 24, 65, 21)
	$_edit_date = GUICtrlCreateInput($aSplitRet[4], 226, 24, 65, 21)
	$_edit_calc_type = GUICtrlCreateCombo("", 298, 24, 79, 21, BitOR($CBS_DROPDOWNLIST, $CBS_AUTOHSCROLL))


	GUICtrlSetData($_edit_calc_type, "Daily Work|Test|Quiz", $aSplitRet[5])
	GUICtrlSetState($save_grade_edit, $GUI_DEFBUTTON)
	GUISetState(@SW_SHOW)

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE;	   Or $hcancelgrade_edit
				GUIDelete($hEdit_Grade)
				Return

			Case $hcancelgrade_edit
				GUIDelete($hEdit_Grade)
				Return

			Case $save_grade_edit
				GUICtrlSetData(GUICtrlRead($Grade_View_Handle), GUICtrlRead($_edit_grade) & "|" & GUICtrlRead($_edit_subject) & "|" & GUICtrlRead($_edit_weekday) & "|" & GUICtrlRead($_edit_date) & "|" & GUICtrlRead($_edit_calc_type))
				_Save_Grade_List($Grade_View_Handle, $sTemp_Student, $sTemp_Subject, $i_Student_Subject_Count, $sTemp_MP)
				GUICtrlSetData($Current_Student_Average, _Get_Average($sTemp_MP, $sTemp_Student, $sTemp_Subject) & "%")
				GUIDelete($hEdit_Grade)
				Return


		EndSwitch
	WEnd


EndFunc   ;==>_Edit_Grade


Func _Save_Grade_List($Grade_View_Handle, $sTemp_Student, $sTemp_Subject, $i_Student_Subject_Count, $sTemp_MP)

	Local $it_Count = 0
	Local $sWrite_Subjects = ""
	$index_temp = _GUICtrlListView_GetItemCount ($Grade_View_Handle)
	IniWrite($sYearName, $sCurrent_MP & $sTemp_Student & "GradeCounts", $sTemp_Subject, $index_temp)
	Do
		$sItem = _GUICtrlListView_GetItemTextString ($Grade_View_Handle, $it_Count)
		$it_Count = $it_Count + 1
		IniWrite($sYearName, $sCurrent_MP & $sTemp_Student & $sTemp_Subject, String($it_Count), $sItem)
	Until $it_Count >= $index_temp

EndFunc   ;==>_Save_Grade_List

Func _Generate_Graph()
	Local $sTemp_Student_List = IniRead($sYearName, "Information", "Students", "")
	$sTemp_Student_List = _ABC_StringTrim ($sTemp_Student_List, "|")
	Local $aTemp_Student_List = StringSplit($sTemp_Student_List, "|")
	Local $iStudent_Loop_Count = 0

	Local $aParentWin_Pos = WinGetPos($BlackBoard_Main_GUI_Handle, "")

	$hGuiProgress = GUICreate("Generating Graph Please Wait...", 551, 75, $aParentWin_Pos[0] + 100, $aParentWin_Pos[1] + 100, BitOR($WS_DLGFRAME, $WS_CLIPSIBLINGS), -1, $BlackBoard_Main_GUI_Handle)
	$Generate_Progress = GUICtrlCreateProgress(8, 16, 529, 17)

	GUISetState(@SW_SHOW)

	$sTemp_Subject_Name = IniRead($sYearName, $aTemp_Student_List[1], "Subjects", "")
	$aTemp_Subject_Name = StringSplit($sTemp_Subject_Name, "|")

	$Temp_Calc = ($aTemp_Student_List[0] * 10) * $aTemp_Subject_Name[0]
	$Temp_Counter = 100 / $Temp_Calc
	Local $IGUIGCounter
	For $iStudent_Loop_Count = 1 To $aTemp_Student_List[0]
		;FileWriteLine($hOutFile, $aTemp_Student_List[$iStudent_Loop_Count])


		$sTemp_Subject_Name = IniRead($sYearName, $aTemp_Student_List[$iStudent_Loop_Count], "Subjects", "")

		Local $iSubject_Loop_Count = 0
		$sTemp_Subject_Name = _ABC_StringTrim ($sTemp_Subject_Name, "|")
		Local $aTemp_Subject_Name = StringSplit($sTemp_Subject_Name, "|")

		Local $sWrite_Grade


		;Loop over Student's subjects
		Local $iMark = 0

		For $iMark = 1 To 10
			; $sWrite_Grade = "MP "&$iMark
			$s_MarkingPeriod = "MP " & String($iMark)
			$sLoop_SubjectList = ""
			$sWrite_Grade = ""
			For $iTemp_Subject_Count = 1 To $aTemp_Subject_Name[0]
				$IGUIGCounter = $IGUIGCounter + $Temp_Counter
				GUICtrlSetData($Generate_Progress, $IGUIGCounter)
				;$aTemp_Subject_Name[$iTemp_Subject_Count]
				$sLoop_SubjectList = $sLoop_SubjectList & ",'" & $aTemp_Subject_Name[$iTemp_Subject_Count] & "'"


				$sTempGrade = _Get_Average("Marking Period " & $iMark, $aTemp_Student_List[$iStudent_Loop_Count], $aTemp_Subject_Name[$iTemp_Subject_Count])

				If $sTempGrade = "0" Then $sTempGrade = ""
				If $sWrite_Grade > "" Then
					$sWrite_Grade = $sWrite_Grade & "," & $sTempGrade
				Else
					$sWrite_Grade = $sTempGrade
				EndIf

				; FileWriteLine($hOutFile, $aTemp_Subject_Name[$iTemp_Subject_Count]&".........."&$sTempGrade)
			Next
			Html_Graph_Add_Data($s_MarkingPeriod, $sWrite_Grade)
		Next
		Html_Graph_HC_Close(_ABC_StringTrim ($sLoop_SubjectList, ","), $aTemp_Student_List[$iStudent_Loop_Count], $aTemp_Student_List[$iStudent_Loop_Count], "")
	Next

	Html_Write_File()

	Run('"' & $Temp_BH_Dir & '\GoogleChromePortable\GoogleChromePortable.exe" --app="' & $Temp_BH_Dir & "\BlackBoard_HD_Graph.html" & '"')

	AutoItSetOption("WinTitleMatchMode", 3) ;1=start, 2=subStr, 3=exact, 4=advanced, -1 to -4=Nocase


	WinWait("BlackBoard HD Graph", "", 60)

	GUIDelete($hGuiProgress)

EndFunc   ;==>_Generate_Graph
Func _Generate_Graph1()
	RunWait("cmd.exe /c del /q *.*", $Temp_BH_Dir & "\bmp", @SW_HIDE)

	;;Loop over students
	Local $sOutput_Name = ""


	Local $sTemp_Student_List = IniRead($sYearName, "Information", "Students", "")
	$sTemp_Student_List = _ABC_StringTrim ($sTemp_Student_List, "|")
	Local $aTemp_Student_List = StringSplit($sTemp_Student_List, "|")
	Local $iStudent_Loop_Count = 0

	For $iStudent_Loop_Count = 1 To $aTemp_Student_List[0]
		;FileWriteLine($hOutFile, $aTemp_Student_List[$iStudent_Loop_Count])

		$sTemp_Subject_Name = IniRead($sYearName, $aTemp_Student_List[$iStudent_Loop_Count], "Subjects", "")

		Local $iSubject_Loop_Count = 0
		$sTemp_Subject_Name = _ABC_StringTrim ($sTemp_Subject_Name, "|")
		Local $aTemp_Subject_Name = StringSplit($sTemp_Subject_Name, "|")

		Local $iColor = 1
		Local $sWrite_Grade


		;Loop over Student's subjects
		Local $aValues[15][15]
		Local $aDataColored[15][15]
		$aValues[0][0] = 14
		For $iTemp_Subject_Count = 1 To $aTemp_Subject_Name[0]
			$sWrite_Grade = $aTemp_Subject_Name[$iTemp_Subject_Count]


			$sTempGrade = _Get_Average($sCurrent_MP, $aTemp_Student_List[$iStudent_Loop_Count], $aTemp_Subject_Name[$iTemp_Subject_Count])
			If $sTempGrade = "0" Then $sTempGrade = ""

			;$sOutput_Name = $sOutput_Name &"|"&$sTempGrade&","&$aTemp_Subject_Name[$iTemp_Subject_Count]

			$aValues[$iTemp_Subject_Count][0] = $aTemp_Subject_Name[$iTemp_Subject_Count]
			$aValues[$iTemp_Subject_Count][1] = $sTempGrade




			; FileWriteLine($hOutFile, $aTemp_Subject_Name[$iTemp_Subject_Count]&".........."&$sTempGrade)
		Next

		_GDIPlus_Startup ()
		$sImg = _3D_BarChart($aValues, $aTemp_Student_List[$iStudent_Loop_Count], "Averages for " & $sCurrent_MP, 700, 400, 0x00FFFFFF, $Temp_BH_Dir & "\bmp")
		_GDIPlus_Shutdown ()
		;ConsoleWrite($sImg & @CR)
		; RunWait("cmd.exe /c graphic.exe "&chr(34)&$aTemp_Subject_Name[0]&"|"&@ScriptDir&"\bmp\"&$aTemp_Student_List[$iStudent_Loop_Count]&".bmp|"&$aTemp_Student_List[$iStudent_Loop_Count]&$sOutput_Name&Chr(34), @ScriptDir, @SW_HIDE)
		; ConsoleWrite("cmd.exe /c graphic.exe "&chr(34)&$aTemp_Subject_Name[0]&"|"&@ScriptDir&"\bmp\"&$aTemp_Student_List[$iStudent_Loop_Count]&".bmp|"&$aTemp_Student_List[$iStudent_Loop_Count]&$sOutput_Name&@CR)
		$sOutput_Name = ""




	Next

	$hTemp_Handle = _Html_Init ($Temp_BH_Dir & "\BlackBoard_HD_Graph.html")

	_Html_SendCSS ("li", "display:inline;")

	_Html_Set_Title ($hTemp_Handle, "BlackBoard HD Graph")

	$atempbmp = _FileListToArray($Temp_BH_Dir & "\bmp", "*.png")

	Local $i

	$ilist_handle = _Html_AddList ()
	For $i = 1 To $atempbmp[0]
		_Html_AddListItem ($ilist_handle, StringTrimRight($atempbmp[$i], 4), StringTrimRight($atempbmp[$i], 4))
	Next
	_Html_Close_List ($hTemp_Handle, $ilist_handle)

	$i = 0

	_Html_AddDiv ($hTemp_Handle, "Align=Center")

	For $i = 1 To $atempbmp[0]
		_Html_Internal_Link ($hTemp_Handle, StringTrimRight($atempbmp[$i], 4))
		$graphpic = _Html_AddPic ($hTemp_Handle, "bmp\" & $atempbmp[$i], StringTrimRight($atempbmp[$i], 4) & "'s Averages for " & $sCurrent_MP)
		_Html_NewLine ($hTemp_Handle)
		_Html_Internal ($hTemp_Handle, "Back to Top", "#Top")
		_Html_NewLine ($hTemp_Handle, 3)
	Next

	_Html_CloseDiv ($hTemp_Handle)
	_Html_Close ($hTemp_Handle)

	; ShellExecute(@ScriptDir&"\BlackBoard_HD_Graph.html")

	$aParentWin_Pos = WinGetPos($BlackBoard_Main_GUI_Handle, "")

	Local $oIE_Graph = ObjCreate("Shell.Explorer.2")
	$Graph_Gui = GUICreate("BlackBoard HD Graph", 890, 631, $aParentWin_Pos[0] + 100, $aParentWin_Pos[1] + 100, -1, -1, $BlackBoard_Main_GUI_Handle)
	Local $Graph_File_Menu = GUICtrlCreateMenu("&File")
	Local $Graph_File_Menu_Print = GUICtrlCreateMenuItem("Print", $Graph_File_Menu)
	Local $Export_As_PDF = GUICtrlCreateMenuItem("Export as PDF", $Graph_File_Menu)

	GUICtrlCreateObj($oIE_Graph, 0, 0, 889, 629)
	$oIE_Graph.navigate ($Temp_BH_Dir & "\BlackBoard_HD_Graph.html")
	GUISetState(@SW_SHOW)


	While 1
		$nMsg_Graph = GUIGetMsg()
		Switch $nMsg_Graph
			Case $GUI_EVENT_CLOSE
				GUIDelete($Graph_Gui)
				ExitLoop

			Case $Graph_File_Menu_Print
				$oIE_Graph.document.parentwindow.Print ()

			Case $Export_As_PDF
				$PDF_File_Name = FileSaveDialog("Save File As", "::{450D8FBA-AD25-11D0-98A8-0800361B1103}", "PDF Files (*.pdf)", $FD_PATHMUSTEXIST)
				If $PDF_File_Name <> "" Then
					_ABC_Make_PDF ($Temp_BH_Dir & "\bmp", $PDF_File_Name)
				EndIf

		EndSwitch
	WEnd

EndFunc   ;==>_Generate_Graph1

;Generate Report as html
;'''''''''''''''''''''''''
;'''''''''''''''''''''''''
;''''''''''''''''''''''''

#Region +++++Generate Report++++++
Func _Generate_Report()
	Local $iType = IniRead($sYearName, "Information", "ReportViewStyle", "0")
	_Generate_Html_Report($iType)

	$aParentWin_Pos = WinGetPos($BlackBoard_Main_GUI_Handle, "")

	Local $oIE = ObjCreate("Shell.Explorer.2")
	$Report_Gui = GUICreate("BlackBoard HD Report", 840, 531, $aParentWin_Pos[0] + 100, $aParentWin_Pos[1] + 100, -1, -1, $BlackBoard_Main_GUI_Handle)
	;	Local $Report_File_Menu = GUICtrlCreateMenu("&File")
	;	Local $Report_File_Menu_Print = GUICtrlCreateMenuItem("Print", $Report_File_Menu)
	$Report_File_Menu = _GUICtrlMenu_CreateMenu ()
	_GUICtrlMenu_InsertMenuItem ($Report_File_Menu, 0, "&Print", $e_Report_File_Menu_Print)

	$Report_Generate_Menu = _GUICtrlMenu_CreateMenu ()
	_GUICtrlMenu_InsertMenuItem ($Report_Generate_Menu, 0, "&Letter Grades", $e_idSet_Number)
	_GUICtrlMenu_InsertMenuItem ($Report_Generate_Menu, 1, "&Number Grades", $e_idSet_Letter)

	$hMain = _GUICtrlMenu_CreateMenu ()
	_GUICtrlMenu_InsertMenuItem ($hMain, 0, "&File", 0, $Report_File_Menu)
	_GUICtrlMenu_InsertMenuItem ($hMain, 1, "&View", 0, $Report_Generate_Menu)

	_GUICtrlMenu_SetMenu ($Report_Gui, $hMain)
	GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")


	GUICtrlCreateObj($oIE, 0, 0, 839, 529)
	$oIE.navigate ($Temp_BH_Dir & "\BlackBoard_HD_Report.html")
	GUISetState(@SW_SHOW)
	_GUICtrlMenu_CheckMenuItem ($Report_Generate_Menu, $iType)



	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				GUIDelete($Report_Gui)
				ExitLoop


		EndSwitch

		Switch $sMenuVal
			Case "Print"
				$oIE.document.parentwindow.Print ()

			Case "Number"
				_GUICtrlMenu_CheckMenuItem ($Report_Generate_Menu, 0)
				_GUICtrlMenu_CheckMenuItem ($Report_Generate_Menu, 1, False)
				_Generate_Html_Report(0)
				$oIE.document.execCommand ("Refresh")
				IniWrite($sYearName, "Information", "ReportViewStyle", "0")

			Case "Letter"
				_GUICtrlMenu_CheckMenuItem ($Report_Generate_Menu, 1)
				_GUICtrlMenu_CheckMenuItem ($Report_Generate_Menu, 0, False)
				_Generate_Html_Report(1)
				$oIE.document.execCommand ("Refresh")
				IniWrite($sYearName, "Information", "ReportViewStyle", "1")

		EndSwitch
		$sMenuVal = ""
	WEnd


	;FileClose($hOutFile)

EndFunc   ;==>_Generate_Report


Func _Generate_Html_Report($iType)

	$hTemp_Handle = _Html_Init ($Temp_BH_Dir & "\BlackBoard_HD_Report.html")
	_Html_Set_Title ($hTemp_Handle, "BlackBoard HD Report")
	;add a wider column for the subject names
	_Html_AddColumn ($hTemp_Handle, 86)

	;add ten narrow columns for the grades to go in
	_Html_AddColumn ($hTemp_Handle, 50, 10)

	_Html_SendCSS ("table", "width: 586px; border:8px ridge #7B7B7B;text-align:center;")


	;Local $hOutFile = FileOpen($Temp_BH_Dir&"\Averages.txt", 2)

	Local $sTemp_Student_List = IniRead($sYearName, "Information", "Students", "")
	$sTemp_Student_List = _ABC_StringTrim ($sTemp_Student_List, "|")
	Local $aTemp_Student_List = StringSplit($sTemp_Student_List, "|")
	Local $iStudent_Loop_Count = 0

	;;Loop over students

	For $iStudent_Loop_Count = 1 To $aTemp_Student_List[0]
		;FileWriteLine($hOutFile, $aTemp_Student_List[$iStudent_Loop_Count])

		Local $html_table1 = _Html_InitTable ($hTemp_Handle, $aTemp_Student_List[$iStudent_Loop_Count])
		_Html_Table_SetCaptionFont ($html_table1, 25, "normal", "Times New Roman")
		Local $html_temp_handle = _Html_AddRow ($hTemp_Handle, "|Mp 1|Mp 2|Mp 3|Mp 4|Mp 5|Mp 6|Mp 7|Mp 8|Mp 9|Mp 10")
		_Html_SetRowColor ($html_temp_handle, "#C5D9F1")

		$sTemp_Subject_Name = IniRead($sYearName, $aTemp_Student_List[$iStudent_Loop_Count], "Subjects", "")

		Local $iSubject_Loop_Count = 0
		$sTemp_Subject_Name = _ABC_StringTrim ($sTemp_Subject_Name, "|")
		Local $aTemp_Subject_Name = StringSplit($sTemp_Subject_Name, "|")

		Local $iColor = 1
		Local $sWrite_Grade


		;Loop over Student's subjects

		For $iTemp_Subject_Count = 1 To $aTemp_Subject_Name[0]
			$sWrite_Grade = $aTemp_Subject_Name[$iTemp_Subject_Count]
			If $iColor = 1 Then
				$iColor = 0
				Local $sColor = 'white'
			Else
				$iColor = 1
				Local $sColor = '#DDDDDD'
			EndIf
			;Loop MPs
			Local $iMark = 0
			For $iMark = 1 To 10
				$sTempGrade = _Get_Average("Marking Period " & $iMark, $aTemp_Student_List[$iStudent_Loop_Count], $aTemp_Subject_Name[$iTemp_Subject_Count])
				If $sTempGrade = "0" Then $sTempGrade = ""
				If $sTempGrade > "" Then
					If $iType = 0 Then $sTempGrade = _Grade2Letter($sTempGrade)
				EndIf
				If $sWrite_Grade > "" Then
					$sWrite_Grade = $sWrite_Grade & "|" & $sTempGrade
				Else
					$sWrite_Grade = $sTempGrade
				EndIf

				; FileWriteLine($hOutFile, $aTemp_Subject_Name[$iTemp_Subject_Count]&".........."&$sTempGrade)
			Next

			Local $html_temp_handle = _Html_AddRow ($hTemp_Handle, $sWrite_Grade)
			_Html_SetRowColor ($html_temp_handle, $sColor)


		Next
		_Html_CloseTable ($hTemp_Handle)




	Next
	_Html_Close ($hTemp_Handle)
EndFunc   ;==>_Generate_Html_Report

Func WM_COMMAND($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg, $lParam
	Switch _WinAPI_LoWord ($wParam)

		Case $e_idSet_Number
			$sMenuVal = "Number"

		Case $e_Report_File_Menu_Print
			$sMenuVal = "Print"

		Case $e_idSet_Letter
			$sMenuVal = "Letter"
	EndSwitch

	Return $GUI_RUNDEFMSG

EndFunc   ;==>WM_COMMAND



#EndRegion ++++End Region Generate Report+++++

#Region ++++Profile PIC+++++
Func _Get_Profile_Pic($AddStudent_Gui)

	$aParentWin_Pos = WinGetPos($AddStudent_Gui, "")

	Local $sFileOpenDialog = FileOpenDialog("Select Image File", @MyDocumentsDir & "\", "Image File(*.jpg;*.jpeg;*.bmp)", $FD_FILEMUSTEXIST, -1, $AddStudent_Gui)
	If $sFileOpenDialog = "" Then Return

	$h_Profile_Pic_GUI = GUICreate("Select Profile Picture", 500, 400, $aParentWin_Pos[0] + 100, $aParentWin_Pos[1] + 100, -1, -1, $AddStudent_Gui)
	GUISetBkColor(0x000000, $h_Profile_Pic_GUI)

	GUISetState(@SW_SHOW)
	; Initialize GDI+ library
	_GDIPlus_Startup ()

	; Draw bitmap to GUI
	$hBitmap = _GDIPlus_BitmapCreateFromFile ($sFileOpenDialog)
	$hGraphic = _GDIPlus_GraphicsCreateFromHWND ($h_Profile_Pic_GUI)
	$aNewPicSize = _ResizePic($hBitmap)
	$hBitmap = _GDIPlus_ImageResize ($hBitmap, $aNewPicSize[0], $aNewPicSize[1])
	_GDIPlus_GraphicsDrawImage ($hGraphic, $hBitmap, 0, 0)
	_GDIPlus_GraphicsSetSmoothingMode ($hGraphic, $GDIP_SMOOTHINGMODE_HIGHQUALITY) ;Sets the graphics object rendering quality (antialiasing)
	$hPen = _GDIPlus_PenCreate (0xFF000000, 1)
	_GDIPlus_PenSetDashStyle ($hPen, $GDIP_DASHSTYLEDASHDOT)
	$hPath1 = _GDIPlus_PathCreate () ;Create new path object
	_GDIPlus_PathAddRectangle ($hPath1, 20, 10, 64, 64)
	$hgdraw = _GDIPlus_GraphicsDrawPath ($hGraphic, $hPath1, $hPen) ;Draw path1 to graphics handle (GUI)


	GUIRegisterMsg($WM_LBUTTONDOWN, "WM_LBUTTONDOWN")
	GUIRegisterMsg($WM_MOUSEMOVE, "WM_MOUSEMOVE")
	GUIRegisterMsg($WM_MOUSEWHEEL, "WM_MOUSEWHEEL")
	$sRetVal = ""

	Do
		If $s_Profile_Pic_Done > "" Then
			$sRetVal = $s_Profile_Pic_Done
			$s_Profile_Pic_Done = ""
			ExitLoop
		EndIf
	Until GUIGetMsg() = $GUI_EVENT_CLOSE

	GUIRegisterMsg($WM_LBUTTONDOWN, "")
	GUIRegisterMsg($WM_MOUSEMOVE, "")
	GUIRegisterMsg($WM_MOUSEWHEEL, "")
	GUIDelete($h_Profile_Pic_GUI)

	Return $sRetVal

EndFunc   ;==>_Get_Profile_Pic



Func WM_LBUTTONDOWN($hWnd, $iMsg, $wParam, $lParam)

	#forceref $hWnd, $iMsg, $wParam

	$g_iMouseX = BitAND($lParam, 0x0000FFFF)
	$g_iMouseY = BitShift($lParam, 16)
	$hClone = _GDIPlus_BitmapCloneArea ($hBitmap, $g_iMouseX, $g_iMouseY, $iScale, $iScale, $GDIP_PXF24RGB)
	$hClone = _GDIPlus_ImageResize ($hClone, 64, 64)
	; Save bitmap to file
	_GDIPlus_ImageSaveToFile ($hClone, $Temp_BH_Dir & "\clone.bmp")
	; Clean up resources
	_GDIPlus_ImageDispose ($hClone)
	;_GDIPlus_ImageDispose($hImage)
	_WinAPI_DeleteObject ($hBitmap)
	; Shut down GDI+ library
	_GDIPlus_Shutdown ()
	GUIDelete($h_Profile_Pic_GUI)
	GUIRegisterMsg($WM_LBUTTONDOWN, "")
	GUIRegisterMsg($WM_MOUSEMOVE, "")
	GUIRegisterMsg($WM_MOUSEWHEEL, "")
	$s_Profile_Pic_Done = $Temp_BH_Dir & "\clone.bmp"
	Return

EndFunc   ;==>WM_LBUTTONDOWN

Func WM_MOUSEMOVE($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg
	$aCursorInfo = GUIGetCursorInfo($h_Profile_Pic_GUI)
	_GDIPlus_GraphicsDrawImage ($hGraphic, $hBitmap, 0, 0)
	_GDIPlus_GraphicsSetSmoothingMode ($hGraphic, $GDIP_SMOOTHINGMODE_HIGHQUALITY) ;Sets the graphics object rendering quality (antialiasing)
	$hPath1 = _GDIPlus_PathCreate () ;Create new path object
	_GDIPlus_PathAddRectangle ($hPath1, $aCursorInfo[0], $aCursorInfo[1], $iScale, $iScale)
	_GDIPlus_GraphicsDrawPath ($hGraphic, $hPath1, $hPen) ;Draw path1 to graphics handle (GUI)

	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_MOUSEMOVE



Func WM_MOUSEWHEEL($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg, $lParam
	$aCursorIfo = GUIGetCursorInfo($h_Profile_Pic_GUI)
	$tPoint2 = _WinAPI_GetMousePos (True, $hWnd)


	Switch BitShift($wParam, 16)
		Case 120
			$iScale = $iScale + 2
		Case Else
			$iScale = $iScale - 2
	EndSwitch

	_GDIPlus_GraphicsDrawImage ($hGraphic, $hBitmap, 0, 0)
	_GDIPlus_GraphicsSetSmoothingMode ($hGraphic, $GDIP_SMOOTHINGMODE_HIGHQUALITY) ;Sets the graphics object rendering quality (antialiasing)
	$hPath1 = _GDIPlus_PathCreate () ;Create new path object
	_GDIPlus_PathAddRectangle ($hPath1, $aCursorIfo[0], $aCursorIfo[1], $iScale, $iScale)
	$hgdraw = _GDIPlus_GraphicsDrawPath ($hGraphic, $hPath1, $hPen) ;Draw path1 to graphics handle (GUI)

	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_MOUSEWHEEL


Func _ResizePic($hPic)
	Local $aRetVal[2]
	Local $iWindowWidth = 500
	Local $iWindowHeight = 400
	Local $iImageWidth = _GDIPlus_ImageGetWidth ($hPic)
	Local $iImageHeight = _GDIPlus_ImageGetHeight ($hPic)



	If $iImageWidth > $iWindowWidth Then
		$ira = $iImageWidth / $iImageHeight
		$idi = $iImageWidth - $iWindowWidth
		$ihdi = $idi / $ira

		$iImageWidth = $iImageWidth - $idi
		$iImageHeight = $iImageHeight - $ihdi

	EndIf

	If $iImageHeight > $iWindowHeight Then
		$ira = $iImageWidth / $iImageHeight
		$idi = $iImageHeight - $iWindowHeight
		$iwdi = $idi / $ira
		ConsoleWrite($ira & @CR)

		$iImageWidth = Round($iImageWidth - $iwdi)
		$iImageHeight = Round($iImageHeight - $idi)
	EndIf

	$aRetVal[1] = $iImageHeight
	$aRetVal[0] = $iImageWidth
	ConsoleWrite($aRetVal[0] & "++" & $aRetVal[1])
	Return $aRetVal

EndFunc   ;==>_ResizePic

#EndRegion ++++ProfilePic+++++

#Region +++++++BarChart+++++++
; #FUNCTION# ====================================================================================================================
; Name ..........: _3D_BarChart
; Description ...:
; Syntax ........: _3D_BarChart($aArray[, $sTitle = "Main title"[, $sSubTitle = "Subtitle"[, $iW = 500[, $iH = 350[,
; $lBgColor = 0x80efefef]]]]])
; Parameters ....: $aArray - An array of data
; |[i][0]=Label,
; |[i][1]=Value,
; |[i][2]=Colour.
; $sTitle - [optional] Title of the graph. Default is "Main title".
; $sSubTitle - [optional] Subtitle of the graph. Default is "Subtitle".
; $iW - [optional] Width of the canvas. Default is 500.
; $iH - [optional] Height of the canvas. Default is 350.
; $lBgColor - [optional] Background colour. Default is 0x80efefef.
; $sFolder - [optional] Folder path (without the leading "") where the images are saved. Default is @ScriptDir.
; Return values .: Image full path
; Return values .: None
; Author ........: Mihai Iancu (taietel at yahoo dot com)
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _3D_BarChart($aArray, $sTitle = "Main title", $sSubTitle = "Subtitle", $iW = 500, $iH = 350, $lBgColor = 0x80efefef, $sFolder = "")

	Local $Biggest_Text
	For $i = 1 To $aArray[0][0]
		If __Chart_GetTextLabelWidth($aArray[$i][0], "Times New Roman", 13) > $Biggest_Text Then
			ConsoleWrite("BT=" & $Biggest_Text & @CRLF)
			$Biggest_Text = __Chart_GetTextLabelWidth($aArray[$i][0], "Times New Roman", 13)
		EndIf
	Next

	If $sFolder = "" Then $sFolder = @ScriptDir
	Local $hGraphic, $hBrush, $hFormat, $hFamily, $hFont, $tLayout, $iDepth = 30
	Local $aResult = DllCall($__g_hGDIPDll, "uint", "GdipCreateBitmapFromScan0", "int", $iW + 16, "int", $iH + ($Biggest_Text + 41), "int", 0, "int", 0x0026200A, "ptr", 0, "int*", 0)
	If @error Then Return SetError(@error, @extended, 0)
	$hGraphic = _GDIPlus_ImageGetGraphicsContext ($aResult[6])
	_GDIPlus_GraphicsSetTextRenderingHint ($hGraphic, $GDIP_TEXTRENDERINGHINT_ANTIALIASGRIDFIT)
	_GDIPlus_GraphicsClear ($hGraphic, 0x00FFFFFF)
	_GDIPlus_GraphicsSetSmoothingMode ($hGraphic, 2)
	;create canvas
	$hBrush = _GDIPlus_BrushCreateSolid ($lBgColor)


	Local $hPen = _GDIPlus_PenCreate (0x70101080)
	Local $aPoints[5][2]
	$aPoints[0][0] = 4
	$aPoints[1][0] = 0
	$aPoints[1][1] = 0
	$aPoints[2][0] = $iW + 15
	$aPoints[2][1] = 0
	$aPoints[3][0] = $iW + 15
	$aPoints[3][1] = $iH + ($Biggest_Text + 40)
	$aPoints[4][0] = 0
	$aPoints[4][1] = $iH + ($Biggest_Text + 40)
	_GDIPlus_GraphicsFillPolygon ($hGraphic, $aPoints, $hBrush)

	_GDIPlus_GraphicsDrawPolygon ($hGraphic, $aPoints, $hPen)
	_GDIPlus_PenDispose ($hPen)

	Local $iMax = 0, $iMin = 0
	For $i = 1 To $aArray[0][0]
		If $iMax < $aArray[$i][1] Then $iMax = _Round($aArray[$i][1])
		If $iMin > $aArray[$i][1] Then $iMin = $aArray[$i][1]
		$aArray[$i][2] = "0xFA" & Hex((Random(80, 190, 1) * 0x10000) + (Random(80, 190, 1) * 0x100) + Random(30, 190, 1), 6)
	Next
	Local $iFH = Floor($iH / 28)
	$hBrush = _GDIPlus_BrushCreateSolid (0xF0EF0000)
	$hFormat = _GDIPlus_StringFormatCreate ()
	_GDIPlus_StringFormatSetAlign ($hFormat, 1)
	$hFamily = _GDIPlus_FontFamilyCreate ("Open Sans")
	Local $x = ($iH + $iDepth - 60) / $aArray[0][0]
	If $x < 20 Then
		$hFont = _GDIPlus_FontCreate ($hFamily, 8, 0)
	Else
		$hFont = _GDIPlus_FontCreate ($hFamily, 10, 0)
	EndIf
	_GDIPlus_BrushSetSolidColor ($hBrush, 0xf8efefcc)
	;create bars
	For $j = 1 To $aArray[0][0]
		Local $iInalt = Int($aArray[$j][1] * ($iH - 2 * $iDepth) / $iMax)
		For $i = 0 To Int(2 * $iDepth / 3) Step 0.5
			If $i = Int(2 * $iDepth / 3) Then
				_GDIPlus_BrushSetSolidColor ($hBrush, $aArray[$j][2] + 0x111111)
			Else
				_GDIPlus_BrushSetSolidColor ($hBrush, $aArray[$j][2])
			EndIf
			If $aArray[$j][1] <> 0 Then
				_GDIPlus_GraphicsFillRect ($hGraphic, $iDepth - $i + ($j - 1) * Floor(($iW - $iDepth) / $aArray[0][0]) + 10, $iH + $i - $iInalt - $iDepth + 30, Int(Floor(($iW - $iDepth) / $aArray[0][0]) - 10), $iInalt, $hBrush)
			Else
				_GDIPlus_GraphicsFillRect ($hGraphic, $iDepth - $i + ($j - 1) * Floor(($iW - $iDepth) / $aArray[0][0]) + 10, $iH + $i - $iDepth + 26, Int(Floor(($iW - $iDepth) / $aArray[0][0]) - 10), 4, $hBrush)
			EndIf
		Next
		_GDIPlus_StringFormatSetAlign ($hFormat, 1)
		_GDIPlus_BrushSetSolidColor ($hBrush, $aArray[$j][2] + 0xd5222222)
		If $aArray[$j][1] <> 0 Then
			$tLayout = _GDIPlus_RectFCreate (Int($iDepth / 3) + ($j - 1) * Floor(($iW - $iDepth) / $aArray[0][0]) + 10, $iH + Int(2 * $iDepth / 3) - $iInalt - $iDepth + 12, Int($iDepth / 3) + Floor(($iW - $iDepth) / $aArray[0][0]), 2 * $iFH)
		Else
			$tLayout = _GDIPlus_RectFCreate (Int($iDepth / 3) + ($j - 1) * Floor(($iW - $iDepth) / $aArray[0][0]) + 10, $iH + Int(2 * $iDepth / 3) - $iDepth + 8, Int($iDepth / 3) + Floor(($iW - $iDepth) / $aArray[0][0]), 2 * $iFH)
		EndIf
		_GDIPlus_GraphicsDrawStringEx ($hGraphic, $aArray[$j][1], $hFont, $tLayout, $hFormat, $hBrush)
		_GDIPlus_BrushSetSolidColor ($hBrush, 0xe0111111)
		If $aArray[$j][1] <> 0 Then
			$tLayout = _GDIPlus_RectFCreate (Int($iDepth / 3) + ($j - 1) * Floor(($iW - $iDepth) / $aArray[0][0]) + 9, $iH + Int(2 * $iDepth / 3) - $iInalt - $iDepth + 11, Int($iDepth / 3) + Floor(($iW - $iDepth) / $aArray[0][0]), 2 * $iFH)
		Else
			$tLayout = _GDIPlus_RectFCreate (Int($iDepth / 3) + ($j - 1) * Floor(($iW - $iDepth) / $aArray[0][0]) + 9, $iH + Int(2 * $iDepth / 3) - $iDepth + 7, Int($iDepth / 3) + Floor(($iW - $iDepth) / $aArray[0][0]), 2 * $iFH)
		EndIf
		_GDIPlus_GraphicsDrawStringEx ($hGraphic, $aArray[$j][1], $hFont, $tLayout, $hFormat, $hBrush)
	Next

	;title
	_GDIPlus_StringFormatSetAlign ($hFormat, 1)
	$hFont = _GDIPlus_FontCreate ($hFamily, 20, 0)
	_GDIPlus_BrushSetSolidColor ($hBrush, 0x8fdedede)
	$tLayout = _GDIPlus_RectFCreate (1, 10, $iW + 116, 35)
	_GDIPlus_GraphicsDrawStringEx ($hGraphic, $sTitle, $hFont, $tLayout, $hFormat, $hBrush)
	_GDIPlus_BrushSetSolidColor ($hBrush, 0xee950000)
	$tLayout = _GDIPlus_RectFCreate (0, 10 - 1, $iW + 115, 35)
	_GDIPlus_GraphicsDrawStringEx ($hGraphic, $sTitle, $hFont, $tLayout, $hFormat, $hBrush)
	;subtitle
	$hFont = _GDIPlus_FontCreate ($hFamily, 10, 0)
	_GDIPlus_BrushSetSolidColor ($hBrush, 0xFFb00000)
	$tLayout = _GDIPlus_RectFCreate (1, 40, $iW + 115, 20)
	_GDIPlus_GraphicsDrawStringEx ($hGraphic, $sSubTitle, $hFont, $tLayout, $hFormat, $hBrush)

	For $i = 1 To $aArray[0][0]
		_GDIPlus_GraphicsDrawString ($hGraphic, $aArray[$i][0], 26 + (47 * ($i - 1)), $iH + 30, "Times New Roman", 13, 0x0002)
	Next

	Local $sImage = $sFolder & "\" & $sTitle & ".png"
	_GDIPlus_ImageSaveToFile ($aResult[6], $sImage)
	_GDIPlus_FontDispose ($hFont)
	_GDIPlus_FontFamilyDispose ($hFamily)
	_GDIPlus_StringFormatDispose ($hFormat)
	_GDIPlus_BrushDispose ($hBrush)
	_GDIPlus_BitmapDispose ($aResult[6])
	_GDIPlus_GraphicsDispose ($hGraphic)
	Return $sImage
EndFunc   ;==>_3D_BarChart

Func _Round($iNumber)
	If IsNumber($iNumber) Then $iNumber = Round(Ceiling($iNumber / 10) * 10, -1)
	Return $iNumber
EndFunc   ;==>_Round

Func __Chart_GetTextLabelWidth($s_WinText, $s_TextFont, $i_FontSize, $i_FontWeight = -1)
	Local Const $DEFAULT_CHARSET = 0 ; ANSI character set
	Local Const $OUT_CHARACTER_PRECIS = 2
	Local Const $CLIP_DEFAULT_PRECIS = 0
	Local Const $PROOF_QUALITY = 2
	Local Const $FIXED_PITCH = 1
	Local Const $RGN_XOR = 3
	Local Const $LOGPIXELSY = 90


	Local $h_WinTitle = "Get Label Width"
	If $i_FontWeight = "" Or $i_FontWeight = -1 Then $i_FontWeight = 600 ; default Font weight
	Local $h_GUI = GUICreate($h_WinTitle, 10, 10, -100, -100, $WS_POPUPWINDOW, $WS_EX_TOOLWINDOW)
	Local $hDC = DllCall("user32.dll", "int", "GetDC", "hwnd", $h_GUI)

	Local $intDeviceCap = DllCall("gdi32.dll", "long", "GetDeviceCaps", "int", $hDC[0], "long", $LOGPIXELSY)
	$intDeviceCap = $intDeviceCap[0]

	Local $intFontHeight = DllCall("kernel32.dll", "long", "MulDiv", "long", $i_FontSize, "long", $intDeviceCap, "long", 72)
	$intFontHeight = -$intFontHeight[0]

	Local $hMyFont = DllCall("gdi32.dll", "hwnd", "CreateFont", "int", $intFontHeight, _
			"int", 0, "int", 0, "int", 0, "int", $i_FontWeight, "int", 0, _
			"int", 0, "int", 0, "int", $DEFAULT_CHARSET, _
			"int", $OUT_CHARACTER_PRECIS, "int", $CLIP_DEFAULT_PRECIS, _
			"int", $PROOF_QUALITY, "int", $FIXED_PITCH, "str", $s_TextFont)
	DllCall("gdi32.dll", "hwnd", "SelectObject", "int", $hDC[0], "hwnd", $hMyFont[0])

	Local $res = DllStructCreate("int;int")

	Local $ret = DllCall("gdi32.dll", "int", "GetTextExtentPoint32", "int", $hDC[0], "str", $s_WinText, "long", StringLen($s_WinText), "ptr", DllStructGetPtr($res))

	Local $intLabelWidth = DllStructGetData($res, 1)

	GUIDelete($h_GUI)
	Return $intLabelWidth
EndFunc   ;==>__Chart_GetTextLabelWidth

#EndRegion BarChart

Func _Delete_Student($_Student_Name)

EndFunc   ;==>_Delete_Student

Func _About()
	Local $aParentWin_Pos = WinGetPos($BlackBoard_Main_GUI_Handle, "")

	Local $ABOUT_AUTHOR = "Questions/Comments/Bugs"
	Local $MAILTO = "codefoxtech@gmail.com"
	Local $ABOUT = "BlackBoard HD  v1.0" & @CRLF & @CRLF & _
			"Author: The CodeFox" & @CRLF & _
			"License type: Freeware"

	$About_Form = GUICreate("About BlackBoard HD", 345, 130, $aParentWin_Pos[0] + (476 - 172), $aParentWin_Pos[1] + (295 - 65), BitOR($WS_SYSMENU, $WS_CAPTION, $WS_POPUPWINDOW, $WS_BORDER, $WS_CLIPSIBLINGS), -1, $BlackBoard_Main_GUI_Handle)
	$Label1 = GUICtrlCreateLabel($ABOUT, 200, 24, 200, 65)
	$LABEL2 = _GuiCtrlCreateHyperlink($ABOUT_AUTHOR, 200, 89, 200, 20, 0x0000ff, 'E-Mail ' & $MAILTO & ' (comments/questions)')
	$idPic = GUICtrlCreateIcon(@ScriptDir & "\BlackBoard2.ico", "", 12, 8, 150, 114); 150, 114)

	GUISetState(@SW_SHOW)
	Local $aMsg = 0
	While 1
		$aMsg = GUIGetMsg(1)
		Select
			Case $aMsg[0] = $GUI_EVENT_CLOSE And $aMsg[1] = $About_Form
				GUIDelete($About_Form)
				Return
			Case $aMsg[0] = $LABEL2
				ShellExecute("mailto:s.fox01@icloud.com ?subject=Regarding BlackBoard HD (v1.0)")
				;_INetMail($MAILTO, "Regarding BlackBoard HD (v1.0)", "")

		EndSelect
	WEnd


EndFunc   ;==>_About


Func _Help()
	$aParentWin_Pos = WinGetPos($BlackBoard_Main_GUI_Handle, "")

	Local $oIE = ObjCreate("Shell.Explorer.2")
	$Help_GUI = GUICreate("BlackBoard HD Help", 875, 531, $aParentWin_Pos[0] + 100, $aParentWin_Pos[1] + 100, -1, -1, $BlackBoard_Main_GUI_Handle)
	;	Local $Report_File_Menu = GUICtrlCreateMenu("&File")
	;	Local $Report_File_Menu_Print = GUICtrlCreateMenuItem("Print", $Report_File_Menu)

	GUICtrlCreateObj($oIE, 0, 0, 875, 529)
	$oIE.navigate (@ScriptDir & "\Help\Introduction.html")
	GUISetState(@SW_SHOW)
	Local $aMsg = 0
	While 1
		$aMsg = GUIGetMsg(1)
		Select
			Case $aMsg[0] = $GUI_EVENT_CLOSE And $aMsg[1] = $Help_GUI
				GUIDelete($Help_GUI)
				Return
		EndSelect
	WEnd
EndFunc   ;==>_Help


Func _Send_Bug($sYearName)
	; $sYearName
	_ScreenCapture_CaptureWnd (@TempDir & "\27488269.jpg", $BlackBoard_Main_GUI_Handle)
	$Bug_Zip = _Zip_Create (@TempDir & "\" & @UserName & "_BlackBoardHD_Bug_ReportData.zip")
	ConsoleWrite("Creating File" & @CRLF)
	_Zip_AddFile ($Bug_Zip, $sYearName)
	_Zip_AddFile ($Bug_Zip, @TempDir & "\27488269.jpg")
	ConsoleWrite("Adding Screenshot" & @CRLF)
	$Bug_Change = StringReplace($sYearName, ".bh", "_Data")
	ConsoleWrite($Bug_Change & @CR)
	_Zip_AddFolder ($Bug_Zip, $Bug_Change)
	ConsoleWrite("Renaming File" & @CRLF)
	Sleep(5000)
	FileMove(@TempDir & "\" & @UserName & "_BlackBoardHD_Bug_ReportData.zip", @TempDir & "\" & @UserName & "_BlackBoardHD_Bug_ReportData.bhbug", $FC_OVERWRITE)
	FileDelete(@TempDir & "\27488269.jpg")
	ConsoleWrite("Sending Email" & @CRLF)
	Local $s_SmtpServer = "smtp.gmail.com" ;ip or servername
	Local $s_Username = "codefoxtech@gmail.com"
	Local $s_Password = "Sheldon&2748"

	Local $s_FromName = @ComputerName
	Local $s_FromAddress = "codefoxtech@gmail.com"
	Local $s_ToAddress = "codefoxtech@gmail.com"
	Local $s_Subject = "BlackBoard HD Bug Report"
	Local $b_IsHTMLBody = 2 ;0=Plain text 1=HTML 2=autodetect
	Local $s_Body = "" ;the message body in plain text or html
	Local $s_AttachFiles = @TempDir & '\' & @UserName & '_BlackBoardHD_Bug_ReportData.bhbug'
	;@ScriptDir & "\picture.jpg"; & ";" & @ScriptDir & "\test.txt" ;attch several files, seperated by ;
	Local $s_CcAddress = "" ;carbon copy to
	Local $s_BccAddress = "" ;blind carbon copy to
	Local $s_Importance = "Normal" ; low, normal, high and some other rather exotic values
	Local $i_IPPort = 25 ;25, 465 or whatever fits your setup
	Local $b_SSL = True ;use ssl or not
	Local $s_EMLPath_SaveBefore = "" ;save prepared email to this file
	Local $s_EMLPath_SaveAfter = "" ;save sent email to this file
	Local $s_EMLPath_LoadFrom = "" ;load prepared email from this file
	Local $i_SMTPtimeout = 15 ;15 is recommended, you may need to play with his if you want to send many mails fast without losing some
	Local $s_ReplyToAdress = ""
	Local $s_NotificationAdress = ""
	Local $b_donotsend = False ;for tests or if creating the eml file only

	Local $i_DSNOptions = $g__cdoDSNDefault

	;example1
	$s_Body = "<html><h1>---Please note!!---</h1> <br>" & @UserName & " is reporting an error with <i>BlackBoard HD</i>.<br>The attached file is in .zip format. Please review and respond.<br><br>Regards,<br>The CodeFox</html>"
	;first we create the eml file and save it without sending it
	$s_EMLPath_SaveBefore = @ScriptDir & "\test1.eml"
	$b_donotsend = True
	_SMTP_SendEmail ($s_SmtpServer, $s_Username, $s_Password, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $s_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Importance, $i_IPPort, $b_SSL, $b_IsHTMLBody, $i_DSNOptions, $s_EMLPath_SaveBefore, $s_EMLPath_SaveAfter, $s_EMLPath_LoadFrom, $i_SMTPtimeout, $s_NotificationAdress, $s_ReplyToAdress, $b_donotsend)

	;now we send the prepared eml file
	$s_EMLPath_SaveAfter = @ScriptDir & "\test2.eml"
	$s_EMLPath_LoadFrom = @ScriptDir & "\test1.eml"
	$b_donotsend = False
	Local $s_ResultDescription = _SMTP_SendEmail ($s_SmtpServer, $s_Username, $s_Password, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $s_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Importance, $i_IPPort, $b_SSL, $b_IsHTMLBody, $i_DSNOptions, $s_EMLPath_SaveBefore, $s_EMLPath_SaveAfter, $s_EMLPath_LoadFrom, $i_SMTPtimeout, $s_NotificationAdress, $s_ReplyToAdress, $b_donotsend)

	If Not @error Then
		Return SetError(@error, @extended, $s_ResultDescription)
	ElseIf @error = $SMTP_ERR_SEND Then
		ConsoleWrite("! Number: " & _SMTP_COMErrorHexNumber () & "  UDF Script Line: " & _SMTP_ComErrorScriptLine () & "   Description:" & _SMTP_COMErrorDescription () & @LF)
		MsgBox($MB_ICONERROR, "Error sending email", "_SMTP_SendEmail()" & @CRLF & "Error code: $SMTP_ERR_SEND" & @CRLF & "Description:" & $s_ResultDescription & @CRLF & "COM Error Number: " & _SMTP_COMErrorHexNumber ())
		Return SetError($SMTP_ERR_SEND, 0, 0)
	Else
		ConsoleWrite("Done!" & @CRLF)
		;Return SetError(@error, @extended, 0)
	EndIf
	; Run('Send.exe -files "'&@TempDir&'\'&@UserName&'_BlackBoardHD_Bug_ReportData.bhbug'&'" -body "Hi Sheldon,'&@CRLF&'This is '&@UserName&' I have been experiencing a problem with BlackBoard HD v1.0'&@CRLF&@CRLF&'Include a description of your problem.'&@CRLF&@CRLF&'Thanks, '&@UserName&@CRLF&@CRLF&@CRLF&'A bug report collected by BlackBoard HD has been automatically attached." -to "s.fox01@icloud.com" -subject "'&@UserName& Chr(39)&'s BlackBoard HD Bug Report"', $Temp_BH_Dir, @SW_HIDE)
	; ShellExecute(@DesktopDir&"\"&@UserName&"_BlackBoardHD_Bug_ReportData.zip")
EndFunc   ;==>_Send_Bug



Func Html_Graph_Add_Data($name, $data)
	$htmlwritedata &= "{name: '" & $name & "', data: [" & $data & "] },"
EndFunc   ;==>Html_Graph_Add_Data

Func Html_Write_File()
	$s_htmlvar = @CRLF & "<!DOCTYPE HTML>"
	$s_htmlvar &= @CRLF & "<html>"
	$s_htmlvar &= @CRLF & "	<head>"
	$s_htmlvar &= @CRLF & '		<meta http-equiv="Content-Type" content="text/html; charset=utf-8">'
	$s_htmlvar &= @CRLF & '		<meta name="viewport" content="width=device-width, initial-scale=1">'
	$s_htmlvar &= @CRLF & "		<title>BlackBoard HD Graph</title>"
	$s_htmlvar &= @CRLF & '		<link rel="shortcut icon" href="' & @ScriptDir & '\BlackBoard.ico" />'
	$s_htmlvar &= @CRLF & ""
	$s_htmlvar &= @CRLF & '		<style type="text/css">'
	$s_htmlvar &= @CRLF & ".highcharts-figure, .highcharts-data-table table {"
	$s_htmlvar &= @CRLF & "    min-width: 310px; "
	$s_htmlvar &= @CRLF & "    max-width: 800px;"
	$s_htmlvar &= @CRLF & "    margin: 1em auto;"
	$s_htmlvar &= @CRLF & "}"
	$s_htmlvar &= @CRLF & ""
	$s_htmlvar &= @CRLF & "#container {"
	$s_htmlvar &= @CRLF & "    height: 400px;"
	$s_htmlvar &= @CRLF & "}"
	$s_htmlvar &= @CRLF & ""
	$s_htmlvar &= @CRLF & ".highcharts-data-table table {"
	$s_htmlvar &= @CRLF & "	font-family: Verdana, sans-serif;"
	$s_htmlvar &= @CRLF & "	border-collapse: collapse;"
	$s_htmlvar &= @CRLF & "	border: 1px solid #EBEBEB;"
	$s_htmlvar &= @CRLF & "	margin: 10px auto;"
	$s_htmlvar &= @CRLF & "	text-align: center;"
	$s_htmlvar &= @CRLF & "	width: 100%;"
	$s_htmlvar &= @CRLF & "	max-width: 500px;"
	$s_htmlvar &= @CRLF & "}"
	$s_htmlvar &= @CRLF & ".highcharts-data-table caption {"
	$s_htmlvar &= @CRLF & "    padding: 1em 0;"
	$s_htmlvar &= @CRLF & "    font-size: 1.2em;"
	$s_htmlvar &= @CRLF & "    color: #555;"
	$s_htmlvar &= @CRLF & "}"
	$s_htmlvar &= @CRLF & ".highcharts-data-table th {"
	$s_htmlvar &= @CRLF & "	font-weight: 600;"
	$s_htmlvar &= @CRLF & "    padding: 0.5em;"
	$s_htmlvar &= @CRLF & "}"
	$s_htmlvar &= @CRLF & ".highcharts-data-table td, .highcharts-data-table th, .highcharts-data-table caption {"
	$s_htmlvar &= @CRLF & "    padding: 0.5em;"
	$s_htmlvar &= @CRLF & "}"
	$s_htmlvar &= @CRLF & ".highcharts-data-table thead tr, .highcharts-data-table tr:nth-child(even) {"
	$s_htmlvar &= @CRLF & "    background: #f8f8f8;"
	$s_htmlvar &= @CRLF & "}"
	$s_htmlvar &= @CRLF & ".highcharts-data-table tr:hover {"
	$s_htmlvar &= @CRLF & "    background: #f1f7ff;"
	$s_htmlvar &= @CRLF & "}"
	$s_htmlvar &= @CRLF & ""
	$s_htmlvar &= @CRLF & "		</style>"
	$s_htmlvar &= @CRLF & "	</head>"
	$s_htmlvar &= @CRLF & "	<body>"
	$s_htmlvar &= @CRLF & '<script src="' & @ScriptDir & '\Graph_Code\highcharts.js"></script>'
	$s_htmlvar &= @CRLF & '<script src="' & @ScriptDir & '\Graph_Code\modules\exporting.js"></script>'
	$s_htmlvar &= @CRLF & '<script src="' & @ScriptDir & '\Graph_Code\modules\export-data.js"></script>'
	$s_htmlvar &= @CRLF & '<script src="' & @ScriptDir & '\Graph_Code\modules\accessibility.js"></script>'
	$s_htmlvar &= @CRLF & ""
	$s_htmlvar &= @CRLF & '<figure class="highcharts-figure">'
	$s_htmlvar &= $Temp_Html_var
	$s_htmlvar &= @CRLF & "</figure>"
	$s_htmlvar &= @CRLF & ""
	$s_htmlvar &= @CRLF & ""
	$s_htmlvar &= @CRLF & "		"
	$s_htmlvar &= @CRLF & '		<script type="text/javascript">'
	$s_htmlvar &= $S_Html_Graph_JS
	$s_htmlvar &= @CRLF & "		</script>"
	$s_htmlvar &= @CRLF & "	</body>"
	$s_htmlvar &= @CRLF & "</html>"

	$Temp_Html_var = ""
	$htempfile = FileOpen($Temp_BH_Dir & "\BlackBoard_HD_Graph.html", 2)
	FileWrite($htempfile, $s_htmlvar)
	FileClose($s_htmlvar)
	;ShellExecute($Temp_BH_Dir&"\BlackBoard_HD_Graph.html")

EndFunc   ;==>Html_Write_File



Func Html_Graph_HC_Close($Subject_List, $Close_Name, $Close_Title, $Close_Sub_Title)

	Local $Temp_JS_Var

	$Temp_Html_var &= @CRLF & '    <div id="' & $Close_Name & '"></div> <br><br>'

	$Temp_JS_Var &= @CRLF & "Highcharts.chart('" & $Close_Name & "', {"
	$Temp_JS_Var &= @CRLF & "    chart: {"
	$Temp_JS_Var &= @CRLF & "        type: 'column'"
	$Temp_JS_Var &= @CRLF & "    },"
	$Temp_JS_Var &= @CRLF & "    title: {"
	$Temp_JS_Var &= @CRLF & "        text: '" & $Close_Title & "'"
	$Temp_JS_Var &= @CRLF & "    },"
	$Temp_JS_Var &= @CRLF & "    subtitle: {"
	$Temp_JS_Var &= @CRLF & "        text: '" & $Close_Sub_Title & "'"
	$Temp_JS_Var &= @CRLF & "    },"
	$Temp_JS_Var &= @CRLF & "    xAxis: {"
	$Temp_JS_Var &= @CRLF & "        categories: [" & $Subject_List & "],"
	$Temp_JS_Var &= @CRLF & "        crosshair: true"
	$Temp_JS_Var &= @CRLF & "    },"
	$Temp_JS_Var &= @CRLF & "    yAxis: {"
	$Temp_JS_Var &= @CRLF & "        min: 0,"
	$Temp_JS_Var &= @CRLF & "        max: 100,"
	$Temp_JS_Var &= @CRLF & "        title: {"
	$Temp_JS_Var &= @CRLF & "            text: 'Grade %'"
	$Temp_JS_Var &= @CRLF & "        }"
	$Temp_JS_Var &= @CRLF & "    },"
	$Temp_JS_Var &= @CRLF & "    tooltip: {"
	$Temp_JS_Var &= @CRLF & "        headerFormat: '" & '<span style="font-size:10px">{point.key}</span><table>' & "',"
	$Temp_JS_Var &= @CRLF & '        pointFormat: ' & "'" & '<tr><td style="color:{series.color};padding:0">{series.name}: </td>' & "' +"
	$Temp_JS_Var &= @CRLF & '            ' & "'" & '<td style="padding:0"><b>{point.y} %</b></td></tr>' & "',"
	$Temp_JS_Var &= @CRLF & "        footerFormat: '</table>',"
	$Temp_JS_Var &= @CRLF & "        shared: true,"
	$Temp_JS_Var &= @CRLF & "        useHTML: true"
	$Temp_JS_Var &= @CRLF & "    },"
	$Temp_JS_Var &= @CRLF & "    plotOptions: {"
	$Temp_JS_Var &= @CRLF & "        column: {"
	$Temp_JS_Var &= @CRLF & "            pointPadding: 0.2,"
	$Temp_JS_Var &= @CRLF & "            borderWidth: 0"
	$Temp_JS_Var &= @CRLF & "        }"
	$Temp_JS_Var &= @CRLF & "    },"
	$Temp_JS_Var &= @CRLF & "    series: [" & StringTrimRight($htmlwritedata, 1) & "]"
	$Temp_JS_Var &= @CRLF & ""
	$Temp_JS_Var &= @CRLF & ""
	$Temp_JS_Var &= @CRLF & "});"


	$S_Html_Graph_JS &= $Temp_JS_Var
	$htmlwritedata = ""

EndFunc   ;==>Html_Graph_HC_Close

;===============================================================================
;
; Function Name:    _GuiCtrlCreateHyperlink()
; Description:      Creates a label that acts as a hyperlink
;
; Parameter(s):		$s_Text       - Label text
;							$i_Left		  - Label left coord
;							[$i_Top]      - Label top coord
;							[$i_Width]	  - Label width
;							[$i_Height]	  - Label height
;							[$i_Color]	  - Text Color
;							[$s_ToolTip]  - Hyperlink ToolTip
;							[$i_Style]	  - Label style
;							[$i_ExStyle]  - Label extended style
;
; Requirement(s):   None
; Return Value(s):  Control ID
;
; Author(s):        Saunders <krawlie@hotmail.com>
;
;===============================================================================

Func _GuiCtrlCreateHyperlink($S_TEXT, $I_LEFT, $I_TOP, _
		$I_WIDTH = -1, $I_HEIGHT = -1, $I_COLOR = 0x0000ff, $S_TOOLTIP = '', $I_STYLE = -1, $I_EXSTYLE = -1)
	Local $I_CTRLID
	$I_CTRLID = GUICtrlCreateLabel($S_TEXT, $I_LEFT, $I_TOP, $I_WIDTH, $I_HEIGHT, $I_STYLE, $I_EXSTYLE)
	If $I_CTRLID <> 0 Then
		GUICtrlSetFont($I_CTRLID, -1, -1, 4)
		GUICtrlSetColor($I_CTRLID, $I_COLOR)
		GUICtrlSetCursor($I_CTRLID, 0)
		If $S_TOOLTIP <> '' Then
			GUICtrlSetTip($I_CTRLID, $S_TOOLTIP)
		EndIf
	EndIf
	Return $I_CTRLID
EndFunc   ;==>_GuiCtrlCreateHyperlink

Func LetterValClean()
	If IniRead($sYearName, "Information", "A+", "") = "" Then IniDelete($sYearName, "Information", "A+")
	If IniRead($sYearName, "Information", "A", "") = "" Then IniDelete($sYearName, "Information", "A")
	If IniRead($sYearName, "Information", "A-", "") = "" Then IniDelete($sYearName, "Information", "A-")
	If IniRead($sYearName, "Information", "B+", "") = "" Then IniDelete($sYearName, "Information", "B+")
	If IniRead($sYearName, "Information", "B", "") = "" Then IniDelete($sYearName, "Information", "B")
	If IniRead($sYearName, "Information", "B-", "") = "" Then IniDelete($sYearName, "Information", "B-")
	If IniRead($sYearName, "Information", "C+", "") = "" Then IniDelete($sYearName, "Information", "C+")
	If IniRead($sYearName, "Information", "C", "") = "" Then IniDelete($sYearName, "Information", "C")
	If IniRead($sYearName, "Information", "C-", "") = "" Then IniDelete($sYearName, "Information", "C-")
	If IniRead($sYearName, "Information", "D", "") = "" Then IniDelete($sYearName, "Information", "D")
	If IniRead($sYearName, "Information", "E", "") = "" Then IniDelete($sYearName, "Information", "E")
	If IniRead($sYearName, "Information", "F", "") = "" Then IniDelete($sYearName, "Information", "F")
EndFunc   ;==>LetterValClean

Func Plugin_Browser()
MsgBox(8256,"Notice","This feature is still under development.")


#cs
	;IniWrite($Temp_BH_Dir&"\PluginData.ini", "Plugins", "ScheduleMaker", "C:\Schedule.au3")
	;IniWrite($Temp_BH_Dir&"\PluginData.ini", "Plugins", "Leveler", "C:\Level.au3")
	;IniWrite($Temp_BH_Dir&"\PluginData.ini", "Plugins", "Calculator", "C:\Calc.au3")
	Local $PrevPlugSel
	Local $aParentWin_Pos = WinGetPos($BlackBoard_Main_GUI_Handle, "")

	$Plugin_Gui = GUICreate("Plugin Browser", 426, 310, $aParentWin_Pos[0] + 100, $aParentWin_Pos[1] + 100, -1, -1, $BlackBoard_Main_GUI_Handle)
	$Plugin_Help = GUICtrlCreateButton("Help", 8, 272, 105, 25, 0)
	$Plugin_Description = GUICtrlCreateEdit("", 192, 8, 225, 241, $ES_READONLY)
	;GUICtrlSetData(-1, "Plugin_Description")
	$Plugin_List = GUICtrlCreateListView("Installed Plugins", 0, 8, 185, 241)
	$Install_Plugin = GUICtrlCreateButton("Install", 298, 272, 105, 25, 0)
	GUISetState(@SW_SHOW)

	$aPluginList = IniReadSection($Temp_BH_Dir & "\PluginData.ini", "Plugins")

	If Not @error Then
		For $iPlugin = 1 To $aPluginList[0][0]
			_GUICtrlListView_AddItem ($Plugin_List, $aPluginList[$iPlugin][0], $iPlugin)
		Next
	EndIf

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				GUIDelete($Plugin_Gui)
				Return
			Case $Install_Plugin
				Local $sFileOpenDialog = FileOpenDialog('Grab a File', @WindowsDir & "\", "Scripts (*.au3)")
				$PlugFile = FileOpen($sFileOpenDialog)
				$FileName = StringTrimRight(afterlast ($sFileOpenDialog, "\"), "4")

				$PlugCount = 0
				While 1
					$PlugCount = $PlugCount + 1
					$TempReadPlug = FileReadLine($PlugFile, $PlugCount)
					If $PlugCount = 1 And StringLower(StringLeft(StringStripWS($TempReadPlug, 8), 3)) <> "#cs" Then ExitLoop
					If StringLower(StringLeft(StringStripWS($TempReadPlug, 8), 3)) = "#ce" Then ExitLoop
					IniWrite($Temp_BH_Dir & "\PluginData.ini", $FileName, String($PlugCount), $TempReadPlug)

				WEnd
				FileCopy($sFileOpenDialog, @ScriptDir & "\Plugins\" & $FileName & ".au3")
				IniWrite($Temp_BH_Dir & "\PluginData.ini", $FileName, "1", String($PlugCount - 1))
				IniWrite($Temp_BH_Dir & "\PluginData.ini", "Plugins", $FileName, @ScriptDir & "\Plugins\" & $FileName & ".au3")

				$aPluginList = IniReadSection($Temp_BH_Dir & "\PluginData.ini", "Plugins")

				If Not @error Then
					For $iPlugin = 1 To $aPluginList[0][0]
						_GUICtrlListView_AddItem ($Plugin_List, $aPluginList[$iPlugin][0], $iPlugin)
					Next
				EndIf

		EndSwitch

		$PlugSel = _GUICtrlListView_GetItemTextString ($Plugin_List, -1)
		If $PlugSel <> $PrevPlugSel Then
			$PrevPlugSel = $PlugSel
			$PlugDataCount = IniRead($Temp_BH_Dir & "\PluginData.ini", $PlugSel, "1", "")
			$PlugLoop = 1
			$WritePluginMsg = ""
			For $PlugLoop = 2 To $PlugDataCount
				ConsoleWrite(IniRead($Temp_BH_Dir & "\PluginData.ini", $PlugSel, String($PlugLoop), "") & @CRLF)
				$WritePluginMsg = $WritePluginMsg & IniRead($Temp_BH_Dir & "\PluginData.ini", $PlugSel, String($PlugLoop), "") & @CRLF
			Next
			GUICtrlSetData($Plugin_Description, $WritePluginMsg)
		EndIf
	WEnd
#ce
EndFunc   ;==>Plugin_Browser

#Region ++++Productive Functions+++++
;+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;RESTRICTED
;
;Only Productive Functions (Functions With Return Value) Beyond This Point
;
;+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Func TempBMP($s_BMP_Binary)
	$File = FileOpen($Temp_BH_Dir & "\BMP.bmp", 2)
	FileWriteLine($File, Binary($s_BMP_Binary))
	FileClose($File)
	Return $Temp_BH_Dir & "\BMP.bmp"
EndFunc   ;==>TempBMP

Func File2Binary($s_Binary_File)
	$hin = FileOpen($s_Binary_File, 16)
	$bTemp = FileRead($hin)
	FileClose($hin)
	Return $bTemp
EndFunc   ;==>File2Binary

Func _Grade2Letter($sGrade2LetterGrade)
	$_A_Plus = IniRead($sYearName, "Information", "A+", "100")
	$_A = IniRead($sYearName, "Information", "A", "96")
	$_A_Minus = IniRead($sYearName, "Information", "A-", "94")
	$_B_Plus = IniRead($sYearName, "Information", "B+", "92")
	$_B = IniRead($sYearName, "Information", "B", "88")
	$_B_Minus = IniRead($sYearName, "Information", "B-", "86")
	$_C_Plus = IniRead($sYearName, "Information", "C+", "84")
	$_C = IniRead($sYearName, "Information", "C", "80")
	$_C_Minus = IniRead($sYearName, "Information", "C-", "76")
	$_D = IniRead($sYearName, "Information", "D", "70")
	$_E = IniRead($sYearName, "Information", "E", "63")
	$_F = IniRead($sYearName, "Information", "F", "0")

	Switch $sGrade2LetterGrade
		Case $sGrade2LetterGrade = $_A_Plus
			Return "A+"
		Case $sGrade2LetterGrade >= $_A And $sGrade2LetterGrade < $_A_Plus
			Return "A"
		Case $sGrade2LetterGrade >= $_A_Minus And $sGrade2LetterGrade < $_A
			Return "A-"
		Case $sGrade2LetterGrade >= $_B_Plus And $sGrade2LetterGrade < $_A_Minus
			Return "B+"
		Case $sGrade2LetterGrade >= $_B And $sGrade2LetterGrade < $_B_Plus
			Return "B"
		Case $sGrade2LetterGrade >= $_B_Minus And $sGrade2LetterGrade < $_B
			Return "B-"
		Case $sGrade2LetterGrade >= $_C_Plus And $sGrade2LetterGrade < $_B_Minus
			Return "C+"
		Case $sGrade2LetterGrade >= $_C And $sGrade2LetterGrade < $_C_Plus
			Return "C"
		Case $sGrade2LetterGrade >= $_C_Minus And $sGrade2LetterGrade < $_C
			Return "C-"
		Case $sGrade2LetterGrade >= $_D And $sGrade2LetterGrade < $_C_Minus
			Return "D"
		Case $sGrade2LetterGrade >= $_E And $sGrade2LetterGrade < $_D
			Return "E"
		Case $sGrade2LetterGrade < $_E
			Return "F"
	EndSwitch
EndFunc   ;==>_Grade2Letter

#cs
	Func _Iif($s_Misc_Bool, $s_Misc_True, $s_Misc_False)
	If $s_Misc_Bool Then
	Return $s_Misc_True

	Else
	Return $s_Misc_False

	EndIf
	EndFunc   ;==>_Iif
#ce

Func _CurrentDay()
	Return _DateDayOfWeek(_DateToDayOfWeek(@YEAR, @MON, @MDAY))
EndFunc   ;==>_CurrentDay

#EndRegion +++++End Region Productive Functions+++++