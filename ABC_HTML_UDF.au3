

Global $s_Html_RowCount
Global $s_Html_Style=".cell{border:.5pt solid;border-top:none;border-left:none}"
Global $i_Html_Misc
Global $i_Html_Pic_Handle
Global $a_Html_List[30]
Global $i_Html_List


Func _Html_Init($hABC_Html_Name)

   Local $sTempFileWrite

  $hABC_Html_Handle = FileOpen($hABC_Html_Name, 2)

  $sTempFileWrite = '<HTML>'&@CR _
  &'<HEAD>'&@CR _
  &'<META HTTP-EQUIV="CONTENT-TYPE" CONTENT="text/html; charset=windows-1252">'&@CR _
  &'<META NAME="GENERATOR" CONTENT="ABC Tech Html Generator For AutoIt">'&@CR _
  &'<META NAME="AUTHOR" CONTENT="@FoxSoft">'&@CR _
  &'</HEAD>'

  FileWriteLine($hABC_Html_Handle, $sTempFileWrite)

  Return $hABC_Html_Handle

EndFunc


Func _Html_Internal($HtmlHandle, $s_InternalLink_Name, $s__Html_Link_Target)
   FileWriteLine($HtmlHandle, '<a class="link" href="'&$s__Html_Link_Target&'">'&$s_InternalLink_Name&'</a>')
EndFunc


Func _Html_Internal_Link($HtmlHandle, $s_InternalLink_Name)
   FileWriteLine($HtmlHandle, '<A name="'&$s_InternalLink_Name&'">')
EndFunc

Func _Html_Set_Title($HtmlHandle, $Html_Title)
   FileWriteLine($HtmlHandle, "<title>"&$Html_Title&"</title>")
EndFunc

Func _Html_Close($HtmlHandle)

  FileWriteLine($HtmlHandle, '<style>')

  FileWriteLine($HtmlHandle, $s_Html_Style)
  FileWriteLine($HtmlHandle, '</style>')

  FileWriteLine($HtmlHandle, '</HTML>')

  FileClose($HtmlHandle)
  $s_Html_Style=".cell{border:.5pt solid;border-top:none;border-left:none}"
  $i_Html_List = 0
EndFunc

Func _Html_AddList()
   $i_Html_List = $i_Html_List + 1
    $a_Html_List[$i_Html_List] = '<ul>'
	Return $i_Html_List
 EndFunc

Func _Html_AddListItem($i_Html_List_Handle, $sHtmlListItem, $s_Html_List_Link_Target="")
   If $s_Html_List_Link_Target > "" Then
   $a_Html_List[$i_Html_List_Handle] = $a_Html_List[$i_Html_List_Handle]&@CR&'<li class=list'&$i_Html_List_Handle&'><a href="#'&$s_Html_List_Link_Target&'">'&$sHtmlListItem&'</a></li>'
   Else
   $a_Html_List[$i_Html_List_Handle] = $a_Html_List[$i_Html_List_Handle]&@CR&'<li class=list'&$i_Html_List_Handle&'>'&$sHtmlListItem&'</li>'
   EndIf
EndFunc

Func _Html_Close_List($HtmlHandle, $i_Html_List_Handle)
   $a_Html_List[$i_Html_List_Handle] = $a_Html_List[$i_Html_List_Handle]&@CR&"</ul>"
   FileWriteLine($HtmlHandle, $a_Html_List[$i_Html_List_Handle])
EndFunc


Func _Html_AddDiv($HtmlHandle, $sHtmlNotes="align=center")
   FileWriteLine($HtmlHandle, "<div "&$sHtmlNotes&">")
EndFunc

Func _Html_CloseDiv($HtmlHandle)
   FileWriteLine($HtmlHandle, "</div>")
EndFunc

Func _Html_AddPic($HtmlHandle, $HtmlPicFile, $HtmlPicName="")
   $i_Html_Pic_Handle = $i_Html_Pic_Handle+1
   FileWriteLine($HtmlHandle, '<img id = htmlpic'&$i_Html_Pic_Handle&' src="'&$HtmlPicFile&'"  alt="'&$HtmlPicName&'"/>')
   Return "#htmlpic"&$i_Html_Pic_Handle
EndFunc

Func _Html_NewLine($HtmlHandle, $iHtml_NewLine_Count=1)
   Local $iHtml_I
   For $iHtml_I = 1 To $iHtml_NewLine_Count
	  FileWriteLine($HtmlHandle, "<br></br>")
   Next
EndFunc

Func _Html_SendCSS($h_Html_Handle, $s_Html_Command)
   $s_Html_Style = $s_Html_Style&@CR&$h_Html_Handle&'{'&$s_Html_Command&'}'
EndFunc




#cs
   Table Functions

#ce

#Region +++++Table Functions++++++



Func _Html_InitTable($HtmlHandle, $sCaption)

   $i_Html_Misc = $i_Html_Misc + 1
   Local $sTempFileWrite
   Local $s_Html_id = "table"&$i_Html_Misc

   $sTempFileWrite = '<TABLE Align=center FRAME=VOID CELLSPACING=0 RULES=NONE BORDER=0>'&@CR _
   &'<caption id = "'&$s_Html_id&'">'&$sCaption&'</caption>'&@CR _

   FileWriteLine($HtmlHandle, $sTempFileWrite)

   Return "#table"&$i_Html_Misc
EndFunc


Func _Html_AddRow($HtmlHandle, $s_Html_Table_Text)
   Local $i = 0
   Local $sTempFileWrite
   $s_Html_RowCount = $s_Html_RowCount + 1

   $a_Html_Exploded = StringSplit($s_Html_Table_Text, "|")

   FileWriteLine($HtmlHandle, '<tr class = body id = "row'&$s_Html_RowCount&'">')
   For $i = 1 To $a_Html_Exploded[0]
	  FileWriteLine($HtmlHandle, "<td class=cell>"&$a_Html_Exploded[$i]&"</td>")
   Next

   FileWriteLine($HtmlHandle, "</tr>")

   Return "#row"&$s_Html_RowCount

EndFunc



Func _Html_CloseTable($HtmlHandle)
   FileWriteLine($HtmlHandle, "</table>")
   FileWriteLine($HtmlHandle, "<br></br>")
EndFunc



Func _Html_SetRowColor($h_Html_Handle, $s_Html_Color)
   $s_Html_Style = $s_Html_Style&@CR&$h_Html_Handle&'{background: '&$s_Html_Color&';}'
EndFunc

Func _Html_AddColumn($HtmlHandle, $s_Html_Width, $s_Html_ColNum=1)
   FileWriteLine($HtmlHandle, "<col width="&$s_Html_Width&" span="&$s_Html_ColNum&">")
EndFunc


Func _Html_Table_SetCaptionFont($h_Html_TableHandle, $s_Html_FontSize="8.5", $s_Html_FontWeight="Bold", $s_Html_FontName="MS Sans Serif")
   $s_Html_Style = $s_Html_Style&@CR&$h_Html_TableHandle&'{font-size:'&$s_Html_FontSize&'px;font-weight:'&$s_Html_FontWeight&';font-family:"'&$s_Html_FontName&'";}'
EndFunc


Func _Html_Table_SetBodyFont($h_Html_TableHandle, $s_Html_FontSize="8.5", $s_Html_FontWeight="Bold", $s_Html_FontName="MS Sans Serif")
   $s_Html_Style = $s_Html_Style&@CR&'.body{font-size:'&$s_Html_FontSize&'px;font-weight:'&$s_Html_FontWeight&';font-family:"'&$s_Html_FontName&'";}'
EndFunc

#EndRegion +++++Table Functions+++++
