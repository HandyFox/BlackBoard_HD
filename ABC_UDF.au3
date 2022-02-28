#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         Sheldon Fox (ABC Technologies)

 Script Function:
	Collection of ABC Technologies Reusable Code Snippets

#ce ----------------------------------------------------------------------------



Func _ABC_StringTrim($inputString, $sCharacters)

   If StringLeft($inputString, 1) = $sCharacters Then
	  $inputString = StringTrimLeft($inputString, 1)
   EndIf

   If StringRight($inputString, 1) = $sCharacters Then
	  $inputString = StringTrimRight($inputString, 1)
   EndIf

   Return $inputString

EndFunc

Func afterlast($source, $match)

  $j = StringInStr($source, $match, 0, -1)
if $j then Return StringMid($source,$j+Stringlen($match))
EndFunc



