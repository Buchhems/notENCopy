#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=C:\Users\buc\Downloads\Notencopy.ico
#AutoIt3Wrapper_Res_Comment=Erstellt von Sebastian Buch. Kontakt: buch@hems.de
#AutoIt3Wrapper_Res_Description=Hilft bei der Eingabe der Noten in die LUS-Datenbank
#AutoIt3Wrapper_Res_Fileversion=1.19
#AutoIt3Wrapper_Res_LegalCopyright=Benutzung auf eigene Gefahr.Software ohne Gewähr
#AutoIt3Wrapper_Res_Language=1031
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <array.au3>
#include <Excel.au3>

Global $iHilfe		;hilfe-Flag
Global $iNoten		;Notentabelle
Global $iAnzSchue   ;Anzahl Schueler in Klasse, wird später festgelegt für **Reihenanzahl**
Global $iAnzSpalte	;Anzahl Fächer / Einträge  / **Spalten** die übernommen werden in ENC
Global $fnameXL		;Fensternamen für Excel
Global $fnameLUSD	;und ENC

Local $bFileInstall = True ; Change to True and ammend the file paths accordingly.

; This will install the file C:\Test.bmp to the script location.
If $bFileInstall Then FileInstall("C:\Users\buc\Downloads\Notencopy.jpg", @ScriptDir & "\Notencopy.jpg")

#Region ### START Koda GUI section ### Form=
$Form1_1 = GUICreate("NotENCopy", 163, 298, 192, 124)
GUISetIcon("C:\Users\buc\Downloads\Notencopy.ico")
$Pic1 = GUICtrlCreatePic(@ScriptDir & "\Notencopy.jpg", 8, 8, 148, 113)
$Button1 = GUICtrlCreateButton("Anleitung", 8, 136, 147, 25)
GUICtrlSetFont(-1, 10, 400, 0, "Arial")
GUICtrlSetBkColor(-1, 0xC0DCC0)
$Button2 = GUICtrlCreateButton("Excel => ENC", 8, 168, 147, 33)
GUICtrlSetFont(-1, 10, 400, 0, "Arial")
$Button3 = GUICtrlCreateButton("Zwischenablage => ENC", 8, 208, 147, 41)
GUICtrlSetFont(-1, 10, 400, 0, "Arial")
$Label1 = GUICtrlCreateLabel("buch@hems.de               v1.19", 8, 272, 142, 17)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

	GUICtrlSetState($Button1, $GUI_DISABLE)
	GUICtrlSetState($Button2, $GUI_DISABLE)
    GUICtrlSetState($Button3, $GUI_DISABLE) ;Solange die Dateien nicht gestartet wurden, ausgrauen!
MsgBox(64, "OK", "Bitte ENC starten und einloggen!" & @CRLF & "(Accountdaten im PDF. XML-Datei nicht vergessen!)")

	WinWaitActive("LUSD - Externe Notenerfassung", "Klassenweise Erfassung")		;Warten bis eingeloggt
	$fnameLUSD = WinGetHandle("LUSD - Externe Notenerfassung", "")
	WinActivate($fnameLUSD)

	MsgBox(64, "OK", "ENC aktiv. Bitte in erstes Feld klicken", 3)
	GUICtrlSetState($Button1, $GUI_ENABLE)
	GUICtrlSetState($Button3, $GUI_ENABLE)

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $Button1
			hilfe()
		Case $Button2
			ausExcelholen()
			;eintragen()
		Case $Button3
			ausZAholen()
			eintragen2D()
	EndSwitch
WEnd

Func hilfe()
#Region ### START Koda GUI section ### Form=
$Form2 = GUICreate("Form2", 347, 144, 165, 192)
$Label2 = GUICtrlCreateLabel("NotENCopy", 96, 16, 156, 36)
GUICtrlSetFont(-1, 20, 400, 0, "Arial")
$Label3 = GUICtrlCreateLabel("von buch@hems.de", 112, 104, 109, 17, $WS_BORDER)
$Label4 = GUICtrlCreateLabel("Dieses Tool ermöglicht es Noten aus Excel o. LibreOffice", 0, 56, 341, 14)
GUICtrlSetFont(-1, 10, 400, 0, "Arial")
$Label5 = GUICtrlCreateLabel("oder der Zwischenablage in den ENC einzutragen,", 16, 72, 317, 22)
GUICtrlSetFont(-1, 10, 400, 0, "Arial")
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			GUIDelete($Form2)
	EndSwitch
WEnd
EndFunc

Func ausExcelholen()
	;;Buggy. Ich bekomme es einfach nicht hin die Daten aus excel abzuholen, da ja schon eine Datei läuft.
	;;Vielleicht über Dateipfad einlesen versuchen?

	MsgBox(64, "OK", "Bitte Excel-Datei starten!" & @CRLF & "Es darf nur eine Excel-Datei mit dem Namen 'FOS' offen sein!")
	WinWaitActive("[CLASS:XLMAIN]", "FOS")
	$fnameXL = WinGetTitle("[CLASS:XLMAIN]", "FOS") 	; $fnameXL ist Verweis auf Excel-Fenster

	$oExcel = _Excel_BookAttach("FOS", $fnameXL)
	If @error = 1 Then MsgBox(0, "Error code is: "&@error, "Could not Attach Excel Book returning: " &$oExcel)

	MsgBox(64, "OK", "blablabla")
	;$oExcel.Sheets("DE").Activate
	;$oExcel.worksheets("DE").select	; select info tab
	;$LastRow = $oExcel.ActiveSheet.UsedRange.Rows.Count               ; determine last row

	$mydata = _Excel_RangeRead($oExcel)   ; read range from A1 to AI;  Anzahl Schüler
	;	If IsArray($mydata) Then
	_ArrayDisplay($mydata)                  ; display spreadsheet data in the array
	;_Excel_BookClose($datawb,True)
	MsgBox(64, "ENDE","ENDE")
EndFunc

Func ausZAholen()
	Local $sClipboard = clipget()												;Hole dir Daten aus der Zwischenablage

	$iNoten = StringSplit(StringReplace($sClipboard, @CRLF, " "), " ")			;Array für alle Werte des jew. Faches
	;_ArrayDisplay($iNoten)
	_ArrayDelete($iNoten, 0)		;;Erster Eintrag wird nicht benötigt.
	_ArrayDelete($iNoten, UBound($iNoten)-1) ;;letztes Element enthält Müll

	$instr = $iNoten[0]
	$tempstr = StringReplace($instr,'	','')
	$iAnzSpalte = StringLen($instr) - StringLen($tempstr) + 1		;Hack um die Spaltenanzahl rauszubekommen

	;_ArrayDisplay($iNoten)		;DEBUG
	 $iAnzSchue = UBound($iNoten)
EndFunc

Func eintragen2D()
	WinActivate($fnameLUSD)
	Sleep(500)

	If $iHilfe = 7 Then
		MsgBox(64, "OK", "Bitte ins 1.Feld klicken!" & @CRLF & "Notenpunkte müssen ZWEISTELLIG angegeben werden!")
		Sleep(2000)
	EndIf

	For $i = 0 To $iAnzSchue-1 Step 1
			Send($iNoten[$i])
			Send("{ENTER}")
		For $j = 0 to $iAnzSpalte-2 Step 1
			Send("{LEFT}")
			Sleep(100)
		Next
		Sleep(300)
	Next
EndFunc
