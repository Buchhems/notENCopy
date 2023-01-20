#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=C:\Users\Sebastian\Downloads\Notencopy.ico
#AutoIt3Wrapper_Res_Comment=Erstellt von Sebastian Buch.
#AutoIt3Wrapper_Res_Description=Hilft bei der Eingabe der Noten in die LUS-Datenbank
#AutoIt3Wrapper_Res_Fileversion=1.23
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

; This will install the file jpg to the script location.
If $bFileInstall Then FileInstall("Notencopy.jpg", @ScriptDir & "\Notencopy.jpg")

#Region ### START Koda GUI section ### Form=
$Form1_1 = GUICreate("NotENCopy", 163, 298, 192, 124)
GUISetIcon("Notencopy.ico")
$Pic1 = GUICtrlCreatePic("Notencopy.jpg", 8, 8, 148, 113)
$Button1 = GUICtrlCreateButton("Hilfe", 8, 136, 147, 25)
GUICtrlSetFont(-1, 10, 400, 0, "Arial")
GUICtrlSetBkColor(-1, 0xC0DCC0)
$Button2 = GUICtrlCreateButton("mit ENC verbinden", 8, 168, 147, 33)
GUICtrlSetFont(-1, 10, 400, 0, "Arial")
$Button3 = GUICtrlCreateButton("Zwischenablage >> ENC", 8, 208, 147, 41)
GUICtrlSetFont(-1, 10, 400, 0, "Arial")
$Label1 = GUICtrlCreateLabel("buc @ hems.de    v1.24", 8, 272, 142, 17)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

GUICtrlSetState($Button1, $GUI_ENABLE)
GUICtrlSetState($Button2, $GUI_ENABLE)
GUICtrlSetState($Button3, $GUI_DISABLE) ;Solange der ENC nicht gestartet wurde, ausgrauen!
    
While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $Button1
			hilfe()
		Case $Button2
			encStart()
		Case $Button3
			ausZAholen()
			eintragen2D()
	EndSwitch
WEnd

Func hilfe()
MsgBox(32, "NotENCopy Hilfe", "NotENCopy ist ein Tool, dass Notenlisten aus der Zwischenablage (aus Excel oder Libreoffice kopiert) direkt in den externen Notenclient (ENC) der Lehrer- und Schülerdatenbank (LUSD) des Landes Hessen kopieren kann." & @CRLF &  @CRLF &"1. Zunächst klickt man auf 'mit ENC verbinden', startet einmalig den ENC und loggt sich ein. Dazu wird das Passwort eingegeben und die XML-Datei ausgewählt werden." & @CRLF & "2. Jetzt kopiert man die Notenspalte(n) und klickt in NotENCopy auf 'Zwischenablage >> ENC'. Diesen zweiten Schritt kann man beliebig oft durchführen, da die Verbindung zum ENC nur einmalig (pro Lehrkraft/Klasse) hergestellt werden muss.")
EndFunc

Func encStart()
	MsgBox(64, "ENC starten", "Bitte ENC starten und einloggen!" & @CRLF & "1. Accountdaten liegen im PDF" & @CRLF & "2. Richtige XML-Datei auswählen!")

	WinWaitActive("LUSD - Externe Notenerfassung", "Klassenweise Erfassung")		;Warten bis eingeloggt
	$fnameLUSD = WinGetHandle("LUSD - Externe Notenerfassung", "")
		
	MsgBox(64, "ENC aktiv", "ENC aktiv." & @CRLF &"Bitte in das gewünschte Startfeld/Zelle des ENC klicken", 5)
	GUICtrlSetState($Button3, $GUI_ENABLE)
	GUICtrlSetState($Button2, $GUI_DISABLE) ; ENC ist verbunden
	GUICtrlSetData($Button2, "mit ENC verbunden!")
EndFunc

Func ausZAholen()
	MsgBox(64, "Zwischenablage in ENC kopieren", "Zwischenablage in ENC kopieren"  & @CRLF & @CRLF & "Falls noch nicht geschehen:" & @CRLF & "Vor dem klicken auf OK mit [Strg+C] Spalte(n) aus Excel/LibreOffice in die Zwischenablage kopieren.")
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
	lusdAktiv()
	Sleep(500)
	If $iHilfe = 7 Then
		MsgBox(64, "OK", "Bitte ins 1.Feld klicken!" & @CRLF & "Notenpunkte müssen ZWEISTELLIG angegeben werden!")
		Sleep(2000)
	EndIf

	For $i = 0 To $iAnzSchue-1 Step 1
		lusdAktiv()
		Send($iNoten[$i])
		lusdAktiv()
		Send("{ENTER}")
		For $j = 0 to $iAnzSpalte-2 Step 1
			lusdAktiv()
			Send("{LEFT}")
			Sleep(100)
		Next
		Sleep(300)
	Next
EndFunc

Func lusdAktiv()
	If NOT WinActive($fnameLUSD) Then ; Check ob LUSD gerade auf ist und direkt sichtbar.
			WinActivate($fnameLUSD)
			Sleep(100)
	EndIf
EndFunc
