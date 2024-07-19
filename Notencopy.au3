#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=C:\Users\Sebastian\Downloads\Notencopy.ico
#AutoIt3Wrapper_Res_Comment=Erstellt von Sebastian Buch.
#AutoIt3Wrapper_Res_Description=Hilft bei der Eingabe der Noten in die LUS-Datenbank
#AutoIt3Wrapper_Res_Fileversion=1.25
#AutoIt3Wrapper_Res_LegalCopyright=Benutzung auf eigene Gefahr.Software ohne Gewähr
#AutoIt3Wrapper_Res_Language=1031
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <array.au3>
#include <Excel.au3>

Global $iNoten		;Notentabelle
Global $iAnzSchue   ;Anzahl Schueler in Klasse, wird später festgelegt für **Reihenanzahl**
Global $iAnzSpalte	;Anzahl Fächer / Einträge  / **Spalten** die übernommen werden in ENC
Global $fnameXL		;Fensternamen für Excel
Global $fnameLUSD	;und ENC
Global $Response 		;Abbruch Kopieraktion

Local $bFileInstall = True ; Change to True and ammend the file paths accordingly.

; This will install the file jpg to the script location.
If $bFileInstall Then FileInstall("Notencopy.jpg", @ScriptDir & "\Notencopy.jpg")

#Region ### START Koda GUI section ### Form=
$Form1_1 = GUICreate("NotENCopy", 163, 260, 192, 124)
GUISetIcon("Notencopy.ico")
$Pic1 = GUICtrlCreatePic("Notencopy.jpg", 8, 8, 148, 113)
$Button1 = GUICtrlCreateButton("Hilfe", 8, 208, 150, 25)
GUICtrlSetFont(-1, 10, 400, 0, "Arial")
GUICtrlSetBkColor(-1, 0xC0DCC0)
$Button2 = GUICtrlCreateButton("Mit ENC verbinden", 8, 128, 150, 33)
GUICtrlSetFont(-1, 10, 400, 0, "Arial")
$Button3 = GUICtrlCreateButton("Zwischenablage >> ENC", 8, 160, 150, 41)
GUICtrlSetFont(-1, 10, 400, 0, "Arial")
$Label1 = GUICtrlCreateLabel("buc @ hems.de                v1.25", 10, 240, 150, 17)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

GUICtrlSetState($Button1, $GUI_ENABLE)
GUICtrlSetState($Button2, $GUI_ENABLE)
GUICtrlSetState($Button3, $GUI_DISABLE) 													;Solange der ENC nicht gestartet wurde, ausgrauen!
    
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
MsgBox(32, "Hilfe", "NotENCopy ist ein Tool, dass Notenlisten aus Excel/Zwischenablage direkt in den externen Notenclient (ENC) der LUSD des Landes Hessen kopieren kann." & @CRLF &  @CRLF &"1. Zunächst klickt man auf 'Mit ENC verbinden' und folgt den Anweisungen." & @CRLF & "2. Jetzt kopiert man die Notenspalte(n) aus Excel und klickt in NotENCopy auf 'Zwischenablage >> ENC'." & @CRLF &  @CRLF &"Diesen zweiten Schritt kann man beliebig oft durchführen, da die Verbindung zum ENC nur einmalig (pro Lehrkraft/Klasse) hergestellt werden muss." & @CRLF &  @CRLF &"Häufige Fehler:" & @CRLF &"Zeilen überspringen: Übereinstimmung Anzahl Spalten Excel und ENC kontrollieren." & @CRLF &"Verrückte Spalte: Einträge in Excel kontrollieren (Zweistellige Note).")
EndFunc

Func encStart()
	MsgBox(64, "ENC starten", "Bitte ENC starten und einloggen!" & @CRLF & "1. Accountdaten liegen im PDF" & @CRLF & "2. Richtige XML-Datei auswählen!")

	If Not WinWaitActive("LUSD - Externe Notenerfassung", "Klassenweise Erfassung", 20) Then
		MsgBox(48, "Fehler", "ENC konnte gefunden werden.")
		Return
	EndIf
	
	$fnameLUSD = WinGetHandle("LUSD - Externe Notenerfassung", "")
		
	MsgBox(64, "MIT ENC VERBUNDEN", "Mit ENC verbunden!", 5)
	GUICtrlSetState($Button3, $GUI_ENABLE)
	GUICtrlSetState($Button2, $GUI_DISABLE) 												; ENC ist verbunden
	GUICtrlSetData($Button2, "MIT ENC VERBUNDEN!")
EndFunc

Func ausZAholen()
	$Response = MsgBox(4 + 32, "Pre-flight Checklist", "1. GEWÜNSCHTE STARTZELLE des ENC ausgewählt?" & @CRLF & @CRLF & "2. Noten/Fehlzeiten aus Excel in die ZWISCHENABLAGE kopiert?" & @CRLF & @CRLF & "3. Falls BG/FOS: Liegen die Noten in ZWEISTELLIGER ZIFFERNFORM vor (01,02,...)?")
	
	If $Response = 7 Then
		MsgBox(0, "Abbruch", "Sie waren wohl noch nicht so weit....")
		Return
	Else
																												; Der Benutzer hat "Abbrechen" geklickt
		Local $sClipboard = ClipGet()																;Hole dir Daten aus der Zwischenablage
		If @error Then
			MsgBox(48, "Fehler", "Konnte nicht auf die Zwischenablage zugreifen.")
			Return
		EndIf	
	
		$iNoten = StringSplit(StringReplace($sClipboard, @CRLF, " "), " ")			;Array für alle Werte des jew. Faches
		;_ArrayDisplay($iNoten)
		_ArrayDelete($iNoten, 0)																	;Erster Eintrag wird nicht benötigt, bzw. hat keine relevanten Daten
																												;_ArrayDelete($iNoten, UBound($iNoten)-1) ;;letztes Element enthält Müll

		$instr = $iNoten[0]
		$tempstr = StringReplace($instr,'	','')
		$iAnzSpalte = StringLen($instr) - StringLen($tempstr) + 1					;Hack um die Spaltenanzahl rauszubekommen
		;_ArrayDisplay($iNoten)		;DEBUG
		$iAnzSchue = UBound($iNoten)
	EndIf
EndFunc

Func eintragen2D()
	If  $Response = 6 Then
		lusdAktiv()
		Sleep(500)
		
		For $i = 0 To $iAnzSchue-1 Step 1
			lusdAktiv()
			Send($iNoten[$i])
			lusdAktiv()
			Send("{ENTER}")
			For $j = 0 To $iAnzSpalte-2 Step 1
				lusdAktiv()
				Send("{LEFT}")
				Sleep(100)
			Next
			Sleep(300)
		Next
	EndIf
EndFunc

Func lusdAktiv()
	If NOT WinActive($fnameLUSD) Then 														; Check ob LUSD gerade auf ist und direkt sichtbar.
			WinActivate($fnameLUSD)
			Sleep(100)
	EndIf
EndFunc
