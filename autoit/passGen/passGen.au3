#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=PassGen.ico
#AutoIt3Wrapper_Outfile=PassGen.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Description=Passwort-Generator
#AutoIt3Wrapper_Res_Fileversion=1.0.0.2
#AutoIt3Wrapper_Res_LegalCopyright=Thomas Stephan (oscar at elektronik-kurs dot de)
#AutoIt3Wrapper_Res_Language=1031
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Allow_Decompile=n
#include<GuiSlider.au3>
#include<Misc.au3>
#include<String.au3>
#include<ButtonConstants.au3>
#include<ComboConstants.au3>
#include<EditConstants.au3>
#include<GUIConstantsEx.au3>
#include<SliderConstants.au3>
#include<StaticConstants.au3>
#include <GuiImageList.au3>
#include <GuiButton.au3>

Global $title = 'PassGen' ; Fenstertitel


If _Singleton($title, 1) = 0 Then ; testen, ob das Programm bereits läuft
	MsgBox(48, $title, 'Das Programm läuft bereits!') ; Benutzer informieren
	Exit ; zweite Instanz beenden
EndIf
Opt('TrayMenuMode', 1) ; Tray-Standardmenü ausschalten
Opt('TrayAutoPause', 0) ; AutoPause ausschalten


Global $key[5] ; Array deklarieren (5 = 0...4)
$key[0] = '1234567890' ; Level 1 = nur Zahlen
$key[1] = _StringRepeat($key[0], 4) & 'abcdefghijklmnopqrstuvwxyz' ; Level 2 = Zahlen (4fach um den prozentualen Anteil zu erhöhen) und Buchstaben (nur Kleinschreibung)
$key[2] = $key[1] & 'ABCDEFGHIJKLMNOPQRSTUVWXYZ' ; Level 3 = Zahlen und Buchstaben (gemischte Groß-/Kleinschreibung)
$key[3] = $key[2] & '@!$%&/()=<>|,.-;:_#+*~?\' & Chr(34) & Chr(39) ; Level 5 = zusätzlich alle übrigen Sonderzeichen einer deutschen Tastatur
$key[4] = $key[3] & 'öäüÖÄÜß' ; Level 4 = zusätzlich Umlaute

Global $GUI = GUICreate($title, 640, 180, -1, -1) ; Fenster erstellen

GUISetIcon(@SystemDir & '\shell32.dll', -212) ; Icon in der Titelleiste setzen
Global $TrayShow  = TrayCreateItem('Passwort-Generator anzeigen') ; Tray-Menü erstellen
Global $TrayExit  = TrayCreateItem('Beenden') ; Tray-Menü erstellen

GUICtrlCreateGroup('Complexity', 10, 8, 130, 55) ; Gruppe 'Stärke' erstellen
Global $Combo = GUICtrlCreateCombo ('', 20, 30, 110, 20, $CBS_DROPDOWNLIST) ; Auswahlbox für die Passwort-Stärke
GUICtrlSetData(-1, '1 = Leicht|2 = Mittel|3 = Hoch|4 = Sehr Hoch', '2 = Mittel') ; Werte der Auswahlbox
GUICtrlCreateGroup ('', -99, -99, 1, 1) ; Ende der Grupppe

GUICtrlCreateGroup('Passwordlength', 150, 8, 480, 55) ; Gruppe 'Passwortlänge' erstellen
Global $Slider = GUICtrlCreateSlider(160, 30, 400, 20, BitOR($TBS_TOOLTIPS, $TBS_AUTOTICKS)) ; Slider für Passwortlänge
GUICtrlSetLimit(-1, 36, 3) ; Slider auf Werte zwischen 3 und 36 limitieren
_GUICtrlSlider_SetTicFreq($Slider, 3) ; alle 3 Stellen ein Strich setzen
_GUICtrlSlider_SetPageSize($Slider, 3) ; bei Klick vor oder hinter dem aktuellen Wert, um 3 Stellen weiterspringen
_GUICtrlSlider_SetPos($Slider, 16) ; Slider auf den Wert 16 setzen
Global $Range = GUICtrlCreateLabel('16', 580, 30, 40, 20, $SS_CENTER) ; Die Anzeige für die Passwortlänge (Startwert=18)
GUICtrlSetFont(-1, 12, 400, 0, 'Arial') ; Schriftart und -größe der Anzeige
;GUICtrlSetBKColor(-1, 0x000000) ; Hintergrundfarbe der Anzeige
;GUICtrlSetColor(-1, 0xFFFF00) ; Vordergrundfarbe der Anzeige
GUICtrlCreateGroup ('', -99, -99, 1, 1) ; Ende der Grupppe

Global $Button = GUICtrlCreateButton('PassGen', 12, 70, 90, 25, $BS_DEFPUSHBUTTON) ; Startbutton erstellen
GUICtrlSetFont(-1, 10, 400, 0, 'Arial') ; Schriftart und -größe des Buttons

If(@OSVersion = "WIN_VISTA" OR @OSVersion = "WIN_7" OR @OSVersion = "WIN_8" OR @OSVersion = "WIN_2008" OR @OSVersion = "WIN_2008R2") Then
	_GUICtrlButton_SetImageList($Button, _GetImageListHandle(@SystemDir & "\shell32.dll", 238))
EndIf

GUICtrlCreateGroup('Passwort', 10, 102, 620, 65) ; Gruppe 'Passwort' erstellen
Global $Code = GUICtrlCreateInput('', 20, 120, 600, 40, $ES_READONLY) ; Ausgabefeld für das Passwort
GUICtrlSetFont(-1, 20, 400, 0, 'Courier New Bold') ; Schriftart und -größe des Ausgabefeldes
GUICtrlCreateGroup ('', -99, -99, 1, 1) ; Ende der Grupppe

;GUICtrlSetState($Code, $GUI_FOCUS) ; den Focus auf das Ausgabefeld, damit man das Passwort einfach per Kontextmenü ('Kopieren') in die Zwischenablage packen kann
Passwort() ; Funktion aufrufen, damit beim Programmstart bereits ein Passwort angezeigt wird
GUISetState() ; GUI-Fenster anzeigen

While 1 ; MessageLoop-Schleife
	Switch GUIGetMsg() ; Anhand des eingetretenen GUI-Ereignisses die entsprechenden Befehle ausführen
		Case $Combo ; Benutzer hat die Passwort-Stärke geändert
			Passwort() ; Funktion Passwort() aufrufen
		Case $Slider ; Benutzer hat den Slider bewegt
			GUICtrlSetData($Range, GUICtrlRead($Slider)) ; Sliderwert auslesen und in das Anzeigefeld schreiben
			Passwort() ; Funktion Passwort() aufrufen
		Case $Button ; Benutzer hat auf den 'Start'-Button geklickt
			Passwort() ; Funktion Passwort() aufrufen
		Case $GUI_EVENT_MINIMIZE ; Benutzer hat auf Minimieren geklickt
			Opt('TrayIconHide', 0) ; Tray-Menü anzeigen
			TraySetIcon(@SystemDir & '\shell32.dll', -212) ; Icon für Tray-Menü setzen
			GUISetState(@SW_HIDE, $GUI) ; Fenster verstecken
			While 2 ; Tray-Menü-Schleife
				Switch TrayGetMsg() ; Anhand des eingetretenen Tray-Ereignisses die entsprechenden Befehle ausführen
					Case $TrayShow ; wurde 'Anzeigen' aufgerufen, dann...
						Opt('TrayIconHide', 1) ; Tray-Menü wieder verstecken
						GUISetState(@SW_SHOW, $GUI) ; Fenster anzeigen
						GUISetState(@SW_RESTORE, $GUI) ; und wiederherstellen (minimieren rückgängig machen)
						ExitLoop ; Tray-Menü-Schleife verlassen
					Case $TrayExit ; Benutzer hat 'Beenden' ausgewählt
						Exit ; Programm beenden
				EndSwitch
			WEnd
		Case $GUI_EVENT_CLOSE ; Benutzer hat auf 'X' geklickt oder 'ESC' gedrückt
			Exit ; Programm beenden
	EndSwitch
Wend

Func Passwort()
	Local $choice = Number(StringLeft(GUICtrlRead($Combo), 1)) ; den 1. Buchstaben von links (Zahl) der Combobox (Passwort-Stärke) auslesen und als $key-Auswahl speichern
	Local $i, $pass = '' ; Passwort = Leerstring
	For $i = 1 To Number(GUICtrlRead($Range)) ; Schleife von 1 bis eingestellter Passwortlänge
		$pass &= StringMid($key[$choice-1], Random(1, StringLen($key[$choice-1]), 1), 1) ; dem Passwort ein zufälliges Zeichen aus dem ausgewählten Stärkestring hinzufügen
	Next ; Schleife fortsetzen
	GUICtrlSetData($Code, $pass) ; das generierte Passwort in das Ausgabefeld schreiben
	;GUICtrlSetState($Code, $GUI_FOCUS) ; den Focus auf das Ausgabefeld, damit man das Passwort einfach per Kontextmenü ('Kopieren') in die Zwischenablage packen kann
EndFunc

; Verwendet die Imagelist um ein Bild zu setzen und Text auf den Buttons darzustellen
Func _GetImageListHandle($sFile, $nIconID = 0, $fLarge = False)
    Local $iSize = 16
    If $fLarge Then $iSize = 32

    Local $hImage = _GUIImageList_Create($iSize, $iSize, 5, 3)
    If StringUpper(StringMid($sFile, StringLen($sFile) - 2)) = "BMP" Then
        _GUIImageList_AddBitmap($hImage, $sFile)
    Else
        _GUIImageList_AddIcon($hImage, $sFile, $nIconID, $fLarge)
    EndIf
    Return $hImage
EndFunc   ;==>_GetImageListHandle