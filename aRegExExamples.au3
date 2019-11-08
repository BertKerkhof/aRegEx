#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile=aRegExExamples.exe
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Description=String functions
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_LegalCopyright=Bert Kerkhof 2019-11-07 Apache 2.0 license
#AutoIt3Wrapper_Res_SaveSource=n
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 3 -w 4 -w 5
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/tc 2 /reel
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/pe /sf /sv /rm
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

; The latest version of this source is published at GitHub in the
; BertKerkhof repository "aRegEx".

#include-once
#include <EditConstants.au3>; Delivered With AutoIT
#include <GuiEdit.au3>; Delivered With AutoIT
#include <GUIConstantsEx.au3>; Delivered With AutoIT
#include <WindowsConstants.au3>; Delivered With AutoIT
#include <aDrosteArray.au3>; GitHub published by Bert Kerkhof
#include <FileJuggler.au3>; GitHub published by Bert Kerkhof
#include <aRegEx.au3>; Github published by Bert Kerkhof

; Author: kerkhof.bert@gmail.com
; Tested with: AutoIt version 3.3.14.5 and win10


; aRegEx examples ===================================================

; #FUNCTION#
; Name ..........: TestTitleCase
; Description ...: Test the sTitleCase function
; Syntax ........: TestTitleCase()
; Parameters ....: None
; Returns .......: None
; Author ........: Bert Kerkhof
; Remark ........: sTitleCase is an aRegEx function. Its alternative
;                  StringTitleCase is in the StringSupport module.
;                  An other function _StringTitleCase is delivered with AutoIT
;                  in the Include folder, String.au3 file.

Func TestTitleCase()
  Local $S = "Although it started as a simple automation tool, AutoIt now has functions " & _
      "and features that allow it to be used as a general purpose scripting language."
  MsgBox(0, "TitleCase", sTitleCase($S))
EndFunc   ;==>TestTitleCase
; TestTitleCase()


; #FUNCTION#
; Name ..........: CsvReader
; Description ...: Select and read a table of comma separated values
; Syntax ........: CsvReader()
; Parameters ....: None
; Returns .......: None
; Comment .......: Census bureau's offer tables on the Internet. This is how you
;                  can get these into your system. A table about renewable energy
;                  is supplied together with this source on GitHub.
; Author ........: Bert Kerkhof

Func CsvReader()
  Local $sFile = FileOpenDialog("CsvReader", @MyDocumentsDir, "Table (*.csv;*.txt)", 3)
  If @error Then Exit
  Local $sInput = FileRead($sFile)
  Local $amArray = amRegExCapture($sInput, "((?:[\w\- ]+,)+[\w\- ]+)\r*\n")
  Local $sOut = ($amArray[0]) ? mRegExdotHeader($amArray[1], $sInput) : ""
  For $I = 1 To $amArray[0]
    $sOut &= sRecite(StringSplit(($amArray[$I])[1], ","), 16) & @CRLF
  Next
  $sOut &= mRegExdotFooter($amArray[$amArray[0]], $sInput)
  DisplayBox("CsvReader", $sOut)
EndFunc   ;==>CsvReader
; CsvReader()

; #FUNCTION#
; Name ..........: NeedleInHaystack
; Description ...: Search and display address info found in text- and AutoIT sources
; Syntax ........: NeedleInHaystack()
; Parameters ....: None
; Returns .......: None
; Comment .......: fRits said me: man you're so busy with the programs, don't forget to add
;                  sample patterns. The users want examples. Adapt the collection to your needs.
;                  This program can scan folders such as Program files\AutoIT3\Examples.
;                  A file 'NeedleInHaystackMaterials.txt' is delivered together with this
;                  source code on GitHub.
; Author ........: Bert Kerkhof

Func NeedleInHaystack()
  aD("Email address", "\b[\w+\.]{5,}@[\w]{3,}.\w{2,5}\b")
  aD("Dutch phone number", "\b\(*\d{3,4}(?:\h\-\h|\-|\h)\d{6,7}\b")
  aD("American phone number", "\s\([2-9][0-8][0-9]\)[2-9][0-9]{2}[\-\.][0-9]{4}\s")
  aD("International phone number", "\b\+[0-9]{1,3}\h[0-9\-]{4,14}\b")
  aD("ISBN", "\bISBN:\h\d\d\d\-\d-\d\d\d\-\d\d\d\d\d\-\d\b")
  aD("Canadian postal code", "\b[A-VXY][0-9][A-Z]\h[0-9][A-Z][0-9]\b")
  aD("UK postal code", "\b[A-Z]{1,2}[0-9][A-Z]\h[0-9][A-Z]{2}\b")
  aD("URL", "\b(?:https:\/\/|ftp:\/\/|www\/)\w+\.\w{2,5}\b")
  Local $aFiles = aSelectFiles("txt|au3", "SubFolders")
  If @error Then Exit
  Local $sRoot = @WorkingDir, $sOut = $sRoot & @CRLF, $nOffset = StringLen($sRoot) + 1
  Local $aaPattern = aD()
  Local $sFile, $sString, $aFound, $aPattern, $Lfirst, $Handle, $N = 0
  For $I = 1 To $aFiles[0]
    $sFile = $aFiles[$I]
    $Handle = FileOpen($sFile)
    $sString = FileRead($Handle)
    FileClose($Handle)
    $Lfirst = True
    For $P = 1 To $aaPattern[0]
      $aPattern = $aaPattern[$P]
      $aFound = aRegExCapture($sString, $aPattern[2])
      If $aFound[0] Then
        If $Lfirst Then
          $sOut &= Spaces(2) & StringMid($sFile, $nOffset) & @CRLF
          $Lfirst = False
        EndIf
        $sOut &= Spaces(4) & $aPattern[1] & ":" & @CRLF
        $N += $aFound[0]
        For $J = 1 To $aFound[0]
          $sOut &= Spaces(6) & $aFound[$J] & @CRLF
        Next
      EndIf
    Next
  Next
  DisplayBox("Needle in Haystack", $sOut & $N & ' found.' & @CRLF)
EndFunc   ;==>NeedleInHaystack
; NeedleInHaystack()

; #FUNCTION#
; Name ..........: DisplayBox
; Description ...: Helper function for NeedleInHaystack and CsvReader
; Syntax ........: DisplayBox($sTitle[, $sContent = ""])
; Parameters ....: $sTitle .... Title to display in the header
;                  $sContent .. Content to display
; Returns .......: None
; Author ........: Bert Kerkhof

Func DisplayBox($sTitle, $sContent = "")
  GUICreate($sTitle, 600, 406) ;
  GUICtrlCreateEdit($sContent, 0, 0, 600, 406, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $WS_HSCROLL, $WS_VSCROLL))
  GUICtrlSetFont(-1, 10, 400, 0, "Courier New")
  GUICtrlSetBkColor(-1, 0xFFFF80)
  GUICtrlSetState(-1, 256)
  GUISetState(@SW_SHOW)
  While True
    Local $nMsg = GUIGetMsg()
    Switch $nMsg
      Case $GUI_EVENT_CLOSE
        Exit
    EndSwitch
  WEnd
EndFunc   ;==>DisplayBox

; End =================================================================
