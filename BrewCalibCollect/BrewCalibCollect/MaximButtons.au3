; script to hit the MaximDL buttons to reset the image calibration FileSaveDialog

; assumes MaximDL is already running

#include <MsgBoxConstants.au3>

; Retrieve runstring parameters.
; First param is the number of frames to be combined; for example,
; 5 darks and 15 flats would be 20 subs to be combined.
; I am assuming the subs need 2 seconds each to combine, so we know how long to wait for the combination
; step to last
$combineCount = $CmdLine[0]
if ($CmdLine[0] = 0) Then
   $combineCount = 5
Else
   $combineCount = $CmdLine[1]
EndIf
if ($combineCount < 5) then
   $combineCount = 6
EndIf
$combineMsec = $combineCount * 2000   ; in msec

if (ProcessExists ("MaxIm_DL.exe") = 0) Then
   ; need to start Maxim
   Run ("C:\Program Files (x86)\Diffraction Limited\MaxIm DL V6\MaxIm_DL.exe", "")
   sleep(3000)
EndIf

$mainMaxim = WinWait("MaxIm DL Pro 6","",2)
if ($MainMaxim = 0) then
   MsgBox($MB_OK,"MaximButtons","Maxim Window not found")
   Exit
EndIf
WinActivate("MaxIm DL Pro 6")

; send Alt-P Alt-L to select Process/Set Calibration menu item
send("{ALTDOWN}pl{ALTUP}")
if (WinWaitActive("[CLASS:#32770; TITLE:MaxIm DL Pro 6]","",5) <> 0) Then
;if (WinWaitActive("MaxIm DL Pro 6","",2) <> 0) Then
   ; got the intermediate dialog indicating Maxim is missing some FileSaveDialog
   sleep(1000)
   send("{ENTER}")
   sleep(500)
endif

Local $hWnd = WinWaitActive("Set Calibration","",5)
if ($hWnd  <> 0) Then
   ; OK, we have the dialog
   ; click the AutoGenerate button
   ControlClick($hWnd, "", "[CLASS:Button; INSTANCE:5]")
   WinWaitActive("Set Calibration","",5)
   sleep(1)

   ; click the Replace Masters button
   ControlClick($hWnd, "", "[CLASS:Button; INSTANCE:6]")
   sleep($combineMsec)
   WinWaitActive("Set Calibration","",1)

   ; Click OK
   ControlClick($hWnd, "", "[CLASS:Button; INSTANCE:1]")
Else
   msgbox($MB_OK, "Set Calibration Not Found","Did not get the Set Calibration dialog")
EndIf

;msgbox($MB_OK, "All done","Script complete")
