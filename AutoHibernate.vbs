Const wshYes = 6
Const wshNo = 7
Const wshCancel = 2
Const wshYesNoCancelDialog = 3
Const wshQuestionMark = 32
Const timeout_sec = 30
Const snooze_min = 10
EndOfLoop = 1

Set objShell = CreateObject("Wscript.Shell")

Do While EndOfLoop = 1
EndOfLoop = 0

intReturn = objShell.Popup(timeout_sec & " Sec to Hibernate!" & vbCrLf & vbCrLf & "Yes => Hibernate Now!"_
& vbCrLf & "No => Snooze for " & snooze_min & " min" & vbCrLf & "Cancel => Do nothing", timeout_sec, _
"Auto Hibernate", wshYesNoCancelDialog + wshQuestionMark)

If intReturn = wshYes Then
  objShell.run "cmd.exe /C shutdown /h"

ElseIf intReturn = wshNo Then
  WScript.Sleep(10*60*1000)
  EndOfLoop = 1

ElseIf intReturn = wshCancel Then
Exit Do

Else 
  objShell.run "cmd.exe /C shutdown /h"
End If

Loop