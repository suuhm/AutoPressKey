Rem VBS Script Press Keys X secs MIL
' Version 1.0 - (c) 2021 suuhm
' SENDKEYS INFO:
' https://ss64.com/vb/sendkeys.html

Dim Message, intVal, MyValue, Title, Default, objResult
' Dim intVal, * as Integer

Message = "SendKey interval?"
Title = "Input the seconds of sendkey interval here..."
Default = "For Example: 60 secs for every minute keydown or NP (Notepad focusscript)"

' Display message, title, and default value.
MyValue = InputBox(Message, Title, Default)
' MyValue = InputBox(Message, Title, , , , "DEMO.HLP", 10)
' MyValue = InputBox(Message, Title, Default, 100, 100)

VT = VarType(MyValue)
   If IsNumeric(MyValue) Then 
   VT = 2
End If
' https://docs.microsoft.com/en-us/previous-versions/visualstudio/aa263402(v=vs.60)?redirectedfrom=MSDN
' Int -> 2 // Str -> 8

MsgBox "Your TYPE-INPUT VALUE is: " + CStr(VT), vbError, "Now Starting Script..."

If VT = 8 And MyValue <> "NP" Then
   MsgBox "WRONG INPUT INT OR NP!!", vbWarning, "Stopping Script..."
   Quit 0

Elseif VT <> 2 And MyValue = "NP" Then
   MsgBox "RUNNING NOTEPAD>>", vbError, "Starting and write some stuff in notepad..."
   Set objShell = CreateObject("WScript.Shell")    
   Call objShell.Run("%windir%\system32\notepad.exe")
   objShell.AppActivate("Notepad") 'Fokus NP
   sleep (5000) ' 5 secs

   For i=1 To 7 Step 1
     ' ### TESTING OPEN NP AND SEND HELLO
     objResult = objShell.sendkeys("LET ME PRESS A KEY FOR YOU >> Nr. " + CStr(i) + ": {ENTER}")
     sleep (2000) ' 2 Sekunden
   Next

Elseif CInt(MyValue) >= 1 Then

   intVal = CInt(MyValue)*1000 ' Millisecs to secs to mins *60
   intTimer = 0
   i2 = intVal
   MsgBox "Input: " + CStr(intVal) + " milliseconds! (" + CStr(intVal/1000/60) + ") minutes" , vbInformation, "Starting Script and minimize..."

   Set objShell = CreateObject("WScript.Shell")  
   
   ' Minimize window
   sleep 333
   objShell.SendKeys "% n"

   sleep 2000
   Do While True
     objResult = objShell.sendkeys("{NUMLOCK}{NUMLOCK}")
     ' MsgBox "Script laeuft seit: " + CStr(intTimer) + " Millisecs", vbWarning, "Stopping Script..."
     ' intTimer = intTimer + intVal
     sleep i2
     ' window.setTimeOut "o1", intVal
   Loop

Else ' MyValue < 1
  MsgBox "WRONG INPUT PLS ENTER SECONDS OR NP!!", vbWarning, "Stopping Script..."
  Quit 0
End If
