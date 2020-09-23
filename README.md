<div align="center">

## Inter\-Process Messaging


</div>

### Description

This provides a SIMPLE solution to sending data between any 2 vb applications. No ActiveX, DDE, COM, DCOM or OLE requirements. The transfer process is easy to follow and you should have little difficulty passing to more controls if required.
 
### More Info
 
In your own program substitue "Receiving AppName" with the name of your program (on both sending and receiving ends). For every control that is to have data passed to it just create an entry as per the textbox examples;

SaveSetting sAppName, "InterProcess Handles", "Text1", Str$(Text1.hWnd)

The sample forms provided show text being passed to 3 textboxes, but you could easily modify to pass to combo boxes etc. The solution involves teporary storage in the windows registry of the necessary windows handles.

If you want the recipient program to act upon the received text you could simply add a change procedure such as;

Private Sub Text1_Change()

MsgBox "Change detected"

End Sub

I hope you find this solution usefull and if you like it, please vote for me.

When text data is sent to the receiving program it is placed straight in to the control and are ready for use.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Phobos](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/phobos.md)
**Level**          |Intermediate
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/phobos-inter-process-messaging__1-11707/archive/master.zip)

### API Declarations

```
The following declarations are present at the start of the message sending program. No other API's are required.
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SETTEXT = &HC
```


### Source Code

```
'********* MESSAGE SENDING PROGRAM **********
'
' This program will send text messages to another vb program.
' The messages will be placed directly into the text boxes.
' Add 1 wide command button (Command1) to a blank form, double
' click on the form, then copy and paste the following source code.
' (This will be a separate project called message sender)
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SETTEXT = &HC
' This program will send test messages to another vb program.
' The recipient must be running when the command button is pressed.
Private Sub Command1_Click()
  Dim sAppName As String, sSection As String
  ' Here we must supply the name of the program which is to receive messages.
  sAppName = "Receiving AppName"
  If Not InterProcMsg(sAppName, "Text1", "Message to Text1") Then
    ' Notify if the message could not be sent.
    MsgBox "Could not send message sent to Text1"
  End If
  If Not InterProcMsg(sAppName, "Text2", "Message to Text2") Then
    ' Notify if the message could not be sent.
    MsgBox "Could not send message sent to Text2"
  End If
  If Not InterProcMsg(sAppName, "Text3", "Message to Text3") Then
    ' Notify if the message could not be sent.
    MsgBox "Could not send message sent to Text3"
  End If
End Sub
Function InterProcMsg(sAppName As String, sKey As String, sValue As String) As Boolean
On Error GoTo Err_InterProcMsg
  ' This routine will place a text message (sValue) into a control on a form
  ' running on another program.
  '
  ' In order for this to work the recipient program must be running,
  ' and must have stored the required windows handles into the windows registry.
  Dim sSection As String, lRequiredHandle As Long, SentOK As Boolean
  sSection = "InterProcess Handles"
  ' First we obtain the required handle from the registry.
  lRequiredHandle = GetSetting(sAppName, sSection, sKey)
  ' If a valid handle was found the send the message passed in the string 'sValue'.
  If lRequiredHandle = 0 Then
    SentOK = False   ' Message not sent (handle not found)
  Else
    Call SendMessage(lRequiredHandle, WM_SETTEXT, ByVal 0&, ByVal sValue)
    SentOK = True    ' Message sent
  End If
Exit_InterProcMsg:
  ' Exit the function with InterProcMsg set to either
  '    TRUE if message sent to the other program without problems, or
  '    FALSE if the message could not be sent.
  InterProcMsg = SentOK
  Exit Function
Err_InterProcMsg:
  ' Error handler to catch and process any unexpected errors.
  MsgBox "Error" & Str$(Err) & " in routine InterProcMsg on sending form: " & Error$(Err)
  SentOK = False   ' Message not sent (due to unexpected error)
  GoTo Exit_InterProcMsg
End Function
Private Sub Form_Load()
  ' Add a prompt to the command button.
  Command1.Caption = "Send Messages to the other program"
End Sub
'
'********* MESSAGE RECEIVING PROGRAM **********
'
' This program will receive text messages from another vb program.
' The messages will be placed directly into the text boxes.
' Add 3 text boxes (text1, text2 and text3) to a blank form, double
' click on the form, then copy and paste the following source code.
' (This will be a separate project called message receiver)
Option Explicit
Private Sub Form_Load()
  ' To allow the sending program to write to our textboxes, we make a
  ' temporary saving of windows handles of the textboxes to the registry.
  Dim sAppName As String
  ' Here we must supply the name of this program
  ' (the name must match that given in the sending program).
  sAppName = "Receiving AppName"
  ' Now we store the windows handles for the forms textboxes.
  SaveSetting sAppName, "InterProcess Handles", "Text1", Str$(Text1.hWnd)
  SaveSetting sAppName, "InterProcess Handles", "Text2", Str$(Text2.hWnd)
  SaveSetting sAppName, "InterProcess Handles", "Text3", Str$(Text3.hWnd)
End Sub
Private Sub Form_Unload(Cancel As Integer)
  ' The program has now finished, so we can now remove
  ' our InterProcess handle values from the registry.
  DeleteSetting "Receiving AppName", "InterProcess Handles"
End Sub
```

