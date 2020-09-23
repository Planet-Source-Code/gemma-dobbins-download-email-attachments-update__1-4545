<div align="center">

## Download Email Attachments:UPDATE


</div>

### Description

Updated email program. This code allows you to download multiple

attachments and copy them into a directory. The program then

replys to the author with a message or/and attachment automatically.
 
### More Info
 
MapiSession Control, MapiMessage Control,

2 command buttons, 1 text box.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Gemma Dobbins](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/gemma-dobbins.md)
**Level**          |Unknown
**User Rating**    |4.0 (28 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/gemma-dobbins-download-email-attachments-update__1-4545/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
  MAPISession1.DownLoadMail = False
  MAPISession1.SignOn
  MAPIMessages1.SessionID = MAPISession1.SessionID
  MAPIMessages1.MsgIndex = -1
  MAPIMessages1.Compose
  MAPIMessages1.Send True
  MAPISession1.SignOff
End Sub
Private Sub Command2_Click()
  MAPISession1.NewSession = True
  MAPISession1.Action = 1 'session_signon
  MAPIMessages1.SessionID = MAPISession1.SessionID
  MAPIMessages1.FetchUnreadOnly = True
  MAPIMessages1.Action = 1 'message_fetch
     Dim i As Integer
    Text1.Text = MAPIMessages1.MsgNoteText
     For i = 0 To MAPIMessages1.AttachmentCount - 1
       MAPIMessages1.AttachmentIndex = i
       Dim intLenFileName As Integer
       Dim intStrPos As Integer
       intLenFileName = Len(MAPIMessages1.AttachmentPathName)
       For intStrPos = intLenFileName To 1 Step -1
         If InStr(1, _
             Right$(MAPIMessages1.AttachmentPathName, _
                 intLenFileName - (intStrPos - 1)), _
             "\", 1) Then
           strNewFileName = _
            Right$(MAPIMessages1.AttachmentPathName, _
                intLenFileName - intStrPos)
           Exit For
         End If
       Next
       FileCopy MAPIMessages1.AttachmentPathName, _
           "c:\" & strNewFileName
     Next
     Mail
     MAPIMessages1.Delete
  MAPISession1.SignOff
End Sub
Private Function Mail()
 Dim o As New Outlook.Application
 Dim m As Object
 Set m = o.CreateItem(olMailItem)
 m.To = MAPIMessages1.MsgOrigAddress
 m.Subject = "Fantastic!!!"
 m.Attachments.Add "C:\Fantastic.txt"
 m.Show ' this can be taken out if you want an automated program
 m.Send
End Function
```

