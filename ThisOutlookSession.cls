VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Function getEmailAddress(objRec As Recipient) As String
    getEmailAddress = objRec.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
End Function

Function getMailRecipientString(objRec As Recipient) As String
    getMailRecipientString = ""
    Select Case objRec.Type
        Case 1
            getMailRecipientString = "To:"
        Case 2
            getMailRecipientString = "Cc:"
        Case 3
            getMailRecipientString = "Bcc:"
    End Select
End Function

Function checkMailName(objRec As Recipient) As String
    Select Case objRec.Name = objRec.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
        Case True
            checkMailName = ""
        Case False
            checkMailName = objRec.Name
        End Select
End Function

Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
     On Error GoTo Exception
    
     Dim maxCnt As Integer
     Dim strInternal As String
     Dim strExternal As String
     Dim strBody As String
     Dim maxMailAddressWaring As String
     Dim NumInternal As Integer
     Dim NumExternal As Integer
     maxMailAddressWaring = ""
     maxCnt = 0
     strInternal = ""
     strExternal = ""
     strBody = Item.Body
     NumInternal = 0
     NumExternal = 0
    
    ' 添付ファイルチェック
    If InStr(strSubject & strBody, "添付") > 0 And Item.Attachments.Count = 0 Then
        If MsgBox("添付ファイルを忘れている可能性があります。本当に送信しますか？", vbYesNo + vbQuestion) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    Dim email As String
    Dim objRec As Recipient
    maxMailAddressWaring = ""
    For Each objRec In Item.Recipients
        email = getEmailAddress(objRec)
        ename = checkMailName(objRec)
        If email Like "*empathy.co*" Then
            strInternal = strInternal & getMailRecipientString(objRec) & NumInternal & ":" & " 社内" & ":" & ename & " " & vbCrLf
            NumInternal = NumInternal + 1
        Else
            strExternal = strExternal & getMailRecipientString(objRec) & NumExternal & ":" & " 社外" & ":" & ename & " " & email & vbCrLf
            NumExternal = NumExternal + 1
        End If
        
        maxCnt = maxCnt + 1
        If maxCnt >= 20 Then
            maxMailAddressWaring = "＊＊＊送信先に指定したアドレスが20件以上あります＊＊＊" & vbCrLf
        End If
    Next
    
    Dim strMsg As String
    strMsg = "件名:" & Item.Subject & vbCrLf & vbCrLf & maxMailAddressWaring
    If strExternal <> "" Then
        strMsg = strMsg & vbCrLf & "※※以下のメールアドレスは社外のメールアドレスです※※" & vbCrLf & strExternal & vbCrLf
    End If
    If strInternal <> "" Then
        strMsg = strMsg & vbCrLf & strInternal & vbCrLf
    End If
    strMsg = strMsg & "上記の宛先に、メールを送信してもよろしいですか?"
    If MsgBox(strMsg, vbExclamation + vbYesNo + vbDefaultButton2) <> vbYes Then
        Cancel = True
    End If
    
    On Error GoTo 0
     Exit Sub
    
Exception:
     MsgBox CStr(Err.Number) & ":" & Err.Description, vbOKOnly + vbCritical
     Cancel = True
     Exit Sub
End Sub



