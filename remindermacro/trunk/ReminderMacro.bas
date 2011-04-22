Attribute VB_Name = "ReminderMacro"
'$Id$
'
'Reminder Macro TRUNK
'
'Reminder Macro is part of the macros4outlook project
'see https://sourceforge.net/apps/mediawiki/macros4outlook/index.php?title=Reminder_Macro or
'    http://sourceforge.net/projects/macros4outlook/ for more information
'
'For more information on Outlook see http://www.microsoft.com/outlook
'Outlook is (C) by Microsoft

Option Explicit

'------------------------------------------------------------------------------------------
' Procedure : Jeremy's Application_ItemSend Event 1.0
' Author    : Jeremy Gollehon
' Purpose   : Warn on blank Subject line and/or no attachment (using keyword check).
'             Program works with all message types (only tested in Outlook 2003).
'
' DateTime  : 7/05/2004,  - Original concept code
'           : 8/17/2004,  - Some optimization and fixing of logic errors.
'           : 8/18/2004,  - Added functionality for all message types.
'                         - Now searches Subject and Body for keywords.
'                         - In Reply/forward's only non-quoted section of body is searched.
'           : 8/19/2004,  - (Armen Stein) Changed array declaration to a Split, so that new
'                           search words can be easily added in a constant.
'           : 8/20/2004,  - Check to make sure code only runs on MailItem type.
'           : 8/23/2004   - Added ExactMatch function: check's to be sure the exact
'                           keyword/keyphrase was found. Eg. "here it is" vs "where it is"
'                         - Added EmbeddedAttachCount function (code mostly taken from
'                           Outlookcode.com). It's used to determine the number of embedded
'                           attachments and exlude them from the attachment count.  This code
'                           uses the Redemption dll (http://www.dimastr.com/redemption)
'                           which must be installed/registered in Windows, and referenced,
'                           Tools> References...> SafeOutlook Library, in Outlook VBA.
'           : 7/13/2006   - DM: Removed dependecies to "Outlook redemption" library
'                           Released as v1.0 by macros4outlook project
'------------------------------------------------------------------------------------------

Sub CheckMailText(ByVal Item As Object, Cancel As Boolean)
  Dim bCancelSend As Boolean
  Dim sTextToSearch As String
  Dim sKeyWords As String
  Dim vKeyWords() As String
  Dim iStartOfQuote As Long
  Dim iAttachmentCount As Long
  Dim i As Long

  If TypeName(Item) <> "MailItem" Then Exit Sub

  'Add keywords/phrases here.  Use lowercase words in the following array.
  sKeyWords = "attach;attached;attachment;enclosed;here's;here it is;anhang;angehängt;anlage;anbei"

  'CHECK FOR BLANK SUBJECT LINE
  If Trim(Item.Subject) = "" Then
    bCancelSend = MsgBox("This message does not have a subject." & vbNewLine & _
                         "Do you wish to continue sending anyway?", _
                         vbYesNo + vbExclamation, "No Subject") = vbNo
  End If

  'CHECK BODY AND SUBJECT FOR ATTACMENT KEYWORDS.
  'Set TextToSearch variable to message Body based
  'on message type and find start of quoted text.
  Select Case Item.BodyFormat
    Case olFormatHTML
      iStartOfQuote = InStr(Item.HTMLBody, "<DIV class=OutlookMessageHeader") - 1
      sTextToSearch = Item.HTMLBody
    Case olFormatRichText
      iStartOfQuote = InStr(Item.Body, "_____________________________________________") - 1
      sTextToSearch = Item.Body
    Case olFormatPlain
      iStartOfQuote = InStr(Item.Body, "-----Original Message-----") - 1
      sTextToSearch = Item.Body
  End Select
  'Adjust TextToSearch if there is quoted text
  If iStartOfQuote > 0 Then sTextToSearch = Left(sTextToSearch, iStartOfQuote)
  'Add Subject to the search text if not a Reply
  If Left(Item.Subject, 3) <> "RE:" Then
    sTextToSearch = Item.Subject & " " & sTextToSearch
  End If
  'Change to all lowercase for string comparison
  sTextToSearch = LCase(sTextToSearch)
  'Replace undesired characters with spaces to help with ExactMatch function
  sTextToSearch = Replace(sTextToSearch, ",", " ")
  sTextToSearch = Replace(sTextToSearch, ".", " ")
  sTextToSearch = Replace(sTextToSearch, "?", " ")
  sTextToSearch = Replace(sTextToSearch, "!", " ")
  sTextToSearch = Replace(sTextToSearch, Chr(34), " ")  'quotes
  sTextToSearch = Replace(sTextToSearch, "<", " ")  'beginning of html tag
  sTextToSearch = Replace(sTextToSearch, ">", " ")  'end of html tag
  sTextToSearch = Replace(sTextToSearch, "&", " ")  'beginning of html Character Entities
  sTextToSearch = Replace(sTextToSearch, ";", " ")  'end of html Character Entities
  
  'Start the search
  If Not bCancelSend Then
    iAttachmentCount = Item.Attachments.count 'DM: - EmbeddedAttachCount(Item)
    If iAttachmentCount = 0 Then
      vKeyWords = Split(sKeyWords, ";")
      For i = LBound(vKeyWords) To UBound(vKeyWords)
        If InStr(sTextToSearch, vKeyWords(i)) > 0 Then
          If StrExactMatch(sTextToSearch, vKeyWords(i)) Then
            bCancelSend = MsgBox("It appears you were going to send an attachment but nothing is attached." & vbNewLine & _
                                 "Do you wish to continue sending anyway?" & vbNewLine & vbNewLine & _
                                 "Word/Phrase found:  " & vKeyWords(i), _
                                 vbYesNo + vbExclamation, "Attachment Not Found") = vbNo
            Exit For
          End If
        End If
      Next i
    End If
  End If

  'Cancel sending message if answered yes to either message box.
  Cancel = bCancelSend
End Sub


Private Function StrExactMatch(sLookIn As String, sLookFor As String) As Boolean
  '- Add padding to sLookin in case sLookfor is at
  '  the very beginning or very end of the sLookIn.
  '- Add padding to sLookFor to ensure an exact match
  StrExactMatch = (InStr(" " & sLookIn & " ", " " & sLookFor & " ") > 0) _
                    Or (InStr(sLookIn, vbCrLf & sLookFor & " ") > 0) _
                    Or (InStr(sLookIn, " " & sLookFor & vbCrLf) > 0) _
                    Or (InStr(sLookIn, vbCrLf & sLookFor & vbCrLf) > 0)
End Function

