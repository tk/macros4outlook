Attribute VB_Name = "Tools"
Option Explicit
   
Global InterceptorCollection As New Collection




Public Sub MarkMailAsUnread(MyMail As MailItem)
    MyMail.UnRead = True
End Sub

Public Sub ReadCurrentMailItemRTF()
    Dim rtf As String, ret As Integer
    rtf = Space(99999)
    ret = ReadRTF("MAPI", GetCurrentItem.EntryID, Session.GetDefaultFolder(olFolderInbox).StoreID, rtf)
    rtf = Trim(rtf)
    
    Debug.Print "RTF READ:" & ret & vbCrLf & rtf
End Sub

Public Sub TestColors()
    Dim mi As MailItem
    'Set mi = Session.GetDefaultFolder(olFolderInbox).Items(99)
    Set mi = GetCurrentItem()
    'mi.Display
    
    Dim answer As MailItem
    Set answer = mi.reply
    Set mi = Nothing
    
    answer.BodyFormat = olFormatRichText
    
    Dim mid As String
    'mid = QuoteColorizerMacro.ColorizeMailItem(answer)
    answer.Display
    Set answer = Nothing 'answer bodyformat changes here to 1 for some stupid reason...
    
    'Call Tools.DisplayMailItemByID(mid)
End Sub


Public Sub FranksMacro(CurrentItem As MailItem)
    'put mails with me as the ONLY recipient into one folder, all others into another
    
    'declare mapifolders to move to here...
    
    
    If (CurrentItem.Recipients.count > 1) Then
        'put into "uninteresting" folder...
        'CurrentItem.Move(...)
    Else
        'put into "interesting" folder
        'CurrentItem.Move
    End If
    
End Sub
