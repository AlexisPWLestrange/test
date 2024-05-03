Attribute VB_Name = "Module1"
Dim olExp As Explorer
Dim olSel As Selection
Dim senderName As AddressEntry
Dim olMail As MailItem
Dim olAppt As AppointmentItem
Dim oPA As PropertyAccessor

Dim olName As NameSpace
Dim olFldr As Folder

Sub initVariables()
    
    ' Unused
    Set olName = Application.GetNamespace("MAPI")
    Set olFldr = olName.GetDefaultFolder(olFolderDrafts)
    
    ' In Use
    Set olExp = Application.ActiveExplorer
    Set olSel = olExp.Selection
    
End Sub

'Sub test()
'
'    Debug.Print Time & ": " & getTo
'
'End Sub
'
'Sub testcode()
'
'    Dim strDate, strMonth, strYear As String
'
'    strDate = DateValue("01/01/2000")
'    strMonth = Month(strDate)
'    strYear = Year(strDate)
'    Debug.Print Time & ": " & DateSerial(strYear, strMonth + 1, 0)
'    Debug.Print Time & ": " & strMonth - 1
'
'End Sub

Sub CopyEmails()
    
    
    
End Sub

Sub Attach_Item()
    
    
    
End Sub

Sub Subject_DateCode_To_Date()
    
    Dim textPos As Long
    Dim newText, dateType As String
    
    Call initVariables
    
    Set olMail = olSel.Item(1)
    
    'Debug.Print Time & ": " &
    
    With olMail
        textPos = InStr(1, .Subject, "^") + 1
        .Subject = Replace(.Subject, "^" & Mid(.Subject, textPos), getDate(Mid(.Subject, textPos)))
    End With
    
End Sub

Sub Body_DateCode_To_Date()
    
    Dim textPos As Long
    Dim newText, dateType As String
    
    Call initVariables
    
    Set olMail = olSel.Item(1)
    
    With olMail
        textPos = InStr(1, .Body, "^") + 1
        .HTMLBody = Replace(.HTMLBody, "^" & Mid(.Body, textPos, 3), getDate(Mid(.Body, textPos, 3)))
    End With
    
End Sub

Function getDate(ByVal dateType As String)
    
    Dim dateText As String
    
    Select Case dateType
        Case "LM", "LMX" ' Last Month
            dateText = Format(CDate(DateAdd("M", -1, Now)), "MMMM YYYY") ' i.e. March 2024
        Case "M", "MXX" ' This Month
            dateText = Format(CDate(Date), "MMMM YYYY") ' i.e. April 2024
        Case "D", "DXX" ' Date
            dateText = Format(CDate(Date), "DD.MM.YYYY") ' i.e. 01.04.2024
        Case "DL", "DLX" ' Date Long
            dateText = Format(CDate(Date), "DD MMMM YYYY") ' i.e. 01 April 2024
        Case "YD", "YDX" ' Yesterday's Date
            dateText = Format(CDate(Date - 1), "DD.MM.YYYY") ' i.e. 31.03.2024
        Case "YDL" ' Yesterday's Date Long
            dateText = Format(CDate(Date - 1), "DD MMMM YYYY") ' i.e. 31 March 2024
        Case Else ' In case the dateType is not found, return the dateType code
            dateText = "^" & dateType
    End Select
        
    getDate = dateText
    
End Function

Function getTo()
    
    Dim strSenderID As String
    Const PR_SENT_REPRESENTING_ENTRYID As String = _
    "http://schemas.microsoft.com/mapi/proptag/0x00410102"
    Dim MsgTxt As String
    Dim x As Long
    
    Call initVariables
    
    For x = 1 To olSel.Count
        If olSel.Item(x).Class = OlObjectClass.olMail Then
            ' For mail item, use the SenderName property.
            Set olMail = olSel.Item(x)
            MsgTxt = MsgTxt & olMail.To & ";"
        ElseIf olSel.Item(x).Class = OlObjectClass.olAppointment Then
            ' For appointment item, use the Organizer property.
            Set olAppt = olSel.Item(x)
            MsgTxt = MsgTxt & olAppt.Organizer & ";"
        Else
            ' For other items, use the property accessor to get the sender ID,
            ' then get the address entry to display the sender name.
            Set oPA = olSel.Item(x).PropertyAccessor
            strSenderID = oPA.GetProperty(PR_SENT_REPRESENTING_ENTRYID)
            Set senderName = Application.Session.GetAddressEntryFromID(strSenderID)
            MsgTxt = MsgTxt & senderName.Name & ";"
        End If
    Next x
    
    getTo = MsgTxt
    
End Function

Option Explicit
Sub GetValueUsingRegEx()
    
    ' Set reference to VB Script library
    ' Microsoft VBScript Regular Expressions 5.5
    
    Dim olMail As MailItem
    Dim Reg1 As RegExp
    Dim M1 As MatchCollection
    Dim M As Match
    
    Set olMail = Application.ActiveExplorer().Selection(1)
   ' Debug.Print olMail.Body
   
    Set Reg1 = New RegExp
    
    With Reg1
        .Pattern = "(\d{11}|\d{3}-\d{8})"
        .Global = True
    End With
    
    If Reg1.test(olMail.Body) Then
        Set M1 = Reg1.Execute(olMail.Body)
        For Each M In M1
            MsgBox M.SubMatches(0)
        Next
    End If
    
End Sub

Sub AutomateReplyWithSearchString()

    Dim myInspector As Outlook.Inspector
    Dim myObject As Object
    Dim myItem As Outlook.MailItem
    Dim myDoc As Word.Document
    Dim mySelection As Word.Selection
    Dim strItem As String
    Dim strGreeting As String

    Set myInspector = Application.ActiveInspector
    Set myObject = myInspector.CurrentItem

    'The active inspector is displaying a mail item.
    If myObject.MessageClass = "IPM.Note" And myInspector.IsWordMail = True Then
        Set myItem = myInspector.CurrentItem

        'Grab the body of the message using a Word Document object.
        Set myDoc = myInspector.WordEditor
        myDoc.Range.Find.ClearFormatting
        Set mySelection = myDoc.Application.Selection
        With mySelection.Find

            .Text = "xxx-xxxxxxxx"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True

        End With

        If mySelection.Find.Execute = True Then
            strItem = mySelection.Text

            'Mail item is in compose mode in the inspector
            If myItem.Sent = False Then
                strGreeting = "With reference to " + strItem
                myDoc.Range.InsertBefore (strGreeting)
            End If
        Else
            MsgBox "There is no item number in this message."

        End If
    End If
End Sub
