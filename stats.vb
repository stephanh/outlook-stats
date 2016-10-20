Sub HowManyEmails()
    Dim strFilename As String
    Dim objFolder As Outlook.Folder
    
    strFilename = InputBox("Enter a filename (including path) to save the email counts to.", MACRO_NAME)
    
    Dim fso As Object
    Dim fo As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fo = fso.CreateTextFile(strFilename)
    
    Set objFolder = Session.GetDefaultFolder(olFolderInbox)
    ProcessFolder2 objFolder, fo

    fo.Close

    Set fo = Nothing
    Set fso = Nothing
    Set objFolder = Nothing
    
    MsgBox "Done exporting", vbInformation + vbOKOnly, "Export Email counts"
End Sub

Sub ProcessFolder2(olkFld As Outlook.Folder, fo As Object)
    Dim dateStr As String
    Dim olkSub As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim myItem As Object
    Dim dict As Object
    Dim msg As String
    
    msg = ""
    
    Set dict = CreateObject("Scripting.Dictionary")
    Set myItems = olkFld.Items
    myItems.SetColumns ("ReceivedTime")
    ' Determine date of each message:
    For Each myItem In myItems
        If myItem.Class = olMail Then
            dateStr = GetDate(myItem.ReceivedTime)
            If Not dict.Exists(dateStr) Then
                dict(dateStr) = 0
            End If
            dict(dateStr) = CLng(dict(dateStr)) + 1
        End If
    Next myItem

    ' Output counts per day:
    For Each o In dict.Keys
        msg = msg & olkFld.Name & "," & o & "," & dict(o) & vbCrLf
    Next
    
    fo.Write msg
    
    Set dict = Nothing
    For Each olkSub In olkFld.Folders
        ProcessFolder2 olkSub, fo
    Next
    Set olkSub = Nothing
End Sub

Function GetDate(dt As Date) As String
    GetDate = Year(dt) & "-" & Month(dt) & "-" & Day(dt)
End Function

Sub CountTimeSpent()
Dim oOLApp As Outlook.Application
'Dim oSelection As Outlook.Selection
Dim oItem As Object
Dim iDuration As Long
Dim iTotalWork As Long
Dim iMileage As Long
Dim iResult As Integer
Dim bShowiMileage As Boolean
 
bShowiMileage = False
 
iDuration = 0
iTotalWork = 0
iMileage = 0
 
'On Error Resume Next
 
    Set oOLApp = CreateObject("Outlook.Application")
'Set oSelection = oOLApp.ActiveExplorer.Selection
Dim oSelection As Object
Set oSelection = Session.GetDefaultFolder(olFolderCalendar)

 
    For Each oItem In oSelection.Items
    'MsgBox "Item: " & oItem.Class

If oItem.Class = olAppointment Then
iDuration = iDuration + oItem.Duration
End If
Next
 
Dim MsgBoxText As String
MsgBoxText = "Total time spent: " & vbNewLine & iDuration & " minutes"
 
If iDuration > 60 Then
MsgBoxText = MsgBoxText & HoursMsg(iDuration)
End If
 
If iTotalWork > 0 Then
MsgBoxText = MsgBoxText & vbNewLine & vbNewLine & "Total work recorded; " & vbNewLine & iTotalWork & " minutes"
 
If iTotalWork > 60 Then
MsgBoxText = MsgBoxText & HoursMsg(iTotalWork)
End If
End If
 
If bShowiMileage = True Then
MsgBoxText = MsgBoxText & vbNewLine & vbNewLine & "Total iMileage; " & iMileage
End If
 
    iResult = MsgBox(MsgBoxText, vbInformation, "Items Time spent")
 
ExitSub:
Set oItem = Nothing
Set oSelection = Nothing
Set oOLApp = Nothing
End Sub
 
Function HoursMsg(TotalMinutes As Long) As String
Dim iHours As Long
Dim iMinutes As Long
iHours = TotalMinutes \ 60
iMinutes = TotalMinutes Mod 60
HoursMsg = " (" & iHours & " Hours and " & iMinutes & " Minutes)"
End Function
