Attribute VB_Name = "CountMailDays"
Const Multiplier = 1000

Sub CountEmailDays()

    Dim objOutlook As Object, objnSpace As Object, objFolder As MAPIFolder
    Dim EmailCount As Integer
    Set objOutlook = CreateObject("Outlook.Application")
    Set objnSpace = objOutlook.GetNamespace("MAPI")

    On Error Resume Next
    
    'Get active folder
    Set objFolder = Application.ActiveExplorer.CurrentFolder
    
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "No such folder."
        Exit Sub
    End If

    EmailCount = objFolder.Items.Count

    Dim emailDays, delta As Integer
    Dim myItems As Outlook.Items
    Dim msg As String
    
    Set myItems = objFolder.Items
    myItems.SetColumns ("ReceivedTime")
    
    ' Determine date of each message:
    emailDays = 0
    For Each myItem In myItems
        delta = Multiplier * ((Date + Time) - (myItem.ReceivedTime))
        emailDays = emailDays + delta
    Next myItem

    ' Output results:
    MsgBox objFolder.FullFolderPath & ":" & vbCrLf & vbCrLf & emailDays / Multiplier & " письмодней инбокса, всего " & EmailCount & " писем."

    Set objFolder = Nothing
    Set objnSpace = Nothing
    Set objOutlook = Nothing
End Sub
