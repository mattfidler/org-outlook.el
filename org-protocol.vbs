'**************************************
' Name: URLEncode Function
' Description:Encodes a string to create legally formatted
'QueryString for URL. This function is more flexible
'than the IIS Server.Encode function because you can
'pass in the WHOLE URL and only the QueryString data
'will be converted. IIS strangely converts EVERYTHING
'(ie "http://" becomes "http%3A%2F%2F").
' By: Markus Diersbock
'
' Inputs:sRawURL - String to Encode
'
' Returns:Encoded String
'
'This code is copyrighted and has' limited warranties.
'Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=43806&lngWId=1'for details.
'**************************************

' Changed by Matthew Fidler to have http:// become http%3A%2F%2F
' Also changed to have spaces be %20 instead of +


Public Function URLEncode(sRawURL As String) As String
    On Error GoTo Catch
    Dim iLoop As Integer
    Dim sRtn As String
    Dim sTmp As String
    Const sValidChars = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    If Len(sRawURL) > 0 Then
        ' Loop through each char
        For iLoop = 1 To Len(sRawURL)
            sTmp = Mid(sRawURL, iLoop, 1)
            If InStr(1, sValidChars, sTmp, vbBinaryCompare) = 0 Then
                ' If not ValidChar, convert to HEX and prefix with %
                sTmp = Hex(Asc(sTmp))
                If Len(sTmp) = 1 Then
                    sTmp = "%0" & sTmp
                Else
                    sTmp = "%" & sTmp
                End If
            End If
            sRtn = sRtn & sTmp
        Next iLoop
        URLEncode = sRtn
    End If
Finally:
    Exit Function
Catch:
    URLEncode = ""
    Resume Finally
End Function


Sub CreateTaskFromItem()
    Dim T As Variant
    Dim Outlook As New Outlook.Application
    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")

    
    Dim orgfile As Variant
    Dim Pos As Integer
    Dim taskf As Object
    
    Set myNamespace = Outlook.GetNamespace("MAPI")
    Set myPersonalFolder = myNamespace.Folders.Item("Matt")
    Set allPersonalFolders = myPersonalFolder.Folders
    
    T = ""
    For Each Folder In allPersonalFolders
        If Folder.Name = "@ActionTasks" Then
            Set taskf = Folder
            Exit For
        End If
    Next
    
    ' Send selected text to clipboard.
    SendKeys ("%E")
    SendKeys ("C")
    DoEvents
    
    
    Set objWeb = CreateObject("InternetExplorer.Application")
    
        
    If Outlook.Application.ActiveExplorer.Selection.Count > 0 Then
        For i = 1 To Outlook.Application.ActiveExplorer.Selection.Count
                Set objMail = Outlook.ActiveExplorer.Selection.Item(i)
                Set objMail = objMail.Move(taskf)
                objMail.Save 'Maybe this will update EntryID
                T = "org-protocol:/outlook:/o/" + URLEncode(objMail.EntryID) _
                    + "/" + URLEncode(objMail.Subject) _
                    + "/" + URLEncode(objMail.SenderName) _
                    + "/" + URLEncode(objMail.SenderEmailAddress)
                    '+ "/" + URLEncode(objMail.Body)
        objWeb.Navigate T
        objWeb.Visible = True
        Next
    End If
End Sub