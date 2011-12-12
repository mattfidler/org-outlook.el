Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA"_
                     ( _
                      ByVal hWnd As Long, _
                      ByVal lpOperation As String, _
                      ByVal lpFile As String, _
                      ByVal lpParameters As String, _
                      ByVal lpDirectory As String, _
                      ByVal nShowCmd As Long) As Long
                     
                     ''Slightly Modified http://www.freevbcode.com/ShowCode.Asp?ID=5137
Function URLEncode(EncodeStr As String) As String
        Dim i As Integer
        Dim erg As String

        erg = EncodeStr

        ' *** First replace '%' chr
        erg = Replace(erg, \"%\", Chr(1))

        ' *** then '+' chr
        erg = Replace(erg, \"+\", Chr(2))

        For i = 0 To 255
                Select Case i
                        ' *** Allowed 'regular' characters
                        Case 37, 43, 48 To 57, 65 To 90, 97 To 122

                        Case 1  ' *** Replace original %
                                erg = Replace(erg, Chr(i), \"%25\")

                        Case 2  ' *** Replace original +
                                erg = Replace(erg, Chr(i), \"%2B\")

                        Case 32
                                erg = Replace(erg, Chr(i), \"%20\") 'org-protocol likes %20 instead of +

                        Case 3 To 15
                                erg = Replace(erg, Chr(i), \"%0\" & Hex(i))
                                
                        Case Else
                                erg = Replace(erg, Chr(i), \"%\" & Hex(i))

                End Select
        Next

        URLEncode = erg

End Function


Sub CreateTaskFromItem()
        Dim T As Variant
        Dim Outlook As New Outlook.Application
        Dim ie As Object
        Set ie = CreateObject(\"InternetExplorer.Application\")


        Dim orgfile As Variant
        Dim Pos As Integer
        Dim taskf As Object

        Set myNamespace = Outlook.GetNamespace(\"MAPI\")

        ' Change this to be your personal folder item.  If it remains
        ' on the server it keeps the Outlook ID originally given.  If
        ' you move it to another folder, it will assign it to another
        ' ID, but keep that ID as long as you don't move it back to the
        ' server. (*sigh*  I wish it kept the same ID.)

        ' Technically this is unnecessary, but with my limited exchange
        ' account size,  I move my emails to \"Personal Folders\\@ActionTasks\" and
        ' then (possibly) refile from there. 
        
        Set myPersonalFolder = myNamespace.Folders.item(\"Personal Folders\")
        Set allPersonalFolders = myPersonalFolder.Folders

        T = \"\"
        For Each Folder In allPersonalFolders
                If Folder.Name = \"@ActionTasks\" Then
                        Set taskf = Folder
                        Exit For
                End If
        Next

        ' End moving message.

        If Outlook.Application.ActiveExplorer.Selection.Count > 0 Then
                For i = 1 To Outlook.Application.ActiveExplorer.Selection.Count
                        Set objMail = Outlook.ActiveExplorer.Selection.item(i)
                        Set objMail = objMail.Move(taskf)
                        objMail.Save 'Maybe this will update EntryID
                        ' Note that o is the Outlook capture template.
                        T = \"org-protocol:/outlook:/o/\" + URLEncode(objMail.EntryID) _
                          + \"/\" + URLEncode(objMail.Subject) _
                          + \"/\" + URLEncode(objMail.SenderName) _
                          + \"/\" + URLEncode(objMail.SenderEmailAddress)
                        ShellExecute 0, \"open\", T, vbNullString, vbNullString, vbNormalFocus
                Next
        End If
End Sub