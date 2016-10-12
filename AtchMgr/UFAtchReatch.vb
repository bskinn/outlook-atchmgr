Option Explicit On
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Word = Microsoft.Office.Interop.Word
Imports VBRx = VBScript_RegExp_55
Imports MSVB = Microsoft.VisualBasic

Public Class UFAtchReatch

    Private filesColl As Collection, mi As Outlook.MailItem

    Private Sub UFAtchReatch_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Dim iter As Long, lf As LinkedFile

        ' Initialize all deletes to FALSE
        For iter = 0 To Me.LBxFileList.Items.Count - 1
            ' Typed object for convenience
            lf = filesColl.Item(iter + 1)

            ' Only set 'deleteable possible' if it's hashed
            ' All entries should only ever be the filename at this point
            If lf.isHashed Then
                Me.LBxFileList.Items(iter) = padDelString("No") & Me.LBxFileList.Items(iter)
            Else
                Me.LBxFileList.Items(iter) = padDelString("(N/A)") & Me.LBxFileList.Items(iter)
            End If
        Next iter
    End Sub

    Private Sub UFAtchReatch_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As EventArgs) Handles BtnCancel.Click
        Me.Close()
    End Sub

    Private Sub BtnDoReattach_Click(sender As Object, e As EventArgs) Handles BtnDoReattach.Click
        Dim iter As Integer, lf As LinkedFile, fs As Scripting.FileSystemObject, wd As Word.Document
        Dim ns As Outlook.NameSpace, app As Outlook.Application
        Dim wdRg As Word.Range
        Dim hl As Word.Hyperlink, iter2 As Long, foundHL As Boolean

        ' Link the file system
        fs = CreateObject("Scripting.FileSystemObject")
        app = GetObject(, "Outlook.Application")
        ns = app.GetNamespace("MAPI")

        ' Iterate through the listbox
        For iter = 0 To Me.LBxFileList.Items.Count - 1
            ' If checked, process the entry
            If Me.LBxFileList.SelectedIndices.Contains(iter) Then
                ' Link the lf object
                lf = filesColl.Item(iter + 1)

                ' (Re-)attach the file; this returns the item to non-edit mode for a message
                '  that is not a draft-in-progress.  Check to ensure file still exists before
                '  attaching -- if the same file is also linked via a hashed block, it could have
                '  disappeared
                If fs.FileExists(lf.LinkAddress) Then
                    Call mi.Attachments.Add(lf.LinkAddress)
                Else
                    MsgBox("File """ & lf.dispName & """ no longer exists; cannot attach.", _
                                vbOKOnly + vbExclamation, "Cannot attach file")
                End If

                ' Save here in the event that something crashy happens
                mi.Save()

                ' If was a detached link, cull the paragraph
                If lf.isHashed Then
                    ' If not a draft-in-progress, must restore edit mode
                    If Not mi.Parent.EntryID = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts).EntryID And Not _
                                        mi.Parent.EntryID = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderOutbox).EntryID Then
                        ' Message is not a draft-in-progress and must be reset to edit mode
                        Call mi.GetInspector.CommandBars.ExecuteMso("EditMessage")
                    End If

                    ' Reattach the editor
                    wd = mi.GetInspector.WordEditor

                    ' Appear to always need to reassign Hyperlink after attaching the file
                    ' Must re-search for the hyperlink in the document
                    foundHL = False
                    For iter2 = 1 To wd.Hyperlinks.Count
                        hl = wd.Hyperlinks(iter2)
                        If (InStr(lf.LinkAddress, hl.Address) > 0) And _
                                    hl.TextToDisplay = lf.LinkText Then
                            foundHL = True
                            Exit For
                        End If
                    Next iter2

                    ' If the hyperlink was not retrieved, complain and do not process
                    If Not foundHL Then
                        ' Block not found
                        MsgBox("Detached file annotation block for """ & lf.dispName & _
                                """ not found. Skipping removal of annotation block.", _
                                vbOKOnly + vbInformation, "Annotation block not found")
                    Else
                        ' Block found; strip
                        ' Bind the whole paragraph
                        wdRg = hl.Range.Paragraphs(1).Range

                        ' Simplest way to delete paragraph content
                        wdRg.Text = ""
                        Call wdRg.Delete(Word.WdUnits.wdCharacter, 1)

                        ' Save here
                        mi.Save()
                    End If

                    ' Check whether to delete the stored file
                    If MSVB.Left(Me.LBxFileList.Items(iter), 3) = "Yes" Then
                        ' Do delete
                        fs.DeleteFile(lf.LinkAddress, True)
                    End If
                End If  ' isHashed
            End If  ' .Selected
        Next iter

        ' Save the message one last time, just in case.
        mi.Save()

        ' Close the form
        Me.Close()

    End Sub

    Public Sub popFormStuff(listColl As Collection, mItem As Outlook.MailItem)
        Dim itm As Object, lf As LinkedFile

        filesColl = listColl
        mi = mItem

        ' Iterate over all the items in the collection
        For Each itm In listColl
            ' Set to typed object for convenience
            lf = itm

            ' Add the display name to the listbox
            Me.LBxFileList.Items.Add(lf.dispName)
        Next itm

    End Sub

    Private Function swapYesNo(str As String) As String
        If str = "No" Then
            swapYesNo = "Yes"
        ElseIf str = "Yes" Then
            swapYesNo = "No"
        Else
            ' Do nothing; leave it the same
            swapYesNo = str
        End If
    End Function

    Private Sub LBxFileList_MouseUp(sender As Object, e As Windows.Forms.MouseEventArgs) Handles LBxFileList.MouseUp
        'If Button = 2 Then  ' Right mouse button
        '    'MsgBox Y
        '    With LBxFileList
        '        If Y >= 0.75 And Y <= 11.25 Then
        '            .List(0 + .TopIndex, 1) = swapYesNo(.List(0 + .TopIndex, 1))
        '        ElseIf Y >= 14.25 And Y <= 24.75 Then
        '            .List(1 + .TopIndex, 1) = swapYesNo(.List(1 + .TopIndex, 1))
        '        ElseIf Y >= 26.25 And Y <= 37.55 Then
        '            .List(2 + .TopIndex, 1) = swapYesNo(.List(2 + .TopIndex, 1))
        '        ElseIf Y >= 39 And Y <= 51 Then
        '            .List(3 + .TopIndex, 1) = swapYesNo(.List(3 + .TopIndex, 1))
        '        ElseIf Y >= 52.55 And Y <= 63.05 Then
        '            .List(4 + .TopIndex, 1) = swapYesNo(.List(4 + .TopIndex, 1))
        '        ElseIf Y >= 64.5 And Y <= 75.8 Then
        '            .List(5 + .TopIndex, 1) = swapYesNo(.List(5 + .TopIndex, 1))
        '        ElseIf Y >= 78.05 And Y <= 87.8 Then
        '            .List(6 + .TopIndex, 1) = swapYesNo(.List(6 + .TopIndex, 1))
        '        ElseIf Y >= 90.05 And Y <= 100.55 Then
        '            .List(7 + .TopIndex, 1) = swapYesNo(.List(7 + .TopIndex, 1))
        '        End If
        '    End With
        'End If
    End Sub

    Private Sub LBxFileList_ValueMemberChanged(sender As Object, e As EventArgs) Handles LBxFileList.ValueMemberChanged
        ' THis may not be the right event...
        'Dim iter As Long, somethingSelected As Boolean

        '' Initialize to nothing found
        'somethingSelected = False

        '' See if anything selected
        'For iter = 0 To LBxFileList.ListCount - 1
        '    somethingSelected = somethingSelected Or LBxFileList.Selected(iter)
        'Next iter

        '' If nothing selected, disable the 'do reattach' button
        'BtnDoReattach.Enabled = somethingSelected
    End Sub
End Class