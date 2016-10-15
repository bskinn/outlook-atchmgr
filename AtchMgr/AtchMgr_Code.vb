Option Explicit On
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Word = Microsoft.Office.Interop.Word
Imports VBRx = VBScript_RegExp_55

Module AtchMgr_Code

    Public Const padWidth As Integer = 10

    Sub EditForAtchAnnot()
        Dim insp As Outlook.Inspector, app As Outlook.Application
        Dim wdEd As Word.Document
        Dim itm As Object

        app = GetObject(, "Outlook.Application")

        If app.ActiveInspector() Is Nothing Then Exit Sub

        insp = app.ActiveInspector()
        wdEd = insp.WordEditor
        itm = insp.CurrentItem

        With itm
            ' Close and reopen item to provide consistent state
            Call .Close(Outlook.OlInspectorClose.olSave)
            .Display()
            Call .GetInspector.CommandBars.ExecuteMso("EditMessage")  ' Activate edit mode (presumes valid)
            Call .GetInspector.CommandBars.ExecuteMso("MessageFormatHtml")  ' Set to HTML mode
        End With

        wdEd.StoryRanges(Word.WdStoryType.wdMainTextStory).InsertBefore("[#" & vbCrLf & vbCrLf)
        With wdEd.StoryRanges(Word.WdStoryType.wdMainTextStory).Paragraphs(1).Range
            .Font.TextColor.RGB = RGB(120, 113, 68)
            .Font.Italic = True
            .Characters(2).Select()
            .Characters(2).Delete(Word.WdUnits.wdCharacter, 1)
        End With

    End Sub

    Public Sub ReattachAttachments()
        ' Will be better done by collecting all (re)attach-able files and presenting in a checkbox-enabled
        '  UserForm, permitting user to select which ones to (re)attach and whether or not to retain the
        '  out-of-message copies of any that were originally macro-detached
        '
        ' Macro in present form will reattach all linked, locally-accessible files, including any that were
        '  linked in the original text of the message.  This may not be the desired behavior.

        Dim wd As Word.Document, wdRg As Word.Range
        Dim fs As New Scripting.FileSystemObject
        Dim hl As Word.Hyperlink
        Dim insp As Outlook.Inspector, app As Outlook.Application, ns As Outlook.NameSpace
        Dim mi As Outlook.MailItem
        Dim lf As LinkedFile, lfColl As New Collection, lfIter As LinkedFile, itm As Object
        Dim alreadyLinked As Boolean

        Dim rx As New VBRx.RegExp
        Dim frm As New UFAtchReatch

        Dim fNameFull As String, fName As String, fPath As String
        Dim fullAddress As String
        Dim iter As Long, atchsExist As Boolean

        app = GetObject(, "Outlook.Application")

        ' If no active inspector, just silently drop
        If app.ActiveInspector() Is Nothing Then Exit Sub

        ns = app.GetNamespace("MAPI")
        insp = app.ActiveInspector()
        wd = insp.WordEditor
        'fs = CreateObject("Scripting.FileSystemObject")

        ' Set up RegEx
        With rx
            .Global = False
            .IgnoreCase = True
            .Multiline = False
            .Pattern = "[a-z]:\\"
        End With

        ' Check each hyperlink for whether it's local or not; if any found, indicate as such
        atchsExist = False
        For Each hl In wd.Hyperlinks
            If Left(hl.Address, 2) = "\\" Or rx.Test(hl.Address) Then
                atchsExist = True
                Exit For
            End If
            '(rx.Test(hl.Address) And Not Left(hl.Address, 4) = "http") Then atchsExist = True
        Next hl

        ' If found attachments to reattach, perform reattachment
        If atchsExist Then
            mi = insp.CurrentItem  ' Comment these lines
            Call mi.Close(Outlook.OlInspectorClose.olSave)
            mi.Display()
            insp = mi.GetInspector
            wd = insp.WordEditor
            If Not mi.Parent.EntryID = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts).EntryID And Not _
                    mi.Parent.EntryID = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderOutbox).EntryID Then
                Call insp.CommandBars.ExecuteMso("EditMessage")  ' To here
                ' Breaks if drafting a message using 'Resend Message' w/o having saved a draft first
            End If

            iter = 1  ' Must initialize
            Do While iter <= wd.Hyperlinks.Count
                hl = wd.Hyperlinks(iter)
                If Left(hl.Address, 2) = "\\" Or rx.Test(hl.Address) Then
                    '(rx.Test(hl.Address) And Not Left(hl.Address, 4) = "http") Then
                    ' Local file; attempt reattach
                    ' Parse into folder and filename
                    ' Have to tag on the anchor name, if present -- "#" is an assumption, but is the
                    '   primary character encountered thus far in my uses
                    fullAddress = hl.Address
                    If Len(hl.SubAddress) > 0 Then fullAddress = fullAddress & "#" & hl.SubAddress

                    fNameFull = fs.GetAbsolutePathName(fullAddress)
                    fPath = fs.GetParentFolderName(fNameFull)
                    fName = fs.GetFileName(fNameFull)

                    If fs.FolderExists(fPath) Then      ' Folder exists (add notification if not exist)
                        If fs.FileExists(fNameFull) Then    ' File exists (add notification if not exist)
                            ' Check if hl already linked
                            alreadyLinked = False
                            For Each itm In lfColl
                                lfIter = itm
                                If lfIter.matchesHyperlink(hl) Then
                                    alreadyLinked = True
                                    Exit For
                                End If
                            Next itm

                            ' Check for already linked
                            If Not alreadyLinked Then
                                ' Set up the linked file object if not already linked
                                lf = New LinkedFile
                                'lf.DeleteHashed = False
                                'lf.ID = lfColl.Count + 1
                                lf.setHyperlink(hl, fName)
                                'lf.ProcessFile = False

                                ' Flag for whether it's hashed or not
                                wdRg = hl.Range.Paragraphs(1).Range
                                If Left(wdRg.Text, 3) = "###" And Right(wdRg.Text, 4) = "###" & Chr(13) Then
                                    lf.isHashed = True
                                Else
                                    lf.isHashed = False
                                End If

                                ' Add the linked file to the collection
                                lfColl.Add(lf)
                            End If  ' alreadyLinked
                        End If  ' FileExists
                    End If  ' FolderExists
                End If  ' local link

                ' Increase the iterator
                iter = iter + 1
            Loop    ' Until all hl's checked

            ' Populate the reattachment form
            frm.popFormStuff(lfColl, mi)

            ' Show the form
            frm.Activate()

        End If  ' Local links found

        ' Dereference variables
        fs = Nothing
        wd = Nothing
        insp = Nothing
        rx = Nothing
        hl = Nothing
        mi = Nothing

    End Sub

    Public Sub DetachAttachment()
        Dim atch As Outlook.Attachment, atchSel As Outlook.AttachmentSelection
        Dim insp As Outlook.Inspector, app As Outlook.Application
        Dim itm As Object
        Dim sh As Shell32.Shell
        Dim fs As New Scripting.FileSystemObject
        Dim fld As Shell32.Folder2
        Dim fName As String, extn As String, baseName As String, fullSavePath As String
        Dim atchName As String
        Dim okfName As Boolean
        Dim iter As Long
        'Dim bodyHTML As String, timeRef As Long

        MsgBox("Made it into DetachAttachment")

        ' Bind script object, inspector,  and attachment selection
        sh = CreateObject("Shell.Application")
        app = GetObject(, "Outlook.Application")
        insp = app.ActiveInspector()
        itm = insp.CurrentItem
        atchSel = insp.AttachmentSelection

        ' Create filesystem object
        'fs = CreateObject("Scripting.FileSystemObject")

        ' Initialize folder to Nothing; should be redundant
        fld = Nothing

        ' Close-with-save the Item and re-open for consistent, stable state DOESN'T WORK
        'Call itm.Close(olSave)
        'itm.Display
        'Set insp = itm.GetInspector

        ' Loop selected attachments, asking to detach or not
        For Each atch In atchSel
            ' Query for fld if none yet selected
            If fld Is Nothing Then fld = sh.BrowseForFolder(0, "Detach file(s) to:", 1, 17)

            ' Check whether folder was selected
            If Not fld Is Nothing Then
                ' Request filename from user, checking/confirming overwrite
                okfName = False ' Initialize no-good filename
                Do
                    atchName = atch.FileName  ' Store attachment filename
                    fName = atch.FileName ' Initialize working filename
                    ' Split working filename into base and extension
                    iter = Len(fName)
                    Do Until Mid(fName, iter, 1) = "." Or iter = 1 : iter = iter - 1 : Loop
                    If iter > 1 Then ' period found
                        baseName = Left(fName, iter - 1)
                        extn = Right(fName, Len(fName) - iter)
                    Else ' period not found
                        baseName = fName
                        extn = ""
                    End If

                    ' Query filename; stop exec if zero-length return (user cancel); reconstruct with extension
                    fName = InputBox("Save to filename:", "Enter File Name", baseName)
                    If Len(fName) < 1 Then Exit For ' okfName already False; fragile if another surrounding For..Next added
                    If Len(extn) > 0 Then fName = fName & "." & extn
                    ' Check whether filename ok based on file existence
                    okfName = Not fs.FileExists( _
                                IIf(Right(fld.Self.Path, 1) = "\", _
                                        fld.Self.Path, _
                                        fld.Self.Path & "\" _
                                    ) & cleanFilename(fName))
                    ' If exists (name is not okay), ask if overwrite ok
                    If Not okfName Then
                        ' Need to deal with not-ok filename
                        Select Case MsgBox("File exists" & Chr(10) & Chr(10) & "Overwrite?", _
                                    vbYesNoCancel + vbExclamation, "Confirm Overwrite")
                            Case vbYes
                                ' Go ahead and overwrite
                                okfName = True
                            Case vbNo
                                ' Just pass through the not-ok-filename flag
                            Case Else
                                ' Presumably just vbCancel is possible; exit sub
                                Exit For ' This is weak; addition of another wrapping For..Next will break code
                        End Select
                    End If
                Loop Until okfName

                ' If anything survives filename cleaning...
                If Not cleanFilename(fName) = "" Then
                    ' Save atch to path\filename and delete
                    fullSavePath = IIf(Right(fld.Self.Path, 1) = "\", _
                                            fld.Self.Path, _
                                            fld.Self.Path & "\" _
                                        ) & cleanFilename(fName)
                    If Len(fName) = Len(cleanFilename(fName)) Then
                        ' To robustify, add wait loop that checks to be sure saved-out file exists
                        '  and has nonzero size (filesize match is not useful; attached file does not
                        '  have identical reported bytesize as on-disk file)
                        Call atch.SaveAsFile(fullSavePath)  ' FRAGILE if Save op unexpectedly fails!
                        Call atch.Delete()
                        'timeRef = Timer: Do While Timer <= timeRef + 0.5: DoEvents: Loop
                        Call tagTextIntoEmail(itm, atchName, fullSavePath)
                        Call itm.Save()
                    Else
                        ' Some invalid characters stripped; notify and save
                        Select Case MsgBox("Invalid characters have been stripped from the indicated " & _
                                    "filename. File will be saved as:" & Chr(10) & Chr(10) & _
                                    cleanFilename(fName), vbOKCancel + vbExclamation, _
                                    "Invalid Characters Removed")
                            Case vbOK
                                Call atch.SaveAsFile(fullSavePath)  ' FRAGILE if Save op unexpectedly fails!
                                Call atch.Delete()
                                'timeRef = Timer: Do While Timer <= timeRef + 0.5: DoEvents: Loop
                                Call tagTextIntoEmail(itm, atchName, fullSavePath)
                                Call itm.Save()
                            Case Else
                                ' Do nothing
                        End Select
                    End If
                End If
            Else
                ' No folder is set; presume that user cancelled & wants to exit routine
                Exit For
            End If
        Next atch

        ' Dereference objects
        sh = Nothing
        fld = Nothing
        itm = Nothing
        insp = Nothing
        atchSel = Nothing
        atch = Nothing

    End Sub

    Public Sub tagTextIntoEmail(itm As Object, atchName As String, saveName As String)
        ' Might still be able to use Word.Document
        Dim wd As Word.Document, newRg As Word.Range, editRg As Word.Range
        'Dim mi As Outlook.MailItem
        Dim idx As Long
        Const verbStr As String = "' detached to "

        With itm
            ' Close and reopen item to provide consistent state
            Call .Close(Outlook.OlInspectorClose.olSave)
            .Display()
            Call .GetInspector.CommandBars.ExecuteMso("EditMessage")  ' Activate edit mode (presumes valid)
            Call .GetInspector.CommandBars.ExecuteMso("MessageFormatHtml")  ' Set to HTML mode
        End With

        ' Attach Word Document for editing
        wd = itm.GetInspector.WordEditor

        ' Insert attachment detachment notification text
        Call wd.Content.InsertBefore("###Attachment '" & atchName & verbStr & saveName & Chr(13) & Chr(13))

        ' Bind the newly added paragraph's Range
        newRg = wd.Content.Paragraphs(1).Range
        With newRg
            ' Change color to red
            .Font.ColorIndex = Word.WdColorIndex.wdRed
            ' Identify where link location starts
            idx = InStr(.Text, verbStr) + Len(verbStr)
            ' Set editing Range to first character
            editRg = .Characters(idx)
            ' Extend editing Range to end of paragraph
            Call editRg.MoveEnd(Word.WdUnits.wdParagraph, 1)
            ' Deselect hard return
            Call editRg.MoveEnd(Word.WdUnits.wdCharacter, -1)
            ' Append hashes
            Call editRg.InsertAfter("###")
            ' Deselect hashes
            Call editRg.MoveEnd(Word.WdUnits.wdCharacter, -3)
            ' Apply hyperlink
            Call wd.Hyperlinks.Add(editRg, editRg.Text)

        End With

    End Sub

    Public Function cleanFilename(ByVal fn As String) As String
        Dim badchrs As New List(Of String)(), val As Long, st As String, ch As String
        badchrs.Add("\")
        badchrs.Add("/")
        badchrs.Add(":")
        badchrs.Add("""")
        badchrs.Add("*")
        badchrs.Add("?")
        badchrs.Add("<")
        badchrs.Add(">")
        badchrs.Add("|")


        ' Set string to shorthand variable
        st = fn

        ' Search for and remove bad characters
        For Each val In badchrs
            ch = badchrs(val)
            Do While InStr(st, ch) > 0
                st = Left(st, InStr(st, ch) - 1) & Right(st, Len(st) - InStr(st, ch))
            Loop
        Next val

        ' Set the cleaned filename to the output variable
        cleanFilename = st

    End Function

    Public Function padDelString(ByRef padStr As String) As String
        padDelString = padStr & StrDup(padWidth - Len(padStr), " ")
    End Function

End Module