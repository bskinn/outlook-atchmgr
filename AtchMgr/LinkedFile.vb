Option Explicit On
Imports Word = Microsoft.Office.Interop.Word
Imports Outlook = Microsoft.Office.Interop.Outlook

Public Class LinkedFile

    Private hashTagged As Boolean

    ' Should be enough to uniquely specify a Hyperlink, though in the rare case
    '  of exactly the same file being linked with exactly the same text, confusion
    '  might result.  Could proofread entire Hyperlinks at some point?
    Private URLAddress As String  ' Matches to .Address
    Private URLText As String    ' Matches to .TextToDisplay

    Private hlObj As Word.Hyperlink   ' Will need this to cull 'detached file' info blocks?
    Private fileName As String ' Filename to display in listbox


    Sub New()
        hashTagged = False
    End Sub

    Public Sub setHyperlink(hl As Word.Hyperlink, name As String)
        ' Extracts the URL .Address and .TextToDisplay strings for storage
        '  Fragile to try to retrieve Hyperlink location within the WordEditor,
        '  so not going to mess with it
        If hl.SubAddress <> "" Then
            URLAddress = hl.Address & "#" & hl.SubAddress
        Else
            URLAddress = hl.Address
        End If

        URLText = hl.TextToDisplay
        hlObj = hl
        fileName = name
    End Sub

    Public Function getHyperlink() As Word.Hyperlink
        getHyperlink = hlObj
    End Function

    Public Function matchesHyperlink(hl As Word.Hyperlink)
        ' If target and text are identical, indicate the Hyperlink matches
        '  Potentially problematic if multiple Hyperlinks exist in the document
        '  with identical .Address and .TextToDisplay properties, but for
        '  reattachment purposes, one wouldn't want to reattach the same file twice
        '  anyways.
        '    Hm. Something to check for when parsing the Hyperlinks in the WordEditor:
        '    Crosscheck the .Address of the Hyperlink under examination, and if it
        '    points to an .Address already linked then exclude it from the list of files to reattach
        '    and from the Collection of LinkedFile's?
        If hlObj.Address = hl.Address And hlObj.TextToDisplay = hl.TextToDisplay Then
            matchesHyperlink = True
        Else
            matchesHyperlink = False
        End If
    End Function

    ReadOnly Property LinkAddress() As String
        Get
            LinkAddress = URLAddress
        End Get
    End Property

    ReadOnly Property LinkText() As String
        Get
            LinkText = URLText
        End Get
    End Property

    Property isHashed() As Boolean
        Get
            isHashed = hashTagged
        End Get

        Set(hashed As Boolean)
            hashTagged = hashed
        End Set
    End Property

    ReadOnly Property dispName() As String
        Get
            dispName = fileName
        End Get
    End Property


End Class
