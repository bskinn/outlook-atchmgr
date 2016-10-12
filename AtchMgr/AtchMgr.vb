Public Class AtchMgr

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        MsgBox("Loaded!")
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        MsgBox("Unloaded!")
    End Sub

End Class
