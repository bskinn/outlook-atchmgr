Public Class AtchMgr

    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New RibbonDetach()
    End Function

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        MsgBox("Loaded!")
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        MsgBox("Unloaded!")
    End Sub

End Class
