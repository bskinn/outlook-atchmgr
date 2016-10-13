Imports Microsoft.Office.Tools.Ribbon

Public Class RibBtnDetach

    Private Sub RibBtnDetach_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub DoDetach_Click(sender As Object, e As RibbonControlEventArgs) Handles DoDetach.Click
        MsgBox("Clicked Detach Button!")
        'AtchMgr_Code.DetachAttachment()
    End Sub
End Class
