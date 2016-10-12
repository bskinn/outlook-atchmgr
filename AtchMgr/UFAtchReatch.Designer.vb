<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UFAtchReatch
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.LBxFileList = New System.Windows.Forms.ListBox()
        Me.LblFileListBox = New System.Windows.Forms.Label()
        Me.LblDelFileCol = New System.Windows.Forms.Label()
        Me.BtnDoReattach = New System.Windows.Forms.Button()
        Me.BtnCancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'LBxFileList
        '
        Me.LBxFileList.ColumnWidth = 100
        Me.LBxFileList.FormattingEnabled = True
        Me.LBxFileList.Location = New System.Drawing.Point(10, 76)
        Me.LBxFileList.Name = "LBxFileList"
        Me.LBxFileList.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
        Me.LBxFileList.Size = New System.Drawing.Size(392, 95)
        Me.LBxFileList.TabIndex = 0
        Me.LBxFileList.TabStop = False
        '
        'LblFileListBox
        '
        Me.LblFileListBox.AutoSize = True
        Me.LblFileListBox.Location = New System.Drawing.Point(269, 20)
        Me.LblFileListBox.Name = "LblFileListBox"
        Me.LblFileListBox.Size = New System.Drawing.Size(115, 13)
        Me.LblFileListBox.TabIndex = 0
        Me.LblFileListBox.Text = "Select files to reattach:"
        Me.LblFileListBox.Visible = False
        '
        'LblDelFileCol
        '
        Me.LblDelFileCol.Location = New System.Drawing.Point(12, 20)
        Me.LblDelFileCol.Name = "LblDelFileCol"
        Me.LblDelFileCol.Size = New System.Drawing.Size(63, 43)
        Me.LblDelFileCol.TabIndex = 0
        Me.LblDelFileCol.Text = "Delete File? (right-click to toggle)"
        '
        'BtnDoReattach
        '
        Me.BtnDoReattach.Enabled = False
        Me.BtnDoReattach.Location = New System.Drawing.Point(109, 179)
        Me.BtnDoReattach.Name = "BtnDoReattach"
        Me.BtnDoReattach.Size = New System.Drawing.Size(103, 32)
        Me.BtnDoReattach.TabIndex = 1
        Me.BtnDoReattach.Text = "Reattach Files"
        Me.BtnDoReattach.UseVisualStyleBackColor = True
        '
        'BtnCancel
        '
        Me.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnCancel.Location = New System.Drawing.Point(228, 179)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(103, 32)
        Me.BtnCancel.TabIndex = 2
        Me.BtnCancel.Text = "Cancel"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'UFAtchReatch
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnCancel
        Me.ClientSize = New System.Drawing.Size(414, 222)
        Me.Controls.Add(Me.BtnCancel)
        Me.Controls.Add(Me.BtnDoReattach)
        Me.Controls.Add(Me.LblFileListBox)
        Me.Controls.Add(Me.LBxFileList)
        Me.Controls.Add(Me.LblDelFileCol)
        Me.MaximumSize = New System.Drawing.Size(430, 260)
        Me.Name = "UFAtchReatch"
        Me.Text = "Select Files To Reattach"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LBxFileList As System.Windows.Forms.ListBox
    Friend WithEvents LblFileListBox As System.Windows.Forms.Label
    Friend WithEvents LblDelFileCol As System.Windows.Forms.Label
    Friend WithEvents BtnDoReattach As System.Windows.Forms.Button
    Friend WithEvents BtnCancel As System.Windows.Forms.Button
End Class
