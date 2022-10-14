<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class HistoChecker
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.view = New System.Windows.Forms.ListBox()
        Me.vbProjList = New System.Windows.Forms.ComboBox()
        Me.vbCompList = New System.Windows.Forms.ComboBox()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.btn_refresh = New System.Windows.Forms.Button()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'view
        '
        Me.view.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.view.FormattingEnabled = True
        Me.view.Location = New System.Drawing.Point(0, 39)
        Me.view.Name = "view"
        Me.view.Size = New System.Drawing.Size(406, 381)
        Me.view.TabIndex = 0
        '
        'vbProjList
        '
        Me.vbProjList.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.vbProjList.FormattingEnabled = True
        Me.vbProjList.Location = New System.Drawing.Point(3, 3)
        Me.vbProjList.Name = "vbProjList"
        Me.vbProjList.Size = New System.Drawing.Size(143, 21)
        Me.vbProjList.TabIndex = 1
        '
        'vbCompList
        '
        Me.vbCompList.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.vbCompList.FormattingEnabled = True
        Me.vbCompList.Location = New System.Drawing.Point(3, 3)
        Me.vbCompList.Name = "vbCompList"
        Me.vbCompList.Size = New System.Drawing.Size(158, 21)
        Me.vbCompList.TabIndex = 2
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(85, 3)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.vbProjList)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.vbCompList)
        Me.SplitContainer1.Size = New System.Drawing.Size(317, 31)
        Me.SplitContainer1.SplitterDistance = 149
        Me.SplitContainer1.TabIndex = 3
        '
        'btn_refresh
        '
        Me.btn_refresh.Location = New System.Drawing.Point(4, 6)
        Me.btn_refresh.Name = "btn_refresh"
        Me.btn_refresh.Size = New System.Drawing.Size(75, 27)
        Me.btn_refresh.TabIndex = 4
        Me.btn_refresh.Text = "Refresh"
        Me.btn_refresh.UseVisualStyleBackColor = True
        '
        'HistoChecker
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.btn_refresh)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.view)
        Me.Name = "HistoChecker"
        Me.Size = New System.Drawing.Size(406, 431)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents view As Windows.Forms.ListBox
    Friend WithEvents vbProjList As Windows.Forms.ComboBox
    Friend WithEvents vbCompList As Windows.Forms.ComboBox
    Friend WithEvents SplitContainer1 As Windows.Forms.SplitContainer
    Friend WithEvents btn_refresh As Windows.Forms.Button
End Class
