<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUserSQL
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmUserSQL))
        Me.txtUserSQL = New System.Windows.Forms.TextBox()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RunToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.SaveQueryToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CopyToClipboardToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.UpdateToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.InsertToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.JoinToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.OuterJoinToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.LeftOuterJoinToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.RightOuterJoinToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripButton4 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripButton3 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSplitButton1 = New System.Windows.Forms.ToolStripSplitButton()
        Me.UpdateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.InsertToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.JoinToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OuterJoinToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LeftOuterJoinToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RightOuterJoinToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripButton2 = New System.Windows.Forms.ToolStripButton()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtUserId = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboDSNs = New System.Windows.Forms.ComboBox()
        Me.btnRun = New System.Windows.Forms.Button()
        Me.chkAllowNonSelect = New System.Windows.Forms.CheckBox()
        Me.MenuStrip1.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtUserSQL
        '
        Me.txtUserSQL.AcceptsReturn = True
        Me.txtUserSQL.AcceptsTab = True
        Me.txtUserSQL.Location = New System.Drawing.Point(12, 127)
        Me.txtUserSQL.Multiline = True
        Me.txtUserSQL.Name = "txtUserSQL"
        Me.txtUserSQL.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtUserSQL.Size = New System.Drawing.Size(493, 184)
        Me.txtUserSQL.TabIndex = 3
        Me.txtUserSQL.WordWrap = False
        '
        'MenuStrip1
        '
        Me.MenuStrip1.AllowMerge = False
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(518, 24)
        Me.MenuStrip1.TabIndex = 1
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.RunToolStripMenuItem, Me.ToolStripSeparator3, Me.SaveQueryToolStripMenuItem, Me.CopyToClipboardToolStripMenuItem, Me.ToolStripMenuItem1, Me.ToolStripSeparator2, Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'RunToolStripMenuItem
        '
        Me.RunToolStripMenuItem.Image = CType(resources.GetObject("RunToolStripMenuItem.Image"), System.Drawing.Image)
        Me.RunToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.RunToolStripMenuItem.Name = "RunToolStripMenuItem"
        Me.RunToolStripMenuItem.ShortcutKeyDisplayString = "F5"
        Me.RunToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F5
        Me.RunToolStripMenuItem.Size = New System.Drawing.Size(208, 22)
        Me.RunToolStripMenuItem.Text = "&Run..."
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(205, 6)
        '
        'SaveQueryToolStripMenuItem
        '
        Me.SaveQueryToolStripMenuItem.Image = CType(resources.GetObject("SaveQueryToolStripMenuItem.Image"), System.Drawing.Image)
        Me.SaveQueryToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.SaveQueryToolStripMenuItem.Name = "SaveQueryToolStripMenuItem"
        Me.SaveQueryToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.A), System.Windows.Forms.Keys)
        Me.SaveQueryToolStripMenuItem.Size = New System.Drawing.Size(208, 22)
        Me.SaveQueryToolStripMenuItem.Text = "S&ave Query..."
        '
        'CopyToClipboardToolStripMenuItem
        '
        Me.CopyToClipboardToolStripMenuItem.Image = CType(resources.GetObject("CopyToClipboardToolStripMenuItem.Image"), System.Drawing.Image)
        Me.CopyToClipboardToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.CopyToClipboardToolStripMenuItem.Name = "CopyToClipboardToolStripMenuItem"
        Me.CopyToClipboardToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.B), System.Windows.Forms.Keys)
        Me.CopyToClipboardToolStripMenuItem.Size = New System.Drawing.Size(208, 22)
        Me.CopyToClipboardToolStripMenuItem.Text = "Copy to Clip&board"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UpdateToolStripMenuItem1, Me.InsertToolStripMenuItem1, Me.JoinToolStripMenuItem1, Me.OuterJoinToolStripMenuItem1, Me.LeftOuterJoinToolStripMenuItem1, Me.RightOuterJoinToolStripMenuItem1})
        Me.ToolStripMenuItem1.Image = CType(resources.GetObject("ToolStripMenuItem1.Image"), System.Drawing.Image)
        Me.ToolStripMenuItem1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(208, 22)
        Me.ToolStripMenuItem1.Text = "Insert Template"
        '
        'UpdateToolStripMenuItem1
        '
        Me.UpdateToolStripMenuItem1.Name = "UpdateToolStripMenuItem1"
        Me.UpdateToolStripMenuItem1.Size = New System.Drawing.Size(159, 22)
        Me.UpdateToolStripMenuItem1.Text = "Update"
        '
        'InsertToolStripMenuItem1
        '
        Me.InsertToolStripMenuItem1.Name = "InsertToolStripMenuItem1"
        Me.InsertToolStripMenuItem1.Size = New System.Drawing.Size(159, 22)
        Me.InsertToolStripMenuItem1.Text = "Insert"
        '
        'JoinToolStripMenuItem1
        '
        Me.JoinToolStripMenuItem1.Name = "JoinToolStripMenuItem1"
        Me.JoinToolStripMenuItem1.Size = New System.Drawing.Size(159, 22)
        Me.JoinToolStripMenuItem1.Text = "Join"
        '
        'OuterJoinToolStripMenuItem1
        '
        Me.OuterJoinToolStripMenuItem1.Name = "OuterJoinToolStripMenuItem1"
        Me.OuterJoinToolStripMenuItem1.Size = New System.Drawing.Size(159, 22)
        Me.OuterJoinToolStripMenuItem1.Text = "Outer Join"
        '
        'LeftOuterJoinToolStripMenuItem1
        '
        Me.LeftOuterJoinToolStripMenuItem1.Name = "LeftOuterJoinToolStripMenuItem1"
        Me.LeftOuterJoinToolStripMenuItem1.Size = New System.Drawing.Size(159, 22)
        Me.LeftOuterJoinToolStripMenuItem1.Text = "Left Outer Join"
        '
        'RightOuterJoinToolStripMenuItem1
        '
        Me.RightOuterJoinToolStripMenuItem1.Name = "RightOuterJoinToolStripMenuItem1"
        Me.RightOuterJoinToolStripMenuItem1.Size = New System.Drawing.Size(159, 22)
        Me.RightOuterJoinToolStripMenuItem1.Text = "Right Outer Join"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(205, 6)
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Image = CType(resources.GetObject("ExitToolStripMenuItem.Image"), System.Drawing.Image)
        Me.ExitToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(208, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButton4, Me.ToolStripSeparator4, Me.ToolStripButton3, Me.ToolStripButton1, Me.ToolStripSplitButton1, Me.ToolStripSeparator1, Me.ToolStripButton2})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 24)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(518, 25)
        Me.ToolStrip1.TabIndex = 2
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'ToolStripButton4
        '
        Me.ToolStripButton4.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton4.Image = CType(resources.GetObject("ToolStripButton4.Image"), System.Drawing.Image)
        Me.ToolStripButton4.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton4.Name = "ToolStripButton4"
        Me.ToolStripButton4.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton4.Text = "Run"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(6, 25)
        '
        'ToolStripButton3
        '
        Me.ToolStripButton3.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton3.Image = CType(resources.GetObject("ToolStripButton3.Image"), System.Drawing.Image)
        Me.ToolStripButton3.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton3.Name = "ToolStripButton3"
        Me.ToolStripButton3.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton3.Text = "Save Query"
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton1.Image = CType(resources.GetObject("ToolStripButton1.Image"), System.Drawing.Image)
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton1.Text = "Copy to Clipboard"
        '
        'ToolStripSplitButton1
        '
        Me.ToolStripSplitButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripSplitButton1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UpdateToolStripMenuItem, Me.InsertToolStripMenuItem, Me.JoinToolStripMenuItem, Me.OuterJoinToolStripMenuItem, Me.LeftOuterJoinToolStripMenuItem, Me.RightOuterJoinToolStripMenuItem})
        Me.ToolStripSplitButton1.Image = CType(resources.GetObject("ToolStripSplitButton1.Image"), System.Drawing.Image)
        Me.ToolStripSplitButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripSplitButton1.Name = "ToolStripSplitButton1"
        Me.ToolStripSplitButton1.Size = New System.Drawing.Size(32, 22)
        Me.ToolStripSplitButton1.Text = "Insert Template"
        '
        'UpdateToolStripMenuItem
        '
        Me.UpdateToolStripMenuItem.Name = "UpdateToolStripMenuItem"
        Me.UpdateToolStripMenuItem.Size = New System.Drawing.Size(159, 22)
        Me.UpdateToolStripMenuItem.Text = "Update"
        '
        'InsertToolStripMenuItem
        '
        Me.InsertToolStripMenuItem.Name = "InsertToolStripMenuItem"
        Me.InsertToolStripMenuItem.Size = New System.Drawing.Size(159, 22)
        Me.InsertToolStripMenuItem.Text = "Insert"
        '
        'JoinToolStripMenuItem
        '
        Me.JoinToolStripMenuItem.Name = "JoinToolStripMenuItem"
        Me.JoinToolStripMenuItem.Size = New System.Drawing.Size(159, 22)
        Me.JoinToolStripMenuItem.Text = "Join"
        '
        'OuterJoinToolStripMenuItem
        '
        Me.OuterJoinToolStripMenuItem.Name = "OuterJoinToolStripMenuItem"
        Me.OuterJoinToolStripMenuItem.Size = New System.Drawing.Size(159, 22)
        Me.OuterJoinToolStripMenuItem.Text = "Outer Join"
        '
        'LeftOuterJoinToolStripMenuItem
        '
        Me.LeftOuterJoinToolStripMenuItem.Name = "LeftOuterJoinToolStripMenuItem"
        Me.LeftOuterJoinToolStripMenuItem.Size = New System.Drawing.Size(159, 22)
        Me.LeftOuterJoinToolStripMenuItem.Text = "Left Outer Join"
        '
        'RightOuterJoinToolStripMenuItem
        '
        Me.RightOuterJoinToolStripMenuItem.Name = "RightOuterJoinToolStripMenuItem"
        Me.RightOuterJoinToolStripMenuItem.Size = New System.Drawing.Size(159, 22)
        Me.RightOuterJoinToolStripMenuItem.Text = "Right Outer Join"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'ToolStripButton2
        '
        Me.ToolStripButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton2.Image = CType(resources.GetObject("ToolStripButton2.Image"), System.Drawing.Image)
        Me.ToolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton2.Name = "ToolStripButton2"
        Me.ToolStripButton2.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton2.Text = "Close this window"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(291, 56)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(56, 13)
        Me.Label6.TabIndex = 24
        Me.Label6.Text = "Password:"
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(294, 72)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(113, 20)
        Me.txtPassword.TabIndex = 2
        Me.txtPassword.UseSystemPasswordChar = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(173, 56)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(46, 13)
        Me.Label5.TabIndex = 23
        Me.Label5.Text = "User ID:"
        '
        'txtUserId
        '
        Me.txtUserId.Location = New System.Drawing.Point(176, 72)
        Me.txtUserId.Name = "txtUserId"
        Me.txtUserId.Size = New System.Drawing.Size(112, 20)
        Me.txtUserId.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 55)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(70, 13)
        Me.Label1.TabIndex = 26
        Me.Label1.Text = "Data Source:"
        '
        'cboDSNs
        '
        Me.cboDSNs.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cboDSNs.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cboDSNs.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cboDSNs.Location = New System.Drawing.Point(12, 71)
        Me.cboDSNs.Name = "cboDSNs"
        Me.cboDSNs.Size = New System.Drawing.Size(158, 21)
        Me.cboDSNs.Sorted = True
        Me.cboDSNs.TabIndex = 0
        '
        'btnRun
        '
        Me.btnRun.Location = New System.Drawing.Point(430, 72)
        Me.btnRun.Name = "btnRun"
        Me.btnRun.Size = New System.Drawing.Size(75, 23)
        Me.btnRun.TabIndex = 4
        Me.btnRun.Text = "&Run"
        Me.btnRun.UseVisualStyleBackColor = True
        '
        'chkAllowNonSelect
        '
        Me.chkAllowNonSelect.AutoSize = True
        Me.chkAllowNonSelect.Location = New System.Drawing.Point(12, 99)
        Me.chkAllowNonSelect.Name = "chkAllowNonSelect"
        Me.chkAllowNonSelect.Size = New System.Drawing.Size(184, 17)
        Me.chkAllowNonSelect.TabIndex = 27
        Me.chkAllowNonSelect.Text = "Run as SELECT to see result set."
        Me.chkAllowNonSelect.UseVisualStyleBackColor = True
        '
        'frmUserSQL
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(518, 330)
        Me.Controls.Add(Me.chkAllowNonSelect)
        Me.Controls.Add(Me.btnRun)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cboDSNs)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtUserId)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.txtUserSQL)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmUserSQL"
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.Text = "User SQL"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtUserSQL As System.Windows.Forms.TextBox
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripButton1 As System.Windows.Forms.ToolStripButton
    Friend WithEvents RunToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtUserId As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboDSNs As System.Windows.Forms.ComboBox
    Friend WithEvents btnRun As System.Windows.Forms.Button
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripButton2 As System.Windows.Forms.ToolStripButton
    Friend WithEvents CopyToClipboardToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents SaveQueryToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripButton3 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButton4 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSeparator4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSplitButton1 As System.Windows.Forms.ToolStripSplitButton
    Friend WithEvents UpdateToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents InsertToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents JoinToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OuterJoinToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LeftOuterJoinToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RightOuterJoinToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UpdateToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents InsertToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents JoinToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OuterJoinToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LeftOuterJoinToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RightOuterJoinToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents chkAllowNonSelect As System.Windows.Forms.CheckBox
End Class
