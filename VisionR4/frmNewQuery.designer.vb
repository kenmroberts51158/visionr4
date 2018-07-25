<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNewQuery
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmNewQuery))
        Me.cboDSNs = New System.Windows.Forms.ComboBox()
        Me.btnConnect = New System.Windows.Forms.Button()
        Me.ListBoxFields = New System.Windows.Forms.ListBox()
        Me.cboTables = New System.Windows.Forms.ComboBox()
        Me.ListBoxSelectedFields = New System.Windows.Forms.ListBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.btnRunQuery = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CurrentSQLToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RecordCountToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripButton2 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton3 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton7 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripButton5 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton6 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnUseKey = New System.Windows.Forms.Button()
        Me.ListBoxWhereClause = New System.Windows.Forms.ListBox()
        Me.cboAndOr = New System.Windows.Forms.ComboBox()
        Me.btnAddSearch = New System.Windows.Forms.Button()
        Me.cboSearchValue = New System.Windows.Forms.ComboBox()
        Me.cboBool = New System.Windows.Forms.ComboBox()
        Me.cboSelectedField = New System.Windows.Forms.ComboBox()
        Me.txtUserId = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cboOrderBy = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.rdoAscending = New System.Windows.Forms.RadioButton()
        Me.rdoDescending = New System.Windows.Forms.RadioButton()
        Me.ckbDistinct = New System.Windows.Forms.CheckBox()
        Me.MenuStrip1.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cboDSNs
        '
        Me.cboDSNs.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cboDSNs.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cboDSNs.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cboDSNs.Location = New System.Drawing.Point(12, 72)
        Me.cboDSNs.Name = "cboDSNs"
        Me.cboDSNs.Size = New System.Drawing.Size(158, 21)
        Me.cboDSNs.Sorted = True
        Me.cboDSNs.TabIndex = 0
        '
        'btnConnect
        '
        Me.btnConnect.Location = New System.Drawing.Point(413, 69)
        Me.btnConnect.Name = "btnConnect"
        Me.btnConnect.Size = New System.Drawing.Size(100, 23)
        Me.btnConnect.TabIndex = 3
        Me.btnConnect.Text = "Connect"
        Me.btnConnect.UseVisualStyleBackColor = True
        '
        'ListBoxFields
        '
        Me.ListBoxFields.FormattingEnabled = True
        Me.ListBoxFields.Location = New System.Drawing.Point(12, 155)
        Me.ListBoxFields.Name = "ListBoxFields"
        Me.ListBoxFields.Size = New System.Drawing.Size(207, 108)
        Me.ListBoxFields.TabIndex = 5
        '
        'cboTables
        '
        Me.cboTables.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.cboTables.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cboTables.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cboTables.Location = New System.Drawing.Point(12, 113)
        Me.cboTables.Name = "cboTables"
        Me.cboTables.Size = New System.Drawing.Size(471, 21)
        Me.cboTables.TabIndex = 4
        '
        'ListBoxSelectedFields
        '
        Me.ListBoxSelectedFields.FormattingEnabled = True
        Me.ListBoxSelectedFields.Location = New System.Drawing.Point(294, 155)
        Me.ListBoxSelectedFields.Name = "ListBoxSelectedFields"
        Me.ListBoxSelectedFields.Size = New System.Drawing.Size(189, 108)
        Me.ListBoxSelectedFields.TabIndex = 10
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(226, 155)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(62, 23)
        Me.Button2.TabIndex = 6
        Me.Button2.Text = ">>>"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(226, 184)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(62, 23)
        Me.Button3.TabIndex = 7
        Me.Button3.Text = ">"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(225, 213)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(62, 23)
        Me.Button4.TabIndex = 8
        Me.Button4.Text = "<"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(225, 242)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(62, 23)
        Me.Button5.TabIndex = 9
        Me.Button5.Text = "<<<"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.Image = CType(resources.GetObject("Button6.Image"), System.Drawing.Image)
        Me.Button6.Location = New System.Drawing.Point(490, 177)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(23, 30)
        Me.Button6.TabIndex = 11
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Button7
        '
        Me.Button7.Image = CType(resources.GetObject("Button7.Image"), System.Drawing.Image)
        Me.Button7.Location = New System.Drawing.Point(489, 213)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(25, 30)
        Me.Button7.TabIndex = 12
        Me.Button7.UseVisualStyleBackColor = True
        '
        'btnRunQuery
        '
        Me.btnRunQuery.Enabled = False
        Me.btnRunQuery.Location = New System.Drawing.Point(409, 431)
        Me.btnRunQuery.Name = "btnRunQuery"
        Me.btnRunQuery.Size = New System.Drawing.Size(105, 37)
        Me.btnRunQuery.TabIndex = 18
        Me.btnRunQuery.Text = "&Run Query"
        Me.btnRunQuery.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(70, 13)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Data Source:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 97)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(42, 13)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Tables:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 139)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(83, 13)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "Fields Available:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(291, 139)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(194, 13)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "Fields Selected: *an empty list means all"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.AllowMerge = False
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.ViewToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(528, 24)
        Me.MenuStrip1.TabIndex = 13
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SaveToolStripMenuItem, Me.ToolStripSeparator5, Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(35, 20)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'SaveToolStripMenuItem
        '
        Me.SaveToolStripMenuItem.Image = CType(resources.GetObject("SaveToolStripMenuItem.Image"), System.Drawing.Image)
        Me.SaveToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.SaveToolStripMenuItem.Name = "SaveToolStripMenuItem"
        Me.SaveToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.A), System.Windows.Forms.Keys)
        Me.SaveToolStripMenuItem.Size = New System.Drawing.Size(178, 22)
        Me.SaveToolStripMenuItem.Text = "S&ave Query..."
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(175, 6)
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Image = CType(resources.GetObject("ExitToolStripMenuItem.Image"), System.Drawing.Image)
        Me.ExitToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.ShortcutKeyDisplayString = ""
        Me.ExitToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(178, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'ViewToolStripMenuItem
        '
        Me.ViewToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CurrentSQLToolStripMenuItem, Me.RecordCountToolStripMenuItem, Me.ToolStripSeparator2, Me.ToolStripMenuItem1, Me.ToolStripMenuItem2})
        Me.ViewToolStripMenuItem.Name = "ViewToolStripMenuItem"
        Me.ViewToolStripMenuItem.Size = New System.Drawing.Size(41, 20)
        Me.ViewToolStripMenuItem.Text = "&View"
        '
        'CurrentSQLToolStripMenuItem
        '
        Me.CurrentSQLToolStripMenuItem.Image = CType(resources.GetObject("CurrentSQLToolStripMenuItem.Image"), System.Drawing.Image)
        Me.CurrentSQLToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.CurrentSQLToolStripMenuItem.Name = "CurrentSQLToolStripMenuItem"
        Me.CurrentSQLToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.Q), System.Windows.Forms.Keys)
        Me.CurrentSQLToolStripMenuItem.Size = New System.Drawing.Size(186, 22)
        Me.CurrentSQLToolStripMenuItem.Text = "Current S&QL..."
        '
        'RecordCountToolStripMenuItem
        '
        Me.RecordCountToolStripMenuItem.Image = CType(resources.GetObject("RecordCountToolStripMenuItem.Image"), System.Drawing.Image)
        Me.RecordCountToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.RecordCountToolStripMenuItem.Name = "RecordCountToolStripMenuItem"
        Me.RecordCountToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.T), System.Windows.Forms.Keys)
        Me.RecordCountToolStripMenuItem.Size = New System.Drawing.Size(186, 22)
        Me.RecordCountToolStripMenuItem.Text = "Record Coun&t..."
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(183, 6)
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Image = CType(resources.GetObject("ToolStripMenuItem1.Image"), System.Drawing.Image)
        Me.ToolStripMenuItem1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(186, 22)
        Me.ToolStripMenuItem1.Text = "View &Schema..."
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.Image = CType(resources.GetObject("ToolStripMenuItem2.Image"), System.Drawing.Image)
        Me.ToolStripMenuItem2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.K), System.Windows.Forms.Keys)
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(186, 22)
        Me.ToolStripMenuItem2.Text = "View &Keys..."
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButton2, Me.ToolStripButton3, Me.ToolStripButton7, Me.ToolStripSeparator1, Me.ToolStripButton5, Me.ToolStripButton6, Me.ToolStripSeparator4, Me.ToolStripButton1})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 24)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(528, 25)
        Me.ToolStrip1.TabIndex = 14
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'ToolStripButton2
        '
        Me.ToolStripButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton2.Image = CType(resources.GetObject("ToolStripButton2.Image"), System.Drawing.Image)
        Me.ToolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton2.Name = "ToolStripButton2"
        Me.ToolStripButton2.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton2.Text = "View Schema"
        '
        'ToolStripButton3
        '
        Me.ToolStripButton3.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton3.Image = CType(resources.GetObject("ToolStripButton3.Image"), System.Drawing.Image)
        Me.ToolStripButton3.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton3.Name = "ToolStripButton3"
        Me.ToolStripButton3.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton3.Text = "View Keys"
        '
        'ToolStripButton7
        '
        Me.ToolStripButton7.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton7.Image = CType(resources.GetObject("ToolStripButton7.Image"), System.Drawing.Image)
        Me.ToolStripButton7.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton7.Name = "ToolStripButton7"
        Me.ToolStripButton7.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton7.Text = "Get record counts"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'ToolStripButton5
        '
        Me.ToolStripButton5.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton5.Image = CType(resources.GetObject("ToolStripButton5.Image"), System.Drawing.Image)
        Me.ToolStripButton5.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton5.Name = "ToolStripButton5"
        Me.ToolStripButton5.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton5.Text = "Current SQL"
        '
        'ToolStripButton6
        '
        Me.ToolStripButton6.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton6.Image = CType(resources.GetObject("ToolStripButton6.Image"), System.Drawing.Image)
        Me.ToolStripButton6.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton6.Name = "ToolStripButton6"
        Me.ToolStripButton6.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton6.Text = "Save Query"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(6, 25)
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton1.Image = CType(resources.GetObject("ToolStripButton1.Image"), System.Drawing.Image)
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton1.Text = "Close this window"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnUseKey)
        Me.GroupBox1.Controls.Add(Me.ListBoxWhereClause)
        Me.GroupBox1.Controls.Add(Me.cboAndOr)
        Me.GroupBox1.Controls.Add(Me.btnAddSearch)
        Me.GroupBox1.Controls.Add(Me.cboSearchValue)
        Me.GroupBox1.Controls.Add(Me.cboBool)
        Me.GroupBox1.Controls.Add(Me.cboSelectedField)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 271)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(500, 152)
        Me.GroupBox1.TabIndex = 13
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Selection Criteria:"
        '
        'btnUseKey
        '
        Me.btnUseKey.Location = New System.Drawing.Point(130, 47)
        Me.btnUseKey.Name = "btnUseKey"
        Me.btnUseKey.Size = New System.Drawing.Size(75, 23)
        Me.btnUseKey.TabIndex = 3
        Me.btnUseKey.Text = "Use Key"
        Me.btnUseKey.UseVisualStyleBackColor = True
        '
        'ListBoxWhereClause
        '
        Me.ListBoxWhereClause.FormattingEnabled = True
        Me.ListBoxWhereClause.Location = New System.Drawing.Point(6, 77)
        Me.ListBoxWhereClause.Name = "ListBoxWhereClause"
        Me.ListBoxWhereClause.Size = New System.Drawing.Size(482, 69)
        Me.ListBoxWhereClause.TabIndex = 6
        '
        'cboAndOr
        '
        Me.cboAndOr.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAndOr.FormattingEnabled = True
        Me.cboAndOr.Items.AddRange(New Object() {" AND ", " OR "})
        Me.cboAndOr.Location = New System.Drawing.Point(212, 47)
        Me.cboAndOr.Name = "cboAndOr"
        Me.cboAndOr.Size = New System.Drawing.Size(62, 21)
        Me.cboAndOr.TabIndex = 4
        '
        'btnAddSearch
        '
        Me.btnAddSearch.Location = New System.Drawing.Point(281, 47)
        Me.btnAddSearch.Name = "btnAddSearch"
        Me.btnAddSearch.Size = New System.Drawing.Size(207, 23)
        Me.btnAddSearch.TabIndex = 5
        Me.btnAddSearch.Text = "&Add Search"
        Me.btnAddSearch.UseVisualStyleBackColor = True
        '
        'cboSearchValue
        '
        Me.cboSearchValue.FormattingEnabled = True
        Me.cboSearchValue.Items.AddRange(New Object() {"[parameter]"})
        Me.cboSearchValue.Location = New System.Drawing.Point(281, 19)
        Me.cboSearchValue.Name = "cboSearchValue"
        Me.cboSearchValue.Size = New System.Drawing.Size(207, 21)
        Me.cboSearchValue.TabIndex = 2
        '
        'cboBool
        '
        Me.cboBool.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboBool.FormattingEnabled = True
        Me.cboBool.Items.AddRange(New Object() {"=", "<>", ">", "<", "<=", ">="})
        Me.cboBool.Location = New System.Drawing.Point(213, 19)
        Me.cboBool.Name = "cboBool"
        Me.cboBool.Size = New System.Drawing.Size(62, 21)
        Me.cboBool.TabIndex = 1
        '
        'cboSelectedField
        '
        Me.cboSelectedField.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSelectedField.FormattingEnabled = True
        Me.cboSelectedField.Location = New System.Drawing.Point(6, 19)
        Me.cboSelectedField.Name = "cboSelectedField"
        Me.cboSelectedField.Size = New System.Drawing.Size(200, 21)
        Me.cboSelectedField.TabIndex = 0
        '
        'txtUserId
        '
        Me.txtUserId.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.txtUserId.Location = New System.Drawing.Point(176, 72)
        Me.txtUserId.Name = "txtUserId"
        Me.txtUserId.Size = New System.Drawing.Size(112, 20)
        Me.txtUserId.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(173, 56)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(46, 13)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "User ID:"
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
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(291, 56)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(56, 13)
        Me.Label6.TabIndex = 20
        Me.Label6.Text = "Password:"
        '
        'cboOrderBy
        '
        Me.cboOrderBy.FormattingEnabled = True
        Me.cboOrderBy.Location = New System.Drawing.Point(19, 440)
        Me.cboOrderBy.Name = "cboOrderBy"
        Me.cboOrderBy.Size = New System.Drawing.Size(200, 21)
        Me.cboOrderBy.TabIndex = 14
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(16, 426)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(50, 13)
        Me.Label7.TabIndex = 22
        Me.Label7.Text = "Order by:"
        '
        'rdoAscending
        '
        Me.rdoAscending.AutoSize = True
        Me.rdoAscending.Checked = True
        Me.rdoAscending.Location = New System.Drawing.Point(224, 431)
        Me.rdoAscending.Name = "rdoAscending"
        Me.rdoAscending.Size = New System.Drawing.Size(75, 17)
        Me.rdoAscending.TabIndex = 15
        Me.rdoAscending.TabStop = True
        Me.rdoAscending.Text = "Ascending"
        Me.rdoAscending.UseVisualStyleBackColor = True
        '
        'rdoDescending
        '
        Me.rdoDescending.AutoSize = True
        Me.rdoDescending.Location = New System.Drawing.Point(224, 451)
        Me.rdoDescending.Name = "rdoDescending"
        Me.rdoDescending.Size = New System.Drawing.Size(82, 17)
        Me.rdoDescending.TabIndex = 16
        Me.rdoDescending.Text = "Descending"
        Me.rdoDescending.UseVisualStyleBackColor = True
        '
        'ckbDistinct
        '
        Me.ckbDistinct.AutoSize = True
        Me.ckbDistinct.Location = New System.Drawing.Point(313, 443)
        Me.ckbDistinct.Name = "ckbDistinct"
        Me.ckbDistinct.Size = New System.Drawing.Size(67, 17)
        Me.ckbDistinct.TabIndex = 17
        Me.ckbDistinct.Text = "Distinct?"
        Me.ckbDistinct.UseVisualStyleBackColor = True
        '
        'frmNewQuery
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(528, 476)
        Me.Controls.Add(Me.ckbDistinct)
        Me.Controls.Add(Me.rdoDescending)
        Me.Controls.Add(Me.rdoAscending)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.cboOrderBy)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtUserId)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnRunQuery)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.ListBoxSelectedFields)
        Me.Controls.Add(Me.cboTables)
        Me.Controls.Add(Me.ListBoxFields)
        Me.Controls.Add(Me.btnConnect)
        Me.Controls.Add(Me.cboDSNs)
        Me.Controls.Add(Me.MenuStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MaximizeBox = False
        Me.Name = "frmNewQuery"
        Me.ShowInTaskbar = False
        Me.Text = "New Query"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cboDSNs As System.Windows.Forms.ComboBox
    Friend WithEvents btnConnect As System.Windows.Forms.Button
    Friend WithEvents ListBoxFields As System.Windows.Forms.ListBox
    Friend WithEvents cboTables As System.Windows.Forms.ComboBox
    Friend WithEvents ListBoxSelectedFields As System.Windows.Forms.ListBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents btnRunQuery As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripButton1 As System.Windows.Forms.ToolStripButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cboSelectedField As System.Windows.Forms.ComboBox
    Friend WithEvents cboBool As System.Windows.Forms.ComboBox
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnAddSearch As System.Windows.Forms.Button
    Friend WithEvents cboSearchValue As System.Windows.Forms.ComboBox
    Friend WithEvents ListBoxWhereClause As System.Windows.Forms.ListBox
    Friend WithEvents cboAndOr As System.Windows.Forms.ComboBox
    Friend WithEvents txtUserId As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnUseKey As System.Windows.Forms.Button
    Friend WithEvents ToolStripButton2 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripButton3 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ViewToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CurrentSQLToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripButton5 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents cboOrderBy As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents rdoAscending As System.Windows.Forms.RadioButton
    Friend WithEvents rdoDescending As System.Windows.Forms.RadioButton
    Friend WithEvents ckbDistinct As System.Windows.Forms.CheckBox
    Friend WithEvents SaveToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator5 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripButton6 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButton7 As System.Windows.Forms.ToolStripButton
    Friend WithEvents RecordCountToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
