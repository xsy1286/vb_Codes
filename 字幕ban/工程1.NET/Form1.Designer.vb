<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Form1
#Region "Windows 窗体设计器生成的代码 "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'此调用是 Windows 窗体设计器所必需的。
		InitializeComponent()
	End Sub
	'窗体重写释放，以清理组件列表。
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Windows 窗体设计器所必需的
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents HScroll2 As System.Windows.Forms.HScrollBar
	Public WithEvents HScroll1 As System.Windows.Forms.HScrollBar
	Public WithEvents Check2 As System.Windows.Forms.CheckBox
	Public WithEvents Check1 As System.Windows.Forms.CheckBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	'注意: 以下过程是 Windows 窗体设计器所必需的
	'可以使用 Windows 窗体设计器来修改它。
	'不要使用代码编辑器修改它。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.HScroll2 = New System.Windows.Forms.HScrollBar
        Me.HScroll1 = New System.Windows.Forms.HScrollBar
        Me.Check2 = New System.Windows.Forms.CheckBox
        Me.Check1 = New System.Windows.Forms.CheckBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Timer1
        '
        Me.Timer1.Interval = 1
        '
        'HScroll2
        '
        Me.HScroll2.Cursor = System.Windows.Forms.Cursors.Default
        Me.HScroll2.LargeChange = 1
        Me.HScroll2.Location = New System.Drawing.Point(16, 48)
        Me.HScroll2.Maximum = 800
        Me.HScroll2.Minimum = 100
        Me.HScroll2.Name = "HScroll2"
        Me.HScroll2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.HScroll2.Size = New System.Drawing.Size(145, 17)
        Me.HScroll2.TabIndex = 3
        Me.HScroll2.TabStop = True
        Me.HScroll2.Value = 100
        '
        'HScroll1
        '
        Me.HScroll1.Cursor = System.Windows.Forms.Cursors.Default
        Me.HScroll1.LargeChange = 1
        Me.HScroll1.Location = New System.Drawing.Point(16, 32)
        Me.HScroll1.Maximum = 1280
        Me.HScroll1.Minimum = 300
        Me.HScroll1.Name = "HScroll1"
        Me.HScroll1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.HScroll1.Size = New System.Drawing.Size(145, 17)
        Me.HScroll1.TabIndex = 2
        Me.HScroll1.TabStop = True
        Me.HScroll1.Value = 300
        '
        'Check2
        '
        Me.Check2.BackColor = System.Drawing.SystemColors.Control
        Me.Check2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Check2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Check2.Location = New System.Drawing.Point(192, 48)
        Me.Check2.Name = "Check2"
        Me.Check2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Check2.Size = New System.Drawing.Size(73, 17)
        Me.Check2.TabIndex = 1
        Me.Check2.Text = "可以移动"
        Me.Check2.UseVisualStyleBackColor = False
        '
        'Check1
        '
        Me.Check1.BackColor = System.Drawing.SystemColors.Control
        Me.Check1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Check1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Check1.Location = New System.Drawing.Point(192, 16)
        Me.Check1.Name = "Check1"
        Me.Check1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Check1.Size = New System.Drawing.Size(56, 26)
        Me.Check1.TabIndex = 0
        Me.Check1.Text = "置顶"
        Me.Check1.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(40, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(105, 17)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "    窗口大小"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(315, 77)
        Me.Controls.Add(Me.HScroll2)
        Me.Controls.Add(Me.HScroll1)
        Me.Controls.Add(Me.Check2)
        Me.Controls.Add(Me.Check1)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form1"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub
#End Region 
End Class