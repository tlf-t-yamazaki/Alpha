<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmLotInp
#Region "Windows フォーム デザイナによって生成されたコード "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'この呼び出しは、Windows フォーム デザイナで必要です。
		InitializeComponent()
	End Sub
	'Form は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Windows フォーム デザイナで必要です。
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents CmndCancel As System.Windows.Forms.Button
	Public WithEvents CmndOk As System.Windows.Forms.Button
	Public WithEvents TextLOT As System.Windows.Forms.TextBox
	Public WithEvents _LblLOT_2 As System.Windows.Forms.Label
	Public WithEvents _Lbl_0 As System.Windows.Forms.Label
	Public WithEvents DataInp As System.Windows.Forms.GroupBox
    'Public WithEvents Lbl As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'Public WithEvents LblLOT As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使って変更できます。
    'コード エディタを使用して、変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmLotInp))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.CmndCancel = New System.Windows.Forms.Button
		Me.CmndOk = New System.Windows.Forms.Button
		Me.DataInp = New System.Windows.Forms.GroupBox
		Me.TextLOT = New System.Windows.Forms.TextBox
		Me._LblLOT_2 = New System.Windows.Forms.Label
		Me._Lbl_0 = New System.Windows.Forms.Label
        'Me.Lbl = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
        'Me.LblLOT = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
        Me.DataInp.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
        'CType(Me.Lbl, System.ComponentModel.ISupportInitialize).BeginInit()
        'CType(Me.LblLOT, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.ClientSize = New System.Drawing.Size(325, 154)
		Me.Location = New System.Drawing.Point(312, 281)
		Me.KeyPreview = True
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.Enabled = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmLotInp"
		Me.CmndCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.CmndCancel
		Me.CmndCancel.Text = "Cancel"
		Me.CmndCancel.CausesValidation = False
		Me.CmndCancel.Size = New System.Drawing.Size(73, 25)
		Me.CmndCancel.Location = New System.Drawing.Point(248, 128)
		Me.CmndCancel.TabIndex = 2
		Me.CmndCancel.BackColor = System.Drawing.SystemColors.Control
		Me.CmndCancel.Enabled = True
		Me.CmndCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.CmndCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.CmndCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.CmndCancel.TabStop = True
		Me.CmndCancel.Name = "CmndCancel"
		Me.CmndOk.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CmndOk.Text = "OK"
		Me.CmndOk.CausesValidation = False
		Me.AcceptButton = Me.CmndOk
		Me.CmndOk.Size = New System.Drawing.Size(73, 25)
		Me.CmndOk.Location = New System.Drawing.Point(168, 128)
		Me.CmndOk.TabIndex = 1
		Me.CmndOk.BackColor = System.Drawing.SystemColors.Control
		Me.CmndOk.Enabled = True
		Me.CmndOk.ForeColor = System.Drawing.SystemColors.ControlText
		Me.CmndOk.Cursor = System.Windows.Forms.Cursors.Default
		Me.CmndOk.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.CmndOk.TabStop = True
		Me.CmndOk.Name = "CmndOk"
		Me.DataInp.Size = New System.Drawing.Size(312, 107)
		Me.DataInp.Location = New System.Drawing.Point(8, 8)
		Me.DataInp.TabIndex = 3
		Me.DataInp.BackColor = System.Drawing.SystemColors.Control
		Me.DataInp.Enabled = True
		Me.DataInp.ForeColor = System.Drawing.SystemColors.ControlText
		Me.DataInp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.DataInp.Visible = True
		Me.DataInp.Name = "DataInp"
		Me.TextLOT.AutoSize = False
		Me.TextLOT.CausesValidation = False
		Me.TextLOT.Size = New System.Drawing.Size(225, 18)
		Me.TextLOT.Location = New System.Drawing.Point(72, 48)
		Me.TextLOT.Maxlength = 20
		Me.TextLOT.TabIndex = 0
		Me.TextLOT.AcceptsReturn = True
		Me.TextLOT.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.TextLOT.BackColor = System.Drawing.SystemColors.Window
		Me.TextLOT.Enabled = True
		Me.TextLOT.ForeColor = System.Drawing.SystemColors.WindowText
		Me.TextLOT.HideSelection = True
		Me.TextLOT.ReadOnly = False
		Me.TextLOT.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.TextLOT.MultiLine = False
		Me.TextLOT.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.TextLOT.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.TextLOT.TabStop = True
		Me.TextLOT.Visible = True
		Me.TextLOT.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TextLOT.Name = "TextLOT"
		Me._LblLOT_2.Text = "ロット番号"
		Me._LblLOT_2.Size = New System.Drawing.Size(65, 17)
		Me._LblLOT_2.Location = New System.Drawing.Point(16, 48)
		Me._LblLOT_2.TabIndex = 5
		Me._LblLOT_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._LblLOT_2.BackColor = System.Drawing.SystemColors.Control
		Me._LblLOT_2.Enabled = True
		Me._LblLOT_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._LblLOT_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._LblLOT_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._LblLOT_2.UseMnemonic = True
		Me._LblLOT_2.Visible = True
		Me._LblLOT_2.AutoSize = False
		Me._LblLOT_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._LblLOT_2.Name = "_LblLOT_2"
		Me._Lbl_0.Text = "ロット番号を入力して下さい。"
		Me._Lbl_0.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
		Me._Lbl_0.Size = New System.Drawing.Size(233, 17)
		Me._Lbl_0.Location = New System.Drawing.Point(8, 16)
		Me._Lbl_0.TabIndex = 4
		Me._Lbl_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Lbl_0.BackColor = System.Drawing.SystemColors.Control
		Me._Lbl_0.Enabled = True
		Me._Lbl_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Lbl_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Lbl_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Lbl_0.UseMnemonic = True
		Me._Lbl_0.Visible = True
		Me._Lbl_0.AutoSize = False
		Me._Lbl_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Lbl_0.Name = "_Lbl_0"
		Me.Controls.Add(CmndCancel)
		Me.Controls.Add(CmndOk)
		Me.Controls.Add(DataInp)
		Me.DataInp.Controls.Add(TextLOT)
		Me.DataInp.Controls.Add(_LblLOT_2)
		Me.DataInp.Controls.Add(_Lbl_0)
        'Me.Lbl.SetIndex(_Lbl_0, CType(0, Short))
        'Me.LblLOT.SetIndex(_LblLOT_2, CType(2, Short))
        'CType(Me.LblLOT, System.ComponentModel.ISupportInitialize).EndInit()
        'CType(Me.Lbl, System.ComponentModel.ISupportInitialize).EndInit()
        Me.DataInp.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class