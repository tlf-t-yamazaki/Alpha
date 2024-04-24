<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmTxTyMsgbox
#Region "Windows フォーム デザイナによって生成されたコード "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'この呼び出しは、Windows フォーム デザイナで必要です。
		InitializeComponent()
		Form_Initialize_renamed()
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
	Public WithEvents cmdOKTxTy As System.Windows.Forms.Button
	Public WithEvents cmdCAN As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents Label1 As System.Windows.Forms.Label
	'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
	'Windows フォーム デザイナを使って変更できます。
	'コード エディタを使用して、変更しないでください。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTxTyMsgbox))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdOKTxTy = New System.Windows.Forms.Button
        Me.cmdCAN = New System.Windows.Forms.Button
        Me.cmdOK = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'cmdOKTxTy
        '
        Me.cmdOKTxTy.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOKTxTy.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOKTxTy.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOKTxTy.Location = New System.Drawing.Point(124, 112)
        Me.cmdOKTxTy.Name = "cmdOKTxTy"
        Me.cmdOKTxTy.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOKTxTy.Size = New System.Drawing.Size(89, 25)
        Me.cmdOKTxTy.TabIndex = 3
        Me.cmdOKTxTy.Text = "TX(&T)"
        Me.cmdOKTxTy.UseVisualStyleBackColor = False
        Me.cmdOKTxTy.Visible = False
        '
        'cmdCAN
        '
        Me.cmdCAN.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCAN.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCAN.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCAN.Location = New System.Drawing.Point(232, 112)
        Me.cmdCAN.Name = "cmdCAN"
        Me.cmdCAN.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCAN.Size = New System.Drawing.Size(89, 25)
        Me.cmdCAN.TabIndex = 1
        Me.cmdCAN.Text = "いいえ（&N)"
        Me.cmdCAN.UseVisualStyleBackColor = False
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.Location = New System.Drawing.Point(16, 112)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.Size = New System.Drawing.Size(89, 25)
        Me.cmdOK.TabIndex = 0
        Me.cmdOK.Text = "はい(&Y)"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(16, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(297, 33)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Label1"
        '
        'frmTxTyMsgbox
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(337, 160)
        Me.Controls.Add(Me.cmdOKTxTy)
        Me.Controls.Add(Me.cmdCAN)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(184, 250)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmTxTyMsgbox"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "画面終了確認"
        Me.ResumeLayout(False)

    End Sub
#End Region 
End Class