<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMessage
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
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

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.BtnCancel = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Grp1 = New System.Windows.Forms.GroupBox()
        Me.BtnOK = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.LblCaption = New System.Windows.Forms.Label()
        Me.Grp1.SuspendLayout()
        Me.SuspendLayout()
        '
        'BtnCancel
        '
        Me.BtnCancel.Font = New System.Drawing.Font("ＭＳ ゴシック", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnCancel.Location = New System.Drawing.Point(329, 163)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(130, 31)
        Me.BtnCancel.TabIndex = 4
        Me.BtnCancel.Text = "Cancel"
        Me.BtnCancel.UseVisualStyleBackColor = True
        Me.BtnCancel.Visible = False
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("ＭＳ ゴシック", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.Location = New System.Drawing.Point(7, 113)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(554, 44)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Label2"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Grp1
        '
        Me.Grp1.Controls.Add(Me.BtnOK)
        Me.Grp1.Controls.Add(Me.BtnCancel)
        Me.Grp1.Controls.Add(Me.Label2)
        Me.Grp1.Controls.Add(Me.Label1)
        Me.Grp1.Controls.Add(Me.LblCaption)
        Me.Grp1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Grp1.Location = New System.Drawing.Point(0, -5)
        Me.Grp1.Name = "Grp1"
        Me.Grp1.Size = New System.Drawing.Size(568, 203)
        Me.Grp1.TabIndex = 2
        Me.Grp1.TabStop = False
        '
        'BtnOK
        '
        Me.BtnOK.Font = New System.Drawing.Font("ＭＳ ゴシック", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnOK.Location = New System.Drawing.Point(107, 162)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(130, 31)
        Me.BtnOK.TabIndex = 5
        Me.BtnOK.Text = "OK"
        Me.BtnOK.UseVisualStyleBackColor = True
        Me.BtnOK.Visible = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Font = New System.Drawing.Font("ＭＳ ゴシック", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(7, 63)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(554, 44)
        Me.Label1.TabIndex = 2
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'LblCaption
        '
        Me.LblCaption.BackColor = System.Drawing.SystemColors.Control
        Me.LblCaption.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.LblCaption.Location = New System.Drawing.Point(7, 14)
        Me.LblCaption.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LblCaption.Name = "LblCaption"
        Me.LblCaption.Size = New System.Drawing.Size(554, 44)
        Me.LblCaption.TabIndex = 1
        Me.LblCaption.Text = "原点復帰中"
        Me.LblCaption.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'frmMessage
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(570, 204)
        Me.ControlBox = False
        Me.Controls.Add(Me.Grp1)
        Me.Font = New System.Drawing.Font("ＭＳ ゴシック", 12.0!)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "frmMessage"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmMessage"
        Me.TopMost = True
        Me.Grp1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents BtnCancel As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Grp1 As System.Windows.Forms.GroupBox
    Friend WithEvents BtnOK As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LblCaption As System.Windows.Forms.Label
End Class
