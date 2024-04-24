<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class formReStartPosSet
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
        Me.LblPlateNo = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtPlateNo = New System.Windows.Forms.TextBox()
        Me.txtBlockX = New System.Windows.Forms.TextBox()
        Me.txtBlockY = New System.Windows.Forms.TextBox()
        Me.BtnOK = New System.Windows.Forms.Button()
        Me.BtnCancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'LblPlateNo
        '
        Me.LblPlateNo.AutoSize = True
        Me.LblPlateNo.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.LblPlateNo.Location = New System.Drawing.Point(57, 45)
        Me.LblPlateNo.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LblPlateNo.Name = "LblPlateNo"
        Me.LblPlateNo.Size = New System.Drawing.Size(72, 16)
        Me.LblPlateNo.TabIndex = 0
        Me.LblPlateNo.Text = "基板番号"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.Label1.Location = New System.Drawing.Point(57, 90)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(18, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "X"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.Label2.Location = New System.Drawing.Point(58, 135)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(17, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Y"
        '
        'txtPlateNo
        '
        Me.txtPlateNo.Font = New System.Drawing.Font("ＭＳ ゴシック", 14.25!, System.Drawing.FontStyle.Bold)
        Me.txtPlateNo.Location = New System.Drawing.Point(192, 35)
        Me.txtPlateNo.Margin = New System.Windows.Forms.Padding(4)
        Me.txtPlateNo.Name = "txtPlateNo"
        Me.txtPlateNo.Size = New System.Drawing.Size(100, 26)
        Me.txtPlateNo.TabIndex = 3
        '
        'txtBlockX
        '
        Me.txtBlockX.Font = New System.Drawing.Font("ＭＳ ゴシック", 14.25!, System.Drawing.FontStyle.Bold)
        Me.txtBlockX.Location = New System.Drawing.Point(192, 80)
        Me.txtBlockX.Margin = New System.Windows.Forms.Padding(4)
        Me.txtBlockX.MaxLength = 3
        Me.txtBlockX.Name = "txtBlockX"
        Me.txtBlockX.Size = New System.Drawing.Size(100, 26)
        Me.txtBlockX.TabIndex = 4
        '
        'txtBlockY
        '
        Me.txtBlockY.Font = New System.Drawing.Font("ＭＳ ゴシック", 14.25!, System.Drawing.FontStyle.Bold)
        Me.txtBlockY.Location = New System.Drawing.Point(193, 125)
        Me.txtBlockY.Margin = New System.Windows.Forms.Padding(4)
        Me.txtBlockY.MaxLength = 3
        Me.txtBlockY.Name = "txtBlockY"
        Me.txtBlockY.Size = New System.Drawing.Size(100, 26)
        Me.txtBlockY.TabIndex = 5
        '
        'BtnOK
        '
        Me.BtnOK.Location = New System.Drawing.Point(177, 185)
        Me.BtnOK.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(116, 35)
        Me.BtnOK.TabIndex = 6
        Me.BtnOK.Text = "OK"
        Me.BtnOK.UseVisualStyleBackColor = True
        '
        'BtnCancel
        '
        Me.BtnCancel.Location = New System.Drawing.Point(317, 185)
        Me.BtnCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(116, 35)
        Me.BtnCancel.TabIndex = 7
        Me.BtnCancel.Text = "Cancel"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'formReStartPosSet
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(500, 247)
        Me.ControlBox = False
        Me.Controls.Add(Me.BtnCancel)
        Me.Controls.Add(Me.BtnOK)
        Me.Controls.Add(Me.txtBlockY)
        Me.Controls.Add(Me.txtBlockX)
        Me.Controls.Add(Me.txtPlateNo)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.LblPlateNo)
        Me.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.Name = "formReStartPosSet"
        Me.Text = "再測定開始位置指定"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LblPlateNo As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtPlateNo As System.Windows.Forms.TextBox
    Friend WithEvents txtBlockX As System.Windows.Forms.TextBox
    Friend WithEvents txtBlockY As System.Windows.Forms.TextBox
    Friend WithEvents BtnOK As System.Windows.Forms.Button
    Friend WithEvents BtnCancel As System.Windows.Forms.Button
End Class
