<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLotFileDisp
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
        Me.ListFile = New System.Windows.Forms.ListBox()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ListFile
        '
        Me.ListFile.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ListFile.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.ListFile.FormattingEnabled = True
        Me.ListFile.HorizontalScrollbar = True
        Me.ListFile.ItemHeight = 19
        Me.ListFile.Location = New System.Drawing.Point(12, 12)
        Me.ListFile.Name = "ListFile"
        Me.ListFile.Size = New System.Drawing.Size(561, 213)
        Me.ListFile.TabIndex = 3
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(401, 245)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(142, 51)
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'frmLotFileDisp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(585, 322)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.ListFile)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "frmLotFileDisp"
        Me.Text = "自動運転登録ファイル"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ListFile As ListBox
    Friend WithEvents btnClose As Button
End Class
