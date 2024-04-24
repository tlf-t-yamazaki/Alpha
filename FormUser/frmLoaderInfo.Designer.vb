<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmLoaderInfo
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.lblTact = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.lblSupplySlot = New System.Windows.Forms.Label()
        Me.lblSupplyMag = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblExchange = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.lblTrimming = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblTact
        '
        Me.lblTact.AutoSize = True
        Me.lblTact.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblTact.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblTact.Location = New System.Drawing.Point(160, 1)
        Me.lblTact.Name = "lblTact"
        Me.lblTact.Size = New System.Drawing.Size(98, 29)
        Me.lblTact.TabIndex = 1
        Me.lblTact.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label4.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label4.Location = New System.Drawing.Point(4, 1)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(149, 29)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "タクトタイム"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.[Single]
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 60.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 40.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.lblSupplySlot, 1, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.lblSupplyMag, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Label4, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.lblTact, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label6, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.lblExchange, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label8, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.lblTrimming, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Label3, 0, 4)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(2, 2)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 5
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(262, 152)
        Me.TableLayoutPanel1.TabIndex = 1
        '
        'lblSupplySlot
        '
        Me.lblSupplySlot.AutoSize = True
        Me.lblSupplySlot.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblSupplySlot.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblSupplySlot.Location = New System.Drawing.Point(160, 121)
        Me.lblSupplySlot.Name = "lblSupplySlot"
        Me.lblSupplySlot.Size = New System.Drawing.Size(98, 30)
        Me.lblSupplySlot.TabIndex = 12
        Me.lblSupplySlot.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblSupplyMag
        '
        Me.lblSupplyMag.AutoSize = True
        Me.lblSupplyMag.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblSupplyMag.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblSupplyMag.Location = New System.Drawing.Point(160, 91)
        Me.lblSupplyMag.Name = "lblSupplyMag"
        Me.lblSupplyMag.Size = New System.Drawing.Size(98, 29)
        Me.lblSupplyMag.TabIndex = 11
        Me.lblSupplyMag.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.AutoEllipsis = True
        Me.Label6.AutoSize = True
        Me.Label6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label6.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label6.Location = New System.Drawing.Point(4, 31)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(149, 29)
        Me.Label6.TabIndex = 2
        Me.Label6.Text = "基板交換時間"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblExchange
        '
        Me.lblExchange.AutoSize = True
        Me.lblExchange.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblExchange.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblExchange.Location = New System.Drawing.Point(160, 31)
        Me.lblExchange.Name = "lblExchange"
        Me.lblExchange.Size = New System.Drawing.Size(98, 29)
        Me.lblExchange.TabIndex = 3
        Me.lblExchange.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label8.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label8.Location = New System.Drawing.Point(4, 61)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(149, 29)
        Me.Label8.TabIndex = 4
        Me.Label8.Text = "トリミング時間"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblTrimming
        '
        Me.lblTrimming.AutoSize = True
        Me.lblTrimming.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblTrimming.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblTrimming.Location = New System.Drawing.Point(160, 61)
        Me.lblTrimming.Name = "lblTrimming"
        Me.lblTrimming.Size = New System.Drawing.Size(98, 29)
        Me.lblTrimming.TabIndex = 5
        Me.lblTrimming.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(22, 96)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 19)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "供給マガジン"
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.Location = New System.Drawing.Point(19, 126)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(118, 19)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "トリミング枚数"
        '
        'frmLoaderInfo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(268, 156)
        Me.ControlBox = False
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmLoaderInfo"
        Me.Text = "Loader Information"
        Me.TopMost = True
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lblTact As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents Label6 As Label
    Friend WithEvents lblExchange As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents lblTrimming As Label
    Friend WithEvents lblSupplySlot As Label
    Friend WithEvents lblSupplyMag As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Label3 As Label
End Class
