<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFineAdjust
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

    'Windows フォーム デザイナで必要です。

    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。

    'Windows フォーム デザイナを使用して変更できます。  
    'コード エディタを使って変更しないでください。

    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFineAdjust))
        Me.btnTrimming = New System.Windows.Forms.Button()
        Me.grpBpOff = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtBpOffY = New System.Windows.Forms.TextBox()
        Me.txtBpOffX = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.BtnADJ = New System.Windows.Forms.Button()
        Me.GrpArrow = New System.Windows.Forms.GroupBox()
        Me.BtnJOG_6 = New System.Windows.Forms.Button()
        Me.BtnJOG_5 = New System.Windows.Forms.Button()
        Me.BtnJOG_7 = New System.Windows.Forms.Button()
        Me.BtnJOG_4 = New System.Windows.Forms.Button()
        Me.BtnHI = New System.Windows.Forms.Button()
        Me.BtnJOG_3 = New System.Windows.Forms.Button()
        Me.BtnJOG_2 = New System.Windows.Forms.Button()
        Me.BtnJOG_1 = New System.Windows.Forms.Button()
        Me.BtnJOG_0 = New System.Windows.Forms.Button()
        Me.BtnZ = New System.Windows.Forms.Button()
        Me.BtnRESET = New System.Windows.Forms.Button()
        Me.BtnSTART = New System.Windows.Forms.Button()
        Me.BtnHALT = New System.Windows.Forms.Button()
        Me.GrpPithPanel = New System.Windows.Forms.GroupBox()
        Me.TBarPause = New System.Windows.Forms.TrackBar()
        Me.TBarHiPitch = New System.Windows.Forms.TrackBar()
        Me.TBarLowPitch = New System.Windows.Forms.TrackBar()
        Me.LblTchMoval2 = New System.Windows.Forms.Label()
        Me.LblTchMoval1 = New System.Windows.Forms.Label()
        Me.LblTchMoval0 = New System.Windows.Forms.Label()
        Me.LblPitch2 = New System.Windows.Forms.Label()
        Me.LblPitch1 = New System.Windows.Forms.Label()
        Me.LblPitch0 = New System.Windows.Forms.Label()
        Me.TmKeyCheck = New System.Windows.Forms.Timer(Me.components)
        Me.BtnTenKey = New System.Windows.Forms.Button()
        Me.BtnLaser = New System.Windows.Forms.Button()
        Me.BtnClickEnable = New System.Windows.Forms.Button()
        Me.BtnLoaderInfo = New System.Windows.Forms.Button()
        Me.btnAutoInfo = New System.Windows.Forms.Button()
        Me.grpBpOff.SuspendLayout()
        Me.GrpArrow.SuspendLayout()
        Me.GrpPithPanel.SuspendLayout()
        CType(Me.TBarPause, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TBarHiPitch, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TBarLowPitch, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnTrimming
        '
        Me.btnTrimming.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.btnTrimming.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnTrimming.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnTrimming.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnTrimming.Location = New System.Drawing.Point(158, 141)
        Me.btnTrimming.Name = "btnTrimming"
        Me.btnTrimming.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnTrimming.Size = New System.Drawing.Size(145, 37)
        Me.btnTrimming.TabIndex = 278
        Me.btnTrimming.Text = "START"
        Me.btnTrimming.UseVisualStyleBackColor = False
        Me.btnTrimming.Visible = False
        '
        'grpBpOff
        '
        Me.grpBpOff.Controls.Add(Me.Label2)
        Me.grpBpOff.Controls.Add(Me.Label1)
        Me.grpBpOff.Controls.Add(Me.txtBpOffY)
        Me.grpBpOff.Controls.Add(Me.txtBpOffX)
        Me.grpBpOff.Font = New System.Drawing.Font("ＭＳ ゴシック", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.grpBpOff.Location = New System.Drawing.Point(4, 50)
        Me.grpBpOff.Name = "grpBpOff"
        Me.grpBpOff.Size = New System.Drawing.Size(595, 76)
        Me.grpBpOff.TabIndex = 283
        Me.grpBpOff.TabStop = False
        Me.grpBpOff.Text = "Beam Position Offset"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(346, 36)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(98, 16)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Y Position"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(23, 36)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(98, 16)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "X Position"
        '
        'txtBpOffY
        '
        Me.txtBpOffY.Location = New System.Drawing.Point(450, 33)
        Me.txtBpOffY.Name = "txtBpOffY"
        Me.txtBpOffY.ReadOnly = True
        Me.txtBpOffY.Size = New System.Drawing.Size(127, 23)
        Me.txtBpOffY.TabIndex = 1
        Me.txtBpOffY.TabStop = False
        Me.txtBpOffY.Text = "0"
        '
        'txtBpOffX
        '
        Me.txtBpOffX.Location = New System.Drawing.Point(149, 33)
        Me.txtBpOffX.Name = "txtBpOffX"
        Me.txtBpOffX.ReadOnly = True
        Me.txtBpOffX.Size = New System.Drawing.Size(127, 23)
        Me.txtBpOffX.TabIndex = 0
        Me.txtBpOffX.TabStop = False
        Me.txtBpOffX.Text = "0"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("ＭＳ ゴシック", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.Location = New System.Drawing.Point(7, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(70, 27)
        Me.Label3.TabIndex = 284
        Me.Label3.Text = "調整"
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.Text = "NotifyIcon1"
        Me.NotifyIcon1.Visible = True
        '
        'BtnADJ
        '
        Me.BtnADJ.Enabled = False
        Me.BtnADJ.Font = New System.Drawing.Font("ＭＳ ゴシック", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnADJ.Location = New System.Drawing.Point(450, 141)
        Me.BtnADJ.Name = "BtnADJ"
        Me.BtnADJ.Size = New System.Drawing.Size(143, 37)
        Me.BtnADJ.TabIndex = 0
        Me.BtnADJ.Text = "ADJ ON"
        Me.BtnADJ.UseVisualStyleBackColor = True
        Me.BtnADJ.Visible = False
        '
        'GrpArrow
        '
        Me.GrpArrow.Controls.Add(Me.BtnJOG_6)
        Me.GrpArrow.Controls.Add(Me.BtnJOG_5)
        Me.GrpArrow.Controls.Add(Me.BtnJOG_7)
        Me.GrpArrow.Controls.Add(Me.BtnJOG_4)
        Me.GrpArrow.Controls.Add(Me.BtnHI)
        Me.GrpArrow.Controls.Add(Me.BtnJOG_3)
        Me.GrpArrow.Controls.Add(Me.BtnJOG_2)
        Me.GrpArrow.Controls.Add(Me.BtnJOG_1)
        Me.GrpArrow.Controls.Add(Me.BtnJOG_0)
        Me.GrpArrow.Controls.Add(Me.BtnZ)
        Me.GrpArrow.Controls.Add(Me.BtnRESET)
        Me.GrpArrow.Controls.Add(Me.BtnSTART)
        Me.GrpArrow.Controls.Add(Me.BtnHALT)
        Me.GrpArrow.Controls.Add(Me.GrpPithPanel)
        Me.GrpArrow.Font = New System.Drawing.Font("ＭＳ ゴシック", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.GrpArrow.Location = New System.Drawing.Point(4, 271)
        Me.GrpArrow.Name = "GrpArrow"
        Me.GrpArrow.Size = New System.Drawing.Size(608, 250)
        Me.GrpArrow.TabIndex = 328
        Me.GrpArrow.TabStop = False
        '
        'BtnJOG_6
        '
        Me.BtnJOG_6.BackColor = System.Drawing.SystemColors.Control
        Me.BtnJOG_6.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnJOG_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnJOG_6.Image = CType(resources.GetObject("BtnJOG_6.Image"), System.Drawing.Image)
        Me.BtnJOG_6.Location = New System.Drawing.Point(19, 163)
        Me.BtnJOG_6.Name = "BtnJOG_6"
        Me.BtnJOG_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.BtnJOG_6.Size = New System.Drawing.Size(73, 73)
        Me.BtnJOG_6.TabIndex = 50
        Me.BtnJOG_6.TabStop = False
        Me.BtnJOG_6.UseVisualStyleBackColor = False
        '
        'BtnJOG_5
        '
        Me.BtnJOG_5.BackColor = System.Drawing.SystemColors.Control
        Me.BtnJOG_5.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnJOG_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnJOG_5.Image = CType(resources.GetObject("BtnJOG_5.Image"), System.Drawing.Image)
        Me.BtnJOG_5.Location = New System.Drawing.Point(19, 17)
        Me.BtnJOG_5.Name = "BtnJOG_5"
        Me.BtnJOG_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.BtnJOG_5.Size = New System.Drawing.Size(73, 73)
        Me.BtnJOG_5.TabIndex = 49
        Me.BtnJOG_5.TabStop = False
        Me.BtnJOG_5.UseVisualStyleBackColor = False
        '
        'BtnJOG_7
        '
        Me.BtnJOG_7.BackColor = System.Drawing.SystemColors.Control
        Me.BtnJOG_7.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnJOG_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnJOG_7.Image = CType(resources.GetObject("BtnJOG_7.Image"), System.Drawing.Image)
        Me.BtnJOG_7.Location = New System.Drawing.Point(166, 163)
        Me.BtnJOG_7.Name = "BtnJOG_7"
        Me.BtnJOG_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.BtnJOG_7.Size = New System.Drawing.Size(73, 73)
        Me.BtnJOG_7.TabIndex = 48
        Me.BtnJOG_7.TabStop = False
        Me.BtnJOG_7.UseVisualStyleBackColor = False
        '
        'BtnJOG_4
        '
        Me.BtnJOG_4.AllowDrop = True
        Me.BtnJOG_4.BackColor = System.Drawing.SystemColors.Control
        Me.BtnJOG_4.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnJOG_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnJOG_4.Image = CType(resources.GetObject("BtnJOG_4.Image"), System.Drawing.Image)
        Me.BtnJOG_4.Location = New System.Drawing.Point(166, 17)
        Me.BtnJOG_4.Name = "BtnJOG_4"
        Me.BtnJOG_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.BtnJOG_4.Size = New System.Drawing.Size(73, 73)
        Me.BtnJOG_4.TabIndex = 47
        Me.BtnJOG_4.TabStop = False
        Me.BtnJOG_4.UseVisualStyleBackColor = False
        '
        'BtnHI
        '
        Me.BtnHI.AutoSize = True
        Me.BtnHI.BackColor = System.Drawing.SystemColors.Control
        Me.BtnHI.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnHI.Font = New System.Drawing.Font("ＭＳ ゴシック", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnHI.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnHI.Location = New System.Drawing.Point(93, 90)
        Me.BtnHI.Name = "BtnHI"
        Me.BtnHI.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.BtnHI.Size = New System.Drawing.Size(73, 73)
        Me.BtnHI.TabIndex = 46
        Me.BtnHI.TabStop = False
        Me.BtnHI.Text = "HI"
        Me.BtnHI.UseVisualStyleBackColor = False
        '
        'BtnJOG_3
        '
        Me.BtnJOG_3.BackColor = System.Drawing.SystemColors.Control
        Me.BtnJOG_3.CausesValidation = False
        Me.BtnJOG_3.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnJOG_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnJOG_3.Image = CType(resources.GetObject("BtnJOG_3.Image"), System.Drawing.Image)
        Me.BtnJOG_3.Location = New System.Drawing.Point(166, 90)
        Me.BtnJOG_3.Name = "BtnJOG_3"
        Me.BtnJOG_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.BtnJOG_3.Size = New System.Drawing.Size(73, 73)
        Me.BtnJOG_3.TabIndex = 45
        Me.BtnJOG_3.TabStop = False
        Me.BtnJOG_3.UseVisualStyleBackColor = False
        '
        'BtnJOG_2
        '
        Me.BtnJOG_2.BackColor = System.Drawing.SystemColors.Control
        Me.BtnJOG_2.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnJOG_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnJOG_2.Image = CType(resources.GetObject("BtnJOG_2.Image"), System.Drawing.Image)
        Me.BtnJOG_2.Location = New System.Drawing.Point(19, 90)
        Me.BtnJOG_2.Name = "BtnJOG_2"
        Me.BtnJOG_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.BtnJOG_2.Size = New System.Drawing.Size(73, 73)
        Me.BtnJOG_2.TabIndex = 44
        Me.BtnJOG_2.TabStop = False
        Me.BtnJOG_2.UseVisualStyleBackColor = False
        '
        'BtnJOG_1
        '
        Me.BtnJOG_1.AccessibleRole = System.Windows.Forms.AccessibleRole.None
        Me.BtnJOG_1.BackColor = System.Drawing.SystemColors.Control
        Me.BtnJOG_1.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnJOG_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnJOG_1.Image = CType(resources.GetObject("BtnJOG_1.Image"), System.Drawing.Image)
        Me.BtnJOG_1.Location = New System.Drawing.Point(93, 17)
        Me.BtnJOG_1.Name = "BtnJOG_1"
        Me.BtnJOG_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.BtnJOG_1.Size = New System.Drawing.Size(73, 73)
        Me.BtnJOG_1.TabIndex = 43
        Me.BtnJOG_1.TabStop = False
        Me.BtnJOG_1.UseVisualStyleBackColor = False
        '
        'BtnJOG_0
        '
        Me.BtnJOG_0.BackColor = System.Drawing.SystemColors.Control
        Me.BtnJOG_0.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnJOG_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnJOG_0.Image = CType(resources.GetObject("BtnJOG_0.Image"), System.Drawing.Image)
        Me.BtnJOG_0.Location = New System.Drawing.Point(93, 163)
        Me.BtnJOG_0.Name = "BtnJOG_0"
        Me.BtnJOG_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.BtnJOG_0.Size = New System.Drawing.Size(73, 73)
        Me.BtnJOG_0.TabIndex = 42
        Me.BtnJOG_0.TabStop = False
        Me.BtnJOG_0.UseVisualStyleBackColor = False
        '
        'BtnZ
        '
        Me.BtnZ.BackColor = System.Drawing.SystemColors.Control
        Me.BtnZ.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnZ.Font = New System.Drawing.Font("ＭＳ ゴシック", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnZ.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnZ.Location = New System.Drawing.Point(512, 13)
        Me.BtnZ.Name = "BtnZ"
        Me.BtnZ.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.BtnZ.Size = New System.Drawing.Size(90, 30)
        Me.BtnZ.TabIndex = 41
        Me.BtnZ.TabStop = False
        Me.BtnZ.Text = "Z Off"
        Me.BtnZ.UseVisualStyleBackColor = False
        Me.BtnZ.Visible = False
        '
        'BtnRESET
        '
        Me.BtnRESET.BackColor = System.Drawing.SystemColors.Control
        Me.BtnRESET.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnRESET.Font = New System.Drawing.Font("ＭＳ ゴシック", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnRESET.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnRESET.Location = New System.Drawing.Point(512, 47)
        Me.BtnRESET.Name = "BtnRESET"
        Me.BtnRESET.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.BtnRESET.Size = New System.Drawing.Size(90, 30)
        Me.BtnRESET.TabIndex = 40
        Me.BtnRESET.TabStop = False
        Me.BtnRESET.Text = "RESET"
        Me.BtnRESET.UseVisualStyleBackColor = False
        '
        'BtnSTART
        '
        Me.BtnSTART.AccessibleRole = System.Windows.Forms.AccessibleRole.None
        Me.BtnSTART.BackColor = System.Drawing.SystemColors.Control
        Me.BtnSTART.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnSTART.Font = New System.Drawing.Font("ＭＳ ゴシック", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnSTART.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnSTART.Location = New System.Drawing.Point(323, 47)
        Me.BtnSTART.Name = "BtnSTART"
        Me.BtnSTART.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.BtnSTART.Size = New System.Drawing.Size(90, 30)
        Me.BtnSTART.TabIndex = 39
        Me.BtnSTART.TabStop = False
        Me.BtnSTART.Text = "START"
        Me.BtnSTART.UseVisualStyleBackColor = False
        Me.BtnSTART.Visible = False
        '
        'BtnHALT
        '
        Me.BtnHALT.BackColor = System.Drawing.SystemColors.Control
        Me.BtnHALT.Cursor = System.Windows.Forms.Cursors.Default
        Me.BtnHALT.Font = New System.Drawing.Font("ＭＳ ゴシック", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnHALT.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BtnHALT.Location = New System.Drawing.Point(419, 47)
        Me.BtnHALT.Name = "BtnHALT"
        Me.BtnHALT.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.BtnHALT.Size = New System.Drawing.Size(90, 30)
        Me.BtnHALT.TabIndex = 38
        Me.BtnHALT.TabStop = False
        Me.BtnHALT.Text = "HALT"
        Me.BtnHALT.UseVisualStyleBackColor = False
        Me.BtnHALT.Visible = False
        '
        'GrpPithPanel
        '
        Me.GrpPithPanel.BackColor = System.Drawing.SystemColors.Control
        Me.GrpPithPanel.Controls.Add(Me.TBarPause)
        Me.GrpPithPanel.Controls.Add(Me.TBarHiPitch)
        Me.GrpPithPanel.Controls.Add(Me.TBarLowPitch)
        Me.GrpPithPanel.Controls.Add(Me.LblTchMoval2)
        Me.GrpPithPanel.Controls.Add(Me.LblTchMoval1)
        Me.GrpPithPanel.Controls.Add(Me.LblTchMoval0)
        Me.GrpPithPanel.Controls.Add(Me.LblPitch2)
        Me.GrpPithPanel.Controls.Add(Me.LblPitch1)
        Me.GrpPithPanel.Controls.Add(Me.LblPitch0)
        Me.GrpPithPanel.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.GrpPithPanel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.GrpPithPanel.Location = New System.Drawing.Point(253, 81)
        Me.GrpPithPanel.Name = "GrpPithPanel"
        Me.GrpPithPanel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.GrpPithPanel.Size = New System.Drawing.Size(344, 155)
        Me.GrpPithPanel.TabIndex = 35
        Me.GrpPithPanel.TabStop = False
        Me.GrpPithPanel.Text = "XYZ/BP MOVING PITCH"
        '
        'TBarPause
        '
        Me.TBarPause.AutoSize = False
        Me.TBarPause.Location = New System.Drawing.Point(83, 119)
        Me.TBarPause.Name = "TBarPause"
        Me.TBarPause.Size = New System.Drawing.Size(253, 26)
        Me.TBarPause.TabIndex = 18
        Me.TBarPause.TabStop = False
        '
        'TBarHiPitch
        '
        Me.TBarHiPitch.AllowDrop = True
        Me.TBarHiPitch.AutoSize = False
        Me.TBarHiPitch.Location = New System.Drawing.Point(83, 73)
        Me.TBarHiPitch.Name = "TBarHiPitch"
        Me.TBarHiPitch.Size = New System.Drawing.Size(253, 26)
        Me.TBarHiPitch.TabIndex = 17
        Me.TBarHiPitch.TabStop = False
        '
        'TBarLowPitch
        '
        Me.TBarLowPitch.AutoSize = False
        Me.TBarLowPitch.Location = New System.Drawing.Point(83, 25)
        Me.TBarLowPitch.Name = "TBarLowPitch"
        Me.TBarLowPitch.Size = New System.Drawing.Size(253, 26)
        Me.TBarLowPitch.TabIndex = 16
        Me.TBarLowPitch.TabStop = False
        '
        'LblTchMoval2
        '
        Me.LblTchMoval2.BackColor = System.Drawing.Color.Transparent
        Me.LblTchMoval2.Cursor = System.Windows.Forms.Cursors.Default
        Me.LblTchMoval2.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.LblTchMoval2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblTchMoval2.Location = New System.Drawing.Point(96, 106)
        Me.LblTchMoval2.Name = "LblTchMoval2"
        Me.LblTchMoval2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LblTchMoval2.Size = New System.Drawing.Size(225, 17)
        Me.LblTchMoval2.TabIndex = 15
        Me.LblTchMoval2.Text = "0.0000"
        Me.LblTchMoval2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LblTchMoval1
        '
        Me.LblTchMoval1.AccessibleRole = System.Windows.Forms.AccessibleRole.None
        Me.LblTchMoval1.BackColor = System.Drawing.Color.Transparent
        Me.LblTchMoval1.Cursor = System.Windows.Forms.Cursors.Default
        Me.LblTchMoval1.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.LblTchMoval1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblTchMoval1.Location = New System.Drawing.Point(96, 62)
        Me.LblTchMoval1.Name = "LblTchMoval1"
        Me.LblTchMoval1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LblTchMoval1.Size = New System.Drawing.Size(225, 17)
        Me.LblTchMoval1.TabIndex = 14
        Me.LblTchMoval1.Text = "0.0000"
        Me.LblTchMoval1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LblTchMoval0
        '
        Me.LblTchMoval0.AllowDrop = True
        Me.LblTchMoval0.BackColor = System.Drawing.Color.Transparent
        Me.LblTchMoval0.Cursor = System.Windows.Forms.Cursors.Default
        Me.LblTchMoval0.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.LblTchMoval0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblTchMoval0.Location = New System.Drawing.Point(96, 14)
        Me.LblTchMoval0.Name = "LblTchMoval0"
        Me.LblTchMoval0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LblTchMoval0.Size = New System.Drawing.Size(225, 17)
        Me.LblTchMoval0.TabIndex = 13
        Me.LblTchMoval0.Text = "0.0000"
        Me.LblTchMoval0.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'LblPitch2
        '
        Me.LblPitch2.AccessibleRole = System.Windows.Forms.AccessibleRole.None
        Me.LblPitch2.AutoSize = True
        Me.LblPitch2.BackColor = System.Drawing.SystemColors.Control
        Me.LblPitch2.Cursor = System.Windows.Forms.Cursors.Default
        Me.LblPitch2.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.LblPitch2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblPitch2.Location = New System.Drawing.Point(8, 126)
        Me.LblPitch2.Name = "LblPitch2"
        Me.LblPitch2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LblPitch2.Size = New System.Drawing.Size(65, 12)
        Me.LblPitch2.TabIndex = 12
        Me.LblPitch2.Text = "PAUSE TIME"
        '
        'LblPitch1
        '
        Me.LblPitch1.AutoSize = True
        Me.LblPitch1.BackColor = System.Drawing.SystemColors.Control
        Me.LblPitch1.Cursor = System.Windows.Forms.Cursors.Default
        Me.LblPitch1.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.LblPitch1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblPitch1.Location = New System.Drawing.Point(8, 78)
        Me.LblPitch1.Name = "LblPitch1"
        Me.LblPitch1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LblPitch1.Size = New System.Drawing.Size(65, 12)
        Me.LblPitch1.TabIndex = 11
        Me.LblPitch1.Text = "HIGH PITCH"
        '
        'LblPitch0
        '
        Me.LblPitch0.AccessibleRole = System.Windows.Forms.AccessibleRole.None
        Me.LblPitch0.AutoSize = True
        Me.LblPitch0.BackColor = System.Drawing.SystemColors.Control
        Me.LblPitch0.Cursor = System.Windows.Forms.Cursors.Default
        Me.LblPitch0.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblPitch0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblPitch0.Location = New System.Drawing.Point(8, 30)
        Me.LblPitch0.Name = "LblPitch0"
        Me.LblPitch0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LblPitch0.Size = New System.Drawing.Size(59, 12)
        Me.LblPitch0.TabIndex = 10
        Me.LblPitch0.Text = "LOW PITCH"
        '
        'TmKeyCheck
        '
        Me.TmKeyCheck.Interval = 3
        '
        'BtnTenKey
        '
        Me.BtnTenKey.BackColor = System.Drawing.Color.Pink
        Me.BtnTenKey.Enabled = False
        Me.BtnTenKey.Font = New System.Drawing.Font("ＭＳ ゴシック", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnTenKey.Location = New System.Drawing.Point(450, 184)
        Me.BtnTenKey.Name = "BtnTenKey"
        Me.BtnTenKey.Size = New System.Drawing.Size(143, 37)
        Me.BtnTenKey.TabIndex = 329
        Me.BtnTenKey.Text = "Ten Key On"
        Me.BtnTenKey.UseVisualStyleBackColor = False
        Me.BtnTenKey.Visible = False
        '
        'BtnLaser
        '
        Me.BtnLaser.BackColor = System.Drawing.SystemColors.Control
        Me.BtnLaser.Font = New System.Drawing.Font("ＭＳ ゴシック", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnLaser.Location = New System.Drawing.Point(305, 141)
        Me.BtnLaser.Name = "BtnLaser"
        Me.BtnLaser.Size = New System.Drawing.Size(143, 37)
        Me.BtnLaser.TabIndex = 332
        Me.BtnLaser.Text = "LASER"
        Me.BtnLaser.UseVisualStyleBackColor = False
        '
        'BtnClickEnable
        '
        Me.BtnClickEnable.BackColor = System.Drawing.SystemColors.Control
        Me.BtnClickEnable.Enabled = False
        Me.BtnClickEnable.Font = New System.Drawing.Font("ＭＳ ゴシック", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnClickEnable.Location = New System.Drawing.Point(12, 141)
        Me.BtnClickEnable.Name = "BtnClickEnable"
        Me.BtnClickEnable.Size = New System.Drawing.Size(143, 37)
        Me.BtnClickEnable.TabIndex = 333
        Me.BtnClickEnable.Text = "Click Move"
        Me.BtnClickEnable.UseVisualStyleBackColor = False
        Me.BtnClickEnable.Visible = False
        '
        'BtnLoaderInfo
        '
        Me.BtnLoaderInfo.BackColor = System.Drawing.SystemColors.Control
        Me.BtnLoaderInfo.Enabled = False
        Me.BtnLoaderInfo.Font = New System.Drawing.Font("ＭＳ ゴシック", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnLoaderInfo.Location = New System.Drawing.Point(12, 228)
        Me.BtnLoaderInfo.Name = "BtnLoaderInfo"
        Me.BtnLoaderInfo.Size = New System.Drawing.Size(143, 37)
        Me.BtnLoaderInfo.TabIndex = 334
        Me.BtnLoaderInfo.Text = "ローダ情報"
        Me.BtnLoaderInfo.UseVisualStyleBackColor = False
        Me.BtnLoaderInfo.Visible = False
        '
        'btnAutoInfo
        '
        Me.btnAutoInfo.BackColor = System.Drawing.SystemColors.Control
        Me.btnAutoInfo.Font = New System.Drawing.Font("ＭＳ ゴシック", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnAutoInfo.Location = New System.Drawing.Point(161, 228)
        Me.btnAutoInfo.Name = "btnAutoInfo"
        Me.btnAutoInfo.Size = New System.Drawing.Size(143, 37)
        Me.btnAutoInfo.TabIndex = 335
        Me.btnAutoInfo.Text = "自動運転確認"
        Me.btnAutoInfo.UseVisualStyleBackColor = False
        '
        'frmFineAdjust
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(615, 877)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnAutoInfo)
        Me.Controls.Add(Me.BtnLoaderInfo)
        Me.Controls.Add(Me.BtnClickEnable)
        Me.Controls.Add(Me.BtnLaser)
        Me.Controls.Add(Me.BtnTenKey)
        Me.Controls.Add(Me.GrpArrow)
        Me.Controls.Add(Me.BtnADJ)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.grpBpOff)
        Me.Controls.Add(Me.btnTrimming)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(656, 46)
        Me.Name = "frmFineAdjust"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "ADJFINE"
        Me.TopMost = True
        Me.grpBpOff.ResumeLayout(False)
        Me.grpBpOff.PerformLayout()
        Me.GrpArrow.ResumeLayout(False)
        Me.GrpArrow.PerformLayout()
        Me.GrpPithPanel.ResumeLayout(False)
        Me.GrpPithPanel.PerformLayout()
        CType(Me.TBarPause, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TBarHiPitch, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TBarLowPitch, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents btnTrimming As System.Windows.Forms.Button
    Friend WithEvents grpBpOff As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtBpOffY As System.Windows.Forms.TextBox
    Friend WithEvents txtBpOffX As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents BtnADJ As System.Windows.Forms.Button
    Friend WithEvents GrpArrow As System.Windows.Forms.GroupBox
    Public WithEvents BtnJOG_6 As System.Windows.Forms.Button
    Public WithEvents BtnJOG_5 As System.Windows.Forms.Button
    Public WithEvents BtnJOG_7 As System.Windows.Forms.Button
    Public WithEvents BtnJOG_4 As System.Windows.Forms.Button
    Public WithEvents BtnHI As System.Windows.Forms.Button
    Public WithEvents BtnJOG_3 As System.Windows.Forms.Button
    Public WithEvents BtnJOG_2 As System.Windows.Forms.Button
    Public WithEvents BtnJOG_1 As System.Windows.Forms.Button
    Public WithEvents BtnJOG_0 As System.Windows.Forms.Button
    Public WithEvents BtnZ As System.Windows.Forms.Button
    Public WithEvents BtnRESET As System.Windows.Forms.Button
    Public WithEvents BtnSTART As System.Windows.Forms.Button
    Public WithEvents BtnHALT As System.Windows.Forms.Button
    Public WithEvents GrpPithPanel As System.Windows.Forms.GroupBox
    Friend WithEvents TBarPause As System.Windows.Forms.TrackBar
    Friend WithEvents TBarHiPitch As System.Windows.Forms.TrackBar
    Friend WithEvents TBarLowPitch As System.Windows.Forms.TrackBar
    Public WithEvents LblTchMoval2 As System.Windows.Forms.Label
    Public WithEvents LblTchMoval1 As System.Windows.Forms.Label
    Public WithEvents LblTchMoval0 As System.Windows.Forms.Label
    Public WithEvents LblPitch2 As System.Windows.Forms.Label
    Public WithEvents LblPitch1 As System.Windows.Forms.Label
    Public WithEvents LblPitch0 As System.Windows.Forms.Label
    Friend WithEvents TmKeyCheck As System.Windows.Forms.Timer
    Friend WithEvents BtnTenKey As System.Windows.Forms.Button
    Friend WithEvents BtnLaser As System.Windows.Forms.Button
    Friend WithEvents BtnClickEnable As Button
    Friend WithEvents BtnLoaderInfo As Button
    Friend WithEvents btnAutoInfo As Button
End Class
