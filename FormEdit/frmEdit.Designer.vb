Namespace FormEdit
    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmEdit
#Region "Windows �t�H�[�� �f�U�C�i�ɂ���Đ������ꂽ�R�[�h "
        <System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
            MyBase.New()
            '���̌Ăяo���́AWindows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
            InitializeComponent()
            Form_Initialize_Renamed()
        End Sub

        'Form �́A�R���|�[�l���g�ꗗ�Ɍ㏈�������s���邽�߂� dispose ���I�[�o�[���C�h���܂��B
        <System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
            If Disposing Then
                If Not components Is Nothing Then
                    components.Dispose()
                End If
            End If
            MyBase.Dispose(Disposing)
        End Sub

        'Windows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
        Private components As System.ComponentModel.IContainer
        Public WithEvents MTab As System.Windows.Forms.TabControl
        Public WithEvents CmndCancel As System.Windows.Forms.Button
        Public WithEvents CmndOK As System.Windows.Forms.Button
        Public WithEvents LblToolTip As System.Windows.Forms.Label
        Public WithEvents LblGuid As System.Windows.Forms.Label
        Public WithEvents LblFPATH As System.Windows.Forms.Label
        'Public WithEvents LblT0_2 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
        'Public WithEvents TxtT0_2 As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
        '����: �ȉ��̃v���V�[�W���� Windows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
        'Windows �t�H�[�� �f�U�C�i���g���ĕύX�ł��܂��B
        '�R�[�h �G�f�B�^���g�p���āA�ύX���Ȃ��ł��������B
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
            Me.MTab = New System.Windows.Forms.TabControl()
            Me.CmndCancel = New System.Windows.Forms.Button()
            Me.CmndOK = New System.Windows.Forms.Button()
            Me.LblToolTip = New System.Windows.Forms.Label()
            Me.LblGuid = New System.Windows.Forms.Label()
            Me.LblFPATH = New System.Windows.Forms.Label()
            Me.CmndKey = New System.Windows.Forms.Button()
            Me.SuspendLayout()
            '
            'MTab
            '
            Me.MTab.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.MTab.ItemSize = New System.Drawing.Size(42, 26)
            Me.MTab.Location = New System.Drawing.Point(12, 52)
            Me.MTab.Name = "MTab"
            Me.MTab.SelectedIndex = 1
            Me.MTab.Size = New System.Drawing.Size(1250, 800)
            Me.MTab.TabIndex = 0
            '
            'CmndCancel
            '
            Me.CmndCancel.BackColor = System.Drawing.SystemColors.Control
            Me.CmndCancel.Cursor = System.Windows.Forms.Cursors.Default
            Me.CmndCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.CmndCancel.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.CmndCancel.ForeColor = System.Drawing.SystemColors.ControlText
            Me.CmndCancel.Location = New System.Drawing.Point(1162, 858)
            Me.CmndCancel.Name = "CmndCancel"
            Me.CmndCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
            Me.CmndCancel.Size = New System.Drawing.Size(100, 50)
            Me.CmndCancel.TabIndex = 1
            Me.CmndCancel.Text = "Cancel"
            Me.CmndCancel.UseVisualStyleBackColor = False
            '
            'CmndOK
            '
            Me.CmndOK.BackColor = System.Drawing.SystemColors.Control
            Me.CmndOK.Cursor = System.Windows.Forms.Cursors.Default
            Me.CmndOK.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.CmndOK.ForeColor = System.Drawing.SystemColors.ControlText
            Me.CmndOK.Location = New System.Drawing.Point(1056, 858)
            Me.CmndOK.Name = "CmndOK"
            Me.CmndOK.RightToLeft = System.Windows.Forms.RightToLeft.No
            Me.CmndOK.Size = New System.Drawing.Size(100, 50)
            Me.CmndOK.TabIndex = 0
            Me.CmndOK.Text = "OK"
            Me.CmndOK.UseVisualStyleBackColor = False
            '
            'LblToolTip
            '
            Me.LblToolTip.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
            Me.LblToolTip.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
            Me.LblToolTip.Cursor = System.Windows.Forms.Cursors.Default
            Me.LblToolTip.Font = New System.Drawing.Font("�l�r �o�S�V�b�N", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.LblToolTip.ForeColor = System.Drawing.SystemColors.ControlText
            Me.LblToolTip.Location = New System.Drawing.Point(12, 855)
            Me.LblToolTip.Name = "LblToolTip"
            Me.LblToolTip.RightToLeft = System.Windows.Forms.RightToLeft.No
            Me.LblToolTip.Size = New System.Drawing.Size(1038, 36)
            Me.LblToolTip.TabIndex = 116
            Me.LblToolTip.Text = "LblToolTip"
            '
            'LblGuid
            '
            Me.LblGuid.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
            Me.LblGuid.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
            Me.LblGuid.Cursor = System.Windows.Forms.Cursors.Default
            Me.LblGuid.Font = New System.Drawing.Font("�l�r �S�V�b�N", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.LblGuid.ForeColor = System.Drawing.SystemColors.ControlText
            Me.LblGuid.Location = New System.Drawing.Point(12, 895)
            Me.LblGuid.Name = "LblGuid"
            Me.LblGuid.RightToLeft = System.Windows.Forms.RightToLeft.No
            Me.LblGuid.Size = New System.Drawing.Size(1038, 120)
            Me.LblGuid.TabIndex = 87
            Me.LblGuid.Text = "LblGuid"
            '
            'LblFPATH
            '
            Me.LblFPATH.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
            Me.LblFPATH.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
            Me.LblFPATH.Cursor = System.Windows.Forms.Cursors.Default
            Me.LblFPATH.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.LblFPATH.ForeColor = System.Drawing.SystemColors.ControlText
            Me.LblFPATH.Location = New System.Drawing.Point(12, 9)
            Me.LblFPATH.Name = "LblFPATH"
            Me.LblFPATH.RightToLeft = System.Windows.Forms.RightToLeft.No
            Me.LblFPATH.Size = New System.Drawing.Size(1250, 38)
            Me.LblFPATH.TabIndex = 6
            Me.LblFPATH.Text = "LblFPATH"
            '
            'CmndKey
            '
            Me.CmndKey.Enabled = False
            Me.CmndKey.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
            Me.CmndKey.Location = New System.Drawing.Point(1162, 9)
            Me.CmndKey.Name = "CmndKey"
            Me.CmndKey.Size = New System.Drawing.Size(100, 38)
            Me.CmndKey.TabIndex = 117
            Me.CmndKey.Text = "Keyboard"
            Me.CmndKey.UseVisualStyleBackColor = True
            Me.CmndKey.Visible = False
            '
            'frmEdit
            '
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
            Me.BackColor = System.Drawing.SystemColors.Control
            Me.ClientSize = New System.Drawing.Size(1280, 1024)
            Me.ControlBox = False
            Me.Controls.Add(Me.CmndKey)
            Me.Controls.Add(Me.MTab)
            Me.Controls.Add(Me.CmndCancel)
            Me.Controls.Add(Me.CmndOK)
            Me.Controls.Add(Me.LblToolTip)
            Me.Controls.Add(Me.LblGuid)
            Me.Controls.Add(Me.LblFPATH)
            Me.Cursor = System.Windows.Forms.Cursors.Default
            Me.DoubleBuffered = True
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
            Me.KeyPreview = True
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Name = "frmEdit"
            Me.RightToLeft = System.Windows.Forms.RightToLeft.No
            Me.ShowInTaskbar = False
            Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
            Me.ResumeLayout(False)

        End Sub
        Friend WithEvents CmndKey As System.Windows.Forms.Button
#End Region
    End Class
End Namespace
