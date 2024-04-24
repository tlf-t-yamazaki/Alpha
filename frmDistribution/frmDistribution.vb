'===============================================================================
'   Description  : ���z�}�\������
'
'   Copyright(C) : TOWA LASERFRONT CORP. 2010
'
'===============================================================================
Option Strict Off
Option Explicit On
Friend Class frmDistribution
    Inherits System.Windows.Forms.Form
#Region "�v���C�x�[�g�萔��`"
    '===========================================================================
    '   �萔��`
    '===========================================================================
    ''----- ��ʺ�� �----
    '' ����۰����ڰĂ���
    'Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Integer, ByVal dwExtraInfo As Integer)
    'Private Const VK_SNAPSHOT As Short = &H2CS          ' PrtSc key
    'Private Const VK_LMENU As Short = &HA4S             ' Alt key
    'Private Const KEYEVENTF_KEYUP As Short = &H2S       ' ����UP���
    'Private Const KEYEVENTF_EXTENDEDKEY As Short = &H1S ' ���݂͊g������

    ' ��ʕ\���ʒu�I�t�Z�b�g
    'Private Const DISP_X_OFFSET As Integer = 4                         '###065
    'Private Const DISP_Y_OFFSET As Integer = 20                        '###065
    Private Const DISP_X_OFFSET As Integer = 0                          '###065
    Private Const DISP_Y_OFFSET As Integer = 0                          '###065

#End Region

#Region "�����o�ϐ���`"
    '===========================================================================
    '   �����o�ϐ���`
    '===========================================================================
    Private m_bInitDistForm As Boolean
    Private m_bFgDispGrp As Boolean                                ' �\�����̎��(TRUE:IT FALSE:FT)

    Private giDistributionResNo As Integer                       ' ���z�}�\����R�ԍ�
    Private Const MAX_SCALE_NUM As Integer = 999999999           ' ���̍ő�l
    Private Const MAX_SCALE_RNUM As Integer = 12                 ' ���̕\����R��

    Private dblAverage(MAXRNO) As Double                         ' ���ϒl
    Private dblDeviationIT(MAXRNO) As Double                     ' �W���΍�(IT)
    Private dblDeviationFT(MAXRNO) As Double                     ' �W���΍�(FT)
    Private dblAverageIT(MAXRNO) As Double                       ' IT���ϒl
    Private dblAverageFT(MAXRNO) As Double                       ' FT���ϒl
    Private glRegistNum(MAX_SCALE_RNUM) As Integer               ' ���z�O���t��R��
    Private glRegistNumIT(MAXRNO, MAX_SCALE_RNUM) As Integer     ' ���z�O���t��R�� �Ƽ��ý�
    Private glRegistNumFT(MAXRNO, MAX_SCALE_RNUM) As Integer     ' ���z�O���t��R�� ̧���ý�

    Private dblMinIT(MAXRNO) As Double                           ' �ŏ��l�Ƽ��
    Private dblMaxIT(MAXRNO) As Double                           ' �ő�l�Ƽ��
    Private dblMinFT(MAXRNO) As Double                           ' �ŏ��ļ���
    Private dblMaxFT(MAXRNO) As Double                           ' �ő�ļ���
    Private dblOKRateIT(MAXRNO) As Double                        ' �Ǖi���Ƽ��
    Private dblNGRateIT(MAXRNO) As Double                        ' �s�Ǖi���Ƽ��
    Private dblOKRateFT(MAXRNO) As Double                        ' �Ǖi��̧���
    Private dblNGRateFT(MAXRNO) As Double                        ' �s�Ǖi��̧���

    Private gDistRegNumLblAry(MAX_SCALE_RNUM) As System.Windows.Forms.Label  ' ���z�O���t��R���z��
    Private gDistGrpPerLblAry(MAX_SCALE_RNUM) As System.Windows.Forms.Label  ' ���z�O���t%�z��
    Private gDistShpGrpLblAry(MAX_SCALE_RNUM) As System.Windows.Forms.Label  ' ���z�O���t�z��

    Private gITNx_cnt(MAXRNO) As Integer                         'IT �Z�o�pܰ���
    Private gITNg_cnt(MAXRNO) As Integer                         'IT NG���L�^
    Private gFTNx_cnt(MAXRNO) As Integer                         'FT �Z�o�pܰ���
    Private gFTNg_cnt(MAXRNO) As Integer                         'FT NG���L�^

    Public TotalFT(MAXRNO) As Double                            ' FT ���v
    Public TotalIT(MAXRNO) As Double                            ' IT ���v
    Public TotalSum2FT(MAXRNO) As Double                        ' FT�Q��a 
    Public TotalSum2IT(MAXRNO) As Double                        ' IT�Q��a
#End Region

    ''V2.2.0.0�O 
    '' �W�v�f�[�^�ۑ��p 
    'Structure TOTAL_DATA_MULTI

    '    <VBFixedArray(MAX_RES_USER)> Dim gITNx_cnt() As Integer     ' IT �Z�o�pܰ���
    '    <VBFixedArray(MAX_RES_USER)> Dim gITNg_cnt() As Integer     ' IT NG���L�^
    '    <VBFixedArray(MAX_RES_USER)> Dim gFTNx_cnt() As Integer     ' FT �Z�o�pܰ���
    '    <VBFixedArray(MAX_RES_USER)> Dim gFTNg_cnt() As Integer     ' FT NG���L�^
    '    <VBFixedArray(MAX_RES_USER)> Dim dblAverage() As Double     ' ���ϒl
    '    <VBFixedArray(MAX_RES_USER)> Dim dblDeviationIT() As Double ' �W���΍�(IT)
    '    <VBFixedArray(MAX_RES_USER)> Dim dblDeviationFT() As Double ' �W���΍�(FT)
    '    <VBFixedArray(MAX_RES_USER)> Dim dblAverageIT() As Double   ' IT���ϒl
    '    <VBFixedArray(MAX_RES_USER)> Dim dblAverageFT() As Double   ' FT���ϒl
    '    <VBFixedArray(MAX_RES_USER)> Dim TotalIT() As Double        ' IT ���v
    '    <VBFixedArray(MAX_RES_USER)> Dim TotalFT() As Double        ' FT ���v
    '    <VBFixedArray(MAX_RES_USER)> Dim TotalSum2IT() As Double    ' IT�Q��a
    '    <VBFixedArray(MAX_RES_USER)> Dim TotalSum2FT() As Double    ' FT�Q��a
    '    <VBFixedArray(MAX_RES_USER)> Dim dblMinIT() As Double       ' IT�ŏ��ļ���
    '    <VBFixedArray(MAX_RES_USER)> Dim dblMaxIT() As Double       ' IT�ő�ļ���
    '    <VBFixedArray(MAX_RES_USER)> Dim dblMinFT() As Double       ' FT�ŏ��ļ���
    '    <VBFixedArray(MAX_RES_USER)> Dim dblMaxFT() As Double       ' FT�ő�ļ���
    '    <VBFixedArray(MAX_RES_USER)> Dim TrimCounter() As Double    ' �g���~���O���J�E���^�[
    '    <VBFixedArray(MAX_RES_USER)> Dim Total_TrimCounter() As Double ' �g���~���O���J�E���^�[


    '    Public stCounter1 As RESULT_PARAM                        ' �\���p�f�[�^��`

    '    '���̍\���̂̃C���X�^���X������������ɂ́A"Initialize" ���Ăяo���Ȃ���΂Ȃ�܂���B
    '    Public Sub Initialize()
    '        ReDim gITNx_cnt(MAX_RES_USER)                       ' IT �Z�o�pܰ���
    '        ReDim gITNg_cnt(MAX_RES_USER)                       ' IT NG���L�^
    '        ReDim gFTNx_cnt(MAX_RES_USER)                       ' FT �Z�o�pܰ���
    '        ReDim gFTNg_cnt(MAX_RES_USER)                       ' FT NG���L�^
    '        ReDim dblAverage(MAX_RES_USER)                      ' ���ϒl
    '        ReDim dblDeviationIT(MAX_RES_USER)                  ' �W���΍�(IT)
    '        ReDim dblDeviationFT(MAX_RES_USER)                  ' �W���΍�(FT)
    '        ReDim dblAverageIT(MAX_RES_USER)                    ' IT���ϒl
    '        ReDim dblAverageFT(MAX_RES_USER)                    ' FT���ϒl

    '        ReDim TotalIT(MAX_RES_USER)                         ' IT ���v
    '        ReDim TotalFT(MAX_RES_USER)                         ' FT ���v
    '        ReDim TotalSum2IT(MAX_RES_USER)                     ' IT�Q��a 
    '        ReDim TotalSum2FT(MAX_RES_USER)                     ' FT�Q��a 

    '        ReDim dblMinIT(MAX_RES_USER)                        ' IT�ŏ��ļ���
    '        ReDim dblMaxIT(MAX_RES_USER)                        ' IT�ő�ļ���
    '        ReDim dblMinFT(MAX_RES_USER)                        ' FT�ŏ��ļ���
    '        ReDim dblMaxFT(MAX_RES_USER)                        ' FT�ő�ļ���
    '        ReDim TrimCounter(MAX_RES_USER)                     ' �g���~���O���J�E���^�[
    '        ReDim Total_TrimCounter(MAX_RES_USER)               ' �g���~���O���J�E���^�[�g�[�^��
    '    End Sub

    'End Structure
    '' ������R�l�擾�p�̏W�v�f�[�^�ۑ��p 
    'Public stToTalDataMulti(MAX_RES_USER) As TOTAL_DATA_MULTI

    'V2.2.0.0�O 
#Region "�t�H�[��������"
    '''=========================================================================
    '''<summary>̫�я�����������</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub InitializeForm()
        Dim strMSG As String

        Try
            ' ���z�}�\���p���x���z��̏�����
            gDistRegNumLblAry(0) = Me.LblRegN_00             ' ���z�O���t��R���z��(0�`11)
            gDistRegNumLblAry(1) = Me.LblRegN_01
            gDistRegNumLblAry(2) = Me.LblRegN_02
            gDistRegNumLblAry(3) = Me.LblRegN_03
            gDistRegNumLblAry(4) = Me.LblRegN_04
            gDistRegNumLblAry(5) = Me.LblRegN_05
            gDistRegNumLblAry(6) = Me.LblRegN_06
            gDistRegNumLblAry(7) = Me.LblRegN_07
            gDistRegNumLblAry(8) = Me.LblRegN_08
            gDistRegNumLblAry(9) = Me.LblRegN_09
            gDistRegNumLblAry(10) = Me.LblRegN_10
            gDistRegNumLblAry(11) = Me.LblRegN_11

            gDistGrpPerLblAry(0) = Me.LblGrpPer_00           ' ���z�O���t%�z��(0�`11)
            gDistGrpPerLblAry(1) = Me.LblGrpPer_01
            gDistGrpPerLblAry(2) = Me.LblGrpPer_02
            gDistGrpPerLblAry(3) = Me.LblGrpPer_03
            gDistGrpPerLblAry(4) = Me.LblGrpPer_04
            gDistGrpPerLblAry(5) = Me.LblGrpPer_05
            gDistGrpPerLblAry(6) = Me.LblGrpPer_06
            gDistGrpPerLblAry(7) = Me.LblGrpPer_07
            gDistGrpPerLblAry(8) = Me.LblGrpPer_08
            gDistGrpPerLblAry(9) = Me.LblGrpPer_09
            gDistGrpPerLblAry(10) = Me.LblGrpPer_10
            gDistGrpPerLblAry(11) = Me.LblGrpPer_11

            gDistShpGrpLblAry(0) = Me.LblShpGrp_00                      ' ���z�O���t�z��(0�`11)
            gDistShpGrpLblAry(1) = Me.LblShpGrp_01
            gDistShpGrpLblAry(2) = Me.LblShpGrp_02
            gDistShpGrpLblAry(3) = Me.LblShpGrp_03
            gDistShpGrpLblAry(4) = Me.LblShpGrp_04
            gDistShpGrpLblAry(5) = Me.LblShpGrp_05
            gDistShpGrpLblAry(6) = Me.LblShpGrp_06
            gDistShpGrpLblAry(7) = Me.LblShpGrp_07
            gDistShpGrpLblAry(8) = Me.LblShpGrp_08
            gDistShpGrpLblAry(9) = Me.LblShpGrp_09
            gDistShpGrpLblAry(10) = Me.LblShpGrp_10
            gDistShpGrpLblAry(11) = Me.LblShpGrp_11

            'DistRegItLblAry(i) = New System.Windows.Forms.Label     ' ���z�O���t��R��(IT)�z��
            'DistRegFtLblAry(i) = New System.Windows.Forms.Label     ' ���z�O���t��R��(FT)�z��

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "frmDistribution.InitializeForm() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�C�j�V����/�t�@�C�i�����z�}�̕\�����"
    Public Function DisplayInitialMode() As Boolean
        Return m_bFgDispGrp
    End Function
#End Region

#Region "���z�}�ۑ��{�^������������"
    '''=========================================================================
    '''<summary>���z�}�ۑ��{�^������������</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub cmdGraphSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdGraphSave.Click

        ' �{�^������
        cmdGraphSave.Enabled = False
        cmdInitial.Enabled = False
        cmdFinal.Enabled = False

        ' ��ʂ��n�[�h�R�s�[���������
        Call SaveWindowPic(True, False)

        ' �������b�Z�[�W

        ' �{�^������
        cmdGraphSave.Enabled = True
        cmdInitial.Enabled = True
        cmdFinal.Enabled = True

    End Sub
#End Region

#Region "���z�}�ۑ�����"
    '''=========================================================================
    '''<summary>���z�}�ۑ��{�^������������</summary>
    '''<remarks>PrintScreen�L�[�������Ɠ����̏������s��</remarks>
    '''=========================================================================
    Private Sub SaveWindowPic(Optional ByRef ActWind As Boolean = True, Optional ByRef PrintOn As Boolean = False)

        Dim msg As String               'V4.7.0.0�B

        Try
            If (String.IsNullOrEmpty(IO.Path.GetFileNameWithoutExtension(gsDataFileName))) Then Exit Sub 'V4.7.0.0�B

            Dim fileName As String
            Dim bFileSave As Boolean
            'Dim bitMap As New Bitmap(Me.Width, Me.Height)
            bFileSave = False
            fileName = ""

            ''�A�N�e�B�u��Window���N���b�v�{�[�h�փR�s�[
            'SendKeys.SendWait("%{PRTSC}")

            '' �N���b�v�{�[�h����f�[�^�擾
            'Dim obj As IDataObject = Clipboard.GetDataObject()

            'If obj IsNot Nothing Then
            '    Dim dispImage As Image = DirectCast(obj.GetData(DataFormats.Bitmap), Image)

            '    If dispImage IsNot Nothing Then
            '        If m_bFgDispGrp = True Then
            '            fileName = gSysPrm.stLOG.gsLoggingDir & "IT_MAP" & Now.ToString("yyMMddhhmmss") & ".BMP"
            '        Else
            '            fileName = gSysPrm.stLOG.gsLoggingDir & "FT_MAP" & Now.ToString("yyMMddhhmmss") & ".BMP"
            '        End If

            '        dispImage.Save(fileName)
            '        bFileSave = True
            '    End If
            'End If

            ' �د���ް�ނ�÷��(Bitmap�ȊO�H)����߰����Ă����Ԃ���
            ' dispImage��Nothing�ƂȂ��ĕۑ�����Ȃ����ߕύX              'V4.7.0.0�B
            Dim ITFT As String
            If (True = m_bFgDispGrp) Then
                ITFT = "_IT_MAP"
            Else
                ITFT = "_FT_MAP"
            End If

            fileName = gSysPrm.stLOG.gsLoggingDir & _
                IO.Path.GetFileNameWithoutExtension(IO.Path.GetFileNameWithoutExtension(gsDataFileName)) & _
                ITFT & Now.ToString("yyMMddHHmmss") & ".BMP"

            Using bmp As New Bitmap(Me.Width, Me.Height)
                Me.DrawToBitmap(bmp, New Rectangle(0, 0, Me.Width, Me.Height))
                bmp.Save(fileName)
                bFileSave = True
            End Using

            '���ʂ̕\��
            If (bFileSave = True) Then
                If gSysPrm.stTMN.giMsgTyp = 0 Then
                    'MsgBox("�ۑ������I" & vbCrLf & " (" & fileName & ")")
                    msg = "�ۑ������I" & vbCrLf & " (" & fileName & ")"
                Else
                    'MsgBox("Save completion." & vbCrLf & " (" & fileName & ")")
                    msg = "Save completion." & vbCrLf & " (" & fileName & ")"
                End If
            Else
                If gSysPrm.stTMN.giMsgTyp = 0 Then
                    'MsgBox("�ۑ��ł��܂���ł����B")
                    msg = "�ۑ��ł��܂���ł����B"
                Else
                    'MsgBox("I was not able to save it.")
                    msg = "I was not able to save it."
                End If
            End If

            'Exit Sub

        Catch ex As Exception
            If gSysPrm.stTMN.giMsgTyp = 0 Then
                'MsgBox("�ۑ��ł��܂���ł����B")
                msg = "�ۑ��ł��܂���ł����B" & Environment.NewLine & ex.Message
            Else
                'MsgBox("I was not able to save it.")
                msg = "I was not able to save it." & Environment.NewLine & ex.Message
            End If
        End Try

        ' ���ɉB��Ȃ��悤�ɑΉ�       'V4.7.0.0�B
        MessageBox.Show(msg, cmdGraphSave.Text, MessageBoxButtons.OK, MessageBoxIcon.None, _
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
    End Sub
#End Region

#Region "�t�@�C�i���e�X�g���z�}�\���{�^������������"
    '''=========================================================================
    '''<summary>�t�@�C�i���e�X�g���z�}�\���{�^������������</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub cmdFinal_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdFinal.Click
        m_bFgDispGrp = False
        Call RedrawGraph()                                              ' ���z�}�\������
    End Sub
#End Region

#Region "�C�j�V�����e�X�g���z�}�\���{�^������������"
    '''=========================================================================
    '''<summary>�C�j�V�����e�X�g���z�}�\���{�^������������</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub cmdInitial_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdInitial.Click
        m_bFgDispGrp = True
        Call RedrawGraph()
    End Sub
#End Region

#Region "�t�H�[�����[�h������"
    '''=========================================================================
    '''<summary>�t�H�[�����[�h������</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub frmDistribution_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        'Dim utdClientPoint As tagPOINT
        'Dim lngWin32apiResultCode As Integer
        Dim setLocation As System.Drawing.Point

        '���������s
        If (m_bInitDistForm = False) Then
            InitializeForm()
            m_bInitDistForm = True
        End If

        'bFgfrmDistribution = True                           ' ���Y���̕\���׸�ON

        'Video�̏�ɕ\������B
        setLocation = Form1.VideoLibrary1.Location
        setLocation.X = setLocation.X + DISP_X_OFFSET
        setLocation.Y = setLocation.Y + DISP_Y_OFFSET
        Me.Location = setLocation

        lblRegistTitle.Text = PIC_TRIM_09
        lblGoodTitle.Text = PIC_TRIM_03
        lblNgTitle.Text = PIC_TRIM_04
        lblMinTitle.Text = PIC_TRIM_05
        lblMaxTitle.Text = PIC_TRIM_06
        lblAverage.Text = PIC_TRIM_07
        lblDeviation.Text = PIC_TRIM_08
        cmdInitial.Text = PIC_TRIM_01
        cmdFinal.Text = PIC_TRIM_02

        ' ���z�}�ޯ�ϯ�ߕۑ�
        cmdGraphSave.Visible = True
        cmdGraphSave.Text = PIC_TRIM_10
        RedrawGraph()

        '��ɍőO�ʂɕ\������B
        Me.TopMost = True
    End Sub
#End Region

#Region "�t�H�[�J�X�����������̏���"
    '''=========================================================================
    '''<summary>���M���O�J�n(�W��)</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub frmDistribution_LostFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.LostFocus
        '    Unload Me
    End Sub
#End Region

#Region "�t�H�[���A�����[�h������"
    '''=========================================================================
    '''<summary>�t�H�[���A�����[�h������</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub frmDistribution_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        'bFgfrmDistribution = False                      ' ���Y���̕\���׸�OFF
        Form1.chkDistributeOnOff.Checked = False

        If (gSysPrm.stTMN.giMsgTyp = 0) Then
            Form1.chkDistributeOnOff.Text = "���Y�O���t�@�\��"
        Else
            Form1.chkDistributeOnOff.Text = "Distribute ON"
        End If
    End Sub
#End Region

#Region "���z�}�\������"
    '''=========================================================================
    '''<summary>���z�}�\������</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub RedrawGraph()

        Dim iCnt As Short                                   ' ����
        Dim lMax As Integer
        Dim lScale As Integer
        Dim lScaleMax As Integer
        Dim dblGraphDiv As Double
        Dim dblGraphTop As Double
        Dim dtemp As Double         ' ###203 
        Dim dblTest_LowLimit, dblTest_HighLimit As Double
        Dim dblTemp As Double

        lMax = 0
        If (m_bFgDispGrp) Then

            lblGraphAccumulationTitle.Text = MSG_TRIM_04                ' "�C�j�V�����e�X�g�@���z�}"
            If Double.MaxValue = dblMinIT(giDistributionResNo) Then
                lblMinValue.Text = "0.000"
            Else
                lblMinValue.Text = dblMinIT(giDistributionResNo).ToString("0.000")               ' �ŏ��l
            End If
            If Double.MinValue = dblMaxIT(giDistributionResNo) Then
                lblMaxValue.Text = "0.000"
            Else
                lblMaxValue.Text = dblMaxIT(giDistributionResNo).ToString("0.000")               ' �ő�l
            End If

            For iCnt = 0 To (MAX_SCALE_RNUM - 1)
                glRegistNum(iCnt) = glRegistNumIT(giDistributionResNo, iCnt)                 ' ���z�O���t��R��
                If lMax < glRegistNum(iCnt) Then
                    lMax = glRegistNum(iCnt)
                End If

                gDistRegNumLblAry(iCnt).Text = CStr(glRegistNum(iCnt)) ' ���z�O���t��R��
            Next

            'OK/NG���̕\��
            lblGoodChip.Text = CStr(gITNx_cnt(giDistributionResNo))                      ' OK��
            lblNgChip.Text = CStr(gITNg_cnt(giDistributionResNo))                        ' NG��
            lblOKRate.Text = dblOKRateIT(giDistributionResNo).ToString("0.000")              ' �Ǖi��
            lblNGRate.Text = dblNGRateIT(giDistributionResNo).ToString("0.000")              ' �s�Ǖi��
        Else
            lblGraphAccumulationTitle.Text = MSG_TRIM_05                ' "�t�@�C�i���e�X�g�@���z�}"
            If Double.MaxValue = dblMinFT(giDistributionResNo) Then
                lblMinValue.Text = "0.000"
            Else
                lblMinValue.Text = dblMinFT(giDistributionResNo).ToString("0.000")               ' �ŏ��l
            End If
            If Double.MinValue = dblMaxFT(giDistributionResNo) Then
                lblMaxValue.Text = "0.000"
            Else
                lblMaxValue.Text = dblMaxFT(giDistributionResNo).ToString("0.000")               ' �ő�l
            End If

            For iCnt = 0 To (MAX_SCALE_RNUM - 1)

                glRegistNum(iCnt) = glRegistNumFT(giDistributionResNo, iCnt)

                If lMax < glRegistNum(iCnt) Then
                    lMax = glRegistNum(iCnt)
                End If
                gDistRegNumLblAry(iCnt).Text = CStr(glRegistNum(iCnt)) ' ���z�O���t��R��
            Next
            'OK/NG���̕\��
            lblGoodChip.Text = CStr(gFTNx_cnt(giDistributionResNo))                      ' OK��
            lblNgChip.Text = CStr(gFTNg_cnt(giDistributionResNo))                        ' NG��
            lblOKRate.Text = dblOKRateFT(giDistributionResNo).ToString("0.000")              ' �Ǖi��
            lblNGRate.Text = dblNGRateFT(giDistributionResNo).ToString("0.000")              ' �s�Ǖi��
        End If

        'lblGoodChip.Text = CStr(lOkChip)                               ' OK��
        'lblNgChip.Text = CStr(lNgChip)                                 ' NG��


        '������������
        ' �덷�ް�������(IT)
        '' '' ''Call Form1.GetMoveMode(digL, digH, digSW)
        If gITNx_cnt(giDistributionResNo) >= 0 Then
            'If (gDigL = 0) Then                                        ' x0���[�h ?
            '' '' ''If (digL = 0) Then                                  ' x0���[�h ?
            '###154 �v�Z�͌��ʎ擾���ɂ��̓s�x���s����
            '' ���ϒl�擾
            'dblAverageIT = Form1.Utility1.GetAverage(gITNx, gITNx_cnt + 1)
            '' �W���΍��̎擾
            'dblDeviationIT = Form1.Utility1.GetDeviation(gITNx, gITNx_cnt + 1, dblAverageIT)
            'TotalDeviationDebug = TotalDeviationDebug '###154
            'TotalAverageDebug = TotalAverageDebug '###154
            '' '' ''End If
        End If

        ' �덷�ް�������(FT)
        If gFTNx_cnt(giDistributionResNo) >= 0 Then
            '###154            ' ���ϒl�擾
            '###154            dblAverageFT = Form1.Utility1.GetAverage(gFTNx, gFTNx_cnt + 1)
            '###154     ' �W���΍��̎擾
            '###154         dblDeviationFT = Form1.Utility1.GetDeviation(gFTNx, gFTNx_cnt + 1, dblAverageFT)
            'dblAverageFT = TotalAverageDebug '###154
            'dblDeviationFT = TotalDeviationDebug '###154
        End If
        '��������������

        If (m_bFgDispGrp) Then
            lblDeviationValue.Text = dblDeviationIT(giDistributionResNo).ToString("0.000000") ' �W���΍�(IT)
        Else
            lblDeviationValue.Text = dblDeviationFT(giDistributionResNo).ToString("0.000000") ' �W���΍�(FT)
        End If

        If (m_bFgDispGrp) Then
            dblAverage(giDistributionResNo) = dblAverageIT(giDistributionResNo)
        Else
            dblAverage(giDistributionResNo) = dblAverageFT(giDistributionResNo)
        End If
        lblAverageValue.Text = dblAverage(giDistributionResNo).ToString("0.000")     ' ���ϒl

        lScaleMax = 0                                           ' �I�[�g�X�P�[�����O
        lScale = 100
        Do
            If (lScale > lMax) Then                             ' lScale < ��R�� ?
                lScaleMax = lScale
            ElseIf ((lScale * 2) > lMax) Then
                lScaleMax = (lScale * 2)
            ElseIf ((lScale * 5) > lMax) Then
                lScaleMax = (lScale * 5)
            End If
            lScale = lScale * 10
        Loop While (0 = lScaleMax) And (MAX_SCALE_NUM > lScale)

        If (0 = lScaleMax) Then
            lScaleMax = MAX_SCALE_NUM + 1
        End If


        If (m_bFgDispGrp) Then
            If giDistributionResNo = 0 Then
                For i As Integer = 1 To stPLT.RCount
                    If Not UserModule.IsMarking(i) Then
                        dblTemp = stREG(i).dblITL
                        If stREG(i).intMode = JUDGE_MODE_ABSOLUTE Then   ' ��Βl�̎��䗦�ɕϊ�
                            dblTemp = dblTemp / stREG(i).dblNOM * 100.0
                        End If
                        If dblTest_LowLimit > dblTemp Then
                            dblTest_LowLimit = dblTemp
                        End If
                        dblTemp = stREG(i).dblITH
                        If stREG(i).intMode = JUDGE_MODE_ABSOLUTE Then   ' ��Βl�̎��䗦�ɕϊ�
                            dblTemp = dblTemp / stREG(i).dblNOM * 100.0
                        End If
                        If dblTest_HighLimit < dblTemp Then
                            dblTest_HighLimit = dblTemp
                        End If
                    End If
                Next
            Else
                dblTest_LowLimit = stREG(giDistributionResNo).dblITL
                dblTest_HighLimit = stREG(giDistributionResNo).dblITH
            End If
        Else
            If giDistributionResNo = 0 Then
                For i As Integer = 1 To stPLT.RCount
                    If Not UserModule.IsMarking(i) Then
                        dblTemp = stREG(i).dblFTL
                        If stREG(i).intMode = JUDGE_MODE_ABSOLUTE Then   ' ��Βl�̎��䗦�ɕϊ�
                            dblTemp = dblTemp / stREG(i).dblNOM * 100.0
                        End If
                        If dblTest_LowLimit > dblTemp Then
                            dblTest_LowLimit = dblTemp
                        End If
                        dblTemp = stREG(i).dblFTH
                        If stREG(i).intMode = JUDGE_MODE_ABSOLUTE Then   ' ��Βl�̎��䗦�ɕϊ�
                            dblTemp = dblTemp / stREG(i).dblNOM * 100.0
                        End If
                        If dblTest_HighLimit < dblTemp Then
                            dblTest_HighLimit = dblTemp
                        End If
                    End If
                Next
            Else
                dblTest_LowLimit = stREG(giDistributionResNo).dblFTL
                dblTest_HighLimit = stREG(giDistributionResNo).dblFTH
            End If
        End If

        If giDistributionResNo > 0 And stREG(giDistributionResNo).intMode = JUDGE_MODE_ABSOLUTE Then   ' ��Βl�̎��䗦�ɕϊ�
            dblTest_LowLimit = dblTest_LowLimit / stREG(giDistributionResNo).dblNOM * 100.0
            dblTest_HighLimit = dblTest_HighLimit / stREG(giDistributionResNo).dblNOM * 100.0
        End If

        If ((0 >= dblTest_LowLimit) And (0 <= dblTest_HighLimit)) Then
            dblGraphDiv = (dblTest_HighLimit * 1.5 - dblTest_LowLimit * 1.5) / 10
            dblGraphTop = dblTest_HighLimit * 1.5
        ElseIf ((0 >= dblTest_LowLimit) And (0 > dblTest_HighLimit)) Then
            dblGraphDiv = (dblTest_HighLimit / 1.5 - dblTest_LowLimit * 1.5) / 10
            dblGraphTop = dblTest_HighLimit * 1.5
        ElseIf ((0 < dblTest_LowLimit) And (0 <= dblTest_HighLimit)) Then
            dblGraphDiv = (dblTest_HighLimit * 1.5 - dblTest_LowLimit / 1.5) / 10
            dblGraphTop = dblTest_HighLimit * 1.5
        Else
            dblGraphDiv = 0.3
            dblGraphTop = 1.5
        End If

        gDistGrpPerLblAry(0).Text = "�`" & dblGraphTop.ToString("0.00")
        For iCnt = 1 To 11
            'gDistGrpPerLblAry(iCnt).Text = (dblGraphTop - (dblGraphDiv * (iCnt - 1)).ToString("0.00")) & "�`"
            ' ###203 
            dtemp = (dblGraphTop - (dblGraphDiv * (iCnt - 1)))
            If ((-0.001 < dtemp) And (dtemp < 0.001)) Then
                gDistGrpPerLblAry(iCnt).Text = "0�`"
            Else
                gDistGrpPerLblAry(iCnt).Text = (dtemp.ToString("0.00")) & "�`"
            End If
            ' ###203
        Next

        picGraphAccumulationDrawLine(lScaleMax)
        Call picGraphAccumulationPrintRegistNum()           ' ���z�O���t�ɒ�R����ݒ肷��

    End Sub
#End Region

#Region "���z�}�\���T�u"
    '''=========================================================================
    ''' <summary>
    ''' ���z�}�\���T�u
    ''' </summary>
    ''' <param name="lScaleMax">(INP)�X�P�[��</param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub picGraphAccumulationDrawLine(ByRef lScaleMax As Integer)

        Dim i As Short
        Dim x As Short

        For i = 0 To (MAX_SCALE_RNUM - 1)
            '            x = CShort((glRegistNum(i) * 473) \ lScaleMax)   ' ���z�O���t��R��
            x = CShort((glRegistNum(i) * 250) \ lScaleMax)   ' ���z�O���t��R��
            'If (473 < x) Then
            If (250 < x) Then
                '                x = 473
                x = 250
            End If
            gDistShpGrpLblAry(i).Width = x
        Next
        lblRegistUnit.Text = CStr(lScaleMax \ 2)            ' ��R���̔����̐� 

    End Sub
#End Region

#Region "���z�O���t�ɒ�R����ݒ肷��"
    '''=========================================================================
    '''<summary>���z�O���t�ɒ�R����ݒ肷��</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub picGraphAccumulationPrintRegistNum()

        Dim i As Short

        For i = 0 To (MAX_SCALE_RNUM - 1)
            gDistRegNumLblAry(i).Text = CStr(glRegistNum(i))  ' ���z�O���t��R��
        Next

    End Sub
#End Region

#Region "���z�}�f�[�^�N���A"
    Public Sub ClearCounter()
        Try

            For i As Integer = 0 To MAXRNO

                dblAverage(i) = 0.0                     ' ���ϒl
                dblDeviationIT(i) = 0.0                 ' �W���΍�(IT)
                dblDeviationFT(i) = 0.0                 ' �W���΍�(FT)
                dblAverageIT(i) = 0.0                   ' IT���ϒl
                dblAverageFT(i) = 0.0                   ' FT���ϒl

                For j As Integer = 0 To MAX_SCALE_RNUM
                    glRegistNum(j) = 0               ' ���z�O���t��R��
                    glRegistNumIT(i, j) = 0             ' ���z�O���t��R�� �Ƽ��ý�
                    glRegistNumFT(i, j) = 0             ' ���z�O���t��R�� ̧���ý�
                Next

                dblMinIT(i) = Double.MaxValue           ' �ŏ��l�Ƽ��
                dblMaxIT(i) = Double.MinValue           ' �ő�l�Ƽ��
                dblMinFT(i) = Double.MaxValue           ' �ŏ��ļ���
                dblMaxFT(i) = Double.MinValue           ' �ő�ļ���

                dblOKRateIT(i) = 0.0                    ' �Ǖi���Ƽ��
                dblNGRateIT(i) = 0.0                    ' �s�Ǖi���Ƽ��
                dblOKRateFT(i) = 0.0                    ' �Ǖi��̧���
                dblNGRateFT(i) = 0.0                    ' �s�Ǖi��̧���

                gITNx_cnt(i) = 0                        'IT �Z�o�pܰ���
                gITNg_cnt(i) = 0                        'IT NG���L�^
                gFTNx_cnt(i) = 0                        'FT �Z�o�pܰ���
                gFTNg_cnt(i) = 0                        'FT NG���L�^

                TotalFT(i) = 0.0                        ' FT ���v      
                TotalIT(i) = 0.0                        ' IT ���v
                TotalSum2FT(i) = 0.0                    ' FT�Q��a      
                TotalSum2IT(i) = 0.0                    ' IT�Q��a     
            Next


            'V2.2.0.0�O��
            For MultiCnt As Integer = 0 To MAX_RES_USER
                With stToTalDataMulti(MultiCnt)
                    .Initialize()

                    For rn As Integer = 0 To MAX_RES_USER
                        .gITNx_cnt(rn) = 0                      ' IT �Z�o�pܰ���
                        .gITNg_cnt(rn) = 0                      ' IT NG���L�^
                        .gFTNx_cnt(rn) = 0                      ' FT �Z�o�pܰ���
                        .gFTNg_cnt(rn) = 0                      ' FT NG���L�^
                        .dblAverage(rn) = 0                     ' ���ϒl
                        .dblDeviationIT(rn) = 0                 ' �W���΍�(IT)
                        .dblDeviationFT(rn) = 0                 ' �W���΍�(FT)
                        .dblAverageIT(rn) = 0                   ' IT���ϒl
                        .dblAverageFT(rn) = 0                   ' FT���ϒl
                        .TotalIT(rn) = 0.0                      ' IT ���v      
                        .TotalFT(rn) = 0.0                      ' FT ���v      
                        .TotalSum2IT(rn) = 0                    ' IT�Q��a
                        .TotalSum2FT(rn) = 0                    ' FT�Q��a
                        .dblMinIT(rn) = Double.MaxValue         ' IT�ŏ��ļ���
                        .dblMaxIT(rn) = Double.MinValue         ' IT�ő�ļ���
                        .dblMinFT(rn) = Double.MaxValue         ' FT�ŏ��ļ���
                        .dblMaxFT(rn) = Double.MinValue         ' FT�ő�ļ���

                    Next rn
                End With
            Next MultiCnt
            'V2.2.0.0�O��

        Catch ex As Exception
            Call Z_PRINT("frmDistribution.ClearCounter() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub
#End Region

#Region "���z�}�f�[�^�ݒ�"
    Public Sub SetGraphData(ByVal dTop As Double, ByVal dDiv As Double, ByVal dGap As Double, ByVal rn As Integer, ByRef iRegistNum(,) As Integer)
        Try
            If ((dTop - (dDiv * 0)) < dGap) Then
                ' ���̍ŏ�i�l-(���͈͍̔��݈ʒu*0)�@<�@���̏ꍇ
                iRegistNum(rn, 0) = iRegistNum(rn, 0) + 1
            ElseIf ((dTop - (dDiv * 1)) < dGap) Then
                ' ���̍ŏ�i�l-(���͈͍̔��݈ʒu*1)�@<�@���̏ꍇ
                iRegistNum(rn, 1) = iRegistNum(rn, 1) + 1
            ElseIf ((dTop - (dDiv * 2)) < dGap) Then
                ' ���̍ŏ�i�l-(���͈͍̔��݈ʒu*2)�@<�@���̏ꍇ
                iRegistNum(rn, 2) = iRegistNum(rn, 2) + 1
            ElseIf ((dTop - (dDiv * 3)) < dGap) Then
                ' ���̍ŏ�i�l-(���͈͍̔��݈ʒu*3)�@<�@���̏ꍇ
                iRegistNum(rn, 3) = iRegistNum(rn, 3) + 1
            ElseIf ((dTop - (dDiv * 4)) < dGap) Then
                ' ���̍ŏ�i�l-(���͈͍̔��݈ʒu*4)�@<�@���̏ꍇ
                iRegistNum(rn, 4) = iRegistNum(rn, 4) + 1
            ElseIf ((dTop - (dDiv * 5)) < dGap) Then
                ' ���̍ŏ�i�l-(���͈͍̔��݈ʒu*5)�@<�@���̏ꍇ
                iRegistNum(rn, 5) = iRegistNum(rn, 5) + 1
            ElseIf ((dTop - (dDiv * 6)) < dGap) Then
                ' ���̍ŏ�i�l-(���͈͍̔��݈ʒu*6)�@<�@���̏ꍇ
                iRegistNum(rn, 6) = iRegistNum(rn, 6) + 1
            ElseIf ((dTop - (dDiv * 7)) < dGap) Then
                ' ���̍ŏ�i�l-(���͈͍̔��݈ʒu*7)�@<�@���̏ꍇ
                iRegistNum(rn, 7) = iRegistNum(rn, 7) + 1
            ElseIf ((dTop - (dDiv * 8)) < dGap) Then
                ' ���̍ŏ�i�l-(���͈͍̔��݈ʒu*8)�@<�@���̏ꍇ
                iRegistNum(rn, 8) = iRegistNum(rn, 8) + 1
            ElseIf ((dTop - (dDiv * 9)) < dGap) Then
                ' ���̍ŏ�i�l-(���͈͍̔��݈ʒu*9)�@<�@���̏ꍇ
                iRegistNum(rn, 9) = iRegistNum(rn, 9) + 1
            ElseIf ((dTop - (dDiv * 10)) < dGap) Then
                ' ���̍ŏ�i�l-(���͈͍̔��݈ʒu*10)�@<�@���̏ꍇ
                iRegistNum(rn, 10) = iRegistNum(rn, 10) + 1
            Else
                ' ��L�����ȊO�̏ꍇ
                iRegistNum(rn, 11) = iRegistNum(rn, 11) + 1
            End If
        Catch ex As Exception
            Call Z_PRINT("frmDistribution.SetGraphData() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub

    ''' <summary>
    ''' �W���΍��̌v�Z�Ɗe��f�[�^�̕ۑ�
    ''' </summary>
    ''' <param name="rn"></param>
    ''' <param name="dblGap"></param>
    ''' <param name="Sum"></param>
    ''' <param name="Total"></param>
    ''' <param name="TotalSum2"></param>
    ''' <param name="dblDeviation"></param>
    ''' <param name="dblMin"></param>
    ''' <param name="dblMax"></param>
    ''' <remarks></remarks>
    Private Sub SetDeviation(ByRef rn As Integer, ByVal dblGap As Double, ByRef Sum() As Integer, ByRef Total() As Double, ByRef TotalSum2() As Double, ByRef dblDeviation() As Double, ByRef dblMin() As Double, ByRef dblMax() As Double, ByRef Average() As Double)
        Try
            Sum(rn) = Sum(rn) + 1                                                           ' �f�[�^���J�E���^�[�P�J�E���g�A�b�v
            Total(rn) = Total(rn) + dblGap                                                  ' �f�[�^�̍��v�i�a�j
            Average(rn) = Total(rn) / Sum(rn)                                                   ' ���ϒl

            TotalSum2(rn) = TotalSum2(rn) + (dblGap * dblGap)                               ' �Q��a
            dblDeviation(rn) = Math.Sqrt((TotalSum2(rn) / Sum(rn)) - (Average(rn) * Average(rn)))   ' �W���΍�

            'V2.2.0.031��
            ' ���l�ɂȂ��Ă��Ȃ��ꍇ�͂O�Ƃ���
            If Double.IsNaN(dblDeviation(rn)) Then
                dblDeviation(rn) = 0.0
            End If
            'V2.2.0.031��

            '(�W���΍��Z�o���C��)
            If (dblMin(rn) > dblGap) Then                                                   ' �ŏ�
                dblMin(rn) = dblGap
            End If
            If (dblMax(rn) < dblGap) Then                                                   ' �ő�
                dblMax(rn) = dblGap
            End If

        Catch ex As Exception
            Call Z_PRINT("frmDistribution.SetDeviation() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Sub

    '''=========================================================================
    ''' <summary>
    ''' �t�@�C�i���e�X�g�e�X�g���z�}
    ''' </summary>
    ''' <param name="JudgeMode">IT or FT</param>
    ''' <param name="rn">��R�ԍ�</param>
    ''' <param name="MeasureValue">����l</param>
    ''' <param name="Judge">����</param>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub TrimLoggingGraph(ByVal JudgeMode As Integer, ByVal rn As Integer, ByVal MeasureValue As Double, ByVal Judge As Integer)


        Try

            Dim dblGraphDiv As Double                                       ' �O���t�͈͍��ݒl
            Dim dblGraphTop As Double                                       ' �O���t�ŏ�i�l
            Dim dblGraphDivAll As Double                                       ' �O���t�͈͍��ݒl
            Dim dblGraphTopAll As Double                                       ' �O���t�ŏ�i�l
            Dim dblGap As Double
            Dim dblTemp As Double

            Dim dblTest_LowLimit, dblTest_HighLimit As Double
            Dim dblTest_LowLimitAll As Double = Double.MaxValue, dblTest_HighLimitAll As Double = Double.MinValue

            If JudgeMode = INITIAL_TEST Then
                dblTest_LowLimit = stREG(rn).dblITL
                dblTest_HighLimit = stREG(rn).dblITH
                For i As Integer = 1 To stPLT.RCount
                    If Not UserModule.IsMarking(rn) Then
                        dblTemp = stREG(rn).dblITL
                        If stREG(rn).intMode = JUDGE_MODE_ABSOLUTE Then   ' ��Βl�̎��䗦�ɕϊ�
                            dblTemp = dblTemp / stREG(rn).dblNOM * 100.0
                        End If
                        If dblTest_LowLimitAll > dblTemp Then
                            dblTest_LowLimitAll = dblTemp
                        End If
                        dblTemp = stREG(rn).dblITH
                        If stREG(rn).intMode = JUDGE_MODE_ABSOLUTE Then   ' ��Βl�̎��䗦�ɕϊ�
                            dblTemp = dblTemp / stREG(rn).dblNOM * 100.0
                        End If
                        If dblTest_HighLimitAll < dblTemp Then
                            dblTest_HighLimitAll = dblTemp
                        End If
                    End If
                Next
            ElseIf JudgeMode = FINAL_TEST Then
                dblTest_LowLimit = stREG(rn).dblFTL
                dblTest_HighLimit = stREG(rn).dblFTH
                For i As Integer = 1 To stPLT.RCount
                    If Not UserModule.IsMarking(rn) Then
                        dblTemp = stREG(rn).dblFTL
                        If stREG(rn).intMode = JUDGE_MODE_ABSOLUTE Then   ' ��Βl�̎��䗦�ɕϊ�
                            dblTemp = dblTemp / stREG(rn).dblNOM * 100.0
                        End If
                        If dblTest_LowLimitAll > dblTemp Then
                            dblTest_LowLimitAll = dblTemp
                        End If
                        dblTemp = stREG(rn).dblFTH
                        If stREG(rn).intMode = JUDGE_MODE_ABSOLUTE Then   ' ��Βl�̎��䗦�ɕϊ�
                            dblTemp = dblTemp / stREG(rn).dblNOM * 100.0
                        End If
                        If dblTest_HighLimitAll < dblTemp Then
                            dblTest_HighLimitAll = dblTemp
                        End If
                    End If
                Next
            End If

            If stREG(rn).intMode = JUDGE_MODE_ABSOLUTE Then   ' ��Βl�̎��䗦�ɕϊ�
                dblTest_LowLimit = dblTest_LowLimit / stREG(rn).dblNOM * 100.0
                dblTest_HighLimit = dblTest_HighLimit / stREG(rn).dblNOM * 100.0
            End If

            ' ���݂̒�R�̌v�Z
            ' ̧���ý�(LOW�Я�)��̧���ý�(HIGH�Я�)�̒l����������B
            If ((0 >= dblTest_LowLimit) And (0 <= dblTest_HighLimit)) Then
                ' ̧���ý�(LOW�Я�)��0�ȉ���̧���ý�(HIGH�Я�)��0�ȏ�̏ꍇ
                dblGraphDiv = (dblTest_HighLimit * 1.5 - dblTest_LowLimit * 1.5) / 10
                dblGraphTop = dblTest_HighLimit * 1.5

            ElseIf ((0 >= dblTest_LowLimit) And (0 > dblTest_HighLimit)) Then
                ' ̧���ý�(LOW�Я�)��0�ȉ���̧���ý�(HIGH�Я�)����菬�����ꍇ
                dblGraphDiv = (dblTest_HighLimit / 1.5 - dblTest_LowLimit * 1.5) / 10
                dblGraphTop = dblTest_HighLimit * 1.5

            ElseIf ((0 < dblTest_LowLimit) And (0 <= dblTest_HighLimit)) Then
                ' ̧���ý�(LOW�Я�)��0���傫����̧���ý�(HIGH�Я�)��0�ȏ�̏ꍇ
                dblGraphDiv = (dblTest_HighLimit * 1.5 - dblTest_LowLimit / 1.5) / 10
                dblGraphTop = dblTest_HighLimit * 1.5
            Else
                ' ��L�����ȊO�̏ꍇ
                dblGraphDiv = 0.3
                dblGraphTop = 1.5
            End If

            ' �S��R�̌v�Z
            ' ̧���ý�(LOW�Я�)��̧���ý�(HIGH�Я�)�̒l����������B
            If ((0 >= dblTest_LowLimitAll) And (0 <= dblTest_HighLimitAll)) Then
                ' ̧���ý�(LOW�Я�)��0�ȉ���̧���ý�(HIGH�Я�)��0�ȏ�̏ꍇ
                dblGraphDivAll = (dblTest_HighLimitAll * 1.5 - dblTest_LowLimitAll * 1.5) / 10
                dblGraphTopAll = dblTest_HighLimitAll * 1.5

            ElseIf ((0 >= dblTest_LowLimitAll) And (0 > dblTest_HighLimitAll)) Then
                ' ̧���ý�(LOW�Я�)��0�ȉ���̧���ý�(HIGH�Я�)����菬�����ꍇ
                dblGraphDivAll = (dblTest_HighLimitAll / 1.5 - dblTest_LowLimitAll * 1.5) / 10
                dblGraphTopAll = dblTest_HighLimitAll * 1.5

            ElseIf ((0 < dblTest_LowLimitAll) And (0 <= dblTest_HighLimitAll)) Then
                ' ̧���ý�(LOW�Я�)��0���傫����̧���ý�(HIGH�Я�)��0�ȏ�̏ꍇ
                dblGraphDivAll = (dblTest_HighLimitAll * 1.5 - dblTest_LowLimitAll / 1.5) / 10
                dblGraphTopAll = dblTest_HighLimitAll * 1.5
            Else
                ' ��L�����ȊO�̏ꍇ
                dblGraphDivAll = 0.3
                dblGraphTopAll = 1.5
            End If

            ' �����Z�o����B�@̧���ýČ���/���ݸޖڕW�l*100�@-�@100
            dblGap = (MeasureValue / stREG(rn).dblNOM) * 100.0# - 100.0#

            If JudgeMode = INITIAL_TEST Then
                SetGraphData(dblGraphTop, dblGraphDiv, dblGap, rn, glRegistNumIT)
                SetGraphData(dblGraphTopAll, dblGraphDivAll, dblGap, 0, glRegistNumIT)
            ElseIf JudgeMode = FINAL_TEST Then
                SetGraphData(dblGraphTop, dblGraphDiv, dblGap, rn, glRegistNumFT)
                SetGraphData(dblGraphTopAll, dblGraphDivAll, dblGap, 0, glRegistNumFT)
            End If


            If Judge = eJudge.JG_OK Then
                If JudgeMode = INITIAL_TEST Then
                    Call SetDeviation(rn, dblGap, gITNx_cnt, TotalIT, TotalSum2IT, dblDeviationIT, dblMinIT, dblMaxIT, dblAverageIT)  ' ����R
                    Call SetDeviation(0, dblGap, gITNx_cnt, TotalIT, TotalSum2IT, dblDeviationIT, dblMinIT, dblMaxIT, dblAverageIT)   ' �S��R
                ElseIf JudgeMode = FINAL_TEST Then
                    Call SetDeviation(rn, dblGap, gFTNx_cnt, TotalFT, TotalSum2FT, dblDeviationFT, dblMinFT, dblMaxFT, dblAverageFT)  ' ����R
                    Call SetDeviation(0, dblGap, gFTNx_cnt, TotalFT, TotalSum2FT, dblDeviationFT, dblMinFT, dblMaxFT, dblAverageFT)   ' �S��R
                End If
            Else
                'NG�J�E���g�����L�^
                If JudgeMode = INITIAL_TEST Then
                    gITNg_cnt(rn) = gITNg_cnt(rn) + 1
                    gITNg_cnt(0) = gITNg_cnt(0) + 1
                ElseIf JudgeMode = FINAL_TEST Then
                    gFTNg_cnt(rn) = gFTNg_cnt(rn) + 1
                    gFTNg_cnt(0) = gFTNg_cnt(0) + 1
                End If
            End If

            ' �Ǖi���A�s�Ǖi��
            If JudgeMode = INITIAL_TEST Then
                dblOKRateIT(rn) = gITNx_cnt(rn) / (gITNx_cnt(rn) + gITNg_cnt(rn)) * 100.0
                dblNGRateIT(rn) = gITNg_cnt(rn) / (gITNx_cnt(rn) + gITNg_cnt(rn)) * 100.0
                dblOKRateIT(0) = gITNx_cnt(0) / (gITNx_cnt(0) + gITNg_cnt(0)) * 100.0
                dblNGRateIT(0) = gITNg_cnt(0) / (gITNx_cnt(0) + gITNg_cnt(0)) * 100.0
            ElseIf JudgeMode = FINAL_TEST Then
                dblOKRateFT(rn) = gFTNx_cnt(rn) / (gFTNx_cnt(rn) + gFTNg_cnt(rn)) * 100.0
                dblNGRateFT(rn) = gFTNg_cnt(rn) / (gFTNx_cnt(rn) + gFTNg_cnt(rn)) * 100.0
                dblOKRateFT(0) = gFTNx_cnt(0) / (gFTNx_cnt(0) + gFTNg_cnt(0)) * 100.0
                dblNGRateFT(0) = gFTNg_cnt(0) / (gFTNx_cnt(0) + gFTNg_cnt(0)) * 100.0
            End If

        Catch ex As Exception
            Call Z_PRINT("frmDistribution.TrimLoggingGraph() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub
#End Region

#Region "�\����R�ԍ��̐ݒ�"
    Public Sub SetDistributionResNo(ByRef No As Integer)
        Try
            giDistributionResNo = No
        Catch ex As Exception
            Call Z_PRINT("SetDistributionResNo.SetDeviation() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub
#End Region

#Region "���v�f�[�^�̂ݏ���"
    ''' <summary>
    ''' ���v�f�[�^�̕ۑ�
    ''' </summary>
    ''' <param name="JudgeMode">IT or FT</param>
    ''' <param name="rn">��R�ԍ�</param>
    ''' <param name="MeasureValue">����l</param>
    ''' <param name="Judge">����</param>
    ''' <remarks></remarks>
    Public Sub StatisticalDataSave(ByVal JudgeMode As Integer, ByVal rn As Integer, ByVal MeasureValue As Double, ByVal Judge As Integer)

        Try

            Dim dblGap As Double

            ' �����Z�o����B�@̧���ýČ���/���ݸޖڕW�l*100�@-�@100
            '�΍��łȂ��^�l��o�^����B            dblGap = (MeasureValue / stREG(rn).dblNOM) * 100.0# - 100.0#
            dblGap = MeasureValue

            If Judge = eJudge.JG_OK Then
                If JudgeMode = INITIAL_TEST Then
                    Call SetDeviation(rn, dblGap, gITNx_cnt, TotalIT, TotalSum2IT, dblDeviationIT, dblMinIT, dblMaxIT, dblAverageIT)  ' ����R
                    Call SetDeviation(0, dblGap, gITNx_cnt, TotalIT, TotalSum2IT, dblDeviationIT, dblMinIT, dblMaxIT, dblAverageIT)   ' �S��R
                ElseIf JudgeMode = FINAL_TEST Then
                    Call SetDeviation(rn, dblGap, gFTNx_cnt, TotalFT, TotalSum2FT, dblDeviationFT, dblMinFT, dblMaxFT, dblAverageFT)  ' ����R
                    Call SetDeviation(0, dblGap, gFTNx_cnt, TotalFT, TotalSum2FT, dblDeviationFT, dblMinFT, dblMaxFT, dblAverageFT)   ' �S��R
                    'V2.2.0.0�O��
                    If stMultiBlock.gMultiBlock <> 0 Then
                        ' ������R�l�擾�̏ꍇ�̏W�v�f�[�^�ۑ�
                        With stToTalDataMulti(stExecBlkData.DataNo)
                            Call SetDeviation(rn, dblGap, .gFTNx_cnt, .TotalFT, .TotalSum2FT, .dblDeviationFT, .dblMinFT, .dblMaxFT, .dblAverageFT)  ' ����R
                            Call SetDeviation(0, dblGap, .gFTNx_cnt, .TotalFT, .TotalSum2FT, .dblDeviationFT, .dblMinFT, .dblMaxFT, .dblAverageFT)   ' �S��R
                        End With
                    End If
                    'V2.2.0.0�O��
                End If
                Else
                'NG�J�E���g�����L�^
                If JudgeMode = INITIAL_TEST Then
                    gITNg_cnt(rn) = gITNg_cnt(rn) + 1
                    gITNg_cnt(0) = gITNg_cnt(0) + 1
                ElseIf JudgeMode = FINAL_TEST Then
                    gFTNg_cnt(rn) = gFTNg_cnt(rn) + 1
                    gFTNg_cnt(0) = gFTNg_cnt(0) + 1
                End If
            End If

            ' �Ǖi���A�s�Ǖi��
            If JudgeMode = INITIAL_TEST Then
                dblOKRateIT(rn) = gITNx_cnt(rn) / (gITNx_cnt(rn) + gITNg_cnt(rn)) * 100.0
                dblNGRateIT(rn) = gITNg_cnt(rn) / (gITNx_cnt(rn) + gITNg_cnt(rn)) * 100.0
                dblOKRateIT(0) = gITNx_cnt(0) / (gITNx_cnt(0) + gITNg_cnt(0)) * 100.0
                dblNGRateIT(0) = gITNg_cnt(0) / (gITNx_cnt(0) + gITNg_cnt(0)) * 100.0
            ElseIf JudgeMode = FINAL_TEST Then
                dblOKRateFT(rn) = gFTNx_cnt(rn) / (gFTNx_cnt(rn) + gFTNg_cnt(rn)) * 100.0
                dblNGRateFT(rn) = gFTNg_cnt(rn) / (gFTNx_cnt(rn) + gFTNg_cnt(rn)) * 100.0
                dblOKRateFT(0) = gFTNx_cnt(0) / (gFTNx_cnt(0) + gFTNg_cnt(0)) * 100.0
                dblNGRateFT(0) = gFTNg_cnt(0) / (gFTNx_cnt(0) + gFTNg_cnt(0)) * 100.0
            End If

        Catch ex As Exception
            Call Z_PRINT("frmDistribution.StatisticalDataSave() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub

    ''' <summary>
    ''' ���v�f�[�^�̎擾
    ''' </summary>
    ''' <param name="JudgeMode"></param>
    ''' <param name="rn">��R�ԍ�</param>
    ''' <param name="dMin">�ŏ��l</param>
    ''' <param name="dMax">�ő�l</param>
    ''' <param name="dAve">���ϒl</param>
    ''' <param name="dDev">�W���΍�</param>
    ''' <remarks></remarks>
    Public Sub StatisticalDataGet(ByVal JudgeMode As Integer, ByRef rn As Integer, ByRef dMin As Double, ByRef dMax As Double, ByRef dAve As Double, ByRef dDev As Double)
        Try
            giDistributionResNo = rn

            If (JudgeMode = INITIAL_TEST) Then

                If Double.MaxValue = dblMinIT(giDistributionResNo) Then
                    dMin = 0.0
                Else
                    dMin = dblMinIT(giDistributionResNo)            ' �ŏ��l
                End If
                If Double.MinValue = dblMaxIT(giDistributionResNo) Then
                    dMax = 0.0
                Else
                    dMax = dblMaxIT(giDistributionResNo)            ' �ő�l
                End If

                'OK/NG���̕\��
                'lblGoodChip.Text = CStr(gITNx_cnt(giDistributionResNo))                      ' OK��
                'lblNgChip.Text = CStr(gITNg_cnt(giDistributionResNo))                        ' NG��
                'lblOKRate.Text = dblOKRateIT(giDistributionResNo).ToString("0.000")          ' �Ǖi��
                'lblNGRate.Text = dblNGRateIT(giDistributionResNo).ToString("0.000")          ' �s�Ǖi��
            Else
                If Double.MaxValue = dblMinFT(giDistributionResNo) Then
                    dMin = 0.0
                Else
                    dMin = dblMinFT(giDistributionResNo)               ' �ŏ��l
                End If
                If Double.MinValue = dblMaxFT(giDistributionResNo) Then
                    dMax = 0.0
                Else
                    dMax = dblMaxFT(giDistributionResNo)               ' �ő�l
                End If

                'OK/NG���̕\��
                'lblGoodChip.Text = CStr(gFTNx_cnt(giDistributionResNo))                      ' OK��
                'lblNgChip.Text = CStr(gFTNg_cnt(giDistributionResNo))                        ' NG��
                'lblOKRate.Text = dblOKRateFT(giDistributionResNo).ToString("0.000")          ' �Ǖi��
                'lblNGRate.Text = dblNGRateFT(giDistributionResNo).ToString("0.000")          ' �s�Ǖi��
            End If


            If (JudgeMode = INITIAL_TEST) Then
                dAve = dblAverageIT(giDistributionResNo)    ' ����(IT)
            Else
                dAve = dblAverageFT(giDistributionResNo)    ' ����(FT)
            End If

            If (JudgeMode = INITIAL_TEST) Then
                dDev = dblDeviationIT(giDistributionResNo) ' �W���΍�(IT)
            Else
                dDev = dblDeviationFT(giDistributionResNo) ' �W���΍�(FT)
            End If

        Catch ex As Exception
            Call Z_PRINT("frmDistribution.StatisticalDataGet() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Sub

    ' 'V2.2.0.0�O��
    ''' <summary>
    ''' ���v�f�[�^�̎擾
    ''' </summary>
    ''' <param name="JudgeMode"></param>
    ''' <param name="rn">��R�ԍ�</param>
    ''' <param name="dMin">�ŏ��l</param>
    ''' <param name="dMax">�ő�l</param>
    ''' <param name="dAve">���ϒl</param>
    ''' <param name="dDev">�W���΍�</param>
    ''' <remarks></remarks>
    Public Sub StatisticalDataGetMulti(ByVal JudgeMode As Integer, ByRef rn As Integer, ByRef dMin As Double, ByRef dMax As Double, ByRef dAve As Double, ByRef dDev As Double, ByVal MultiBlkNo As Integer)

        Try
            giDistributionResNo = rn

            With stToTalDataMulti(MultiBlkNo)

                If (JudgeMode = INITIAL_TEST) Then

                    If Double.MaxValue = .dblMinIT(giDistributionResNo) Then
                        dMin = 0.0
                    Else
                        dMin = .dblMinIT(giDistributionResNo)            ' �ŏ��l
                    End If
                    If Double.MinValue = .dblMaxIT(giDistributionResNo) Then
                        dMax = 0.0
                    Else
                        dMax = .dblMaxIT(giDistributionResNo)            ' �ő�l
                    End If

                    'OK/NG���̕\��
                    'lblGoodChip.Text = CStr(gITNx_cnt(giDistributionResNo))                      ' OK��
                    'lblNgChip.Text = CStr(gITNg_cnt(giDistributionResNo))                        ' NG��
                    'lblOKRate.Text = dblOKRateIT(giDistributionResNo).ToString("0.000")          ' �Ǖi��
                    'lblNGRate.Text = dblNGRateIT(giDistributionResNo).ToString("0.000")          ' �s�Ǖi��
                Else
                    If Double.MaxValue = .dblMinFT(giDistributionResNo) Then
                        dMin = 0.0
                    Else
                        dMin = .dblMinFT(giDistributionResNo)               ' �ŏ��l
                    End If
                    If Double.MinValue = .dblMaxFT(giDistributionResNo) Then
                        dMax = 0.0
                    Else
                        dMax = .dblMaxFT(giDistributionResNo)               ' �ő�l
                    End If

                    'OK/NG���̕\��
                    'lblGoodChip.Text = CStr(gFTNx_cnt(giDistributionResNo))                      ' OK��
                    'lblNgChip.Text = CStr(gFTNg_cnt(giDistributionResNo))                        ' NG��
                    'lblOKRate.Text = dblOKRateFT(giDistributionResNo).ToString("0.000")          ' �Ǖi��
                    'lblNGRate.Text = dblNGRateFT(giDistributionResNo).ToString("0.000")          ' �s�Ǖi��
                End If


                If (JudgeMode = INITIAL_TEST) Then
                    dAve = .dblAverageIT(giDistributionResNo)    ' ����(IT)
                Else
                    dAve = .dblAverageFT(giDistributionResNo)    ' ����(FT)
                End If

                If (JudgeMode = INITIAL_TEST) Then
                    dDev = .dblDeviationIT(giDistributionResNo) ' �W���΍�(IT)
                Else
                    dDev = .dblDeviationFT(giDistributionResNo) ' �W���΍�(FT)
                End If

            End With


        Catch ex As Exception
            Call Z_PRINT("frmDistribution.StatisticalDataGetMulti() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Sub

    ''' <summary>
    ''' NG�J�E���^�[
    ''' </summary>
    ''' <returns></returns>
    Public Function CalcNgCounter() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then

                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.NG_Counter = .stCounter1.ITHigh + .stCounter1.ITLow + .stCounter1.ITOpen + .stCounter1.FTHigh + .stCounter1.FTLow + .stCounter1.FTOpen + .stCounter1.Pattern + .stCounter1.VaNG + .stCounter1.StdNg
                    .stCounter1.Total_NG_Counter = .stCounter1.Total_ITHigh + .stCounter1.Total_ITLow + .stCounter1.Total_ITOpen + .stCounter1.Total_FTHigh + .stCounter1.Total_FTLow + .stCounter1.Total_FTOpen + .stCounter1.Total_Pattern + .stCounter1.Total_VaNG + .stCounter1.Total_StdNg
                End With

            End If

        Catch ex As Exception

        End Try

    End Function
    ' 'V2.2.0.0�O��

    ' 'V2.2.0.0�O��
    ''' <summary>
    ''' OK�J�E���^�[�̍X�V ������R�l�p�̃J�E���^�[
    ''' </summary>
    ''' <returns></returns>
    Public Function SetOkCounterMulti() As Integer
        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.OK_Counter = .stCounter1.OK_Counter + 1
                    .stCounter1.Total_OK_Counter = .stCounter1.Total_OK_Counter + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' IT-HI�J�E���^�[�̍X�V ������R�l�p�̃J�E���^�[
    ''' </summary>
    ''' <returns></returns>
    Public Function SetITHICounterMulti() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.ITHigh = .stCounter1.ITHigh + 1
                    .stCounter1.Total_ITHigh = .stCounter1.Total_ITHigh + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' IT-LO�J�E���^�[�̍X�V ������R�l�p�̃J�E���^�[
    ''' </summary>
    ''' <returns></returns>
    Public Function SetITLowCounterMulti() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.ITLow = .stCounter1.ITLow + 1
                    .stCounter1.Total_ITLow = .stCounter1.Total_ITLow + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' IT-LO�J�E���^�[�̍X�V ������R�l�p�̃J�E���^�[
    ''' </summary>
    ''' <returns></returns>
    Public Function SetITLOCounterMulti() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.ITLow = .stCounter1.ITLow + 1
                    .stCounter1.Total_ITLow = .stCounter1.Total_ITLow + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function


    ''' <summary>
    ''' IT-OPEN�J�E���^�[�̍X�V ������R�l�p�̃J�E���^�[
    ''' </summary>
    ''' <returns></returns>
    Public Function SetITOpenCounterMulti() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.ITOpen = .stCounter1.ITOpen + 1
                    .stCounter1.Total_ITOpen = .stCounter1.Total_ITOpen + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' FT-HI�J�E���^�[�̍X�V ������R�l�p�̃J�E���^�[
    ''' </summary>
    ''' <returns></returns>
    Public Function SetFTHighCounterMulti() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.FTHigh = .stCounter1.FTHigh + 1
                    .stCounter1.Total_FTHigh = .stCounter1.Total_FTHigh + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' FT-LO�J�E���^�[�̍X�V ������R�l�p�̃J�E���^�[
    ''' </summary>
    ''' <returns></returns>
    Public Function SetFTLOCounterMulti() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.FTLow = .stCounter1.FTLow + 1
                    .stCounter1.Total_FTLow = .stCounter1.Total_FTLow + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' FT-Open�J�E���^�[�̍X�V ������R�l�p�̃J�E���^�[
    ''' </summary>
    ''' <returns></returns>
    Public Function SetFTOpenCounterMulti() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.FTOpen = .stCounter1.FTOpen + 1
                    .stCounter1.Total_FTOpen = .stCounter1.Total_FTOpen + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' Pattern�J�E���^�[�̍X�V ������R�l�p�̃J�E���^�[
    ''' </summary>
    ''' <returns></returns>
    Public Function SetPatternCounterMulti() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.Pattern = .stCounter1.Pattern + 1
                    .stCounter1.Total_Pattern = .stCounter1.Total_Pattern + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' �ϓ���(VaNG)�J�E���^�[�̍X�V ������R�l�p�̃J�E���^�[
    ''' </summary>
    ''' <returns></returns>
    Public Function SetVaNGCounterMulti() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.VaNG = .stCounter1.VaNG + 1
                    .stCounter1.Total_VaNG = .stCounter1.Total_VaNG + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' ���x�Z���T�[�X�^���_�[�h����NG(StdNg)�J�E���^�[�̍X�V ������R�l�p�̃J�E���^�[
    ''' </summary>
    ''' <returns></returns>
    Public Function SetStdNgCounterMulti() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.StdNg = .stCounter1.StdNg + 1
                    .stCounter1.Total_StdNg = .stCounter1.Total_StdNg + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' �㏸������NG(ValLow)�J�E���^�[�̍X�V ������R�l�p�̃J�E���^�[
    ''' </summary>
    ''' <returns></returns>
    Public Function SetValLowCounterMulti() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.ValLow = .stCounter1.ValLow + 1
                    .stCounter1.Total_ValLow = .stCounter1.Total_ValLow + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function


    ''' <summary>
    ''' �g���~���O�J�E���^�[�̍X�V ������R�l�p�̃J�E���^�[
    ''' </summary>
    ''' <returns></returns>
    Public Function SetTrimCounterMulti() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.TrimCounter = .stCounter1.TrimCounter + 1
                    .stCounter1.Total_TrimCounter = .stCounter1.Total_TrimCounter + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function
    ' 'V2.2.0.0�O��

    'V2.2.0.0�O��


    ''' <summary>
    ''' �}���`�u���b�N�̒�R�l�J�E���^�̃N���A
    ''' </summary>
    ''' <returns></returns>
    Public Function ClearMultiLotCountData() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                For blk As Integer = 0 To MAXBlock
                    With stToTalDataMulti(blk)

                        .stCounter1.TrimCounter = 0                   ' ���ݸސ�(ܰ�������)
                        .stCounter1.OK_Counter = 0                    ' OK��
                        .stCounter1.NG_Counter = 0                    ' NG��
                        .stCounter1.ITHigh = 0                        ' �����������l�ُ�
                        .stCounter1.ITLow = 0                         ' �������艺���l�ُ�
                        .stCounter1.ITOpen = 0                        ' ����l�ُ�
                        .stCounter1.FTHigh = 0                        ' �ŏI�������l�ُ�
                        .stCounter1.FTLow = 0                         ' �ŏI���艺���l�ُ�
                        .stCounter1.FTOpen = 0                        ' ����l�ُ�
                        .stCounter1.Pattern = 0                       ' �J�b�g�ʒu�␳�̔��� 'V1.2.0.0�B
                        .stCounter1.VaNG = 0                          ' �đ���ω��ʃG���[V2.0.0.0�A
                        .stCounter1.StdNg = 0                         ' �X�^���_�[�h��R����G���[ 'V2.0.0.0�N
                        .stCounter1.ValLow = 0                        ' �J�b�g���㏸���ω�Low�ُ�        'V2.2.0.029
                        .stCounter1.ValHigh = 0                       ' �J�b�g���㏸���ω�High�ُ�       'V2.2.0.029
                        ' ���b�g�ʎZ
                        .stCounter1.PlateCounter = 0                  ' ��J�E���^�[
                        .stCounter1.Total_TrimCounter = 0             ' ��R�g�[�^��������
                        .stCounter1.Total_OK_Counter = 0              ' OK��
                        .stCounter1.Total_NG_Counter = 0              ' NG��
                        .stCounter1.Total_ITHigh = 0                  ' �����������l�ُ�
                        .stCounter1.Total_ITLow = 0                   ' �������艺���l�ُ�
                        .stCounter1.Total_ITOpen = 0                  ' ����l�ُ�
                        .stCounter1.Total_FTHigh = 0                  ' �ŏI�������l�ُ�
                        .stCounter1.Total_FTLow = 0                   ' �ŏI���艺���l�ُ�
                        .stCounter1.Total_FTOpen = 0                  ' ����l�ُ�
                        .stCounter1.Total_Pattern = 0                 ' �J�b�g�ʒu�␳�̔��� 'V1.2.0.0�B
                        .stCounter1.Total_VaNG = 0                    ' �đ���ω��ʃG���[V2.0.0.0�A
                        .stCounter1.Total_StdNg = 0                   ' �X�^���_�[�h��R����G���[ 'V2.0.0.0�N
                        .stCounter1.Total_ValLow = 0                  ' �J�b�g���㏸���ω�Low�ُ�        'V2.2.0.029
                        .stCounter1.Total_ValHigh = 0                 ' �J�b�g���㏸���ω�High�ُ�       'V2.2.0.029

                        For cnt As Integer = 0 To MAX_RES_USER
                            .TrimCounter(cnt) = 0                    ' �g���~���O���J�E���^�[
                            .Total_TrimCounter(cnt) = 0              ' �g���~���O���J�E���^�[�g�[�^�� 
                        Next cnt

                    End With
                Next blk
            End If

        Catch ex As Exception

        End Try

    End Function
    'V2.2.0.0�O��

    'V2.2.0.0�O��


    ''' <summary>
    ''' �}���`�u���b�N�̒�R�l�J�E���^�̃N���A
    ''' </summary>
    ''' <returns></returns>
    Public Function ClearMultiCountPlateData() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                For blk As Integer = 0 To MAXBlock
                    With stToTalDataMulti(blk)

                        .stCounter1.TrimCounter = 0                   ' ���ݸސ�(ܰ�������)
                        .stCounter1.OK_Counter = 0                    ' OK��
                        .stCounter1.NG_Counter = 0                    ' NG��
                        .stCounter1.ITHigh = 0                        ' �����������l�ُ�
                        .stCounter1.ITLow = 0                         ' �������艺���l�ُ�
                        .stCounter1.ITOpen = 0                        ' ����l�ُ�
                        .stCounter1.FTHigh = 0                        ' �ŏI�������l�ُ�
                        .stCounter1.FTLow = 0                         ' �ŏI���艺���l�ُ�
                        .stCounter1.FTOpen = 0                        ' ����l�ُ�
                        .stCounter1.Pattern = 0                       ' �J�b�g�ʒu�␳�̔��� 'V1.2.0.0�B
                        .stCounter1.VaNG = 0                          ' �đ���ω��ʃG���[V2.0.0.0�A
                        .stCounter1.StdNg = 0                         ' �X�^���_�[�h��R����G���[ 'V2.0.0.0�N
                        .stCounter1.ValLow = 0                        ' �J�b�g���㏸���ω�Low�ُ�        'V2.2.0.029
                        .stCounter1.ValHigh = 0                       ' �J�b�g���㏸���ω�High�ُ�       'V2.2.0.029

                        For cnt As Integer = 0 To MAX_RES_USER
                            .TrimCounter(cnt) = 0                    ' �g���~���O���J�E���^�[
                        Next cnt

                    End With
                Next blk
            End If

        Catch ex As Exception

        End Try

    End Function
    'V2.2.0.0�O��


#End Region

End Class