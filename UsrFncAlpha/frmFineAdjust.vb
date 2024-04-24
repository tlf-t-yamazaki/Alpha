'===============================================================================
'   Description  : �g���~���O���s���ꎞ��~����
'
'   Copyright(C) : TOWA LASERFRONT CORP. 2012
'
'===============================================================================
Option Strict Off
Option Explicit On

Imports LaserFront.Trimmer.DllSysPrm.SysParam

Friend Class frmFineAdjust
    Inherits System.Windows.Forms.Form
    Implements ICommonMethods              'V2.2.0.0�@

    '========================================================================================
    '   �萔�E�ϐ���`
    '========================================================================================
#Region "�萔�E�ϐ���`"
    '===========================================================================
    '   �萔��`
    '===========================================================================
    Public Const MOVE_NEXT As Integer = 0
    Public Const MOVE_NOT As Integer = 1

    '----- �������[�h -----
    Private Const MD_INI As Integer = 0                                 ' �����G���g�����[�h
    Private Const MD_CHK As Integer = 1                                 ' �p���G���g�����[�h

    '===========================================================================
    '   �����o�ϐ���`
    '===========================================================================
    Private m_BlockSizeX As Double
    Private m_BlockSizeY As Double
    Private m_bpOffX As Double
    Private m_bpOffY As Double
    Private m_sysPrm As SYSPARAM_PARAM
    Private stJOG As JOG_PARAM                                          ' �����(BP��JOG����)�p�p�����[�^ (Globals.vb�̋��ʊ֐����g�p)
    Private dblTchMoval(3) As Double                                    ' �߯��ޔ���(0=�߯�, 1=HIGH�߯�, 2=Pause Time(Sec))
    Private mExit_flg As Short                                          ' ����
    Private mMd As Integer = MD_INI                                     ' �������[�h
    Private m_TenKeyFlg As Boolean = False
    Private m_LaserOnOffFlag As Boolean = False

#End Region

    '========================================================================================
    '   ���\�b�h��`
    '========================================================================================
#Region "�����l�ݒ菈��"
    '''=========================================================================
    ''' <summary>�����l�ݒ菈��</summary>
    ''' <param name="SysPrm"></param>
    ''' <param name="digL"></param>
    ''' <param name="digH"></param>
    ''' <param name="curPltNo"></param>
    ''' <param name="curBlkNo"></param>
    '''=========================================================================
    Public Sub SetInitialData(ByRef SysPrm As SYSPARAM_PARAM, _
                        ByVal digL As Integer, ByVal digH As Integer, _
                        ByRef curPltNo As Integer, ByRef curBlkNo As Integer)

        Try
            'CbDigSwH.SelectedIndex = digH
            'CbDigSwL.SelectedIndex = digL
            gCurBlockNo = curBlkNo
            gCurPlateNo = curPltNo
            m_sysPrm = SysPrm
            'gFrmEndStatus = cFRS_NORMAL

            If (gbChkboxHalt = True) Then                                       '###009
                BtnADJ.Text = "ADJ ON"                                          '###009
                BtnADJ.BackColor = System.Drawing.Color.Yellow                  '###009
            Else                                                                '###009
                BtnADJ.Text = "ADJ OFF"                                         '###009
                BtnADJ.BackColor = System.Drawing.SystemColors.Control          '###009
            End If                                                              '###009

            ' ���x�����ݒ�(���{��/�p��)
            'BtnEdit.Text = "�f�[�^�ҏW"                                      ' "�f�[�^�ҏW" ###014
            '-----###204 -----
            Me.Label3.Text = "����"                                    ' "����" 
            'CbDigSwH.Items(0) = "�O�F�\���Ȃ�"
            'CbDigSwH.Items(1) = "�P�F�m�f�̂ݕ\��"
            'CbDigSwH.Items(2) = "�Q�F�S�ĕ\��"
            '-----###204 -----
            '----- ###268�� -----
            '�uTen Key On/Off�v�{�^���̏����l���V�X�p�����ݒ肷��
            If (giTenKey_Btn = 0) Then                                          ' �ꎞ��~��ʂł́uTen Key On/Off�v�{�^���̏����l(0:ON(����l), 1:OFF)
                gbTenKeyFlg = True
                BtnTenKey.Text = "Ten Key On"
                BtnTenKey.BackColor = System.Drawing.Color.Pink
            Else
                gbTenKeyFlg = False
                BtnTenKey.Text = "Ten Key Off"
                BtnTenKey.BackColor = System.Drawing.SystemColors.Control
            End If

            'gbTenKeyFlg = True                                                 ' �uTen Key On�v��� ###242
            '----- ###268�� -----
            '----- ###269�� -----
            ' �ꎞ��~��ʂł̃V�X�p���uBP�I�t�Z�b�g��������/���Ȃ��v�w��ɂ����{�^������ݒ肷��
            Call Sub_SetBtnArrowEnable()
            '----- ###269�� -----

            'for�@��R����
            '�ڕW�l
            '�J�b�g�I�t
            '�X�s�[�h
            '���H�����ԍ�
            'next
            'txtExCamPosX.Text = m_sysPrm.stDEV.gfExCmX.ToString


            'V2.2.0.0�C ��
            If giMouseClickMove = 1 Then
                BtnClickEnable.BackColor = SystemColors.Control
                gbTenKeyFlg = False
            Else
                gbTenKeyFlg = True
            End If
            'V2.2.0.0�C ��


        Catch ex As Exception
            Dim strMsg As String
            strMsg = "frmFineAdjust.Form_Initialize_Renamed() TRAP ERROR = " + ex.Message
            MsgBox(strMsg)
        End Try

    End Sub
#End Region
    '-----###269��-----
#Region "���{�^����������/�񊈐�������"
    '''=========================================================================
    ''' <summary>���{�^����������/�񊈐�������</summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub Sub_SetBtnArrowEnable()

        Dim bFlg As Boolean
        Dim strMsg As String

        Try
            '  �ꎞ��~��ʂł̃V�X�p���uBP�I�t�Z�b�g��������/���Ȃ��v�w��ɂ����{�^������ݒ肷��
            If (giBpAdj_HALT = 0) Then                                          ' BP�I�t�Z�b�g�������� ?
                bFlg = True

                'V2.2.0.0�@��
                Form1.SetActiveJogMethod(AddressOf Me.JogKeyDown,
                                                  AddressOf Me.JogKeyUp,
                                                  AddressOf Me.MoveToCenter)    'V6.0.0.0�I
                'V2.2.0.0�@��

            Else                                                                ' BP�I�t�Z�b�g�������Ȃ�
                bFlg = False
                gbTenKeyFlg = False
                BtnTenKey.Enabled = False                                       '�uTen Key Off�v�{�^���񊈐���
                BtnTenKey.Text = "Ten Key Off"
                BtnTenKey.BackColor = System.Drawing.SystemColors.Control

                Form1.SetActiveJogMethod(Nothing, Nothing, Nothing)    'V2.2.0.0�@

            End If

            If giMouseClickMove = 1 Then
                BtnClickEnable.Enabled = True
                BtnClickEnable.Visible = True
                BtnClickEnable.BackColor = SystemColors.Control
            Else
                BtnClickEnable.Enabled = False
                BtnClickEnable.Visible = False
            End If

            ' ���{�^��������/�񊈐���
            BtnJOG_0.Enabled = bFlg
            BtnJOG_1.Enabled = bFlg
            BtnJOG_2.Enabled = bFlg
            BtnJOG_3.Enabled = bFlg
            BtnJOG_4.Enabled = bFlg
            BtnJOG_5.Enabled = bFlg
            BtnJOG_6.Enabled = bFlg
            BtnJOG_7.Enabled = bFlg
            BtnHI.Enabled = bFlg

            ' Moving Pitch������/�񊈐���
            GrpPithPanel.Enabled = bFlg

        Catch ex As Exception
            strMsg = "frmFineAdjust.Sub_SetBtnArrowEnable() TRAP ERROR = " + ex.Message
            MsgBox(strMsg)
        End Try
    End Sub
#End Region
    '----- ###269��-----
    '----- ###260��-----
#Region "�^�C�}�[��~"
    '''=========================================================================
    ''' <summary>�^�C�}�[��~</summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function Sub_StopTimer() As Integer

        TmKeyCheck.Enabled = False

    End Function
#End Region
    '----- ###260��-----

#Region "�X�e�[�W�|�W�V�����擾����"
    '''=========================================================================
    ''' <summary>�X�e�[�W�|�W�V�����擾�����i���s��Ɏ擾�j</summary>
    '''=========================================================================
    Public Sub GetStagePosInfo(ByRef pltNo As Integer, ByRef blkNo As Integer)

        Try
            pltNo = gCurPlateNo
            blkNo = gCurBlockNo

        Catch ex As Exception
            Dim strMsg As String
            strMsg = "frmFineAdjust.GetStagePosInfo() TRAP ERROR = " + ex.Message
            MsgBox(strMsg)
        End Try
    End Sub
#End Region

#Region "�I���߂�l�擾����"
    '''=========================================================================
    ''' <summary>�I���߂�l�擾�����i���s��Ɏ擾�j</summary>
    '''=========================================================================
    Public Function GetReturnVal() As Integer

        Try
            Return (mExit_flg)

        Catch ex As Exception
            Dim strMsg As String
            strMsg = "frmFineAdjust.GetReturnVal() TRAP ERROR = " + ex.Message
            MsgBox(strMsg)
        End Try
    End Function
#End Region

    '========================================================================================
    '   ��ʏ���
    '========================================================================================
#Region "�t�H�[�����[�h����"
    '''=========================================================================
    ''' <summary>�t�H�[�����[�h����</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub frmFineAdjust_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Dim stPos As System.Drawing.Point
        'Dim stGetPos As System.Drawing.Point
        'Dim r As Integer                                                ' ###237
        Dim strMsg As String

        Try
            ' �\���ʒu�̒���
            'stPos = Form1.Text4.PointToScreen(stGetPos)
            'stPos.X = stPos.X - 2
            'stPos.Y = stPos.Y - 2
            'Me.Location = stPos
            Me.Location = New Point(Form1.VideoLibrary1.Location.X + Form1.VideoLibrary1.Size.Width + 6, Form1.Grpcmds.Location.Y)
            'Me.Height = Form1.frmInfo.Location.Y - Form1.Grpcmds.Location.Y
            Me.Height = Form1.Grpcmds.Size.Height
            ' BpOffset�̌��ݒl�ݒ�
            GetBpOffset(m_bpOffX, m_bpOffY)
            txtBpOffX.Text = m_bpOffX.ToString
            txtBpOffY.Text = m_bpOffY.ToString

            ' BlockSize�̌��ݒl�擾
            GetBlockSize(m_BlockSizeX, m_BlockSizeY)

            '----- ###139�� -----
            ' ���C����ʂ́u���Y�O���t�\��/��\���{�^���v���瓖��ʂ́u���Y�O���t�\��/��\���{�^���v��ݒ肷��
            'If (gTkyKnd = KND_CHIP Or gTkyKnd = KND_NET) Then
            '    chkDistributeOnOff.Text = Form1.chkDistributeOnOff.Text
            '    chkDistributeOnOff.Checked = Form1.chkDistributeOnOff.Checked
            '    GrpDistribute.Visible = True                        '�u���Y�O���t�{�^���v�\��

            'Else
            'GrpDistribute.Visible = False                       '�u���Y�O���t�{�^���v��\��
            'End If
            '----- ###139�� -----

            '----- ###237�� -----
            ' ���H�����ԍ���ݒ肷��(FL��)
            'If (gSysPrm.stRAT.giOsc_Res = OSCILLATOR_FL) Then
            '    Call QRATE(stCND.Freq(ADJ_CND_NUM))                     ' Q���[�g�ݒ�(KHz)
            '    r = FLSET(FLMD_CNDSET, ADJ_CND_NUM)                     ' ���H�����ԍ��ݒ�(�ꎞ��~��ʗp)
            'Else
            '' '' ''Call QRATE(gSysPrm.stDEV.gfLaserQrate)                  ' Q���[�g�ݒ�(KHz) �����[�U�����pQ���[�g��ݒ�
            'End If
            '----- ###237�� -----

            Call PrepareMessages(gSysPrm.stTMN.giMsgTyp)
            ' �t�H�[�J�X�̐ݒ�(����ɂ���ăe���L�[�̃C�x���g���擾�ł���)
            Me.KeyPreview = True
            Me.Activate()                                               ' ###046

        Catch ex As Exception
            strMsg = "frmFineAdjust.frmFineAdjust_Load() TRAP ERROR = " + ex.Message
            MsgBox(strMsg)
        End Try
    End Sub
#End Region

#Region "�t�H�[�����\�����ꂽ���̏���"
    '''=========================================================================
    ''' <summary>�t�H�[�����\�����ꂽ���̏���</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub frmFineAdjust_Shown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Shown

        Dim r As Integer = cFRS_NORMAL
        Dim strMSG As String

        Try
            ' �ꎞ��~��ʏ������C����Call����
            mExit_flg = 0                                               ' �I���t���O = 0
            Call ZCONRST()                                              ' �ݿ�ٷ�ׯ�����
            TmKeyCheck.Interval = 10
            TmKeyCheck.Enabled = True                                   ' �^�C�}�[�J�n
            Return

            ' �g���b�v�G���[������
        Catch ex As Exception
            strMSG = "frmFineAdjust.frmFineAdjust_Shown() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            mExit_flg = cERR_TRAP                                       ' Return�l = ��O�G���[
        End Try

        gbExitFlg = True                                                ' �I���t���OON
        Call LASEROFF()                                                 ' ###237
        Me.Close()
    End Sub
#End Region

#Region "�L�[���̓`�F�b�N�^�C�}�[����"
    '''=========================================================================
    ''' <summary>�L�[���̓`�F�b�N�^�C�}�[����</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub TmKeyCheck_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles TmKeyCheck.Tick

        Dim r As Short
        Dim strMSG As String

        Try

            If gbExitFlg Then
                Exit Sub
            End If
            ' �L�[���̓`�F�b�N����
            TmKeyCheck.Enabled = False                                  ' �^�C�}�[��~
            r = MainProc(mMd)                                           ' �ꎞ��~��ʏ���
            If (r = cFRS_NORMAL) Then                                   ' ����߂� 
                TmKeyCheck.Enabled = True                               ' �^�C�}�[�J�n
                Return
            End If

            '----- ###219�� -----
            ' Z �L�[�����Ȃ� Z On/OFF���� 
            If (r = cFRS_ERR_Z) Then                                    ' Z SW���� ?
                If (stJOG.bZ = True) Then                               ' Z ON ? 
                    r = Prob_On()
                Else                                                    ' Z OFF
                    r = Prob_Off()
                End If
                ' �G���[�Ȃ烁�b�Z�[�W��\�����ăG���[���^�[��
                r = Form1.System1.EX_ZGETSRVSIGNAL(gSysPrm, r, 0)       ' �G���[�Ȃ烁�b�Z�[�W��\������
                If (r <> cFRS_NORMAL) Then
                    mExit_flg = r                                       ' �G���[���^�[�� 
                    Return
                End If

                ' Z�����v�̓_��/����
                If (stJOG.bZ = True) Then
                    Call LAMP_CTRL(LAMP_Z, True)
                Else
                    Call LAMP_CTRL(LAMP_Z, False)
                End If

                TmKeyCheck.Enabled = True                               ' �^�C�}�[�J�n
                Return
            End If
            '----- ###219�� -----

            ' START/RESET�L�[�����܂��̓G���[�Ȃ�I��
            If (r = cFRS_ERR_START) Then r = cFRS_NORMAL

            mExit_flg = r

            ' �g���b�v�G���[������
        Catch ex As Exception
            strMSG = "frmFineAdjust.TmKeyCheck_Tick() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            mExit_flg = cERR_TRAP                                       ' Return�l = ��O�G���[
        End Try

        gbExitFlg = True                                                ' �I���t���OON
        Call LASEROFF()                                                 ' ###237
        Me.Close()
    End Sub
#End Region

#Region "���C������"
    '''=========================================================================
    ''' <summary>���C������"</summary>
    ''' <param name="Md">(I/O)�������[�h
    ''' �@�@�@�@�@�@�@�@�@�@�@MD_INI=�����G���g��, MD_CHK=�p���G���g��</param>
    ''' <returns>cFRS_NORMAL   = OK(START��)
    '''          cFRS_ERR_RST  = Cancel(RESET��)
    '''          -1�ȉ�        = �G���[</returns>
    '''=========================================================================
    Private Function MainProc(ByRef Md As Integer) As Short

        Dim mdAdjx As Double = 0.0                                      ' ��ެ�ĈʒuX(���g�p)
        Dim mdAdjy As Double = 0.0                                      ' ��ެ�ĈʒuY(���g�p)
        Dim r As Short
        Dim strMSG As String
        Dim cControl As Control = Me.ActiveControl

        Try
            '-------------------------------------------------------------------
            '   ��������
            '-------------------------------------------------------------------
            If (Md = MD_INI) Then                                       ' �����G���g��
                ' JOG�p�����[�^�ݒ� 
                stJOG.Md = MODE_BP                                      ' ���[�h(1:BP�ړ�)
                stJOG.Md2 = MD2_BUTN                                    ' ���̓��[�h(0:������ݓ���, 1:�ݿ�ٓ���)
                '                                                       ' �L�[�̗L��(1)/����(0)�w��
                'stJOG.Opt = CONSOLE_SW_RESET + CONSOLE_SW_START
                stJOG.Opt = CONSOLE_SW_RESET + CONSOLE_SW_START + CONSOLE_SW_ZSW ' ###219
                stJOG.PosX = 0.0                                        ' BP X�ʒu(BP�̾��X)
                stJOG.PosY = 0.0                                        ' BP Y�ʒu(BP�̾��Y)
                stJOG.BpOffX = mdAdjx + m_bpOffX                        ' BP�̾��X 
                stJOG.BpOffY = mdAdjy + m_bpOffY                        ' BP�̾��Y 
                stJOG.BszX = m_BlockSizeX                               ' ��ۯ�����X 
                stJOG.BszY = m_BlockSizeY                               ' ��ۯ�����Y
                txtBpOffX.ShortcutsEnabled = False                      ' ###047 �E�N���b�N���j���[��\�����Ȃ� 
                txtBpOffY.ShortcutsEnabled = False                      '  
                stJOG.TextX = txtBpOffX                                 ' BP X�ʒu�\���p÷���ޯ��
                stJOG.TextY = txtBpOffY                                 ' BP Y�ʒu�\���p÷���ޯ��
                stJOG.cgX = m_bpOffX                                    ' �ړ���X (BP�̾��X)
                stJOG.cgY = m_bpOffY                                    ' �ړ���Y (BP�̾��Y)
                stJOG.BtnHI = BtnHI                                     ' HI�{�^��
                stJOG.BtnZ = BtnZ                                       ' Z�{�^��
                stJOG.BtnSTART = BtnSTART                               ' START�{�^��
                stJOG.BtnRESET = BtnRESET                               ' RESET�{�^��
                stJOG.BtnHALT = BtnHALT                                 ' HALT�{�^��
                Call JogEzInit(stJOG, TBarLowPitch, TBarHiPitch, TBarPause, LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval)
                stJOG.Flg = -1                                          ' �e��ʂ�OK/Cancel���݉����׸�
                Md = MD_CHK
                stJOG.bZ = False                                        ' Jog��Z�L�[��� = Z Off ###219
                Call LAMP_CTRL(LAMP_Z, False)                           ' ###219 
            End If

STP_RETRY:
            'Call Me.Focus()                                            ' �� ��������ƃe���L�[��KeyUp/KeyDown�C�x���g�������Ă��Ȃ��Ȃ�

            ' ����~���`�F�b�N
            r = Form1.System1.Sys_Err_Chk_EX(gSysPrm, giAppMode)
            If (r <> cFRS_NORMAL) Then                                  ' ����~�����o ?
                Call Form1.AppEndDataSave()                                       ' ��ċ����I�������ް��ۑ��m�F
                Call Form1.AplicationForcedEnding()                     ' ��ċ����I������
                End                                                     ' �A�v�������I��
                '                Return (r)
            End If

            ''----- ###209�� -----
            '' �J�o�[���m�F����(SL436R���Ŏ蓮���[�h��)
            'If (gSysPrm.stTMN.gsKeimei = MACHINE_TYPE_SL436) And (bFgAutoMode = False) Then
            '    Call COVERLATCH_CLEAR()                                 ' �J�o�[�J���b�`�̃N���A
            '    r = FrmReset.Sub_CoverCheck()
            '    If (r < cFRS_NORMAL) Then                               ' ����~�����o ?
            '        Return (r)
            '    End If
            'End If
            ''----- ###209�� -----
            'V2.2.0.0�@ ��
            'If System.Windows.Forms.Form.ActiveForm IsNot Nothing Then
            '    If System.Windows.Forms.Form.ActiveForm.Text <> "ADJFINE" Then
            '        Call ClearInpKey()
            '    End If
            'End If
            'V2.2.0.0�@ ��
            ' �R���\�[���L�[���̓��͑҂�
            'stJOG.Flg = -1                                             ' �e��ʂ�OK/Cancel���݉����׸�
            r = JogEzMove_Ex(stJOG, gSysPrm, TBarLowPitch, TBarHiPitch, TBarPause, LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval)
            If (r < cFRS_NORMAL) Then                                   ' �G���[�Ȃ�I��
                Return (r)
            End If

            ' �R���\�[���L�[�`�F�b�N
            If (r = cFRS_ERR_START) Then                                ' START SW���� ?
                ' DIG-SW�ݒ�
                'Call Form1.SetMoveMode(CbDigSwL.SelectedIndex, CbDigSwH.SelectedIndex)
                ' BP�I�t�Z�b�g�X�V(�^�C�~���O�ɂ���ċ󔒂œ����Ă���ꍇ�g���b�v�G���[�ƂȂ�̂Ń`�F�b�N���� ###014)
                If (txtBpOffX.Text <> "") And (txtBpOffY.Text <> "") Then
                    Call SetBpOffset(Double.Parse(txtBpOffX.Text), Double.Parse(txtBpOffY.Text))
                End If
                Return (cFRS_ERR_START)

            ElseIf (r = cFRS_ERR_RST) Then                              ' RESET SW���� ?
                Return (cFRS_ERR_RST)

                '----- ###219�� -----
            ElseIf (r = cFRS_ERR_Z) Then                                ' Z SW���� ?
                Return (cFRS_ERR_Z)
                '----- ###219�� -----
            End If

            'Loop While (stJOG.Flg = -1)

            '' ����ʂ���OK/Cancel���݉����Ȃ�r�ɖߒl��ݒ肷��
            'If (stJOG.Flg <> -1) Then
            '    r = stJOG.Flg
            'End If

STP_END:
            Return (r)                                                  ' Return�l�ݒ�

            ' �g���b�v�G���[������
        Catch ex As Exception
            strMSG = "frmFineAdjust.MainProc() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Return�l = ��O�G���[
        End Try
    End Function
#End Region

    '========================================================================================
    '   ���C����ʂ̃{�^������������
    '========================================================================================
#Region "ADJ���݉���������"
    '''=========================================================================
    ''' <summary>ADJ���݉��������� ###009</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub BtnADJ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnADJ.Click
        Dim strMSG As String

        Try
            If (BtnADJ.Text = "ADJ OFF") Then
                gbChkboxHalt = True
                BtnADJ.Text = "ADJ ON"
                BtnADJ.BackColor = System.Drawing.Color.Yellow
            Else
                gbChkboxHalt = False
                BtnADJ.Text = "ADJ OFF"
                BtnADJ.BackColor = System.Drawing.SystemColors.Control
            End If

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "FrmFineAdjust.BtnADJ_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "Ten Key On/Off�{�^������������"
    '''=========================================================================
    ''' <summary>Ten Key On/Off�{�^������������ ###057</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnTenKey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnTenKey.Click

        'Dim InpKey As UShort
        Dim strMSG As String

        Try
            Call SubBtnTenKey_Click()                                   ' ###139

            '' InpKey��HI SW�ȊO��OFF����' ###139
            'GetInpKey(InpKey)
            'If (InpKey And cBIT_HI) Then                                ' HI SW ON ?
            '    InpKey = cBIT_HI
            'Else
            '    InpKey = 0
            'End If
            'PutInpKey(InpKey)

            '' Ten Key On/Off�{�^���ݒ�
            'If (BtnTenKey.Text = "Ten Key Off") Then
            '    gbTenKeyFlg = True
            '    BtnTenKey.Text = "Ten Key On"
            '    BtnTenKey.BackColor = System.Drawing.Color.Pink
            'Else
            '    gbTenKeyFlg = False
            '    BtnTenKey.Text = "Ten Key Off"
            '    BtnTenKey.BackColor = System.Drawing.SystemColors.Control
            'End If

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "FrmFineAdjust.BtnTenKey_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "Ten Key On/Off�{�^������������"
    '''=========================================================================
    ''' <summary>Ten Key On/Off�{�^������������ ###139</summary>
    '''=========================================================================
    Private Sub SubBtnTenKey_Click()

        Dim InpKey As UShort
        Dim strMSG As String

        Try
            ' InpKey��HI SW�ȊO��OFF����
            GetInpKey(InpKey)
            If (InpKey And cBIT_HI) Then                                ' HI SW ON ?
                InpKey = cBIT_HI
            Else
                InpKey = 0
            End If
            PutInpKey(InpKey)

            ' Ten Key On/Off�{�^���ݒ�
            If (BtnTenKey.Text = "Ten Key Off") Then
                gbTenKeyFlg = True
                BtnTenKey.Text = "Ten Key On"
                BtnTenKey.BackColor = System.Drawing.Color.Pink
            Else
                gbTenKeyFlg = False
                BtnTenKey.Text = "Ten Key Off"
                BtnTenKey.BackColor = System.Drawing.SystemColors.Control
            End If

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "FrmFineAdjust.SubBtnTenKey_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

    '----- ###237�� -----
#Region "LASER�{�^������������"
    '''=========================================================================
    ''' <summary>LASER�{�^������������</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnLaser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnLaser.Click

        Dim r As Integer
        Dim strMSG As String

        Try
            ' LASER�ˏo�\/�s�̐؂�ւ�
            If (BtnLaser.BackColor = System.Drawing.SystemColors.Control) Then
                ' LASER�ˏo�\�Ƃ���
                BtnLaser.BackColor = System.Drawing.Color.OrangeRed
            Else
                ' LASER�ˏo�s�Ƃ���
                BtnLaser.BackColor = System.Drawing.SystemColors.Control
                r = LASEROFF()
                m_LaserOnOffFlag = False
            End If

            ' �g���b�v�G���[������
        Catch ex As Exception
            strMSG = "frmFineAdjust.BtnLaser_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region
    '----- ###237�� -----

    '========================================================================================
    '   ���ʊ֐�
    '========================================================================================
#Region "�X�e�[�W�ړ�����"
    '''=========================================================================
    ''' <summary>�X�e�[�W�ړ�����</summary>
    ''' <param name="pltNo"></param>
    ''' <param name="blkNo"></param>
    ''' <returns></returns>
    '''=========================================================================
    Private Function MoveTargetStagePos(ByVal pltNo As Integer, ByVal blkNo As Integer) As Integer

        Dim intRet As Integer
        Dim nextStgX As Double
        Dim nextStgY As Double
        Dim dispPltX As Integer
        Dim dispPltY As Integer
        Dim dispBlkX As Integer
        Dim dispBlkY As Integer
        'Dim retBlkNoX As Integer
        'Dim retBlkNoY As Integer
        Dim dispCurStgGrpNoX As Integer
        Dim dispCurStgGrpNoY As Integer
        Dim dispCurBlkNoX As Integer
        Dim dispCurBlkNoY As Integer
        Dim dispCurPltNoX As Integer
        Dim dispCurPltNoY As Integer

        Try
            MoveTargetStagePos = MOVE_NEXT
            intRet = GetTargetStagePos(pltNo, blkNo, nextStgX, nextStgY, dispPltX, dispPltY, dispBlkX, dispBlkY)
            If intRet = BLOCK_END Then
                ' �������Ȃ��ŏI��
                MoveTargetStagePos = MOVE_NOT
                Exit Function
            ElseIf intRet = PLATE_BLOCK_END Then
                ' �������Ȃ��ŏI��
                MoveTargetStagePos = MOVE_NOT
                Exit Function
            End If

            '---------------------------------------------------------------------
            '   �\���p�e�|�W�V�����̔ԍ���ݒ�i�v���[�g/�X�e�[�W�O���[�v/�u���b�N�j
            '---------------------------------------------------------------------
            Dim bRet As Boolean
            bRet = GetDisplayPosInfo(dispBlkX, dispBlkY, _
                            dispCurStgGrpNoX, dispCurStgGrpNoY, dispCurBlkNoX, dispCurBlkNoY)

            '---------------------------------------------------------------------
            '   ���O�\��������̐ݒ�
            '---------------------------------------------------------------------
            dispCurPltNoX = dispPltX : dispCurPltNoY = dispPltY         '###056
            Call DisplayStartLog(dispCurPltNoX, dispCurPltNoY, _
                            dispCurStgGrpNoX, dispCurStgGrpNoY, dispCurBlkNoX, dispCurBlkNoY)
            '' '' '' �X�e�[�W�̓���
            ' '' ''intRet = Form1.System1.EX_START(gSysPrm, nextStgX + typPlateInfo.dblTableOffsetXDir + gfCorrectPosX, _
            ' '' ''                        nextStgY + typPlateInfo.dblTableOffsetYDir + gfCorrectPosY, 0)
        Catch ex As Exception
            Dim strMsg As String
            strMsg = "frmFineAdjust.btnTrimming_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMsg)
        End Try
    End Function
#End Region

#Region "���{�^��������/�񊈐���"
    '''=========================================================================
    ''' <summary>���{�^��������/�񊈐��� ###139</summary>
    ''' <param name="OnOff"></param>
    '''=========================================================================
    Private Sub SetBtnArowEnable(ByVal OnOff As Boolean)

        Dim strMSG As String

        Try
            ' ���{�^��������/�񊈐���
            BtnJOG_0.Enabled = OnOff
            BtnJOG_1.Enabled = OnOff
            BtnJOG_2.Enabled = OnOff
            BtnJOG_3.Enabled = OnOff
            BtnJOG_4.Enabled = OnOff
            BtnJOG_5.Enabled = OnOff
            BtnJOG_6.Enabled = OnOff
            BtnJOG_7.Enabled = OnOff
            BtnHI.Enabled = OnOff

            ' Ten Key�{�^��������/�񊈐���
            BtnTenKey.Enabled = OnOff

            ' Ten Key�{�^����On/Off�ɂ���
            If (OnOff = False) Then
                ' ���{�^���񊈐����Ȃ�Ten Key�{�^����Off�ɂ��ăe���L�[���͂�s�Ƃ���
                If (BtnTenKey.Text = "Ten Key On") Then
                    m_TenKeyFlg = True
                    Call SubBtnTenKey_Click()
                End If
            Else
                ' Ten Key�{�^����Off�ɂ����ꍇ��Ten Key�{�^����On�ɂ��ăe���L�[���͂��Ƃ���
                If (m_TenKeyFlg = True) Then
                    m_TenKeyFlg = False
                    Call SubBtnTenKey_Click()
                End If
            End If

        Catch ex As Exception
            strMSG = "frmFineAdjust.SetBtnArowEnable() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

    '===============================================================================
    '   Description  : �i�n�f�����ʏ���
    '
    '   Copyright(C) : TOWA LASERFRONT CORP. 2012
    '
    '===============================================================================
    '========================================================================================
    '   �{�^������������
    '========================================================================================
#Region "RESET�{�^������������"
    '''=========================================================================
    ''' <summary>RESET�{�^������������</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnRESET_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnRESET.Click
        mExit_flg = cFRS_ERR_RST                                        ' Return�l = Cancel(RESET��)  
        gbExitFlg = True                                                ' �I���t���OON
        Me.Close()
    End Sub
#End Region

#Region "HI�{�^������������"
    '''=========================================================================
    ''' <summary>HI�{�^������������</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnHI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnHI.Click
        Call SubBtnHI_Click(stJOG)
    End Sub
#End Region

#Region "���{�^���̃}�E�X�N���b�N������"
    '''=========================================================================
    ''' <summary>���{�^���̃}�E�X�N���b�N������</summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub BtnJOG_0_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_0.MouseDown
        Call SubBtnJOG_0_MouseDown()                                    ' +Y ON
    End Sub
    Private Sub BtnJOG_0_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_0.MouseUp
        Call SubBtnJOG_0_MouseUp()                                      ' +Y OFF
    End Sub

    Private Sub BtnJOG_1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_1.MouseDown
        Call SubBtnJOG_1_MouseDown()                                    ' -Y ON
    End Sub
    Private Sub BtnJOG_1_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_1.MouseUp
        Call SubBtnJOG_1_MouseUp()                                      ' -Y OFF
    End Sub

    Private Sub BtnJOG_2_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_2.MouseDown
        Call SubBtnJOG_2_MouseDown()                                    ' +X ON
    End Sub
    Private Sub BtnJOG_2_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_2.MouseUp
        Call SubBtnJOG_2_MouseUp()                                      ' +X OFF
    End Sub

    Private Sub BtnJOG_3_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_3.MouseDown
        Call SubBtnJOG_3_MouseDown()                                    ' -X ON
    End Sub
    Private Sub BtnJOG_3_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_3.MouseUp
        Call SubBtnJOG_3_MouseUp()                                      ' -X OFF
    End Sub

    Private Sub BtnJOG_4_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_4.MouseDown
        Call SubBtnJOG_4_MouseDown()                                    ' -X -Y ON
    End Sub
    Private Sub BtnJOG_4_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_4.MouseUp
        Call SubBtnJOG_4_MouseUp()                                      ' -X -Y OFF
    End Sub

    Private Sub BtnJOG_5_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_5.MouseDown
        Call SubBtnJOG_5_MouseDown()                                    ' +X -Y ON
    End Sub
    Private Sub BtnJOG_5_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_5.MouseUp
        Call SubBtnJOG_5_MouseUp()                                      ' +X -Y OFF
    End Sub

    Private Sub BtnJOG_6_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_6.MouseDown
        Call SubBtnJOG_6_MouseDown()                                    ' +X +Y ON
    End Sub
    Private Sub BtnJOG_6_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_6.MouseUp
        Call SubBtnJOG_6_MouseUp()                                      ' +X +Y OFF
    End Sub

    Private Sub BtnJOG_7_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_7.MouseDown
        Call SubBtnJOG_7_MouseDown()                                    ' -X +Y ON
    End Sub
    Private Sub BtnJOG_7_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_7.MouseUp
        Call SubBtnJOG_7_MouseUp()                                      ' -X +Y OFF
    End Sub
#End Region
    '----- ###219 -----
#Region "Z�{�^������������"
    '''=========================================================================
    '''<summary>RESET�{�^������������</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnZ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnZ.Click
        Call SubBtnZ_Click(stJOG)
    End Sub
#End Region
    '----- ###219 -----

    '========================================================================================
    '   �e���L�[���͏���
    '========================================================================================
#Region "�L�[�_�E��������"
    '''=========================================================================
    ''' <summary>�L�[�_�E��������</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub frmFineAdjust_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        JogKeyDown(e)

        'Dim KeyCode As Short = e.KeyCode

        ''----- ###237�� -----
        '' LASER�ˏo�\�Łu*�L�[�v�����Ȃ�LASER�ˏo����()
        'If (BtnLaser.BackColor = System.Drawing.Color.OrangeRed) And (KeyCode = System.Windows.Forms.Keys.Multiply) Then
        '    ' ���[�UON
        '    If (m_LaserOnOffFlag = False) Then
        '        Call LASERON()
        '        m_LaserOnOffFlag = True
        '        Console.WriteLine("frmFineAdjust_KeyDown() Laser On")
        '    End If
        'End If
        ''----- ###237�� -----

        '' �e���L�[���̓t���O��OFF�Ȃ�NOP ###057
        'If (gbTenKeyFlg = False) Then
        '    e.Handled = False           ' V6.1.3.0�H
        '    Exit Sub
        'End If

        '' �e���L�[�_�E���Ȃ�InpKey�Ƀe���L�[�R�[�h��ݒ肷��
        'Call Sub_10KeyDown(KeyCode)
        'If (KeyCode = System.Windows.Forms.Keys.NumPad5) Then       ' 5�� (KeyCode = 101(&H65)
        '    Call BtnHI_Click(sender, e)                             ' HI�{�^�� ON/OFF
        'End If
        ''Call Me.Focus()

    End Sub

    Public Sub JogKeyDown(ByVal e As KeyEventArgs) Implements ICommonMethods.JogKeyDown    'V2.2.0.0�@

        Dim KeyCode As Short = e.KeyCode

        '----- ###237�� -----
        ' LASER�ˏo�\�Łu*�L�[�v�����Ȃ�LASER�ˏo����()
        If (BtnLaser.BackColor = System.Drawing.Color.OrangeRed) And (KeyCode = System.Windows.Forms.Keys.Multiply) Then
            ' ���[�UON
            If (m_LaserOnOffFlag = False) Then
                Call LASERON()
                m_LaserOnOffFlag = True
                Console.WriteLine("frmFineAdjust_KeyDown() Laser On")
            End If
        End If
        '----- ###237�� -----

        ' �e���L�[���̓t���O��OFF�Ȃ�NOP ###057
        If (gbTenKeyFlg = False) Then
            e.Handled = False           ' V6.1.3.0�H
            Exit Sub
        End If

        ' �e���L�[�_�E���Ȃ�InpKey�Ƀe���L�[�R�[�h��ݒ肷��
        Call Sub_10KeyDown(KeyCode)
        If (KeyCode = System.Windows.Forms.Keys.NumPad5) Then       ' 5�� (KeyCode = 101(&H65)
            Call BtnHI_Click(BtnHI, e)                             ' HI�{�^�� ON/OFF
        End If
        'Call Me.Focus()



        ''V6.0.0.0�J        Dim KeyCode As Short = e.KeyCode
        'Dim KeyCode As Keys = e.KeyCode             'V6.0.0.0�J
        'Dim r As Integer

        ''----- ###237�� -----
        '' LASER�ˏo�\�Łu*�L�[�v�����Ȃ�LASER�ˏo����()
        'If (BtnLaser.BackColor = System.Drawing.Color.OrangeRed) And (KeyCode = System.Windows.Forms.Keys.Multiply) Then
        '    ' ���[�UON
        '    If (m_LaserOnOffFlag = False) Then
        '        ' DIG-SW�ݒ�
        '        Call Form1.SetMoveMode(CbDigSwL.SelectedIndex, CbDigSwH.SelectedIndex) 'V5.0.0.1�K

        '        ''V4.0.0.0-86
        '        r = GetLaserOffIO(False)
        '        If r = 1 Then
        '            Me.ShowInTaskbar = False 'V5.0.0.1�K
        '            Me.Activate()  'V5.0.0.1�K
        '            'frmFineAdjust_KeyUp(sender, e)

        '            Exit Sub
        '        End If
        '        ''V4.0.0.0-86
        '        Call LASERON()
        '        m_LaserOnOffFlag = True
        '        Console.WriteLine("frmFineAdjust_KeyDown() Laser On")
        '    End If
        'End If
        ''----- ###237�� -----

        '' �e���L�[���̓t���O��OFF�Ȃ�NOP ###057
        ''V7.0.0.0�N        If (gbTenKeyFlg = False) Then Exit Sub
        'If (gbTenKeyFlg = False) OrElse (False = _firstResistor) Then   'V7.0.0.0�N
        '    e.Handled = False           ' V6.1.3.0�H
        '    Exit Sub
        'End If

        '' �e���L�[�_�E���Ȃ�InpKey�Ƀe���L�[�R�[�h��ݒ肷��
        ''V6.0.0.0�K       'Call Sub_10KeyDown(KeyCode)
        'Sub_10KeyDown(KeyCode, stJOG)             'V6.0.0.0�K
        'If (KeyCode = System.Windows.Forms.Keys.NumPad5) Then       ' 5�� (KeyCode = 101(&H65)
        '    'Call BtnHI_Click(sender, e)                             ' HI�{�^�� ON/OFF
        '    Call BtnHI_Click(BtnHI, e)                              ' HI�{�^�� ON/OFF     'V6.0.0.0�I
        'End If
        ''Call Me.Focus()

    End Sub
#End Region

#Region "�L�[�A�b�v������"
    '''=========================================================================
    ''' <summary>�L�[�A�b�v������</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub frmFineAdjust_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp

        Me.JogKeyUp(e)                  'V6.0.0.0�J

        'Dim KeyCode As Short = e.KeyCode

        ''----- ###237�� -----
        '' LASER Off����
        'If (m_LaserOnOffFlag = True) Then
        '    Call LASEROFF()
        '    m_LaserOnOffFlag = False
        '    Console.WriteLine("frmFineAdjust_KeyUp() Laser Off")
        'End If
        ''----- ###237�� -----

        '' �e���L�[���̓t���O��OFF�Ȃ�NOP ###057
        'If (gbTenKeyFlg = False) Then Exit Sub

        '' �e���L�[�A�b�v�Ȃ�InpKey�̃e���L�[�R�[�h��OFF����
        'Call Sub_10KeyUp(KeyCode)
        ''Call Me.Focus()

    End Sub


    Public Sub JogKeyUp(ByVal e As KeyEventArgs) Implements ICommonMethods.JogKeyUp        'V2.2.0.0�@

        Dim KeyCode As Short = e.KeyCode

        '----- ###237�� -----
        ' LASER Off����
        If (m_LaserOnOffFlag = True) Then
            Call LASEROFF()
            m_LaserOnOffFlag = False
            Console.WriteLine("frmFineAdjust_KeyUp() Laser Off")
        End If
        '----- ###237�� -----

        ' �e���L�[���̓t���O��OFF�Ȃ�NOP ###057
        If (gbTenKeyFlg = False) Then Exit Sub

        ' �e���L�[�A�b�v�Ȃ�InpKey�̃e���L�[�R�[�h��OFF����
        Call Sub_10KeyUp(KeyCode)
        'Call Me.Focus()

        ''V6.0.0.0�J        Dim KeyCode As Short = e.KeyCode
        'Dim KeyCode As Keys = e.KeyCode             'V6.0.0.0�J

        ''----- ###237�� -----
        '' LASER Off����
        'If (m_LaserOnOffFlag = True) Then
        '    Call LASEROFF()
        '    m_LaserOnOffFlag = False
        '    Console.WriteLine("frmFineAdjust_KeyUp() Laser Off")
        'End If
        ''----- ###237�� -----

        '' �e���L�[���̓t���O��OFF�Ȃ�NOP ###057
        ''V6.0.1.0�B        If (gbTenKeyFlg = False) Then Exit Sub
        ''V7.0.0.0�N        If (False = gbTenKeyFlg) Then       'V6.0.1.0�B
        'If (False = gbTenKeyFlg) OrElse (False = _firstResistor) Then   'V7.0.0.0�N
        '    'V6.1.3.0�H
        '    If (giBpAdj_HALT = 0) Then
        '        Sub_10KeyUp(Keys.None, stJOG)   'V6.0.1.0�B
        '    End If
        '    'V6.1.3.0�H
        'Else
        '    ' �e���L�[�A�b�v�Ȃ�InpKey�̃e���L�[�R�[�h��OFF����
        '    'V6.0.0.0�K        Call Sub_10KeyUp(KeyCode)
        '    Sub_10KeyUp(KeyCode, stJOG)                   'V6.0.0.0�K
        '    'Call Me.Focus()
        'End If

    End Sub

#End Region

    '========================================================================================
    '   �g���b�N�o�[����
    '========================================================================================
#Region "�g���b�N�o�[�̃X���C�_�[�ړ��C�x���g"
    '''=========================================================================
    ''' <summary>�g���b�N�o�[�̃X���C�_�[�ړ��C�x���g</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub TBarLowPitch_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBarLowPitch.Scroll
        Call SetSliderPitch(IDX_PIT, TBarLowPitch, TBarHiPitch, TBarPause, LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval)
    End Sub

    Private Sub TBarHiPitch_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBarHiPitch.Scroll
        Call SetSliderPitch(IDX_HPT, TBarLowPitch, TBarHiPitch, TBarPause, LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval)
    End Sub

    Private Sub TBarPause_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBarPause.Scroll
        Call SetSliderPitch(IDX_PAU, TBarLowPitch, TBarHiPitch, TBarPause, LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval)
    End Sub

    ''' <summary>
    ''' �ꎞ��~��ʂŃL���v�`���[��ʂ��N���b�N�����Ƃ��ɁA���삷��A���Ȃ��̐ݒ�       'V2.2.0.0�C
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnClickEnable_Click(sender As Object, e As EventArgs) Handles BtnClickEnable.Click

        Try

            If BtnClickEnable.BackColor = SystemColors.Control Then
                BtnClickEnable.BackColor = Color.Aqua
                gbTenKeyFlg = True
            Else
                BtnClickEnable.BackColor = SystemColors.Control
                gbTenKeyFlg = False
            End If

        Catch ex As Exception

        End Try

    End Sub
#End Region

#Region "���C���������s"
    ''' <summary>���C���������s</summary>
    ''' <returns>���s����</returns>
    ''' <remarks>'V2.2.0.0�@</remarks>
    Public Function Execute() As Integer Implements ICommonMethods.Execute
        ' DO NOTHING
    End Function
#End Region

#Region "�J�����摜�N���b�N�ʒu���摜�Z���^�[�Ɉړ����鏈��"
    ''' <summary>�J�����摜�N���b�N�ʒu���摜�Z���^�[�Ɉړ����鏈��</summary>
    ''' <param name="distanceX">�摜�Z���^�[����̋���X</param>
    ''' <param name="distanceY">�摜�Z���^�[����̋���Y</param>
    ''' <remarks>'V6.0.0.0�J</remarks>
    Public Sub MoveToCenter(ByVal distanceX As Decimal, ByVal distanceY As Decimal) _
        Implements ICommonMethods.MoveToCenter

        ' �e���L�[���̓t���O��OFF�Ȃ�NOP 
        If (gbTenKeyFlg = False) Then
            Exit Sub
        End If

        UserModule.MoveToCenter(distanceX, distanceY, stJOG)

    End Sub

    ''' <summary>
    ''' ���[�_���\���{�^�� 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnLoaderInfo_Click(sender As Object, e As EventArgs) Handles BtnLoaderInfo.Click

        Try

            objLoaderInfo.Show()

        Catch ex As Exception

        End Try



    End Sub

    ''' <summary>
    ''' �����^�]�̓o�^����Ă���t�@�C����\�����A�I���A���s���A���������킩��悤�ɂ��� 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnAutoInfo_Click(sender As Object, e As EventArgs) Handles btnAutoInfo.Click

        Dim objForm As frmLotFileDisp

        Try

            objForm = New frmLotFileDisp

            objForm.Show(Me)



        Catch ex As Exception

        End Try
    End Sub
#End Region

End Class