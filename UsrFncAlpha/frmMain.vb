'==============================================================================
'   Description : ���C����ʏ���
'
'   Copyright(C): TOWA LASERFRONT CORP. 2018
'
'==============================================================================
Option Strict Off
Option Explicit On

Imports System.Text                     ' ###lstLog
Imports System.Globalization
Imports System.Threading
Imports Microsoft.Win32
Imports LaserFront.Trimmer
Imports System.Collections.Generic
Imports LaserFront.Trimmer.DefWin32Fnc
Imports UsrFunc.My.Resources
Imports DllPlcIf

Friend Class Form1
    Inherits System.Windows.Forms.Form

    '=========================================================================
    '   �ϐ���`
    '=========================================================================
#Region "�t�H�[�����O���[�o���ϐ�"
    '=====================================================================
    ' �t�H�[�����O���[�o���ϐ�
    '=====================================================================
    Private Const CUR_CRS_LINEX As Short = 8            ' �۽ײ�X�\���ʒu�̕␳�l
    Private Const CUR_CRS_LINEY As Short = 13           ' �۽ײ�Y�\���ʒu�̕␳�l

    Private gflgCmpEndProcess As Boolean                ' �I�����������t���O�iTrue=�I���������s�ς݁AFalse=�I���������s�ς݂łȂ��j
    Private gfclamp As Boolean                          ' �N�����vON/OFF
    Private pbVideoCapture As Boolean                   ' �r�f�I�L���v�`���[�J�n�t���O
    Public gPrevTrimMode As Short                       ' �f�W�^���r�v�ޔ���
    Public giTrimErr As Short                           ' ��ϰ �װ �׸� ���װ���͸���߸����OFF����ϓ��쒆OFF��۰�ް�ɑ��M���Ȃ�
    '                                                   ' B0 : �z���װ(EXIT)
    '                                                   ' B1 : ���̑��װ
    '                                                   ' B2 : �W�o�@�װь��o
    '                                                   ' B3 : ���ЯĤ���װ�����ѱ��
    '                                                   ' B4 : ����~
    '                                                   ' B5 : ������װ

    Private pbVideoInit As Boolean                      ' �r�f�IInit�t���O
    'Private Const cTEMPLATPATH As String = "C:\TRIM\VIDEO"  ' Video.OCX�p����ڰ�̧�ق̕ۑ��ꏊ
    Private Const WORK_DIR_PATH As String = "C:\TRIM"       ' ��Ɨp̫��ް
    Private gbChkboxHalt As Boolean = False             ' ADJ�{�^�����(ON=ADJ ON, OFF=ADJ OFF) ###009
    Private gbAdjOnStatus As Boolean = False            ' �`�c�i�{�^���ł̒�~��

#End Region

#Region "�J�����摜�N���b�N�ړ��֘A"     'V2.2.0.0�@
    Private _jogKeyDown As Action(Of KeyEventArgs) = Nothing
    Private _jogKeyUp As Action(Of KeyEventArgs) = Nothing
    ''' <summary><para>�\������JOG�𐧌䂷��KeyDown,KeyUp���̏��������C���t�H�[���ɁA</para>
    ''' <para>�J�����摜MouseClick���̏�����DllVideo�ɐݒ肷��</para></summary>
    ''' <param name="keyDown"></param>
    ''' <param name="keyUp"></param>
    ''' <param name="moveToCenter">�J�����摜�N���b�N�ʒu���摜�Z���^�[�Ɉړ����鏈��</param>
    Friend Sub SetActiveJogMethod(ByVal keyDown As Action(Of KeyEventArgs),
                                  ByVal keyUp As Action(Of KeyEventArgs),
                                  ByVal moveToCenter As Action(Of Decimal, Decimal))
        _jogKeyDown = keyDown
        _jogKeyUp = keyUp

        '�J�����摜�\��PictureBox�N���b�N�ʒu��JOG�o�R�ŉ摜�Z���^�[�Ɉړ�����
        VideoLibrary1.MoveToCenter = moveToCenter
    End Sub
#End Region

    '=========================================================================
    '   �t�H�[���̏�����/�I������
    '=========================================================================
#Region "�V���b�g�_�E������-�����I��"
    '''=========================================================================
    '''<summary>�V���b�g�_�E������-�����I��</summary>
    '''<param name="sender"></param> 
    '''<param name="e"></param> 
    '''=========================================================================
    Private Sub SystemEvents_SessionEnding(
            ByVal sender As Object,
            ByVal e As SessionEndingEventArgs)
        If e.Reason = SessionEndReasons.SystemShutdown Then
            Call AplicationForcedEnding()
        End If
    End Sub
#End Region

#Region "�t�H�[������������"
    '''=========================================================================
    '''<summary>�t�H�[������������</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub Form_Initialize_Renamed()

        Dim r As Short
        Dim strMSG As String

        Try
            '-----------------------------------------------------------------------
            '   ���d�N���h�~Mutex�n���h��
            '-----------------------------------------------------------------------
            If gmhUserPro.WaitOne(0, False) = False Then
                '' ���łɋN������Ă���ꍇ
                '   �����b�Z�[�W�{�b�N�X���r�s�`�q�s�{�^�����͑҂��Ȃǂ̏�ԂŁA���ɉ�邱�Ƃ�����̂ŁA�\���͂�߂�B
                'MessageBox.Show("Cannot run TKY's family.(Another Process of TKY's family is already running.", "Trimmer Program", _
                '                MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, _
                '                MessageBoxOptions.ServiceNotification, False)
                End
            End If




            ' �V���b�g�_�E���C�x���g�����֐�
            AddHandler SystemEvents.SessionEnding, AddressOf SystemEvents_SessionEnding

            ChDir(WORK_DIR_PATH)
            Timer1.Enabled = False                                      ' �Ď��^�C�}�[��~

            ' Intime����m�F
#If cOFFLINEcDEBUG = 0 Then
            r = ISALIVE_INTIME()
            If (r = ERR_INTIME_NOTMOVE) Then
                '�G���[���b�Z�[�W�̕\�� (System1.TrmMsgBox�͂����ł͎g�p�ł��Ȃ��ׁA�W�����b�Z�[�W�{�b�N�X)
                MessageBox.Show("Real-time control module has not loaded.", "Trimmer Program", MessageBoxButtons.OK,
                                MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.DefaultDesktopOnly, False)
                End                                                     ' �A�v���I�� 
            End If
#End If
            '----------------------------------------------------------------------------
            '   �t���O��������
            '----------------------------------------------------------------------------
            ' �t���O������
            gbInitialized = False
            pbVideoInit = False
            pbVideoCapture = False                                      ' �r�f�I�L���v�`���[�J�n�t���O
            pbLoadFlg = False                                           ' �f�[�^���[�h�ς݃t���O
            gflgResetStart = False                                      ' �������t���O
            gfclamp = False                                             ' �N�����vOFF
            giAppMode = APP_MODE_IDLE                                   ' ����Ӱ�� = �g���}���u�A�C�h����
            FlgGPIB = False                                                 ' GPIB������Flag
            FlgUpd = False                                              ' �f�[�^�X�V Flag
            giTrimErr = 0                                               ' ��ϰ �װ �׸ޏ�����
            fStartTrim = False                                          ' �X�^�[�gTRIM�t���O OFF
            gflgCmpEndProcess = False                                   ' �I�����������t���O
            '                                                           ' �f�B�W�^��SW������ 
            DGH = DGSW_HI_DISP                                          ' �f�B�W�^��SWH = �S�ĕ\��
            DGL = TRIM_MODE_ITTRFT                                      ' �f�B�W�^��SWL = �C�j�V�����e�X�g�{�g���~���O�{�t�@�C�i���e�X�g���s
            DGSW = DGH * 10 + DGL                                       ' �f�B�W�^��SW

            ' �\���̂̏�����
            Call Init_Struct()

            '----------------------------------------------------------------------------
            '   �g�p����n�b�w�̏����ݒ���s��
            '----------------------------------------------------------------------------
            Call Ocx_Initialize()                                       ' �������Ұ���ذ�ޑO�ōs��
            '                                                           ' �����READ�O��Form_Load()�ɐ��䂪�n��̂Œ���
            '----------------------------------------------------------------------------
            '   �V�X�e���ݒ�t�@�C�����[�h
            '   ���V�X�e���p�����[�^�̑��M��OcxSystem��SetOptionFlg()�ōs��
            '----------------------------------------------------------------------------
            gSysPrm.Initialize()
            Call DllSysPrmSysParam_definst.SetAppKind(KND_USER)
            Call DllSysPrmSysParam_definst.GetSystemParameter(gSysPrm)   ' �V�X�e���ݒ�t�@�C�����[�h
            Call PrepareMessages(gSysPrm.stTMN.giMsgTyp)                 ' ���b�Z�[�W�����ݒ菈��
            Call Me.System1.OperationLogDelete(gSysPrm)                  ' �Â����샍�O�t�@�C�����폜����
            Call Me.System1.OperationLogging(gSysPrm, MSG_OPLOG_START, "")
            Call Me.System1.SetSysParam(gSysPrm)                         ' OcxSystem�p�̃V�X�e���p�����[�^��ݒ肷��

            ' ���O��ʕ\���N���A��������V�X�p�����擾����
            gDspClsCount = GetPrivateProfileInt("SPECIALFUNCTION", "DISP_CLS_USR", 5, SYSPARAMPATH)
            If (gDspClsCount <= 0) Then gDspClsCount = 1
            gDspCounter = 0                                             ' ���O��ʕ\��������J�E���^

            ' EXTOUT LED����r�b�g(BIT4-7)���V�X�p�����ݒ肷��
            glLedBit = Val(GetPrivateProfileString_S("IO_CONTROL", "ILUM_BIT", SYSPARAMPATH, "16"))

            gGpibMultiMeterCount = GetPrivateProfileInt("SPECIALFUNCTION", "MULTIMETER_COUNT", 5, SYSPARAMPATH) ' �}���`���[�^��IT�Ƒ���̑���񐔍Ō�̒l���g�p����B
            bDebugLogOut = GetPrivateProfileInt("LOGGING", "LOG_DEBUG", 0, SYSPARAMPATH)                '�ʏ�̃f�o�b�O���O�̏o�͗L��       'V1.2.0.2
            bNgCutDebugLogOut = GetPrivateProfileInt("LOGGING", "LOG_DEBUG_NGCUT", 0, SYSPARAMPATH)     '�m�f�J�b�g�p�f�o�b�O���O�̏o�͗L�� 'V1.2.0.2
            bCutVariationDebugLogOut = GetPrivateProfileInt("LOGGING", "LOG_DEBUG_CUTVA", 0, SYSPARAMPATH)  '�m�f�J�b�g�p�f�o�b�O���O�̏o�͗L�� 'V2.1.0.0�@
            'V2.0.0.0�L��
            If Integer.Parse(GetPrivateProfileString_S("USER", "RELAY_BOARD", SYSPARAMPATH, "0")) = 2 Then
                bRelayBoard = True
            Else
                bRelayBoard = False
            End If
            'V2.0.0.0�L��
            ''V2.2.1.3�A��
            giAutoOperationDebugLogOut = GetPrivateProfileInt("LOGGING", "LOG_DEBUG_AUTOMODE", 0, SYSPARAMPATH)  '���b�g�����p�f�o�b�O���O�̏o�͗L�� 
            ''V2.2.1.3�A

            ''V2.0.0.0�A��
            If Integer.Parse(GetPrivateProfileString_S("USER", "POWER_ON_OFF_TRIM_MEAS", SYSPARAMPATH, "0")) <> 0 Then
                bPowerOnOffUse = True
            Else
                bPowerOnOffUse = False
            End If
            ''V2.0.0.0�A��
            ''V2.2.0.0�B��
            If Integer.Parse(GetPrivateProfileString_S("OPT_TEACH", "DISABLE_BLUE_CROSSLINE", SYSPARAMPATH, "0")) <> 0 Then
                giBlueCrossDisable = 1
            Else
                giBlueCrossDisable = 0
            End If
            ''V2.2.0.0�B��
            ''V2.2.0.0�C��
            If Integer.Parse(GetPrivateProfileString_S("SPECIALFUNCTION", "ADJ_MOUSECLICK_DISABLE", SYSPARAMPATH, "0")) <> 0 Then
                giMouseClickMove = 1
            Else
                giMouseClickMove = 0
            End If
            ''V2.2.0.0�C��
            ''V2.2.0.0�D��
            If Integer.Parse(GetPrivateProfileString_S("DEVICE_CONST", "LOADER_TYPE", SYSPARAMPATH, "0")) <> 0 Then
                giLoaderType = 1
                btnLoaderInfo.Visible = True
                Call COVERCHK_ONOFF(0)                         ' �u�Œ�J�o�[�J�`�F�b�N����v�ɂ���
                btnCycleStop.Visible = True         'V2.2.2.0�@
            Else
                giLoaderType = 0
                btnLoaderInfo.Visible = False
                btnCycleStop.Visible = False        'V2.2.2.0�@
            End If
            ''V2.2.0.0�D��


            ''V2.2.0.0�E��
            If Integer.Parse(GetPrivateProfileString_S("SPECIALFUNCTION", "CUT_STOP", SYSPARAMPATH, "0")) <> 0 Then
                giCutStop = 1
                Me.btnCutStop.Visible = True
            Else
                giCutStop = 0
                Me.btnCutStop.Visible = False
            End If
            ''V2.2.0.0�E��

            ''V2.2.0.0�F��
            ' �T�C�N����~�@�\
            If Integer.Parse(GetPrivateProfileString_S("SPECIALFUNCTION", "CYCLE_STOP", SYSPARAMPATH, "0")) <> 0 Then
                giClcleStop = 1
                Me.btnCycleStop.Visible = True
            Else
                giClcleStop = 0
                Me.btnCycleStop.Visible = False
            End If
            ''V2.2.0.0�F��


            'V2.2.2.0�@��
            ' �����J�����ԍ��擾
            INTERNAL_CAMERA = Integer.Parse(GetPrivateProfileString_S("OPT_VIDEO", "INTERNAL_CAMERA_PORT", SYSPARAMPATH, "0"))
            ' �O���J�����ԍ��擾
            EXTERNAL_CAMERA = Integer.Parse(GetPrivateProfileString_S("OPT_VIDEO", "EXTERNAL_CAMERA_PORT", SYSPARAMPATH, "1"))
            'V2.2.2.0�@��

            '----------------------------------------------------------------------------
            ' ���[�U�[��`�ϐ��̏���������
            '----------------------------------------------------------------------------
            Call Set_UserForm(Z0)                                       ' ���C�����
            Call Me.System1.SetSignalTower(0, &HFFFFS)                  ' ������ܰ������(On,Off)
            Call GetFncDefParameter()                                   ' �@�\�I���`�e�[�u���ݒ�
            Call GetPasFuncDefParameter()                               ' �p�X���[�h��`�e�[�u���ݒ�

            '----------------------------------------------------------------------------
            '   �I�u�W�F�N�g�ݒ�
            '----------------------------------------------------------------------------
            ObjGpib = New GpibMaster                                    ' �f�o�h�a�ʐM�p�I�u�W�F�N�g
            frmAutoObj = New FormDataSelect(Me)                             ' �����^�]������޼ު��

            '----------------------------------------------------------------------------
            '   ��ʕ\�����ڂ�ݒ肵��ʂɕ\������
            '----------------------------------------------------------------------------
            gPrevTrimMode = -1

            ' form�֘A(Form Load������)
            Me.Picture1.Top = gSysPrm.stDVR.giCrossLineX + 59            ' CLOSS LINE X(����)
            Me.Picture2.Left = gSysPrm.stDVR.giCrossLineY + 8            ' CLOSS LINE Y(�c��)

            ' OK/NG�\���ر
            Call Disp_Result(0, 0)

            ' ۸މ�ʊg��\���ؑ�(��߼�݋@�\)
            If (gSysPrm.stSPF.giDispCh = 0) Then                         ' �g��\�����Ȃ� ?
                cmdExpansion.Visible = False
            Else
                cmdExpansion.Visible = True
            End If

            Dim ctxStr As String
            If (0 = gSysPrm.stTMN.giMsgTyp) Then
                ctxStr = "�R�s�[ (&C)"
            Else
                ctxStr = "Copy (&C)"
            End If

            Me.ctxMenuLstBox.Items.Add(
                ctxStr, Nothing, New EventHandler(AddressOf lstLog_Copy))       ' ###lstLog
            Me.lstLog.Items.Add(" ")
            Me.txtlog.Text = " "            ''V2.2.0.0�P��

            ' ���[�U�p���[�����֘A���ڂ̐ݒ�(���{��/�p��)
            SetLaserItems()

            ' �N�����v/�z��OFF
            r = Me.System1.ClampVacume_Ctrl(gSysPrm, 0, giAppMode, giTrimErr, False)
            If (r <> cFRS_NORMAL) Then                                  ' �G���[ ?
                Call AplicationForcedEnding()                           ' ��ċ����I������
                End                                                     ' �A�v�������I��
                Return
            End If

            ''V2.2.0.0�D��
            ' TLF�����[�_�̏ꍇ�����^�]�؂�ւ����o�͂���
            If giLoaderType = 1 Then
                Call Me.System1.Z_ATLDSET(0, clsLoaderIf.LOUT_AUTO Or clsLoaderIf.LOUT_REDY)                    'V1.2.0.0�C ���[�_�[�o��(ON=����,OFF=�Ȃ�)
            End If
            ''V2.2.0.0�D��

            ''V2.2.0.028 ��
            giTablePosUpd = Int32.Parse(GetPrivateProfileString_S("OPT_VIDEO", "TABLE_POS_UPDATE", "C:\TRIM\tky.ini", "0"))
            ''V2.2.0.028 ��

            giRecogPointCorrLine = Int16.Parse(GetPrivateProfileString_S("OPT_VIDEO", "CUTPOSCORR_BASELINE", "C:\TRIM\tky.ini", "0"))   ' V2.2.1.2�@

            'V2.2.0.030��
            giLaserOffMode = 0
            If (gSysPrm.stRAT.giOsc_Res = OSCILLATOR_SP) Then
                btnLaserOff.Visible = True
                btnLaserOff.Enabled = True
            Else
                btnLaserOff.Visible = False
                btnLaserOff.Enabled = False
            End If
            'V2.2.0.030��

            '---------------------------------------------------------------------------
            '   �N����̍ŏ��̌��o��۰�ގ���Ӱ��/���쒆�̏ꍇ�́A��~�ɐؑւ���悤�m�F����
            '---------------------------------------------------------------------------
            ' ���[�_����
            giHostMode = cHOSTcMODEcMANUAL                              ' ۰��Ӱ�� = �蓮Ӱ��
            gbHostConnected = False                                     ' �z�X�g�ڑ���� = ���ڑ�(۰�ޖ�)
            giHostRun = 0                                               ' ۰�ޒ�~��
            Call Me.System1.ReadHostCommand(gSysPrm, giHostMode, gbHostConnected, giHostRun, giAppMode, pbLoadFlg)

            ' �N����۰�ގ���Ӱ��/���쒆����
            r = Me.System1.Form_Reset(cGMODE_LDR_CHK, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
            If (r <> cFRS_NORMAL) Then                                  ' �G���[(����~) ?
                'Call AplicationForcedEnding()                           ' ��ċ����I������
                End                                                     ' �A�v�������I��
                Return
            End If

            stCounter.LotCounter = 0                            ' ���b�g�J�E���^�[������

            ' �N���X���C���␳�̏�����
            ' ObjCrossLine.CrossLineParamINitial(Me.Picture2, Me.Picture1, Me.CrosLineX, Me.CrosLineY, 0.0, 0.0)
            ObjCrossLine.CrossLineParamINitial(AddressOf VideoLibrary1.GetCrossLineCenter,
                                               AddressOf VideoLibrary1.SetCorrCrossVisible,
                                               AddressOf VideoLibrary1.SetCorrCrossCenter,
                                               0.0, 0.0)
            Me.CrosLineX.BringToFront()
            Me.CrosLineY.BringToFront()

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "Form_Initialize() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�n�b�w�̏����ݒ�"
    '''=========================================================================
    '''<summary>�g�p����n�b�w�̏����ݒ���s��</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub Ocx_Initialize()

        Dim strMSG As String

        Try
            Dim i As Short
            Dim r As Short
            Dim onoff(cMAXOptFlgNUM) As Short           ' ���߲ٵ�߼��(�ő吔)

            '---------------------------------------------------------------------------
            '   OCX�p�I�u�W�F�N�g��ݒ肷��
            '---------------------------------------------------------------------------
            ObjMain = Me                                ' Form1�N���X
            ObjSys = System1                            ' OcxSystem.ocx        '������
            ObjUtl = Utility1                           ' OcxUtility.ocx
            ObjHlp = HelpVersion1                       ' OcxAbout.ocx
            ObjPas = Password1                          ' OcxPassword.ocx
            ObjMTC = ManualTeach1                       ' OcxManualTeach.ocx
            ObjTch = Teaching1                          ' Teach.ocx
            ObjPrb = Probe1                             ' Probe.ocx
            ObjVdo = VideoLibrary1                      ' Video.ocx
            ObjLoader = New clsLoaderIf()                    ' Loader�N���X     'V2.2.0.0�D
            '@@888 ObjfrmResetLoader = New frmResetLoader()    ' Loader���Z�b�g�t�H�[�� 'V2.2.0.0�D
            ObjPlcIf = New DllPlcIf.DllMelsecPLCIf()
            objLoaderInfo = New frmLoaderInfo()         ' Loader�֌W���\��'V2.2.0.0�D
            ObjSys.frmResetLoaderInitial()              'V2.2.0.0�D

            '---------------------------------------------------------------------------
            '   OcxSystem.ocx�p�̏����ݒ菈�����s��
            '---------------------------------------------------------------------------
            ' OcxSystem�p�̃I�u�W�F�N�g��ݒ肷��
            For i = 0 To 31
                ObjMON(i) = HostSignal(i)
            Next i
            'HostSignal(0).BackColor = Color.Black

            'Call ObjSys.SetOcxUtilityObject(Utility1)  ' OcxUtility.ocx
            Call System1.SetOcxUtilityObject(ObjUtl)    ' OcxUtility.ocx
            'r = System1.SetMainObject_EX(txtLog, ObjMON)       ' Main��޼ު��
            r = System1.SetMainObject_EX()              ' Main��޼ު��
            Call System1.SetSystemObject(System1)       ' System.ocx
            ' �e���W���[���̃��\�b�h��ݒ肷��(OcxSystem�p)
            gparModules = New MainModules               ' �e�����\�b�h�ďo���I�u�W�F�N�g
            Call System1.SetMainObject(gparModules)

            VideoLibrary1.SetMainObject(gparModules)    ' �e���W���[���̃��\�b�h��ݒ肷��B

            ObjVdo.SetCrossLineObject(gparModules)      ' �N���X���C���\���p�I�u�W�F�N�g 'V2.2.1.2�@


            ' ���߲ٵ�߼�݂�ݒ肷��
#If cOFFLINEcDEBUG = 0 Then                             ' ���ޯ��Ӱ�ނłȂ� ?
            onoff(0) = 0                                ' OffLine���ޯ���׸�OFF
            Call DebugMode(0, 0)                        ' DllTrimFunc.dll�ޯ���׸�ON
#Else
            onoff(0) = 1                                ' OffLine���ޯ���׸�ON
            Call DebugMode(1, 0)                        ' DllTrimFunc.dll�ޯ���׸�OFF
#End If

#If cIOcMONITORcENABLED = 0 Then
            onoff(1) = 0                                ' IO����\��(0=�\�����Ȃ�, 1=�\������)
#Else
            onoff(1) = 1
#End If
            ' ���߲ٵ�߼�݂�ݒ肵�V�X�e���p�����[�^��INtime���֑��M����
            r = Me.System1.SetOptionFlg(cMAXOptFlgNUM, onoff)
            If (r <> cFRS_NORMAL) Then
                strMSG = "Me.System1.SetOptionFlg Error (r = " & r.ToString("0") & ")"
                Call MsgBox(strMSG, MsgBoxStyle.OkOnly)
                Call AplicationForcedEnding()                   ' ��ċ����I������
                End                                             ' �A�v�������I��
                Return
            End If

            '---------------------------------------------------------------------------
            '   OcxAbout.ocx�p�̏����ݒ菈�����s��
            '---------------------------------------------------------------------------
            Call HelpVersion1.SetOcxUtilityObject(Utility1) ' OcxUtility.ocx

            '---------------------------------------------------------------------------
            '   OcxPassword.ocx�p�̏����ݒ菈�����s��
            '---------------------------------------------------------------------------
            Call Password1.SetOcxUtilityObject(Utility1)    ' OcxUtility.ocx

            '---------------------------------------------------------------------------
            '   OcxManualTeach.ocx�p�̏����ݒ菈�����s��
            '---------------------------------------------------------------------------
            Call ManualTeach1.SetOcxUtilityObject(Utility1) ' OcxUtility1.ocx
            Call ManualTeach1.SetSystemObject(System1)      ' System.ocx

            '---------------------------------------------------------------------------
            '   DllgSysPrm.dll�p�̏����ݒ菈�����s��
            '---------------------------------------------------------------------------
            Call DllSysPrmSysParam_definst.SetOcxUtilityObjectForSysprm(Utility1)

            '---------------------------------------------------------------------------
            '   Teach.ocx�p�̏����ݒ菈�����s��
            '---------------------------------------------------------------------------
            Call Teaching1.SetOcxUtilityObject(Utility1)    ' OcxUtility1.ocx
            Call Teaching1.SetSystemObject(System1)         ' System.ocx

            '---------------------------------------------------------------------------
            '   Probe.ocx�p�̏����ݒ菈�����s��
            '---------------------------------------------------------------------------
            Call Probe1.SetOcxUtilityObject(Utility1)       ' OcxUtility1.ocx
            Call Probe1.SetSystemObject(System1)            ' System.ocx

            '---------------------------------------------------------------------------
            '   Video.ocx�p�̏����ݒ菈�����s��
            '---------------------------------------------------------------------------
            Call VideoLibrary1.SetOcxUtilityObject(Utility1) ' OcxUtility1.ocx
            Call VideoLibrary1.SetSystemObject(System1)      ' System.ocx
            Exit Sub

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "Ocx_Initialize() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try

    End Sub
#End Region

#Region "�t�H�[�����[�h���̏���"
    '''=========================================================================
    ''' <summary>�t�H�[�����[�h���̏���</summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Dim i As Short
        Dim r As Short
        Dim dispSize As System.Drawing.Size
        Dim dispPos As System.Drawing.Point
        Dim strMSG As String                                            ' ү���ޕҏW��

        Try
            AddHandler CbDigSwL.MouseWheel, AddressOf CbDigSwL_MouseWheel   'V2.0.0.0�E

            Me.Visible = False                                          ' ��ʔ�\��

            Me.AutoRunnningDisp.BackColor = System.Drawing.Color.Yellow  ' �w�i�F = ���F
            Me.AutoRunnningDisp.Text = "�����^�]������"

            ' I/O����\��
#If cIOcMONITORcENABLED = 0 Then
            For i = 0 To 31
                HostSignal(i).Visible = False                           ' I/O�����\��
            Next
            Label2.Visible = False                                      ' H��\��
            Label3.Visible = False                                      ' L��\��
            Label4.Visible = False                                      ' L��\��
            Label5.Visible = False                                      ' H��\��
#End If

            ' ���ޯ�ޗpνĺ���ޔ�\��/�\��
#If (cIOcHostComandcENABLED = 0) Then
            For i = 0 To 8
                DEBUG_HST_CMD(i).Visible = False                        ' ���ޯ�ޗpνĺ���ޔ�\��
            Next
#Else
		For i = 0 To 8
            DEBUG_HST_CMD(i).Visible = True                             ' ���ޯ�ޗpνĺ���ޕ\��
		Next
#End If
            ' ��ʕ\���ʒu�̐ݒ�
            dispPos.X = 0
            dispPos.Y = 0
            Me.Location = dispPos

            ' ��ʃT�C�Y�̐ݒ�
            dispSize.Height = 1024
            dispSize.Width = 1280
            Me.Size = dispSize

            '�{�^���\����
            Call SetButtonImage()                                       ' �t�H�[���̃{�^�����̐ݒ�(���{��/�p��)
            Call Btn_Enb_OnOff(2)                                       ' �{�^�����̕\��/��\��
            Call Btn_Enb_OnOff(1)                                       ' �{�^��������/�񊈐���
            Call Disp_frmInfo(COUNTER.COUNTUP, COUNTER.INITIAL_DISP)    ' �g���~���O���ʕ\��(���ݸޑO(�Sܰ�))

            ' Ocx���ޯ��Ӱ�ސݒ�
#If cOFFLINEcDEBUG Then
            VideoLibrary1.cOFFLINEcDEBUG = &H3141S
            Teaching1.cOFFLINEcDEBUG = &H3141S
            Probe1.cOFFLINEcDEBUG = &H3141S
            'trimmer.cOFFLINEcDEBUG = &H3141
            ctl_LaserTeach1.cOFFLINEcDEBUG = &H3141S
            Me.AutoScroll = True
#End If

            ' Video.ocx��DbgOn/Off���݂̗L��/�����w��(�f�o�b�O�p)
            '�@�f�o�b�O���ϐ����e��\�������邽��
#If cDBGRdraw Then                                                      ' Video.ocx��DbgOn/Off���ݗL���Ƃ��� ?
		VideoLibrary1.cDBGRdraw = &H3142
#End If
            Text2.Text = ""

            ' �R���g���[�����\���ɂ���
            Probe1.Visible = False
            Teaching1.Visible = False
            HelpVersion1.Visible = False

            ' �v���[�u�ʒu���킹�̃R���g���[���̕\���ʒu���w�肷��
            Probe1.Left = Text2.Left
            Probe1.Top = Text2.Top

            ' �e�B�[�`���O�̃R���g���[���̕\���ʒu���w�肷��
            Teaching1.Left = Text2.Left
            Teaching1.Top = Text2.Top
            Call Z_PRINT("")

            UserSub.LaserCalibrationModeLoad()                          'V2.1.0.0�A ���[�U�p���[���j�^�����O���[�h�擾�{�^���\��
            UserSub.LaserCalibrationSet(POWER_CHECK_LOT)                'V2.1.0.0�A ���[�U�p���[���j�^�����O���s�L���ݒ�

            '---------------------------------------------------------------------------
            '   ���u����������
            '---------------------------------------------------------------------------
            Call Me.Initialize_VideoLib()                               ' �r�f�I���C�u����������
            Call Me.VideoLibrary1.VideoStop()                           ' ���_���A�����ŕ\�����t���[�Y�\�������邽�߈�U��~
            '-------------------------------------------------------------------
            '   ���_���A����
            '-------------------------------------------------------------------
            Call Me.Initialize_TrimMachine()                            ' ���_���A������FL�ւ̏������t�@�C�����t
            Me.Visible = True                                           ' ��ʕ\��
            Me.Refresh()                                                'V2.0.0.3�@
            '-------------------------------------------------------------------
            '   �f�[�^���[�h
            '-------------------------------------------------------------------
            Call GetLotInf()                                            ' INI�t�@�C���ۑ����̃��b�g��񃍁[�h

            'V2.2.0.0�O�� 
            stMultiBlock.gMultiBlock = 0
            stMultiBlock.Initialize()
            For i = 0 To 5
                stMultiBlock.BLOCK_DATA(i).DataNo = i + 1           ' DataNo
                stMultiBlock.BLOCK_DATA(i).Initialize()
                stMultiBlock.BLOCK_DATA(i).gBlockCnt = 0            ' �u���b�N��
            Next
            ''V2.2.0.0�O��

            r = UserVal()                                               ' �f�[�^�����ݒ�
            If (r = 1) Then                                             ' �f�[�^���[�h�G���[ ?
                pbLoadFlg = False                                       ' �f�[�^���[�h�σt���O = False
                strMSG = "Data load Error : " & gsDataFileName & vbCrLf
                Call Z_PRINT(strMSG)
            ElseIf (r = 2) Then                                         ' �V�X�e���ϐ��ݒ�G���[ ?
                Call Me.System1.TrmMsgBox(gSysPrm, "�V�X�e���ϐ��ݒ�G���[!!(Aplication End)", MsgBoxStyle.OkOnly, My.Application.Info.Title)
                Call AplicationForcedEnding()                           ' ��ċ����I������
                End                                                     ' �A�v�������I��
                Return
            Else
                pbLoadFlg = True                                        ' �f�[�^���[�h�σt���O = True
                strMSG = "Data loaded : " & gsDataFileName & vbCrLf
                Call Z_PRINT(strMSG)
            End If
            If (pbLoadFlg = True) Then
                LblDataFileName.Text = gsDataFileName
            Else
                LblDataFileName.Text = ""
            End If

            Call UserSub.SetStartCheckStatus(True)      ' �ݒ��ʂ̊m�F�L����

            UserBas.TrimmingDataChange()        ' ###1041�@

            If UserSub.IsTrimType1() Or UserSub.IsTrimType4() Then       ' ###1033 ���x�Z���T�[�̎��ȊO�́A�펞�_��'V2.0.0.0�@sTrimType4()�ǉ� 
                UserBas.BackLight_Off()         ' ###1033
            Else                                ' ###1033
                UserBas.BackLight_On()          ' ###1033
            End If                              ' ###1033
            '###1040�E            Call SetATTRateToScreen(True)       ' ###1040�B �g���~���O�f�[�^�ł̂`�s�s�������̐ݒ�

            '-----------------------------------------------------------------------
            '   FL���։��H�����𑗐M����(FL���ŉ��H�����t�@�C��������ꍇ)
            '-----------------------------------------------------------------------
            Call SendFlParam(gsDataFileName)

            ' �����J�����ɐ؂�ւ���
            'V2.2.2.0�@��
            ' �����J�����ԍ��擾
            'V2.2.2.0�@�@Call Me.VideoLibrary1.ChangeCamera(0)
            Call Me.VideoLibrary1.ChangeCamera(INTERNAL_CAMERA)
            'V2.2.2.0�@��
            Call Me.VideoLibrary1.VideoStart()

            ' �R���\�[���̃��b�`����
            Call ZCONRST()

            ' �����v����
            Call Me.System1.sLampOnOff(LAMP_START, True)                ' START����ON
            Call Me.System1.sLampOnOff(LAMP_RESET, True)                ' RESET����ON

            ' �摜�\���v���O�������N������
            'V2.2.0.0�@ Execute_GazouProc(ObjGazou, DISPGAZOU_PATH, DISPGAZOU_WRK, INTERNAL_CAMERA)

            '-----------------------------------------------------------------------
            '   �Ď��^�C�}�[�J�n
            '-----------------------------------------------------------------------
            giAppMode = APP_MODE_IDLE                                   ' ����Ӱ�� = �g���}���u�A�C�h����
            Timer1.Interval = 10                                        ' �Ď��^�C�}�[�l(msec)
            Timer1.Enabled = True                                       ' �Ď��^�C�}�[�J�n

            gObjFrmDistribute = New frmDistribution                     ' ���z�}�f�[�^�I�u�W�F�N�g���� 'V2.0.0.0�H

            'V2.2.0.0�O��
            ' ������R�l���s���̌��ʊi�[������
            For rn As Integer = 0 To MAX_RES_USER - 1
                stToTalDataMulti(rn).Initialize()
            Next rn
            'V2.2.0.0�O��

            gObjFrmDistribute.ClearCounter()                            ' ���z�}�f�[�^������           'V2.0.0.0�H

            'V1.2.0.2            If gSysPrm.stLOG.giLoggingMode = 1 Then                     ' �f�o�b�O���O�̏o�͗L��
            'V1.2.0.2               bDebugLogOut = True
            'V1.2.0.2            Else
            'V1.2.0.2               bDebugLogOut = False
            'V1.2.0.2            End If

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "Form1_Load() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�t�H�[���̃{�^�����̐ݒ�(���{��/�p��)"
    '''=========================================================================
    ''' <summary>�t�H�[���̃{�^�����̐ݒ�(���{��/�p��)</summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub SetButtonImage()

        Dim strMSG As String

        Try
            ' �f�B�W�^���X�C�b�`�̐ݒ�(���{�� / �p��)
            SetDigSwImage()

            ' �{�^�����̐ݒ�(���{��/�p��)
            If (gSysPrm.stTMN.giMsgTyp = 0) Then
                '-------------------------------------------------------------------
                '   ���{��ݒ�
                '-------------------------------------------------------------------
                cmdLotInfo.Text = "�f�[�^�ݒ�" & vbCrLf & "(F1)"
                cmdLoad.Text = "���[�h" & vbCrLf & "(F2)"
                cmdSave.Text = "�Z�[�u" & vbCrLf & "(F3)"
                cmdEdit.Text = "�ҏW" & vbCrLf & "(F4)"
                cmdLaserTeach.Text = "���[�U"
                cmdProbeTeaching.Text = "�v���[�u" & vbCrLf & "(F6)"
                cmdTeaching.Text = "�e�B�[�`���O" & vbCrLf & "(F7)"
                cmdLotChg.Text = "�����^�]" & vbCrLf & "(F5)"
                cmdCutPosTeach.Text = "�J�b�g�ʒu�␳" & vbCrLf & "(F8)"
                BtnRECOG.Text = "�p�^�[���o�^" & vbCrLf & "(F9)"
                cmdExit.Text = "�I��"

            Else
                '-------------------------------------------------------------------
                '   �p��ݒ�
                '-------------------------------------------------------------------
                cmdLotInfo.Text = "DATA" & vbCrLf & "(F1)"
                cmdLoad.Text = "LOAD" & vbCrLf & "(F2)"
                cmdSave.Text = "SAVE" & vbCrLf & "(F3)"
                cmdEdit.Text = "EDIT" & vbCrLf & "(F4)"
                cmdLaserTeach.Text = "LASER"
                cmdProbeTeaching.Text = "PROBE" & vbCrLf & "(F6)"
                cmdTeaching.Text = "TEACH" & vbCrLf & "(F7)"
                cmdLotChg.Text = "LOT CHANGE" & vbCrLf & "(F5)"
                cmdCutPosTeach.Text = "CUT POS RECOG" & vbCrLf & "(F8)"
                BtnRECOG.Text = "RECOG" & vbCrLf & "(F9)"
                cmdExit.Text = "END"
            End If

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "SetButtonImage() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�f�B�W�^���X�C�b�`�̐ݒ�(���{��/�p��)"
    '''=========================================================================
    ''' <summary>�f�B�W�^���X�C�b�`�̐ݒ�(���{��/�p��)</summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub SetDigSwImage()

        Dim strMSG As String

        Try
            If (gSysPrm.stTMN.giMsgTyp = 0) Then
                '-------------------------------------------------------------------
                '   ���{��ݒ�
                '-------------------------------------------------------------------
                ' �f�B�W�^���X�C�b�`HI
                CbDigSwH.Items.Clear()
                CbDigSwH.Items.Add("�O�F�\���Ȃ�")
                CbDigSwH.Items.Add("�P�F�m�f�̂ݕ\��")
                CbDigSwH.Items.Add("�Q�F�S�ĕ\��")

                ' �f�B�W�^���X�C�b�`LO
                CbDigSwL.Items.Clear()
                CbDigSwL.Items.Add("�O�F�g���~���O")
                CbDigSwL.Items.Add("�P�F����")
                CbDigSwL.Items.Add("�Q�F�J�b�g���s")
                CbDigSwL.Items.Add("�R�F�X�e�b�v�����s�[�g")
                CbDigSwL.Items.Add("�S�F����}�[�L���O���[�h")  'V1.0.4.3�I
                CbDigSwL.Items.Add("�T�F�d�����[�h")            'V2.0.0.0�A
                CbDigSwL.Items.Add("�U�F����l�ϓ�����")        'V2.0.0.0�A

            Else
                '-------------------------------------------------------------------
                '   �p��ݒ�
                '-------------------------------------------------------------------
                ' �f�B�W�^���X�C�b�`HI
                CbDigSwH.Items.Clear()
                CbDigSwH.Items.Add("�O�FNo Display")
                CbDigSwH.Items.Add("�P�FDisplay only NG Logs")
                CbDigSwH.Items.Add("�Q�FDisplay All Logs")

                ' �f�B�W�^���X�C�b�`LO
                CbDigSwL.Items.Clear()
                CbDigSwL.Items.Add("0:Trimming")
                CbDigSwL.Items.Add("1:Measure")
                CbDigSwL.Items.Add("2:Cutting")
                CbDigSwL.Items.Add("3:Step And Repeat")
                CbDigSwL.Items.Add("4:Measure Marking")  'V1.0.4.3�I
                CbDigSwL.Items.Add("5:Power")            'V2.0.0.0�A
                CbDigSwL.Items.Add("6:Meas Variation")   'V2.0.0.0�A
            End If

            ' �f�B�W�^���X�C�b�`�̏����ݒ�
            LblDIGSW_HI.Visible = True
            LblDIGSW_LO.Visible = True

            CbDigSwH.Visible = True
            CbDigSwL.Visible = True
            CbDigSwH.SelectedIndex = DGH
            CbDigSwL.SelectedIndex = DGL

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "SetDigSwImage() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "���[�U�p���[�����֘A���ڂ̐ݒ�(���{��/�p��)"
    '''=========================================================================
    ''' <summary>
    ''' ���[�U�p���[�����������̐ݒ�(���{��/�p��)
    ''' </summary>
    ''' <param name="bMode">True:ATT�̐ݒ�AFalse:�p���[�������ATT�f�[�^�̕ۑ�</param>
    ''' <remarks>###1040�B�ŕ���</remarks>
    '''=========================================================================
    Public Function SetATTRateToScreen(ByVal bMode As Boolean) As Boolean   'V2.1.0.0�ASub����Function�֕ύX

        Dim strMSG As String
        Dim iRtn As Integer

        Try
            ' ���������V�X�p�����\������("������ = 99.9%")
            ' ��۰�ر��Ȱ��̐ݒ��OcxSystem�̌��_���A�����ōs����
            If (gSysPrm.stRMC.giRmCtrl2 >= 2 And
                gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then           ' ۰�ر��Ȱ�����L(RMCTRL2�Ή����L��) ?

                'V2.1.0.0�A                If pbLoadFlg And stLASER.iTrimAtt = 1 Then                                      ' ###1040�B
                If pbLoadFlg AndAlso (stLASER.iTrimAtt = 1 OrElse stLASER.iAttNo > 0) Then            'V2.1.0.0�A
                    If bMode Then                                                               ' ###1040�B ���[�U�p���[������
                        'V2.1.0.0�E�J�o�[�J�̃G���[                        Call ATTRESET() 'V2.1.0.0�E
                        iRtn = LATTSET(stLASER.iFixAtt, stLASER.dblRotAtt)                      ' ###1040�B
                        iRtn = ObjSys.EX_ZGETSRVSIGNAL(gSysPrm, iRtn, 0)                        ' ###1040�B
                        If (iRtn <> cFRS_NORMAL) Then                                           ' ###1040�B
                            Call Z_PRINT("���[�^���[�A�b�e�l�[�^�̐ݒ肪�ُ�I�����܂����B[" & CStr(iRtn) & "]")          ' ###1040�B
                            Return (False)                                                       'V2.1.0.0�A
                            'V2.1.0.0�A                            MsgBox("���[�^���[�A�b�e�l�[�^�̐ݒ肪�ُ�I�����܂����B[" & CStr(iRtn) & "]")          ' ###1040�B
                            'V2.1.0.0�A                            Exit Function                                                            ' ###1040�B
                        End If                                                                  ' ###1040�B
                    Else                                                                        ' ###1040�B
                        stLASER.dblRotPar = gSysPrm.stRAT.gfAttRate                             ' ###1040�B ������(%)
                        stLASER.dblRotAtt = gSysPrm.stRAT.giAttRot                              ' ###1040�B ���[�^���[�A�b�e�l�[�^�̉�]��(0-FFF)
                        stLASER.iFixAtt = gSysPrm.stRAT.giAttFix                                ' ###1040�B �Œ�A�b�e�l�[�^(0:OFF, 1:ON)
                    End If

                    If (gSysPrm.stTMN.giMsgTyp = 0) Then                                        ' ###1040�B
                        strMSG = "������ " + CDbl(stLASER.dblRotPar).ToString("##0.00") + " %"   ' ###1040�B
                    Else                                                                        ' ###1040�B
                        strMSG = "ATT. " + CDbl(stLASER.dblRotPar).ToString("##0.00") + " %"     ' ###1040�B
                    End If                                                                      ' ###1040�B
                Else                                                                            ' ###1040�B
                    If (gSysPrm.stTMN.giMsgTyp = 0) Then
                        strMSG = "������ " + CDbl(gSysPrm.stRAT.gfAttRate).ToString("##0.0") + " %"
                    Else
                        strMSG = "ATT. " + CDbl(gSysPrm.stRAT.gfAttRate).ToString("##0.0") + " %"
                    End If
                End If                                                                          ' ###1040�B
                Me.LblRotAtt.Text = strMSG                              ' �������\��
                Me.LblRotAtt.Visible = True
            Else
                Me.LblRotAtt.Visible = False
            End If

            Return (True)                                                                        'V2.1.0.0�A

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "SetATTRateToScreen() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (False)                                                       'V2.1.0.0�A
        End Try
    End Function
#End Region

#Region "���[�U�p���[�����֘A���ڂ̐ݒ�(���{��/�p��)"
    '''=========================================================================
    '''<summary>���[�U�p���[�����֘A���ڂ̐ݒ�(���{��/�p��)</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub SetLaserItems()

        Dim strMSG As String

        Try

            '###1040�E            Call SetATTRateToScreen(True)
            Call SetATTRateToScreen(False)   '###1040�E

            ' ����l���V�X�p�����\������
            ' ��۸��ыN������ڰ�ް��ܰ�ݒ�l�́u-----�v�\���Ƃ���
            If (gSysPrm.stRMC.giRmCtrl2 >= 3) And (gSysPrm.stRMC.giPMonHi = 1) Then ' RMCTRL2 >=3 �� ����l�\�� ?
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    strMSG = "���[�U�p���[�ݒ�l�@---- W"
                Else
                    strMSG = "Laser Power ---- W"
                End If
                LblMes.Text = strMSG                                    ' ����p���[[W]�̕\��
            Else
                LblMes.Visible = False                                  ' ����l��\��
            End If

            ' ��d���l��\������
            LblCur.Visible = False                                      ' ��d���l��\��
            If (gSysPrm.stTMN.giMsgTyp = 0) Then
                strMSG = "��d���l "
            Else
                strMSG = "Fixed Current Val "
            End If

            Select Case (gSysPrm.stSPF.giProcPower2)
                Case 0 ' �w��Ȃ�(�W��)
                    LblCur.Text = strMSG & "0.25A"                      ' "��d���l 0.25A"
                Case 1
                    LblCur.Text = strMSG & "1.00A"                      ' "��d���l 1.00A"
                Case 2
                    LblCur.Text = strMSG & "0.75A"                      ' "��d���l 0.75A"
                Case 3
                    LblCur.Text = strMSG & "0.50A"                      ' "��d���l 0.50A"
            End Select

            ' ���H�d�͐ݒ� = 4(��d��1A)�̎��ɕ\��
            If (gSysPrm.stSPF.giProcPower = 4) And (gSysPrm.stSPF.giProcPower2 <> 0) Then
                LblCur.Visible = True
            Else
                LblCur.Visible = False
            End If

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "SetLaserItems() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "MV10 �r�f�I���C�u��������������"
    '''=========================================================================
    ''' <summary>MV10 �r�f�I���C�u��������������</summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub Initialize_VideoLib()

        Dim lRet As Integer
        Dim s As String
        Dim strMSG As String

        Try
            '---------------------------------------------------------------------------
            '   �r�f�I���C�u����������������
            '---------------------------------------------------------------------------
            If pbVideoCapture = False Then
                pbVideoCapture = True
                'ChDir(My.Application.Info.DirectoryPath)
                ChDir(WORK_DIR_PATH)                                    ' MvcPt2.ini�̂���̫��ް����Ɨp̫��ް�Ƃ���

                If (gSysPrm.stDEV.giEXCAM = 0) Then                      ' �����J����?
                    VideoLibrary1.pp36_x = gSysPrm.stGRV.gfPixelX        ' �s�N�Z���lX(um)
                    VideoLibrary1.pp36_y = gSysPrm.stGRV.gfPixelY        ' �s�N�Z���lY(um)
                Else
                    VideoLibrary1.pp36_x = gSysPrm.stGRV.gfEXCAM_PixelX  ' �O������߸�ْlX(um)
                    VideoLibrary1.pp36_y = gSysPrm.stGRV.gfEXCAM_PixelY  ' �O������߸�ْlY(um)
                End If

                VideoLibrary1.OverLay = True
                lRet = VideoLibrary1.Init_Library()                     ' �r�f�I���C�u����������
                If (lRet <> 0) Then                                     ' Video.OCX�G���[ ?
                    Select Case lRet
                        Case cFRS_VIDEO_INI
                            s = "VIDEOLIB: Already initialized."
                        Case cFRS_VIDEO_PRP
                            s = "VIDEOLIB: Invalid property value."
                        Case cFRS_MVC_UTL
                            s = "VIDEOLIB: Error in MvcUtil"
                        Case cFRS_MVC_PT2
                            s = "VIDEOLIB: Error in MvcPt2"
                        Case cFRS_MVC_10
                            s = "VIDEOLIB: Error in Mvc10"
                        Case Else
                            s = "VIDEOLIB: Unexpected error 2"
                    End Select
                    Call Me.System1.TrmMsgBox(gSysPrm, s, MsgBoxStyle.OkOnly, My.Application.Info.Title)
                Else
                    ' ���C�u��������������
                    pbVideoInit = True
                End If

            End If

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "Initialize_VideoLib() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "���_���A������FL�ւ̏������t�@�C�����t"
    '''=========================================================================
    ''' <summary>���_���A������FL�ւ̏������t�@�C�����t</summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub Initialize_TrimMachine()

        Dim r As Short
        Dim strSetFileName As String = ""
        Dim strMSG As String

        Try
            '---------------------------------------------------------------------------
            '   ���_���A������FL�ւ̏������t�@�C�����t
            '---------------------------------------------------------------------------
            If (gflgResetStart = False) Then                            ' �����ݒ�ς݂łȂ� ?

                ' ����ڰ�̧�ق̕ۑ��ꏊ��"C:\TRIM"�ɐݒ肷��(VideoStart()��Ɏw�肷��)
                ' (��)�Ǘ�̧�فuPt2Template.xxx�v�͋N��̫��ނɍ쐬�����B
                r = Me.VideoLibrary1.SetTemplatePass(cTEMPLATPATH)

                Call InitFunction()                                     ' DllTrimFunc.dll������
                If (gflgResetStart = False) Then                        ' �����ݒ�ς݂łȂ� ?
                    If (giLoaderType <> 0) Then
                        ' �d�����b�N(�ω����E�����b�N)����������
                        r = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_ON)
                    End If

                    ' ���_���A
                    r = sResetStart()
                    If (r <> cFRS_NORMAL) Then                          ' �G���[ ?
                        If (r <> cFRS_ERR_RST) Then
                            r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE2)
                        End If

                        'V2.2.0.0�D��
                        If (giLoaderType <> 0) Then
                            Call Me.VideoLibrary1.ChangeCamera(0)
                            Call Me.VideoLibrary1.VideoStart()
                        End If

                        Call AplicationForcedEnding()                   ' ��ċ����I������
                        End                                             ' �A�v�������I��
                        Return
                    End If
                    gflgResetStart = True                               ' �����ݒ�ς�ON

                    If (giLoaderType <> 0) Then
                        ' �d�����b�N(�ω����E�����b�N)����������
                        r = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_OFF)
                    End If

                    'Call BackLight_Off()                                ' �o�b�N���C�g�k�d�c����n�e�e

                End If

#If cOSCILLATORcFLcUSE Then
                '-----------------------------------------------------------------------
                '   FL���։��H�����𑗐M����(FL��)
                '-----------------------------------------------------------------------
                If (gSysPrm.stRAT.giOsc_Res = OSCILLATOR_FL) Then
                    If (pbLoadFlg = True) Then                          ' �f�[�^���[�h�ς� ? 
                        strSetFileName = gsDataFileName                 ' ���[�h�����f�[�^�t�@�C���ɑΉ�������H�����t�@�C����
                    Else
                        strSetFileName = DEF_FLPRM_SETFILENAME          ' �f�t�H���g�̉��H�����t�@�C����
                    End If
                    ' FL�p���H�����t�@�C�������[�h����FL���։��H�����𑗐M����
                    r = SendTrimCondInfToFL(stCND, strSetFileName, strSetFileName)
                    If (r = SerialErrorCode.rRS_FLCND_XMLNONE) And (strSetFileName <> DEF_FLPRM_SETFILENAME) Then
                        ' ���[�h�����f�[�^�t�@�C���ɑΉ�����XML�t�@�C�������݂��Ȃ��ꍇ�̓f�t�H���g�̉��H�����𑗐M����
                        strSetFileName = DEF_FLPRM_SETFILENAME          ' �f�t�H���g�̉��H�����t�@�C����
                        r = SendTrimCondInfToFL(stCND, strSetFileName, strSetFileName)
                    End If
                    If (r <> SerialErrorCode.rRS_OK) Then
                        '"�e�k�ʐM�ُ�B�e�k�Ƃ̒ʐM�Ɏ��s���܂����B" + vbCrLf + "�e�k�Ɛ������ڑ��ł��Ă��邩�m�F���Ă��������B"
                        strMSG = MSG_150
                        Call MsgBox(strMSG, MsgBoxStyle.OkOnly, "")
                    End If
                End If
#End If
                ' �����}�K�W���̏����擾 
                r = ObjSys.Sub_GetNowProcessMgInfo(gisupplyMgNum, gisupplyMgStepNum, gistoreMgNum, gistoreMgStepNum)

            End If

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "Initialize_TrimMachine TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�t�H�[���A�����[�h���̏���"
    '''=========================================================================
    ''' <summary>�t�H�[���A�����[�h���̏���</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub Form1_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        Dim lRet As Integer
        Dim strMSG As String

        Try
            ' EXTBIT OFF
            Call EXTOUT1(0, &HFFFFS)                                    ' EXTBIT(On=0, Off=�S�r�b�g)
            Call EXTOUT2(0, &HFFFFS)                                    ' EXTBIT2(On=0, Off=�S�r�b�g)

            ' �g���}���f�B�M��OFF���M(���۰�ް)
            Call Sub_ATLDSET(COM_STS_TRM_STATE, COM_STS_TRM_READY)      ' ���[�_�[�o��(ON=�g���}���쒆, ,OFF=�g���}���f�B)

            Call IoMonitor(gdwATLDDATA, 1)

            '' �X���C�g�J�o�[�X�g�b�p�߂�(�S�J�ʒu��)
            'Call Me.System1.BigCover_Ctrl(gSysPrm, 1)

            ' ������ܰ����(On=0, Off=��ި, ���_���A��, �����^�]��)
            Call Me.System1.SetSignalTower(0, &HFFFFS)

            ' ���C�u�����I��
            If pbVideoInit = True Then
                lRet = VideoLibrary1.Close_Library
                If lRet <> 0 Then
                    Select Case lRet
                        Case cFRS_VIDEO_INI
                            Call Me.System1.TrmMsgBox(gSysPrm, "Video library: Not initialized.", MsgBoxStyle.OkOnly, My.Application.Info.Title)
                        Case Else
                            'MsgBox "�\�����ʃG���["
                            Call Me.System1.TrmMsgBox(gSysPrm, "Video library: Unexpected error.", MsgBoxStyle.OkOnly, My.Application.Info.Title)
                    End Select
                End If
            End If

            ' �f�o�h�a�I������
            ObjGpib.Gpib_Term(gDevId)

            ' ���샍�O�o��
            Call Me.System1.OperationLogging(gSysPrm, MSG_OPLOG_END, "FormClosed") ' "���[�U�v���O�����I��"

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "Form1_FormClosed() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

    '========================================================================================
    '   �R�}���h�{�^���������̏���
    '========================================================================================
#Region "�X�^�[�g�{�^��������(�f�o�b�O�p)"
    '''=========================================================================
    ''' <summary>�X�^�[�g�{�^��������</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub cmdStart_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdStart.Click

        Dim s As String
        Dim r As Short
        Dim strMSG As String

        Try
            ' ��������
            If pbLoadFlg = False Then
                s = MSG_DataNotLoad                         ' �ް���۰��
                Call Z_PRINT(s + vbCrLf)
                Call Beep()
                Exit Sub
            End If
            Timer1.Enabled = False                          '�^�C�}�[��~

            ' ���샍�O�o��(���ݸ�)
            Call Me.System1.OperationLogging(gSysPrm, MSG_OPLOG_TRIMST, "DEBUG")

            ' �g���~���O����
            r = User()

            ' �㏈��
            Timer1.Enabled = True                           ' �^�C�}�[�J�n
            Call ZCONRST()                                  ' �ݿ��SWׯ�����

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "cmdStart_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�I���{�^��������"
    '''=========================================================================
    ''' <summary>�I���{�^��������</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub cmdExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExit.Click

        Dim s As String
        Dim r As Short
        Dim strMSG As String

        Try
            ' �g���}���쒆�M��ON���M(���۰�ް)
            Call Sub_ATLDSET(COM_STS_TRM_STATE, 0)                      ' ���[�_�[�o��(ON=�g���}���쒆,OFF=�Ȃ�)

            ' �I���m�Fү���ނ�ݒ肷��
            giAppMode = APP_MODE_EXIT                                   ' ����Ӱ�ސݒ�

            Timer1.Enabled = False                                      ' �Ď��^�C�}�[��~
            Call ZCONRST()                                              ' ���b�`����

            '' �g���}���u�A�C�h�����ȊO�Ȃ�NOP
            'If giAppMode Then GoTo STP_END

            If pbLoadFlg = False Then                                   ' �f�[�^�����[�h ?
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    s = "           �I�����܂����H           "
                Else
                    s = "      Are you sure to quit ?      "
                End If

            ElseIf (FlgUpd = TriState.True) Then                        ' �f�[�^���[�h�ς݂ōX�V ?
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    s = "�ҏW���̃f�[�^������܂��B" & vbCrLf & "�A�v���P�[�V�������I�����Ă�낵���ł����H"
                Else
                    s = "  Please make sure to save the data to the disk before quit this program.  " & vbCrLf & "  Are you sure to quit?  "
                End If

            Else
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    s = "�A�v���P�[�V�������I�����Ă�낵���ł����H"
                Else
                    s = "      Are you sure to quit ?      "
                End If
            End If

            ' �I���m�Fү���ޕ\��
            r = Me.System1.TrmMsgBox(gSysPrm, s, MsgBoxStyle.OkCancel, "QUIT")
            If (r = cFRS_ERR_ADV) Then                                  ' OK(ADV��) ?
                '�@�\�t�g�����I������
                Call AplicationForcedEnding()
                End
                Exit Sub
            End If

STP_END:
            Call ZCONRST()                                              ' �ݿ�ٷ� ׯ�����
            giAppMode = APP_MODE_IDLE                                   ' ����Ӱ�� = �g���}���u�A�C�h����
            Call Sub_ATLDSET(0, COM_STS_TRM_STATE)                      ' ���[�_�[�o��(ON=�Ȃ�,OFF=�g���}���쒆)
            Timer1.Enabled = True                                       ' �Ď��^�C�}�[�J�n

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "cmdExit_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "���[�h�{�^������������"
    Public Function TrimDataLoad(ByVal sFileName As String) As Boolean
        Dim s As String
        Dim r As Short

        gsDataFileName = sFileName                     ' �f�[�^�t�@�C�����ݒ�

        ' ���ݒ�̑��u�̓d����OFF����
        r = V_Off()                                     ' DC�d�����u �d��OFF����

        Call Z_CLS()                            ' �f�[�^���[�h�Ń��O��ʃN���A                   ###lstLog'V2.0.0.0�N

        ' �g���~���O�f�[�^�ݒ�
        r = UserVal()                                   ' �f�[�^�����ݒ�
        If (r <> 0) Then                                ' �G���[ ?
            pbLoadFlg = False                           ' �f�[�^���[�h�σt���O = False
            s = "Data load Error : " & sFileName & vbCrLf
            LblDataFileName.Text = ""
            Call Z_PRINT(s)
            Return (False)
        Else
            'V2.0.0.0�N            Call Z_CLS()                                ' �f�[�^���[�h�Ń��O��ʃN���A                   ###lstLog
            gDspCounter = 0                             ' ���O��ʕ\��������J�E���^�N���A
            pbLoadFlg = True                            ' �f�[�^���[�h�σt���O = True
            s = "Data loaded : " & sFileName & vbCrLf
            Call Z_PRINT(s)

            Call Me.System1.OperationLogging(gSysPrm, MSG_OPLOG_FUNC01, "File='" & sFileName & "' MANUAL")

            ' �g���~���O�f�[�^�t�@�C���������b�g���t�@�C���ɏo�͂���
            Call PutLotInf()
            ' �t�@�C���p�X���̕\��
            'V2.1.0.0�C��
            If sFileName.Length > 60 Then
                LblDataFileName.Text = sFileName
            Else
                'V2.1.0.0�C��
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    LblDataFileName.Text = "�f�[�^�t�@�C���� " & sFileName
                Else
                    LblDataFileName.Text = "File name " & sFileName
                End If
            End If          'V2.1.0.0�C

#If cOSCILLATORcFLcUSE Then
            '-----------------------------------------------------------------------
            '   FL���։��H�����𑗐M����(FL���ŉ��H�����t�@�C��������ꍇ)
            '-----------------------------------------------------------------------
            Call SendFlParam(sFileName)
#End If
            UserBas.TrimmingDataChange()    ' ###1041�@
            UserSub.LaserCalibrationSet(POWER_CHECK_LOT)            'V2.1.0.0�A ���[�U�p���[���j�^�����O���s�L���ݒ�

            'If giLoaderType <> 0 Then   '�N�����v�z������ݒ�
            '    ObjSys.setClampVaccumConfig(stUserData.intClampVacume)
            'End If

            If UserSub.IsTrimType1() Or UserSub.IsTrimType4() Then       ' ���x�Z���T�[�̎��ȊO�́A�펞�_��'V2.0.0.0�@sTrimType4()�ǉ�
                UserBas.BackLight_Off()
            Else
                UserBas.BackLight_On()
            End If
        End If
        '###1040�E        Call SetATTRateToScreen(True)           ' ###1040�B �g���~���O�f�[�^�ł̂`�s�s�������̐ݒ�

        Return (True)
    End Function

    '''=========================================================================
    ''' <summary>���[�h�{�^������������</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub cmdLoad_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdLoad.Click

        Dim rslt As Short
        Dim s As String
        Dim r As Short
        Dim strMSG As String
        Dim result As System.Windows.Forms.DialogResult

        Try
            '-----------------------------------------------------------------------
            '   ��������
            '-----------------------------------------------------------------------
            ' �g���}���쒆�M��ON���M(���۰�ް)
            Call Sub_ATLDSET(COM_STS_TRM_STATE, 0)              ' ���[�_�[�o��(ON=�g���}���쒆,OFF=�Ȃ�)

            ' �g���}���u�A�C�h�����ȊO�Ȃ�NOP
            If giAppMode Then GoTo STP_END
            giAppMode = APP_MODE_LOAD                           ' ����Ӱ�� = �t�@�C�����[�h(F1)

            ' �p�X���[�h����(�I�v�V����)
            rslt = Func_Password(F_LOAD)
            If (rslt <> True) Then
                GoTo STP_END                                    ' �߽ܰ�ޓ��ʹװ�Ȃ�EXIT
            End If

            ' �f�[�^���[�h�ς݃`�F�b�N
            If (pbLoadFlg = True) Then
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    s = "���[�h�ς݂̃f�[�^������܂��B�ʂ̃f�[�^�����[�h���܂����H"
                Else
                    s = "Current data will be lost. Are you sure to load another data?"
                End If
                r = Me.System1.TrmMsgBox(gSysPrm, s, MsgBoxStyle.OkCancel, cAPPcTITLE)
                If (r = cFRS_ERR_RST) Then ' Cancel(RESET��) ?
                    Call Z_PRINT("Canceled data load." + vbCrLf)
                    GoTo STP_END
                End If
            End If

            'V2.2.0.0�D��
            If giLoaderType <> 0 Then
                ChkLoaderInfoDisp(0)
            End If
            'V2.2.0.0�D��

            '-----------------------------------------------------------------------
            '   �y̧�ق��J���z�޲�۸ނ�\������
            '-----------------------------------------------------------------------
#If cKEYBOARDcUSE <> 1 Then
            ' �\�t�g�E�F�A�L�[�{�[�h���N������
            Dim procHandle As Process
            procHandle = New Process
            Call StartSoftwareKeyBoard(procHandle)              ' �\�t�g�E�F�A�L�[�{�[�h���N������
#End If

            FileDlgOpen.InitialDirectory = "C:\TRIMDATA\DATA"
            FileDlgOpen.FileName = ""
            FileDlgOpen.Filter = "*.txt|*.txt"
            FileDlgOpen.ShowReadOnly = False
            FileDlgOpen.CheckFileExists = True
            FileDlgOpen.CheckPathExists = True

            ' �y̧�ق��J���z�޲�۸ނ�\������
            result = FileDlgOpen.ShowDialog()

#If cKEYBOARDcUSE <> 1 Then
            ' �\�t�g�E�F�A�L�[�{�[�h���I������
            Call EndSoftwareKeyBoard(procHandle)
#End If

            ' OK�ȊO�̏ꍇ
            If (result <> Windows.Forms.DialogResult.OK) Then
                GoTo Cansel                                     ' Cansel�w��Ȃ�I��
            End If

            '-----------------------------------------------------------------------
            '   �f�[�^�t�@�C�������[�h����
            '-----------------------------------------------------------------------
            'If (FileDlgOpen.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            ' �f�[�^�t�@�C�����[�h
            If FileDlgOpen.FileName <> "" Then
                If Not TrimDataLoad(FileDlgOpen.FileName) Then
                    GoTo Cansel
                End If
            End If

            '-----------------------------------------------------------------------
            '   �I������
            '-----------------------------------------------------------------------
Cansel:
            ChDrive("C")                                        ' ChDrive���Ȃ��Ǝ��N����FD�h���C�u�����ɍs����,
            ChDir(My.Application.Info.DirectoryPath)            ' "MVCutil.dll���Ȃ�"�ƂȂ�N���ł��Ȃ��Ȃ�

STP_END:
            Call ZCONRST()                                      ' �ݿ�ٷ� ׯ�����
            giAppMode = APP_MODE_IDLE                           ' ����Ӱ�� = �g���}���u�A�C�h����
            Call Sub_ATLDSET(0, COM_STS_TRM_STATE)              ' ���[�_�[�o��(ON=�Ȃ�,OFF=�g���}���쒆)
            'V2.2.0.0�D��
            If giLoaderType <> 0 Then
                ChkLoaderInfoDisp(1)
            End If
            'V2.2.0.0�D��

            Exit Sub

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "CmdLoad_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            GoTo STP_END
        End Try
    End Sub
#End Region

#Region "FL���։��H�����𑗐M����"
    '''=========================================================================
    ''' <summary>FL���։��H�����𑗐M����</summary>
    ''' <param name="DataFileName">(INP)�g���~���O�f�[�^�t�@�C����</param>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub SendFlParam(ByVal DataFileName As String)

        Dim strMSG As String

        Try
#If cOSCILLATORcFLcUSE Then
        Dim strXmlFName As String
        Dim r As Integer
            '-----------------------------------------------------------------------
            '   FL���։��H�����𑗐M����(FL���ŉ��H�����t�@�C��������ꍇ)
            '-----------------------------------------------------------------------
            If (gSysPrm.stRAT.giOsc_Res = OSCILLATOR_FL) Then
                ' ���H�����t�@�C�������݂��邩�`�F�b�N
                strXmlFName = ""
                r = GetFLCndFileName(DataFileName, strXmlFName, True)
                If (r = SerialErrorCode.rRS_OK) Then                    ' ���H�����t�@�C�������݂��� ?
                    ' �f�[�^���M���̃��b�Z�[�W�\��
                    strMSG = MSG_148
                    Call Z_PRINT(strMSG + vbCrLf)                       ' ү���ޕ\��(���O���)

                    ' FL�p���H�����t�@�C�������[�h����FL���։��H�����𑗐M����
                    r = SendTrimCondInfToFL(stCND, DataFileName, strXmlFName)
                    If (r = SerialErrorCode.rRS_OK) Then
                        ' "FL�։��H�����𑗐M���܂����B"
                        strMSG = MSG_147 & vbCrLf & " (SendDdata File Name = " & strXmlFName & ")" + vbCrLf
                        Call Z_PRINT(strMSG)                            ' ү���ޕ\��(���O���)
                    Else
                        strMSG = MSG_152                                ' "���H�����̑��M�Ɏ��s���܂����B�ēx�f�[�^�����[�h���邩�A�ҏW��ʂ�����H�����̐ݒ���s���Ă��������B"
                        Call MsgBox(strMSG, MsgBoxStyle.OkOnly, "")
                        Call Z_PRINT(strMSG + vbCrLf)                   ' ү���ޕ\��(���O���)
                    End If
                End If
            End If
#End If

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "SendFlParam() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�Z�[�u�{�^������������"
    '''=========================================================================
    ''' <summary>�Z�[�u�{�^������������</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks>�p�����[�^�ύX�������s���A�ύX��̃p�����[�^���f�[�^�t�@�C���֏�����</remarks>
    '''=========================================================================
    Public Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click

        Dim s As String
        Dim r As Short
        Dim strMSG As String
        Dim result As System.Windows.Forms.DialogResult

        Try
            '-----------------------------------------------------------------------
            '   ��������
            '-----------------------------------------------------------------------
            ' �g���}���쒆�M��ON���M(���۰�ް)
            Call Sub_ATLDSET(COM_STS_TRM_STATE, 0)              ' ���[�_�[�o��(ON=�g���}���쒆,OFF=�Ȃ�)

            ' �p�X���[�h����
            r = Func_Password(F_SAVE)
            If (r <> True) Then
                GoTo STP_END                                    ' �߽ܰ�ޓ��ʹװ�Ȃ�EXIT
            End If

            ' ��������
            FlgCan = False                                      ' Cancel Flag = false
            'If (giAppMode) Then                                 ' �g���}���u�A�C�h�����ȊO�Ȃ�NOP
            '    GoTo STP_END
            'End If
            giAppMode = APP_MODE_SAVE                           ' ��ʃX�e�[�^�X = �t�@�C���Z�[�u(F2)
            If pbLoadFlg = False Then
                s = MSG_DataNotLoad                             ' �ް���۰��
                Call Z_PRINT(s)
                Call Beep()
                GoTo STP_END
            End If

            'V2.2.0.0�D��
            If giLoaderType <> 0 Then
                ChkLoaderInfoDisp(0)
            End If
            'V2.2.0.0�D��

            '-----------------------------------------------------------------------
            '   �y���O��t���ĕۑ��z�޲�۸ނ�\������
            '-----------------------------------------------------------------------
#If cKEYBOARDcUSE <> 1 Then
            ' �\�t�g�E�F�A�L�[�{�[�h���N������
            Dim procHandle As Process
            procHandle = New Process
            Call StartSoftwareKeyBoard(procHandle)              ' �\�t�g�E�F�A�L�[�{�[�h���N������
#End If

            '�y���O��t���ĕۑ��z�޲�۸ނ�\������
            FileDlgSave.FileName = gsDataFileName
            FileDlgSave.Filter = "*.txt | *.txt"
            FileDlgSave.OverwritePrompt = True                  ' ���ɑ��݂��Ă���ꍇ�̓��b�Z�[�W �{�b�N�X��\��
            result = FileDlgSave.ShowDialog()                   ' ��̧�ٖ��w��Ȃ��ł͖߂��Ă��Ȃ��A�g���q�t�Ŗ߂��Ă���

#If cKEYBOARDcUSE <> 1 Then
            ' �\�t�g�E�F�A�L�[�{�[�h���I������
            Call EndSoftwareKeyBoard(procHandle)
#End If

            ' OK�ȊO�Ȃ�I��
            If (result <> Windows.Forms.DialogResult.OK) Then
                GoTo STP_TRM                                    ' Cansel�w��Ȃ�I��
            End If

            '-----------------------------------------------------------------------
            '   �f�[�^�t�@�C�����Z�[�u����
            '-----------------------------------------------------------------------
            'If (FileDlgSave.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            If (FileDlgSave.FileName <> "") Then
                If rData_save((FileDlgSave.FileName)) <> 0 Then       ' �f�[�^�t�@�C���Z�[�u
                    GoTo STP_END
                Else
                    gsDataFileName = FileDlgSave.FileName
                    Call Z_PRINT("Data saved : " & FileDlgSave.FileName & vbCrLf)
                End If

                ' �g���~���O�f�[�^�t�@�C���������b�g���t�@�C���ɏo�͂���
                Call PutLotInf()

                '-----------------------------------------------------------------------
                '   ���샍�O�����o�͂���
                '-----------------------------------------------------------------------
                Call Me.System1.OperationLogging(gSysPrm, MSG_OPLOG_FUNC02, "File='" & gsDataFileName & "' MANUAL")

                ' �t�@�C���p�X���̕\��
                'V2.1.0.0�C��
                If gsDataFileName.Length > 60 Then
                    LblDataFileName.Text = gsDataFileName
                Else
                    'V2.1.0.0�C��
                    If (gSysPrm.stTMN.giMsgTyp = 0) Then
                        LblDataFileName.Text = "�f�[�^�t�@�C���� " & gsDataFileName
                    Else
                        LblDataFileName.Text = "File name " & gsDataFileName
                    End If
                End If          'V2.1.0.0�C

#If cOSCILLATORcFLcUSE Then
                Dim strXmlFName As String
                '-----------------------------------------------------------------------
                '   FL�����猻�݂̉��H��������M����FL�p���H�����t�@�C�������C�g����(FL��)
                '-----------------------------------------------------------------------
                If (gSysPrm.stRAT.giOsc_Res = OSCILLATOR_FL) Then ' FL(̧��ްڰ��) ? 
                    strXmlFName = ""
                    r = RcvTrimCondInfToFL(stCND, gsDataFileName, strXmlFName)
                    If (r = SerialErrorCode.rRS_OK) Then
                        ' "���H�����t�@�C�����쐬���܂����B"
                        strMSG = MSG_142 + vbCrLf + " (File Name = " + strXmlFName + ")" + vbCrLf
                        Call Z_PRINT(strMSG)                    ' ү���ޕ\��(���O���)
                    End If
                End If
#End If

                FlgUpd = TriState.False                         ' �f�[�^�X�V Flag OFF
            End If

STP_TRM:
            ChDrive("C")                                        ' ChDrive���Ȃ��Ǝ��N����FD�h���C�u�����ɍs����,"MVCutil.dll���Ȃ�"�ƂȂ�N���ł��Ȃ��Ȃ�
            ChDir(My.Application.Info.DirectoryPath)

STP_END:
            Call ZCONRST()                                      ' �ݿ�ٷ� ׯ�����
            giAppMode = APP_MODE_IDLE                           ' ����Ӱ�� = �g���}���u�A�C�h����
            Call Sub_ATLDSET(0, COM_STS_TRM_STATE)              ' ���[�_�[�o��(ON=�Ȃ�,OFF=�g���}���쒆)

            'V2.2.0.0�D��
            If giLoaderType <> 0 Then
                ChkLoaderInfoDisp(1)
            End If
            'V2.2.0.0�D��

            Exit Sub

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "CmdSave_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            GoTo STP_END
        End Try
    End Sub
#End Region

#Region "�ҏW�{�^������������"
    '''=========================================================================
    ''' <summary>�ҏW�{�^������������</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub cmdEdit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdEdit.Click

        Dim s As String
        Dim r As Short
        'Dim ExeFile As String
        'V2.2.1.6�@ Dim fForm As System.Windows.Forms.Form
        Dim fForm As FormEdit.frmEdit  'V2.2.1.6�@

        Dim strMSG As String
        Dim retbtn As Integer            'V2.2.1.6�@

        Try
            If giAppMode <> APP_MODE_IDLE Then
                Return
            Else
                giAppMode = APP_MODE_EDIT                               ' �A�v�����[�h = �f�[�^�ҏW
            End If
            ' �g���}���쒆�M��ON���M(���۰�ް)
            Call Sub_ATLDSET(COM_STS_TRM_STATE, 0)                      ' ���[�_�[�o��(ON=�g���}���쒆,OFF=�Ȃ�)

            If giLoaderType <> 0 Then
                Call Me.VideoLibrary1.VideoStop()                           ' ���_���A�����ŕ\�����t���[�Y�\�������邽�߈�U��~
                Timer1.Enabled = False
                'V2.2.0.0�D��
                ChkLoaderInfoDisp(0)
                'V2.2.0.0�D��
            End If

            ' �p�X���[�h����
            r = Func_Password(F_EDIT)
            If (r <> True) Then
                giAppMode = APP_MODE_IDLE                               ' ����Ӱ�� = �g���}���u�A�C�h����
                GoTo STP_END                                            ' �߽ܰ�ޓ��ʹװ�Ȃ�EXIT
            End If

            ' �f�[�^���[�h�`�F�b�N
            If pbLoadFlg = False Then
                s = MSG_DataNotLoad                                     ' �ް���۰��
                Call Z_PRINT(s)
                Call Beep()
                GoTo STP_END
            End If

            ' �f�[�^�ҏW
            Call Me.System1.OperationLogging(gSysPrm, MSG_OPLOG_FUNC03, "")
            FlgUpdGPIB = 0                                              ' GPIB�f�[�^�X�VFlag Off
            'fForm = New frmEdit                                         ' frm��޼ު�Đ���
            fForm = New FormEdit.frmEdit                                ' frm��޼ު�Đ���
            fForm.ShowDialog()                                          ' �f�[�^�ҏW
            retbtn = fForm.GetResult()
            fForm.Dispose()                                             ' frm��޼ު�ĊJ��

            ' GPIB�f�[�^�X�V�Ȃ�GPIB���������s��
            If (FlgUpdGPIB = 1) Then
                Call GPIB_Init()
            End If

            '    ' NOTEPAD�Ńf�[�^�t�@�C�����J��
            '    If giAppMode Then Exit Sub
            '    giAppMode = GSTAT_EDIT                                 ' ��ʃX�e�[�^�X = �ҏW��ʕ\��  (F3)
            '    #If cOFFLINEcDEBUG Then
            '        ExeFile = "notepad.exe " + gsDataFileName
            '    #Else
            '        ExeFile = "C:\WINNT\system32\notepad.exe " + gsDataFileName
            '    #End If
            '    r = Shell(ExeFile, vbNormalFocus)

            If LaserFront.Trimmer.DllVideo.VideoLibrary.IsDigitalCamera Then
                ObjVdo.StdMagnification = CDec(stPLT.dblStdMagnification)         ' �����J�����\���{����ݒ� 
            End If


STP_END:

            'V2.2.0.021��
            '�v���[�u�}�X�^�[�e�[�u������f�[�^��W�J����
            If (stPLT.ProbNo > 0) And (DialogResult.OK = retbtn) Then       ' �w��̃v���[�u�f�[�^��Ǎ��ݐݒ肷�� 
                'V2.2.1.6�@��
                '--------------------------------------------------------------------------
                '   �m�Fү���ނ�\������
                '--------------------------------------------------------------------------
                strMSG = "�v���[�u�}�X�^�[��Ǎ��݂܂����H"
                Dim ret As Integer = MsgBox(strMSG, DirectCast(
                            MsgBoxStyle.OkCancel +
                            MsgBoxStyle.Information, MsgBoxStyle),
                            My.Application.Info.Title)
                If (ret = MsgBoxResult.Ok) Then ' Cancel(RESET��) ?
                    ConvProbeData(stPLT.ProbNo)
                End If
                'V2.2.1.6�@��
            End If
            'V2.2.0.021��

            UserBas.TrimmingDataChange()    ' ###1041�@

            If UserSub.IsTrimType1() Or UserSub.IsTrimType4() Then       ' ���x�Z���T�[�̎��ȊO�́A�펞�_��'V2.0.0.0�@sTrimType4()�ǉ�
                UserBas.BackLight_Off()
            Else
                UserBas.BackLight_On()
            End If

            Call ZCONRST()                                              ' �ݿ�ٷ� ׯ�����
            giAppMode = APP_MODE_IDLE                                   ' ����Ӱ�� = �g���}���u�A�C�h����
            Call Sub_ATLDSET(0, COM_STS_TRM_STATE)                      ' ���[�_�[�o��(ON=�Ȃ�,OFF=�g���}���쒆)

            If giLoaderType <> 0 Then
                Call Me.VideoLibrary1.VideoStart() '                    ' 
                Timer1.Enabled = True
                'V2.2.0.0�D��
                ChkLoaderInfoDisp(1)
                'V2.2.0.0�D��
            End If

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "cmdEdit_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        Finally
            giAppMode = APP_MODE_IDLE                                   ' ����Ӱ�� = �g���}���u�A�C�h����
        End Try
    End Sub
#End Region

#Region "���[�U�����{�^������������"
    '''=========================================================================
    ''' <summary>���[�U�����{�^������������</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    '''=========================================================================
    Private Sub cmdLaserTeach_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdLaserTeach.Click

        Dim strMSG As String

        Try
            ' ���[�U���������s����
            cmdLaserTeach_Proc()

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "cmdLaserTeach_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "���[�U�������s��"
    '''=========================================================================
    ''' <summary>���[�U�������s��</summary>
    '''=========================================================================
    Public Sub cmdLaserTeach_Proc()

        Dim r As Short
        Dim strMSG As String

        Try
            ' �R�}���h���s�O����
            r = Sub_cmdInit_Proc(APP_MODE_LASER, F_LASER)
            If (r = cFRS_ERR_RST) Then                                  ' Cancel(RESET��) ?
                If (gSysPrm.stSPF.giWithStartSw = 0) Then                ' ����SW�����҂�(��߼��)�łȂ� ?
                    GoTo STP_END
                Else                                                    ' ����SW�����҂�(��߼��)
                    GoTo STP_TRM
                End If
            End If
            If (r < cFRS_NORMAL) Then                                   ' �G���[ ? 
                ' �J�o�[�J���o(����SW�����҂�(��߼��)����������)
                If (r = cFRS_ERR_CVR) Or (r = cFRS_ERR_SCVR) Or (r = cFRS_ERR_LATCH) Then
                    r = cFRS_NORMAL
                    GoTo STP_TRM                                        ' �J�o�[�J���o�Ȃ�ײ�޶�ް�۽ނ���READY��Ԃ�
                Else
                    GoTo STP_END
                End If
            End If

            ' ���[�U�[�����������s��
            gbInitialized = False                                       ' flg = ���_���A��
            SetLaserItemsVisible(0)                                     ' ���[�U�p���[�����֘A���ڂ��\���Ƃ���
            r = User_LaserTeach()                                       ' ���[�U�[������ʕ\��
            SetLaserItemsVisible(1)                                     ' ���[�U�p���[�����֘A���ڂ�\���Ƃ���

            If r = cFRS_ERR_CVR Then                                    ' ���[�U�R�}���h����➑̃J�o�[�J�́ASub_cmdTerm_Proc()���ŋ����I���Ƃ���B
                r = cFRS_ERR_EMG
            End If
            ' �R�}���h���s�㏈��
STP_TRM:
            Call Sub_cmdTerm_Proc(APP_MODE_LASER, r)

            ' �R�}���h�I������
STP_END:
            Call Sub_cmdEnd_Proc()

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "cmdLaserTeach_Proc() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

    'V2.1.0.0�A��
#Region "���[�U�������s��"
    ''' <summary>
    ''' ���[�U�p���[�L�����u���[�V����
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub cmdLaserTeach_Calibration()

        Dim r As Short
        Dim strMSG As String

        Try
            ' �R�}���h���s�O����
            r = Sub_cmdInit_Proc(APP_MODE_LASER, F_LASER)
            If (r = cFRS_ERR_RST) Then                                  ' Cancel(RESET��) ?
                If (gSysPrm.stSPF.giWithStartSw = 0) Then                ' ����SW�����҂�(��߼��)�łȂ� ?
                    GoTo STP_END
                Else                                                    ' ����SW�����҂�(��߼��)
                    GoTo STP_TRM
                End If
            End If
            If (r < cFRS_NORMAL) Then                                   ' �G���[ ? 
                ' �J�o�[�J���o(����SW�����҂�(��߼��)����������)
                If (r = cFRS_ERR_CVR) Or (r = cFRS_ERR_SCVR) Or (r = cFRS_ERR_LATCH) Then
                    r = cFRS_NORMAL
                    GoTo STP_TRM                                        ' �J�o�[�J���o�Ȃ�ײ�޶�ް�۽ނ���READY��Ԃ�
                Else
                    GoTo STP_END
                End If
            End If

            ' ���[�U�[�����������s��
            gbInitialized = False                                       ' flg = ���_���A��
            SetLaserItemsVisible(0)                                     ' ���[�U�p���[�����֘A���ڂ��\���Ƃ���
            r = User_LaserTeach(True)                                   ' ������True�́A���[�U�p���[�L�����u���[�V����
            SetLaserItemsVisible(1)                                     ' ���[�U�p���[�����֘A���ڂ�\���Ƃ���

            If r = cFRS_ERR_CVR Then                                    ' ���[�U�R�}���h����➑̃J�o�[�J�́ASub_cmdTerm_Proc()���ŋ����I���Ƃ���B
                r = cFRS_ERR_EMG
            End If
            ' �R�}���h���s�㏈��
STP_TRM:
            Call Sub_cmdTerm_Proc(APP_MODE_LASER, r)

            ' �R�}���h�I������
STP_END:
            Call Sub_cmdEnd_Proc()

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "cmdLaserTeach_Calibration() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region
    'V2.1.0.0�A��

#Region "���b�g�ؑփ{�^������������"
    '''=========================================================================
    ''' <summary>���b�g�ؑփ{�^������������</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub cmdLotchg_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdLotChg.Click

        Dim Rtn As Short
        Dim strMSG As String

        Try

            ''V2.2.0.0�D��
            ' TLF�����[�_�̏ꍇ�����^�]���̊�����`�F�b�N
            If giLoaderType = 1 Then
                Timer1.Enabled = False                                      ' �Ď��^�C�}�[��~
                Rtn = ObjLoader.Sub_SubstrateNothingCheck(Me.System1)
                If Rtn <> cFRS_NORMAL Then
                    GoTo STP_END
                End If
                stCounter.PlateCounter = 0
            End If
            ''V2.2.0.0�D��

            ' �g���}���쒆�M��ON���M(���۰�ް)
            Call Sub_ATLDSET(COM_STS_TRM_STATE, 0)                      ' ���[�_�[�o��(ON=�g���}���쒆,OFF=�Ȃ�)

            ' ���R�}���h���s�� ?
            If giAppMode Then GoTo STP_END
            giAppMode = APP_MODE_LOTCHG                                 ' ����Ӱ�� = ���b�g�ؑ�


            ' ���샍�O�o��
            Call Me.System1.OperationLogging(gSysPrm, MSG_OPLOG_FUNC06, "MANUAL")

            'V2.2.0.0�D��
            If giLoaderType = 1 Then
                ChkLoaderInfoDisp(0)
                MarkingCount = 0             ' �}�[�L���O�p�J�E���^�N���A	V2.2.1.7�B
                LotMarkingAlarmCnt = 0       ' �}�[�L���O���A���[���J�E���^�N���A	V2.2.1.7�B
            End If
            'V2.2.0.0�D��

            ' ���b�gNO.���͏���
            'frmObj = New FormDataSelect(Me)                ' Form����
            frmAutoObj.ShowDialog()                         ' ���b�gNO.����
            Rtn = frmAutoObj.sGetReturn()
            'frmAutoObj.Close()                             ' Form�A�����[�h
            '
            Call COVERLATCH_CLEAR()                         '�J�o�[�J���b�`�N���A V2.2.0.035 

            If Rtn = cFRS_ERR_START Then
                Call Me.System1.OperationLogging(gSysPrm, "�����^�]�n�j", "MANUAL")   'V2.0.0.2�@
                Call UserSub.SetStartCheckStatus(True)                  ' �ݒ��ʂ̊m�F�L����'V2.0.0.2�@���b�g�X�^�[�g�����̒��ŏ�����������������ׂɂ����ł��ݒ肷��
                'V2.0.0.2�@                If UserSub.IsSpecialTrimType() Then             ' ���[�U�v���O�������ꏈ��
                If Not UserBas.LotStartSetting() Then       ' ���b�g�X�^�[�g�������i����f�[�^�w�b�_�[���쐬�Ȃǁj
                    frmAutoObj.gbFgAutoOperation = False     ' �����^�]����
                    Me.AutoRunnningDisp.BackColor = System.Drawing.Color.Yellow  ' �w�i�F = ���F
                    Me.AutoRunnningDisp.Text = "�����^�]������"
                End If
                'V2.0.0.2�@            End If                '

                ''V2.2.0.0�D��
                ' TLF�����[�_�̏ꍇ�����^�]�؂�ւ����o�͂���
                If giLoaderType = 1 Then
                    Call Sub_ATLDSET(clsLoaderIf.LOUT_AUTO Or clsLoaderIf.LOUT_STOP, clsLoaderIf.LOUT_REDY)                    'V1.2.0.0�C ���[�_�[�o��(ON=����,OFF=�Ȃ�)
                    ObjLoader.gbIniFlg = 0
                End If
                ''V2.2.0.0�D��

                If frmAutoObj.gbFgAutoOperation Then
                    'If stUserData.iLotChange = 2 Or stUserData.iLotChange = 3 Then
                    Rtn = System1.ReadHostCommand_ForVBNET(giHostMode, gbHostConnected, giHostRun, giAppMode, pbLoadFlg)  ' ���[�_����f�[�^����͂���
                    If (giHostMode <> cHOSTcMODEcAUTO) Then
                        ' ���[�_�������ɐ؂�ւ��܂ő҂�
                        Rtn = Me.System1.Form_Reset(cGMODE_LDR_CHK_AUTO, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
                        If (Rtn <= cFRS_ERR_EMG) Then
                            GoTo STP_ERR_EXIT                       ' ����~���̃G���[�Ȃ�A�v�������I��
                        ElseIf (Rtn = cFRS_ERR_RST) Then              ' �L�����Z���Ȃ�R�}���h�I��
                            ''V2.2.0.0�D��
                            ' TLF�����[�_�̏ꍇ�����^�]�؂�ւ����o�͂���
                            If giLoaderType = 1 Then
                                Call Sub_ATLDSET(0, clsLoaderIf.LOUT_AUTO)                    'V1.2.0.0�C ���[�_�[�o��(ON=����,OFF=�Ȃ�)
                                ObjLoader.gbIniFlg = 0
                            End If
                            ''V2.2.0.0�D��
                            frmAutoObj.SetAutoOpeCancel(True)               ' V2.2.1.1�A
                            Call frmAutoObj.AutoOperationEnd()              ' V2.2.1.1�A

                            frmAutoObj.gbFgAutoOperation = False
                            Me.AutoRunnningDisp.BackColor = System.Drawing.Color.Yellow  ' �w�i�F = ���F
                            Me.AutoRunnningDisp.Text = "�����^�]������"
                            GoTo STP_END
                        End If
                    End If
                    'End If
                    Call UserSub.SetStartCheckStatus(True)                  'V1.2.0.0�C �ݒ��ʂ̊m�F�L����
                    ''V2.2.0.0�D��
                    SetAutoOpeStartTime()
                    If giLoaderType = 1 Then
                        ' TLF�����[�_�̏ꍇ�����^�]�؂�ւ����o�͂���
                        Call Sub_ATLDSET(clsLoaderIf.LOUT_AUTO Or clsLoaderIf.LOUT_STOP, clsLoaderIf.LOUT_REDY)                    'V1.2.0.0�C ���[�_�[�o��(ON=����,OFF=�Ȃ�)
                        ObjLoader.gbIniFlg = 0
                        ' �d�����b�N(�ω����E�����b�N)����������
                        Rtn = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_ON)
                        If (Rtn = cFRS_TO_EXLOCK) Then                                ' �u�O�ʔ����b�N�����^�C���A�E�g�v�Ȃ�߂�l���uRESET�v�ɂ���

                            GoTo STP_END
                        End If

                    Else
                        Call Sub_ATLDSET(0, COM_STS_LOT_END)                    'V1.2.0.0�C ���[�_�[�o��(ON=�Ȃ�,OFF=���b�g�I��)
                    End If
                    ''V2.2.0.0�D��
                End If
            Else                                                                            'V2.0.0.2�@
                Call Me.System1.OperationLogging(gSysPrm, "�����^�]�L�����Z��", "MANUAL")   'V2.0.0.2�@
            End If

STP_END:
            'V2.2.0.0�D��
            If giLoaderType = 1 Then
                ChkLoaderInfoDisp(1)
            End If
            'V2.2.0.0�D��

            Call ZCONRST()                                              ' �ݿ�ٷ� ׯ�����
            giAppMode = APP_MODE_IDLE                                   ' ����Ӱ�� = �g���}���u�A�C�h����
            Timer1.Enabled = True                                       ' �Ď��^�C�}�[�J�n
            Call Sub_ATLDSET(0, COM_STS_TRM_STATE)                      ' ���[�_�[�o��(ON=�Ȃ�,OFF=�g���}���쒆)
            Exit Sub

            ' ��ċ����I������
STP_ERR_EXIT:
            Call AppEndDataSave()                                       ' ��ċ����I�������ް��ۑ��m�F
            Call AplicationForcedEnding()                               ' ��ċ����I������
            End
            Exit Sub

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "cmdLotchg_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        Finally
            giAppMode = APP_MODE_IDLE                                   ' ����Ӱ�� = �g���}���u�A�C�h����
        End Try
    End Sub
#End Region

#Region "�v���[�u�{�^������������"
    '''=========================================================================
    ''' <summary>�v���[�u�{�^������������</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    '''=========================================================================
    Private Sub cmdProbeTeaching_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdProbeTeaching.Click

        Dim strMSG As String

        Try
            ' �v���[�u�R�}���h�����s����
            cmdProbeTeaching_Proc()

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "cmdProbeTeaching_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�v���[�u�R�}���h���s"
    '''=========================================================================
    ''' <summary>�v���[�u�R�}���h���s</summary>
    '''=========================================================================
    Public Sub cmdProbeTeaching_Proc()

        Dim r As Short
        Dim strMSG As String

        Try
            ' �R�}���h���s�O����
            r = Sub_cmdInit_Proc(APP_MODE_PROBE, F_PROBE)
            If (r = cFRS_ERR_RST) Then                                  ' Cancel(RESET��) ?
                If (gSysPrm.stSPF.giWithStartSw = 0) Then                ' ����SW�����҂�(��߼��)�łȂ� ?
                    GoTo STP_END
                Else                                                    ' ����SW�����҂�(��߼��)
                    GoTo STP_TRM
                End If
            End If
            If (r < cFRS_NORMAL) Then                                   ' �G���[ ?
                ' �J�o�[�J���o(����SW�����҂�(��߼��)����������)
                If (r = cFRS_ERR_CVR) Or (r = cFRS_ERR_SCVR) Or (r = cFRS_ERR_LATCH) Then
                    r = cFRS_NORMAL
                    GoTo STP_TRM                                        ' �J�o�[�J���o�Ȃ�ײ�޶�ް�۽ނ���READY��Ԃ�
                Else
                    GoTo STP_END
                End If
            End If

            ' �v���[�u�e�B�[�`���O����
            r = User_ProbeTeaching()                                    ' �v���[�u�e�B�[�`���O

            ' �R�}���h���s�㏈��
STP_TRM:
            Call Sub_cmdTerm_Proc(APP_MODE_PROBE, r)

            ' �R�}���h�I������
STP_END:
            Call Sub_cmdEnd_Proc()

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "cmdProbeTeaching_Proc() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "TEACH�{�^������������"
    '''=========================================================================
    ''' <summary>TEACH�{�^������������</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub cmdTeaching_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTeaching.Click

        Dim strMSG As String

        Try
            ' �X�^�[�g�|�W�V���� �e�B�[�`���O���s��
            cmdTeaching_Proc()

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "BtnTEACH_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�X�^�[�g�|�W�V���� �e�B�[�`���O���s��"
    '''=========================================================================
    ''' <summary>�X�^�[�g�|�W�V���� �e�B�[�`���O���s��</summary>
    '''=========================================================================
    Public Sub cmdTeaching_Proc()

        Dim r As Short
        'Dim ObjGazou As Process = Nothing                               ' Process�I�u�W�F�N�g
        Dim strMSG As String

        Try
            ' �R�}���h���s�O����
            r = Sub_cmdInit_Proc(APP_MODE_TEACH, F_TEACH)
            If (r = cFRS_ERR_RST) Then                                  ' Cancel(RESET��) ?
                If (gSysPrm.stSPF.giWithStartSw = 0) Then                ' ����SW�����҂�(��߼��)�łȂ� ?
                    GoTo STP_END
                Else                                                    ' ����SW�����҂�(��߼��)
                    GoTo STP_TRM
                End If
            End If
            If (r < cFRS_NORMAL) Then                                   ' �G���[ ?
                ' �J�o�[�J���o(����SW�����҂�(��߼��)����������)
                If (r = cFRS_ERR_CVR) Or (r = cFRS_ERR_SCVR) Or (r = cFRS_ERR_LATCH) Then
                    r = cFRS_NORMAL
                    GoTo STP_TRM                                        ' �J�o�[�J���o�Ȃ�ײ�޶�ް�۽ނ���READY��Ԃ�
                Else
                    GoTo STP_END
                End If
            End If

            '-----------------------------------------------------------------------------
            '   �e�B�[�`���O�������s��
            '-----------------------------------------------------------------------------
            gbInitialized = False

            ' �摜�\���v���O�������N������(�J�b�g�g���[�X�p)
            ' ���摜�\���v���O�����̋N����OcxTeach�ōs�����߂����ł͋N�����Ȃ�
            'r = Exec_GazouProc(ObjGazou, DISPGAZOU_PATH, DISPGAZOU_WRK, 0)

            Call ObjSys.Ilum_Ctrl(gSysPrm, Z1, ZOPT)                ' ���Ȱ������ߓ_��(����L���L��)

            ' �e�B�[�`���O��ʏ���
            r = User_teaching()

            Call ObjSys.Ilum_Ctrl(gSysPrm, Z0, ZOPT)                ' ���Ȱ������ߏ���(����L���L��)

            ' �摜�\���v���O�������I������
            'End_GazouProc(ObjGazou)

            ' �R�}���h���s�㏈��
STP_TRM:
            Call Sub_cmdTerm_Proc(APP_MODE_TEACH, r)

            ' �R�}���h�I������
STP_END:

            Call Sub_cmdEnd_Proc()

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "cmdTeaching_Proc() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�J�b�g�ʒu�␳�{�^������������"
    '''=========================================================================
    ''' <summary>�J�b�g�ʒu�␳�{�^������������</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub cmdCutPosTeach_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCutPosTeach.Click

        Dim r As Short
        Dim strMSG As String

        Try
            ' �R�}���h���s�O����
            r = Sub_cmdInit_Proc(APP_MODE_CUTPOS, F_CUTPOS)
            If (r = cFRS_ERR_RST) Then                                  ' Cancel(RESET��) ?
                If (gSysPrm.stSPF.giWithStartSw = 0) Then                ' ����SW�����҂�(��߼��)�łȂ� ?
                    GoTo STP_END
                Else                                                    ' ����SW�����҂�(��߼��)
                    GoTo STP_TRM
                End If
            End If
            If (r = cFRS_FNG_CPOS) Then                                 ' �J�b�g�ʒu�␳�Ώۂ̒�R���Ȃ� ?
                GoTo STP_TRM
            End If

            If (r < cFRS_NORMAL) Then                                   ' �G���[ ?
                ' �J�o�[�J���o(����SW�����҂�(��߼��)����������)
                If (r = cFRS_ERR_CVR) Or (r = cFRS_ERR_SCVR) Or (r = cFRS_ERR_LATCH) Then
                    r = cFRS_NORMAL
                    GoTo STP_TRM                                        ' �J�o�[�J���o�Ȃ�ײ�޶�ް�۽ނ���READY��Ԃ�
                Else
                    GoTo STP_END
                End If
            End If

            ' �J�b�g�ʒu�␳�ׂ̈̉摜�o�^�������s��
            gbInitialized = False
            ChDir(My.Application.Info.DirectoryPath)
            r = User_CutpositionTeach()                                 ' �摜�o�^����

            ' �R�}���h���s�㏈��
STP_TRM:
            Call Sub_cmdTerm_Proc(APP_MODE_CUTPOS, r)

            ' �R�}���h�I������
STP_END:
            Call Sub_cmdEnd_Proc()

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "cmdCutPosTeach_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "RECOG�{�^������������"
    '''=========================================================================
    ''' <summary>RECOG�{�^������������</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnRECOG_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnRECOG.Click

        Dim strMSG As String

        Try
            ' RECOG�������s��
            BtnRECOG_Proc()

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "BtnRECOG_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "RECOG�������s��"
    '''=========================================================================
    ''' <summary>RECOG�������s��</summary>
    '''=========================================================================
    Public Sub BtnRECOG_Proc()

        Dim r As Short
        Dim strMSG As String

        Try
            ' �R�}���h���s�O����
            r = Sub_cmdInit_Proc(APP_MODE_RECOG, F_RECOG)
            If (r = cFRS_ERR_RST) Then                                  ' Cancel(RESET��) ?
                If (gSysPrm.stSPF.giWithStartSw = 0) Then                ' ����SW�����҂�(��߼��)�łȂ� ?
                    GoTo STP_END
                Else                                                    ' ����SW�����҂�(��߼��)
                    GoTo STP_TRM
                End If
            End If
            If (r < cFRS_NORMAL) Then                                   ' �G���[ ?
                ' �J�o�[�J���o(����SW�����҂�(��߼��)����������)
                If (r = cFRS_ERR_CVR) Or (r = cFRS_ERR_SCVR) Or (r = cFRS_ERR_LATCH) Then
                    r = cFRS_NORMAL
                    GoTo STP_TRM                                        ' �J�o�[�J���o�Ȃ�ײ�޶�ް�۽ނ���READY��Ԃ�
                Else
                    GoTo STP_END
                End If
            End If

            Call ObjSys.Ilum_Ctrl(gSysPrm, Z1, ZOPT)                    ' ���Ȱ������ߓ_��(����L���L��)

            ' �ƕ␳�ׂ̈̉摜�o�^�������s��
            r = User_PatternTeach()                                     ' �摜�o�^����

            ' �R�}���h���s�㏈��
STP_TRM:
            Call Sub_cmdTerm_Proc(APP_MODE_RECOG, r)

            ' �R�}���h�I������
STP_END:
            Call Sub_cmdEnd_Proc()

            Call Me.System1.Ilum_Ctrl(gSysPrm, Z0, ZOPT)            ' ���Ȱ������ߏ���(����L���L��)

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "BtnRECOG_Proc() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�s�w�{�^������������"
    '''=========================================================================
    ''' <summary>
    ''' �s�w�{�^������������
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub CmdTx_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdTx.Click
        Try
            TxTyTeach_Proc(APP_MODE_TX)
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Me)
        End Try
    End Sub
#End Region

#Region "�s�x�{�^������������"
    '''=========================================================================
    ''' <summary>
    ''' �s�x�{�^������������
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub CmdTy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdTy.Click
        Try
            TxTyTeach_Proc(APP_MODE_TY)
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Me)
        End Try
    End Sub
#End Region

#Region "�s�w�A�s�x�����̎��s"
    '''=========================================================================
    ''' <summary>
    ''' �s�w�A�s�x�����̎��s
    ''' </summary>
    ''' <param name="AppMode">APP_MODE_TX�܂���APP_MODE_TY�̃��[�h</param>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub TxTyTeach_Proc(ByVal AppMode As Short)

        Dim r As Short
        Dim strMSG As String
        Dim FncIdx As Short

        Try
            ' �R�}���h���s�O����
            If AppMode = APP_MODE_TX Then
                FncIdx = F_TX
            Else
                FncIdx = F_TY
            End If

            r = Sub_cmdInit_Proc(AppMode, FncIdx)
            If (r = cFRS_ERR_RST) Then                                  ' Cancel(RESET��) ?
                If (gSysPrm.stSPF.giWithStartSw = 0) Then                ' ����SW�����҂�(��߼��)�łȂ� ?
                    GoTo STP_END
                Else                                                    ' ����SW�����҂�(��߼��)
                    GoTo STP_TRM
                End If
            End If
            If (r < cFRS_NORMAL) Then                                   ' �G���[ ?
                ' �J�o�[�J���o(����SW�����҂�(��߼��)����������)
                If (r = cFRS_ERR_CVR) Or (r = cFRS_ERR_SCVR) Or (r = cFRS_ERR_LATCH) Then
                    r = cFRS_NORMAL
                    GoTo STP_TRM                                        ' �J�o�[�J���o�Ȃ�ײ�޶�ް�۽ނ���READY��Ԃ�
                Else
                    GoTo STP_END
                End If
            End If

            Call ObjSys.Ilum_Ctrl(gSysPrm, Z1, ZOPT)                    ' ���Ȱ������ߓ_��(����L���L��)

            Call ZCONRST()                                              ' �ݿ��SWׯ�����

            ' �ƕ␳�ׂ̈̉摜�o�^�������s��
            r = User_TxTyTeach()                                        ' �s�w�C�s�x����

            ' �R�}���h���s�㏈��
STP_TRM:
            Call Sub_cmdTerm_Proc(AppMode, r)

            ' �R�}���h�I������
STP_END:

            Call Sub_cmdEnd_Proc()

            Call Me.System1.Ilum_Ctrl(gSysPrm, Z0, ZOPT)            ' ���Ȱ������ߏ���(����L���L��)

            If r = cFRS_TxTy Then
                ' �X�^�[�g�|�W�V���� �e�B�[�`���O���s��
                cmdTeaching_Proc()
            End If

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "TxTyTeach_Proc() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region
    '========================================================================================
    '   �e�R�}���h���s�O/�㏈��
    '========================================================================================
#Region "�R�}���h���s�O����"
    '''=========================================================================
    '''<summary>�R�}���h���s�O����</summary>
    '''<param name="gSts">  ����Ӱ��(giAppMode�Q��)</param>
    '''<param name="FncIdx">�@�\�I���`�e�[�u���̲��ޯ��</param>
    ''' <returns>  0  = ����
    '''            3  = Reset SW����
    '''           -80 = �f�[�^�����[�h
    '''           -81 = ���R�}���h���s��
    '''           -82 = �߽ܰ�ޓ��ʹװ
    ''' </returns>
    '''<remarks>����~/�W�o�@�ُ펞�͓��֐����ſ�ċ����I������</remarks>
    '''=========================================================================
    Private Function Sub_cmdInit_Proc(ByRef gSts As Short, ByRef FncIdx As Short) As Short

        Dim r As Short
        Dim InterlockSts As Integer
        Dim SwitchSts As Long
        Dim s As String
        Dim strMSG As String
        Dim iRtn As Integer

        Try
            ' �g���}���쒆�M��ON���M(���۰�ް)
            Call Sub_ATLDSET(COM_STS_TRM_STATE, 0)                      ' ���[�_�[�o��(ON=�g���}���쒆,OFF=�Ȃ�)

            ' ���샍�O�o��
            Call Sub_OprLog(gSts)

            ' �p�X���[�h����(����)
            r = Func_Password(FncIdx)
            If (r <> True) Then                                         ' �߽ܰ�ޓ��ʹװ ?
                Return (cFRS_FNG_PASS)                                  ' Return�l = �߽ܰ�ޓ��ʹװ
            End If

            ' ���R�}���h���s�� ?
            If giAppMode <> APP_MODE_IDLE Then                          ' ���R�}���h���s�� ?
                Return (cFRS_FNG_CMD)                                   ' Return�l = ���R�}���h���s��
            End If

            '' �R�}���h���s�O�̃`�F�b�N
            'r = CmdExec_Check(gSts)
            'If (r <> cFRS_NORMAL) Then                                  ' �`�F�b�N�G���[ ?
            '    Return (r)                                              ' Return�l = �`�F�b�N�G���[(Cancel(RESET��)��Ԃ�)
            'End If

            giAppMode = gSts                                            ' ����Ӱ�ސݒ�

            ' �f�[�^���[�h�ς݃`�F�b�N
            If (pbLoadFlg = False) Then                                 ' �f�[�^�����[�h ?
                s = MSG_DataNotLoad                                     ' �ް���۰��
                Call Z_PRINT(s)                                         ' "Data is not loaded. Please Load the data file."
                Call Beep()
                Return (cFRS_FNG_DATA)                                  ' Return�l = �ް���۰��
                Exit Function
            End If

            ' �R�}���h���s�O�̃`�F�b�N
            r = CmdExec_Check(giAppMode)
            If (r <> cFRS_NORMAL) Then                                  ' �`�F�b�N�G���[ ?
                Return (r)                                              ' Return�l = �`�F�b�N�G���[(Cancel(RESET��)��Ԃ�)
            End If

            ' �W�o�@�ُ�`�F�b�N
            r = Me.System1.CheckDustVaccumeAlarm(gSysPrm)
            If (r <> 0) Then                                            ' �G���[�Ȃ�W�o�@�ُ팟�o���b�Z�[�W�\��
                Call Me.System1.Form_Reset(cGMODE_ERR_DUST, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
                GoTo STP_ERR_EXIT                                       ' ��ċ����I��
            End If

            '' ���Ȱ������ߓ_��(����L���L��)
            'If (gSts <> APP_MODE_LASER) Then                            ' LASER����ނ͓_�����Ȃ�
            '    Call Me.System1.Ilum_Ctrl(gSysPrm, Z1, ZOPT)
            'End If

            ' ����m�F���(START/RESET�҂�)
            r = Me.System1.Form_Reset(cGMODE_START_RESET, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
            Me.Refresh()
            If (r = cFRS_ERR_RST) Then                                  ' Reset SW���� ?
                Return (cFRS_ERR_RST)                                   ' Return�l = Reset SW����
            End If
            If (r = cFRS_NORMAL) Or (r = cFRS_ERR_START) Then           ' ���� ?
                ' �N�����v/�z��ON
                r = System1.ClampVacume_Ctrl(gSysPrm, 1, giAppMode, 1)
                If (r <> cFRS_NORMAL) Then GoTo STP_ERR_EXIT

                ' �V�O�i���^���[���_��(�e�B�[�`���O��) 
                ' ���A���C���^�[���b�N������(���_��)�D��
                r = INTERLOCK_CHECK(InterlockSts, SwitchSts)
                If (InterlockSts = INTERLOCK_STS_DISABLE_NO) Then       ' �C���^�[���b�N���Ȃ物�_��
                    r = Me.System1.SetSignalTower(SIGOUT_YLW_ON, &HFFFF)
                End If
            ElseIf (r <= cFRS_ERR_EMG) Then                             ' �װ(����~��)�Ȃ��ċ����I��
                GoTo STP_ERR_EXIT
            Else
                Sub_cmdInit_Proc = r                                    ' Return�l�ݒ�
            End If

            Call UserSub.ClampVacumeChange()         'V2.0.0.0�M

            ' �{�^�������\���ɂ���
            cmdHelp.Visible = False                                     ' Version�{�^����\�� 
            Me.Grpcmds.Visible = False                                  ' �R�}���h�{�^���O���[�v�{�b�N�X��\��
            Me.GrpMode.Visible = False                                  ' �f�B�W�^��SW�O���[�v�{�b�N�X��\��
            Me.frmInfo.Visible = False                                  ' ���ʕ\�����\��
            BtnStartPosSet.Enabled = False                              'V2.0.0.0�A
            gbInitialized = False
            ButtonLaserCalibration.Visible = False                      'V2.1.0.0�A
            btnCutStop.Visible = False                                  'V2.2.0.0�E
            btnLoaderInfo.Visible = False                               'V2.2.0.0�D

            Timer1.Enabled = False          ' @@@888 

            SetMagnifyBar(True)                                         ' V2.2.0.0�@

            ChkLoaderInfoDisp(0)                              'V2.2.0.0�D

            If gSts = APP_MODE_LASER Then
                ' Z�����_�ֈړ�
                iRtn = EX_ZMOVE(0)
                If (iRtn <> cFRS_NORMAL) Then                              ' �G���[ ?(���b�Z�[�W�͕\���ς�) 
                    Call Me.System1.TrmMsgBox(gSysPrm, "SETZOFFPOS �y�����_�ʒu�ړ����ُ�I�����܂����B", MsgBoxStyle.OkOnly, My.Application.Info.Title)
                    Return (cERR_END)
                End If
            Else
                iRtn = SetZOff_Prob_Off()                        ' INTIME�����̑ҋ@�ʒu��ύX���đҋ@�ʒu�ړ�����B
                If iRtn <> cFRS_NORMAL Then
                    Call Me.System1.TrmMsgBox(gSysPrm, "SETZOFFPOS �y���ҋ@�ʒu�ύX���ُ�I�����܂����B", MsgBoxStyle.OkOnly, My.Application.Info.Title)
                    Return (cERR_END)
                End If
            End If

            Return (cFRS_NORMAL)

            ' ��ċ����I������
STP_ERR_EXIT:
            Call AppEndDataSave()                                       ' ��ċ����I�������ް��ۑ��m�F
            Call AplicationForcedEnding()                               ' ��ċ����I������
            End
            Return (r)

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "Sub_cmdInit_Proc() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Function
#End Region

#Region "�e�R�}���h���s�O�̃`�F�b�N����"
    '''=========================================================================
    ''' <summary>�e�R�}���h���s�O�̃`�F�b�N����</summary>
    ''' <param name="iAppMode">(INP) ����Ӱ��(giAppMode�Q��)</param>
    ''' <returns>cFRS_NORMAL   = ����
    '''          cFRS_FNG_CPOS = �J�b�g�ʒu�␳�Ώۂ̒�R���Ȃ�
    '''          ��L�ȊO�̃G���[
    ''' </returns>
    '''=========================================================================
    Private Function CmdExec_Check(ByRef iAppMode As Short) As Short

        Dim bFlg As Boolean
        Dim Rn As Integer
        Dim RtnCode As Short
        Dim strMSG As String

        Try
            '-------------------------------------------------------------------
            '   �R�}���h���s�O�̃`�F�b�N���s��
            '-------------------------------------------------------------------
            ' �J�b�g�ʒu�␳�R�}���h��
            If (iAppMode = APP_MODE_CUTPOS) Then
                ' �p�^�[���o�^�f�[�^�����邩�`�F�b�N����
                bFlg = False
                For Rn = 1 To stPLT.PtnCount                            ' �p�^�[���o�^�����J��Ԃ�
                    ' �p�^�[���o�^���� ?
                    'V1.0.4.3�E                    If (stPTN(Rn).PtnFlg <> CUT_PATTERN_NONE And stPTN(Rn).PtnFlg <> 3) Then
                    If (stPTN(Rn).PtnFlg <> CUT_PATTERN_NONE) Then
                        bFlg = True
                        Exit For
                    End If
                Next Rn

                ' �p�^�[���o�^�f�[�^���Ȃ��ꍇ�͏������Ȃ�
                If (bFlg = False) Then
                    strMSG = MSG_153                                    ' "�J�b�g�ʒu�␳�Ώۂ̒�R������܂���"
                    RtnCode = cFRS_FNG_CPOS                             ' Return�l = �J�b�g�ʒu�␳�Ώۂ̒�R���Ȃ�
                    GoTo STP_ERR_EXIT                                   ' ���b�Z�[�W�\����G���[�߂�
                End If
            End If

            ' 'V2.2.0.0�D ��
            If giLoaderType <> 0 Then
                If (iAppMode <> APP_MODE_LASER) Then
                    ' �ڕ���Ɋ�����鎖���`�F�b�N����(�蓮���[�h��(OPTION))
                    RtnCode = ObjLoader.Sub_SubstrateExistCheck(System1)
                    If (RtnCode <> cFRS_NORMAL) Then                              ' �G���[ ?
                        RtnCode = cFRS_ERR_RST
                        Return RtnCode
                    End If
                Else
                    'V2.2.0.023��
                    ' �d�����b�N(�ω����E�����b�N)�����b�N����
                    RtnCode = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_ON)                             ' �d�����b�N
                    If (RtnCode <> cFRS_NORMAL) Then                                  ' �G���[ ?(���b�Z�[�W�͕\����)
                        RtnCode = cFRS_ERR_RST
                        Return RtnCode
                    End If
                    'V2.2.0.023��
                End If
            End If
            ' 'V2.2.0.0�D ��


            Return (cFRS_NORMAL)                                        ' Return�l = ����

            '-------------------------------------------------------------------
            '   ���b�Z�[�W�\����G���[�߂�
            '-------------------------------------------------------------------
STP_ERR_EXIT:
            MsgBox(strMSG, MsgBoxStyle.Exclamation)
            Return (RtnCode)                                            ' Return�l = �`�F�b�N�G���[

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "CmdExec_Check() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cFRS_ERR_RST)                                       ' Return�l = �`�F�b�N�G���[(Cancel(RESET��)��Ԃ�)
        End Try
    End Function
#End Region

#Region "�R�}���h���s�㏈��"
    '''=========================================================================
    '''<summary>�R�}���h���s�㏈��</summary>
    '''<param name="gSts">(INP) ��ʃX�e�[�^�X(giAppMode�Q��)</param>
    '''<param name="sts"> (INP) �R�}���h���s�X�e�[�^�X(�G���[�ԍ�)</param>
    '''<remarks>����~�����͓��֐����ſ�ċ����I������</remarks>
    '''=========================================================================
    Private Sub Sub_cmdTerm_Proc(ByRef gSts As Short, ByRef sts As Short)

        Dim r As Short
        Dim strMSG As String

        Try
            ' �e�R�}���h���s�G���[�Ȃ烁�b�Z�[�W�\��
            If (sts < cFRS_NORMAL) Then                                 ' �R�}���h���s�G���[ ?
                If (sts = cFRS_ERR_PTN) Then                            ' �ȉ��̃g���~���ONG���̃G���[�̓\�t�g�����I�����Ȃ�
                ElseIf (sts = cFRS_TRIM_NG) Then                        '  
                ElseIf (sts = cFRS_ERR_TRIM) Then
                ElseIf (sts = cFRS_ERR_PT2) Then
                ElseIf (sts = cFRS_FNG_CPOS) Then                       ' �J�b�g�ʒu�␳�Ώۂ̒�R���Ȃ� ?
                    GoTo STP_END
                ElseIf (sts <= cFRS_ERR_EMG) Then                       ' ��ċ����I�� ?
                    ' �N�����v/�z��OFF
                    r = Me.System1.ClampVacume_Ctrl(gSysPrm, 0, giAppMode, giTrimErr)
                    GoTo STP_ERR_EXIT                                   ' ��ċ����I��
                End If
            End If

            ' �e�[�u�����_�ړ�
            r = Me.System1.Form_Reset(cGMODE_ORG_MOVE, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
            If (r <= cFRS_ERR_EMG) Then GoTo STP_ERR_EXIT '             ' �G���[�Ȃ��ċ����I��

            'V2.1.0.1�@            ' �Ǝ����_���A(���������͢����+����������͢�蓮�ŕ␳�Ȃ��Ń����Ұ�=���_���A�w�裎�) ���Ƃ���̏ꍇ
            'V2.1.0.1�@            If ((stThta.iPP30 = 0) And (gSysPrm.stDEV.giTheta <> 0)) Or _
            'V2.1.0.1�@               ((stThta.iPP30 = 2) And (gSysPrm.stDEV.giTheta <> 0)) Or _
            'V2.1.0.1�@               ((stThta.iPP30 = 1) And (stThta.iPP31 = 0) And (gSysPrm.stSPF.giThetaParam = 1) And (gSysPrm.stDEV.giTheta <> 0)) Then
            'V2.1.0.1�@                Call ROUND4(0.0#)                                       ' �Ƃ����_�ɖ߂�
            'V2.1.0.1�@            End If
            Call ROUND4(0.0#)                                           'V2.1.0.1�@ �Ƃ����_�ɖ߂�

            ' �ײ�޶�ް���������
            If (gSysPrm.stSPF.giWithStartSw = 0) Then                    ' ����SW�����҂�(��߼��)�łȂ� ?
                r = System1.Z_COPEN(gSysPrm, giAppMode, giTrimErr, False)
                If (r <= cFRS_ERR_EMG) Then GoTo STP_ERR_EXIT '         ' �G���[�Ȃ��ċ����I��
            Else
                ' ����ۯ����Ȃ�ײ�޶�ް�J�҂�
                If (Me.System1.InterLockSwRead() And BIT_INTERLOCK_DISABLE) = 0 Then
                    r = Me.System1.Form_Reset(cGMODE_OPT_END, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
                    If (r <= cFRS_ERR_EMG) Then GoTo STP_ERR_EXIT '     ' �G���[�Ȃ��ċ����I��
                End If
            End If

STP_END:
            ' �{�^������\������
            cmdHelp.Visible = True                                      ' Version�{�^���\�� 
            Me.Grpcmds.Visible = True                                   ' �R�}���h�{�^���O���[�v�{�b�N�X�\��
            Me.GrpMode.Visible = True                                   ' �f�B�W�^��SW�O���[�v�{�b�N�X�\��
            Me.frmInfo.Visible = True                                   ' ���ʕ\����\��
            BtnStartPosSet.Enabled = True                               'V2.0.0.0�A
            gbInitialized = True                                        ' True=���_���A��
            'V2.1.0.0�A��
            If UserSub.IsLaserCaribrarionUse() Then
                ButtonLaserCalibration.Visible = True
            End If
            'V2.1.0.0�A��
            'V2.2.0.0�E��
            If giCutStop <> 0 Then
                btnCutStop.Visible = True
            End If
            'V2.2.0.0�E��
            'V2.2.0.0�D��
            If giLoaderType <> 0 Then
                btnLoaderInfo.Visible = True
                ' �d�����b�N(�ω����E�����b�N)����������
                r = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_OFF)
            End If
            'V2.2.0.0�D��

            SetMagnifyBar(False)                                         ' V2.2.0.0�@

            Exit Sub

            ' ��ċ����I������
STP_ERR_EXIT:
            Call AppEndDataSave()                                       ' ��ċ����I�������ް��ۑ��m�F
            Call AplicationForcedEnding()                               ' ��ċ����I������
            End

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "Sub_cmdTerm_Proc() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�R�}���h�I������"
    '''=========================================================================
    '''<summary>�R�}���h�I������</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub Sub_cmdEnd_Proc()

        Dim r As Short
        Dim strMSG As String

        Try
            ' �N�����v/�z��OFF
            r = Me.System1.ClampVacume_Ctrl(gSysPrm, 0, giAppMode, 1, False)
            r = Me.System1.SetSignalTower(0, &HFFFFS)                   ' �V�O�i���^���[����(On=�Ȃ�, Off=�S�ޯ�) 

            ' �㏈��
            Call Me.System1.sLampOnOff(LAMP_START, True)                ' START����ON
            Call Me.System1.sLampOnOff(LAMP_RESET, True)                ' RESET����ON
            Call Me.System1.sLampOnOff(LAMP_Z, False)                   ' PRB����OFF

            Call ZCONRST()                                              ' �ݿ�ٷ� ׯ�����
            giAppMode = APP_MODE_IDLE                                   ' ����Ӱ�� = �g���}���u�A�C�h����
            Call Me.System1.Ilum_Ctrl(gSysPrm, Z0, ZOPT)                 ' ���Ȱ������ߏ���(����L���L��)
            Timer1.Enabled = True                                       ' �Ď��^�C�}�[�J�n
            Call Sub_ATLDSET(0, COM_STS_TRM_STATE)                      ' ���[�_�[�o��(ON=�Ȃ�,OFF=�g���}���쒆)

            ' 
            Call Z_PRINT(" " & vbCrLf)

            ChkLoaderInfoDisp(1)                              'V2.2.0.0�D

            Exit Sub

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "Sub_cmdEnd_Proc() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

    '========================================================================================
    '   ���̑��̃{�^���������̏���
    '========================================================================================
#Region "LOG�{�^������������"
    '''=========================================================================
    ''' <summary>LOG�{�^������������</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub CmdLog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim ExeFile As String
        Dim strFNAME As String
        Dim r As Double
        Dim strMSG As String

        Try
            ' �y̧�ق��J���z�޲�۸ސݒ�
            strFNAME = ""
            FileDlgOpen.FileName = ""
            FileDlgOpen.ShowReadOnly = False
            FileDlgOpen.CheckFileExists = True
            FileDlgOpen.CheckPathExists = True
            FileDlgOpen.InitialDirectory = "C:\TRIMDATA\LOG"
            FileDlgOpen.Filter = "*.LOG|*.LOG"

            ' �y̧�ق��J���z�޲�۸ޕ\��
            If (FileDlgOpen.ShowDialog() = Windows.Forms.DialogResult.OK) Then
                If (FileDlgOpen.FileName = "") Then Exit Sub
                strFNAME = FileDlgOpen.FileName                         ' ���O�t�@�C�����ݒ�
            End If

            '    ' ���O�t�@�C�����Ȃ����NOP
            '    If (gsLogFileName = "") Then Exit Sub                  ' ���O�t�@�C�����Ȃ����NOP
            '    strFName = gsLogFileName

            ' NOTEPAD�Ń��O�t�@�C�����J��
#If cOFFLINEcDEBUG Then
            ExeFile = "notepad.exe " & strFNAME
#Else
            'ExeFile = "C:\WINNT\system32\notepad.exe " + strFNAME
            ExeFile = "notepad.exe " & strFNAME
#End If
            r = Shell(ExeFile, 1)

Cansel:

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "CmdLog_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "Clear�{�^������������"
    '''=========================================================================
    ''' <summary>Clear�{�^������������</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub CmdClr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim r As Short
        Dim strMSG As String

        Try
            ' �N���A�m�F���b�Z�[�W�ݒ�
            If (gSysPrm.stTMN.giMsgTyp = 0) Then
                strMSG = "   ���Y�������������܂����H   "
            Else
                strMSG = "   Are you sure to Clear Trimming Result ?   "
            End If

            ' �N���A�m�F���b�Z�[�W��\������
            r = Me.System1.TrmMsgBox(gSysPrm, strMSG, MsgBoxStyle.OkCancel, cAPPcTITLE)

            ' Cancel(RESET��))�Ȃ�EXIT
            If (r = cFRS_ERR_RST) Then Exit Sub

            ' ���Y���N���A
            Call Disp_frmInfo(COUNTER.PRODUCT_INIT, COUNTER.NONE)                                    ' ���Y��������(frmInfo��ʂ��ĕ\��)
            Call PutLotInf()                                            ' ���b�g���Z�[�u

            'V2.0.0.0�H��
            If (Not gObjFrmDistribute Is Nothing) Then                  ' ���z�}�f�[�^�N���A
                gObjFrmDistribute.ClearCounter()
            End If
            'V2.0.0.0�H��

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "CmdClr_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�o�[�W�����_�C�A���O�{�b�N�X�̕\��"
    '''=========================================================================
    ''' <summary>�o�[�W�����_�C�A���O�{�b�N�X�̕\��</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
        Dim r As Short
        Dim pstPRM As DllAbout.HelpVersion.HelpVer_PARAM            ' �o�[�W�������\���֐��p�p�����[�^(OCX�Œ�`)
        Dim strVER(3) As String
        Dim strMSG As String
        Dim EqType As String ''V2.2.0.0�I

        Try
            ' �{�^�������\���ɂ���
            cmdHelp.Visible = False                                     ' Version�{�^����\�� 
            Me.Grpcmds.Visible = False                                  ' �R�}���h�{�^���O���[�v�{�b�N�X��\��
            Me.GrpMode.Visible = False                                  ' �f�B�W�^��SW�O���[�v�{�b�N�X��\��
            Me.frmInfo.Visible = False                                  ' ���ʕ\�����\��

            ' �\����pstPRM�̔z��̏����� ���z��̗v�f����OcxAbout.ocx�Œ�`�Ɠ����ɂ���K�v����
            pstPRM.strTtl = New String(4) {}
            pstPRM.strModule = New String(20) {}
            pstPRM.strVer = New String(20) {}


            'V2.2.0.0�I��
            Dim strVersion = GetPrivateProfileString_S("TMENU", "VERSION_NAME", SYSPARAMPATH, "")
            If strVersion.ToString().Trim <> "" Then
                EqType = strVersion
            Else
                EqType = gSysPrm.stTMN.gsKeimei
            End If
            'V2.2.0.0�I��


            ' �o�[�W�������\���֐��p�p�����[�^��ݒ肷��
            pstPRM.iTtlNum = 3                              ' �^�C�g��������̐��@'V2.2.0.0�I
            pstPRM.strTtl(0) = My.Application.Info.Title    ' �A�v���� 
            pstPRM.strTtl(1) = "LMP-" + EqType + gSysPrm.stDEV.gsDevice_No + "-000 " +
                               My.Application.Info.Version.Major.ToString("0") & "." &
                               My.Application.Info.Version.Minor.ToString("0") & "." &
                               My.Application.Info.Version.Build.ToString("0") & "." &
                               My.Application.Info.Version.Revision.ToString("0")
            pstPRM.strTtl(2) = "(c) TOWA LASERFRONT CORP."

            pstPRM.iVerNum = 15                             ' �o�[�W�������̐�
            pstPRM.strModule(0) = "RT MODULE"               ' 1."RT MODULE"
            pstPRM.strVer(0) = DLL_PATH + "INTRIM_SL432.rta"
            pstPRM.strModule(1) = "DllTrimFnc.dll"          ' 2."DllTrimFnc.dll"
            pstPRM.strVer(1) = DLL_PATH + pstPRM.strModule(1)
            pstPRM.strModule(2) = "DllSysPrm.dll"           ' 3."DllSysPrm.dll"
            pstPRM.strVer(2) = DLL_PATH + pstPRM.strModule(2)
            pstPRM.strModule(3) = "DllSystem.dll"           ' 4."DllSystem.dll"
            pstPRM.strVer(3) = DLL_PATH + pstPRM.strModule(3)
            pstPRM.strModule(4) = "DllAbout.dll"            ' 5."DllAbout.dll"
            pstPRM.strVer(4) = DLL_PATH + pstPRM.strModule(4)
            pstPRM.strModule(5) = "DllUtility.dll"          ' 6."DllUtility.dll"
            pstPRM.strVer(5) = DLL_PATH + pstPRM.strModule(5)
            pstPRM.strModule(6) = "DllLaserTeach.dll"       ' 7."DllLaserTeach.dll"
            pstPRM.strVer(6) = DLL_PATH + pstPRM.strModule(6)
            pstPRM.strModule(7) = "DllManualTeach.dll"      ' 8."DllManualTeach.dll"
            pstPRM.strVer(7) = DLL_PATH + pstPRM.strModule(7)
            pstPRM.strModule(8) = "DllPassword.dll"         ' 9."DllPassword.dll"
            pstPRM.strVer(8) = DLL_PATH + pstPRM.strModule(8)
            pstPRM.strModule(9) = "DllProbeTeach.dll"       '10."DllProbeTeach.dll"
            pstPRM.strVer(9) = DLL_PATH + pstPRM.strModule(9)
            pstPRM.strModule(10) = "DllTeach.dll"           '11."DllTeach.dll"
            pstPRM.strVer(10) = DLL_PATH + pstPRM.strModule(10)
            pstPRM.strModule(11) = "DllVideo.dll"           '12."DllVideo.dll"
            pstPRM.strVer(11) = DLL_PATH + pstPRM.strModule(11)

            ' �VDll(C#�ō쐬) 
            pstPRM.strModule(12) = "DllSerialIO.dll"        '13."DllSerialIO.dll"
            pstPRM.strVer(12) = DLL_PATH + pstPRM.strModule(12)
            pstPRM.strModule(13) = "DllCndXMLIO.dll"        '14."DllCndXMLIO.dll"
            pstPRM.strVer(13) = DLL_PATH + pstPRM.strModule(13)
            pstPRM.strModule(14) = "DllFLCom.dll"           '15."DllFLCom.dll"
            pstPRM.strVer(14) = DLL_PATH + pstPRM.strModule(14)

            ' �o�[�W�������\���ʒu��ݒ肷��
            HelpVersion1.Left = Text2.Location.X            ' Left = Text4�ʒu 
            HelpVersion1.Top = cmdHelp.Location.Y           ' Top  = Version���݈ʒu 

            ' �o�[�W�������\��
            HelpVersion1.Visible = True
            HelpVersion1.BringToFront()                     ' �őO�ʂ֕\��

            r = HelpVersion1.Version_Disp(pstPRM)
            HelpVersion1.Visible = False

            ' �{�^������\������
            cmdHelp.Visible = True                                      ' Version�{�^���\�� 
            Me.Grpcmds.Visible = True                                   ' �R�}���h�{�^���O���[�v�{�b�N�X�\��
            Me.GrpMode.Visible = True                                   ' �f�B�W�^��SW�O���[�v�{�b�N�X�\��
            Me.frmInfo.Visible = True                                   ' ���ʕ\����\��

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "cmdHelp_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "ADJ�{�^������������"
    '''=========================================================================
    '''<summary>ADJ�{�^������������</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub BtnADJ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnADJ.Click

        Dim strMSG As String

        Try
            If (BtnADJ.Text = "ADJ OFF") Then
                BtnADJ.Text = "ADJ ON"
                BtnADJ.BackColor = System.Drawing.Color.Yellow
                gbChkboxHalt = True
            Else
                BtnADJ.Text = "ADJ OFF"
                BtnADJ.BackColor = System.Drawing.SystemColors.Control
                gbChkboxHalt = False
            End If
            BtnADJ.Refresh()

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "BtnADJ_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
    ''' <summary>
    ''' �`�c�i�{�^���̂n�m�A�n�e�e�擾
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetBtnADJStatus() As Boolean
        Return (gbChkboxHalt)
    End Function
#End Region

#Region "Expansion���݉���������(۸މ�ʊg��)"
    '''=========================================================================
    ''' <summary>Expansion���݉���������(۸މ�ʊg��)</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub cmdExpansion_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExpansion.Click

        'If cmdExpansion.Text = "�g��\��" Or cmdExpansion.Text = "Expansion" Then
        '    ' �g��۸މ��
        '    txtLog.Top = 40
        '    txtLog.Height = 689
        '    cmdExpansion.Top = txtLog.Top - 22
        '    If (gSysPrm.stTMN.giMsgTyp = 0) Then
        '        cmdExpansion.Text = "�ʏ�\��"
        '    Else
        '        cmdExpansion.Text = "Normal"
        '    End If
        '    txtLog.Font = VB6.FontChangeSize(txtLog.Font, 10)
        '    txtLog.BringToFront()

        'Else
        '    ' �ʏ�۸މ��
        '    txtLog.Top = 544
        '    txtLog.Height = 192
        '    cmdExpansion.Top = txtLog.Top - 22
        '    If (gSysPrm.stTMN.giMsgTyp = 0) Then
        '        cmdExpansion.Text = "�g��\��"
        '    Else
        '        cmdExpansion.Text = "Expansion"
        '    End If
        '    txtLog.Font = VB6.FontChangeSize(txtLog.Font, gSysPrm.stLOG.gdLogTextFontSize)
        '    txtLog.SendToBack()
        'End If

        If cmdExpansion.Text = "�g��\��" Or cmdExpansion.Text = "Expansion" Then
            ' �g��۸މ��
            lstLog.Top = 40
            lstLog.Height = 689
            cmdExpansion.Top = lstLog.Top - 22
            If (gSysPrm.stTMN.giMsgTyp = 0) Then
                cmdExpansion.Text = "�ʏ�\��"
            Else
                cmdExpansion.Text = "Normal"
            End If
            lstLog.Font = VB6.FontChangeSize(lstLog.Font, 10)
            lstLog.BringToFront()

        Else
            ' �ʏ�۸މ��
            lstLog.Top = 544
            lstLog.Height = 192
            cmdExpansion.Top = lstLog.Top - 22
            If (gSysPrm.stTMN.giMsgTyp = 0) Then
                cmdExpansion.Text = "�g��\��"
            Else
                cmdExpansion.Text = "Expansion"
            End If
            lstLog.Font = VB6.FontChangeSize(lstLog.Font, gSysPrm.stLOG.gdLogTextFontSize)
            lstLog.SendToBack()
        End If

    End Sub
#End Region

#Region "Expansion����(�L��/����)"
    '''=========================================================================
    ''' <summary>Expansion����(�L��/����)</summary>
    ''' <param name="MODE">True=�L��, False=����</param>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub ExpansionOnOff(ByRef MODE As Boolean)

        If gSysPrm.stSPF.giDispCh = 1 Then
            If Not (Me.cmdExpansion.Visible) = MODE Then
                Me.cmdExpansion.Visible = MODE
                If MODE = True Then Me.cmdExpansion.BackColor = System.Drawing.ColorTranslator.FromOle(&HFF00)
                If MODE = False Then Me.cmdExpansion.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            End If
        End If

    End Sub
#End Region

#Region "۸މ�ʐؑ�"
    '''=========================================================================
    ''' <summary>۸މ�ʐؑ�</summary>
    ''' <param name="MODE">True=�L��, False=����</param>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub LogWindowCh(ByRef MODE As Short)

        If (gSysPrm.stSPF.giDispCh = 1) Then             ' �g��\������ ?
            If MODE = 0 Then                            ' �ʏ��� ?
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    Me.cmdExpansion.Text = "�ʏ�\��"
                Else
                    Me.cmdExpansion.Text = "Normal"
                End If

            ElseIf MODE = 1 Then                        ' �g�廲�� ?
                Me.cmdExpansion.Text = "�g��\��"
            End If
            Call Me.cmdExpansion_Click(Me.cmdExpansion, New System.EventArgs())
        End If

    End Sub
#End Region

#Region "�t�@���N�V�����L�[����������"
    '''=========================================================================
    ''' <summary>�t�@���N�V�����L�[����������</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub Form1_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown

#If cKEYBOARDcUSE Then

        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000

        Dim ShiftDown As Boolean
        Dim AltDown As Boolean
        Dim CtrlDown As Boolean

        ' �g���}���u�A�C�h�����łȂ����NOP
        If giAppMode Then
            If Not ((gbAdjOnStatus And KeyCode = System.Windows.Forms.Keys.F11) Or KeyCode = System.Windows.Forms.Keys.F12) Then
                'V2.2.0.032
                If giAppMode = APP_MODE_FINEADJ Or giAppMode = APP_MODE_TEACH Then
                    If (_jogKeyDown IsNot Nothing) Then         'V6.0.0.0�I
                        _jogKeyDown.Invoke(eventArgs)
                    End If
                End If
                Exit Sub
            End If
        End If

        ShiftDown = (Shift And VB6.ShiftConstants.ShiftMask) > 0
        AltDown = (Shift And VB6.ShiftConstants.AltMask) > 0
        CtrlDown = (Shift And VB6.ShiftConstants.CtrlMask) > 0

        Select Case KeyCode
            Case System.Windows.Forms.Keys.F1
                Call cmdLotInfo_Click(cmdLotInfo, New System.EventArgs())   ' �f�[�^�ݒ�

            Case System.Windows.Forms.Keys.F2
                If (stFNC(F_LOAD).iDEF = 0) Then Exit Sub ' �I��s�Ȃ�EXIT
                Call cmdLoad_Click(cmdLoad, New System.EventArgs()) ' �f�[�^���[�h

            Case System.Windows.Forms.Keys.F3
                If (stFNC(F_SAVE).iDEF = 0) Then Exit Sub ' �I��s�Ȃ�EXIT
                Call cmdSave_Click(cmdSave, New System.EventArgs()) ' �f�[�^�Z�[�u

            Case System.Windows.Forms.Keys.F4
                If (stFNC(F_EDIT).iDEF = 0) Then Exit Sub ' �I��s�Ȃ�EXIT
                Call cmdEdit_Click(cmdEdit, New System.EventArgs()) ' EDIT

            Case System.Windows.Forms.Keys.F5
                Call cmdLotchg_Click(cmdLotChg, New System.EventArgs())   ' �����^�]
                'Call cmdPrint_Click(cmdPrint, New System.EventArgs())   ' ���

            Case System.Windows.Forms.Keys.F6
                If (stFNC(F_PROBE).iDEF = 0) Then Exit Sub ' �I��s�Ȃ�EXIT
                Call cmdProbeTeaching_Click(cmdProbeTeaching, New System.EventArgs())   ' �v���[�u�ʒu�e�B�[�`���O
                'If (stFNC(F_LASER).iDEF = 0) Then Exit Sub ' �I��s�Ȃ�EXIT
                'Call cmdLaserTeach_Click(cmdLaserTeach, New System.EventArgs()) ' ���[�U

            Case System.Windows.Forms.Keys.F7
                If (stFNC(F_TEACH).iDEF = 0) Then Exit Sub ' �I��s�Ȃ�EXIT
                Call cmdTeaching_Click(cmdTeaching, New System.EventArgs()) ' �e�B�[�`���O(F8)

            Case System.Windows.Forms.Keys.F8
                If (stFNC(F_CUTPOS).iDEF = 0) Then Exit Sub ' �I��s�Ȃ�EXIT
                Call cmdCutPosTeach_Click(cmdCutPosTeach, New System.EventArgs()) ' �J�b�g�ʒu�e�B�[�`���O

            Case System.Windows.Forms.Keys.F9
                If (stFNC(F_RECOG).iDEF = 0) Then Exit Sub ' �I��s�Ȃ�EXIT
                Call BtnRECOG_Click(BtnRECOG, New System.EventArgs()) ' �p�^�[���o�^

            Case System.Windows.Forms.Keys.F10
                '                Call cmdExit_Click(cmdExit, New System.EventArgs()) ' END(F11)

            Case System.Windows.Forms.Keys.F11
                CbDigSwL.Focus()        ' MoveMode �̕ύX
                If CbDigSwL.SelectedIndex >= CbDigSwL.Items.Count - 1 Then
                    CbDigSwL.SelectedIndex = 0
                Else
                    CbDigSwL.SelectedIndex = CbDigSwL.SelectedIndex + 1
                End If

            Case System.Windows.Forms.Keys.F12
                Call BtnADJ_Click(eventSender, eventArgs)           ' ADJ ON/OFF
        End Select

        If (_jogKeyDown IsNot Nothing) Then         'V6.0.0.0�I
            _jogKeyDown.Invoke(eventArgs)
        End If

#End If
    End Sub
#End Region

#Region "�L�[�A�b�v������"          'V2.2.0.032
    Private Sub Form1_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        If (_jogKeyUp IsNot Nothing) Then
            _jogKeyUp.Invoke(e)
        End If
    End Sub
#End Region


    '========================================================================================
    '   �^�C�}�[�C�x���g����
    '========================================================================================
#Region "�����N���^�C�}�[����"
    '''=========================================================================
    ''' <summary>�����N���^�C�}�[����</summary>
    ''' <remarks></remarks>
    '''=========================================================================

    Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick

        Dim strMSG As String                                            ' ү���ޕҏW��
        Dim r As Short
        Dim iRtn As Integer
        Dim swStatus As Integer
        Dim interlockStatus As Integer
        Dim sldCvrSts As Integer
        Dim LdIDat As UInteger                          'V1.2.0.0�C
        Dim iAppMode As Short                           'V1.2.0.0�C ����Ӱ��
        Dim bAutoLoaderAuto As Boolean = False          'V2.0.0.0�N �I�[�g���[�_�[�����蓮�t���O
        Dim coverSts As Long                            'V2.2.0.0�D

        Try
            '---------------------------------------------------------------------------
            '   ��������
            '---------------------------------------------------------------------------
            Timer1.Enabled = False                                      ' �Ď��^�C�}�[��~

            ' �����/LOAD/SAVE/EDIT/LOTCHG����ވȊO�̏ꍇ(OCX�g�p�����)����ϰ��~���Ă��̂܂ܔ�����
            ' OCX����Ԃ��Ă�����Ď���ϰ���J�n����
            If (giAppMode <> APP_MODE_IDLE) And (giAppMode <> APP_MODE_LOAD) And
               (giAppMode <> APP_MODE_SAVE) And (giAppMode <> APP_MODE_LOTCHG) And
               (giAppMode <> APP_MODE_EDIT) And (giAppMode <> APP_MODE_LOTNO) Then
                Call ZCONRST()                                          ' �R���\�[���L�[���b�`����
                Exit Sub
            End If

            '---------------------------------------------------------------------------
            '   �Ď������J�n
            '---------------------------------------------------------------------------		
            ' ����~���`�F�b�N(�g���}���u�A�C�h����)
            r = Me.System1.Sys_Err_Chk_EX(gSysPrm, giAppMode)           ' ����~/��ް/�����/�W�o�@/Ͻ�����������
            If (r <> cFRS_NORMAL) Then                                  ' ����~�����o ?
                GoTo TimerErr                                           ' �A�v�������I��
            End If

            '---------------------------------------------------------------------------
            '   �C���^�[���b�N��Ԏ擾
            '---------------------------------------------------------------------------
            r = DispInterLockSts()                                      ' �C���^�[���b�N��Ԃ̕\��/��\��
            ''V2.2.0.0�D��
            '            If (r = INTERLOCK_STS_DISABLE_FULL) Then                    ' �C���^�[���b�N�S���� ?
            If (r <> INTERLOCK_STS_DISABLE_NO) Then                    ' �C���^�[���b�N�����łȂ� ?
                ' �C���^�[���b�N�����X�C�b�`ON�ŁA�J�o�[�ُ͈�
                '    If (System1.InterLockSwRead() And BIT_COVER_CLOSE) Then
                '#If cATLcDEN = 1 Then
                '				Call Sub_ATLDSET(0, &HFFFF)                             ' �S��OFF�Ƃ���B
                '#End If
                r = COVER_CHECK(coverSts)                           ' �Œ�J�o�[��Ԏ擾(0=�Œ�J�o�[�J, 1=�Œ�J�o�[��))
                    If (coverSts = 1) Then                              ' �Œ�J�o�[�� ?
                        ' �n�[�h�E�F�A�G���[(�J�o�[�����Ă܂�)���b�Z�[�W�\��
                        Call System1.Form_Reset(cGMODE_ERR_HW, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
                        GoTo TimerErr                                       ' �A�v�������I��
                    End If

                '    End If
            End If
            ''V2.2.0.0�D��

            ' IO���j�^�\��(�f�o�b�O�p)
#If cIOcMONITORcENABLED = 1 Then                                        ' IO����\������ ?
            ObjSys.Z_ATLDGET(LdIDat)                                        ' ���[�_�[����
            If (gwPrevHcmd <> LdIDat) Then                                  ' �O��f�[�^����ω����������H
                Call IoMonitor(LdIDat, 0)                                   ' IO����\��
                gwPrevHcmd = LdIDat
            End If
#End If

            Dim bChangeManual As Boolean = False                        ' V2.2.0.0�D 

            '---------------------------------------------------------------------------
            '   ���[�_�����Ȃ烍�[�_����f�[�^����͂���i�R�}���h��M�j
            '---------------------------------------------------------------------------
            If (giAppMode = APP_MODE_IDLE) Then                         ' �A�C�h����Ԏ��Ƀ`�F�b�N����
                ' ���[�_�����Ń��[�_�L�肩�烍�[�_�����̕ω��̓G���[�Ƃ���
                If (giHostMode = cHOSTcMODEcAUTO) And (gbHostConnected = False) Then
                    ' ���[�_���蓮&��~�ɐ؂�ւ��܂ő҂�
                    r = Me.System1.Form_Reset(cGMODE_LDR_ERR, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
                    If (r <= cFRS_ERR_EMG) Then GoTo TimerErr ' ����~���̃G���[�Ȃ�A�v�������I��
                End If

                ''V2.2.0.0�D��
                ' TLF�����[�_�̏ꍇ�����^�]�؂�ւ����o�͂���
                If giLoaderType = 1 Then
                    ' ���[�_����f�[�^����͂���
                    r = System1.ReadHostCommand_ForVBNET(giHostMode, gbHostConnected, giHostRun, giAppMode, pbLoadFlg)

                    If frmAutoObj.gbFgAutoOperation = True Then
                        Call Btn_Enb_OnOff(0)                               ' �{�^���񊈐���
                        bAutoLoaderAuto = True

                        ' �K���N�����v�z���L�ɂ���
                        ObjSys.setClampVaccumConfig(0)

                        ' ���b�g�؂�ւ��t���O�̃N���A
                        ObjLoader.SetLotChangeFlg(0)        'V2.2.1.1�G 

                        r = ObjLoader.LoaderGlassHandlingProc(Me.System1)

                        Dim tmptactTime As Double = ((gdTrimtime.Minutes * 60) + gdTrimtime.Seconds + (gdTrimtime.Milliseconds / 1000.0)) * 10
                        ObjSys.Sub_SetTrimmingTime(tmptactTime)

                        ' ��������ԏ�����
                        Dim dummy As Integer
                        Dim SupplyMag As Integer = 0
                        Dim SupplySlot As Integer = 0
                        Dim StoreMag As Integer = 0
                        Dim StoreSlot As Integer = 0

                        ObjSys.Sub_GetProcessTime(gitacktTime, gichangePlateTime, dummy)
                        gichangePlateTime = gitacktTime - tmptactTime
                        ObjSys.Sub_SetChangePlateTime(gichangePlateTime)
                        objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_TACT, gitacktTime)
                        objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_EXCHANGE, gichangePlateTime)
                        objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_TRIMMING, tmptactTime)

                        ObjSys.Sub_GetNowProcessMgInfo(SupplyMag, SupplySlot, StoreMag, StoreSlot)
                        objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_SUPPLY_MAGAGINE, SupplyMag)
                        objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_SUPPLY_SLOT, SupplySlot)
                        'V2.2.0.037�@objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_STORE_MAGAGINE, StoreMag)
                        'V2.2.0.037�@objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_STORE_SLOT, StoreSlot)

                        If r = cFRS_ERR_EMG Then
                            '����~

                        ElseIf r = cFRS_ERR_LOTEND Then
                            ' �����^�]�̏I��
                            Call Sub_ATLDSET(0, clsLoaderIf.LOUT_AUTO Or clsLoaderIf.LOUT_SUPLY Or clsLoaderIf.LOUT_STS_RUN Or clsLoaderIf.LOUT_REQ_COLECT Or clsLoaderIf.LOUT_DISCHRAGE)                             ' ���[�_�o��(ON=��v���܂��͋����ʒu������+��ϒ�~��+��, OFF=�����ʒu�������܂��͊�v��)
                            frmAutoObj.gbFgAutoOperation = False

                            fStartTrim = False                       ' �X�^�[�gTRIM�t���O��OFF
                            Call Me.System1.ClampVacume_Ctrl(gSysPrm, 0, giAppMode, 0)

                            ' �d�����b�N(�ω����E�����b�N)����������
                            r = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_OFF)

                            'V2.2.0.022��
                            Call UserSub.LotEnd()                           ' ���b�g�I�����̃f�[�^�o��
                            Call Printer.Print(False)                       ' ���b�g�����
                            UserBas.stCounter.LotPrint = True               ' ���b�g�I�����̈�����s�ς݂�True
                            'V2.2.0.022��

                            DispMarkAlarmList()         ' �}�[�N�󎚂̃G���[���X�g����ʂɕ\��         V2.2.1.7�B

                            ObjLoader.Loader_EndAutoDrive(Me.System1)
                            frmAutoObj.gbFgAutoOperation = False

                            ' �}�[�N�󎚂̏ꍇ�A�g���~���O�ɖ߂� V2.2.1.7�F��
                            If UserSub.IsTrimType5() = True Then
                                CbDigSwL.SelectedIndex = 0
                            End If
                            'V2.2.1.7�F��

                            'V2.2.0.0�O�� 
                            stMultiBlock.gMultiBlock = 0
                            stMultiBlock.Initialize()
                            For i As Integer = 0 To 5
                                stMultiBlock.BLOCK_DATA(i).DataNo = i + 1           ' DataNo
                                stMultiBlock.BLOCK_DATA(i).Initialize()
                                stMultiBlock.BLOCK_DATA(i).gBlockCnt = 0            ' �u���b�N��
                            Next
                            ''V2.2.0.0�O��

                        ElseIf r = cFRS_ERR_RST Then
                            ' ���f���Ŏ��̊�͏������Ȃ��B 
                            fStartTrim = False                       ' �X�^�[�gTRIM�t���O��OFF
                            ' �@���菜�����b�Z�[�W��\������
                            r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE)

                            frmAutoObj.SetAutoOpeCancel(False)               ' V2.2.1.1�A
                            Call frmAutoObj.AutoOperationEnd()              ' V2.2.1.1�A
                            frmAutoObj.gbFgAutoOperation = False

                            r = sResetStart()
                            If (r <> cFRS_NORMAL) Then                          ' �G���[ ?
                                r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE2)
                                Call AppEndDataSave()                           ' ��ċ����I�������ް��ۑ��m�F
                                Call AplicationForcedEnding()                   ' ��ċ����I������
                                End                                             ' �A�v�������I��
                                Return
                            End If
                            ' �d�����b�N(�ω����E�����b�N)����������
                            r = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_OFF)

                        ElseIf r = cFRS_ERR_LDRTO Or r = cFRS_ERR_LDR1 Or r = cFRS_ERR_LDR2 Or r = cFRS_ERR_LDR3 Then
                            ' ���f���Ŏ��̊�͏������Ȃ��B 
                            fStartTrim = False                       ' �X�^�[�gTRIM�t���O��OFF
                            ' �@���菜�����b�Z�[�W��\������
                            r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE)

                            frmAutoObj.SetAutoOpeCancel(False)               ' V2.2.1.1�A
                            Call frmAutoObj.AutoOperationEnd()              ' V2.2.1.1�A
                            frmAutoObj.gbFgAutoOperation = False

                            r = sResetStart()
                            If (r <> cFRS_NORMAL) Then                          ' �G���[ ?
                                r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE2)
                                Call AppEndDataSave()                           ' ��ċ����I�������ް��ۑ��m�F
                                Call AplicationForcedEnding()                   ' ��ċ����I������
                                End                                             ' �A�v�������I��
                                Return
                            End If
                            ' �d�����b�N(�ω����E�����b�N)����������
                            r = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_OFF)

                        ElseIf r = cFRS_NORMAL Then
                            ' ����X�e�[�W�ɍڂ����̂�1���������J�n���� 
                            fStartTrim = True                       ' �X�^�[�gTRIM�t���O��ON
                            ' Lot�؂�ւ��M���̃`�F�b�N
                            ObjSys.Z_ATLDGET(LdIDat)                                        ' ���[�_�[����
                            If LdIDat = clsLoaderIf.LINP_TRM_LOTCHANGE_START Then
                                r = clsLoaderIf.LINP_TRM_LOTCHANGE_START
                            End If

                        Else
                            ' ���f���Ŏ��̊�͏������Ȃ��B 
                            fStartTrim = False                       ' �X�^�[�gTRIM�t���O��OFF
                            Call frmAutoObj.AutoOperationEnd()              ' V2.2.1.1�A
                            frmAutoObj.gbFgAutoOperation = False

                            ' �@���菜�����b�Z�[�W��\������
                            r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE)

                            ' �d�����b�N(�ω����E�����b�N)����������
                            r = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_OFF)

                        End If
                        ' �g���~���O���s�㏈��
                        Call System1.AutoLoaderFlgReset()                       ' �I�[�g���[�_�[�t���O���Z�b�g

                        If frmAutoObj.gbFgAutoOperation = False Then
                            Call Sub_ATLDSET(0, clsLoaderIf.LINP_AUTO)                    ' ���[�_�[�o��(ON=�Ȃ�,OFF=����)
                        End If
                    Else
                        Call Btn_Enb_OnOff(1)                               ' �{�^���񊈐���
                    End If
                Else

                    ' ���ޯ�ޗp۰�ޓ���Ӱ��
#If cIOcHostComandcENABLED = 1 Then                                     ' ���ޯ�ޗp۰�ޓ���Ӱ�� ? 
        ' �z�X�g�ڑ���Ԑݒ�
        If ((gDebugHostCmd And cHSTcRDY) = cHSTcRDY) Then
            gbHostConnected = True                                      ' �z�X�g�ڑ����(True=�ڑ�(۰�ޗL))
        Else
            gbHostConnected = False                                     ' �z�X�g�ڑ����(False=���ڑ�(۰�ޖ�))
        End If
        ' ۰��Ӱ�ސݒ�
        If ((gDebugHostCmd And cHSTcAUTO) = cHSTcAUTO) Then
            giHostMode = cHOSTcMODEcAUTO                                ' ۰��Ӱ��(1:����Ӱ��)
        Else
            giHostMode = cHOSTcMODEcMANUAL                              ' ۰��Ӱ��(0:�蓮Ӱ��)
        End If
        ' ۰�ޓ��쒆�ݒ�
        If ((gDebugHostCmd And cHSTcSTATE) = cHSTcSTATE) Then
            giHostRun = 0                                               ' ۰�ޓ��쒆(0:��~)
        Else
            giHostRun = 1                                               ' ۰�ޓ��쒆(1:���쒆)
        End If
        ' ��ϰ���Đݒ�
        If ((gDebugHostCmd And cHSTcTRMCMD) = cHSTcTRMCMD) Then
            r = cHSTcTRMCMD                                             ' ��ϰ����
        End If

        ' ���ޯ�ޗp����ސݒ�
        LdIDat = gDebugHostCmd
#End If

                    ' ���[�_����f�[�^����͂���
                    r = System1.ReadHostCommand_ForVBNET(giHostMode, gbHostConnected, giHostRun, giAppMode, pbLoadFlg)


                    ' ���[�_�������œ��쒆�̓{�^���񊈐���
                    If (giHostMode = cHOSTcMODEcAUTO) And (gbHostConnected = True) Then
                        Call Btn_Enb_OnOff(0)                               ' �{�^���񊈐���
                        bAutoLoaderAuto = True                              'V2.0.0.0�N
                    Else
                        If frmAutoObj.gbFgAutoOperation Then
                            'V1.2.0.0�CAutoOperationEnd()�̒��Ɉړ� Call UserSub.SetStartCheckStatus(True)          ' �ݒ��ʂ̊m�F�L����
                            Call frmAutoObj.AutoOperationEnd()
                            frmAutoObj.gbFgAutoOperation = False            ' �����^�]�I��
                            Me.AutoRunnningDisp.BackColor = System.Drawing.Color.Yellow  ' �w�i�F = ���F
                            Me.AutoRunnningDisp.Text = "�����^�]������"
                        End If
                        Call Btn_Enb_OnOff(1)                               ' �{�^��������
                        bAutoLoaderAuto = False                             'V2.0.0.0�N
                    End If

                    ' ���[�_�����M�����f�[�^���`�F�b�N����
                    If gbHostConnected = True And r >= 0 Then
                        If giHostMode = cHOSTcMODEcAUTO Then                ' ���۰�ގ��� ?
                            Select Case r
                                Case cHSTcTRMCMD                            ' �R�}���h���g���}���H�w���Ȃ�
                                    fStartTrim = True                       ' �X�^�[�gTRIM�t���O��ON
                                    Call System1.OperationLogging(gSysPrm, MSG_OPLOG_TRIMST & "[" & DGL.ToString & "]" & Me.AutoRunnningDisp.Text, "HOSTCMD")
                                Case cHSTcLOTCHANGE                         ' �R�}���h�����b�g�؂�ւ����g���}���H�w���Ȃ�
                                    fStartTrim = True                       ' �X�^�[�gTRIM�t���O��ON
                                    Call System1.OperationLogging(gSysPrm, MSG_OPLOG_LOTCHG & "[" & DGL.ToString & "]" & Me.AutoRunnningDisp.Text, "HOSTCMD")
                            End Select
                        End If
                    End If
                    'V1.2.0.0�C��
                    ObjSys.Z_ATLDGET(LdIDat)                                        ' ���[�_�[����
                    If (gwPrevHcmd <> LdIDat) Then                                  ' �O��f�[�^����ω����������H
                        Call IoMonitor(LdIDat, 0)                                   ' IO����\��
                        gwPrevHcmd = LdIDat
                    End If
                    'V2.0.0.0�N                If frmAutoObj.gbFgAutoOperation And Not fStartTrim Then                     ' �����^�]���[�h�ŃA�C�h����Ԃ̎� 
                    If bAutoLoaderAuto And Not fStartTrim Then                                  ' �����^�]���[�h�ŃA�C�h����Ԃ̎� 'V2.0.0.0�N�����^�]���[�h�łȂ������^�]������
                        If gwPrevHcmd And cHSTcCLAMP_ON Then                                    ' �N�����v�J�M����M
                            If gbClampOpen Then
                                r = Me.System1.ClampCtrl(gSysPrm, 0, giTrimErr)                 ' �N�����v/�z��OFF
                                If (r = cFRS_NORMAL) Then
                                    Call Sub_ATLDSET(COM_STS_CLAMP_ON, 0)                       ' ���[�_�[�o��(ON=�ڕ������ߊJ,OFF=�Ȃ�)
                                    gbClampOpen = False
                                Else
                                    Call Z_PRINT("�N�����v�J�G���[���������܂����B�蓮�ɐ؂�ւ��Ă��������B" & vbCrLf)
                                    Call System.Threading.Thread.Sleep(1000)                    ' Wait(ms)
                                End If
                            End If

                        End If
                        If gwPrevHcmd And cHSTcABS_ON Then                                      ' �z���I�t�M����M
                            If gbVaccumeOff Then
                                Call Me.System1.AbsVaccume(gSysPrm, 0, giAppMode, giTrimErr)    ' �o�L���[���̐���(1=�z��ON, 0=�z��OFF)
                                Call Me.System1.Adsorption(gSysPrm, 0)                          ' �z���j�󐧌�(1:�z��, 0:�z���j��)
                                Call Sub_ATLDSET(COM_STS_ABS_ON, 0)                             ' ���[�_�[�o��(ON=�z��:�I�t,OFF=�Ȃ�)
                                gbVaccumeOff = False
                            End If
                        End If
                    End If
                    'V1.2.0.0�C��

                End If

            End If

            'Dim bChangeManual As Boolean = False
            Dim iLotChg As Integer

            'V2.2.1.1�G��
            Dim lotcnt As Integer = ObjLoader.GetLotChangeFlg()
            'V2.2.1.1�G iLotChg = IsLotChange(giHostMode, r, fStartTrim)
            iLotChg = IsLotChange(giHostMode, r, fStartTrim, lotcnt)
            'V2.2.1.1�G��

            If iLotChg > 0 Then                                         ' ���b�g�؂�ւ������̏ꍇ
                If frmAutoObj.gbFgAutoOperation Then                    ' �����^�]��
LOT_CHG:            'V2.2.1.1�G 
                    If frmAutoObj.LotChangeExecuteCheck() Then          ' ���b�g�؂�ւ��\�̏ꍇ
                        If frmAutoObj.LotChangeExecute() Then
                            stCounter.LotCounter = stCounter.LotCounter + 1

                            'V2.2.1.1�G ��
                            '�t���O�ł̃��b�g�؂�ւ���������̏ꍇ�����ōs��
                            If lotcnt >= 1 Then
                                lotcnt = lotcnt - 1

                                ObjLoader.SetLotChangeFlg(lotcnt)
                                ' ���b�g�؂�ւ��񐔕��؂�ւ����s��
                                If lotcnt > 0 Then
                                    GoTo LOT_CHG
                                End If
                            End If
                            'V2.2.1.1�G ��

                            If Not UserBas.LotStartSetting() Then       ' ���b�g�X�^�[�g�������i����f�[�^�w�b�_�[���쐬�Ȃǁj
                                bChangeManual = True
                            End If
                            'V2.1.0.0�C��
                        Else
                            bChangeManual = True
                            'V2.1.0.0�C��
                        End If
                    Else
                        Call Z_PRINT("���b�g�؂�ւ��M�����󂯂܂������A���̃��b�g�̓G���g���[����Ă��܂���B" & vbCrLf)
                        bChangeManual = True
                    End If
                End If
            End If

            If bChangeManual Then
                fStartTrim = False                              ' �X�^�[�gTRIM�t���O��OFF
                Call frmAutoObj.AutoOperationEnd()
                'V1.2.0.0�CAutoOperationEnd()�̒��Ɉړ�                 Call UserSub.SetStartCheckStatus(True)          ' �ݒ��ʂ̊m�F�L����
                Buzzer()                                        'V1.1.0.1�A �I�����u�U�[
                'V1.1.0.1�A System1.SetSignalTower(SIGOUT_RED_ON Or SIGOUT_BZ1_ON, &HFFFFS)
                'V1.2.0.0�C Call Me.System1.TrmMsgBox(gSysPrm, "���b�g�؂�ւ��G���[���������܂����B", MsgBoxStyle.OkOnly, My.Application.Info.Title)
                'V1.1.0.1�A System1.SetSignalTower(0, SIGOUT_RED_ON Or SIGOUT_BZ1_ON)

                ''V2.2.0.0�D��
                ' TLF�����[�_�̏ꍇ�����^�]�؂�ւ����o�͂���
                If giLoaderType = 1 Then
                    Call Sub_ATLDSET(0, clsLoaderIf.LINP_AUTO)                    ' ���[�_�[�o��(ON=�Ȃ�,OFF=����)
                End If

                ' ���[�_���蓮&��~�ɐ؂�ւ��܂ő҂�
                If giHostMode = cHOSTcMODEcAUTO Then
                    r = Me.System1.Form_Reset(cGMODE_LDR_CHK, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
                    If (r <= cFRS_ERR_EMG) Then
                        GoTo TimerErr                           ' ����~���̃G���[�Ȃ�A�v�������I��
                    ElseIf (r = cFRS_ERR_RST) Then              ' �L�����Z���Ȃ�R�}���h�I��
                        GoTo TimerErr
                    End If
                End If


                ''V2.2.0.0�D��
                ' TLF�����[�_�̏ꍇ�A���菜�����b�Z�[�W��\������
                If giLoaderType = 1 Then
                    r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE)
                    frmAutoObj.SetAutoOpeCancel(False)               ' V2.2.1.1�A
                    Call frmAutoObj.AutoOperationEnd()              ' V2.2.1.1�A

                    r = sResetStart()
                    If (r <> cFRS_NORMAL) Then                          ' �G���[ ?
                        r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE2)
                        Call AppEndDataSave()                           ' ��ċ����I�������ް��ۑ��m�F     
                        Call AplicationForcedEnding()                   ' ��ċ����I������
                        End                                             ' �A�v�������I��
                        Return
                    End If
                Else
                    'V2.2.2.0�C��
                    '#0005/#0050�̏ꍇ�����Ŗ߂�
                    ' �}�[�N�󎚂̏ꍇ�A�g���~���O�ɖ߂� V2.2.1.7�F��
                    If UserSub.IsTrimType5() = True Then
                        CbDigSwL.SelectedIndex = 0
                    End If
                    'V2.2.1.7�F��
                    'V2.2.2.0�C��

                End If
                ''V2.2.0.0�D��

            End If


            '---------------------------------------------------------------------------
            '   ���[�_�}�j���A����(����у��[�_������)�́A�ȉ��̏������s��
            '---------------------------------------------------------------------------
            If (giHostMode = cHOSTcMODEcMANUAL) Then

                r = INTERLOCK_CHECK(interlockStatus, swStatus)
                ' TLF�����[�_�̏ꍇ�����^�]�؂�ւ����o�͂���
                If giLoaderType = 0 Then                ''V2.2.0.0�D�@
                    If (r <> ERR_CLEAR) Then                                    ' �����b�Z�[�W�\����ǉ����� 
                        '➑̃J�o�[�J�̏ꍇ�A�C���^�[���b�N�������J�o�[���Ď�����
                        If (r = ERR_OPN_CVR) Then
                            iRtn = System1.Form_Reset(cGMODE_CVR_CLOSEWAIT, gSysPrm, giAppMode, False, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
                            If (iRtn <= cFRS_ERR_EMG) Then GoTo TimerErr ' ����~���̃G���[�Ȃ�A�v�������I��
                            ' �R���\�[���L�[���b�`�N���A
                            Call ZCONRST()
                        ElseIf (r = ERR_OPN_SCVR Or r = ERR_OPN_CVRLTC) Then
                            ' SL432R�̏ꍇ�̓J�o�[�J���b�`�͖�������B
                            If (gSysPrm.stTMN.gsKeimei <> MACHINE_TYPE_SL432) Then
                                'V2.2.0.0�D��
                                iRtn = System1.Sub_CoverCheck(gSysPrm, 0, False)
                                If (iRtn <= cFRS_ERR_EMG) Then GoTo TimerErr ' ����~���̃G���[�Ȃ�A�v�������I��
                                Call COVERLATCH_CLEAR()                                     ' �J�o�[�J���b�`�̃N���A
                                ' GoTo TimerErr
                                'V2.2.0.0�D��
                            End If
                        Else
                            GoTo TimerErr
                        End If
                    End If
                End If

                '---------------------------------------------------------------------------
                '   �r�s�`�q�s �r�v�̉����`�F�b�N
                '---------------------------------------------------------------------------
                If (giAppMode = APP_MODE_IDLE) Then                     ' �A�C�h�����[�h���Ƀ`�F�b�N����
                    r = START_SWCHECK(0, swStatus)                      ' �g���}�[ START SW �����`�F�b�N
                    If (swStatus = cFRS_ERR_START) Then

                        ' 'V2.2.0.0�D ��
                        If giLoaderType <> 0 Then
                            ' �ڕ���Ɋ�����鎖���`�F�b�N����(�蓮���[�h��(OPTION))
                            r = ObjLoader.Sub_SubstrateExistCheck(System1)
                            If (r <> cFRS_NORMAL) Then                              ' �G���[ ?
                                If (r = cFRS_ERR_RST) Then                          ' �����(Cancel(RESET��)�@?
                                    Timer1.Enabled = True                           ' �^�C�}�[�ċN��
                                    Exit Sub
                                End If
                                GoTo TimerErr                                       ' ���̑��̃G���[�Ȃ�A�v�������I��
                            End If
                        End If
                        ' 'V2.2.0.0�D ��

                        ' �g���}���쒆�M��ON���M(�I�[�g���[�_�[)
                        Call Sub_ATLDSET(COM_STS_TRM_STATE, COM_STS_TRM_NG Or COM_STS_PTN_NG Or COM_STS_TRM_ERR)          ' ���[�_�[�o��(ON=�g���}���쒆,OFF=)
                        Call System1.OperationLogging(gSysPrm, MSG_OPLOG_TRIMST, "START SW ON")
                        ' �N�����v/�z��ON
                        r = Me.System1.ClampCtrl(gSysPrm, 1, giTrimErr)
                        If (r <> cFRS_NORMAL) Then
                            GoTo TimerErr
                        End If
                        r = Me.System1.AbsVaccume(gSysPrm, 1, APP_MODE_TRIM, giTrimErr) ' APP_MODE_TRIM�̃��[�h�̎��̂݃o�L���[���`�F�b�N���s����B
                        If (r <> cFRS_NORMAL) Then
                            Call ZCONRST()                                              ' ���b�`����
                            Call Me.System1.Adsorption(gSysPrm, 0)
                            Call Me.System1.ClampCtrl(gSysPrm, 0, giTrimErr)
                            GoTo TimerExit
                        End If
                        Call Me.System1.Adsorption(gSysPrm, 1)

                        ' �X�^�[�gSW�𗣂��܂ő҂�INTRIM�ɂĊĎ��̖������[�v
                        Call START_SWCHECK(1, swStatus)

                        ' ����SW�����҂����Ȃ��ꍇ�̓��b�Z�[�W�\������
                        If (gSysPrm.stSPF.giWithStartSw = 0) And (interlockStatus = INTERLOCK_STS_DISABLE_NO) Then
                            ' "���ӁI�I�I�@�X���C�h�J�o�[�������ŕ��܂��B"(Red,Blue)
                            ' 'V2.2.0.0�@ r = Me.System1.Form_MsgDispStartReset(MSG_SPRASH31, MSG_SPRASH32, &HFF, &HFF0000)
                            r = Me.System1.Form_MsgDispStartReset(MSG_SPRASH31, MSG_SPRASH32, Color.Blue, Color.Red)           'V2.2.0.0�@
                            If (r = cFRS_ERR_RST) Then
                                ' RESET SW�����Ȃ�ErrorSkip��
                                GoTo TimerExit
                            End If
                        End If


                        If giLoaderType = 1 And frmAutoObj.gbFgAutoOperation = True Then
                            ObjLoader.DispLoaderInfo()
                        End If

                        ' �X���C�h�J�o�[���N���[�Y����(�蓮/����)
                        If (gSysPrm.stSPF.giWithStartSw = 1) And (giHostMode <> cHOSTcMODEcAUTO) Then
                            If (interlockStatus = INTERLOCK_STS_DISABLE_NO) Then
                                ' �X���C�h�J�o�[���b�Z�[�W�\�� (����SW�����҂�(�I�v�V����) �Ń��[�_�����^�]���łȂ��ꍇ)
                                r = Me.System1.Form_Reset(cGMODE_START_RESET, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, True)
                                If (r = cFRS_ERR_START) Then
                                    ' START SW�����Ȃ�X�^�[�gTRIM�t���OON
                                    fStartTrim = True
                                Else
                                    ' RESET SW�����Ȃ�ErrorSkip��
                                    GoTo ErrorSkip
                                End If
                            Else
                                ' START SW�����Ȃ�X�^�[�gTRIM�t���OON
                                fStartTrim = True
                            End If

                        Else
                            If (interlockStatus = INTERLOCK_STS_DISABLE_NO) Then
                                ' �X���C�h�J�o�[�������N���[�Y����
                                If gSysPrm.stTMN.giOnline = TYPE_MANUAL Then
                                    ' XY_SLIDE�������� ?(XY_SLIDE��������̓��[�_����̃X�^�[�g�v�����݂̂̂��ߒʏ퓮��Ƃ���) 
                                    r = Me.System1.Z_CCLOSE(gSysPrm, giAppMode, 0, False, 0.0, 0.0)
                                End If
                                If gSysPrm.stTMN.giOnline = TYPE_ONLINE Then
                                    ' XY_SLIDE�ʏ퓮��
                                    r = Me.System1.Z_CCLOSE(gSysPrm, giAppMode, 0, False, 0.0, 0.0)
                                End If
                            Else
                                fStartTrim = True
                            End If
                        End If

                    Else
                        '---------------------------------------------------------------------------
                        '   �X���C�h�J�o�[��Ԃ̃`�F�b�N(SL432R��)
                        '---------------------------------------------------------------------------
                        ' �C���^�[���b�N������SL432R�n�̏ꍇ�Ƀ`�F�b�N���� 
                        If (interlockStatus = INTERLOCK_STS_DISABLE_NO) And (gSysPrm.stTMN.gsKeimei = MACHINE_TYPE_SL432) Then
                            ' �X���C�h�J�o�[�̏�Ԏ擾�iINTRIM�ł�IO�擾�ׁ݂̂̈A�G���[���Ԃ鎖�͂Ȃ��j
                            r = SLIDECOVER_GETSTS(sldCvrSts)

                            ' �X���C�h�J�o�[��Ԃ̃`�F�b�N
                            If (sldCvrSts = SLIDECOVER_MOVING) Then

                                ' �X���C�h�J�o�[���ԂȂ�g���}���쒆�M��ON���M(�I�[�g���[�_�[)
                                Call Sub_ATLDSET(COM_STS_TRM_STATE, 0)  ' ���[�_�[�o��(ON=�g���}���쒆,OFF=�Ȃ�)

                                If (gfclamp = False) Then
                                    ' �X���C�h�J�o�[���ԁA�N�����vOFF�̏ꍇ�F�N�����v��ON����B
                                    ' �N�����v/�z��ON
                                    r = Me.System1.ClampCtrl(gSysPrm, 1, giTrimErr)
                                    If (r <> cFRS_NORMAL) Then
                                        GoTo TimerErr
                                    End If
                                    r = Me.System1.AbsVaccume(gSysPrm, 1, APP_MODE_TRIM, giTrimErr) ' APP_MODE_TRIM�̃��[�h�̎��̂݃o�L���[���`�F�b�N���s����B
                                    If (r <> cFRS_NORMAL) Then
                                        Call ZCONRST()                                              ' ���b�`����
                                        Call Me.System1.Adsorption(gSysPrm, 0)
                                        Call Me.System1.ClampCtrl(gSysPrm, 0, giTrimErr)
                                        GoTo TimerExit
                                    End If
                                    Call Me.System1.Adsorption(gSysPrm, 1)
                                    gfclamp = True
                                End If

                            ElseIf (sldCvrSts = SLIDECOVER_OPEN) Then
                                ' �X���C�h�J�o�[���I�[�v����ԂȂ�g���}���쒆�M��OFF���M(�I�[�g���[�_�[)
                                Call Sub_ATLDSET(0, COM_STS_TRM_STATE)  ' ���[�_�[�o��(ON=�Ȃ�, OFF=�g���}���쒆)

                                ' �X���C�h�J�o�[���I�[�v����ԂŁA�N�����vON�̏ꍇ�F�N�����v��OFF����B
                                gfclamp = False
                                ' �N�����v/�z��OFF
                                r = Me.System1.ClampCtrl(gSysPrm, 0, giTrimErr)
                                If (r <> cFRS_NORMAL) Then
                                    GoTo TimerErr
                                End If
                                Call Me.System1.AbsVaccume(gSysPrm, 0, giAppMode, giTrimErr)
                                Call Me.System1.Adsorption(gSysPrm, 0)
                            ElseIf (sldCvrSts = SLIDECOVER_CLOSE) Then
                                ' �g���}���쒆�M��ON���M(�I�[�g���[�_�[)
                                Call Sub_ATLDSET(COM_STS_TRM_STATE, 0)  ' ���[�_�[�o��(ON=�g���}���쒆,OFF=�Ȃ�)

                                ' �X���C�h�J�o�[��
                                gfclamp = False
                                Call COVERLATCH_CLEAR()                 ' ��ް�Jׯ��ر
                                Call System1.OperationLogging(gSysPrm, MSG_OPLOG_TRIMST & "[" & DGL.ToString & "]", "SLIDE COVER CLOSED")

                                ' ����m�F���(START/RESET�҂�)
                                If (gSysPrm.stSPF.giWithStartSw = 1) Then
                                    r = System1.Form_Reset(cGMODE_START_RESET, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
                                    If (r = cFRS_ERR_RST) Then
                                        ' RESET SW�����Ȃ�ErrorSkip��
                                        Call UserSub.SetStartCheckStatus(True)          ' �ݒ��ʂ̊m�F�L����
                                        GoTo ErrorSkip
                                    ElseIf (r <> cFRS_ERR_START) Then
                                        GoTo TimerErr
                                    End If
                                End If
                                fStartTrim = True                       ' �X�^�[�gTRIM�t���OON
                            End If
                        End If
                    End If
                End If
            End If

            '---------------------------------------------------------------------------
            '   �X�^�[�gTRIM�t���O��ON�Ȃ�A�ȉ��̏������s��
            '---------------------------------------------------------------------------
            If fStartTrim = True Then
                ' �f�[�^�̓��[�h�ς݂�
                If pbLoadFlg = False Then                               ' �f�[�^�����[�h ?
                    strMSG = MSG_DataNotLoad                            ' "�f�[�^�����[�h"
                    Call Z_PRINT(strMSG)                                ' ���b�Z�[�W�\��
                    GoTo ErrorSkip
                End If

                ' �g���}���u�A�C�h���łȂ����ErrorSkip��
                gfclamp = False                                         ' FLG = �N�����vOFF
                If giAppMode Then GoTo ErrorSkip '                      ' �g���}���u�A�C�h���łȂ� ?
                gbInitialized = False                                   ' flg = ���_���A��

                If Not frmAutoObj.gbFgAutoOperation Then            ' �����^�]���łȂ���
                    If UserSub.IsSpecialTrimType() Then             ' ���[�U�v���O�������ꏈ��
                        If Not UserBas.LotStartSetting() Then       ' ���b�g�X�^�[�g�������i����f�[�^�w�b�_�[���쐬�Ȃǁj
                            GoTo ErrorSkip
                        End If
                    End If
                End If

                'V2.1.0.0�A�� ���[�U�[�p���[�L�����u���[�V�����@�\��ֈړ�                Call SetATTRateToScreen(True)           '###1040�E �g���~���O�f�[�^�ł̂`�s�s�������̐ݒ�

                ' ���[�_�փf�[�^�o��
                ' 'V2.2.0.0�D �� TLF�����[�_���͑O���̌��ʂ�ۑ�
                If giLoaderType = 1 Then
                    Call Sub_ATLDSET(0, clsLoaderIf.LOUT_REDY Or clsLoaderIf.LOUT_STOP Or clsLoaderIf.LOUT_TRM_NG)
                    ObjLoader.SetLotAbort(0)

                Else
                    'V1.2.0.0�C#If cATLcDEN = 0 Then
                    ' ON=�g���}���쒆, OFF=�g���~���ONG,�p�^�[���F��NG
                    'V1.2.0.0�C                Call Sub_ATLDSET(COM_STS_TRM_STATE, COM_STS_TRM_NG Or COM_STS_PTN_NG)
                    Call Sub_ATLDSET(COM_STS_TRM_STATE, COM_STS_TRM_NG Or COM_STS_PTN_NG Or COM_STS_CLAMP_ON Or COM_STS_ABS_ON)
                    'V1.2.0.0�C#Else
                    'V1.2.0.0�C                			Call Sub_ATLDSET(COM_STS_TRM_STATE Or COM_STS_1ST_CMD, COM_STS_TRM_OK Or COM_STS_TRM_PRB Or COM_STS_TRM_NG)
                    'V1.2.0.0�C#End If
                End If
                ' 'V2.2.0.0�D ��
                '-----------------------------------------------------------------------
                '   ������ܰ����(��߼��)
                '-----------------------------------------------------------------------
                giTrimErr = 0                                           ' ��ϰ �װ �׸ޏ�����
                If (giHostMode = cHOSTcMODEcAUTO) Then                  ' ۰�ގ���Ӱ�� ?
                    ' ������ܰ����(On=�����^�]�� , Off=�S�r�b�g)
                    r = System1.SetSignalTower(SIGOUT_GRN_ON, &HFFFFS)
                Else
                    ' ������ܰ����(On=�Ȃ�,Off=�S�r�b�g)
                    r = System1.SetSignalTower(0, &HFFFFS)
                End If

                Call System1.sLampOnOff(LAMP_START, True)               ' START�����v�_��
                giAppMode = APP_MODE_TRIM                               ' ����Ӱ�� = �g���~���O��

                '-----------------------------------------------------------------------
                '   �X���C�h�J�o�[�����N���[�Y
                '-----------------------------------------------------------------------
                ' �I�[�g���[�_���� ?
                If giHostMode = cHOSTcMODEcAUTO Then                    ' ���[�_�������[�h ?
                    ' �X���C�h�J�o�[�������N���[�Y����
                    If gSysPrm.stTMN.giOnline = TYPE_MANUAL Then
                        ' XY_SLIDE�������� ?
                        r = Me.System1.Z_CCLOSE(gSysPrm, giAppMode, 0, True, gSysPrm.stDEV.gfTrimX, gSysPrm.stDEV.gfTrimY)
                    End If
                    If gSysPrm.stTMN.giOnline = TYPE_ONLINE Then
                        ' XY_SLIDE�ʏ퓮��
                        r = Me.System1.Z_CCLOSE(gSysPrm, giAppMode, 0, False, 0.0, 0.0)
                    End If
                    If (r <= cFRS_ERR_EMG) Then GoTo TimerErr '         ' ����~���̃G���[�Ȃ�A�v�������I��

                Else                                                    ' �ײ�޶�ް�����۰��
                    If (gSysPrm.stSPF.giWithStartSw = 0) Then           ' ����SW���������ݸފJ�n(��߼��)���͎����۰�ނ��Ȃ�
                        r = System1.Z_CCLOSE(gSysPrm, giAppMode, giTrimErr, False, 0, 0)
                        If (r <= cFRS_ERR_EMG) Then GoTo TimerErr '     ' ����~���̃G���[�Ȃ�A�v�������I��
                    End If
                End If

                iRtn = SetZOff_Prob_Off()                        ' INTIME�����̑ҋ@�ʒu��ύX���đҋ@�ʒu�ړ�����B###1041�@
                If iRtn <> cFRS_NORMAL Then
                    Call Me.System1.TrmMsgBox(gSysPrm, "SETZOFFPOS �y���ҋ@�ʒu�ύX���ُ�I�����܂����B", MsgBoxStyle.OkOnly, My.Application.Info.Title)
                    GoTo ErrorSkip
                End If

                ' 'V2.2.0.0�D �� TLF�����[�_���͑O���̌��ʂ�ۑ�
                If giLoaderType = 1 Then
                    swMesureTrimtime.Stop()
                    swMesureTrimtime.Reset()
                    swMesureTrimtime.Start()
                End If

                Call UserSub.ClampVacumeChange()         'V2.0.0.0�M
                'V2.1.0.0�A�� ���[�U�[�p���[�L�����u���[�V�����@�\
                '-----------------------------------------------------------------------
                '   ���[�U�[�p���[�̃��j�^�����O���s
                '-----------------------------------------------------------------------
                If UserSub.LaserCalibrationExecute() Then
                    Dim tmpiAttFix As Short = gSysPrm.stRAT.giAttFix
                    Dim tmpiAttRot As Short = gSysPrm.stRAT.giAttRot
                    Dim tmpfAttRate As Double = gSysPrm.stRAT.gfAttRate
                    If UserSub.LaserCalibrationFullPowerGet(stLASER.dblPowerAdjustTarget, stLASER.dblPowerAdjustToleLevel) Then
                        stLASER.dblPowerAdjustQRate = stLASER.intQR / 10.0#
                        r = AutoLaserPowerADJ(True)                         ' ���[�U�p���[�������s
                        If (r = cFRS_ERR_RST) Then                          ' Cancel(RESET��) ?

                            'V2.2.1.1�C��
                            If giLoaderType = 1 Then
                                ' �t���p���[�`�F�b�N���G���[�̏ꍇ�ɂ̓��b�g���I�����ă��C����ʂɖ߂� 

                                ' ���b�g�������f 
                                ' ���f���Ŏ��̊�͏������Ȃ��B 
                                fStartTrim = False                       ' �X�^�[�gTRIM�t���O��OFF
                                ' �@���菜�����b�Z�[�W��\������
                                r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE)

                                frmAutoObj.SetAutoOpeCancel(False)               ' V2.2.1.1�A
                                Call frmAutoObj.AutoOperationEnd()              ' V2.2.1.1�A
                                frmAutoObj.gbFgAutoOperation = False

                                ' ���_���A�m�F 
                                r = sResetStart()
                                If (r <> cFRS_NORMAL) Then                          ' �G���[ ?
                                    '���_���A�G���[�̏ꍇ�̓v���O�����I�� 
                                    r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE2)
                                    Call AppEndDataSave()                           ' ��ċ����I�������ް��ۑ��m�F
                                    Call AplicationForcedEnding()                   ' ��ċ����I������
                                    End                                             ' �A�v�������I��
                                    Return
                                End If
                                ' �d�����b�N(�ω����E�����b�N)����������
                                r = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_OFF)
                                GoTo ErrorSkip                              ' ��������蓮�ɐ؂�ւ������

                            Else
                                ' "���[�_�M���������ł�", "���[�_���蓮�ɐ؂�ւ��Ă�������"
                                r = Me.System1.Form_Reset(cGMODE_LDR_CHK, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
                                If (r = cFRS_NORMAL) Then
                                    UserSub.LaserCalibrationSet(POWER_CHECK_LOT) '���[�U�p���[���j�^�����O���s�L���ݒ�
                                    GoTo ErrorSkip                              ' ��������蓮�ɐ؂�ւ������
                                Else
                                    GoTo TimerErr                               ' �i"Cancel�{�^�������Ńv���O�������I�����܂�"�́AcFRS_ERR_RST�j�A�v�������I��
                                End If
                            End If
                            'V2.2.1.1�C��

                        ElseIf (r <> cFRS_NORMAL) Then                      ' �G���[ ?(���G���[���b�Z�[�W�͕\���ς�) 
                            GoTo TimerErr                                   ' �A�v�������I��
                        End If
                    Else
                        GoTo ErrorSkip                              ' ��������蓮�ɐ؂�ւ������
                    End If
                    Call ATTRESET()
                End If
                If SetATTRateToScreen(True) = False Then
                    GoTo TimerErr                                       ' �A�v�������I��
                End If
                'V2.1.0.0�A��

                'V2.2.1.7�B��
                If frmAutoObj.gbFgAutoOperation = True Then
                    MarkingCount = MarkingCount + 1             ' �}�[�L���O�p�J�E���^�N���A	
                Else ' frmAutoObj.gbFgAutoOperation = False Then
                    ' �蓮�̏ꍇ�J�E���^��1�ŁA�}�[�L���O��-1���ĊJ�n�ԍ������̂܂܃}�[�N�� 
                    MarkingCount = 1             ' �}�[�L���O�p�J�E���^�N���A	
                End If
                'V2.2.1.7�B��
                'V2.2.1.7�B��

                '-----------------------------------------------------------------------
                '   �g���~���O���s
                '-----------------------------------------------------------------------
                iRtn = User()                                           ' �g���~���O���s


                ' 'V2.2.0.0�D �� TLF�����[�_���͑O���̌��ʂ�ۑ�
                If giLoaderType = 1 And frmAutoObj.gbFgAutoOperation = True Then
                        ObjLoader.m_lTrimResult = iRtn
                        ' 'V2.2.0.0�D �� TLF�����[�_���͑O���̌��ʂ�ۑ�
                        swMesureTrimtime.Stop()
                        '' �g���~���O���ԏ����� 
                        gdTrimtime = swMesureTrimtime.Elapsed

                        '' ��������ԏ�����
                        'Dim dummy As Integer
                        Dim SupplyMag As Integer = 0
                        Dim SupplySlot As Integer = 0
                        Dim StoreMag As Integer = 0
                        Dim StoreSlot As Integer = 0

                        ObjSys.Sub_GetNowProcessMgInfo(SupplyMag, SupplySlot, StoreMag, StoreSlot)
                        objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_SUPPLY_MAGAGINE, SupplyMag)
                        objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_SUPPLY_SLOT, SupplySlot)
                        'V2.2.0.037�@objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_STORE_MAGAGINE, StoreMag)
                        'V2.2.0.037�@objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_STORE_SLOT, StoreSlot)

                        If iRtn = cFRS_ERR_RST Then
                            ObjLoader.SetLotAbort(1)
                        End If

                    End If
                    ' 'V2.2.0.0�D ��

                    If (iRtn >= cFRS_NORMAL) Then                           ' ����/RESET SW���� ?


                        ' ���ݸ�NG/��ϰ�װ/����ݔF���װ��
                    ElseIf (iRtn = cFRS_TRIM_NG) Or (iRtn = cFRS_ERR_TRIM) Or (iRtn = cFRS_ERR_PTN) Then

                        '' ➑̃J�o�[�J/�X���C�h�J�o�[�J/�J�o�[�J���b�`���o���͋����I�����Ȃ�
                        'ElseIf (iRtn = cFRS_ERR_CVR) Or (iRtn = cFRS_ERR_SCVR) Or (iRtn = cFRS_ERR_LATCH) Then

                        ' ����~���̃A�v�������I���G���[��
                    Else ' �N�����v/�z��OFF(۰�ް�ւ���ϰ���쒆OFF���Ȃ�)
                        r = System1.ClampVacume_Ctrl(gSysPrm, 0, giAppMode, 1, False)
                        Call ZSLCOVEROPEN(0)                                ' �ײ�޶�ް����������OFF
                        Call ZSLCOVERCLOSE(0)                               ' �ײ�޶�ް�۰�������OFF
                        GoTo TimerErr                                       ' �A�v�������I��
                    End If

ErrorSkip:
                    If giHostMode = cHOSTcMODEcMANUAL Then                  ' ���[�_�}�j���A���H
                        Call ZCONRST()                                      ' �ݿ��SWׯ�����
                    End If

                    '-----------------------------------------------------------------------
                    '   �X���C�h�J�o�[�����I�[�v��
                    '-----------------------------------------------------------------------

                    If (giLoaderType = 1) AndAlso (giHostMode = cHOSTcMODEcAUTO) Then
                        ' ��r�o�ʒu�ւ̈ړ� 
                        r = ObjLoader.MoveGlassOutPos()
                        If (r <> cFRS_NORMAL) Then                                  ' �G���[ ?
                            frmAutoObj.gbFgAutoOperation = False
                            Call Sub_ATLDSET(0, clsLoaderIf.LINP_AUTO)                    ' ���[�_�[�o��(ON=�Ȃ�,OFF=����)
                            GoTo TimerErr                                       ' �A�v�������I��
                        End If

                    Else
                        'If (giHostMode <> cHOSTcMODEcAUTO) Then                 ' ���[�_�������[�h�łȂ� ? 
                        If (gSysPrm.stTMN.giOnline = 1) Or ((giHostMode <> cHOSTcMODEcAUTO) And (gSysPrm.stTMN.giOnline = 2)) Then
                            r = System1.EX_SBACK(gSysPrm)                       ' �߰�����ׂ�۰�ވʒu�ɖ߂�
                            If (r <= cFRS_ERR_EMG) Then GoTo TimerErr '         ' ����~���̃G���[�Ȃ�A�v�������I��
                            'V2.1.0.1�@                    ' �Ǝ����_���A(���������͢����+����������͢�蓮�ŕ␳�Ȃ��Ń����Ұ�=���_���A�w�裎�) ���Ƃ���̏ꍇ
                            'V2.1.0.1�@                    If ((stThta.iPP30 = 0) And (gSysPrm.stDEV.giTheta <> 0)) Or _
                            'V2.1.0.1�@                       ((stThta.iPP30 = 2) And (gSysPrm.stDEV.giTheta <> 0)) Or _
                            'V2.1.0.1�@                       ((stThta.iPP30 = 1) And (stThta.iPP31 = 0) And (gSysPrm.stSPF.giThetaParam = 1) And (gSysPrm.stDEV.giTheta <> 0)) Then
                            'V2.1.0.1�@                        Call ROUND4(0.0#)                               ' �Ƃ����_�ɖ߂�
                            'V2.1.0.1�@                    End If
                        End If


                    End If

                    Call ROUND4(0.0#)                               'V2.1.0.1�@�K���Ƃ����_�ɖ߂�

                    If UserSub.IsTRIM_MODE_ITTRFT() And Not UserSub.GetStartCheckStatus() And iRtn <> cFRS_ERR_RST And r <> cFRS_ERR_RST Then
                        UserBas.stCounter.EndTime = DateTime.Now()          ' ������I�����ԕۑ� '###1030�B
                        Buzzer()                                            ' �I�����u�U�[
                    End If

                    ' �X���C�h�J�o�[�����I�[�v��
                    ' ����SW���������ݸފJ�n(��߼��)���͎�������݂��Ȃ�
                    If (gSysPrm.stSPF.giWithStartSw = 0) Or (giHostMode = cHOSTcMODEcAUTO) Then
                        'V1.2.0.0�C�� ����Ӱ�ށ@Z_COPEN��giAppMode��iAppMode�ɕύX
                        iAppMode = giAppMode
                        'V2.0.0.0�N                    If frmAutoObj.gbFgAutoOperation And giAppMode = APP_MODE_TRIM Then  '�����^�]���́AZ_COPEN���ŃN�����v�J���s��Ȃ��B
                        If bAutoLoaderAuto And giAppMode = APP_MODE_TRIM Then  '�����^�]���́AZ_COPEN���ŃN�����v�J���s��Ȃ��B
                            iAppMode = APP_MODE_TRIM_AUTO            'V2.0.0.0�N�@APP_MODE_AUTO����APP_MODE_TRIM_AUTO�֕ύX
                        End If
                        'V1.2.0.0�C��
                        If (gSysPrm.stTMN.giOnline = 1) Then                ' �ײ�޶�ް���������(XY_SLIDE�ʏ퓮��)
                            r = System1.Z_COPEN(gSysPrm, iAppMode, giTrimErr, False)
                            If (r <= cFRS_ERR_EMG) Then GoTo TimerErr '     ' ����~���̃G���[�Ȃ�A�v�������I��
                        End If
                        If (gSysPrm.stTMN.giOnline = 2) Then                ' �ײ�޶�ް���������(XY_SLIDE��������)
                            r = System1.Z_COPEN(gSysPrm, iAppMode, giTrimErr, True)
                            If (r <= cFRS_ERR_EMG) Then GoTo TimerErr '     ' ����~���̃G���[�Ȃ�A�v�������I��
                        End If

                        ' �g���~���O�I�����̶�ް�J�҂�(��߼��) (�C���^�[���b�N���̏ꍇ)
                    Else
                        If (System1.InterLockSwRead() And BIT_INTERLOCK_DISABLE) = 0 Then
                            r = System1.Form_Reset(cGMODE_OPT_END, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
                            If (r <= cFRS_ERR_EMG) Then GoTo TimerErr '     ' ����~���̃G���[�Ȃ�A�v�������I��
                            r = System1.ClampVacume_Ctrl(gSysPrm, 0, giAppMode, 1, False)
                            If (r <> cFRS_NORMAL) Then GoTo TimerErr '      ' �A�v�������I��
                        Else
                            r = System1.ClampVacume_Ctrl(gSysPrm, 0, giAppMode, 1, False)
                            If (r <= cFRS_ERR_EMG) Then GoTo TimerErr '     ' ����~���̃G���[�Ȃ�A�v�������I��
                        End If
                    End If

                    'V2.1.0.1�@                ' �Ǝ����_���A(���������͢����+����������͢�蓮�ŕ␳�Ȃ��Ń����Ұ�=���_���A�w�裎�) ���Ƃ���̏ꍇ
                    'V2.1.0.1�@                If ((stThta.iPP30 = 0) And (gSysPrm.stDEV.giTheta <> 0)) Or _
                    'V2.1.0.1�@                   ((stThta.iPP30 = 2) And (gSysPrm.stDEV.giTheta <> 0)) Or _
                    'V2.1.0.1�@                   ((stThta.iPP30 = 1) And (stThta.iPP31 = 0) And (gSysPrm.stSPF.giThetaParam = 1) And (gSysPrm.stDEV.giTheta <> 0)) Then
                    'V2.1.0.1�@                    Call ROUND4(0.0#)                                   ' �Ƃ����_�ɖ߂�
                    'V2.1.0.1�@                End If

                    ' ���ߐݒ�
                    Call System1.sLampOnOff(LAMP_RESET, True)               ' RESET�����v�_��
                    Call System1.sLampOnOff(LAMP_START, True)               ' START�����v�_��

                    'V2.0.0.2�@                If (iRtn <> cFRS_ERR_PTN And DGL <> TRIM_MODE_POWER) Then                          ' ###1040�D �Ƃ̃p�^�[���F���m�f���ɏ�������������Ă��܂��C���B'V2.0.0.0�ATRIM_MODE_POWER�ǉ�
                    If UserSub.IsTRIM_MODE_ITTRFT() And (iRtn <> cFRS_ERR_PTN And DGL <> TRIM_MODE_POWER) Then                   ' ###1040�D �Ƃ̃p�^�[���F���m�f���ɏ�������������Ă��܂��C���B'V2.0.0.0�ATRIM_MODE_POWER�ǉ�'V2.0.0.2�@ �J�b�g���s�����O����ׂ�IsTRIM_MODE_ITTRFT()�ǉ�
                        Call UserSub.SubstrateEnd()                         ' ��P�ʂ̌��ʏo��
                    End If

                    gbClampOpen = True        'V1.2.0.0�C �N�����v�J��ԉ���
                    gbVaccumeOff = True       'V1.2.0.0�C �z���I�t��ԉ���

                    '-----------------------------------------------------------------------
                    '   �g���~���O���ʂ����[�_�֏o�͂���
                    '-----------------------------------------------------------------------
                    'V2.0.0.1�B��
                    If iRtn = cFRS_TRIM_NG Then
                        If UserSub.IsTRIM_MODE_ITTRFT() And PlateNGJudgeByCounter() Then

                            'V2.2.0.036��
                            If (giLoaderType = 1) AndAlso (giHostMode = cHOSTcMODEcAUTO) Then

                                ' �V�O�i���^���[����(On=�����^�]��(�Γ_��),Off=�S�ޯ�)
                                Call Me.System1.SetSignalTowerCtrl(Me.System1.SIGNAL_ALARM)

                                '  "NG�����ݒ�l�𒴂��܂����B" "START�L�[�F�������s�CRESET�L�[�F�����I��"
                                Dim ret As Integer = ObjLoader.Sub_CallFrmMsgDisp(Me.System1, cGMODE_MSG_DSP, cFRS_ERR_START + cFRS_ERR_RST, True,
                                    My.Resources.MSG_SPRASH56, My.Resources.MSG_SPRASH35, "", System.Drawing.Color.Red, System.Drawing.Color.Black, System.Drawing.Color.Black)

                                If (ret = cFRS_ERR_START) Then
                                    '���s
                                    Call Me.System1.SetSignalTowerCtrl(Me.System1.SIGNAL_OPERATION)

                                Else
                                    Call Me.System1.SetSignalTowerCtrl(Me.System1.SIGNAL_IDLE)

                                    ' ���b�g�������f 
                                    ' ���f���Ŏ��̊�͏������Ȃ��B 
                                    fStartTrim = False                       ' �X�^�[�gTRIM�t���O��OFF
                                    ' �@���菜�����b�Z�[�W��\������
                                    r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE)

                                    frmAutoObj.SetAutoOpeCancel(False)               ' V2.2.1.1�A
                                    Call frmAutoObj.AutoOperationEnd()              ' V2.2.1.1�A
                                    frmAutoObj.gbFgAutoOperation = False

                                    ' ���_���A�m�F 
                                    r = sResetStart()
                                    If (r <> cFRS_NORMAL) Then                          ' �G���[ ?
                                        '���_���A�G���[�̏ꍇ�̓v���O�����I�� 
                                        r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE2)
                                        Call AppEndDataSave()                           ' ��ċ����I�������ް��ۑ��m�F
                                        Call AplicationForcedEnding()                   ' ��ċ����I������
                                        End                                             ' �A�v�������I��
                                        Return
                                    End If
                                    ' �d�����b�N(�ω����E�����b�N)����������
                                    r = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_OFF)

                                End If
                            End If
                            'V2.2.0.036��

                            iRtn = cFRS_TRIM_NG
                        Else
                            iRtn = cFRS_NORMAL
                        End If
                    End If
                    'V2.0.0.1�B��

                    If (iRtn <> cFRS_NORMAL And iRtn <> cFRS_ERR_RST) Then      'V2.0.0.1�B cFRS_ERR_RST�ǉ�
                        ' �G���[��
                        If giLoaderType = 1 Then
                            Call Sub_ATLDSET(0, COM_STS_TRM_STATE)
                        Else
                            Call Sub_ATLDSET(COM_STS_TRM_NG, COM_STS_TRM_STATE)
                        End If

                        DebugLogOut("�g���~���O�s�ǐM��(BIT1)�o�� Result=[" & iRtn.ToString & "]")
                    Else
                        ' ���펞
                        If giLoaderType = 1 Then
                            Call Sub_ATLDSET(0, COM_STS_TRM_STATE)
                        Else
                            Call Sub_ATLDSET(0, COM_STS_TRM_STATE Or COM_STS_TRM_NG)
                        End If

                    End If
                    'If (r = cFRS_ERR_PTN) Then                             ' �p�^�[���F���G���[ ?
                    '    Call Sub_ATLDSET(COM_STS_PTN_NG Or COM_STS_TRM_NG, COM_STS_TRM_STATE)
                    'ElseIf (r <> cFRS_NORMAL) Then                         ' �G���[ ?
                    '    Call Sub_ATLDSET(COM_STS_TRM_NG, COM_STS_TRM_STATE Or COM_STS_PTN_NG)
                    'Else                                            ' ����
                    '    Call Sub_ATLDSET(0, COM_STS_TRM_STATE Or COM_STS_TRM_NG Or COM_STS_PTN_NG)
                    'End If

                    'V2.2.0.0�D��
                    ' TLF�����[�_�̏ꍇ�A�蓮�Ȃ烍�b�N����
                    If (giLoaderType <> 0) AndAlso (frmAutoObj.gbFgAutoOperation = False) Then
                        r = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_OFF)
                        If (r <> cFRS_NORMAL) Then                                  ' �G���[ ?(���b�Z�[�W�͕\����)

                        End If
                    End If

                    ' �g���~���O���s�㏈��
                    Call System1.AutoLoaderFlgReset()                       ' �I�[�g���[�_�[�t���O���Z�b�g
                    giAppMode = APP_MODE_IDLE                               ' ����Ӱ�� =�g���}���u�A�C�h����
                    fStartTrim = False                                      ' �X�^�[�gTRIM�t���O OFF

                    '---------------------------------------------------------------------------
                    '   �X�^�[�gTRIM�t���O��OFF�Ȃ�A�ȉ��̏������s��
                    '---------------------------------------------------------------------------
                Else
                    ' �}�j���A�����[�_��RESET SW�������́u���_���A�����v���s��(�A�C�h�����[�h���Ƀ`�F�b�N����)
                    If (giAppMode = APP_MODE_IDLE) Then
                    STARTRESET_SWCHECK(1, swStatus)
                    ' �}�j���A�����[�_��RESET �r�v�����H
                    If giHostMode = cHOSTcMODEcMANUAL And swStatus = cFRS_ERR_RST Then
                        Call System1.sLampOnOff(LAMP_RESET, True)       ' RESET�����vON
                        ' ���_���A
                        r = sResetStart()                               ' RESET/START�L�[�҂�
                        ' ����~���Ȃ�A�v�������I��
                        If (r < cFRS_NORMAL) Then
                            r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE2)

                            GoTo TimerErr
                        End If
                        If (r = cFRS_ERR_RST) Then                      ' RESET�L�[���� ? 
                            ' ����ۯ����Ȃ�ײ�޶�ް�J�҂�
                            ' V2.2.0.0�D If (Me.System1.InterLockSwRead() And BIT_INTERLOCK_DISABLE) = 0 Then
                            If (Me.System1.InterLockSwRead() And BIT_INTERLOCK_DISABLE) = 0 AndAlso giLoaderType = 0 Then   'V2.2.0.0�D
                                r = Me.System1.Form_Reset(cGMODE_OPT_END, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
                                If (r < cFRS_NORMAL) Then GoTo TimerErr ' �G���[�Ȃ�A�v�������I��
                            End If
                        End If

                        ' �����v����
                        Call Me.System1.sLampOnOff(LAMP_START, True)    ' START����ON
                        Call Me.System1.sLampOnOff(LAMP_RESET, True)    ' RESET����ON
                        GoTo TimerExit
                    End If
                End If
            End If

            '---------------------------------------------------------------------------
            '   �I������
            '---------------------------------------------------------------------------
TimerExit:
            Timer1.Enabled = True                                       ' �Ď��^�C�}�[�J�n
            Exit Sub

            '---------------------------------------------------------------------------
            '   �A�v�������I��
            '---------------------------------------------------------------------------
TimerErr:
            Call AppEndDataSave()                                       ' ��ċ����I�������ް��ۑ��m�F
            Call AplicationForcedEnding()                               ' ��ċ����I������
            End                                                         ' �A�v�������I��

        Catch ex As Exception
            Call Z_PRINT("Timer1.Tick() TRAP ERROR = " + ex.Message & vbCrLf)
        End Try
    End Sub
#End Region

    '========================================================================================
    '   ���̑��̃C�x���g����
    '========================================================================================
#Region "�f�W�^��SW�̑I�����ڂ��ς�����ꍇ�̏���"
    '''=========================================================================
    ''' <summary>�f�W�^��SW�̑I�����ڂ��ς�����ꍇ�̏���</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub CbDigSwL_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CbDigSwL.SelectedIndexChanged
        DGL = CbDigSwL.SelectedIndex                            ' �f�W�^���r�v(Low)
        DGSW = (DGH * 10) + DGL                                 ' �f�W�^���r�v
    End Sub

    Private Sub CbDigSwH_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CbDigSwH.SelectedIndexChanged
        DGH = CbDigSwH.SelectedIndex                            ' �f�W�^���r�v(Hight) 
        DGSW = (DGH * 10) + DGL                                ' �f�W�^���r�v
    End Sub
#End Region

    '========================================================================================
    '   ���ʊ֐�
    '========================================================================================
#Region "���_���A�T�u"
    '''=========================================================================
    '''<summary>���_���A�T�u</summary>
    ''' <returns>0:����, 0�ȊO:�G���[</returns>
    '''<remarks></remarks>
    '''=========================================================================
    Public Function sResetStart() As Short

        Dim r As Short = cFRS_NORMAL
        Dim rtn As Short = cFRS_NORMAL
        Dim strMSG As String

        Try
            ' ���_���A
#If cOFFLINEcDEBUG = 1 Then
            Return (cFRS_NORMAL)                                                    ' Return�l = ����
#End If
            ' ���_���A

            Call SETAXISSPDY(SETAXISSPDY_DEFALT)                                    ' 'V2.0.0.0�N�x���X�e�[�W���x�����ɖ߂��B25000����15000�֕ύX

            Call Sub_ATLDSET(COM_STS_TRM_STATE Or COM_STS_LOT_END, 0)               ' ���[�_�[�o��(ON=�g���}���쒆 �@'V1.2.0.0�C���b�g�I��(0:������, 1:�I�����),OFF=�Ȃ�)


            'V2.2.0.0�D��
            ' TLF�����[�_�̏ꍇ�A���[�_���_���A���s��
            If (giLoaderType <> 0) Then
                Call ATTRESET()                                             'V2.1.0.0�E
                ' �A���[�����Z�b�g����
                ObjSys.W_RESET()
                Call Sub_ATLDSET(0, clsLoaderIf.LOUT_AUTO Or clsLoaderIf.LOUT_STOP Or clsLoaderIf.LOUT_SUPLY Or clsLoaderIf.LOUT_STS_RUN Or clsLoaderIf.LOUT_REQ_COLECT Or clsLoaderIf.LOUT_DISCHRAGE)                             ' ���[�_�o��(ON=��v���܂��͋����ʒu������+��ϒ�~��+��, OFF=�����ʒu�������܂��͊�v��)

                Call Sub_ATLDSET(clsLoaderIf.LOUT_STS_RUN, clsLoaderIf.LOUT_REDY)               ' ���[�_�[�o��(ON=�g���}���쒆 �@'V1.2.0.0�C���b�g�I��(0:������, 1:�I�����),OFF=�Ȃ�)

            End If
            'V2.2.0.0�D��


            r = System1.Form_Reset(cGMODE_ORG, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
            If (r = cFRS_ERR_RST) Then                                              ' RESET ?
                rtn = r                                                             ' Return�l�ݒ�
                GoTo STP_END
            End If
            If (r <> cFRS_NORMAL) Then                                              ' �G���[ ?

                '���[�_���_���A�^�C���A�E�g�̏ꍇ�ɂ́A
                If (r = cFRS_ERR_LDRTO) Then                          ' ���[�_�ʐM�^�C���A�E�g ?
                    ' rtnCode = Sub_CallFrmRset(ObjSys, cGMODE_LDR_TMOUT)     ' �G���[���b�Z�[�W�\��
                    AutoOperationDebugLogOut("sResetStart() r = cFRS_ERR_LDRTO")       ''V2.2.1.3�A

                    r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_TMOUT)
                Else
                    ' ���[�_�A���[�����b�Z�[�W�쐬 & ���[�_�A���[����ʕ\��
                    r = ObjSys.Sub_CallFormLoaderAlarm(cGMODE_LDR_ALARM, ObjPlcIf)
                End If
                Call Sub_ATLDSET(&H0, clsLoaderIf.LOUT_AUTO)        ' ���[�_�蓮���[�h�ؑւ�(���[�_�o��(ON=�Ȃ�, OFF=����))

                Return (r)
            End If

            ' ����SW�����҂�(��߼��)���ͽײ�޶�ް��������݂��Ȃ��̂�
            ' ү���ޕ\����ײ�޶�ް�J�҂�
            If ((System1.InterLockSwRead() And BIT_INTERLOCK_DISABLE) = 0) And
                (gSysPrm.stSPF.giWithStartSw = 1) Then                               ' ����ۯ����Ž���SW�����҂�(��߼��) ?
                r = System1.Form_Reset(cGMODE_OPT_END, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
                If (r <> cFRS_NORMAL And r <> cFRS_ERR_RST) Then                    ' �G���[ ?
                    Return (r)                                                      ' Return�l�ݒ�
                End If
                rtn = r
            End If

            'V2.2.0.0�D��
            ' TLF�����[�_�̏ꍇ�A���[�_���_���A���s��
            If (giLoaderType <> 0) Then
                r = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_OFF)
                If (r <> cFRS_NORMAL) Then                                  ' �G���[ ?(���b�Z�[�W�͕\����)
                    Return (r)
                End If
            End If

            ' �N�����v/�z��OFF
            r = System1.ClampVacume_Ctrl(gSysPrm, 0, giAppMode, 1, False)
            If (r <> cFRS_NORMAL) Then                                              ' �G���[ ?
                Return (r)                                                          ' Return�l�ݒ�
            End If



STP_END:
            Call Sub_ATLDSET(0, COM_STS_TRM_STATE)                                  ' ���[�_�[�o��(ON=�Ȃ�,OFF=�g���}���쒆)
            Return (rtn)

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "sResetStart() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            sResetStart = cERR_TRAP                                                 ' Return�l = ��O�G���[
        End Try
    End Function
#End Region

#Region "�C���^�[���b�N��Ԃ̕\��/��\��"
    '''=========================================================================
    '''<summary>�C���^�[���b�N��Ԃ̕\��/��\��</summary>
    ''' <returns>�C���^�[���b�N���
    '''          INTERLOCK_STS_DISABLE_FULL = �C���^�[���b�N�S����
    '''          INTERLOCK_STS_DISABLE_PART = �C���^�[���b�N�ꕔ�����i�X�e�[�W����\�j
    '''          INTERLOCK_STS_DISABLE_NO   = �C���^�[���b�N��ԁi�����Ȃ��j
    ''' </returns>
    '''<remarks></remarks>
    '''=========================================================================
    Public Function DispInterLockSts() As Integer

        Dim r As Integer
        Dim InterlockSts As Integer
        Dim SwitchSts As Long
        Dim strMSG As String

        Try
            ' �C���^�[���b�N��Ԃɂ��X�e�[�^�X�\����ύX
            r = INTERLOCK_CHECK(InterlockSts, SwitchSts)
#If cOFFLINEcDEBUG Then
            InterlockSts = INTERLOCK_STS_DISABLE_FULL
#End If
            If (InterlockSts = INTERLOCK_STS_DISABLE_FULL) Then         ' �C���^�[���b�N�S���� ?
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    strMSG = "�C���^�[���b�N�S������"
                Else
                    strMSG = "Under Interlock Release"
                End If
                Me.lblInterLockMSG.Text = strMSG
                Me.lblInterLockMSG.Visible = True
                'V2.2.0.0�D ��
                If giLoaderType = 1 Then
                    Call Sub_ATLDSET(clsLoaderIf.LOUT_INTLOK_DISABLE, 0)
                End If
                'V2.2.0.0�D ��

            ElseIf (InterlockSts = INTERLOCK_STS_DISABLE_PART) Then     ' �C���^�[���b�N�ꕔ�����i�X�e�[�W����\�j ?
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    strMSG = "�C���^�[���b�N�ꕔ������"
                Else
                    strMSG = "Under Interlock Part Release"
                End If
                Me.lblInterLockMSG.Text = strMSG                        '�u�C���^�[���b�N�ꕔ�������v�\��
                Me.lblInterLockMSG.Visible = True
                'V2.2.0.0�D ��
                If giLoaderType = 1 Then
                    Call Sub_ATLDSET(clsLoaderIf.LOUT_INTLOK_DISABLE, 0)
                End If
                'V2.2.0.0�D ��

            Else                                                        ' �C���^�[���b�N��
                Me.lblInterLockMSG.Visible = False
                'V2.2.0.0�D ��
                If giLoaderType = 1 Then
                    Call Sub_ATLDSET(0, clsLoaderIf.LOUT_INTLOK_DISABLE)
                End If
                'V2.2.0.0�D ��
            End If

            Return (InterlockSts)

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "DispInterLockSts() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)
        End Try
    End Function
#End Region

#Region "�őO�ʕ\�����b�Z�[�W�{�b�N�X"
    ''' <summary>
    ''' �őO�ʕ\�����b�Z�[�W�{�b�N�X
    ''' </summary>
    ''' <param name="DispStr"></param>      ' �\�����b�Z�[�W
    ''' <param name="title"></param>        ' �\���^�C�g��(�ȗ��F�f�t�H���g��)
    ''' <param name="Button"></param>       ' �{�^�����(�ȗ��F�f�t�H���gOK�{�^���̂�)
    ''' <returns></returns>                 ' �������{�^���̎��
    ''' <remarks></remarks>
    Public Function MsgBoxForeground(ByVal DispStr As String, Optional ByVal title As String = "", Optional ByVal Button As MessageBoxButtons = vbOKOnly) As DialogResult
        Dim ret As DialogResult
        Using dummyForm As New Form()
            dummyForm.TopMost = False
            dummyForm.Width = 0
            dummyForm.Height = 0
            dummyForm.ControlBox = False
            dummyForm.Show()
            dummyForm.Visible = False
            dummyForm.TopMost = True
            ret = MessageBox.Show(dummyForm, DispStr, title, Button)
        End Using
        Return ret
    End Function
#End Region

#Region "��ċ����I�������ް��ۑ��m�F"
    '''=========================================================================
    '''<summary>��ċ����I�������ް��ۑ��m�F</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub AppEndDataSave()

        Dim ret As Short
        Dim strMSG As String

        Try
            ' �g���}���쒆�M��ON���M(���۰�ް)
            Call Sub_ATLDSET(COM_STS_TRM_STATE, 0)                      ' ���[�_�[�o��(ON=�g���}���쒆,OFF=�Ȃ�)

            'V2.2.0.0�@             Call FinalEnd_GazouProc(ObjGazou)                           'DispGazou�����I��

            V_Off()                                                     ' DC�d�����u �d��OFF

            ' �ҏW���̃f�[�^���� ?
            If (FlgUpd = True) Then
                If gSysPrm.stTMN.giMsgTyp = 0 Then
                    ret = MsgBoxForeground("�A�v���P�[�V�������I�����܂��B" & vbCrLf & "�g���~���O�f�[�^��ۑ����܂����H")
                Else
                    ret = MsgBoxForeground("Quits the program." & vbCrLf & "Do you store trimming data?")
                End If
                If ret = MsgBoxResult.Ok Then
                    ' �f�[�^�ۑ�
                    Call Me.cmdSave_Click(Me.cmdSave, New System.EventArgs())
                    If gSysPrm.stTMN.giMsgTyp = 0 Then
                        ret = MsgBoxForeground("�f�[�^�̕ۑ����������܂����B" & vbCrLf & "�A�v���P�[�V�������I�����܂��B")
                    Else
                        ret = MsgBoxForeground("A save of data was completed." & vbCrLf & "Quits the program.")
                    End If
                Else
                    ' �f�[�^�ۑ��Ȃ�
                    If gSysPrm.stTMN.giMsgTyp = 0 Then
                        ret = MsgBoxForeground("�A�v���P�[�V�������I�����܂��B")
                    Else
                        ret = MsgBoxForeground("Quits the program.")
                    End If
                End If
            Else
                If gSysPrm.stTMN.giMsgTyp = 0 Then
                    ret = MsgBoxForeground("�A�v���P�[�V�������I�����܂��B")
                Else
                    ret = MsgBoxForeground("Quits the program.")
                End If
            End If
            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "AppEndDataSave() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "��ċ����I������"
    '''=========================================================================
    '''<summary>��ċ����I������</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub AplicationForcedEnding()

        Dim lRet As Integer
        Dim hProcInf As New System.Diagnostics.ProcessStartInfo()
        'Dim ret As Short

        Try
            'V2.2.0.0�@            Call FinalEnd_GazouProc(ObjGazou)

            If frmAutoObj.gbFgAutoOperation Then
                lRet = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE2)
            End If

            ' �g���}���f�B�M��OFF���M(���۰�ް)
            Call Sub_ATLDSET(COM_STS_TRM_STATE, COM_STS_TRM_READY)          ' ���[�_�[�o��(ON=�g���}���쒆, ,OFF=�g���}���f�B)

            ' �V�O�i���^���[������(On=0, Off=�S�ޯ�)
            Call Me.System1.SetSignalTower(0, &HFFFFS)

            'V2.2.0.0�D��
            ' ���[�_�ʐM�N���[�Y 
            If giLoaderType = 1 Then
                ObjSys.ClosePLCThread()
                Call Sub_ATLDSET(&H0, clsLoaderIf.LOUT_AUTO)        ' ���[�_�蓮���[�h�ؑւ�(���[�_�o��(ON=�Ȃ�, OFF=����))

                ' �d�����b�N(�ω����E�����b�N)����������
                ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_OFF)

            End If
            'V2.2.0.0�D��

            ' �X���C�h�J�o�[�I�[�v��/�N���[�Y�o���uOFF
            If (gSysPrm.stTMN.gsKeimei = MACHINE_TYPE_SL432) Then            ' SL432R�n ? 
                Call ZSLCOVERCLOSE(0)                                       ' �X���C�h�J�o�[�N���[�Y�o���uOFF
                Call ZSLCOVEROPEN(0)                                        ' �X���C�h�J�o�[�I�[�v���o���uOFF
            End If

            ' �T�[�{�A���[���N���A
            Call CLEAR_SERVO_ALARM(1, 1)

            ' �r�f�I���C�u�����I������
            If (pbVideoInit = True) Then
                lRet = VideoLibrary1.Close_Library
                If (lRet <> 0) Then
                    Select Case lRet
                        Case cFRS_VIDEO_INI
                            'Call System1.TrmMsgBox(ggSysPrm, "Video library: Not initialized.", MsgBoxStyle.OkOnly, My.Application.Info.Title)
                            Call MsgBox("Video library: Not initialized.", MsgBoxStyle.OkOnly, My.Application.Info.Title) ' 2011.09.01
                        Case Else
                            ' "�\�����ʃG���["
                            'Call System1.TrmMsgBox(ggSysPrm, "Video library: Unexpected error.", MsgBoxStyle.OkOnly, My.Application.Info.Title)
                            Call MsgBox("Video library: Unexpected error.", MsgBoxStyle.OkOnly, My.Application.Info.Title) ' 2011.09.01
                    End Select
                End If
            End If

            ' ���샍�O�o��("���[�U�v���O�����I��")
            Call Me.System1.OperationLogging(gSysPrm, MSG_OPLOG_END, "")
            gflgCmpEndProcess = True

            ' �N�����v�y�уo�L���[��OFF 
            Call Me.System1.ClampVacume_Ctrl(gSysPrm, 0, giAppMode, giTrimErr, False)

            ' �f�o�h�a�I������
            ObjGpib.Gpib_Term(gDevId)

            ' �����vOFF
            Call LAMP_CTRL(LAMP_START, False)                               ' START�����vOFF 
            Call LAMP_CTRL(LAMP_RESET, False)                               ' RESET�����vOFF 
            Call LAMP_CTRL(LAMP_Z, False)                                   ' Z�����vOFF 

            ' �g���}���쒆�M��ON���M(���۰�ް)�O�̈�
            Call Sub_ATLDSET(COM_STS_TRM_STATE, 0)                          ' ���[�_�[�o��(ON=�g���}���쒆,OFF=�Ȃ�)
            '-----------------------------------------------------------------------
            '   �N���X���C���I�u�W�F�N�g�̉�� ###1040�E 
            '-----------------------------------------------------------------------
            'V2.2.0.0�@ ObjTch.SetCrossLineObject(0)                                   ' ###1040�E 
            'V2.2.0.0�@ ObjMTC.SetCrossLineObject(0)                                   ' ###1040�E 

            '-----------------------------------------------------------------------
            '   Mutex�̉��
            '-----------------------------------------------------------------------
            gmhUserPro.ReleaseMutex()

            '-----------------------------------------------------------------------
            '   �C�x���g�̉��
            '-----------------------------------------------------------------------
            RemoveHandler SystemEvents.SessionEnding, AddressOf SystemEvents_SessionEnding

            '-----------------------------------------------------------------------
            '�I����Videolib�֌W�ŃG���[���������邽�ߋ����I�ɊO������A�v�����I��������B
            '-----------------------------------------------------------------------
            hProcInf.FileName = APP_FORCEEND
            hProcInf.Arguments = System.Diagnostics.Process.GetCurrentProcess.ProcessName
            Call System.Diagnostics.Process.Start(hProcInf)
            System.Threading.Thread.Sleep(2000) ' �I����҂��Ȃ��Ǝ��̏����֐i�ݍċN����INtime�Ƃ̕s�������N���āu�G�A�[���ቺ���o�v����������B
        Catch ex As Exception
            ' ���샍�O�o��("���[�U�v���O�����I��")
            Call System1.OperationLogging(gSysPrm, MSG_OPLOG_END, "Exception")
            gflgCmpEndProcess = True
            'MsgBox("Execption error !" & vbCrLf & "error msg = " & ex.Message)
        End Try
    End Sub
#End Region

#Region "frmInfo��ʃ{�^��������/�񊈐���"
    '''=========================================================================
    '''<summary>frmInfo��ʃ{�^��������/�񊈐���</summary>
    '''<param name="Flg">(INP) 0=�{�^���񊈐���, 1=�{�^��������</param>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub SBtn_Enb_OnOff(ByRef Flg As Short)

        Dim strMSG As String

        Try
            ' �{�^��������/�񊈐���
            If (Flg = 1) Then                                           ' �{�^�������� ?
                ' �{�^��������

            Else
                ' �{�^���񊈐���
            End If

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "SBtn_Enb_OnOff() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "frmInfo��ʍ��ځ@������/�񊈐���"

    Public Sub Set_UserForm(ByRef Flg As Short)
        ' �{�^��������/�񊈐���
        If (Flg = 1) Then                                           ' �{�^�������� ?
            '' '' ''LblLOTHed.Visible = True
            '' '' ''LblLOT.Visible = True
            '' '' ''_Lbl_1.Visible = True
            '' '' ''_Lbl_2.Visible = True
            '' '' ''_Lbl_3.Visible = True
            '' '' ''_Lbl_4.Visible = True
            '' '' ''_Lbl_5.Visible = True
            '' '' ''_Lbl_6.Visible = True
            '' '' ''_Lbl_7.Visible = True

            '' '' ''LblN_0.Visible = True
            '' '' ''LblN_1.Visible = True
            '' '' ''LblN_2.Visible = True
            '' '' ''LblN_3.Visible = True
            '' '' ''LblN_4.Visible = True

        Else
            '' '' ''LblLOTHed.Visible = False
            '' '' ''LblLOT.Visible = False

            '' '' ''_Lbl_1.Visible = False
            '' '' ''_Lbl_2.Visible = False
            '' '' ''_Lbl_3.Visible = False
            '' '' ''_Lbl_4.Visible = False
            '' '' ''_Lbl_5.Visible = False
            '' '' ''_Lbl_6.Visible = False
            '' '' ''_Lbl_7.Visible = False

            '' '' ''LblN_0.Visible = False
            '' '' ''LblN_1.Visible = False
            '' '' ''LblN_2.Visible = False
            '' '' ''LblN_3.Visible = False
            '' '' ''LblN_4.Visible = False


        End If

    End Sub

#End Region

#Region "�{�^��������/�񊈐���"
    '''=========================================================================
    '''<summary>�{�^��������/�񊈐���</summary>
    '''<param name="Flg">(INP) 0=�{�^���񊈐���, 1=�{�^��������, 2=�{�^�����̕\��/��\��</param>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub Btn_Enb_OnOff(ByRef Flg As Short)

        Dim strMSG As String

        Try
            'LblDIGSW_HI.Visible = True
            LblDIGSW_HI.Visible = False
            LblDIGSW_LO.Visible = True                      ' "DSW="�@�\��
            'CbDigSwH.Visible = True
            CbDigSwH.Visible = False
            CbDigSwL.Visible = True
            '---------------------------------------------------------------------------
            '   �{�^��������������
            '---------------------------------------------------------------------------
            If (Flg = 1) Then

                'V2.1.0.0�A��
                If UserSub.IsLaserCaribrarionUse() Then
                    ButtonLaserCalibration.Enabled = True
                End If
                'V2.1.0.0�A��

                cmdHelp.Enabled = True                      ' HELP(About)
                cmdExit.Enabled = True                      ' END(F12)
                'cmdStart.Enabled = True                     ' START(Debug)

                If (stFNC(F_LOAD).iDEF = 0) Then            ' LOAD(F1)
                    cmdLoad.Enabled = False
                Else
                    cmdLoad.Enabled = True
                End If
                If (stFNC(F_SAVE).iDEF = 0) Then            ' SAVE(F2)
                    cmdSave.Enabled = False
                Else
                    cmdSave.Enabled = True
                End If
                If (stFNC(F_EDIT).iDEF = 0) Then            ' EDIT(F3)
                    cmdEdit.Enabled = False
                Else
                    cmdEdit.Enabled = True
                End If
                'If (stFNC(F_MSTCHK).iDEF = 0) Then         ' Ͻ�����(F4)
                '    cmdMstChk.Enabled = False
                'Else
                '    cmdMstChk.Enabled = True
                'End If
                If (stFNC(F_LASER).iDEF = 0) Then           ' LASER(F5)
                    cmdLaserTeach.Enabled = False
                    cmdLaserCalibration.Enabled = False     'V2.1.0.0�A
                Else
                    cmdLaserTeach.Enabled = True
                    'V2.1.0.0�A��
                    If UserSub.IsLaserCaribrarionUse Then
                        cmdLaserCalibration.Enabled = True
                    End If
                    'V2.1.0.0�A��
                End If
                If (stFNC(F_LOTCHG).iDEF = 0) Then          ' ۯĐؑ�(S-F6)
                    cmdLotChg.Enabled = False
                Else
                    cmdLotChg.Enabled = True
                End If
                If (stFNC(F_PROBE).iDEF = 0) Then           ' PROBE(F7)
                    cmdProbeTeaching.Enabled = False
                Else
                    cmdProbeTeaching.Enabled = True
                End If
                If (stFNC(F_TEACH).iDEF = 0) Then           ' TEACH(F8)
                    cmdTeaching.Enabled = False
                Else
                    cmdTeaching.Enabled = True
                End If
                If (stFNC(F_CUTPOS).iDEF = 0) Then          ' CutPosTeach(S-F8)
                    cmdCutPosTeach.Enabled = False
                Else
                    cmdCutPosTeach.Enabled = True
                End If
                If (stFNC(F_RECOG).iDEF = 0) Then          ' RECOG(F9)
                    BtnRECOG.Enabled = False
                Else
                    BtnRECOG.Enabled = True
                End If
                'V2.0.0.0�@��
                If (stFNC(F_TX).iDEF > 0) Then          ' TX(F9)
                    CmdTx.Enabled = True
                Else
                    CmdTx.Enabled = False
                End If
                If (stFNC(F_TY).iDEF > 0) Then          ' TY(F10)
                    CmdTy.Enabled = True
                Else
                    CmdTy.Enabled = False
                End If
                'V2.0.0.0�@��

                BtnStartPosSet.Enabled = True           'V2.0.0.0�A

                cmdClamp.Enabled = True                    'V2.2.1.1�H

                ' ���[�U�[���ꏈ�� START
                cmdLotInfo.Enabled = True
                cmdPrint.Enabled = True
                ' ���[�U�[���ꏈ�� END
                '---------------------------------------------------------------------------
                '   �{�^����񊈐�������
                '---------------------------------------------------------------------------
            ElseIf (Flg = 0) Then

                ButtonLaserCalibration.Enabled = False      'V2.1.0.0�A

                cmdHelp.Enabled = False                     ' HELP(About)
                cmdStart.Enabled = False                    ' START(Debug) 
                cmdExit.Enabled = False                     ' END(F10)
                cmdLoad.Enabled = False                     ' LOAD(F1)
                cmdSave.Enabled = False                     ' SAVE(F2)
                cmdEdit.Enabled = False                     ' EDIT(F3)
                'cmdMstChk.Enabled = False                  ' Ͻ�����(F4)
                cmdLaserTeach.Enabled = False               ' LASER(F5)
                cmdLaserCalibration.Enabled = False         'V2.1.0.0�A ���[�U�L�����u���[�V����
                cmdLotChg.Enabled = False                   ' ۯĐؑ�(S-F6)
                cmdProbeTeaching.Enabled = False            ' PROBE(F7)
                cmdTeaching.Enabled = False                 ' TEACH(F8)
                cmdCutPosTeach.Enabled = False              ' CutPosTeach(S-F8)
                BtnRECOG.Enabled = False                    ' RECOG(F9)
                'V2.0.0.0�@��
                CmdTx.Enabled = False                       ' TX(F9)
                CmdTy.Enabled = False                       ' TY(F10)
                'V2.0.0.0�@��

                BtnStartPosSet.Enabled = False              'V2.0.0.0�A

                cmdClamp.Enabled = False                    'V2.2.1.1�H

                ' ���[�U�[���ꏈ�� START
                cmdLotInfo.Enabled = False
                cmdPrint.Enabled = False
                ' ���[�U�[���ꏈ�� END
                '---------------------------------------------------------------------------
                '   �{�^�����̕\��/��\����ݒ肷��
                '---------------------------------------------------------------------------
            Else
                Grpcmds.Visible = True
                frmInfo.Visible = True                      ' �g���~���O���ʕ\���t���[��
                'txtLog.Visible = True
                If giTxtLogType <> 0 Then
                    txtlog.Visible = True                       ' 'V2.2.0.0�P
                Else
                    lstLog.Visible = True                       ' ###lstLog
                End If

#If cOFFLINEcDEBUG Then
                'cmdStart.Visible = True
#Else
                cmdStart.Visible = False
#End If
                cmdExit.Visible = True                      ' END(F12)
                If (stFNC(F_LOAD).iDEF >= 0) Then           ' LOAD(F1)
                    cmdLoad.Visible = True
                Else
                    cmdLoad.Visible = False
                End If
                If (stFNC(F_SAVE).iDEF >= 0) Then           ' SAVE(F2)
                    cmdSave.Visible = True
                Else
                    cmdSave.Visible = False
                End If
                If (stFNC(F_EDIT).iDEF >= 0) Then           ' EDIT(F3)
                    cmdEdit.Visible = True
                Else
                    cmdEdit.Visible = False
                End If
                'If (stFNC(F_MSTCHK).iDEF >= 0) Then        ' Ͻ�����(F4)
                '    cmdMstChk.Visible = True
                'Else
                '    cmdMstChk.Visible = False
                'End If
                If (stFNC(F_LASER).iDEF >= 0) Then          ' LASER(F5)
                    cmdLaserTeach.Visible = True
                    'V2.1.0.0�A��
                    If UserSub.IsLaserCaribrarionUse Then
                        cmdLaserCalibration.Visible = True
                    End If
                    'V2.1.0.0�A��
                Else
                    cmdLaserTeach.Visible = False
                    cmdLaserCalibration.Visible = False     'V2.1.0.0�A
                End If
                If (stFNC(F_LOTCHG).iDEF >= 0) Then         ' ۯĐؑ�(S-F6)
                    cmdLotChg.Visible = True
                Else
                    cmdLotChg.Visible = False
                End If
                If (stFNC(F_PROBE).iDEF >= 0) Then          ' PROBE(F7)
                    cmdProbeTeaching.Visible = True
                Else
                    cmdProbeTeaching.Visible = False
                End If
                If (stFNC(F_TEACH).iDEF >= 0) Then          ' TEACH(F8)
                    cmdTeaching.Visible = True
                Else
                    cmdTeaching.Visible = False
                End If
                If (stFNC(F_CUTPOS).iDEF >= 0) Then         ' CutPosTeach(S-F8)
                    cmdCutPosTeach.Visible = True
                Else
                    cmdCutPosTeach.Visible = False
                End If
                If (stFNC(F_RECOG).iDEF >= 0) Then          ' RECOG(F9)
                    BtnRECOG.Visible = True
                Else
                    BtnRECOG.Visible = False
                End If

                'V2.0.0.0�@��
                If (stFNC(F_TX).iDEF > 0) Then          ' TX(F9)
                    CmdTx.Visible = True
                Else
                    CmdTx.Visible = False
                End If
                If (stFNC(F_TY).iDEF > 0) Then          ' TY(F10)
                    CmdTy.Visible = True
                Else
                    CmdTy.Visible = False
                End If
                'V2.0.0.0�@��

                cmdClamp.Enabled = True                    'V2.2.1.1�H

            End If

STP_END:

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "Btn_Enb_OnOff() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "���[�U�p���[�����֘A���ڂ̕\��/��\���ݒ�"
    '''=========================================================================
    '''<summary>���[�U�p���[�����֘A���ڂ̕\��/��\���ݒ�</summary>
    ''' <param name="Md">(INP)0=�\�����Ȃ�, 1=�\������</param>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub SetLaserItemsVisible(ByVal Md As Integer)

        Dim strMSG As String

        Try
            ' ���������V�X�p�����\������("������ = 99.9%")
            Me.LblRotAtt.Visible = False                                ' ��������\��
            If (Md = 1) Then                                            ' �\������ ?
                If (gSysPrm.stRMC.giRmCtrl2 >= 2 And
                    gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then       ' ۰�ر��Ȱ�����L��FL�łȂ� ?
                    Me.LblRotAtt.Visible = True
                End If
            End If

            ' ����l���V�X�p�����\������
            Me.LblMes.Visible = False                                   ' ����l��\��
            If (Md = 1) Then                                            ' �\������ ?
                ' RMCTRL2 >=3 �� ����l�\�� ?
                If (gSysPrm.stRMC.giRmCtrl2 >= 3) And (gSysPrm.stRMC.giPMonHi = 1) Then
                    LblMes.Visible = False                              ' ����l��\��
                End If
            End If

            ' ��d���l��\������
            LblCur.Visible = False                                      ' ��d���l��\��
            If (Md = 1) Then                                            ' �\������ ?
                ' ���H�d�͐ݒ� = 4(��d��1A)�̎��ɕ\������
                If (gSysPrm.stSPF.giProcPower = 4) And (gSysPrm.stSPF.giProcPower2 <> 0) Then
                    LblCur.Visible = True
                End If
            End If

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "SetLaserItemsVisible() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "���샍�O�o�̓T�u"
    '''=========================================================================
    '''<summary>���샍�O�o�̓T�u</summary>
    '''<param name="gSts">����Ӱ��(giAppMode�Q��)</param>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub Sub_OprLog(ByRef gSts As Short)

        Dim strMSG As String

        Try
            ' ���O���b�Z�[�W�ݒ�
            Select Case (gSts)
                Case APP_MODE_LASER
                    strMSG = MSG_OPLOG_FUNC05       ' "���[�U����"
                Case APP_MODE_PROBE
                    strMSG = MSG_OPLOG_FUNC07       ' "�v���[�u�ʒu���킹"
                Case APP_MODE_PROBE2
                    strMSG = MSG_OPLOG_FUNC10       ' "�v���[�u�ʒu���킹�Q"
                Case APP_MODE_TEACH
                    strMSG = MSG_OPLOG_FUNC08       ' "�e�B�[�`���O"
                Case APP_MODE_CUTPOS
                    strMSG = MSG_OPLOG_FUNC08S      ' "�J�b�g�␳�ʒu�e�B�[�`���O"
                Case APP_MODE_RECOG
                    strMSG = MSG_OPLOG_FUNC09       ' "�p�^�[���o�^"
                Case Else
                    Exit Sub
            End Select

            ' ���샍�O�o��
            Call Me.System1.OperationLogging(gSysPrm, strMSG, "MANUAL")

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "Sub_OprLog() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�z�X�g�R�}���h�@�V�~�����[�V����(DEBUG�p)"
    '''=========================================================================
    '''<summary>�z�X�g�R�}���h�@�V�~�����[�V����(DEBUG�p)</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub DEBUG_HST_CMD_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles DEBUG_HST_CMD.Click

        Dim Index As Short = DEBUG_HST_CMD.GetIndex(eventSender)

        ' �z�X�g�R�}���h�@�V�~�����[�V����(DEBUG�p)
        Call DEBUG_ReadHostCommand(Index)               ' ���[�_���̓T�u(�f�o�b�O�p)

    End Sub
#End Region

    '========================================================================================
    '   �e�R�}���h(���[�U�R���g���[���Ƀt�H�[��������OCX�̏ꍇ)���t�H�[�J�X���������ꍇ�̏���
    '   �e���L�[��UP/Down�C�x���g�������Ă��Ȃ��Ȃ邽��OCX�Ƀt�H�[�J��ݒ肷��
    '========================================================================================
#Region "�r�f�I�摜���N���b�N���ăt�H�[�J�X���������ꍇ�̏���"
    '''=========================================================================
    ''' <summary>�r�f�I�摜���N���b�N���ăt�H�[�J�X���������ꍇ�̏���</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>�EOcxTeach��DispGazou.EXE���s���̂���Enter�C�x���g�͓����Ă��Ȃ�
    '''          �@DispGazou.EXE��OcxTeach�ŋN������
    '''          �EEnter�C�x���g��Form��ACTIVE�R���g���[���ɂȂ������ɔ���</remarks>
    '''=========================================================================
    Private Sub VideoLibrary1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try
            CmdSetFocus()

        Catch ex As Exception
        End Try
    End Sub
#End Region

#Region "���O�\����`�掞"
    ''' <summary>ListBox�����s�\���E�܂�Ԃ��\��</summary>
    ''' <remarks>###lstLog</remarks>
    Private Sub lstLog_DrawItem(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles lstLog.DrawItem
        Dim lstBox As ListBox = DirectCast(sender, ListBox)
        With lstBox
            If (0 = .Items.Count) OrElse (.Items Is Nothing) OrElse (e.Index < 0) Then Exit Sub
            Dim strItem As String = .GetItemText(.Items(e.Index))
            e.DrawBackground()
            e.Graphics.DrawString(strItem, e.Font, New SolidBrush(e.ForeColor), e.Bounds)
            e.DrawFocusRectangle()
        End With
    End Sub

    ''' <summary>ListBox�����s�\���E�܂�Ԃ��\��</summary>
    ''' <remarks>###lstLog</remarks>
    Private Sub lstLog_MeasureItem(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MeasureItemEventArgs) Handles lstLog.MeasureItem
        Dim lstBox As ListBox = DirectCast(sender, ListBox)
        With lstBox
            If (0 = .Items.Count) OrElse (.Items Is Nothing) OrElse (e.Index < 0) Then Exit Sub
            Dim strItem As String = .GetItemText(.Items(e.Index))
            Dim z As SizeF = e.Graphics.MeasureString(
                strItem, .Font, Convert.ToInt32(e.Graphics.VisibleClipBounds.Width))
            e.ItemWidth = Convert.ToInt32(z.Width)
            e.ItemHeight = Convert.ToInt32(z.Height)
        End With
    End Sub
#End Region

#Region "���O�\����N���b�N��"
    ''V2.2.0.0�P��    �L���ɂ���
    Private Sub txtLog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtlog.Click
        Try
            CmdSetFocus()

        Catch ex As Exception
        End Try
    End Sub
    ''V2.2.0.0�P��

    Private Sub lstLog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstLog.Click
        Try
            CmdSetFocus()

        Catch ex As Exception
        End Try
    End Sub

    Private Sub lstLog_Copy(ByVal sender As Object, ByVal e As EventArgs)
        Dim sb As New StringBuilder(256)
        For Each item As Object In Me.lstLog.SelectedItems
            sb.Append(item)
        Next
        If (0 < sb.Length) Then
            Clipboard.SetText(sb.ToString())
        End If
    End Sub

    ''V2.2.0.0�P��
    Private Sub txtLog_Copy(ByVal sender As Object, ByVal e As EventArgs)
        Dim sb As New StringBuilder(256)
        For Each item As Object In Me.txtlog.Lines
            sb.Append(item)
        Next
        If (0 < sb.Length) Then
            Clipboard.SetText(sb.ToString())
        End If
    End Sub
    ''V2.2.0.0�P��



#End Region

#Region "�e�R�}���h(OCX)�Ƀt�H�[�J�X��ݒ肷��"
    '''=========================================================================
    ''' <summary>�e�R�}���h(OCX)�Ƀt�H�[�J�X��ݒ肷��</summary>
    '''=========================================================================
    Private Sub CmdSetFocus()

        Dim strMSG As String

        Try
            Select Case (giAppMode)
                Case APP_MODE_PROBE
                    ' �v���[�u�R�}���h���s�� ?
                    Probe1.Focus()                                      ' OcxProbe�Ƀt�H�[�J�X���Z�b�g���� 

                Case APP_MODE_TEACH
                    ' �e�B�[�`�R�}���h���s�� ?
                    Teaching1.Focus()                                   ' OcxTeach�Ƀt�H�[�J�X���Z�b�g���� 
                    Teaching1.JogSetFocus()                             ' ��Probe�ƈႢ���̂����ꂪ�Ȃ��ƃe���L�[�������Ȃ�
            End Select

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "frmMain.CmdSetFocus() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

    Private Sub cmdLotInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLotInfo.Click

        Dim s As String
        Dim strMSG As String

        Try
            ' �g���}���쒆�M��ON���M(���۰�ް)
            Call Sub_ATLDSET(COM_STS_TRM_STATE, 0)                      ' ���[�_�[�o��(ON=�g���}���쒆,OFF=�Ȃ�)

            ' �g���}���u�A�C�h�����ȊO�Ȃ�NOP
            If giAppMode Then Exit Sub
            giAppMode = APP_MODE_LOTNO                                   ' ����Ӱ�� = ���b�g�ԍ��ݒ蒆

            ' �f�[�^���[�h�`�F�b�N
            If pbLoadFlg = False Then
                s = MSG_DataNotLoad                                     ' �ް���۰��
                Call Z_PRINT(s)
                Call Beep()
                GoTo STP_END
            End If

            If Not UserSub.IsSpecialTrimType Then
                Call Z_PRINT("�g���~���O�f�[�^�̐��i��ʂ��w�薳���ɐݒ肳��Ă��܂��B" & vbCrLf)
                Call Beep()
                GoTo STP_END
            End If
            ' �f�[�^�ҏW
            Call Me.System1.OperationLogging(gSysPrm, MSG_OPLOG_LOTSET, "")

            ChkLoaderInfoDisp(0)                              'V2.2.0.0�D

            Dim Rtn As Short
            Dim fLotInf As New FormEdit.frmLotInfoInput(True)
            fLotInf.ShowDialog(Me)
            Rtn = fLotInf.sGetReturn()
            fLotInf.Dispose()
            If Rtn = cFRS_ERR_START Then                                ' �n�j���^�[��
                Call UserSub.SetStartCheckStatus(True)                  ' �ݒ��ʂ̊m�F�L����
            End If

STP_END:
            Call ZCONRST()                                              ' �ݿ�ٷ� ׯ�����
            giAppMode = APP_MODE_IDLE                                   ' ����Ӱ�� = �g���}���u�A�C�h����
            Call Sub_ATLDSET(0, COM_STS_TRM_STATE)                      ' ���[�_�[�o��(ON=�Ȃ�,OFF=�g���}���쒆)

            ChkLoaderInfoDisp(1)                              'V2.2.0.0�D

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "cmdLotInfo_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        Finally
            giAppMode = APP_MODE_IDLE                                   ' ����Ӱ�� = �g���}���u�A�C�h����
        End Try

    End Sub

#Region "������ݸد������"
    '''=========================================================================
    ''' <summary>������ݸد������</summary>
    '''=========================================================================
    Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrint.Click
        Call UserSub.LotEnd()                                   ' ���b�g�I�����̃f�[�^�o��
        Call Printer.Print(True)                                ' �m�Fү���ނ̖߂�l�ɂ�����������Ȃ�
        Call UserSub.SetStartCheckStatus(True)                  ' ������s�����烍�b�g�I���Ƃ݂Ȃ��B
    End Sub
#End Region

    '==========================================================================
    '   �X�e�b�v�ړ�����
    '==========================================================================
#Region "�X�e�b�v�ړ��{�^���\��"
    Public Sub StepMoveButtonOn()
        BtnForward.Enabled = True
        BtnForward.Visible = True
        BtnBackword.Enabled = True
        BtnBackword.Visible = True
        gbAdjOnStatus = True        ' �`�c�i��~��
    End Sub
#End Region

#Region "�X�e�b�v�ړ��{�^����\��"
    Public Sub StepMoveButtonOff()
        BtnForward.Enabled = False
        BtnForward.Visible = False
        BtnBackword.Enabled = False
        BtnBackword.Visible = False
        gbAdjOnStatus = False       ' �`�c�i���~��
    End Sub
#End Region
#Region "�X�e�b�v�ړ��{�^������"
    Private Sub BtnForward_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnForward.Click
        Call UserBas.StepMove(1)
    End Sub

    Private Sub BtnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBackword.Click
        Call UserBas.StepMove(-1)
    End Sub
#End Region
    'V2.0.0.0�H��
    ''' <summary>
    ''' �đ���̊J�n�ʒu�̎w��{�^������
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnStartPosSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnStartPosSet.Click
        Try
            Dim fReStartPosSet As New formReStartPosSet
            If giAppMode <> APP_MODE_IDLE Then
                Return
            Else
                giAppMode = APP_MODE_EDIT                               ' �A�v�����[�h = �f�[�^�ҏW
            End If

            ' �g���}���쒆�M��ON���M(���۰�ް)
            Call Sub_ATLDSET(COM_STS_TRM_STATE, 0)                      ' ���[�_�[�o��(ON=�g���}���쒆,OFF=�Ȃ�)

            fReStartPosSet.ShowDialog(Me)

            giAppMode = APP_MODE_IDLE                                   ' ����Ӱ�� = �g���}���u�A�C�h����
            Call Sub_ATLDSET(0, COM_STS_TRM_STATE)                      ' ���[�_�[�o��(ON=�Ȃ�,OFF=�g���}���쒆)

        Catch ex As Exception
            MsgBox("frmMain.BtnStartPosSet_Click() TRAP ERROR = " + ex.Message)
        End Try
    End Sub
    'V2.0.0.0�H��

    'V2.0.0.0�E��
    Private Sub CbDigSwL_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles CbDigSwL.MouseWheel

        Dim eventArgs As HandledMouseEventArgs = DirectCast(e, HandledMouseEventArgs)
        eventArgs.Handled = True

    End Sub
    'V2.0.0.0�E��
    'V2.0.0.0�H��
#Region "�O���t�\��/��\���{�^������������"
    Public Sub changefrmDistStatus(ByVal DispOnOff As Integer)
        Try

            If (DispOnOff = 1) Then
                '���v�\����ON
                gObjFrmDistribute.Show()
                gObjFrmDistribute.RedrawGraph()  '###218 
                '�{�^���\���̕ύX
                'If (gSysPrm.stTMN.giMsgTyp = 0) Then
                '    chkDistributeOnOff.Text = "���Y�O���t�@��\��"
                'Else
                '    chkDistributeOnOff.Text = "Distribute OFF"
                'End If
                chkDistributeOnOff.Text = Form1_019
            Else
                '���v�\����OFF
                gObjFrmDistribute.hide()

                '�{�^���\���̕ύX
                'If (gSysPrm.stTMN.giMsgTyp = 0) Then
                '    chkDistributeOnOff.Text = "���Y�O���t�@�\��"
                'Else
                '    chkDistributeOnOff.Text = "Distribute ON"
                'End If
                chkDistributeOnOff.Text = Form1_020
            End If

            Exit Sub

        Catch ex As Exception
            MsgBox("changefrmDistStatus() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' ���v��ʃ{�^���̗L����������
    ''' </summary>
    ''' <param name="Flag"></param>
    ''' <remarks></remarks>
    Public Sub chkDistributeOnOffEnableSet(ByVal Flag As Boolean)
        Try
            chkDistributeOnOff.Enabled = Flag
            CCmb_DistributeResList.Enabled = Flag
        Catch ex As Exception
            MsgBox("chkDistributeOnOffEnableSet() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' ���v��ʕ\����\��
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub chkDistributeOnOff_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDistributeOnOff.CheckedChanged
        Try
            If chkDistributeOnOff.Checked = True Then
                ''���v�\����ON
                ' ���v��ʃ{�^����L������
                gObjFrmDistribute.cmdGraphSave.Enabled = True
                gObjFrmDistribute.cmdInitial.Enabled = True
                gObjFrmDistribute.cmdFinal.Enabled = True
                CCmb_DistributeResList.Enabled = False
                changefrmDistStatus(1)
            Else
                ''���v�\����OFF
                CCmb_DistributeResList.Enabled = True
                changefrmDistStatus(0)
            End If
        Catch ex As Exception
            MsgBox("chkDistributeOnOff_CheckedChanged() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Sub

    Private Sub CCmb_DistributeResList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CCmb_DistributeResList.SelectedIndexChanged
        Try
            If (Not gObjFrmDistribute Is Nothing) Then
                gObjFrmDistribute.SetDistributionResNo(CCmb_DistributeResList.SelectedIndex + 1)
                stPLT.DistributionResNo = CCmb_DistributeResList.SelectedIndex + 1
                StatisticalDataDisp()
            End If
        Catch ex As Exception
            MsgBox("CCmb_DistributeResList_SelectedIndexChanged() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' ���z�}�̃{�^�����\������\�����Ԃ��B
    ''' </summary>
    ''' <returns>�\��:True ��\��:False</returns>
    ''' <remarks></remarks>
    Public Function GetDistributeOnOffStatus() As Boolean
        Try
            If chkDistributeOnOff.Checked = True Then
                Return (True)
            Else
                Return (False)
            End If
        Catch ex As Exception
            MsgBox("GetDistributeOnOffStatus() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Function
    ''' <summary>
    ''' ���z�}�̕\���A��\��
    ''' </summary>
    ''' <param name="Flag">1:�\�� 0:��\��</param>
    ''' <remarks></remarks>
    Public Sub DistributeOnOff(ByVal Flag As Integer)
        Try
            If Flag = 1 Then
                chkDistributeOnOff.Checked = True
            Else
                chkDistributeOnOff.Checked = False
            End If
        Catch ex As Exception
            MsgBox("DistributeOnOff() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Sub

#End Region

#Region "���v�f�[�^�̕\���X�V"
    Public Sub StatisticalDataDisp()
        Try
            Dim JudgeMode As Integer = FINAL_TEST
            Dim dMin As Double, dMax As Double, dAve As Double, dDev As Double

            If (Not gObjFrmDistribute Is Nothing) Then

                Call gObjFrmDistribute.StatisticalDataGet(JudgeMode, stPLT.DistributionResNo, dMin, dMax, dAve, dDev)

                Me.LabelStaticNom.Text = stREG(GetRNumByCircuit(1, stPLT.DistributionResNo)).dblNOM.ToString(TARGET_DIGIT_DEFINE)
                If dMin.ToString(TARGET_DIGIT_DEFINE).Length > 11 Then
                    Me.LabelStaticMin.Text = dMin.ToString(TARGET_DIGIT_DEFINE).Substring(0, 11)
                Else
                    Me.LabelStaticMin.Text = dMin.ToString(TARGET_DIGIT_DEFINE)
                End If
                If dMax.ToString(TARGET_DIGIT_DEFINE).Length > 11 Then
                    Me.LabelStaticMax.Text = dMax.ToString(TARGET_DIGIT_DEFINE).Substring(0, 11)
                Else
                    Me.LabelStaticMax.Text = dMax.ToString(TARGET_DIGIT_DEFINE)
                End If
                If dAve.ToString(TARGET_DIGIT_DEFINE).Length > 11 Then
                    Me.LabelStaticAve.Text = dAve.ToString(TARGET_DIGIT_DEFINE).Substring(0, 11)
                Else
                    Me.LabelStaticAve.Text = dAve.ToString(TARGET_DIGIT_DEFINE)
                End If
                If dDev.ToString(TARGET_DIGIT_DEFINE).Length > 11 Then
                    Me.LabelStaticDev.Text = dDev.ToString(TARGET_DIGIT_DEFINE).Substring(0, 11)
                Else
                    Me.LabelStaticDev.Text = dDev.ToString(TARGET_DIGIT_DEFINE)
                End If
            Else
                Me.LabelStaticMin.Text = ""
                Me.LabelStaticMax.Text = ""
                Me.LabelStaticAve.Text = ""
                Me.LabelStaticDev.Text = ""
            End If

        Catch ex As Exception
            MsgBox("frmMain.StatisticalDataDisp() TRAP ERROR = " + ex.Message)
        End Try
    End Sub
#End Region
    'V2.0.0.0�H��

    'V2.1.0.0�A��
#Region "���[�U�L�����u���[�V����"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cmdLaserCalibration_Click(sender As System.Object, e As System.EventArgs) Handles cmdLaserCalibration.Click
        Try
            cmdLaserTeach_Calibration()
        Catch ex As Exception
            MsgBox("frmMain.cmdLaserCalibration_Click() TRAP ERROR = " + ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' ���[�U�p���[���j�^�����O�`�F�b�N�{�^��������
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ButtonLaserCalibration_Click(sender As System.Object, e As System.EventArgs) Handles ButtonLaserCalibration.Click
        Try
            Dim LaserCalibrationMode As Integer = UserSub.LaserCalibrationModeGet()

            Select Case (LaserCalibrationMode)
                Case POWER_CHECK_NONE
                    LaserCalibrationMode = POWER_CHECK_START
                Case POWER_CHECK_START
                    LaserCalibrationMode = POWER_CHECK_LOT
                Case POWER_CHECK_LOT
                    LaserCalibrationMode = POWER_CHECK_NONE
            End Select

            UserSub.LaserCalibrationModeSet(LaserCalibrationMode)

            UserSub.LaserCalibrationModeUpdate()

        Catch ex As Exception
            MsgBox("frmMain.ButtonLaserCalibration_Click() TRAP ERROR = " + ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' ���[�_���̉�ʕ\��
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnLoaderInfo_Click(sender As Object, e As EventArgs) Handles btnLoaderInfo.Click

        Try


            If IsNothing(objLoaderInfo) = True Then
                Return
            End If

            If btnLoaderInfo.BackColor = Color.LightGreen Then
                btnLoaderInfo.BackColor = SystemColors.Control
                objLoaderInfo.Hide()
                objLoaderInfo.saveLoaderInfoDisp = 0
            Else
                btnLoaderInfo.BackColor = Color.LightGreen
                objLoaderInfo.saveLoaderInfoDisp = 1
                ObjLoader.DispLoaderInfo()

                objLoaderInfo.Show(Me)
            End If

        Catch ex As Exception

        End Try


    End Sub

    'V2.2.0.0�D��
    ''' <summary>
    ''' LoaderInfo��ʂ̕\����Ԃ�ۑ��A�擾���邵�ď�Ԃ����킹��  
    ''' </summary>
    ''' <param name="mode"></param>
    ''' <returns></returns>
    Public Function ChkLoaderInfoDisp(ByVal mode As Integer) As Integer

        Try

            If giLoaderType = 0 Then
                Return 0
            End If

            If mode = 0 Then        ' ��\��
                If btnLoaderInfo.BackColor = Color.LightGreen Then
                    objLoaderInfo.saveLoaderInfoDisp = 1    ' 
                End If
                objLoaderInfo.Hide()
            Else                    ' �\��
                If objLoaderInfo.saveLoaderInfoDisp = 1 Then
                    objLoaderInfo.Show()
                End If
            End If

        Catch ex As Exception

        End Try

    End Function
    'V2.2.0.0�D��



    ''' <summary>
    ''' �J�b�g���̒�~�@�\�{�^��������������  'V2.2.0.0�E
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnCutStop_Click(sender As Object, e As EventArgs) Handles btnCutStop.Click
        Try
            If giCutStop = 0 Then
                Return
            End If

            If btnCutStop.BackColor = Color.Yellow Then
                btnCutStop.BackColor = SystemColors.Control
            Else
                btnCutStop.BackColor = Color.Yellow
            End If


        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "STOP�{�^���̏�Ԃ𔻒肷�� "
    ''' <summary>
    ''' STOP�{�^���̏�Ԃ𔻒肷��    'V2.2.0.0�E
    ''' </summary>
    ''' <returns></returns>
    Public Function JudgeStop() As Integer

        Try

            If btnCutStop.BackColor = Color.Yellow Then
                JudgeStop = True
            Else
                JudgeStop = False
            End If


        Catch ex As Exception

        End Try

    End Function
#End Region

#Region "�T�C�N����~�{�^�������������̏���"
    ''' <summary>�@
    ''' �T�C�N����~�{�^�������������̏����@'V2.2.0.0�F
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnCycleStop_Click(sender As Object, e As EventArgs) Handles btnCycleStop.Click

        Try

            If btnCycleStop.BackColor = Color.Yellow Then
                btnCycleStop.BackColor = SystemColors.Control
            Else
                btnCycleStop.BackColor = Color.Yellow
            End If


        Catch ex As Exception

        End Try
    End Sub

#End Region

#Region "CYCLE STOP�{�^���̏�Ԃ𔻒肷�� "
    ''' <summary>
    ''' CYCLE STOP�{�^���̏�Ԃ𔻒肷��    'V2.2.0.0�F
    ''' </summary>
    ''' <returns></returns>
    Public Function JudgeCycleStop() As Integer

        Try

            If giClcleStop = 0 Then
                Return 0
            End If

            If btnCycleStop.BackColor = Color.Yellow Then
                JudgeCycleStop = 1
            Else
                JudgeCycleStop = 0
            End If


        Catch ex As Exception

        End Try

    End Function
#End Region


    ''' <summary>
    ''' ���[�UOFF�{�^�����������Ƃ��̏��� 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnLaserOff_Click(sender As Object, e As EventArgs) Handles btnLaserOff.Click
        Dim ret As Integer
        Dim mode As Integer = 0

        Try

            If btnLaserOff.BackColor = Color.Red Then
                btnLaserOff.BackColor = SystemColors.Control
                giLaserOffMode = 0
                mode = 1
            Else
                btnLaserOff.BackColor = Color.Red
                giLaserOffMode = 1
                mode = 0
            End If
            ret = DefTrimFnc.SPLASER_EXTDIODESET(mode)

        Catch ex As Exception

        End Try


    End Sub

    ''' <summary>
    ''' ��ʏ�̃{�^������N�����v�̕ˊJ����
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub cmdClamp_Click(sender As Object, e As EventArgs) Handles cmdClamp.Click
        Dim r As Integer

        Try

            cmdClamp.BackColor = Color.Yellow
            cmdClamp.Enabled = False

            ' �ڕ���N�����vON   
            r = System1.ClampCtrl(gSysPrm, 1, 0)
            If (r <> cFRS_NORMAL) Then

            End If

            Sleep(500)

            ' �ڕ���N�����vOFF 
            r = System1.ClampCtrl(gSysPrm, 0, 0)
            If (r <> cFRS_NORMAL) Then

            End If

        Catch ex As Exception


        Finally
            cmdClamp.Enabled = True
            cmdClamp.BackColor = SystemColors.ButtonFace

        End Try



    End Sub
    'V2.1.0.0�A��

    'V2.2.1.7�B��
    ''' <summary>
    ''' �A���[���Ń}�[�N�󎚂��Ȃ�������̈ꗗ����ʃ��O�ɕ\�� 
    ''' </summary>
    Public Sub DispMarkAlarmList()
        Dim i As Integer

        Try
            'V2.2.1.7�E ��
            ' �}�[�N�󎚂Ŗ�����΃A���[���\�����Ȃ��B 
            If UserSub.IsTrimType5() <> True Then
                Return
            End If
            'V2.2.1.7�E ��

            If LotMarkingAlarmCnt > 0 Then

                Call Z_PRINT("�}�[�N�󎚎��A���[�����X�g" & vbCrLf)

                For i = 1 To LotMarkingAlarmCnt
                    Call Z_PRINT(gMarkAlarmList(i).AlarmTrimData & ":" & gMarkAlarmList(i).LotCount & "����" & vbCrLf)
                Next

            Else
                'Call Z_PRINT("�}�[�N�󎚁F�S����" & vbCrLf)
            End If

        Catch ex As Exception

        End Try


    End Sub
    'V2.2.1.7�B��

End Class

#Region "�e�R�}���h���s�T�u�t�H�[���p���ʃC���^�[�t�F�[�X"
''' <summary>�e�R�}���h���s�T�u�t�H�[���p���ʃC���^�[�t�F�[�X</summary>
''' <remarks>'V2.2.0.0�@</remarks>
Public Interface ICommonMethods
    ''' <summary>�T�u�t�H�[���������s</summary>
    ''' <returns>���s���� sGetReturn</returns>
    ''' <remarks>'V2.2.0.0�@</remarks>
    Function Execute() As Integer

    ''' <summary>�T�u�t�H�[��KeyDown���̏���</summary>
    ''' <param name="e"></param>
    Sub JogKeyDown(ByVal e As KeyEventArgs)

    ''' <summary>�T�u�t�H�[��KeyUp���̏���</summary>
    ''' <param name="e"></param>
    Sub JogKeyUp(ByVal e As KeyEventArgs)

    ''' <summary>�J�����摜�N���b�N�ʒu���摜�Z���^�[�Ɉړ����鏈��</summary>
    ''' <param name="distanceX">�摜�Z���^�[����̋���X</param>
    ''' <param name="distanceY">�摜�Z���^�[����̋���Y</param>
    Sub MoveToCenter(ByVal distanceX As Decimal, ByVal distanceY As Decimal)
End Interface
#End Region

