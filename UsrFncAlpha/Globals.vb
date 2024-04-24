'===============================================================================
'   Description : �O���[�o���萔�̒�`
'
'   Copyright(C): TOWA LASERFRONT CORP. 2018
'
'===============================================================================
Option Strict Off
Option Explicit On

Imports System.IO
Imports System.Reflection
Imports LaserFront.Trimmer.DllAbout
Imports LaserFront.Trimmer.DllManualTeach
Imports LaserFront.Trimmer.DllPassword
Imports LaserFront.Trimmer.DllProbeTeach
Imports LaserFront.Trimmer.DllSysPrm
Imports LaserFront.Trimmer.DllSystem
Imports LaserFront.Trimmer.DllTeach
Imports LaserFront.Trimmer.DllUtility
Imports LaserFront.Trimmer.DllVideo
Imports TrimClassLibrary
Imports DllPlcIf                        'V2.2.0.0�D 
Imports System.Windows.Forms.Control
Imports LaserFront.Trimmer.DefWin32Fnc
Imports System.Runtime.InteropServices      '@@@888

Imports LaserFront.Trimmer
Imports LaserFront.Trimmer.DllSysPrm.SysParam
Imports UsrFunc.FormEdit

Module Globals_define
#Region "�O���[�o���萔/�ϐ��̒�`"

    '   ���d�N���h�~Mutex�n���h��
    Public gmhUserPro As System.Threading.Mutex = New System.Threading.Mutex(False, Application.ProductName)

    '---------------------------------------------------------------------------
    '   �A�v���P�[�V������/�A�v���P�[�V�������/�A�v���P�[�V�������[�h
    '---------------------------------------------------------------------------
    '----- �����I���p�A�v���P�[�V���� -----
    Public Const APP_FORCEEND As String = "c:\Trim\ForceEndProcess.exe"

    '-------------------------------------------------------------------------------
    '   �t�@�C���p�X��
    '-------------------------------------------------------------------------------
    Public Const OCX_PATH As String = "c:\Trim\ocx\"       '----- OCX�o�^�p�X
    Public Const DLL_PATH As String = "c:\Trim\"            '----- DLL�o�^�p�X
    Public Const SYSPARAMPATH As String = "C:\TRIM\tky.ini"
    Public Const USER_SYSPARAMPATH As String = "C:\TRIM\UserFunc.ini"         'V2.1.0.0�A


    'COPYDATASTRUCT�\����
    Public Structure COPYDATASTRUCT
        Public dwData As Int32   '���M����32�r�b�g�l
        Public cbData As Int32        'lpData�̃o�C�g��
        Public lpData As String     '���M����f�[�^�ւ̃|�C���^(0���\)
    End Structure

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Function SendMessage(
                           ByVal hWnd As IntPtr,
                           ByVal wMsg As Int32,
                           ByVal wParam As Int32,
                           ByVal lParam As Int32) As Integer
    End Function


    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Function SendMessage(
                            ByVal hWnd As IntPtr,
                            ByVal wMsg As Int32,
                            ByVal wParam As Int32,
                            ByRef lParam As COPYDATASTRUCT) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Unicode, EntryPoint:="SendMessage")>
    Public Function SendMessageString(ByVal hWnd As IntPtr,
                                      ByVal wMsg As UInt32,
                                      ByVal wParam As Int32,
                                      <[In], MarshalAs(UnmanagedType.LPWStr)>
                                      lParam As String) As Integer
    End Function

    '-------------------------------------------------------------------------------
    '   �V�X�e���p�����[�^(�`����DllgSysPrm.dll�Œ�`)
    '-------------------------------------------------------------------------------
    Public DllSysPrmSysParam_definst As New DllSysPrm.SysParam
    Public gSysPrm As SYSPARAM_PARAM           ' �V�X�e���p�����[�^

    '-------------------------------------------------------------------------------
    '   �I�u�W�F�N�g��`
    '-------------------------------------------------------------------------------
    '----- Form1�N���X -----
    Public ObjMain As Form1                             ' Form1�N���X              ###lstLog

    '----- .NET��DLL -----
    Public ObjSys As SystemNET                          ' DllSystem.dll
    Public ObjUtl As Utility                            ' DllUtility.dll
    Public ObjHlp As HelpVersion                        ' DllAbout.dll
    Public ObjPas As Password                           ' DllPassword.dll
    Public ObjMTC As ManualTeach                        ' DllManualTeach.dll
    Public ObjTch As Teaching                           ' DllTeach.dll
    Public ObjPrb As Probe                              ' DllProbeTeach.dll
    Public ObjVdo As VideoLibrary                       ' DllVideo.dll
    'Public ObjPrt As Object                             ' OcxPrint.ocx
    Public ObjMON(32) As Object
    Public gparModules As MainModules                   ' �e�����\�b�h�ďo���I�u�W�F�N�g(DllSystem�p)
    Public TrimClassCommon As New TrimClassLibrary.Common()             ' ���ʊ֐�

    '-------------------------------------------------------------------------------
    '   �ő�l/�ŏ��l
    '-------------------------------------------------------------------------------
    Public Const cMAXOptFlgNUM As Short = 5                 ' OcxSystem�p���߲ٵ�߼�݂̐� (�ő�5��)

    '-------------------------------------------------------------------------------
    '   Lamp ON/OFF����p���ߔԍ�
    '-------------------------------------------------------------------------------
    Public Const LAMP_START As Short = 0
    Public Const LAMP_RESET As Short = 1
    Public Const LAMP_Z As Short = 2
    Public Const LAMP_HALT As Short = 5
    Public Const cSTS_HALTSW_ON As Integer = 4              ' HALT�X�C�b�`ON Switch���

    '-------------------------------------------------------------------------------
    ' �g���h�^�n�@�r�b�g��`
    '-------------------------------------------------------------------------------
    Public Const EXT_BIT0 As Integer = &H1
    Public Const EXT_BIT1 As Integer = &H2
    Public Const EXT_BIT2 As Integer = &H4
    Public Const EXT_BIT3 As Integer = &H8
    Public Const EXT_BIT4 As Integer = &H10
    Public Const EXT_BIT5 As Integer = &H20
    Public Const EXT_BIT6 As Integer = &H40
    Public Const EXT_BIT7 As Integer = &H80
    Public Const EXT_BIT8 As Integer = &H100
    Public Const EXT_BIT9 As Integer = &H200
    Public Const EXT_BIT10 As Integer = &H400
    Public Const EXT_BIT11 As Integer = &H800
    Public Const EXT_BIT12 As Integer = &H1000
    Public Const EXT_BIT13 As Integer = &H2000
    Public Const EXT_BIT14 As Integer = &H4000
    Public Const EXT_BIT15 As Integer = &H8000

    '-------------------------------------------------------------------------------
    ' �g��I/O�ėp��`
    '-------------------------------------------------------------------------------
    Public Const EXT_IN0 As UShort = &H10                   ' B04: ���[�U���t�P
    Public Const EXT_IN1 As UShort = &H20                   ' B05: ���[�U���t�Q
    Public Const EXT_IN2 As UShort = &H40                   ' B06: ���[�U���t�R
    Public Const EXT_IN3 As UShort = &H80                   ' B07: ���[�U���t�S

    '-------------------------------------------------------------------------------
    '   Interlock switch bits (ADR. 0x21E8)
    '-------------------------------------------------------------------------------
    Public Const BIT_SLIDECOVER_CLOSE As Short = &H100S     ' B8 : �ײ�޶�ް��
    Public Const BIT_SLIDECOVER_OPEN As Short = &H200S      ' B9 : �ײ�޶�ް�J
    Public Const BIT_EMERGENCY_SW As Short = &H400S         ' B10: �ϰ�ުݼ�SW
    Public Const BIT_EMERGENCY_RESET As Short = &H800S      ' B11: �ϰ�ުݼ�ؾ��
    Public Const BIT_INTERLOCK_DISABLE As Short = &H1000S   ' B12: ����ۯ�����SW
    Public Const BIT_SERVO_ALARM As Short = &H2000S         ' B13: ���ޱװ�
    Public Const BIT_COVER_CLOSE As Short = &H4000S         ' B14: ��ް��
    '                                                       ' B15: ��ް&�ײ�޶�ް��

    Public Const INTERLOCK_STS_DISABLE_NO As Short = (0)    ' �C���^�[���b�N���(�����Ȃ�)
    Public Const INTERLOCK_STS_DISABLE_PART As Short = (1)  ' �C���^�[���b�N�ꕔ�����i�X�e�[�W���\�j
    Public Const INTERLOCK_STS_DISABLE_FULL As Short = 2    ' �C���^�[���b�N�S����
    Public Const SLIDECOVER_OPEN As Short = (1)             ' Bit0 : �X���C�h�J�o�[�F�I�[�v��
    Public Const SLIDECOVER_CLOSE As Short = (2)            ' Bit1 : �X���C�h�J�o�[�F�N���[�Y
    Public Const SLIDECOVER_MOVING As Short = (4)           ' Bit2 : �X���C�h�J�o�[�F���쒆
    '----- �V�O�i���^���[������ -----  
    Public Const SIGTOWR_NORMAL As Short = 0                ' �W���R�F����
    Public Const SIGTOWR_SPCIAL As Short = 1                ' �S�F����(����)

    '-------------------------------------------------------------------------------
    '   �V�O�i���^���[�R�F����(�W��)SL432R/SL436R����
    '   �@�蓮�^�]�� ������������� ���_��(���_���A����, ���f�B(�蓮))
    '   �A�C���^�[���b�N���������� ���_��(H/W�Ő���)
    '   �B�e�B�[�`���O������������ ���_��
    '   �C���_���A�� ������������� �Γ_��
    '   �D����~�� ������������� �ԓ_���{�u�U�[�n�m �� H/W��������ׂȂ�
    '   �E�����^�]�� �������������
    '     ��)����^�]���@�@�@�F�Γ_��
    '     ��)�S�}�K�W���I�����F�ԓ_�Ł{�u�U�[�n�m
    '     ��)�A���[�����@�@�@�F�ԓ_�Ł{�u�U�[�n�m�i�A���A�D���D��j
    '-------------------------------------------------------------------------------
    '----- OUTPUT -----                                     ' ON���̈Ӗ�
    '                                                       ' B0 : ���g�p
    '                                                       ' :
    '                                                       ' B7 : ���g�p
    Public Const SIGOUT_GRN_ON As UShort = &H100            ' B8 : �Γ_��  (�����^�]��)
    Public Const SIGOUT_YLW_ON As UShort = &H200            ' B9 : ���_��  (�e�B�[�`���O��)
    Public Const SIGOUT_RED_ON As UShort = &H400            ' B10: �ԓ_��  (����~) �����g�p(H/W�Ő���)
    Public Const SIGOUT_GRN_BLK As UShort = &H800           ' B11: �Γ_��  (���_���A��)
    Public Const SIGOUT_YLW_BLK As UShort = &H1000          ' B12: ���_��  (�C���^�[���b�N������)
    Public Const SIGOUT_RED_BLK As UShort = &H2000          ' B13: �ԓ_��  (�ُ�/�S�}�K�W���I��) ��+�u�U�[�P
    Public Const SIGOUT_BZ1_ON As UShort = &H4000           ' B14: �u�U�[�P(�ُ�) ��+�ԓ_��
    '                                                       ' B15: ���g�p

    '-------------------------------------------------------------------------------
    '   �g���d�w�s�a�h�s(���16�r�b�g ADR. 213A)
    '   ���V�O�i���^���[�S�F����(����)
    '-------------------------------------------------------------------------------
    '----- OUTPUT -----                                     ' ON���̈Ӗ�
    '                                                       ' B0 (B16): ���g�p
    '                                                       ' :
    '                                                       ' B3 (B19): ���g�p
    '                                                       ' B4 (B20): ���g�p
    '                                                       ' :
    '                                                       ' B7 (B23): ���g�p

    Public Const EXTOUT_RED_ON As UShort = &H100            ' B8 (B24): �ԓ_��  (����~) �����g�p(H/W�Ő���)
    Public Const EXTOUT_RED_BLK As UShort = &H200           ' B9 (B25): �ԓ_��  (�ُ�) ��+�u�U�[�P
    Public Const EXTOUT_YLW_ON As UShort = &H400            ' B10(B26): ���F�_��(���_���A����, ���f�B(�蓮))
    Public Const EXTOUT_YLW_BLK As UShort = &H800           ' B11(B27): ���F�_��(���_���A��)

    Public Const EXTOUT_GRN_ON As UShort = &H1000           ' B12(B28): �Γ_��  (�����^�]��)
    Public Const EXTOUT_GRN_BLK As UShort = &H2000          ' B13(B29): �Γ_��  (-) �����g�p
    Public Const EXTOUT_BZ1_ON As UShort = &H4000           ' B14(B30): �u�U�[�P(�ُ�) ��+�ԓ_��
    '                                                       ' B15(B31): ���g�p

    '-------------------------------------------------------------------------------
    '   ���[�_�[�h�^�n�r�b�g(ADR. 219A)
    '-------------------------------------------------------------------------------
    '----- ��ϰ  �� ۰�ް -----
    ' ��Bit0�`Bit4,Bit7���W����
    Public Const COM_STS_TRM_STATE As Short = &H1S          ' B0 : �g���}��~(0:��~,1:���쒆)
    Public Const COM_STS_TRM_NG As Short = &H2S             ' B1 : �g���~���O�m�f(0:����, 1:NG)
    Public Const COM_STS_PTN_NG As Short = &H4S             ' B2 : �p�^�[���F���G���[(0:����, 1:�G���[)
    Public Const COM_STS_TRM_ERR As Short = &H8S            ' B3 : �g���}�G���[(0:����, 1:�G���[)
    Public Const COM_STS_TRM_READY As Short = &H10S         ' B4 : �g���}���f�B(0:ɯ���ި, 1:��ި)
    Public Const COM_STS_LOT_END As Short = &H20S           ' B5 : ���b�g�I��(0:������, 1:�I�����)�@'V1.2.0.0�C�@
    Public Const COM_STS_ABS_ON As Short = &H40S            ' B6 : �z��(0:�I��, 1:�I�t)�@�@�@�@�@�@�@'V1.2.0.0�C
    Public Const COM_STS_CLAMP_ON As Short = &H80S          ' B7 : �ڕ������ߊJ��(0:��, 1:�J)�@�@�@ 'V1.2.0.0�C
    'V1.2.0.0�C    '                                                       ' B5 : ���g�p
    'V1.2.0.0�C    '                                                       ' B6 : ���g�p
    'V1.2.0.0�C    Public Const COM_STS_ABS_ON As Short = &H80S            ' B7 : �ڕ������ߊJ��(0:��, 1:�J)

    '----- ۰�ް  �� ��ϰ -----
    ' ��Bit0�`Bit3�܂ł��W����
    Public Const cHSTcRDY As Short = 1                      ' B0 : ���۰�ް�L��(0:��, 1:�L)
    Public Const cHSTcAUTO As Short = 2                     ' B1 : ���۰�ްӰ��(1=����Ӱ��, 0=�蓮Ӱ��)
    Public Const cHSTcSTATE As Short = 4                    ' B2 : ���۰�ް���쒆(0=���쒆, 1=��~)
    Public Const cHSTcTRMCMD As Short = 8                   ' B3 : ��ϰ����(1=��ϰ����) ��ׯ�
    Public Const cHSTcLOTCHANGE As Short = &H10             ' B4 : ���b�g�؂�ւ��M��
    Public Const cHSTcABS_ON As Short = &H40S               ' B6 : �z��(0:�I��, 1:�I�t)�@�@�@�@�@�@�@'V1.2.0.0�C
    Public Const cHSTcCLAMP_ON As Short = &H80S             ' B7 : �ڕ������ߊJ��(0:��, 1:�J)�@�@�@ 'V1.2.0.0�C

    Public gdwATLDDATA As UInteger                          ' ���[�_�o�̓f�[�^
    Public gDebugHostCmd As UInteger                        ' ���[�_���̓f�[�^(���ޯ�ޗp)
    Public gwPrevHcmd As UInteger                           ' ���[�_���̓f�[�^�ޔ���
    Public gbClampOpen As Boolean = True                    'V1.2.0.0�C �N�����v�J�\���:True ���ɊJ:False
    Public gbVaccumeOff As Boolean = True                   'V1.2.0.0�C �z���I�t�\���:True ���ɃI�t:False

    '-------------------------------------------------------------------------------
    '   gMode(frmReset�̏������[�h) ��100�`266�͖߂�l�ɂ��g�p���܂�(-101�`-266)
    '-------------------------------------------------------------------------------
    Public gFRsetFlg As Short                               ' frmReset�׸�(0:�����l, 1:frmReset������)
    Public gMode As Short                                   ' �������[�h�ޔ���

    Public Const cGMODE_ORG As Short = 0                    '  0 : ���_���A
    Public Const cGMODE_ORG_MOVE As Short = 1               '  1 : ���_�ʒu�ړ�
    Public Const cGMODE_START_RESET As Short = 2            '  2 : ����m�F���(START/RESET�҂�)
    '                                                       '  3 :
    '                                                       '  4 :
    Public Const cGMODE_EMG As Short = 5                    '  5 : ����~���b�Z�[�W�\��
    '                                                       '  6 :
    Public Const cGMODE_SCVR_OPN As Short = 7               '  7 : �g���~���O���̃X���C�h�J�o�[�J���b�Z�[�W�\��
    Public Const cGMODE_CVR_OPN As Short = 8                '  8 : �g���~���O����➑̃J�o�[�J���b�Z�[�W�\��
    Public Const cGMODE_SCVRMSG As Short = 9                '  9 : �X���C�h�J�o�[�J���b�Z�[�W�\��(�g���~���O���ȊO)
    Public Const cGMODE_CVRMSG As Short = 10                ' 10 : ➑̃J�o�[�J�m�F���b�Z�[�W�\��(�g���~���O���ȊO)
    Public Const cGMODE_ERR_HW As Short = 11                ' 11 : �n�[�h�E�F�A�G���[(�J�o�[�����Ă܂�)���b�Z�[�W�\��
    Public Const cGMODE_ERR_HW2 As Short = 12               ' 12 : �n�[�h�E�F�A�G���[���b�Z�[�W�\��
    Public Const cGMODE_CVR_LATCH As Short = 13             ' 13 : �J�o�[�J���b�`���b�Z�[�W�\��
    Public Const cGMODE_CVR_CLOSEWAIT As Short = 14         ' 14 : ➑̃J�o�[�N���[�Y�������̓C���^�[���b�N�����҂�

    Public Const cGMODE_ERR_DUST As Short = 20              ' 20 : �W�o�@�ُ팟�o���b�Z�[�W�\��
    Public Const cGMODE_ERR_AIR As Short = 21               ' 21 : �G�A�[���G���[���o���b�Z�[�W�\��

    Public Const cGMODE_ERR_HING As Short = 40              ' 40 : �A��HI-NG�װ(ADV�������҂�)
    Public Const cGMODE_SWAP As Short = 41                  ' 41 : �����(START�������҂�)
    Public Const cGMODE_XYMOVE As Short = 42                ' 42 : �I������ð��وړ��m�F(START�������҂�)
    Public Const cGMODE_LDR_ALARM As Short = 44             ' 44 : ���[�_�A���[������                  'V2.2.0.0�D 
    Public Const cGMODE_LDR_START As Short = 45             ' 45 : �����^�]�J�n(START�������҂�)       'V2.2.0.0�D 
    Public Const cGMODE_LDR_TMOUT As Short = 46             ' 46 : ���[�_�ʐM�^�C���A�E�g              'V2.2.0.0�D 
    Public Const cGMODE_LDR_END As Short = 47               ' 47 : �����^�]�I��(START�������҂�)       'V2.2.0.0�D   

    Public Const cGMODE_AUTO_LASER As Short = 50            ' 50 : �������[�U�p���[����

    Public Const cGMODE_LDR_CHK As Short = 60               ' 60 : ���[�_��ԃ`�F�b�N(�N����۰�ގ���Ӱ��/���쒆)
    Public Const cGMODE_LDR_ERR As Short = 61               ' 61 : ���[�_��ԃG���[(۰�ގ�����۰�ޖ�)
    Public Const cGMODE_LDR_MNL As Short = 62               ' 62 : �J�o�[�J��̃��[�_�蓮���[�h����
    Public Const cGMODE_LDR_WKREMOVE As Short = 63          ' 63 : �c���菜�����b�Z�[�W     'V2.2.0.0�D
    Public Const cGMODE_LDR_RSTAUTO As Short = 64           ' 64 : �����^�]���~���b�Z�[�W      'V2.2.0.0�D
    Public Const cGMODE_LDR_WKREMOVE2 As Short = 65         ' 65 : �c���菜�����b�Z�[�W(APP�I��)  'V2.2.0.0�D

    Public Const cGMODE_LDR_CHK_AUTO As Short = 67          ' 63 : ���[�_��ԃ`�F�b�N(�����^�]��),���[�_�������ɐ؂�ւ��܂ő҂�'V1.0.4.3�K
    Public Const cGMODE_LDR_STAGE_ORG As Short = 66         ' 66 : �X�e�[�W���_�ړ�v

    Public Const cGMODE_OPT_START As Short = 70             ' 70 : ���ݸފJ�n���̽���SW�����҂�
    Public Const cGMODE_OPT_END As Short = 71               ' 71 : ���ݸޏI�����̽ײ�޶�ް�J�҂�

    Public Const cGMODE_MSG_DSP As Short = 90               ' 90 : �w�胁�b�Z�[�W�\��(ADV�������҂�)

    ' ���~�b�g�Z���T�[& ���G���[ & �^�C���A�E�g���b�Z�[�W
    Public Const cGMODE_TO_AXISX As Short = 101             ' 101: X���G���[(�^�C���A�E�g)
    Public Const cGMODE_TO_AXISY As Short = 102             ' 102: Y���G���[(�^�C���A�E�g)
    Public Const cGMODE_TO_AXISZ As Short = 103             ' 103: Z���G���[(�^�C���A�E�g)
    Public Const cGMODE_TO_AXIST As Short = 104             ' 104: �Ǝ��G���[(�^�C���A�E�g)
    '�y�\�t�g���~�b�g�G���[�z
    Public Const cGMODE_SL_AXISX As Short = 105             ' 105: X���\�t�g���~�b�g�G���[
    Public Const cGMODE_SL_AXISY As Short = 106             ' 106: Y���\�t�g���~�b�g�G���[
    Public Const cGMODE_SL_AXISZ As Short = 107             ' 107: Z���\�t�g���~�b�g�G���[
    Public Const cGMODE_SL_BPX As Short = 110               ' 110: BP X���\�t�g���~�b�g�G���[
    Public Const cGMODE_SL_BPY As Short = 111               ' 111: BP Y���\�t�g���~�b�g�G���[

    Public Const cGMODE_TO_ROTATT As Short = 108            ' 108: ���[�^���A�b�e�l�[�^�G���[(�^�C���A�E�g)
    Public Const cGMODE_TO_AXISZ2 As Short = 109            ' 109: Z2���G���[(�^�C���A�E�g)

    Public Const cGMODE_SRV_ARM As Short = 202              ' 202: �T�[�{�A���[��
    Public Const cGMODE_AXISX_LIM As Short = 203            ' 203: X�����~�b�g
    Public Const cGMODE_AXISY_LIM As Short = 204            ' 204: Y�����~�b�g
    Public Const cGMODE_AXISZ_LIM As Short = 205            ' 205: Z�����~�b�g
    Public Const cGMODE_AXIST_LIM As Short = 206            ' 206: �Ǝ����~�b�g
    Public Const cGMODE_RATT_LIM As Short = 207             ' 207: ���[�^���[�A�b�e�l�[�^���~�b�g
    Public Const cGMODE_AXISZ2_LIM As Short = 208           ' 208: Z2�����~�b�g

    Public Const cGMODE_BASE_ERR As Short = 200             ' ���G���[�x�[�X�ԍ�
    '�yX���G���[�z
    Public Const cGMODE_AXISX_AOFF As Short = 211           ' 211: X���G���[(Bit All Off)
    Public Const cGMODE_AXISX_AON As Short = 212            ' 212: X���G���[(Bit All On)
    Public Const cGMODE_AXISX_ARM As Short = 213            ' 213: X���A���[��
    Public Const cGMODE_AXISX_PML As Short = 214            ' 214: �}X�����~�b�g
    Public Const cGMODE_AXISX_PLM As Short = 215            ' 215: +X�����~�b�g
    Public Const cGMODE_AXISX_MLM As Short = 216            ' 216: -X�����~�b�g
    '�yY���G���[�z
    Public Const cGMODE_AXISY_AOFF As Short = 221           ' 221: Y���G���[(Bit All Off)
    Public Const cGMODE_AXISY_AON As Short = 222            ' 222: Y���G���[(Bit All On)
    Public Const cGMODE_AXISY_ARM As Short = 223            ' 223: Y���A���[��
    Public Const cGMODE_AXISY_PML As Short = 224            ' 224: �}Y�����~�b�g
    Public Const cGMODE_AXISY_PLM As Short = 225            ' 225: +Y�����~�b�g
    Public Const cGMODE_AXISY_MLM As Short = 226            ' 226: -Y�����~�b�g
    '�yZ���G���[�z
    Public Const cGMODE_AXISZ_AOFF As Short = 231           ' 231: Z���G���[(Bit All Off)
    Public Const cGMODE_AXISZ_AON As Short = 232            ' 232: Z���G���[(Bit All On)
    Public Const cGMODE_AXISZ_ARM As Short = 233            ' 233: Z���A���[��
    Public Const cGMODE_AXISZ_PML As Short = 234            ' 234: �}Z�����~�b�g
    Public Const cGMODE_AXISZ_PLM As Short = 235            ' 235: +Z�����~�b�g
    Public Const cGMODE_AXISZ_MLM As Short = 236            ' 236: -Z�����~�b�g
    Public Const cGMODE_AXISZ_ORG As Short = 237            ' 237: Z�����_���A������
    '�y�Ǝ��G���[�z
    Public Const cGMODE_AXIST_AOFF As Short = 241           ' 241: �Ǝ��G���[(Bit All Off)
    Public Const cGMODE_AXIST_AON As Short = 242            ' 242: �Ǝ��G���[(Bit All On)
    Public Const cGMODE_AXIST_ARM As Short = 243            ' 243: �Ǝ��A���[��
    Public Const cGMODE_AXIST_PML As Short = 244            ' 244: �}�Ǝ����~�b�g
    Public Const cGMODE_AXIST_PLM As Short = 245            ' 245: +�Ǝ����~�b�g
    Public Const cGMODE_AXIST_MLM As Short = 246            ' 246: -�Ǝ����~�b�g
    '�yZ2���G���[�z
    Public Const cGMODE_AXISZ2_AOFF As Short = 251          ' 251: Z2���G���[(Bit All Off)
    Public Const cGMODE_AXISZ2_AON As Short = 252           ' 252: Z2���G���[(Bit All On)
    Public Const cGMODE_AXISZ2_ARM As Short = 253           ' 253: Z2���A���[��
    Public Const cGMODE_AXISZ2_PML As Short = 254           ' 254: �}Z2�����~�b�g
    Public Const cGMODE_AXISZ2_PLM As Short = 255           ' 255: +Z2�����~�b�g
    Public Const cGMODE_AXISZ2_MLM As Short = 256           ' 256: -Z2�����~�b�g
    Public Const cGMODE_AXISZ2_ORG As Short = 257           ' 257: Z2�����_���A������
    '�y۰�ر��Ȱ���װ�z
    Public Const cGMODE_ROTATT_AOFF As Short = 261          ' 261: ۰�ر��Ȱ���װ(Bit All Off)
    Public Const cGMODE_ROTATT_AON As Short = 262           ' 262: ۰�ر��Ȱ���װ(Bit All On)
    Public Const cGMODE_ROTATT_ARM As Short = 263           ' 263: ۰�ر��Ȱ���װ�
    Public Const cGMODE_ROTATT_PML As Short = 264           ' 264: �}۰�ر��Ȱ���Я�
    Public Const cGMODE_ROTATT_PLM As Short = 265           ' 265: +۰�ر��Ȱ���Я�
    Public Const cGMODE_ROTATT_MLM As Short = 266           ' 266: -۰�ر��Ȱ���Я�

    '-------------------------------------------------------------------------------
    '   ���샍�O���b�Z�[�W
    '-------------------------------------------------------------------------------
    '----- ���샍�O���b�Z�[�W -----
    Public MSG_OPLOG_START As String                        ' "���[�U�v���O�����N��"
    Public MSG_OPLOG_FUNC01 As String                       ' "�f�[�^���[�h"
    Public MSG_OPLOG_FUNC02 As String                       ' "�f�[�^�Z�[�u"
    Public MSG_OPLOG_FUNC03 As String                       ' "�f�[�^�ҏW"
    Public MSG_OPLOG_FUNC04 As String                       ' "�}�X�^�`�F�b�N"(����)
    Public MSG_OPLOG_FUNC05 As String                       ' "���[�U����"
    Public MSG_OPLOG_FUNC06 As String                       ' "���b�g�ؑ�"
    Public MSG_OPLOG_FUNC07 As String                       ' "�v���[�u�ʒu���킹"
    Public MSG_OPLOG_FUNC07_2 As String                     ' "�v���[�u�ʒu���킹2"
    Public MSG_OPLOG_FUNC08 As String                       ' "�e�B�[�`���O"
    Public MSG_OPLOG_FUNC08S As String                      ' "�J�b�g�␳�ʒu�e�B�[�`���O"
    Public MSG_OPLOG_FUNC09 As String                       ' "�p�^�[���o�^"
    Public MSG_OPLOG_FUNC10 As String                       ' "�v���[�u�ʒu���킹�Q"
    Public MSG_OPLOG_FUNC11 As String                       ' "�f�[�^�ݒ�"
    Public MSG_OPLOG_END As String                          ' "���[�U�v���O�����I��"
    Public MSG_OPLOG_TRIMST As String                       ' "�g���~���O"
    Public MSG_OPLOG_LOTCHG As String                       ' "���b�g�ؑ֐M����M"
    Public MSG_OPLOG_STOP As String                         ' "�g���}���u��~"
    Public MSG_OPLOG_LOTSET As String                       ' "���b�g���f�[�^�ݒ�"

    '----- ���b�Z�[�W -----
    Public MSG_DataNotLoad As String                        ' �f�[�^�����[�h
    Public MSG_SPRASH31 As String
    Public MSG_SPRASH32 As String
    Public MSG_SPRASH52 As String
    Public MSG_105 As String
    Public MSG_136 As String
    Public MSG_137 As String
    Public MSG_138 As String
    Public MSG_139 As String
    Public MSG_140 As String
    Public MSG_141 As String
    Public MSG_142 As String
    Public MSG_143 As String
    Public MSG_144 As String
    Public MSG_145 As String
    Public MSG_146 As String
    Public MSG_147 As String
    Public MSG_148 As String
    Public MSG_149 As String
    Public MSG_150 As String
    Public MSG_151 As String
    Public MSG_152 As String
    Public MSG_153 As String

    ' �s�w�C�s�x�֌W�@START
    ' ��frmMsgBox ��ʏI���m�F
    Public MSG_CLOSE_LABEL01 As String
    Public MSG_CLOSE_LABEL02 As String
    Public MSG_CLOSE_LABEL03 As String
    Public MSG_EXECUTE_TXTYLABEL As String 'TX,TY
    Public TITLE_TX As String '�`�b�v�T�C�Y(TX)�e�B�[�`���O
    Public TITLE_TY As String '�X�e�b�v�T�C�Y(TY)�e�B�[�`���O
    Public LBL_TXTY_TEACH_03 As String '�␳��
    Public LBL_TXTY_TEACH_04 As String '�␳�䗦
    Public LBL_TXTY_TEACH_05 As String '���߻��� (mm)
    Public LBL_TXTY_TEACH_07 As String '�␳�O
    Public LBL_TXTY_TEACH_08 As String '�␳��
    Public LBL_TXTY_TEACH_09 As String '��ٰ�߲������(mm)
    Public LBL_TXTY_TEACH_11 As String '�X�e�b�v�C���^�[�o��(mm)(�ǉ�)
    Public LBL_TXTY_TEACH_12 As String '��P��_
    Public LBL_TXTY_TEACH_13 As String '��Q��_
    Public LBL_TXTY_TEACH_14 As String '�O���[�v
    Public LBL_CMD_CANCEL As String
    Public CMD_CANCEL As String '�L�����Z��
    Public INFO_MSG13 As String '"�`�b�v�T�C�Y�@�e�B�[�`���O"
    Public INFO_MSG14 As String '"�X�e�b�v�ԃC���^�[�o���@�e�B�[�`���O"��"�X�e�[�W�O���[�v�Ԋu�e�B�[�`���O"
    Public INFO_MSG15 As String '"�X�e�b�v�I�t�Z�b�g�ʁ@�e�B�[�`���O"
    Public INFO_MSG16 As String '"��ʒu�����킹�ĉ������B"
    Public INFO_MSG17 As String '"�ړ�:[���]  ����:[START]  ���f:[RESET]" & vbCrLf & "[HALT]�łP�O�̏����ɖ߂�܂��B"
    Public INFO_MSG18 As String '"��1�O���[�v�A��1��R��ʒu�̃e�B�[�`���O"
    Public INFO_MSG19 As String '"��"
    Public INFO_MSG20 As String '"�O���[�v�A�ŏI��R��ʒu�̃e�B�[�`���O"
    Public INFO_MSG23 As String '"�O���[�v�ԃC���^�[�o���@�e�B�[�`���O"��"�a�o�O���[�v�Ԋu�e�B�[�`���O"
    Public INFO_MSG28 As String '"�O���[�v�A�ŏI�[�ʒu�̃e�B�[�`���O"
    Public INFO_MSG29 As String '"�O���[�v�A�Ő�[�ʒu�̃e�B�[�`���O"
    Public INFO_MSG30 As String '"�T�[�L�b�g�Ԋu�e�B�[�`���O"
    Public INFO_MSG31 As String '"�X�e�b�v�I�t�Z�b�g�ʂ̃e�B�[�`���O"
    Public INFO_MSG32 As String '"(�s�w)"   '###084
    Public INFO_MSG33 As String '"(�s�x)"   '###084
    Public INFO_MSG34 As String '"�X�e�b�v�T�C�Y�@�e�B�[�`���O"
    ' �s�w�C�s�x�֌W�@END

    '----- �摜�\���v���O�����̕\���ʒu -----
    'Public Const FORM_X As Integer = 4                                  ' �R���g���[���㕔���[���WX
    'Public Const FORM_Y As Integer = 20                                 ' �R���g���[���㕔���[���WY
    Public Const FORM_X As Integer = 0                                  ' �R���g���[���㕔���[���WX
    Public Const FORM_Y As Integer = 0                                  ' �R���g���[���㕔���[���WY

    '----- �摜�\���v���O�����̋N���p -----
    Public Const DISPGAZOU_PATH As String = "C:\TRIM\DispGazouSmall.exe"    ' �摜�\���v���O������
    Public Const DISPGAZOU_WRK As String = "C:\\TRIM"                       ' ��ƃt�H���_

    '----- �n�� -----
    Public Const MACHINE_TYPE_SL432 As String = "SL432R"
    Public Const MACHINE_TYPE_SL436 As String = "SL436R"

    '----- ���O��ʕ\���p -----�@
    Public gDspClsCount As Integer = 5                                  ' ���O��ʕ\���N���A�����
    Public gDspCounter As Integer = 0                                   ' ���O��ʕ\��������J�E���^

    Public gPlateCount As Integer = 0                                   ' �����(�f�o�b�O�p)

    '----- GPIB ����p -----                               
    Public gGpibMultiMeterCount As Integer = 5                          ' �O��������IT�Ƒ���ł̑���񐔁i����l�����肵�Ȃ��̂ōŌ�̑���l���g�p����B�j

    '----- EXTOUT LED����r�b�g -----                               
    Public glLedBit As Long                                             ' LED����r�b�g(EXTOUT) 
    Public Const INITIAL_TEST As Integer = 0                ' �����e�X�g
    Public Const FINAL_TEST As Integer = 1                  ' �ŏI�e�X�g

    Public Const SETAXISSPDY_DEFALT As UInteger = 15000    ' �x���X�e�[�W���x�̕ύX�@�\�����l 'V2.0.0.0�N

    Public giStageYDir As Integer = 1                       ' �X�e�[�WY�̈ړ�����(CW(1), CCW(-1))    'V2.2.0.0�@ 

    '---------------------------------------------------------------------------
    ' �g���~���O���샂�[�h
    '---------------------------------------------------------------------------
    '^^^^^ �f�B�W�^��SW HI�@��` -----
    Public Const DGSW_HI_NODISP As Integer = 0              ' �\���Ȃ�
    Public Const DGSW_HI_NGDISP As Integer = 1              ' �m�f�̂ݕ\��
    Public Const DGSW_HI_DISP As Integer = 2                ' �S�ĕ\��

    '----- �f�B�W�^��SW LOW�@��` -----
    Public Const TRIM_MODE_ITTRFT As Integer = 0            ' �C�j�V�����e�X�g�{�g���~���O�{�t�@�C�i���e�X�g���s
    Public Const TRIM_MODE_MEAS As Integer = 1              ' ������s
    Public Const TRIM_MODE_CUT As Integer = 2               ' �J�b�g���s
    Public Const TRIM_MODE_STPRPT As Integer = 3            ' �X�e�b�v�����s�[�g���s
    Public Const TRIM_MODE_MEAS_MARK As Integer = 4         ' 'V1.0.4.3�I����}�[�L���O���[�h�E�t�@�C�i������̂�
    Public Const TRIM_MODE_POWER As Integer = 5             ' �d�����[�h 'V2.0.0.0�A
    Public Const TRIM_VARIATION_MEAS As Integer = 6         ' ����l�ϓ����� 'V2.0.0.0�A

    'Public Const TRIM_MODE_TRFT As Integer = 1              ' �g���~���O�{�t�@�C�i���e�X�g���s
    'Public Const TRIM_MODE_FT As Integer = 2                ' �t�@�C�i���e�X�g���s�i����j
    'Public Const TRIM_MODE_MEAS As Integer = 3              ' ������s
    'Public Const TRIM_MODE_POSCHK As Integer = 4            ' �|�W�V�����`�F�b�N
    'Public Const TRIM_MODE_CUT As Integer = 5               ' �J�b�g���s
    'Public Const TRIM_MODE_STPRPT As Integer = 6            ' �X�e�b�v�����s�[�g���s

    '---------------------------------------------------------------------------
    ' SLIDE COVER+XY�ړ���������
    '---------------------------------------------------------------------------
    Public Const TYPE_OFFLINE As Short = 0                  ' OFFLINE
    Public Const TYPE_ONLINE As Short = 1                   ' ONLINE
    Public Const TYPE_MANUAL As Short = 2                   ' SLIDE COVER+XY�ړ���������

    '---------------------------------------------------------------------------
    ' �J�b�g���샂�[�h �i0:�g���~���O�A1:�e�B�[�`���O�A2:�����J�b�g�j
    '---------------------------------------------------------------------------
    Public Const TRIM_MODE As Integer = 0                   ' �X�g���[�g�@�g���~���O���[�h
    Public Const TEACH_MODE As Integer = 1                  ' �X�g���[�g�@�e�B�[�`���O���[�h
    Public Const FORCE_MODE As Integer = 2                  ' �X�g���[�g�@�����J�b�g���[�h


    Public Const CUT_MODE_NORMAL As Integer = 0             ' �m�[�}��
    Public Const CUT_MODE_RETURN As Integer = 1             ' ���^�[���J�b�g
    Public Const CUT_MODE_RETRACE As Integer = 2            ' ���g���[�X�J�b�g
    Public Const CUT_MODE_NANAME As Integer = 4             ' �΂߃J�b�g

    '---------------------------------------------------------------------------
    ' �J�b�g���@ �i1:�ׯ�ݸށA2:INDEX�A3:NG ) 
    '---------------------------------------------------------------------------
    Public Const CNS_CUTM_TR As Integer = 1                 ' �g���b�L���O
    Public Const CNS_CUTM_IX As Integer = 2                 ' �C���f�b�N�X
    Public Const CNS_CUTM_NG As Integer = 3                 ' �m�f
    Public Const CNS_CUTM_NON_POS_IX As Integer = 4         ' �|�W�V���j���O�����C���f�b�N�X

    '---------------------------------------------------------------------------
    ' �J�b�g�`�� �i1:ST�J�b�g�A2:L�J�b�g�A3:SP�J�b�g 4:IX�J�b�g�j
    '---------------------------------------------------------------------------
    'Public Const CNS_CUTP_ST As Integer = 1                 ' ST�J�b�g
    'Public Const CNS_CUTP_L As Integer = 2                  ' L�J�b�g
    'Public Const CNS_CUTP_SP As Integer = 3                 ' SP�J�b�g
    'Public Const CNS_CUTP_IX As Integer = 4                 ' IX�J�b�g
    'Public Const CNS_CUTP_M As Integer = 19                 ' �����}�[�L���O�@###1042�@ 
    Public Const CNS_CUTP_NORMAL As Integer = 0             ' �m�[�}���J�b�g�E�J�b�g���[�h�w��p 'V1.0.4.3�G
    Public Const CNS_CUTP_ST As Integer = 1                 ' ST�J�b�g
    Public Const CNS_CUTP_ST_TR As Integer = 2              ' V1.1.0.0�B�X�g���[�g�E���g���[�X(RETRACE)�J�b�g 
    Public Const CNS_CUTP_L As Integer = 3                  ' V1.1.0.0�BL�J�b�g
    Public Const CNS_CUTP_M As Integer = 4                  ' V1.1.0.0�B�����}�[�L���O�@###1042�@ 
    Public Const CNS_CUTP_U As Integer = 5                  ' U�J�b�g�ǉ��@ 'V2.2.0.0�A
    Public Const CNS_CUTP_SP As Integer = 6                 ' SP�J�b�g V1.1.0.0�B�ԍ��ύX   'V2.2.0.0�A5->6
    Public Const CNS_CUTP_IX As Integer = 7                 ' IX�J�b�g V1.1.0.0�B�ԍ��ύX   'V2.2.0.0�A6->7

    '---------------------------------------------------------------------------
    ' ���胂�[�h �i0:��R����A1:�d������A2:�O������i�f�o�h�a�j�j
    '---------------------------------------------------------------------------
    Public Const MEAS_MODE_RESISTOR As Integer = 0          ' ��R����
    Public Const MEAS_MODE_VOLTAGE As Integer = 1           ' �d������
    Public Const MEAS_MODE_EXTERNAL As Integer = 2          ' �O������i�f�o�h�a�j

    '---------------------------------------------------------------------------
    ' ���萸�x �i0:��������A1:�����x����j
    '---------------------------------------------------------------------------
    Public Const MEAS_TYP_FAST As Integer = 0               ' ��������
    Public Const MEAS_TYP_HIPRECI As Integer = 1            ' �����x����


    '---------------------------------------------------------------------------
    ' ���萸�x �i0:��������A1:�����x����j
    '---------------------------------------------------------------------------
    Public Const MEAS_RNGSET_AUTO As Integer = 0            ' �I�[�g�����W�ݒ�
    Public Const MEAS_RNGSET_FIX_TAR As Integer = 1         ' �Œ背���W�ݒ�-�ڕW�l�ݒ�
    Public Const MEAS_RNGSET_FIX_NO As Integer = 2          ' �Œ背���W�ݒ�-�����W�ԍ��ݒ�

    'V1.0.4.3�E��
    '---------------------------------------------------------------------------
    ' �p�^�[���F��(0:����, 1:�L��, 2:�蓮, 3:�����m�f���肠��j
    '---------------------------------------------------------------------------
    Public Const CUT_PATTERN_NONE As Integer = 0            ' 0:����
    Public Const CUT_PATTERN_AUTO As Integer = 1            ' 1:�L��
    Public Const CUT_PATTERN_MANUAL As Integer = 2          ' 2:�蓮
    Public Const CUT_PATTERN_AUTO_NG As Integer = 3         ' 3:�����m�f���肠��

    '-------------------------------------------------------------------------------
    '   �X���[�v��`��`    
    '-------------------------------------------------------------------------------
    Public Const SLP_VTRIMPLS As Integer = 1                ' �{�d���g���~���O
    Public Const SLP_VTRIMMNS As Integer = 2                ' �|�d���g���~���O
    Public Const SLP_RTRM As Integer = 4                    ' �@��R�g���~���O
    Public Const SLP_VMES As Integer = 5                    ' �@�d������
    Public Const SLP_RMES As Integer = 6                    ' �@��R����
    Public Const SLP_NG_MARK As Integer = 7                 ' �@�m�f�}�[�N
    Public Const SLP_OK_MARK As Integer = 8                 ' �@�n�j�}�[�N
    Public Const SLP_ATRIMPLS As Integer = 9                ' �{�d���g���~���O
    Public Const SLP_ATRIMMNS As Integer = 10               ' �|�d���g���~���O
    Public Const SLP_AMES As Integer = 11                   ' �@�d������
    Public Const SLP_MARK As Integer = 12                   ' �@�}�[�N�� 'V2.2.1.7�@
    'V1.0.4.3�E��
    'V1.0.4.3�F��
    Public Const DEF_DIR_CW As Integer = 1                  ' ���v�����̉�]�iClock Wise)
    Public Const DEF_DIR_CCW As Integer = 2                 ' �����v�����̉�]�iCounter Clock Wise)
    'V1.0.4.3�F��

    'V1.0.4.3�F��
    ' �p���[����(FL�p)
    Public Const CUT_CND_L1 As Integer = 1              ' L1���H�����ݒ�
    Public Const CUT_CND_L2 As Integer = 2              ' L2���H�����ݒ�
    Public Const CUT_CND_L3 As Integer = 3              ' L3���H�����ݒ�
    Public Const CUT_CND_L4 As Integer = 4             ' L4���H�����ݒ�
    'V1.0.4.3�F��

    '-------------------------------------------------------------------------------
    '   �t�@�C�o�[���[�U�p��`
    '-------------------------------------------------------------------------------
    '----- ���U���� -----
    Public Const OSCILLATOR_FL As Integer = 3               ' FL(̧��ްڰ��)
    Public Const OSCILLATOR_SP As Integer = 5               ' SP���[�U

#If cOSCILLATORcFLcUSE Then
    Public stCND As TrimCondInfo                            ' �g���}�[���H����(�`����`��Rs232c.vb�Q��)

    '----- RS232C�|�[�g����`
    Public stCOM As ComInfo                                 ' �|�[�g���(�`����`��Rs232c.vb�Q��)
#End If
    Public Const cTIMEOUT As Long = 10000                   ' �����҃^�C�}�l(ms)

    '----- FL�����f�t�H���g�ݒ�t�@�C�� ----
    Public Const DEF_FLPRM_SETFILEPATH As String = "c:\TRIM\"
    Public Const DEF_FLPRM_SETFILENAME As String = "c:\TRIM\defaultFlParamSet.xml"

    Public giTenKey_Btn As Short = 0                        ' �ꎞ��~��ʂł́uTen Key On/Off�v�{�^���̏����l(0:ON(����l), 1:OFF)
    Public giBpAdj_HALT As Short = 0                        ' �ꎞ��~��ʂł́uBP�I�t�Z�b�g��������/���Ȃ��v(0:��������(����l), 1:�������Ȃ�)
    Public Const BLOCK_END As Short = 1         ' �u���b�N�I�� 
    Public Const PLATE_BLOCK_END As Short = 2   ' �v���[�g�E�u���b�N�I��

    ' �s�w�C�s�x�֌W�@START
    Public Const MaxCntStep As Short = 256                  ' �ï�ߍő匏��
    Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
    Public Const HWND_TOPMOST As Short = -1                 ' �E�B���h�E���őO�ʂɕ\��
    Public Const SWP_NOSIZE As Short = &H1S                 ' ���݂̃T�C�Y���ێ�
    Public Const SWP_NOMOVE As Short = &H2S                 ' ���݂̈ʒu���ێ�
    Public Const KND_CHIP As Short = 1
    Public Const MaxCntCut As Short = MAXCTN                ' �ő嶯Đ�
    Public Const MaxCutInfo As Short = MAXCTN               ' �ő嶯ď��
    Public Const CNS_CUTP_ST2 As String = "T"               ' �߼޼��ݸޖ���ST�J�b�g
    Public Const CNS_CUTP_IX2 As String = "X"               ' �߼޼��ݸޖ������ޯ��
    ' �s�w�C�s�x�֌W�@END

    '----- �p���[���[�^�̃f�[�^�擾�擾 -----
    Public Const PM_DTTYPE_NONE As Short = 0                ' �Ȃ�
    Public Const PM_DTTYPE_IO As Short = 1                  ' �h�^�n�ǎ��
    Public Const PM_DTTYPE_USB As Short = 2                 ' �t�r�a

    Public INTERNAL_CAMERA As Integer = 0                   ' �������
    Public EXTERNAL_CAMERA As Integer = 1                   ' �O�����

    'V2.0.0.0�M��
    Public Const CLAMP_VACCUME_USE As Short = 1             '�N�����v�z���L��
    Public Const CLAMP_ONLY_USE As Short = 2                '�N�����v�̂�
    Public Const VACCUME_ONLY_USE As Short = 3              '�z���̂�
    'V2.0.0.0�M��

#Region "�O���[�o���ϐ��̒�`"
    Public gCurPlateNo As Integer
    Public gCurBlockNo As Integer
    Public gbExitFlg As Boolean
    Public gbTenKeyFlg As Boolean = True                            ' �e���L�[���̓t���O
    Public gbChkboxHalt As Boolean = True                           ' ADJ�{�^�����(ON=ADJ ON, OFF=ADJ OFF)
    Public gObjADJ As frmFineAdjust = Nothing                              ' �ꎞ��~��ʃI�u�W�F�N�g
    '-------------------------------------------------------------------------------
    '   �f�o�h�a�ʐM�p��`
    '-------------------------------------------------------------------------------
    Public ObjGpib As GpibMaster = Nothing                  ' �f�o�h�a�ʐM�p�I�u�W�F�N�g
    Public gstrDeviceName As String = "GPIB000"             ' �f�o�h�a�f�o�C�X��(�f�o�C�X�}�l�[�W���Œ�`�������O)
    Public gDevId As Short = -1                             ' �f�o�C�X�h�c
    Public gEOI As Short = 1                                ' EOI(0:�o�͂��Ȃ�, 0�ȊO:�o�͂���) 2013/3/9 0����1�֕ύX

    '-------------------------------------------------------------------------------
    '   ���̑��̒�`
    '-------------------------------------------------------------------------------
    '----- �A�v���P�[�V������ʒ�` -----  
    Public Const KND_USER As Integer = 9                    ' ���[�U�v���O����
    Public frmAutoObj As FormDataSelect                     ' �����^�]Form��޼ު��
    Public ObjCrossLine As New TrimClassLibrary.TrimCrossLineClass()

    Public Const TARGET_DIGIT_DEFINE As String = "0.0000000"      'V2.0.0.0�D

    Public ObjLoader As clsLoaderIf = Nothing                               'V2.2.0.0�D
    Public ObjPlcIf As DllPlcIf.DllMelsecPLCIf                              'V2.2.0.0�D
    Public objLoaderInfo As frmLoaderInfo                                   'V2.2.0.0�D
    Public swMesureTrimtime As New System.Diagnostics.Stopwatch()           '�������Ԃ̌v���p       'V2.2.0.0�D
    Public gdTrimtime As New TimeSpan                                       '�g���~���O���ԕۑ��p   'V2.2.0.0�D
    Public gitacktTime As Integer                                           '�^�N�g�^�C���ۑ��p     'V2.2.0.0�D
    Public gichangePlateTime As Integer                                     '��������ԕۑ��p     'V2.2.0.0�D
    Public MarkingCount As Integer                                          '�}�[�L���O��������J�E���g�p    'V2.2.1.7�B


    '-------------------------------------------------------------------------------
    ' �������^�]..
    '-------------------------------------------------------------------------------
    Public MSG_AUTO_01 As String '���샂�[�h
    Public MSG_AUTO_02 As String '�}�K�W�����[�h
    Public MSG_AUTO_03 As String '���b�g���[�h
    Public MSG_AUTO_04 As String '�G���h���X���[�h
    Public MSG_AUTO_05 As String '�f�[�^�t�@�C��
    Public MSG_AUTO_06 As String '�o�^�ς݃f�[�^�t�@�C��
    Public MSG_AUTO_07 As String '���X�g��1���
    Public MSG_AUTO_08 As String '���X�g��1����
    Public MSG_AUTO_09 As String '���X�g����폜
    Public MSG_AUTO_10 As String '���X�g���N���A
    Public MSG_AUTO_11 As String '�o�^
    Public MSG_AUTO_12 As String 'OK
    Public MSG_AUTO_13 As String '�L�����Z��
    Public MSG_AUTO_14 As String '�f�[�^�I��'
    Public MSG_AUTO_15 As String '�o�^���X�g��S�č폜���܂��B
    Public MSG_AUTO_16 As String '��낵���ł����H
    Public MSG_AUTO_17 As String '�G���h���X���[�h���͕����̃f�[�^�t�@�C���͑I���ł��܂���B
    Public MSG_AUTO_18 As String '�f�[�^�t�@�C����I�����Ă��������B
    Public MSG_AUTO_19 As String '�ҏW���̃f�[�^��ۑ����܂����H
    Public MSG_AUTO_20 As String '���H�����t�@�C�������݂��܂���B

    ' �s�w�C�s�x�֌W�@START
    Public gCmpTrimDataFlg As Short                         ' �f�[�^�X�V�t���O(0=�X�V�Ȃ�, 1=�X�V����)
    Public gTkyKnd As Short = KND_CHIP                      ' �A�v���P�[�V�������
    Public gfCorrectPosX As Double                          ' �ƕ␳����XYð��ق����X(mm) ��ThetaCorrection()�Őݒ�
    Public gfCorrectPosY As Double                          ' �ƕ␳����XYð��ق����Y(mm)
    ' �s�w�C�s�x�֌W�@END

    ' �p���[����(FL�p)
    Public MSG_AUTOPOWER_01 As String
    Public MSG_AUTOPOWER_02 As String
    Public MSG_AUTOPOWER_03 As String
    Public MSG_AUTOPOWER_04 As String
    Public MSG_AUTOPOWER_05 As String

    ' ���z�}
    Public PIC_TRIM_01 As String '�C�j�V�����e�X�g�@���z�}
    Public PIC_TRIM_02 As String '�t�@�C�i���e�X�g�@���z�}
    Public PIC_TRIM_03 As String '�Ǖi
    Public PIC_TRIM_04 As String '�s�Ǖi
    Public PIC_TRIM_05 As String '�ŏ�%
    Public PIC_TRIM_06 As String '�ő�%
    Public PIC_TRIM_07 As String '����%
    Public PIC_TRIM_08 As String '�W���΍�
    Public PIC_TRIM_09 As String '��R��
    Public PIC_TRIM_10 As String '���z�}�ۑ� 
    Public MSG_TRIM_04 As String '�C�j�V�����e�X�g�@���z�}
    Public MSG_TRIM_05 As String '�t�@�C�i���e�X�g�@���z�}

    '-------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------
    '   �J�b�g�ʒu�␳��`(�␳���s)    
    '-------------------------------------------------------------------------------
    Public Const PTN_NONE As Integer = 0                    ' �␳���s 0:�Ȃ�
    Public Const PTN_AUTO As Integer = 1                    ' �␳���s 1:����
    Public Const PTN_MANUAL As Integer = 2                  ' �␳���s 2:�蓮
    Public Const PTN_AUTO_JUDGE As Integer = 3              ' �␳���s 3:�����m�f���肠��

    '-------------------------------------------------------------------------------
    '   ���胂�[�h��`    
    '-------------------------------------------------------------------------------
    Public Const MEAS_JUDGE_NONE As Integer = 0             ' ����Ȃ�
    Public Const MEAS_JUDGE_IT As Integer = 1               ' IT�̂�
    Public Const MEAS_JUDGE_FT As Integer = 2               ' FT�̂�
    Public Const MEAS_JUDGE_BOTH As Integer = 3             ' IT,FT����
#End Region

#Region "���蔻�胂�[�h"
    Public Const JUDGE_MODE_RATIO As Integer = 0            ' 0:�䗦(%)
    Public Const JUDGE_MODE_ABSOLUTE As Integer = 1         ' 1:���l(��Βl)
#End Region

#Region "�s�w�A�s�x�p�\���̒�`"

    '-------------------------------------------------------------------------------
    '   �s�w�A�s�x�p�v���[�g�f�[�^
    '-------------------------------------------------------------------------------
    Public Structure PlateInfo
        Dim intBlockCntXDir As Short                        ' ��ۯ����w
        Dim intBlockCntYDir As Short                        ' ��ۯ����x
        Dim dblBlockSizeXDir As Double                      ' �u���b�N�T�C�Y�w   
        Dim dblBlockSizeYDir As Double                      ' �u���b�N�T�C�Y�x   
        Dim dblTableOffsetXDir As Double                    ' ð��وʒu�̾��X
        Dim dblTableOffsetYDir As Double                    ' ð��وʒu�̾��Y
        Dim dblBpOffSetXDir As Double                       ' �ްшʒu�̾��X
        Dim dblBpOffSetYDir As Double                       ' �ްшʒu�̾��Y
        Dim intResistDir As Short                           ' ��R���ѕ���
        Dim intResistCntInBlock As Short                    ' 1�u���b�N����R��
        Dim intResistCntInGroup As Short                    ' 1�O���[�v����R��
        Dim intGroupCntInBlockXBp As Short                  ' �u���b�N���a�o�O���[�v��(�T�[�L�b�g��)
        Dim intGroupCntInBlockYStage As Short               ' �u���b�N���X�e�[�W�O���[�v��
        Dim dblChipSizeXDir As Double                       ' ���߻���X
        Dim dblChipSizeYDir As Double                       ' ���߻���Y
        Dim dblStepOffsetXDir As Double                     ' �ï�ߵ̾�ė�X
        Dim dblStepOffsetYDir As Double                     ' �ï�ߵ̾�ė�Y
        Dim dblBpGrpItv As Double                           ' BP�O���[�v�Ԋu�i�ȑO��CHIP�̃O���[�v�Ԋu�j
        Dim dblStgGrpItvX As Double                         ' X�����̃X�e�[�W�O���[�v�Ԋu�i�ȑO�̂b�g�h�o�̃X�e�b�v�ԃC���^�[�o���j
        Dim dblStgGrpItvY As Double                         ' Y�����̃X�e�[�W�O���[�v�Ԋu�i�ȑO�̂b�g�h�o�̃X�e�b�v�ԃC���^�[�o���j
        Dim intBlkCntInStgGrpX As Short                     ' X�����̃X�e�[�W�O���[�v���u���b�N��
        Dim intBlkCntInStgGrpY As Short                     ' Y�����̃X�e�[�W�O���[�v���u���b�N��
    End Structure
    Public typPlateInfo As PlateInfo                        ' ��ڰ��ް�
    '--------------------------------------------------------------------------
    '   �J�b�g�f�[�^�\���̌`����`
    '--------------------------------------------------------------------------
    Public Structure CutList
        Dim intCutNo As Short                               ' ��Ĕԍ�(1�`n)
        Dim dblStartPointX As Double                        ' �����߲��X
        Dim dblStartPointY As Double                        ' �����߲��Y
        Dim dblTeachPointX As Double                        ' è��ݸ��߲��X
        Dim dblTeachPointY As Double                        ' è��ݸ��߲��Y
        Dim strCutType As String                            ' ��Č`��
        Dim intCutAngle As Short                            ' �J�b�g�p�x     'V2.2.0.0�A
        Dim intLTurnDir As Short                            ' �^�[������     'V2.2.0.0�A
    End Structure
    '--------------------------------------------------------------------------
    '   ��R�f�[�^�\���̌`����`
    '--------------------------------------------------------------------------
    Public Structure ResistorInfo
        Dim intResNo As Short                               ' ��R�ԍ�(1�`9999)
        Dim intCutCount As Short                            ' ��Đ�
        <VBFixedArray(MaxCutInfo)> Dim ArrCut() As CutList  ' ��ď��
        ' �\���̂̏�����
        Public Sub Initialize()
            ReDim ArrCut(MaxCutInfo)
        End Sub
    End Structure

    Public typResistorInfoArray(MAXRNO) As ResistorInfo     ' ��R�ް�

#End Region

#End Region

    'V2.0.0.0�H��
#Region "���z�}"
    '----- ���Y�Ǘ��O���t�t�H�[���I�u�W�F�N�g
    Public gObjFrmDistribute As Object                          ' frmDistribute

#End Region
    'V2.0.0.0�H��

    '=========================================================================
    '   ���b�Z�[�W�ݒ菈��
    '=========================================================================
#Region "���b�Z�[�W�����ݒ菈��"
    '''=========================================================================
    '''<summary>���b�Z�[�W�����ݒ菈��</summary>
    '''<param name="language">(INP) 0=���{��, 1=�p��</param>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub PrepareMessages(ByVal language As Short)

        Select Case language
            Case 0
                Call PrepareMessagesJapanese()
            Case 1
                Call PrepareMessagesEnglish()
            Case Else
                Call PrepareMessagesEnglish()
        End Select

    End Sub
#End Region

#Region "���b�Z�[�W�����ݒ�(���{��)"
    '''=========================================================================
    '''<summary>���b�Z�[�W�����ݒ�(���{��)</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub PrepareMessagesJapanese()

        ' �G���[���b�Z�[�W
        MSG_DataNotLoad = "�f�[�^�������[�h�ł��B�f�[�^�����[�h���ĉ������B" & vbCrLf
        MSG_SPRASH31 = "���ӁI�I�I"
        MSG_SPRASH32 = "�X���C�h�J�o�[�������ŕ��܂�"
        MSG_SPRASH52 = "�X���C�h�J�o�[�����Ă��܂��B" & ControlChars.NewLine & "�g���~���O���J�n���܂��B"

        MSG_105 = "�O�̉�ʂɖ߂�܂��B��낵���ł����H�@�@�@�@�@�@�@�@�@�@�@�@"

        MSG_136 = "�V���A���|�[�g�n�o�d�m�G���["
        MSG_137 = "�V���A���|�[�g�b�k�n�r�d�G���["
        MSG_138 = "�V���A���|�[�g���M�G���["
        MSG_139 = "�V���A���|�[�g��M�G���["
        MSG_140 = "�e�k���̉��H�����̐ݒ肪����܂���B" + vbCrLf + "�ēx�f�[�^�����[�h���邩�A�ҏW��ʂ�����H�����̐ݒ���s���Ă��������B"
        MSG_141 = "�e�k�����H�����̃��[�h�Ɏ��s���܂����B"
        MSG_142 = "���H�����t�@�C�����쐬���܂���"
        MSG_143 = "�f�[�^�����[�h���܂���"
        MSG_144 = "�f�[�^���[�h�m�f"
        MSG_145 = "�f�[�^���Z�[�u���܂���"
        MSG_146 = "�f�[�^�Z�[�u�m�f"
        MSG_147 = "�e�k�։��H�����𑗐M���܂����B"
        MSG_148 = "�e�k�փf�[�^���M���E�E�E�E�E�E"
        MSG_150 = "�e�k�ʐM�ُ�B�e�k�Ƃ̒ʐM�Ɏ��s���܂����B" + vbCrLf + "�e�k�Ɛ������ڑ��ł��Ă��邩�m�F���Ă��������B"
        MSG_151 = "���H�����̐ݒ�Ɏ��s���܂����B"
        MSG_152 = "���H�����̑��M�Ɏ��s���܂����B" + vbCrLf + "�ēx�f�[�^�����[�h���邩�A�ҏW��ʂ�����H�����̐ݒ���s���Ă��������B"
        MSG_153 = "�J�b�g�ʒu�␳�Ώۂ̒�R������܂���"

        ' ���샍�O�@���b�Z�[�W
        MSG_OPLOG_START = "���[�U�v���O�����N��"
        MSG_OPLOG_FUNC01 = "�f�[�^���[�h"
        MSG_OPLOG_FUNC02 = "�f�[�^�Z�[�u"
        MSG_OPLOG_FUNC03 = "�f�[�^�ҏW"
        MSG_OPLOG_FUNC04 = "�}�X�^�`�F�b�N"
        MSG_OPLOG_FUNC05 = "���[�U����"
        MSG_OPLOG_FUNC06 = "���b�g�ؑ�"
        MSG_OPLOG_FUNC07 = "�v���[�u�ʒu���킹"
        MSG_OPLOG_FUNC08 = "�e�B�[�`���O"
        MSG_OPLOG_FUNC08S = "�J�b�g�␳�ʒu�e�B�[�`���O"
        MSG_OPLOG_FUNC09 = "�p�^�[���o�^"
        MSG_OPLOG_FUNC10 = "�v���[�u�ʒu���킹�Q"
        MSG_OPLOG_FUNC11 = "�f�[�^�ݒ�"
        MSG_OPLOG_END = "���[�U�v���O�����I��"
        MSG_OPLOG_TRIMST = "�g���~���O"
        MSG_OPLOG_LOTCHG = "���b�g�ؑ֐M����M"
        MSG_OPLOG_STOP = "�g���}���u��~"
        MSG_OPLOG_LOTSET = "���b�g���f�[�^�ݒ�"

        ' �������^�]..
        MSG_AUTO_01 = "���샂�[�h"
        MSG_AUTO_02 = "�}�K�W�����[�h"
        MSG_AUTO_03 = "���b�g���[�h"
        MSG_AUTO_04 = "�G���h���X���[�h"
        MSG_AUTO_05 = "�f�[�^�t�@�C��"
        MSG_AUTO_06 = "�o�^�ς݃f�[�^�t�@�C��"
        MSG_AUTO_07 = "���X�g��1���"
        MSG_AUTO_08 = "���X�g��1����"
        MSG_AUTO_09 = "���X�g����폜"
        MSG_AUTO_10 = "���X�g���N���A"
        MSG_AUTO_11 = "���o�^��"
        MSG_AUTO_12 = "OK"
        MSG_AUTO_13 = "�L�����Z��"
        MSG_AUTO_14 = "�f�[�^�o�^"
        MSG_AUTO_15 = "�o�^���X�g��S�č폜���܂��B"
        MSG_AUTO_16 = "��낵���ł����H"
        MSG_AUTO_17 = "�G���h���X���[�h���͕����̃f�[�^�t�@�C���͑I���ł��܂���B"
        MSG_AUTO_18 = "�f�[�^�t�@�C����I�����Ă��������B"
        MSG_AUTO_19 = "�ҏW���̃f�[�^��ۑ����܂����H"
        MSG_AUTO_20 = "���H�����t�@�C�������݂��܂���B"

        ' �s�w�C�s�x�֌W�@START
        ' frmMsgBox(��ʏI���m�F)
        MSG_CLOSE_LABEL01 = "��ʏI���m�F"
        MSG_CLOSE_LABEL02 = "�͂�(&Y)"
        MSG_CLOSE_LABEL03 = "������(&N)"
        TITLE_TX = "�`�b�v�T�C�Y(TX)�e�B�[�`���O"
        TITLE_TY = "�X�e�b�v�T�C�Y(TY)�e�B�[�`���O"
        LBL_TXTY_TEACH_03 = "�␳��"
        LBL_TXTY_TEACH_04 = "�␳�䗦"
        LBL_TXTY_TEACH_05 = "�`�b�v�T�C�Y(mm)"
        LBL_TXTY_TEACH_07 = "�␳�O"
        LBL_TXTY_TEACH_08 = "�␳��"
        LBL_TXTY_TEACH_09 = "�O���[�v�C���^�[�o��"
        LBL_TXTY_TEACH_11 = "�X�e�b�v�C���^�[�o��"
        LBL_TXTY_TEACH_12 = "��P��_"
        LBL_TXTY_TEACH_13 = "��Q��_"
        LBL_TXTY_TEACH_14 = "�O���[�v"
        LBL_CMD_CANCEL = "�L�����Z�� (&Q)"
        CMD_CANCEL = "�L�����Z��"
        INFO_MSG13 = "�`�b�v�T�C�Y�@�e�B�[�`���O"
        INFO_MSG14 = "�X�e�[�W�O���[�v�Ԋu�e�B�[�`���O"
        INFO_MSG15 = "�X�e�b�v�I�t�Z�b�g�ʁ@�e�B�[�`���O"
        INFO_MSG16 = "�@�@��ʒu�����킹�ĉ������B"
        INFO_MSG17 = "�@�@�ړ��F[���]  ����F[START]  ���f�F[RESET]" '& vbCrLf & "�@�@[HALT]�łP�O�̏����ɖ߂�܂��B"
        INFO_MSG18 = "��1�O���[�v�A��1��R��ʒu�̃e�B�[�`���O"
        INFO_MSG19 = "��"
        INFO_MSG20 = "�O���[�v�A�ŏI��R��ʒu�̃e�B�[�`���O"
        INFO_MSG23 = "�a�o�O���[�v�Ԋu�e�B�[�`���O"
        INFO_MSG28 = "�O���[�v�A�ŏI�[�ʒu�̃e�B�[�`���O"
        INFO_MSG29 = "�O���[�v�A�Ő�[�ʒu�̃e�B�[�`���O"
        INFO_MSG30 = "�T�[�L�b�g�Ԋu�e�B�[�`���O"
        INFO_MSG31 = "�X�e�b�v�I�t�Z�b�g�ʂ̃e�B�[�`���O"
        INFO_MSG32 = " (�s�w)"  '###084
        INFO_MSG33 = " (�s�x)"  '###084
        INFO_MSG34 = "�X�e�b�v�T�C�Y�@�e�B�[�`���O"
        ' �s�w�C�s�x�֌W�@END

        ' �p���[����(FL�p)
        MSG_AUTOPOWER_01 = "�p���[�����J�n"
        MSG_AUTOPOWER_02 = "���H�����ԍ�"
        MSG_AUTOPOWER_03 = "���[�U�p���[�ݒ�l"
        MSG_AUTOPOWER_04 = "�d���l"
        MSG_AUTOPOWER_05 = "�p���[����������"

        ' ���z�}
        MSG_TRIM_04 = "�C�j�V�����e�X�g�@���z�}"
        MSG_TRIM_05 = "�t�@�C�i���e�X�g�@���z�}"
        PIC_TRIM_01 = "�C�j�V�����e�X�g�@���z�}"
        PIC_TRIM_02 = "�t�@�C�i���e�X�g�@���z�}"
        PIC_TRIM_03 = "�Ǖi"
        PIC_TRIM_04 = "�s�Ǖi"
        PIC_TRIM_05 = "�ŏ�%"
        PIC_TRIM_06 = "�ő�%"
        PIC_TRIM_07 = "����%"
        PIC_TRIM_08 = "�W���΍�"
        PIC_TRIM_09 = "��R��"
        PIC_TRIM_10 = "���z�}�ۑ�"
    End Sub
#End Region

#Region "���b�Z�[�W�����ݒ�(�p��)"
    '''=========================================================================
    '''<summary>���b�Z�[�W�����ݒ�(�p��)</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub PrepareMessagesEnglish()

        ' �G���[���b�Z�[�W
        MSG_DataNotLoad = "Data is not loaded. Please Load the data file." & vbCrLf
        MSG_SPRASH31 = "Cautions !!!"
        MSG_SPRASH32 = "Slide Cover Closes Automatically."
        MSG_SPRASH52 = "Slide Cover is closed." & ControlChars.NewLine & "Trimming will be started."           'V1.0.0.1�A

        MSG_136 = "Serial Port Open Error."
        MSG_137 = "Serial Port Close Error."
        MSG_138 = "Serial Port Transmission Error."
        MSG_139 = "Serial Port Reception Error."
        MSG_140 = "There Is No Setting On The FL Side." + vbCrLf + "Please Load Data Or Set Condition From Edit Function."
        MSG_141 = "Condition Reading Error On The FL Side."
        MSG_142 = "Condition File Was Made."
        MSG_143 = "DATA LOAD OK"
        MSG_144 = "DATA LOAD NG"
        MSG_145 = "DATA SAVE OK"
        MSG_146 = "DATA SAVE NG"
        MSG_147 = "DATA SEND TO FL"
        MSG_148 = "Data sending to FL......"
        MSG_150 = "Connection error for FiberLaser." + vbCrLf + "Please confirm the connection."
        MSG_151 = "It Failed In The Setting Of Processing Conditions."
        MSG_152 = "It Failed In The Transmission Of The Condition Data." + vbCrLf + "Please Load Data Or Set Condition From Edit Function."
        MSG_153 = "No resistor to correct cutting position."

        ' ���샍�O�@���b�Z�[�W
        MSG_OPLOG_START = "START USER PROGURAM"
        MSG_OPLOG_FUNC01 = "LOAD"
        MSG_OPLOG_FUNC02 = "SAVE"
        MSG_OPLOG_FUNC03 = "EDIT"
        MSG_OPLOG_FUNC04 = "MASTER CHECK"
        MSG_OPLOG_FUNC05 = "LASER"
        MSG_OPLOG_FUNC06 = "LOT CHANGE"
        MSG_OPLOG_FUNC07 = "PROBE"
        MSG_OPLOG_FUNC08 = "TEACH"
        MSG_OPLOG_FUNC08S = "CUTTING POSITION CORRECTION TEACHING"
        MSG_OPLOG_FUNC09 = "RECOG"
        MSG_OPLOG_FUNC10 = "PROBE2"
        MSG_OPLOG_FUNC11 = "DATA SET"
        MSG_OPLOG_END = "END USER PROGURAM"
        MSG_OPLOG_TRIMST = "TRIMMING"
        MSG_OPLOG_LOTCHG = "LOT CHANGE RECEIVE"
        MSG_OPLOG_STOP = "TRIMMER STOP"
        MSG_OPLOG_LOTSET = "LOT DATA INPUT"

        ' �s�w�C�s�x�֌W�@START
        'frmMsgBox(��ʏI���m�F)
        MSG_CLOSE_LABEL01 = "Exit?"
        MSG_CLOSE_LABEL02 = "Yes(&Y)"
        MSG_CLOSE_LABEL03 = "No(&N)"
        TITLE_TX = "Chip size (TX) Teaching"
        TITLE_TY = "Step size (TY) Teaching"
        LBL_TXTY_TEACH_03 = "Correct quantity"
        LBL_TXTY_TEACH_04 = "Correct ratio"
        LBL_TXTY_TEACH_05 = "Chip size(mm)"
        LBL_TXTY_TEACH_07 = "Before"
        LBL_TXTY_TEACH_08 = "After"
        LBL_TXTY_TEACH_09 = "Group interval(mm)"
        LBL_TXTY_TEACH_11 = "Step interval"
        LBL_TXTY_TEACH_12 = "The 1st datum point."
        LBL_TXTY_TEACH_13 = "The 2nd datum point."
        LBL_TXTY_TEACH_14 = "Group"
        LBL_CMD_CANCEL = "Cancel (&Q)"
        CMD_CANCEL = "Cancel"
        INFO_MSG13 = "CHIP SIZE TEACHING"
        INFO_MSG14 = "STAGE INTERVAL TEACHING"
        INFO_MSG15 = "STEP OFFSET TEACHING"
        INFO_MSG16 = "    Please unite a standard position."
        INFO_MSG17 = "    MOVE:[Arrow]  OK:[START]  CANCEL:[RESET]" '& vbCrLf & "    It returns to the processing before one by the HALT key."
        INFO_MSG18 = "<Group No.1> The 1st resistance standard position." ''''2009/07/03 NET�ł́uresistance��circuit�v(18,20-22)
        INFO_MSG19 = "<Group No."
        INFO_MSG20 = "> The last resistance standard position."
        INFO_MSG23 = "BP GROUP INTERVAL TEACHING"
        INFO_MSG28 = "> The Final Edge Positionlast."
        INFO_MSG29 = "> The State-Of-The-Art Position."
        INFO_MSG30 = "CIRCUIT INTERVAL TEACHING"
        INFO_MSG31 = "STEP OFFSET TEACHING"
        INFO_MSG32 = " (TX)"  '###084
        INFO_MSG33 = " (TY)"  '###084
        INFO_MSG33 = "STEP SIZE TEACHING"
        ' �s�w�C�s�x�֌W�@END

        ' �p���[����(FL�p)
        MSG_AUTOPOWER_01 = "Start Power Adjustment"
        MSG_AUTOPOWER_02 = "Condition No."
        MSG_AUTOPOWER_03 = "Laser Power"
        MSG_AUTOPOWER_04 = "Current"
        MSG_AUTOPOWER_05 = "Power Adjustment Failed."

        ' ���z�}
        MSG_TRIM_04 = "INITIAL TEST DISTRIBUTION MAP"
        MSG_TRIM_05 = "FINAL TEST DISTRIBUTION MAP"
        PIC_TRIM_01 = "INITIAL TEST DISTRIBUTION MAP"
        PIC_TRIM_02 = "FINAL TEST DISTRIBUTION MAP"
        PIC_TRIM_03 = "OK" '"�Ǖi"
        PIC_TRIM_04 = "NG" '"�s�Ǖi"
        PIC_TRIM_05 = "MIN %" '"�ŏ�%"
        PIC_TRIM_06 = "MAX %" '"�ő�%"
        PIC_TRIM_07 = "AVG %" '"����%"
        PIC_TRIM_08 = "Std Dev" '"�W���΍�"
        PIC_TRIM_09 = "Res Num" '"��R��"
        PIC_TRIM_10 = "DISTRIBUTION MAP SAVE"
    End Sub
#End Region


#Region "���O���(Main.txtLog)�ɕ������\������"
    '''=========================================================================
    '''<summary>���O���(frmMain.txtLog)�ɕ������\������</summary>
    '''<param name="s">(INP) �\��������</param>
    '''<remarks></remarks>
    '''=========================================================================
    Public Function Z_PRINT(Optional ByVal s As String = vbCrLf) As Integer

        Z_PRINT = LogPrint(s)
        Exit Function

    End Function
#End Region

#Region "���O��ʃN���A"
    '''=========================================================================
    '''<summary>���O��ʃN���A</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub Z_CLS()

        Try
            ''V2.2.0.0�P��
            If giTxtLogType <> 0 Then
                Static hWnd As IntPtr = ObjMain.txtlog.Handle
                Const WM_SETTEXT As Integer = &HC
                SendMessageString(hWnd, WM_SETTEXT, 0, "")                  ' �폜
            Else
                'ObjMain.txtLog.Text = ""
                ObjMain.lstLog.Items.Clear()
                ObjMain.lstLog.Items.Add(" ")
            End If
            ''V2.2.0.0�P��


        Catch ex As Exception

        End Try

    End Sub
#End Region

#Region "���O��ʕ\���T�u"
    '''=========================================================================
    '''<summary>���O��ʕ\���T�u</summary>
    '''<param name="s">(INP) �\��������</param>
    '''<remarks></remarks>
    '''=========================================================================
    Private Function LogPrint(ByVal s As String) As Integer

        Dim strMSG As String

        Try


            ''V2.2.0.0�P��
            '' �\���̍Ō�܂ŃX�N���[������
            LogPrint = 0                                ' Return�l = ����
            'ObjMain.txtLog.Text = ObjMain.txtLog.Text + s + "  "
            'ObjMain.txtLog.Focus()
            'ObjMain.txtLog.SelectionStart = ObjMain.txtLog.Text.Length
            'ObjMain.txtLog.ScrollToCaret()
            If giTxtLogType <> 0 Then
                ''V2.2.0.0�P
                Z_PRINT_MSG(s)
            Else
                With ObjMain.lstLog                                         ' ###lstLog
                    .BeginUpdate()
                    .Items.RemoveAt(.Items.Count - 1)
                    .Items.Add(s)
                    .Items.Add(" ")
                    .SelectedIndex = (.Items.Count - 1)
                    .ClearSelected()
                    .EndUpdate()
                End With
            End If
            ''V2.2.0.0�P��

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "LogPrint() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            LogPrint = -1                               ' Return�l = �G���[
        End Try

        Exit Function

    End Function


    ''' <summary>���O��ʂɕ������\������</summary>
    ''' <param name="s"></param>
    ''' <remarks>'#4.12.2.0�C</remarks>
    Public Function Z_PRINT_MSG(ByVal s As String) As Integer

        '#4.12.2.0�C                    ��
        'Static hWnd As IntPtr = ObjMain.lstLog.Handle
        Static hWnd As IntPtr = ObjMain.txtlog.Handle
        Const WM_GETTEXTLENGTH As Integer = &HE
        'Const LB_GETTEXT As Integer = &HE
        Const EM_SETSEL As Integer = &HB1
        Const EM_REPLACESEL As Integer = &HC2
        Const LB_ADDSTRING As Integer = &H180
        Const WM_COPYDATA As Integer = &H4A
        Dim result As Integer

        Try
            Dim test As String = Strings.Right(s, 1)
            Dim len As Integer = SendMessage(hWnd, WM_GETTEXTLENGTH, 0, 0)      ' �������擾
            SendMessage(hWnd, EM_SETSEL, len, len)                              ' �J�[�\���𖖔���
            If test <> "" AndAlso InStr(s, Environment.NewLine) = False Then
                s = s + Environment.NewLine
            End If
            SendMessageString(hWnd, EM_REPLACESEL, 0, s)  ' �e�L�X�g�ɕ������ǉ�����
            'SendMessageString(hWnd, EM_REPLACESEL, 0, s & Environment.NewLine)  ' �e�L�X�g�ɕ������ǉ�����
            ' SendMessageString(hWnd, LB_ADDSTRING, 0, s & Environment.NewLine)  ' �e�L�X�g�ɕ������ǉ�����

            'Dim cds As COPYDATASTRUCT

            'len = s.Length
            'cds.dwData = 0        '�g�p���Ȃ�
            'cds.lpData = s      '�e�L�X�g�̃|�C���^�[���Z�b�g
            'cds.cbData = len + 1     '�������Z�b�g
            ''������𑗂�
            'result = SendMessage(hWnd, WM_COPYDATA, 0, cds)

            Z_PRINT_MSG = cFRS_NORMAL


            ' �g���b�v�G���[������
        Catch ex As Exception
            Dim strMSG As String = "i-TKY.LogPrint() TRAP ERROR = " & ex.Message
            MsgBox(strMSG)
            'MessageBox.Show(Me, strMSG)
        End Try
    End Function

#End Region

    '=========================================================================
    '   ���[�_���o�͏���
    '=========================================================================
#Region "���[�_�o�̓T�u"
    '''=========================================================================
    '''<summary>���[�_�o�̓T�u</summary>
    '''<param name="LDON"> (INP) ON�r�b�g�f�[�^</param>
    '''<param name="LDOFF">(INP) OFF�r�b�g�f�[�^</param>
    '''<remarks></remarks>
    '''=========================================================================
    Public Function Sub_ATLDSET(ByVal LDON As Integer, ByVal LDOFF As Integer) As Integer

        Dim strMSG As String

        Try
            ' ���[�_�[�o��(ON,OFF)
            Sub_ATLDSET = Form1.System1.Z_ATLDSET(LDON, LDOFF)

            ' IO���j�^�\��
            gdwATLDDATA = gdwATLDDATA And (LDOFF Xor &HFFFF)
            gdwATLDDATA = gdwATLDDATA Or LDON
            Call IoMonitor(gdwATLDDATA, 1)

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "Sub_ATLDSET() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
        Exit Function
    End Function
#End Region

#Region "���[�_���̓T�u(�f�o�b�O�p)"
    '''=========================================================================
    '''<summary>���[�_���̓T�u(�f�o�b�O�p)</summary>
    '''<param name="Index"> (INP) ON�r�b�g�f�[�^</param>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub DEBUG_ReadHostCommand(ByVal Index As Integer)

        Dim strMSG As String

        Try
            Select Case Index
                Case 0
                    gDebugHostCmd = gDebugHostCmd Xor cHSTcRDY
                Case 1
                    gDebugHostCmd = gDebugHostCmd Xor cHSTcAUTO
                Case 2
                    gDebugHostCmd = gDebugHostCmd Xor cHSTcSTATE
                Case 3
                    gDebugHostCmd = gDebugHostCmd Xor cHSTcTRMCMD
                Case 4
                    gDebugHostCmd = gDebugHostCmd Xor &H10          ' Bit4:���g�p
                Case 5
                    gDebugHostCmd = gDebugHostCmd Xor &H20          ' Bit5:���g�p
                Case 6
                    gDebugHostCmd = 0
                Case 7
                    gDebugHostCmd = &HFFFF
            End Select

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "DEBUG_ReadHostCommand() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try

        Exit Sub
    End Sub
#End Region

#Region "IO���j�^�\��(���۰�ްI/O)"
    '''=========================================================================
    ''' <summary>IO���j�^�\��(���۰�ްI/O)</summary>
    ''' <param name="whcmd">(INP) I/O�ް�(16BIT)</param>
    ''' <param name="io">   (INP) �ް����(0=۰�ް����ϰ, 1=��ϰ��۰�ް)</param>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub IoMonitor(ByVal whcmd As Integer, ByVal io As Integer)

        Dim strMSG As String

        Try

#If cIOcMONITORcENABLED = 1 Then                    ' IO����\������

            Dim i As Integer


            If io = 0 Then
                For i = 0 To 15
                    If whcmd And (2 ^ i) Then
                        ObjMain.HostSignal(i).BackColor = Color.Red
                        ObjMain.HostSignal(i).Refresh()
                    Else
                        ObjMain.HostSignal(i).BackColor = Color.White
                        ObjMain.HostSignal(i).Refresh()
                    End If
                Next
            Else
                For i = 0 To 15
                    If whcmd And (2 ^ i) Then
                        ' BIT0�̓��쒆�^��~�������n�[�h�Ŕ��]�����Ă���B
                        If i = 0 Then
                            ObjMain.HostSignal(i + 16).BackColor = Color.White
                            ObjMain.HostSignal(i + 16).Refresh()
                        Else
                            ObjMain.HostSignal(i + 16).BackColor = Color.Lime
                            ObjMain.HostSignal(i + 16).Refresh()
                        End If
                    Else
                        If i = 0 Then
                            ObjMain.HostSignal(i + 16).BackColor = Color.Lime
                            ObjMain.HostSignal(i + 16).Refresh()
                        Else
                            ObjMain.HostSignal(i + 16).BackColor = Color.White
                            ObjMain.HostSignal(i + 16).Refresh()
                        End If
                    End If
                Next
            End If

#End If
            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "Globals.IoMonitor() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

    'V2.2.1.7�@ ��
#Region "�e�L�X�g�{�b�N�X�̕����񂪐����ϊ��ł��邩�m�F"
    ''' <summary>�e�L�X�g�{�b�N�X�̕����񂪐����ϊ��ł��邩�m�F�i�e�L�X�g�{�b�N�X�p�j</summary>
    ''' <param name="cTextBox">�m�F����÷���ޯ��</param>
    ''' <returns>(-1)=�װ</returns>
    Public Function CheckNumeric(ByRef cTextBox As cTxt_) As Integer
        Dim ret As Integer = 0
        Try

            '���l�`�F�b�N
            If IsNumeric(cTextBox.Text) Then
                'Nop
            Else
                MsgBox("���l����͂��Ă��������B")
                ret = -1
            End If
        Catch ex As Exception
            ret = -1
        Finally
            CheckNumeric = ret
        End Try

    End Function
#End Region
    'V2.2.1.7�@ ��

End Module