''===============================================================================
''   Description  : �O���[�o���萔�̒�`
''
''   Copyright(C) : OMRON LASERFRONT INC. 2010
''
'===============================================================================
Option Strict Off
Option Explicit On

Imports System.Runtime.InteropServices

Module Globals_Renamed
#Region "�O���[�o���萔/�ϐ��̒�`"
    '    '===========================================================================
    '    '   �O���[�o���萔/�ϐ��̒�`
    '    '===========================================================================
    '    '-------------------------------------------------------------------------------
    '    '   DLL��`
    '    '-------------------------------------------------------------------------------
    '    '----- WIN32 API -----
    '    ' �E�B���h�E�\���̑����API
    '    Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
    '    Public Const HWND_TOPMOST As Short = -1                 ' �E�B���h�E���őO�ʂɕ\��
    '    Public Const SWP_NOSIZE As Short = &H1S                 ' ���݂̃T�C�Y���ێ�
    '    Public Const SWP_NOMOVE As Short = &H2S                 ' ���݂̈ʒu���ێ�

    '    '---------------------------------------------------------------------------
    '    '   �A�v���P�[�V������/�A�v���P�[�V�������/�A�v���P�[�V�������[�h
    '    '---------------------------------------------------------------------------
    '    '----- �����I���p�A�v���P�[�V���� -----
    '    Public Const APP_FORCEEND As String = "c:\Trim\ForceEndProcess.exe"

    '    '----- �p�X�@-----
    '    Public Const OCX_PATH As String = "c:\Trim\ocx\"        '----- OCX�o�^�p�X
    '    Public Const DLL_PATH As String = "c:\Trim\"            '----- DLL�o�^�p�X

    '    '----- �A�v���P�[�V������ -----
    '    Public Const APP_TKY As String = "TKY"
    '    Public Const APP_CHIP As String = "TKYCHIP"
    '    Public Const APP_NET As String = "TKYNET"

    '    '----- �A�v���P�[�V������� -----
    '    Public Const KND_TKY As Short = 0
    '    Public Const KND_CHIP As Short = 1
    '    Public Const KND_NET As Short = 2
    '    Public Const MACHINE_TYPE_SL432 As String = "SL432R"                 ' �n��
    '    Public Const MACHINE_TYPE_SL436 As String = "SL436R"                 ' �n��

    '    Public gAppName As String                               ' �A�v���P�[�V������
    '    Public gTkyKnd As Short                                 ' �A�v���P�[�V�������

    '    '----- �摜�\���v���O�����̕\���ʒu -----
    '    'Public Const FORM_X As Integer = 4                     ' �R���g���[���㕔���[���WX ###050
    '    'Public Const FORM_Y As Integer = 20                    ' �R���g���[���㕔���[���WY ###050
    '    Public Const FORM_X As Integer = 0                      ' �R���g���[���㕔���[���WX ###050
    '    Public Const FORM_Y As Integer = 0                      ' �R���g���[���㕔���[���WY ###050

    '    '�T�u�t�H�[���̕\���ʒu�ڈ�̕\���ʒu�I�t�Z�b�g
    '    Public Const DISPOFF_SUBFORM_TOP As Integer = 12

    '    '----- �V�O�i���^���[������ -----                     ' ###007
    '    Public Const SIGTOWR_NORMAL As Short = 0                ' �W���R�F����
    '    Public Const SIGTOWR_SPCIAL As Short = 1                ' �S�F����(�������Ӱè�ޓa����)

    '    '----- �A�v���P�[�V�������[�h ----- (��)OcxSystem��`�ƈ�v������K�v�L��
    '    Public giAppMode As Short

    '    Public Const APP_MODE_IDLE As Short = 0                 ' �g���}���u�A�C�h����
    '    Public Const APP_MODE_LOAD As Short = 1                 ' �t�@�C�����[�h(F1)
    '    Public Const APP_MODE_SAVE As Short = 2                 ' �t�@�C���Z�[�u(F2)
    '    Public Const APP_MODE_EDIT As Short = 3                 ' �ҏW���      (F3)
    '    '                                                       ' ��
    '    Public Const APP_MODE_LASER As Short = 5                ' ���[�U�[����  (F5)
    '    Public Const APP_MODE_LOTCHG As Short = 6               ' ���b�g�ؑ�    (F6) �����[�U�v���Ή�
    '    Public Const APP_MODE_PROBE As Short = 7                ' �v���[�u      (F7)
    '    Public Const APP_MODE_TEACH As Short = 8                ' �e�B�[�`���O  (F8)
    '    Public Const APP_MODE_RECOG As Short = 9                ' �p�^�[���o�^  (F9)
    '    Public Const APP_MODE_EXIT As Short = 10                ' �I�� �@�@�@�@ (F11)
    '    Public Const APP_MODE_TRIM As Short = 11                ' �g���~���O��
    '    Public Const APP_MODE_CUTPOS As Short = 12              ' ��Ĉʒu�␳   (S-F8)
    '    Public Const APP_MODE_PROBE2 As Short = 13              ' �v���[�u2     (F10) �����[�U�v���Ή�
    '    Public Const APP_MODE_LOGGING As Short = 14             ' ���M���O      (F6) 

    '    ' CHIP,NET�n
    '    Public Const APP_MODE_TTHETA As Short = 40              ' �s��(�Ɗp�x�␳)�e�B�[�`���O
    Public Const APP_MODE_TX As Short = 41                  ' TX�e�B�[�`���O
    '    Public Const APP_MODE_TY As Short = 42                  ' TY�e�B�[�`���O
    '    Public Const APP_MODE_TY2 As Short = 43                 ' TY2�e�B�[�`���O
    '    Public Const APP_MODE_EXCAM_R1TEACH As Short = 44       ' �O���J����R1�e�B�[�`���O�y�O���J�����z
    '    Public Const APP_MODE_EXCAM_TEACH As Short = 45         ' �O���J�����e�B�[�`���O�y�O���J�����z
    Public Const APP_MODE_CARIB_REC As Short = 46           ' �摜�o�^(�L�����u���[�V�����␳�p)�y�O���J�����z
    '    Public Const APP_MODE_CARIB As Short = 47               ' �L�����u���[�V�����y�O���J�����z
    '    Public Const APP_MODE_CUTREVISE_REC As Short = 48       ' �摜�o�^(�J�b�g�ʒu�␳�p)�y�O���J�����z
    Public Const APP_MODE_CUTREVIDE As Short = 49           ' �J�b�g�ʒu�␳�y�O���J�����z
    '    Public Const APP_MODE_AUTO As Short = 50                ' �����^�]�@�@�@
    '    Public Const APP_MODE_LOADERINIT As Short = 51          ' ���[�_���_���A
    '    Public Const APP_MODE_LDR_ALRM As Short = 52            ' ���[�_�A���[�����    '###088
    Public Const APP_MODE_FINEADJ As Short = 53             ' �ꎞ��~���          '###088

    '    ' NET�n
    '    Public Const APP_MODE_CIRCUIT As Short = 60             ' �T�[�L�b�g�e�B�[�`���O

    '    '---------------------------------------------------------------------------
    '    '----- �@�\�I���`�e�[�u���̲��ޯ����` -----          '                         TKY CHIP NET
    '    '                                                       '                (��:�W��,��:��߼��,�~:����߰�)
    '    Public Const F_LOAD As Short = 0                        ' LOAD�{�^��              ��  ��   ��
    '    Public Const F_SAVE As Short = 1                        ' SAVE�{�^��              ��  ��   ��
    '    Public Const F_EDIT As Short = 2                        ' EDIT�{�^��              ��  ��   ��
    '    Public Const F_LASER As Short = 3                       ' LASER�{�^��             ��  ��   ��
    '    Public Const F_LOG As Short = 4                         ' LOGGING�{�^��           ��  ��   ��
    '    Public Const F_PROBE As Short = 5                       ' PROBE�{�^��             ��  ��   ��
    '    Public Const F_TEACH As Short = 6                       ' TEACH�{�^��             ��  ��   ��
    '    Public Const F_CUTPOS As Short = 7                      ' CUTPOS�{�^��            ��  �~   �~
    '    Public Const F_RECOG As Short = 8                       ' RECOG�{�^��             ��  ��   ��
    '    ' CHIP,NET�n
    '    Public Const F_TTHETA As Short = 9                      ' T�ƃ{�^��               �~  ��   ��
    '    Public Const F_TX As Short = 10                         ' TX�{�^��                �~  ��   ��
    '    Public Const F_TY As Short = 11                         ' TY�{�^��                �~  ��   ��
    '    Public Const F_TY2 As Short = 12                        ' TY2�{�^��               �~  ��   ��
    '    Public Const F_EXR1 As Short = 13                       ' �O�����R1è��ݸ�����    �~  ��   ��
    '    Public Const F_EXTEACH As Short = 14                    ' �O�����è��ݸ�����      �~  ��   ��
    '    Public Const F_CARREC As Short = 15                     ' �����ڰ��ݕ␳�o�^����  �~  ��   ��
    '    Public Const F_CAR As Short = 16                        ' �����ڰ�������          �~  ��   ��
    '    Public Const F_CUTREC As Short = 17                     ' ��ĕ␳�o�^����         �~  ��   ��
    '    Public Const F_CUTREV As Short = 18                     ' ��Ĉʒu�␳����         �~  ��   ��
    '    ' NET�n
    '    Public Const F_CIRCUIT As Short = 19                    ' �����è��ݸ�����        �~  �~   ��

    '    ' SL436R CHIP,NET�n 
    '    Public Const F_AUTO As Short = 20                       ' AUTO�{�^��              -   ��   ��
    '    Public Const F_LOADERINI As Short = 21                  ' LOADER INIT�{�^��       -   ��   ��

    '    Public Const MAX_FNCNO As Short = 22                    ' �@�\�I���`�e�[�u���̃f�[�^�� 

    '    '---------------------------------------------------------------------------

    '    '---------------------------------------------------------------------------
    '    '   �ő�l/�ŏ��l
    '    '---------------------------------------------------------------------------
    '    Public Const cMAXOptFlgNUM As Short = 5                 ' OcxSystem�p���߲ٵ�߼�݂̐� (�ő�5��)

    '    '----- �e���͍��ڂ͈̔� -----
    '    Public Const gMIN As Short = 0
    '    Public Const gMAX As Short = 1

    '    '----- ZZMOVE()�̈ړ��w�� -----
    '    Public Const MOVE_RELATIVE As Short = 0                 ' ���Βl�w�� 
    '    Public Const MOVE_ABSOLUTE As Short = 1                 ' ��Βl�w��

    '    '----- ZINPSTS()�̓��͉ӏ��w��  -----
    '    Public Const GET_CONSOLE_INPUT As Short = 1             ' �R���\�[��
    '    Public Const GET_INTERLOCK_INPUT As Short = 2           ' �C���^�[���b�N

    '    '----- �摜�o�^�p�p�����[�^ -----
    '    Public Const PTN_NUM_MAX As Short = 50                  ' �e���v���[�g�ԍ�(1-50)
    '    Public Const GRP_NUM_MAX As Short = 999                 ' ����ڰĸ�ٰ�ߔԍ�(1-999)

    '    Public Const INIT_THRESH_VAL As Double = 0.7            ' 臒l�����l
    '    Public Const INIT_CONTRAST_VAL As Integer = 216         ' �R���g���X�g�����l
    '    Public Const INIT_BRIGHTNESS_VAL As Integer = 0         ' �P�x�����l
    '    Public Const MIN_CONTRAST_VAL As Integer = 0            ' �R���g���X�g�ŏ��l
    '    Public Const MAX_CONTRAST_VAL As Integer = 511          ' �R���g���X�g�ő�l
    '    Public Const MIN_BRIGHTNESS_VAL As Integer = -128       ' �P�x�ŏ��l
    '    Public Const MAX_BRIGHTNESS_VAL As Integer = 127        ' �P�x�ő�l

    '    '----- ���[�_�p ----- 
    '    Public Const LALARM_COUNT As Integer = 128              ' �ő�A���[����
    '    Public Const MG_UP As Integer = 1                       ' �}�K�W���t�o      2013.01.28  '###182
    '    Public Const MG_DOWN As Integer = 0                     ' �}�K�W���c�n�v�m  2013.01.28  '###182


    '    '----- �}�[�L���O��R�ԍ� -----
    '    Public Const MARKING_RESNO_SET As Integer = 1000        ' ��R�ԍ�1000�Ԉȍ~�̓}�[�L���O�p�̒�R�ԍ�

    '    '---------------------------------------------------------------------------
    '    '   �V�X�e���p�����[�^(�`����DllSysprm.dll�Œ�`)
    '    '---------------------------------------------------------------------------
    '    Public gDllSysprmSysParam_definst As New DllSysprm.SysParam
    '    Public gSysPrm As DllSysprm.SYSPARAM_PARAM              ' �V�X�e���p�����[�^
    '    Public OptVideoPrm As DllSysprm.OPT_VIDEO_PRM           ' Video.ocx�p�I�v�V������`
    '    Public giTrimExe_NoWork As Short = 0                    ' �蓮���[�h���A�ڕ���Ɋ�Ȃ��Ńg���~���O���s����(0)/���Ȃ�(1)�@###240
    Public giTenKey_Btn As Short = 0                        ' �ꎞ��~��ʂł́uTen Key On/Off�v�{�^���̏����l(0:ON(����l), 1:OFF)�@###268
    Public giBpAdj_HALT As Short = 0                        ' �ꎞ��~��ʂł́uBP�I�t�Z�b�g��������/���Ȃ��v(0:��������(����l), 1:�������Ȃ�)�@###269

    '    '----- ONLINE -----
    '    Public Const TYPE_OFFLINE As Short = 0                  ' OFFLINE
    '    Public Const TYPE_ONLINE As Short = 1                   ' ONLINE
    '    Public Const TYPE_MANUAL As Short = 2                   ' SLIDE COVER+XY�ړ�����

    '    '----- ProbeType -----
    '    Public Const TYPE_PROBE_NON As Short = 0                ' NON
    '    Public Const TYPE_PROBE_STD As Short = 1                ' STANDARD

    '    '----- XY Table Exist Flag -----
    '    Public Const TYPE_XYTABLE_NON As Short = 0              ' NON
    '    Public Const TYPE_XYTABLE_X As Short = 1                ' X Only
    '    Public Const TYPE_XYTABLE_Y As Short = 2                ' Y Only
    '    Public Const TYPE_XYTABLE_XY As Short = 3               ' XY

    '    '----- �z����ײ���� -----
    '    Public Const VACCUME_ERRRETRY_OFF As Short = 0          ' Not retry
    '    Public Const VACCUME_ERRRETRY_ON As Short = 1           ' Retry
    '    Public Const RET_VACCUME_RETRY As Short = 1
    '    Public Const RET_VACCUME_CANCEL As Short = 2

    '    '----- ���ϲ�� -----
    '    Public Const customROHM As Short = 1                    ' ۰ѓa�����d�l
    '    Public Const customASAHI As Short = 2                   ' �����d�q�a�����d�l
    '    Public Const customSUSUMU As Short = 3                  ' �i�a�����d�l
    '    Public Const customKOA As Short = 4                     ' KOA(���̗�)�a�����d�l
    '    Public Const customKOAEW As Short = 5                   ' KOA(EW)�a�����d�l

    '    '----- �p���[���[�^�̃f�[�^�擾�擾 -----
    '    Public Const PM_DTTYPE_NONE As Short = 0                ' �Ȃ�
    '    Public Const PM_DTTYPE_IO As Short = 1                  ' �h�^�n�ǎ��
    '    Public Const PM_DTTYPE_USB As Short = 2                 ' �t�r�a

    '    '---------------------------------------------------------------------------
    '    '   �X�e�[�W����֌W
    '    '---------------------------------------------------------------------------
    '    ' �X�e�b�v����
    '    Public Const STEP_RPT_NON As Short = 0      ' �X�e�b�v�����s�[�g�����i�Ȃ��j
    '    Public Const STEP_RPT_X As Short = 1        ' �X�e�b�v�����s�[�g�����iX�����j
    '    Public Const STEP_RPT_Y As Short = 2        ' �X�e�b�v�����s�[�g�����iY�����j
    '    Public Const STEP_RPT_CHIPXSTPY As Short = 3 ' �X�e�b�v�����s�[�g�����iX�����`�b�v���X�e�b�v�{Y�����j
    '    Public Const STEP_RPT_CHIPYSTPX As Short = 4 ' �X�e�b�v�����s�[�g�����iY�����`�b�v���X�e�b�v�{X�����j

    '    ' BP�����
    '    Public Const BP_DIR_RIGHTUP As Short = 0    ' BP��E��i�v���X�����j�����@�@�@ 1 �Q �Q 0
    '    Public Const BP_DIR_LEFTUP As Short = 1     ' BP�����i�v���X�����j�����@�@�@�@|�Q|�Q|
    '    Public Const BP_DIR_RIGHTDOWN As Short = 2  ' BP��E���i�v���X�����j����        |�Q|�Q|
    '    Public Const BP_DIR_LEFTDOWN As Short = 3   ' BP������i�v���X�����j�����@�@�@ 3�@�@�@ 2

    '    Public Const BP_DIR_RIGHT As Short = 0      ' BP-X������E
    '    Public Const BP_DIR_LEFT As Short = 1       ' BP-X�������

    Public Const BLOCK_END As Short = 1         ' �u���b�N�I�� 
    Public Const PLATE_BLOCK_END As Short = 2   ' �v���[�g�E�u���b�N�I��

    '    '----- ���̑� -----
    '    ' FLSET�֐��̃��[�h
    '    Public Const FLMD_CNDSET As Integer = 0                 ' ���H�����ݒ�
    '    Public Const FLMD_BIAS_ON As Integer = 1                ' BIAS ON
    '    Public Const FLMD_BIAS_OFF As Integer = 2               ' BIAS OFF(LaserOff�֐�����BIAS OFF�͂��Ă���)


    '    '---------------------------------------------------------------------------
    '    '   ����t���O
    '    '---------------------------------------------------------------------------
    '    Public gCmpTrimDataFlg As Short                         ' �f�[�^�X�V�t���O(0=�X�V�Ȃ�, 1=�X�V����)
    '    Public giTrimErr As Short                               ' ��ϰ �װ �׸� ���װ���͸���߸����OFF����ϓ��쒆OFF��۰�ް�ɑ��M���Ȃ�
    '    '                                                       ' B0 : �z���װ(EXIT)
    '    '                                                       ' B1 : ���̑��װ
    '    '                                                       ' B2 : �W�o�@�װь��o
    '    '                                                       ' B3 : ���ЯĤ���װ�����ѱ��
    '    '                                                       ' B4 : ����~
    '    '                                                       ' B5 : ������װ

    '    Public gLoadDTFlag As Boolean                            ' �ް�۰�ލ��׸�(False:�ް���۰��, True:�ް�۰�ލ�)
    '    Public gbInitialized As Boolean                         ' True=���_���A��, False=���_���A��
    '    'Public bFgfrmDistribution As Boolean                    ' ���Y���̕\���׸�(TRUE:�\�� FALSE:��\��)
    '    Public gLoggingHeader As Boolean                        ' ۸�ͯ�ް�����ݎw���׸�(TRUE:�o��)
    '    Public gESLog_flg As Boolean                            ' ES���O�t���O(Flase=���OOFF, True=���OON)
    '    '' '' ''Public giAdjKeybord As Short                             ' �g���~���O��ADJ�@�\�L�[�{�[�h���(0:���͂Ȃ� 1:�� 2:�� 3:�E 4:�� )
    '    Public gPrevInterlockSw As Short

    '    Public gbCanceled As Boolean ' ���@�e��ʏ�����Private�Ŏ��� 

    '    '-------------------------------------------------------------------------------
    '    '   �I�u�W�F�N�g��`
    '    '-------------------------------------------------------------------------------
    '    '----- VB6��OCX -----
    '    'Public ObjSys As Object                                 ' OcxSystem.ocx
    '    'Public ObjUtl As Object                                 ' OcxUtility.ocx
    '    'Public ObjHlp As Object                                 ' OcxAbout.ocx
    '    'Public ObjPas As Object                                 ' OcxPassword.ocx
    '    'Public ObjMTC As Object                                 ' OcxManualTeach.ocx
    '    'Public ObjTch As Object                                 ' Teach.ocx
    '    'Public ObjPrb As Object                                 ' Probe.ocx
    '    'Public ObjVdo As Object                                 ' Video.ocx
    '    'Public ObjPrt As Object                                ' OcxPrint.ocx
    '    Public ObjMON(32) As Object
    '    Public gparModules As MainModules                                   ' �e�����\�b�h�ďo���I�u�W�F�N�g(OcxSystem�p) '###061
    '    Public ObjCrossLine As New TrimClassLibrary.TrimCrossLineClass()    ' �␳�N���X���C���\���p�I�u�W�F�N�g ###232 

    '    '---------------------------------------------------------------------------
    '    ' �g���~���O���샂�[�h
    '    '---------------------------------------------------------------------------
    '    Public Const TRIM_MODE_ITTRFT As Integer = 0    '�C�j�V�����e�X�g�{�g���~���O�{�t�@�C�i���e�X�g���s
    '    Public Const TRIM_MODE_TRFT As Integer = 1      '�g���~���O�{�t�@�C�i���e�X�g���s
    '    Public Const TRIM_MODE_FT As Integer = 2        '�t�@�C�i���e�X�g���s�i����j
    '    Public Const TRIM_MODE_MEAS As Integer = 3      '������s
    '    Public Const TRIM_MODE_POSCHK As Integer = 4    '�|�W�V�����`�F�b�N
    '    Public Const TRIM_MODE_CUT As Integer = 5       '�J�b�g���s
    '    Public Const TRIM_MODE_STPRPT As Integer = 6    '�X�e�b�v�����s�[�g���s
    '    Public Const TRIM_MODE_TRIMCUT As Integer = 7   '�g���~���O���[�h�ł̃J�b�g���s


    '    '-------------------------------------------------------------------------------
    '    ' �g���~���O����
    '    '-------------------------------------------------------------------------------
    '    '----- �g���~���O���ʒl�iINTRIM�Őݒ�j
    '    '//Trim result
    '    '//0:�����{   1:OK       2:ITNG      3:FTNG     4:SKIP
    '    '//5:RATIO    6:ITHI NG  7:ITLO NG   8:FTHI NG  9:FTLO NG
    '    '//10:        11:        12:         13:        14:
    '    '//15:�ٌ`�ʕt���ɂ��SKIP
    '    Public Const RSLT_NO_JUDGE As Integer = 0
    '    Public Const RSLT_OK As Integer = 1
    '    Public Const RSLT_IT_NG As Integer = 2
    '    Public Const RSLT_FT_NG As Integer = 3
    '    Public Const RSLT_SKIP As Integer = 4
    '    Public Const RSLT_RATIO As Integer = 5
    '    Public Const RSLT_IT_HING As Integer = 6
    '    Public Const RSLT_IT_LONG As Integer = 7
    '    Public Const RSLT_FT_HING As Integer = 8
    '    Public Const RSLT_FT_LONG As Integer = 9
    '    Public Const RSLT_RANGEOVER As Integer = 10
    '    Public Const RSLT_OPENCHK_NG As Integer = 20
    '    Public Const RSLT_SHORTCHK_NG As Integer = 21
    '    Public Const RSLT_IKEI_SKIP As Integer = 15

    '    '----- ���Y�Ǘ��O���t�t�H�[���I�u�W�F�N�g
    '    Public gObjFrmDistribute As Object                      ' frmDistribute

    '    '----- ���Y�Ǘ����p�z�� -----
    '    Public Const MAX_FRAM1_ARY As Integer = 15              ' ���x���z��
    '    '                                                       ' ���Y�Ǘ����̃��x���z��̃C���f�b�N�X 
    '    Public Const FRAM1_ARY_GO As Integer = 0                ' GO��(�T�[�L�b�g�� or ��R��)
    '    Public Const FRAM1_ARY_NG As Integer = 1                ' NG��(�T�[�L�b�g�� or ��R��)
    '    Public Const FRAM1_ARY_NGPER As Integer = 2             ' NG%
    '    Public Const FRAM1_ARY_PLTNUM As Integer = 3            ' PLATE��
    '    Public Const FRAM1_ARY_REGNUM As Integer = 4            ' RESISTOR��
    '    Public Const FRAM1_ARY_ITHING As Integer = 5            ' IT HI NG��
    '    Public Const FRAM1_ARY_FTHING As Integer = 6            ' FT HI NG��
    '    Public Const FRAM1_ARY_ITLONG As Integer = 7            ' IT LO NG��
    '    Public Const FRAM1_ARY_FTLONG As Integer = 8            ' FT LO NG��
    '    Public Const FRAM1_ARY_OVER As Integer = 9              ' OVER��
    '    Public Const FRAM1_ARY_ITHINGP As Integer = 10          ' IT HI NG%
    '    Public Const FRAM1_ARY_FTHINGP As Integer = 11          ' FT HI NG%
    '    Public Const FRAM1_ARY_ITLONGP As Integer = 12          ' IT LO NG%
    '    Public Const FRAM1_ARY_FTLONGP As Integer = 13          ' FT LO NG%
    '    Public Const FRAM1_ARY_OVERP As Integer = 14            ' OVER NG%

    '    Public Fram1LblAry(MAX_FRAM1_ARY) As System.Windows.Forms.Label     ' ���Y�Ǘ����̃��x���z��

    '    '-------------------------------------------------------------------------------
    '    '   gMode(OcxSystem��frmReset()�̏������[�h)
    '    '-------------------------------------------------------------------------------
    '    Public Const cGMODE_ORG As Short = 0                    '  0 : ���_���A
    '    Public Const cGMODE_ORG_MOVE As Short = 1               '  1 : ���_�ʒu�ړ�
    '    Public Const cGMODE_START_RESET As Short = 2            '  2 : ����m�F���(START/RESET�҂�)
    '    '                                                       '  3 :
    '    '                                                       '  4 :
    '    Public Const cGMODE_EMG As Short = 5                    '  5 : ����~���b�Z�[�W�\��
    '    '                                                       '  6 :
    '    Public Const cGMODE_SCVR_OPN As Short = 7               '  7 : �g���~���O���̃X���C�h�J�o�[�J���b�Z�[�W�\��
    '    Public Const cGMODE_CVR_OPN As Short = 8                '  8 : �g���~���O����➑̃J�o�[�J���b�Z�[�W�\��
    '    Public Const cGMODE_SCVRMSG As Short = 9                '  9 : �X���C�h�J�o�[�J���b�Z�[�W�\��(�g���~���O���ȊO)
    '    Public Const cGMODE_CVRMSG As Short = 10                ' 10 : ➑̃J�o�[�J�m�F���b�Z�[�W�\��(�g���~���O���ȊO)
    '    Public Const cGMODE_ERR_HW As Short = 11                ' 11 : �n�[�h�E�F�A�G���[(�J�o�[�����Ă܂�)���b�Z�[�W�\��
    '    Public Const cGMODE_ERR_HW2 As Short = 12               ' 12 : �n�[�h�E�F�A�G���[���b�Z�[�W�\��
    '    Public Const cGMODE_CVR_LATCH As Short = 13             ' 13 : �J�o�[�J���b�`���b�Z�[�W�\��
    '    Public Const cGMODE_CVR_CLOSEWAIT As Short = 14         ' 14 : ➑̃J�o�[�N���[�Y�������̓C���^�[���b�N�����҂�
    '    Public Const cGMODE_ERR_DUST As Short = 20              ' 20 : �W�o�@�ُ팟�o���b�Z�[�W�\��
    '    Public Const cGMODE_ERR_AIR As Short = 21               ' 21 : �G�A�[���G���[���o���b�Z�[�W�\��

    '    Public Const cGMODE_ERR_HING As Short = 40              ' 40 : �A��HI-NG�װ(ADV�������҂�)
    '    Public Const cGMODE_SWAP As Short = 41                  ' 41 : �����(START�������҂�)
    '    Public Const cGMODE_XYMOVE As Short = 42                ' 42 : �I������ð��وړ��m�F(START�������҂�)
    '    Public Const cGMODE_ERR_REPROBE As Short = 43           ' 43 : �ăv���[�r���O���s(START�������҂�) SL436R�p
    '    Public Const cGMODE_LDR_ALARM As Short = 44             ' 44 : ���[�_�A���[������   SL436R�p
    '    Public Const cGMODE_LDR_START As Short = 45             ' 45 : �����^�]�J�n(START�������҂�)   SL436R�p
    '    Public Const cGMODE_LDR_TMOUT As Short = 46             ' 46 : ���[�_�ʐM�^�C���A�E�g  SL436R�p
    '    Public Const cGMODE_LDR_END As Short = 47               ' 47 : �����^�]�I��(START�������҂�)   SL436R�p
    '    Public Const cGMODE_LDR_ORG As Short = 48               ' 48 : ���[�_���_���A  SL436R�p

    '    Public Const cGMODE_AUTO_LASER As Short = 50            ' 50 : �������[�U�p���[����

    '    Public Const cGMODE_LDR_CHK As Short = 60               ' 60 : ���[�_��ԃ`�F�b�N(�N����۰�ގ���Ӱ��/���쒆)
    '    Public Const cGMODE_LDR_ERR As Short = 61               ' 61 : ���[�_��ԃG���[(۰�ގ�����۰�ޖ�)
    '    Public Const cGMODE_LDR_MNL As Short = 62               ' 62 : �J�o�[�J��̃��[�_�蓮���[�h����
    '    Public Const cGMODE_LDR_WKREMOVE As Short = 63          ' 63 : �c���菜�����b�Z�[�W  SL436R�p
    '    Public Const cGMODE_LDR_RSTAUTO As Short = 64           ' 64 : �����^�]���~���b�Z�[�W  SL436R�p ###124
    '    Public Const cGMODE_LDR_WKREMOVE2 As Short = 65         ' 65 : �c���菜�����b�Z�[�W(APP�I��)  SL436R�p ###175
    '    Public Const cGMODE_LDR_STAGE_ORG As Short = 66         ' 66 : �X�e�[�W���_�ړ� SL436R�p ###188

    '    Public Const cGMODE_OPT_START As Short = 70             ' 70 : ���ݸފJ�n���̽���SW�����҂�
    '    Public Const cGMODE_OPT_END As Short = 71               ' 71 : ���ݸޏI�����̽ײ�޶�ް�J�҂�

    '    Public Const cGMODE_MSG_DSP As Short = 90               ' 90 : �w�胁�b�Z�[�W�\��(START�L�[�����҂�)

    '    ' ���~�b�g�Z���T�[& ���G���[ & �^�C���A�E�g���b�Z�[�W
    '    ' ��TrimErrNo.vb�Ɉړ�
    '    '                                                       ' ��(��)
    '    'Public Const cGMODE_TO_AXISX As Short = 101             ' 101: X���G���[(�^�C���A�E�g)
    '    'Public Const cGMODE_TO_AXISY As Short = 102             ' 102: Y���G���[(�^�C���A�E�g)
    '    'Public Const cGMODE_TO_AXISZ As Short = 103             ' 103: Z���G���[(�^�C���A�E�g)
    '    'Public Const cGMODE_TO_AXIST As Short = 104             ' 104: �Ǝ��G���[(�^�C���A�E�g)

    '    ''                                                       '�y�\�t�g���~�b�g�G���[�z
    '    'Public Const cGMODE_SL_AXISX As Short = 105             ' 105: X���\�t�g���~�b�g�G���[
    '    'Public Const cGMODE_SL_AXISY As Short = 106             ' 106: Y���\�t�g���~�b�g�G���[
    '    'Public Const cGMODE_SL_AXISZ As Short = 107             ' 107: Z���\�t�g���~�b�g�G���[
    '    'Public Const cGMODE_SL_BPX As Short = 110               ' 110: BP X���\�t�g���~�b�g�G���[
    '    'Public Const cGMODE_SL_BPY As Short = 111               ' 111: BP Y���\�t�g���~�b�g�G���[

    '    'Public Const cGMODE_TO_ROTATT As Short = 108            ' 108: ���[�^���A�b�e�l�[�^�G���[(�^�C���A�E�g)
    '    'Public Const cGMODE_TO_AXISZ2 As Short = 109            ' 109: Z2���G���[(�^�C���A�E�g)

    '    'Public Const cGMODE_SRV_ARM As Short = 202              ' 202: �T�[�{�A���[��
    '    'Public Const cGMODE_AXISX_LIM As Short = 203            ' 203: X�����~�b�g
    '    'Public Const cGMODE_AXISY_LIM As Short = 204            ' 204: Y�����~�b�g
    '    'Public Const cGMODE_AXISZ_LIM As Short = 205            ' 205: Z�����~�b�g
    '    'Public Const cGMODE_AXIST_LIM As Short = 206            ' 206: �Ǝ����~�b�g
    '    'Public Const cGMODE_RATT_LIM As Short = 207             ' 207: ���[�^���[�A�b�e�l�[�^���~�b�g
    '    'Public Const cGMODE_AXISZ2_LIM As Short = 208           ' 208: Z2�����~�b�g

    '    'Public Const cGMODE_BASE_ERR As Short = 200             ' Base Num.
    '    ''                                                       '�yX���G���[�z
    '    'Public Const cGMODE_AXISX_AOFF As Short = 211           ' 211: X���G���[(Bit All Off)
    '    'Public Const cGMODE_AXISX_AON As Short = 212            ' 212: X���G���[(Bit All On)
    '    'Public Const cGMODE_AXISX_ARM As Short = 213            ' 213: X���A���[��
    '    'Public Const cGMODE_AXISX_PML As Short = 214            ' 214: �}X�����~�b�g
    '    'Public Const cGMODE_AXISX_PLM As Short = 215            ' 215: +X�����~�b�g
    '    'Public Const cGMODE_AXISX_MLM As Short = 216            ' 216: -X�����~�b�g
    '    ''                                                       '�yY���G���[�z
    '    'Public Const cGMODE_AXISY_AOFF As Short = 221           ' 221: Y���G���[(Bit All Off)
    '    'Public Const cGMODE_AXISY_AON As Short = 222            ' 222: Y���G���[(Bit All On)
    '    'Public Const cGMODE_AXISY_ARM As Short = 223            ' 223: Y���A���[��
    '    'Public Const cGMODE_AXISY_PML As Short = 224            ' 224: �}Y�����~�b�g
    '    'Public Const cGMODE_AXISY_PLM As Short = 225            ' 225: +Y�����~�b�g
    '    'Public Const cGMODE_AXISY_MLM As Short = 226            ' 226: -Y�����~�b�g
    '    ''                                                       '�yZ���G���[�z
    '    'Public Const cGMODE_AXISZ_AOFF As Short = 231           ' 231: Z���G���[(Bit All Off)
    '    'Public Const cGMODE_AXISZ_AON As Short = 232            ' 232: Z���G���[(Bit All On)
    '    'Public Const cGMODE_AXISZ_ARM As Short = 233            ' 233: Z���A���[��
    '    'Public Const cGMODE_AXISZ_PML As Short = 234            ' 234: �}Z�����~�b�g
    '    'Public Const cGMODE_AXISZ_PLM As Short = 235            ' 235: +Z�����~�b�g
    '    'Public Const cGMODE_AXISZ_MLM As Short = 236            ' 236: -Z�����~�b�g
    '    'Public Const cGMODE_AXISZ_ORG As Short = 237            ' 237: Z�����_���A������
    '    ''                                                       '�y�Ǝ��G���[�z
    '    'Public Const cGMODE_AXIST_AOFF As Short = 241           ' 241: �Ǝ��G���[(Bit All Off)
    '    'Public Const cGMODE_AXIST_AON As Short = 242            ' 242: �Ǝ��G���[(Bit All On)
    '    'Public Const cGMODE_AXIST_ARM As Short = 243            ' 243: �Ǝ��A���[��
    '    'Public Const cGMODE_AXIST_PML As Short = 244            ' 244: �}�Ǝ����~�b�g
    '    'Public Const cGMODE_AXIST_PLM As Short = 245            ' 245: +�Ǝ����~�b�g
    '    'Public Const cGMODE_AXIST_MLM As Short = 246            ' 246: -�Ǝ����~�b�g
    '    ''                                                       '�yZ2���G���[�z
    '    'Public Const cGMODE_AXISZ2_AOFF As Short = 251          ' 251: Z2���G���[(Bit All Off)
    '    'Public Const cGMODE_AXISZ2_AON As Short = 252           ' 252: Z2���G���[(Bit All On)
    '    'Public Const cGMODE_AXISZ2_ARM As Short = 253           ' 253: Z2���A���[��
    '    'Public Const cGMODE_AXISZ2_PML As Short = 254           ' 254: �}Z2�����~�b�g
    '    'Public Const cGMODE_AXISZ2_PLM As Short = 255           ' 255: +Z2�����~�b�g
    '    'Public Const cGMODE_AXISZ2_MLM As Short = 256           ' 256: -Z2�����~�b�g
    '    'Public Const cGMODE_AXISZ2_ORG As Short = 257           ' 257: Z2�����_���A������
    '    ''                                                       '�y۰�ر��Ȱ���װ�z
    '    'Public Const cGMODE_ROTATT_AOFF As Short = 261          ' 261: ۰�ر��Ȱ���װ(Bit All Off)
    '    'Public Const cGMODE_ROTATT_AON As Short = 262           ' 262: ۰�ر��Ȱ���װ(Bit All On)
    '    'Public Const cGMODE_ROTATT_ARM As Short = 263           ' 263: ۰�ر��Ȱ���װ�
    '    'Public Const cGMODE_ROTATT_PML As Short = 264           ' 264: �}۰�ر��Ȱ���Я�
    '    'Public Const cGMODE_ROTATT_PLM As Short = 265           ' 265: +۰�ر��Ȱ���Я�
    '    'Public Const cGMODE_ROTATT_MLM As Short = 266           ' 266: -۰�ر��Ȱ���Я�

    '    ''-------------------------------------------------------------------------------
    '    ''   DllTrimFnc.dll�̖߂�l(��L�ȊO�̃p�����[�^�G���[��)
    '    ''-------------------------------------------------------------------------------
    '    'Public Const cFNC_ERR_TRIMRTN_ERR As Short = 99         ' �R�}���h���s�G���[(DllTrimFnc����99�ŕԂ��Ă������)
    '    'Public Const cFNC_ERR_CMD_NOTSPT As Short = 301         ' ���T�|�[�g�R�}���h
    '    'Public Const cFNC_ERR_CMD_PRM As Short = 302            ' �p�����[�^�G���[
    '    'Public Const cFNC_ERR_CMD_LIM_L As Short = 303          ' �p�����[�^�����l�G���[
    '    'Public Const cFNC_ERR_CMD_LIM_U As Short = 304          ' �p�����[�^����l�G���[
    '    'Public Const cFNC_ERR_RT2WIN_SEND As Short = 305        ' INTime��Windows���M�G���[
    '    'Public Const cFNC_ERR_RT2WIN_RECV As Short = 306        ' INTime��Windows��M�G���[
    '    'Public Const cFNC_ERR_WIN2RT_SEND As Short = 307        ' Windows��INTime���M�G���[
    '    'Public Const cFNC_ERR_WIN2RT_RECV As Short = 308        ' Windows��INTime��M�G���[

    '    '-------------------------------------------------------------------------------
    '    '   �߂�l(frmReset()��)
    '    '-------------------------------------------------------------------------------
    '    ' ��TrimErrNo.vb�Ɉړ�
    '    'Public Const cFRS_NORMAL As Short = 0                   ' ����
    '    'Public Const cFRS_ERR_ADV As Short = 1                  ' OK(ADV��)       �� START/RESET�҂���
    '    'Public Const cFRS_ERR_START As Short = 1                ' START(ADV��)    �� START/RESET�҂���
    '    'Public Const cFRS_ERR_HLT As Short = 2                  ' HALT��
    '    'Public Const cFRS_ERR_RST As Short = 3                  ' Cancel(RESET��) �� START/RESET�҂���
    '    'Public Const cFRS_ERR_Z As Short = 4                    ' Z��ON/OFF
    '    'Public Const cFRS_TxTy As Short = 5                     ' TX2/TY2����

    '    'Public Const cFRS_ERR_CVR As Short = -1                 ' ➑̃J�o�[�J���o
    '    'Public Const cFRS_ERR_SCVR As Short = -2                ' �X���C�h�J�o�[�J���o
    '    'Public Const cFRS_ERR_LATCH As Short = -3               ' �J�o�[�J���b�`���o

    '    'Public Const cFRS_ERR_EMG As Short = -11                ' ����~
    '    'Public Const cFRS_ERR_DUST As Short = -12               ' �W�o�@�ُ팟�o
    '    'Public Const cFRS_ERR_AIR As Short = -13                ' �G�A�[���G���[���o
    '    'Public Const cFRS_ERR_MVC As Short = -14                ' Ͻ������މ�H��ԃG���[���o
    '    'Public Const cFRS_ERR_HW As Short = -15                 ' �n�[�h�E�F�A�G���[���o

    '    ''----- IO����^�C���A�E�g -----
    '    'Public Const cFRS_TO_SCVR_CL As Short = -21             ' �^�C���A�E�g(�X���C�h�J�o�[�҂�)
    '    'Public Const cFRS_TO_SCVR_OP As Short = -22             ' �^�C���A�E�g(�X���C�h�J�o�[�J�҂�)
    '    'Public Const cFRS_TO_SCVR_ON As Short = -23             ' �^�C���A�E�g(�ײ�޶�ް�į�߰�s�҂�)
    '    'Public Const cFRS_TO_SCVR_OFF As Short = -24            ' �^�C���A�E�g(�ײ�޶�ް�į�߰�ߑ҂�)
    '    'Public Const cFRS_TO_CLAMP_ON As Short = -25            ' �^�C���A�E�g(�N�����v�n�m)
    '    'Public Const cFRS_TO_CLAMP_OFF As Short = -26           ' �^�C���A�E�g(�N�����v�n�e�e)
    '    'Public Const cFRS_TO_PM_DW As Short = -27               ' �^�C���A�E�g(�p���[���[�^���~�ړ�)
    '    'Public Const cFRS_TO_PM_UP As Short = -28               ' �^�C���A�E�g(�p���[���[�^�㏸�ړ�)
    '    'Public Const cFRS_TO_PM_FW As Short = -29               ' �^�C���A�E�g(�p���[���[�^����[�ړ�)
    '    'Public Const cFRS_TO_PM_BK As Short = -30               ' �^�C���A�E�g(�p���[���[�^�ҋ@�[�ړ�)

    '    ''----- ���G���[ & �^�C���A�E�g -----
    '    ''                                                       ' -101�`-266(���G���[ & �^�C���A�E�g)����L�Q��
    '    ''----- Main()�̖߂�l -----
    '    '' ��ʏ����p
    '    'Public Const cFRS_FNG_DATA As Short = -80               ' �f�[�^�����[�h
    '    'Public Const cFRS_FNG_CMD As Short = -81                ' ���R�}���h���s��
    '    'Public Const cFRS_FNG_PASS As Short = -82               ' �߽ܰ�ޓ��ʹװ

    '    '' �g���~���O�p
    '    'Public Const cFRS_TRIM_NG As Short = -90                ' �g���~���ONG
    '    'Public Const cFRS_ERR_TRIM As Short = -91               ' �g���}�G���[
    '    'Public Const cFRS_ERR_PTN As Short = -92                ' �p�^�[���F���G���[

    '    ''----- �p�����[�^�G���[��(���b�Z�[�W�\���͂��Ȃ�) -----
    '    'Public Const cFRS_ERR_CMD_NOTSPT As Short = -301        ' ���T�|�[�g�R�}���h
    '    'Public Const cFRS_ERR_CMD_PRM As Short = -302           ' �p�����[�^�G���[
    '    'Public Const cFRS_ERR_CMD_LIM_L As Short = -303         ' �p�����[�^�����l�G���[
    '    'Public Const cFRS_ERR_CMD_LIM_U As Short = -304         ' �p�����[�^����l�G���[
    '    'Public Const cFRS_ERR_CMD_OBJ As Short = -305           ' �I�u�W�F�N�g���ݒ�(Utility��޼ު�đ�)
    '    'Public Const cFRS_ERR_CMD_EXE As Short = -306           ' �R�}���h���s�G���[(DllTrimFnc����99�ŕԂ��Ă������)
    '    ''                                                       ' (��)cFRS_ERR_CMD_EXE�`cFRS_ERR_CMD_NOTSPT�Ŕ��肵�Ă���ӏ������邽��
    '    ' '' �@�@                                                   �ǉ�����ꍇ�͒���(cFRS_ERR_CMD_EXE�����炵�Ĕԍ���U�蒼��)
    '    ''----- Video.OCX�̃G���[ -----
    '    'Public Const cFRS_VIDEO_PTN As Short = -401             ' �p�^�[���F���G���[
    '    'Public Const cFRS_VIDEO_PT1 As Short = -402             ' �p�^�[���F���G���[(�␳�ʒu1)
    '    'Public Const cFRS_VIDEO_PT2 As Short = -403             ' �p�^�[���F���G���[(�␳�ʒu2)
    '    'Public Const cFRS_VIDEO_COM As Short = -404             ' �ʐM�G���[(CV3000)

    '    'Public Const cFRS_VIDEO_INI As Short = -411             ' ���������s���Ă��܂���
    '    'Public Const cFRS_VIDEO_IN2 As Short = -412             ' �������ς�
    '    'Public Const cFRS_VIDEO_FRM As Short = -413             ' �t�H�[���\����
    '    'Public Const cFRS_VIDEO_PRP As Short = -414             ' �v���p�e�B�l�s��
    '    'Public Const cFRS_VIDEO_GRP As Short = -415             ' ����ڰĸ�ٰ�ߔԍ��װ
    '    'Public Const cFRS_VIDEO_MXT As Short = -416             ' �e���v���[�g�� > MAX

    '    'Public Const cFRS_VIDEO_UXP As Short = -421             ' �\�����ʃG���[
    '    'Public Const cFRS_VIDEO_UX2 As Short = -422             ' �\�����ʃG���[2

    '    'Public Const cFRS_MVC_UTL As Short = -431               ' MvcUtil �G���[
    '    'Public Const cFRS_MVC_PT2 As Short = -432               ' MvcPt2 �G���[
    '    'Public Const cFRS_MVC_10 As Short = -433                ' Mvc10 �G���[

    '    ''----- �t�@�C�����o�̓G���[ -----
    '    'Public Const cFRS_FIOERR_INP As Short = -501            ' �t�@�C�����̓G���[
    '    'Public Const cFRS_FIOERR_OUT As Short = -502            ' �t�@�C���o�̓G���[

    '    'Public Const cERR_TRAP As Short = -999                  ' ��O�G���[

    '    '---------------------------------------------------------------------------
    '    '   �␳�N���X���C���\���p�p�����[�^
    '    '---------------------------------------------------------------------------
    '    Public gstCLC As CLC_PARAM                              ' �␳�N���X���C���\���p�p�����[�^

    '    '---------------------------------------------------------------------------
    '    '   �t�@�C���p�X�֌W
    '    '---------------------------------------------------------------------------
    '    Public gStrTrimFileName As String                       ' ���ݸ��ް�̧�ٖ�

    '    ''''    lib.bas�@�ł����g�p����Ă��Ȃ��B
    '    Public gsDataLogPath As String

    '    Public gbCutPosTeach As Boolean                         ' CutPosTeach(�\����:True, ��\��:False)

    '    '---------------------------------------------------------------------------
    '    '   �ϐ���`
    '    '---------------------------------------------------------------------------

    '    '----- �p�^�[���F���p -----
    '    Public giTempGrpNo As Integer                           ' �e���v���[�g�O���[�v�ԍ�(1�`999)
    '    Public giTempNo As Integer                              ' �e���v���[�g�ԍ�

    '    '----- �J�b�g�ʒu�␳�p�\���� -----
    '    Public Structure CutPosCorrect_Info                     ' �p�^�[���o�^���
    '        Dim intFLG As Short                                 ' �J�b�g�ʒu�␳�t���O(0:���Ȃ�, 1:����)
    '        Dim intGRP As Short                                 ' �p�^�[����ٰ�ߔԍ�(1-999)
    '        Dim intPTN As Short                                 ' �p�^�[���ԍ�(1-50)
    '        Dim dblPosX As Double                               ' �p�^�[���ʒuX(�␳�ʒu�e�B�[�`���O�p)
    '        Dim dblPosY As Double                               ' �p�^�[���ʒuY(�␳�ʒu�e�B�[�`���O�p)
    '        Dim intDisp As Short                                ' �p�^�[���F�����̌����g�\��(0:�Ȃ�, 1:����)
    '    End Structure

    '    Public Const MaxRegNum As Short = 256                   ' ��R���̍ő�l
    '    Public Const MaxCutNum As Short = 30                    ' �J�b�g�̍ő�l
    '    Public Const MaxDataNum As Short = 7681                 ' ��R��*�J�b�g�̍ő吔+1
    '    Public stCutPos(MaxRegNum + 1) As CutPosCorrect_Info        ' �p�^�[���o�^���

    '    Public giCutPosRNum As Short                            ' �J�b�g�ʒu�␳�����R��
    '    'Public giCutPosRSLT(MaxRegNum) As Short                 ' �p�^�[���F������(0:�␳�Ȃ�, 1:OK, 2:NG�����)
    '    'Public gfCutPosDRX(MaxRegNum) As Double                 ' �Y����X
    '    'Public gfCutPosDRY(MaxRegNum) As Double                 ' �Y����Y
    '    Public gfCutPosCoef(MaxRegNum) As Double                '  ��v�x

    '    '----- �ƕ␳�p -----
    '    Public gfCorrectPosX As Double                          ' �ƕ␳����XYð��ق����X(mm) ��ThetaCorrection()�Őݒ�
    '    Public gfCorrectPosY As Double                          ' �ƕ␳����XYð��ق����Y(mm)
    '    Public gbInPattern As Boolean                           ' �ʒu�␳������
    '    Public gbRotCorrectCancel As Short                      ' 0:OK, n < 0: �ʒu�␳���L�����Z������ or �ʒu�␳�G���[

    '    '----- �f�W�^���r�v -----
    '    'Public gDigH As Short                                   ' �f�W�^���r�v(Hight)
    '    'Public gDigL As Short                                   ' �f�W�^���r�v(Low)
    '    'Public gDigSW As Short                                  ' �f�W�^���r�v
    '    Public gPrevTrimMode As Short                           ' �f�W�^���r�v�l�ޔ���

    '    '----- GPIB�p -----
    '    Public giGpibDefAdder As Short = 21                     ' �����ݒ�(�@����ڽ)

    '    '----- ���̑� -----
    '    Public giIX2LOG As Short = 0                            ' IX2���O(0=����, 1=�L��)�@###231
    '    Public giTablePosUpd As Short = 0                       ' �e�[�u��1,2���W���X�V����/���Ȃ�(VIDEO.OCX�p�I�v�V����)�@###234

    '    ''''    ��������False�ɐݒ肵�Ă��邪�ATrue�ɐݒ肳��邱�Ƃ͂Ȃ��B
    '    ''''    �t���O�Ƃ��ċ@�\�͂��Ă��Ȃ��̂ŁA�R�[�h�m�F�̏�폜�B
    '    'Public OKFlag As Boolean                    'OK�{�^�������̗L��

    '    ''''    �������̂�
    '    'Public gRegisterExceptMarkingCnt As Short '��R���i�}�[�L���O��������) @@@007
    '    'Public gsSystemPassword As String
    '    'Public gLoggingEnd As Boolean

    '    ' '' '' ''----- ���Y�Ǘ���� -----
    '    '' '' ''Public glCircuitNgTotal As Integer                      ' �s�ǃT�[�L�b�g��
    '    '' '' ''Public glCircuitGoodTotal As Integer                    ' �Ǖi�T�[�L�b�g��
    '    '' '' ''Public glPlateCount As Integer                          ' �v���[�g������
    '    '' '' ''Public glGoodCount As Integer                           ' �Ǖi��R��
    '    '' '' ''Public glNgCount As Integer                             ' �s�ǒ�R��
    '    '' '' ''Public glITHINGCount As Integer                         ' IT HI NG��
    '    '' '' ''Public glITLONGCount As Integer                         ' IT LO NG��
    '    '' '' ''Public glFTHINGCount As Integer                         ' FT HI NG��
    '    '' '' ''Public glFTLONGCount As Integer                         ' FT LO NG��
    '    '' '' ''Public glITOVERCount As Integer                         ' IT���ް�ݼސ�


    '    Public gfPreviousPrbBpX As Double                       ' BP�_�����W��̈ʒuX (BSIZE+BPOFFSET����)
    '    Public gfPreviousPrbBpY As Double                       '                   Y

    '    ''''------------------------------------------------

    '    ''''---------------------------------------------------
    '    ''''�@090413 minato
    '    ''''    ProbeTeach�Őݒ肵�AResistorGraph�Ŏg�p���Ă���̂݁B
    '    ''''    �����ŏo����悤�Ɍ������B
    '    '---------------------------------------------------------------------------
    '    '   �S��R����̃O���t�\���p
    '    '---------------------------------------------------------------------------
    '    Public giMeasureResistors As Short                      ' ��R��
    '    Public giMeasureResiNum(512) As Double                  ' ��R�ԍ�
    '    Public gfMeasureResiOhm(512) As Double                  ' ���肵����R�l
    '    Public gfResistorTarget(512) As Double                  ' �ڕW�l
    '    Public gfMeasureResiPos(2, 512) As Double               ' �J�b�g�X�^�[�g�|�C���g
    '    Public giMeasureResiRst(512) As Short                   ' �g���~���O����

    '    Public Const cMEASUREcOK As Short = 1                   ' OK
    '    Public Const cMEASUREcIT As Short = 2                   ' IT ERROR
    '    Public Const cMEASUREcFT As Short = 3                   ' FT ERROR
    '    Public Const cMEASUREcNA As Short = 4                   ' ������


    '    '===============================================================================
    '    Public ExitFlag As Short
    '    Public gMode As Short '���[�h

    '    'INI�t�@�C���擾�f�[�^
    '    ''''(2010/11/16) ����m�F�㉺�L�R�����g�͍폜
    '    'Public gStartX As Double '�v���[�u�����lX
    '    'Public gStartY As Double '�v���[�u�����lY

    '    ' ���[�U�[����
    '    ''''    frmReset�ALASER_teaching�@�Ŏg�p
    '    Public gfLaserContXpos As Double
    '    Public gfLaserContYpos As Double

    '    '�摜�n���h��
    '    'Public mlHSKDib As Integer '����
    '    '�\���ʒu
    '    'Public mtDest As RECT
    '    'Public mtSrc As RECT
    '    'Public gVideoStarted As Boolean

    '    ''----- ����Ӱ�� ----- (��)OcxSystem��`�ƈ�v������K�v�L��
    '    'Public giAppMode As Short

    '    ''�f�[�^�ҏW�p�X���[�h�֘A
    '    'Public gbPassSucceeded As Boolean

    '    'Public gLoggingHeader As Boolean                    ' ͯ�ް�����ݎw���׸�(TRUE:�o��)
    '    'Public gbLogHeaderWrite As Boolean ' ���O�̃w�b�_�o�̓t���O @@@082

    '    'Public giOpLogFileHandle As Short ' ���샍�O�t�@�C���̃n���h��
    '    'Public gwTrimmerStatus As Short ' �z�X�g�ʐM�X�e�[�^�X�ێ�

    '    '''' ���M���O�t���O�@09/09/09  SysParam����ڍs


    '    Public Const KUGIRI_CHAR As Short = &H9S ' TAB

    '    'Public gbInPattern As Boolean ' �ʒu�␳������
    '    'Public gbRotCorrectCancel As Short ' 0:OK, n < 0: �ʒu�␳���L�����Z������ or �ʒu�␳�G���[
    '    ''Public gfCorrectPosX As Double                          ' �g�����|�W�V�����␳�lX 
    '    'Public gfCorrectPosY As Double                          ' �g�����|�W�V�����␳�lY
    '    'Public gbPreviousPrbPos As Boolean ' �v���[�u�ʒu���킹��BP/STAGE���W���L�����Ă���
    '    'Public gsCutTypeName(256) As String ' �J�b�g�^�C�v���e�[�u��
    '    'Public gtimerCoverTimeUp As Boolean

    '    ''BP���j�A���e�B�[�␳�l
    '    'Public Const cMAXcBPcLINEARITYcNUM As Short = 21


    '    ''''2009/05/29 minato
    '    ''''    LoaderAlarm.bas�폜�ɂ��ꎞ�ړ�
    '    ''''===============================================
    '    '' ''Public iLoaderAlarmKind As Short ' �װю��(1:�S��~�ُ� 2:���ْ�~ 3:�y�̏� 0:�װі���)
    '    '' ''Public iLoaderAlarmNum As Short ' �������̱װѐ�
    '    '' ''Public strLoaderAlarm() As String ' �װѕ�����
    '    '' ''Public strLoaderAlarmInfo() As String ' �װя��1
    '    '' ''Public strLoaderAlarmExec() As String ' �װя��2(�΍�)
    '    ''''===============================================



    '    'Public gbInitialized As Boolean

    '    '----- ���z�}�p -----
    '    Public Const MAX_SCALE_NUM As Integer = 999999999           ' ���̍ő�l
    '    Public Const MAX_SCALE_RNUM As Integer = 12                 ' ���̕\����R��

    '    Public gDistRegNumLblAry(12) As System.Windows.Forms.Label     ' ���z�O���t��R���z��
    '    Public gDistGrpPerLblAry(12) As System.Windows.Forms.Label     ' ���z�O���t%�z��
    '    Public gDistShpGrpLblAry(12) As System.Windows.Forms.Label     ' ���z�O���t�z��

    '    Public glRegistNum(12) As Integer                            ' ���z�O���t��R��
    '    Public glRegistNumIT(12) As Integer                          ' ���z�O���t��R�� �Ƽ��ý�
    '    Public glRegistNumFT(12) As Integer                          ' ���z�O���t��R�� ̧���ý�

    '    Public lOkChip As Integer                                   ' OK��
    '    Public lNgChip As Integer                                   ' NG��
    '    Public dblMinIT As Double                                   ' �ŏ��l�Ƽ��
    '    Public dblMaxIT As Double                                   ' �ő�l�Ƽ��
    '    Public dblMinFT As Double                                   ' �ŏ��ļ���
    '    Public dblMaxFT As Double                                   ' �ő�ļ���
    '    '' '' ''Public dblGapIT As Double                                   ' �ώZ�덷�Ƽ��
    '    '' '' ''Public dblGapFT As Double                                   ' �ώZ�덷̧���

    '    Public dblAverage As Double                                 ' ���ϒl
    '    Public dblDeviationIT As Double                             ' �W���΍�(IT)
    '    Public dblDeviationFT As Double                             ' �W���΍�(FT)

    '    Public dblAverageIT As Double                               ' IT���ϒl
    '    Public dblAverageFT As Double                               ' FT���ϒl
    '    Public HEIHOUIT As Double                                   ' �����΍�
    '    Public HEIHOUFT As Double                                   ' �����΍�

#End Region

#Region "�O���[�o���ϐ��̒�`"
    '    '===========================================================================
    '    '   �O���[�o���ϐ��̒�`
    '    '===========================================================================

    '    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '    ''''' 2009/04/13 minato
    '    ''''    TKY�ł͎g�p���Ă���O���[�o���ϐ�
    '    ''''    ���ʉ��ׁ̈ATKY�p�Ƃ��Ă͐錾���������

    '    '----- �A���^�]�p(SL436R�p) -----
    '    Public gbFgAutoOperation As Boolean = False                     ' �����^�]�t���O(True:�����^�]��, False:�����^�]���łȂ�) 
    '    Public gsAutoDataFileFullPath() As String                       ' �A���^�]�o�^�f�[�^�t�@�C�����z��
    '    Public giAutoDataFileNum As Short                               ' �A���^�]�o�^�f�[�^�t�@�C����
    '    Public giActMode As Short                                       ' �A���^�]���샂�[�h(0:϶޼��Ӱ�� 1:ۯ�Ӱ�� 2:����ڽӰ��)
    '    Public Const MODE_MAGAZINE As Short = 0                         ' �}�K�W�����[�h
    '    Public Const MODE_LOT As Short = 1                              ' ���b�g���[�h
    '    Public Const MODE_ENDLESS As Short = 2                          ' �G���h���X���[�h
    '    '                                                               ' �ؑւ����[�h(1=�������[�h, 0=�蓮���[�h)
    '    Public Const MODE_MANUAL As Integer = 0                         ' �蓮���[�h
    '    Public Const MODE_AUTO As Integer = 1                           ' �������[�h
    '    Public giErrLoader As Short = 0                                 ' ���[�_�A���[�����o(0:�����o 0�ȊO:�G���[�R�[�h) ###073

    '    '                                                               ' �ȉ��̓V�X�p�����ݒ肷��
    '    Public giOPLDTimeOutFlg As Integer                              ' ���[�_�ʐM�^�C���A�E�g���o(0=���o����, 1=���o����)
    '    Public giOPLDTimeOut As Integer                                 ' ���[�_�ʐM�^�C���A�E�g����(msec)
    '    Public giOPVacFlg As Integer                                    ' �蓮���[�h���̍ڕ���z���A���[�����o(0=���o����, 1=���o����)
    '    Public giOPVacTimeOut As Integer                                ' �蓮���[�h���̍ڕ���z���A���[���^�C���A�E�g����(msec)

    '    Public Const MAXWORK_KND As Integer = 10                        ' �v���[�g�f�[�^�̊�i��̐�
    '    Public giLoaderSpeed As Integer                                 ' ���[�_�������x
    '    Public giLoaderPositionSetting As Integer                       ' ���[�_�ʒu�ݒ�I��ԍ�
    '    Public gfBordTableOutPosX(0 To MAXWORK_KND - 1) As Double       ' ���[�_��e�[�u���r�o�ʒuX
    '    Public gfBordTableOutPosY(0 To MAXWORK_KND - 1) As Double       ' ���[�_��e�[�u���r�o�ʒuY
    '    Public gfBordTableInPosX(0 To MAXWORK_KND - 1) As Double        ' ���[�_��e�[�u�������ʒuX
    '    Public gfBordTableInPosY(0 To MAXWORK_KND - 1) As Double        ' ���[�_��e�[�u�������ʒuY
    '    Public giNgBoxCount(0 To MAXWORK_KND - 1) As Integer            ' NG�r�oBOX�̎��[����(��i�핪)   ###089
    '    Public giNgBoxCounter As Integer = 0                            ' NG�r�oBOX�̎��[�����J�E���^�[     ###089

    '    Public giBreakCounter As Integer = 0                            ' ���ꌇ�������̎��[�����J�E���^�[     ###130 
    '    Public giTwoTakeCounter As Integer = 0                          ' �Q����蔭���̎��[�����J�E���^�[     ###130 

    '    Public m_lTrimResult As Integer = cFRS_NORMAL                   ' ��P�ʂ̃g���~���O����(SL436R�����^�]����NG�r�oBOX�̎��[�����J�E���g�p) ###089
    '    '                                                               ' cFRS_NORMAL (����)
    '    '                                                               ' cFRS_TRIM_NG(�g���~���ONG)
    '    '                                                               ' cFRS_ERR_PTN(�p�^�[���F���G���[) ���Ȃ�
    Public bFgAutoMode As Boolean = False                           ' ���[�_�������[�h�t���O

    '    '----- �A���^�]�p(SL436R�p) -----


    '    '    Public Const cMAXcMARKINGcSTRLEN As Short = 18          ' �}�[�L���O������ő咷(byte)
    '    'Public strPlateDataFileFullPath() As String             ' �A���^�]�o�^ؽ����߽������z��
    '    'Public intPlateDataFileNum As Short                     ' �A���^�]�o�^ؽ����߽������
    '    'Public intActMode As Short                              ' �A���^�]����Ӱ��(0:϶޼��Ӱ�� 1:ۯ�Ӱ�� 2:����ڽӰ��)

    '    'Public INTRTM_Ver As String 'INtime Version
    '    'Public LMP_No As String 'LMP No


    '    '' '' ''Public gfX_2IT As Double ' IT�W���΍��Z�o�p���[�N
    '    '' '' ''Public gfX_2FT As Double ' FT�W���΍��Z�o�p���[�N

    '    Public glITTOTAL As Long                                        ' IT�v�Z�Ώې� ###138
    '    Public glFTTOTAL As Long                                        ' FT�v�Z�Ώې� ###138

    '    'Public gbEditPassword As Short ' �f�[�^���͎��̃p�X���[�h�v��(0:�� 1:�L)
    '    Public gITNx() As Double                                        'IT ����덷(�X)
    '    Public gFTNx() As Double                                        'FT ����덷(�X)

    '    Public gITNx_cnt As Integer                                     'IT �Z�o�pܰ���
    '    Public gITNg_cnt As Integer                                     'IT NG���L�^
    '    Public gFTNx_cnt As Integer                                     'FT �Z�o�pܰ���
    '    Public gFTNg_cnt As Integer                                     'FT NG���L�^
    '    'Public giXmode As Short
    '    Public gLogMode As Integer                                      '۷�ݸ�Ӱ��(0:���Ȃ�, 1:INITIAL TEST, 2:FINAL TEST, 3:INITIAL + FINAL) ###150 

    '    Public StepTab_Mode As Short                                    '(0)Step (1)Group
    '    Public StepFGMove As Short                                      '(0)�Ȃ��@(1)�ï�߸�د�ފԈړ�����[->]  (2)�ï�߸�د�ފԈړ�����[<-]
    '    Public StepTitle(2) As Short                                    '(0)���͂���@(1)���͂Ȃ�

    '    '--ROHM--
    '    Public giLoginPass As Boolean '�N�����߽ܰ�ޓ���(False)NG (True)OK
    '    'Public gsLoginPassword As String                    'ini̧�ٓ����߽ܰ��
    '    '--ROHM(���)--
    '    Public PrnDateR As String '��Ɠ�
    '    Public prnSTART_TIME As String '�J�n����
    '    Public prnSTOP_TIME As String '�I������
    '    Public prnPROG_TIME As String '�J�n�`�I���܂łɗv��������
    '    Public prnOPE_TIME As String '�ғ�����
    '    Public prnALARM_TIME As String '�װтɂ���~��������
    '    Public prnOPE_RATE As String '�ғ���
    '    Public prnMTBF As String '���ό̏�Ԋu
    '    Public prnMTTR As String '���ϕ�������
    '    Public prnLOT_NO As String '���ݸ��ް����ݽ���ް
    '    Public prnQrate As String '���ݸ�Qڰ�
    '    Public prnTrim_Speed As String '���ݸ޶�Ľ�߰��
    '    Public prnTrim_OK As Integer '�Ǖi���ߐ�
    '    Public prnPretest_Lo_Fail As Integer '�����l�����s�ǂ����ߐ�
    '    Public prnPretest_Hi_Fail As Integer '�����l����s�ǂ����ߐ�
    '    Public prnPretest_Open As Integer '�����l����ݕs�ǂ����ߐ�
    '    Public prnCut_NG As Integer '���ݸގ��ɖڕW�l�ɒB���Ȃ��������ߐ�
    '    Public prnPretest_NG_Cut_NG As Integer '�����s��
    '    Public prnFinal_test_Lo_Fail As Integer '���ݸތ�̉����s�ǂ����ߐ�
    '    Public prnFinal_test_Hi_Fail As Integer '���ݸތ�̏���s�ǂ����ߐ�
    '    Public prnFinal_test_Open As Integer '���ݸތ�ɵ���ݴװ�ƂȂ������ߐ�
    '    Public prnYield As String '�Ǖi���ߐ������ߐ�
    '    Public prnYield_Par As Double '��L��%�\��
    '    Public prnPdt_Sheet As Integer '���ݸ޽ð�ނŏ������������
    '    Public prnLot_Sheet As Integer '���u�ɓ������ꂽۯĖ���
    '    Public prnLot_NG_Sheet As Integer 'ۯĒ��̕s�Ǌ��
    '    Public prnEdg_Fail As Integer 'ۯĒ��̔F���s�Ǌ����
    '    Public prnNominal As Double '�ڕW��R�l
    '    Public prnTrim_Target As Double '�␳��̖ڕW��R�l
    '    Public prnTrim_Limit As Double '���ݸޖڕW�␳�l
    '    Public prnMean_Value As Double '���ݸނ��ꂽ���߂̕��ϒ�R�l
    '    Public prn_Par As Double '��L��%�\��
    '    Public prnM_R As Double '���ϒl�̌덷
    '    Public prn3S__x As Double '���ݸނ��ꂽ���߂̌덷�̕W���΍�

    '    Public prnSTtime As Double '�J�n����(double)
    '    Public prnEDtime As Double '�I������(double)
    '    Public prnAlmSTtime As Double '�װђ�~�J�n����(double)
    '    Public prnAlmEDtime As Double '�װђ�~�I������(double)
    '    Public prnAlmCnt As Short '�װє�����
    '    Public prnAlmTotaltime As Double '�װђ�~İ�َ���(double)
    '    Public prnChipTotal As Double '1ۯĕ��̑���R��
    '    Public prnTrim_NG As Integer '�s�Ǖi���ߐ�
    '    Public prnTrim_TotalVal As Double '���ݸނ��ꂽ���߂̒�R�l(���v)
    '    Public prnTrim_TotalValCnt As Double '���ݸނ��ꂽ���߂̒�R�l(�v�Z�p���v)
    '    Public prnTrim_TotalValKT As Short '���ݸނ��ꂽ���߂̒�R�l(��)


    '    'Public bPrnDataLoad As Boolean '�ް�۰��(True)����@(False)2��ڈȍ~

    '    Public sIX2LogFilePath As String 'IX2 LOĢ���߽��
    '    Public gsESLogFilePath As String 'ES LOĢ���߽��

    '    'frmFineAdjust.vb�ł̂ݎg�p����ϐ�
    '    '   �t�H�[���I����ɒl�̎擾���K�v�Ȃ��߁A
    '    '   �O���[�o���ŕϐ���ݒ肷��B
    Public gCurPlateNo As Integer
    Public gCurBlockNo As Integer
    '    Public gFrmEndStatus As Integer

    '    '----- ���O��ʕ\���p -----�@                                   '###013
    '    Public gDspClsCount As Integer                                  ' ���O��ʕ\���N���A�����
    '    Public gDspCounter As Integer                                   ' ���O��ʕ\��������J�E���^

    '    '----- �ꎞ��~��ʗp -----
    Public gbExitFlg As Boolean                                     '###014
    Public gbTenKeyFlg As Boolean = True                            ' �e���L�[���̓t���O ###057
    Public gbChkboxHalt As Boolean = True                           ' ADJ�{�^�����(ON=ADJ ON, OFF=ADJ OFF) ###009
    '    Public gbHaltSW As Boolean = False                              ' HALT SW��ԑޔ� ###255
    Public gObjADJ As Object = Nothing                              ' �ꎞ��~��ʃI�u�W�F�N�g ###053

    '    '----- EXTOUT LED����r�b�g -----                               '###061
    '    Public glLedBit As Long                                         ' LED����r�b�g(EXTOUT) 

    '    '----- GP-IB���� -----
    '    Public bGpib2Flg As Integer = 0                                 ' GP-IB����(�ėp)�t���O(0=����Ȃ�, 1=���䂠��) ###229

#End Region

    '========================================================================================
    '   �W���O����p�ϐ���`(�s�w/�s�x�e�B�[�`���O������)
    '========================================================================================
#Region "�W���O����p�ϐ���`"
    '-------------------------------------------------------------------------------
    '   �W���O����p��`
    '-------------------------------------------------------------------------------
    '    Public giCurrentNo As Integer                               ' �������̍s�ԍ�(�O���b�h�\���p)

    '    '----- JOG����p�p�����[�^�`����`(OcxJOG���g�p���Ȃ��ꍇ) -----
    Public Structure JOG_PARAM
        Dim Md As Short                                         ' �������[�h(0:XYð��وړ�, 1:BP�ړ�, 2:�L�[���͑҂����[�h)
        Dim Md2 As Short                                        ' ���̓��[�h(0:������ݓ���, 1:�ݿ�ٓ���)
        Dim Opt As UShort                                       ' �I�v�V����(�L�[�̗L��(1)/����(0)�w��)
        '                                                       '  BIT0:START�L�[
        '                                                       '  BIT1:RESET�L�[
        '                                                       '  BIT2:Z�L�[
        '                                                       '  BIT3:
        '                                                       '  BIT4:���g�p
        '                                                       '  BIT5:HALT�L�[
        '                                                       '  BIT6:���g�p
        '                                                       '  BIT7-15:���g�p
        Dim Flg As Short                                        ' �e��ʂ�OK/Cancel���݉����׸�(cFRS_ERR_ADV, cFRS_ERR_RST)
        Dim PosX As Double                                      ' BP or ð��� X�ʒu
        Dim PosY As Double                                      ' BP or ð��� Y�ʒu
        Dim BpOffX As Double                                    ' BP�̾��X 
        Dim BpOffY As Double                                    ' BP�̾��Y
        Dim BszX As Double                                      ' ��ۯ�����X 
        Dim BszY As Double                                      ' ��ۯ�����Y
        Dim TextX As Object                                     ' BP or ð��� X�ʒu�\���p÷���ޯ��
        Dim TextY As Object                                     ' BP or ð��� Y�ʒu�\���p÷���ޯ��
        Dim cgX As Double                                       ' �ړ���X 
        Dim cgY As Double                                       ' �ړ���Y
        Dim bZ As Boolean                                       ' Z�L�[  (True:ON, False:OFF)

        Dim BtnHI As Object                                     ' HI�{�^��
        Dim BtnZ As Object                                      ' Z�{�^��
        Dim BtnSTART As Object                                  ' START�{�^��
        Dim BtnHALT As Object                                   ' HALT�{�^��
        Dim BtnRESET As Object                                  ' RESET�{�^��
        Dim CurrentNo As Integer                                ' �������̍s�ԍ�(�O���b�h�\���p)
    End Structure

    '    '----- ZINPSTS�֐�(�R���\�[������)�ߒl -----
    Public Const CONSOLE_SW_START As UShort = &H1           ' bit 0(01)  : START       0/1=������/����
    Public Const CONSOLE_SW_RESET As UShort = &H2           ' bit 1(02)  : RESET       0/1=������/����
    Public Const CONSOLE_SW_ZSW As UShort = &H4             ' bit 2(04)  : Z_ON/OFF_SW 0/1=������/����
    '    Public Const CONSOLE_SW_ZDOWN As UShort = &H8           ' bit 3(08)  : Z_DOWN      1=��ԃZ���X
    '    Public Const CONSOLE_SW_ZUP As UShort = &H10            ' bit 4(10)  : Z_UP        1=��ԃZ���X
    Public Const CONSOLE_SW_HALT As UShort = &H20           ' bit 5(20)  : HALT        0/1=������/����

    '    '----- �R���\�[���L�[SW -----
    '    'Public Const cBIT_ADV As UShort = &H1US                 ' START(ADV)�L�[
    '    'Public Const cBIT_HALT As UShort = &H2US                ' HALT�L�[
    '    'Public Const cBIT_RESET As UShort = &H8US               ' RESET�L�[
    '    'Public Const cBIT_Z As UShort = &H20US                  ' Z�L�[
    Public Const cBIT_HI As UShort = &H100US                ' HI�L�[

    '    '----- �������[�h��` -----
    Public Const MODE_STG As Integer = 0                    ' XY�e�[�u�����[�h
    Public Const MODE_BP As Integer = 1                     ' BP���[�h
    Public Const MODE_KEY As Integer = 2                    ' �L�[���͑҂����[�h

    '    '----- �v���[�u���[�h/�T�u���[�h��` -----
    '    'Public Const MODE_STG      As Integer = 0              ' XY�e�[�u�����[�h
    '    'Public Const MODE_BP       As Integer = 1              ' BP���[�h
    '    Public Const MODE_Z As Integer = 2                      ' ZӰ��
    '    Public Const MODE_TTA As Integer = 3                    ' ��Ӱ��
    '    Public Const MODE_Z2 As Integer = 4                     ' Z2Ӱ��

    '    Public Const MODE_PRB As Integer = 10                   ' �ڐG�ʒu�m�F���[�h
    '    Public Const MODE_RECOG As Integer = 20                 ' �ƕ␳�蓮�ʒu�������[�h
    '    ' ���A�v�����[�h�́u�g���~���O���v
    '    Public Const MODE_POSOFS As Integer = 21                ' �␳�|�W�V�����I�t�Z�b�g�������[�h
    '    ' ���A�v�����[�h�́u�p�^�[���o�^(�ƕ␳)�v

    '    '----- ���̓��[�h -----
    Public Const MD2_BUTN As Integer = 0                    ' ��ʃ{�^������
    '    Public Const MD2_CONS As Integer = 1                    ' �R���\�[������
    '    Public Const MD2_BOTH As Integer = 2                    ' ����

    '    '----- �s�b�`�ő�l/�ŏ��l -----
    Public Const cPT_LO As Double = 0.001                   ' �߯��ŏ��l(mm)
    Public Const cPT_HI As Double = 0.1                     ' �߯��ő�l(mm)
    Public Const cHPT_LO As Double = 0.01                   ' HIGH�߯��ŏ��l(mm)
    Public Const cHPT_HI As Double = 5.0#                   ' HIGH�߯��ő�l(mm)
    Public Const cPAU_LO As Double = 0.05                   ' �|�[�Y�ŏ��l(sec)
    Public Const cPAU_HI As Double = 1.0#                   ' �|�[�Y�ő�l(sec)

    '    '----- �Y���� -----
    Public Const IDX_PIT As Short = 0                       ' �߯�
    Public Const IDX_HPT As Short = 1                       ' HIGH�߯�
    Public Const IDX_PAU As Short = 2                       ' �|�[�Y

    '    '----- ���̑� -----
    '    'Private dblTchMoval(3) As Double                           ' �߯��ޔ���(0=�߯�, 1=HIGH�߯�, 2=Pause Time(Sec))
    Private InpKey As UShort                                    ' �ݿ�ٷ����͈� 
    Private cin As UShort                                       ' �ݿ�ٓ��͒l
    Private bZ As Boolean                                       ' Z�L�[ �ޔ��� (True:ON, False:OFF)
    Private bHI As Boolean                                      ' HI�L�[(True:ON, False:OFF)

    Private mPIT As Double                                      ' �ړ��߯�
    Private X As Double                                         ' �ړ��߯�(X)
    Private Y As Double                                         ' �ړ��߯�(Y)
    '    Private NOWXP As Double                                     ' BP���ݒlX(�۽ײݕ␳�p)
    '    Private NOWYP As Double                                     ' BP���ݒlY(�۽ײݕ␳�p)
    Private mvx As Double                                       ' BP/ð��ٓ��̈ʒuX
    Private mvy As Double                                       ' BP/ð��ٓ��̈ʒuY
    Private mvxBk As Double                                     ' BP/ð��ٓ��̈ʒuX(�ޔ�p)
    Private mvyBk As Double                                     ' BP/ð��ٓ��̈ʒuY(�ޔ�p)
#End Region

    '    '========================================================================================
    '    '   �i�n�f�����ʏ����p���ʊ֐�
    '    '========================================================================================
#Region "�����ݒ菈��"
    '''=========================================================================
    '''<summary>�����ݒ菈��</summary>
    '''<param name="stJOG">       (INP)JOG����p�p�����[�^</param>
    '''<param name="TBarLowPitch">(I/O)�X���C�_�[1(Low�߯�)</param>
    '''<param name="TBarHiPitch"> (I/O)�X���C�_�[2(HIGH�߯�)</param>
    '''<param name="TBarPause">   (I/O)�X���C�_�[3(Pause Time)</param>
    '''<param name="LblTchMoval0">(I/O)�ڐ�1(Low Pich Label)</param>
    '''<param name="LblTchMoval1">(I/O)�ڐ�2(Low�߯� Label)</param>
    '''<param name="LblTchMoval2">(I/O)�ڐ�3(HIGH�߯� Label)</param>
    '''<param name="dblTchMoval"> (I/O)�߯��ޔ���(0=�߯�, 1=HIGH�߯�, 2=Pause Time)</param>
    '''=========================================================================
    Public Sub JogEzInit(ByVal stJOG As JOG_PARAM, _
                         ByRef TBarLowPitch As System.Windows.Forms.TrackBar, _
                         ByRef TBarHiPitch As System.Windows.Forms.TrackBar, _
                         ByRef TBarPause As System.Windows.Forms.TrackBar, _
                         ByRef LblTchMoval0 As System.Windows.Forms.Label, _
                         ByRef LblTchMoval1 As System.Windows.Forms.Label, _
                         ByRef LblTchMoval2 As System.Windows.Forms.Label, _
                         ByRef dblTchMoval() As Double)

        Dim strMSG As String

        Try
            ' �ړ��s�b�`�X���C�_�[�����ݒ�
            If (stJOG.Md = MODE_BP) Then                            ' ���[�h = 1(BP�ړ�) ?
                dblTchMoval(IDX_PIT) = gSysPrm.stSYP.gBpPIT         ' BP�p�߯��ݒ�
                dblTchMoval(IDX_HPT) = gSysPrm.stSYP.gBpHighPIT
                dblTchMoval(IDX_PAU) = gSysPrm.stSYP.gPitPause
            Else
                dblTchMoval(IDX_PIT) = gSysPrm.stSYP.gPIT           ' XY�e�[�u���p�߯��ݒ�
                dblTchMoval(IDX_HPT) = gSysPrm.stSYP.gStageHighPIT
                dblTchMoval(IDX_PAU) = gSysPrm.stSYP.gPitPause
            End If
            Call XyzBpMovingPitchInit(TBarLowPitch, TBarHiPitch, TBarPause, _
                                      LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval)

            Call Form1.System1.SetSysParam(gSysPrm)                 ' �V�X�e���p�����[�^�̐ݒ�(OcxSystem�p)

            InpKey = 0

            ' �g���b�v�G���[������
        Catch ex As Exception
            strMSG = "Globals.JogEzInit() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "BP/XY�e�[�u����JOG����(Do Loop�Ȃ�)"
    '''=========================================================================
    '''<summary>BP/XY�e�[�u����JOG���� ###047</summary>
    '''<param name="stJOG">       (INP)JOG����p�p�����[�^</param>
    '''<param name="TBarLowPitch">(I/O)�X���C�_�[1(Low�߯�)</param>
    '''<param name="TBarHiPitch"> (I/O)�X���C�_�[2(HIGH�߯�)</param>
    '''<param name="TBarPause">   (I/O)�X���C�_�[3(Pause Time)</param>
    '''<param name="LblTchMoval0">(I/O)�ڐ�1(Low Pich Label)</param>
    '''<param name="LblTchMoval1">(I/O)�ڐ�2(Low�߯� Label)</param>
    '''<param name="LblTchMoval2">(I/O)�ڐ�3(HIGH�߯� Label)</param>
    '''<param name="dblTchMoval"> (I/O)�߯��ޔ���(0=�߯�, 1=HIGH�߯�, 2=Pause Time)</param>
    '''<returns>cFRS_ERR_ADV = OK(START��) 
    '''         cFRS_ERR_RST = Cancel(RESET��)
    '''         cFRS_ERR_HLT = HALT��
    '''         -1�ȉ�       = �G���[</returns>
    ''' <remarks>JogEzInit�֐���Call�ςł��邱��</remarks>
    '''=========================================================================
    Public Function JogEzMove_Ex(ByRef stJOG As JOG_PARAM, ByVal SysPrm As DllSysprm.SYSPARAM_PARAM, _
                         ByRef TBarLowPitch As System.Windows.Forms.TrackBar, _
                         ByRef TBarHiPitch As System.Windows.Forms.TrackBar, _
                         ByRef TBarPause As System.Windows.Forms.TrackBar, _
                         ByRef LblTchMoval0 As System.Windows.Forms.Label, _
                         ByRef LblTchMoval1 As System.Windows.Forms.Label, _
                         ByRef LblTchMoval2 As System.Windows.Forms.Label, _
                         ByRef dblTchMoval() As Double) As Integer

        Dim strMSG As String
        Dim r As Short

        Try
            '---------------------------------------------------------------------------
            '   ��������
            '---------------------------------------------------------------------------
            X = 0.0 : Y = 0.0                                           ' �ړ��߯�X,Y
            mvx = stJOG.PosX : mvy = stJOG.PosY                         ' BP or ð��وʒuX,Y
            mvxBk = stJOG.PosX : mvyBk = stJOG.PosY
            ' �L�����u���[�V�������s/�J�b�g�ʒu�␳�y�O���J�����z�� �����΍��W��\�����邽�߃N���A���Ȃ�
            ' �g���~���O���̈ꎞ��~��ʂ��N���A���Ȃ�
            If (giAppMode <> APP_MODE_CARIB_REC) And (giAppMode <> APP_MODE_CUTREVIDE) And _
               (giAppMode <> APP_MODE_FINEADJ) Then                     '###088
                '(giAppMode <> APP_MODE_TRIM) Then                      '###088
                stJOG.cgX = 0.0# : stJOG.cgY = 0.0#                     ' �ړ���X,Y
            End If

            'If (giAppMode = APP_MODE_TRIM) Then                        '###088
            If (giAppMode = APP_MODE_FINEADJ) Then                      '###088
                mvx = stJOG.cgX - stJOG.BpOffX : mvy = stJOG.cgY - stJOG.BpOffY
                mvxBk = mvx : mvyBk = mvy
            End If

            Call Init_Proc(stJOG, TBarLowPitch, TBarHiPitch, TBarPause, LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval)

            ' ���݂̈ʒu��\������(÷���ޯ���̔w�i�F��������(���F)�ɐݒ肷��)
            Call DispPosition(stJOG, 1)
            'Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
            'Call Me.Focus()                                            ' �t�H�[�J�X��ݒ肷��(�e���L�[���͂̂���)
            ''                                                          ' KeyPreview�v���p�e�B��True�ɂ���ƑS�ẴL�[�C�x���g���܂��t�H�[�����󂯎��悤�ɂȂ�B
            '---------------------------------------------------------------------------
            '   �R���\�[���{�^�����̓R���\�[���L�[����̃L�[���͏������s��
            '---------------------------------------------------------------------------
            ' �V�X�e���G���[�`�F�b�N
            r = Form1.System1.SysErrChk_ForVBNET(giAppMode)
            If (r <> cFRS_NORMAL) Then Return (r)

            ' ���b�Z�[�W�|���v
            System.Windows.Forms.Application.DoEvents()

            '----- ###232�� -----
            '' �␳�N���X���C���\������(BP�ړ����[�h��Teach��)
            'If (stJOG.Md = MODE_BP) Then                                ' ���[�h = 1(BP�ړ�) ?
            '    NOWXP = 0.0 : NOWYP = 0.0
            '    If (gSysPrm.stCRL.giDspFlg = 1) Then                    ' �␳�N���X���C���\�� ?
            '        If (gSysPrm.stCRL.giDspFlg = 1) And _
            '           (giAppMode = APP_MODE_TEACH) Then                ' �␳�N���X���C���\�� ?
            '            Call ZGETBPPOS(NOWXP, NOWYP)                    ' BP���݈ʒu�擾
            '            gstCLC.x = NOWXP                                ' BP�ʒuX(mm)
            '            gstCLC.y = NOWYP                                ' BP�ʒuY(mm)
            '            Call CrossLineCorrect(gstCLC)                   ' �␳�N���X���C���\��
            '        End If
            '    End If
            'End If
            '----- ###232�� -----

            ' �R���\�[���{�^�����̓R���\�[���L�[����̃L�[����
            Call ReadConsoleSw(stJOG, cin)                              ' �L�[����

            '-----------------------------------------------------------------------
            '   ���̓L�[�`�F�b�N
            '-----------------------------------------------------------------------
            If (cin And CONSOLE_SW_RESET) Then                          ' RESET SW ?
                ' RESET SW������
                If (stJOG.Opt And CONSOLE_SW_RESET) Then                ' RESET�L�[�L�� ?
                    Return (cFRS_ERR_RST)                               ' Return�l = Cancel(RESET��)
                End If

                ' HALT SW������
            ElseIf (cin And CONSOLE_SW_HALT) Then                       ' HALT SW ?
                If (stJOG.Opt And CONSOLE_SW_HALT) Then                 ' �I�v�V����(0:HALT�L�[����, 1:HALT�L�[�L��)
                    r = cFRS_ERR_HALT                                   ' Return�l = HALT��
                    GoTo STP_END
                End If

                ' START SW������
            ElseIf (cin And CONSOLE_SW_START) Then                      ' START SW ?
                If (stJOG.Opt And CONSOLE_SW_START) Then                ' START�L�[�L�� ?
                    r = cFRS_ERR_START                                  ' Return�l = OK(START��) 
                    GoTo STP_END
                End If

                ' Z SW��ON����OFF(����OFF����ON)�ɐؑւ������
            ElseIf (stJOG.bZ <> bZ) Then
                If (stJOG.Opt And CONSOLE_SW_ZSW) Then                  ' Z�L�[�L�� ?
                    r = cFRS_ERR_Z                                      ' Return�l = Z��ON/OFF
                    stJOG.bZ = bZ                                       ' ON/OFF
                    GoTo STP_END
                End If

                ' ���SW������
            ElseIf cin And &H1E00US Then                                ' ���SW
                '�u�L�[���͑҂����[�h�v�Ȃ牽�����Ȃ�
                If (stJOG.Md = MODE_KEY) Then

                Else
                    If cin And &H100US Then                             ' HI SW ? 
                        mPIT = dblTchMoval(IDX_HPT)                     ' mPIT = �ړ������߯�
                    Else
                        mPIT = dblTchMoval(IDX_PIT)                     ' mPIT = �ړ��ʏ��߯�
                    End If

                    ' XY�e�[�u����Βl�ړ�(�\�t�g���~�b�g�`�F�b�N�L��)
                    r = cFRS_NORMAL
                    If (stJOG.Md = MODE_STG) Then                       ' ���[�h = XY�e�[�u���ړ� ?
                        ' XY�e�[�u����Βl�ړ�
                        r = Sub_XYtableMove(SysPrm, Form1.System1, Form1.Utility1, stJOG)
                        If (r <> cFRS_NORMAL) Then                      ' �װ ?
                            If (Form1.System1.IsSoftLimitXY(r) = False) Then
                                Return (r)                              ' ����ЯĴװ�ȊO�ʹװ����
                            End If
                        End If

                        '  ���[�h = BP�ړ��̏ꍇ
                    ElseIf (stJOG.Md = MODE_BP) Then
                        ' BP��Βl�ړ�
                        r = Sub_BPmove(SysPrm, Form1.System1, Form1.Utility1, stJOG)
                        If (r <> cFRS_NORMAL) Then                      ' BP�ړ��G���[ ?
                            If (Form1.System1.IsSoftLimitBP(r) = False) Then
                                Return (r)                              ' ����ЯĴװ�ȊO�ʹװ����
                            End If
                        End If
                    End If

                    ' �\�t�g���~�b�g�G���[�̏ꍇ�� HI SW�ȊO��OFF����
                    If (r <> cFRS_NORMAL) Then                          ' �װ ?
                        If (stJOG.BtnHI.BackColor = System.Drawing.Color.Yellow) Then
                            InpKey = cBIT_HI                            ' HI SW ON
                        Else
                            InpKey = 0                                  ' HI SW�ȊO��OFF
                        End If
                        r = cFRS_NORMAL                                 ' Retuen�l = ���� ###143 
                    End If

                    ' ���݂̈ʒu��\������
                    Call DispPosition(stJOG, 1)
                    'Call Form1.System1.WAIT(SysPrm.stSYP.gPitPause)    ' Wait(sec)'###251
                    Call Form1.System1.WAIT(dblTchMoval(IDX_PAU))       ' Wait(sec)'###251
                End If

            End If

            '---------------------------------------------------------------------------
            '   �I������
            '---------------------------------------------------------------------------
STP_END:
            'stJOG.PosX = mvx                                            ' �ʒuX,Y�X�V
            'stJOG.PosY = mvy
            Return (r)                                                  ' Return�l�ݒ� 

            ' �g���b�v�G���[������
        Catch ex As Exception
            strMSG = "Globals.JogEzMove_Ex() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Return�l = ��O�G���[ 
        End Try
    End Function
#End Region

    '#Region "BP/XY�e�[�u����JOG����"
    '    '''=========================================================================
    '    '''<summary>BP/XY�e�[�u����JOG����</summary>
    '    '''<param name="stJOG">       (INP)JOG����p�p�����[�^</param>
    '    '''<param name="TBarLowPitch">(I/O)�X���C�_�[1(Low�߯�)</param>
    '    '''<param name="TBarHiPitch"> (I/O)�X���C�_�[2(HIGH�߯�)</param>
    '    '''<param name="TBarPause">   (I/O)�X���C�_�[3(Pause Time)</param>
    '    '''<param name="LblTchMoval0">(I/O)�ڐ�1(Low Pich Label)</param>
    '    '''<param name="LblTchMoval1">(I/O)�ڐ�2(Low�߯� Label)</param>
    '    '''<param name="LblTchMoval2">(I/O)�ڐ�3(HIGH�߯� Label)</param>
    '    '''<param name="dblTchMoval"> (I/O)�߯��ޔ���(0=�߯�, 1=HIGH�߯�, 2=Pause Time)</param>
    '    '''<returns>cFRS_ERR_ADV = OK(START��) 
    '    '''         cFRS_ERR_RST = Cancel(RESET��)
    '    '''         cFRS_ERR_HLT = HALT��
    '    '''         -1�ȉ�       = �G���[</returns>
    '    ''' <remarks>JogEzInit�֐���Call�ςł��邱��</remarks>
    '    '''=========================================================================
    '    Public Function JogEzMove(ByRef stJOG As JOG_PARAM, ByVal SysPrm As DllSysprm.SYSPARAM_PARAM, _
    '                         ByRef TBarLowPitch As System.Windows.Forms.TrackBar, _
    '                         ByRef TBarHiPitch As System.Windows.Forms.TrackBar, _
    '                         ByRef TBarPause As System.Windows.Forms.TrackBar, _
    '                         ByRef LblTchMoval0 As System.Windows.Forms.Label, _
    '                         ByRef LblTchMoval1 As System.Windows.Forms.Label, _
    '                         ByRef LblTchMoval2 As System.Windows.Forms.Label, _
    '                         ByRef dblTchMoval() As Double) As Integer

    '        Dim strMSG As String
    '        Dim r As Short

    '        Try
    '            '---------------------------------------------------------------------------
    '            '   ��������
    '            '---------------------------------------------------------------------------
    '            X = 0.0 : Y = 0.0                                   ' �ړ��߯�X,Y
    '            mvx = stJOG.PosX : mvy = stJOG.PosY                 ' BP or ð��وʒuX,Y
    '            mvxBk = stJOG.PosX : mvyBk = stJOG.PosY
    '            ' �L�����u���[�V�������s/�J�b�g�ʒu�␳�y�O���J�����z�� �����΍��W��\�����邽�߃N���A���Ȃ�
    '            ' �g���~���O���̈ꎞ��~��ʂ��N���A���Ȃ�
    '            If (giAppMode <> APP_MODE_CARIB_REC) And (giAppMode <> APP_MODE_CUTREVIDE) And _
    '               (giAppMode <> APP_MODE_FINEADJ) Then             '###088
    '                '(giAppMode <> APP_MODE_TRIM) Then              '###088
    '                stJOG.cgX = 0.0# : stJOG.cgY = 0.0#             ' �ړ���X,Y
    '            End If
    '            stJOG.Flg = -1
    '            InpKey = 0
    '            Call Init_Proc(stJOG, TBarLowPitch, TBarHiPitch, TBarPause, LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval)

    '            ' ���݂̈ʒu��\������(÷���ޯ���̔w�i�F��������(���F)�ɐݒ肷��)
    '            Call DispPosition(stJOG, 1)
    '            'Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    '            'Call Me.Focus()                                     ' �t�H�[�J�X��ݒ肷��(�e���L�[���͂̂���)
    '            ''                                                   ' KeyPreview�v���p�e�B��True�ɂ���ƑS�ẴL�[�C�x���g���܂��t�H�[�����󂯎��悤�ɂȂ�B
    '            '---------------------------------------------------------------------------
    '            '   �R���\�[���{�^�����̓R���\�[���L�[����̃L�[���͏������s��
    '            '---------------------------------------------------------------------------
    '            Do
    '                ' �V�X�e���G���[�`�F�b�N
    '                r = Form1.System1.SysErrChk_ForVBNET(giAppMode)
    '                If (r <> cFRS_NORMAL) Then GoTo STP_END

    '                ' ���b�Z�[�W�|���v
    '                '  ��VB.NET�̓}���`�X���b�h�Ή��Ȃ̂ŁA�{���̓C�x���g�̊J���ȂǂłȂ��A
    '                '    �X���b�h�𐶐����ăR�[�f�B���O������̂��������B
    '                '    �X���b�h�łȂ��Ă��A�Œ�Ń^�C�}�[�𗘗p����B
    '                System.Windows.Forms.Application.DoEvents()
    '                System.Threading.Thread.Sleep(10)               ' CPU�g�p���������邽�߃X���[�v

    '                '----- ###232�� -----
    '                '' �␳�N���X���C���\������(BP�ړ����[�h��Teach��)
    '                'If (stJOG.Md = MODE_BP) Then                    ' ���[�h = 1(BP�ړ�) ?
    '                '    NOWXP = 0.0 : NOWYP = 0.0
    '                '    If (gSysPrm.stCRL.giDspFlg = 1) Then        ' �␳�N���X���C���\�� ?
    '                '        If (gSysPrm.stCRL.giDspFlg = 1) And _
    '                '           (giAppMode = APP_MODE_TEACH) Then    ' �␳�N���X���C���\�� ?
    '                '            Call ZGETBPPOS(NOWXP, NOWYP)        ' BP���݈ʒu�擾
    '                '            gstCLC.x = NOWXP                    ' BP�ʒuX(mm)
    '                '            gstCLC.y = NOWYP                    ' BP�ʒuY(mm)
    '                '            Call CrossLineCorrect(gstCLC)       ' �␳�N���X���C���\��
    '                '        End If
    '                '    End If
    '                'End If
    '                '----- ###232�� -----

    '                ' �R���\�[���{�^�����̓R���\�[���L�[����̃L�[����
    '                Call ReadConsoleSw(stJOG, cin)                  ' �L�[����

    '                '-----------------------------------------------------------------------
    '                '   ���̓L�[�`�F�b�N
    '                '-----------------------------------------------------------------------
    '                If (cin And CONSOLE_SW_RESET) Then              ' RESET SW ?
    '                    ' RESET SW������
    '                    If (stJOG.Opt And CONSOLE_SW_RESET) Then    ' RESET�L�[�L�� ?
    '                        r = cFRS_ERR_RST                        ' Return�l = Cancel(RESET��)
    '                        Exit Do
    '                    End If

    '                    ' HALT SW������
    '                ElseIf (cin And CONSOLE_SW_HALT) Then           ' HALT SW ?
    '                    If (stJOG.Opt And CONSOLE_SW_HALT) Then     ' �I�v�V����(0:HALT�L�[����, 1:HALT�L�[�L��)
    '                        r = cFRS_ERR_HALT                       ' Return�l = HALT��
    '                        Exit Do
    '                    End If

    '                    ' START SW������
    '                ElseIf (cin And CONSOLE_SW_START) Then          ' START SW ?
    '                    If (stJOG.Opt And CONSOLE_SW_START) Then    ' START�L�[�L�� ?
    '                        'stJOG.PosX = mvx                       ' �ʒuX,Y�X�V
    '                        'stJOG.PosY = mvy
    '                        r = cFRS_ERR_START                      ' Return�l = OK(START��) 
    '                        Exit Do
    '                    End If

    '                    ' Z SW��ON����OFF(����OFF����ON)�ɐؑւ������
    '                ElseIf (stJOG.bZ <> bZ) Then
    '                    If (stJOG.Opt And CONSOLE_SW_ZSW) Then      ' Z�L�[�L�� ?
    '                        r = cFRS_ERR_Z                          ' Return�l = Z��ON/OFF
    '                        stJOG.bZ = bZ                           ' ON/OFF
    '                        Exit Do
    '                    End If

    '                    ' ���SW������
    '                ElseIf cin And &H1E00US Then                    ' ���SW
    '                    '�u�L�[���͑҂����[�h�v�Ȃ牽�����Ȃ�
    '                    If (stJOG.Md = MODE_KEY) Then

    '                    Else
    '                        If cin And &H100US Then                     ' HI SW ? 
    '                            mPIT = dblTchMoval(IDX_HPT)             ' mPIT = �ړ������߯�
    '                        Else
    '                            mPIT = dblTchMoval(IDX_PIT)             ' mPIT = �ړ��ʏ��߯�
    '                        End If

    '                        ' XY�e�[�u����Βl�ړ�(�\�t�g���~�b�g�`�F�b�N�L��)
    '                        r = cFRS_NORMAL
    '                        If (stJOG.Md = MODE_STG) Then                ' ���[�h = XY�e�[�u���ړ� ?
    '                            ' XY�e�[�u����Βl�ړ�
    '                            r = Sub_XYtableMove(SysPrm, Form1.System1, Form1.Utility1, stJOG)
    '                            If (r <> cFRS_NORMAL) Then              ' �װ ?
    '                                If (Form1.System1.IsSoftLimitXY(r) = False) Then
    '                                    GoTo STP_END                    ' ����ЯĴװ�ȊO�ʹװ����
    '                                End If
    '                            End If

    '                            '  ���[�h = BP�ړ��̏ꍇ
    '                        ElseIf (stJOG.Md = MODE_BP) Then
    '                            ' BP��Βl�ړ�
    '                            r = Sub_BPmove(SysPrm, Form1.System1, Form1.Utility1, stJOG)
    '                            If (r <> cFRS_NORMAL) Then              ' BP�ړ��G���[ ?
    '                                If (Form1.System1.IsSoftLimitBP(r) = False) Then
    '                                    GoTo STP_END                    ' ����ЯĴװ�ȊO�ʹװ����
    '                                End If
    '                            End If
    '                        End If

    '                        ' �\�t�g���~�b�g�G���[�̏ꍇ�� HI SW�ȊO��OFF����
    '                        If (r <> cFRS_NORMAL) Then                  ' �װ ?
    '                            If (stJOG.BtnHI.BackColor = System.Drawing.Color.Yellow) Then
    '                                InpKey = cBIT_HI                    ' HI SW ON
    '                            Else
    '                                InpKey = 0                          ' HI SW�ȊO��OFF
    '                            End If
    '                        End If

    '                        ' ���݂̈ʒu��\������
    '                        Call DispPosition(stJOG, 1)
    '                        Call Form1.System1.WAIT(SysPrm.stSYP.gPitPause)    ' Wait(sec)
    '                    End If

    '                End If

    '            Loop While (stJOG.Flg = -1)

    '            '---------------------------------------------------------------------------
    '            '   �I������
    '            '---------------------------------------------------------------------------
    '            ' ���W�\���p÷���ޯ���̔w�i�F�𔒐F�ɐݒ肷��
    '            Call DispPosition(stJOG, 0)

    '            ' �e��ʂ���OK/Cancel���݉��� ?
    '            If (stJOG.Flg <> -1) Then
    '                r = stJOG.Flg
    '            End If

    '            ' OK(START��)�Ȃ�ʒuX,Y�X�V
    '            If (r = cFRS_ERR_START) Then                            ' OK(START��) ?
    '                stJOG.PosX = mvx                                    ' �ʒuX,Y�X�V
    '                stJOG.PosY = mvy
    '            End If

    'STP_END:
    '            Call ZCONRST()                                          ' �ݿ�ٷ�ׯ����� 
    '            Return (r)                                              ' Return�l�ݒ� 

    '            ' �g���b�v�G���[������
    '        Catch ex As Exception
    '            strMSG = "Globals.JogEzMove() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '            Return (cERR_TRAP)                                      ' Return�l = ��O�G���[ 
    '        End Try
    '    End Function
    '#End Region

#Region "�����ݒ菈��"
    '''=========================================================================
    '''<summary>�����ݒ菈��</summary>
    '''<param name="stJOG">       (INP)JOG����p�p�����[�^</param>
    '''<param name="TBarLowPitch">(I/O)�X���C�_�[1(Low�߯�)</param>
    '''<param name="TBarHiPitch"> (I/O)�X���C�_�[2(HIGH�߯�)</param>
    '''<param name="TBarPause">   (I/O)�X���C�_�[3(Pause Time)</param>
    '''<param name="LblTchMoval0">(I/O)�ڐ�1(Low Pich Label)</param>
    '''<param name="LblTchMoval1">(I/O)�ڐ�2(Low�߯� Label)</param>
    '''<param name="LblTchMoval2">(I/O)�ڐ�3(HIGH�߯� Label)</param>
    '''<param name="dblTchMoval"> (I/O)�߯��ޔ���(0=�߯�, 1=HIGH�߯�, 2=Pause Time)</param>
    '''=========================================================================
    Private Sub Init_Proc(ByVal stJOG As JOG_PARAM, _
                         ByRef TBarLowPitch As System.Windows.Forms.TrackBar, _
                         ByRef TBarHiPitch As System.Windows.Forms.TrackBar, _
                         ByRef TBarPause As System.Windows.Forms.TrackBar, _
                         ByRef LblTchMoval0 As System.Windows.Forms.Label, _
                         ByRef LblTchMoval1 As System.Windows.Forms.Label, _
                         ByRef LblTchMoval2 As System.Windows.Forms.Label, _
                         ByRef dblTchMoval() As Double)

        Dim strMSG As String

        Try

            ' �ړ��s�b�`�X���C�_�[�ݒ�(�O��ݒ肵���l)
            Call XyzBpMovingPitchInit(TBarLowPitch, TBarHiPitch, TBarPause, _
                                      LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval)

            ' �{�^���L��/�����ݒ�
            If (stJOG.Opt And CONSOLE_SW_HALT) Then                 ' HALT�L�[�L��/����
                stJOG.BtnHALT.Enabled = True
            Else
                stJOG.BtnHALT.Enabled = False
            End If
            If (stJOG.Opt And CONSOLE_SW_START) Then                ' START�L�[�L��/����
                stJOG.BtnSTART.Enabled = True
            Else
                stJOG.BtnSTART.Enabled = False
            End If
            If (stJOG.Opt And CONSOLE_SW_RESET) Then                ' RESET�L�[�L��/����
                stJOG.BtnRESET.Enabled = True
            Else
                stJOG.BtnRESET.Enabled = False
            End If
            If (stJOG.Opt And CONSOLE_SW_ZSW) Then                  ' Z�L�[�L��/����
                stJOG.BtnZ.Enabled = True
            Else
                stJOG.BtnZ.Enabled = False
            End If

            ' Z�L�[/HI�L�[��ԓ��ޔ�
            bZ = stJOG.bZ                                           ' Z�L�[�ޔ�
            If (bZ = False) Then                                    ' Z�{�^���̔w�i�F��ݒ�
                stJOG.BtnZ.BackColor = System.Drawing.SystemColors.Control ' �w�i�F = �D�F
                stJOG.BtnZ.Text = "Z Off"
            Else
                stJOG.BtnZ.BackColor = System.Drawing.Color.Yellow        ' �w�i�F = ���F
                stJOG.BtnZ.Text = "Z On"
            End If

            If (stJOG.BtnHI.BackColor = System.Drawing.Color.Yellow) Then ' HI�L�[��Ԏ擾
                bHI = True
                InpKey = InpKey Or cBIT_HI                          ' HI SW ON
            Else
                bHI = False
                InpKey = InpKey And Not cBIT_HI                     ' HI SW OFF
            End If

            ' �g���b�v�G���[������
        Catch ex As Exception
            strMSG = "Globals.Init_Proc() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "��ʃ{�^�����̓R���\�[���L�[����̃L�[����"
    '''=========================================================================
    '''<summary>��ʃ{�^�����̓R���\�[���L�[����̃L�[����</summary>
    '''<param name="stJOG">(INP)JOG����p�p�����[�^</param>
    '''<param name="cin">  (OUT)�R���\�[�����͒l</param>
    '''=========================================================================
    Private Sub ReadConsoleSw(ByRef stJOG As JOG_PARAM, ByRef cin As UShort)

        Dim r As Integer
        Dim sw As Long
        Dim strMSG As String

        Try
            ' HALT�L�[���̓`�F�b�N
            r = HALT_SWCHECK(sw)
            If (sw <> 0) Then                                           ' HALT�L�[���� ?
                If (stJOG.Opt And CONSOLE_SW_HALT) Then                 ' HALT�L�[�L�� ?
                    cin = CONSOLE_SW_HALT
                    Exit Sub
                End If
            End If

            ' Z�L�[���̓`�F�b�N
            r = Z_SWCHECK(sw)                                           ' Z�X�C�b�`�̏�Ԃ��`�F�b�N����
            If (sw <> 0) Then                                           ' Z�L�[���� ?
                If (stJOG.Opt And CONSOLE_SW_ZSW) Then                  ' Z�L�[�L�� ?
                    Call SubBtnZ_Click(stJOG)
                    Exit Sub
                End If
            End If

            ' START/RESET�L�[���̓`�F�b�N
            r = STARTRESET_SWCHECK(False, sw)                           ' START/RESET�L�[�����`�F�b�N(�Ď��Ȃ����[�h)

            ' �R���\�[�����͒l�ɕϊ����Đݒ�
            If (sw = cFRS_ERR_START) Then                               ' START�L�[���� ?
                If (stJOG.Opt And CONSOLE_SW_START) Then                ' START�L�[�L�� ?
                    cin = CONSOLE_SW_START
                    Exit Sub
                End If
            ElseIf (sw = cFRS_ERR_RST) Then                             ' RESET�L�[���� ?
                If (stJOG.Opt And CONSOLE_SW_RESET) Then                ' RESET�L�[�L�� ?
                    cin = CONSOLE_SW_RESET
                    Exit Sub
                End If
                '    ElseIf (sw = CONSOLE_SW_ZSW) Then                          ' Z�L�[���� ?
                '        If (stJOG.opt And CONSOLE_SW_ZSW) Then                  ' Z�L�[�L�� ?
                '            cin = CONSOLE_SW_ZSW
                '        End If
            End If

            ' �u��ʃ{�^�����́v
            cin = InpKey                                                ' ��ʃ{�^������

            ' �g���b�v�G���[������
        Catch ex As Exception
            strMSG = "Globals.ReadConsoleSw() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "���W�\��"
    '''=========================================================================
    '''<summary>���W�\��</summary>
    '''<param name="stJOG">(INP)JOG����p�p�����[�^</param>
    '''<param name="Md">   (INP)0=�w�i�F�𔒐F�ɐݒ�, 1=�w�i�F��������(���F)�ɐݒ�</param>
    '''=========================================================================
    Private Sub DispPosition(ByVal stJOG As JOG_PARAM, ByVal MD As Integer)

        Dim xPos As Double = 0.0                    ' ###232
        Dim yPos As Double = 0.0                    ' ###232
        Dim strMSG As String

        Try
            '�u�L�[���͑҂����[�h�v�Ȃ�NOP
            If (stJOG.Md = MODE_KEY) Then Exit Sub

            ' �␳�ʒu�e�B�[�`���O�Ȃ�O���b�h�ɕ\������
            If (giAppMode = APP_MODE_CUTPOS) Then
                stJOG.TextX.set_TextMatrix(stJOG.CurrentNo, 2, (stJOG.PosX + stJOG.cgX).ToString("0.0000"))
                stJOG.TextX.set_TextMatrix(stJOG.CurrentNo, 3, (stJOG.PosY + stJOG.cgY).ToString("0.0000"))
                Exit Sub

                ' �J�b�g�ʒu�␳�y�O���J�����z�Ȃ�O���b�h�ɑ��΍��W��\������
            ElseIf (giAppMode = APP_MODE_CUTREVIDE) Then
                stJOG.TextX.set_TextMatrix(stJOG.CurrentNo, 3, (stJOG.cgX).ToString("0.0000"))    ' �����X
                stJOG.TextX.set_TextMatrix(stJOG.CurrentNo, 4, (stJOG.cgY).ToString("0.0000"))    ' �����Y
                Exit Sub
            End If

            ' �e�L�X�g�{�b�N�X�ɍ��W��\������
            If (MD = 0) Then
                ' �L�����u���[�V�������s���͔w�i�F���D�F�ɐݒ�
                If (giAppMode = APP_MODE_CARIB_REC) Then
                    ' �w�i�F���D�F�ɐݒ�
                    stJOG.TextX.BackColor = System.Drawing.SystemColors.Control
                    stJOG.TextY.BackColor = System.Drawing.SystemColors.Control
                Else
                    ' �w�i�F�𔒐F�ɐݒ�
                    stJOG.TextX.BackColor = System.Drawing.Color.White
                    stJOG.TextY.BackColor = System.Drawing.Color.White
                End If
            Else
                ' �L�����u���[�V�������s���͑��΍��W��\��
                If (giAppMode = APP_MODE_CARIB_REC) Then
                    stJOG.TextX.Text = stJOG.cgX.ToString("0.0000")
                    stJOG.TextY.Text = stJOG.cgY.ToString("0.0000")
                Else
                    ' ���̑��̃��[�h���͐�΍��W��\��
                    stJOG.TextX.Text = (stJOG.PosX + stJOG.cgX).ToString("0.0000")
                    stJOG.TextY.Text = (stJOG.PosY + stJOG.cgY).ToString("0.0000")
                    '----- ###232�� -----
                    ' �g���~���O���̈ꎞ��~��ʕ\�����Ȃ�␳�N���X���C����\������
                    If (giAppMode = APP_MODE_FINEADJ) Or (giAppMode = APP_MODE_TX) Then
                        'xPos = Double.Parse(stJOG.TextX.Text)
                        'yPos = Double.Parse(stJOG.TextY.Text)
                        Call ZGETBPPOS(xPos, yPos)
                        ObjCrossLine.CrossLineDispXY(xPos, yPos)
                    End If
                    '----- ###232�� -----
                End If
                ' �w�i�F��������(���F)�ɐݒ�
                stJOG.TextX.BackColor = System.Drawing.Color.Yellow
                stJOG.TextY.BackColor = System.Drawing.Color.Yellow
            End If

            stJOG.TextX.Refresh()
            stJOG.TextY.Refresh()

            ' �g���b�v�G���[������
        Catch ex As Exception
            strMSG = "Globals.DispPosition() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

    '#Region "�e�B�[�`���O�r�v�擾"
    '    '''=========================================================================
    '    ''' <summary>�e�B�[�`���O�r�v�擾</summary>
    '    ''' <param name="SysPrm">(INP)�V�X�e���p�����[�^</param>
    '    ''' <param name="ObjSys">(INP)OcxSystem�I�u�W�F�N</param>
    '    ''' <returns>0=OFF, 1:ON</returns>
    '    '''=========================================================================
    '    Private Function Z_TEACHSTS(ByVal SysPrm As DllSysprm.SYSPARAM_PARAM, ByVal ObjSys As Object) As Long

    '        Dim r As Integer
    '        Dim strMSG As String

    '        Try
    '            ' �f�[�^���� & ON�r�b�g�`�F�b�N
    '            If (SysPrm.stIOC.giTeachSW = 1) Then                    ' �e�B�[�`���OSW���䂠�� ?
    '                r = ObjSys.Inp_And_Check_Bit(SysPrm.stIOC.glTS_In_Adr, SysPrm.stIOC.glTS_In_ON, SysPrm.stIOC.giTS_In_ON_ST)
    '                If (r = 1) Then                                     ' TEACH_SW ON ?
    '                    r = 1                                           ' TEACH_SW ON
    '                Else
    '                    r = 0                                           ' TEACH_SW OFF
    '                End If

    '            Else
    '                r = 1                                               ' TEACH_SW ON
    '            End If
    '            Return (r)

    '            ' �g���b�v�G���[������
    '        Catch ex As Exception
    '            strMSG = "Globals.Z_TEACHSTS() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '            Return (1)
    '        End Try
    '    End Function
    '#End Region

#Region "BP��Βl�ړ�(�\�t�g���~�b�g�`�F�b�N�L��)"
    '''=========================================================================
    ''' <summary>BP��Βl�ړ�(�\�t�g���~�b�g�`�F�b�N�L��)</summary>
    ''' <param name="SysPrm">(INP)�V�X�e���p�����[�^</param>
    ''' <param name="ObjSys">(INP)OcxSystem�I�u�W�F�N</param>
    ''' <param name="ObjUtl">(INP)OcxUtility�I�u�W�F�N</param>
    ''' <param name="stJOG"> (I/O)JOG����p�p�����[�^</param>
    ''' <returns>0=����, 0�ȊO:�G���[</returns>
    '''=========================================================================
    Private Function Sub_BPmove(ByVal SysPrm As DllSysprm.SYSPARAM_PARAM, ByVal ObjSys As Object, ByVal ObjUtl As Object, ByRef stJOG As JOG_PARAM) As Integer

        Dim r As Integer
        Dim strMSG As String

        Try
            ' BP�ړ��ʂ̎Z�o(��X,Y)
            mvxBk = mvx                                             ' ���݂̈ʒu�ޔ�
            mvyBk = mvy
            Call ObjUtl.GetBPmovePitch(cin, X, Y, mPIT, mvx, mvy, SysPrm.stDEV.giBpDirXy)

            ' BP��Βl�ړ�(�\�t�g���~�b�g�`�F�b�N�L��)
            r = ObjSys.BPMOVE(SysPrm, stJOG.BpOffX, stJOG.BpOffY, stJOG.BszX, stJOG.BszY, mvx, mvy, 1)
            If (r <> cFRS_NORMAL) Then                              ' �װ�Ȃ�װ����(���b�Z�[�W�\���ς�)
                If (ObjSys.IsSoftLimitBP(r) = False) Then
                    GoTo STP_END                                    ' ����ЯĴװ�ȊO�ʹװ����
                End If
                mvx = mvxBk                                         ' BP����ЯĴװ����BP�ʒu��߂�
                mvy = mvyBk
                GoTo STP_END                                        ' BP����ЯĴװ
            End If

            stJOG.cgX = stJOG.cgX + (-1 * X)                        ' BP�ړ���X�X�V (���ړ��ʂ͔��]���Ă���̂�-1���|����)
            stJOG.cgY = stJOG.cgY + (-1 * Y)                        ' BP�ړ���Y�X�V

STP_END:
            Return (r)

            ' �g���b�v�G���[������
        Catch ex As Exception
            strMSG = "Globals.Sub_BPmove() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)
        End Try
    End Function
#End Region

#Region "XY�e�[�u����Βl�ړ�(�\�t�g���~�b�g�`�F�b�N�L��)"
    '''=========================================================================
    ''' <summary>XY�e�[�u����Βl�ړ�(�\�t�g���~�b�g�`�F�b�N�L��)</summary>
    ''' <param name="SysPrm">(INP)�V�X�e���p�����[�^</param>
    ''' <param name="ObjSys">(INP)OcxSystem�I�u�W�F�N</param>
    ''' <param name="ObjUtl">(INP)OcxUtility�I�u�W�F�N</param>
    ''' <param name="stJOG"> (I/O)JOG����p�p�����[�^</param>
    ''' <returns>0=����, 0�ȊO:�G���[</returns>
    '''=========================================================================
    Private Function Sub_XYtableMove(ByVal SysPrm As DllSysprm.SYSPARAM_PARAM, ByVal ObjSys As Object, ByVal ObjUtl As Object, ByRef stJOG As JOG_PARAM) As Integer

        Dim r As Integer
        Dim strMSG As String

        Try
            ' XY�e�[�u���ړ��ʂ̎Z�o(��X,Y)
            mvxBk = X                                               ' ���݂̈ʒu�ޔ�
            mvyBk = Y
            Call ObjUtl.GetXYmovePitch(cin, X, Y, mPIT)

            ' XY�e�[�u����Βl�ړ�(�\�t�g���~�b�g�`�F�b�N�L��)
            r = ObjSys.XYtableMove(SysPrm, mvx + X, mvy + Y)
            If (r <> cFRS_NORMAL) Then                              ' �װ�Ȃ�װ����(���b�Z�[�W�\���ς�)
                If (ObjSys.IsSoftLimitXY(r) = False) Then
                    GoTo STP_END                                    ' ����ЯĴװ�ȊO�ʹװ����
                End If
                X = mvxBk                                           ' ����ЯĴװ����X,Y�ʒu��߂�
                Y = mvyBk
                GoTo STP_END                                        ' ����ЯĴװ
            End If

            mvx = mvx + X                                           ' �e�[�u���ʒuX,Y�X�V(��΍��W)
            mvy = mvy + Y
            stJOG.cgX = stJOG.cgX + X                               ' �e�[�u���ړ���X,Y�X�V
            stJOG.cgY = stJOG.cgY + Y

STP_END:
            Return (r)

            ' �g���b�v�G���[������
        Catch ex As Exception
            strMSG = "Globals.Sub_XYtableMove() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)
        End Try
    End Function
#End Region

    '    '========================================================================================
    '    '   �{�^������������(�i�n�f������)
    '    '========================================================================================
    '#Region "HALT�{�^������������"
    '    '''=========================================================================
    '    '''<summary>HALT�{�^������������</summary>
    '    '''=========================================================================
    '    Public Sub SubBtnHALT_Click()
    '        InpKey = CONSOLE_SW_HALT
    '    End Sub
    '#End Region

    '#Region "START�{�^������������"
    '    '''=========================================================================
    '    '''<summary>START�{�^������������</summary>
    '    '''=========================================================================
    '    Public Sub SubBtnSTART_Click()
    '        InpKey = CONSOLE_SW_START
    '    End Sub
    '#End Region

    '#Region "RESET�{�^������������"
    '    '''=========================================================================
    '    '''<summary>RESET�{�^������������</summary>
    '    '''=========================================================================
    '    Public Sub SubBtnRESET_Click()
    '        InpKey = CONSOLE_SW_RESET
    '    End Sub
    '#End Region

#Region "Z�{�^������������"
    '''=========================================================================
    '''<summary>RESET�{�^������������</summary>
    '''<param name="stJOG">(INP)JOG����p�p�����[�^</param>
    '''=========================================================================
    Public Sub SubBtnZ_Click(ByVal stJOG As JOG_PARAM)

        Dim strMSG As String

        Try
            If (stJOG.BtnZ.BackColor = System.Drawing.Color.Yellow) Then    ' Z SW ON ?
                stJOG.BtnZ.BackColor = System.Drawing.SystemColors.Control
                stJOG.BtnZ.Text = "Z Off"
                InpKey = InpKey And Not CONSOLE_SW_ZSW                      ' Z SW OFF
                bZ = False                                                  ' Z�L�[�ޔ���
            Else
                stJOG.BtnZ.BackColor = System.Drawing.Color.Yellow
                stJOG.BtnZ.Text = "Z On"
                InpKey = InpKey Or CONSOLE_SW_ZSW                           ' Z SW ON
                bZ = True                                                   ' Z�L�[�ޔ���
            End If

            ' �g���b�v�G���[������
        Catch ex As Exception
            strMSG = "Globals.SubBtnZ_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "HI�{�^������������"
    '''=========================================================================
    '''<summary>HI�{�^������������</summary>
    '''<param name="stJOG">(INP)JOG����p�p�����[�^</param>
    '''=========================================================================
    Public Sub SubBtnHI_Click(ByVal stJOG As JOG_PARAM)

        ' �w�i�F��ؑւ���
        If (stJOG.BtnHI.BackColor = System.Drawing.Color.Yellow) Then   ' �w�i�F = ���F ?
            ' �w�i�F���f�t�H���g�ɂ���
            stJOG.BtnHI.BackColor = System.Drawing.SystemColors.Control
            InpKey = InpKey And Not cBIT_HI                             ' HI SW OFF
        Else
            ' �w�i�F�����F�ɂ���
            stJOG.BtnHI.BackColor = System.Drawing.Color.Yellow
            InpKey = InpKey Or cBIT_HI                                  ' HI SW ON
        End If

    End Sub
#End Region

#Region "InpKey���擾����"
    '''=========================================================================
    '''<summary>InpKey���擾����</summary>
    '''<param name="IKey">(OUT)InpKey</param>
    '''=========================================================================
    Public Sub GetInpKey(ByRef IKey As UShort) '###057
        IKey = InpKey
    End Sub
#End Region

#Region "InpKey��ݒ肷��"
    '''=========================================================================
    '''<summary>InpKey��ݒ肷��</summary>
    '''<param name="IKey">(INP)InpKey</param>
    '''=========================================================================
    Public Sub PutInpKey(ByVal IKey As UShort) '###057
        InpKey = IKey
    End Sub
#End Region

#Region "���{�^��������"
    '''=========================================================================
    '''<summary>���{�^��������</summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub SubBtnJOG_0_MouseDown()
        InpKey = InpKey Or &H1000US                         ' +Y ON
    End Sub
    Public Sub SubBtnJOG_0_MouseUp()
        InpKey = InpKey And Not &H1000US                    ' +Y OFF
    End Sub

    Public Sub SubBtnJOG_1_MouseDown()
        InpKey = InpKey Or &H800US                          ' -Y ON
    End Sub
    Public Sub SubBtnJOG_1_MouseUp()
        InpKey = InpKey And Not &H800US                     ' -Y OFF
    End Sub

    Public Sub SubBtnJOG_2_MouseDown()
        InpKey = InpKey Or &H400US                          ' +X ON
    End Sub
    Public Sub SubBtnJOG_2_MouseUp()
        InpKey = InpKey And Not &H400US                     ' +X OFF
    End Sub

    Public Sub SubBtnJOG_3_MouseDown()
        InpKey = InpKey Or &H200US                          ' -X ON
    End Sub
    Public Sub SubBtnJOG_3_MouseUp()
        InpKey = InpKey And Not &H200US                     ' -X OFF
    End Sub

    Public Sub SubBtnJOG_4_MouseDown()
        InpKey = InpKey Or &HA00US                          ' -X -Y ON
    End Sub
    Public Sub SubBtnJOG_4_MouseUp()
        InpKey = InpKey And Not &HA00US                     ' -X -Y OFF
    End Sub

    Public Sub SubBtnJOG_5_MouseDown()
        InpKey = InpKey Or &HC00US                          ' +X -Y ON
    End Sub
    Public Sub SubBtnJOG_5_MouseUp()
        InpKey = InpKey And Not &HC00US                     ' +X -Y OFF
    End Sub

    Public Sub SubBtnJOG_6_MouseDown()
        InpKey = InpKey Or &H1400US                         ' +X +Y ON
    End Sub
    Public Sub SubBtnJOG_6_MouseUp()
        InpKey = InpKey And Not &H1400US                    ' +X +Y OFF
    End Sub

    Public Sub SubBtnJOG_7_MouseDown()
        InpKey = InpKey Or &H1200US                         ' -X +Y ON
    End Sub
    Public Sub SubBtnJOG_7_MouseUp()
        InpKey = InpKey And Not &H1200US                    ' -X +Y OFF
    End Sub
#End Region

    '========================================================================================
    '   �i�n�f�����ʏ����p�g���b�N�o�[����
    '========================================================================================
#Region "�g���b�N�o�[�̃X���C�_�[��ʏ����l�\��"
    '''=========================================================================
    '''<summary>�g���b�N�o�[�̃X���C�_�[��ʏ����l�\��</summary>
    '''<param name="TBarLowPitch">(I/O)�X���C�_�[1(Low�߯�)</param>
    '''<param name="TBarHiPitch"> (I/O)�X���C�_�[2(HIGH�߯�)</param>
    '''<param name="TBarPause">   (I/O)�X���C�_�[3(Pause Time)</param>
    '''<param name="LblTchMoval0">(I/O)�ڐ�1(Low Pich Label)</param>
    '''<param name="LblTchMoval1">(I/O)�ڐ�2(Low�߯� Label)</param>
    '''<param name="LblTchMoval2">(I/O)�ڐ�3(HIGH�߯� Label)</param>
    '''<param name="dblTchMoval"> (I/O)�߯��ޔ���(0=�߯�, 1=HIGH�߯�, 2=Pause Time)</param>
    '''=========================================================================
    Public Sub XyzBpMovingPitchInit(ByRef TBarLowPitch As System.Windows.Forms.TrackBar, _
                                    ByRef TBarHiPitch As System.Windows.Forms.TrackBar, _
                                    ByRef TBarPause As System.Windows.Forms.TrackBar, _
                                    ByRef LblTchMoval0 As System.Windows.Forms.Label, _
                                    ByRef LblTchMoval1 As System.Windows.Forms.Label, _
                                    ByRef LblTchMoval2 As System.Windows.Forms.Label, _
                                    ByRef dblTchMoval() As Double)

        Dim minval As Short

        ' LOW�߯������͈͊O�Ȃ�͈͓��ɕύX����
        If (dblTchMoval(IDX_PIT) < cPT_LO) Then dblTchMoval(IDX_PIT) = cPT_LO
        If (dblTchMoval(IDX_PIT) > cPT_HI) Then dblTchMoval(IDX_PIT) = cPT_HI

        ' LOW�߯��̖ڐ���ݒ肷��
        If (dblTchMoval(IDX_PIT) < 0.002) Then                          ' ����\�ɂ��ŏ��ڐ���ݒ肷��
            minval = 1                                                  ' �ڐ�1�`
        Else
            minval = 2                                                  ' �ڐ�2�`
        End If

        TBarLowPitch.TickFrequency = dblTchMoval(IDX_PIT) * 1000        ' 0.001mm�P��
        TBarLowPitch.Maximum = 100                                      ' �ڐ�1(or 2)�`100(0.001m�`0.1mm)
        TBarLowPitch.Minimum = minval
        '###110
        TBarLowPitch.Value = dblTchMoval(IDX_PIT) * 1000        ' 0.001mm�P��

        ' HIGH�߯������͈͊O�Ȃ�͈͓��ɕύX����
        If (dblTchMoval(IDX_HPT) < cHPT_LO) Then dblTchMoval(IDX_HPT) = cHPT_LO
        If (dblTchMoval(IDX_HPT) > cHPT_HI) Then dblTchMoval(IDX_HPT) = cHPT_HI

        ' HIGH�߯��̖ڐ���ݒ肷��
        TBarHiPitch.TickFrequency = dblTchMoval(IDX_HPT) * 100          ' 0.01mm�P��
        TBarHiPitch.Maximum = 500                                       ' �ڐ�1�`100(0.01m�`5.00mm)
        TBarHiPitch.Minimum = 1
        '###110
        TBarHiPitch.Value = dblTchMoval(IDX_HPT) * 100          ' 0.01mm�P��

        ' Pause Time���͈͊O�Ȃ�͈͓��ɕύX����
        If (dblTchMoval(IDX_PAU) < cPAU_LO) Then dblTchMoval(IDX_PAU) = cPAU_LO
        If (dblTchMoval(IDX_PAU) > cPAU_HI) Then dblTchMoval(IDX_PAU) = cPAU_HI

        ' Pause Time�̖ڐ���ݒ肷��
        TBarPause.TickFrequency = dblTchMoval(IDX_PAU) * 20             ' 0.5�b�P��
        TBarPause.Maximum = 20                                          ' �ڐ�1�`20(0.05�b�`1.00�b)
        TBarPause.Minimum = 1
        '###110
        TBarPause.Value = dblTchMoval(IDX_PAU) * 20             ' 0.5�b�P��

        ' �ړ��s�b�`��\������
        LblTchMoval0.Text = dblTchMoval(IDX_PIT).ToString("0.0000")
        LblTchMoval1.Text = dblTchMoval(IDX_HPT).ToString("0.0000")
        LblTchMoval2.Text = dblTchMoval(IDX_PAU).ToString("0.0000")

    End Sub
#End Region

#Region "�g���b�N�o�[�̃X���C�_�[�ړ�����"
    '''=========================================================================
    '''<summary>�g���b�N�o�[�̃X���C�_�[�ړ�����</summary>
    '''<param name="Index">       (INP)0=LOW�߯�, 1=HIGH�߯�, 2=Pause</param>
    '''<param name="TBarLowPitch">(I/O)�X���C�_�[1(Low�߯�)</param>
    '''<param name="TBarHiPitch"> (I/O)�X���C�_�[2(HIGH�߯�)</param>
    '''<param name="TBarPause">   (I/O)�X���C�_�[3(Pause Time)</param>
    '''<param name="LblTchMoval0">(I/O)�ڐ�1(Low Pich Label)</param>
    '''<param name="LblTchMoval1">(I/O)�ڐ�2(Low�߯� Label)</param>
    '''<param name="LblTchMoval2">(I/O)�ڐ�3(HIGH�߯� Label)</param>
    '''<param name="dblTchMoval"> (I/O)�߯��ޔ���(0=�߯�, 1=HIGH�߯�, 2=Pause Time)</param>
    '''=========================================================================
    Public Sub SetSliderPitch(ByRef Index As Short, _
                              ByRef TBarLowPitch As System.Windows.Forms.TrackBar, _
                              ByRef TBarHiPitch As System.Windows.Forms.TrackBar, _
                              ByRef TBarPause As System.Windows.Forms.TrackBar, _
                              ByRef LblTchMoval0 As System.Windows.Forms.Label, _
                              ByRef LblTchMoval1 As System.Windows.Forms.Label, _
                              ByRef LblTchMoval2 As System.Windows.Forms.Label, _
                              ByRef dblTchMoval() As Double)

        Dim lVal As Integer

        ' BP�̈ړ��s�b�`����ݒ肷��
        Select Case Index
            Case IDX_PIT    ' LOW�߯�
                lVal = TBarLowPitch.Value                       ' �ײ�ޖڐ��l�擾
                dblTchMoval(Index) = 0.001 * lVal               ' LOW�߯��l�ύX
                LblTchMoval0.Text = dblTchMoval(Index).ToString("0.0000")
                LblTchMoval0.Refresh()

            Case IDX_HPT    ' HIGH�߯�
                lVal = TBarHiPitch.Value                        ' �ײ�ޖڐ��l�擾
                dblTchMoval(Index) = 0.01 * lVal                ' HIGH�߯��l�ύX
                LblTchMoval1.Text = dblTchMoval(Index).ToString("0.0000")
                LblTchMoval1.Refresh()

            Case IDX_PAU    ' Pause Time
                lVal = TBarPause.Value                          ' �ײ�ޖڐ��l�擾
                dblTchMoval(Index) = 0.05 * lVal                ' �ړ��s�b�`�Ԃ̃|�[�Y�l�ύX
                LblTchMoval2.Text = dblTchMoval(Index).ToString("0.0000")
                LblTchMoval2.Refresh()
        End Select

    End Sub
#End Region

    '========================================================================================
    '   �i�n�f�����ʏ����p�e���L�[���͏���
    '========================================================================================
#Region "�e���L�[�_�E���T�u���[�`��"
    '''=========================================================================
    '''<summary>�e���L�[�_�E���T�u���[�`��</summary>
    ''' <param name="KeyCode">(INP)�L�[�R�[�h</param>
    '''=========================================================================
    Public Sub Sub_10KeyDown(ByVal KeyCode As Short)

        Dim strMSG As String

        Try
            ' Num Lock��
            Select Case (KeyCode)
                Case System.Windows.Forms.Keys.NumPad2                      ' ��  (KeyCode =  98(&H62)
                    InpKey = InpKey Or &H1000                               ' +Y ON(��)
                Case System.Windows.Forms.Keys.NumPad8                      ' ��  (KeyCode = 104(&H68)
                    InpKey = InpKey Or &H800                                ' -Y ON(��)
                Case System.Windows.Forms.Keys.NumPad4                      ' ��  (KeyCode = 100(&H64)
                    InpKey = InpKey Or &H400                                ' +X ON(��)
                Case System.Windows.Forms.Keys.NumPad6                      ' ��  (KeyCode = 102(&H66)
                    InpKey = InpKey Or &H200                                ' -X ON(��)
                Case System.Windows.Forms.Keys.NumPad9                      ' PgUp(KeyCode = 105(&H69)
                    InpKey = InpKey Or &HA00                                ' -X -Y ON
                Case System.Windows.Forms.Keys.NumPad7                      ' Home(KeyCode = 103(&H67))
                    InpKey = InpKey Or &HC00                                ' +X -Y ON
                Case System.Windows.Forms.Keys.NumPad1                      ' End(KeyCode =   97(&H61)
                    InpKey = InpKey Or &H1400                               ' +X +Y ON
                Case System.Windows.Forms.Keys.NumPad3                      ' PgDn(KeyCode =  99(&H63)
                    InpKey = InpKey Or &H1200                               ' -X +Y ON
                Case System.Windows.Forms.Keys.NumPad5                      ' 5�� (KeyCode = 101(&H65)
                    'Call BtnHI_Click(sender, e)                             ' HI�{�^�� ON/OFF
            End Select

            ' �g���b�v�G���[������
        Catch ex As Exception
            strMSG = "Globals.Sub_10KeyDown() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�e���L�[�A�b�v�T�u���[�`��"
    '''=========================================================================
    '''<summary>�e���L�[�A�b�v�T�u���[�`��</summary>
    ''' <param name="KeyCode">(INP)�L�[�R�[�h</param>
    '''=========================================================================
    Public Sub Sub_10KeyUp(ByVal KeyCode As Short)

        Dim strMSG As String

        Try
            ' Num Lock��
            Select Case (KeyCode)
                Case System.Windows.Forms.Keys.NumPad2                      ' ��  (KeyCode =  98(&H62)
                    InpKey = InpKey And Not &H1000                          ' +Y OFF
                Case System.Windows.Forms.Keys.NumPad8                      ' ��  (KeyCode = 104(&H68)
                    InpKey = InpKey And Not &H800                           ' -Y OFF
                Case System.Windows.Forms.Keys.NumPad4                      ' ��  (KeyCode = 100(&H64)
                    InpKey = InpKey And Not &H400                           ' +X OFF
                Case System.Windows.Forms.Keys.NumPad6                      ' ��  (KeyCode = 102(&H66)
                    InpKey = InpKey And Not &H200                           ' -X OFF
                Case System.Windows.Forms.Keys.NumPad9                      ' PgUp(KeyCode = 105(&H69)
                    InpKey = InpKey And Not &HA00                           ' -X -Y OFF
                Case System.Windows.Forms.Keys.NumPad7                      ' Home(KeyCode = 103(&H67))
                    InpKey = InpKey And Not &HC00                           ' +X -Y OFF
                Case System.Windows.Forms.Keys.NumPad1                      ' End(KeyCode =   97(&H61)
                    InpKey = InpKey And Not &H1400                          ' +X +Y OFF
                Case System.Windows.Forms.Keys.NumPad3                      ' PgDn(KeyCode =  99(&H63)
                    InpKey = InpKey And Not &H1200                          ' -X +Y OFF
            End Select

            ' �g���b�v�G���[������
        Catch ex As Exception
            strMSG = "Globals.Sub_10KeyUp() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "���̓L�[�R�[�h�̃N���A"
    '''=========================================================================
    ''' <summary>
    ''' ���̓L�[�R�[�h�̃N���A
    ''' </summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub ClearInpKey()

        Try
            InpKey = 0

            ' �g���b�v�G���[������
        Catch ex As Exception
            MsgBox("Globals.ClearInpKey() TRAP ERROR = " + ex.Message)
        End Try
    End Sub
#End Region

    '    '===========================================================================
    '    '   �O���[�o�����\�b�h��`
    '    '===========================================================================
    '#Region "�@�B�n�̃p�����[�^�ݒ�"
    '    '''=========================================================================
    '    '''<summary>�@�B�n�̃p�����[�^�ݒ�</summary>
    '    '''<remarks></remarks>
    '    '''=========================================================================
    '    Public Sub SetMechanicalParam()

    '        Dim BpSoftLimitX As Integer
    '        Dim BpSoftLimitY As Integer

    '        With gSysPrm.stDEV
    '            ' ���ʑΉ�
    '            If gSysPrm.stCTM.giSPECIAL = customASAHI And typPlateInfo.strDataName = "2" Then
    '                .gfTrimX = .gfTrimX2                            ' TRIM POSITION X(mm)
    '                .gfTrimY = .gfTrimY2                            ' TRIM POSITION Y(mm)
    '                .gfExCmX = .gfExCmX2                            ' Externla Camera Offset X(mm)
    '                .gfExCmY = .gfExCmY2                            ' Externla Camera Offset Y(mm)
    '                .gfRot_X1 = .gfRot_X2                           ' ��]���S X
    '                .gfRot_Y1 = .gfRot_Y2                           ' ��]���S Y
    '                '(2010/11/16)���L�����͕s�v
    '                'Else
    '                '    gSysPrm.stDEV.gfTrimX = gSysPrm.stDEV.gfTrimX   ' TRIM POSITION X(mm)
    '                '    gSysPrm.stDEV.gfTrimY = gSysPrm.stDEV.gfTrimY   ' TRIM POSITION Y(mm)
    '                '    gSysPrm.stDEV.gfExCmX = gSysPrm.stDEV.gfExCmX   ' Externla Camera Offset X(mm)
    '                '    gSysPrm.stDEV.gfExCmY = gSysPrm.stDEV.gfExCmY   ' Externla Camera Offset Y(mm)
    '                '    gSysPrm.stDEV.gfRot_X1 = gSysPrm.stDEV.gfRot_X1 ' ��]���S X
    '                '    gSysPrm.stDEV.gfRot_Y1 = gSysPrm.stDEV.gfRot_Y1 ' ��]���S Y
    '            End If
    '            ''''(2010/11/16) ����m�F�㉺�L�R�����g�͍폜
    '            'gStartX = gSysPrm.stDEV.gfTrimX
    '            'gStartY = gSysPrm.stDEV.gfTrimY

    '            'BpSize����BP�̃\�t�g���~�b�g�iBP�̃\�t�g�ғ��͈́j��ݒ�
    '            Select Case (.giBpSize)
    '                Case 0
    '                    BpSoftLimitX = 50
    '                    BpSoftLimitY = 50
    '                Case 1
    '                    BpSoftLimitX = 80
    '                    BpSoftLimitY = 80
    '                Case 2
    '                    BpSoftLimitX = 100
    '                    BpSoftLimitY = 60
    '                Case 3
    '                    BpSoftLimitX = 60
    '                    BpSoftLimitY = 100
    '                Case Else
    '                    BpSoftLimitX = 80
    '                    BpSoftLimitY = 80
    '            End Select

    '            '''''2009/07/23 minato
    '            ''''    �g�����|�W�V�������ύX����Ă��邽�߁A
    '            ''''    INTRTM���̃V�X�e���p�����[�^���X�V����K�v������B
    '            Call ZSYSPARAM2(.giPrbTyp, .gfSminMaxZ2, .giZPTimeOn, .giZPTimeOff, _
    '                        .giXYtbl, .gfSmaxX, .gfSmaxY, gSysPrm.stIOC.glAbsTime, _
    '                        .gfTrimX, .gfTrimY, BpSoftLimitX, BpSoftLimitY)
    '        End With
    '    End Sub
    '#End Region

    '#Region "U��Ď��s���ʎ擾"
    '    '''=========================================================================
    '    '''<summary>U��Ď��s���ʎ擾</summary>
    '    '''<param name="rn">(INP) ��R�ԍ�</param>
    '    '''<param name="s"> (OUT) ���s����</param>
    '    '''<returns>0=����, 0�ȊO=�G���[</returns>
    '    '''=========================================================================
    '    Public Function RetrieveUCutResult(ByVal rn As Short, ByRef s As String) As Short

    '        Dim cn As Short
    '        Dim n As Short
    '        Dim f As Double
    '        Dim r As Integer

    '        s = ""
    '        RetrieveUCutResult = 0

    '        If gSysPrm.stSPF.giUCutKind = 0 Then
    '            Exit Function
    '        End If

    '        On Error GoTo ErrTrap

    '        For cn = 1 To typResistorInfoArray(rn).intCutCount
    '            s = typResistorInfoArray(rn).ArrCut(cn).strCutType       ' Cut pattern
    '            If s = "H" Then
    '                s = ""
    '                '  U��Ď��s���ʎ擾
    '                r = UCUT_RESULT(rn, cn, n, f)
    '                If (r <> 0) Then
    '                    MsgBox("Internal error  X001-" & Str(r), MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, gAppName)
    '                    RetrieveUCutResult = 1
    '                    Exit Function
    '                End If

    '                If n = 255 Then                                 ' 255 ��UCUT���s���Ă��Ȃ��ꍇ
    '                    s = Form1.Utility1.sFormat(f, "0.000000", 10 + 7) & " n** "
    '                ElseIf n >= 0 And n <= 19 Then
    '                    n = n + 1
    '                    s = Form1.Utility1.sFormat(f, "0.000000", 10 + 7) & " " & "n" & n.ToString("00") & " "
    '                ElseIf n = 254 Then                             ' �p�����[�^�e�[�u���ɊY�������R�ԍ������������ꍇ
    '                    s = Form1.Utility1.sFormat(f, "0.000000", 10 + 7) & " n** "
    '                Else                                            ' �ςȒl
    '                    RetrieveUCutResult = 2
    '                    Exit Function
    '                End If
    '            Else
    '                s = ""
    '            End If
    '        Next

    '        Exit Function

    'ErrTrap:
    '        Resume ErrTrap1
    'ErrTrap1:
    '        Dim er As Integer
    '        er = Err.Number
    '        On Error GoTo 0
    '        MsgBox("Internal error X002-" & Str(er), MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, gAppName)
    '    End Function
    '#End Region

    '#Region "�ް���R�ԍ����ڼ��ް���R�ԍ����擾����"
    '    '''=========================================================================
    '    '''<summary>�ް���R�ԍ����ڼ��ް���R�ԍ����擾����</summary>
    '    '''<param name="br">(INP) ��R�ԍ�</param>
    '    '''<returns>0�ȏ�=ڼ��ް���R�ԍ�, -1=�Ȃ�</returns>
    '    '''=========================================================================
    '    Public Function GetRatio1BaseNum(ByVal br As Short) As Short

    '        Dim n As Short

    '        For n = 1 To gRegistorCnt
    '            ' �x�[�X��R ?
    '            If typResistorInfoArray(n).intResNo = br Then
    '                GetRatio1BaseNum = n
    '                Exit Function
    '            End If
    '        Next
    '        GetRatio1BaseNum = -1

    '    End Function
    '#End Region

    '#Region "�O���[�v��,�u���b�N��,�`�b�v��(��R��),�`�b�v�T�C�Y���擾����(�s�w/�s�x�e�B�[�`���O�p)"
    '    '''=========================================================================
    '    ''' <summary>�O���[�v��,�u���b�N��,�`�b�v��(��R��),�`�b�v�T�C�Y���擾����</summary>
    '    ''' <param name="AppMode">  (INP)���[�h</param>
    '    ''' <param name="Gn">       (OUT)�O���[�v��</param>
    '    ''' <param name="RnBn">     (OUT)�`�b�v��(�s�w�e�B�[�`���O��)�܂���
    '    '''                              �u���b�N��(�s�x�e�B�[�`���O��)</param>
    '    ''' <param name="DblChipSz">(OUT)�`�b�v�T�C�Y</param>
    '    ''' <returns>0=����, 0�ȊO=�G���[</returns>
    '    '''=========================================================================
    '    Public Function GetChipNumAndSize(ByVal AppMode As Short, ByRef Gn As Short, ByRef RnBn As Short, ByRef DblChipSz As Double) As Short

    '        Dim ChipNum As Short                                        ' �`�b�v��(��R��)
    '        Dim ChipSzX As Double                                       ' �`�b�v�T�C�YX
    '        Dim ChipSzY As Double                                       ' �`�b�v�T�C�YY
    '        Dim strMSG As String

    '        Try
    '            ' �O����(CHIP/NET����)
    '            ChipNum = typPlateInfo.intResistCntInGroup              ' �`�b�v��(��R��) = 1�O���[�v��(1�T�[�L�b�g��)��R��
    '            ChipSzX = typPlateInfo.dblChipSizeXDir                  ' �`�b�v�T�C�YX,Y
    '            ChipSzY = typPlateInfo.dblChipSizeYDir

    '            ' �v���[�g�f�[�^����O���[�v��, �u���b�N��, �`�b�v��(��R��), �`�b�v�T�C�Y���擾����
    '            If (AppMode = APP_MODE_TX) Then
    '                '----- �s�w�e�B�[�`���O�� -----
    '                ' �`�b�v��(��R��)��Ԃ�
    '                RnBn = ChipNum                                      ' 1�O���[�v��(1�T�[�L�b�g��)��R�����Z�b�g
    '                ' �O���[�v����Ԃ�
    '                Gn = typPlateInfo.intGroupCntInBlockXBp             ' �a�o�O���[�v��(�T�[�L�b�g��)���Z�b�g
    '                ' �`�b�v�T�C�Y��Ԃ�
    '                If (typPlateInfo.intResistDir = 0) Then             ' �`�b�v���т�X���� ?
    '                    DblChipSz = System.Math.Abs(ChipSzX)
    '                Else
    '                    DblChipSz = System.Math.Abs(ChipSzY)
    '                End If

    '            Else
    '                '----- �s�x�e�B�[�`���O�� -----
    '                ' �O���[�v����Ԃ�
    '                Gn = typPlateInfo.intGroupCntInBlockYStage          ' �u���b�N��Stage�O���[�v�����Z�b�g
    '                ' �u���b�N���ƃ`�b�v�T�C�Y��Ԃ�
    '                If (typPlateInfo.intResistDir = 0) Then             ' �`�b�v���т�X���� ?
    '                    RnBn = typPlateInfo.intBlockCntYDir             ' �u���b�N��Y���Z�b�g
    '                    DblChipSz = System.Math.Abs(ChipSzY)            ' �`�b�v�T�C�YY���Z�b�g
    '                Else
    '                    RnBn = typPlateInfo.intBlockCntXDir             ' �u���b�N��X���Z�b�g
    '                    DblChipSz = System.Math.Abs(ChipSzX)            ' �`�b�v�T�C�YX���Z�b�g
    '                End If
    '            End If

    '            strMSG = "GetChipNumAndSize() Gn=" + Gn.ToString("0") + ", RnBn=" + RnBn.ToString("0") + ", ChipSZ=" + DblChipSz.ToString("0.00000")
    '            Console.WriteLine(strMSG)
    '            Return (cFRS_NORMAL)                                    ' Return�l = ����

    '            ' �g���b�v�G���[������
    '        Catch ex As Exception
    '            strMSG = "globals.GetChipNumAndSize() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '            Return (cERR_TRAP)                                      ' Return�l = ��O�G���[
    '        End Try
    '    End Function
    '#End Region

    '#Region "�������V�I���̃x�[�X��R�ԍ�(�z��̓Y��)��Ԃ�"
    '    '''=========================================================================
    '    '''<summary>�������V�I���̃x�[�X��R�ԍ�(�z��̓Y��)��Ԃ�</summary>
    '    '''<param name="rr">(INP) ��R�ԍ�</param> 
    '    '''<param name="br">(OUT) �x�[�X��R�ԍ�(�z��̓Y��)</param>  
    '    '''<remarks>�����a�����������V�I�@�\(TKY���ڐA)</remarks>
    '    '''=========================================================================
    '    Public Sub GetRatio3Br(ByRef rr As Short, ByRef br As Short)

    '        Dim i As Short
    '        Dim wRn As Short
    '        Dim wGn As Short
    '        Dim wBr As Short
    '        Dim wBr2 As Short

    '        ' �������V�I���[�h(3�`9)�łȂ���Βʏ�̃x�[�X��R�ԍ���Ԃ�
    '        wRn = typResistorInfoArray(rr).intResNo                         ' ��R�ԍ�
    '        wGn = typResistorInfoArray(rr).intTargetValType                 ' �ڕW�l��ʁi0:��Βl, 1:���V�I�A2�F�v�Z��, 3�`9:��ٰ�ߔԍ��j
    '        wBr = GetRatio1BaseNum(typResistorInfoArray(rr).intBaseResNo)   ' �x�[�X��R�ԍ�(�Y��)
    '        wBr2 = -1
    '        If (wGn < 3) Or (wGn > 9) Then                                  ' �������V�I���[�h(3�`9)�łȂ� ? 
    '            GoTo STP_END
    '        End If

    '        ' �������V�I�Ȃ瑊���ٰ�ߔԍ�����������
    '        For i = 1 To gRegistorCnt                                       ' ��R�����J��Ԃ�
    '            If (wRn <> typResistorInfoArray(i).intResNo) Then           ' ��R�ԍ�=�������g��SKIP
    '                If (wGn = typResistorInfoArray(i).intTargetValType) Then            ' �����ٰ�ߔԍ� ?
    '                    wBr2 = GetRatio1BaseNum(typResistorInfoArray(i).intBaseResNo)   ' �x�[�X��R�ԍ�(�Y��)
    '                    Exit For
    '                End If
    '            End If
    '        Next i

    '        ' �x�[�X��R��FT�l�̑傫�������x�[�X��R�ԍ��Ƃ���
    '        If (wBr2 < 0) Then GoTo STP_END '                               ' �����ٰ�ߔԍ���������Ȃ����� ?
    '        If (gfFinalTest(wBr2) > gfFinalTest(wBr)) Then                  ' �����FT�l���傫�� ?
    '            wBr = wBr2
    '        End If

    'STP_END:
    '        'br = wBr                                                       ' �x�[�X��R�ԍ���Ԃ�
    '        br = wBr - 1                                                    ' �x�[�X��R�ԍ���Ԃ� ###244

    '    End Sub
    '#End Region

    '#Region "���V�I(�v�Z��)���̃x�[�X��R�ԍ�(�z��̓Y��)��Ԃ�"
    '    '''=========================================================================
    '    '''<summary>���V�I(�v�Z��)���̃x�[�X��R�ԍ������R�f�[�^�̔z��̓Y����Ԃ�###123</summary>
    '    '''<param name="br">(INP)�x�[�X��R�ԍ�(�z��̓Y��)</param> 
    '    '''<param name="rr">(OUT)��R�f�[�^�̔z��̓Y��(1 ORG)</param> 
    '    '''<remarks></remarks>
    '    '''=========================================================================
    '    Public Sub GetRatio2Rn(ByVal br As Short, ByRef rr As Short)

    '        Dim Rn As Short

    '        ' �x�[�X��R�ԍ�����������
    '        For Rn = 1 To gRegistorCnt                                      ' ��R�����J��Ԃ�
    '            If (typResistorInfoArray(Rn).intBaseResNo = br) Then
    '                rr = Rn
    '                Exit Sub
    '            End If
    '        Next Rn

    '    End Sub
    '#End Region

    '#Region "Z/Z2�ړ�(ON/OFF) "
    '    '''=========================================================================
    '    '''<summary>Z/Z2�ړ�(ON/OFF) </summary>
    '    '''<param name="MD">  (INP)Ӱ��(0 = OFF�ʒu�ړ�, 1 = ON�ʒu�ړ�)</param> 
    '    '''<param name="Z2ON">(INP)Z2 ON�ʒu(OPTION)</param>  
    '    '''<remarks>0=����, 0�ȊO=�G���[</remarks>
    '    '''=========================================================================
    '    Public Function Sub_Probe_OnOff(ByVal MD As Integer, Optional ByVal Z2ON As Double = 0.0#) As Integer

    '        Dim r As Integer
    '        Dim strMSG As String

    '        Try
    '            ' �y�v���[�u���I���ʒu�ֈړ�
    '            Sub_Probe_OnOff = cFRS_NORMAL                       ' Return�l = ����
    '            If (MD = 1) Then                                    ' ON ?
    '                r = Form1.System1.EX_PROBON(gSysPrm)                   ' Z ON�ʒu�ֈړ�
    '                If (r <> cFRS_NORMAL) Then                      ' �װ ?
    '                    Sub_Probe_OnOff = r                         ' Return�l = ����~��(��ү���ނ͕\����)
    '                    Exit Function
    '                End If
    '                If ((gSysPrm.stDEV.giPrbTyp And 2) = 2) Then    ' ������۰�ނȂ��Ȃ�NOP
    '                    r = Form1.System1.EX_PROBON2(gSysPrm, Z2ON)        ' Z2 ON�ʒu�ֈړ�
    '                    If (r <> cFRS_NORMAL) Then                  ' �װ ?
    '                        Sub_Probe_OnOff = r                     ' Return�l = ����~��(��ү���ނ͕\����)
    '                        Exit Function
    '                    End If
    '                End If

    '                ' �y�v���[�u���I�t�ʒu�ֈړ�
    '            Else
    '                If ((gSysPrm.stDEV.giPrbTyp And 2) = 2) Then    ' ������۰�ނȂ��Ȃ�NOP
    '                    r = Form1.System1.EX_PROBOFF2(gSysPrm)             ' Z2 OFF�ʒu�ֈړ�
    '                    If (r <> cFRS_NORMAL) Then                  ' �װ ?
    '                        Sub_Probe_OnOff = r                     ' Return�l = ����~��(��ү���ނ͕\����)
    '                        Exit Function
    '                    End If
    '                End If
    '                r = Form1.System1.EX_PROBOFF(gSysPrm)                  ' Z OFF�ʒu�ֈړ�
    '                If (r <> cFRS_NORMAL) Then                      ' �װ ?
    '                    Sub_Probe_OnOff = r                         ' Return�l = ����~��(��ү���ނ͕\����)
    '                    Exit Function
    '                End If
    '            End If
    '            Exit Function

    '            ' �g���b�v�G���[������
    '        Catch ex As Exception
    '            strMSG = "globals.Sub_Probe_OnOff() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '            Return (cERR_TRAP)                                  ' Return�l = ��O�G���[
    '        End Try
    '    End Function
    '#End Region

    '#Region "���H��������͂���(FL��)"
    '    '''=========================================================================
    '    '''<summary>���H��������͂���(FL��)</summary>
    '    ''' <param name="CondNum">(I/O)���H�����ԍ�</param>
    '    ''' <param name="dQrate"> (I/O)Q���[�g(KHz)</param>
    '    ''' <param name="Owner">  (INP)�I�[�i�[</param>
    '    ''' <returns>0=����, 0�ȊO=�G���[</returns>
    '    ''' <remarks>�L�����u���[�V�����A�J�b�g�ʒu�␳(�O���J����)�̏\���J�b�g�p</remarks>
    '    '''=========================================================================
    '    Public Function Sub_FlCond(ByRef CondNum As Integer, ByRef dQrate As Double, ByVal Owner As IWin32Window) As Integer

    '        Dim r As Integer
    '        Dim ObjForm As Object = Nothing
    '        Dim strMSG As String

    '        Try
    '            ' ���H��������͂���(FL��)
    '            r = cFRS_NORMAL                                             ' Return�l = ����
    '            If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then          ' FL�łȂ���
    '                CondNum = 0                                             ' ���H�����ԍ�(Dmy)

    '            Else                                                        ' FL���͉��H��������͂���
    '                ' ���H�������͉�ʕ\��
    '                ObjForm = New FrmFlCond()                               ' �I�u�W�F�N�g����
    '                Call ObjForm.ShowDialog(Owner, CondNum)                 ' ���H�������͉�ʕ\��
    '                r = ObjForm.GetResult(CondNum, dQrate)                  ' ���H�����擾

    '                ' �I�u�W�F�N�g�J��
    '                If (ObjForm Is Nothing = False) Then
    '                    Call ObjForm.Close()                                ' �I�u�W�F�N�g�J��
    '                    Call ObjForm.Dispose()                              ' ���\�[�X�J��
    '                End If
    '            End If

    '            Return (r)                                                  ' Return�l�ݒ�

    '            ' �g���b�v�G���[������
    '        Catch ex As Exception
    '            strMSG = "globals.Sub_FlCond() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '            Return (cERR_TRAP)                                          ' Return�l = �ׯ�ߴװ����
    '        End Try
    '    End Function
    '#End Region

    '#Region "�\���J�b�g���s��"
    '    '''=========================================================================
    '    '''<summary>�\���J�b�g���s��</summary>
    '    ''' <param name="BPx">         (INP)�J�b�g�ʒuX</param>
    '    ''' <param name="BPy">         (INP)�J�b�g�ʒuY</param>
    '    ''' <param name="CondNum">     (INP)���H�����ԍ�(FL�p)</param>
    '    ''' <param name="dQrate">      (INP)Q���[�g(KHz)</param>
    '    ''' <param name="dblCutLength">(INP)�J�b�g��</param>
    '    ''' <param name="dblCutSpeed"> (INP)�J�b�g���x</param>
    '    ''' <returns>0=����, 0�ȊO=�G���[</returns>
    '    ''' <remarks>���\����Ă̒�����BP���ړ����Ă�������
    '    '''            �L�����u���[�V�����A�J�b�g�ʒu�␳(�O���J����)�̏\���J�b�g�p</remarks>
    '    '''=========================================================================
    '    Public Function CrossCutExec(ByVal BPx As Double, ByVal BPy As Double, ByVal CondNum As Integer, _
    '                                 ByVal dQrate As Double, ByVal dblCutLength As Double, ByVal dblCutSpeed As Double) As Integer

    '        Dim r As Integer
    '        Dim intXANG As Integer
    '        Dim intYANG As Integer
    '        Dim strMSG As String
    '        Dim stCutCmnPrm As CUT_COMMON_PRM                               ' �J�b�g�p�����[�^

    '        Try
    '            '-------------------------------------------------------------------
    '            '   ��������
    '            '-------------------------------------------------------------------
    '            Call InitCutParam(stCutCmnPrm)                              ' �J�b�g�p�����[�^������

    '            ' �J�b�g�p�x��ݒ肷��
    '            Select Case (gSysPrm.stDEV.giBpDirXy)
    '                Case 0      ' x��, y��
    '                    intXANG = 180
    '                    intYANG = 270
    '                Case 1      ' x��, y��
    '                    intXANG = 0
    '                    intYANG = 270
    '                Case 2      ' x��, y��
    '                    intXANG = 180
    '                    intYANG = 90
    '                Case 3      ' x��, y��
    '                    intXANG = 0
    '                    intYANG = 90
    '            End Select

    '            ' �J�b�g�p�����[�^(�J�b�g���\����)��ݒ肷��
    '            stCutCmnPrm.CutInfo.srtMoveMode = 2                         ' ���샂�[�h�i0:�g���~���O�A1:�e�B�[�`���O�A2:�����J�b�g�j
    '            stCutCmnPrm.CutInfo.srtCutMode = 4                          ' �J�b�g���[�h�́u�΂߁v
    '            stCutCmnPrm.CutInfo.dblTarget = 1000.0#                     ' �ڕW�l = 1�Ƃ���
    '            stCutCmnPrm.CutInfo.srtSlope = 4                            ' 4:��R����{�X���[�v
    '            stCutCmnPrm.CutInfo.srtMeasType = 0                         ' ����^�C�v(0:����(3��)�A1:�����x(2000��)
    '            stCutCmnPrm.CutInfo.dblAngle = intXANG                      ' �J�b�g�p�x(X��)

    '            ' �J�b�g�p�����[�^(���H�ݒ�\����)��ݒ肷��
    '            stCutCmnPrm.CutCond.CutLen.dblL1 = dblCutLength             ' �J�b�g��(Line1�p)
    '            stCutCmnPrm.CutCond.SpdOwd.dblL1 = dblCutSpeed              ' �J�b�g�X�s�[�h�i���H�j(Line1�p)
    '            stCutCmnPrm.CutCond.QRateOwd.dblL1 = dQrate                 ' �J�b�gQ���[�g�i���H�j(Line1�p)
    '            stCutCmnPrm.CutCond.CondOwd.srtL1 = CondNum                 ' �J�b�g�����ԍ��i���H�j(Line1�p)

    '            ' Q���[�g(FL���ȊO)�܂��͉��H�����ԍ�(FL��)��ݒ肷��
    '            If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then          ' FL�łȂ� ?
    '                Call QRATE(dQrate)                                      ' Q���[�g�ݒ�(KHz)
    '            Else                                                        ' ���H�����ԍ���ݒ肷��(FL��)
    '                Call QRATE(dQrate)                                      ' Q���[�g�ݒ�(KHz)
    '                r = FLSET(FLMD_CNDSET, CondNum)                         ' ���H�����ԍ��ݒ�
    '                If (r <> cFRS_NORMAL) Then GoTo STP_ERR_FL
    '            End If

    '            '-------------------------------------------------------------------
    '            '   �\���J�b�g��X�����J�b�g����
    '            '-------------------------------------------------------------------
    '            ' BP��X���n�_�ֈړ�����(��Βl�ړ�)
    '            r = Form1.System1.EX_MOVE(gSysPrm, BPx - (dblCutLength / 2), BPy, 1)
    '            If (r <> cFRS_NORMAL) Then                                  ' �G���[ ?
    '                Return (r)
    '            End If
    '            ' �\���J�b�g��X�����J�b�g����
    '            r = Sub_CrossCut(stCutCmnPrm)                               ' X���J�b�g
    '            If (r <> cFRS_NORMAL) Then                                  ' �G���[ ?
    '                Return (r)
    '            End If
    '            ' BP�𒆐S�֖߂�
    '            r = Form1.System1.EX_MOVE(gSysPrm, BPx, BPy, 1)
    '            If (r <> cFRS_NORMAL) Then                                  ' �G���[ ?
    '                Return (r)
    '            End If
    '            Call System.Threading.Thread.Sleep(500)                     ' Wait(msec)

    '            '-------------------------------------------------------------------
    '            '   �\���J�b�g��Y�����J�b�g����
    '            '-------------------------------------------------------------------
    '            ' BP��Y���n�_�ֈړ�����(��Βl�ړ�)
    '            r = Form1.System1.EX_MOVE(gSysPrm, BPx, BPy - (dblCutLength / 2), 1)
    '            If (r <> cFRS_NORMAL) Then                                  ' �G���[ ?
    '                Return (r)
    '            End If
    '            ' �\���J�b�g��Y�����J�b�g����
    '            stCutCmnPrm.CutInfo.dblAngle = intYANG                      ' �J�b�g�p�x(Y��)
    '            r = Sub_CrossCut(stCutCmnPrm)                               ' Y���J�b�g
    '            If (r <> cFRS_NORMAL) Then                                  ' �G���[ ?
    '                Return (r)
    '            End If
    '            ' BP�𒆐S�֖߂�
    '            r = Form1.System1.EX_MOVE(gSysPrm, BPx, BPy, 1)
    '            If (r <> cFRS_NORMAL) Then                                  ' �G���[ ?
    '                Return (r)
    '            End If
    '            Call System.Threading.Thread.Sleep(500)                     ' Wait(msec)

    '            Return (cFRS_NORMAL)

    '            ' ���H�����ԍ��̐ݒ�G���[��(FL��)
    'STP_ERR_FL:
    '            strMSG = MSG_151                                            ' "���H�����̐ݒ�Ɏ��s���܂����"
    '            Call Form1.System1.TrmMsgBox(gSysPrm, strMSG, vbOKOnly, gAppName)
    '            Return (cFRS_ERR_RST)                                       ' Return�l = Cancel

    '            ' �g���b�v�G���[������
    '        Catch ex As Exception
    '            strMSG = "globals.CrossCutExec() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '            Return (cERR_TRAP)                                          ' Return�l = �ׯ�ߴװ����
    '        End Try
    '    End Function
    '#End Region

    '#Region "�\���J�b�g��X���܂���Y�����J�b�g����"
    '    '''=========================================================================
    '    '''<summary>�\���J�b�g��X���܂���Y�����J�b�g����</summary>
    '    ''' <param name="stCutCmnPrm">(INP)�J�b�g�p�����[�^</param>
    '    ''' <returns>0=����, 0�ȊO=�G���[</returns>
    '    ''' <remarks>���\���J�b�g�ʒu��BP���ړ����Ă�������
    '    '''            �L�����u���[�V�����A�J�b�g�ʒu�␳(�O���J����)�̏\���J�b�g�p</remarks>
    '    '''=========================================================================
    '    Private Function Sub_CrossCut(ByRef stCutCmnPrm As CUT_COMMON_PRM) As Integer

    '        Dim r As Integer
    '        Dim strMSG As String

    '        Try
    '            ' �\���J�b�g��X���܂���Y�����J�b�g����
    '            r = TRIM_ST(stCutCmnPrm)                                    ' ST�J�b�g
    '            r = Form1.System1.EX_ZGETSRVSIGNAL(gSysPrm, r, 0)
    '            If (r < cFRS_NORMAL) Then                                   ' �G���[ ?
    '                Return (r)
    '            End If
    '            Return (cFRS_NORMAL)

    '            ' �g���b�v�G���[������
    '        Catch ex As Exception
    '            strMSG = "globals.Sub_CrossCut() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '            Return (cERR_TRAP)                                          ' Return�l = �ׯ�ߴװ����
    '        End Try
    '    End Function
    '#End Region

    '#Region "�J�b�g�p�����[�^������������"
    '    '''=========================================================================
    '    '''<summary>�J�b�g�p�����[�^������������</summary>
    '    ''' <param name="pstCutCmnPrm">(I/O)�J�b�g�p�����[�^</param>
    '    ''' <remarks>�L�����u���[�V�����A�J�b�g�ʒu�␳(�O���J����)�̏\���J�b�g�p</remarks>
    '    '''=========================================================================
    '    Private Sub InitCutParam(ByRef pstCutCmnPrm As CUT_COMMON_PRM)

    '        Dim strMSG As String

    '        Try
    '            ' �J�b�g�p�����[�^������������(�J�b�g���\����)
    '            pstCutCmnPrm.CutInfo.srtMoveMode = 1                        ' ���샂�[�h�i0:�g���~���O�A1:�e�B�[�`���O�A2:�����J�b�g�j
    '            pstCutCmnPrm.CutInfo.srtCutMode = 0                         ' �J�b�g���[�h(0:�m�[�}���A1:���^�[���A2:���g���[�X�A3:�΂߁j
    '            pstCutCmnPrm.CutInfo.dblTarget = 0.0#                       ' �ڕW�l
    '            pstCutCmnPrm.CutInfo.srtSlope = 4                           ' 4:��R����{�X���[�v
    '            pstCutCmnPrm.CutInfo.srtMeasType = 0                        ' ����^�C�v(0:����(3��)�A1:�����x(2000��)
    '            pstCutCmnPrm.CutInfo.dblAngle = 0.0#                        ' �J�b�g�p�x
    '            pstCutCmnPrm.CutInfo.dblLTP = 0.0#                          ' L�^�[���|�C���g
    '            pstCutCmnPrm.CutInfo.srtLTDIR = 0                           ' L�^�[����̕���
    '            pstCutCmnPrm.CutInfo.dblRADI = 0.0#                         ' R����]���a�iU�J�b�g�Ŏg�p�j
    '            '                                                           ' For Hook Or UCut
    '            pstCutCmnPrm.CutInfo.dblRADI2 = 0.0#                        ' R2����]���a�iU�J�b�g�Ŏg�p�j
    '            pstCutCmnPrm.CutInfo.srtHkOrUType = 0                       ' HookCut(3)��U�J�b�g�i3�ȊO�j�̎w��B
    '            '                                                           ' For Index
    '            pstCutCmnPrm.CutInfo.srtIdxScnCnt = 0                       ' �C���f�b�N�X/�X�L�����J�b�g��(1�`32767)
    '            pstCutCmnPrm.CutInfo.srtIdxMeasMode = 0                     ' �C���f�b�N�X���胂�[�h�i0:��R�A1:�d���A2:�O���j
    '            '                                                           ' For EdgeSense
    '            pstCutCmnPrm.CutInfo.dblEsPoint = 0.0#                      ' �G�b�W�Z���X�|�C���g
    '            pstCutCmnPrm.CutInfo.dblRdrJdgVal = 0.0#                    ' ���_�[��������ω���
    '            pstCutCmnPrm.CutInfo.dblMinJdgVal = 0.0#                    ' ���_�[�J�b�g��Œዖ�e�ω���
    '            pstCutCmnPrm.CutInfo.srtEsAftCutCnt = 0                     ' ���_�[�ؔ�����̃J�b�g�񐔁i����񐔁j
    '            pstCutCmnPrm.CutInfo.srtMinOvrNgCnt = 0                     ' ���_�[���o����A�Œ�ω��ʂ̘A��Over���e��
    '            pstCutCmnPrm.CutInfo.srtMinOvrNgMode = 0                    ' �A��Over����NG�����i0:NG���薢���{, 1:NG������{�B���_�[���؂�, 2:NG���薢���{�B���_�[�؏グ�j
    '            '                                                           ' For Scan
    '            pstCutCmnPrm.CutInfo.dblStepPitch = 0.0#                    ' �X�e�b�v�ړ��s�b�`
    '            pstCutCmnPrm.CutInfo.srtStepDir = 0                         ' �X�e�b�v����

    '            ' �J�b�g�p�����[�^������������(���H�ݒ�\����)
    '            pstCutCmnPrm.CutCond.CutLen.dblL1 = 0.0#                    ' �J�b�g��(Line1�p)
    '            pstCutCmnPrm.CutCond.CutLen.dblL2 = 0.0#                    ' �J�b�g��(Line2�p)
    '            pstCutCmnPrm.CutCond.CutLen.dblL3 = 0.0#                    ' �J�b�g��(Line3�p)
    '            pstCutCmnPrm.CutCond.CutLen.dblL4 = 0.0#                    ' �J�b�g��(Line4�p)

    '            pstCutCmnPrm.CutCond.SpdOwd.dblL1 = 0.0#                    ' �J�b�g�X�s�[�h�i���H�j(Line1�p)
    '            pstCutCmnPrm.CutCond.SpdOwd.dblL2 = 0.0#                    ' �J�b�g�X�s�[�h�i���H�j(Line2�p)
    '            pstCutCmnPrm.CutCond.SpdOwd.dblL3 = 0.0#                    ' �J�b�g�X�s�[�h�i���H�j(Line3�p)
    '            pstCutCmnPrm.CutCond.SpdOwd.dblL4 = 0.0#                    ' �J�b�g�X�s�[�h�i���H�j(Line4�p)

    '            pstCutCmnPrm.CutCond.SpdRet.dblL1 = 0.0#                    ' �J�b�g�X�s�[�h�i���H�j(Line1�p)
    '            pstCutCmnPrm.CutCond.SpdRet.dblL2 = 0.0#                    ' �J�b�g�X�s�[�h�i���H�j(Line2�p)
    '            pstCutCmnPrm.CutCond.SpdRet.dblL3 = 0.0#                    ' �J�b�g�X�s�[�h�i���H�j(Line3�p)
    '            pstCutCmnPrm.CutCond.SpdRet.dblL4 = 0.0#                    ' �J�b�g�X�s�[�h�i���H�j(Line4�p)

    '            pstCutCmnPrm.CutCond.QRateOwd.dblL1 = 0.0#                  ' �J�b�gQ���[�g�i���H�j(Line1�p)
    '            pstCutCmnPrm.CutCond.QRateOwd.dblL2 = 0.0#                  ' �J�b�gQ���[�g�i���H�j(Line2�p)
    '            pstCutCmnPrm.CutCond.QRateOwd.dblL3 = 0.0#                  ' �J�b�gQ���[�g�i���H�j(Line3�p)
    '            pstCutCmnPrm.CutCond.QRateOwd.dblL4 = 0.0#                  ' �J�b�gQ���[�g�i���H�j(Line4�p)

    '            pstCutCmnPrm.CutCond.QRateRet.dblL1 = 0.0#                  ' �J�b�gQ���[�g�i���H�j(Line1�p)
    '            pstCutCmnPrm.CutCond.QRateRet.dblL2 = 0.0#                  ' �J�b�gQ���[�g�i���H�j(Line2�p)
    '            pstCutCmnPrm.CutCond.QRateRet.dblL3 = 0.0#                  ' �J�b�gQ���[�g�i���H�j(Line3�p)
    '            pstCutCmnPrm.CutCond.QRateRet.dblL4 = 0.0#                  ' �J�b�gQ���[�g�i���H�j(Line4�p)

    '            pstCutCmnPrm.CutCond.CondOwd.srtL1 = 0                      ' �J�b�g�����ԍ��i���H�j(Line1�p)
    '            pstCutCmnPrm.CutCond.CondOwd.srtL2 = 0                      ' �J�b�g�����ԍ��i���H�j(Line2�p)
    '            pstCutCmnPrm.CutCond.CondOwd.srtL3 = 0                      ' �J�b�g�����ԍ��i���H�j(Line3�p)
    '            pstCutCmnPrm.CutCond.CondOwd.srtL4 = 0                      ' �J�b�g�����ԍ��i���H�j(Line4�p)

    '            pstCutCmnPrm.CutCond.CondRet.srtL1 = 0                      ' �J�b�g�����ԍ��i���H�j(Line1�p)
    '            pstCutCmnPrm.CutCond.CondRet.srtL2 = 0                      ' �J�b�g�����ԍ��i���H�j(Line2�p)
    '            pstCutCmnPrm.CutCond.CondRet.srtL3 = 0                      ' �J�b�g�����ԍ��i���H�j(Line3�p)
    '            pstCutCmnPrm.CutCond.CondRet.srtL4 = 0                      ' �J�b�g�����ԍ��i���H�j(Line4�p)

    '            Exit Sub

    '            ' �g���b�v�G���[������
    '        Catch ex As Exception
    '            strMSG = "globals.InitCutParam() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '        End Try
    '        Exit Sub
    '    End Sub
    '#End Region

    '#Region "�p�^�[���F�������s���A����ʂ�Ԃ�"
    '    '''=========================================================================
    '    ''' <summary>�p�^�[���F�������s���A����ʂ�Ԃ�</summary>
    '    ''' <param name="GrpNo">    (INP)�O���[�v�ԍ�</param>
    '    ''' <param name="TmpNo">    (INP)�p�^�[���ԍ�</param>
    '    ''' <param name="fCorrectX">(OUT)�����X</param> 
    '    ''' <param name="fCorrectY">(OUT)�����Y</param>
    '    ''' <param name="coef">     (OUT)���֌W��</param> 
    '    ''' <returns>cFRS_NORMAL  = ����
    '    '''          cFRS_ERR_PTN = �p�^�[���p�^�[���F���G���[
    '    '''          ��L�ȊO     = ���̑��G���[</returns>
    '    ''' <remarks>�E�p�^�[���F���ʒu�փe�[�u���͈ړ��ςł��邱��
    '    '''          �E�L�����u���[�V�����A�J�b�g�ʒu�␳(�O���J����)�p
    '    ''' </remarks>
    '    '''=========================================================================
    '    Public Function Sub_PatternMatching(ByRef GrpNo As Short, ByRef TmpNo As Short, ByRef fCorrectX As Double, ByRef fCorrectY As Double, ByRef coef As Double) As Integer

    '        Dim ret As Short = cFRS_NORMAL
    '        Dim crx As Double = 0.0                                         ' �����X
    '        Dim cry As Double = 0.0                                         ' �����Y
    '        Dim fcoeff As Double = 0.0                                      ' ���֒l
    '        Dim Thresh As Double = 0.0                                      ' 臒l
    '        Dim r As Integer = cFRS_NORMAL                                  ' �֐��l
    '        Dim strMSG As String = ""

    '        Try
    '#If VIDEO_CAPTURE = 1 Then
    '            fCorrectX = 0.0
    '            fCorrectY = 0.0
    '            coef = 0.8
    '            Return (cFRS_NORMAL)   
    '#End If
    '            ' �p�^�[���}�b�`���O���̃e���v���[�g�O���[�v�ԍ���ݒ肷��(������ƒx���Ȃ�)
    '            If (giTempGrpNo <> GrpNo) Then                              ' �e���v���[�g�O���[�v�ԍ����ς���� ?
    '                giTempGrpNo = GrpNo                                     ' ���݂̃e���v���[�g�O���[�v�ԍ���ޔ�
    '                Form1.VideoLibrary1.SelectTemplateGroup(GrpNo)          ' �e���v���[�g�O���[�v�ԍ��ݒ�
    '            End If

    '            ' 臒l�擾
    '            Thresh = gDllSysprmSysParam_definst.GetPtnMatchThresh(GrpNo, TmpNo)
    '            coef = 0.0                                                  ' ��v�x

    '            ' �p�^�[���}�b�`���O(�O���J����)���s��(Video.ocx���g�p)
    '            ret = Form1.VideoLibrary1.PatternMatching_EX(TmpNo, 1, True, crx, cry, fcoeff)
    '            If (ret = cFRS_NORMAL) Then
    '                r = cFRS_NORMAL                                         ' Return�l = ����
    '                fCorrectX = crx                                         ' �����X
    '                fCorrectY = cry                                         ' �����Y
    '                '' �}�b�`�����p�^�[���̑���ʒu���炸��ʂ����߂�
    '                'fCorrectX = crx / 1000.0#
    '                'fCorrectY = -cry / 1000.0#
    '                coef = fcoeff                                           ' ���֌W��
    '                strMSG = "�p�^�[���F������"
    '                If (fcoeff < Thresh) Then
    '                    r = cFRS_ERR_PT2                                    ' �p�^�[���F���G���[(臒l�G���[)
    '                    strMSG = "�p�^�[���F���G���[(臒l�G���[)"
    '                End If
    '                strMSG = strMSG + " (���֌W��=" + Format(fcoeff, "0.000") + " �����X=" + Format(crx, "0.0000") + ", �����X=" + Format(cry, "0.0000") + ")"
    '            Else
    '                r = cFRS_ERR_PTN                                        ' �p�^�[���F���G���[(�p�^�[����������Ȃ�����)
    '                strMSG = "�p�^�[���F���G���[(�p�^�[����������Ȃ�����)"
    '            End If

    '            ' �㏈��
    '            Console.WriteLine(strMSG)
    '            Return (r)

    '            ' �g���b�v�G���[������ 
    '        Catch ex As Exception
    '            strMSG = "globals.Sub_PatternMatching() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '            Return (cERR_TRAP)                                          ' Return�l = ��O�G���[
    '        End Try
    '    End Function
    '#End Region

    '#Region "�w��u���b�N�̒����փe�[�u�����ړ�����"
    '    '''=========================================================================
    '    '''<summary>�w��u���b�N�̒����փe�[�u�����ړ�����</summary>
    '    '''<param name="intCamera">(INP)��׎��(0:������� 1:�O�����)</param>
    '    '''<param name="iXPlate">(INP)XPlateNo</param> 
    '    '''<param name="iYPlate">(INP)YPlateNo</param>  
    '    '''<param name="iXBlock">(INP)XBlockNo</param> 
    '    '''<param name="iYBlock">(INP)YBlockNo</param>   
    '    '''<remarks>�\����Ĉʒu��è��ݸ��߲�Ă�����ڰ��ް�
    '    '''         ��PP47�̒l�����ꂽ�Ƃ��낪���S�ƂȂ�
    '    '''         �������ڰĂ��w�肵�Ă��Ӗ��Ȃ�</remarks>
    '    '''=========================================================================
    '    Public Function XYTableMoveBlock(ByRef intCamera As Short, ByRef iXPlate As Short, ByRef iYPlate As Short, ByRef iXBlock As Short, ByRef iYBlock As Short) As Short

    '        Dim dblX As Double
    '        Dim dblY As Double
    '        Dim dblRotX As Double
    '        Dim dblRotY As Double
    '        Dim dblPSX As Double
    '        Dim dblPSY As Double
    '        Dim dblBsoX As Double
    '        Dim dblBsoY As Double
    '        Dim dblBSX As Double
    '        Dim dblBSY As Double
    '        Dim intCDir As Short
    '        Dim dblTrimPosX As Double
    '        Dim dblTrimPosY As Double
    '        Dim dblTOffsX As Double
    '        Dim dblTOffsY As Double
    '        Dim dblStepInterval As Double
    '        Dim Del_x As Double
    '        Dim Del_y As Double
    '        Dim r As Short
    '        Dim strMSG As String

    '        Try
    '            dblRotX = 0
    '            dblRotY = 0

    '            ' ����߼޼��X,Y�擾
    '            dblTrimPosX = gSysPrm.stDEV.gfTrimX                 ' ����߼޼��X,Y�擾
    '            dblTrimPosY = gSysPrm.stDEV.gfTrimY
    '            ' ð��وʒu�̾�Ă̎擾
    '            dblTOffsX = typPlateInfo.dblTableOffsetXDir : dblTOffsY = typPlateInfo.dblTableOffsetYDir

    '            Call CalcBlockSize(dblBSX, dblBSY)                  ' ��ۯ����ގZ�o

    '            ' ��ۯ����޵̾�ĎZ�o�@��ۯ�����/2 ��ۯ��̏ی���XY�Ƃ���1 ð��ق̏ی���1
    '            dblBsoX = (dblBSX / 2.0#) * 1 * 1                   ' Table.BDirX * Table.dir
    '            dblBsoY = (dblBSY / 2) * 1                          ' Table.BDirY;

    '            ' �ƕ␳��ѵ̾��X,Y
    '            Del_x = gfCorrectPosX
    '            Del_y = gfCorrectPosY

    '            ' giBpDirXy ���W�n�̐ݒ�(���ѐݒ�)
    '            ' 0:XY NOM(�E��)  1:X REV(����)  2:Y REV(�E��)  3:XY REV(����)
    '            ' ���ݸވʒu���W (+or-) ��]���a + ð��ٵ̾�� (+or-) ��ۯ����޵̾�� + ð��ٕ␳��
    '            Select Case gSysPrm.stDEV.giBpDirXy

    '                Case 0 ' x��, y��
    '                    dblX = dblTrimPosX + dblRotX + dblTOffsX + dblBsoX + Del_x
    '                    dblY = dblTrimPosY + dblRotY + dblTOffsY + dblBsoY + Del_y

    '                Case 1 ' x��, y��
    '                    dblX = dblTrimPosX - dblRotX + dblTOffsX - dblBsoX + Del_x
    '                    dblY = dblTrimPosY + dblRotY + dblTOffsY + dblBsoY + Del_y

    '                Case 2 ' x��, y��
    '                    dblX = dblTrimPosX + dblRotX + dblTOffsX + dblBsoX + Del_x
    '                    dblY = dblTrimPosY - dblRotY + dblTOffsY - dblBsoY + Del_y

    '                Case 3 ' x��, y��
    '                    dblX = dblTrimPosX - dblRotX + dblTOffsX - dblBsoX + Del_x
    '                    dblY = dblTrimPosY - dblRotY + dblTOffsY - dblBsoY + Del_y

    '            End Select

    '            If (1 = intCamera) Then                             ' �O����׈ʒu���Z ?
    '                dblX = dblX + gSysPrm.stDEV.gfExCmX
    '                dblY = dblY + gSysPrm.stDEV.gfExCmY
    '            End If

    '            '�ï�ߊԊu�̎Z�o
    '            intCDir = typPlateInfo.intResistDir                 ' �`�b�v���ѕ����擾(CHIP-NET�̂�)

    '            If intCDir = 0 Then                                 ' X����
    '                dblStepInterval = CalcStepInterval(iYBlock)     ' �ï�߲�����َZ�o(Y��)
    '                If gSysPrm.stDEV.giBpDirXy = 0 Or gSysPrm.stDEV.giBpDirXy = 1 Then ' ð���Y�������]�Ȃ�
    '                    dblY = dblY + dblStepInterval
    '                Else                                            ' ð���Y�������]
    '                    dblY = dblY - dblStepInterval
    '                End If
    '            Else                                                ' Y����
    '                dblStepInterval = CalcStepInterval(iXBlock)     ' �ï�߲�����َZ�o(X��)
    '                If gSysPrm.stDEV.giBpDirXy = 0 Or gSysPrm.stDEV.giBpDirXy = 2 Then ' ð���X�������]�Ȃ�
    '                    dblX = dblX + dblStepInterval
    '                Else                                            ' ð���X�������]
    '                    dblX = dblX - dblStepInterval
    '                End If
    '            End If

    '            ' ��ڰ�/��ۯ��ʒu�̑��΍��W�v�Z
    '            dblPSX = 0.0 : dblPSY = 0.0                         ' ��ڰĻ��ގ擾(0�Œ�)
    '            Select Case gSysPrm.stDEV.giBpDirXy

    '                Case 0 ' x��, y��
    '                    dblX = dblX + ((dblPSX * CInt(iXPlate - 1)) + (dblBSX * CInt(iXBlock - 1)))
    '                    dblY = dblY + ((dblPSY * CInt(iYPlate - 1)) + (dblBSY * CInt(iYBlock - 1)))

    '                Case 1 ' x��, y��
    '                    dblX = dblX - ((dblPSX * CInt(iXPlate - 1)) + (dblBSX * CInt(iXBlock - 1)))
    '                    dblY = dblY + ((dblPSY * CInt(iYPlate - 1)) + (dblBSY * CInt(iYBlock - 1)))

    '                Case 2 ' x��, y��
    '                    dblX = dblX + ((dblPSX * CInt(iXPlate - 1)) + (dblBSX * CInt(iXBlock - 1)))
    '                    dblY = dblY - ((dblPSY * CInt(iYPlate - 1)) + (dblBSY * CInt(iYBlock - 1)))

    '                Case 3 ' x��, y��
    '                    dblX = dblX - ((dblPSX * CInt(iXPlate - 1)) + (dblBSX * CInt(iXBlock - 1)))
    '                    dblY = dblY - ((dblPSY * CInt(iYPlate - 1)) + (dblBSY * CInt(iYBlock - 1)))

    '            End Select

    '            ' �w����ڰ�/��ۯ��ʒu��XYð��ِ�Βl�ړ�
    '            r = Form1.System1.XYtableMove(gSysPrm, dblX, dblY)
    '            Return (r)                                      ' Return

    '            ' �g���b�v�G���[������
    '        Catch ex As Exception
    '            strMSG = "globals.XYTableMoveBlock() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '            Return (cERR_TRAP)                              ' Return�l = ��O�G���[ 
    '        End Try

    '    End Function
    '#End Region

    '#Region "GPIB�R�}���h��ݒ肷��"
    '    '''=========================================================================
    '    '''<summary>GPIB�R�}���h��ݒ肷��</summary>
    '    '''<param name="pltInfo">(OUT)�v���[�g�f�[�^</param>
    '    '''=========================================================================
    '    Public Sub SetGpibCommand(ByRef pltInfo As PlateInfo)

    '        Dim strDAT As String
    '        Dim strMSG As String

    '        Try
    '            ' ADEX AX-1152�p�ݒ�R�}���h��ݒ肷��
    '            pltInfo.intGpibDefAdder = giGpibDefAdder                ' GPIB�A�h���X 
    '            pltInfo.intGpibDefDelimiter = 0                         ' �����ݒ�(�����)(�Œ�)
    '            pltInfo.intGpibDefTimiout = 100                         ' �����ݒ�(��ѱ��)(�Œ�)
    '            If (pltInfo.intGpibMeasSpeed = 0) Then                  ' ���葬�x(0:�ᑬ, 1:����)
    '                strDAT = "W0"
    '            Else
    '                strDAT = "W1"
    '            End If

    '            '// ���胂�[�h�Ő؂�ւ�
    '            If (pltInfo.intGpibMeasMode = 0) Then                   ' ���胂�[�h(0:���, 1:�΍�)
    '                strDAT = strDAT + "FR"                              ' ���胂�[�h=���
    '                strDAT = strDAT + "LL00000" + "LH15000"             ' ����/������~�b�g�̐ݒ�
    '            Else

    '                strDAT = strDAT + "FD"                              ' ���胂�[�h=�΍�
    '                strDAT = strDAT + "DL-5000" + "DH+5000"             ' ����/������~�b�g�̐ݒ�
    '            End If

    '            pltInfo.strGpibInitCmnd1 = strDAT                       ' �����������1
    '            pltInfo.strGpibInitCmnd2 = ""                           ' �����������2
    '            pltInfo.strGpibTriggerCmnd = "E"                        ' �ض޺����

    '            ' �g���b�v�G���[������
    '        Catch ex As Exception
    '            strMSG = "globals.SetGpibCommand() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '        End Try
    '    End Sub
    '#End Region
    '    '----- ###211�� -----
    '#Region "START/RESET�L�[�����҂��T�u���[�`��"
    '    '''=========================================================================
    '    ''' <summary>START/RESET�L�[�����҂��T�u���[�`��</summary>
    '    ''' <param name="Md">(INP)cFRS_ERR_START                = START�L�[�����҂�
    '    '''                       cFRS_ERR_RST                  = RESET�L�[�����҂�
    '    '''                       cFRS_ERR_START + cFRS_ERR_RST = START/RESET�L�[�����҂�
    '    ''' </param>
    '    ''' <param name="bZ">(INP)True=Z�L�[�����`�F�b�N����, False=���Ȃ� ###220</param>
    '    ''' <returns>cFRS_ERR_START = START�L�[����
    '    '''          cFRS_ERR_RST   = RESET�L�[����
    '    '''          cFRS_ERR_Z     = Z�L�[����
    '    '''          ��L�ȊO=�G���[
    '    ''' </returns>
    '    '''=========================================================================
    '    Public Function WaitStartRestKey(ByVal Md As Integer, ByVal bZ As Boolean) As Integer

    '        Dim sts As Long = 0
    '        Dim r As Long = 0
    '        Dim ExitFlag As Integer
    '        Dim strMSG As String

    '        Try
    '            ' �p�����[�^�`�F�b�N
    '            If (Md = 0) Then
    '                Return (-1 * ERR_CMD_PRM)                               ' �p�����[�^�G���[
    '            End If

    '#If cOFFLINEcDEBUG Then                                                 ' OffLine���ޯ��ON ?(��FormReset���őO�ʕ\���Ȃ̂ŉ��L�̂悤�ɂ��Ȃ���MsgBox���őO�ʕ\������Ȃ�)
    '            Dim Dr As System.Windows.Forms.DialogResult
    '            Dr = MessageBox.Show("START SW CHECK", "Debug", MessageBoxButtons.OKCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
    '            If (Dr = System.Windows.Forms.DialogResult.OK) Then
    '                ExitFlag = cFRS_ERR_START                               ' Return�l = START�L�[����
    '            Else
    '                ExitFlag = cFRS_ERR_RST                                 ' Return�l = RESET�L�[����
    '            End If
    '            Return (ExitFlag)
    '#End If
    '            ' START/RESET�L�[�����҂�
    '            Call ZCONRST()                                              ' �R���\�[���L�[���b�`����
    '            ExitFlag = -1
    '            Call Form1.System1.SetSysParam(gSysPrm)                     ' �V�X�e���p�����[�^�̐ݒ�(OcxSystem�p)

    '            ' START/RESET�L�[�����҂�
    '            Do
    '                r = STARTRESET_SWCHECK(False, sts)                      ' START/RESET SW�����`�F�b�N
    '                If (sts = cFRS_ERR_RST) And ((Md = cFRS_ERR_RST) Or (Md = cFRS_ERR_START + cFRS_ERR_RST)) Then
    '                    ExitFlag = cFRS_ERR_RST                             ' ExitFlag = Cancel(RESET�L�[)
    '                ElseIf (sts = cFRS_ERR_START) And ((Md = cFRS_ERR_START) Or (Md = cFRS_ERR_START + cFRS_ERR_RST)) Then
    '                    ExitFlag = cFRS_ERR_START                           ' ExitFlag = OK(START�L�[)
    '                End If
    '                '----- ###220�� -----
    '                If (bZ = True) Then
    '                    r = Z_SWCHECK(sts)                                  ' Z SW�����`�F�b�N
    '                    If (sts <> 0) Then
    '                        ExitFlag = cFRS_ERR_Z                           ' ExitFlag = Z�L�[����
    '                    End If
    '                End If
    '                '----- ###220�� -----
    '                System.Windows.Forms.Application.DoEvents()             ' ���b�Z�[�W�|���v
    '                Call System.Threading.Thread.Sleep(100)                 ' Wait(msec)

    '                ' �V�X�e���G���[�`�F�b�N
    '                r = Form1.System1.SysErrChk_ForVBNET(giAppMode)
    '                If (r <> cFRS_NORMAL) Then                              ' ����~��(���b�Z�[�W�͕\����) ?
    '                    ExitFlag = r
    '                    Exit Do
    '                End If
    '            Loop While (ExitFlag = -1)

    '            Call ZCONRST()                                              ' �R���\�[���L�[���b�`����
    '            Return (ExitFlag)

    '            ' �g���b�v�G���[������ 
    '        Catch ex As Exception
    '            strMSG = "Globals.WaitRestKey() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '            Return (cERR_TRAP)                                          ' Retuen�l = �ׯ�ߴװ����
    '        End Try
    '    End Function
    '#End Region
    '    '----- ###211�� -----

    '    '===========================================================================
    '    '   �ėp�^�C�}�[
    '    '===========================================================================
    '    Private bTmTimeOut As Boolean                                       ' �^�C���A�E�g�t���O

    '#Region "�ėp�^�C�}�[����"
    '    '''=========================================================================
    '    ''' <summary>�ėp�^�C�}�[����</summary>
    '    ''' <param name="TimerTM">(I/O)�^�C�}�[</param>
    '    ''' <param name="TimeVal">(INP)�^�C���A�E�g�l(msec)</param>
    '    ''' <remarks>�^�C�}�[���������ꍇ��TimerTM_Dispose��Call���ă^�C�}�[��j�����鎖</remarks>
    '    '''=========================================================================
    '    Public Sub TimerTM_Create(ByRef TimerTM As System.Threading.Timer, ByVal TimeVal As Integer)

    '        Dim strMSG As String

    '        Try
    '            ' �^�C���A�E�g�`�F�b�N�p�^�C�}�[�I�u�W�F�N�g�̍쐬(TimerTM_Tick��TimeVal msec�Ԋu�Ŏ��s����)
    '            bTmTimeOut = False                                          ' �^�C���A�E�g�t���OOFF
    '            'TimerTM = New System.Threading.Timer(New System.Threading.TimerCallback(AddressOf TimerTM_Tick), Nothing, TimeVal, TimeVal)
    '            TimerTM = New System.Threading.Timer(New System.Threading.TimerCallback(AddressOf TimerTM_Tick), Nothing, System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)

    '            ' �g���b�v�G���[������ 
    '        Catch ex As Exception
    '            strMSG = "globals.TimerTM_Create() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '        End Try
    '    End Sub
    '#End Region

    '#Region "�ėp�^�C�}�[�J�n"
    '    '''=========================================================================
    '    ''' <summary>�ėp�^�C�}�[�J�n</summary>
    '    ''' <param name="TimerTM">(INP)�^�C�}�[</param>
    '    '''=========================================================================
    '    Public Sub TimerTM_Start(ByRef TimerTM As System.Threading.Timer, ByVal TimeVal As Integer)

    '        Dim strMSG As String

    '        Try
    '            If (TimerTM Is Nothing) Then Return
    '            TimerTM.Change(TimeVal, TimeVal)
    '            Exit Sub

    '            ' �g���b�v�G���[������ 
    '        Catch ex As Exception
    '            strMSG = "globals.TimerTM_Start() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '        End Try
    '    End Sub
    '#End Region

    '#Region "�ėp�^�C�}�[��~(�R�[���o�b�N���\�b�h(TimerTM_Tick)�̌ďo�����~����)"
    '    '''=========================================================================
    '    ''' <summary>�ėp�^�C�}�[��~(�R�[���o�b�N���\�b�h(TimerTM_Tick)�̌ďo�����~����)</summary>
    '    ''' <param name="TimerTM">(INP)�^�C�}�[</param>
    '    '''=========================================================================
    '    Public Sub TimerTM_Stop(ByRef TimerTM As System.Threading.Timer)

    '        Dim strMSG As String

    '        Try
    '            ' �R�[���o�b�N���\�b�h�̌ďo�����~����
    '            If (TimerTM Is Nothing) Then Return
    '            TimerTM.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)
    '            Exit Sub

    '            ' �g���b�v�G���[������ 
    '        Catch ex As Exception
    '            strMSG = "globals.TimerTM_Stop() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '        End Try
    '    End Sub
    '#End Region

    '#Region "�ėp�^�C�}�[��j������"
    '    '''=========================================================================
    '    ''' <summary>�ėp�^�C�}�[��j������</summary>
    '    ''' <param name="TimerTM">(I/O)�^�C�}�[</param>
    '    '''=========================================================================
    '    Public Sub TimerTM_Dispose(ByRef TimerTM As System.Threading.Timer)

    '        Dim strMSG As String

    '        Try
    '            ' �R�[���o�b�N���\�b�h�̌ďo�����~����
    '            If (TimerTM Is Nothing) Then Return
    '            TimerTM.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)
    '            TimerTM.Dispose()                                           ' �^�C�}�[��j������
    '            Exit Sub

    '            ' �g���b�v�G���[������ 
    '        Catch ex As Exception
    '            strMSG = "globals.TimerTM_Dispose() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '        End Try
    '    End Sub
    '#End Region

    '#Region "�^�C���A�E�g�t���O��Ԃ�"
    '    '''=========================================================================
    '    ''' <summary>�^�C���A�E�g�t���O��Ԃ�</summary>
    '    ''' <returns>Trur=�^�C���A�E�g, False=�^�C���A�E�g�łȂ�</returns>
    '    '''=========================================================================
    '    Public Function TimerTM_Sts() As Boolean

    '        Dim strMSG As String

    '        Try
    '            ' �^�C���A�E�g�t���O��Ԃ�
    '            Return (bTmTimeOut)

    '            ' �g���b�v�G���[������ 
    '        Catch ex As Exception
    '            strMSG = "globals.TimerTM_Sts() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '            Return (bTmTimeOut)
    '        End Try
    '    End Function
    '#End Region

    '#Region "�^�C�}�[�C�x���g(�w��^�C�}�Ԋu���o�߂������ɔ���)"
    '    '''=========================================================================
    '    ''' <summary>�^�C�}�[�C�x���g(�w��^�C�}�Ԋu���o�߂������ɔ���)</summary>
    '    ''' <param name="Sts">(INP)</param>
    '    '''=========================================================================
    '    Private Sub TimerTM_Tick(ByVal Sts As Object)

    '        Dim strMSG As String

    '        Try
    '            bTmTimeOut = True                                           ' �^�C���A�E�g�t���OON
    '            Exit Sub

    '            ' �g���b�v�G���[������ 
    '        Catch ex As Exception
    '            strMSG = "globals.TimerTM_Tick() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '        End Try
    '    End Sub
    '#End Region

    '    '========================================================================================
    '    '   ���z�}�\���֘A����
    '    '========================================================================================
    '    '' '' ''#Region "���z�}�\��"
    '    '' '' ''    '''=========================================================================
    '    '' '' ''    '''<summary>���z�}�\��</summary>
    '    '' '' ''    '''<remarks></remarks>
    '    '' '' ''    '''=========================================================================
    '    '' '' ''    Public Sub picGraphAccumulationRedraw()

    '    '' '' ''        Dim iCnt As Short '����
    '    '' '' ''        Dim lMax As Integer
    '    '' '' ''        Dim lScale As Integer
    '    '' '' ''        Dim lScaleMax As Integer
    '    '' '' ''        Dim dblGraphDiv As Double
    '    '' '' ''        Dim dblGraphTop As Double
    '    '' '' ''        Dim digL As Integer
    '    '' '' ''        Dim digH As Integer
    '    '' '' ''        Dim digSW As Integer

    '    '' '' ''        lMax = 0
    '    '' '' ''        If (bFgDispGrp) Then
    '    '' '' ''            Form1.lblGraphAccumulationTitle.Text = MSG_TRIM_04
    '    '' '' ''            Form1.lblMinValue.Text = dblMinIT.ToString("0.000") ' �ŏ��l
    '    '' '' ''            Form1.lblMaxValue.Text = dblMaxIT.ToString("0.000") ' �ő�l

    '    '' '' ''            For iCnt = 0 To (MAX_SCALE_RNUM - 1)

    '    '' '' ''                glRegistNum(iCnt) = glRegistNumIT(iCnt)

    '    '' '' ''                If lMax < glRegistNum(iCnt) Then
    '    '' '' ''                    lMax = glRegistNum(iCnt)
    '    '' '' ''                End If

    '    '' '' ''                gDistRegNumLblAry(iCnt).Text = CStr(glRegistNum(iCnt)) ' ���z�O���t��R��

    '    '' '' ''            Next
    '    '' '' ''        Else

    '    '' '' ''            Form1.lblGraphAccumulationTitle.Text = MSG_TRIM_05
    '    '' '' ''            Form1.lblMinValue.Text = dblMinFT.ToString("0.000") ' �ŏ��l
    '    '' '' ''            Form1.lblMaxValue.Text = dblMaxFT.ToString("0.000") ' �ő�l

    '    '' '' ''            For iCnt = 0 To (MAX_SCALE_RNUM - 1)

    '    '' '' ''                glRegistNum(iCnt) = glRegistNumFT(iCnt)

    '    '' '' ''                If lMax < glRegistNum(iCnt) Then
    '    '' '' ''                    lMax = glRegistNum(iCnt)
    '    '' '' ''                End If

    '    '' '' ''                gDistRegNumLblAry(iCnt).Text = CStr(glRegistNum(iCnt)) ' ���z�O���t��R��

    '    '' '' ''            Next
    '    '' '' ''        End If

    '    '' '' ''        Form1.lblGoodChip.Text = CStr(lOkChip)                        ' OK��
    '    '' '' ''        Form1.lblNgChip.Text = CStr(lNgChip)                          ' NG��

    '    '' '' ''        ' �덷�ް�������(IT)
    '    '' '' ''        Call Form1.GetMoveMode(digL, digH, digSW)
    '    '' '' ''        If ITNx_cnt >= 0 Then
    '    '' '' ''            If (digL = 0) Then                                 ' x0���[�h ?
    '    '' '' ''                ' ���ϒl�擾
    '    '' '' ''                dblAverageIT = Form1.Utility1.GetAverage(ITNx, ITNx_cnt + 1)
    '    '' '' ''                ' �W���΍��̎擾
    '    '' '' ''                dblDeviationIT = Form1.Utility1.GetDeviation(ITNx, ITNx_cnt + 1, dblAverageIT)
    '    '' '' ''            End If
    '    '' '' ''        End If

    '    '' '' ''        ' �덷�ް�������(FT)
    '    '' '' ''        If FTNx_cnt >= 0 Then
    '    '' '' ''            ' ���ϒl�擾
    '    '' '' ''            dblAverageFT = Form1.Utility1.GetAverage(FTNx, FTNx_cnt + 1)
    '    '' '' ''            ' �W���΍��̎擾
    '    '' '' ''            dblDeviationFT = Form1.Utility1.GetDeviation(FTNx, FTNx_cnt + 1, dblAverageFT)
    '    '' '' ''        End If

    '    '' '' ''        If (bFgDispGrp) Then
    '    '' '' ''            Form1.lblDeviationValue.Text = dblDeviationIT.ToString("0.000000") ' �W���΍�(IT)
    '    '' '' ''        Else
    '    '' '' ''            Form1.lblDeviationValue.Text = dblDeviationFT.ToString("0.000000") ' �W���΍�(FT)
    '    '' '' ''        End If

    '    '' '' ''        If (bFgDispGrp) Then
    '    '' '' ''            dblAverage = dblAverageIT
    '    '' '' ''        Else
    '    '' '' ''            dblAverage = dblAverageFT
    '    '' '' ''        End If
    '    '' '' ''        Form1.lblAverageValue.Text = dblAverage.ToString("0.000")             ' ���ϒl

    '    '' '' ''        lScaleMax = 0 ' �I�[�g�X�P�[�����O
    '    '' '' ''        lScale = 100
    '    '' '' ''        Do
    '    '' '' ''            If (lScale > lMax) Then
    '    '' '' ''                lScaleMax = lScale
    '    '' '' ''            ElseIf ((lScale * 2) > lMax) Then
    '    '' '' ''                lScaleMax = (lScale * 2)
    '    '' '' ''            ElseIf ((lScale * 5) > lMax) Then
    '    '' '' ''                lScaleMax = (lScale * 5)
    '    '' '' ''            End If
    '    '' '' ''            lScale = lScale * 10
    '    '' '' ''        Loop While (0 = lScaleMax) And (MAX_SCALE_NUM > lScale)
    '    '' '' ''        If (0 = lScaleMax) Then
    '    '' '' ''            lScaleMax = MAX_SCALE_NUM + 1
    '    '' '' ''        End If

    '    '' '' ''        If (bFgDispGrp) Then
    '    '' '' ''            If ((0 >= typResistorInfoArray(1).dblInitTest_LowLimit) And (0 <= typResistorInfoArray(1).dblInitTest_HighLimit)) Then
    '    '' '' ''                dblGraphDiv = (typResistorInfoArray(1).dblInitTest_HighLimit * 1.5 - typResistorInfoArray(1).dblInitTest_LowLimit * 1.5) / 10
    '    '' '' ''                dblGraphTop = typResistorInfoArray(1).dblInitTest_HighLimit * 1.5
    '    '' '' ''            ElseIf ((0 >= typResistorInfoArray(1).dblInitTest_LowLimit) And (0 > typResistorInfoArray(1).dblInitTest_HighLimit)) Then
    '    '' '' ''                dblGraphDiv = (typResistorInfoArray(1).dblInitTest_HighLimit / 1.5 - typResistorInfoArray(1).dblInitTest_LowLimit * 1.5) / 10
    '    '' '' ''                dblGraphTop = typResistorInfoArray(1).dblInitTest_HighLimit * 1.5
    '    '' '' ''            ElseIf ((0 < typResistorInfoArray(1).dblInitTest_LowLimit) And (0 <= typResistorInfoArray(1).dblInitTest_HighLimit)) Then
    '    '' '' ''                dblGraphDiv = (typResistorInfoArray(1).dblInitTest_HighLimit * 1.5 - typResistorInfoArray(1).dblInitTest_LowLimit / 1.5) / 10
    '    '' '' ''                dblGraphTop = typResistorInfoArray(1).dblInitTest_HighLimit * 1.5
    '    '' '' ''            Else
    '    '' '' ''                dblGraphDiv = 0.3
    '    '' '' ''                dblGraphTop = 1.5
    '    '' '' ''            End If
    '    '' '' ''        Else
    '    '' '' ''            If ((0 >= typResistorInfoArray(1).dblFinalTest_LowLimit) And (0 <= typResistorInfoArray(1).dblFinalTest_HighLimit)) Then
    '    '' '' ''                dblGraphDiv = (typResistorInfoArray(1).dblFinalTest_HighLimit * 1.5 - typResistorInfoArray(1).dblFinalTest_LowLimit * 1.5) / 10
    '    '' '' ''                dblGraphTop = typResistorInfoArray(1).dblFinalTest_HighLimit * 1.5
    '    '' '' ''            ElseIf ((0 >= typResistorInfoArray(1).dblFinalTest_LowLimit) And (0 > typResistorInfoArray(1).dblFinalTest_HighLimit)) Then
    '    '' '' ''                dblGraphDiv = (typResistorInfoArray(1).dblFinalTest_HighLimit / 1.5 - typResistorInfoArray(1).dblFinalTest_LowLimit * 1.5) / 10
    '    '' '' ''                dblGraphTop = typResistorInfoArray(1).dblFinalTest_HighLimit * 1.5
    '    '' '' ''            ElseIf ((0 < typResistorInfoArray(1).dblFinalTest_LowLimit) And (0 <= typResistorInfoArray(1).dblFinalTest_HighLimit)) Then
    '    '' '' ''                dblGraphDiv = (typResistorInfoArray(1).dblFinalTest_HighLimit * 1.5 - typResistorInfoArray(1).dblFinalTest_LowLimit / 1.5) / 10
    '    '' '' ''                dblGraphTop = typResistorInfoArray(1).dblFinalTest_HighLimit * 1.5
    '    '' '' ''            Else
    '    '' '' ''                dblGraphDiv = 0.3
    '    '' '' ''                dblGraphTop = 1.5
    '    '' '' ''            End If
    '    '' '' ''        End If

    '    '' '' ''        gDistGrpPerLblAry(0).Text = "�`" & dblGraphTop.ToString("0.00")
    '    '' '' ''        For iCnt = 1 To 11
    '    '' '' ''            gDistGrpPerLblAry(iCnt).Text = (dblGraphTop - (dblGraphDiv * (iCnt - 1))).ToString("0.00") & "�`"
    '    '' '' ''        Next

    '    '' '' ''        picGraphAccumulationDrawSubLine()
    '    '' '' ''        picGraphAccumulationDrawLine(lScaleMax)
    '    '' '' ''        picGraphAccumulationPrintRegistNum()        ' ���z�O���t�ɒ�R����ݒ肷��
    '    '' '' ''    End Sub
    '    '' '' ''#End Region

    '#Region "���z�}�\���T�u"
    '    '''=========================================================================
    '    '''<summary>���z�}�\���T�u</summary>
    '    '''<remarks></remarks>
    '    '''=========================================================================
    '    Public Sub picGraphAccumulationDrawSubLine()
    '        'Dim i As Short

    '        '      'UPGRADE_ISSUE: PictureBox ���\�b�h picGraphAccumulation.Line �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"' ���N���b�N���Ă��������B
    '        'picGraphAccumulation.Line (56, 16) - (56, 112), RGB(0, 255, 0)
    '        '      'UPGRADE_ISSUE: PictureBox ���\�b�h picGraphAccumulation.Line �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"' ���N���b�N���Ă��������B
    '        'picGraphAccumulation.Line (56, 112) - (288, 112), RGB(0, 255, 0)
    '        '      For i = 0 To 10
    '        '          'UPGRADE_ISSUE: PictureBox ���\�b�h picGraphAccumulation.Line �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"' ���N���b�N���Ă��������B
    '        '	picGraphAccumulation.Line (56, 24 + (i * 8)) - (288, 24 + (i * 8)), RGB(0, 0, 128)
    '        '      Next
    '        '      'UPGRADE_ISSUE: PictureBox ���\�b�h picGraphAccumulation.Line �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"' ���N���b�N���Ă��������B
    '        'picGraphAccumulation.Line (172, 112) - (172, 116), RGB(0, 255, 0)

    '    End Sub
    '#End Region

    '    '' '' ''#Region "���z�}�\���T�u"
    '    '' '' ''    '''=========================================================================
    '    '' '' ''    '''<summary>���z�}�\���T�u</summary>
    '    '' '' ''    '''<remarks></remarks>
    '    '' '' ''    '''=========================================================================
    '    '' '' ''    Public Sub picGraphAccumulationDrawLine(ByRef lScaleMax As Integer)
    '    '' '' ''        Dim i As Short
    '    '' '' ''        Dim X As Short

    '    '' '' ''        For i = 0 To 11
    '    '' '' ''            X = CShort((glRegistNum(i) * 232) \ lScaleMax) ' ���z�O���t��R��
    '    '' '' ''            If (232 < X) Then
    '    '' '' ''                X = 232
    '    '' '' ''            End If
    '    '' '' ''            '         'UPGRADE_ISSUE: PictureBox ���\�b�h picGraphAccumulation.Line �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"' ���N���b�N���Ă��������B
    '    '' '' ''            'picGraphAccumulation.Line (56, 18 + (i * 8)) - (288, 22 + (i * 8)), RGB(0, 0, 0), BF
    '    '' '' ''            '         'UPGRADE_ISSUE: PictureBox ���\�b�h picGraphAccumulation.Line �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"' ���N���b�N���Ă��������B
    '    '' '' ''            'picGraphAccumulation.Line (56, 18 + (i * 8)) - (56 + X, 22 + (i * 8)), RGB(0, 255, 255), BF
    '    '' '' ''        Next
    '    '' '' ''        Form1.lblRegistUnit.Text = CStr(lScaleMax \ 2)
    '    '' '' ''    End Sub
    '    '' '' ''#End Region

    '    '' '' ''#Region "���z�O���t�ɒ�R����ݒ肷��"
    '    '' '' ''    '''=========================================================================
    '    '' '' ''    '''<summary>���z�O���t�ɒ�R����ݒ肷��</summary>
    '    '' '' ''    '''<remarks></remarks>
    '    '' '' ''    '''=========================================================================
    '    '' '' ''    Public Sub picGraphAccumulationPrintRegistNum()
    '    '' '' ''        Dim i As Short

    '    '' '' ''        For i = 0 To (MAX_SCALE_RNUM - 1)
    '    '' '' ''            gDistRegNumLblAry(i).Text = CStr(glRegistNum(i))  ' ���z�O���t��R��
    '    '' '' ''        Next

    '    '' '' ''    End Sub
    '    '' '' ''#End Region

    '    '' '' ''#Region "�O���t�N���b�N������"
    '    '' '' ''    '''=========================================================================
    '    '' '' ''    '''<summary>�O���t�N���b�N������</summary>
    '    '' '' ''    '''<remarks></remarks>
    '    '' '' ''    '''=========================================================================
    '    '' '' ''    Public Sub lblGraphClick_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblGraphClick.Click
    '    '' '' ''        ' �O���t�N���b�N
    '    '' '' ''        frmDistribution.Show()
    '    '' '' ''    End Sub
    '    '' '' ''#End Region


    End Module