'===============================================================================
'   Description  : DllTrimFnc.dll�֐��̒�`
'
'   Copyright(C) : TOWA LASERFRONT CORP. 2010
'
'===============================================================================
Option Strict Off
Option Explicit On

Imports System.Runtime.InteropServices

Module DefTrimFnc

#Region "�萔/�ϐ���`"
    '===========================================================================
    '   �萔/�ϐ���`
    '===========================================================================
    Public Const cMAXcMARKINGcSTRLEN As Integer = 18        ' �}�[�L���O������ő咷(byte)
    Public Const cMAXcSENDcPRMCNT As Integer = 32           ' VB��INTRIM�̑��M�R�}���h�p�����[�^�ő吔
    'Public Const cResultMax As Integer = 256                ' �g���~���O���ʃf�[�^�̍ő�z��
    'Public Const cResultAry As Integer = 999                ' �g���~���O���ʃf�[�^�̍ő吔
    Public Const cAxisMAX As Integer = 5                    ' �ő厲��(X,Y,Z,Theta,Z2)
    Public Const cRsultTky As Integer = 4                   ' TKY�߂�l
    Public Const cRetAxisPos As Integer = 5               ' �e���̌��ݒl(X,Y,Z,Theta,Z2)
    Public Const cRetBpPos As Integer = 2                      ' BP�̌��ݒl
    Public Const cMAXcRESISTORS As Integer = 256            ' �ő��R��
#End Region

#Region "�g���~���O�f�[�^�\���̌`����`(VB����INtime)"
    '===========================================================================
    '   �g���~���O�f�[�^�\���̌`����`(VB����INtime)
    '===========================================================================
    '---------------------------------------------------------------------------
    '   ���[�U�v���O���������J�b�g���ʃp�����[�^(VB��INtime)
    '---------------------------------------------------------------------------
    '------ �e�J�b�g�̃p�����[�^(�����A�X�s�[�h�AQ���[�g)�ۑ��̈�
    Public Structure DBL_CUTCOND_ARRAY
        Dim dblL1 As Double                                    ' Line1�p�̃p�����[�^
        Dim dblL2 As Double                                    ' Line2�p�̃p�����[�^
        Dim dblL3 As Double                                    ' Line3�p�̃p�����[�^
        Dim dblL4 As Double                                    ' Line4�p�̃p�����[�^
    End Structure
    Public Structure SRT_CUTCOND_ARRAY
        Dim srtL1 As Short                                     ' Line1�p�̃p�����[�^
        Dim srtL2 As Short                                    ' Line2�p�̃p�����[�^
        Dim srtL3 As Short                                    ' Line3�p�̃p�����[�^
        Dim srtL4 As Short                                    ' Line4�p�̃p�����[�^
    End Structure
    '------ ���H�ݒ�\����
    Public Structure CUT_COND_STRUCT
        Dim CutLen As DBL_CUTCOND_ARRAY                     '�J�b�g�����
        Dim SpdOwd As DBL_CUTCOND_ARRAY                     '�J�b�g�X�s�[�h�i���H�j
        Dim SpdRet As DBL_CUTCOND_ARRAY                     '�J�b�g�X�s�[�h�i���H�j
        Dim QRateOwd As DBL_CUTCOND_ARRAY                   '�J�b�gQ���[�g�i���H�j
        Dim QRateRet As DBL_CUTCOND_ARRAY                   '�J�b�gQ���[�g�i���H�j
        Dim CondOwd As SRT_CUTCOND_ARRAY                    '�J�b�g�����ԍ��i���H�j
        Dim CondRet As SRT_CUTCOND_ARRAY                    '�J�b�g�����ԍ��i���H�j
    End Structure

    'V1.0.4.3�F��
    '------ ���H�ݒ�\���� �U�_�^�[���|�C���g�k�J�b�g�p
    Public Structure CUT_COND_STRUCT_L6
        Dim dCutLen_1 As Double    ' �J�b�g���P�`�V
        Dim dCutLen_2 As Double    ' �J�b�g���P�`�V
        Dim dCutLen_3 As Double    ' �J�b�g���P�`�V
        Dim dCutLen_4 As Double    ' �J�b�g���P�`�V
        Dim dCutLen_5 As Double    ' �J�b�g���P�`�V
        Dim dCutLen_6 As Double    ' �J�b�g���P�`�V
        Dim dCutLen_7 As Double    ' �J�b�g���P�`�V
        Dim dQRate_1 As Double     ' �p���[�g�P�`�V
        Dim dQRate_2 As Double     ' �p���[�g�P�`�V
        Dim dQRate_3 As Double     ' �p���[�g�P�`�V
        Dim dQRate_4 As Double     ' �p���[�g�P�`�V
        Dim dQRate_5 As Double     ' �p���[�g�P�`�V
        Dim dQRate_6 As Double     ' �p���[�g�P�`�V
        Dim dQRate_7 As Double     ' �p���[�g�P�`�V
        Dim dSpeed_1 As Double     ' ���x�P�`�V
        Dim dSpeed_2 As Double     ' ���x�P�`�V
        Dim dSpeed_3 As Double     ' ���x�P�`�V
        Dim dSpeed_4 As Double     ' ���x�P�`�V
        Dim dSpeed_5 As Double     ' ���x�P�`�V
        Dim dSpeed_6 As Double     ' ���x�P�`�V
        Dim dSpeed_7 As Double     ' ���x�P�`�V
        Dim dAngle_1 As Double     ' �p�x�P�`�V
        Dim dAngle_2 As Double     ' �p�x�P�`�V
        Dim dAngle_3 As Double     ' �p�x�P�`�V
        Dim dAngle_4 As Double     ' �p�x�P�`�V
        Dim dAngle_5 As Double     ' �p�x�P�`�V
        Dim dAngle_6 As Double     ' �p�x�P�`�V
        Dim dAngle_7 As Double     ' �p�x�P�`�V
        Dim dTurnPoint_1 As Double ' �^�[���|�C���g�P�`�U
        Dim dTurnPoint_2 As Double ' �^�[���|�C���g�P�`�U
        Dim dTurnPoint_3 As Double ' �^�[���|�C���g�P�`�U
        Dim dTurnPoint_4 As Double ' �^�[���|�C���g�P�`�U
        Dim dTurnPoint_5 As Double ' �^�[���|�C���g�P�`�U
        Dim dTurnPoint_6 As Double ' �^�[���|�C���g�P�`�U
    End Structure
    'V1.0.4.3�F��

    '------ �J�b�g���\����
    Public Structure CUT_INFO_STRUCT
        Dim srtMoveMode As Short                                '���샂�[�h�i0:�g���~���O�A1:�e�B�[�`���O�A2:�����J�b�g�j
        Dim srtCutMode As Short                                 '�J�b�g���[�h(0:�m�[�}���A1:���^�[���A2:���g���[�X�A3:�΂߁j
        Dim dblTarget As Double                                 '�ڕW�l
        Dim srtSlope As Short                                   '�X���[�v�ݒ�(1:�d������{�X���[�v�A2:�d������|�X���[�v�A4:��R����{�X���[�v�A5:��R����|�X���[�v�j
        Dim srtMeasType As Short                                '����^�C�v(0:����(3��)�A1:�����x(2000��)�A2:�iIDX�̂݁j�O���@��A3:���薳���A5�`:�w��񐔑���j
        Dim dblAngle As Double                                  '�J�b�g�p�x
        Dim dblLTP As Double                                    'L�^�[���|�C���g
        Dim srtLTDIR As Short                                   'L�^�[����̕���
        Dim dblRADI As Double                                   'R����]���a�iU�J�b�g�Ŏg�p�j
        'for Hook and U
        Dim dblRADI2 As Double                                  'R2����]���a�iU�J�b�g�Ŏg�p�j  
        Dim srtHkOrUType As Short                               'HookCut(3)��U�J�b�g�i3�ȊO�j�̎w��B
        'for Index
        Dim srtIdxScnCnt As Short                               '�C���f�b�N�X/�X�L�����J�b�g��(1�`32767)
        Dim srtIdxMeasMode As Short                             '�C���f�b�N�X���胂�[�h�i0:��R�A1:�d���A2:�O���j
        'for EdgeSense
        Dim dblEsPoint As Double                                '�G�b�W�Z���X�|�C���g
        Dim dblRdrJdgVal As Double                              '���_�[��������ω���
        Dim dblMinJdgVal As Double                              '���_�[�J�b�g��Œዖ�e�ω���
        Dim srtEsAftCutCnt As Short                             '���_�[�ؔ�����̃J�b�g�񐔁i����񐔁j
        Dim srtMinOvrNgCnt As Short                             '���_�[���o����A�Œ�ω��ʂ̘A��Over���e��
        Dim srtMinOvrNgMode As Short                            '�A��Over����NG�����i0:NG���薢���{, 1:NG������{�B���_�[���؂�, 2:NG���薢���{�B���_�[�؏グ�j	
        'for Scan
        Dim dblStepPitch As Double                              '�X�e�b�v�ړ��s�b�`
        Dim srtStepDir As Short                                 '�X�e�b�v����
    End Structure
    '-----------------------------------------------------------------
    '���[�U�v���O���������J�b�g�p�����[�^ �U�_�^�[���|�C���g�k�J�b�g�p
    '-----------------------------------------------------------------
    Public Structure CUT_COMMON_PRM
        Dim CutInfo As CUT_INFO_STRUCT
        Dim CutCond As CUT_COND_STRUCT
    End Structure

    'V1.0.4.3�F��
    '---------------------------------------
    '���[�U�v���O���������J�b�g�p�����[�^
    '---------------------------------------
    Public Structure CUT_COMMON_PRM_L6
        Dim CutInfo As CUT_INFO_STRUCT
        Dim CutCond As CUT_COND_STRUCT_L6
    End Structure
    'V1.0.4.3�F��

    '---------------------------------------
    'iTKY�����J�b�g�p�����[�^ - C���ꑤ�ł�Union���g�p
    '---------------------------------------
    '---------------------------------------------------------------------------
    '   �J�b�g�^�C�v�ʃp�����[�^�`����`(VB��INtime)
    '---------------------------------------------------------------------------
    '----- ST cut -----
    Public Structure PRM_CUT_ST                             ' ST cut�p�����[�^�`����`
        'Dim DIR As UShort                                   ' �J�b�g����(1:+X, 2:-X, 3:+Y, 4:-Y)
        Dim MODE As UShort                                  ' ���샂�[�h(0:NOM, 1:���^�[��, 2:���g���[�X, 3:�΂�)
        Dim angle As UShort                                 ' �΂߃J�b�g�p�x(0�`359)
        Dim Length As Double                                ' �ő�J�b�e�B���O��(0.0001�`20.0000(mm))
    End Structure

    '----- L cut -----
    Public Structure PRM_CUT_L                              ' L cut�p�����[�^�`����`
        'Dim DIR As UShort                                   ' �J�b�g����(1:+X, 2:-X, 3:+Y, 4:-Y)
        'Dim tdir As UShort                                  ' L�^�[������(1:+X, 2:-X, 3:+Y, 4:-Y)
        Dim MODE As UShort                                  ' ���샂�[�h(0:NOM, 1:���^�[��, 2:���g���[�X, 3:�΂�)
        Dim angle As UShort                                 ' �΂߃J�b�g�p�x(0�`359)
        Dim tdir As UShort                                  ' L�^�[������(1:+X, 2:-X, 3:+Y, 4:-Y)
        Dim turn As Double                                  ' L�^�[���|�C���g(0.0�`100.0(%))
        Dim L1 As Double                                    ' L1 �ő�J�b�e�B���O��(0.0001�`20.0000(mm))
        Dim L2 As Double                                    ' L2 �ő�J�b�e�B���O��(0.0001�`20.0000(mm))
        Dim r As Double                                     ' �^�[���̉~�ʔ��a(mm)
    End Structure

    '----- HOOK cut -----
    Public Structure PRM_CUT_HOOK                           ' HOOK cut�p�����[�^�`����`
        'Dim DIR As UShort                                   ' �J�b�g����(1:+X, 2:-X, 3:+Y, 4:-Y)
        Dim MODE As UShort                                  ' ���샂�[�h(0:NOM, 1:���^�[��, 2:���g���[�X, 3:�΂�)
        Dim angle As UShort                                 ' �΂߃J�b�g�p�x(0�`359)
        Dim tdir As UShort                                  ' L�^�[������(1:+X, 2:-X, 3:+Y, 4:-Y)
        Dim turn As Double                                  ' L�^�[���|�C���g(0.0�`100.0(%))
        Dim L1 As Double                                    ' L1 �ő�J�b�e�B���O��(0.0001�`20.0000(mm))
        Dim r1 As Double                                    ' �^�[��1�̉~�ʔ��a(mm)
        Dim L2 As Double                                    ' L2 �ő�J�b�e�B���O��(0.0001�`20.0000(mm))
        Dim r2 As Double                                    ' �^�[��2�̉~�ʔ��a(mm)
        Dim L3 As Double                                    ' L3 �ő�J�b�e�B���O��(0.00001�`20.0000(mm))
    End Structure

    '----- INDEX cut -----
    Public Structure PRM_CUT_INDEX                          ' INDEX cut�p�����[�^�`����`
        Dim angle As UShort                                   ' �J�b�g�p�x(0�`359)
        Dim maxindex As UShort                              ' �C���f�b�N�X��(1�`32767)
        Dim measMode As UShort                              ' ���胂�[�h(0:��R, 1:�d��)
        Dim measType As UShort                              ' ���蔻��^�C�v(0:����, 1:�����x, 2:�O��) 
        Dim Length As Double                                ' �C���f�b�N�X��(0.0001�`20.0000(mm))
    End Structure

    '----- SCAN cut -----
    Public Structure PRM_CUT_SCAN                           ' SCAN cut�p�����[�^�`����`
        Dim angle As UShort                                 ' �J�b�g�p�x(0�`359)
        Dim sdir As UShort                                  ' �X�e�b�v����(1:+X, 2:-X, 3:+Y, 4:-Y)
        Dim lines As UShort                                 ' �{��(1�`n)
        Dim Length As Double                                ' �J�b�e�B���O��(0.0001�`20.0000(mm))
        Dim pitch As Double                                 ' �s�b�`(0.0001�`20.0000(mm))
    End Structure

    '----- Letter Marking -----
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure PRM_CUT_MARKING                        ' Letter Marking�p�����[�^�`����`
        '                                                   ' ����
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=cMAXcMARKINGcSTRLEN)> _
            Dim str() As Byte
        Dim magnify As Double                               ' �{��(�P�`999)
        Dim DIR As UShort                                   ' �����̌���(1:0, 2:90, 3:180, 4:270)
        ' ���̍\���̂�����������ɂ́A"Initialize" ���Ăяo���Ȃ���΂Ȃ�܂���B 
        Public Sub Initialize()
            ReDim str(cMAXcMARKINGcSTRLEN - 1)
        End Sub
    End Structure

    '----- ES cut -----
    Public Structure PRM_CUT_ES                             ' ES cut�p�����[�^�`����`
        Dim angle As UShort                                 ' �J�b�g�p�x(0�`359)
        Dim L1 As Double                                    ' L1 �ő�J�b�e�B���O��(0.0001�`20.0000(mm))
        Dim EsPoint As Double                               ' ES�߲��(-99.9999�`0.0000%)
        Dim ESWide As Double                                ' ES����ω���(0.0�`100.0%)
        Dim ESWide2 As Double                               ' ES��ω���(0.0�`100.0%)
        Dim EScount As UShort                               ' ES��m�F��(0�`20)
        Dim CTcount As UShort                               ' ���޾ݽ��A��NG�m�F�񐔁@
        Dim wJudgeNg As UShort                              ' NG���肷��/���Ȃ��i0:TRUE/1:FALSE�j
    End Structure

    '----- UCUT�p�����[�^(1�v�f) -----
    Public Structure UCUT_PARAM_EL                          ' UCUT�p�����[�^(1�v�f)�`����`
        Dim RATIO As Double                                 ' �ڕW�l�ɑ΂��鏉���l�̍�(%)
        Dim LTP As Double                                   ' L�^�[���|�C���g(0.0�`100.0%)
        Dim LTP2 As Double                                  ' L�^�[���|�C���g2(0.0�`100.0%)
        Dim L1 As Double                                    ' L1 �ő�J�b�e�B���O��(0.0001�`20.0000mm)
        Dim L2 As Double                                    ' L2 �ő�J�b�e�B���O��(0.0001�`20.0000mm)
        Dim r As Double                                     ' �~�ʔ��a (mm)
        Dim V As Double                                     ' ���x(mm/s)
        Dim NOM As Double                                   ' �ڕW�l
        Dim Flg As Boolean                                  ' �f�[�^�L��(���g�p)
    End Structure

    '----- UCUT�p�����[�^ -----
    Public Structure S_UCUTPARAM_EL                         ' UCUT�p�����[�^�`����`
        Dim RNO As UShort
        Dim NOM As Double
        Dim PRM_UNIT As UCUT_PARAM_EL
    End Structure

    '----- UCUT�p�����[�^�e�[�u��(1��R��) -----
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure S_UCUTPARAM                            ' UCUT�p�����[�^�`����`
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=20)> _
            Dim EL() As S_UCUTPARAM_EL       ' UCUT�p�����[�^ 

        ' ���̍\���̂�����������ɂ́A"Initialize" ���Ăяo���Ȃ���΂Ȃ�܂���B 
        Public Sub Initialize()
            ReDim EL(19)
        End Sub
    End Structure
    Public Structure S_CUT_CONDITION                         ' ���H�����ݒ�\����
        Dim cutSetNo As UShort                              ' (FL����)���H�����ԍ�
        Dim cutSpd As Double                                ' �J�b�g�X�s�[�h
        Dim cutQRate As Double                              ' �J�b�gQ���[�g
        Dim bUse As Boolean                                 ' INTRIM���ł̂ݎg�p
    End Structure

    '-------------------------------------------------------------------------------
    '   �J�b�g�f�[�^�`����`(VB��INtime)
    '-------------------------------------------------------------------------------
    <StructLayout(LayoutKind.Sequential)> _
        Public Structure TRIM_CUT_DATA          ' �J�b�g�f�[�^�`����`
        Dim wCutNo As UShort                        ' �J�b�g�ԍ� 1-20
        Dim wDelayTime As UShort                    ' ��d������㑪��x������(0-32767msec) 
        Dim wCutType As UShort                      ' �J�b�g�`��(1:st, 2:L, 3:HK, 4:IX ��)
        Dim wMoveMode As UShort                     ' �J�b�g���[�h�F�g���~���O�J�b�g(0)�A�����J�b�g(2)
        Dim wDoPosition As UShort                   ' �|�W�V���j���O(0:����, 1:�Ȃ�)
        Dim fCutStartX As Double                    ' �J�b�g�X�^�[�g���WX(-80.0000�`+80.0000)
        Dim fCutStartY As Double                    ' �J�b�g�X�^�[�g���WY(-80.0000�`+80.0000)
        'Dim CP5 As Double                          ' [@@@�폜]�J�b�g�X�s�[�h(0.1�`409.0mm/s)
        'Dim CP6 As Double                          ' [@@@�폜]���[�U�[Q�X�C�b�`���[�g(0.1�`50.0KHz) ��FL���͉��H�����ԍ�1��QڰĂ�ݒ�
        Dim fCutOff As Double                       ' �J�b�g�I�t %(-99.999 �` +999.999)
        Dim fAveDataRate As Double                  ' �J�b�g�f�[�^���ω���(0.0�`100.0, 0%)(���g�p)
        Dim bUcutNo As Byte                         ' [@@@�ǉ�]U�J�b�g �p�����[�^�w�莞�A�I�����ꂽ�e�[�u���ԍ���ۑ�(INTRIM���ł̂ݎg�p�j
        Dim fInitialVal As Double                   ' [@@@�ǉ�]U�J�b�g �J�b�g���Ƃ̏����l������ۑ�(INTRIM���ł̂ݎg�p�j
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=cCNDNUM)> _
             Dim CutCnd() As S_CUT_CONDITION '���H�����i�����ԍ��A�J�b�g�X�s�[�h�AQrate�j1�`8
        '        <VBFixedArray(cCNDNUM - 1)> Dim CP72() As Byte      ' ���H�����ԍ�1�`4(FL�p) 
        'Dim dummy As PRM_CUT_HOOK                          ' �J�b�g�p�����[�^(union) ��INTRIM���ł̂ݎg�p

        ' ���̍\���̂�����������ɂ́A"Initialize" ���Ăяo���Ȃ���΂Ȃ�܂���B 
        Public Sub Initialize()
            ReDim CutCnd(cCNDNUM - 1)
        End Sub
    End Structure

    '-------------------------------------------------------------------------------
    '   ��R�f�[�^�`����`(VB��INtime)
    '-------------------------------------------------------------------------------
    Public Structure TRIM_RESISTOR_DATA     ' ��R�f�[�^�`����`
        Dim wResNo As UShort                    ' ��R�ԍ�(1-999=�g���~���O, 1000-9999=�}�[�L���O)
        Dim wMeasMode As UShort                 ' [@@@�ǉ�]���胂�[�h�@0�F��R�A1�F�d��
        Dim wMeasType As UShort                 ' ���葪��(0:����, 1:�����x, 2:�O���j
        Dim wCircuit As UShort                  ' �T�[�L�b�g(��R��������T�[�L�b�g�ԍ�)
        Dim wHiProbNo As UShort                 ' �n�C���v���[�u�ԍ�
        Dim wLoProbNo As UShort                 ' ���[���v���[�u�ԍ�
        Dim w1stAG As UShort                    ' ��1�A�N�e�B�u�K�[�h�ԍ�
        Dim w2ndAG As UShort                    ' ��2�A�N�e�B�u�K�[�h�ԍ�
        Dim w3rdAG As UShort                    ' ��3�A�N�e�B�u�K�[�h�ԍ�
        Dim w4thAG As UShort                    ' ��4�A�N�e�B�u�K�[�h�ԍ�
        Dim w5thAG As UShort                    ' ��5�A�N�e�B�u�K�[�h�ԍ�
        Dim dwExBits As UInteger                ' External bits
        Dim wPauseTime As UShort                ' �|�[�Y�^�C��(External bits�o�͌�̃E�F�C�g) (msec)
        Dim wRatioMode As UShort                ' �ڕW�l�w��(0:��Βl, 1:���V�I, 2:�v�Z��)
        Dim wBaseReg As UShort                  ' �x�[�X��RNo.(���V�I���̊��R�ԍ�)
        Dim fTargetVal As Double                ' �g���~���O�ڕW�l(ohm)
        Dim wSlope As UShort                    ' �d���ω��X���[�v(0:+�X���[�v, 1:-�X���[�v) ����ڰ��ް��̑��胂�[�h=�d���̏ꍇ�L��
        Dim fITLimitH As Double                 ' IT Limit H(-99.99�`9999.99%)
        Dim fITLimitL As Double                 ' IT Limit L(-99.99�`9999.99%)
        Dim fFTLimitH As Double                 ' FT Limit H(-99.99�`9999.99%)
        Dim fFTLimitL As Double                 ' FT Limit L(-99.99�`9999.99%)
        Dim wCutCnt As UShort                   ' �J�b�g��(1�`20)
        Dim wCorrectFlg As UShort               ' �J�b�g�ʒu�␳�t���O(0:�␳���Ȃ�, 1:�␳����)
        'Dim PR14_Ha As Double                  ' [@@@�폜]�C�j�V����OK�e�X�gHIGH���~�b�g(SL436K�p)
        'Dim PR14_La As Double                  ' [@@@�폜]�C�j�V����OK�e�X�gLOW���~�b�g (SL436K�p)
        Dim fCutMag As Double                   ' �؏グ�{��(CHIP�̂�)
        Dim bTrimEnd As Boolean                 ' �g���~���O�����t���O�iINTRIM���Ŏg�p�j
        Dim pCutData As UInteger                ' �J�b�g�f�[�^�|�C���^(INTRIM���Ŏg�p)
    End Structure

    '-------------------------------------------------------------------------------
    '   �v���[�g�f�[�^�`����`(VB��INtime)
    '-------------------------------------------------------------------------------
    Public Structure TRIM_PLATE_DATA                        ' �v���[�g�f�[�^�`����`
        Dim wCircuitCnt As UShort                           ' �T�[�L�b�g��
        Dim wRegistCnt As UShort                            ' ��R��
        Dim wResCntInCrt As UShort                          ' �T�[�L�b�g����R��
        'Dim wTrimMode As UShort                             ' ���胂�[�h(0:��R, 1:�d��)
        Dim wDelayTrim As UShort                            ' �f�B���C�g����(0=�Ȃ�, 1=�ިڲ��т����s����, 2=�ިڲ���2�����s����)
        Dim fBPOffsetX As Double                            ' BP�I�t�Z�b�gX(mm)
        Dim fBPOffsetY As Double                            ' BP�I�t�Z�b�gY(mm)
        Dim fAdjustOffsetX As Double                        ' �A�W���X�g�ʒuX(mm)
        Dim fAdjustOffsetY As Double                        ' �A�W���X�g�ʒuY(mm)
        Dim fNgCriterion As Double                          ' NG����(%)
        Dim fZStepPos As Double                             ' Z���ï��&��߰Ĉʒu
        Dim fZTrimPos As Double                             ' Z������Ĉʒu
        Dim fReProbingX As Double                           ' ����۰��ݸ�X�ړ���
        Dim fReProbingY As Double                           ' ����۰��ݸ�Y�ړ���
        Dim wReProbingCnt As UShort                         ' ����۰��ݸމ�
        Dim wInitialOK As UShort                            ' IT��FT�͈͎��̏���(0:�J�b�g���s 1:�J�b�g�����s))
        Dim wNGMark As UShort                               ' NGϰ�ݸނ���/���Ȃ�)(SL436K�p)
        Dim w4Terminal As UShort                            ' 4�[�q�������������/���Ȃ�)(SL436K�p)
        Dim wLogMode As UShort                              ' ۷�ݸ�Ӱ��
        '                                                   ' 0:���Ȃ�, 1:INITIAL TEST, 2:FINAL TEST, 3:INITIAL + FINAL)	
        Dim bTrimCutEnd As Boolean                          ' �J�b�g�I�t�ڕW�ő�l�ɓ��B������J�b�g���I������iTRUE�j/���Ȃ��iFALSE�j
    End Structure

    '-------------------------------------------------------------------------------
    '   GPIB�ݒ�p�f�[�^�`����`(VB��INtime) ��CHIP�p
    '-------------------------------------------------------------------------------
    '<StructLayout(LayoutKind.Sequential)> _
    'Public Structure TRIM_PLATE_GPIB                        ' GPIB�ݒ�p�f�[�^�`����`
    '    Dim wGPIBmode As UShort                             ' GP-IB����(0:���Ȃ� 1:����)
    '    Dim wDelim As UShort                                ' �����(0:CR+LF 1:CR 2:LF 3:NONE)
    '    Dim wTimeout As UShort                              ' ��ѱ��(0�`1000)(100ms�P��)
    '    Dim wAddress As UShort                              ' �@����ڽ(0�`30)
    '    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=40)> _
    '        Dim strI() As Byte           ' �����������(MAX40byte)
    '    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=10)> _
    '        Dim strT() As Byte           ' �ض޺����(10byte)
    '    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=6)> _
    '        Dim wReserve() As Byte        ' �\��(6byte)  
    '    Dim wMeasurementMode As UShort                      ' ���胂�[�h(0:���, 1:�΍�) 

    '    ' ���̍\���̂�����������ɂ́A"Initialize" ���Ăяo���Ȃ���΂Ȃ�܂���B 
    '    Public Sub Initialize()
    '        ReDim strI(39)
    '        ReDim strT(9)
    '        ReDim wReserve(5)
    '    End Sub
    'End Structure
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure TRIM_PLATE_GPIB                        ' GPIB�ݒ�p�f�[�^�`����` '###002
        Dim wGPIBmode As UShort                             ' GP-IB����(0:���Ȃ� 1:����)
        Dim wDelim As UShort                                ' �����(0:CR+LF 1:CR 2:LF 3:NONE)
        Dim wTimeout As UShort                              ' ��ѱ��(0�`1000)(100ms�P��)
        Dim wAddress As UShort                              ' �@����ڽ(0�`30)
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=40)> _
            Dim strI() As Byte           ' �����������(MAX40byte)
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=10)> _
            Dim strT() As Byte           ' �ض޺����(10byte)
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=6)> _
            Dim wReserve() As Byte        ' �\��(6byte)  
        Dim wMeasurementMode As UShort                      ' ���胂�[�h(0:���, 1:�΍�) 

        ' ���̍\���̂�����������ɂ́A"Initialize" ���Ăяo���Ȃ���΂Ȃ�܂���B 
        Public Sub Initialize()
            ReDim strI(39)
            ReDim strT(9)
            ReDim wReserve(5)
        End Sub
    End Structure

    '-------------------------------------------------------------------------------
    '   �g���~���O�v���f�[�^�`����`(VB��INtime)
    '-------------------------------------------------------------------------------
    '----- �g���~���O�v���f�[�^(GPIB�ݒ�p�f�[�^) -----
    Public Structure TRIM_DAT_GPIB                          ' �g���~���O�v���f�[�^(GPIB�ݒ�p�f�[�^)�`����`
        Dim cmdNo As UInteger                               ' �R�}���hNo.(DllTrimFnc�Őݒ肷��̂Ŗ��g�p)
        Dim type As UShort                                  ' �f�[�^�^�C�v(8:GPIB�f�[�^)
        Dim index_reg As UShort                             ' ��R�f�[�^�E�C���f�b�N�X
        Dim index_cut As UShort                             ' �J�b�g�f�[�^�E�C���f�b�N�X
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET���(0:TKY, 1:CHIP, 2:NET)
        Dim prmGPIB As TRIM_PLATE_GPIB                      ' GPIB�ݒ�p�f�[�^
    End Structure

    '-------------------------------------------------------------------------------
    '   �g���~���O�v���f�[�^�`����`(VB��DllTrimFunc)
    '-------------------------------------------------------------------------------
    '----- �J�b�g�ʒu�␳�\����(�z���0�I���W��) ###059 -----
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure CUTPOS_CORRECT_DATA                              ' �v���f�[�^�`����`
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=cMAXcRESISTORS)> _
            Dim corrPosX() As Double                            ' double �^�p�����[�^(X���W�␳�l)
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=cMAXcRESISTORS)> _
            Dim corrPosY() As Double                          ' double �^�p�����[�^(X���W�␳�l)
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=cMAXcRESISTORS)> _
        Dim corrResult() As UInteger

        ' ���̍\���̂�����������ɂ́A"Initialize" ���Ăяo���Ȃ���΂Ȃ�܂���B 
        Public Sub Initialize()
            ReDim corrPosX(cMAXcRESISTORS)
            ReDim corrPosY(cMAXcRESISTORS)
            ReDim corrResult(cMAXcRESISTORS)
        End Sub
    End Structure

    '<StructLayout(LayoutKind.Sequential)> _
    'Public Structure CUTPOS_CORRECT_DATA                              ' �v���f�[�^�`����`
    '    '        Dim cmdNo As UInteger                               ' �R�}���hNo.(DllTrimFnc�Őݒ肷��̂Ŗ��g�p)
    '    <MarshalAs(UnmanagedType.ByValArray, SizeConst:=cMAXcRESISTORS + 1)> _
    '        Dim corrPosX() As Double                            ' double �^�p�����[�^(X���W�␳�l)
    '    <MarshalAs(UnmanagedType.ByValArray, SizeConst:=cMAXcRESISTORS + 1)> _
    '        Dim corrPosY() As Double                          ' double �^�p�����[�^(X���W�␳�l)
    '    <MarshalAs(UnmanagedType.ByValArray, SizeConst:=cMAXcRESISTORS + 1)> _
    '    'Dim corrResult() As UShort                            

    '    ' ���̍\���̂�����������ɂ́A"Initialize" ���Ăяo���Ȃ���΂Ȃ�܂���B 
    '    Public Sub Initialize()
    '        ReDim corrPosX(cMAXcRESISTORS)
    '        ReDim corrPosY(cMAXcRESISTORS)
    '        ReDim corrResult(cMAXcRESISTORS)
    '    End Sub
    'End Structure

    '-------------------------------------------------------------------------------
    '   �v��/�����f�[�^(�R�}���h)�`����`(VB����INtime)
    '-------------------------------------------------------------------------------
    '----- �v���f�[�^(VB��INtime) -----
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure S_CMD_DAT                              ' �v���f�[�^�`����`
        Dim cmdNo As UInteger                               ' �R�}���hNo.(DllTrimFnc�Őݒ肷��̂Ŗ��g�p)
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=cMAXcSENDcPRMCNT)> _
            Dim dbPara() As Double                          ' double �^�p�����[�^(dbPara(0-32))
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=cMAXcSENDcPRMCNT)> _
            Dim dwPara() As Integer                         ' long	 �^�p�����[�^(dbPara(0-32))
        Dim flgTrim As UInteger                             ' 0:���ݸޒ��łȂ�, 1:���ݸޒ�(IRQ0�����֎~) 

        ' ���̍\���̂�����������ɂ́A"Initialize" ���Ăяo���Ȃ���΂Ȃ�܂���B 
        Public Sub Initialize()
            ReDim dbPara(cMAXcSENDcPRMCNT - 1)
            ReDim dwPara(cMAXcSENDcPRMCNT - 1)
        End Sub
    End Structure

    '----- ���V�I���[�h�Q�v�Z���f�[�^(VB��INtime) -----
    Public Structure S_RATIO2EXP                            ' ���V�I���[�h�Q�v�Z���f�[�^�`����`
        Dim cmdNo As UInteger                               ' �R�}���hNo.(DllTrimFnc�Őݒ肷��̂Ŗ��g�p)
        Dim RNO As UInteger                                 ' ��R�ԍ�
        Dim strExp As String                                ' �v�Z��������
    End Structure

    '----- �����f�[�^(�R�}���h)(VB��INtime) -----
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure S_RES_DAT                              ' �����f�[�^�`����`
        Dim status As UInteger                              ' 0:����, 0�ȊO:�s����
        Dim dwerrno As UInteger                             ' �G���[�ԍ�(0:����)
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=cAxisMAX)> _
            Dim signal() As UInteger                        ' ���X�e�[�^�X
        '                                                   ' [0]:X��
        '                                                   ' [1]:Y��
        '                                                   ' [2]:Z��
        '                                                   ' [3]:�Ǝ�
        '                                                   ' I/O���͏��
        ' '' '' ''<MarshalAs(UnmanagedType.ByValArray, SizeConst:=INP_MAX)> _
        ' '' '' ''    Dim in_dat() As UInteger
        '' '' '' ''                                                   ' [0]:�R���\�[��SW�Z���X
        '' '' '' ''                                                   ' [1]:�C���^�[���b�N�֌WSW�Z���X
        '' '' '' ''                                                   ' [2]:�I�[�g���[�_LO
        '' '' '' ''                                                   ' [3]:�I�[�g���[�_HI
        '' '' '' ''                                                   ' [4]:�Œ�A�b�e�l�[�^
        '' '' '' ''                                                   ' I/O�o�͏��
        ' '' '' ''<MarshalAs(UnmanagedType.ByValArray, SizeConst:=OUT_MAX)> _
        ' '' '' ''    Dim outdat() As UInteger
        '' '' '' ''                                                   ' [0]:�R���\�[������
        '' '' '' ''                                                   ' [1]:�T�[�{�p���[
        '' '' '' ''                                                   ' [2]:�I�[�g���[�_LO
        '' '' '' ''                                                   ' [3]:�I�[�g���[�_HI
        '' '' '' ''                                                   ' [4]:�V�O�i���^���[(���g�p)
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=cRsultTky)> _
            Dim wData() As UInteger                         ' TKY�ߒl

        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=cRetAxisPos)> _
            Dim pos() As Double                             ' ���݈ʒu
        '                                                   ' [0]:X��
        '                                                   ' [1]:Y��
        '                                                   ' [2]:Z��
        '                                                   ' [3]:Theta
        '                                                   ' [4]:Z2
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=cRetBpPos)> _
            Dim bppos() As Double
        '                                                   ' [0]:BPX
        '                                                   ' [1]:BPY
        Dim fData As Double                                 ' �ߒl(����l��)

        ' ���̍\���̂�����������ɂ́A"Initialize" ���Ăяo���Ȃ���΂Ȃ�܂���B 
        Public Sub Initialize()
            ReDim signal(cAxisMAX - 1)
            '' '' ''ReDim in_dat(INP_MAX - 1)
            '' '' ''ReDim outdat(OUT_MAX - 1)
            ReDim wData(cRsultTky - 1)
            ReDim pos(cRetAxisPos - 1)
            ReDim bppos(cRetBpPos - 1)
        End Sub
    End Structure
#End Region

#Region "DllTrimFnc.dll�֐��̒�`"
    '===========================================================================
    '   DllTrimFnc.dll�֐��̒�`
    '===========================================================================
    '#If Not Debug Then
    Public Declare Function ALDFLGRST Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ALDFLGRST@0" () As Integer
    Public Declare Function BIFC Lib "C:\TRIM\DllTrimFnc.dll" Alias "_BIFC@8" (ByVal tim As Short, ByVal brdIdx As Short) As Integer
    Public Declare Function BP_CALIBRATION Lib "C:\TRIM\DllTrimFnc.dll" Alias "_BP_CALIBRATION@32" (ByVal GainX As Double, ByVal GainY As Double, ByVal OfsX As Double, ByVal OfsY As Double) As Integer
    Public Declare Function BPLINEARITY Lib "C:\TRIM\DllTrimFnc.dll" Alias "_BPLINEARITY@20" (ByVal XY As Short, ByVal IDX As Short, ByVal Flg As Short, ByVal Val_Renamed As Double) As Integer
    Public Declare Function BP_MOVE Lib "C:\TRIM\DllTrimFnc.dll" Alias "_MOVE@20" (ByVal dpStx As Double, ByVal dpSty As Double, ByVal Flg As Short) As Integer
    Public Declare Function BPOFF Lib "C:\TRIM\DllTrimFnc.dll" Alias "_BPOFF@16" (ByVal BPOX As Double, ByVal BPOY As Double) As Integer
    Public Declare Function BSIZE Lib "C:\TRIM\DllTrimFnc.dll" Alias "_BSIZE@16" (ByVal BSX As Double, ByVal BSY As Double) As Integer
    Public Declare Function CIRCUT Lib "C:\TRIM\DllTrimFnc.dll" Alias "_CIRCUT@32" (ByVal V As Double, ByVal RADI As Double, ByVal ANG2 As Double, ByVal ANG As Double) As Integer
    Public Declare Function CTRIM Lib "C:\TRIM\DllTrimFnc.dll" Alias "_CTRIM@48" (ByVal X As Double, ByVal y As Double, ByVal VX As Double, ByVal VY As Double, ByVal LIMX As Double, ByVal LIMY As Double) As Integer
    Public Declare Function CUTPOSCOR Lib "C:\TRIM\DllTrimFnc.dll" Alias "_CUTPOSCOR@16" (ByVal rn As Short, ByRef POSX() As Double, ByRef POSY() As Double, ByRef Flg() As Short) As Integer
    Public Declare Function CUTPOSCOR_ALL Lib "C:\TRIM\DllTrimFnc.dll" Alias "_CUTPOSCOR_ALL@8" (ByVal resCnt As Integer, ByRef corrData As CUTPOS_CORRECT_DATA) As Integer
    Public Declare Function DebugMode Lib "C:\TRIM\DllTrimFnc.dll" Alias "_DebugMode@8" (ByVal MODE As Short, ByVal level As Short) As Integer
    Public Declare Function DREAD Lib "C:\TRIM\DllTrimFnc.dll" Alias "_DREAD@4" (ByRef DGSW As Short) As Integer
    Public Declare Function DREAD2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_DREAD2@12" (ByRef DGL As Short, ByRef DGH As Short, ByRef DGSW As Short) As Integer
    Public Declare Function DSCAN Lib "C:\TRIM\DllTrimFnc.dll" Alias "_DSCAN@12" (ByVal HP As Short, ByVal LP As Short, ByVal GP As Short) As Integer
    Public Declare Function EMGRESET Lib "C:\TRIM\DllTrimFnc.dll" Alias "_EMGRESET@0" () As Integer
    Public Declare Function EXTIN Lib "C:\TRIM\DllTrimFnc.dll" Alias "_EXTIN@4" (ByRef EIN As Integer) As Integer
    Public Declare Function EXTOUT Lib "C:\TRIM\DllTrimFnc.dll" Alias "_EXTOUT@4" (ByVal ODAT As Integer) As Integer
    Public Declare Function EXTOUT1 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_EXTOUT1@8" (ByVal EON As Integer, ByVal EOFF As Integer) As Integer
    Public Declare Function EXTOUT2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_EXTOUT2@8" (ByVal EON As Integer, ByVal EOFF As Integer) As Integer
    Public Declare Function EXTOUT3 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_EXTOUT3@8" (ByVal EON As Integer, ByVal EOFF As Integer) As Integer
    Public Declare Function EXTOUT4 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_EXTOUT4@8" (ByVal EON As Integer, ByVal EOFF As Integer) As Integer
    Public Declare Function EXTOUT5 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_EXTOUT5@8" (ByVal EON As Integer, ByVal EOFF As Integer) As Integer
    Public Declare Function EXTRSTSET Lib "C:\TRIM\DllTrimFnc.dll" Alias "_EXTRSTSET@4" (ByVal ODATA As Integer) As Integer
    Public Declare Function FAST_WMEAS Lib "C:\TRIM\DllTrimFnc.dll" Alias "_FAST_WMEAS@8" (ByRef MR As Double, ByVal OSC As Short) As Integer
    Public Declare Function FPRESET Lib "C:\TRIM\DllTrimFnc.dll" Alias "_FPRESET@0" () As Integer
    Public Declare Function FPSET Lib "C:\TRIM\DllTrimFnc.dll" Alias "_FPSET@0" () As Integer
    Public Declare Function FPSET2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_FPSET2@4" (ByVal tim As Integer) As Integer
    'Public Declare Function GET_VERSION Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GET_VERSION@4" (ByRef VER As String) As Integer
    Public Declare Function GET_VERSION Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GET_VERSION@4" (ByRef VER As Double) As Integer
    Public Declare Function GETERRSTS Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GETERRSTS@4" (ByRef ERRSTS As Integer) As Integer
    'Public Declare Function GETSETTIME Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GETSETTIME@0" () As Integer
    Public Declare Function GET_Z2_POS Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GET_Z2_POS@4" (ByRef Z2 As Double) As Integer
    'Public Declare Function GPBActRen Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GPBActRen@4" (ByVal brdIdx As Short) As Integer
    'Public Declare Function GPBAdrStRead Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GPBAdrStRead@8" (ByRef btadrst As Short, ByVal brdIdx As Short) As Integer
    'Public Declare Function GPBClrRen Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GPBClrRen@4" (ByVal brdIdx As Short) As Integer
    'Public Declare Function GPBExeSpoll Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GPBExeSpoll@20" (ByRef bttlks As Short, ByVal wtlknum As Short, ByRef bttlk As Short, ByRef btstb As Short, ByVal brdIdx As Short) As Integer
    'Public Declare Function GPBGetAdrs Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GPBGetAdrs@8" (ByRef btadrs As Short, ByVal brdIdx As Short) As Integer
    'Public Declare Function GPBGetDlm Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GPBGetDlm@8" (ByRef btdlm As Short, ByVal brdIdx As Short) As Integer
    'Public Declare Function GPBGetTimeout Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GPBGetTimeout@8" (ByRef wtim As Short, ByVal brdIdx As Short) As Integer
    'Public Declare Function GPBIfc Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GPBIfc@8" (ByVal wtim As Short, ByVal brdIdx As Short) As Integer
    'Public Declare Function GPBInit Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GPBInit@4" (ByVal brdIdx As Short) As Integer
    'Public Declare Function GPBRecvData Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GPBRecvData@16" (ByRef btdata As Short, ByVal wsize As Short, ByRef wrecv As Short, ByVal brdIdx As Short) As Integer
    'Public Declare Function GPBSendCmd Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GPBSendCmd@12" (ByVal btcmd As String, ByVal wsize As Short, ByVal brdIdx As Short) As Integer
    'Public Declare Function GPBSendData Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GPBSendData@16" (ByVal btdata As String, ByVal wsize As Short, ByVal weoi As Short, ByVal brdIdx As Short) As Integer
    'Public Declare Function GPBSetDlm Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GPBSetDlm@8" (ByVal btdlm As Short, ByVal brdIdx As Short) As Integer
    'Public Declare Function GPBSetTimeout Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GPBSetTimeout@8" (ByVal wtim As Short, ByVal brdIdx As Short) As Integer
    'Public Declare Function GPBWaitForSRQ Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GPBWaitForSRQ@8" (ByVal timeout As Short, ByVal brdIdx As Short) As Integer
    'Public Declare Function GPERecv Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GPERecv@28" (ByVal bttlk As Short, ByRef btlsns As Short, ByVal wlsnnum As Short, ByRef btmsge As Short, ByVal wsize As Short, ByRef wrecv As Short, ByVal brdIdx As Short) As Integer
    'Public Declare Function GPESend Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GPESend@20" (ByRef btlsns As Short, ByVal wlsnnum As Short, ByVal btmsge As String, ByVal wsize As Short, ByVal brdIdx As Short) As Integer
    'Public Declare Function GPSGetSrqTkn Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GPSGetSrqTkn@8" (ByRef hSrqSem As Integer, ByVal brdIdx As Short) As Integer
    'Public Declare Function GPSInit Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GPSInit@0" () As Integer
    'Public Declare Function GPSLExeSRQ Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GPSLExeSRQ@12" (ByVal weoi As Short, ByVal wdevst As Short, ByVal brdIdx As Short) As Integer
    'Public Declare Function GPSLock Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GPSLock@4" (ByVal timeout As Short) As Integer
    'Public Declare Function GPSUnlock Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GPSUnlock@0" () As Integer
    'Public Declare Function HCUT2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_HCUT2@40" (ByVal LTDIR As Short, ByVal L1 As Double, ByVal L2 As Double, ByVal L3 As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    Public Declare Function IACLEAR Lib "C:\TRIM\DllTrimFnc.dll" Alias "_IACLEAR@0" () As Integer
    Public Declare Function ICLEAR Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ICLEAR@4" (ByVal GADR As Short) As Integer
    Public Declare Function ICUT2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ICUT2@24" (ByVal n As Short, ByVal L As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    Public Declare Function IDELIM Lib "C:\TRIM\DllTrimFnc.dll" Alias "_IDELIM@4" (ByVal DLM As Short) As Integer
    Public Declare Function ILUM Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ILUM@4" (ByVal sw As Short) As Integer
    Public Declare Function InitFunction Lib "C:\TRIM\DllTrimFnc.dll" Alias "_InitFunction@0" () As Integer
    Public Declare Function INP16 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_INP16@8" (ByVal ADR As Integer, ByRef DAT As Integer) As Integer
    ''Public Declare Function INtimeGWInitialize Lib "C:\TRIM\DllTrimFnc.dll" Alias "_INtimeGWInitialize@0" () As Integer
    ''Public Declare Function INtimeGWTerminate Lib "C:\TRIM\DllTrimFnc.dll" Alias "_INtimeGWTerminate@0" () As Integer
    Public Declare Function IREAD Lib "C:\TRIM\DllTrimFnc.dll" Alias "_IREAD@8" (ByVal GADR As Short, ByVal DAT As String) As Integer
    Public Declare Function IREAD2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_IREAD2@12" (ByVal GADR As Short, ByVal GADR2 As Short, ByVal DAT As String) As Integer
    Public Declare Function IREADM Lib "C:\TRIM\DllTrimFnc.dll" Alias "_IREADM@16" (ByVal GADR As Short, ByRef MAX As Short, ByRef DAT As String, ByVal DLM As String) As Integer
    Public Declare Function IRHVAL Lib "C:\TRIM\DllTrimFnc.dll" Alias "_IRHVAL@12" (ByVal GADR As Short, ByVal HED As Short, ByRef DAT As Double) As Integer
    Public Declare Function IRHVAL2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_IRHVAL2@16" (ByVal GADR As Short, ByVal GADR2 As Short, ByVal HED As Short, ByRef DAT As Double) As Integer
    Public Declare Function IRMVAL Lib "C:\TRIM\DllTrimFnc.dll" Alias "_IRMVAL@16" (ByVal GADR As Short, ByRef MAX As Short, ByRef DAT As Double, ByRef DLM As String) As Integer
    Public Declare Function IRVAL Lib "C:\TRIM\DllTrimFnc.dll" Alias "_IRVAL@8" (ByVal GADR As Short, ByRef DAT As Double) As Integer
    Public Declare Function IRVAL2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_IRVAL2@12" (ByVal GADR As Short, ByVal GADR2 As Short, ByRef DAT As Double) As Integer
    Public Declare Function ITIMEOUT Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ITIMEOUT@4" (ByVal tim As Short) As Integer
    ''Public Declare Function ITIMESET Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ITIMESET@4" (ByVal MODE As Short) As Integer
    Public Declare Function ITRIGGER Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ITRIGGER@4" (ByVal GADR As Short) As Integer
    Public Declare Function IWRITE Lib "C:\TRIM\DllTrimFnc.dll" Alias "_IWRITE@8" (ByVal GADR As Short, ByVal DAT As String) As Integer
    Public Declare Function IWRITE2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_IWRITE2@12" (ByVal GADR As Short, ByVal GADR2 As Short, ByVal DAT As String) As Integer
    Public Declare Function LASEROFF Lib "C:\TRIM\DllTrimFnc.dll" Alias "_LASEROFF@0" () As Integer
    Public Declare Function LASERON Lib "C:\TRIM\DllTrimFnc.dll" Alias "_LASERON@0" () As Integer
    Public Declare Function LATTSET Lib "C:\TRIM\DllTrimFnc.dll" Alias "_LATTSET@8" (ByVal FAT As Integer, ByVal RAT As Integer) As Integer
    Public Declare Function LCUT2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_LCUT2@32" (ByVal LTDIR As Short, ByVal L1 As Double, ByVal L2 As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    Public Declare Function RANGE_SET Lib "C:\TRIM\DllTrimFnc.dll" Alias "_RANGE_SET@8" (ByVal MSDEV As String, ByVal rangeNo As Integer) As Integer
    Public Declare Function MFSET Lib "C:\TRIM\DllTrimFnc.dll" Alias "_MFSET@4" (ByVal MSDEV As String) As Integer
    Public Declare Function ATTRESET Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ATTRESET@0" () As Integer     'V2.1.0.0�A

    '�i�V�KIF�j���d���Ή���
    Public Declare Function MFSET_EX Lib "C:\TRIM\DllTrimFnc.dll" Alias "_MFSET_EX@12" (ByVal MSDEV As String, ByVal target As Double) As Integer

    Public Declare Function MSCAN Lib "C:\TRIM\DllTrimFnc.dll" Alias "_MSCAN@28" (ByVal HP As Short, ByVal LP As Short, ByVal GP1 As Short, ByVal GP2 As Short, ByVal GP3 As Short, ByVal GP4 As Short, ByVal GP5 As Short) As Integer
    Public Declare Function NO_OPERATION Lib "C:\TRIM\DllTrimFnc.dll" Alias "_NO_OPERATION@20" (ByRef X As Double, ByRef y As Double, ByRef z As Double, ByRef BPx As Double, ByRef BPy As Double) As Integer
    'Public Declare Function GET_STATUS Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GET_STATUS@20" (ByRef X As Double, ByRef y As Double, ByRef z As Double, ByRef BPx As Double, ByRef BPy As Double) As Integer
    Public Declare Function GET_STATUS Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GET_STATUS@24" (ByVal getBpMode As Integer, ByRef X As Double, ByRef y As Double, ByRef z As Double, ByRef BPx As Double, ByRef BPy As Double) As Integer
    Public Declare Function OUT16 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_OUT16@8" (ByVal ADR As Integer, ByVal DAT As Integer) As Integer
    Public Declare Function OUTBIT Lib "C:\TRIM\DllTrimFnc.dll" Alias "_OUTBIT@12" (ByVal CATEGORY As Short, ByVal BITNUM As Short, ByVal BON As Short) As Integer
    Public Declare Function PIN16 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_PIN16@8" (ByVal ADR As Integer, ByRef DAT As Integer) As Integer
    Public Declare Function PROBOFF Lib "C:\TRIM\DllTrimFnc.dll" Alias "_PROBOFF@0" () As Integer
    Public Declare Function PROBOFF_EX Lib "C:\TRIM\DllTrimFnc.dll" Alias "_PROBOFF_EX@8" (ByVal Pos As Double) As Integer
    Public Declare Function PROBOFF2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_PROBOFF2@0" () As Integer
    Public Declare Function PROBON Lib "C:\TRIM\DllTrimFnc.dll" Alias "_PROBON@0" () As Integer
    Public Declare Function PROBON2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_PROBON2@8" (ByVal Z2ON As Double) As Integer
    Public Declare Function PROBUP Lib "C:\TRIM\DllTrimFnc.dll" Alias "_PROBUP@8" (ByVal UP As Double) As Integer
    Public Declare Function PROCPOWER Lib "C:\TRIM\DllTrimFnc.dll" Alias "_PROCPOWER@4" (ByVal POWER As Short) As Integer
    Public Declare Sub PROP_SET Lib "C:\TRIM\DllTrimFnc.dll" Alias "_PROP_SET@48" (ByVal ZON As Double, ByVal ZOFF As Double, ByVal POSX As Double, ByVal POSY As Double, ByVal SmaxX As Double, ByVal SmaxY As Double)
    Public Declare Function QRATE Lib "C:\TRIM\DllTrimFnc.dll" Alias "_QRATE@8" (ByVal QR As Double) As Integer
    Public Declare Function RangeCorrect Lib "C:\TRIM\DllTrimFnc.dll" Alias "_RangeCorrect@24" (ByVal IDX As Short, ByVal Val_Renamed As Double, ByVal Flg As Short, ByVal RMin As Short, ByVal RMax As Short) As Integer
    Public Declare Function RATIO2EXP Lib "C:\TRIM\DllTrimFnc.dll" Alias "_RATIO2EXP@8" (ByVal RNO As Integer, ByVal MKSTR As String) As Integer
    Public Declare Function RBACK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_RBACK@0" () As Integer
    Public Declare Function RESET_Renamed Lib "C:\TRIM\DllTrimFnc.dll" Alias "_RESET@0" () As Integer
    Public Declare Function RINIT Lib "C:\TRIM\DllTrimFnc.dll" Alias "_RINIT@0" () As Integer
    Public Declare Function RMEAS Lib "C:\TRIM\DllTrimFnc.dll" Alias "_RMEAS@12" (ByVal MODE As Short, ByVal DVM As Short, ByRef r As Double) As Integer
    Public Declare Function RMeasHL Lib "C:\TRIM\DllTrimFnc.dll" Alias "_RMeasHL@28" (ByVal HP As Short, ByVal LP As Short, ByVal MODE As Short, ByVal NOM As Double, ByRef r As Double, ByRef ad As Short) As Integer
    Public Declare Function ROUND Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ROUND@4" (ByVal PLS As Integer) As Integer
    Public Declare Function ROUND4 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ROUND4@8" (ByVal ANG As Double) As Integer
    Public Declare Function RTEST Lib "C:\TRIM\DllTrimFnc.dll" Alias "_RTEST@36" (ByVal NOM As Double, ByVal MODE As Short, ByVal LOW As Double, ByVal HIGH As Double, ByVal JM As Short, ByVal DVM As Short) As Integer
    Public Declare Function RTRACK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_RTRACK@12" (ByVal NOM As Double, ByVal JM As Short) As Integer
    Public Declare Function SBACK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_SBACK@0" () As Integer
    Public Declare Function SETDLY Lib "C:\TRIM\DllTrimFnc.dll" Alias "_SETDLY@4" (ByVal DTIME As Integer) As Integer
    Public Declare Function SLIDECOVERCHK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_SLIDECOVERCHK@4" (ByVal CHK As Short) As Integer
    Public Declare Function SMOVE Lib "C:\TRIM\DllTrimFnc.dll" Alias "_SMOVE@16" (ByVal XD As Double, ByVal YD As Double) As Integer
    Public Declare Function SMOVE2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_SMOVE2@16" (ByVal XP As Double, ByVal YP As Double) As Integer
    Public Declare Function SMOVE_EX Lib "C:\TRIM\DllTrimFnc.dll" Alias "_SMOVE_EX@20" (ByVal XD As Double, ByVal YD As Double, ByVal OnOff As Short) As Integer
    Public Declare Function SMOVE2_EX Lib "C:\TRIM\DllTrimFnc.dll" Alias "_SMOVE_EX@24" (ByVal XP As Double, ByVal YP As Double, ByVal OnOff As Short, ByVal jogMode As Integer) As Integer
    'Public Declare Function SMOVE2_EX Lib "C:\TRIM\DllTrimFnc.dll" Alias "_SMOVE_EX@20" (ByVal XP As Double, ByVal YP As Double, ByVal OnOff As Short) As Integer
    '    Public Declare Function START Lib "C:\TRIM\DllTrimFnc.dll" Alias "_START@20" (ByVal Z1 As Short, ByVal XOFF As Double, ByVal YOFF As Double) As Integer
    Public Declare Function START Lib "C:\TRIM\DllTrimFnc.dll" Alias "_START@16" (ByVal XOFF As Double, ByVal YOFF As Double) As Integer
    Public Declare Function STCUT Lib "C:\TRIM\DllTrimFnc.dll" Alias "_STCUT@60" (ByVal L As Double, ByVal V As Double, ByVal NOM As Double, ByVal CUTOFF As Double, ByVal V2 As Double, ByVal Q2 As Double, ByVal DIR_Renamed As Short, ByVal CUTMODE As Short, ByVal MODE As Short) As Integer
    Public Declare Function SYSINIT Lib "C:\TRIM\DllTrimFnc.dll" Alias "_SYSINIT@16" (ByVal ZOFF As Double, ByVal ZON As Double) As Integer
    Public Declare Function TEST Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TEST@36" (ByVal X As Double, ByVal NOM As Double, ByVal MODE As Short, ByVal LOW As Double, ByVal HIGH As Double) As Integer
    Public Declare Function TRIM_NGMARK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIM_NGMARK@32" (ByVal POSX As Double, ByVal POSY As Double, ByVal TM As Short, ByVal SN As Short, ByVal sw As Short, ByVal Flg As Short) As Integer
    'Public Declare Function TRIM_RESULT_WORD Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIM_RESULT@24" (ByVal KD As Short, ByVal SN As Short, ByVal NM As Short, ByVal CI As Short, ByVal DI As Short, ByRef Res() As UShort) As Integer
    Public Declare Function TRIM_RESULT_WORD Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIM_RESULT@24" (ByVal KD As Short, ByVal SN As Short, ByVal NM As Short, ByVal CI As Short, ByVal DI As Short, ByRef Res As UShort) As Integer
    Public Declare Function TRIM_RESULT_Double Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIM_RESULT@24" (ByVal KD As Short, ByVal SN As Short, ByVal NM As Short, ByVal CI As Short, ByVal DI As Short, ByRef Res As Double) As Integer
    'Public Declare Function TRIM_RESULT_Double Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIM_RESULT@24" (ByVal KD As Short, ByVal SN As Short, ByVal NM As Short, ByVal CI As Short, ByVal DI As Short, ByRef Res As TRIM_RES_Double) As Integer
    Public Declare Function TRIM80 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIM80@24" (ByVal X As Double, ByVal y As Double, ByVal V As Double) As Integer

    ' �u���b�N�P�ʂ̃g���~���O����
    Public Declare Function TRIMBLOCK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMBLOCK@20" (ByVal MD As Short, ByVal HZ As Short, ByVal RI As Short, ByVal CI As Short, ByVal NG As Short) As Integer
    ' �v���[�g�f�[�^���M
    Public Declare Function TRIMDATA_PLATE Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA_PLATE@8" (ByRef msg As TRIM_PLATE_DATA, ByVal tkyKnd As Integer) As Integer
    '    Public Declare Function TRIMDATA_GPIB Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA_GPIB@8" (ByRef msg As TRIM_DAT_GPIB, ByRef sts As S_RES_DAT) As Integer
    Public Declare Function TRIMDATA_GPIB Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA_GPIB@8" (ByRef msg As TRIM_PLATE_GPIB, ByVal tkyKnd As Integer) As Integer
    Public Declare Function TRIMDATA_RESISTOR Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA_RESISTOR@8" (ByRef msg As TRIM_RESISTOR_DATA, ByVal resNo As Integer) As Integer
    Public Declare Function TRIMDATA_CUTDATA Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA_CUTDATA@12" (ByRef msg As TRIM_CUT_DATA, ByVal resNo As Integer, ByVal cutNo As Integer) As Integer
    ' �J�b�g�p�����[�^���M
    Public Declare Function TRIMDATA_CutST Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA_CUTPRM@12" (ByRef msg As PRM_CUT_ST, ByVal resNo As Integer, ByVal cutNo As Integer) As Integer
    Public Declare Function TRIMDATA_CutL Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA_CUTPRM@12" (ByRef msg As PRM_CUT_L, ByVal resNo As Integer, ByVal cutNo As Integer) As Integer
    Public Declare Function TRIMDATA_CutHK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA_CUTPRM@12" (ByRef msg As PRM_CUT_HOOK, ByVal resNo As Integer, ByVal cutNo As Integer) As Integer
    Public Declare Function TRIMDATA_CutIX Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA_CUTPRM@12" (ByRef msg As PRM_CUT_INDEX, ByVal resNo As Integer, ByVal cutNo As Integer) As Integer
    Public Declare Function TRIMDATA_CutSC Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA_CUTPRM@12" (ByRef msg As PRM_CUT_SCAN, ByVal resNo As Integer, ByVal cutNo As Integer) As Integer
    Public Declare Function TRIMDATA_CutMK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA_CUTPRM@12" (ByRef msg As PRM_CUT_MARKING, ByVal resNo As Integer, ByVal cutNo As Integer) As Integer
    Public Declare Function TRIMDATA_CutES Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA_CUTPRM@12" (ByRef msg As PRM_CUT_ES, ByVal resNo As Integer, ByVal cutNo As Integer) As Integer

    '    'Public Declare Function TRIMBLOCK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMBLOCK@24" (ByVal MD As Short, ByVal HZ As Short, ByVal RI As Short, ByVal CI As Short, ByVal NG As Short, ByRef sts As S_RES_DAT) As Integer
    '    Public Declare Function TRIMBLOCK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMBLOCK@16" (ByVal MD As Short, ByVal HZ As Short, ByVal CI As Short, ByVal NG As Short) As Integer
    ''    Public Declare Function TRIMDATA Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA_PLATE@8" (ByRef msg As TRIM_DAT_PLATE, ByRef sts As S_RES_DAT) As Integer
    '    Public Declare Function TRIMDATA_PLATE Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA_PLATE@8" (ByRef msg As TRIM_DAT_PLATE, ByVal tkyKnd As Integer) As Integer
    ''    Public Declare Function TRIMDATA_GPIB Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA_GPIB@8" (ByRef msg As TRIM_DAT_GPIB, ByRef sts As S_RES_DAT) As Integer
    '    Public Declare Function TRIMDATA_RESISTOR Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA_RESISTOR@8" (ByRef msg As TRIM_DAT_RESISTOR, byval resNo as Integer ) As Integer
    '    Public Declare Function TRIMDATA_CUTDATA Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA_CUTDATA@12" (ByRef msg As TRIM_DAT_CUT, byval resNo as Integer, byval cutNo as Integer) As Integer
    '    Public Declare Function TRIMDATA_CUTPRM Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA_CUTPRM@12" (ByRef msg As TRIM_DAT_CUT, byval resNo as Integer, byval cutNo as Integer ) As Integer
    '    Public Declare Function TRIMDATA_CutST Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA@8" (ByRef msg As TRIM_DAT_CUT_ST, ByRef sts As S_RES_DAT) As Integer
    '    Public Declare Function TRIMDATA_CutL Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA@8" (ByRef msg As TRIM_DAT_CUT_L, ByRef sts As S_RES_DAT) As Integer
    '    Public Declare Function TRIMDATA_CutHK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA@8" (ByRef msg As TRIM_DAT_CUT_HOOK, ByRef sts As S_RES_DAT) As Integer
    '    Public Declare Function TRIMDATA_CutIX Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA@8" (ByRef msg As TRIM_DAT_CUT_INDEX, ByRef sts As S_RES_DAT) As Integer
    '    Public Declare Function TRIMDATA_CutSC Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA@8" (ByRef msg As TRIM_DAT_CUT_SCAN, ByRef sts As S_RES_DAT) As Integer
    '    Public Declare Function TRIMDATA_CutMK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA@8" (ByRef msg As TRIM_DAT_CUT_MARKING, ByRef sts As S_RES_DAT) As Integer
    '    Public Declare Function TRIMDATA_CutC Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA@8" (ByRef msg As TRIM_DAT_CUT_C, ByRef sts As S_RES_DAT) As Integer
    '    Public Declare Function TRIMDATA_CutES Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA@8" (ByRef msg As TRIM_DAT_CUT_ES, ByRef sts As S_RES_DAT) As Integer
    '    Public Declare Function TRIMDATA_CutES2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA@8" (ByRef msg As TRIM_DAT_CUT_ES2, ByRef sts As S_RES_DAT) As Integer
    '    Public Declare Function TRIMDATA_CutZ Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMDATA@8" (ByRef msg As TRIM_DAT_CUT_Z, ByRef sts As S_RES_DAT) As Integer
    Public Declare Function TRIMEND Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIMEND@0" () As Integer
    'Public Declare Function TSTEP Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TSTEP@8" (ByVal BNX As Short, ByVal BNY As Short) As Integer
    Public Declare Function TSTEP Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TSTEP@24" (ByVal BNX As Short, ByVal BNY As Short, ByVal stepOffX As Double, ByVal stepOffY As Double) As Integer
    Public Declare Function UCUT2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_UCUT2@40" (ByVal LTDIR As Short, ByVal L1 As Double, ByVal L2 As Double, ByVal RADI As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    Public Declare Function UCUT_PARAMSET Lib "C:\TRIM\DllTrimFnc.dll" Alias "_UCUT_PARAMSET@24" (ByVal MD As Short, ByVal KD As Short, ByVal RNO As Short, ByVal IDX As Short, ByVal EL As Short, ByRef pstPRM As UCUT_PARAM_EL) As Integer
    Public Declare Function UCUT_RESULT Lib "C:\TRIM\DllTrimFnc.dll" Alias "_UCUT_RESULT@16" (ByVal RNO As Short, ByVal CNO As Short, ByRef UcutNO As Short, ByRef InitVal As Double) As Integer
    Public Declare Function UCUT4RESULT Lib "C:\TRIM\DllTrimFnc.dll" Alias "_UCUT4RESULT@8" (ByRef sRegNo_p As Short, ByRef sCutNo_p As Short) As Integer
    Public Declare Function VCIRTRIM Lib "C:\TRIM\DllTrimFnc.dll" Alias "_VCIRTRIM@44" (ByVal SLP As Short, ByVal NOM As Double, ByVal V As Double, ByVal RADI As Double, ByVal ANG2 As Double, ByVal ANG As Double) As Integer
    Public Declare Function VCTRIM Lib "C:\TRIM\DllTrimFnc.dll" Alias "_VCTRIM@64" (ByVal SLP As Short, ByVal NOM As Double, ByVal MD As Short, ByVal X As Double, ByVal y As Double, ByVal VX As Double, ByVal VY As Double, ByVal LIMX As Double, ByVal LIMY As Double) As Integer
    Public Declare Function VHTRIM2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_VHTRIM2@64" (ByVal SLP As Short, ByVal NOM As Double, ByVal MD As Short, ByVal LTP As Double, ByVal LTDIR As Short, ByVal L1 As Double, ByVal L2 As Double, ByVal L3 As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    Public Declare Function VITRIM2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_VITRIM2@40" (ByVal SLP As Short, ByVal NOM As Double, ByVal MD As Short, ByVal n As Short, ByVal L As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    Public Declare Function VLTRIM2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_VLTRIM2@56" (ByVal SLP As Short, ByVal NOM As Double, ByVal MD As Short, ByVal LTP As Double, ByVal LTDIR As Short, ByVal L1 As Double, ByVal L2 As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    Public Declare Function VMEAS Lib "C:\TRIM\DllTrimFnc.dll" Alias "_VMEAS@12" (ByVal MODE As Short, ByVal DVM As Short, ByRef V As Double) As Integer
    Public Declare Function VRangeCorrect Lib "C:\TRIM\DllTrimFnc.dll" Alias "_VRangeCorrect@24" (ByVal IDX As Short, ByVal Val_Renamed As Double, ByVal Flg As Short, ByVal RMin As Short, ByVal RMax As Short) As Integer
    Public Declare Function VTEST Lib "C:\TRIM\DllTrimFnc.dll" Alias "_VTEST@36" (ByVal NOM As Double, ByVal MODE As Short, ByVal LOW As Double, ByVal HIGH As Double, ByVal JM As Short, ByVal DVM As Short) As Integer
    Public Declare Function VTRACK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_VTRACK@16" (ByVal SLP As Short, ByVal NOM As Double, ByVal JM As Short) As Integer
    Public Declare Function VUTRIM2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_VUTRIM2@64" (ByVal SLP As Short, ByVal NOM As Double, ByVal MD As Short, ByVal LTP As Double, ByVal LTDIR As Short, ByVal L1 As Double, ByVal L2 As Double, ByVal RADI As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    Public Declare Function VUTRIM4 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_VUTRIM4@88" (ByVal SLP As Short, ByVal NOM As Double, ByVal MD As Short, ByVal LTP As Double, ByVal LTDIR As Short, ByVal L1 As Double, ByVal L2 As Double, ByVal RADI As Double, ByVal V As Double, ByVal ANG As Short, ByVal trmd As Short, ByVal trl As Double, ByVal cn As Short, ByVal DT As Short, ByVal MODE As Short) As Integer
    Public Declare Function VUTRIM3 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_VUTRIM3@72" (ByVal SLP As Short, ByVal NOM As Double, ByVal MD As Short, ByVal LTP As Double, ByVal LTDIR As Short, ByVal L1 As Double, ByVal L2 As Double, ByVal RADI As Double, ByVal RADI2 As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    '(2011/06/03)
    '   ���g�p�̈׍폜����
    'Public Declare Function XYOFF Lib "C:\TRIM\DllTrimFnc.dll" Alias "_XYOFF@16" (ByVal XOFF As Double, ByVal YOFF As Double) As Integer
    Public Declare Function ZABSVACCUME Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZABSVACCUME@4" (ByVal ZON As Integer) As Integer
    Public Declare Function ZATLDGET Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZATLDGET@4" (ByRef LDIN As Integer) As Integer
    Public Declare Function ZATLDSET Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZATLDSET@8" (ByVal LDON As Integer, ByVal LDOFF As Integer) As Integer
    Public Declare Function ZBPLOGICALCOORD Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZBPLOGICALCOORD@4" (ByVal COORD As Integer) As Integer
    Public Declare Function ZCONRST Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZCONRST@0" () As Integer
    Public Declare Function ZGETBPPOS Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZGETBPPOS@8" (ByRef XP As Double, ByRef YP As Double) As Integer
    Public Declare Function ZGETDCVRANG Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZGETDCVRANG@4" (ByRef VMAX As Double) As Integer
    Public Declare Function ZGETPHPOS Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZGETPHPOS@8" (ByRef NOWXP As Double, ByRef NOWYP As Double) As Integer
    Public Declare Function ZGETPHPOS2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZGETPHPOS2@8" (ByRef NOWXP As Double, ByRef NOWYP As Double) As Integer

    Public Declare Function ZGETSRVSIGNAL Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZGETSRVSIGNAL@16" (ByRef X As Integer, ByRef y As Integer, ByRef z As Integer, ByRef t As Integer) As Integer
    'Public Declare Function ZGETTRMPOS Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZGETTRMPOS@24" (ByRef TRIMX As Double, ByRef TRIMY As Double, ByRef RCX As Double, ByRef RCY As Double, ByRef SMAX As Double, ByRef SMAY As Double) As Integer
    Public Declare Function ZINPSTS Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZINPSTS@8" (ByVal sw As Integer, ByRef sts As Integer) As Integer
    Public Declare Function ZLATCHOFF Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZLATCHOFF@0" () As Integer
    Public Declare Function ZZMOVE Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZMOVE@12" (ByVal z As Double, ByVal MD As Short) As Integer
    Public Declare Function ZZMOVE2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZMOVE2@12" (ByVal z As Double, ByVal MD As Short) As Integer
    Public Declare Function ZRCIRTRIM Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZRCIRTRIM@44" (ByVal NOM As Double, ByVal RNG As Short, ByVal V As Double, ByVal RADI As Double, ByVal ANG2 As Double, ByVal ANG As Double) As Integer
    Public Declare Function ZRTRIM2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZRTRIM2@32" (ByVal NOM As Double, ByVal RNG As Short, ByVal L As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    Public Declare Function ZSELXYZSPD Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZSELXYZSPD@4" (ByVal SPD As Integer) As Integer
    Public Declare Function ZSETBPTIME Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZSETBPTIME@8" (ByVal BPTIME As Integer, ByVal EPTIME As Integer) As Integer
    Public Declare Function ZSETPOS2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZSETPOS2@24" (ByVal POS2X As Double, ByVal POS2Y As Double, ByVal POS2Z As Double) As Integer
    Public Declare Function ZSETUCUT Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZSETUCUT@40" (ByVal MD As Short, ByVal RNO As Short, ByVal Index As Short, ByVal EL As Short, ByVal RATIO As Double, ByVal LTP As Double, ByVal LTP2 As Double) As Integer
    Public Declare Function ZSLCOVERCLOSE Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZSLCOVERCLOSE@4" (ByVal ZONOFF As Short) As Integer
    Public Declare Function ZSLCOVEROPEN Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZSLCOVEROPEN@4" (ByVal ZONOFF As Short) As Integer
    Public Declare Function ZSTGXYMODE Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZSTGXYMODE@4" (ByVal MODE As Integer) As Integer
    Public Declare Function ZSTOPSTS Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZSTOPSTS@0" () As Integer
    Public Declare Function ZSTOPSTS2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZSTOPSTS2@0" () As Integer
    Public Declare Function ZSYSPARAM1 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZSYSPARAM1@54" (ByVal POWERCYCLE As Short, ByVal THETA As Short, ByVal BPDIRXY As Short, ByVal BPSIZE As Short, ByVal DCSCANNER As Short, ByVal DCVRANGE As Short, ByVal LRANGE As Short, ByVal LDPOSX As Double, ByVal LDPOSY As Double, ByVal FPSUP As Short, ByVal DELAYSKIP As Short, ByVal OSC As Short) As Integer
    'Public Declare Function ZSYSPARAM2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZSYSPARAM2@60" (ByVal PRBTYP As Short, ByVal SMINMAXZ2 As Double, ByVal ZPTIMEON As Short, ByVal ZPTIMEOFF As Short, ByVal XYTBL As Short, ByVal SmaxX As Double, ByVal SmaxY As Double, ByVal ABSTIME As Integer, ByVal TRIMX As Double, ByVal TRIMY As Double) As Integer
    Public Declare Function ZSYSPARAM2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZSYSPARAM2@68" (ByVal PRBTYP As Short, ByVal SMINMAXZ2 As Double, ByVal ZPTIMEON As Short, ByVal ZPTIMEOFF As Short, ByVal XYTBL As Short, ByVal SmaxX As Double, ByVal SmaxY As Double, ByVal ABSTIME As Integer, ByVal TRIMX As Double, ByVal TRIMY As Double, ByVal BpMoveLimX As Integer, ByVal BpMoveLimY As Integer) As Integer
    'Public Declare Function ZSYSPARAM3 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZSYSPARAM3@16" (ByVal ProcPower2 As Short, ByVal GrvTime As Integer, ByVal UcutType As Short, ByVal ExtBit As Integer) As Integer
    'Public Declare Function ZSYSPARAM3 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZSYSPARAM3@20" (ByVal ProcPower2 As Short, ByVal GrvTime As Integer, ByVal UcutType As Short, ByVal ExtBit As Integer, ByVal PosSpd As Integer) As Integer '###021
    Public Declare Function ZSYSPARAM3 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZSYSPARAM3@24" (ByVal ProcPower2 As Short, ByVal GrvTime As Integer, ByVal UcutType As Short, ByVal ExtBit As Integer, ByVal PosSpd As Integer, ByVal BiasOn_AddTime As Integer) As Integer
    Public Declare Function ZTIMERINIT Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZTIMERINIT@0" () As Integer
    Public Declare Function ZVMEAS Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZVMEAS@12" (ByVal MODE As Short, ByVal DVM As Short, ByRef V As Double) As Integer
    Public Declare Function ZWAIT Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZWAIT@4" (ByVal lngWaitMilliSec As Integer) As Integer
    Public Declare Function ZZGETRTMODULEINFO Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZZGETRTMODULEINFO@0" () As Integer
    Public Declare Function Z_INIT Lib "C:\TRIM\DllTrimFnc.dll" Alias "_Z_INIT@0" () As Integer
    'About TRIMMING
    Public Declare Function ZRANGTRIM Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ZRANGTRIM@32" (ByVal NOM As Double, ByVal RNG As Short, ByVal L As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    Public Declare Function VTRIM2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_VTRIM2@32" (ByVal SLP As Short, ByVal NOM As Double, ByVal L As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    Public Declare Function CUT2 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_CUT2@20" (ByVal L As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    Public Declare Function CMARK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_CMARK@40" (ByVal MKSTR As String, ByVal STX As Double, ByVal STY As Double, ByVal HIGH As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    '###1042�@    Public Declare Function TrimMK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TrimMK@52" (ByVal MKSTR As String, ByVal STX As Double, ByVal STY As Double, ByVal HIGH As Double, ByVal V As Double, ByVal ANG As Short, ByVal QRate1 As Double, ByVal condNoCut1 As Short) As Integer
    Public Declare Function TrimMK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIM_MK@56" (ByVal MKSTR As String, ByVal STX As Double, ByVal STY As Double, ByVal HIGH As Double, ByVal V As Double, ByVal ANG As Short, ByVal QRate1 As Double, ByVal condNoCut1 As Short, ByVal moveMode As Short) As Integer '###1042�@

    '�V�KI/F
    '    Public Declare Function TRIM_ST Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIM_ST@76" (ByVal MOVEMODE As Integer, ByVal CUTMODE As Integer, ByVal POS As Integer, ByVal SLP As Integer, ByVal NOM As Double, ByVal L As Double, ByVal V As Double, ByVal V_RET As Double, ByVal ANG As Integer, ByVal QRATE As Double, ByVal QRATE_RET As Double, ByVal CUTCOND_NO As Integer, ByVal CUTCOND_NO_RET As Integer) As Long
    'Public Declare Function TRIM_ST Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIM_ST@60" (ByVal MOVEMODE As short, ByVal CUTMODE As short, ByVal SLP As short, ByVal NOM As Double, ByVal L As Double, ByVal V As Double, ByVal V_RET As Double, ByVal ANG As short, ByVal QRATE As Double, ByVal QRATE_RET As Double, ByVal CUTCOND_NO As short, ByVal CUTCOND_NO_RET As short) As  Integer
    'Public Declare Function TRIM_L Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIM_ST@116" (ByVal MOVEMODE As short, ByVal CUTMODE As short, ByVal SLP As short, ByVal NOM As Double, ByVal MD As short, ByVal LTP As Double, ByVal LTDIR As short, ByVal L1 As Double, ByVal L2 As Double, ByVal V As Double, ByVal V2 As Double, ByVal V_RET As Double, ByVal V_RET2 As Double, ByVal ANG As short, ByVal QRATE As Double, ByVal QRATE2 As Double, ByVal QRATE_RET As Double, ByVal QRATE_RET2 As Double, ByVal CUTCOND_NO As short, ByVal CUTCOND_NO2 As short, ByVal CUTCOND_NO_RET As short, ByVal CUTCOND_NO_RET2 As short) As Integer 
    'Public Declare Function TRIM_HkU Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIM_HkU@168" (ByVal MOVEMODE As short, ByVal CUTMODE As short, ByVal SLP As short, ByVal NOM As Double, ByVal MD As short, ByVal LTP As Double, ByVal LTDIR As short, ByVal L1 As Double, ByVal L2 As Double, ByVal L3 As Double, ByVal RADI As Double, ByVal V1 As Double, ByVal V2 As Double, ByVal V3 As Double, ByVal V1_RET As Double, ByVal V2_RET As Double, ByVal V3_RET As Double, ByVal ANG As short, ByVal QRATE1 As Double, ByVal QRATE2 As Double, ByVal QRATE3 As Double, ByVal QRATE1_RET As Double, ByVal QRATE2_RET As Double, ByVal QRATE3_RET As Double, ByVal CUTCOND_NO1 As short, ByVal CUTCOND_NO2 As short, ByVal CUTCOND_NO3 As short, ByVal CUTCOND_NO1_RET As short, ByVal CUTCOND_NO2_RET As short, ByVal CUTCOND_NO3_RET As short) As Integer 
    Public Declare Function TRIM_ST Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIM_ST@4" (ByRef CutCmnPrm As CUT_COMMON_PRM) As Integer
    Public Declare Function TRIM_L Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIM_L@4" (ByRef CutCmnPrm As CUT_COMMON_PRM) As Integer
    Public Declare Function TRIM_L6 Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIM_L6@4" (ByRef CutCmnPrm As CUT_COMMON_PRM_L6) As Integer
    Public Declare Function TRIM_HkU Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIM_HkU@4" (ByRef CutCmnPrm As CUT_COMMON_PRM) As Integer
    Public Declare Function TRIM_ES Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIM_ESU@4" (ByRef CutCmnPrm As CUT_COMMON_PRM) As Integer
    Public Declare Function TRIM_IX Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIM_IX@4" (ByRef CutCmnPrm As CUT_COMMON_PRM) As Integer
    Public Declare Function MEASURE Lib "C:\TRIM\DllTrimFnc.dll" Alias "_MEASURE@28" (ByVal MEASMODE As Short, ByVal RANGSETTYPE As Short, ByVal MEASTYPE As Short, ByVal TARGET As Double, ByVal RANGE As Short, ByRef RESULT As Double) As Integer
    Public Declare Function FLSET Lib "C:\TRIM\DllTrimFnc.dll" Alias "_FLSET@8" (ByVal mode As Short, ByVal cutCondNo As Short) As Integer
    Public Declare Function SET_FL_ERRLOG Lib "C:\TRIM\DllTrimFnc.dll" Alias "_SET_FL_ERRLOG@4" (ByRef ErrCode As Integer) As Integer
    Public Declare Function TRIM_LWithR Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TRIM_LWithR@4" (ByRef CutCmnPrm As CUT_COMMON_PRM) As Integer           'V2.2.0.0�A 

    ' �V�K�ǉ��R�}���h�i�V�g���}SL43xR�����j
    Public Declare Function SYSTEM_RESET Lib "C:\TRIM\DllTrimFnc.dll" Alias "_SYSTEM_RESET@0" () As Integer
    Public Declare Function SERVO_POWER Lib "C:\TRIM\DllTrimFnc.dll" Alias "_SERVO_POWER@16" (ByVal XAxisOnOff As Integer, ByVal YAxisOnOff As Integer, ByVal ZAxisOnOff As Integer, ByVal TAxisOnOff As Integer) As Integer
    Public Declare Function CLEAR_SERVO_ALARM Lib "C:\TRIM\DllTrimFnc.dll" Alias "_CLEAR_SERVO_ALARM@8" (ByVal XY As Integer, ByVal ZT As Integer) As Integer
    Public Declare Function AXIS_X_INIT Lib "C:\TRIM\DllTrimFnc.dll" Alias "_AXIS_X_INIT@0" () As Integer
    Public Declare Function AXIS_Y_INIT Lib "C:\TRIM\DllTrimFnc.dll" Alias "_AXIS_Y_INIT@0" () As Integer
    Public Declare Function AXIS_Z_INIT Lib "C:\TRIM\DllTrimFnc.dll" Alias "_AXIS_Z_INIT@0" () As Integer
    Public Declare Function GET_ALLAXIS_STATUS Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GET_ALLAXIS_STATUS@8" (ByRef err As Long, ByRef AllStatus As Long) As Integer
    Public Declare Function LAMP_CTRL Lib "C:\TRIM\DllTrimFnc.dll" Alias "_LAMP_CTRL@8" (ByVal LampNo As Integer, ByVal OnOff As Boolean) As Integer
    Public Declare Function COVERLATCH_CLEAR Lib "C:\TRIM\DllTrimFnc.dll" Alias "_COVERLATCH_CLEAR@0" () As Integer
    Public Declare Function COVERLATCH_CHECK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_COVERLATCH_CHECK@4" (ByRef LatchSts As Long) As Integer
    Public Declare Function COVER_CHECK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_COVER_CHECK@4" (ByRef SwitchSts As Long) As Integer
    Public Declare Function INTERLOCK_CHECK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_INTERLOCK_CHECK@8" (ByRef InterlockSts As Integer, ByRef SwitchSts As Long) As Integer
    Public Declare Function ORG_INTERLOCK_CHECK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ORG_INTERLOCK_CHECK@8" (ByRef InterlockSts As Integer, ByRef SwitchSts As Long) As Integer
    Public Declare Function SLIDECOVER_MOVINGCHK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_SLIDECOVER_MOVINGCHK@12" (ByVal OpenCloseChk As Integer, ByVal UseReset As Integer, ByRef SwitchSts As Long) As Integer
    Public Declare Function SLIDECOVER_CLOSECHK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_SLIDECOVER_CLOSECHK@4" (ByRef slidecoverSts As Long) As Integer
    Public Declare Function SLIDECOVER_GETSTS Lib "C:\TRIM\DllTrimFnc.dll" Alias "_SLIDECOVER_GETSTS@4" (ByRef slidecoverSts As Long) As Integer
    Public Declare Function START_SWWAIT Lib "C:\TRIM\DllTrimFnc.dll" Alias "_START_SWWAIT@4" (ByRef SwitchSts As Long) As Integer
    Public Declare Function START_SWCHECK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_START_SWCHECK@8" (ByVal bReleaseCheck As Integer, ByRef SwitchSts As Long) As Integer
    Public Declare Function HALT_SWCHECK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_HALT_SWCHECK@4" (ByRef SwitchSts As Long) As Integer
    Public Declare Function STARTRESET_SWWAIT Lib "C:\TRIM\DllTrimFnc.dll" Alias "_STARTRESET_SWWAIT@4" (ByRef SwitchSts As Long) As Integer
    Public Declare Function ORG_STARTRESET_SWWAIT Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ORG_STARTRESET_SWWAIT@4" (ByRef SwitchSts As Long) As Integer
    Public Declare Function STARTRESET_SWCHECK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_STARTRESET_SWCHECK@8" (ByVal bReleaseCheck As Integer, ByRef SwitchSts As Long) As Integer
    Public Declare Function GET_Z_POS Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GET_Z_POS@4" (ByRef ZPos As Double) As Integer
    Public Declare Function GET_QRATE Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GET_QRATE@4" (ByRef QRate As Double) As Integer
    Public Declare Function CONSOLE_SWCHECK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_CONSOLE_SWCHECK@8" (ByVal BbReleaseCheck As Boolean, ByRef SwitchChk As Long) As Integer
    Public Declare Function Z_SWCHECK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_Z_SWCHECK@4" (ByRef SwitchChk As Long) As Integer
    Public Declare Function EMGSTS_CHECK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_EMGSTS_CHECK@4" (ByRef Status As Integer) As Long
    Public Declare Function ISALIVE_INTIME Lib "C:\TRIM\DllTrimFnc.dll" Alias "_ISALIVE_INTIME@0" () As Integer
    Public Declare Function TERMINATE_INTIME Lib "C:\TRIM\DllTrimFnc.dll" Alias "_TERMINATE_INTIME@0" () As Integer
    Public Declare Function BP_GET_CALIBDATA Lib "C:\TRIM\DllTrimFnc.dll" Alias "_BP_GET_CALIBDATA@16" (ByRef gainX As Double, ByRef gainY As Double, ByRef offsetX As Double, ByRef offsetY As Double) As Integer
    Public Declare Function SIGNALTOWER_CTRLEX Lib "C:\TRIM\DllTrimFnc.dll" Alias "_SIGNALTOWER_CTRLEX@8" (ByVal OnBit As Integer, ByVal OffBit As Integer) As Integer
    Public Declare Function SETAXISSPDX Lib "C:\TRIM\DllTrimFnc.dll" Alias "_SETAXISSPDX@4" (ByVal XH As UInteger) As Integer
    Public Declare Function SETAXISSPDY Lib "C:\TRIM\DllTrimFnc.dll" Alias "_SETAXISSPDY@4" (ByVal YH As UInteger) As Integer ' ###1040�C
    Public Declare Function SETLOADPOS Lib "C:\TRIM\DllTrimFnc.dll" Alias "_SETLOADPOS@16" (ByVal LDPOSX As Double, ByVal LDPOSY As Double) As Integer
    Public Declare Function SETZOFFPOS Lib "C:\TRIM\DllTrimFnc.dll" Alias "_SETZOFFPOS@8" (ByVal Pos As Double) As Integer '###1041�@
    Public Declare Function GET_CUT_LENGTH Lib "C:\TRIM\DllTrimFnc.dll" Alias "_GET_CUT_LENGTH@4" (ByRef Length As Double) As Integer
    Public Declare Function COVERCHK_ONOFF Lib "C:\TRIM\DllTrimFnc.dll" Alias "_COVERCHK_ONOFF@4" (ByVal mode As Short) As Integer       ' 'V2.2.0.0�D
    Public Declare Function SPLASER_EXTDIODESET Lib "C:\TRIM\DllTrimFnc.dll" Alias "_SPLASER_EXTDIODESET@4" (ByVal mode As Short) As Integer       ''V2.2.0.030



    '�f�o�b�O/���u�]���p�R�}���h
    Public Declare Function SETLOG_ALLTARGET Lib "C:\TRIM\DllTrimFnc.dll" Alias "_SETLOG_ALLTARGET@36" (ByVal base As Short, ByVal io As Short, ByVal laser As Short, ByVal bp As Short, ByVal meas As Short, ByVal trim As Short, ByVal correct As Short, ByVal stage As Short, ByVal loader As Short) As Integer
    Public Declare Function SETLOG_TARGET Lib "C:\TRIM\DllTrimFnc.dll" Alias "_SETLOG_TARGET@8" (ByVal segNo As Integer, ByVal status As UInteger) As Integer
    Public Declare Function PERFORMCHK Lib "C:\TRIM\DllTrimFnc.dll" Alias "_PERFORMCHK@12" (ByVal ADDR As UInteger, ByVal COUNT As UInteger, ByVal WAIT As UInteger) As Integer
    Public Declare Function SETAXISSPD Lib "C:\TRIM\DllTrimFnc.dll" Alias "_SETAXISSPD@24" (ByVal XL As UInteger, ByVal XH As UInteger, ByVal YL As UInteger, ByVal YH As UInteger, ByVal ZL As UInteger, ByVal ZH As UInteger) As Integer
    Public Declare Function LSI_RESET Lib "C:\TRIM\DllTrimFnc.dll" Alias "_LSI_RESET@0" () As Integer

    '#Else
    '    '===========================================================================
    '    '   DllTrimFnc.dll�֐��̒�`
    '    '===========================================================================
    '    Public Declare Function ALDFLGRST Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ALDFLGRST@0" () As Integer
    '    Public Declare Function BIFC Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_BIFC@8" (ByVal tim As Short, ByVal brdIdx As Short) As Integer
    '    Public Declare Function BP_CALIBRATION Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_BP_CALIBRATION@32" (ByVal GainX As Double, ByVal GainY As Double, ByVal OfsX As Double, ByVal OfsY As Double) As Integer
    '    Public Declare Function BPLINEARITY Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_BPLINEARITY@20" (ByVal XY As Short, ByVal IDX As Short, ByVal Flg As Short, ByVal Val_Renamed As Double) As Integer
    '    Public Declare Function BP_MOVE Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_MOVE@20" (ByVal dpStx As Double, ByVal dpSty As Double, ByVal Flg As Short) As Integer
    '    Public Declare Function BPOFF Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_BPOFF@16" (ByVal BPOX As Double, ByVal BPOY As Double) As Integer
    '    Public Declare Function BSIZE Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_BSIZE@16" (ByVal BSX As Double, ByVal BSY As Double) As Integer
    '    Public Declare Function CIRCUT Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_CIRCUT@32" (ByVal V As Double, ByVal RADI As Double, ByVal ANG2 As Double, ByVal ANG As Double) As Integer
    '    Public Declare Function CTRIM Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_CTRIM@48" (ByVal X As Double, ByVal y As Double, ByVal VX As Double, ByVal VY As Double, ByVal LIMX As Double, ByVal LIMY As Double) As Integer
    '    Public Declare Function CUTPOSCOR Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_CUTPOSCOR@16" (ByVal rn As Short, ByRef POSX() As Double, ByRef POSY() As Double, ByRef Flg() As Short) As Integer
    '    Public Declare Function CUTPOSCOR_ALL Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_CUTPOSCOR_ALL@8" (ByVal resCnt As Integer, ByRef corrData As CUTPOS_CORRECT_DATA) As Integer
    '    Public Declare Function DebugMode Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_DebugMode@8" (ByVal MODE As Short, ByVal level As Short) As Integer
    '    Public Declare Function DREAD Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_DREAD@4" (ByRef DGSW As Short) As Integer
    '    Public Declare Function DREAD2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_DREAD2@12" (ByRef DGL As Short, ByRef DGH As Short, ByRef DGSW As Short) As Integer
    '    Public Declare Function DSCAN Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_DSCAN@12" (ByVal HP As Short, ByVal LP As Short, ByVal GP As Short) As Integer
    '    Public Declare Function EMGRESET Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_EMGRESET@0" () As Integer
    '    Public Declare Function EXTIN Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_EXTIN@4" (ByRef EIN As Integer) As Integer
    '    Public Declare Function EXTOUT Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_EXTOUT@4" (ByVal ODAT As Integer) As Integer
    '    Public Declare Function EXTOUT1 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_EXTOUT1@8" (ByVal EON As Integer, ByVal EOFF As Integer) As Integer
    '    Public Declare Function EXTOUT2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_EXTOUT2@8" (ByVal EON As Integer, ByVal EOFF As Integer) As Integer
    '    Public Declare Function EXTOUT3 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_EXTOUT3@8" (ByVal EON As Integer, ByVal EOFF As Integer) As Integer
    '    Public Declare Function EXTOUT4 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_EXTOUT4@8" (ByVal EON As Integer, ByVal EOFF As Integer) As Integer
    '    Public Declare Function EXTOUT5 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_EXTOUT5@8" (ByVal EON As Integer, ByVal EOFF As Integer) As Integer
    '    Public Declare Function EXTRSTSET Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_EXTRSTSET@4" (ByVal ODATA As Integer) As Integer
    '    Public Declare Function FAST_WMEAS Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_FAST_WMEAS@8" (ByRef MR As Double, ByVal OSC As Short) As Integer
    '    Public Declare Function FPRESET Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_FPRESET@0" () As Integer
    '    Public Declare Function FPSET Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_FPSET@0" () As Integer
    '    Public Declare Function FPSET2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_FPSET2@4" (ByVal tim As Integer) As Integer
    '    'Public Declare Function GET_VERSION Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GET_VERSION@4" (ByRef VER As String) As Integer
    '    Public Declare Function GET_VERSION Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GET_VERSION@4" (ByRef VER As Double) As Integer
    '    Public Declare Function GETERRSTS Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GETERRSTS@4" (ByRef ERRSTS As Integer) As Integer
    '    Public Declare Function GETSETTIME Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GETSETTIME@0" () As Integer
    '    Public Declare Function GET_Z2_POS Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GET_Z2_POS@4" (ByRef Z2 As Double) As Integer
    '    'Public Declare Function GPBActRen Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GPBActRen@4" (ByVal brdIdx As Short) As Integer
    '    'Public Declare Function GPBAdrStRead Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GPBAdrStRead@8" (ByRef btadrst As Short, ByVal brdIdx As Short) As Integer
    '    'Public Declare Function GPBClrRen Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GPBClrRen@4" (ByVal brdIdx As Short) As Integer
    '    'Public Declare Function GPBExeSpoll Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GPBExeSpoll@20" (ByRef bttlks As Short, ByVal wtlknum As Short, ByRef bttlk As Short, ByRef btstb As Short, ByVal brdIdx As Short) As Integer
    '    'Public Declare Function GPBGetAdrs Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GPBGetAdrs@8" (ByRef btadrs As Short, ByVal brdIdx As Short) As Integer
    '    'Public Declare Function GPBGetDlm Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GPBGetDlm@8" (ByRef btdlm As Short, ByVal brdIdx As Short) As Integer
    '    'Public Declare Function GPBGetTimeout Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GPBGetTimeout@8" (ByRef wtim As Short, ByVal brdIdx As Short) As Integer
    '    'Public Declare Function GPBIfc Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GPBIfc@8" (ByVal wtim As Short, ByVal brdIdx As Short) As Integer
    '    'Public Declare Function GPBInit Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GPBInit@4" (ByVal brdIdx As Short) As Integer
    '    'Public Declare Function GPBRecvData Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GPBRecvData@16" (ByRef btdata As Short, ByVal wsize As Short, ByRef wrecv As Short, ByVal brdIdx As Short) As Integer
    '    'Public Declare Function GPBSendCmd Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GPBSendCmd@12" (ByVal btcmd As String, ByVal wsize As Short, ByVal brdIdx As Short) As Integer
    '    'Public Declare Function GPBSendData Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GPBSendData@16" (ByVal btdata As String, ByVal wsize As Short, ByVal weoi As Short, ByVal brdIdx As Short) As Integer
    '    'Public Declare Function GPBSetDlm Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GPBSetDlm@8" (ByVal btdlm As Short, ByVal brdIdx As Short) As Integer
    '    'Public Declare Function GPBSetTimeout Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GPBSetTimeout@8" (ByVal wtim As Short, ByVal brdIdx As Short) As Integer
    '    'Public Declare Function GPBWaitForSRQ Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GPBWaitForSRQ@8" (ByVal timeout As Short, ByVal brdIdx As Short) As Integer
    '    'Public Declare Function GPERecv Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GPERecv@28" (ByVal bttlk As Short, ByRef btlsns As Short, ByVal wlsnnum As Short, ByRef btmsge As Short, ByVal wsize As Short, ByRef wrecv As Short, ByVal brdIdx As Short) As Integer
    '    'Public Declare Function GPESend Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GPESend@20" (ByRef btlsns As Short, ByVal wlsnnum As Short, ByVal btmsge As String, ByVal wsize As Short, ByVal brdIdx As Short) As Integer
    '    'Public Declare Function GPSGetSrqTkn Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GPSGetSrqTkn@8" (ByRef hSrqSem As Integer, ByVal brdIdx As Short) As Integer
    '    'Public Declare Function GPSInit Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GPSInit@0" () As Integer
    '    'Public Declare Function GPSLExeSRQ Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GPSLExeSRQ@12" (ByVal weoi As Short, ByVal wdevst As Short, ByVal brdIdx As Short) As Integer
    '    'Public Declare Function GPSLock Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GPSLock@4" (ByVal timeout As Short) As Integer
    '    'Public Declare Function GPSUnlock Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GPSUnlock@0" () As Integer
    '    Public Declare Function HCUT2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_HCUT2@40" (ByVal LTDIR As Short, ByVal L1 As Double, ByVal L2 As Double, ByVal L3 As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    '    Public Declare Function IACLEAR Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_IACLEAR@0" () As Integer
    '    Public Declare Function ICLEAR Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ICLEAR@4" (ByVal GADR As Short) As Integer
    '    Public Declare Function ICUT2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ICUT2@24" (ByVal n As Short, ByVal L As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    '    Public Declare Function IDELIM Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_IDELIM@4" (ByVal DLM As Short) As Integer
    '    Public Declare Function ILUM Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ILUM@4" (ByVal sw As Short) As Integer
    '    Public Declare Function InitFunction Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_InitFunction@0" () As Integer
    '    Public Declare Function INP16 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_INP16@8" (ByVal ADR As Integer, ByRef DAT As Integer) As Integer

    '    'Public Declare Function INtimeGWInitialize Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_INtimeGWInitialize@0" () As Integer
    '    'Public Declare Function INtimeGWTerminate Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_INtimeGWTerminate@0" () As Integer
    '    Public Declare Function IREAD Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_IREAD@8" (ByVal GADR As Short, ByVal DAT As String) As Integer
    '    Public Declare Function IREAD2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_IREAD2@12" (ByVal GADR As Short, ByVal GADR2 As Short, ByVal DAT As String) As Integer
    '    Public Declare Function IREADM Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_IREADM@16" (ByVal GADR As Short, ByRef MAX As Short, ByRef DAT As String, ByVal DLM As String) As Integer
    '    Public Declare Function IRHVAL Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_IRHVAL@12" (ByVal GADR As Short, ByVal HED As Short, ByRef DAT As Double) As Integer
    '    Public Declare Function IRHVAL2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_IRHVAL2@16" (ByVal GADR As Short, ByVal GADR2 As Short, ByVal HED As Short, ByRef DAT As Double) As Integer
    '    Public Declare Function IRMVAL Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_IRMVAL@16" (ByVal GADR As Short, ByRef MAX As Short, ByRef DAT As Double, ByRef DLM As String) As Integer
    '    Public Declare Function IRVAL Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_IRVAL@8" (ByVal GADR As Short, ByRef DAT As Double) As Integer
    '    Public Declare Function IRVAL2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_IRVAL2@12" (ByVal GADR As Short, ByVal GADR2 As Short, ByRef DAT As Double) As Integer
    '    Public Declare Function ITIMEOUT Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ITIMEOUT@4" (ByVal tim As Short) As Integer
    '    Public Declare Function ITIMESET Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ITIMESET@4" (ByVal MODE As Short) As Integer
    '    Public Declare Function ITRIGGER Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ITRIGGER@4" (ByVal GADR As Short) As Integer
    '    Public Declare Function IWRITE Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_IWRITE@8" (ByVal GADR As Short, ByVal DAT As String) As Integer
    '    Public Declare Function IWRITE2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_IWRITE2@12" (ByVal GADR As Short, ByVal GADR2 As Short, ByVal DAT As String) As Integer
    '    Public Declare Function LASEROFF Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_LASEROFF@0" () As Integer
    '    Public Declare Function LASERON Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_LASERON@0" () As Integer
    '    Public Declare Function LATTSET Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_LATTSET@8" (ByVal FAT As Integer, ByVal RAT As Integer) As Integer
    '    Public Declare Function LCUT2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_LCUT2@32" (ByVal LTDIR As Short, ByVal L1 As Double, ByVal L2 As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    '    Public Declare Function RANGE_SET Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_RANGE_SET@8" (ByVal MSDEV As String, ByVal rangeNo As Integer) As Integer
    '    Public Declare Function MFSET Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_MFSET@4" (ByVal MSDEV As String) As Integer
    '    '�i�V�KIF�j���d���Ή���
    '    Public Declare Function MFSET_EX Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_MFSET_EX@12" (ByVal MSDEV As String, ByVal target As Double) As Integer

    '    Public Declare Function MSCAN Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_MSCAN@28" (ByVal HP As Short, ByVal LP As Short, ByVal GP1 As Short, ByVal GP2 As Short, ByVal GP3 As Short, ByVal GP4 As Short, ByVal GP5 As Short) As Integer
    '    Public Declare Function NO_OPERATION Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_NO_OPERATION@20" (ByRef X As Double, ByRef y As Double, ByRef z As Double, ByRef BPx As Double, ByRef BPy As Double) As Integer
    '    Public Declare Function GET_STATUS Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GET_STATUS@20" (ByRef X As Double, ByRef y As Double, ByRef z As Double, ByRef BPx As Double, ByRef BPy As Double) As Integer
    '    Public Declare Function OUT16 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_OUT16@8" (ByVal ADR As Integer, ByVal DAT As Integer) As Integer
    '    Public Declare Function OUTBIT Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_OUTBIT@12" (ByVal CATEGORY As Short, ByVal BITNUM As Short, ByVal BON As Short) As Integer
    '    Public Declare Function PIN16 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_PIN16@8" (ByVal ADR As Integer, ByRef DAT As Integer) As Integer
    '    Public Declare Function PROBOFF Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_PROBOFF@0" () As Integer
    '    Public Declare Function PROBOFF_EX Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_PROBOFF_EX@8" (ByVal Pos As Double) As Integer
    '    Public Declare Function PROBOFF2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_PROBOFF2@0" () As Integer
    '    Public Declare Function PROBON Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_PROBON@0" () As Integer
    '    Public Declare Function PROBON2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_PROBON2@8" (ByVal Z2ON As Double) As Integer
    '    Public Declare Function PROBUP Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_PROBUP@8" (ByVal UP As Double) As Integer
    '    Public Declare Function PROCPOWER Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_PROCPOWER@4" (ByVal POWER As Short) As Integer
    '    Public Declare Sub PROP_SET Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_PROP_SET@48" (ByVal ZON As Double, ByVal ZOFF As Double, ByVal POSX As Double, ByVal POSY As Double, ByVal SmaxX As Double, ByVal SmaxY As Double)
    '    Public Declare Function QRATE Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_QRATE@8" (ByVal QR As Double) As Integer
    '    Public Declare Function RangeCorrect Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_RangeCorrect@24" (ByVal IDX As Short, ByVal Val_Renamed As Double, ByVal Flg As Short, ByVal RMin As Short, ByVal RMax As Short) As Integer
    '    Public Declare Function RATIO2EXP Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_RATIO2EXP@8" (ByVal RNO As Integer, ByVal MKSTR As String) As Integer
    '    Public Declare Function RBACK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_RBACK@0" () As Integer
    '    Public Declare Function RESET_Renamed Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_RESET@0" () As Integer
    '    Public Declare Function RINIT Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_RINIT@0" () As Integer
    '    Public Declare Function RMEAS Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_RMEAS@12" (ByVal MODE As Short, ByVal DVM As Short, ByRef r As Double) As Integer
    '    Public Declare Function RMeasHL Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_RMeasHL@28" (ByVal HP As Short, ByVal LP As Short, ByVal MODE As Short, ByVal NOM As Double, ByRef r As Double, ByRef ad As Short) As Integer
    '    Public Declare Function ROUND Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ROUND@4" (ByVal PLS As Integer) As Integer
    '    Public Declare Function ROUND4 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ROUND4@8" (ByVal ANG As Double) As Integer
    '    Public Declare Function RTEST Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_RTEST@36" (ByVal NOM As Double, ByVal MODE As Short, ByVal LOW As Double, ByVal HIGH As Double, ByVal JM As Short, ByVal DVM As Short) As Integer
    '    Public Declare Function RTRACK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_RTRACK@12" (ByVal NOM As Double, ByVal JM As Short) As Integer
    '    Public Declare Function SBACK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_SBACK@0" () As Integer
    '    Public Declare Function SETDLY Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_SETDLY@4" (ByVal DTIME As Integer) As Integer
    '    Public Declare Function SLIDECOVERCHK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_SLIDECOVERCHK@4" (ByVal CHK As Short) As Integer
    '    Public Declare Function SMOVE Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_SMOVE@16" (ByVal XD As Double, ByVal YD As Double) As Integer
    '    Public Declare Function SMOVE2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_SMOVE2@16" (ByVal XP As Double, ByVal YP As Double) As Integer
    '    Public Declare Function SMOVE_EX Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_SMOVE_EX@20" (ByVal XD As Double, ByVal YD As Double, ByVal OnOff As Short) As Integer
    '    Public Declare Function SMOVE2_EX Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_SMOVE2_EX@24" (ByVal XP As Double, ByVal YP As Double, ByVal OnOff As Short, ByVal jogMode As Integer) As Integer
    '    'Public Declare Function SMOVE2_EX Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_SMOVE2_EX@20" (ByVal XP As Double, ByVal YP As Double, ByVal OnOff As Short) As Integer
    '    '    Public Declare Function START Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_START@20" (ByVal Z1 As Short, ByVal XOFF As Double, ByVal YOFF As Double) As Integer
    '    Public Declare Function START Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_START@16" (ByVal XOFF As Double, ByVal YOFF As Double) As Integer
    '    Public Declare Function STCUT Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_STCUT@60" (ByVal L As Double, ByVal V As Double, ByVal NOM As Double, ByVal CUTOFF As Double, ByVal V2 As Double, ByVal Q2 As Double, ByVal CUT_DIR As Short, ByVal CUTMODE As Short, ByVal MODE As Short) As Integer
    '    Public Declare Function SYSINIT Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_SYSINIT@16" (ByVal ZOFF As Double, ByVal ZON As Double) As Integer
    '    Public Declare Function TEST Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TEST@36" (ByVal X As Double, ByVal NOM As Double, ByVal MODE As Short, ByVal LOW As Double, ByVal HIGH As Double) As Integer
    '    Public Declare Function TRIM_NGMARK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIM_NGMARK@32" (ByVal POSX As Double, ByVal POSY As Double, ByVal TM As Short, ByVal SN As Short, ByVal sw As Short, ByVal Flg As Short) As Integer
    '    'Public Declare Function TRIM_RESULT_WORD Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIM_RESULT@24" (ByVal KD As Short, ByVal SN As Short, ByVal NM As Short, ByVal CI As Short, ByVal DI As Short, ByRef Res() As UShort) As Integer
    '    Public Declare Function TRIM_RESULT_WORD Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIM_RESULT@24" (ByVal KD As Short, ByVal SN As Short, ByVal NM As Short, ByVal CI As Short, ByVal DI As Short, ByRef Res As UShort) As Integer
    '    Public Declare Function TRIM_RESULT_Double Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIM_RESULT@24" (ByVal KD As Short, ByVal SN As Short, ByVal NM As Short, ByVal CI As Short, ByVal DI As Short, ByRef Res As Double) As Integer
    '    'Public Declare Function TRIM_RESULT_Double Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIM_RESULT@24" (ByVal KD As Short, ByVal SN As Short, ByVal NM As Short, ByVal CI As Short, ByVal DI As Short, ByRef Res As TRIM_RES_Double) As Integer
    '    Public Declare Function TRIM80 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIM80@24" (ByVal X As Double, ByVal y As Double, ByVal V As Double) As Integer


    '    ' �u���b�N�P�ʂ̃g���~���O����
    '    Public Declare Function TRIMBLOCK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMBLOCK@20" (ByVal MD As Short, ByVal HZ As Short, ByVal RI As Short, ByVal CI As Short, ByVal NG As Short) As Integer
    '    ' �v���[�g�f�[�^���M
    '    Public Declare Function TRIMDATA_PLATE Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA_PLATE@8" (ByRef msg As TRIM_PLATE_DATA, ByVal tkyKnd As Integer) As Integer
    '    '    Public Declare Function TRIMDATA_GPIB Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA_GPIB@8" (ByRef msg As TRIM_DAT_GPIB, ByRef sts As S_RES_DAT) As Integer
    '    Public Declare Function TRIMDATA_GPIB Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA_GPIB@8" (ByRef msg As TRIM_PLATE_GPIB, ByVal tkyKnd As Integer) As Integer
    '    Public Declare Function TRIMDATA_RESISTOR Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA_RESISTOR@8" (ByRef msg As TRIM_RESISTOR_DATA, ByVal resNo As Integer) As Integer
    '    Public Declare Function TRIMDATA_CUTDATA Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA_CUTDATA@12" (ByRef msg As TRIM_CUT_DATA, ByVal resNo As Integer, ByVal cutNo As Integer) As Integer
    '    ' �J�b�g�p�����[�^���M
    '    Public Declare Function TRIMDATA_CutST Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA_CUTPRM@12" (ByRef msg As PRM_CUT_ST, ByVal resNo As Integer, ByVal cutNo As Integer) As Integer
    '    Public Declare Function TRIMDATA_CutL Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA_CUTPRM@12" (ByRef msg As PRM_CUT_L, ByVal resNo As Integer, ByVal cutNo As Integer) As Integer
    '    Public Declare Function TRIMDATA_CutHK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA_CUTPRM@12" (ByRef msg As PRM_CUT_HOOK, ByVal resNo As Integer, ByVal cutNo As Integer) As Integer
    '    Public Declare Function TRIMDATA_CutIX Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA_CUTPRM@12" (ByRef msg As PRM_CUT_INDEX, ByVal resNo As Integer, ByVal cutNo As Integer) As Integer
    '    Public Declare Function TRIMDATA_CutSC Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA_CUTPRM@12" (ByRef msg As PRM_CUT_SCAN, ByVal resNo As Integer, ByVal cutNo As Integer) As Integer
    '    Public Declare Function TRIMDATA_CutMK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA_CUTPRM@12" (ByRef msg As PRM_CUT_MARKING, ByVal resNo As Integer, ByVal cutNo As Integer) As Integer
    '    Public Declare Function TRIMDATA_CutES Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA_CUTPRM@12" (ByRef msg As PRM_CUT_ES, ByVal resNo As Integer, ByVal cutNo As Integer) As Integer


    '    ''Public Declare Function TRIMBLOCK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMBLOCK@24" (ByVal MD As Short, ByVal HZ As Short, ByVal RI As Short, ByVal CI As Short, ByVal NG As Short, ByRef sts As S_RES_DAT) As Integer
    '    ''Public Declare Function TRIMBLOCK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMBLOCK@24" (ByVal MD As Short, ByVal HZ As Short, ByVal CI As Short, ByVal NG As Short, ByRef sts As S_RES_DAT) As Integer
    '    'Public Declare Function TRIMBLOCK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMBLOCK@16" (ByVal MD As Short, ByVal HZ As Short, ByVal CI As Short, ByVal NG As Short) As Integer

    '    'Public Declare Function TRIMDATA_PLATE Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA_PLATE@8" (ByRef msg As TRIM_DAT_PLATE, ByVal tkyKnd As Integer) As Integer
    '    ''    Public Declare Function TRIMDATA_GPIB Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA_GPIB@8" (ByRef msg As TRIM_DAT_GPIB, ByRef sts As S_RES_DAT) As Integer
    '    'Public Declare Function TRIMDATA_RESISTOR Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA_RESISTOR@8" (ByRef msg As TRIM_DAT_RESISTOR, ByVal resNo As Integer) As Integer
    '    'Public Declare Function TRIMDATA_CUTDATA Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA_CUTDATA@12" (ByRef msg As TRIM_DAT_CUT, ByVal resNo As Integer, ByVal cutNo As Integer) As Integer
    '    'Public Declare Function TRIMDATA_CUTPRM Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA_CUTPRM@12" (ByRef msg As TRIM_DAT_CUT, ByVal resNo As Integer, ByVal cutNo As Integer) As Integer

    '    'Public Declare Function TRIMDATA_CutST Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA@8" (ByRef msg As TRIM_DAT_CUT_ST, ByRef sts As S_RES_DAT) As Integer
    '    'Public Declare Function TRIMDATA_CutL Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA@8" (ByRef msg As TRIM_DAT_CUT_L, ByRef sts As S_RES_DAT) As Integer
    '    'Public Declare Function TRIMDATA_CutHK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA@8" (ByRef msg As TRIM_DAT_CUT_HOOK, ByRef sts As S_RES_DAT) As Integer
    '    'Public Declare Function TRIMDATA_CutIX Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA@8" (ByRef msg As TRIM_DAT_CUT_INDEX, ByRef sts As S_RES_DAT) As Integer
    '    'Public Declare Function TRIMDATA_CutSC Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA@8" (ByRef msg As TRIM_DAT_CUT_SCAN, ByRef sts As S_RES_DAT) As Integer
    '    'Public Declare Function TRIMDATA_CutMK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA@8" (ByRef msg As TRIM_DAT_CUT_MARKING, ByRef sts As S_RES_DAT) As Integer
    '    'Public Declare Function TRIMDATA_CutC Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA@8" (ByRef msg As TRIM_DAT_CUT_C, ByRef sts As S_RES_DAT) As Integer
    '    'Public Declare Function TRIMDATA_CutES Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA@8" (ByRef msg As TRIM_DAT_CUT_ES, ByRef sts As S_RES_DAT) As Integer
    '    'Public Declare Function TRIMDATA_CutES2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA@8" (ByRef msg As TRIM_DAT_CUT_ES2, ByRef sts As S_RES_DAT) As Integer
    '    'Public Declare Function TRIMDATA_CutZ Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMDATA@8" (ByRef msg As TRIM_DAT_CUT_Z, ByRef sts As S_RES_DAT) As Integer
    '    Public Declare Function TRIMEND Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIMEND@0" () As Integer
    '    'Public Declare Function TSTEP Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TSTEP@8" (ByVal BNX As Short, ByVal BNY As Short) As Integer
    '    Public Declare Function TSTEP Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TSTEP@24" (ByVal BNX As Short, ByVal BNY As Short, ByVal stepOffX As Double, ByVal stepOffY As Double) As Integer
    '    Public Declare Function UCUT2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_UCUT2@40" (ByVal LTDIR As Short, ByVal L1 As Double, ByVal L2 As Double, ByVal RADI As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    '    Public Declare Function UCUT_PARAMSET Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_UCUT_PARAMSET@24" (ByVal MD As Short, ByVal KD As Short, ByVal RNO As Short, ByVal IDX As Short, ByVal EL As Short, ByRef pstPRM As UCUT_PARAM_EL) As Integer
    '    Public Declare Function UCUT_RESULT Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_UCUT_RESULT@16" (ByVal RNO As Short, ByVal CNO As Short, ByRef UcutNO As Short, ByRef InitVal As Double) As Integer
    '    Public Declare Function UCUT4RESULT Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_UCUT4RESULT@8" (ByRef sRegNo_p As Short, ByRef sCutNo_p As Short) As Integer
    '    Public Declare Function VCIRTRIM Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_VCIRTRIM@44" (ByVal SLP As Short, ByVal NOM As Double, ByVal V As Double, ByVal RADI As Double, ByVal ANG2 As Double, ByVal ANG As Double) As Integer
    '    Public Declare Function VCTRIM Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_VCTRIM@64" (ByVal SLP As Short, ByVal NOM As Double, ByVal MD As Short, ByVal X As Double, ByVal y As Double, ByVal VX As Double, ByVal VY As Double, ByVal LIMX As Double, ByVal LIMY As Double) As Integer
    '    Public Declare Function VHTRIM2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_VHTRIM2@64" (ByVal SLP As Short, ByVal NOM As Double, ByVal MD As Short, ByVal LTP As Double, ByVal LTDIR As Short, ByVal L1 As Double, ByVal L2 As Double, ByVal L3 As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    '    Public Declare Function VITRIM2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_VITRIM2@40" (ByVal SLP As Short, ByVal NOM As Double, ByVal MD As Short, ByVal n As Short, ByVal L As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    '    Public Declare Function VLTRIM2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_VLTRIM2@56" (ByVal SLP As Short, ByVal NOM As Double, ByVal MD As Short, ByVal LTP As Double, ByVal LTDIR As Short, ByVal L1 As Double, ByVal L2 As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    '    Public Declare Function VMEAS Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_VMEAS@12" (ByVal MODE As Short, ByVal DVM As Short, ByRef V As Double) As Integer
    '    Public Declare Function VRangeCorrect Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_VRangeCorrect@24" (ByVal IDX As Short, ByVal Val_Renamed As Double, ByVal Flg As Short, ByVal RMin As Short, ByVal RMax As Short) As Integer
    '    Public Declare Function VTEST Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_VTEST@36" (ByVal NOM As Double, ByVal MODE As Short, ByVal LOW As Double, ByVal HIGH As Double, ByVal JM As Short, ByVal DVM As Short) As Integer
    '    Public Declare Function VTRACK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_VTRACK@16" (ByVal SLP As Short, ByVal NOM As Double, ByVal JM As Short) As Integer
    '    Public Declare Function VUTRIM2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_VUTRIM2@64" (ByVal SLP As Short, ByVal NOM As Double, ByVal MD As Short, ByVal LTP As Double, ByVal LTDIR As Short, ByVal L1 As Double, ByVal L2 As Double, ByVal RADI As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    '    Public Declare Function VUTRIM4 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_VUTRIM4@88" (ByVal SLP As Short, ByVal NOM As Double, ByVal MD As Short, ByVal LTP As Double, ByVal LTDIR As Short, ByVal L1 As Double, ByVal L2 As Double, ByVal RADI As Double, ByVal V As Double, ByVal ANG As Short, ByVal trmd As Short, ByVal trl As Double, ByVal cn As Short, ByVal DT As Short, ByVal MODE As Short) As Integer
    '    '(2011/06/03)
    '    '   ���g�p�̈׍폜����
    '    'Public Declare Function XYOFF Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_XYOFF@16" (ByVal XOFF As Double, ByVal YOFF As Double) As Integer
    '    Public Declare Function ZABSVACCUME Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZABSVACCUME@4" (ByVal ZON As Integer) As Integer
    '    Public Declare Function ZATLDGET Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZATLDGET@4" (ByRef LDIN As Integer) As Integer
    '    Public Declare Function ZATLDSET Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZATLDSET@8" (ByVal LDON As Integer, ByVal LDOFF As Integer) As Integer
    '    Public Declare Function ZBPLOGICALCOORD Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZBPLOGICALCOORD@4" (ByVal COORD As Integer) As Integer
    '    Public Declare Function ZCONRST Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZCONRST@0" () As Integer
    '    Public Declare Function ZGETBPPOS Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZGETBPPOS@8" (ByRef XP As Double, ByRef YP As Double) As Integer
    '    Public Declare Function ZGETDCVRANG Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZGETDCVRANG@4" (ByRef VMAX As Double) As Integer
    '    Public Declare Function ZGETPHPOS Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZGETPHPOS@8" (ByRef NOWXP As Double, ByRef NOWYP As Double) As Integer

    '    Public Declare Function ZGETSRVSIGNAL Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZGETSRVSIGNAL@16" (ByRef X As Integer, ByRef y As Integer, ByRef z As Integer, ByRef t As Integer) As Integer
    '    'Public Declare Function ZGETTRMPOS Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZGETTRMPOS@24" (ByRef TRIMX As Double, ByRef TRIMY As Double, ByRef RCX As Double, ByRef RCY As Double, ByRef SMAX As Double, ByRef SMAY As Double) As Integer
    '    Public Declare Function ZINPSTS Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZINPSTS@8" (ByVal sw As Integer, ByRef sts As Integer) As Integer
    '    Public Declare Function ZLATCHOFF Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZLATCHOFF@0" () As Integer
    '    Public Declare Function ZZMOVE Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZMOVE@12" (ByVal z As Double, ByVal MD As Short) As Integer
    '    Public Declare Function ZZMOVE2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZMOVE2@12" (ByVal z As Double, ByVal MD As Short) As Integer
    '    Public Declare Function ZRCIRTRIM Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZRCIRTRIM@44" (ByVal NOM As Double, ByVal RNG As Short, ByVal V As Double, ByVal RADI As Double, ByVal ANG2 As Double, ByVal ANG As Double) As Integer
    '    Public Declare Function ZRTRIM2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZRTRIM2@32" (ByVal NOM As Double, ByVal RNG As Short, ByVal L As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    '    Public Declare Function ZSELXYZSPD Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZSELXYZSPD@4" (ByVal SPD As Integer) As Integer
    '    Public Declare Function ZSETBPTIME Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZSETBPTIME@8" (ByVal BPTIME As Integer, ByVal EPTIME As Integer) As Integer
    '    Public Declare Function ZSETPOS2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZSETPOS2@24" (ByVal POS2X As Double, ByVal POS2Y As Double, ByVal POS2Z As Double) As Integer
    '    Public Declare Function ZSETUCUT Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZSETUCUT@40" (ByVal MD As Short, ByVal RNO As Short, ByVal Index As Short, ByVal EL As Short, ByVal RATIO As Double, ByVal LTP As Double, ByVal LTP2 As Double) As Integer
    '    Public Declare Function ZSLCOVERCLOSE Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZSLCOVERCLOSE@4" (ByVal ZONOFF As Short) As Integer
    '    Public Declare Function ZSLCOVEROPEN Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZSLCOVEROPEN@4" (ByVal ZONOFF As Short) As Integer
    '    Public Declare Function ZSTGXYMODE Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZSTGXYMODE@4" (ByVal MODE As Integer) As Integer
    '    Public Declare Function ZSTOPSTS Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZSTOPSTS@0" () As Integer
    '    Public Declare Function ZSTOPSTS2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZSTOPSTS2@0" () As Integer
    '    '    Public Declare Function ZSYSPARAM1 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZSYSPARAM1@52" (ByVal POWERCYCLE As Short, ByVal THETA As Short, ByVal BPDIRXY As Short, ByVal BPSIZE As Short, ByVal DCSCANNER As Short, ByVal DCVRANGE As Short, ByVal LRANGE As Short, ByVal LDPOSX As Double, ByVal LDPOSY As Double, ByVal FPSUP As Short, ByVal DELAYSKIP As Short) As Integer
    '    Public Declare Function ZSYSPARAM1 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZSYSPARAM1@54" (ByVal POWERCYCLE As Short, ByVal THETA As Short, ByVal BPDIRXY As Short, ByVal BPSIZE As Short, ByVal DCSCANNER As Short, ByVal DCVRANGE As Short, ByVal LRANGE As Short, ByVal LDPOSX As Double, ByVal LDPOSY As Double, ByVal FPSUP As Short, ByVal DELAYSKIP As Short, ByVal OSC As Short) As Integer
    '    'Public Declare Function ZSYSPARAM2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZSYSPARAM2@60" (ByVal PRBTYP As Short, ByVal SMINMAXZ2 As Double, ByVal ZPTIMEON As Short, ByVal ZPTIMEOFF As Short, ByVal XYTBL As Short, ByVal SmaxX As Double, ByVal SmaxY As Double, ByVal ABSTIME As Integer, ByVal TRIMX As Double, ByVal TRIMY As Double) As Integer
    '    Public Declare Function ZSYSPARAM2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZSYSPARAM2@68" (ByVal PRBTYP As Short, ByVal SMINMAXZ2 As Double, ByVal ZPTIMEON As Short, ByVal ZPTIMEOFF As Short, ByVal XYTBL As Short, ByVal SmaxX As Double, ByVal SmaxY As Double, ByVal ABSTIME As Integer, ByVal TRIMX As Double, ByVal TRIMY As Double, ByVal BpMoveLimX As Integer, ByVal BpMoveLimY As Integer) As Integer
    '    'Public Declare Function ZSYSPARAM3 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZSYSPARAM3@16" (ByVal ProcPower2 As Short, ByVal GrvTime As Integer, ByVal UcutType As Short, ByVal ExtBit As Integer) As Integer
    '    'Public Declare Function ZSYSPARAM3 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZSYSPARAM3@20" (ByVal ProcPower2 As Short, ByVal GrvTime As Integer, ByVal UcutType As Short, ByVal ExtBit As Integer, ByVal PosSpd As Integer) As Integer '###021
    '    Public Declare Function ZSYSPARAM3 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZSYSPARAM3@24" (ByVal ProcPower2 As Short, ByVal GrvTime As Integer, ByVal UcutType As Short, ByVal ExtBit As Integer, ByVal PosSpd As Integer, ByVal BiasOn_AddTime As Integer) As Integer
    '    Public Declare Function ZTIMERINIT Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZTIMERINIT@0" () As Integer
    '    Public Declare Function ZVMEAS Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZVMEAS@12" (ByVal MODE As Short, ByVal DVM As Short, ByRef V As Double) As Integer
    '    Public Declare Function ZWAIT Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZWAIT@4" (ByVal lngWaitMilliSec As Integer) As Integer
    '    Public Declare Function ZZGETRTMODULEINFO Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZZGETRTMODULEINFO@0" () As Integer
    '    Public Declare Function Z_INIT Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_Z_INIT@0" () As Integer
    '    'About TRIMMING
    '    Public Declare Function ZRANGTRIM Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ZRANGTRIM@32" (ByVal NOM As Double, ByVal RNG As Short, ByVal L As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    '    Public Declare Function VTRIM2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_VTRIM2@32" (ByVal SLP As Short, ByVal NOM As Double, ByVal L As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    '    Public Declare Function CUT2 Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_CUT2@20" (ByVal L As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    '    Public Declare Function CMARK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_CMARK@40" (ByVal MKSTR As String, ByVal STX As Double, ByVal STY As Double, ByVal HIGH As Double, ByVal V As Double, ByVal ANG As Short) As Integer
    '    Public Declare Function TrimMK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIM_MK@52" (ByVal MKSTR As String, ByVal STX As Double, ByVal STY As Double, ByVal HIGH As Double, ByVal V As Double, ByVal ANG As Short, ByVal QRate1 As Double, ByVal condNoCut1 As Short) As Integer

    '    '�V�KI/F
    '    '    Public Declare Function TRIM_ST Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIM_ST@76" (ByVal MOVEMODE As Integer, ByVal CUTMODE As Integer, ByVal POS As Integer, ByVal SLP As Integer, ByVal NOM As Double, ByVal L As Double, ByVal V As Double, ByVal V_RET As Double, ByVal ANG As Integer, ByVal QRATE As Double, ByVal QRATE_RET As Double, ByVal CUTCOND_NO As Integer, ByVal CUTCOND_NO_RET As Integer) As Long
    '    'Public Declare Function TRIM_ST Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIM_ST@60" (ByVal MOVEMODE As short, ByVal CUTMODE As short, ByVal SLP As short, ByVal NOM As Double, ByVal L As Double, ByVal V As Double, ByVal V_RET As Double, ByVal ANG As short, ByVal QRATE As Double, ByVal QRATE_RET As Double, ByVal CUTCOND_NO As short, ByVal CUTCOND_NO_RET As short) As  Integer
    '    'Public Declare Function TRIM_L Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIM_ST@116" (ByVal MOVEMODE As short, ByVal CUTMODE As short, ByVal SLP As short, ByVal NOM As Double, ByVal MD As short, ByVal LTP As Double, ByVal LTDIR As short, ByVal L1 As Double, ByVal L2 As Double, ByVal V As Double, ByVal V2 As Double, ByVal V_RET As Double, ByVal V_RET2 As Double, ByVal ANG As short, ByVal QRATE As Double, ByVal QRATE2 As Double, ByVal QRATE_RET As Double, ByVal QRATE_RET2 As Double, ByVal CUTCOND_NO As short, ByVal CUTCOND_NO2 As short, ByVal CUTCOND_NO_RET As short, ByVal CUTCOND_NO_RET2 As short) As Integer 
    '    'Public Declare Function TRIM_HkU Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIM_HkU@168" (ByVal MOVEMODE As short, ByVal CUTMODE As short, ByVal SLP As short, ByVal NOM As Double, ByVal MD As short, ByVal LTP As Double, ByVal LTDIR As short, ByVal L1 As Double, ByVal L2 As Double, ByVal L3 As Double, ByVal RADI As Double, ByVal V1 As Double, ByVal V2 As Double, ByVal V3 As Double, ByVal V1_RET As Double, ByVal V2_RET As Double, ByVal V3_RET As Double, ByVal ANG As short, ByVal QRATE1 As Double, ByVal QRATE2 As Double, ByVal QRATE3 As Double, ByVal QRATE1_RET As Double, ByVal QRATE2_RET As Double, ByVal QRATE3_RET As Double, ByVal CUTCOND_NO1 As short, ByVal CUTCOND_NO2 As short, ByVal CUTCOND_NO3 As short, ByVal CUTCOND_NO1_RET As short, ByVal CUTCOND_NO2_RET As short, ByVal CUTCOND_NO3_RET As short) As Integer 
    '    Public Declare Function TRIM_ST Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIM_ST@4" (ByRef CutCmnPrm As CUT_COMMON_PRM) As Integer
    '    Public Declare Function TRIM_L Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIM_L@4" (ByRef CutCmnPrm As CUT_COMMON_PRM) As Integer
    '    Public Declare Function TRIM_HkU Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIM_HkU@4" (ByRef CutCmnPrm As CUT_COMMON_PRM) As Integer
    '    Public Declare Function TRIM_ES Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIM_ES@4" (ByRef CutCmnPrm As CUT_COMMON_PRM) As Integer
    '    Public Declare Function TRIM_IX Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TRIM_IX@4" (ByRef CutCmnPrm As CUT_COMMON_PRM) As Integer
    '    Public Declare Function MEASURE Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_MEASURE@28" (ByVal MEASMODE As Short, ByVal RANGSETTYPE As Short, ByVal MEASTYPE As Short, ByVal TARGET As Double, ByVal RANGE As Short, ByRef RESULT As Double) As Integer
    '    Public Declare Function FLSET Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_FLSET@8" (ByVal mode As Short, ByVal cutCondNo As Short) As Integer
    '    Public Declare Function SET_FL_ERRLOG Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_SET_FL_ERRLOG@4" (ByRef ErrCode As Integer) As Integer

    '    Public Declare Function SYSTEM_RESET Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_SYSTEM_RESET@0" () As Integer
    '    Public Declare Function SERVO_POWER Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_SERVO_POWER@16" (ByVal XAxisOnOff As Integer, ByVal YAxisOnOff As Integer, ByVal ZAxisOnOff As Integer, ByVal TAxisOnOff As Integer) As Integer
    '    Public Declare Function CLEAR_SERVO_ALARM Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_CLEAR_SERVO_ALARM@8" (ByVal XY As Integer, ByVal ZT As Integer) As Integer
    '    Public Declare Function AXIS_X_INIT Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_AXIS_X_INIT@0" () As Integer
    '    Public Declare Function AXIS_Y_INIT Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_AXIS_Y_INIT@0" () As Integer
    '    Public Declare Function AXIS_Z_INIT Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_AXIS_Z_INIT@0" () As Integer
    '    Public Declare Function GET_ALLAXIS_STATUS Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GET_ALLAXIS_STATUS@8" (ByRef err As Long, ByRef AllStatus As Long) As Integer
    '    Public Declare Function LAMP_CTRL Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_LAMP_CTRL@8" (ByVal LampNo As Integer, ByVal OnOff As Boolean) As Integer
    '    Public Declare Function COVERLATCH_CLEAR Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_COVERLATCH_CLEAR@0" () As Integer
    '    Public Declare Function COVERLATCH_CHECK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_COVERLATCH_CHECK@4" (ByRef LatchSts As Long) As Integer
    '    Public Declare Function COVER_CHECK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_COVER_CHECK@4" (ByRef SwitchSts As Long) As Integer
    '    Public Declare Function INTERLOCK_CHECK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_INTERLOCK_CHECK@8" (ByRef InterlockSts As Integer, ByRef SwitchSts As Long) As Integer
    '    Public Declare Function ORG_INTERLOCK_CHECK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ORG_INTERLOCK_CHECK@8" (ByRef InterlockSts As Integer, ByRef SwitchSts As Long) As Integer
    '    Public Declare Function SLIDECOVER_MOVINGCHK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_SLIDECOVER_MOVINGCHK@12" (ByVal OpenCloseChk As Integer, ByVal UseReset As Integer, ByRef SwitchSts As Long) As Integer
    '    Public Declare Function SLIDECOVER_CLOSECHK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_SLIDECOVER_CLOSECHK@4" (ByRef slidecoverSts As Long) As Integer
    '    Public Declare Function SLIDECOVER_GETSTS Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_SLIDECOVER_GETSTS@4" (ByRef slidecoverSts As Long) As Integer
    '    Public Declare Function START_SWWAIT Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_START_SWWAIT@4" (ByRef SwitchSts As Long) As Integer
    '    Public Declare Function START_SWCHECK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_START_SWCHECK@8" (ByVal bReleaseCheck As Integer, ByRef SwitchSts As Long) As Integer
    '    Public Declare Function HALT_SWCHECK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_HALT_SWCHECK@4" (ByRef SwitchSts As Long) As Integer
    '    Public Declare Function STARTRESET_SWWAIT Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_STARTRESET_SWWAIT@4" (ByRef SwitchSts As Long) As Integer
    '    Public Declare Function ORG_STARTRESET_SWWAIT Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ORG_STARTRESET_SWWAIT@4" (ByRef SwitchSts As Long) As Integer
    '    Public Declare Function STARTRESET_SWCHECK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_STARTRESET_SWCHECK@8" (ByVal bReleaseCheck As Integer, ByRef SwitchSts As Long) As Integer
    '    Public Declare Function GET_Z_POS Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GET_Z_POS@4" (ByRef ZPos As Double) As Integer
    '    Public Declare Function GET_QRATE Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GET_QRATE@4" (ByRef QRate As Double) As Integer
    '    Public Declare Function CONSOLE_SWCHECK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_CONSOLE_SWCHECK@8" (ByVal BbReleaseCheck As Boolean, ByRef SwitchChk As Long) As Integer
    '    Public Declare Function Z_SWCHECK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_Z_SWCHECK@4" (ByRef SwitchChk As Long) As Integer
    '    Public Declare Function EMGSTS_CHECK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_EMGSTS_CHECK@4" (ByRef Status As Integer) As Long
    '    Public Declare Function ISALIVE_INTIME Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_ISALIVE_INTIME@0" () As Integer
    '    Public Declare Function TERMINATE_INTIME Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_TERMINATE_INTIME@0" () As Integer
    '    Public Declare Function BP_GET_CALIBDATA Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_BP_GET_CALIBDATA@16" (ByRef gainX As Double, ByRef gainY As Double, ByRef offsetX As Double, ByRef offsetY As Double) As Integer
    '    Public Declare Function SIGNALTOWER_CTRLEX Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_SIGNALTOWER_CTRLEX@8" (ByVal OnBit As Integer, ByVal OffBit As Integer) As Integer

    '    '�f�o�b�O/���u�]���p�R�}���h
    '    Public Declare Function SETLOG_ALLTARGET Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_SETLOG_ALLTARGET@36" (ByVal base As Short, ByVal io As Short, ByVal laser As Short, ByVal bp As Short, ByVal meas As Short, ByVal trim As Short, ByVal correct As Short, ByVal stage As Short, ByVal loader As Short) As Integer
    '    'Public Declare Function SETLOG_ALLTARGET Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_SETLOG_ALLTARGET@12" (ByRef AllStatus[] As UInteger) As Integer
    '    Public Declare Function SETLOG_TARGET Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_SETLOG_TARGET@8" (ByVal segNo As Integer, ByVal status As UInteger) As Integer
    '    'Public Declare Function GETLOG_TARGET Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_GETLOG_TARGET@" () As Integer
    '    Public Declare Function PERFORMCHK Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_PERFORMCHK@12" (ByVal ADDR As UInteger, ByVal COUNT As UInteger, ByVal WAIT As UInteger) As Integer
    '    Public Declare Function SETAXISSPD Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_SETAXISSPD@24" (ByVal XL As UInteger, ByVal XH As UInteger, ByVal YL As UInteger, ByVal YH As UInteger, ByVal ZL As UInteger, ByVal ZH As UInteger) As Integer
    '    Public Declare Function LSI_RESET Lib "C:\DevOnly\Trimmer\Update\Source\Modules\DLL\DllTrimFunc\debug\DllTrimFnc.dll" Alias "_LSI_RESET@0" () As Integer

    '#End If
#End Region
End Module