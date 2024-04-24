'===============================================================================
'   Description : �u���b�N�P�ʂ̃g���~���O����
'
'   Copyright(C) Laser Front 2010
'
'===============================================================================
Option Strict Off
Option Explicit On
Module Trimming
#Region "�O���[�o���萔/�ϐ��̒�`"
    '-------------------------------------------------------------------------------
    '   �萔��`
    '-------------------------------------------------------------------------------
    '----- �ő�l/�ŏ��l -----
    Public Const cMAXcMARKINGcSTRLEN As Integer = 18        ' �}�[�L���O������ő咷(byte)
    Public Const cCNDNUM As Integer = 4                     ' 1��Ă̍ő���H������(FL�p)
    Public Const cResultMax As Integer = 256                ' �g���~���O���ʃf�[�^�̍ő�z��
    Public Const cResultAry As Integer = 999                ' �g���~���O���ʃf�[�^�̍ő吔

    '----- ���o�� -----
    Public Const INP_MAX As Integer = 5                     ' ��Signal��Ԃ̐�
    Public Const INP_ICSLSS As Integer = 0                  ' [0]:�R���\�[��SW�Z���X
    Public Const INP_IITLKS As Integer = 1                  ' [1]:�C���^�[���b�N�֌WSW�Z���X
    Public Const INP_AUTLODL As Integer = 2                 ' [2]:�I�[�g���[�_LO
    Public Const INP_AUTLODH As Integer = 3                 ' [3]:�I�[�g���[�_HI
    Public Const INP_ATTNATE As Integer = 4                 ' [4]:�Œ�A�b�e�l�[�^

    Public Const OUT_MAX As Integer = 4                     ' ��Signal��Ԃ̐�
    Public Const OUT_OCSLLN As Integer = 0                  ' [0]:�R���\�[������
    Public Const OUT_OSYSCTL As Integer = 1                 ' [1]:�T�[�{�p���[
    Public Const OUT_AUTLODL As Integer = 2                 ' [2]:�I�[�g���[�_LO
    Public Const OUT_AUTLODH As Integer = 3                 ' [3]:�I�[�g���[�_HI
    Public Const OUT_SIGNALT As Integer = 4                 ' [4]:�V�O�i���^���[(���g�p)
    Public Const OUT_Z2CONT As Integer = 6                  ' [5]:Z2�T�[�{�p���[

    '----- �g���~���O�v���f�[�^�̃f�[�^�^�C�v -----
    Public Const DATTYPE_PLATE As UShort = 1                ' �v���[�g�f�[�^
    Public Const DATTYPE_REGI As UShort = 2                 ' ��R�f�[�^
    Public Const DATTYPE_CUT As UShort = 3                  ' �J�b�g�f�[�^
    Public Const DATTYPE_PARAM As UShort = 4                ' �J�b�g�p�����[�^
    Public Const DATTYPE_GPIB As UShort = 8                 ' GPIB�ݒ�p�f�[�^

    '-------------------------------------------------------------------------------
    '   �J�b�g�^�C�v�ʃp�����[�^�`����`(VB��INtime)
    '-------------------------------------------------------------------------------
    '----- ST cut -----
    Public Structure PRM_CUT_ST                             ' ST cut�p�����[�^�`����`
        Dim DIR As UShort                                   ' �J�b�g����(1:+X, 2:-X, 3:+Y, 4:-Y)
        Dim MODE As UShort                                  ' ���샂�[�h(0:NOM, 1:���^�[��, 2:���g���[�X, 3:�΂�)
        Dim angle As UShort                                 ' �΂߃J�b�g�p�x(0�`359)
        Dim Length As Double                                ' �ő�J�b�e�B���O��(0.0001�`20.0000(mm))
        Dim spd2 As Double                                  ' ���H�X�s�[�h(mm/s)
        Dim qrate2 As Double                                ' ���^�[��/���g���[�X��Qrate2(KHz)
        'Dim chenge As Double                                ' �؂�ւ��|�C���g(0.0�`100.0%)(SL436K�p)
    End Structure

    '----- L cut -----
    Public Structure PRM_CUT_L                              ' L cut�p�����[�^�`����`
        Dim DIR As UShort                                   ' �J�b�g����(1:+X, 2:-X, 3:+Y, 4:-Y)
        Dim tdir As UShort                                  ' L�^�[������(1:+X, 2:-X, 3:+Y, 4:-Y)
        Dim MODE As UShort                                  ' ���샂�[�h(0:NOM, 1:���^�[��, 2:���g���[�X, 3:�΂�)
        Dim angle As UShort                                 ' �΂߃J�b�g�p�x(0�`359)
        Dim turn As Double                                  ' L�^�[���|�C���g(0.0�`100.0(%))
        Dim L1 As Double                                    ' L1 �ő�J�b�e�B���O��(0.0001�`20.0000(mm))
        Dim L2 As Double                                    ' L2 �ő�J�b�e�B���O��(0.0001�`20.0000(mm))
        Dim r As Double                                     ' �^�[���̉~�ʔ��a(mm)
        Dim spd2 As Double                                  ' ���H�X�s�[�h(mm/s)
        Dim qrate2 As Double                                ' Qrate2(KHz)
        <VBFixedArray(1)> Dim qrate3() As Double            ' Qrate3(KHz) FL��������/��ڰ�����Qrate
        '                                                   ' Qrate3[0]: ���H�����ԍ�3��QڰĂ�ݒ�(L��ݑO��Qrate)
        '                                                   ' Qrate3[1]: ���H�����ԍ�4��QڰĂ�ݒ�(L��݌��Qrate)
        <VBFixedArray(2)> Dim spd3() As Double              ' FL����L��݌�/����/��ڰ����̃X�s�[�h(mm/s)
        '                                                   ' Spd3[0]: L��݌�̽�߰��
        '                                                   ' Spd3[1]: ����/��ڰ�����L��ݑO�̽�߰��
        '                                                   ' Spd3[2]: ����/��ڰ�����L��݌�̽�߰��
        ' ���̍\���̂�����������ɂ́A"Initialize" ���Ăяo���Ȃ���΂Ȃ�܂���B 
        Public Sub Initialize()
            ReDim qrate3(1)
            ReDim spd3(2)
        End Sub
    End Structure

    '----- HOOK cut -----
    Public Structure PRM_CUT_HOOK                           ' HOOK cut�p�����[�^�`����`
        Dim DIR As UShort                                   ' �J�b�g����(1:+X, 2:-X, 3:+Y, 4:-Y)
        Dim tdir As UShort                                  ' L�^�[������(1:+X, 2:-X, 3:+Y, 4:-Y)
        Dim turn As Double                                  ' L�^�[���|�C���g(0.0�`100.0(%))
        Dim L1 As Double                                    ' L1 �ő�J�b�e�B���O��(0.0001�`20.0000(mm))
        Dim r1 As Double                                    ' �^�[��1�̉~�ʔ��a(mm)
        Dim L2 As Double                                    ' L2 �ő�J�b�e�B���O��(0.0001�`20.0000(mm))
        Dim r2 As Double                                    ' �^�[��2�̉~�ʔ��a(mm)
        Dim L3 As Double                                    ' L3 �ő�J�b�e�B���O��(0.00001�`20.0000(mm))
        <VBFixedArray(1)> Dim qrate2() As Double            ' Qrate2(KHz) FL����L2/L3��Qrate
        '                                                   ' Qrate2[0]: ���H�����ԍ�2��QڰĂ�ݒ�(L2��Qrate)
        '                                                   ' Qrate2[1]: ���H�����ԍ�3��QڰĂ�ݒ�(L3��Qrate)
        <VBFixedArray(1)> Dim spd2() As Double              ' FL����L2/L3�̃X�s�[�h(mm/s)
        '                                                   ' Spd2[0]: L2�̽�߰��
        '                                                   ' Spd2[1]: L3�̽�߰��
        ' ���̍\���̂�����������ɂ́A"Initialize" ���Ăяo���Ȃ���΂Ȃ�܂���B 
        Public Sub Initialize()
            ReDim qrate2(1)
            ReDim spd2(1)
        End Sub
    End Structure

    '----- INDEX cut -----
    Public Structure PRM_CUT_INDEX                          ' INDEX cut�p�����[�^�`����`
        Dim DIR As UShort                                   ' �J�b�g����(1:+X, 2:-X, 3:+Y, 4:-Y)
        Dim maxindex As UShort                              ' �C���f�b�N�X��(1�`32767)
        Dim measure As UShort                               ' ���胂�[�h(0:����, 1:�����x)
        Dim Length As Double                                ' �C���f�b�N�X��(0.0001�`20.0000(mm))
    End Structure

    '----- SCAN cut -----
    Public Structure PRM_CUT_SCAN                           ' SCAN cut�p�����[�^�`����`
        Dim DIR As UShort                                   ' �J�b�g����(1:+X, 2:-X, 3:+Y, 4:-Y)
        Dim sdir As UShort                                  ' �X�e�b�v����(1:+X, 2:-X, 3:+Y, 4:-Y)
        Dim lines As UShort                                 ' �{��(1�`n)
        Dim Length As Double                                ' �J�b�e�B���O��(0.0001�`20.0000(mm))
        Dim pitch As Double                                 ' �s�b�`(0.0001�`20.0000(mm))
    End Structure

    '----- Letter Marking -----
    Public Structure PRM_CUT_MARKING                        ' Letter Marking�p�����[�^�`����`
        '                                                   ' ����
        <VBFixedArray(cMAXcMARKINGcSTRLEN - 1)> Dim str() As Byte
        Dim magnify As Double                               ' �{��(�P�`999)
        Dim DIR As UShort                                   ' �����̌���(1:0, 2:90, 3:180, 4:270)
        ' ���̍\���̂�����������ɂ́A"Initialize" ���Ăяo���Ȃ���΂Ȃ�܂���B 
        Public Sub Initialize()
            ReDim str(cMAXcMARKINGcSTRLEN - 1)
        End Sub
    End Structure

    '----- C cut -----
    Public Structure PRM_CUT_C                              ' C cut�p�����[�^�`����`
        Dim DIR As UShort                                   ' �J�b�g����(1:CW, 2:CCW)
        Dim angle As UShort                                 ' �J�b�g�p�x(0�`359)
        Dim count As UShort                                 ' ��
        Dim st_r As Double                                  ' �~�ʔ��a (mm)
        Dim pitch As Double                                 ' �s�b�`
    End Structure

    '----- ES cut -----
    Public Structure PRM_CUT_ES                             ' ES cut�p�����[�^�`����`
        Dim DIR As UShort                                   ' �J�b�g����  1:+X, 2:-X, 3:+Y, 4:-Y
        Dim L1 As Double                                    ' L1 �ő�J�b�e�B���O��(0.0001�`20.0000(mm))
        Dim EsPoint As Double                               ' ES�߲��(-99.9999�`0.0000%))
        Dim ESchangerate As Double                          ' ES����ω���(0.0�`100.0%))
        Dim EScutlen As Double                              ' ES�㶯Ē�(0.0001�`20.0000(mm))
    End Structure

    '----- ES2 cut -----
    Public Structure PRM_CUT_ES2                            ' ES2 cut�p�����[�^�`����`
        Dim DIR As UShort                                   ' �J�b�g����(1:+X, 2:-X, 3:+Y, 4:-Y)
        Dim L1 As Double                                    ' L1 �ő�J�b�e�B���O��(0.0001�`20.0000(mm))
        Dim EsPoint As Double                               ' ES�߲��(-99.9999�`0.0000%)
        Dim ESWide As Double                                ' ES����ω���(0.0�`100.0%)
        Dim ESWide2 As Double                               ' ES��ω���(0.0�`100.0%)
        Dim EScount As UShort                               ' ES��m�F��(0�`20)
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
    Public Structure S_UCUTPARAM                            ' UCUT�p�����[�^�`����`
        <VBFixedArray(19)> Dim EL() As S_UCUTPARAM_EL       ' UCUT�p�����[�^ 
        ' ���̍\���̂�����������ɂ́A"Initialize" ���Ăяo���Ȃ���΂Ȃ�܂���B 
        Public Sub Initialize()
            ReDim EL(19)
        End Sub
    End Structure

    '-------------------------------------------------------------------------------
    '   �J�b�g�f�[�^�`����`(VB��INtime)
    '-------------------------------------------------------------------------------
    Public Structure PRM_CUT_DATA                           ' �J�b�g�f�[�^�`����`
        Dim CP1 As UShort                                   ' �J�b�g�ԍ� 1-20
        Dim CP2 As UShort                                   ' ��d������㑪��x������(0-32767msec) 
        Dim CP3 As UShort                                   ' �J�b�g�`��(1:st, 2:L, 3:HK, 4:IX ��)
        Dim cp4_x As Double                                 ' �J�b�g�X�^�[�g���WX(-80.0000�`+80.0000)
        Dim cp4_y As Double                                 ' �J�b�g�X�^�[�g���WY(-80.0000�`+80.0000)
        Dim CP5 As Double                                   ' �J�b�g�X�s�[�h(0.1�`409.0mm/s)
        Dim CP6 As Double                                   ' ���[�U�[Q�X�C�b�`���[�g(0.1�`50.0KHz) ��FL���͉��H�����ԍ�1��QڰĂ�ݒ�
        Dim CP7 As Double                                   ' �J�b�g�I�t %(-99.999 �` +999.999)
        Dim CP71 As Double                                  ' �J�b�g�f�[�^���ω���(0.0�`100.0, 0%)(���g�p)
        <VBFixedArray(cCNDNUM - 1)> Dim CP72() As Byte      ' ���H�����ԍ�1�`4(FL�p) 
        'Dim CP50 As UShort                                  ' �p���X������(0:���� 1:�L��)(SL436K�p)
        'Dim CP51 As Double                                  ' �p���X������(SL436K�p)
        'Dim CP52 As Double                                  ' LSw�p���X������(�O���V���b�^)(SL436K�p)
        Dim dummy As PRM_CUT_HOOK                           ' �J�b�g�p�����[�^(union) ��union��`�Ȃ��̂ōő�̂��̂��w��

        ' ���̍\���̂�����������ɂ́A"Initialize" ���Ăяo���Ȃ���΂Ȃ�܂���B 
        Public Sub Initialize()
            ReDim CP72(cCNDNUM - 1)
        End Sub
    End Structure

    '-------------------------------------------------------------------------------
    '   ��R�f�[�^�`����`(VB��INtime)
    '-------------------------------------------------------------------------------
    Public Structure PRM_REGISTER                           ' ��R�f�[�^�`����`
        Dim PR1a As UShort                                  ' ��R�ԍ�(1-999=�g���~���O, 1000-9999=�}�[�L���O)
        Dim PR2a As UShort                                  ' ���葪��(0:����, 1:�����x
        Dim PR3a As UShort                                  ' �T�[�L�b�g(��R��������T�[�L�b�g�ԍ�)
        Dim PR4_ha As UShort                                ' �n�C���v���[�u�ԍ�
        Dim PR4_la As UShort                                ' ���[���v���[�u�ԍ�
        Dim PR4_g1a As UShort                               ' ��1�A�N�e�B�u�K�[�h�ԍ�
        Dim PR4_g2a As UShort                               ' ��2�A�N�e�B�u�K�[�h�ԍ�
        Dim PR4_g3a As UShort                               ' ��3�A�N�e�B�u�K�[�h�ԍ�
        Dim PR4_g4a As UShort                               ' ��4�A�N�e�B�u�K�[�h�ԍ�
        Dim PR4_g5a As UShort                               ' ��5�A�N�e�B�u�K�[�h�ԍ�
        Dim PR5a As UInteger                                ' External bits
        Dim PR6a As UShort                                  ' �|�[�Y�^�C��(External bits�o�͌�̃E�F�C�g) (msec)
        Dim PR7a As UShort                                  ' �ڕW�l�w��(0:��Βl, 1:���V�I, 2:�v�Z��)
        Dim PR8a As UShort                                  ' �x�[�X��RNo.(���V�I���̊��R�ԍ�)
        Dim PR9a As Double                                  ' �g���~���O�ڕW�l(ohm)
        Dim PR10a As UShort                                 ' �d���ω��X���[�v(0:+�X���[�v, 1:-�X���[�v) ����ڰ��ް��̑��胂�[�h=�d���̏ꍇ�L��
        Dim PR11_Ha As Double                               ' IT Limit H(-99.99�`9999.99%)
        Dim PR11_La As Double                               ' IT Limit L(-99.99�`9999.99%)
        Dim PR12_Ha As Double                               ' FT Limit H(-99.99�`9999.99%)
        Dim PR12_La As Double                               ' FT Limit L(-99.99�`9999.99%)
        Dim PR13a As UShort                                 ' �J�b�g��(1�`20)
        Dim PR14 As UShort                                  ' �J�b�g�ʒu�␳�t���O(0:�␳���Ȃ�, 1:�␳����)
        Dim PR14_Ha As Double                               ' �C�j�V����OK�e�X�gHIGH���~�b�g(SL436K�p)
        Dim PR14_La As Double                               ' �C�j�V����OK�e�X�gLOW���~�b�g (SL436K�p)
        Dim fCutMag As Double                               ' �؏グ�{��(CHIP�̂�)
        Dim pCutData As UInteger                            ' �J�b�g�f�[�^�|�C���^(INTIME���Ŏg�p)
    End Structure

    '-------------------------------------------------------------------------------
    '   �v���[�g�f�[�^�`����`(VB��INtime)
    '-------------------------------------------------------------------------------
    Public Structure TRIM_PLATE_DATA                        ' �v���[�g�f�[�^�`����`
        Dim wCircuitCnt As UShort                           ' �T�[�L�b�g��
        Dim wRegistCnt As UShort                            ' ��R��
        Dim wTrimMode As UShort                             ' ���胂�[�h(0:��R, 1:�d��)
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
        Dim wInitialOK As UShort                            ' �Ƽ��OKýėL��(0:���� 1:�L��))(SL436K�p)
        Dim wNGMark As UShort                               ' NGϰ�ݸނ���/���Ȃ�)(SL436K�p)
        Dim w4Terminal As UShort                            ' 4�[�q�������������/���Ȃ�)(SL436K�p)
        Dim wLogMode As UShort                              ' ۷�ݸ�Ӱ��
        '                                                   ' 0:���Ȃ�, 1:INITIAL TEST, 2:FINAL TEST, 3:INITIAL + FINAL)	
    End Structure

    '-------------------------------------------------------------------------------
    '   GPIB�ݒ�p�f�[�^�`����`(VB��INtime)
    '-------------------------------------------------------------------------------
    Public Structure TRIM_PLATE_GPIB                        ' GPIB�ݒ�p�f�[�^�`����`
        Dim wGPIBmode As UShort                             ' GP-IB����(0:���Ȃ� 1:����)
        Dim wDelim As UShort                                ' �����(0:CR+LF 1:CR 2:LF 3:NONE)
        Dim wTimeout As UShort                              ' ��ѱ��(0�`1000)(100ms�P��)
        Dim wAddress As UShort                              ' �@����ڽ(0�`30)
        <VBFixedArray(39)> Dim strI() As Byte               ' �����������(MAX40byte)
        <VBFixedArray(9)> Dim strT() As Byte                ' �ض޺����(10byte)
        <VBFixedArray(5)> Dim wReserve() As Byte            ' �\��(6byte)  
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
    '----- �g���~���O�v���f�[�^(�v���[�g�f�[�^) -----
    Public Structure TRIM_DAT_PLATE                         ' �g���~���O�v���f�[�^(�v���[�g�f�[�^)�`����`
        Dim cmdNo As UInteger                               ' �R�}���hNo.(DllTrimFnc�Őݒ肷��̂Ŗ��g�p)
        Dim type As UShort                                  ' �f�[�^�^�C�v(1:�v���[�g�f�[�^)
        Dim index_reg As UShort                             ' ��R�f�[�^�E�C���f�b�N�X
        Dim index_cut As UShort                             ' �J�b�g�f�[�^�E�C���f�b�N�X
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET���(0:TKY, 1:CHIP, 2:NET)
        Dim prmPlate As TRIM_PLATE_DATA                     ' �v���[�g�f�[�^
    End Structure

    '----- �g���~���O�v���f�[�^(GPIB�ݒ�p�f�[�^) -----
    Public Structure TRIM_DAT_GPIB                          ' �g���~���O�v���f�[�^(GPIB�ݒ�p�f�[�^)�`����`
        Dim cmdNo As UInteger                               ' �R�}���hNo.(DllTrimFnc�Őݒ肷��̂Ŗ��g�p)
        Dim type As UShort                                  ' �f�[�^�^�C�v(8:GPIB�f�[�^)
        Dim index_reg As UShort                             ' ��R�f�[�^�E�C���f�b�N�X
        Dim index_cut As UShort                             ' �J�b�g�f�[�^�E�C���f�b�N�X
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET���(0:TKY, 1:CHIP, 2:NET)
        Dim prmGPIB As TRIM_PLATE_GPIB                      ' GPIB�ݒ�p�f�[�^
    End Structure

    '----- �g���~���O�v���f�[�^(��R�f�[�^) -----
    Public Structure TRIM_DAT_REGI                          ' �g���~���O�v���f�[�^(��R�f�[�^)�`����`
        Dim cmdNo As UInteger                               ' �R�}���hNo.(DllTrimFnc�Őݒ肷��̂Ŗ��g�p)
        Dim type As UShort                                  ' �f�[�^�^�C�v(2:��R�f�[�^)
        Dim index_reg As UShort                             ' ��R�f�[�^�E�C���f�b�N�X
        Dim index_cut As UShort                             ' �J�b�g�f�[�^�E�C���f�b�N�X
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET���(0:TKY, 1:CHIP, 2:NET)
        Dim prmReg As PRM_REGISTER                          ' ��R�f�[�^
    End Structure

    '----- �g���~���O�v���f�[�^(�J�b�g�f�[�^) -----
    Public Structure TRIM_DAT_CUT                           ' �g���~���O�v���f�[�^(�J�b�g�f�[�^)�`����`
        Dim cmdNo As UInteger                               ' �R�}���hNo.(DllTrimFnc�Őݒ肷��̂Ŗ��g�p)
        Dim type As UShort                                  ' �f�[�^�^�C�v(3:�J�b�g�f�[�^)
        Dim index_reg As UShort                             ' ��R�f�[�^�E�C���f�b�N�X
        Dim index_cut As UShort                             ' �J�b�g�f�[�^�E�C���f�b�N�X
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET���(0:TKY, 1:CHIP, 2:NET)
        Dim prmCut As PRM_CUT_DATA                          ' �J�b�g�f�[�^
    End Structure

    '----- �g���~���O�v���f�[�^(ST cut/ST cut2�p�����[�^) -----
    Public Structure TRIM_DAT_CUT_ST                        ' �g���~���O�v���f�[�^(�J�b�g�p�����[�^)�`����`
        Dim cmdNo As UInteger                               ' �R�}���hNo.(DllTrimFnc�Őݒ肷��̂Ŗ��g�p)
        Dim type As UShort                                  ' �f�[�^�^�C�v(4:�J�b�g�p�����[�^)
        Dim index_reg As UShort                             ' ��R�f�[�^�E�C���f�b�N�X
        Dim index_cut As UShort                             ' �J�b�g�f�[�^�E�C���f�b�N�X
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET���(0:TKY, 1:CHIP, 2:NET)
        Dim c As PRM_CUT_ST                                 ' ST cut�p�����[�^
    End Structure

    '----- �g���~���O�v���f�[�^(L cut�p�����[�^) -----
    Public Structure TRIM_DAT_CUT_L                         ' �g���~���O�v���f�[�^(�J�b�g�p�����[�^)�`����`
        Dim cmdNo As UInteger                               ' �R�}���hNo.(DllTrimFnc�Őݒ肷��̂Ŗ��g�p)
        Dim type As UShort                                  ' �f�[�^�^�C�v(4:�J�b�g�p�����[�^)
        Dim index_reg As UShort                             ' ��R�f�[�^�E�C���f�b�N�X
        Dim index_cut As UShort                             ' �J�b�g�f�[�^�E�C���f�b�N�X
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET���(0:TKY, 1:CHIP, 2:NET)
        Dim c As PRM_CUT_L                                  ' L cut�p�����[�^
    End Structure

    '----- �g���~���O�v���f�[�^(HOOK cut�p�����[�^) -----
    Public Structure TRIM_DAT_CUT_HOOK                      ' �g���~���O�v���f�[�^(�J�b�g�p�����[�^)�`����`
        Dim cmdNo As UInteger                               ' �R�}���hNo.(DllTrimFnc�Őݒ肷��̂Ŗ��g�p)
        Dim type As UShort                                  ' �f�[�^�^�C�v(4:�J�b�g�p�����[�^)
        Dim index_reg As UShort                             ' ��R�f�[�^�E�C���f�b�N�X
        Dim index_cut As UShort                             ' �J�b�g�f�[�^�E�C���f�b�N�X
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET���(0:TKY, 1:CHIP, 2:NET)
        Dim c As PRM_CUT_HOOK                               ' HOOK cut�p�����[�^
    End Structure

    '----- �g���~���O�v���f�[�^(INDEX cut�p�����[�^) -----
    Public Structure TRIM_DAT_CUT_INDEX                     ' �g���~���O�v���f�[�^(�J�b�g�p�����[�^)�`����`
        Dim cmdNo As UInteger                               ' �R�}���hNo.(DllTrimFnc�Őݒ肷��̂Ŗ��g�p)
        Dim type As UShort                                  ' �f�[�^�^�C�v(4:�J�b�g�p�����[�^)
        Dim index_reg As UShort                             ' ��R�f�[�^�E�C���f�b�N�X
        Dim index_cut As UShort                             ' �J�b�g�f�[�^�E�C���f�b�N�X
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET���(0:TKY, 1:CHIP, 2:NET)
        Dim c As PRM_CUT_INDEX                              ' INDEX cut�p�����[�^
    End Structure

    '----- �g���~���O�v���f�[�^(SCAN cut�p�����[�^) -----
    Public Structure TRIM_DAT_CUT_SCAN                     ' �g���~���O�v���f�[�^(�J�b�g�p�����[�^)�`����`
        Dim cmdNo As UInteger                               ' �R�}���hNo.(DllTrimFnc�Őݒ肷��̂Ŗ��g�p)
        Dim type As UShort                                  ' �f�[�^�^�C�v(4:�J�b�g�p�����[�^)
        Dim index_reg As UShort                             ' ��R�f�[�^�E�C���f�b�N�X
        Dim index_cut As UShort                             ' �J�b�g�f�[�^�E�C���f�b�N�X
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET���(0:TKY, 1:CHIP, 2:NET)
        Dim c As PRM_CUT_SCAN                               ' SCAN cut�p�����[�^
    End Structure

    '----- �g���~���O�v���f�[�^(Letter Marking�p�����[�^) -----
    Public Structure TRIM_DAT_CUT_MARKING                   ' �g���~���O�v���f�[�^(�J�b�g�p�����[�^)�`����`
        Dim cmdNo As UInteger                               ' �R�}���hNo.(DllTrimFnc�Őݒ肷��̂Ŗ��g�p)
        Dim type As UShort                                  ' �f�[�^�^�C�v(4:�J�b�g�p�����[�^)
        Dim index_reg As UShort                             ' ��R�f�[�^�E�C���f�b�N�X
        Dim index_cut As UShort                             ' �J�b�g�f�[�^�E�C���f�b�N�X
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET���(0:TKY, 1:CHIP, 2:NET)
        Dim c As PRM_CUT_MARKING                            ' Letter Marking�p�����[�^
    End Structure

    '----- �g���~���O�v���f�[�^(C cut�p�����[�^) -----
    Public Structure TRIM_DAT_CUT_C                         ' �g���~���O�v���f�[�^(�J�b�g�p�����[�^)�`����`
        Dim cmdNo As UInteger                               ' �R�}���hNo.(DllTrimFnc�Őݒ肷��̂Ŗ��g�p)
        Dim type As UShort                                  ' �f�[�^�^�C�v(4:�J�b�g�p�����[�^)
        Dim index_reg As UShort                             ' ��R�f�[�^�E�C���f�b�N�X
        Dim index_cut As UShort                             ' �J�b�g�f�[�^�E�C���f�b�N�X
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET���(0:TKY, 1:CHIP, 2:NET)
        Dim c As PRM_CUT_C                                  ' C cut�p�����[�^
    End Structure

    '----- �g���~���O�v���f�[�^(ES cut�p�����[�^) -----
    Public Structure TRIM_DAT_CUT_ES                        ' �g���~���O�v���f�[�^(�J�b�g�p�����[�^)�`����`
        Dim cmdNo As UInteger                               ' �R�}���hNo.(DllTrimFnc�Őݒ肷��̂Ŗ��g�p)
        Dim type As UShort                                  ' �f�[�^�^�C�v(4:�J�b�g�p�����[�^)
        Dim index_reg As UShort                             ' ��R�f�[�^�E�C���f�b�N�X
        Dim index_cut As UShort                             ' �J�b�g�f�[�^�E�C���f�b�N�X
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET���(0:TKY, 1:CHIP, 2:NET)
        Dim c As PRM_CUT_ES                                 ' ES cut�p�����[�^
    End Structure

    '----- �g���~���O�v���f�[�^(ES2 cut�p�����[�^) -----
    Public Structure TRIM_DAT_CUT_ES2                       ' �g���~���O�v���f�[�^(�J�b�g�p�����[�^)�`����`
        Dim cmdNo As UInteger                               ' �R�}���hNo.(DllTrimFnc�Őݒ肷��̂Ŗ��g�p)
        Dim type As UShort                                  ' �f�[�^�^�C�v(4:�J�b�g�p�����[�^)
        Dim index_reg As UShort                             ' ��R�f�[�^�E�C���f�b�N�X
        Dim index_cut As UShort                             ' �J�b�g�f�[�^�E�C���f�b�N�X
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET���(0:TKY, 1:CHIP, 2:NET)
        Dim c As PRM_CUT_ES2                                ' ES2 cut�p�����[�^
    End Structure

    '----- �g���~���O�v���f�[�^(Z cut(NOP)�p�����[�^) -----
    Public Structure TRIM_DAT_CUT_Z                         ' �g���~���O�v���f�[�^(�J�b�g�p�����[�^)�`����`
        Dim cmdNo As UInteger                               ' �R�}���hNo.(DllTrimFnc�Őݒ肷��̂Ŗ��g�p)
        Dim type As UShort                                  ' �f�[�^�^�C�v(4:�J�b�g�p�����[�^)
        Dim index_reg As UShort                             ' ��R�f�[�^�E�C���f�b�N�X
        Dim index_cut As UShort                             ' �J�b�g�f�[�^�E�C���f�b�N�X
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET���(0:TKY, 1:CHIP, 2:NET)
    End Structure

    '-------------------------------------------------------------------------------
    '   �����f�[�^(�g���~���O���ʃf�[�^)�`����`(VB��INtime)
    '-------------------------------------------------------------------------------
    '----- �g���~���O���ʃf�[�^(WORD�^�f�[�^�p) -----
    Public Structure TRIM_RESULT_WORD                       ' �g���~���O���ʃf�[�^(WORD�^�f�[�^�p)�`����`
        Dim wTxSize As UShort                               ' �]���T�C�Y(DllTrimFnc�Őݒ肷��)
        <VBFixedArray(cResultMax - 1)> Dim wd() As UShort   ' ����(wd[0]�`wd[255])
        ' ���̍\���̂��g�p����ɂ�"Initialize"���Ăяo���Ȃ���΂Ȃ�Ȃ��B 
        Public Sub Initialize()
            ReDim wd(cResultMax - 1)
        End Sub
    End Structure

    '----- �g���~���O���ʃf�[�^(Double�^�f�[�^�p) -----
    Public Structure TRIM_RESULT_Double                     ' �g���~���O���ʃf�[�^(WORD�^�f�[�^�p)�`����`
        Dim wTxSize As UShort                               ' �]���T�C�Y(DllTrimFnc�Őݒ肷��)
        <VBFixedArray(cResultMax - 1)> Dim dd() As Double   ' ����(dd[0]�`dd[255])
        ' ���̍\���̂��g�p����ɂ�"Initialize"���Ăяo���Ȃ���΂Ȃ�Ȃ��B 
        Public Sub Initialize()
            ReDim dd(cResultMax - 1)
        End Sub
    End Structure

    '-------------------------------------------------------------------------------
    '   �v��/�����f�[�^(�R�}���h)�`����`(VB����INtime)
    '-------------------------------------------------------------------------------
    '----- �v���f�[�^(VB��INtime) -----
    Public Structure S_CMD_DAT                              ' �v���f�[�^�`����`
        Dim cmdNo As UInteger                               ' �R�}���hNo.(DllTrimFnc�Őݒ肷��̂Ŗ��g�p)
        <VBFixedArray(9)> Dim dbPara() As Double            ' double �^�p�����[�^(dbPara(0-9))
        <VBFixedArray(9)> Dim dwPara() As Integer           ' long	 �^�p�����[�^(dbPara(0-9))
        Dim flgTrim As UInteger                             ' 0:���ݸޒ��łȂ�, 1:���ݸޒ�(IRQ0�����֎~) 

        ' ���̍\���̂�����������ɂ́A"Initialize" ���Ăяo���Ȃ���΂Ȃ�܂���B 
        Public Sub Initialize()
            ReDim dbPara(9)
            ReDim dwPara(9)
        End Sub
    End Structure

    '----- ���V�I���[�h�Q�v�Z���f�[�^(VB��INtime) -----
    Public Structure S_RATIO2EXP                            ' ���V�I���[�h�Q�v�Z���f�[�^�`����`
        Dim cmdNo As UInteger                               ' �R�}���hNo.(DllTrimFnc�Őݒ肷��̂Ŗ��g�p)
        Dim RNO As UInteger                                 ' ��R�ԍ�
        Dim strExp As String                                ' �v�Z��������
    End Structure

    '----- �����f�[�^(�R�}���h)(VB��INtime) -----
    Public Structure S_RES_DAT                              ' �����f�[�^�`����`
        Dim status As Integer                               ' 0:����, 0�ȊO:�s���� (�������Ȃ��Ȃ̂�-1��ݒ肵�Ă���?)
        Dim dwerrno As Integer                              ' �G���[�ԍ�(0:����)
        <VBFixedArray(3)> Dim signal() As UInteger          ' ���X�e�[�^�X
        '                                                   ' [0]:X��
        '                                                   ' [1]:Y��
        '                                                   ' [2]:Z��
        '                                                   ' [3]:�Ǝ�
        '                                                   ' I/O���͏��
        <VBFixedArray(INP_MAX - 1)> Dim in_dat() As UInteger
        '                                                   ' [0]:�R���\�[��SW�Z���X
        '                                                   ' [1]:�C���^�[���b�N�֌WSW�Z���X
        '                                                   ' [2]:�I�[�g���[�_LO
        '                                                   ' [3]:�I�[�g���[�_HI
        '                                                   ' [4]:�Œ�A�b�e�l�[�^
        '                                                   ' I/O�o�͏��
        <VBFixedArray(OUT_MAX - 1)> Dim outdat() As UInteger
        '                                                   ' [0]:�R���\�[������
        '                                                   ' [1]:�T�[�{�p���[
        '                                                   ' [2]:�I�[�g���[�_LO
        '                                                   ' [3]:�I�[�g���[�_HI
        '                                                   ' [4]:�V�O�i���^���[(���g�p)
        <VBFixedArray(3)> Dim wData() As UInteger           ' TKY�ߒl
        <VBFixedArray(4)> Dim pos() As Double               ' ���݈ʒu
        '                                                   ' [0]:X��
        '                                                   ' [1]:Y��
        '                                                   ' [2]:Z��
        '                                                   ' [3]:BPX
        '                                                   ' [4]:BPY
        Dim fData As Double                                 ' �ߒl(����l��)

        ' ���̍\���̂�����������ɂ́A"Initialize" ���Ăяo���Ȃ���΂Ȃ�܂���B 
        Public Sub Initialize()
            ReDim signal(3)
            ReDim in_dat(INP_MAX - 1)
            ReDim outdat(OUT_MAX - 1)
            ReDim pos(4)
        End Sub
    End Structure

    '-------------------------------------------------------------------------------
    '   �v��/�����f�[�^��`
    '-------------------------------------------------------------------------------
    Public stSCMD As S_CMD_DAT                              ' �v���f�[�^(�R�}���h)(VB��INtime)
    Public stSRES As S_RES_DAT                              ' �����f�[�^(�R�}���h)(VB��INtime)

    '----- �g���~���O�v���f�[�^(VB��INtime) -----
    Public stTPLT As TRIM_DAT_PLATE                         ' �v���[�g�f�[�^
    Public stTGPI As TRIM_DAT_GPIB                          ' GPIB�ݒ�f�[�^
    Public stTREG As TRIM_DAT_REGI                          ' ��R�f�[�^
    Public stTCUT As TRIM_DAT_CUT                           ' �J�b�g�f�[�^
    '                                                       ' �J�b�g�p�����[�^ 
    Public stCutST As TRIM_DAT_CUT_ST                       ' ST cut�p�����[�^
    Public stCutL As TRIM_DAT_CUT_L                         ' L cut�p�����[�^
    Public stCutHK As TRIM_DAT_CUT_HOOK                     ' HOOK cut�p�����[�^
    Public stCutIX As TRIM_DAT_CUT_INDEX                    ' INDEX cut�p�����[�^
    Public stCutSC As TRIM_DAT_CUT_SCAN                     ' SCAN cut�p�����[�^
    Public stCutMK As TRIM_DAT_CUT_MARKING                  ' Letter Marking�p�����[�^
    Public stCutC As TRIM_DAT_CUT_C                         ' C cut�p�����[�^
    Public stCutES As TRIM_DAT_CUT_ES                       ' ES cut�p�����[�^
    Public stCutE2 As TRIM_DAT_CUT_ES2                      ' ES2 cut�p�����[�^
    Public stCutZ As TRIM_DAT_CUT_Z                         ' Z cut(NOP)�p�����[�^

    '----- �g���~���O���ʃf�[�^ -----
    Public stResultWd As TRIM_RESULT_WORD                   ' �g���~���O���ʃf�[�^(WORD�^�f�[�^�p)
    Public stResultDd As TRIM_RESULT_Double                 ' �g���~���O���ʃf�[�^(Double�^�f�[�^�p)

    Public gwTrimResult(cResultAry - 1) As UShort           ' ����(gwTrimResult[0]�`gwTrimResult[999])
    Public gfInitialTest(cResultAry - 1) As Double          ' IT����l(gwTrimResult[0]�`gwTrimResult[999])


#End Region

End Module
