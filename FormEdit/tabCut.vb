Option Explicit On
Option Strict Off

Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports LaserFront.Trimmer.DefWin32Fnc

Namespace FormEdit
    Friend Class tabCut
        Inherits tabBase

#Region "�׽"
#Region "NG��ėp"
        ''' <summary>NG��Ċ֘A�̺��۰ق̗L���������ؑւ���</summary>
        Private Class NGCut
            Friend m_CtlArr() As Control    ' NG��ĂŎg�p������۰�
            Friend WriteOnly Property Enabled() As Boolean
                Set(ByVal value As Boolean)
                    For Each ctl As Control In m_CtlArr
                        ctl.Enabled = value
                    Next
                End Set
            End Property
        End Class
#End Region

#Region "�������ݶ�ėp"
        ''' <summary>�������ݶ�Ċ֘A�̕\�����\���������Ȃ�</summary>
        Private Class Serpentine
            Friend m_ctlArr() As Control    ' �������ݶ�ĂŎg�p������۰�
            Friend m_lblArr() As Label      ' ��Ē�1/2,��ĕ���1/2������
            ' { {��Ē�       }, {��ĕ���(0�`360��)   } }
            ' { {��Ē�1/2(mm)}, {��ĕ���1/2(0�`360��)} }
            Friend m_strLbl(1, 1) As String ' ���قɕ\�����镶����

            Friend WriteOnly Property Visible() As Boolean
                Set(ByVal value As Boolean)
                    Dim tf As Integer = 0
                    If (True = value) Then tf = 1
                    For i As Integer = 0 To (m_lblArr.Length - 1) Step 1
                        m_lblArr(i).Text = m_strLbl(tf, i) ' ���قɕ������ݒ�
                    Next i

                    For Each ctl As Control In m_ctlArr
                        ctl.Visible = value ' ���۰ق̕\�����\����ݒ�
                    Next
                End Set
            End Property

        End Class
#End Region

#Region "��ď����p�׽"
        ''' <summary>FL���H�����ŕ\�����鶯ď�������ݒ肷��</summary>
        Private Class CutCondition
            Friend m_ctlArr(,) As Control ' �\�����\���������Ȃ����۰�

            ''' <summary>�����̒l���̂ݏ�����\�����A�ȍ~���\���ɂ���</summary>
            ''' <value>�\�����������</value>
            Friend WriteOnly Property Visible() As Integer
                Set(ByVal value As Integer)
                    Dim tf As Boolean
                    If (3 < value) Then value = 3
                    For i As Integer = 0 To (m_ctlArr.GetLength(0) - 1) Step 1
                        If (i < value) Then
                            tf = True
                        Else
                            tf = False
                        End If

                        For j As Integer = 0 To (m_ctlArr.GetLength(1) - 1) Step 1
                            m_ctlArr(i, j).Visible = tf
                        Next j
                    Next i
                End Set
            End Property

            ''' <summary>�����̒l���̂�FL�ݒ�No.�����ޯ����L���ɂ��A�ȍ~�𖳌��ɂ���</summary>
            ''' <value>�L���ɂ��������</value>
            Friend WriteOnly Property Enabled() As Integer
                Set(ByVal value As Integer)
                    Dim tf As Boolean
                    If (4 < value) Then value = 4
                    For i As Integer = 0 To (m_ctlArr.GetLength(0) - 1) Step 1
                        If (i < value) Then
                            tf = True
                        Else
                            tf = False
                        End If
                        m_ctlArr(i, 1).Enabled = tf ' �e�k�ݒ�m���D
                        'm_ctlArr(i, 4).Enabled = tf ' �r�s�d�f�{��
                        m_ctlArr(i, 5).Enabled = tf ' �ڕW�p���[
                        m_ctlArr(i, 6).Enabled = tf ' ���e�͈�
                    Next i
                End Set
            End Property

        End Class
#End Region
#End Region

#Region "�錾"
        Private Const CUT_CUT As Integer = 2        ' m_CtlCut�ł̲��ޯ��(��ĕ��@)
        Private Const CUT_CTYPE As Integer = 3      ' m_CtlCut�ł̲��ޯ��(��Č`��)
        Private Const CUT_QRATE As Integer = 5      ' m_CtlCut�ł̲��ޯ��(�p���[�g)
        Private Const CUT_SPEED As Integer = 6      ' m_CtlCut�ł̲��ޯ��(���x)
        Private Const CUT_START_X As Integer = 7    ' m_CtlCut�ł̲��ޯ��(�J�b�g�ʒu�w)
        Private Const CUT_START_Y As Integer = 8    ' m_CtlCut�ł̲��ޯ��(�J�b�g�ʒu�x)
        Private Const CUT_START_2_X As Integer = 9  ' m_CtlCut�ł̲��ޯ��(�J�b�g�ʒu�Q�w)
        Private Const CUT_START_2_Y As Integer = 10 ' m_CtlCut�ł̲��ޯ��(�J�b�g�ʒu�Q�x)
        Private Const CUT_LEN_1 As Integer = 11     ' m_CtlCut�ł̲��ޯ��(�J�b�g���P)
        Private Const CUT_LEN_2 As Integer = 12     ' m_CtlCut�ł̲��ޯ��(�J�b�g���Q)
        Private Const CUT_DIR_1 As Integer = 13     ' m_CtlCut�ł̲��ޯ��(�J�b�g����)
        Private Const CUT_DIR_2 As Integer = 14     ' m_CtlCut�ł̲��ޯ��(�J�b�g�����Q)
        Private Const CUT_OFF As Integer = 15       ' m_CtlCut�ł̲��ޯ��(�J�b�g�I�t)
        Private Const CUT_LTP As Integer = 16       ' m_CtlCut�ł̲��ޯ��(�k�^�[���|�C���g)
        Private Const CUT_MTYPE As Integer = 17     ' m_CtlCut�ł̲��ޯ��(����@��)
        Private Const CUT_TMM As Integer = 18       ' m_CtlCut�ł̲��ޯ��(����Ӱ��)
        Private Const CUT_LETTER As Integer = 19    '###1042�@ m_CtlCut�ł̲��ޯ��(����Ӱ��)

        'V2.2.1.7�@ ��
        Private Const CUT_MARK_FIX As Integer = 20   ' m_CtlCut�ł̲��ޯ��(�󎚌Œ蕔)
        Private Const CUT_ST_NUM As Integer = 21     ' m_CtlCut�ł̲��ޯ��(�J�n�ԍ�)
        Private Const CUT_REPEAT_CNT As Integer = 22 ' m_CtlCut�ł̲��ޯ��(�d����)

        Private Const CUT_VAR_REPEAT As Integer = 23 '###1042�@ m_CtlCut�ł̲��ޯ��(���s�[�g�L��)
        Private Const CUT_VARIATION As Integer = 24 '###1042�@ m_CtlCut�ł̲��ޯ��(����L��)
        Private Const CUT_RATE As Integer = 25      '###1042�@ m_CtlCut�ł̲��ޯ��(�㏸��)
        Private Const CUT_VAR_LO As Integer = 26    '###1042�@ m_CtlCut�ł̲��ޯ��(�����l)
        Private Const CUT_VAR_HI As Integer = 27    '###1042�@ m_CtlCut�ł̲��ޯ��(����l)
        Private Const CUT_RETRACE As Integer = 28      'V2.0.0.0�F���g���[�X�̖{�� 'V2.1.0.0�@ 20����J�b�g���̒�R�l�ω��ʔ��荀�ڂT���V�t�g��25

        ''V2.1.0.0�@��
        'Private Const CUT_VAR_REPEAT As Integer = 20 '###1042�@ m_CtlCut�ł̲��ޯ��(���s�[�g�L��)
        'Private Const CUT_VARIATION As Integer = 21 '###1042�@ m_CtlCut�ł̲��ޯ��(����L��)
        'Private Const CUT_RATE As Integer = 22      '###1042�@ m_CtlCut�ł̲��ޯ��(�㏸��)
        'Private Const CUT_VAR_LO As Integer = 23    '###1042�@ m_CtlCut�ł̲��ޯ��(�����l)
        'Private Const CUT_VAR_HI As Integer = 24    '###1042�@ m_CtlCut�ł̲��ޯ��(����l)
        ''V2.1.0.0�@��
        'Private Const CUT_RETRACE As Integer = 25      'V2.0.0.0�F���g���[�X�̖{�� 'V2.1.0.0�@ 20����J�b�g���̒�R�l�ω��ʔ��荀�ڂT���V�t�g��25
        'V2.2.1.7�@ ��

        'V2.0.0.0�F        Private Const CUT_TR_Q As Integer = 20      '###1042�@ m_CtlCut�ł̲��ޯ��(����Ӱ��)
        'V2.0.0.0�F        Private Const CUT_TR_SPEED As Integer = 21  '###1042�@ m_CtlCut�ł̲��ޯ��(����Ӱ��)

        Private Const IDX_MTYPE As Integer = 4      ' m_CtlIdxCut�ł�2�����ڂ̲��ޯ��(����@��)
        Private Const IDX_TMM As Integer = 5        ' m_CtlIdxCut�ł�2�����ڂ̲��ޯ��(����Ӱ��)

        'V1.0.4.3�B CNS_CUTM_TR�ɕύX        Private Const CMB_CUT_TRACK As Integer = 1  ' ��ĕ��@�����ޯ����ؽĲ��ޯ��+1(�ׯ�ݸ�)
        'V1.0.4.3�B CNS_CUTM_IX�ɕύX        Private Const CMB_CUT_IDX As Integer = 2    ' ��ĕ��@�����ޯ����ؽĲ��ޯ��+1(���ޯ��)
        'V1.0.4.3�B CNS_CUTM_NG�ɕύX        Private Const CMB_CUT_NG As Integer = 3     ' ��ĕ��@�����ޯ����ؽĲ��ޯ��+1(NG���)

        'V1.0.4.3�B CNS_CUTP_ST�ɕύX        Private Const CMB_CTYP_STR As Integer = 1   ' ��Č`������ޯ����ؽĲ��ޯ��+1(��ڰ�)
        'V1.0.4.3�B CNS_CUTP_L�ɕύX        Private Const CMB_CTYP_LCUT As Integer = 2  ' ��Č`������ޯ����ؽĲ��ޯ��+1(L���)
        'V1.0.4.3�B CNS_CUTP_SP�ɕύX        Private Const CMB_CTYP_SPT As Integer = 3   ' ��Č`������ޯ����ؽĲ��ޯ��+1(��������)

        Private m_CtlCut() As Control           ' ��ĸ�ٰ���ޯ���̺��۰ٔz��
        Private m_CtlIdxCut(,) As Control       ' ���ޯ����ĸ�ٰ���ޯ���̺��۰ٔz��
        Private m_CtlFLCnd(,) As Control        ' FL���H������ٰ���ޯ���̺��۰ٔz��
        Private m_CtlLCut(,) As Control         'V1.0.4.3�B L�J�b�g��ٰ���ޯ���̺��۰ٔz��
        Private m_CtlRetraceCut(,) As Control   'V2.0.0.0�F ���g���[�X�J�b�g��ٰ���ޯ���̺��۰ٔz��
        Private Const LCUT_DIR_IDX As Integer = 3      'V1.0.4.3�B m_CtlLCut�Ŋp�x�i�v���_�E���j�̂Q�����ڂ̔z��ԍ�

        Private m_CtlUCut() As Control         ' U�J�b�g�p�p�����[�^�ǉ�      'V2.2.0.0�A

        Private m_NGCut As NGCut                ' NG��Ċ֘A�̗L���������ؑւ���
        'Private m_Serpentine As Serpentine      ' �������ݶ�Ċ֘A�̕\�����\���������Ȃ�
        Private m_CutCondition As CutCondition  ' FL���H�����ŕ\�����鶯ď�������ݒ肷��

        Private Const MAX_DEGREES As Integer = 359
        Private AngleArray(MAX_DEGREES, 1) As Integer           'V1.0.4.3�A
        Private AngleArrayForLcut(MAX_DEGREES, 1) As Integer    'V1.0.4.3�A
        Private AngleArrayForUcut(MAX_DEGREES, 1) As Integer    'V2.2.0.0�A

        'V2.0.0.0��
        ''' <summary>
        ''' �덷�̃��x���\������
        ''' </summary>
        ''' <remarks></remarks>
        Private m_strDev As String

        ''' <summary>
        ''' �덷�䗦�w�莞�̍ŏ��l/�ő�l
        ''' </summary>
        ''' <remarks>0=�ŏ��l�A1=�ő�l</remarks>
        Private m_strDEVRaite(2) As String

        ''' <summary>
        ''' �덷��Βl�w�莞�̍ŏ��l/�ő�l
        ''' </summary>
        ''' <remarks>0=�ŏ��l�A1=�ő�l</remarks>
        Private m_strDEVAbsolute(2) As String

        ''' <summary>
        ''' ��ĕ��@�̃R���{�{�b�N�X�f�[�^���X�g
        ''' </summary>
        ''' <remarks></remarks>
        Private m_lstCutMethod As New List(Of ComboDataStruct)

        Private m_lstCutType As New List(Of ComboDataStruct)
        'V2.0.0.0��
#End Region

#Region "�ݽ�׸�"
        ''' <summary>�ݽ�׸�</summary>
        Friend Sub New(ByRef mainEdit As frmEdit, ByVal tabIdx As Integer)
            ' ���̌Ăяo���́AWindows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
            InitializeComponent()

            ' InitializeComponent() �Ăяo���̌�ŏ�������ǉ����܂��B
            Call InitAllControl(mainEdit, tabIdx)
        End Sub
#End Region

#Region "����������"
        ''' <summary>���۰ُ���������</summary>
        ''' <param name="mainEdit">ҲݕҏW��ʂւ̎Q��</param>
        ''' <param name="tabIdx">Ҳ���޺��۰ُ�̲��ޯ��</param>
        Protected Overrides Sub InitAllControl(ByRef mainEdit As frmEdit, ByVal tabIdx As Integer)
            Dim GrpArray() As cGrp_     ' ��ٰ���ޯ���̕\���ݒ�Ŏg�p����
            Dim LblArray() As cLbl_     ' ���قւ̕\���ݒ�Ŏg�p����
            Dim CtlArray() As Control   ' ���ٷ��ɂ��̫����ړ��Ŏg�p����

            m_MainEdit = mainEdit       ' ҲݕҏW��ʂւ̎Q�Ƃ�ݒ�
            m_TabIdx = tabIdx           ' ҲݕҏW�����޺��۰ُ�ł̲��ޯ��

            Try
                ' ��ĕ��@�̏����f�[�^
                Call InitCutMethodData()

                ' �J�b�g�`��̏����f�[�^
                Call InitCutTypeData()

                ' EDIT_DEF_User.ini������ޖ���ݒ�
                TAB_NAME = GetPrivateProfileString_S("CUT_LABEL", "TAB_NAM", m_sPath, "????")

                ' ���U���ʂ�̧���ڰ�ނłȂ��ꍇ��FL���H������ٰ���ޯ�����\���ɂ���
                If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then
                    CGrp_2.Visible = False
                Else ' ̧���ڰ�ނ̏ꍇ��ĸ�ٰ���ޯ����QڰĂ��\���ɂ���
                    CLbl_7.Visible = False
                    CTxt_1.Visible = False
                End If

                ' �ǉ���폜����H�������݂̐ݒ�
                With mainEdit
                    CBtn_Add.SetLblToolTip(.LblToolTip)
                    CBtn_Add.Text = GetPrivateProfileString_S("BUTTON_LABEL", "BTN_ADD", m_sPath, "ADD")
                    CBtn_Del.SetLblToolTip(.LblToolTip)
                    CBtn_Del.Text = GetPrivateProfileString_S("BUTTON_LABEL", "BTN_DEL", m_sPath, "DEL")
                    CBtn_FLC.SetLblToolTip(.LblToolTip)
                    CBtn_FLC.Text = GetPrivateProfileString_S("BUTTON_LABEL", "BTN_FLC", m_sPath, "Condition")
                End With

                ' ----------------------------------------------------------
                ' EDIT_DEF_User.ini�����ٰ���ޯ���ɕ\������ݒ�
                ' ----------------------------------------------------------
                ' 'V1.0.4.3�B CGrp_3 �ǉ� 'V2.0.0.0�F CGrp_4 �ǉ� 'V2.2.0.0�A  CGrp_5�ǉ�
                GrpArray = New cGrp_() {
                    CGrp_0, CGrp_1, CGrp_2, CGrp_3, CGrp_4, CGrp_5
                }
                For i As Integer = 0 To (GrpArray.Length - 1) Step 1
                    With GrpArray(i)
                        .TabIndex = ((254 - GrpArray.Length) + i) ' ���ٷ��ɂ��̫����ړ��ŕK�v
                        .Tag = i
                        .Text = GetPrivateProfileString_S( _
                                    "CUT_LABEL", (i.ToString("000") & "_GRP"), m_sPath, "????")
                    End With
                Next i
                ' �ǉ���폜���݂�����
                CPnl_Btn.TabIndex = 254 ' ���۰ٔz�u�\�ő吔(�Ō�ɐݒ�)

                ' ----------------------------------------------------------
                ' EDIT_DEF_User.ini�������قɕ\������ݒ�
                ' ----------------------------------------------------------
                LblArray = New cLbl_() {
                    CLbl_0, CLbl_1, CLbl_2, CLbl_3, CLbl_4, CLbl_5, CLbl_6, CLbl_7, CLbl_8,
                    CLbl_9, CLbl_10, CLbl_11, CLbl_12,
                    CLbl_13, CLbl_14, CLbl_15, CLbl_16, _
 _
                    CLbl_17, CLbl_18, CLbl_19, CLbl_20, CLbl_21, CLbl_22,
                    CLbl_23, CLbl_24, CLbl_25, CLbl_26, CLbl_27, _
 _
                    CLbl_28, CLbl_29, CLbl_30, CLbl_31,
                    CLbl_32, CLbl_33, CLbl_34, CLbl_35,
                    CLbl_80, CLbl_81, CLbl_82 'V2.2.1.7�@
                }
                For i As Integer = 0 To (LblArray.Length - 1) Step 1
                    LblArray(i).Text = GetPrivateProfileString_S( _
                            "CUT_LABEL", (i.ToString("000") & "_LBL"), m_sPath, "????")
                Next i

                ' ----------------------------------------------------------
                ' NG��ėp�׽�̐ݒ�
                ' ----------------------------------------------------------
                m_NGCut = New NGCut()
                With m_NGCut
                    .m_CtlArr = New Control() { _
                        CTxt_9, CCmb_6, CCmb_7 _
                    }
                End With

                ' ----------------------------------------------------------
                ' �������ݶ�ėp�׽�̐ݒ�
                ' ----------------------------------------------------------
                'm_Serpentine = New Serpentine()
                'With m_Serpentine
                '    .m_ctlArr = New Control() { _
                '        CLbl_6, CTxt_0, _
                '        CLbl_10, CTxt_5, CTxt_6, _
                '        CTxt_8, _
                '        CCmb_5 _
                '    }

                '    .m_lblArr = New Label() { _
                '        CLbl_11, CLbl_12 _
                '    }

                '    For i As Integer = 0 To (.m_strLbl.GetLength(0) - 1) Step 1
                '        For j As Integer = 0 To (.m_strLbl.GetLength(1) - 1) Step 1
                '            Dim no As Integer = Convert.ToInt32((i.ToString() & j.ToString()), 2)
                '            .m_strLbl(i, j) = GetPrivateProfileString_S( _
                '                    "CUT_SERPENTINE", (no.ToString("000") & "_LBL"), m_sPath, "????")
                '        Next j
                '    Next i
                'End With

                ' ----------------------------------------------------------
                ' ��ď����p�׽�̐ݒ�
                ' ----------------------------------------------------------
                m_CutCondition = New CutCondition()
                With m_CutCondition
                    .m_ctlArr = New Control(,) { _
                        {CLbl_32, CCmb_18, CTxt_31, CTxt_32, CTxt_33}, _
                        {CLbl_33, CCmb_19, CTxt_34, CTxt_35, CTxt_36}, _
                        {CLbl_34, CCmb_20, CTxt_37, CTxt_38, CTxt_39}, _
                        {CLbl_35, CCmb_21, CTxt_40, CTxt_41, CTxt_42} _
                    }
                End With

                ' ----------------------------------------------------------
                ' ��ĸ�ٰ���ޯ�����̺��۰ق�ݒ�(CtlArray�̏��Ԃƍ��킹��)
                ' ----------------------------------------------------------
                ' ###1042�@ CTxt_43�i�����j�ǉ� 'V2.1.0.0�@CCmb_22,CCmb_23,CTxt_46,CTxt_47,CTxt_48�ǉ�
                m_CtlCut = New Control() {
                    CCmb_0, CCmb_1, CCmb_2, CCmb_3, CTxt_0, CTxt_1, CTxt_2,
                    CTxt_3, CTxt_4, CTxt_5, CTxt_6, CTxt_7, CTxt_8, CCmb_4, CCmb_5,
                    CTxt_9, CTxt_10, CCmb_6, CCmb_7, CTxt_43,
                    CTxt_80, CTxt_81, CTxt_82, 'V2.2.1.7�@
                    CCmb_22, CCmb_23, CTxt_46, CTxt_47, CTxt_48,
                    CRT_Num
                }
                Call SetControlData(m_CtlCut) ' m_Serpentine,m_CutCondition�̐ݒ����ɂ����Ȃ�
                CCmb_4.DropDownStyle = ComboBoxStyle.DropDown   'V1.1.0.0�@ �p�x����͉�
                'CCmb_5.DropDownStyle = ComboBoxStyle.DropDown   'V1.1.0.0�@ �p�x����͉�

                ' ----------------------------------------------------------
                ' ���ޯ����ĸ�ٰ���ޯ�����̺��۰ق�ݒ�(CtlArray�̏��Ԃƍ��킹��)
                ' ----------------------------------------------------------
                m_CtlIdxCut = New Control(,) { _
                    {CTxt_11, CTxt_12, CTxt_13, CTxt_14, CCmb_8, CCmb_9}, _
                    {CTxt_15, CTxt_16, CTxt_17, CTxt_18, CCmb_10, CCmb_11}, _
                    {CTxt_19, CTxt_20, CTxt_21, CTxt_22, CCmb_12, CCmb_13}, _
                    {CTxt_23, CTxt_24, CTxt_25, CTxt_26, CCmb_14, CCmb_15}, _
                    {CTxt_27, CTxt_28, CTxt_29, CTxt_30, CCmb_16, CCmb_17} _
                }
                Call SetControlData(m_CtlIdxCut)

                'V1.0.4.3�B ADD ��
                ' ----------------------------------------------------------
                ' �k�J�b�g�p�����[�^��ٰ���ޯ�����̺��۰ق�ݒ�(CtlArray�̏��Ԃƍ��킹��)
                ' ----------------------------------------------------------
                m_CtlLCut = New Control(,) { _
                    {CL_Len1, CL_Q1, CL_Spd1, CCmb_Dir_1, CL_Tpt1}, _
                    {CL_Len2, CL_Q2, CL_Spd2, CCmb_Dir_2, CL_Tpt2}, _
                    {CL_Len3, CL_Q3, CL_Spd3, CCmb_Dir_3, CL_Tpt3}, _
                    {CL_Len4, CL_Q4, CL_Spd4, CCmb_Dir_4, CL_Tpt4}, _
                    {CL_Len5, CL_Q5, CL_Spd5, CCmb_Dir_5, CL_Tpt5}, _
                    {CL_Len6, CL_Q6, CL_Spd6, CCmb_Dir_6, CL_Tpt6}, _
                    {CL_Len7, CL_Q7, CL_Spd7, CCmb_Dir_7, CL_Tpt7} _
                }
                Call SetControlData(m_CtlLCut)
                'V1.0.4.3�B ADD ��
                'V2.0.0.0�F ADD ��
                ' ----------------------------------------------------------
                ' ���g���[�X�J�b�g�p�����[�^��ٰ���ޯ�����̺��۰ق�ݒ�(CtlArray�̏��Ԃƍ��킹��)
                ' ----------------------------------------------------------
                m_CtlRetraceCut = New Control(,) { _
                    {RTOFFX_1, RTOFFY_1, RTQRATE_1, RTV_1}, _
                    {RTOFFX_2, RTOFFY_2, RTQRATE_2, RTV_2}, _
                    {RTOFFX_3, RTOFFY_3, RTQRATE_3, RTV_3}, _
                    {RTOFFX_4, RTOFFY_4, RTQRATE_4, RTV_4}, _
                    {RTOFFX_5, RTOFFY_5, RTQRATE_5, RTV_5}, _
                    {RTOFFX_6, RTOFFY_6, RTQRATE_6, RTV_6}, _
                    {RTOFFX_7, RTOFFY_7, RTQRATE_7, RTV_7}, _
                    {RTOFFX_8, RTOFFY_8, RTQRATE_8, RTV_8}, _
                    {RTOFFX_9, RTOFFY_9, RTQRATE_9, RTV_9}, _
                    {RTOFFX_10, RTOFFY_10, RTQRATE_10, RTV_10} _
                }
                Call SetControlData(m_CtlRetraceCut)
                'V2.0.0.0�F ADD ��
                'V2.2.0.0�A��
                ' ----------------------------------------------------------
                ' U�J�b�g�p�����[�^�F��ٰ���ޯ�����̺��۰ق�ݒ�(CtlArray�̏��Ԃƍ��킹��)
                ' ----------------------------------------------------------
                m_CtlUCut = New Control() {
                    UclTxt_1, UclTxt_2, UcR1Txt_1, UcR1Txt_2,
                    UcqTxt_1, UcspdTxt_1, Ucdircmb, UcTurnCmb,
                    UcTurnTxt
                }
                Call SetControlData(m_CtlUCut)
                'V2.2.0.0�A��

                ' ----------------------------------------------------------
                ' FL���H������ٰ���ޯ�����̺��۰ق�ݒ�(CtlArray�̏��Ԃƍ��킹��)
                ' ----------------------------------------------------------
                m_CtlFLCnd = New Control(,) { _
                    {CCmb_18, CTxt_31, CTxt_32, CTxt_33}, _
                    {CCmb_19, CTxt_34, CTxt_35, CTxt_36}, _
                    {CCmb_20, CTxt_37, CTxt_38, CTxt_39}, _
                    {CCmb_21, CTxt_40, CTxt_41, CTxt_42} _
                }
                Call SetControlData(m_CtlFLCnd)

                ' ----------------------------------------------------------
                ' ���ׂĂ�÷���ޯ��������ޯ������݂ɑ΂��A̫����ړ��ݒ�������Ȃ�
                ' ��޷��A���ٷ��ɂ��̫����ړ����鏇�Ԃź��۰ق�CtlArray�ɐݒ肷��
                ' �g�p���Ȃ����۰ق� Enabled=False �܂��� Visible=False �ɂ���
                ' ----------------------------------------------------------
                ' ###1042�@ CTxt_43�i�����j�ǉ�'V2.1.0.0�@CCmb_22,CCmb_23,CTxt_46,CTxt_47,CTxt_48�ǉ�
                'V2.0.0.0�F CTxt_44, CTxt_45�폜�ACRT_Num�ǉ�
                CtlArray = New Control() {
                    CCmb_0, CCmb_1, CCmb_2, CCmb_3, CTxt_0, CTxt_1, CTxt_2,
                    CTxt_3, CTxt_4, CTxt_5, CTxt_6, CTxt_7, CTxt_8, CCmb_4, CCmb_5,
                    CTxt_9, CTxt_10, CCmb_6, CCmb_7, CTxt_43,
                    CTxt_80, CTxt_81, CTxt_82,'V2.2.1.7�@
                    CCmb_22, CCmb_23, CTxt_46, CTxt_47, CTxt_48,
                    CRT_Num, _
 _
                    CTxt_11, CTxt_12, CTxt_13, CTxt_14, CCmb_8, CCmb_9,
                    CTxt_15, CTxt_16, CTxt_17, CTxt_18, CCmb_10, CCmb_11,
                    CTxt_19, CTxt_20, CTxt_21, CTxt_22, CCmb_12, CCmb_13,
                    CTxt_23, CTxt_24, CTxt_25, CTxt_26, CCmb_14, CCmb_15,
                    CTxt_27, CTxt_28, CTxt_29, CTxt_30, CCmb_16, CCmb_17, _
 _
                    CL_Len1, CL_Q1, CL_Spd1, CCmb_Dir_1, CL_Tpt1,
                    CL_Len2, CL_Q2, CL_Spd2, CCmb_Dir_2, CL_Tpt2,
                    CL_Len3, CL_Q3, CL_Spd3, CCmb_Dir_3, CL_Tpt3,
                    CL_Len4, CL_Q4, CL_Spd4, CCmb_Dir_4, CL_Tpt4,
                    CL_Len5, CL_Q5, CL_Spd5, CCmb_Dir_5, CL_Tpt5,
                    CL_Len6, CL_Q6, CL_Spd6, CCmb_Dir_6, CL_Tpt6,
                    CL_Len7, CL_Q7, CL_Spd7, CCmb_Dir_7, _
 _
                    CRT_Num,
                    RTOFFX_1, RTOFFY_1, RTQRATE_1, RTV_1,
                    RTOFFX_2, RTOFFY_2, RTQRATE_2, RTV_2,
                    RTOFFX_3, RTOFFY_3, RTQRATE_3, RTV_3,
                    RTOFFX_4, RTOFFY_4, RTQRATE_4, RTV_4,
                    RTOFFX_5, RTOFFY_5, RTQRATE_5, RTV_5,
                    RTOFFX_6, RTOFFY_6, RTQRATE_6, RTV_6,
                    RTOFFX_7, RTOFFY_7, RTQRATE_7, RTV_7,
                    RTOFFX_8, RTOFFY_8, RTQRATE_8, RTV_8,
                    RTOFFX_9, RTOFFY_9, RTQRATE_9, RTV_9,
                    RTOFFX_10, RTOFFY_10, RTQRATE_10, RTV_10, _
 _
                    UclTxt_1, UclTxt_2, UcR1Txt_1, UcR1Txt_2,
                    UcqTxt_1, UcspdTxt_1, Ucdircmb, UcTurnCmb,
                    UcTurnTxt, _
 _
                    CCmb_18, CCmb_19, CCmb_20, CCmb_21, CBtn_FLC,
                    CBtn_Add, CBtn_Del
                }
                Call SetTabIndex(CtlArray) ' ��޲��ޯ����KeyDown����Ă�ݒ肷��

                ' ----------------------------------------------------------
                ' ��ʕ\������̫����������۰ق�ݒ肷��
                ' ----------------------------------------------------------
                FIRST_CONTROL = CtlArray(0) ' ��R�ԍ������ޯ�����I�������悤�ɂ���

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#Region "���������ɺ����ޯ���̐ݒ�������Ȃ�"
        ''' <summary>���������ɺ����ޯ����ؽĥү���ސݒ�������Ȃ�</summary>
        ''' <param name="cCombo">�ݒ�������Ȃ������ޯ��</param>
        Protected Overrides Sub InitComboBox(ByRef cCombo As cCmb_)
            Dim tag As Integer
            Dim Cnt As Integer  'V1.0.4.3�A
            Dim i As Integer

            Try
                With cCombo
                    tag = DirectCast(.Tag, Integer)
                    .Items.Clear()
                    Select Case (DirectCast(.Parent.Tag, Integer)) ' ��ٰ���ޯ�������
                        ' ------------------------------------------------------------------------------
                        Case 0 ' ��ĸ�ٰ���ޯ��
                            Select Case (tag)
                                Case 0 ' ��R�ԍ�
                                    '.Items.Add("") ' ڲ��Ĳ���ĂŐݒ肳���
                                Case 1 ' ��Ĕԍ�
                                    '.Items.Add("") ' ڲ��Ĳ���ĂŐݒ肳���
                                Case 2 ' ��ĕ��@(1:�ׯ�ݸ�, 2:���ޯ��)
                                    For i = 0 To m_lstCutMethod.Count - 1 Step 1
                                        .Items.Add(m_lstCutMethod(i).Name)
                                    Next i
                                Case 3 ' ��Č`��(1:��ڰ�, 2:L���, 3:�������ݶ��)
                                    For i = 0 To m_lstCutType.Count - 1 Step 1
                                        .Items.Add(m_lstCutType(i).Name)
                                    Next i
                                Case 4, 5 ' ��ĕ���1, ��ĕ���2(90���P�ʁ@0���`360��)
                                    'V1.0.4.3�A ADD START��
                                    For Cnt = 0 To MAX_DEGREES
                                        AngleArray(Cnt, 0) = Cnt
                                        AngleArray(Cnt, 1) = -1
                                    Next
                                    'V1.0.4.3�A ADD END��
                                    .Items.Add("     0��")
                                    AngleArray(0, 1) = 0       'V1.0.4.3�A
                                    .Items.Add("    90��")
                                    AngleArray(90, 1) = 1       'V1.0.4.3�A
                                    .Items.Add("   180��")
                                    AngleArray(180, 1) = 2       'V1.0.4.3�A
                                    .Items.Add("   270��")
                                    AngleArray(270, 1) = 3       'V1.0.4.3�A
                                    .Items.Add("    10��")
                                    AngleArray(10, 1) = 4       'V1.0.4.3�A
                                    .Items.Add("    20��")
                                    AngleArray(20, 1) = 5       'V1.0.4.3�A
                                    .Items.Add("    30��")
                                    AngleArray(30, 1) = 6       'V1.0.4.3�A
                                    .Items.Add("    40��")
                                    AngleArray(40, 1) = 7       'V1.0.4.3�A
                                    .Items.Add("    50��")
                                    AngleArray(50, 1) = 8       'V1.0.4.3�A
                                    .Items.Add("    60��")
                                    AngleArray(60, 1) = 9       'V1.0.4.3�A
                                    .Items.Add("    70��")
                                    AngleArray(70, 1) = 10       'V1.0.4.3�A
                                    .Items.Add("    80��")
                                    AngleArray(80, 1) = 11       'V1.0.4.3�A
                                    .Items.Add("   100��")
                                    AngleArray(100, 1) = 12       'V1.0.4.3�A
                                    .Items.Add("   110��")
                                    AngleArray(110, 1) = 13       'V1.0.4.3�A
                                    .Items.Add("   120��")
                                    AngleArray(120, 1) = 14       'V1.0.4.3�A
                                    .Items.Add("   130��")
                                    AngleArray(130, 1) = 15       'V1.0.4.3�A
                                    .Items.Add("   140��")
                                    AngleArray(140, 1) = 16       'V1.0.4.3�A
                                    .Items.Add("   150��")
                                    AngleArray(150, 1) = 17       'V1.0.4.3�A
                                    .Items.Add("   160��")
                                    AngleArray(160, 1) = 18       'V1.0.4.3�A
                                    .Items.Add("   170��")
                                    AngleArray(170, 1) = 19       'V1.0.4.3�A
                                Case 6 ' ����@��(0:���������, 1�ȏ�͊O�������ԍ�)
                                    '.Items.Add("") ' ڲ��Ĳ���ĂŐݒ肳���
                                Case 7 ' ����Ӱ��(0:����, 1:�����x)
                                    .Items.Add("����")
                                    .Items.Add("�����x")
                                    'V2.1.0.0�@��
                                Case 8
                                    .Items.Add("�Ȃ�")
                                    .Items.Add("����")
                                Case 9
                                    .Items.Add("�Ȃ�")
                                    .Items.Add("����")
                                    'V2.1.0.0�@��
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case 1 ' ���ޯ����ĸ�ٰ���ޯ��
                            Select Case (tag)
                                Case 0 ' ����@��(0:���������, 1�ȏ�͊O�������ԍ�)
                                    '.Items.Add("") ' ڲ��Ĳ���ĂŐݒ肳���
                                Case 1 ' ����Ӱ��(0:����, 1:�����x)
                                    '.Items.Add("") ' ڲ��Ĳ���ĂŐݒ肳���
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case 2 ' FL���H������ٰ���ޯ��
                            Select Case (tag)
                                Case 0 ' FL�ݒ�No.
                                    For i = 0 To 31 Step 1
                                        .Items.Add(String.Format("{0,5:#0}", i))
                                    Next i
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                            'V1.0.4.3�B��
                        Case 3 ' �k�J�b�g�p�����[�^�O���[�v�{�b�N�X
                            Select Case (tag)
                                Case 0 ' FL�ݒ�No.
                                    .DropDownStyle = ComboBoxStyle.DropDown
                                    For Cnt = 0 To MAX_DEGREES
                                        AngleArrayForLcut(Cnt, 0) = Cnt
                                        AngleArrayForLcut(Cnt, 1) = -1
                                    Next
                                    .Items.Add("     0��")
                                    AngleArrayForLcut(0, 1) = 0       'V1.0.4.3�A
                                    .Items.Add("    90��")
                                    AngleArrayForLcut(90, 1) = 1       'V1.0.4.3�A
                                    .Items.Add("   180��")
                                    AngleArrayForLcut(180, 1) = 2       'V1.0.4.3�A
                                    .Items.Add("   270��")
                                    AngleArrayForLcut(270, 1) = 3       'V1.0.4.3�A
                                    .Items.Add("    10��")
                                    AngleArrayForLcut(10, 1) = 4       'V1.0.4.3�A
                                    .Items.Add("    20��")
                                    AngleArrayForLcut(20, 1) = 5       'V1.0.4.3�A
                                    .Items.Add("    30��")
                                    AngleArrayForLcut(30, 1) = 6       'V1.0.4.3�A
                                    .Items.Add("    40��")
                                    AngleArrayForLcut(40, 1) = 7       'V1.0.4.3�A
                                    .Items.Add("    50��")
                                    AngleArrayForLcut(50, 1) = 8       'V1.0.4.3�A
                                    .Items.Add("    60��")
                                    AngleArrayForLcut(60, 1) = 9       'V1.0.4.3�A
                                    .Items.Add("    70��")
                                    AngleArrayForLcut(70, 1) = 10       'V1.0.4.3�A
                                    .Items.Add("    80��")
                                    AngleArrayForLcut(80, 1) = 11       'V1.0.4.3�A
                                    .Items.Add("   100��")
                                    AngleArrayForLcut(100, 1) = 12       'V1.0.4.3�A
                                    .Items.Add("   110��")
                                    AngleArrayForLcut(110, 1) = 13       'V1.0.4.3�A
                                    .Items.Add("   120��")
                                    AngleArrayForLcut(120, 1) = 14       'V1.0.4.3�A
                                    .Items.Add("   130��")
                                    AngleArrayForLcut(130, 1) = 15       'V1.0.4.3�A
                                    .Items.Add("   140��")
                                    AngleArrayForLcut(140, 1) = 16       'V1.0.4.3�A
                                    .Items.Add("   150��")
                                    AngleArrayForLcut(150, 1) = 17       'V1.0.4.3�A
                                    .Items.Add("   160��")
                                    AngleArrayForLcut(160, 1) = 18       'V1.0.4.3�A
                                    .Items.Add("   170��")
                                    AngleArrayForLcut(170, 1) = 19       'V1.0.4.3�A
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                            'V1.0.4.3�B��

                            'V2.2.0.0�A��
                        Case 5 ' �t�J�b�g�p�����[�^�O���[�v�{�b�N�X
                            Select Case (tag)
                                Case 0 ' �p�x
                                    For Cnt = 0 To MAX_DEGREES
                                        AngleArrayForUcut(Cnt, 0) = Cnt
                                        AngleArrayForUcut(Cnt, 1) = -1
                                    Next
                                    'V1.0.4.3�A ADD END��
                                    .Items.Add("     0��")
                                    AngleArrayForUcut(0, 1) = 0
                                    .Items.Add("    90��")
                                    AngleArrayForUcut(90, 1) = 1
                                    .Items.Add("   180��")
                                    AngleArrayForUcut(180, 1) = 2
                                    .Items.Add("   270��")
                                    AngleArrayForUcut(270, 1) = 3

                                Case 1 ' �^�[������ 
                                    .Items.Add("�b�v")
                                    .Items.Add("�b�b�v")
                            End Select
                            'V2.2.0.0�A��
                        Case Else
                            Throw New Exception("Parent.Tag - Case Else")
                    End Select

                            Call .SetStrTip("�h���b�v�_�E�����X�g����I�����Ă�������") ' °�����ү���ނ̐ݒ�
                    Call .SetLblToolTip(m_MainEdit.LblToolTip) ' ҲݕҏW��ʂ�°����ߎQ�Ɛݒ�
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex, cCombo.Name)
            End Try

        End Sub
#End Region

#Region "����������÷���ޯ���̐ݒ�������Ȃ�"
        ''' <summary>����������÷���ޯ���̏㉺���l�ү���ސݒ�������Ȃ�</summary>
        ''' <param name="cTextBox">�ݒ�������Ȃ�÷���ޯ��</param>
        Protected Overrides Sub InitTextBox(ByRef cTextBox As cTxt_)
            Dim strMin As String = ""           ' �ݒ肷��ϐ��̍ő�l
            Dim strMax As String = ""           ' �ݒ肷��ϐ��̍ŏ��l
            Dim strMsg As String = ""           ' �װ�ŕ\�����鍀�ږ�
            Dim no As String = ""
            Dim tag As Integer
            Dim strFlg As Boolean = False       ' �i�[����l�̎��(False=���l,True=������) ###1042�@
            Try
                tag = DirectCast(cTextBox.Tag, Integer)
                no = tag.ToString("000")

                Select Case (DirectCast(cTextBox.Parent.Tag, Integer)) ' ��ٰ���ޯ�������
                    ' ------------------------------------------------------------------------------
                    Case 0 ' ��ĸ�ٰ���ޯ��
                        ' �ް������װ���̕\����
                        strMsg = GetPrivateProfileString_S("CUT_CUT", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0 ' �������ݖ{��
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "10")
                            Case 1 ' Qڰ�
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "0.1")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "40.0")
                            Case 2 ' ���x
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "0.1")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "400.0")
                            Case 3 ' ��ĈʒuX
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "-80.0")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "80.0")
                            Case 4 ' ��ĈʒuY
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "-80.0")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "80.0")
                            Case 5 ' ��Ĉʒu2X
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "-80.0")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "80.0")
                            Case 6 ' ��Ĉʒu2Y
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "-80.0")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "80.0")
                            Case 7 ' ��Ē�1
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "0.001")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "20.000")
                            Case 8 ' ��Ē�2
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "0.001")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "20.000")
                            Case 9 ' ��ĵ�
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "-99.99")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "99.99")
                            Case 10 ' L����߲��
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "100.0")
                            Case 11 ' �����@###1042�@
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "10")
                                strFlg = True

                                'V2.2.1.7�@ ��
                            Case 12 ' �󎚌Œ蕔
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "14")
                                strFlg = True
                            Case 13 ' �J�n�ԍ�
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "6")
                                strFlg = True
                            Case 14 ' �d����
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "100")
                            Case 15 ' �㏸��
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "-99.99")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "99.99")
                            Case 16 ' �����l
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "-99.99")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "99.99")
                            Case 17 ' ����l
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "-99.99")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "99.99")

                                '    'V2.1.0.0�@��
                                'Case 12 ' �㏸��
                                '    strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "-99.99")
                                '    strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "99.99")
                                'Case 13 ' �����l
                                '    strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "-99.99")
                                '    strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "99.99")
                                'Case 14 ' ����l
                                '    strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "-99.99")
                                '    strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "99.99")
                                'V2.2.1.7�@ ��

                                'V2.1.0.0�@��
                                'V2.0.0.0�F                            Case 12 ' ���g���[�XQ���[�g V1.0.4.3�B
                                'V2.0.0.0�F                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "1")
                                'V2.0.0.0�F                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "10")
                                'V2.0.0.0�F                            Case 13 ' ���g���[�X���x V1.0.4.3�B
                                'V2.0.0.0�F                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "1")
                                'V2.0.0.0�F                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "10")
                            Case Else
                                Throw New Exception("Case " & tag & ": Nothing")
                        End Select
                        ' ------------------------------------------------------------------------------
                    Case 1 ' ���ޯ����ĸ�ٰ���ޯ��
                        ' �ް������װ���̕\����
                        strMsg = GetPrivateProfileString_S("CUT_INDEX", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0 ' ��ĉ�
                                strMin = GetPrivateProfileString_S("CUT_INDEX", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("CUT_INDEX", (no & "_MAX"), m_sPath, "999")
                            Case 1 ' ��Ē�(���ޯ���߯�)
                                strMin = GetPrivateProfileString_S("CUT_INDEX", (no & "_MIN"), m_sPath, "0.000")
                                strMax = GetPrivateProfileString_S("CUT_INDEX", (no & "_MAX"), m_sPath, "20.000")
                            Case 2 ' �߯����߰��
                                strMin = GetPrivateProfileString_S("CUT_INDEX", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("CUT_INDEX", (no & "_MAX"), m_sPath, "32767")
                            Case 3 ' �덷
                                strMin = GetPrivateProfileString_S("CUT_INDEX", (no & "_MIN"), m_sPath, "0.00")
                                strMax = GetPrivateProfileString_S("CUT_INDEX", (no & "_MAX"), m_sPath, "99.99")
                            Case Else
                                Throw New Exception("Case " & tag & ": Nothing")
                        End Select
                        ' ------------------------------------------------------------------------------
                    Case 2 ' FL���H������ٰ���ޯ��
                        ' �ް������װ���̕\����
                        strMsg = GetPrivateProfileString_S("CUT_FLCOND", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0 ' �d���l
                                strMin = GetPrivateProfileString_S("CUT_FLCOND", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("CUT_FLCOND", (no & "_MAX"), m_sPath, "1000")
                            Case 1 ' Qڰ�
                                strMin = GetPrivateProfileString_S("CUT_FLCOND", (no & "_MIN"), m_sPath, "0.1")
                                strMax = GetPrivateProfileString_S("CUT_FLCOND", (no & "_MAX"), m_sPath, "40.0")
                            Case 2 ' STEG�{��
                                strMin = GetPrivateProfileString_S("CUT_FLCOND", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("CUT_FLCOND", (no & "_MAX"), m_sPath, "15")
                            Case Else
                                Throw New Exception("Case " & tag & ": Nothing")
                        End Select
                        ' ------------------------------------------------------------------------------
                    Case 3
                        strMsg = GetPrivateProfileString_S("CUT_LCUT_PARA", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0 ' �J�b�g��
                                strMin = GetPrivateProfileString_S("CUT_LCUT_PARA", (no & "_MIN"), m_sPath, "0.001")
                                strMax = GetPrivateProfileString_S("CUT_LCUT_PARA", (no & "_MAX"), m_sPath, "20.000")
                            Case 1 ' Q���[�g
                                strMin = GetPrivateProfileString_S("CUT_LCUT_PARA", (no & "_MIN"), m_sPath, "0.1")
                                strMax = GetPrivateProfileString_S("CUT_LCUT_PARA", (no & "_MAX"), m_sPath, "40.0")
                            Case 2 ' ���x
                                strMin = GetPrivateProfileString_S("CUT_LCUT_PARA", (no & "_MIN"), m_sPath, "0.1")
                                strMax = GetPrivateProfileString_S("CUT_LCUT_PARA", (no & "_MAX"), m_sPath, "400.0")
                            Case 3 ' L����߲��
                                strMin = GetPrivateProfileString_S("CUT_LCUT_PARA", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("CUT_LCUT_PARA", (no & "_MAX"), m_sPath, "100.0")
                            Case Else
                                Throw New Exception("Case " & tag & ": Nothing")
                        End Select
                        'V2.0.0.0�F��
                    Case 4
                        strMsg = GetPrivateProfileString_S("CUT_RETRACE_PARA", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0 ' ���g���[�X�̃I�t�Z�b�g�w
                                strMin = GetPrivateProfileString_S("CUT_RETRACE_PARA", (no & "_MIN"), m_sPath, "-10.0")
                                strMax = GetPrivateProfileString_S("CUT_RETRACE_PARA", (no & "_MAX"), m_sPath, "10.0")
                            Case 1 ' ���g���[�X�̃I�t�Z�b�g�x
                                strMin = GetPrivateProfileString_S("CUT_RETRACE_PARA", (no & "_MIN"), m_sPath, "-10.0")
                                strMax = GetPrivateProfileString_S("CUT_RETRACE_PARA", (no & "_MAX"), m_sPath, "10.0")
                            Case 2 ' �X�g���[�g�J�b�g�E���g���[�X��Q���[�g
                                strMin = GetPrivateProfileString_S("CUT_RETRACE_PARA", (no & "_MIN"), m_sPath, "0.1")
                                strMax = GetPrivateProfileString_S("CUT_RETRACE_PARA", (no & "_MAX"), m_sPath, "40.0")
                            Case 3 ' �X�g���[�g�J�b�g�E���g���[�X�̃g�������x
                                strMin = GetPrivateProfileString_S("CUT_RETRACE_PARA", (no & "_MIN"), m_sPath, "0.1")
                                strMax = GetPrivateProfileString_S("CUT_RETRACE_PARA", (no & "_MAX"), m_sPath, "400.0")
                            Case 18 ' �X�g���[�g�J�b�g�E���g���[�X�{��'V2.0.0.0�F  'V2.1.0.0�@ Case 12���� Case 15�֕ύX 'V2.2.1.7�@ Case 15���� Case 18�֕ύX
                                strMsg = GetPrivateProfileString_S("CUT_CUT", (no & "_MSG"), m_sPath, "??????")
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "10")
                            Case Else
                                Throw New Exception("Case " & tag & ": Nothing")
                        End Select
                        'V2.0.0.0�F��

                        'V2.2.0.0�A ��
                    Case 5      ' U�J�b�g�p�����[�^ 
                        Select Case (tag)
                            Case 0 ' �k�P�J�b�g��
                                strMin = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MAX"), m_sPath, "20.0")
                            Case 1 ' �k�Q�J�b�g��
                                strMin = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MAX"), m_sPath, "20.0")
                            Case 2 ' �q�P
                                strMin = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MAX"), m_sPath, "20.0")
                            Case 3 ' �q�Q
                                strMin = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MAX"), m_sPath, "20.0")
                            Case 4 ' �p���[�g 
                                strMin = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MIN"), m_sPath, "0.1")
                                strMax = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MAX"), m_sPath, "40.0")
                            Case 5 ' ���x 
                                strMin = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MIN"), m_sPath, "0.1")
                                strMax = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MAX"), m_sPath, "400.0")
                            Case 6 ' L�^�[���|�C���g
                                strMin = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MAX"), m_sPath, "100.0")
                        End Select
                        'V2.2.0.0�A ��

                    Case Else
                                Throw New Exception("Parent.Tag - Case Else")
                        End Select

                        With cTextBox
                    Call .SetStrMsg(strMsg) ' �ް������װ���̕\�����ݒ�
                    Call .SetMinMax(strMin, strMax) ' �����l�����l�̐ݒ�
                    If (False = strFlg) Then                                                    '###1042�@
                        Call .SetStrTip(strMin & "�`" & strMax & "�͈̔͂Ŏw�肵�ĉ�����") ' °�����ү���ނ̐ݒ�
                    Else                                                                        '###1042�@
                        Call .SetStrTip(strMin & "�`" & strMax & "�����͈̔͂Ŏw�肵�ĉ�����")  '###1042�@
                        .MaxLength = Integer.Parse(strMax)                                      '###1042�@ SetControlData()���̏������f�Ŏg�p����
                        .TextAlign = HorizontalAlignment.Left                                   '###1042�@
                    End If                                                                      '###1042�@
                    Call .SetLblToolTip(m_MainEdit.LblToolTip) ' ҲݕҏW��ʂ�°����ߎQ�Ɛݒ�
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex, cTextBox.Name)
            End Try

        End Sub
#End Region

#Region "÷���ޯ��������ޯ���ɒl��\������"
        ''' <summary>÷���ޯ��������ޯ���ɒl��\������</summary>
        Protected Overrides Sub SetDataToText()
            Try
                Me.SuspendLayout()

                With m_MainEdit ' ��R�����Đ��Ƃ�0�ɂ͂Ȃ�Ȃ��d�l�̂��ߕs�v�Ǝv����
                    ' ��R���m�F
                    If (.W_PLT.RCount < 1) Then
                        m_ResNo = 1
                    End If
                    ' ��Đ��m�F
                    If (.W_REG(m_ResNo).intTNN < 1) Then
                        m_CutNo = 1
                    End If
                End With

                ' ------------------------
                If (SLP_VMES <> m_MainEdit.W_REG(m_ResNo).intSLP) AndAlso (SLP_RMES <> m_MainEdit.W_REG(m_ResNo).intSLP) Then
                    ' ��R�̽۰�߂� 5:�d������̂�, 6:��R����̂� �ł͂Ȃ��ꍇ
                    For Each ctl As Control In CGrp_0.Controls
                        ' ��ĸ�ٰ���ޯ�����̺��۰ق�L���ɂ���(�X�̐ݒ�͈ȍ~�ł����Ȃ�)
                        ctl.Enabled = True
                    Next
                End If
                ' ------------------------

                Call ChangedCutShape(m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCTYP) ' �֘A���۰ق̕\�����\����ݒ�

                ' ��ĸ�ٰ���ޯ���ݒ�
                Call SetCutData()

                ' ���ޯ����ĸ�ٰ���ޯ���ݒ�
                Call SetIdxCutData()

                ' �k�J�b�g�p�����[�^�O���[�v�{�b�N�X���̐ݒ�
                Call SetLCutParamData()

                '���g���[�X�J�b�g�p�����[�^�O���[�v�{�b�N�X���̐ݒ�
                SetRetraceCutParamData()    'V2.0.0.0�F

                'V2.2.0.0�A��
                ' �t�J�b�g�p�����[�^�̒ǉ�
                Call SetUCutParamData()
                'V2.2.0.0�A��

                ' ���U���ʂ�̧���ڰ�ނ̏ꍇ
                If (gSysPrm.stRAT.giOsc_Res = OSCILLATOR_FL) Then
                    ' FL���H������ٰ���ޯ���ݒ�
                    Call SetFLCndData()
                End If

                ' ------------------------
                If UserModule.IsMeasureOnly(m_MainEdit.W_REG, m_ResNo) Then
                    ' ��R�̽۰�߂� 5:�d������̂�, 6:��R����̂� �̏ꍇ
                    For Each ctl As Control In CGrp_0.Controls
                        ' ��ĸ�ٰ���ޯ�����̺��۰ق𖳌��ɂ���
                        If (Not ctl Is CCmb_0) Then
                            ' ��R�ԍ������ޯ�����������񖳌��ɂ����̫������߂�Ȃ��Ȃ�
                            ctl.Enabled = False
                        End If
                    Next
                    CLbl_0.Enabled = True       ' ��R������
                    CLblRN_0.Enabled = True     ' ��R��
                    CLbl_1.Enabled = True       ' ��R�ԍ�����
                    CLbl_2.Enabled = True       ' ��Đ�����
                    CLblCN_0.Enabled = True     ' ��Đ�
                    CLblCN_0.Text = 0           ' ��Đ��̕\����0�ɂ���
                End If
                ' ------------------------

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            Finally
                Me.ResumeLayout()
                Me.Refresh()
            End Try

        End Sub

#Region "��ĸ�ٰ���ޯ�����̐ݒ�"
        ''' <summary>��ĸ�ٰ���ޯ������÷���ޯ��������ޯ���ɒl��ݒ肷��</summary>
        Private Sub SetCutData()
            Dim idx As Integer

            Try
                With m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo)
                    For i As Integer = 0 To (m_CtlCut.Length - 1) Step 1
                        Select Case (i)
                            Case 0 ' ����R��, ��R�ԍ�
                                Dim rCnt As Integer = m_MainEdit.W_PLT.RCount
                                Dim cCombo As cCmb_ = DirectCast(m_CtlCut(i), cCmb_)
                                CLblRN_0.Text = rCnt.ToString() ' ����R��
                                With cCombo ' ��R�ԍ�
                                    .Items.Clear()
                                    For j As Integer = 1 To rCnt Step 1
                                        '.Items.Add(String.Format("{0,5:#0}", j)) ' ����R�����J��Ԃ�
                                        .Items.Add(j.ToString(0) & ":" & m_MainEdit.W_REG(j).strRNO) ' ����R�����J��Ԃ�
                                    Next j
                                End With
                                Call NoEventIndexChange(cCombo, (m_ResNo - 1)) ' �w���R�ԍ���ݒ�

                            Case 1 ' ����Đ�, ��Ĕԍ�
                                Dim cCnt As Integer = m_MainEdit.W_REG(m_ResNo).intTNN
                                Dim cCombo As cCmb_ = DirectCast(m_CtlCut(i), cCmb_)
                                CLblCN_0.Text = cCnt.ToString() ' ����Đ�
                                With cCombo ' ��Ĕԍ�
                                    .Items.Clear()
                                    For j As Integer = 1 To cCnt Step 1
                                        .Items.Add(String.Format("{0,5:#0}", j)) ' ����Đ����J��Ԃ�
                                    Next j
                                End With
                                Call NoEventIndexChange(cCombo, (m_CutNo - 1)) ' �w�趯Ĕԍ���ݒ�

                            Case 2 ' ��ĕ��@(1:�ׯ�ݸ�, 2:���ޯ��(ST��Ă̂�), 3:NG���)
                                If UserModule.IsMarking(m_MainEdit.W_REG, m_ResNo) Then             ' �}�[�L���O�̎�
#If cFORCEcCUT Then
                                    m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCUT = CNS_CUTM_FC   ' �J�b�g���@�������J�b�g�ɌŒ肷��B
#End If
                                    CCmb_2.Enabled = False
                                Else
                                    CCmb_2.Enabled = True
                                End If

                                idx = GetComboBoxValue2Index(.intCUT, Me.m_lstCutMethod)

                                Call NoEventIndexChange(DirectCast(m_CtlCut(i), cCmb_), idx)

                                Call ChangedCutMethod(.intCUT)

                            Case 3 ' ��Č`��(1:��ڰ�, 2:L���)
                                If m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCTYP = CNS_CUTP_M Then
                                    m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCUT = CNS_CUTM_TR
                                End If
                                Call NoEventIndexChange(DirectCast(m_CtlCut(i), cCmb_), GetComboBoxValue2Index(.intCTYP, Me.m_lstCutType))
                                Call ChangedCutShape(m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCTYP) ' �֘A���۰ق̕\�����\����ݒ�

                            Case 4 ' �������ݖ{��
                                m_CtlCut(i).Text = (.intNum).ToString()
                            Case 5 ' Qڰ�(0.1KHz��KHz)
                                m_CtlCut(i).Text = (.intQF1 / 10).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                            Case 6 ' ���x
                                m_CtlCut(i).Text = (.dblV1).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                            Case 7 ' ��ĈʒuX
                                m_CtlCut(i).Text = (.dblSTX).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                            Case 8 ' ��ĈʒuY
                                m_CtlCut(i).Text = (.dblSTY).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                            Case 9 ' ��Ĉʒu2X
                                m_CtlCut(i).Text = (.dblSX2).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                            Case 10 ' ��Ĉʒu2Y
                                m_CtlCut(i).Text = (.dblSY2).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                            Case 11 ' ��Ē�1
                                m_CtlCut(i).Text = (.dblDL2).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                            Case 12 ' ��Ē�2
                                m_CtlCut(i).Text = (.dblDL3).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                            Case 13 ' ��ĕ���1, ��ĕ���2
                                Dim iWK As Integer
                                Dim index As Integer
                                iWK = .intANG
                                Select Case (iWK)
                                    Case 0                  'V1.0.4.3�A ADD
                                        index = 0   ' 0��   'V1.0.4.3�A ADD
                                    Case 90
                                        index = 1   ' 90��
                                    Case 180
                                        index = 2   ' 180��
                                    Case 270
                                        index = 3   ' 270��
                                    Case 10
                                        index = 4   ' 10��
                                    Case 20
                                        index = 5   ' 20��
                                    Case 30
                                        index = 6   ' 30��
                                    Case 40
                                        index = 7   ' 40��
                                    Case 50
                                        index = 8   ' 50��
                                    Case 60
                                        index = 9   ' 60��
                                    Case 70
                                        index = 10  ' 70��
                                    Case 80
                                        index = 11  ' 80��
                                    Case 100
                                        index = 12  ' 100��
                                    Case 110
                                        index = 13  ' 110��
                                    Case 120
                                        index = 14  ' 120��
                                    Case 130
                                        index = 15  ' 130��
                                    Case 140
                                        index = 16  ' 140��
                                    Case 150
                                        index = 17  ' 150��
                                    Case 160
                                        index = 18  ' 160��
                                    Case 170
                                        index = 19  ' 170��
                                    Case Else
                                        'V1.0.4.3�A                                        index = 0   ' 0��
                                        index = Add_CCmb_4_Item(iWK)   'V1.0.4.3�A
                                End Select
                                Call NoEventIndexChange(DirectCast(m_CtlCut(i), cCmb_), index)

                            Case 14 ' ��ĕ���2
                                Dim iWK As Integer
                                iWK = .intANG2
                            Case 15 ' ��ĵ�
                                m_CtlCut(i).Text = (.dblCOF).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                            Case 16 ' L����߲��
                                m_CtlCut(i).Text = (.dblLTP).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                            Case 17 ' ����@��(0=��������, 1�ȏ�O������@��ԍ�)
                                ' GP-IB�o�^�@�햼��\������(�O���d��������)
                                Dim cCombo As cCmb_ = DirectCast(m_CtlCut(i), cCmb_)
                                Dim ctrg As String ' �ضް�����
                                Dim type As Integer = .intMType
                                Dim cnt As Integer = 0 ' ؽĂɒǉ��������ڐ�
                                idx = 0 ' �I��������ޯ��
                                cCombo.Items.Clear()
                                cCombo.Items.Add(" 0:���������")
                                With m_MainEdit
                                    If (0 < .W_PLT.GCount) Then ' GP-IB����@�킪�o�^����Ă���ꍇ
                                        For j As Integer = 1 To (.W_PLT.GCount) Step 1
                                            ctrg = .W_GPIB(j).strCTRG
                                            ' �ضް����ނ���̏ꍇ�A�O�������Ƃ���ؽĂɒǉ�
                                            If ("" <> ctrg) Then
                                                If (Not .W_GPIB(j).strGNAM Is Nothing) Then
                                                    cCombo.Items.Add(String.Format("{0,2:#0}", j) & ":" & .W_GPIB(j).strGNAM)
                                                Else
                                                    cCombo.Items.Add(String.Format("{0,2:#0}", j) & ":")
                                                End If
                                                ' �ǉ�����ؽĂ��ı���
                                                cnt = (cnt + 1)
                                                ' .intType(GP-IB�o�^�ԍ�)�Ɠ������ڂ�ؽĂɒǉ����ꂽ�ꍇ��
                                                ' ���̍��ڂ�I�����邽�߲��ޯ����ݒ肷��
                                                ' �g�p���̋@�킪�폜���ꂽ�ꍇ�AGP-IB��ޓ��̏�����
                                                ' .intMType��0�ƂȂ邽�ߓ�������킪�I�������
                                                If (type = j) Then idx = cnt
                                            End If
                                        Next j

                                    Else ' GP-IB����@��̓o�^���Ȃ��ꍇ
                                        .W_REG(m_ResNo).STCUT(m_CutNo).intMType = 0
                                        idx = 0 ' ���������
                                    End If
                                End With

                                ' ��ĕ��@��NG��Ă܂��͊O�������̏ꍇ����Ӱ�ނ𖳌��ɂ���
                                If (CNS_CUTM_NG = .intCUT) OrElse (0 < idx) Then
                                    m_CtlCut(CUT_TMM).Enabled = False ' ����Ӱ�ޖ���
                                Else
                                    m_CtlCut(CUT_TMM).Enabled = True ' ����Ӱ�ޗL��
                                End If
                                If (CNS_CUTM_NG = .intCUT Or CNS_CUTM_TR = .intCUT) Then       ' �g���b�L���O�܂���NG�J�b�g�̏ꍇ�́A��������̂�
                                    .intMType = 0
                                    m_CtlCut(CUT_MTYPE).Enabled = False
                                End If
                                Call NoEventIndexChange(cCombo, idx)

                            Case 18 ' ����Ӱ��(0:����, 1:�����x)
                                Call NoEventIndexChange(DirectCast(m_CtlCut(i), cCmb_), .intTMM)
                            Case 19 '###1042�@
                                m_CtlCut(i).Text = .cFormat
                                'V2.1.0.0�@��

                                'V2.2.1.7�@ ��
                            Case 20 ' �󎚌Œ蕔
                                m_CtlCut(i).Text = .cMarkFix
                            Case 21 ' �J�n�ԍ�
                                m_CtlCut(i).Text = .cMarkStartNum
                            Case 22 ' �d����
                                m_CtlCut(i).Text = .intMarkRepeatCnt

                            Case 23 ' ���s�[�g�L��
                                Call NoEventIndexChange(DirectCast(m_CtlCut(i), cCmb_), .iVariationRepeat)
                            Case 24 ' ����L��
                                Call NoEventIndexChange(DirectCast(m_CtlCut(i), cCmb_), .iVariation)
                            Case 25 ' �㏸��
                                m_CtlCut(i).Text = (.dRateOfUp).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                            Case 26 ' �����l
                                m_CtlCut(i).Text = (.dVariationLow).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                            Case 27 ' ����l
                                m_CtlCut(i).Text = (.dVariationHi).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                                'V2.1.0.0�@��
                            Case 28 ' ���g���[�X�{��'V2.0.0.0�F 'V2.1.0.0�@Case 20����Case 25�֕ύX
                                m_CtlCut(i).Text = (.intRetraceCnt).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())   'V2.0.0.0�F


                                'Case 20 ' ���s�[�g�L��
                                '    Call NoEventIndexChange(DirectCast(m_CtlCut(i), cCmb_), .iVariationRepeat)
                                'Case 21 ' ����L��
                                '    Call NoEventIndexChange(DirectCast(m_CtlCut(i), cCmb_), .iVariation)
                                'Case 22 ' �㏸��
                                '    m_CtlCut(i).Text = (.dRateOfUp).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                                'Case 23 ' �����l
                                '    m_CtlCut(i).Text = (.dVariationLow).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                                'Case 24 ' ����l
                                '    m_CtlCut(i).Text = (.dVariationHi).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                                '    'V2.1.0.0�@��
                                'Case 25 ' ���g���[�X�{��'V2.0.0.0�F 'V2.1.0.0�@Case 20����Case 25�֕ύX
                                '    m_CtlCut(i).Text = (.intRetraceCnt).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())   'V2.0.0.0�F

                                'V2.2.1.7�@��

                                'V1.0.4.3�B ADD ��
                                'V2.0.0.0�F                            Case 20 ' Qڰ�(0.1KHz��KHz)
                                'V2.0.0.0�F                                m_CtlCut(i).Text = (.intQF2 / 10).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                                'V2.0.0.0�F                            Case 21 ' ���x
                                'V2.0.0.0�F                                m_CtlCut(i).Text = (.dblV2).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                                'V1.0.4.3�B ADD ��
                            Case Else
                                Throw New Exception("Case " & i & ": Nothing")
                        End Select
                    Next i
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

#Region "�I�����ꂽ��ĕ��@�ɂ��֘A������۰ق̗L���������ύX����"
        ''' <summary>�I�����ꂽ��ĕ��@�ɂ��֘A������۰ق̗L���������ύX����</summary>
        ''' <param name="intCUT">1:�ׯ�ݸ�, 2:���ޯ�����, 3:NG���</param>
        Private Sub ChangedCutMethod(ByVal intCUT As Short)
            Try
                Select Case (intCUT)
                    Case CNS_CUTM_IX ' ���ޯ����Ă̏ꍇ
                        CGrp_1.Enabled = True ' ���ޯ����ĸ�ٰ���ޯ����\��
                        m_NGCut.Enabled = True ' NG��Ăł͎g�p���Ȃ����۰ق�L���ɂ���
                    Case CNS_CUTM_NG ' NG��Ă̏ꍇ
                        CGrp_1.Enabled = False ' ���ޯ����ĸ�ٰ���ޯ���𖳌��ɂ���
                        m_NGCut.Enabled = False ' NG��Ăł͎g�p���Ȃ����۰ق𖳌��ɂ���
                    Case Else ' �ׯ�ݸ�
                        CGrp_1.Enabled = False ' ���ޯ����ĸ�ٰ���ޯ���𖳌��ɂ���
                        m_NGCut.Enabled = True ' NG��Ăł͎g�p���Ȃ����۰ق�L���ɂ���
                End Select

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#Region "�I�����ꂽ��Č`��ɂ��֘A������۰ق̕\�����\����ύX����"
        'V2.0.0.0        ''' <summary>�I�����ꂽ��Č`��ɂ��֘A������۰ق̕\�����\����ύX����</summary>
        'V2.0.0.0        ''' <param name="selectedIdx">0:��ڰ�, 1:L���, 2:��������</param>
        'V2.0.0.0        Private Sub ChangedCutShape(ByVal selectedIdx As Integer)
        ''' <summary>
        ''' �I�����ꂽ��Č`��ɂ��֘A������۰ق̕\�����\����ύX����
        ''' </summary>
        ''' <param name="intCTYP">1:��ڰ�, 2:L���, 3:��������</param>
        ''' <remarks></remarks>
        Private Sub ChangedCutShape(ByVal intCTYP As Short)
            Try
                '###1042�@��
                Dim strMin As String = ""           ' �ݒ肷��ϐ��̍ő�l
                Dim strMax As String = ""           ' �ݒ肷��ϐ��̍ŏ��l
                Dim strMsg As String = ""           ' �װ�ŕ\�����鍀�ږ�

                strMsg = GetPrivateProfileString_S("CUT_CUT", ("007_MSG"), m_sPath, "??????")
                strMin = GetPrivateProfileString_S("CUT_CUT", ("007_MIN"), m_sPath, "0.001")
                strMax = GetPrivateProfileString_S("CUT_CUT", ("007_MAX"), m_sPath, "20.000")
                With DirectCast(m_CtlCut(CUT_LEN_1), cTxt_) ' �ڕW�l÷���ޯ��
                    Call .SetStrMsg(strMsg) ' �ް������װ���̕\�����ݒ�
                    Call .SetMinMax(strMin, strMax) ' �����l�����l�̐ݒ�
                    Call .SetStrTip(strMin & "�`" & strMax & "�͈̔͂Ŏw�肵�ĉ�����") ' °�����ү���ނ̐ݒ�
                End With
                '###1042�@��
                ' V1.1.0.0�B �r�s���g���[�X�A�k�J�b�g�ǉ��A�����}�[�L���O�ǉ��@�ԍ���DEFINE��
                ' �֘A���۰ق̕\�����\����ݒ�
                m_CtlCut(CUT_QRATE).Enabled = True     'V1.0.4.3�B�p���[�g
                m_CtlCut(CUT_SPEED).Enabled = True     'V1.0.4.3�B���x
                'V2.0.0.0                Select Case (selectedIdx)
                Select Case (intCTYP) ' ��Č`��
                    Case CNS_CUTP_ST ' ��ڰ�
                        'm_Serpentine.Visible = False
                        'V2.0.0.0                        m_CutCondition.Display = 1
                        CLbl_14.Visible = False ' L����߲��
                        m_CtlCut(CUT_LTP).Visible = False
                        'V2.0.0.0�R                        CLbl_36.Visible = False   ' ����        '###1042�@
                        m_CtlCut(CUT_LETTER).Enabled = False    '###1042�@
                        m_CtlCut(CUT_LEN_1).Enabled = True          'V1.0.4.3�B�J�b�g��
                        m_CtlCut(CUT_DIR_1).Enabled = True          'V1.0.4.3�B�J�b�g����
                        m_CtlCut(CUT_OFF).Enabled = True            'V1.0.4.3�B�J�b�g�I�t
                        m_CtlCut(CUT_START_2_X).Enabled = False     'V1.0.4.3�B�I�t�Z�b�gX
                        m_CtlCut(CUT_START_2_Y).Enabled = False     'V1.0.4.3�B�I�t�Z�b�gY
                        'V2.0.0.0�F                        m_CtlCut(CUT_TR_Q).Enabled = False          'V1.0.4.3�B���g���[�X�p���[�g
                        'V2.0.0.0�F                        m_CtlCut(CUT_TR_SPEED).Enabled = False      'V1.0.4.3�B���g���[�X���x
                        If (m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCUT = CNS_CUTM_IX) Then        'V2.2.0.0�I
                            CGrp_1.Enabled = True                       '�C���f�b�N�X�J�b�g'V2.1.0.0�@
                        End If
                        CGrp_3.Enabled = False                      'V1.0.4.3�B�k�J�b�g�p�����[�^
                        CGrp_4.Enabled = False                      ''V2.0.0.0�F���g���[�X�J�b�g�p�����[�^
                        CGrp_5.Enabled = False                      '�t�J�b�g�p�����[�^      'V2.2.0.0�A 
                        'V2.1.0.0�@��
                        m_CtlCut(CUT_VAR_REPEAT).Enabled = True     '���s�[�g�L��
                        m_CtlCut(CUT_VARIATION).Enabled = True      '����L��
                        m_CtlCut(CUT_RATE).Enabled = True           '�㏸��
                        m_CtlCut(CUT_VAR_LO).Enabled = True         '�����l
                        m_CtlCut(CUT_VAR_HI).Enabled = True         '����l
                        'V2.1.0.0�@��

                        'V2.2.1.7�@��
                        CLbl_80.Visible = False
                        CLbl_81.Visible = False
                        CLbl_82.Visible = False
                        m_CtlCut(CUT_MARK_FIX).Visible = False
                        m_CtlCut(CUT_ST_NUM).Visible = False
                        m_CtlCut(CUT_REPEAT_CNT).Visible = False
                        'V2.2.1.7�@ ��

                    Case CNS_CUTP_ST_TR ' �X�g���[�g�E���g���[�X(RETRACE)�J�b�g
                        'm_Serpentine.Visible = False
                        'V2.0.0.0                        m_CutCondition.Display = 1
                        CLbl_14.Visible = False ' L����߲��
                        m_CtlCut(CUT_LTP).Visible = False
                        'V2.0.0.0�R                        CLbl_36.Visible = False   ' ����        '###1042�@
                        m_CtlCut(CUT_LETTER).Enabled = False    '###1042�@
                        m_CtlCut(CUT_LEN_1).Enabled = True          'V1.0.4.3�B�J�b�g��
                        m_CtlCut(CUT_DIR_1).Enabled = True          'V1.0.4.3�B�J�b�g����
                        m_CtlCut(CUT_OFF).Enabled = True            'V1.0.4.3�B�J�b�g�I�t
                        'V2.0.0.0�F                        CLbl_10.Visible = True                      'V1.0.4.3�B�I�t�Z�b�g
                        m_CtlCut(CUT_START_2_X).Enabled = True      'V1.0.4.3�B�I�t�Z�b�gX
                        m_CtlCut(CUT_START_2_Y).Enabled = True      'V1.0.4.3�B�I�t�Z�b�gY
                        'V2.0.0.0�F                        m_CtlCut(CUT_TR_Q).Enabled = True           'V1.0.4.3�B���g���[�X�̂p���[�g
                        'V2.0.0.0�F                        m_CtlCut(CUT_TR_SPEED).Enabled = True       'V1.0.4.3�B���g���[�X�̑��x
                        If (m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCUT = CNS_CUTM_IX) Then        'V2.2.0.0�I
                            CGrp_1.Enabled = True                       '�C���f�b�N�X�J�b�g'V2.1.0.0�@
                        End If
                        CGrp_3.Enabled = False                      'V1.0.4.3�B�k�J�b�g�p�����[�^
                        CGrp_4.Enabled = True                      ''V2.0.0.0�F���g���[�X�J�b�g�p�����[�^
                        CGrp_5.Enabled = False                      '�t�J�b�g�p�����[�^      'V2.2.0.0�A 
                        'V2.1.0.0�@��
                        m_CtlCut(CUT_VAR_REPEAT).Enabled = True     '���s�[�g�L��
                        m_CtlCut(CUT_VARIATION).Enabled = True      '����L��
                        m_CtlCut(CUT_RATE).Enabled = True           '�㏸��
                        m_CtlCut(CUT_VAR_LO).Enabled = True         '�����l
                        m_CtlCut(CUT_VAR_HI).Enabled = True         '����l
                        'V2.1.0.0�@��

                        'V2.2.1.7�@ ��
                        CLbl_80.Visible = False
                        CLbl_81.Visible = False
                        CLbl_82.Visible = False
                        m_CtlCut(CUT_MARK_FIX).Visible = False
                        m_CtlCut(CUT_ST_NUM).Visible = False
                        m_CtlCut(CUT_REPEAT_CNT).Visible = False
                        'V2.2.1.7�@ ��

                    Case CNS_CUTP_L ' L���
                        'm_Serpentine.Visible = False
                        'V2.0.0.0                        m_CutCondition.Display = 2
                        'V1.0.4.3�B                        CLbl_14.Visible = True ' L����߲��
                        'V1.0.4.3�B                        m_CtlCut(CUT_LTP).Visible = True
                        CLbl_14.Visible = False ' L����߲��         'V1.0.4.3�B
                        m_CtlCut(CUT_QRATE).Enabled = False         'V1.0.4.3�B�p���[�g
                        m_CtlCut(CUT_SPEED).Enabled = False         'V1.0.4.3�B���x
                        m_CtlCut(CUT_LTP).Visible = False           'V1.0.4.3�B�^�[���|�C���g
                        m_CtlCut(CUT_LEN_1).Enabled = False         'V1.0.4.3�B�J�b�g��
                        m_CtlCut(CUT_DIR_1).Enabled = False         'V1.0.4.3�B�J�b�g����
                        'V2.0.0.0�R                        CLbl_36.Visible = False   ' ����        '###1042�@
                        m_CtlCut(CUT_LETTER).Enabled = False    '###1042�@
                        m_CtlCut(CUT_OFF).Enabled = True            'V1.0.4.3�B�J�b�g�I�t
                        m_CtlCut(CUT_START_2_X).Enabled = False     'V1.0.4.3�B�I�t�Z�b�gX
                        m_CtlCut(CUT_START_2_Y).Enabled = False     'V1.0.4.3�B�I�t�Z�b�gY
                        'V2.0.0.0�F                        m_CtlCut(CUT_TR_Q).Enabled = False          'V1.0.4.3�B���g���[�X�p���[�g
                        'V2.0.0.0�F                        m_CtlCut(CUT_TR_SPEED).Enabled = False      'V1.0.4.3�B���g���[�X���x
                        If (m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCUT = CNS_CUTM_IX) Then        'V2.2.0.0�I
                            CGrp_1.Enabled = True                       '�C���f�b�N�X�J�b�g'V2.1.0.0�@
                        End If
                        CGrp_3.Enabled = True                       'V1.0.4.3�B�k�J�b�g�p�����[�^
                        CGrp_4.Enabled = False                      ''V2.0.0.0�F���g���[�X�J�b�g�p�����[�^
                        CGrp_5.Enabled = False                      '�t�J�b�g�p�����[�^      'V2.2.0.0�A 
                        'V2.1.0.0�@��
                        m_CtlCut(CUT_VAR_REPEAT).Enabled = True     '���s�[�g�L��
                        m_CtlCut(CUT_VARIATION).Enabled = True      '����L��
                        m_CtlCut(CUT_RATE).Enabled = True           '�㏸��
                        m_CtlCut(CUT_VAR_LO).Enabled = True         '�����l
                        m_CtlCut(CUT_VAR_HI).Enabled = True         '����l
                        'V2.1.0.0�@��

                        'V2.2.1.7�@ ��
                        CLbl_80.Visible = False
                        CLbl_81.Visible = False
                        CLbl_82.Visible = False
                        m_CtlCut(CUT_MARK_FIX).Visible = False
                        m_CtlCut(CUT_ST_NUM).Visible = False
                        m_CtlCut(CUT_REPEAT_CNT).Visible = False
                        'V2.2.1.7�@ ��

                    Case CNS_CUTP_M                             '###1042�@
                        ''m_Serpentine.Visible = False
                        'V2.0.0.0                        m_CutCondition.Display = 2
                        CLbl_14.Visible = False ' L����߲��
                        m_CtlCut(CUT_LTP).Visible = False
                        CLbl_36.Visible = True                      ' ����         '###1042�@
                        m_CtlCut(CUT_LETTER).Enabled = True         '###1042�@������
                        m_CtlCut(CUT_LEN_1).Enabled = True          'V1.0.4.3�B�J�b�g��
                        m_CtlCut(CUT_DIR_1).Enabled = True          'V1.0.4.3�B�J�b�g����
                        m_CtlCut(CUT_DIR_1).Enabled = True          'V1.0.4.3�B�J�b�g����
                        m_CtlCut(CUT_OFF).Enabled = False           'V1.0.4.3�B�J�b�g�I�t
                        m_CtlCut(CUT_START_2_X).Enabled = False     'V1.0.4.3�B�I�t�Z�b�gX
                        m_CtlCut(CUT_START_2_Y).Enabled = False     'V1.0.4.3�B�I�t�Z�b�gY
                        'V2.0.0.0�F                        m_CtlCut(CUT_TR_Q).Enabled = False          'V1.0.4.3�B���g���[�X�p���[�g
                        'V2.0.0.0�F                        m_CtlCut(CUT_TR_SPEED).Enabled = False      'V1.0.4.3�B���g���[�X���x
                        CGrp_1.Enabled = False                       '�C���f�b�N�X�J�b�g'V2.1.0.0�@
                        CGrp_4.Enabled = False                      ''V2.0.0.0�F���g���[�X�J�b�g�p�����[�^
                        CGrp_3.Enabled = False                      'V1.0.4.3�B�k�J�b�g�p�����[�^
                        CGrp_5.Enabled = False                      '�t�J�b�g�p�����[�^      'V2.2.0.0�A 
                        'V2.1.0.0�@��
                        m_CtlCut(CUT_VAR_REPEAT).Enabled = False    '���s�[�g�L��
                        m_CtlCut(CUT_VARIATION).Enabled = False     '����L��
                        m_CtlCut(CUT_RATE).Enabled = False          '�㏸��
                        m_CtlCut(CUT_VAR_LO).Enabled = False        '�����l
                        m_CtlCut(CUT_VAR_HI).Enabled = False        '����l
                        'V2.1.0.0�@��

                        '###1042�@��
                        With DirectCast(m_CtlCut(CUT_LEN_1), cTxt_) ' �ڕW�l÷���ޯ��
                            strMin = "0.1"
                            strMax = "10.0"
                            strMsg = "��������"
                            Call .SetStrMsg(strMsg) ' �ް������װ���̕\�����ݒ�
                            Call .SetMinMax(strMin, strMax) ' �����l�����l�̐ݒ�
                            Call .SetStrTip(strMin & "�`" & strMax & "�͈̔͂Ŏw�肵�ĉ�����") ' °�����ү���ނ̐ݒ�
                        End With
                        '###1042�@��

                        'V2.2.1.7�@ ��
                        If ((m_MainEdit.W_stUserData.iTrimType = 5) And (m_MainEdit.W_REG(m_ResNo).intSLP = SLP_MARK)) Then
                            CLbl_80.Visible = True
                            CLbl_81.Visible = True
                            CLbl_82.Visible = True
                            m_CtlCut(CUT_MARK_FIX).Visible = True
                            m_CtlCut(CUT_ST_NUM).Visible = True
                            m_CtlCut(CUT_REPEAT_CNT).Visible = True

                            CLbl_36.Visible = False
                            m_CtlCut(CUT_LETTER).Visible = False                      ' ����
                        Else
                            CLbl_80.Visible = False
                            CLbl_81.Visible = False
                            CLbl_82.Visible = False

                            m_CtlCut(CUT_MARK_FIX).Visible = False
                            m_CtlCut(CUT_ST_NUM).Visible = False
                            m_CtlCut(CUT_REPEAT_CNT).Visible = False

                            CLbl_36.Visible = True
                            m_CtlCut(CUT_LETTER).Visible = True                      ' ����
                        End If
                        'V2.2.1.7�@ ��

                        'V2.2.0.0�A ��
                    Case CNS_CUTP_U  ' U���
                        CLbl_14.Visible = False ' L����߲��         '
                        m_CtlCut(CUT_QRATE).Enabled = False         '�p���[�g
                        m_CtlCut(CUT_SPEED).Enabled = False         '���x
                        m_CtlCut(CUT_LTP).Visible = False           '�^�[���|�C���g
                        m_CtlCut(CUT_LEN_1).Enabled = False         '�J�b�g��
                        m_CtlCut(CUT_DIR_1).Enabled = False         '�J�b�g����
                        m_CtlCut(CUT_LETTER).Enabled = False
                        m_CtlCut(CUT_OFF).Enabled = True            '�J�b�g�I�t
                        m_CtlCut(CUT_START_2_X).Enabled = False     '�I�t�Z�b�gX
                        m_CtlCut(CUT_START_2_Y).Enabled = False     '�I�t�Z�b�gY
                        If (m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCUT = CNS_CUTM_IX) Then        'V2.2.0.0�I
                            CGrp_1.Enabled = True                       '�C���f�b�N�X�J�b�g
                        End If
                        CGrp_3.Enabled = False                      '�k�J�b�g�p�����[�^
                        CGrp_4.Enabled = False                      '���g���[�X�J�b�g�p�����[�^
                        CGrp_5.Enabled = True                       '�t�J�b�g�p�����[�^      'V2.2.0.0�A 
                        m_CtlCut(CUT_VAR_REPEAT).Enabled = True     '���s�[�g�L��
                        m_CtlCut(CUT_VARIATION).Enabled = True      '����L��
                        m_CtlCut(CUT_RATE).Enabled = True           '�㏸��
                        m_CtlCut(CUT_VAR_LO).Enabled = True         '�����l
                        m_CtlCut(CUT_VAR_HI).Enabled = True         '����l
                        'V2.2.0.0�A ��

                        'V2.2.1.7�@ ��
                        CLbl_80.Visible = False
                        CLbl_81.Visible = False
                        CLbl_82.Visible = False
                        m_CtlCut(CUT_MARK_FIX).Visible = False
                        m_CtlCut(CUT_ST_NUM).Visible = False
                        m_CtlCut(CUT_REPEAT_CNT).Visible = False
                        'V2.2.1.7�@ ��

                        'Case 2 ' ��������
                        '    m_Serpentine.Visible = True
                        '    m_CutCondition.Display = 2
                        '    CLbl_14.Visible = False ' L����߲��
                        '    m_CtlCut(CUT_LTP).Visible = False
                        '    CLbl_36.Visible = False   ' ����        '###1042�@
                        '    m_CtlCut(CUT_LETTER).Visible = True     '###1042�@
                    Case Else
                        'm_Serpentine.Visible = False
                        'V2.0.0.0                        m_CutCondition.Display = 1
                        CLbl_14.Visible = False ' L����߲��
                        m_CtlCut(CUT_LETTER).Enabled = False    '###1042�@
                        m_CtlCut(CUT_LEN_1).Enabled = True          'V1.0.4.3�B�J�b�g��
                        m_CtlCut(CUT_DIR_1).Enabled = True          'V1.0.4.3�B�J�b�g����
                        m_CtlCut(CUT_LTP).Visible = False
                        m_CtlCut(CUT_OFF).Enabled = True            'V1.0.4.3�B�J�b�g�I�t
                        m_CtlCut(CUT_START_2_X).Enabled = False     'V1.0.4.3�B�I�t�Z�b�gX
                        m_CtlCut(CUT_START_2_Y).Enabled = False     'V1.0.4.3�B�I�t�Z�b�gY
                        'V2.0.0.0�F                        m_CtlCut(CUT_TR_Q).Enabled = False          'V1.0.4.3�B���g���[�X�p���[�g
                        'V2.0.0.0�F                        m_CtlCut(CUT_TR_SPEED).Enabled = False      'V1.0.4.3�B���g���[�X���x
                        If (m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCUT = CNS_CUTM_IX) Then        'V2.2.0.0�I
                            CGrp_1.Enabled = True                       '�C���f�b�N�X�J�b�g'V2.1.0.0�@
                        End If
                        CGrp_3.Enabled = False                      'V1.0.4.3�B�k�J�b�g�p�����[�^
                        CGrp_4.Enabled = False                      ''V2.0.0.0�F���g���[�X�J�b�g�p�����[�^
                        CGrp_5.Enabled = False                      '�t�J�b�g�p�����[�^      'V2.2.0.0�A 
                        '    �J�b�g���O�̎��̈�                    Throw New Exception("Case " & selectedIdx & ": Nothing")

                        'V2.2.1.7�@ ��
                        CLbl_80.Visible = False
                        CLbl_81.Visible = False
                        CLbl_82.Visible = False
                        m_CtlCut(CUT_MARK_FIX).Visible = False
                        m_CtlCut(CUT_ST_NUM).Visible = False
                        m_CtlCut(CUT_REPEAT_CNT).Visible = False
                        'V2.2.1.7�@ ��

                End Select

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#End Region

#Region "���ޯ����ĸ�ٰ���ޯ�����̐ݒ�"
        ''' <summary>���ޯ����ĸ�ٰ���ޯ������÷���ޯ��������ޯ���ɒl��ݒ肷��</summary>
        Private Sub SetIdxCutData()
            Try
                With m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo)
                    For i As Integer = 0 To (m_CtlIdxCut.GetLength(0) - 1) Step 1
                        For j As Integer = 0 To (m_CtlIdxCut.GetLength(1) - 1) Step 1
                            Select Case (j)
                                Case 0 ' ��ĉ�1-5
                                    m_CtlIdxCut(i, j).Text = (.intIXN(i + 1)).ToString()
                                Case 1 ' ��Ē�1-5
                                    m_CtlIdxCut(i, j).Text = (.dblDL1(i + 1)).ToString(DirectCast(m_CtlIdxCut(i, j), cTxt_).GetStrFormat())
                                Case 2 ' �߯����߰��1-5(ms)
                                    m_CtlIdxCut(i, j).Text = (.lngPAU(i + 1)).ToString()
                                Case 3 ' �덷1-5(%)
                                    m_CtlIdxCut(i, j).Text = (.dblDEV(i + 1)).ToString(DirectCast(m_CtlIdxCut(i, j), cTxt_).GetStrFormat())
                                Case 4 ' ����@��(0:���������, 1�ȏ�͊O�������ԍ�)
                                    Dim cCombo As cCmb_ = DirectCast(m_CtlIdxCut(i, j), cCmb_)
                                    Dim ctrg As String ' �ضް�����
                                    Dim type As Integer = Convert.ToInt32(.intIXMType(i + 1))
                                    Dim cnt As Integer = 0 ' ؽĂɒǉ��������ڐ�
                                    Dim idx As Integer = 0 ' �I��������ޯ��
                                    Dim ctlIdx As Integer = GetCtlIdx(cCombo, DirectCast(cCombo.Tag, Integer)) ' 1�����ڂ̲��ޯ��
                                    cCombo.Items.Clear()
                                    cCombo.Items.Add(" 0:���������")
                                    With m_MainEdit
                                        If (0 < .W_PLT.GCount) Then ' GP-IB����@�킪�o�^����Ă���ꍇ
                                            For k As Integer = 1 To (.W_PLT.GCount) Step 1
                                                ctrg = .W_GPIB(k).strCTRG
                                                ' �ضް����ނ���̏ꍇ�A�O�������Ƃ���ؽĂɒǉ�
                                                If ("" <> ctrg) Then
                                                    If (Not .W_GPIB(k).strGNAM Is Nothing) Then
                                                        cCombo.Items.Add(String.Format("{0,2:#0}", k) & ":" & .W_GPIB(k).strGNAM)
                                                    Else
                                                        cCombo.Items.Add(String.Format("{0,2:#0}", k) & ":")
                                                    End If
                                                    ' �ǉ�����ؽĂ��ı���
                                                    cnt = (cnt + 1)
                                                    ' .intType(GP-IB�o�^�ԍ�)�Ɠ������ڂ�ؽĂɒǉ����ꂽ�ꍇ��
                                                    ' ���̍��ڂ�I�����邽�߲��ޯ����ݒ肷��
                                                    ' �g�p���̋@�킪�폜���ꂽ�ꍇ�AGP-IB��ޓ��̏�����
                                                    ' .intIXMType��0�ƂȂ邽�ߓ�������킪�I�������
                                                    If (type = k) Then idx = cnt
                                                End If
                                            Next k

                                            If (0 < idx) Then
                                                m_CtlIdxCut(ctlIdx, IDX_TMM).Enabled = False ' ����Ӱ�ޖ���
                                            Else
                                                m_CtlIdxCut(ctlIdx, IDX_TMM).Enabled = True ' ����Ӱ�ޗL��
                                            End If

                                        Else ' GP-IB����@��̓o�^���Ȃ��ꍇ
                                            .W_REG(m_ResNo).STCUT(m_CutNo).intIXMType(i + 1) = 0
                                            idx = 0 ' ���������
                                            m_CtlIdxCut(ctlIdx, IDX_TMM).Enabled = True ' ����Ӱ�ޗL��
                                        End If
                                    End With
                                    Call NoEventIndexChange(cCombo, idx)

                                Case 5 ' ����Ӱ��(0:����, 1:�����x)
                                    Dim cCombo As cCmb_ = DirectCast(m_CtlIdxCut(i, j), cCmb_)
                                    With cCombo
                                        .Items.Clear()
                                        .Items.Add("����")
                                        .Items.Add("�����x")
                                    End With
                                    Call NoEventIndexChange(cCombo, Convert.ToInt32(.intIXTMM(i + 1)))

                                Case Else
                                    Throw New Exception("i = " & i & ", Case Else")
                            End Select
                        Next j
                    Next i
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#Region "FL���H������ٰ���ޯ�����̒l��ݒ�"
        ''' <summary>FL���H������ٰ���ޯ������÷���ޯ��������ޯ���ɒl��ݒ肷��</summary>
        Private Sub SetFLCndData()
            Try
#If cOSCILLATORcFLcUSE Then
                With m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo)
                    For i As Integer = 0 To (m_CtlFLCnd.GetLength(0) - 1) Step 1
                        For j As Integer = 0 To (m_CtlFLCnd.GetLength(1) - 1) Step 1
                            Select Case (j)
                                Case 0 ' FL�ݒ�No.
                                    m_CtlFLCnd(i, j).Text = (.intCND(j + 1)).ToString("##0")
                                    Call NoEventIndexChange(DirectCast(m_CtlFLCnd(i, j), cCmb_), _
                                                                        Convert.ToInt32(.intCND(j + 1)))
                                Case 1 ' �d���l
                                    m_CtlFLCnd(i, j).Text = (stCND.Curr(.intCND(j + 1))).ToString("###0")
                                Case 2 ' Qڰ�
                                    m_CtlFLCnd(i, j).Text = (stCND.Freq(.intCND(j + 1))).ToString("#0.0")
                                Case 3 ' STEG�{��
                                    m_CtlFLCnd(i, j).Text = (stCND.Steg(.intCND(j + 1))).ToString("#0")
                                Case Else
                                    Throw New Exception("i = " & i & ", Case " & j & ": Nothing")
                            End Select
                        Next j
                    Next i
                End With
#End If

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#Region "�k�J�b�g�p�����[�^�O���[�v�{�b�N�X���̐ݒ�"
        ''' <summary>�k�J�b�g�p�����[�^�O���[�v�{�b�N�X���̃e�L�X�g�{�b�N�X�E�R���{�{�b�N�X�ɒl��ݒ肷��</summary>
        Private Sub SetLCutParamData()
            Try
                With m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo)
                    For i As Integer = 0 To (m_CtlLCut.GetLength(0) - 1) Step 1
                        For j As Integer = 0 To (m_CtlLCut.GetLength(1) - 1) Step 1
                            Select Case (j)
                                Case 0 ' �J�b�g��
                                    m_CtlLCut(i, j).Text = (.dCutLen(i + 1)).ToString(DirectCast(m_CtlLCut(i, j), cTxt_).GetStrFormat())
                                Case 1 ' �p���[�g
                                    m_CtlLCut(i, j).Text = (.dQRate(i + 1) / 10.0).ToString(DirectCast(m_CtlLCut(i, j), cTxt_).GetStrFormat())
                                Case 2 ' ���x
                                    m_CtlLCut(i, j).Text = (.dSpeed(i + 1)).ToString(DirectCast(m_CtlLCut(i, j), cTxt_).GetStrFormat())
                                Case 3 ' �p�x
                                    Dim index As Integer
                                    Select Case (.dAngle(i + 1))
                                        Case 0
                                            index = 0   ' 0��
                                        Case 90
                                            index = 1   ' 90��
                                        Case 180
                                            index = 2   ' 180��
                                        Case 270
                                            index = 3   ' 270��
                                        Case 10
                                            index = 4   ' 10��
                                        Case 20
                                            index = 5   ' 20��
                                        Case 30
                                            index = 6   ' 30��
                                        Case 40
                                            index = 7   ' 40��
                                        Case 50
                                            index = 8   ' 50��
                                        Case 60
                                            index = 9   ' 60��
                                        Case 70
                                            index = 10  ' 70��
                                        Case 80
                                            index = 11  ' 80��
                                        Case 100
                                            index = 12  ' 100��
                                        Case 110
                                            index = 13  ' 110��
                                        Case 120
                                            index = 14  ' 120��
                                        Case 130
                                            index = 15  ' 130��
                                        Case 140
                                            index = 16  ' 140��
                                        Case 150
                                            index = 17  ' 150��
                                        Case 160
                                            index = 18  ' 160��
                                        Case 170
                                            index = 19  ' 170��
                                        Case Else
                                            index = Add_CCmb_Dir_X_Item(.dAngle(i + 1), DirectCast(m_CtlLCut(i, j), cCmb_))
                                    End Select
                                    Call NoEventIndexChange(DirectCast(m_CtlLCut(i, j), cCmb_), index)
                                Case 4 ' �^�[���|�C���g
                                    m_CtlLCut(i, j).Text = (.dTurnPoint(i + 1)).ToString(DirectCast(m_CtlLCut(i, j), cTxt_).GetStrFormat())
                                Case Else
                                    Throw New Exception("i = " & i & ", Case Else")
                            End Select
                        Next j
                    Next i
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

        'V2.0.0.0�F ADD ��
#Region "���g���[�X�J�b�g�p�����[�^�O���[�v�{�b�N�X���̐ݒ�"
        ''' <summary>���g���[�X�J�b�g�p�����[�^�O���[�v�{�b�N�X���̃e�L�X�g�{�b�N�X�E�R���{�{�b�N�X�ɒl��ݒ肷��</summary>
        Private Sub SetRetraceCutParamData()
            Try
                With m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo)
                    For i As Integer = 0 To (m_CtlRetraceCut.GetLength(0) - 1) Step 1
                        For j As Integer = 0 To (m_CtlRetraceCut.GetLength(1) - 1) Step 1
                            Select Case (j)
                                Case 0 ' ���g���[�X�̃I�t�Z�b�g�w
                                    m_CtlRetraceCut(i, j).Text = (.dblRetraceOffX(i + 1)).ToString(DirectCast(m_CtlRetraceCut(i, j), cTxt_).GetStrFormat())
                                Case 1 ' ���g���[�X�̃I�t�Z�b�g�x
                                    m_CtlRetraceCut(i, j).Text = (.dblRetraceOffY(i + 1)).ToString(DirectCast(m_CtlRetraceCut(i, j), cTxt_).GetStrFormat())
                                Case 2 ' �X�g���[�g�J�b�g�E���g���[�X��Q���[�g(0.1KHz)�Ɏg�p
                                    m_CtlRetraceCut(i, j).Text = (.dblRetraceQrate(i + 1) / 10.0).ToString(DirectCast(m_CtlRetraceCut(i, j), cTxt_).GetStrFormat())
                                Case 3 ' ���x
                                    m_CtlRetraceCut(i, j).Text = (.dblRetraceSpeed(i + 1)).ToString(DirectCast(m_CtlRetraceCut(i, j), cTxt_).GetStrFormat())
                                Case Else
                                    Throw New Exception("i = " & i & ", Case Else")
                            End Select
                        Next j
                    Next i
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region
        'V2.0.0.0�F ADD ��
#End Region

#Region "���ׂĂ�÷���ޯ�����ް������������Ȃ�"
#If cCUTDATAcCHECKcBYDATA Then
        ''' <summary>���ׂĂ�÷���ޯ�����ް������������Ȃ�</summary>
        ''' <returns>0=����, 1=�װ</returns>
        Friend Overrides Function CheckAllTextData() As Integer
            Dim ret As Integer
            Try
                ret = CutDataCheckByDataOnly()
            Catch ex As Exception
                Call MsgBox_Exception(ex)
                ret = 1
            Finally
                m_CheckFlg = False ' �����I��
                CheckAllTextData = ret
            End Try

        End Function
#Region "�f�[�^�݂̂̃`�F�b�N����"
        ''' <summary>
        ''' �J�b�g�f�[�^�f�[�^�݂̂̃`�F�b�N
        ''' </summary>
        ''' <returns>0=����, 1=�G���[</returns>
        ''' <remarks></remarks>
        Private Function CutDataCheckByDataOnly() As Integer
            Dim strMSG As String = "", strMin As String = "", strMax As String = ""

            CutDataCheckByDataOnly = 0

            ' �������ݖ{��
            Dim intNumMin As Short = Short.Parse(GetPrivateProfileString_S("CUT_CUT", "000_MIN", m_sPath, "1"))
            Dim intNumMax As Short = Short.Parse(GetPrivateProfileString_S("CUT_CUT", "000_MAX", m_sPath, "10"))
            ' Qڰ�
            Dim intQF1Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "001_MIN", m_sPath, "0.1"))
            Dim intQF1Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "001_MAX", m_sPath, "40.0"))
            ' ���x
            Dim dblV1Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "002_MIN", m_sPath, "0.1"))
            Dim dblV1Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "002_MAX", m_sPath, "400.0"))
            ' ��ĈʒuX
            Dim dblSTXMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "003_MIN", m_sPath, "-80.0"))
            Dim dblSTXMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "003_MAX", m_sPath, "80.0"))
            ' ��ĈʒuY
            Dim dblSTYMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "004_MIN", m_sPath, "-80.0"))
            Dim dblSTYMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "004_MAX", m_sPath, "80.0"))
            ' ��Ĉʒu2X
            Dim dblSX2Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "005_MIN", m_sPath, "-80.0"))
            Dim dblSX2Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "005_MAX", m_sPath, "80.0"))
            ' ��Ĉʒu2Y
            Dim dblSY2Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "006_MIN", m_sPath, "-80.0"))
            Dim dblSY2Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "006_MAX", m_sPath, "80.0"))
            ' ��Ē�1
            Dim dblDL2Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "007_MIN", m_sPath, "0.001"))
            Dim dblDL2Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "007_MAX", m_sPath, "20.000"))
            ' ��Ē�2
            Dim dblDL3Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "008_MIN", m_sPath, "0.001"))
            Dim dblDL3Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "008_MAX", m_sPath, "20.000"))
            ' ��ĵ�
            Dim dblCOFMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "009_MIN", m_sPath, "-99.99"))
            Dim dblCOFMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "009_MAX", m_sPath, "99.99"))
            ' L����߲��
            Dim dblLTPMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "010_MIN", m_sPath, "0.0"))
            Dim dblLTPMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "010_MAX", m_sPath, "100.0"))

            ' ��ĉ�
            Dim intIXNMin As Short = Short.Parse(GetPrivateProfileString_S("CUT_INDEX", "000_MIN", m_sPath, "0"))
            Dim intIXNMax As Short = Short.Parse(GetPrivateProfileString_S("CUT_INDEX", "000_MAX", m_sPath, "999"))
            ' ��Ē�(���ޯ���߯�)
            Dim dblDL1Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_INDEX", "001_MIN", m_sPath, "0.000"))
            Dim dblDL1Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_INDEX", "001_MAX", m_sPath, "20.000"))
            ' �߯����߰��
            Dim lngPAUMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_INDEX", "002_MIN", m_sPath, "0"))
            Dim lngPAUMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_INDEX", "002_MAX", m_sPath, "32767"))
            ' �덷
            Dim dblDEVMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_INDEX", "003_MIN", m_sPath, "0.00"))
            Dim dblDEVMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_INDEX", "003_MAX", m_sPath, "99.99"))


            '' �d���l
            'Dim CurrMin As Integer = Integer.Parse(GetPrivateProfileString_S("CUT_FLCOND", "000_MIN", m_sPath, "1"))
            'Dim CurrMax As Integer = Integer.Parse(GetPrivateProfileString_S("CUT_FLCOND", "000_MAX", m_sPath, "1000"))
            '' Qڰ�
            'Dim FreqMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_FLCOND", "001_MIN", m_sPath, "0.1"))
            'Dim FreqMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_FLCOND", "001_MAX", m_sPath, "40.0"))
            '' STEG�{��
            'Dim StegMin As Integer = Integer.Parse(GetPrivateProfileString_S("CUT_FLCOND", "002_MIN", m_sPath, "1"))
            'Dim StegMax As Integer = Integer.Parse(GetPrivateProfileString_S("CUT_FLCOND", "002_MAX", m_sPath, "15"))
            '' �ڕW�p���[�i�v�j
            'Dim dblPowerAdjustTargetMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_FLCOND", "003_MIN", m_sPath, "0.01"))
            'Dim dblPowerAdjustTargetMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_FLCOND", "003_MAX", m_sPath, "20.0"))
            '' ���e�͈́i�}�v�j
            'Dim dblPowerAdjustToleLevelMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_FLCOND", "004_MIN", m_sPath, "0.01"))
            'Dim dblPowerAdjustToleLevelMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_FLCOND", "004_MAX", m_sPath, "10.0"))


            ' Qڰ�
            'V2.1.0.0�@            Dim intQF2Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "012_MIN", m_sPath, "0.1"))
            'V2.1.0.0�@            Dim intQF2Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "012_MAX", m_sPath, "40.0"))
            ' ���x
            'V2.1.0.0�@            Dim dblV2Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "013_MIN", m_sPath, "0.1"))
            'V2.1.0.0�@            Dim dblV2Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "013_MAX", m_sPath, "400.0"))
            ' ����
            Dim intLetterLenMin As Integer = Integer.Parse(GetPrivateProfileString_S("CUT_CUT", "011_MIN", m_sPath, "1"))
            Dim intLetterLenMax As Integer = Integer.Parse(GetPrivateProfileString_S("CUT_CUT", "011_MAX", m_sPath, "10"))

            ' �U�_�^�[���|�C���g�k�J�b�g�g���~���O
            ' �J�b�g��
            Dim dCutLenMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_LCUT_PARA", ("000_MIN"), m_sPath, "0.001"))
            Dim dCutLenMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_LCUT_PARA", ("000_MAX"), m_sPath, "20.000"))
            ' Q���[�g
            Dim dQRateMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_LCUT_PARA", ("001_MIN"), m_sPath, "0.1"))
            Dim dQRateMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_LCUT_PARA", ("001_MAX"), m_sPath, "40.0"))
            ' ���x
            Dim dSpeedMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_LCUT_PARA", ("002_MIN"), m_sPath, "0.1"))
            Dim dSpeedMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_LCUT_PARA", ("002_MAX"), m_sPath, "400.0"))
            ' L����߲��
            Dim dTurnPointMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_LCUT_PARA", ("003_MIN"), m_sPath, "0.0"))
            Dim dTurnPointMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_LCUT_PARA", ("003_MAX"), m_sPath, "100.0"))



            ' �X�g���[�g�J�b�g�{�� 'V2.1.0.0�@ �J�b�g���̒�R�l�ω��ʔ���@�\���ڒǉ��@�R���ڃV�t�g����012_����015_�֕ύX
            ' �X�g���[�g�J�b�g�{�� 'V2.2.1.7�@ �J�b�g���̒�R�l�ω��ʔ���@�\���ڒǉ��@�R���ڃV�t�g����015_����018_�֕ύX
            Dim intRetraceCntMin As Short = Short.Parse(GetPrivateProfileString_S("CUT_CUT", ("018_MIN"), m_sPath, "1"))
            Dim intRetraceCntMax As Short = Short.Parse(GetPrivateProfileString_S("CUT_CUT", ("018_MAX"), m_sPath, "10"))
            ' ���g���[�X�̃I�t�Z�b�g�w
            Dim dblRetraceOffXMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_RETRACE_PARA", ("000_MIN"), m_sPath, "-10.0"))
            Dim dblRetraceOffXMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_RETRACE_PARA", ("000_MAX"), m_sPath, "10.0"))
            ' ���g���[�X�̃I�t�Z�b�g�x
            Dim dblRetraceOffYMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_RETRACE_PARA", ("001_MIN"), m_sPath, "-10.0"))
            Dim dblRetraceOffYMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_RETRACE_PARA", ("001_MAX"), m_sPath, "10.0"))
            ' �X�g���[�g�J�b�g�E���g���[�X��Q���[�g
            Dim dblRetraceQrateMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_RETRACE_PARA", ("002_MIN"), m_sPath, "0.1"))
            Dim dblRetraceQrateMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_RETRACE_PARA", ("002_MAX"), m_sPath, "40.0"))
            ' �X�g���[�g�J�b�g�E���g���[�X�̃g�������x
            Dim dblRetraceSpeedMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_RETRACE_PARA", ("003_MIN"), m_sPath, "0.1"))
            Dim dblRetraceSpeedMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_RETRACE_PARA", ("003_MAX"), m_sPath, "400.0"))

            'V2.2.0.0�A��
            'U�J�b�g�f�[�^�̏���l�擾 
            Dim dblUcutLen1Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("000_MIN"), m_sPath, "0.0"))
            Dim dblUcutLen1Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("000_MAX"), m_sPath, "20.0"))
            Dim dblUcutLen2Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("001_MIN"), m_sPath, "0.0"))
            Dim dblUcutLen2Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("001_MAX"), m_sPath, "20.0"))
            Dim dblUcutR1Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("002_MIN"), m_sPath, "0.0"))
            Dim dblUcutR1Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("002_MAX"), m_sPath, "20.0"))
            Dim dblUcutR2Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("003_MIN"), m_sPath, "0.0"))
            Dim dblUcutR2Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("003_MAX"), m_sPath, "20.0"))
            Dim dblUcutQMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("004_MIN"), m_sPath, "0.1"))
            Dim dblUcutQMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("004_MAX"), m_sPath, "40.0"))
            Dim dblUcutSpdMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("005_MIN"), m_sPath, "0.1"))
            Dim dblUcutSpdMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("005_MAX"), m_sPath, "400.0"))
            Dim dblUcutLturnMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("006_MIN"), m_sPath, "0.0"))
            Dim dblUcutLturnMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("006_MAX"), m_sPath, "100.0"))
            'V2.2.0.0�A��


            With m_MainEdit
                For iRn As Integer = 1 To .W_PLT.RCount                     ' ��R�����J�Ԃ�
                    If UserModule.IsCutResistorIncMarking(.W_REG, iRn) Then
                        For iCn As Integer = 1 To .W_REG(iRn).intTNN            ' �J�b�g�����J�Ԃ�
                            strMSG = "��R�ԍ�" & iRn.ToString & " �J�b�g�ԍ�" & iCn.ToString
                            m_ResNo = iRn : m_CutNo = iCn

                            If .W_REG(iRn).STCUT(iCn).intCTYP = CNS_CUTP_SP And (.W_REG(iRn).STCUT(iCn).intNum < intNumMin Or intNumMax < .W_REG(iRn).STCUT(iCn).intNum) Then          ' �������ݖ{��
                                strMin = intNumMin.ToString : strMax = intNumMax.ToString
                                strMSG = strMSG & "�̖{����" : GoTo ERR_MESSAGE
                            ElseIf (.W_REG(iRn).STCUT(iCn).intQF1 / 10.0) < intQF1Min Or intQF1Max < (.W_REG(iRn).STCUT(iCn).intQF1 / 10.0) Then      ' Qڰ�
                                strMin = intQF1Min.ToString : strMax = intQF1Max.ToString
                                strMSG = strMSG & "�p���[�g��" : GoTo ERR_MESSAGE
                            ElseIf .W_REG(iRn).STCUT(iCn).dblV1 < dblV1Min Or dblV1Max < .W_REG(iRn).STCUT(iCn).dblV1 Then          ' ���x
                                strMin = dblV1Min.ToString : strMax = dblV1Max.ToString
                                strMSG = strMSG & "���x��" : GoTo ERR_MESSAGE
                            ElseIf .W_REG(iRn).STCUT(iCn).dblSTX < dblSTXMin Or dblSTXMax < .W_REG(iRn).STCUT(iCn).dblSTX Then      ' ��ĈʒuX
                                strMin = dblSTXMin.ToString : strMax = dblSTXMax.ToString
                                strMSG = strMSG & "�J�b�g�ʒu�w��" : GoTo ERR_MESSAGE
                            ElseIf .W_REG(iRn).STCUT(iCn).dblSTY < dblSTYMin Or dblSTYMax < .W_REG(iRn).STCUT(iCn).dblSTY Then      ' ��ĈʒuY
                                strMin = dblSTYMin.ToString : strMax = dblSTYMax.ToString
                                strMSG = strMSG & "�J�b�g�ʒu�x��" : GoTo ERR_MESSAGE
                            ElseIf .W_REG(iRn).STCUT(iCn).dblSX2 < dblSX2Min Or dblSX2Max < .W_REG(iRn).STCUT(iCn).dblSX2 Then      ' ��Ĉʒu2X
                                strMin = dblSX2Min.ToString : strMax = dblSX2Max.ToString
                                strMSG = strMSG & "�J�b�g�ʒu�w�Q��" : GoTo ERR_MESSAGE
                            ElseIf .W_REG(iRn).STCUT(iCn).dblSY2 < dblSY2Min Or dblSY2Max < .W_REG(iRn).STCUT(iCn).dblSY2 Then      ' ��Ĉʒu2Y
                                strMin = dblSY2Min.ToString : strMax = dblSY2Max.ToString
                                strMSG = strMSG & "�J�b�g�ʒu�x�Q��" : GoTo ERR_MESSAGE
                            ElseIf .W_REG(iRn).STCUT(iCn).dblDL2 < dblDL2Min Or dblDL2Max < .W_REG(iRn).STCUT(iCn).dblDL2 Then      ' ��Ē�1
                                strMin = dblDL2Min.ToString : strMax = dblDL2Max.ToString
                                strMSG = strMSG & "�J�b�g���P��" : GoTo ERR_MESSAGE
                                'ElseIf .W_REG(iRn).STCUT(iCn).intCTYP > 1 And (.W_REG(iRn).STCUT(iCn).dblDL3 < dblDL3Min Or dblDL3Max < .W_REG(iRn).STCUT(iCn).dblDL3) Then      ' ��Ē�2
                                '    strMin = dblDL3Min.ToString : strMax = dblDL3Max.ToString
                                '    strMSG = strMSG & "�J�b�g���Q��" : GoTo ERR_MESSAGE
                            ElseIf .W_REG(iRn).STCUT(iCn).dblCOF < dblCOFMin Or dblCOFMax < .W_REG(iRn).STCUT(iCn).dblCOF Then      ' ��ĵ�
                                strMin = dblCOFMin.ToString : strMax = dblCOFMax.ToString
                                strMSG = strMSG & "�J�b�g�I�t��" : GoTo ERR_MESSAGE
                            ElseIf .W_REG(iRn).STCUT(iCn).dblLTP < dblLTPMin Or dblLTPMax < .W_REG(iRn).STCUT(iCn).dblLTP Then      ' L����߲��
                                strMin = dblLTPMin.ToString : strMax = dblLTPMax.ToString
                                strMSG = strMSG & "�k�^�[���|�C���g��" : GoTo ERR_MESSAGE
                            ElseIf ((.W_REG(iRn).STCUT(iCn).intCTYP = CNS_CUTP_M) And (.W_REG(iRn).intSLP <> SLP_MARK)) Then       ' �����}�[�L���O 'V2.2.1.7�@
                                Dim iLen As Integer = .W_REG(iRn).STCUT(iCn).cFormat.Length
                                If (iLen < intLetterLenMin) Or (intLetterLenMax < iLen) Then
                                    strMin = intLetterLenMin.ToString("0") : strMax = intLetterLenMax.ToString("0")
                                    strMSG = strMSG & "������" : GoTo ERR_MESSAGE
                                End If
                                If .W_REG(iRn).STCUT(iCn).dblDL2 < 0.1 Or .W_REG(iRn).STCUT(iCn).dblDL2 > 10.0 Then
                                    strMin = "1.0" : strMax = "10.0"
                                    strMSG = strMSG & "����������" : GoTo ERR_MESSAGE
                                End If
                            End If
                            ' ARATA
                            If .W_REG(iRn).STCUT(iCn).intCUT = CNS_CUTM_IX Then   ' ��ĕ��@(1:�ׯ�ݸ�, 2:���ޯ��, 3:NG, 4:�����J�b�g(FULL CUT))
                                For idx As Integer = 1 To MAXIDX
                                    If .W_REG(iRn).STCUT(iCn).intIXN(idx) < intIXNMin Or intIXNMax < .W_REG(iRn).STCUT(iCn).intIXN(idx) Then                ' ��ĉ�
                                        strMin = intIXNMin.ToString : strMax = intIXNMax.ToString
                                        strMSG = strMSG & "�C���f�b�N�X�ԍ�" & idx.ToString & "�̃C���f�b�N�X�J�b�g����" : GoTo ERR_MESSAGE
                                    Else
                                        If .W_REG(iRn).STCUT(iCn).intIXN(idx) > 0 Then
                                            If .W_REG(iRn).STCUT(iCn).dblDL1(idx) < dblDL1Min Or dblDL1Max < .W_REG(iRn).STCUT(iCn).dblDL1(idx) Then        ' ��Ē�(���ޯ���߯�)
                                                strMin = dblDL1Min.ToString : strMax = dblDL1Max.ToString
                                                strMSG = strMSG & "�C���f�b�N�X�ԍ�" & idx.ToString & "�̃J�b�g����" : GoTo ERR_MESSAGE
                                            ElseIf .W_REG(iRn).STCUT(iCn).lngPAU(idx) < lngPAUMin Or lngPAUMax < .W_REG(iRn).STCUT(iCn).lngPAU(idx) Then    ' �߯����߰��
                                                strMin = lngPAUMin.ToString : strMax = lngPAUMax.ToString
                                                strMSG = strMSG & "�C���f�b�N�X�ԍ�" & idx.ToString & "�̃s�b�`�ԃ|�[�Y��" : GoTo ERR_MESSAGE
                                            ElseIf .W_REG(iRn).STCUT(iCn).dblDEV(idx) < dblDEVMin Or dblDEVMax < .W_REG(iRn).STCUT(iCn).dblDEV(idx) Then    ' �덷
                                                strMin = dblDEVMin.ToString : strMax = dblDEVMax.ToString
                                                strMSG = strMSG & "�C���f�b�N�X�ԍ�" & idx.ToString & "�̌덷��" : GoTo ERR_MESSAGE
                                            End If
                                        End If
                                    End If
                                Next
                            End If

                            If .W_REG(iRn).STCUT(iCn).intCTYP = CNS_CUTP_L Then   ' �U�_�^�[���|�C���g�k�J�b�g�g���~���O
                                For idx As Integer = 1 To MAX_LCUT
                                    If .W_REG(iRn).STCUT(iCn).dCutLen(idx) < dCutLenMin Or dCutLenMax < .W_REG(iRn).STCUT(iCn).dCutLen(idx) Then
                                        strMin = dCutLenMin.ToString : strMax = dCutLenMax.ToString
                                        strMSG = strMSG & "�ԍ�" & idx.ToString & "�̃J�b�g����" : GoTo ERR_MESSAGE
                                    ElseIf .W_REG(iRn).STCUT(iCn).dQRate(idx) / 10.0 < dQRateMin Or dQRateMax < .W_REG(iRn).STCUT(iCn).dQRate(idx) / 10.0 Then
                                        strMin = dQRateMin.ToString : strMax = dQRateMax.ToString
                                        strMSG = strMSG & "�ԍ�" & idx.ToString & "�̂p���[�g��" : GoTo ERR_MESSAGE
                                    ElseIf .W_REG(iRn).STCUT(iCn).dSpeed(idx) < dSpeedMin Or dSpeedMax < .W_REG(iRn).STCUT(iCn).dSpeed(idx) Then
                                        strMin = dSpeedMin.ToString : strMax = dSpeedMax.ToString
                                        strMSG = strMSG & "�ԍ�" & idx.ToString & "�̑��x��" : GoTo ERR_MESSAGE
                                    ElseIf idx < MAX_LCUT And (.W_REG(iRn).STCUT(iCn).dTurnPoint(idx) < dTurnPointMin Or dTurnPointMax < .W_REG(iRn).STCUT(iCn).dTurnPoint(idx)) Then
                                        strMin = dTurnPointMin.ToString : strMax = dTurnPointMax.ToString
                                        strMSG = strMSG & "�ԍ�" & idx.ToString & "�̃^�[���|�C���g��" : GoTo ERR_MESSAGE
                                    End If
                                Next
                            End If
                            If .W_REG(iRn).STCUT(iCn).intCTYP = CNS_CUTP_ST_TR Then   ' �X�g���[�g�J�b�g�E���g���[�X
                                If .W_REG(iRn).STCUT(iCn).intRetraceCnt < intRetraceCntMin Or intRetraceCntMax < .W_REG(iRn).STCUT(iCn).intRetraceCnt Then      ' Qڰ�'V2.1.0.0�@intQF1Min��intRetraceCntMin,intQF1Max��intRetraceCntMax�C��
                                    strMin = intRetraceCntMin.ToString : strMax = intRetraceCntMax.ToString
                                    strMSG = strMSG & "���g���[�X�{����" : GoTo ERR_MESSAGE
                                Else
                                    For idx As Integer = 1 To .W_REG(iRn).STCUT(iCn).intRetraceCnt

                                        If .W_REG(iRn).STCUT(iCn).dblRetraceOffX(idx) < dblRetraceOffXMin Or dblRetraceOffXMax < .W_REG(iRn).STCUT(iCn).dblRetraceOffX(idx) Then
                                            strMin = dblRetraceOffXMin.ToString : strMax = dblRetraceOffXMax.ToString
                                            strMSG = strMSG & "�ԍ�" & idx.ToString & "�̃��g���[�X�̃I�t�Z�b�g�w��" : GoTo ERR_MESSAGE
                                        ElseIf .W_REG(iRn).STCUT(iCn).dblRetraceOffY(idx) < dblRetraceOffYMin Or dblRetraceOffYMax < .W_REG(iRn).STCUT(iCn).dblRetraceOffY(idx) Then
                                            strMin = dblRetraceOffYMin.ToString : strMax = dblRetraceOffYMax.ToString
                                            strMSG = strMSG & "�ԍ�" & idx.ToString & "�̃��g���[�X�̃I�t�Z�b�g�x��" : GoTo ERR_MESSAGE
                                        ElseIf .W_REG(iRn).STCUT(iCn).dblRetraceQrate(idx) / 10.0 < dblRetraceQrateMin Or dblRetraceQrateMax < .W_REG(iRn).STCUT(iCn).dblRetraceQrate(idx) / 10.0 Then
                                            strMin = dblRetraceQrateMin.ToString : strMax = dblRetraceQrateMax.ToString
                                            strMSG = strMSG & "�ԍ�" & idx.ToString & "�̃��g���[�X�p���[�g��" : GoTo ERR_MESSAGE
                                        ElseIf .W_REG(iRn).STCUT(iCn).dblRetraceSpeed(idx) < dblRetraceSpeedMin Or dblRetraceSpeedMax < .W_REG(iRn).STCUT(iCn).dblRetraceSpeed(idx) Then
                                            strMin = dblRetraceSpeedMin.ToString : strMax = dblRetraceSpeedMax.ToString
                                            strMSG = strMSG & "�ԍ�" & idx.ToString & "�̃��g���[�X�̑��x��" : GoTo ERR_MESSAGE
                                        End If
                                    Next
                                End If
                            End If
#If cOSCILLATORcFLcUSE And cFLcAUTOcPOWER Then
                            If (gSysPrm.stRAT.giOsc_Res = OSCILLATOR_FL) Then
                                For idx As Integer = 1 To MAXCND

                                    'If .W_FLCND.Curr(.W_REG(iRn).STCUT(iCn).intCND(idx)) < CurrMin Or CurrMax < .W_FLCND.Curr(.W_REG(iRn).STCUT(iCn).intCND(idx)) Then                ' �d���l
                                    '    strMin = CurrMin.ToString : strMax = CurrMax.ToString
                                    '    strMSG = strMSG & "�J�b�g�����ԍ�" & idx.ToString & "�̓d���l��" : GoTo ERR_MESSAGE
                                    'ElseIf .W_FLCND.Freq(.W_REG(iRn).STCUT(iCn).intCND(idx)) < FreqMin Or FreqMax < .W_FLCND.Freq(.W_REG(iRn).STCUT(iCn).intCND(idx)) Then            ' STEG�{��
                                    '    strMin = FreqMin.ToString : strMax = FreqMax.ToString
                                    '    strMSG = strMSG & "�J�b�g�����ԍ�" & idx.ToString & "�̂p���[�g��" : GoTo ERR_MESSAGE
                                    'ElseIf .W_FLCND.Steg(.W_REG(iRn).STCUT(iCn).intCND(idx)) < StegMin Or StegMax < .W_FLCND.Steg(.W_REG(iRn).STCUT(iCn).intCND(idx)) Then            ' Qڰ�
                                    '    strMin = StegMin.ToString : strMax = StegMax.ToString
                                    '    strMSG = strMSG & "�J�b�g�����ԍ�" & idx.ToString & "�̂r�s�d�f�{����" : GoTo ERR_MESSAGE
                                    'End If

                                    dblPowerAdjustTargetMax = ObjFiberLaser.GetMaxPower(.W_FLCND.Freq(.W_REG(iRn).STCUT(iCn).intCND(idx)), .W_FLCND.Steg(.W_REG(iRn).STCUT(iCn).intCND(idx)))
                                    If .W_FLCND.dblPowerAdjustTarget(.W_REG(iRn).STCUT(iCn).intCND(idx)) < dblPowerAdjustTargetMin Or dblPowerAdjustTargetMax < .W_FLCND.dblPowerAdjustTarget(.W_REG(iRn).STCUT(iCn).intCND(idx)) Then                    ' �ڕW�p���[�i�v�j
                                        strMin = dblPowerAdjustTargetMin.ToString : strMax = dblPowerAdjustTargetMax.ToString
                                        strMSG = strMSG & "�J�b�g�����ԍ�" & idx.ToString & "�̖ڕW�p���[�i�v�j��" : GoTo ERR_MESSAGE
                                    ElseIf .W_FLCND.dblPowerAdjustToleLevel(.W_REG(iRn).STCUT(iCn).intCND(idx)) < dblPowerAdjustToleLevelMin Or dblPowerAdjustToleLevelMax < .W_FLCND.dblPowerAdjustToleLevel(.W_REG(iRn).STCUT(iCn).intCND(idx)) Then    ' ���e�͈́i�}�v�j
                                        strMin = dblPowerAdjustToleLevelMin.ToString : strMax = dblPowerAdjustToleLevelMax.ToString
                                        strMSG = strMSG & "�J�b�g�����ԍ�" & idx.ToString & "�̋��e�͈́i�}�v�j��" : GoTo ERR_MESSAGE
                                    End If

                                    If .W_REG(iRn).STCUT(iCn).intCTYP = CNS_CUTP_ST And idx = 1 Then  ' �X�g���[�g
                                        Exit For
                                    ElseIf .W_REG(iRn).STCUT(iCn).intCTYP = CNS_CUTP_L And idx = 2 Then  ' �X�g���[�g
                                        Exit For
                                    End If
                                Next
                            End If
#End If
                            'V2.2.0.0�A��
                            'U�J�b�g�p�����[�^�̃`�F�b�N
                            If .W_REG(iRn).STCUT(iCn).intCTYP = CNS_CUTP_U Then   ' U�J�b�g�g���~���O
                                If .W_REG(iRn).STCUT(iCn).dUCutL1 < dblUcutLen1Min Or dblUcutLen1Max < .W_REG(iRn).STCUT(iCn).dUCutL1 Then      ' L1�J�b�g��
                                    strMin = dblUcutLen1Min.ToString : strMax = dblUcutLen1Max.ToString
                                    strMSG = strMSG & "L1�J�b�g����" : GoTo ERR_MESSAGE
                                End If
                                If .W_REG(iRn).STCUT(iCn).dUCutL2 < dblUcutLen1Min Or dblUcutLen1Max < .W_REG(iRn).STCUT(iCn).dUCutL2 Then      ' L2�J�b�g��
                                    strMin = dblUcutLen2Min.ToString : strMax = dblUcutLen2Max.ToString
                                    strMSG = strMSG & "L2�J�b�g����" : GoTo ERR_MESSAGE
                                End If
                                If .W_REG(iRn).STCUT(iCn).dblUCutR1 < dblUcutR1Min Or dblUcutR1Max < .W_REG(iRn).STCUT(iCn).dblUCutR1 Then     'R1
                                    strMin = dblUcutR1Min.ToString : strMax = dblUcutR1Max.ToString
                                    strMSG = strMSG & "R1���a��" : GoTo ERR_MESSAGE
                                End If
                                If .W_REG(iRn).STCUT(iCn).dblUCutR2 < dblUcutR2Min Or dblUcutR2Max < .W_REG(iRn).STCUT(iCn).dblUCutR2 Then     'R2
                                    strMin = dblUcutR2Min.ToString : strMax = dblUcutR2Max.ToString
                                    strMSG = strMSG & "R2���a��" : GoTo ERR_MESSAGE
                                End If
                                If (.W_REG(iRn).STCUT(iCn).intUCutQF1 / 10.0) < dblUcutQMin Or dblUcutQMax < (.W_REG(iRn).STCUT(iCn).intUCutQF1 / 10.0) Then     'Q���[�g
                                    strMin = dblUcutQMin.ToString : strMax = dblUcutQMax.ToString
                                    strMSG = strMSG & "Q���[�g��" : GoTo ERR_MESSAGE
                                End If
                                If .W_REG(iRn).STCUT(iCn).dblUCutV1 < dblUcutSpdMin Or dblUcutSpdMax < .W_REG(iRn).STCUT(iCn).dblUCutV1 Then     '���x
                                    strMin = dblUcutSpdMin.ToString : strMax = dblUcutSpdMax.ToString
                                    strMSG = strMSG & "���x��" : GoTo ERR_MESSAGE
                                End If
                                If .W_REG(iRn).STCUT(iCn).dblUCutTurnP < dblUcutLturnMin Or dblUcutLturnMax < .W_REG(iRn).STCUT(iCn).dblUCutTurnP Then     '
                                    strMin = dblUcutSpdMin.ToString : strMax = dblUcutSpdMax.ToString
                                    strMSG = strMSG & "�^�[���|�C���g��" : GoTo ERR_MESSAGE
                                End If
                            End If
                            'V2.2.0.0�A��

                        Next iCn
                    End If
                Next iRn
            End With

            Exit Function
ERR_MESSAGE:
            m_MainEdit.MTab.SelectedIndex = m_TabIdx  ' ��ޕ\���ؑ�
            Call SetDataToText()    ' ���������R�ԍ��A��Ĕԍ����ް�����۰قɾ�Ă���
            CutDataCheckByDataOnly = 1
            strMSG = strMSG & strMin & "�`" & strMax & "�͈̔͂Ŏw�肵�ĉ�����"
            Call MsgBox(strMSG, DirectCast( _
                        MsgBoxStyle.OkOnly + _
                        MsgBoxStyle.Information, MsgBoxStyle), _
                        My.Application.Info.Title)

        End Function
#End Region
#Else
        ''' <summary>���ׂĂ�÷���ޯ�����ް������������Ȃ�</summary>
        ''' <returns>0=����, 1=�װ</returns>
        Friend Overrides Function CheckAllTextData() As Integer
            Dim ret As Integer
            Try
                m_CheckFlg = True ' ������(tabBase_Layout�ɂĎg�p)
                With m_MainEdit
                    .MTab.SelectedIndex = m_TabIdx  ' ��ޕ\���ؑ�

                    For rn As Integer = 1 To .W_PLT.RCount Step 1
                        m_ResNo = rn
                        With .W_REG(rn)

                            ' ------------------------
                            If (SLP_VMES = m_MainEdit.W_REG(m_ResNo).intSLP) OrElse (SLP_RMES = m_MainEdit.W_REG(m_ResNo).intSLP) Then
                                ' ��R�̽۰�߂� 5:�d������̂�, 6:��R����̂� �̏ꍇ
                                Continue For    ' ��Ă������͂����Ȃ�Ȃ�(���̒�R��)
                            End If
                            ' ------------------------

                            ' TODO: ��Đ���0�ɂȂ邱�Ƃ͂Ȃ��d�l�̂��ߕs�v�Ǝv����
                            If (.intTNN < 1) Then ' ��Đ� < 1 ?
                                Dim strMsg As String
                                strMsg = "��R�ԍ�" & rn.ToString("0") & "�̃J�b�g�f�[�^������܂���B" & vbCrLf
                                strMsg = strMsg & "�J�b�g�f�[�^��o�^���Ă��������B"
                                Call MsgBox(strMsg, MsgBoxStyle.OkOnly, My.Application.Info.Title)
                                Exit Try
                            End If

                            ' ��Đ����J�Ԃ�
                            For cn As Integer = 1 To .intTNN Step 1
                                m_CutNo = cn
                                With .STCUT(cn)

                                    ' ���������R�ԍ��A��Ĕԍ����ް�����۰قɾ�Ă���
                                    Call SetDataToText()

                                    ' ��ĸ�ٰ���ޯ��
                                    ret = CheckControlData(m_CtlCut)
                                    If (ret <> 0) Then Exit Try

                                    ' ��ĕ��@�����ޯ����Ă̏ꍇ(1:�ׯ�ݸ�, 2:���ޯ��, 3:NG���)
                                    If (CNS_CUTM_IX = .intCUT) Then
                                        ' ���ޯ����ĸ�ٰ���ޯ��
                                        ret = CheckControlData(m_CtlIdxCut)
                                        If (ret <> 0) Then Exit Try
                                    End If

                                    ' ��������
                                    ret = CheckRelation()
                                    If (ret <> 0) Then Exit Try
                                End With
                            Next cn

                        End With
                    Next rn
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
                ret = 1
            Finally
                m_CheckFlg = False ' �����I��
                CheckAllTextData = ret
            End Try

        End Function
#End If
#End Region

#Region "�ް������֐����Ăяo��"
        ''' <summary>÷���ޯ�����ް������֐����Ăяo��</summary>
        ''' <param name="cTextBox">��������÷���ޯ��</param>
        ''' <returns>0=����, 1=�װ</returns>
        Protected Overrides Function CheckTextData(ByRef cTextBox As cTxt_) As Integer
            Dim strMsg As String
            Dim tag As Integer
            Dim ret As Integer
            Dim dblWK As Double
            Dim i As Integer
            Try
                ' ��R�ް��o�^������
                ' TODO: ��R����0�ɂȂ邱�Ƃ͂Ȃ��d�l�̂��ߕs�v�Ǝv����
                If (m_ResNo < 1) Then
                    strMsg = "��R�f�[�^������܂���B��R�f�[�^��o�^���Ă��������B"
                    Call MsgBox(strMsg, MsgBoxStyle.OkOnly, My.Application.Info.Title)
                    ret = 1
                    Exit Try
                End If

                With m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo)
                    ' ��Đ�����
                    ' TODO: ��Đ���0�ɂȂ邱�Ƃ͂Ȃ��d�l�̂��ߕs�v�Ǝv����
                    If (m_MainEdit.W_REG(m_ResNo).intTNN < 1) Then ' ��Đ� < 1 ?
                        strMsg = "��R�ԍ�" & m_ResNo.ToString("0") & "�̃J�b�g�f�[�^������܂���B" & vbCrLf
                        strMsg = strMsg & "�ǉ��{�^�����������ăJ�b�g�f�[�^��o�^���Ă��������B"
                        Call MsgBox(strMsg, MsgBoxStyle.OkOnly, My.Application.Info.Title)
                        ret = 1
                        Exit Try
                    End If

                    tag = DirectCast(cTextBox.Tag, Integer)
                    Select Case (DirectCast(cTextBox.Parent.Tag, Integer))
                        ' ------------------------------------------------------------------------------
                        Case 0 ' ��ĸ�ٰ���ޯ��
                            Select Case (tag)
                                Case 0 ' �������ݖ{��(�������ݶ�Ď��L��)
                                    If (CNS_CUTP_SP = .intCTYP) Then ' �������ݶ�Ă̏ꍇ
                                        ret = CheckShortData(cTextBox, .intNum)
                                    End If
                                Case 1 ' Qڰ�
                                    If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then ' ̧���ڰ�ނł͂Ȃ��ꍇ
                                        dblWK = .intQF1 / 10.0
                                        ret = CheckDoubleData(cTextBox, dblWK)
                                        .intQF1 = Convert.ToInt16(dblWK * 10) ' (KHz��0.1KHz)
                                    End If
                                Case 2 ' ���x
                                    ret = CheckDoubleData(cTextBox, .dblV1)
                                Case 3 ' ��ĈʒuX
                                    ret = CheckDoubleData(cTextBox, .dblSTX)
                                Case 4 ' ��ĈʒuY
                                    ret = CheckDoubleData(cTextBox, .dblSTY)
                                Case 5 ' ��Ĉʒu2X(�������ݶ�Ď��L��)
                                    If (CNS_CUTP_ST_TR = .intCTYP) Then 'V1.0.4.3�B ���g���[�X�̏ꍇ�@�T�[�y���^�C���iCNS_CUTP_SP�j����ύX
                                        ret = CheckDoubleData(cTextBox, .dblSX2)
                                    End If
                                Case 6 ' ��Ĉʒu2Y(�������ݶ�Ď��L��)
                                    If (CNS_CUTP_ST_TR = .intCTYP) Then 'V1.0.4.3�B ���g���[�X�̏ꍇ�@�T�[�y���^�C���iCNS_CUTP_SP�j����ύX
                                        ret = CheckDoubleData(cTextBox, .dblSY2)
                                    End If
                                Case 7 ' ��Ē�1
                                    ret = CheckDoubleData(cTextBox, .dblDL2)
                                Case 8 ' ��Ē�2
                                    If (CNS_CUTP_ST <> .intCTYP) Then ' ��ڰĶ�Ăł͂Ȃ��ꍇ
                                        ret = CheckDoubleData(cTextBox, .dblDL3)
                                    End If
                                Case 9 ' ��ĵ�
                                    If (CNS_CUTM_NG <> .intCUT) Then ' NG��Ăł͂Ȃ��ꍇ
                                        ret = CheckDoubleData(cTextBox, .dblCOF)
                                    End If
                                Case 10 ' L����߲��
                                    If (CNS_CUTP_L = .intCTYP) Then ' L��Ă̏ꍇ
                                        ret = CheckDoubleData(cTextBox, .dblLTP)
                                    End If
                                Case 11 ' ###1042�@�@���̓`�F�b�N�ǉ�����
                                    If (CNS_CUTP_M = .intCTYP) Then ' �����}�[�L���O
                                        ret = CheckStrData(cTextBox, .cFormat)
                                    End If
                                    'V2.2.1.7�@ ��
                                Case 12 ' �󎚌Œ蕔
                                    If (CNS_CUTP_M = .intCTYP) Then ' �����}�[�L���O
                                        ret = CheckStrData(cTextBox, .cMarkFix)
                                    End If
                                Case 13 ' �J�n�ԍ�
                                    If (CNS_CUTP_M = .intCTYP) Then ' �����}�[�L���O
                                        Dim Inputstr = cTextBox.Text.Trim()
                                        If Inputstr <> "" Then
                                            ret = CheckNumeric(cTextBox)
                                            If (ret <> -1) Then
                                                ret = CheckStrData(cTextBox, .cMarkStartNum)
                                            Else
                                                cTextBox.Text = .cMarkStartNum
                                            End If
                                        Else
                                            .cMarkStartNum = ""
                                        End If
                                    End If
                                Case 14 ' �d����
                                    If (CNS_CUTP_M = .intCTYP) Then ' �����}�[�L���O
                                        ret = CheckShortData(cTextBox, .intMarkRepeatCnt)
                                    End If
                                                                        'V2.1.0.0�@��
                                Case 15 ' �㏸��
                                    ret = CheckDoubleData(cTextBox, .dRateOfUp)
                                    If ret = 0 And .iVariationRepeat = 1 Then
                                        UserSub.CutVariationDataCopy(m_MainEdit.W_PLT, m_MainEdit.W_REG, m_ResNo, m_CutNo)
                                    End If
                                Case 16 ' �����l
                                    ret = CheckDoubleData(cTextBox, .dVariationLow)
                                    If ret = 0 And .iVariationRepeat = 1 Then
                                        UserSub.CutVariationDataCopy(m_MainEdit.W_PLT, m_MainEdit.W_REG, m_ResNo, m_CutNo)
                                    End If
                                Case 17 ' ����l
                                    ret = CheckDoubleData(cTextBox, .dVariationHi)
                                    If ret = 0 And .iVariationRepeat = 1 Then
                                        UserSub.CutVariationDataCopy(m_MainEdit.W_PLT, m_MainEdit.W_REG, m_ResNo, m_CutNo)
                                    End If
                                    'V2.1.0.0�@��

                                    '    'V2.1.0.0�@��
                                    'Case 12 ' �㏸��
                                    '    ret = CheckDoubleData(cTextBox, .dRateOfUp)
                                    '    If ret = 0 And .iVariationRepeat = 1 Then
                                    '        UserSub.CutVariationDataCopy(m_MainEdit.W_PLT, m_MainEdit.W_REG, m_ResNo, m_CutNo)
                                    '    End If
                                    'Case 13 ' �����l
                                    '    ret = CheckDoubleData(cTextBox, .dVariationLow)
                                    '    If ret = 0 And .iVariationRepeat = 1 Then
                                    '        UserSub.CutVariationDataCopy(m_MainEdit.W_PLT, m_MainEdit.W_REG, m_ResNo, m_CutNo)
                                    '    End If
                                    'Case 14 ' ����l
                                    '    ret = CheckDoubleData(cTextBox, .dVariationHi)
                                    '    If ret = 0 And .iVariationRepeat = 1 Then
                                    '        UserSub.CutVariationDataCopy(m_MainEdit.W_PLT, m_MainEdit.W_REG, m_ResNo, m_CutNo)
                                    '    End If
                                    '    'V2.1.0.0�@��
                                    'V2.2.1.7�@ ��
                                    'V1.0.4.3�D��
                                    'V2.1.0.0�@                                Case 12 ' �X�g���[�g�J�b�g�E���g���[�X
                                    'V2.1.0.0�@                                    If (CNS_CUTP_ST_TR = .intCTYP) Then ' �X�g���[�g�J�b�g�E���g���[�X�̏ꍇ
                                    'V2.1.0.0�@                                        If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then ' ̧���ڰ�ނł͂Ȃ��ꍇ
                                    'V2.1.0.0�@                                            ret = CheckDoubleData(cTextBox, dblWK)
                                    'V2.1.0.0�@                                            .intQF2 = Convert.ToInt16(dblWK * 10) ' (KHz��0.1KHz)
                                    'V2.1.0.0�@                                        End If
                                    'V2.1.0.0�@                                    End If
                                    'V2.1.0.0�@                                Case 13 ' ���x�Q
                                    'V2.1.0.0�@                                    If (CNS_CUTP_ST_TR = .intCTYP) Then ' �X�g���[�g�J�b�g�E���g���[�X�̏ꍇ
                                    'V2.1.0.0�@                                        ret = CheckDoubleData(cTextBox, .dblV2)
                                    'V2.1.0.0�@                                    End If
                                    'V1.0.4.3�D��
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case 1 ' ���ޯ����ĸ�ٰ���ޯ��
                            If (CNS_CUTM_IX = .intCUT) Then ' ��ĕ��@�����ޯ����Ă̏ꍇ
                                Select Case (tag)
                                    Case 0 ' ��Đ�
                                        Dim cutNo As Integer = (GetCtlIdx(cTextBox, tag) + 1)
                                        ret = CheckShortData(cTextBox, .intIXN(cutNo))
                                    Case 1 ' ��Ē�(���ޯ���߯�)
                                        Dim cutNo As Integer = (GetCtlIdx(cTextBox, tag) + 1)
                                        ret = CheckDoubleData(cTextBox, .dblDL1(cutNo))
                                    Case 2 ' �߰��
                                        Dim cutNo As Integer = (GetCtlIdx(cTextBox, tag) + 1)
                                        ret = CheckIntData(cTextBox, .lngPAU(cutNo))
                                    Case 3 ' �덷
                                        Dim cntNo As Integer = (GetCtlIdx(cTextBox, tag) + 1)
                                        ret = CheckDoubleData(cTextBox, .dblDEV(cntNo))
                                    Case Else
                                        Throw New Exception("Case " & tag & ": Nothing")
                                End Select
                            End If
                            ' ------------------------------------------------------------------------------
                        Case 2 ' FL���H������ٰ���ޯ��(�\���݂̂̂��������Ȃ�)
                            Throw New Exception("Parent.Tag - Case " & tag & ": Nothing")
                            ' ------------------------------------------------------------------------------
                        Case 3 '�k�J�b�g�p�����[�^�O���[�v�{�b�N�X 
                            If (CNS_CUTP_L = .intCTYP) Then ' �J�b�g���@���A�k�J�b�g�̏ꍇ
                                Dim cutNo As Integer = (GetCtlLCutIdx(cTextBox, tag) + 1)
                                Select Case (tag)
                                    Case 0 ' �J�b�g��

                                        ret = CheckDoubleData(cTextBox, .dCutLen(cutNo))
                                        .dblDL2 = .dCutLen(1)
                                        .dblDL3 = 0.0
                                        For i = 2 To MAX_LCUT
                                            .dblDL3 = .dblDL3 + .dCutLen(i)
                                        Next
                                    Case 1 ' �p���[�g
                                        If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then ' ̧���ڰ�ނł͂Ȃ��ꍇ
                                            dblWK = .dQRate(cutNo) / 10.0
                                            ret = CheckDoubleData(cTextBox, dblWK)
                                            .dQRate(cutNo) = Convert.ToInt16(dblWK * 10.0) ' (KHz��0.1KHz)
                                        End If
                                    Case 2 ' ���x
                                        ret = CheckDoubleData(cTextBox, .dSpeed(cutNo))
                                    Case 3 ' �^�[���|�C���g
                                        ret = CheckDoubleData(cTextBox, .dTurnPoint(cutNo))
                                    Case Else
                                        Throw New Exception("Case " & tag & ": Nothing")
                                End Select
                            End If
                            'V2.0.0.0�F��
                        Case 4 '���g���[�X�J�b�g�J�b�g�p�����[�^�O���[�v�{�b�N�X 
                            If (CNS_CUTP_ST_TR = .intCTYP) Then ' �J�b�g���@���A���g���[�X�J�b�g�̏ꍇ
                                If (m_CtlCut(CUT_RETRACE) Is cTextBox) Then
                                    ret = CheckShortData(cTextBox, .intRetraceCnt)
                                Else
                                    Dim cutNo As Integer = (GetCtlRetraceCutIdx(cTextBox, tag) + 1)
                                    Select Case (tag)
                                        Case 0 ' ���g���[�X�̃I�t�Z�b�g�w
                                            ret = CheckDoubleData(cTextBox, .dblRetraceOffX(cutNo))
                                        Case 1 ' ���g���[�X�̃I�t�Z�b�g�x
                                            ret = CheckDoubleData(cTextBox, .dblRetraceOffY(cutNo))
                                        Case 2 ' �X�g���[�g�J�b�g�E���g���[�X��Q���[�g(0.1KHz)�Ɏg�p
                                            If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then ' ̧���ڰ�ނł͂Ȃ��ꍇ
                                                dblWK = .dblRetraceQrate(cutNo) / 10.0
                                                ret = CheckDoubleData(cTextBox, dblWK)
                                                .dblRetraceQrate(cutNo) = Convert.ToInt16(dblWK * 10.0) ' (KHz��0.1KHz)
                                            End If
                                        Case 3 ' ���x
                                            ret = CheckDoubleData(cTextBox, .dblRetraceSpeed(cutNo))
                                        Case Else
                                            Throw New Exception("Case(Retrace) " & tag & ": Nothing")
                                    End Select
                                End If
                            End If
                            'V2.0.0.0�F��
                            'V2.2.0.0�A��
                        Case 5  ' �t�J�b�g�p�����[�^ 
                            If (.intCTYP = CNS_CUTP_U) Then

                                Select Case (tag)
                                    Case 0 ' �J�b�g��
                                        ret = CheckDoubleData(cTextBox, .dUCutL1)
                                    Case 1 ' �J�b�g��
                                        ret = CheckDoubleData(cTextBox, .dUCutL2)
                                    Case 2 ' R1
                                        ret = CheckDoubleData(cTextBox, .dblUCutR1)
                                    Case 3 ' R2
                                        ret = CheckDoubleData(cTextBox, .dblUCutR2)
                                    Case 4 ' �p���[�g
                                        If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then ' ̧���ڰ�ނł͂Ȃ��ꍇ
                                            dblWK = .intUCutQF1 / 10.0
                                            ret = CheckDoubleData(cTextBox, dblWK)
                                            .intUCutQF1 = Convert.ToInt16(dblWK * 10.0) ' (KHz��0.1KHz)
                                        End If
                                    Case 5 ' ���x
                                        ret = CheckDoubleData(cTextBox, .dblUCutV1)
                                    Case 6 ' �^�[���|�C���g
                                        ret = CheckDoubleData(cTextBox, .dblUCutTurnP)

                                    Case Else
                                        Throw New Exception("Case " & tag & ": Nothing")
                                End Select

                            End If

                            'V2.2.0.0�A��
                        Case Else
                            Throw New Exception("Parent.Tag - Case " & tag & ": Nothing")
                    End Select
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex, cTextBox.Name)
                ret = 1
            Finally
                CheckTextData = ret
            End Try

        End Function
#End Region

        '        'V2.2.1.7�@ ��
        '#Region "�e�L�X�g�{�b�N�X�̕����񂪐����ϊ��ł��邩�m�F"
        '        ''' <summary>�e�L�X�g�{�b�N�X�̕����񂪐����ϊ��ł��邩�m�F�i�e�L�X�g�{�b�N�X�p�j</summary>
        '        ''' <param name="cTextBox">�m�F����÷���ޯ��</param>
        '        ''' <returns>(-1)=�װ</returns>
        '        Private Function CheckNumeric(ByRef cTextBox As cTxt_) As Integer
        '            Dim ret As Integer = 0
        '            Try

        '                '���l�`�F�b�N
        '                If IsNumeric(cTextBox.Text) Then
        '                    'Nop
        '                Else
        '                    MsgBox("���l����͂��Ă��������B")
        '                    ret = -1
        '                End If
        '            Catch ex As Exception
        '                Call MsgBox_Exception(ex, cTextBox.Name)
        '                ret = -1
        '            Finally
        '                CheckNumeric = ret
        '            End Try

        '        End Function
        '#End Region
        '        'V2.2.1.7�@ ��

#Region "���ޯ�����÷���ޯ���̶�Ĕԍ���Ԃ�"
        ''' <summary>m_CtlIdxCut(,)�ł�1�����ڂ̲��ޯ����Ԃ�(÷���ޯ���p)</summary>
        ''' <param name="cTextBox">�m�F����÷���ޯ��</param>
        ''' <param name="tag">÷���ޯ�������</param>
        ''' <returns>(-1)=�װ, 0~4=���ޯ��</returns>
        Private Function GetCtlIdx(ByRef cTextBox As cTxt_, ByVal tag As Integer) As Integer
            Dim ret As Integer = (-1)
            Try
                For i As Integer = 0 To (m_CtlIdxCut.GetLength(0) - 1) Step 1
                    If (m_CtlIdxCut(i, tag) Is cTextBox) Then
                        ret = i
                        Exit For
                    End If
                Next i
            Catch ex As Exception
                Call MsgBox_Exception(ex, cTextBox.Name)
                ret = (-1)
            Finally
                GetCtlIdx = ret
            End Try

        End Function

        ''' <summary>m_CtlIdxCut(,)�ł�1�����ڂ̲��ޯ����Ԃ�(�����ޯ���p)</summary>
        ''' <param name="cCombo">�m�F��������ޯ��</param>
        ''' <param name="tag">÷���ޯ�������</param>
        ''' <returns>(-1)=�װ, 0~4=���ޯ��</returns>
        Private Function GetCtlIdx(ByRef cCombo As cCmb_, ByVal tag As Integer) As Integer
            Dim ret As Integer = (-1)
            Try
                For i As Integer = 0 To (m_CtlIdxCut.GetLength(0) - 1) Step 1
                    If (m_CtlIdxCut(i, IDX_MTYPE + tag) Is cCombo) Then
                        ret = i
                        Exit For
                    End If
                Next i

            Catch ex As Exception
                Call MsgBox_Exception(ex, cCombo.Name)
                ret = (-1)
            Finally
                GetCtlIdx = ret
            End Try

        End Function
#End Region

#Region "�k�J�b�g�e�L�X�g�{�b�N�X�̃J�b�g�ԍ���Ԃ�"
        ''' <summary>
        ''' m_CtlLCut(,)�ł�1�����ڂ̃C���f�b�N�X��Ԃ�(�e�L�X�g�{�b�N�X�p)
        ''' </summary>
        ''' <param name="cTextBox">�m�F����e�L�X�g�{�b�N�X</param>
        ''' <param name="tag">�e�L�X�g�{�b�N�X�̃^�O</param>
        ''' <returns>(-1)=�G���[, 0~6=�C���f�b�N�X</returns>
        ''' <remarks></remarks>
        Private Function GetCtlLCutIdx(ByRef cTextBox As cTxt_, ByVal tag As Integer) As Integer
            Dim ret As Integer = (-1)
            Try
                If tag >= LCUT_DIR_IDX Then
                    tag = tag + 1
                End If
                For i As Integer = 0 To (m_CtlLCut.GetLength(0) - 1) Step 1
                    If (m_CtlLCut(i, tag) Is cTextBox) Then
                        ret = i
                        Exit For
                    End If
                Next i
            Catch ex As Exception
                Call MsgBox_Exception(ex, cTextBox.Name)
                ret = (-1)
            Finally
                GetCtlLCutIdx = ret
            End Try

        End Function

        ''' <summary>GetCtlLCutIdx(,)�ł�1�����ڂ̲��ޯ����Ԃ�(�����ޯ���p)</summary>
        ''' <param name="cCombo">�m�F��������ޯ��</param>
        ''' <param name="tag">÷���ޯ�������</param>
        ''' <returns>(-1)=�װ, 0~4=���ޯ��</returns>
        Private Function GetCtlLCutIdx(ByRef cCombo As cCmb_, ByVal tag As Integer) As Integer
            Dim ret As Integer = (-1)
            Try
                For i As Integer = 0 To (m_CtlLCut.GetLength(0) - 1) Step 1
                    If (m_CtlLCut(i, LCUT_DIR_IDX + tag) Is cCombo) Then
                        ret = i
                        Exit For
                    End If
                Next i

            Catch ex As Exception
                Call MsgBox_Exception(ex, cCombo.Name)
                ret = (-1)
            Finally
                GetCtlLCutIdx = ret
            End Try

        End Function
#End Region

        'V2.0.0.0�F��
#Region "���g���[�X�J�b�g�e�L�X�g�{�b�N�X�̃J�b�g�ԍ���Ԃ�"
        ''' <summary>
        ''' m_CtlRetraceCut(,)�ł�1�����ڂ̃C���f�b�N�X��Ԃ�(�e�L�X�g�{�b�N�X�p)
        ''' </summary>
        ''' <param name="cTextBox">�m�F����e�L�X�g�{�b�N�X</param>
        ''' <param name="tag">�e�L�X�g�{�b�N�X�̃^�O</param>
        ''' <returns>(-1)=�G���[, 0~9=�C���f�b�N�X</returns>
        ''' <remarks></remarks>
        Private Function GetCtlRetraceCutIdx(ByRef cTextBox As cTxt_, ByVal tag As Integer) As Integer
            Dim ret As Integer = (-1)
            Try
                For i As Integer = 0 To (m_CtlRetraceCut.GetLength(0) - 1) Step 1
                    If (m_CtlRetraceCut(i, tag) Is cTextBox) Then
                        ret = i
                        Exit For
                    End If
                Next i
            Catch ex As Exception
                Call MsgBox_Exception(ex, cTextBox.Name)
                ret = (-1)
            Finally
                GetCtlRetraceCutIdx = ret
            End Try

        End Function
#End Region
        'V2.0.0.0�F��

#Region "��������"
        ''' <summary>������������</summary>
        ''' <returns>0 = ����, 1 = �װ</returns>
        Protected Overrides Function CheckRelation() As Integer
            Dim strMsg As String
            Dim errIdx As Integer
            Dim ctlArray() As Control

            CheckRelation = 0 ' Return�l = ����
            Try
                With m_MainEdit
                    ctlArray = m_CtlCut ' ���(����)��ٰ���ޯ��
                    '---------------------------------------------------------------------------
                    '   �O���@��̎w�肪�Ȃ��ꍇ�̊O������w��`�F�b�N
                    '---------------------------------------------------------------------------
                    ' ������(0=��������, 1�ȏ�=�O������)
                    If (1 <= .W_REG(m_ResNo).STCUT(m_CutNo).intMType) And (.W_PLT.GCount <= 0) Then
                        strMsg = "���փ`�F�b�N�G���[" & vbCrLf
                        strMsg = strMsg & "�O���@��̎w�肪�Ȃ��ꍇ�͊O�������̎w��͂ł��܂���B"
                        errIdx = CUT_MTYPE
                        GoTo STP_ERR
                    End If

                    '---------------------------------------------------------------------------
                    '   ��ĕ���1,2�`�F�b�N(�k�J�b�g/�������ݶ�Ď�)
                    '---------------------------------------------------------------------------
                    'V1.0.4.3�F�J�b�g�����́A�V�J�b�g�ʂɂO����R�T�X�x�͈̔͂Ŏw��\�B
                    'V1.0.4.3�F                    If (.W_REG(m_ResNo).STCUT(m_CutNo).intCTYP = CNS_CUTP_L) Or (.W_REG(m_ResNo).STCUT(m_CutNo).intCTYP = CNS_CUTP_SP) Then
                    'V1.0.4.3�F                        Select Case (.W_REG(m_ResNo).STCUT(m_CutNo).intANG) ' ��ĕ���1(90���P�ʁ@0���`360��)
                    'V1.0.4.3�F                            Case 0, 180 ' ��ĕ���1 = 0,180���Ȃ� ��ĕ���2 = 90,270���ȊO�G���[
                    'V1.0.4.3�F                        If (.W_REG(m_ResNo).STCUT(m_CutNo).intANG2 = 0) Or (.W_REG(m_ResNo).STCUT(m_CutNo).intANG2 = 180) Then
                    'V1.0.4.3�F                        strMsg = "���փ`�F�b�N�G���[" & vbCrLf
                    'V1.0.4.3�F                        strMsg = strMsg & "�J�b�g�����P�ƃJ�b�g�����Q�̑g���킹�w�肪����������܂���B"
                    'V1.0.4.3�F                        GoTo STP_ERR
                    'V1.0.4.3�F                    End If
                    'V1.0.4.3�F                            Case 90, 270 ' ��ĕ���1 = 90,270���Ȃ� ��ĕ���2 = 0,180���ȊO�G���[
                    'V1.0.4.3�F                        If (.W_REG(m_ResNo).STCUT(m_CutNo).intANG2 = 90) Or (.W_REG(m_ResNo).STCUT(m_CutNo).intANG2 = 270) Then
                    'V1.0.4.3�F                        strMsg = "���փ`�F�b�N�G���[" & vbCrLf
                    'V1.0.4.3�F                        strMsg = strMsg & "�J�b�g�����P�ƃJ�b�g�����Q�̑g���킹�w�肪����������܂���B"
                    'V1.0.4.3�F                        GoTo STP_ERR
                    'V1.0.4.3�F                    End If
                    'V1.0.4.3�F                        End Select
                    'V1.0.4.3�F                    End If
                    '###1042�@��
                    '---------------------------------------------------------------------------
                    '   �����}�[�L���O�̎��̊p�x�`�F�b�N
                    '---------------------------------------------------------------------------
                    If .W_REG(m_ResNo).STCUT(m_CutNo).intCTYP = CNS_CUTP_M Then
                        If .W_REG(m_ResNo).STCUT(m_CutNo).intANG Mod 90 > 0 Then
                            strMsg = "���փ`�F�b�N�G���[" & vbCrLf
                            strMsg = strMsg & "�����̊p�x�́A0��,90��,180��,270���̂ݗL���ł��B"
                            '                            errIdx = CUT_DIR_1
                            Call MsgBox(strMsg, DirectCast(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                            CheckRelation = 1 ' Return�l = �װ
                            Exit Function
                        End If
                        If .W_REG(m_ResNo).STCUT(m_CutNo).dblDL2 < 0 Or 10.0 < .W_REG(m_ResNo).STCUT(m_CutNo).dblDL2 Then
                            strMsg = "���փ`�F�b�N�G���[" & vbCrLf
                            strMsg = strMsg & "�����̍����́A0.1mm�`10.0mm�͈̔͂Ŏw�肵�ĉ������B"
                            errIdx = CUT_LEN_1
                            GoTo STP_ERR
                        End If
                    End If
                    '###1042�@��
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

            Exit Function
STP_ERR:
            Call MsgBox_CheckErr(DirectCast(ctlArray(errIdx), cTxt_), strMsg)
            CheckRelation = 1 ' Return�l = �װ

        End Function
#End Region

#Region "�ǉ���폜���݊֘A����"
        ''' <summary>����ް���ǉ��܂��͍폜���A�����ް�������������</summary>
        ''' <param name="addDel">1=Add, (-1)=Del</param>
        Private Sub SortCutData(ByVal addDel As Integer)
            Dim iStart As Integer
            Dim iEnd As Integer
            Dim dir As Integer = (-1) * addDel ' Add=(-1), Del=1�ɂ���
            Try
                With m_MainEdit.W_REG(m_ResNo)
                    If (1 = addDel) Then ' �ǉ��̏ꍇ
                        .intTNN = Convert.ToInt16(.intTNN + 1) ' �o�^��Đ���ǉ�����
                        iStart = .intTNN ' �o�^����Ă��鶯Đ�����
                        iEnd = (m_CutNo + 1) ' �ǉ����鶯��ް��ԍ�+1�܂ŁA�O���ް������ɂ��炷
                    Else ' �폜�̏ꍇ
                        iStart = m_CutNo ' �폜���鶯��ް��ԍ�����
                        iEnd = (.intTNN - 1) ' �o�^����Ă��鶯��ް���-1�܂ŁA�����ް���O�ɂ��炷
                    End If

                    For cn As Integer = iStart To iEnd Step dir
                        .STCUT(cn).intCUT = .STCUT(cn + dir).intCUT     ' ��ĕ��@(1:�ׯ�ݸ�, 2:���ޯ��, 3:�߼޼��ݸޖ������ޯ��)
                        .STCUT(cn).intCTYP = .STCUT(cn + dir).intCTYP   ' ��Č`��(1:��ڰ�, 2:L���)
                        .STCUT(cn).intNum = .STCUT(cn + dir).intNum     ' ��Ė{��(�������ݶ�Ă̂�)
                        .STCUT(cn).dblSTX = .STCUT(cn + dir).dblSTX     ' ���ݸ� ���ē_ X
                        .STCUT(cn).dblSTY = .STCUT(cn + dir).dblSTY     ' ���ݸ� ���ē_ Y
                        .STCUT(cn).dblSX2 = .STCUT(cn + dir).dblSX2     ' ���ݸ� ���ē_2 X
                        .STCUT(cn).dblSY2 = .STCUT(cn + dir).dblSY2     ' ���ݸ� ���ē_2 Y
                        .STCUT(cn).dblCOF = .STCUT(cn + dir).dblCOF     ' ��ĵ�(%)
                        .STCUT(cn).intTMM = .STCUT(cn + dir).intTMM     ' Ӱ��(0:����(����ڰ���ϕ�Ӱ��), 1:�����x(�ϕ�Ӱ��))
                        .STCUT(cn).intMType = .STCUT(cn + dir).intMType ' �����^�O�������
                        .STCUT(cn).intQF1 = .STCUT(cn + dir).intQF1     ' Qڰ�(0.1KHz)
                        .STCUT(cn).dblV1 = .STCUT(cn + dir).dblV1       ' ��ё��x(mm/s)
                        .STCUT(cn).intQF2 = .STCUT(cn + dir).intQF2     ' V1.0.4.3�B�X�g���[�g�J�b�g�E���g���[�X��Q���[�g(0.1KHz)�Ɏg�p
                        .STCUT(cn).dblV2 = .STCUT(cn + dir).dblV2       ' V1.0.4.3�B�X�g���[�g�J�b�g�E���g���[�X�̃g�������x(mm/s)�Ɏg�p
                        .STCUT(cn).dblDL2 = .STCUT(cn + dir).dblDL2     ' ��2�̶�Ē�(�ЯĶ�ė�mm(L��ݑO))
                        .STCUT(cn).dblDL3 = .STCUT(cn + dir).dblDL3     ' ��3�̶�Ē�(�ЯĶ�ė�mm(L��݌�))
                        .STCUT(cn).intANG = .STCUT(cn + dir).intANG     ' ��ĕ���1
                        .STCUT(cn).intANG2 = .STCUT(cn + dir).intANG2   ' ��ĕ���2
                        .STCUT(cn).dblLTP = .STCUT(cn + dir).dblLTP     ' L��� �߲��(%)
                        .STCUT(cn).cFormat = .STCUT(cn + dir).cFormat   '###1042�@ �����f�[�^
                        .STCUT(cn).cMarkFix = .STCUT(cn + dir).cMarkFix   '�󎚌Œ蕔 'V2.2.1.7�@
                        .STCUT(cn).cMarkStartNum = .STCUT(cn + dir).cMarkStartNum   '�J�n�ԍ� 'V2.2.1.7�@
                        .STCUT(cn).intMarkRepeatCnt = .STCUT(cn + dir).intMarkRepeatCnt   '�J�n�ԍ� 'V2.2.1.7�@

                        'V2.1.0.0�@�� �J�b�g���̒�R�l�ω��ʔ���@�\�ǉ�
                        .STCUT(cn).iVariationRepeat = .STCUT(cn + dir).iVariationRepeat     ' ���s�[�g�L��
                        .STCUT(cn).iVariation = .STCUT(cn + dir).iVariation                 ' ����L��
                        .STCUT(cn).dRateOfUp = .STCUT(cn + dir).dRateOfUp                   ' �㏸��
                        .STCUT(cn).dVariationLow = .STCUT(cn + dir).dVariationLow           ' �����l
                        .STCUT(cn).dVariationHi = .STCUT(cn + dir).dVariationHi             ' ����l
                        'V2.1.0.0�@��

                        ' ���ޯ����ď��ݒ�
                        For ix As Integer = 1 To MAXIDX Step 1 ' MAX���ޯ����Đ����J�Ԃ�
                            .STCUT(cn).intIXN(ix) = .STCUT(cn + dir).intIXN(ix) ' ��ĉ�1-5
                            .STCUT(cn).dblDL1(ix) = .STCUT(cn + dir).dblDL1(ix) ' ��Ē�1-5
                            .STCUT(cn).lngPAU(ix) = .STCUT(cn + dir).lngPAU(ix) ' �߯����߰��1-5
                            .STCUT(cn).dblDEV(ix) = .STCUT(cn + dir).dblDEV(ix) ' �덷1-5(%)
                            .STCUT(cn).intIXMType(ix) = .STCUT(cn + dir).intIXMType(ix) ' ����@��
                            .STCUT(cn).intIXTMM(ix) = .STCUT(cn + dir).intIXTMM(ix)     ' ����Ӱ��
                        Next ix

                        ' FL���H����
                        For fl As Integer = 1 To MAXCND Step 1
                            .STCUT(cn).intCND(fl) = .STCUT(cn + dir).intCND(fl) ' FL�ݒ�No.
                        Next fl

                        'V1.0.4.3�B ADD ��
                        Dim i As Integer
                        For i = 1 To MAX_LCUT
                            .STCUT(cn).dCutLen(i) = .STCUT(cn + dir).dCutLen(i)         ' �J�b�g���P�`�V�@���^�[�������g�p
                            .STCUT(cn).dQRate(i) = .STCUT(cn + dir).dQRate(i)           ' �p���[�g�P�`�V�@���^�[�������g�p
                            .STCUT(cn).dSpeed(i) = .STCUT(cn + dir).dSpeed(i)           ' ���x�P�`�V
                            .STCUT(cn).dAngle(i) = .STCUT(cn + dir).dAngle(i)           ' �p�x�P�`�V
                            .STCUT(cn).dTurnPoint(i) = .STCUT(cn + dir).dTurnPoint(i)   ' �^�[���|�C���g�P�`�U
                        Next
                        'V1.0.4.3�B ADD ��

                        'V2.0.0.0�F ADD ��
                        .STCUT(cn).intRetraceCnt = .STCUT(cn + dir).intRetraceCnt       ' ���g���[�X�J�b�g�{��
                        For i = 1 To MAX_LCUT
                            .STCUT(cn).dblRetraceOffX(i) = .STCUT(cn + dir).dblRetraceOffX(i)       ' ���g���[�X�̃I�t�Z�b�g�w
                            .STCUT(cn).dblRetraceOffY(i) = .STCUT(cn + dir).dblRetraceOffY(i)       ' ���g���[�X�̃I�t�Z�b�g�x
                            .STCUT(cn).dblRetraceQrate(i) = .STCUT(cn + dir).dblRetraceQrate(i)     ' �X�g���[�g�J�b�g�E���g���[�X��Q���[�g(0.1KHz)�Ɏg�p
                            .STCUT(cn).dblRetraceSpeed(i) = .STCUT(cn + dir).dblRetraceSpeed(i)     ' �X�g���[�g�J�b�g�E���g���[�X�̃g�������x(mm/s)�Ɏg�p
                        Next
                        'V2.0.0.0�F ADD ��

                        'V2.2.0.0�A��
                        'U�J�b�g�p�����[�^�̒ǉ�
                        .STCUT(cn).dUCutL1 = 0.0          ' L1
                        .STCUT(cn).dUCutL2 = 0.0          ' L2
                        .STCUT(cn).intUCutQF1 = 0.1       ' Q���[�g
                        .STCUT(cn).dblUCutV1 = 0.1        ' ���x
                        .STCUT(cn).intUCutANG = 0         ' �p�x
                        .STCUT(cn).dblUCutTurnP = 0       ' �^�[���|�C���g
                        .STCUT(cn).intUCutTurnDir = 1     ' �^�[������
                        .STCUT(cn).dblUCutR1 = 0          ' R1
                        .STCUT(cn).dblUCutR2 = 0          ' R2
                        'V2.2.0.0�A��

                    Next cn

                    ' �߂ĕs�v�ƂȂ����ް�������������
                    If (1 = addDel) Then ' �ǉ��̏ꍇ
                        Call InitCutData(m_ResNo, m_CutNo) ' �ǉ������ް���������
                    Else ' �폜�̏ꍇ
                        Call InitCutData(m_ResNo, .intTNN) ' �Ō���ް���������
                        .intTNN = Convert.ToInt16(.intTNN - 1) ' �o�^��Đ���-1����

                        ' �ŏI��Ă̍폜�Ȃ猻�̶݂�Ĕԍ����ŏI��Ĕԍ��Ƃ���
                        If (.intTNN < m_CutNo) Then m_CutNo = .intTNN
                    End If
                End With

                ' ����ް�����ʍ��ڂɐݒ�
                Call SetDataToText()
                FIRST_CONTROL.Select() ' ̵����ݒ�

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#Region "�����"
        ''' <summary>�����ޯ���̲��ޯ�����ύX���ꂽ���̏���</summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Protected Overrides Sub cCmb_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
            Dim cCombo As cCmb_
            Dim tag As Integer
            Dim idx As Integer
            Dim Cnt As Integer      'V1.0.4.3�A
            Dim ctlIdx As Integer   'V1.0.4.3�B
            Try
                cCombo = DirectCast(sender, cCmb_)
                tag = DirectCast(cCombo.Tag, Integer)
                idx = cCombo.SelectedIndex
                With m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo)
                    Select Case (DirectCast(cCombo.Parent.Tag, Integer)) ' ��ٰ���ޯ�������
                        ' ------------------------------------------------------------------------------
                        Case 0 ' ��ĸ�ٰ���ޯ��
                            Select Case (tag)
                                Case 0 ' ��R�ԍ�
                                    m_ResNo = (idx + 1) ' ���̑����m_CutNo��1�ɂȂ�
                                    ' �Ή������ް���÷���ޯ���A��Ĕԍ������ޯ���ɾ�Ă���
                                    Call SetDataToText() ' ��Ĕԍ���1��ݒ肷��
                                Case 1 ' ��Ĕԍ�
                                    m_CutNo = (idx + 1)
                                    ' �Ή������ް���÷���ޯ���A��Ĕԍ������ޯ���ɾ�Ă���
                                    Call SetDataToText()

                                Case 2 ' ��ĕ��@(1:�ׯ�ݸ�, 2:���ޯ��, 3:NG���)
                                    .intCUT = Convert.ToInt16(idx + 1)
                                    Call ChangedCutMethod(idx + 1) ' �֘A���۰ق̗L���������ݒ�

                                    ' NG��Ă܂��͑���@�킪�O�������̏ꍇ����Ӱ�ނ𖳌��ɂ���
                                    If (CNS_CUTM_NG = .intCUT) OrElse (0 < .intMType) Then
                                        m_CtlCut(CUT_TMM).Enabled = False ' ����Ӱ�ޖ���
                                    Else
                                        m_CtlCut(CUT_TMM).Enabled = True ' ����Ӱ�ޗL��
                                    End If

                                    If (CNS_CUTM_TR = .intCUT Or CNS_CUTM_NG = .intCUT) Then       ' �g���b�L���O�܂���NG�J�b�g�̏ꍇ�́A��������̂�
                                        .intMType = 0
                                        m_CtlCut(CUT_MTYPE).Enabled = False
                                        Call SetCutData()
                                    Else
                                        m_CtlCut(CUT_MTYPE).Enabled = True
                                    End If

                                    'V2.0.0.5�@                                    'V2.0.0.0�N �C���f�b�N�X�Ńg���b�L���O��L�J�b�g����
                                    'V2.0.0.5�@                                    If m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCUT = CNS_CUTM_IX And m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCTYP = CNS_CUTP_L Then
                                    'V2.0.0.5�@                                        CGrp_1.Enabled = False                      ' �C���f�b�N�X
                                    'V2.0.0.5�@                                        CGrp_3.Enabled = False                       '�k�J�b�g�p�����[�^
                                    'V2.0.0.5�@                                    End If
                                    'V2.0.0.5�@                                    'V2.0.0.0�N��

                                Case 3 ' ��Č`��(1:��ڰ�, 2:L���, 3:��������)
                                    .intCTYP = GetComboBoxName2Value(cCombo.Text, Me.m_lstCutType)
                                    Call ChangedCutShape(.intCTYP) ' �֘A���۰ق̕\�����\����ݒ�
                                    'V2.0.0.5�@                                    'V2.0.0.0�N���C���f�b�N�X�Ńg���b�L���O��L�J�b�g����
                                    'V2.0.0.5�@                                    If m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCUT = CNS_CUTM_IX And m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCTYP = CNS_CUTP_L Then
                                    'V2.0.0.5�@                                        CGrp_1.Enabled = False                      ' �C���f�b�N�X
                                    'V2.0.0.5�@                                        CGrp_3.Enabled = False                       '�k�J�b�g�p�����[�^
                                    'V2.0.0.5�@                                    End If
                                    'V2.0.0.5�@                                    'V2.0.0.0�N��
                                Case 4, 5 ' ��ĕ���1, ��ĕ���2
                                    Dim iWK As Short = 0
                                    'V1.0.4.3�A �R�����g����
                                    'Select Case (idx) ' ��ĕ���(90���P�ʁ@0���`360��)
                                    '    Case 0 : iWK = 0    ' 0��
                                    '    Case 1 : iWK = 90   ' 90��
                                    '    Case 2 : iWK = 180  ' 180��
                                    '    Case 3 : iWK = 270  ' 270��
                                    '    Case 4 : iWK = 10   ' 10��
                                    '    Case 5 : iWK = 20   ' 20��
                                    '    Case 6 : iWK = 30   ' 30��
                                    '    Case 7 : iWK = 40   ' 40��
                                    '    Case 8 : iWK = 50   ' 50��
                                    '    Case 9 : iWK = 60   ' 60��
                                    '    Case 10 : iWK = 70  ' 70��
                                    '    Case 11 : iWK = 80  ' 80��
                                    '    Case 12 : iWK = 100 ' 100��
                                    '    Case 13 : iWK = 110 ' 110��
                                    '    Case 14 : iWK = 120 ' 120��
                                    '    Case 15 : iWK = 130 ' 130��
                                    '    Case 16 : iWK = 140 ' 140��
                                    '    Case 17 : iWK = 150 ' 150��
                                    '    Case 18 : iWK = 160 ' 160��
                                    '    Case 19 : iWK = 170 ' 170��
                                    '    Case Else ' DO NOTHING
                                    'End Select
                                    'V1.0.4.3�A��
                                    For Cnt = 0 To MAX_DEGREES
                                        If AngleArray(Cnt, 1) = idx Then
                                            iWK = AngleArray(Cnt, 0)
                                            Exit For
                                        End If
                                    Next
                                    'V1.0.4.3�A��
                                    ' �ҏW�ް��̶�ĕ�����ݒ肷��
                                    If (4 = tag) Then ' ��ĕ���1(90���P�ʁ@0���`360��)
                                        .intANG = iWK
                                    Else ' ��ĕ���2(90���P�ʁ@0���`360��)
                                        .intANG2 = iWK
                                    End If

                                Case 6 ' ����@��(0:���������, 1�ȏ�͊O�������ԍ�)
                                    ' �o�^����Ă��鑪��@��ؽĂ̐��l��ݒ肷��( 1:NAME=1, 10:NAME=10)
                                    .intMType = Convert.ToInt16((cCombo.Text).Substring(0, 2))

                                    ' ����@�킪�O�������̏ꍇ����Ӱ�ނ𖳌��ɂ���
                                    If (0 < idx) Then
                                        m_CtlCut(CUT_TMM).Enabled = False ' ����Ӱ�ޖ���
                                    Else
                                        m_CtlCut(CUT_TMM).Enabled = True ' ����Ӱ�ޗL��
                                    End If

                                Case 7 ' ����Ӱ��
                                    .intTMM = Convert.ToInt16(idx)
                                    'V2.1.0.0�@��
                                Case 8 ' �J�b�g���̒�R�l�ω��ʔ���@�\���s�[�g�L��
                                    Dim bChange As Boolean = False
                                    If .iVariationRepeat <> Convert.ToInt16(idx) Then
                                        ' �u����v����u�Ȃ��v�ɕς���������Ȃ��S�Ă���ׂɃR�s�[����B
                                        bChange = True
                                    End If
                                    .iVariationRepeat = Convert.ToInt16(idx)
                                    If .iVariationRepeat = 1 OrElse bChange Then
                                        UserSub.CutVariationDataCopy(m_MainEdit.W_PLT, m_MainEdit.W_REG, m_ResNo, m_CutNo)
                                    End If
                                Case 9 ' �J�b�g���̒�R�l�ω��ʔ���@�\����L��
                                    .iVariation = Convert.ToInt16(idx)
                                    If .iVariationRepeat = 1 Then
                                        UserSub.CutVariationDataCopy(m_MainEdit.W_PLT, m_MainEdit.W_REG, m_ResNo, m_CutNo)
                                    End If
                                    'V2.1.0.0�@��
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case 1 ' ���ޯ����ĸ�ٰ���ޯ��
                            ctlIdx = GetCtlIdx(cCombo, tag) ' 1�����ڂ̲��ޯ��

                            Select Case (tag)
                                Case 0 ' ����@��(0:���������, 1�ȏ�͊O�������ԍ�)
                                    ' �o�^����Ă��鑪��@��ؽĂ̐��l��ݒ肷��( 1:NAME=1, 10:NAME=10)
                                    .intIXMType(ctlIdx + 1) = Convert.ToInt16((cCombo.Text).Substring(0, 2))
                                    ' �O�������̏ꍇ����Ӱ�ނ𖳌��ɂ���
                                    If (0 < idx) Then
                                        m_CtlIdxCut(ctlIdx, IDX_TMM).Enabled = False
                                    Else
                                        m_CtlIdxCut(ctlIdx, IDX_TMM).Enabled = True
                                    End If

                                Case 1 ' ����Ӱ��(0:����, 1:�����x)
                                    .intIXTMM(ctlIdx + 1) = Convert.ToInt16(idx)
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case 2 ' FL���H������ٰ���ޯ��
                            Select Case (tag)
                                Case 0 ' FL�ݒ�No.(0�`31)
                                    Dim cndNo As Integer ' ��ď����ԍ�
                                    For i As Integer = 0 To (m_CtlFLCnd.GetLength(0) - 1) Step 1
                                        If (m_CtlFLCnd(i, 0) Is cCombo) Then
                                            cndNo = (i + 1)
                                            Exit For
                                        End If
                                    Next i
                                    .intCND(cndNo) = Convert.ToInt16(idx)

                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                            'V1.0.4.3�B��
                            ' ------------------------------------------------------------------------------
                        Case 3 ' �k�J�b�g�p�����[�^�O���[�v�{�b�N�X
                            ctlIdx = GetCtlLCutIdx(cCombo, tag) ' 1�����ڂ̲��ޯ��
                            Select Case (tag)
                                Case 0 ' �J�b�g��
                                    For Cnt = 0 To MAX_DEGREES
                                        If AngleArrayForLcut(Cnt, 1) = idx Then
                                            .dAngle(ctlIdx + 1) = AngleArrayForLcut(Cnt, 0)
                                            Exit For
                                        End If
                                    Next
                                    If ctlIdx = 0 Then
                                        .intANG = .dAngle(ctlIdx + 1)
                                    End If
                                    If .dAngle(2) >= .dAngle(1) Then
                                        If .dAngle(2) - .dAngle(1) > 180.0 Then
                                            .intANG2 = DEF_DIR_CW
                                        Else
                                            .intANG2 = DEF_DIR_CCW
                                        End If
                                    Else
                                        If .dAngle(1) - .dAngle(2) > 180.0 Then
                                            .intANG2 = DEF_DIR_CCW
                                        Else
                                            .intANG2 = DEF_DIR_CW
                                        End If
                                    End If
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                            'V1.0.4.3�B��
                            'V2.2.0.0�A��
                        Case 5 ' �t�J�b�g�p�����[�^�O���[�v�{�b�N�X

                            ctlIdx = GetCtlLCutIdx(cCombo, tag) ' 1�����ڂ̲��ޯ��
                            Select Case (tag)
                                Case 0 ' �J�b�g�p�x
                                    For Cnt = 0 To MAX_DEGREES
                                        If AngleArrayForUcut(Cnt, 1) = idx Then
                                            .intUCutANG = AngleArrayForUcut(Cnt, 0)
                                            Exit For
                                        End If
                                    Next

                                Case 1 ' �^�[������
                                    .intUCutTurnDir = Convert.ToInt16(idx) + 1
                            End Select
                            'V2.2.0.0�A��

                        Case Else
                            Throw New Exception("Parent.Tag - Case Else")

                    End Select
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

        ''' <summary>�ǉ����ݸد����̏���</summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Private Sub CBtn_Add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBtn_Add.Click
            Dim refOpt As Short ' ��߼������(0=�O�ɒǉ� ,1=��ɒǉ�)
            Dim ret As Integer
            Try
                ' �o�^������
                With m_MainEdit
                    If (.W_PLT.RCount < 1) Then Exit Sub ' ��R�ް��Ȃ��Ȃ�NOP
                    If (MAXCTN <= .W_REG(m_ResNo).intTNN) Then ' �J�b�g�� >= 9 ?
                        Dim strMsg As String = "����ȏ�J�b�g�f�[�^�͓o�^�ł��܂���B"
                        Call MsgBox(strMsg, DirectCast( _
                                    MsgBoxStyle.OkOnly + _
                                    MsgBoxStyle.Information, MsgBoxStyle), _
                                    My.Application.Info.Title)
                        Exit Sub
                    End If
                End With

                ' �m�Fү���ނ�\��("�J�b�g�f�[�^��ǉ����܂�")
                ret = MsgBox_AddClick("�J�b�g�f�[�^", refOpt) ' ү���ޕ\��
                If (ret <> cFRS_ERR_ADV) Then Exit Sub ' Cancel�Ȃ�Return
                If (refOpt = 1) Then ' �\���ް��̌�ɒǉ� ?
                    m_CutNo = (m_CutNo + 1) ' �ǉ������ް��̔ԍ� = ���݂��ް��ԍ� + 1
                Else ' �\���ް��̑O�ɒǉ�
                    m_CutNo = m_CutNo ' �ǉ������ް��̔ԍ� = ���݂��ް��ԍ�
                End If

                ' �ް���1��ɂ��炵�Ēǉ�����
                Call SortCutData(1)

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

        ''' <summary>�폜���ݸد����̏���</summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Private Sub CBtn_Del_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBtn_Del.Click
            Dim strMsg As String ' ү���ޕҏW��
            Dim ret As Integer
            Try
                ' �m�Fү���ނ�\��
                If (1 = m_MainEdit.W_REG(m_ResNo).intTNN) Then Exit Sub ' ��R���J�b�g��1�Ȃ�NOP
                strMsg = "���݂̃J�b�g�f�[�^���폜���܂��B��낵���ł����H"
                ret = MsgBox(strMsg, DirectCast( _
                            MsgBoxStyle.OkCancel + _
                            MsgBoxStyle.Exclamation + _
                            MsgBoxStyle.DefaultButton2, MsgBoxStyle), _
                            My.Application.Info.Title)

                If (ret = MsgBoxResult.Cancel) Then Exit Sub ' Cancel(RESET��) ?

                ' �����ް���1�O�ɂ߂�
                Call SortCutData(-1)

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

        ''' <summary>���H�������ݸد����̏���</summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Private Sub CBtn_FLS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBtn_FLC.Click
            Dim ret As Integer
            Try
#If cOSCILLATORcFLcUSE Then

                Dim fls As Process = Process.Start("C:\TRIM\FLSetup.exe") ' FLSetup.exe���N������
                fls.WaitForExit() ' �I����҂�

                ' FL�����猻�݂̉��H������Ҳ݉�ʂ̉��H�����ް��Ɏ�M����
                ret = TrimCondInfoRcv(stCND)
                If (0 <> ret) Then ' �װ�̏ꍇ
                    Dim strMsg As String = "�e�k�����H�����̃��[�h�Ɏ��s���܂����B"
                    Call MsgBox(strMsg, DirectCast( _
                                MsgBoxStyle.OkOnly + _
                                MsgBoxStyle.Critical, MsgBoxStyle), _
                                My.Application.Info.Title)
                    Exit Sub
                End If

                ' �ް�����M���AҲ݉�ʂ̉��H�����ް����X�V���ꂽ�ꍇ
                Call m_MainEdit.ReadFlConditionData() ' �ҏW��ʂ�FL���H�����ް����X�V����
                Call SetFLCndData() ' ������FL���H���������̕\�����X�V
                Me.Refresh()
#Else
                ret = 0
#End If

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

        'V2.0.0.0��
#Region "��ĕ��@�̏����ݒ�"
        ''' <summary>
        ''' ��ĕ��@�̏����ݒ�
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub InitCutMethodData()
            Dim cd As New ComboDataStruct

            cd.SetData("�g���b�L���O", CNS_CUTM_TR)
            m_lstCutMethod.Add(cd)
            cd.SetData("�C���f�b�N�X", CNS_CUTM_IX)
            m_lstCutMethod.Add(cd)
#If cFORCEcCUT Then
            cd.SetData("�����J�b�g", CNS_CUTM_FC)
            m_lstCutMethod.Add(cd)
#End If
        End Sub
#End Region

#Region "�J�b�g�`��̏����ݒ�"
        ''' <summary>
        ''' �J�b�g�`��̏����ݒ�
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub InitCutTypeData()
            Try
                Dim ctyp As New ComboDataStruct

                ctyp.SetData("�X�g���[�g", CNS_CUTP_ST)
                m_lstCutType.Add(ctyp)
                ctyp.SetData("���g���[�X", CNS_CUTP_ST_TR)
                m_lstCutType.Add(ctyp)
                ctyp.SetData("�k�J�b�g", CNS_CUTP_L)
                m_lstCutType.Add(ctyp)
                ctyp.SetData("����", CNS_CUTP_M)
                m_lstCutType.Add(ctyp)
                ctyp.SetData("�t�J�b�g", CNS_CUTP_U)        'V2.2.0.0�A 
                m_lstCutType.Add(ctyp)                      'V2.2.0.0�A 


            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try
        End Sub
#End Region

#Region "��R�̔��胂�[�h���ύX���ꂽ���ɃC���f�b�N�X�̌덷�̐ݒ��ύX����"
        ''' <summary>��R�̔��胂�[�h���ύX���ꂽ���ɃC���f�b�N�X�̌덷�̐ݒ��ύX����</summary>
        ''' <param name="nJudge">���胂�[�h(0:�䗦(%), 1:���l(��Βl))</param>
        ''' <param name="ctlText">�e�L�X�g�R���g���[��</param>
        ''' <param name="strUnit">�P��</param>
        Private Sub ChangedJudge(ByVal nJudge As Integer, ByVal ctlText As Control, ByVal strUnit As String)
            Dim strMin As String
            Dim strMax As String

            Try
                If nJudge = JUDGE_MODE_RATIO Then ' �䗦
                    strMin = m_strDEVRaite(0)
                    strMax = m_strDEVRaite(1)

                    ' �덷���x���̕ύX
                    CLbl_20.Text = String.Format("{0}�E�䗦(%)", m_strDev)
                Else
                    strMin = m_strDEVAbsolute(0)
                    strMax = m_strDEVAbsolute(1)

                    ' �덷���x���̕ύX
                    CLbl_20.Text = String.Format("{0}�E��Βl({1})", m_strDev, strUnit)
                End If

                If TypeOf ctlText Is cTxt_ Then
                    With DirectCast(ctlText, cTxt_) ' �덷÷���ޯ��
                        Call .SetMinMax(strMin, strMax) ' �����l�����l�̐ݒ�
                        Call .SetStrTip(strMin & "�`" & strMax & "�͈̔͂Ŏw�肵�ĉ�����") ' °�����ү���ނ̐ݒ�
                    End With
                End If

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region
        'V2.0.0.0��

        ''' <summary>
        ''' �p�x�̃R���{�{�b�N�X�ɐV���ɐݒ肳�ꂽ�p�x��ǉ�����B
        ''' </summary>
        ''' <param name="IDegree"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Add_CCmb_4_Item(ByRef IDegree As Integer) As Integer
            Dim Cnt As Integer

            Try
                If AngleArray(IDegree, 1) = -1 Then
                    CCmb_4.Items.Add("   " + IDegree.ToString("0") + "��")
                    Dim iMax As Integer = 0
                    For Cnt = 0 To MAX_DEGREES
                        If AngleArray(Cnt, 1) > iMax Then
                            iMax = AngleArray(Cnt, 1)
                        End If
                    Next
                    AngleArray(IDegree, 1) = iMax + 1
                End If
                CCmb_4.SelectedIndex = AngleArray(IDegree, 1)

                Return (CCmb_4.SelectedIndex)
            Catch ex As Exception
                MsgBox("CCmb_4_KeyDown() TRAP ERROR = " + ex.Message)
            End Try

        End Function

        Private Sub CCmb_4_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CCmb_4.KeyDown
            Dim iDegree As Integer

            ' Try       V2.2.0.0�H
            If e.KeyValue <> Keys.Enter Then
                Exit Sub
            End If
            If Not IsNumeric(CCmb_4.Text) Then
                Exit Sub
            End If

            ' ��V2.2.0.0�H
            Try
                iDegree = Integer.Parse(CCmb_4.Text)
            Catch
                Call MsgBox("��������͂��Ă�������", DirectCast(
                                                MsgBoxStyle.OkOnly +
                                                MsgBoxStyle.Information, MsgBoxStyle),
                                                My.Application.Info.Title)
                Add_CCmb_4_Item(m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intANG)
                Exit Sub
            End Try

            Try
                ' ��V2.2.0.0�H
                If m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCTYP = CNS_CUTP_M And (iDegree Mod 90 > 0) Then
                    Call MsgBox("�����̊p�x�́A0��,90��,180��,270���̂ݗL���ł��B", DirectCast(
                                                MsgBoxStyle.OkOnly +
                                                MsgBoxStyle.Information, MsgBoxStyle),
                                                My.Application.Info.Title)
                    Add_CCmb_4_Item(m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intANG)
                    Exit Sub
                End If

                If iDegree < 0 Or MAX_DEGREES < iDegree Then
                    Call MsgBox("0�`359�͈̔͂Ŏw�肵�ĉ�����", DirectCast(
                                                MsgBoxStyle.OkOnly +
                                                MsgBoxStyle.Information, MsgBoxStyle),
                                                My.Application.Info.Title)
                    Add_CCmb_4_Item(m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intANG)
                    Exit Sub
                End If

                Call Add_CCmb_4_Item(iDegree)

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

        'V1.0.4.3�F��
        ''' <summary>
        ''' �p�x�̃R���{�{�b�N�X�ɐV���ɐݒ肳�ꂽ�p�x��ǉ�����B�k�J�b�g�p
        ''' </summary>
        ''' <param name="IDegree"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        ''' 
        Private Function Add_CCmb_Dir_X_Item(ByRef IDegree As Integer, ByRef cCombo As cCmb_) As Integer
            Dim Cnt As Integer

            Try
                If AngleArrayForLcut(IDegree, 1) = -1 Then
                    Dim cCombo2 As cCmb_
                    For Cnt = 0 To (m_CtlLCut.GetLength(0) - 1) Step 1
                        cCombo2 = DirectCast(m_CtlLCut(Cnt, 3), cCmb_)
                        cCombo2.Items.Add("   " + IDegree.ToString("0") + "��")
                    Next Cnt
                    Dim iMax As Integer = 0
                    For Cnt = 0 To MAX_DEGREES
                        If AngleArrayForLcut(Cnt, 1) > iMax Then
                            iMax = AngleArrayForLcut(Cnt, 1)
                        End If
                    Next
                    AngleArrayForLcut(IDegree, 1) = iMax + 1
                End If
                cCombo.SelectedIndex = AngleArrayForLcut(IDegree, 1)

                Return (cCombo.SelectedIndex)
            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Function

        ''' <summary>
        ''' �k�J�b�g�p�����[�^�̊p�x���̓C�x���g����
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub CCmb_Dir_1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CCmb_Dir_7.KeyDown, CCmb_Dir_6.KeyDown, CCmb_Dir_5.KeyDown, CCmb_Dir_4.KeyDown, CCmb_Dir_3.KeyDown, CCmb_Dir_2.KeyDown, CCmb_Dir_1.KeyDown
            Dim iDegree As Integer

            ' Try       V2.2.0.0�H
            If e.KeyValue <> Keys.Enter Then
                Exit Sub
            End If
            If Not IsNumeric(sender.Text) Then
                Exit Sub
            End If

            ' ��V2.2.0.0�H
            Try
                iDegree = Integer.Parse(sender.Text)
            Catch
                Call MsgBox("��������͂��Ă�������", DirectCast(
                                                MsgBoxStyle.OkOnly +
                                                MsgBoxStyle.Information, MsgBoxStyle),
                                                My.Application.Info.Title)
                Add_CCmb_Dir_X_Item(m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intANG, sender)
                Exit Sub
            End Try

            Try
                ' ��V2.2.0.0�H

                If iDegree < 0 Or MAX_DEGREES < iDegree Then
                    Call MsgBox("0�`359�͈̔͂Ŏw�肵�ĉ�����", DirectCast(
                                                MsgBoxStyle.OkOnly +
                                                MsgBoxStyle.Information, MsgBoxStyle),
                                                My.Application.Info.Title)
                    Add_CCmb_Dir_X_Item(m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intANG, sender)
                    Exit Sub
                End If

                Call Add_CCmb_Dir_X_Item(iDegree, sender)

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
        'V1.0.4.3�F��


        'V2.2.0.0�A��
        ''' <summary>
        '''          �t�J�b�g�p�����[�^�̒ǉ�
        ''' </summary>
        Private Sub SetUCutParamData()

            Try
                With m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo)
                    For i As Integer = 0 To (m_CtlUCut.GetLength(0) - 1) Step 1
                        Select Case (i)
                            Case 0 ' L1�J�b�g��
                                m_CtlUCut(i).Text = (.dUCutL1).ToString(DirectCast(m_CtlUCut(i), cTxt_).GetStrFormat())
                            Case 1 ' L2�J�b�g��
                                m_CtlUCut(i).Text = (.dUCutL2).ToString(DirectCast(m_CtlUCut(i), cTxt_).GetStrFormat())
                            Case 2 ' �q�P���a
                                m_CtlUCut(i).Text = (.dblUCutR1).ToString(DirectCast(m_CtlUCut(i), cTxt_).GetStrFormat())
                            Case 3 ' �q�Q���a
                                m_CtlUCut(i).Text = (.dblUCutR2).ToString(DirectCast(m_CtlUCut(i), cTxt_).GetStrFormat())
                            Case 4 ' �p���[�g
                                m_CtlUCut(i).Text = (.intUCutQF1 / 10.0).ToString(DirectCast(m_CtlUCut(i), cTxt_).GetStrFormat())
                            Case 5 ' ���x
                                m_CtlUCut(i).Text = (.dblUCutV1).ToString(DirectCast(m_CtlUCut(i), cTxt_).GetStrFormat())
                            Case 6 ' �p�x
                                Dim index As Integer
                                Select Case (.intUCutANG)
                                    Case 0
                                        index = 0   ' 0��
                                    Case 90
                                        index = 1   ' 90��
                                    Case 180
                                        index = 2   ' 180��
                                    Case 270
                                        index = 3   ' 270��
                                    Case Else
                                        index = Add_CCmb_Dir_X_Item(.dAngle(i + 1), DirectCast(m_CtlUCut(i), cCmb_))
                                End Select
                                Call NoEventIndexChange(DirectCast(m_CtlUCut(i), cCmb_), index)
                            Case 7 ' �^�[������ 
                                Dim index As Integer
                                Select Case (.intUCutTurnDir)
                                    Case 1
                                        index = 0   ' �b�v
                                    Case Else
                                        index = 1   ' �b�b�v
                                End Select
                                Call NoEventIndexChange(DirectCast(m_CtlUCut(i), cCmb_), index)

                            Case 8 ' �^�[���|�C���g
                                m_CtlUCut(i).Text = (.dblUCutTurnP.ToString(DirectCast(m_CtlUCut(i), cTxt_).GetStrFormat()))

                            Case Else
                                Throw New Exception("i = " & i & ", Case Else")
                        End Select
                    Next i
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub


    End Class
End Namespace
