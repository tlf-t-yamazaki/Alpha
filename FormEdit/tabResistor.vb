Option Explicit On
Option Strict Off

Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports LaserFront.Trimmer.DefWin32Fnc

Namespace FormEdit
    Friend Class tabResistor
        Inherits tabBase

#Region "�錾"
        '0:��R
        '1:�J�b�g��
        '2:��R��
        '3:�X���[�v
        '4:�����[�r�b�g
        '5:�ڕW�l
        '6:�\���P��
        '7:���胂�[�h
        '8:����2
        '9:IT����
        '10:IT���
        '11:FT����
        '12:FT���
        '13:����@��
        '14:���胂�[�h
        '15:�đ����
        '16:�|�[�Y����
        '17:IT��
        '18:FT��
        '19�T�[�L�b�g
        Private Const RES_NOM As Integer = 5    ' m_CtlRes�ł̲��ޯ��(�ڕW�l)
        Private Const RES_SLOPE As Integer = 3  ' m_CtlRes�ł̲��ޯ��(�X���[�v)
        Private Const RES_MTYPE As Integer = 13 ' m_CtlRes�ł̲��ޯ��(����@��)
        Private Const RES_TMM1 As Integer = 14  ' m_CtlRes�ł̲��ޯ��(����Ӱ��)
        Private Const PRB_PRH As Integer = 0    ' m_CtlProbe�ł̲��ޯ��(HI����۰��)
        Private Const PRB_PRL As Integer = 1    ' m_CtlProbe�ł̲��ޯ��(LO����۰��)

        Private m_voltNOM_Min As String         ' ���ݸ� �ڕW�l(V)
        Private m_voltNOM_Max As String
        Private m_ohmNOM_Min As String          ' ���ݸ� �ڕW�l(��)
        Private m_ohmNOM_Max As String
        Private m_ITH_Min As String             ' ��������l(ITHI)(%�p)
        Private m_ITH_Max As String
        Private m_ITL_Min As String             ' ��������l(ITLO)(%�p)
        Private m_ITL_Max As String
        Private m_FTH_Min As String             ' �I������l(FTHI)(%�p)
        Private m_FTH_Max As String
        Private m_FTL_Min As String             ' �I������l(FTLO)(%�p)
        Private m_FTL_Max As String

        Private GRP_MIN As Integer              ' ��Ĉʒu�␳��ٰ�ߔԍ��ŏ��l
        Private GRP_MAX As Integer              ' ��Ĉʒu�␳��ٰ�ߔԍ��ő�l
        Private PTN_MIN As Integer              ' ��Ĉʒu�␳����ݔԍ��ŏ��l
        Private PTN_MAX As Integer              ' ��Ĉʒu�␳����ݔԍ��ő�l

        Private m_CtlRes() As Control           ' ��R��ٰ���ޯ���̺��۰ٔz��
        Private m_CtlProbe() As Control         ' ��۰�޸�ٰ���ޯ���̺��۰ٔz��
        Private m_CtlCutCorr() As Control       ' ��Ĉʒu�␳��ٰ���ޯ���̺��۰ٔz��

        Private m_IntialFinal() As cTxt_        ' IT,FT÷���ޯ���z��
        Private m_CutPosCorr() As Control       ' ��Ĉʒu�␳�ؑւ����ɗL��������ɂ�����۰�

        'V2.0.0.0��
        ''' <summary>
        ''' �X���[�v�R���{�{�b�N�X�f�[�^���X�g
        ''' </summary>
        ''' <remarks></remarks>
        Private m_lstSlope As New List(Of ComboDataStruct)

        ''' <summary>
        ''' ���胂�[�h�R���{�{�b�N�X�f�[�^���X�g
        ''' </summary>
        ''' <remarks></remarks>
        Private m_lstMeasMode As New List(Of ComboDataStruct)
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

            ' �X���[�v�f�[�^������
            Call InitSlopeData()        'V2.0.0.0

            ' ���胂�[�h�f�[�^������
            Call InitMeasModeData()     'V2.0.0.0

            Try
                ' EDIT_DEF_User.ini������ޖ���ݒ�
                TAB_NAME = GetPrivateProfileString_S("RESISTOR_LABEL", "TAB_NAM", m_sPath, "????")

                ' ��Ĉʒu�␳��ٰ�ߔԍ������ݔԍ��̏㉺���l
                GRP_MIN = Convert.ToInt32(GetPrivateProfileString_S("PATTERN_LABEL", "GRP_MIN", m_sPath, "1"))
                GRP_MAX = Convert.ToInt32(GetPrivateProfileString_S("PATTERN_LABEL", "GRP_MAX", m_sPath, "999"))
                PTN_MIN = Convert.ToInt32(GetPrivateProfileString_S("PATTERN_LABEL", "PTN_MIN", m_sPath, "1"))
                PTN_MAX = Convert.ToInt32(GetPrivateProfileString_S("PATTERN_LABEL", "PTN_MAX", m_sPath, "50"))

                ' �ǉ���폜���݂̐ݒ�
                With mainEdit
                    CBtn_Add.SetLblToolTip(.LblToolTip)
                    CBtn_Add.Text = GetPrivateProfileString_S("BUTTON_LABEL", "BTN_ADD", m_sPath, "ADD")
                    CBtn_Del.SetLblToolTip(.LblToolTip)
                    CBtn_Del.Text = GetPrivateProfileString_S("BUTTON_LABEL", "BTN_DEL", m_sPath, "DEL")
                End With

                ' ----------------------------------------------------------
                ' EDIT_DEF_User.ini�����ٰ���ޯ���ɕ\������ݒ�
                ' ----------------------------------------------------------
                GrpArray = New cGrp_() { _
                    CGrp_0, CGrp_1, CGrp_2 _
                }
                For i As Integer = 0 To (GrpArray.Length - 1) Step 1
                    With GrpArray(i)
                        .TabIndex = ((254 - GrpArray.Length) + i) ' ���ٷ��ɂ��̫����ړ��ŕK�v
                        .Tag = i
                        .Text = GetPrivateProfileString_S( _
                                "RESISTOR_LABEL", (i.ToString("000") & "_GRP"), m_sPath, "????")
                    End With
                Next i
                ' �ǉ���폜���݂�����
                CPnl_Btn.TabIndex = 254 ' ���۰ٔz�u�\�ő吔(�Ō�ɐݒ�)

                ' ----------------------------------------------------------
                ' EDIT_DEF_User.ini�������قɕ\������ݒ�
                ' ----------------------------------------------------------
                LblArray = New cLbl_() { _
                    CLbl_0, CLbl_1, CLbl_2, CLbl_3, CLbl_4, CLbl_5, _
                    CLbl_6, CLbl_7, _
                    CLbl_8, CLbl_9, CLbl_10, CLbl_11, CLbl_12, _
                    CLbl_13, CLbl_14, CLbl_24, CLbl_25, _
                    CLbl_22, CLbl_28, CLbl_29, _
                    CLbl_26, CLbl_27, _
                    CLbl_16, CLbl_15, CLbl_17, _
                    CLbl_18, CLbl_19, CLbl_20, CLbl_21 _
                }
                For i As Integer = 0 To (LblArray.GetLength(0) - 1) Step 1
                    LblArray(i).Text = GetPrivateProfileString_S( _
                            "RESISTOR_LABEL", (i.ToString("000") & "_LBL"), m_sPath, "????")
                Next i

                ' ----------------------------------------------------------
                ' ��R��ٰ���ޯ�����̺��۰ق�ݒ�(CtlArray�̏��Ԃƍ��킹��)
                ' ----------------------------------------------------------
                m_CtlRes = New Control() { _
                    CCmb_0, CTxt_0, CTxt_1, CCmb_1, CTxt_2, _
                    CTxt_3, CTxt_4, _
                    CCmb_2, CCmb_16, CTxt_5, CTxt_6, CTxt_7, CTxt_8, _
                    CCmb_3, CCmb_4, CTxt_14, CTxt_15, CTxt_16, CTxt_17, CTxt_18, _
                    CCmb_10, CCmb_11, CCmb_12, CCmb_13, CCmb_14, CCmb_15 _
                }
                Call SetControlData(m_CtlRes)

                ' IT, FT�֘A��÷���ޯ���z��
                m_IntialFinal = New cTxt_() { _
                    CTxt_5, CTxt_6, CTxt_7, CTxt_8 _
                }

                ' ----------------------------------------------------------
                ' ��۰�޸�ٰ���ޯ�����̺��۰ق�ݒ�(CtlArray�̏��Ԃƍ��킹��)
                ' ----------------------------------------------------------
                m_CtlProbe = New Control() { _
                    CTxt_9, CTxt_10, CTxt_11 _
                }
                Call SetControlData(m_CtlProbe)

                ' ----------------------------------------------------------
                ' ��Ĉʒu�␳��ٰ���ޯ�����̺��۰ق�ݒ�(CtlArray�̏��Ԃƍ��킹��)
                ' ----------------------------------------------------------
                m_CtlCutCorr = New Control() { _
                    CCmb_5, CCmb_6, CCmb_7, CTxt_12, CTxt_13 _
                }
                Call SetControlData(m_CtlCutCorr)

                m_CutPosCorr = New Control() { _
                    CCmb_6, CCmb_7, CTxt_12, CTxt_13 _
                }

                ' ----------------------------------------------------------
                ' ���ׂĂ�÷���ޯ��������ޯ������݂ɑ΂��A̫����ړ��ݒ�������Ȃ�
                ' ��޷��A���ٷ��ɂ��̫����ړ����鏇�Ԃź��۰ق�CtlArray�ɐݒ肷��
                ' �g�p���Ȃ����۰ق� Enabled=False �܂��� Visible=False �ɂ���
                ' ----------------------------------------------------------
                CtlArray = New Control() { _
                    CCmb_0, CTxt_0, CTxt_1, CCmb_1, CTxt_2, _
                    CTxt_3, CTxt_4, _
                    CCmb_2, CCmb_16, CTxt_5, CTxt_6, CTxt_7, CTxt_8, _
                    CCmb_3, CCmb_4, CTxt_14, CTxt_15, CTxt_16, CTxt_17, CTxt_18, _
 _
                    CTxt_10, CTxt_9, CTxt_11, _
 _
                    CCmb_5, CCmb_6, CCmb_7, CTxt_12, CTxt_13, _
 _
                    CBtn_Add, CBtn_Del _
                }
                Call SetTabIndex(CtlArray) ' ��޲��ޯ����KeyDown����Ă�ݒ肷��

                ' ----------------------------------------------------------
                ' ��ʕ\������̫����������۰ق�ݒ肷��
                ' ----------------------------------------------------------
                FIRST_CONTROL = CtlArray(0)

                ' ImeMode�̐ؑւ�L���ɂ��邽�߲���Ă�ݒ肷��
                CTxt_4.ImeMode = Windows.Forms.ImeMode.Off  ' ��̫�Ă͉p�����͂Ƃ���
                AddHandler CTxt_4.Validating, AddressOf MyBase.cTxt_Validating  ' �\���P��

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
            Dim i As Integer

            Try
                With cCombo
                    tag = DirectCast(.Tag, Integer)
                    .Items.Clear()
                    Select Case (DirectCast(.Parent.Tag, Integer))
                        ' ------------------------------------------------------------------------------
                        Case 0 ' ��R��ٰ���ޯ��
                            Select Case (tag)
                                Case 0 ' ��R�ԍ�
                                    '.Items.Add("") ' ڲ��Ĳ���ĂōĐݒ肳���
                                Case 1 ' �۰��
                                    For i = 0 To Me.m_lstSlope.Count - 1 Step 1
                                        .Items.Add(Me.m_lstSlope(i).Name)
                                    Next i

                                Case 2 ' ����Ӱ��(0:�䗦(%), 1:���l(��Βl))
                                    .Items.Add("�䗦(%)")
                                    .Items.Add("��Βl")
                                    'V2.0.0.0��

                                Case 3  ' ���胂�[�h
                                    For i = 0 To Me.m_lstMeasMode.Count - 1 Step 1
                                        .Items.Add(Me.m_lstMeasMode(i).Name)
                                    Next i
                                    'V2.0.0.0��


                                Case 4 ' ����@��'(0:���������, 1:�O�������)
                                    '.Items.Add("") ' ڲ��Ĳ���ĂōĐݒ肳���
                                Case 5 ' ����Ӱ�� 
                                    .Items.Add("����")
                                    .Items.Add("�����x")
                                    'V2.0.0.0��
                                Case 6 ' ON�@��
                                    '.Items.Add("") ' ڲ��Ĳ���ĂŐݒ肳���
                                Case 7 ' ON�@��
                                    '.Items.Add("") ' ڲ��Ĳ���ĂŐݒ肳���
                                Case 8 ' ON�@��
                                    '.Items.Add("") ' ڲ��Ĳ���ĂŐݒ肳���
                                Case 9 ' OFF�@��
                                    '.Items.Add("") ' ڲ��Ĳ���ĂŐݒ肳���
                                Case 10 ' OFF�@��
                                    '.Items.Add("") ' ڲ��Ĳ���ĂŐݒ肳���
                                Case 11 ' OFF�@��
                                    '.Items.Add("") ' ڲ��Ĳ���ĂŐݒ肳���
                                    'V2.0.0.0��
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case 1 ' ��۰�޸�ٰ���ޯ��
                            Throw New Exception("Parent.Tag - Case 1")
                            ' ------------------------------------------------------------------------------
                        Case 2 ' ��Ĉʒu�␳��ٰ���ޯ��
                            Select Case (tag)
                                Case 0 ' �␳���s(0:�Ȃ�, 1:����, 2:�蓮)
                                    .Items.Add("�Ȃ�")
                                    .Items.Add("����")
                                    .Items.Add("�蓮")
                                    .Items.Add("�����m�f���肠��")      'V1.0.4.3�E
                                Case 1 ' ��ٰ�ߔԍ�(1-999)
                                    For i = GRP_MIN To GRP_MAX Step 1
                                        .Items.Add(String.Format("{0,5:#0}", i))
                                    Next i
                                Case 2 ' ����ݔԍ�(1-50)
                                    For i = PTN_MIN To PTN_MAX Step 1
                                        .Items.Add(String.Format("{0,5:#0}", i))
                                    Next i
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
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
        ''' <summary>����������÷���ޯ���̐ݒ�������Ȃ�</summary>
        ''' <param name="cTextBox">�ݒ�������Ȃ�÷���ޯ��</param>
        Protected Overrides Sub InitTextBox(ByRef cTextBox As cTxt_)
            Dim strMin As String = ""           ' �ݒ肷��ϐ��̍ő�l
            Dim strMax As String = ""           ' �ݒ肷��ϐ��̍ŏ��l
            Dim strMsg As String = ""           ' �װ�ŕ\�����鍀�ږ�
            Dim no As String = ""
            Dim tag As Integer
            Dim strFlg As Boolean = False       ' �i�[����l�̎��(False=���l,True=������)
            Dim hexFlg As Boolean = False       ' �i�[���镶����̎��(False=10�i��,True=16�i��)
            Try
                With cTextBox
                    tag = DirectCast(.Tag, Integer)
                    no = tag.ToString("000")
                    Select Case (DirectCast(.Parent.Tag, Integer))
                        ' ------------------------------------------------------------------------------
                        Case 0 ' ��R��ٰ���ޯ��
                            ' �ް������װ���̕\����
                            strMsg = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MSG"), m_sPath, "??????")
                            Select Case (tag)
                                Case 0 ' ��Đ�
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "1")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "10")
                                    If Integer.Parse(strMax) > MAXCTN Then
                                        strMax = MAXCTN.ToString
                                    End If
                                Case 1 ' ��R��
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "1")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "4")
                                    strFlg = True
                                Case 2 ' �ڰ�ޯ�
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "0")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "16777215")
                                    strFlg = True
                                    hexFlg = True
                                Case 3 ' �ڕW�l
                                    ' (V)
                                    m_voltNOM_Min = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_VMN"), m_sPath, "-32.0")
                                    m_voltNOM_Max = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_VMX"), m_sPath, "32.0")
                                    ' (��)
                                    m_ohmNOM_Min = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_RMN"), m_sPath, "0.1")
                                    m_ohmNOM_Max = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_RMX"), m_sPath, "60000000.0")
                                    ' �����l�Ƃ��ēd���̖ڕW�l��ݒ肷��
                                    strMin = m_voltNOM_Min
                                    strMax = m_voltNOM_Max
                                Case 4 ' �\���P��
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "1")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "2")
                                    strFlg = True
                                Case 5 ' IT �����l(%�p)
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "-99.99")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "99.99")
                                    m_ITL_Min = strMin
                                    m_ITL_Max = strMax
                                Case 6 ' IT ����l(%�p)
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "-99.99")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "99.99")
                                    m_ITH_Min = strMin
                                    m_ITH_Max = strMax
                                Case 7 ' FT �����l(%�p)
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "-99.99")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "99.99")
                                    m_FTL_Min = strMin
                                    m_FTL_Max = strMax
                                Case 8 ' FT ����l(%�p)
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "-99.99")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "99.99")
                                    m_FTH_Min = strMin
                                    m_FTH_Max = strMax
                                    'V2.0.0.0��
                                Case 9  ' �đ����
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "0")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "10")
                                Case 10 ' �đ���܂ł��߰�ގ���(ms)
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "0")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "32767")
                                    'V2.0.0.0��
                                    'V2.0.0.0�G��
                                Case 11  ' IT�����
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "0")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "99")
                                Case 12 ' FT�����
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "0")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "99")
                                    'V2.0.0.0�G��
                                    'V2.0.0.0�I��
                                Case 13 ' �T�[�L�b�g�ԍ�
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "0")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "99")
                                    'V2.0.0.0�I��
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case 1 ' ��۰�޸�ٰ���ޯ��
                            ' �ް������װ���̕\����
                            strMsg = GetPrivateProfileString_S("RESISTOR_PROBE", (no & "_MSG"), m_sPath, "??????")
                            Select Case (tag)
                                Case 0 ' LO���ԍ�
                                    strMin = GetPrivateProfileString_S("RESISTOR_PROBE", (no & "_MIN"), m_sPath, "0")
                                    strMax = GetPrivateProfileString_S("RESISTOR_PROBE", (no & "_MAX"), m_sPath, "255")
                                Case 1 ' HI���ԍ�
                                    strMin = GetPrivateProfileString_S("RESISTOR_PROBE", (no & "_MIN"), m_sPath, "0")
                                    strMax = GetPrivateProfileString_S("RESISTOR_PROBE", (no & "_MAX"), m_sPath, "255")
                                Case 2 ' �ް�ޔԍ�
                                    strMin = GetPrivateProfileString_S("RESISTOR_PROBE", (no & "_MIN"), m_sPath, "0")
                                    strMax = GetPrivateProfileString_S("RESISTOR_PROBE", (no & "_MAX"), m_sPath, "255")
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case 2 ' ��Ĉʒu�␳��ٰ���ޯ��
                            ' �ް������װ���̕\����
                            strMsg = GetPrivateProfileString_S("RESISTOR_CUTCORR", (no & "_MSG"), m_sPath, "??????")
                            Select Case (tag)
                                Case 0 ' ����݈ʒuX
                                    strMin = GetPrivateProfileString_S("RESISTOR_CUTCORR", (no & "_MIN"), m_sPath, "-80.0")
                                    strMax = GetPrivateProfileString_S("RESISTOR_CUTCORR", (no & "_MAX"), m_sPath, "80.0")
                                Case 1 ' ����݈ʒuY
                                    strMin = GetPrivateProfileString_S("RESISTOR_CUTCORR", (no & "_MIN"), m_sPath, "-80.0")
                                    strMax = GetPrivateProfileString_S("RESISTOR_CUTCORR", (no & "_MAX"), m_sPath, "80.0")
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select                            ' ------------------------------------------------------------------------------
                        Case Else
                            Throw New Exception("Parent.Tag - Case Else")
                    End Select

                    Call .SetStrMsg(strMsg) ' �ް������װ���̕\�����ݒ�
                    Call .SetMinMax(strMin, strMax) ' �����l�����l�̐ݒ�
                    Dim msg As String
                    If (False = strFlg) Then ' (False=���l,True=������)
                        msg = "�͈̔͂Ŏw�肵�ĉ�����"
                    Else
                        If (False = hexFlg) Then ' (True=16�i��������)
                            msg = "�����͈̔͂Ŏw�肵�ĉ�����"
                            .MaxLength = Integer.Parse(strMax) ' SetControlData()���̏������f�Ŏg�p����
                            .TextAlign = HorizontalAlignment.Left
                        Else ' 16�i��������
                            msg = "�͈̔͂Ŏw�肵�ĉ�����"
                            ' 10�i���������16�i��������ɕϊ�����������̕�����
                            .MaxLength = ((Integer.Parse(strMax)).ToString("X")).Length
                            ' °����ߗp�ɕϊ�
                            strMin = ((Integer.Parse(strMin)).ToString("X")).ToUpper & "(Hex)"
                            strMax = ((Integer.Parse(strMax)).ToString("X")).ToUpper & "(Hex)"
                        End If
                    End If
                    Call .SetStrTip(strMin & "�`" & strMax & msg) ' °�����ү���ނ̐ݒ�
                    Call .SetLblToolTip(m_MainEdit.LblToolTip) ' ҲݕҏW��ʂ�°����ߎQ�Ɛݒ�
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#Region "÷���ޯ��������ޯ���ɒl��\������"
        ''' <summary>÷���ޯ��������ޯ���ɒl��\������</summary>
        Protected Overrides Sub SetDataToText()
            Try
                If (m_MainEdit.W_PLT.RCount < 1) Then ' ��R�� = 0 ?
                    m_ResNo = 1
                End If

                Call ChangeSlopeList()      'V2.2.1.7�@

                ' ��R��ٰ���ޯ���ݒ�
                Call SetResData()

                ' ��۰�޸�ٰ���ޯ���ݒ�
                Call SetProbeData()

                ' ��Ĉʒu�␳��ٰ���ޯ���ݒ�
                Call SetCutPosData()

                Me.Refresh()

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#Region "��R��ٰ���ޯ�����̐ݒ�"
        ''' <summary>��R��ٰ���ޯ������÷���ޯ��������ޯ���ɒl��ݒ肷��</summary>
        Private Sub SetResData()
            Dim idx As Integer

            Try
                With m_MainEdit.W_REG(m_ResNo)
                    For i As Integer = 0 To (m_CtlRes.Length - 1) Step 1
                        Select Case (i)
                            Case 0 ' ��R��, ��R�ԍ�
                                Dim rCnt As Integer = m_MainEdit.W_PLT.RCount
                                Dim cCombo As cCmb_ = DirectCast(m_CtlRes(i), cCmb_)

                                CLblRN_0.Text = rCnt.ToString() ' ��R��
                                With cCombo ' ��R�ԍ�
                                    .Items.Clear()
                                    For j As Integer = 1 To rCnt Step 1
                                        .Items.Add(String.Format("{0,5:#0}", j)) ' ����R�����J��Ԃ�
                                    Next j
                                End With
                                Call NoEventIndexChange(cCombo, (m_ResNo - 1)) ' �w���R�ԍ���ݒ�

                            Case 1 ' ��Đ�
                                ' Case 3 �۰�߂Őݒ�������Ȃ�
                                'm_CtlRes(i).Text = (.intTNN).ToString()
                            Case 2  ' ��R��
                                m_CtlRes(i).Text = .strRNO
                            Case 3 ' �۰��(1:+�d��, 2:-�d��, 4:��R, 5:�d������̂�, 6:��R����̂�, 7:NGϰ�ݸ�)
                                idx = GetComboBoxValue2Index(.intSLP, Me.m_lstSlope)

                                Call NoEventIndexChange(DirectCast(m_CtlRes(i), cCmb_), idx)
                                ' �㉺���l��ύX����K�v���Ȃ����ߺ��ı��
                                'Call ChangedSlope(.intMode, .intSLP)

                                ' ��Đ�÷���ޯ���̐ݒ�
                                If UserModule.IsMarking(m_MainEdit.W_REG, m_ResNo) Then
                                    'm_MainEdit.W_REG(m_ResNo).intMeasMode = MEAS_JUDGE_NONE
                                    ChangeSlopeAllOnOff(False)
                                Else
                                    'm_MainEdit.W_REG(m_ResNo).intMeasMode = MEAS_JUDGE_FT
                                    ChangeSlopeAllOnOff(True)
                                End If

                                If UserModule.IsMeasureOnly(m_MainEdit.W_REG, m_ResNo) Then
                                    ' �۰�߂� 7:�d������̂�, 9:��R����̂� �̏ꍇ
                                    CTxt_0.Text = 0                     ' ��Đ���0�Ƃ���
                                    CTxt_0.Enabled = False              ' �����ɂ���
                                Else
                                    CTxt_0.Text = (.intTNN).ToString()  ' ��Đ�
                                    CTxt_0.Enabled = True               ' �L���ɂ���
                                End If

                            Case 4 ' �ڰ�ޯ�
                                m_CtlRes(i).Text = (.lngRel.ToString("X")).ToUpper
                            Case 5 ' �ڕW�l
                                m_CtlRes(i).Text = (.dblNOM).ToString(DirectCast(m_CtlRes(i), cTxt_).GetStrFormat())
                                ' �۰�߂ɂ��͈͂��Đݒ肷��K�v������
                                Call ChangedSlope(.intMode, .intSLP, .dblNOM)

                            Case 6 ' �\���P��
                                m_CtlRes(i).Text = .strTANI
                            Case 7 ' ����Ӱ��(0:�䗦(%), 1:��Βl)
                                Call NoEventIndexChange(DirectCast(m_CtlRes(i), cCmb_), .intMode)
                                Call ChangedMode(.intMode, .intSLP, .dblNOM)

                            Case 8  ' ���胂�[�h
                                idx = GetComboBoxValue2Index(.intMeasMode, Me.m_lstMeasMode)

                                Call NoEventIndexChange(DirectCast(m_CtlRes(i), cCmb_), idx)

                            Case 9 ' IT�����l
                                m_CtlRes(i).Text = (.dblITL).ToString(DirectCast(m_CtlRes(i), cTxt_).GetStrFormat())
                            Case 10 ' IT����l
                                m_CtlRes(i).Text = (.dblITH).ToString(DirectCast(m_CtlRes(i), cTxt_).GetStrFormat())
                            Case 11 ' FT�����l
                                m_CtlRes(i).Text = (.dblFTL).ToString(DirectCast(m_CtlRes(i), cTxt_).GetStrFormat())
                            Case 12 ' FT����l
                                m_CtlRes(i).Text = (.dblFTH).ToString(DirectCast(m_CtlRes(i), cTxt_).GetStrFormat())
                            Case 13 ' ����@��(0=��������, 1�ȏ�O������@��ԍ�)
                                ' GP-IB�o�^�@�햼��\������(�O���d��������)
                                Dim cCombo As cCmb_ = DirectCast(m_CtlRes(i), cCmb_)
                                Dim type As Integer = Convert.ToInt32(.intMType)
                                Dim cnt As Integer = 0 ' ؽĂɒǉ��������ڐ�
                                idx = 0 ' �I��������ޯ��
                                cCombo.Items.Clear()
                                cCombo.Items.Add(" 0:���������")
                                With m_MainEdit
                                    If (0 < .W_PLT.GCount) Then ' GP-IB����@�킪�o�^����Ă���ꍇ
                                        For j As Integer = 1 To (.W_PLT.GCount) Step 1
                                            ' �ضް����ނ���̏ꍇ�A�O�������Ƃ���ؽĂɒǉ�
                                            If (.W_GPIB(j).strCTRG <> "") Then
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

                                        If (0 < idx) Then
                                            m_CtlRes(RES_TMM1).Enabled = False ' ����Ӱ�ޖ���
                                        Else
                                            m_CtlRes(RES_TMM1).Enabled = True ' ����Ӱ�ޗL��
                                        End If

                                    Else ' GP-IB����@��̓o�^���Ȃ��ꍇ
                                        .W_REG(m_ResNo).intMType = 0
                                        idx = 0 ' ���������
                                        m_CtlRes(RES_TMM1).Enabled = True ' ����Ӱ�ޗL��
                                    End If
                                End With
                                Call NoEventIndexChange(cCombo, idx)

                            Case 14 ' ����Ӱ��
                                Call NoEventIndexChange(DirectCast(m_CtlRes(i), cCmb_), .intTMM1)
                                'V2.0.0.0��
                            Case 15 ' �đ����
                                m_CtlRes(i).Text = (.intReMeas).ToString(DirectCast(m_CtlRes(i), cTxt_).GetStrFormat())
                            Case 16 ' �đ���܂ł��߰�ގ���
                                m_CtlRes(i).Text = (.intReMeas_Time).ToString(DirectCast(m_CtlRes(i), cTxt_).GetStrFormat())
                                'V2.0.0.0��
                                'V2.0.0.0�G��
                            Case 17 ' IT�����
                                m_CtlRes(i).Text = (.intITReMeas).ToString(DirectCast(m_CtlRes(i), cTxt_).GetStrFormat())
                            Case 18 ' FT�����
                                m_CtlRes(i).Text = (.intFTReMeas).ToString(DirectCast(m_CtlRes(i), cTxt_).GetStrFormat())
                                'V2.0.0.0�G��
                                'V2.0.0.0�I��
                            Case 19 ' �đ����
                                m_CtlRes(i).Text = (.intCircuitNo).ToString(DirectCast(m_CtlRes(i), cTxt_).GetStrFormat())
                                'V2.0.0.0�I��
                                'V2.0.0.0�A��
                            Case 20 ' ON�@��
                                Call CreateOnOffExtEquList(i, .intOnExtEqu(1))
                            Case 21 ' ON�@��
                                Call CreateOnOffExtEquList(i, .intOnExtEqu(2))
                            Case 22 ' ON�@��
                                Call CreateOnOffExtEquList(i, .intOnExtEqu(3))
                            Case 23 ' OFF�@��
                                Call CreateOnOffExtEquList(i, .intOffExtEqu(1))
                            Case 24 ' OFF�@��
                                Call CreateOnOffExtEquList(i, .intOffExtEqu(2))
                            Case 25 ' OFF�@��
                                Call CreateOnOffExtEquList(i, .intOffExtEqu(3))
                                'V2.0.0.0�A��
                            Case Else
                                Throw New Exception("Case " & i & ": Nothing")
                        End Select
                    Next i
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

        ''' <summary>ON�@��OFF�@������ޯ����ؽĂ��쐬���A���ڂ�I������</summary>
        ''' <param name="i">Case �ԍ�</param>
        ''' <param name="OnOffExtEqu">�Ή�����l�ւ̎Q��</param>
        ''' <remarks>GP-IB�o�^�@�햼(�O���d���̂�)��ؽĂɕ\������</remarks>
        Private Sub CreateOnOffExtEquList(ByVal i As Integer, ByRef OnOffExtEqu As Short)
            Dim idx As Integer = 0 ' �I����������ޯ�����ޯ��
            Dim cCombo As cCmb_ = DirectCast(m_CtlRes(i), cCmb_)
            cCombo.Items.Clear()
            cCombo.Items.Add(" 0:�Ȃ�")
            With m_MainEdit
                If (0 < .W_PLT.GCount) Then ' GP-IB�@�킪�o�^����Ă���ꍇ
                    Dim gpibNo As Integer = Convert.ToInt32(OnOffExtEqu) ' �ϐ��ɐݒ肳��Ă���GP-IB�o�^�ԍ�
                    Dim cnt As Integer = 0 ' ؽĂɒǉ��������ڐ�

                    For j As Integer = 1 To (.W_PLT.GCount) Step 1
                        ' ON����ނ�OFF�R�}���h���ݒ肳��Ă���ꍇ�A�O���d���Ƃ���ؽĂɒǉ�
                        If ("" <> .W_GPIB(j).strCON) AndAlso ("" <> .W_GPIB(j).strCOFF) Then
                            If (Not .W_GPIB(j).strGNAM Is Nothing) Then
                                cCombo.Items.Add(String.Format("{0,2:#0}", j) & ":" & .W_GPIB(j).strGNAM)
                            Else
                                cCombo.Items.Add(String.Format("{0,2:#0}", j) & ":")
                            End If
                            ' �ǉ�����ؽĂ��ı���
                            cnt = (cnt + 1)
                            ' OnOffExtEqu(GP-IB�o�^�ԍ�)�Ɠ������ڂ�ؽĂɒǉ����ꂽ�ꍇ��
                            ' ���̍��ڂ�I�����邽�߲��ޯ����ݒ肷��
                            ' �g�p���̋@�킪�폜���ꂽ�ꍇ�AGP-IB��ޓ��̏���(ResetResCutData)��
                            ' OnOffExtEqu��0�ƂȂ邽�� 0:�Ȃ� ���I�������
                            If (gpibNo = j) Then idx = cnt
                        End If
                    Next j

                Else ' GP-IB�@��̓o�^���Ȃ��ꍇ
                    OnOffExtEqu = 0
                    idx = 0 ' 0:�Ȃ�
                End If
            End With
            Call NoEventIndexChange(cCombo, idx)

        End Sub
#End Region

#Region "��۰�޸�ٰ���ޯ�����̐ݒ�"
        ''' <summary>��۰�޸�ٰ���ޯ������÷���ޯ��������ޯ���ɒl��ݒ肷��</summary>
        Private Sub SetProbeData()
            Try
                With m_MainEdit.W_REG(m_ResNo)
                    For i As Integer = 0 To (m_CtlProbe.Length - 1) Step 1
                        Select Case (i)
                            Case 0 ' LO���ԍ�
                                m_CtlProbe(i).Text = (.intPRL).ToString("#0")
                            Case 1 ' HI���ԍ�
                                m_CtlProbe(i).Text = (.intPRH).ToString("#0")
                            Case 2 ' �ް�ޔԍ�
                                m_CtlProbe(i).Text = (.intPRG).ToString("##0")
                            Case Else
                                Throw New Exception("Case " & i & ": Nothing")
                        End Select
                    Next i
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#Region "��Ĉʒu�␳��ٰ���ޯ�����̐ݒ�"
        ''' <summary>��Ĉʒu�␳��ٰ���ޯ������÷���ޯ��������ޯ���ɒl��ݒ肷��</summary>
        Private Sub SetCutPosData()
            Try
                With m_MainEdit.W_PTN(m_ResNo)
                    For i As Integer = 0 To (m_CtlCutCorr.Length - 1) Step 1
                        Select Case (i)
                            Case 0 ' �␳���s
                                If (SLP_VMES = m_MainEdit.W_REG(m_ResNo).intSLP) OrElse (SLP_RMES = m_MainEdit.W_REG(m_ResNo).intSLP) Then
                                    ' �۰�߂� 5:�d������̂�, 6:��R����̂� �̏ꍇ
                                    .PtnFlg = 0                         ' �␳���s����
                                    Call NoEventIndexChange(CCmb_5, 0)  ' �␳���s�����ޯ��
                                    Call ChangedCorrection(.PtnFlg)     ' �֘A���۰ق̗L���������ύX
                                    Dim cnt As Integer = 0
                                    For j As Integer = 1 To m_MainEdit.W_PLT.RCount Step 1  ' ��R����
                                        ' �␳���s����̏ꍇ�ɶ��ı���
                                        If (1 <= m_MainEdit.W_PTN(j).PtnFlg) Then cnt = (cnt + 1)
                                    Next j
                                    '                                    m_MainEdit.W_PLT.PtnCount = Convert.ToInt16(cnt) ' ����ݓo�^����ݒ�
                                    m_MainEdit.W_PLT.PtnCount = m_MainEdit.W_PLT.RCount ' ����ݓo�^����ݒ�
                                    CGrp_2.Enabled = False              ' ��Ĉʒu�␳��ٰ���ޯ���𖳌��ɂ���
                                    Exit For                            ' �ȍ~�̐ݒ�͂����Ȃ�Ȃ�

                                Else
                                    Call NoEventIndexChange(DirectCast(m_CtlCutCorr(i), cCmb_), .PtnFlg)
                                    Call ChangedCorrection(.PtnFlg)     ' �֘A���۰ق̗L���������ύX
                                    CGrp_2.Enabled = True               ' ��Ĉʒu�␳��ٰ���ޯ����L���ɂ���
                                End If

                            Case 1 ' ��ٰ�ߔԍ�(1-999)
                                Call NoEventIndexChange(DirectCast(m_CtlCutCorr(i), cCmb_), (.intGRP - 1))
                            Case 2 ' ����ݔԍ�(1-50)
                                Call NoEventIndexChange(DirectCast(m_CtlCutCorr(i), cCmb_), (.intPTN - 1))
                            Case 3 ' ����݈ʒuX
                                m_CtlCutCorr(i).Text = (.dblPosX).ToString(DirectCast(m_CtlCutCorr(i), cTxt_).GetStrFormat())
                            Case 4 ' ����݈ʒuY
                                m_CtlCutCorr(i).Text = (.dblPosY).ToString(DirectCast(m_CtlCutCorr(i), cTxt_).GetStrFormat())
                            Case Else
                                Throw New Exception("Case " & i & ": Nothing")
                        End Select
                    Next i
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#Region "���ׂĂ�÷���ޯ�����ް������������Ȃ�"
        ''' <summary>���ׂĂ�÷���ޯ�����ް������������Ȃ�</summary>
        ''' <returns>0=����, 1=�װ</returns>
        Friend Overrides Function CheckAllTextData() As Integer
            Dim ret As Integer
            Try
                m_CheckFlg = True ' ������(tabBase_Layout�ɂĎg�p)
                With m_MainEdit
                    .MTab.SelectedIndex = m_TabIdx ' ��ޕ\���ؑ�

                    ' TODO: ��R����0�ɂȂ邱�Ƃ͂Ȃ��d�l�̂��ߕs�v�Ǝv����
                    If (.W_PLT.RCount < 1) Then ' ��R�� < 1 ?
                        Dim strMsg As String
                        strMsg = "��R�f�[�^������܂���B��R�f�[�^��o�^���Ă��������B"
                        Call MsgBox(strMsg, MsgBoxStyle.OkOnly, My.Application.Info.Title)
                        ret = 1
                        Exit Try
                    End If

                    For rn As Integer = 1 To .W_PLT.RCount Step 1
                        m_ResNo = rn
                        ' ���������R�ԍ����ް�����۰قɾ�Ă���
                        Call SetDataToText()

                        ' ��R��ٰ���ޯ��
                        ret = CheckControlData(m_CtlRes)
                        If (ret <> 0) Then Exit Try

                        ' ��۰�޸�ٰ���ޯ��
                        ret = CheckControlData(m_CtlProbe)
                        If (ret <> 0) Then Exit Try

                        ' ��Ĉʒu�␳��ٰ���ޯ��
                        ret = CheckControlData(m_CtlCutCorr)
                        If (ret <> 0) Then Exit Try

                        ' ��������
                        ret = CheckRelation()
                        If (ret <> 0) Then Exit Try
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
#End Region

#Region "�ް������֐����Ăяo��"
        ''' <summary>÷���ޯ�����ް������֐����Ăяo��</summary>
        ''' <param name="cTextBox">��������÷���ޯ��</param>
        ''' <returns>0=����, 1=�װ</returns>
        Protected Overrides Function CheckTextData(ByRef cTextBox As cTxt_) As Integer
            Dim tag As Integer
            Dim ret As Integer
            Try
                ' ��R�ް��o�^������
                ' TODO: ��R����0�ɂȂ邱�Ƃ͂Ȃ��d�l�̂��ߕs�v�Ǝv����
                If (m_ResNo < 1) Then
                    Dim strMSG As String
                    strMSG = "��R�f�[�^������܂���B" & _
                                            "�ǉ��{�^�����������Ē�R�f�[�^��o�^���Ă��������B"
                    Call MsgBox(strMSG, MsgBoxStyle.OkOnly, My.Application.Info.Title)
                    ret = 1
                    Exit Try
                End If

                tag = DirectCast(cTextBox.Tag, Integer)
                With m_MainEdit.W_REG(m_ResNo)
                    Select Case (DirectCast(cTextBox.Parent.Tag, Integer))
                        ' ------------------------------------------------------------------------------
                        Case 0 ' ��R��ٰ���ޯ��
                            Select Case (tag)
                                Case 0 ' ��Đ�
                                    ' �۰�߂� 5:�d������̂�, 6:��R����̂� �ł͂Ȃ��ꍇ�̂������������Ȃ�
                                    If UserModule.IsCutResistorIncMarking(m_MainEdit.W_REG, m_ResNo) Then
                                        Dim cnt As Integer = .intTNN ' �ύX�O�̒l��ێ�
                                        ret = CheckShortData(cTextBox, .intTNN)
                                        If (cnt <> .intTNN) Then
                                            If (cnt < .intTNN) Then ' �ǉ����ꂽ�ꍇ
                                                For i As Integer = (cnt + 1) To .intTNN Step 1
                                                    Call InitCutData(m_ResNo, i) ' �ǉ����ꂽ�ް���������
                                                Next i
                                            Else ' �폜���ꂽ�ꍇ
                                                For i As Integer = (.intTNN + 1) To cnt Step 1
                                                    Call InitCutData(m_ResNo, i) ' �폜���ꂽ�ް���������
                                                Next i
                                            End If
                                            m_CutNo = 1 ' �������̶�Ĕԍ�
                                        End If
                                    End If

                                Case 1 ' ��R��
                                    ret = CheckStrData(cTextBox, .strRNO)
                                Case 2 ' �ڰ�ޯ�
                                    ret = CheckHexData(cTextBox, .lngRel)
                                Case 3 ' �ڕW�l
                                    If Not UserModule.IsMarking(m_MainEdit.W_REG, m_ResNo) Then
                                        ret = CheckDoubleData(cTextBox, .dblNOM)
                                        If ret = 0 Then
                                            ' �����l/����l��ύX
                                            Call ChangedMode(.intMode, .intSLP, .dblNOM)
                                        End If
                                    End If
                                Case 4 ' �\���P��
                                    ret = CheckStrData(cTextBox, .strTANI)
                                Case 5 ' IT �����l
                                    ret = CheckDoubleData(cTextBox, .dblITL)
                                Case 6 ' IT ����l
                                    ret = CheckDoubleData(cTextBox, .dblITH)
                                Case 7 ' FT �����l
                                    ret = CheckDoubleData(cTextBox, .dblFTL)
                                Case 8 ' FT ����l
                                    ret = CheckDoubleData(cTextBox, .dblFTH)
                                    'V2.0.0.0��
                                Case 9 ' �đ����
                                    ret = CheckShortData(cTextBox, .intReMeas)
                                Case 10 ' �đ���܂ł��߰�ގ���(ms)
                                    ret = CheckShortData(cTextBox, .intReMeas_Time)
                                    'V2.0.0.0��
                                    'V2.0.0.0�G��
                                Case 11  ' IT�����
                                    If Not UserModule.IsMarking(m_MainEdit.W_REG, m_ResNo) Then ' �}�[�L���O�̓`�F�b�N����
                                        ret = CheckShortData(cTextBox, .intITReMeas)
                                    End If
                                Case 12 ' IFT�����
                                    If Not UserModule.IsMarking(m_MainEdit.W_REG, m_ResNo) Then ' �}�[�L���O�̓`�F�b�N����
                                        ret = CheckShortData(cTextBox, .intFTReMeas)
                                    End If
                                    'V2.0.0.0�G��
                                    'V2.0.0.0�I��
                                Case 13 ' �T�[�L�b�g�ԍ�
                                    If Not UserModule.IsMarking(m_MainEdit.W_REG, m_ResNo) Then ' �}�[�L���O�̓`�F�b�N����
                                        ret = CheckShortData(cTextBox, .intCircuitNo)
                                    End If
                                    'V2.0.0.0�I��
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case 1 ' ��۰�޸�ٰ���ޯ��
                            If Not UserModule.IsMarking(m_MainEdit.W_REG, m_ResNo) Then ' �}�[�L���O�̓`�F�b�N����
                            Select Case (tag)
                                Case 0 ' LO���ԍ�
                                    ret = CheckShortData(cTextBox, .intPRL)
                                Case 1 ' HI���ԍ�
                                    ret = CheckShortData(cTextBox, .intPRH)
                                Case 2 ' �ް�ޔԍ�
                                    ret = CheckShortData(cTextBox, .intPRG)
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            End If
                            ' ------------------------------------------------------------------------------
                        Case 2 ' ��Ĉʒu�␳��ٰ���ޯ��
                            With m_MainEdit.W_PTN(m_ResNo)
                                If (1 <= .PtnFlg) Then ' �␳���s����Ȃ�m�F�������Ȃ�
                                    Select Case (tag)
                                        Case 0 ' ����݈ʒuX
                                            ret = CheckDoubleData(cTextBox, .dblPosX)
                                        Case 1 ' ����݈ʒuY
                                            ret = CheckDoubleData(cTextBox, .dblPosY)
                                        Case Else
                                            Throw New Exception("Case " & tag & ": Nothing")
                                    End Select
                                End If
                            End With
                            ' ------------------------------------------------------------------------------
                        Case Else
                            Throw New Exception("Parent.Tag - Case Else")
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

#Region "��������"
        ''' <summary>������������</summary>
        ''' <returns>0 = ����, 1 = �װ</returns>
        Protected Overrides Function CheckRelation() As Integer
            Dim strMsg As String
            Dim errIdx As Integer
            Dim ctlArray() As Control ' ��ٰ���ޯ�����Ƃ̺��۰ٔz����Q��
            Dim dMin As Double
            Dim dMax As Double

            CheckRelation = 0 ' Return�l = ����
            Try
                With m_MainEdit
                    ' ------------------------------------------------------------------------------
                    ctlArray = m_CtlRes ' ��R��ٰ���ޯ��
                    ' ������(0=��������, 1<=�O������)
                    If (.W_PLT.GCount <= 0) AndAlso (1 <= .W_REG(m_ResNo).intMType) Then
                        strMsg = "���փ`�F�b�N�G���[" & vbCrLf
                        strMsg = strMsg & "�O���@��̎w�肪�Ȃ��ꍇ�͊O�������̎w��͂ł��܂���B"
                        GoTo STP_ERR
                    End If

                    With .W_REG(m_ResNo)
                        ' IT �����l <= �ڕW�l <= IT ����l ?
                        If (0 = .intMode) Then ' ����Ӱ��(0:�䗦(%), 1:���l(��Βl))
                            dMin = .dblNOM + (.dblNOM * .dblITL * 0.01)     ' Low�ЯĒl  (LOW = (NOM*(100+Lo)/100))
                            dMax = .dblNOM + (.dblNOM * .dblITH * 0.01)     ' High�ЯĒl (HIGH= (NOM*(100+Hi)/100))
                        Else
                            dMin = .dblITL
                            dMax = .dblITH
                        End If
                        If (.dblNOM < dMin) OrElse (dMax < .dblNOM) Then
                            errIdx = RES_NOM
                            strMsg = "���փ`�F�b�N�G���[" + vbCrLf
                            strMsg = strMsg + "IT �����l <= �ڕW�l <= IT ����l�ƂȂ�悤�Ɏw�肵�Ă��������B"
                            GoTo STP_ERR
                        End If

                        ' FT �����l <= �ڕW�l <= FT ����l ?
                        If (0 = .intMode) Then ' ����Ӱ��(0:�䗦(%), 1:���l(��Βl))
                            dMin = .dblNOM + (.dblNOM * .dblFTL * 0.01)     ' Low�ЯĒl  (LOW = (NOM*(100+Lo)/100))
                            dMax = .dblNOM + (.dblNOM * .dblFTH * 0.01)     ' High�ЯĒl (HIGH= (NOM*(100+Hi)/100))
                        Else
                            dMin = .dblFTL
                            dMax = .dblFTH
                        End If
                        If (.dblNOM < dMin) OrElse (dMax < .dblNOM) Then
                            errIdx = RES_NOM
                            strMsg = "���փ`�F�b�N�G���[" + vbCrLf
                            strMsg = strMsg + "FT �����l <= �ڕW�l <= FT ����l�ƂȂ�悤�Ɏw�肵�Ă��������B"
                            GoTo STP_ERR
                        End If
                        ' ------------------------------------------------------------------------------
                        ctlArray = m_CtlProbe ' ��۰�޸�ٰ���ޯ��
                        ' ������۰�ޔԍ�����
                        'V2.0.0.0�N                        If (.intPRL = .intPRH) Then
                        If (.intPRL = .intPRH) And Not UserModule.IsMarking(m_MainEdit.W_REG, m_ResNo) Then 'V2.0.0.0�N
                            errIdx = PRB_PRH
                            strMsg = "���փ`�F�b�N�G���[" & vbCrLf
                            strMsg = strMsg & "����v���[�u�ԍ��̎w��͂ł��܂���B"
                            GoTo STP_ERR
                        End If

                    End With ' .W_REG(m_ResNo)
                End With ' m_MainEdit

                Exit Function
STP_ERR:
                Call MsgBox_CheckErr(DirectCast(ctlArray(errIdx), cTxt_), strMsg)
                CheckRelation = 1 ' Return�l = �װ

            Catch ex As Exception
                Call MsgBox_Exception(ex)
                CheckRelation = 1 ' Return�l = �װ
            End Try

        End Function
#End Region

#Region "�����ޯ���֘A����"
#Region "�۰�߂��ύX���ꂽ���ɖڕW�l÷���ޯ���̐ݒ��ύX����"
        ''' <summary>�۰�߂��ύX���ꂽ���ɖڕW�l÷���ޯ���̐ݒ��ύX����</summary>
        ''' <param name="mode">����Ӱ��</param>
        ''' <param name="slp">�۰��(1:+�d��, 2:-�d��, 4:��R, 5:�d������̂�, 6:��R����̂�, 7:NGϰ�ݸ�)</param>
        ''' <param name="dNOM">�ڕW�l</param>
        Private Sub ChangedSlope(ByVal mode As Integer, ByVal slp As Integer, ByVal dNOM As Double)
            Dim strMin As String
            Dim strMax As String
            Try
                If (SLP_VTRIMPLS = slp Or SLP_VTRIMMNS = slp Or SLP_VMES = slp) Then ' (�d���̂�)
                    strMin = m_voltNOM_Min
                    strMax = m_voltNOM_Max
                Else
                    strMin = m_ohmNOM_Min
                    strMax = m_ohmNOM_Max
                End If

                With DirectCast(m_CtlRes(RES_NOM), cTxt_) ' �ڕW�l÷���ޯ��
                    Call .SetMinMax(strMin, strMax) ' �����l�����l�̐ݒ�
                    Call .SetStrTip(strMin & "�`" & strMax & "�͈̔͂Ŏw�肵�ĉ�����") ' °�����ү���ނ̐ݒ�
                End With

                Call ChangedMode(mode, slp, dNOM)

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#Region "����Ӱ�ނ��ύX���ꂽ����IT,FT÷���ޯ���̐ݒ��ύX����"
        ''' <summary>����Ӱ�ނ��ύX���ꂽ����IT,FT÷���ޯ���̐ݒ��ύX����</summary>
        ''' <param name="mode">����Ӱ��(0:�䗦(%), 1:���l(��Βl)</param>
        ''' <param name="slp">�۰��(1:+�d��, 2:-�d��, 4:��R, 5:�d������̂�, 6:��R����̂�, 7:NGϰ�ݸ�)</param>
        ''' <param name="dNOM">�ڕW�l</param>
        Private Sub ChangedMode(ByVal mode As Integer, ByVal slp As Integer, ByVal dNOM As Double)
            Dim Length As Integer = (m_IntialFinal.Length - 1)
            Dim strMin(Length) As String
            Dim strMax(Length) As String
            Dim i As Integer
            Dim dValue As Double

            Try
                If (JUDGE_MODE_RATIO = mode) Then ' (0:�䗦(%), 1:���l(��Βl))
                    strMin(0) = m_ITL_Min   ' IT �����l
                    strMax(0) = m_ITL_Max
                    strMin(1) = m_ITH_Min   ' IT ����l
                    strMax(1) = m_ITH_Max
                    strMin(2) = m_FTL_Min   ' FT �����l
                    strMax(2) = m_FTL_Max
                    strMin(3) = m_FTH_Min   ' FT ����l
                    strMax(3) = m_FTH_Max
                Else

                    For i = 0 To Length Step 1
                        If (slp = SLP_RTRM) Or (slp = SLP_RMES) Then    ' ��R�̏ꍇ
                            If TypeOf Me.m_CtlRes(RES_NOM) Is cTxt_ Then
                                dValue = -(dNOM - Double.Parse(m_ohmNOM_Min))
                                strMin(i) = dValue.ToString(DirectCast(m_CtlRes(RES_NOM), cTxt_).GetStrFormat())
                                dValue = Double.Parse(m_ohmNOM_Max) - dNOM
                                strMax(i) = dValue.ToString(DirectCast(m_CtlRes(RES_NOM), cTxt_).GetStrFormat())
                            End If
                        Else                                            ' �d���̏ꍇ
                            strMin(i) = m_voltNOM_Min
                            strMax(i) = m_voltNOM_Max
                        End If
                    Next i
                End If

                For i = 0 To Length Step 1
                    With m_IntialFinal(i)
                        Call .SetMinMax(strMin(i), strMax(i)) ' �����l�����l�̐ݒ�
                        Call .SetStrTip(strMin(i) & "�`" & strMax(i) & "�͈̔͂Ŏw�肵�ĉ�����") ' °�����ү���ނ̐ݒ�
                    End With
                Next

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#Region "�␳���s�����ޯ���̲��ޯ�����ύX���ꂽ���Ɋ֘A���۰ق̗L���������ύX����"
        ''' <summary>�␳���s�����ޯ���̲��ޯ�����ύX���ꂽ���Ɋ֘A���۰ق̗L���������ύX����</summary>
        ''' <param name="idx">0:�Ȃ�, 1:����, 2:����+�蓮</param>
        Private Sub ChangedCorrection(ByVal idx As Integer)
            Dim tf As Boolean
            'V1.0.4.3�E            If (0 = idx) OrElse (3 = idx) Then ' �␳�Ȃ�
            If (CUT_PATTERN_NONE = idx) Then ' �␳�Ȃ�
                tf = False
            Else ' �␳����
                tf = True
            End If

            For Each ctl As Control In m_CutPosCorr
                ctl.Enabled = tf
            Next
        End Sub
#End Region
#End Region

#Region "�ǉ���폜���݊֘A����"
        ''' <summary>��R�ް���ǉ��܂��͍폜���A���̒�R�ް�������������</summary>
        ''' <param name="addDel">1=Add, (-1)=Del</param>
        Private Sub SortResistorData(ByVal addDel As Integer)
            Dim iStart As Integer
            Dim iEnd As Integer
            Dim dir As Integer = (-1) * addDel ' Add=(-1), Del=1�ɂ���
            Try
                With m_MainEdit
                    If (1 = addDel) Then ' �ǉ��̏ꍇ
                        .W_PLT.RCount = Convert.ToInt16(.W_PLT.RCount + 1) ' �o�^��R����ǉ�����
                        iStart = .W_PLT.RCount ' �o�^����Ă����R������
                        iEnd = (m_ResNo + 1) ' �ǉ������R�ް��ԍ�+1�܂ŁA�O���ް������ɂ��炷
                    Else ' �폜�̏ꍇ
                        iStart = m_ResNo ' �폜�����R�ް��ԍ�����
                        iEnd = (.W_PLT.RCount - 1) ' �o�^����Ă����R�ް���-1�܂ŁA�����ް���O�ɂ��炷
                    End If

                    For rn As Integer = iStart To iEnd Step dir
                        With .W_REG(rn)
                            .strRNO = m_MainEdit.W_REG(rn + dir).strRNO     ' ��R��
                            .strTANI = m_MainEdit.W_REG(rn + dir).strTANI   ' �\���P��("V","��" ��)
                            .intSLP = m_MainEdit.W_REG(rn + dir).intSLP     ' �d���ω��۰��(1:+V, 2:-V, 4:��R)
                            .lngRel = m_MainEdit.W_REG(rn + dir).lngRel     ' �ڰ�ޯ�
                            .dblNOM = m_MainEdit.W_REG(rn + dir).dblNOM     ' ���ݸ� �ڕW�l
                            .dblITL = m_MainEdit.W_REG(rn + dir).dblITL     ' �������艺���l (ITLO)
                            .dblITH = m_MainEdit.W_REG(rn + dir).dblITH     ' �����������l (ITHI)
                            .dblFTL = m_MainEdit.W_REG(rn + dir).dblFTL     ' �I�����艺���l (FTLO)
                            .dblFTH = m_MainEdit.W_REG(rn + dir).dblFTH     ' �I���������l (FTHI)
                            .intMode = m_MainEdit.W_REG(rn + dir).intMode   ' ����Ӱ��(0:�䗦(%), 1:���l(��Βl))
                            .intMeasMode = m_MainEdit.W_REG(rn + dir).intMeasMode       ' ���胂�[�h(0:�Ȃ�, 1:IT�̂� 2:FT�̂� 3:IT,FT����)
                            .intTMM1 = m_MainEdit.W_REG(rn + dir).intTMM1               ' ���[�h(0:����(�R���p���[�^��ϕ����[�h), 1:�����x(�ϕ����[�h))
                            .intPRH = m_MainEdit.W_REG(rn + dir).intPRH     ' HI����۰�ޔԍ�
                            .intPRL = m_MainEdit.W_REG(rn + dir).intPRL     ' LO����۰�ޔԍ�
                            .intPRG = m_MainEdit.W_REG(rn + dir).intPRG     ' �ް����۰�ޔԍ�
                            .intMType = m_MainEdit.W_REG(rn + dir).intMType ' ������(0=��������, 1=�O������)
                            .intTNN = m_MainEdit.W_REG(rn + dir).intTNN     ' ��Đ�(1�`9)
                            'V2.0.0.0��
                            .intReMeas = m_MainEdit.W_REG(rn + dir).intReMeas           ' �đ����
                            .intReMeas_Time = m_MainEdit.W_REG(rn + dir).intReMeas_Time ' ON����߰�ގ���(ms)
                            For i As Integer = 1 To EXTEQU Step 1
                                .intOnExtEqu(i) = m_MainEdit.W_REG(rn + dir).intOnExtEqu(i)     ' �n�m�@��P�i�n�m����O���@��P�`�R�j
                                .intOffExtEqu(i) = m_MainEdit.W_REG(rn + dir).intOffExtEqu(i)   ' �n�e�e�@��P�i�n�e�e����O���@��P�`�R�j
                            Next i
                            'V2.0.0.0��
                            'V2.0.0.0�G��
                            .intITReMeas = m_MainEdit.W_REG(rn + dir).intITReMeas       ' IT�����
                            .intFTReMeas = m_MainEdit.W_REG(rn + dir).intFTReMeas   ' FT�����
                            'V2.0.0.0�G��
                            'V2.0.0.0�I��
                            .intCircuitNo = m_MainEdit.W_REG(rn + dir).intCircuitNo     ' �đ����
                            'V2.0.0.0�I��

                            For cn As Integer = 1 To MAXCTN Step 1
                                With .STCUT(cn)
                                    .intCUT = m_MainEdit.W_REG(rn + dir).STCUT(cn).intCUT     ' ��ĕ��@(1:�ׯ�ݸ�, 2:���ޯ��, 3:�߼޼��ݸޖ������ޯ��)
                                    .intCTYP = m_MainEdit.W_REG(rn + dir).STCUT(cn).intCTYP   ' ��Č`��(1:��ڰ�, 2:L���)
                                    .dblSTX = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblSTX     ' ���ݸ� ���ē_ X
                                    .dblSTY = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblSTY     ' ���ݸ� ���ē_ Y
                                    .dblSX2 = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblSX2     ' ���ݸ� ���ē_2 X
                                    .dblSY2 = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblSY2     ' ���ݸ� ���ē_2 Y
                                    .dblDL2 = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblDL2     ' ��2�̶�Ē�(�ЯĶ�ė�mm(L��ݑO))
                                    .dblDL3 = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblDL3     ' ��3�̶�Ē�(�ЯĶ�ė�mm(L��݌�))
                                    .intANG = m_MainEdit.W_REG(rn + dir).STCUT(cn).intANG     ' ��ĕ���1
                                    .intANG2 = m_MainEdit.W_REG(rn + dir).STCUT(cn).intANG2   ' ��ĕ���2
                                    .intQF1 = m_MainEdit.W_REG(rn + dir).STCUT(cn).intQF1     ' Qڰ�(0.1KHz)
                                    .dblV1 = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblV1       ' ��ё��x(mm/s)
                                    .dblCOF = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblCOF     ' ��ĵ�(%)
                                    .dblLTP = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblLTP     ' L��� �߲��(%)
                                    .intTMM = m_MainEdit.W_REG(rn + dir).STCUT(cn).intTMM     ' Ӱ��(0:����(����ڰ���ϕ�Ӱ��), 1:�����x(�ϕ�Ӱ��))
                                    .intMType = m_MainEdit.W_REG(rn + dir).STCUT(cn).intMType ' �����^�O�������
                                    .cFormat = m_MainEdit.W_REG(rn + dir).STCUT(cn).cFormat   '###1042�@ �����f�[�^
                                    'V2.1.0.0�@�� �J�b�g���̒�R�l�ω��ʔ���@�\�ǉ�
                                    .iVariationRepeat = m_MainEdit.W_REG(rn + dir).STCUT(cn).iVariationRepeat   ' ���s�[�g�L��
                                    .iVariation = m_MainEdit.W_REG(rn + dir).STCUT(cn).iVariation               ' ����L��
                                    .dRateOfUp = m_MainEdit.W_REG(rn + dir).STCUT(cn).dRateOfUp                 ' �㏸��
                                    .dVariationLow = m_MainEdit.W_REG(rn + dir).STCUT(cn).dVariationLow         ' �����l
                                    .dVariationHi = m_MainEdit.W_REG(rn + dir).STCUT(cn).dVariationHi           ' ����l
                                    'V2.1.0.0�@��

                                    'V2.1.0.0�C ADD ��
                                    For i As Integer = 1 To MAX_LCUT Step 1         ' MAX���Đ����J�Ԃ�
                                        .dCutLen(i) = m_MainEdit.W_REG(rn + dir).STCUT(cn).dCutLen(i)                   ' �J�b�g��
                                        .dQRate(i) = m_MainEdit.W_REG(rn + dir).STCUT(cn).dQRate(i)                     ' �p���[�g
                                        .dSpeed(i) = m_MainEdit.W_REG(rn + dir).STCUT(cn).dSpeed(i)                     ' ���x
                                        .dAngle(i) = m_MainEdit.W_REG(rn + dir).STCUT(cn).dAngle(i)                     ' �����i�p�x�j
                                        .dTurnPoint(i) = m_MainEdit.W_REG(rn + dir).STCUT(cn).dTurnPoint(i)             ' �k�^�[���|�C���g
                                    Next i
                                    .intRetraceCnt = m_MainEdit.W_REG(rn + dir).STCUT(cn).intRetraceCnt                       ' ���g���[�X�J�b�g�{��
                                    For i As Integer = 1 To MAX_RETRACECUT Step 1           ' MAX���Đ����J�Ԃ�
                                        .dblRetraceOffX(i) = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblRetraceOffX(i)     ' ���g���[�X�̃I�t�Z�b�g�w
                                        .dblRetraceOffY(i) = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblRetraceOffY(i)     ' ���g���[�X�̃I�t�Z�b�g�x
                                        .dblRetraceQrate(i) = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblRetraceQrate(i)   ' �X�g���[�g�J�b�g�E���g���[�X��Q���[�g(0.1KHz)�Ɏg�p
                                        .dblRetraceSpeed(i) = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblRetraceSpeed(i)   ' �X�g���[�g�J�b�g�E���g���[�X�̃g�������x(mm/s)�Ɏg�p
                                    Next i
                                    'V2.1.0.0�C ADD ��

                                    ' ���ޯ����ď��ݒ�
                                    For ix As Integer = 1 To MAXIDX Step 1 ' MAX���ޯ����Đ����J�Ԃ�
                                        .intIXN(ix) = m_MainEdit.W_REG(rn + dir).STCUT(cn).intIXN(ix) ' ���ޯ����Đ�1-5
                                        .dblDL1(ix) = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblDL1(ix) ' ��Ē�1-5
                                        .lngPAU(ix) = m_MainEdit.W_REG(rn + dir).STCUT(cn).lngPAU(ix) ' �߯����߰�ގ���1-5
                                        .dblDEV(ix) = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblDEV(ix) ' �덷1-5(%)
                                        .intIXMType(ix) = m_MainEdit.W_REG(rn + dir).STCUT(cn).intIXMType(ix) ' ����@��
                                        .intIXTMM(ix) = m_MainEdit.W_REG(rn + dir).STCUT(cn).intIXTMM(ix) ' ����Ӱ��
                                    Next ix

                                    ' FL���H����
                                    For fl As Integer = 1 To MAXCND Step 1
                                        .intCND(fl) = m_MainEdit.W_REG(rn + dir).STCUT(cn).intCND(fl) ' FL�ݒ�No.
                                    Next fl
                                End With ' .STCUT(cn)
                            Next cn
                        End With ' .W_REG(rn)

                        ' ��Ĉʒu�␳
                        With .W_PTN(rn)
                            .PtnFlg = m_MainEdit.W_PTN(rn + dir).PtnFlg     ' �␳���s(0:�Ȃ�, 1:����, 2:����+�蓮)
                            .intGRP = m_MainEdit.W_PTN(rn + dir).intGRP     ' ��ٰ�ߔԍ�
                            .intPTN = m_MainEdit.W_PTN(rn + dir).intPTN     ' ����ݔԍ�
                            .dblPosX = m_MainEdit.W_PTN(rn + dir).dblPosX   ' �����X
                            .dblPosY = m_MainEdit.W_PTN(rn + dir).dblPosY   ' �����Y
                            .dblDRX = m_MainEdit.W_PTN(rn + dir).dblDRX     ' ����ʕۑ�ܰ�X
                            .dblDRY = m_MainEdit.W_PTN(rn + dir).dblDRY     ' ����ʕۑ�ܰ�Y
                        End With
                    Next rn

                    ' �߂ĕs�v�ƂȂ����ް�������������
                    If (1 = addDel) Then ' �ǉ��̏ꍇ
                        Call InitResData(m_ResNo)       ' �ǉ�������R�ް���������
                    Else ' �폜�̏ꍇ
                        Dim lastRn As Integer = Convert.ToInt32(.W_PLT.RCount)
                        Call InitResData(lastRn)        ' �Ō���ް���������
                        .W_PLT.RCount = Convert.ToInt16(lastRn - 1) ' �o�^��R����-1����

                        ' �ŏI��R�̍폜�Ȃ猻�݂̒�R�ԍ����ŏI��R�ԍ��Ƃ���
                        If (.W_PLT.RCount < m_ResNo) Then m_ResNo = .W_PLT.RCount
                    End If
                End With

                ' ��R�ް�����ʍ��ڂɐݒ�
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
            Try
                cCombo = DirectCast(sender, cCmb_)
                tag = DirectCast(cCombo.Tag, Integer)
                idx = cCombo.SelectedIndex

                With m_MainEdit
                    Select Case (DirectCast(cCombo.Parent.Tag, Integer)) ' ��ٰ���ޯ�������
                        ' ------------------------------------------------------------------------------
                        Case 0 ' ��R��ٰ���ޯ��
                            Select Case (tag)
                                Case 0 ' ��R�ԍ�
                                    m_ResNo = (idx + 1)
                                    ' �Ή������ް���÷���ޯ���ɾ�Ă���
                                    Call SetDataToText()

                                    If UserModule.IsMarking(m_MainEdit.W_REG, m_ResNo) Then
                                        m_MainEdit.W_REG(m_ResNo).intMeasMode = MEAS_JUDGE_NONE
                                        ChangeSlopeAllOnOff(False)
                                    Else
                                        ChangeSlopeAllOnOff(True)

                                        If m_MainEdit.W_REG(m_ResNo).intMType > 0 Then
                                            m_CtlRes(RES_TMM1).Enabled = False
                                        Else
                                            m_CtlRes(RES_TMM1).Enabled = True
                                        End If
                                    End If

                                Case 1 ' �۰��
                                    Dim iSlp As Integer

                                    iSlp = GetComboBoxName2Value(cCombo.Text, Me.m_lstSlope)

                                    With .W_REG(m_ResNo)
                                        .intSLP = Convert.ToInt16(iSlp)         ' �۰�߂�ݒ�

                                        If UserModule.IsMarking(m_MainEdit.W_REG, m_ResNo) Then
                                            m_MainEdit.W_REG(m_ResNo).intMeasMode = MEAS_JUDGE_NONE
                                            'Call NoEventIndexChange(CCmb_16, 0) ' �␳���s�����ޯ��
                                            ChangeSlopeAllOnOff(False)
                                        Else
                                            'm_MainEdit.W_REG(m_ResNo).intMeasMode = MEAS_JUDGE_FT
                                            'CCmb_16.Enabled = True
                                            ChangeSlopeAllOnOff(True)
                                        End If

                                        Call ChangedSlope(.intMode, iSlp, .dblNOM)

                                        ' ------------------------
                                        'If (SLP_VMES <= iSlp And iSlp <= SLP_RMES) Then
                                        If UserModule.IsMeasureOnly(m_MainEdit.W_REG, m_ResNo) Then
                                            ' �۰�߂� 7:�d������̂�, 9:��R����̂� �̏ꍇ
                                            CTxt_0.Text = 0                     ' ��Đ��̕\����0�ɂ���
                                            CTxt_0.Enabled = False              ' �����ɂ���
                                            If (1 < .intTNN) Then               ' ���Ƃ̶�Đ�
                                                For i As Integer = 2 To .intTNN Step 1
                                                    Call InitCutData(m_ResNo, i) ' 2�ȍ~�̶���ް���������
                                                Next i
                                                .intTNN = 1                     ' ��Đ���1�ɂ���
                                                m_CutNo = 1                     ' �������̶�Ĕԍ�
                                            End If

                                            With m_MainEdit.W_PTN(m_ResNo)
                                                .PtnFlg = PTN_NONE                 ' �␳���s����
                                                Call NoEventIndexChange(CCmb_5, 0) ' �␳���s�����ޯ��
                                                Call ChangedCorrection(.PtnFlg) ' �֘A���۰ق̗L���������ύX
                                                Dim cnt As Integer = 0
                                                m_MainEdit.W_PLT.PtnCount = m_MainEdit.W_PLT.RCount

                                            End With
                                            CGrp_2.Enabled = False              ' ��Ĉʒu�␳��ٰ���ޯ���𖳌��ɂ���

                                        Else
                                            CTxt_0.Text = (.intTNN).ToString()  ' ��Đ�
                                            CTxt_0.Enabled = True               ' �L���ɂ���

                                            CGrp_2.Enabled = True               ' ��Ĉʒu�␳��ٰ���ޯ����L���ɂ���
                                        End If

                                        ' ------------------------
                                    End With

                                Case 2 ' ����Ӱ��
                                    ' (0:�䗦(%), 1:���l(��Βl))
                                    .W_REG(m_ResNo).intMode = Convert.ToInt16(idx)
                                    ' �֘A������۰ق̍ő奍ŏ��l��°����߂Ȃǂ̐ݒ��ύX����
                                    Call ChangedMode(.W_REG(m_ResNo).intMode, .W_REG(m_ResNo).intSLP, .W_REG(m_ResNo).dblNOM)

                                Case 3  ' ���胂�[�h
                                    .W_REG(m_ResNo).intMeasMode = GetComboBoxName2Value(cCombo.Text, Me.m_lstMeasMode)

                                Case 4 ' ����@��(0:���������, 1�ȏ�͊O�������)
                                    ' �o�^����Ă��鑪��@��ؽĂ̐��l��ݒ肷��( 1:NAME=1, 10:NAME=10)
                                    .W_REG(m_ResNo).intMType = Short.Parse((cCombo.Text).Substring(0, 2))
                                    ' �O�������̏ꍇ����Ӱ�ނ𖳌��ɂ���
                                    If (0 < idx) Then
                                        m_CtlRes(RES_TMM1).Enabled = False
                                    Else
                                        m_CtlRes(RES_TMM1).Enabled = True
                                    End If
                                Case 5 ' ����Ӱ��
                                    .W_REG(m_ResNo).intTMM1 = Convert.ToInt16(idx)
                                    'V2.0.0.0�A
                                Case 6 ' ON�@��
                                    ' �o�^����Ă��鑪��@��ؽĂ̐��l��ݒ肷��( 1:NAME=1, 10:NAME=10)
                                    .W_REG(m_ResNo).intOnExtEqu(1) = Short.Parse((cCombo.Text).Substring(0, 2))
                                Case 7 ' ON�@��
                                    .W_REG(m_ResNo).intOnExtEqu(2) = Short.Parse((cCombo.Text).Substring(0, 2))
                                Case 8 ' ON�@��
                                    .W_REG(m_ResNo).intOnExtEqu(3) = Short.Parse((cCombo.Text).Substring(0, 2))
                                Case 9 ' OFF�@��
                                    ' �o�^����Ă��鑪��@��ؽĂ̐��l��ݒ肷��( 1:NAME=1, 10:NAME=10)
                                    .W_REG(m_ResNo).intOffExtEqu(1) = Short.Parse((cCombo.Text).Substring(0, 2))
                                Case 10 ' OFF�@��
                                    .W_REG(m_ResNo).intOffExtEqu(2) = Short.Parse((cCombo.Text).Substring(0, 2))
                                Case 11 ' OFF�@��
                                    .W_REG(m_ResNo).intOffExtEqu(3) = Short.Parse((cCombo.Text).Substring(0, 2))
                                    'V2.0.0.0�A
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case 1 ' ��۰�޸�ٰ���ޯ��
                            Throw New Exception("Parent.Tag - Case 1")
                            ' ------------------------------------------------------------------------------
                        Case 2 ' ��Ĉʒu�␳��ٰ���ޯ��
                            With .W_PTN(m_ResNo)
                                Select Case (tag)
                                    Case 0 ' �␳���s(0:�Ȃ�, 1:����, 2:����+�蓮)
                                        .PtnFlg = Convert.ToInt16(idx)
                                        Call ChangedCorrection(idx) ' �֘A���۰ق̗L���������ύX
                                        Dim cnt As Integer = 0
                                        For i As Integer = 1 To m_MainEdit.W_PLT.RCount Step 1
                                            If (1 <= m_MainEdit.W_PTN(i).PtnFlg) Then cnt = (cnt + 1) ' �␳���s����̏ꍇ�ɶ��ı���
                                        Next i
                                        'm_MainEdit.W_PLT.PtnCount = Convert.ToInt16(cnt) ' ����ݓo�^����ݒ�
                                        m_MainEdit.W_PLT.PtnCount = m_MainEdit.W_PLT.RCount ' ����ݓo�^����ݒ�

                                    Case 1 ' ��ٰ�ߔԍ�(1-999)
                                        .intGRP = Convert.ToInt16(idx + 1)
                                    Case 2 ' ����ݔԍ�(1-50)
                                        .intPTN = Convert.ToInt16(idx + 1)
                                    Case Else
                                        Throw New Exception("Case " & tag & ": Nothing")
                                End Select
                            End With
                            ' ------------------------------------------------------------------------------
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
            Dim strMsg As String ' ү�����ޯ���̷��߼�ݕ\���p
            Dim refOpt As Short ' ��߼������(0=�O�ɒǉ� ,1=��ɒǉ�)
            Dim ret As Integer
            Try
                ' �o�^������
                If (MAXRNO <= m_MainEdit.W_PLT.RCount) Then ' �o�^��OK ?
                    strMsg = "����ȏ��R�f�[�^�͓o�^�ł��܂���B"
                    Call MsgBox(strMsg, DirectCast( _
                                MsgBoxStyle.OkOnly + _
                                MsgBoxStyle.Information, MsgBoxStyle), _
                                My.Application.Info.Title)
                    Exit Sub
                End If

                ' �m�Fү���ނ�\��("��R�f�[�^��ǉ����܂�")
                ret = MsgBox_AddClick("��R�f�[�^", refOpt) ' ү���ޕ\��
                If (ret <> cFRS_ERR_ADV) Then Exit Sub ' Cancel�Ȃ�Return
                If (refOpt = 1) Then ' �\���ް��̌�ɒǉ� ?
                    m_ResNo = (m_ResNo + 1) ' m_ResNo = ���݂��ް��ԍ� + 1
                Else ' �\���ް��̑O�ɒǉ�
                    m_ResNo = m_ResNo ' m_ResNo = ���݂��ް��ԍ�
                End If

                ' �ް���1��ɂ��炷
                Call SortResistorData(1)

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
                If (1 = m_MainEdit.W_PLT.RCount) Then Exit Sub ' �o�^��1�Ȃ�NOP
                strMsg = "���݂̒�R�f�[�^���폜���܂��B��낵���ł����H"
                ret = MsgBox(strMsg, DirectCast( _
                            MsgBoxStyle.OkCancel + _
                            MsgBoxStyle.Exclamation + _
                            MsgBoxStyle.DefaultButton2, MsgBoxStyle), _
                            My.Application.Info.Title)
                If (ret = MsgBoxResult.Cancel) Then Exit Sub ' Cancel(RESET��) ?

                ' �ް���1�O�ɂ߂�
                Call SortResistorData(-1)

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

        'V2.0.0.0��
#Region "�X���[�v�f�[�^�����ݒ�"
        ''' <summary>
        ''' �X���[�v�f�[�^�����ݒ�
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub InitSlopeData()
            Dim data As New ComboDataStruct
#If VOLTAGE_USE Then
            data.SetData("�{�d���g���~���O", SLP_VTRIMPLS)
            Me.m_lstSlope.Add(data)
            data.SetData("�|�d���g���~���O", SLP_VTRIMMNS)
            Me.m_lstSlope.Add(data)
#End If
            data.SetData("��R�g���~���O", SLP_RTRM)
            Me.m_lstSlope.Add(data)
#If VOLTAGE_USE Then
            data.SetData("�d������̂�", SLP_VMES)
            Me.m_lstSlope.Add(data)
#End If
            data.SetData("��R����̂�", SLP_RMES)
            Me.m_lstSlope.Add(data)

#If NG_MARKING_USE Then
            data.SetData("�m�f�}�[�L���O", SLP_NG_MARK)
            Me.m_lstSlope.Add(data)
#End If

#If OK_MARKING_USE Then
            data.SetData("�n�j�}�[�L���O", SLP_OK_MARK)
            Me.m_lstSlope.Add(data)
#End If

#If OK_NG_MARKING_USE Then
            data.SetData("�m�f�}�[�L���O", SLP_NG_MARK)
            Me.m_lstSlope.Add(data)
            data.SetData("�n�j�}�[�L���O", SLP_OK_MARK)
            Me.m_lstSlope.Add(data)
#End If
            'V2.2.1.7�@ ��
            If (m_MainEdit.W_stUserData.iTrimType = 5) Then
                data.SetData("�}�[�N��", SLP_MARK)
                Me.m_lstSlope.Add(data)
            End If

            'data.SetData("�}�[�N��", SLP_MARK)
            'Me.m_lstSlope.Add(data)
            'V2.2.1.7�@ ��

        End Sub
#End Region

#Region "���胂�[�h�f�[�^�����ݒ�"
        ''' <summary>
        ''' ���胂�[�h�f�[�^�����ݒ�
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub InitMeasModeData()
            Dim data As New ComboDataStruct

            data.SetData("�Ȃ�", MEAS_JUDGE_NONE)
            Me.m_lstMeasMode.Add(data)
            data.SetData("IT�̂�", MEAS_JUDGE_IT)
            Me.m_lstMeasMode.Add(data)
            data.SetData("FT�̂�", MEAS_JUDGE_FT)
            Me.m_lstMeasMode.Add(data)
            data.SetData("IT,FT����", MEAS_JUDGE_BOTH)
            Me.m_lstMeasMode.Add(data)
        End Sub
#End Region

#Region "�X���[�v�ɂ����͕s�\�̕ύX"
        ''' <summary>
        ''' �m�f�}�[�L���O�̎����̍��ڂ͑S�Ăn�e�e����B
        ''' </summary>
        ''' <param name="OnOff"></param>
        ''' <remarks></remarks>
        Private Sub ChangeSlopeAllOnOff(ByVal OnOff As Boolean)
            Try
                If OnOff Then
                    ' ��R�f�[�^
                    For i As Integer = 4 To (m_CtlRes.Length - 1) Step 1
                        m_CtlRes(i).Enabled = True
                    Next i
                    ' �v���[�u
                    For i As Integer = 0 To (m_CtlProbe.Length - 1) Step 1
                        m_CtlProbe(i).Enabled = True
                    Next i

#If ADDITIONAL_GPIB Then        ' �ǉ�GPIB�@��
                    If m_MainEdit.W_USER.iTrimType = 2 Or m_MainEdit.W_USER.iTrimType = 3 Then
                        CCmb_GPIB.Enabled = True
                        CTxt_GPIB_LO.Enabled = True
                        CTxt_GPIB_HI.Enabled = True
                    Else
                        CCmb_GPIB.Enabled = False
                        CTxt_GPIB_LO.Enabled = False
                        CTxt_GPIB_HI.Enabled = False
                    End If
#End If

                Else
                    ' ��R�f�[�^
                    For i As Integer = 4 To (m_CtlRes.Length - 1) Step 1
                        m_CtlRes(i).Enabled = False
                    Next i
                    ' �v���[�u
                    For i As Integer = 0 To (m_CtlProbe.Length - 1) Step 1
                        m_CtlProbe(i).Enabled = False
                    Next i
                End If

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try
        End Sub
#End Region

#Region "�X���[�v���X�g�̕ύX"
        Public Sub ChangeSlopeList()
            Try
                Dim cCombo As cCmb_ = DirectCast(m_CtlRes(RES_SLOPE), cCmb_)
                m_lstSlope = Nothing
                m_lstSlope = New List(Of ComboDataStruct)
                InitSlopeData()
                With cCombo
                    .Items.Clear()
                    For i As Integer = 0 To Me.m_lstSlope.Count - 1 Step 1
                        .Items.Add(Me.m_lstSlope(i).Name)
                    Next i
                End With
            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try
        End Sub
#End Region
        'V2.0.0.0��

    End Class
End Namespace
