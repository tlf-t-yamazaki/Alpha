Option Explicit On
Option Strict Off

Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports LaserFront.Trimmer.DefWin32Fnc

Namespace FormEdit
    Friend Class tabUser
        Inherits tabBase

#Region "�錾"
        Private Const SYS_ZON As Integer = 11       ' ���������Ŏg�p(m_CtlSystem�ł̲��ޯ��)
        Private Const SYS_ZOFF As Integer = 12      ' ���������Ŏg�p(m_CtlSystem�ł̲��ޯ��)

        Private m_CtlSystem() As Control            ' USER�̺��۰ٔz��
        Private m_CtlResistor(,) As Control         ' ��R�ݒ�̒�R�ԍ��ޯ���̺��۰ٔz��
        Private m_DispResistor(,) As Control        ' ��R���ɉ����ĕ\��/��\����؂�ւ�����۰�
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
                ' EDIT_DEF_User.ini������ޖ���ݒ�
                TAB_NAME = GetPrivateProfileString_S("USER_LABEL", "TAB_NAM", m_sPath, "????")

                ' ----------------------------------------------------------
                ' EDIT_DEF_User.ini�����ٰ���ޯ���ɕ\������ݒ�
                ' ----------------------------------------------------------
                GrpArray = New cGrp_() { _
                    CGrp_0, CGrp_1, CGrp_2, CGrp_3, CGrp_4 _
                }
                For i As Integer = 0 To (GrpArray.Length - 1) Step 1
                    With GrpArray(i)
                        .TabIndex = ((254 - GrpArray.Length) + i) ' ���ٷ��ɂ��̫����ړ��ŕK�v
                        .Tag = i
                        .Text = GetPrivateProfileString_S( _
                            "USER_LABEL", (i.ToString("000") & "_GRP"), m_sPath, "????")
                    End With
                Next i

                ' ----------------------------------------------------------
                ' EDIT_DEF_User.ini�������قɕ\������ݒ�
                ' ----------------------------------------------------------
                'V2.1.0.0�@ CLbl_26�ǉ�
                LblArray = New cLbl_() { _
                    CLbl_0, CLbl_1, CLbl_2, _
                    CLbl_3, CLbl_4, CLbl_5, _
                    CLbl_6, CLbl_7, CLbl_8, _
                    CLbl_9, CLbl_10, CLbl_11, _
                    CLbl_12, CLbl_13, CLbl_14, _
                    CLbl_15, CLbl_16, CLbl_17, _
                    CLbl_18, CLbl_19, CLbl_20, _
                    CLbl_21, CLbl_26, CLbl_22 _
                }
                For i As Integer = 0 To (LblArray.Length - 1) Step 1
                    LblArray(i).Text = GetPrivateProfileString_S( _
                        "USER_LABEL", (i.ToString("000") & "_LBL"), m_sPath, "????")
                Next i

                ' ----------------------------------------------------------
                ' ���Ѹ�ٰ���ޯ�����̺��۰ق�ݒ�(CtlArray�̏��Ԃƍ��킹��)
                ' ----------------------------------------------------------
                'V2.1.0.0�B ���x�Z���T�[���ꌳ�Ǘ��I��ԍ��ǉ� CTxt_37,CTxt_38,�O��(CTxt_7)��Tab�I�[�_�[����iNo�̎��ɕύX
                m_CtlSystem = New Control() { _
                    CCmb_0, _
                    CTxt_0, CTxt_1, CTxt_2, CTxt_3, CCmb_1, CCmb_2, CTxt_4, CTxt_5, CTxt_6, CCmb_10, _
                    CCmb_3, CTxt_28, CTxt_30, CTxt_31, CTxt_37, CTxt_7, CTxt_8, CTxt_29, CTxt_38, CTxt_9, CTxt_10, _
                    CTxt_11, CTxt_12 _
                }
                'V2.0.0.0�J                m_CtlSystem = New Control() { _
                'V2.0.0.0�J                    CCmb_0, CTxt_0, CTxt_1, CTxt_2, CTxt_3, CCmb_1, CCmb_2, CTxt_4, CTxt_5, CTxt_6, CCmb_3, CCmb_4, _
                'V2.0.0.0�J                    CTxt_7, CTxt_8, CTxt_9, CTxt_10, CTxt_11, CTxt_12 _
                'V2.0.0.0�J                }
                Call SetControlData(m_CtlSystem)

                ' ----------------------------------------------------------
                ' ��R�ݒ�̒�R�ԍ��̺��۰ق�ݒ�(CtlArray�̏��Ԃƍ��킹��)
                ' ----------------------------------------------------------
                'V2.1.0.0�@ CTxt_32�`CTxt_36 ����p�ڕW�l�Z�o�W���ǉ�
                m_CtlResistor = New Control(,) { _
                    {CCmb_5, CTxt_13, CTxt_14, CTxt_32, CTxt_15}, _
                    {CCmb_6, CTxt_16, CTxt_17, CTxt_33, CTxt_18}, _
                    {CCmb_7, CTxt_19, CTxt_20, CTxt_34, CTxt_21}, _
                    {CCmb_8, CTxt_22, CTxt_23, CTxt_35, CTxt_24}, _
                    {CCmb_9, CTxt_25, CTxt_26, CTxt_36, CTxt_27} _
                }
                Call SetControlData(m_CtlResistor)

                'V2.1.0.0�@ CTxt_32�`CTxt_36 ����p�ڕW�l�Z�o�W���ǉ�
                m_DispResistor = New Control(,) { _
                    {Label1, CCmb_5, CTxt_13, CTxt_14, CTxt_32, CTxt_15}, _
                    {Label2, CCmb_6, CTxt_16, CTxt_17, CTxt_33, CTxt_18}, _
                    {Label3, CCmb_7, CTxt_19, CTxt_20, CTxt_34, CTxt_21}, _
                    {Label4, CCmb_8, CTxt_22, CTxt_23, CTxt_35, CTxt_24}, _
                    {Label5, CCmb_9, CTxt_25, CTxt_26, CTxt_36, CTxt_27} _
                }

                ' ----------------------------------------------------------
                ' ���ׂĂ�÷���ޯ��������ޯ������݂ɑ΂��A̫����ړ��ݒ�������Ȃ�
                ' �g�p���Ȃ����۰ق� Enabled=False �܂��� Visible=False �ɂ���
                ' ----------------------------------------------------------
                'V2.1.0.0�@ CTxt_32�`CTxt_36 ����p�ڕW�l�Z�o�W���ǉ�
                'V2.1.0.0�B ���x�Z���T�[���ꌳ�Ǘ��I��ԍ��ǉ� CTxt_37,CTxt_38,�O��(CTxt_7)��Tab�I�[�_�[����iNo�̎��ɕύX
                CtlArray = New Control() { _
                    CCmb_0, _
                    CTxt_0, CTxt_1, CTxt_2, CTxt_3, CCmb_1, CCmb_2, CTxt_4, CTxt_5, CTxt_6, CCmb_10, _
                    CCmb_3, CTxt_28, CTxt_30, CTxt_31, CTxt_37, CTxt_7, CTxt_8, CTxt_29, CTxt_38, CTxt_9, CTxt_10, _
                    CTxt_11, CTxt_12, _
                    CCmb_5, CTxt_13, CTxt_14, CTxt_32, CTxt_15, _
                    CCmb_6, CTxt_16, CTxt_17, CTxt_33, CTxt_18, _
                    CCmb_7, CTxt_19, CTxt_20, CTxt_34, CTxt_21, _
                    CCmb_8, CTxt_22, CTxt_23, CTxt_35, CTxt_24, _
                    CCmb_9, CTxt_25, CTxt_26, CTxt_36, CTxt_27 _
                }
                Call SetTabIndex(CtlArray) ' ��޲��ޯ����KeyDown����Ă�ݒ肷��

                ' ----------------------------------------------------------
                ' ��ʕ\������̫����������۰ق�ݒ肷��
                ' ----------------------------------------------------------
                FIRST_CONTROL = CtlArray(0)

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#Region "���������ɺ����ޯ���̐ݒ�������Ȃ�"
        ''' <summary>���������ɺ����ޯ���̐ݒ�������Ȃ�</summary>
        ''' <param name="cCombo">�ݒ�������Ȃ������ޯ��</param>
        Protected Overrides Sub InitComboBox(ByRef cCombo As cCmb_)
            Dim tag As Integer
            Try
                With cCombo
                    tag = DirectCast(.Tag, Integer)
                    .Items.Clear()
                    Select Case (DirectCast(.Parent.Tag, Integer)) ' ��ٰ���ޯ�������
                        ' ------------------------------------------------------------------------------
                        Case 0 To 3 ' ��ۯ�,ۯď��,���ʐݒ�,���x�ݻ���ٰ���ޯ��
                            Select Case (tag)
                                Case 0 ' ���i���
                                    .Items.Add("�w��Ȃ�")
                                    .Items.Add("���x�Z���T�[")
                                    .Items.Add("��R�g���~���O")
                                    .Items.Add("�`�b�v��R�g���~���O")    'V1.0.4.3�C
                                    .Items.Add("�`�b�v���x�Z���T�[") 'V2.0.0.0�@
                                    .Items.Add("�}�[�N��") 'V2.2.1.7�@
                                Case 1 ' �g���~���O���x
                                    .Items.Add("1:����")
                                    .Items.Add("2:�����x")
                                    .Items.Add("3:�ݒ�l")
                                Case 2 ' ���b�g�I������
                                    .Items.Add("0:�I���������薳��")
                                    .Items.Add("1:����")
                                    .Items.Add("2:���[�_�M��")
                                    .Items.Add("3:�������M��")
                                    'V2.0.0.0�M��
                                Case 3  ' �N�����v�Ƌz���̗L�薳��
                                    .Items.Add("�N�����v�z���L��")
                                    .Items.Add("�N�����v�̂�")
                                    .Items.Add("�z���̂�")
                                    'V2.0.0.0�M��
                                Case 4 ' ��R�P��
                                    .Items.Add("1:��")
                                    .Items.Add("2:�j��")
                                    'V2.0.0.0�J                                Case 4 ' �Q�Ɖ��x
                                    'V2.0.0.0�J                         .Items.Add("1:�O��")
                                    'V2.0.0.0�J              .Items.Add("2:�Q�T��")
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                        Case 4 ' �␳�l��ٰ���ޯ��
                            Select Case (tag)
                                Case 0 ' ��R�P��
                                    .Items.Add("1:��")
                                    .Items.Add("2:�j��")
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                        Case Else
                            Throw New Exception("Parent.Tag - Case Else")
                    End Select

                    '.SelectedIndex = 0
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

            Try
                tag = DirectCast(cTextBox.Tag, Integer)
                no = tag.ToString("000")

                Select Case (DirectCast(cTextBox.Parent.Tag, Integer)) ' ��ٰ���ޯ�������
                    ' ------------------------------------------------------------------------------
                    Case 0 To 3 ' ���Ѹ�ٰ���ޯ��
                        ' �ް������װ���̕\����
                        strMsg = GetPrivateProfileString_S("USER_VALUE", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0  ' ���[�U��
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "15")
                                strFlg = True
                            Case 1  ' ���[�U�@���b�g�m���D
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "15")
                                strFlg = True
                            Case 2  ' �p�^�[���m���D
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "15")
                                strFlg = True
                            Case 3  ' �v���O�����m���D
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "15")
                                strFlg = True
                            Case 4 ' ��������
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "99999")
                            Case 5  ' �␳�p�x
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "999")
                            Case 6  ' ����f�q��
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "9999")
                                'V2.0.0.0�J��
                            Case 7  ' �ݒ艷�x
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "100")
                            Case 8  ' ��\���l(ppm/��)
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "-9999.0000000")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "9999.0000000")
                            Case 9  ' ��\���l(ppm/��)
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "-9999.0000000")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "9999.0000000")
                            Case 10  ' ���x�Z���T�[���ꌳ�Ǘ��I��ԍ�
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "99")
                                Dim TempNo As Integer = Integer.Parse(strMax)
                                Dim TableNo As Integer = UserSub.TemperatureTableMaxNumberGet()
                                If TableNo < TempNo Then
                                    strMax = TableNo.ToString("0")
                                End If
                                'V2.1.0.0�B��
                            Case 11  ' �O���@'V2.1.0.0�B Case 8����11�ֈړ�
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "0.0000001")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "100000000.0000000")
                            Case 12  ' ���l(ppm/��)
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "-9999.0000000")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "9999.0000000")
                            Case 13  ' ���l(ppm/��)
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "-9999.0000000")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "9999.0000000")
                                'V2.0.0.0�J��
                                'V2.1.0.0�B��
                            Case 14  ' ���x�Z���T�[���ꌳ�Ǘ��I��ԍ�
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "99")
                                Dim TempNo As Integer = Integer.Parse(strMax)
                                Dim TableNo As Integer = UserSub.TemperatureTableMaxNumberGet()
                                If TableNo < TempNo Then
                                    strMax = TableNo.ToString("0")
                                End If
                                'V2.1.0.0�B��
                                'V2.0.0.0�J                            Case 7  ' �W����R�l
                                'V2.0.0.0�J                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "0.01")
                                'V2.0.0.0�J                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "100000000")
                                'V2.0.0.0�J                            Case 8  ' ��R���x�W��
                                'V2.0.0.0�J                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "0")
                                'V2.0.0.0�J                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "1")
                            Case Else
                                Throw New Exception("Case " & tag & ": Nothing")
                        End Select
                        ' ------------------------------------------------------------------------------
                    Case 4 ' �␳�l��ٰ���ޯ��
                        strMsg = GetPrivateProfileString_S("USER_VALUE", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0 To 3 ' ���[�U���@'V2.1.0.0�@  Case 0 To 2 ����  Case 0 To 3 �֕ύX
                                For i As Integer = 0 To (m_CtlResistor.GetLength(0) - 1) Step 1
                                    For j As Integer = 0 To (m_CtlResistor.GetLength(1) - 1) Step 1
                                        If (m_CtlResistor(i, j) Is cTextBox) Then
                                            tag = j + 18        'V2.1.0.0�B�@011_MSG = �m���D,014_MSG = �m���D���ǉ��ɂȂ����̂łQ�����B16��18�֕ύX
                                            no = tag.ToString("000")
                                            strMsg = GetPrivateProfileString_S("USER_VALUE", (no & "_MSG"), m_sPath, "??????")
                                            Select Case (j)
                                                Case 1 ' �␳�l 'V2.0.0.0�K �␳�l�̍��ڂ�ppm���͂ɕύX
                                                    strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "99999")
                                                    strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "-99999")
                                                Case 2 ' �ڕW�l�Z�o�W��                                                        
                                                    strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "0")
                                                    strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "9999.9")
                                                    'V2.1.0.0�@��
                                                Case 3 ' ����p�ڕW�l�Z�o�W��                                                        
                                                    strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "0")
                                                    strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "9999.9")
                                                    'V2.1.0.0�@��
                                                Case 4 ' ���葬�x��ύX����J�b�g�m���D                                        'V2.1.0.0�@ Case 3����Case 4�֕ύX
                                                    strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "1")
                                                    strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "99")
                                                Case Else
                                                    Throw New Exception("Case " & tag & ": Nothing")
                                            End Select
                                            Exit For
                                        End If
                                    Next j
                                Next i
                                'V2.1.0.0�B ���x�Z���T�[���ꌳ�Ǘ��I��ԍ��ǉ�Case�ԍ��Q�J�グ13��15
                            Case 15 ' �t�@�C�i�����~�b�g High[%]
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "-99.99")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "99.99")
                            Case 16 ' �t�@�C�i�����~�b�g Lo[%]
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "-99.99")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "99.99")
                            Case 17  ' ���Βl���~�b�g High[%]
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "-99.99")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "99.99")
                            Case 18  ' ���Βl���~�b�g Lo[%]
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "-99.99")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "99.99")
                            Case Else
                                Throw New Exception("Case " & tag & ": Nothing")
                        End Select

                    Case Else
                        Throw New Exception("Parent.Tag - Case Else")
                End Select

                With cTextBox
                    Call .SetStrMsg(strMsg) ' �ް������װ���̕\�����ݒ�
                    Call .SetMinMax(strMin, strMax) ' �����l�����l�̐ݒ�
                    Dim strKind As String
                    If (False = strFlg) Then ' (False=���l,True=������)
                        strKind = "�͈̔͂Ŏw�肵�ĉ�����"
                    Else
                        strKind = "�����͈̔͂Ŏw�肵�ĉ�����"
                        .MaxLength = Convert.ToInt32(strMax) ' SetControlData()���̏������f�Ŏg�p����
                        .TextAlign = HorizontalAlignment.Left
                    End If
                    Call .SetStrTip(strMin & "�`" & strMax & strKind) ' °�����ү���ނ̐ݒ�
                    Call .SetLblToolTip(m_MainEdit.LblToolTip) ' ҲݕҏW��ʂ�°����ߎQ�Ɛݒ�
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex, cTextBox.Name)
            End Try

        End Sub
#End Region

#Region "÷���ޯ��������ޯ���ɒl��\������"
        ''' <summary>÷���ޯ���ɒl��ݒ肷��</summary>
        Protected Overrides Sub SetDataToText()
            Try
                Call SetUserData()
                Me.Refresh()

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

#Region "���Ѹ�ٰ���ޯ�����̐ݒ�"
        ''' <summary>���Ѹ�ٰ���ޯ������÷���ޯ��������ޯ���ɒl��ݒ肷��</summary>
        Private Sub SetUserData()
            Try
                With m_MainEdit.W_stUserData
                    For i As Integer = 0 To (m_CtlSystem.Length - 1) Step 1
                        Select Case (i)
                            Case 0 ' ���i���
                                Dim cCombo As cCmb_ = DirectCast(m_CtlSystem(i), cCmb_)
                                Call NoEventIndexChange(cCombo, .iTrimType) ' �w�萻�i��ʂ�ݒ�
                            Case 1  ' �I�y���[�^��
                                m_CtlSystem(i).Text = .sOperator
                            Case 2  ' ���[�U���b�gNo.
                                m_CtlSystem(i).Text = .sLotNumber
                            Case 3  ' �p�^�[���m���D
                                m_CtlSystem(i).Text = .sPatternNo
                            Case 4  ' �v���O�����m���D
                                m_CtlSystem(i).Text = .sProgramNo
                            Case 5  ' �g���~���O���x
                                Dim cCombo As cCmb_ = DirectCast(m_CtlSystem(i), cCmb_)
                                Call NoEventIndexChange(cCombo, (.iTrimSpeed - 1))
                            Case 6  ' ���b�g�I������
                                Dim cCombo As cCmb_ = DirectCast(m_CtlSystem(i), cCmb_)
                                Call NoEventIndexChange(cCombo, .iLotChange)
                            Case 7  ' ���b�g��������
                                m_CtlSystem(i).Text = (.lLotEndSL).ToString()
                            Case 8  ' �J�b�g�ʒu�␳�p�x
                                m_CtlSystem(i).Text = (.lCutHosei).ToString()
                            Case 9  ' ���b�g�I��������f�q��
                                m_CtlSystem(i).Text = (.lPrintRes).ToString()
                            Case 10 ' �N�����v�Ƌz���̗L�薳�� 'V2.0.0.0�M
                                Dim cCombo As cCmb_ = DirectCast(m_CtlSystem(i), cCmb_)
                                Call NoEventIndexChange(cCombo, (.intClampVacume - 1))
                            Case 11 ' ���x�Z���T�[ ��R�����W 1:��, 2:K��
                                Dim cCombo As cCmb_ = DirectCast(m_CtlSystem(i), cCmb_)
                                Call NoEventIndexChange(cCombo, (.iTempResUnit - 1))
                                'V2.0.0.0�J��
                            Case 12  ' �ݒ艷�x
                                m_CtlSystem(i).Text = (.iTempTemp).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 13  ' ��\���l(ppm/��)
                                m_CtlSystem(i).Text = (.dDaihyouAlpha).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 14  ' ��\���l(ppm/��)
                                m_CtlSystem(i).Text = (.dDaihyouBeta).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                                'V2.1.0.0�B��
                            Case 15 ' ���x�Z���T�[���ꌳ�Ǘ��I��ԍ��ǉ� CTxt_37 �ȍ~Case�ԍ��P���Z
                                m_CtlSystem(i).Text = (.iTempSensorInfNoDaihyou).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                                'V2.1.0.0�B��
                            Case 16  ' �O�� 'V2.1.0.0�BCase 13����16�ֈړ�
                                m_CtlSystem(i).Text = (.dTemperatura0).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 17  ' ���l(ppm/��)
                                m_CtlSystem(i).Text = (.dAlpha).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 18  ' ���l(ppm/��)
                                m_CtlSystem(i).Text = (.dBeta).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                                'V2.1.0.0�B��
                            Case 19 ' ���x�Z���T�[���ꌳ�Ǘ��I��ԍ��ǉ� CTxt_38 �ȍ~Case�ԍ��P���Z
                                m_CtlSystem(i).Text = (.iTempSensorInfNoStd).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                                'V2.1.0.0�B��
                                'V2.0.0.0�J                            Case 11 ' �Q�Ɖ��x	�P�F�O�� �܂��� �Q�F�Q
                                'V2.0.0.0�J                                Dim cCombo As cCmb_ = DirectCast(m_CtlSystem(i), cCmb_)
                                'V2.0.0.0�J                                Call NoEventIndexChange(cCombo, (.iTempTemp - 1))
                                'V2.0.0.0�J                            Case 12 ' �W����R�l �O�� 0.01�`100M
                                'V2.0.0.0�J                                m_CtlSystem(i).Text = (.dStandardRes0).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                                'V2.0.0.0�J                            Case 13 ' �W����R�l �Q�T�� 0.01�`100M
                                'V2.0.0.0�J                                m_CtlSystem(i).Text = (.dStandardRes25).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 20 ' �t�@�C�i�����~�b�g�@Hight[%]
                                m_CtlSystem(i).Text = (.dFinalLimitHigh).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 21 ' �t�@�C�i�����~�b�g�@Lo[%]
                                m_CtlSystem(i).Text = (.dFinalLimitLow).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 22 ' ���Βl���~�b�g�@Hight[%]
                                m_CtlSystem(i).Text = (.dRelativeHigh).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 23 ' ���Βl���~�b�g�@Lo[%]
                                m_CtlSystem(i).Text = (.dRelativeLow).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case Else
                                Throw New Exception("Case " & i & ": Nothing")
                        End Select
                    Next i

                    For i As Integer = 0 To (m_CtlResistor.GetLength(0) - 1) Step 1
                        For j As Integer = 0 To (m_CtlResistor.GetLength(1) - 1) Step 1
                            Select Case (j)
                                Case 0 ' �␳�l ��R�����W 1:��, 2:K��
                                    Dim cCombo As cCmb_ = DirectCast(m_CtlResistor(i, j), cCmb_)
                                    Call NoEventIndexChange(cCombo, (.iResUnit(i + 1) - 1))
                                Case 1 ' �␳�l�i�m�~�i���l�Z�o�W���j
                                    m_CtlResistor(i, j).Text = (.dNomCalcCoff(i + 1)).ToString(DirectCast(m_CtlResistor(i, j), cTxt_).GetStrFormat())
                                Case 2 ' �ڕW�l�Z�o�W��
                                    m_CtlResistor(i, j).Text = (.dTargetCoff(i + 1)).ToString(DirectCast(m_CtlResistor(i, j), cTxt_).GetStrFormat())
                                    'V2.1.0.0�@��
                                Case 3 ' ����p�ڕW�l�Z�o�W��
                                    m_CtlResistor(i, j).Text = (.dTargetCoffJudge(i + 1)).ToString(DirectCast(m_CtlResistor(i, j), cTxt_).GetStrFormat())
                                    'V2.1.0.0�@��
                                Case 4 ' ���葬�x��ύX����J�b�g�m���D'V2.1.0.0�@�@Case 3����Case 4�֕ύX
                                    m_CtlResistor(i, j).Text = (.iChangeSpeed(i + 1)).ToString(DirectCast(m_CtlResistor(i, j), cTxt_).GetStrFormat())
                                Case Else
                                    Throw New Exception("Case " & Tag & ": Nothing")
                            End Select
                        Next j
                    Next i
                    Call setDispResistor()
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region
#End Region

#Region "��R���ɉ����ĺ��۰ق̕\��/��\����؂�ւ���"
        Private Sub setDispResistor()
            Dim ResCnt As Short   'V1.2.0.0�A

            'V1.2.0.0�A��
            ResCnt = UserBas.GetRCountExceptMeasure()
            'V2.0.0.0�I            If UserSub.IsTrimType3 Then
            'V2.0.0.0�I                ResCnt = 1                  ' �`�b�v�́A��R�P�̃f�[�^�ŏ�������B
            'V2.0.0.0�I            End If
            'V1.2.0.0�A��

            ' ��R�����̂ݕ\������
            For i As Integer = 0 To (m_DispResistor.GetLength(0) - 1) Step 1
                If (i < ResCnt) Then
                    For j As Integer = 0 To (m_DispResistor.GetLength(1) - 1) Step 1
                        m_DispResistor(i, j).Visible = True
                        m_DispResistor(i, j).Enabled = True
                    Next
                Else
                    For j As Integer = 0 To (m_DispResistor.GetLength(1) - 1) Step 1
                        m_DispResistor(i, j).Visible = False
                        m_DispResistor(i, j).Enabled = False
                    Next
                End If
            Next

        End Sub
#End Region

#Region "���ׂĂ�÷���ޯ�����ް������������Ȃ�"
        ''' <summary>���ׂĂ�÷���ޯ�����ް������������Ȃ�</summary>
        ''' <returns>0=����, 1=�װ</returns>
        Friend Overrides Function CheckAllTextData() As Integer
            Dim ret As Integer
            Try
                m_CheckFlg = True ' ������(tabBase_Layout�ɂĎg�p)
                m_MainEdit.MTab.SelectedIndex = m_TabIdx ' ��ޕ\���ؑ�

                ' ���������ް�����۰قɾ�Ă���
                Call SetDataToText()

                ' ���Ѹ�ٰ���ޯ��
                ret = CheckControlData(m_CtlSystem)
                If (ret <> 0) Then Exit Try

                ' �␳�l��ٰ���ޯ��
                ret = CheckControlData(m_CtlResistor)
                If (ret <> 0) Then Exit Try

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
                tag = DirectCast(cTextBox.Tag, Integer)

                With m_MainEdit
                    Select Case (DirectCast(cTextBox.Parent.Tag, Integer))
                        ' ------------------------------------------------------------------------------

                        Case 1 To 3 ' ���Ѹ�ٰ���ޯ��
                            With .W_stUserData
                                Select Case (tag)
                                    Case 0 ' �I�y���[�^��
                                        ret = CheckStrData(cTextBox, .sOperator)
                                    Case 1 ' ���[�U�@���b�g�ԍ�
                                        ret = CheckStrData(cTextBox, .sLotNumber)
                                    Case 2 ' �p�^�[���m���D
                                        ret = CheckStrData(cTextBox, .sPatternNo)
                                    Case 3 ' �v���O�����m���D
                                        ret = CheckStrData(cTextBox, .sProgramNo)
                                    Case 4 ' ��������
                                        ret = CheckLongData(cTextBox, .lLotEndSL)
                                    Case 5 ' �J�b�g�ʒu�␳�p�x
                                        ret = CheckLongData(cTextBox, .lCutHosei)
                                    Case 6 ' ���b�g�I��������f�q��
                                        ret = CheckLongData(cTextBox, .lPrintRes)
                                        'V2.0.0.0�J                                    Case 7 ' �W����R�l	�O��0.01�`100M
                                        'V2.0.0.0�J                                        ret = CheckDoubleData(cTextBox, .dStandardRes0)
                                        'V2.0.0.0�J                                        If ret = 0 Then
                                        'V2.0.0.0�J                                            .dResTempCoff = GetResTempCoff(.dStandardRes0, .dStandardRes25)
                                        'V2.0.0.0�J                                            LabelTempCoff.Text = (.dResTempCoff).ToString("0.000")
                                        'V2.0.0.0�J                                        End If
                                        'V2.0.0.0�J                                    Case 8 ' �W����R�l	�Q�T�� 0.01�`100M
                                        'V2.0.0.0�J                                        ret = CheckDoubleData(cTextBox, .dStandardRes25)
                                        'V2.0.0.0�J                                        If ret = 0 Then
                                        'V2.0.0.0�J                                            .dResTempCoff = GetResTempCoff(.dStandardRes0, .dStandardRes25)
                                        'V2.0.0.0�J                                            LabelTempCoff.Text = (.dResTempCoff).ToString("0.000")
                                        'V2.0.0.0�J                                        End If
                                    Case 7  ' �ݒ艷�x
                                        ret = CheckIntData(cTextBox, .iTempTemp)
                                    Case 8  ' ��\���l(ppm/��)
                                        ret = CheckDoubleData(cTextBox, .dDaihyouAlpha)
                                    Case 9  ' ��\���l(ppm/��)
                                        ret = CheckDoubleData(cTextBox, .dDaihyouBeta)
                                        'V2.1.0.0�B��
                                    Case 10  ' ��\���x�Z���T�[���ꌳ�Ǘ��I��ԍ�
                                        Dim iSaveNo As Integer = .iTempSensorInfNoDaihyou
                                        ret = CheckIntData(cTextBox, .iTempSensorInfNoDaihyou)
                                        If ret = 0 And .iTempSensorInfNoDaihyou > 0 Then ' ���x���X�V
                                            Dim dDummy As Double
                                            If TemperatureTableDataGet(.iTempSensorInfNoDaihyou, dDummy, .dDaihyouAlpha, .dDaihyouBeta) Then
                                                Call SetDataToText()
                                            Else
                                                Call MsgBox_CheckErr(cTextBox, "���x�Z���T�[��񂪎擾�ł��܂���ł����BNo=[" & .iTempSensorInfNoDaihyou.ToString("0") & "]", iSaveNo.ToString())
                                                .iTempSensorInfNoDaihyou = iSaveNo
                                            End If
                                        End If
                                        'V2.1.0.0�B��
                                    Case 11  ' �O��   'V2.1.0.0�B Case8����11�ֈړ�
                                        ret = CheckDoubleData(cTextBox, .dTemperatura0)
                                    Case 12  ' ���l(ppm/��)
                                        ret = CheckDoubleData(cTextBox, .dAlpha)
                                    Case 13  ' ���l(ppm/��)
                                        ret = CheckDoubleData(cTextBox, .dBeta)
                                        'V2.1.0.0�B��
                                    Case 14  ' ���x�Z���T�[���ꌳ�Ǘ��I��ԍ�
                                        Dim iSaveNo As Integer = .iTempSensorInfNoStd
                                        ret = CheckIntData(cTextBox, .iTempSensorInfNoStd)
                                        If ret = 0 And .iTempSensorInfNoStd > 0 Then ' ���x���X�V
                                            If TemperatureTableDataGet(.iTempSensorInfNoStd, .dTemperatura0, .dAlpha, .dBeta) Then
                                                Call SetDataToText()
                                            Else
                                                Call MsgBox_CheckErr(cTextBox, "���x�Z���T�[��񂪎擾�ł��܂���ł����BNo=[" & .iTempSensorInfNoStd.ToString("0") & "]", iSaveNo.ToString())
                                                .iTempSensorInfNoStd = iSaveNo
                                            End If
                                        End If
                                        'V2.1.0.0�B��
                                    Case Else
                                        Throw New Exception("Case " & tag & ": Nothing")
                                End Select
                            End With
                            ' ------------------------------------------------------------------------------
                        Case 4
                            With .W_stUserData
                                Select Case (tag)
                                    Case 0 To 3 'V2.1.0.0�@2����3�֕ύX ' �␳�l���瑪�葬�x��ύX����J�b�g�m���D�܂ŁiSetControlData�Őݒ肳���e�L�X�g�{�b�N�X�̏��ԁA�S�ȍ~�Ƃ̓Z�b�g�����R���g���[�����قȂ�j
                                        Dim bStop As Boolean
                                        bStop = False
                                        For i As Integer = 0 To (m_CtlResistor.GetLength(0) - 1) Step 1
                                            For j As Integer = 0 To (m_CtlResistor.GetLength(1) - 1) Step 1
                                                If (m_CtlResistor(i, j) Is cTextBox) Then
                                                    Select Case (j)
                                                        Case 1 ' �␳�l
                                                            ret = CheckDoubleData(cTextBox, .dNomCalcCoff(i + 1))
                                                        Case 2 ' �ڕW�l�Z�o�W��                                                        
                                                            ret = CheckDoubleData(cTextBox, .dTargetCoff(i + 1))
                                                            'V2.1.0.0�@��
                                                        Case 3 ' ����p�ڕW�l�Z�o�W��                                                        
                                                            ret = CheckDoubleData(cTextBox, .dTargetCoffJudge(i + 1))
                                                            'V2.1.0.0�@��
                                                        Case 4 ' ���葬�x��ύX����J�b�g�m���D                'V2.1.0.0�@ Case 3����Case 4�֕ύX                        
                                                            ret = CheckIntData(cTextBox, .iChangeSpeed(i + 1))
                                                        Case Else
                                                            Throw New Exception("Case " & tag & ": Nothing")
                                                    End Select
                                                    bStop = True
                                                    Exit For
                                                End If
                                            Next j
                                            If bStop Then
                                                Exit For
                                            End If
                                        Next i
                                        'V2.1.0.0�B No.�̃e�L�X�g���ڂ��Q�ǉ������̂ňȍ~�Q�����Z 13��15
                                    Case 15 ' �t�@�C�i�����~�b�g�@Hight[%]
                                        ret = CheckDoubleData(cTextBox, .dFinalLimitHigh)
                                    Case 16 ' �t�@�C�i�����~�b�g�@Lo[%]
                                        ret = CheckDoubleData(cTextBox, .dFinalLimitLow)
                                    Case 17 ' ���Βl���~�b�g�@Hight[%]
                                        ret = CheckDoubleData(cTextBox, .dRelativeHigh)
                                    Case 18 ' ���Βl���~�b�g�@Lo[%]
                                        ret = CheckDoubleData(cTextBox, .dRelativeLow)
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
                Call MsgBox_Exception(ex, cTextBox.Name)
                ret = 1
            Finally
                CheckTextData = ret
            End Try

        End Function
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
                        Case 0 To 3
                            Select Case (tag)
                                Case 0 ' ���i���( 0:���x�Z���T�[, 1:��R�g���~���O )
                                    .W_stUserData.iTrimType = Convert.ToInt16(idx)
                                Case 1 ' �g���~���O���x(1:����, 2:�����x�C3:�ݒ�l)
                                    .W_stUserData.iTrimSpeed = Convert.ToInt16(idx + 1)
                                Case 2 ' ���b�g�I������(0:�I���������薳��,1:����,2:���[�_�M��,3:�������M��)
                                    .W_stUserData.iLotChange = Convert.ToInt16(idx)
                                Case 3 ' �N�����v�Ƌz���̗L�薳��                               'V2.0.0.0�M
                                    .W_stUserData.intClampVacume = Convert.ToInt16(idx + 1)     'V2.0.0.0�M
                                Case 4 ' ��R�����W(1:��, 2:K��)
                                    .W_stUserData.iTempResUnit = Convert.ToInt16(idx + 1)
                                    'V2.0.0.0�J                                Case 4 ' �Q�Ɖ��x(�P�F�O�� �܂��� �Q�F�Q�T��)
                                    'V2.0.0.0�J                                    .W_stUserData.iTempTemp = Convert.ToInt16(idx + 1)
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                        Case 4 ' ��R�ݒ�
                            For i As Integer = 0 To (m_CtlResistor.GetLength(0) - 1) Step 1
                                If (m_CtlResistor(i, tag) Is cCombo) Then
                                    Debug.WriteLine("DATA=" & i)
                                    .W_stUserData.iResUnit(i + 1) = Convert.ToInt16(idx + 1)
                                    Exit For
                                End If
                            Next i
                            ' ------------------------------------------------------------------------------
                        Case Else
                            Throw New Exception("Parent.Tag - Case Else")
                    End Select
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

    End Class
End Namespace
