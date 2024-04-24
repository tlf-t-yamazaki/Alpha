Option Explicit On
Option Strict Off

Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports LaserFront.Trimmer.DefWin32Fnc

Namespace FormEdit
    Friend Class tabSystem
        Inherits tabBase

#Region "�錾"
        Private Const SYS_ZON As Integer = 11       ' ���������Ŏg�p(m_CtlSystem�ł̲��ޯ��)
        Private Const SYS_ZOFF As Integer = 12      ' ���������Ŏg�p(m_CtlSystem�ł̲��ޯ��)

        Private m_CtlSystem() As Control            ' ���Ѹ�ٰ���ޯ���̺��۰ٔz��
        Private m_CtlLaser() As Control             ' ڰ�ް��ܰ������ٰ���ޯ���̺��۰ٔz��
        Private m_CtlTeachBlock() As Control        ' ###1040 �e�B�[�`���O�E�u���b�N��ٰ���ޯ���̺��۰ٔz��
        Private m_CtlStageSpeed() As Control        ' ###1040 �X�e�[�W���x��ٰ���ޯ���̺��۰ٔz��
        Private m_CtlDisMagnify() As Control        ' 'V2.2.0.0�A �f�W�^���J�����\���{�� ��ٰ���ޯ���̺��۰ٔz��

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
                TAB_NAME = GetPrivateProfileString_S("SYSTEM_LABEL", "TAB_NAM", m_sPath, "????")

                ' ���U���ʂ�̧���ڰ�ނ̏ꍇ��ڰ�ް��ܰ������ٰ���ޯ�����\���ɂ���
                If (gSysPrm.stRAT.giOsc_Res = OSCILLATOR_FL) Then
                    CGrp_5.Visible = False
                End If

                ' ----------------------------------------------------------
                ' EDIT_DEF_User.ini�����ٰ���ޯ���ɕ\������ݒ� ###1040 Grp_6, CGrp_7�ǉ�     ' V2.2.0.0�A CGrp_8�ǉ�
                ' ----------------------------------------------------------
                'V2.2.2.0�@ ��
                If giLoaderType <> 0 Then
                    '#0128
                    GrpArray = New cGrp_() {
                    CGrp_0, CGrp_1, CGrp_2, CGrp_3, CGrp_4,
                    CGrp_5, CGrp_6, CGrp_7, CGrp_8
                    }
                Else
                    '#0005,#0050
                    GrpArray = New cGrp_() {
                    CGrp_0, CGrp_1, CGrp_2, CGrp_3, CGrp_4,
                    CGrp_5, CGrp_6, CGrp_7
                    }
                    CGrp_8.Visible = False
                End If
                'V2.2.2.0�@ ��
                For i As Integer = 0 To (GrpArray.Length - 1) Step 1
                    With GrpArray(i)
                        .TabIndex = ((254 - GrpArray.Length) + i) ' ���ٷ��ɂ��̫����ړ��ŕK�v
                        .Tag = 0
                        .Text = GetPrivateProfileString_S(
                            "SYSTEM_LABEL", (i.ToString("000") & "_GRP"), m_sPath, "????")
                    End With
                Next i
                ' ��ٰ���ޯ��������q�ɂȂ��Ă��邽��
                ' ���тƼ��ѓ�=0, ڰ����ܰ����=1�Ƃ���
                CGrp_5.Tag = 1
                CGrp_6.Tag = 2      ' ###1040
                CGrp_7.Tag = 3      ' ###1040
                CGrp_8.Tag = 4      ' V2.2.0.0�A 

                ' �ǉ���폜���݂����� (���݂Ȃ�)
                'CPnl_Btn.TabIndex = 254 ' ���۰ٔz�u�\�ő吔(�Ō�ɐݒ�)

                ' ----------------------------------------------------------
                ' EDIT_DEF_User.ini�������قɕ\������ݒ� ###1040 CLbl_12, CLbl_13, CLbl_14 �ǉ�
                ' 'V1.2.0.0�@ CLbl_15�`�b�v�T�C�Y�ǉ�  'V2.2.0.0�A CLbl_17�ǉ� 
                ' 'V2.2.0.0�N CLbl_18��۰�ނ̒ǉ�       
                ' ----------------------------------------------------------
                LblArray = New cLbl_() {
                    CLbl_0, CLbl_1, CLbl_2,
                    CLbl_3, CLbl_4, CLbl_5,
                    CLbl_6, CLbl_7, CLbl_8,
                    CLbl_9,
                    CLbl_10, CLbl_11, CLbl_12, CLbl_13, CLbl_14, CLbl_15, CLbl_17, CLbl_18
                }
                For i As Integer = 0 To (LblArray.Length - 1) Step 1
                    LblArray(i).Text = GetPrivateProfileString_S( _
                        "SYSTEM_LABEL", (i.ToString("000") & "_LBL"), m_sPath, "????")
                Next i

                ' ----------------------------------------------------------
                ' ���Ѹ�ٰ���ޯ�����̺��۰ق�ݒ�(CtlArray�̏��Ԃƍ��킹��)
                ' V1.2.0.0�@ �`�b�v�T�C�Y CTxt_20,CTxt_21�ǉ�
                ' V2.2.0.0�N CTxt_24�F�v���[�u�̒ǉ�
                ' ----------------------------------------------------------
                m_CtlSystem = New Control() {
                    CTxt_0, CTxt_1, CTxt_2, CTxt_3, CTxt_4,
                    CTxt_5, CTxt_6, CTxt_7, CTxt_8, CTxt_9, CTxt_10,
                    CTxt_11, CTxt_12, CCmb_0,
                    CTxt_13,
                    CTxt_20, CTxt_21,
                     CTxt_24
                }
                Call SetControlData(m_CtlSystem)

                ' ----------------------------------------------------------
                ' ڰ�ް��ܰ������ٰ���ޯ���̺��۰ق�ݒ�(CtlArray�̏��Ԃƍ��킹��)
                ' ----------------------------------------------------------
                'V2.1.0.0�A�@CTxt_22(�A�b�e�l�[�^No.)�ǉ��@
                m_CtlLaser = New Control() { _
                    CTxt_14, CTxt_15, CTxt_16, CTxt_22 _
                }
                Call SetControlData(m_CtlLaser)

                ' ----------------------------------------------------------------------------------------
                ' ###1040 �e�B�[�`���O�E�u���b�N��ٰ���ޯ���̺��۰ق�ݒ�(CtlArray�̏��Ԃƍ��킹��)
                ' ----------------------------------------------------------------------------------------
                m_CtlTeachBlock = New Control() { _
                    CTxt_17, CTxt_18 _
                }
                Call SetControlData(m_CtlTeachBlock)

                ' ----------------------------------------------------------------------------------------
                ' ###1040 �X�e�[�W���x��ٰ���ޯ���̺��۰ق�ݒ�(CtlArray�̏��Ԃƍ��킹��)
                ' ----------------------------------------------------------------------------------------
                m_CtlStageSpeed = New Control() { _
                    CTxt_19 _
                }
                Call SetControlData(m_CtlStageSpeed)
                ' ----------------------------------------------------------------------------------------
                ' ' 'V2.2.0.0�A �i���ٰ���ޯ���̺��۰ق�ݒ�(CtlArray�̏��Ԃƍ��킹��)
                ' ----------------------------------------------------------------------------------------
                m_CtlDisMagnify = New Control() {
                    CTxt_23
                }
                Call SetControlData(m_CtlDisMagnify)

                ' ----------------------------------------------------------
                ' ���ׂĂ�÷���ޯ��������ޯ������݂ɑ΂��A̫����ړ��ݒ�������Ȃ�
                ' �g�p���Ȃ����۰ق� Enabled=False �܂��� Visible=False �ɂ���
                ' ###1040 CTxt_16�`19�܂Œǉ�
                ' V1.2.0.0�@ �`�b�v�T�C�Y CTxt_20,CTxt_21�ǉ�
                ' V2.1.0.0�A�@CTxt_22(�A�b�e�l�[�^No.)�ǉ��@
                ' 'V2.2.0.0�A CTxt_23:�f�W�^���J�����\���{���ǉ� 
                ' 'V2.2.0.0�N CTxt_24�F�v���[�uNo�ǉ�
                ' ---------------------------------------------------------- 
                CtlArray = New Control() {
                    CTxt_0, CTxt_1, CTxt_2, CTxt_3, CTxt_4, CTxt_20, CTxt_21,
                    CTxt_5, CTxt_6, CTxt_7, CTxt_8, CTxt_9, CTxt_10,
                    CTxt_11, CTxt_12, CCmb_0, CTxt_24,
                    CTxt_13,
                    CTxt_14, CTxt_15, CTxt_16, CTxt_22, CTxt_17, CTxt_18, CTxt_19, CTxt_23
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
                        Case 0 ' ���Ѹ�ٰ���ޯ��
                            Select Case (tag)
                                Case 0 ' ��۰����ײ
                                    .Items.Add("�Ȃ�")
                                    .Items.Add("����")
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case Else
                            Throw New Exception("Parent.Tag - Case Else")
                    End Select
                    .SelectedIndex = 0

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
            Try
                tag = DirectCast(cTextBox.Tag, Integer)
                no = tag.ToString("000")

                Select Case (DirectCast(cTextBox.Parent.Tag, Integer)) ' ��ٰ���ޯ�������
                    ' ------------------------------------------------------------------------------
                    Case 0 ' ���Ѹ�ٰ���ޯ��
                        ' �ް������װ���̕\����
                        strMsg = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0  ' ��ۯ�����X
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "0.001")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "80.0")
                            Case 1  ' ��ۯ����ނx
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "0.001")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "80.0")
                            Case 2  ' ��ۯ���X
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "50")
                            Case 3  ' ��ۯ���Y
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "50")
                            Case 4 ' ��R��
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "20")
                                If Integer.Parse(strMax) > MAXRNO Then
                                    strMax = MAXRNO.ToString
                                End If
                            Case 5  ' ð��وʒu�̾��X
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "-245.000")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "245.000")
                            Case 6  ' ð��وʒu�̾��Y
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "-245.000")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "245.000")
                            Case 7  ' �ްшʒu�̾��X
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "-80.000")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "80.000")
                            Case 8  ' �ްшʒu�̾��Y
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "-80.000")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "80.000")
                                'V2.0.0.0�@                            Case 9 ' ��ެ���߲��X
                                'V2.0.0.0�@                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "0.001")
                                'V2.0.0.0�@                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "80.000")
                                'V2.0.0.0�@                            Case 10 ' ��ެ���߲��Y
                                'V2.0.0.0�@                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "0.001")
                                'V2.0.0.0�@                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "80.000")
                                'V2.0.0.0�@��ެ���߲��X����X�e�b�v�I�t�Z�b�gX�֕ύX��
                            Case 9 ' �X�e�b�v�I�t�Z�b�gX 
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "-80.000")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "80.000")
                            Case 10 ' �X�e�b�v�I�t�Z�b�gY
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "-80.000")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "80.000")
                                'V2.0.0.0�@��
                            Case 11  ' ��۰�ސڐG�ʒu
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "0.000")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "40.000")
                            Case 12  ' ��۰�ޑҋ@�ʒu
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "0.000")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "40.000")
                            Case 13 ' �O���@�퐔
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "10")
                                'V1.2.0.0�@��
                            Case 14  ' �`�b�v�T�C�YX
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "0.0001")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "50.0")
                            Case 15  ' �`�b�v�T�C�YY
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "0.0001")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "50.0")
                                'V1.2.0.0�@��
                                'V2.2.0.0�N ��
                            Case 16
                                ' �v���[�u�ԍ� 
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "10")
                                'V2.2.0.0�N ��
                            Case Else
                                Throw New Exception("Case " & tag & ": Nothing")
                        End Select
                        ' ------------------------------------------------------------------------------
                    Case 1 ' ڰ�ް��ܰ������ٰ���ޯ��
                        If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then ' ̧���ڰ�ް�łȂ��ꍇ
                            ' �ް������װ���̕\����
                            strMsg = GetPrivateProfileString_S("SYSTEM_LASER", (no & "_MSG"), m_sPath, "??????")
                            Select Case (tag)
                                Case 0 ' Qڰ�
                                    strMin = GetPrivateProfileString_S("SYSTEM_LASER", (no & "_MIN"), m_sPath, "0.1")
                                    strMax = GetPrivateProfileString_S("SYSTEM_LASER", (no & "_MAX"), m_sPath, "40.0")
                                Case 1 ' �ݒ���ܰ
                                    strMin = GetPrivateProfileString_S("SYSTEM_LASER", (no & "_MIN"), m_sPath, "0.1")
                                    strMax = GetPrivateProfileString_S("SYSTEM_LASER", (no & "_MAX"), m_sPath, "10.0")
                                Case 2 ' �A�b�e�l�[�^�������i0:�ۑ��� 1:�ۑ��j###1040�B
                                    strMin = GetPrivateProfileString_S("SYSTEM_LASER", (no & "_MIN"), m_sPath, "0")
                                    strMax = GetPrivateProfileString_S("SYSTEM_LASER", (no & "_MAX"), m_sPath, "1")
                                Case 3  'V2.1.0.0�A��CTxt_22(�A�b�e�l�[�^No.)�ǉ�
                                    strMin = GetPrivateProfileString_S("SYSTEM_LASER", (no & "_MIN"), m_sPath, "0")
                                    strMax = GetPrivateProfileString_S("SYSTEM_LASER", (no & "_MAX"), m_sPath, "99")
                                    Dim TempNo As Integer = Integer.Parse(strMax)
                                    Dim TableNo As Integer = UserSub.LaserCalibrationMaxNumberGet()
                                    If TableNo < TempNo Then
                                        strMax = TableNo.ToString("0")
                                    End If
                                    'V2.1.0.0�A��
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                        End If
                        ' ------------------------------------------------------------------------------
                    Case 2 ' �e�B�[�`���O�E�u���b�N��ٰ���ޯ�� ###1040�@
                        strMsg = GetPrivateProfileString_S("SYSTEM_TEACHBLOCK", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0  ' �e�B�[�`���O�E�u���b�N�w
                                strMin = GetPrivateProfileString_S("SYSTEM_TEACHBLOCK", (no & "_MIN"), m_sPath, "1")
                                strMax = m_MainEdit.W_PLT.BNX
                            Case 1  ' �e�B�[�`���O�E�u���b�N�x
                                strMin = GetPrivateProfileString_S("SYSTEM_TEACHBLOCK", (no & "_MIN"), m_sPath, "1")
                                strMax = m_MainEdit.W_PLT.BNY
                            Case Else
                                Throw New Exception("Case " & tag & ": Nothing")
                        End Select
                    Case 3 ' �X�e�[�W���x��ٰ���ޯ�� ###1040�C
                        strMsg = GetPrivateProfileString_S("SYSTEM_STAGESPEED", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0  ' �x��
                                strMin = GetPrivateProfileString_S("SYSTEM_STAGESPEED", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("SYSTEM_STAGESPEED", (no & "_MAX"), m_sPath, "50")
                            Case Else
                                Throw New Exception("Case " & tag & ": Nothing")
                        End Select
                    Case 4  '�i��O���[�v�{�b�N�X         'V2.2.0.0�A
                        strMsg = GetPrivateProfileString_S("SYSTEM_KIND", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0  ' �\���{�� 
                                strMin = GetPrivateProfileString_S("SYSTEM_KIND", (no & "_MIN"), m_sPath, "0.5")
                                strMax = GetPrivateProfileString_S("SYSTEM_KIND", (no & "_MAX"), m_sPath, "2.0")
                            Case Else
                                Throw New Exception("Case " & tag & ": Nothing")
                        End Select

                    Case Else
                        Throw New Exception("Parent.Tag - Case Else")
                End Select

                With cTextBox
                    Call .SetStrMsg(strMsg) ' �ް������װ���̕\�����ݒ�
                    Call .SetMinMax(strMin, strMax) ' �����l�����l�̐ݒ�
                    Call .SetStrTip(strMin & "�`" & strMax & "�͈̔͂Ŏw�肵�ĉ�����") ' °�����ү���ނ̐ݒ�
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
                ' ���Ѹ�ٰ���ޯ���ݒ�
                Call SetSystemData()

                ' ڰ����ܰ������ٰ���ޯ���ݒ�
                Call SetLaserData()

                Call SetTeachBlockData()    ' ###1040 �e�B�[�`���O�E�u���b�N

                Call SetStageSpeedData()    ' ###1040 �X�e�[�W���x

                Call SetKindData()          ' 'V2.2.0.0�A �\���{�� 

                Me.Refresh()

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

#Region "���Ѹ�ٰ���ޯ�����̐ݒ�"
        ''' <summary>���Ѹ�ٰ���ޯ������÷���ޯ��������ޯ���ɒl��ݒ肷��</summary>
        Private Sub SetSystemData()
            Try
                With m_MainEdit.W_PLT
                    For i As Integer = 0 To (m_CtlSystem.Length - 1) Step 1
                        Select Case (i)
                            Case 0 ' ��ۯ�����X(mm)
                                m_CtlSystem(i).Text = (.zsx).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 1 ' ��ۯ�����Y(mm)
                                m_CtlSystem(i).Text = (.zsy).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 2 ' ��ۯ���X
                                m_CtlSystem(i).Text = (.BNX).ToString()
                            Case 3 ' ��ۯ���Y
                                m_CtlSystem(i).Text = (.BNY).ToString()
                            Case 4 ' ��R��
                                m_CtlSystem(i).Text = (.RCount).ToString()
                            Case 5  ' ð��وʒu�̾��X(mm)
                                m_CtlSystem(i).Text = (.z_xoff).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 6 ' ð��وʒu�̾��Y(mm)
                                m_CtlSystem(i).Text = (.z_yoff).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 7 ' �ްшʒu�̾��X(mm)
                                m_CtlSystem(i).Text = (.BPOX).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 8 ' �ްшʒu�̾��Y(mm)
                                m_CtlSystem(i).Text = (.BPOY).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                                'V2.0.0.0�@                            Case 9 ' ��ެ���߲��X
                                'V2.0.0.0�@                                m_CtlSystem(i).Text = (.ADJX).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                                'V2.0.0.0�@                            Case 10 ' ��ެ���߲��Y
                                'V2.0.0.0�@                                m_CtlSystem(i).Text = (.ADJY).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                                'V2.0.0.0�@��
                            Case 9 ' �X�e�b�v�I�t�Z�b�g��X
                                m_CtlSystem(i).Text = (.dblStepOffsetXDir).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 10 ' �X�e�b�v�I�t�Z�b�g��Y
                                m_CtlSystem(i).Text = (.dblStepOffsetYDir).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                                'V2.0.0.0�@��
                            Case 11  ' ��۰�ސڐG�ʒu(mm)
                                m_CtlSystem(i).Text = (.Z_ZON).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 12 ' ��۰�ޑҋ@�ʒu(mm)
                                m_CtlSystem(i).Text = (.Z_ZOFF).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 13 ' ��۰����ײ(0:��, 1:�L)
                                Call NoEventIndexChange(DirectCast(m_CtlSystem(i), cCmb_), .PrbRetry)
                            Case 14 ' �O���@�퐔
                                m_CtlSystem(i).Text = (.GCount).ToString()
                                'V1.2.0.0�@��
                            Case 15 ' �`�b�v�T�C�YX(mm)
                                m_CtlSystem(i).Text = (.dblChipSizeXDir).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 16 ' �`�b�v�T�C�YY(mm)
                                m_CtlSystem(i).Text = (.dblChipSizeYDir).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                                'V1.2.0.0�@��
                            '    'V2.2.0.0�A��
                            'Case 17 ' �\���{��
                            '    m_CtlSystem(i).Text = (.dblStdMagnification).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            '    'V2.2.0.0�A��
                                'V2.2.0.0�N ��
                            Case 17
                                m_CtlSystem(i).Text = (.ProbNo).ToString()
                                'V2.2.0.0�N ��
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

#Region "ڰ����ܰ������ٰ���ޯ�����̐ݒ�"
        ''' <summary>ڰ����ܰ������ٰ���ޯ������÷���ޯ��������ޯ���ɒl��ݒ肷��</summary>
        Private Sub SetLaserData()
            Try
                With m_MainEdit.W_LASER
                    For i As Integer = 0 To (m_CtlLaser.Length - 1) Step 1
                        Select Case (i)
                            Case 0 ' Qڰ�
                                m_CtlLaser(i).Text = (.intQR / 10).ToString(DirectCast(m_CtlLaser(i), cTxt_).GetStrFormat())
                            Case 1 ' �ݒ���ܰ(W)
                                m_CtlLaser(i).Text = (.dblspecPower).ToString(DirectCast(m_CtlLaser(i), cTxt_).GetStrFormat())
                            Case 2 ' �A�b�e�l�[�^�������i0:�ۑ��� 1:�ۑ��j###1040
                                m_CtlLaser(i).Text = (.iTrimAtt).ToString(DirectCast(m_CtlLaser(i), cTxt_).GetStrFormat())
                            Case 3  'V2.1.0.0�A�@�A�b�e�l�[�^No.
                                m_CtlLaser(i).Text = (.iAttNo).ToString(DirectCast(m_CtlLaser(i), cTxt_).GetStrFormat())
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
#Region "�e�B�[�`���O�E�u���b�N��ٰ���ޯ�����̐ݒ�"
        ''' <summary>�e�B�[�`���O�E�u���b�N��ٰ���ޯ������÷���ޯ��������ޯ���ɒl��ݒ肷��</summary>
        Private Sub SetTeachBlockData()
            Try
                With m_MainEdit.W_PLT
                    For i As Integer = 0 To (m_CtlTeachBlock.Length - 1) Step 1
                        Select Case (i)
                            Case 0 ' �u���b�N�w
                                m_CtlTeachBlock(i).Text = (.TeachBlockX).ToString(DirectCast(m_CtlTeachBlock(i), cTxt_).GetStrFormat())
                            Case 1 ' �u���b�N�x
                                m_CtlTeachBlock(i).Text = (.TeachBlockY).ToString(DirectCast(m_CtlTeachBlock(i), cTxt_).GetStrFormat())
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
#Region "�X�e�[�W���x��ٰ���ޯ�����̐ݒ�"
        ''' <summary>�X�e�[�W���x��ٰ���ޯ������÷���ޯ��������ޯ���ɒl��ݒ肷��</summary>
        Private Sub SetStageSpeedData()
            Try
                With m_MainEdit.W_PLT
                    For i As Integer = 0 To (m_CtlStageSpeed.Length - 1) Step 1
                        Select Case (i)
                            Case 0 ' �x�����x
                                m_CtlStageSpeed(i).Text = (.StageSpeedY).ToString(DirectCast(m_CtlStageSpeed(i), cTxt_).GetStrFormat())
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
#Region "�i���ٰ���ޯ�����̐ݒ�"
        Private Sub SetKindData()
            Try
                With m_MainEdit.W_PLT
                    For i As Integer = 0 To (m_CtlDisMagnify.Length - 1) Step 1
                        Select Case (i)
                            Case 0 ' �\���{��
                                m_CtlDisMagnify(i).Text = (.dblStdMagnification).ToString(DirectCast(m_CtlDisMagnify(i), cTxt_).GetStrFormat())
                            Case Else
                                Throw New Exception("Case " & i & ": Nothing")
                        End Select
                    Next i
                End With
            Catch ex As Exception

            End Try
        End Sub
#End Region
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

                ' �e�B�[�`���O�E�u���b�N��ٰ���ޯ��
                ret = CheckControlData(m_CtlTeachBlock) ' ###1040�@ 
                If (ret <> 0) Then Exit Try ' ###1040�@

                ' ��������
                ret = CheckRelation()
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
                        Case 0 ' ���Ѹ�ٰ���ޯ��
                            With .W_PLT
                                Select Case (tag)
                                    Case 0 ' ��ۯ�����X
                                        ret = CheckDoubleData(cTextBox, .zsx)
                                    Case 1 ' ��ۯ�����Y
                                        ret = CheckDoubleData(cTextBox, .zsy)
                                    Case 2 ' ��ۯ���X
                                        ret = CheckShortData(cTextBox, .BNX)
                                    Case 3 ' ��ۯ���Y
                                        ret = CheckShortData(cTextBox, .BNY)
                                    Case 4 ' ��R��
                                        Dim cnt As Integer = .RCount ' �ύX�O�̒l��ێ�
                                        'V2.0.0.0�I��
                                        Dim RCount As Short
                                        Dim FromNo As Short = 1
                                        Dim CircuitNo As Integer
                                        If m_MainEdit.W_stUserData.iTrimType = 3 Then
                                            RCount = UserSub.CircuitResistorCount(m_MainEdit.W_PLT, m_MainEdit.W_REG)
                                        Else

                                            'V2.0.0.0�@��
                                            FromNo = 1
                                            For j As Integer = 1 To cnt
                                                If UserModule.IsCutResistor(m_MainEdit.W_REG, j) Then
                                                    FromNo = j
                                                    Exit For
                                                End If
                                            Next
                                            'V2.0.0.0�@��
                                        End If
                                        'V2.0.0.0�I��
                                        ret = CheckShortData(cTextBox, .RCount)
                                        If (cnt <> .RCount) Then
                                            If (cnt < .RCount) Then ' �ǉ����ꂽ�ꍇ
                                                If m_MainEdit.W_stUserData.iTrimType = 3 Then
                                                    cnt = RCount
                                                End If
                                                For i As Integer = (cnt + 1) To .RCount Step 1
                                                    Call InitResData(i) ' �ǉ����ꂽ�ް���������
                                                    'V2.0.0.0�I��
                                                    If m_MainEdit.W_stUserData.iTrimType = 3 And RCount > 1 Then
                                                        m_MainEdit.W_REG(i).Initialize()
                                                        CircuitNo = i \ RCount
                                                        FromNo = i Mod RCount
                                                        If FromNo = 0 Then
                                                            FromNo = RCount
                                                        Else
                                                            CircuitNo = CircuitNo + 1
                                                        End If
                                                        Call CopyResistorData(m_MainEdit.W_REG(i), m_MainEdit.W_REG(FromNo))
                                                        m_MainEdit.W_PTN(i).PtnFlg = m_MainEdit.W_PTN(FromNo).PtnFlg
                                                        m_MainEdit.W_PTN(i).intGRP = m_MainEdit.W_PTN(FromNo).intGRP
                                                        m_MainEdit.W_PTN(i).intPTN = m_MainEdit.W_PTN(FromNo).intPTN
                                                        m_MainEdit.W_REG(i).intCircuitNo = CircuitNo
                                                    Else
                                                        'V2.0.0.0�I��
                                                        'V1.2.0.0�D��
                                                        m_MainEdit.W_REG(i).Initialize()
                                                        'V2.0.0.0�@�R�s�[�����P����FromNo�֕ύX
                                                        Call CopyResistorData(m_MainEdit.W_REG(i), m_MainEdit.W_REG(FromNo))
                                                        m_MainEdit.W_PTN(i).PtnFlg = m_MainEdit.W_PTN(FromNo).PtnFlg
                                                        m_MainEdit.W_PTN(i).intGRP = m_MainEdit.W_PTN(FromNo).intGRP
                                                        m_MainEdit.W_PTN(i).intPTN = m_MainEdit.W_PTN(FromNo).intPTN
                                                        'V1.2.0.0�D��
                                                    End If                                                                  'V2.0.0.0�I
                                                Next i
                                                'V2.0.0.0�@��
                                                If m_MainEdit.W_stUserData.iTrimType = 1 Or m_MainEdit.W_stUserData.iTrimType = 4 Then
                                                    For i As Integer = 1 To .RCount
                                                        If UserModule.IsCutResistor(m_MainEdit.W_REG, i) Then
                                                            m_MainEdit.W_REG(i).intCircuitNo = i
                                                        End If
                                                    Next
                                                End If
                                                'V2.0.0.0�@��
                                            Else ' �폜���ꂽ�ꍇ
                                                For i As Integer = (.RCount + 1) To cnt Step 1
                                                    Call InitResData(i) ' �폜���ꂽ�ް���������
                                                Next i
                                            End If
                                            m_ResNo = 1 ' �������̒�R�ԍ�
                                        End If

                                    Case 5 ' ð��وʒu�̾��X
                                        ret = CheckDoubleData(cTextBox, .z_xoff)
                                    Case 6 ' ð��وʒu�̾��Y
                                        ret = CheckDoubleData(cTextBox, .z_yoff)
                                    Case 7 ' �ްшʒu�̾��X
                                        ret = CheckDoubleData(cTextBox, .BPOX)
                                    Case 8 ' �ްшʒu�̾��Y
                                        ret = CheckDoubleData(cTextBox, .BPOY)
                                        'V2.0.0.0�@                                    Case 9 ' ��ެ���߲��X
                                        'V2.0.0.0�@                                        ret = CheckDoubleData(cTextBox, .ADJX)
                                        'V2.0.0.0�@                                    Case 10 ' ��ެ���߲��Y
                                        'V2.0.0.0�@                                        ret = CheckDoubleData(cTextBox, .ADJY)
                                        'V2.0.0.0�@��
                                    Case 9 ' �X�e�b�v�I�t�Z�b�gX
                                        ret = CheckDoubleData(cTextBox, .dblStepOffsetXDir)
                                    Case 10 ' �X�e�b�v�I�t�Z�b�gY
                                        ret = CheckDoubleData(cTextBox, .dblStepOffsetYDir)
                                        'V2.0.0.0�@��
                                    Case 11 ' ��۰�ސڐG�ʒu
                                        ret = CheckDoubleData(cTextBox, .Z_ZON)
                                    Case 12 ' ��۰�ޑҋ@�ʒu
                                        ret = CheckDoubleData(cTextBox, .Z_ZOFF)
                                    Case 13 ' �O���@�퐔
                                        Dim cnt As Integer = .GCount ' �ύX�O�̒l��ێ�
                                        ret = CheckShortData(cTextBox, .GCount)
                                        If (cnt <> .GCount) Then
                                            If (cnt < .GCount) Then ' �ǉ����ꂽ�ꍇ
                                                For i As Integer = (cnt + 1) To .GCount Step 1
                                                    Call InitGpibData(i) ' �ǉ����ꂽ�ް���������
                                                Next i
                                            Else ' �폜���ꂽ�ꍇ
                                                For i As Integer = (.GCount + 1) To cnt Step 1
                                                    Call InitGpibData(i) ' �폜���ꂽ�ް���������
                                                Next i
                                            End If
                                            m_GpibNo = 1 ' ��������GP-IB�o�^�ԍ�
                                        End If
                                        'V1.2.0.0�@��
                                    Case 14 ' �`�b�v�T�C�YX
                                        ret = CheckDoubleData(cTextBox, .dblChipSizeXDir)
                                    Case 15 ' �`�b�v�T�C�YY
                                        ret = CheckDoubleData(cTextBox, .dblChipSizeYDir)
                                    'V2.2.0.0�N��
                                    Case 16 ' ��۰��No
                                        ret = CheckShortData(cTextBox, .ProbNo)
                                        'V2.2.0.0�N��
                                    Case Else
                                        Throw New Exception("Case " & tag & ": Nothing")
                                        'V1.2.0.0�@��
                                End Select
                            End With
                            ' ------------------------------------------------------------------------------
                        Case 1 ' ڰ����ܰ������ٰ���ޯ��
                            If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then ' ̧���ڰ�ް�łȂ��ꍇ
                                Select Case (tag)
                                    Case 0 ' Qڰ�
                                        Dim dblWK As Double
                                        ret = CheckDoubleData(cTextBox, dblWK)
                                        .W_LASER.intQR = Convert.ToInt16(dblWK * 10) ' KHz �� 0.1KHz
                                    Case 1 ' �ݒ���ܰ
                                        ret = CheckDoubleData(cTextBox, .W_LASER.dblspecPower)
                                    Case 2 ' �A�b�e�l�[�^�������i0:�ۑ��� 1:�ۑ��j
                                        ret = CheckShortData(cTextBox, .W_LASER.iTrimAtt)
                                    Case 3 'V2.1.0.0�A�@�A�b�e�l�[�^No.
                                        ret = CheckShortData(cTextBox, .W_LASER.iAttNo)
                                    Case Else
                                        Throw New Exception("Case " & tag & ": Nothing")
                                End Select
                            End If
                            ' ------------------------------------------------------------------------------
                        Case 2 ' ###1040 �e�B�[�`���O�E�u���b�N��ٰ���ޯ��
                            With .W_PLT
                                Call SetControlData(m_CtlTeachBlock)    ' �u���b�N���ɂ�������~�b�g��ύX����B
                                Select Case (tag)
                                    Case 0 ' �e�B�[�`���O�u���b�N�ʒu�w
                                        ret = CheckShortData(cTextBox, .TeachBlockX)
                                    Case 1 ' �e�B�[�`���O�u���b�N�ʒu�x
                                        ret = CheckShortData(cTextBox, .TeachBlockY)
                                End Select
                            End With
                            ' ------------------------------------------------------------------------------
                        Case 3 ' ###1040 �X�e�[�W���x��ٰ���ޯ��
                            With .W_PLT
                                Select Case (tag)
                                    Case 0 ' �X�e�[�W�E�X�s�[�h�x
                                        ret = CheckShortData(cTextBox, .StageSpeedY)
                                End Select
                            End With
                            ' ------------------------------------------------------------------------------
                        Case 4 'V2.2.0.0�A �i���ٰ���ޯ��
                            With .W_PLT
                                Select Case (tag)
                                    Case 0 ' �\���{��
                                        ret = CheckDoubleData(cTextBox, .dblStdMagnification)
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

#Region "��������"
        ''' <summary>������������</summary>
        ''' <returns>0 = ����, 1 = �װ</returns>
        Protected Overrides Function CheckRelation() As Integer
            Dim strMsg As String
            Dim errIdx As Integer
            CheckRelation = 0 ' Return�l = ����
            Try
                ' ��۰��OFF�ʒu(mm) >= ��۰��ON�ʒu(mm) ?
                With m_MainEdit
                    If (.W_PLT.Z_ZOFF >= .W_PLT.Z_ZON) Then
                        errIdx = SYS_ZOFF
                        strMsg = "���փ`�F�b�N�G���[" & vbCrLf
                        strMsg = strMsg & DirectCast(m_CtlSystem(SYS_ZOFF), cTxt_).GetStrMsg() & " < " _
                                        & DirectCast(m_CtlSystem(SYS_ZON), cTxt_).GetStrMsg() & _
                                        "�ƂȂ�悤�Ɏw�肵�Ă��������B"
                        GoTo STP_ERR
                    End If
                End With
                Exit Function
STP_ERR:
                Call MsgBox_CheckErr(DirectCast(m_CtlSystem(errIdx), cTxt_), strMsg)
                CheckRelation = 1 ' Return�l = �װ

            Catch ex As Exception
                Call MsgBox_Exception(ex)
                CheckRelation = 1 ' Return�l = �װ
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
                        Case 0 ' ���Ѹ�ٰ���ޯ��
                            Select Case (tag)
                                Case 0 ' ��۰����ײ(1:�L, 0:��)
                                    .W_PLT.PrbRetry = Convert.ToInt16(idx)
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case 1 ' ڰ����ܰ������ٰ���ޯ��
                            Throw New Exception("Parent.Tag - Case 1")
                        Case Else
                            Throw New Exception("Parent.Tag - Case Else")
                    End Select
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

        Private Sub CTxt_20_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CTxt_20.TextChanged

        End Sub
        Private Sub CTxt_21_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CTxt_21.TextChanged

        End Sub
        Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
            Try
                Dim CutNo As Short
                Dim ret As Integer = 1
                Dim FromNo As Short     'V2.0.0.0�I
                Dim OrderNo As Short    'V2.0.0.0�I
                '--------------------------------------------------------------------------
                '   �m�Fү���ނ�\������
                '--------------------------------------------------------------------------
                Dim strMsg As String = "�`�b�v�T�C�Y�𔽉f���܂��B��낵���ł����H"
                ret = MsgBox(strMsg, DirectCast( _
                            MsgBoxStyle.OkCancel + _
                            MsgBoxStyle.Information, MsgBoxStyle), _
                            My.Application.Info.Title)
                If (ret = MsgBoxResult.Cancel) Then ' Cancel(RESET��) ?
                    Exit Sub
                End If

                'V2.0.0.0�@��
                m_MainEdit.W_PLT.dblStepOffsetXDir = 0.0                       ' �X�e�b�v�I�t�Z�b�g��X
                m_MainEdit.W_PLT.dblStepOffsetYDir = 0.0                       ' �X�e�b�v�I�t�Z�b�g��Y
                'V2.0.0.0�@��

                'V2.0.0.0�I��
                If m_MainEdit.W_stUserData.iTrimType = 3 And UserSub.CircuitResistorCount(m_MainEdit.W_PLT, m_MainEdit.W_REG) > 1 Then
                    Dim RCount As Short = UserSub.CircuitResistorCount(m_MainEdit.W_PLT, m_MainEdit.W_REG)
                    With m_MainEdit
                        For rn As Integer = RCount + 1 To .W_PLT.RCount Step 1
                            With .W_REG(rn)
                                If .intSLP <> SLP_OK_MARK And .intSLP <> SLP_NG_MARK Then   ' OK,NG�}�[�L���O�͎��{�R�s�[���Ȃ�
                                    OrderNo = UserSub.GetResNumberInCircuit(m_MainEdit.W_REG, rn)   ' �T�[�L�b�g���̒�R�̏���
                                    FromNo = UserSub.GetRNumByCircuit(m_MainEdit.W_PLT, m_MainEdit.W_REG, 1, OrderNo)   ' ���T�[�L�b�g���̏��Ԃ̒�R�ԍ�
                                    ' ��Đ����J�Ԃ�
                                    CutNo = .intTNN
                                    If CutNo > m_MainEdit.W_REG(FromNo).intTNN Then
                                        CutNo = m_MainEdit.W_REG(FromNo).intTNN
                                    End If
                                    For cn As Integer = 1 To CutNo Step 1
                                        With .STCUT(cn)
                                            .dblSTX = m_MainEdit.W_REG(FromNo).STCUT(cn).dblSTX + m_MainEdit.W_PLT.dblChipSizeXDir * (m_MainEdit.W_REG(rn).intCircuitNo - 1)
                                            .dblSTY = m_MainEdit.W_REG(FromNo).STCUT(cn).dblSTY + m_MainEdit.W_PLT.dblChipSizeYDir * (m_MainEdit.W_REG(rn).intCircuitNo - 1)
                                        End With
                                    Next cn
                                    m_MainEdit.W_PTN(rn).dblPosX = stPTN(FromNo).dblPosX + m_MainEdit.W_PLT.dblChipSizeXDir * (m_MainEdit.W_REG(rn).intCircuitNo - 1)
                                    m_MainEdit.W_PTN(rn).dblPosY = stPTN(FromNo).dblPosY + m_MainEdit.W_PLT.dblChipSizeYDir * (m_MainEdit.W_REG(rn).intCircuitNo - 1)
                                End If
                            End With
                        Next rn
                    End With
                Else
                    'V2.0.0.0�I��

                    FromNo = 0
                    Dim Circuit As Short = 0
                    With m_MainEdit
                        'V2.0.0.0�@                        For rn As Integer = 2 To .W_PLT.RCount Step 1
                        For rn As Integer = 1 To .W_PLT.RCount Step 1
                            With .W_REG(rn)
                                'V2.0.0.0�@                                If .intSLP <> SLP_OK_MARK And .intSLP <> SLP_NG_MARK Then   ' OK,NG�}�[�L���O�͎��{�R�s�[���Ȃ�
                                If UserModule.IsCutResistor(stREG, rn) Then
                                    If FromNo = 0 Then
                                        FromNo = rn
                                        Circuit = 2
                                        Continue For
                                    End If
                                    ' ��Đ����J�Ԃ�
                                    CutNo = .intTNN
                                    If CutNo > m_MainEdit.W_REG(FromNo).intTNN Then
                                        CutNo = m_MainEdit.W_REG(FromNo).intTNN
                                    End If
                                    For cn As Integer = 1 To CutNo Step 1
                                        With .STCUT(cn)
                                            .dblSTX = m_MainEdit.W_REG(FromNo).STCUT(cn).dblSTX + m_MainEdit.W_PLT.dblChipSizeXDir * (Circuit - 1)
                                            .dblSTY = m_MainEdit.W_REG(FromNo).STCUT(cn).dblSTY + m_MainEdit.W_PLT.dblChipSizeYDir * (Circuit - 1)
                                        End With
                                    Next cn
                                    'V1.2.0.0�D��
                                    m_MainEdit.W_PTN(rn).dblPosX = stPTN(FromNo).dblPosX + m_MainEdit.W_PLT.dblChipSizeXDir * (Circuit - 1)
                                    m_MainEdit.W_PTN(rn).dblPosY = stPTN(FromNo).dblPosY + m_MainEdit.W_PLT.dblChipSizeYDir * (Circuit - 1)
                                    'V1.2.0.0�D��
                                    Circuit = Circuit + 1
                                End If
                            End With
                        Next rn
                    End With
                End If      'V2.0.0.0�I

                m_MainEdit.LblToolTip.Text = "�`�b�v�T�C�Y�̔��f������ɏI�����܂����B"
            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
    End Class
End Namespace
