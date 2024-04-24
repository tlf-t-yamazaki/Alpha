Option Explicit On
Option Strict Off

Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports LaserFront.Trimmer.DefWin32Fnc

Namespace FormEdit
    Friend Class tabGpib
        Inherits tabBase

#Region "�錾"
        Private Const GPIB_GNAM As Integer = 3  ' m_CtlGpib�ł̲��ޯ��(�@�햼)
        Private Const GPIB_CON As Integer = 7   ' m_CtlGpib�ł̲��ޯ��(ON�����)
        Private Const GPIB_COFF As Integer = 9  ' m_CtlGpib�ł̲��ޯ��(OFF�����)
        Private Const GPIB_CTRG As Integer = 11 ' m_CtlGpib�ł̲��ޯ��(�ضް�����)

        Private m_CtlGpib() As Control          ' GP-IB��ٰ���ޯ���̺��۰ٔz��
        Private m_TrgCmdFlg As Boolean          ' �ضް����� ����(�����)=True,�Ȃ�(�d��)=False
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

            m_TabIdx = tabIdx           ' ҲݕҏW�����޺��۰ُ�ł̲��ޯ��
            m_MainEdit = mainEdit       ' ҲݕҏW��ʂւ̎Q�Ƃ�ݒ�

            Try
                ' EDIT_DEF_User.ini������ޖ���ݒ�
                TAB_NAME = GetPrivateProfileString_S("GPIB_LABEL", "TAB_NAM", m_sPath, "????")

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
                    CGrp_0 _
                }
                For i As Integer = 0 To (GrpArray.Length - 1) Step 1
                    With GrpArray(i)
                        .TabIndex = ((254 - GrpArray.Length) + i) ' ���ٷ��ɂ��̫����ړ��ŕK�v
                        .Tag = i
                        .Text = GetPrivateProfileString_S( _
                            "GPIB_LABEL", (i.ToString("000") & "_GRP"), m_sPath, "????")
                    End With
                Next i
                ' �ǉ���폜���݂�����
                CPnl_Btn.TabIndex = 254 ' ���۰ٔz�u�\�ő吔(�Ō�ɐݒ�)

                ' ----------------------------------------------------------
                ' EDIT_DEF_User.ini�������قɕ\������ݒ�
                ' ----------------------------------------------------------
                LblArray = New cLbl_() { _
                    CLbl_0, CLbl_1, CLbl_2, CLbl_3, CLbl_4, _
                    CLbl_5, _
                    CLbl_6, CLbl_7, _
                    CLbl_8, CLbl_9, _
                    CLbl_10 _
                }
                For i As Integer = 0 To (LblArray.Length - 1) Step 1
                    LblArray(i).Text = GetPrivateProfileString_S( _
                        "GPIB_LABEL", (i.ToString("000") & "_LBL"), m_sPath, "????")
                Next i

                ' ----------------------------------------------------------
                ' GP-IB��ٰ���ޯ�����̺��۰ق�ݒ�
                ' ----------------------------------------------------------
                m_CtlGpib = New Control() { _
                    CCmb_0, CTxt_0, CCmb_1, CTxt_1, _
                    CTxt_2, CTxt_8, CTxt_9, _
                    CTxt_3, CTxt_4, _
                    CTxt_5, CTxt_6, _
                    CTxt_7 _
                }
                Call SetControlData(m_CtlGpib)

                ' ----------------------------------------------------------
                ' ���ׂĂ�÷���ޯ��������ޯ������݂ɑ΂��A̫����ړ��ݒ�������Ȃ�
                ' ��޷��A���ٷ��ɂ��̫����ړ����鏇�Ԃź��۰ق�CtlArray�ɐݒ肷��
                ' �g�p���Ȃ����۰ق� Enabled=False �܂��� Visible=False �ɂ���
                ' ----------------------------------------------------------
                CtlArray = New Control() { _
                    CCmb_0, CTxt_0, CCmb_1, CTxt_1, _
                    CTxt_2, CTxt_8, CTxt_9, _
                    CTxt_3, CTxt_4, _
                    CTxt_5, CTxt_6, _
                    CTxt_7, _
                    CBtn_Add, CBtn_Del _
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
        ''' <summary>���������ɺ����ޯ����ؽĥү���ސݒ�������Ȃ�</summary>
        ''' <param name="cCombo">�ݒ�������Ȃ������ޯ��</param>
        Protected Overrides Sub InitComboBox(ByRef cCombo As cCmb_)
            Dim tag As Integer
            Try
                With cCombo
                    tag = DirectCast(.Tag, Integer)
                    .Items.Clear()
                    Select Case (DirectCast(.Parent.Tag, Integer)) ' ��ٰ���ޯ�������
                        ' ------------------------------------------------------------------------------
                        Case 0 ' GP-IB��ٰ���ޯ��
                            Select Case (tag)
                                Case 0 ' �o�^�ԍ�
                                    '.Items.Add("") ' ڲ��Ĳ���ĂōĐݒ肳���
                                Case 1 ' �����(0:CRLF, 1:CR, 2:LF, 3:�Ȃ�)
                                    .Items.Add("�Ȃ�")                      'V2.1.0.0�C
                                    .Items.Add("CRLF")                      'V2.1.0.0�C
                                    .Items.Add("CR")                        'V2.1.0.0�C
                                    .Items.Add("LF")                        'V2.1.0.0�C
                                    'V2.1.0.0�C                                    .Items.Add("CRLF")
                                    'V2.1.0.0�C                                    .Items.Add("LF")
                                    'V2.1.0.0�C                                    .Items.Add("�Ȃ�")
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
        ''' <summary>����������÷���ޯ���̏㉺���l�ү���ސݒ�������Ȃ�</summary>
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
                    Case 0 ' GP-IB��ٰ���ޯ��
                        ' �ް������װ���̕\����
                        strMsg = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0 ' ���ڽ
                                strMin = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MAX"), m_sPath, "30")
                            Case 1 ' �@�햼
                                strMin = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MAX"), m_sPath, "20")
                                strFlg = True
                            Case 2 ' �ݒ�����
                                strMin = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MAX"), m_sPath, "100")
                                strFlg = True
                            Case 3 ' �ݒ�����(2�i��)
                                strMin = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MAX"), m_sPath, "100")
                                strFlg = True
                            Case 4 ' �ݒ�����(3�i��)
                                strMin = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MAX"), m_sPath, "100")
                                strFlg = True
                            Case 5 ' ON�����
                                strMin = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MAX"), m_sPath, "50")
                                strFlg = True
                            Case 6 ' ON����߰�ގ���
                                strMin = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MAX"), m_sPath, "32767")
                            Case 7 ' OFF�����
                                strMin = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MAX"), m_sPath, "50")
                                strFlg = True
                            Case 8 ' OFF����߰�ގ���
                                strMin = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MAX"), m_sPath, "32767")
                            Case 9 ' �ضް�����
                                strMin = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MAX"), m_sPath, "50")
                                strFlg = True
                            Case Else
                                Throw New Exception("Case " & tag & ": Nothing")
                        End Select
                        ' ------------------------------------------------------------------------------
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

#Region "÷���ޯ���ɒl��ݒ肷��"
        ''' <summary>÷���ޯ���ɒl��ݒ肷��</summary>
        Protected Overrides Sub SetDataToText()
            Try
                Me.SuspendLayout()
                ' GP-IB��ٰ���ޯ���ݒ�
                Call SetGpibData()

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            Finally
                Me.ResumeLayout()
                Me.Refresh()
            End Try

        End Sub
#End Region

#Region "GP-IB��ٰ���ޯ�����̐ݒ�"
        ''' <summary>GP-IB��ٰ���ޯ������÷���ޯ��������ޯ���ɒl��ݒ肷��</summary>
        Private Sub SetGpibData()
            With m_MainEdit
                If (.W_PLT.GCount < 1) Then ' �o�^�� = 0 ?
                    CLblNum.Text = "0" ' �o�^��
                    m_GpibNo = 1
                Else
                    CLblNum.Text = (.W_PLT.GCount).ToString() ' �o�^��
                End If
            End With

            Try
                With m_MainEdit.W_GPIB(m_GpibNo)
                    For i As Integer = 0 To (m_CtlGpib.Length - 1) Step 1
                        Select Case (i)
                            Case 0 ' �o�^�ԍ�
                                Dim gCnt As Integer = m_MainEdit.W_PLT.GCount ' �o�^��
                                Dim cCombo As cCmb_ = DirectCast(m_CtlGpib(i), cCmb_)
                                With cCombo
                                    .Items.Clear()
                                    If (0 < gCnt) Then
                                        For j As Integer = 1 To gCnt Step 1
                                            .Items.Add(String.Format("{0,5:#0}", j))
                                        Next
                                    Else
                                        .Items.Add(String.Format("{0,5:#0}", m_GpibNo))
                                    End If
                                End With
                                Call NoEventIndexChange(cCombo, (m_GpibNo - 1))

                            Case 1 ' ���ڽ
                                m_CtlGpib(i).Text = (.intGAD).ToString()
                            Case 2 ' �����(0:CRLF, 1:CR, 2:LF, 3:�Ȃ�)
                                Dim cCombo As cCmb_ = DirectCast(m_CtlGpib(i), cCmb_)
                                Call NoEventIndexChange(cCombo, .intDLM)
                            Case 3 ' �@�햼
                                m_CtlGpib(i).Text = .strGNAM
                            Case 4 ' �ݒ�����(1�i��)
                                m_CtlGpib(i).Text = (.strCCMD1).ToString()  'V2.0.0.0�C
                                'V2.0.0.0�C                                If .strCCMD.Length >= 100 Then
                                'V2.0.0.0�C                                    m_CtlGpib(i + 0).Text = .strCCMD.Substring(0, 100)
                                'V2.0.0.0�C                                Else
                                'V2.0.0.0�C                                    m_CtlGpib(i + 0).Text = .strCCMD
                                'V2.0.0.0�C                                End If
                            Case 5  ' �ݒ�����(2�i��)
                                m_CtlGpib(i).Text = (.strCCMD2).ToString()  'V2.0.0.0�C
                                'V2.0.0.0�C                                If .strCCMD.Length < 100 Then
                                'V2.0.0.0�C                                    m_CtlGpib(i).Text = String.Empty
                                'V2.0.0.0�C                                ElseIf 100 < .strCCMD.Length And .strCCMD.Length <= 200 Then
                                'V2.0.0.0�C                                    m_CtlGpib(i).Text = .strCCMD.Substring(100)
                                'V2.0.0.0�C                                Else
                                'V2.0.0.0�C                                    m_CtlGpib(i).Text = .strCCMD.Substring(100, 100)
                                'V2.0.0.0�C                                End If
                            Case 6  ' �ݒ�����(3�i��)
                                m_CtlGpib(i).Text = (.strCCMD3).ToString()  'V2.0.0.0�C
                                'V2.0.0.0�C                                    If .strCCMD.Length < 200 Then
                                'V2.0.0.0�C                                        m_CtlGpib(i).Text = String.Empty
                                'V2.0.0.0�C                                    ElseIf 200 < .strCCMD.Length Then
                                'V2.0.0.0�C                                        m_CtlGpib(i).Text = .strCCMD.Substring(200)
                                'V2.0.0.0�C                                    End If
                            Case 7 ' ON�����
                                m_CtlGpib(i).Text = .strCON
                            Case 8 ' ON����߰�ގ���
                                m_CtlGpib(i).Text = (.lngPOWON).ToString()
                            Case 9 ' OFF�����
                                m_CtlGpib(i).Text = .strCOFF
                            Case 10 ' OFF����߰�ގ���
                                m_CtlGpib(i).Text = (.lngPOWOFF).ToString()
                            Case 11 ' �ضް�����
                                m_CtlGpib(i).Text = .strCTRG
                                If ("" = .strCTRG) Then ' �ضް�����
                                    m_TrgCmdFlg = False ' �Ȃ�(�d��)=False
                                Else
                                    m_TrgCmdFlg = True ' ����(�����)=True
                                End If
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

                    If (.W_PLT.GCount < 1) Then ' �ް��Ȃ��Ȃ�NOP
                        ret = 0
                        Exit Try
                    End If

                    For gn As Integer = 1 To .W_PLT.GCount Step 1
                        m_GpibNo = gn
                        ' ��������o�^�ԍ����ް�����۰قɾ�Ă���
                        Call SetDataToText()
                        ret = CheckControlData(m_CtlGpib)
                        If (ret <> 0) Then Exit Try

                        ' ��������
                        ret = CheckRelation()
                        If (ret <> 0) Then Exit Try
                    Next gn
                End With

                Call CheckDataUpdate() ' �ް��X�V�m�F

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
            Dim ret As Integer = 1
            Dim strTmp As String = String.Empty
            Try
                With m_MainEdit.W_GPIB(m_GpibNo)
                    tag = DirectCast(cTextBox.Parent.Tag, Integer)
                    Select Case (tag)
                        ' ------------------------------------------------------------------------------
                        Case 0 ' GP-IB
                            Select Case (DirectCast(cTextBox.Tag, Integer))
                                Case 0  ' ���ڽ
                                    ret = CheckShortData(cTextBox, .intGAD)
                                Case 1  ' �@�햼
                                    ret = CheckStrData(cTextBox, .strGNAM)
                                Case 2  ' �ݒ�����
                                    ret = CheckStrData(cTextBox, .strCCMD1)         'V2.0.0.0�C
                                    'V2.0.0.0�C                                    ret = CheckStrData(cTextBox, strTmp)
                                    'V2.0.0.0�C                                    .strCCMD = strTmp + m_CtlGpib(5).Text + m_CtlGpib(6).Text
                                Case 3  ' �ݒ�����(2�i��)
                                    ret = CheckStrData(cTextBox, .strCCMD2)         'V2.0.0.0�C
                                    'V2.0.0.0�C                                    ret = CheckStrData(cTextBox, strTmp)
                                    'V2.0.0.0�C                                    .strCCMD = m_CtlGpib(4).Text + strTmp + m_CtlGpib(6).Text
                                Case 4  ' �ݒ�����(3�i��)
                                    ret = CheckStrData(cTextBox, .strCCMD3)         'V2.0.0.0�C
                                    'V2.0.0.0�C                                    ret = CheckStrData(cTextBox, strTmp)
                                    'V2.0.0.0�C                                    .strCCMD = m_CtlGpib(4).Text + m_CtlGpib(5).Text + strTmp
                                Case 5  ' ON�����
                                    ret = CheckStrData(cTextBox, .strCON)
                                Case 6  ' ON����߰�ގ���
                                    ret = CheckIntData(cTextBox, .lngPOWON)
                                Case 7  ' OFF�����
                                    ret = CheckStrData(cTextBox, .strCOFF)
                                Case 8  ' OFF����߰�ގ���
                                    ret = CheckIntData(cTextBox, .lngPOWOFF)
                                Case 9  ' �ضް�����
                                    ret = CheckStrData(cTextBox, .strCTRG)
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case Else
                            Throw New Exception("Parent.Tag - Case Else")
                    End Select
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
                ret = 1
            Finally
                CheckTextData = ret
            End Try

        End Function
#End Region

#Region "��������"
        Protected Overrides Function CheckRelation() As Integer
            Dim erridx As Integer
            Dim strMsg As String

            CheckRelation = 0 ' Return�l = ����
            Try
                With m_MainEdit.W_GPIB(m_GpibNo)
                    '---------------------------------------------------------------
                    ' �R�}���h���͂Ȃ�
                    '---------------------------------------------------------------
                    If ("" = .strCON) AndAlso ("" = .strCOFF) AndAlso ("" = .strCTRG) Then
                        erridx = GPIB_CON
                        strMsg = "���փ`�F�b�N�G���[" & vbCrLf & _
                                "�R�}���h����͂��ĉ������B"
                        GoTo STP_ERR
                    End If

                    '---------------------------------------------------------------
                    ' ON�܂���OFF���ضް����ނ����͂���Ă���
                    '---------------------------------------------------------------
                    'If ("" <> .strCTRG) AndAlso _
                    '        (("" <> .strCON) OrElse ("" <> .strCOFF)) Then
                    '    erridx = GPIB_CTRG
                    '    strMsg = "���փ`�F�b�N�G���[" & vbCrLf & _
                    '            "�R�}���h�̑g�ݍ��킹������������܂���B"
                    '    GoTo STP_ERR
                    'End If

                    '---------------------------------------------------------------
                    ' ON���͂��肩��OFF������
                    '---------------------------------------------------------------
                    If ("" <> .strCON) AndAlso ("" = .strCOFF) Then
                        erridx = GPIB_COFF
                        strMsg = "���փ`�F�b�N�G���[" & vbCrLf & _
                                "�R�}���h�̑g�ݍ��킹������������܂���B"
                        GoTo STP_ERR
                    End If

                    '---------------------------------------------------------------
                    ' ON�����͂���OFF���͂���
                    '---------------------------------------------------------------
                    If ("" = .strCON) AndAlso ("" <> .strCOFF) Then
                        erridx = GPIB_CON
                        strMsg = "���փ`�F�b�N�G���[" & vbCrLf & _
                                "�R�}���h�̑g�ݍ��킹������������܂���B"
                        GoTo STP_ERR
                    End If
                End With

                Exit Function
STP_ERR:
                If (TypeOf m_CtlGpib(erridx) Is cTxt_) Then
                    ' ÷���ޯ�����װ�̏ꍇ
                    Call MsgBox_CheckErr(DirectCast(m_CtlGpib(erridx), cTxt_), strMsg)
                ElseIf (TypeOf m_CtlGpib(erridx) Is cCmb_) Then
                    ' �����ޯ�����װ�̏ꍇ
                    Call MsgBox_CheckErr(DirectCast(m_CtlGpib(erridx), cCmb_), strMsg)
                Else
                    ' DO NOTHING
                End If
                CheckRelation = 1 ' Return�l = �װ

            Catch ex As Exception
                Call MsgBox_Exception(ex)
                CheckRelation = 1 ' Return�l = �װ
            End Try

        End Function
#End Region

#Region "�ް��̍X�V���m�F����"
        ''' <summary>�ް��̍X�V���������ꍇ�AGPIB�ް��X�VFlag��ON�ɂ���</summary>
        Private Sub CheckDataUpdate()
            Dim flg As Integer = 0
            Try
                With m_MainEdit
                    If (.W_PLT.GCount <> stPLT.GCount) Then flg = 1 : Exit Try ' �o�^��
                    For i As Integer = 1 To stPLT.GCount Step 1
                        With .W_GPIB(i)
                            If (.intGAD <> stGPIB(i).intGAD) Then flg = 1 : Exit Try ' ���ڽ
                            If (.intDLM <> stGPIB(i).intDLM) Then flg = 1 : Exit Try ' �����
                            If (.strGNAM <> stGPIB(i).strGNAM) Then flg = 1 : Exit Try ' �@�햼
                            'V2.0.0.0�C                            If (.strCCMD <> stGPIB(i).strCCMD) Then flg = 1 : Exit Try ' �ݒ�����
                            If (.strCCMD1 <> stGPIB(i).strCCMD1) Then flg = 1 : Exit Try ' �ݒ�����'V2.0.0.0�C
                            If (.strCCMD2 <> stGPIB(i).strCCMD2) Then flg = 1 : Exit Try ' �ݒ�����'V2.0.0.0�C
                            If (.strCCMD3 <> stGPIB(i).strCCMD3) Then flg = 1 : Exit Try ' �ݒ�����'V2.0.0.0�C
                            If (.strCON <> stGPIB(i).strCON) Then flg = 1 : Exit Try ' ON�����
                            If (.strCOFF <> stGPIB(i).strCOFF) Then flg = 1 : Exit Try ' OFF�����
                            If (.lngPOWON <> stGPIB(i).lngPOWON) Then flg = 1 : Exit Try ' ON����߰�ގ���
                            If (.lngPOWOFF <> stGPIB(i).lngPOWOFF) Then flg = 1 : Exit Try ' OFF����߰�ގ���
                            If (.strCTRG <> stGPIB(i).strCTRG) Then flg = 1 : Exit Try ' �ضް�����
                        End With
                    Next i
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            Finally
                FlgUpdGPIB = Convert.ToInt16(flg) ' GPIB�ް��X�VFlag ON=1
            End Try

        End Sub
#End Region

#Region "�ǉ���폜���݊֘A����"
        ''' <summary>GP-IB�ް���ǉ��܂��͍폜���A�����ް�������������</summary>
        ''' <param name="addDel">1=Add, (-1)=Del</param>
        Private Sub SortGpibData(ByVal addDel As Integer)
            Dim iStart As Integer
            Dim iEnd As Integer
            Dim dir As Integer = (-1) * addDel ' Add=(-1), Del=1�ɂ���
            Try
                With m_MainEdit
                    If (1 = addDel) Then ' �ǉ��̏ꍇ
                        .W_PLT.GCount = Convert.ToInt16(.W_PLT.GCount + 1) ' �o�^����ǉ�����
                        iStart = .W_PLT.GCount ' �o�^������
                        iEnd = (m_GpibNo + 1) ' �ǉ������ް��̓o�^�ԍ�+1�܂ŁA�O���ް������ɂ��炷
                    Else ' �폜�̏ꍇ
                        iStart = m_GpibNo ' �폜�����ް��̓o�^�ԍ�����
                        iEnd = (.W_PLT.GCount - 1) ' �o�^����Ă���o�^��-1�܂ŁA�����ް���O�ɂ��炷
                    End If

                    For i As Integer = iStart To iEnd Step dir
                        .W_GPIB(i).intGAD = .W_GPIB(i + dir).intGAD         ' ���ڽ
                        .W_GPIB(i).strGNAM = .W_GPIB(i + dir).strGNAM       ' �@�햼
                        .W_GPIB(i).intDLM = .W_GPIB(i + dir).intDLM         ' �����(0:CRLF, 1:CR, 2:LF, 3:�Ȃ�)
                        'V2.0.0.0�C                        .W_GPIB(i).strCCMD = .W_GPIB(i + dir).strCCMD       ' �ݒ�����
                        .W_GPIB(i).strCCMD1 = .W_GPIB(i + dir).strCCMD1       ' �ݒ�����'V2.0.0.0�C
                        .W_GPIB(i).strCCMD2 = .W_GPIB(i + dir).strCCMD2       ' �ݒ�����'V2.0.0.0�C
                        .W_GPIB(i).strCCMD3 = .W_GPIB(i + dir).strCCMD3       ' �ݒ�����'V2.0.0.0�C
                        .W_GPIB(i).strCON = .W_GPIB(i + dir).strCON         ' ON�����
                        .W_GPIB(i).lngPOWON = .W_GPIB(i + dir).lngPOWON     ' ON����߰�ގ���(ms)
                        .W_GPIB(i).strCOFF = .W_GPIB(i + dir).strCOFF       ' OFF�����
                        .W_GPIB(i).lngPOWOFF = .W_GPIB(i + dir).lngPOWOFF   ' OFF����߰�ގ���(ms)
                        .W_GPIB(i).strCTRG = .W_GPIB(i + dir).strCTRG       ' �ضް�����
                    Next i

                    ' �ǉ��܂��͍폜�����o�^�ԍ��ȍ~�̓o�^�ԍ����g�p����Ă���ꍇ�ɂ��̒l��ύX����
                    Call ResetResCutData(addDel)

                    ' �߂ĕs�v�ƂȂ����ް�������������
                    If (1 = addDel) Then ' �ǉ��̏ꍇ
                        Call InitGpibData(m_GpibNo) ' �ǉ������ް���������
                    Else ' �폜�̏ꍇ
                        Call InitGpibData(.W_PLT.GCount) ' �ŏI�o�^�ԍ����ް���������
                        .W_PLT.GCount = Convert.ToInt16(.W_PLT.GCount - 1) ' �o�^����-1����

                        ' �ŏI�ް��̍폜�Ȃ猻�݂̓o�^�ԍ����ŏI�o�^�ԍ��Ƃ���
                        If (.W_PLT.GCount < m_GpibNo) Then m_GpibNo = .W_PLT.GCount
                    End If
                End With

                ' GP-IB�ް�����ʍ��ڂɐݒ�
                Call SetDataToText()
                FIRST_CONTROL.Select() ' ̵����ݒ�

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

        ''' <summary>�ǉ���폜�����ް��̓o�^�ԍ��ȍ~�̔ԍ����g�p����Ă���ꍇ�ɂ��̒l��ύX����</summary>
        ''' <param name="addDel">1=Add, (-1)=Del</param>
        ''' <param name="delTrgCmd">�ضް����ނ̍폜=True</param>
        Private Sub ResetResCutData(ByVal addDel As Integer, Optional ByVal delTrgCmd As Boolean = False)
            Try
                With m_MainEdit
                    If (1 = addDel) AndAlso (False = delTrgCmd) Then ' �ǉ��̏ꍇ
                        For rn As Integer = 1 To .W_PLT.RCount Step 1
                            With .W_REG(rn)
                                ' �ǉ��ް��̓o�^�ԍ�����̔ԍ����g�p���Ă���ꍇ�A�ԍ���(+1)����
                                If (m_GpibNo <= .intMType) Then ' ��R
                                    .intMType = Convert.ToInt16(.intMType + addDel)
                                End If

                                ' ------------------------
                                For i As Integer = 1 To EXTEQU Step 1
                                    If (m_GpibNo <= .intOnExtEqu(i)) Then   ' ��R��� > ON�@��
                                        ' �ǉ��ް��̓o�^�ԍ�����̔ԍ����g�p���Ă���ꍇ�A�ԍ���(+1)����
                                        .intOnExtEqu(i) = Convert.ToInt16(.intOnExtEqu(i) + addDel)
                                    End If
                                    If (m_GpibNo <= .intOffExtEqu(i)) Then  ' ��R��� > OFF�@��
                                        ' �ǉ��ް��̓o�^�ԍ�����̔ԍ����g�p���Ă���ꍇ�A�ԍ���(+1)����
                                        .intOffExtEqu(i) = Convert.ToInt16(.intOffExtEqu(i) + addDel)
                                    End If
                                Next i
                                ' ------------------------

                                For cn As Integer = 1 To (.STCUT.Length - 1) Step 1 ' ���
                                    With .STCUT(cn)
                                        If (m_GpibNo <= .intMType) Then
                                            .intMType = Convert.ToInt16(.intMType + addDel)
                                        End If

                                        For ix As Integer = 1 To (.intIXMType.Length - 1) Step 1 ' ���ޯ�����
                                            If (m_GpibNo <= .intIXMType(ix)) Then
                                                .intIXMType(ix) = Convert.ToInt16(.intIXMType(ix) + addDel)
                                            End If
                                        Next ix
                                    End With
                                Next cn

                            End With
                        Next rn

                    Else ' �폜�̏ꍇ
                        For rn As Integer = 1 To .W_PLT.RCount Step 1
                            With .W_REG(rn)
                                If (m_GpibNo < .intMType) AndAlso (False = delTrgCmd) Then ' ��R
                                    ' �폜�ް��̓o�^�ԍ�����̔ԍ����g�p���Ă���ꍇ�A�ԍ���(-1)����
                                    .intMType = Convert.ToInt16(.intMType + addDel)
                                ElseIf (m_GpibNo = .intMType) Then ' ��R
                                    ' �폜�ް��A�܂����ضް����ނ��폜���ꂽ�ް���
                                    ' �o�^�ԍ����g�p���̏ꍇ�A0(������R����)�Ƃ���
                                    .intMType = 0
                                Else
                                    ' DO NOTHING
                                End If

                                ' ------------------------
                                For i As Integer = 1 To EXTEQU Step 1
                                    ' ��R��� > ON�@��
                                    If (m_GpibNo < .intOnExtEqu(i)) AndAlso (False = delTrgCmd) Then
                                        ' �폜�ް��̓o�^�ԍ�����̔ԍ����g�p���Ă���ꍇ�A�ԍ���(-1)����
                                        .intOnExtEqu(i) = Convert.ToInt16(.intOnExtEqu(i) + addDel)
                                    ElseIf (m_GpibNo = .intOnExtEqu(i)) Then
                                        ' �폜�ް��A�܂����ضް����ނ��폜���ꂽ�ް���
                                        ' �o�^�ԍ����g�p���̏ꍇ�A0(�Ȃ�)�Ƃ���
                                        .intOnExtEqu(i) = 0
                                    Else
                                        ' DO NOTHING
                                    End If

                                    ' ��R��� > OFF�@��
                                    If (m_GpibNo < .intOffExtEqu(i)) AndAlso (False = delTrgCmd) Then
                                        ' �폜�ް��̓o�^�ԍ�����̔ԍ����g�p���Ă���ꍇ�A�ԍ���(-1)����
                                        .intOffExtEqu(i) = Convert.ToInt16(.intOffExtEqu(i) + addDel)
                                    ElseIf (m_GpibNo = .intOffExtEqu(i)) Then
                                        ' �폜�ް��A�܂����ضް����ނ��폜���ꂽ�ް���
                                        ' �o�^�ԍ����g�p���̏ꍇ�A0(�Ȃ�)�Ƃ���
                                        .intOffExtEqu(i) = 0
                                    Else
                                        ' DO NOTHING
                                    End If
                                Next i
                                ' ------------------------

                                For cn As Integer = 1 To (.STCUT.Length - 1) Step 1 ' ���
                                    With .STCUT(cn)
                                        ' �폜�ް��̓o�^�ԍ�����̔ԍ����g�p���Ă���ꍇ�A�ԍ���(-1)����
                                        If (m_GpibNo < .intMType) AndAlso (False = delTrgCmd) Then
                                            .intMType = Convert.ToInt16(.intMType + addDel)
                                        ElseIf (m_GpibNo = .intMType) Then
                                            ' �폜�ް��A�܂����ضް����ނ��폜���ꂽ�ް���
                                            ' �o�^�ԍ����g�p���̏ꍇ�A0(������R����)�Ƃ���
                                            .intMType = 0
                                        Else
                                            ' DO NOTHING
                                        End If

                                        For ix As Integer = 1 To (.intIXMType.Length - 1) Step 1 ' ���ޯ�����
                                            ' �폜�ް��̓o�^�ԍ�����̔ԍ����g�p���Ă���ꍇ�A�ԍ���(-1)����
                                            If (m_GpibNo < .intIXMType(ix)) AndAlso (False = delTrgCmd) Then
                                                .intIXMType(ix) = Convert.ToInt16(.intIXMType(ix) + addDel)
                                            ElseIf (m_GpibNo = .intIXMType(ix)) Then
                                                ' �폜�ް��A�܂����ضް����ނ��폜���ꂽ�ް���
                                                ' �o�^�ԍ����g�p���̏ꍇ�A0(������R����)�Ƃ���
                                                .intIXMType(ix) = 0
                                            Else
                                                ' DO NOTHING
                                            End If
                                        Next ix
                                    End With
                                Next cn

                            End With
                        Next rn

                    End If
                End With

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
                        Case 0 ' GP-IB��ٰ���ޯ��
                            Select Case (tag)
                                Case 0 ' �o�^�ԍ�
                                    m_GpibNo = (idx + 1)
                                    ' �Ή������ް���÷���ޯ���ɾ�Ă���
                                    Call SetDataToText()
                                Case 1 ' �����(0:CRLF, 1:CR, 2:LF, 3:�Ȃ�)
                                    .W_GPIB(m_GpibNo).intDLM = Convert.ToInt16(idx)
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
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
            Dim strMsg As String
            Dim refOpt As Short ' ��߼������(0=�O�ɒǉ� ,1=��ɒǉ�)
            Dim ret As Integer
            Try
                With m_MainEdit
                    ' �o�^������
                    If (MAXGNO <= .W_PLT.GCount) Then ' �o�^��OK ?
                        strMsg = "����ȏ�f�o�|�h�a�f�[�^�͓o�^�ł��܂���B"
                        Call MsgBox(strMsg, DirectCast( _
                                    MsgBoxStyle.OkOnly + _
                                    MsgBoxStyle.Information, MsgBoxStyle), _
                                    My.Application.Info.Title)
                        Exit Sub
                    End If

                    If (0 < .W_PLT.GCount) Then ' �o�^����1�ȏ�̏ꍇ
                        ' �m�Fү���ނ�\��("�f�o�|�h�a�f�[�^��ǉ����܂�")
                        ret = MsgBox_AddClick("�f�o�|�h�a�f�[�^", refOpt) ' ү���ޕ\��
                        If (ret <> cFRS_ERR_ADV) Then Exit Sub ' Cancel�Ȃ�Return

                        If (refOpt = 1) Then ' �\���ް��̌�ɒǉ� ?
                            m_GpibNo = (m_GpibNo + 1) ' �ǉ������ް��̓o�^�ԍ� = ���݂̓o�^�ԍ��ԍ� + 1
                        Else ' �\���ް��̑O�ɒǉ�
                            m_GpibNo = m_GpibNo ' �ǉ������ް��̓o�^�ԍ� = ���݂̓o�^�ԍ�
                        End If
                        ' �ް���1��ɂ��炵�Ēǉ�����
                        Call SortGpibData(1)

                    Else ' �o�^����0�̏ꍇ
                        With m_CtlGpib(GPIB_GNAM) ' �@�햼���͊m�F
                            If ("" = .Text) OrElse (.Text Is Nothing) Then
                                .Select()
                                .BackColor = Color.Yellow
                                strMsg = DirectCast(m_CtlGpib(GPIB_GNAM), cTxt_).GetStrMsg & "�̓��͂������Ȃ��Ă��������B"
                                Call MsgBox(strMsg, DirectCast( _
                                            MsgBoxStyle.OkOnly + _
                                            MsgBoxStyle.Information, MsgBoxStyle), _
                                            My.Application.Info.Title)
                                Exit Sub
                            End If
                        End With

                        strMsg = "�f�o�|�h�a�f�[�^��o�^���܂����H"
                        If (MsgBoxResult.Ok = MsgBox(strMsg, DirectCast( _
                                                    MsgBoxStyle.OkCancel + _
                                                    MsgBoxStyle.Information, MsgBoxStyle), _
                                                    My.Application.Info.Title)) Then
                            .W_PLT.GCount = 1 ' �o�^����ݒ�
                            m_GpibNo = 1

                            ' GP-IB�ް�����ʍ��ڂɐݒ�
                            Call SetDataToText()
                            FIRST_CONTROL.Select() ' ̵����ݒ�
                        End If

                    End If
                End With

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
                If (0 = m_MainEdit.W_PLT.GCount) Then Exit Sub ' �o�^��0�Ȃ�NOP
                strMsg = "���݂̂f�o�|�h�a�f�[�^���폜���܂��B��낵���ł����H"
                ret = MsgBox(strMsg, DirectCast( _
                            MsgBoxStyle.OkCancel + _
                            MsgBoxStyle.Exclamation + _
                            MsgBoxStyle.DefaultButton2, MsgBoxStyle), _
                            My.Application.Info.Title)
                If (ret = MsgBoxResult.Cancel) Then Exit Sub ' Cancel(RESET��) ?

                ' �����ް���1�O�ɂ߂�
                Call SortGpibData(-1)

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

        ''' <summary>�ضް�����÷���ޯ����̫�����������(�ύX���ꂽ�\��������)���ɂ����Ȃ�����</summary>
        ''' <param name="sender">�ضް�����÷���ޯ��</param>
        ''' <param name="e"></param>
        Private Sub CTxt_7_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CTxt_7.Leave
            ' ��ނ܂��͓o�^�ԍ���ؑւ����ꍇ�ASetGpibData()�̏����ɂ��
            ' m_TrgCmdFlg �� �ضް����ނ���(�����)=True,�Ȃ�(�d��)=False �ƂȂ�
            ' �ǉ��܂��͍폜���ꂽ�ꍇ�Am_TrgCmdFlg �� �ضް����ނȂ�(�d��)=False �ƂȂ�
            Dim txt As String
            Try
                txt = DirectCast(sender, cTxt_).Text
                If (0 < m_MainEdit.W_PLT.GCount) Then ' �o�^������ꍇ
                    If (False = m_TrgCmdFlg) AndAlso ("" <> txt) Then ' �ضް����ނȂ��̏�Ԃ�����͂��������ꍇ
                        ' �g�p���̓o�^�ԍ��͕ς��Ȃ����߁A�׸ނ̂ݐؑւ���
                        m_TrgCmdFlg = True
                        Exit Sub
                    ElseIf (True = m_TrgCmdFlg) AndAlso ("" = txt) Then ' �ضް����ނ���̏�Ԃ���폜���ꂽ�ꍇ
                        ' �ضް����ނ��폜���ꂽ�o�^�ԍ����g�p���Ă����R�ް��̒l��0(���������)�ɂ���
                        Call ResetResCutData(-1, True)
                        m_TrgCmdFlg = False
                        Exit Sub
                    Else
                        ' DO NOTHING
                    End If
                End If

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

    End Class
End Namespace
