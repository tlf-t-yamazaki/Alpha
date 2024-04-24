Option Explicit On
Option Strict Off

Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports LaserFront.Trimmer.DefWin32Fnc

Namespace FormEdit
    Friend Class tabPattern
        Inherits tabBase

#Region "�錾"
        Private GRP_MIN As Integer              ' ��Ĉʒu�␳��ٰ�ߔԍ��ŏ��l
        Private GRP_MAX As Integer              ' ��Ĉʒu�␳��ٰ�ߔԍ��ő�l
        Private PTN_MIN As Integer              ' ��Ĉʒu�␳����ݔԍ��ŏ��l
        Private PTN_MAX As Integer              ' ��Ĉʒu�␳����ݔԍ��ő�l

        Private m_CtlTheta() As Control         ' �ƕ␳��ٰ���ޯ���̺��۰ٔz��
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
                TAB_NAME = GetPrivateProfileString_S("PATTERN_LABEL", "TAB_NAM", m_sPath, "????")

                ' ��Ĉʒu�␳��ٰ�ߔԍ������ݔԍ��̏㉺���l
                GRP_MIN = Convert.ToInt32(GetPrivateProfileString_S("PATTERN_LABEL", "GRP_MIN", m_sPath, "1"))
                GRP_MAX = Convert.ToInt32(GetPrivateProfileString_S("PATTERN_LABEL", "GRP_MAX", m_sPath, "999"))
                PTN_MIN = Convert.ToInt32(GetPrivateProfileString_S("PATTERN_LABEL", "PTN_MIN", m_sPath, "1"))
                PTN_MAX = Convert.ToInt32(GetPrivateProfileString_S("PATTERN_LABEL", "PTN_MAX", m_sPath, "50"))

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
                            "PATTERN_LABEL", (i.ToString("000") & "_GRP"), m_sPath, "????")
                    End With
                Next i
                ' �ǉ���폜���݂�����(����q�̂��ߐݒ肵�Ȃ�)
                'CPnl_Btn.TabIndex = 254 ' ���۰ٔz�u�\�ő吔(�Ō�ɐݒ�)

                ' ----------------------------------------------------------
                ' EDIT_DEF_User.ini�������قɕ\������ݒ�
                ' ----------------------------------------------------------
                LblArray = New cLbl_() { _
                    CLbl_0, CLbl_1, CLbl_2, CLbl_3, CLbl_4, _
                    CLbl_5, CLbl_6, CLbl_7, CLbl_8 _
                }
                For i As Integer = 0 To (LblArray.Length - 1) Step 1
                    LblArray(i).Text = GetPrivateProfileString_S( _
                        "PATTERN_LABEL", (i.ToString("000") & "_LBL"), m_sPath, "????")
                Next i

                ' ----------------------------------------------------------
                ' �ƕ␳��ٰ���ޯ�����̺��۰ق�ݒ�(CtlArray�̏��Ԃƍ��킹��)
                ' ----------------------------------------------------------
                m_CtlTheta = New Control() { _
                    CCmb_0, CCmb_1, CCmb_2, CCmb_3, CTxt_0, CTxt_1, _
                    CTxt_2, CTxt_3, CTxt_4, CCmb_4, CTxt_5, CTxt_6 _
                }
                Call SetControlData(m_CtlTheta)

                ' ----------------------------------------------------------
                ' ���ׂĂ�÷���ޯ��������ޯ������݂ɑ΂��A̫����ړ��ݒ�������Ȃ�
                ' �g�p���Ȃ����۰ق� Enabled=False �܂��� Visible=False �ɂ���
                ' ----------------------------------------------------------
                CtlArray = New Control() { _
                    CCmb_0, CCmb_1, CCmb_2, CCmb_3, CTxt_0, CTxt_1, _
                    CTxt_2, CTxt_3, CTxt_4, CCmb_4, CTxt_5, CTxt_6 _
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
                        Case 0 ' �ƕ␳
                            Select Case (tag)
                                Case 0 ' �ʒu�␳Ӱ��
                                    .Items.Add(("����"))
                                    .Items.Add(("�蓮"))
                                    .Items.Add(("����+����"))
                                Case 1 ' �ʒu�␳���@
                                    .Items.Add(("�␳�Ȃ�"))
                                    .Items.Add(("�␳����"))
                                Case 2 ' ��ٰ�ߔԍ�(1-999)
                                    For i As Integer = GRP_MIN To GRP_MAX
                                        .Items.Add(String.Format("{0,5:##0}", i))
                                    Next i
                                Case 3 ' ����ݔԍ��P(1-50)
                                    For i As Integer = PTN_MIN To PTN_MAX
                                        .Items.Add(String.Format("{0,5:#0}", i))
                                    Next i
                                Case 4 ' ����ݔԍ�2(1-50)
                                    For i As Integer = PTN_MIN To PTN_MAX
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
        ''' <summary>����������÷���ޯ���̏㉺���l�ү���ސݒ�������Ȃ�</summary>
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
                Select Case (DirectCast(cTextBox.Parent.Tag, Integer))
                    ' ------------------------------------------------------------------------------
                    Case 0 ' �ƕ␳
                        strMsg = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0  ' ����ݍ��W1X
                                strMin = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MAX"), m_sPath, "245.0")
                            Case 1  ' ����ݍ��W1Y
                                strMin = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MAX"), m_sPath, "245.0")
                            Case 2 ' �␳�߼޼�ݵ̾��X
                                strMin = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MAX"), m_sPath, "80.0") ' TODO: �㉺���l�m�F �ƕ␳
                            Case 3 ' �␳�߼޼�ݵ̾��Y
                                strMin = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MAX"), m_sPath, "80.0") ' TODO: �㉺���l�m�F �ƕ␳
                            Case 4  ' �Ɗp�x
                                strMin = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MIN"), m_sPath, "-5")
                                strMax = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MAX"), m_sPath, "5")
                            Case 5  ' ����ݍ��W1X
                                strMin = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MAX"), m_sPath, "245.0")
                            Case 6  ' ����ݍ��W2Y
                                strMin = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MAX"), m_sPath, "245.0")
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
                    Call .SetStrTip(strMin & "�`" & strMax & "�͈̔͂Ŏw�肵�ĉ�����") ' °�����ү���ނ̐ݒ�
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
                ' �ƕ␳��ٰ���ޯ���ݒ�
                Call SetThetaData()

                Me.Refresh()

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

#Region "�ƕ␳��ٰ���ޯ�����̐ݒ�"
        ''' <summary>�ƕ␳��ٰ���ޯ������÷���ޯ��������ޯ���ɒl��ݒ肷��</summary>
        Private Sub SetThetaData()
            Try
                With m_MainEdit.W_THE
                    For i As Integer = 0 To (m_CtlTheta.Length - 1) Step 1
                        Select Case (i)
                            Case 0 ' �ʒu�␳Ӱ��(0:����, 1:�蓮)
                                Dim cCombo As cCmb_ = DirectCast(m_CtlTheta(i), cCmb_)
                                Call NoEventIndexChange(cCombo, .iPP30)
                            Case 1 ' �ʒu�␳���@(0:�␳�Ȃ�, 1:�␳����)
                                Dim cCombo As cCmb_ = DirectCast(m_CtlTheta(i), cCmb_)
                                Call NoEventIndexChange(cCombo, .iPP31)
                            Case 2 ' ��ٰ�ߔԍ�(1-999)
                                Dim cCombo As cCmb_ = DirectCast(m_CtlTheta(i), cCmb_)
                                Call NoEventIndexChange(cCombo, (.iPP38 - 1))
                            Case 3 ' ����ݔԍ�1(1-50)
                                Dim cCombo As cCmb_ = DirectCast(m_CtlTheta(i), cCmb_)
                                Call NoEventIndexChange(cCombo, (.iPP37_1 - 1))
                            Case 4 ' ����݈ʒu1X
                                m_CtlTheta(i).Text = (.fpp32_x).ToString(DirectCast(m_CtlTheta(i), cTxt_).GetStrFormat())
                            Case 5 ' ����݈ʒu1Y
                                m_CtlTheta(i).Text = (.fpp32_y).ToString(DirectCast(m_CtlTheta(i), cTxt_).GetStrFormat())
                            Case 6 ' �␳�߼޼�ݵ̾��X
                                m_CtlTheta(i).Text = (.fpp34_x).ToString(DirectCast(m_CtlTheta(i), cTxt_).GetStrFormat())
                            Case 7 ' �␳�߼޼�ݵ̾��Y
                                m_CtlTheta(i).Text = (.fpp34_y).ToString(DirectCast(m_CtlTheta(i), cTxt_).GetStrFormat())
                            Case 8 ' �摜�F���p�x�␳
                                m_CtlTheta(i).Text = (.fTheta).ToString(DirectCast(m_CtlTheta(i), cTxt_).GetStrFormat())
                            Case 9 ' ����ݔԍ�2(1-50)
                                Dim cCombo As cCmb_ = DirectCast(m_CtlTheta(i), cCmb_)
                                Call NoEventIndexChange(cCombo, (.iPP37_2 - 1))
                            Case 10 ' ����݈ʒu2X
                                m_CtlTheta(i).Text = (.fpp33_x).ToString(DirectCast(m_CtlTheta(i), cTxt_).GetStrFormat())
                            Case 11 ' ����݈ʒu2Y
                                m_CtlTheta(i).Text = (.fpp33_y).ToString(DirectCast(m_CtlTheta(i), cTxt_).GetStrFormat())
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
#End Region

#Region "���ׂĂ�÷���ޯ�����ް������������Ȃ�"
        ''' <summary>���ׂĂ�÷���ޯ�����ް������������Ȃ�</summary>
        ''' <returns>0=����, 1=�װ</returns>
        Friend Overrides Function CheckAllTextData() As Integer
            Dim ret As Integer
            Try
                m_CheckFlg = True ' ������(tabBase_Layout�ɂĎg�p)
                m_MainEdit.MTab.SelectedIndex = m_TabIdx  ' ��ޕ\���ؑ�

                ' ���������ް�����۰قɾ�Ă���
                Call SetDataToText()
                Call CheckControlData(m_CtlTheta)
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

#Region "÷���ޯ�����ް������֐����Ăяo��"
        ''' <summary>÷���ޯ�����ް������֐����Ăяo��</summary>
        ''' <param name="cTextBox">��������÷���ޯ��</param>
        ''' <returns>0=����, 1=�װ</returns>
        Protected Overrides Function CheckTextData(ByRef cTextBox As cTxt_) As Integer
            Dim tag As Integer
            Dim ret As Integer
            Try
                tag = DirectCast(cTextBox.Tag, Integer)
                With m_MainEdit
                    Select Case (DirectCast(cTextBox.Parent.Tag, Integer)) ' ��ٰ���ޯ�������
                        ' ------------------------------------------------------------------------------
                        Case 0 ' �ƕ␳��ٰ���ޯ��
                            With .W_THE
                                Select Case (tag)
                                    Case 0 ' �����1���WX
                                        ret = CheckDoubleData(cTextBox, .fpp32_x)
                                    Case 1 ' �����1���WY
                                        ret = CheckDoubleData(cTextBox, .fpp32_y)
                                    Case 2 ' �␳�߼޼�ݵ̾��X
                                        ret = CheckDoubleData(cTextBox, .fpp34_x)
                                    Case 3 ' �␳�߼޼�ݵ̾��Y
                                        ret = CheckDoubleData(cTextBox, .fpp34_y)
                                    Case 4 ' �Ǝ��p�x
                                        ret = CheckDoubleData(cTextBox, .fTheta)
                                    Case 5 ' �����2���WX
                                        ret = CheckDoubleData(cTextBox, .fpp33_x)
                                    Case 6 ' �����2���WY
                                        ret = CheckDoubleData(cTextBox, .fpp33_y)
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
                        Case 0 ' �ƕ␳
                            Select Case (tag)
                                Case 0 ' �ʒu�␳Ӱ��(0:����, 1:�蓮)
                                    .W_THE.iPP30 = Convert.ToInt16(idx)
                                Case 1 ' �ʒu�␳���@(0:�␳�Ȃ�, 1:�␳����)
                                    .W_THE.iPP31 = Convert.ToInt16(idx)
                                Case 2 ' ��ٰ�ߔԍ�(1-999)
                                    .W_THE.iPP38 = Convert.ToInt16(idx + 1)
                                Case 3 ' ����ݔԍ�1(1-50)
                                    .W_THE.iPP37_1 = Convert.ToInt16(idx + 1)
                                Case 4 ' ����ݔԍ�2(1-50)
                                    .W_THE.iPP37_2 = Convert.ToInt16(idx + 1)
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
#End Region

    End Class
End Namespace
