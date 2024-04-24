Option Explicit On
Option Strict Off

Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports LaserFront.Trimmer.DefWin32Fnc

Namespace FormEdit
    Friend Class tabUser2
        Inherits tabBase

#Region "�錾"
        Private Const SYS_ZON As Integer = 11       ' ���������Ŏg�p(m_CtlSystem�ł̲��ޯ��)
        Private Const SYS_ZOFF As Integer = 12      ' ���������Ŏg�p(m_CtlSystem�ł̲��ޯ��)

        Private m_CtlSystem() As Control            ' USER�̺��۰ٔz��
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
                TAB_NAME = GetPrivateProfileString_S("VOLT_LABEL", "TAB_NAM", m_sPath, "????")

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
                            "VOLT_LABEL", (i.ToString("000") & "_GRP"), m_sPath, "????")
                    End With
                Next i

                ' ----------------------------------------------------------
                ' EDIT_DEF_User.ini�������قɕ\������ݒ�
                ' ----------------------------------------------------------
                LblArray = New cLbl_() { _
                    CLbl_0, CLbl_1, CLbl_2, _
                    CLbl_3, CLbl_4, CLbl_5 _
                }
                For i As Integer = 0 To (LblArray.Length - 1) Step 1
                    LblArray(i).Text = GetPrivateProfileString_S( _
                        "VOLT_LABEL", (i.ToString("000") & "_LBL"), m_sPath, "????")
                Next i

                ' ----------------------------------------------------------
                ' ���Ѹ�ٰ���ޯ�����̺��۰ق�ݒ�(CtlArray�̏��Ԃƍ��킹��)
                ' ----------------------------------------------------------
                m_CtlSystem = New Control() { _
                    CTxt_0, CTxt_1, CTxt_2, CTxt_3, CTxt_4, CTxt_5 _
                }
                Call SetControlData(m_CtlSystem)

              

                ' ----------------------------------------------------------
                ' ���ׂĂ�÷���ޯ��������ޯ������݂ɑ΂��A̫����ړ��ݒ�������Ȃ�
                ' �g�p���Ȃ����۰ق� Enabled=False �܂��� Visible=False �ɂ���
                ' ----------------------------------------------------------
                CtlArray = New Control() { _
                    CTxt_0, CTxt_1, CTxt_2, CTxt_3, CTxt_4, CTxt_5 _
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
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                        Case 4 ' �␳�l��ٰ���ޯ��
                            Select Case (tag)
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
                        strMsg = GetPrivateProfileString_S("VOLT_VALUE", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0  ' ��i
                                strMin = GetPrivateProfileString_S("VOLT_VALUE", (no & "_MIN"), m_sPath, "0.001")
                                strMax = GetPrivateProfileString_S("VOLT_VALUE", (no & "_MAX"), m_sPath, "2.000")
                            Case 1  ' ��i�d���̔{��
                                strMin = GetPrivateProfileString_S("VOLT_VALUE", (no & "_MIN"), m_sPath, "0.01")
                                strMax = GetPrivateProfileString_S("VOLT_VALUE", (no & "_MAX"), m_sPath, "10.00")
                            Case 2  ' ��R�̌�
                                strMin = GetPrivateProfileString_S("VOLT_VALUE", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("VOLT_VALUE", (no & "_MAX"), m_sPath, "10")
                            Case 3  ' �d������
                                strMin = GetPrivateProfileString_S("VOLT_VALUE", (no & "_MIN"), m_sPath, "0.01")
                                strMax = GetPrivateProfileString_S("VOLT_VALUE", (no & "_MAX"), m_sPath, "10.00")
                            Case 4 ' ����b��
                                strMin = GetPrivateProfileString_S("VOLT_VALUE", (no & "_MIN"), m_sPath, "0.01")
                                strMax = GetPrivateProfileString_S("VOLT_VALUE", (no & "_MAX"), m_sPath, "60.00")
                            Case 5 ' �ω���
                                strMin = GetPrivateProfileString_S("VOLT_VALUE", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("VOLT_VALUE", (no & "_MAX"), m_sPath, "9999")
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
                            Case 0  ' ��i
                                m_CtlSystem(i).Text = (.dRated).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 1  ' ��i�d���̔{��
                                m_CtlSystem(i).Text = (.dMagnification).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 2  ' ��R�̌�
                                m_CtlSystem(i).Text = (.dResNumber).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 3  ' �d������
                                m_CtlSystem(i).Text = (.dCurrentLimit).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 4 ' ����b��
                                m_CtlSystem(i).Text = (.dAppliedSecond).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 5 ' �ω���
                                m_CtlSystem(i).Text = (.dVariation).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())

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
                m_MainEdit.MTab.SelectedIndex = m_TabIdx ' ��ޕ\���ؑ�

                ' ���������ް�����۰قɾ�Ă���
                Call SetDataToText()

                ' ���Ѹ�ٰ���ޯ��
                ret = CheckControlData(m_CtlSystem)
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

                        Case 0
                            With .W_stUserData
                                Select Case (tag)
                                    Case 0 ' ��i
                                        ret = CheckDoubleData(cTextBox, .dRated)
                                    Case 1 ' ��i�d���̔{��
                                        ret = CheckDoubleData(cTextBox, .dMagnification)
                                    Case 2 ' ��R��
                                        ret = CheckIntData(cTextBox, .dResNumber)
                                    Case 3 ' �d������
                                        ret = CheckDoubleData(cTextBox, .dCurrentLimit)
                                    Case 4 ' ����b��
                                        ret = CheckDoubleData(cTextBox, .dAppliedSecond)
                                    Case 5 ' �ω���
                                        ret = CheckDoubleData(cTextBox, .dVariation)
                                    Case Else
                                        Throw New Exception("Case " & tag & ": Nothing")
                                End Select
                            End With
                            ' ------------------------------------------------------------------------------
                        Case Else
                            Throw New Exception("Parent.Tag Case " & cTextBox.Parent.Tag & ": Nothing")
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
