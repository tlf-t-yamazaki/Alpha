Option Explicit On
Option Strict Off

Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports LaserFront.Trimmer.DefWin32Fnc

Namespace FormEdit
    Friend Class tabUser2
        Inherits tabBase

#Region "宣言"
        Private Const SYS_ZON As Integer = 11       ' 相関ﾁｪｯｸで使用(m_CtlSystemでのｲﾝﾃﾞｯｸｽ)
        Private Const SYS_ZOFF As Integer = 12      ' 相関ﾁｪｯｸで使用(m_CtlSystemでのｲﾝﾃﾞｯｸｽ)

        Private m_CtlSystem() As Control            ' USERのｺﾝﾄﾛｰﾙ配列
#End Region

#Region "ｺﾝｽﾄﾗｸﾀ"
        ''' <summary>ｺﾝｽﾄﾗｸﾀ</summary>
        Friend Sub New(ByRef mainEdit As frmEdit, ByVal tabIdx As Integer)
            ' この呼び出しは、Windows フォーム デザイナで必要です。
            InitializeComponent()

            ' InitializeComponent() 呼び出しの後で初期化を追加します。
            Call InitAllControl(mainEdit, tabIdx)
        End Sub
#End Region

#Region "初期化処理"
        ''' <summary>ｺﾝﾄﾛｰﾙ初期化処理</summary>
        ''' <param name="mainEdit">ﾒｲﾝ編集画面への参照</param>
        ''' <param name="tabIdx">ﾒｲﾝﾀﾌﾞｺﾝﾄﾛｰﾙ上のｲﾝﾃﾞｯｸｽ</param>
        Protected Overrides Sub InitAllControl(ByRef mainEdit As frmEdit, ByVal tabIdx As Integer)
            Dim GrpArray() As cGrp_     ' ｸﾞﾙｰﾌﾟﾎﾞｯｸｽの表示設定で使用する
            Dim LblArray() As cLbl_     ' ﾗﾍﾞﾙへの表示設定で使用する
            Dim CtlArray() As Control   ' ｶｰｿﾙｷｰによるﾌｫｰｶｽ移動で使用する

            m_MainEdit = mainEdit       ' ﾒｲﾝ編集画面への参照を設定
            m_TabIdx = tabIdx           ' ﾒｲﾝ編集画面ﾀﾌﾞｺﾝﾄﾛｰﾙ上でのｲﾝﾃﾞｯｸｽ

            Try
                ' EDIT_DEF_User.iniからﾀﾌﾞ名を設定
                TAB_NAME = GetPrivateProfileString_S("VOLT_LABEL", "TAB_NAM", m_sPath, "????")

                ' ----------------------------------------------------------
                ' EDIT_DEF_User.iniからｸﾞﾙｰﾌﾟﾎﾞｯｸｽに表示名を設定
                ' ----------------------------------------------------------
                GrpArray = New cGrp_() { _
                    CGrp_0 _
                }
                For i As Integer = 0 To (GrpArray.Length - 1) Step 1
                    With GrpArray(i)
                        .TabIndex = ((254 - GrpArray.Length) + i) ' ｶｰｿﾙｷｰによるﾌｫｰｶｽ移動で必要
                        .Tag = i
                        .Text = GetPrivateProfileString_S( _
                            "VOLT_LABEL", (i.ToString("000") & "_GRP"), m_sPath, "????")
                    End With
                Next i

                ' ----------------------------------------------------------
                ' EDIT_DEF_User.iniからﾗﾍﾞﾙに表示名を設定
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
                ' ｼｽﾃﾑｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のｺﾝﾄﾛｰﾙを設定(CtlArrayの順番と合わせる)
                ' ----------------------------------------------------------
                m_CtlSystem = New Control() { _
                    CTxt_0, CTxt_1, CTxt_2, CTxt_3, CTxt_4, CTxt_5 _
                }
                Call SetControlData(m_CtlSystem)

              

                ' ----------------------------------------------------------
                ' すべてのﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽ･ﾎﾞﾀﾝに対し、ﾌｫｰｶｽ移動設定をおこなう
                ' 使用しないｺﾝﾄﾛｰﾙは Enabled=False または Visible=False にする
                ' ----------------------------------------------------------
                CtlArray = New Control() { _
                    CTxt_0, CTxt_1, CTxt_2, CTxt_3, CTxt_4, CTxt_5 _
                }
                Call SetTabIndex(CtlArray) ' ﾀﾌﾞｲﾝﾃﾞｯｸｽとKeyDownｲﾍﾞﾝﾄを設定する

                ' ----------------------------------------------------------
                ' 画面表示時にﾌｫｰｶｽされるｺﾝﾄﾛｰﾙを設定する
                ' ----------------------------------------------------------
                FIRST_CONTROL = CtlArray(0)

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#Region "初期化時にｺﾝﾎﾞﾎﾞｯｸｽの設定をおこなう"
        ''' <summary>初期化時にｺﾝﾎﾞﾎﾞｯｸｽの設定をおこなう</summary>
        ''' <param name="cCombo">設定をおこなうｺﾝﾎﾞﾎﾞｯｸｽ</param>
        Protected Overrides Sub InitComboBox(ByRef cCombo As cCmb_)
            Dim tag As Integer
            Try
                With cCombo
                    tag = DirectCast(.Tag, Integer)
                    .Items.Clear()
                    Select Case (DirectCast(.Parent.Tag, Integer)) ' ｸﾞﾙｰﾌﾟﾎﾞｯｸｽのﾀｸﾞ
                        ' ------------------------------------------------------------------------------
                        Case 0 To 3 ' ﾌﾞﾛｯｸ,ﾛｯﾄ情報,共通設定,温度ｾﾝｻｰｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            Select Case (tag)
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                        Case 4 ' 補正値ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            Select Case (tag)
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                        Case Else
                            Throw New Exception("Parent.Tag - Case Else")
                    End Select

                    '.SelectedIndex = 0
                    Call .SetStrTip("ドロップダウンリストから選択してください") ' ﾂｰﾙﾁｯﾌﾟﾒｯｾｰｼﾞの設定
                    Call .SetLblToolTip(m_MainEdit.LblToolTip) ' ﾒｲﾝ編集画面のﾂｰﾙﾁｯﾌﾟ参照設定
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex, cCombo.Name)
            End Try

        End Sub
#End Region

#Region "初期化時にﾃｷｽﾄﾎﾞｯｸｽの設定をおこなう"
        ''' <summary>初期化時にﾃｷｽﾄﾎﾞｯｸｽの設定をおこなう</summary>
        ''' <param name="cTextBox">設定をおこなうﾃｷｽﾄﾎﾞｯｸｽ</param>
        Protected Overrides Sub InitTextBox(ByRef cTextBox As cTxt_)
            Dim strMin As String = ""           ' 設定する変数の最大値
            Dim strMax As String = ""           ' 設定する変数の最小値
            Dim strMsg As String = ""           ' ｴﾗｰで表示する項目名
            Dim no As String = ""
            Dim tag As Integer
            Dim strFlg As Boolean = False       ' 格納する値の種類(False=数値,True=文字列)

            Try
                tag = DirectCast(cTextBox.Tag, Integer)
                no = tag.ToString("000")

                Select Case (DirectCast(cTextBox.Parent.Tag, Integer)) ' ｸﾞﾙｰﾌﾟﾎﾞｯｸｽのﾀｸﾞ
                    ' ------------------------------------------------------------------------------
                    Case 0 To 3 ' ｼｽﾃﾑｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                        ' ﾃﾞｰﾀﾁｪｯｸｴﾗｰ時の表示名
                        strMsg = GetPrivateProfileString_S("VOLT_VALUE", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0  ' 定格
                                strMin = GetPrivateProfileString_S("VOLT_VALUE", (no & "_MIN"), m_sPath, "0.001")
                                strMax = GetPrivateProfileString_S("VOLT_VALUE", (no & "_MAX"), m_sPath, "2.000")
                            Case 1  ' 定格電圧の倍率
                                strMin = GetPrivateProfileString_S("VOLT_VALUE", (no & "_MIN"), m_sPath, "0.01")
                                strMax = GetPrivateProfileString_S("VOLT_VALUE", (no & "_MAX"), m_sPath, "10.00")
                            Case 2  ' 抵抗の個数
                                strMin = GetPrivateProfileString_S("VOLT_VALUE", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("VOLT_VALUE", (no & "_MAX"), m_sPath, "10")
                            Case 3  ' 電流制限
                                strMin = GetPrivateProfileString_S("VOLT_VALUE", (no & "_MIN"), m_sPath, "0.01")
                                strMax = GetPrivateProfileString_S("VOLT_VALUE", (no & "_MAX"), m_sPath, "10.00")
                            Case 4 ' 印加秒数
                                strMin = GetPrivateProfileString_S("VOLT_VALUE", (no & "_MIN"), m_sPath, "0.01")
                                strMax = GetPrivateProfileString_S("VOLT_VALUE", (no & "_MAX"), m_sPath, "60.00")
                            Case 5 ' 変化量
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
                    Call .SetStrMsg(strMsg) ' ﾃﾞｰﾀﾁｪｯｸｴﾗｰ時の表示名設定
                    Call .SetMinMax(strMin, strMax) ' 下限値･上限値の設定
                    Dim strKind As String
                    If (False = strFlg) Then ' (False=数値,True=文字列)
                        strKind = "の範囲で指定して下さい"
                    Else
                        strKind = "文字の範囲で指定して下さい"
                        .MaxLength = Convert.ToInt32(strMax) ' SetControlData()内の条件判断で使用する
                        .TextAlign = HorizontalAlignment.Left
                    End If
                    Call .SetStrTip(strMin & "～" & strMax & strKind) ' ﾂｰﾙﾁｯﾌﾟﾒｯｾｰｼﾞの設定
                    Call .SetLblToolTip(m_MainEdit.LblToolTip) ' ﾒｲﾝ編集画面のﾂｰﾙﾁｯﾌﾟ参照設定
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex, cTextBox.Name)
            End Try

        End Sub
#End Region

#Region "ﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽに値を表示する"
        ''' <summary>ﾃｷｽﾄﾎﾞｯｸｽに値を設定する</summary>
        Protected Overrides Sub SetDataToText()
            Try
                Call SetUserData()
                Me.Refresh()

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

#Region "ｼｽﾃﾑｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内の設定"
        ''' <summary>ｼｽﾃﾑｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽに値を設定する</summary>
        Private Sub SetUserData()
            Try
                With m_MainEdit.W_stUserData
                    For i As Integer = 0 To (m_CtlSystem.Length - 1) Step 1
                        Select Case (i)
                            Case 0  ' 定格
                                m_CtlSystem(i).Text = (.dRated).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 1  ' 定格電圧の倍率
                                m_CtlSystem(i).Text = (.dMagnification).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 2  ' 抵抗の個数
                                m_CtlSystem(i).Text = (.dResNumber).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 3  ' 電流制限
                                m_CtlSystem(i).Text = (.dCurrentLimit).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 4 ' 印加秒数
                                m_CtlSystem(i).Text = (.dAppliedSecond).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 5 ' 変化量
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

#Region "すべてのﾃｷｽﾄﾎﾞｯｸｽのﾃﾞｰﾀﾁｪｯｸをおこなう"
        ''' <summary>すべてのﾃｷｽﾄﾎﾞｯｸｽのﾃﾞｰﾀﾁｪｯｸをおこなう</summary>
        ''' <returns>0=正常, 1=ｴﾗｰ</returns>
        Friend Overrides Function CheckAllTextData() As Integer
            Dim ret As Integer
            Try
                m_CheckFlg = True ' ﾁｪｯｸ中(tabBase_Layoutにて使用)
                m_MainEdit.MTab.SelectedIndex = m_TabIdx ' ﾀﾌﾞ表示切替

                ' ﾁｪｯｸするﾃﾞｰﾀをｺﾝﾄﾛｰﾙにｾｯﾄする
                Call SetDataToText()

                ' ｼｽﾃﾑｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                ret = CheckControlData(m_CtlSystem)
                If (ret <> 0) Then Exit Try

            Catch ex As Exception
                Call MsgBox_Exception(ex)
                ret = 1
            Finally
                m_CheckFlg = False ' ﾁｪｯｸ終了
                CheckAllTextData = ret
            End Try

        End Function
#End Region

#Region "ﾃﾞｰﾀﾁｪｯｸ関数を呼び出す"
        ''' <summary>ﾃｷｽﾄﾎﾞｯｸｽのﾃﾞｰﾀﾁｪｯｸ関数を呼び出す</summary>
        ''' <param name="cTextBox">ﾁｪｯｸするﾃｷｽﾄﾎﾞｯｸｽ</param>
        ''' <returns>0=正常, 1=ｴﾗｰ</returns>
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
                                    Case 0 ' 定格
                                        ret = CheckDoubleData(cTextBox, .dRated)
                                    Case 1 ' 定格電圧の倍率
                                        ret = CheckDoubleData(cTextBox, .dMagnification)
                                    Case 2 ' 抵抗個数
                                        ret = CheckIntData(cTextBox, .dResNumber)
                                    Case 3 ' 電流制限
                                        ret = CheckDoubleData(cTextBox, .dCurrentLimit)
                                    Case 4 ' 印加秒数
                                        ret = CheckDoubleData(cTextBox, .dAppliedSecond)
                                    Case 5 ' 変化量
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

#Region "ｲﾍﾞﾝﾄ"
        ''' <summary>ｺﾝﾎﾞﾎﾞｯｸｽのｲﾝﾃﾞｯｸｽが変更された時の処理</summary>
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
                    Select Case (DirectCast(cCombo.Parent.Tag, Integer)) ' ｸﾞﾙｰﾌﾟﾎﾞｯｸｽのﾀｸﾞ
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
