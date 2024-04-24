Option Explicit On
Option Strict Off

Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports LaserFront.Trimmer.DefWin32Fnc

Namespace FormEdit
    Friend Class tabPattern
        Inherits tabBase

#Region "宣言"
        Private GRP_MIN As Integer              ' ｶｯﾄ位置補正ｸﾞﾙｰﾌﾟ番号最小値
        Private GRP_MAX As Integer              ' ｶｯﾄ位置補正ｸﾞﾙｰﾌﾟ番号最大値
        Private PTN_MIN As Integer              ' ｶｯﾄ位置補正ﾊﾟﾀｰﾝ番号最小値
        Private PTN_MAX As Integer              ' ｶｯﾄ位置補正ﾊﾟﾀｰﾝ番号最大値

        Private m_CtlTheta() As Control         ' θ補正ｸﾞﾙｰﾌﾟﾎﾞｯｸｽのｺﾝﾄﾛｰﾙ配列
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

            m_TabIdx = tabIdx           ' ﾒｲﾝ編集画面ﾀﾌﾞｺﾝﾄﾛｰﾙ上でのｲﾝﾃﾞｯｸｽ
            m_MainEdit = mainEdit       ' ﾒｲﾝ編集画面への参照を設定

            Try
                ' EDIT_DEF_User.iniからﾀﾌﾞ名を設定
                TAB_NAME = GetPrivateProfileString_S("PATTERN_LABEL", "TAB_NAM", m_sPath, "????")

                ' ｶｯﾄ位置補正ｸﾞﾙｰﾌﾟ番号･ﾊﾟﾀｰﾝ番号の上下限値
                GRP_MIN = Convert.ToInt32(GetPrivateProfileString_S("PATTERN_LABEL", "GRP_MIN", m_sPath, "1"))
                GRP_MAX = Convert.ToInt32(GetPrivateProfileString_S("PATTERN_LABEL", "GRP_MAX", m_sPath, "999"))
                PTN_MIN = Convert.ToInt32(GetPrivateProfileString_S("PATTERN_LABEL", "PTN_MIN", m_sPath, "1"))
                PTN_MAX = Convert.ToInt32(GetPrivateProfileString_S("PATTERN_LABEL", "PTN_MAX", m_sPath, "50"))

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
                            "PATTERN_LABEL", (i.ToString("000") & "_GRP"), m_sPath, "????")
                    End With
                Next i
                ' 追加･削除ﾎﾞﾀﾝのﾊﾟﾈﾙ(入れ子のため設定しない)
                'CPnl_Btn.TabIndex = 254 ' ｺﾝﾄﾛｰﾙ配置可能最大数(最後に設定)

                ' ----------------------------------------------------------
                ' EDIT_DEF_User.iniからﾗﾍﾞﾙに表示名を設定
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
                ' θ補正ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のｺﾝﾄﾛｰﾙを設定(CtlArrayの順番と合わせる)
                ' ----------------------------------------------------------
                m_CtlTheta = New Control() { _
                    CCmb_0, CCmb_1, CCmb_2, CCmb_3, CTxt_0, CTxt_1, _
                    CTxt_2, CTxt_3, CTxt_4, CCmb_4, CTxt_5, CTxt_6 _
                }
                Call SetControlData(m_CtlTheta)

                ' ----------------------------------------------------------
                ' すべてのﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽ･ﾎﾞﾀﾝに対し、ﾌｫｰｶｽ移動設定をおこなう
                ' 使用しないｺﾝﾄﾛｰﾙは Enabled=False または Visible=False にする
                ' ----------------------------------------------------------
                CtlArray = New Control() { _
                    CCmb_0, CCmb_1, CCmb_2, CCmb_3, CTxt_0, CTxt_1, _
                    CTxt_2, CTxt_3, CTxt_4, CCmb_4, CTxt_5, CTxt_6 _
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
        ''' <summary>初期化時にｺﾝﾎﾞﾎﾞｯｸｽのﾘｽﾄ･ﾒｯｾｰｼﾞ設定をおこなう</summary>
        ''' <param name="cCombo">設定をおこなうｺﾝﾎﾞﾎﾞｯｸｽ</param>
        Protected Overrides Sub InitComboBox(ByRef cCombo As cCmb_)
            Dim tag As Integer
            Try
                With cCombo
                    tag = DirectCast(.Tag, Integer)
                    .Items.Clear()
                    Select Case (DirectCast(.Parent.Tag, Integer)) ' ｸﾞﾙｰﾌﾟﾎﾞｯｸｽのﾀｸﾞ
                        ' ------------------------------------------------------------------------------
                        Case 0 ' θ補正
                            Select Case (tag)
                                Case 0 ' 位置補正ﾓｰﾄﾞ
                                    .Items.Add(("自動"))
                                    .Items.Add(("手動"))
                                    .Items.Add(("自動+微調"))
                                Case 1 ' 位置補正方法
                                    .Items.Add(("補正なし"))
                                    .Items.Add(("補正あり"))
                                Case 2 ' ｸﾞﾙｰﾌﾟ番号(1-999)
                                    For i As Integer = GRP_MIN To GRP_MAX
                                        .Items.Add(String.Format("{0,5:##0}", i))
                                    Next i
                                Case 3 ' ﾊﾟﾀｰﾝ番号１(1-50)
                                    For i As Integer = PTN_MIN To PTN_MAX
                                        .Items.Add(String.Format("{0,5:#0}", i))
                                    Next i
                                Case 4 ' ﾊﾟﾀｰﾝ番号2(1-50)
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

                    Call .SetStrTip("ドロップダウンリストから選択してください") ' ﾂｰﾙﾁｯﾌﾟﾒｯｾｰｼﾞの設定
                    Call .SetLblToolTip(m_MainEdit.LblToolTip) ' ﾒｲﾝ編集画面のﾂｰﾙﾁｯﾌﾟ参照設定
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex, cCombo.Name)
            End Try

        End Sub
#End Region

#Region "初期化時にﾃｷｽﾄﾎﾞｯｸｽの設定をおこなう"
        ''' <summary>初期化時にﾃｷｽﾄﾎﾞｯｸｽの上下限値･ﾒｯｾｰｼﾞ設定をおこなう</summary>
        ''' <param name="cTextBox">設定をおこなうﾃｷｽﾄﾎﾞｯｸｽ</param>
        Protected Overrides Sub InitTextBox(ByRef cTextBox As cTxt_)
            Dim strMin As String = ""           ' 設定する変数の最大値
            Dim strMax As String = ""           ' 設定する変数の最小値
            Dim strMsg As String = ""           ' ｴﾗｰで表示する項目名
            Dim no As String = ""
            Dim tag As Integer
            Try
                tag = DirectCast(cTextBox.Tag, Integer)
                no = tag.ToString("000")
                Select Case (DirectCast(cTextBox.Parent.Tag, Integer))
                    ' ------------------------------------------------------------------------------
                    Case 0 ' θ補正
                        strMsg = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0  ' ﾊﾟﾀｰﾝ座標1X
                                strMin = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MAX"), m_sPath, "245.0")
                            Case 1  ' ﾊﾟﾀｰﾝ座標1Y
                                strMin = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MAX"), m_sPath, "245.0")
                            Case 2 ' 補正ﾎﾟｼﾞｼｮﾝｵﾌｾｯﾄX
                                strMin = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MAX"), m_sPath, "80.0") ' TODO: 上下限値確認 θ補正
                            Case 3 ' 補正ﾎﾟｼﾞｼｮﾝｵﾌｾｯﾄY
                                strMin = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MAX"), m_sPath, "80.0") ' TODO: 上下限値確認 θ補正
                            Case 4  ' θ角度
                                strMin = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MIN"), m_sPath, "-5")
                                strMax = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MAX"), m_sPath, "5")
                            Case 5  ' ﾊﾟﾀｰﾝ座標1X
                                strMin = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("PATTERN_THETA", (no & "_MAX"), m_sPath, "245.0")
                            Case 6  ' ﾊﾟﾀｰﾝ座標2Y
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
                    Call .SetStrMsg(strMsg) ' ﾃﾞｰﾀﾁｪｯｸｴﾗｰ時の表示名設定
                    Call .SetMinMax(strMin, strMax) ' 下限値･上限値の設定
                    Call .SetStrTip(strMin & "～" & strMax & "の範囲で指定して下さい") ' ﾂｰﾙﾁｯﾌﾟﾒｯｾｰｼﾞの設定
                    Call .SetLblToolTip(m_MainEdit.LblToolTip) ' ﾒｲﾝ編集画面のﾂｰﾙﾁｯﾌﾟ参照設定
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex, cTextBox.Name)
            End Try

        End Sub
#End Region

#Region "ﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽに値を表示する"
        ''' <summary>ﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽに値を表示する</summary>
        Protected Overrides Sub SetDataToText()
            Try
                ' θ補正ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ設定
                Call SetThetaData()

                Me.Refresh()

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

#Region "θ補正ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内の設定"
        ''' <summary>θ補正ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽに値を設定する</summary>
        Private Sub SetThetaData()
            Try
                With m_MainEdit.W_THE
                    For i As Integer = 0 To (m_CtlTheta.Length - 1) Step 1
                        Select Case (i)
                            Case 0 ' 位置補正ﾓｰﾄﾞ(0:自動, 1:手動)
                                Dim cCombo As cCmb_ = DirectCast(m_CtlTheta(i), cCmb_)
                                Call NoEventIndexChange(cCombo, .iPP30)
                            Case 1 ' 位置補正方法(0:補正なし, 1:補正あり)
                                Dim cCombo As cCmb_ = DirectCast(m_CtlTheta(i), cCmb_)
                                Call NoEventIndexChange(cCombo, .iPP31)
                            Case 2 ' ｸﾞﾙｰﾌﾟ番号(1-999)
                                Dim cCombo As cCmb_ = DirectCast(m_CtlTheta(i), cCmb_)
                                Call NoEventIndexChange(cCombo, (.iPP38 - 1))
                            Case 3 ' ﾊﾟﾀｰﾝ番号1(1-50)
                                Dim cCombo As cCmb_ = DirectCast(m_CtlTheta(i), cCmb_)
                                Call NoEventIndexChange(cCombo, (.iPP37_1 - 1))
                            Case 4 ' ﾊﾟﾀｰﾝ位置1X
                                m_CtlTheta(i).Text = (.fpp32_x).ToString(DirectCast(m_CtlTheta(i), cTxt_).GetStrFormat())
                            Case 5 ' ﾊﾟﾀｰﾝ位置1Y
                                m_CtlTheta(i).Text = (.fpp32_y).ToString(DirectCast(m_CtlTheta(i), cTxt_).GetStrFormat())
                            Case 6 ' 補正ﾎﾟｼﾞｼｮﾝｵﾌｾｯﾄX
                                m_CtlTheta(i).Text = (.fpp34_x).ToString(DirectCast(m_CtlTheta(i), cTxt_).GetStrFormat())
                            Case 7 ' 補正ﾎﾟｼﾞｼｮﾝｵﾌｾｯﾄY
                                m_CtlTheta(i).Text = (.fpp34_y).ToString(DirectCast(m_CtlTheta(i), cTxt_).GetStrFormat())
                            Case 8 ' 画像認識角度補正
                                m_CtlTheta(i).Text = (.fTheta).ToString(DirectCast(m_CtlTheta(i), cTxt_).GetStrFormat())
                            Case 9 ' ﾊﾟﾀｰﾝ番号2(1-50)
                                Dim cCombo As cCmb_ = DirectCast(m_CtlTheta(i), cCmb_)
                                Call NoEventIndexChange(cCombo, (.iPP37_2 - 1))
                            Case 10 ' ﾊﾟﾀｰﾝ位置2X
                                m_CtlTheta(i).Text = (.fpp33_x).ToString(DirectCast(m_CtlTheta(i), cTxt_).GetStrFormat())
                            Case 11 ' ﾊﾟﾀｰﾝ位置2Y
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

#Region "すべてのﾃｷｽﾄﾎﾞｯｸｽのﾃﾞｰﾀﾁｪｯｸをおこなう"
        ''' <summary>すべてのﾃｷｽﾄﾎﾞｯｸｽのﾃﾞｰﾀﾁｪｯｸをおこなう</summary>
        ''' <returns>0=正常, 1=ｴﾗｰ</returns>
        Friend Overrides Function CheckAllTextData() As Integer
            Dim ret As Integer
            Try
                m_CheckFlg = True ' ﾁｪｯｸ中(tabBase_Layoutにて使用)
                m_MainEdit.MTab.SelectedIndex = m_TabIdx  ' ﾀﾌﾞ表示切替

                ' ﾁｪｯｸするﾃﾞｰﾀをｺﾝﾄﾛｰﾙにｾｯﾄする
                Call SetDataToText()
                Call CheckControlData(m_CtlTheta)
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

#Region "ﾃｷｽﾄﾎﾞｯｸｽのﾃﾞｰﾀﾁｪｯｸ関数を呼び出す"
        ''' <summary>ﾃｷｽﾄﾎﾞｯｸｽのﾃﾞｰﾀﾁｪｯｸ関数を呼び出す</summary>
        ''' <param name="cTextBox">ﾁｪｯｸするﾃｷｽﾄﾎﾞｯｸｽ</param>
        ''' <returns>0=正常, 1=ｴﾗｰ</returns>
        Protected Overrides Function CheckTextData(ByRef cTextBox As cTxt_) As Integer
            Dim tag As Integer
            Dim ret As Integer
            Try
                tag = DirectCast(cTextBox.Tag, Integer)
                With m_MainEdit
                    Select Case (DirectCast(cTextBox.Parent.Tag, Integer)) ' ｸﾞﾙｰﾌﾟﾎﾞｯｸｽのﾀｸﾞ
                        ' ------------------------------------------------------------------------------
                        Case 0 ' θ補正ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            With .W_THE
                                Select Case (tag)
                                    Case 0 ' ﾊﾟﾀｰﾝ1座標X
                                        ret = CheckDoubleData(cTextBox, .fpp32_x)
                                    Case 1 ' ﾊﾟﾀｰﾝ1座標Y
                                        ret = CheckDoubleData(cTextBox, .fpp32_y)
                                    Case 2 ' 補正ﾎﾟｼﾞｼｮﾝｵﾌｾｯﾄX
                                        ret = CheckDoubleData(cTextBox, .fpp34_x)
                                    Case 3 ' 補正ﾎﾟｼﾞｼｮﾝｵﾌｾｯﾄY
                                        ret = CheckDoubleData(cTextBox, .fpp34_y)
                                    Case 4 ' θ軸角度
                                        ret = CheckDoubleData(cTextBox, .fTheta)
                                    Case 5 ' ﾊﾟﾀｰﾝ2座標X
                                        ret = CheckDoubleData(cTextBox, .fpp33_x)
                                    Case 6 ' ﾊﾟﾀｰﾝ2座標Y
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
                        Case 0 ' θ補正
                            Select Case (tag)
                                Case 0 ' 位置補正ﾓｰﾄﾞ(0:自動, 1:手動)
                                    .W_THE.iPP30 = Convert.ToInt16(idx)
                                Case 1 ' 位置補正方法(0:補正なし, 1:補正あり)
                                    .W_THE.iPP31 = Convert.ToInt16(idx)
                                Case 2 ' ｸﾞﾙｰﾌﾟ番号(1-999)
                                    .W_THE.iPP38 = Convert.ToInt16(idx + 1)
                                Case 3 ' ﾊﾟﾀｰﾝ番号1(1-50)
                                    .W_THE.iPP37_1 = Convert.ToInt16(idx + 1)
                                Case 4 ' ﾊﾟﾀｰﾝ番号2(1-50)
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
