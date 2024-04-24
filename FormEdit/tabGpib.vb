Option Explicit On
Option Strict Off

Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports LaserFront.Trimmer.DefWin32Fnc

Namespace FormEdit
    Friend Class tabGpib
        Inherits tabBase

#Region "宣言"
        Private Const GPIB_GNAM As Integer = 3  ' m_CtlGpibでのｲﾝﾃﾞｯｸｽ(機器名)
        Private Const GPIB_CON As Integer = 7   ' m_CtlGpibでのｲﾝﾃﾞｯｸｽ(ONｺﾏﾝﾄﾞ)
        Private Const GPIB_COFF As Integer = 9  ' m_CtlGpibでのｲﾝﾃﾞｯｸｽ(OFFｺﾏﾝﾄﾞ)
        Private Const GPIB_CTRG As Integer = 11 ' m_CtlGpibでのｲﾝﾃﾞｯｸｽ(ﾄﾘｶﾞｰｺﾏﾝﾄﾞ)

        Private m_CtlGpib() As Control          ' GP-IBｸﾞﾙｰﾌﾟﾎﾞｯｸｽのｺﾝﾄﾛｰﾙ配列
        Private m_TrgCmdFlg As Boolean          ' ﾄﾘｶﾞｰｺﾏﾝﾄﾞ あり(測定器)=True,なし(電源)=False
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
                TAB_NAME = GetPrivateProfileString_S("GPIB_LABEL", "TAB_NAM", m_sPath, "????")

                ' 追加･削除ﾎﾞﾀﾝの設定
                With mainEdit
                    CBtn_Add.SetLblToolTip(.LblToolTip)
                    CBtn_Add.Text = GetPrivateProfileString_S("BUTTON_LABEL", "BTN_ADD", m_sPath, "ADD")
                    CBtn_Del.SetLblToolTip(.LblToolTip)
                    CBtn_Del.Text = GetPrivateProfileString_S("BUTTON_LABEL", "BTN_DEL", m_sPath, "DEL")
                End With

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
                            "GPIB_LABEL", (i.ToString("000") & "_GRP"), m_sPath, "????")
                    End With
                Next i
                ' 追加･削除ﾎﾞﾀﾝのﾊﾟﾈﾙ
                CPnl_Btn.TabIndex = 254 ' ｺﾝﾄﾛｰﾙ配置可能最大数(最後に設定)

                ' ----------------------------------------------------------
                ' EDIT_DEF_User.iniからﾗﾍﾞﾙに表示名を設定
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
                ' GP-IBｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のｺﾝﾄﾛｰﾙを設定
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
                ' すべてのﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽ･ﾎﾞﾀﾝに対し、ﾌｫｰｶｽ移動設定をおこなう
                ' ﾀﾌﾞｷｰ、ｶｰｿﾙｷｰによりﾌｫｰｶｽ移動する順番でｺﾝﾄﾛｰﾙをCtlArrayに設定する
                ' 使用しないｺﾝﾄﾛｰﾙは Enabled=False または Visible=False にする
                ' ----------------------------------------------------------
                CtlArray = New Control() { _
                    CCmb_0, CTxt_0, CCmb_1, CTxt_1, _
                    CTxt_2, CTxt_8, CTxt_9, _
                    CTxt_3, CTxt_4, _
                    CTxt_5, CTxt_6, _
                    CTxt_7, _
                    CBtn_Add, CBtn_Del _
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
                        Case 0 ' GP-IBｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            Select Case (tag)
                                Case 0 ' 登録番号
                                    '.Items.Add("") ' ﾚｲｱｳﾄｲﾍﾞﾝﾄで再設定される
                                Case 1 ' ﾃﾞﾘﾐﾀ(0:CRLF, 1:CR, 2:LF, 3:なし)
                                    .Items.Add("なし")                      'V2.1.0.0④
                                    .Items.Add("CRLF")                      'V2.1.0.0④
                                    .Items.Add("CR")                        'V2.1.0.0④
                                    .Items.Add("LF")                        'V2.1.0.0④
                                    'V2.1.0.0④                                    .Items.Add("CRLF")
                                    'V2.1.0.0④                                    .Items.Add("LF")
                                    'V2.1.0.0④                                    .Items.Add("なし")
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

            Dim strFlg As Boolean = False       ' 格納する値の種類(False=数値,True=文字列)
            Try
                tag = DirectCast(cTextBox.Tag, Integer)
                no = tag.ToString("000")

                Select Case (DirectCast(cTextBox.Parent.Tag, Integer)) ' ｸﾞﾙｰﾌﾟﾎﾞｯｸｽのﾀｸﾞ
                    ' ------------------------------------------------------------------------------
                    Case 0 ' GP-IBｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                        ' ﾃﾞｰﾀﾁｪｯｸｴﾗｰ時の表示名
                        strMsg = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0 ' ｱﾄﾞﾚｽ
                                strMin = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MAX"), m_sPath, "30")
                            Case 1 ' 機器名
                                strMin = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MAX"), m_sPath, "20")
                                strFlg = True
                            Case 2 ' 設定ｺﾏﾝﾄﾞ
                                strMin = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MAX"), m_sPath, "100")
                                strFlg = True
                            Case 3 ' 設定ｺﾏﾝﾄﾞ(2段目)
                                strMin = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MAX"), m_sPath, "100")
                                strFlg = True
                            Case 4 ' 設定ｺﾏﾝﾄﾞ(3段目)
                                strMin = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MAX"), m_sPath, "100")
                                strFlg = True
                            Case 5 ' ONｺﾏﾝﾄﾞ
                                strMin = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MAX"), m_sPath, "50")
                                strFlg = True
                            Case 6 ' ON後のﾎﾟｰｽﾞ時間
                                strMin = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MAX"), m_sPath, "32767")
                            Case 7 ' OFFｺﾏﾝﾄﾞ
                                strMin = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MAX"), m_sPath, "50")
                                strFlg = True
                            Case 8 ' OFF後のﾎﾟｰｽﾞ時間
                                strMin = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("GPIB_GPIB", (no & "_MAX"), m_sPath, "32767")
                            Case 9 ' ﾄﾘｶﾞｰｺﾏﾝﾄﾞ
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

#Region "ﾃｷｽﾄﾎﾞｯｸｽに値を設定する"
        ''' <summary>ﾃｷｽﾄﾎﾞｯｸｽに値を設定する</summary>
        Protected Overrides Sub SetDataToText()
            Try
                Me.SuspendLayout()
                ' GP-IBｸﾞﾙｰﾌﾟﾎﾞｯｸｽ設定
                Call SetGpibData()

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            Finally
                Me.ResumeLayout()
                Me.Refresh()
            End Try

        End Sub
#End Region

#Region "GP-IBｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内の設定"
        ''' <summary>GP-IBｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽに値を設定する</summary>
        Private Sub SetGpibData()
            With m_MainEdit
                If (.W_PLT.GCount < 1) Then ' 登録数 = 0 ?
                    CLblNum.Text = "0" ' 登録数
                    m_GpibNo = 1
                Else
                    CLblNum.Text = (.W_PLT.GCount).ToString() ' 登録数
                End If
            End With

            Try
                With m_MainEdit.W_GPIB(m_GpibNo)
                    For i As Integer = 0 To (m_CtlGpib.Length - 1) Step 1
                        Select Case (i)
                            Case 0 ' 登録番号
                                Dim gCnt As Integer = m_MainEdit.W_PLT.GCount ' 登録数
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

                            Case 1 ' ｱﾄﾞﾚｽ
                                m_CtlGpib(i).Text = (.intGAD).ToString()
                            Case 2 ' ﾃﾞﾘﾐﾀ(0:CRLF, 1:CR, 2:LF, 3:なし)
                                Dim cCombo As cCmb_ = DirectCast(m_CtlGpib(i), cCmb_)
                                Call NoEventIndexChange(cCombo, .intDLM)
                            Case 3 ' 機器名
                                m_CtlGpib(i).Text = .strGNAM
                            Case 4 ' 設定ｺﾏﾝﾄﾞ(1段目)
                                m_CtlGpib(i).Text = (.strCCMD1).ToString()  'V2.0.0.0④
                                'V2.0.0.0④                                If .strCCMD.Length >= 100 Then
                                'V2.0.0.0④                                    m_CtlGpib(i + 0).Text = .strCCMD.Substring(0, 100)
                                'V2.0.0.0④                                Else
                                'V2.0.0.0④                                    m_CtlGpib(i + 0).Text = .strCCMD
                                'V2.0.0.0④                                End If
                            Case 5  ' 設定ｺﾏﾝﾄﾞ(2段目)
                                m_CtlGpib(i).Text = (.strCCMD2).ToString()  'V2.0.0.0④
                                'V2.0.0.0④                                If .strCCMD.Length < 100 Then
                                'V2.0.0.0④                                    m_CtlGpib(i).Text = String.Empty
                                'V2.0.0.0④                                ElseIf 100 < .strCCMD.Length And .strCCMD.Length <= 200 Then
                                'V2.0.0.0④                                    m_CtlGpib(i).Text = .strCCMD.Substring(100)
                                'V2.0.0.0④                                Else
                                'V2.0.0.0④                                    m_CtlGpib(i).Text = .strCCMD.Substring(100, 100)
                                'V2.0.0.0④                                End If
                            Case 6  ' 設定ｺﾏﾝﾄﾞ(3段目)
                                m_CtlGpib(i).Text = (.strCCMD3).ToString()  'V2.0.0.0④
                                'V2.0.0.0④                                    If .strCCMD.Length < 200 Then
                                'V2.0.0.0④                                        m_CtlGpib(i).Text = String.Empty
                                'V2.0.0.0④                                    ElseIf 200 < .strCCMD.Length Then
                                'V2.0.0.0④                                        m_CtlGpib(i).Text = .strCCMD.Substring(200)
                                'V2.0.0.0④                                    End If
                            Case 7 ' ONｺﾏﾝﾄﾞ
                                m_CtlGpib(i).Text = .strCON
                            Case 8 ' ON後のﾎﾟｰｽﾞ時間
                                m_CtlGpib(i).Text = (.lngPOWON).ToString()
                            Case 9 ' OFFｺﾏﾝﾄﾞ
                                m_CtlGpib(i).Text = .strCOFF
                            Case 10 ' OFF後のﾎﾟｰｽﾞ時間
                                m_CtlGpib(i).Text = (.lngPOWOFF).ToString()
                            Case 11 ' ﾄﾘｶﾞｰｺﾏﾝﾄﾞ
                                m_CtlGpib(i).Text = .strCTRG
                                If ("" = .strCTRG) Then ' ﾄﾘｶﾞｰｺﾏﾝﾄﾞ
                                    m_TrgCmdFlg = False ' なし(電源)=False
                                Else
                                    m_TrgCmdFlg = True ' あり(測定器)=True
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

#Region "すべてのﾃｷｽﾄﾎﾞｯｸｽのﾃﾞｰﾀﾁｪｯｸをおこなう"
        ''' <summary>すべてのﾃｷｽﾄﾎﾞｯｸｽのﾃﾞｰﾀﾁｪｯｸをおこなう</summary>
        ''' <returns>0=正常, 1=ｴﾗｰ</returns>
        Friend Overrides Function CheckAllTextData() As Integer
            Dim ret As Integer
            Try
                m_CheckFlg = True ' ﾁｪｯｸ中(tabBase_Layoutにて使用)
                With m_MainEdit
                    .MTab.SelectedIndex = m_TabIdx ' ﾀﾌﾞ表示切替

                    If (.W_PLT.GCount < 1) Then ' ﾃﾞｰﾀなしならNOP
                        ret = 0
                        Exit Try
                    End If

                    For gn As Integer = 1 To .W_PLT.GCount Step 1
                        m_GpibNo = gn
                        ' ﾁｪｯｸする登録番号のﾃﾞｰﾀをｺﾝﾄﾛｰﾙにｾｯﾄする
                        Call SetDataToText()
                        ret = CheckControlData(m_CtlGpib)
                        If (ret <> 0) Then Exit Try

                        ' 相関ﾁｪｯｸ
                        ret = CheckRelation()
                        If (ret <> 0) Then Exit Try
                    Next gn
                End With

                Call CheckDataUpdate() ' ﾃﾞｰﾀ更新確認

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
            Dim ret As Integer = 1
            Dim strTmp As String = String.Empty
            Try
                With m_MainEdit.W_GPIB(m_GpibNo)
                    tag = DirectCast(cTextBox.Parent.Tag, Integer)
                    Select Case (tag)
                        ' ------------------------------------------------------------------------------
                        Case 0 ' GP-IB
                            Select Case (DirectCast(cTextBox.Tag, Integer))
                                Case 0  ' ｱﾄﾞﾚｽ
                                    ret = CheckShortData(cTextBox, .intGAD)
                                Case 1  ' 機器名
                                    ret = CheckStrData(cTextBox, .strGNAM)
                                Case 2  ' 設定ｺﾏﾝﾄﾞ
                                    ret = CheckStrData(cTextBox, .strCCMD1)         'V2.0.0.0④
                                    'V2.0.0.0④                                    ret = CheckStrData(cTextBox, strTmp)
                                    'V2.0.0.0④                                    .strCCMD = strTmp + m_CtlGpib(5).Text + m_CtlGpib(6).Text
                                Case 3  ' 設定ｺﾏﾝﾄﾞ(2段目)
                                    ret = CheckStrData(cTextBox, .strCCMD2)         'V2.0.0.0④
                                    'V2.0.0.0④                                    ret = CheckStrData(cTextBox, strTmp)
                                    'V2.0.0.0④                                    .strCCMD = m_CtlGpib(4).Text + strTmp + m_CtlGpib(6).Text
                                Case 4  ' 設定ｺﾏﾝﾄﾞ(3段目)
                                    ret = CheckStrData(cTextBox, .strCCMD3)         'V2.0.0.0④
                                    'V2.0.0.0④                                    ret = CheckStrData(cTextBox, strTmp)
                                    'V2.0.0.0④                                    .strCCMD = m_CtlGpib(4).Text + m_CtlGpib(5).Text + strTmp
                                Case 5  ' ONｺﾏﾝﾄﾞ
                                    ret = CheckStrData(cTextBox, .strCON)
                                Case 6  ' ON後のﾎﾟｰｽﾞ時間
                                    ret = CheckIntData(cTextBox, .lngPOWON)
                                Case 7  ' OFFｺﾏﾝﾄﾞ
                                    ret = CheckStrData(cTextBox, .strCOFF)
                                Case 8  ' OFF後のﾎﾟｰｽﾞ時間
                                    ret = CheckIntData(cTextBox, .lngPOWOFF)
                                Case 9  ' ﾄﾘｶﾞｰｺﾏﾝﾄﾞ
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

#Region "相関ﾁｪｯｸ"
        Protected Overrides Function CheckRelation() As Integer
            Dim erridx As Integer
            Dim strMsg As String

            CheckRelation = 0 ' Return値 = 正常
            Try
                With m_MainEdit.W_GPIB(m_GpibNo)
                    '---------------------------------------------------------------
                    ' コマンド入力なし
                    '---------------------------------------------------------------
                    If ("" = .strCON) AndAlso ("" = .strCOFF) AndAlso ("" = .strCTRG) Then
                        erridx = GPIB_CON
                        strMsg = "相関チェックエラー" & vbCrLf & _
                                "コマンドを入力して下さい。"
                        GoTo STP_ERR
                    End If

                    '---------------------------------------------------------------
                    ' ONまたはOFFとﾄﾘｶﾞｰｺﾏﾝﾄﾞが入力されている
                    '---------------------------------------------------------------
                    'If ("" <> .strCTRG) AndAlso _
                    '        (("" <> .strCON) OrElse ("" <> .strCOFF)) Then
                    '    erridx = GPIB_CTRG
                    '    strMsg = "相関チェックエラー" & vbCrLf & _
                    '            "コマンドの組み合わせが正しくありません。"
                    '    GoTo STP_ERR
                    'End If

                    '---------------------------------------------------------------
                    ' ON入力ありかつOFF未入力
                    '---------------------------------------------------------------
                    If ("" <> .strCON) AndAlso ("" = .strCOFF) Then
                        erridx = GPIB_COFF
                        strMsg = "相関チェックエラー" & vbCrLf & _
                                "コマンドの組み合わせが正しくありません。"
                        GoTo STP_ERR
                    End If

                    '---------------------------------------------------------------
                    ' ON未入力かつOFF入力あり
                    '---------------------------------------------------------------
                    If ("" = .strCON) AndAlso ("" <> .strCOFF) Then
                        erridx = GPIB_CON
                        strMsg = "相関チェックエラー" & vbCrLf & _
                                "コマンドの組み合わせが正しくありません。"
                        GoTo STP_ERR
                    End If
                End With

                Exit Function
STP_ERR:
                If (TypeOf m_CtlGpib(erridx) Is cTxt_) Then
                    ' ﾃｷｽﾄﾎﾞｯｸｽがｴﾗｰの場合
                    Call MsgBox_CheckErr(DirectCast(m_CtlGpib(erridx), cTxt_), strMsg)
                ElseIf (TypeOf m_CtlGpib(erridx) Is cCmb_) Then
                    ' ｺﾝﾎﾞﾎﾞｯｸｽがｴﾗｰの場合
                    Call MsgBox_CheckErr(DirectCast(m_CtlGpib(erridx), cCmb_), strMsg)
                Else
                    ' DO NOTHING
                End If
                CheckRelation = 1 ' Return値 = ｴﾗｰ

            Catch ex As Exception
                Call MsgBox_Exception(ex)
                CheckRelation = 1 ' Return値 = ｴﾗｰ
            End Try

        End Function
#End Region

#Region "ﾃﾞｰﾀの更新を確認する"
        ''' <summary>ﾃﾞｰﾀの更新があった場合、GPIBﾃﾞｰﾀ更新FlagをONにする</summary>
        Private Sub CheckDataUpdate()
            Dim flg As Integer = 0
            Try
                With m_MainEdit
                    If (.W_PLT.GCount <> stPLT.GCount) Then flg = 1 : Exit Try ' 登録数
                    For i As Integer = 1 To stPLT.GCount Step 1
                        With .W_GPIB(i)
                            If (.intGAD <> stGPIB(i).intGAD) Then flg = 1 : Exit Try ' ｱﾄﾞﾚｽ
                            If (.intDLM <> stGPIB(i).intDLM) Then flg = 1 : Exit Try ' ﾃﾞﾘﾐﾀ
                            If (.strGNAM <> stGPIB(i).strGNAM) Then flg = 1 : Exit Try ' 機器名
                            'V2.0.0.0④                            If (.strCCMD <> stGPIB(i).strCCMD) Then flg = 1 : Exit Try ' 設定ｺﾏﾝﾄﾞ
                            If (.strCCMD1 <> stGPIB(i).strCCMD1) Then flg = 1 : Exit Try ' 設定ｺﾏﾝﾄﾞ'V2.0.0.0④
                            If (.strCCMD2 <> stGPIB(i).strCCMD2) Then flg = 1 : Exit Try ' 設定ｺﾏﾝﾄﾞ'V2.0.0.0④
                            If (.strCCMD3 <> stGPIB(i).strCCMD3) Then flg = 1 : Exit Try ' 設定ｺﾏﾝﾄﾞ'V2.0.0.0④
                            If (.strCON <> stGPIB(i).strCON) Then flg = 1 : Exit Try ' ONｺﾏﾝﾄﾞ
                            If (.strCOFF <> stGPIB(i).strCOFF) Then flg = 1 : Exit Try ' OFFｺﾏﾝﾄﾞ
                            If (.lngPOWON <> stGPIB(i).lngPOWON) Then flg = 1 : Exit Try ' ON後のﾎﾟｰｽﾞ時間
                            If (.lngPOWOFF <> stGPIB(i).lngPOWOFF) Then flg = 1 : Exit Try ' OFF後のﾎﾟｰｽﾞ時間
                            If (.strCTRG <> stGPIB(i).strCTRG) Then flg = 1 : Exit Try ' ﾄﾘｶﾞｰｺﾏﾝﾄﾞ
                        End With
                    Next i
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            Finally
                FlgUpdGPIB = Convert.ToInt16(flg) ' GPIBﾃﾞｰﾀ更新Flag ON=1
            End Try

        End Sub
#End Region

#Region "追加･削除ﾎﾞﾀﾝ関連処理"
        ''' <summary>GP-IBﾃﾞｰﾀを追加または削除し、そのﾃﾞｰﾀを初期化する</summary>
        ''' <param name="addDel">1=Add, (-1)=Del</param>
        Private Sub SortGpibData(ByVal addDel As Integer)
            Dim iStart As Integer
            Dim iEnd As Integer
            Dim dir As Integer = (-1) * addDel ' Add=(-1), Del=1にする
            Try
                With m_MainEdit
                    If (1 = addDel) Then ' 追加の場合
                        .W_PLT.GCount = Convert.ToInt16(.W_PLT.GCount + 1) ' 登録数を追加する
                        iStart = .W_PLT.GCount ' 登録数から
                        iEnd = (m_GpibNo + 1) ' 追加するﾃﾞｰﾀの登録番号+1まで、前のﾃﾞｰﾀを後ろにずらす
                    Else ' 削除の場合
                        iStart = m_GpibNo ' 削除するﾃﾞｰﾀの登録番号から
                        iEnd = (.W_PLT.GCount - 1) ' 登録されている登録数-1まで、後ろのﾃﾞｰﾀを前にずらす
                    End If

                    For i As Integer = iStart To iEnd Step dir
                        .W_GPIB(i).intGAD = .W_GPIB(i + dir).intGAD         ' ｱﾄﾞﾚｽ
                        .W_GPIB(i).strGNAM = .W_GPIB(i + dir).strGNAM       ' 機器名
                        .W_GPIB(i).intDLM = .W_GPIB(i + dir).intDLM         ' ﾃﾞﾘﾐﾀ(0:CRLF, 1:CR, 2:LF, 3:なし)
                        'V2.0.0.0④                        .W_GPIB(i).strCCMD = .W_GPIB(i + dir).strCCMD       ' 設定ｺﾏﾝﾄﾞ
                        .W_GPIB(i).strCCMD1 = .W_GPIB(i + dir).strCCMD1       ' 設定ｺﾏﾝﾄﾞ'V2.0.0.0④
                        .W_GPIB(i).strCCMD2 = .W_GPIB(i + dir).strCCMD2       ' 設定ｺﾏﾝﾄﾞ'V2.0.0.0④
                        .W_GPIB(i).strCCMD3 = .W_GPIB(i + dir).strCCMD3       ' 設定ｺﾏﾝﾄﾞ'V2.0.0.0④
                        .W_GPIB(i).strCON = .W_GPIB(i + dir).strCON         ' ONｺﾏﾝﾄﾞ
                        .W_GPIB(i).lngPOWON = .W_GPIB(i + dir).lngPOWON     ' ON後のﾎﾟｰｽﾞ時間(ms)
                        .W_GPIB(i).strCOFF = .W_GPIB(i + dir).strCOFF       ' OFFｺﾏﾝﾄﾞ
                        .W_GPIB(i).lngPOWOFF = .W_GPIB(i + dir).lngPOWOFF   ' OFF後のﾎﾟｰｽﾞ時間(ms)
                        .W_GPIB(i).strCTRG = .W_GPIB(i + dir).strCTRG       ' ﾄﾘｶﾞｰｺﾏﾝﾄﾞ
                    Next i

                    ' 追加または削除した登録番号以降の登録番号が使用されている場合にその値を変更する
                    Call ResetResCutData(addDel)

                    ' つめて不要となったﾃﾞｰﾀを初期化する
                    If (1 = addDel) Then ' 追加の場合
                        Call InitGpibData(m_GpibNo) ' 追加したﾃﾞｰﾀを初期化
                    Else ' 削除の場合
                        Call InitGpibData(.W_PLT.GCount) ' 最終登録番号のﾃﾞｰﾀを初期化
                        .W_PLT.GCount = Convert.ToInt16(.W_PLT.GCount - 1) ' 登録数を-1する

                        ' 最終ﾃﾞｰﾀの削除なら現在の登録番号を最終登録番号とする
                        If (.W_PLT.GCount < m_GpibNo) Then m_GpibNo = .W_PLT.GCount
                    End If
                End With

                ' GP-IBﾃﾞｰﾀを画面項目に設定
                Call SetDataToText()
                FIRST_CONTROL.Select() ' ﾌｵｰｶｽ設定

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

        ''' <summary>追加･削除したﾃﾞｰﾀの登録番号以降の番号が使用されている場合にその値を変更する</summary>
        ''' <param name="addDel">1=Add, (-1)=Del</param>
        ''' <param name="delTrgCmd">ﾄﾘｶﾞｰｺﾏﾝﾄﾞの削除=True</param>
        Private Sub ResetResCutData(ByVal addDel As Integer, Optional ByVal delTrgCmd As Boolean = False)
            Try
                With m_MainEdit
                    If (1 = addDel) AndAlso (False = delTrgCmd) Then ' 追加の場合
                        For rn As Integer = 1 To .W_PLT.RCount Step 1
                            With .W_REG(rn)
                                ' 追加ﾃﾞｰﾀの登録番号より後の番号を使用している場合、番号を(+1)する
                                If (m_GpibNo <= .intMType) Then ' 抵抗
                                    .intMType = Convert.ToInt16(.intMType + addDel)
                                End If

                                ' ------------------------
                                For i As Integer = 1 To EXTEQU Step 1
                                    If (m_GpibNo <= .intOnExtEqu(i)) Then   ' 抵抗ﾀﾌﾞ > ON機器
                                        ' 追加ﾃﾞｰﾀの登録番号より後の番号を使用している場合、番号を(+1)する
                                        .intOnExtEqu(i) = Convert.ToInt16(.intOnExtEqu(i) + addDel)
                                    End If
                                    If (m_GpibNo <= .intOffExtEqu(i)) Then  ' 抵抗ﾀﾌﾞ > OFF機器
                                        ' 追加ﾃﾞｰﾀの登録番号より後の番号を使用している場合、番号を(+1)する
                                        .intOffExtEqu(i) = Convert.ToInt16(.intOffExtEqu(i) + addDel)
                                    End If
                                Next i
                                ' ------------------------

                                For cn As Integer = 1 To (.STCUT.Length - 1) Step 1 ' ｶｯﾄ
                                    With .STCUT(cn)
                                        If (m_GpibNo <= .intMType) Then
                                            .intMType = Convert.ToInt16(.intMType + addDel)
                                        End If

                                        For ix As Integer = 1 To (.intIXMType.Length - 1) Step 1 ' ｲﾝﾃﾞｯｸｽｶｯﾄ
                                            If (m_GpibNo <= .intIXMType(ix)) Then
                                                .intIXMType(ix) = Convert.ToInt16(.intIXMType(ix) + addDel)
                                            End If
                                        Next ix
                                    End With
                                Next cn

                            End With
                        Next rn

                    Else ' 削除の場合
                        For rn As Integer = 1 To .W_PLT.RCount Step 1
                            With .W_REG(rn)
                                If (m_GpibNo < .intMType) AndAlso (False = delTrgCmd) Then ' 抵抗
                                    ' 削除ﾃﾞｰﾀの登録番号より後の番号を使用している場合、番号を(-1)する
                                    .intMType = Convert.ToInt16(.intMType + addDel)
                                ElseIf (m_GpibNo = .intMType) Then ' 抵抗
                                    ' 削除ﾃﾞｰﾀ、またはﾄﾘｶﾞｰｺﾏﾝﾄﾞが削除されたﾃﾞｰﾀの
                                    ' 登録番号を使用中の場合、0(内部抵抗測定)とする
                                    .intMType = 0
                                Else
                                    ' DO NOTHING
                                End If

                                ' ------------------------
                                For i As Integer = 1 To EXTEQU Step 1
                                    ' 抵抗ﾀﾌﾞ > ON機器
                                    If (m_GpibNo < .intOnExtEqu(i)) AndAlso (False = delTrgCmd) Then
                                        ' 削除ﾃﾞｰﾀの登録番号より後の番号を使用している場合、番号を(-1)する
                                        .intOnExtEqu(i) = Convert.ToInt16(.intOnExtEqu(i) + addDel)
                                    ElseIf (m_GpibNo = .intOnExtEqu(i)) Then
                                        ' 削除ﾃﾞｰﾀ、またはﾄﾘｶﾞｰｺﾏﾝﾄﾞが削除されたﾃﾞｰﾀの
                                        ' 登録番号を使用中の場合、0(なし)とする
                                        .intOnExtEqu(i) = 0
                                    Else
                                        ' DO NOTHING
                                    End If

                                    ' 抵抗ﾀﾌﾞ > OFF機器
                                    If (m_GpibNo < .intOffExtEqu(i)) AndAlso (False = delTrgCmd) Then
                                        ' 削除ﾃﾞｰﾀの登録番号より後の番号を使用している場合、番号を(-1)する
                                        .intOffExtEqu(i) = Convert.ToInt16(.intOffExtEqu(i) + addDel)
                                    ElseIf (m_GpibNo = .intOffExtEqu(i)) Then
                                        ' 削除ﾃﾞｰﾀ、またはﾄﾘｶﾞｰｺﾏﾝﾄﾞが削除されたﾃﾞｰﾀの
                                        ' 登録番号を使用中の場合、0(なし)とする
                                        .intOffExtEqu(i) = 0
                                    Else
                                        ' DO NOTHING
                                    End If
                                Next i
                                ' ------------------------

                                For cn As Integer = 1 To (.STCUT.Length - 1) Step 1 ' ｶｯﾄ
                                    With .STCUT(cn)
                                        ' 削除ﾃﾞｰﾀの登録番号より後の番号を使用している場合、番号を(-1)する
                                        If (m_GpibNo < .intMType) AndAlso (False = delTrgCmd) Then
                                            .intMType = Convert.ToInt16(.intMType + addDel)
                                        ElseIf (m_GpibNo = .intMType) Then
                                            ' 削除ﾃﾞｰﾀ、またはﾄﾘｶﾞｰｺﾏﾝﾄﾞが削除されたﾃﾞｰﾀの
                                            ' 登録番号を使用中の場合、0(内部抵抗測定)とする
                                            .intMType = 0
                                        Else
                                            ' DO NOTHING
                                        End If

                                        For ix As Integer = 1 To (.intIXMType.Length - 1) Step 1 ' ｲﾝﾃﾞｯｸｽｶｯﾄ
                                            ' 削除ﾃﾞｰﾀの登録番号より後の番号を使用している場合、番号を(-1)する
                                            If (m_GpibNo < .intIXMType(ix)) AndAlso (False = delTrgCmd) Then
                                                .intIXMType(ix) = Convert.ToInt16(.intIXMType(ix) + addDel)
                                            ElseIf (m_GpibNo = .intIXMType(ix)) Then
                                                ' 削除ﾃﾞｰﾀ、またはﾄﾘｶﾞｰｺﾏﾝﾄﾞが削除されたﾃﾞｰﾀの
                                                ' 登録番号を使用中の場合、0(内部抵抗測定)とする
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
                        Case 0 ' GP-IBｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            Select Case (tag)
                                Case 0 ' 登録番号
                                    m_GpibNo = (idx + 1)
                                    ' 対応するﾃﾞｰﾀをﾃｷｽﾄﾎﾞｯｸｽにｾｯﾄする
                                    Call SetDataToText()
                                Case 1 ' ﾃﾞﾘﾐﾀ(0:CRLF, 1:CR, 2:LF, 3:なし)
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

        ''' <summary>追加ﾎﾞﾀﾝｸﾘｯｸ時の処理</summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Private Sub CBtn_Add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBtn_Add.Click
            Dim strMsg As String
            Dim refOpt As Short ' ｵﾌﾟｼｮﾝﾎﾞﾀﾝ(0=前に追加 ,1=後に追加)
            Dim ret As Integer
            Try
                With m_MainEdit
                    ' 登録数ﾁｪｯｸ
                    If (MAXGNO <= .W_PLT.GCount) Then ' 登録数OK ?
                        strMsg = "これ以上ＧＰ－ＩＢデータは登録できません。"
                        Call MsgBox(strMsg, DirectCast( _
                                    MsgBoxStyle.OkOnly + _
                                    MsgBoxStyle.Information, MsgBoxStyle), _
                                    My.Application.Info.Title)
                        Exit Sub
                    End If

                    If (0 < .W_PLT.GCount) Then ' 登録数が1以上の場合
                        ' 確認ﾒｯｾｰｼﾞを表示("ＧＰ－ＩＢデータを追加します")
                        ret = MsgBox_AddClick("ＧＰ－ＩＢデータ", refOpt) ' ﾒｯｾｰｼﾞ表示
                        If (ret <> cFRS_ERR_ADV) Then Exit Sub ' CancelならReturn

                        If (refOpt = 1) Then ' 表示ﾃﾞｰﾀの後に追加 ?
                            m_GpibNo = (m_GpibNo + 1) ' 追加するﾃﾞｰﾀの登録番号 = 現在の登録番号番号 + 1
                        Else ' 表示ﾃﾞｰﾀの前に追加
                            m_GpibNo = m_GpibNo ' 追加するﾃﾞｰﾀの登録番号 = 現在の登録番号
                        End If
                        ' ﾃﾞｰﾀを1個後にずらして追加する
                        Call SortGpibData(1)

                    Else ' 登録数が0の場合
                        With m_CtlGpib(GPIB_GNAM) ' 機器名入力確認
                            If ("" = .Text) OrElse (.Text Is Nothing) Then
                                .Select()
                                .BackColor = Color.Yellow
                                strMsg = DirectCast(m_CtlGpib(GPIB_GNAM), cTxt_).GetStrMsg & "の入力をおこなってください。"
                                Call MsgBox(strMsg, DirectCast( _
                                            MsgBoxStyle.OkOnly + _
                                            MsgBoxStyle.Information, MsgBoxStyle), _
                                            My.Application.Info.Title)
                                Exit Sub
                            End If
                        End With

                        strMsg = "ＧＰ－ＩＢデータを登録しますか？"
                        If (MsgBoxResult.Ok = MsgBox(strMsg, DirectCast( _
                                                    MsgBoxStyle.OkCancel + _
                                                    MsgBoxStyle.Information, MsgBoxStyle), _
                                                    My.Application.Info.Title)) Then
                            .W_PLT.GCount = 1 ' 登録数を設定
                            m_GpibNo = 1

                            ' GP-IBﾃﾞｰﾀを画面項目に設定
                            Call SetDataToText()
                            FIRST_CONTROL.Select() ' ﾌｵｰｶｽ設定
                        End If

                    End If
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

        ''' <summary>削除ﾎﾞﾀﾝｸﾘｯｸ時の処理</summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Private Sub CBtn_Del_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBtn_Del.Click
            Dim strMsg As String ' ﾒｯｾｰｼﾞ編集域
            Dim ret As Integer
            Try
                ' 確認ﾒｯｾｰｼﾞを表示
                If (0 = m_MainEdit.W_PLT.GCount) Then Exit Sub ' 登録数0ならNOP
                strMsg = "現在のＧＰ－ＩＢデータを削除します。よろしいですか？"
                ret = MsgBox(strMsg, DirectCast( _
                            MsgBoxStyle.OkCancel + _
                            MsgBoxStyle.Exclamation + _
                            MsgBoxStyle.DefaultButton2, MsgBoxStyle), _
                            My.Application.Info.Title)
                If (ret = MsgBoxResult.Cancel) Then Exit Sub ' Cancel(RESETｷｰ) ?

                ' 後ろのﾃﾞｰﾀを1個前につめる
                Call SortGpibData(-1)

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

        ''' <summary>ﾄﾘｶﾞｰｺﾏﾝﾄﾞﾃｷｽﾄﾎﾞｯｸｽがﾌｫｰｶｽを失った(変更された可能性がある)時におこなう処理</summary>
        ''' <param name="sender">ﾄﾘｶﾞｰｺﾏﾝﾄﾞﾃｷｽﾄﾎﾞｯｸｽ</param>
        ''' <param name="e"></param>
        Private Sub CTxt_7_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CTxt_7.Leave
            ' ﾀﾌﾞまたは登録番号を切替えた場合、SetGpibData()の処理により
            ' m_TrgCmdFlg は ﾄﾘｶﾞｰｺﾏﾝﾄﾞあり(測定器)=True,なし(電源)=False となる
            ' 追加または削除された場合、m_TrgCmdFlg は ﾄﾘｶﾞｰｺﾏﾝﾄﾞなし(電源)=False となる
            Dim txt As String
            Try
                txt = DirectCast(sender, cTxt_).Text
                If (0 < m_MainEdit.W_PLT.GCount) Then ' 登録がある場合
                    If (False = m_TrgCmdFlg) AndAlso ("" <> txt) Then ' ﾄﾘｶﾞｰｺﾏﾝﾄﾞなしの状態から入力があった場合
                        ' 使用中の登録番号は変わらないため、ﾌﾗｸﾞのみ切替える
                        m_TrgCmdFlg = True
                        Exit Sub
                    ElseIf (True = m_TrgCmdFlg) AndAlso ("" = txt) Then ' ﾄﾘｶﾞｰｺﾏﾝﾄﾞありの状態から削除された場合
                        ' ﾄﾘｶﾞｰｺﾏﾝﾄﾞが削除された登録番号を使用している抵抗ﾃﾞｰﾀの値を0(内部測定器)にする
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
