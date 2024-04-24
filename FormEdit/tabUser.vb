Option Explicit On
Option Strict Off

Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports LaserFront.Trimmer.DefWin32Fnc

Namespace FormEdit
    Friend Class tabUser
        Inherits tabBase

#Region "宣言"
        Private Const SYS_ZON As Integer = 11       ' 相関ﾁｪｯｸで使用(m_CtlSystemでのｲﾝﾃﾞｯｸｽ)
        Private Const SYS_ZOFF As Integer = 12      ' 相関ﾁｪｯｸで使用(m_CtlSystemでのｲﾝﾃﾞｯｸｽ)

        Private m_CtlSystem() As Control            ' USERのｺﾝﾄﾛｰﾙ配列
        Private m_CtlResistor(,) As Control         ' 抵抗設定の抵抗番号ﾎﾞｯｸｽのｺﾝﾄﾛｰﾙ配列
        Private m_DispResistor(,) As Control        ' 抵抗数に応じて表示/非表示を切り替えるｺﾝﾄﾛｰﾙ
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
                TAB_NAME = GetPrivateProfileString_S("USER_LABEL", "TAB_NAM", m_sPath, "????")

                ' ----------------------------------------------------------
                ' EDIT_DEF_User.iniからｸﾞﾙｰﾌﾟﾎﾞｯｸｽに表示名を設定
                ' ----------------------------------------------------------
                GrpArray = New cGrp_() { _
                    CGrp_0, CGrp_1, CGrp_2, CGrp_3, CGrp_4 _
                }
                For i As Integer = 0 To (GrpArray.Length - 1) Step 1
                    With GrpArray(i)
                        .TabIndex = ((254 - GrpArray.Length) + i) ' ｶｰｿﾙｷｰによるﾌｫｰｶｽ移動で必要
                        .Tag = i
                        .Text = GetPrivateProfileString_S( _
                            "USER_LABEL", (i.ToString("000") & "_GRP"), m_sPath, "????")
                    End With
                Next i

                ' ----------------------------------------------------------
                ' EDIT_DEF_User.iniからﾗﾍﾞﾙに表示名を設定
                ' ----------------------------------------------------------
                'V2.1.0.0① CLbl_26追加
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
                ' ｼｽﾃﾑｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のｺﾝﾄﾛｰﾙを設定(CtlArrayの順番と合わせる)
                ' ----------------------------------------------------------
                'V2.1.0.0③ 温度センサー情報一元管理選択番号追加 CTxt_37,CTxt_38,０℃(CTxt_7)のTabオーダーを上段Noの次に変更
                m_CtlSystem = New Control() { _
                    CCmb_0, _
                    CTxt_0, CTxt_1, CTxt_2, CTxt_3, CCmb_1, CCmb_2, CTxt_4, CTxt_5, CTxt_6, CCmb_10, _
                    CCmb_3, CTxt_28, CTxt_30, CTxt_31, CTxt_37, CTxt_7, CTxt_8, CTxt_29, CTxt_38, CTxt_9, CTxt_10, _
                    CTxt_11, CTxt_12 _
                }
                'V2.0.0.0⑪                m_CtlSystem = New Control() { _
                'V2.0.0.0⑪                    CCmb_0, CTxt_0, CTxt_1, CTxt_2, CTxt_3, CCmb_1, CCmb_2, CTxt_4, CTxt_5, CTxt_6, CCmb_3, CCmb_4, _
                'V2.0.0.0⑪                    CTxt_7, CTxt_8, CTxt_9, CTxt_10, CTxt_11, CTxt_12 _
                'V2.0.0.0⑪                }
                Call SetControlData(m_CtlSystem)

                ' ----------------------------------------------------------
                ' 抵抗設定の抵抗番号のｺﾝﾄﾛｰﾙを設定(CtlArrayの順番と合わせる)
                ' ----------------------------------------------------------
                'V2.1.0.0① CTxt_32～CTxt_36 判定用目標値算出係数追加
                m_CtlResistor = New Control(,) { _
                    {CCmb_5, CTxt_13, CTxt_14, CTxt_32, CTxt_15}, _
                    {CCmb_6, CTxt_16, CTxt_17, CTxt_33, CTxt_18}, _
                    {CCmb_7, CTxt_19, CTxt_20, CTxt_34, CTxt_21}, _
                    {CCmb_8, CTxt_22, CTxt_23, CTxt_35, CTxt_24}, _
                    {CCmb_9, CTxt_25, CTxt_26, CTxt_36, CTxt_27} _
                }
                Call SetControlData(m_CtlResistor)

                'V2.1.0.0① CTxt_32～CTxt_36 判定用目標値算出係数追加
                m_DispResistor = New Control(,) { _
                    {Label1, CCmb_5, CTxt_13, CTxt_14, CTxt_32, CTxt_15}, _
                    {Label2, CCmb_6, CTxt_16, CTxt_17, CTxt_33, CTxt_18}, _
                    {Label3, CCmb_7, CTxt_19, CTxt_20, CTxt_34, CTxt_21}, _
                    {Label4, CCmb_8, CTxt_22, CTxt_23, CTxt_35, CTxt_24}, _
                    {Label5, CCmb_9, CTxt_25, CTxt_26, CTxt_36, CTxt_27} _
                }

                ' ----------------------------------------------------------
                ' すべてのﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽ･ﾎﾞﾀﾝに対し、ﾌｫｰｶｽ移動設定をおこなう
                ' 使用しないｺﾝﾄﾛｰﾙは Enabled=False または Visible=False にする
                ' ----------------------------------------------------------
                'V2.1.0.0① CTxt_32～CTxt_36 判定用目標値算出係数追加
                'V2.1.0.0③ 温度センサー情報一元管理選択番号追加 CTxt_37,CTxt_38,０℃(CTxt_7)のTabオーダーを上段Noの次に変更
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
                                Case 0 ' 製品種別
                                    .Items.Add("指定なし")
                                    .Items.Add("温度センサー")
                                    .Items.Add("抵抗トリミング")
                                    .Items.Add("チップ抵抗トリミング")    'V1.0.4.3④
                                    .Items.Add("チップ温度センサー") 'V2.0.0.0①
                                    .Items.Add("マーク印字") 'V2.2.1.7①
                                Case 1 ' トリミング速度
                                    .Items.Add("1:高速")
                                    .Items.Add("2:高精度")
                                    .Items.Add("3:設定値")
                                Case 2 ' ロット終了条件
                                    .Items.Add("0:終了条件判定無し")
                                    .Items.Add("1:枚数")
                                    .Items.Add("2:ローダ信号")
                                    .Items.Add("3:枚数＆信号")
                                    'V2.0.0.0⑭↓
                                Case 3  ' クランプと吸着の有り無し
                                    .Items.Add("クランプ吸着有り")
                                    .Items.Add("クランプのみ")
                                    .Items.Add("吸着のみ")
                                    'V2.0.0.0⑭↑
                                Case 4 ' 抵抗単位
                                    .Items.Add("1:Ω")
                                    .Items.Add("2:ＫΩ")
                                    'V2.0.0.0⑪                                Case 4 ' 参照温度
                                    'V2.0.0.0⑪                         .Items.Add("1:０℃")
                                    'V2.0.0.0⑪              .Items.Add("2:２５℃")
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                        Case 4 ' 補正値ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            Select Case (tag)
                                Case 0 ' 抵抗単位
                                    .Items.Add("1:Ω")
                                    .Items.Add("2:ＫΩ")
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
                        strMsg = GetPrivateProfileString_S("USER_VALUE", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0  ' ユーザ名
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "15")
                                strFlg = True
                            Case 1  ' レーザ　ロットＮｏ．
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "15")
                                strFlg = True
                            Case 2  ' パターンＮｏ．
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "15")
                                strFlg = True
                            Case 3  ' プログラムＮｏ．
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "15")
                                strFlg = True
                            Case 4 ' 処理枚数
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "99999")
                            Case 5  ' 補正頻度
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "999")
                            Case 6  ' 印刷素子数
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "9999")
                                'V2.0.0.0⑪↓
                            Case 7  ' 設定温度
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "100")
                            Case 8  ' 代表α値(ppm/℃)
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "-9999.0000000")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "9999.0000000")
                            Case 9  ' 代表β値(ppm/℃)
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "-9999.0000000")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "9999.0000000")
                            Case 10  ' 温度センサー情報一元管理選択番号
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "99")
                                Dim TempNo As Integer = Integer.Parse(strMax)
                                Dim TableNo As Integer = UserSub.TemperatureTableMaxNumberGet()
                                If TableNo < TempNo Then
                                    strMax = TableNo.ToString("0")
                                End If
                                'V2.1.0.0③↑
                            Case 11  ' ０℃　'V2.1.0.0③ Case 8から11へ移動
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "0.0000001")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "100000000.0000000")
                            Case 12  ' α値(ppm/℃)
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "-9999.0000000")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "9999.0000000")
                            Case 13  ' β値(ppm/℃)
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "-9999.0000000")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "9999.0000000")
                                'V2.0.0.0⑪↑
                                'V2.1.0.0③↓
                            Case 14  ' 温度センサー情報一元管理選択番号
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "99")
                                Dim TempNo As Integer = Integer.Parse(strMax)
                                Dim TableNo As Integer = UserSub.TemperatureTableMaxNumberGet()
                                If TableNo < TempNo Then
                                    strMax = TableNo.ToString("0")
                                End If
                                'V2.1.0.0③↑
                                'V2.0.0.0⑪                            Case 7  ' 標準抵抗値
                                'V2.0.0.0⑪                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "0.01")
                                'V2.0.0.0⑪                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "100000000")
                                'V2.0.0.0⑪                            Case 8  ' 抵抗温度係数
                                'V2.0.0.0⑪                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "0")
                                'V2.0.0.0⑪                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "1")
                            Case Else
                                Throw New Exception("Case " & tag & ": Nothing")
                        End Select
                        ' ------------------------------------------------------------------------------
                    Case 4 ' 補正値ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                        strMsg = GetPrivateProfileString_S("USER_VALUE", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0 To 3 ' ユーザ名　'V2.1.0.0①  Case 0 To 2 から  Case 0 To 3 へ変更
                                For i As Integer = 0 To (m_CtlResistor.GetLength(0) - 1) Step 1
                                    For j As Integer = 0 To (m_CtlResistor.GetLength(1) - 1) Step 1
                                        If (m_CtlResistor(i, j) Is cTextBox) Then
                                            tag = j + 18        'V2.1.0.0③　011_MSG = Ｎｏ．,014_MSG = Ｎｏ．が追加になったので２つずれる。16→18へ変更
                                            no = tag.ToString("000")
                                            strMsg = GetPrivateProfileString_S("USER_VALUE", (no & "_MSG"), m_sPath, "??????")
                                            Select Case (j)
                                                Case 1 ' 補正値 'V2.0.0.0⑫ 補正値の項目をppm入力に変更
                                                    strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "99999")
                                                    strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "-99999")
                                                Case 2 ' 目標値算出係数                                                        
                                                    strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "0")
                                                    strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "9999.9")
                                                    'V2.1.0.0①↓
                                                Case 3 ' 判定用目標値算出係数                                                        
                                                    strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "0")
                                                    strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "9999.9")
                                                    'V2.1.0.0①↑
                                                Case 4 ' 測定速度を変更するカットＮｏ．                                        'V2.1.0.0① Case 3からCase 4へ変更
                                                    strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "1")
                                                    strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "99")
                                                Case Else
                                                    Throw New Exception("Case " & tag & ": Nothing")
                                            End Select
                                            Exit For
                                        End If
                                    Next j
                                Next i
                                'V2.1.0.0③ 温度センサー情報一元管理選択番号追加Case番号２繰上げ13→15
                            Case 15 ' ファイナルリミット High[%]
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "-99.99")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "99.99")
                            Case 16 ' ファイナルリミット Lo[%]
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "-99.99")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "99.99")
                            Case 17  ' 相対値リミット High[%]
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "-99.99")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "99.99")
                            Case 18  ' 相対値リミット Lo[%]
                                strMin = GetPrivateProfileString_S("USER_VALUE", (no & "_MIN"), m_sPath, "-99.99")
                                strMax = GetPrivateProfileString_S("USER_VALUE", (no & "_MAX"), m_sPath, "99.99")
                            Case Else
                                Throw New Exception("Case " & tag & ": Nothing")
                        End Select

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
                            Case 0 ' 製品種別
                                Dim cCombo As cCmb_ = DirectCast(m_CtlSystem(i), cCmb_)
                                Call NoEventIndexChange(cCombo, .iTrimType) ' 指定製品種別を設定
                            Case 1  ' オペレータ名
                                m_CtlSystem(i).Text = .sOperator
                            Case 2  ' レーザロットNo.
                                m_CtlSystem(i).Text = .sLotNumber
                            Case 3  ' パターンＮｏ．
                                m_CtlSystem(i).Text = .sPatternNo
                            Case 4  ' プログラムＮｏ．
                                m_CtlSystem(i).Text = .sProgramNo
                            Case 5  ' トリミング速度
                                Dim cCombo As cCmb_ = DirectCast(m_CtlSystem(i), cCmb_)
                                Call NoEventIndexChange(cCombo, (.iTrimSpeed - 1))
                            Case 6  ' ロット終了条件
                                Dim cCombo As cCmb_ = DirectCast(m_CtlSystem(i), cCmb_)
                                Call NoEventIndexChange(cCombo, .iLotChange)
                            Case 7  ' ロット処理枚数
                                m_CtlSystem(i).Text = (.lLotEndSL).ToString()
                            Case 8  ' カット位置補正頻度
                                m_CtlSystem(i).Text = (.lCutHosei).ToString()
                            Case 9  ' ロット終了時印刷素子数
                                m_CtlSystem(i).Text = (.lPrintRes).ToString()
                            Case 10 ' クランプと吸着の有り無し 'V2.0.0.0⑭
                                Dim cCombo As cCmb_ = DirectCast(m_CtlSystem(i), cCmb_)
                                Call NoEventIndexChange(cCombo, (.intClampVacume - 1))
                            Case 11 ' 温度センサー 抵抗レンジ 1:Ω, 2:KΩ
                                Dim cCombo As cCmb_ = DirectCast(m_CtlSystem(i), cCmb_)
                                Call NoEventIndexChange(cCombo, (.iTempResUnit - 1))
                                'V2.0.0.0⑪↓
                            Case 12  ' 設定温度
                                m_CtlSystem(i).Text = (.iTempTemp).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 13  ' 代表α値(ppm/℃)
                                m_CtlSystem(i).Text = (.dDaihyouAlpha).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 14  ' 代表β値(ppm/℃)
                                m_CtlSystem(i).Text = (.dDaihyouBeta).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                                'V2.1.0.0③↓
                            Case 15 ' 温度センサー情報一元管理選択番号追加 CTxt_37 以降Case番号１つ加算
                                m_CtlSystem(i).Text = (.iTempSensorInfNoDaihyou).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                                'V2.1.0.0③↑
                            Case 16  ' ０℃ 'V2.1.0.0③Case 13から16へ移動
                                m_CtlSystem(i).Text = (.dTemperatura0).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 17  ' α値(ppm/℃)
                                m_CtlSystem(i).Text = (.dAlpha).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 18  ' β値(ppm/℃)
                                m_CtlSystem(i).Text = (.dBeta).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                                'V2.1.0.0③↓
                            Case 19 ' 温度センサー情報一元管理選択番号追加 CTxt_38 以降Case番号１つ加算
                                m_CtlSystem(i).Text = (.iTempSensorInfNoStd).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                                'V2.1.0.0③↑
                                'V2.0.0.0⑪                            Case 11 ' 参照温度	１：０℃ または ２：２
                                'V2.0.0.0⑪                                Dim cCombo As cCmb_ = DirectCast(m_CtlSystem(i), cCmb_)
                                'V2.0.0.0⑪                                Call NoEventIndexChange(cCombo, (.iTempTemp - 1))
                                'V2.0.0.0⑪                            Case 12 ' 標準抵抗値 ０℃ 0.01～100M
                                'V2.0.0.0⑪                                m_CtlSystem(i).Text = (.dStandardRes0).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                                'V2.0.0.0⑪                            Case 13 ' 標準抵抗値 ２５℃ 0.01～100M
                                'V2.0.0.0⑪                                m_CtlSystem(i).Text = (.dStandardRes25).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 20 ' ファイナルリミット　Hight[%]
                                m_CtlSystem(i).Text = (.dFinalLimitHigh).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 21 ' ファイナルリミット　Lo[%]
                                m_CtlSystem(i).Text = (.dFinalLimitLow).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 22 ' 相対値リミット　Hight[%]
                                m_CtlSystem(i).Text = (.dRelativeHigh).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 23 ' 相対値リミット　Lo[%]
                                m_CtlSystem(i).Text = (.dRelativeLow).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case Else
                                Throw New Exception("Case " & i & ": Nothing")
                        End Select
                    Next i

                    For i As Integer = 0 To (m_CtlResistor.GetLength(0) - 1) Step 1
                        For j As Integer = 0 To (m_CtlResistor.GetLength(1) - 1) Step 1
                            Select Case (j)
                                Case 0 ' 補正値 抵抗レンジ 1:Ω, 2:KΩ
                                    Dim cCombo As cCmb_ = DirectCast(m_CtlResistor(i, j), cCmb_)
                                    Call NoEventIndexChange(cCombo, (.iResUnit(i + 1) - 1))
                                Case 1 ' 補正値（ノミナル値算出係数）
                                    m_CtlResistor(i, j).Text = (.dNomCalcCoff(i + 1)).ToString(DirectCast(m_CtlResistor(i, j), cTxt_).GetStrFormat())
                                Case 2 ' 目標値算出係数
                                    m_CtlResistor(i, j).Text = (.dTargetCoff(i + 1)).ToString(DirectCast(m_CtlResistor(i, j), cTxt_).GetStrFormat())
                                    'V2.1.0.0①↓
                                Case 3 ' 判定用目標値算出係数
                                    m_CtlResistor(i, j).Text = (.dTargetCoffJudge(i + 1)).ToString(DirectCast(m_CtlResistor(i, j), cTxt_).GetStrFormat())
                                    'V2.1.0.0①↑
                                Case 4 ' 測定速度を変更するカットＮｏ．'V2.1.0.0①　Case 3からCase 4へ変更
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

#Region "抵抗数に応じてｺﾝﾄﾛｰﾙの表示/非表示を切り替える"
        Private Sub setDispResistor()
            Dim ResCnt As Short   'V1.2.0.0②

            'V1.2.0.0②↓
            ResCnt = UserBas.GetRCountExceptMeasure()
            'V2.0.0.0⑩            If UserSub.IsTrimType3 Then
            'V2.0.0.0⑩                ResCnt = 1                  ' チップは、抵抗１つのデータで処理する。
            'V2.0.0.0⑩            End If
            'V1.2.0.0②↑

            ' 抵抗数分のみ表示する
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

                ' 補正値ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                ret = CheckControlData(m_CtlResistor)
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

                        Case 1 To 3 ' ｼｽﾃﾑｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            With .W_stUserData
                                Select Case (tag)
                                    Case 0 ' オペレータ名
                                        ret = CheckStrData(cTextBox, .sOperator)
                                    Case 1 ' レーザ　ロット番号
                                        ret = CheckStrData(cTextBox, .sLotNumber)
                                    Case 2 ' パターンＮｏ．
                                        ret = CheckStrData(cTextBox, .sPatternNo)
                                    Case 3 ' プログラムＮｏ．
                                        ret = CheckStrData(cTextBox, .sProgramNo)
                                    Case 4 ' 処理枚数
                                        ret = CheckLongData(cTextBox, .lLotEndSL)
                                    Case 5 ' カット位置補正頻度
                                        ret = CheckLongData(cTextBox, .lCutHosei)
                                    Case 6 ' ロット終了時印刷素子数
                                        ret = CheckLongData(cTextBox, .lPrintRes)
                                        'V2.0.0.0⑪                                    Case 7 ' 標準抵抗値	０℃0.01～100M
                                        'V2.0.0.0⑪                                        ret = CheckDoubleData(cTextBox, .dStandardRes0)
                                        'V2.0.0.0⑪                                        If ret = 0 Then
                                        'V2.0.0.0⑪                                            .dResTempCoff = GetResTempCoff(.dStandardRes0, .dStandardRes25)
                                        'V2.0.0.0⑪                                            LabelTempCoff.Text = (.dResTempCoff).ToString("0.000")
                                        'V2.0.0.0⑪                                        End If
                                        'V2.0.0.0⑪                                    Case 8 ' 標準抵抗値	２５℃ 0.01～100M
                                        'V2.0.0.0⑪                                        ret = CheckDoubleData(cTextBox, .dStandardRes25)
                                        'V2.0.0.0⑪                                        If ret = 0 Then
                                        'V2.0.0.0⑪                                            .dResTempCoff = GetResTempCoff(.dStandardRes0, .dStandardRes25)
                                        'V2.0.0.0⑪                                            LabelTempCoff.Text = (.dResTempCoff).ToString("0.000")
                                        'V2.0.0.0⑪                                        End If
                                    Case 7  ' 設定温度
                                        ret = CheckIntData(cTextBox, .iTempTemp)
                                    Case 8  ' 代表α値(ppm/℃)
                                        ret = CheckDoubleData(cTextBox, .dDaihyouAlpha)
                                    Case 9  ' 代表β値(ppm/℃)
                                        ret = CheckDoubleData(cTextBox, .dDaihyouBeta)
                                        'V2.1.0.0③↓
                                    Case 10  ' 代表温度センサー情報一元管理選択番号
                                        Dim iSaveNo As Integer = .iTempSensorInfNoDaihyou
                                        ret = CheckIntData(cTextBox, .iTempSensorInfNoDaihyou)
                                        If ret = 0 And .iTempSensorInfNoDaihyou > 0 Then ' 温度情報更新
                                            Dim dDummy As Double
                                            If TemperatureTableDataGet(.iTempSensorInfNoDaihyou, dDummy, .dDaihyouAlpha, .dDaihyouBeta) Then
                                                Call SetDataToText()
                                            Else
                                                Call MsgBox_CheckErr(cTextBox, "温度センサー情報が取得できませんでした。No=[" & .iTempSensorInfNoDaihyou.ToString("0") & "]", iSaveNo.ToString())
                                                .iTempSensorInfNoDaihyou = iSaveNo
                                            End If
                                        End If
                                        'V2.1.0.0③↑
                                    Case 11  ' ０℃   'V2.1.0.0③ Case8から11へ移動
                                        ret = CheckDoubleData(cTextBox, .dTemperatura0)
                                    Case 12  ' α値(ppm/℃)
                                        ret = CheckDoubleData(cTextBox, .dAlpha)
                                    Case 13  ' β値(ppm/℃)
                                        ret = CheckDoubleData(cTextBox, .dBeta)
                                        'V2.1.0.0③↓
                                    Case 14  ' 温度センサー情報一元管理選択番号
                                        Dim iSaveNo As Integer = .iTempSensorInfNoStd
                                        ret = CheckIntData(cTextBox, .iTempSensorInfNoStd)
                                        If ret = 0 And .iTempSensorInfNoStd > 0 Then ' 温度情報更新
                                            If TemperatureTableDataGet(.iTempSensorInfNoStd, .dTemperatura0, .dAlpha, .dBeta) Then
                                                Call SetDataToText()
                                            Else
                                                Call MsgBox_CheckErr(cTextBox, "温度センサー情報が取得できませんでした。No=[" & .iTempSensorInfNoStd.ToString("0") & "]", iSaveNo.ToString())
                                                .iTempSensorInfNoStd = iSaveNo
                                            End If
                                        End If
                                        'V2.1.0.0③↑
                                    Case Else
                                        Throw New Exception("Case " & tag & ": Nothing")
                                End Select
                            End With
                            ' ------------------------------------------------------------------------------
                        Case 4
                            With .W_stUserData
                                Select Case (tag)
                                    Case 0 To 3 'V2.1.0.0①2から3へ変更 ' 補正値から測定速度を変更するカットＮｏ．まで（SetControlDataで設定されるテキストボックスの順番、４以降とはセットされるコントロールが異なる）
                                        Dim bStop As Boolean
                                        bStop = False
                                        For i As Integer = 0 To (m_CtlResistor.GetLength(0) - 1) Step 1
                                            For j As Integer = 0 To (m_CtlResistor.GetLength(1) - 1) Step 1
                                                If (m_CtlResistor(i, j) Is cTextBox) Then
                                                    Select Case (j)
                                                        Case 1 ' 補正値
                                                            ret = CheckDoubleData(cTextBox, .dNomCalcCoff(i + 1))
                                                        Case 2 ' 目標値算出係数                                                        
                                                            ret = CheckDoubleData(cTextBox, .dTargetCoff(i + 1))
                                                            'V2.1.0.0①↓
                                                        Case 3 ' 判定用目標値算出係数                                                        
                                                            ret = CheckDoubleData(cTextBox, .dTargetCoffJudge(i + 1))
                                                            'V2.1.0.0①↑
                                                        Case 4 ' 測定速度を変更するカットＮｏ．                'V2.1.0.0① Case 3からCase 4へ変更                        
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
                                        'V2.1.0.0③ No.のテキスト項目が２つ追加したので以降２つずつ加算 13→15
                                    Case 15 ' ファイナルリミット　Hight[%]
                                        ret = CheckDoubleData(cTextBox, .dFinalLimitHigh)
                                    Case 16 ' ファイナルリミット　Lo[%]
                                        ret = CheckDoubleData(cTextBox, .dFinalLimitLow)
                                    Case 17 ' 相対値リミット　Hight[%]
                                        ret = CheckDoubleData(cTextBox, .dRelativeHigh)
                                    Case 18 ' 相対値リミット　Lo[%]
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
                                Case 0 ' 製品種別( 0:温度センサー, 1:抵抗トリミング )
                                    .W_stUserData.iTrimType = Convert.ToInt16(idx)
                                Case 1 ' トリミング速度(1:高速, 2:高精度，3:設定値)
                                    .W_stUserData.iTrimSpeed = Convert.ToInt16(idx + 1)
                                Case 2 ' ロット終了条件(0:終了条件判定無し,1:枚数,2:ローダ信号,3:枚数＆信号)
                                    .W_stUserData.iLotChange = Convert.ToInt16(idx)
                                Case 3 ' クランプと吸着の有り無し                               'V2.0.0.0⑭
                                    .W_stUserData.intClampVacume = Convert.ToInt16(idx + 1)     'V2.0.0.0⑭
                                Case 4 ' 抵抗レンジ(1:Ω, 2:KΩ)
                                    .W_stUserData.iTempResUnit = Convert.ToInt16(idx + 1)
                                    'V2.0.0.0⑪                                Case 4 ' 参照温度(１：０℃ または ２：２５℃)
                                    'V2.0.0.0⑪                                    .W_stUserData.iTempTemp = Convert.ToInt16(idx + 1)
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                        Case 4 ' 抵抗設定
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
