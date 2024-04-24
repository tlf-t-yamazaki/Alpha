Option Explicit On
Option Strict Off

Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports LaserFront.Trimmer.DefWin32Fnc

Namespace FormEdit
    Friend Class tabSystem
        Inherits tabBase

#Region "宣言"
        Private Const SYS_ZON As Integer = 11       ' 相関ﾁｪｯｸで使用(m_CtlSystemでのｲﾝﾃﾞｯｸｽ)
        Private Const SYS_ZOFF As Integer = 12      ' 相関ﾁｪｯｸで使用(m_CtlSystemでのｲﾝﾃﾞｯｸｽ)

        Private m_CtlSystem() As Control            ' ｼｽﾃﾑｸﾞﾙｰﾌﾟﾎﾞｯｸｽのｺﾝﾄﾛｰﾙ配列
        Private m_CtlLaser() As Control             ' ﾚｰｻﾞｰﾊﾟﾜｰ調整ｸﾞﾙｰﾌﾟﾎﾞｯｸｽのｺﾝﾄﾛｰﾙ配列
        Private m_CtlTeachBlock() As Control        ' ###1040 ティーチング・ブロックｸﾞﾙｰﾌﾟﾎﾞｯｸｽのｺﾝﾄﾛｰﾙ配列
        Private m_CtlStageSpeed() As Control        ' ###1040 ステージ速度ｸﾞﾙｰﾌﾟﾎﾞｯｸｽのｺﾝﾄﾛｰﾙ配列
        Private m_CtlDisMagnify() As Control        ' 'V2.2.0.0② デジタルカメラ表示倍率 ｸﾞﾙｰﾌﾟﾎﾞｯｸｽのｺﾝﾄﾛｰﾙ配列

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
                TAB_NAME = GetPrivateProfileString_S("SYSTEM_LABEL", "TAB_NAM", m_sPath, "????")

                ' 発振器種別がﾌｧｲﾊﾞﾚｰｻﾞの場合はﾚｰｻﾞｰﾊﾟﾜｰ調整ｸﾞﾙｰﾌﾟﾎﾞｯｸｽを非表示にする
                If (gSysPrm.stRAT.giOsc_Res = OSCILLATOR_FL) Then
                    CGrp_5.Visible = False
                End If

                ' ----------------------------------------------------------
                ' EDIT_DEF_User.iniからｸﾞﾙｰﾌﾟﾎﾞｯｸｽに表示名を設定 ###1040 Grp_6, CGrp_7追加     ' V2.2.0.0② CGrp_8追加
                ' ----------------------------------------------------------
                'V2.2.2.0① ↓
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
                'V2.2.2.0① ↑
                For i As Integer = 0 To (GrpArray.Length - 1) Step 1
                    With GrpArray(i)
                        .TabIndex = ((254 - GrpArray.Length) + i) ' ｶｰｿﾙｷｰによるﾌｫｰｶｽ移動で必要
                        .Tag = 0
                        .Text = GetPrivateProfileString_S(
                            "SYSTEM_LABEL", (i.ToString("000") & "_GRP"), m_sPath, "????")
                    End With
                Next i
                ' ｸﾞﾙｰﾌﾟﾎﾞｯｸｽが入れ子になっているため
                ' ｼｽﾃﾑとｼｽﾃﾑ内=0, ﾚｰｻﾞﾊﾟﾜｰ調整=1とする
                CGrp_5.Tag = 1
                CGrp_6.Tag = 2      ' ###1040
                CGrp_7.Tag = 3      ' ###1040
                CGrp_8.Tag = 4      ' V2.2.0.0② 

                ' 追加･削除ﾎﾞﾀﾝのﾊﾟﾈﾙ (ﾎﾞﾀﾝなし)
                'CPnl_Btn.TabIndex = 254 ' ｺﾝﾄﾛｰﾙ配置可能最大数(最後に設定)

                ' ----------------------------------------------------------
                ' EDIT_DEF_User.iniからﾗﾍﾞﾙに表示名を設定 ###1040 CLbl_12, CLbl_13, CLbl_14 追加
                ' 'V1.2.0.0① CLbl_15チップサイズ追加  'V2.2.0.0② CLbl_17追加 
                ' 'V2.2.0.0⑮ CLbl_18ﾌﾟﾛｰﾌﾞの追加       
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
                ' ｼｽﾃﾑｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のｺﾝﾄﾛｰﾙを設定(CtlArrayの順番と合わせる)
                ' V1.2.0.0① チップサイズ CTxt_20,CTxt_21追加
                ' V2.2.0.0⑮ CTxt_24：プローブの追加
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
                ' ﾚｰｻﾞｰﾊﾟﾜｰ調整ｸﾞﾙｰﾌﾟﾎﾞｯｸｽのｺﾝﾄﾛｰﾙを設定(CtlArrayの順番と合わせる)
                ' ----------------------------------------------------------
                'V2.1.0.0②　CTxt_22(アッテネータNo.)追加　
                m_CtlLaser = New Control() { _
                    CTxt_14, CTxt_15, CTxt_16, CTxt_22 _
                }
                Call SetControlData(m_CtlLaser)

                ' ----------------------------------------------------------------------------------------
                ' ###1040 ティーチング・ブロックｸﾞﾙｰﾌﾟﾎﾞｯｸｽのｺﾝﾄﾛｰﾙを設定(CtlArrayの順番と合わせる)
                ' ----------------------------------------------------------------------------------------
                m_CtlTeachBlock = New Control() { _
                    CTxt_17, CTxt_18 _
                }
                Call SetControlData(m_CtlTeachBlock)

                ' ----------------------------------------------------------------------------------------
                ' ###1040 ステージ速度ｸﾞﾙｰﾌﾟﾎﾞｯｸｽのｺﾝﾄﾛｰﾙを設定(CtlArrayの順番と合わせる)
                ' ----------------------------------------------------------------------------------------
                m_CtlStageSpeed = New Control() { _
                    CTxt_19 _
                }
                Call SetControlData(m_CtlStageSpeed)
                ' ----------------------------------------------------------------------------------------
                ' ' 'V2.2.0.0② 品種ｸﾞﾙｰﾌﾟﾎﾞｯｸｽのｺﾝﾄﾛｰﾙを設定(CtlArrayの順番と合わせる)
                ' ----------------------------------------------------------------------------------------
                m_CtlDisMagnify = New Control() {
                    CTxt_23
                }
                Call SetControlData(m_CtlDisMagnify)

                ' ----------------------------------------------------------
                ' すべてのﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽ･ﾎﾞﾀﾝに対し、ﾌｫｰｶｽ移動設定をおこなう
                ' 使用しないｺﾝﾄﾛｰﾙは Enabled=False または Visible=False にする
                ' ###1040 CTxt_16～19まで追加
                ' V1.2.0.0① チップサイズ CTxt_20,CTxt_21追加
                ' V2.1.0.0②　CTxt_22(アッテネータNo.)追加　
                ' 'V2.2.0.0② CTxt_23:デジタルカメラ表示倍率追加 
                ' 'V2.2.0.0⑮ CTxt_24：プローブNo追加
                ' ---------------------------------------------------------- 
                CtlArray = New Control() {
                    CTxt_0, CTxt_1, CTxt_2, CTxt_3, CTxt_4, CTxt_20, CTxt_21,
                    CTxt_5, CTxt_6, CTxt_7, CTxt_8, CTxt_9, CTxt_10,
                    CTxt_11, CTxt_12, CCmb_0, CTxt_24,
                    CTxt_13,
                    CTxt_14, CTxt_15, CTxt_16, CTxt_22, CTxt_17, CTxt_18, CTxt_19, CTxt_23
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
                        Case 0 ' ｼｽﾃﾑｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            Select Case (tag)
                                Case 0 ' ﾌﾟﾛｰﾌﾞﾘﾄﾗｲ
                                    .Items.Add("なし")
                                    .Items.Add("あり")
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case Else
                            Throw New Exception("Parent.Tag - Case Else")
                    End Select
                    .SelectedIndex = 0

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
            Try
                tag = DirectCast(cTextBox.Tag, Integer)
                no = tag.ToString("000")

                Select Case (DirectCast(cTextBox.Parent.Tag, Integer)) ' ｸﾞﾙｰﾌﾟﾎﾞｯｸｽのﾀｸﾞ
                    ' ------------------------------------------------------------------------------
                    Case 0 ' ｼｽﾃﾑｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                        ' ﾃﾞｰﾀﾁｪｯｸｴﾗｰ時の表示名
                        strMsg = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0  ' ﾌﾞﾛｯｸｻｲｽﾞX
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "0.001")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "80.0")
                            Case 1  ' ﾌﾞﾛｯｸｻｲｽﾞＹ
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "0.001")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "80.0")
                            Case 2  ' ﾌﾞﾛｯｸ数X
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "50")
                            Case 3  ' ﾌﾞﾛｯｸ数Y
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "50")
                            Case 4 ' 抵抗数
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "20")
                                If Integer.Parse(strMax) > MAXRNO Then
                                    strMax = MAXRNO.ToString
                                End If
                            Case 5  ' ﾃｰﾌﾞﾙ位置ｵﾌｾｯﾄX
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "-245.000")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "245.000")
                            Case 6  ' ﾃｰﾌﾞﾙ位置ｵﾌｾｯﾄY
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "-245.000")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "245.000")
                            Case 7  ' ﾋﾞｰﾑ位置ｵﾌｾｯﾄX
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "-80.000")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "80.000")
                            Case 8  ' ﾋﾞｰﾑ位置ｵﾌｾｯﾄY
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "-80.000")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "80.000")
                                'V2.0.0.0①                            Case 9 ' ｱｼﾞｬｽﾄﾎﾟｲﾝﾄX
                                'V2.0.0.0①                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "0.001")
                                'V2.0.0.0①                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "80.000")
                                'V2.0.0.0①                            Case 10 ' ｱｼﾞｬｽﾄﾎﾟｲﾝﾄY
                                'V2.0.0.0①                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "0.001")
                                'V2.0.0.0①                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "80.000")
                                'V2.0.0.0①ｱｼﾞｬｽﾄﾎﾟｲﾝﾄXからステップオフセットXへ変更↓
                            Case 9 ' ステップオフセットX 
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "-80.000")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "80.000")
                            Case 10 ' ステップオフセットY
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "-80.000")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "80.000")
                                'V2.0.0.0①↑
                            Case 11  ' ﾌﾟﾛｰﾌﾞ接触位置
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "0.000")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "40.000")
                            Case 12  ' ﾌﾟﾛｰﾌﾞ待機位置
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "0.000")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "40.000")
                            Case 13 ' 外部機器数
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "10")
                                'V1.2.0.0①↓
                            Case 14  ' チップサイズX
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "0.0001")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "50.0")
                            Case 15  ' チップサイズY
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "0.0001")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "50.0")
                                'V1.2.0.0①↑
                                'V2.2.0.0⑮ ↓
                            Case 16
                                ' プローブ番号 
                                strMin = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("SYSTEM_SYSTEM", (no & "_MAX"), m_sPath, "10")
                                'V2.2.0.0⑮ ↑
                            Case Else
                                Throw New Exception("Case " & tag & ": Nothing")
                        End Select
                        ' ------------------------------------------------------------------------------
                    Case 1 ' ﾚｰｻﾞｰﾊﾟﾜｰ調整ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                        If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then ' ﾌｧｲﾊﾞﾚｰｻﾞｰでない場合
                            ' ﾃﾞｰﾀﾁｪｯｸｴﾗｰ時の表示名
                            strMsg = GetPrivateProfileString_S("SYSTEM_LASER", (no & "_MSG"), m_sPath, "??????")
                            Select Case (tag)
                                Case 0 ' Qﾚｰﾄ
                                    strMin = GetPrivateProfileString_S("SYSTEM_LASER", (no & "_MIN"), m_sPath, "0.1")
                                    strMax = GetPrivateProfileString_S("SYSTEM_LASER", (no & "_MAX"), m_sPath, "40.0")
                                Case 1 ' 設定ﾊﾟﾜｰ
                                    strMin = GetPrivateProfileString_S("SYSTEM_LASER", (no & "_MIN"), m_sPath, "0.1")
                                    strMax = GetPrivateProfileString_S("SYSTEM_LASER", (no & "_MAX"), m_sPath, "10.0")
                                Case 2 ' アッテネータ減衰率（0:保存無 1:保存）###1040③
                                    strMin = GetPrivateProfileString_S("SYSTEM_LASER", (no & "_MIN"), m_sPath, "0")
                                    strMax = GetPrivateProfileString_S("SYSTEM_LASER", (no & "_MAX"), m_sPath, "1")
                                Case 3  'V2.1.0.0②↓CTxt_22(アッテネータNo.)追加
                                    strMin = GetPrivateProfileString_S("SYSTEM_LASER", (no & "_MIN"), m_sPath, "0")
                                    strMax = GetPrivateProfileString_S("SYSTEM_LASER", (no & "_MAX"), m_sPath, "99")
                                    Dim TempNo As Integer = Integer.Parse(strMax)
                                    Dim TableNo As Integer = UserSub.LaserCalibrationMaxNumberGet()
                                    If TableNo < TempNo Then
                                        strMax = TableNo.ToString("0")
                                    End If
                                    'V2.1.0.0②↑
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                        End If
                        ' ------------------------------------------------------------------------------
                    Case 2 ' ティーチング・ブロックｸﾞﾙｰﾌﾟﾎﾞｯｸｽ ###1040①
                        strMsg = GetPrivateProfileString_S("SYSTEM_TEACHBLOCK", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0  ' ティーチング・ブロックＸ
                                strMin = GetPrivateProfileString_S("SYSTEM_TEACHBLOCK", (no & "_MIN"), m_sPath, "1")
                                strMax = m_MainEdit.W_PLT.BNX
                            Case 1  ' ティーチング・ブロックＹ
                                strMin = GetPrivateProfileString_S("SYSTEM_TEACHBLOCK", (no & "_MIN"), m_sPath, "1")
                                strMax = m_MainEdit.W_PLT.BNY
                            Case Else
                                Throw New Exception("Case " & tag & ": Nothing")
                        End Select
                    Case 3 ' ステージ速度ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ ###1040④
                        strMsg = GetPrivateProfileString_S("SYSTEM_STAGESPEED", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0  ' Ｙ軸
                                strMin = GetPrivateProfileString_S("SYSTEM_STAGESPEED", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("SYSTEM_STAGESPEED", (no & "_MAX"), m_sPath, "50")
                            Case Else
                                Throw New Exception("Case " & tag & ": Nothing")
                        End Select
                    Case 4  '品種グループボックス         'V2.2.0.0②
                        strMsg = GetPrivateProfileString_S("SYSTEM_KIND", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0  ' 表示倍率 
                                strMin = GetPrivateProfileString_S("SYSTEM_KIND", (no & "_MIN"), m_sPath, "0.5")
                                strMax = GetPrivateProfileString_S("SYSTEM_KIND", (no & "_MAX"), m_sPath, "2.0")
                            Case Else
                                Throw New Exception("Case " & tag & ": Nothing")
                        End Select

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
        ''' <summary>ﾃｷｽﾄﾎﾞｯｸｽに値を設定する</summary>
        Protected Overrides Sub SetDataToText()
            Try
                ' ｼｽﾃﾑｸﾞﾙｰﾌﾟﾎﾞｯｸｽ設定
                Call SetSystemData()

                ' ﾚｰｻﾞﾊﾟﾜｰ調整ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ設定
                Call SetLaserData()

                Call SetTeachBlockData()    ' ###1040 ティーチング・ブロック

                Call SetStageSpeedData()    ' ###1040 ステージ速度

                Call SetKindData()          ' 'V2.2.0.0② 表示倍率 

                Me.Refresh()

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

#Region "ｼｽﾃﾑｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内の設定"
        ''' <summary>ｼｽﾃﾑｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽに値を設定する</summary>
        Private Sub SetSystemData()
            Try
                With m_MainEdit.W_PLT
                    For i As Integer = 0 To (m_CtlSystem.Length - 1) Step 1
                        Select Case (i)
                            Case 0 ' ﾌﾞﾛｯｸｻｲｽﾞX(mm)
                                m_CtlSystem(i).Text = (.zsx).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 1 ' ﾌﾞﾛｯｸｻｲｽﾞY(mm)
                                m_CtlSystem(i).Text = (.zsy).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 2 ' ﾌﾞﾛｯｸ数X
                                m_CtlSystem(i).Text = (.BNX).ToString()
                            Case 3 ' ﾌﾞﾛｯｸ数Y
                                m_CtlSystem(i).Text = (.BNY).ToString()
                            Case 4 ' 抵抗数
                                m_CtlSystem(i).Text = (.RCount).ToString()
                            Case 5  ' ﾃｰﾌﾞﾙ位置ｵﾌｾｯﾄX(mm)
                                m_CtlSystem(i).Text = (.z_xoff).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 6 ' ﾃｰﾌﾞﾙ位置ｵﾌｾｯﾄY(mm)
                                m_CtlSystem(i).Text = (.z_yoff).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 7 ' ﾋﾞｰﾑ位置ｵﾌｾｯﾄX(mm)
                                m_CtlSystem(i).Text = (.BPOX).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 8 ' ﾋﾞｰﾑ位置ｵﾌｾｯﾄY(mm)
                                m_CtlSystem(i).Text = (.BPOY).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                                'V2.0.0.0①                            Case 9 ' ｱｼﾞｬｽﾄﾎﾟｲﾝﾄX
                                'V2.0.0.0①                                m_CtlSystem(i).Text = (.ADJX).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                                'V2.0.0.0①                            Case 10 ' ｱｼﾞｬｽﾄﾎﾟｲﾝﾄY
                                'V2.0.0.0①                                m_CtlSystem(i).Text = (.ADJY).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                                'V2.0.0.0①↓
                            Case 9 ' ステップオフセット量X
                                m_CtlSystem(i).Text = (.dblStepOffsetXDir).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 10 ' ステップオフセット量Y
                                m_CtlSystem(i).Text = (.dblStepOffsetYDir).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                                'V2.0.0.0①↑
                            Case 11  ' ﾌﾟﾛｰﾌﾞ接触位置(mm)
                                m_CtlSystem(i).Text = (.Z_ZON).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 12 ' ﾌﾟﾛｰﾌﾞ待機位置(mm)
                                m_CtlSystem(i).Text = (.Z_ZOFF).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 13 ' ﾌﾟﾛｰﾌﾞﾘﾄﾗｲ(0:無, 1:有)
                                Call NoEventIndexChange(DirectCast(m_CtlSystem(i), cCmb_), .PrbRetry)
                            Case 14 ' 外部機器数
                                m_CtlSystem(i).Text = (.GCount).ToString()
                                'V1.2.0.0①↓
                            Case 15 ' チップサイズX(mm)
                                m_CtlSystem(i).Text = (.dblChipSizeXDir).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            Case 16 ' チップサイズY(mm)
                                m_CtlSystem(i).Text = (.dblChipSizeYDir).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                                'V1.2.0.0①↑
                            '    'V2.2.0.0②↓
                            'Case 17 ' 表示倍率
                            '    m_CtlSystem(i).Text = (.dblStdMagnification).ToString(DirectCast(m_CtlSystem(i), cTxt_).GetStrFormat())
                            '    'V2.2.0.0②↑
                                'V2.2.0.0⑮ ↓
                            Case 17
                                m_CtlSystem(i).Text = (.ProbNo).ToString()
                                'V2.2.0.0⑮ ↑
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

#Region "ﾚｰｻﾞﾊﾟﾜｰ調整ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内の設定"
        ''' <summary>ﾚｰｻﾞﾊﾟﾜｰ調整ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽに値を設定する</summary>
        Private Sub SetLaserData()
            Try
                With m_MainEdit.W_LASER
                    For i As Integer = 0 To (m_CtlLaser.Length - 1) Step 1
                        Select Case (i)
                            Case 0 ' Qﾚｰﾄ
                                m_CtlLaser(i).Text = (.intQR / 10).ToString(DirectCast(m_CtlLaser(i), cTxt_).GetStrFormat())
                            Case 1 ' 設定ﾊﾟﾜｰ(W)
                                m_CtlLaser(i).Text = (.dblspecPower).ToString(DirectCast(m_CtlLaser(i), cTxt_).GetStrFormat())
                            Case 2 ' アッテネータ減衰率（0:保存無 1:保存）###1040
                                m_CtlLaser(i).Text = (.iTrimAtt).ToString(DirectCast(m_CtlLaser(i), cTxt_).GetStrFormat())
                            Case 3  'V2.1.0.0②　アッテネータNo.
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
#Region "ティーチング・ブロックｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内の設定"
        ''' <summary>ティーチング・ブロックｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽに値を設定する</summary>
        Private Sub SetTeachBlockData()
            Try
                With m_MainEdit.W_PLT
                    For i As Integer = 0 To (m_CtlTeachBlock.Length - 1) Step 1
                        Select Case (i)
                            Case 0 ' ブロックＸ
                                m_CtlTeachBlock(i).Text = (.TeachBlockX).ToString(DirectCast(m_CtlTeachBlock(i), cTxt_).GetStrFormat())
                            Case 1 ' ブロックＹ
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
#Region "ステージ速度ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内の設定"
        ''' <summary>ステージ速度ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽに値を設定する</summary>
        Private Sub SetStageSpeedData()
            Try
                With m_MainEdit.W_PLT
                    For i As Integer = 0 To (m_CtlStageSpeed.Length - 1) Step 1
                        Select Case (i)
                            Case 0 ' Ｙ軸速度
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
#Region "品種ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内の設定"
        Private Sub SetKindData()
            Try
                With m_MainEdit.W_PLT
                    For i As Integer = 0 To (m_CtlDisMagnify.Length - 1) Step 1
                        Select Case (i)
                            Case 0 ' 表示倍率
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

                ' ティーチング・ブロックｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                ret = CheckControlData(m_CtlTeachBlock) ' ###1040① 
                If (ret <> 0) Then Exit Try ' ###1040①

                ' 相関ﾁｪｯｸ
                ret = CheckRelation()
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
                        Case 0 ' ｼｽﾃﾑｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            With .W_PLT
                                Select Case (tag)
                                    Case 0 ' ﾌﾞﾛｯｸｻｲｽﾞX
                                        ret = CheckDoubleData(cTextBox, .zsx)
                                    Case 1 ' ﾌﾟﾛｯｸｻｲｽﾞY
                                        ret = CheckDoubleData(cTextBox, .zsy)
                                    Case 2 ' ﾌﾞﾛｯｸ数X
                                        ret = CheckShortData(cTextBox, .BNX)
                                    Case 3 ' ﾌﾞﾛｯｸ数Y
                                        ret = CheckShortData(cTextBox, .BNY)
                                    Case 4 ' 抵抗数
                                        Dim cnt As Integer = .RCount ' 変更前の値を保持
                                        'V2.0.0.0⑩↓
                                        Dim RCount As Short
                                        Dim FromNo As Short = 1
                                        Dim CircuitNo As Integer
                                        If m_MainEdit.W_stUserData.iTrimType = 3 Then
                                            RCount = UserSub.CircuitResistorCount(m_MainEdit.W_PLT, m_MainEdit.W_REG)
                                        Else

                                            'V2.0.0.0①↓
                                            FromNo = 1
                                            For j As Integer = 1 To cnt
                                                If UserModule.IsCutResistor(m_MainEdit.W_REG, j) Then
                                                    FromNo = j
                                                    Exit For
                                                End If
                                            Next
                                            'V2.0.0.0①↑
                                        End If
                                        'V2.0.0.0⑩↑
                                        ret = CheckShortData(cTextBox, .RCount)
                                        If (cnt <> .RCount) Then
                                            If (cnt < .RCount) Then ' 追加された場合
                                                If m_MainEdit.W_stUserData.iTrimType = 3 Then
                                                    cnt = RCount
                                                End If
                                                For i As Integer = (cnt + 1) To .RCount Step 1
                                                    Call InitResData(i) ' 追加されたﾃﾞｰﾀを初期化
                                                    'V2.0.0.0⑩↓
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
                                                        'V2.0.0.0⑩↑
                                                        'V1.2.0.0⑤↓
                                                        m_MainEdit.W_REG(i).Initialize()
                                                        'V2.0.0.0①コピー元を１からFromNoへ変更
                                                        Call CopyResistorData(m_MainEdit.W_REG(i), m_MainEdit.W_REG(FromNo))
                                                        m_MainEdit.W_PTN(i).PtnFlg = m_MainEdit.W_PTN(FromNo).PtnFlg
                                                        m_MainEdit.W_PTN(i).intGRP = m_MainEdit.W_PTN(FromNo).intGRP
                                                        m_MainEdit.W_PTN(i).intPTN = m_MainEdit.W_PTN(FromNo).intPTN
                                                        'V1.2.0.0⑤↑
                                                    End If                                                                  'V2.0.0.0⑩
                                                Next i
                                                'V2.0.0.0①↓
                                                If m_MainEdit.W_stUserData.iTrimType = 1 Or m_MainEdit.W_stUserData.iTrimType = 4 Then
                                                    For i As Integer = 1 To .RCount
                                                        If UserModule.IsCutResistor(m_MainEdit.W_REG, i) Then
                                                            m_MainEdit.W_REG(i).intCircuitNo = i
                                                        End If
                                                    Next
                                                End If
                                                'V2.0.0.0①↑
                                            Else ' 削除された場合
                                                For i As Integer = (.RCount + 1) To cnt Step 1
                                                    Call InitResData(i) ' 削除されたﾃﾞｰﾀを初期化
                                                Next i
                                            End If
                                            m_ResNo = 1 ' 処理中の抵抗番号
                                        End If

                                    Case 5 ' ﾃｰﾌﾞﾙ位置ｵﾌｾｯﾄX
                                        ret = CheckDoubleData(cTextBox, .z_xoff)
                                    Case 6 ' ﾃｰﾌﾞﾙ位置ｵﾌｾｯﾄY
                                        ret = CheckDoubleData(cTextBox, .z_yoff)
                                    Case 7 ' ﾋﾞｰﾑ位置ｵﾌｾｯﾄX
                                        ret = CheckDoubleData(cTextBox, .BPOX)
                                    Case 8 ' ﾋﾞｰﾑ位置ｵﾌｾｯﾄY
                                        ret = CheckDoubleData(cTextBox, .BPOY)
                                        'V2.0.0.0①                                    Case 9 ' ｱｼﾞｬｽﾄﾎﾟｲﾝﾄX
                                        'V2.0.0.0①                                        ret = CheckDoubleData(cTextBox, .ADJX)
                                        'V2.0.0.0①                                    Case 10 ' ｱｼﾞｬｽﾄﾎﾟｲﾝﾄY
                                        'V2.0.0.0①                                        ret = CheckDoubleData(cTextBox, .ADJY)
                                        'V2.0.0.0①↓
                                    Case 9 ' ステップオフセットX
                                        ret = CheckDoubleData(cTextBox, .dblStepOffsetXDir)
                                    Case 10 ' ステップオフセットY
                                        ret = CheckDoubleData(cTextBox, .dblStepOffsetYDir)
                                        'V2.0.0.0①↑
                                    Case 11 ' ﾌﾟﾛｰﾌﾞ接触位置
                                        ret = CheckDoubleData(cTextBox, .Z_ZON)
                                    Case 12 ' ﾌﾟﾛｰﾌﾞ待機位置
                                        ret = CheckDoubleData(cTextBox, .Z_ZOFF)
                                    Case 13 ' 外部機器数
                                        Dim cnt As Integer = .GCount ' 変更前の値を保持
                                        ret = CheckShortData(cTextBox, .GCount)
                                        If (cnt <> .GCount) Then
                                            If (cnt < .GCount) Then ' 追加された場合
                                                For i As Integer = (cnt + 1) To .GCount Step 1
                                                    Call InitGpibData(i) ' 追加されたﾃﾞｰﾀを初期化
                                                Next i
                                            Else ' 削除された場合
                                                For i As Integer = (.GCount + 1) To cnt Step 1
                                                    Call InitGpibData(i) ' 削除されたﾃﾞｰﾀを初期化
                                                Next i
                                            End If
                                            m_GpibNo = 1 ' 処理中のGP-IB登録番号
                                        End If
                                        'V1.2.0.0①↓
                                    Case 14 ' チップサイズX
                                        ret = CheckDoubleData(cTextBox, .dblChipSizeXDir)
                                    Case 15 ' チップサイズY
                                        ret = CheckDoubleData(cTextBox, .dblChipSizeYDir)
                                    'V2.2.0.0⑮↓
                                    Case 16 ' ﾌﾟﾛｰﾌﾞNo
                                        ret = CheckShortData(cTextBox, .ProbNo)
                                        'V2.2.0.0⑮↑
                                    Case Else
                                        Throw New Exception("Case " & tag & ": Nothing")
                                        'V1.2.0.0①↑
                                End Select
                            End With
                            ' ------------------------------------------------------------------------------
                        Case 1 ' ﾚｰｻﾞﾊﾟﾜｰ調整ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then ' ﾌｧｲﾊﾞﾚｰｻﾞｰでない場合
                                Select Case (tag)
                                    Case 0 ' Qﾚｰﾄ
                                        Dim dblWK As Double
                                        ret = CheckDoubleData(cTextBox, dblWK)
                                        .W_LASER.intQR = Convert.ToInt16(dblWK * 10) ' KHz → 0.1KHz
                                    Case 1 ' 設定ﾊﾟﾜｰ
                                        ret = CheckDoubleData(cTextBox, .W_LASER.dblspecPower)
                                    Case 2 ' アッテネータ減衰率（0:保存無 1:保存）
                                        ret = CheckShortData(cTextBox, .W_LASER.iTrimAtt)
                                    Case 3 'V2.1.0.0②　アッテネータNo.
                                        ret = CheckShortData(cTextBox, .W_LASER.iAttNo)
                                    Case Else
                                        Throw New Exception("Case " & tag & ": Nothing")
                                End Select
                            End If
                            ' ------------------------------------------------------------------------------
                        Case 2 ' ###1040 ティーチング・ブロックｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            With .W_PLT
                                Call SetControlData(m_CtlTeachBlock)    ' ブロック数により上限リミットを変更する。
                                Select Case (tag)
                                    Case 0 ' ティーチングブロック位置Ｘ
                                        ret = CheckShortData(cTextBox, .TeachBlockX)
                                    Case 1 ' ティーチングブロック位置Ｙ
                                        ret = CheckShortData(cTextBox, .TeachBlockY)
                                End Select
                            End With
                            ' ------------------------------------------------------------------------------
                        Case 3 ' ###1040 ステージ速度ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            With .W_PLT
                                Select Case (tag)
                                    Case 0 ' ステージ・スピードＹ
                                        ret = CheckShortData(cTextBox, .StageSpeedY)
                                End Select
                            End With
                            ' ------------------------------------------------------------------------------
                        Case 4 'V2.2.0.0② 品種ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            With .W_PLT
                                Select Case (tag)
                                    Case 0 ' 表示倍率
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

#Region "相関ﾁｪｯｸ"
        ''' <summary>相関ﾁｪｯｸ処理</summary>
        ''' <returns>0 = 正常, 1 = ｴﾗｰ</returns>
        Protected Overrides Function CheckRelation() As Integer
            Dim strMsg As String
            Dim errIdx As Integer
            CheckRelation = 0 ' Return値 = 正常
            Try
                ' ﾌﾟﾛｰﾌﾞOFF位置(mm) >= ﾌﾟﾛｰﾌﾞON位置(mm) ?
                With m_MainEdit
                    If (.W_PLT.Z_ZOFF >= .W_PLT.Z_ZON) Then
                        errIdx = SYS_ZOFF
                        strMsg = "相関チェックエラー" & vbCrLf
                        strMsg = strMsg & DirectCast(m_CtlSystem(SYS_ZOFF), cTxt_).GetStrMsg() & " < " _
                                        & DirectCast(m_CtlSystem(SYS_ZON), cTxt_).GetStrMsg() & _
                                        "となるように指定してください。"
                        GoTo STP_ERR
                    End If
                End With
                Exit Function
STP_ERR:
                Call MsgBox_CheckErr(DirectCast(m_CtlSystem(errIdx), cTxt_), strMsg)
                CheckRelation = 1 ' Return値 = ｴﾗｰ

            Catch ex As Exception
                Call MsgBox_Exception(ex)
                CheckRelation = 1 ' Return値 = ｴﾗｰ
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
                        Case 0 ' ｼｽﾃﾑｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            Select Case (tag)
                                Case 0 ' ﾌﾟﾛｰﾌﾞﾘﾄﾗｲ(1:有, 0:無)
                                    .W_PLT.PrbRetry = Convert.ToInt16(idx)
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case 1 ' ﾚｰｻﾞﾊﾟﾜｰ調整ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
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
                Dim FromNo As Short     'V2.0.0.0⑩
                Dim OrderNo As Short    'V2.0.0.0⑩
                '--------------------------------------------------------------------------
                '   確認ﾒｯｾｰｼﾞを表示する
                '--------------------------------------------------------------------------
                Dim strMsg As String = "チップサイズを反映します。よろしいですか？"
                ret = MsgBox(strMsg, DirectCast( _
                            MsgBoxStyle.OkCancel + _
                            MsgBoxStyle.Information, MsgBoxStyle), _
                            My.Application.Info.Title)
                If (ret = MsgBoxResult.Cancel) Then ' Cancel(RESETｷｰ) ?
                    Exit Sub
                End If

                'V2.0.0.0①↓
                m_MainEdit.W_PLT.dblStepOffsetXDir = 0.0                       ' ステップオフセット量X
                m_MainEdit.W_PLT.dblStepOffsetYDir = 0.0                       ' ステップオフセット量Y
                'V2.0.0.0①↑

                'V2.0.0.0⑩↓
                If m_MainEdit.W_stUserData.iTrimType = 3 And UserSub.CircuitResistorCount(m_MainEdit.W_PLT, m_MainEdit.W_REG) > 1 Then
                    Dim RCount As Short = UserSub.CircuitResistorCount(m_MainEdit.W_PLT, m_MainEdit.W_REG)
                    With m_MainEdit
                        For rn As Integer = RCount + 1 To .W_PLT.RCount Step 1
                            With .W_REG(rn)
                                If .intSLP <> SLP_OK_MARK And .intSLP <> SLP_NG_MARK Then   ' OK,NGマーキングは実施コピーしない
                                    OrderNo = UserSub.GetResNumberInCircuit(m_MainEdit.W_REG, rn)   ' サーキット内の抵抗の順番
                                    FromNo = UserSub.GetRNumByCircuit(m_MainEdit.W_PLT, m_MainEdit.W_REG, 1, OrderNo)   ' 第一サーキット内の順番の抵抗番号
                                    ' ｶｯﾄ数分繰返す
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
                    'V2.0.0.0⑩↑

                    FromNo = 0
                    Dim Circuit As Short = 0
                    With m_MainEdit
                        'V2.0.0.0①                        For rn As Integer = 2 To .W_PLT.RCount Step 1
                        For rn As Integer = 1 To .W_PLT.RCount Step 1
                            With .W_REG(rn)
                                'V2.0.0.0①                                If .intSLP <> SLP_OK_MARK And .intSLP <> SLP_NG_MARK Then   ' OK,NGマーキングは実施コピーしない
                                If UserModule.IsCutResistor(stREG, rn) Then
                                    If FromNo = 0 Then
                                        FromNo = rn
                                        Circuit = 2
                                        Continue For
                                    End If
                                    ' ｶｯﾄ数分繰返す
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
                                    'V1.2.0.0⑤↓
                                    m_MainEdit.W_PTN(rn).dblPosX = stPTN(FromNo).dblPosX + m_MainEdit.W_PLT.dblChipSizeXDir * (Circuit - 1)
                                    m_MainEdit.W_PTN(rn).dblPosY = stPTN(FromNo).dblPosY + m_MainEdit.W_PLT.dblChipSizeYDir * (Circuit - 1)
                                    'V1.2.0.0⑤↑
                                    Circuit = Circuit + 1
                                End If
                            End With
                        Next rn
                    End With
                End If      'V2.0.0.0⑩

                m_MainEdit.LblToolTip.Text = "チップサイズの反映が正常に終了しました。"
            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
    End Class
End Namespace
