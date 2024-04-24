Option Explicit On
Option Strict Off

Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports LaserFront.Trimmer.DefWin32Fnc

Namespace FormEdit
    Friend Class tabResistor
        Inherits tabBase

#Region "宣言"
        '0:抵抗
        '1:カット数
        '2:抵抗名
        '3:スロープ
        '4:リレービット
        '5:目標値
        '6:表示単位
        '7:判定モード
        '8:その2
        '9:IT下限
        '10:IT上限
        '11:FT下限
        '12:FT上限
        '13:測定機器
        '14:測定モード
        '15:再測定回数
        '16:ポーズ時間
        '17:IT回数
        '18:FT回数
        '19サーキット
        Private Const RES_NOM As Integer = 5    ' m_CtlResでのｲﾝﾃﾞｯｸｽ(目標値)
        Private Const RES_SLOPE As Integer = 3  ' m_CtlResでのｲﾝﾃﾞｯｸｽ(スロープ)
        Private Const RES_MTYPE As Integer = 13 ' m_CtlResでのｲﾝﾃﾞｯｸｽ(測定機器)
        Private Const RES_TMM1 As Integer = 14  ' m_CtlResでのｲﾝﾃﾞｯｸｽ(測定ﾓｰﾄﾞ)
        Private Const PRB_PRH As Integer = 0    ' m_CtlProbeでのｲﾝﾃﾞｯｸｽ(HI側ﾌﾟﾛｰﾌﾞ)
        Private Const PRB_PRL As Integer = 1    ' m_CtlProbeでのｲﾝﾃﾞｯｸｽ(LO側ﾌﾟﾛｰﾌﾞ)

        Private m_voltNOM_Min As String         ' ﾄﾘﾐﾝｸﾞ 目標値(V)
        Private m_voltNOM_Max As String
        Private m_ohmNOM_Min As String          ' ﾄﾘﾐﾝｸﾞ 目標値(Ω)
        Private m_ohmNOM_Max As String
        Private m_ITH_Min As String             ' 初期判定値(ITHI)(%用)
        Private m_ITH_Max As String
        Private m_ITL_Min As String             ' 初期判定値(ITLO)(%用)
        Private m_ITL_Max As String
        Private m_FTH_Min As String             ' 終了判定値(FTHI)(%用)
        Private m_FTH_Max As String
        Private m_FTL_Min As String             ' 終了判定値(FTLO)(%用)
        Private m_FTL_Max As String

        Private GRP_MIN As Integer              ' ｶｯﾄ位置補正ｸﾞﾙｰﾌﾟ番号最小値
        Private GRP_MAX As Integer              ' ｶｯﾄ位置補正ｸﾞﾙｰﾌﾟ番号最大値
        Private PTN_MIN As Integer              ' ｶｯﾄ位置補正ﾊﾟﾀｰﾝ番号最小値
        Private PTN_MAX As Integer              ' ｶｯﾄ位置補正ﾊﾟﾀｰﾝ番号最大値

        Private m_CtlRes() As Control           ' 抵抗ｸﾞﾙｰﾌﾟﾎﾞｯｸｽのｺﾝﾄﾛｰﾙ配列
        Private m_CtlProbe() As Control         ' ﾌﾟﾛｰﾌﾞｸﾞﾙｰﾌﾟﾎﾞｯｸｽのｺﾝﾄﾛｰﾙ配列
        Private m_CtlCutCorr() As Control       ' ｶｯﾄ位置補正ｸﾞﾙｰﾌﾟﾎﾞｯｸｽのｺﾝﾄﾛｰﾙ配列

        Private m_IntialFinal() As cTxt_        ' IT,FTﾃｷｽﾄﾎﾞｯｸｽ配列
        Private m_CutPosCorr() As Control       ' ｶｯﾄ位置補正切替え時に有効･無効にするｺﾝﾄﾛｰﾙ

        'V2.0.0.0↓
        ''' <summary>
        ''' スロープコンボボックスデータリスト
        ''' </summary>
        ''' <remarks></remarks>
        Private m_lstSlope As New List(Of ComboDataStruct)

        ''' <summary>
        ''' 測定モードコンボボックスデータリスト
        ''' </summary>
        ''' <remarks></remarks>
        Private m_lstMeasMode As New List(Of ComboDataStruct)
        'V2.0.0.0↑
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

            ' スロープデータ初期化
            Call InitSlopeData()        'V2.0.0.0

            ' 測定モードデータ初期化
            Call InitMeasModeData()     'V2.0.0.0

            Try
                ' EDIT_DEF_User.iniからﾀﾌﾞ名を設定
                TAB_NAME = GetPrivateProfileString_S("RESISTOR_LABEL", "TAB_NAM", m_sPath, "????")

                ' ｶｯﾄ位置補正ｸﾞﾙｰﾌﾟ番号･ﾊﾟﾀｰﾝ番号の上下限値
                GRP_MIN = Convert.ToInt32(GetPrivateProfileString_S("PATTERN_LABEL", "GRP_MIN", m_sPath, "1"))
                GRP_MAX = Convert.ToInt32(GetPrivateProfileString_S("PATTERN_LABEL", "GRP_MAX", m_sPath, "999"))
                PTN_MIN = Convert.ToInt32(GetPrivateProfileString_S("PATTERN_LABEL", "PTN_MIN", m_sPath, "1"))
                PTN_MAX = Convert.ToInt32(GetPrivateProfileString_S("PATTERN_LABEL", "PTN_MAX", m_sPath, "50"))

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
                    CGrp_0, CGrp_1, CGrp_2 _
                }
                For i As Integer = 0 To (GrpArray.Length - 1) Step 1
                    With GrpArray(i)
                        .TabIndex = ((254 - GrpArray.Length) + i) ' ｶｰｿﾙｷｰによるﾌｫｰｶｽ移動で必要
                        .Tag = i
                        .Text = GetPrivateProfileString_S( _
                                "RESISTOR_LABEL", (i.ToString("000") & "_GRP"), m_sPath, "????")
                    End With
                Next i
                ' 追加･削除ﾎﾞﾀﾝのﾊﾟﾈﾙ
                CPnl_Btn.TabIndex = 254 ' ｺﾝﾄﾛｰﾙ配置可能最大数(最後に設定)

                ' ----------------------------------------------------------
                ' EDIT_DEF_User.iniからﾗﾍﾞﾙに表示名を設定
                ' ----------------------------------------------------------
                LblArray = New cLbl_() { _
                    CLbl_0, CLbl_1, CLbl_2, CLbl_3, CLbl_4, CLbl_5, _
                    CLbl_6, CLbl_7, _
                    CLbl_8, CLbl_9, CLbl_10, CLbl_11, CLbl_12, _
                    CLbl_13, CLbl_14, CLbl_24, CLbl_25, _
                    CLbl_22, CLbl_28, CLbl_29, _
                    CLbl_26, CLbl_27, _
                    CLbl_16, CLbl_15, CLbl_17, _
                    CLbl_18, CLbl_19, CLbl_20, CLbl_21 _
                }
                For i As Integer = 0 To (LblArray.GetLength(0) - 1) Step 1
                    LblArray(i).Text = GetPrivateProfileString_S( _
                            "RESISTOR_LABEL", (i.ToString("000") & "_LBL"), m_sPath, "????")
                Next i

                ' ----------------------------------------------------------
                ' 抵抗ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のｺﾝﾄﾛｰﾙを設定(CtlArrayの順番と合わせる)
                ' ----------------------------------------------------------
                m_CtlRes = New Control() { _
                    CCmb_0, CTxt_0, CTxt_1, CCmb_1, CTxt_2, _
                    CTxt_3, CTxt_4, _
                    CCmb_2, CCmb_16, CTxt_5, CTxt_6, CTxt_7, CTxt_8, _
                    CCmb_3, CCmb_4, CTxt_14, CTxt_15, CTxt_16, CTxt_17, CTxt_18, _
                    CCmb_10, CCmb_11, CCmb_12, CCmb_13, CCmb_14, CCmb_15 _
                }
                Call SetControlData(m_CtlRes)

                ' IT, FT関連のﾃｷｽﾄﾎﾞｯｸｽ配列
                m_IntialFinal = New cTxt_() { _
                    CTxt_5, CTxt_6, CTxt_7, CTxt_8 _
                }

                ' ----------------------------------------------------------
                ' ﾌﾟﾛｰﾌﾞｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のｺﾝﾄﾛｰﾙを設定(CtlArrayの順番と合わせる)
                ' ----------------------------------------------------------
                m_CtlProbe = New Control() { _
                    CTxt_9, CTxt_10, CTxt_11 _
                }
                Call SetControlData(m_CtlProbe)

                ' ----------------------------------------------------------
                ' ｶｯﾄ位置補正ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のｺﾝﾄﾛｰﾙを設定(CtlArrayの順番と合わせる)
                ' ----------------------------------------------------------
                m_CtlCutCorr = New Control() { _
                    CCmb_5, CCmb_6, CCmb_7, CTxt_12, CTxt_13 _
                }
                Call SetControlData(m_CtlCutCorr)

                m_CutPosCorr = New Control() { _
                    CCmb_6, CCmb_7, CTxt_12, CTxt_13 _
                }

                ' ----------------------------------------------------------
                ' すべてのﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽ･ﾎﾞﾀﾝに対し、ﾌｫｰｶｽ移動設定をおこなう
                ' ﾀﾌﾞｷｰ、ｶｰｿﾙｷｰによりﾌｫｰｶｽ移動する順番でｺﾝﾄﾛｰﾙをCtlArrayに設定する
                ' 使用しないｺﾝﾄﾛｰﾙは Enabled=False または Visible=False にする
                ' ----------------------------------------------------------
                CtlArray = New Control() { _
                    CCmb_0, CTxt_0, CTxt_1, CCmb_1, CTxt_2, _
                    CTxt_3, CTxt_4, _
                    CCmb_2, CCmb_16, CTxt_5, CTxt_6, CTxt_7, CTxt_8, _
                    CCmb_3, CCmb_4, CTxt_14, CTxt_15, CTxt_16, CTxt_17, CTxt_18, _
 _
                    CTxt_10, CTxt_9, CTxt_11, _
 _
                    CCmb_5, CCmb_6, CCmb_7, CTxt_12, CTxt_13, _
 _
                    CBtn_Add, CBtn_Del _
                }
                Call SetTabIndex(CtlArray) ' ﾀﾌﾞｲﾝﾃﾞｯｸｽとKeyDownｲﾍﾞﾝﾄを設定する

                ' ----------------------------------------------------------
                ' 画面表示時にﾌｫｰｶｽされるｺﾝﾄﾛｰﾙを設定する
                ' ----------------------------------------------------------
                FIRST_CONTROL = CtlArray(0)

                ' ImeModeの切替を有効にするためｲﾍﾞﾝﾄを設定する
                CTxt_4.ImeMode = Windows.Forms.ImeMode.Off  ' ﾃﾞﾌｫﾙﾄは英字入力とする
                AddHandler CTxt_4.Validating, AddressOf MyBase.cTxt_Validating  ' 表示単位

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
            Dim i As Integer

            Try
                With cCombo
                    tag = DirectCast(.Tag, Integer)
                    .Items.Clear()
                    Select Case (DirectCast(.Parent.Tag, Integer))
                        ' ------------------------------------------------------------------------------
                        Case 0 ' 抵抗ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            Select Case (tag)
                                Case 0 ' 抵抗番号
                                    '.Items.Add("") ' ﾚｲｱｳﾄｲﾍﾞﾝﾄで再設定される
                                Case 1 ' ｽﾛｰﾌﾟ
                                    For i = 0 To Me.m_lstSlope.Count - 1 Step 1
                                        .Items.Add(Me.m_lstSlope(i).Name)
                                    Next i

                                Case 2 ' 判定ﾓｰﾄﾞ(0:比率(%), 1:数値(絶対値))
                                    .Items.Add("比率(%)")
                                    .Items.Add("絶対値")
                                    'V2.0.0.0↓

                                Case 3  ' 測定モード
                                    For i = 0 To Me.m_lstMeasMode.Count - 1 Step 1
                                        .Items.Add(Me.m_lstMeasMode(i).Name)
                                    Next i
                                    'V2.0.0.0↑


                                Case 4 ' 測定機器'(0:内部測定器, 1:外部測定器)
                                    '.Items.Add("") ' ﾚｲｱｳﾄｲﾍﾞﾝﾄで再設定される
                                Case 5 ' 測定ﾓｰﾄﾞ 
                                    .Items.Add("高速")
                                    .Items.Add("高精度")
                                    'V2.0.0.0↓
                                Case 6 ' ON機器
                                    '.Items.Add("") ' ﾚｲｱｳﾄｲﾍﾞﾝﾄで設定される
                                Case 7 ' ON機器
                                    '.Items.Add("") ' ﾚｲｱｳﾄｲﾍﾞﾝﾄで設定される
                                Case 8 ' ON機器
                                    '.Items.Add("") ' ﾚｲｱｳﾄｲﾍﾞﾝﾄで設定される
                                Case 9 ' OFF機器
                                    '.Items.Add("") ' ﾚｲｱｳﾄｲﾍﾞﾝﾄで設定される
                                Case 10 ' OFF機器
                                    '.Items.Add("") ' ﾚｲｱｳﾄｲﾍﾞﾝﾄで設定される
                                Case 11 ' OFF機器
                                    '.Items.Add("") ' ﾚｲｱｳﾄｲﾍﾞﾝﾄで設定される
                                    'V2.0.0.0↑
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case 1 ' ﾌﾟﾛｰﾌﾞｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            Throw New Exception("Parent.Tag - Case 1")
                            ' ------------------------------------------------------------------------------
                        Case 2 ' ｶｯﾄ位置補正ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            Select Case (tag)
                                Case 0 ' 補正実行(0:なし, 1:自動, 2:手動)
                                    .Items.Add("なし")
                                    .Items.Add("自動")
                                    .Items.Add("手動")
                                    .Items.Add("自動ＮＧ判定あり")      'V1.0.4.3⑥
                                Case 1 ' ｸﾞﾙｰﾌﾟ番号(1-999)
                                    For i = GRP_MIN To GRP_MAX Step 1
                                        .Items.Add(String.Format("{0,5:#0}", i))
                                    Next i
                                Case 2 ' ﾊﾟﾀｰﾝ番号(1-50)
                                    For i = PTN_MIN To PTN_MAX Step 1
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
        ''' <summary>初期化時にﾃｷｽﾄﾎﾞｯｸｽの設定をおこなう</summary>
        ''' <param name="cTextBox">設定をおこなうﾃｷｽﾄﾎﾞｯｸｽ</param>
        Protected Overrides Sub InitTextBox(ByRef cTextBox As cTxt_)
            Dim strMin As String = ""           ' 設定する変数の最大値
            Dim strMax As String = ""           ' 設定する変数の最小値
            Dim strMsg As String = ""           ' ｴﾗｰで表示する項目名
            Dim no As String = ""
            Dim tag As Integer
            Dim strFlg As Boolean = False       ' 格納する値の種類(False=数値,True=文字列)
            Dim hexFlg As Boolean = False       ' 格納する文字列の種類(False=10進数,True=16進数)
            Try
                With cTextBox
                    tag = DirectCast(.Tag, Integer)
                    no = tag.ToString("000")
                    Select Case (DirectCast(.Parent.Tag, Integer))
                        ' ------------------------------------------------------------------------------
                        Case 0 ' 抵抗ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            ' ﾃﾞｰﾀﾁｪｯｸｴﾗｰ時の表示名
                            strMsg = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MSG"), m_sPath, "??????")
                            Select Case (tag)
                                Case 0 ' ｶｯﾄ数
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "1")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "10")
                                    If Integer.Parse(strMax) > MAXCTN Then
                                        strMax = MAXCTN.ToString
                                    End If
                                Case 1 ' 抵抗名
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "1")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "4")
                                    strFlg = True
                                Case 2 ' ﾘﾚｰﾋﾞｯﾄ
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "0")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "16777215")
                                    strFlg = True
                                    hexFlg = True
                                Case 3 ' 目標値
                                    ' (V)
                                    m_voltNOM_Min = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_VMN"), m_sPath, "-32.0")
                                    m_voltNOM_Max = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_VMX"), m_sPath, "32.0")
                                    ' (Ω)
                                    m_ohmNOM_Min = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_RMN"), m_sPath, "0.1")
                                    m_ohmNOM_Max = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_RMX"), m_sPath, "60000000.0")
                                    ' 初期値として電圧の目標値を設定する
                                    strMin = m_voltNOM_Min
                                    strMax = m_voltNOM_Max
                                Case 4 ' 表示単位
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "1")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "2")
                                    strFlg = True
                                Case 5 ' IT 下限値(%用)
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "-99.99")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "99.99")
                                    m_ITL_Min = strMin
                                    m_ITL_Max = strMax
                                Case 6 ' IT 上限値(%用)
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "-99.99")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "99.99")
                                    m_ITH_Min = strMin
                                    m_ITH_Max = strMax
                                Case 7 ' FT 下限値(%用)
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "-99.99")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "99.99")
                                    m_FTL_Min = strMin
                                    m_FTL_Max = strMax
                                Case 8 ' FT 上限値(%用)
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "-99.99")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "99.99")
                                    m_FTH_Min = strMin
                                    m_FTH_Max = strMax
                                    'V2.0.0.0↓
                                Case 9  ' 再測定回数
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "0")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "10")
                                Case 10 ' 再測定までのﾎﾟｰｽﾞ時間(ms)
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "0")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "32767")
                                    'V2.0.0.0↑
                                    'V2.0.0.0⑧↓
                                Case 11  ' IT測定回数
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "0")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "99")
                                Case 12 ' FT測定回数
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "0")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "99")
                                    'V2.0.0.0⑧↑
                                    'V2.0.0.0⑩↓
                                Case 13 ' サーキット番号
                                    strMin = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MIN"), m_sPath, "0")
                                    strMax = GetPrivateProfileString_S("RESISTOR_RESISTOR", (no & "_MAX"), m_sPath, "99")
                                    'V2.0.0.0⑩↑
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case 1 ' ﾌﾟﾛｰﾌﾞｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            ' ﾃﾞｰﾀﾁｪｯｸｴﾗｰ時の表示名
                            strMsg = GetPrivateProfileString_S("RESISTOR_PROBE", (no & "_MSG"), m_sPath, "??????")
                            Select Case (tag)
                                Case 0 ' LO側番号
                                    strMin = GetPrivateProfileString_S("RESISTOR_PROBE", (no & "_MIN"), m_sPath, "0")
                                    strMax = GetPrivateProfileString_S("RESISTOR_PROBE", (no & "_MAX"), m_sPath, "255")
                                Case 1 ' HI側番号
                                    strMin = GetPrivateProfileString_S("RESISTOR_PROBE", (no & "_MIN"), m_sPath, "0")
                                    strMax = GetPrivateProfileString_S("RESISTOR_PROBE", (no & "_MAX"), m_sPath, "255")
                                Case 2 ' ｶﾞｰﾄﾞ番号
                                    strMin = GetPrivateProfileString_S("RESISTOR_PROBE", (no & "_MIN"), m_sPath, "0")
                                    strMax = GetPrivateProfileString_S("RESISTOR_PROBE", (no & "_MAX"), m_sPath, "255")
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case 2 ' ｶｯﾄ位置補正ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            ' ﾃﾞｰﾀﾁｪｯｸｴﾗｰ時の表示名
                            strMsg = GetPrivateProfileString_S("RESISTOR_CUTCORR", (no & "_MSG"), m_sPath, "??????")
                            Select Case (tag)
                                Case 0 ' ﾊﾟﾀｰﾝ位置X
                                    strMin = GetPrivateProfileString_S("RESISTOR_CUTCORR", (no & "_MIN"), m_sPath, "-80.0")
                                    strMax = GetPrivateProfileString_S("RESISTOR_CUTCORR", (no & "_MAX"), m_sPath, "80.0")
                                Case 1 ' ﾊﾟﾀｰﾝ位置Y
                                    strMin = GetPrivateProfileString_S("RESISTOR_CUTCORR", (no & "_MIN"), m_sPath, "-80.0")
                                    strMax = GetPrivateProfileString_S("RESISTOR_CUTCORR", (no & "_MAX"), m_sPath, "80.0")
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select                            ' ------------------------------------------------------------------------------
                        Case Else
                            Throw New Exception("Parent.Tag - Case Else")
                    End Select

                    Call .SetStrMsg(strMsg) ' ﾃﾞｰﾀﾁｪｯｸｴﾗｰ時の表示名設定
                    Call .SetMinMax(strMin, strMax) ' 下限値･上限値の設定
                    Dim msg As String
                    If (False = strFlg) Then ' (False=数値,True=文字列)
                        msg = "の範囲で指定して下さい"
                    Else
                        If (False = hexFlg) Then ' (True=16進数文字列)
                            msg = "文字の範囲で指定して下さい"
                            .MaxLength = Integer.Parse(strMax) ' SetControlData()内の条件判断で使用する
                            .TextAlign = HorizontalAlignment.Left
                        Else ' 16進数文字列
                            msg = "の範囲で指定して下さい"
                            ' 10進数文字列を16進数文字列に変換した文字列の文字数
                            .MaxLength = ((Integer.Parse(strMax)).ToString("X")).Length
                            ' ﾂｰﾙﾁｯﾌﾟ用に変換
                            strMin = ((Integer.Parse(strMin)).ToString("X")).ToUpper & "(Hex)"
                            strMax = ((Integer.Parse(strMax)).ToString("X")).ToUpper & "(Hex)"
                        End If
                    End If
                    Call .SetStrTip(strMin & "～" & strMax & msg) ' ﾂｰﾙﾁｯﾌﾟﾒｯｾｰｼﾞの設定
                    Call .SetLblToolTip(m_MainEdit.LblToolTip) ' ﾒｲﾝ編集画面のﾂｰﾙﾁｯﾌﾟ参照設定
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#Region "ﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽに値を表示する"
        ''' <summary>ﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽに値を表示する</summary>
        Protected Overrides Sub SetDataToText()
            Try
                If (m_MainEdit.W_PLT.RCount < 1) Then ' 抵抗数 = 0 ?
                    m_ResNo = 1
                End If

                Call ChangeSlopeList()      'V2.2.1.7①

                ' 抵抗ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ設定
                Call SetResData()

                ' ﾌﾟﾛｰﾌﾞｸﾞﾙｰﾌﾟﾎﾞｯｸｽ設定
                Call SetProbeData()

                ' ｶｯﾄ位置補正ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ設定
                Call SetCutPosData()

                Me.Refresh()

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#Region "抵抗ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内の設定"
        ''' <summary>抵抗ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽに値を設定する</summary>
        Private Sub SetResData()
            Dim idx As Integer

            Try
                With m_MainEdit.W_REG(m_ResNo)
                    For i As Integer = 0 To (m_CtlRes.Length - 1) Step 1
                        Select Case (i)
                            Case 0 ' 抵抗数, 抵抗番号
                                Dim rCnt As Integer = m_MainEdit.W_PLT.RCount
                                Dim cCombo As cCmb_ = DirectCast(m_CtlRes(i), cCmb_)

                                CLblRN_0.Text = rCnt.ToString() ' 抵抗数
                                With cCombo ' 抵抗番号
                                    .Items.Clear()
                                    For j As Integer = 1 To rCnt Step 1
                                        .Items.Add(String.Format("{0,5:#0}", j)) ' 総抵抗数分繰り返す
                                    Next j
                                End With
                                Call NoEventIndexChange(cCombo, (m_ResNo - 1)) ' 指定抵抗番号を設定

                            Case 1 ' ｶｯﾄ数
                                ' Case 3 ｽﾛｰﾌﾟで設定をおこなう
                                'm_CtlRes(i).Text = (.intTNN).ToString()
                            Case 2  ' 抵抗名
                                m_CtlRes(i).Text = .strRNO
                            Case 3 ' ｽﾛｰﾌﾟ(1:+電圧, 2:-電圧, 4:抵抗, 5:電圧測定のみ, 6:抵抗測定のみ, 7:NGﾏｰｷﾝｸﾞ)
                                idx = GetComboBoxValue2Index(.intSLP, Me.m_lstSlope)

                                Call NoEventIndexChange(DirectCast(m_CtlRes(i), cCmb_), idx)
                                ' 上下限値を変更する必要がないためｺﾒﾝﾄｱｳﾄ
                                'Call ChangedSlope(.intMode, .intSLP)

                                ' ｶｯﾄ数ﾃｷｽﾄﾎﾞｯｸｽの設定
                                If UserModule.IsMarking(m_MainEdit.W_REG, m_ResNo) Then
                                    'm_MainEdit.W_REG(m_ResNo).intMeasMode = MEAS_JUDGE_NONE
                                    ChangeSlopeAllOnOff(False)
                                Else
                                    'm_MainEdit.W_REG(m_ResNo).intMeasMode = MEAS_JUDGE_FT
                                    ChangeSlopeAllOnOff(True)
                                End If

                                If UserModule.IsMeasureOnly(m_MainEdit.W_REG, m_ResNo) Then
                                    ' ｽﾛｰﾌﾟが 7:電圧測定のみ, 9:抵抗測定のみ の場合
                                    CTxt_0.Text = 0                     ' ｶｯﾄ数を0とする
                                    CTxt_0.Enabled = False              ' 無効にする
                                Else
                                    CTxt_0.Text = (.intTNN).ToString()  ' ｶｯﾄ数
                                    CTxt_0.Enabled = True               ' 有効にする
                                End If

                            Case 4 ' ﾘﾚｰﾋﾞｯﾄ
                                m_CtlRes(i).Text = (.lngRel.ToString("X")).ToUpper
                            Case 5 ' 目標値
                                m_CtlRes(i).Text = (.dblNOM).ToString(DirectCast(m_CtlRes(i), cTxt_).GetStrFormat())
                                ' ｽﾛｰﾌﾟにより範囲を再設定する必要がある
                                Call ChangedSlope(.intMode, .intSLP, .dblNOM)

                            Case 6 ' 表示単位
                                m_CtlRes(i).Text = .strTANI
                            Case 7 ' 判定ﾓｰﾄﾞ(0:比率(%), 1:絶対値)
                                Call NoEventIndexChange(DirectCast(m_CtlRes(i), cCmb_), .intMode)
                                Call ChangedMode(.intMode, .intSLP, .dblNOM)

                            Case 8  ' 測定モード
                                idx = GetComboBoxValue2Index(.intMeasMode, Me.m_lstMeasMode)

                                Call NoEventIndexChange(DirectCast(m_CtlRes(i), cCmb_), idx)

                            Case 9 ' IT下限値
                                m_CtlRes(i).Text = (.dblITL).ToString(DirectCast(m_CtlRes(i), cTxt_).GetStrFormat())
                            Case 10 ' IT上限値
                                m_CtlRes(i).Text = (.dblITH).ToString(DirectCast(m_CtlRes(i), cTxt_).GetStrFormat())
                            Case 11 ' FT下限値
                                m_CtlRes(i).Text = (.dblFTL).ToString(DirectCast(m_CtlRes(i), cTxt_).GetStrFormat())
                            Case 12 ' FT上限値
                                m_CtlRes(i).Text = (.dblFTH).ToString(DirectCast(m_CtlRes(i), cTxt_).GetStrFormat())
                            Case 13 ' 測定機器(0=内部測定, 1以上外部測定機器番号)
                                ' GP-IB登録機器名を表示する(外部電源を除く)
                                Dim cCombo As cCmb_ = DirectCast(m_CtlRes(i), cCmb_)
                                Dim type As Integer = Convert.ToInt32(.intMType)
                                Dim cnt As Integer = 0 ' ﾘｽﾄに追加した項目数
                                idx = 0 ' 選択するｲﾝﾃﾞｯｸｽ
                                cCombo.Items.Clear()
                                cCombo.Items.Add(" 0:内部測定器")
                                With m_MainEdit
                                    If (0 < .W_PLT.GCount) Then ' GP-IB測定機器が登録されている場合
                                        For j As Integer = 1 To (.W_PLT.GCount) Step 1
                                            ' ﾄﾘｶﾞｰｺﾏﾝﾄﾞありの場合、外部測定器としてﾘｽﾄに追加
                                            If (.W_GPIB(j).strCTRG <> "") Then
                                                If (Not .W_GPIB(j).strGNAM Is Nothing) Then
                                                    cCombo.Items.Add(String.Format("{0,2:#0}", j) & ":" & .W_GPIB(j).strGNAM)
                                                Else
                                                    cCombo.Items.Add(String.Format("{0,2:#0}", j) & ":")
                                                End If
                                                ' 追加したﾘｽﾄをｶｳﾝﾄｱｯﾌﾟ
                                                cnt = (cnt + 1)
                                                ' .intType(GP-IB登録番号)と同じ項目がﾘｽﾄに追加された場合に
                                                ' その項目を選択するためｲﾝﾃﾞｯｸｽを設定する
                                                ' 使用中の機器が削除された場合、GP-IBﾀﾌﾞ内の処理で
                                                ' .intMTypeが0となるため内部測定器が選択される
                                                If (type = j) Then idx = cnt
                                            End If
                                        Next j

                                        If (0 < idx) Then
                                            m_CtlRes(RES_TMM1).Enabled = False ' 測定ﾓｰﾄﾞ無効
                                        Else
                                            m_CtlRes(RES_TMM1).Enabled = True ' 測定ﾓｰﾄﾞ有効
                                        End If

                                    Else ' GP-IB測定機器の登録がない場合
                                        .W_REG(m_ResNo).intMType = 0
                                        idx = 0 ' 内部測定器
                                        m_CtlRes(RES_TMM1).Enabled = True ' 測定ﾓｰﾄﾞ有効
                                    End If
                                End With
                                Call NoEventIndexChange(cCombo, idx)

                            Case 14 ' 測定ﾓｰﾄﾞ
                                Call NoEventIndexChange(DirectCast(m_CtlRes(i), cCmb_), .intTMM1)
                                'V2.0.0.0↓
                            Case 15 ' 再測定回数
                                m_CtlRes(i).Text = (.intReMeas).ToString(DirectCast(m_CtlRes(i), cTxt_).GetStrFormat())
                            Case 16 ' 再測定までのﾎﾟｰｽﾞ時間
                                m_CtlRes(i).Text = (.intReMeas_Time).ToString(DirectCast(m_CtlRes(i), cTxt_).GetStrFormat())
                                'V2.0.0.0↑
                                'V2.0.0.0⑧↓
                            Case 17 ' IT測定回数
                                m_CtlRes(i).Text = (.intITReMeas).ToString(DirectCast(m_CtlRes(i), cTxt_).GetStrFormat())
                            Case 18 ' FT測定回数
                                m_CtlRes(i).Text = (.intFTReMeas).ToString(DirectCast(m_CtlRes(i), cTxt_).GetStrFormat())
                                'V2.0.0.0⑧↑
                                'V2.0.0.0⑩↓
                            Case 19 ' 再測定回数
                                m_CtlRes(i).Text = (.intCircuitNo).ToString(DirectCast(m_CtlRes(i), cTxt_).GetStrFormat())
                                'V2.0.0.0⑩↑
                                'V2.0.0.0②↓
                            Case 20 ' ON機器
                                Call CreateOnOffExtEquList(i, .intOnExtEqu(1))
                            Case 21 ' ON機器
                                Call CreateOnOffExtEquList(i, .intOnExtEqu(2))
                            Case 22 ' ON機器
                                Call CreateOnOffExtEquList(i, .intOnExtEqu(3))
                            Case 23 ' OFF機器
                                Call CreateOnOffExtEquList(i, .intOffExtEqu(1))
                            Case 24 ' OFF機器
                                Call CreateOnOffExtEquList(i, .intOffExtEqu(2))
                            Case 25 ' OFF機器
                                Call CreateOnOffExtEquList(i, .intOffExtEqu(3))
                                'V2.0.0.0②↑
                            Case Else
                                Throw New Exception("Case " & i & ": Nothing")
                        End Select
                    Next i
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

        ''' <summary>ON機器･OFF機器ｺﾝﾎﾞﾎﾞｯｸｽのﾘｽﾄを作成し、項目を選択する</summary>
        ''' <param name="i">Case 番号</param>
        ''' <param name="OnOffExtEqu">対応する値への参照</param>
        ''' <remarks>GP-IB登録機器名(外部電源のみ)をﾘｽﾄに表示する</remarks>
        Private Sub CreateOnOffExtEquList(ByVal i As Integer, ByRef OnOffExtEqu As Short)
            Dim idx As Integer = 0 ' 選択するｺﾝﾎﾞﾎﾞｯｸｽｲﾝﾃﾞｯｸｽ
            Dim cCombo As cCmb_ = DirectCast(m_CtlRes(i), cCmb_)
            cCombo.Items.Clear()
            cCombo.Items.Add(" 0:なし")
            With m_MainEdit
                If (0 < .W_PLT.GCount) Then ' GP-IB機器が登録されている場合
                    Dim gpibNo As Integer = Convert.ToInt32(OnOffExtEqu) ' 変数に設定されているGP-IB登録番号
                    Dim cnt As Integer = 0 ' ﾘｽﾄに追加した項目数

                    For j As Integer = 1 To (.W_PLT.GCount) Step 1
                        ' ONｺﾏﾝﾄﾞとOFFコマンドが設定されている場合、外部電源としてﾘｽﾄに追加
                        If ("" <> .W_GPIB(j).strCON) AndAlso ("" <> .W_GPIB(j).strCOFF) Then
                            If (Not .W_GPIB(j).strGNAM Is Nothing) Then
                                cCombo.Items.Add(String.Format("{0,2:#0}", j) & ":" & .W_GPIB(j).strGNAM)
                            Else
                                cCombo.Items.Add(String.Format("{0,2:#0}", j) & ":")
                            End If
                            ' 追加したﾘｽﾄをｶｳﾝﾄｱｯﾌﾟ
                            cnt = (cnt + 1)
                            ' OnOffExtEqu(GP-IB登録番号)と同じ項目がﾘｽﾄに追加された場合に
                            ' その項目を選択するためｲﾝﾃﾞｯｸｽを設定する
                            ' 使用中の機器が削除された場合、GP-IBﾀﾌﾞ内の処理(ResetResCutData)で
                            ' OnOffExtEquが0となるため 0:なし が選択される
                            If (gpibNo = j) Then idx = cnt
                        End If
                    Next j

                Else ' GP-IB機器の登録がない場合
                    OnOffExtEqu = 0
                    idx = 0 ' 0:なし
                End If
            End With
            Call NoEventIndexChange(cCombo, idx)

        End Sub
#End Region

#Region "ﾌﾟﾛｰﾌﾞｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内の設定"
        ''' <summary>ﾌﾟﾛｰﾌﾞｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽに値を設定する</summary>
        Private Sub SetProbeData()
            Try
                With m_MainEdit.W_REG(m_ResNo)
                    For i As Integer = 0 To (m_CtlProbe.Length - 1) Step 1
                        Select Case (i)
                            Case 0 ' LO側番号
                                m_CtlProbe(i).Text = (.intPRL).ToString("#0")
                            Case 1 ' HI側番号
                                m_CtlProbe(i).Text = (.intPRH).ToString("#0")
                            Case 2 ' ｶﾞｰﾄﾞ番号
                                m_CtlProbe(i).Text = (.intPRG).ToString("##0")
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

#Region "ｶｯﾄ位置補正ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内の設定"
        ''' <summary>ｶｯﾄ位置補正ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽに値を設定する</summary>
        Private Sub SetCutPosData()
            Try
                With m_MainEdit.W_PTN(m_ResNo)
                    For i As Integer = 0 To (m_CtlCutCorr.Length - 1) Step 1
                        Select Case (i)
                            Case 0 ' 補正実行
                                If (SLP_VMES = m_MainEdit.W_REG(m_ResNo).intSLP) OrElse (SLP_RMES = m_MainEdit.W_REG(m_ResNo).intSLP) Then
                                    ' ｽﾛｰﾌﾟが 5:電圧測定のみ, 6:抵抗測定のみ の場合
                                    .PtnFlg = 0                         ' 補正実行無し
                                    Call NoEventIndexChange(CCmb_5, 0)  ' 補正実行ｺﾝﾎﾞﾎﾞｯｸｽ
                                    Call ChangedCorrection(.PtnFlg)     ' 関連ｺﾝﾄﾛｰﾙの有効･無効を変更
                                    Dim cnt As Integer = 0
                                    For j As Integer = 1 To m_MainEdit.W_PLT.RCount Step 1  ' 抵抗数分
                                        ' 補正実行ありの場合にｶｳﾝﾄｱｯﾌﾟ
                                        If (1 <= m_MainEdit.W_PTN(j).PtnFlg) Then cnt = (cnt + 1)
                                    Next j
                                    '                                    m_MainEdit.W_PLT.PtnCount = Convert.ToInt16(cnt) ' ﾊﾟﾀｰﾝ登録数を設定
                                    m_MainEdit.W_PLT.PtnCount = m_MainEdit.W_PLT.RCount ' ﾊﾟﾀｰﾝ登録数を設定
                                    CGrp_2.Enabled = False              ' ｶｯﾄ位置補正ｸﾞﾙｰﾌﾟﾎﾞｯｸｽを無効にする
                                    Exit For                            ' 以降の設定はおこなわない

                                Else
                                    Call NoEventIndexChange(DirectCast(m_CtlCutCorr(i), cCmb_), .PtnFlg)
                                    Call ChangedCorrection(.PtnFlg)     ' 関連ｺﾝﾄﾛｰﾙの有効･無効を変更
                                    CGrp_2.Enabled = True               ' ｶｯﾄ位置補正ｸﾞﾙｰﾌﾟﾎﾞｯｸｽを有効にする
                                End If

                            Case 1 ' ｸﾞﾙｰﾌﾟ番号(1-999)
                                Call NoEventIndexChange(DirectCast(m_CtlCutCorr(i), cCmb_), (.intGRP - 1))
                            Case 2 ' ﾊﾟﾀｰﾝ番号(1-50)
                                Call NoEventIndexChange(DirectCast(m_CtlCutCorr(i), cCmb_), (.intPTN - 1))
                            Case 3 ' ﾊﾟﾀｰﾝ位置X
                                m_CtlCutCorr(i).Text = (.dblPosX).ToString(DirectCast(m_CtlCutCorr(i), cTxt_).GetStrFormat())
                            Case 4 ' ﾊﾟﾀｰﾝ位置Y
                                m_CtlCutCorr(i).Text = (.dblPosY).ToString(DirectCast(m_CtlCutCorr(i), cTxt_).GetStrFormat())
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

                    ' TODO: 抵抗数が0になることはない仕様のため不要と思われる
                    If (.W_PLT.RCount < 1) Then ' 抵抗数 < 1 ?
                        Dim strMsg As String
                        strMsg = "抵抗データがありません。抵抗データを登録してください。"
                        Call MsgBox(strMsg, MsgBoxStyle.OkOnly, My.Application.Info.Title)
                        ret = 1
                        Exit Try
                    End If

                    For rn As Integer = 1 To .W_PLT.RCount Step 1
                        m_ResNo = rn
                        ' ﾁｪｯｸする抵抗番号のﾃﾞｰﾀをｺﾝﾄﾛｰﾙにｾｯﾄする
                        Call SetDataToText()

                        ' 抵抗ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                        ret = CheckControlData(m_CtlRes)
                        If (ret <> 0) Then Exit Try

                        ' ﾌﾟﾛｰﾌﾞｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                        ret = CheckControlData(m_CtlProbe)
                        If (ret <> 0) Then Exit Try

                        ' ｶｯﾄ位置補正ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                        ret = CheckControlData(m_CtlCutCorr)
                        If (ret <> 0) Then Exit Try

                        ' 相関ﾁｪｯｸ
                        ret = CheckRelation()
                        If (ret <> 0) Then Exit Try
                    Next rn
                End With

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
                ' 抵抗ﾃﾞｰﾀ登録数ﾁｪｯｸ
                ' TODO: 抵抗数が0になることはない仕様のため不要と思われる
                If (m_ResNo < 1) Then
                    Dim strMSG As String
                    strMSG = "抵抗データがありません。" & _
                                            "追加ボタンを押下して抵抗データを登録してください。"
                    Call MsgBox(strMSG, MsgBoxStyle.OkOnly, My.Application.Info.Title)
                    ret = 1
                    Exit Try
                End If

                tag = DirectCast(cTextBox.Tag, Integer)
                With m_MainEdit.W_REG(m_ResNo)
                    Select Case (DirectCast(cTextBox.Parent.Tag, Integer))
                        ' ------------------------------------------------------------------------------
                        Case 0 ' 抵抗ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            Select Case (tag)
                                Case 0 ' ｶｯﾄ数
                                    ' ｽﾛｰﾌﾟが 5:電圧測定のみ, 6:抵抗測定のみ ではない場合のみﾁｪｯｸをおこなう
                                    If UserModule.IsCutResistorIncMarking(m_MainEdit.W_REG, m_ResNo) Then
                                        Dim cnt As Integer = .intTNN ' 変更前の値を保持
                                        ret = CheckShortData(cTextBox, .intTNN)
                                        If (cnt <> .intTNN) Then
                                            If (cnt < .intTNN) Then ' 追加された場合
                                                For i As Integer = (cnt + 1) To .intTNN Step 1
                                                    Call InitCutData(m_ResNo, i) ' 追加されたﾃﾞｰﾀを初期化
                                                Next i
                                            Else ' 削除された場合
                                                For i As Integer = (.intTNN + 1) To cnt Step 1
                                                    Call InitCutData(m_ResNo, i) ' 削除されたﾃﾞｰﾀを初期化
                                                Next i
                                            End If
                                            m_CutNo = 1 ' 処理中のｶｯﾄ番号
                                        End If
                                    End If

                                Case 1 ' 抵抗名
                                    ret = CheckStrData(cTextBox, .strRNO)
                                Case 2 ' ﾘﾚｰﾋﾞｯﾄ
                                    ret = CheckHexData(cTextBox, .lngRel)
                                Case 3 ' 目標値
                                    If Not UserModule.IsMarking(m_MainEdit.W_REG, m_ResNo) Then
                                        ret = CheckDoubleData(cTextBox, .dblNOM)
                                        If ret = 0 Then
                                            ' 下限値/上限値を変更
                                            Call ChangedMode(.intMode, .intSLP, .dblNOM)
                                        End If
                                    End If
                                Case 4 ' 表示単位
                                    ret = CheckStrData(cTextBox, .strTANI)
                                Case 5 ' IT 下限値
                                    ret = CheckDoubleData(cTextBox, .dblITL)
                                Case 6 ' IT 上限値
                                    ret = CheckDoubleData(cTextBox, .dblITH)
                                Case 7 ' FT 下限値
                                    ret = CheckDoubleData(cTextBox, .dblFTL)
                                Case 8 ' FT 上限値
                                    ret = CheckDoubleData(cTextBox, .dblFTH)
                                    'V2.0.0.0↓
                                Case 9 ' 再測定回数
                                    ret = CheckShortData(cTextBox, .intReMeas)
                                Case 10 ' 再測定までのﾎﾟｰｽﾞ時間(ms)
                                    ret = CheckShortData(cTextBox, .intReMeas_Time)
                                    'V2.0.0.0↑
                                    'V2.0.0.0⑧↓
                                Case 11  ' IT測定回数
                                    If Not UserModule.IsMarking(m_MainEdit.W_REG, m_ResNo) Then ' マーキングはチェック無し
                                        ret = CheckShortData(cTextBox, .intITReMeas)
                                    End If
                                Case 12 ' IFT測定回数
                                    If Not UserModule.IsMarking(m_MainEdit.W_REG, m_ResNo) Then ' マーキングはチェック無し
                                        ret = CheckShortData(cTextBox, .intFTReMeas)
                                    End If
                                    'V2.0.0.0⑧↑
                                    'V2.0.0.0⑩↓
                                Case 13 ' サーキット番号
                                    If Not UserModule.IsMarking(m_MainEdit.W_REG, m_ResNo) Then ' マーキングはチェック無し
                                        ret = CheckShortData(cTextBox, .intCircuitNo)
                                    End If
                                    'V2.0.0.0⑩↑
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case 1 ' ﾌﾟﾛｰﾌﾞｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            If Not UserModule.IsMarking(m_MainEdit.W_REG, m_ResNo) Then ' マーキングはチェック無し
                            Select Case (tag)
                                Case 0 ' LO側番号
                                    ret = CheckShortData(cTextBox, .intPRL)
                                Case 1 ' HI側番号
                                    ret = CheckShortData(cTextBox, .intPRH)
                                Case 2 ' ｶﾞｰﾄﾞ番号
                                    ret = CheckShortData(cTextBox, .intPRG)
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            End If
                            ' ------------------------------------------------------------------------------
                        Case 2 ' ｶｯﾄ位置補正ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            With m_MainEdit.W_PTN(m_ResNo)
                                If (1 <= .PtnFlg) Then ' 補正実行ありなら確認をおこなう
                                    Select Case (tag)
                                        Case 0 ' ﾊﾟﾀｰﾝ位置X
                                            ret = CheckDoubleData(cTextBox, .dblPosX)
                                        Case 1 ' ﾊﾟﾀｰﾝ位置Y
                                            ret = CheckDoubleData(cTextBox, .dblPosY)
                                        Case Else
                                            Throw New Exception("Case " & tag & ": Nothing")
                                    End Select
                                End If
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
            Dim ctlArray() As Control ' ｸﾞﾙｰﾌﾟﾎﾞｯｸｽごとのｺﾝﾄﾛｰﾙ配列を参照
            Dim dMin As Double
            Dim dMax As Double

            CheckRelation = 0 ' Return値 = 正常
            Try
                With m_MainEdit
                    ' ------------------------------------------------------------------------------
                    ctlArray = m_CtlRes ' 抵抗ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                    ' 測定種別(0=内部測定, 1<=外部測定)
                    If (.W_PLT.GCount <= 0) AndAlso (1 <= .W_REG(m_ResNo).intMType) Then
                        strMsg = "相関チェックエラー" & vbCrLf
                        strMsg = strMsg & "外部機器の指定がない場合は外部測定器の指定はできません。"
                        GoTo STP_ERR
                    End If

                    With .W_REG(m_ResNo)
                        ' IT 下限値 <= 目標値 <= IT 上限値 ?
                        If (0 = .intMode) Then ' 判定ﾓｰﾄﾞ(0:比率(%), 1:数値(絶対値))
                            dMin = .dblNOM + (.dblNOM * .dblITL * 0.01)     ' Lowﾘﾐｯﾄ値  (LOW = (NOM*(100+Lo)/100))
                            dMax = .dblNOM + (.dblNOM * .dblITH * 0.01)     ' Highﾘﾐｯﾄ値 (HIGH= (NOM*(100+Hi)/100))
                        Else
                            dMin = .dblITL
                            dMax = .dblITH
                        End If
                        If (.dblNOM < dMin) OrElse (dMax < .dblNOM) Then
                            errIdx = RES_NOM
                            strMsg = "相関チェックエラー" + vbCrLf
                            strMsg = strMsg + "IT 下限値 <= 目標値 <= IT 上限値となるように指定してください。"
                            GoTo STP_ERR
                        End If

                        ' FT 下限値 <= 目標値 <= FT 上限値 ?
                        If (0 = .intMode) Then ' 判定ﾓｰﾄﾞ(0:比率(%), 1:数値(絶対値))
                            dMin = .dblNOM + (.dblNOM * .dblFTL * 0.01)     ' Lowﾘﾐｯﾄ値  (LOW = (NOM*(100+Lo)/100))
                            dMax = .dblNOM + (.dblNOM * .dblFTH * 0.01)     ' Highﾘﾐｯﾄ値 (HIGH= (NOM*(100+Hi)/100))
                        Else
                            dMin = .dblFTL
                            dMax = .dblFTH
                        End If
                        If (.dblNOM < dMin) OrElse (dMax < .dblNOM) Then
                            errIdx = RES_NOM
                            strMsg = "相関チェックエラー" + vbCrLf
                            strMsg = strMsg + "FT 下限値 <= 目標値 <= FT 上限値となるように指定してください。"
                            GoTo STP_ERR
                        End If
                        ' ------------------------------------------------------------------------------
                        ctlArray = m_CtlProbe ' ﾌﾟﾛｰﾌﾞｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                        ' 同一ﾌﾟﾛｰﾌﾞ番号ﾁｪｯｸ
                        'V2.0.0.0⑮                        If (.intPRL = .intPRH) Then
                        If (.intPRL = .intPRH) And Not UserModule.IsMarking(m_MainEdit.W_REG, m_ResNo) Then 'V2.0.0.0⑮
                            errIdx = PRB_PRH
                            strMsg = "相関チェックエラー" & vbCrLf
                            strMsg = strMsg & "同一プローブ番号の指定はできません。"
                            GoTo STP_ERR
                        End If

                    End With ' .W_REG(m_ResNo)
                End With ' m_MainEdit

                Exit Function
STP_ERR:
                Call MsgBox_CheckErr(DirectCast(ctlArray(errIdx), cTxt_), strMsg)
                CheckRelation = 1 ' Return値 = ｴﾗｰ

            Catch ex As Exception
                Call MsgBox_Exception(ex)
                CheckRelation = 1 ' Return値 = ｴﾗｰ
            End Try

        End Function
#End Region

#Region "ｺﾝﾎﾞﾎﾞｯｸｽ関連処理"
#Region "ｽﾛｰﾌﾟが変更された時に目標値ﾃｷｽﾄﾎﾞｯｸｽの設定を変更する"
        ''' <summary>ｽﾛｰﾌﾟが変更された時に目標値ﾃｷｽﾄﾎﾞｯｸｽの設定を変更する</summary>
        ''' <param name="mode">判定ﾓｰﾄﾞ</param>
        ''' <param name="slp">ｽﾛｰﾌﾟ(1:+電圧, 2:-電圧, 4:抵抗, 5:電圧測定のみ, 6:抵抗測定のみ, 7:NGﾏｰｷﾝｸﾞ)</param>
        ''' <param name="dNOM">目標値</param>
        Private Sub ChangedSlope(ByVal mode As Integer, ByVal slp As Integer, ByVal dNOM As Double)
            Dim strMin As String
            Dim strMax As String
            Try
                If (SLP_VTRIMPLS = slp Or SLP_VTRIMMNS = slp Or SLP_VMES = slp) Then ' (電圧のみ)
                    strMin = m_voltNOM_Min
                    strMax = m_voltNOM_Max
                Else
                    strMin = m_ohmNOM_Min
                    strMax = m_ohmNOM_Max
                End If

                With DirectCast(m_CtlRes(RES_NOM), cTxt_) ' 目標値ﾃｷｽﾄﾎﾞｯｸｽ
                    Call .SetMinMax(strMin, strMax) ' 下限値･上限値の設定
                    Call .SetStrTip(strMin & "～" & strMax & "の範囲で指定して下さい") ' ﾂｰﾙﾁｯﾌﾟﾒｯｾｰｼﾞの設定
                End With

                Call ChangedMode(mode, slp, dNOM)

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#Region "判定ﾓｰﾄﾞが変更された時にIT,FTﾃｷｽﾄﾎﾞｯｸｽの設定を変更する"
        ''' <summary>判定ﾓｰﾄﾞが変更された時にIT,FTﾃｷｽﾄﾎﾞｯｸｽの設定を変更する</summary>
        ''' <param name="mode">判定ﾓｰﾄﾞ(0:比率(%), 1:数値(絶対値)</param>
        ''' <param name="slp">ｽﾛｰﾌﾟ(1:+電圧, 2:-電圧, 4:抵抗, 5:電圧測定のみ, 6:抵抗測定のみ, 7:NGﾏｰｷﾝｸﾞ)</param>
        ''' <param name="dNOM">目標値</param>
        Private Sub ChangedMode(ByVal mode As Integer, ByVal slp As Integer, ByVal dNOM As Double)
            Dim Length As Integer = (m_IntialFinal.Length - 1)
            Dim strMin(Length) As String
            Dim strMax(Length) As String
            Dim i As Integer
            Dim dValue As Double

            Try
                If (JUDGE_MODE_RATIO = mode) Then ' (0:比率(%), 1:数値(絶対値))
                    strMin(0) = m_ITL_Min   ' IT 下限値
                    strMax(0) = m_ITL_Max
                    strMin(1) = m_ITH_Min   ' IT 上限値
                    strMax(1) = m_ITH_Max
                    strMin(2) = m_FTL_Min   ' FT 下限値
                    strMax(2) = m_FTL_Max
                    strMin(3) = m_FTH_Min   ' FT 上限値
                    strMax(3) = m_FTH_Max
                Else

                    For i = 0 To Length Step 1
                        If (slp = SLP_RTRM) Or (slp = SLP_RMES) Then    ' 抵抗の場合
                            If TypeOf Me.m_CtlRes(RES_NOM) Is cTxt_ Then
                                dValue = -(dNOM - Double.Parse(m_ohmNOM_Min))
                                strMin(i) = dValue.ToString(DirectCast(m_CtlRes(RES_NOM), cTxt_).GetStrFormat())
                                dValue = Double.Parse(m_ohmNOM_Max) - dNOM
                                strMax(i) = dValue.ToString(DirectCast(m_CtlRes(RES_NOM), cTxt_).GetStrFormat())
                            End If
                        Else                                            ' 電圧の場合
                            strMin(i) = m_voltNOM_Min
                            strMax(i) = m_voltNOM_Max
                        End If
                    Next i
                End If

                For i = 0 To Length Step 1
                    With m_IntialFinal(i)
                        Call .SetMinMax(strMin(i), strMax(i)) ' 下限値･上限値の設定
                        Call .SetStrTip(strMin(i) & "～" & strMax(i) & "の範囲で指定して下さい") ' ﾂｰﾙﾁｯﾌﾟﾒｯｾｰｼﾞの設定
                    End With
                Next

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#Region "補正実行ｺﾝﾎﾞﾎﾞｯｸｽのｲﾝﾃﾞｯｸｽが変更された時に関連ｺﾝﾄﾛｰﾙの有効･無効を変更する"
        ''' <summary>補正実行ｺﾝﾎﾞﾎﾞｯｸｽのｲﾝﾃﾞｯｸｽが変更された時に関連ｺﾝﾄﾛｰﾙの有効･無効を変更する</summary>
        ''' <param name="idx">0:なし, 1:自動, 2:自動+手動</param>
        Private Sub ChangedCorrection(ByVal idx As Integer)
            Dim tf As Boolean
            'V1.0.4.3⑥            If (0 = idx) OrElse (3 = idx) Then ' 補正なし
            If (CUT_PATTERN_NONE = idx) Then ' 補正なし
                tf = False
            Else ' 補正あり
                tf = True
            End If

            For Each ctl As Control In m_CutPosCorr
                ctl.Enabled = tf
            Next
        End Sub
#End Region
#End Region

#Region "追加･削除ﾎﾞﾀﾝ関連処理"
        ''' <summary>抵抗ﾃﾞｰﾀを追加または削除し、その抵抗ﾃﾞｰﾀを初期化する</summary>
        ''' <param name="addDel">1=Add, (-1)=Del</param>
        Private Sub SortResistorData(ByVal addDel As Integer)
            Dim iStart As Integer
            Dim iEnd As Integer
            Dim dir As Integer = (-1) * addDel ' Add=(-1), Del=1にする
            Try
                With m_MainEdit
                    If (1 = addDel) Then ' 追加の場合
                        .W_PLT.RCount = Convert.ToInt16(.W_PLT.RCount + 1) ' 登録抵抗数を追加する
                        iStart = .W_PLT.RCount ' 登録されている抵抗数から
                        iEnd = (m_ResNo + 1) ' 追加する抵抗ﾃﾞｰﾀ番号+1まで、前のﾃﾞｰﾀを後ろにずらす
                    Else ' 削除の場合
                        iStart = m_ResNo ' 削除する抵抗ﾃﾞｰﾀ番号から
                        iEnd = (.W_PLT.RCount - 1) ' 登録されている抵抗ﾃﾞｰﾀ数-1まで、後ろのﾃﾞｰﾀを前にずらす
                    End If

                    For rn As Integer = iStart To iEnd Step dir
                        With .W_REG(rn)
                            .strRNO = m_MainEdit.W_REG(rn + dir).strRNO     ' 抵抗名
                            .strTANI = m_MainEdit.W_REG(rn + dir).strTANI   ' 表示単位("V","Ω" 等)
                            .intSLP = m_MainEdit.W_REG(rn + dir).intSLP     ' 電圧変化ｽﾛｰﾌﾟ(1:+V, 2:-V, 4:抵抗)
                            .lngRel = m_MainEdit.W_REG(rn + dir).lngRel     ' ﾘﾚｰﾋﾞｯﾄ
                            .dblNOM = m_MainEdit.W_REG(rn + dir).dblNOM     ' ﾄﾘﾐﾝｸﾞ 目標値
                            .dblITL = m_MainEdit.W_REG(rn + dir).dblITL     ' 初期判定下限値 (ITLO)
                            .dblITH = m_MainEdit.W_REG(rn + dir).dblITH     ' 初期判定上限値 (ITHI)
                            .dblFTL = m_MainEdit.W_REG(rn + dir).dblFTL     ' 終了判定下限値 (FTLO)
                            .dblFTH = m_MainEdit.W_REG(rn + dir).dblFTH     ' 終了判定上限値 (FTHI)
                            .intMode = m_MainEdit.W_REG(rn + dir).intMode   ' 判定ﾓｰﾄﾞ(0:比率(%), 1:数値(絶対値))
                            .intMeasMode = m_MainEdit.W_REG(rn + dir).intMeasMode       ' 測定モード(0:なし, 1:ITのみ 2:FTのみ 3:IT,FT両方)
                            .intTMM1 = m_MainEdit.W_REG(rn + dir).intTMM1               ' モード(0:高速(コンパレータ非積分モード), 1:高精度(積分モード))
                            .intPRH = m_MainEdit.W_REG(rn + dir).intPRH     ' HI側ﾌﾟﾛｰﾌﾞ番号
                            .intPRL = m_MainEdit.W_REG(rn + dir).intPRL     ' LO側ﾌﾟﾛｰﾌﾞ番号
                            .intPRG = m_MainEdit.W_REG(rn + dir).intPRG     ' ｶﾞｰﾄﾞﾌﾟﾛｰﾌﾞ番号
                            .intMType = m_MainEdit.W_REG(rn + dir).intMType ' 測定種別(0=内部測定, 1=外部測定)
                            .intTNN = m_MainEdit.W_REG(rn + dir).intTNN     ' ｶｯﾄ数(1～9)
                            'V2.0.0.0↓
                            .intReMeas = m_MainEdit.W_REG(rn + dir).intReMeas           ' 再測定回数
                            .intReMeas_Time = m_MainEdit.W_REG(rn + dir).intReMeas_Time ' ON後のﾎﾟｰｽﾞ時間(ms)
                            For i As Integer = 1 To EXTEQU Step 1
                                .intOnExtEqu(i) = m_MainEdit.W_REG(rn + dir).intOnExtEqu(i)     ' ＯＮ機器１（ＯＮする外部機器１～３）
                                .intOffExtEqu(i) = m_MainEdit.W_REG(rn + dir).intOffExtEqu(i)   ' ＯＦＦ機器１（ＯＦＦする外部機器１～３）
                            Next i
                            'V2.0.0.0↑
                            'V2.0.0.0⑧↓
                            .intITReMeas = m_MainEdit.W_REG(rn + dir).intITReMeas       ' IT測定回数
                            .intFTReMeas = m_MainEdit.W_REG(rn + dir).intFTReMeas   ' FT測定回数
                            'V2.0.0.0⑧↑
                            'V2.0.0.0⑩↓
                            .intCircuitNo = m_MainEdit.W_REG(rn + dir).intCircuitNo     ' 再測定回数
                            'V2.0.0.0⑩↑

                            For cn As Integer = 1 To MAXCTN Step 1
                                With .STCUT(cn)
                                    .intCUT = m_MainEdit.W_REG(rn + dir).STCUT(cn).intCUT     ' ｶｯﾄ方法(1:ﾄﾗｯｷﾝｸﾞ, 2:ｲﾝﾃﾞｯｸｽ, 3:ﾎﾟｼﾞｼｮﾆﾝｸﾞ無しｲﾝﾃﾞｯｸｽ)
                                    .intCTYP = m_MainEdit.W_REG(rn + dir).STCUT(cn).intCTYP   ' ｶｯﾄ形状(1:ｽﾄﾚｰﾄ, 2:Lｶｯﾄ)
                                    .dblSTX = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblSTX     ' ﾄﾘﾐﾝｸﾞ ｽﾀｰﾄ点 X
                                    .dblSTY = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblSTY     ' ﾄﾘﾐﾝｸﾞ ｽﾀｰﾄ点 Y
                                    .dblSX2 = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblSX2     ' ﾄﾘﾐﾝｸﾞ ｽﾀｰﾄ点2 X
                                    .dblSY2 = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblSY2     ' ﾄﾘﾐﾝｸﾞ ｽﾀｰﾄ点2 Y
                                    .dblDL2 = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblDL2     ' 第2のｶｯﾄ長(ﾘﾐｯﾄｶｯﾄ量mm(Lﾀｰﾝ前))
                                    .dblDL3 = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblDL3     ' 第3のｶｯﾄ長(ﾘﾐｯﾄｶｯﾄ量mm(Lﾀｰﾝ後))
                                    .intANG = m_MainEdit.W_REG(rn + dir).STCUT(cn).intANG     ' ｶｯﾄ方向1
                                    .intANG2 = m_MainEdit.W_REG(rn + dir).STCUT(cn).intANG2   ' ｶｯﾄ方向2
                                    .intQF1 = m_MainEdit.W_REG(rn + dir).STCUT(cn).intQF1     ' Qﾚｰﾄ(0.1KHz)
                                    .dblV1 = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblV1       ' ﾄﾘﾑ速度(mm/s)
                                    .dblCOF = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblCOF     ' ｶｯﾄｵﾌ(%)
                                    .dblLTP = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblLTP     ' Lﾀｰﾝ ﾎﾟｲﾝﾄ(%)
                                    .intTMM = m_MainEdit.W_REG(rn + dir).STCUT(cn).intTMM     ' ﾓｰﾄﾞ(0:高速(ｺﾝﾊﾟﾚｰﾀ非積分ﾓｰﾄﾞ), 1:高精度(積分ﾓｰﾄﾞ))
                                    .intMType = m_MainEdit.W_REG(rn + dir).STCUT(cn).intMType ' 内部／外部測定器
                                    .cFormat = m_MainEdit.W_REG(rn + dir).STCUT(cn).cFormat   '###1042① 文字データ
                                    'V2.1.0.0①↓ カット毎の抵抗値変化量判定機能追加
                                    .iVariationRepeat = m_MainEdit.W_REG(rn + dir).STCUT(cn).iVariationRepeat   ' リピート有無
                                    .iVariation = m_MainEdit.W_REG(rn + dir).STCUT(cn).iVariation               ' 判定有無
                                    .dRateOfUp = m_MainEdit.W_REG(rn + dir).STCUT(cn).dRateOfUp                 ' 上昇率
                                    .dVariationLow = m_MainEdit.W_REG(rn + dir).STCUT(cn).dVariationLow         ' 下限値
                                    .dVariationHi = m_MainEdit.W_REG(rn + dir).STCUT(cn).dVariationHi           ' 上限値
                                    'V2.1.0.0①↑

                                    'V2.1.0.0④ ADD ↓
                                    For i As Integer = 1 To MAX_LCUT Step 1         ' MAXｽｶｯﾄ数分繰返す
                                        .dCutLen(i) = m_MainEdit.W_REG(rn + dir).STCUT(cn).dCutLen(i)                   ' カット長
                                        .dQRate(i) = m_MainEdit.W_REG(rn + dir).STCUT(cn).dQRate(i)                     ' Ｑレート
                                        .dSpeed(i) = m_MainEdit.W_REG(rn + dir).STCUT(cn).dSpeed(i)                     ' 速度
                                        .dAngle(i) = m_MainEdit.W_REG(rn + dir).STCUT(cn).dAngle(i)                     ' 方向（角度）
                                        .dTurnPoint(i) = m_MainEdit.W_REG(rn + dir).STCUT(cn).dTurnPoint(i)             ' Ｌターンポイント
                                    Next i
                                    .intRetraceCnt = m_MainEdit.W_REG(rn + dir).STCUT(cn).intRetraceCnt                       ' リトレースカット本数
                                    For i As Integer = 1 To MAX_RETRACECUT Step 1           ' MAXｽｶｯﾄ数分繰返す
                                        .dblRetraceOffX(i) = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblRetraceOffX(i)     ' リトレースのオフセットＸ
                                        .dblRetraceOffY(i) = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblRetraceOffY(i)     ' リトレースのオフセットＹ
                                        .dblRetraceQrate(i) = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblRetraceQrate(i)   ' ストレートカット・リトレースのQレート(0.1KHz)に使用
                                        .dblRetraceSpeed(i) = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblRetraceSpeed(i)   ' ストレートカット・リトレースのトリム速度(mm/s)に使用
                                    Next i
                                    'V2.1.0.0④ ADD ↑

                                    ' ｲﾝﾃﾞｯｸｽｶｯﾄ情報設定
                                    For ix As Integer = 1 To MAXIDX Step 1 ' MAXｲﾝﾃﾞｯｸｽｶｯﾄ数分繰返す
                                        .intIXN(ix) = m_MainEdit.W_REG(rn + dir).STCUT(cn).intIXN(ix) ' ｲﾝﾃﾞｯｸｽｶｯﾄ数1-5
                                        .dblDL1(ix) = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblDL1(ix) ' ｶｯﾄ長1-5
                                        .lngPAU(ix) = m_MainEdit.W_REG(rn + dir).STCUT(cn).lngPAU(ix) ' ﾋﾟｯﾁ間ﾎﾟｰｽﾞ時間1-5
                                        .dblDEV(ix) = m_MainEdit.W_REG(rn + dir).STCUT(cn).dblDEV(ix) ' 誤差1-5(%)
                                        .intIXMType(ix) = m_MainEdit.W_REG(rn + dir).STCUT(cn).intIXMType(ix) ' 測定機器
                                        .intIXTMM(ix) = m_MainEdit.W_REG(rn + dir).STCUT(cn).intIXTMM(ix) ' 測定ﾓｰﾄﾞ
                                    Next ix

                                    ' FL加工条件
                                    For fl As Integer = 1 To MAXCND Step 1
                                        .intCND(fl) = m_MainEdit.W_REG(rn + dir).STCUT(cn).intCND(fl) ' FL設定No.
                                    Next fl
                                End With ' .STCUT(cn)
                            Next cn
                        End With ' .W_REG(rn)

                        ' ｶｯﾄ位置補正
                        With .W_PTN(rn)
                            .PtnFlg = m_MainEdit.W_PTN(rn + dir).PtnFlg     ' 補正実行(0:なし, 1:自動, 2:自動+手動)
                            .intGRP = m_MainEdit.W_PTN(rn + dir).intGRP     ' ｸﾞﾙｰﾌﾟ番号
                            .intPTN = m_MainEdit.W_PTN(rn + dir).intPTN     ' ﾊﾟﾀｰﾝ番号
                            .dblPosX = m_MainEdit.W_PTN(rn + dir).dblPosX   ' ﾊﾟﾀｰﾝX
                            .dblPosY = m_MainEdit.W_PTN(rn + dir).dblPosY   ' ﾊﾟﾀｰﾝY
                            .dblDRX = m_MainEdit.W_PTN(rn + dir).dblDRX     ' ずれ量保存ﾜｰｸX
                            .dblDRY = m_MainEdit.W_PTN(rn + dir).dblDRY     ' ずれ量保存ﾜｰｸY
                        End With
                    Next rn

                    ' つめて不要となったﾃﾞｰﾀを初期化する
                    If (1 = addDel) Then ' 追加の場合
                        Call InitResData(m_ResNo)       ' 追加した抵抗ﾃﾞｰﾀを初期化
                    Else ' 削除の場合
                        Dim lastRn As Integer = Convert.ToInt32(.W_PLT.RCount)
                        Call InitResData(lastRn)        ' 最後のﾃﾞｰﾀを初期化
                        .W_PLT.RCount = Convert.ToInt16(lastRn - 1) ' 登録抵抗数を-1する

                        ' 最終抵抗の削除なら現在の抵抗番号を最終抵抗番号とする
                        If (.W_PLT.RCount < m_ResNo) Then m_ResNo = .W_PLT.RCount
                    End If
                End With

                ' 抵抗ﾃﾞｰﾀを画面項目に設定
                Call SetDataToText()
                FIRST_CONTROL.Select() ' ﾌｵｰｶｽ設定

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
                        Case 0 ' 抵抗ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            Select Case (tag)
                                Case 0 ' 抵抗番号
                                    m_ResNo = (idx + 1)
                                    ' 対応するﾃﾞｰﾀをﾃｷｽﾄﾎﾞｯｸｽにｾｯﾄする
                                    Call SetDataToText()

                                    If UserModule.IsMarking(m_MainEdit.W_REG, m_ResNo) Then
                                        m_MainEdit.W_REG(m_ResNo).intMeasMode = MEAS_JUDGE_NONE
                                        ChangeSlopeAllOnOff(False)
                                    Else
                                        ChangeSlopeAllOnOff(True)

                                        If m_MainEdit.W_REG(m_ResNo).intMType > 0 Then
                                            m_CtlRes(RES_TMM1).Enabled = False
                                        Else
                                            m_CtlRes(RES_TMM1).Enabled = True
                                        End If
                                    End If

                                Case 1 ' ｽﾛｰﾌﾟ
                                    Dim iSlp As Integer

                                    iSlp = GetComboBoxName2Value(cCombo.Text, Me.m_lstSlope)

                                    With .W_REG(m_ResNo)
                                        .intSLP = Convert.ToInt16(iSlp)         ' ｽﾛｰﾌﾟを設定

                                        If UserModule.IsMarking(m_MainEdit.W_REG, m_ResNo) Then
                                            m_MainEdit.W_REG(m_ResNo).intMeasMode = MEAS_JUDGE_NONE
                                            'Call NoEventIndexChange(CCmb_16, 0) ' 補正実行ｺﾝﾎﾞﾎﾞｯｸｽ
                                            ChangeSlopeAllOnOff(False)
                                        Else
                                            'm_MainEdit.W_REG(m_ResNo).intMeasMode = MEAS_JUDGE_FT
                                            'CCmb_16.Enabled = True
                                            ChangeSlopeAllOnOff(True)
                                        End If

                                        Call ChangedSlope(.intMode, iSlp, .dblNOM)

                                        ' ------------------------
                                        'If (SLP_VMES <= iSlp And iSlp <= SLP_RMES) Then
                                        If UserModule.IsMeasureOnly(m_MainEdit.W_REG, m_ResNo) Then
                                            ' ｽﾛｰﾌﾟが 7:電圧測定のみ, 9:抵抗測定のみ の場合
                                            CTxt_0.Text = 0                     ' ｶｯﾄ数の表示を0にする
                                            CTxt_0.Enabled = False              ' 無効にする
                                            If (1 < .intTNN) Then               ' もとのｶｯﾄ数
                                                For i As Integer = 2 To .intTNN Step 1
                                                    Call InitCutData(m_ResNo, i) ' 2以降のｶｯﾄﾃﾞｰﾀを初期化
                                                Next i
                                                .intTNN = 1                     ' ｶｯﾄ数を1にする
                                                m_CutNo = 1                     ' 処理中のｶｯﾄ番号
                                            End If

                                            With m_MainEdit.W_PTN(m_ResNo)
                                                .PtnFlg = PTN_NONE                 ' 補正実行無し
                                                Call NoEventIndexChange(CCmb_5, 0) ' 補正実行ｺﾝﾎﾞﾎﾞｯｸｽ
                                                Call ChangedCorrection(.PtnFlg) ' 関連ｺﾝﾄﾛｰﾙの有効･無効を変更
                                                Dim cnt As Integer = 0
                                                m_MainEdit.W_PLT.PtnCount = m_MainEdit.W_PLT.RCount

                                            End With
                                            CGrp_2.Enabled = False              ' ｶｯﾄ位置補正ｸﾞﾙｰﾌﾟﾎﾞｯｸｽを無効にする

                                        Else
                                            CTxt_0.Text = (.intTNN).ToString()  ' ｶｯﾄ数
                                            CTxt_0.Enabled = True               ' 有効にする

                                            CGrp_2.Enabled = True               ' ｶｯﾄ位置補正ｸﾞﾙｰﾌﾟﾎﾞｯｸｽを有効にする
                                        End If

                                        ' ------------------------
                                    End With

                                Case 2 ' 判定ﾓｰﾄﾞ
                                    ' (0:比率(%), 1:数値(絶対値))
                                    .W_REG(m_ResNo).intMode = Convert.ToInt16(idx)
                                    ' 関連するｺﾝﾄﾛｰﾙの最大･最小値とﾂｰﾙﾁｯﾌﾟなどの設定を変更する
                                    Call ChangedMode(.W_REG(m_ResNo).intMode, .W_REG(m_ResNo).intSLP, .W_REG(m_ResNo).dblNOM)

                                Case 3  ' 測定モード
                                    .W_REG(m_ResNo).intMeasMode = GetComboBoxName2Value(cCombo.Text, Me.m_lstMeasMode)

                                Case 4 ' 測定機器(0:内部測定器, 1以上は外部測定器)
                                    ' 登録されている測定機器ﾘｽﾄの数値を設定する( 1:NAME=1, 10:NAME=10)
                                    .W_REG(m_ResNo).intMType = Short.Parse((cCombo.Text).Substring(0, 2))
                                    ' 外部測定器の場合測定ﾓｰﾄﾞを無効にする
                                    If (0 < idx) Then
                                        m_CtlRes(RES_TMM1).Enabled = False
                                    Else
                                        m_CtlRes(RES_TMM1).Enabled = True
                                    End If
                                Case 5 ' 測定ﾓｰﾄﾞ
                                    .W_REG(m_ResNo).intTMM1 = Convert.ToInt16(idx)
                                    'V2.0.0.0②
                                Case 6 ' ON機器
                                    ' 登録されている測定機器ﾘｽﾄの数値を設定する( 1:NAME=1, 10:NAME=10)
                                    .W_REG(m_ResNo).intOnExtEqu(1) = Short.Parse((cCombo.Text).Substring(0, 2))
                                Case 7 ' ON機器
                                    .W_REG(m_ResNo).intOnExtEqu(2) = Short.Parse((cCombo.Text).Substring(0, 2))
                                Case 8 ' ON機器
                                    .W_REG(m_ResNo).intOnExtEqu(3) = Short.Parse((cCombo.Text).Substring(0, 2))
                                Case 9 ' OFF機器
                                    ' 登録されている測定機器ﾘｽﾄの数値を設定する( 1:NAME=1, 10:NAME=10)
                                    .W_REG(m_ResNo).intOffExtEqu(1) = Short.Parse((cCombo.Text).Substring(0, 2))
                                Case 10 ' OFF機器
                                    .W_REG(m_ResNo).intOffExtEqu(2) = Short.Parse((cCombo.Text).Substring(0, 2))
                                Case 11 ' OFF機器
                                    .W_REG(m_ResNo).intOffExtEqu(3) = Short.Parse((cCombo.Text).Substring(0, 2))
                                    'V2.0.0.0②
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case 1 ' ﾌﾟﾛｰﾌﾞｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            Throw New Exception("Parent.Tag - Case 1")
                            ' ------------------------------------------------------------------------------
                        Case 2 ' ｶｯﾄ位置補正ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            With .W_PTN(m_ResNo)
                                Select Case (tag)
                                    Case 0 ' 補正実行(0:なし, 1:自動, 2:自動+手動)
                                        .PtnFlg = Convert.ToInt16(idx)
                                        Call ChangedCorrection(idx) ' 関連ｺﾝﾄﾛｰﾙの有効･無効を変更
                                        Dim cnt As Integer = 0
                                        For i As Integer = 1 To m_MainEdit.W_PLT.RCount Step 1
                                            If (1 <= m_MainEdit.W_PTN(i).PtnFlg) Then cnt = (cnt + 1) ' 補正実行ありの場合にｶｳﾝﾄｱｯﾌﾟ
                                        Next i
                                        'm_MainEdit.W_PLT.PtnCount = Convert.ToInt16(cnt) ' ﾊﾟﾀｰﾝ登録数を設定
                                        m_MainEdit.W_PLT.PtnCount = m_MainEdit.W_PLT.RCount ' ﾊﾟﾀｰﾝ登録数を設定

                                    Case 1 ' ｸﾞﾙｰﾌﾟ番号(1-999)
                                        .intGRP = Convert.ToInt16(idx + 1)
                                    Case 2 ' ﾊﾟﾀｰﾝ番号(1-50)
                                        .intPTN = Convert.ToInt16(idx + 1)
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
                Call MsgBox_Exception(ex)
            End Try

        End Sub

        ''' <summary>追加ﾎﾞﾀﾝｸﾘｯｸ時の処理</summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Private Sub CBtn_Add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBtn_Add.Click
            Dim strMsg As String ' ﾒｯｾｰｼﾞﾎﾞｯｸｽのｷｬﾌﾟｼｮﾝ表示用
            Dim refOpt As Short ' ｵﾌﾟｼｮﾝﾎﾞﾀﾝ(0=前に追加 ,1=後に追加)
            Dim ret As Integer
            Try
                ' 登録数ﾁｪｯｸ
                If (MAXRNO <= m_MainEdit.W_PLT.RCount) Then ' 登録数OK ?
                    strMsg = "これ以上抵抗データは登録できません。"
                    Call MsgBox(strMsg, DirectCast( _
                                MsgBoxStyle.OkOnly + _
                                MsgBoxStyle.Information, MsgBoxStyle), _
                                My.Application.Info.Title)
                    Exit Sub
                End If

                ' 確認ﾒｯｾｰｼﾞを表示("抵抗データを追加します")
                ret = MsgBox_AddClick("抵抗データ", refOpt) ' ﾒｯｾｰｼﾞ表示
                If (ret <> cFRS_ERR_ADV) Then Exit Sub ' CancelならReturn
                If (refOpt = 1) Then ' 表示ﾃﾞｰﾀの後に追加 ?
                    m_ResNo = (m_ResNo + 1) ' m_ResNo = 現在のﾃﾞｰﾀ番号 + 1
                Else ' 表示ﾃﾞｰﾀの前に追加
                    m_ResNo = m_ResNo ' m_ResNo = 現在のﾃﾞｰﾀ番号
                End If

                ' ﾃﾞｰﾀを1個後にずらす
                Call SortResistorData(1)

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
                If (1 = m_MainEdit.W_PLT.RCount) Then Exit Sub ' 登録数1ならNOP
                strMsg = "現在の抵抗データを削除します。よろしいですか？"
                ret = MsgBox(strMsg, DirectCast( _
                            MsgBoxStyle.OkCancel + _
                            MsgBoxStyle.Exclamation + _
                            MsgBoxStyle.DefaultButton2, MsgBoxStyle), _
                            My.Application.Info.Title)
                If (ret = MsgBoxResult.Cancel) Then Exit Sub ' Cancel(RESETｷｰ) ?

                ' ﾃﾞｰﾀを1個前につめる
                Call SortResistorData(-1)

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

        'V2.0.0.0↓
#Region "スロープデータ初期設定"
        ''' <summary>
        ''' スロープデータ初期設定
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub InitSlopeData()
            Dim data As New ComboDataStruct
#If VOLTAGE_USE Then
            data.SetData("＋電圧トリミング", SLP_VTRIMPLS)
            Me.m_lstSlope.Add(data)
            data.SetData("－電圧トリミング", SLP_VTRIMMNS)
            Me.m_lstSlope.Add(data)
#End If
            data.SetData("抵抗トリミング", SLP_RTRM)
            Me.m_lstSlope.Add(data)
#If VOLTAGE_USE Then
            data.SetData("電圧測定のみ", SLP_VMES)
            Me.m_lstSlope.Add(data)
#End If
            data.SetData("抵抗測定のみ", SLP_RMES)
            Me.m_lstSlope.Add(data)

#If NG_MARKING_USE Then
            data.SetData("ＮＧマーキング", SLP_NG_MARK)
            Me.m_lstSlope.Add(data)
#End If

#If OK_MARKING_USE Then
            data.SetData("ＯＫマーキング", SLP_OK_MARK)
            Me.m_lstSlope.Add(data)
#End If

#If OK_NG_MARKING_USE Then
            data.SetData("ＮＧマーキング", SLP_NG_MARK)
            Me.m_lstSlope.Add(data)
            data.SetData("ＯＫマーキング", SLP_OK_MARK)
            Me.m_lstSlope.Add(data)
#End If
            'V2.2.1.7① ↓
            If (m_MainEdit.W_stUserData.iTrimType = 5) Then
                data.SetData("マーク印字", SLP_MARK)
                Me.m_lstSlope.Add(data)
            End If

            'data.SetData("マーク印字", SLP_MARK)
            'Me.m_lstSlope.Add(data)
            'V2.2.1.7① ↑

        End Sub
#End Region

#Region "測定モードデータ初期設定"
        ''' <summary>
        ''' 測定モードデータ初期設定
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub InitMeasModeData()
            Dim data As New ComboDataStruct

            data.SetData("なし", MEAS_JUDGE_NONE)
            Me.m_lstMeasMode.Add(data)
            data.SetData("ITのみ", MEAS_JUDGE_IT)
            Me.m_lstMeasMode.Add(data)
            data.SetData("FTのみ", MEAS_JUDGE_FT)
            Me.m_lstMeasMode.Add(data)
            data.SetData("IT,FT両方", MEAS_JUDGE_BOTH)
            Me.m_lstMeasMode.Add(data)
        End Sub
#End Region

#Region "スロープによる入力不可可能の変更"
        ''' <summary>
        ''' ＮＧマーキングの時他の項目は全てＯＦＦする。
        ''' </summary>
        ''' <param name="OnOff"></param>
        ''' <remarks></remarks>
        Private Sub ChangeSlopeAllOnOff(ByVal OnOff As Boolean)
            Try
                If OnOff Then
                    ' 抵抗データ
                    For i As Integer = 4 To (m_CtlRes.Length - 1) Step 1
                        m_CtlRes(i).Enabled = True
                    Next i
                    ' プローブ
                    For i As Integer = 0 To (m_CtlProbe.Length - 1) Step 1
                        m_CtlProbe(i).Enabled = True
                    Next i

#If ADDITIONAL_GPIB Then        ' 追加GPIB機器
                    If m_MainEdit.W_USER.iTrimType = 2 Or m_MainEdit.W_USER.iTrimType = 3 Then
                        CCmb_GPIB.Enabled = True
                        CTxt_GPIB_LO.Enabled = True
                        CTxt_GPIB_HI.Enabled = True
                    Else
                        CCmb_GPIB.Enabled = False
                        CTxt_GPIB_LO.Enabled = False
                        CTxt_GPIB_HI.Enabled = False
                    End If
#End If

                Else
                    ' 抵抗データ
                    For i As Integer = 4 To (m_CtlRes.Length - 1) Step 1
                        m_CtlRes(i).Enabled = False
                    Next i
                    ' プローブ
                    For i As Integer = 0 To (m_CtlProbe.Length - 1) Step 1
                        m_CtlProbe(i).Enabled = False
                    Next i
                End If

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try
        End Sub
#End Region

#Region "スロープリストの変更"
        Public Sub ChangeSlopeList()
            Try
                Dim cCombo As cCmb_ = DirectCast(m_CtlRes(RES_SLOPE), cCmb_)
                m_lstSlope = Nothing
                m_lstSlope = New List(Of ComboDataStruct)
                InitSlopeData()
                With cCombo
                    .Items.Clear()
                    For i As Integer = 0 To Me.m_lstSlope.Count - 1 Step 1
                        .Items.Add(Me.m_lstSlope(i).Name)
                    Next i
                End With
            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try
        End Sub
#End Region
        'V2.0.0.0↑

    End Class
End Namespace
