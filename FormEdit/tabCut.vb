Option Explicit On
Option Strict Off

Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports LaserFront.Trimmer.DefWin32Fnc

Namespace FormEdit
    Friend Class tabCut
        Inherits tabBase

#Region "ｸﾗｽ"
#Region "NGｶｯﾄ用"
        ''' <summary>NGｶｯﾄ関連のｺﾝﾄﾛｰﾙの有効･無効を切替える</summary>
        Private Class NGCut
            Friend m_CtlArr() As Control    ' NGｶｯﾄで使用するｺﾝﾄﾛｰﾙ
            Friend WriteOnly Property Enabled() As Boolean
                Set(ByVal value As Boolean)
                    For Each ctl As Control In m_CtlArr
                        ctl.Enabled = value
                    Next
                End Set
            End Property
        End Class
#End Region

#Region "ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ用"
        ''' <summary>ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ関連の表示･非表示をおこなう</summary>
        Private Class Serpentine
            Friend m_ctlArr() As Control    ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄで使用するｺﾝﾄﾛｰﾙ
            Friend m_lblArr() As Label      ' ｶｯﾄ長1/2,ｶｯﾄ方向1/2のﾗﾍﾞﾙ
            ' { {ｶｯﾄ長       }, {ｶｯﾄ方向(0～360°)   } }
            ' { {ｶｯﾄ長1/2(mm)}, {ｶｯﾄ方向1/2(0～360°)} }
            Friend m_strLbl(1, 1) As String ' ﾗﾍﾞﾙに表示する文字列

            Friend WriteOnly Property Visible() As Boolean
                Set(ByVal value As Boolean)
                    Dim tf As Integer = 0
                    If (True = value) Then tf = 1
                    For i As Integer = 0 To (m_lblArr.Length - 1) Step 1
                        m_lblArr(i).Text = m_strLbl(tf, i) ' ﾗﾍﾞﾙに文字列を設定
                    Next i

                    For Each ctl As Control In m_ctlArr
                        ctl.Visible = value ' ｺﾝﾄﾛｰﾙの表示･非表示を設定
                    Next
                End Set
            End Property

        End Class
#End Region

#Region "ｶｯﾄ条件用ｸﾗｽ"
        ''' <summary>FL加工条件で表示するｶｯﾄ条件数を設定する</summary>
        Private Class CutCondition
            Friend m_ctlArr(,) As Control ' 表示･非表示をおこなうｺﾝﾄﾛｰﾙ

            ''' <summary>引数の値分のみ条件を表示し、以降を非表示にする</summary>
            ''' <value>表示する条件数</value>
            Friend WriteOnly Property Visible() As Integer
                Set(ByVal value As Integer)
                    Dim tf As Boolean
                    If (3 < value) Then value = 3
                    For i As Integer = 0 To (m_ctlArr.GetLength(0) - 1) Step 1
                        If (i < value) Then
                            tf = True
                        Else
                            tf = False
                        End If

                        For j As Integer = 0 To (m_ctlArr.GetLength(1) - 1) Step 1
                            m_ctlArr(i, j).Visible = tf
                        Next j
                    Next i
                End Set
            End Property

            ''' <summary>引数の値分のみFL設定No.ｺﾝﾎﾞﾎﾞｯｸｽを有効にし、以降を無効にする</summary>
            ''' <value>有効にする条件数</value>
            Friend WriteOnly Property Enabled() As Integer
                Set(ByVal value As Integer)
                    Dim tf As Boolean
                    If (4 < value) Then value = 4
                    For i As Integer = 0 To (m_ctlArr.GetLength(0) - 1) Step 1
                        If (i < value) Then
                            tf = True
                        Else
                            tf = False
                        End If
                        m_ctlArr(i, 1).Enabled = tf ' ＦＬ設定Ｎｏ．
                        'm_ctlArr(i, 4).Enabled = tf ' ＳＴＥＧ本数
                        m_ctlArr(i, 5).Enabled = tf ' 目標パワー
                        m_ctlArr(i, 6).Enabled = tf ' 許容範囲
                    Next i
                End Set
            End Property

        End Class
#End Region
#End Region

#Region "宣言"
        Private Const CUT_CUT As Integer = 2        ' m_CtlCutでのｲﾝﾃﾞｯｸｽ(ｶｯﾄ方法)
        Private Const CUT_CTYPE As Integer = 3      ' m_CtlCutでのｲﾝﾃﾞｯｸｽ(ｶｯﾄ形状)
        Private Const CUT_QRATE As Integer = 5      ' m_CtlCutでのｲﾝﾃﾞｯｸｽ(Ｑレート)
        Private Const CUT_SPEED As Integer = 6      ' m_CtlCutでのｲﾝﾃﾞｯｸｽ(速度)
        Private Const CUT_START_X As Integer = 7    ' m_CtlCutでのｲﾝﾃﾞｯｸｽ(カット位置Ｘ)
        Private Const CUT_START_Y As Integer = 8    ' m_CtlCutでのｲﾝﾃﾞｯｸｽ(カット位置Ｙ)
        Private Const CUT_START_2_X As Integer = 9  ' m_CtlCutでのｲﾝﾃﾞｯｸｽ(カット位置２Ｘ)
        Private Const CUT_START_2_Y As Integer = 10 ' m_CtlCutでのｲﾝﾃﾞｯｸｽ(カット位置２Ｙ)
        Private Const CUT_LEN_1 As Integer = 11     ' m_CtlCutでのｲﾝﾃﾞｯｸｽ(カット長１)
        Private Const CUT_LEN_2 As Integer = 12     ' m_CtlCutでのｲﾝﾃﾞｯｸｽ(カット長２)
        Private Const CUT_DIR_1 As Integer = 13     ' m_CtlCutでのｲﾝﾃﾞｯｸｽ(カット方向)
        Private Const CUT_DIR_2 As Integer = 14     ' m_CtlCutでのｲﾝﾃﾞｯｸｽ(カット方向２)
        Private Const CUT_OFF As Integer = 15       ' m_CtlCutでのｲﾝﾃﾞｯｸｽ(カットオフ)
        Private Const CUT_LTP As Integer = 16       ' m_CtlCutでのｲﾝﾃﾞｯｸｽ(Ｌターンポイント)
        Private Const CUT_MTYPE As Integer = 17     ' m_CtlCutでのｲﾝﾃﾞｯｸｽ(測定機器)
        Private Const CUT_TMM As Integer = 18       ' m_CtlCutでのｲﾝﾃﾞｯｸｽ(測定ﾓｰﾄﾞ)
        Private Const CUT_LETTER As Integer = 19    '###1042① m_CtlCutでのｲﾝﾃﾞｯｸｽ(測定ﾓｰﾄﾞ)

        'V2.2.1.7① ↓
        Private Const CUT_MARK_FIX As Integer = 20   ' m_CtlCutでのｲﾝﾃﾞｯｸｽ(印字固定部)
        Private Const CUT_ST_NUM As Integer = 21     ' m_CtlCutでのｲﾝﾃﾞｯｸｽ(開始番号)
        Private Const CUT_REPEAT_CNT As Integer = 22 ' m_CtlCutでのｲﾝﾃﾞｯｸｽ(重複回数)

        Private Const CUT_VAR_REPEAT As Integer = 23 '###1042① m_CtlCutでのｲﾝﾃﾞｯｸｽ(リピート有無)
        Private Const CUT_VARIATION As Integer = 24 '###1042① m_CtlCutでのｲﾝﾃﾞｯｸｽ(判定有無)
        Private Const CUT_RATE As Integer = 25      '###1042① m_CtlCutでのｲﾝﾃﾞｯｸｽ(上昇率)
        Private Const CUT_VAR_LO As Integer = 26    '###1042① m_CtlCutでのｲﾝﾃﾞｯｸｽ(下限値)
        Private Const CUT_VAR_HI As Integer = 27    '###1042① m_CtlCutでのｲﾝﾃﾞｯｸｽ(上限値)
        Private Const CUT_RETRACE As Integer = 28      'V2.0.0.0⑦リトレースの本数 'V2.1.0.0① 20からカット毎の抵抗値変化量判定項目５個分シフトで25

        ''V2.1.0.0①↓
        'Private Const CUT_VAR_REPEAT As Integer = 20 '###1042① m_CtlCutでのｲﾝﾃﾞｯｸｽ(リピート有無)
        'Private Const CUT_VARIATION As Integer = 21 '###1042① m_CtlCutでのｲﾝﾃﾞｯｸｽ(判定有無)
        'Private Const CUT_RATE As Integer = 22      '###1042① m_CtlCutでのｲﾝﾃﾞｯｸｽ(上昇率)
        'Private Const CUT_VAR_LO As Integer = 23    '###1042① m_CtlCutでのｲﾝﾃﾞｯｸｽ(下限値)
        'Private Const CUT_VAR_HI As Integer = 24    '###1042① m_CtlCutでのｲﾝﾃﾞｯｸｽ(上限値)
        ''V2.1.0.0①↑
        'Private Const CUT_RETRACE As Integer = 25      'V2.0.0.0⑦リトレースの本数 'V2.1.0.0① 20からカット毎の抵抗値変化量判定項目５個分シフトで25
        'V2.2.1.7① ↑

        'V2.0.0.0⑦        Private Const CUT_TR_Q As Integer = 20      '###1042① m_CtlCutでのｲﾝﾃﾞｯｸｽ(測定ﾓｰﾄﾞ)
        'V2.0.0.0⑦        Private Const CUT_TR_SPEED As Integer = 21  '###1042① m_CtlCutでのｲﾝﾃﾞｯｸｽ(測定ﾓｰﾄﾞ)

        Private Const IDX_MTYPE As Integer = 4      ' m_CtlIdxCutでの2次元目のｲﾝﾃﾞｯｸｽ(測定機器)
        Private Const IDX_TMM As Integer = 5        ' m_CtlIdxCutでの2次元目のｲﾝﾃﾞｯｸｽ(測定ﾓｰﾄﾞ)

        'V1.0.4.3③ CNS_CUTM_TRに変更        Private Const CMB_CUT_TRACK As Integer = 1  ' ｶｯﾄ方法ｺﾝﾎﾞﾎﾞｯｸｽのﾘｽﾄｲﾝﾃﾞｯｸｽ+1(ﾄﾗｯｷﾝｸﾞ)
        'V1.0.4.3③ CNS_CUTM_IXに変更        Private Const CMB_CUT_IDX As Integer = 2    ' ｶｯﾄ方法ｺﾝﾎﾞﾎﾞｯｸｽのﾘｽﾄｲﾝﾃﾞｯｸｽ+1(ｲﾝﾃﾞｯｸｽ)
        'V1.0.4.3③ CNS_CUTM_NGに変更        Private Const CMB_CUT_NG As Integer = 3     ' ｶｯﾄ方法ｺﾝﾎﾞﾎﾞｯｸｽのﾘｽﾄｲﾝﾃﾞｯｸｽ+1(NGｶｯﾄ)

        'V1.0.4.3③ CNS_CUTP_STに変更        Private Const CMB_CTYP_STR As Integer = 1   ' ｶｯﾄ形状ｺﾝﾎﾞﾎﾞｯｸｽのﾘｽﾄｲﾝﾃﾞｯｸｽ+1(ｽﾄﾚｰﾄ)
        'V1.0.4.3③ CNS_CUTP_Lに変更        Private Const CMB_CTYP_LCUT As Integer = 2  ' ｶｯﾄ形状ｺﾝﾎﾞﾎﾞｯｸｽのﾘｽﾄｲﾝﾃﾞｯｸｽ+1(Lｶｯﾄ)
        'V1.0.4.3③ CNS_CUTP_SPに変更        Private Const CMB_CTYP_SPT As Integer = 3   ' ｶｯﾄ形状ｺﾝﾎﾞﾎﾞｯｸｽのﾘｽﾄｲﾝﾃﾞｯｸｽ+1(ｻｰﾍﾟﾝﾀｲﾝ)

        Private m_CtlCut() As Control           ' ｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽのｺﾝﾄﾛｰﾙ配列
        Private m_CtlIdxCut(,) As Control       ' ｲﾝﾃﾞｯｸｽｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽのｺﾝﾄﾛｰﾙ配列
        Private m_CtlFLCnd(,) As Control        ' FL加工条件ｸﾞﾙｰﾌﾟﾎﾞｯｸｽのｺﾝﾄﾛｰﾙ配列
        Private m_CtlLCut(,) As Control         'V1.0.4.3③ Lカットｸﾞﾙｰﾌﾟﾎﾞｯｸｽのｺﾝﾄﾛｰﾙ配列
        Private m_CtlRetraceCut(,) As Control   'V2.0.0.0⑦ リトレースカットｸﾞﾙｰﾌﾟﾎﾞｯｸｽのｺﾝﾄﾛｰﾙ配列
        Private Const LCUT_DIR_IDX As Integer = 3      'V1.0.4.3③ m_CtlLCutで角度（プルダウン）の２次元目の配列番号

        Private m_CtlUCut() As Control         ' Uカット用パラメータ追加      'V2.2.0.0②

        Private m_NGCut As NGCut                ' NGｶｯﾄ関連の有効･無効を切替える
        'Private m_Serpentine As Serpentine      ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ関連の表示･非表示をおこなう
        Private m_CutCondition As CutCondition  ' FL加工条件で表示するｶｯﾄ条件数を設定する

        Private Const MAX_DEGREES As Integer = 359
        Private AngleArray(MAX_DEGREES, 1) As Integer           'V1.0.4.3②
        Private AngleArrayForLcut(MAX_DEGREES, 1) As Integer    'V1.0.4.3②
        Private AngleArrayForUcut(MAX_DEGREES, 1) As Integer    'V2.2.0.0②

        'V2.0.0.0↓
        ''' <summary>
        ''' 誤差のラベル表示文字
        ''' </summary>
        ''' <remarks></remarks>
        Private m_strDev As String

        ''' <summary>
        ''' 誤差比率指定時の最小値/最大値
        ''' </summary>
        ''' <remarks>0=最小値、1=最大値</remarks>
        Private m_strDEVRaite(2) As String

        ''' <summary>
        ''' 誤差絶対値指定時の最小値/最大値
        ''' </summary>
        ''' <remarks>0=最小値、1=最大値</remarks>
        Private m_strDEVAbsolute(2) As String

        ''' <summary>
        ''' ｶｯﾄ方法のコンボボックスデータリスト
        ''' </summary>
        ''' <remarks></remarks>
        Private m_lstCutMethod As New List(Of ComboDataStruct)

        Private m_lstCutType As New List(Of ComboDataStruct)
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

            Try
                ' ｶｯﾄ方法の初期データ
                Call InitCutMethodData()

                ' カット形状の初期データ
                Call InitCutTypeData()

                ' EDIT_DEF_User.iniからﾀﾌﾞ名を設定
                TAB_NAME = GetPrivateProfileString_S("CUT_LABEL", "TAB_NAM", m_sPath, "????")

                ' 発振器種別がﾌｧｲﾊﾞﾚｰｻﾞでない場合はFL加工条件ｸﾞﾙｰﾌﾟﾎﾞｯｸｽを非表示にする
                If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then
                    CGrp_2.Visible = False
                Else ' ﾌｧｲﾊﾞﾚｰｻﾞの場合ｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽのQﾚｰﾄを非表示にする
                    CLbl_7.Visible = False
                    CTxt_1.Visible = False
                End If

                ' 追加･削除･加工条件ﾎﾞﾀﾝの設定
                With mainEdit
                    CBtn_Add.SetLblToolTip(.LblToolTip)
                    CBtn_Add.Text = GetPrivateProfileString_S("BUTTON_LABEL", "BTN_ADD", m_sPath, "ADD")
                    CBtn_Del.SetLblToolTip(.LblToolTip)
                    CBtn_Del.Text = GetPrivateProfileString_S("BUTTON_LABEL", "BTN_DEL", m_sPath, "DEL")
                    CBtn_FLC.SetLblToolTip(.LblToolTip)
                    CBtn_FLC.Text = GetPrivateProfileString_S("BUTTON_LABEL", "BTN_FLC", m_sPath, "Condition")
                End With

                ' ----------------------------------------------------------
                ' EDIT_DEF_User.iniからｸﾞﾙｰﾌﾟﾎﾞｯｸｽに表示名を設定
                ' ----------------------------------------------------------
                ' 'V1.0.4.3③ CGrp_3 追加 'V2.0.0.0⑦ CGrp_4 追加 'V2.2.0.0②  CGrp_5追加
                GrpArray = New cGrp_() {
                    CGrp_0, CGrp_1, CGrp_2, CGrp_3, CGrp_4, CGrp_5
                }
                For i As Integer = 0 To (GrpArray.Length - 1) Step 1
                    With GrpArray(i)
                        .TabIndex = ((254 - GrpArray.Length) + i) ' ｶｰｿﾙｷｰによるﾌｫｰｶｽ移動で必要
                        .Tag = i
                        .Text = GetPrivateProfileString_S( _
                                    "CUT_LABEL", (i.ToString("000") & "_GRP"), m_sPath, "????")
                    End With
                Next i
                ' 追加･削除ﾎﾞﾀﾝのﾊﾟﾈﾙ
                CPnl_Btn.TabIndex = 254 ' ｺﾝﾄﾛｰﾙ配置可能最大数(最後に設定)

                ' ----------------------------------------------------------
                ' EDIT_DEF_User.iniからﾗﾍﾞﾙに表示名を設定
                ' ----------------------------------------------------------
                LblArray = New cLbl_() {
                    CLbl_0, CLbl_1, CLbl_2, CLbl_3, CLbl_4, CLbl_5, CLbl_6, CLbl_7, CLbl_8,
                    CLbl_9, CLbl_10, CLbl_11, CLbl_12,
                    CLbl_13, CLbl_14, CLbl_15, CLbl_16, _
 _
                    CLbl_17, CLbl_18, CLbl_19, CLbl_20, CLbl_21, CLbl_22,
                    CLbl_23, CLbl_24, CLbl_25, CLbl_26, CLbl_27, _
 _
                    CLbl_28, CLbl_29, CLbl_30, CLbl_31,
                    CLbl_32, CLbl_33, CLbl_34, CLbl_35,
                    CLbl_80, CLbl_81, CLbl_82 'V2.2.1.7①
                }
                For i As Integer = 0 To (LblArray.Length - 1) Step 1
                    LblArray(i).Text = GetPrivateProfileString_S( _
                            "CUT_LABEL", (i.ToString("000") & "_LBL"), m_sPath, "????")
                Next i

                ' ----------------------------------------------------------
                ' NGｶｯﾄ用ｸﾗｽの設定
                ' ----------------------------------------------------------
                m_NGCut = New NGCut()
                With m_NGCut
                    .m_CtlArr = New Control() { _
                        CTxt_9, CCmb_6, CCmb_7 _
                    }
                End With

                ' ----------------------------------------------------------
                ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ用ｸﾗｽの設定
                ' ----------------------------------------------------------
                'm_Serpentine = New Serpentine()
                'With m_Serpentine
                '    .m_ctlArr = New Control() { _
                '        CLbl_6, CTxt_0, _
                '        CLbl_10, CTxt_5, CTxt_6, _
                '        CTxt_8, _
                '        CCmb_5 _
                '    }

                '    .m_lblArr = New Label() { _
                '        CLbl_11, CLbl_12 _
                '    }

                '    For i As Integer = 0 To (.m_strLbl.GetLength(0) - 1) Step 1
                '        For j As Integer = 0 To (.m_strLbl.GetLength(1) - 1) Step 1
                '            Dim no As Integer = Convert.ToInt32((i.ToString() & j.ToString()), 2)
                '            .m_strLbl(i, j) = GetPrivateProfileString_S( _
                '                    "CUT_SERPENTINE", (no.ToString("000") & "_LBL"), m_sPath, "????")
                '        Next j
                '    Next i
                'End With

                ' ----------------------------------------------------------
                ' ｶｯﾄ条件用ｸﾗｽの設定
                ' ----------------------------------------------------------
                m_CutCondition = New CutCondition()
                With m_CutCondition
                    .m_ctlArr = New Control(,) { _
                        {CLbl_32, CCmb_18, CTxt_31, CTxt_32, CTxt_33}, _
                        {CLbl_33, CCmb_19, CTxt_34, CTxt_35, CTxt_36}, _
                        {CLbl_34, CCmb_20, CTxt_37, CTxt_38, CTxt_39}, _
                        {CLbl_35, CCmb_21, CTxt_40, CTxt_41, CTxt_42} _
                    }
                End With

                ' ----------------------------------------------------------
                ' ｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のｺﾝﾄﾛｰﾙを設定(CtlArrayの順番と合わせる)
                ' ----------------------------------------------------------
                ' ###1042① CTxt_43（文字）追加 'V2.1.0.0①CCmb_22,CCmb_23,CTxt_46,CTxt_47,CTxt_48追加
                m_CtlCut = New Control() {
                    CCmb_0, CCmb_1, CCmb_2, CCmb_3, CTxt_0, CTxt_1, CTxt_2,
                    CTxt_3, CTxt_4, CTxt_5, CTxt_6, CTxt_7, CTxt_8, CCmb_4, CCmb_5,
                    CTxt_9, CTxt_10, CCmb_6, CCmb_7, CTxt_43,
                    CTxt_80, CTxt_81, CTxt_82, 'V2.2.1.7①
                    CCmb_22, CCmb_23, CTxt_46, CTxt_47, CTxt_48,
                    CRT_Num
                }
                Call SetControlData(m_CtlCut) ' m_Serpentine,m_CutConditionの設定より後におこなう
                CCmb_4.DropDownStyle = ComboBoxStyle.DropDown   'V1.1.0.0① 角度手入力可
                'CCmb_5.DropDownStyle = ComboBoxStyle.DropDown   'V1.1.0.0① 角度手入力可

                ' ----------------------------------------------------------
                ' ｲﾝﾃﾞｯｸｽｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のｺﾝﾄﾛｰﾙを設定(CtlArrayの順番と合わせる)
                ' ----------------------------------------------------------
                m_CtlIdxCut = New Control(,) { _
                    {CTxt_11, CTxt_12, CTxt_13, CTxt_14, CCmb_8, CCmb_9}, _
                    {CTxt_15, CTxt_16, CTxt_17, CTxt_18, CCmb_10, CCmb_11}, _
                    {CTxt_19, CTxt_20, CTxt_21, CTxt_22, CCmb_12, CCmb_13}, _
                    {CTxt_23, CTxt_24, CTxt_25, CTxt_26, CCmb_14, CCmb_15}, _
                    {CTxt_27, CTxt_28, CTxt_29, CTxt_30, CCmb_16, CCmb_17} _
                }
                Call SetControlData(m_CtlIdxCut)

                'V1.0.4.3③ ADD ↓
                ' ----------------------------------------------------------
                ' Ｌカットパラメータｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のｺﾝﾄﾛｰﾙを設定(CtlArrayの順番と合わせる)
                ' ----------------------------------------------------------
                m_CtlLCut = New Control(,) { _
                    {CL_Len1, CL_Q1, CL_Spd1, CCmb_Dir_1, CL_Tpt1}, _
                    {CL_Len2, CL_Q2, CL_Spd2, CCmb_Dir_2, CL_Tpt2}, _
                    {CL_Len3, CL_Q3, CL_Spd3, CCmb_Dir_3, CL_Tpt3}, _
                    {CL_Len4, CL_Q4, CL_Spd4, CCmb_Dir_4, CL_Tpt4}, _
                    {CL_Len5, CL_Q5, CL_Spd5, CCmb_Dir_5, CL_Tpt5}, _
                    {CL_Len6, CL_Q6, CL_Spd6, CCmb_Dir_6, CL_Tpt6}, _
                    {CL_Len7, CL_Q7, CL_Spd7, CCmb_Dir_7, CL_Tpt7} _
                }
                Call SetControlData(m_CtlLCut)
                'V1.0.4.3③ ADD ↑
                'V2.0.0.0⑦ ADD ↓
                ' ----------------------------------------------------------
                ' リトレースカットパラメータｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のｺﾝﾄﾛｰﾙを設定(CtlArrayの順番と合わせる)
                ' ----------------------------------------------------------
                m_CtlRetraceCut = New Control(,) { _
                    {RTOFFX_1, RTOFFY_1, RTQRATE_1, RTV_1}, _
                    {RTOFFX_2, RTOFFY_2, RTQRATE_2, RTV_2}, _
                    {RTOFFX_3, RTOFFY_3, RTQRATE_3, RTV_3}, _
                    {RTOFFX_4, RTOFFY_4, RTQRATE_4, RTV_4}, _
                    {RTOFFX_5, RTOFFY_5, RTQRATE_5, RTV_5}, _
                    {RTOFFX_6, RTOFFY_6, RTQRATE_6, RTV_6}, _
                    {RTOFFX_7, RTOFFY_7, RTQRATE_7, RTV_7}, _
                    {RTOFFX_8, RTOFFY_8, RTQRATE_8, RTV_8}, _
                    {RTOFFX_9, RTOFFY_9, RTQRATE_9, RTV_9}, _
                    {RTOFFX_10, RTOFFY_10, RTQRATE_10, RTV_10} _
                }
                Call SetControlData(m_CtlRetraceCut)
                'V2.0.0.0⑦ ADD ↑
                'V2.2.0.0②↓
                ' ----------------------------------------------------------
                ' Uカットパラメータ：ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のｺﾝﾄﾛｰﾙを設定(CtlArrayの順番と合わせる)
                ' ----------------------------------------------------------
                m_CtlUCut = New Control() {
                    UclTxt_1, UclTxt_2, UcR1Txt_1, UcR1Txt_2,
                    UcqTxt_1, UcspdTxt_1, Ucdircmb, UcTurnCmb,
                    UcTurnTxt
                }
                Call SetControlData(m_CtlUCut)
                'V2.2.0.0②↑

                ' ----------------------------------------------------------
                ' FL加工条件ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のｺﾝﾄﾛｰﾙを設定(CtlArrayの順番と合わせる)
                ' ----------------------------------------------------------
                m_CtlFLCnd = New Control(,) { _
                    {CCmb_18, CTxt_31, CTxt_32, CTxt_33}, _
                    {CCmb_19, CTxt_34, CTxt_35, CTxt_36}, _
                    {CCmb_20, CTxt_37, CTxt_38, CTxt_39}, _
                    {CCmb_21, CTxt_40, CTxt_41, CTxt_42} _
                }
                Call SetControlData(m_CtlFLCnd)

                ' ----------------------------------------------------------
                ' すべてのﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽ･ﾎﾞﾀﾝに対し、ﾌｫｰｶｽ移動設定をおこなう
                ' ﾀﾌﾞｷｰ、ｶｰｿﾙｷｰによりﾌｫｰｶｽ移動する順番でｺﾝﾄﾛｰﾙをCtlArrayに設定する
                ' 使用しないｺﾝﾄﾛｰﾙは Enabled=False または Visible=False にする
                ' ----------------------------------------------------------
                ' ###1042① CTxt_43（文字）追加'V2.1.0.0①CCmb_22,CCmb_23,CTxt_46,CTxt_47,CTxt_48追加
                'V2.0.0.0⑦ CTxt_44, CTxt_45削除、CRT_Num追加
                CtlArray = New Control() {
                    CCmb_0, CCmb_1, CCmb_2, CCmb_3, CTxt_0, CTxt_1, CTxt_2,
                    CTxt_3, CTxt_4, CTxt_5, CTxt_6, CTxt_7, CTxt_8, CCmb_4, CCmb_5,
                    CTxt_9, CTxt_10, CCmb_6, CCmb_7, CTxt_43,
                    CTxt_80, CTxt_81, CTxt_82,'V2.2.1.7①
                    CCmb_22, CCmb_23, CTxt_46, CTxt_47, CTxt_48,
                    CRT_Num, _
 _
                    CTxt_11, CTxt_12, CTxt_13, CTxt_14, CCmb_8, CCmb_9,
                    CTxt_15, CTxt_16, CTxt_17, CTxt_18, CCmb_10, CCmb_11,
                    CTxt_19, CTxt_20, CTxt_21, CTxt_22, CCmb_12, CCmb_13,
                    CTxt_23, CTxt_24, CTxt_25, CTxt_26, CCmb_14, CCmb_15,
                    CTxt_27, CTxt_28, CTxt_29, CTxt_30, CCmb_16, CCmb_17, _
 _
                    CL_Len1, CL_Q1, CL_Spd1, CCmb_Dir_1, CL_Tpt1,
                    CL_Len2, CL_Q2, CL_Spd2, CCmb_Dir_2, CL_Tpt2,
                    CL_Len3, CL_Q3, CL_Spd3, CCmb_Dir_3, CL_Tpt3,
                    CL_Len4, CL_Q4, CL_Spd4, CCmb_Dir_4, CL_Tpt4,
                    CL_Len5, CL_Q5, CL_Spd5, CCmb_Dir_5, CL_Tpt5,
                    CL_Len6, CL_Q6, CL_Spd6, CCmb_Dir_6, CL_Tpt6,
                    CL_Len7, CL_Q7, CL_Spd7, CCmb_Dir_7, _
 _
                    CRT_Num,
                    RTOFFX_1, RTOFFY_1, RTQRATE_1, RTV_1,
                    RTOFFX_2, RTOFFY_2, RTQRATE_2, RTV_2,
                    RTOFFX_3, RTOFFY_3, RTQRATE_3, RTV_3,
                    RTOFFX_4, RTOFFY_4, RTQRATE_4, RTV_4,
                    RTOFFX_5, RTOFFY_5, RTQRATE_5, RTV_5,
                    RTOFFX_6, RTOFFY_6, RTQRATE_6, RTV_6,
                    RTOFFX_7, RTOFFY_7, RTQRATE_7, RTV_7,
                    RTOFFX_8, RTOFFY_8, RTQRATE_8, RTV_8,
                    RTOFFX_9, RTOFFY_9, RTQRATE_9, RTV_9,
                    RTOFFX_10, RTOFFY_10, RTQRATE_10, RTV_10, _
 _
                    UclTxt_1, UclTxt_2, UcR1Txt_1, UcR1Txt_2,
                    UcqTxt_1, UcspdTxt_1, Ucdircmb, UcTurnCmb,
                    UcTurnTxt, _
 _
                    CCmb_18, CCmb_19, CCmb_20, CCmb_21, CBtn_FLC,
                    CBtn_Add, CBtn_Del
                }
                Call SetTabIndex(CtlArray) ' ﾀﾌﾞｲﾝﾃﾞｯｸｽとKeyDownｲﾍﾞﾝﾄを設定する

                ' ----------------------------------------------------------
                ' 画面表示時にﾌｫｰｶｽされるｺﾝﾄﾛｰﾙを設定する
                ' ----------------------------------------------------------
                FIRST_CONTROL = CtlArray(0) ' 抵抗番号ｺﾝﾎﾞﾎﾞｯｸｽが選択されるようにする

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
            Dim Cnt As Integer  'V1.0.4.3②
            Dim i As Integer

            Try
                With cCombo
                    tag = DirectCast(.Tag, Integer)
                    .Items.Clear()
                    Select Case (DirectCast(.Parent.Tag, Integer)) ' ｸﾞﾙｰﾌﾟﾎﾞｯｸｽのﾀｸﾞ
                        ' ------------------------------------------------------------------------------
                        Case 0 ' ｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            Select Case (tag)
                                Case 0 ' 抵抗番号
                                    '.Items.Add("") ' ﾚｲｱｳﾄｲﾍﾞﾝﾄで設定される
                                Case 1 ' ｶｯﾄ番号
                                    '.Items.Add("") ' ﾚｲｱｳﾄｲﾍﾞﾝﾄで設定される
                                Case 2 ' ｶｯﾄ方法(1:ﾄﾗｯｷﾝｸﾞ, 2:ｲﾝﾃﾞｯｸｽ)
                                    For i = 0 To m_lstCutMethod.Count - 1 Step 1
                                        .Items.Add(m_lstCutMethod(i).Name)
                                    Next i
                                Case 3 ' ｶｯﾄ形状(1:ｽﾄﾚｰﾄ, 2:Lｶｯﾄ, 3:ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ)
                                    For i = 0 To m_lstCutType.Count - 1 Step 1
                                        .Items.Add(m_lstCutType(i).Name)
                                    Next i
                                Case 4, 5 ' ｶｯﾄ方向1, ｶｯﾄ方向2(90°単位　0°～360°)
                                    'V1.0.4.3② ADD START↓
                                    For Cnt = 0 To MAX_DEGREES
                                        AngleArray(Cnt, 0) = Cnt
                                        AngleArray(Cnt, 1) = -1
                                    Next
                                    'V1.0.4.3② ADD END↑
                                    .Items.Add("     0°")
                                    AngleArray(0, 1) = 0       'V1.0.4.3②
                                    .Items.Add("    90°")
                                    AngleArray(90, 1) = 1       'V1.0.4.3②
                                    .Items.Add("   180°")
                                    AngleArray(180, 1) = 2       'V1.0.4.3②
                                    .Items.Add("   270°")
                                    AngleArray(270, 1) = 3       'V1.0.4.3②
                                    .Items.Add("    10°")
                                    AngleArray(10, 1) = 4       'V1.0.4.3②
                                    .Items.Add("    20°")
                                    AngleArray(20, 1) = 5       'V1.0.4.3②
                                    .Items.Add("    30°")
                                    AngleArray(30, 1) = 6       'V1.0.4.3②
                                    .Items.Add("    40°")
                                    AngleArray(40, 1) = 7       'V1.0.4.3②
                                    .Items.Add("    50°")
                                    AngleArray(50, 1) = 8       'V1.0.4.3②
                                    .Items.Add("    60°")
                                    AngleArray(60, 1) = 9       'V1.0.4.3②
                                    .Items.Add("    70°")
                                    AngleArray(70, 1) = 10       'V1.0.4.3②
                                    .Items.Add("    80°")
                                    AngleArray(80, 1) = 11       'V1.0.4.3②
                                    .Items.Add("   100°")
                                    AngleArray(100, 1) = 12       'V1.0.4.3②
                                    .Items.Add("   110°")
                                    AngleArray(110, 1) = 13       'V1.0.4.3②
                                    .Items.Add("   120°")
                                    AngleArray(120, 1) = 14       'V1.0.4.3②
                                    .Items.Add("   130°")
                                    AngleArray(130, 1) = 15       'V1.0.4.3②
                                    .Items.Add("   140°")
                                    AngleArray(140, 1) = 16       'V1.0.4.3②
                                    .Items.Add("   150°")
                                    AngleArray(150, 1) = 17       'V1.0.4.3②
                                    .Items.Add("   160°")
                                    AngleArray(160, 1) = 18       'V1.0.4.3②
                                    .Items.Add("   170°")
                                    AngleArray(170, 1) = 19       'V1.0.4.3②
                                Case 6 ' 測定機器(0:内部測定器, 1以上は外部測定器番号)
                                    '.Items.Add("") ' ﾚｲｱｳﾄｲﾍﾞﾝﾄで設定される
                                Case 7 ' 測定ﾓｰﾄﾞ(0:高速, 1:高精度)
                                    .Items.Add("高速")
                                    .Items.Add("高精度")
                                    'V2.1.0.0①↓
                                Case 8
                                    .Items.Add("なし")
                                    .Items.Add("あり")
                                Case 9
                                    .Items.Add("なし")
                                    .Items.Add("あり")
                                    'V2.1.0.0①↑
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case 1 ' ｲﾝﾃﾞｯｸｽｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            Select Case (tag)
                                Case 0 ' 測定機器(0:内部測定器, 1以上は外部測定器番号)
                                    '.Items.Add("") ' ﾚｲｱｳﾄｲﾍﾞﾝﾄで設定される
                                Case 1 ' 測定ﾓｰﾄﾞ(0:高速, 1:高精度)
                                    '.Items.Add("") ' ﾚｲｱｳﾄｲﾍﾞﾝﾄで設定される
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case 2 ' FL加工条件ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            Select Case (tag)
                                Case 0 ' FL設定No.
                                    For i = 0 To 31 Step 1
                                        .Items.Add(String.Format("{0,5:#0}", i))
                                    Next i
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                            'V1.0.4.3③↓
                        Case 3 ' Ｌカットパラメータグループボックス
                            Select Case (tag)
                                Case 0 ' FL設定No.
                                    .DropDownStyle = ComboBoxStyle.DropDown
                                    For Cnt = 0 To MAX_DEGREES
                                        AngleArrayForLcut(Cnt, 0) = Cnt
                                        AngleArrayForLcut(Cnt, 1) = -1
                                    Next
                                    .Items.Add("     0°")
                                    AngleArrayForLcut(0, 1) = 0       'V1.0.4.3②
                                    .Items.Add("    90°")
                                    AngleArrayForLcut(90, 1) = 1       'V1.0.4.3②
                                    .Items.Add("   180°")
                                    AngleArrayForLcut(180, 1) = 2       'V1.0.4.3②
                                    .Items.Add("   270°")
                                    AngleArrayForLcut(270, 1) = 3       'V1.0.4.3②
                                    .Items.Add("    10°")
                                    AngleArrayForLcut(10, 1) = 4       'V1.0.4.3②
                                    .Items.Add("    20°")
                                    AngleArrayForLcut(20, 1) = 5       'V1.0.4.3②
                                    .Items.Add("    30°")
                                    AngleArrayForLcut(30, 1) = 6       'V1.0.4.3②
                                    .Items.Add("    40°")
                                    AngleArrayForLcut(40, 1) = 7       'V1.0.4.3②
                                    .Items.Add("    50°")
                                    AngleArrayForLcut(50, 1) = 8       'V1.0.4.3②
                                    .Items.Add("    60°")
                                    AngleArrayForLcut(60, 1) = 9       'V1.0.4.3②
                                    .Items.Add("    70°")
                                    AngleArrayForLcut(70, 1) = 10       'V1.0.4.3②
                                    .Items.Add("    80°")
                                    AngleArrayForLcut(80, 1) = 11       'V1.0.4.3②
                                    .Items.Add("   100°")
                                    AngleArrayForLcut(100, 1) = 12       'V1.0.4.3②
                                    .Items.Add("   110°")
                                    AngleArrayForLcut(110, 1) = 13       'V1.0.4.3②
                                    .Items.Add("   120°")
                                    AngleArrayForLcut(120, 1) = 14       'V1.0.4.3②
                                    .Items.Add("   130°")
                                    AngleArrayForLcut(130, 1) = 15       'V1.0.4.3②
                                    .Items.Add("   140°")
                                    AngleArrayForLcut(140, 1) = 16       'V1.0.4.3②
                                    .Items.Add("   150°")
                                    AngleArrayForLcut(150, 1) = 17       'V1.0.4.3②
                                    .Items.Add("   160°")
                                    AngleArrayForLcut(160, 1) = 18       'V1.0.4.3②
                                    .Items.Add("   170°")
                                    AngleArrayForLcut(170, 1) = 19       'V1.0.4.3②
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                            'V1.0.4.3③↑

                            'V2.2.0.0②↓
                        Case 5 ' Ｕカットパラメータグループボックス
                            Select Case (tag)
                                Case 0 ' 角度
                                    For Cnt = 0 To MAX_DEGREES
                                        AngleArrayForUcut(Cnt, 0) = Cnt
                                        AngleArrayForUcut(Cnt, 1) = -1
                                    Next
                                    'V1.0.4.3② ADD END↑
                                    .Items.Add("     0°")
                                    AngleArrayForUcut(0, 1) = 0
                                    .Items.Add("    90°")
                                    AngleArrayForUcut(90, 1) = 1
                                    .Items.Add("   180°")
                                    AngleArrayForUcut(180, 1) = 2
                                    .Items.Add("   270°")
                                    AngleArrayForUcut(270, 1) = 3

                                Case 1 ' ターン方向 
                                    .Items.Add("ＣＷ")
                                    .Items.Add("ＣＣＷ")
                            End Select
                            'V2.2.0.0②↑
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
            Dim strFlg As Boolean = False       ' 格納する値の種類(False=数値,True=文字列) ###1042①
            Try
                tag = DirectCast(cTextBox.Tag, Integer)
                no = tag.ToString("000")

                Select Case (DirectCast(cTextBox.Parent.Tag, Integer)) ' ｸﾞﾙｰﾌﾟﾎﾞｯｸｽのﾀｸﾞ
                    ' ------------------------------------------------------------------------------
                    Case 0 ' ｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                        ' ﾃﾞｰﾀﾁｪｯｸｴﾗｰ時の表示名
                        strMsg = GetPrivateProfileString_S("CUT_CUT", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0 ' ｻｰﾍﾟﾝﾀｲﾝ本数
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "10")
                            Case 1 ' Qﾚｰﾄ
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "0.1")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "40.0")
                            Case 2 ' 速度
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "0.1")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "400.0")
                            Case 3 ' ｶｯﾄ位置X
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "-80.0")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "80.0")
                            Case 4 ' ｶｯﾄ位置Y
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "-80.0")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "80.0")
                            Case 5 ' ｶｯﾄ位置2X
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "-80.0")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "80.0")
                            Case 6 ' ｶｯﾄ位置2Y
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "-80.0")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "80.0")
                            Case 7 ' ｶｯﾄ長1
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "0.001")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "20.000")
                            Case 8 ' ｶｯﾄ長2
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "0.001")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "20.000")
                            Case 9 ' ｶｯﾄｵﾌ
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "-99.99")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "99.99")
                            Case 10 ' Lﾀｰﾝﾎﾟｲﾝﾄ
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "100.0")
                            Case 11 ' 文字　###1042①
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "10")
                                strFlg = True

                                'V2.2.1.7① ↓
                            Case 12 ' 印字固定部
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "14")
                                strFlg = True
                            Case 13 ' 開始番号
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "6")
                                strFlg = True
                            Case 14 ' 重複回数
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "100")
                            Case 15 ' 上昇率
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "-99.99")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "99.99")
                            Case 16 ' 下限値
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "-99.99")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "99.99")
                            Case 17 ' 上限値
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "-99.99")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "99.99")

                                '    'V2.1.0.0①↓
                                'Case 12 ' 上昇率
                                '    strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "-99.99")
                                '    strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "99.99")
                                'Case 13 ' 下限値
                                '    strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "-99.99")
                                '    strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "99.99")
                                'Case 14 ' 上限値
                                '    strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "-99.99")
                                '    strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "99.99")
                                'V2.2.1.7① ↑

                                'V2.1.0.0①↑
                                'V2.0.0.0⑦                            Case 12 ' リトレースQレート V1.0.4.3③
                                'V2.0.0.0⑦                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "1")
                                'V2.0.0.0⑦                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "10")
                                'V2.0.0.0⑦                            Case 13 ' リトレース速度 V1.0.4.3③
                                'V2.0.0.0⑦                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "1")
                                'V2.0.0.0⑦                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "10")
                            Case Else
                                Throw New Exception("Case " & tag & ": Nothing")
                        End Select
                        ' ------------------------------------------------------------------------------
                    Case 1 ' ｲﾝﾃﾞｯｸｽｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                        ' ﾃﾞｰﾀﾁｪｯｸｴﾗｰ時の表示名
                        strMsg = GetPrivateProfileString_S("CUT_INDEX", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0 ' ｶｯﾄ回数
                                strMin = GetPrivateProfileString_S("CUT_INDEX", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("CUT_INDEX", (no & "_MAX"), m_sPath, "999")
                            Case 1 ' ｶｯﾄ長(ｲﾝﾃﾞｯｸｽﾋﾟｯﾁ)
                                strMin = GetPrivateProfileString_S("CUT_INDEX", (no & "_MIN"), m_sPath, "0.000")
                                strMax = GetPrivateProfileString_S("CUT_INDEX", (no & "_MAX"), m_sPath, "20.000")
                            Case 2 ' ﾋﾟｯﾁ間ﾎﾟｰｽﾞ
                                strMin = GetPrivateProfileString_S("CUT_INDEX", (no & "_MIN"), m_sPath, "0")
                                strMax = GetPrivateProfileString_S("CUT_INDEX", (no & "_MAX"), m_sPath, "32767")
                            Case 3 ' 誤差
                                strMin = GetPrivateProfileString_S("CUT_INDEX", (no & "_MIN"), m_sPath, "0.00")
                                strMax = GetPrivateProfileString_S("CUT_INDEX", (no & "_MAX"), m_sPath, "99.99")
                            Case Else
                                Throw New Exception("Case " & tag & ": Nothing")
                        End Select
                        ' ------------------------------------------------------------------------------
                    Case 2 ' FL加工条件ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                        ' ﾃﾞｰﾀﾁｪｯｸｴﾗｰ時の表示名
                        strMsg = GetPrivateProfileString_S("CUT_FLCOND", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0 ' 電流値
                                strMin = GetPrivateProfileString_S("CUT_FLCOND", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("CUT_FLCOND", (no & "_MAX"), m_sPath, "1000")
                            Case 1 ' Qﾚｰﾄ
                                strMin = GetPrivateProfileString_S("CUT_FLCOND", (no & "_MIN"), m_sPath, "0.1")
                                strMax = GetPrivateProfileString_S("CUT_FLCOND", (no & "_MAX"), m_sPath, "40.0")
                            Case 2 ' STEG本数
                                strMin = GetPrivateProfileString_S("CUT_FLCOND", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("CUT_FLCOND", (no & "_MAX"), m_sPath, "15")
                            Case Else
                                Throw New Exception("Case " & tag & ": Nothing")
                        End Select
                        ' ------------------------------------------------------------------------------
                    Case 3
                        strMsg = GetPrivateProfileString_S("CUT_LCUT_PARA", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0 ' カット長
                                strMin = GetPrivateProfileString_S("CUT_LCUT_PARA", (no & "_MIN"), m_sPath, "0.001")
                                strMax = GetPrivateProfileString_S("CUT_LCUT_PARA", (no & "_MAX"), m_sPath, "20.000")
                            Case 1 ' Qレート
                                strMin = GetPrivateProfileString_S("CUT_LCUT_PARA", (no & "_MIN"), m_sPath, "0.1")
                                strMax = GetPrivateProfileString_S("CUT_LCUT_PARA", (no & "_MAX"), m_sPath, "40.0")
                            Case 2 ' 速度
                                strMin = GetPrivateProfileString_S("CUT_LCUT_PARA", (no & "_MIN"), m_sPath, "0.1")
                                strMax = GetPrivateProfileString_S("CUT_LCUT_PARA", (no & "_MAX"), m_sPath, "400.0")
                            Case 3 ' Lﾀｰﾝﾎﾟｲﾝﾄ
                                strMin = GetPrivateProfileString_S("CUT_LCUT_PARA", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("CUT_LCUT_PARA", (no & "_MAX"), m_sPath, "100.0")
                            Case Else
                                Throw New Exception("Case " & tag & ": Nothing")
                        End Select
                        'V2.0.0.0⑦↓
                    Case 4
                        strMsg = GetPrivateProfileString_S("CUT_RETRACE_PARA", (no & "_MSG"), m_sPath, "??????")
                        Select Case (tag)
                            Case 0 ' リトレースのオフセットＸ
                                strMin = GetPrivateProfileString_S("CUT_RETRACE_PARA", (no & "_MIN"), m_sPath, "-10.0")
                                strMax = GetPrivateProfileString_S("CUT_RETRACE_PARA", (no & "_MAX"), m_sPath, "10.0")
                            Case 1 ' リトレースのオフセットＹ
                                strMin = GetPrivateProfileString_S("CUT_RETRACE_PARA", (no & "_MIN"), m_sPath, "-10.0")
                                strMax = GetPrivateProfileString_S("CUT_RETRACE_PARA", (no & "_MAX"), m_sPath, "10.0")
                            Case 2 ' ストレートカット・リトレースのQレート
                                strMin = GetPrivateProfileString_S("CUT_RETRACE_PARA", (no & "_MIN"), m_sPath, "0.1")
                                strMax = GetPrivateProfileString_S("CUT_RETRACE_PARA", (no & "_MAX"), m_sPath, "40.0")
                            Case 3 ' ストレートカット・リトレースのトリム速度
                                strMin = GetPrivateProfileString_S("CUT_RETRACE_PARA", (no & "_MIN"), m_sPath, "0.1")
                                strMax = GetPrivateProfileString_S("CUT_RETRACE_PARA", (no & "_MAX"), m_sPath, "400.0")
                            Case 18 ' ストレートカット・リトレース本数'V2.0.0.0⑦  'V2.1.0.0① Case 12から Case 15へ変更 'V2.2.1.7① Case 15から Case 18へ変更
                                strMsg = GetPrivateProfileString_S("CUT_CUT", (no & "_MSG"), m_sPath, "??????")
                                strMin = GetPrivateProfileString_S("CUT_CUT", (no & "_MIN"), m_sPath, "1")
                                strMax = GetPrivateProfileString_S("CUT_CUT", (no & "_MAX"), m_sPath, "10")
                            Case Else
                                Throw New Exception("Case " & tag & ": Nothing")
                        End Select
                        'V2.0.0.0⑦↑

                        'V2.2.0.0② ↓
                    Case 5      ' Uカットパラメータ 
                        Select Case (tag)
                            Case 0 ' Ｌ１カット長
                                strMin = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MAX"), m_sPath, "20.0")
                            Case 1 ' Ｌ２カット長
                                strMin = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MAX"), m_sPath, "20.0")
                            Case 2 ' Ｒ１
                                strMin = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MAX"), m_sPath, "20.0")
                            Case 3 ' Ｒ２
                                strMin = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MAX"), m_sPath, "20.0")
                            Case 4 ' Ｑレート 
                                strMin = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MIN"), m_sPath, "0.1")
                                strMax = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MAX"), m_sPath, "40.0")
                            Case 5 ' 速度 
                                strMin = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MIN"), m_sPath, "0.1")
                                strMax = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MAX"), m_sPath, "400.0")
                            Case 6 ' Lターンポイント
                                strMin = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MIN"), m_sPath, "0.0")
                                strMax = GetPrivateProfileString_S("CUT_UCUT_PARA", (no & "_MAX"), m_sPath, "100.0")
                        End Select
                        'V2.2.0.0② ↑

                    Case Else
                                Throw New Exception("Parent.Tag - Case Else")
                        End Select

                        With cTextBox
                    Call .SetStrMsg(strMsg) ' ﾃﾞｰﾀﾁｪｯｸｴﾗｰ時の表示名設定
                    Call .SetMinMax(strMin, strMax) ' 下限値･上限値の設定
                    If (False = strFlg) Then                                                    '###1042①
                        Call .SetStrTip(strMin & "～" & strMax & "の範囲で指定して下さい") ' ﾂｰﾙﾁｯﾌﾟﾒｯｾｰｼﾞの設定
                    Else                                                                        '###1042①
                        Call .SetStrTip(strMin & "～" & strMax & "文字の範囲で指定して下さい")  '###1042①
                        .MaxLength = Integer.Parse(strMax)                                      '###1042① SetControlData()内の条件判断で使用する
                        .TextAlign = HorizontalAlignment.Left                                   '###1042①
                    End If                                                                      '###1042①
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
                Me.SuspendLayout()

                With m_MainEdit ' 抵抗数･ｶｯﾄ数とも0にはならない仕様のため不要と思われる
                    ' 抵抗数確認
                    If (.W_PLT.RCount < 1) Then
                        m_ResNo = 1
                    End If
                    ' ｶｯﾄ数確認
                    If (.W_REG(m_ResNo).intTNN < 1) Then
                        m_CutNo = 1
                    End If
                End With

                ' ------------------------
                If (SLP_VMES <> m_MainEdit.W_REG(m_ResNo).intSLP) AndAlso (SLP_RMES <> m_MainEdit.W_REG(m_ResNo).intSLP) Then
                    ' 抵抗のｽﾛｰﾌﾟが 5:電圧測定のみ, 6:抵抗測定のみ ではない場合
                    For Each ctl As Control In CGrp_0.Controls
                        ' ｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のｺﾝﾄﾛｰﾙを有効にする(個々の設定は以降でおこなう)
                        ctl.Enabled = True
                    Next
                End If
                ' ------------------------

                Call ChangedCutShape(m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCTYP) ' 関連ｺﾝﾄﾛｰﾙの表示･非表示を設定

                ' ｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽ設定
                Call SetCutData()

                ' ｲﾝﾃﾞｯｸｽｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽ設定
                Call SetIdxCutData()

                ' Ｌカットパラメータグループボックス内の設定
                Call SetLCutParamData()

                'リトレースカットパラメータグループボックス内の設定
                SetRetraceCutParamData()    'V2.0.0.0⑦

                'V2.2.0.0②↓
                ' Ｕカットパラメータの追加
                Call SetUCutParamData()
                'V2.2.0.0②↑

                ' 発振器種別がﾌｧｲﾊﾞﾚｰｻﾞの場合
                If (gSysPrm.stRAT.giOsc_Res = OSCILLATOR_FL) Then
                    ' FL加工条件ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ設定
                    Call SetFLCndData()
                End If

                ' ------------------------
                If UserModule.IsMeasureOnly(m_MainEdit.W_REG, m_ResNo) Then
                    ' 抵抗のｽﾛｰﾌﾟが 5:電圧測定のみ, 6:抵抗測定のみ の場合
                    For Each ctl As Control In CGrp_0.Controls
                        ' ｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のｺﾝﾄﾛｰﾙを無効にする
                        If (Not ctl Is CCmb_0) Then
                            ' 抵抗番号ｺﾝﾎﾞﾎﾞｯｸｽをいったん無効にするとﾌｫｰｶｽが戻らなくなる
                            ctl.Enabled = False
                        End If
                    Next
                    CLbl_0.Enabled = True       ' 抵抗数ﾗﾍﾞﾙ
                    CLblRN_0.Enabled = True     ' 抵抗数
                    CLbl_1.Enabled = True       ' 抵抗番号ﾗﾍﾞﾙ
                    CLbl_2.Enabled = True       ' ｶｯﾄ数ﾗﾍﾞﾙ
                    CLblCN_0.Enabled = True     ' ｶｯﾄ数
                    CLblCN_0.Text = 0           ' ｶｯﾄ数の表示を0にする
                End If
                ' ------------------------

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            Finally
                Me.ResumeLayout()
                Me.Refresh()
            End Try

        End Sub

#Region "ｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内の設定"
        ''' <summary>ｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽに値を設定する</summary>
        Private Sub SetCutData()
            Dim idx As Integer

            Try
                With m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo)
                    For i As Integer = 0 To (m_CtlCut.Length - 1) Step 1
                        Select Case (i)
                            Case 0 ' 総抵抗数, 抵抗番号
                                Dim rCnt As Integer = m_MainEdit.W_PLT.RCount
                                Dim cCombo As cCmb_ = DirectCast(m_CtlCut(i), cCmb_)
                                CLblRN_0.Text = rCnt.ToString() ' 総抵抗数
                                With cCombo ' 抵抗番号
                                    .Items.Clear()
                                    For j As Integer = 1 To rCnt Step 1
                                        '.Items.Add(String.Format("{0,5:#0}", j)) ' 総抵抗数分繰り返す
                                        .Items.Add(j.ToString(0) & ":" & m_MainEdit.W_REG(j).strRNO) ' 総抵抗数分繰り返す
                                    Next j
                                End With
                                Call NoEventIndexChange(cCombo, (m_ResNo - 1)) ' 指定抵抗番号を設定

                            Case 1 ' 総ｶｯﾄ数, ｶｯﾄ番号
                                Dim cCnt As Integer = m_MainEdit.W_REG(m_ResNo).intTNN
                                Dim cCombo As cCmb_ = DirectCast(m_CtlCut(i), cCmb_)
                                CLblCN_0.Text = cCnt.ToString() ' 総ｶｯﾄ数
                                With cCombo ' ｶｯﾄ番号
                                    .Items.Clear()
                                    For j As Integer = 1 To cCnt Step 1
                                        .Items.Add(String.Format("{0,5:#0}", j)) ' 総ｶｯﾄ数分繰り返す
                                    Next j
                                End With
                                Call NoEventIndexChange(cCombo, (m_CutNo - 1)) ' 指定ｶｯﾄ番号を設定

                            Case 2 ' ｶｯﾄ方法(1:ﾄﾗｯｷﾝｸﾞ, 2:ｲﾝﾃﾞｯｸｽ(STｶｯﾄのみ), 3:NGｶｯﾄ)
                                If UserModule.IsMarking(m_MainEdit.W_REG, m_ResNo) Then             ' マーキングの時
#If cFORCEcCUT Then
                                    m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCUT = CNS_CUTM_FC   ' カット方法を強制カットに固定する。
#End If
                                    CCmb_2.Enabled = False
                                Else
                                    CCmb_2.Enabled = True
                                End If

                                idx = GetComboBoxValue2Index(.intCUT, Me.m_lstCutMethod)

                                Call NoEventIndexChange(DirectCast(m_CtlCut(i), cCmb_), idx)

                                Call ChangedCutMethod(.intCUT)

                            Case 3 ' ｶｯﾄ形状(1:ｽﾄﾚｰﾄ, 2:Lｶｯﾄ)
                                If m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCTYP = CNS_CUTP_M Then
                                    m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCUT = CNS_CUTM_TR
                                End If
                                Call NoEventIndexChange(DirectCast(m_CtlCut(i), cCmb_), GetComboBoxValue2Index(.intCTYP, Me.m_lstCutType))
                                Call ChangedCutShape(m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCTYP) ' 関連ｺﾝﾄﾛｰﾙの表示･非表示を設定

                            Case 4 ' ｻｰﾍﾟﾝﾀｲﾝ本数
                                m_CtlCut(i).Text = (.intNum).ToString()
                            Case 5 ' Qﾚｰﾄ(0.1KHz→KHz)
                                m_CtlCut(i).Text = (.intQF1 / 10).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                            Case 6 ' 速度
                                m_CtlCut(i).Text = (.dblV1).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                            Case 7 ' ｶｯﾄ位置X
                                m_CtlCut(i).Text = (.dblSTX).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                            Case 8 ' ｶｯﾄ位置Y
                                m_CtlCut(i).Text = (.dblSTY).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                            Case 9 ' ｶｯﾄ位置2X
                                m_CtlCut(i).Text = (.dblSX2).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                            Case 10 ' ｶｯﾄ位置2Y
                                m_CtlCut(i).Text = (.dblSY2).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                            Case 11 ' ｶｯﾄ長1
                                m_CtlCut(i).Text = (.dblDL2).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                            Case 12 ' ｶｯﾄ長2
                                m_CtlCut(i).Text = (.dblDL3).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                            Case 13 ' ｶｯﾄ方向1, ｶｯﾄ方向2
                                Dim iWK As Integer
                                Dim index As Integer
                                iWK = .intANG
                                Select Case (iWK)
                                    Case 0                  'V1.0.4.3② ADD
                                        index = 0   ' 0°   'V1.0.4.3② ADD
                                    Case 90
                                        index = 1   ' 90°
                                    Case 180
                                        index = 2   ' 180°
                                    Case 270
                                        index = 3   ' 270°
                                    Case 10
                                        index = 4   ' 10°
                                    Case 20
                                        index = 5   ' 20°
                                    Case 30
                                        index = 6   ' 30°
                                    Case 40
                                        index = 7   ' 40°
                                    Case 50
                                        index = 8   ' 50°
                                    Case 60
                                        index = 9   ' 60°
                                    Case 70
                                        index = 10  ' 70°
                                    Case 80
                                        index = 11  ' 80°
                                    Case 100
                                        index = 12  ' 100°
                                    Case 110
                                        index = 13  ' 110°
                                    Case 120
                                        index = 14  ' 120°
                                    Case 130
                                        index = 15  ' 130°
                                    Case 140
                                        index = 16  ' 140°
                                    Case 150
                                        index = 17  ' 150°
                                    Case 160
                                        index = 18  ' 160°
                                    Case 170
                                        index = 19  ' 170°
                                    Case Else
                                        'V1.0.4.3②                                        index = 0   ' 0°
                                        index = Add_CCmb_4_Item(iWK)   'V1.0.4.3②
                                End Select
                                Call NoEventIndexChange(DirectCast(m_CtlCut(i), cCmb_), index)

                            Case 14 ' ｶｯﾄ方向2
                                Dim iWK As Integer
                                iWK = .intANG2
                            Case 15 ' ｶｯﾄｵﾌ
                                m_CtlCut(i).Text = (.dblCOF).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                            Case 16 ' Lﾀｰﾝﾎﾟｲﾝﾄ
                                m_CtlCut(i).Text = (.dblLTP).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                            Case 17 ' 測定機器(0=内部測定, 1以上外部測定機器番号)
                                ' GP-IB登録機器名を表示する(外部電源を除く)
                                Dim cCombo As cCmb_ = DirectCast(m_CtlCut(i), cCmb_)
                                Dim ctrg As String ' ﾄﾘｶﾞｰｺﾏﾝﾄﾞ
                                Dim type As Integer = .intMType
                                Dim cnt As Integer = 0 ' ﾘｽﾄに追加した項目数
                                idx = 0 ' 選択するｲﾝﾃﾞｯｸｽ
                                cCombo.Items.Clear()
                                cCombo.Items.Add(" 0:内部測定器")
                                With m_MainEdit
                                    If (0 < .W_PLT.GCount) Then ' GP-IB測定機器が登録されている場合
                                        For j As Integer = 1 To (.W_PLT.GCount) Step 1
                                            ctrg = .W_GPIB(j).strCTRG
                                            ' ﾄﾘｶﾞｰｺﾏﾝﾄﾞありの場合、外部測定器としてﾘｽﾄに追加
                                            If ("" <> ctrg) Then
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

                                    Else ' GP-IB測定機器の登録がない場合
                                        .W_REG(m_ResNo).STCUT(m_CutNo).intMType = 0
                                        idx = 0 ' 内部測定器
                                    End If
                                End With

                                ' ｶｯﾄ方法がNGｶｯﾄまたは外部測定器の場合測定ﾓｰﾄﾞを無効にする
                                If (CNS_CUTM_NG = .intCUT) OrElse (0 < idx) Then
                                    m_CtlCut(CUT_TMM).Enabled = False ' 測定ﾓｰﾄﾞ無効
                                Else
                                    m_CtlCut(CUT_TMM).Enabled = True ' 測定ﾓｰﾄﾞ有効
                                End If
                                If (CNS_CUTM_NG = .intCUT Or CNS_CUTM_TR = .intCUT) Then       ' トラッキングまたはNGカットの場合は、内部測定のみ
                                    .intMType = 0
                                    m_CtlCut(CUT_MTYPE).Enabled = False
                                End If
                                Call NoEventIndexChange(cCombo, idx)

                            Case 18 ' 測定ﾓｰﾄﾞ(0:高速, 1:高精度)
                                Call NoEventIndexChange(DirectCast(m_CtlCut(i), cCmb_), .intTMM)
                            Case 19 '###1042①
                                m_CtlCut(i).Text = .cFormat
                                'V2.1.0.0①↓

                                'V2.2.1.7① ↓
                            Case 20 ' 印字固定部
                                m_CtlCut(i).Text = .cMarkFix
                            Case 21 ' 開始番号
                                m_CtlCut(i).Text = .cMarkStartNum
                            Case 22 ' 重複回数
                                m_CtlCut(i).Text = .intMarkRepeatCnt

                            Case 23 ' リピート有無
                                Call NoEventIndexChange(DirectCast(m_CtlCut(i), cCmb_), .iVariationRepeat)
                            Case 24 ' 判定有無
                                Call NoEventIndexChange(DirectCast(m_CtlCut(i), cCmb_), .iVariation)
                            Case 25 ' 上昇率
                                m_CtlCut(i).Text = (.dRateOfUp).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                            Case 26 ' 下限値
                                m_CtlCut(i).Text = (.dVariationLow).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                            Case 27 ' 上限値
                                m_CtlCut(i).Text = (.dVariationHi).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                                'V2.1.0.0①↑
                            Case 28 ' リトレース本数'V2.0.0.0⑦ 'V2.1.0.0①Case 20からCase 25へ変更
                                m_CtlCut(i).Text = (.intRetraceCnt).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())   'V2.0.0.0⑦


                                'Case 20 ' リピート有無
                                '    Call NoEventIndexChange(DirectCast(m_CtlCut(i), cCmb_), .iVariationRepeat)
                                'Case 21 ' 判定有無
                                '    Call NoEventIndexChange(DirectCast(m_CtlCut(i), cCmb_), .iVariation)
                                'Case 22 ' 上昇率
                                '    m_CtlCut(i).Text = (.dRateOfUp).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                                'Case 23 ' 下限値
                                '    m_CtlCut(i).Text = (.dVariationLow).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                                'Case 24 ' 上限値
                                '    m_CtlCut(i).Text = (.dVariationHi).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                                '    'V2.1.0.0①↑
                                'Case 25 ' リトレース本数'V2.0.0.0⑦ 'V2.1.0.0①Case 20からCase 25へ変更
                                '    m_CtlCut(i).Text = (.intRetraceCnt).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())   'V2.0.0.0⑦

                                'V2.2.1.7①↑

                                'V1.0.4.3③ ADD ↓
                                'V2.0.0.0⑦                            Case 20 ' Qﾚｰﾄ(0.1KHz→KHz)
                                'V2.0.0.0⑦                                m_CtlCut(i).Text = (.intQF2 / 10).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                                'V2.0.0.0⑦                            Case 21 ' 速度
                                'V2.0.0.0⑦                                m_CtlCut(i).Text = (.dblV2).ToString(DirectCast(m_CtlCut(i), cTxt_).GetStrFormat())
                                'V1.0.4.3③ ADD ↑
                            Case Else
                                Throw New Exception("Case " & i & ": Nothing")
                        End Select
                    Next i
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

#Region "選択されたｶｯﾄ方法により関連するｺﾝﾄﾛｰﾙの有効･無効を変更する"
        ''' <summary>選択されたｶｯﾄ方法により関連するｺﾝﾄﾛｰﾙの有効･無効を変更する</summary>
        ''' <param name="intCUT">1:ﾄﾗｯｷﾝｸﾞ, 2:ｲﾝﾃﾞｯｸｽｶｯﾄ, 3:NGｶｯﾄ</param>
        Private Sub ChangedCutMethod(ByVal intCUT As Short)
            Try
                Select Case (intCUT)
                    Case CNS_CUTM_IX ' ｲﾝﾃﾞｯｸｽｶｯﾄの場合
                        CGrp_1.Enabled = True ' ｲﾝﾃﾞｯｸｽｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽを表示
                        m_NGCut.Enabled = True ' NGｶｯﾄでは使用しないｺﾝﾄﾛｰﾙを有効にする
                    Case CNS_CUTM_NG ' NGｶｯﾄの場合
                        CGrp_1.Enabled = False ' ｲﾝﾃﾞｯｸｽｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽを無効にする
                        m_NGCut.Enabled = False ' NGｶｯﾄでは使用しないｺﾝﾄﾛｰﾙを無効にする
                    Case Else ' ﾄﾗｯｷﾝｸﾞ
                        CGrp_1.Enabled = False ' ｲﾝﾃﾞｯｸｽｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽを無効にする
                        m_NGCut.Enabled = True ' NGｶｯﾄでは使用しないｺﾝﾄﾛｰﾙを有効にする
                End Select

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#Region "選択されたｶｯﾄ形状により関連するｺﾝﾄﾛｰﾙの表示･非表示を変更する"
        'V2.0.0.0        ''' <summary>選択されたｶｯﾄ形状により関連するｺﾝﾄﾛｰﾙの表示･非表示を変更する</summary>
        'V2.0.0.0        ''' <param name="selectedIdx">0:ｽﾄﾚｰﾄ, 1:Lｶｯﾄ, 2:ｻｰﾍﾟﾝﾀｲﾝ</param>
        'V2.0.0.0        Private Sub ChangedCutShape(ByVal selectedIdx As Integer)
        ''' <summary>
        ''' 選択されたｶｯﾄ形状により関連するｺﾝﾄﾛｰﾙの表示･非表示を変更する
        ''' </summary>
        ''' <param name="intCTYP">1:ｽﾄﾚｰﾄ, 2:Lｶｯﾄ, 3:ｻｰﾍﾟﾝﾀｲﾝ</param>
        ''' <remarks></remarks>
        Private Sub ChangedCutShape(ByVal intCTYP As Short)
            Try
                '###1042①↓
                Dim strMin As String = ""           ' 設定する変数の最大値
                Dim strMax As String = ""           ' 設定する変数の最小値
                Dim strMsg As String = ""           ' ｴﾗｰで表示する項目名

                strMsg = GetPrivateProfileString_S("CUT_CUT", ("007_MSG"), m_sPath, "??????")
                strMin = GetPrivateProfileString_S("CUT_CUT", ("007_MIN"), m_sPath, "0.001")
                strMax = GetPrivateProfileString_S("CUT_CUT", ("007_MAX"), m_sPath, "20.000")
                With DirectCast(m_CtlCut(CUT_LEN_1), cTxt_) ' 目標値ﾃｷｽﾄﾎﾞｯｸｽ
                    Call .SetStrMsg(strMsg) ' ﾃﾞｰﾀﾁｪｯｸｴﾗｰ時の表示名設定
                    Call .SetMinMax(strMin, strMax) ' 下限値･上限値の設定
                    Call .SetStrTip(strMin & "～" & strMax & "の範囲で指定して下さい") ' ﾂｰﾙﾁｯﾌﾟﾒｯｾｰｼﾞの設定
                End With
                '###1042①↑
                ' V1.1.0.0③ ＳＴリトレース、Ｌカット追加、文字マーキング追加　番号をDEFINE化
                ' 関連ｺﾝﾄﾛｰﾙの表示･非表示を設定
                m_CtlCut(CUT_QRATE).Enabled = True     'V1.0.4.3③Ｑレート
                m_CtlCut(CUT_SPEED).Enabled = True     'V1.0.4.3③速度
                'V2.0.0.0                Select Case (selectedIdx)
                Select Case (intCTYP) ' ｶｯﾄ形状
                    Case CNS_CUTP_ST ' ｽﾄﾚｰﾄ
                        'm_Serpentine.Visible = False
                        'V2.0.0.0                        m_CutCondition.Display = 1
                        CLbl_14.Visible = False ' Lﾀｰﾝﾎﾟｲﾝﾄ
                        m_CtlCut(CUT_LTP).Visible = False
                        'V2.0.0.0⑲                        CLbl_36.Visible = False   ' 文字        '###1042①
                        m_CtlCut(CUT_LETTER).Enabled = False    '###1042①
                        m_CtlCut(CUT_LEN_1).Enabled = True          'V1.0.4.3③カット長
                        m_CtlCut(CUT_DIR_1).Enabled = True          'V1.0.4.3③カット方向
                        m_CtlCut(CUT_OFF).Enabled = True            'V1.0.4.3③カットオフ
                        m_CtlCut(CUT_START_2_X).Enabled = False     'V1.0.4.3③オフセットX
                        m_CtlCut(CUT_START_2_Y).Enabled = False     'V1.0.4.3③オフセットY
                        'V2.0.0.0⑦                        m_CtlCut(CUT_TR_Q).Enabled = False          'V1.0.4.3③リトレースＱレート
                        'V2.0.0.0⑦                        m_CtlCut(CUT_TR_SPEED).Enabled = False      'V1.0.4.3③リトレース速度
                        If (m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCUT = CNS_CUTM_IX) Then        'V2.2.0.0⑩
                            CGrp_1.Enabled = True                       'インデックスカット'V2.1.0.0①
                        End If
                        CGrp_3.Enabled = False                      'V1.0.4.3③Ｌカットパラメータ
                        CGrp_4.Enabled = False                      ''V2.0.0.0⑦リトレースカットパラメータ
                        CGrp_5.Enabled = False                      'Ｕカットパラメータ      'V2.2.0.0② 
                        'V2.1.0.0①↓
                        m_CtlCut(CUT_VAR_REPEAT).Enabled = True     'リピート有無
                        m_CtlCut(CUT_VARIATION).Enabled = True      '判定有無
                        m_CtlCut(CUT_RATE).Enabled = True           '上昇率
                        m_CtlCut(CUT_VAR_LO).Enabled = True         '下限値
                        m_CtlCut(CUT_VAR_HI).Enabled = True         '上限値
                        'V2.1.0.0①↑

                        'V2.2.1.7①↓
                        CLbl_80.Visible = False
                        CLbl_81.Visible = False
                        CLbl_82.Visible = False
                        m_CtlCut(CUT_MARK_FIX).Visible = False
                        m_CtlCut(CUT_ST_NUM).Visible = False
                        m_CtlCut(CUT_REPEAT_CNT).Visible = False
                        'V2.2.1.7① ↑

                    Case CNS_CUTP_ST_TR ' ストレート・リトレース(RETRACE)カット
                        'm_Serpentine.Visible = False
                        'V2.0.0.0                        m_CutCondition.Display = 1
                        CLbl_14.Visible = False ' Lﾀｰﾝﾎﾟｲﾝﾄ
                        m_CtlCut(CUT_LTP).Visible = False
                        'V2.0.0.0⑲                        CLbl_36.Visible = False   ' 文字        '###1042①
                        m_CtlCut(CUT_LETTER).Enabled = False    '###1042①
                        m_CtlCut(CUT_LEN_1).Enabled = True          'V1.0.4.3③カット長
                        m_CtlCut(CUT_DIR_1).Enabled = True          'V1.0.4.3③カット方向
                        m_CtlCut(CUT_OFF).Enabled = True            'V1.0.4.3③カットオフ
                        'V2.0.0.0⑦                        CLbl_10.Visible = True                      'V1.0.4.3③オフセット
                        m_CtlCut(CUT_START_2_X).Enabled = True      'V1.0.4.3③オフセットX
                        m_CtlCut(CUT_START_2_Y).Enabled = True      'V1.0.4.3③オフセットY
                        'V2.0.0.0⑦                        m_CtlCut(CUT_TR_Q).Enabled = True           'V1.0.4.3③リトレースのＱレート
                        'V2.0.0.0⑦                        m_CtlCut(CUT_TR_SPEED).Enabled = True       'V1.0.4.3③リトレースの速度
                        If (m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCUT = CNS_CUTM_IX) Then        'V2.2.0.0⑩
                            CGrp_1.Enabled = True                       'インデックスカット'V2.1.0.0①
                        End If
                        CGrp_3.Enabled = False                      'V1.0.4.3③Ｌカットパラメータ
                        CGrp_4.Enabled = True                      ''V2.0.0.0⑦リトレースカットパラメータ
                        CGrp_5.Enabled = False                      'Ｕカットパラメータ      'V2.2.0.0② 
                        'V2.1.0.0①↓
                        m_CtlCut(CUT_VAR_REPEAT).Enabled = True     'リピート有無
                        m_CtlCut(CUT_VARIATION).Enabled = True      '判定有無
                        m_CtlCut(CUT_RATE).Enabled = True           '上昇率
                        m_CtlCut(CUT_VAR_LO).Enabled = True         '下限値
                        m_CtlCut(CUT_VAR_HI).Enabled = True         '上限値
                        'V2.1.0.0①↑

                        'V2.2.1.7① ↓
                        CLbl_80.Visible = False
                        CLbl_81.Visible = False
                        CLbl_82.Visible = False
                        m_CtlCut(CUT_MARK_FIX).Visible = False
                        m_CtlCut(CUT_ST_NUM).Visible = False
                        m_CtlCut(CUT_REPEAT_CNT).Visible = False
                        'V2.2.1.7① ↑

                    Case CNS_CUTP_L ' Lｶｯﾄ
                        'm_Serpentine.Visible = False
                        'V2.0.0.0                        m_CutCondition.Display = 2
                        'V1.0.4.3③                        CLbl_14.Visible = True ' Lﾀｰﾝﾎﾟｲﾝﾄ
                        'V1.0.4.3③                        m_CtlCut(CUT_LTP).Visible = True
                        CLbl_14.Visible = False ' Lﾀｰﾝﾎﾟｲﾝﾄ         'V1.0.4.3③
                        m_CtlCut(CUT_QRATE).Enabled = False         'V1.0.4.3③Ｑレート
                        m_CtlCut(CUT_SPEED).Enabled = False         'V1.0.4.3③速度
                        m_CtlCut(CUT_LTP).Visible = False           'V1.0.4.3③ターンポイント
                        m_CtlCut(CUT_LEN_1).Enabled = False         'V1.0.4.3③カット長
                        m_CtlCut(CUT_DIR_1).Enabled = False         'V1.0.4.3③カット方向
                        'V2.0.0.0⑲                        CLbl_36.Visible = False   ' 文字        '###1042①
                        m_CtlCut(CUT_LETTER).Enabled = False    '###1042①
                        m_CtlCut(CUT_OFF).Enabled = True            'V1.0.4.3③カットオフ
                        m_CtlCut(CUT_START_2_X).Enabled = False     'V1.0.4.3③オフセットX
                        m_CtlCut(CUT_START_2_Y).Enabled = False     'V1.0.4.3③オフセットY
                        'V2.0.0.0⑦                        m_CtlCut(CUT_TR_Q).Enabled = False          'V1.0.4.3③リトレースＱレート
                        'V2.0.0.0⑦                        m_CtlCut(CUT_TR_SPEED).Enabled = False      'V1.0.4.3③リトレース速度
                        If (m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCUT = CNS_CUTM_IX) Then        'V2.2.0.0⑩
                            CGrp_1.Enabled = True                       'インデックスカット'V2.1.0.0①
                        End If
                        CGrp_3.Enabled = True                       'V1.0.4.3③Ｌカットパラメータ
                        CGrp_4.Enabled = False                      ''V2.0.0.0⑦リトレースカットパラメータ
                        CGrp_5.Enabled = False                      'Ｕカットパラメータ      'V2.2.0.0② 
                        'V2.1.0.0①↓
                        m_CtlCut(CUT_VAR_REPEAT).Enabled = True     'リピート有無
                        m_CtlCut(CUT_VARIATION).Enabled = True      '判定有無
                        m_CtlCut(CUT_RATE).Enabled = True           '上昇率
                        m_CtlCut(CUT_VAR_LO).Enabled = True         '下限値
                        m_CtlCut(CUT_VAR_HI).Enabled = True         '上限値
                        'V2.1.0.0①↑

                        'V2.2.1.7① ↓
                        CLbl_80.Visible = False
                        CLbl_81.Visible = False
                        CLbl_82.Visible = False
                        m_CtlCut(CUT_MARK_FIX).Visible = False
                        m_CtlCut(CUT_ST_NUM).Visible = False
                        m_CtlCut(CUT_REPEAT_CNT).Visible = False
                        'V2.2.1.7① ↑

                    Case CNS_CUTP_M                             '###1042①
                        ''m_Serpentine.Visible = False
                        'V2.0.0.0                        m_CutCondition.Display = 2
                        CLbl_14.Visible = False ' Lﾀｰﾝﾎﾟｲﾝﾄ
                        m_CtlCut(CUT_LTP).Visible = False
                        CLbl_36.Visible = True                      ' 文字         '###1042①
                        m_CtlCut(CUT_LETTER).Enabled = True         '###1042①文字列
                        m_CtlCut(CUT_LEN_1).Enabled = True          'V1.0.4.3③カット長
                        m_CtlCut(CUT_DIR_1).Enabled = True          'V1.0.4.3③カット方向
                        m_CtlCut(CUT_DIR_1).Enabled = True          'V1.0.4.3③カット方向
                        m_CtlCut(CUT_OFF).Enabled = False           'V1.0.4.3③カットオフ
                        m_CtlCut(CUT_START_2_X).Enabled = False     'V1.0.4.3③オフセットX
                        m_CtlCut(CUT_START_2_Y).Enabled = False     'V1.0.4.3③オフセットY
                        'V2.0.0.0⑦                        m_CtlCut(CUT_TR_Q).Enabled = False          'V1.0.4.3③リトレースＱレート
                        'V2.0.0.0⑦                        m_CtlCut(CUT_TR_SPEED).Enabled = False      'V1.0.4.3③リトレース速度
                        CGrp_1.Enabled = False                       'インデックスカット'V2.1.0.0①
                        CGrp_4.Enabled = False                      ''V2.0.0.0⑦リトレースカットパラメータ
                        CGrp_3.Enabled = False                      'V1.0.4.3③Ｌカットパラメータ
                        CGrp_5.Enabled = False                      'Ｕカットパラメータ      'V2.2.0.0② 
                        'V2.1.0.0①↓
                        m_CtlCut(CUT_VAR_REPEAT).Enabled = False    'リピート有無
                        m_CtlCut(CUT_VARIATION).Enabled = False     '判定有無
                        m_CtlCut(CUT_RATE).Enabled = False          '上昇率
                        m_CtlCut(CUT_VAR_LO).Enabled = False        '下限値
                        m_CtlCut(CUT_VAR_HI).Enabled = False        '上限値
                        'V2.1.0.0①↑

                        '###1042①↓
                        With DirectCast(m_CtlCut(CUT_LEN_1), cTxt_) ' 目標値ﾃｷｽﾄﾎﾞｯｸｽ
                            strMin = "0.1"
                            strMax = "10.0"
                            strMsg = "文字高さ"
                            Call .SetStrMsg(strMsg) ' ﾃﾞｰﾀﾁｪｯｸｴﾗｰ時の表示名設定
                            Call .SetMinMax(strMin, strMax) ' 下限値･上限値の設定
                            Call .SetStrTip(strMin & "～" & strMax & "の範囲で指定して下さい") ' ﾂｰﾙﾁｯﾌﾟﾒｯｾｰｼﾞの設定
                        End With
                        '###1042①↑

                        'V2.2.1.7① ↓
                        If ((m_MainEdit.W_stUserData.iTrimType = 5) And (m_MainEdit.W_REG(m_ResNo).intSLP = SLP_MARK)) Then
                            CLbl_80.Visible = True
                            CLbl_81.Visible = True
                            CLbl_82.Visible = True
                            m_CtlCut(CUT_MARK_FIX).Visible = True
                            m_CtlCut(CUT_ST_NUM).Visible = True
                            m_CtlCut(CUT_REPEAT_CNT).Visible = True

                            CLbl_36.Visible = False
                            m_CtlCut(CUT_LETTER).Visible = False                      ' 文字
                        Else
                            CLbl_80.Visible = False
                            CLbl_81.Visible = False
                            CLbl_82.Visible = False

                            m_CtlCut(CUT_MARK_FIX).Visible = False
                            m_CtlCut(CUT_ST_NUM).Visible = False
                            m_CtlCut(CUT_REPEAT_CNT).Visible = False

                            CLbl_36.Visible = True
                            m_CtlCut(CUT_LETTER).Visible = True                      ' 文字
                        End If
                        'V2.2.1.7① ↑

                        'V2.2.0.0② ↓
                    Case CNS_CUTP_U  ' Uｶｯﾄ
                        CLbl_14.Visible = False ' Lﾀｰﾝﾎﾟｲﾝﾄ         '
                        m_CtlCut(CUT_QRATE).Enabled = False         'Ｑレート
                        m_CtlCut(CUT_SPEED).Enabled = False         '速度
                        m_CtlCut(CUT_LTP).Visible = False           'ターンポイント
                        m_CtlCut(CUT_LEN_1).Enabled = False         'カット長
                        m_CtlCut(CUT_DIR_1).Enabled = False         'カット方向
                        m_CtlCut(CUT_LETTER).Enabled = False
                        m_CtlCut(CUT_OFF).Enabled = True            'カットオフ
                        m_CtlCut(CUT_START_2_X).Enabled = False     'オフセットX
                        m_CtlCut(CUT_START_2_Y).Enabled = False     'オフセットY
                        If (m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCUT = CNS_CUTM_IX) Then        'V2.2.0.0⑩
                            CGrp_1.Enabled = True                       'インデックスカット
                        End If
                        CGrp_3.Enabled = False                      'Ｌカットパラメータ
                        CGrp_4.Enabled = False                      'リトレースカットパラメータ
                        CGrp_5.Enabled = True                       'Ｕカットパラメータ      'V2.2.0.0② 
                        m_CtlCut(CUT_VAR_REPEAT).Enabled = True     'リピート有無
                        m_CtlCut(CUT_VARIATION).Enabled = True      '判定有無
                        m_CtlCut(CUT_RATE).Enabled = True           '上昇率
                        m_CtlCut(CUT_VAR_LO).Enabled = True         '下限値
                        m_CtlCut(CUT_VAR_HI).Enabled = True         '上限値
                        'V2.2.0.0② ↑

                        'V2.2.1.7① ↓
                        CLbl_80.Visible = False
                        CLbl_81.Visible = False
                        CLbl_82.Visible = False
                        m_CtlCut(CUT_MARK_FIX).Visible = False
                        m_CtlCut(CUT_ST_NUM).Visible = False
                        m_CtlCut(CUT_REPEAT_CNT).Visible = False
                        'V2.2.1.7① ↑

                        'Case 2 ' ｻｰﾍﾟﾝﾀｲﾝ
                        '    m_Serpentine.Visible = True
                        '    m_CutCondition.Display = 2
                        '    CLbl_14.Visible = False ' Lﾀｰﾝﾎﾟｲﾝﾄ
                        '    m_CtlCut(CUT_LTP).Visible = False
                        '    CLbl_36.Visible = False   ' 文字        '###1042①
                        '    m_CtlCut(CUT_LETTER).Visible = True     '###1042①
                    Case Else
                        'm_Serpentine.Visible = False
                        'V2.0.0.0                        m_CutCondition.Display = 1
                        CLbl_14.Visible = False ' Lﾀｰﾝﾎﾟｲﾝﾄ
                        m_CtlCut(CUT_LETTER).Enabled = False    '###1042①
                        m_CtlCut(CUT_LEN_1).Enabled = True          'V1.0.4.3③カット長
                        m_CtlCut(CUT_DIR_1).Enabled = True          'V1.0.4.3③カット方向
                        m_CtlCut(CUT_LTP).Visible = False
                        m_CtlCut(CUT_OFF).Enabled = True            'V1.0.4.3③カットオフ
                        m_CtlCut(CUT_START_2_X).Enabled = False     'V1.0.4.3③オフセットX
                        m_CtlCut(CUT_START_2_Y).Enabled = False     'V1.0.4.3③オフセットY
                        'V2.0.0.0⑦                        m_CtlCut(CUT_TR_Q).Enabled = False          'V1.0.4.3③リトレースＱレート
                        'V2.0.0.0⑦                        m_CtlCut(CUT_TR_SPEED).Enabled = False      'V1.0.4.3③リトレース速度
                        If (m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCUT = CNS_CUTM_IX) Then        'V2.2.0.0⑩
                            CGrp_1.Enabled = True                       'インデックスカット'V2.1.0.0①
                        End If
                        CGrp_3.Enabled = False                      'V1.0.4.3③Ｌカットパラメータ
                        CGrp_4.Enabled = False                      ''V2.0.0.0⑦リトレースカットパラメータ
                        CGrp_5.Enabled = False                      'Ｕカットパラメータ      'V2.2.0.0② 
                        '    カット数０の時の為                    Throw New Exception("Case " & selectedIdx & ": Nothing")

                        'V2.2.1.7① ↓
                        CLbl_80.Visible = False
                        CLbl_81.Visible = False
                        CLbl_82.Visible = False
                        m_CtlCut(CUT_MARK_FIX).Visible = False
                        m_CtlCut(CUT_ST_NUM).Visible = False
                        m_CtlCut(CUT_REPEAT_CNT).Visible = False
                        'V2.2.1.7① ↑

                End Select

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#End Region

#Region "ｲﾝﾃﾞｯｸｽｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内の設定"
        ''' <summary>ｲﾝﾃﾞｯｸｽｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽに値を設定する</summary>
        Private Sub SetIdxCutData()
            Try
                With m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo)
                    For i As Integer = 0 To (m_CtlIdxCut.GetLength(0) - 1) Step 1
                        For j As Integer = 0 To (m_CtlIdxCut.GetLength(1) - 1) Step 1
                            Select Case (j)
                                Case 0 ' ｶｯﾄ回数1-5
                                    m_CtlIdxCut(i, j).Text = (.intIXN(i + 1)).ToString()
                                Case 1 ' ｶｯﾄ長1-5
                                    m_CtlIdxCut(i, j).Text = (.dblDL1(i + 1)).ToString(DirectCast(m_CtlIdxCut(i, j), cTxt_).GetStrFormat())
                                Case 2 ' ﾋﾟｯﾁ間ﾎﾟｰｽﾞ1-5(ms)
                                    m_CtlIdxCut(i, j).Text = (.lngPAU(i + 1)).ToString()
                                Case 3 ' 誤差1-5(%)
                                    m_CtlIdxCut(i, j).Text = (.dblDEV(i + 1)).ToString(DirectCast(m_CtlIdxCut(i, j), cTxt_).GetStrFormat())
                                Case 4 ' 測定機器(0:内部測定器, 1以上は外部測定器番号)
                                    Dim cCombo As cCmb_ = DirectCast(m_CtlIdxCut(i, j), cCmb_)
                                    Dim ctrg As String ' ﾄﾘｶﾞｰｺﾏﾝﾄﾞ
                                    Dim type As Integer = Convert.ToInt32(.intIXMType(i + 1))
                                    Dim cnt As Integer = 0 ' ﾘｽﾄに追加した項目数
                                    Dim idx As Integer = 0 ' 選択するｲﾝﾃﾞｯｸｽ
                                    Dim ctlIdx As Integer = GetCtlIdx(cCombo, DirectCast(cCombo.Tag, Integer)) ' 1次元目のｲﾝﾃﾞｯｸｽ
                                    cCombo.Items.Clear()
                                    cCombo.Items.Add(" 0:内部測定器")
                                    With m_MainEdit
                                        If (0 < .W_PLT.GCount) Then ' GP-IB測定機器が登録されている場合
                                            For k As Integer = 1 To (.W_PLT.GCount) Step 1
                                                ctrg = .W_GPIB(k).strCTRG
                                                ' ﾄﾘｶﾞｰｺﾏﾝﾄﾞありの場合、外部測定器としてﾘｽﾄに追加
                                                If ("" <> ctrg) Then
                                                    If (Not .W_GPIB(k).strGNAM Is Nothing) Then
                                                        cCombo.Items.Add(String.Format("{0,2:#0}", k) & ":" & .W_GPIB(k).strGNAM)
                                                    Else
                                                        cCombo.Items.Add(String.Format("{0,2:#0}", k) & ":")
                                                    End If
                                                    ' 追加したﾘｽﾄをｶｳﾝﾄｱｯﾌﾟ
                                                    cnt = (cnt + 1)
                                                    ' .intType(GP-IB登録番号)と同じ項目がﾘｽﾄに追加された場合に
                                                    ' その項目を選択するためｲﾝﾃﾞｯｸｽを設定する
                                                    ' 使用中の機器が削除された場合、GP-IBﾀﾌﾞ内の処理で
                                                    ' .intIXMTypeが0となるため内部測定器が選択される
                                                    If (type = k) Then idx = cnt
                                                End If
                                            Next k

                                            If (0 < idx) Then
                                                m_CtlIdxCut(ctlIdx, IDX_TMM).Enabled = False ' 測定ﾓｰﾄﾞ無効
                                            Else
                                                m_CtlIdxCut(ctlIdx, IDX_TMM).Enabled = True ' 測定ﾓｰﾄﾞ有効
                                            End If

                                        Else ' GP-IB測定機器の登録がない場合
                                            .W_REG(m_ResNo).STCUT(m_CutNo).intIXMType(i + 1) = 0
                                            idx = 0 ' 内部測定器
                                            m_CtlIdxCut(ctlIdx, IDX_TMM).Enabled = True ' 測定ﾓｰﾄﾞ有効
                                        End If
                                    End With
                                    Call NoEventIndexChange(cCombo, idx)

                                Case 5 ' 測定ﾓｰﾄﾞ(0:高速, 1:高精度)
                                    Dim cCombo As cCmb_ = DirectCast(m_CtlIdxCut(i, j), cCmb_)
                                    With cCombo
                                        .Items.Clear()
                                        .Items.Add("高速")
                                        .Items.Add("高精度")
                                    End With
                                    Call NoEventIndexChange(cCombo, Convert.ToInt32(.intIXTMM(i + 1)))

                                Case Else
                                    Throw New Exception("i = " & i & ", Case Else")
                            End Select
                        Next j
                    Next i
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#Region "FL加工条件ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内の値を設定"
        ''' <summary>FL加工条件ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽに値を設定する</summary>
        Private Sub SetFLCndData()
            Try
#If cOSCILLATORcFLcUSE Then
                With m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo)
                    For i As Integer = 0 To (m_CtlFLCnd.GetLength(0) - 1) Step 1
                        For j As Integer = 0 To (m_CtlFLCnd.GetLength(1) - 1) Step 1
                            Select Case (j)
                                Case 0 ' FL設定No.
                                    m_CtlFLCnd(i, j).Text = (.intCND(j + 1)).ToString("##0")
                                    Call NoEventIndexChange(DirectCast(m_CtlFLCnd(i, j), cCmb_), _
                                                                        Convert.ToInt32(.intCND(j + 1)))
                                Case 1 ' 電流値
                                    m_CtlFLCnd(i, j).Text = (stCND.Curr(.intCND(j + 1))).ToString("###0")
                                Case 2 ' Qﾚｰﾄ
                                    m_CtlFLCnd(i, j).Text = (stCND.Freq(.intCND(j + 1))).ToString("#0.0")
                                Case 3 ' STEG本数
                                    m_CtlFLCnd(i, j).Text = (stCND.Steg(.intCND(j + 1))).ToString("#0")
                                Case Else
                                    Throw New Exception("i = " & i & ", Case " & j & ": Nothing")
                            End Select
                        Next j
                    Next i
                End With
#End If

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#Region "Ｌカットパラメータグループボックス内の設定"
        ''' <summary>Ｌカットパラメータグループボックス内のテキストボックス・コンボボックスに値を設定する</summary>
        Private Sub SetLCutParamData()
            Try
                With m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo)
                    For i As Integer = 0 To (m_CtlLCut.GetLength(0) - 1) Step 1
                        For j As Integer = 0 To (m_CtlLCut.GetLength(1) - 1) Step 1
                            Select Case (j)
                                Case 0 ' カット長
                                    m_CtlLCut(i, j).Text = (.dCutLen(i + 1)).ToString(DirectCast(m_CtlLCut(i, j), cTxt_).GetStrFormat())
                                Case 1 ' Ｑレート
                                    m_CtlLCut(i, j).Text = (.dQRate(i + 1) / 10.0).ToString(DirectCast(m_CtlLCut(i, j), cTxt_).GetStrFormat())
                                Case 2 ' 速度
                                    m_CtlLCut(i, j).Text = (.dSpeed(i + 1)).ToString(DirectCast(m_CtlLCut(i, j), cTxt_).GetStrFormat())
                                Case 3 ' 角度
                                    Dim index As Integer
                                    Select Case (.dAngle(i + 1))
                                        Case 0
                                            index = 0   ' 0°
                                        Case 90
                                            index = 1   ' 90°
                                        Case 180
                                            index = 2   ' 180°
                                        Case 270
                                            index = 3   ' 270°
                                        Case 10
                                            index = 4   ' 10°
                                        Case 20
                                            index = 5   ' 20°
                                        Case 30
                                            index = 6   ' 30°
                                        Case 40
                                            index = 7   ' 40°
                                        Case 50
                                            index = 8   ' 50°
                                        Case 60
                                            index = 9   ' 60°
                                        Case 70
                                            index = 10  ' 70°
                                        Case 80
                                            index = 11  ' 80°
                                        Case 100
                                            index = 12  ' 100°
                                        Case 110
                                            index = 13  ' 110°
                                        Case 120
                                            index = 14  ' 120°
                                        Case 130
                                            index = 15  ' 130°
                                        Case 140
                                            index = 16  ' 140°
                                        Case 150
                                            index = 17  ' 150°
                                        Case 160
                                            index = 18  ' 160°
                                        Case 170
                                            index = 19  ' 170°
                                        Case Else
                                            index = Add_CCmb_Dir_X_Item(.dAngle(i + 1), DirectCast(m_CtlLCut(i, j), cCmb_))
                                    End Select
                                    Call NoEventIndexChange(DirectCast(m_CtlLCut(i, j), cCmb_), index)
                                Case 4 ' ターンポイント
                                    m_CtlLCut(i, j).Text = (.dTurnPoint(i + 1)).ToString(DirectCast(m_CtlLCut(i, j), cTxt_).GetStrFormat())
                                Case Else
                                    Throw New Exception("i = " & i & ", Case Else")
                            End Select
                        Next j
                    Next i
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

        'V2.0.0.0⑦ ADD ↓
#Region "リトレースカットパラメータグループボックス内の設定"
        ''' <summary>リトレースカットパラメータグループボックス内のテキストボックス・コンボボックスに値を設定する</summary>
        Private Sub SetRetraceCutParamData()
            Try
                With m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo)
                    For i As Integer = 0 To (m_CtlRetraceCut.GetLength(0) - 1) Step 1
                        For j As Integer = 0 To (m_CtlRetraceCut.GetLength(1) - 1) Step 1
                            Select Case (j)
                                Case 0 ' リトレースのオフセットＸ
                                    m_CtlRetraceCut(i, j).Text = (.dblRetraceOffX(i + 1)).ToString(DirectCast(m_CtlRetraceCut(i, j), cTxt_).GetStrFormat())
                                Case 1 ' リトレースのオフセットＹ
                                    m_CtlRetraceCut(i, j).Text = (.dblRetraceOffY(i + 1)).ToString(DirectCast(m_CtlRetraceCut(i, j), cTxt_).GetStrFormat())
                                Case 2 ' ストレートカット・リトレースのQレート(0.1KHz)に使用
                                    m_CtlRetraceCut(i, j).Text = (.dblRetraceQrate(i + 1) / 10.0).ToString(DirectCast(m_CtlRetraceCut(i, j), cTxt_).GetStrFormat())
                                Case 3 ' 速度
                                    m_CtlRetraceCut(i, j).Text = (.dblRetraceSpeed(i + 1)).ToString(DirectCast(m_CtlRetraceCut(i, j), cTxt_).GetStrFormat())
                                Case Else
                                    Throw New Exception("i = " & i & ", Case Else")
                            End Select
                        Next j
                    Next i
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region
        'V2.0.0.0⑦ ADD ↑
#End Region

#Region "すべてのﾃｷｽﾄﾎﾞｯｸｽのﾃﾞｰﾀﾁｪｯｸをおこなう"
#If cCUTDATAcCHECKcBYDATA Then
        ''' <summary>すべてのﾃｷｽﾄﾎﾞｯｸｽのﾃﾞｰﾀﾁｪｯｸをおこなう</summary>
        ''' <returns>0=正常, 1=ｴﾗｰ</returns>
        Friend Overrides Function CheckAllTextData() As Integer
            Dim ret As Integer
            Try
                ret = CutDataCheckByDataOnly()
            Catch ex As Exception
                Call MsgBox_Exception(ex)
                ret = 1
            Finally
                m_CheckFlg = False ' ﾁｪｯｸ終了
                CheckAllTextData = ret
            End Try

        End Function
#Region "データのみのチェック処理"
        ''' <summary>
        ''' カットデータデータのみのチェック
        ''' </summary>
        ''' <returns>0=正常, 1=エラー</returns>
        ''' <remarks></remarks>
        Private Function CutDataCheckByDataOnly() As Integer
            Dim strMSG As String = "", strMin As String = "", strMax As String = ""

            CutDataCheckByDataOnly = 0

            ' ｻｰﾍﾟﾝﾀｲﾝ本数
            Dim intNumMin As Short = Short.Parse(GetPrivateProfileString_S("CUT_CUT", "000_MIN", m_sPath, "1"))
            Dim intNumMax As Short = Short.Parse(GetPrivateProfileString_S("CUT_CUT", "000_MAX", m_sPath, "10"))
            ' Qﾚｰﾄ
            Dim intQF1Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "001_MIN", m_sPath, "0.1"))
            Dim intQF1Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "001_MAX", m_sPath, "40.0"))
            ' 速度
            Dim dblV1Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "002_MIN", m_sPath, "0.1"))
            Dim dblV1Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "002_MAX", m_sPath, "400.0"))
            ' ｶｯﾄ位置X
            Dim dblSTXMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "003_MIN", m_sPath, "-80.0"))
            Dim dblSTXMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "003_MAX", m_sPath, "80.0"))
            ' ｶｯﾄ位置Y
            Dim dblSTYMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "004_MIN", m_sPath, "-80.0"))
            Dim dblSTYMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "004_MAX", m_sPath, "80.0"))
            ' ｶｯﾄ位置2X
            Dim dblSX2Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "005_MIN", m_sPath, "-80.0"))
            Dim dblSX2Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "005_MAX", m_sPath, "80.0"))
            ' ｶｯﾄ位置2Y
            Dim dblSY2Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "006_MIN", m_sPath, "-80.0"))
            Dim dblSY2Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "006_MAX", m_sPath, "80.0"))
            ' ｶｯﾄ長1
            Dim dblDL2Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "007_MIN", m_sPath, "0.001"))
            Dim dblDL2Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "007_MAX", m_sPath, "20.000"))
            ' ｶｯﾄ長2
            Dim dblDL3Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "008_MIN", m_sPath, "0.001"))
            Dim dblDL3Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "008_MAX", m_sPath, "20.000"))
            ' ｶｯﾄｵﾌ
            Dim dblCOFMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "009_MIN", m_sPath, "-99.99"))
            Dim dblCOFMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "009_MAX", m_sPath, "99.99"))
            ' Lﾀｰﾝﾎﾟｲﾝﾄ
            Dim dblLTPMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "010_MIN", m_sPath, "0.0"))
            Dim dblLTPMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "010_MAX", m_sPath, "100.0"))

            ' ｶｯﾄ回数
            Dim intIXNMin As Short = Short.Parse(GetPrivateProfileString_S("CUT_INDEX", "000_MIN", m_sPath, "0"))
            Dim intIXNMax As Short = Short.Parse(GetPrivateProfileString_S("CUT_INDEX", "000_MAX", m_sPath, "999"))
            ' ｶｯﾄ長(ｲﾝﾃﾞｯｸｽﾋﾟｯﾁ)
            Dim dblDL1Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_INDEX", "001_MIN", m_sPath, "0.000"))
            Dim dblDL1Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_INDEX", "001_MAX", m_sPath, "20.000"))
            ' ﾋﾟｯﾁ間ﾎﾟｰｽﾞ
            Dim lngPAUMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_INDEX", "002_MIN", m_sPath, "0"))
            Dim lngPAUMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_INDEX", "002_MAX", m_sPath, "32767"))
            ' 誤差
            Dim dblDEVMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_INDEX", "003_MIN", m_sPath, "0.00"))
            Dim dblDEVMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_INDEX", "003_MAX", m_sPath, "99.99"))


            '' 電流値
            'Dim CurrMin As Integer = Integer.Parse(GetPrivateProfileString_S("CUT_FLCOND", "000_MIN", m_sPath, "1"))
            'Dim CurrMax As Integer = Integer.Parse(GetPrivateProfileString_S("CUT_FLCOND", "000_MAX", m_sPath, "1000"))
            '' Qﾚｰﾄ
            'Dim FreqMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_FLCOND", "001_MIN", m_sPath, "0.1"))
            'Dim FreqMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_FLCOND", "001_MAX", m_sPath, "40.0"))
            '' STEG本数
            'Dim StegMin As Integer = Integer.Parse(GetPrivateProfileString_S("CUT_FLCOND", "002_MIN", m_sPath, "1"))
            'Dim StegMax As Integer = Integer.Parse(GetPrivateProfileString_S("CUT_FLCOND", "002_MAX", m_sPath, "15"))
            '' 目標パワー（Ｗ）
            'Dim dblPowerAdjustTargetMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_FLCOND", "003_MIN", m_sPath, "0.01"))
            'Dim dblPowerAdjustTargetMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_FLCOND", "003_MAX", m_sPath, "20.0"))
            '' 許容範囲（±Ｗ）
            'Dim dblPowerAdjustToleLevelMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_FLCOND", "004_MIN", m_sPath, "0.01"))
            'Dim dblPowerAdjustToleLevelMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_FLCOND", "004_MAX", m_sPath, "10.0"))


            ' Qﾚｰﾄ
            'V2.1.0.0①            Dim intQF2Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "012_MIN", m_sPath, "0.1"))
            'V2.1.0.0①            Dim intQF2Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "012_MAX", m_sPath, "40.0"))
            ' 速度
            'V2.1.0.0①            Dim dblV2Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "013_MIN", m_sPath, "0.1"))
            'V2.1.0.0①            Dim dblV2Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_CUT", "013_MAX", m_sPath, "400.0"))
            ' 文字
            Dim intLetterLenMin As Integer = Integer.Parse(GetPrivateProfileString_S("CUT_CUT", "011_MIN", m_sPath, "1"))
            Dim intLetterLenMax As Integer = Integer.Parse(GetPrivateProfileString_S("CUT_CUT", "011_MAX", m_sPath, "10"))

            ' ６点ターンポイントＬカットトリミング
            ' カット長
            Dim dCutLenMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_LCUT_PARA", ("000_MIN"), m_sPath, "0.001"))
            Dim dCutLenMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_LCUT_PARA", ("000_MAX"), m_sPath, "20.000"))
            ' Qレート
            Dim dQRateMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_LCUT_PARA", ("001_MIN"), m_sPath, "0.1"))
            Dim dQRateMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_LCUT_PARA", ("001_MAX"), m_sPath, "40.0"))
            ' 速度
            Dim dSpeedMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_LCUT_PARA", ("002_MIN"), m_sPath, "0.1"))
            Dim dSpeedMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_LCUT_PARA", ("002_MAX"), m_sPath, "400.0"))
            ' Lﾀｰﾝﾎﾟｲﾝﾄ
            Dim dTurnPointMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_LCUT_PARA", ("003_MIN"), m_sPath, "0.0"))
            Dim dTurnPointMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_LCUT_PARA", ("003_MAX"), m_sPath, "100.0"))



            ' ストレートカット本数 'V2.1.0.0① カット毎の抵抗値変化量判定機能項目追加　３項目シフトして012_から015_へ変更
            ' ストレートカット本数 'V2.2.1.7① カット毎の抵抗値変化量判定機能項目追加　３項目シフトして015_から018_へ変更
            Dim intRetraceCntMin As Short = Short.Parse(GetPrivateProfileString_S("CUT_CUT", ("018_MIN"), m_sPath, "1"))
            Dim intRetraceCntMax As Short = Short.Parse(GetPrivateProfileString_S("CUT_CUT", ("018_MAX"), m_sPath, "10"))
            ' リトレースのオフセットＸ
            Dim dblRetraceOffXMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_RETRACE_PARA", ("000_MIN"), m_sPath, "-10.0"))
            Dim dblRetraceOffXMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_RETRACE_PARA", ("000_MAX"), m_sPath, "10.0"))
            ' リトレースのオフセットＹ
            Dim dblRetraceOffYMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_RETRACE_PARA", ("001_MIN"), m_sPath, "-10.0"))
            Dim dblRetraceOffYMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_RETRACE_PARA", ("001_MAX"), m_sPath, "10.0"))
            ' ストレートカット・リトレースのQレート
            Dim dblRetraceQrateMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_RETRACE_PARA", ("002_MIN"), m_sPath, "0.1"))
            Dim dblRetraceQrateMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_RETRACE_PARA", ("002_MAX"), m_sPath, "40.0"))
            ' ストレートカット・リトレースのトリム速度
            Dim dblRetraceSpeedMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_RETRACE_PARA", ("003_MIN"), m_sPath, "0.1"))
            Dim dblRetraceSpeedMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_RETRACE_PARA", ("003_MAX"), m_sPath, "400.0"))

            'V2.2.0.0②↓
            'Uカットデータの上限値取得 
            Dim dblUcutLen1Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("000_MIN"), m_sPath, "0.0"))
            Dim dblUcutLen1Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("000_MAX"), m_sPath, "20.0"))
            Dim dblUcutLen2Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("001_MIN"), m_sPath, "0.0"))
            Dim dblUcutLen2Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("001_MAX"), m_sPath, "20.0"))
            Dim dblUcutR1Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("002_MIN"), m_sPath, "0.0"))
            Dim dblUcutR1Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("002_MAX"), m_sPath, "20.0"))
            Dim dblUcutR2Min As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("003_MIN"), m_sPath, "0.0"))
            Dim dblUcutR2Max As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("003_MAX"), m_sPath, "20.0"))
            Dim dblUcutQMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("004_MIN"), m_sPath, "0.1"))
            Dim dblUcutQMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("004_MAX"), m_sPath, "40.0"))
            Dim dblUcutSpdMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("005_MIN"), m_sPath, "0.1"))
            Dim dblUcutSpdMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("005_MAX"), m_sPath, "400.0"))
            Dim dblUcutLturnMin As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("006_MIN"), m_sPath, "0.0"))
            Dim dblUcutLturnMax As Double = Double.Parse(GetPrivateProfileString_S("CUT_UCUT_PARA", ("006_MAX"), m_sPath, "100.0"))
            'V2.2.0.0②↑


            With m_MainEdit
                For iRn As Integer = 1 To .W_PLT.RCount                     ' 抵抗数分繰返す
                    If UserModule.IsCutResistorIncMarking(.W_REG, iRn) Then
                        For iCn As Integer = 1 To .W_REG(iRn).intTNN            ' カット数分繰返す
                            strMSG = "抵抗番号" & iRn.ToString & " カット番号" & iCn.ToString
                            m_ResNo = iRn : m_CutNo = iCn

                            If .W_REG(iRn).STCUT(iCn).intCTYP = CNS_CUTP_SP And (.W_REG(iRn).STCUT(iCn).intNum < intNumMin Or intNumMax < .W_REG(iRn).STCUT(iCn).intNum) Then          ' ｻｰﾍﾟﾝﾀｲﾝ本数
                                strMin = intNumMin.ToString : strMax = intNumMax.ToString
                                strMSG = strMSG & "の本数を" : GoTo ERR_MESSAGE
                            ElseIf (.W_REG(iRn).STCUT(iCn).intQF1 / 10.0) < intQF1Min Or intQF1Max < (.W_REG(iRn).STCUT(iCn).intQF1 / 10.0) Then      ' Qﾚｰﾄ
                                strMin = intQF1Min.ToString : strMax = intQF1Max.ToString
                                strMSG = strMSG & "Ｑレートを" : GoTo ERR_MESSAGE
                            ElseIf .W_REG(iRn).STCUT(iCn).dblV1 < dblV1Min Or dblV1Max < .W_REG(iRn).STCUT(iCn).dblV1 Then          ' 速度
                                strMin = dblV1Min.ToString : strMax = dblV1Max.ToString
                                strMSG = strMSG & "速度を" : GoTo ERR_MESSAGE
                            ElseIf .W_REG(iRn).STCUT(iCn).dblSTX < dblSTXMin Or dblSTXMax < .W_REG(iRn).STCUT(iCn).dblSTX Then      ' ｶｯﾄ位置X
                                strMin = dblSTXMin.ToString : strMax = dblSTXMax.ToString
                                strMSG = strMSG & "カット位置Ｘを" : GoTo ERR_MESSAGE
                            ElseIf .W_REG(iRn).STCUT(iCn).dblSTY < dblSTYMin Or dblSTYMax < .W_REG(iRn).STCUT(iCn).dblSTY Then      ' ｶｯﾄ位置Y
                                strMin = dblSTYMin.ToString : strMax = dblSTYMax.ToString
                                strMSG = strMSG & "カット位置Ｙを" : GoTo ERR_MESSAGE
                            ElseIf .W_REG(iRn).STCUT(iCn).dblSX2 < dblSX2Min Or dblSX2Max < .W_REG(iRn).STCUT(iCn).dblSX2 Then      ' ｶｯﾄ位置2X
                                strMin = dblSX2Min.ToString : strMax = dblSX2Max.ToString
                                strMSG = strMSG & "カット位置Ｘ２を" : GoTo ERR_MESSAGE
                            ElseIf .W_REG(iRn).STCUT(iCn).dblSY2 < dblSY2Min Or dblSY2Max < .W_REG(iRn).STCUT(iCn).dblSY2 Then      ' ｶｯﾄ位置2Y
                                strMin = dblSY2Min.ToString : strMax = dblSY2Max.ToString
                                strMSG = strMSG & "カット位置Ｙ２を" : GoTo ERR_MESSAGE
                            ElseIf .W_REG(iRn).STCUT(iCn).dblDL2 < dblDL2Min Or dblDL2Max < .W_REG(iRn).STCUT(iCn).dblDL2 Then      ' ｶｯﾄ長1
                                strMin = dblDL2Min.ToString : strMax = dblDL2Max.ToString
                                strMSG = strMSG & "カット長１を" : GoTo ERR_MESSAGE
                                'ElseIf .W_REG(iRn).STCUT(iCn).intCTYP > 1 And (.W_REG(iRn).STCUT(iCn).dblDL3 < dblDL3Min Or dblDL3Max < .W_REG(iRn).STCUT(iCn).dblDL3) Then      ' ｶｯﾄ長2
                                '    strMin = dblDL3Min.ToString : strMax = dblDL3Max.ToString
                                '    strMSG = strMSG & "カット長２を" : GoTo ERR_MESSAGE
                            ElseIf .W_REG(iRn).STCUT(iCn).dblCOF < dblCOFMin Or dblCOFMax < .W_REG(iRn).STCUT(iCn).dblCOF Then      ' ｶｯﾄｵﾌ
                                strMin = dblCOFMin.ToString : strMax = dblCOFMax.ToString
                                strMSG = strMSG & "カットオフを" : GoTo ERR_MESSAGE
                            ElseIf .W_REG(iRn).STCUT(iCn).dblLTP < dblLTPMin Or dblLTPMax < .W_REG(iRn).STCUT(iCn).dblLTP Then      ' Lﾀｰﾝﾎﾟｲﾝﾄ
                                strMin = dblLTPMin.ToString : strMax = dblLTPMax.ToString
                                strMSG = strMSG & "Ｌターンポイントを" : GoTo ERR_MESSAGE
                            ElseIf ((.W_REG(iRn).STCUT(iCn).intCTYP = CNS_CUTP_M) And (.W_REG(iRn).intSLP <> SLP_MARK)) Then       ' 文字マーキング 'V2.2.1.7①
                                Dim iLen As Integer = .W_REG(iRn).STCUT(iCn).cFormat.Length
                                If (iLen < intLetterLenMin) Or (intLetterLenMax < iLen) Then
                                    strMin = intLetterLenMin.ToString("0") : strMax = intLetterLenMax.ToString("0")
                                    strMSG = strMSG & "文字を" : GoTo ERR_MESSAGE
                                End If
                                If .W_REG(iRn).STCUT(iCn).dblDL2 < 0.1 Or .W_REG(iRn).STCUT(iCn).dblDL2 > 10.0 Then
                                    strMin = "1.0" : strMax = "10.0"
                                    strMSG = strMSG & "文字高さを" : GoTo ERR_MESSAGE
                                End If
                            End If
                            ' ARATA
                            If .W_REG(iRn).STCUT(iCn).intCUT = CNS_CUTM_IX Then   ' ｶｯﾄ方法(1:ﾄﾗｯｷﾝｸﾞ, 2:ｲﾝﾃﾞｯｸｽ, 3:NG, 4:強制カット(FULL CUT))
                                For idx As Integer = 1 To MAXIDX
                                    If .W_REG(iRn).STCUT(iCn).intIXN(idx) < intIXNMin Or intIXNMax < .W_REG(iRn).STCUT(iCn).intIXN(idx) Then                ' ｶｯﾄ回数
                                        strMin = intIXNMin.ToString : strMax = intIXNMax.ToString
                                        strMSG = strMSG & "インデックス番号" & idx.ToString & "のインデックスカット数を" : GoTo ERR_MESSAGE
                                    Else
                                        If .W_REG(iRn).STCUT(iCn).intIXN(idx) > 0 Then
                                            If .W_REG(iRn).STCUT(iCn).dblDL1(idx) < dblDL1Min Or dblDL1Max < .W_REG(iRn).STCUT(iCn).dblDL1(idx) Then        ' ｶｯﾄ長(ｲﾝﾃﾞｯｸｽﾋﾟｯﾁ)
                                                strMin = dblDL1Min.ToString : strMax = dblDL1Max.ToString
                                                strMSG = strMSG & "インデックス番号" & idx.ToString & "のカット長を" : GoTo ERR_MESSAGE
                                            ElseIf .W_REG(iRn).STCUT(iCn).lngPAU(idx) < lngPAUMin Or lngPAUMax < .W_REG(iRn).STCUT(iCn).lngPAU(idx) Then    ' ﾋﾟｯﾁ間ﾎﾟｰｽﾞ
                                                strMin = lngPAUMin.ToString : strMax = lngPAUMax.ToString
                                                strMSG = strMSG & "インデックス番号" & idx.ToString & "のピッチ間ポーズを" : GoTo ERR_MESSAGE
                                            ElseIf .W_REG(iRn).STCUT(iCn).dblDEV(idx) < dblDEVMin Or dblDEVMax < .W_REG(iRn).STCUT(iCn).dblDEV(idx) Then    ' 誤差
                                                strMin = dblDEVMin.ToString : strMax = dblDEVMax.ToString
                                                strMSG = strMSG & "インデックス番号" & idx.ToString & "の誤差を" : GoTo ERR_MESSAGE
                                            End If
                                        End If
                                    End If
                                Next
                            End If

                            If .W_REG(iRn).STCUT(iCn).intCTYP = CNS_CUTP_L Then   ' ６点ターンポイントＬカットトリミング
                                For idx As Integer = 1 To MAX_LCUT
                                    If .W_REG(iRn).STCUT(iCn).dCutLen(idx) < dCutLenMin Or dCutLenMax < .W_REG(iRn).STCUT(iCn).dCutLen(idx) Then
                                        strMin = dCutLenMin.ToString : strMax = dCutLenMax.ToString
                                        strMSG = strMSG & "番号" & idx.ToString & "のカット長を" : GoTo ERR_MESSAGE
                                    ElseIf .W_REG(iRn).STCUT(iCn).dQRate(idx) / 10.0 < dQRateMin Or dQRateMax < .W_REG(iRn).STCUT(iCn).dQRate(idx) / 10.0 Then
                                        strMin = dQRateMin.ToString : strMax = dQRateMax.ToString
                                        strMSG = strMSG & "番号" & idx.ToString & "のＱレートを" : GoTo ERR_MESSAGE
                                    ElseIf .W_REG(iRn).STCUT(iCn).dSpeed(idx) < dSpeedMin Or dSpeedMax < .W_REG(iRn).STCUT(iCn).dSpeed(idx) Then
                                        strMin = dSpeedMin.ToString : strMax = dSpeedMax.ToString
                                        strMSG = strMSG & "番号" & idx.ToString & "の速度を" : GoTo ERR_MESSAGE
                                    ElseIf idx < MAX_LCUT And (.W_REG(iRn).STCUT(iCn).dTurnPoint(idx) < dTurnPointMin Or dTurnPointMax < .W_REG(iRn).STCUT(iCn).dTurnPoint(idx)) Then
                                        strMin = dTurnPointMin.ToString : strMax = dTurnPointMax.ToString
                                        strMSG = strMSG & "番号" & idx.ToString & "のターンポイントを" : GoTo ERR_MESSAGE
                                    End If
                                Next
                            End If
                            If .W_REG(iRn).STCUT(iCn).intCTYP = CNS_CUTP_ST_TR Then   ' ストレートカット・リトレース
                                If .W_REG(iRn).STCUT(iCn).intRetraceCnt < intRetraceCntMin Or intRetraceCntMax < .W_REG(iRn).STCUT(iCn).intRetraceCnt Then      ' Qﾚｰﾄ'V2.1.0.0①intQF1Min→intRetraceCntMin,intQF1Max→intRetraceCntMax修正
                                    strMin = intRetraceCntMin.ToString : strMax = intRetraceCntMax.ToString
                                    strMSG = strMSG & "リトレース本数を" : GoTo ERR_MESSAGE
                                Else
                                    For idx As Integer = 1 To .W_REG(iRn).STCUT(iCn).intRetraceCnt

                                        If .W_REG(iRn).STCUT(iCn).dblRetraceOffX(idx) < dblRetraceOffXMin Or dblRetraceOffXMax < .W_REG(iRn).STCUT(iCn).dblRetraceOffX(idx) Then
                                            strMin = dblRetraceOffXMin.ToString : strMax = dblRetraceOffXMax.ToString
                                            strMSG = strMSG & "番号" & idx.ToString & "のリトレースのオフセットＸを" : GoTo ERR_MESSAGE
                                        ElseIf .W_REG(iRn).STCUT(iCn).dblRetraceOffY(idx) < dblRetraceOffYMin Or dblRetraceOffYMax < .W_REG(iRn).STCUT(iCn).dblRetraceOffY(idx) Then
                                            strMin = dblRetraceOffYMin.ToString : strMax = dblRetraceOffYMax.ToString
                                            strMSG = strMSG & "番号" & idx.ToString & "のリトレースのオフセットＹを" : GoTo ERR_MESSAGE
                                        ElseIf .W_REG(iRn).STCUT(iCn).dblRetraceQrate(idx) / 10.0 < dblRetraceQrateMin Or dblRetraceQrateMax < .W_REG(iRn).STCUT(iCn).dblRetraceQrate(idx) / 10.0 Then
                                            strMin = dblRetraceQrateMin.ToString : strMax = dblRetraceQrateMax.ToString
                                            strMSG = strMSG & "番号" & idx.ToString & "のリトレースＱレートを" : GoTo ERR_MESSAGE
                                        ElseIf .W_REG(iRn).STCUT(iCn).dblRetraceSpeed(idx) < dblRetraceSpeedMin Or dblRetraceSpeedMax < .W_REG(iRn).STCUT(iCn).dblRetraceSpeed(idx) Then
                                            strMin = dblRetraceSpeedMin.ToString : strMax = dblRetraceSpeedMax.ToString
                                            strMSG = strMSG & "番号" & idx.ToString & "のリトレースの速度を" : GoTo ERR_MESSAGE
                                        End If
                                    Next
                                End If
                            End If
#If cOSCILLATORcFLcUSE And cFLcAUTOcPOWER Then
                            If (gSysPrm.stRAT.giOsc_Res = OSCILLATOR_FL) Then
                                For idx As Integer = 1 To MAXCND

                                    'If .W_FLCND.Curr(.W_REG(iRn).STCUT(iCn).intCND(idx)) < CurrMin Or CurrMax < .W_FLCND.Curr(.W_REG(iRn).STCUT(iCn).intCND(idx)) Then                ' 電流値
                                    '    strMin = CurrMin.ToString : strMax = CurrMax.ToString
                                    '    strMSG = strMSG & "カット条件番号" & idx.ToString & "の電流値を" : GoTo ERR_MESSAGE
                                    'ElseIf .W_FLCND.Freq(.W_REG(iRn).STCUT(iCn).intCND(idx)) < FreqMin Or FreqMax < .W_FLCND.Freq(.W_REG(iRn).STCUT(iCn).intCND(idx)) Then            ' STEG本数
                                    '    strMin = FreqMin.ToString : strMax = FreqMax.ToString
                                    '    strMSG = strMSG & "カット条件番号" & idx.ToString & "のＱレートを" : GoTo ERR_MESSAGE
                                    'ElseIf .W_FLCND.Steg(.W_REG(iRn).STCUT(iCn).intCND(idx)) < StegMin Or StegMax < .W_FLCND.Steg(.W_REG(iRn).STCUT(iCn).intCND(idx)) Then            ' Qﾚｰﾄ
                                    '    strMin = StegMin.ToString : strMax = StegMax.ToString
                                    '    strMSG = strMSG & "カット条件番号" & idx.ToString & "のＳＴＥＧ本数を" : GoTo ERR_MESSAGE
                                    'End If

                                    dblPowerAdjustTargetMax = ObjFiberLaser.GetMaxPower(.W_FLCND.Freq(.W_REG(iRn).STCUT(iCn).intCND(idx)), .W_FLCND.Steg(.W_REG(iRn).STCUT(iCn).intCND(idx)))
                                    If .W_FLCND.dblPowerAdjustTarget(.W_REG(iRn).STCUT(iCn).intCND(idx)) < dblPowerAdjustTargetMin Or dblPowerAdjustTargetMax < .W_FLCND.dblPowerAdjustTarget(.W_REG(iRn).STCUT(iCn).intCND(idx)) Then                    ' 目標パワー（Ｗ）
                                        strMin = dblPowerAdjustTargetMin.ToString : strMax = dblPowerAdjustTargetMax.ToString
                                        strMSG = strMSG & "カット条件番号" & idx.ToString & "の目標パワー（Ｗ）を" : GoTo ERR_MESSAGE
                                    ElseIf .W_FLCND.dblPowerAdjustToleLevel(.W_REG(iRn).STCUT(iCn).intCND(idx)) < dblPowerAdjustToleLevelMin Or dblPowerAdjustToleLevelMax < .W_FLCND.dblPowerAdjustToleLevel(.W_REG(iRn).STCUT(iCn).intCND(idx)) Then    ' 許容範囲（±Ｗ）
                                        strMin = dblPowerAdjustToleLevelMin.ToString : strMax = dblPowerAdjustToleLevelMax.ToString
                                        strMSG = strMSG & "カット条件番号" & idx.ToString & "の許容範囲（±Ｗ）を" : GoTo ERR_MESSAGE
                                    End If

                                    If .W_REG(iRn).STCUT(iCn).intCTYP = CNS_CUTP_ST And idx = 1 Then  ' ストレート
                                        Exit For
                                    ElseIf .W_REG(iRn).STCUT(iCn).intCTYP = CNS_CUTP_L And idx = 2 Then  ' ストレート
                                        Exit For
                                    End If
                                Next
                            End If
#End If
                            'V2.2.0.0②↓
                            'Uカットパラメータのチェック
                            If .W_REG(iRn).STCUT(iCn).intCTYP = CNS_CUTP_U Then   ' Uカットトリミング
                                If .W_REG(iRn).STCUT(iCn).dUCutL1 < dblUcutLen1Min Or dblUcutLen1Max < .W_REG(iRn).STCUT(iCn).dUCutL1 Then      ' L1カット長
                                    strMin = dblUcutLen1Min.ToString : strMax = dblUcutLen1Max.ToString
                                    strMSG = strMSG & "L1カット長を" : GoTo ERR_MESSAGE
                                End If
                                If .W_REG(iRn).STCUT(iCn).dUCutL2 < dblUcutLen1Min Or dblUcutLen1Max < .W_REG(iRn).STCUT(iCn).dUCutL2 Then      ' L2カット長
                                    strMin = dblUcutLen2Min.ToString : strMax = dblUcutLen2Max.ToString
                                    strMSG = strMSG & "L2カット長を" : GoTo ERR_MESSAGE
                                End If
                                If .W_REG(iRn).STCUT(iCn).dblUCutR1 < dblUcutR1Min Or dblUcutR1Max < .W_REG(iRn).STCUT(iCn).dblUCutR1 Then     'R1
                                    strMin = dblUcutR1Min.ToString : strMax = dblUcutR1Max.ToString
                                    strMSG = strMSG & "R1半径を" : GoTo ERR_MESSAGE
                                End If
                                If .W_REG(iRn).STCUT(iCn).dblUCutR2 < dblUcutR2Min Or dblUcutR2Max < .W_REG(iRn).STCUT(iCn).dblUCutR2 Then     'R2
                                    strMin = dblUcutR2Min.ToString : strMax = dblUcutR2Max.ToString
                                    strMSG = strMSG & "R2半径を" : GoTo ERR_MESSAGE
                                End If
                                If (.W_REG(iRn).STCUT(iCn).intUCutQF1 / 10.0) < dblUcutQMin Or dblUcutQMax < (.W_REG(iRn).STCUT(iCn).intUCutQF1 / 10.0) Then     'Qレート
                                    strMin = dblUcutQMin.ToString : strMax = dblUcutQMax.ToString
                                    strMSG = strMSG & "Qレートを" : GoTo ERR_MESSAGE
                                End If
                                If .W_REG(iRn).STCUT(iCn).dblUCutV1 < dblUcutSpdMin Or dblUcutSpdMax < .W_REG(iRn).STCUT(iCn).dblUCutV1 Then     '速度
                                    strMin = dblUcutSpdMin.ToString : strMax = dblUcutSpdMax.ToString
                                    strMSG = strMSG & "速度を" : GoTo ERR_MESSAGE
                                End If
                                If .W_REG(iRn).STCUT(iCn).dblUCutTurnP < dblUcutLturnMin Or dblUcutLturnMax < .W_REG(iRn).STCUT(iCn).dblUCutTurnP Then     '
                                    strMin = dblUcutSpdMin.ToString : strMax = dblUcutSpdMax.ToString
                                    strMSG = strMSG & "ターンポイントを" : GoTo ERR_MESSAGE
                                End If
                            End If
                            'V2.2.0.0②↑

                        Next iCn
                    End If
                Next iRn
            End With

            Exit Function
ERR_MESSAGE:
            m_MainEdit.MTab.SelectedIndex = m_TabIdx  ' ﾀﾌﾞ表示切替
            Call SetDataToText()    ' ﾁｪｯｸする抵抗番号、ｶｯﾄ番号のﾃﾞｰﾀをｺﾝﾄﾛｰﾙにｾｯﾄする
            CutDataCheckByDataOnly = 1
            strMSG = strMSG & strMin & "～" & strMax & "の範囲で指定して下さい"
            Call MsgBox(strMSG, DirectCast( _
                        MsgBoxStyle.OkOnly + _
                        MsgBoxStyle.Information, MsgBoxStyle), _
                        My.Application.Info.Title)

        End Function
#End Region
#Else
        ''' <summary>すべてのﾃｷｽﾄﾎﾞｯｸｽのﾃﾞｰﾀﾁｪｯｸをおこなう</summary>
        ''' <returns>0=正常, 1=ｴﾗｰ</returns>
        Friend Overrides Function CheckAllTextData() As Integer
            Dim ret As Integer
            Try
                m_CheckFlg = True ' ﾁｪｯｸ中(tabBase_Layoutにて使用)
                With m_MainEdit
                    .MTab.SelectedIndex = m_TabIdx  ' ﾀﾌﾞ表示切替

                    For rn As Integer = 1 To .W_PLT.RCount Step 1
                        m_ResNo = rn
                        With .W_REG(rn)

                            ' ------------------------
                            If (SLP_VMES = m_MainEdit.W_REG(m_ResNo).intSLP) OrElse (SLP_RMES = m_MainEdit.W_REG(m_ResNo).intSLP) Then
                                ' 抵抗のｽﾛｰﾌﾟが 5:電圧測定のみ, 6:抵抗測定のみ の場合
                                Continue For    ' ｶｯﾄのﾁｪｯｸはおこなわない(次の抵抗へ)
                            End If
                            ' ------------------------

                            ' TODO: ｶｯﾄ数が0になることはない仕様のため不要と思われる
                            If (.intTNN < 1) Then ' ｶｯﾄ数 < 1 ?
                                Dim strMsg As String
                                strMsg = "抵抗番号" & rn.ToString("0") & "のカットデータがありません。" & vbCrLf
                                strMsg = strMsg & "カットデータを登録してください。"
                                Call MsgBox(strMsg, MsgBoxStyle.OkOnly, My.Application.Info.Title)
                                Exit Try
                            End If

                            ' ｶｯﾄ数分繰返す
                            For cn As Integer = 1 To .intTNN Step 1
                                m_CutNo = cn
                                With .STCUT(cn)

                                    ' ﾁｪｯｸする抵抗番号、ｶｯﾄ番号のﾃﾞｰﾀをｺﾝﾄﾛｰﾙにｾｯﾄする
                                    Call SetDataToText()

                                    ' ｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                                    ret = CheckControlData(m_CtlCut)
                                    If (ret <> 0) Then Exit Try

                                    ' ｶｯﾄ方法がｲﾝﾃﾞｯｸｽｶｯﾄの場合(1:ﾄﾗｯｷﾝｸﾞ, 2:ｲﾝﾃﾞｯｸｽ, 3:NGｶｯﾄ)
                                    If (CNS_CUTM_IX = .intCUT) Then
                                        ' ｲﾝﾃﾞｯｸｽｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                                        ret = CheckControlData(m_CtlIdxCut)
                                        If (ret <> 0) Then Exit Try
                                    End If

                                    ' 相関ﾁｪｯｸ
                                    ret = CheckRelation()
                                    If (ret <> 0) Then Exit Try
                                End With
                            Next cn

                        End With
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
#End If
#End Region

#Region "ﾃﾞｰﾀﾁｪｯｸ関数を呼び出す"
        ''' <summary>ﾃｷｽﾄﾎﾞｯｸｽのﾃﾞｰﾀﾁｪｯｸ関数を呼び出す</summary>
        ''' <param name="cTextBox">ﾁｪｯｸするﾃｷｽﾄﾎﾞｯｸｽ</param>
        ''' <returns>0=正常, 1=ｴﾗｰ</returns>
        Protected Overrides Function CheckTextData(ByRef cTextBox As cTxt_) As Integer
            Dim strMsg As String
            Dim tag As Integer
            Dim ret As Integer
            Dim dblWK As Double
            Dim i As Integer
            Try
                ' 抵抗ﾃﾞｰﾀ登録数ﾁｪｯｸ
                ' TODO: 抵抗数が0になることはない仕様のため不要と思われる
                If (m_ResNo < 1) Then
                    strMsg = "抵抗データがありません。抵抗データを登録してください。"
                    Call MsgBox(strMsg, MsgBoxStyle.OkOnly, My.Application.Info.Title)
                    ret = 1
                    Exit Try
                End If

                With m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo)
                    ' ｶｯﾄ数ﾁｪｯｸ
                    ' TODO: ｶｯﾄ数が0になることはない仕様のため不要と思われる
                    If (m_MainEdit.W_REG(m_ResNo).intTNN < 1) Then ' ｶｯﾄ数 < 1 ?
                        strMsg = "抵抗番号" & m_ResNo.ToString("0") & "のカットデータがありません。" & vbCrLf
                        strMsg = strMsg & "追加ボタンを押下してカットデータを登録してください。"
                        Call MsgBox(strMsg, MsgBoxStyle.OkOnly, My.Application.Info.Title)
                        ret = 1
                        Exit Try
                    End If

                    tag = DirectCast(cTextBox.Tag, Integer)
                    Select Case (DirectCast(cTextBox.Parent.Tag, Integer))
                        ' ------------------------------------------------------------------------------
                        Case 0 ' ｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            Select Case (tag)
                                Case 0 ' ｻｰﾍﾟﾝﾀｲﾝ本数(ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ時有効)
                                    If (CNS_CUTP_SP = .intCTYP) Then ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄの場合
                                        ret = CheckShortData(cTextBox, .intNum)
                                    End If
                                Case 1 ' Qﾚｰﾄ
                                    If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then ' ﾌｧｲﾊﾞﾚｰｻﾞではない場合
                                        dblWK = .intQF1 / 10.0
                                        ret = CheckDoubleData(cTextBox, dblWK)
                                        .intQF1 = Convert.ToInt16(dblWK * 10) ' (KHz→0.1KHz)
                                    End If
                                Case 2 ' 速度
                                    ret = CheckDoubleData(cTextBox, .dblV1)
                                Case 3 ' ｶｯﾄ位置X
                                    ret = CheckDoubleData(cTextBox, .dblSTX)
                                Case 4 ' ｶｯﾄ位置Y
                                    ret = CheckDoubleData(cTextBox, .dblSTY)
                                Case 5 ' ｶｯﾄ位置2X(ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ時有効)
                                    If (CNS_CUTP_ST_TR = .intCTYP) Then 'V1.0.4.3③ リトレースの場合　サーペンタイン（CNS_CUTP_SP）から変更
                                        ret = CheckDoubleData(cTextBox, .dblSX2)
                                    End If
                                Case 6 ' ｶｯﾄ位置2Y(ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ時有効)
                                    If (CNS_CUTP_ST_TR = .intCTYP) Then 'V1.0.4.3③ リトレースの場合　サーペンタイン（CNS_CUTP_SP）から変更
                                        ret = CheckDoubleData(cTextBox, .dblSY2)
                                    End If
                                Case 7 ' ｶｯﾄ長1
                                    ret = CheckDoubleData(cTextBox, .dblDL2)
                                Case 8 ' ｶｯﾄ長2
                                    If (CNS_CUTP_ST <> .intCTYP) Then ' ｽﾄﾚｰﾄｶｯﾄではない場合
                                        ret = CheckDoubleData(cTextBox, .dblDL3)
                                    End If
                                Case 9 ' ｶｯﾄｵﾌ
                                    If (CNS_CUTM_NG <> .intCUT) Then ' NGｶｯﾄではない場合
                                        ret = CheckDoubleData(cTextBox, .dblCOF)
                                    End If
                                Case 10 ' Lﾀｰﾝﾎﾟｲﾝﾄ
                                    If (CNS_CUTP_L = .intCTYP) Then ' Lｶｯﾄの場合
                                        ret = CheckDoubleData(cTextBox, .dblLTP)
                                    End If
                                Case 11 ' ###1042①　入力チェック追加する
                                    If (CNS_CUTP_M = .intCTYP) Then ' 文字マーキング
                                        ret = CheckStrData(cTextBox, .cFormat)
                                    End If
                                    'V2.2.1.7① ↓
                                Case 12 ' 印字固定部
                                    If (CNS_CUTP_M = .intCTYP) Then ' 文字マーキング
                                        ret = CheckStrData(cTextBox, .cMarkFix)
                                    End If
                                Case 13 ' 開始番号
                                    If (CNS_CUTP_M = .intCTYP) Then ' 文字マーキング
                                        Dim Inputstr = cTextBox.Text.Trim()
                                        If Inputstr <> "" Then
                                            ret = CheckNumeric(cTextBox)
                                            If (ret <> -1) Then
                                                ret = CheckStrData(cTextBox, .cMarkStartNum)
                                            Else
                                                cTextBox.Text = .cMarkStartNum
                                            End If
                                        Else
                                            .cMarkStartNum = ""
                                        End If
                                    End If
                                Case 14 ' 重複回数
                                    If (CNS_CUTP_M = .intCTYP) Then ' 文字マーキング
                                        ret = CheckShortData(cTextBox, .intMarkRepeatCnt)
                                    End If
                                                                        'V2.1.0.0①↓
                                Case 15 ' 上昇率
                                    ret = CheckDoubleData(cTextBox, .dRateOfUp)
                                    If ret = 0 And .iVariationRepeat = 1 Then
                                        UserSub.CutVariationDataCopy(m_MainEdit.W_PLT, m_MainEdit.W_REG, m_ResNo, m_CutNo)
                                    End If
                                Case 16 ' 下限値
                                    ret = CheckDoubleData(cTextBox, .dVariationLow)
                                    If ret = 0 And .iVariationRepeat = 1 Then
                                        UserSub.CutVariationDataCopy(m_MainEdit.W_PLT, m_MainEdit.W_REG, m_ResNo, m_CutNo)
                                    End If
                                Case 17 ' 上限値
                                    ret = CheckDoubleData(cTextBox, .dVariationHi)
                                    If ret = 0 And .iVariationRepeat = 1 Then
                                        UserSub.CutVariationDataCopy(m_MainEdit.W_PLT, m_MainEdit.W_REG, m_ResNo, m_CutNo)
                                    End If
                                    'V2.1.0.0①↑

                                    '    'V2.1.0.0①↓
                                    'Case 12 ' 上昇率
                                    '    ret = CheckDoubleData(cTextBox, .dRateOfUp)
                                    '    If ret = 0 And .iVariationRepeat = 1 Then
                                    '        UserSub.CutVariationDataCopy(m_MainEdit.W_PLT, m_MainEdit.W_REG, m_ResNo, m_CutNo)
                                    '    End If
                                    'Case 13 ' 下限値
                                    '    ret = CheckDoubleData(cTextBox, .dVariationLow)
                                    '    If ret = 0 And .iVariationRepeat = 1 Then
                                    '        UserSub.CutVariationDataCopy(m_MainEdit.W_PLT, m_MainEdit.W_REG, m_ResNo, m_CutNo)
                                    '    End If
                                    'Case 14 ' 上限値
                                    '    ret = CheckDoubleData(cTextBox, .dVariationHi)
                                    '    If ret = 0 And .iVariationRepeat = 1 Then
                                    '        UserSub.CutVariationDataCopy(m_MainEdit.W_PLT, m_MainEdit.W_REG, m_ResNo, m_CutNo)
                                    '    End If
                                    '    'V2.1.0.0①↑
                                    'V2.2.1.7① ↑
                                    'V1.0.4.3⑤↓
                                    'V2.1.0.0①                                Case 12 ' ストレートカット・リトレース
                                    'V2.1.0.0①                                    If (CNS_CUTP_ST_TR = .intCTYP) Then ' ストレートカット・リトレースの場合
                                    'V2.1.0.0①                                        If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then ' ﾌｧｲﾊﾞﾚｰｻﾞではない場合
                                    'V2.1.0.0①                                            ret = CheckDoubleData(cTextBox, dblWK)
                                    'V2.1.0.0①                                            .intQF2 = Convert.ToInt16(dblWK * 10) ' (KHz→0.1KHz)
                                    'V2.1.0.0①                                        End If
                                    'V2.1.0.0①                                    End If
                                    'V2.1.0.0①                                Case 13 ' 速度２
                                    'V2.1.0.0①                                    If (CNS_CUTP_ST_TR = .intCTYP) Then ' ストレートカット・リトレースの場合
                                    'V2.1.0.0①                                        ret = CheckDoubleData(cTextBox, .dblV2)
                                    'V2.1.0.0①                                    End If
                                    'V1.0.4.3⑤↑
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case 1 ' ｲﾝﾃﾞｯｸｽｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            If (CNS_CUTM_IX = .intCUT) Then ' ｶｯﾄ方法がｲﾝﾃﾞｯｸｽｶｯﾄの場合
                                Select Case (tag)
                                    Case 0 ' ｶｯﾄ数
                                        Dim cutNo As Integer = (GetCtlIdx(cTextBox, tag) + 1)
                                        ret = CheckShortData(cTextBox, .intIXN(cutNo))
                                    Case 1 ' ｶｯﾄ長(ｲﾝﾃﾞｯｸｽﾋﾟｯﾁ)
                                        Dim cutNo As Integer = (GetCtlIdx(cTextBox, tag) + 1)
                                        ret = CheckDoubleData(cTextBox, .dblDL1(cutNo))
                                    Case 2 ' ﾎﾟｰｽﾞ
                                        Dim cutNo As Integer = (GetCtlIdx(cTextBox, tag) + 1)
                                        ret = CheckIntData(cTextBox, .lngPAU(cutNo))
                                    Case 3 ' 誤差
                                        Dim cntNo As Integer = (GetCtlIdx(cTextBox, tag) + 1)
                                        ret = CheckDoubleData(cTextBox, .dblDEV(cntNo))
                                    Case Else
                                        Throw New Exception("Case " & tag & ": Nothing")
                                End Select
                            End If
                            ' ------------------------------------------------------------------------------
                        Case 2 ' FL加工条件ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ(表示のみのためﾁｪｯｸなし)
                            Throw New Exception("Parent.Tag - Case " & tag & ": Nothing")
                            ' ------------------------------------------------------------------------------
                        Case 3 'Ｌカットパラメータグループボックス 
                            If (CNS_CUTP_L = .intCTYP) Then ' カット方法が、Ｌカットの場合
                                Dim cutNo As Integer = (GetCtlLCutIdx(cTextBox, tag) + 1)
                                Select Case (tag)
                                    Case 0 ' カット長

                                        ret = CheckDoubleData(cTextBox, .dCutLen(cutNo))
                                        .dblDL2 = .dCutLen(1)
                                        .dblDL3 = 0.0
                                        For i = 2 To MAX_LCUT
                                            .dblDL3 = .dblDL3 + .dCutLen(i)
                                        Next
                                    Case 1 ' Ｑレート
                                        If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then ' ﾌｧｲﾊﾞﾚｰｻﾞではない場合
                                            dblWK = .dQRate(cutNo) / 10.0
                                            ret = CheckDoubleData(cTextBox, dblWK)
                                            .dQRate(cutNo) = Convert.ToInt16(dblWK * 10.0) ' (KHz→0.1KHz)
                                        End If
                                    Case 2 ' 速度
                                        ret = CheckDoubleData(cTextBox, .dSpeed(cutNo))
                                    Case 3 ' ターンポイント
                                        ret = CheckDoubleData(cTextBox, .dTurnPoint(cutNo))
                                    Case Else
                                        Throw New Exception("Case " & tag & ": Nothing")
                                End Select
                            End If
                            'V2.0.0.0⑦↓
                        Case 4 'リトレースカットカットパラメータグループボックス 
                            If (CNS_CUTP_ST_TR = .intCTYP) Then ' カット方法が、リトレースカットの場合
                                If (m_CtlCut(CUT_RETRACE) Is cTextBox) Then
                                    ret = CheckShortData(cTextBox, .intRetraceCnt)
                                Else
                                    Dim cutNo As Integer = (GetCtlRetraceCutIdx(cTextBox, tag) + 1)
                                    Select Case (tag)
                                        Case 0 ' リトレースのオフセットＸ
                                            ret = CheckDoubleData(cTextBox, .dblRetraceOffX(cutNo))
                                        Case 1 ' リトレースのオフセットＹ
                                            ret = CheckDoubleData(cTextBox, .dblRetraceOffY(cutNo))
                                        Case 2 ' ストレートカット・リトレースのQレート(0.1KHz)に使用
                                            If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then ' ﾌｧｲﾊﾞﾚｰｻﾞではない場合
                                                dblWK = .dblRetraceQrate(cutNo) / 10.0
                                                ret = CheckDoubleData(cTextBox, dblWK)
                                                .dblRetraceQrate(cutNo) = Convert.ToInt16(dblWK * 10.0) ' (KHz→0.1KHz)
                                            End If
                                        Case 3 ' 速度
                                            ret = CheckDoubleData(cTextBox, .dblRetraceSpeed(cutNo))
                                        Case Else
                                            Throw New Exception("Case(Retrace) " & tag & ": Nothing")
                                    End Select
                                End If
                            End If
                            'V2.0.0.0⑦↑
                            'V2.2.0.0②↓
                        Case 5  ' Ｕカットパラメータ 
                            If (.intCTYP = CNS_CUTP_U) Then

                                Select Case (tag)
                                    Case 0 ' カット長
                                        ret = CheckDoubleData(cTextBox, .dUCutL1)
                                    Case 1 ' カット長
                                        ret = CheckDoubleData(cTextBox, .dUCutL2)
                                    Case 2 ' R1
                                        ret = CheckDoubleData(cTextBox, .dblUCutR1)
                                    Case 3 ' R2
                                        ret = CheckDoubleData(cTextBox, .dblUCutR2)
                                    Case 4 ' Ｑレート
                                        If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then ' ﾌｧｲﾊﾞﾚｰｻﾞではない場合
                                            dblWK = .intUCutQF1 / 10.0
                                            ret = CheckDoubleData(cTextBox, dblWK)
                                            .intUCutQF1 = Convert.ToInt16(dblWK * 10.0) ' (KHz→0.1KHz)
                                        End If
                                    Case 5 ' 速度
                                        ret = CheckDoubleData(cTextBox, .dblUCutV1)
                                    Case 6 ' ターンポイント
                                        ret = CheckDoubleData(cTextBox, .dblUCutTurnP)

                                    Case Else
                                        Throw New Exception("Case " & tag & ": Nothing")
                                End Select

                            End If

                            'V2.2.0.0②↑
                        Case Else
                            Throw New Exception("Parent.Tag - Case " & tag & ": Nothing")
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

        '        'V2.2.1.7① ↓
        '#Region "テキストボックスの文字列が数字変換できるか確認"
        '        ''' <summary>テキストボックスの文字列が数字変換できるか確認（テキストボックス用）</summary>
        '        ''' <param name="cTextBox">確認するﾃｷｽﾄﾎﾞｯｸｽ</param>
        '        ''' <returns>(-1)=ｴﾗｰ</returns>
        '        Private Function CheckNumeric(ByRef cTextBox As cTxt_) As Integer
        '            Dim ret As Integer = 0
        '            Try

        '                '数値チェック
        '                If IsNumeric(cTextBox.Text) Then
        '                    'Nop
        '                Else
        '                    MsgBox("数値を入力してください。")
        '                    ret = -1
        '                End If
        '            Catch ex As Exception
        '                Call MsgBox_Exception(ex, cTextBox.Name)
        '                ret = -1
        '            Finally
        '                CheckNumeric = ret
        '            End Try

        '        End Function
        '#End Region
        '        'V2.2.1.7① ↑

#Region "ｲﾝﾃﾞｯｸｽｶｯﾄﾃｷｽﾄﾎﾞｯｸｽのｶｯﾄ番号を返す"
        ''' <summary>m_CtlIdxCut(,)での1次元目のｲﾝﾃﾞｯｸｽを返す(ﾃｷｽﾄﾎﾞｯｸｽ用)</summary>
        ''' <param name="cTextBox">確認するﾃｷｽﾄﾎﾞｯｸｽ</param>
        ''' <param name="tag">ﾃｷｽﾄﾎﾞｯｸｽのﾀｸﾞ</param>
        ''' <returns>(-1)=ｴﾗｰ, 0~4=ｲﾝﾃﾞｯｸｽ</returns>
        Private Function GetCtlIdx(ByRef cTextBox As cTxt_, ByVal tag As Integer) As Integer
            Dim ret As Integer = (-1)
            Try
                For i As Integer = 0 To (m_CtlIdxCut.GetLength(0) - 1) Step 1
                    If (m_CtlIdxCut(i, tag) Is cTextBox) Then
                        ret = i
                        Exit For
                    End If
                Next i
            Catch ex As Exception
                Call MsgBox_Exception(ex, cTextBox.Name)
                ret = (-1)
            Finally
                GetCtlIdx = ret
            End Try

        End Function

        ''' <summary>m_CtlIdxCut(,)での1次元目のｲﾝﾃﾞｯｸｽを返す(ｺﾝﾎﾞﾎﾞｯｸｽ用)</summary>
        ''' <param name="cCombo">確認するｺﾝﾎﾞﾎﾞｯｸｽ</param>
        ''' <param name="tag">ﾃｷｽﾄﾎﾞｯｸｽのﾀｸﾞ</param>
        ''' <returns>(-1)=ｴﾗｰ, 0~4=ｲﾝﾃﾞｯｸｽ</returns>
        Private Function GetCtlIdx(ByRef cCombo As cCmb_, ByVal tag As Integer) As Integer
            Dim ret As Integer = (-1)
            Try
                For i As Integer = 0 To (m_CtlIdxCut.GetLength(0) - 1) Step 1
                    If (m_CtlIdxCut(i, IDX_MTYPE + tag) Is cCombo) Then
                        ret = i
                        Exit For
                    End If
                Next i

            Catch ex As Exception
                Call MsgBox_Exception(ex, cCombo.Name)
                ret = (-1)
            Finally
                GetCtlIdx = ret
            End Try

        End Function
#End Region

#Region "Ｌカットテキストボックスのカット番号を返す"
        ''' <summary>
        ''' m_CtlLCut(,)での1次元目のインデックスを返す(テキストボックス用)
        ''' </summary>
        ''' <param name="cTextBox">確認するテキストボックス</param>
        ''' <param name="tag">テキストボックスのタグ</param>
        ''' <returns>(-1)=エラー, 0~6=インデックス</returns>
        ''' <remarks></remarks>
        Private Function GetCtlLCutIdx(ByRef cTextBox As cTxt_, ByVal tag As Integer) As Integer
            Dim ret As Integer = (-1)
            Try
                If tag >= LCUT_DIR_IDX Then
                    tag = tag + 1
                End If
                For i As Integer = 0 To (m_CtlLCut.GetLength(0) - 1) Step 1
                    If (m_CtlLCut(i, tag) Is cTextBox) Then
                        ret = i
                        Exit For
                    End If
                Next i
            Catch ex As Exception
                Call MsgBox_Exception(ex, cTextBox.Name)
                ret = (-1)
            Finally
                GetCtlLCutIdx = ret
            End Try

        End Function

        ''' <summary>GetCtlLCutIdx(,)での1次元目のｲﾝﾃﾞｯｸｽを返す(ｺﾝﾎﾞﾎﾞｯｸｽ用)</summary>
        ''' <param name="cCombo">確認するｺﾝﾎﾞﾎﾞｯｸｽ</param>
        ''' <param name="tag">ﾃｷｽﾄﾎﾞｯｸｽのﾀｸﾞ</param>
        ''' <returns>(-1)=ｴﾗｰ, 0~4=ｲﾝﾃﾞｯｸｽ</returns>
        Private Function GetCtlLCutIdx(ByRef cCombo As cCmb_, ByVal tag As Integer) As Integer
            Dim ret As Integer = (-1)
            Try
                For i As Integer = 0 To (m_CtlLCut.GetLength(0) - 1) Step 1
                    If (m_CtlLCut(i, LCUT_DIR_IDX + tag) Is cCombo) Then
                        ret = i
                        Exit For
                    End If
                Next i

            Catch ex As Exception
                Call MsgBox_Exception(ex, cCombo.Name)
                ret = (-1)
            Finally
                GetCtlLCutIdx = ret
            End Try

        End Function
#End Region

        'V2.0.0.0⑦↓
#Region "リトレースカットテキストボックスのカット番号を返す"
        ''' <summary>
        ''' m_CtlRetraceCut(,)での1次元目のインデックスを返す(テキストボックス用)
        ''' </summary>
        ''' <param name="cTextBox">確認するテキストボックス</param>
        ''' <param name="tag">テキストボックスのタグ</param>
        ''' <returns>(-1)=エラー, 0~9=インデックス</returns>
        ''' <remarks></remarks>
        Private Function GetCtlRetraceCutIdx(ByRef cTextBox As cTxt_, ByVal tag As Integer) As Integer
            Dim ret As Integer = (-1)
            Try
                For i As Integer = 0 To (m_CtlRetraceCut.GetLength(0) - 1) Step 1
                    If (m_CtlRetraceCut(i, tag) Is cTextBox) Then
                        ret = i
                        Exit For
                    End If
                Next i
            Catch ex As Exception
                Call MsgBox_Exception(ex, cTextBox.Name)
                ret = (-1)
            Finally
                GetCtlRetraceCutIdx = ret
            End Try

        End Function
#End Region
        'V2.0.0.0⑦↑

#Region "相関ﾁｪｯｸ"
        ''' <summary>相関ﾁｪｯｸ処理</summary>
        ''' <returns>0 = 正常, 1 = ｴﾗｰ</returns>
        Protected Overrides Function CheckRelation() As Integer
            Dim strMsg As String
            Dim errIdx As Integer
            Dim ctlArray() As Control

            CheckRelation = 0 ' Return値 = 正常
            Try
                With m_MainEdit
                    ctlArray = m_CtlCut ' ｶｯﾄ(共通)ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                    '---------------------------------------------------------------------------
                    '   外部機器の指定がない場合の外部測定指定チェック
                    '---------------------------------------------------------------------------
                    ' 測定種別(0=内部測定, 1以上=外部測定)
                    If (1 <= .W_REG(m_ResNo).STCUT(m_CutNo).intMType) And (.W_PLT.GCount <= 0) Then
                        strMsg = "相関チェックエラー" & vbCrLf
                        strMsg = strMsg & "外部機器の指定がない場合は外部測定器の指定はできません。"
                        errIdx = CUT_MTYPE
                        GoTo STP_ERR
                    End If

                    '---------------------------------------------------------------------------
                    '   ｶｯﾄ方向1,2チェック(Ｌカット/ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ時)
                    '---------------------------------------------------------------------------
                    'V1.0.4.3⑦カット方向は、７カット個別に０から３５９度の範囲で指定可能。
                    'V1.0.4.3⑦                    If (.W_REG(m_ResNo).STCUT(m_CutNo).intCTYP = CNS_CUTP_L) Or (.W_REG(m_ResNo).STCUT(m_CutNo).intCTYP = CNS_CUTP_SP) Then
                    'V1.0.4.3⑦                        Select Case (.W_REG(m_ResNo).STCUT(m_CutNo).intANG) ' ｶｯﾄ方向1(90°単位　0°～360°)
                    'V1.0.4.3⑦                            Case 0, 180 ' ｶｯﾄ方向1 = 0,180°なら ｶｯﾄ方向2 = 90,270°以外エラー
                    'V1.0.4.3⑦                        If (.W_REG(m_ResNo).STCUT(m_CutNo).intANG2 = 0) Or (.W_REG(m_ResNo).STCUT(m_CutNo).intANG2 = 180) Then
                    'V1.0.4.3⑦                        strMsg = "相関チェックエラー" & vbCrLf
                    'V1.0.4.3⑦                        strMsg = strMsg & "カット方向１とカット方向２の組合わせ指定が正しくありません。"
                    'V1.0.4.3⑦                        GoTo STP_ERR
                    'V1.0.4.3⑦                    End If
                    'V1.0.4.3⑦                            Case 90, 270 ' ｶｯﾄ方向1 = 90,270°なら ｶｯﾄ方向2 = 0,180°以外エラー
                    'V1.0.4.3⑦                        If (.W_REG(m_ResNo).STCUT(m_CutNo).intANG2 = 90) Or (.W_REG(m_ResNo).STCUT(m_CutNo).intANG2 = 270) Then
                    'V1.0.4.3⑦                        strMsg = "相関チェックエラー" & vbCrLf
                    'V1.0.4.3⑦                        strMsg = strMsg & "カット方向１とカット方向２の組合わせ指定が正しくありません。"
                    'V1.0.4.3⑦                        GoTo STP_ERR
                    'V1.0.4.3⑦                    End If
                    'V1.0.4.3⑦                        End Select
                    'V1.0.4.3⑦                    End If
                    '###1042①↓
                    '---------------------------------------------------------------------------
                    '   文字マーキングの時の角度チェック
                    '---------------------------------------------------------------------------
                    If .W_REG(m_ResNo).STCUT(m_CutNo).intCTYP = CNS_CUTP_M Then
                        If .W_REG(m_ResNo).STCUT(m_CutNo).intANG Mod 90 > 0 Then
                            strMsg = "相関チェックエラー" & vbCrLf
                            strMsg = strMsg & "文字の角度は、0°,90°,180°,270°のみ有効です。"
                            '                            errIdx = CUT_DIR_1
                            Call MsgBox(strMsg, DirectCast(MsgBoxStyle.OkOnly + MsgBoxStyle.Information, MsgBoxStyle), My.Application.Info.Title)
                            CheckRelation = 1 ' Return値 = ｴﾗｰ
                            Exit Function
                        End If
                        If .W_REG(m_ResNo).STCUT(m_CutNo).dblDL2 < 0 Or 10.0 < .W_REG(m_ResNo).STCUT(m_CutNo).dblDL2 Then
                            strMsg = "相関チェックエラー" & vbCrLf
                            strMsg = strMsg & "文字の高さは、0.1mm～10.0mmの範囲で指定して下さい。"
                            errIdx = CUT_LEN_1
                            GoTo STP_ERR
                        End If
                    End If
                    '###1042①↑
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

            Exit Function
STP_ERR:
            Call MsgBox_CheckErr(DirectCast(ctlArray(errIdx), cTxt_), strMsg)
            CheckRelation = 1 ' Return値 = ｴﾗｰ

        End Function
#End Region

#Region "追加･削除ﾎﾞﾀﾝ関連処理"
        ''' <summary>ｶｯﾄﾃﾞｰﾀを追加または削除し、そのﾃﾞｰﾀを初期化する</summary>
        ''' <param name="addDel">1=Add, (-1)=Del</param>
        Private Sub SortCutData(ByVal addDel As Integer)
            Dim iStart As Integer
            Dim iEnd As Integer
            Dim dir As Integer = (-1) * addDel ' Add=(-1), Del=1にする
            Try
                With m_MainEdit.W_REG(m_ResNo)
                    If (1 = addDel) Then ' 追加の場合
                        .intTNN = Convert.ToInt16(.intTNN + 1) ' 登録ｶｯﾄ数を追加する
                        iStart = .intTNN ' 登録されているｶｯﾄ数から
                        iEnd = (m_CutNo + 1) ' 追加するｶｯﾄﾃﾞｰﾀ番号+1まで、前のﾃﾞｰﾀを後ろにずらす
                    Else ' 削除の場合
                        iStart = m_CutNo ' 削除するｶｯﾄﾃﾞｰﾀ番号から
                        iEnd = (.intTNN - 1) ' 登録されているｶｯﾄﾃﾞｰﾀ数-1まで、後ろのﾃﾞｰﾀを前にずらす
                    End If

                    For cn As Integer = iStart To iEnd Step dir
                        .STCUT(cn).intCUT = .STCUT(cn + dir).intCUT     ' ｶｯﾄ方法(1:ﾄﾗｯｷﾝｸﾞ, 2:ｲﾝﾃﾞｯｸｽ, 3:ﾎﾟｼﾞｼｮﾆﾝｸﾞ無しｲﾝﾃﾞｯｸｽ)
                        .STCUT(cn).intCTYP = .STCUT(cn + dir).intCTYP   ' ｶｯﾄ形状(1:ｽﾄﾚｰﾄ, 2:Lｶｯﾄ)
                        .STCUT(cn).intNum = .STCUT(cn + dir).intNum     ' ｶｯﾄ本数(ｻｰﾍﾟﾝﾀｲﾝｶｯﾄのみ)
                        .STCUT(cn).dblSTX = .STCUT(cn + dir).dblSTX     ' ﾄﾘﾐﾝｸﾞ ｽﾀｰﾄ点 X
                        .STCUT(cn).dblSTY = .STCUT(cn + dir).dblSTY     ' ﾄﾘﾐﾝｸﾞ ｽﾀｰﾄ点 Y
                        .STCUT(cn).dblSX2 = .STCUT(cn + dir).dblSX2     ' ﾄﾘﾐﾝｸﾞ ｽﾀｰﾄ点2 X
                        .STCUT(cn).dblSY2 = .STCUT(cn + dir).dblSY2     ' ﾄﾘﾐﾝｸﾞ ｽﾀｰﾄ点2 Y
                        .STCUT(cn).dblCOF = .STCUT(cn + dir).dblCOF     ' ｶｯﾄｵﾌ(%)
                        .STCUT(cn).intTMM = .STCUT(cn + dir).intTMM     ' ﾓｰﾄﾞ(0:高速(ｺﾝﾊﾟﾚｰﾀ非積分ﾓｰﾄﾞ), 1:高精度(積分ﾓｰﾄﾞ))
                        .STCUT(cn).intMType = .STCUT(cn + dir).intMType ' 内部／外部測定器
                        .STCUT(cn).intQF1 = .STCUT(cn + dir).intQF1     ' Qﾚｰﾄ(0.1KHz)
                        .STCUT(cn).dblV1 = .STCUT(cn + dir).dblV1       ' ﾄﾘﾑ速度(mm/s)
                        .STCUT(cn).intQF2 = .STCUT(cn + dir).intQF2     ' V1.0.4.3③ストレートカット・リトレースのQレート(0.1KHz)に使用
                        .STCUT(cn).dblV2 = .STCUT(cn + dir).dblV2       ' V1.0.4.3③ストレートカット・リトレースのトリム速度(mm/s)に使用
                        .STCUT(cn).dblDL2 = .STCUT(cn + dir).dblDL2     ' 第2のｶｯﾄ長(ﾘﾐｯﾄｶｯﾄ量mm(Lﾀｰﾝ前))
                        .STCUT(cn).dblDL3 = .STCUT(cn + dir).dblDL3     ' 第3のｶｯﾄ長(ﾘﾐｯﾄｶｯﾄ量mm(Lﾀｰﾝ後))
                        .STCUT(cn).intANG = .STCUT(cn + dir).intANG     ' ｶｯﾄ方向1
                        .STCUT(cn).intANG2 = .STCUT(cn + dir).intANG2   ' ｶｯﾄ方向2
                        .STCUT(cn).dblLTP = .STCUT(cn + dir).dblLTP     ' Lﾀｰﾝ ﾎﾟｲﾝﾄ(%)
                        .STCUT(cn).cFormat = .STCUT(cn + dir).cFormat   '###1042① 文字データ
                        .STCUT(cn).cMarkFix = .STCUT(cn + dir).cMarkFix   '印字固定部 'V2.2.1.7①
                        .STCUT(cn).cMarkStartNum = .STCUT(cn + dir).cMarkStartNum   '開始番号 'V2.2.1.7①
                        .STCUT(cn).intMarkRepeatCnt = .STCUT(cn + dir).intMarkRepeatCnt   '開始番号 'V2.2.1.7①

                        'V2.1.0.0①↓ カット毎の抵抗値変化量判定機能追加
                        .STCUT(cn).iVariationRepeat = .STCUT(cn + dir).iVariationRepeat     ' リピート有無
                        .STCUT(cn).iVariation = .STCUT(cn + dir).iVariation                 ' 判定有無
                        .STCUT(cn).dRateOfUp = .STCUT(cn + dir).dRateOfUp                   ' 上昇率
                        .STCUT(cn).dVariationLow = .STCUT(cn + dir).dVariationLow           ' 下限値
                        .STCUT(cn).dVariationHi = .STCUT(cn + dir).dVariationHi             ' 上限値
                        'V2.1.0.0①↑

                        ' ｲﾝﾃﾞｯｸｽｶｯﾄ情報設定
                        For ix As Integer = 1 To MAXIDX Step 1 ' MAXｲﾝﾃﾞｯｸｽｶｯﾄ数分繰返す
                            .STCUT(cn).intIXN(ix) = .STCUT(cn + dir).intIXN(ix) ' ｶｯﾄ回数1-5
                            .STCUT(cn).dblDL1(ix) = .STCUT(cn + dir).dblDL1(ix) ' ｶｯﾄ長1-5
                            .STCUT(cn).lngPAU(ix) = .STCUT(cn + dir).lngPAU(ix) ' ﾋﾟｯﾁ間ﾎﾟｰｽﾞ1-5
                            .STCUT(cn).dblDEV(ix) = .STCUT(cn + dir).dblDEV(ix) ' 誤差1-5(%)
                            .STCUT(cn).intIXMType(ix) = .STCUT(cn + dir).intIXMType(ix) ' 測定機器
                            .STCUT(cn).intIXTMM(ix) = .STCUT(cn + dir).intIXTMM(ix)     ' 測定ﾓｰﾄﾞ
                        Next ix

                        ' FL加工条件
                        For fl As Integer = 1 To MAXCND Step 1
                            .STCUT(cn).intCND(fl) = .STCUT(cn + dir).intCND(fl) ' FL設定No.
                        Next fl

                        'V1.0.4.3③ ADD ↓
                        Dim i As Integer
                        For i = 1 To MAX_LCUT
                            .STCUT(cn).dCutLen(i) = .STCUT(cn + dir).dCutLen(i)         ' カット長１～７　リターン時も使用
                            .STCUT(cn).dQRate(i) = .STCUT(cn + dir).dQRate(i)           ' Ｑレート１～７　リターン時も使用
                            .STCUT(cn).dSpeed(i) = .STCUT(cn + dir).dSpeed(i)           ' 速度１～７
                            .STCUT(cn).dAngle(i) = .STCUT(cn + dir).dAngle(i)           ' 角度１～７
                            .STCUT(cn).dTurnPoint(i) = .STCUT(cn + dir).dTurnPoint(i)   ' ターンポイント１～６
                        Next
                        'V1.0.4.3③ ADD ↑

                        'V2.0.0.0⑦ ADD ↓
                        .STCUT(cn).intRetraceCnt = .STCUT(cn + dir).intRetraceCnt       ' リトレースカット本数
                        For i = 1 To MAX_LCUT
                            .STCUT(cn).dblRetraceOffX(i) = .STCUT(cn + dir).dblRetraceOffX(i)       ' リトレースのオフセットＸ
                            .STCUT(cn).dblRetraceOffY(i) = .STCUT(cn + dir).dblRetraceOffY(i)       ' リトレースのオフセットＹ
                            .STCUT(cn).dblRetraceQrate(i) = .STCUT(cn + dir).dblRetraceQrate(i)     ' ストレートカット・リトレースのQレート(0.1KHz)に使用
                            .STCUT(cn).dblRetraceSpeed(i) = .STCUT(cn + dir).dblRetraceSpeed(i)     ' ストレートカット・リトレースのトリム速度(mm/s)に使用
                        Next
                        'V2.0.0.0⑦ ADD ↑

                        'V2.2.0.0②↓
                        'Uカットパラメータの追加
                        .STCUT(cn).dUCutL1 = 0.0          ' L1
                        .STCUT(cn).dUCutL2 = 0.0          ' L2
                        .STCUT(cn).intUCutQF1 = 0.1       ' Qレート
                        .STCUT(cn).dblUCutV1 = 0.1        ' 速度
                        .STCUT(cn).intUCutANG = 0         ' 角度
                        .STCUT(cn).dblUCutTurnP = 0       ' ターンポイント
                        .STCUT(cn).intUCutTurnDir = 1     ' ターン方向
                        .STCUT(cn).dblUCutR1 = 0          ' R1
                        .STCUT(cn).dblUCutR2 = 0          ' R2
                        'V2.2.0.0②↑

                    Next cn

                    ' つめて不要となったﾃﾞｰﾀを初期化する
                    If (1 = addDel) Then ' 追加の場合
                        Call InitCutData(m_ResNo, m_CutNo) ' 追加したﾃﾞｰﾀを初期化
                    Else ' 削除の場合
                        Call InitCutData(m_ResNo, .intTNN) ' 最後のﾃﾞｰﾀを初期化
                        .intTNN = Convert.ToInt16(.intTNN - 1) ' 登録ｶｯﾄ数を-1する

                        ' 最終ｶｯﾄの削除なら現在のｶｯﾄ番号を最終ｶｯﾄ番号とする
                        If (.intTNN < m_CutNo) Then m_CutNo = .intTNN
                    End If
                End With

                ' ｶｯﾄﾃﾞｰﾀを画面項目に設定
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
            Dim Cnt As Integer      'V1.0.4.3②
            Dim ctlIdx As Integer   'V1.0.4.3③
            Try
                cCombo = DirectCast(sender, cCmb_)
                tag = DirectCast(cCombo.Tag, Integer)
                idx = cCombo.SelectedIndex
                With m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo)
                    Select Case (DirectCast(cCombo.Parent.Tag, Integer)) ' ｸﾞﾙｰﾌﾟﾎﾞｯｸｽのﾀｸﾞ
                        ' ------------------------------------------------------------------------------
                        Case 0 ' ｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            Select Case (tag)
                                Case 0 ' 抵抗番号
                                    m_ResNo = (idx + 1) ' この代入でm_CutNoが1になる
                                    ' 対応するﾃﾞｰﾀをﾃｷｽﾄﾎﾞｯｸｽ、ｶｯﾄ番号ｺﾝﾎﾞﾎﾞｯｸｽにｾｯﾄする
                                    Call SetDataToText() ' ｶｯﾄ番号は1を設定する
                                Case 1 ' ｶｯﾄ番号
                                    m_CutNo = (idx + 1)
                                    ' 対応するﾃﾞｰﾀをﾃｷｽﾄﾎﾞｯｸｽ、ｶｯﾄ番号ｺﾝﾎﾞﾎﾞｯｸｽにｾｯﾄする
                                    Call SetDataToText()

                                Case 2 ' ｶｯﾄ方法(1:ﾄﾗｯｷﾝｸﾞ, 2:ｲﾝﾃﾞｯｸｽ, 3:NGｶｯﾄ)
                                    .intCUT = Convert.ToInt16(idx + 1)
                                    Call ChangedCutMethod(idx + 1) ' 関連ｺﾝﾄﾛｰﾙの有効･無効を設定

                                    ' NGｶｯﾄまたは測定機器が外部測定器の場合測定ﾓｰﾄﾞを無効にする
                                    If (CNS_CUTM_NG = .intCUT) OrElse (0 < .intMType) Then
                                        m_CtlCut(CUT_TMM).Enabled = False ' 測定ﾓｰﾄﾞ無効
                                    Else
                                        m_CtlCut(CUT_TMM).Enabled = True ' 測定ﾓｰﾄﾞ有効
                                    End If

                                    If (CNS_CUTM_TR = .intCUT Or CNS_CUTM_NG = .intCUT) Then       ' トラッキングまたはNGカットの場合は、内部測定のみ
                                        .intMType = 0
                                        m_CtlCut(CUT_MTYPE).Enabled = False
                                        Call SetCutData()
                                    Else
                                        m_CtlCut(CUT_MTYPE).Enabled = True
                                    End If

                                    'V2.0.0.5①                                    'V2.0.0.0⑮ インデックスでトラッキングはLカット無効
                                    'V2.0.0.5①                                    If m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCUT = CNS_CUTM_IX And m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCTYP = CNS_CUTP_L Then
                                    'V2.0.0.5①                                        CGrp_1.Enabled = False                      ' インデックス
                                    'V2.0.0.5①                                        CGrp_3.Enabled = False                       'Ｌカットパラメータ
                                    'V2.0.0.5①                                    End If
                                    'V2.0.0.5①                                    'V2.0.0.0⑮↑

                                Case 3 ' ｶｯﾄ形状(1:ｽﾄﾚｰﾄ, 2:Lｶｯﾄ, 3:ｻｰﾍﾟﾝﾀｲﾝ)
                                    .intCTYP = GetComboBoxName2Value(cCombo.Text, Me.m_lstCutType)
                                    Call ChangedCutShape(.intCTYP) ' 関連ｺﾝﾄﾛｰﾙの表示･非表示を設定
                                    'V2.0.0.5①                                    'V2.0.0.0⑮↓インデックスでトラッキングはLカット無効
                                    'V2.0.0.5①                                    If m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCUT = CNS_CUTM_IX And m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCTYP = CNS_CUTP_L Then
                                    'V2.0.0.5①                                        CGrp_1.Enabled = False                      ' インデックス
                                    'V2.0.0.5①                                        CGrp_3.Enabled = False                       'Ｌカットパラメータ
                                    'V2.0.0.5①                                    End If
                                    'V2.0.0.5①                                    'V2.0.0.0⑮↑
                                Case 4, 5 ' ｶｯﾄ方向1, ｶｯﾄ方向2
                                    Dim iWK As Short = 0
                                    'V1.0.4.3② コメント化↓
                                    'Select Case (idx) ' ｶｯﾄ方向(90°単位　0°～360°)
                                    '    Case 0 : iWK = 0    ' 0°
                                    '    Case 1 : iWK = 90   ' 90°
                                    '    Case 2 : iWK = 180  ' 180°
                                    '    Case 3 : iWK = 270  ' 270°
                                    '    Case 4 : iWK = 10   ' 10°
                                    '    Case 5 : iWK = 20   ' 20°
                                    '    Case 6 : iWK = 30   ' 30°
                                    '    Case 7 : iWK = 40   ' 40°
                                    '    Case 8 : iWK = 50   ' 50°
                                    '    Case 9 : iWK = 60   ' 60°
                                    '    Case 10 : iWK = 70  ' 70°
                                    '    Case 11 : iWK = 80  ' 80°
                                    '    Case 12 : iWK = 100 ' 100°
                                    '    Case 13 : iWK = 110 ' 110°
                                    '    Case 14 : iWK = 120 ' 120°
                                    '    Case 15 : iWK = 130 ' 130°
                                    '    Case 16 : iWK = 140 ' 140°
                                    '    Case 17 : iWK = 150 ' 150°
                                    '    Case 18 : iWK = 160 ' 160°
                                    '    Case 19 : iWK = 170 ' 170°
                                    '    Case Else ' DO NOTHING
                                    'End Select
                                    'V1.0.4.3②↓
                                    For Cnt = 0 To MAX_DEGREES
                                        If AngleArray(Cnt, 1) = idx Then
                                            iWK = AngleArray(Cnt, 0)
                                            Exit For
                                        End If
                                    Next
                                    'V1.0.4.3②↑
                                    ' 編集ﾃﾞｰﾀのｶｯﾄ方向を設定する
                                    If (4 = tag) Then ' ｶｯﾄ方向1(90°単位　0°～360°)
                                        .intANG = iWK
                                    Else ' ｶｯﾄ方向2(90°単位　0°～360°)
                                        .intANG2 = iWK
                                    End If

                                Case 6 ' 測定機器(0:内部測定器, 1以上は外部測定器番号)
                                    ' 登録されている測定機器ﾘｽﾄの数値を設定する( 1:NAME=1, 10:NAME=10)
                                    .intMType = Convert.ToInt16((cCombo.Text).Substring(0, 2))

                                    ' 測定機器が外部測定器の場合測定ﾓｰﾄﾞを無効にする
                                    If (0 < idx) Then
                                        m_CtlCut(CUT_TMM).Enabled = False ' 測定ﾓｰﾄﾞ無効
                                    Else
                                        m_CtlCut(CUT_TMM).Enabled = True ' 測定ﾓｰﾄﾞ有効
                                    End If

                                Case 7 ' 測定ﾓｰﾄﾞ
                                    .intTMM = Convert.ToInt16(idx)
                                    'V2.1.0.0①↓
                                Case 8 ' カット毎の抵抗値変化量判定機能リピート有無
                                    Dim bChange As Boolean = False
                                    If .iVariationRepeat <> Convert.ToInt16(idx) Then
                                        ' 「あり」から「なし」に変わった時もなし全てする為にコピーする。
                                        bChange = True
                                    End If
                                    .iVariationRepeat = Convert.ToInt16(idx)
                                    If .iVariationRepeat = 1 OrElse bChange Then
                                        UserSub.CutVariationDataCopy(m_MainEdit.W_PLT, m_MainEdit.W_REG, m_ResNo, m_CutNo)
                                    End If
                                Case 9 ' カット毎の抵抗値変化量判定機能判定有無
                                    .iVariation = Convert.ToInt16(idx)
                                    If .iVariationRepeat = 1 Then
                                        UserSub.CutVariationDataCopy(m_MainEdit.W_PLT, m_MainEdit.W_REG, m_ResNo, m_CutNo)
                                    End If
                                    'V2.1.0.0①↑
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case 1 ' ｲﾝﾃﾞｯｸｽｶｯﾄｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            ctlIdx = GetCtlIdx(cCombo, tag) ' 1次元目のｲﾝﾃﾞｯｸｽ

                            Select Case (tag)
                                Case 0 ' 測定機器(0:内部測定器, 1以上は外部測定器番号)
                                    ' 登録されている測定機器ﾘｽﾄの数値を設定する( 1:NAME=1, 10:NAME=10)
                                    .intIXMType(ctlIdx + 1) = Convert.ToInt16((cCombo.Text).Substring(0, 2))
                                    ' 外部測定器の場合測定ﾓｰﾄﾞを無効にする
                                    If (0 < idx) Then
                                        m_CtlIdxCut(ctlIdx, IDX_TMM).Enabled = False
                                    Else
                                        m_CtlIdxCut(ctlIdx, IDX_TMM).Enabled = True
                                    End If

                                Case 1 ' 測定ﾓｰﾄﾞ(0:高速, 1:高精度)
                                    .intIXTMM(ctlIdx + 1) = Convert.ToInt16(idx)
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                        Case 2 ' FL加工条件ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ
                            Select Case (tag)
                                Case 0 ' FL設定No.(0～31)
                                    Dim cndNo As Integer ' ｶｯﾄ条件番号
                                    For i As Integer = 0 To (m_CtlFLCnd.GetLength(0) - 1) Step 1
                                        If (m_CtlFLCnd(i, 0) Is cCombo) Then
                                            cndNo = (i + 1)
                                            Exit For
                                        End If
                                    Next i
                                    .intCND(cndNo) = Convert.ToInt16(idx)

                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                            'V1.0.4.3③↓
                            ' ------------------------------------------------------------------------------
                        Case 3 ' Ｌカットパラメータグループボックス
                            ctlIdx = GetCtlLCutIdx(cCombo, tag) ' 1次元目のｲﾝﾃﾞｯｸｽ
                            Select Case (tag)
                                Case 0 ' カット長
                                    For Cnt = 0 To MAX_DEGREES
                                        If AngleArrayForLcut(Cnt, 1) = idx Then
                                            .dAngle(ctlIdx + 1) = AngleArrayForLcut(Cnt, 0)
                                            Exit For
                                        End If
                                    Next
                                    If ctlIdx = 0 Then
                                        .intANG = .dAngle(ctlIdx + 1)
                                    End If
                                    If .dAngle(2) >= .dAngle(1) Then
                                        If .dAngle(2) - .dAngle(1) > 180.0 Then
                                            .intANG2 = DEF_DIR_CW
                                        Else
                                            .intANG2 = DEF_DIR_CCW
                                        End If
                                    Else
                                        If .dAngle(1) - .dAngle(2) > 180.0 Then
                                            .intANG2 = DEF_DIR_CCW
                                        Else
                                            .intANG2 = DEF_DIR_CW
                                        End If
                                    End If
                                Case Else
                                    Throw New Exception("Case " & tag & ": Nothing")
                            End Select
                            ' ------------------------------------------------------------------------------
                            'V1.0.4.3③↑
                            'V2.2.0.0②↓
                        Case 5 ' Ｕカットパラメータグループボックス

                            ctlIdx = GetCtlLCutIdx(cCombo, tag) ' 1次元目のｲﾝﾃﾞｯｸｽ
                            Select Case (tag)
                                Case 0 ' カット角度
                                    For Cnt = 0 To MAX_DEGREES
                                        If AngleArrayForUcut(Cnt, 1) = idx Then
                                            .intUCutANG = AngleArrayForUcut(Cnt, 0)
                                            Exit For
                                        End If
                                    Next

                                Case 1 ' ターン方向
                                    .intUCutTurnDir = Convert.ToInt16(idx) + 1
                            End Select
                            'V2.2.0.0②↑

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
            Dim refOpt As Short ' ｵﾌﾟｼｮﾝﾎﾞﾀﾝ(0=前に追加 ,1=後に追加)
            Dim ret As Integer
            Try
                ' 登録数ﾁｪｯｸ
                With m_MainEdit
                    If (.W_PLT.RCount < 1) Then Exit Sub ' 抵抗ﾃﾞｰﾀなしならNOP
                    If (MAXCTN <= .W_REG(m_ResNo).intTNN) Then ' カット数 >= 9 ?
                        Dim strMsg As String = "これ以上カットデータは登録できません。"
                        Call MsgBox(strMsg, DirectCast( _
                                    MsgBoxStyle.OkOnly + _
                                    MsgBoxStyle.Information, MsgBoxStyle), _
                                    My.Application.Info.Title)
                        Exit Sub
                    End If
                End With

                ' 確認ﾒｯｾｰｼﾞを表示("カットデータを追加します")
                ret = MsgBox_AddClick("カットデータ", refOpt) ' ﾒｯｾｰｼﾞ表示
                If (ret <> cFRS_ERR_ADV) Then Exit Sub ' CancelならReturn
                If (refOpt = 1) Then ' 表示ﾃﾞｰﾀの後に追加 ?
                    m_CutNo = (m_CutNo + 1) ' 追加するﾃﾞｰﾀの番号 = 現在のﾃﾞｰﾀ番号 + 1
                Else ' 表示ﾃﾞｰﾀの前に追加
                    m_CutNo = m_CutNo ' 追加するﾃﾞｰﾀの番号 = 現在のﾃﾞｰﾀ番号
                End If

                ' ﾃﾞｰﾀを1個後にずらして追加する
                Call SortCutData(1)

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
                If (1 = m_MainEdit.W_REG(m_ResNo).intTNN) Then Exit Sub ' 抵抗内カット数1ならNOP
                strMsg = "現在のカットデータを削除します。よろしいですか？"
                ret = MsgBox(strMsg, DirectCast( _
                            MsgBoxStyle.OkCancel + _
                            MsgBoxStyle.Exclamation + _
                            MsgBoxStyle.DefaultButton2, MsgBoxStyle), _
                            My.Application.Info.Title)

                If (ret = MsgBoxResult.Cancel) Then Exit Sub ' Cancel(RESETｷｰ) ?

                ' 後ろのﾃﾞｰﾀを1個前につめる
                Call SortCutData(-1)

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

        ''' <summary>加工条件ﾎﾞﾀﾝｸﾘｯｸ時の処理</summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Private Sub CBtn_FLS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBtn_FLC.Click
            Dim ret As Integer
            Try
#If cOSCILLATORcFLcUSE Then

                Dim fls As Process = Process.Start("C:\TRIM\FLSetup.exe") ' FLSetup.exeを起動する
                fls.WaitForExit() ' 終了を待つ

                ' FL側から現在の加工条件をﾒｲﾝ画面の加工条件ﾃﾞｰﾀに受信する
                ret = TrimCondInfoRcv(stCND)
                If (0 <> ret) Then ' ｴﾗｰの場合
                    Dim strMsg As String = "ＦＬ側加工条件のリードに失敗しました。"
                    Call MsgBox(strMsg, DirectCast( _
                                MsgBoxStyle.OkOnly + _
                                MsgBoxStyle.Critical, MsgBoxStyle), _
                                My.Application.Info.Title)
                    Exit Sub
                End If

                ' ﾃﾞｰﾀを受信し、ﾒｲﾝ画面の加工条件ﾃﾞｰﾀが更新された場合
                Call m_MainEdit.ReadFlConditionData() ' 編集画面のFL加工条件ﾃﾞｰﾀを更新する
                Call SetFLCndData() ' ｶｯﾄﾀﾌﾞFL加工条件部分の表示を更新
                Me.Refresh()
#Else
                ret = 0
#End If

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

        'V2.0.0.0↓
#Region "ｶｯﾄ方法の初期設定"
        ''' <summary>
        ''' ｶｯﾄ方法の初期設定
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub InitCutMethodData()
            Dim cd As New ComboDataStruct

            cd.SetData("トラッキング", CNS_CUTM_TR)
            m_lstCutMethod.Add(cd)
            cd.SetData("インデックス", CNS_CUTM_IX)
            m_lstCutMethod.Add(cd)
#If cFORCEcCUT Then
            cd.SetData("強制カット", CNS_CUTM_FC)
            m_lstCutMethod.Add(cd)
#End If
        End Sub
#End Region

#Region "カット形状の初期設定"
        ''' <summary>
        ''' カット形状の初期設定
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub InitCutTypeData()
            Try
                Dim ctyp As New ComboDataStruct

                ctyp.SetData("ストレート", CNS_CUTP_ST)
                m_lstCutType.Add(ctyp)
                ctyp.SetData("リトレース", CNS_CUTP_ST_TR)
                m_lstCutType.Add(ctyp)
                ctyp.SetData("Ｌカット", CNS_CUTP_L)
                m_lstCutType.Add(ctyp)
                ctyp.SetData("文字", CNS_CUTP_M)
                m_lstCutType.Add(ctyp)
                ctyp.SetData("Ｕカット", CNS_CUTP_U)        'V2.2.0.0② 
                m_lstCutType.Add(ctyp)                      'V2.2.0.0② 


            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try
        End Sub
#End Region

#Region "抵抗の判定モードが変更された時にインデックスの誤差の設定を変更する"
        ''' <summary>抵抗の判定モードが変更された時にインデックスの誤差の設定を変更する</summary>
        ''' <param name="nJudge">判定モード(0:比率(%), 1:数値(絶対値))</param>
        ''' <param name="ctlText">テキストコントロール</param>
        ''' <param name="strUnit">単位</param>
        Private Sub ChangedJudge(ByVal nJudge As Integer, ByVal ctlText As Control, ByVal strUnit As String)
            Dim strMin As String
            Dim strMax As String

            Try
                If nJudge = JUDGE_MODE_RATIO Then ' 比率
                    strMin = m_strDEVRaite(0)
                    strMax = m_strDEVRaite(1)

                    ' 誤差ラベルの変更
                    CLbl_20.Text = String.Format("{0}・比率(%)", m_strDev)
                Else
                    strMin = m_strDEVAbsolute(0)
                    strMax = m_strDEVAbsolute(1)

                    ' 誤差ラベルの変更
                    CLbl_20.Text = String.Format("{0}・絶対値({1})", m_strDev, strUnit)
                End If

                If TypeOf ctlText Is cTxt_ Then
                    With DirectCast(ctlText, cTxt_) ' 誤差ﾃｷｽﾄﾎﾞｯｸｽ
                        Call .SetMinMax(strMin, strMax) ' 下限値･上限値の設定
                        Call .SetStrTip(strMin & "～" & strMax & "の範囲で指定して下さい") ' ﾂｰﾙﾁｯﾌﾟﾒｯｾｰｼﾞの設定
                    End With
                End If

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region
        'V2.0.0.0↑

        ''' <summary>
        ''' 角度のコンボボックスに新たに設定された角度を追加する。
        ''' </summary>
        ''' <param name="IDegree"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Add_CCmb_4_Item(ByRef IDegree As Integer) As Integer
            Dim Cnt As Integer

            Try
                If AngleArray(IDegree, 1) = -1 Then
                    CCmb_4.Items.Add("   " + IDegree.ToString("0") + "°")
                    Dim iMax As Integer = 0
                    For Cnt = 0 To MAX_DEGREES
                        If AngleArray(Cnt, 1) > iMax Then
                            iMax = AngleArray(Cnt, 1)
                        End If
                    Next
                    AngleArray(IDegree, 1) = iMax + 1
                End If
                CCmb_4.SelectedIndex = AngleArray(IDegree, 1)

                Return (CCmb_4.SelectedIndex)
            Catch ex As Exception
                MsgBox("CCmb_4_KeyDown() TRAP ERROR = " + ex.Message)
            End Try

        End Function

        Private Sub CCmb_4_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CCmb_4.KeyDown
            Dim iDegree As Integer

            ' Try       V2.2.0.0⑨
            If e.KeyValue <> Keys.Enter Then
                Exit Sub
            End If
            If Not IsNumeric(CCmb_4.Text) Then
                Exit Sub
            End If

            ' ↓V2.2.0.0⑨
            Try
                iDegree = Integer.Parse(CCmb_4.Text)
            Catch
                Call MsgBox("整数を入力してください", DirectCast(
                                                MsgBoxStyle.OkOnly +
                                                MsgBoxStyle.Information, MsgBoxStyle),
                                                My.Application.Info.Title)
                Add_CCmb_4_Item(m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intANG)
                Exit Sub
            End Try

            Try
                ' ↑V2.2.0.0⑨
                If m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intCTYP = CNS_CUTP_M And (iDegree Mod 90 > 0) Then
                    Call MsgBox("文字の角度は、0°,90°,180°,270°のみ有効です。", DirectCast(
                                                MsgBoxStyle.OkOnly +
                                                MsgBoxStyle.Information, MsgBoxStyle),
                                                My.Application.Info.Title)
                    Add_CCmb_4_Item(m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intANG)
                    Exit Sub
                End If

                If iDegree < 0 Or MAX_DEGREES < iDegree Then
                    Call MsgBox("0～359の範囲で指定して下さい", DirectCast(
                                                MsgBoxStyle.OkOnly +
                                                MsgBoxStyle.Information, MsgBoxStyle),
                                                My.Application.Info.Title)
                    Add_CCmb_4_Item(m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intANG)
                    Exit Sub
                End If

                Call Add_CCmb_4_Item(iDegree)

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

        'V1.0.4.3⑦↓
        ''' <summary>
        ''' 角度のコンボボックスに新たに設定された角度を追加する。Ｌカット用
        ''' </summary>
        ''' <param name="IDegree"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        ''' 
        Private Function Add_CCmb_Dir_X_Item(ByRef IDegree As Integer, ByRef cCombo As cCmb_) As Integer
            Dim Cnt As Integer

            Try
                If AngleArrayForLcut(IDegree, 1) = -1 Then
                    Dim cCombo2 As cCmb_
                    For Cnt = 0 To (m_CtlLCut.GetLength(0) - 1) Step 1
                        cCombo2 = DirectCast(m_CtlLCut(Cnt, 3), cCmb_)
                        cCombo2.Items.Add("   " + IDegree.ToString("0") + "°")
                    Next Cnt
                    Dim iMax As Integer = 0
                    For Cnt = 0 To MAX_DEGREES
                        If AngleArrayForLcut(Cnt, 1) > iMax Then
                            iMax = AngleArrayForLcut(Cnt, 1)
                        End If
                    Next
                    AngleArrayForLcut(IDegree, 1) = iMax + 1
                End If
                cCombo.SelectedIndex = AngleArrayForLcut(IDegree, 1)

                Return (cCombo.SelectedIndex)
            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Function

        ''' <summary>
        ''' Ｌカットパラメータの角度入力イベント処理
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub CCmb_Dir_1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CCmb_Dir_7.KeyDown, CCmb_Dir_6.KeyDown, CCmb_Dir_5.KeyDown, CCmb_Dir_4.KeyDown, CCmb_Dir_3.KeyDown, CCmb_Dir_2.KeyDown, CCmb_Dir_1.KeyDown
            Dim iDegree As Integer

            ' Try       V2.2.0.0⑨
            If e.KeyValue <> Keys.Enter Then
                Exit Sub
            End If
            If Not IsNumeric(sender.Text) Then
                Exit Sub
            End If

            ' ↓V2.2.0.0⑨
            Try
                iDegree = Integer.Parse(sender.Text)
            Catch
                Call MsgBox("整数を入力してください", DirectCast(
                                                MsgBoxStyle.OkOnly +
                                                MsgBoxStyle.Information, MsgBoxStyle),
                                                My.Application.Info.Title)
                Add_CCmb_Dir_X_Item(m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intANG, sender)
                Exit Sub
            End Try

            Try
                ' ↑V2.2.0.0⑨

                If iDegree < 0 Or MAX_DEGREES < iDegree Then
                    Call MsgBox("0～359の範囲で指定して下さい", DirectCast(
                                                MsgBoxStyle.OkOnly +
                                                MsgBoxStyle.Information, MsgBoxStyle),
                                                My.Application.Info.Title)
                    Add_CCmb_Dir_X_Item(m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo).intANG, sender)
                    Exit Sub
                End If

                Call Add_CCmb_Dir_X_Item(iDegree, sender)

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
        'V1.0.4.3⑦↑


        'V2.2.0.0②↓
        ''' <summary>
        '''          Ｕカットパラメータの追加
        ''' </summary>
        Private Sub SetUCutParamData()

            Try
                With m_MainEdit.W_REG(m_ResNo).STCUT(m_CutNo)
                    For i As Integer = 0 To (m_CtlUCut.GetLength(0) - 1) Step 1
                        Select Case (i)
                            Case 0 ' L1カット長
                                m_CtlUCut(i).Text = (.dUCutL1).ToString(DirectCast(m_CtlUCut(i), cTxt_).GetStrFormat())
                            Case 1 ' L2カット長
                                m_CtlUCut(i).Text = (.dUCutL2).ToString(DirectCast(m_CtlUCut(i), cTxt_).GetStrFormat())
                            Case 2 ' Ｒ１半径
                                m_CtlUCut(i).Text = (.dblUCutR1).ToString(DirectCast(m_CtlUCut(i), cTxt_).GetStrFormat())
                            Case 3 ' Ｒ２半径
                                m_CtlUCut(i).Text = (.dblUCutR2).ToString(DirectCast(m_CtlUCut(i), cTxt_).GetStrFormat())
                            Case 4 ' Ｑレート
                                m_CtlUCut(i).Text = (.intUCutQF1 / 10.0).ToString(DirectCast(m_CtlUCut(i), cTxt_).GetStrFormat())
                            Case 5 ' 速度
                                m_CtlUCut(i).Text = (.dblUCutV1).ToString(DirectCast(m_CtlUCut(i), cTxt_).GetStrFormat())
                            Case 6 ' 角度
                                Dim index As Integer
                                Select Case (.intUCutANG)
                                    Case 0
                                        index = 0   ' 0°
                                    Case 90
                                        index = 1   ' 90°
                                    Case 180
                                        index = 2   ' 180°
                                    Case 270
                                        index = 3   ' 270°
                                    Case Else
                                        index = Add_CCmb_Dir_X_Item(.dAngle(i + 1), DirectCast(m_CtlUCut(i), cCmb_))
                                End Select
                                Call NoEventIndexChange(DirectCast(m_CtlUCut(i), cCmb_), index)
                            Case 7 ' ターン方向 
                                Dim index As Integer
                                Select Case (.intUCutTurnDir)
                                    Case 1
                                        index = 0   ' ＣＷ
                                    Case Else
                                        index = 1   ' ＣＣＷ
                                End Select
                                Call NoEventIndexChange(DirectCast(m_CtlUCut(i), cCmb_), index)

                            Case 8 ' ターンポイント
                                m_CtlUCut(i).Text = (.dblUCutTurnP.ToString(DirectCast(m_CtlUCut(i), cTxt_).GetStrFormat()))

                            Case Else
                                Throw New Exception("i = " & i & ", Case Else")
                        End Select
                    Next i
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub


    End Class
End Namespace
