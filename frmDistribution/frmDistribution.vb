'===============================================================================
'   Description  : 分布図表示処理
'
'   Copyright(C) : TOWA LASERFRONT CORP. 2010
'
'===============================================================================
Option Strict Off
Option Explicit On
Friend Class frmDistribution
    Inherits System.Windows.Forms.Form
#Region "プライベート定数定義"
    '===========================================================================
    '   定数定義
    '===========================================================================
    ''----- 画面ｺﾋﾟ ｰ----
    '' ｷｰｽﾄﾛｰｸをｼｭﾐﾚｰﾄする
    'Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Integer, ByVal dwExtraInfo As Integer)
    'Private Const VK_SNAPSHOT As Short = &H2CS          ' PrtSc key
    'Private Const VK_LMENU As Short = &HA4S             ' Alt key
    'Private Const KEYEVENTF_KEYUP As Short = &H2S       ' ｷｰはUP状態
    'Private Const KEYEVENTF_EXTENDEDKEY As Short = &H1S ' ｽｷｬﾝは拡張ｺｰﾄﾞ

    ' 画面表示位置オフセット
    'Private Const DISP_X_OFFSET As Integer = 4                         '###065
    'Private Const DISP_Y_OFFSET As Integer = 20                        '###065
    Private Const DISP_X_OFFSET As Integer = 0                          '###065
    Private Const DISP_Y_OFFSET As Integer = 0                          '###065

#End Region

#Region "メンバ変数定義"
    '===========================================================================
    '   メンバ変数定義
    '===========================================================================
    Private m_bInitDistForm As Boolean
    Private m_bFgDispGrp As Boolean                                ' 表示ｸﾞﾗﾌ種別(TRUE:IT FALSE:FT)

    Private giDistributionResNo As Integer                       ' 分布図表示抵抗番号
    Private Const MAX_SCALE_NUM As Integer = 999999999           ' ｸﾞﾗﾌ最大値
    Private Const MAX_SCALE_RNUM As Integer = 12                 ' ｸﾞﾗﾌ表示抵抗数

    Private dblAverage(MAXRNO) As Double                         ' 平均値
    Private dblDeviationIT(MAXRNO) As Double                     ' 標準偏差(IT)
    Private dblDeviationFT(MAXRNO) As Double                     ' 標準偏差(FT)
    Private dblAverageIT(MAXRNO) As Double                       ' IT平均値
    Private dblAverageFT(MAXRNO) As Double                       ' FT平均値
    Private glRegistNum(MAX_SCALE_RNUM) As Integer               ' 分布グラフ抵抗数
    Private glRegistNumIT(MAXRNO, MAX_SCALE_RNUM) As Integer     ' 分布グラフ抵抗数 ｲﾆｼｬﾙﾃｽﾄ
    Private glRegistNumFT(MAXRNO, MAX_SCALE_RNUM) As Integer     ' 分布グラフ抵抗数 ﾌｧｲﾅﾙﾃｽﾄ

    Private dblMinIT(MAXRNO) As Double                           ' 最小値ｲﾆｼｬﾙ
    Private dblMaxIT(MAXRNO) As Double                           ' 最大値ｲﾆｼｬﾙ
    Private dblMinFT(MAXRNO) As Double                           ' 最小値ﾌｧｲﾅﾙ
    Private dblMaxFT(MAXRNO) As Double                           ' 最大値ﾌｧｲﾅﾙ
    Private dblOKRateIT(MAXRNO) As Double                        ' 良品率ｲﾆｼｬﾙ
    Private dblNGRateIT(MAXRNO) As Double                        ' 不良品率ｲﾆｼｬﾙ
    Private dblOKRateFT(MAXRNO) As Double                        ' 良品率ﾌｧｲﾅﾙ
    Private dblNGRateFT(MAXRNO) As Double                        ' 不良品率ﾌｧｲﾅﾙ

    Private gDistRegNumLblAry(MAX_SCALE_RNUM) As System.Windows.Forms.Label  ' 分布グラフ抵抗数配列
    Private gDistGrpPerLblAry(MAX_SCALE_RNUM) As System.Windows.Forms.Label  ' 分布グラフ%配列
    Private gDistShpGrpLblAry(MAX_SCALE_RNUM) As System.Windows.Forms.Label  ' 分布グラフ配列

    Private gITNx_cnt(MAXRNO) As Integer                         'IT 算出用ﾜｰｸ数
    Private gITNg_cnt(MAXRNO) As Integer                         'IT NG数記録
    Private gFTNx_cnt(MAXRNO) As Integer                         'FT 算出用ﾜｰｸ数
    Private gFTNg_cnt(MAXRNO) As Integer                         'FT NG数記録

    Public TotalFT(MAXRNO) As Double                            ' FT 合計
    Public TotalIT(MAXRNO) As Double                            ' IT 合計
    Public TotalSum2FT(MAXRNO) As Double                        ' FT２乗和 
    Public TotalSum2IT(MAXRNO) As Double                        ' IT２乗和
#End Region

    ''V2.2.0.0⑯ 
    '' 集計データ保存用 
    'Structure TOTAL_DATA_MULTI

    '    <VBFixedArray(MAX_RES_USER)> Dim gITNx_cnt() As Integer     ' IT 算出用ﾜｰｸ数
    '    <VBFixedArray(MAX_RES_USER)> Dim gITNg_cnt() As Integer     ' IT NG数記録
    '    <VBFixedArray(MAX_RES_USER)> Dim gFTNx_cnt() As Integer     ' FT 算出用ﾜｰｸ数
    '    <VBFixedArray(MAX_RES_USER)> Dim gFTNg_cnt() As Integer     ' FT NG数記録
    '    <VBFixedArray(MAX_RES_USER)> Dim dblAverage() As Double     ' 平均値
    '    <VBFixedArray(MAX_RES_USER)> Dim dblDeviationIT() As Double ' 標準偏差(IT)
    '    <VBFixedArray(MAX_RES_USER)> Dim dblDeviationFT() As Double ' 標準偏差(FT)
    '    <VBFixedArray(MAX_RES_USER)> Dim dblAverageIT() As Double   ' IT平均値
    '    <VBFixedArray(MAX_RES_USER)> Dim dblAverageFT() As Double   ' FT平均値
    '    <VBFixedArray(MAX_RES_USER)> Dim TotalIT() As Double        ' IT 合計
    '    <VBFixedArray(MAX_RES_USER)> Dim TotalFT() As Double        ' FT 合計
    '    <VBFixedArray(MAX_RES_USER)> Dim TotalSum2IT() As Double    ' IT２乗和
    '    <VBFixedArray(MAX_RES_USER)> Dim TotalSum2FT() As Double    ' FT２乗和
    '    <VBFixedArray(MAX_RES_USER)> Dim dblMinIT() As Double       ' IT最小値ﾌｧｲﾅﾙ
    '    <VBFixedArray(MAX_RES_USER)> Dim dblMaxIT() As Double       ' IT最大値ﾌｧｲﾅﾙ
    '    <VBFixedArray(MAX_RES_USER)> Dim dblMinFT() As Double       ' FT最小値ﾌｧｲﾅﾙ
    '    <VBFixedArray(MAX_RES_USER)> Dim dblMaxFT() As Double       ' FT最大値ﾌｧｲﾅﾙ
    '    <VBFixedArray(MAX_RES_USER)> Dim TrimCounter() As Double    ' トリミング数カウンター
    '    <VBFixedArray(MAX_RES_USER)> Dim Total_TrimCounter() As Double ' トリミング数カウンター


    '    Public stCounter1 As RESULT_PARAM                        ' 表示用データ定義

    '    'この構造体のインスタンスを初期化するには、"Initialize" を呼び出さなければなりません。
    '    Public Sub Initialize()
    '        ReDim gITNx_cnt(MAX_RES_USER)                       ' IT 算出用ﾜｰｸ数
    '        ReDim gITNg_cnt(MAX_RES_USER)                       ' IT NG数記録
    '        ReDim gFTNx_cnt(MAX_RES_USER)                       ' FT 算出用ﾜｰｸ数
    '        ReDim gFTNg_cnt(MAX_RES_USER)                       ' FT NG数記録
    '        ReDim dblAverage(MAX_RES_USER)                      ' 平均値
    '        ReDim dblDeviationIT(MAX_RES_USER)                  ' 標準偏差(IT)
    '        ReDim dblDeviationFT(MAX_RES_USER)                  ' 標準偏差(FT)
    '        ReDim dblAverageIT(MAX_RES_USER)                    ' IT平均値
    '        ReDim dblAverageFT(MAX_RES_USER)                    ' FT平均値

    '        ReDim TotalIT(MAX_RES_USER)                         ' IT 合計
    '        ReDim TotalFT(MAX_RES_USER)                         ' FT 合計
    '        ReDim TotalSum2IT(MAX_RES_USER)                     ' IT２乗和 
    '        ReDim TotalSum2FT(MAX_RES_USER)                     ' FT２乗和 

    '        ReDim dblMinIT(MAX_RES_USER)                        ' IT最小値ﾌｧｲﾅﾙ
    '        ReDim dblMaxIT(MAX_RES_USER)                        ' IT最大値ﾌｧｲﾅﾙ
    '        ReDim dblMinFT(MAX_RES_USER)                        ' FT最小値ﾌｧｲﾅﾙ
    '        ReDim dblMaxFT(MAX_RES_USER)                        ' FT最大値ﾌｧｲﾅﾙ
    '        ReDim TrimCounter(MAX_RES_USER)                     ' トリミング数カウンター
    '        ReDim Total_TrimCounter(MAX_RES_USER)               ' トリミング数カウンタートータル
    '    End Sub

    'End Structure
    '' 複数抵抗値取得用の集計データ保存用 
    'Public stToTalDataMulti(MAX_RES_USER) As TOTAL_DATA_MULTI

    'V2.2.0.0⑯ 
#Region "フォーム初期化"
    '''=========================================================================
    '''<summary>ﾌｫｰﾑ初期化時処理</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub InitializeForm()
        Dim strMSG As String

        Try
            ' 分布図表示用ラベル配列の初期化
            gDistRegNumLblAry(0) = Me.LblRegN_00             ' 分布グラフ抵抗数配列(0～11)
            gDistRegNumLblAry(1) = Me.LblRegN_01
            gDistRegNumLblAry(2) = Me.LblRegN_02
            gDistRegNumLblAry(3) = Me.LblRegN_03
            gDistRegNumLblAry(4) = Me.LblRegN_04
            gDistRegNumLblAry(5) = Me.LblRegN_05
            gDistRegNumLblAry(6) = Me.LblRegN_06
            gDistRegNumLblAry(7) = Me.LblRegN_07
            gDistRegNumLblAry(8) = Me.LblRegN_08
            gDistRegNumLblAry(9) = Me.LblRegN_09
            gDistRegNumLblAry(10) = Me.LblRegN_10
            gDistRegNumLblAry(11) = Me.LblRegN_11

            gDistGrpPerLblAry(0) = Me.LblGrpPer_00           ' 分布グラフ%配列(0～11)
            gDistGrpPerLblAry(1) = Me.LblGrpPer_01
            gDistGrpPerLblAry(2) = Me.LblGrpPer_02
            gDistGrpPerLblAry(3) = Me.LblGrpPer_03
            gDistGrpPerLblAry(4) = Me.LblGrpPer_04
            gDistGrpPerLblAry(5) = Me.LblGrpPer_05
            gDistGrpPerLblAry(6) = Me.LblGrpPer_06
            gDistGrpPerLblAry(7) = Me.LblGrpPer_07
            gDistGrpPerLblAry(8) = Me.LblGrpPer_08
            gDistGrpPerLblAry(9) = Me.LblGrpPer_09
            gDistGrpPerLblAry(10) = Me.LblGrpPer_10
            gDistGrpPerLblAry(11) = Me.LblGrpPer_11

            gDistShpGrpLblAry(0) = Me.LblShpGrp_00                      ' 分布グラフ配列(0～11)
            gDistShpGrpLblAry(1) = Me.LblShpGrp_01
            gDistShpGrpLblAry(2) = Me.LblShpGrp_02
            gDistShpGrpLblAry(3) = Me.LblShpGrp_03
            gDistShpGrpLblAry(4) = Me.LblShpGrp_04
            gDistShpGrpLblAry(5) = Me.LblShpGrp_05
            gDistShpGrpLblAry(6) = Me.LblShpGrp_06
            gDistShpGrpLblAry(7) = Me.LblShpGrp_07
            gDistShpGrpLblAry(8) = Me.LblShpGrp_08
            gDistShpGrpLblAry(9) = Me.LblShpGrp_09
            gDistShpGrpLblAry(10) = Me.LblShpGrp_10
            gDistShpGrpLblAry(11) = Me.LblShpGrp_11

            'DistRegItLblAry(i) = New System.Windows.Forms.Label     ' 分布グラフ抵抗数(IT)配列
            'DistRegFtLblAry(i) = New System.Windows.Forms.Label     ' 分布グラフ抵抗数(FT)配列

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "frmDistribution.InitializeForm() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "イニシャル/ファイナル分布図の表示状態"
    Public Function DisplayInitialMode() As Boolean
        Return m_bFgDispGrp
    End Function
#End Region

#Region "分布図保存ボタン押下時処理"
    '''=========================================================================
    '''<summary>分布図保存ボタン押下時処理</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub cmdGraphSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdGraphSave.Click

        ' ボタン制御
        cmdGraphSave.Enabled = False
        cmdInitial.Enabled = False
        cmdFinal.Enabled = False

        ' 画面をハードコピーし印刷する
        Call SaveWindowPic(True, False)

        ' 完了メッセージ

        ' ボタン制御
        cmdGraphSave.Enabled = True
        cmdInitial.Enabled = True
        cmdFinal.Enabled = True

    End Sub
#End Region

#Region "分布図保存処理"
    '''=========================================================================
    '''<summary>分布図保存ボタン押下時処理</summary>
    '''<remarks>PrintScreenキー押下時と同等の処理を行う</remarks>
    '''=========================================================================
    Private Sub SaveWindowPic(Optional ByRef ActWind As Boolean = True, Optional ByRef PrintOn As Boolean = False)

        Dim msg As String               'V4.7.0.0③

        Try
            If (String.IsNullOrEmpty(IO.Path.GetFileNameWithoutExtension(gsDataFileName))) Then Exit Sub 'V4.7.0.0③

            Dim fileName As String
            Dim bFileSave As Boolean
            'Dim bitMap As New Bitmap(Me.Width, Me.Height)
            bFileSave = False
            fileName = ""

            ''アクティブなWindowをクリップボードへコピー
            'SendKeys.SendWait("%{PRTSC}")

            '' クリップボードからデータ取得
            'Dim obj As IDataObject = Clipboard.GetDataObject()

            'If obj IsNot Nothing Then
            '    Dim dispImage As Image = DirectCast(obj.GetData(DataFormats.Bitmap), Image)

            '    If dispImage IsNot Nothing Then
            '        If m_bFgDispGrp = True Then
            '            fileName = gSysPrm.stLOG.gsLoggingDir & "IT_MAP" & Now.ToString("yyMMddhhmmss") & ".BMP"
            '        Else
            '            fileName = gSysPrm.stLOG.gsLoggingDir & "FT_MAP" & Now.ToString("yyMMddhhmmss") & ".BMP"
            '        End If

            '        dispImage.Save(fileName)
            '        bFileSave = True
            '    End If
            'End If

            ' ｸﾘｯﾌﾟﾎﾞｰﾄﾞにﾃｷｽﾄ(Bitmap以外？)がｺﾋﾟｰされている状態だと
            ' dispImageがNothingとなって保存されないため変更              'V4.7.0.0③
            Dim ITFT As String
            If (True = m_bFgDispGrp) Then
                ITFT = "_IT_MAP"
            Else
                ITFT = "_FT_MAP"
            End If

            fileName = gSysPrm.stLOG.gsLoggingDir & _
                IO.Path.GetFileNameWithoutExtension(IO.Path.GetFileNameWithoutExtension(gsDataFileName)) & _
                ITFT & Now.ToString("yyMMddHHmmss") & ".BMP"

            Using bmp As New Bitmap(Me.Width, Me.Height)
                Me.DrawToBitmap(bmp, New Rectangle(0, 0, Me.Width, Me.Height))
                bmp.Save(fileName)
                bFileSave = True
            End Using

            '結果の表示
            If (bFileSave = True) Then
                If gSysPrm.stTMN.giMsgTyp = 0 Then
                    'MsgBox("保存完了！" & vbCrLf & " (" & fileName & ")")
                    msg = "保存完了！" & vbCrLf & " (" & fileName & ")"
                Else
                    'MsgBox("Save completion." & vbCrLf & " (" & fileName & ")")
                    msg = "Save completion." & vbCrLf & " (" & fileName & ")"
                End If
            Else
                If gSysPrm.stTMN.giMsgTyp = 0 Then
                    'MsgBox("保存できませんでした。")
                    msg = "保存できませんでした。"
                Else
                    'MsgBox("I was not able to save it.")
                    msg = "I was not able to save it."
                End If
            End If

            'Exit Sub

        Catch ex As Exception
            If gSysPrm.stTMN.giMsgTyp = 0 Then
                'MsgBox("保存できませんでした。")
                msg = "保存できませんでした。" & Environment.NewLine & ex.Message
            Else
                'MsgBox("I was not able to save it.")
                msg = "I was not able to save it." & Environment.NewLine & ex.Message
            End If
        End Try

        ' 後ろに隠れないように対応       'V4.7.0.0③
        MessageBox.Show(msg, cmdGraphSave.Text, MessageBoxButtons.OK, MessageBoxIcon.None, _
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
    End Sub
#End Region

#Region "ファイナルテスト分布図表示ボタン押下時処理"
    '''=========================================================================
    '''<summary>ファイナルテスト分布図表示ボタン押下時処理</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub cmdFinal_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdFinal.Click
        m_bFgDispGrp = False
        Call RedrawGraph()                                              ' 分布図表示処理
    End Sub
#End Region

#Region "イニシャルテスト分布図表示ボタン押下時処理"
    '''=========================================================================
    '''<summary>イニシャルテスト分布図表示ボタン押下時処理</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub cmdInitial_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdInitial.Click
        m_bFgDispGrp = True
        Call RedrawGraph()
    End Sub
#End Region

#Region "フォームロード時処理"
    '''=========================================================================
    '''<summary>フォームロード時処理</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub frmDistribution_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        'Dim utdClientPoint As tagPOINT
        'Dim lngWin32apiResultCode As Integer
        Dim setLocation As System.Drawing.Point

        '初期化実行
        If (m_bInitDistForm = False) Then
            InitializeForm()
            m_bInitDistForm = True
        End If

        'bFgfrmDistribution = True                           ' 生産ｸﾞﾗﾌ表示ﾌﾗｸﾞON

        'Videoの上に表示する。
        setLocation = Form1.VideoLibrary1.Location
        setLocation.X = setLocation.X + DISP_X_OFFSET
        setLocation.Y = setLocation.Y + DISP_Y_OFFSET
        Me.Location = setLocation

        lblRegistTitle.Text = PIC_TRIM_09
        lblGoodTitle.Text = PIC_TRIM_03
        lblNgTitle.Text = PIC_TRIM_04
        lblMinTitle.Text = PIC_TRIM_05
        lblMaxTitle.Text = PIC_TRIM_06
        lblAverage.Text = PIC_TRIM_07
        lblDeviation.Text = PIC_TRIM_08
        cmdInitial.Text = PIC_TRIM_01
        cmdFinal.Text = PIC_TRIM_02

        ' 分布図ﾋﾞｯﾄﾏｯﾌﾟ保存
        cmdGraphSave.Visible = True
        cmdGraphSave.Text = PIC_TRIM_10
        RedrawGraph()

        '常に最前面に表示する。
        Me.TopMost = True
    End Sub
#End Region

#Region "フォーカスを失った時の処理"
    '''=========================================================================
    '''<summary>ロギング開始(標準)</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub frmDistribution_LostFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.LostFocus
        '    Unload Me
    End Sub
#End Region

#Region "フォームアンロード時処理"
    '''=========================================================================
    '''<summary>フォームアンロード時処理</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub frmDistribution_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        'bFgfrmDistribution = False                      ' 生産ｸﾞﾗﾌ表示ﾌﾗｸﾞOFF
        Form1.chkDistributeOnOff.Checked = False

        If (gSysPrm.stTMN.giMsgTyp = 0) Then
            Form1.chkDistributeOnOff.Text = "生産グラフ　表示"
        Else
            Form1.chkDistributeOnOff.Text = "Distribute ON"
        End If
    End Sub
#End Region

#Region "分布図表示処理"
    '''=========================================================================
    '''<summary>分布図表示処理</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub RedrawGraph()

        Dim iCnt As Short                                   ' ｶｳﾝﾀ
        Dim lMax As Integer
        Dim lScale As Integer
        Dim lScaleMax As Integer
        Dim dblGraphDiv As Double
        Dim dblGraphTop As Double
        Dim dtemp As Double         ' ###203 
        Dim dblTest_LowLimit, dblTest_HighLimit As Double
        Dim dblTemp As Double

        lMax = 0
        If (m_bFgDispGrp) Then

            lblGraphAccumulationTitle.Text = MSG_TRIM_04                ' "イニシャルテスト　分布図"
            If Double.MaxValue = dblMinIT(giDistributionResNo) Then
                lblMinValue.Text = "0.000"
            Else
                lblMinValue.Text = dblMinIT(giDistributionResNo).ToString("0.000")               ' 最小値
            End If
            If Double.MinValue = dblMaxIT(giDistributionResNo) Then
                lblMaxValue.Text = "0.000"
            Else
                lblMaxValue.Text = dblMaxIT(giDistributionResNo).ToString("0.000")               ' 最大値
            End If

            For iCnt = 0 To (MAX_SCALE_RNUM - 1)
                glRegistNum(iCnt) = glRegistNumIT(giDistributionResNo, iCnt)                 ' 分布グラフ抵抗数
                If lMax < glRegistNum(iCnt) Then
                    lMax = glRegistNum(iCnt)
                End If

                gDistRegNumLblAry(iCnt).Text = CStr(glRegistNum(iCnt)) ' 分布グラフ抵抗数
            Next

            'OK/NG数の表示
            lblGoodChip.Text = CStr(gITNx_cnt(giDistributionResNo))                      ' OK数
            lblNgChip.Text = CStr(gITNg_cnt(giDistributionResNo))                        ' NG数
            lblOKRate.Text = dblOKRateIT(giDistributionResNo).ToString("0.000")              ' 良品率
            lblNGRate.Text = dblNGRateIT(giDistributionResNo).ToString("0.000")              ' 不良品率
        Else
            lblGraphAccumulationTitle.Text = MSG_TRIM_05                ' "ファイナルテスト　分布図"
            If Double.MaxValue = dblMinFT(giDistributionResNo) Then
                lblMinValue.Text = "0.000"
            Else
                lblMinValue.Text = dblMinFT(giDistributionResNo).ToString("0.000")               ' 最小値
            End If
            If Double.MinValue = dblMaxFT(giDistributionResNo) Then
                lblMaxValue.Text = "0.000"
            Else
                lblMaxValue.Text = dblMaxFT(giDistributionResNo).ToString("0.000")               ' 最大値
            End If

            For iCnt = 0 To (MAX_SCALE_RNUM - 1)

                glRegistNum(iCnt) = glRegistNumFT(giDistributionResNo, iCnt)

                If lMax < glRegistNum(iCnt) Then
                    lMax = glRegistNum(iCnt)
                End If
                gDistRegNumLblAry(iCnt).Text = CStr(glRegistNum(iCnt)) ' 分布グラフ抵抗数
            Next
            'OK/NG数の表示
            lblGoodChip.Text = CStr(gFTNx_cnt(giDistributionResNo))                      ' OK数
            lblNgChip.Text = CStr(gFTNg_cnt(giDistributionResNo))                        ' NG数
            lblOKRate.Text = dblOKRateFT(giDistributionResNo).ToString("0.000")              ' 良品率
            lblNGRate.Text = dblNGRateFT(giDistributionResNo).ToString("0.000")              ' 不良品率
        End If

        'lblGoodChip.Text = CStr(lOkChip)                               ' OK数
        'lblNgChip.Text = CStr(lNgChip)                                 ' NG数


        '■■■■■■
        ' 誤差ﾃﾞｰﾀがある(IT)
        '' '' ''Call Form1.GetMoveMode(digL, digH, digSW)
        If gITNx_cnt(giDistributionResNo) >= 0 Then
            'If (gDigL = 0) Then                                        ' x0モード ?
            '' '' ''If (digL = 0) Then                                  ' x0モード ?
            '###154 計算は結果取得時にその都度実行する
            '' 平均値取得
            'dblAverageIT = Form1.Utility1.GetAverage(gITNx, gITNx_cnt + 1)
            '' 標準偏差の取得
            'dblDeviationIT = Form1.Utility1.GetDeviation(gITNx, gITNx_cnt + 1, dblAverageIT)
            'TotalDeviationDebug = TotalDeviationDebug '###154
            'TotalAverageDebug = TotalAverageDebug '###154
            '' '' ''End If
        End If

        ' 誤差ﾃﾞｰﾀがある(FT)
        If gFTNx_cnt(giDistributionResNo) >= 0 Then
            '###154            ' 平均値取得
            '###154            dblAverageFT = Form1.Utility1.GetAverage(gFTNx, gFTNx_cnt + 1)
            '###154     ' 標準偏差の取得
            '###154         dblDeviationFT = Form1.Utility1.GetDeviation(gFTNx, gFTNx_cnt + 1, dblAverageFT)
            'dblAverageFT = TotalAverageDebug '###154
            'dblDeviationFT = TotalDeviationDebug '###154
        End If
        '■■■■■■■

        If (m_bFgDispGrp) Then
            lblDeviationValue.Text = dblDeviationIT(giDistributionResNo).ToString("0.000000") ' 標準偏差(IT)
        Else
            lblDeviationValue.Text = dblDeviationFT(giDistributionResNo).ToString("0.000000") ' 標準偏差(FT)
        End If

        If (m_bFgDispGrp) Then
            dblAverage(giDistributionResNo) = dblAverageIT(giDistributionResNo)
        Else
            dblAverage(giDistributionResNo) = dblAverageFT(giDistributionResNo)
        End If
        lblAverageValue.Text = dblAverage(giDistributionResNo).ToString("0.000")     ' 平均値

        lScaleMax = 0                                           ' オートスケーリング
        lScale = 100
        Do
            If (lScale > lMax) Then                             ' lScale < 抵抗数 ?
                lScaleMax = lScale
            ElseIf ((lScale * 2) > lMax) Then
                lScaleMax = (lScale * 2)
            ElseIf ((lScale * 5) > lMax) Then
                lScaleMax = (lScale * 5)
            End If
            lScale = lScale * 10
        Loop While (0 = lScaleMax) And (MAX_SCALE_NUM > lScale)

        If (0 = lScaleMax) Then
            lScaleMax = MAX_SCALE_NUM + 1
        End If


        If (m_bFgDispGrp) Then
            If giDistributionResNo = 0 Then
                For i As Integer = 1 To stPLT.RCount
                    If Not UserModule.IsMarking(i) Then
                        dblTemp = stREG(i).dblITL
                        If stREG(i).intMode = JUDGE_MODE_ABSOLUTE Then   ' 絶対値の時比率に変換
                            dblTemp = dblTemp / stREG(i).dblNOM * 100.0
                        End If
                        If dblTest_LowLimit > dblTemp Then
                            dblTest_LowLimit = dblTemp
                        End If
                        dblTemp = stREG(i).dblITH
                        If stREG(i).intMode = JUDGE_MODE_ABSOLUTE Then   ' 絶対値の時比率に変換
                            dblTemp = dblTemp / stREG(i).dblNOM * 100.0
                        End If
                        If dblTest_HighLimit < dblTemp Then
                            dblTest_HighLimit = dblTemp
                        End If
                    End If
                Next
            Else
                dblTest_LowLimit = stREG(giDistributionResNo).dblITL
                dblTest_HighLimit = stREG(giDistributionResNo).dblITH
            End If
        Else
            If giDistributionResNo = 0 Then
                For i As Integer = 1 To stPLT.RCount
                    If Not UserModule.IsMarking(i) Then
                        dblTemp = stREG(i).dblFTL
                        If stREG(i).intMode = JUDGE_MODE_ABSOLUTE Then   ' 絶対値の時比率に変換
                            dblTemp = dblTemp / stREG(i).dblNOM * 100.0
                        End If
                        If dblTest_LowLimit > dblTemp Then
                            dblTest_LowLimit = dblTemp
                        End If
                        dblTemp = stREG(i).dblFTH
                        If stREG(i).intMode = JUDGE_MODE_ABSOLUTE Then   ' 絶対値の時比率に変換
                            dblTemp = dblTemp / stREG(i).dblNOM * 100.0
                        End If
                        If dblTest_HighLimit < dblTemp Then
                            dblTest_HighLimit = dblTemp
                        End If
                    End If
                Next
            Else
                dblTest_LowLimit = stREG(giDistributionResNo).dblFTL
                dblTest_HighLimit = stREG(giDistributionResNo).dblFTH
            End If
        End If

        If giDistributionResNo > 0 And stREG(giDistributionResNo).intMode = JUDGE_MODE_ABSOLUTE Then   ' 絶対値の時比率に変換
            dblTest_LowLimit = dblTest_LowLimit / stREG(giDistributionResNo).dblNOM * 100.0
            dblTest_HighLimit = dblTest_HighLimit / stREG(giDistributionResNo).dblNOM * 100.0
        End If

        If ((0 >= dblTest_LowLimit) And (0 <= dblTest_HighLimit)) Then
            dblGraphDiv = (dblTest_HighLimit * 1.5 - dblTest_LowLimit * 1.5) / 10
            dblGraphTop = dblTest_HighLimit * 1.5
        ElseIf ((0 >= dblTest_LowLimit) And (0 > dblTest_HighLimit)) Then
            dblGraphDiv = (dblTest_HighLimit / 1.5 - dblTest_LowLimit * 1.5) / 10
            dblGraphTop = dblTest_HighLimit * 1.5
        ElseIf ((0 < dblTest_LowLimit) And (0 <= dblTest_HighLimit)) Then
            dblGraphDiv = (dblTest_HighLimit * 1.5 - dblTest_LowLimit / 1.5) / 10
            dblGraphTop = dblTest_HighLimit * 1.5
        Else
            dblGraphDiv = 0.3
            dblGraphTop = 1.5
        End If

        gDistGrpPerLblAry(0).Text = "～" & dblGraphTop.ToString("0.00")
        For iCnt = 1 To 11
            'gDistGrpPerLblAry(iCnt).Text = (dblGraphTop - (dblGraphDiv * (iCnt - 1)).ToString("0.00")) & "～"
            ' ###203 
            dtemp = (dblGraphTop - (dblGraphDiv * (iCnt - 1)))
            If ((-0.001 < dtemp) And (dtemp < 0.001)) Then
                gDistGrpPerLblAry(iCnt).Text = "0～"
            Else
                gDistGrpPerLblAry(iCnt).Text = (dtemp.ToString("0.00")) & "～"
            End If
            ' ###203
        Next

        picGraphAccumulationDrawLine(lScaleMax)
        Call picGraphAccumulationPrintRegistNum()           ' 分布グラフに抵抗数を設定する

    End Sub
#End Region

#Region "分布図表示サブ"
    '''=========================================================================
    ''' <summary>
    ''' 分布図表示サブ
    ''' </summary>
    ''' <param name="lScaleMax">(INP)スケール</param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub picGraphAccumulationDrawLine(ByRef lScaleMax As Integer)

        Dim i As Short
        Dim x As Short

        For i = 0 To (MAX_SCALE_RNUM - 1)
            '            x = CShort((glRegistNum(i) * 473) \ lScaleMax)   ' 分布グラフ抵抗数
            x = CShort((glRegistNum(i) * 250) \ lScaleMax)   ' 分布グラフ抵抗数
            'If (473 < x) Then
            If (250 < x) Then
                '                x = 473
                x = 250
            End If
            gDistShpGrpLblAry(i).Width = x
        Next
        lblRegistUnit.Text = CStr(lScaleMax \ 2)            ' 抵抗数の半分の数 

    End Sub
#End Region

#Region "分布グラフに抵抗数を設定する"
    '''=========================================================================
    '''<summary>分布グラフに抵抗数を設定する</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub picGraphAccumulationPrintRegistNum()

        Dim i As Short

        For i = 0 To (MAX_SCALE_RNUM - 1)
            gDistRegNumLblAry(i).Text = CStr(glRegistNum(i))  ' 分布グラフ抵抗数
        Next

    End Sub
#End Region

#Region "分布図データクリア"
    Public Sub ClearCounter()
        Try

            For i As Integer = 0 To MAXRNO

                dblAverage(i) = 0.0                     ' 平均値
                dblDeviationIT(i) = 0.0                 ' 標準偏差(IT)
                dblDeviationFT(i) = 0.0                 ' 標準偏差(FT)
                dblAverageIT(i) = 0.0                   ' IT平均値
                dblAverageFT(i) = 0.0                   ' FT平均値

                For j As Integer = 0 To MAX_SCALE_RNUM
                    glRegistNum(j) = 0               ' 分布グラフ抵抗数
                    glRegistNumIT(i, j) = 0             ' 分布グラフ抵抗数 ｲﾆｼｬﾙﾃｽﾄ
                    glRegistNumFT(i, j) = 0             ' 分布グラフ抵抗数 ﾌｧｲﾅﾙﾃｽﾄ
                Next

                dblMinIT(i) = Double.MaxValue           ' 最小値ｲﾆｼｬﾙ
                dblMaxIT(i) = Double.MinValue           ' 最大値ｲﾆｼｬﾙ
                dblMinFT(i) = Double.MaxValue           ' 最小値ﾌｧｲﾅﾙ
                dblMaxFT(i) = Double.MinValue           ' 最大値ﾌｧｲﾅﾙ

                dblOKRateIT(i) = 0.0                    ' 良品率ｲﾆｼｬﾙ
                dblNGRateIT(i) = 0.0                    ' 不良品率ｲﾆｼｬﾙ
                dblOKRateFT(i) = 0.0                    ' 良品率ﾌｧｲﾅﾙ
                dblNGRateFT(i) = 0.0                    ' 不良品率ﾌｧｲﾅﾙ

                gITNx_cnt(i) = 0                        'IT 算出用ﾜｰｸ数
                gITNg_cnt(i) = 0                        'IT NG数記録
                gFTNx_cnt(i) = 0                        'FT 算出用ﾜｰｸ数
                gFTNg_cnt(i) = 0                        'FT NG数記録

                TotalFT(i) = 0.0                        ' FT 合計      
                TotalIT(i) = 0.0                        ' IT 合計
                TotalSum2FT(i) = 0.0                    ' FT２乗和      
                TotalSum2IT(i) = 0.0                    ' IT２乗和     
            Next


            'V2.2.0.0⑯↓
            For MultiCnt As Integer = 0 To MAX_RES_USER
                With stToTalDataMulti(MultiCnt)
                    .Initialize()

                    For rn As Integer = 0 To MAX_RES_USER
                        .gITNx_cnt(rn) = 0                      ' IT 算出用ﾜｰｸ数
                        .gITNg_cnt(rn) = 0                      ' IT NG数記録
                        .gFTNx_cnt(rn) = 0                      ' FT 算出用ﾜｰｸ数
                        .gFTNg_cnt(rn) = 0                      ' FT NG数記録
                        .dblAverage(rn) = 0                     ' 平均値
                        .dblDeviationIT(rn) = 0                 ' 標準偏差(IT)
                        .dblDeviationFT(rn) = 0                 ' 標準偏差(FT)
                        .dblAverageIT(rn) = 0                   ' IT平均値
                        .dblAverageFT(rn) = 0                   ' FT平均値
                        .TotalIT(rn) = 0.0                      ' IT 合計      
                        .TotalFT(rn) = 0.0                      ' FT 合計      
                        .TotalSum2IT(rn) = 0                    ' IT２乗和
                        .TotalSum2FT(rn) = 0                    ' FT２乗和
                        .dblMinIT(rn) = Double.MaxValue         ' IT最小値ﾌｧｲﾅﾙ
                        .dblMaxIT(rn) = Double.MinValue         ' IT最大値ﾌｧｲﾅﾙ
                        .dblMinFT(rn) = Double.MaxValue         ' FT最小値ﾌｧｲﾅﾙ
                        .dblMaxFT(rn) = Double.MinValue         ' FT最大値ﾌｧｲﾅﾙ

                    Next rn
                End With
            Next MultiCnt
            'V2.2.0.0⑯↑

        Catch ex As Exception
            Call Z_PRINT("frmDistribution.ClearCounter() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub
#End Region

#Region "分布図データ設定"
    Public Sub SetGraphData(ByVal dTop As Double, ByVal dDiv As Double, ByVal dGap As Double, ByVal rn As Integer, ByRef iRegistNum(,) As Integer)
        Try
            If ((dTop - (dDiv * 0)) < dGap) Then
                ' ｸﾞﾗﾌ最上段値-(ｸﾞﾗﾌ範囲刻み位置*0)　<　差の場合
                iRegistNum(rn, 0) = iRegistNum(rn, 0) + 1
            ElseIf ((dTop - (dDiv * 1)) < dGap) Then
                ' ｸﾞﾗﾌ最上段値-(ｸﾞﾗﾌ範囲刻み位置*1)　<　差の場合
                iRegistNum(rn, 1) = iRegistNum(rn, 1) + 1
            ElseIf ((dTop - (dDiv * 2)) < dGap) Then
                ' ｸﾞﾗﾌ最上段値-(ｸﾞﾗﾌ範囲刻み位置*2)　<　差の場合
                iRegistNum(rn, 2) = iRegistNum(rn, 2) + 1
            ElseIf ((dTop - (dDiv * 3)) < dGap) Then
                ' ｸﾞﾗﾌ最上段値-(ｸﾞﾗﾌ範囲刻み位置*3)　<　差の場合
                iRegistNum(rn, 3) = iRegistNum(rn, 3) + 1
            ElseIf ((dTop - (dDiv * 4)) < dGap) Then
                ' ｸﾞﾗﾌ最上段値-(ｸﾞﾗﾌ範囲刻み位置*4)　<　差の場合
                iRegistNum(rn, 4) = iRegistNum(rn, 4) + 1
            ElseIf ((dTop - (dDiv * 5)) < dGap) Then
                ' ｸﾞﾗﾌ最上段値-(ｸﾞﾗﾌ範囲刻み位置*5)　<　差の場合
                iRegistNum(rn, 5) = iRegistNum(rn, 5) + 1
            ElseIf ((dTop - (dDiv * 6)) < dGap) Then
                ' ｸﾞﾗﾌ最上段値-(ｸﾞﾗﾌ範囲刻み位置*6)　<　差の場合
                iRegistNum(rn, 6) = iRegistNum(rn, 6) + 1
            ElseIf ((dTop - (dDiv * 7)) < dGap) Then
                ' ｸﾞﾗﾌ最上段値-(ｸﾞﾗﾌ範囲刻み位置*7)　<　差の場合
                iRegistNum(rn, 7) = iRegistNum(rn, 7) + 1
            ElseIf ((dTop - (dDiv * 8)) < dGap) Then
                ' ｸﾞﾗﾌ最上段値-(ｸﾞﾗﾌ範囲刻み位置*8)　<　差の場合
                iRegistNum(rn, 8) = iRegistNum(rn, 8) + 1
            ElseIf ((dTop - (dDiv * 9)) < dGap) Then
                ' ｸﾞﾗﾌ最上段値-(ｸﾞﾗﾌ範囲刻み位置*9)　<　差の場合
                iRegistNum(rn, 9) = iRegistNum(rn, 9) + 1
            ElseIf ((dTop - (dDiv * 10)) < dGap) Then
                ' ｸﾞﾗﾌ最上段値-(ｸﾞﾗﾌ範囲刻み位置*10)　<　差の場合
                iRegistNum(rn, 10) = iRegistNum(rn, 10) + 1
            Else
                ' 上記条件以外の場合
                iRegistNum(rn, 11) = iRegistNum(rn, 11) + 1
            End If
        Catch ex As Exception
            Call Z_PRINT("frmDistribution.SetGraphData() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub

    ''' <summary>
    ''' 標準偏差の計算と各種データの保存
    ''' </summary>
    ''' <param name="rn"></param>
    ''' <param name="dblGap"></param>
    ''' <param name="Sum"></param>
    ''' <param name="Total"></param>
    ''' <param name="TotalSum2"></param>
    ''' <param name="dblDeviation"></param>
    ''' <param name="dblMin"></param>
    ''' <param name="dblMax"></param>
    ''' <remarks></remarks>
    Private Sub SetDeviation(ByRef rn As Integer, ByVal dblGap As Double, ByRef Sum() As Integer, ByRef Total() As Double, ByRef TotalSum2() As Double, ByRef dblDeviation() As Double, ByRef dblMin() As Double, ByRef dblMax() As Double, ByRef Average() As Double)
        Try
            Sum(rn) = Sum(rn) + 1                                                           ' データ数カウンター１カウントアップ
            Total(rn) = Total(rn) + dblGap                                                  ' データの合計（和）
            Average(rn) = Total(rn) / Sum(rn)                                                   ' 平均値

            TotalSum2(rn) = TotalSum2(rn) + (dblGap * dblGap)                               ' ２乗和
            dblDeviation(rn) = Math.Sqrt((TotalSum2(rn) / Sum(rn)) - (Average(rn) * Average(rn)))   ' 標準偏差

            'V2.2.0.031↓
            ' 数値になっていない場合は０とする
            If Double.IsNaN(dblDeviation(rn)) Then
                dblDeviation(rn) = 0.0
            End If
            'V2.2.0.031↑

            '(標準偏差算出式修正)
            If (dblMin(rn) > dblGap) Then                                                   ' 最小
                dblMin(rn) = dblGap
            End If
            If (dblMax(rn) < dblGap) Then                                                   ' 最大
                dblMax(rn) = dblGap
            End If

        Catch ex As Exception
            Call Z_PRINT("frmDistribution.SetDeviation() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Sub

    '''=========================================================================
    ''' <summary>
    ''' ファイナルテストテスト分布図
    ''' </summary>
    ''' <param name="JudgeMode">IT or FT</param>
    ''' <param name="rn">抵抗番号</param>
    ''' <param name="MeasureValue">測定値</param>
    ''' <param name="Judge">判定</param>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub TrimLoggingGraph(ByVal JudgeMode As Integer, ByVal rn As Integer, ByVal MeasureValue As Double, ByVal Judge As Integer)


        Try

            Dim dblGraphDiv As Double                                       ' グラフ範囲刻み値
            Dim dblGraphTop As Double                                       ' グラフ最上段値
            Dim dblGraphDivAll As Double                                       ' グラフ範囲刻み値
            Dim dblGraphTopAll As Double                                       ' グラフ最上段値
            Dim dblGap As Double
            Dim dblTemp As Double

            Dim dblTest_LowLimit, dblTest_HighLimit As Double
            Dim dblTest_LowLimitAll As Double = Double.MaxValue, dblTest_HighLimitAll As Double = Double.MinValue

            If JudgeMode = INITIAL_TEST Then
                dblTest_LowLimit = stREG(rn).dblITL
                dblTest_HighLimit = stREG(rn).dblITH
                For i As Integer = 1 To stPLT.RCount
                    If Not UserModule.IsMarking(rn) Then
                        dblTemp = stREG(rn).dblITL
                        If stREG(rn).intMode = JUDGE_MODE_ABSOLUTE Then   ' 絶対値の時比率に変換
                            dblTemp = dblTemp / stREG(rn).dblNOM * 100.0
                        End If
                        If dblTest_LowLimitAll > dblTemp Then
                            dblTest_LowLimitAll = dblTemp
                        End If
                        dblTemp = stREG(rn).dblITH
                        If stREG(rn).intMode = JUDGE_MODE_ABSOLUTE Then   ' 絶対値の時比率に変換
                            dblTemp = dblTemp / stREG(rn).dblNOM * 100.0
                        End If
                        If dblTest_HighLimitAll < dblTemp Then
                            dblTest_HighLimitAll = dblTemp
                        End If
                    End If
                Next
            ElseIf JudgeMode = FINAL_TEST Then
                dblTest_LowLimit = stREG(rn).dblFTL
                dblTest_HighLimit = stREG(rn).dblFTH
                For i As Integer = 1 To stPLT.RCount
                    If Not UserModule.IsMarking(rn) Then
                        dblTemp = stREG(rn).dblFTL
                        If stREG(rn).intMode = JUDGE_MODE_ABSOLUTE Then   ' 絶対値の時比率に変換
                            dblTemp = dblTemp / stREG(rn).dblNOM * 100.0
                        End If
                        If dblTest_LowLimitAll > dblTemp Then
                            dblTest_LowLimitAll = dblTemp
                        End If
                        dblTemp = stREG(rn).dblFTH
                        If stREG(rn).intMode = JUDGE_MODE_ABSOLUTE Then   ' 絶対値の時比率に変換
                            dblTemp = dblTemp / stREG(rn).dblNOM * 100.0
                        End If
                        If dblTest_HighLimitAll < dblTemp Then
                            dblTest_HighLimitAll = dblTemp
                        End If
                    End If
                Next
            End If

            If stREG(rn).intMode = JUDGE_MODE_ABSOLUTE Then   ' 絶対値の時比率に変換
                dblTest_LowLimit = dblTest_LowLimit / stREG(rn).dblNOM * 100.0
                dblTest_HighLimit = dblTest_HighLimit / stREG(rn).dblNOM * 100.0
            End If

            ' 現在の抵抗の計算
            ' ﾌｧｲﾅﾙﾃｽﾄ(LOWﾘﾐｯﾄ)とﾌｧｲﾅﾙﾃｽﾄ(HIGHﾘﾐｯﾄ)の値をﾁｪｯｸする。
            If ((0 >= dblTest_LowLimit) And (0 <= dblTest_HighLimit)) Then
                ' ﾌｧｲﾅﾙﾃｽﾄ(LOWﾘﾐｯﾄ)が0以下でﾌｧｲﾅﾙﾃｽﾄ(HIGHﾘﾐｯﾄ)が0以上の場合
                dblGraphDiv = (dblTest_HighLimit * 1.5 - dblTest_LowLimit * 1.5) / 10
                dblGraphTop = dblTest_HighLimit * 1.5

            ElseIf ((0 >= dblTest_LowLimit) And (0 > dblTest_HighLimit)) Then
                ' ﾌｧｲﾅﾙﾃｽﾄ(LOWﾘﾐｯﾄ)が0以下でﾌｧｲﾅﾙﾃｽﾄ(HIGHﾘﾐｯﾄ)がより小さい場合
                dblGraphDiv = (dblTest_HighLimit / 1.5 - dblTest_LowLimit * 1.5) / 10
                dblGraphTop = dblTest_HighLimit * 1.5

            ElseIf ((0 < dblTest_LowLimit) And (0 <= dblTest_HighLimit)) Then
                ' ﾌｧｲﾅﾙﾃｽﾄ(LOWﾘﾐｯﾄ)が0より大きくてﾌｧｲﾅﾙﾃｽﾄ(HIGHﾘﾐｯﾄ)が0以上の場合
                dblGraphDiv = (dblTest_HighLimit * 1.5 - dblTest_LowLimit / 1.5) / 10
                dblGraphTop = dblTest_HighLimit * 1.5
            Else
                ' 上記条件以外の場合
                dblGraphDiv = 0.3
                dblGraphTop = 1.5
            End If

            ' 全抵抗の計算
            ' ﾌｧｲﾅﾙﾃｽﾄ(LOWﾘﾐｯﾄ)とﾌｧｲﾅﾙﾃｽﾄ(HIGHﾘﾐｯﾄ)の値をﾁｪｯｸする。
            If ((0 >= dblTest_LowLimitAll) And (0 <= dblTest_HighLimitAll)) Then
                ' ﾌｧｲﾅﾙﾃｽﾄ(LOWﾘﾐｯﾄ)が0以下でﾌｧｲﾅﾙﾃｽﾄ(HIGHﾘﾐｯﾄ)が0以上の場合
                dblGraphDivAll = (dblTest_HighLimitAll * 1.5 - dblTest_LowLimitAll * 1.5) / 10
                dblGraphTopAll = dblTest_HighLimitAll * 1.5

            ElseIf ((0 >= dblTest_LowLimitAll) And (0 > dblTest_HighLimitAll)) Then
                ' ﾌｧｲﾅﾙﾃｽﾄ(LOWﾘﾐｯﾄ)が0以下でﾌｧｲﾅﾙﾃｽﾄ(HIGHﾘﾐｯﾄ)がより小さい場合
                dblGraphDivAll = (dblTest_HighLimitAll / 1.5 - dblTest_LowLimitAll * 1.5) / 10
                dblGraphTopAll = dblTest_HighLimitAll * 1.5

            ElseIf ((0 < dblTest_LowLimitAll) And (0 <= dblTest_HighLimitAll)) Then
                ' ﾌｧｲﾅﾙﾃｽﾄ(LOWﾘﾐｯﾄ)が0より大きくてﾌｧｲﾅﾙﾃｽﾄ(HIGHﾘﾐｯﾄ)が0以上の場合
                dblGraphDivAll = (dblTest_HighLimitAll * 1.5 - dblTest_LowLimitAll / 1.5) / 10
                dblGraphTopAll = dblTest_HighLimitAll * 1.5
            Else
                ' 上記条件以外の場合
                dblGraphDivAll = 0.3
                dblGraphTopAll = 1.5
            End If

            ' 差を算出する。　ﾌｧｲﾅﾙﾃｽﾄ結果/ﾄﾘﾐﾝｸﾞ目標値*100　-　100
            dblGap = (MeasureValue / stREG(rn).dblNOM) * 100.0# - 100.0#

            If JudgeMode = INITIAL_TEST Then
                SetGraphData(dblGraphTop, dblGraphDiv, dblGap, rn, glRegistNumIT)
                SetGraphData(dblGraphTopAll, dblGraphDivAll, dblGap, 0, glRegistNumIT)
            ElseIf JudgeMode = FINAL_TEST Then
                SetGraphData(dblGraphTop, dblGraphDiv, dblGap, rn, glRegistNumFT)
                SetGraphData(dblGraphTopAll, dblGraphDivAll, dblGap, 0, glRegistNumFT)
            End If


            If Judge = eJudge.JG_OK Then
                If JudgeMode = INITIAL_TEST Then
                    Call SetDeviation(rn, dblGap, gITNx_cnt, TotalIT, TotalSum2IT, dblDeviationIT, dblMinIT, dblMaxIT, dblAverageIT)  ' 現抵抗
                    Call SetDeviation(0, dblGap, gITNx_cnt, TotalIT, TotalSum2IT, dblDeviationIT, dblMinIT, dblMaxIT, dblAverageIT)   ' 全抵抗
                ElseIf JudgeMode = FINAL_TEST Then
                    Call SetDeviation(rn, dblGap, gFTNx_cnt, TotalFT, TotalSum2FT, dblDeviationFT, dblMinFT, dblMaxFT, dblAverageFT)  ' 現抵抗
                    Call SetDeviation(0, dblGap, gFTNx_cnt, TotalFT, TotalSum2FT, dblDeviationFT, dblMinFT, dblMaxFT, dblAverageFT)   ' 全抵抗
                End If
            Else
                'NGカウント数を記録
                If JudgeMode = INITIAL_TEST Then
                    gITNg_cnt(rn) = gITNg_cnt(rn) + 1
                    gITNg_cnt(0) = gITNg_cnt(0) + 1
                ElseIf JudgeMode = FINAL_TEST Then
                    gFTNg_cnt(rn) = gFTNg_cnt(rn) + 1
                    gFTNg_cnt(0) = gFTNg_cnt(0) + 1
                End If
            End If

            ' 良品率、不良品率
            If JudgeMode = INITIAL_TEST Then
                dblOKRateIT(rn) = gITNx_cnt(rn) / (gITNx_cnt(rn) + gITNg_cnt(rn)) * 100.0
                dblNGRateIT(rn) = gITNg_cnt(rn) / (gITNx_cnt(rn) + gITNg_cnt(rn)) * 100.0
                dblOKRateIT(0) = gITNx_cnt(0) / (gITNx_cnt(0) + gITNg_cnt(0)) * 100.0
                dblNGRateIT(0) = gITNg_cnt(0) / (gITNx_cnt(0) + gITNg_cnt(0)) * 100.0
            ElseIf JudgeMode = FINAL_TEST Then
                dblOKRateFT(rn) = gFTNx_cnt(rn) / (gFTNx_cnt(rn) + gFTNg_cnt(rn)) * 100.0
                dblNGRateFT(rn) = gFTNg_cnt(rn) / (gFTNx_cnt(rn) + gFTNg_cnt(rn)) * 100.0
                dblOKRateFT(0) = gFTNx_cnt(0) / (gFTNx_cnt(0) + gFTNg_cnt(0)) * 100.0
                dblNGRateFT(0) = gFTNg_cnt(0) / (gFTNx_cnt(0) + gFTNg_cnt(0)) * 100.0
            End If

        Catch ex As Exception
            Call Z_PRINT("frmDistribution.TrimLoggingGraph() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub
#End Region

#Region "表示抵抗番号の設定"
    Public Sub SetDistributionResNo(ByRef No As Integer)
        Try
            giDistributionResNo = No
        Catch ex As Exception
            Call Z_PRINT("SetDistributionResNo.SetDeviation() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub
#End Region

#Region "統計データのみ処理"
    ''' <summary>
    ''' 統計データの保存
    ''' </summary>
    ''' <param name="JudgeMode">IT or FT</param>
    ''' <param name="rn">抵抗番号</param>
    ''' <param name="MeasureValue">測定値</param>
    ''' <param name="Judge">判定</param>
    ''' <remarks></remarks>
    Public Sub StatisticalDataSave(ByVal JudgeMode As Integer, ByVal rn As Integer, ByVal MeasureValue As Double, ByVal Judge As Integer)

        Try

            Dim dblGap As Double

            ' 差を算出する。　ﾌｧｲﾅﾙﾃｽﾄ結果/ﾄﾘﾐﾝｸﾞ目標値*100　-　100
            '偏差でなく真値を登録する。            dblGap = (MeasureValue / stREG(rn).dblNOM) * 100.0# - 100.0#
            dblGap = MeasureValue

            If Judge = eJudge.JG_OK Then
                If JudgeMode = INITIAL_TEST Then
                    Call SetDeviation(rn, dblGap, gITNx_cnt, TotalIT, TotalSum2IT, dblDeviationIT, dblMinIT, dblMaxIT, dblAverageIT)  ' 現抵抗
                    Call SetDeviation(0, dblGap, gITNx_cnt, TotalIT, TotalSum2IT, dblDeviationIT, dblMinIT, dblMaxIT, dblAverageIT)   ' 全抵抗
                ElseIf JudgeMode = FINAL_TEST Then
                    Call SetDeviation(rn, dblGap, gFTNx_cnt, TotalFT, TotalSum2FT, dblDeviationFT, dblMinFT, dblMaxFT, dblAverageFT)  ' 現抵抗
                    Call SetDeviation(0, dblGap, gFTNx_cnt, TotalFT, TotalSum2FT, dblDeviationFT, dblMinFT, dblMaxFT, dblAverageFT)   ' 全抵抗
                    'V2.2.0.0⑯↓
                    If stMultiBlock.gMultiBlock <> 0 Then
                        ' 複数抵抗値取得の場合の集計データ保存
                        With stToTalDataMulti(stExecBlkData.DataNo)
                            Call SetDeviation(rn, dblGap, .gFTNx_cnt, .TotalFT, .TotalSum2FT, .dblDeviationFT, .dblMinFT, .dblMaxFT, .dblAverageFT)  ' 現抵抗
                            Call SetDeviation(0, dblGap, .gFTNx_cnt, .TotalFT, .TotalSum2FT, .dblDeviationFT, .dblMinFT, .dblMaxFT, .dblAverageFT)   ' 全抵抗
                        End With
                    End If
                    'V2.2.0.0⑯↑
                End If
                Else
                'NGカウント数を記録
                If JudgeMode = INITIAL_TEST Then
                    gITNg_cnt(rn) = gITNg_cnt(rn) + 1
                    gITNg_cnt(0) = gITNg_cnt(0) + 1
                ElseIf JudgeMode = FINAL_TEST Then
                    gFTNg_cnt(rn) = gFTNg_cnt(rn) + 1
                    gFTNg_cnt(0) = gFTNg_cnt(0) + 1
                End If
            End If

            ' 良品率、不良品率
            If JudgeMode = INITIAL_TEST Then
                dblOKRateIT(rn) = gITNx_cnt(rn) / (gITNx_cnt(rn) + gITNg_cnt(rn)) * 100.0
                dblNGRateIT(rn) = gITNg_cnt(rn) / (gITNx_cnt(rn) + gITNg_cnt(rn)) * 100.0
                dblOKRateIT(0) = gITNx_cnt(0) / (gITNx_cnt(0) + gITNg_cnt(0)) * 100.0
                dblNGRateIT(0) = gITNg_cnt(0) / (gITNx_cnt(0) + gITNg_cnt(0)) * 100.0
            ElseIf JudgeMode = FINAL_TEST Then
                dblOKRateFT(rn) = gFTNx_cnt(rn) / (gFTNx_cnt(rn) + gFTNg_cnt(rn)) * 100.0
                dblNGRateFT(rn) = gFTNg_cnt(rn) / (gFTNx_cnt(rn) + gFTNg_cnt(rn)) * 100.0
                dblOKRateFT(0) = gFTNx_cnt(0) / (gFTNx_cnt(0) + gFTNg_cnt(0)) * 100.0
                dblNGRateFT(0) = gFTNg_cnt(0) / (gFTNx_cnt(0) + gFTNg_cnt(0)) * 100.0
            End If

        Catch ex As Exception
            Call Z_PRINT("frmDistribution.StatisticalDataSave() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub

    ''' <summary>
    ''' 統計データの取得
    ''' </summary>
    ''' <param name="JudgeMode"></param>
    ''' <param name="rn">抵抗番号</param>
    ''' <param name="dMin">最小値</param>
    ''' <param name="dMax">最大値</param>
    ''' <param name="dAve">平均値</param>
    ''' <param name="dDev">標準偏差</param>
    ''' <remarks></remarks>
    Public Sub StatisticalDataGet(ByVal JudgeMode As Integer, ByRef rn As Integer, ByRef dMin As Double, ByRef dMax As Double, ByRef dAve As Double, ByRef dDev As Double)
        Try
            giDistributionResNo = rn

            If (JudgeMode = INITIAL_TEST) Then

                If Double.MaxValue = dblMinIT(giDistributionResNo) Then
                    dMin = 0.0
                Else
                    dMin = dblMinIT(giDistributionResNo)            ' 最小値
                End If
                If Double.MinValue = dblMaxIT(giDistributionResNo) Then
                    dMax = 0.0
                Else
                    dMax = dblMaxIT(giDistributionResNo)            ' 最大値
                End If

                'OK/NG数の表示
                'lblGoodChip.Text = CStr(gITNx_cnt(giDistributionResNo))                      ' OK数
                'lblNgChip.Text = CStr(gITNg_cnt(giDistributionResNo))                        ' NG数
                'lblOKRate.Text = dblOKRateIT(giDistributionResNo).ToString("0.000")          ' 良品率
                'lblNGRate.Text = dblNGRateIT(giDistributionResNo).ToString("0.000")          ' 不良品率
            Else
                If Double.MaxValue = dblMinFT(giDistributionResNo) Then
                    dMin = 0.0
                Else
                    dMin = dblMinFT(giDistributionResNo)               ' 最小値
                End If
                If Double.MinValue = dblMaxFT(giDistributionResNo) Then
                    dMax = 0.0
                Else
                    dMax = dblMaxFT(giDistributionResNo)               ' 最大値
                End If

                'OK/NG数の表示
                'lblGoodChip.Text = CStr(gFTNx_cnt(giDistributionResNo))                      ' OK数
                'lblNgChip.Text = CStr(gFTNg_cnt(giDistributionResNo))                        ' NG数
                'lblOKRate.Text = dblOKRateFT(giDistributionResNo).ToString("0.000")          ' 良品率
                'lblNGRate.Text = dblNGRateFT(giDistributionResNo).ToString("0.000")          ' 不良品率
            End If


            If (JudgeMode = INITIAL_TEST) Then
                dAve = dblAverageIT(giDistributionResNo)    ' 平均(IT)
            Else
                dAve = dblAverageFT(giDistributionResNo)    ' 平均(FT)
            End If

            If (JudgeMode = INITIAL_TEST) Then
                dDev = dblDeviationIT(giDistributionResNo) ' 標準偏差(IT)
            Else
                dDev = dblDeviationFT(giDistributionResNo) ' 標準偏差(FT)
            End If

        Catch ex As Exception
            Call Z_PRINT("frmDistribution.StatisticalDataGet() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Sub

    ' 'V2.2.0.0⑯↓
    ''' <summary>
    ''' 統計データの取得
    ''' </summary>
    ''' <param name="JudgeMode"></param>
    ''' <param name="rn">抵抗番号</param>
    ''' <param name="dMin">最小値</param>
    ''' <param name="dMax">最大値</param>
    ''' <param name="dAve">平均値</param>
    ''' <param name="dDev">標準偏差</param>
    ''' <remarks></remarks>
    Public Sub StatisticalDataGetMulti(ByVal JudgeMode As Integer, ByRef rn As Integer, ByRef dMin As Double, ByRef dMax As Double, ByRef dAve As Double, ByRef dDev As Double, ByVal MultiBlkNo As Integer)

        Try
            giDistributionResNo = rn

            With stToTalDataMulti(MultiBlkNo)

                If (JudgeMode = INITIAL_TEST) Then

                    If Double.MaxValue = .dblMinIT(giDistributionResNo) Then
                        dMin = 0.0
                    Else
                        dMin = .dblMinIT(giDistributionResNo)            ' 最小値
                    End If
                    If Double.MinValue = .dblMaxIT(giDistributionResNo) Then
                        dMax = 0.0
                    Else
                        dMax = .dblMaxIT(giDistributionResNo)            ' 最大値
                    End If

                    'OK/NG数の表示
                    'lblGoodChip.Text = CStr(gITNx_cnt(giDistributionResNo))                      ' OK数
                    'lblNgChip.Text = CStr(gITNg_cnt(giDistributionResNo))                        ' NG数
                    'lblOKRate.Text = dblOKRateIT(giDistributionResNo).ToString("0.000")          ' 良品率
                    'lblNGRate.Text = dblNGRateIT(giDistributionResNo).ToString("0.000")          ' 不良品率
                Else
                    If Double.MaxValue = .dblMinFT(giDistributionResNo) Then
                        dMin = 0.0
                    Else
                        dMin = .dblMinFT(giDistributionResNo)               ' 最小値
                    End If
                    If Double.MinValue = .dblMaxFT(giDistributionResNo) Then
                        dMax = 0.0
                    Else
                        dMax = .dblMaxFT(giDistributionResNo)               ' 最大値
                    End If

                    'OK/NG数の表示
                    'lblGoodChip.Text = CStr(gFTNx_cnt(giDistributionResNo))                      ' OK数
                    'lblNgChip.Text = CStr(gFTNg_cnt(giDistributionResNo))                        ' NG数
                    'lblOKRate.Text = dblOKRateFT(giDistributionResNo).ToString("0.000")          ' 良品率
                    'lblNGRate.Text = dblNGRateFT(giDistributionResNo).ToString("0.000")          ' 不良品率
                End If


                If (JudgeMode = INITIAL_TEST) Then
                    dAve = .dblAverageIT(giDistributionResNo)    ' 平均(IT)
                Else
                    dAve = .dblAverageFT(giDistributionResNo)    ' 平均(FT)
                End If

                If (JudgeMode = INITIAL_TEST) Then
                    dDev = .dblDeviationIT(giDistributionResNo) ' 標準偏差(IT)
                Else
                    dDev = .dblDeviationFT(giDistributionResNo) ' 標準偏差(FT)
                End If

            End With


        Catch ex As Exception
            Call Z_PRINT("frmDistribution.StatisticalDataGetMulti() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Sub

    ''' <summary>
    ''' NGカウンター
    ''' </summary>
    ''' <returns></returns>
    Public Function CalcNgCounter() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then

                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.NG_Counter = .stCounter1.ITHigh + .stCounter1.ITLow + .stCounter1.ITOpen + .stCounter1.FTHigh + .stCounter1.FTLow + .stCounter1.FTOpen + .stCounter1.Pattern + .stCounter1.VaNG + .stCounter1.StdNg
                    .stCounter1.Total_NG_Counter = .stCounter1.Total_ITHigh + .stCounter1.Total_ITLow + .stCounter1.Total_ITOpen + .stCounter1.Total_FTHigh + .stCounter1.Total_FTLow + .stCounter1.Total_FTOpen + .stCounter1.Total_Pattern + .stCounter1.Total_VaNG + .stCounter1.Total_StdNg
                End With

            End If

        Catch ex As Exception

        End Try

    End Function
    ' 'V2.2.0.0⑯↑

    ' 'V2.2.0.0⑯↓
    ''' <summary>
    ''' OKカウンターの更新 複数抵抗値用のカウンター
    ''' </summary>
    ''' <returns></returns>
    Public Function SetOkCounterMulti() As Integer
        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.OK_Counter = .stCounter1.OK_Counter + 1
                    .stCounter1.Total_OK_Counter = .stCounter1.Total_OK_Counter + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' IT-HIカウンターの更新 複数抵抗値用のカウンター
    ''' </summary>
    ''' <returns></returns>
    Public Function SetITHICounterMulti() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.ITHigh = .stCounter1.ITHigh + 1
                    .stCounter1.Total_ITHigh = .stCounter1.Total_ITHigh + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' IT-LOカウンターの更新 複数抵抗値用のカウンター
    ''' </summary>
    ''' <returns></returns>
    Public Function SetITLowCounterMulti() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.ITLow = .stCounter1.ITLow + 1
                    .stCounter1.Total_ITLow = .stCounter1.Total_ITLow + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' IT-LOカウンターの更新 複数抵抗値用のカウンター
    ''' </summary>
    ''' <returns></returns>
    Public Function SetITLOCounterMulti() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.ITLow = .stCounter1.ITLow + 1
                    .stCounter1.Total_ITLow = .stCounter1.Total_ITLow + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function


    ''' <summary>
    ''' IT-OPENカウンターの更新 複数抵抗値用のカウンター
    ''' </summary>
    ''' <returns></returns>
    Public Function SetITOpenCounterMulti() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.ITOpen = .stCounter1.ITOpen + 1
                    .stCounter1.Total_ITOpen = .stCounter1.Total_ITOpen + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' FT-HIカウンターの更新 複数抵抗値用のカウンター
    ''' </summary>
    ''' <returns></returns>
    Public Function SetFTHighCounterMulti() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.FTHigh = .stCounter1.FTHigh + 1
                    .stCounter1.Total_FTHigh = .stCounter1.Total_FTHigh + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' FT-LOカウンターの更新 複数抵抗値用のカウンター
    ''' </summary>
    ''' <returns></returns>
    Public Function SetFTLOCounterMulti() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.FTLow = .stCounter1.FTLow + 1
                    .stCounter1.Total_FTLow = .stCounter1.Total_FTLow + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' FT-Openカウンターの更新 複数抵抗値用のカウンター
    ''' </summary>
    ''' <returns></returns>
    Public Function SetFTOpenCounterMulti() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.FTOpen = .stCounter1.FTOpen + 1
                    .stCounter1.Total_FTOpen = .stCounter1.Total_FTOpen + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' Patternカウンターの更新 複数抵抗値用のカウンター
    ''' </summary>
    ''' <returns></returns>
    Public Function SetPatternCounterMulti() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.Pattern = .stCounter1.Pattern + 1
                    .stCounter1.Total_Pattern = .stCounter1.Total_Pattern + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' 変動量(VaNG)カウンターの更新 複数抵抗値用のカウンター
    ''' </summary>
    ''' <returns></returns>
    Public Function SetVaNGCounterMulti() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.VaNG = .stCounter1.VaNG + 1
                    .stCounter1.Total_VaNG = .stCounter1.Total_VaNG + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' 温度センサースタンダード測定NG(StdNg)カウンターの更新 複数抵抗値用のカウンター
    ''' </summary>
    ''' <returns></returns>
    Public Function SetStdNgCounterMulti() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.StdNg = .stCounter1.StdNg + 1
                    .stCounter1.Total_StdNg = .stCounter1.Total_StdNg + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' 上昇率判定NG(ValLow)カウンターの更新 複数抵抗値用のカウンター
    ''' </summary>
    ''' <returns></returns>
    Public Function SetValLowCounterMulti() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.ValLow = .stCounter1.ValLow + 1
                    .stCounter1.Total_ValLow = .stCounter1.Total_ValLow + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function


    ''' <summary>
    ''' トリミングカウンターの更新 複数抵抗値用のカウンター
    ''' </summary>
    ''' <returns></returns>
    Public Function SetTrimCounterMulti() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                With stToTalDataMulti(stExecBlkData.DataNo)
                    .stCounter1.TrimCounter = .stCounter1.TrimCounter + 1
                    .stCounter1.Total_TrimCounter = .stCounter1.Total_TrimCounter + 1
                End With
            End If

        Catch ex As Exception

        End Try

    End Function
    ' 'V2.2.0.0⑯↑

    'V2.2.0.0⑯↓


    ''' <summary>
    ''' マルチブロックの抵抗値カウンタのクリア
    ''' </summary>
    ''' <returns></returns>
    Public Function ClearMultiLotCountData() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                For blk As Integer = 0 To MAXBlock
                    With stToTalDataMulti(blk)

                        .stCounter1.TrimCounter = 0                   ' ﾄﾘﾐﾝｸﾞ数(ﾜｰｸ投入数)
                        .stCounter1.OK_Counter = 0                    ' OK数
                        .stCounter1.NG_Counter = 0                    ' NG数
                        .stCounter1.ITHigh = 0                        ' 初期測定上限値異常
                        .stCounter1.ITLow = 0                         ' 初期測定下限値異常
                        .stCounter1.ITOpen = 0                        ' 測定値異常
                        .stCounter1.FTHigh = 0                        ' 最終測定上限値異常
                        .stCounter1.FTLow = 0                         ' 最終測定下限値異常
                        .stCounter1.FTOpen = 0                        ' 測定値異常
                        .stCounter1.Pattern = 0                       ' カット位置補正の判定 'V1.2.0.0③
                        .stCounter1.VaNG = 0                          ' 再測定変化量エラーV2.0.0.0②
                        .stCounter1.StdNg = 0                         ' スタンダード抵抗測定エラー 'V2.0.0.0⑮
                        .stCounter1.ValLow = 0                        ' カット毎上昇率変化Low異常        'V2.2.0.029
                        .stCounter1.ValHigh = 0                       ' カット毎上昇率変化High異常       'V2.2.0.029
                        ' ロット通算
                        .stCounter1.PlateCounter = 0                  ' 基板カウンター
                        .stCounter1.Total_TrimCounter = 0             ' 抵抗トータル処理数
                        .stCounter1.Total_OK_Counter = 0              ' OK数
                        .stCounter1.Total_NG_Counter = 0              ' NG数
                        .stCounter1.Total_ITHigh = 0                  ' 初期測定上限値異常
                        .stCounter1.Total_ITLow = 0                   ' 初期測定下限値異常
                        .stCounter1.Total_ITOpen = 0                  ' 測定値異常
                        .stCounter1.Total_FTHigh = 0                  ' 最終測定上限値異常
                        .stCounter1.Total_FTLow = 0                   ' 最終測定下限値異常
                        .stCounter1.Total_FTOpen = 0                  ' 測定値異常
                        .stCounter1.Total_Pattern = 0                 ' カット位置補正の判定 'V1.2.0.0③
                        .stCounter1.Total_VaNG = 0                    ' 再測定変化量エラーV2.0.0.0②
                        .stCounter1.Total_StdNg = 0                   ' スタンダード抵抗測定エラー 'V2.0.0.0⑮
                        .stCounter1.Total_ValLow = 0                  ' カット毎上昇率変化Low異常        'V2.2.0.029
                        .stCounter1.Total_ValHigh = 0                 ' カット毎上昇率変化High異常       'V2.2.0.029

                        For cnt As Integer = 0 To MAX_RES_USER
                            .TrimCounter(cnt) = 0                    ' トリミング数カウンター
                            .Total_TrimCounter(cnt) = 0              ' トリミング数カウンタートータル 
                        Next cnt

                    End With
                Next blk
            End If

        Catch ex As Exception

        End Try

    End Function
    'V2.2.0.0⑯↑

    'V2.2.0.0⑯↓


    ''' <summary>
    ''' マルチブロックの抵抗値カウンタのクリア
    ''' </summary>
    ''' <returns></returns>
    Public Function ClearMultiCountPlateData() As Integer

        Try

            If stMultiBlock.gMultiBlock <> 0 Then
                For blk As Integer = 0 To MAXBlock
                    With stToTalDataMulti(blk)

                        .stCounter1.TrimCounter = 0                   ' ﾄﾘﾐﾝｸﾞ数(ﾜｰｸ投入数)
                        .stCounter1.OK_Counter = 0                    ' OK数
                        .stCounter1.NG_Counter = 0                    ' NG数
                        .stCounter1.ITHigh = 0                        ' 初期測定上限値異常
                        .stCounter1.ITLow = 0                         ' 初期測定下限値異常
                        .stCounter1.ITOpen = 0                        ' 測定値異常
                        .stCounter1.FTHigh = 0                        ' 最終測定上限値異常
                        .stCounter1.FTLow = 0                         ' 最終測定下限値異常
                        .stCounter1.FTOpen = 0                        ' 測定値異常
                        .stCounter1.Pattern = 0                       ' カット位置補正の判定 'V1.2.0.0③
                        .stCounter1.VaNG = 0                          ' 再測定変化量エラーV2.0.0.0②
                        .stCounter1.StdNg = 0                         ' スタンダード抵抗測定エラー 'V2.0.0.0⑮
                        .stCounter1.ValLow = 0                        ' カット毎上昇率変化Low異常        'V2.2.0.029
                        .stCounter1.ValHigh = 0                       ' カット毎上昇率変化High異常       'V2.2.0.029

                        For cnt As Integer = 0 To MAX_RES_USER
                            .TrimCounter(cnt) = 0                    ' トリミング数カウンター
                        Next cnt

                    End With
                Next blk
            End If

        Catch ex As Exception

        End Try

    End Function
    'V2.2.0.0⑯↑


#End Region

End Class