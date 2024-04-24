'===============================================================================
'   Description  : ＴＸ/ＴＹティーチング画面処理
'
'   Copyright(C) : TOWA LASERFRONT CORP. 2011
'
'===============================================================================
Option Strict Off
Option Explicit On

Friend Class frmTxTyTeach
    Inherits System.Windows.Forms.Form
    Implements ICommonMethods              'V2.2.0.0①

#Region "プライベート変数定義"
    Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Integer) As Integer '###228

    '===========================================================================
    '   変数定義
    '===========================================================================
    '----- オブジェクト定義 -----
    Private objTxTyMsgbox As frmTxTyMsgbox                          ' TX/TYﾃｨﾁﾝｸﾞ終了確認画面ｵﾌﾞｼﾞｪｸﾄ 
    Private stJOG As JOG_PARAM                                      ' 矢印画面(BPのJOG操作)用パラメータ
    Private mExit_flg As Short                                      ' ＴＸ/ＴＹティーチング結果

    '----- データグリッドビュー用 -----
    Private RowAry() As DataGridViewRow                             ' Row(行)オブジェクト配列
    Private ColAry() As DataGridViewColumn                          ' Col(列)オブジェクト配列

    Private miStepBlock(MaxCntStep) As Short                        ' 変更前のﾌﾞﾛｯｸ数
    Private mdStepInterval(MaxCntStep) As Double                    ' 変更前のｽﾃｯﾌﾟ間ｲﾝﾀｰﾊﾞﾙ
    'Private tmpStepInterval(MaxCntStep) As Double                   ' 変更後のｽﾃｯﾌﾟ間ｲﾝﾀｰﾊﾞﾙ
    Private ZRPosX(1 + (MaxCntStep * 2)) As Double                  ' ずれ量X(0:第1基準点　1:第2基準点)
    Private ZRPosY(1 + (MaxCntStep * 2)) As Double                  ' ずれ量Y(0:第1基準点　1:第2基準点)

    '----- ステップオフセット量算出域(TX/TY共通) -----              ' ###091
    Private StepOffSetX(2) As Double                                ' オフセット位置X(0:第1基準点　1:第2基準点)
    Private StepOffSetY(2) As Double                                ' オフセット位置X(0:第1基準点　1:第2基準点)

    'Private mdBlockCnt(MaxCntResist) As Double                      ' 変更前のブロック数
    'Private mdStepBlockX(MaxCntResist) As Double                    ' 変更前のブロックインターバルX
    'Private mdStepBlockY(MaxCntResist) As Double                    ' 変更前のブロックインターバルY
    'Private tmpStepBlockX(MaxCntResist) As Double                   ' 変更後のブロックインターバルX
    'Private tmpStepBlockY(MaxCntResist) As Double                   ' 変更後のブロックインターバルY

    '----- その他 -----
    Private dblTchMoval(3) As Double                                ' ﾋﾟｯﾁ退避域(0=ﾋﾟｯﾁ, 1=HIGHﾋﾟｯﾁ, 2=Pause Time(Sec))

    ' 以下は未使用のため削除 ?
    'Private ChipSize As Double                          ' 補正後のﾁｯﾌﾟｻｲｽﾞ
    'Private Ino As Short                                ' 処理中のｲﾝﾀｰﾊﾞﾙNo
    'Private TBLx As Double                                          ' ﾃｰﾌﾞﾙ移動座標X(絶対値)
    'Private TBLy As Double                                          ' ﾃｰﾌﾞﾙ移動座標Y(絶対値)
    'Private KJPosX As Double                                        ' 基準位置X(第1ﾌﾞﾛｯｸの先頭基準位置)
    'Private KJPosY As Double                                        ' 基準位置Y(第1ﾌﾞﾛｯｸの先頭基準位置)
    'Private Pos2ndX As Double                                       ' 基準位置X(第1ﾌﾞﾛｯｸの最終基準位置)
    'Private Pos2ndY As Double                                       ' 基準位置Y(第1ﾌﾞﾛｯｸの最終基準位置)
    'Private EffBN As Short                              ' 有効ﾌﾞﾛｯｸ数
    'Private EffGN As Short                              ' 有効ｸﾞﾙｰﾌﾟ数

    '----- トリミングパラメータ -----
    'Private mdTbOffx As Double                          ' ﾃｰﾌﾞﾙ位置ｵﾌｾｯﾄX
    'Private mdTbOffy As Double                          ' ﾃｰﾌﾞﾙ位置ｵﾌｾｯﾄY
    'Private mdBpOffx As Double                          ' BP位置ｵﾌｾｯﾄX
    'Private mdBpOffy As Double                          ' BP位置ｵﾌｾｯﾄY
    'Private mdAdjx As Double                            ' ｱｼﾞｬｽﾄ位置ｵﾌｾｯﾄX
    'Private mdAdjy As Double                            ' ｱｼﾞｬｽﾄ位置ｵﾌｾｯﾄY
    'Private miBNx As Short                              ' ﾌﾞﾛｯｸ数X
    'Private miBNy As Short                              ' ﾌﾞﾛｯｸ数Y
    'Private miCdir As Short                             ' ﾁｯﾌﾟ並び方向
    'Private miChpNum As Short                           ' ﾁｯﾌﾟ数
    'Private mdCSx As Double                             ' ﾁｯﾌﾟｻｲｽﾞX
    'Private mdCSy As Double                             ' ﾁｯﾌﾟｻｲｽﾞY
    'Private miGNx As Short                              ' ｸﾞﾙｰﾌﾟ数X
    'Private miGNy As Short                              ' ｸﾞﾙｰﾌﾟ数Y
    'Private mdBSx As Double                             ' ﾌﾞﾛｯｸｻｲｽﾞX
    'Private mdBSy As Double                             ' ﾌﾞﾛｯｸｻｲｽﾞY
    'Private dStepOffx As Double                         ' ｽﾃｯﾌﾟｵﾌｾｯﾄ量X
    'Private dStepOffy As Double                         ' ｽﾃｯﾌﾟｵﾌｾｯﾄ量Y
    'Private miPP30 As Short                             ' 補正ﾓｰﾄﾞ
    'Private miPP31 As Short                             ' 補正方法
    'Private mdSpx As Double                             ' ｽﾀｰﾄﾎﾟｲﾝﾄX
    'Private mdSpy As Double                             ' ｽﾀｰﾄﾎﾟｲﾝﾄY

#End Region

#Region "終了結果を返す"
    '''=========================================================================
    '''　<summary>終了結果を返す</summary>
    '''　<returns>cFRS_ERR_START = OK(STARTｷｰ)
    '''  　       cFRS_ERR_RST   = Cancel(RESETｷｰ)
    '''    　     cFRS_TxTy      = TX2(Teach)/TY2押下
    '''      　   -1以下         = エラー</returns>
    '''=========================================================================
    Public ReadOnly Property sGetReturn() As Short
        Get
            sGetReturn = mExit_flg
        End Get
    End Property
#End Region

    '=========================================================================
    '   フォームの初期化/終了処理
    '=========================================================================
#Region "Form Initialize時処理"
    '''=========================================================================
    '''<summary>Form Initialize時処理</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub Form_Initialize_Renamed()

        Dim strMSG As String

        Try
            ' データグリッドビュー用オブジェクト配列初期化
            RowAry = Nothing                                        ' Row(行)オブジェクト配列
            ColAry = Nothing                                        ' Col(列)オブジェクト配列

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTyTeach.Form_Initialize_Renamed() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "Form 終了時処理"
    '''=========================================================================
    '''<summary>Form 終了時処理</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub frmTxTeach_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        Dim strMSG As String

        Try
            ' データグリッドビューオブジェクトの開放
            Call TermGridView(RowAry, ColAry)

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTyTeach.frmTxTeach_FormClosed() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "Form Load時処理"
    '''=========================================================================
    '''<summary>Form Load時処理</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub frmTxTeach_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Dim strMSG As String

        Try
            ' 初期処理
            Call SetDispMsg()                                       ' ラベル等を設定する(日本語/英語)
            mExit_flg = -1                                          ' 終了フラグ = 初期化

            ' ｳｲﾝﾄﾞｳをｺﾝﾄﾛｰﾙに重ねる
            'Me.Top = Form1.Text4.Top + DISPOFF_SUBFORM_TOP
            'Me.Left = Form1.Text4.Left
            Me.Location = New Point(Form1.VideoLibrary1.Location.X + Form1.VideoLibrary1.Size.Width + 6, Form1.Grpcmds.Location.Y)

            ' トリミングデータより必要なパラメータを取得する
            Call SetTrimData()

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTyTeach.frmTxTeach_Load() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "ラベル等を設定する(日本語/英語)"
    '''=========================================================================
    '''<summary>ラベル等を設定する(日本語/英語)</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub SetDispMsg()

        Dim strMSG As String

        Try
            ' ラベル等を設定する
            With Me
                If (giAppMode = APP_MODE_TX) Then
                    ' ＴＸティーチング時
                    ' label
                    .lblTitle.Text = TITLE_TX                       ' ﾀｲﾄﾙ
                    .lblTitle2.Text = ""
                    .LblDisp_8.Text = LBL_TXTY_TEACH_04             ' 補正比率
                    .LblDisp_9.Text = LBL_TXTY_TEACH_05             ' チップサイズ(mm)
                    .LblDisp_10.Text = LBL_TXTY_TEACH_09            ' グループインターバル
                Else
                    ' ＴＹティーチング時
                    ' label
                    .lblTitle.Text = TITLE_TY                       ' ﾀｲﾄﾙ
                    .lblTitle2.Text = ""
                    .LblDisp_8.Text = LBL_TXTY_TEACH_04             ' 補正比率
                    .LblDisp_9.Text = LBL_TXTY_TEACH_05             ' チップサイズ(mm)
                    .LblDisp_10.Text = LBL_TXTY_TEACH_11            ' ステップインターバル
                End If

                ' ＴＸ/ＴＹティーチング共通
                ' frame
                .GrpFram_0.Text = LBL_TXTY_TEACH_12                 ' 第１基準点
                .GrpFram_1.Text = LBL_TXTY_TEACH_13                 ' 第２基準点
                .GrpFram_2.Text = LBL_TXTY_TEACH_03                 ' 補正量
                ' button
                .cmdOK.Text = "OK"
                .cmdCancel.Text = CMD_CANCEL
            End With

            ' 表示関連初期化
            Call InitDisp()

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTyTeach.SetDispMsg() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "表示関連初期化"
    '''=========================================================================
    '''<summary>表示関連初期化</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub InitDisp()

        Dim strMSG As String

        Try
            ' 座標表示クリア
            With Me
                Me.TxtPosX.Text = ""                                ' 第1基準点XY
                Me.TxtPosY.Text = ""
                Me.TxtPos2X.Text = ""                               ' 第2基準点XY
                Me.TxtPos2Y.Text = ""

                Me.LblResult_0.Text = ""                            ' 補正比率
                Me.LblResult_1.Text = ""                            ' チップサイズ
                Me.LblResult_2.Text = ""                            ' チップサイズ(補正後)
            End With

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTyTeach.InitDisp() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "Form Activated時処理"
    '''=========================================================================
    '''<summary>Form Activated時処理</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub frmTxTeach_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated

        'V2.2.0.0①↓　Execute()で実行
        'Dim strMSG As String

        'Try
        '    ' ＴＸ/ＴＹティーチング開始
        '    If (mExit_flg <> -1) Then Exit Sub
        '    mExit_flg = 0                                           ' 終了フラグ = 0
        '    mExit_flg = TxTyMainProc()                              ' ＴＸまたはＴＹティーチング開始

        '    ' 補正クロスラインを非表示とする
        '    Form1.CrosLineX.Visible = False
        '    Form1.CrosLineY.Visible = False

        '    ' トラップエラー発生時
        'Catch ex As Exception
        '    strMSG = "frmTxTeach.frmTxTeach_Activated() TRAP ERROR = " + ex.Message
        '    MsgBox(strMSG)
        '    mExit_flg = cERR_TRAP                                   ' Return値 = 例外エラー
        'End Try
        'Me.Close()
        'V2.2.0.0①↑

    End Sub
#End Region

    '=========================================================================
    '   ＴＸまたはＴＹティーチング処理
    '=========================================================================
#Region "終了確認メッセージ表示処理"
    '''=========================================================================
    '''<summary>終了確認メッセージ表示処理</summary>
    '''<remarks>Exit_flg = cFRS_ERR_START(OK(STARTｷｰ))
    '''                    cFRS_ERR_RST(Cancel(RESETｷｰ))
    '''                    cFRS_TxTy(TX2/TY2押下)</remarks>
    '''=========================================================================
    Private Sub Disp_TxTyMsgBox(ByRef Exit_flg As Short)

        Dim strMSG As String

        Try
            If (giAppMode = APP_MODE_TX) Then
                MSG_EXECUTE_TXTYLABEL = "Teach"
            Else
                MSG_EXECUTE_TXTYLABEL = "TY2"
            End If
            objTxTyMsgbox = New frmTxTyMsgbox()                 ' TX/TYﾃｨﾁﾝｸﾞ終了確認画面ｵﾌﾞｼﾞｪｸﾄ生成 

            ' "前の画面に戻ります。よろしいですか？"
            objTxTyMsgbox.ShowDialog()                          ' 終了確認メッセージ表示 
            Exit_flg = objTxTyMsgbox.sGetReturn()               ' Exit_flg = 設定

            Call objTxTyMsgbox.Close()                          ' TX/TYﾃｨﾁﾝｸﾞ終了確認画面ｵﾌﾞｼﾞｪｸﾄ開放
            Call objTxTyMsgbox.Dispose()                        ' リソース開放
            objTxTyMsgbox = Nothing

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTyTeach.Disp_TxTyMsgBox() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "OKﾎﾞﾀﾝ押下時処理"
    '''=========================================================================
    ''' <summary>OKﾎﾞﾀﾝ押下時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub cmdOK_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles cmdOK.MouseUp  'V1.16.0.0⑬
        'Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

        stJOG.Flg = cFRS_ERR_START                                      ' OK(STARTｷｰ)

    End Sub
#End Region

#Region "ｷｬﾝｾﾙﾎﾞﾀﾝ押下時処理"
    '''=========================================================================
    ''' <summary>ｷｬﾝｾﾙﾎﾞﾀﾝ押下時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub cmdCancel_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles cmdCancel.MouseUp 'V1.16.0.0⑬
        'Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

        stJOG.Flg = cFRS_ERR_RST                           ' Cancel(RESETｷｰ)

    End Sub
#End Region

#Region "ＴＸまたはＴＹティーチングのメイン処理"
    '''=========================================================================
    '''<summary>ＴＸまたはＴＹティーチングのメイン処理"</summary>
    '''<returns>cFRS_ERR_START=OK(STARTｷｰ)
    '''         cFRS_ERR_RST  =Cancel(RESETｷｰ)
    '''         cFRS_TxTy     =TX2/TY2押下
    '''         -1以下        =エラー</returns>
    '''=========================================================================
    Private Function TxTyMainProc() As Short

        Dim i As Short
        'Dim rtn As Short                                    ' TxTyMainProc戻値 
        Dim GrpNum As Short                                 ' ｸﾞﾙｰﾌﾟ数
        Dim RnBn As Short                                   ' チップ数(抵抗数)(TX時), ブロック数(TY時)
        Dim GrpCnt As Short                                 ' ｸﾞﾙｰﾌﾟ数ｶｳﾝﾀ
        Dim ChipSize As Double                              ' 補正後のﾁｯﾌﾟｻｲｽﾞ
        Dim tmpChipSize As Double                           ' ﾁｯﾌﾟｻｲｽﾞ退避域
        Dim CSPoint(1) As Double                            ' ﾁｯﾌﾟｻｲｽﾞ算出用(0:第1基準点, 1:第2基準点)
        Dim mdBpOffx As Double                              ' BP位置ｵﾌｾｯﾄX
        Dim mdBpOffy As Double                              ' BP位置ｵﾌｾｯﾄY
        Dim mdAdjx As Double                                ' ｱｼﾞｬｽﾄ位置X(未使用)
        Dim mdAdjy As Double                                ' ｱｼﾞｬｽﾄ位置Y(未使用)
        Dim dStepOffx As Double                             ' ｽﾃｯﾌﾟｵﾌｾｯﾄ量X
        Dim dStepOffy As Double                             ' ｽﾃｯﾌﾟｵﾌｾｯﾄ量Y
        Dim mdSpx As Double                                 ' ｽﾀｰﾄﾎﾟｲﾝﾄX
        Dim mdSpy As Double                                 ' ｽﾀｰﾄﾎﾟｲﾝﾄY
        Dim mdBSx As Double                                 ' ﾌﾞﾛｯｸｻｲｽﾞX
        Dim mdBSy As Double                                 ' ﾌﾞﾛｯｸｻｲｽﾞY
        Dim KJPosX As Double                                ' 先頭ブロックの先頭基準位置(BP位置X/ﾃｰﾌﾞﾙ移動座標X)
        Dim KJPosY As Double                                ' 先頭ブロックの先頭基準位置(BP位置Y/ﾃｰﾌﾞﾙ移動座標Y)
        Dim r As Short
        Dim StepNum As Integer = 0
        Dim strMSG As String

        Try
            '-------------------------------------------------------------------
            '   初期処理
            '-------------------------------------------------------------------
            ' グループ数, ブロック数(TY時)またはチップ数(抵抗数)(TX時), チップサイズを取得する
            Me.cmdOK.Enabled = False                                    ' 第２基準へ行くまで有効としない'V3.0.0.0③
            r = GetChipNumAndSize(giAppMode, GrpNum, RnBn, ChipSize)
            tmpChipSize = ChipSize                                      ' チップサイズ退避
            GrpCnt = 0                                                  ' ｸﾞﾙｰﾌﾟ数ｶｳﾝﾀ
            mdAdjx = 0.0 : mdAdjy = 0.0                                 ' ｱｼﾞｬｽﾄ位置X(未使用)
            mdBpOffx = typPlateInfo.dblBpOffSetXDir                     ' BP位置ｵﾌｾｯﾄX,Y設定
            mdBpOffy = typPlateInfo.dblBpOffSetYDir
            dStepOffx = typPlateInfo.dblStepOffsetXDir                  ' ｽﾃｯﾌﾟｵﾌｾｯﾄ量X,Y(TYﾃｨｰﾁ用)
            dStepOffy = typPlateInfo.dblStepOffsetYDir
            Call CalcBlockSize(mdBSx, mdBSy)                            ' ブロックサイズXY設定

            ' データグリッドビューオブジェクトの生成
            Call InitGridView(GridView, GrpNum, RowAry, ColAry)         ' グループ数｢2｣以上の時にｸﾞﾘｯﾄﾞ表示する

STP_RETRY:
            ' 座標表示等クリア
            For i = 0 To 1 + MaxCntStep                                 ' グループインターバル初期化
                ZRPosX(i) = 0.0#
                ZRPosY(i) = 0.0#
            Next i
            For i = 0 To 1                                              ' ステップオフセット量算出域初期化 ###091
                StepOffSetX(i) = 0.0#
                StepOffSetY(i) = 0.0#
            Next i

            Call InitDisp()                                             ' 座標表示クリア
            Call ClearGridView(RowAry)                                  ' データグリッドビューのインターバル表示をクリアする
            ChipSize = tmpChipSize                                      ' チップサイズ設定

            ' BPを第1ﾌﾞﾛｯｸ、第1抵抗ｽﾀｰﾄﾎﾟｲﾝﾄに移動する
            If (giAppMode = APP_MODE_TY) Then
                Call BpMoveOrigin_Ex()                                  ' BPをﾌﾞﾛｯｸの右上に移動する
                Call XYTableMoveTopBlock(KJPosX, KJPosY)                ' 先頭ﾌﾞﾛｯｸの先頭基準テーブル位置X,Yを取得する
            End If
            Call GetCutStartPoint(1, 1, mdSpx, mdSpy)                   ' 第一抵抗の第一カットのｽﾀｰﾄﾎﾟｲﾝﾄを取得
            r = Form1.System1.EX_MOVE(gSysPrm, mdSpx, mdSpy, 1)         ' BPを第一抵抗の第一カットのｽﾀｰﾄﾎﾟｲﾝﾄに移動
            If (r < cFRS_NORMAL) Then
                Return (r)                                              ' Return値 = エラー
            End If
            If (giAppMode = APP_MODE_TX) Then
                KJPosX = mdSpx                                          ' KJPosX = 第1ﾌﾞﾛｯｸの先頭基準BP位置X
                KJPosY = mdSpy                                          ' KJPosY = 第1ﾌﾞﾛｯｸの先頭基準BP位置Y
            End If

            ' 矢印画面(BPのJOG操作)用パラメータを初期化する
            If (giAppMode = APP_MODE_TX) Then                           ' TX ? 
                stJOG.Md = MODE_BP                                      ' モード(1:BP移動)
            Else
                stJOG.Md = MODE_STG                                     ' モード(0:XYテーブル移動)
            End If
            stJOG.Md2 = MD2_BUTN                                        ' 入力モード(0:画面ﾎﾞﾀﾝ入力, 1:ｺﾝｿｰﾙ入力)
            '                                                           ' キーの有効(1)/無効(0)指定
            stJOG.Opt = CONSOLE_SW_RESET + CONSOLE_SW_START + CONSOLE_SW_HALT
            stJOG.PosX = KJPosX                                         ' BP X位置/XYﾃｰﾌﾞﾙ X位置
            stJOG.PosY = KJPosY                                         ' BP Y位置/XYﾃｰﾌﾞﾙ Y位置
            stJOG.BpOffX = mdAdjx + mdBpOffx                            ' BPｵﾌｾｯﾄX 
            stJOG.BpOffY = mdAdjy + mdBpOffy                            ' BPｵﾌｾｯﾄY 
            stJOG.BszX = mdBSx                                          ' ﾌﾞﾛｯｸｻｲｽﾞX 
            stJOG.BszY = mdBSy                                          ' ﾌﾞﾛｯｸｻｲｽﾞY
            stJOG.TextX = TxtPosX                                       ' BP X位置/XYﾃｰﾌﾞﾙ X位置表示用ﾃｷｽﾄﾎﾞｯｸｽ
            stJOG.TextY = TxtPosY                                       ' BP Y位置/XYﾃｰﾌﾞﾙ Y位置表示用ﾃｷｽﾄﾎﾞｯｸｽ
            stJOG.cgX = 0.0                                             ' 移動量X 
            stJOG.cgY = 0.0                                             ' 移動量Y 
            stJOG.BtnHI = BtnHI                                         ' HIボタン
            stJOG.BtnZ = BtnZ                                           ' Zボタン
            stJOG.BtnSTART = BtnSTART                                   ' STARTボタン
            stJOG.BtnHALT = BtnHALT                                     ' HALTボタン
            stJOG.BtnRESET = BtnRESET                                   ' RESETボタン
            Call JogEzInit(stJOG, TBarLowPitch, TBarHiPitch, TBarPause, LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval)
            stJOG.Flg = -1                                              ' 親画面のOK/Cancelﾎﾞﾀﾝ押下ﾌﾗｸﾞ
            Call Me.Focus()

STP_CHIPSIZE:
            '-------------------------------------------------------------------
            '   チップサイズＸ(TX時)またはチップサイズＹ(TY時)のティーチング処理
            '   ※CSPoint(0:第1基準点, 1:第2基準点)にBPまたはテーブル位置を設定する
            '-------------------------------------------------------------------
            ' チップサイズのティーチング処理
            r = Sub_Jog1(RnBn, ChipSize, KJPosX, KJPosY, GrpNum, GrpCnt, CSPoint)
            Timer1.Enabled = False                                      ' ###228
            If (r < cFRS_NORMAL) Then                                   ' エラーならエラーリターン
                Return (r)
            End If
            If (r <> cFRS_NORMAL) Then                                  ' 処理継続以外(RESET SW(Cancelﾎﾞﾀﾝ)押下/OKボタン押下)ならトリミングデータ更新へ
                GoTo STP_END
            End If

STP_INTVAL:
            '-------------------------------------------------------------------
            '   ＢＰグループ(サーキット)間隔(TX時)または
            '   ステージグループ間隔(TY時)のティーチング処理
            '     mdStepInterval()にインターバル値を設定する
            '-------------------------------------------------------------------
            ' インターバルのティーチング処理
            r = Sub_Jog2(RnBn, ChipSize, KJPosX, KJPosY, GrpNum, GrpCnt, CSPoint)
            Timer1.Enabled = False                                      ' ###228
            If (r < cFRS_NORMAL) Then                                   ' エラーならエラーリターン
                Return (r)
            End If
            If (r = cFRS_ERR_HALT) Then                                 ' HALT SW押下ならチップサイズティーチング処理へ
                GoTo STP_CHIPSIZE
            End If
            If (r <> cFRS_NORMAL) Then                                  ' 処理継続以外(RESET SW(Cancelﾎﾞﾀﾝ)押下/OKボタン押下)ならトリミングデータ更新へ
                GoTo STP_END
            End If


            '###121
            '-------------------------------------------------------------------
            '   ステップオフセット量の算出(ＴＹティーチング時)
            '   TYティーチング最終ブロックティーチング時のX方向のずれ量がステップオフセットとなる
            '-------------------------------------------------------------------
            ' ステップオフセット量のティーチング処理

            'Nop Nop Nop    2012.11.20 kami del
            'r = Sub_Jog3(RnBn, ChipSize, KJPosX, KJPosY, GrpNum, GrpCnt, CSPoint, dStepOffx, dStepOffy)
            'If (r < cFRS_NORMAL) Then                                   ' エラーならエラーリターン
            '    Return (r)
            'End If
            'If (r = cFRS_ERR_HALT) Then                                 ' HALT SW押下ならステージグループ間隔のティーチング処理へ
            '    GoTo STP_INTVAL
            'End If
            'Nop Nop Nop    2012.11.20 kami del
            '###121


STP_END:
            '-------------------------------------------------------------------
            '   トリミングデータ更新
            '-------------------------------------------------------------------
            If (r <> cFRS_ERR_RST) Then                                 ' Cancel(RESET SW)押下以外 ?
                Call SetTrimParameter(ChipSize, GrpCnt)
                gCmpTrimDataFlg = 1                                     ' データ更新フラグ = 1(更新あり)
            End If
            Return (r)                                                  ' Return値設定

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTyTeach.TxTyMainProc() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Return値 = 例外エラー
        End Try
    End Function
#End Region

#Region "チップサイズＸまたはチップサイズＹのティーチング処理"
    '''=========================================================================
    ''' <summary>チップサイズＸまたはチップサイズＹのティーチング処理</summary>
    ''' <param name="RnBn">    (INP)チップ数(抵抗数)(TX時), ブロック数(TY時)</param>
    ''' <param name="ChipSize">(I/O)チップサイズ</param>
    ''' <param name="KJPosX">  (INP)先頭ブロックの先頭基準位置X</param>
    ''' <param name="KJPosY">  (INP)先頭ブロックの先頭基準位置Y</param>
    ''' <param name="GrpNum">  (INP)グループ数</param>
    ''' <param name="GrpCnt">  (OUT)グループカウンタ</param>
    ''' <param name="CSPoint"> (OUT)ﾁｯﾌﾟｻｲｽﾞ算出用(0:第1基準点, 1:第2基準点)</param>
    ''' <returns>cFRS_ERR_START = OK(STARTｷｰ)押下
    '''          cFRS_ERR_RST   = Cancel(RESETｷｰ)押下
    '''          cFRS_NORMAL    = 正常(グループ間インターバル処理へ)
    '''          -1以下         = エラー</returns>
    '''=========================================================================
    Private Function Sub_Jog1(ByVal RnBn As Short, ByRef ChipSize As Double, ByVal KJPosX As Double, ByVal KJPosY As Double, _
                              ByVal GrpNum As Short, ByRef GrpCnt As Short, ByRef CSPoint() As Double) As Short

        Dim PosX As Double                                              ' 現在のBP位置X/ﾃｰﾌﾞﾙ移動座標X(絶対値)
        Dim PosY As Double                                              ' 現在のBP位置Y/ﾃｰﾌﾞﾙ移動座標Y(絶対値)
        Dim WkChipSize As Double
        Dim iPos As Short
        Dim r As Short
        Dim rtn As Short                                                ' 戻値 
        Dim strMSG As String

        Try
            ' 初期処理
            Timer1.Enabled = True                                       ' ###228
            iPos = 0                                                    ' iPos = 第1基準点(0:第1基準点位置, 1:第2基準点位置)
            GrpCnt = 1                                                  ' ｸﾞﾙｰﾌﾟ数ｶｳﾝﾀ
            WkChipSize = ChipSize                                       ' チップサイズ退避 

            If (giAppMode = APP_MODE_TX) Then                           ' タイトル設定 ###084
                Me.lblTitle2.Text = INFO_MSG13 + INFO_MSG33             ' タイトル = "チップサイズ　ティーチング (TX)"
            Else
                Me.lblTitle2.Text = INFO_MSG34 + INFO_MSG32             ' タイトル = "チップサイズ　ティーチング (TY)"
            End If
            'Me.lblTitle2.Text = INFO_MSG13                             ' タイトル = "チップサイズ　ティーチング"

            stJOG.Flg = -1                                              ' 親画面のOK/Cancelﾎﾞﾀﾝ押下ﾌﾗｸﾞ
            Me.GrpFram_1.Visible = True                                 ' 第二基準点表示(Sub_Jog3()で非表示にするため)
            Call InitDisp()                                             ' 座標表示クリア
            Call ClearGridView(RowAry)                                  ' データグリッドビューのインターバル表示をクリアする
            If (giAppMode = APP_MODE_TY) Then                           ' ＴＹティーチング時はブロック数をステージグループ内ブロック数とする
                If (typPlateInfo.intResistDir = 0) Then                 ' チップ並びはX方向 ?
                    RnBn = typPlateInfo.intBlkCntInStgGrpY              ' ブロック数 = ステージグループ内ブロック数Y 
                Else
                    RnBn = typPlateInfo.intBlkCntInStgGrpX              ' ブロック数 = ステージグループ内ブロック数X
                End If
            End If

STP_RETRY:
            Call Me.Focus()
            Do
                ' BP移動またはXYﾃｰﾌﾞﾙ移動(第1ｸﾞﾙｰﾌﾟ第1基準位置/第2基準位置)
                If (giAppMode = APP_MODE_TX) Then                       ' ＴＸティーチング時
                    ' BPを第1基準位置または第2基準位置移動する
                    r = XYBPMoveSetBlock(iPos, iPos, GrpCnt, RnBn, ChipSize, KJPosX, KJPosY, PosX, PosY)
                Else                                                    ' ＴＹティーチング時
                    ' XYﾃｰﾌﾞﾙを第1基準位置または第2基準位置移動する
                    r = XYTableMoveSetBlock(iPos, iPos, GrpCnt, RnBn, ChipSize, KJPosX, KJPosY, PosX, PosY)
                End If
                If (r <> cFRS_NORMAL) Then                              ' エラー ?
                    Return (r)                                          ' エラーリターン
                End If

                ' 表示メッセージ等を設定する 
                If (iPos = 0) Then                                      ' 第1基準点 ? 
                    '"第1グループ、第1抵抗基準位置のティーチング"+"基準位置を合わせて下さい。"+"移動:[矢印]  決定:[START]  中断:[RESET]" "
                    lblInfo.Text = INFO_MSG18 & vbCrLf & INFO_MSG16 & vbCrLf & INFO_MSG17
                    stJOG.TextX = TxtPosX                               ' 第1基準点座標X位置表示用ﾃｷｽﾄﾎﾞｯｸｽ
                    stJOG.TextY = TxtPosY                               ' 第1基準点座標Y位置表示用ﾃｷｽﾄﾎﾞｯｸｽ
                    TxtPosX.Enabled = True
                    TxtPosY.Enabled = True
                Else                                                    ' 第2基準点 
                    '"第nグループ、最終抵抗基準位置のティーチング"+"基準位置を合わせて下さい。"+"移動:[矢印]  決定:[START]  中断:[RESET]" & vbCrLf & "[HALT]で１つ前の処理に戻ります。"
                    lblInfo.Text = INFO_MSG19 & iPos & INFO_MSG20 & vbCrLf & INFO_MSG16 & vbCrLf & INFO_MSG17
                    stJOG.TextX = TxtPos2X                              ' 第2基準点座標X位置表示用ﾃｷｽﾄﾎﾞｯｸｽ
                    stJOG.TextY = TxtPos2Y                              ' 第2基準点座標Y位置表示用ﾃｷｽﾄﾎﾞｯｸｽ
                    TxtPos2X.Enabled = True
                    TxtPos2Y.Enabled = True
                    Me.cmdOK.Enabled = True                             ' 第２基準へ行くまで有効としない'V3.0.0.0③
                End If

                ' 第1基準点/第2基準点のティーチング"処理
                stJOG.PosX = PosX                                       ' BP XまたはXYﾃｰﾌﾞﾙ X絶対位置
                stJOG.PosY = PosY                                       ' BP YまたはXYﾃｰﾌﾞﾙ Y絶対位置
                stJOG.Flg = -1                                          ' 親画面のOK/Cancelﾎﾞﾀﾝ押下ﾌﾗｸﾞ
                r = JogEzMove(stJOG, gSysPrm, TBarLowPitch, TBarHiPitch, TBarPause, LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval, Me)  'V2.2.0.0① 
                If (r < cFRS_NORMAL) Then                               ' エラーなら終了
                    Return (r)
                End If

                ' 第1基準点/第2基準点のオフセット値更新
                ZRPosX(iPos) = ZRPosX(iPos) + stJOG.cgX
                ZRPosY(iPos) = ZRPosY(iPos) + stJOG.cgY
                ' V1.16.0.1① ADD START-----------------------------------------------------
                If (iPos = 0) Then                                      ' 第1基準点 ? 
                    If (typPlateInfo.intResistDir = 0) Then             ' ﾁｯﾌﾟ並びはX方向 ?
                        If (giAppMode = APP_MODE_TX) Then               ' ＴＸティーチング時
                            ZRPosY(1) = ZRPosY(iPos)
                        End If
                    Else
                        If (giAppMode = APP_MODE_TX) Then               ' ＴＸティーチング時
                            ZRPosX(1) = ZRPosX(iPos)
                        End If
                    End If
                End If
                ' V1.16.0.1① ADD END  -----------------------------------------------------

                ' コンソールキーチェック
                If (r = cFRS_ERR_START) Then                            ' START SW押下 ?
                    ' 第1基準点/第2基準点の座標更新
                    If (typPlateInfo.intResistDir = 0) Then             ' ﾁｯﾌﾟ並びはX方向 ?
                        If (giAppMode = APP_MODE_TX) Then               ' ＴＸティーチング時
                            CSPoint(iPos) = stJOG.PosX                  ' CSPoint(0:第1基準点位置, 1:第2基準点位置)に設定する
                        Else                                            ' ＴＹティーチング時
                            CSPoint(iPos) = stJOG.PosY
                        End If
                    Else
                        If (giAppMode = APP_MODE_TX) Then               ' ＴＸティーチング時
                            CSPoint(iPos) = stJOG.PosY                  ' CSPoint(0:第1基準点位置, 1:第2基準点位置)に設定する
                        Else                                            ' ＴＹティーチング時
                            CSPoint(iPos) = stJOG.PosX
                        End If
                    End If

                    ' ステップオフセット位置更新(第1基準点/第2基準点) ###091
                    If (iPos = 0) Then                                  ' 第1基準点
                        StepOffSetX(0) = Double.Parse(TxtPosX.Text)
                        StepOffSetY(0) = Double.Parse(TxtPosY.Text)
                    Else                                                ' 第2基準点
                        StepOffSetX(1) = Double.Parse(TxtPos2X.Text)
                        StepOffSetY(1) = Double.Parse(TxtPos2Y.Text)
                    End If

                    If (iPos >= 1) Then                                 ' 第1基準点/第2基準点のﾃｨｰﾁﾝｸﾞ終了 ?
                        If (GrpNum <= 1) Then                           ' グループ数 <= 1 ? 
                            r = cFRS_ERR_START                          ' Return値 = OK(START SW)押下 
                        Else
                            r = cFRS_NORMAL                             ' Return値 = グループ間インターバル処理へ 
                        End If
                        Exit Do
                    Else
                        If (stJOG.Flg = -1) Then
                            iPos = 1                                    ' iPos = 第2基準点のﾃｨｰﾁﾝｸﾞ処理
                        End If
                    End If

                    ' HALT SW押下時は１つ前のティーチングへ戻る
                ElseIf (r = cFRS_ERR_HALT) Then                         ' HALT SW押下 ?
                    If (iPos = 0) Then                                  ' 第1基準点のティーチングなら処理続行

                    Else                                                ' 第2基準点なら第1基準点のティーチングへ
                        ' 第2基準点の座標更新 'V3.0.0.0③↓　HALTキーで第１基準点に戻ってからOKをする場合の為に第２基準点の値を格納しておく。
                        If (typPlateInfo.intResistDir = 0) Then             ' ﾁｯﾌﾟ並びはX方向 ?
                            If (giAppMode = APP_MODE_TX) Then               ' ＴＸティーチング時
                                CSPoint(iPos) = stJOG.PosX                  ' CSPoint(0:第1基準点位置, 1:第2基準点位置)に設定する
                            Else                                            ' ＴＹティーチング時
                                CSPoint(iPos) = stJOG.PosY
                            End If
                        Else
                            If (giAppMode = APP_MODE_TX) Then               ' ＴＸティーチング時
                                CSPoint(iPos) = stJOG.PosY                  ' CSPoint(0:第1基準点位置, 1:第2基準点位置)に設定する
                            Else                                            ' ＴＹティーチング時
                                CSPoint(iPos) = stJOG.PosX
                            End If
                        End If
                        StepOffSetX(1) = Double.Parse(TxtPos2X.Text)
                        StepOffSetY(1) = Double.Parse(TxtPos2Y.Text)
                        ' 第2基準点の座標更新 'V3.0.0.0③↑
                        iPos = 0                                        ' iPos = 第1基準点
                    End If

                    '  RESET SW押下時は終了確認メッセージ表示へ
                ElseIf (r = cFRS_ERR_RST) Then                          ' RESET SW押下 ?
                    Exit Do
                End If

                Call ZCONRST()                                          ' ｺﾝｿｰﾙｷｰﾗｯﾁ解除 
            Loop While (stJOG.Flg = -1)

            ' 当画面からOK/Cancelﾎﾞﾀﾝ押下ならrに戻値を設定する
            If (stJOG.Flg <> -1) Then
                r = stJOG.Flg
            End If

            '-------------------------------------------------------------------
            '   チップサイズを算出する
            '-------------------------------------------------------------------
            'V3.0.0.0③            If ((r <> cFRS_ERR_RST) And (iPos >= 1)) Then               ' OK(STARTｷｰ)押下で第2基準点のﾃｨｰﾁﾝｸﾞ終了 ?
            If ((r <> cFRS_ERR_RST) And (Me.cmdOK.Enabled = True)) Then ' OK(STARTｷｰ)押下で第2基準点のﾃｨｰﾁﾝｸﾞ終了 ?
                '  ＴＸティーチング時 → チップサイズ = (第2基準点-第1基準点) / (チップ数-1)) 
                '  ＴＹティーチング時 → チップサイズ = (第2基準点-第1基準点) / (ブロック数-1)) 
                'V2.0.0.0①                WkChipSize = System.Math.Abs(CDbl((CSPoint(1) - CSPoint(0)) / (RnBn - 1)))
                WkChipSize = CDbl((CSPoint(1) - CSPoint(0)) / (RnBn - 1))   'V2.0.0.0①
                strMSG = WkChipSize.ToString("0.0000")
                WkChipSize = Double.Parse(strMSG)
                Call DispChipSize(WkChipSize)                           ' 更新前/更新後のチップサイズを表示する
            End If

            '-------------------------------------------------------------------
            '   終了確認メッセージ表示 
            '-------------------------------------------------------------------
            ' 終了確認メッセージ表示 
            If (r = cFRS_ERR_RST) Then                                  ' Cancel(RESETｷｰ)押下時 ?
                ' "前の画面に戻ります。よろしいですか？　　　
                If (giAppMode = APP_MODE_TX) Then                           ' TX ? 
                    rtn = Form1.System1.TrmMsgBox(gSysPrm, MSG_105, MsgBoxStyle.OkCancel, TITLE_TX)
                Else
                    rtn = Form1.System1.TrmMsgBox(gSysPrm, MSG_105, MsgBoxStyle.OkCancel, TITLE_TY)
                End If
            ElseIf (r = cFRS_ERR_START) Then                            ' OKボタン押下時
                Call Disp_TxTyMsgBox(rtn)                               ' 終了確認メッセージ表示 
            Else                                                        ' START SW押下で第1基準点/第2基準点のﾃｨｰﾁﾝｸﾞ終了時は
                GoTo STP_END                                            ' 終了確認メッセージ表示は表示せずチップサイズ更新後処理継続へ
            End If

            ' Cancel(RESETｷｰ)押下時は処理を継続する
            ' Cancel(RESETｷｰ)押下時は処理を継続する
            If (rtn = cFRS_ERR_RST) Then
                Me.LblResult_0.Text = ""                                ' 補正比率
                Me.LblResult_1.Text = ""                                ' チップサイズ
                Me.LblResult_2.Text = ""                                ' チップサイズ(補正後)
                GoTo STP_RETRY                                          ' 処理継続へ
            End If

            ' TX2(Teach)またはTY2押下
            If (rtn = cFRS_TxTy) Then
                r = rtn
            End If

STP_END:
            ' チップサイズ更新
            If (r <> cFRS_ERR_RST) Then                                 ' Cancel(RESETｷｰ)押下以外 ?
                ChipSize = WkChipSize                                   ' チップサイズ更新
            End If
            Return (r)                                                  ' Return値設定

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTyTeach.Sub_Jog1() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Return値 = 例外エラー
        End Try
    End Function
#End Region

#Region "ＢＰグループ(サーキット)間隔(TX時)またはステージグループ間隔(TY時)のティーチング処理"
    '''=========================================================================
    ''' <summary>ＢＰグループ(サーキット)間隔(TX時)またはステージグループ間隔(TY時)のティーチング処理</summary>
    ''' <param name="RnBn">    (INP)チップ数(抵抗数)(TX時), ブロック数(TY時)</param>
    ''' <param name="ChipSize">(INP)チップサイズ</param>
    ''' <param name="KJPosX">  (INP)先頭ブロックの先頭基準位置X</param>
    ''' <param name="KJPosY">  (INP)先頭ブロックの先頭基準位置Y</param>
    ''' <param name="GrpNum">  (INP)グループ数</param>
    ''' <param name="GrpCnt">  (I/O)グループ数カウンタ</param>
    ''' <param name="CSPoint"> (OUT)チップサイズ算出用(0:第1基準点, 1:第2基準点)</param>
    ''' <returns>cFRS_ERR_START = OK(STARTｷｰ)押下
    '''          cFRS_ERR_RST   = Cancel(RESETｷｰ)押下
    '''          cFRS_ERR_HALT  = HALTｷｰ押下(チップサイズティーチング処理へ)
    '''          cFRS_NORMAL    = 正常(ステップオフセット量の算出へ)
    '''          -1以下         = エラー</returns>
    ''' <remarks>インターバル値をmdStepInterval()に設定</remarks>
    '''=========================================================================
    Private Function Sub_Jog2(ByVal RnBn As Short, ByVal ChipSize As Double, ByVal KJPosX As Double, ByVal KJPosY As Double, _
                              ByVal GrpNum As Short, ByRef GrpCnt As Short, ByRef CSPoint() As Double) As Short

        Dim PosX As Double                                              ' 現在のBP位置X/ﾃｰﾌﾞﾙ移動座標X(絶対値)
        Dim PosY As Double                                              ' 現在のBP位置Y/ﾃｰﾌﾞﾙ移動座標Y(絶対値)
        Dim iPos As Short
        Dim iFlg As Short                                               ' iFlg:2=第nｸﾞﾙｰﾌﾟ,最終基準位置のﾃｨｰﾁﾝｸﾞ, 3=第n+1ｸﾞﾙｰﾌﾟ,先頭基準位置のﾃｨｰﾁﾝｸﾞ 
        Dim r As Short
        Dim rtn As Short                                                ' 戻値 
        Dim strMSG As String

        Try
            ' グループ数｢1｣以下ならNOP
            If (GrpNum <= 1) Then
                Return (cFRS_ERR_START)
            End If

            ' 初期処理
            GrpCnt = 1                                                  ' グループ数カウンタ初期化
            iFlg = 2                                                    ' iFlg = 第nｸﾞﾙｰﾌﾟ,最終基準位置のﾃｨｰﾁﾝｸﾞ
            Me.GrpFram_1.Visible = True                                 ' 第二基準点表示(Sub_Jog3()で非表示にするため)
            stJOG.Flg = -1                                              ' 親画面のOK/Cancelﾎﾞﾀﾝ押下ﾌﾗｸﾞ
            Me.TxtPosX.Text = ""                                        ' 座標表示(ｸﾘｱ)第1基準点XY
            Me.TxtPosY.Text = ""
            Me.TxtPos2X.Text = ""                                       ' 座標表示(ｸﾘｱ)第2基準点XY
            Me.TxtPos2Y.Text = ""
            Call ClearGridView(RowAry)                                  ' データグリッドビューのインターバル表示をクリアする

            If (giAppMode = APP_MODE_TX) Then                           ' ＴＸティーチング ? 
                GrpNum = 2                                              ' BPグループ数(サーキット数)は２として処理する 

            Else                                                        ' ＴＹティーチング時
                GrpNum = 2                                              ' ステージグループ数は２として処理する
                If (typPlateInfo.intResistDir = 0) Then                 ' チップ並びはX方向 ?
                    RnBn = typPlateInfo.intBlkCntInStgGrpY              ' ブロック数 = ステージグループ内ブロック数Y 
                Else
                    RnBn = typPlateInfo.intBlkCntInStgGrpX              ' ブロック数 = ステージグループ内ブロック数X
                End If
            End If

            ' タイトルメッセージの設定 
            If (giAppMode = APP_MODE_TX) Then                           ' ＴＸティーチング時
                If (gTkyKnd = KND_CHIP) Then
                    Me.lblTitle2.Text = INFO_MSG23                      ' タイトル = "ＢＰグループ間隔ティーチング"
                Else
                    Me.lblTitle2.Text = INFO_MSG30                      ' タイトル = "サーキット間隔ティーチング"
                End If
            Else                                                        ' ＴＹティーチング時
                Me.lblTitle2.Text = INFO_MSG14                          ' タイトル = "ステージグループ間隔ティーチング"
            End If

            ' グループ数分またはステージグループ数分のインターバルをティーチング
            Timer1.Enabled = True                                       ' ###228
            Do
                'グループ数(サーキット数)分またはステップ数分の処理完了？
                If (GrpCnt >= GrpNum) Then                              ' インターバルティーチング終了 ?
                    ' ＴＹティーチング時でプレートデータのステップオフセット量X,Y指定なしの場合はステップオフセット量の算出処理へ
                    If ((giAppMode = APP_MODE_TY) And (typPlateInfo.dblStepOffsetXDir = 0) And (typPlateInfo.dblStepOffsetYDir = 0)) Then
                        r = cFRS_NORMAL                                 ' Return値 = ステップオフセット量の算出処理へ 
                    Else
                        r = cFRS_ERR_START                              ' Return値 = OK(START SW)押下 
                    End If
                    Exit Do
                End If

STP_RETRY:
                ' 第(GrpCnt)ｸﾞﾙｰﾌﾟ,最終基準点のﾃｨｰﾁﾝｸﾞ
                If (iFlg = 2) Then                                      ' iFlg(2=第nｸﾞﾙｰﾌﾟ,最終基準位置のﾃｨｰﾁﾝｸﾞ)?
                    '"第" n "グループ、最終端位置のティーチング"+"基準位置を合わせて下さい。"+"移動:[矢印]  決定:[START]  中断:[RESET]" "
                    lblInfo.Text = INFO_MSG19 & GrpCnt & INFO_MSG28 & vbCrLf & INFO_MSG16 & vbCrLf & INFO_MSG17
                    stJOG.TextX = TxtPosX                               ' 第1基準点座標X位置表示用ﾃｷｽﾄﾎﾞｯｸｽ
                    stJOG.TextY = TxtPosY                               ' 第1基準点座標Y位置表示用ﾃｷｽﾄﾎﾞｯｸｽ
                    iPos = 2 * GrpCnt
                Else
                    '"第" n+1 "グループ、最先端位置のティーチング"+"基準位置を合わせて下さい。"+"移動:[矢印]  決定:[START]  中断:[RESET]" "
                    lblInfo.Text = INFO_MSG19 & GrpCnt + 1 & INFO_MSG29 & vbCrLf & INFO_MSG16 & vbCrLf & INFO_MSG17
                    stJOG.TextX = TxtPos2X                              ' 第2基準点座標X位置表示用ﾃｷｽﾄﾎﾞｯｸｽ
                    stJOG.TextY = TxtPos2Y                              ' 第2基準点座標Y位置表示用ﾃｷｽﾄﾎﾞｯｸｽ
                    iPos = (2 * GrpCnt) + 1
                End If

                ' BP移動またはXYテーブル移動(第[GrpCnt]ｸﾞﾙｰﾌﾟ,最終基準位置/第[GrpCnt+1]ｸﾞﾙｰﾌﾟ,先頭基準位置)
                If (giAppMode = APP_MODE_TX) Then                       ' ＴＸティーチング時
                    ' BPを第[GrpCnt]ｸﾞﾙｰﾌﾟの最終基準位置または第[GrpCnt+1]の先頭基準位置に移動する
                    r = XYBPMoveSetBlock(iFlg, iPos, GrpCnt, RnBn, ChipSize, KJPosX, KJPosY, PosX, PosY)
                Else                                                    ' ＴＹティーチング時
                    ' XYテーブルを第[GrpCnt]ｸﾞﾙｰﾌﾟの最終基準位置または第[GrpCnt+1]の先頭基準位置に移動する
                    r = XYTableMoveSetBlock(iFlg, iPos, GrpCnt, RnBn, ChipSize, KJPosX, KJPosY, PosX, PosY)
                End If
                If (r <> cFRS_NORMAL) Then
                    Return (r)                                          ' エラーリターン
                End If

                ' ティーチング処理
                stJOG.PosX = PosX                                       ' BP XまたはXYﾃｰﾌﾞﾙ X絶対位置
                stJOG.PosY = PosY                                       ' BP YまたはXYﾃｰﾌﾞﾙ X絶対位置
                stJOG.Flg = -1                                          ' 親画面のOK/Cancelﾎﾞﾀﾝ押下ﾌﾗｸﾞ
                r = JogEzMove(stJOG, gSysPrm, TBarLowPitch, TBarHiPitch, TBarPause, LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval, Me) 'V2.2.0.0①    引数Me追加 
                If (r < cFRS_NORMAL) Then                               ' エラーなら終了
                    Return (r)
                End If

                ' 最終基準位置/先頭基準位置の絶対値座標更新
                ZRPosX(iPos) = ZRPosX(iPos) + stJOG.cgX
                ZRPosY(iPos) = ZRPosY(iPos) + stJOG.cgY
                ' V1.16.0.1① ADD START-----------------------------------------------------
                If (iPos = (2 * GrpCnt)) Then                           ' 第1基準点 ? 
                    If (typPlateInfo.intResistDir = 0) Then             ' ﾁｯﾌﾟ並びはX方向 ?
                        If (giAppMode = APP_MODE_TX) Then               ' ＴＸティーチング時
                            ZRPosY((2 * GrpCnt) + 1) = ZRPosY(iPos)
                        End If
                    Else
                        If (giAppMode = APP_MODE_TX) Then               ' ＴＸティーチング時
                            ZRPosX((2 * GrpCnt) + 1) = ZRPosX(iPos)
                        End If
                    End If
                End If
                ' V1.16.0.1① ADD END  -----------------------------------------------------

                ' コンソールキーチェック
                If (r = cFRS_ERR_START) Then                            ' START SW押下 ?
                    ' インターバル値算出
                    If (iFlg = 3) Then                                  ' 第n+1ｸﾞﾙｰﾌﾟ,先頭基準位置のﾃｨｰﾁﾝｸﾞ ?
                        ' 第n+1ｸﾞﾙｰﾌﾟ,先頭基準位置の座標更新
                        If (typPlateInfo.intResistDir = 0) Then         ' ﾁｯﾌﾟ並びはX方向 ?
                            If (giAppMode = APP_MODE_TX) Then           ' ＴＸティーチング時
                                CSPoint(1) = stJOG.PosX                 ' CSPoint(1:第n+1ｸﾞﾙｰﾌﾟ,先頭基準位置)に設定する
                            Else                                        ' ＴＹティーチング時
                                CSPoint(1) = stJOG.PosY
                            End If
                        Else
                            If (giAppMode = APP_MODE_TX) Then           ' ＴＸティーチング時
                                CSPoint(1) = stJOG.PosY
                            Else                                        ' ＴＹティーチング時
                                CSPoint(1) = stJOG.PosX                 ' CSPoint(1:第n+1ｸﾞﾙｰﾌﾟ,先頭基準位置)に設定する
                            End If
                        End If
                        ' インターバル値を算出 (ｲﾝﾀｰﾊﾞﾙ値 = 第2基準点-第1基準点)
                        '                        mdStepInterval(GrpCnt) = CDbl(CSPoint(1) - CSPoint(0))
                        mdStepInterval(1) = CDbl(CSPoint(1) - CSPoint(0))
                        iFlg = 2                                        ' iFlg = 第nｸﾞﾙｰﾌﾟ、最終基準位置
                        GrpCnt = GrpCnt + 1                             ' グループﾟ数カウンタ更新
                        If (GrpCnt < GrpNum) Then                       ' 最終グループでない ?
                            ' 座標表示(ｸﾘｱ)
                            Me.TxtPosX.Text = ""                        ' 第1基準点XY
                            Me.TxtPosY.Text = ""
                            Me.TxtPos2X.Text = ""                       ' 第2基準点XY
                            Me.TxtPos2Y.Text = ""
                        End If
                        Exit Do                                         '   ###121

                    Else
                        ' 第nｸﾞﾙｰﾌﾟ,最終基準位置の座標更新
                        If (typPlateInfo.intResistDir = 0) Then         ' ﾁｯﾌﾟ並びはX方向 ?
                            If (giAppMode = APP_MODE_TX) Then           ' ＴＸティーチング時
                                CSPoint(0) = stJOG.PosX                 ' CSPoint(0:第nｸﾞﾙｰﾌﾟ,最終基準位置)に設定する
                            Else                                        ' ＴＹティーチング時
                                CSPoint(0) = stJOG.PosY
                            End If
                        Else
                            If (giAppMode = APP_MODE_TX) Then           ' ＴＸティーチング時
                                CSPoint(0) = stJOG.PosY                 ' CSPoint(0:第nｸﾞﾙｰﾌﾟ,最終基準位置)に設定する
                            Else                                        ' ＴＹティーチング時
                                CSPoint(0) = stJOG.PosX
                            End If
                        End If
                        iFlg = 3                                        ' 第n+1ｸﾞﾙｰﾌﾟ,先頭基準位置
                    End If

                    ' HALT SW押下時は１つ前のティーチングへ戻る
                ElseIf (r = cFRS_ERR_HALT) Then                         ' HALT SW押下 ?
                    If (iFlg = 2) Then                                  ' 第nｸﾞﾙｰﾌﾟ最終基準位置のティーチング ?
                        If (GrpCnt <= 1) Then                           ' 最初のグループなら
                            Return (r)                                  ' チップサイズティーチング処理へ戻る
                        Else
                            GrpCnt = GrpCnt - 1                         ' グループカウンタ -= 1 
                            iFlg = 3                                    ' iFlg = 第n+1ｸﾞﾙｰﾌﾟ,先頭基準位置のﾃｨｰﾁﾝｸﾞへ
                        End If
                    Else                                                ' 第n+1ｸﾞﾙｰﾌﾟ先頭基準位置のティーチング
                        If (GrpCnt <= 1) Then
                            iFlg = 2                                    ' iFlg = 第nｸﾞﾙｰﾌﾟ,最終基準位置のﾃｨｰﾁﾝｸﾞへ
                        End If
                    End If

                    '  RESET SW押下時は終了確認メッセージ表示へ
                ElseIf (r = cFRS_ERR_RST) Then                          ' RESET SW押下 ?
                    Exit Do
                End If
            Loop While (stJOG.Flg = -1)

            ' 当画面からOK/Cancelﾎﾞﾀﾝ押下ならrに戻値を設定する
            If (stJOG.Flg <> -1) Then
                r = stJOG.Flg
            End If

            '-------------------------------------------------------------------
            '   グループ間インターバル値(補正前/補正後)を表示する
            '-------------------------------------------------------------------
            If (r <> cFRS_ERR_RST) Then                                 ' OK(STARTｷｰ)押下で第2基準点のﾃｨｰﾁﾝｸﾞ終了 ?
                Call DispGridView(RowAry, ChipSize)                     ' グループ間インターバル値(補正前/補正後)を表示する
            End If

            '-------------------------------------------------------------------
            '   終了確認メッセージ表示 
            '-------------------------------------------------------------------
            ' 終了確認メッセージ表示 
            If (r = cFRS_ERR_RST) Then                                  ' Cancel(RESETｷｰ)押下時 ?
                ' "前の画面に戻ります。よろしいですか？　　　
                If (giAppMode = APP_MODE_TX) Then                           ' TX ? 
                    rtn = Form1.System1.TrmMsgBox(gSysPrm, MSG_105, MsgBoxStyle.OkCancel, TITLE_TX)
                Else
                    rtn = Form1.System1.TrmMsgBox(gSysPrm, MSG_105, MsgBoxStyle.OkCancel, TITLE_TY)
                End If
            ElseIf (r = cFRS_ERR_START) Then                            ' OKボタン押下時
                Call Disp_TxTyMsgBox(rtn)                               ' 終了確認メッセージ表示 
            Else                                                        ' START SW押下で第1基準点/第2基準点のﾃｨｰﾁﾝｸﾞ終了時は
                Return (r)                                              ' 終了確認メッセージ表示は表示せず処理継続へ
            End If

            ' Cancel(RESETｷｰ)押下時は処理を継続する                     ' ###080↓
            If (rtn = cFRS_ERR_RST) Then

                'If (GrpCnt >= GrpNum) Then                              ' インターバルティーチング終了 ?
                '    GrpCnt = GrpCnt - 1                                 ' 最終グループ番号にする 
                '    iFlg = 3                                            ' 第n+1ｸﾞﾙｰﾌﾟ,先頭基準位置
                'End If

                If (GrpCnt < GrpNum) Then                               ' ###080 インターバルティーチング終了 ?
                    iFlg = 2                                            ' 第n+1ｸﾞﾙｰﾌﾟ,先頭基準位置
                Else
                    GrpCnt = GrpCnt - 1                                 ' 最終グループ番号にする 
                    iFlg = 3
                End If

                GoTo STP_RETRY                                          ' 処理継続へ
            End If                                                      ' ###080↑

            ' TX2(Teach)またはTY2押下
            If (rtn = cFRS_TxTy) Then
                r = rtn
            End If

            Return (r)                                                  ' Return値設定

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTyTeach.Sub_Jog2() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Return値 = 例外エラー
        End Try
    End Function
#End Region

#Region "ステップオフセット量の算出(ＴＹティーチング時)"
    '''=========================================================================
    ''' <summary>ステップオフセット量の算出(ＴＹティーチング時)</summary>
    ''' <param name="RnBn">     (INP)ブロック数</param>
    ''' <param name="ChipSize"> (INP)チップサイズ</param>
    ''' <param name="KJPosX">   (INP)先頭ブロックの先頭基準位置X</param>
    ''' <param name="KJPosY">   (INP)先頭ブロックの先頭基準位置Y</param>
    ''' <param name="GrpNum">   (INP)ｸﾞﾙｰﾌﾟ数</param>
    ''' <param name="GrpCnt">   (INP)ｸﾞﾙｰﾌﾟ数ｶｳﾝﾀ</param>
    ''' <param name="CSPoint">  (OUT)ﾁｯﾌﾟｻｲｽﾞ算出用(0:第1基準点, 1:第2基準点)</param>
    ''' <param name="dStepOffx">(INP)ｽﾃｯﾌﾟｵﾌｾｯﾄ量X</param>
    ''' <param name="dStepOffy">(INP)ｽﾃｯﾌﾟｵﾌｾｯﾄ量Y</param>
    ''' <returns>cFRS_ERR_ADV  = OK(STARTｷｰ)
    '''          cFRS_ERR_RST  = Cancel(RESETｷｰ)
    '''          cFRS_ERR_HALT = HALTｷｰ押下(ステージグループ間隔のティーチング処理へ)
    '''          -1以下        = エラー</returns>
    ''' <remarks>・TYティーチング最終ブロックティーチング時のX方向のずれ量がステップオフセットとなる
    '''          ・プレートデータのステップオフセット量XYが｢0｣の時のみ処理を行う
    ''' </remarks>
    '''=========================================================================
    Private Function Sub_Jog3(ByVal RnBn As Short, ByVal ChipSize As Double, ByVal KJPosX As Double, ByVal KJPosY As Double, _
                              ByVal GrpNum As Short, ByVal GrpCnt As Short, ByRef CSPoint() As Double, ByVal dStepOffx As Double, ByVal dStepOffy As Double) As Short

        Dim PosX As Double                                              ' 現在のBP位置X/ﾃｰﾌﾞﾙ移動座標X(絶対値)
        Dim PosY As Double                                              ' 現在のBP位置Y/ﾃｰﾌﾞﾙ移動座標Y(絶対値)
        Dim iPos As Short
        Dim iFlg As Short                                               ' iFlg:2=第nｸﾞﾙｰﾌﾟ,最終基準位置のﾃｨｰﾁﾝｸﾞ, 3=第n+1ｸﾞﾙｰﾌﾟ,先頭基準位置のﾃｨｰﾁﾝｸﾞ 
        Dim r As Short
        Dim rtn As Short                                                ' 戻値 
        Dim strMSG As String

        Try
            ' ＴＹティーチングでなければNOP
            If (giAppMode <> APP_MODE_TY) Then
                Return (cFRS_ERR_START)
            End If

            ' ｸﾞﾙｰﾌﾟ数｢1｣以下ならNOP
            If (GrpNum <= 1) Then
                Return (cFRS_ERR_START)
            End If

            ' プレートデータのステップオフセット量XYが｢0｣なら処理する
            If (dStepOffx <> 0) And (dStepOffy <> 0) Then
                Return (cFRS_ERR_START)
            End If

            ' 最終ｸﾞﾙｰﾌﾟ最終基準点のﾃｨｰﾁﾝｸﾞ
            Me.lblTitle2.Text = INFO_MSG15                              '"ステップオフセット量　ティーチング"
            iFlg = 4                                                    ' iFlg = 第nｸﾞﾙｰﾌﾟ,最終ブロック基準位置のﾃｨｰﾁﾝｸﾞ
            iPos = 2 * GrpCnt
            ' XYテーブル移動(最終ｸﾞﾙｰﾌﾟ、最終基準位置)
            r = XYTableMoveSetBlock(iFlg, iPos, GrpCnt, RnBn, ChipSize, KJPosX, KJPosY, PosX, PosY)
            ' 座標表示
            Me.TxtPosX.Text = ""                                        ' 第1基準点XY
            Me.TxtPosY.Text = ""

            Me.GrpFram_1.Visible = False                                ' 第二基準点非表示 
            '"ステップオフセット位置のティーチング"+"基準位置を合わせて下さい。"+"移動:[矢印]  決定:[START]  中断:[RESET]"
            lblInfo.Text = INFO_MSG31 & vbCrLf & INFO_MSG16 & vbCrLf & INFO_MSG17

STP_RETRY:
            ' ティーチング処理
            Do
                stJOG.PosX = PosX                                       ' XYﾃｰﾌﾞﾙ X絶対位置
                stJOG.PosY = PosY                                       ' XYﾃｰﾌﾞﾙ Y絶対位置
                stJOG.TextX = TxtPosX                                   ' XYﾃｰﾌﾞﾙ X位置表示用ﾃｷｽﾄﾎﾞｯｸｽ
                stJOG.TextY = TxtPosY                                   ' XYﾃｰﾌﾞﾙ Y位置表示用ﾃｷｽﾄﾎﾞｯｸｽ
                stJOG.Flg = -1                                          ' 親画面のOK/Cancelﾎﾞﾀﾝ押下ﾌﾗｸﾞ
                r = JogEzMove(stJOG, gSysPrm, TBarLowPitch, TBarHiPitch, TBarPause, LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval, Me)   'V2.2.0.0① 引数Me追加
                If (r < cFRS_NORMAL) Then                               ' エラーなら終了
                    Return (r)
                End If

                ' 最終グループ、最終抵抗位置の絶対値座標更新
                ZRPosX(iPos) = ZRPosX(iPos) + stJOG.cgX
                ZRPosY(iPos) = ZRPosY(iPos) + stJOG.cgY

                ' コンソールキーチェック
                If (r = cFRS_ERR_START) Then                            ' START SW押下 ?
                    ' 最終グループ、最終抵抗位置の座標更新
                    If (typPlateInfo.intResistDir = 0) Then             ' ﾁｯﾌﾟ並びはX方向 ?
                        CSPoint(1) = stJOG.PosY
                    Else
                        CSPoint(1) = stJOG.PosX
                    End If

                    ' ステップオフセット位置更新(第2基準点) ###091
                    StepOffSetX(1) = Double.Parse(TxtPosX.Text)
                    StepOffSetY(1) = Double.Parse(TxtPosY.Text)

                    Exit Do

                ElseIf (r = cFRS_ERR_HALT) Then                         ' HALT SW押下 ?
                    Return (r)                                          ' ステージグループ間隔のティーチング処理へ戻る

                ElseIf (r = cFRS_ERR_RST) Then                          ' RESET SW押下 ?
                    Exit Do
                End If

            Loop While (stJOG.Flg = -1)

            ' 当画面からOK/Cancelﾎﾞﾀﾝ押下ならrに戻値を設定する
            If (stJOG.Flg <> -1) Then
                r = stJOG.Flg
            End If

            '-------------------------------------------------------------------
            '   終了確認メッセージ表示 
            '-------------------------------------------------------------------
            ' 終了確認メッセージ表示 
            If (r = cFRS_ERR_RST) Then                                  ' Cancel(RESETｷｰ)押下時 ?
                ' "前の画面に戻ります。よろしいですか？　　　
                If (giAppMode = APP_MODE_TX) Then                           ' TX ? 
                    rtn = Form1.System1.TrmMsgBox(gSysPrm, MSG_105, MsgBoxStyle.OkCancel, TITLE_TX)
                Else
                    rtn = Form1.System1.TrmMsgBox(gSysPrm, MSG_105, MsgBoxStyle.OkCancel, TITLE_TY)
                End If
            ElseIf (r = cFRS_ERR_START) Then                            ' OKボタン押下時
                Call Disp_TxTyMsgBox(rtn)                               ' 終了確認メッセージ表示 
            Else                                                        ' START SW押下で第1基準点/第2基準点のﾃｨｰﾁﾝｸﾞ終了時は
                Return (r)                                              ' 終了確認メッセージ表示は表示せず処理継続へ
            End If

            ' Cancel(RESETｷｰ)押下時は処理を継続する
            If (rtn = cFRS_ERR_RST) Then
                If (GrpCnt >= GrpNum) Then                              ' インターバルティーチング終了 ?
                    GrpCnt = GrpCnt - 1                                 ' 最終グループ番号にする 
                    iFlg = 3                                            ' 第n+1ｸﾞﾙｰﾌﾟ,先頭基準位置
                End If
                GoTo STP_RETRY                                          ' 処理継続へ
            End If

            ' TX2(Teach)またはTY2押下
            If (rtn = cFRS_TxTy) Then
                r = rtn
            End If

            Return (r)                                                  ' Return値設定

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTyTeach.Sub_Jog3() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                      ' Return値 = 例外エラー
        End Try
    End Function
#End Region

    '=========================================================================
    '   トリミングパラメータの取得/更新処理
    '=========================================================================
#Region "トリミングデータより必要なパラメータを取得する(ＴＸ/ＴＹティーチング)"
    '''=========================================================================
    '''<summary>トリミングデータより必要なパラメータを取得する</summary>
    '''<remarks>各インターバルデータを取得する</remarks>
    '''=========================================================================
    Private Sub SetTrimData()

        Dim mdCSx As Double                                                     ' チップサイズX
        Dim mdCSy As Double                                                     ' チップサイズY
        Dim i As Short
        Dim strMSG As String

        Try
            '---------------------------------------------------------------
            '   チップサイズX,Yを取得する
            '---------------------------------------------------------------
            mdCSx = typPlateInfo.dblChipSizeXDir                                ' チップサイズX,Y(CHIP/NET共通)
            mdCSy = typPlateInfo.dblChipSizeYDir
            mdStepInterval(0) = 0.0#

            '---------------------------------------------------------------
            '   各インターバルデータを取得する(ＴＸティーチング時)
            '---------------------------------------------------------------
            ' プレートデータのBPグループ間隔(サーキット間隔)を取得する
            If (giAppMode = APP_MODE_TX) Then                                   ' ＴＹティーチング ?
                For i = 1 To MaxCntStep
                    miStepBlock(i) = typPlateInfo.intResistCntInGroup           ' 1グループ(1サーキット)内抵抗数
                    mdStepInterval(i) = typPlateInfo.dblBpGrpItv                ' BPグループ(サーキット)間隔
                Next i
            End If

            '---------------------------------------------------------------
            '   各インターバルデータを取得する(ＴＹティーチング時)
            '---------------------------------------------------------------
            ' プレートデータからブロック数とインターバル値を取得する(ステップデータ構造体は未使用)
            If (giAppMode = APP_MODE_TY) Then                                   ' ＴＹティーチング ?
                For i = 1 To MaxCntStep
                    If (typPlateInfo.intResistDir = 0) Then                     ' チップ並び方向 = X方向 ?
                        miStepBlock(i) = typPlateInfo.intBlockCntYDir           ' ブロック数Y
                        mdStepInterval(i) = typPlateInfo.dblStgGrpItvY          ' Y方向ステージグループ間隔
                    Else
                        miStepBlock(i) = typPlateInfo.intBlockCntXDir           ' ブロック数X
                        mdStepInterval(i) = typPlateInfo.dblStgGrpItvX          ' X方向ステージグループ間隔
                    End If
                Next i
            End If

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTeach.SetTrimData() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "トリミングパラメータ更新(ＴＸ/ＴＹティーチング)"
    '''=========================================================================
    '''<summary>トリミングパラメータ更新</summary>
    ''' <param name="ChipSize">(INP)チップサイズ</param>
    ''' <param name="GrpCnt">  (INP)ｸﾞﾙｰﾌﾟ数ｶｳﾝﾀ</param>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub SetTrimParameter(ByVal ChipSize As Double, ByVal GrpCnt As Short)

        Dim OffSet As Double = 0.0                                      ' ###091
        Dim strMSG As String

        Try
            '---------------------------------------------------------------
            '   トリミングパラメータ更新(ＴＹティーチング時)
            '---------------------------------------------------------------
            If (giAppMode = APP_MODE_TY) Then
                Call SetTrimParamToGlobalArea_TY(ChipSize, GrpCnt)
                Return
            End If

            '---------------------------------------------------------------
            '   トリミングパラメータ更新(ＴＸティーチング時)
            '---------------------------------------------------------------
            ' チップサイズを更新する
            If (typPlateInfo.intResistDir = 0) Then                     ' チップ並び方向 = X方向 ?
                typPlateInfo.dblChipSizeXDir = ChipSize
            Else
                typPlateInfo.dblChipSizeYDir = ChipSize
            End If

            ' インターバルデータを更新する
            If (RowAry Is Nothing = False) Then
                typPlateInfo.dblBpGrpItv = mdStepInterval(1)            ' BPグループ(サーキット)間隔
            End If

            ' ステップオフセット量(BPオフセット量)を求める ###091
            ' (TXティーチング最終ブロックティーチング時のY方向のずれ量がオフセットとなる)
            If (typPlateInfo.intResistDir = 0) Then                     ' チップ並び方向 = X方向 ?
                OffSet = StepOffSetY(1) - StepOffSetY(0)                ' オフセット位置Y(0:第1基準点　1:第2基準点)
            Else
                OffSet = StepOffSetX(1) - StepOffSetX(0)                ' オフセット位置X(0:第1基準点　1:第2基準点)
            End If

            ' トリミングパラメータ(カット位置XY, ブロックサイズ)を更新する
            'Call SetTrimParamToGlobalArea(ChipSize)                    '###090
            Call SetTrimParamToGlobalArea(ChipSize, OffSet)             '###090

            Return

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTyTeach.SetTrimParameter() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "トリミングパラメータ(カット位置XY, ブロックサイズ)更新(ＴＸティーチング)【CHIP/NET用】"
    '''=========================================================================
    ''' <summary>トリミングパラメータ(カット位置XY, ブロックサイズ)(ＴＸティーチング)【CHIP/NET用】</summary>
    ''' <param name="ChipSize">(INP)チップサイズ</param>
    ''' <param name="OffSet">  (INP)ステップオフセット量(BPオフセット量) ###090</param>
    ''' <returns>0=正常, 0以外:エラー</returns>
    '''=========================================================================
    Private Function SetTrimParamToGlobalArea(ByVal ChipSize As Double, ByVal OffSet As Double) As Short
        'Private Function SetTrimParamToGlobalArea(ByVal ChipSize As Double) As Short

        Dim dTeachPos(MaxCntCut) As Double
        Dim dStartPos(MaxCntCut) As Double
        Dim dAddGrpInt As Double
        Dim iChpNum As Short
        Dim iChpCnt As Short
        Dim iCutNum As Short
        Dim iCutCnt As Short
        Dim strMSG As String
        Dim OffSetReg As Double                                         ' ###090
        Dim WkDbl As Double                                             ' ###264
        Dim cin As Integer                                              ' ###090
        Dim r As Integer                                                ' ###090
        Dim bOfsFlg As Boolean = False                                  ' ###090
        Dim DmyX As Double = 0.0                                        ' ###090
        Dim DmyY As Double = 0.0                                        ' ###090
        Dim mPIT As Double = 0.0                                        ' ###090
        Dim iCircuitCnt As Short                                            'V2.0.0.0⑩
        Try
            ' 初期処理
            iChpNum = typPlateInfo.intResistCntInBlock                  ' ブロック内抵抗数(マーキング用抵抗含む)
            r = GetCutPosOffset(OffSet, iChpNum, OffSetReg, cin)        ' BPオフセット量から１抵抗当たりのオフセット量を求める ###090
            If (r = cFRS_NORMAL) Then bOfsFlg = True '                  ' bOfsFlg = BPオフセットの加算を実行する ###090
            mPIT = 0.0                                                  ' ピッチ初期化 ###090

            ' トリミングパラメータ(カット位置XY)を更新する
            For iChpCnt = 1 To iChpNum                                  ' チップ数(抵抗数)分設定する
                iCutNum = typResistorInfoArray(iChpCnt).intCutCount     ' カット数
                dAddGrpInt = AddGrpInterval(iChpCnt)                    ' BPグループ(サーキット)間隔取得
                'V2.0.0.0⑩↓
                If UserSub.IsTrimType3() Then
                    iCircuitCnt = UserSub.GetCircuitNoFromResNo(stPLT, stREG, iChpCnt)
                Else
                    iCircuitCnt = iChpCnt
                End If
                'V2.0.0.0⑩↑

                ' カット数分設定する
                For iCutCnt = 1 To iCutNum

                    ' 第1抵抗の場合
                    'V2.0.0.0⑩                    If iChpCnt = 1 Then
                    If iCircuitCnt = 1 Then             'V2.0.0.0⑩
                        If (typPlateInfo.intResistDir = 0) Then         ' チップ並び方向 = X方向 ?
                            ' ｽﾀｰﾄﾎﾟｲﾝﾄを取得
                            dStartPos(iCutCnt) = typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).dblStartPointX
                            ' ﾃｨｰﾁﾝｸﾞﾎﾟｲﾝﾄを取得
                            dTeachPos(iCutCnt) = typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).dblTeachPointX

                            ' ﾃｨｰﾁﾝｸﾞﾎﾟｲﾝﾄをｽﾀｰﾄﾎﾟｲﾝﾄにｺﾋﾟｰ    
                            If (gSysPrm.stCTM.giTEACH_P = 2) Then       '  ﾃｨｰﾁﾝｸﾞﾎﾟｲﾝﾄをｽﾀｰﾄﾎﾟｲﾝﾄにｺﾋﾟｰ ?
                                dStartPos(iCutCnt) = dTeachPos(iCutCnt) '
                                typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).dblStartPointX = dStartPos(iCutCnt)
                            End If
                        Else                                            ' チップ並び方向 = Y方向の場合 
                            ' ｽﾀｰﾄﾎﾟｲﾝﾄを取得
                            dStartPos(iCutCnt) = typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).dblStartPointY
                            ' ﾃｨｰﾁﾝｸﾞﾎﾟｲﾝﾄを取得
                            dTeachPos(iCutCnt) = typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).dblTeachPointY

                            ' ﾃｨｰﾁﾝｸﾞﾎﾟｲﾝﾄをｽﾀｰﾄﾎﾟｲﾝﾄにｺﾋﾟｰ 
                            If (gSysPrm.stCTM.giTEACH_P = 2) Then       ' ﾃｨｰﾁﾝｸﾞﾎﾟｲﾝﾄをｽﾀｰﾄﾎﾟｲﾝﾄにｺﾋﾟｰ ?
                                dStartPos(iCutCnt) = dTeachPos(iCutCnt) '
                                typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).dblStartPointY = dStartPos(iCutCnt)
                            End If
                        End If

                        ' 第1抵抗以外の場合
                    Else
                        If (typPlateInfo.intResistDir = 0) Then         ' チップ並び方向 = X方向 ?
                            ' ｽﾀｰﾄﾎﾟｲﾝﾄ更新
                            typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).dblStartPointX = dStartPos(iCutCnt) + (ChipSize * (iChpCnt - 1)) + dAddGrpInt

                            ' ﾃｨｰﾁﾝｸﾞﾎﾟｲﾝﾄへ反映
                            If (gSysPrm.stCTM.giTEACH_P = 1) Or (gSysPrm.stCTM.giTEACH_P = 2) Then
                                ' ﾃｨｰﾁﾝｸﾞﾎﾟｲﾝﾄ
                                typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).dblTeachPointX = dTeachPos(iCutCnt) + (ChipSize * (iChpCnt - 1)) + dAddGrpInt
                            End If

                            ' カット位置X,Yにオフセット量を加算する(BP基準コーナーを考慮) ###090
                            If (bOfsFlg = True) Then
                                '----- ###264↓ -----
                                ' 第n抵抗の第nカット位置yにオフセット量を加算する
                                WkDbl = typResistorInfoArray(1).ArrCut(iCutCnt).dblStartPointY
                                Call Form1.Utility1.GetBPmovePitch(cin, DmyX, DmyY, mPIT, typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).dblStartPointX, WkDbl, gSysPrm.stDEV.giBpDirXy)
                                typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).dblStartPointY = WkDbl
                                'Call Form1.Utility1.GetBPmovePitch(cin, DmyX, DmyY, mPIT, typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).dblStartPointX, typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).dblStartPointY, gSysPrm.stDEV.giBpDirXy)
                                '----- ###264↑ -----
                            End If

                            ' ﾎﾟｼﾞｼｮﾆﾝｸﾞなしIXｶｯﾄ,ﾎﾟｼﾞｼｮﾆﾝｸﾞなしSTｶｯﾄは使用できないので初期化
                            Select Case typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).strCutType
                                Case CNS_CUTP_IX2, CNS_CUTP_ST2
                                    typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).dblStartPointX = 0.0
                                    typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).dblTeachPointX = 0.0
                            End Select
                        Else                                            ' チップ並び方向 = Y方向の場合 
                            ' ｽﾀｰﾄﾎﾟｲﾝﾄ
                            'V2.0.0.0⑩↓
                            If UserSub.IsTrimType3() Then
                                Dim OrderNo As Short = UserSub.GetResNumberInCircuit(stREG, iChpCnt)   ' サーキット内の抵抗の順番
                                Dim FromNo As Short = UserSub.GetRNumByCircuit(stPLT, stREG, 1, OrderNo)   ' 第一サーキット内の順番の抵抗番号
                                typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).dblStartPointY = typResistorInfoArray(FromNo).ArrCut(iCutCnt).dblStartPointY + (ChipSize * (iCircuitCnt - 1)) + dAddGrpInt
                            Else
                                'V2.0.0.0⑩↑
                                typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).dblStartPointY = dStartPos(iCutCnt) + (ChipSize * (iChpCnt - 1)) + dAddGrpInt
                            End If                                  'V2.0.0.0⑩

                            ' ﾃｨｰﾁﾝｸﾞﾎﾟｲﾝﾄへ反映
                            If (gSysPrm.stCTM.giTEACH_P = 1) Or (gSysPrm.stCTM.giTEACH_P = 2) Then
                                ' ﾃｨｰﾁﾝｸﾞﾎﾟｲﾝﾄ
                                typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).dblTeachPointY = dTeachPos(iCutCnt) + (ChipSize * (iChpCnt - 1)) + dAddGrpInt
                            End If

                            If (bOfsFlg = True) Then
                                '----- ###264↓ -----
                                ' 第n抵抗の第nカット位置xにオフセット量を加算する
                                WkDbl = typResistorInfoArray(1).ArrCut(iCutCnt).dblStartPointX
                                Call Form1.Utility1.GetBPmovePitch(cin, DmyX, DmyY, mPIT, WkDbl, typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).dblStartPointY, gSysPrm.stDEV.giBpDirXy)
                                typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).dblStartPointX = WkDbl
                                'Call Form1.Utility1.GetBPmovePitch(cin, DmyX, DmyY, mPIT, typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).dblStartPointX, typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).dblStartPointY, gSysPrm.stDEV.giBpDirXy)
                                '----- ###264↑ -----
                            End If

                            ' ﾎﾟｼﾞｼｮﾆﾝｸﾞなしIXｶｯﾄ,ﾎﾟｼﾞｼｮﾆﾝｸﾞなしSTｶｯﾄは使用できないので初期化
                            Select Case typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).strCutType
                                Case CNS_CUTP_IX2, CNS_CUTP_ST2
                                    typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).dblStartPointY = 0.0
                                    typResistorInfoArray(iChpCnt).ArrCut(iCutCnt).dblTeachPointY = 0.0
                            End Select
                        End If

                    End If
                Next iCutCnt                                            ' 次カットへ

                ' ピッチ更新 ###090
                mPIT = mPIT + OffSetReg                                 ' ピッチ = ピッチ + １抵抗当たりのオフセット量 

            Next iChpCnt                                                ' 次抵抗へ

            ' プロックサイズを設定する
            Call CalcBlockSize(typPlateInfo.dblBlockSizeXDir, typPlateInfo.dblBlockSizeYDir)

            Return (cFRS_NORMAL)                                        ' Return値 = 正常

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTyTeach.SetTrimParamToGlobalArea() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Return値 = 例外エラー
        End Try
    End Function
#End Region

#Region "BPオフセット量から１抵抗当たりのオフセット量を求める"
    '''=========================================================================
    ''' <summary>BPオフセット量から１抵抗当たりのオフセット量を求める ###090</summary>
    ''' <param name="OffSet">       (INP)オフセット量</param>
    ''' <param name="registorCnt">  (INP)抵抗数</param>
    ''' <param name="OffSetReg">    (OUT)１抵抗当たりのオフセット量</param>
    ''' <param name="cin">          (OUT)BP移動方向(コンソールの矢印キーコード)</param>
    ''' <returns>cFRS_NORMAL  =  オフセット量を設定した
    '''          cFRS_ERR_RST =  オフセット量を設定しない </returns>
    '''=========================================================================
    Private Function GetCutPosOffset(ByVal OffSet As Double, ByVal registorCnt As Integer, ByRef OffSetReg As Double, ByRef cin As Integer) As Integer

        Dim strMSG As String

        Try
            '------------------------------------------------------------------------
            '   初期処理
            '------------------------------------------------------------------------
            '----- V1.20.0.1①↓ -----
            'If (OffSet = 0.0) Then Return (cFRS_ERR_RST) '              ' オフセット量が0ならNOP 
            ' オフセット量が0の場合もReturn値 = cFRS_NORMAL(オフセット量を設定した)で返す
            If (OffSet = 0.0) Then
                OffSetReg = 0.0                                         ' １抵抗当たりのオフセット量 = 0.0 
                Return (cFRS_NORMAL)                                    ' Return値 = オフセット量を設定した
            End If
            '----- V1.20.0.1①↑ -----

            If (typPlateInfo.intResistDir = 0) Then                     ' チップ並び方向 = X方向の場合
                ' BP移動方向を求める
                Select Case (gSysPrm.stDEV.giBpDirXy)
                    Case 0, 1                                           ' y↓
                        If (OffSet > 0.0) Then
                            cin = &H1000                                ' BP移動方向 =↓(+Y)
                        Else
                            cin = &H800                                 ' BP移動方向 =↑(-Y)
                        End If
                    Case 2, 3                                           ' y↑
                        If (OffSet > 0.0) Then
                            cin = &H800                                 ' BP移動方向 =↑(+Y)
                        Else
                            cin = &H1000                                ' BP移動方向 =↓(-Y)
                        End If
                End Select

            ElseIf (typPlateInfo.intResistDir = 1) Then                 ' チップ並び方向 = Y方向の場合
                ' BP移動方向を求める
                Select Case (gSysPrm.stDEV.giBpDirXy)
                    Case 0, 2                                           ' x←
                        If (OffSet > 0.0) Then
                            'cin = &H200                                 ' BP移動方向 =←(+X)
                            cin = &H400                                 ' BP移動方向 =→(-X) V1.14.0.0③
                        Else
                            'cin = &H400                                 ' BP移動方向 =→(-X)
                            cin = &H200                                 ' BP移動方向 =←(+X) V1.14.0.0③
                        End If
                    Case 1, 4                                           ' x→
                        If (OffSet > 0.0) Then
                            'cin = &H400                                 ' BP移動方向 =→(+X)
                            cin = &H200                                 ' BP移動方向 =←(+X) V1.14.0.0③
                        Else
                            'cin = &H200                                 ' BP移動方向 =←(-X)
                            cin = &H400                                 ' BP移動方向 =→(-X) V1.14.0.0③
                        End If
                End Select
            End If

            ' １抵抗当たりのオフセット量(OffSetReg)を求める
            OffSetReg = Math.Abs(OffSet / (registorCnt - 1))            ' １抵抗当たりのオフセット量 = オフセット量 / 抵抗数 - 1 

            Return (cFRS_NORMAL)                                        ' Return値 = オフセット量を設定した

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "frmTxTyTeach.GetCutPosOffset() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cFRS_ERR_RST)
        End Try
    End Function
#End Region

#Region "トリミングパラメータ更新処理(ＴＹティーチング)【CHIP/NET用】"
    '''=========================================================================
    '''<summary>トリミングパラメータ更新処理(ＴＹティーチング)【CHIP/NET用】</summary>
    ''' <param name="ChipSize">(INP)チップサイズ</param>
    ''' <param name="GrpCnt">  (INP)ｸﾞﾙｰﾌﾟ数ｶｳﾝﾀ</param>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub SetTrimParamToGlobalArea_TY(ByVal ChipSize As Double, ByVal GrpCnt As Short)

        Dim lRow As Integer
        Dim RowMax As Integer
        'Dim iPos As Integer
        'Dim SpOff As Double
        Dim dData As Double
        Dim strMSG As String

        Try
            ' チップサイズを更新する
            If (typPlateInfo.intResistDir = 0) Then                     ' チップ並び方向 = X方向 ?
                typPlateInfo.dblChipSizeYDir = ChipSize
            Else
                typPlateInfo.dblChipSizeXDir = ChipSize
            End If

            ' プロックサイズを設定する ###113
            Call CalcBlockSize(typPlateInfo.dblBlockSizeXDir, typPlateInfo.dblBlockSizeYDir)

            ' ステージグループ間隔を更新する
            If (RowAry Is Nothing = False) Then
                RowMax = RowAry.Length
                If (RowMax >= 1) Then
                    lRow = 1
                    If (typPlateInfo.intResistDir = 0) Then                 ' チップ並び方向 = X方向 ?
                        typPlateInfo.dblStgGrpItvY = mdStepInterval(lRow)   ' Y方向ステージグループ間隔
                    Else
                        typPlateInfo.dblStgGrpItvX = mdStepInterval(lRow)   ' X方向ステージグループ間隔
                    End If
                End If
            End If

            ' ステップオフセット量を設定する
            ' (TYティーチング最終ブロックティーチング時のX方向のずれ量がステップオフセットとなる) ###091
            If (typPlateInfo.intResistDir = 0) Then                     ' チップ並び方向 = X方向 ?
                dData = StepOffSetX(1) - StepOffSetX(0)                 ' オフセット位置Y(0:第1基準点　1:第2基準点)
            Else
                dData = StepOffSetY(1) - StepOffSetY(0)                 ' オフセット位置X(0:第1基準点　1:第2基準点)
            End If
            'dData = 0.0#
            'iPos = 2 * GrpCnt
            'For lRow = 0 To iPos
            '    If (typPlateInfo.intResistDir = 0) Then                 ' チップ並び方向 = X方向 ?
            '        dData = dData + ZRPosX(lRow)                        ' ZRPosX(ずれ量X(0:第1基準点, 1:第2基準点...))
            '    Else
            '        dData = dData + ZRPosY(lRow)                        ' ZRPosY(ずれ量Y(0:第1基準点, 1:第2基準点...))
            '    End If
            'Next lRow

            ' プレートデータのステップオフセット量を更新する(ステップオフセット量が0の場合のみ設定する) 
            'If ((typPlateInfo.dblStepOffsetXDir = 0) And (typPlateInfo.dblStepOffsetYDir = 0)) Then '###263
            If (typPlateInfo.intResistDir = 0) Then                 ' チップ並び方向 = X方向 ?
                ' ﾄﾘﾐﾝｸﾞ位置座標 + ｷｬﾘﾌﾞﾚｰｼｮﾝﾃｰﾌﾞﾙｵﾌｾｯﾄ + 補正位置  (+or-) ﾃｰﾌﾞﾙ補正量
                '###249 ↓
                Select Case (gSysPrm.stDEV.giBpDirXy)
                    Case 0, 2      ' x←, y↓
                        typPlateInfo.dblStepOffsetXDir = CDbl(dData.ToString("0.0000"))
                    Case 1, 3      ' x→, y↓
                        typPlateInfo.dblStepOffsetXDir = -1 * CDbl(dData.ToString("0.0000"))
                End Select
                '###249 ↑                   typPlateInfo.dblStepOffsetXDir = CDbl(dData.ToString("0.0000"))
                'SpOff = (dData - ZRPosX(0)) / (typPlateInfo.intBlockCntYDir - 1)
                'typPlateInfo.dblStepOffsetXDir = CDbl(SpOff.ToString("0.0000"))
            Else
                '###249 ↓
                Select Case (gSysPrm.stDEV.giBpDirXy)
                    Case 0, 1      ' y↓
                        typPlateInfo.dblStepOffsetYDir = CDbl(dData.ToString("0.0000"))
                    Case 2, 3      ' y↑
                        typPlateInfo.dblStepOffsetYDir = -1 * CDbl(dData.ToString("0.0000"))
                End Select
                '###249 ↑                   typPlateInfo.dblStepOffsetYDir = CDbl(dData.ToString("0.0000"))
                'SpOff = (dData - ZRPosY(0)) / (typPlateInfo.intBlockCntXDir - 1)
                'typPlateInfo.dblStepOffsetYDir = CDbl(SpOff.ToString("0.0000"))
            End If
            'End If '###263

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTyTeach.SetTrimParamToGlobalArea_TY() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "指定抵抗に対応するBPグループ(サーキット)間隔取得"
    '''=========================================================================
    '''<summary>指定抵抗に対応するBPグループ(サーキット)間隔取得</summary>
    '''<param name="Rno">(INP)抵抗データインデックス</param>
    '''<returns>BPグループ(サーキット)間隔</returns>
    '''=========================================================================
    Private Function AddGrpInterval(ByRef Rno As Short) As Double

        Dim iGrCnt As Integer
        Dim dRet As Double                                              ' 戻り値
        Dim strMSG As String

        Try
            ' BPグループ(サーキット)数が1なら0を返す 
            If (typPlateInfo.intGroupCntInBlockXBp <= 1) Then
                Return (0.0)
            End If

            ' グループ番号 = 抵抗番号 / 1グループ(1サーキット)内抵抗数
            iGrCnt = Rno \ typPlateInfo.intResistCntInGroup
            If ((Rno Mod typPlateInfo.intResistCntInGroup) = 0) Then
                iGrCnt = iGrCnt - 1                                     ' 余り0ならグループ番号-1 
            End If

            ' インターバル = グループ番号 * BPグループ(サーキット)間隔
            dRet = iGrCnt * typPlateInfo.dblBpGrpItv
            Return (dRet)

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTyTeach.AddGrpInterval() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (0.0)
        End Try
    End Function
#End Region

    '=========================================================================
    '   データグリッドビューの表示処理
    '=========================================================================
#Region "データグリッドビューオブジェクトの生成"
    '''=========================================================================
    '''<summary>データグリッドビューオブジェクトの生成</summary>
    ''' <param name="ObjGrid">(I/O)データグリッドビューオブジェクト</param>
    ''' <param name="GrpNum"> (INP)グループ数</param>
    ''' <param name="RowAry"> (OUT)Row(行)オブジェクト配列オブジェクト</param>
    ''' <param name="ColAry"> (OUT)Col(列)オブジェクト配列オブジェクト</param>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub InitGridView(ByVal ObjGrid As DataGridView, ByVal GrpNum As Short, ByRef RowAry() As DataGridViewRow, ByRef ColAry() As DataGridViewColumn)

        Dim RowIdx As Integer
        Dim ColIdx As Integer
        Dim RowCount As Integer
        Dim ColCount As Integer
        Dim strMSG As String

        Try
            ' グリッドの表示/非表示を設定する
            If (GrpNum <= 1) Then                                   ' ｸﾞﾙｰﾌﾟ数1以下なら非表示
                LblDisp_10.Visible = False                          ' ラベル非表示
                GridView.Visible = False                            ' ﾌﾚｷｼﾌﾞﾙｸﾞﾘｯﾄﾞ非表示
                Exit Sub
            End If

            ' Row(行)/Col(列)オブジェクト配列生成
            RowCount = GrpNum - 1                                   ' 最大行数 = ｸﾞﾙｰﾌﾟ数(0 ORG)
            ColCount = 3                                            ' 最大列数 = 3
            RowAry = New DataGridViewRow(RowCount) {}               ' Row(行)オブジェクト配列生成
            ColAry = New DataGridViewColumn(ColCount) {}            ' Col(列)オブジェクト配列生成

            ' グリッド属性を設定する
            ObjGrid.ReadOnly = True                                 ' 編集禁止
            ObjGrid.ColumnHeadersVisible = True                     ' 列ヘッダー表示
            'ObjGrid.RowHeadersVisible = True                        ' 行ヘッダー表示
            ObjGrid.RowHeadersVisible = False                       ' 行ヘッダー非表示
            ObjGrid.AutoGenerateColumns = False                     ' 列が自動的に作成されないようにする
            ObjGrid.AllowUserToAddRows = False                      ' 新しい行を追加できないようにする 
            ObjGrid.AllowUserToResizeColumns = False                ' 列の幅を変更できないようにする
            ObjGrid.AllowUserToResizeRows = False                   ' 行の高さを変更できないようにする
            '                                                       ' ヘッダーの列の幅/行の高さを変更できないようにする
            ObjGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            ObjGrid.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

            ' 列ヘッダーを設定する
            For ColIdx = 0 To (ColCount - 1)
                ' Col(列)オブジェクトを作成する
                Dim ObjCol As DataGridViewTextBoxColumn = New DataGridViewTextBoxColumn()
                ObjGrid.Columns.Add(ObjCol)                         ' Col(列)を追加する
                ColAry(ColIdx) = ObjCol                             ' Col(列)オブジェクト配列に設定
                Select Case (ColIdx)
                    Case 0                                          ' 列ヘッダー設定(0列 列見出し = なし)
                        ObjGrid.Columns(ColIdx).HeaderText = ""
                    Case 1                                          ' 列ヘッダー設定(1列 列見出し = "補正前")
                        ObjGrid.Columns(ColIdx).HeaderText = LBL_TXTY_TEACH_07
                    Case Else                                       ' 列ヘッダー設定(2列 列見出し = "補正後")
                        ObjGrid.Columns(ColIdx).HeaderText = LBL_TXTY_TEACH_08
                End Select

                ObjGrid.Columns(ColIdx).Name = "Col" + ColIdx.ToString()                        ' 列の名前 
                ObjGrid.Columns(ColIdx).Width = Len(ObjGrid.Columns(ColIdx).HeaderText) + 120   ' 列の幅 
                '                                                                               ' ヘッダーテキストの配置を上下左右とも中央にする 
                ObjGrid.Columns(ColIdx).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                '                                                                               ' 列のテキストの配置を右詰にする 
                ObjGrid.Columns(ColIdx).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Next ColIdx

            ' すべての列の並び替えを禁止する
            For Each Col As DataGridViewColumn In ObjGrid.Columns
                Col.SortMode = DataGridViewColumnSortMode.NotSortable
            Next Col

            ' 行情報を設定する
            For RowIdx = 0 To (RowCount - 1)
                Dim ObjRow As DataGridViewRow = New DataGridViewRow()   ' Row(行)オブジェクト作成
                RowAry(RowIdx) = ObjRow                                 ' Row(行)オブジェクト配列に設定
                RowAry(RowIdx).CreateCells(ObjGrid)                     ' Col(列)情報を取得する

                ' 行(ヘッダーグループ1-2, グループ2-3 .....)をRow(行)オブジェクトに設定する
                strMSG = LBL_TXTY_TEACH_14 + " " + (RowIdx + 1).ToString("0") + "-" + (RowIdx + 2).ToString("0")
                RowAry(RowIdx).Cells(0).Value = strMSG
                ObjGrid.Rows.Add(RowAry(RowIdx))                        ' グリッドのRow(行)プロパティにセット
                ObjGrid.Rows(RowIdx).HeaderCell.Value = ""              ' 行(ヘッダーグループ1-2, グループ2-3 .....)
                '                                                       ' テキストの配置を右詰めにする 
                ObjGrid.Rows(RowIdx).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
            Next RowIdx

            ' セルのすべての選択を解除
            ObjGrid.ClearSelection()

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTeach.InitGridView() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "データグリッドビューオブジェクトの開放"
    '''=========================================================================
    ''' <summary>データグリッドビューオブジェクトの開放</summary>
    ''' <param name="RowAry"> (I/O)Row(行)オブジェクト配列オブジェクト</param>
    ''' <param name="ColAry"> (I/O)Col(列)オブジェクト配列オブジェクト</param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub TermGridView(ByRef RowAry() As DataGridViewRow, ByRef ColAry() As DataGridViewColumn)

        Dim Idx As Integer
        Dim Count As Integer
        Dim strMSG As String

        Try
            ' Row(行)オブジェクト配列開放
            If (RowAry Is Nothing = False) Then
                Count = RowAry.Length
                For Idx = 0 To (Count - 1)
                    If (RowAry(Idx) Is Nothing = False) Then
                        RowAry(Idx).Dispose()                       ' リソース開放 
                    End If
                Next Idx
                RowAry = Nothing
            End If

            ' Col(列)オブジェクト配列開放
            If (ColAry Is Nothing = False) Then
                Count = ColAry.Length
                For Idx = 0 To (Count - 1)
                    If (ColAry(Idx) Is Nothing = False) Then
                        ColAry(Idx).Dispose()                       ' リソース開放 
                    End If
                Next Idx
                ColAry = Nothing
            End If

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTyTeach.TermGridView() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "データグリッドビューのインターバル表示をクリアする"
    '''=========================================================================
    '''<summary>データグリッドビューのインターバル表示をクリアする</summary>
    ''' <param name="RowAry"> (I/O)Row(行)オブジェクト配列オブジェクト</param>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub ClearGridView(ByVal RowAry() As DataGridViewRow)

        Dim RowIdx As Integer
        Dim RowMax As Integer
        Dim strMSG As String

        Try
            If (RowAry Is Nothing = True) Then Exit Sub
            RowMax = RowAry.Length
            For RowIdx = 0 To (RowMax - 1)
                If (RowAry(RowIdx) Is Nothing = False) Then
                    RowAry(RowIdx).Cells(1).Value = ""                  ' 補正前インターバル表示クリア
                    RowAry(RowIdx).Cells(2).Value = ""                  ' 補正後インターバル表示クリア
                End If
            Next RowIdx

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTyTeach.ClearGridView() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "データグリッドビューへインターバル値(補正前/補正後)を表示する"
    '''=========================================================================
    ''' <summary>データグリッドビューへインターバル値を表示する</summary>
    ''' <param name="RowAry">  (I/O)Row(行)オブジェクト配列オブジェクト</param>
    ''' <param name="ChipSize">(INP)チップサイズ</param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub DispGridView(ByVal RowAry() As DataGridViewRow, ByVal ChipSize As Double)

        Dim RowIdx As Integer
        Dim RowMax As Integer
        Dim StgGrpItv As Double
        Dim strMSG As String

        Try
            ' Row(行)オブジェクト配列なしならNOP
            If (RowAry Is Nothing = True) Then Exit Sub

            ' 第2基準点表示なしならNOP
            If (TxtPos2X.Text = "") Then Exit Sub

            ' データグリッドビューへインターバル値を表示する
            RowMax = RowAry.Length - 2
            For RowIdx = 0 To RowMax                                    ' グリッド行数分繰り返す 
                If (giAppMode = APP_MODE_TX) Then                       ' ＴＸティーチング時(CHIP/NET共通)
                    ' 補正前インターバル(ＢＰグループ(サーキット)間隔)表示
                    RowAry(RowIdx).Cells(1).Value = typPlateInfo.dblBpGrpItv.ToString("0.0000")
                    ' 補正後インターバル(ＢＰグループ(サーキット)間隔)表示
                    'RowAry(RowIdx).Cells(2).Value = mdStepInterval(RowIdx + 1).ToString("0.0000")
                    RowAry(RowIdx).Cells(2).Value = mdStepInterval(1).ToString("0.0000")                ' ###081

                Else                                                    ' ＴＹティーチング時
                    ' 補正前インターバル表示
                    If (typPlateInfo.intResistDir = 0) Then             ' チップ並び方向 = X方向 ?
                        StgGrpItv = typPlateInfo.dblStgGrpItvY          ' Y方向ステージグループ間隔
                    Else
                        StgGrpItv = typPlateInfo.dblStgGrpItvX          ' X方向ステージグループ間隔
                    End If
                    RowAry(RowIdx).Cells(1).Value = StgGrpItv.ToString("0.0000")
                    ' 補正後インターバル表示
                    'RowAry(RowIdx).Cells(2).Value = mdStepInterval(RowIdx + 1).ToString("0.0000")
                    RowAry(RowIdx).Cells(2).Value = mdStepInterval(1).ToString("0.0000")                ' ###081
                End If
            Next RowIdx

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTyTeach.DispGridView() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

    '=========================================================================
    '   その他共通関数
    '=========================================================================
#Region "更新前/更新後のチップサイズを表示する"
    '''=========================================================================
    '''<summary>更新前/更新後のチップサイズを表示する</summary>
    ''' <param name="ChipSize">(INP)更新後チップサイズ</param>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub DispChipSize(ByVal ChipSize As Double)

        Dim mdCSx As Double                                         ' 更新前のチップサイズX
        Dim mdCSy As Double                                         ' 更新前のチップサイズY
        Dim strMSG As String

        Try
            ' 更新前のチップサイズを取得する(CHIP/NET共通)
            mdCSx = typPlateInfo.dblChipSizeXDir
            mdCSy = typPlateInfo.dblChipSizeYDir

            ' 更新前/更新後のチップサイズを表示する
            If (giAppMode = APP_MODE_TX) Then
                ' ＴＸティーチング時
                With Me
                    If (typPlateInfo.intResistDir = 0) Then         ' チップ並びはX方向 ?
                        ' チップサイズX(旧)
                        .LblResult_1.Text = mdCSx.ToString("0.0000")
                        ' 補正比率
                        .LblResult_0.Text = System.Math.Abs(ChipSize / mdCSx).ToString("0.0000")
                    Else
                        ' チップサイズY(旧)
                        .LblResult_1.Text = mdCSy.ToString("0.0000")
                        ' 補正比率
                        .LblResult_0.Text = System.Math.Abs(ChipSize / mdCSy).ToString("0.0000")
                    End If
                    ' チップサイズ(新)
                    .LblResult_2.Text = ChipSize.ToString("0.0000")
                End With
            Else
                ' ＴＹティーチング時
                With Me
                    If (typPlateInfo.intResistDir = 0) Then         ' チップ並びはX方向 ?
                        ' チップサイズX(旧)
                        .LblResult_1.Text = mdCSy.ToString("0.0000")
                        ' 補正比率
                        .LblResult_0.Text = System.Math.Abs(ChipSize / mdCSy).ToString("0.0000")
                    Else
                        ' チップサイズY(旧)
                        .LblResult_1.Text = mdCSx.ToString("0.0000")
                        ' 補正比率
                        .LblResult_0.Text = System.Math.Abs(ChipSize / mdCSx).ToString("0.0000")
                    End If
                    ' チップサイズ(新)
                    .LblResult_2.Text = ChipSize.ToString("0.0000")
                End With
            End If

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTyTeach.DispChipSize() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "ブロック移動(BP移動)【TXティーチング用】"
    '''=========================================================================
    ''' <summary>ブロック移動(BP移動)【TXティーチング用】</summary>
    ''' <param name="MODE">    (INP)0=第1ｸﾞﾙｰﾌﾟ第1基準位置 , 1=第1ｸﾞﾙｰﾌﾟ最終基準位置
    '''                             2=第nｸﾞﾙｰﾌﾟ最終基準位置, 3=第n+1ｸﾞﾙｰﾌﾟ、先頭基準位置置</param>
    ''' <param name="iPos">    (INP)ずれ量X,Yﾃｰﾌﾞﾙ(ZRPosX(),ZRPosY())ｲﾝﾃﾞｯｸｽ</param>
    ''' <param name="Ino">     (INP)処理中のｸﾞﾙｰﾌﾟ番号</param>
    ''' <param name="ChipNum"> (INP)チップ数</param>
    ''' <param name="ChipSize">(INP)チップサイズ</param>
    ''' <param name="KJPosX">  (INP)先頭ブロックの先頭基準位置X</param>
    ''' <param name="KJPosY">  (INP)先頭ブロックの先頭基準位置y</param>
    ''' <param name="TBLx">    (OUT)BP位置X</param>
    ''' <param name="TBLy">    (OUT)BP位置y</param>
    ''' <returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Private Function XYBPMoveSetBlock(ByRef MODE As Short, ByRef iPos As Short, ByVal Ino As Short, ByVal ChipNum As Short, _
                                      ByVal ChipSize As Double, ByVal KJPosX As Double, ByVal KJPosY As Double, _
                                      ByRef TBLx As Double, ByRef TBLy As Double) As Short

        'Dim i As Short
        'Dim iBlock As Short
        'Dim dIntv As Double
        'Dim ZZx As Double
        'Dim ZZy As Double
        Dim strMSG As String
        Dim r As Short
        'Dim Gn As Integer
        'Dim GrpSz As Double

        Try
            Select Case MODE
                Case 0                                              ' 第1ｸﾞﾙｰﾌﾟ、第1基準位置
                    ' 第1ﾌﾞﾛｯｸの第一抵抗の第一カット位置 + ずれ量 の位置を求める
                    TBLx = KJPosX + ZRPosX(0)                       ' 第1ﾌﾞﾛｯｸの先頭基準位置X + 第1基準点のずれ量X
                    TBLy = KJPosY + ZRPosY(0)                       ' 第1ﾌﾞﾛｯｸの先頭基準位置Y + 第1基準点のずれ量Y

                Case 1                                              ' 第1ｸﾞﾙｰﾌﾟ、最終基準位置
                    If (typPlateInfo.intResistDir = 0) Then         ' ﾁｯﾌﾟ並び X方向 ?
                        ' ﾁｯﾌﾟｻｲｽﾞ * ﾁｯﾌﾟ数-1 + 第1ﾌﾞﾛｯｸの先頭基準位置X + 第1基準点のずれ量X の位置を求める
                        TBLx = (ChipSize * (ChipNum - 1)) + KJPosX + ZRPosX(1)
                        TBLy = KJPosY + ZRPosY(1)
                    Else
                        TBLx = KJPosX + ZRPosX(1)
                        TBLy = (ChipSize * (ChipNum - 1)) + KJPosY + ZRPosY(1)
                    End If

                Case 2                                              ' 第nｸﾞﾙｰﾌﾟ、最終基準位置
                    If (typPlateInfo.intResistDir = 0) Then         ' ﾁｯﾌﾟ並び X方向 ?
                        ' ﾁｯﾌﾟｻｲｽﾞ * ﾁｯﾌﾟ数 + 第1ﾌﾞﾛｯｸの先頭基準位置X + 第1基準点のずれ量X の位置を求める
                        TBLx = (ChipSize * ChipNum) + KJPosX + ZRPosX(MODE)
                        TBLy = KJPosY + ZRPosY(MODE)
                    Else
                        TBLx = KJPosX + ZRPosX(MODE)
                        TBLy = (ChipSize * ChipNum) + KJPosY + ZRPosX(MODE)
                    End If

                Case 3                                              ' 第n+1ｸﾞﾙｰﾌﾟ、先頭基準位置
                    If (typPlateInfo.intResistDir = 0) Then         ' ﾁｯﾌﾟ並び X方向 ?
                        ' ﾁｯﾌﾟｻｲｽﾞ * ﾁｯﾌﾟ数 + 第1ﾌﾞﾛｯｸの先頭基準位置X + 第1基準点のずれ量X の位置を求める
                        TBLx = (ChipSize * ChipNum) + KJPosX + ZRPosX(MODE) + mdStepInterval(MODE)
                        TBLy = KJPosY + ZRPosY(MODE)
                    Else
                        TBLx = KJPosX + ZRPosX(MODE)
                        TBLy = (ChipSize * ChipNum) + KJPosY + ZRPosX(MODE) + mdStepInterval(MODE)
                    End If

            End Select

            ' BP移動(絶対値指定)
            r = Form1.System1.EX_MOVE(gSysPrm, TBLx, TBLy, 1)
            Return (r)                                      ' Return

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTeach.XYBPMoveSetBlock() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                              ' Return値 = 例外エラー 
        End Try

    End Function
#End Region

#Region "先頭ブロックの先頭基準位置を返す"
    '''=========================================================================
    ''' <summary>先頭ブロックの先頭基準位置を返す</summary>
    ''' <param name="PosX">(OUT)座標X</param>
    ''' <param name="PosY">(OUT)座標Y</param>
    ''' <returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Private Function XYTableMoveTopBlock(ByRef PosX As Double, ByRef PosY As Double) As Integer

        Dim dblTrimPosX As Double                                   ' ﾄﾘﾑﾎﾟｼﾞｼｮﾝX
        Dim dblTrimPosY As Double                                   ' ﾄﾘﾑﾎﾟｼﾞｼｮﾝY
        Dim dblBSX As Double                                        ' ﾌﾞﾛｯｸｻｲｽﾞX
        Dim dblBSY As Double                                        ' ﾌﾞﾛｯｸｻｲｽﾞY
        Dim dblBsoX As Double                                       ' ﾌﾞﾛｯｸｻｲｽﾞｵﾌｾｯﾄX
        Dim dblBsoY As Double                                       ' ﾌﾞﾛｯｸｻｲｽﾞｵﾌｾｯﾄY
        Dim Del_x As Double                                         ' θ補正量X
        Dim Del_y As Double                                         ' θ補正量Y
        Dim dblRotX As Double                                       ' 回転半径X
        Dim dblRotY As Double                                       ' 回転半径Y
        Dim dblX As Double                                          ' 移動座標X
        Dim dblY As Double                                          ' 移動座標Y
        Dim mdTbOffx As Double                                      ' ﾃｰﾌﾞﾙ位置ｵﾌｾｯﾄX
        Dim mdTbOffy As Double                                      ' ﾃｰﾌﾞﾙ位置ｵﾌｾｯﾄY

        Dim strMSG As String

        Try
            ' ﾄﾘﾑﾎﾟｼﾞｼｮﾝX,Y取得
            ''''(2010/11/16) 動作確認後下記コメントは削除
            'dblTrimPosX = gStartX                                  ' ﾄﾘﾑﾎﾟｼﾞｼｮﾝX,Y取得
            'dblTrimPosY = gStartY
            dblRotX = 0
            dblRotY = 0
            dblTrimPosX = gSysPrm.stDEV.gfTrimX                     ' ﾄﾘﾑﾎﾟｼﾞｼｮﾝX,Y取得
            dblTrimPosY = gSysPrm.stDEV.gfTrimY
            mdTbOffx = typPlateInfo.dblTableOffsetXDir              ' ﾃｰﾌﾞﾙ位置ｵﾌｾｯﾄX,Yの取得
            mdTbOffy = typPlateInfo.dblTableOffsetYDir
            Call CalcBlockSize(dblBSX, dblBSY)                      ' ﾌﾞﾛｯｸｻｲｽﾞ算出

            ' ﾌﾞﾛｯｸｻｲｽﾞｵﾌｾｯﾄ算出(ﾌﾞﾛｯｸｻｲｽﾞ/2 ﾌﾞﾛｯｸの象限はXYともに1 ﾃｰﾌﾞﾙの象限も1)
            dblBsoX = (dblBSX / 2.0#) * 1 * 1
            dblBsoY = (dblBSY / 2.0#) * 1
            ''----- V4.0.0.0-40↓ -----
            '' SL36S時でステージYの原点位置が上の場合、ブロックサイズの1/2は加算しない
            'If (giMachineKd = MACHINE_KD_RS) Then
            '    If (giStageYOrg = STGY_ORG_UP) Then
            '        dblBsoY = 0
            '    End If
            'End If
            ''----- V4.0.0.0-40↑ -----

            ' θ補正ﾄﾘﾑｵﾌｾｯﾄX,Y
            Del_x = gfCorrectPosX
            Del_y = gfCorrectPosY

            ' giBpDirXy 座標系の設定(ｼｽﾃﾑ設定)
            ' 0:XY NOM(右上)  1:X REV(左上)  2:Y REV(右下)  3:XY REV(左下)
            ' ﾄﾘﾐﾝｸﾞ位置座標 (+or-) 回転半径 + ﾃｰﾌﾞﾙｵﾌｾｯﾄ (+or-) ﾌﾞﾛｯｸｻｲｽﾞｵﾌｾｯﾄ + ﾃｰﾌﾞﾙ補正量
            Select Case gSysPrm.stDEV.giBpDirXy
                Case 0 ' x←, y↓
                    dblX = dblTrimPosX + dblRotX + mdTbOffx + dblBsoX + Del_x
                    dblY = dblTrimPosY + dblRotY + mdTbOffy + dblBsoY + Del_y
                Case 1 ' x→, y↓
                    ''###249                    dblX = dblTrimPosX - dblRotX + mdTbOffx - dblBsoX + Del_x
                    'dblX = dblTrimPosX - dblRotX - mdTbOffx - dblBsoX + Del_x
                    dblX = dblTrimPosX - dblRotX - mdTbOffx - dblBsoX - Del_x
                    dblY = dblTrimPosY + dblRotY + mdTbOffy + dblBsoY + Del_y
                Case 2 ' x←, y↑
                    dblX = dblTrimPosX + dblRotX + mdTbOffx + dblBsoX + Del_x
                    'dblY = dblTrimPosY - dblRotY + mdTbOffy - dblBsoY - Del_y ' V1.20.0.0⑤
                    dblY = dblTrimPosY - dblRotY - mdTbOffy - dblBsoY - Del_y  ' V1.20.0.0⑤
                Case 3 ' x→, y↑
                    ''###249                    dblX = dblTrimPosX - dblRotX + mdTbOffx - dblBsoX + Del_x
                    'dblX = dblTrimPosX - dblRotX - mdTbOffx - dblBsoX + Del_x
                    dblX = dblTrimPosX - dblRotX - mdTbOffx - dblBsoX + Del_x
                    'dblY = dblTrimPosY - dblRotY + mdTbOffy - dblBsoY + Del_y
                    'dblY = dblTrimPosY - dblRotY + mdTbOffy - dblBsoY - Del_y' V1.20.0.0⑤
                    dblY = dblTrimPosY - dblRotY - mdTbOffy - dblBsoY - Del_y ' V1.20.0.0⑤
            End Select

            PosX = dblX
            PosY = dblY
            Return (cFRS_NORMAL)                                    ' Return値 = 正常

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTyTeach.XYTableMoveTopBlock() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                      ' Return値 = 例外エラー 
        End Try
    End Function
#End Region

#Region "ブロック移動(XYテーブル移動)【TYティーチング用】"
    '''=========================================================================
    ''' <summary>ブロック移動(XYテーブル移動)【TYティーチング用】</summary>
    ''' <param name="MODE">    (INP)0=第1ｸﾞﾙｰﾌﾟ第1基準位置 , 1=第1ｸﾞﾙｰﾌﾟ最終基準位置
    '''                             2=第1ｸﾞﾙｰﾌﾟ最終基準位置, 3=第2ｸﾞﾙｰﾌﾟ、先頭基準位置
    '''                             4=最終ブロック基準位置</param>
    ''' <param name="iPos">    (INP)ずれ量X,Yﾃｰﾌﾞﾙ(ZRPosX(),ZRPosY())ｲﾝﾃﾞｯｸｽ
    ''' 　　　　　　　　　　　　　　0=第1ｸﾞﾙｰﾌﾟ第1基準位置 , 1=第1ｸﾞﾙｰﾌﾟ最終基準位置
    ''' 　　　　　　　　　　　　　　2=第1ｸﾞﾙｰﾌﾟ最終基準位置, 3=第2ｸﾞﾙｰﾌﾟ先頭基準位置
    ''' </param>
    ''' <param name="Ino">     (INP)処理中のｲﾝﾀｰﾊﾞﾙNo</param>
    ''' <param name="Bn">      (INP)ブロック数</param>
    ''' <param name="ChipSize">(INP)チップサイズ</param>
    ''' <param name="KJPosX">  (INP)先頭ブロックの先頭基準テーブル位置X</param>
    ''' <param name="KJPosY">  (INP)先頭ブロックの先頭基準テーブル位置y</param>
    ''' <param name="TBLx">    (OUT)テーブル位置X</param>
    ''' <param name="TBLy">    (OUT)テーブル位置y</param>
    ''' <returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Private Function XYTableMoveSetBlock(ByRef MODE As Short, ByVal iPos As Short, ByVal Ino As Short, ByVal Bn As Short, _
                                         ByVal ChipSize As Double, ByVal KJPosX As Double, ByVal KJPosY As Double, _
                                         ByRef TBLx As Double, ByRef TBLy As Double) As Short

        Dim wkx As Double
        Dim wky As Double
        Dim i As Short
        Dim iBlock As Short = 0
        Dim dIntv As Double = 0
        Dim ZZx As Double
        Dim ZZy As Double
        Dim r As Short
        Dim strMSG As String

        Try
            Select Case MODE
                Case 0                                                  ' 第1ｸﾞﾙｰﾌﾟ、第1基準位置(チップサイズティーチング用)
                    ' テーブル位置を第1ｸﾞﾙｰﾌﾟ第1基準位置に設定する
                    TBLx = KJPosX + ZRPosX(0)                           ' X = 第1基準位置X + 第1基準位置ずれ量X
                    TBLy = KJPosY + ZRPosY(0)                           ' Y = 第1基準位置Y + 第1基準位置ずれ量Y

                Case 1                                                  ' 第1ｸﾞﾙｰﾌﾟ、最終基準位置(チップサイズティーチング用)
                    ' テーブル位置を第1ｸﾞﾙｰﾌﾟ最終ブロック位置に設定する
                    If (typPlateInfo.intResistDir = 0) Then             ' チップ並びはX方向 ?
                        wkx = ZRPosX(0)                                 ' X = 第1基準位置のずれ量X 
                        wky = (ChipSize * (Bn - 1)) + ZRPosY(0)         ' Y = チップサイズ*(ブロック数-1) + 第1基準位置のずれ量Y
                    Else
                        wkx = (ChipSize * (Bn - 1)) + ZRPosX(0)         ' X = チップサイズ*(ブロック数-1) + 第1基準位置のずれ量X
                        wky = ZRPosY(0)                                 ' Y = 第1基準位置のずれ量Y 
                    End If

                    ' テーブル移動位置設定
                    Call Sub_TblPosXY(wkx, wky)
                    TBLx = KJPosX + wkx + ZRPosX(iPos)                  ' X = 第1基準位置X + 計算値X + 現在のずれ量X
                    TBLy = KJPosY + wky + ZRPosY(iPos)                  ' Y = 第1基準位置Y + 計算値Y + 現在のずれ量Y

                Case 2                                                  ' 第1ｸﾞﾙｰﾌﾟ、最終基準位置(ステージグループ間隔ティーチング用)
                    ' テーブル位置を第1ｸﾞﾙｰﾌﾟの最終ブロック位置に設定する
                    For i = 0 To iPos - 1                               ' ずれ量を加算する
                        ZZx = ZZx + ZRPosX(i)
                        ZZy = ZZy + ZRPosY(i)
                    Next i

                    If (typPlateInfo.intResistDir = 0) Then             ' チップ並びはX方向 ?
                        wkx = ZZx                                       ' X = 第1基準位置のずれ量X 
                        wky = (ChipSize * Bn) + ZRPosY(0)               ' Y = チップサイズ*(ブロック数-1) + 第1基準位置のずれ量Y
                    Else
                        wkx = (ChipSize * Bn) + ZRPosX(0)               ' X = チップサイズ*(ブロック数-1) + 第1基準位置のずれ量X
                        wky = ZZy                                       ' Y = 第1基準位置のずれ量Y 
                    End If

                    ' テーブル移動位置設定
                    Call Sub_TblPosXY(wkx, wky)
                    TBLx = KJPosX + wkx + ZRPosX(iPos)                  ' X = 第1基準位置X + 計算値X + 現在のずれ量X
                    TBLy = KJPosY + wky + ZRPosY(iPos)                  ' Y = 第1基準位置Y + 計算値Y + 現在のずれ量Y

                Case 3                                                  ' 第2ｸﾞﾙｰﾌﾟ、先頭基準位置(ステージグループ間隔ティーチング用)
                    ' テーブル位置を第2ｸﾞﾙｰﾌﾟの先頭ブロック位置に設定する
                    For i = 0 To iPos - 1                               ' ずれ量を加算する
                        ZZx = ZZx + ZRPosX(i)
                        ZZy = ZZy + ZRPosY(i)
                    Next i
                    dIntv = mdStepInterval(1)                           ' ステージグループ間隔 

                    If (typPlateInfo.intResistDir = 0) Then             ' チップ並びはX方向 ?
                        wkx = ZZx                                       ' X = 第1基準位置のずれ量X 
                        wky = (ChipSize * Bn) + dIntv + ZRPosY(0)       ' Y = チップサイズ*ブロック数 + グループ間隔 + 第1基準位置のずれ量Y
                    Else
                        wkx = (ChipSize * Bn) + dIntv + ZRPosX(0)       ' X = チップサイズ*ブロック数 + グループ間隔 + 第1基準位置のずれ量X
                        wky = ZZy                                       ' Y = 第1基準位置のずれ量Y 
                    End If

                    ' テーブル移動位置設定
                    Call Sub_TblPosXY(wkx, wky)
                    TBLx = KJPosX + wkx + ZRPosX(iPos)                  ' X = 第1基準位置X + 計算値X + 現在のずれ量X
                    TBLy = KJPosY + wky + ZRPosY(iPos)                  ' Y = 第1基準位置Y + 計算値Y + 現在のずれ量Y

                Case 4                                                  ' ステップオフセットティーチング用
                    ' テーブル位置を最終ブロック位置に設定する
                    For i = 0 To iPos - 1                               ' ずれ量を加算する
                        ZZx = ZZx + ZRPosX(i)
                        ZZy = ZZy + ZRPosY(i)
                    Next i

                    ' ステージグループ間隔 = (ステージグループ数 -1) * ステージグループ間隔
                    dIntv = (typPlateInfo.intGroupCntInBlockYStage - 1) * mdStepInterval(1)

                    If (typPlateInfo.intResistDir = 0) Then             ' チップ並びはX方向 ?
                        wkx = ZZx                                       ' X = 第1基準位置のずれ量X 
                        wky = (ChipSize * Bn) + dIntv + ZRPosY(0)       ' Y = チップサイズ*ブロック数 + グループ間隔 + 第1基準位置のずれ量Y
                    Else
                        wkx = (ChipSize * Bn) + dIntv + ZRPosX(0)       ' X = チップサイズ*ブロック数 + グループ間隔 + 第1基準位置のずれ量X
                        wky = ZZy                                       ' Y = 第1基準位置のずれ量Y 
                    End If

                    ' テーブル移動位置設定
                    Call Sub_TblPosXY(wkx, wky)
                    TBLx = KJPosX + wkx + ZRPosX(iPos)                  ' X = 第1基準位置X + 計算値X + 現在のずれ量X
                    TBLy = KJPosY + wky + ZRPosY(iPos)                  ' Y = 第1基準位置Y + 計算値Y + 現在のずれ量Y

            End Select

            ' テーブル移動
            r = Form1.System1.XYtableMove(gSysPrm, TBLx, TBLy)
            Return (r)

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTyTeach.XYTableMoveSetBlock() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Return値 = 例外エラー 
        End Try
    End Function
#End Region

#Region "テーブル位置計算サブルーチン"
    '''=========================================================================
    ''' <summary>テーブル位置計算サブルーチン</summary>
    ''' <param name="wkx">    (I/O)テーブル位置X</param>
    ''' <param name="wky">    (I/O)テーブル位置y</param>
    '''=========================================================================
    Private Sub Sub_TblPosXY(ByRef wkx As Double, ByRef wky As Double)

        Dim strMSG As String

        Try

            If (typPlateInfo.intResistDir = 0) Then                     ' チップ並びはX方向 ?
                Select Case gSysPrm.stDEV.giBpDirXy
                    Case 0 'x←, y↓
                        wky = wky
                    Case 1 'x→, y↓
                        wky = wky
                    Case 2 'x←, y↑
                        wky = -wky
                    Case 3 'x→, y↑
                        wky = -wky
                End Select
            Else
                Select Case gSysPrm.stDEV.giBpDirXy
                    Case 0 'x←, y↓
                        wkx = wkx
                    Case 1 'x→, y↓
                        wkx = -wkx
                    Case 2 'x←, y↑
                        wkx = wkx
                    Case 3 'x→, y↑
                        wkx = -wkx
                End Select
            End If

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTyTeach.XYTableMoveSetBlock() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

    '========================================================================================
    '   ボタン押下時処理
    '========================================================================================
#Region "HALTボタン押下時処理"
    '''=========================================================================
    '''<summary>HALTボタン押下時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnHALT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnHALT.Click
        'Call SubBtnHALT_Click() ' '###265
    End Sub
#End Region

#Region "STARTボタン押下時処理"
    '''=========================================================================
    '''<summary>STARTボタン押下時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnSTART_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSTART.Click
        'Call SubBtnSTART_Click() ' '###265
    End Sub
#End Region

#Region "RESETボタン押下時処理"
    '''=========================================================================
    '''<summary>RESETボタン押下時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnRESET_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnRESET.Click
        'Call SubBtnRESET_Click() '###265
    End Sub
#End Region

    Private Sub BtnSTART_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnSTART.MouseDown
        Call SubBtnSTART_Click()
    End Sub

    Private Sub BtnHALT_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnHALT.MouseDown
        Call SubBtnHALT_Click()
    End Sub

    Private Sub BtnRESET_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnRESET.MouseDown
        Call SubBtnRESET_Click()
    End Sub


#Region "Zボタン押下時処理"
    '''=========================================================================
    '''<summary>RESETボタン押下時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnZ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnZ.Click
        Call SubBtnZ_Click(stJOG)
    End Sub
#End Region

#Region "HIボタン押下時処理"
    '''=========================================================================
    '''<summary>HIボタン押下時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnHI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnHI.Click
        'Call SubBtnHI_Click(stJOG) 
    End Sub
#End Region
    Private Sub BtnHI_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnHI.MouseDown
        Call SubBtnHI_Click(stJOG)
    End Sub

#Region "矢印ボタン押下時"
    '''=========================================================================
    '''<summary>矢印ボタン押下時</summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub BtnJOG_0_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_0.MouseDown
        Timer1.Enabled = False
        Call SubBtnJOG_0_MouseDown()                                    ' +Y ON
    End Sub
    Private Sub BtnJOG_0_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_0.MouseUp
        Call SubBtnJOG_0_MouseUp()                                      ' +Y OFF
        Timer1.Enabled = True
    End Sub

    Private Sub BtnJOG_1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_1.MouseDown
        Timer1.Enabled = False
        Call SubBtnJOG_1_MouseDown()                                    ' -Y ON
    End Sub
    Private Sub BtnJOG_1_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_1.MouseUp
        Call SubBtnJOG_1_MouseUp()                                      ' -Y OFF
        Timer1.Enabled = True
    End Sub
    Private Sub BtnJOG_2_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_2.MouseDown
        Timer1.Enabled = False
        Call SubBtnJOG_2_MouseDown()                                    ' +X ON
    End Sub
    Private Sub BtnJOG_2_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_2.MouseUp
        Call SubBtnJOG_2_MouseUp()                                      ' +X OFF
        Timer1.Enabled = True
    End Sub

    Private Sub BtnJOG_3_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_3.MouseDown
        Timer1.Enabled = False
        Call SubBtnJOG_3_MouseDown()                                    ' -X ON
    End Sub
    Private Sub BtnJOG_3_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_3.MouseUp
        Call SubBtnJOG_3_MouseUp()                                      ' -X OFF
        Timer1.Enabled = True
    End Sub

    Private Sub BtnJOG_4_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_4.MouseDown
        Timer1.Enabled = False
        Call SubBtnJOG_4_MouseDown()                                    ' -X -Y ON
    End Sub
    Private Sub BtnJOG_4_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_4.MouseUp
        Call SubBtnJOG_4_MouseUp()                                      ' -X -Y OFF
        Timer1.Enabled = True
    End Sub

    Private Sub BtnJOG_5_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_5.MouseDown
        Timer1.Enabled = False
        Call SubBtnJOG_5_MouseDown()                                    ' +X -Y ON
    End Sub
    Private Sub BtnJOG_5_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_5.MouseUp
        Call SubBtnJOG_5_MouseUp()                                      ' +X -Y OFF
        Timer1.Enabled = True
    End Sub

    Private Sub BtnJOG_6_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_6.MouseDown
        Timer1.Enabled = False
        Call SubBtnJOG_6_MouseDown()                                    ' +X +Y ON
    End Sub
    Private Sub BtnJOG_6_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_6.MouseUp
        Call SubBtnJOG_6_MouseUp()                                      ' +X +Y OFF
        Timer1.Enabled = True
    End Sub

    Private Sub BtnJOG_7_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_7.MouseDown
        Timer1.Enabled = False
        Call SubBtnJOG_7_MouseDown()                                    ' -X +Y ON
    End Sub
    Private Sub BtnJOG_7_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_7.MouseUp
        Call SubBtnJOG_7_MouseUp()                                      ' -X +Y OFF
        Timer1.Enabled = True
    End Sub
#End Region

    '========================================================================================
    '   テンキー入力処理
    '========================================================================================
#Region "キーダウン時処理"
    '''=========================================================================
    '''<summary>キーダウン時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub frmTxTyTeach_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown

        Dim KeyCode As Short = e.KeyCode
        Console.WriteLine("frmTxTyTeach.frmTxTyTeach_KeyDown()")

        ' テンキーダウンならInpKeyにテンキーコードを設定する
        GrpMain.Enabled = False
        Call Sub_10KeyDown(KeyCode)
        If (KeyCode = System.Windows.Forms.Keys.NumPad5) Then           ' 5ｷｰ (KeyCode = 101(&H65)
            Call SubBtnHI_Click(stJOG)
        End If
        Call Me.Focus()

    End Sub
#End Region

    '========================================================================================
    '   トラックバー処理
    '========================================================================================
#Region "トラックバーのスライダー移動イベント"
    '''=========================================================================
    ''' <summary>トラックバーのスライダー移動イベント</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub TBarLowPitch_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBarLowPitch.Scroll
        Call SetSliderPitch(IDX_PIT, TBarLowPitch, TBarHiPitch, TBarPause, LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval)
    End Sub

    Private Sub TBarHiPitch_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBarHiPitch.Scroll
        Call SetSliderPitch(IDX_HPT, TBarLowPitch, TBarHiPitch, TBarPause, LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval)
    End Sub

    Private Sub TBarPause_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBarPause.Scroll
        Call SetSliderPitch(IDX_PAU, TBarLowPitch, TBarHiPitch, TBarPause, LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval)
    End Sub
#End Region

    '----- ###228↓ -----
#Region "キーステータス取得"
    '''=========================================================================
    ''' <summary>キーステータス取得</summary>
    ''' <param name="Code">(INP)キーコード</param>
    ''' <returns></returns>
    '''=========================================================================
    Private Function Sub_GetKeyState(ByVal Code As Integer) As Integer

        Dim keyState As Integer

        keyState = GetKeyState(Code)
        Return (keyState)
    End Function
#End Region

#Region "キーステータスチェックタイマー"
    '''=========================================================================
    ''' <summary>キーステータスチェックタイマー</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick

    End Sub
#End Region


#Region "Execute処理"
    '''=========================================================================
    '''<summary>Execute処理</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Public Function Execute() As Integer Implements ICommonMethods.Execute

        Dim strMSG As String

        Try
            ' ＴＸ/ＴＹティーチング開始
            If (mExit_flg <> -1) Then Exit Function
            mExit_flg = 0                                           ' 終了フラグ = 0
            mExit_flg = TxTyMainProc()                              ' ＴＸまたはＴＹティーチング開始

            ' 補正クロスラインを非表示とする
            Form1.CrosLineX.Visible = False
            Form1.CrosLineY.Visible = False

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmTxTeach.frmTxTeach_Activated() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            mExit_flg = cERR_TRAP                                   ' Return値 = 例外エラー
        End Try
        Me.Close()

    End Function

#End Region

    '========================================================================================
    '   テンキー入力処理
    '========================================================================================

#Region "キーアップ時処理"
    '''=========================================================================
    '''<summary>キーアップ時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub frmTxTyTeach_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
        JogKeyUp(e)
    End Sub


    '========================================================================================
    '   テンキー入力処理
    '========================================================================================
#Region "キーダウン時処理"
    '''=========================================================================
    ''' <summary>キーダウン時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub frmFineAdjust_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        JogKeyDown(e)

    End Sub

    Public Sub JogKeyDown(ByVal e As KeyEventArgs) Implements ICommonMethods.JogKeyDown

        Dim KeyCode As Short = e.KeyCode

        ' テンキー入力フラグがOFFならNOP
        If (gbTenKeyFlg = False) Then
            e.Handled = False
            Exit Sub
        End If

        ' テンキーダウンならInpKeyにテンキーコードを設定する
        Call Sub_10KeyDown(KeyCode)
        If (KeyCode = System.Windows.Forms.Keys.NumPad5) Then       ' 5ｷｰ (KeyCode = 101(&H65)
            Call BtnHI_Click(BtnHI, e)                             ' HIボタン ON/OFF
        End If

    End Sub
#End Region

#Region "キーアップ時処理"
    '''=========================================================================
    ''' <summary>キーアップ時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub frmFineAdjust_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp

        Me.JogKeyUp(e)

    End Sub


    Public Sub JogKeyUp(ByVal e As KeyEventArgs) Implements ICommonMethods.JogKeyUp

        Dim KeyCode As Short = e.KeyCode

        ' テンキー入力フラグがOFFならNOP 
        If (gbTenKeyFlg = False) Then Exit Sub

        ' テンキーアップならInpKeyのテンキーコードをOFFする
        Call Sub_10KeyUp(KeyCode)
        'Call Me.Focus()

    End Sub

#End Region
#Region "カメラ画像クリック位置を画像センターに移動する処理"
    ''' <summary>カメラ画像クリック位置を画像センターに移動する処理</summary>
    ''' <param name="distanceX">画像センターからの距離X</param>
    ''' <param name="distanceY">画像センターからの距離Y</param>
    ''' <remarks></remarks>
    Public Sub MoveToCenter(ByVal distanceX As Decimal, ByVal distanceY As Decimal) _
        Implements ICommonMethods.MoveToCenter

        ' テンキー入力フラグがOFFならNOP 
        If (gbTenKeyFlg = False) Then
            Exit Sub
        End If

        UserModule.MoveToCenter(distanceX, distanceY, stJOG)
    End Sub
#End Region


#End Region


End Class
