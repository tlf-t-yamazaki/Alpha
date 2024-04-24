'===============================================================================
'   Description  : トリミング実行時一時停止処理
'
'   Copyright(C) : TOWA LASERFRONT CORP. 2012
'
'===============================================================================
Option Strict Off
Option Explicit On

Imports LaserFront.Trimmer.DllSysPrm.SysParam

Friend Class frmFineAdjust
    Inherits System.Windows.Forms.Form
    Implements ICommonMethods              'V2.2.0.0①

    '========================================================================================
    '   定数・変数定義
    '========================================================================================
#Region "定数・変数定義"
    '===========================================================================
    '   定数定義
    '===========================================================================
    Public Const MOVE_NEXT As Integer = 0
    Public Const MOVE_NOT As Integer = 1

    '----- 処理モード -----
    Private Const MD_INI As Integer = 0                                 ' 初期エントリモード
    Private Const MD_CHK As Integer = 1                                 ' 継続エントリモード

    '===========================================================================
    '   メンバ変数定義
    '===========================================================================
    Private m_BlockSizeX As Double
    Private m_BlockSizeY As Double
    Private m_bpOffX As Double
    Private m_bpOffY As Double
    Private m_sysPrm As SYSPARAM_PARAM
    Private stJOG As JOG_PARAM                                          ' 矢印画面(BPのJOG操作)用パラメータ (Globals.vbの共通関数を使用)
    Private dblTchMoval(3) As Double                                    ' ﾋﾟｯﾁ退避域(0=ﾋﾟｯﾁ, 1=HIGHﾋﾟｯﾁ, 2=Pause Time(Sec))
    Private mExit_flg As Short                                          ' 結果
    Private mMd As Integer = MD_INI                                     ' 処理モード
    Private m_TenKeyFlg As Boolean = False
    Private m_LaserOnOffFlag As Boolean = False

#End Region

    '========================================================================================
    '   メソッド定義
    '========================================================================================
#Region "初期値設定処理"
    '''=========================================================================
    ''' <summary>初期値設定処理</summary>
    ''' <param name="SysPrm"></param>
    ''' <param name="digL"></param>
    ''' <param name="digH"></param>
    ''' <param name="curPltNo"></param>
    ''' <param name="curBlkNo"></param>
    '''=========================================================================
    Public Sub SetInitialData(ByRef SysPrm As SYSPARAM_PARAM, _
                        ByVal digL As Integer, ByVal digH As Integer, _
                        ByRef curPltNo As Integer, ByRef curBlkNo As Integer)

        Try
            'CbDigSwH.SelectedIndex = digH
            'CbDigSwL.SelectedIndex = digL
            gCurBlockNo = curBlkNo
            gCurPlateNo = curPltNo
            m_sysPrm = SysPrm
            'gFrmEndStatus = cFRS_NORMAL

            If (gbChkboxHalt = True) Then                                       '###009
                BtnADJ.Text = "ADJ ON"                                          '###009
                BtnADJ.BackColor = System.Drawing.Color.Yellow                  '###009
            Else                                                                '###009
                BtnADJ.Text = "ADJ OFF"                                         '###009
                BtnADJ.BackColor = System.Drawing.SystemColors.Control          '###009
            End If                                                              '###009

            ' ラベル名設定(日本語/英語)
            'BtnEdit.Text = "データ編集"                                      ' "データ編集" ###014
            '-----###204 -----
            Me.Label3.Text = "調整"                                    ' "調整" 
            'CbDigSwH.Items(0) = "０：表示なし"
            'CbDigSwH.Items(1) = "１：ＮＧのみ表示"
            'CbDigSwH.Items(2) = "２：全て表示"
            '-----###204 -----
            '----- ###268↓ -----
            '「Ten Key On/Off」ボタンの初期値をシスパラより設定する
            If (giTenKey_Btn = 0) Then                                          ' 一時停止画面での「Ten Key On/Off」ボタンの初期値(0:ON(既定値), 1:OFF)
                gbTenKeyFlg = True
                BtnTenKey.Text = "Ten Key On"
                BtnTenKey.BackColor = System.Drawing.Color.Pink
            Else
                gbTenKeyFlg = False
                BtnTenKey.Text = "Ten Key Off"
                BtnTenKey.BackColor = System.Drawing.SystemColors.Control
            End If

            'gbTenKeyFlg = True                                                 ' 「Ten Key On」状態 ###242
            '----- ###268↑ -----
            '----- ###269↓ -----
            ' 一時停止画面でのシスパラ「BPオフセット調整する/しない」指定により矢印ボタン等を設定する
            Call Sub_SetBtnArrowEnable()
            '----- ###269↑ -----

            'for　抵抗数分
            '目標値
            'カットオフ
            'スピード
            '加工条件番号
            'next
            'txtExCamPosX.Text = m_sysPrm.stDEV.gfExCmX.ToString


            'V2.2.0.0④ ↓
            If giMouseClickMove = 1 Then
                BtnClickEnable.BackColor = SystemColors.Control
                gbTenKeyFlg = False
            Else
                gbTenKeyFlg = True
            End If
            'V2.2.0.0④ ↑


        Catch ex As Exception
            Dim strMsg As String
            strMsg = "frmFineAdjust.Form_Initialize_Renamed() TRAP ERROR = " + ex.Message
            MsgBox(strMsg)
        End Try

    End Sub
#End Region
    '-----###269↓-----
#Region "矢印ボタンを活性化/非活性化する"
    '''=========================================================================
    ''' <summary>矢印ボタンを活性化/非活性化する</summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub Sub_SetBtnArrowEnable()

        Dim bFlg As Boolean
        Dim strMsg As String

        Try
            '  一時停止画面でのシスパラ「BPオフセット調整する/しない」指定により矢印ボタン等を設定する
            If (giBpAdj_HALT = 0) Then                                          ' BPオフセット調整する ?
                bFlg = True

                'V2.2.0.0①↓
                Form1.SetActiveJogMethod(AddressOf Me.JogKeyDown,
                                                  AddressOf Me.JogKeyUp,
                                                  AddressOf Me.MoveToCenter)    'V6.0.0.0⑩
                'V2.2.0.0①↑

            Else                                                                ' BPオフセット調整しない
                bFlg = False
                gbTenKeyFlg = False
                BtnTenKey.Enabled = False                                       '「Ten Key Off」ボタン非活性化
                BtnTenKey.Text = "Ten Key Off"
                BtnTenKey.BackColor = System.Drawing.SystemColors.Control

                Form1.SetActiveJogMethod(Nothing, Nothing, Nothing)    'V2.2.0.0①

            End If

            If giMouseClickMove = 1 Then
                BtnClickEnable.Enabled = True
                BtnClickEnable.Visible = True
                BtnClickEnable.BackColor = SystemColors.Control
            Else
                BtnClickEnable.Enabled = False
                BtnClickEnable.Visible = False
            End If

            ' 矢印ボタン活性化/非活性化
            BtnJOG_0.Enabled = bFlg
            BtnJOG_1.Enabled = bFlg
            BtnJOG_2.Enabled = bFlg
            BtnJOG_3.Enabled = bFlg
            BtnJOG_4.Enabled = bFlg
            BtnJOG_5.Enabled = bFlg
            BtnJOG_6.Enabled = bFlg
            BtnJOG_7.Enabled = bFlg
            BtnHI.Enabled = bFlg

            ' Moving Pitch活性化/非活性化
            GrpPithPanel.Enabled = bFlg

        Catch ex As Exception
            strMsg = "frmFineAdjust.Sub_SetBtnArrowEnable() TRAP ERROR = " + ex.Message
            MsgBox(strMsg)
        End Try
    End Sub
#End Region
    '----- ###269↑-----
    '----- ###260↓-----
#Region "タイマー停止"
    '''=========================================================================
    ''' <summary>タイマー停止</summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function Sub_StopTimer() As Integer

        TmKeyCheck.Enabled = False

    End Function
#End Region
    '----- ###260↑-----

#Region "ステージポジション取得処理"
    '''=========================================================================
    ''' <summary>ステージポジション取得処理（実行後に取得）</summary>
    '''=========================================================================
    Public Sub GetStagePosInfo(ByRef pltNo As Integer, ByRef blkNo As Integer)

        Try
            pltNo = gCurPlateNo
            blkNo = gCurBlockNo

        Catch ex As Exception
            Dim strMsg As String
            strMsg = "frmFineAdjust.GetStagePosInfo() TRAP ERROR = " + ex.Message
            MsgBox(strMsg)
        End Try
    End Sub
#End Region

#Region "終了戻り値取得処理"
    '''=========================================================================
    ''' <summary>終了戻り値取得処理（実行後に取得）</summary>
    '''=========================================================================
    Public Function GetReturnVal() As Integer

        Try
            Return (mExit_flg)

        Catch ex As Exception
            Dim strMsg As String
            strMsg = "frmFineAdjust.GetReturnVal() TRAP ERROR = " + ex.Message
            MsgBox(strMsg)
        End Try
    End Function
#End Region

    '========================================================================================
    '   画面処理
    '========================================================================================
#Region "フォームロード処理"
    '''=========================================================================
    ''' <summary>フォームロード処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub frmFineAdjust_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Dim stPos As System.Drawing.Point
        'Dim stGetPos As System.Drawing.Point
        'Dim r As Integer                                                ' ###237
        Dim strMsg As String

        Try
            ' 表示位置の調整
            'stPos = Form1.Text4.PointToScreen(stGetPos)
            'stPos.X = stPos.X - 2
            'stPos.Y = stPos.Y - 2
            'Me.Location = stPos
            Me.Location = New Point(Form1.VideoLibrary1.Location.X + Form1.VideoLibrary1.Size.Width + 6, Form1.Grpcmds.Location.Y)
            'Me.Height = Form1.frmInfo.Location.Y - Form1.Grpcmds.Location.Y
            Me.Height = Form1.Grpcmds.Size.Height
            ' BpOffsetの現在値設定
            GetBpOffset(m_bpOffX, m_bpOffY)
            txtBpOffX.Text = m_bpOffX.ToString
            txtBpOffY.Text = m_bpOffY.ToString

            ' BlockSizeの現在値取得
            GetBlockSize(m_BlockSizeX, m_BlockSizeY)

            '----- ###139↓ -----
            ' メイン画面の「生産グラフ表示/非表示ボタン」から当画面の「生産グラフ表示/非表示ボタン」を設定する
            'If (gTkyKnd = KND_CHIP Or gTkyKnd = KND_NET) Then
            '    chkDistributeOnOff.Text = Form1.chkDistributeOnOff.Text
            '    chkDistributeOnOff.Checked = Form1.chkDistributeOnOff.Checked
            '    GrpDistribute.Visible = True                        '「生産グラフボタン」表示

            'Else
            'GrpDistribute.Visible = False                       '「生産グラフボタン」非表示
            'End If
            '----- ###139↑ -----

            '----- ###237↓ -----
            ' 加工条件番号を設定する(FL時)
            'If (gSysPrm.stRAT.giOsc_Res = OSCILLATOR_FL) Then
            '    Call QRATE(stCND.Freq(ADJ_CND_NUM))                     ' Qレート設定(KHz)
            '    r = FLSET(FLMD_CNDSET, ADJ_CND_NUM)                     ' 加工条件番号設定(一時停止画面用)
            'Else
            '' '' ''Call QRATE(gSysPrm.stDEV.gfLaserQrate)                  ' Qレート設定(KHz) ※レーザ調整用Qレートを設定
            'End If
            '----- ###237↑ -----

            Call PrepareMessages(gSysPrm.stTMN.giMsgTyp)
            ' フォーカスの設定(これによってテンキーのイベントが取得できる)
            Me.KeyPreview = True
            Me.Activate()                                               ' ###046

        Catch ex As Exception
            strMsg = "frmFineAdjust.frmFineAdjust_Load() TRAP ERROR = " + ex.Message
            MsgBox(strMsg)
        End Try
    End Sub
#End Region

#Region "フォームが表示された時の処理"
    '''=========================================================================
    ''' <summary>フォームが表示された時の処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub frmFineAdjust_Shown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Shown

        Dim r As Integer = cFRS_NORMAL
        Dim strMSG As String

        Try
            ' 一時停止画面処理メインをCallする
            mExit_flg = 0                                               ' 終了フラグ = 0
            Call ZCONRST()                                              ' ｺﾝｿｰﾙｷｰﾗｯﾁ解除
            TmKeyCheck.Interval = 10
            TmKeyCheck.Enabled = True                                   ' タイマー開始
            Return

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmFineAdjust.frmFineAdjust_Shown() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            mExit_flg = cERR_TRAP                                       ' Return値 = 例外エラー
        End Try

        gbExitFlg = True                                                ' 終了フラグON
        Call LASEROFF()                                                 ' ###237
        Me.Close()
    End Sub
#End Region

#Region "キー入力チェックタイマー処理"
    '''=========================================================================
    ''' <summary>キー入力チェックタイマー処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub TmKeyCheck_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles TmKeyCheck.Tick

        Dim r As Short
        Dim strMSG As String

        Try

            If gbExitFlg Then
                Exit Sub
            End If
            ' キー入力チェック処理
            TmKeyCheck.Enabled = False                                  ' タイマー停止
            r = MainProc(mMd)                                           ' 一時停止画面処理
            If (r = cFRS_NORMAL) Then                                   ' 正常戻り 
                TmKeyCheck.Enabled = True                               ' タイマー開始
                Return
            End If

            '----- ###219↓ -----
            ' Z キー押下なら Z On/OFFする 
            If (r = cFRS_ERR_Z) Then                                    ' Z SW押下 ?
                If (stJOG.bZ = True) Then                               ' Z ON ? 
                    r = Prob_On()
                Else                                                    ' Z OFF
                    r = Prob_Off()
                End If
                ' エラーならメッセージを表示してエラーリターン
                r = Form1.System1.EX_ZGETSRVSIGNAL(gSysPrm, r, 0)       ' エラーならメッセージを表示する
                If (r <> cFRS_NORMAL) Then
                    mExit_flg = r                                       ' エラーリターン 
                    Return
                End If

                ' Zランプの点灯/消灯
                If (stJOG.bZ = True) Then
                    Call LAMP_CTRL(LAMP_Z, True)
                Else
                    Call LAMP_CTRL(LAMP_Z, False)
                End If

                TmKeyCheck.Enabled = True                               ' タイマー開始
                Return
            End If
            '----- ###219↑ -----

            ' START/RESETキー押下またはエラーなら終了
            If (r = cFRS_ERR_START) Then r = cFRS_NORMAL

            mExit_flg = r

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmFineAdjust.TmKeyCheck_Tick() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            mExit_flg = cERR_TRAP                                       ' Return値 = 例外エラー
        End Try

        gbExitFlg = True                                                ' 終了フラグON
        Call LASEROFF()                                                 ' ###237
        Me.Close()
    End Sub
#End Region

#Region "メイン処理"
    '''=========================================================================
    ''' <summary>メイン処理"</summary>
    ''' <param name="Md">(I/O)処理モード
    ''' 　　　　　　　　　　　MD_INI=初期エントリ, MD_CHK=継続エントリ</param>
    ''' <returns>cFRS_NORMAL   = OK(STARTｷｰ)
    '''          cFRS_ERR_RST  = Cancel(RESETｷｰ)
    '''          -1以下        = エラー</returns>
    '''=========================================================================
    Private Function MainProc(ByRef Md As Integer) As Short

        Dim mdAdjx As Double = 0.0                                      ' ｱｼﾞｬｽﾄ位置X(未使用)
        Dim mdAdjy As Double = 0.0                                      ' ｱｼﾞｬｽﾄ位置Y(未使用)
        Dim r As Short
        Dim strMSG As String
        Dim cControl As Control = Me.ActiveControl

        Try
            '-------------------------------------------------------------------
            '   初期処理
            '-------------------------------------------------------------------
            If (Md = MD_INI) Then                                       ' 初期エントリ
                ' JOGパラメータ設定 
                stJOG.Md = MODE_BP                                      ' モード(1:BP移動)
                stJOG.Md2 = MD2_BUTN                                    ' 入力モード(0:画面ﾎﾞﾀﾝ入力, 1:ｺﾝｿｰﾙ入力)
                '                                                       ' キーの有効(1)/無効(0)指定
                'stJOG.Opt = CONSOLE_SW_RESET + CONSOLE_SW_START
                stJOG.Opt = CONSOLE_SW_RESET + CONSOLE_SW_START + CONSOLE_SW_ZSW ' ###219
                stJOG.PosX = 0.0                                        ' BP X位置(BPｵﾌｾｯﾄX)
                stJOG.PosY = 0.0                                        ' BP Y位置(BPｵﾌｾｯﾄY)
                stJOG.BpOffX = mdAdjx + m_bpOffX                        ' BPｵﾌｾｯﾄX 
                stJOG.BpOffY = mdAdjy + m_bpOffY                        ' BPｵﾌｾｯﾄY 
                stJOG.BszX = m_BlockSizeX                               ' ﾌﾞﾛｯｸｻｲｽﾞX 
                stJOG.BszY = m_BlockSizeY                               ' ﾌﾞﾛｯｸｻｲｽﾞY
                txtBpOffX.ShortcutsEnabled = False                      ' ###047 右クリックメニューを表示しない 
                txtBpOffY.ShortcutsEnabled = False                      '  
                stJOG.TextX = txtBpOffX                                 ' BP X位置表示用ﾃｷｽﾄﾎﾞｯｸｽ
                stJOG.TextY = txtBpOffY                                 ' BP Y位置表示用ﾃｷｽﾄﾎﾞｯｸｽ
                stJOG.cgX = m_bpOffX                                    ' 移動量X (BPｵﾌｾｯﾄX)
                stJOG.cgY = m_bpOffY                                    ' 移動量Y (BPｵﾌｾｯﾄY)
                stJOG.BtnHI = BtnHI                                     ' HIボタン
                stJOG.BtnZ = BtnZ                                       ' Zボタン
                stJOG.BtnSTART = BtnSTART                               ' STARTボタン
                stJOG.BtnRESET = BtnRESET                               ' RESETボタン
                stJOG.BtnHALT = BtnHALT                                 ' HALTボタン
                Call JogEzInit(stJOG, TBarLowPitch, TBarHiPitch, TBarPause, LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval)
                stJOG.Flg = -1                                          ' 親画面のOK/Cancelﾎﾞﾀﾝ押下ﾌﾗｸﾞ
                Md = MD_CHK
                stJOG.bZ = False                                        ' JogのZキー状態 = Z Off ###219
                Call LAMP_CTRL(LAMP_Z, False)                           ' ###219 
            End If

STP_RETRY:
            'Call Me.Focus()                                            ' ← これをやるとテンキーのKeyUp/KeyDownイベントが入ってこなくなる

            ' 非常停止等チェック
            r = Form1.System1.Sys_Err_Chk_EX(gSysPrm, giAppMode)
            If (r <> cFRS_NORMAL) Then                                  ' 非常停止等検出 ?
                Call Form1.AppEndDataSave()                                       ' ｿﾌﾄ強制終了時のﾃﾞｰﾀ保存確認
                Call Form1.AplicationForcedEnding()                     ' ｿﾌﾄ強制終了処理
                End                                                     ' アプリ強制終了
                '                Return (r)
            End If

            ''----- ###209↓ -----
            '' カバー閉を確認する(SL436R時で手動モード時)
            'If (gSysPrm.stTMN.gsKeimei = MACHINE_TYPE_SL436) And (bFgAutoMode = False) Then
            '    Call COVERLATCH_CLEAR()                                 ' カバー開ラッチのクリア
            '    r = FrmReset.Sub_CoverCheck()
            '    If (r < cFRS_NORMAL) Then                               ' 非常停止等検出 ?
            '        Return (r)
            '    End If
            'End If
            ''----- ###209↑ -----
            'V2.2.0.0① ↓
            'If System.Windows.Forms.Form.ActiveForm IsNot Nothing Then
            '    If System.Windows.Forms.Form.ActiveForm.Text <> "ADJFINE" Then
            '        Call ClearInpKey()
            '    End If
            'End If
            'V2.2.0.0① ↑
            ' コンソールキー等の入力待ち
            'stJOG.Flg = -1                                             ' 親画面のOK/Cancelﾎﾞﾀﾝ押下ﾌﾗｸﾞ
            r = JogEzMove_Ex(stJOG, gSysPrm, TBarLowPitch, TBarHiPitch, TBarPause, LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval)
            If (r < cFRS_NORMAL) Then                                   ' エラーなら終了
                Return (r)
            End If

            ' コンソールキーチェック
            If (r = cFRS_ERR_START) Then                                ' START SW押下 ?
                ' DIG-SW設定
                'Call Form1.SetMoveMode(CbDigSwL.SelectedIndex, CbDigSwH.SelectedIndex)
                ' BPオフセット更新(タイミングによって空白で入ってくる場合トラップエラーとなるのでチェックする ###014)
                If (txtBpOffX.Text <> "") And (txtBpOffY.Text <> "") Then
                    Call SetBpOffset(Double.Parse(txtBpOffX.Text), Double.Parse(txtBpOffY.Text))
                End If
                Return (cFRS_ERR_START)

            ElseIf (r = cFRS_ERR_RST) Then                              ' RESET SW押下 ?
                Return (cFRS_ERR_RST)

                '----- ###219↓ -----
            ElseIf (r = cFRS_ERR_Z) Then                                ' Z SW押下 ?
                Return (cFRS_ERR_Z)
                '----- ###219↑ -----
            End If

            'Loop While (stJOG.Flg = -1)

            '' 当画面からOK/Cancelﾎﾞﾀﾝ押下ならrに戻値を設定する
            'If (stJOG.Flg <> -1) Then
            '    r = stJOG.Flg
            'End If

STP_END:
            Return (r)                                                  ' Return値設定

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmFineAdjust.MainProc() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Return値 = 例外エラー
        End Try
    End Function
#End Region

    '========================================================================================
    '   メイン画面のボタン押下時処理
    '========================================================================================
#Region "ADJﾎﾞﾀﾝ押下時処理"
    '''=========================================================================
    ''' <summary>ADJﾎﾞﾀﾝ押下時処理 ###009</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub BtnADJ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnADJ.Click
        Dim strMSG As String

        Try
            If (BtnADJ.Text = "ADJ OFF") Then
                gbChkboxHalt = True
                BtnADJ.Text = "ADJ ON"
                BtnADJ.BackColor = System.Drawing.Color.Yellow
            Else
                gbChkboxHalt = False
                BtnADJ.Text = "ADJ OFF"
                BtnADJ.BackColor = System.Drawing.SystemColors.Control
            End If

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "FrmFineAdjust.BtnADJ_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "Ten Key On/Offボタン押下時処理"
    '''=========================================================================
    ''' <summary>Ten Key On/Offボタン押下時処理 ###057</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnTenKey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnTenKey.Click

        'Dim InpKey As UShort
        Dim strMSG As String

        Try
            Call SubBtnTenKey_Click()                                   ' ###139

            '' InpKeyのHI SW以外はOFFする' ###139
            'GetInpKey(InpKey)
            'If (InpKey And cBIT_HI) Then                                ' HI SW ON ?
            '    InpKey = cBIT_HI
            'Else
            '    InpKey = 0
            'End If
            'PutInpKey(InpKey)

            '' Ten Key On/Offボタン設定
            'If (BtnTenKey.Text = "Ten Key Off") Then
            '    gbTenKeyFlg = True
            '    BtnTenKey.Text = "Ten Key On"
            '    BtnTenKey.BackColor = System.Drawing.Color.Pink
            'Else
            '    gbTenKeyFlg = False
            '    BtnTenKey.Text = "Ten Key Off"
            '    BtnTenKey.BackColor = System.Drawing.SystemColors.Control
            'End If

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "FrmFineAdjust.BtnTenKey_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "Ten Key On/Offボタン押下時処理"
    '''=========================================================================
    ''' <summary>Ten Key On/Offボタン押下時処理 ###139</summary>
    '''=========================================================================
    Private Sub SubBtnTenKey_Click()

        Dim InpKey As UShort
        Dim strMSG As String

        Try
            ' InpKeyのHI SW以外はOFFする
            GetInpKey(InpKey)
            If (InpKey And cBIT_HI) Then                                ' HI SW ON ?
                InpKey = cBIT_HI
            Else
                InpKey = 0
            End If
            PutInpKey(InpKey)

            ' Ten Key On/Offボタン設定
            If (BtnTenKey.Text = "Ten Key Off") Then
                gbTenKeyFlg = True
                BtnTenKey.Text = "Ten Key On"
                BtnTenKey.BackColor = System.Drawing.Color.Pink
            Else
                gbTenKeyFlg = False
                BtnTenKey.Text = "Ten Key Off"
                BtnTenKey.BackColor = System.Drawing.SystemColors.Control
            End If

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "FrmFineAdjust.SubBtnTenKey_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

    '----- ###237↓ -----
#Region "LASERボタン押下時処理"
    '''=========================================================================
    ''' <summary>LASERボタン押下時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnLaser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnLaser.Click

        Dim r As Integer
        Dim strMSG As String

        Try
            ' LASER射出可能/不可の切り替え
            If (BtnLaser.BackColor = System.Drawing.SystemColors.Control) Then
                ' LASER射出可能とする
                BtnLaser.BackColor = System.Drawing.Color.OrangeRed
            Else
                ' LASER射出不可とする
                BtnLaser.BackColor = System.Drawing.SystemColors.Control
                r = LASEROFF()
                m_LaserOnOffFlag = False
            End If

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "frmFineAdjust.BtnLaser_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region
    '----- ###237↑ -----

    '========================================================================================
    '   共通関数
    '========================================================================================
#Region "ステージ移動処理"
    '''=========================================================================
    ''' <summary>ステージ移動処理</summary>
    ''' <param name="pltNo"></param>
    ''' <param name="blkNo"></param>
    ''' <returns></returns>
    '''=========================================================================
    Private Function MoveTargetStagePos(ByVal pltNo As Integer, ByVal blkNo As Integer) As Integer

        Dim intRet As Integer
        Dim nextStgX As Double
        Dim nextStgY As Double
        Dim dispPltX As Integer
        Dim dispPltY As Integer
        Dim dispBlkX As Integer
        Dim dispBlkY As Integer
        'Dim retBlkNoX As Integer
        'Dim retBlkNoY As Integer
        Dim dispCurStgGrpNoX As Integer
        Dim dispCurStgGrpNoY As Integer
        Dim dispCurBlkNoX As Integer
        Dim dispCurBlkNoY As Integer
        Dim dispCurPltNoX As Integer
        Dim dispCurPltNoY As Integer

        Try
            MoveTargetStagePos = MOVE_NEXT
            intRet = GetTargetStagePos(pltNo, blkNo, nextStgX, nextStgY, dispPltX, dispPltY, dispBlkX, dispBlkY)
            If intRet = BLOCK_END Then
                ' 何もしないで終了
                MoveTargetStagePos = MOVE_NOT
                Exit Function
            ElseIf intRet = PLATE_BLOCK_END Then
                ' 何もしないで終了
                MoveTargetStagePos = MOVE_NOT
                Exit Function
            End If

            '---------------------------------------------------------------------
            '   表示用各ポジションの番号を設定（プレート/ステージグループ/ブロック）
            '---------------------------------------------------------------------
            Dim bRet As Boolean
            bRet = GetDisplayPosInfo(dispBlkX, dispBlkY, _
                            dispCurStgGrpNoX, dispCurStgGrpNoY, dispCurBlkNoX, dispCurBlkNoY)

            '---------------------------------------------------------------------
            '   ログ表示文字列の設定
            '---------------------------------------------------------------------
            dispCurPltNoX = dispPltX : dispCurPltNoY = dispPltY         '###056
            Call DisplayStartLog(dispCurPltNoX, dispCurPltNoY, _
                            dispCurStgGrpNoX, dispCurStgGrpNoY, dispCurBlkNoX, dispCurBlkNoY)
            '' '' '' ステージの動作
            ' '' ''intRet = Form1.System1.EX_START(gSysPrm, nextStgX + typPlateInfo.dblTableOffsetXDir + gfCorrectPosX, _
            ' '' ''                        nextStgY + typPlateInfo.dblTableOffsetYDir + gfCorrectPosY, 0)
        Catch ex As Exception
            Dim strMsg As String
            strMsg = "frmFineAdjust.btnTrimming_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMsg)
        End Try
    End Function
#End Region

#Region "矢印ボタン活性化/非活性化"
    '''=========================================================================
    ''' <summary>矢印ボタン活性化/非活性化 ###139</summary>
    ''' <param name="OnOff"></param>
    '''=========================================================================
    Private Sub SetBtnArowEnable(ByVal OnOff As Boolean)

        Dim strMSG As String

        Try
            ' 矢印ボタン活性化/非活性化
            BtnJOG_0.Enabled = OnOff
            BtnJOG_1.Enabled = OnOff
            BtnJOG_2.Enabled = OnOff
            BtnJOG_3.Enabled = OnOff
            BtnJOG_4.Enabled = OnOff
            BtnJOG_5.Enabled = OnOff
            BtnJOG_6.Enabled = OnOff
            BtnJOG_7.Enabled = OnOff
            BtnHI.Enabled = OnOff

            ' Ten Keyボタン活性化/非活性化
            BtnTenKey.Enabled = OnOff

            ' Ten KeyボタンをOn/Offにする
            If (OnOff = False) Then
                ' 矢印ボタン非活性化ならTen KeyボタンをOffにしてテンキー入力を不可とする
                If (BtnTenKey.Text = "Ten Key On") Then
                    m_TenKeyFlg = True
                    Call SubBtnTenKey_Click()
                End If
            Else
                ' Ten KeyボタンをOffにした場合はTen KeyボタンをOnにしてテンキー入力を可とする
                If (m_TenKeyFlg = True) Then
                    m_TenKeyFlg = False
                    Call SubBtnTenKey_Click()
                End If
            End If

        Catch ex As Exception
            strMSG = "frmFineAdjust.SetBtnArowEnable() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

    '===============================================================================
    '   Description  : ＪＯＧ操作画面処理
    '
    '   Copyright(C) : TOWA LASERFRONT CORP. 2012
    '
    '===============================================================================
    '========================================================================================
    '   ボタン押下時処理
    '========================================================================================
#Region "RESETボタン押下時処理"
    '''=========================================================================
    ''' <summary>RESETボタン押下時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnRESET_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnRESET.Click
        mExit_flg = cFRS_ERR_RST                                        ' Return値 = Cancel(RESETｷｰ)  
        gbExitFlg = True                                                ' 終了フラグON
        Me.Close()
    End Sub
#End Region

#Region "HIボタン押下時処理"
    '''=========================================================================
    ''' <summary>HIボタン押下時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnHI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnHI.Click
        Call SubBtnHI_Click(stJOG)
    End Sub
#End Region

#Region "矢印ボタンのマウスクリック時処理"
    '''=========================================================================
    ''' <summary>矢印ボタンのマウスクリック時処理</summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub BtnJOG_0_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_0.MouseDown
        Call SubBtnJOG_0_MouseDown()                                    ' +Y ON
    End Sub
    Private Sub BtnJOG_0_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_0.MouseUp
        Call SubBtnJOG_0_MouseUp()                                      ' +Y OFF
    End Sub

    Private Sub BtnJOG_1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_1.MouseDown
        Call SubBtnJOG_1_MouseDown()                                    ' -Y ON
    End Sub
    Private Sub BtnJOG_1_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_1.MouseUp
        Call SubBtnJOG_1_MouseUp()                                      ' -Y OFF
    End Sub

    Private Sub BtnJOG_2_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_2.MouseDown
        Call SubBtnJOG_2_MouseDown()                                    ' +X ON
    End Sub
    Private Sub BtnJOG_2_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_2.MouseUp
        Call SubBtnJOG_2_MouseUp()                                      ' +X OFF
    End Sub

    Private Sub BtnJOG_3_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_3.MouseDown
        Call SubBtnJOG_3_MouseDown()                                    ' -X ON
    End Sub
    Private Sub BtnJOG_3_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_3.MouseUp
        Call SubBtnJOG_3_MouseUp()                                      ' -X OFF
    End Sub

    Private Sub BtnJOG_4_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_4.MouseDown
        Call SubBtnJOG_4_MouseDown()                                    ' -X -Y ON
    End Sub
    Private Sub BtnJOG_4_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_4.MouseUp
        Call SubBtnJOG_4_MouseUp()                                      ' -X -Y OFF
    End Sub

    Private Sub BtnJOG_5_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_5.MouseDown
        Call SubBtnJOG_5_MouseDown()                                    ' +X -Y ON
    End Sub
    Private Sub BtnJOG_5_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_5.MouseUp
        Call SubBtnJOG_5_MouseUp()                                      ' +X -Y OFF
    End Sub

    Private Sub BtnJOG_6_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_6.MouseDown
        Call SubBtnJOG_6_MouseDown()                                    ' +X +Y ON
    End Sub
    Private Sub BtnJOG_6_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_6.MouseUp
        Call SubBtnJOG_6_MouseUp()                                      ' +X +Y OFF
    End Sub

    Private Sub BtnJOG_7_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_7.MouseDown
        Call SubBtnJOG_7_MouseDown()                                    ' -X +Y ON
    End Sub
    Private Sub BtnJOG_7_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles BtnJOG_7.MouseUp
        Call SubBtnJOG_7_MouseUp()                                      ' -X +Y OFF
    End Sub
#End Region
    '----- ###219 -----
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
    '----- ###219 -----

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

        'Dim KeyCode As Short = e.KeyCode

        ''----- ###237↓ -----
        '' LASER射出可能で「*キー」押下ならLASER射出する()
        'If (BtnLaser.BackColor = System.Drawing.Color.OrangeRed) And (KeyCode = System.Windows.Forms.Keys.Multiply) Then
        '    ' レーザON
        '    If (m_LaserOnOffFlag = False) Then
        '        Call LASERON()
        '        m_LaserOnOffFlag = True
        '        Console.WriteLine("frmFineAdjust_KeyDown() Laser On")
        '    End If
        'End If
        ''----- ###237↑ -----

        '' テンキー入力フラグがOFFならNOP ###057
        'If (gbTenKeyFlg = False) Then
        '    e.Handled = False           ' V6.1.3.0⑨
        '    Exit Sub
        'End If

        '' テンキーダウンならInpKeyにテンキーコードを設定する
        'Call Sub_10KeyDown(KeyCode)
        'If (KeyCode = System.Windows.Forms.Keys.NumPad5) Then       ' 5ｷｰ (KeyCode = 101(&H65)
        '    Call BtnHI_Click(sender, e)                             ' HIボタン ON/OFF
        'End If
        ''Call Me.Focus()

    End Sub

    Public Sub JogKeyDown(ByVal e As KeyEventArgs) Implements ICommonMethods.JogKeyDown    'V2.2.0.0①

        Dim KeyCode As Short = e.KeyCode

        '----- ###237↓ -----
        ' LASER射出可能で「*キー」押下ならLASER射出する()
        If (BtnLaser.BackColor = System.Drawing.Color.OrangeRed) And (KeyCode = System.Windows.Forms.Keys.Multiply) Then
            ' レーザON
            If (m_LaserOnOffFlag = False) Then
                Call LASERON()
                m_LaserOnOffFlag = True
                Console.WriteLine("frmFineAdjust_KeyDown() Laser On")
            End If
        End If
        '----- ###237↑ -----

        ' テンキー入力フラグがOFFならNOP ###057
        If (gbTenKeyFlg = False) Then
            e.Handled = False           ' V6.1.3.0⑨
            Exit Sub
        End If

        ' テンキーダウンならInpKeyにテンキーコードを設定する
        Call Sub_10KeyDown(KeyCode)
        If (KeyCode = System.Windows.Forms.Keys.NumPad5) Then       ' 5ｷｰ (KeyCode = 101(&H65)
            Call BtnHI_Click(BtnHI, e)                             ' HIボタン ON/OFF
        End If
        'Call Me.Focus()



        ''V6.0.0.0⑪        Dim KeyCode As Short = e.KeyCode
        'Dim KeyCode As Keys = e.KeyCode             'V6.0.0.0⑪
        'Dim r As Integer

        ''----- ###237↓ -----
        '' LASER射出可能で「*キー」押下ならLASER射出する()
        'If (BtnLaser.BackColor = System.Drawing.Color.OrangeRed) And (KeyCode = System.Windows.Forms.Keys.Multiply) Then
        '    ' レーザON
        '    If (m_LaserOnOffFlag = False) Then
        '        ' DIG-SW設定
        '        Call Form1.SetMoveMode(CbDigSwL.SelectedIndex, CbDigSwH.SelectedIndex) 'V5.0.0.1⑫

        '        ''V4.0.0.0-86
        '        r = GetLaserOffIO(False)
        '        If r = 1 Then
        '            Me.ShowInTaskbar = False 'V5.0.0.1⑫
        '            Me.Activate()  'V5.0.0.1⑫
        '            'frmFineAdjust_KeyUp(sender, e)

        '            Exit Sub
        '        End If
        '        ''V4.0.0.0-86
        '        Call LASERON()
        '        m_LaserOnOffFlag = True
        '        Console.WriteLine("frmFineAdjust_KeyDown() Laser On")
        '    End If
        'End If
        ''----- ###237↑ -----

        '' テンキー入力フラグがOFFならNOP ###057
        ''V7.0.0.0⑮        If (gbTenKeyFlg = False) Then Exit Sub
        'If (gbTenKeyFlg = False) OrElse (False = _firstResistor) Then   'V7.0.0.0⑮
        '    e.Handled = False           ' V6.1.3.0⑨
        '    Exit Sub
        'End If

        '' テンキーダウンならInpKeyにテンキーコードを設定する
        ''V6.0.0.0⑫       'Call Sub_10KeyDown(KeyCode)
        'Sub_10KeyDown(KeyCode, stJOG)             'V6.0.0.0⑫
        'If (KeyCode = System.Windows.Forms.Keys.NumPad5) Then       ' 5ｷｰ (KeyCode = 101(&H65)
        '    'Call BtnHI_Click(sender, e)                             ' HIボタン ON/OFF
        '    Call BtnHI_Click(BtnHI, e)                              ' HIボタン ON/OFF     'V6.0.0.0⑩
        'End If
        ''Call Me.Focus()

    End Sub
#End Region

#Region "キーアップ時処理"
    '''=========================================================================
    ''' <summary>キーアップ時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub frmFineAdjust_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp

        Me.JogKeyUp(e)                  'V6.0.0.0⑪

        'Dim KeyCode As Short = e.KeyCode

        ''----- ###237↓ -----
        '' LASER Offする
        'If (m_LaserOnOffFlag = True) Then
        '    Call LASEROFF()
        '    m_LaserOnOffFlag = False
        '    Console.WriteLine("frmFineAdjust_KeyUp() Laser Off")
        'End If
        ''----- ###237↑ -----

        '' テンキー入力フラグがOFFならNOP ###057
        'If (gbTenKeyFlg = False) Then Exit Sub

        '' テンキーアップならInpKeyのテンキーコードをOFFする
        'Call Sub_10KeyUp(KeyCode)
        ''Call Me.Focus()

    End Sub


    Public Sub JogKeyUp(ByVal e As KeyEventArgs) Implements ICommonMethods.JogKeyUp        'V2.2.0.0①

        Dim KeyCode As Short = e.KeyCode

        '----- ###237↓ -----
        ' LASER Offする
        If (m_LaserOnOffFlag = True) Then
            Call LASEROFF()
            m_LaserOnOffFlag = False
            Console.WriteLine("frmFineAdjust_KeyUp() Laser Off")
        End If
        '----- ###237↑ -----

        ' テンキー入力フラグがOFFならNOP ###057
        If (gbTenKeyFlg = False) Then Exit Sub

        ' テンキーアップならInpKeyのテンキーコードをOFFする
        Call Sub_10KeyUp(KeyCode)
        'Call Me.Focus()

        ''V6.0.0.0⑪        Dim KeyCode As Short = e.KeyCode
        'Dim KeyCode As Keys = e.KeyCode             'V6.0.0.0⑪

        ''----- ###237↓ -----
        '' LASER Offする
        'If (m_LaserOnOffFlag = True) Then
        '    Call LASEROFF()
        '    m_LaserOnOffFlag = False
        '    Console.WriteLine("frmFineAdjust_KeyUp() Laser Off")
        'End If
        ''----- ###237↑ -----

        '' テンキー入力フラグがOFFならNOP ###057
        ''V6.0.1.0③        If (gbTenKeyFlg = False) Then Exit Sub
        ''V7.0.0.0⑮        If (False = gbTenKeyFlg) Then       'V6.0.1.0③
        'If (False = gbTenKeyFlg) OrElse (False = _firstResistor) Then   'V7.0.0.0⑮
        '    'V6.1.3.0⑨
        '    If (giBpAdj_HALT = 0) Then
        '        Sub_10KeyUp(Keys.None, stJOG)   'V6.0.1.0③
        '    End If
        '    'V6.1.3.0⑨
        'Else
        '    ' テンキーアップならInpKeyのテンキーコードをOFFする
        '    'V6.0.0.0⑫        Call Sub_10KeyUp(KeyCode)
        '    Sub_10KeyUp(KeyCode, stJOG)                   'V6.0.0.0⑫
        '    'Call Me.Focus()
        'End If

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

    ''' <summary>
    ''' 一時停止画面でキャプチャー画面をクリックしたときに、動作する、しないの設定       'V2.2.0.0④
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnClickEnable_Click(sender As Object, e As EventArgs) Handles BtnClickEnable.Click

        Try

            If BtnClickEnable.BackColor = SystemColors.Control Then
                BtnClickEnable.BackColor = Color.Aqua
                gbTenKeyFlg = True
            Else
                BtnClickEnable.BackColor = SystemColors.Control
                gbTenKeyFlg = False
            End If

        Catch ex As Exception

        End Try

    End Sub
#End Region

#Region "メイン処理実行"
    ''' <summary>メイン処理実行</summary>
    ''' <returns>実行結果</returns>
    ''' <remarks>'V2.2.0.0①</remarks>
    Public Function Execute() As Integer Implements ICommonMethods.Execute
        ' DO NOTHING
    End Function
#End Region

#Region "カメラ画像クリック位置を画像センターに移動する処理"
    ''' <summary>カメラ画像クリック位置を画像センターに移動する処理</summary>
    ''' <param name="distanceX">画像センターからの距離X</param>
    ''' <param name="distanceY">画像センターからの距離Y</param>
    ''' <remarks>'V6.0.0.0⑪</remarks>
    Public Sub MoveToCenter(ByVal distanceX As Decimal, ByVal distanceY As Decimal) _
        Implements ICommonMethods.MoveToCenter

        ' テンキー入力フラグがOFFならNOP 
        If (gbTenKeyFlg = False) Then
            Exit Sub
        End If

        UserModule.MoveToCenter(distanceX, distanceY, stJOG)

    End Sub

    ''' <summary>
    ''' ローダ情報表示ボタン 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnLoaderInfo_Click(sender As Object, e As EventArgs) Handles BtnLoaderInfo.Click

        Try

            objLoaderInfo.Show()

        Catch ex As Exception

        End Try



    End Sub

    ''' <summary>
    ''' 自動運転の登録されているファイルを表示し、終了、実行中、未処理をわかるようにする 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnAutoInfo_Click(sender As Object, e As EventArgs) Handles btnAutoInfo.Click

        Dim objForm As frmLotFileDisp

        Try

            objForm = New frmLotFileDisp

            objForm.Show(Me)



        Catch ex As Exception

        End Try
    End Sub
#End Region

End Class