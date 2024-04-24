'==============================================================================
'   Description : メイン画面処理
'
'   Copyright(C): TOWA LASERFRONT CORP. 2018
'
'==============================================================================
Option Strict Off
Option Explicit On

Imports System.Text                     ' ###lstLog
Imports System.Globalization
Imports System.Threading
Imports Microsoft.Win32
Imports LaserFront.Trimmer
Imports System.Collections.Generic
Imports LaserFront.Trimmer.DefWin32Fnc
Imports UsrFunc.My.Resources
Imports DllPlcIf

Friend Class Form1
    Inherits System.Windows.Forms.Form

    '=========================================================================
    '   変数定義
    '=========================================================================
#Region "フォーム内グローバル変数"
    '=====================================================================
    ' フォーム内グローバル変数
    '=====================================================================
    Private Const CUR_CRS_LINEX As Short = 8            ' ｸﾛｽﾗｲﾝX表示位置の補正値
    Private Const CUR_CRS_LINEY As Short = 13           ' ｸﾛｽﾗｲﾝY表示位置の補正値

    Private gflgCmpEndProcess As Boolean                ' 終了処理完了フラグ（True=終了処理実行済み、False=終了処理実行済みでない）
    Private gfclamp As Boolean                          ' クランプON/OFF
    Private pbVideoCapture As Boolean                   ' ビデオキャプチャー開始フラグ
    Public gPrevTrimMode As Short                       ' デジタルＳＷ退避域
    Public giTrimErr As Short                           ' ﾄﾘﾏｰ ｴﾗｰ ﾌﾗｸﾞ ※ｴﾗｰ時はｸﾗﾝﾌﾟｸﾗﾝﾌﾟOFF時ﾄﾘﾏ動作中OFFをﾛｰﾀﾞｰに送信しない
    '                                                   ' B0 : 吸着ｴﾗｰ(EXIT)
    '                                                   ' B1 : その他ｴﾗｰ
    '                                                   ' B2 : 集塵機ｱﾗｰﾑ検出
    '                                                   ' B3 : 軸ﾘﾐｯﾄ､軸ｴﾗｰ､軸ﾀｲﾑｱｳﾄ
    '                                                   ' B4 : 非常停止
    '                                                   ' B5 : ｴｱｰ圧ｴﾗｰ

    Private pbVideoInit As Boolean                      ' ビデオInitフラグ
    'Private Const cTEMPLATPATH As String = "C:\TRIM\VIDEO"  ' Video.OCX用ﾃﾝﾌﾟﾚｰﾄﾌｧｲﾙの保存場所
    Private Const WORK_DIR_PATH As String = "C:\TRIM"       ' 作業用ﾌｫﾙﾀﾞｰ
    Private gbChkboxHalt As Boolean = False             ' ADJボタン状態(ON=ADJ ON, OFF=ADJ OFF) ###009
    Private gbAdjOnStatus As Boolean = False            ' ＡＤＪボタンでの停止中

#End Region

#Region "カメラ画像クリック移動関連"     'V2.2.0.0①
    Private _jogKeyDown As Action(Of KeyEventArgs) = Nothing
    Private _jogKeyUp As Action(Of KeyEventArgs) = Nothing
    ''' <summary><para>表示中のJOGを制御するKeyDown,KeyUp時の処理をメインフォームに、</para>
    ''' <para>カメラ画像MouseClick時の処理をDllVideoに設定する</para></summary>
    ''' <param name="keyDown"></param>
    ''' <param name="keyUp"></param>
    ''' <param name="moveToCenter">カメラ画像クリック位置を画像センターに移動する処理</param>
    Friend Sub SetActiveJogMethod(ByVal keyDown As Action(Of KeyEventArgs),
                                  ByVal keyUp As Action(Of KeyEventArgs),
                                  ByVal moveToCenter As Action(Of Decimal, Decimal))
        _jogKeyDown = keyDown
        _jogKeyUp = keyUp

        'カメラ画像表示PictureBoxクリック位置をJOG経由で画像センターに移動する
        VideoLibrary1.MoveToCenter = moveToCenter
    End Sub
#End Region

    '=========================================================================
    '   フォームの初期化/終了処理
    '=========================================================================
#Region "シャットダウン処理-強制終了"
    '''=========================================================================
    '''<summary>シャットダウン処理-強制終了</summary>
    '''<param name="sender"></param> 
    '''<param name="e"></param> 
    '''=========================================================================
    Private Sub SystemEvents_SessionEnding(
            ByVal sender As Object,
            ByVal e As SessionEndingEventArgs)
        If e.Reason = SessionEndReasons.SystemShutdown Then
            Call AplicationForcedEnding()
        End If
    End Sub
#End Region

#Region "フォーム初期化処理"
    '''=========================================================================
    '''<summary>フォーム初期化処理</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub Form_Initialize_Renamed()

        Dim r As Short
        Dim strMSG As String

        Try
            '-----------------------------------------------------------------------
            '   多重起動防止Mutexハンドル
            '-----------------------------------------------------------------------
            If gmhUserPro.WaitOne(0, False) = False Then
                '' すでに起動されている場合
                '   →メッセージボックスがＳＴＡＲＴボタン入力待ちなどの状態で、後ろに回ることがあるので、表示はやめる。
                'MessageBox.Show("Cannot run TKY's family.(Another Process of TKY's family is already running.", "Trimmer Program", _
                '                MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, _
                '                MessageBoxOptions.ServiceNotification, False)
                End
            End If




            ' シャットダウンイベント処理関数
            AddHandler SystemEvents.SessionEnding, AddressOf SystemEvents_SessionEnding

            ChDir(WORK_DIR_PATH)
            Timer1.Enabled = False                                      ' 監視タイマー停止

            ' Intime動作確認
#If cOFFLINEcDEBUG = 0 Then
            r = ISALIVE_INTIME()
            If (r = ERR_INTIME_NOTMOVE) Then
                'エラーメッセージの表示 (System1.TrmMsgBoxはここでは使用できない為、標準メッセージボックス)
                MessageBox.Show("Real-time control module has not loaded.", "Trimmer Program", MessageBoxButtons.OK,
                                MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.DefaultDesktopOnly, False)
                End                                                     ' アプリ終了 
            End If
#End If
            '----------------------------------------------------------------------------
            '   フラグ等初期化
            '----------------------------------------------------------------------------
            ' フラグ初期化
            gbInitialized = False
            pbVideoInit = False
            pbVideoCapture = False                                      ' ビデオキャプチャー開始フラグ
            pbLoadFlg = False                                           ' データロード済みフラグ
            gflgResetStart = False                                      ' 初期化フラグ
            gfclamp = False                                             ' クランプOFF
            giAppMode = APP_MODE_IDLE                                   ' ｱﾌﾟﾘﾓｰﾄﾞ = トリマ装置アイドル中
            FlgGPIB = False                                                 ' GPIB初期化Flag
            FlgUpd = False                                              ' データ更新 Flag
            giTrimErr = 0                                               ' ﾄﾘﾏｰ ｴﾗｰ ﾌﾗｸﾞ初期化
            fStartTrim = False                                          ' スタートTRIMフラグ OFF
            gflgCmpEndProcess = False                                   ' 終了処理完了フラグ
            '                                                           ' ディジタルSW初期化 
            DGH = DGSW_HI_DISP                                          ' ディジタルSWH = 全て表示
            DGL = TRIM_MODE_ITTRFT                                      ' ディジタルSWL = イニシャルテスト＋トリミング＋ファイナルテスト実行
            DGSW = DGH * 10 + DGL                                       ' ディジタルSW

            ' 構造体の初期化
            Call Init_Struct()

            '----------------------------------------------------------------------------
            '   使用するＯＣＸの初期設定を行う
            '----------------------------------------------------------------------------
            Call Ocx_Initialize()                                       ' ｼｽﾃﾑﾊﾟﾗﾒｰﾀのﾘｰﾄﾞ前で行う
            '                                                           ' ｼｽﾊﾟﾗREAD前にForm_Load()に制御が渡るので注意
            '----------------------------------------------------------------------------
            '   システム設定ファイルリード
            '   ※システムパラメータの送信はOcxSystemのSetOptionFlg()で行う
            '----------------------------------------------------------------------------
            gSysPrm.Initialize()
            Call DllSysPrmSysParam_definst.SetAppKind(KND_USER)
            Call DllSysPrmSysParam_definst.GetSystemParameter(gSysPrm)   ' システム設定ファイルリード
            Call PrepareMessages(gSysPrm.stTMN.giMsgTyp)                 ' メッセージ初期設定処理
            Call Me.System1.OperationLogDelete(gSysPrm)                  ' 古い操作ログファイルを削除する
            Call Me.System1.OperationLogging(gSysPrm, MSG_OPLOG_START, "")
            Call Me.System1.SetSysParam(gSysPrm)                         ' OcxSystem用のシステムパラメータを設定する

            ' ログ画面表示クリア基板枚数をシスパラより取得する
            gDspClsCount = GetPrivateProfileInt("SPECIALFUNCTION", "DISP_CLS_USR", 5, SYSPARAMPATH)
            If (gDspClsCount <= 0) Then gDspClsCount = 1
            gDspCounter = 0                                             ' ログ画面表示基板枚数カウンタ

            ' EXTOUT LED制御ビット(BIT4-7)をシスパラより設定する
            glLedBit = Val(GetPrivateProfileString_S("IO_CONTROL", "ILUM_BIT", SYSPARAMPATH, "16"))

            gGpibMultiMeterCount = GetPrivateProfileInt("SPECIALFUNCTION", "MULTIMETER_COUNT", 5, SYSPARAMPATH) ' マルチメータのITと測定の測定回数最後の値を使用する。
            bDebugLogOut = GetPrivateProfileInt("LOGGING", "LOG_DEBUG", 0, SYSPARAMPATH)                '通常のデバッグログの出力有無       'V1.2.0.2
            bNgCutDebugLogOut = GetPrivateProfileInt("LOGGING", "LOG_DEBUG_NGCUT", 0, SYSPARAMPATH)     'ＮＧカット用デバッグログの出力有無 'V1.2.0.2
            bCutVariationDebugLogOut = GetPrivateProfileInt("LOGGING", "LOG_DEBUG_CUTVA", 0, SYSPARAMPATH)  'ＮＧカット用デバッグログの出力有無 'V2.1.0.0①
            'V2.0.0.0⑬↓
            If Integer.Parse(GetPrivateProfileString_S("USER", "RELAY_BOARD", SYSPARAMPATH, "0")) = 2 Then
                bRelayBoard = True
            Else
                bRelayBoard = False
            End If
            'V2.0.0.0⑬↑
            ''V2.2.1.3②↓
            giAutoOperationDebugLogOut = GetPrivateProfileInt("LOGGING", "LOG_DEBUG_AUTOMODE", 0, SYSPARAMPATH)  'ロット処理用デバッグログの出力有無 
            ''V2.2.1.3②

            ''V2.0.0.0②↓
            If Integer.Parse(GetPrivateProfileString_S("USER", "POWER_ON_OFF_TRIM_MEAS", SYSPARAMPATH, "0")) <> 0 Then
                bPowerOnOffUse = True
            Else
                bPowerOnOffUse = False
            End If
            ''V2.0.0.0②↑
            ''V2.2.0.0③↓
            If Integer.Parse(GetPrivateProfileString_S("OPT_TEACH", "DISABLE_BLUE_CROSSLINE", SYSPARAMPATH, "0")) <> 0 Then
                giBlueCrossDisable = 1
            Else
                giBlueCrossDisable = 0
            End If
            ''V2.2.0.0③↑
            ''V2.2.0.0④↓
            If Integer.Parse(GetPrivateProfileString_S("SPECIALFUNCTION", "ADJ_MOUSECLICK_DISABLE", SYSPARAMPATH, "0")) <> 0 Then
                giMouseClickMove = 1
            Else
                giMouseClickMove = 0
            End If
            ''V2.2.0.0④↑
            ''V2.2.0.0⑤↓
            If Integer.Parse(GetPrivateProfileString_S("DEVICE_CONST", "LOADER_TYPE", SYSPARAMPATH, "0")) <> 0 Then
                giLoaderType = 1
                btnLoaderInfo.Visible = True
                Call COVERCHK_ONOFF(0)                         ' 「固定カバー開チェックあり」にする
                btnCycleStop.Visible = True         'V2.2.2.0①
            Else
                giLoaderType = 0
                btnLoaderInfo.Visible = False
                btnCycleStop.Visible = False        'V2.2.2.0①
            End If
            ''V2.2.0.0⑤↑


            ''V2.2.0.0⑥↓
            If Integer.Parse(GetPrivateProfileString_S("SPECIALFUNCTION", "CUT_STOP", SYSPARAMPATH, "0")) <> 0 Then
                giCutStop = 1
                Me.btnCutStop.Visible = True
            Else
                giCutStop = 0
                Me.btnCutStop.Visible = False
            End If
            ''V2.2.0.0⑥↑

            ''V2.2.0.0⑦↓
            ' サイクル停止機能
            If Integer.Parse(GetPrivateProfileString_S("SPECIALFUNCTION", "CYCLE_STOP", SYSPARAMPATH, "0")) <> 0 Then
                giClcleStop = 1
                Me.btnCycleStop.Visible = True
            Else
                giClcleStop = 0
                Me.btnCycleStop.Visible = False
            End If
            ''V2.2.0.0⑦↑


            'V2.2.2.0①↓
            ' 内部カメラ番号取得
            INTERNAL_CAMERA = Integer.Parse(GetPrivateProfileString_S("OPT_VIDEO", "INTERNAL_CAMERA_PORT", SYSPARAMPATH, "0"))
            ' 外部カメラ番号取得
            EXTERNAL_CAMERA = Integer.Parse(GetPrivateProfileString_S("OPT_VIDEO", "EXTERNAL_CAMERA_PORT", SYSPARAMPATH, "1"))
            'V2.2.2.0①↑

            '----------------------------------------------------------------------------
            ' ユーザー定義変数の初期化処理
            '----------------------------------------------------------------------------
            Call Set_UserForm(Z0)                                       ' メイン画面
            Call Me.System1.SetSignalTower(0, &HFFFFS)                  ' ｼｸﾞﾅﾙﾀﾜｰ初期化(On,Off)
            Call GetFncDefParameter()                                   ' 機能選択定義テーブル設定
            Call GetPasFuncDefParameter()                               ' パスワード定義テーブル設定

            '----------------------------------------------------------------------------
            '   オブジェクト設定
            '----------------------------------------------------------------------------
            ObjGpib = New GpibMaster                                    ' ＧＰＩＢ通信用オブジェクト
            frmAutoObj = New FormDataSelect(Me)                             ' 自動運転処理ｵﾌﾞｼﾞｪｸﾄ

            '----------------------------------------------------------------------------
            '   画面表示項目を設定し画面に表示する
            '----------------------------------------------------------------------------
            gPrevTrimMode = -1

            ' form関連(Form Loadが走る)
            Me.Picture1.Top = gSysPrm.stDVR.giCrossLineX + 59            ' CLOSS LINE X(横線)
            Me.Picture2.Left = gSysPrm.stDVR.giCrossLineY + 8            ' CLOSS LINE Y(縦線)

            ' OK/NG表示ｸﾘｱ
            Call Disp_Result(0, 0)

            ' ﾛｸﾞ画面拡大表示切替(ｵﾌﾟｼｮﾝ機能)
            If (gSysPrm.stSPF.giDispCh = 0) Then                         ' 拡大表示しない ?
                cmdExpansion.Visible = False
            Else
                cmdExpansion.Visible = True
            End If

            Dim ctxStr As String
            If (0 = gSysPrm.stTMN.giMsgTyp) Then
                ctxStr = "コピー (&C)"
            Else
                ctxStr = "Copy (&C)"
            End If

            Me.ctxMenuLstBox.Items.Add(
                ctxStr, Nothing, New EventHandler(AddressOf lstLog_Copy))       ' ###lstLog
            Me.lstLog.Items.Add(" ")
            Me.txtlog.Text = " "            ''V2.2.0.0⑰↑

            ' レーザパワー調整関連項目の設定(日本語/英語)
            SetLaserItems()

            ' クランプ/吸着OFF
            r = Me.System1.ClampVacume_Ctrl(gSysPrm, 0, giAppMode, giTrimErr, False)
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
                Call AplicationForcedEnding()                           ' ｿﾌﾄ強制終了処理
                End                                                     ' アプリ強制終了
                Return
            End If

            ''V2.2.0.0⑤↓
            ' TLF製ローダの場合自動運転切り替えを出力する
            If giLoaderType = 1 Then
                Call Me.System1.Z_ATLDSET(0, clsLoaderIf.LOUT_AUTO Or clsLoaderIf.LOUT_REDY)                    'V1.2.0.0④ ローダー出力(ON=自動,OFF=なし)
            End If
            ''V2.2.0.0⑤↑

            ''V2.2.0.028 ↓
            giTablePosUpd = Int32.Parse(GetPrivateProfileString_S("OPT_VIDEO", "TABLE_POS_UPDATE", "C:\TRIM\tky.ini", "0"))
            ''V2.2.0.028 ↑

            giRecogPointCorrLine = Int16.Parse(GetPrivateProfileString_S("OPT_VIDEO", "CUTPOSCORR_BASELINE", "C:\TRIM\tky.ini", "0"))   ' V2.2.1.2①

            'V2.2.0.030↓
            giLaserOffMode = 0
            If (gSysPrm.stRAT.giOsc_Res = OSCILLATOR_SP) Then
                btnLaserOff.Visible = True
                btnLaserOff.Enabled = True
            Else
                btnLaserOff.Visible = False
                btnLaserOff.Enabled = False
            End If
            'V2.2.0.030↑

            '---------------------------------------------------------------------------
            '   起動後の最初の検出がﾛｰﾀﾞ自動ﾓｰﾄﾞ/動作中の場合は、停止に切替えるよう確認する
            '---------------------------------------------------------------------------
            ' ローダ入力
            giHostMode = cHOSTcMODEcMANUAL                              ' ﾛｰﾀﾞﾓｰﾄﾞ = 手動ﾓｰﾄﾞ
            gbHostConnected = False                                     ' ホスト接続状態 = 未接続(ﾛｰﾀﾞ無)
            giHostRun = 0                                               ' ﾛｰﾀﾞ停止中
            Call Me.System1.ReadHostCommand(gSysPrm, giHostMode, gbHostConnected, giHostRun, giAppMode, pbLoadFlg)

            ' 起動時ﾛｰﾀﾞ自動ﾓｰﾄﾞ/動作中ﾁｪｯｸ
            r = Me.System1.Form_Reset(cGMODE_LDR_CHK, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
            If (r <> cFRS_NORMAL) Then                                  ' エラー(非常停止) ?
                'Call AplicationForcedEnding()                           ' ｿﾌﾄ強制終了処理
                End                                                     ' アプリ強制終了
                Return
            End If

            stCounter.LotCounter = 0                            ' ロットカウンター初期化

            ' クロスライン補正の初期化
            ' ObjCrossLine.CrossLineParamINitial(Me.Picture2, Me.Picture1, Me.CrosLineX, Me.CrosLineY, 0.0, 0.0)
            ObjCrossLine.CrossLineParamINitial(AddressOf VideoLibrary1.GetCrossLineCenter,
                                               AddressOf VideoLibrary1.SetCorrCrossVisible,
                                               AddressOf VideoLibrary1.SetCorrCrossCenter,
                                               0.0, 0.0)
            Me.CrosLineX.BringToFront()
            Me.CrosLineY.BringToFront()

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Form_Initialize() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "ＯＣＸの初期設定"
    '''=========================================================================
    '''<summary>使用するＯＣＸの初期設定を行う</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub Ocx_Initialize()

        Dim strMSG As String

        Try
            Dim i As Short
            Dim r As Short
            Dim onoff(cMAXOptFlgNUM) As Short           ' ｺﾝﾊﾟｲﾙｵﾌﾟｼｮﾝ(最大数)

            '---------------------------------------------------------------------------
            '   OCX用オブジェクトを設定する
            '---------------------------------------------------------------------------
            ObjMain = Me                                ' Form1クラス
            ObjSys = System1                            ' OcxSystem.ocx        '★★★
            ObjUtl = Utility1                           ' OcxUtility.ocx
            ObjHlp = HelpVersion1                       ' OcxAbout.ocx
            ObjPas = Password1                          ' OcxPassword.ocx
            ObjMTC = ManualTeach1                       ' OcxManualTeach.ocx
            ObjTch = Teaching1                          ' Teach.ocx
            ObjPrb = Probe1                             ' Probe.ocx
            ObjVdo = VideoLibrary1                      ' Video.ocx
            ObjLoader = New clsLoaderIf()                    ' Loaderクラス     'V2.2.0.0⑤
            '@@888 ObjfrmResetLoader = New frmResetLoader()    ' Loaderリセットフォーム 'V2.2.0.0⑤
            ObjPlcIf = New DllPlcIf.DllMelsecPLCIf()
            objLoaderInfo = New frmLoaderInfo()         ' Loader関係情報表示'V2.2.0.0⑤
            ObjSys.frmResetLoaderInitial()              'V2.2.0.0⑤

            '---------------------------------------------------------------------------
            '   OcxSystem.ocx用の初期設定処理を行う
            '---------------------------------------------------------------------------
            ' OcxSystem用のオブジェクトを設定する
            For i = 0 To 31
                ObjMON(i) = HostSignal(i)
            Next i
            'HostSignal(0).BackColor = Color.Black

            'Call ObjSys.SetOcxUtilityObject(Utility1)  ' OcxUtility.ocx
            Call System1.SetOcxUtilityObject(ObjUtl)    ' OcxUtility.ocx
            'r = System1.SetMainObject_EX(txtLog, ObjMON)       ' Mainｵﾌﾞｼﾞｪｸﾄ
            r = System1.SetMainObject_EX()              ' Mainｵﾌﾞｼﾞｪｸﾄ
            Call System1.SetSystemObject(System1)       ' System.ocx
            ' 親モジュールのメソッドを設定する(OcxSystem用)
            gparModules = New MainModules               ' 親側メソッド呼出しオブジェクト
            Call System1.SetMainObject(gparModules)

            VideoLibrary1.SetMainObject(gparModules)    ' 親モジュールのメソッドを設定する。

            ObjVdo.SetCrossLineObject(gparModules)      ' クロスライン表示用オブジェクト 'V2.2.1.2①


            ' ｺﾝﾊﾟｲﾙｵﾌﾟｼｮﾝを設定する
#If cOFFLINEcDEBUG = 0 Then                             ' ﾃﾞﾊﾞｯｸﾞﾓｰﾄﾞでない ?
            onoff(0) = 0                                ' OffLineﾃﾞﾊﾞｯｸﾞﾌﾗｸﾞOFF
            Call DebugMode(0, 0)                        ' DllTrimFunc.dllﾊﾞｯｸﾞﾌﾗｸﾞON
#Else
            onoff(0) = 1                                ' OffLineﾃﾞﾊﾞｯｸﾞﾌﾗｸﾞON
            Call DebugMode(1, 0)                        ' DllTrimFunc.dllﾊﾞｯｸﾞﾌﾗｸﾞOFF
#End If

#If cIOcMONITORcENABLED = 0 Then
            onoff(1) = 0                                ' IOﾓﾆﾀ表示(0=表示しない, 1=表示する)
#Else
            onoff(1) = 1
#End If
            ' ｺﾝﾊﾟｲﾙｵﾌﾟｼｮﾝを設定しシステムパラメータをINtime側へ送信する
            r = Me.System1.SetOptionFlg(cMAXOptFlgNUM, onoff)
            If (r <> cFRS_NORMAL) Then
                strMSG = "Me.System1.SetOptionFlg Error (r = " & r.ToString("0") & ")"
                Call MsgBox(strMSG, MsgBoxStyle.OkOnly)
                Call AplicationForcedEnding()                   ' ｿﾌﾄ強制終了処理
                End                                             ' アプリ強制終了
                Return
            End If

            '---------------------------------------------------------------------------
            '   OcxAbout.ocx用の初期設定処理を行う
            '---------------------------------------------------------------------------
            Call HelpVersion1.SetOcxUtilityObject(Utility1) ' OcxUtility.ocx

            '---------------------------------------------------------------------------
            '   OcxPassword.ocx用の初期設定処理を行う
            '---------------------------------------------------------------------------
            Call Password1.SetOcxUtilityObject(Utility1)    ' OcxUtility.ocx

            '---------------------------------------------------------------------------
            '   OcxManualTeach.ocx用の初期設定処理を行う
            '---------------------------------------------------------------------------
            Call ManualTeach1.SetOcxUtilityObject(Utility1) ' OcxUtility1.ocx
            Call ManualTeach1.SetSystemObject(System1)      ' System.ocx

            '---------------------------------------------------------------------------
            '   DllgSysPrm.dll用の初期設定処理を行う
            '---------------------------------------------------------------------------
            Call DllSysPrmSysParam_definst.SetOcxUtilityObjectForSysprm(Utility1)

            '---------------------------------------------------------------------------
            '   Teach.ocx用の初期設定処理を行う
            '---------------------------------------------------------------------------
            Call Teaching1.SetOcxUtilityObject(Utility1)    ' OcxUtility1.ocx
            Call Teaching1.SetSystemObject(System1)         ' System.ocx

            '---------------------------------------------------------------------------
            '   Probe.ocx用の初期設定処理を行う
            '---------------------------------------------------------------------------
            Call Probe1.SetOcxUtilityObject(Utility1)       ' OcxUtility1.ocx
            Call Probe1.SetSystemObject(System1)            ' System.ocx

            '---------------------------------------------------------------------------
            '   Video.ocx用の初期設定処理を行う
            '---------------------------------------------------------------------------
            Call VideoLibrary1.SetOcxUtilityObject(Utility1) ' OcxUtility1.ocx
            Call VideoLibrary1.SetSystemObject(System1)      ' System.ocx
            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Ocx_Initialize() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try

    End Sub
#End Region

#Region "フォームロード時の処理"
    '''=========================================================================
    ''' <summary>フォームロード時の処理</summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Dim i As Short
        Dim r As Short
        Dim dispSize As System.Drawing.Size
        Dim dispPos As System.Drawing.Point
        Dim strMSG As String                                            ' ﾒｯｾｰｼﾞ編集域

        Try
            AddHandler CbDigSwL.MouseWheel, AddressOf CbDigSwL_MouseWheel   'V2.0.0.0⑥

            Me.Visible = False                                          ' 画面非表示

            Me.AutoRunnningDisp.BackColor = System.Drawing.Color.Yellow  ' 背景色 = 黄色
            Me.AutoRunnningDisp.Text = "自動運転解除中"

            ' I/Oﾓﾆﾀ表示
#If cIOcMONITORcENABLED = 0 Then
            For i = 0 To 31
                HostSignal(i).Visible = False                           ' I/Oﾓﾆﾀ非表示
            Next
            Label2.Visible = False                                      ' H非表示
            Label3.Visible = False                                      ' L非表示
            Label4.Visible = False                                      ' L非表示
            Label5.Visible = False                                      ' H非表示
#End If

            ' ﾃﾞﾊﾞｯｸﾞ用ﾎｽﾄｺﾏﾝﾄﾞ非表示/表示
#If (cIOcHostComandcENABLED = 0) Then
            For i = 0 To 8
                DEBUG_HST_CMD(i).Visible = False                        ' ﾃﾞﾊﾞｯｸﾞ用ﾎｽﾄｺﾏﾝﾄﾞ非表示
            Next
#Else
		For i = 0 To 8
            DEBUG_HST_CMD(i).Visible = True                             ' ﾃﾞﾊﾞｯｸﾞ用ﾎｽﾄｺﾏﾝﾄﾞ表示
		Next
#End If
            ' 画面表示位置の設定
            dispPos.X = 0
            dispPos.Y = 0
            Me.Location = dispPos

            ' 画面サイズの設定
            dispSize.Height = 1024
            dispSize.Width = 1280
            Me.Size = dispSize

            'ボタン表示他
            Call SetButtonImage()                                       ' フォームのボタン名の設定(日本語/英語)
            Call Btn_Enb_OnOff(2)                                       ' ボタン等の表示/非表示
            Call Btn_Enb_OnOff(1)                                       ' ボタン活性化/非活性化
            Call Disp_frmInfo(COUNTER.COUNTUP, COUNTER.INITIAL_DISP)    ' トリミング結果表示(ﾄﾘﾐﾝｸﾞ前(全ﾜｰｸ))

            ' Ocxﾃﾞﾊﾞｯｸﾞﾓｰﾄﾞ設定
#If cOFFLINEcDEBUG Then
            VideoLibrary1.cOFFLINEcDEBUG = &H3141S
            Teaching1.cOFFLINEcDEBUG = &H3141S
            Probe1.cOFFLINEcDEBUG = &H3141S
            'trimmer.cOFFLINEcDEBUG = &H3141
            ctl_LaserTeach1.cOFFLINEcDEBUG = &H3141S
            Me.AutoScroll = True
#End If

            ' Video.ocxのDbgOn/Offﾎﾞﾀﾝの有効/無効指定(デバッグ用)
            '　デバッグ時変数内容を表示させるため
#If cDBGRdraw Then                                                      ' Video.ocxのDbgOn/Offﾎﾞﾀﾝ有効とする ?
		VideoLibrary1.cDBGRdraw = &H3142
#End If
            Text2.Text = ""

            ' コントロールを非表示にする
            Probe1.Visible = False
            Teaching1.Visible = False
            HelpVersion1.Visible = False

            ' プローブ位置合わせのコントロールの表示位置を指定する
            Probe1.Left = Text2.Left
            Probe1.Top = Text2.Top

            ' ティーチングのコントロールの表示位置を指定する
            Teaching1.Left = Text2.Left
            Teaching1.Top = Text2.Top
            Call Z_PRINT("")

            UserSub.LaserCalibrationModeLoad()                          'V2.1.0.0② レーザパワーモニタリングモード取得ボタン表示
            UserSub.LaserCalibrationSet(POWER_CHECK_LOT)                'V2.1.0.0② レーザパワーモニタリング実行有無設定

            '---------------------------------------------------------------------------
            '   装置初期化処理
            '---------------------------------------------------------------------------
            Call Me.Initialize_VideoLib()                               ' ビデオライブラリ初期化
            Call Me.VideoLibrary1.VideoStop()                           ' 原点復帰処理で表示がフリーズ可能性があるため一旦停止
            '-------------------------------------------------------------------
            '   原点復帰処理
            '-------------------------------------------------------------------
            Call Me.Initialize_TrimMachine()                            ' 原点復帰処理とFLへの初期化ファイル送付
            Me.Visible = True                                           ' 画面表示
            Me.Refresh()                                                'V2.0.0.3①
            '-------------------------------------------------------------------
            '   データロード
            '-------------------------------------------------------------------
            Call GetLotInf()                                            ' INIファイル保存時のロット情報ロード

            'V2.2.0.0⑯↓ 
            stMultiBlock.gMultiBlock = 0
            stMultiBlock.Initialize()
            For i = 0 To 5
                stMultiBlock.BLOCK_DATA(i).DataNo = i + 1           ' DataNo
                stMultiBlock.BLOCK_DATA(i).Initialize()
                stMultiBlock.BLOCK_DATA(i).gBlockCnt = 0            ' ブロック数
            Next
            ''V2.2.0.0⑯↑

            r = UserVal()                                               ' データ初期設定
            If (r = 1) Then                                             ' データロードエラー ?
                pbLoadFlg = False                                       ' データロード済フラグ = False
                strMSG = "Data load Error : " & gsDataFileName & vbCrLf
                Call Z_PRINT(strMSG)
            ElseIf (r = 2) Then                                         ' システム変数設定エラー ?
                Call Me.System1.TrmMsgBox(gSysPrm, "システム変数設定エラー!!(Aplication End)", MsgBoxStyle.OkOnly, My.Application.Info.Title)
                Call AplicationForcedEnding()                           ' ｿﾌﾄ強制終了処理
                End                                                     ' アプリ強制終了
                Return
            Else
                pbLoadFlg = True                                        ' データロード済フラグ = True
                strMSG = "Data loaded : " & gsDataFileName & vbCrLf
                Call Z_PRINT(strMSG)
            End If
            If (pbLoadFlg = True) Then
                LblDataFileName.Text = gsDataFileName
            Else
                LblDataFileName.Text = ""
            End If

            Call UserSub.SetStartCheckStatus(True)      ' 設定画面の確認有効化

            UserBas.TrimmingDataChange()        ' ###1041①

            If UserSub.IsTrimType1() Or UserSub.IsTrimType4() Then       ' ###1033 温度センサーの時以外は、常時点灯'V2.0.0.0①sTrimType4()追加 
                UserBas.BackLight_Off()         ' ###1033
            Else                                ' ###1033
                UserBas.BackLight_On()          ' ###1033
            End If                              ' ###1033
            '###1040⑥            Call SetATTRateToScreen(True)       ' ###1040③ トリミングデータでのＡＴＴ減衰率の設定

            '-----------------------------------------------------------------------
            '   FL側へ加工条件を送信する(FL時で加工条件ファイルがある場合)
            '-----------------------------------------------------------------------
            Call SendFlParam(gsDataFileName)

            ' 内部カメラに切り替える
            'V2.2.2.0①↓
            ' 内部カメラ番号取得
            'V2.2.2.0①　Call Me.VideoLibrary1.ChangeCamera(0)
            Call Me.VideoLibrary1.ChangeCamera(INTERNAL_CAMERA)
            'V2.2.2.0①↑
            Call Me.VideoLibrary1.VideoStart()

            ' コンソールのラッチ解除
            Call ZCONRST()

            ' ランプ制御
            Call Me.System1.sLampOnOff(LAMP_START, True)                ' STARTﾗﾝﾌﾟON
            Call Me.System1.sLampOnOff(LAMP_RESET, True)                ' RESETﾗﾝﾌﾟON

            ' 画像表示プログラムを起動する
            'V2.2.0.0① Execute_GazouProc(ObjGazou, DISPGAZOU_PATH, DISPGAZOU_WRK, INTERNAL_CAMERA)

            '-----------------------------------------------------------------------
            '   監視タイマー開始
            '-----------------------------------------------------------------------
            giAppMode = APP_MODE_IDLE                                   ' ｱﾌﾟﾘﾓｰﾄﾞ = トリマ装置アイドル中
            Timer1.Interval = 10                                        ' 監視タイマー値(msec)
            Timer1.Enabled = True                                       ' 監視タイマー開始

            gObjFrmDistribute = New frmDistribution                     ' 分布図データオブジェクト生成 'V2.0.0.0⑨

            'V2.2.0.0⑯↓
            ' 複数抵抗値実行時の結果格納初期化
            For rn As Integer = 0 To MAX_RES_USER - 1
                stToTalDataMulti(rn).Initialize()
            Next rn
            'V2.2.0.0⑯↑

            gObjFrmDistribute.ClearCounter()                            ' 分布図データ初期化           'V2.0.0.0⑨

            'V1.2.0.2            If gSysPrm.stLOG.giLoggingMode = 1 Then                     ' デバッグログの出力有無
            'V1.2.0.2               bDebugLogOut = True
            'V1.2.0.2            Else
            'V1.2.0.2               bDebugLogOut = False
            'V1.2.0.2            End If

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Form1_Load() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "フォームのボタン名の設定(日本語/英語)"
    '''=========================================================================
    ''' <summary>フォームのボタン名の設定(日本語/英語)</summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub SetButtonImage()

        Dim strMSG As String

        Try
            ' ディジタルスイッチの設定(日本語 / 英語)
            SetDigSwImage()

            ' ボタン名の設定(日本語/英語)
            If (gSysPrm.stTMN.giMsgTyp = 0) Then
                '-------------------------------------------------------------------
                '   日本語設定
                '-------------------------------------------------------------------
                cmdLotInfo.Text = "データ設定" & vbCrLf & "(F1)"
                cmdLoad.Text = "ロード" & vbCrLf & "(F2)"
                cmdSave.Text = "セーブ" & vbCrLf & "(F3)"
                cmdEdit.Text = "編集" & vbCrLf & "(F4)"
                cmdLaserTeach.Text = "レーザ"
                cmdProbeTeaching.Text = "プローブ" & vbCrLf & "(F6)"
                cmdTeaching.Text = "ティーチング" & vbCrLf & "(F7)"
                cmdLotChg.Text = "自動運転" & vbCrLf & "(F5)"
                cmdCutPosTeach.Text = "カット位置補正" & vbCrLf & "(F8)"
                BtnRECOG.Text = "パターン登録" & vbCrLf & "(F9)"
                cmdExit.Text = "終了"

            Else
                '-------------------------------------------------------------------
                '   英語設定
                '-------------------------------------------------------------------
                cmdLotInfo.Text = "DATA" & vbCrLf & "(F1)"
                cmdLoad.Text = "LOAD" & vbCrLf & "(F2)"
                cmdSave.Text = "SAVE" & vbCrLf & "(F3)"
                cmdEdit.Text = "EDIT" & vbCrLf & "(F4)"
                cmdLaserTeach.Text = "LASER"
                cmdProbeTeaching.Text = "PROBE" & vbCrLf & "(F6)"
                cmdTeaching.Text = "TEACH" & vbCrLf & "(F7)"
                cmdLotChg.Text = "LOT CHANGE" & vbCrLf & "(F5)"
                cmdCutPosTeach.Text = "CUT POS RECOG" & vbCrLf & "(F8)"
                BtnRECOG.Text = "RECOG" & vbCrLf & "(F9)"
                cmdExit.Text = "END"
            End If

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "SetButtonImage() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "ディジタルスイッチの設定(日本語/英語)"
    '''=========================================================================
    ''' <summary>ディジタルスイッチの設定(日本語/英語)</summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub SetDigSwImage()

        Dim strMSG As String

        Try
            If (gSysPrm.stTMN.giMsgTyp = 0) Then
                '-------------------------------------------------------------------
                '   日本語設定
                '-------------------------------------------------------------------
                ' ディジタルスイッチHI
                CbDigSwH.Items.Clear()
                CbDigSwH.Items.Add("０：表示なし")
                CbDigSwH.Items.Add("１：ＮＧのみ表示")
                CbDigSwH.Items.Add("２：全て表示")

                ' ディジタルスイッチLO
                CbDigSwL.Items.Clear()
                CbDigSwL.Items.Add("０：トリミング")
                CbDigSwL.Items.Add("１：測定")
                CbDigSwL.Items.Add("２：カット実行")
                CbDigSwL.Items.Add("３：ステップ＆リピート")
                CbDigSwL.Items.Add("４：測定マーキングモード")  'V1.0.4.3⑩
                CbDigSwL.Items.Add("５：電源モード")            'V2.0.0.0②
                CbDigSwL.Items.Add("６：測定値変動測定")        'V2.0.0.0②

            Else
                '-------------------------------------------------------------------
                '   英語設定
                '-------------------------------------------------------------------
                ' ディジタルスイッチHI
                CbDigSwH.Items.Clear()
                CbDigSwH.Items.Add("０：No Display")
                CbDigSwH.Items.Add("１：Display only NG Logs")
                CbDigSwH.Items.Add("２：Display All Logs")

                ' ディジタルスイッチLO
                CbDigSwL.Items.Clear()
                CbDigSwL.Items.Add("0:Trimming")
                CbDigSwL.Items.Add("1:Measure")
                CbDigSwL.Items.Add("2:Cutting")
                CbDigSwL.Items.Add("3:Step And Repeat")
                CbDigSwL.Items.Add("4:Measure Marking")  'V1.0.4.3⑩
                CbDigSwL.Items.Add("5:Power")            'V2.0.0.0②
                CbDigSwL.Items.Add("6:Meas Variation")   'V2.0.0.0②
            End If

            ' ディジタルスイッチの初期設定
            LblDIGSW_HI.Visible = True
            LblDIGSW_LO.Visible = True

            CbDigSwH.Visible = True
            CbDigSwL.Visible = True
            CbDigSwH.SelectedIndex = DGH
            CbDigSwL.SelectedIndex = DGL

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "SetDigSwImage() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "レーザパワー調整関連項目の設定(日本語/英語)"
    '''=========================================================================
    ''' <summary>
    ''' レーザパワー調整減衰率の設定(日本語/英語)
    ''' </summary>
    ''' <param name="bMode">True:ATTの設定、False:パワー調整後のATTデータの保存</param>
    ''' <remarks>###1040③で分離</remarks>
    '''=========================================================================
    Public Function SetATTRateToScreen(ByVal bMode As Boolean) As Boolean   'V2.1.0.0②SubからFunctionへ変更

        Dim strMSG As String
        Dim iRtn As Integer

        Try
            ' 減衰率をシスパラより表示する("減衰率 = 99.9%")
            ' ※ﾛｰﾀﾘｱｯﾃﾈｰﾀの設定はOcxSystemの原点復帰処理で行われる
            If (gSysPrm.stRMC.giRmCtrl2 >= 2 And
                gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then           ' ﾛｰﾀﾘｱｯﾃﾈｰﾀ制御有(RMCTRL2対応時有効) ?

                'V2.1.0.0②                If pbLoadFlg And stLASER.iTrimAtt = 1 Then                                      ' ###1040③
                If pbLoadFlg AndAlso (stLASER.iTrimAtt = 1 OrElse stLASER.iAttNo > 0) Then            'V2.1.0.0②
                    If bMode Then                                                               ' ###1040③ レーザパワー調整後
                        'V2.1.0.0⑥カバー開のエラー                        Call ATTRESET() 'V2.1.0.0⑥
                        iRtn = LATTSET(stLASER.iFixAtt, stLASER.dblRotAtt)                      ' ###1040③
                        iRtn = ObjSys.EX_ZGETSRVSIGNAL(gSysPrm, iRtn, 0)                        ' ###1040③
                        If (iRtn <> cFRS_NORMAL) Then                                           ' ###1040③
                            Call Z_PRINT("ロータリーアッテネータの設定が異常終了しました。[" & CStr(iRtn) & "]")          ' ###1040③
                            Return (False)                                                       'V2.1.0.0②
                            'V2.1.0.0②                            MsgBox("ロータリーアッテネータの設定が異常終了しました。[" & CStr(iRtn) & "]")          ' ###1040③
                            'V2.1.0.0②                            Exit Function                                                            ' ###1040③
                        End If                                                                  ' ###1040③
                    Else                                                                        ' ###1040③
                        stLASER.dblRotPar = gSysPrm.stRAT.gfAttRate                             ' ###1040③ 減衰率(%)
                        stLASER.dblRotAtt = gSysPrm.stRAT.giAttRot                              ' ###1040③ ロータリーアッテネータの回転量(0-FFF)
                        stLASER.iFixAtt = gSysPrm.stRAT.giAttFix                                ' ###1040③ 固定アッテネータ(0:OFF, 1:ON)
                    End If

                    If (gSysPrm.stTMN.giMsgTyp = 0) Then                                        ' ###1040③
                        strMSG = "減衰率 " + CDbl(stLASER.dblRotPar).ToString("##0.00") + " %"   ' ###1040③
                    Else                                                                        ' ###1040③
                        strMSG = "ATT. " + CDbl(stLASER.dblRotPar).ToString("##0.00") + " %"     ' ###1040③
                    End If                                                                      ' ###1040③
                Else                                                                            ' ###1040③
                    If (gSysPrm.stTMN.giMsgTyp = 0) Then
                        strMSG = "減衰率 " + CDbl(gSysPrm.stRAT.gfAttRate).ToString("##0.0") + " %"
                    Else
                        strMSG = "ATT. " + CDbl(gSysPrm.stRAT.gfAttRate).ToString("##0.0") + " %"
                    End If
                End If                                                                          ' ###1040③
                Me.LblRotAtt.Text = strMSG                              ' 減衰率表示
                Me.LblRotAtt.Visible = True
            Else
                Me.LblRotAtt.Visible = False
            End If

            Return (True)                                                                        'V2.1.0.0②

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "SetATTRateToScreen() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (False)                                                       'V2.1.0.0②
        End Try
    End Function
#End Region

#Region "レーザパワー調整関連項目の設定(日本語/英語)"
    '''=========================================================================
    '''<summary>レーザパワー調整関連項目の設定(日本語/英語)</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub SetLaserItems()

        Dim strMSG As String

        Try

            '###1040⑥            Call SetATTRateToScreen(True)
            Call SetATTRateToScreen(False)   '###1040⑥

            ' 測定値をシスパラより表示する
            ' ﾌﾟﾛｸﾞﾗﾑ起動時はﾚｰｻﾞｰﾊﾟﾜｰ設定値は「-----」表示とする
            If (gSysPrm.stRMC.giRmCtrl2 >= 3) And (gSysPrm.stRMC.giPMonHi = 1) Then ' RMCTRL2 >=3 で 測定値表示 ?
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    strMSG = "レーザパワー設定値　---- W"
                Else
                    strMSG = "Laser Power ---- W"
                End If
                LblMes.Text = strMSG                                    ' 測定パワー[W]の表示
            Else
                LblMes.Visible = False                                  ' 測定値非表示
            End If

            ' 定電流値を表示する
            LblCur.Visible = False                                      ' 定電流値非表示
            If (gSysPrm.stTMN.giMsgTyp = 0) Then
                strMSG = "定電流値 "
            Else
                strMSG = "Fixed Current Val "
            End If

            Select Case (gSysPrm.stSPF.giProcPower2)
                Case 0 ' 指定なし(標準)
                    LblCur.Text = strMSG & "0.25A"                      ' "定電流値 0.25A"
                Case 1
                    LblCur.Text = strMSG & "1.00A"                      ' "定電流値 1.00A"
                Case 2
                    LblCur.Text = strMSG & "0.75A"                      ' "定電流値 0.75A"
                Case 3
                    LblCur.Text = strMSG & "0.50A"                      ' "定電流値 0.50A"
            End Select

            ' 加工電力設定 = 4(定電流1A)の時に表示
            If (gSysPrm.stSPF.giProcPower = 4) And (gSysPrm.stSPF.giProcPower2 <> 0) Then
                LblCur.Visible = True
            Else
                LblCur.Visible = False
            End If

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "SetLaserItems() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "MV10 ビデオライブラリ初期化処理"
    '''=========================================================================
    ''' <summary>MV10 ビデオライブラリ初期化処理</summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub Initialize_VideoLib()

        Dim lRet As Integer
        Dim s As String
        Dim strMSG As String

        Try
            '---------------------------------------------------------------------------
            '   ビデオライブラリを初期化する
            '---------------------------------------------------------------------------
            If pbVideoCapture = False Then
                pbVideoCapture = True
                'ChDir(My.Application.Info.DirectoryPath)
                ChDir(WORK_DIR_PATH)                                    ' MvcPt2.iniのあるﾌｫﾙﾀﾞｰを作業用ﾌｫﾙﾀﾞｰとする

                If (gSysPrm.stDEV.giEXCAM = 0) Then                      ' 内部カメラ?
                    VideoLibrary1.pp36_x = gSysPrm.stGRV.gfPixelX        ' ピクセル値X(um)
                    VideoLibrary1.pp36_y = gSysPrm.stGRV.gfPixelY        ' ピクセル値Y(um)
                Else
                    VideoLibrary1.pp36_x = gSysPrm.stGRV.gfEXCAM_PixelX  ' 外部ｶﾒﾗﾋﾟｸｾﾙ値X(um)
                    VideoLibrary1.pp36_y = gSysPrm.stGRV.gfEXCAM_PixelY  ' 外部ｶﾒﾗﾋﾟｸｾﾙ値Y(um)
                End If

                VideoLibrary1.OverLay = True
                lRet = VideoLibrary1.Init_Library()                     ' ビデオライブラリ初期化
                If (lRet <> 0) Then                                     ' Video.OCXエラー ?
                    Select Case lRet
                        Case cFRS_VIDEO_INI
                            s = "VIDEOLIB: Already initialized."
                        Case cFRS_VIDEO_PRP
                            s = "VIDEOLIB: Invalid property value."
                        Case cFRS_MVC_UTL
                            s = "VIDEOLIB: Error in MvcUtil"
                        Case cFRS_MVC_PT2
                            s = "VIDEOLIB: Error in MvcPt2"
                        Case cFRS_MVC_10
                            s = "VIDEOLIB: Error in Mvc10"
                        Case Else
                            s = "VIDEOLIB: Unexpected error 2"
                    End Select
                    Call Me.System1.TrmMsgBox(gSysPrm, s, MsgBoxStyle.OkOnly, My.Application.Info.Title)
                Else
                    ' ライブラリ初期化完了
                    pbVideoInit = True
                End If

            End If

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Initialize_VideoLib() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "原点復帰処理とFLへの初期化ファイル送付"
    '''=========================================================================
    ''' <summary>原点復帰処理とFLへの初期化ファイル送付</summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub Initialize_TrimMachine()

        Dim r As Short
        Dim strSetFileName As String = ""
        Dim strMSG As String

        Try
            '---------------------------------------------------------------------------
            '   原点復帰処理とFLへの初期化ファイル送付
            '---------------------------------------------------------------------------
            If (gflgResetStart = False) Then                            ' 初期設定済みでない ?

                ' ﾃﾝﾌﾟﾚｰﾄﾌｧｲﾙの保存場所を"C:\TRIM"に設定する(VideoStart()後に指定する)
                ' (注)管理ﾌｧｲﾙ「Pt2Template.xxx」は起動ﾌｫﾙﾀﾞに作成される。
                r = Me.VideoLibrary1.SetTemplatePass(cTEMPLATPATH)

                Call InitFunction()                                     ' DllTrimFunc.dll初期化
                If (gflgResetStart = False) Then                        ' 初期設定済みでない ?
                    If (giLoaderType <> 0) Then
                        ' 電磁ロック(観音扉右側ロック)を解除する
                        r = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_ON)
                    End If

                    ' 原点復帰
                    r = sResetStart()
                    If (r <> cFRS_NORMAL) Then                          ' エラー ?
                        If (r <> cFRS_ERR_RST) Then
                            r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE2)
                        End If

                        'V2.2.0.0⑤↓
                        If (giLoaderType <> 0) Then
                            Call Me.VideoLibrary1.ChangeCamera(0)
                            Call Me.VideoLibrary1.VideoStart()
                        End If

                        Call AplicationForcedEnding()                   ' ｿﾌﾄ強制終了処理
                        End                                             ' アプリ強制終了
                        Return
                    End If
                    gflgResetStart = True                               ' 初期設定済みON

                    If (giLoaderType <> 0) Then
                        ' 電磁ロック(観音扉右側ロック)を解除する
                        r = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_OFF)
                    End If

                    'Call BackLight_Off()                                ' バックライトＬＥＤ制御ＯＦＦ

                End If

#If cOSCILLATORcFLcUSE Then
                '-----------------------------------------------------------------------
                '   FL側へ加工条件を送信する(FL時)
                '-----------------------------------------------------------------------
                If (gSysPrm.stRAT.giOsc_Res = OSCILLATOR_FL) Then
                    If (pbLoadFlg = True) Then                          ' データロード済み ? 
                        strSetFileName = gsDataFileName                 ' ロードしたデータファイルに対応する加工条件ファイル名
                    Else
                        strSetFileName = DEF_FLPRM_SETFILENAME          ' デフォルトの加工条件ファイル名
                    End If
                    ' FL用加工条件ファイルをリードしてFL側へ加工条件を送信する
                    r = SendTrimCondInfToFL(stCND, strSetFileName, strSetFileName)
                    If (r = SerialErrorCode.rRS_FLCND_XMLNONE) And (strSetFileName <> DEF_FLPRM_SETFILENAME) Then
                        ' ロードしたデータファイルに対応するXMLファイルが存在しない場合はデフォルトの加工条件を送信する
                        strSetFileName = DEF_FLPRM_SETFILENAME          ' デフォルトの加工条件ファイル名
                        r = SendTrimCondInfToFL(stCND, strSetFileName, strSetFileName)
                    End If
                    If (r <> SerialErrorCode.rRS_OK) Then
                        '"ＦＬ通信異常。ＦＬとの通信に失敗しました。" + vbCrLf + "ＦＬと正しく接続できているか確認してください。"
                        strMSG = MSG_150
                        Call MsgBox(strMSG, MsgBoxStyle.OkOnly, "")
                    End If
                End If
#End If
                ' 処理マガジンの情報を取得 
                r = ObjSys.Sub_GetNowProcessMgInfo(gisupplyMgNum, gisupplyMgStepNum, gistoreMgNum, gistoreMgStepNum)

            End If

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Initialize_TrimMachine TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "フォームアンロード時の処理"
    '''=========================================================================
    ''' <summary>フォームアンロード時の処理</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub Form1_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        Dim lRet As Integer
        Dim strMSG As String

        Try
            ' EXTBIT OFF
            Call EXTOUT1(0, &HFFFFS)                                    ' EXTBIT(On=0, Off=全ビット)
            Call EXTOUT2(0, &HFFFFS)                                    ' EXTBIT2(On=0, Off=全ビット)

            ' トリマレディ信号OFF送信(ｵｰﾄﾛｰﾀﾞｰ)
            Call Sub_ATLDSET(COM_STS_TRM_STATE, COM_STS_TRM_READY)      ' ローダー出力(ON=トリマ動作中, ,OFF=トリマレディ)

            Call IoMonitor(gdwATLDDATA, 1)

            '' スライトカバーストッパ戻り(全開位置へ)
            'Call Me.System1.BigCover_Ctrl(gSysPrm, 1)

            ' ｼｸﾞﾅﾙﾀﾜｰ制御(On=0, Off=ﾚﾃﾞｨ, 原点復帰中, 自動運転中)
            Call Me.System1.SetSignalTower(0, &HFFFFS)

            ' ライブラリ終了
            If pbVideoInit = True Then
                lRet = VideoLibrary1.Close_Library
                If lRet <> 0 Then
                    Select Case lRet
                        Case cFRS_VIDEO_INI
                            Call Me.System1.TrmMsgBox(gSysPrm, "Video library: Not initialized.", MsgBoxStyle.OkOnly, My.Application.Info.Title)
                        Case Else
                            'MsgBox "予期せぬエラー"
                            Call Me.System1.TrmMsgBox(gSysPrm, "Video library: Unexpected error.", MsgBoxStyle.OkOnly, My.Application.Info.Title)
                    End Select
                End If
            End If

            ' ＧＰＩＢ終了処理
            ObjGpib.Gpib_Term(gDevId)

            ' 操作ログ出力
            Call Me.System1.OperationLogging(gSysPrm, MSG_OPLOG_END, "FormClosed") ' "ユーザプログラム終了"

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Form1_FormClosed() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

    '========================================================================================
    '   コマンドボタン押下時の処理
    '========================================================================================
#Region "スタートボタン押下時(デバッグ用)"
    '''=========================================================================
    ''' <summary>スタートボタン押下時</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub cmdStart_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdStart.Click

        Dim s As String
        Dim r As Short
        Dim strMSG As String

        Try
            ' 初期処理
            If pbLoadFlg = False Then
                s = MSG_DataNotLoad                         ' ﾃﾞｰﾀ未ﾛｰﾄﾞ
                Call Z_PRINT(s + vbCrLf)
                Call Beep()
                Exit Sub
            End If
            Timer1.Enabled = False                          'タイマー停止

            ' 操作ログ出力(ﾄﾘﾐﾝｸﾞ)
            Call Me.System1.OperationLogging(gSysPrm, MSG_OPLOG_TRIMST, "DEBUG")

            ' トリミング処理
            r = User()

            ' 後処理
            Timer1.Enabled = True                           ' タイマー開始
            Call ZCONRST()                                  ' ｺﾝｿｰﾙSWﾗｯﾁ解除

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "cmdStart_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "終了ボタン押下時"
    '''=========================================================================
    ''' <summary>終了ボタン押下時</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub cmdExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExit.Click

        Dim s As String
        Dim r As Short
        Dim strMSG As String

        Try
            ' トリマ動作中信号ON送信(ｵｰﾄﾛｰﾀﾞｰ)
            Call Sub_ATLDSET(COM_STS_TRM_STATE, 0)                      ' ローダー出力(ON=トリマ動作中,OFF=なし)

            ' 終了確認ﾒｯｾｰｼﾞを設定する
            giAppMode = APP_MODE_EXIT                                   ' ｱﾌﾟﾘﾓｰﾄﾞ設定

            Timer1.Enabled = False                                      ' 監視タイマー停止
            Call ZCONRST()                                              ' ラッチ解除

            '' トリマ装置アイドル中以外ならNOP
            'If giAppMode Then GoTo STP_END

            If pbLoadFlg = False Then                                   ' データ未ロード ?
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    s = "           終了しますか？           "
                Else
                    s = "      Are you sure to quit ?      "
                End If

            ElseIf (FlgUpd = TriState.True) Then                        ' データロード済みで更新 ?
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    s = "編集中のデータがあります。" & vbCrLf & "アプリケーションを終了してよろしいですか？"
                Else
                    s = "  Please make sure to save the data to the disk before quit this program.  " & vbCrLf & "  Are you sure to quit?  "
                End If

            Else
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    s = "アプリケーションを終了してよろしいですか？"
                Else
                    s = "      Are you sure to quit ?      "
                End If
            End If

            ' 終了確認ﾒｯｾｰｼﾞ表示
            r = Me.System1.TrmMsgBox(gSysPrm, s, MsgBoxStyle.OkCancel, "QUIT")
            If (r = cFRS_ERR_ADV) Then                                  ' OK(ADVｷｰ) ?
                '　ソフト強制終了処理
                Call AplicationForcedEnding()
                End
                Exit Sub
            End If

STP_END:
            Call ZCONRST()                                              ' ｺﾝｿｰﾙｷｰ ﾗｯﾁ解除
            giAppMode = APP_MODE_IDLE                                   ' ｱﾌﾟﾘﾓｰﾄﾞ = トリマ装置アイドル中
            Call Sub_ATLDSET(0, COM_STS_TRM_STATE)                      ' ローダー出力(ON=なし,OFF=トリマ動作中)
            Timer1.Enabled = True                                       ' 監視タイマー開始

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "cmdExit_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "ロードボタン押下時処理"
    Public Function TrimDataLoad(ByVal sFileName As String) As Boolean
        Dim s As String
        Dim r As Short

        gsDataFileName = sFileName                     ' データファイル名設定

        ' 旧設定の装置の電圧をOFFする
        r = V_Off()                                     ' DC電源装置 電圧OFF処理

        Call Z_CLS()                            ' データロードでログ画面クリア                   ###lstLog'V2.0.0.0⑮

        ' トリミングデータ設定
        r = UserVal()                                   ' データ初期設定
        If (r <> 0) Then                                ' エラー ?
            pbLoadFlg = False                           ' データロード済フラグ = False
            s = "Data load Error : " & sFileName & vbCrLf
            LblDataFileName.Text = ""
            Call Z_PRINT(s)
            Return (False)
        Else
            'V2.0.0.0⑮            Call Z_CLS()                                ' データロードでログ画面クリア                   ###lstLog
            gDspCounter = 0                             ' ログ画面表示基板枚数カウンタクリア
            pbLoadFlg = True                            ' データロード済フラグ = True
            s = "Data loaded : " & sFileName & vbCrLf
            Call Z_PRINT(s)

            Call Me.System1.OperationLogging(gSysPrm, MSG_OPLOG_FUNC01, "File='" & sFileName & "' MANUAL")

            ' トリミングデータファイル名をロット情報ファイルに出力する
            Call PutLotInf()
            ' ファイルパス名の表示
            'V2.1.0.0④↓
            If sFileName.Length > 60 Then
                LblDataFileName.Text = sFileName
            Else
                'V2.1.0.0④↑
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    LblDataFileName.Text = "データファイル名 " & sFileName
                Else
                    LblDataFileName.Text = "File name " & sFileName
                End If
            End If          'V2.1.0.0④

#If cOSCILLATORcFLcUSE Then
            '-----------------------------------------------------------------------
            '   FL側へ加工条件を送信する(FL時で加工条件ファイルがある場合)
            '-----------------------------------------------------------------------
            Call SendFlParam(sFileName)
#End If
            UserBas.TrimmingDataChange()    ' ###1041①
            UserSub.LaserCalibrationSet(POWER_CHECK_LOT)            'V2.1.0.0② レーザパワーモニタリング実行有無設定

            'If giLoaderType <> 0 Then   'クランプ吸着動作設定
            '    ObjSys.setClampVaccumConfig(stUserData.intClampVacume)
            'End If

            If UserSub.IsTrimType1() Or UserSub.IsTrimType4() Then       ' 温度センサーの時以外は、常時点灯'V2.0.0.0①sTrimType4()追加
                UserBas.BackLight_Off()
            Else
                UserBas.BackLight_On()
            End If
        End If
        '###1040⑥        Call SetATTRateToScreen(True)           ' ###1040③ トリミングデータでのＡＴＴ減衰率の設定

        Return (True)
    End Function

    '''=========================================================================
    ''' <summary>ロードボタン押下時処理</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub cmdLoad_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdLoad.Click

        Dim rslt As Short
        Dim s As String
        Dim r As Short
        Dim strMSG As String
        Dim result As System.Windows.Forms.DialogResult

        Try
            '-----------------------------------------------------------------------
            '   初期処理
            '-----------------------------------------------------------------------
            ' トリマ動作中信号ON送信(ｵｰﾄﾛｰﾀﾞｰ)
            Call Sub_ATLDSET(COM_STS_TRM_STATE, 0)              ' ローダー出力(ON=トリマ動作中,OFF=なし)

            ' トリマ装置アイドル中以外ならNOP
            If giAppMode Then GoTo STP_END
            giAppMode = APP_MODE_LOAD                           ' ｱﾌﾟﾘﾓｰﾄﾞ = ファイルロード(F1)

            ' パスワード入力(オプション)
            rslt = Func_Password(F_LOAD)
            If (rslt <> True) Then
                GoTo STP_END                                    ' ﾊﾟｽﾜｰﾄﾞ入力ｴﾗｰならEXIT
            End If

            ' データロード済みチェック
            If (pbLoadFlg = True) Then
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    s = "ロード済みのデータがあります。別のデータをロードしますか？"
                Else
                    s = "Current data will be lost. Are you sure to load another data?"
                End If
                r = Me.System1.TrmMsgBox(gSysPrm, s, MsgBoxStyle.OkCancel, cAPPcTITLE)
                If (r = cFRS_ERR_RST) Then ' Cancel(RESETｷｰ) ?
                    Call Z_PRINT("Canceled data load." + vbCrLf)
                    GoTo STP_END
                End If
            End If

            'V2.2.0.0⑤↓
            If giLoaderType <> 0 Then
                ChkLoaderInfoDisp(0)
            End If
            'V2.2.0.0⑤↑

            '-----------------------------------------------------------------------
            '   【ﾌｧｲﾙを開く】ﾀﾞｲｱﾛｸﾞを表示する
            '-----------------------------------------------------------------------
#If cKEYBOARDcUSE <> 1 Then
            ' ソフトウェアキーボードを起動する
            Dim procHandle As Process
            procHandle = New Process
            Call StartSoftwareKeyBoard(procHandle)              ' ソフトウェアキーボードを起動する
#End If

            FileDlgOpen.InitialDirectory = "C:\TRIMDATA\DATA"
            FileDlgOpen.FileName = ""
            FileDlgOpen.Filter = "*.txt|*.txt"
            FileDlgOpen.ShowReadOnly = False
            FileDlgOpen.CheckFileExists = True
            FileDlgOpen.CheckPathExists = True

            ' 【ﾌｧｲﾙを開く】ﾀﾞｲｱﾛｸﾞを表示する
            result = FileDlgOpen.ShowDialog()

#If cKEYBOARDcUSE <> 1 Then
            ' ソフトウェアキーボードを終了する
            Call EndSoftwareKeyBoard(procHandle)
#End If

            ' OK以外の場合
            If (result <> Windows.Forms.DialogResult.OK) Then
                GoTo Cansel                                     ' Cansel指定なら終了
            End If

            '-----------------------------------------------------------------------
            '   データファイルをリードする
            '-----------------------------------------------------------------------
            'If (FileDlgOpen.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            ' データファイルリード
            If FileDlgOpen.FileName <> "" Then
                If Not TrimDataLoad(FileDlgOpen.FileName) Then
                    GoTo Cansel
                End If
            End If

            '-----------------------------------------------------------------------
            '   終了処理
            '-----------------------------------------------------------------------
Cansel:
            ChDrive("C")                                        ' ChDriveしないと次起動時FDドライブを見に行って,
            ChDir(My.Application.Info.DirectoryPath)            ' "MVCutil.dllがない"となり起動できなくなる

STP_END:
            Call ZCONRST()                                      ' ｺﾝｿｰﾙｷｰ ﾗｯﾁ解除
            giAppMode = APP_MODE_IDLE                           ' ｱﾌﾟﾘﾓｰﾄﾞ = トリマ装置アイドル中
            Call Sub_ATLDSET(0, COM_STS_TRM_STATE)              ' ローダー出力(ON=なし,OFF=トリマ動作中)
            'V2.2.0.0⑤↓
            If giLoaderType <> 0 Then
                ChkLoaderInfoDisp(1)
            End If
            'V2.2.0.0⑤↑

            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "CmdLoad_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            GoTo STP_END
        End Try
    End Sub
#End Region

#Region "FL側へ加工条件を送信する"
    '''=========================================================================
    ''' <summary>FL側へ加工条件を送信する</summary>
    ''' <param name="DataFileName">(INP)トリミングデータファイル名</param>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub SendFlParam(ByVal DataFileName As String)

        Dim strMSG As String

        Try
#If cOSCILLATORcFLcUSE Then
        Dim strXmlFName As String
        Dim r As Integer
            '-----------------------------------------------------------------------
            '   FL側へ加工条件を送信する(FL時で加工条件ファイルがある場合)
            '-----------------------------------------------------------------------
            If (gSysPrm.stRAT.giOsc_Res = OSCILLATOR_FL) Then
                ' 加工条件ファイルが存在するかチェック
                strXmlFName = ""
                r = GetFLCndFileName(DataFileName, strXmlFName, True)
                If (r = SerialErrorCode.rRS_OK) Then                    ' 加工条件ファイルが存在する ?
                    ' データ送信中のメッセージ表示
                    strMSG = MSG_148
                    Call Z_PRINT(strMSG + vbCrLf)                       ' ﾒｯｾｰｼﾞ表示(ログ画面)

                    ' FL用加工条件ファイルをリードしてFL側へ加工条件を送信する
                    r = SendTrimCondInfToFL(stCND, DataFileName, strXmlFName)
                    If (r = SerialErrorCode.rRS_OK) Then
                        ' "FLへ加工条件を送信しました。"
                        strMSG = MSG_147 & vbCrLf & " (SendDdata File Name = " & strXmlFName & ")" + vbCrLf
                        Call Z_PRINT(strMSG)                            ' ﾒｯｾｰｼﾞ表示(ログ画面)
                    Else
                        strMSG = MSG_152                                ' "加工条件の送信に失敗しました。再度データをロードするか、編集画面から加工条件の設定を行ってください。"
                        Call MsgBox(strMSG, MsgBoxStyle.OkOnly, "")
                        Call Z_PRINT(strMSG + vbCrLf)                   ' ﾒｯｾｰｼﾞ表示(ログ画面)
                    End If
                End If
            End If
#End If

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "SendFlParam() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "セーブボタン押下時処理"
    '''=========================================================================
    ''' <summary>セーブボタン押下時処理</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks>パラメータ変更処理を行い、変更後のパラメータをデータファイルへ書込む</remarks>
    '''=========================================================================
    Public Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click

        Dim s As String
        Dim r As Short
        Dim strMSG As String
        Dim result As System.Windows.Forms.DialogResult

        Try
            '-----------------------------------------------------------------------
            '   初期処理
            '-----------------------------------------------------------------------
            ' トリマ動作中信号ON送信(ｵｰﾄﾛｰﾀﾞｰ)
            Call Sub_ATLDSET(COM_STS_TRM_STATE, 0)              ' ローダー出力(ON=トリマ動作中,OFF=なし)

            ' パスワード入力
            r = Func_Password(F_SAVE)
            If (r <> True) Then
                GoTo STP_END                                    ' ﾊﾟｽﾜｰﾄﾞ入力ｴﾗｰならEXIT
            End If

            ' 初期処理
            FlgCan = False                                      ' Cancel Flag = false
            'If (giAppMode) Then                                 ' トリマ装置アイドル中以外ならNOP
            '    GoTo STP_END
            'End If
            giAppMode = APP_MODE_SAVE                           ' 画面ステータス = ファイルセーブ(F2)
            If pbLoadFlg = False Then
                s = MSG_DataNotLoad                             ' ﾃﾞｰﾀ未ﾛｰﾄﾞ
                Call Z_PRINT(s)
                Call Beep()
                GoTo STP_END
            End If

            'V2.2.0.0⑤↓
            If giLoaderType <> 0 Then
                ChkLoaderInfoDisp(0)
            End If
            'V2.2.0.0⑤↑

            '-----------------------------------------------------------------------
            '   【名前を付けて保存】ﾀﾞｲｱﾛｸﾞを表示する
            '-----------------------------------------------------------------------
#If cKEYBOARDcUSE <> 1 Then
            ' ソフトウェアキーボードを起動する
            Dim procHandle As Process
            procHandle = New Process
            Call StartSoftwareKeyBoard(procHandle)              ' ソフトウェアキーボードを起動する
#End If

            '【名前を付けて保存】ﾀﾞｲｱﾛｸﾞを表示する
            FileDlgSave.FileName = gsDataFileName
            FileDlgSave.Filter = "*.txt | *.txt"
            FileDlgSave.OverwritePrompt = True                  ' 既に存在している場合はメッセージ ボックスを表示
            result = FileDlgSave.ShowDialog()                   ' ※ﾌｧｲﾙ名指定なしでは戻ってこない、拡張子付で戻ってくる

#If cKEYBOARDcUSE <> 1 Then
            ' ソフトウェアキーボードを終了する
            Call EndSoftwareKeyBoard(procHandle)
#End If

            ' OK以外なら終了
            If (result <> Windows.Forms.DialogResult.OK) Then
                GoTo STP_TRM                                    ' Cansel指定なら終了
            End If

            '-----------------------------------------------------------------------
            '   データファイルをセーブする
            '-----------------------------------------------------------------------
            'If (FileDlgSave.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            If (FileDlgSave.FileName <> "") Then
                If rData_save((FileDlgSave.FileName)) <> 0 Then       ' データファイルセーブ
                    GoTo STP_END
                Else
                    gsDataFileName = FileDlgSave.FileName
                    Call Z_PRINT("Data saved : " & FileDlgSave.FileName & vbCrLf)
                End If

                ' トリミングデータファイル名をロット情報ファイルに出力する
                Call PutLotInf()

                '-----------------------------------------------------------------------
                '   操作ログ等を出力する
                '-----------------------------------------------------------------------
                Call Me.System1.OperationLogging(gSysPrm, MSG_OPLOG_FUNC02, "File='" & gsDataFileName & "' MANUAL")

                ' ファイルパス名の表示
                'V2.1.0.0④↓
                If gsDataFileName.Length > 60 Then
                    LblDataFileName.Text = gsDataFileName
                Else
                    'V2.1.0.0④↑
                    If (gSysPrm.stTMN.giMsgTyp = 0) Then
                        LblDataFileName.Text = "データファイル名 " & gsDataFileName
                    Else
                        LblDataFileName.Text = "File name " & gsDataFileName
                    End If
                End If          'V2.1.0.0④

#If cOSCILLATORcFLcUSE Then
                Dim strXmlFName As String
                '-----------------------------------------------------------------------
                '   FL側から現在の加工条件を受信してFL用加工条件ファイルをライトする(FL時)
                '-----------------------------------------------------------------------
                If (gSysPrm.stRAT.giOsc_Res = OSCILLATOR_FL) Then ' FL(ﾌｧｲﾊﾞｰﾚｰｻﾞ) ? 
                    strXmlFName = ""
                    r = RcvTrimCondInfToFL(stCND, gsDataFileName, strXmlFName)
                    If (r = SerialErrorCode.rRS_OK) Then
                        ' "加工条件ファイルを作成しました。"
                        strMSG = MSG_142 + vbCrLf + " (File Name = " + strXmlFName + ")" + vbCrLf
                        Call Z_PRINT(strMSG)                    ' ﾒｯｾｰｼﾞ表示(ログ画面)
                    End If
                End If
#End If

                FlgUpd = TriState.False                         ' データ更新 Flag OFF
            End If

STP_TRM:
            ChDrive("C")                                        ' ChDriveしないと次起動時FDドライブを見に行って,"MVCutil.dllがない"となり起動できなくなる
            ChDir(My.Application.Info.DirectoryPath)

STP_END:
            Call ZCONRST()                                      ' ｺﾝｿｰﾙｷｰ ﾗｯﾁ解除
            giAppMode = APP_MODE_IDLE                           ' ｱﾌﾟﾘﾓｰﾄﾞ = トリマ装置アイドル中
            Call Sub_ATLDSET(0, COM_STS_TRM_STATE)              ' ローダー出力(ON=なし,OFF=トリマ動作中)

            'V2.2.0.0⑤↓
            If giLoaderType <> 0 Then
                ChkLoaderInfoDisp(1)
            End If
            'V2.2.0.0⑤↑

            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "CmdSave_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            GoTo STP_END
        End Try
    End Sub
#End Region

#Region "編集ボタン押下時処理"
    '''=========================================================================
    ''' <summary>編集ボタン押下時処理</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub cmdEdit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdEdit.Click

        Dim s As String
        Dim r As Short
        'Dim ExeFile As String
        'V2.2.1.6① Dim fForm As System.Windows.Forms.Form
        Dim fForm As FormEdit.frmEdit  'V2.2.1.6①

        Dim strMSG As String
        Dim retbtn As Integer            'V2.2.1.6①

        Try
            If giAppMode <> APP_MODE_IDLE Then
                Return
            Else
                giAppMode = APP_MODE_EDIT                               ' アプリモード = データ編集
            End If
            ' トリマ動作中信号ON送信(ｵｰﾄﾛｰﾀﾞｰ)
            Call Sub_ATLDSET(COM_STS_TRM_STATE, 0)                      ' ローダー出力(ON=トリマ動作中,OFF=なし)

            If giLoaderType <> 0 Then
                Call Me.VideoLibrary1.VideoStop()                           ' 原点復帰処理で表示がフリーズ可能性があるため一旦停止
                Timer1.Enabled = False
                'V2.2.0.0⑤↓
                ChkLoaderInfoDisp(0)
                'V2.2.0.0⑤↑
            End If

            ' パスワード入力
            r = Func_Password(F_EDIT)
            If (r <> True) Then
                giAppMode = APP_MODE_IDLE                               ' ｱﾌﾟﾘﾓｰﾄﾞ = トリマ装置アイドル中
                GoTo STP_END                                            ' ﾊﾟｽﾜｰﾄﾞ入力ｴﾗｰならEXIT
            End If

            ' データロードチェック
            If pbLoadFlg = False Then
                s = MSG_DataNotLoad                                     ' ﾃﾞｰﾀ未ﾛｰﾄﾞ
                Call Z_PRINT(s)
                Call Beep()
                GoTo STP_END
            End If

            ' データ編集
            Call Me.System1.OperationLogging(gSysPrm, MSG_OPLOG_FUNC03, "")
            FlgUpdGPIB = 0                                              ' GPIBデータ更新Flag Off
            'fForm = New frmEdit                                         ' frmｵﾌﾞｼﾞｪｸﾄ生成
            fForm = New FormEdit.frmEdit                                ' frmｵﾌﾞｼﾞｪｸﾄ生成
            fForm.ShowDialog()                                          ' データ編集
            retbtn = fForm.GetResult()
            fForm.Dispose()                                             ' frmｵﾌﾞｼﾞｪｸﾄ開放

            ' GPIBデータ更新ならGPIB初期化を行う
            If (FlgUpdGPIB = 1) Then
                Call GPIB_Init()
            End If

            '    ' NOTEPADでデータファイルを開く
            '    If giAppMode Then Exit Sub
            '    giAppMode = GSTAT_EDIT                                 ' 画面ステータス = 編集画面表示  (F3)
            '    #If cOFFLINEcDEBUG Then
            '        ExeFile = "notepad.exe " + gsDataFileName
            '    #Else
            '        ExeFile = "C:\WINNT\system32\notepad.exe " + gsDataFileName
            '    #End If
            '    r = Shell(ExeFile, vbNormalFocus)

            If LaserFront.Trimmer.DllVideo.VideoLibrary.IsDigitalCamera Then
                ObjVdo.StdMagnification = CDec(stPLT.dblStdMagnification)         ' 内部カメラ表示倍率を設定 
            End If


STP_END:

            'V2.2.0.021↓
            'プローブマスターテーブルからデータを展開する
            If (stPLT.ProbNo > 0) And (DialogResult.OK = retbtn) Then       ' 指定のプローブデータを読込み設定する 
                'V2.2.1.6①↓
                '--------------------------------------------------------------------------
                '   確認ﾒｯｾｰｼﾞを表示する
                '--------------------------------------------------------------------------
                strMSG = "プローブマスターを読込みますか？"
                Dim ret As Integer = MsgBox(strMSG, DirectCast(
                            MsgBoxStyle.OkCancel +
                            MsgBoxStyle.Information, MsgBoxStyle),
                            My.Application.Info.Title)
                If (ret = MsgBoxResult.Ok) Then ' Cancel(RESETｷｰ) ?
                    ConvProbeData(stPLT.ProbNo)
                End If
                'V2.2.1.6①↑
            End If
            'V2.2.0.021↑

            UserBas.TrimmingDataChange()    ' ###1041①

            If UserSub.IsTrimType1() Or UserSub.IsTrimType4() Then       ' 温度センサーの時以外は、常時点灯'V2.0.0.0①sTrimType4()追加
                UserBas.BackLight_Off()
            Else
                UserBas.BackLight_On()
            End If

            Call ZCONRST()                                              ' ｺﾝｿｰﾙｷｰ ﾗｯﾁ解除
            giAppMode = APP_MODE_IDLE                                   ' ｱﾌﾟﾘﾓｰﾄﾞ = トリマ装置アイドル中
            Call Sub_ATLDSET(0, COM_STS_TRM_STATE)                      ' ローダー出力(ON=なし,OFF=トリマ動作中)

            If giLoaderType <> 0 Then
                Call Me.VideoLibrary1.VideoStart() '                    ' 
                Timer1.Enabled = True
                'V2.2.0.0⑤↓
                ChkLoaderInfoDisp(1)
                'V2.2.0.0⑤↑
            End If

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "cmdEdit_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        Finally
            giAppMode = APP_MODE_IDLE                                   ' ｱﾌﾟﾘﾓｰﾄﾞ = トリマ装置アイドル中
        End Try
    End Sub
#End Region

#Region "レーザ調整ボタン押下時処理"
    '''=========================================================================
    ''' <summary>レーザ調整ボタン押下時処理</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    '''=========================================================================
    Private Sub cmdLaserTeach_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdLaserTeach.Click

        Dim strMSG As String

        Try
            ' レーザ調整を実行する
            cmdLaserTeach_Proc()

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "cmdLaserTeach_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "レーザ調整を行う"
    '''=========================================================================
    ''' <summary>レーザ調整を行う</summary>
    '''=========================================================================
    Public Sub cmdLaserTeach_Proc()

        Dim r As Short
        Dim strMSG As String

        Try
            ' コマンド実行前処理
            r = Sub_cmdInit_Proc(APP_MODE_LASER, F_LASER)
            If (r = cFRS_ERR_RST) Then                                  ' Cancel(RESETｷｰ) ?
                If (gSysPrm.stSPF.giWithStartSw = 0) Then                ' ｽﾀｰﾄSW押下待ち(ｵﾌﾟｼｮﾝ)でない ?
                    GoTo STP_END
                Else                                                    ' ｽﾀｰﾄSW押下待ち(ｵﾌﾟｼｮﾝ)
                    GoTo STP_TRM
                End If
            End If
            If (r < cFRS_NORMAL) Then                                   ' エラー ? 
                ' カバー開検出(ｽﾀｰﾄSW押下待ち(ｵﾌﾟｼｮﾝ)時発生する)
                If (r = cFRS_ERR_CVR) Or (r = cFRS_ERR_SCVR) Or (r = cFRS_ERR_LATCH) Then
                    r = cFRS_NORMAL
                    GoTo STP_TRM                                        ' カバー開検出ならｽﾗｲﾄﾞｶﾊﾞｰｸﾛｽﾞしてREADY状態へ
                Else
                    GoTo STP_END
                End If
            End If

            ' レーザー調整処理を行う
            gbInitialized = False                                       ' flg = 原点復帰未
            SetLaserItemsVisible(0)                                     ' レーザパワー調整関連項目を非表示とする
            r = User_LaserTeach()                                       ' レーザー調整画面表示
            SetLaserItemsVisible(1)                                     ' レーザパワー調整関連項目を表示とする

            If r = cFRS_ERR_CVR Then                                    ' レーザコマンド中に筐体カバー開は、Sub_cmdTerm_Proc()内で強制終了とする。
                r = cFRS_ERR_EMG
            End If
            ' コマンド実行後処理
STP_TRM:
            Call Sub_cmdTerm_Proc(APP_MODE_LASER, r)

            ' コマンド終了処理
STP_END:
            Call Sub_cmdEnd_Proc()

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "cmdLaserTeach_Proc() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

    'V2.1.0.0②↓
#Region "レーザ調整を行う"
    ''' <summary>
    ''' レーザパワーキャリブレーション
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub cmdLaserTeach_Calibration()

        Dim r As Short
        Dim strMSG As String

        Try
            ' コマンド実行前処理
            r = Sub_cmdInit_Proc(APP_MODE_LASER, F_LASER)
            If (r = cFRS_ERR_RST) Then                                  ' Cancel(RESETｷｰ) ?
                If (gSysPrm.stSPF.giWithStartSw = 0) Then                ' ｽﾀｰﾄSW押下待ち(ｵﾌﾟｼｮﾝ)でない ?
                    GoTo STP_END
                Else                                                    ' ｽﾀｰﾄSW押下待ち(ｵﾌﾟｼｮﾝ)
                    GoTo STP_TRM
                End If
            End If
            If (r < cFRS_NORMAL) Then                                   ' エラー ? 
                ' カバー開検出(ｽﾀｰﾄSW押下待ち(ｵﾌﾟｼｮﾝ)時発生する)
                If (r = cFRS_ERR_CVR) Or (r = cFRS_ERR_SCVR) Or (r = cFRS_ERR_LATCH) Then
                    r = cFRS_NORMAL
                    GoTo STP_TRM                                        ' カバー開検出ならｽﾗｲﾄﾞｶﾊﾞｰｸﾛｽﾞしてREADY状態へ
                Else
                    GoTo STP_END
                End If
            End If

            ' レーザー調整処理を行う
            gbInitialized = False                                       ' flg = 原点復帰未
            SetLaserItemsVisible(0)                                     ' レーザパワー調整関連項目を非表示とする
            r = User_LaserTeach(True)                                   ' 引数がTrueは、レーザパワーキャリブレーション
            SetLaserItemsVisible(1)                                     ' レーザパワー調整関連項目を表示とする

            If r = cFRS_ERR_CVR Then                                    ' レーザコマンド中に筐体カバー開は、Sub_cmdTerm_Proc()内で強制終了とする。
                r = cFRS_ERR_EMG
            End If
            ' コマンド実行後処理
STP_TRM:
            Call Sub_cmdTerm_Proc(APP_MODE_LASER, r)

            ' コマンド終了処理
STP_END:
            Call Sub_cmdEnd_Proc()

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "cmdLaserTeach_Calibration() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region
    'V2.1.0.0②↑

#Region "ロット切替ボタン押下時処理"
    '''=========================================================================
    ''' <summary>ロット切替ボタン押下時処理</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub cmdLotchg_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdLotChg.Click

        Dim Rtn As Short
        Dim strMSG As String

        Try

            ''V2.2.0.0⑤↓
            ' TLF製ローダの場合自動運転時の基板無しチェック
            If giLoaderType = 1 Then
                Timer1.Enabled = False                                      ' 監視タイマー停止
                Rtn = ObjLoader.Sub_SubstrateNothingCheck(Me.System1)
                If Rtn <> cFRS_NORMAL Then
                    GoTo STP_END
                End If
                stCounter.PlateCounter = 0
            End If
            ''V2.2.0.0⑤↑

            ' トリマ動作中信号ON送信(ｵｰﾄﾛｰﾀﾞｰ)
            Call Sub_ATLDSET(COM_STS_TRM_STATE, 0)                      ' ローダー出力(ON=トリマ動作中,OFF=なし)

            ' 他コマンド実行中 ?
            If giAppMode Then GoTo STP_END
            giAppMode = APP_MODE_LOTCHG                                 ' ｱﾌﾟﾘﾓｰﾄﾞ = ロット切替


            ' 操作ログ出力
            Call Me.System1.OperationLogging(gSysPrm, MSG_OPLOG_FUNC06, "MANUAL")

            'V2.2.0.0⑤↓
            If giLoaderType = 1 Then
                ChkLoaderInfoDisp(0)
                MarkingCount = 0             ' マーキング用カウンタクリア	V2.2.1.7③
                LotMarkingAlarmCnt = 0       ' マーキング時アラームカウンタクリア	V2.2.1.7③
            End If
            'V2.2.0.0⑤↑

            ' ロットNO.入力処理
            'frmObj = New FormDataSelect(Me)                ' Form生成
            frmAutoObj.ShowDialog()                         ' ロットNO.入力
            Rtn = frmAutoObj.sGetReturn()
            'frmAutoObj.Close()                             ' Formアンロード
            '
            Call COVERLATCH_CLEAR()                         'カバー開ラッチクリア V2.2.0.035 

            If Rtn = cFRS_ERR_START Then
                Call Me.System1.OperationLogging(gSysPrm, "自動運転ＯＫ", "MANUAL")   'V2.0.0.2①
                Call UserSub.SetStartCheckStatus(True)                  ' 設定画面の確認有効化'V2.0.0.2①ロットスタート処理の中で初期化処理をさせる為にここでも設定する
                'V2.0.0.2①                If UserSub.IsSpecialTrimType() Then             ' ユーザプログラム特殊処理
                If Not UserBas.LotStartSetting() Then       ' ロットスタート時処理（印刷データヘッダー情報作成など）
                    frmAutoObj.gbFgAutoOperation = False     ' 自動運転解除
                    Me.AutoRunnningDisp.BackColor = System.Drawing.Color.Yellow  ' 背景色 = 黄色
                    Me.AutoRunnningDisp.Text = "自動運転解除中"
                End If
                'V2.0.0.2①            End If                '

                ''V2.2.0.0⑤↓
                ' TLF製ローダの場合自動運転切り替えを出力する
                If giLoaderType = 1 Then
                    Call Sub_ATLDSET(clsLoaderIf.LOUT_AUTO Or clsLoaderIf.LOUT_STOP, clsLoaderIf.LOUT_REDY)                    'V1.2.0.0④ ローダー出力(ON=自動,OFF=なし)
                    ObjLoader.gbIniFlg = 0
                End If
                ''V2.2.0.0⑤↑

                If frmAutoObj.gbFgAutoOperation Then
                    'If stUserData.iLotChange = 2 Or stUserData.iLotChange = 3 Then
                    Rtn = System1.ReadHostCommand_ForVBNET(giHostMode, gbHostConnected, giHostRun, giAppMode, pbLoadFlg)  ' ローダからデータを入力する
                    If (giHostMode <> cHOSTcMODEcAUTO) Then
                        ' ローダが自動に切り替わるまで待つ
                        Rtn = Me.System1.Form_Reset(cGMODE_LDR_CHK_AUTO, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
                        If (Rtn <= cFRS_ERR_EMG) Then
                            GoTo STP_ERR_EXIT                       ' 非常停止等のエラーならアプリ強制終了
                        ElseIf (Rtn = cFRS_ERR_RST) Then              ' キャンセルならコマンド終了
                            ''V2.2.0.0⑤↓
                            ' TLF製ローダの場合自動運転切り替えを出力する
                            If giLoaderType = 1 Then
                                Call Sub_ATLDSET(0, clsLoaderIf.LOUT_AUTO)                    'V1.2.0.0④ ローダー出力(ON=自動,OFF=なし)
                                ObjLoader.gbIniFlg = 0
                            End If
                            ''V2.2.0.0⑤↑
                            frmAutoObj.SetAutoOpeCancel(True)               ' V2.2.1.1②
                            Call frmAutoObj.AutoOperationEnd()              ' V2.2.1.1②

                            frmAutoObj.gbFgAutoOperation = False
                            Me.AutoRunnningDisp.BackColor = System.Drawing.Color.Yellow  ' 背景色 = 黄色
                            Me.AutoRunnningDisp.Text = "自動運転解除中"
                            GoTo STP_END
                        End If
                    End If
                    'End If
                    Call UserSub.SetStartCheckStatus(True)                  'V1.2.0.0④ 設定画面の確認有効化
                    ''V2.2.0.0⑤↓
                    SetAutoOpeStartTime()
                    If giLoaderType = 1 Then
                        ' TLF製ローダの場合自動運転切り替えを出力する
                        Call Sub_ATLDSET(clsLoaderIf.LOUT_AUTO Or clsLoaderIf.LOUT_STOP, clsLoaderIf.LOUT_REDY)                    'V1.2.0.0④ ローダー出力(ON=自動,OFF=なし)
                        ObjLoader.gbIniFlg = 0
                        ' 電磁ロック(観音扉右側ロック)を解除する
                        Rtn = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_ON)
                        If (Rtn = cFRS_TO_EXLOCK) Then                                ' 「前面扉ロック解除タイムアウト」なら戻り値を「RESET」にする

                            GoTo STP_END
                        End If

                    Else
                        Call Sub_ATLDSET(0, COM_STS_LOT_END)                    'V1.2.0.0④ ローダー出力(ON=なし,OFF=ロット終了)
                    End If
                    ''V2.2.0.0⑤↑
                End If
            Else                                                                            'V2.0.0.2①
                Call Me.System1.OperationLogging(gSysPrm, "自動運転キャンセル", "MANUAL")   'V2.0.0.2①
            End If

STP_END:
            'V2.2.0.0⑤↓
            If giLoaderType = 1 Then
                ChkLoaderInfoDisp(1)
            End If
            'V2.2.0.0⑤↑

            Call ZCONRST()                                              ' ｺﾝｿｰﾙｷｰ ﾗｯﾁ解除
            giAppMode = APP_MODE_IDLE                                   ' ｱﾌﾟﾘﾓｰﾄﾞ = トリマ装置アイドル中
            Timer1.Enabled = True                                       ' 監視タイマー開始
            Call Sub_ATLDSET(0, COM_STS_TRM_STATE)                      ' ローダー出力(ON=なし,OFF=トリマ動作中)
            Exit Sub

            ' ｿﾌﾄ強制終了処理
STP_ERR_EXIT:
            Call AppEndDataSave()                                       ' ｿﾌﾄ強制終了時のﾃﾞｰﾀ保存確認
            Call AplicationForcedEnding()                               ' ｿﾌﾄ強制終了処理
            End
            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "cmdLotchg_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        Finally
            giAppMode = APP_MODE_IDLE                                   ' ｱﾌﾟﾘﾓｰﾄﾞ = トリマ装置アイドル中
        End Try
    End Sub
#End Region

#Region "プローブボタン押下時処理"
    '''=========================================================================
    ''' <summary>プローブボタン押下時処理</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    '''=========================================================================
    Private Sub cmdProbeTeaching_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdProbeTeaching.Click

        Dim strMSG As String

        Try
            ' プローブコマンドを実行する
            cmdProbeTeaching_Proc()

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "cmdProbeTeaching_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "プローブコマンド実行"
    '''=========================================================================
    ''' <summary>プローブコマンド実行</summary>
    '''=========================================================================
    Public Sub cmdProbeTeaching_Proc()

        Dim r As Short
        Dim strMSG As String

        Try
            ' コマンド実行前処理
            r = Sub_cmdInit_Proc(APP_MODE_PROBE, F_PROBE)
            If (r = cFRS_ERR_RST) Then                                  ' Cancel(RESETｷｰ) ?
                If (gSysPrm.stSPF.giWithStartSw = 0) Then                ' ｽﾀｰﾄSW押下待ち(ｵﾌﾟｼｮﾝ)でない ?
                    GoTo STP_END
                Else                                                    ' ｽﾀｰﾄSW押下待ち(ｵﾌﾟｼｮﾝ)
                    GoTo STP_TRM
                End If
            End If
            If (r < cFRS_NORMAL) Then                                   ' エラー ?
                ' カバー開検出(ｽﾀｰﾄSW押下待ち(ｵﾌﾟｼｮﾝ)時発生する)
                If (r = cFRS_ERR_CVR) Or (r = cFRS_ERR_SCVR) Or (r = cFRS_ERR_LATCH) Then
                    r = cFRS_NORMAL
                    GoTo STP_TRM                                        ' カバー開検出ならｽﾗｲﾄﾞｶﾊﾞｰｸﾛｽﾞしてREADY状態へ
                Else
                    GoTo STP_END
                End If
            End If

            ' プローブティーチング処理
            r = User_ProbeTeaching()                                    ' プローブティーチング

            ' コマンド実行後処理
STP_TRM:
            Call Sub_cmdTerm_Proc(APP_MODE_PROBE, r)

            ' コマンド終了処理
STP_END:
            Call Sub_cmdEnd_Proc()

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "cmdProbeTeaching_Proc() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "TEACHボタン押下時処理"
    '''=========================================================================
    ''' <summary>TEACHボタン押下時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub cmdTeaching_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTeaching.Click

        Dim strMSG As String

        Try
            ' スタートポジション ティーチングを行う
            cmdTeaching_Proc()

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "BtnTEACH_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "スタートポジション ティーチングを行う"
    '''=========================================================================
    ''' <summary>スタートポジション ティーチングを行う</summary>
    '''=========================================================================
    Public Sub cmdTeaching_Proc()

        Dim r As Short
        'Dim ObjGazou As Process = Nothing                               ' Processオブジェクト
        Dim strMSG As String

        Try
            ' コマンド実行前処理
            r = Sub_cmdInit_Proc(APP_MODE_TEACH, F_TEACH)
            If (r = cFRS_ERR_RST) Then                                  ' Cancel(RESETｷｰ) ?
                If (gSysPrm.stSPF.giWithStartSw = 0) Then                ' ｽﾀｰﾄSW押下待ち(ｵﾌﾟｼｮﾝ)でない ?
                    GoTo STP_END
                Else                                                    ' ｽﾀｰﾄSW押下待ち(ｵﾌﾟｼｮﾝ)
                    GoTo STP_TRM
                End If
            End If
            If (r < cFRS_NORMAL) Then                                   ' エラー ?
                ' カバー開検出(ｽﾀｰﾄSW押下待ち(ｵﾌﾟｼｮﾝ)時発生する)
                If (r = cFRS_ERR_CVR) Or (r = cFRS_ERR_SCVR) Or (r = cFRS_ERR_LATCH) Then
                    r = cFRS_NORMAL
                    GoTo STP_TRM                                        ' カバー開検出ならｽﾗｲﾄﾞｶﾊﾞｰｸﾛｽﾞしてREADY状態へ
                Else
                    GoTo STP_END
                End If
            End If

            '-----------------------------------------------------------------------------
            '   ティーチング処理を行う
            '-----------------------------------------------------------------------------
            gbInitialized = False

            ' 画像表示プログラムを起動する(カットトレース用)
            ' ※画像表示プログラムの起動はOcxTeachで行うためここでは起動しない
            'r = Exec_GazouProc(ObjGazou, DISPGAZOU_PATH, DISPGAZOU_WRK, 0)

            Call ObjSys.Ilum_Ctrl(gSysPrm, Z1, ZOPT)                ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ点灯(制御有時有効)

            ' ティーチング画面処理
            r = User_teaching()

            Call ObjSys.Ilum_Ctrl(gSysPrm, Z0, ZOPT)                ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ消灯(制御有時有効)

            ' 画像表示プログラムを終了する
            'End_GazouProc(ObjGazou)

            ' コマンド実行後処理
STP_TRM:
            Call Sub_cmdTerm_Proc(APP_MODE_TEACH, r)

            ' コマンド終了処理
STP_END:

            Call Sub_cmdEnd_Proc()

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "cmdTeaching_Proc() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "カット位置補正ボタン押下時処理"
    '''=========================================================================
    ''' <summary>カット位置補正ボタン押下時処理</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub cmdCutPosTeach_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCutPosTeach.Click

        Dim r As Short
        Dim strMSG As String

        Try
            ' コマンド実行前処理
            r = Sub_cmdInit_Proc(APP_MODE_CUTPOS, F_CUTPOS)
            If (r = cFRS_ERR_RST) Then                                  ' Cancel(RESETｷｰ) ?
                If (gSysPrm.stSPF.giWithStartSw = 0) Then                ' ｽﾀｰﾄSW押下待ち(ｵﾌﾟｼｮﾝ)でない ?
                    GoTo STP_END
                Else                                                    ' ｽﾀｰﾄSW押下待ち(ｵﾌﾟｼｮﾝ)
                    GoTo STP_TRM
                End If
            End If
            If (r = cFRS_FNG_CPOS) Then                                 ' カット位置補正対象の抵抗がない ?
                GoTo STP_TRM
            End If

            If (r < cFRS_NORMAL) Then                                   ' エラー ?
                ' カバー開検出(ｽﾀｰﾄSW押下待ち(ｵﾌﾟｼｮﾝ)時発生する)
                If (r = cFRS_ERR_CVR) Or (r = cFRS_ERR_SCVR) Or (r = cFRS_ERR_LATCH) Then
                    r = cFRS_NORMAL
                    GoTo STP_TRM                                        ' カバー開検出ならｽﾗｲﾄﾞｶﾊﾞｰｸﾛｽﾞしてREADY状態へ
                Else
                    GoTo STP_END
                End If
            End If

            ' カット位置補正の為の画像登録処理を行う
            gbInitialized = False
            ChDir(My.Application.Info.DirectoryPath)
            r = User_CutpositionTeach()                                 ' 画像登録処理

            ' コマンド実行後処理
STP_TRM:
            Call Sub_cmdTerm_Proc(APP_MODE_CUTPOS, r)

            ' コマンド終了処理
STP_END:
            Call Sub_cmdEnd_Proc()

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "cmdCutPosTeach_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "RECOGボタン押下時処理"
    '''=========================================================================
    ''' <summary>RECOGボタン押下時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnRECOG_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnRECOG.Click

        Dim strMSG As String

        Try
            ' RECOG処理を行う
            BtnRECOG_Proc()

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "BtnRECOG_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "RECOG処理を行う"
    '''=========================================================================
    ''' <summary>RECOG処理を行う</summary>
    '''=========================================================================
    Public Sub BtnRECOG_Proc()

        Dim r As Short
        Dim strMSG As String

        Try
            ' コマンド実行前処理
            r = Sub_cmdInit_Proc(APP_MODE_RECOG, F_RECOG)
            If (r = cFRS_ERR_RST) Then                                  ' Cancel(RESETｷｰ) ?
                If (gSysPrm.stSPF.giWithStartSw = 0) Then                ' ｽﾀｰﾄSW押下待ち(ｵﾌﾟｼｮﾝ)でない ?
                    GoTo STP_END
                Else                                                    ' ｽﾀｰﾄSW押下待ち(ｵﾌﾟｼｮﾝ)
                    GoTo STP_TRM
                End If
            End If
            If (r < cFRS_NORMAL) Then                                   ' エラー ?
                ' カバー開検出(ｽﾀｰﾄSW押下待ち(ｵﾌﾟｼｮﾝ)時発生する)
                If (r = cFRS_ERR_CVR) Or (r = cFRS_ERR_SCVR) Or (r = cFRS_ERR_LATCH) Then
                    r = cFRS_NORMAL
                    GoTo STP_TRM                                        ' カバー開検出ならｽﾗｲﾄﾞｶﾊﾞｰｸﾛｽﾞしてREADY状態へ
                Else
                    GoTo STP_END
                End If
            End If

            Call ObjSys.Ilum_Ctrl(gSysPrm, Z1, ZOPT)                    ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ点灯(制御有時有効)

            ' θ補正の為の画像登録処理を行う
            r = User_PatternTeach()                                     ' 画像登録処理

            ' コマンド実行後処理
STP_TRM:
            Call Sub_cmdTerm_Proc(APP_MODE_RECOG, r)

            ' コマンド終了処理
STP_END:
            Call Sub_cmdEnd_Proc()

            Call Me.System1.Ilum_Ctrl(gSysPrm, Z0, ZOPT)            ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ消灯(制御有時有効)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "BtnRECOG_Proc() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "ＴＸボタン押下時処理"
    '''=========================================================================
    ''' <summary>
    ''' ＴＸボタン押下時処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub CmdTx_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdTx.Click
        Try
            TxTyTeach_Proc(APP_MODE_TX)
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Me)
        End Try
    End Sub
#End Region

#Region "ＴＹボタン押下時処理"
    '''=========================================================================
    ''' <summary>
    ''' ＴＹボタン押下時処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub CmdTy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdTy.Click
        Try
            TxTyTeach_Proc(APP_MODE_TY)
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Me)
        End Try
    End Sub
#End Region

#Region "ＴＸ、ＴＹ処理の実行"
    '''=========================================================================
    ''' <summary>
    ''' ＴＸ、ＴＹ処理の実行
    ''' </summary>
    ''' <param name="AppMode">APP_MODE_TXまたはAPP_MODE_TYのモード</param>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub TxTyTeach_Proc(ByVal AppMode As Short)

        Dim r As Short
        Dim strMSG As String
        Dim FncIdx As Short

        Try
            ' コマンド実行前処理
            If AppMode = APP_MODE_TX Then
                FncIdx = F_TX
            Else
                FncIdx = F_TY
            End If

            r = Sub_cmdInit_Proc(AppMode, FncIdx)
            If (r = cFRS_ERR_RST) Then                                  ' Cancel(RESETｷｰ) ?
                If (gSysPrm.stSPF.giWithStartSw = 0) Then                ' ｽﾀｰﾄSW押下待ち(ｵﾌﾟｼｮﾝ)でない ?
                    GoTo STP_END
                Else                                                    ' ｽﾀｰﾄSW押下待ち(ｵﾌﾟｼｮﾝ)
                    GoTo STP_TRM
                End If
            End If
            If (r < cFRS_NORMAL) Then                                   ' エラー ?
                ' カバー開検出(ｽﾀｰﾄSW押下待ち(ｵﾌﾟｼｮﾝ)時発生する)
                If (r = cFRS_ERR_CVR) Or (r = cFRS_ERR_SCVR) Or (r = cFRS_ERR_LATCH) Then
                    r = cFRS_NORMAL
                    GoTo STP_TRM                                        ' カバー開検出ならｽﾗｲﾄﾞｶﾊﾞｰｸﾛｽﾞしてREADY状態へ
                Else
                    GoTo STP_END
                End If
            End If

            Call ObjSys.Ilum_Ctrl(gSysPrm, Z1, ZOPT)                    ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ点灯(制御有時有効)

            Call ZCONRST()                                              ' ｺﾝｿｰﾙSWﾗｯﾁ解除

            ' θ補正の為の画像登録処理を行う
            r = User_TxTyTeach()                                        ' ＴＸ，ＴＹ処理

            ' コマンド実行後処理
STP_TRM:
            Call Sub_cmdTerm_Proc(AppMode, r)

            ' コマンド終了処理
STP_END:

            Call Sub_cmdEnd_Proc()

            Call Me.System1.Ilum_Ctrl(gSysPrm, Z0, ZOPT)            ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ消灯(制御有時有効)

            If r = cFRS_TxTy Then
                ' スタートポジション ティーチングを行う
                cmdTeaching_Proc()
            End If

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "TxTyTeach_Proc() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region
    '========================================================================================
    '   各コマンド実行前/後処理
    '========================================================================================
#Region "コマンド実行前処理"
    '''=========================================================================
    '''<summary>コマンド実行前処理</summary>
    '''<param name="gSts">  ｱﾌﾟﾘﾓｰﾄﾞ(giAppMode参照)</param>
    '''<param name="FncIdx">機能選択定義テーブルのｲﾝﾃﾞｯｸｽ</param>
    ''' <returns>  0  = 正常
    '''            3  = Reset SW押下
    '''           -80 = データ未ロード
    '''           -81 = 他コマンド実行中
    '''           -82 = ﾊﾟｽﾜｰﾄﾞ入力ｴﾗｰ
    ''' </returns>
    '''<remarks>非常停止/集塵機異常時は当関数内でｿﾌﾄ強制終了する</remarks>
    '''=========================================================================
    Private Function Sub_cmdInit_Proc(ByRef gSts As Short, ByRef FncIdx As Short) As Short

        Dim r As Short
        Dim InterlockSts As Integer
        Dim SwitchSts As Long
        Dim s As String
        Dim strMSG As String
        Dim iRtn As Integer

        Try
            ' トリマ動作中信号ON送信(ｵｰﾄﾛｰﾀﾞｰ)
            Call Sub_ATLDSET(COM_STS_TRM_STATE, 0)                      ' ローダー出力(ON=トリマ動作中,OFF=なし)

            ' 操作ログ出力
            Call Sub_OprLog(gSts)

            ' パスワード入力(特注)
            r = Func_Password(FncIdx)
            If (r <> True) Then                                         ' ﾊﾟｽﾜｰﾄﾞ入力ｴﾗｰ ?
                Return (cFRS_FNG_PASS)                                  ' Return値 = ﾊﾟｽﾜｰﾄﾞ入力ｴﾗｰ
            End If

            ' 他コマンド実行中 ?
            If giAppMode <> APP_MODE_IDLE Then                          ' 他コマンド実行中 ?
                Return (cFRS_FNG_CMD)                                   ' Return値 = 他コマンド実行中
            End If

            '' コマンド実行前のチェック
            'r = CmdExec_Check(gSts)
            'If (r <> cFRS_NORMAL) Then                                  ' チェックエラー ?
            '    Return (r)                                              ' Return値 = チェックエラー(Cancel(RESETｷｰ)を返す)
            'End If

            giAppMode = gSts                                            ' ｱﾌﾟﾘﾓｰﾄﾞ設定

            ' データロード済みチェック
            If (pbLoadFlg = False) Then                                 ' データ未ロード ?
                s = MSG_DataNotLoad                                     ' ﾃﾞｰﾀ未ﾛｰﾄﾞ
                Call Z_PRINT(s)                                         ' "Data is not loaded. Please Load the data file."
                Call Beep()
                Return (cFRS_FNG_DATA)                                  ' Return値 = ﾃﾞｰﾀ未ﾛｰﾄﾞ
                Exit Function
            End If

            ' コマンド実行前のチェック
            r = CmdExec_Check(giAppMode)
            If (r <> cFRS_NORMAL) Then                                  ' チェックエラー ?
                Return (r)                                              ' Return値 = チェックエラー(Cancel(RESETｷｰ)を返す)
            End If

            ' 集塵機異常チェック
            r = Me.System1.CheckDustVaccumeAlarm(gSysPrm)
            If (r <> 0) Then                                            ' エラーなら集塵機異常検出メッセージ表示
                Call Me.System1.Form_Reset(cGMODE_ERR_DUST, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
                GoTo STP_ERR_EXIT                                       ' ｿﾌﾄ強制終了
            End If

            '' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ点灯(制御有時有効)
            'If (gSts <> APP_MODE_LASER) Then                            ' LASERｺﾏﾝﾄﾞは点灯しない
            '    Call Me.System1.Ilum_Ctrl(gSysPrm, Z1, ZOPT)
            'End If

            ' 操作確認画面(START/RESET待ち)
            r = Me.System1.Form_Reset(cGMODE_START_RESET, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
            Me.Refresh()
            If (r = cFRS_ERR_RST) Then                                  ' Reset SW押下 ?
                Return (cFRS_ERR_RST)                                   ' Return値 = Reset SW押下
            End If
            If (r = cFRS_NORMAL) Or (r = cFRS_ERR_START) Then           ' 正常 ?
                ' クランプ/吸着ON
                r = System1.ClampVacume_Ctrl(gSysPrm, 1, giAppMode, 1)
                If (r <> cFRS_NORMAL) Then GoTo STP_ERR_EXIT

                ' シグナルタワー黄点灯(ティーチング中) 
                ' ※但しインターロック解除中(黄点滅)優先
                r = INTERLOCK_CHECK(InterlockSts, SwitchSts)
                If (InterlockSts = INTERLOCK_STS_DISABLE_NO) Then       ' インターロック中なら黄点灯
                    r = Me.System1.SetSignalTower(SIGOUT_YLW_ON, &HFFFF)
                End If
            ElseIf (r <= cFRS_ERR_EMG) Then                             ' ｴﾗｰ(非常停止等)ならｿﾌﾄ強制終了
                GoTo STP_ERR_EXIT
            Else
                Sub_cmdInit_Proc = r                                    ' Return値設定
            End If

            Call UserSub.ClampVacumeChange()         'V2.0.0.0⑭

            ' ボタン等を非表示にする
            cmdHelp.Visible = False                                     ' Versionボタン非表示 
            Me.Grpcmds.Visible = False                                  ' コマンドボタングループボックス非表示
            Me.GrpMode.Visible = False                                  ' ディジタルSWグループボックス非表示
            Me.frmInfo.Visible = False                                  ' 結果表示域非表示
            BtnStartPosSet.Enabled = False                              'V2.0.0.0②
            gbInitialized = False
            ButtonLaserCalibration.Visible = False                      'V2.1.0.0②
            btnCutStop.Visible = False                                  'V2.2.0.0⑥
            btnLoaderInfo.Visible = False                               'V2.2.0.0⑤

            Timer1.Enabled = False          ' @@@888 

            SetMagnifyBar(True)                                         ' V2.2.0.0①

            ChkLoaderInfoDisp(0)                              'V2.2.0.0⑤

            If gSts = APP_MODE_LASER Then
                ' Zを原点へ移動
                iRtn = EX_ZMOVE(0)
                If (iRtn <> cFRS_NORMAL) Then                              ' エラー ?(メッセージは表示済み) 
                    Call Me.System1.TrmMsgBox(gSysPrm, "SETZOFFPOS Ｚ軸原点位置移動が異常終了しました。", MsgBoxStyle.OkOnly, My.Application.Info.Title)
                    Return (cERR_END)
                End If
            Else
                iRtn = SetZOff_Prob_Off()                        ' INTIME内部の待機位置を変更して待機位置移動する。
                If iRtn <> cFRS_NORMAL Then
                    Call Me.System1.TrmMsgBox(gSysPrm, "SETZOFFPOS Ｚ軸待機位置変更が異常終了しました。", MsgBoxStyle.OkOnly, My.Application.Info.Title)
                    Return (cERR_END)
                End If
            End If

            Return (cFRS_NORMAL)

            ' ｿﾌﾄ強制終了処理
STP_ERR_EXIT:
            Call AppEndDataSave()                                       ' ｿﾌﾄ強制終了時のﾃﾞｰﾀ保存確認
            Call AplicationForcedEnding()                               ' ｿﾌﾄ強制終了処理
            End
            Return (r)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Sub_cmdInit_Proc() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Function
#End Region

#Region "各コマンド実行前のチェック処理"
    '''=========================================================================
    ''' <summary>各コマンド実行前のチェック処理</summary>
    ''' <param name="iAppMode">(INP) ｱﾌﾟﾘﾓｰﾄﾞ(giAppMode参照)</param>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          cFRS_FNG_CPOS = カット位置補正対象の抵抗がない
    '''          上記以外のエラー
    ''' </returns>
    '''=========================================================================
    Private Function CmdExec_Check(ByRef iAppMode As Short) As Short

        Dim bFlg As Boolean
        Dim Rn As Integer
        Dim RtnCode As Short
        Dim strMSG As String

        Try
            '-------------------------------------------------------------------
            '   コマンド実行前のチェックを行う
            '-------------------------------------------------------------------
            ' カット位置補正コマンド時
            If (iAppMode = APP_MODE_CUTPOS) Then
                ' パターン登録データがあるかチェックする
                bFlg = False
                For Rn = 1 To stPLT.PtnCount                            ' パターン登録数分繰り返す
                    ' パターン登録あり ?
                    'V1.0.4.3⑥                    If (stPTN(Rn).PtnFlg <> CUT_PATTERN_NONE And stPTN(Rn).PtnFlg <> 3) Then
                    If (stPTN(Rn).PtnFlg <> CUT_PATTERN_NONE) Then
                        bFlg = True
                        Exit For
                    End If
                Next Rn

                ' パターン登録データがない場合は処理しない
                If (bFlg = False) Then
                    strMSG = MSG_153                                    ' "カット位置補正対象の抵抗がありません"
                    RtnCode = cFRS_FNG_CPOS                             ' Return値 = カット位置補正対象の抵抗がない
                    GoTo STP_ERR_EXIT                                   ' メッセージ表示後エラー戻り
                End If
            End If

            ' 'V2.2.0.0⑤ ↓
            If giLoaderType <> 0 Then
                If (iAppMode <> APP_MODE_LASER) Then
                    ' 載物台に基板がある事をチェックする(手動モード時(OPTION))
                    RtnCode = ObjLoader.Sub_SubstrateExistCheck(System1)
                    If (RtnCode <> cFRS_NORMAL) Then                              ' エラー ?
                        RtnCode = cFRS_ERR_RST
                        Return RtnCode
                    End If
                Else
                    'V2.2.0.023↓
                    ' 電磁ロック(観音扉右側ロック)をロックする
                    RtnCode = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_ON)                             ' 電磁ロック
                    If (RtnCode <> cFRS_NORMAL) Then                                  ' エラー ?(メッセージは表示済)
                        RtnCode = cFRS_ERR_RST
                        Return RtnCode
                    End If
                    'V2.2.0.023↑
                End If
            End If
            ' 'V2.2.0.0⑤ ↑


            Return (cFRS_NORMAL)                                        ' Return値 = 正常

            '-------------------------------------------------------------------
            '   メッセージ表示後エラー戻り
            '-------------------------------------------------------------------
STP_ERR_EXIT:
            MsgBox(strMSG, MsgBoxStyle.Exclamation)
            Return (RtnCode)                                            ' Return値 = チェックエラー

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "CmdExec_Check() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cFRS_ERR_RST)                                       ' Return値 = チェックエラー(Cancel(RESETｷｰ)を返す)
        End Try
    End Function
#End Region

#Region "コマンド実行後処理"
    '''=========================================================================
    '''<summary>コマンド実行後処理</summary>
    '''<param name="gSts">(INP) 画面ステータス(giAppMode参照)</param>
    '''<param name="sts"> (INP) コマンド実行ステータス(エラー番号)</param>
    '''<remarks>非常停止時等は当関数内でｿﾌﾄ強制終了する</remarks>
    '''=========================================================================
    Private Sub Sub_cmdTerm_Proc(ByRef gSts As Short, ByRef sts As Short)

        Dim r As Short
        Dim strMSG As String

        Try
            ' 各コマンド実行エラーならメッセージ表示
            If (sts < cFRS_NORMAL) Then                                 ' コマンド実行エラー ?
                If (sts = cFRS_ERR_PTN) Then                            ' 以下のトリミングNG等のエラーはソフト強制終了しない
                ElseIf (sts = cFRS_TRIM_NG) Then                        '  
                ElseIf (sts = cFRS_ERR_TRIM) Then
                ElseIf (sts = cFRS_ERR_PT2) Then
                ElseIf (sts = cFRS_FNG_CPOS) Then                       ' カット位置補正対象の抵抗がない ?
                    GoTo STP_END
                ElseIf (sts <= cFRS_ERR_EMG) Then                       ' ｿﾌﾄ強制終了 ?
                    ' クランプ/吸着OFF
                    r = Me.System1.ClampVacume_Ctrl(gSysPrm, 0, giAppMode, giTrimErr)
                    GoTo STP_ERR_EXIT                                   ' ｿﾌﾄ強制終了
                End If
            End If

            ' テーブル原点移動
            r = Me.System1.Form_Reset(cGMODE_ORG_MOVE, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
            If (r <= cFRS_ERR_EMG) Then GoTo STP_ERR_EXIT '             ' エラーならﾌﾄ強制終了

            'V2.1.0.1①            ' θ軸原点復帰(｢自動｣又は｢自動+微調整｣又は｢手動で補正なしでθﾊﾟﾗﾒｰﾀ=原点復帰指定｣時) ※θありの場合
            'V2.1.0.1①            If ((stThta.iPP30 = 0) And (gSysPrm.stDEV.giTheta <> 0)) Or _
            'V2.1.0.1①               ((stThta.iPP30 = 2) And (gSysPrm.stDEV.giTheta <> 0)) Or _
            'V2.1.0.1①               ((stThta.iPP30 = 1) And (stThta.iPP31 = 0) And (gSysPrm.stSPF.giThetaParam = 1) And (gSysPrm.stDEV.giTheta <> 0)) Then
            'V2.1.0.1①                Call ROUND4(0.0#)                                       ' θを原点に戻す
            'V2.1.0.1①            End If
            Call ROUND4(0.0#)                                           'V2.1.0.1① θを原点に戻す

            ' ｽﾗｲﾄﾞｶﾊﾞｰ自動ｵｰﾌﾟﾝ
            If (gSysPrm.stSPF.giWithStartSw = 0) Then                    ' ｽﾀｰﾄSW押下待ち(ｵﾌﾟｼｮﾝ)でない ?
                r = System1.Z_COPEN(gSysPrm, giAppMode, giTrimErr, False)
                If (r <= cFRS_ERR_EMG) Then GoTo STP_ERR_EXIT '         ' エラーならﾌﾄ強制終了
            Else
                ' ｲﾝﾀｰﾛｯｸ時ならｽﾗｲﾄﾞｶﾊﾞｰ開待ち
                If (Me.System1.InterLockSwRead() And BIT_INTERLOCK_DISABLE) = 0 Then
                    r = Me.System1.Form_Reset(cGMODE_OPT_END, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
                    If (r <= cFRS_ERR_EMG) Then GoTo STP_ERR_EXIT '     ' エラーならﾌﾄ強制終了
                End If
            End If

STP_END:
            ' ボタン等を表示する
            cmdHelp.Visible = True                                      ' Versionボタン表示 
            Me.Grpcmds.Visible = True                                   ' コマンドボタングループボックス表示
            Me.GrpMode.Visible = True                                   ' ディジタルSWグループボックス表示
            Me.frmInfo.Visible = True                                   ' 結果表示域表示
            BtnStartPosSet.Enabled = True                               'V2.0.0.0②
            gbInitialized = True                                        ' True=原点復帰済
            'V2.1.0.0②↓
            If UserSub.IsLaserCaribrarionUse() Then
                ButtonLaserCalibration.Visible = True
            End If
            'V2.1.0.0②↑
            'V2.2.0.0⑥↓
            If giCutStop <> 0 Then
                btnCutStop.Visible = True
            End If
            'V2.2.0.0⑥↑
            'V2.2.0.0⑤↓
            If giLoaderType <> 0 Then
                btnLoaderInfo.Visible = True
                ' 電磁ロック(観音扉右側ロック)を解除する
                r = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_OFF)
            End If
            'V2.2.0.0⑤↑

            SetMagnifyBar(False)                                         ' V2.2.0.0①

            Exit Sub

            ' ｿﾌﾄ強制終了処理
STP_ERR_EXIT:
            Call AppEndDataSave()                                       ' ｿﾌﾄ強制終了時のﾃﾞｰﾀ保存確認
            Call AplicationForcedEnding()                               ' ｿﾌﾄ強制終了処理
            End

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Sub_cmdTerm_Proc() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "コマンド終了処理"
    '''=========================================================================
    '''<summary>コマンド終了処理</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub Sub_cmdEnd_Proc()

        Dim r As Short
        Dim strMSG As String

        Try
            ' クランプ/吸着OFF
            r = Me.System1.ClampVacume_Ctrl(gSysPrm, 0, giAppMode, 1, False)
            r = Me.System1.SetSignalTower(0, &HFFFFS)                   ' シグナルタワー制御(On=なし, Off=全ﾋﾞｯﾄ) 

            ' 後処理
            Call Me.System1.sLampOnOff(LAMP_START, True)                ' STARTﾗﾝﾌﾟON
            Call Me.System1.sLampOnOff(LAMP_RESET, True)                ' RESETﾗﾝﾌﾟON
            Call Me.System1.sLampOnOff(LAMP_Z, False)                   ' PRBﾗﾝﾌﾟOFF

            Call ZCONRST()                                              ' ｺﾝｿｰﾙｷｰ ﾗｯﾁ解除
            giAppMode = APP_MODE_IDLE                                   ' ｱﾌﾟﾘﾓｰﾄﾞ = トリマ装置アイドル中
            Call Me.System1.Ilum_Ctrl(gSysPrm, Z0, ZOPT)                 ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ消灯(制御有時有効)
            Timer1.Enabled = True                                       ' 監視タイマー開始
            Call Sub_ATLDSET(0, COM_STS_TRM_STATE)                      ' ローダー出力(ON=なし,OFF=トリマ動作中)

            ' 
            Call Z_PRINT(" " & vbCrLf)

            ChkLoaderInfoDisp(1)                              'V2.2.0.0⑤

            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Sub_cmdEnd_Proc() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

    '========================================================================================
    '   その他のボタン押下時の処理
    '========================================================================================
#Region "LOGボタン押下時処理"
    '''=========================================================================
    ''' <summary>LOGボタン押下時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub CmdLog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim ExeFile As String
        Dim strFNAME As String
        Dim r As Double
        Dim strMSG As String

        Try
            ' 【ﾌｧｲﾙを開く】ﾀﾞｲｱﾛｸﾞ設定
            strFNAME = ""
            FileDlgOpen.FileName = ""
            FileDlgOpen.ShowReadOnly = False
            FileDlgOpen.CheckFileExists = True
            FileDlgOpen.CheckPathExists = True
            FileDlgOpen.InitialDirectory = "C:\TRIMDATA\LOG"
            FileDlgOpen.Filter = "*.LOG|*.LOG"

            ' 【ﾌｧｲﾙを開く】ﾀﾞｲｱﾛｸﾞ表示
            If (FileDlgOpen.ShowDialog() = Windows.Forms.DialogResult.OK) Then
                If (FileDlgOpen.FileName = "") Then Exit Sub
                strFNAME = FileDlgOpen.FileName                         ' ログファイル名設定
            End If

            '    ' ログファイルがなければNOP
            '    If (gsLogFileName = "") Then Exit Sub                  ' ログファイルがなければNOP
            '    strFName = gsLogFileName

            ' NOTEPADでログファイルを開く
#If cOFFLINEcDEBUG Then
            ExeFile = "notepad.exe " & strFNAME
#Else
            'ExeFile = "C:\WINNT\system32\notepad.exe " + strFNAME
            ExeFile = "notepad.exe " & strFNAME
#End If
            r = Shell(ExeFile, 1)

Cansel:

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "CmdLog_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "Clearボタン押下時処理"
    '''=========================================================================
    ''' <summary>Clearボタン押下時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub CmdClr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim r As Short
        Dim strMSG As String

        Try
            ' クリア確認メッセージ設定
            If (gSysPrm.stTMN.giMsgTyp = 0) Then
                strMSG = "   生産数を初期化しますか？   "
            Else
                strMSG = "   Are you sure to Clear Trimming Result ?   "
            End If

            ' クリア確認メッセージを表示する
            r = Me.System1.TrmMsgBox(gSysPrm, strMSG, MsgBoxStyle.OkCancel, cAPPcTITLE)

            ' Cancel(RESETｷｰ))ならEXIT
            If (r = cFRS_ERR_RST) Then Exit Sub

            ' 生産数クリア
            Call Disp_frmInfo(COUNTER.PRODUCT_INIT, COUNTER.NONE)                                    ' 生産数初期化(frmInfo画面も再表示)
            Call PutLotInf()                                            ' ロット情報セーブ

            'V2.0.0.0⑨↓
            If (Not gObjFrmDistribute Is Nothing) Then                  ' 分布図データクリア
                gObjFrmDistribute.ClearCounter()
            End If
            'V2.0.0.0⑨↑

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "CmdClr_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "バージョンダイアログボックスの表示"
    '''=========================================================================
    ''' <summary>バージョンダイアログボックスの表示</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
        Dim r As Short
        Dim pstPRM As DllAbout.HelpVersion.HelpVer_PARAM            ' バージョン情報表示関数用パラメータ(OCXで定義)
        Dim strVER(3) As String
        Dim strMSG As String
        Dim EqType As String ''V2.2.0.0⑩

        Try
            ' ボタン等を非表示にする
            cmdHelp.Visible = False                                     ' Versionボタン非表示 
            Me.Grpcmds.Visible = False                                  ' コマンドボタングループボックス非表示
            Me.GrpMode.Visible = False                                  ' ディジタルSWグループボックス非表示
            Me.frmInfo.Visible = False                                  ' 結果表示域非表示

            ' 構造体pstPRMの配列の初期化 ※配列の要素数はOcxAbout.ocxで定義と同じにする必要あり
            pstPRM.strTtl = New String(4) {}
            pstPRM.strModule = New String(20) {}
            pstPRM.strVer = New String(20) {}


            'V2.2.0.0⑩↓
            Dim strVersion = GetPrivateProfileString_S("TMENU", "VERSION_NAME", SYSPARAMPATH, "")
            If strVersion.ToString().Trim <> "" Then
                EqType = strVersion
            Else
                EqType = gSysPrm.stTMN.gsKeimei
            End If
            'V2.2.0.0⑩↑


            ' バージョン情報表示関数用パラメータを設定する
            pstPRM.iTtlNum = 3                              ' タイトル文字列の数　'V2.2.0.0⑩
            pstPRM.strTtl(0) = My.Application.Info.Title    ' アプリ名 
            pstPRM.strTtl(1) = "LMP-" + EqType + gSysPrm.stDEV.gsDevice_No + "-000 " +
                               My.Application.Info.Version.Major.ToString("0") & "." &
                               My.Application.Info.Version.Minor.ToString("0") & "." &
                               My.Application.Info.Version.Build.ToString("0") & "." &
                               My.Application.Info.Version.Revision.ToString("0")
            pstPRM.strTtl(2) = "(c) TOWA LASERFRONT CORP."

            pstPRM.iVerNum = 15                             ' バージョン情報の数
            pstPRM.strModule(0) = "RT MODULE"               ' 1."RT MODULE"
            pstPRM.strVer(0) = DLL_PATH + "INTRIM_SL432.rta"
            pstPRM.strModule(1) = "DllTrimFnc.dll"          ' 2."DllTrimFnc.dll"
            pstPRM.strVer(1) = DLL_PATH + pstPRM.strModule(1)
            pstPRM.strModule(2) = "DllSysPrm.dll"           ' 3."DllSysPrm.dll"
            pstPRM.strVer(2) = DLL_PATH + pstPRM.strModule(2)
            pstPRM.strModule(3) = "DllSystem.dll"           ' 4."DllSystem.dll"
            pstPRM.strVer(3) = DLL_PATH + pstPRM.strModule(3)
            pstPRM.strModule(4) = "DllAbout.dll"            ' 5."DllAbout.dll"
            pstPRM.strVer(4) = DLL_PATH + pstPRM.strModule(4)
            pstPRM.strModule(5) = "DllUtility.dll"          ' 6."DllUtility.dll"
            pstPRM.strVer(5) = DLL_PATH + pstPRM.strModule(5)
            pstPRM.strModule(6) = "DllLaserTeach.dll"       ' 7."DllLaserTeach.dll"
            pstPRM.strVer(6) = DLL_PATH + pstPRM.strModule(6)
            pstPRM.strModule(7) = "DllManualTeach.dll"      ' 8."DllManualTeach.dll"
            pstPRM.strVer(7) = DLL_PATH + pstPRM.strModule(7)
            pstPRM.strModule(8) = "DllPassword.dll"         ' 9."DllPassword.dll"
            pstPRM.strVer(8) = DLL_PATH + pstPRM.strModule(8)
            pstPRM.strModule(9) = "DllProbeTeach.dll"       '10."DllProbeTeach.dll"
            pstPRM.strVer(9) = DLL_PATH + pstPRM.strModule(9)
            pstPRM.strModule(10) = "DllTeach.dll"           '11."DllTeach.dll"
            pstPRM.strVer(10) = DLL_PATH + pstPRM.strModule(10)
            pstPRM.strModule(11) = "DllVideo.dll"           '12."DllVideo.dll"
            pstPRM.strVer(11) = DLL_PATH + pstPRM.strModule(11)

            ' 新Dll(C#で作成) 
            pstPRM.strModule(12) = "DllSerialIO.dll"        '13."DllSerialIO.dll"
            pstPRM.strVer(12) = DLL_PATH + pstPRM.strModule(12)
            pstPRM.strModule(13) = "DllCndXMLIO.dll"        '14."DllCndXMLIO.dll"
            pstPRM.strVer(13) = DLL_PATH + pstPRM.strModule(13)
            pstPRM.strModule(14) = "DllFLCom.dll"           '15."DllFLCom.dll"
            pstPRM.strVer(14) = DLL_PATH + pstPRM.strModule(14)

            ' バージョン情報表示位置を設定する
            HelpVersion1.Left = Text2.Location.X            ' Left = Text4位置 
            HelpVersion1.Top = cmdHelp.Location.Y           ' Top  = Versionﾎﾞﾀﾝ位置 

            ' バージョン情報表示
            HelpVersion1.Visible = True
            HelpVersion1.BringToFront()                     ' 最前面へ表示

            r = HelpVersion1.Version_Disp(pstPRM)
            HelpVersion1.Visible = False

            ' ボタン等を表示する
            cmdHelp.Visible = True                                      ' Versionボタン表示 
            Me.Grpcmds.Visible = True                                   ' コマンドボタングループボックス表示
            Me.GrpMode.Visible = True                                   ' ディジタルSWグループボックス表示
            Me.frmInfo.Visible = True                                   ' 結果表示域表示

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "cmdHelp_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "ADJボタン押下時処理"
    '''=========================================================================
    '''<summary>ADJボタン押下時処理</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub BtnADJ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnADJ.Click

        Dim strMSG As String

        Try
            If (BtnADJ.Text = "ADJ OFF") Then
                BtnADJ.Text = "ADJ ON"
                BtnADJ.BackColor = System.Drawing.Color.Yellow
                gbChkboxHalt = True
            Else
                BtnADJ.Text = "ADJ OFF"
                BtnADJ.BackColor = System.Drawing.SystemColors.Control
                gbChkboxHalt = False
            End If
            BtnADJ.Refresh()

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "BtnADJ_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
    ''' <summary>
    ''' ＡＤＪボタンのＯＮ、ＯＦＦ取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetBtnADJStatus() As Boolean
        Return (gbChkboxHalt)
    End Function
#End Region

#Region "Expansionﾎﾞﾀﾝ押下時処理(ﾛｸﾞ画面拡大)"
    '''=========================================================================
    ''' <summary>Expansionﾎﾞﾀﾝ押下時処理(ﾛｸﾞ画面拡大)</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub cmdExpansion_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExpansion.Click

        'If cmdExpansion.Text = "拡大表示" Or cmdExpansion.Text = "Expansion" Then
        '    ' 拡大ﾛｸﾞ画面
        '    txtLog.Top = 40
        '    txtLog.Height = 689
        '    cmdExpansion.Top = txtLog.Top - 22
        '    If (gSysPrm.stTMN.giMsgTyp = 0) Then
        '        cmdExpansion.Text = "通常表示"
        '    Else
        '        cmdExpansion.Text = "Normal"
        '    End If
        '    txtLog.Font = VB6.FontChangeSize(txtLog.Font, 10)
        '    txtLog.BringToFront()

        'Else
        '    ' 通常ﾛｸﾞ画面
        '    txtLog.Top = 544
        '    txtLog.Height = 192
        '    cmdExpansion.Top = txtLog.Top - 22
        '    If (gSysPrm.stTMN.giMsgTyp = 0) Then
        '        cmdExpansion.Text = "拡大表示"
        '    Else
        '        cmdExpansion.Text = "Expansion"
        '    End If
        '    txtLog.Font = VB6.FontChangeSize(txtLog.Font, gSysPrm.stLOG.gdLogTextFontSize)
        '    txtLog.SendToBack()
        'End If

        If cmdExpansion.Text = "拡大表示" Or cmdExpansion.Text = "Expansion" Then
            ' 拡大ﾛｸﾞ画面
            lstLog.Top = 40
            lstLog.Height = 689
            cmdExpansion.Top = lstLog.Top - 22
            If (gSysPrm.stTMN.giMsgTyp = 0) Then
                cmdExpansion.Text = "通常表示"
            Else
                cmdExpansion.Text = "Normal"
            End If
            lstLog.Font = VB6.FontChangeSize(lstLog.Font, 10)
            lstLog.BringToFront()

        Else
            ' 通常ﾛｸﾞ画面
            lstLog.Top = 544
            lstLog.Height = 192
            cmdExpansion.Top = lstLog.Top - 22
            If (gSysPrm.stTMN.giMsgTyp = 0) Then
                cmdExpansion.Text = "拡大表示"
            Else
                cmdExpansion.Text = "Expansion"
            End If
            lstLog.Font = VB6.FontChangeSize(lstLog.Font, gSysPrm.stLOG.gdLogTextFontSize)
            lstLog.SendToBack()
        End If

    End Sub
#End Region

#Region "Expansionﾎﾞﾀﾝ(有効/無効)"
    '''=========================================================================
    ''' <summary>Expansionﾎﾞﾀﾝ(有効/無効)</summary>
    ''' <param name="MODE">True=有効, False=無効</param>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub ExpansionOnOff(ByRef MODE As Boolean)

        If gSysPrm.stSPF.giDispCh = 1 Then
            If Not (Me.cmdExpansion.Visible) = MODE Then
                Me.cmdExpansion.Visible = MODE
                If MODE = True Then Me.cmdExpansion.BackColor = System.Drawing.ColorTranslator.FromOle(&HFF00)
                If MODE = False Then Me.cmdExpansion.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            End If
        End If

    End Sub
#End Region

#Region "ﾛｸﾞ画面切替"
    '''=========================================================================
    ''' <summary>ﾛｸﾞ画面切替</summary>
    ''' <param name="MODE">True=有効, False=無効</param>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub LogWindowCh(ByRef MODE As Short)

        If (gSysPrm.stSPF.giDispCh = 1) Then             ' 拡大表示する ?
            If MODE = 0 Then                            ' 通常ｻｲｽﾞ ?
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    Me.cmdExpansion.Text = "通常表示"
                Else
                    Me.cmdExpansion.Text = "Normal"
                End If

            ElseIf MODE = 1 Then                        ' 拡大ｻｲｽﾞ ?
                Me.cmdExpansion.Text = "拡大表示"
            End If
            Call Me.cmdExpansion_Click(Me.cmdExpansion, New System.EventArgs())
        End If

    End Sub
#End Region

#Region "ファンクションキー押下時処理"
    '''=========================================================================
    ''' <summary>ファンクションキー押下時処理</summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub Form1_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown

#If cKEYBOARDcUSE Then

        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000

        Dim ShiftDown As Boolean
        Dim AltDown As Boolean
        Dim CtrlDown As Boolean

        ' トリマ装置アイドル中でなければNOP
        If giAppMode Then
            If Not ((gbAdjOnStatus And KeyCode = System.Windows.Forms.Keys.F11) Or KeyCode = System.Windows.Forms.Keys.F12) Then
                'V2.2.0.032
                If giAppMode = APP_MODE_FINEADJ Or giAppMode = APP_MODE_TEACH Then
                    If (_jogKeyDown IsNot Nothing) Then         'V6.0.0.0⑩
                        _jogKeyDown.Invoke(eventArgs)
                    End If
                End If
                Exit Sub
            End If
        End If

        ShiftDown = (Shift And VB6.ShiftConstants.ShiftMask) > 0
        AltDown = (Shift And VB6.ShiftConstants.AltMask) > 0
        CtrlDown = (Shift And VB6.ShiftConstants.CtrlMask) > 0

        Select Case KeyCode
            Case System.Windows.Forms.Keys.F1
                Call cmdLotInfo_Click(cmdLotInfo, New System.EventArgs())   ' データ設定

            Case System.Windows.Forms.Keys.F2
                If (stFNC(F_LOAD).iDEF = 0) Then Exit Sub ' 選択不可ならEXIT
                Call cmdLoad_Click(cmdLoad, New System.EventArgs()) ' データロード

            Case System.Windows.Forms.Keys.F3
                If (stFNC(F_SAVE).iDEF = 0) Then Exit Sub ' 選択不可ならEXIT
                Call cmdSave_Click(cmdSave, New System.EventArgs()) ' データセーブ

            Case System.Windows.Forms.Keys.F4
                If (stFNC(F_EDIT).iDEF = 0) Then Exit Sub ' 選択不可ならEXIT
                Call cmdEdit_Click(cmdEdit, New System.EventArgs()) ' EDIT

            Case System.Windows.Forms.Keys.F5
                Call cmdLotchg_Click(cmdLotChg, New System.EventArgs())   ' 自動運転
                'Call cmdPrint_Click(cmdPrint, New System.EventArgs())   ' 印刷

            Case System.Windows.Forms.Keys.F6
                If (stFNC(F_PROBE).iDEF = 0) Then Exit Sub ' 選択不可ならEXIT
                Call cmdProbeTeaching_Click(cmdProbeTeaching, New System.EventArgs())   ' プローブ位置ティーチング
                'If (stFNC(F_LASER).iDEF = 0) Then Exit Sub ' 選択不可ならEXIT
                'Call cmdLaserTeach_Click(cmdLaserTeach, New System.EventArgs()) ' レーザ

            Case System.Windows.Forms.Keys.F7
                If (stFNC(F_TEACH).iDEF = 0) Then Exit Sub ' 選択不可ならEXIT
                Call cmdTeaching_Click(cmdTeaching, New System.EventArgs()) ' ティーチング(F8)

            Case System.Windows.Forms.Keys.F8
                If (stFNC(F_CUTPOS).iDEF = 0) Then Exit Sub ' 選択不可ならEXIT
                Call cmdCutPosTeach_Click(cmdCutPosTeach, New System.EventArgs()) ' カット位置ティーチング

            Case System.Windows.Forms.Keys.F9
                If (stFNC(F_RECOG).iDEF = 0) Then Exit Sub ' 選択不可ならEXIT
                Call BtnRECOG_Click(BtnRECOG, New System.EventArgs()) ' パターン登録

            Case System.Windows.Forms.Keys.F10
                '                Call cmdExit_Click(cmdExit, New System.EventArgs()) ' END(F11)

            Case System.Windows.Forms.Keys.F11
                CbDigSwL.Focus()        ' MoveMode の変更
                If CbDigSwL.SelectedIndex >= CbDigSwL.Items.Count - 1 Then
                    CbDigSwL.SelectedIndex = 0
                Else
                    CbDigSwL.SelectedIndex = CbDigSwL.SelectedIndex + 1
                End If

            Case System.Windows.Forms.Keys.F12
                Call BtnADJ_Click(eventSender, eventArgs)           ' ADJ ON/OFF
        End Select

        If (_jogKeyDown IsNot Nothing) Then         'V6.0.0.0⑩
            _jogKeyDown.Invoke(eventArgs)
        End If

#End If
    End Sub
#End Region

#Region "キーアップ時処理"          'V2.2.0.032
    Private Sub Form1_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        If (_jogKeyUp IsNot Nothing) Then
            _jogKeyUp.Invoke(e)
        End If
    End Sub
#End Region


    '========================================================================================
    '   タイマーイベント処理
    '========================================================================================
#Region "周期起動タイマー処理"
    '''=========================================================================
    ''' <summary>周期起動タイマー処理</summary>
    ''' <remarks></remarks>
    '''=========================================================================

    Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick

        Dim strMSG As String                                            ' ﾒｯｾｰｼﾞ編集域
        Dim r As Short
        Dim iRtn As Integer
        Dim swStatus As Integer
        Dim interlockStatus As Integer
        Dim sldCvrSts As Integer
        Dim LdIDat As UInteger                          'V1.2.0.0④
        Dim iAppMode As Short                           'V1.2.0.0④ ｱﾌﾟﾘﾓｰﾄﾞ
        Dim bAutoLoaderAuto As Boolean = False          'V2.0.0.0⑮ オートローダー自動手動フラグ
        Dim coverSts As Long                            'V2.2.0.0⑤

        Try
            '---------------------------------------------------------------------------
            '   初期処理
            '---------------------------------------------------------------------------
            Timer1.Enabled = False                                      ' 監視タイマー停止

            ' ｱｲﾄﾞﾙ/LOAD/SAVE/EDIT/LOTCHGｺﾏﾝﾄﾞ以外の場合(OCX使用ｺﾏﾝﾄﾞ)はﾀｲﾏｰ停止してそのまま抜けて
            ' OCXから返ってきたら監視ﾀｲﾏｰを開始する
            If (giAppMode <> APP_MODE_IDLE) And (giAppMode <> APP_MODE_LOAD) And
               (giAppMode <> APP_MODE_SAVE) And (giAppMode <> APP_MODE_LOTCHG) And
               (giAppMode <> APP_MODE_EDIT) And (giAppMode <> APP_MODE_LOTNO) Then
                Call ZCONRST()                                          ' コンソールキーラッチ解除
                Exit Sub
            End If

            '---------------------------------------------------------------------------
            '   監視処理開始
            '---------------------------------------------------------------------------		
            ' 非常停止等チェック(トリマ装置アイドル中)
            r = Me.System1.Sys_Err_Chk_EX(gSysPrm, giAppMode)           ' 非常停止/ｶﾊﾞｰ/ｴｱｰ圧/集塵機/ﾏｽﾀｰﾊﾞﾙﾌﾞﾁｪｯｸ
            If (r <> cFRS_NORMAL) Then                                  ' 非常停止等検出 ?
                GoTo TimerErr                                           ' アプリ強制終了
            End If

            '---------------------------------------------------------------------------
            '   インターロック状態取得
            '---------------------------------------------------------------------------
            r = DispInterLockSts()                                      ' インターロック状態の表示/非表示
            ''V2.2.0.0⑤↓
            '            If (r = INTERLOCK_STS_DISABLE_FULL) Then                    ' インターロック全解除 ?
            If (r <> INTERLOCK_STS_DISABLE_NO) Then                    ' インターロック解除でない ?
                ' インターロック解除スイッチONで、カバー閉は異常
                '    If (System1.InterLockSwRead() And BIT_COVER_CLOSE) Then
                '#If cATLcDEN = 1 Then
                '				Call Sub_ATLDSET(0, &HFFFF)                             ' 全てOFFとする。
                '#End If
                r = COVER_CHECK(coverSts)                           ' 固定カバー状態取得(0=固定カバー開, 1=固定カバー閉))
                    If (coverSts = 1) Then                              ' 固定カバー閉 ?
                        ' ハードウェアエラー(カバーが閉じてます)メッセージ表示
                        Call System1.Form_Reset(cGMODE_ERR_HW, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
                        GoTo TimerErr                                       ' アプリ強制終了
                    End If

                '    End If
            End If
            ''V2.2.0.0⑤↑

            ' IOモニタ表示(デバッグ用)
#If cIOcMONITORcENABLED = 1 Then                                        ' IOﾓﾆﾀ表示する ?
            ObjSys.Z_ATLDGET(LdIDat)                                        ' ローダー入力
            If (gwPrevHcmd <> LdIDat) Then                                  ' 前回データから変化があった？
                Call IoMonitor(LdIDat, 0)                                   ' IOﾓﾆﾀ表示
                gwPrevHcmd = LdIDat
            End If
#End If

            Dim bChangeManual As Boolean = False                        ' V2.2.0.0⑤ 

            '---------------------------------------------------------------------------
            '   ローダ自動ならローダからデータを入力する（コマンド受信）
            '---------------------------------------------------------------------------
            If (giAppMode = APP_MODE_IDLE) Then                         ' アイドル状態時にチェックする
                ' ローダ自動でローダ有りからローダ無しの変化はエラーとする
                If (giHostMode = cHOSTcMODEcAUTO) And (gbHostConnected = False) Then
                    ' ローダが手動&停止に切り替わるまで待つ
                    r = Me.System1.Form_Reset(cGMODE_LDR_ERR, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
                    If (r <= cFRS_ERR_EMG) Then GoTo TimerErr ' 非常停止等のエラーならアプリ強制終了
                End If

                ''V2.2.0.0⑤↓
                ' TLF製ローダの場合自動運転切り替えを出力する
                If giLoaderType = 1 Then
                    ' ローダからデータを入力する
                    r = System1.ReadHostCommand_ForVBNET(giHostMode, gbHostConnected, giHostRun, giAppMode, pbLoadFlg)

                    If frmAutoObj.gbFgAutoOperation = True Then
                        Call Btn_Enb_OnOff(0)                               ' ボタン非活性化
                        bAutoLoaderAuto = True

                        ' 必ずクランプ吸着有にする
                        ObjSys.setClampVaccumConfig(0)

                        ' ロット切り替えフラグのクリア
                        ObjLoader.SetLotChangeFlg(0)        'V2.2.1.1⑧ 

                        r = ObjLoader.LoaderGlassHandlingProc(Me.System1)

                        Dim tmptactTime As Double = ((gdTrimtime.Minutes * 60) + gdTrimtime.Seconds + (gdTrimtime.Milliseconds / 1000.0)) * 10
                        ObjSys.Sub_SetTrimmingTime(tmptactTime)

                        ' 基板交換時間書込み
                        Dim dummy As Integer
                        Dim SupplyMag As Integer = 0
                        Dim SupplySlot As Integer = 0
                        Dim StoreMag As Integer = 0
                        Dim StoreSlot As Integer = 0

                        ObjSys.Sub_GetProcessTime(gitacktTime, gichangePlateTime, dummy)
                        gichangePlateTime = gitacktTime - tmptactTime
                        ObjSys.Sub_SetChangePlateTime(gichangePlateTime)
                        objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_TACT, gitacktTime)
                        objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_EXCHANGE, gichangePlateTime)
                        objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_TRIMMING, tmptactTime)

                        ObjSys.Sub_GetNowProcessMgInfo(SupplyMag, SupplySlot, StoreMag, StoreSlot)
                        objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_SUPPLY_MAGAGINE, SupplyMag)
                        objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_SUPPLY_SLOT, SupplySlot)
                        'V2.2.0.037　objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_STORE_MAGAGINE, StoreMag)
                        'V2.2.0.037　objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_STORE_SLOT, StoreSlot)

                        If r = cFRS_ERR_EMG Then
                            '非常停止

                        ElseIf r = cFRS_ERR_LOTEND Then
                            ' 自動運転の終了
                            Call Sub_ATLDSET(0, clsLoaderIf.LOUT_AUTO Or clsLoaderIf.LOUT_SUPLY Or clsLoaderIf.LOUT_STS_RUN Or clsLoaderIf.LOUT_REQ_COLECT Or clsLoaderIf.LOUT_DISCHRAGE)                             ' ローダ出力(ON=基板要求または供給位置決完了+ﾄﾘﾏ停止中+他, OFF=供給位置決完了または基板要求)
                            frmAutoObj.gbFgAutoOperation = False

                            fStartTrim = False                       ' スタートTRIMフラグをOFF
                            Call Me.System1.ClampVacume_Ctrl(gSysPrm, 0, giAppMode, 0)

                            ' 電磁ロック(観音扉右側ロック)を解除する
                            r = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_OFF)

                            'V2.2.0.022↓
                            Call UserSub.LotEnd()                           ' ロット終了時のデータ出力
                            Call Printer.Print(False)                       ' ロット情報印刷
                            UserBas.stCounter.LotPrint = True               ' ロット終了時の印刷実行済みでTrue
                            'V2.2.0.022↑

                            DispMarkAlarmList()         ' マーク印字のエラーリストを画面に表示         V2.2.1.7③

                            ObjLoader.Loader_EndAutoDrive(Me.System1)
                            frmAutoObj.gbFgAutoOperation = False

                            ' マーク印字の場合、トリミングに戻す V2.2.1.7⑦↓
                            If UserSub.IsTrimType5() = True Then
                                CbDigSwL.SelectedIndex = 0
                            End If
                            'V2.2.1.7⑦↑

                            'V2.2.0.0⑯↓ 
                            stMultiBlock.gMultiBlock = 0
                            stMultiBlock.Initialize()
                            For i As Integer = 0 To 5
                                stMultiBlock.BLOCK_DATA(i).DataNo = i + 1           ' DataNo
                                stMultiBlock.BLOCK_DATA(i).Initialize()
                                stMultiBlock.BLOCK_DATA(i).gBlockCnt = 0            ' ブロック数
                            Next
                            ''V2.2.0.0⑯↑

                        ElseIf r = cFRS_ERR_RST Then
                            ' 中断等で次の基板は処理しない。 
                            fStartTrim = False                       ' スタートTRIMフラグをOFF
                            ' 　基板取り除きメッセージを表示する
                            r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE)

                            frmAutoObj.SetAutoOpeCancel(False)               ' V2.2.1.1②
                            Call frmAutoObj.AutoOperationEnd()              ' V2.2.1.1②
                            frmAutoObj.gbFgAutoOperation = False

                            r = sResetStart()
                            If (r <> cFRS_NORMAL) Then                          ' エラー ?
                                r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE2)
                                Call AppEndDataSave()                           ' ｿﾌﾄ強制終了時のﾃﾞｰﾀ保存確認
                                Call AplicationForcedEnding()                   ' ｿﾌﾄ強制終了処理
                                End                                             ' アプリ強制終了
                                Return
                            End If
                            ' 電磁ロック(観音扉右側ロック)を解除する
                            r = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_OFF)

                        ElseIf r = cFRS_ERR_LDRTO Or r = cFRS_ERR_LDR1 Or r = cFRS_ERR_LDR2 Or r = cFRS_ERR_LDR3 Then
                            ' 中断等で次の基板は処理しない。 
                            fStartTrim = False                       ' スタートTRIMフラグをOFF
                            ' 　基板取り除きメッセージを表示する
                            r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE)

                            frmAutoObj.SetAutoOpeCancel(False)               ' V2.2.1.1②
                            Call frmAutoObj.AutoOperationEnd()              ' V2.2.1.1②
                            frmAutoObj.gbFgAutoOperation = False

                            r = sResetStart()
                            If (r <> cFRS_NORMAL) Then                          ' エラー ?
                                r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE2)
                                Call AppEndDataSave()                           ' ｿﾌﾄ強制終了時のﾃﾞｰﾀ保存確認
                                Call AplicationForcedEnding()                   ' ｿﾌﾄ強制終了処理
                                End                                             ' アプリ強制終了
                                Return
                            End If
                            ' 電磁ロック(観音扉右側ロック)を解除する
                            r = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_OFF)

                        ElseIf r = cFRS_NORMAL Then
                            ' 基板をステージに載せたので1枚処理を開始する 
                            fStartTrim = True                       ' スタートTRIMフラグをON
                            ' Lot切り替え信号のチェック
                            ObjSys.Z_ATLDGET(LdIDat)                                        ' ローダー入力
                            If LdIDat = clsLoaderIf.LINP_TRM_LOTCHANGE_START Then
                                r = clsLoaderIf.LINP_TRM_LOTCHANGE_START
                            End If

                        Else
                            ' 中断等で次の基板は処理しない。 
                            fStartTrim = False                       ' スタートTRIMフラグをOFF
                            Call frmAutoObj.AutoOperationEnd()              ' V2.2.1.1②
                            frmAutoObj.gbFgAutoOperation = False

                            ' 　基板取り除きメッセージを表示する
                            r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE)

                            ' 電磁ロック(観音扉右側ロック)を解除する
                            r = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_OFF)

                        End If
                        ' トリミング実行後処理
                        Call System1.AutoLoaderFlgReset()                       ' オートローダーフラグリセット

                        If frmAutoObj.gbFgAutoOperation = False Then
                            Call Sub_ATLDSET(0, clsLoaderIf.LINP_AUTO)                    ' ローダー出力(ON=なし,OFF=自動)
                        End If
                    Else
                        Call Btn_Enb_OnOff(1)                               ' ボタン非活性化
                    End If
                Else

                    ' ﾃﾞﾊﾞｯｸﾞ用ﾛｰﾀﾞ入力ﾓｰﾄﾞ
#If cIOcHostComandcENABLED = 1 Then                                     ' ﾃﾞﾊﾞｯｸﾞ用ﾛｰﾀﾞ入力ﾓｰﾄﾞ ? 
        ' ホスト接続状態設定
        If ((gDebugHostCmd And cHSTcRDY) = cHSTcRDY) Then
            gbHostConnected = True                                      ' ホスト接続状態(True=接続(ﾛｰﾀﾞ有))
        Else
            gbHostConnected = False                                     ' ホスト接続状態(False=未接続(ﾛｰﾀﾞ無))
        End If
        ' ﾛｰﾀﾞﾓｰﾄﾞ設定
        If ((gDebugHostCmd And cHSTcAUTO) = cHSTcAUTO) Then
            giHostMode = cHOSTcMODEcAUTO                                ' ﾛｰﾀﾞﾓｰﾄﾞ(1:自動ﾓｰﾄﾞ)
        Else
            giHostMode = cHOSTcMODEcMANUAL                              ' ﾛｰﾀﾞﾓｰﾄﾞ(0:手動ﾓｰﾄﾞ)
        End If
        ' ﾛｰﾀﾞ動作中設定
        If ((gDebugHostCmd And cHSTcSTATE) = cHSTcSTATE) Then
            giHostRun = 0                                               ' ﾛｰﾀﾞ動作中(0:停止)
        Else
            giHostRun = 1                                               ' ﾛｰﾀﾞ動作中(1:動作中)
        End If
        ' ﾄﾘﾏｰｽﾀｰﾄ設定
        If ((gDebugHostCmd And cHSTcTRMCMD) = cHSTcTRMCMD) Then
            r = cHSTcTRMCMD                                             ' ﾄﾘﾏｰｽﾀｰﾄ
        End If

        ' ﾃﾞﾊﾞｯｸﾞ用ｺﾏﾝﾄﾞ設定
        LdIDat = gDebugHostCmd
#End If

                    ' ローダからデータを入力する
                    r = System1.ReadHostCommand_ForVBNET(giHostMode, gbHostConnected, giHostRun, giAppMode, pbLoadFlg)


                    ' ローダ自動時で動作中はボタン非活性化
                    If (giHostMode = cHOSTcMODEcAUTO) And (gbHostConnected = True) Then
                        Call Btn_Enb_OnOff(0)                               ' ボタン非活性化
                        bAutoLoaderAuto = True                              'V2.0.0.0⑮
                    Else
                        If frmAutoObj.gbFgAutoOperation Then
                            'V1.2.0.0④AutoOperationEnd()の中に移動 Call UserSub.SetStartCheckStatus(True)          ' 設定画面の確認有効化
                            Call frmAutoObj.AutoOperationEnd()
                            frmAutoObj.gbFgAutoOperation = False            ' 自動運転終了
                            Me.AutoRunnningDisp.BackColor = System.Drawing.Color.Yellow  ' 背景色 = 黄色
                            Me.AutoRunnningDisp.Text = "自動運転解除中"
                        End If
                        Call Btn_Enb_OnOff(1)                               ' ボタン活性化
                        bAutoLoaderAuto = False                             'V2.0.0.0⑮
                    End If

                    ' ローダから受信したデータをチェックする
                    If gbHostConnected = True And r >= 0 Then
                        If giHostMode = cHOSTcMODEcAUTO Then                ' ｵｰﾄﾛｰﾀﾞ自動 ?
                            Select Case r
                                Case cHSTcTRMCMD                            ' コマンドがトリマ加工指示なら
                                    fStartTrim = True                       ' スタートTRIMフラグをON
                                    Call System1.OperationLogging(gSysPrm, MSG_OPLOG_TRIMST & "[" & DGL.ToString & "]" & Me.AutoRunnningDisp.Text, "HOSTCMD")
                                Case cHSTcLOTCHANGE                         ' コマンドがロット切り替え＆トリマ加工指示なら
                                    fStartTrim = True                       ' スタートTRIMフラグをON
                                    Call System1.OperationLogging(gSysPrm, MSG_OPLOG_LOTCHG & "[" & DGL.ToString & "]" & Me.AutoRunnningDisp.Text, "HOSTCMD")
                            End Select
                        End If
                    End If
                    'V1.2.0.0④↓
                    ObjSys.Z_ATLDGET(LdIDat)                                        ' ローダー入力
                    If (gwPrevHcmd <> LdIDat) Then                                  ' 前回データから変化があった？
                        Call IoMonitor(LdIDat, 0)                                   ' IOﾓﾆﾀ表示
                        gwPrevHcmd = LdIDat
                    End If
                    'V2.0.0.0⑮                If frmAutoObj.gbFgAutoOperation And Not fStartTrim Then                     ' 自動運転モードでアイドル状態の時 
                    If bAutoLoaderAuto And Not fStartTrim Then                                  ' 自動運転モードでアイドル状態の時 'V2.0.0.0⑮自動運転モードでない自動運転もある
                        If gwPrevHcmd And cHSTcCLAMP_ON Then                                    ' クランプ開信号受信
                            If gbClampOpen Then
                                r = Me.System1.ClampCtrl(gSysPrm, 0, giTrimErr)                 ' クランプ/吸着OFF
                                If (r = cFRS_NORMAL) Then
                                    Call Sub_ATLDSET(COM_STS_CLAMP_ON, 0)                       ' ローダー出力(ON=載物台ｸﾗﾝﾌﾟ開,OFF=なし)
                                    gbClampOpen = False
                                Else
                                    Call Z_PRINT("クランプ開エラーが発生しました。手動に切り替えてください。" & vbCrLf)
                                    Call System.Threading.Thread.Sleep(1000)                    ' Wait(ms)
                                End If
                            End If

                        End If
                        If gwPrevHcmd And cHSTcABS_ON Then                                      ' 吸着オフ信号受信
                            If gbVaccumeOff Then
                                Call Me.System1.AbsVaccume(gSysPrm, 0, giAppMode, giTrimErr)    ' バキュームの制御(1=吸着ON, 0=吸着OFF)
                                Call Me.System1.Adsorption(gSysPrm, 0)                          ' 吸着破壊制御(1:吸着, 0:吸着破壊)
                                Call Sub_ATLDSET(COM_STS_ABS_ON, 0)                             ' ローダー出力(ON=吸着:オフ,OFF=なし)
                                gbVaccumeOff = False
                            End If
                        End If
                    End If
                    'V1.2.0.0④↑

                End If

            End If

            'Dim bChangeManual As Boolean = False
            Dim iLotChg As Integer

            'V2.2.1.1⑧↓
            Dim lotcnt As Integer = ObjLoader.GetLotChangeFlg()
            'V2.2.1.1⑧ iLotChg = IsLotChange(giHostMode, r, fStartTrim)
            iLotChg = IsLotChange(giHostMode, r, fStartTrim, lotcnt)
            'V2.2.1.1⑧↑

            If iLotChg > 0 Then                                         ' ロット切り替え条件の場合
                If frmAutoObj.gbFgAutoOperation Then                    ' 自動運転中
LOT_CHG:            'V2.2.1.1⑧ 
                    If frmAutoObj.LotChangeExecuteCheck() Then          ' ロット切り替え可能の場合
                        If frmAutoObj.LotChangeExecute() Then
                            stCounter.LotCounter = stCounter.LotCounter + 1

                            'V2.2.1.1⑧ ↓
                            'フラグでのロット切り替えが複数回の場合ここで行う
                            If lotcnt >= 1 Then
                                lotcnt = lotcnt - 1

                                ObjLoader.SetLotChangeFlg(lotcnt)
                                ' ロット切り替え回数分切り替えを行う
                                If lotcnt > 0 Then
                                    GoTo LOT_CHG
                                End If
                            End If
                            'V2.2.1.1⑧ ↑

                            If Not UserBas.LotStartSetting() Then       ' ロットスタート時処理（印刷データヘッダー情報作成など）
                                bChangeManual = True
                            End If
                            'V2.1.0.0④↓
                        Else
                            bChangeManual = True
                            'V2.1.0.0④↑
                        End If
                    Else
                        Call Z_PRINT("ロット切り替え信号を受けましたが、次のロットはエントリーされていません。" & vbCrLf)
                        bChangeManual = True
                    End If
                End If
            End If

            If bChangeManual Then
                fStartTrim = False                              ' スタートTRIMフラグをOFF
                Call frmAutoObj.AutoOperationEnd()
                'V1.2.0.0④AutoOperationEnd()の中に移動                 Call UserSub.SetStartCheckStatus(True)          ' 設定画面の確認有効化
                Buzzer()                                        'V1.1.0.1② 終了時ブザー
                'V1.1.0.1② System1.SetSignalTower(SIGOUT_RED_ON Or SIGOUT_BZ1_ON, &HFFFFS)
                'V1.2.0.0④ Call Me.System1.TrmMsgBox(gSysPrm, "ロット切り替えエラーが発生しました。", MsgBoxStyle.OkOnly, My.Application.Info.Title)
                'V1.1.0.1② System1.SetSignalTower(0, SIGOUT_RED_ON Or SIGOUT_BZ1_ON)

                ''V2.2.0.0⑤↓
                ' TLF製ローダの場合自動運転切り替えを出力する
                If giLoaderType = 1 Then
                    Call Sub_ATLDSET(0, clsLoaderIf.LINP_AUTO)                    ' ローダー出力(ON=なし,OFF=自動)
                End If

                ' ローダが手動&停止に切り替わるまで待つ
                If giHostMode = cHOSTcMODEcAUTO Then
                    r = Me.System1.Form_Reset(cGMODE_LDR_CHK, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
                    If (r <= cFRS_ERR_EMG) Then
                        GoTo TimerErr                           ' 非常停止等のエラーならアプリ強制終了
                    ElseIf (r = cFRS_ERR_RST) Then              ' キャンセルならコマンド終了
                        GoTo TimerErr
                    End If
                End If


                ''V2.2.0.0⑤↓
                ' TLF製ローダの場合、基板取り除きメッセージを表示する
                If giLoaderType = 1 Then
                    r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE)
                    frmAutoObj.SetAutoOpeCancel(False)               ' V2.2.1.1②
                    Call frmAutoObj.AutoOperationEnd()              ' V2.2.1.1②

                    r = sResetStart()
                    If (r <> cFRS_NORMAL) Then                          ' エラー ?
                        r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE2)
                        Call AppEndDataSave()                           ' ｿﾌﾄ強制終了時のﾃﾞｰﾀ保存確認     
                        Call AplicationForcedEnding()                   ' ｿﾌﾄ強制終了処理
                        End                                             ' アプリ強制終了
                        Return
                    End If
                Else
                    'V2.2.2.0④↓
                    '#0005/#0050の場合ここで戻す
                    ' マーク印字の場合、トリミングに戻す V2.2.1.7⑦↓
                    If UserSub.IsTrimType5() = True Then
                        CbDigSwL.SelectedIndex = 0
                    End If
                    'V2.2.1.7⑦↑
                    'V2.2.2.0④↑

                End If
                ''V2.2.0.0⑤↑

            End If


            '---------------------------------------------------------------------------
            '   ローダマニュアル時(およびローダ無し時)は、以下の処理を行う
            '---------------------------------------------------------------------------
            If (giHostMode = cHOSTcMODEcMANUAL) Then

                r = INTERLOCK_CHECK(interlockStatus, swStatus)
                ' TLF製ローダの場合自動運転切り替えを出力する
                If giLoaderType = 0 Then                ''V2.2.0.0⑤　
                    If (r <> ERR_CLEAR) Then                                    ' ※メッセージ表示を追加する 
                        '筐体カバー開の場合、インターロック解除かカバー閉を監視する
                        If (r = ERR_OPN_CVR) Then
                            iRtn = System1.Form_Reset(cGMODE_CVR_CLOSEWAIT, gSysPrm, giAppMode, False, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
                            If (iRtn <= cFRS_ERR_EMG) Then GoTo TimerErr ' 非常停止等のエラーならアプリ強制終了
                            ' コンソールキーラッチクリア
                            Call ZCONRST()
                        ElseIf (r = ERR_OPN_SCVR Or r = ERR_OPN_CVRLTC) Then
                            ' SL432Rの場合はカバー開ラッチは無視する。
                            If (gSysPrm.stTMN.gsKeimei <> MACHINE_TYPE_SL432) Then
                                'V2.2.0.0⑤↓
                                iRtn = System1.Sub_CoverCheck(gSysPrm, 0, False)
                                If (iRtn <= cFRS_ERR_EMG) Then GoTo TimerErr ' 非常停止等のエラーならアプリ強制終了
                                Call COVERLATCH_CLEAR()                                     ' カバー開ラッチのクリア
                                ' GoTo TimerErr
                                'V2.2.0.0⑤↑
                            End If
                        Else
                            GoTo TimerErr
                        End If
                    End If
                End If

                '---------------------------------------------------------------------------
                '   ＳＴＡＲＴ ＳＷの押下チェック
                '---------------------------------------------------------------------------
                If (giAppMode = APP_MODE_IDLE) Then                     ' アイドルモード時にチェックする
                    r = START_SWCHECK(0, swStatus)                      ' トリマー START SW 押下チェック
                    If (swStatus = cFRS_ERR_START) Then

                        ' 'V2.2.0.0⑤ ↓
                        If giLoaderType <> 0 Then
                            ' 載物台に基板がある事をチェックする(手動モード時(OPTION))
                            r = ObjLoader.Sub_SubstrateExistCheck(System1)
                            If (r <> cFRS_NORMAL) Then                              ' エラー ?
                                If (r = cFRS_ERR_RST) Then                          ' 基板無し(Cancel(RESETｷｰ)　?
                                    Timer1.Enabled = True                           ' タイマー再起動
                                    Exit Sub
                                End If
                                GoTo TimerErr                                       ' その他のエラーならアプリ強制終了
                            End If
                        End If
                        ' 'V2.2.0.0⑤ ↑

                        ' トリマ動作中信号ON送信(オートローダー)
                        Call Sub_ATLDSET(COM_STS_TRM_STATE, COM_STS_TRM_NG Or COM_STS_PTN_NG Or COM_STS_TRM_ERR)          ' ローダー出力(ON=トリマ動作中,OFF=)
                        Call System1.OperationLogging(gSysPrm, MSG_OPLOG_TRIMST, "START SW ON")
                        ' クランプ/吸着ON
                        r = Me.System1.ClampCtrl(gSysPrm, 1, giTrimErr)
                        If (r <> cFRS_NORMAL) Then
                            GoTo TimerErr
                        End If
                        r = Me.System1.AbsVaccume(gSysPrm, 1, APP_MODE_TRIM, giTrimErr) ' APP_MODE_TRIMのモードの時のみバキュームチェックが行われる。
                        If (r <> cFRS_NORMAL) Then
                            Call ZCONRST()                                              ' ラッチ解除
                            Call Me.System1.Adsorption(gSysPrm, 0)
                            Call Me.System1.ClampCtrl(gSysPrm, 0, giTrimErr)
                            GoTo TimerExit
                        End If
                        Call Me.System1.Adsorption(gSysPrm, 1)

                        ' スタートSWを離すまで待つ→INTRIMにて監視の無限ループ
                        Call START_SWCHECK(1, swStatus)

                        ' ｽﾀｰﾄSW押下待ちしない場合はメッセージ表示する
                        If (gSysPrm.stSPF.giWithStartSw = 0) And (interlockStatus = INTERLOCK_STS_DISABLE_NO) Then
                            ' "注意！！！　スライドカバーが自動で閉じます。"(Red,Blue)
                            ' 'V2.2.0.0① r = Me.System1.Form_MsgDispStartReset(MSG_SPRASH31, MSG_SPRASH32, &HFF, &HFF0000)
                            r = Me.System1.Form_MsgDispStartReset(MSG_SPRASH31, MSG_SPRASH32, Color.Blue, Color.Red)           'V2.2.0.0①
                            If (r = cFRS_ERR_RST) Then
                                ' RESET SW押下ならErrorSkipへ
                                GoTo TimerExit
                            End If
                        End If


                        If giLoaderType = 1 And frmAutoObj.gbFgAutoOperation = True Then
                            ObjLoader.DispLoaderInfo()
                        End If

                        ' スライドカバーをクローズする(手動/自動)
                        If (gSysPrm.stSPF.giWithStartSw = 1) And (giHostMode <> cHOSTcMODEcAUTO) Then
                            If (interlockStatus = INTERLOCK_STS_DISABLE_NO) Then
                                ' スライドカバー閉メッセージ表示 (ｽﾀｰﾄSW押下待ち(オプション) でローダ自動運転中でない場合)
                                r = Me.System1.Form_Reset(cGMODE_START_RESET, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, True)
                                If (r = cFRS_ERR_START) Then
                                    ' START SW押下ならスタートTRIMフラグON
                                    fStartTrim = True
                                Else
                                    ' RESET SW押下ならErrorSkipへ
                                    GoTo ErrorSkip
                                End If
                            Else
                                ' START SW押下ならスタートTRIMフラグON
                                fStartTrim = True
                            End If

                        Else
                            If (interlockStatus = INTERLOCK_STS_DISABLE_NO) Then
                                ' スライドカバーを自動クローズする
                                If gSysPrm.stTMN.giOnline = TYPE_MANUAL Then
                                    ' XY_SLIDE同時動作 ?(XY_SLIDE同時動作はローダからのスタート要求時のみのため通常動作とする) 
                                    r = Me.System1.Z_CCLOSE(gSysPrm, giAppMode, 0, False, 0.0, 0.0)
                                End If
                                If gSysPrm.stTMN.giOnline = TYPE_ONLINE Then
                                    ' XY_SLIDE通常動作
                                    r = Me.System1.Z_CCLOSE(gSysPrm, giAppMode, 0, False, 0.0, 0.0)
                                End If
                            Else
                                fStartTrim = True
                            End If
                        End If

                    Else
                        '---------------------------------------------------------------------------
                        '   スライドカバー状態のチェック(SL432R時)
                        '---------------------------------------------------------------------------
                        ' インターロック解除でSL432R系の場合にチェックする 
                        If (interlockStatus = INTERLOCK_STS_DISABLE_NO) And (gSysPrm.stTMN.gsKeimei = MACHINE_TYPE_SL432) Then
                            ' スライドカバーの状態取得（INTRIMではIO取得のみの為、エラーが返る事はない）
                            r = SLIDECOVER_GETSTS(sldCvrSts)

                            ' スライドカバー状態のチェック
                            If (sldCvrSts = SLIDECOVER_MOVING) Then

                                ' スライドカバー中間ならトリマ動作中信号ON送信(オートローダー)
                                Call Sub_ATLDSET(COM_STS_TRM_STATE, 0)  ' ローダー出力(ON=トリマ動作中,OFF=なし)

                                If (gfclamp = False) Then
                                    ' スライドカバー中間、クランプOFFの場合：クランプをONする。
                                    ' クランプ/吸着ON
                                    r = Me.System1.ClampCtrl(gSysPrm, 1, giTrimErr)
                                    If (r <> cFRS_NORMAL) Then
                                        GoTo TimerErr
                                    End If
                                    r = Me.System1.AbsVaccume(gSysPrm, 1, APP_MODE_TRIM, giTrimErr) ' APP_MODE_TRIMのモードの時のみバキュームチェックが行われる。
                                    If (r <> cFRS_NORMAL) Then
                                        Call ZCONRST()                                              ' ラッチ解除
                                        Call Me.System1.Adsorption(gSysPrm, 0)
                                        Call Me.System1.ClampCtrl(gSysPrm, 0, giTrimErr)
                                        GoTo TimerExit
                                    End If
                                    Call Me.System1.Adsorption(gSysPrm, 1)
                                    gfclamp = True
                                End If

                            ElseIf (sldCvrSts = SLIDECOVER_OPEN) Then
                                ' スライドカバーがオープン状態ならトリマ動作中信号OFF送信(オートローダー)
                                Call Sub_ATLDSET(0, COM_STS_TRM_STATE)  ' ローダー出力(ON=なし, OFF=トリマ動作中)

                                ' スライドカバーがオープン状態で、クランプONの場合：クランプをOFFする。
                                gfclamp = False
                                ' クランプ/吸着OFF
                                r = Me.System1.ClampCtrl(gSysPrm, 0, giTrimErr)
                                If (r <> cFRS_NORMAL) Then
                                    GoTo TimerErr
                                End If
                                Call Me.System1.AbsVaccume(gSysPrm, 0, giAppMode, giTrimErr)
                                Call Me.System1.Adsorption(gSysPrm, 0)
                            ElseIf (sldCvrSts = SLIDECOVER_CLOSE) Then
                                ' トリマ動作中信号ON送信(オートローダー)
                                Call Sub_ATLDSET(COM_STS_TRM_STATE, 0)  ' ローダー出力(ON=トリマ動作中,OFF=なし)

                                ' スライドカバー閉
                                gfclamp = False
                                Call COVERLATCH_CLEAR()                 ' ｶﾊﾞｰ開ﾗｯﾁｸﾘｱ
                                Call System1.OperationLogging(gSysPrm, MSG_OPLOG_TRIMST & "[" & DGL.ToString & "]", "SLIDE COVER CLOSED")

                                ' 操作確認画面(START/RESET待ち)
                                If (gSysPrm.stSPF.giWithStartSw = 1) Then
                                    r = System1.Form_Reset(cGMODE_START_RESET, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
                                    If (r = cFRS_ERR_RST) Then
                                        ' RESET SW押下ならErrorSkipへ
                                        Call UserSub.SetStartCheckStatus(True)          ' 設定画面の確認有効化
                                        GoTo ErrorSkip
                                    ElseIf (r <> cFRS_ERR_START) Then
                                        GoTo TimerErr
                                    End If
                                End If
                                fStartTrim = True                       ' スタートTRIMフラグON
                            End If
                        End If
                    End If
                End If
            End If

            '---------------------------------------------------------------------------
            '   スタートTRIMフラグがONなら、以下の処理を行う
            '---------------------------------------------------------------------------
            If fStartTrim = True Then
                ' データはロード済みか
                If pbLoadFlg = False Then                               ' データ未ロード ?
                    strMSG = MSG_DataNotLoad                            ' "データ未ロード"
                    Call Z_PRINT(strMSG)                                ' メッセージ表示
                    GoTo ErrorSkip
                End If

                ' トリマ装置アイドルでなければErrorSkipへ
                gfclamp = False                                         ' FLG = クランプOFF
                If giAppMode Then GoTo ErrorSkip '                      ' トリマ装置アイドルでない ?
                gbInitialized = False                                   ' flg = 原点復帰未

                If Not frmAutoObj.gbFgAutoOperation Then            ' 自動運転中でない時
                    If UserSub.IsSpecialTrimType() Then             ' ユーザプログラム特殊処理
                        If Not UserBas.LotStartSetting() Then       ' ロットスタート時処理（印刷データヘッダー情報作成など）
                            GoTo ErrorSkip
                        End If
                    End If
                End If

                'V2.1.0.0②↓ レーザーパワーキャリブレーション機能後へ移動                Call SetATTRateToScreen(True)           '###1040⑥ トリミングデータでのＡＴＴ減衰率の設定

                ' ローダへデータ出力
                ' 'V2.2.0.0⑤ ↓ TLF製ローダ時は前回基板の結果を保存
                If giLoaderType = 1 Then
                    Call Sub_ATLDSET(0, clsLoaderIf.LOUT_REDY Or clsLoaderIf.LOUT_STOP Or clsLoaderIf.LOUT_TRM_NG)
                    ObjLoader.SetLotAbort(0)

                Else
                    'V1.2.0.0④#If cATLcDEN = 0 Then
                    ' ON=トリマ動作中, OFF=トリミングNG,パターン認識NG
                    'V1.2.0.0④                Call Sub_ATLDSET(COM_STS_TRM_STATE, COM_STS_TRM_NG Or COM_STS_PTN_NG)
                    Call Sub_ATLDSET(COM_STS_TRM_STATE, COM_STS_TRM_NG Or COM_STS_PTN_NG Or COM_STS_CLAMP_ON Or COM_STS_ABS_ON)
                    'V1.2.0.0④#Else
                    'V1.2.0.0④                			Call Sub_ATLDSET(COM_STS_TRM_STATE Or COM_STS_1ST_CMD, COM_STS_TRM_OK Or COM_STS_TRM_PRB Or COM_STS_TRM_NG)
                    'V1.2.0.0④#End If
                End If
                ' 'V2.2.0.0⑤ ↑
                '-----------------------------------------------------------------------
                '   ｼｸﾞﾅﾙﾀﾜｰ制御(ｵﾌﾟｼｮﾝ)
                '-----------------------------------------------------------------------
                giTrimErr = 0                                           ' ﾄﾘﾏｰ ｴﾗｰ ﾌﾗｸﾞ初期化
                If (giHostMode = cHOSTcMODEcAUTO) Then                  ' ﾛｰﾀﾞ自動ﾓｰﾄﾞ ?
                    ' ｼｸﾞﾅﾙﾀﾜｰ制御(On=自動運転中 , Off=全ビット)
                    r = System1.SetSignalTower(SIGOUT_GRN_ON, &HFFFFS)
                Else
                    ' ｼｸﾞﾅﾙﾀﾜｰ制御(On=なし,Off=全ビット)
                    r = System1.SetSignalTower(0, &HFFFFS)
                End If

                Call System1.sLampOnOff(LAMP_START, True)               ' STARTランプ点灯
                giAppMode = APP_MODE_TRIM                               ' ｱﾌﾟﾘﾓｰﾄﾞ = トリミング中

                '-----------------------------------------------------------------------
                '   スライドカバー自動クローズ
                '-----------------------------------------------------------------------
                ' オートローダ自動 ?
                If giHostMode = cHOSTcMODEcAUTO Then                    ' ローダ自動モード ?
                    ' スライドカバーを自動クローズする
                    If gSysPrm.stTMN.giOnline = TYPE_MANUAL Then
                        ' XY_SLIDE同時動作 ?
                        r = Me.System1.Z_CCLOSE(gSysPrm, giAppMode, 0, True, gSysPrm.stDEV.gfTrimX, gSysPrm.stDEV.gfTrimY)
                    End If
                    If gSysPrm.stTMN.giOnline = TYPE_ONLINE Then
                        ' XY_SLIDE通常動作
                        r = Me.System1.Z_CCLOSE(gSysPrm, giAppMode, 0, False, 0.0, 0.0)
                    End If
                    If (r <= cFRS_ERR_EMG) Then GoTo TimerErr '         ' 非常停止等のエラーならアプリ強制終了

                Else                                                    ' ｽﾗｲﾄﾞｶﾊﾞｰ自動ｸﾛｰｽﾞ
                    If (gSysPrm.stSPF.giWithStartSw = 0) Then           ' ｽﾀｰﾄSW押下でﾄﾘﾐﾝｸﾞ開始(ｵﾌﾟｼｮﾝ)時は自動ｸﾛｰｽﾞしない
                        r = System1.Z_CCLOSE(gSysPrm, giAppMode, giTrimErr, False, 0, 0)
                        If (r <= cFRS_ERR_EMG) Then GoTo TimerErr '     ' 非常停止等のエラーならアプリ強制終了
                    End If
                End If

                iRtn = SetZOff_Prob_Off()                        ' INTIME内部の待機位置を変更して待機位置移動する。###1041①
                If iRtn <> cFRS_NORMAL Then
                    Call Me.System1.TrmMsgBox(gSysPrm, "SETZOFFPOS Ｚ軸待機位置変更が異常終了しました。", MsgBoxStyle.OkOnly, My.Application.Info.Title)
                    GoTo ErrorSkip
                End If

                ' 'V2.2.0.0⑤ ↓ TLF製ローダ時は前回基板の結果を保存
                If giLoaderType = 1 Then
                    swMesureTrimtime.Stop()
                    swMesureTrimtime.Reset()
                    swMesureTrimtime.Start()
                End If

                Call UserSub.ClampVacumeChange()         'V2.0.0.0⑭
                'V2.1.0.0②↓ レーザーパワーキャリブレーション機能
                '-----------------------------------------------------------------------
                '   レーザーパワーのモニタリング実行
                '-----------------------------------------------------------------------
                If UserSub.LaserCalibrationExecute() Then
                    Dim tmpiAttFix As Short = gSysPrm.stRAT.giAttFix
                    Dim tmpiAttRot As Short = gSysPrm.stRAT.giAttRot
                    Dim tmpfAttRate As Double = gSysPrm.stRAT.gfAttRate
                    If UserSub.LaserCalibrationFullPowerGet(stLASER.dblPowerAdjustTarget, stLASER.dblPowerAdjustToleLevel) Then
                        stLASER.dblPowerAdjustQRate = stLASER.intQR / 10.0#
                        r = AutoLaserPowerADJ(True)                         ' レーザパワー調整実行
                        If (r = cFRS_ERR_RST) Then                          ' Cancel(RESETｷｰ) ?

                            'V2.2.1.1④↓
                            If giLoaderType = 1 Then
                                ' フルパワーチェックがエラーの場合にはロットを終了してメイン画面に戻る 

                                ' ロット処理中断 
                                ' 中断等で次の基板は処理しない。 
                                fStartTrim = False                       ' スタートTRIMフラグをOFF
                                ' 　基板取り除きメッセージを表示する
                                r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE)

                                frmAutoObj.SetAutoOpeCancel(False)               ' V2.2.1.1②
                                Call frmAutoObj.AutoOperationEnd()              ' V2.2.1.1②
                                frmAutoObj.gbFgAutoOperation = False

                                ' 原点復帰確認 
                                r = sResetStart()
                                If (r <> cFRS_NORMAL) Then                          ' エラー ?
                                    '原点復帰エラーの場合はプログラム終了 
                                    r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE2)
                                    Call AppEndDataSave()                           ' ｿﾌﾄ強制終了時のﾃﾞｰﾀ保存確認
                                    Call AplicationForcedEnding()                   ' ｿﾌﾄ強制終了処理
                                    End                                             ' アプリ強制終了
                                    Return
                                End If
                                ' 電磁ロック(観音扉右側ロック)を解除する
                                r = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_OFF)
                                GoTo ErrorSkip                              ' 自動から手動に切り替わった時

                            Else
                                ' "ローダ信号が自動です", "ローダを手動に切り替えてください"
                                r = Me.System1.Form_Reset(cGMODE_LDR_CHK, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
                                If (r = cFRS_NORMAL) Then
                                    UserSub.LaserCalibrationSet(POWER_CHECK_LOT) 'レーザパワーモニタリング実行有無設定
                                    GoTo ErrorSkip                              ' 自動から手動に切り替わった時
                                Else
                                    GoTo TimerErr                               ' （"Cancelボタン押下でプログラムを終了します"は、cFRS_ERR_RST）アプリ強制終了
                                End If
                            End If
                            'V2.2.1.1④↑

                        ElseIf (r <> cFRS_NORMAL) Then                      ' エラー ?(※エラーメッセージは表示済み) 
                            GoTo TimerErr                                   ' アプリ強制終了
                        End If
                    Else
                        GoTo ErrorSkip                              ' 自動から手動に切り替わった時
                    End If
                    Call ATTRESET()
                End If
                If SetATTRateToScreen(True) = False Then
                    GoTo TimerErr                                       ' アプリ強制終了
                End If
                'V2.1.0.0②↑

                'V2.2.1.7③↓
                If frmAutoObj.gbFgAutoOperation = True Then
                    MarkingCount = MarkingCount + 1             ' マーキング用カウンタクリア	
                Else ' frmAutoObj.gbFgAutoOperation = False Then
                    ' 手動の場合カウンタは1で、マーキング時-1して開始番号をそのままマーク印字 
                    MarkingCount = 1             ' マーキング用カウンタクリア	
                End If
                'V2.2.1.7③↑
                'V2.2.1.7③↑

                '-----------------------------------------------------------------------
                '   トリミング実行
                '-----------------------------------------------------------------------
                iRtn = User()                                           ' トリミング実行


                ' 'V2.2.0.0⑤ ↓ TLF製ローダ時は前回基板の結果を保存
                If giLoaderType = 1 And frmAutoObj.gbFgAutoOperation = True Then
                        ObjLoader.m_lTrimResult = iRtn
                        ' 'V2.2.0.0⑤ ↓ TLF製ローダ時は前回基板の結果を保存
                        swMesureTrimtime.Stop()
                        '' トリミング時間書込み 
                        gdTrimtime = swMesureTrimtime.Elapsed

                        '' 基板交換時間書込み
                        'Dim dummy As Integer
                        Dim SupplyMag As Integer = 0
                        Dim SupplySlot As Integer = 0
                        Dim StoreMag As Integer = 0
                        Dim StoreSlot As Integer = 0

                        ObjSys.Sub_GetNowProcessMgInfo(SupplyMag, SupplySlot, StoreMag, StoreSlot)
                        objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_SUPPLY_MAGAGINE, SupplyMag)
                        objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_SUPPLY_SLOT, SupplySlot)
                        'V2.2.0.037　objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_STORE_MAGAGINE, StoreMag)
                        'V2.2.0.037　objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_STORE_SLOT, StoreSlot)

                        If iRtn = cFRS_ERR_RST Then
                            ObjLoader.SetLotAbort(1)
                        End If

                    End If
                    ' 'V2.2.0.0⑤ ↑

                    If (iRtn >= cFRS_NORMAL) Then                           ' 正常/RESET SW押下 ?


                        ' ﾄﾘﾐﾝｸﾞNG/ﾄﾘﾏｰｴﾗｰ/ﾊﾟﾀｰﾝ認識ｴﾗｰ時
                    ElseIf (iRtn = cFRS_TRIM_NG) Or (iRtn = cFRS_ERR_TRIM) Or (iRtn = cFRS_ERR_PTN) Then

                        '' 筐体カバー開/スライドカバー開/カバー開ラッチ検出時は強制終了しない
                        'ElseIf (iRtn = cFRS_ERR_CVR) Or (iRtn = cFRS_ERR_SCVR) Or (iRtn = cFRS_ERR_LATCH) Then

                        ' 非常停止等のアプリ強制終了エラー時
                    Else ' クランプ/吸着OFF(ﾛｰﾀﾞｰへはﾄﾘﾏｰ動作中OFFしない)
                        r = System1.ClampVacume_Ctrl(gSysPrm, 0, giAppMode, 1, False)
                        Call ZSLCOVEROPEN(0)                                ' ｽﾗｲﾄﾞｶﾊﾞｰｵｰﾌﾟﾝﾊﾞﾙﾌﾞOFF
                        Call ZSLCOVERCLOSE(0)                               ' ｽﾗｲﾄﾞｶﾊﾞｰｸﾛｰｽﾞﾊﾞﾙﾌﾞOFF
                        GoTo TimerErr                                       ' アプリ強制終了
                    End If

ErrorSkip:
                    If giHostMode = cHOSTcMODEcMANUAL Then                  ' ローダマニュアル？
                        Call ZCONRST()                                      ' ｺﾝｿｰﾙSWﾗｯﾁ解除
                    End If

                    '-----------------------------------------------------------------------
                    '   スライドカバー自動オープン
                    '-----------------------------------------------------------------------

                    If (giLoaderType = 1) AndAlso (giHostMode = cHOSTcMODEcAUTO) Then
                        ' 基板排出位置への移動 
                        r = ObjLoader.MoveGlassOutPos()
                        If (r <> cFRS_NORMAL) Then                                  ' エラー ?
                            frmAutoObj.gbFgAutoOperation = False
                            Call Sub_ATLDSET(0, clsLoaderIf.LINP_AUTO)                    ' ローダー出力(ON=なし,OFF=自動)
                            GoTo TimerErr                                       ' アプリ強制終了
                        End If

                    Else
                        'If (giHostMode <> cHOSTcMODEcAUTO) Then                 ' ローダ自動モードでない ? 
                        If (gSysPrm.stTMN.giOnline = 1) Or ((giHostMode <> cHOSTcMODEcAUTO) And (gSysPrm.stTMN.giOnline = 2)) Then
                            r = System1.EX_SBACK(gSysPrm)                       ' ﾊﾟｰﾂﾊﾝﾄﾞﾗをﾛｰﾄﾞ位置に戻す
                            If (r <= cFRS_ERR_EMG) Then GoTo TimerErr '         ' 非常停止等のエラーならアプリ強制終了
                            'V2.1.0.1①                    ' θ軸原点復帰(｢自動｣又は｢自動+微調整｣又は｢手動で補正なしでθﾊﾟﾗﾒｰﾀ=原点復帰指定｣時) ※θありの場合
                            'V2.1.0.1①                    If ((stThta.iPP30 = 0) And (gSysPrm.stDEV.giTheta <> 0)) Or _
                            'V2.1.0.1①                       ((stThta.iPP30 = 2) And (gSysPrm.stDEV.giTheta <> 0)) Or _
                            'V2.1.0.1①                       ((stThta.iPP30 = 1) And (stThta.iPP31 = 0) And (gSysPrm.stSPF.giThetaParam = 1) And (gSysPrm.stDEV.giTheta <> 0)) Then
                            'V2.1.0.1①                        Call ROUND4(0.0#)                               ' θを原点に戻す
                            'V2.1.0.1①                    End If
                        End If


                    End If

                    Call ROUND4(0.0#)                               'V2.1.0.1①必ずθを原点に戻す

                    If UserSub.IsTRIM_MODE_ITTRFT() And Not UserSub.GetStartCheckStatus() And iRtn <> cFRS_ERR_RST And r <> cFRS_ERR_RST Then
                        UserBas.stCounter.EndTime = DateTime.Now()          ' 基板処理終了時間保存 '###1030③
                        Buzzer()                                            ' 終了時ブザー
                    End If

                    ' スライドカバー自動オープン
                    ' ｽﾀｰﾄSW押下でﾄﾘﾐﾝｸﾞ開始(ｵﾌﾟｼｮﾝ)時は自動ｵｰﾌﾟﾝしない
                    If (gSysPrm.stSPF.giWithStartSw = 0) Or (giHostMode = cHOSTcMODEcAUTO) Then
                        'V1.2.0.0④↓ ｱﾌﾟﾘﾓｰﾄﾞ　Z_COPENのgiAppModeをiAppModeに変更
                        iAppMode = giAppMode
                        'V2.0.0.0⑮                    If frmAutoObj.gbFgAutoOperation And giAppMode = APP_MODE_TRIM Then  '自動運転時は、Z_COPEN内でクランプ開を行わない。
                        If bAutoLoaderAuto And giAppMode = APP_MODE_TRIM Then  '自動運転時は、Z_COPEN内でクランプ開を行わない。
                            iAppMode = APP_MODE_TRIM_AUTO            'V2.0.0.0⑮　APP_MODE_AUTOからAPP_MODE_TRIM_AUTOへ変更
                        End If
                        'V1.2.0.0④↑
                        If (gSysPrm.stTMN.giOnline = 1) Then                ' ｽﾗｲﾄﾞｶﾊﾞｰ自動ｵｰﾌﾟﾝ(XY_SLIDE通常動作)
                            r = System1.Z_COPEN(gSysPrm, iAppMode, giTrimErr, False)
                            If (r <= cFRS_ERR_EMG) Then GoTo TimerErr '     ' 非常停止等のエラーならアプリ強制終了
                        End If
                        If (gSysPrm.stTMN.giOnline = 2) Then                ' ｽﾗｲﾄﾞｶﾊﾞｰ自動ｵｰﾌﾟﾝ(XY_SLIDE同時動作)
                            r = System1.Z_COPEN(gSysPrm, iAppMode, giTrimErr, True)
                            If (r <= cFRS_ERR_EMG) Then GoTo TimerErr '     ' 非常停止等のエラーならアプリ強制終了
                        End If

                        ' トリミング終了時のｶﾊﾞｰ開待ち(ｵﾌﾟｼｮﾝ) (インターロック中の場合)
                    Else
                        If (System1.InterLockSwRead() And BIT_INTERLOCK_DISABLE) = 0 Then
                            r = System1.Form_Reset(cGMODE_OPT_END, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
                            If (r <= cFRS_ERR_EMG) Then GoTo TimerErr '     ' 非常停止等のエラーならアプリ強制終了
                            r = System1.ClampVacume_Ctrl(gSysPrm, 0, giAppMode, 1, False)
                            If (r <> cFRS_NORMAL) Then GoTo TimerErr '      ' アプリ強制終了
                        Else
                            r = System1.ClampVacume_Ctrl(gSysPrm, 0, giAppMode, 1, False)
                            If (r <= cFRS_ERR_EMG) Then GoTo TimerErr '     ' 非常停止等のエラーならアプリ強制終了
                        End If
                    End If

                    'V2.1.0.1①                ' θ軸原点復帰(｢自動｣又は｢自動+微調整｣又は｢手動で補正なしでθﾊﾟﾗﾒｰﾀ=原点復帰指定｣時) ※θありの場合
                    'V2.1.0.1①                If ((stThta.iPP30 = 0) And (gSysPrm.stDEV.giTheta <> 0)) Or _
                    'V2.1.0.1①                   ((stThta.iPP30 = 2) And (gSysPrm.stDEV.giTheta <> 0)) Or _
                    'V2.1.0.1①                   ((stThta.iPP30 = 1) And (stThta.iPP31 = 0) And (gSysPrm.stSPF.giThetaParam = 1) And (gSysPrm.stDEV.giTheta <> 0)) Then
                    'V2.1.0.1①                    Call ROUND4(0.0#)                                   ' θを原点に戻す
                    'V2.1.0.1①                End If

                    ' ﾗﾝﾌﾟ設定
                    Call System1.sLampOnOff(LAMP_RESET, True)               ' RESETランプ点灯
                    Call System1.sLampOnOff(LAMP_START, True)               ' STARTランプ点灯

                    'V2.0.0.2①                If (iRtn <> cFRS_ERR_PTN And DGL <> TRIM_MODE_POWER) Then                          ' ###1040⑤ θのパターン認識ＮＧ時に処理基板が印刷されてしまう修正。'V2.0.0.0②TRIM_MODE_POWER追加
                    If UserSub.IsTRIM_MODE_ITTRFT() And (iRtn <> cFRS_ERR_PTN And DGL <> TRIM_MODE_POWER) Then                   ' ###1040⑤ θのパターン認識ＮＧ時に処理基板が印刷されてしまう修正。'V2.0.0.0②TRIM_MODE_POWER追加'V2.0.0.2① カット実行を除外する為にIsTRIM_MODE_ITTRFT()追加
                        Call UserSub.SubstrateEnd()                         ' 基板単位の結果出力
                    End If

                    gbClampOpen = True        'V1.2.0.0④ クランプ開状態解除
                    gbVaccumeOff = True       'V1.2.0.0④ 吸着オフ状態解除

                    '-----------------------------------------------------------------------
                    '   トリミング結果をローダへ出力する
                    '-----------------------------------------------------------------------
                    'V2.0.0.1③↓
                    If iRtn = cFRS_TRIM_NG Then
                        If UserSub.IsTRIM_MODE_ITTRFT() And PlateNGJudgeByCounter() Then

                            'V2.2.0.036↓
                            If (giLoaderType = 1) AndAlso (giHostMode = cHOSTcMODEcAUTO) Then

                                ' シグナルタワー制御(On=自動運転中(緑点灯),Off=全ﾋﾞｯﾄ)
                                Call Me.System1.SetSignalTowerCtrl(Me.System1.SIGNAL_ALARM)

                                '  "NG率が設定値を超えました。" "STARTキー：処理続行，RESETキー：処理終了"
                                Dim ret As Integer = ObjLoader.Sub_CallFrmMsgDisp(Me.System1, cGMODE_MSG_DSP, cFRS_ERR_START + cFRS_ERR_RST, True,
                                    My.Resources.MSG_SPRASH56, My.Resources.MSG_SPRASH35, "", System.Drawing.Color.Red, System.Drawing.Color.Black, System.Drawing.Color.Black)

                                If (ret = cFRS_ERR_START) Then
                                    '続行
                                    Call Me.System1.SetSignalTowerCtrl(Me.System1.SIGNAL_OPERATION)

                                Else
                                    Call Me.System1.SetSignalTowerCtrl(Me.System1.SIGNAL_IDLE)

                                    ' ロット処理中断 
                                    ' 中断等で次の基板は処理しない。 
                                    fStartTrim = False                       ' スタートTRIMフラグをOFF
                                    ' 　基板取り除きメッセージを表示する
                                    r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE)

                                    frmAutoObj.SetAutoOpeCancel(False)               ' V2.2.1.1②
                                    Call frmAutoObj.AutoOperationEnd()              ' V2.2.1.1②
                                    frmAutoObj.gbFgAutoOperation = False

                                    ' 原点復帰確認 
                                    r = sResetStart()
                                    If (r <> cFRS_NORMAL) Then                          ' エラー ?
                                        '原点復帰エラーの場合はプログラム終了 
                                        r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE2)
                                        Call AppEndDataSave()                           ' ｿﾌﾄ強制終了時のﾃﾞｰﾀ保存確認
                                        Call AplicationForcedEnding()                   ' ｿﾌﾄ強制終了処理
                                        End                                             ' アプリ強制終了
                                        Return
                                    End If
                                    ' 電磁ロック(観音扉右側ロック)を解除する
                                    r = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_OFF)

                                End If
                            End If
                            'V2.2.0.036↑

                            iRtn = cFRS_TRIM_NG
                        Else
                            iRtn = cFRS_NORMAL
                        End If
                    End If
                    'V2.0.0.1③↑

                    If (iRtn <> cFRS_NORMAL And iRtn <> cFRS_ERR_RST) Then      'V2.0.0.1③ cFRS_ERR_RST追加
                        ' エラー時
                        If giLoaderType = 1 Then
                            Call Sub_ATLDSET(0, COM_STS_TRM_STATE)
                        Else
                            Call Sub_ATLDSET(COM_STS_TRM_NG, COM_STS_TRM_STATE)
                        End If

                        DebugLogOut("トリミング不良信号(BIT1)出力 Result=[" & iRtn.ToString & "]")
                    Else
                        ' 正常時
                        If giLoaderType = 1 Then
                            Call Sub_ATLDSET(0, COM_STS_TRM_STATE)
                        Else
                            Call Sub_ATLDSET(0, COM_STS_TRM_STATE Or COM_STS_TRM_NG)
                        End If

                    End If
                    'If (r = cFRS_ERR_PTN) Then                             ' パターン認識エラー ?
                    '    Call Sub_ATLDSET(COM_STS_PTN_NG Or COM_STS_TRM_NG, COM_STS_TRM_STATE)
                    'ElseIf (r <> cFRS_NORMAL) Then                         ' エラー ?
                    '    Call Sub_ATLDSET(COM_STS_TRM_NG, COM_STS_TRM_STATE Or COM_STS_PTN_NG)
                    'Else                                            ' 正常
                    '    Call Sub_ATLDSET(0, COM_STS_TRM_STATE Or COM_STS_TRM_NG Or COM_STS_PTN_NG)
                    'End If

                    'V2.2.0.0⑤↓
                    ' TLF製ローダの場合、手動ならロック解除
                    If (giLoaderType <> 0) AndAlso (frmAutoObj.gbFgAutoOperation = False) Then
                        r = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_OFF)
                        If (r <> cFRS_NORMAL) Then                                  ' エラー ?(メッセージは表示済)

                        End If
                    End If

                    ' トリミング実行後処理
                    Call System1.AutoLoaderFlgReset()                       ' オートローダーフラグリセット
                    giAppMode = APP_MODE_IDLE                               ' ｱﾌﾟﾘﾓｰﾄﾞ =トリマ装置アイドル中
                    fStartTrim = False                                      ' スタートTRIMフラグ OFF

                    '---------------------------------------------------------------------------
                    '   スタートTRIMフラグがOFFなら、以下の処理を行う
                    '---------------------------------------------------------------------------
                Else
                    ' マニュアルローダでRESET SW押下時は「原点復帰処理」を行う(アイドルモード時にチェックする)
                    If (giAppMode = APP_MODE_IDLE) Then
                    STARTRESET_SWCHECK(1, swStatus)
                    ' マニュアルローダでRESET ＳＷ押下？
                    If giHostMode = cHOSTcMODEcMANUAL And swStatus = cFRS_ERR_RST Then
                        Call System1.sLampOnOff(LAMP_RESET, True)       ' RESETランプON
                        ' 原点復帰
                        r = sResetStart()                               ' RESET/STARTキー待ち
                        ' 非常停止等ならアプリ強制終了
                        If (r < cFRS_NORMAL) Then
                            r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE2)

                            GoTo TimerErr
                        End If
                        If (r = cFRS_ERR_RST) Then                      ' RESETキー押下 ? 
                            ' ｲﾝﾀｰﾛｯｸ時ならｽﾗｲﾄﾞｶﾊﾞｰ開待ち
                            ' V2.2.0.0⑤ If (Me.System1.InterLockSwRead() And BIT_INTERLOCK_DISABLE) = 0 Then
                            If (Me.System1.InterLockSwRead() And BIT_INTERLOCK_DISABLE) = 0 AndAlso giLoaderType = 0 Then   'V2.2.0.0⑤
                                r = Me.System1.Form_Reset(cGMODE_OPT_END, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
                                If (r < cFRS_NORMAL) Then GoTo TimerErr ' エラーならアプリ強制終了
                            End If
                        End If

                        ' ランプ制御
                        Call Me.System1.sLampOnOff(LAMP_START, True)    ' STARTﾗﾝﾌﾟON
                        Call Me.System1.sLampOnOff(LAMP_RESET, True)    ' RESETﾗﾝﾌﾟON
                        GoTo TimerExit
                    End If
                End If
            End If

            '---------------------------------------------------------------------------
            '   終了処理
            '---------------------------------------------------------------------------
TimerExit:
            Timer1.Enabled = True                                       ' 監視タイマー開始
            Exit Sub

            '---------------------------------------------------------------------------
            '   アプリ強制終了
            '---------------------------------------------------------------------------
TimerErr:
            Call AppEndDataSave()                                       ' ｿﾌﾄ強制終了時のﾃﾞｰﾀ保存確認
            Call AplicationForcedEnding()                               ' ｿﾌﾄ強制終了処理
            End                                                         ' アプリ強制終了

        Catch ex As Exception
            Call Z_PRINT("Timer1.Tick() TRAP ERROR = " + ex.Message & vbCrLf)
        End Try
    End Sub
#End Region

    '========================================================================================
    '   その他のイベント処理
    '========================================================================================
#Region "デジタルSWの選択項目が変わった場合の処理"
    '''=========================================================================
    ''' <summary>デジタルSWの選択項目が変わった場合の処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub CbDigSwL_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CbDigSwL.SelectedIndexChanged
        DGL = CbDigSwL.SelectedIndex                            ' デジタルＳＷ(Low)
        DGSW = (DGH * 10) + DGL                                 ' デジタルＳＷ
    End Sub

    Private Sub CbDigSwH_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CbDigSwH.SelectedIndexChanged
        DGH = CbDigSwH.SelectedIndex                            ' デジタルＳＷ(Hight) 
        DGSW = (DGH * 10) + DGL                                ' デジタルＳＷ
    End Sub
#End Region

    '========================================================================================
    '   共通関数
    '========================================================================================
#Region "原点復帰サブ"
    '''=========================================================================
    '''<summary>原点復帰サブ</summary>
    ''' <returns>0:正常, 0以外:エラー</returns>
    '''<remarks></remarks>
    '''=========================================================================
    Public Function sResetStart() As Short

        Dim r As Short = cFRS_NORMAL
        Dim rtn As Short = cFRS_NORMAL
        Dim strMSG As String

        Try
            ' 原点復帰
#If cOFFLINEcDEBUG = 1 Then
            Return (cFRS_NORMAL)                                                    ' Return値 = 正常
#End If
            ' 原点復帰

            Call SETAXISSPDY(SETAXISSPDY_DEFALT)                                    ' 'V2.0.0.0⑮Ｙ軸ステージ速度を元に戻す。25000から15000へ変更

            Call Sub_ATLDSET(COM_STS_TRM_STATE Or COM_STS_LOT_END, 0)               ' ローダー出力(ON=トリマ動作中 　'V1.2.0.0④ロット終了(0:処理中, 1:終了状態),OFF=なし)


            'V2.2.0.0⑤↓
            ' TLF製ローダの場合、ローダ原点復帰を行う
            If (giLoaderType <> 0) Then
                Call ATTRESET()                                             'V2.1.0.0⑥
                ' アラームリセットする
                ObjSys.W_RESET()
                Call Sub_ATLDSET(0, clsLoaderIf.LOUT_AUTO Or clsLoaderIf.LOUT_STOP Or clsLoaderIf.LOUT_SUPLY Or clsLoaderIf.LOUT_STS_RUN Or clsLoaderIf.LOUT_REQ_COLECT Or clsLoaderIf.LOUT_DISCHRAGE)                             ' ローダ出力(ON=基板要求または供給位置決完了+ﾄﾘﾏ停止中+他, OFF=供給位置決完了または基板要求)

                Call Sub_ATLDSET(clsLoaderIf.LOUT_STS_RUN, clsLoaderIf.LOUT_REDY)               ' ローダー出力(ON=トリマ動作中 　'V1.2.0.0④ロット終了(0:処理中, 1:終了状態),OFF=なし)

            End If
            'V2.2.0.0⑤↑


            r = System1.Form_Reset(cGMODE_ORG, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
            If (r = cFRS_ERR_RST) Then                                              ' RESET ?
                rtn = r                                                             ' Return値設定
                GoTo STP_END
            End If
            If (r <> cFRS_NORMAL) Then                                              ' エラー ?

                'ローダ原点復帰タイムアウトの場合には、
                If (r = cFRS_ERR_LDRTO) Then                          ' ローダ通信タイムアウト ?
                    ' rtnCode = Sub_CallFrmRset(ObjSys, cGMODE_LDR_TMOUT)     ' エラーメッセージ表示
                    AutoOperationDebugLogOut("sResetStart() r = cFRS_ERR_LDRTO")       ''V2.2.1.3②

                    r = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_TMOUT)
                Else
                    ' ローダアラームメッセージ作成 & ローダアラーム画面表示
                    r = ObjSys.Sub_CallFormLoaderAlarm(cGMODE_LDR_ALARM, ObjPlcIf)
                End If
                Call Sub_ATLDSET(&H0, clsLoaderIf.LOUT_AUTO)        ' ローダ手動モード切替え(ローダ出力(ON=なし, OFF=自動))

                Return (r)
            End If

            ' ｽﾀｰﾄSW押下待ち(ｵﾌﾟｼｮﾝ)時はｽﾗｲﾄﾞｶﾊﾞｰ自動ｵｰﾌﾟﾝしないので
            ' ﾒｯｾｰｼﾞ表示後ｽﾗｲﾄﾞｶﾊﾞｰ開待ち
            If ((System1.InterLockSwRead() And BIT_INTERLOCK_DISABLE) = 0) And
                (gSysPrm.stSPF.giWithStartSw = 1) Then                               ' ｲﾝﾀｰﾛｯｸ時でｽﾀｰﾄSW押下待ち(ｵﾌﾟｼｮﾝ) ?
                r = System1.Form_Reset(cGMODE_OPT_END, gSysPrm, giAppMode, gbInitialized, stPLT.Z_ZOFF, stPLT.Z_ZON, pbLoadFlg)
                If (r <> cFRS_NORMAL And r <> cFRS_ERR_RST) Then                    ' エラー ?
                    Return (r)                                                      ' Return値設定
                End If
                rtn = r
            End If

            'V2.2.0.0⑤↓
            ' TLF製ローダの場合、ローダ原点復帰を行う
            If (giLoaderType <> 0) Then
                r = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_OFF)
                If (r <> cFRS_NORMAL) Then                                  ' エラー ?(メッセージは表示済)
                    Return (r)
                End If
            End If

            ' クランプ/吸着OFF
            r = System1.ClampVacume_Ctrl(gSysPrm, 0, giAppMode, 1, False)
            If (r <> cFRS_NORMAL) Then                                              ' エラー ?
                Return (r)                                                          ' Return値設定
            End If



STP_END:
            Call Sub_ATLDSET(0, COM_STS_TRM_STATE)                                  ' ローダー出力(ON=なし,OFF=トリマ動作中)
            Return (rtn)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "sResetStart() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            sResetStart = cERR_TRAP                                                 ' Return値 = 例外エラー
        End Try
    End Function
#End Region

#Region "インターロック状態の表示/非表示"
    '''=========================================================================
    '''<summary>インターロック状態の表示/非表示</summary>
    ''' <returns>インターロック状態
    '''          INTERLOCK_STS_DISABLE_FULL = インターロック全解除
    '''          INTERLOCK_STS_DISABLE_PART = インターロック一部解除（ステージ動作可能）
    '''          INTERLOCK_STS_DISABLE_NO   = インターロック状態（解除なし）
    ''' </returns>
    '''<remarks></remarks>
    '''=========================================================================
    Public Function DispInterLockSts() As Integer

        Dim r As Integer
        Dim InterlockSts As Integer
        Dim SwitchSts As Long
        Dim strMSG As String

        Try
            ' インターロック状態によりステータス表示を変更
            r = INTERLOCK_CHECK(InterlockSts, SwitchSts)
#If cOFFLINEcDEBUG Then
            InterlockSts = INTERLOCK_STS_DISABLE_FULL
#End If
            If (InterlockSts = INTERLOCK_STS_DISABLE_FULL) Then         ' インターロック全解除 ?
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    strMSG = "インターロック全解除中"
                Else
                    strMSG = "Under Interlock Release"
                End If
                Me.lblInterLockMSG.Text = strMSG
                Me.lblInterLockMSG.Visible = True
                'V2.2.0.0⑤ ↓
                If giLoaderType = 1 Then
                    Call Sub_ATLDSET(clsLoaderIf.LOUT_INTLOK_DISABLE, 0)
                End If
                'V2.2.0.0⑤ ↑

            ElseIf (InterlockSts = INTERLOCK_STS_DISABLE_PART) Then     ' インターロック一部解除（ステージ動作可能） ?
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    strMSG = "インターロック一部解除中"
                Else
                    strMSG = "Under Interlock Part Release"
                End If
                Me.lblInterLockMSG.Text = strMSG                        '「インターロック一部解除中」表示
                Me.lblInterLockMSG.Visible = True
                'V2.2.0.0⑤ ↓
                If giLoaderType = 1 Then
                    Call Sub_ATLDSET(clsLoaderIf.LOUT_INTLOK_DISABLE, 0)
                End If
                'V2.2.0.0⑤ ↑

            Else                                                        ' インターロック中
                Me.lblInterLockMSG.Visible = False
                'V2.2.0.0⑤ ↓
                If giLoaderType = 1 Then
                    Call Sub_ATLDSET(0, clsLoaderIf.LOUT_INTLOK_DISABLE)
                End If
                'V2.2.0.0⑤ ↑
            End If

            Return (InterlockSts)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "DispInterLockSts() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)
        End Try
    End Function
#End Region

#Region "最前面表示メッセージボックス"
    ''' <summary>
    ''' 最前面表示メッセージボックス
    ''' </summary>
    ''' <param name="DispStr"></param>      ' 表示メッセージ
    ''' <param name="title"></param>        ' 表示タイトル(省略可：デフォルト空白)
    ''' <param name="Button"></param>       ' ボタン種別(省略可：デフォルトOKボタンのみ)
    ''' <returns></returns>                 ' 押したボタンの種別
    ''' <remarks></remarks>
    Public Function MsgBoxForeground(ByVal DispStr As String, Optional ByVal title As String = "", Optional ByVal Button As MessageBoxButtons = vbOKOnly) As DialogResult
        Dim ret As DialogResult
        Using dummyForm As New Form()
            dummyForm.TopMost = False
            dummyForm.Width = 0
            dummyForm.Height = 0
            dummyForm.ControlBox = False
            dummyForm.Show()
            dummyForm.Visible = False
            dummyForm.TopMost = True
            ret = MessageBox.Show(dummyForm, DispStr, title, Button)
        End Using
        Return ret
    End Function
#End Region

#Region "ｿﾌﾄ強制終了時のﾃﾞｰﾀ保存確認"
    '''=========================================================================
    '''<summary>ｿﾌﾄ強制終了時のﾃﾞｰﾀ保存確認</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub AppEndDataSave()

        Dim ret As Short
        Dim strMSG As String

        Try
            ' トリマ動作中信号ON送信(ｵｰﾄﾛｰﾀﾞｰ)
            Call Sub_ATLDSET(COM_STS_TRM_STATE, 0)                      ' ローダー出力(ON=トリマ動作中,OFF=なし)

            'V2.2.0.0①             Call FinalEnd_GazouProc(ObjGazou)                           'DispGazou強制終了

            V_Off()                                                     ' DC電源装置 電圧OFF

            ' 編集中のデータあり ?
            If (FlgUpd = True) Then
                If gSysPrm.stTMN.giMsgTyp = 0 Then
                    ret = MsgBoxForeground("アプリケーションを終了します。" & vbCrLf & "トリミングデータを保存しますか？")
                Else
                    ret = MsgBoxForeground("Quits the program." & vbCrLf & "Do you store trimming data?")
                End If
                If ret = MsgBoxResult.Ok Then
                    ' データ保存
                    Call Me.cmdSave_Click(Me.cmdSave, New System.EventArgs())
                    If gSysPrm.stTMN.giMsgTyp = 0 Then
                        ret = MsgBoxForeground("データの保存が完了しました。" & vbCrLf & "アプリケーションを終了します。")
                    Else
                        ret = MsgBoxForeground("A save of data was completed." & vbCrLf & "Quits the program.")
                    End If
                Else
                    ' データ保存なし
                    If gSysPrm.stTMN.giMsgTyp = 0 Then
                        ret = MsgBoxForeground("アプリケーションを終了します。")
                    Else
                        ret = MsgBoxForeground("Quits the program.")
                    End If
                End If
            Else
                If gSysPrm.stTMN.giMsgTyp = 0 Then
                    ret = MsgBoxForeground("アプリケーションを終了します。")
                Else
                    ret = MsgBoxForeground("Quits the program.")
                End If
            End If
            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "AppEndDataSave() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "ｿﾌﾄ強制終了処理"
    '''=========================================================================
    '''<summary>ｿﾌﾄ強制終了処理</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub AplicationForcedEnding()

        Dim lRet As Integer
        Dim hProcInf As New System.Diagnostics.ProcessStartInfo()
        'Dim ret As Short

        Try
            'V2.2.0.0①            Call FinalEnd_GazouProc(ObjGazou)

            If frmAutoObj.gbFgAutoOperation Then
                lRet = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_WKREMOVE2)
            End If

            ' トリマレディ信号OFF送信(ｵｰﾄﾛｰﾀﾞｰ)
            Call Sub_ATLDSET(COM_STS_TRM_STATE, COM_STS_TRM_READY)          ' ローダー出力(ON=トリマ動作中, ,OFF=トリマレディ)

            ' シグナルタワー初期化(On=0, Off=全ﾋﾞｯﾄ)
            Call Me.System1.SetSignalTower(0, &HFFFFS)

            'V2.2.0.0⑤↓
            ' ローダ通信クローズ 
            If giLoaderType = 1 Then
                ObjSys.ClosePLCThread()
                Call Sub_ATLDSET(&H0, clsLoaderIf.LOUT_AUTO)        ' ローダ手動モード切替え(ローダ出力(ON=なし, OFF=自動))

                ' 電磁ロック(観音扉右側ロック)を解除する
                ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_OFF)

            End If
            'V2.2.0.0⑤↑

            ' スライドカバーオープン/クローズバルブOFF
            If (gSysPrm.stTMN.gsKeimei = MACHINE_TYPE_SL432) Then            ' SL432R系 ? 
                Call ZSLCOVERCLOSE(0)                                       ' スライドカバークローズバルブOFF
                Call ZSLCOVEROPEN(0)                                        ' スライドカバーオープンバルブOFF
            End If

            ' サーボアラームクリア
            Call CLEAR_SERVO_ALARM(1, 1)

            ' ビデオライブラリ終了処理
            If (pbVideoInit = True) Then
                lRet = VideoLibrary1.Close_Library
                If (lRet <> 0) Then
                    Select Case lRet
                        Case cFRS_VIDEO_INI
                            'Call System1.TrmMsgBox(ggSysPrm, "Video library: Not initialized.", MsgBoxStyle.OkOnly, My.Application.Info.Title)
                            Call MsgBox("Video library: Not initialized.", MsgBoxStyle.OkOnly, My.Application.Info.Title) ' 2011.09.01
                        Case Else
                            ' "予期せぬエラー"
                            'Call System1.TrmMsgBox(ggSysPrm, "Video library: Unexpected error.", MsgBoxStyle.OkOnly, My.Application.Info.Title)
                            Call MsgBox("Video library: Unexpected error.", MsgBoxStyle.OkOnly, My.Application.Info.Title) ' 2011.09.01
                    End Select
                End If
            End If

            ' 操作ログ出力("ユーザプログラム終了")
            Call Me.System1.OperationLogging(gSysPrm, MSG_OPLOG_END, "")
            gflgCmpEndProcess = True

            ' クランプ及びバキュームOFF 
            Call Me.System1.ClampVacume_Ctrl(gSysPrm, 0, giAppMode, giTrimErr, False)

            ' ＧＰＩＢ終了処理
            ObjGpib.Gpib_Term(gDevId)

            ' ランプOFF
            Call LAMP_CTRL(LAMP_START, False)                               ' STARTランプOFF 
            Call LAMP_CTRL(LAMP_RESET, False)                               ' RESETランプOFF 
            Call LAMP_CTRL(LAMP_Z, False)                                   ' ZランプOFF 

            ' トリマ動作中信号ON送信(ｵｰﾄﾛｰﾀﾞｰ)念の為
            Call Sub_ATLDSET(COM_STS_TRM_STATE, 0)                          ' ローダー出力(ON=トリマ動作中,OFF=なし)
            '-----------------------------------------------------------------------
            '   クロスラインオブジェクトの解放 ###1040⑥ 
            '-----------------------------------------------------------------------
            'V2.2.0.0① ObjTch.SetCrossLineObject(0)                                   ' ###1040⑥ 
            'V2.2.0.0① ObjMTC.SetCrossLineObject(0)                                   ' ###1040⑥ 

            '-----------------------------------------------------------------------
            '   Mutexの解放
            '-----------------------------------------------------------------------
            gmhUserPro.ReleaseMutex()

            '-----------------------------------------------------------------------
            '   イベントの解放
            '-----------------------------------------------------------------------
            RemoveHandler SystemEvents.SessionEnding, AddressOf SystemEvents_SessionEnding

            '-----------------------------------------------------------------------
            '終了時Videolib関係でエラーが発生するため強制的に外部からアプリを終了させる。
            '-----------------------------------------------------------------------
            hProcInf.FileName = APP_FORCEEND
            hProcInf.Arguments = System.Diagnostics.Process.GetCurrentProcess.ProcessName
            Call System.Diagnostics.Process.Start(hProcInf)
            System.Threading.Thread.Sleep(2000) ' 終了を待たないと次の処理へ進み再起動でINtimeとの不整合が起きて「エアー圧低下検出」が発生する。
        Catch ex As Exception
            ' 操作ログ出力("ユーザプログラム終了")
            Call System1.OperationLogging(gSysPrm, MSG_OPLOG_END, "Exception")
            gflgCmpEndProcess = True
            'MsgBox("Execption error !" & vbCrLf & "error msg = " & ex.Message)
        End Try
    End Sub
#End Region

#Region "frmInfo画面ボタン活性化/非活性化"
    '''=========================================================================
    '''<summary>frmInfo画面ボタン活性化/非活性化</summary>
    '''<param name="Flg">(INP) 0=ボタン非活性化, 1=ボタン活性化</param>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub SBtn_Enb_OnOff(ByRef Flg As Short)

        Dim strMSG As String

        Try
            ' ボタン活性化/非活性化
            If (Flg = 1) Then                                           ' ボタン活性化 ?
                ' ボタン活性化

            Else
                ' ボタン非活性化
            End If

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "SBtn_Enb_OnOff() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "frmInfo画面項目　活性化/非活性化"

    Public Sub Set_UserForm(ByRef Flg As Short)
        ' ボタン活性化/非活性化
        If (Flg = 1) Then                                           ' ボタン活性化 ?
            '' '' ''LblLOTHed.Visible = True
            '' '' ''LblLOT.Visible = True
            '' '' ''_Lbl_1.Visible = True
            '' '' ''_Lbl_2.Visible = True
            '' '' ''_Lbl_3.Visible = True
            '' '' ''_Lbl_4.Visible = True
            '' '' ''_Lbl_5.Visible = True
            '' '' ''_Lbl_6.Visible = True
            '' '' ''_Lbl_7.Visible = True

            '' '' ''LblN_0.Visible = True
            '' '' ''LblN_1.Visible = True
            '' '' ''LblN_2.Visible = True
            '' '' ''LblN_3.Visible = True
            '' '' ''LblN_4.Visible = True

        Else
            '' '' ''LblLOTHed.Visible = False
            '' '' ''LblLOT.Visible = False

            '' '' ''_Lbl_1.Visible = False
            '' '' ''_Lbl_2.Visible = False
            '' '' ''_Lbl_3.Visible = False
            '' '' ''_Lbl_4.Visible = False
            '' '' ''_Lbl_5.Visible = False
            '' '' ''_Lbl_6.Visible = False
            '' '' ''_Lbl_7.Visible = False

            '' '' ''LblN_0.Visible = False
            '' '' ''LblN_1.Visible = False
            '' '' ''LblN_2.Visible = False
            '' '' ''LblN_3.Visible = False
            '' '' ''LblN_4.Visible = False


        End If

    End Sub

#End Region

#Region "ボタン活性化/非活性化"
    '''=========================================================================
    '''<summary>ボタン活性化/非活性化</summary>
    '''<param name="Flg">(INP) 0=ボタン非活性化, 1=ボタン活性化, 2=ボタン等の表示/非表示</param>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub Btn_Enb_OnOff(ByRef Flg As Short)

        Dim strMSG As String

        Try
            'LblDIGSW_HI.Visible = True
            LblDIGSW_HI.Visible = False
            LblDIGSW_LO.Visible = True                      ' "DSW="　表示
            'CbDigSwH.Visible = True
            CbDigSwH.Visible = False
            CbDigSwL.Visible = True
            '---------------------------------------------------------------------------
            '   ボタンを活性化する
            '---------------------------------------------------------------------------
            If (Flg = 1) Then

                'V2.1.0.0②↓
                If UserSub.IsLaserCaribrarionUse() Then
                    ButtonLaserCalibration.Enabled = True
                End If
                'V2.1.0.0②↑

                cmdHelp.Enabled = True                      ' HELP(About)
                cmdExit.Enabled = True                      ' END(F12)
                'cmdStart.Enabled = True                     ' START(Debug)

                If (stFNC(F_LOAD).iDEF = 0) Then            ' LOAD(F1)
                    cmdLoad.Enabled = False
                Else
                    cmdLoad.Enabled = True
                End If
                If (stFNC(F_SAVE).iDEF = 0) Then            ' SAVE(F2)
                    cmdSave.Enabled = False
                Else
                    cmdSave.Enabled = True
                End If
                If (stFNC(F_EDIT).iDEF = 0) Then            ' EDIT(F3)
                    cmdEdit.Enabled = False
                Else
                    cmdEdit.Enabled = True
                End If
                'If (stFNC(F_MSTCHK).iDEF = 0) Then         ' ﾏｽﾀﾁｪｯｸ(F4)
                '    cmdMstChk.Enabled = False
                'Else
                '    cmdMstChk.Enabled = True
                'End If
                If (stFNC(F_LASER).iDEF = 0) Then           ' LASER(F5)
                    cmdLaserTeach.Enabled = False
                    cmdLaserCalibration.Enabled = False     'V2.1.0.0②
                Else
                    cmdLaserTeach.Enabled = True
                    'V2.1.0.0②↓
                    If UserSub.IsLaserCaribrarionUse Then
                        cmdLaserCalibration.Enabled = True
                    End If
                    'V2.1.0.0②↑
                End If
                If (stFNC(F_LOTCHG).iDEF = 0) Then          ' ﾛｯﾄ切替(S-F6)
                    cmdLotChg.Enabled = False
                Else
                    cmdLotChg.Enabled = True
                End If
                If (stFNC(F_PROBE).iDEF = 0) Then           ' PROBE(F7)
                    cmdProbeTeaching.Enabled = False
                Else
                    cmdProbeTeaching.Enabled = True
                End If
                If (stFNC(F_TEACH).iDEF = 0) Then           ' TEACH(F8)
                    cmdTeaching.Enabled = False
                Else
                    cmdTeaching.Enabled = True
                End If
                If (stFNC(F_CUTPOS).iDEF = 0) Then          ' CutPosTeach(S-F8)
                    cmdCutPosTeach.Enabled = False
                Else
                    cmdCutPosTeach.Enabled = True
                End If
                If (stFNC(F_RECOG).iDEF = 0) Then          ' RECOG(F9)
                    BtnRECOG.Enabled = False
                Else
                    BtnRECOG.Enabled = True
                End If
                'V2.0.0.0①↓
                If (stFNC(F_TX).iDEF > 0) Then          ' TX(F9)
                    CmdTx.Enabled = True
                Else
                    CmdTx.Enabled = False
                End If
                If (stFNC(F_TY).iDEF > 0) Then          ' TY(F10)
                    CmdTy.Enabled = True
                Else
                    CmdTy.Enabled = False
                End If
                'V2.0.0.0①↑

                BtnStartPosSet.Enabled = True           'V2.0.0.0②

                cmdClamp.Enabled = True                    'V2.2.1.1⑨

                ' ユーザー特殊処理 START
                cmdLotInfo.Enabled = True
                cmdPrint.Enabled = True
                ' ユーザー特殊処理 END
                '---------------------------------------------------------------------------
                '   ボタンを非活性化する
                '---------------------------------------------------------------------------
            ElseIf (Flg = 0) Then

                ButtonLaserCalibration.Enabled = False      'V2.1.0.0②

                cmdHelp.Enabled = False                     ' HELP(About)
                cmdStart.Enabled = False                    ' START(Debug) 
                cmdExit.Enabled = False                     ' END(F10)
                cmdLoad.Enabled = False                     ' LOAD(F1)
                cmdSave.Enabled = False                     ' SAVE(F2)
                cmdEdit.Enabled = False                     ' EDIT(F3)
                'cmdMstChk.Enabled = False                  ' ﾏｽﾀﾁｪｯｸ(F4)
                cmdLaserTeach.Enabled = False               ' LASER(F5)
                cmdLaserCalibration.Enabled = False         'V2.1.0.0② レーザキャリブレーション
                cmdLotChg.Enabled = False                   ' ﾛｯﾄ切替(S-F6)
                cmdProbeTeaching.Enabled = False            ' PROBE(F7)
                cmdTeaching.Enabled = False                 ' TEACH(F8)
                cmdCutPosTeach.Enabled = False              ' CutPosTeach(S-F8)
                BtnRECOG.Enabled = False                    ' RECOG(F9)
                'V2.0.0.0①↓
                CmdTx.Enabled = False                       ' TX(F9)
                CmdTy.Enabled = False                       ' TY(F10)
                'V2.0.0.0①↑

                BtnStartPosSet.Enabled = False              'V2.0.0.0②

                cmdClamp.Enabled = False                    'V2.2.1.1⑨

                ' ユーザー特殊処理 START
                cmdLotInfo.Enabled = False
                cmdPrint.Enabled = False
                ' ユーザー特殊処理 END
                '---------------------------------------------------------------------------
                '   ボタン等の表示/非表示を設定する
                '---------------------------------------------------------------------------
            Else
                Grpcmds.Visible = True
                frmInfo.Visible = True                      ' トリミング結果表示フレーム
                'txtLog.Visible = True
                If giTxtLogType <> 0 Then
                    txtlog.Visible = True                       ' 'V2.2.0.0⑰
                Else
                    lstLog.Visible = True                       ' ###lstLog
                End If

#If cOFFLINEcDEBUG Then
                'cmdStart.Visible = True
#Else
                cmdStart.Visible = False
#End If
                cmdExit.Visible = True                      ' END(F12)
                If (stFNC(F_LOAD).iDEF >= 0) Then           ' LOAD(F1)
                    cmdLoad.Visible = True
                Else
                    cmdLoad.Visible = False
                End If
                If (stFNC(F_SAVE).iDEF >= 0) Then           ' SAVE(F2)
                    cmdSave.Visible = True
                Else
                    cmdSave.Visible = False
                End If
                If (stFNC(F_EDIT).iDEF >= 0) Then           ' EDIT(F3)
                    cmdEdit.Visible = True
                Else
                    cmdEdit.Visible = False
                End If
                'If (stFNC(F_MSTCHK).iDEF >= 0) Then        ' ﾏｽﾀﾁｪｯｸ(F4)
                '    cmdMstChk.Visible = True
                'Else
                '    cmdMstChk.Visible = False
                'End If
                If (stFNC(F_LASER).iDEF >= 0) Then          ' LASER(F5)
                    cmdLaserTeach.Visible = True
                    'V2.1.0.0②↓
                    If UserSub.IsLaserCaribrarionUse Then
                        cmdLaserCalibration.Visible = True
                    End If
                    'V2.1.0.0②↑
                Else
                    cmdLaserTeach.Visible = False
                    cmdLaserCalibration.Visible = False     'V2.1.0.0②
                End If
                If (stFNC(F_LOTCHG).iDEF >= 0) Then         ' ﾛｯﾄ切替(S-F6)
                    cmdLotChg.Visible = True
                Else
                    cmdLotChg.Visible = False
                End If
                If (stFNC(F_PROBE).iDEF >= 0) Then          ' PROBE(F7)
                    cmdProbeTeaching.Visible = True
                Else
                    cmdProbeTeaching.Visible = False
                End If
                If (stFNC(F_TEACH).iDEF >= 0) Then          ' TEACH(F8)
                    cmdTeaching.Visible = True
                Else
                    cmdTeaching.Visible = False
                End If
                If (stFNC(F_CUTPOS).iDEF >= 0) Then         ' CutPosTeach(S-F8)
                    cmdCutPosTeach.Visible = True
                Else
                    cmdCutPosTeach.Visible = False
                End If
                If (stFNC(F_RECOG).iDEF >= 0) Then          ' RECOG(F9)
                    BtnRECOG.Visible = True
                Else
                    BtnRECOG.Visible = False
                End If

                'V2.0.0.0①↓
                If (stFNC(F_TX).iDEF > 0) Then          ' TX(F9)
                    CmdTx.Visible = True
                Else
                    CmdTx.Visible = False
                End If
                If (stFNC(F_TY).iDEF > 0) Then          ' TY(F10)
                    CmdTy.Visible = True
                Else
                    CmdTy.Visible = False
                End If
                'V2.0.0.0①↑

                cmdClamp.Enabled = True                    'V2.2.1.1⑨

            End If

STP_END:

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Btn_Enb_OnOff() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "レーザパワー調整関連項目の表示/非表示設定"
    '''=========================================================================
    '''<summary>レーザパワー調整関連項目の表示/非表示設定</summary>
    ''' <param name="Md">(INP)0=表示しない, 1=表示する</param>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub SetLaserItemsVisible(ByVal Md As Integer)

        Dim strMSG As String

        Try
            ' 減衰率をシスパラより表示する("減衰率 = 99.9%")
            Me.LblRotAtt.Visible = False                                ' 減衰率非表示
            If (Md = 1) Then                                            ' 表示する ?
                If (gSysPrm.stRMC.giRmCtrl2 >= 2 And
                    gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then       ' ﾛｰﾀﾘｱｯﾃﾈｰﾀ制御有でFLでない ?
                    Me.LblRotAtt.Visible = True
                End If
            End If

            ' 測定値をシスパラより表示する
            Me.LblMes.Visible = False                                   ' 測定値非表示
            If (Md = 1) Then                                            ' 表示する ?
                ' RMCTRL2 >=3 で 測定値表示 ?
                If (gSysPrm.stRMC.giRmCtrl2 >= 3) And (gSysPrm.stRMC.giPMonHi = 1) Then
                    LblMes.Visible = False                              ' 測定値非表示
                End If
            End If

            ' 定電流値を表示する
            LblCur.Visible = False                                      ' 定電流値非表示
            If (Md = 1) Then                                            ' 表示する ?
                ' 加工電力設定 = 4(定電流1A)の時に表示する
                If (gSysPrm.stSPF.giProcPower = 4) And (gSysPrm.stSPF.giProcPower2 <> 0) Then
                    LblCur.Visible = True
                End If
            End If

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "SetLaserItemsVisible() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "操作ログ出力サブ"
    '''=========================================================================
    '''<summary>操作ログ出力サブ</summary>
    '''<param name="gSts">ｱﾌﾟﾘﾓｰﾄﾞ(giAppMode参照)</param>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub Sub_OprLog(ByRef gSts As Short)

        Dim strMSG As String

        Try
            ' ログメッセージ設定
            Select Case (gSts)
                Case APP_MODE_LASER
                    strMSG = MSG_OPLOG_FUNC05       ' "レーザ調整"
                Case APP_MODE_PROBE
                    strMSG = MSG_OPLOG_FUNC07       ' "プローブ位置合わせ"
                Case APP_MODE_PROBE2
                    strMSG = MSG_OPLOG_FUNC10       ' "プローブ位置合わせ２"
                Case APP_MODE_TEACH
                    strMSG = MSG_OPLOG_FUNC08       ' "ティーチング"
                Case APP_MODE_CUTPOS
                    strMSG = MSG_OPLOG_FUNC08S      ' "カット補正位置ティーチング"
                Case APP_MODE_RECOG
                    strMSG = MSG_OPLOG_FUNC09       ' "パターン登録"
                Case Else
                    Exit Sub
            End Select

            ' 操作ログ出力
            Call Me.System1.OperationLogging(gSysPrm, strMSG, "MANUAL")

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Sub_OprLog() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "ホストコマンド　シミュレーション(DEBUG用)"
    '''=========================================================================
    '''<summary>ホストコマンド　シミュレーション(DEBUG用)</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub DEBUG_HST_CMD_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles DEBUG_HST_CMD.Click

        Dim Index As Short = DEBUG_HST_CMD.GetIndex(eventSender)

        ' ホストコマンド　シミュレーション(DEBUG用)
        Call DEBUG_ReadHostCommand(Index)               ' ローダ入力サブ(デバッグ用)

    End Sub
#End Region

    '========================================================================================
    '   各コマンド(ユーザコントロールにフォームがあるOCXの場合)がフォーカスを失った場合の処理
    '   テンキーのUP/Downイベントが入ってこなくなるためOCXにフォーカを設定する
    '========================================================================================
#Region "ビデオ画像をクリックしてフォーカスを失った場合の処理"
    '''=========================================================================
    ''' <summary>ビデオ画像をクリックしてフォーカスを失った場合の処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>・OcxTeachはDispGazou.EXE実行中のためEnterイベントは入ってこない
    '''          　DispGazou.EXEはOcxTeachで起動する
    '''          ・EnterイベントはFormのACTIVEコントロールになった時に発生</remarks>
    '''=========================================================================
    Private Sub VideoLibrary1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try
            CmdSetFocus()

        Catch ex As Exception
        End Try
    End Sub
#End Region

#Region "ログ表示域描画時"
    ''' <summary>ListBox複数行表示・折り返し表示</summary>
    ''' <remarks>###lstLog</remarks>
    Private Sub lstLog_DrawItem(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles lstLog.DrawItem
        Dim lstBox As ListBox = DirectCast(sender, ListBox)
        With lstBox
            If (0 = .Items.Count) OrElse (.Items Is Nothing) OrElse (e.Index < 0) Then Exit Sub
            Dim strItem As String = .GetItemText(.Items(e.Index))
            e.DrawBackground()
            e.Graphics.DrawString(strItem, e.Font, New SolidBrush(e.ForeColor), e.Bounds)
            e.DrawFocusRectangle()
        End With
    End Sub

    ''' <summary>ListBox複数行表示・折り返し表示</summary>
    ''' <remarks>###lstLog</remarks>
    Private Sub lstLog_MeasureItem(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MeasureItemEventArgs) Handles lstLog.MeasureItem
        Dim lstBox As ListBox = DirectCast(sender, ListBox)
        With lstBox
            If (0 = .Items.Count) OrElse (.Items Is Nothing) OrElse (e.Index < 0) Then Exit Sub
            Dim strItem As String = .GetItemText(.Items(e.Index))
            Dim z As SizeF = e.Graphics.MeasureString(
                strItem, .Font, Convert.ToInt32(e.Graphics.VisibleClipBounds.Width))
            e.ItemWidth = Convert.ToInt32(z.Width)
            e.ItemHeight = Convert.ToInt32(z.Height)
        End With
    End Sub
#End Region

#Region "ログ表示域クリック時"
    ''V2.2.0.0⑰↓    有効にする
    Private Sub txtLog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtlog.Click
        Try
            CmdSetFocus()

        Catch ex As Exception
        End Try
    End Sub
    ''V2.2.0.0⑰↑

    Private Sub lstLog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstLog.Click
        Try
            CmdSetFocus()

        Catch ex As Exception
        End Try
    End Sub

    Private Sub lstLog_Copy(ByVal sender As Object, ByVal e As EventArgs)
        Dim sb As New StringBuilder(256)
        For Each item As Object In Me.lstLog.SelectedItems
            sb.Append(item)
        Next
        If (0 < sb.Length) Then
            Clipboard.SetText(sb.ToString())
        End If
    End Sub

    ''V2.2.0.0⑰↓
    Private Sub txtLog_Copy(ByVal sender As Object, ByVal e As EventArgs)
        Dim sb As New StringBuilder(256)
        For Each item As Object In Me.txtlog.Lines
            sb.Append(item)
        Next
        If (0 < sb.Length) Then
            Clipboard.SetText(sb.ToString())
        End If
    End Sub
    ''V2.2.0.0⑰↑



#End Region

#Region "各コマンド(OCX)にフォーカスを設定する"
    '''=========================================================================
    ''' <summary>各コマンド(OCX)にフォーカスを設定する</summary>
    '''=========================================================================
    Private Sub CmdSetFocus()

        Dim strMSG As String

        Try
            Select Case (giAppMode)
                Case APP_MODE_PROBE
                    ' プローブコマンド実行中 ?
                    Probe1.Focus()                                      ' OcxProbeにフォーカスをセットする 

                Case APP_MODE_TEACH
                    ' ティーチコマンド実行中 ?
                    Teaching1.Focus()                                   ' OcxTeachにフォーカスをセットする 
                    Teaching1.JogSetFocus()                             ' ←Probeと違い何故かこれがないとテンキーが効かない
            End Select

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "frmMain.CmdSetFocus() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

    Private Sub cmdLotInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLotInfo.Click

        Dim s As String
        Dim strMSG As String

        Try
            ' トリマ動作中信号ON送信(ｵｰﾄﾛｰﾀﾞｰ)
            Call Sub_ATLDSET(COM_STS_TRM_STATE, 0)                      ' ローダー出力(ON=トリマ動作中,OFF=なし)

            ' トリマ装置アイドル中以外ならNOP
            If giAppMode Then Exit Sub
            giAppMode = APP_MODE_LOTNO                                   ' ｱﾌﾟﾘﾓｰﾄﾞ = ロット番号設定中

            ' データロードチェック
            If pbLoadFlg = False Then
                s = MSG_DataNotLoad                                     ' ﾃﾞｰﾀ未ﾛｰﾄﾞ
                Call Z_PRINT(s)
                Call Beep()
                GoTo STP_END
            End If

            If Not UserSub.IsSpecialTrimType Then
                Call Z_PRINT("トリミングデータの製品種別が指定無しに設定されています。" & vbCrLf)
                Call Beep()
                GoTo STP_END
            End If
            ' データ編集
            Call Me.System1.OperationLogging(gSysPrm, MSG_OPLOG_LOTSET, "")

            ChkLoaderInfoDisp(0)                              'V2.2.0.0⑤

            Dim Rtn As Short
            Dim fLotInf As New FormEdit.frmLotInfoInput(True)
            fLotInf.ShowDialog(Me)
            Rtn = fLotInf.sGetReturn()
            fLotInf.Dispose()
            If Rtn = cFRS_ERR_START Then                                ' ＯＫリターン
                Call UserSub.SetStartCheckStatus(True)                  ' 設定画面の確認有効化
            End If

STP_END:
            Call ZCONRST()                                              ' ｺﾝｿｰﾙｷｰ ﾗｯﾁ解除
            giAppMode = APP_MODE_IDLE                                   ' ｱﾌﾟﾘﾓｰﾄﾞ = トリマ装置アイドル中
            Call Sub_ATLDSET(0, COM_STS_TRM_STATE)                      ' ローダー出力(ON=なし,OFF=トリマ動作中)

            ChkLoaderInfoDisp(1)                              'V2.2.0.0⑤

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "cmdLotInfo_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        Finally
            giAppMode = APP_MODE_IDLE                                   ' ｱﾌﾟﾘﾓｰﾄﾞ = トリマ装置アイドル中
        End Try

    End Sub

#Region "印刷ﾎﾞﾀﾝｸﾘｯｸｲﾍﾞﾝﾄ"
    '''=========================================================================
    ''' <summary>印刷ﾎﾞﾀﾝｸﾘｯｸｲﾍﾞﾝﾄ</summary>
    '''=========================================================================
    Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrint.Click
        Call UserSub.LotEnd()                                   ' ロット終了時のデータ出力
        Call Printer.Print(True)                                ' 確認ﾒｯｾｰｼﾞの戻り値により印刷をおこなう
        Call UserSub.SetStartCheckStatus(True)                  ' 印刷を行ったらロット終了とみなす。
    End Sub
#End Region

    '==========================================================================
    '   ステップ移動処理
    '==========================================================================
#Region "ステップ移動ボタン表示"
    Public Sub StepMoveButtonOn()
        BtnForward.Enabled = True
        BtnForward.Visible = True
        BtnBackword.Enabled = True
        BtnBackword.Visible = True
        gbAdjOnStatus = True        ' ＡＤＪ停止中
    End Sub
#End Region

#Region "ステップ移動ボタン非表示"
    Public Sub StepMoveButtonOff()
        BtnForward.Enabled = False
        BtnForward.Visible = False
        BtnBackword.Enabled = False
        BtnBackword.Visible = False
        gbAdjOnStatus = False       ' ＡＤＪ非停止中
    End Sub
#End Region
#Region "ステップ移動ボタン処理"
    Private Sub BtnForward_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnForward.Click
        Call UserBas.StepMove(1)
    End Sub

    Private Sub BtnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBackword.Click
        Call UserBas.StepMove(-1)
    End Sub
#End Region
    'V2.0.0.0⑨↓
    ''' <summary>
    ''' 再測定の開始位置の指定ボタン処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnStartPosSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnStartPosSet.Click
        Try
            Dim fReStartPosSet As New formReStartPosSet
            If giAppMode <> APP_MODE_IDLE Then
                Return
            Else
                giAppMode = APP_MODE_EDIT                               ' アプリモード = データ編集
            End If

            ' トリマ動作中信号ON送信(ｵｰﾄﾛｰﾀﾞｰ)
            Call Sub_ATLDSET(COM_STS_TRM_STATE, 0)                      ' ローダー出力(ON=トリマ動作中,OFF=なし)

            fReStartPosSet.ShowDialog(Me)

            giAppMode = APP_MODE_IDLE                                   ' ｱﾌﾟﾘﾓｰﾄﾞ = トリマ装置アイドル中
            Call Sub_ATLDSET(0, COM_STS_TRM_STATE)                      ' ローダー出力(ON=なし,OFF=トリマ動作中)

        Catch ex As Exception
            MsgBox("frmMain.BtnStartPosSet_Click() TRAP ERROR = " + ex.Message)
        End Try
    End Sub
    'V2.0.0.0⑨↑

    'V2.0.0.0⑥↓
    Private Sub CbDigSwL_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles CbDigSwL.MouseWheel

        Dim eventArgs As HandledMouseEventArgs = DirectCast(e, HandledMouseEventArgs)
        eventArgs.Handled = True

    End Sub
    'V2.0.0.0⑥↑
    'V2.0.0.0⑨↓
#Region "グラフ表示/非表示ボタン押下時処理"
    Public Sub changefrmDistStatus(ByVal DispOnOff As Integer)
        Try

            If (DispOnOff = 1) Then
                '統計表示のON
                gObjFrmDistribute.Show()
                gObjFrmDistribute.RedrawGraph()  '###218 
                'ボタン表示の変更
                'If (gSysPrm.stTMN.giMsgTyp = 0) Then
                '    chkDistributeOnOff.Text = "生産グラフ　非表示"
                'Else
                '    chkDistributeOnOff.Text = "Distribute OFF"
                'End If
                chkDistributeOnOff.Text = Form1_019
            Else
                '統計表示のOFF
                gObjFrmDistribute.hide()

                'ボタン表示の変更
                'If (gSysPrm.stTMN.giMsgTyp = 0) Then
                '    chkDistributeOnOff.Text = "生産グラフ　表示"
                'Else
                '    chkDistributeOnOff.Text = "Distribute ON"
                'End If
                chkDistributeOnOff.Text = Form1_020
            End If

            Exit Sub

        Catch ex As Exception
            MsgBox("changefrmDistStatus() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 統計画面ボタンの有効化無効化
    ''' </summary>
    ''' <param name="Flag"></param>
    ''' <remarks></remarks>
    Public Sub chkDistributeOnOffEnableSet(ByVal Flag As Boolean)
        Try
            chkDistributeOnOff.Enabled = Flag
            CCmb_DistributeResList.Enabled = Flag
        Catch ex As Exception
            MsgBox("chkDistributeOnOffEnableSet() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 統計画面表示非表示
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub chkDistributeOnOff_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDistributeOnOff.CheckedChanged
        Try
            If chkDistributeOnOff.Checked = True Then
                ''統計表示のON
                ' 統計画面ボタンを有効する
                gObjFrmDistribute.cmdGraphSave.Enabled = True
                gObjFrmDistribute.cmdInitial.Enabled = True
                gObjFrmDistribute.cmdFinal.Enabled = True
                CCmb_DistributeResList.Enabled = False
                changefrmDistStatus(1)
            Else
                ''統計表示のOFF
                CCmb_DistributeResList.Enabled = True
                changefrmDistStatus(0)
            End If
        Catch ex As Exception
            MsgBox("chkDistributeOnOff_CheckedChanged() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Sub

    Private Sub CCmb_DistributeResList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CCmb_DistributeResList.SelectedIndexChanged
        Try
            If (Not gObjFrmDistribute Is Nothing) Then
                gObjFrmDistribute.SetDistributionResNo(CCmb_DistributeResList.SelectedIndex + 1)
                stPLT.DistributionResNo = CCmb_DistributeResList.SelectedIndex + 1
                StatisticalDataDisp()
            End If
        Catch ex As Exception
            MsgBox("CCmb_DistributeResList_SelectedIndexChanged() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 分布図のボタンが表示か非表示か返す。
    ''' </summary>
    ''' <returns>表示:True 非表示:False</returns>
    ''' <remarks></remarks>
    Public Function GetDistributeOnOffStatus() As Boolean
        Try
            If chkDistributeOnOff.Checked = True Then
                Return (True)
            Else
                Return (False)
            End If
        Catch ex As Exception
            MsgBox("GetDistributeOnOffStatus() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Function
    ''' <summary>
    ''' 分布図の表示、非表示
    ''' </summary>
    ''' <param name="Flag">1:表示 0:非表示</param>
    ''' <remarks></remarks>
    Public Sub DistributeOnOff(ByVal Flag As Integer)
        Try
            If Flag = 1 Then
                chkDistributeOnOff.Checked = True
            Else
                chkDistributeOnOff.Checked = False
            End If
        Catch ex As Exception
            MsgBox("DistributeOnOff() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Sub

#End Region

#Region "統計データの表示更新"
    Public Sub StatisticalDataDisp()
        Try
            Dim JudgeMode As Integer = FINAL_TEST
            Dim dMin As Double, dMax As Double, dAve As Double, dDev As Double

            If (Not gObjFrmDistribute Is Nothing) Then

                Call gObjFrmDistribute.StatisticalDataGet(JudgeMode, stPLT.DistributionResNo, dMin, dMax, dAve, dDev)

                Me.LabelStaticNom.Text = stREG(GetRNumByCircuit(1, stPLT.DistributionResNo)).dblNOM.ToString(TARGET_DIGIT_DEFINE)
                If dMin.ToString(TARGET_DIGIT_DEFINE).Length > 11 Then
                    Me.LabelStaticMin.Text = dMin.ToString(TARGET_DIGIT_DEFINE).Substring(0, 11)
                Else
                    Me.LabelStaticMin.Text = dMin.ToString(TARGET_DIGIT_DEFINE)
                End If
                If dMax.ToString(TARGET_DIGIT_DEFINE).Length > 11 Then
                    Me.LabelStaticMax.Text = dMax.ToString(TARGET_DIGIT_DEFINE).Substring(0, 11)
                Else
                    Me.LabelStaticMax.Text = dMax.ToString(TARGET_DIGIT_DEFINE)
                End If
                If dAve.ToString(TARGET_DIGIT_DEFINE).Length > 11 Then
                    Me.LabelStaticAve.Text = dAve.ToString(TARGET_DIGIT_DEFINE).Substring(0, 11)
                Else
                    Me.LabelStaticAve.Text = dAve.ToString(TARGET_DIGIT_DEFINE)
                End If
                If dDev.ToString(TARGET_DIGIT_DEFINE).Length > 11 Then
                    Me.LabelStaticDev.Text = dDev.ToString(TARGET_DIGIT_DEFINE).Substring(0, 11)
                Else
                    Me.LabelStaticDev.Text = dDev.ToString(TARGET_DIGIT_DEFINE)
                End If
            Else
                Me.LabelStaticMin.Text = ""
                Me.LabelStaticMax.Text = ""
                Me.LabelStaticAve.Text = ""
                Me.LabelStaticDev.Text = ""
            End If

        Catch ex As Exception
            MsgBox("frmMain.StatisticalDataDisp() TRAP ERROR = " + ex.Message)
        End Try
    End Sub
#End Region
    'V2.0.0.0⑨↑

    'V2.1.0.0②↓
#Region "レーザキャリブレーション"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cmdLaserCalibration_Click(sender As System.Object, e As System.EventArgs) Handles cmdLaserCalibration.Click
        Try
            cmdLaserTeach_Calibration()
        Catch ex As Exception
            MsgBox("frmMain.cmdLaserCalibration_Click() TRAP ERROR = " + ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' レーザパワーモニタリングチェックボタン押下時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ButtonLaserCalibration_Click(sender As System.Object, e As System.EventArgs) Handles ButtonLaserCalibration.Click
        Try
            Dim LaserCalibrationMode As Integer = UserSub.LaserCalibrationModeGet()

            Select Case (LaserCalibrationMode)
                Case POWER_CHECK_NONE
                    LaserCalibrationMode = POWER_CHECK_START
                Case POWER_CHECK_START
                    LaserCalibrationMode = POWER_CHECK_LOT
                Case POWER_CHECK_LOT
                    LaserCalibrationMode = POWER_CHECK_NONE
            End Select

            UserSub.LaserCalibrationModeSet(LaserCalibrationMode)

            UserSub.LaserCalibrationModeUpdate()

        Catch ex As Exception
            MsgBox("frmMain.ButtonLaserCalibration_Click() TRAP ERROR = " + ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' ローダ情報の画面表示
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnLoaderInfo_Click(sender As Object, e As EventArgs) Handles btnLoaderInfo.Click

        Try


            If IsNothing(objLoaderInfo) = True Then
                Return
            End If

            If btnLoaderInfo.BackColor = Color.LightGreen Then
                btnLoaderInfo.BackColor = SystemColors.Control
                objLoaderInfo.Hide()
                objLoaderInfo.saveLoaderInfoDisp = 0
            Else
                btnLoaderInfo.BackColor = Color.LightGreen
                objLoaderInfo.saveLoaderInfoDisp = 1
                ObjLoader.DispLoaderInfo()

                objLoaderInfo.Show(Me)
            End If

        Catch ex As Exception

        End Try


    End Sub

    'V2.2.0.0⑤↓
    ''' <summary>
    ''' LoaderInfo画面の表示状態を保存、取得するして状態を合わせる  
    ''' </summary>
    ''' <param name="mode"></param>
    ''' <returns></returns>
    Public Function ChkLoaderInfoDisp(ByVal mode As Integer) As Integer

        Try

            If giLoaderType = 0 Then
                Return 0
            End If

            If mode = 0 Then        ' 非表示
                If btnLoaderInfo.BackColor = Color.LightGreen Then
                    objLoaderInfo.saveLoaderInfoDisp = 1    ' 
                End If
                objLoaderInfo.Hide()
            Else                    ' 表示
                If objLoaderInfo.saveLoaderInfoDisp = 1 Then
                    objLoaderInfo.Show()
                End If
            End If

        Catch ex As Exception

        End Try

    End Function
    'V2.2.0.0⑤↑



    ''' <summary>
    ''' カット毎の停止機能ボタンを押した処理  'V2.2.0.0⑥
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnCutStop_Click(sender As Object, e As EventArgs) Handles btnCutStop.Click
        Try
            If giCutStop = 0 Then
                Return
            End If

            If btnCutStop.BackColor = Color.Yellow Then
                btnCutStop.BackColor = SystemColors.Control
            Else
                btnCutStop.BackColor = Color.Yellow
            End If


        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "STOPボタンの状態を判定する "
    ''' <summary>
    ''' STOPボタンの状態を判定する    'V2.2.0.0⑥
    ''' </summary>
    ''' <returns></returns>
    Public Function JudgeStop() As Integer

        Try

            If btnCutStop.BackColor = Color.Yellow Then
                JudgeStop = True
            Else
                JudgeStop = False
            End If


        Catch ex As Exception

        End Try

    End Function
#End Region

#Region "サイクル停止ボタンを押した時の処理"
    ''' <summary>　
    ''' サイクル停止ボタンを押した時の処理　'V2.2.0.0⑦
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnCycleStop_Click(sender As Object, e As EventArgs) Handles btnCycleStop.Click

        Try

            If btnCycleStop.BackColor = Color.Yellow Then
                btnCycleStop.BackColor = SystemColors.Control
            Else
                btnCycleStop.BackColor = Color.Yellow
            End If


        Catch ex As Exception

        End Try
    End Sub

#End Region

#Region "CYCLE STOPボタンの状態を判定する "
    ''' <summary>
    ''' CYCLE STOPボタンの状態を判定する    'V2.2.0.0⑦
    ''' </summary>
    ''' <returns></returns>
    Public Function JudgeCycleStop() As Integer

        Try

            If giClcleStop = 0 Then
                Return 0
            End If

            If btnCycleStop.BackColor = Color.Yellow Then
                JudgeCycleStop = 1
            Else
                JudgeCycleStop = 0
            End If


        Catch ex As Exception

        End Try

    End Function
#End Region


    ''' <summary>
    ''' レーザOFFボタンを押したときの処理 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnLaserOff_Click(sender As Object, e As EventArgs) Handles btnLaserOff.Click
        Dim ret As Integer
        Dim mode As Integer = 0

        Try

            If btnLaserOff.BackColor = Color.Red Then
                btnLaserOff.BackColor = SystemColors.Control
                giLaserOffMode = 0
                mode = 1
            Else
                btnLaserOff.BackColor = Color.Red
                giLaserOffMode = 1
                mode = 0
            End If
            ret = DefTrimFnc.SPLASER_EXTDIODESET(mode)

        Catch ex As Exception

        End Try


    End Sub

    ''' <summary>
    ''' 画面上のボタンからクランプの閉⇒開動作
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub cmdClamp_Click(sender As Object, e As EventArgs) Handles cmdClamp.Click
        Dim r As Integer

        Try

            cmdClamp.BackColor = Color.Yellow
            cmdClamp.Enabled = False

            ' 載物台クランプON   
            r = System1.ClampCtrl(gSysPrm, 1, 0)
            If (r <> cFRS_NORMAL) Then

            End If

            Sleep(500)

            ' 載物台クランプOFF 
            r = System1.ClampCtrl(gSysPrm, 0, 0)
            If (r <> cFRS_NORMAL) Then

            End If

        Catch ex As Exception


        Finally
            cmdClamp.Enabled = True
            cmdClamp.BackColor = SystemColors.ButtonFace

        End Try



    End Sub
    'V2.1.0.0②↑

    'V2.2.1.7③↓
    ''' <summary>
    ''' アラームでマーク印字しなかった基板の一覧を画面ログに表示 
    ''' </summary>
    Public Sub DispMarkAlarmList()
        Dim i As Integer

        Try
            'V2.2.1.7⑥ ↓
            ' マーク印字で無ければアラーム表示しない。 
            If UserSub.IsTrimType5() <> True Then
                Return
            End If
            'V2.2.1.7⑥ ↑

            If LotMarkingAlarmCnt > 0 Then

                Call Z_PRINT("マーク印字時アラームリスト" & vbCrLf)

                For i = 1 To LotMarkingAlarmCnt
                    Call Z_PRINT(gMarkAlarmList(i).AlarmTrimData & ":" & gMarkAlarmList(i).LotCount & "枚目" & vbCrLf)
                Next

            Else
                'Call Z_PRINT("マーク印字：全完了" & vbCrLf)
            End If

        Catch ex As Exception

        End Try


    End Sub
    'V2.2.1.7③↑

End Class

#Region "各コマンド実行サブフォーム用共通インターフェース"
''' <summary>各コマンド実行サブフォーム用共通インターフェース</summary>
''' <remarks>'V2.2.0.0①</remarks>
Public Interface ICommonMethods
    ''' <summary>サブフォーム処理実行</summary>
    ''' <returns>実行結果 sGetReturn</returns>
    ''' <remarks>'V2.2.0.0①</remarks>
    Function Execute() As Integer

    ''' <summary>サブフォームKeyDown時の処理</summary>
    ''' <param name="e"></param>
    Sub JogKeyDown(ByVal e As KeyEventArgs)

    ''' <summary>サブフォームKeyUp時の処理</summary>
    ''' <param name="e"></param>
    Sub JogKeyUp(ByVal e As KeyEventArgs)

    ''' <summary>カメラ画像クリック位置を画像センターに移動する処理</summary>
    ''' <param name="distanceX">画像センターからの距離X</param>
    ''' <param name="distanceY">画像センターからの距離Y</param>
    Sub MoveToCenter(ByVal distanceX As Decimal, ByVal distanceY As Decimal)
End Interface
#End Region

