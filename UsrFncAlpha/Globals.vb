'===============================================================================
'   Description : グローバル定数の定義
'
'   Copyright(C): TOWA LASERFRONT CORP. 2018
'
'===============================================================================
Option Strict Off
Option Explicit On

Imports System.IO
Imports System.Reflection
Imports LaserFront.Trimmer.DllAbout
Imports LaserFront.Trimmer.DllManualTeach
Imports LaserFront.Trimmer.DllPassword
Imports LaserFront.Trimmer.DllProbeTeach
Imports LaserFront.Trimmer.DllSysPrm
Imports LaserFront.Trimmer.DllSystem
Imports LaserFront.Trimmer.DllTeach
Imports LaserFront.Trimmer.DllUtility
Imports LaserFront.Trimmer.DllVideo
Imports TrimClassLibrary
Imports DllPlcIf                        'V2.2.0.0⑤ 
Imports System.Windows.Forms.Control
Imports LaserFront.Trimmer.DefWin32Fnc
Imports System.Runtime.InteropServices      '@@@888

Imports LaserFront.Trimmer
Imports LaserFront.Trimmer.DllSysPrm.SysParam
Imports UsrFunc.FormEdit

Module Globals_define
#Region "グローバル定数/変数の定義"

    '   多重起動防止Mutexハンドル
    Public gmhUserPro As System.Threading.Mutex = New System.Threading.Mutex(False, Application.ProductName)

    '---------------------------------------------------------------------------
    '   アプリケーション名/アプリケーション種別/アプリケーションモード
    '---------------------------------------------------------------------------
    '----- 強制終了用アプリケーション -----
    Public Const APP_FORCEEND As String = "c:\Trim\ForceEndProcess.exe"

    '-------------------------------------------------------------------------------
    '   ファイルパス名
    '-------------------------------------------------------------------------------
    Public Const OCX_PATH As String = "c:\Trim\ocx\"       '----- OCX登録パス
    Public Const DLL_PATH As String = "c:\Trim\"            '----- DLL登録パス
    Public Const SYSPARAMPATH As String = "C:\TRIM\tky.ini"
    Public Const USER_SYSPARAMPATH As String = "C:\TRIM\UserFunc.ini"         'V2.1.0.0②


    'COPYDATASTRUCT構造体
    Public Structure COPYDATASTRUCT
        Public dwData As Int32   '送信する32ビット値
        Public cbData As Int32        'lpDataのバイト数
        Public lpData As String     '送信するデータへのポインタ(0も可能)
    End Structure

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Function SendMessage(
                           ByVal hWnd As IntPtr,
                           ByVal wMsg As Int32,
                           ByVal wParam As Int32,
                           ByVal lParam As Int32) As Integer
    End Function


    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Function SendMessage(
                            ByVal hWnd As IntPtr,
                            ByVal wMsg As Int32,
                            ByVal wParam As Int32,
                            ByRef lParam As COPYDATASTRUCT) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Unicode, EntryPoint:="SendMessage")>
    Public Function SendMessageString(ByVal hWnd As IntPtr,
                                      ByVal wMsg As UInt32,
                                      ByVal wParam As Int32,
                                      <[In], MarshalAs(UnmanagedType.LPWStr)>
                                      lParam As String) As Integer
    End Function

    '-------------------------------------------------------------------------------
    '   システムパラメータ(形式はDllgSysPrm.dllで定義)
    '-------------------------------------------------------------------------------
    Public DllSysPrmSysParam_definst As New DllSysPrm.SysParam
    Public gSysPrm As SYSPARAM_PARAM           ' システムパラメータ

    '-------------------------------------------------------------------------------
    '   オブジェクト定義
    '-------------------------------------------------------------------------------
    '----- Form1クラス -----
    Public ObjMain As Form1                             ' Form1クラス              ###lstLog

    '----- .NETのDLL -----
    Public ObjSys As SystemNET                          ' DllSystem.dll
    Public ObjUtl As Utility                            ' DllUtility.dll
    Public ObjHlp As HelpVersion                        ' DllAbout.dll
    Public ObjPas As Password                           ' DllPassword.dll
    Public ObjMTC As ManualTeach                        ' DllManualTeach.dll
    Public ObjTch As Teaching                           ' DllTeach.dll
    Public ObjPrb As Probe                              ' DllProbeTeach.dll
    Public ObjVdo As VideoLibrary                       ' DllVideo.dll
    'Public ObjPrt As Object                             ' OcxPrint.ocx
    Public ObjMON(32) As Object
    Public gparModules As MainModules                   ' 親側メソッド呼出しオブジェクト(DllSystem用)
    Public TrimClassCommon As New TrimClassLibrary.Common()             ' 共通関数

    '-------------------------------------------------------------------------------
    '   最大値/最小値
    '-------------------------------------------------------------------------------
    Public Const cMAXOptFlgNUM As Short = 5                 ' OcxSystem用ｺﾝﾊﾟｲﾙｵﾌﾟｼｮﾝの数 (最大5個)

    '-------------------------------------------------------------------------------
    '   Lamp ON/OFF制御用ﾗﾝﾌﾟ番号
    '-------------------------------------------------------------------------------
    Public Const LAMP_START As Short = 0
    Public Const LAMP_RESET As Short = 1
    Public Const LAMP_Z As Short = 2
    Public Const LAMP_HALT As Short = 5
    Public Const cSTS_HALTSW_ON As Integer = 4              ' HALTスイッチON Switch状態

    '-------------------------------------------------------------------------------
    ' 拡張Ｉ／Ｏ　ビット定義
    '-------------------------------------------------------------------------------
    Public Const EXT_BIT0 As Integer = &H1
    Public Const EXT_BIT1 As Integer = &H2
    Public Const EXT_BIT2 As Integer = &H4
    Public Const EXT_BIT3 As Integer = &H8
    Public Const EXT_BIT4 As Integer = &H10
    Public Const EXT_BIT5 As Integer = &H20
    Public Const EXT_BIT6 As Integer = &H40
    Public Const EXT_BIT7 As Integer = &H80
    Public Const EXT_BIT8 As Integer = &H100
    Public Const EXT_BIT9 As Integer = &H200
    Public Const EXT_BIT10 As Integer = &H400
    Public Const EXT_BIT11 As Integer = &H800
    Public Const EXT_BIT12 As Integer = &H1000
    Public Const EXT_BIT13 As Integer = &H2000
    Public Const EXT_BIT14 As Integer = &H4000
    Public Const EXT_BIT15 As Integer = &H8000

    '-------------------------------------------------------------------------------
    ' 拡張I/O汎用定義
    '-------------------------------------------------------------------------------
    Public Const EXT_IN0 As UShort = &H10                   ' B04: ユーザ割付１
    Public Const EXT_IN1 As UShort = &H20                   ' B05: ユーザ割付２
    Public Const EXT_IN2 As UShort = &H40                   ' B06: ユーザ割付３
    Public Const EXT_IN3 As UShort = &H80                   ' B07: ユーザ割付４

    '-------------------------------------------------------------------------------
    '   Interlock switch bits (ADR. 0x21E8)
    '-------------------------------------------------------------------------------
    Public Const BIT_SLIDECOVER_CLOSE As Short = &H100S     ' B8 : ｽﾗｲﾄﾞｶﾊﾞｰ閉
    Public Const BIT_SLIDECOVER_OPEN As Short = &H200S      ' B9 : ｽﾗｲﾄﾞｶﾊﾞｰ開
    Public Const BIT_EMERGENCY_SW As Short = &H400S         ' B10: ｴﾏｰｼﾞｪﾝｼｰSW
    Public Const BIT_EMERGENCY_RESET As Short = &H800S      ' B11: ｴﾏｰｼﾞｪﾝｼｰﾘｾｯﾄ
    Public Const BIT_INTERLOCK_DISABLE As Short = &H1000S   ' B12: ｲﾝﾀｰﾛｯｸ解除SW
    Public Const BIT_SERVO_ALARM As Short = &H2000S         ' B13: ｻｰﾎﾞｱﾗｰﾑ
    Public Const BIT_COVER_CLOSE As Short = &H4000S         ' B14: ｶﾊﾞｰ閉
    '                                                       ' B15: ｶﾊﾞｰ&ｽﾗｲﾄﾞｶﾊﾞｰ閉

    Public Const INTERLOCK_STS_DISABLE_NO As Short = (0)    ' インターロック状態(解除なし)
    Public Const INTERLOCK_STS_DISABLE_PART As Short = (1)  ' インターロック一部解除（ステージ可動可能）
    Public Const INTERLOCK_STS_DISABLE_FULL As Short = 2    ' インターロック全解除
    Public Const SLIDECOVER_OPEN As Short = (1)             ' Bit0 : スライドカバー：オープン
    Public Const SLIDECOVER_CLOSE As Short = (2)            ' Bit1 : スライドカバー：クローズ
    Public Const SLIDECOVER_MOVING As Short = (4)           ' Bit2 : スライドカバー：動作中
    '----- シグナルタワー制御種別 -----  
    Public Const SIGTOWR_NORMAL As Short = 0                ' 標準３色制御
    Public Const SIGTOWR_SPCIAL As Short = 1                ' ４色制御(特注)

    '-------------------------------------------------------------------------------
    '   シグナルタワー３色制御(標準)SL432R/SL436R共通
    '   ①手動運転中 ･････････････ 無点灯(原点復帰完了, レディ(手動))
    '   ②インターロック解除中････ 黄点滅(H/Wで制御)
    '   ③ティーチング中･･････････ 黄点灯
    '   ④原点復帰時 ･････････････ 緑点滅
    '   ⑤非常停止中 ･････････････ 赤点灯＋ブザーＯＮ ← H/Wが落ちる為なし
    '   ⑥自動運転中 ･････････････
    '     ａ)正常運転時　　　：緑点灯
    '     ｂ)全マガジン終了時：赤点滅＋ブザーＯＮ
    '     ｃ)アラーム時　　　：赤点滅＋ブザーＯＮ（但し、⑤項優先）
    '-------------------------------------------------------------------------------
    '----- OUTPUT -----                                     ' ON時の意味
    '                                                       ' B0 : 未使用
    '                                                       ' :
    '                                                       ' B7 : 未使用
    Public Const SIGOUT_GRN_ON As UShort = &H100            ' B8 : 緑点灯  (自動運転中)
    Public Const SIGOUT_YLW_ON As UShort = &H200            ' B9 : 黄点灯  (ティーチング中)
    Public Const SIGOUT_RED_ON As UShort = &H400            ' B10: 赤点灯  (非常停止) ※未使用(H/Wで制御)
    Public Const SIGOUT_GRN_BLK As UShort = &H800           ' B11: 緑点滅  (原点復帰時)
    Public Const SIGOUT_YLW_BLK As UShort = &H1000          ' B12: 黄点滅  (インターロック解除中)
    Public Const SIGOUT_RED_BLK As UShort = &H2000          ' B13: 赤点滅  (異常/全マガジン終了) ※+ブザー１
    Public Const SIGOUT_BZ1_ON As UShort = &H4000           ' B14: ブザー１(異常) ※+赤点滅
    '                                                       ' B15: 未使用

    '-------------------------------------------------------------------------------
    '   拡張ＥＸＴＢＩＴ(上位16ビット ADR. 213A)
    '   ※シグナルタワー４色制御(特注)
    '-------------------------------------------------------------------------------
    '----- OUTPUT -----                                     ' ON時の意味
    '                                                       ' B0 (B16): 未使用
    '                                                       ' :
    '                                                       ' B3 (B19): 未使用
    '                                                       ' B4 (B20): 未使用
    '                                                       ' :
    '                                                       ' B7 (B23): 未使用

    Public Const EXTOUT_RED_ON As UShort = &H100            ' B8 (B24): 赤点灯  (非常停止) ※未使用(H/Wで制御)
    Public Const EXTOUT_RED_BLK As UShort = &H200           ' B9 (B25): 赤点滅  (異常) ※+ブザー１
    Public Const EXTOUT_YLW_ON As UShort = &H400            ' B10(B26): 黄色点灯(原点復帰完了, レディ(手動))
    Public Const EXTOUT_YLW_BLK As UShort = &H800           ' B11(B27): 黄色点滅(原点復帰中)

    Public Const EXTOUT_GRN_ON As UShort = &H1000           ' B12(B28): 緑点灯  (自動運転中)
    Public Const EXTOUT_GRN_BLK As UShort = &H2000          ' B13(B29): 緑点滅  (-) ※未使用
    Public Const EXTOUT_BZ1_ON As UShort = &H4000           ' B14(B30): ブザー１(異常) ※+赤点滅
    '                                                       ' B15(B31): 未使用

    '-------------------------------------------------------------------------------
    '   ローダーＩ／Ｏビット(ADR. 219A)
    '-------------------------------------------------------------------------------
    '----- ﾄﾘﾏｰ  → ﾛｰﾀﾞｰ -----
    ' ※Bit0～Bit4,Bit7が標準版
    Public Const COM_STS_TRM_STATE As Short = &H1S          ' B0 : トリマ停止(0:停止,1:動作中)
    Public Const COM_STS_TRM_NG As Short = &H2S             ' B1 : トリミングＮＧ(0:正常, 1:NG)
    Public Const COM_STS_PTN_NG As Short = &H4S             ' B2 : パターン認識エラー(0:正常, 1:エラー)
    Public Const COM_STS_TRM_ERR As Short = &H8S            ' B3 : トリマエラー(0:正常, 1:エラー)
    Public Const COM_STS_TRM_READY As Short = &H10S         ' B4 : トリマレディ(0:ﾉｯﾄﾚﾃﾞｨ, 1:ﾚﾃﾞｨ)
    Public Const COM_STS_LOT_END As Short = &H20S           ' B5 : ロット終了(0:処理中, 1:終了状態)　'V1.2.0.0④　
    Public Const COM_STS_ABS_ON As Short = &H40S            ' B6 : 吸着(0:オン, 1:オフ)　　　　　　　'V1.2.0.0④
    Public Const COM_STS_CLAMP_ON As Short = &H80S          ' B7 : 載物台ｸﾗﾝﾌﾟ開閉(0:閉, 1:開)　　　 'V1.2.0.0④
    'V1.2.0.0④    '                                                       ' B5 : 未使用
    'V1.2.0.0④    '                                                       ' B6 : 未使用
    'V1.2.0.0④    Public Const COM_STS_ABS_ON As Short = &H80S            ' B7 : 載物台ｸﾗﾝﾌﾟ開閉(0:閉, 1:開)

    '----- ﾛｰﾀﾞｰ  → ﾄﾘﾏｰ -----
    ' ※Bit0～Bit3までが標準版
    Public Const cHSTcRDY As Short = 1                      ' B0 : ｵｰﾄﾛｰﾀﾞｰ有無(0:無, 1:有)
    Public Const cHSTcAUTO As Short = 2                     ' B1 : ｵｰﾄﾛｰﾀﾞｰﾓｰﾄﾞ(1=自動ﾓｰﾄﾞ, 0=手動ﾓｰﾄﾞ)
    Public Const cHSTcSTATE As Short = 4                    ' B2 : ｵｰﾄﾛｰﾀﾞｰ動作中(0=動作中, 1=停止)
    Public Const cHSTcTRMCMD As Short = 8                   ' B3 : ﾄﾘﾏｰｽﾀｰﾄ(1=ﾄﾘﾏｰｽﾀｰﾄ) ※ﾗｯﾁ
    Public Const cHSTcLOTCHANGE As Short = &H10             ' B4 : ロット切り替え信号
    Public Const cHSTcABS_ON As Short = &H40S               ' B6 : 吸着(0:オン, 1:オフ)　　　　　　　'V1.2.0.0④
    Public Const cHSTcCLAMP_ON As Short = &H80S             ' B7 : 載物台ｸﾗﾝﾌﾟ開閉(0:閉, 1:開)　　　 'V1.2.0.0④

    Public gdwATLDDATA As UInteger                          ' ローダ出力データ
    Public gDebugHostCmd As UInteger                        ' ローダ入力データ(ﾃﾞﾊﾞｯｸﾞ用)
    Public gwPrevHcmd As UInteger                           ' ローダ入力データ退避域
    Public gbClampOpen As Boolean = True                    'V1.2.0.0④ クランプ開可能状態:True 既に開:False
    Public gbVaccumeOff As Boolean = True                   'V1.2.0.0④ 吸着オフ可能状態:True 既にオフ:False

    '-------------------------------------------------------------------------------
    '   gMode(frmResetの処理モード) ※100～266は戻り値にも使用します(-101～-266)
    '-------------------------------------------------------------------------------
    Public gFRsetFlg As Short                               ' frmResetﾌﾗｸﾞ(0:初期値, 1:frmReset処理中)
    Public gMode As Short                                   ' 処理モード退避域

    Public Const cGMODE_ORG As Short = 0                    '  0 : 原点復帰
    Public Const cGMODE_ORG_MOVE As Short = 1               '  1 : 原点位置移動
    Public Const cGMODE_START_RESET As Short = 2            '  2 : 操作確認画面(START/RESET待ち)
    '                                                       '  3 :
    '                                                       '  4 :
    Public Const cGMODE_EMG As Short = 5                    '  5 : 非常停止メッセージ表示
    '                                                       '  6 :
    Public Const cGMODE_SCVR_OPN As Short = 7               '  7 : トリミング中のスライドカバー開メッセージ表示
    Public Const cGMODE_CVR_OPN As Short = 8                '  8 : トリミング中の筐体カバー開メッセージ表示
    Public Const cGMODE_SCVRMSG As Short = 9                '  9 : スライドカバー開メッセージ表示(トリミング中以外)
    Public Const cGMODE_CVRMSG As Short = 10                ' 10 : 筐体カバー開確認メッセージ表示(トリミング中以外)
    Public Const cGMODE_ERR_HW As Short = 11                ' 11 : ハードウェアエラー(カバーが閉じてます)メッセージ表示
    Public Const cGMODE_ERR_HW2 As Short = 12               ' 12 : ハードウェアエラーメッセージ表示
    Public Const cGMODE_CVR_LATCH As Short = 13             ' 13 : カバー開ラッチメッセージ表示
    Public Const cGMODE_CVR_CLOSEWAIT As Short = 14         ' 14 : 筐体カバークローズもしくはインターロック解除待ち

    Public Const cGMODE_ERR_DUST As Short = 20              ' 20 : 集塵機異常検出メッセージ表示
    Public Const cGMODE_ERR_AIR As Short = 21               ' 21 : エアー圧エラー検出メッセージ表示

    Public Const cGMODE_ERR_HING As Short = 40              ' 40 : 連続HI-NGｴﾗｰ(ADVｷｰ押下待ち)
    Public Const cGMODE_SWAP As Short = 41                  ' 41 : 基板交換(STARTｷｰ押下待ち)
    Public Const cGMODE_XYMOVE As Short = 42                ' 42 : 終了時のﾃｰﾌﾞﾙ移動確認(STARTｷｰ押下待ち)
    Public Const cGMODE_LDR_ALARM As Short = 44             ' 44 : ローダアラーム発生                  'V2.2.0.0⑤ 
    Public Const cGMODE_LDR_START As Short = 45             ' 45 : 自動運転開始(STARTｷｰ押下待ち)       'V2.2.0.0⑤ 
    Public Const cGMODE_LDR_TMOUT As Short = 46             ' 46 : ローダ通信タイムアウト              'V2.2.0.0⑤ 
    Public Const cGMODE_LDR_END As Short = 47               ' 47 : 自動運転終了(STARTｷｰ押下待ち)       'V2.2.0.0⑤   

    Public Const cGMODE_AUTO_LASER As Short = 50            ' 50 : 自動レーザパワー調整

    Public Const cGMODE_LDR_CHK As Short = 60               ' 60 : ローダ状態チェック(起動時ﾛｰﾀﾞ自動ﾓｰﾄﾞ/動作中)
    Public Const cGMODE_LDR_ERR As Short = 61               ' 61 : ローダ状態エラー(ﾛｰﾀﾞ自動でﾛｰﾀﾞ無)
    Public Const cGMODE_LDR_MNL As Short = 62               ' 62 : カバー開後のローダ手動モード処理
    Public Const cGMODE_LDR_WKREMOVE As Short = 63          ' 63 : 残基板取り除きメッセージ     'V2.2.0.0⑤
    Public Const cGMODE_LDR_RSTAUTO As Short = 64           ' 64 : 自動運転中止メッセージ      'V2.2.0.0⑤
    Public Const cGMODE_LDR_WKREMOVE2 As Short = 65         ' 65 : 残基板取り除きメッセージ(APP終了)  'V2.2.0.0⑤

    Public Const cGMODE_LDR_CHK_AUTO As Short = 67          ' 63 : ローダ状態チェック(自動運転時),ローダが自動に切り替わるまで待つ'V1.0.4.3⑫
    Public Const cGMODE_LDR_STAGE_ORG As Short = 66         ' 66 : ステージ原点移動v

    Public Const cGMODE_OPT_START As Short = 70             ' 70 : ﾄﾘﾐﾝｸﾞ開始時のｽﾀｰﾄSW押下待ち
    Public Const cGMODE_OPT_END As Short = 71               ' 71 : ﾄﾘﾐﾝｸﾞ終了時のｽﾗｲﾄﾞｶﾊﾞｰ開待ち

    Public Const cGMODE_MSG_DSP As Short = 90               ' 90 : 指定メッセージ表示(ADVｷｰ押下待ち)

    ' リミットセンサー& 軸エラー & タイムアウトメッセージ
    Public Const cGMODE_TO_AXISX As Short = 101             ' 101: X軸エラー(タイムアウト)
    Public Const cGMODE_TO_AXISY As Short = 102             ' 102: Y軸エラー(タイムアウト)
    Public Const cGMODE_TO_AXISZ As Short = 103             ' 103: Z軸エラー(タイムアウト)
    Public Const cGMODE_TO_AXIST As Short = 104             ' 104: θ軸エラー(タイムアウト)
    '【ソフトリミットエラー】
    Public Const cGMODE_SL_AXISX As Short = 105             ' 105: X軸ソフトリミットエラー
    Public Const cGMODE_SL_AXISY As Short = 106             ' 106: Y軸ソフトリミットエラー
    Public Const cGMODE_SL_AXISZ As Short = 107             ' 107: Z軸ソフトリミットエラー
    Public Const cGMODE_SL_BPX As Short = 110               ' 110: BP X軸ソフトリミットエラー
    Public Const cGMODE_SL_BPY As Short = 111               ' 111: BP Y軸ソフトリミットエラー

    Public Const cGMODE_TO_ROTATT As Short = 108            ' 108: ロータリアッテネータエラー(タイムアウト)
    Public Const cGMODE_TO_AXISZ2 As Short = 109            ' 109: Z2軸エラー(タイムアウト)

    Public Const cGMODE_SRV_ARM As Short = 202              ' 202: サーボアラーム
    Public Const cGMODE_AXISX_LIM As Short = 203            ' 203: X軸リミット
    Public Const cGMODE_AXISY_LIM As Short = 204            ' 204: Y軸リミット
    Public Const cGMODE_AXISZ_LIM As Short = 205            ' 205: Z軸リミット
    Public Const cGMODE_AXIST_LIM As Short = 206            ' 206: θ軸リミット
    Public Const cGMODE_RATT_LIM As Short = 207             ' 207: ロータリーアッテネータリミット
    Public Const cGMODE_AXISZ2_LIM As Short = 208           ' 208: Z2軸リミット

    Public Const cGMODE_BASE_ERR As Short = 200             ' 軸エラーベース番号
    '【X軸エラー】
    Public Const cGMODE_AXISX_AOFF As Short = 211           ' 211: X軸エラー(Bit All Off)
    Public Const cGMODE_AXISX_AON As Short = 212            ' 212: X軸エラー(Bit All On)
    Public Const cGMODE_AXISX_ARM As Short = 213            ' 213: X軸アラーム
    Public Const cGMODE_AXISX_PML As Short = 214            ' 214: ±X軸リミット
    Public Const cGMODE_AXISX_PLM As Short = 215            ' 215: +X軸リミット
    Public Const cGMODE_AXISX_MLM As Short = 216            ' 216: -X軸リミット
    '【Y軸エラー】
    Public Const cGMODE_AXISY_AOFF As Short = 221           ' 221: Y軸エラー(Bit All Off)
    Public Const cGMODE_AXISY_AON As Short = 222            ' 222: Y軸エラー(Bit All On)
    Public Const cGMODE_AXISY_ARM As Short = 223            ' 223: Y軸アラーム
    Public Const cGMODE_AXISY_PML As Short = 224            ' 224: ±Y軸リミット
    Public Const cGMODE_AXISY_PLM As Short = 225            ' 225: +Y軸リミット
    Public Const cGMODE_AXISY_MLM As Short = 226            ' 226: -Y軸リミット
    '【Z軸エラー】
    Public Const cGMODE_AXISZ_AOFF As Short = 231           ' 231: Z軸エラー(Bit All Off)
    Public Const cGMODE_AXISZ_AON As Short = 232            ' 232: Z軸エラー(Bit All On)
    Public Const cGMODE_AXISZ_ARM As Short = 233            ' 233: Z軸アラーム
    Public Const cGMODE_AXISZ_PML As Short = 234            ' 234: ±Z軸リミット
    Public Const cGMODE_AXISZ_PLM As Short = 235            ' 235: +Z軸リミット
    Public Const cGMODE_AXISZ_MLM As Short = 236            ' 236: -Z軸リミット
    Public Const cGMODE_AXISZ_ORG As Short = 237            ' 237: Z軸原点復帰未完了
    '【θ軸エラー】
    Public Const cGMODE_AXIST_AOFF As Short = 241           ' 241: θ軸エラー(Bit All Off)
    Public Const cGMODE_AXIST_AON As Short = 242            ' 242: θ軸エラー(Bit All On)
    Public Const cGMODE_AXIST_ARM As Short = 243            ' 243: θ軸アラーム
    Public Const cGMODE_AXIST_PML As Short = 244            ' 244: ±θ軸リミット
    Public Const cGMODE_AXIST_PLM As Short = 245            ' 245: +θ軸リミット
    Public Const cGMODE_AXIST_MLM As Short = 246            ' 246: -θ軸リミット
    '【Z2軸エラー】
    Public Const cGMODE_AXISZ2_AOFF As Short = 251          ' 251: Z2軸エラー(Bit All Off)
    Public Const cGMODE_AXISZ2_AON As Short = 252           ' 252: Z2軸エラー(Bit All On)
    Public Const cGMODE_AXISZ2_ARM As Short = 253           ' 253: Z2軸アラーム
    Public Const cGMODE_AXISZ2_PML As Short = 254           ' 254: ±Z2軸リミット
    Public Const cGMODE_AXISZ2_PLM As Short = 255           ' 255: +Z2軸リミット
    Public Const cGMODE_AXISZ2_MLM As Short = 256           ' 256: -Z2軸リミット
    Public Const cGMODE_AXISZ2_ORG As Short = 257           ' 257: Z2軸原点復帰未完了
    '【ﾛｰﾀﾘｱｯﾃﾈｰﾀｰｴﾗｰ】
    Public Const cGMODE_ROTATT_AOFF As Short = 261          ' 261: ﾛｰﾀﾘｱｯﾃﾈｰﾀｰｴﾗｰ(Bit All Off)
    Public Const cGMODE_ROTATT_AON As Short = 262           ' 262: ﾛｰﾀﾘｱｯﾃﾈｰﾀｰｴﾗｰ(Bit All On)
    Public Const cGMODE_ROTATT_ARM As Short = 263           ' 263: ﾛｰﾀﾘｱｯﾃﾈｰﾀｰｱﾗｰﾑ
    Public Const cGMODE_ROTATT_PML As Short = 264           ' 264: ±ﾛｰﾀﾘｱｯﾃﾈｰﾀｰﾘﾐｯﾄ
    Public Const cGMODE_ROTATT_PLM As Short = 265           ' 265: +ﾛｰﾀﾘｱｯﾃﾈｰﾀｰﾘﾐｯﾄ
    Public Const cGMODE_ROTATT_MLM As Short = 266           ' 266: -ﾛｰﾀﾘｱｯﾃﾈｰﾀｰﾘﾐｯﾄ

    '-------------------------------------------------------------------------------
    '   操作ログメッセージ
    '-------------------------------------------------------------------------------
    '----- 操作ログメッセージ -----
    Public MSG_OPLOG_START As String                        ' "ユーザプログラム起動"
    Public MSG_OPLOG_FUNC01 As String                       ' "データロード"
    Public MSG_OPLOG_FUNC02 As String                       ' "データセーブ"
    Public MSG_OPLOG_FUNC03 As String                       ' "データ編集"
    Public MSG_OPLOG_FUNC04 As String                       ' "マスタチェック"(特注)
    Public MSG_OPLOG_FUNC05 As String                       ' "レーザ調整"
    Public MSG_OPLOG_FUNC06 As String                       ' "ロット切替"
    Public MSG_OPLOG_FUNC07 As String                       ' "プローブ位置合わせ"
    Public MSG_OPLOG_FUNC07_2 As String                     ' "プローブ位置合わせ2"
    Public MSG_OPLOG_FUNC08 As String                       ' "ティーチング"
    Public MSG_OPLOG_FUNC08S As String                      ' "カット補正位置ティーチング"
    Public MSG_OPLOG_FUNC09 As String                       ' "パターン登録"
    Public MSG_OPLOG_FUNC10 As String                       ' "プローブ位置合わせ２"
    Public MSG_OPLOG_FUNC11 As String                       ' "データ設定"
    Public MSG_OPLOG_END As String                          ' "ユーザプログラム終了"
    Public MSG_OPLOG_TRIMST As String                       ' "トリミング"
    Public MSG_OPLOG_LOTCHG As String                       ' "ロット切替信号受信"
    Public MSG_OPLOG_STOP As String                         ' "トリマ装置停止"
    Public MSG_OPLOG_LOTSET As String                       ' "ロット情報データ設定"

    '----- メッセージ -----
    Public MSG_DataNotLoad As String                        ' データ未ロード
    Public MSG_SPRASH31 As String
    Public MSG_SPRASH32 As String
    Public MSG_SPRASH52 As String
    Public MSG_105 As String
    Public MSG_136 As String
    Public MSG_137 As String
    Public MSG_138 As String
    Public MSG_139 As String
    Public MSG_140 As String
    Public MSG_141 As String
    Public MSG_142 As String
    Public MSG_143 As String
    Public MSG_144 As String
    Public MSG_145 As String
    Public MSG_146 As String
    Public MSG_147 As String
    Public MSG_148 As String
    Public MSG_149 As String
    Public MSG_150 As String
    Public MSG_151 As String
    Public MSG_152 As String
    Public MSG_153 As String

    ' ＴＸ，ＴＹ関係　START
    ' ■frmMsgBox 画面終了確認
    Public MSG_CLOSE_LABEL01 As String
    Public MSG_CLOSE_LABEL02 As String
    Public MSG_CLOSE_LABEL03 As String
    Public MSG_EXECUTE_TXTYLABEL As String 'TX,TY
    Public TITLE_TX As String 'チップサイズ(TX)ティーチング
    Public TITLE_TY As String 'ステップサイズ(TY)ティーチング
    Public LBL_TXTY_TEACH_03 As String '補正量
    Public LBL_TXTY_TEACH_04 As String '補正比率
    Public LBL_TXTY_TEACH_05 As String 'ﾁｯﾌﾟｻｲｽﾞ (mm)
    Public LBL_TXTY_TEACH_07 As String '補正前
    Public LBL_TXTY_TEACH_08 As String '補正後
    Public LBL_TXTY_TEACH_09 As String 'ｸﾞﾙｰﾌﾟｲﾝﾀｰﾊﾞﾙ(mm)
    Public LBL_TXTY_TEACH_11 As String 'ステップインターバル(mm)(追加)
    Public LBL_TXTY_TEACH_12 As String '第１基準点
    Public LBL_TXTY_TEACH_13 As String '第２基準点
    Public LBL_TXTY_TEACH_14 As String 'グループ
    Public LBL_CMD_CANCEL As String
    Public CMD_CANCEL As String 'キャンセル
    Public INFO_MSG13 As String '"チップサイズ　ティーチング"
    Public INFO_MSG14 As String '"ステップ間インターバル　ティーチング"→"ステージグループ間隔ティーチング"
    Public INFO_MSG15 As String '"ステップオフセット量　ティーチング"
    Public INFO_MSG16 As String '"基準位置を合わせて下さい。"
    Public INFO_MSG17 As String '"移動:[矢印]  決定:[START]  中断:[RESET]" & vbCrLf & "[HALT]で１つ前の処理に戻ります。"
    Public INFO_MSG18 As String '"第1グループ、第1抵抗基準位置のティーチング"
    Public INFO_MSG19 As String '"第"
    Public INFO_MSG20 As String '"グループ、最終抵抗基準位置のティーチング"
    Public INFO_MSG23 As String '"グループ間インターバル　ティーチング"→"ＢＰグループ間隔ティーチング"
    Public INFO_MSG28 As String '"グループ、最終端位置のティーチング"
    Public INFO_MSG29 As String '"グループ、最先端位置のティーチング"
    Public INFO_MSG30 As String '"サーキット間隔ティーチング"
    Public INFO_MSG31 As String '"ステップオフセット量のティーチング"
    Public INFO_MSG32 As String '"(ＴＸ)"   '###084
    Public INFO_MSG33 As String '"(ＴＹ)"   '###084
    Public INFO_MSG34 As String '"ステップサイズ　ティーチング"
    ' ＴＸ，ＴＹ関係　END

    '----- 画像表示プログラムの表示位置 -----
    'Public Const FORM_X As Integer = 4                                  ' コントロール上部左端座標X
    'Public Const FORM_Y As Integer = 20                                 ' コントロール上部左端座標Y
    Public Const FORM_X As Integer = 0                                  ' コントロール上部左端座標X
    Public Const FORM_Y As Integer = 0                                  ' コントロール上部左端座標Y

    '----- 画像表示プログラムの起動用 -----
    Public Const DISPGAZOU_PATH As String = "C:\TRIM\DispGazouSmall.exe"    ' 画像表示プログラム名
    Public Const DISPGAZOU_WRK As String = "C:\\TRIM"                       ' 作業フォルダ

    '----- 系名 -----
    Public Const MACHINE_TYPE_SL432 As String = "SL432R"
    Public Const MACHINE_TYPE_SL436 As String = "SL436R"

    '----- ログ画面表示用 -----　
    Public gDspClsCount As Integer = 5                                  ' ログ画面表示クリア基板枚数
    Public gDspCounter As Integer = 0                                   ' ログ画面表示基板枚数カウンタ

    Public gPlateCount As Integer = 0                                   ' 基板枚数(デバッグ用)

    '----- GPIB 制御用 -----                               
    Public gGpibMultiMeterCount As Integer = 5                          ' 外部測定器でITと測定での測定回数（測定値が安定しないので最後の測定値を使用する。）

    '----- EXTOUT LED制御ビット -----                               
    Public glLedBit As Long                                             ' LED制御ビット(EXTOUT) 
    Public Const INITIAL_TEST As Integer = 0                ' 初期テスト
    Public Const FINAL_TEST As Integer = 1                  ' 最終テスト

    Public Const SETAXISSPDY_DEFALT As UInteger = 15000    ' Ｙ軸ステージ速度の変更機能初期値 'V2.0.0.0⑮

    Public giStageYDir As Integer = 1                       ' ステージYの移動方向(CW(1), CCW(-1))    'V2.2.0.0① 

    '---------------------------------------------------------------------------
    ' トリミング動作モード
    '---------------------------------------------------------------------------
    '^^^^^ ディジタルSW HI　定義 -----
    Public Const DGSW_HI_NODISP As Integer = 0              ' 表示なし
    Public Const DGSW_HI_NGDISP As Integer = 1              ' ＮＧのみ表示
    Public Const DGSW_HI_DISP As Integer = 2                ' 全て表示

    '----- ディジタルSW LOW　定義 -----
    Public Const TRIM_MODE_ITTRFT As Integer = 0            ' イニシャルテスト＋トリミング＋ファイナルテスト実行
    Public Const TRIM_MODE_MEAS As Integer = 1              ' 測定実行
    Public Const TRIM_MODE_CUT As Integer = 2               ' カット実行
    Public Const TRIM_MODE_STPRPT As Integer = 3            ' ステップ＆リピート実行
    Public Const TRIM_MODE_MEAS_MARK As Integer = 4         ' 'V1.0.4.3⑩測定マーキングモード・ファイナル測定のみ
    Public Const TRIM_MODE_POWER As Integer = 5             ' 電源モード 'V2.0.0.0②
    Public Const TRIM_VARIATION_MEAS As Integer = 6         ' 測定値変動測定 'V2.0.0.0②

    'Public Const TRIM_MODE_TRFT As Integer = 1              ' トリミング＋ファイナルテスト実行
    'Public Const TRIM_MODE_FT As Integer = 2                ' ファイナルテスト実行（判定）
    'Public Const TRIM_MODE_MEAS As Integer = 3              ' 測定実行
    'Public Const TRIM_MODE_POSCHK As Integer = 4            ' ポジションチェック
    'Public Const TRIM_MODE_CUT As Integer = 5               ' カット実行
    'Public Const TRIM_MODE_STPRPT As Integer = 6            ' ステップ＆リピート実行

    '---------------------------------------------------------------------------
    ' SLIDE COVER+XY移動同時動作
    '---------------------------------------------------------------------------
    Public Const TYPE_OFFLINE As Short = 0                  ' OFFLINE
    Public Const TYPE_ONLINE As Short = 1                   ' ONLINE
    Public Const TYPE_MANUAL As Short = 2                   ' SLIDE COVER+XY移動同時動作

    '---------------------------------------------------------------------------
    ' カット動作モード （0:トリミング、1:ティーチング、2:強制カット）
    '---------------------------------------------------------------------------
    Public Const TRIM_MODE As Integer = 0                   ' ストレート　トリミングモード
    Public Const TEACH_MODE As Integer = 1                  ' ストレート　ティーチングモード
    Public Const FORCE_MODE As Integer = 2                  ' ストレート　強制カットモード


    Public Const CUT_MODE_NORMAL As Integer = 0             ' ノーマル
    Public Const CUT_MODE_RETURN As Integer = 1             ' リターンカット
    Public Const CUT_MODE_RETRACE As Integer = 2            ' リトレースカット
    Public Const CUT_MODE_NANAME As Integer = 4             ' 斜めカット

    '---------------------------------------------------------------------------
    ' カット方法 （1:ﾄﾗｯｷﾝｸﾞ、2:INDEX、3:NG ) 
    '---------------------------------------------------------------------------
    Public Const CNS_CUTM_TR As Integer = 1                 ' トラッキング
    Public Const CNS_CUTM_IX As Integer = 2                 ' インデックス
    Public Const CNS_CUTM_NG As Integer = 3                 ' ＮＧ
    Public Const CNS_CUTM_NON_POS_IX As Integer = 4         ' ポジショニング無しインデックス

    '---------------------------------------------------------------------------
    ' カット形状 （1:STカット、2:Lカット、3:SPカット 4:IXカット）
    '---------------------------------------------------------------------------
    'Public Const CNS_CUTP_ST As Integer = 1                 ' STカット
    'Public Const CNS_CUTP_L As Integer = 2                  ' Lカット
    'Public Const CNS_CUTP_SP As Integer = 3                 ' SPカット
    'Public Const CNS_CUTP_IX As Integer = 4                 ' IXカット
    'Public Const CNS_CUTP_M As Integer = 19                 ' 文字マーキング　###1042① 
    Public Const CNS_CUTP_NORMAL As Integer = 0             ' ノーマルカット・カットモード指定用 'V1.0.4.3⑧
    Public Const CNS_CUTP_ST As Integer = 1                 ' STカット
    Public Const CNS_CUTP_ST_TR As Integer = 2              ' V1.1.0.0③ストレート・リトレース(RETRACE)カット 
    Public Const CNS_CUTP_L As Integer = 3                  ' V1.1.0.0③Lカット
    Public Const CNS_CUTP_M As Integer = 4                  ' V1.1.0.0③文字マーキング　###1042① 
    Public Const CNS_CUTP_U As Integer = 5                  ' Uカット追加　 'V2.2.0.0②
    Public Const CNS_CUTP_SP As Integer = 6                 ' SPカット V1.1.0.0③番号変更   'V2.2.0.0②5->6
    Public Const CNS_CUTP_IX As Integer = 7                 ' IXカット V1.1.0.0③番号変更   'V2.2.0.0②6->7

    '---------------------------------------------------------------------------
    ' 測定モード （0:抵抗測定、1:電圧測定、2:外部測定（ＧＰＩＢ））
    '---------------------------------------------------------------------------
    Public Const MEAS_MODE_RESISTOR As Integer = 0          ' 抵抗測定
    Public Const MEAS_MODE_VOLTAGE As Integer = 1           ' 電圧測定
    Public Const MEAS_MODE_EXTERNAL As Integer = 2          ' 外部測定（ＧＰＩＢ）

    '---------------------------------------------------------------------------
    ' 測定精度 （0:高速測定、1:高精度測定）
    '---------------------------------------------------------------------------
    Public Const MEAS_TYP_FAST As Integer = 0               ' 高速測定
    Public Const MEAS_TYP_HIPRECI As Integer = 1            ' 高精度測定


    '---------------------------------------------------------------------------
    ' 測定精度 （0:高速測定、1:高精度測定）
    '---------------------------------------------------------------------------
    Public Const MEAS_RNGSET_AUTO As Integer = 0            ' オートレンジ設定
    Public Const MEAS_RNGSET_FIX_TAR As Integer = 1         ' 固定レンジ設定-目標値設定
    Public Const MEAS_RNGSET_FIX_NO As Integer = 2          ' 固定レンジ設定-レンジ番号設定

    'V1.0.4.3⑥↓
    '---------------------------------------------------------------------------
    ' パターン認識(0:無し, 1:有り, 2:手動, 3:自動ＮＧ判定あり）
    '---------------------------------------------------------------------------
    Public Const CUT_PATTERN_NONE As Integer = 0            ' 0:無し
    Public Const CUT_PATTERN_AUTO As Integer = 1            ' 1:有り
    Public Const CUT_PATTERN_MANUAL As Integer = 2          ' 2:手動
    Public Const CUT_PATTERN_AUTO_NG As Integer = 3         ' 3:自動ＮＧ判定あり

    '-------------------------------------------------------------------------------
    '   スロープ定義定義    
    '-------------------------------------------------------------------------------
    Public Const SLP_VTRIMPLS As Integer = 1                ' ＋電圧トリミング
    Public Const SLP_VTRIMMNS As Integer = 2                ' －電圧トリミング
    Public Const SLP_RTRM As Integer = 4                    ' 　抵抗トリミング
    Public Const SLP_VMES As Integer = 5                    ' 　電圧測定
    Public Const SLP_RMES As Integer = 6                    ' 　抵抗測定
    Public Const SLP_NG_MARK As Integer = 7                 ' 　ＮＧマーク
    Public Const SLP_OK_MARK As Integer = 8                 ' 　ＯＫマーク
    Public Const SLP_ATRIMPLS As Integer = 9                ' ＋電流トリミング
    Public Const SLP_ATRIMMNS As Integer = 10               ' －電流トリミング
    Public Const SLP_AMES As Integer = 11                   ' 　電流測定
    Public Const SLP_MARK As Integer = 12                   ' 　マーク印字 'V2.2.1.7①
    'V1.0.4.3⑥↑
    'V1.0.4.3⑦↓
    Public Const DEF_DIR_CW As Integer = 1                  ' 時計方向の回転（Clock Wise)
    Public Const DEF_DIR_CCW As Integer = 2                 ' 半時計方向の回転（Counter Clock Wise)
    'V1.0.4.3⑦↑

    'V1.0.4.3⑦↓
    ' パワー調整(FL用)
    Public Const CUT_CND_L1 As Integer = 1              ' L1加工条件設定
    Public Const CUT_CND_L2 As Integer = 2              ' L2加工条件設定
    Public Const CUT_CND_L3 As Integer = 3              ' L3加工条件設定
    Public Const CUT_CND_L4 As Integer = 4             ' L4加工条件設定
    'V1.0.4.3⑦↑

    '-------------------------------------------------------------------------------
    '   ファイバーレーザ用定義
    '-------------------------------------------------------------------------------
    '----- 発振器種別 -----
    Public Const OSCILLATOR_FL As Integer = 3               ' FL(ﾌｧｲﾊﾞｰﾚｰｻﾞ)
    Public Const OSCILLATOR_SP As Integer = 5               ' SPレーザ

#If cOSCILLATORcFLcUSE Then
    Public stCND As TrimCondInfo                            ' トリマー加工条件(形式定義はRs232c.vb参照)

    '----- RS232Cポート情報定義
    Public stCOM As ComInfo                                 ' ポート情報(形式定義はRs232c.vb参照)
#End If
    Public Const cTIMEOUT As Long = 10000                   ' 応答待タイマ値(ms)

    '----- FL向けデフォルト設定ファイル ----
    Public Const DEF_FLPRM_SETFILEPATH As String = "c:\TRIM\"
    Public Const DEF_FLPRM_SETFILENAME As String = "c:\TRIM\defaultFlParamSet.xml"

    Public giTenKey_Btn As Short = 0                        ' 一時停止画面での「Ten Key On/Off」ボタンの初期値(0:ON(既定値), 1:OFF)
    Public giBpAdj_HALT As Short = 0                        ' 一時停止画面での「BPオフセット調整する/しない」(0:調整する(既定値), 1:調整しない)
    Public Const BLOCK_END As Short = 1         ' ブロック終了 
    Public Const PLATE_BLOCK_END As Short = 2   ' プレート・ブロック終了

    ' ＴＸ，ＴＹ関係　START
    Public Const MaxCntStep As Short = 256                  ' ｽﾃｯﾌﾟ最大件数
    Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
    Public Const HWND_TOPMOST As Short = -1                 ' ウィンドウを最前面に表示
    Public Const SWP_NOSIZE As Short = &H1S                 ' 現在のサイズを維持
    Public Const SWP_NOMOVE As Short = &H2S                 ' 現在の位置を維持
    Public Const KND_CHIP As Short = 1
    Public Const MaxCntCut As Short = MAXCTN                ' 最大ｶｯﾄ数
    Public Const MaxCutInfo As Short = MAXCTN               ' 最大ｶｯﾄ情報数
    Public Const CNS_CUTP_ST2 As String = "T"               ' ﾎﾟｼﾞｼｮﾆﾝｸﾞ無しSTカット
    Public Const CNS_CUTP_IX2 As String = "X"               ' ﾎﾟｼﾞｼｮﾆﾝｸﾞ無しｲﾝﾃﾞｯｸｽ
    ' ＴＸ，ＴＹ関係　END

    '----- パワーメータのデータ取得取得 -----
    Public Const PM_DTTYPE_NONE As Short = 0                ' なし
    Public Const PM_DTTYPE_IO As Short = 1                  ' Ｉ／Ｏ読取り
    Public Const PM_DTTYPE_USB As Short = 2                 ' ＵＳＢ

    Public INTERNAL_CAMERA As Integer = 0                   ' 内部ｶﾒﾗ
    Public EXTERNAL_CAMERA As Integer = 1                   ' 外部ｶﾒﾗ

    'V2.0.0.0⑭↓
    Public Const CLAMP_VACCUME_USE As Short = 1             'クランプ吸着有り
    Public Const CLAMP_ONLY_USE As Short = 2                'クランプのみ
    Public Const VACCUME_ONLY_USE As Short = 3              '吸着のみ
    'V2.0.0.0⑭↑

#Region "グローバル変数の定義"
    Public gCurPlateNo As Integer
    Public gCurBlockNo As Integer
    Public gbExitFlg As Boolean
    Public gbTenKeyFlg As Boolean = True                            ' テンキー入力フラグ
    Public gbChkboxHalt As Boolean = True                           ' ADJボタン状態(ON=ADJ ON, OFF=ADJ OFF)
    Public gObjADJ As frmFineAdjust = Nothing                              ' 一時停止画面オブジェクト
    '-------------------------------------------------------------------------------
    '   ＧＰＩＢ通信用定義
    '-------------------------------------------------------------------------------
    Public ObjGpib As GpibMaster = Nothing                  ' ＧＰＩＢ通信用オブジェクト
    Public gstrDeviceName As String = "GPIB000"             ' ＧＰＩＢデバイス名(デバイスマネージャで定義した名前)
    Public gDevId As Short = -1                             ' デバイスＩＤ
    Public gEOI As Short = 1                                ' EOI(0:出力しない, 0以外:出力する) 2013/3/9 0から1へ変更

    '-------------------------------------------------------------------------------
    '   その他の定義
    '-------------------------------------------------------------------------------
    '----- アプリケーション種別定義 -----  
    Public Const KND_USER As Integer = 9                    ' ユーザプログラム
    Public frmAutoObj As FormDataSelect                     ' 自動運転Formｵﾌﾞｼﾞｪｸﾄ
    Public ObjCrossLine As New TrimClassLibrary.TrimCrossLineClass()

    Public Const TARGET_DIGIT_DEFINE As String = "0.0000000"      'V2.0.0.0⑤

    Public ObjLoader As clsLoaderIf = Nothing                               'V2.2.0.0⑤
    Public ObjPlcIf As DllPlcIf.DllMelsecPLCIf                              'V2.2.0.0⑤
    Public objLoaderInfo As frmLoaderInfo                                   'V2.2.0.0⑤
    Public swMesureTrimtime As New System.Diagnostics.Stopwatch()           '処理時間の計測用       'V2.2.0.0⑤
    Public gdTrimtime As New TimeSpan                                       'トリミング時間保存用   'V2.2.0.0⑤
    Public gitacktTime As Integer                                           'タクトタイム保存用     'V2.2.0.0⑤
    Public gichangePlateTime As Integer                                     '基板交換時間保存用     'V2.2.0.0⑤
    Public MarkingCount As Integer                                          'マーキング処理基板数カウント用    'V2.2.1.7③


    '-------------------------------------------------------------------------------
    ' ■自動運転..
    '-------------------------------------------------------------------------------
    Public MSG_AUTO_01 As String '動作モード
    Public MSG_AUTO_02 As String 'マガジンモード
    Public MSG_AUTO_03 As String 'ロットモード
    Public MSG_AUTO_04 As String 'エンドレスモード
    Public MSG_AUTO_05 As String 'データファイル
    Public MSG_AUTO_06 As String '登録済みデータファイル
    Public MSG_AUTO_07 As String 'リストの1つ上へ
    Public MSG_AUTO_08 As String 'リストの1つ下へ
    Public MSG_AUTO_09 As String 'リストから削除
    Public MSG_AUTO_10 As String 'リストをクリア
    Public MSG_AUTO_11 As String '登録
    Public MSG_AUTO_12 As String 'OK
    Public MSG_AUTO_13 As String 'キャンセル
    Public MSG_AUTO_14 As String 'データ選択'
    Public MSG_AUTO_15 As String '登録リストを全て削除します。
    Public MSG_AUTO_16 As String 'よろしいですか？
    Public MSG_AUTO_17 As String 'エンドレスモード時は複数のデータファイルは選択できません。
    Public MSG_AUTO_18 As String 'データファイルを選択してください。
    Public MSG_AUTO_19 As String '編集中のデータを保存しますか？
    Public MSG_AUTO_20 As String '加工条件ファイルが存在しません。

    ' ＴＸ，ＴＹ関係　START
    Public gCmpTrimDataFlg As Short                         ' データ更新フラグ(0=更新なし, 1=更新あり)
    Public gTkyKnd As Short = KND_CHIP                      ' アプリケーション種別
    Public gfCorrectPosX As Double                          ' θ補正時のXYﾃｰﾌﾞﾙずれ量X(mm) ※ThetaCorrection()で設定
    Public gfCorrectPosY As Double                          ' θ補正時のXYﾃｰﾌﾞﾙずれ量Y(mm)
    ' ＴＸ，ＴＹ関係　END

    ' パワー調整(FL用)
    Public MSG_AUTOPOWER_01 As String
    Public MSG_AUTOPOWER_02 As String
    Public MSG_AUTOPOWER_03 As String
    Public MSG_AUTOPOWER_04 As String
    Public MSG_AUTOPOWER_05 As String

    ' 分布図
    Public PIC_TRIM_01 As String 'イニシャルテスト　分布図
    Public PIC_TRIM_02 As String 'ファイナルテスト　分布図
    Public PIC_TRIM_03 As String '良品
    Public PIC_TRIM_04 As String '不良品
    Public PIC_TRIM_05 As String '最小%
    Public PIC_TRIM_06 As String '最大%
    Public PIC_TRIM_07 As String '平均%
    Public PIC_TRIM_08 As String '標準偏差
    Public PIC_TRIM_09 As String '抵抗数
    Public PIC_TRIM_10 As String '分布図保存 
    Public MSG_TRIM_04 As String 'イニシャルテスト　分布図
    Public MSG_TRIM_05 As String 'ファイナルテスト　分布図

    '-------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------
    '   カット位置補正定義(補正実行)    
    '-------------------------------------------------------------------------------
    Public Const PTN_NONE As Integer = 0                    ' 補正実行 0:なし
    Public Const PTN_AUTO As Integer = 1                    ' 補正実行 1:自動
    Public Const PTN_MANUAL As Integer = 2                  ' 補正実行 2:手動
    Public Const PTN_AUTO_JUDGE As Integer = 3              ' 補正実行 3:自動ＮＧ判定あり

    '-------------------------------------------------------------------------------
    '   測定モード定義    
    '-------------------------------------------------------------------------------
    Public Const MEAS_JUDGE_NONE As Integer = 0             ' 測定なし
    Public Const MEAS_JUDGE_IT As Integer = 1               ' ITのみ
    Public Const MEAS_JUDGE_FT As Integer = 2               ' FTのみ
    Public Const MEAS_JUDGE_BOTH As Integer = 3             ' IT,FT両方
#End Region

#Region "判定判定モード"
    Public Const JUDGE_MODE_RATIO As Integer = 0            ' 0:比率(%)
    Public Const JUDGE_MODE_ABSOLUTE As Integer = 1         ' 1:数値(絶対値)
#End Region

#Region "ＴＸ、ＴＹ用構造体定義"

    '-------------------------------------------------------------------------------
    '   ＴＸ、ＴＹ用プレートデータ
    '-------------------------------------------------------------------------------
    Public Structure PlateInfo
        Dim intBlockCntXDir As Short                        ' ﾌﾞﾛｯｸ数Ｘ
        Dim intBlockCntYDir As Short                        ' ﾌﾞﾛｯｸ数Ｙ
        Dim dblBlockSizeXDir As Double                      ' ブロックサイズＸ   
        Dim dblBlockSizeYDir As Double                      ' ブロックサイズＹ   
        Dim dblTableOffsetXDir As Double                    ' ﾃｰﾌﾞﾙ位置ｵﾌｾｯﾄX
        Dim dblTableOffsetYDir As Double                    ' ﾃｰﾌﾞﾙ位置ｵﾌｾｯﾄY
        Dim dblBpOffSetXDir As Double                       ' ﾋﾞｰﾑ位置ｵﾌｾｯﾄX
        Dim dblBpOffSetYDir As Double                       ' ﾋﾞｰﾑ位置ｵﾌｾｯﾄY
        Dim intResistDir As Short                           ' 抵抗並び方向
        Dim intResistCntInBlock As Short                    ' 1ブロック内抵抗数
        Dim intResistCntInGroup As Short                    ' 1グループ内抵抗数
        Dim intGroupCntInBlockXBp As Short                  ' ブロック内ＢＰグループ数(サーキット数)
        Dim intGroupCntInBlockYStage As Short               ' ブロック内ステージグループ数
        Dim dblChipSizeXDir As Double                       ' ﾁｯﾌﾟｻｲｽﾞX
        Dim dblChipSizeYDir As Double                       ' ﾁｯﾌﾟｻｲｽﾞY
        Dim dblStepOffsetXDir As Double                     ' ｽﾃｯﾌﾟｵﾌｾｯﾄ量X
        Dim dblStepOffsetYDir As Double                     ' ｽﾃｯﾌﾟｵﾌｾｯﾄ量Y
        Dim dblBpGrpItv As Double                           ' BPグループ間隔（以前のCHIPのグループ間隔）
        Dim dblStgGrpItvX As Double                         ' X方向のステージグループ間隔（以前のＣＨＩＰのステップ間インターバル）
        Dim dblStgGrpItvY As Double                         ' Y方向のステージグループ間隔（以前のＣＨＩＰのステップ間インターバル）
        Dim intBlkCntInStgGrpX As Short                     ' X方向のステージグループ内ブロック数
        Dim intBlkCntInStgGrpY As Short                     ' Y方向のステージグループ内ブロック数
    End Structure
    Public typPlateInfo As PlateInfo                        ' ﾌﾟﾚｰﾄﾃﾞｰﾀ
    '--------------------------------------------------------------------------
    '   カットデータ構造体形式定義
    '--------------------------------------------------------------------------
    Public Structure CutList
        Dim intCutNo As Short                               ' ｶｯﾄ番号(1～n)
        Dim dblStartPointX As Double                        ' ｽﾀｰﾄﾎﾟｲﾝﾄX
        Dim dblStartPointY As Double                        ' ｽﾀｰﾄﾎﾟｲﾝﾄY
        Dim dblTeachPointX As Double                        ' ﾃｨｰﾁﾝｸﾞﾎﾟｲﾝﾄX
        Dim dblTeachPointY As Double                        ' ﾃｨｰﾁﾝｸﾞﾎﾟｲﾝﾄY
        Dim strCutType As String                            ' ｶｯﾄ形状
        Dim intCutAngle As Short                            ' カット角度     'V2.2.0.0②
        Dim intLTurnDir As Short                            ' ターン方向     'V2.2.0.0②
    End Structure
    '--------------------------------------------------------------------------
    '   抵抗データ構造体形式定義
    '--------------------------------------------------------------------------
    Public Structure ResistorInfo
        Dim intResNo As Short                               ' 抵抗番号(1～9999)
        Dim intCutCount As Short                            ' ｶｯﾄ数
        <VBFixedArray(MaxCutInfo)> Dim ArrCut() As CutList  ' ｶｯﾄ情報
        ' 構造体の初期化
        Public Sub Initialize()
            ReDim ArrCut(MaxCutInfo)
        End Sub
    End Structure

    Public typResistorInfoArray(MAXRNO) As ResistorInfo     ' 抵抗ﾃﾞｰﾀ

#End Region

#End Region

    'V2.0.0.0⑨↓
#Region "分布図"
    '----- 生産管理グラフフォームオブジェクト
    Public gObjFrmDistribute As Object                          ' frmDistribute

#End Region
    'V2.0.0.0⑨↑

    '=========================================================================
    '   メッセージ設定処理
    '=========================================================================
#Region "メッセージ初期設定処理"
    '''=========================================================================
    '''<summary>メッセージ初期設定処理</summary>
    '''<param name="language">(INP) 0=日本語, 1=英語</param>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub PrepareMessages(ByVal language As Short)

        Select Case language
            Case 0
                Call PrepareMessagesJapanese()
            Case 1
                Call PrepareMessagesEnglish()
            Case Else
                Call PrepareMessagesEnglish()
        End Select

    End Sub
#End Region

#Region "メッセージ初期設定(日本語)"
    '''=========================================================================
    '''<summary>メッセージ初期設定(日本語)</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub PrepareMessagesJapanese()

        ' エラーメッセージ
        MSG_DataNotLoad = "データが未ロードです。データをロードして下さい。" & vbCrLf
        MSG_SPRASH31 = "注意！！！"
        MSG_SPRASH32 = "スライドカバーが自動で閉じます"
        MSG_SPRASH52 = "スライドカバーが閉じています。" & ControlChars.NewLine & "トリミングを開始します。"

        MSG_105 = "前の画面に戻ります。よろしいですか？　　　　　　　　　　　　"

        MSG_136 = "シリアルポートＯＰＥＮエラー"
        MSG_137 = "シリアルポートＣＬＯＳＥエラー"
        MSG_138 = "シリアルポート送信エラー"
        MSG_139 = "シリアルポート受信エラー"
        MSG_140 = "ＦＬ側の加工条件の設定がありません。" + vbCrLf + "再度データをロードするか、編集画面から加工条件の設定を行ってください。"
        MSG_141 = "ＦＬ側加工条件のリードに失敗しました。"
        MSG_142 = "加工条件ファイルを作成しました"
        MSG_143 = "データをロードしました"
        MSG_144 = "データロードＮＧ"
        MSG_145 = "データをセーブしました"
        MSG_146 = "データセーブＮＧ"
        MSG_147 = "ＦＬへ加工条件を送信しました。"
        MSG_148 = "ＦＬへデータ送信中・・・・・・"
        MSG_150 = "ＦＬ通信異常。ＦＬとの通信に失敗しました。" + vbCrLf + "ＦＬと正しく接続できているか確認してください。"
        MSG_151 = "加工条件の設定に失敗しました。"
        MSG_152 = "加工条件の送信に失敗しました。" + vbCrLf + "再度データをロードするか、編集画面から加工条件の設定を行ってください。"
        MSG_153 = "カット位置補正対象の抵抗がありません"

        ' 操作ログ　メッセージ
        MSG_OPLOG_START = "ユーザプログラム起動"
        MSG_OPLOG_FUNC01 = "データロード"
        MSG_OPLOG_FUNC02 = "データセーブ"
        MSG_OPLOG_FUNC03 = "データ編集"
        MSG_OPLOG_FUNC04 = "マスタチェック"
        MSG_OPLOG_FUNC05 = "レーザ調整"
        MSG_OPLOG_FUNC06 = "ロット切替"
        MSG_OPLOG_FUNC07 = "プローブ位置合わせ"
        MSG_OPLOG_FUNC08 = "ティーチング"
        MSG_OPLOG_FUNC08S = "カット補正位置ティーチング"
        MSG_OPLOG_FUNC09 = "パターン登録"
        MSG_OPLOG_FUNC10 = "プローブ位置合わせ２"
        MSG_OPLOG_FUNC11 = "データ設定"
        MSG_OPLOG_END = "ユーザプログラム終了"
        MSG_OPLOG_TRIMST = "トリミング"
        MSG_OPLOG_LOTCHG = "ロット切替信号受信"
        MSG_OPLOG_STOP = "トリマ装置停止"
        MSG_OPLOG_LOTSET = "ロット情報データ設定"

        ' ■自動運転..
        MSG_AUTO_01 = "動作モード"
        MSG_AUTO_02 = "マガジンモード"
        MSG_AUTO_03 = "ロットモード"
        MSG_AUTO_04 = "エンドレスモード"
        MSG_AUTO_05 = "データファイル"
        MSG_AUTO_06 = "登録済みデータファイル"
        MSG_AUTO_07 = "リストの1つ上へ"
        MSG_AUTO_08 = "リストの1つ下へ"
        MSG_AUTO_09 = "リストから削除"
        MSG_AUTO_10 = "リストをクリア"
        MSG_AUTO_11 = "↓登録↓"
        MSG_AUTO_12 = "OK"
        MSG_AUTO_13 = "キャンセル"
        MSG_AUTO_14 = "データ登録"
        MSG_AUTO_15 = "登録リストを全て削除します。"
        MSG_AUTO_16 = "よろしいですか？"
        MSG_AUTO_17 = "エンドレスモード時は複数のデータファイルは選択できません。"
        MSG_AUTO_18 = "データファイルを選択してください。"
        MSG_AUTO_19 = "編集中のデータを保存しますか？"
        MSG_AUTO_20 = "加工条件ファイルが存在しません。"

        ' ＴＸ，ＴＹ関係　START
        ' frmMsgBox(画面終了確認)
        MSG_CLOSE_LABEL01 = "画面終了確認"
        MSG_CLOSE_LABEL02 = "はい(&Y)"
        MSG_CLOSE_LABEL03 = "いいえ(&N)"
        TITLE_TX = "チップサイズ(TX)ティーチング"
        TITLE_TY = "ステップサイズ(TY)ティーチング"
        LBL_TXTY_TEACH_03 = "補正量"
        LBL_TXTY_TEACH_04 = "補正比率"
        LBL_TXTY_TEACH_05 = "チップサイズ(mm)"
        LBL_TXTY_TEACH_07 = "補正前"
        LBL_TXTY_TEACH_08 = "補正後"
        LBL_TXTY_TEACH_09 = "グループインターバル"
        LBL_TXTY_TEACH_11 = "ステップインターバル"
        LBL_TXTY_TEACH_12 = "第１基準点"
        LBL_TXTY_TEACH_13 = "第２基準点"
        LBL_TXTY_TEACH_14 = "グループ"
        LBL_CMD_CANCEL = "キャンセル (&Q)"
        CMD_CANCEL = "キャンセル"
        INFO_MSG13 = "チップサイズ　ティーチング"
        INFO_MSG14 = "ステージグループ間隔ティーチング"
        INFO_MSG15 = "ステップオフセット量　ティーチング"
        INFO_MSG16 = "　　基準位置を合わせて下さい。"
        INFO_MSG17 = "　　移動：[矢印]  決定：[START]  中断：[RESET]" '& vbCrLf & "　　[HALT]で１つ前の処理に戻ります。"
        INFO_MSG18 = "第1グループ、第1抵抗基準位置のティーチング"
        INFO_MSG19 = "第"
        INFO_MSG20 = "グループ、最終抵抗基準位置のティーチング"
        INFO_MSG23 = "ＢＰグループ間隔ティーチング"
        INFO_MSG28 = "グループ、最終端位置のティーチング"
        INFO_MSG29 = "グループ、最先端位置のティーチング"
        INFO_MSG30 = "サーキット間隔ティーチング"
        INFO_MSG31 = "ステップオフセット量のティーチング"
        INFO_MSG32 = " (ＴＸ)"  '###084
        INFO_MSG33 = " (ＴＹ)"  '###084
        INFO_MSG34 = "ステップサイズ　ティーチング"
        ' ＴＸ，ＴＹ関係　END

        ' パワー調整(FL用)
        MSG_AUTOPOWER_01 = "パワー調整開始"
        MSG_AUTOPOWER_02 = "加工条件番号"
        MSG_AUTOPOWER_03 = "レーザパワー設定値"
        MSG_AUTOPOWER_04 = "電流値"
        MSG_AUTOPOWER_05 = "パワー調整未完了"

        ' 分布図
        MSG_TRIM_04 = "イニシャルテスト　分布図"
        MSG_TRIM_05 = "ファイナルテスト　分布図"
        PIC_TRIM_01 = "イニシャルテスト　分布図"
        PIC_TRIM_02 = "ファイナルテスト　分布図"
        PIC_TRIM_03 = "良品"
        PIC_TRIM_04 = "不良品"
        PIC_TRIM_05 = "最小%"
        PIC_TRIM_06 = "最大%"
        PIC_TRIM_07 = "平均%"
        PIC_TRIM_08 = "標準偏差"
        PIC_TRIM_09 = "抵抗数"
        PIC_TRIM_10 = "分布図保存"
    End Sub
#End Region

#Region "メッセージ初期設定(英語)"
    '''=========================================================================
    '''<summary>メッセージ初期設定(英語)</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub PrepareMessagesEnglish()

        ' エラーメッセージ
        MSG_DataNotLoad = "Data is not loaded. Please Load the data file." & vbCrLf
        MSG_SPRASH31 = "Cautions !!!"
        MSG_SPRASH32 = "Slide Cover Closes Automatically."
        MSG_SPRASH52 = "Slide Cover is closed." & ControlChars.NewLine & "Trimming will be started."           'V1.0.0.1②

        MSG_136 = "Serial Port Open Error."
        MSG_137 = "Serial Port Close Error."
        MSG_138 = "Serial Port Transmission Error."
        MSG_139 = "Serial Port Reception Error."
        MSG_140 = "There Is No Setting On The FL Side." + vbCrLf + "Please Load Data Or Set Condition From Edit Function."
        MSG_141 = "Condition Reading Error On The FL Side."
        MSG_142 = "Condition File Was Made."
        MSG_143 = "DATA LOAD OK"
        MSG_144 = "DATA LOAD NG"
        MSG_145 = "DATA SAVE OK"
        MSG_146 = "DATA SAVE NG"
        MSG_147 = "DATA SEND TO FL"
        MSG_148 = "Data sending to FL......"
        MSG_150 = "Connection error for FiberLaser." + vbCrLf + "Please confirm the connection."
        MSG_151 = "It Failed In The Setting Of Processing Conditions."
        MSG_152 = "It Failed In The Transmission Of The Condition Data." + vbCrLf + "Please Load Data Or Set Condition From Edit Function."
        MSG_153 = "No resistor to correct cutting position."

        ' 操作ログ　メッセージ
        MSG_OPLOG_START = "START USER PROGURAM"
        MSG_OPLOG_FUNC01 = "LOAD"
        MSG_OPLOG_FUNC02 = "SAVE"
        MSG_OPLOG_FUNC03 = "EDIT"
        MSG_OPLOG_FUNC04 = "MASTER CHECK"
        MSG_OPLOG_FUNC05 = "LASER"
        MSG_OPLOG_FUNC06 = "LOT CHANGE"
        MSG_OPLOG_FUNC07 = "PROBE"
        MSG_OPLOG_FUNC08 = "TEACH"
        MSG_OPLOG_FUNC08S = "CUTTING POSITION CORRECTION TEACHING"
        MSG_OPLOG_FUNC09 = "RECOG"
        MSG_OPLOG_FUNC10 = "PROBE2"
        MSG_OPLOG_FUNC11 = "DATA SET"
        MSG_OPLOG_END = "END USER PROGURAM"
        MSG_OPLOG_TRIMST = "TRIMMING"
        MSG_OPLOG_LOTCHG = "LOT CHANGE RECEIVE"
        MSG_OPLOG_STOP = "TRIMMER STOP"
        MSG_OPLOG_LOTSET = "LOT DATA INPUT"

        ' ＴＸ，ＴＹ関係　START
        'frmMsgBox(画面終了確認)
        MSG_CLOSE_LABEL01 = "Exit?"
        MSG_CLOSE_LABEL02 = "Yes(&Y)"
        MSG_CLOSE_LABEL03 = "No(&N)"
        TITLE_TX = "Chip size (TX) Teaching"
        TITLE_TY = "Step size (TY) Teaching"
        LBL_TXTY_TEACH_03 = "Correct quantity"
        LBL_TXTY_TEACH_04 = "Correct ratio"
        LBL_TXTY_TEACH_05 = "Chip size(mm)"
        LBL_TXTY_TEACH_07 = "Before"
        LBL_TXTY_TEACH_08 = "After"
        LBL_TXTY_TEACH_09 = "Group interval(mm)"
        LBL_TXTY_TEACH_11 = "Step interval"
        LBL_TXTY_TEACH_12 = "The 1st datum point."
        LBL_TXTY_TEACH_13 = "The 2nd datum point."
        LBL_TXTY_TEACH_14 = "Group"
        LBL_CMD_CANCEL = "Cancel (&Q)"
        CMD_CANCEL = "Cancel"
        INFO_MSG13 = "CHIP SIZE TEACHING"
        INFO_MSG14 = "STAGE INTERVAL TEACHING"
        INFO_MSG15 = "STEP OFFSET TEACHING"
        INFO_MSG16 = "    Please unite a standard position."
        INFO_MSG17 = "    MOVE:[Arrow]  OK:[START]  CANCEL:[RESET]" '& vbCrLf & "    It returns to the processing before one by the HALT key."
        INFO_MSG18 = "<Group No.1> The 1st resistance standard position." ''''2009/07/03 NETでは「resistance→circuit」(18,20-22)
        INFO_MSG19 = "<Group No."
        INFO_MSG20 = "> The last resistance standard position."
        INFO_MSG23 = "BP GROUP INTERVAL TEACHING"
        INFO_MSG28 = "> The Final Edge Positionlast."
        INFO_MSG29 = "> The State-Of-The-Art Position."
        INFO_MSG30 = "CIRCUIT INTERVAL TEACHING"
        INFO_MSG31 = "STEP OFFSET TEACHING"
        INFO_MSG32 = " (TX)"  '###084
        INFO_MSG33 = " (TY)"  '###084
        INFO_MSG33 = "STEP SIZE TEACHING"
        ' ＴＸ，ＴＹ関係　END

        ' パワー調整(FL用)
        MSG_AUTOPOWER_01 = "Start Power Adjustment"
        MSG_AUTOPOWER_02 = "Condition No."
        MSG_AUTOPOWER_03 = "Laser Power"
        MSG_AUTOPOWER_04 = "Current"
        MSG_AUTOPOWER_05 = "Power Adjustment Failed."

        ' 分布図
        MSG_TRIM_04 = "INITIAL TEST DISTRIBUTION MAP"
        MSG_TRIM_05 = "FINAL TEST DISTRIBUTION MAP"
        PIC_TRIM_01 = "INITIAL TEST DISTRIBUTION MAP"
        PIC_TRIM_02 = "FINAL TEST DISTRIBUTION MAP"
        PIC_TRIM_03 = "OK" '"良品"
        PIC_TRIM_04 = "NG" '"不良品"
        PIC_TRIM_05 = "MIN %" '"最小%"
        PIC_TRIM_06 = "MAX %" '"最大%"
        PIC_TRIM_07 = "AVG %" '"平均%"
        PIC_TRIM_08 = "Std Dev" '"標準偏差"
        PIC_TRIM_09 = "Res Num" '"抵抗数"
        PIC_TRIM_10 = "DISTRIBUTION MAP SAVE"
    End Sub
#End Region


#Region "ログ画面(Main.txtLog)に文字列を表示する"
    '''=========================================================================
    '''<summary>ログ画面(frmMain.txtLog)に文字列を表示する</summary>
    '''<param name="s">(INP) 表示文字列</param>
    '''<remarks></remarks>
    '''=========================================================================
    Public Function Z_PRINT(Optional ByVal s As String = vbCrLf) As Integer

        Z_PRINT = LogPrint(s)
        Exit Function

    End Function
#End Region

#Region "ログ画面クリア"
    '''=========================================================================
    '''<summary>ログ画面クリア</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub Z_CLS()

        Try
            ''V2.2.0.0⑰↓
            If giTxtLogType <> 0 Then
                Static hWnd As IntPtr = ObjMain.txtlog.Handle
                Const WM_SETTEXT As Integer = &HC
                SendMessageString(hWnd, WM_SETTEXT, 0, "")                  ' 削除
            Else
                'ObjMain.txtLog.Text = ""
                ObjMain.lstLog.Items.Clear()
                ObjMain.lstLog.Items.Add(" ")
            End If
            ''V2.2.0.0⑰↑


        Catch ex As Exception

        End Try

    End Sub
#End Region

#Region "ログ画面表示サブ"
    '''=========================================================================
    '''<summary>ログ画面表示サブ</summary>
    '''<param name="s">(INP) 表示文字列</param>
    '''<remarks></remarks>
    '''=========================================================================
    Private Function LogPrint(ByVal s As String) As Integer

        Dim strMSG As String

        Try


            ''V2.2.0.0⑰↓
            '' 表示の最後までスクロールする
            LogPrint = 0                                ' Return値 = 正常
            'ObjMain.txtLog.Text = ObjMain.txtLog.Text + s + "  "
            'ObjMain.txtLog.Focus()
            'ObjMain.txtLog.SelectionStart = ObjMain.txtLog.Text.Length
            'ObjMain.txtLog.ScrollToCaret()
            If giTxtLogType <> 0 Then
                ''V2.2.0.0⑰
                Z_PRINT_MSG(s)
            Else
                With ObjMain.lstLog                                         ' ###lstLog
                    .BeginUpdate()
                    .Items.RemoveAt(.Items.Count - 1)
                    .Items.Add(s)
                    .Items.Add(" ")
                    .SelectedIndex = (.Items.Count - 1)
                    .ClearSelected()
                    .EndUpdate()
                End With
            End If
            ''V2.2.0.0⑰↑

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "LogPrint() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            LogPrint = -1                               ' Return値 = エラー
        End Try

        Exit Function

    End Function


    ''' <summary>ログ画面に文字列を表示する</summary>
    ''' <param name="s"></param>
    ''' <remarks>'#4.12.2.0④</remarks>
    Public Function Z_PRINT_MSG(ByVal s As String) As Integer

        '#4.12.2.0④                    ↓
        'Static hWnd As IntPtr = ObjMain.lstLog.Handle
        Static hWnd As IntPtr = ObjMain.txtlog.Handle
        Const WM_GETTEXTLENGTH As Integer = &HE
        'Const LB_GETTEXT As Integer = &HE
        Const EM_SETSEL As Integer = &HB1
        Const EM_REPLACESEL As Integer = &HC2
        Const LB_ADDSTRING As Integer = &H180
        Const WM_COPYDATA As Integer = &H4A
        Dim result As Integer

        Try
            Dim test As String = Strings.Right(s, 1)
            Dim len As Integer = SendMessage(hWnd, WM_GETTEXTLENGTH, 0, 0)      ' 文字数取得
            SendMessage(hWnd, EM_SETSEL, len, len)                              ' カーソルを末尾へ
            If test <> "" AndAlso InStr(s, Environment.NewLine) = False Then
                s = s + Environment.NewLine
            End If
            SendMessageString(hWnd, EM_REPLACESEL, 0, s)  ' テキストに文字列を追加する
            'SendMessageString(hWnd, EM_REPLACESEL, 0, s & Environment.NewLine)  ' テキストに文字列を追加する
            ' SendMessageString(hWnd, LB_ADDSTRING, 0, s & Environment.NewLine)  ' テキストに文字列を追加する

            'Dim cds As COPYDATASTRUCT

            'len = s.Length
            'cds.dwData = 0        '使用しない
            'cds.lpData = s      'テキストのポインターをセット
            'cds.cbData = len + 1     '長さをセット
            ''文字列を送る
            'result = SendMessage(hWnd, WM_COPYDATA, 0, cds)

            Z_PRINT_MSG = cFRS_NORMAL


            ' トラップエラー発生時
        Catch ex As Exception
            Dim strMSG As String = "i-TKY.LogPrint() TRAP ERROR = " & ex.Message
            MsgBox(strMSG)
            'MessageBox.Show(Me, strMSG)
        End Try
    End Function

#End Region

    '=========================================================================
    '   ローダ入出力処理
    '=========================================================================
#Region "ローダ出力サブ"
    '''=========================================================================
    '''<summary>ローダ出力サブ</summary>
    '''<param name="LDON"> (INP) ONビットデータ</param>
    '''<param name="LDOFF">(INP) OFFビットデータ</param>
    '''<remarks></remarks>
    '''=========================================================================
    Public Function Sub_ATLDSET(ByVal LDON As Integer, ByVal LDOFF As Integer) As Integer

        Dim strMSG As String

        Try
            ' ローダー出力(ON,OFF)
            Sub_ATLDSET = Form1.System1.Z_ATLDSET(LDON, LDOFF)

            ' IOモニタ表示
            gdwATLDDATA = gdwATLDDATA And (LDOFF Xor &HFFFF)
            gdwATLDDATA = gdwATLDDATA Or LDON
            Call IoMonitor(gdwATLDDATA, 1)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Sub_ATLDSET() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
        Exit Function
    End Function
#End Region

#Region "ローダ入力サブ(デバッグ用)"
    '''=========================================================================
    '''<summary>ローダ入力サブ(デバッグ用)</summary>
    '''<param name="Index"> (INP) ONビットデータ</param>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub DEBUG_ReadHostCommand(ByVal Index As Integer)

        Dim strMSG As String

        Try
            Select Case Index
                Case 0
                    gDebugHostCmd = gDebugHostCmd Xor cHSTcRDY
                Case 1
                    gDebugHostCmd = gDebugHostCmd Xor cHSTcAUTO
                Case 2
                    gDebugHostCmd = gDebugHostCmd Xor cHSTcSTATE
                Case 3
                    gDebugHostCmd = gDebugHostCmd Xor cHSTcTRMCMD
                Case 4
                    gDebugHostCmd = gDebugHostCmd Xor &H10          ' Bit4:未使用
                Case 5
                    gDebugHostCmd = gDebugHostCmd Xor &H20          ' Bit5:未使用
                Case 6
                    gDebugHostCmd = 0
                Case 7
                    gDebugHostCmd = &HFFFF
            End Select

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "DEBUG_ReadHostCommand() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try

        Exit Sub
    End Sub
#End Region

#Region "IOモニタ表示(ｵｰﾄﾛｰﾀﾞｰI/O)"
    '''=========================================================================
    ''' <summary>IOモニタ表示(ｵｰﾄﾛｰﾀﾞｰI/O)</summary>
    ''' <param name="whcmd">(INP) I/Oﾃﾞｰﾀ(16BIT)</param>
    ''' <param name="io">   (INP) ﾃﾞｰﾀ種別(0=ﾛｰﾀﾞｰ→ﾄﾘﾏｰ, 1=ﾄﾘﾏｰ→ﾛｰﾀﾞｰ)</param>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub IoMonitor(ByVal whcmd As Integer, ByVal io As Integer)

        Dim strMSG As String

        Try

#If cIOcMONITORcENABLED = 1 Then                    ' IOﾓﾆﾀ表示する

            Dim i As Integer


            If io = 0 Then
                For i = 0 To 15
                    If whcmd And (2 ^ i) Then
                        ObjMain.HostSignal(i).BackColor = Color.Red
                        ObjMain.HostSignal(i).Refresh()
                    Else
                        ObjMain.HostSignal(i).BackColor = Color.White
                        ObjMain.HostSignal(i).Refresh()
                    End If
                Next
            Else
                For i = 0 To 15
                    If whcmd And (2 ^ i) Then
                        ' BIT0の動作中／停止中だけハードで反転させている。
                        If i = 0 Then
                            ObjMain.HostSignal(i + 16).BackColor = Color.White
                            ObjMain.HostSignal(i + 16).Refresh()
                        Else
                            ObjMain.HostSignal(i + 16).BackColor = Color.Lime
                            ObjMain.HostSignal(i + 16).Refresh()
                        End If
                    Else
                        If i = 0 Then
                            ObjMain.HostSignal(i + 16).BackColor = Color.Lime
                            ObjMain.HostSignal(i + 16).Refresh()
                        Else
                            ObjMain.HostSignal(i + 16).BackColor = Color.White
                            ObjMain.HostSignal(i + 16).Refresh()
                        End If
                    End If
                Next
            End If

#End If
            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Globals.IoMonitor() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

    'V2.2.1.7① ↓
#Region "テキストボックスの文字列が数字変換できるか確認"
    ''' <summary>テキストボックスの文字列が数字変換できるか確認（テキストボックス用）</summary>
    ''' <param name="cTextBox">確認するﾃｷｽﾄﾎﾞｯｸｽ</param>
    ''' <returns>(-1)=ｴﾗｰ</returns>
    Public Function CheckNumeric(ByRef cTextBox As cTxt_) As Integer
        Dim ret As Integer = 0
        Try

            '数値チェック
            If IsNumeric(cTextBox.Text) Then
                'Nop
            Else
                MsgBox("数値を入力してください。")
                ret = -1
            End If
        Catch ex As Exception
            ret = -1
        Finally
            CheckNumeric = ret
        End Try

    End Function
#End Region
    'V2.2.1.7① ↑

End Module