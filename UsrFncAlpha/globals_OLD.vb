''===============================================================================
''   Description  : グローバル定数の定義
''
''   Copyright(C) : OMRON LASERFRONT INC. 2010
''
'===============================================================================
Option Strict Off
Option Explicit On

Imports System.Runtime.InteropServices

Module Globals_Renamed
#Region "グローバル定数/変数の定義"
    '    '===========================================================================
    '    '   グローバル定数/変数の定義
    '    '===========================================================================
    '    '-------------------------------------------------------------------------------
    '    '   DLL定義
    '    '-------------------------------------------------------------------------------
    '    '----- WIN32 API -----
    '    ' ウィンドウ表示の操作のAPI
    '    Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
    '    Public Const HWND_TOPMOST As Short = -1                 ' ウィンドウを最前面に表示
    '    Public Const SWP_NOSIZE As Short = &H1S                 ' 現在のサイズを維持
    '    Public Const SWP_NOMOVE As Short = &H2S                 ' 現在の位置を維持

    '    '---------------------------------------------------------------------------
    '    '   アプリケーション名/アプリケーション種別/アプリケーションモード
    '    '---------------------------------------------------------------------------
    '    '----- 強制終了用アプリケーション -----
    '    Public Const APP_FORCEEND As String = "c:\Trim\ForceEndProcess.exe"

    '    '----- パス　-----
    '    Public Const OCX_PATH As String = "c:\Trim\ocx\"        '----- OCX登録パス
    '    Public Const DLL_PATH As String = "c:\Trim\"            '----- DLL登録パス

    '    '----- アプリケーション名 -----
    '    Public Const APP_TKY As String = "TKY"
    '    Public Const APP_CHIP As String = "TKYCHIP"
    '    Public Const APP_NET As String = "TKYNET"

    '    '----- アプリケーション種別 -----
    '    Public Const KND_TKY As Short = 0
    '    Public Const KND_CHIP As Short = 1
    '    Public Const KND_NET As Short = 2
    '    Public Const MACHINE_TYPE_SL432 As String = "SL432R"                 ' 系名
    '    Public Const MACHINE_TYPE_SL436 As String = "SL436R"                 ' 系名

    '    Public gAppName As String                               ' アプリケーション名
    '    Public gTkyKnd As Short                                 ' アプリケーション種別

    '    '----- 画像表示プログラムの表示位置 -----
    '    'Public Const FORM_X As Integer = 4                     ' コントロール上部左端座標X ###050
    '    'Public Const FORM_Y As Integer = 20                    ' コントロール上部左端座標Y ###050
    '    Public Const FORM_X As Integer = 0                      ' コントロール上部左端座標X ###050
    '    Public Const FORM_Y As Integer = 0                      ' コントロール上部左端座標Y ###050

    '    'サブフォームの表示位置目印の表示位置オフセット
    '    Public Const DISPOFF_SUBFORM_TOP As Integer = 12

    '    '----- シグナルタワー制御種別 -----                     ' ###007
    '    Public Const SIGTOWR_NORMAL As Short = 0                ' 標準３色制御
    '    Public Const SIGTOWR_SPCIAL As Short = 1                ' ４色制御(日立ｵｰﾄﾓｰﾃｨﾌﾞ殿特注)

    '    '----- アプリケーションモード ----- (注)OcxSystem定義と一致させる必要有り
    '    Public giAppMode As Short

    '    Public Const APP_MODE_IDLE As Short = 0                 ' トリマ装置アイドル中
    '    Public Const APP_MODE_LOAD As Short = 1                 ' ファイルロード(F1)
    '    Public Const APP_MODE_SAVE As Short = 2                 ' ファイルセーブ(F2)
    '    Public Const APP_MODE_EDIT As Short = 3                 ' 編集画面      (F3)
    '    '                                                       ' 空き
    '    Public Const APP_MODE_LASER As Short = 5                ' レーザー調整  (F5)
    '    Public Const APP_MODE_LOTCHG As Short = 6               ' ロット切替    (F6) ※ユーザプロ対応
    '    Public Const APP_MODE_PROBE As Short = 7                ' プローブ      (F7)
    '    Public Const APP_MODE_TEACH As Short = 8                ' ティーチング  (F8)
    '    Public Const APP_MODE_RECOG As Short = 9                ' パターン登録  (F9)
    '    Public Const APP_MODE_EXIT As Short = 10                ' 終了 　　　　 (F11)
    '    Public Const APP_MODE_TRIM As Short = 11                ' トリミング中
    '    Public Const APP_MODE_CUTPOS As Short = 12              ' ｶｯﾄ位置補正   (S-F8)
    '    Public Const APP_MODE_PROBE2 As Short = 13              ' プローブ2     (F10) ※ユーザプロ対応
    '    Public Const APP_MODE_LOGGING As Short = 14             ' ロギング      (F6) 

    '    ' CHIP,NET系
    '    Public Const APP_MODE_TTHETA As Short = 40              ' Ｔθ(θ角度補正)ティーチング
    Public Const APP_MODE_TX As Short = 41                  ' TXティーチング
    '    Public Const APP_MODE_TY As Short = 42                  ' TYティーチング
    '    Public Const APP_MODE_TY2 As Short = 43                 ' TY2ティーチング
    '    Public Const APP_MODE_EXCAM_R1TEACH As Short = 44       ' 外部カメラR1ティーチング【外部カメラ】
    '    Public Const APP_MODE_EXCAM_TEACH As Short = 45         ' 外部カメラティーチング【外部カメラ】
    Public Const APP_MODE_CARIB_REC As Short = 46           ' 画像登録(キャリブレーション補正用)【外部カメラ】
    '    Public Const APP_MODE_CARIB As Short = 47               ' キャリブレーション【外部カメラ】
    '    Public Const APP_MODE_CUTREVISE_REC As Short = 48       ' 画像登録(カット位置補正用)【外部カメラ】
    Public Const APP_MODE_CUTREVIDE As Short = 49           ' カット位置補正【外部カメラ】
    '    Public Const APP_MODE_AUTO As Short = 50                ' 自動運転　　　
    '    Public Const APP_MODE_LOADERINIT As Short = 51          ' ローダ原点復帰
    '    Public Const APP_MODE_LDR_ALRM As Short = 52            ' ローダアラーム画面    '###088
    Public Const APP_MODE_FINEADJ As Short = 53             ' 一時停止画面          '###088

    '    ' NET系
    '    Public Const APP_MODE_CIRCUIT As Short = 60             ' サーキットティーチング

    '    '---------------------------------------------------------------------------
    '    '----- 機能選択定義テーブルのｲﾝﾃﾞｯｸｽ定義 -----          '                         TKY CHIP NET
    '    '                                                       '                (○:標準,△:ｵﾌﾟｼｮﾝ,×:未ｻﾎﾟｰﾄ)
    '    Public Const F_LOAD As Short = 0                        ' LOADボタン              ○  ○   ○
    '    Public Const F_SAVE As Short = 1                        ' SAVEボタン              ○  ○   ○
    '    Public Const F_EDIT As Short = 2                        ' EDITボタン              ○  ○   ○
    '    Public Const F_LASER As Short = 3                       ' LASERボタン             ○  ○   ○
    '    Public Const F_LOG As Short = 4                         ' LOGGINGボタン           ○  ○   ○
    '    Public Const F_PROBE As Short = 5                       ' PROBEボタン             ○  ○   ○
    '    Public Const F_TEACH As Short = 6                       ' TEACHボタン             ○  ○   ○
    '    Public Const F_CUTPOS As Short = 7                      ' CUTPOSボタン            △  ×   ×
    '    Public Const F_RECOG As Short = 8                       ' RECOGボタン             ○  ○   ○
    '    ' CHIP,NET系
    '    Public Const F_TTHETA As Short = 9                      ' Tθボタン               ×  △   △
    '    Public Const F_TX As Short = 10                         ' TXボタン                ×  ○   ○
    '    Public Const F_TY As Short = 11                         ' TYボタン                ×  ○   ○
    '    Public Const F_TY2 As Short = 12                        ' TY2ボタン               ×  △   △
    '    Public Const F_EXR1 As Short = 13                       ' 外部ｶﾒﾗR1ﾃｨｰﾁﾝｸﾞﾎﾞﾀﾝ    ×  △   △
    '    Public Const F_EXTEACH As Short = 14                    ' 外部ｶﾒﾗﾃｨｰﾁﾝｸﾞﾎﾞﾀﾝ      ×  △   △
    '    Public Const F_CARREC As Short = 15                     ' ｷｬﾘﾌﾞﾚｰｼｮﾝ補正登録ﾎﾞﾀﾝ  ×  △   △
    '    Public Const F_CAR As Short = 16                        ' ｷｬﾘﾌﾞﾚｰｼｮﾝﾎﾞﾀﾝ          ×  △   △
    '    Public Const F_CUTREC As Short = 17                     ' ｶｯﾄ補正登録ﾎﾞﾀﾝ         ×  △   △
    '    Public Const F_CUTREV As Short = 18                     ' ｶｯﾄ位置補正ﾎﾞﾀﾝ         ×  △   △
    '    ' NET系
    '    Public Const F_CIRCUIT As Short = 19                    ' ｻｰｷｯﾄﾃｨｰﾁﾝｸﾞﾎﾞﾀﾝ        ×  ×   ○

    '    ' SL436R CHIP,NET系 
    '    Public Const F_AUTO As Short = 20                       ' AUTOボタン              -   ○   ○
    '    Public Const F_LOADERINI As Short = 21                  ' LOADER INITボタン       -   ○   ○

    '    Public Const MAX_FNCNO As Short = 22                    ' 機能選択定義テーブルのデータ数 

    '    '---------------------------------------------------------------------------

    '    '---------------------------------------------------------------------------
    '    '   最大値/最小値
    '    '---------------------------------------------------------------------------
    '    Public Const cMAXOptFlgNUM As Short = 5                 ' OcxSystem用ｺﾝﾊﾟｲﾙｵﾌﾟｼｮﾝの数 (最大5個)

    '    '----- 各入力項目の範囲 -----
    '    Public Const gMIN As Short = 0
    '    Public Const gMAX As Short = 1

    '    '----- ZZMOVE()の移動指定 -----
    '    Public Const MOVE_RELATIVE As Short = 0                 ' 相対値指定 
    '    Public Const MOVE_ABSOLUTE As Short = 1                 ' 絶対値指定

    '    '----- ZINPSTS()の入力箇所指定  -----
    '    Public Const GET_CONSOLE_INPUT As Short = 1             ' コンソール
    '    Public Const GET_INTERLOCK_INPUT As Short = 2           ' インターロック

    '    '----- 画像登録用パラメータ -----
    '    Public Const PTN_NUM_MAX As Short = 50                  ' テンプレート番号(1-50)
    '    Public Const GRP_NUM_MAX As Short = 999                 ' ﾃﾝﾌﾟﾚｰﾄｸﾞﾙｰﾌﾟ番号(1-999)

    '    Public Const INIT_THRESH_VAL As Double = 0.7            ' 閾値初期値
    '    Public Const INIT_CONTRAST_VAL As Integer = 216         ' コントラスト初期値
    '    Public Const INIT_BRIGHTNESS_VAL As Integer = 0         ' 輝度初期値
    '    Public Const MIN_CONTRAST_VAL As Integer = 0            ' コントラスト最小値
    '    Public Const MAX_CONTRAST_VAL As Integer = 511          ' コントラスト最大値
    '    Public Const MIN_BRIGHTNESS_VAL As Integer = -128       ' 輝度最小値
    '    Public Const MAX_BRIGHTNESS_VAL As Integer = 127        ' 輝度最大値

    '    '----- ローダ用 ----- 
    '    Public Const LALARM_COUNT As Integer = 128              ' 最大アラーム数
    '    Public Const MG_UP As Integer = 1                       ' マガジンＵＰ      2013.01.28  '###182
    '    Public Const MG_DOWN As Integer = 0                     ' マガジンＤＯＷＮ  2013.01.28  '###182


    '    '----- マーキング抵抗番号 -----
    '    Public Const MARKING_RESNO_SET As Integer = 1000        ' 抵抗番号1000番以降はマーキング用の抵抗番号

    '    '---------------------------------------------------------------------------
    '    '   システムパラメータ(形式はDllSysprm.dllで定義)
    '    '---------------------------------------------------------------------------
    '    Public gDllSysprmSysParam_definst As New DllSysprm.SysParam
    '    Public gSysPrm As DllSysprm.SYSPARAM_PARAM              ' システムパラメータ
    '    Public OptVideoPrm As DllSysprm.OPT_VIDEO_PRM           ' Video.ocx用オプション定義
    '    Public giTrimExe_NoWork As Short = 0                    ' 手動モード時、載物台に基板なしでトリミング実行する(0)/しない(1)　###240
    Public giTenKey_Btn As Short = 0                        ' 一時停止画面での「Ten Key On/Off」ボタンの初期値(0:ON(既定値), 1:OFF)　###268
    Public giBpAdj_HALT As Short = 0                        ' 一時停止画面での「BPオフセット調整する/しない」(0:調整する(既定値), 1:調整しない)　###269

    '    '----- ONLINE -----
    '    Public Const TYPE_OFFLINE As Short = 0                  ' OFFLINE
    '    Public Const TYPE_ONLINE As Short = 1                   ' ONLINE
    '    Public Const TYPE_MANUAL As Short = 2                   ' SLIDE COVER+XY移動処理

    '    '----- ProbeType -----
    '    Public Const TYPE_PROBE_NON As Short = 0                ' NON
    '    Public Const TYPE_PROBE_STD As Short = 1                ' STANDARD

    '    '----- XY Table Exist Flag -----
    '    Public Const TYPE_XYTABLE_NON As Short = 0              ' NON
    '    Public Const TYPE_XYTABLE_X As Short = 1                ' X Only
    '    Public Const TYPE_XYTABLE_Y As Short = 2                ' Y Only
    '    Public Const TYPE_XYTABLE_XY As Short = 3               ' XY

    '    '----- 吸着ﾘﾄﾗｲ処理 -----
    '    Public Const VACCUME_ERRRETRY_OFF As Short = 0          ' Not retry
    '    Public Const VACCUME_ERRRETRY_ON As Short = 1           ' Retry
    '    Public Const RET_VACCUME_RETRY As Short = 1
    '    Public Const RET_VACCUME_CANCEL As Short = 2

    '    '----- ｶｽﾀﾏｲｽﾞ -----
    '    Public Const customROHM As Short = 1                    ' ﾛｰﾑ殿向け仕様
    '    Public Const customASAHI As Short = 2                   ' 朝日電子殿向け仕様
    '    Public Const customSUSUMU As Short = 3                  ' 進殿向け仕様
    '    Public Const customKOA As Short = 4                     ' KOA(匠の里)殿向け仕様
    '    Public Const customKOAEW As Short = 5                   ' KOA(EW)殿向け仕様

    '    '----- パワーメータのデータ取得取得 -----
    '    Public Const PM_DTTYPE_NONE As Short = 0                ' なし
    '    Public Const PM_DTTYPE_IO As Short = 1                  ' Ｉ／Ｏ読取り
    '    Public Const PM_DTTYPE_USB As Short = 2                 ' ＵＳＢ

    '    '---------------------------------------------------------------------------
    '    '   ステージ動作関係
    '    '---------------------------------------------------------------------------
    '    ' ステップ方向
    '    Public Const STEP_RPT_NON As Short = 0      ' ステップ＆リピート方向（なし）
    '    Public Const STEP_RPT_X As Short = 1        ' ステップ＆リピート方向（X方向）
    '    Public Const STEP_RPT_Y As Short = 2        ' ステップ＆リピート方向（Y方向）
    '    Public Const STEP_RPT_CHIPXSTPY As Short = 3 ' ステップ＆リピート方向（X方向チップ幅ステップ＋Y方向）
    '    Public Const STEP_RPT_CHIPYSTPX As Short = 4 ' ステップ＆リピート方向（Y方向チップ幅ステップ＋X方向）

    '    ' BP基準方向
    '    Public Const BP_DIR_RIGHTUP As Short = 0    ' BP基準右上（プラス方向）←↓　　　 1 ＿ ＿ 0
    '    Public Const BP_DIR_LEFTUP As Short = 1     ' BP基準左上（プラス方向）↓→　　　　|＿|＿|
    '    Public Const BP_DIR_RIGHTDOWN As Short = 2  ' BP基準右下（プラス方向）←↑        |＿|＿|
    '    Public Const BP_DIR_LEFTDOWN As Short = 3   ' BP基準左下（プラス方向）↑→　　　 3　　　 2

    '    Public Const BP_DIR_RIGHT As Short = 0      ' BP-X方向基準右
    '    Public Const BP_DIR_LEFT As Short = 1       ' BP-X方向基準左

    Public Const BLOCK_END As Short = 1         ' ブロック終了 
    Public Const PLATE_BLOCK_END As Short = 2   ' プレート・ブロック終了

    '    '----- その他 -----
    '    ' FLSET関数のモード
    '    Public Const FLMD_CNDSET As Integer = 0                 ' 加工条件設定
    '    Public Const FLMD_BIAS_ON As Integer = 1                ' BIAS ON
    '    Public Const FLMD_BIAS_OFF As Integer = 2               ' BIAS OFF(LaserOff関数内でBIAS OFFはしている)


    '    '---------------------------------------------------------------------------
    '    '   制御フラグ
    '    '---------------------------------------------------------------------------
    '    Public gCmpTrimDataFlg As Short                         ' データ更新フラグ(0=更新なし, 1=更新あり)
    '    Public giTrimErr As Short                               ' ﾄﾘﾏｰ ｴﾗｰ ﾌﾗｸﾞ ※ｴﾗｰ時はｸﾗﾝﾌﾟｸﾗﾝﾌﾟOFF時ﾄﾘﾏ動作中OFFをﾛｰﾀﾞｰに送信しない
    '    '                                                       ' B0 : 吸着ｴﾗｰ(EXIT)
    '    '                                                       ' B1 : その他ｴﾗｰ
    '    '                                                       ' B2 : 集塵機ｱﾗｰﾑ検出
    '    '                                                       ' B3 : 軸ﾘﾐｯﾄ､軸ｴﾗｰ､軸ﾀｲﾑｱｳﾄ
    '    '                                                       ' B4 : 非常停止
    '    '                                                       ' B5 : ｴｱｰ圧ｴﾗｰ

    '    Public gLoadDTFlag As Boolean                            ' ﾃﾞｰﾀﾛｰﾄﾞ済ﾌﾗｸﾞ(False:ﾃﾞｰﾀ未ﾛｰﾄﾞ, True:ﾃﾞｰﾀﾛｰﾄﾞ済)
    '    Public gbInitialized As Boolean                         ' True=原点復帰済, False=原点復帰未
    '    'Public bFgfrmDistribution As Boolean                    ' 生産ｸﾞﾗﾌ表示ﾌﾗｸﾞ(TRUE:表示 FALSE:非表示)
    '    Public gLoggingHeader As Boolean                        ' ﾛｸﾞﾍｯﾀﾞｰ書込み指示ﾌﾗｸﾞ(TRUE:出力)
    '    Public gESLog_flg As Boolean                            ' ESログフラグ(Flase=ログOFF, True=ログON)
    '    '' '' ''Public giAdjKeybord As Short                             ' トリミング中ADJ機能キーボード矢印(0:入力なし 1:上 2:下 3:右 4:左 )
    '    Public gPrevInterlockSw As Short

    '    Public gbCanceled As Boolean ' ←　各画面処理でPrivateで持つ 

    '    '-------------------------------------------------------------------------------
    '    '   オブジェクト定義
    '    '-------------------------------------------------------------------------------
    '    '----- VB6のOCX -----
    '    'Public ObjSys As Object                                 ' OcxSystem.ocx
    '    'Public ObjUtl As Object                                 ' OcxUtility.ocx
    '    'Public ObjHlp As Object                                 ' OcxAbout.ocx
    '    'Public ObjPas As Object                                 ' OcxPassword.ocx
    '    'Public ObjMTC As Object                                 ' OcxManualTeach.ocx
    '    'Public ObjTch As Object                                 ' Teach.ocx
    '    'Public ObjPrb As Object                                 ' Probe.ocx
    '    'Public ObjVdo As Object                                 ' Video.ocx
    '    'Public ObjPrt As Object                                ' OcxPrint.ocx
    '    Public ObjMON(32) As Object
    '    Public gparModules As MainModules                                   ' 親側メソッド呼出しオブジェクト(OcxSystem用) '###061
    '    Public ObjCrossLine As New TrimClassLibrary.TrimCrossLineClass()    ' 補正クロスライン表示用オブジェクト ###232 

    '    '---------------------------------------------------------------------------
    '    ' トリミング動作モード
    '    '---------------------------------------------------------------------------
    '    Public Const TRIM_MODE_ITTRFT As Integer = 0    'イニシャルテスト＋トリミング＋ファイナルテスト実行
    '    Public Const TRIM_MODE_TRFT As Integer = 1      'トリミング＋ファイナルテスト実行
    '    Public Const TRIM_MODE_FT As Integer = 2        'ファイナルテスト実行（判定）
    '    Public Const TRIM_MODE_MEAS As Integer = 3      '測定実行
    '    Public Const TRIM_MODE_POSCHK As Integer = 4    'ポジションチェック
    '    Public Const TRIM_MODE_CUT As Integer = 5       'カット実行
    '    Public Const TRIM_MODE_STPRPT As Integer = 6    'ステップ＆リピート実行
    '    Public Const TRIM_MODE_TRIMCUT As Integer = 7   'トリミングモードでのカット実行


    '    '-------------------------------------------------------------------------------
    '    ' トリミング結果
    '    '-------------------------------------------------------------------------------
    '    '----- トリミング結果値（INTRIMで設定）
    '    '//Trim result
    '    '//0:未実施   1:OK       2:ITNG      3:FTNG     4:SKIP
    '    '//5:RATIO    6:ITHI NG  7:ITLO NG   8:FTHI NG  9:FTLO NG
    '    '//10:        11:        12:         13:        14:
    '    '//15:異形面付けによりSKIP
    '    Public Const RSLT_NO_JUDGE As Integer = 0
    '    Public Const RSLT_OK As Integer = 1
    '    Public Const RSLT_IT_NG As Integer = 2
    '    Public Const RSLT_FT_NG As Integer = 3
    '    Public Const RSLT_SKIP As Integer = 4
    '    Public Const RSLT_RATIO As Integer = 5
    '    Public Const RSLT_IT_HING As Integer = 6
    '    Public Const RSLT_IT_LONG As Integer = 7
    '    Public Const RSLT_FT_HING As Integer = 8
    '    Public Const RSLT_FT_LONG As Integer = 9
    '    Public Const RSLT_RANGEOVER As Integer = 10
    '    Public Const RSLT_OPENCHK_NG As Integer = 20
    '    Public Const RSLT_SHORTCHK_NG As Integer = 21
    '    Public Const RSLT_IKEI_SKIP As Integer = 15

    '    '----- 生産管理グラフフォームオブジェクト
    '    Public gObjFrmDistribute As Object                      ' frmDistribute

    '    '----- 生産管理情報用配列 -----
    '    Public Const MAX_FRAM1_ARY As Integer = 15              ' ラベル配列数
    '    '                                                       ' 生産管理情報のラベル配列のインデックス 
    '    Public Const FRAM1_ARY_GO As Integer = 0                ' GO数(サーキット数 or 抵抗数)
    '    Public Const FRAM1_ARY_NG As Integer = 1                ' NG数(サーキット数 or 抵抗数)
    '    Public Const FRAM1_ARY_NGPER As Integer = 2             ' NG%
    '    Public Const FRAM1_ARY_PLTNUM As Integer = 3            ' PLATE数
    '    Public Const FRAM1_ARY_REGNUM As Integer = 4            ' RESISTOR数
    '    Public Const FRAM1_ARY_ITHING As Integer = 5            ' IT HI NG数
    '    Public Const FRAM1_ARY_FTHING As Integer = 6            ' FT HI NG数
    '    Public Const FRAM1_ARY_ITLONG As Integer = 7            ' IT LO NG数
    '    Public Const FRAM1_ARY_FTLONG As Integer = 8            ' FT LO NG数
    '    Public Const FRAM1_ARY_OVER As Integer = 9              ' OVER数
    '    Public Const FRAM1_ARY_ITHINGP As Integer = 10          ' IT HI NG%
    '    Public Const FRAM1_ARY_FTHINGP As Integer = 11          ' FT HI NG%
    '    Public Const FRAM1_ARY_ITLONGP As Integer = 12          ' IT LO NG%
    '    Public Const FRAM1_ARY_FTLONGP As Integer = 13          ' FT LO NG%
    '    Public Const FRAM1_ARY_OVERP As Integer = 14            ' OVER NG%

    '    Public Fram1LblAry(MAX_FRAM1_ARY) As System.Windows.Forms.Label     ' 生産管理情報のラベル配列

    '    '-------------------------------------------------------------------------------
    '    '   gMode(OcxSystemのfrmReset()の処理モード)
    '    '-------------------------------------------------------------------------------
    '    Public Const cGMODE_ORG As Short = 0                    '  0 : 原点復帰
    '    Public Const cGMODE_ORG_MOVE As Short = 1               '  1 : 原点位置移動
    '    Public Const cGMODE_START_RESET As Short = 2            '  2 : 操作確認画面(START/RESET待ち)
    '    '                                                       '  3 :
    '    '                                                       '  4 :
    '    Public Const cGMODE_EMG As Short = 5                    '  5 : 非常停止メッセージ表示
    '    '                                                       '  6 :
    '    Public Const cGMODE_SCVR_OPN As Short = 7               '  7 : トリミング中のスライドカバー開メッセージ表示
    '    Public Const cGMODE_CVR_OPN As Short = 8                '  8 : トリミング中の筐体カバー開メッセージ表示
    '    Public Const cGMODE_SCVRMSG As Short = 9                '  9 : スライドカバー開メッセージ表示(トリミング中以外)
    '    Public Const cGMODE_CVRMSG As Short = 10                ' 10 : 筐体カバー開確認メッセージ表示(トリミング中以外)
    '    Public Const cGMODE_ERR_HW As Short = 11                ' 11 : ハードウェアエラー(カバーが閉じてます)メッセージ表示
    '    Public Const cGMODE_ERR_HW2 As Short = 12               ' 12 : ハードウェアエラーメッセージ表示
    '    Public Const cGMODE_CVR_LATCH As Short = 13             ' 13 : カバー開ラッチメッセージ表示
    '    Public Const cGMODE_CVR_CLOSEWAIT As Short = 14         ' 14 : 筐体カバークローズもしくはインターロック解除待ち
    '    Public Const cGMODE_ERR_DUST As Short = 20              ' 20 : 集塵機異常検出メッセージ表示
    '    Public Const cGMODE_ERR_AIR As Short = 21               ' 21 : エアー圧エラー検出メッセージ表示

    '    Public Const cGMODE_ERR_HING As Short = 40              ' 40 : 連続HI-NGｴﾗｰ(ADVｷｰ押下待ち)
    '    Public Const cGMODE_SWAP As Short = 41                  ' 41 : 基板交換(STARTｷｰ押下待ち)
    '    Public Const cGMODE_XYMOVE As Short = 42                ' 42 : 終了時のﾃｰﾌﾞﾙ移動確認(STARTｷｰ押下待ち)
    '    Public Const cGMODE_ERR_REPROBE As Short = 43           ' 43 : 再プロービング失敗(STARTｷｰ押下待ち) SL436R用
    '    Public Const cGMODE_LDR_ALARM As Short = 44             ' 44 : ローダアラーム発生   SL436R用
    '    Public Const cGMODE_LDR_START As Short = 45             ' 45 : 自動運転開始(STARTｷｰ押下待ち)   SL436R用
    '    Public Const cGMODE_LDR_TMOUT As Short = 46             ' 46 : ローダ通信タイムアウト  SL436R用
    '    Public Const cGMODE_LDR_END As Short = 47               ' 47 : 自動運転終了(STARTｷｰ押下待ち)   SL436R用
    '    Public Const cGMODE_LDR_ORG As Short = 48               ' 48 : ローダ原点復帰  SL436R用

    '    Public Const cGMODE_AUTO_LASER As Short = 50            ' 50 : 自動レーザパワー調整

    '    Public Const cGMODE_LDR_CHK As Short = 60               ' 60 : ローダ状態チェック(起動時ﾛｰﾀﾞ自動ﾓｰﾄﾞ/動作中)
    '    Public Const cGMODE_LDR_ERR As Short = 61               ' 61 : ローダ状態エラー(ﾛｰﾀﾞ自動でﾛｰﾀﾞ無)
    '    Public Const cGMODE_LDR_MNL As Short = 62               ' 62 : カバー開後のローダ手動モード処理
    '    Public Const cGMODE_LDR_WKREMOVE As Short = 63          ' 63 : 残基板取り除きメッセージ  SL436R用
    '    Public Const cGMODE_LDR_RSTAUTO As Short = 64           ' 64 : 自動運転中止メッセージ  SL436R用 ###124
    '    Public Const cGMODE_LDR_WKREMOVE2 As Short = 65         ' 65 : 残基板取り除きメッセージ(APP終了)  SL436R用 ###175
    '    Public Const cGMODE_LDR_STAGE_ORG As Short = 66         ' 66 : ステージ原点移動 SL436R用 ###188

    '    Public Const cGMODE_OPT_START As Short = 70             ' 70 : ﾄﾘﾐﾝｸﾞ開始時のｽﾀｰﾄSW押下待ち
    '    Public Const cGMODE_OPT_END As Short = 71               ' 71 : ﾄﾘﾐﾝｸﾞ終了時のｽﾗｲﾄﾞｶﾊﾞｰ開待ち

    '    Public Const cGMODE_MSG_DSP As Short = 90               ' 90 : 指定メッセージ表示(STARTキー押下待ち)

    '    ' リミットセンサー& 軸エラー & タイムアウトメッセージ
    '    ' ※TrimErrNo.vbに移動
    '    '                                                       ' ※(注)
    '    'Public Const cGMODE_TO_AXISX As Short = 101             ' 101: X軸エラー(タイムアウト)
    '    'Public Const cGMODE_TO_AXISY As Short = 102             ' 102: Y軸エラー(タイムアウト)
    '    'Public Const cGMODE_TO_AXISZ As Short = 103             ' 103: Z軸エラー(タイムアウト)
    '    'Public Const cGMODE_TO_AXIST As Short = 104             ' 104: θ軸エラー(タイムアウト)

    '    ''                                                       '【ソフトリミットエラー】
    '    'Public Const cGMODE_SL_AXISX As Short = 105             ' 105: X軸ソフトリミットエラー
    '    'Public Const cGMODE_SL_AXISY As Short = 106             ' 106: Y軸ソフトリミットエラー
    '    'Public Const cGMODE_SL_AXISZ As Short = 107             ' 107: Z軸ソフトリミットエラー
    '    'Public Const cGMODE_SL_BPX As Short = 110               ' 110: BP X軸ソフトリミットエラー
    '    'Public Const cGMODE_SL_BPY As Short = 111               ' 111: BP Y軸ソフトリミットエラー

    '    'Public Const cGMODE_TO_ROTATT As Short = 108            ' 108: ロータリアッテネータエラー(タイムアウト)
    '    'Public Const cGMODE_TO_AXISZ2 As Short = 109            ' 109: Z2軸エラー(タイムアウト)

    '    'Public Const cGMODE_SRV_ARM As Short = 202              ' 202: サーボアラーム
    '    'Public Const cGMODE_AXISX_LIM As Short = 203            ' 203: X軸リミット
    '    'Public Const cGMODE_AXISY_LIM As Short = 204            ' 204: Y軸リミット
    '    'Public Const cGMODE_AXISZ_LIM As Short = 205            ' 205: Z軸リミット
    '    'Public Const cGMODE_AXIST_LIM As Short = 206            ' 206: θ軸リミット
    '    'Public Const cGMODE_RATT_LIM As Short = 207             ' 207: ロータリーアッテネータリミット
    '    'Public Const cGMODE_AXISZ2_LIM As Short = 208           ' 208: Z2軸リミット

    '    'Public Const cGMODE_BASE_ERR As Short = 200             ' Base Num.
    '    ''                                                       '【X軸エラー】
    '    'Public Const cGMODE_AXISX_AOFF As Short = 211           ' 211: X軸エラー(Bit All Off)
    '    'Public Const cGMODE_AXISX_AON As Short = 212            ' 212: X軸エラー(Bit All On)
    '    'Public Const cGMODE_AXISX_ARM As Short = 213            ' 213: X軸アラーム
    '    'Public Const cGMODE_AXISX_PML As Short = 214            ' 214: ±X軸リミット
    '    'Public Const cGMODE_AXISX_PLM As Short = 215            ' 215: +X軸リミット
    '    'Public Const cGMODE_AXISX_MLM As Short = 216            ' 216: -X軸リミット
    '    ''                                                       '【Y軸エラー】
    '    'Public Const cGMODE_AXISY_AOFF As Short = 221           ' 221: Y軸エラー(Bit All Off)
    '    'Public Const cGMODE_AXISY_AON As Short = 222            ' 222: Y軸エラー(Bit All On)
    '    'Public Const cGMODE_AXISY_ARM As Short = 223            ' 223: Y軸アラーム
    '    'Public Const cGMODE_AXISY_PML As Short = 224            ' 224: ±Y軸リミット
    '    'Public Const cGMODE_AXISY_PLM As Short = 225            ' 225: +Y軸リミット
    '    'Public Const cGMODE_AXISY_MLM As Short = 226            ' 226: -Y軸リミット
    '    ''                                                       '【Z軸エラー】
    '    'Public Const cGMODE_AXISZ_AOFF As Short = 231           ' 231: Z軸エラー(Bit All Off)
    '    'Public Const cGMODE_AXISZ_AON As Short = 232            ' 232: Z軸エラー(Bit All On)
    '    'Public Const cGMODE_AXISZ_ARM As Short = 233            ' 233: Z軸アラーム
    '    'Public Const cGMODE_AXISZ_PML As Short = 234            ' 234: ±Z軸リミット
    '    'Public Const cGMODE_AXISZ_PLM As Short = 235            ' 235: +Z軸リミット
    '    'Public Const cGMODE_AXISZ_MLM As Short = 236            ' 236: -Z軸リミット
    '    'Public Const cGMODE_AXISZ_ORG As Short = 237            ' 237: Z軸原点復帰未完了
    '    ''                                                       '【θ軸エラー】
    '    'Public Const cGMODE_AXIST_AOFF As Short = 241           ' 241: θ軸エラー(Bit All Off)
    '    'Public Const cGMODE_AXIST_AON As Short = 242            ' 242: θ軸エラー(Bit All On)
    '    'Public Const cGMODE_AXIST_ARM As Short = 243            ' 243: θ軸アラーム
    '    'Public Const cGMODE_AXIST_PML As Short = 244            ' 244: ±θ軸リミット
    '    'Public Const cGMODE_AXIST_PLM As Short = 245            ' 245: +θ軸リミット
    '    'Public Const cGMODE_AXIST_MLM As Short = 246            ' 246: -θ軸リミット
    '    ''                                                       '【Z2軸エラー】
    '    'Public Const cGMODE_AXISZ2_AOFF As Short = 251          ' 251: Z2軸エラー(Bit All Off)
    '    'Public Const cGMODE_AXISZ2_AON As Short = 252           ' 252: Z2軸エラー(Bit All On)
    '    'Public Const cGMODE_AXISZ2_ARM As Short = 253           ' 253: Z2軸アラーム
    '    'Public Const cGMODE_AXISZ2_PML As Short = 254           ' 254: ±Z2軸リミット
    '    'Public Const cGMODE_AXISZ2_PLM As Short = 255           ' 255: +Z2軸リミット
    '    'Public Const cGMODE_AXISZ2_MLM As Short = 256           ' 256: -Z2軸リミット
    '    'Public Const cGMODE_AXISZ2_ORG As Short = 257           ' 257: Z2軸原点復帰未完了
    '    ''                                                       '【ﾛｰﾀﾘｱｯﾃﾈｰﾀｰｴﾗｰ】
    '    'Public Const cGMODE_ROTATT_AOFF As Short = 261          ' 261: ﾛｰﾀﾘｱｯﾃﾈｰﾀｰｴﾗｰ(Bit All Off)
    '    'Public Const cGMODE_ROTATT_AON As Short = 262           ' 262: ﾛｰﾀﾘｱｯﾃﾈｰﾀｰｴﾗｰ(Bit All On)
    '    'Public Const cGMODE_ROTATT_ARM As Short = 263           ' 263: ﾛｰﾀﾘｱｯﾃﾈｰﾀｰｱﾗｰﾑ
    '    'Public Const cGMODE_ROTATT_PML As Short = 264           ' 264: ±ﾛｰﾀﾘｱｯﾃﾈｰﾀｰﾘﾐｯﾄ
    '    'Public Const cGMODE_ROTATT_PLM As Short = 265           ' 265: +ﾛｰﾀﾘｱｯﾃﾈｰﾀｰﾘﾐｯﾄ
    '    'Public Const cGMODE_ROTATT_MLM As Short = 266           ' 266: -ﾛｰﾀﾘｱｯﾃﾈｰﾀｰﾘﾐｯﾄ

    '    ''-------------------------------------------------------------------------------
    '    ''   DllTrimFnc.dllの戻り値(上記以外のパラメータエラー等)
    '    ''-------------------------------------------------------------------------------
    '    'Public Const cFNC_ERR_TRIMRTN_ERR As Short = 99         ' コマンド実行エラー(DllTrimFncから99で返ってくるもの)
    '    'Public Const cFNC_ERR_CMD_NOTSPT As Short = 301         ' 未サポートコマンド
    '    'Public Const cFNC_ERR_CMD_PRM As Short = 302            ' パラメータエラー
    '    'Public Const cFNC_ERR_CMD_LIM_L As Short = 303          ' パラメータ下限値エラー
    '    'Public Const cFNC_ERR_CMD_LIM_U As Short = 304          ' パラメータ上限値エラー
    '    'Public Const cFNC_ERR_RT2WIN_SEND As Short = 305        ' INTime→Windows送信エラー
    '    'Public Const cFNC_ERR_RT2WIN_RECV As Short = 306        ' INTime→Windows受信エラー
    '    'Public Const cFNC_ERR_WIN2RT_SEND As Short = 307        ' Windows→INTime送信エラー
    '    'Public Const cFNC_ERR_WIN2RT_RECV As Short = 308        ' Windows→INTime受信エラー

    '    '-------------------------------------------------------------------------------
    '    '   戻り値(frmReset()他)
    '    '-------------------------------------------------------------------------------
    '    ' ※TrimErrNo.vbに移動
    '    'Public Const cFRS_NORMAL As Short = 0                   ' 正常
    '    'Public Const cFRS_ERR_ADV As Short = 1                  ' OK(ADVｷｰ)       ← START/RESET待ち時
    '    'Public Const cFRS_ERR_START As Short = 1                ' START(ADVｷｰ)    ← START/RESET待ち時
    '    'Public Const cFRS_ERR_HLT As Short = 2                  ' HALTｷｰ
    '    'Public Const cFRS_ERR_RST As Short = 3                  ' Cancel(RESETｷｰ) ← START/RESET待ち時
    '    'Public Const cFRS_ERR_Z As Short = 4                    ' ZｷｰON/OFF
    '    'Public Const cFRS_TxTy As Short = 5                     ' TX2/TY2押下

    '    'Public Const cFRS_ERR_CVR As Short = -1                 ' 筐体カバー開検出
    '    'Public Const cFRS_ERR_SCVR As Short = -2                ' スライドカバー開検出
    '    'Public Const cFRS_ERR_LATCH As Short = -3               ' カバー開ラッチ検出

    '    'Public Const cFRS_ERR_EMG As Short = -11                ' 非常停止
    '    'Public Const cFRS_ERR_DUST As Short = -12               ' 集塵機異常検出
    '    'Public Const cFRS_ERR_AIR As Short = -13                ' エアー圧エラー検出
    '    'Public Const cFRS_ERR_MVC As Short = -14                ' ﾏｽﾀｰﾊﾞﾙﾌﾞ回路状態エラー検出
    '    'Public Const cFRS_ERR_HW As Short = -15                 ' ハードウェアエラー検出

    '    ''----- IO制御タイムアウト -----
    '    'Public Const cFRS_TO_SCVR_CL As Short = -21             ' タイムアウト(スライドカバー閉待ち)
    '    'Public Const cFRS_TO_SCVR_OP As Short = -22             ' タイムアウト(スライドカバー開待ち)
    '    'Public Const cFRS_TO_SCVR_ON As Short = -23             ' タイムアウト(ｽﾗｲﾄﾞｶﾊﾞｰｽﾄｯﾊﾟｰ行待ち)
    '    'Public Const cFRS_TO_SCVR_OFF As Short = -24            ' タイムアウト(ｽﾗｲﾄﾞｶﾊﾞｰｽﾄｯﾊﾟｰ戻待ち)
    '    'Public Const cFRS_TO_CLAMP_ON As Short = -25            ' タイムアウト(クランプＯＮ)
    '    'Public Const cFRS_TO_CLAMP_OFF As Short = -26           ' タイムアウト(クランプＯＦＦ)
    '    'Public Const cFRS_TO_PM_DW As Short = -27               ' タイムアウト(パワーメータ下降移動)
    '    'Public Const cFRS_TO_PM_UP As Short = -28               ' タイムアウト(パワーメータ上昇移動)
    '    'Public Const cFRS_TO_PM_FW As Short = -29               ' タイムアウト(パワーメータ測定端移動)
    '    'Public Const cFRS_TO_PM_BK As Short = -30               ' タイムアウト(パワーメータ待機端移動)

    '    ''----- 軸エラー & タイムアウト -----
    '    ''                                                       ' -101～-266(軸エラー & タイムアウト)※上記参照
    '    ''----- Main()の戻り値 -----
    '    '' 画面処理用
    '    'Public Const cFRS_FNG_DATA As Short = -80               ' データ未ロード
    '    'Public Const cFRS_FNG_CMD As Short = -81                ' 他コマンド実行中
    '    'Public Const cFRS_FNG_PASS As Short = -82               ' ﾊﾟｽﾜｰﾄﾞ入力ｴﾗｰ

    '    '' トリミング用
    '    'Public Const cFRS_TRIM_NG As Short = -90                ' トリミングNG
    '    'Public Const cFRS_ERR_TRIM As Short = -91               ' トリマエラー
    '    'Public Const cFRS_ERR_PTN As Short = -92                ' パターン認識エラー

    '    ''----- パラメータエラー等(メッセージ表示はしない) -----
    '    'Public Const cFRS_ERR_CMD_NOTSPT As Short = -301        ' 未サポートコマンド
    '    'Public Const cFRS_ERR_CMD_PRM As Short = -302           ' パラメータエラー
    '    'Public Const cFRS_ERR_CMD_LIM_L As Short = -303         ' パラメータ下限値エラー
    '    'Public Const cFRS_ERR_CMD_LIM_U As Short = -304         ' パラメータ上限値エラー
    '    'Public Const cFRS_ERR_CMD_OBJ As Short = -305           ' オブジェクト未設定(Utilityｵﾌﾞｼﾞｪｸﾄ他)
    '    'Public Const cFRS_ERR_CMD_EXE As Short = -306           ' コマンド実行エラー(DllTrimFncから99で返ってくるもの)
    '    ''                                                       ' (注)cFRS_ERR_CMD_EXE～cFRS_ERR_CMD_NOTSPTで判定している箇所があるため
    '    ' '' 　　                                                   追加する場合は注意(cFRS_ERR_CMD_EXEをずらして番号を振り直す)
    '    ''----- Video.OCXのエラー -----
    '    'Public Const cFRS_VIDEO_PTN As Short = -401             ' パターン認識エラー
    '    'Public Const cFRS_VIDEO_PT1 As Short = -402             ' パターン認識エラー(補正位置1)
    '    'Public Const cFRS_VIDEO_PT2 As Short = -403             ' パターン認識エラー(補正位置2)
    '    'Public Const cFRS_VIDEO_COM As Short = -404             ' 通信エラー(CV3000)

    '    'Public Const cFRS_VIDEO_INI As Short = -411             ' 初期化が行われていません
    '    'Public Const cFRS_VIDEO_IN2 As Short = -412             ' 初期化済み
    '    'Public Const cFRS_VIDEO_FRM As Short = -413             ' フォーム表示中
    '    'Public Const cFRS_VIDEO_PRP As Short = -414             ' プロパティ値不正
    '    'Public Const cFRS_VIDEO_GRP As Short = -415             ' ﾃﾝﾌﾟﾚｰﾄｸﾞﾙｰﾌﾟ番号ｴﾗｰ
    '    'Public Const cFRS_VIDEO_MXT As Short = -416             ' テンプレート数 > MAX

    '    'Public Const cFRS_VIDEO_UXP As Short = -421             ' 予期せぬエラー
    '    'Public Const cFRS_VIDEO_UX2 As Short = -422             ' 予期せぬエラー2

    '    'Public Const cFRS_MVC_UTL As Short = -431               ' MvcUtil エラー
    '    'Public Const cFRS_MVC_PT2 As Short = -432               ' MvcPt2 エラー
    '    'Public Const cFRS_MVC_10 As Short = -433                ' Mvc10 エラー

    '    ''----- ファイル入出力エラー -----
    '    'Public Const cFRS_FIOERR_INP As Short = -501            ' ファイル入力エラー
    '    'Public Const cFRS_FIOERR_OUT As Short = -502            ' ファイル出力エラー

    '    'Public Const cERR_TRAP As Short = -999                  ' 例外エラー

    '    '---------------------------------------------------------------------------
    '    '   補正クロスライン表示用パラメータ
    '    '---------------------------------------------------------------------------
    '    Public gstCLC As CLC_PARAM                              ' 補正クロスライン表示用パラメータ

    '    '---------------------------------------------------------------------------
    '    '   ファイルパス関係
    '    '---------------------------------------------------------------------------
    '    Public gStrTrimFileName As String                       ' ﾄﾘﾐﾝｸﾞﾃﾞｰﾀﾌｧｲﾙ名

    '    ''''    lib.bas　でしか使用されていない。
    '    Public gsDataLogPath As String

    '    Public gbCutPosTeach As Boolean                         ' CutPosTeach(表示中:True, 非表示:False)

    '    '---------------------------------------------------------------------------
    '    '   変数定義
    '    '---------------------------------------------------------------------------

    '    '----- パターン認識用 -----
    '    Public giTempGrpNo As Integer                           ' テンプレートグループ番号(1～999)
    '    Public giTempNo As Integer                              ' テンプレート番号

    '    '----- カット位置補正用構造体 -----
    '    Public Structure CutPosCorrect_Info                     ' パターン登録情報
    '        Dim intFLG As Short                                 ' カット位置補正フラグ(0:しない, 1:する)
    '        Dim intGRP As Short                                 ' パターンｸﾞﾙｰﾌﾟ番号(1-999)
    '        Dim intPTN As Short                                 ' パターン番号(1-50)
    '        Dim dblPosX As Double                               ' パターン位置X(補正位置ティーチング用)
    '        Dim dblPosY As Double                               ' パターン位置Y(補正位置ティーチング用)
    '        Dim intDisp As Short                                ' パターン認識時の検索枠表示(0:なし, 1:あり)
    '    End Structure

    '    Public Const MaxRegNum As Short = 256                   ' 抵抗数の最大値
    '    Public Const MaxCutNum As Short = 30                    ' カットの最大値
    '    Public Const MaxDataNum As Short = 7681                 ' 抵抗数*カットの最大数+1
    '    Public stCutPos(MaxRegNum + 1) As CutPosCorrect_Info        ' パターン登録情報

    '    Public giCutPosRNum As Short                            ' カット位置補正する抵抗数
    '    'Public giCutPosRSLT(MaxRegNum) As Short                 ' パターン認識結果(0:補正なし, 1:OK, 2:NGｽｷｯﾌﾟ)
    '    'Public gfCutPosDRX(MaxRegNum) As Double                 ' ズレ量X
    '    'Public gfCutPosDRY(MaxRegNum) As Double                 ' ズレ量Y
    '    Public gfCutPosCoef(MaxRegNum) As Double                '  一致度

    '    '----- θ補正用 -----
    '    Public gfCorrectPosX As Double                          ' θ補正時のXYﾃｰﾌﾞﾙずれ量X(mm) ※ThetaCorrection()で設定
    '    Public gfCorrectPosY As Double                          ' θ補正時のXYﾃｰﾌﾞﾙずれ量Y(mm)
    '    Public gbInPattern As Boolean                           ' 位置補正処理中
    '    Public gbRotCorrectCancel As Short                      ' 0:OK, n < 0: 位置補正をキャンセルした or 位置補正エラー

    '    '----- デジタルＳＷ -----
    '    'Public gDigH As Short                                   ' デジタルＳＷ(Hight)
    '    'Public gDigL As Short                                   ' デジタルＳＷ(Low)
    '    'Public gDigSW As Short                                  ' デジタルＳＷ
    '    Public gPrevTrimMode As Short                           ' デジタルＳＷ値退避域

    '    '----- GPIB用 -----
    '    Public giGpibDefAdder As Short = 21                     ' 初期設定(機器ｱﾄﾞﾚｽ)

    '    '----- その他 -----
    '    Public giIX2LOG As Short = 0                            ' IX2ログ(0=無効, 1=有効)　###231
    '    Public giTablePosUpd As Short = 0                       ' テーブル1,2座標を更新する/しない(VIDEO.OCX用オプション)　###234

    '    ''''    複数個所でFalseに設定しているが、Trueに設定されることはない。
    '    ''''    フラグとして機能はしていないので、コード確認の上削除。
    '    'Public OKFlag As Boolean                    'OKボタン押下の有無

    '    ''''    初期化のみ
    '    'Public gRegisterExceptMarkingCnt As Short '抵抗数（マーキングを除く数) @@@007
    '    'Public gsSystemPassword As String
    '    'Public gLoggingEnd As Boolean

    '    ' '' '' ''----- 生産管理情報 -----
    '    '' '' ''Public glCircuitNgTotal As Integer                      ' 不良サーキット数
    '    '' '' ''Public glCircuitGoodTotal As Integer                    ' 良品サーキット数
    '    '' '' ''Public glPlateCount As Integer                          ' プレート処理数
    '    '' '' ''Public glGoodCount As Integer                           ' 良品抵抗数
    '    '' '' ''Public glNgCount As Integer                             ' 不良抵抗数
    '    '' '' ''Public glITHINGCount As Integer                         ' IT HI NG数
    '    '' '' ''Public glITLONGCount As Integer                         ' IT LO NG数
    '    '' '' ''Public glFTHINGCount As Integer                         ' FT HI NG数
    '    '' '' ''Public glFTLONGCount As Integer                         ' FT LO NG数
    '    '' '' ''Public glITOVERCount As Integer                         ' ITｵｰﾊﾞｰﾚﾝｼﾞ数


    '    Public gfPreviousPrbBpX As Double                       ' BP論理座標上の位置X (BSIZE+BPOFFSET相対)
    '    Public gfPreviousPrbBpY As Double                       '                   Y

    '    ''''------------------------------------------------

    '    ''''---------------------------------------------------
    '    ''''　090413 minato
    '    ''''    ProbeTeachで設定し、ResistorGraphで使用しているのみ。
    '    ''''    内部で出来るように見直す。
    '    '---------------------------------------------------------------------------
    '    '   全抵抗測定のグラフ表示用
    '    '---------------------------------------------------------------------------
    '    Public giMeasureResistors As Short                      ' 抵抗数
    '    Public giMeasureResiNum(512) As Double                  ' 抵抗番号
    '    Public gfMeasureResiOhm(512) As Double                  ' 測定した抵抗値
    '    Public gfResistorTarget(512) As Double                  ' 目標値
    '    Public gfMeasureResiPos(2, 512) As Double               ' カットスタートポイント
    '    Public giMeasureResiRst(512) As Short                   ' トリミング結果

    '    Public Const cMEASUREcOK As Short = 1                   ' OK
    '    Public Const cMEASUREcIT As Short = 2                   ' IT ERROR
    '    Public Const cMEASUREcFT As Short = 3                   ' FT ERROR
    '    Public Const cMEASUREcNA As Short = 4                   ' 未測定


    '    '===============================================================================
    '    Public ExitFlag As Short
    '    Public gMode As Short 'モード

    '    'INIファイル取得データ
    '    ''''(2010/11/16) 動作確認後下記コメントは削除
    '    'Public gStartX As Double 'プローブ初期値X
    '    'Public gStartY As Double 'プローブ初期値Y

    '    ' レーザー調整
    '    ''''    frmReset、LASER_teaching　で使用
    '    Public gfLaserContXpos As Double
    '    Public gfLaserContYpos As Double

    '    '画像ハンドル
    '    'Public mlHSKDib As Integer '白黒
    '    '表示位置
    '    'Public mtDest As RECT
    '    'Public mtSrc As RECT
    '    'Public gVideoStarted As Boolean

    '    ''----- ｱﾌﾟﾘﾓｰﾄﾞ ----- (注)OcxSystem定義と一致させる必要有り
    '    'Public giAppMode As Short

    '    ''データ編集パスワード関連
    '    'Public gbPassSucceeded As Boolean

    '    'Public gLoggingHeader As Boolean                    ' ﾍｯﾀﾞｰ書込み指示ﾌﾗｸﾞ(TRUE:出力)
    '    'Public gbLogHeaderWrite As Boolean ' ログのヘッダ出力フラグ @@@082

    '    'Public giOpLogFileHandle As Short ' 操作ログファイルのハンドル
    '    'Public gwTrimmerStatus As Short ' ホスト通信ステータス保持

    '    '''' ロギングフラグ　09/09/09  SysParamから移行


    '    Public Const KUGIRI_CHAR As Short = &H9S ' TAB

    '    'Public gbInPattern As Boolean ' 位置補正処理中
    '    'Public gbRotCorrectCancel As Short ' 0:OK, n < 0: 位置補正をキャンセルした or 位置補正エラー
    '    ''Public gfCorrectPosX As Double                          ' トリムポジション補正値X 
    '    'Public gfCorrectPosY As Double                          ' トリムポジション補正値Y
    '    'Public gbPreviousPrbPos As Boolean ' プローブ位置合わせのBP/STAGE座標を記憶している
    '    'Public gsCutTypeName(256) As String ' カットタイプ名テーブル
    '    'Public gtimerCoverTimeUp As Boolean

    '    ''BPリニアリティー補正値
    '    'Public Const cMAXcBPcLINEARITYcNUM As Short = 21


    '    ''''2009/05/29 minato
    '    ''''    LoaderAlarm.bas削除により一時移動
    '    ''''===============================================
    '    '' ''Public iLoaderAlarmKind As Short ' ｱﾗｰﾑ種類(1:全停止異常 2:ｻｲｸﾙ停止 3:軽故障 0:ｱﾗｰﾑ無し)
    '    '' ''Public iLoaderAlarmNum As Short ' 発生中のｱﾗｰﾑ数
    '    '' ''Public strLoaderAlarm() As String ' ｱﾗｰﾑ文字列
    '    '' ''Public strLoaderAlarmInfo() As String ' ｱﾗｰﾑ情報1
    '    '' ''Public strLoaderAlarmExec() As String ' ｱﾗｰﾑ情報2(対策)
    '    ''''===============================================



    '    'Public gbInitialized As Boolean

    '    '----- 分布図用 -----
    '    Public Const MAX_SCALE_NUM As Integer = 999999999           ' ｸﾞﾗﾌ最大値
    '    Public Const MAX_SCALE_RNUM As Integer = 12                 ' ｸﾞﾗﾌ表示抵抗数

    '    Public gDistRegNumLblAry(12) As System.Windows.Forms.Label     ' 分布グラフ抵抗数配列
    '    Public gDistGrpPerLblAry(12) As System.Windows.Forms.Label     ' 分布グラフ%配列
    '    Public gDistShpGrpLblAry(12) As System.Windows.Forms.Label     ' 分布グラフ配列

    '    Public glRegistNum(12) As Integer                            ' 分布グラフ抵抗数
    '    Public glRegistNumIT(12) As Integer                          ' 分布グラフ抵抗数 ｲﾆｼｬﾙﾃｽﾄ
    '    Public glRegistNumFT(12) As Integer                          ' 分布グラフ抵抗数 ﾌｧｲﾅﾙﾃｽﾄ

    '    Public lOkChip As Integer                                   ' OK数
    '    Public lNgChip As Integer                                   ' NG数
    '    Public dblMinIT As Double                                   ' 最小値ｲﾆｼｬﾙ
    '    Public dblMaxIT As Double                                   ' 最大値ｲﾆｼｬﾙ
    '    Public dblMinFT As Double                                   ' 最小値ﾌｧｲﾅﾙ
    '    Public dblMaxFT As Double                                   ' 最大値ﾌｧｲﾅﾙ
    '    '' '' ''Public dblGapIT As Double                                   ' 積算誤差ｲﾆｼｬﾙ
    '    '' '' ''Public dblGapFT As Double                                   ' 積算誤差ﾌｧｲﾅﾙ

    '    Public dblAverage As Double                                 ' 平均値
    '    Public dblDeviationIT As Double                             ' 標準偏差(IT)
    '    Public dblDeviationFT As Double                             ' 標準偏差(FT)

    '    Public dblAverageIT As Double                               ' IT平均値
    '    Public dblAverageFT As Double                               ' FT平均値
    '    Public HEIHOUIT As Double                                   ' 平方偏差
    '    Public HEIHOUFT As Double                                   ' 平方偏差

#End Region

#Region "グローバル変数の定義"
    '    '===========================================================================
    '    '   グローバル変数の定義
    '    '===========================================================================

    '    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '    ''''' 2009/04/13 minato
    '    ''''    TKYでは使用しているグローバル変数
    '    ''''    共通化の為、TKY用としては宣言する方向で

    '    '----- 連続運転用(SL436R用) -----
    '    Public gbFgAutoOperation As Boolean = False                     ' 自動運転フラグ(True:自動運転中, False:自動運転中でない) 
    '    Public gsAutoDataFileFullPath() As String                       ' 連続運転登録データファイル名配列
    '    Public giAutoDataFileNum As Short                               ' 連続運転登録データファイル数
    '    Public giActMode As Short                                       ' 連続運転動作モード(0:ﾏｶﾞｼﾞﾝﾓｰﾄﾞ 1:ﾛｯﾄﾓｰﾄﾞ 2:ｴﾝﾄﾞﾚｽﾓｰﾄﾞ)
    '    Public Const MODE_MAGAZINE As Short = 0                         ' マガジンモード
    '    Public Const MODE_LOT As Short = 1                              ' ロットモード
    '    Public Const MODE_ENDLESS As Short = 2                          ' エンドレスモード
    '    '                                                               ' 切替えモード(1=自動モード, 0=手動モード)
    '    Public Const MODE_MANUAL As Integer = 0                         ' 手動モード
    '    Public Const MODE_AUTO As Integer = 1                           ' 自動モード
    '    Public giErrLoader As Short = 0                                 ' ローダアラーム検出(0:未検出 0以外:エラーコード) ###073

    '    '                                                               ' 以下はシスパラより設定する
    '    Public giOPLDTimeOutFlg As Integer                              ' ローダ通信タイムアウト検出(0=検出無し, 1=検出あり)
    '    Public giOPLDTimeOut As Integer                                 ' ローダ通信タイムアウト時間(msec)
    '    Public giOPVacFlg As Integer                                    ' 手動モード時の載物台吸着アラーム検出(0=検出無し, 1=検出あり)
    '    Public giOPVacTimeOut As Integer                                ' 手動モード時の載物台吸着アラームタイムアウト時間(msec)

    '    Public Const MAXWORK_KND As Integer = 10                        ' プレートデータの基板品種の数
    '    Public giLoaderSpeed As Integer                                 ' ローダ搬送速度
    '    Public giLoaderPositionSetting As Integer                       ' ローダ位置設定選択番号
    '    Public gfBordTableOutPosX(0 To MAXWORK_KND - 1) As Double       ' ローダ基板テーブル排出位置X
    '    Public gfBordTableOutPosY(0 To MAXWORK_KND - 1) As Double       ' ローダ基板テーブル排出位置Y
    '    Public gfBordTableInPosX(0 To MAXWORK_KND - 1) As Double        ' ローダ基板テーブル供給位置X
    '    Public gfBordTableInPosY(0 To MAXWORK_KND - 1) As Double        ' ローダ基板テーブル供給位置Y
    '    Public giNgBoxCount(0 To MAXWORK_KND - 1) As Integer            ' NG排出BOXの収納枚数(基板品種分)   ###089
    '    Public giNgBoxCounter As Integer = 0                            ' NG排出BOXの収納枚数カウンター     ###089

    '    Public giBreakCounter As Integer = 0                            ' 割れ欠け発生の収納枚数カウンター     ###130 
    '    Public giTwoTakeCounter As Integer = 0                          ' ２枚取り発生の収納枚数カウンター     ###130 

    '    Public m_lTrimResult As Integer = cFRS_NORMAL                   ' 基板単位のトリミング結果(SL436R自動運転時のNG排出BOXの収納枚数カウント用) ###089
    '    '                                                               ' cFRS_NORMAL (正常)
    '    '                                                               ' cFRS_TRIM_NG(トリミングNG)
    '    '                                                               ' cFRS_ERR_PTN(パターン認識エラー) ※なし
    Public bFgAutoMode As Boolean = False                           ' ローダ自動モードフラグ

    '    '----- 連続運転用(SL436R用) -----


    '    '    Public Const cMAXcMARKINGcSTRLEN As Short = 18          ' マーキング文字列最大長(byte)
    '    'Public strPlateDataFileFullPath() As String             ' 連続運転登録ﾘｽﾄﾌﾙﾊﾟｽ文字列配列
    '    'Public intPlateDataFileNum As Short                     ' 連続運転登録ﾘｽﾄﾌﾙﾊﾟｽ文字列数
    '    'Public intActMode As Short                              ' 連続運転動作ﾓｰﾄﾞ(0:ﾏｶﾞｼﾞﾝﾓｰﾄﾞ 1:ﾛｯﾄﾓｰﾄﾞ 2:ｴﾝﾄﾞﾚｽﾓｰﾄﾞ)

    '    'Public INTRTM_Ver As String 'INtime Version
    '    'Public LMP_No As String 'LMP No


    '    '' '' ''Public gfX_2IT As Double ' IT標準偏差算出用ワーク
    '    '' '' ''Public gfX_2FT As Double ' FT標準偏差算出用ワーク

    '    Public glITTOTAL As Long                                        ' IT計算対象数 ###138
    '    Public glFTTOTAL As Long                                        ' FT計算対象数 ###138

    '    'Public gbEditPassword As Short ' データ入力時のパスワード要求(0:無 1:有)
    '    Public gITNx() As Double                                        'IT 測定誤差(個々)
    '    Public gFTNx() As Double                                        'FT 測定誤差(個々)

    '    Public gITNx_cnt As Integer                                     'IT 算出用ﾜｰｸ数
    '    Public gITNg_cnt As Integer                                     'IT NG数記録
    '    Public gFTNx_cnt As Integer                                     'FT 算出用ﾜｰｸ数
    '    Public gFTNg_cnt As Integer                                     'FT NG数記録
    '    'Public giXmode As Short
    '    Public gLogMode As Integer                                      'ﾛｷﾞﾝｸﾞﾓｰﾄﾞ(0:しない, 1:INITIAL TEST, 2:FINAL TEST, 3:INITIAL + FINAL) ###150 

    '    Public StepTab_Mode As Short                                    '(0)Step (1)Group
    '    Public StepFGMove As Short                                      '(0)なし　(1)ｽﾃｯﾌﾟｸﾞﾘｯﾄﾞ間移動あり[->]  (2)ｽﾃｯﾌﾟｸﾞﾘｯﾄﾞ間移動あり[<-]
    '    Public StepTitle(2) As Short                                    '(0)入力あり　(1)入力なし

    '    '--ROHM--
    '    Public giLoginPass As Boolean '起動時ﾊﾟｽﾜｰﾄﾞ入力(False)NG (True)OK
    '    'Public gsLoginPassword As String                    'iniﾌｧｲﾙ内のﾊﾟｽﾜｰﾄﾞ
    '    '--ROHM(印刷)--
    '    Public PrnDateR As String '作業日
    '    Public prnSTART_TIME As String '開始時間
    '    Public prnSTOP_TIME As String '終了時間
    '    Public prnPROG_TIME As String '開始～終了までに要した時間
    '    Public prnOPE_TIME As String '稼動時間
    '    Public prnALARM_TIME As String 'ｱﾗｰﾑにより停止した時間
    '    Public prnOPE_RATE As String '稼働率
    '    Public prnMTBF As String '平均故障間隔
    '    Public prnMTTR As String '平均復旧時間
    '    Public prnLOT_NO As String 'ﾄﾘﾐﾝｸﾞﾃﾞｰﾀｼｰｹﾝｽﾅﾝﾊﾞｰ
    '    Public prnQrate As String 'ﾄﾘﾐﾝｸﾞQﾚｰﾄ
    '    Public prnTrim_Speed As String 'ﾄﾘﾐﾝｸﾞｶｯﾄｽﾋﾟｰﾄﾞ
    '    Public prnTrim_OK As Integer '良品ﾁｯﾌﾟ数
    '    Public prnPretest_Lo_Fail As Integer '初期値下限不良のﾁｯﾌﾟ数
    '    Public prnPretest_Hi_Fail As Integer '初期値上限不良のﾁｯﾌﾟ数
    '    Public prnPretest_Open As Integer '初期値ｵｰﾌﾟﾝ不良のﾁｯﾌﾟ数
    '    Public prnCut_NG As Integer 'ﾄﾘﾐﾝｸﾞ時に目標値に達しなかったﾁｯﾌﾟ数
    '    Public prnPretest_NG_Cut_NG As Integer '初期不良
    '    Public prnFinal_test_Lo_Fail As Integer 'ﾄﾘﾐﾝｸﾞ後の下限不良のﾁｯﾌﾟ数
    '    Public prnFinal_test_Hi_Fail As Integer 'ﾄﾘﾐﾝｸﾞ後の上限不良のﾁｯﾌﾟ数
    '    Public prnFinal_test_Open As Integer 'ﾄﾘﾐﾝｸﾞ後にｵｰﾌﾟﾝｴﾗｰとなったﾁｯﾌﾟ数
    '    Public prnYield As String '良品ﾁｯﾌﾟ数÷ﾁｯﾌﾟ数
    '    Public prnYield_Par As Double '上記の%表示
    '    Public prnPdt_Sheet As Integer 'ﾄﾘﾐﾝｸﾞｽﾃｰｼﾞで処理した基板枚数
    '    Public prnLot_Sheet As Integer '装置に投入されたﾛｯﾄ枚数
    '    Public prnLot_NG_Sheet As Integer 'ﾛｯﾄ中の不良基板数
    '    Public prnEdg_Fail As Integer 'ﾛｯﾄ中の認識不良基板枚数
    '    Public prnNominal As Double '目標抵抗値
    '    Public prnTrim_Target As Double '補正後の目標抵抗値
    '    Public prnTrim_Limit As Double 'ﾄﾘﾐﾝｸﾞ目標補正値
    '    Public prnMean_Value As Double 'ﾄﾘﾐﾝｸﾞされたﾁｯﾌﾟの平均抵抗値
    '    Public prn_Par As Double '上記の%表示
    '    Public prnM_R As Double '平均値の誤差
    '    Public prn3S__x As Double 'ﾄﾘﾐﾝｸﾞされたﾁｯﾌﾟの誤差の標準偏差

    '    Public prnSTtime As Double '開始時間(double)
    '    Public prnEDtime As Double '終了時間(double)
    '    Public prnAlmSTtime As Double 'ｱﾗｰﾑ停止開始時間(double)
    '    Public prnAlmEDtime As Double 'ｱﾗｰﾑ停止終了時間(double)
    '    Public prnAlmCnt As Short 'ｱﾗｰﾑ発生回数
    '    Public prnAlmTotaltime As Double 'ｱﾗｰﾑ停止ﾄｰﾀﾙ時間(double)
    '    Public prnChipTotal As Double '1ﾛｯﾄ分の総抵抗数
    '    Public prnTrim_NG As Integer '不良品ﾁｯﾌﾟ数
    '    Public prnTrim_TotalVal As Double 'ﾄﾘﾐﾝｸﾞされたﾁｯﾌﾟの抵抗値(合計)
    '    Public prnTrim_TotalValCnt As Double 'ﾄﾘﾐﾝｸﾞされたﾁｯﾌﾟの抵抗値(計算用合計)
    '    Public prnTrim_TotalValKT As Short 'ﾄﾘﾐﾝｸﾞされたﾁｯﾌﾟの抵抗値(桁)


    '    'Public bPrnDataLoad As Boolean 'ﾃﾞｰﾀﾛｰﾄﾞ(True)初回　(False)2回目以降

    '    Public sIX2LogFilePath As String 'IX2 LOGﾌｧｲﾙﾊﾟｽ名
    '    Public gsESLogFilePath As String 'ES LOGﾌｧｲﾙﾊﾟｽ名

    '    'frmFineAdjust.vbでのみ使用する変数
    '    '   フォーム終了後に値の取得が必要なため、
    '    '   グローバルで変数を設定する。
    Public gCurPlateNo As Integer
    Public gCurBlockNo As Integer
    '    Public gFrmEndStatus As Integer

    '    '----- ログ画面表示用 -----　                                   '###013
    '    Public gDspClsCount As Integer                                  ' ログ画面表示クリア基板枚数
    '    Public gDspCounter As Integer                                   ' ログ画面表示基板枚数カウンタ

    '    '----- 一時停止画面用 -----
    Public gbExitFlg As Boolean                                     '###014
    Public gbTenKeyFlg As Boolean = True                            ' テンキー入力フラグ ###057
    Public gbChkboxHalt As Boolean = True                           ' ADJボタン状態(ON=ADJ ON, OFF=ADJ OFF) ###009
    '    Public gbHaltSW As Boolean = False                              ' HALT SW状態退避 ###255
    Public gObjADJ As Object = Nothing                              ' 一時停止画面オブジェクト ###053

    '    '----- EXTOUT LED制御ビット -----                               '###061
    '    Public glLedBit As Long                                         ' LED制御ビット(EXTOUT) 

    '    '----- GP-IB制御 -----
    '    Public bGpib2Flg As Integer = 0                                 ' GP-IB制御(汎用)フラグ(0=制御なし, 1=制御あり) ###229

#End Region

    '========================================================================================
    '   ジョグ操作用変数定義(ＴＸ/ＴＹティーチング他共通)
    '========================================================================================
#Region "ジョグ操作用変数定義"
    '-------------------------------------------------------------------------------
    '   ジョグ操作用定義
    '-------------------------------------------------------------------------------
    '    Public giCurrentNo As Integer                               ' 処理中の行番号(グリッド表示用)

    '    '----- JOG操作用パラメータ形式定義(OcxJOGを使用しない場合) -----
    Public Structure JOG_PARAM
        Dim Md As Short                                         ' 処理モード(0:XYﾃｰﾌﾞﾙ移動, 1:BP移動, 2:キー入力待ちモード)
        Dim Md2 As Short                                        ' 入力モード(0:画面ﾎﾞﾀﾝ入力, 1:ｺﾝｿｰﾙ入力)
        Dim Opt As UShort                                       ' オプション(キーの有効(1)/無効(0)指定)
        '                                                       '  BIT0:STARTキー
        '                                                       '  BIT1:RESETキー
        '                                                       '  BIT2:Zキー
        '                                                       '  BIT3:
        '                                                       '  BIT4:未使用
        '                                                       '  BIT5:HALTキー
        '                                                       '  BIT6:未使用
        '                                                       '  BIT7-15:未使用
        Dim Flg As Short                                        ' 親画面のOK/Cancelﾎﾞﾀﾝ押下ﾌﾗｸﾞ(cFRS_ERR_ADV, cFRS_ERR_RST)
        Dim PosX As Double                                      ' BP or ﾃｰﾌﾞﾙ X位置
        Dim PosY As Double                                      ' BP or ﾃｰﾌﾞﾙ Y位置
        Dim BpOffX As Double                                    ' BPｵﾌｾｯﾄX 
        Dim BpOffY As Double                                    ' BPｵﾌｾｯﾄY
        Dim BszX As Double                                      ' ﾌﾞﾛｯｸｻｲｽﾞX 
        Dim BszY As Double                                      ' ﾌﾞﾛｯｸｻｲｽﾞY
        Dim TextX As Object                                     ' BP or ﾃｰﾌﾞﾙ X位置表示用ﾃｷｽﾄﾎﾞｯｸｽ
        Dim TextY As Object                                     ' BP or ﾃｰﾌﾞﾙ Y位置表示用ﾃｷｽﾄﾎﾞｯｸｽ
        Dim cgX As Double                                       ' 移動量X 
        Dim cgY As Double                                       ' 移動量Y
        Dim bZ As Boolean                                       ' Zキー  (True:ON, False:OFF)

        Dim BtnHI As Object                                     ' HIボタン
        Dim BtnZ As Object                                      ' Zボタン
        Dim BtnSTART As Object                                  ' STARTボタン
        Dim BtnHALT As Object                                   ' HALTボタン
        Dim BtnRESET As Object                                  ' RESETボタン
        Dim CurrentNo As Integer                                ' 処理中の行番号(グリッド表示用)
    End Structure

    '    '----- ZINPSTS関数(コンソール入力)戻値 -----
    Public Const CONSOLE_SW_START As UShort = &H1           ' bit 0(01)  : START       0/1=未動作/動作
    Public Const CONSOLE_SW_RESET As UShort = &H2           ' bit 1(02)  : RESET       0/1=未動作/動作
    Public Const CONSOLE_SW_ZSW As UShort = &H4             ' bit 2(04)  : Z_ON/OFF_SW 0/1=未動作/動作
    '    Public Const CONSOLE_SW_ZDOWN As UShort = &H8           ' bit 3(08)  : Z_DOWN      1=状態センス
    '    Public Const CONSOLE_SW_ZUP As UShort = &H10            ' bit 4(10)  : Z_UP        1=状態センス
    Public Const CONSOLE_SW_HALT As UShort = &H20           ' bit 5(20)  : HALT        0/1=未動作/動作

    '    '----- コンソールキーSW -----
    '    'Public Const cBIT_ADV As UShort = &H1US                 ' START(ADV)キー
    '    'Public Const cBIT_HALT As UShort = &H2US                ' HALTキー
    '    'Public Const cBIT_RESET As UShort = &H8US               ' RESETキー
    '    'Public Const cBIT_Z As UShort = &H20US                  ' Zキー
    Public Const cBIT_HI As UShort = &H100US                ' HIキー

    '    '----- 処理モード定義 -----
    Public Const MODE_STG As Integer = 0                    ' XYテーブルモード
    Public Const MODE_BP As Integer = 1                     ' BPモード
    Public Const MODE_KEY As Integer = 2                    ' キー入力待ちモード

    '    '----- プローブモード/サブモード定義 -----
    '    'Public Const MODE_STG      As Integer = 0              ' XYテーブルモード
    '    'Public Const MODE_BP       As Integer = 1              ' BPモード
    '    Public Const MODE_Z As Integer = 2                      ' Zﾓｰﾄﾞ
    '    Public Const MODE_TTA As Integer = 3                    ' θﾓｰﾄﾞ
    '    Public Const MODE_Z2 As Integer = 4                     ' Z2ﾓｰﾄﾞ

    '    Public Const MODE_PRB As Integer = 10                   ' 接触位置確認モード
    '    Public Const MODE_RECOG As Integer = 20                 ' θ補正手動位置合せモード
    '    ' ※アプリモードは「トリミング中」
    '    Public Const MODE_POSOFS As Integer = 21                ' 補正ポジションオフセット調整モード
    '    ' ※アプリモードは「パターン登録(θ補正)」

    '    '----- 入力モード -----
    Public Const MD2_BUTN As Integer = 0                    ' 画面ボタン入力
    '    Public Const MD2_CONS As Integer = 1                    ' コンソール入力
    '    Public Const MD2_BOTH As Integer = 2                    ' 両方

    '    '----- ピッチ最大値/最小値 -----
    Public Const cPT_LO As Double = 0.001                   ' ﾋﾟｯﾁ最小値(mm)
    Public Const cPT_HI As Double = 0.1                     ' ﾋﾟｯﾁ最大値(mm)
    Public Const cHPT_LO As Double = 0.01                   ' HIGHﾋﾟｯﾁ最小値(mm)
    Public Const cHPT_HI As Double = 5.0#                   ' HIGHﾋﾟｯﾁ最大値(mm)
    Public Const cPAU_LO As Double = 0.05                   ' ポーズ最小値(sec)
    Public Const cPAU_HI As Double = 1.0#                   ' ポーズ最大値(sec)

    '    '----- 添え字 -----
    Public Const IDX_PIT As Short = 0                       ' ﾋﾟｯﾁ
    Public Const IDX_HPT As Short = 1                       ' HIGHﾋﾟｯﾁ
    Public Const IDX_PAU As Short = 2                       ' ポーズ

    '    '----- その他 -----
    '    'Private dblTchMoval(3) As Double                           ' ﾋﾟｯﾁ退避域(0=ﾋﾟｯﾁ, 1=HIGHﾋﾟｯﾁ, 2=Pause Time(Sec))
    Private InpKey As UShort                                    ' ｺﾝｿｰﾙｷｰ入力域 
    Private cin As UShort                                       ' ｺﾝｿｰﾙ入力値
    Private bZ As Boolean                                       ' Zキー 退避域 (True:ON, False:OFF)
    Private bHI As Boolean                                      ' HIキー(True:ON, False:OFF)

    Private mPIT As Double                                      ' 移動ﾋﾟｯﾁ
    Private X As Double                                         ' 移動ﾋﾟｯﾁ(X)
    Private Y As Double                                         ' 移動ﾋﾟｯﾁ(Y)
    '    Private NOWXP As Double                                     ' BP現在値X(ｸﾛｽﾗｲﾝ補正用)
    '    Private NOWYP As Double                                     ' BP現在値Y(ｸﾛｽﾗｲﾝ補正用)
    Private mvx As Double                                       ' BP/ﾃｰﾌﾞﾙ等の位置X
    Private mvy As Double                                       ' BP/ﾃｰﾌﾞﾙ等の位置Y
    Private mvxBk As Double                                     ' BP/ﾃｰﾌﾞﾙ等の位置X(退避用)
    Private mvyBk As Double                                     ' BP/ﾃｰﾌﾞﾙ等の位置Y(退避用)
#End Region

    '    '========================================================================================
    '    '   ＪＯＧ操作画面処理用共通関数
    '    '========================================================================================
#Region "初期設定処理"
    '''=========================================================================
    '''<summary>初期設定処理</summary>
    '''<param name="stJOG">       (INP)JOG操作用パラメータ</param>
    '''<param name="TBarLowPitch">(I/O)スライダー1(Lowﾋﾟｯﾁ)</param>
    '''<param name="TBarHiPitch"> (I/O)スライダー2(HIGHﾋﾟｯﾁ)</param>
    '''<param name="TBarPause">   (I/O)スライダー3(Pause Time)</param>
    '''<param name="LblTchMoval0">(I/O)目盛1(Low Pich Label)</param>
    '''<param name="LblTchMoval1">(I/O)目盛2(Lowﾋﾟｯﾁ Label)</param>
    '''<param name="LblTchMoval2">(I/O)目盛3(HIGHﾋﾟｯﾁ Label)</param>
    '''<param name="dblTchMoval"> (I/O)ﾋﾟｯﾁ退避域(0=ﾋﾟｯﾁ, 1=HIGHﾋﾟｯﾁ, 2=Pause Time)</param>
    '''=========================================================================
    Public Sub JogEzInit(ByVal stJOG As JOG_PARAM, _
                         ByRef TBarLowPitch As System.Windows.Forms.TrackBar, _
                         ByRef TBarHiPitch As System.Windows.Forms.TrackBar, _
                         ByRef TBarPause As System.Windows.Forms.TrackBar, _
                         ByRef LblTchMoval0 As System.Windows.Forms.Label, _
                         ByRef LblTchMoval1 As System.Windows.Forms.Label, _
                         ByRef LblTchMoval2 As System.Windows.Forms.Label, _
                         ByRef dblTchMoval() As Double)

        Dim strMSG As String

        Try
            ' 移動ピッチスライダー初期設定
            If (stJOG.Md = MODE_BP) Then                            ' モード = 1(BP移動) ?
                dblTchMoval(IDX_PIT) = gSysPrm.stSYP.gBpPIT         ' BP用ﾋﾟｯﾁ設定
                dblTchMoval(IDX_HPT) = gSysPrm.stSYP.gBpHighPIT
                dblTchMoval(IDX_PAU) = gSysPrm.stSYP.gPitPause
            Else
                dblTchMoval(IDX_PIT) = gSysPrm.stSYP.gPIT           ' XYテーブル用ﾋﾟｯﾁ設定
                dblTchMoval(IDX_HPT) = gSysPrm.stSYP.gStageHighPIT
                dblTchMoval(IDX_PAU) = gSysPrm.stSYP.gPitPause
            End If
            Call XyzBpMovingPitchInit(TBarLowPitch, TBarHiPitch, TBarPause, _
                                      LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval)

            Call Form1.System1.SetSysParam(gSysPrm)                 ' システムパラメータの設定(OcxSystem用)

            InpKey = 0

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "Globals.JogEzInit() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "BP/XYテーブルのJOG操作(Do Loopなし)"
    '''=========================================================================
    '''<summary>BP/XYテーブルのJOG操作 ###047</summary>
    '''<param name="stJOG">       (INP)JOG操作用パラメータ</param>
    '''<param name="TBarLowPitch">(I/O)スライダー1(Lowﾋﾟｯﾁ)</param>
    '''<param name="TBarHiPitch"> (I/O)スライダー2(HIGHﾋﾟｯﾁ)</param>
    '''<param name="TBarPause">   (I/O)スライダー3(Pause Time)</param>
    '''<param name="LblTchMoval0">(I/O)目盛1(Low Pich Label)</param>
    '''<param name="LblTchMoval1">(I/O)目盛2(Lowﾋﾟｯﾁ Label)</param>
    '''<param name="LblTchMoval2">(I/O)目盛3(HIGHﾋﾟｯﾁ Label)</param>
    '''<param name="dblTchMoval"> (I/O)ﾋﾟｯﾁ退避域(0=ﾋﾟｯﾁ, 1=HIGHﾋﾟｯﾁ, 2=Pause Time)</param>
    '''<returns>cFRS_ERR_ADV = OK(STARTｷｰ) 
    '''         cFRS_ERR_RST = Cancel(RESETｷｰ)
    '''         cFRS_ERR_HLT = HALTｷｰ
    '''         -1以下       = エラー</returns>
    ''' <remarks>JogEzInit関数をCall済であること</remarks>
    '''=========================================================================
    Public Function JogEzMove_Ex(ByRef stJOG As JOG_PARAM, ByVal SysPrm As DllSysprm.SYSPARAM_PARAM, _
                         ByRef TBarLowPitch As System.Windows.Forms.TrackBar, _
                         ByRef TBarHiPitch As System.Windows.Forms.TrackBar, _
                         ByRef TBarPause As System.Windows.Forms.TrackBar, _
                         ByRef LblTchMoval0 As System.Windows.Forms.Label, _
                         ByRef LblTchMoval1 As System.Windows.Forms.Label, _
                         ByRef LblTchMoval2 As System.Windows.Forms.Label, _
                         ByRef dblTchMoval() As Double) As Integer

        Dim strMSG As String
        Dim r As Short

        Try
            '---------------------------------------------------------------------------
            '   初期処理
            '---------------------------------------------------------------------------
            X = 0.0 : Y = 0.0                                           ' 移動ﾋﾟｯﾁX,Y
            mvx = stJOG.PosX : mvy = stJOG.PosY                         ' BP or ﾃｰﾌﾞﾙ位置X,Y
            mvxBk = stJOG.PosX : mvyBk = stJOG.PosY
            ' キャリブレーション実行/カット位置補正【外部カメラ】時 ※相対座標を表示するためクリアしない
            ' トリミング時の一時停止画面もクリアしない
            If (giAppMode <> APP_MODE_CARIB_REC) And (giAppMode <> APP_MODE_CUTREVIDE) And _
               (giAppMode <> APP_MODE_FINEADJ) Then                     '###088
                '(giAppMode <> APP_MODE_TRIM) Then                      '###088
                stJOG.cgX = 0.0# : stJOG.cgY = 0.0#                     ' 移動量X,Y
            End If

            'If (giAppMode = APP_MODE_TRIM) Then                        '###088
            If (giAppMode = APP_MODE_FINEADJ) Then                      '###088
                mvx = stJOG.cgX - stJOG.BpOffX : mvy = stJOG.cgY - stJOG.BpOffY
                mvxBk = mvx : mvyBk = mvy
            End If

            Call Init_Proc(stJOG, TBarLowPitch, TBarHiPitch, TBarPause, LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval)

            ' 現在の位置を表示する(ﾃｷｽﾄﾎﾞｯｸｽの背景色を処理中(黄色)に設定する)
            Call DispPosition(stJOG, 1)
            'Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
            'Call Me.Focus()                                            ' フォーカスを設定する(テンキー入力のため)
            ''                                                          ' KeyPreviewプロパティをTrueにすると全てのキーイベントをまずフォームが受け取るようになる。
            '---------------------------------------------------------------------------
            '   コンソールボタン又はコンソールキーからのキー入力処理を行う
            '---------------------------------------------------------------------------
            ' システムエラーチェック
            r = Form1.System1.SysErrChk_ForVBNET(giAppMode)
            If (r <> cFRS_NORMAL) Then Return (r)

            ' メッセージポンプ
            System.Windows.Forms.Application.DoEvents()

            '----- ###232↓ -----
            '' 補正クロスライン表示処理(BP移動モードでTeach時)
            'If (stJOG.Md = MODE_BP) Then                                ' モード = 1(BP移動) ?
            '    NOWXP = 0.0 : NOWYP = 0.0
            '    If (gSysPrm.stCRL.giDspFlg = 1) Then                    ' 補正クロスライン表示 ?
            '        If (gSysPrm.stCRL.giDspFlg = 1) And _
            '           (giAppMode = APP_MODE_TEACH) Then                ' 補正クロスライン表示 ?
            '            Call ZGETBPPOS(NOWXP, NOWYP)                    ' BP現在位置取得
            '            gstCLC.x = NOWXP                                ' BP位置X(mm)
            '            gstCLC.y = NOWYP                                ' BP位置Y(mm)
            '            Call CrossLineCorrect(gstCLC)                   ' 補正クロスライン表示
            '        End If
            '    End If
            'End If
            '----- ###232↑ -----

            ' コンソールボタン又はコンソールキーからのキー入力
            Call ReadConsoleSw(stJOG, cin)                              ' キー入力

            '-----------------------------------------------------------------------
            '   入力キーチェック
            '-----------------------------------------------------------------------
            If (cin And CONSOLE_SW_RESET) Then                          ' RESET SW ?
                ' RESET SW押下時
                If (stJOG.Opt And CONSOLE_SW_RESET) Then                ' RESETキー有効 ?
                    Return (cFRS_ERR_RST)                               ' Return値 = Cancel(RESETｷｰ)
                End If

                ' HALT SW押下時
            ElseIf (cin And CONSOLE_SW_HALT) Then                       ' HALT SW ?
                If (stJOG.Opt And CONSOLE_SW_HALT) Then                 ' オプション(0:HALTキー無効, 1:HALTキー有効)
                    r = cFRS_ERR_HALT                                   ' Return値 = HALTｷｰ
                    GoTo STP_END
                End If

                ' START SW押下時
            ElseIf (cin And CONSOLE_SW_START) Then                      ' START SW ?
                If (stJOG.Opt And CONSOLE_SW_START) Then                ' STARTキー有効 ?
                    r = cFRS_ERR_START                                  ' Return値 = OK(STARTｷｰ) 
                    GoTo STP_END
                End If

                ' Z SWがONからOFF(又はOFFからON)に切替わった時
            ElseIf (stJOG.bZ <> bZ) Then
                If (stJOG.Opt And CONSOLE_SW_ZSW) Then                  ' Zキー有効 ?
                    r = cFRS_ERR_Z                                      ' Return値 = ZｷｰON/OFF
                    stJOG.bZ = bZ                                       ' ON/OFF
                    GoTo STP_END
                End If

                ' 矢印SW押下時
            ElseIf cin And &H1E00US Then                                ' 矢印SW
                '「キー入力待ちモード」なら何もしない
                If (stJOG.Md = MODE_KEY) Then

                Else
                    If cin And &H100US Then                             ' HI SW ? 
                        mPIT = dblTchMoval(IDX_HPT)                     ' mPIT = 移動高速ﾋﾟｯﾁ
                    Else
                        mPIT = dblTchMoval(IDX_PIT)                     ' mPIT = 移動通常ﾋﾟｯﾁ
                    End If

                    ' XYテーブル絶対値移動(ソフトリミットチェック有り)
                    r = cFRS_NORMAL
                    If (stJOG.Md = MODE_STG) Then                       ' モード = XYテーブル移動 ?
                        ' XYテーブル絶対値移動
                        r = Sub_XYtableMove(SysPrm, Form1.System1, Form1.Utility1, stJOG)
                        If (r <> cFRS_NORMAL) Then                      ' ｴﾗｰ ?
                            If (Form1.System1.IsSoftLimitXY(r) = False) Then
                                Return (r)                              ' ｿﾌﾄﾘﾐｯﾄｴﾗｰ以外はｴﾗｰﾘﾀｰﾝ
                            End If
                        End If

                        '  モード = BP移動の場合
                    ElseIf (stJOG.Md = MODE_BP) Then
                        ' BP絶対値移動
                        r = Sub_BPmove(SysPrm, Form1.System1, Form1.Utility1, stJOG)
                        If (r <> cFRS_NORMAL) Then                      ' BP移動エラー ?
                            If (Form1.System1.IsSoftLimitBP(r) = False) Then
                                Return (r)                              ' ｿﾌﾄﾘﾐｯﾄｴﾗｰ以外はｴﾗｰﾘﾀｰﾝ
                            End If
                        End If
                    End If

                    ' ソフトリミットエラーの場合は HI SW以外はOFFする
                    If (r <> cFRS_NORMAL) Then                          ' ｴﾗｰ ?
                        If (stJOG.BtnHI.BackColor = System.Drawing.Color.Yellow) Then
                            InpKey = cBIT_HI                            ' HI SW ON
                        Else
                            InpKey = 0                                  ' HI SW以外はOFF
                        End If
                        r = cFRS_NORMAL                                 ' Retuen値 = 正常 ###143 
                    End If

                    ' 現在の位置を表示する
                    Call DispPosition(stJOG, 1)
                    'Call Form1.System1.WAIT(SysPrm.stSYP.gPitPause)    ' Wait(sec)'###251
                    Call Form1.System1.WAIT(dblTchMoval(IDX_PAU))       ' Wait(sec)'###251
                End If

            End If

            '---------------------------------------------------------------------------
            '   終了処理
            '---------------------------------------------------------------------------
STP_END:
            'stJOG.PosX = mvx                                            ' 位置X,Y更新
            'stJOG.PosY = mvy
            Return (r)                                                  ' Return値設定 

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "Globals.JogEzMove_Ex() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Return値 = 例外エラー 
        End Try
    End Function
#End Region

    '#Region "BP/XYテーブルのJOG操作"
    '    '''=========================================================================
    '    '''<summary>BP/XYテーブルのJOG操作</summary>
    '    '''<param name="stJOG">       (INP)JOG操作用パラメータ</param>
    '    '''<param name="TBarLowPitch">(I/O)スライダー1(Lowﾋﾟｯﾁ)</param>
    '    '''<param name="TBarHiPitch"> (I/O)スライダー2(HIGHﾋﾟｯﾁ)</param>
    '    '''<param name="TBarPause">   (I/O)スライダー3(Pause Time)</param>
    '    '''<param name="LblTchMoval0">(I/O)目盛1(Low Pich Label)</param>
    '    '''<param name="LblTchMoval1">(I/O)目盛2(Lowﾋﾟｯﾁ Label)</param>
    '    '''<param name="LblTchMoval2">(I/O)目盛3(HIGHﾋﾟｯﾁ Label)</param>
    '    '''<param name="dblTchMoval"> (I/O)ﾋﾟｯﾁ退避域(0=ﾋﾟｯﾁ, 1=HIGHﾋﾟｯﾁ, 2=Pause Time)</param>
    '    '''<returns>cFRS_ERR_ADV = OK(STARTｷｰ) 
    '    '''         cFRS_ERR_RST = Cancel(RESETｷｰ)
    '    '''         cFRS_ERR_HLT = HALTｷｰ
    '    '''         -1以下       = エラー</returns>
    '    ''' <remarks>JogEzInit関数をCall済であること</remarks>
    '    '''=========================================================================
    '    Public Function JogEzMove(ByRef stJOG As JOG_PARAM, ByVal SysPrm As DllSysprm.SYSPARAM_PARAM, _
    '                         ByRef TBarLowPitch As System.Windows.Forms.TrackBar, _
    '                         ByRef TBarHiPitch As System.Windows.Forms.TrackBar, _
    '                         ByRef TBarPause As System.Windows.Forms.TrackBar, _
    '                         ByRef LblTchMoval0 As System.Windows.Forms.Label, _
    '                         ByRef LblTchMoval1 As System.Windows.Forms.Label, _
    '                         ByRef LblTchMoval2 As System.Windows.Forms.Label, _
    '                         ByRef dblTchMoval() As Double) As Integer

    '        Dim strMSG As String
    '        Dim r As Short

    '        Try
    '            '---------------------------------------------------------------------------
    '            '   初期処理
    '            '---------------------------------------------------------------------------
    '            X = 0.0 : Y = 0.0                                   ' 移動ﾋﾟｯﾁX,Y
    '            mvx = stJOG.PosX : mvy = stJOG.PosY                 ' BP or ﾃｰﾌﾞﾙ位置X,Y
    '            mvxBk = stJOG.PosX : mvyBk = stJOG.PosY
    '            ' キャリブレーション実行/カット位置補正【外部カメラ】時 ※相対座標を表示するためクリアしない
    '            ' トリミング時の一時停止画面もクリアしない
    '            If (giAppMode <> APP_MODE_CARIB_REC) And (giAppMode <> APP_MODE_CUTREVIDE) And _
    '               (giAppMode <> APP_MODE_FINEADJ) Then             '###088
    '                '(giAppMode <> APP_MODE_TRIM) Then              '###088
    '                stJOG.cgX = 0.0# : stJOG.cgY = 0.0#             ' 移動量X,Y
    '            End If
    '            stJOG.Flg = -1
    '            InpKey = 0
    '            Call Init_Proc(stJOG, TBarLowPitch, TBarHiPitch, TBarPause, LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval)

    '            ' 現在の位置を表示する(ﾃｷｽﾄﾎﾞｯｸｽの背景色を処理中(黄色)に設定する)
    '            Call DispPosition(stJOG, 1)
    '            'Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    '            'Call Me.Focus()                                     ' フォーカスを設定する(テンキー入力のため)
    '            ''                                                   ' KeyPreviewプロパティをTrueにすると全てのキーイベントをまずフォームが受け取るようになる。
    '            '---------------------------------------------------------------------------
    '            '   コンソールボタン又はコンソールキーからのキー入力処理を行う
    '            '---------------------------------------------------------------------------
    '            Do
    '                ' システムエラーチェック
    '                r = Form1.System1.SysErrChk_ForVBNET(giAppMode)
    '                If (r <> cFRS_NORMAL) Then GoTo STP_END

    '                ' メッセージポンプ
    '                '  →VB.NETはマルチスレッド対応なので、本来はイベントの開放などでなく、
    '                '    スレッドを生成してコーディングをするのが正しい。
    '                '    スレッドでなくても、最低でタイマーを利用する。
    '                System.Windows.Forms.Application.DoEvents()
    '                System.Threading.Thread.Sleep(10)               ' CPU使用率を下げるためスリープ

    '                '----- ###232↓ -----
    '                '' 補正クロスライン表示処理(BP移動モードでTeach時)
    '                'If (stJOG.Md = MODE_BP) Then                    ' モード = 1(BP移動) ?
    '                '    NOWXP = 0.0 : NOWYP = 0.0
    '                '    If (gSysPrm.stCRL.giDspFlg = 1) Then        ' 補正クロスライン表示 ?
    '                '        If (gSysPrm.stCRL.giDspFlg = 1) And _
    '                '           (giAppMode = APP_MODE_TEACH) Then    ' 補正クロスライン表示 ?
    '                '            Call ZGETBPPOS(NOWXP, NOWYP)        ' BP現在位置取得
    '                '            gstCLC.x = NOWXP                    ' BP位置X(mm)
    '                '            gstCLC.y = NOWYP                    ' BP位置Y(mm)
    '                '            Call CrossLineCorrect(gstCLC)       ' 補正クロスライン表示
    '                '        End If
    '                '    End If
    '                'End If
    '                '----- ###232↑ -----

    '                ' コンソールボタン又はコンソールキーからのキー入力
    '                Call ReadConsoleSw(stJOG, cin)                  ' キー入力

    '                '-----------------------------------------------------------------------
    '                '   入力キーチェック
    '                '-----------------------------------------------------------------------
    '                If (cin And CONSOLE_SW_RESET) Then              ' RESET SW ?
    '                    ' RESET SW押下時
    '                    If (stJOG.Opt And CONSOLE_SW_RESET) Then    ' RESETキー有効 ?
    '                        r = cFRS_ERR_RST                        ' Return値 = Cancel(RESETｷｰ)
    '                        Exit Do
    '                    End If

    '                    ' HALT SW押下時
    '                ElseIf (cin And CONSOLE_SW_HALT) Then           ' HALT SW ?
    '                    If (stJOG.Opt And CONSOLE_SW_HALT) Then     ' オプション(0:HALTキー無効, 1:HALTキー有効)
    '                        r = cFRS_ERR_HALT                       ' Return値 = HALTｷｰ
    '                        Exit Do
    '                    End If

    '                    ' START SW押下時
    '                ElseIf (cin And CONSOLE_SW_START) Then          ' START SW ?
    '                    If (stJOG.Opt And CONSOLE_SW_START) Then    ' STARTキー有効 ?
    '                        'stJOG.PosX = mvx                       ' 位置X,Y更新
    '                        'stJOG.PosY = mvy
    '                        r = cFRS_ERR_START                      ' Return値 = OK(STARTｷｰ) 
    '                        Exit Do
    '                    End If

    '                    ' Z SWがONからOFF(又はOFFからON)に切替わった時
    '                ElseIf (stJOG.bZ <> bZ) Then
    '                    If (stJOG.Opt And CONSOLE_SW_ZSW) Then      ' Zキー有効 ?
    '                        r = cFRS_ERR_Z                          ' Return値 = ZｷｰON/OFF
    '                        stJOG.bZ = bZ                           ' ON/OFF
    '                        Exit Do
    '                    End If

    '                    ' 矢印SW押下時
    '                ElseIf cin And &H1E00US Then                    ' 矢印SW
    '                    '「キー入力待ちモード」なら何もしない
    '                    If (stJOG.Md = MODE_KEY) Then

    '                    Else
    '                        If cin And &H100US Then                     ' HI SW ? 
    '                            mPIT = dblTchMoval(IDX_HPT)             ' mPIT = 移動高速ﾋﾟｯﾁ
    '                        Else
    '                            mPIT = dblTchMoval(IDX_PIT)             ' mPIT = 移動通常ﾋﾟｯﾁ
    '                        End If

    '                        ' XYテーブル絶対値移動(ソフトリミットチェック有り)
    '                        r = cFRS_NORMAL
    '                        If (stJOG.Md = MODE_STG) Then                ' モード = XYテーブル移動 ?
    '                            ' XYテーブル絶対値移動
    '                            r = Sub_XYtableMove(SysPrm, Form1.System1, Form1.Utility1, stJOG)
    '                            If (r <> cFRS_NORMAL) Then              ' ｴﾗｰ ?
    '                                If (Form1.System1.IsSoftLimitXY(r) = False) Then
    '                                    GoTo STP_END                    ' ｿﾌﾄﾘﾐｯﾄｴﾗｰ以外はｴﾗｰﾘﾀｰﾝ
    '                                End If
    '                            End If

    '                            '  モード = BP移動の場合
    '                        ElseIf (stJOG.Md = MODE_BP) Then
    '                            ' BP絶対値移動
    '                            r = Sub_BPmove(SysPrm, Form1.System1, Form1.Utility1, stJOG)
    '                            If (r <> cFRS_NORMAL) Then              ' BP移動エラー ?
    '                                If (Form1.System1.IsSoftLimitBP(r) = False) Then
    '                                    GoTo STP_END                    ' ｿﾌﾄﾘﾐｯﾄｴﾗｰ以外はｴﾗｰﾘﾀｰﾝ
    '                                End If
    '                            End If
    '                        End If

    '                        ' ソフトリミットエラーの場合は HI SW以外はOFFする
    '                        If (r <> cFRS_NORMAL) Then                  ' ｴﾗｰ ?
    '                            If (stJOG.BtnHI.BackColor = System.Drawing.Color.Yellow) Then
    '                                InpKey = cBIT_HI                    ' HI SW ON
    '                            Else
    '                                InpKey = 0                          ' HI SW以外はOFF
    '                            End If
    '                        End If

    '                        ' 現在の位置を表示する
    '                        Call DispPosition(stJOG, 1)
    '                        Call Form1.System1.WAIT(SysPrm.stSYP.gPitPause)    ' Wait(sec)
    '                    End If

    '                End If

    '            Loop While (stJOG.Flg = -1)

    '            '---------------------------------------------------------------------------
    '            '   終了処理
    '            '---------------------------------------------------------------------------
    '            ' 座標表示用ﾃｷｽﾄﾎﾞｯｸｽの背景色を白色に設定する
    '            Call DispPosition(stJOG, 0)

    '            ' 親画面からOK/Cancelﾎﾞﾀﾝ押下 ?
    '            If (stJOG.Flg <> -1) Then
    '                r = stJOG.Flg
    '            End If

    '            ' OK(STARTｷｰ)なら位置X,Y更新
    '            If (r = cFRS_ERR_START) Then                            ' OK(STARTｷｰ) ?
    '                stJOG.PosX = mvx                                    ' 位置X,Y更新
    '                stJOG.PosY = mvy
    '            End If

    'STP_END:
    '            Call ZCONRST()                                          ' ｺﾝｿｰﾙｷｰﾗｯﾁ解除 
    '            Return (r)                                              ' Return値設定 

    '            ' トラップエラー発生時
    '        Catch ex As Exception
    '            strMSG = "Globals.JogEzMove() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '            Return (cERR_TRAP)                                      ' Return値 = 例外エラー 
    '        End Try
    '    End Function
    '#End Region

#Region "初期設定処理"
    '''=========================================================================
    '''<summary>初期設定処理</summary>
    '''<param name="stJOG">       (INP)JOG操作用パラメータ</param>
    '''<param name="TBarLowPitch">(I/O)スライダー1(Lowﾋﾟｯﾁ)</param>
    '''<param name="TBarHiPitch"> (I/O)スライダー2(HIGHﾋﾟｯﾁ)</param>
    '''<param name="TBarPause">   (I/O)スライダー3(Pause Time)</param>
    '''<param name="LblTchMoval0">(I/O)目盛1(Low Pich Label)</param>
    '''<param name="LblTchMoval1">(I/O)目盛2(Lowﾋﾟｯﾁ Label)</param>
    '''<param name="LblTchMoval2">(I/O)目盛3(HIGHﾋﾟｯﾁ Label)</param>
    '''<param name="dblTchMoval"> (I/O)ﾋﾟｯﾁ退避域(0=ﾋﾟｯﾁ, 1=HIGHﾋﾟｯﾁ, 2=Pause Time)</param>
    '''=========================================================================
    Private Sub Init_Proc(ByVal stJOG As JOG_PARAM, _
                         ByRef TBarLowPitch As System.Windows.Forms.TrackBar, _
                         ByRef TBarHiPitch As System.Windows.Forms.TrackBar, _
                         ByRef TBarPause As System.Windows.Forms.TrackBar, _
                         ByRef LblTchMoval0 As System.Windows.Forms.Label, _
                         ByRef LblTchMoval1 As System.Windows.Forms.Label, _
                         ByRef LblTchMoval2 As System.Windows.Forms.Label, _
                         ByRef dblTchMoval() As Double)

        Dim strMSG As String

        Try

            ' 移動ピッチスライダー設定(前回設定した値)
            Call XyzBpMovingPitchInit(TBarLowPitch, TBarHiPitch, TBarPause, _
                                      LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval)

            ' ボタン有効/無効設定
            If (stJOG.Opt And CONSOLE_SW_HALT) Then                 ' HALTキー有効/無効
                stJOG.BtnHALT.Enabled = True
            Else
                stJOG.BtnHALT.Enabled = False
            End If
            If (stJOG.Opt And CONSOLE_SW_START) Then                ' STARTキー有効/無効
                stJOG.BtnSTART.Enabled = True
            Else
                stJOG.BtnSTART.Enabled = False
            End If
            If (stJOG.Opt And CONSOLE_SW_RESET) Then                ' RESETキー有効/無効
                stJOG.BtnRESET.Enabled = True
            Else
                stJOG.BtnRESET.Enabled = False
            End If
            If (stJOG.Opt And CONSOLE_SW_ZSW) Then                  ' Zキー有効/無効
                stJOG.BtnZ.Enabled = True
            Else
                stJOG.BtnZ.Enabled = False
            End If

            ' Zキー/HIキー状態等退避
            bZ = stJOG.bZ                                           ' Zキー退避
            If (bZ = False) Then                                    ' Zボタンの背景色を設定
                stJOG.BtnZ.BackColor = System.Drawing.SystemColors.Control ' 背景色 = 灰色
                stJOG.BtnZ.Text = "Z Off"
            Else
                stJOG.BtnZ.BackColor = System.Drawing.Color.Yellow        ' 背景色 = 黄色
                stJOG.BtnZ.Text = "Z On"
            End If

            If (stJOG.BtnHI.BackColor = System.Drawing.Color.Yellow) Then ' HIキー状態取得
                bHI = True
                InpKey = InpKey Or cBIT_HI                          ' HI SW ON
            Else
                bHI = False
                InpKey = InpKey And Not cBIT_HI                     ' HI SW OFF
            End If

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "Globals.Init_Proc() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "画面ボタン又はコンソールキーからのキー入力"
    '''=========================================================================
    '''<summary>画面ボタン又はコンソールキーからのキー入力</summary>
    '''<param name="stJOG">(INP)JOG操作用パラメータ</param>
    '''<param name="cin">  (OUT)コンソール入力値</param>
    '''=========================================================================
    Private Sub ReadConsoleSw(ByRef stJOG As JOG_PARAM, ByRef cin As UShort)

        Dim r As Integer
        Dim sw As Long
        Dim strMSG As String

        Try
            ' HALTキー入力チェック
            r = HALT_SWCHECK(sw)
            If (sw <> 0) Then                                           ' HALTキー押下 ?
                If (stJOG.Opt And CONSOLE_SW_HALT) Then                 ' HALTキー有効 ?
                    cin = CONSOLE_SW_HALT
                    Exit Sub
                End If
            End If

            ' Zキー入力チェック
            r = Z_SWCHECK(sw)                                           ' Zスイッチの状態をチェックする
            If (sw <> 0) Then                                           ' Zキー押下 ?
                If (stJOG.Opt And CONSOLE_SW_ZSW) Then                  ' Zキー有効 ?
                    Call SubBtnZ_Click(stJOG)
                    Exit Sub
                End If
            End If

            ' START/RESETキー入力チェック
            r = STARTRESET_SWCHECK(False, sw)                           ' START/RESETキー押下チェック(監視なしモード)

            ' コンソール入力値に変換して設定
            If (sw = cFRS_ERR_START) Then                               ' STARTキー押下 ?
                If (stJOG.Opt And CONSOLE_SW_START) Then                ' STARTキー有効 ?
                    cin = CONSOLE_SW_START
                    Exit Sub
                End If
            ElseIf (sw = cFRS_ERR_RST) Then                             ' RESETキー押下 ?
                If (stJOG.Opt And CONSOLE_SW_RESET) Then                ' RESETキー有効 ?
                    cin = CONSOLE_SW_RESET
                    Exit Sub
                End If
                '    ElseIf (sw = CONSOLE_SW_ZSW) Then                          ' Zキー押下 ?
                '        If (stJOG.opt And CONSOLE_SW_ZSW) Then                  ' Zキー有効 ?
                '            cin = CONSOLE_SW_ZSW
                '        End If
            End If

            ' 「画面ボタン入力」
            cin = InpKey                                                ' 画面ボタン入力

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "Globals.ReadConsoleSw() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "座標表示"
    '''=========================================================================
    '''<summary>座標表示</summary>
    '''<param name="stJOG">(INP)JOG操作用パラメータ</param>
    '''<param name="Md">   (INP)0=背景色を白色に設定, 1=背景色を処理中(黄色)に設定</param>
    '''=========================================================================
    Private Sub DispPosition(ByVal stJOG As JOG_PARAM, ByVal MD As Integer)

        Dim xPos As Double = 0.0                    ' ###232
        Dim yPos As Double = 0.0                    ' ###232
        Dim strMSG As String

        Try
            '「キー入力待ちモード」ならNOP
            If (stJOG.Md = MODE_KEY) Then Exit Sub

            ' 補正位置ティーチングならグリッドに表示する
            If (giAppMode = APP_MODE_CUTPOS) Then
                stJOG.TextX.set_TextMatrix(stJOG.CurrentNo, 2, (stJOG.PosX + stJOG.cgX).ToString("0.0000"))
                stJOG.TextX.set_TextMatrix(stJOG.CurrentNo, 3, (stJOG.PosY + stJOG.cgY).ToString("0.0000"))
                Exit Sub

                ' カット位置補正【外部カメラ】ならグリッドに相対座標を表示する
            ElseIf (giAppMode = APP_MODE_CUTREVIDE) Then
                stJOG.TextX.set_TextMatrix(stJOG.CurrentNo, 3, (stJOG.cgX).ToString("0.0000"))    ' ずれ量X
                stJOG.TextX.set_TextMatrix(stJOG.CurrentNo, 4, (stJOG.cgY).ToString("0.0000"))    ' ずれ量Y
                Exit Sub
            End If

            ' テキストボックスに座標を表示する
            If (MD = 0) Then
                ' キャリブレーション実行時は背景色を灰色に設定
                If (giAppMode = APP_MODE_CARIB_REC) Then
                    ' 背景色を灰色に設定
                    stJOG.TextX.BackColor = System.Drawing.SystemColors.Control
                    stJOG.TextY.BackColor = System.Drawing.SystemColors.Control
                Else
                    ' 背景色を白色に設定
                    stJOG.TextX.BackColor = System.Drawing.Color.White
                    stJOG.TextY.BackColor = System.Drawing.Color.White
                End If
            Else
                ' キャリブレーション実行時は相対座標を表示
                If (giAppMode = APP_MODE_CARIB_REC) Then
                    stJOG.TextX.Text = stJOG.cgX.ToString("0.0000")
                    stJOG.TextY.Text = stJOG.cgY.ToString("0.0000")
                Else
                    ' その他のモード時は絶対座標を表示
                    stJOG.TextX.Text = (stJOG.PosX + stJOG.cgX).ToString("0.0000")
                    stJOG.TextY.Text = (stJOG.PosY + stJOG.cgY).ToString("0.0000")
                    '----- ###232↓ -----
                    ' トリミング時の一時停止画面表示中なら補正クロスラインを表示する
                    If (giAppMode = APP_MODE_FINEADJ) Or (giAppMode = APP_MODE_TX) Then
                        'xPos = Double.Parse(stJOG.TextX.Text)
                        'yPos = Double.Parse(stJOG.TextY.Text)
                        Call ZGETBPPOS(xPos, yPos)
                        ObjCrossLine.CrossLineDispXY(xPos, yPos)
                    End If
                    '----- ###232↑ -----
                End If
                ' 背景色を処理中(黄色)に設定
                stJOG.TextX.BackColor = System.Drawing.Color.Yellow
                stJOG.TextY.BackColor = System.Drawing.Color.Yellow
            End If

            stJOG.TextX.Refresh()
            stJOG.TextY.Refresh()

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "Globals.DispPosition() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

    '#Region "ティーチングＳＷ取得"
    '    '''=========================================================================
    '    ''' <summary>ティーチングＳＷ取得</summary>
    '    ''' <param name="SysPrm">(INP)システムパラメータ</param>
    '    ''' <param name="ObjSys">(INP)OcxSystemオブジェク</param>
    '    ''' <returns>0=OFF, 1:ON</returns>
    '    '''=========================================================================
    '    Private Function Z_TEACHSTS(ByVal SysPrm As DllSysprm.SYSPARAM_PARAM, ByVal ObjSys As Object) As Long

    '        Dim r As Integer
    '        Dim strMSG As String

    '        Try
    '            ' データ入力 & ONビットチェック
    '            If (SysPrm.stIOC.giTeachSW = 1) Then                    ' ティーチングSW制御あり ?
    '                r = ObjSys.Inp_And_Check_Bit(SysPrm.stIOC.glTS_In_Adr, SysPrm.stIOC.glTS_In_ON, SysPrm.stIOC.giTS_In_ON_ST)
    '                If (r = 1) Then                                     ' TEACH_SW ON ?
    '                    r = 1                                           ' TEACH_SW ON
    '                Else
    '                    r = 0                                           ' TEACH_SW OFF
    '                End If

    '            Else
    '                r = 1                                               ' TEACH_SW ON
    '            End If
    '            Return (r)

    '            ' トラップエラー発生時
    '        Catch ex As Exception
    '            strMSG = "Globals.Z_TEACHSTS() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '            Return (1)
    '        End Try
    '    End Function
    '#End Region

#Region "BP絶対値移動(ソフトリミットチェック有り)"
    '''=========================================================================
    ''' <summary>BP絶対値移動(ソフトリミットチェック有り)</summary>
    ''' <param name="SysPrm">(INP)システムパラメータ</param>
    ''' <param name="ObjSys">(INP)OcxSystemオブジェク</param>
    ''' <param name="ObjUtl">(INP)OcxUtilityオブジェク</param>
    ''' <param name="stJOG"> (I/O)JOG操作用パラメータ</param>
    ''' <returns>0=正常, 0以外:エラー</returns>
    '''=========================================================================
    Private Function Sub_BPmove(ByVal SysPrm As DllSysprm.SYSPARAM_PARAM, ByVal ObjSys As Object, ByVal ObjUtl As Object, ByRef stJOG As JOG_PARAM) As Integer

        Dim r As Integer
        Dim strMSG As String

        Try
            ' BP移動量の算出(→X,Y)
            mvxBk = mvx                                             ' 現在の位置退避
            mvyBk = mvy
            Call ObjUtl.GetBPmovePitch(cin, X, Y, mPIT, mvx, mvy, SysPrm.stDEV.giBpDirXy)

            ' BP絶対値移動(ソフトリミットチェック有り)
            r = ObjSys.BPMOVE(SysPrm, stJOG.BpOffX, stJOG.BpOffY, stJOG.BszX, stJOG.BszY, mvx, mvy, 1)
            If (r <> cFRS_NORMAL) Then                              ' ｴﾗｰならｴﾗｰﾘﾀｰﾝ(メッセージ表示済み)
                If (ObjSys.IsSoftLimitBP(r) = False) Then
                    GoTo STP_END                                    ' ｿﾌﾄﾘﾐｯﾄｴﾗｰ以外はｴﾗｰﾘﾀｰﾝ
                End If
                mvx = mvxBk                                         ' BPｿﾌﾄﾘﾐｯﾄｴﾗｰ時はBP位置を戻す
                mvy = mvyBk
                GoTo STP_END                                        ' BPｿﾌﾄﾘﾐｯﾄｴﾗｰ
            End If

            stJOG.cgX = stJOG.cgX + (-1 * X)                        ' BP移動量X更新 (※移動量は反転しているので-1を掛ける)
            stJOG.cgY = stJOG.cgY + (-1 * Y)                        ' BP移動量Y更新

STP_END:
            Return (r)

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "Globals.Sub_BPmove() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)
        End Try
    End Function
#End Region

#Region "XYテーブル絶対値移動(ソフトリミットチェック有り)"
    '''=========================================================================
    ''' <summary>XYテーブル絶対値移動(ソフトリミットチェック有り)</summary>
    ''' <param name="SysPrm">(INP)システムパラメータ</param>
    ''' <param name="ObjSys">(INP)OcxSystemオブジェク</param>
    ''' <param name="ObjUtl">(INP)OcxUtilityオブジェク</param>
    ''' <param name="stJOG"> (I/O)JOG操作用パラメータ</param>
    ''' <returns>0=正常, 0以外:エラー</returns>
    '''=========================================================================
    Private Function Sub_XYtableMove(ByVal SysPrm As DllSysprm.SYSPARAM_PARAM, ByVal ObjSys As Object, ByVal ObjUtl As Object, ByRef stJOG As JOG_PARAM) As Integer

        Dim r As Integer
        Dim strMSG As String

        Try
            ' XYテーブル移動量の算出(→X,Y)
            mvxBk = X                                               ' 現在の位置退避
            mvyBk = Y
            Call ObjUtl.GetXYmovePitch(cin, X, Y, mPIT)

            ' XYテーブル絶対値移動(ソフトリミットチェック有り)
            r = ObjSys.XYtableMove(SysPrm, mvx + X, mvy + Y)
            If (r <> cFRS_NORMAL) Then                              ' ｴﾗｰならｴﾗｰﾘﾀｰﾝ(メッセージ表示済み)
                If (ObjSys.IsSoftLimitXY(r) = False) Then
                    GoTo STP_END                                    ' ｿﾌﾄﾘﾐｯﾄｴﾗｰ以外はｴﾗｰﾘﾀｰﾝ
                End If
                X = mvxBk                                           ' ｿﾌﾄﾘﾐｯﾄｴﾗｰ時はX,Y位置を戻す
                Y = mvyBk
                GoTo STP_END                                        ' ｿﾌﾄﾘﾐｯﾄｴﾗｰ
            End If

            mvx = mvx + X                                           ' テーブル位置X,Y更新(絶対座標)
            mvy = mvy + Y
            stJOG.cgX = stJOG.cgX + X                               ' テーブル移動量X,Y更新
            stJOG.cgY = stJOG.cgY + Y

STP_END:
            Return (r)

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "Globals.Sub_XYtableMove() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)
        End Try
    End Function
#End Region

    '    '========================================================================================
    '    '   ボタン押下時処理(ＪＯＧ操作画面)
    '    '========================================================================================
    '#Region "HALTボタン押下時処理"
    '    '''=========================================================================
    '    '''<summary>HALTボタン押下時処理</summary>
    '    '''=========================================================================
    '    Public Sub SubBtnHALT_Click()
    '        InpKey = CONSOLE_SW_HALT
    '    End Sub
    '#End Region

    '#Region "STARTボタン押下時処理"
    '    '''=========================================================================
    '    '''<summary>STARTボタン押下時処理</summary>
    '    '''=========================================================================
    '    Public Sub SubBtnSTART_Click()
    '        InpKey = CONSOLE_SW_START
    '    End Sub
    '#End Region

    '#Region "RESETボタン押下時処理"
    '    '''=========================================================================
    '    '''<summary>RESETボタン押下時処理</summary>
    '    '''=========================================================================
    '    Public Sub SubBtnRESET_Click()
    '        InpKey = CONSOLE_SW_RESET
    '    End Sub
    '#End Region

#Region "Zボタン押下時処理"
    '''=========================================================================
    '''<summary>RESETボタン押下時処理</summary>
    '''<param name="stJOG">(INP)JOG操作用パラメータ</param>
    '''=========================================================================
    Public Sub SubBtnZ_Click(ByVal stJOG As JOG_PARAM)

        Dim strMSG As String

        Try
            If (stJOG.BtnZ.BackColor = System.Drawing.Color.Yellow) Then    ' Z SW ON ?
                stJOG.BtnZ.BackColor = System.Drawing.SystemColors.Control
                stJOG.BtnZ.Text = "Z Off"
                InpKey = InpKey And Not CONSOLE_SW_ZSW                      ' Z SW OFF
                bZ = False                                                  ' Zキー退避域
            Else
                stJOG.BtnZ.BackColor = System.Drawing.Color.Yellow
                stJOG.BtnZ.Text = "Z On"
                InpKey = InpKey Or CONSOLE_SW_ZSW                           ' Z SW ON
                bZ = True                                                   ' Zキー退避域
            End If

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "Globals.SubBtnZ_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "HIボタン押下時処理"
    '''=========================================================================
    '''<summary>HIボタン押下時処理</summary>
    '''<param name="stJOG">(INP)JOG操作用パラメータ</param>
    '''=========================================================================
    Public Sub SubBtnHI_Click(ByVal stJOG As JOG_PARAM)

        ' 背景色を切替える
        If (stJOG.BtnHI.BackColor = System.Drawing.Color.Yellow) Then   ' 背景色 = 黄色 ?
            ' 背景色をデフォルトにする
            stJOG.BtnHI.BackColor = System.Drawing.SystemColors.Control
            InpKey = InpKey And Not cBIT_HI                             ' HI SW OFF
        Else
            ' 背景色を黄色にする
            stJOG.BtnHI.BackColor = System.Drawing.Color.Yellow
            InpKey = InpKey Or cBIT_HI                                  ' HI SW ON
        End If

    End Sub
#End Region

#Region "InpKeyを取得する"
    '''=========================================================================
    '''<summary>InpKeyを取得する</summary>
    '''<param name="IKey">(OUT)InpKey</param>
    '''=========================================================================
    Public Sub GetInpKey(ByRef IKey As UShort) '###057
        IKey = InpKey
    End Sub
#End Region

#Region "InpKeyを設定する"
    '''=========================================================================
    '''<summary>InpKeyを設定する</summary>
    '''<param name="IKey">(INP)InpKey</param>
    '''=========================================================================
    Public Sub PutInpKey(ByVal IKey As UShort) '###057
        InpKey = IKey
    End Sub
#End Region

#Region "矢印ボタン押下時"
    '''=========================================================================
    '''<summary>矢印ボタン押下時</summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub SubBtnJOG_0_MouseDown()
        InpKey = InpKey Or &H1000US                         ' +Y ON
    End Sub
    Public Sub SubBtnJOG_0_MouseUp()
        InpKey = InpKey And Not &H1000US                    ' +Y OFF
    End Sub

    Public Sub SubBtnJOG_1_MouseDown()
        InpKey = InpKey Or &H800US                          ' -Y ON
    End Sub
    Public Sub SubBtnJOG_1_MouseUp()
        InpKey = InpKey And Not &H800US                     ' -Y OFF
    End Sub

    Public Sub SubBtnJOG_2_MouseDown()
        InpKey = InpKey Or &H400US                          ' +X ON
    End Sub
    Public Sub SubBtnJOG_2_MouseUp()
        InpKey = InpKey And Not &H400US                     ' +X OFF
    End Sub

    Public Sub SubBtnJOG_3_MouseDown()
        InpKey = InpKey Or &H200US                          ' -X ON
    End Sub
    Public Sub SubBtnJOG_3_MouseUp()
        InpKey = InpKey And Not &H200US                     ' -X OFF
    End Sub

    Public Sub SubBtnJOG_4_MouseDown()
        InpKey = InpKey Or &HA00US                          ' -X -Y ON
    End Sub
    Public Sub SubBtnJOG_4_MouseUp()
        InpKey = InpKey And Not &HA00US                     ' -X -Y OFF
    End Sub

    Public Sub SubBtnJOG_5_MouseDown()
        InpKey = InpKey Or &HC00US                          ' +X -Y ON
    End Sub
    Public Sub SubBtnJOG_5_MouseUp()
        InpKey = InpKey And Not &HC00US                     ' +X -Y OFF
    End Sub

    Public Sub SubBtnJOG_6_MouseDown()
        InpKey = InpKey Or &H1400US                         ' +X +Y ON
    End Sub
    Public Sub SubBtnJOG_6_MouseUp()
        InpKey = InpKey And Not &H1400US                    ' +X +Y OFF
    End Sub

    Public Sub SubBtnJOG_7_MouseDown()
        InpKey = InpKey Or &H1200US                         ' -X +Y ON
    End Sub
    Public Sub SubBtnJOG_7_MouseUp()
        InpKey = InpKey And Not &H1200US                    ' -X +Y OFF
    End Sub
#End Region

    '========================================================================================
    '   ＪＯＧ操作画面処理用トラックバー処理
    '========================================================================================
#Region "トラックバーのスライダー画面初期値表示"
    '''=========================================================================
    '''<summary>トラックバーのスライダー画面初期値表示</summary>
    '''<param name="TBarLowPitch">(I/O)スライダー1(Lowﾋﾟｯﾁ)</param>
    '''<param name="TBarHiPitch"> (I/O)スライダー2(HIGHﾋﾟｯﾁ)</param>
    '''<param name="TBarPause">   (I/O)スライダー3(Pause Time)</param>
    '''<param name="LblTchMoval0">(I/O)目盛1(Low Pich Label)</param>
    '''<param name="LblTchMoval1">(I/O)目盛2(Lowﾋﾟｯﾁ Label)</param>
    '''<param name="LblTchMoval2">(I/O)目盛3(HIGHﾋﾟｯﾁ Label)</param>
    '''<param name="dblTchMoval"> (I/O)ﾋﾟｯﾁ退避域(0=ﾋﾟｯﾁ, 1=HIGHﾋﾟｯﾁ, 2=Pause Time)</param>
    '''=========================================================================
    Public Sub XyzBpMovingPitchInit(ByRef TBarLowPitch As System.Windows.Forms.TrackBar, _
                                    ByRef TBarHiPitch As System.Windows.Forms.TrackBar, _
                                    ByRef TBarPause As System.Windows.Forms.TrackBar, _
                                    ByRef LblTchMoval0 As System.Windows.Forms.Label, _
                                    ByRef LblTchMoval1 As System.Windows.Forms.Label, _
                                    ByRef LblTchMoval2 As System.Windows.Forms.Label, _
                                    ByRef dblTchMoval() As Double)

        Dim minval As Short

        ' LOWﾋﾟｯﾁがが範囲外なら範囲内に変更する
        If (dblTchMoval(IDX_PIT) < cPT_LO) Then dblTchMoval(IDX_PIT) = cPT_LO
        If (dblTchMoval(IDX_PIT) > cPT_HI) Then dblTchMoval(IDX_PIT) = cPT_HI

        ' LOWﾋﾟｯﾁの目盛を設定する
        If (dblTchMoval(IDX_PIT) < 0.002) Then                          ' 分解能により最小目盛を設定する
            minval = 1                                                  ' 目盛1～
        Else
            minval = 2                                                  ' 目盛2～
        End If

        TBarLowPitch.TickFrequency = dblTchMoval(IDX_PIT) * 1000        ' 0.001mm単位
        TBarLowPitch.Maximum = 100                                      ' 目盛1(or 2)～100(0.001m～0.1mm)
        TBarLowPitch.Minimum = minval
        '###110
        TBarLowPitch.Value = dblTchMoval(IDX_PIT) * 1000        ' 0.001mm単位

        ' HIGHﾋﾟｯﾁがが範囲外なら範囲内に変更する
        If (dblTchMoval(IDX_HPT) < cHPT_LO) Then dblTchMoval(IDX_HPT) = cHPT_LO
        If (dblTchMoval(IDX_HPT) > cHPT_HI) Then dblTchMoval(IDX_HPT) = cHPT_HI

        ' HIGHﾋﾟｯﾁの目盛を設定する
        TBarHiPitch.TickFrequency = dblTchMoval(IDX_HPT) * 100          ' 0.01mm単位
        TBarHiPitch.Maximum = 500                                       ' 目盛1～100(0.01m～5.00mm)
        TBarHiPitch.Minimum = 1
        '###110
        TBarHiPitch.Value = dblTchMoval(IDX_HPT) * 100          ' 0.01mm単位

        ' Pause Timeが範囲外なら範囲内に変更する
        If (dblTchMoval(IDX_PAU) < cPAU_LO) Then dblTchMoval(IDX_PAU) = cPAU_LO
        If (dblTchMoval(IDX_PAU) > cPAU_HI) Then dblTchMoval(IDX_PAU) = cPAU_HI

        ' Pause Timeの目盛を設定する
        TBarPause.TickFrequency = dblTchMoval(IDX_PAU) * 20             ' 0.5秒単位
        TBarPause.Maximum = 20                                          ' 目盛1～20(0.05秒～1.00秒)
        TBarPause.Minimum = 1
        '###110
        TBarPause.Value = dblTchMoval(IDX_PAU) * 20             ' 0.5秒単位

        ' 移動ピッチを表示する
        LblTchMoval0.Text = dblTchMoval(IDX_PIT).ToString("0.0000")
        LblTchMoval1.Text = dblTchMoval(IDX_HPT).ToString("0.0000")
        LblTchMoval2.Text = dblTchMoval(IDX_PAU).ToString("0.0000")

    End Sub
#End Region

#Region "トラックバーのスライダー移動処理"
    '''=========================================================================
    '''<summary>トラックバーのスライダー移動処理</summary>
    '''<param name="Index">       (INP)0=LOWﾋﾟｯﾁ, 1=HIGHﾋﾟｯﾁ, 2=Pause</param>
    '''<param name="TBarLowPitch">(I/O)スライダー1(Lowﾋﾟｯﾁ)</param>
    '''<param name="TBarHiPitch"> (I/O)スライダー2(HIGHﾋﾟｯﾁ)</param>
    '''<param name="TBarPause">   (I/O)スライダー3(Pause Time)</param>
    '''<param name="LblTchMoval0">(I/O)目盛1(Low Pich Label)</param>
    '''<param name="LblTchMoval1">(I/O)目盛2(Lowﾋﾟｯﾁ Label)</param>
    '''<param name="LblTchMoval2">(I/O)目盛3(HIGHﾋﾟｯﾁ Label)</param>
    '''<param name="dblTchMoval"> (I/O)ﾋﾟｯﾁ退避域(0=ﾋﾟｯﾁ, 1=HIGHﾋﾟｯﾁ, 2=Pause Time)</param>
    '''=========================================================================
    Public Sub SetSliderPitch(ByRef Index As Short, _
                              ByRef TBarLowPitch As System.Windows.Forms.TrackBar, _
                              ByRef TBarHiPitch As System.Windows.Forms.TrackBar, _
                              ByRef TBarPause As System.Windows.Forms.TrackBar, _
                              ByRef LblTchMoval0 As System.Windows.Forms.Label, _
                              ByRef LblTchMoval1 As System.Windows.Forms.Label, _
                              ByRef LblTchMoval2 As System.Windows.Forms.Label, _
                              ByRef dblTchMoval() As Double)

        Dim lVal As Integer

        ' BPの移動ピッチ等を設定する
        Select Case Index
            Case IDX_PIT    ' LOWﾋﾟｯﾁ
                lVal = TBarLowPitch.Value                       ' ｽﾗｲﾀﾞ目盛値取得
                dblTchMoval(Index) = 0.001 * lVal               ' LOWﾋﾟｯﾁ値変更
                LblTchMoval0.Text = dblTchMoval(Index).ToString("0.0000")
                LblTchMoval0.Refresh()

            Case IDX_HPT    ' HIGHﾋﾟｯﾁ
                lVal = TBarHiPitch.Value                        ' ｽﾗｲﾀﾞ目盛値取得
                dblTchMoval(Index) = 0.01 * lVal                ' HIGHﾋﾟｯﾁ値変更
                LblTchMoval1.Text = dblTchMoval(Index).ToString("0.0000")
                LblTchMoval1.Refresh()

            Case IDX_PAU    ' Pause Time
                lVal = TBarPause.Value                          ' ｽﾗｲﾀﾞ目盛値取得
                dblTchMoval(Index) = 0.05 * lVal                ' 移動ピッチ間のポーズ値変更
                LblTchMoval2.Text = dblTchMoval(Index).ToString("0.0000")
                LblTchMoval2.Refresh()
        End Select

    End Sub
#End Region

    '========================================================================================
    '   ＪＯＧ操作画面処理用テンキー入力処理
    '========================================================================================
#Region "テンキーダウンサブルーチン"
    '''=========================================================================
    '''<summary>テンキーダウンサブルーチン</summary>
    ''' <param name="KeyCode">(INP)キーコード</param>
    '''=========================================================================
    Public Sub Sub_10KeyDown(ByVal KeyCode As Short)

        Dim strMSG As String

        Try
            ' Num Lock版
            Select Case (KeyCode)
                Case System.Windows.Forms.Keys.NumPad2                      ' ↓  (KeyCode =  98(&H62)
                    InpKey = InpKey Or &H1000                               ' +Y ON(↓)
                Case System.Windows.Forms.Keys.NumPad8                      ' ↑  (KeyCode = 104(&H68)
                    InpKey = InpKey Or &H800                                ' -Y ON(↑)
                Case System.Windows.Forms.Keys.NumPad4                      ' ←  (KeyCode = 100(&H64)
                    InpKey = InpKey Or &H400                                ' +X ON(←)
                Case System.Windows.Forms.Keys.NumPad6                      ' →  (KeyCode = 102(&H66)
                    InpKey = InpKey Or &H200                                ' -X ON(→)
                Case System.Windows.Forms.Keys.NumPad9                      ' PgUp(KeyCode = 105(&H69)
                    InpKey = InpKey Or &HA00                                ' -X -Y ON
                Case System.Windows.Forms.Keys.NumPad7                      ' Home(KeyCode = 103(&H67))
                    InpKey = InpKey Or &HC00                                ' +X -Y ON
                Case System.Windows.Forms.Keys.NumPad1                      ' End(KeyCode =   97(&H61)
                    InpKey = InpKey Or &H1400                               ' +X +Y ON
                Case System.Windows.Forms.Keys.NumPad3                      ' PgDn(KeyCode =  99(&H63)
                    InpKey = InpKey Or &H1200                               ' -X +Y ON
                Case System.Windows.Forms.Keys.NumPad5                      ' 5ｷｰ (KeyCode = 101(&H65)
                    'Call BtnHI_Click(sender, e)                             ' HIボタン ON/OFF
            End Select

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "Globals.Sub_10KeyDown() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "テンキーアップサブルーチン"
    '''=========================================================================
    '''<summary>テンキーアップサブルーチン</summary>
    ''' <param name="KeyCode">(INP)キーコード</param>
    '''=========================================================================
    Public Sub Sub_10KeyUp(ByVal KeyCode As Short)

        Dim strMSG As String

        Try
            ' Num Lock版
            Select Case (KeyCode)
                Case System.Windows.Forms.Keys.NumPad2                      ' ↓  (KeyCode =  98(&H62)
                    InpKey = InpKey And Not &H1000                          ' +Y OFF
                Case System.Windows.Forms.Keys.NumPad8                      ' ↑  (KeyCode = 104(&H68)
                    InpKey = InpKey And Not &H800                           ' -Y OFF
                Case System.Windows.Forms.Keys.NumPad4                      ' ←  (KeyCode = 100(&H64)
                    InpKey = InpKey And Not &H400                           ' +X OFF
                Case System.Windows.Forms.Keys.NumPad6                      ' →  (KeyCode = 102(&H66)
                    InpKey = InpKey And Not &H200                           ' -X OFF
                Case System.Windows.Forms.Keys.NumPad9                      ' PgUp(KeyCode = 105(&H69)
                    InpKey = InpKey And Not &HA00                           ' -X -Y OFF
                Case System.Windows.Forms.Keys.NumPad7                      ' Home(KeyCode = 103(&H67))
                    InpKey = InpKey And Not &HC00                           ' +X -Y OFF
                Case System.Windows.Forms.Keys.NumPad1                      ' End(KeyCode =   97(&H61)
                    InpKey = InpKey And Not &H1400                          ' +X +Y OFF
                Case System.Windows.Forms.Keys.NumPad3                      ' PgDn(KeyCode =  99(&H63)
                    InpKey = InpKey And Not &H1200                          ' -X +Y OFF
            End Select

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "Globals.Sub_10KeyUp() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "入力キーコードのクリア"
    '''=========================================================================
    ''' <summary>
    ''' 入力キーコードのクリア
    ''' </summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub ClearInpKey()

        Try
            InpKey = 0

            ' トラップエラー発生時
        Catch ex As Exception
            MsgBox("Globals.ClearInpKey() TRAP ERROR = " + ex.Message)
        End Try
    End Sub
#End Region

    '    '===========================================================================
    '    '   グローバルメソッド定義
    '    '===========================================================================
    '#Region "機械系のパラメータ設定"
    '    '''=========================================================================
    '    '''<summary>機械系のパラメータ設定</summary>
    '    '''<remarks></remarks>
    '    '''=========================================================================
    '    Public Sub SetMechanicalParam()

    '        Dim BpSoftLimitX As Integer
    '        Dim BpSoftLimitY As Integer

    '        With gSysPrm.stDEV
    '            ' 基板種別対応
    '            If gSysPrm.stCTM.giSPECIAL = customASAHI And typPlateInfo.strDataName = "2" Then
    '                .gfTrimX = .gfTrimX2                            ' TRIM POSITION X(mm)
    '                .gfTrimY = .gfTrimY2                            ' TRIM POSITION Y(mm)
    '                .gfExCmX = .gfExCmX2                            ' Externla Camera Offset X(mm)
    '                .gfExCmY = .gfExCmY2                            ' Externla Camera Offset Y(mm)
    '                .gfRot_X1 = .gfRot_X2                           ' 回転中心 X
    '                .gfRot_Y1 = .gfRot_Y2                           ' 回転中心 Y
    '                '(2010/11/16)下記処理は不要
    '                'Else
    '                '    gSysPrm.stDEV.gfTrimX = gSysPrm.stDEV.gfTrimX   ' TRIM POSITION X(mm)
    '                '    gSysPrm.stDEV.gfTrimY = gSysPrm.stDEV.gfTrimY   ' TRIM POSITION Y(mm)
    '                '    gSysPrm.stDEV.gfExCmX = gSysPrm.stDEV.gfExCmX   ' Externla Camera Offset X(mm)
    '                '    gSysPrm.stDEV.gfExCmY = gSysPrm.stDEV.gfExCmY   ' Externla Camera Offset Y(mm)
    '                '    gSysPrm.stDEV.gfRot_X1 = gSysPrm.stDEV.gfRot_X1 ' 回転中心 X
    '                '    gSysPrm.stDEV.gfRot_Y1 = gSysPrm.stDEV.gfRot_Y1 ' 回転中心 Y
    '            End If
    '            ''''(2010/11/16) 動作確認後下記コメントは削除
    '            'gStartX = gSysPrm.stDEV.gfTrimX
    '            'gStartY = gSysPrm.stDEV.gfTrimY

    '            'BpSizeからBPのソフトリミット（BPのソフト稼動範囲）を設定
    '            Select Case (.giBpSize)
    '                Case 0
    '                    BpSoftLimitX = 50
    '                    BpSoftLimitY = 50
    '                Case 1
    '                    BpSoftLimitX = 80
    '                    BpSoftLimitY = 80
    '                Case 2
    '                    BpSoftLimitX = 100
    '                    BpSoftLimitY = 60
    '                Case 3
    '                    BpSoftLimitX = 60
    '                    BpSoftLimitY = 100
    '                Case Else
    '                    BpSoftLimitX = 80
    '                    BpSoftLimitY = 80
    '            End Select

    '            '''''2009/07/23 minato
    '            ''''    トリムポジションが変更されているため、
    '            ''''    INTRTM側のシステムパラメータを更新する必要がある。
    '            Call ZSYSPARAM2(.giPrbTyp, .gfSminMaxZ2, .giZPTimeOn, .giZPTimeOff, _
    '                        .giXYtbl, .gfSmaxX, .gfSmaxY, gSysPrm.stIOC.glAbsTime, _
    '                        .gfTrimX, .gfTrimY, BpSoftLimitX, BpSoftLimitY)
    '        End With
    '    End Sub
    '#End Region

    '#Region "Uｶｯﾄ実行結果取得"
    '    '''=========================================================================
    '    '''<summary>Uｶｯﾄ実行結果取得</summary>
    '    '''<param name="rn">(INP) 抵抗番号</param>
    '    '''<param name="s"> (OUT) 実行結果</param>
    '    '''<returns>0=正常, 0以外=エラー</returns>
    '    '''=========================================================================
    '    Public Function RetrieveUCutResult(ByVal rn As Short, ByRef s As String) As Short

    '        Dim cn As Short
    '        Dim n As Short
    '        Dim f As Double
    '        Dim r As Integer

    '        s = ""
    '        RetrieveUCutResult = 0

    '        If gSysPrm.stSPF.giUCutKind = 0 Then
    '            Exit Function
    '        End If

    '        On Error GoTo ErrTrap

    '        For cn = 1 To typResistorInfoArray(rn).intCutCount
    '            s = typResistorInfoArray(rn).ArrCut(cn).strCutType       ' Cut pattern
    '            If s = "H" Then
    '                s = ""
    '                '  Uｶｯﾄ実行結果取得
    '                r = UCUT_RESULT(rn, cn, n, f)
    '                If (r <> 0) Then
    '                    MsgBox("Internal error  X001-" & Str(r), MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, gAppName)
    '                    RetrieveUCutResult = 1
    '                    Exit Function
    '                End If

    '                If n = 255 Then                                 ' 255 はUCUT実行していない場合
    '                    s = Form1.Utility1.sFormat(f, "0.000000", 10 + 7) & " n** "
    '                ElseIf n >= 0 And n <= 19 Then
    '                    n = n + 1
    '                    s = Form1.Utility1.sFormat(f, "0.000000", 10 + 7) & " " & "n" & n.ToString("00") & " "
    '                ElseIf n = 254 Then                             ' パラメータテーブルに該当する抵抗番号が無かった場合
    '                    s = Form1.Utility1.sFormat(f, "0.000000", 10 + 7) & " n** "
    '                Else                                            ' 変な値
    '                    RetrieveUCutResult = 2
    '                    Exit Function
    '                End If
    '            Else
    '                s = ""
    '            End If
    '        Next

    '        Exit Function

    'ErrTrap:
    '        Resume ErrTrap1
    'ErrTrap1:
    '        Dim er As Integer
    '        er = Err.Number
    '        On Error GoTo 0
    '        MsgBox("Internal error X002-" & Str(er), MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, gAppName)
    '    End Function
    '#End Region

    '#Region "ﾍﾞｰｽ抵抗番号よりﾚｼｵﾍﾞｰｽ抵抗番号を取得する"
    '    '''=========================================================================
    '    '''<summary>ﾍﾞｰｽ抵抗番号よりﾚｼｵﾍﾞｰｽ抵抗番号を取得する</summary>
    '    '''<param name="br">(INP) 抵抗番号</param>
    '    '''<returns>0以上=ﾚｼｵﾍﾞｰｽ抵抗番号, -1=なし</returns>
    '    '''=========================================================================
    '    Public Function GetRatio1BaseNum(ByVal br As Short) As Short

    '        Dim n As Short

    '        For n = 1 To gRegistorCnt
    '            ' ベース抵抗 ?
    '            If typResistorInfoArray(n).intResNo = br Then
    '                GetRatio1BaseNum = n
    '                Exit Function
    '            End If
    '        Next
    '        GetRatio1BaseNum = -1

    '    End Function
    '#End Region

    '#Region "グループ数,ブロック数,チップ数(抵抗数),チップサイズを取得する(ＴＸ/ＴＹティーチング用)"
    '    '''=========================================================================
    '    ''' <summary>グループ数,ブロック数,チップ数(抵抗数),チップサイズを取得する</summary>
    '    ''' <param name="AppMode">  (INP)モード</param>
    '    ''' <param name="Gn">       (OUT)グループ数</param>
    '    ''' <param name="RnBn">     (OUT)チップ数(ＴＸティーチング時)または
    '    '''                              ブロック数(ＴＹティーチング時)</param>
    '    ''' <param name="DblChipSz">(OUT)チップサイズ</param>
    '    ''' <returns>0=正常, 0以外=エラー</returns>
    '    '''=========================================================================
    '    Public Function GetChipNumAndSize(ByVal AppMode As Short, ByRef Gn As Short, ByRef RnBn As Short, ByRef DblChipSz As Double) As Short

    '        Dim ChipNum As Short                                        ' チップ数(抵抗数)
    '        Dim ChipSzX As Double                                       ' チップサイズX
    '        Dim ChipSzY As Double                                       ' チップサイズY
    '        Dim strMSG As String

    '        Try
    '            ' 前処理(CHIP/NET共通)
    '            ChipNum = typPlateInfo.intResistCntInGroup              ' チップ数(抵抗数) = 1グループ内(1サーキット内)抵抗数
    '            ChipSzX = typPlateInfo.dblChipSizeXDir                  ' チップサイズX,Y
    '            ChipSzY = typPlateInfo.dblChipSizeYDir

    '            ' プレートデータからグループ数, ブロック数, チップ数(抵抗数), チップサイズを取得する
    '            If (AppMode = APP_MODE_TX) Then
    '                '----- ＴＸティーチング時 -----
    '                ' チップ数(抵抗数)を返す
    '                RnBn = ChipNum                                      ' 1グループ内(1サーキット内)抵抗数をセット
    '                ' グループ数を返す
    '                Gn = typPlateInfo.intGroupCntInBlockXBp             ' ＢＰグループ数(サーキット数)をセット
    '                ' チップサイズを返す
    '                If (typPlateInfo.intResistDir = 0) Then             ' チップ並びはX方向 ?
    '                    DblChipSz = System.Math.Abs(ChipSzX)
    '                Else
    '                    DblChipSz = System.Math.Abs(ChipSzY)
    '                End If

    '            Else
    '                '----- ＴＹティーチング時 -----
    '                ' グループ数を返す
    '                Gn = typPlateInfo.intGroupCntInBlockYStage          ' ブロック内Stageグループ数をセット
    '                ' ブロック数とチップサイズを返す
    '                If (typPlateInfo.intResistDir = 0) Then             ' チップ並びはX方向 ?
    '                    RnBn = typPlateInfo.intBlockCntYDir             ' ブロック数Yをセット
    '                    DblChipSz = System.Math.Abs(ChipSzY)            ' チップサイズYをセット
    '                Else
    '                    RnBn = typPlateInfo.intBlockCntXDir             ' ブロック数Xをセット
    '                    DblChipSz = System.Math.Abs(ChipSzX)            ' チップサイズXをセット
    '                End If
    '            End If

    '            strMSG = "GetChipNumAndSize() Gn=" + Gn.ToString("0") + ", RnBn=" + RnBn.ToString("0") + ", ChipSZ=" + DblChipSz.ToString("0.00000")
    '            Console.WriteLine(strMSG)
    '            Return (cFRS_NORMAL)                                    ' Return値 = 正常

    '            ' トラップエラー発生時
    '        Catch ex As Exception
    '            strMSG = "globals.GetChipNumAndSize() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '            Return (cERR_TRAP)                                      ' Return値 = 例外エラー
    '        End Try
    '    End Function
    '#End Region

    '#Region "特注レシオ時のベース抵抗番号(配列の添字)を返す"
    '    '''=========================================================================
    '    '''<summary>特注レシオ時のベース抵抗番号(配列の添字)を返す</summary>
    '    '''<param name="rr">(INP) 抵抗番号</param> 
    '    '''<param name="br">(OUT) ベース抵抗番号(配列の添字)</param>  
    '    '''<remarks>日立殿向け特注レシオ機能(TKYより移植)</remarks>
    '    '''=========================================================================
    '    Public Sub GetRatio3Br(ByRef rr As Short, ByRef br As Short)

    '        Dim i As Short
    '        Dim wRn As Short
    '        Dim wGn As Short
    '        Dim wBr As Short
    '        Dim wBr2 As Short

    '        ' 特注レシオモード(3～9)でなければ通常のベース抵抗番号を返す
    '        wRn = typResistorInfoArray(rr).intResNo                         ' 抵抗番号
    '        wGn = typResistorInfoArray(rr).intTargetValType                 ' 目標値種別（0:絶対値, 1:レシオ、2：計算式, 3～9:ｸﾞﾙｰﾌﾟ番号）
    '        wBr = GetRatio1BaseNum(typResistorInfoArray(rr).intBaseResNo)   ' ベース抵抗番号(添字)
    '        wBr2 = -1
    '        If (wGn < 3) Or (wGn > 9) Then                                  ' 特注レシオモード(3～9)でない ? 
    '            GoTo STP_END
    '        End If

    '        ' 特注レシオなら相手ｸﾞﾙｰﾌﾟ番号を検索する
    '        For i = 1 To gRegistorCnt                                       ' 抵抗数分繰り返す
    '            If (wRn <> typResistorInfoArray(i).intResNo) Then           ' 抵抗番号=自分自身はSKIP
    '                If (wGn = typResistorInfoArray(i).intTargetValType) Then            ' 相手ｸﾞﾙｰﾌﾟ番号 ?
    '                    wBr2 = GetRatio1BaseNum(typResistorInfoArray(i).intBaseResNo)   ' ベース抵抗番号(添字)
    '                    Exit For
    '                End If
    '            End If
    '        Next i

    '        ' ベース抵抗のFT値の大きい方をベース抵抗番号とする
    '        If (wBr2 < 0) Then GoTo STP_END '                               ' 相手ｸﾞﾙｰﾌﾟ番号が見つからなかった ?
    '        If (gfFinalTest(wBr2) > gfFinalTest(wBr)) Then                  ' 相手のFT値が大きい ?
    '            wBr = wBr2
    '        End If

    'STP_END:
    '        'br = wBr                                                       ' ベース抵抗番号を返す
    '        br = wBr - 1                                                    ' ベース抵抗番号を返す ###244

    '    End Sub
    '#End Region

    '#Region "レシオ(計算式)時のベース抵抗番号(配列の添字)を返す"
    '    '''=========================================================================
    '    '''<summary>レシオ(計算式)時のベース抵抗番号から抵抗データの配列の添字を返す###123</summary>
    '    '''<param name="br">(INP)ベース抵抗番号(配列の添字)</param> 
    '    '''<param name="rr">(OUT)抵抗データの配列の添字(1 ORG)</param> 
    '    '''<remarks></remarks>
    '    '''=========================================================================
    '    Public Sub GetRatio2Rn(ByVal br As Short, ByRef rr As Short)

    '        Dim Rn As Short

    '        ' ベース抵抗番号を検索する
    '        For Rn = 1 To gRegistorCnt                                      ' 抵抗数分繰り返す
    '            If (typResistorInfoArray(Rn).intBaseResNo = br) Then
    '                rr = Rn
    '                Exit Sub
    '            End If
    '        Next Rn

    '    End Sub
    '#End Region

    '#Region "Z/Z2移動(ON/OFF) "
    '    '''=========================================================================
    '    '''<summary>Z/Z2移動(ON/OFF) </summary>
    '    '''<param name="MD">  (INP)ﾓｰﾄﾞ(0 = OFF位置移動, 1 = ON位置移動)</param> 
    '    '''<param name="Z2ON">(INP)Z2 ON位置(OPTION)</param>  
    '    '''<remarks>0=正常, 0以外=エラー</remarks>
    '    '''=========================================================================
    '    Public Function Sub_Probe_OnOff(ByVal MD As Integer, Optional ByVal Z2ON As Double = 0.0#) As Integer

    '        Dim r As Integer
    '        Dim strMSG As String

    '        Try
    '            ' Ｚプローブをオン位置へ移動
    '            Sub_Probe_OnOff = cFRS_NORMAL                       ' Return値 = 正常
    '            If (MD = 1) Then                                    ' ON ?
    '                r = Form1.System1.EX_PROBON(gSysPrm)                   ' Z ON位置へ移動
    '                If (r <> cFRS_NORMAL) Then                      ' ｴﾗｰ ?
    '                    Sub_Probe_OnOff = r                         ' Return値 = 非常停止他(※ﾒｯｾｰｼﾞは表示済)
    '                    Exit Function
    '                End If
    '                If ((gSysPrm.stDEV.giPrbTyp And 2) = 2) Then    ' 下方ﾌﾟﾛｰﾌﾞなしならNOP
    '                    r = Form1.System1.EX_PROBON2(gSysPrm, Z2ON)        ' Z2 ON位置へ移動
    '                    If (r <> cFRS_NORMAL) Then                  ' ｴﾗｰ ?
    '                        Sub_Probe_OnOff = r                     ' Return値 = 非常停止他(※ﾒｯｾｰｼﾞは表示済)
    '                        Exit Function
    '                    End If
    '                End If

    '                ' Ｚプローブをオフ位置へ移動
    '            Else
    '                If ((gSysPrm.stDEV.giPrbTyp And 2) = 2) Then    ' 下方ﾌﾟﾛｰﾌﾞなしならNOP
    '                    r = Form1.System1.EX_PROBOFF2(gSysPrm)             ' Z2 OFF位置へ移動
    '                    If (r <> cFRS_NORMAL) Then                  ' ｴﾗｰ ?
    '                        Sub_Probe_OnOff = r                     ' Return値 = 非常停止他(※ﾒｯｾｰｼﾞは表示済)
    '                        Exit Function
    '                    End If
    '                End If
    '                r = Form1.System1.EX_PROBOFF(gSysPrm)                  ' Z OFF位置へ移動
    '                If (r <> cFRS_NORMAL) Then                      ' ｴﾗｰ ?
    '                    Sub_Probe_OnOff = r                         ' Return値 = 非常停止他(※ﾒｯｾｰｼﾞは表示済)
    '                    Exit Function
    '                End If
    '            End If
    '            Exit Function

    '            ' トラップエラー発生時
    '        Catch ex As Exception
    '            strMSG = "globals.Sub_Probe_OnOff() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '            Return (cERR_TRAP)                                  ' Return値 = 例外エラー
    '        End Try
    '    End Function
    '#End Region

    '#Region "加工条件を入力する(FL時)"
    '    '''=========================================================================
    '    '''<summary>加工条件を入力する(FL時)</summary>
    '    ''' <param name="CondNum">(I/O)加工条件番号</param>
    '    ''' <param name="dQrate"> (I/O)Qレート(KHz)</param>
    '    ''' <param name="Owner">  (INP)オーナー</param>
    '    ''' <returns>0=正常, 0以外=エラー</returns>
    '    ''' <remarks>キャリブレーション、カット位置補正(外部カメラ)の十字カット用</remarks>
    '    '''=========================================================================
    '    Public Function Sub_FlCond(ByRef CondNum As Integer, ByRef dQrate As Double, ByVal Owner As IWin32Window) As Integer

    '        Dim r As Integer
    '        Dim ObjForm As Object = Nothing
    '        Dim strMSG As String

    '        Try
    '            ' 加工条件を入力する(FL時)
    '            r = cFRS_NORMAL                                             ' Return値 = 正常
    '            If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then          ' FLでない時
    '                CondNum = 0                                             ' 加工条件番号(Dmy)

    '            Else                                                        ' FL時は加工条件を入力する
    '                ' 加工条件入力画面表示
    '                ObjForm = New FrmFlCond()                               ' オブジェクト生成
    '                Call ObjForm.ShowDialog(Owner, CondNum)                 ' 加工条件入力画面表示
    '                r = ObjForm.GetResult(CondNum, dQrate)                  ' 加工条件取得

    '                ' オブジェクト開放
    '                If (ObjForm Is Nothing = False) Then
    '                    Call ObjForm.Close()                                ' オブジェクト開放
    '                    Call ObjForm.Dispose()                              ' リソース開放
    '                End If
    '            End If

    '            Return (r)                                                  ' Return値設定

    '            ' トラップエラー発生時
    '        Catch ex As Exception
    '            strMSG = "globals.Sub_FlCond() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '            Return (cERR_TRAP)                                          ' Return値 = ﾄﾗｯﾌﾟｴﾗｰ発生
    '        End Try
    '    End Function
    '#End Region

    '#Region "十字カットを行う"
    '    '''=========================================================================
    '    '''<summary>十字カットを行う</summary>
    '    ''' <param name="BPx">         (INP)カット位置X</param>
    '    ''' <param name="BPy">         (INP)カット位置Y</param>
    '    ''' <param name="CondNum">     (INP)加工条件番号(FL用)</param>
    '    ''' <param name="dQrate">      (INP)Qレート(KHz)</param>
    '    ''' <param name="dblCutLength">(INP)カット長</param>
    '    ''' <param name="dblCutSpeed"> (INP)カット速度</param>
    '    ''' <returns>0=正常, 0以外=エラー</returns>
    '    ''' <remarks>※十字ｶｯﾄの中央にBPを移動しておくこと
    '    '''            キャリブレーション、カット位置補正(外部カメラ)の十字カット用</remarks>
    '    '''=========================================================================
    '    Public Function CrossCutExec(ByVal BPx As Double, ByVal BPy As Double, ByVal CondNum As Integer, _
    '                                 ByVal dQrate As Double, ByVal dblCutLength As Double, ByVal dblCutSpeed As Double) As Integer

    '        Dim r As Integer
    '        Dim intXANG As Integer
    '        Dim intYANG As Integer
    '        Dim strMSG As String
    '        Dim stCutCmnPrm As CUT_COMMON_PRM                               ' カットパラメータ

    '        Try
    '            '-------------------------------------------------------------------
    '            '   初期処理
    '            '-------------------------------------------------------------------
    '            Call InitCutParam(stCutCmnPrm)                              ' カットパラメータ初期化

    '            ' カット角度を設定する
    '            Select Case (gSysPrm.stDEV.giBpDirXy)
    '                Case 0      ' x←, y↓
    '                    intXANG = 180
    '                    intYANG = 270
    '                Case 1      ' x→, y↓
    '                    intXANG = 0
    '                    intYANG = 270
    '                Case 2      ' x←, y↑
    '                    intXANG = 180
    '                    intYANG = 90
    '                Case 3      ' x→, y↑
    '                    intXANG = 0
    '                    intYANG = 90
    '            End Select

    '            ' カットパラメータ(カット情報構造体)を設定する
    '            stCutCmnPrm.CutInfo.srtMoveMode = 2                         ' 動作モード（0:トリミング、1:ティーチング、2:強制カット）
    '            stCutCmnPrm.CutInfo.srtCutMode = 4                          ' カットモードは「斜め」
    '            stCutCmnPrm.CutInfo.dblTarget = 1000.0#                     ' 目標値 = 1とする
    '            stCutCmnPrm.CutInfo.srtSlope = 4                            ' 4:抵抗測定＋スロープ
    '            stCutCmnPrm.CutInfo.srtMeasType = 0                         ' 測定タイプ(0:高速(3回)、1:高精度(2000回)
    '            stCutCmnPrm.CutInfo.dblAngle = intXANG                      ' カット角度(X軸)

    '            ' カットパラメータ(加工設定構造体)を設定する
    '            stCutCmnPrm.CutCond.CutLen.dblL1 = dblCutLength             ' カット長(Line1用)
    '            stCutCmnPrm.CutCond.SpdOwd.dblL1 = dblCutSpeed              ' カットスピード（往路）(Line1用)
    '            stCutCmnPrm.CutCond.QRateOwd.dblL1 = dQrate                 ' カットQレート（往路）(Line1用)
    '            stCutCmnPrm.CutCond.CondOwd.srtL1 = CondNum                 ' カット条件番号（往路）(Line1用)

    '            ' Qレート(FL時以外)または加工条件番号(FL時)を設定する
    '            If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then          ' FLでない ?
    '                Call QRATE(dQrate)                                      ' Qレート設定(KHz)
    '            Else                                                        ' 加工条件番号を設定する(FL時)
    '                Call QRATE(dQrate)                                      ' Qレート設定(KHz)
    '                r = FLSET(FLMD_CNDSET, CondNum)                         ' 加工条件番号設定
    '                If (r <> cFRS_NORMAL) Then GoTo STP_ERR_FL
    '            End If

    '            '-------------------------------------------------------------------
    '            '   十字カットのX軸をカットする
    '            '-------------------------------------------------------------------
    '            ' BPをX軸始点へ移動する(絶対値移動)
    '            r = Form1.System1.EX_MOVE(gSysPrm, BPx - (dblCutLength / 2), BPy, 1)
    '            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
    '                Return (r)
    '            End If
    '            ' 十字カットのX軸をカットする
    '            r = Sub_CrossCut(stCutCmnPrm)                               ' X軸カット
    '            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
    '                Return (r)
    '            End If
    '            ' BPを中心へ戻す
    '            r = Form1.System1.EX_MOVE(gSysPrm, BPx, BPy, 1)
    '            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
    '                Return (r)
    '            End If
    '            Call System.Threading.Thread.Sleep(500)                     ' Wait(msec)

    '            '-------------------------------------------------------------------
    '            '   十字カットのY軸をカットする
    '            '-------------------------------------------------------------------
    '            ' BPをY軸始点へ移動する(絶対値移動)
    '            r = Form1.System1.EX_MOVE(gSysPrm, BPx, BPy - (dblCutLength / 2), 1)
    '            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
    '                Return (r)
    '            End If
    '            ' 十字カットのY軸をカットする
    '            stCutCmnPrm.CutInfo.dblAngle = intYANG                      ' カット角度(Y軸)
    '            r = Sub_CrossCut(stCutCmnPrm)                               ' Y軸カット
    '            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
    '                Return (r)
    '            End If
    '            ' BPを中心へ戻す
    '            r = Form1.System1.EX_MOVE(gSysPrm, BPx, BPy, 1)
    '            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
    '                Return (r)
    '            End If
    '            Call System.Threading.Thread.Sleep(500)                     ' Wait(msec)

    '            Return (cFRS_NORMAL)

    '            ' 加工条件番号の設定エラー時(FL時)
    'STP_ERR_FL:
    '            strMSG = MSG_151                                            ' "加工条件の設定に失敗しました｡"
    '            Call Form1.System1.TrmMsgBox(gSysPrm, strMSG, vbOKOnly, gAppName)
    '            Return (cFRS_ERR_RST)                                       ' Return値 = Cancel

    '            ' トラップエラー発生時
    '        Catch ex As Exception
    '            strMSG = "globals.CrossCutExec() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '            Return (cERR_TRAP)                                          ' Return値 = ﾄﾗｯﾌﾟｴﾗｰ発生
    '        End Try
    '    End Function
    '#End Region

    '#Region "十字カットのX軸またはY軸をカットする"
    '    '''=========================================================================
    '    '''<summary>十字カットのX軸またはY軸をカットする</summary>
    '    ''' <param name="stCutCmnPrm">(INP)カットパラメータ</param>
    '    ''' <returns>0=正常, 0以外=エラー</returns>
    '    ''' <remarks>※十字カット位置にBPを移動しておくこと
    '    '''            キャリブレーション、カット位置補正(外部カメラ)の十字カット用</remarks>
    '    '''=========================================================================
    '    Private Function Sub_CrossCut(ByRef stCutCmnPrm As CUT_COMMON_PRM) As Integer

    '        Dim r As Integer
    '        Dim strMSG As String

    '        Try
    '            ' 十字カットのX軸またはY軸をカットする
    '            r = TRIM_ST(stCutCmnPrm)                                    ' STカット
    '            r = Form1.System1.EX_ZGETSRVSIGNAL(gSysPrm, r, 0)
    '            If (r < cFRS_NORMAL) Then                                   ' エラー ?
    '                Return (r)
    '            End If
    '            Return (cFRS_NORMAL)

    '            ' トラップエラー発生時
    '        Catch ex As Exception
    '            strMSG = "globals.Sub_CrossCut() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '            Return (cERR_TRAP)                                          ' Return値 = ﾄﾗｯﾌﾟｴﾗｰ発生
    '        End Try
    '    End Function
    '#End Region

    '#Region "カットパラメータを初期化する"
    '    '''=========================================================================
    '    '''<summary>カットパラメータを初期化する</summary>
    '    ''' <param name="pstCutCmnPrm">(I/O)カットパラメータ</param>
    '    ''' <remarks>キャリブレーション、カット位置補正(外部カメラ)の十字カット用</remarks>
    '    '''=========================================================================
    '    Private Sub InitCutParam(ByRef pstCutCmnPrm As CUT_COMMON_PRM)

    '        Dim strMSG As String

    '        Try
    '            ' カットパラメータを初期化する(カット情報構造体)
    '            pstCutCmnPrm.CutInfo.srtMoveMode = 1                        ' 動作モード（0:トリミング、1:ティーチング、2:強制カット）
    '            pstCutCmnPrm.CutInfo.srtCutMode = 0                         ' カットモード(0:ノーマル、1:リターン、2:リトレース、3:斜め）
    '            pstCutCmnPrm.CutInfo.dblTarget = 0.0#                       ' 目標値
    '            pstCutCmnPrm.CutInfo.srtSlope = 4                           ' 4:抵抗測定＋スロープ
    '            pstCutCmnPrm.CutInfo.srtMeasType = 0                        ' 測定タイプ(0:高速(3回)、1:高精度(2000回)
    '            pstCutCmnPrm.CutInfo.dblAngle = 0.0#                        ' カット角度
    '            pstCutCmnPrm.CutInfo.dblLTP = 0.0#                          ' Lターンポイント
    '            pstCutCmnPrm.CutInfo.srtLTDIR = 0                           ' Lターン後の方向
    '            pstCutCmnPrm.CutInfo.dblRADI = 0.0#                         ' R部回転半径（Uカットで使用）
    '            '                                                           ' For Hook Or UCut
    '            pstCutCmnPrm.CutInfo.dblRADI2 = 0.0#                        ' R2部回転半径（Uカットで使用）
    '            pstCutCmnPrm.CutInfo.srtHkOrUType = 0                       ' HookCut(3)かUカット（3以外）の指定。
    '            '                                                           ' For Index
    '            pstCutCmnPrm.CutInfo.srtIdxScnCnt = 0                       ' インデックス/スキャンカット数(1～32767)
    '            pstCutCmnPrm.CutInfo.srtIdxMeasMode = 0                     ' インデックス測定モード（0:抵抗、1:電圧、2:外部）
    '            '                                                           ' For EdgeSense
    '            pstCutCmnPrm.CutInfo.dblEsPoint = 0.0#                      ' エッジセンスポイント
    '            pstCutCmnPrm.CutInfo.dblRdrJdgVal = 0.0#                    ' ラダー内部判定変化量
    '            pstCutCmnPrm.CutInfo.dblMinJdgVal = 0.0#                    ' ラダーカット後最低許容変化量
    '            pstCutCmnPrm.CutInfo.srtEsAftCutCnt = 0                     ' ラダー切抜け後のカット回数（測定回数）
    '            pstCutCmnPrm.CutInfo.srtMinOvrNgCnt = 0                     ' ラダー抜出し後、最低変化量の連続Over許容数
    '            pstCutCmnPrm.CutInfo.srtMinOvrNgMode = 0                    ' 連続Over時のNG処理（0:NG判定未実施, 1:NG判定実施。ラダー中切り, 2:NG判定未実施。ラダー切上げ）
    '            '                                                           ' For Scan
    '            pstCutCmnPrm.CutInfo.dblStepPitch = 0.0#                    ' ステップ移動ピッチ
    '            pstCutCmnPrm.CutInfo.srtStepDir = 0                         ' ステップ方向

    '            ' カットパラメータを初期化する(加工設定構造体)
    '            pstCutCmnPrm.CutCond.CutLen.dblL1 = 0.0#                    ' カット長(Line1用)
    '            pstCutCmnPrm.CutCond.CutLen.dblL2 = 0.0#                    ' カット長(Line2用)
    '            pstCutCmnPrm.CutCond.CutLen.dblL3 = 0.0#                    ' カット長(Line3用)
    '            pstCutCmnPrm.CutCond.CutLen.dblL4 = 0.0#                    ' カット長(Line4用)

    '            pstCutCmnPrm.CutCond.SpdOwd.dblL1 = 0.0#                    ' カットスピード（往路）(Line1用)
    '            pstCutCmnPrm.CutCond.SpdOwd.dblL2 = 0.0#                    ' カットスピード（往路）(Line2用)
    '            pstCutCmnPrm.CutCond.SpdOwd.dblL3 = 0.0#                    ' カットスピード（往路）(Line3用)
    '            pstCutCmnPrm.CutCond.SpdOwd.dblL4 = 0.0#                    ' カットスピード（往路）(Line4用)

    '            pstCutCmnPrm.CutCond.SpdRet.dblL1 = 0.0#                    ' カットスピード（復路）(Line1用)
    '            pstCutCmnPrm.CutCond.SpdRet.dblL2 = 0.0#                    ' カットスピード（復路）(Line2用)
    '            pstCutCmnPrm.CutCond.SpdRet.dblL3 = 0.0#                    ' カットスピード（復路）(Line3用)
    '            pstCutCmnPrm.CutCond.SpdRet.dblL4 = 0.0#                    ' カットスピード（復路）(Line4用)

    '            pstCutCmnPrm.CutCond.QRateOwd.dblL1 = 0.0#                  ' カットQレート（往路）(Line1用)
    '            pstCutCmnPrm.CutCond.QRateOwd.dblL2 = 0.0#                  ' カットQレート（往路）(Line2用)
    '            pstCutCmnPrm.CutCond.QRateOwd.dblL3 = 0.0#                  ' カットQレート（往路）(Line3用)
    '            pstCutCmnPrm.CutCond.QRateOwd.dblL4 = 0.0#                  ' カットQレート（往路）(Line4用)

    '            pstCutCmnPrm.CutCond.QRateRet.dblL1 = 0.0#                  ' カットQレート（復路）(Line1用)
    '            pstCutCmnPrm.CutCond.QRateRet.dblL2 = 0.0#                  ' カットQレート（復路）(Line2用)
    '            pstCutCmnPrm.CutCond.QRateRet.dblL3 = 0.0#                  ' カットQレート（復路）(Line3用)
    '            pstCutCmnPrm.CutCond.QRateRet.dblL4 = 0.0#                  ' カットQレート（復路）(Line4用)

    '            pstCutCmnPrm.CutCond.CondOwd.srtL1 = 0                      ' カット条件番号（往路）(Line1用)
    '            pstCutCmnPrm.CutCond.CondOwd.srtL2 = 0                      ' カット条件番号（往路）(Line2用)
    '            pstCutCmnPrm.CutCond.CondOwd.srtL3 = 0                      ' カット条件番号（往路）(Line3用)
    '            pstCutCmnPrm.CutCond.CondOwd.srtL4 = 0                      ' カット条件番号（往路）(Line4用)

    '            pstCutCmnPrm.CutCond.CondRet.srtL1 = 0                      ' カット条件番号（復路）(Line1用)
    '            pstCutCmnPrm.CutCond.CondRet.srtL2 = 0                      ' カット条件番号（復路）(Line2用)
    '            pstCutCmnPrm.CutCond.CondRet.srtL3 = 0                      ' カット条件番号（復路）(Line3用)
    '            pstCutCmnPrm.CutCond.CondRet.srtL4 = 0                      ' カット条件番号（復路）(Line4用)

    '            Exit Sub

    '            ' トラップエラー発生時
    '        Catch ex As Exception
    '            strMSG = "globals.InitCutParam() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '        End Try
    '        Exit Sub
    '    End Sub
    '#End Region

    '#Region "パターン認識を実行し、ずれ量を返す"
    '    '''=========================================================================
    '    ''' <summary>パターン認識を実行し、ずれ量を返す</summary>
    '    ''' <param name="GrpNo">    (INP)グループ番号</param>
    '    ''' <param name="TmpNo">    (INP)パターン番号</param>
    '    ''' <param name="fCorrectX">(OUT)ずれ量X</param> 
    '    ''' <param name="fCorrectY">(OUT)ずれ量Y</param>
    '    ''' <param name="coef">     (OUT)相関係数</param> 
    '    ''' <returns>cFRS_NORMAL  = 正常
    '    '''          cFRS_ERR_PTN = パターンパターン認識エラー
    '    '''          上記以外     = その他エラー</returns>
    '    ''' <remarks>・パターン認識位置へテーブルは移動済であること
    '    '''          ・キャリブレーション、カット位置補正(外部カメラ)用
    '    ''' </remarks>
    '    '''=========================================================================
    '    Public Function Sub_PatternMatching(ByRef GrpNo As Short, ByRef TmpNo As Short, ByRef fCorrectX As Double, ByRef fCorrectY As Double, ByRef coef As Double) As Integer

    '        Dim ret As Short = cFRS_NORMAL
    '        Dim crx As Double = 0.0                                         ' ずれ量X
    '        Dim cry As Double = 0.0                                         ' ずれ量Y
    '        Dim fcoeff As Double = 0.0                                      ' 相関値
    '        Dim Thresh As Double = 0.0                                      ' 閾値
    '        Dim r As Integer = cFRS_NORMAL                                  ' 関数値
    '        Dim strMSG As String = ""

    '        Try
    '#If VIDEO_CAPTURE = 1 Then
    '            fCorrectX = 0.0
    '            fCorrectY = 0.0
    '            coef = 0.8
    '            Return (cFRS_NORMAL)   
    '#End If
    '            ' パターンマッチング時のテンプレートグループ番号を設定する(毎回やると遅くなる)
    '            If (giTempGrpNo <> GrpNo) Then                              ' テンプレートグループ番号が変わった ?
    '                giTempGrpNo = GrpNo                                     ' 現在のテンプレートグループ番号を退避
    '                Form1.VideoLibrary1.SelectTemplateGroup(GrpNo)          ' テンプレートグループ番号設定
    '            End If

    '            ' 閾値取得
    '            Thresh = gDllSysprmSysParam_definst.GetPtnMatchThresh(GrpNo, TmpNo)
    '            coef = 0.0                                                  ' 一致度

    '            ' パターンマッチング(外部カメラ)を行う(Video.ocxを使用)
    '            ret = Form1.VideoLibrary1.PatternMatching_EX(TmpNo, 1, True, crx, cry, fcoeff)
    '            If (ret = cFRS_NORMAL) Then
    '                r = cFRS_NORMAL                                         ' Return値 = 正常
    '                fCorrectX = crx                                         ' ずれ量X
    '                fCorrectY = cry                                         ' ずれ量Y
    '                '' マッチしたパターンの測定位置からずれ量を求める
    '                'fCorrectX = crx / 1000.0#
    '                'fCorrectY = -cry / 1000.0#
    '                coef = fcoeff                                           ' 相関係数
    '                strMSG = "パターン認識成功"
    '                If (fcoeff < Thresh) Then
    '                    r = cFRS_ERR_PT2                                    ' パターン認識エラー(閾値エラー)
    '                    strMSG = "パターン認識エラー(閾値エラー)"
    '                End If
    '                strMSG = strMSG + " (相関係数=" + Format(fcoeff, "0.000") + " ずれ量X=" + Format(crx, "0.0000") + ", ずれ量X=" + Format(cry, "0.0000") + ")"
    '            Else
    '                r = cFRS_ERR_PTN                                        ' パターン認識エラー(パターンが見つからなかった)
    '                strMSG = "パターン認識エラー(パターンが見つからなかった)"
    '            End If

    '            ' 後処理
    '            Console.WriteLine(strMSG)
    '            Return (r)

    '            ' トラップエラー発生時 
    '        Catch ex As Exception
    '            strMSG = "globals.Sub_PatternMatching() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '            Return (cERR_TRAP)                                          ' Return値 = 例外エラー
    '        End Try
    '    End Function
    '#End Region

    '#Region "指定ブロックの中央へテーブルを移動する"
    '    '''=========================================================================
    '    '''<summary>指定ブロックの中央へテーブルを移動する</summary>
    '    '''<param name="intCamera">(INP)ｶﾒﾗ種類(0:内部ｶﾒﾗ 1:外部ｶﾒﾗ)</param>
    '    '''<param name="iXPlate">(INP)XPlateNo</param> 
    '    '''<param name="iYPlate">(INP)YPlateNo</param>  
    '    '''<param name="iXBlock">(INP)XBlockNo</param> 
    '    '''<param name="iYBlock">(INP)YBlockNo</param>   
    '    '''<remarks>十字ｶｯﾄ位置はﾃｨｰﾁﾝｸﾞﾎﾟｲﾝﾄからﾌﾟﾚｰﾄﾃﾞｰﾀ
    '    '''         のPP47の値分ずれたところが中心となる
    '    '''         現状はﾌﾟﾚｰﾄを指定しても意味なし</remarks>
    '    '''=========================================================================
    '    Public Function XYTableMoveBlock(ByRef intCamera As Short, ByRef iXPlate As Short, ByRef iYPlate As Short, ByRef iXBlock As Short, ByRef iYBlock As Short) As Short

    '        Dim dblX As Double
    '        Dim dblY As Double
    '        Dim dblRotX As Double
    '        Dim dblRotY As Double
    '        Dim dblPSX As Double
    '        Dim dblPSY As Double
    '        Dim dblBsoX As Double
    '        Dim dblBsoY As Double
    '        Dim dblBSX As Double
    '        Dim dblBSY As Double
    '        Dim intCDir As Short
    '        Dim dblTrimPosX As Double
    '        Dim dblTrimPosY As Double
    '        Dim dblTOffsX As Double
    '        Dim dblTOffsY As Double
    '        Dim dblStepInterval As Double
    '        Dim Del_x As Double
    '        Dim Del_y As Double
    '        Dim r As Short
    '        Dim strMSG As String

    '        Try
    '            dblRotX = 0
    '            dblRotY = 0

    '            ' ﾄﾘﾑﾎﾟｼﾞｼｮﾝX,Y取得
    '            dblTrimPosX = gSysPrm.stDEV.gfTrimX                 ' ﾄﾘﾑﾎﾟｼﾞｼｮﾝX,Y取得
    '            dblTrimPosY = gSysPrm.stDEV.gfTrimY
    '            ' ﾃｰﾌﾞﾙ位置ｵﾌｾｯﾄの取得
    '            dblTOffsX = typPlateInfo.dblTableOffsetXDir : dblTOffsY = typPlateInfo.dblTableOffsetYDir

    '            Call CalcBlockSize(dblBSX, dblBSY)                  ' ﾌﾞﾛｯｸｻｲｽﾞ算出

    '            ' ﾌﾞﾛｯｸｻｲｽﾞｵﾌｾｯﾄ算出　ﾌﾞﾛｯｸｻｲｽﾞ/2 ﾌﾞﾛｯｸの象限はXYともに1 ﾃｰﾌﾞﾙの象限も1
    '            dblBsoX = (dblBSX / 2.0#) * 1 * 1                   ' Table.BDirX * Table.dir
    '            dblBsoY = (dblBSY / 2) * 1                          ' Table.BDirY;

    '            ' θ補正ﾄﾘﾑｵﾌｾｯﾄX,Y
    '            Del_x = gfCorrectPosX
    '            Del_y = gfCorrectPosY

    '            ' giBpDirXy 座標系の設定(ｼｽﾃﾑ設定)
    '            ' 0:XY NOM(右上)  1:X REV(左上)  2:Y REV(右下)  3:XY REV(左下)
    '            ' ﾄﾘﾐﾝｸﾞ位置座標 (+or-) 回転半径 + ﾃｰﾌﾞﾙｵﾌｾｯﾄ (+or-) ﾌﾞﾛｯｸｻｲｽﾞｵﾌｾｯﾄ + ﾃｰﾌﾞﾙ補正量
    '            Select Case gSysPrm.stDEV.giBpDirXy

    '                Case 0 ' x←, y↓
    '                    dblX = dblTrimPosX + dblRotX + dblTOffsX + dblBsoX + Del_x
    '                    dblY = dblTrimPosY + dblRotY + dblTOffsY + dblBsoY + Del_y

    '                Case 1 ' x→, y↓
    '                    dblX = dblTrimPosX - dblRotX + dblTOffsX - dblBsoX + Del_x
    '                    dblY = dblTrimPosY + dblRotY + dblTOffsY + dblBsoY + Del_y

    '                Case 2 ' x←, y↑
    '                    dblX = dblTrimPosX + dblRotX + dblTOffsX + dblBsoX + Del_x
    '                    dblY = dblTrimPosY - dblRotY + dblTOffsY - dblBsoY + Del_y

    '                Case 3 ' x→, y↑
    '                    dblX = dblTrimPosX - dblRotX + dblTOffsX - dblBsoX + Del_x
    '                    dblY = dblTrimPosY - dblRotY + dblTOffsY - dblBsoY + Del_y

    '            End Select

    '            If (1 = intCamera) Then                             ' 外部ｶﾒﾗ位置加算 ?
    '                dblX = dblX + gSysPrm.stDEV.gfExCmX
    '                dblY = dblY + gSysPrm.stDEV.gfExCmY
    '            End If

    '            'ｽﾃｯﾌﾟ間隔の算出
    '            intCDir = typPlateInfo.intResistDir                 ' チップ並び方向取得(CHIP-NETのみ)

    '            If intCDir = 0 Then                                 ' X方向
    '                dblStepInterval = CalcStepInterval(iYBlock)     ' ｽﾃｯﾌﾟｲﾝﾀｰﾊﾞﾙ算出(Y軸)
    '                If gSysPrm.stDEV.giBpDirXy = 0 Or gSysPrm.stDEV.giBpDirXy = 1 Then ' ﾃｰﾌﾞﾙY方向反転なし
    '                    dblY = dblY + dblStepInterval
    '                Else                                            ' ﾃｰﾌﾞﾙY方向反転
    '                    dblY = dblY - dblStepInterval
    '                End If
    '            Else                                                ' Y方向
    '                dblStepInterval = CalcStepInterval(iXBlock)     ' ｽﾃｯﾌﾟｲﾝﾀｰﾊﾞﾙ算出(X軸)
    '                If gSysPrm.stDEV.giBpDirXy = 0 Or gSysPrm.stDEV.giBpDirXy = 2 Then ' ﾃｰﾌﾞﾙX方向反転なし
    '                    dblX = dblX + dblStepInterval
    '                Else                                            ' ﾃｰﾌﾞﾙX方向反転
    '                    dblX = dblX - dblStepInterval
    '                End If
    '            End If

    '            ' ﾌﾟﾚｰﾄ/ﾌﾞﾛｯｸ位置の相対座標計算
    '            dblPSX = 0.0 : dblPSY = 0.0                         ' ﾌﾟﾚｰﾄｻｲｽﾞ取得(0固定)
    '            Select Case gSysPrm.stDEV.giBpDirXy

    '                Case 0 ' x←, y↓
    '                    dblX = dblX + ((dblPSX * CInt(iXPlate - 1)) + (dblBSX * CInt(iXBlock - 1)))
    '                    dblY = dblY + ((dblPSY * CInt(iYPlate - 1)) + (dblBSY * CInt(iYBlock - 1)))

    '                Case 1 ' x→, y↓
    '                    dblX = dblX - ((dblPSX * CInt(iXPlate - 1)) + (dblBSX * CInt(iXBlock - 1)))
    '                    dblY = dblY + ((dblPSY * CInt(iYPlate - 1)) + (dblBSY * CInt(iYBlock - 1)))

    '                Case 2 ' x←, y↑
    '                    dblX = dblX + ((dblPSX * CInt(iXPlate - 1)) + (dblBSX * CInt(iXBlock - 1)))
    '                    dblY = dblY - ((dblPSY * CInt(iYPlate - 1)) + (dblBSY * CInt(iYBlock - 1)))

    '                Case 3 ' x→, y↑
    '                    dblX = dblX - ((dblPSX * CInt(iXPlate - 1)) + (dblBSX * CInt(iXBlock - 1)))
    '                    dblY = dblY - ((dblPSY * CInt(iYPlate - 1)) + (dblBSY * CInt(iYBlock - 1)))

    '            End Select

    '            ' 指定ﾌﾟﾚｰﾄ/ﾌﾞﾛｯｸ位置にXYﾃｰﾌﾞﾙ絶対値移動
    '            r = Form1.System1.XYtableMove(gSysPrm, dblX, dblY)
    '            Return (r)                                      ' Return

    '            ' トラップエラー発生時
    '        Catch ex As Exception
    '            strMSG = "globals.XYTableMoveBlock() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '            Return (cERR_TRAP)                              ' Return値 = 例外エラー 
    '        End Try

    '    End Function
    '#End Region

    '#Region "GPIBコマンドを設定する"
    '    '''=========================================================================
    '    '''<summary>GPIBコマンドを設定する</summary>
    '    '''<param name="pltInfo">(OUT)プレートデータ</param>
    '    '''=========================================================================
    '    Public Sub SetGpibCommand(ByRef pltInfo As PlateInfo)

    '        Dim strDAT As String
    '        Dim strMSG As String

    '        Try
    '            ' ADEX AX-1152用設定コマンドを設定する
    '            pltInfo.intGpibDefAdder = giGpibDefAdder                ' GPIBアドレス 
    '            pltInfo.intGpibDefDelimiter = 0                         ' 初期設定(ﾃﾞﾘﾐﾀ)(固定)
    '            pltInfo.intGpibDefTimiout = 100                         ' 初期設定(ﾀｲﾑｱｳﾄ)(固定)
    '            If (pltInfo.intGpibMeasSpeed = 0) Then                  ' 測定速度(0:低速, 1:高速)
    '                strDAT = "W0"
    '            Else
    '                strDAT = "W1"
    '            End If

    '            '// 測定モードで切り替え
    '            If (pltInfo.intGpibMeasMode = 0) Then                   ' 測定モード(0:絶対, 1:偏差)
    '                strDAT = strDAT + "FR"                              ' 測定モード=絶対
    '                strDAT = strDAT + "LL00000" + "LH15000"             ' 下限/上限リミットの設定
    '            Else

    '                strDAT = strDAT + "FD"                              ' 測定モード=偏差
    '                strDAT = strDAT + "DL-5000" + "DH+5000"             ' 下限/上限リミットの設定
    '            End If

    '            pltInfo.strGpibInitCmnd1 = strDAT                       ' 初期化ｺﾏﾝﾄﾞ1
    '            pltInfo.strGpibInitCmnd2 = ""                           ' 初期化ｺﾏﾝﾄﾞ2
    '            pltInfo.strGpibTriggerCmnd = "E"                        ' ﾄﾘｶﾞｺﾏﾝﾄﾞ

    '            ' トラップエラー発生時
    '        Catch ex As Exception
    '            strMSG = "globals.SetGpibCommand() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '        End Try
    '    End Sub
    '#End Region
    '    '----- ###211↓ -----
    '#Region "START/RESETキー押下待ちサブルーチン"
    '    '''=========================================================================
    '    ''' <summary>START/RESETキー押下待ちサブルーチン</summary>
    '    ''' <param name="Md">(INP)cFRS_ERR_START                = STARTキー押下待ち
    '    '''                       cFRS_ERR_RST                  = RESETキー押下待ち
    '    '''                       cFRS_ERR_START + cFRS_ERR_RST = START/RESETキー押下待ち
    '    ''' </param>
    '    ''' <param name="bZ">(INP)True=Zキー押下チェックする, False=しない ###220</param>
    '    ''' <returns>cFRS_ERR_START = STARTキー押下
    '    '''          cFRS_ERR_RST   = RESETキー押下
    '    '''          cFRS_ERR_Z     = Zキー押下
    '    '''          上記以外=エラー
    '    ''' </returns>
    '    '''=========================================================================
    '    Public Function WaitStartRestKey(ByVal Md As Integer, ByVal bZ As Boolean) As Integer

    '        Dim sts As Long = 0
    '        Dim r As Long = 0
    '        Dim ExitFlag As Integer
    '        Dim strMSG As String

    '        Try
    '            ' パラメータチェック
    '            If (Md = 0) Then
    '                Return (-1 * ERR_CMD_PRM)                               ' パラメータエラー
    '            End If

    '#If cOFFLINEcDEBUG Then                                                 ' OffLineﾃﾞﾊﾞｯｸﾞON ?(↓FormResetが最前面表示なので下記のようにしないとMsgBoxが最前面表示されない)
    '            Dim Dr As System.Windows.Forms.DialogResult
    '            Dr = MessageBox.Show("START SW CHECK", "Debug", MessageBoxButtons.OKCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
    '            If (Dr = System.Windows.Forms.DialogResult.OK) Then
    '                ExitFlag = cFRS_ERR_START                               ' Return値 = STARTキー押下
    '            Else
    '                ExitFlag = cFRS_ERR_RST                                 ' Return値 = RESETキー押下
    '            End If
    '            Return (ExitFlag)
    '#End If
    '            ' START/RESETキー押下待ち
    '            Call ZCONRST()                                              ' コンソールキーラッチ解除
    '            ExitFlag = -1
    '            Call Form1.System1.SetSysParam(gSysPrm)                     ' システムパラメータの設定(OcxSystem用)

    '            ' START/RESETキー押下待ち
    '            Do
    '                r = STARTRESET_SWCHECK(False, sts)                      ' START/RESET SW押下チェック
    '                If (sts = cFRS_ERR_RST) And ((Md = cFRS_ERR_RST) Or (Md = cFRS_ERR_START + cFRS_ERR_RST)) Then
    '                    ExitFlag = cFRS_ERR_RST                             ' ExitFlag = Cancel(RESETキー)
    '                ElseIf (sts = cFRS_ERR_START) And ((Md = cFRS_ERR_START) Or (Md = cFRS_ERR_START + cFRS_ERR_RST)) Then
    '                    ExitFlag = cFRS_ERR_START                           ' ExitFlag = OK(STARTキー)
    '                End If
    '                '----- ###220↓ -----
    '                If (bZ = True) Then
    '                    r = Z_SWCHECK(sts)                                  ' Z SW押下チェック
    '                    If (sts <> 0) Then
    '                        ExitFlag = cFRS_ERR_Z                           ' ExitFlag = Zキー押下
    '                    End If
    '                End If
    '                '----- ###220↑ -----
    '                System.Windows.Forms.Application.DoEvents()             ' メッセージポンプ
    '                Call System.Threading.Thread.Sleep(100)                 ' Wait(msec)

    '                ' システムエラーチェック
    '                r = Form1.System1.SysErrChk_ForVBNET(giAppMode)
    '                If (r <> cFRS_NORMAL) Then                              ' 非常停止等(メッセージは表示済) ?
    '                    ExitFlag = r
    '                    Exit Do
    '                End If
    '            Loop While (ExitFlag = -1)

    '            Call ZCONRST()                                              ' コンソールキーラッチ解除
    '            Return (ExitFlag)

    '            ' トラップエラー発生時 
    '        Catch ex As Exception
    '            strMSG = "Globals.WaitRestKey() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '            Return (cERR_TRAP)                                          ' Retuen値 = ﾄﾗｯﾌﾟｴﾗｰ発生
    '        End Try
    '    End Function
    '#End Region
    '    '----- ###211↑ -----

    '    '===========================================================================
    '    '   汎用タイマー
    '    '===========================================================================
    '    Private bTmTimeOut As Boolean                                       ' タイムアウトフラグ

    '#Region "汎用タイマー生成"
    '    '''=========================================================================
    '    ''' <summary>汎用タイマー生成</summary>
    '    ''' <param name="TimerTM">(I/O)タイマー</param>
    '    ''' <param name="TimeVal">(INP)タイムアウト値(msec)</param>
    '    ''' <remarks>タイマー生成した場合はTimerTM_DisposeをCallしてタイマーを破棄する事</remarks>
    '    '''=========================================================================
    '    Public Sub TimerTM_Create(ByRef TimerTM As System.Threading.Timer, ByVal TimeVal As Integer)

    '        Dim strMSG As String

    '        Try
    '            ' タイムアウトチェック用タイマーオブジェクトの作成(TimerTM_TickをTimeVal msec間隔で実行する)
    '            bTmTimeOut = False                                          ' タイムアウトフラグOFF
    '            'TimerTM = New System.Threading.Timer(New System.Threading.TimerCallback(AddressOf TimerTM_Tick), Nothing, TimeVal, TimeVal)
    '            TimerTM = New System.Threading.Timer(New System.Threading.TimerCallback(AddressOf TimerTM_Tick), Nothing, System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)

    '            ' トラップエラー発生時 
    '        Catch ex As Exception
    '            strMSG = "globals.TimerTM_Create() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '        End Try
    '    End Sub
    '#End Region

    '#Region "汎用タイマー開始"
    '    '''=========================================================================
    '    ''' <summary>汎用タイマー開始</summary>
    '    ''' <param name="TimerTM">(INP)タイマー</param>
    '    '''=========================================================================
    '    Public Sub TimerTM_Start(ByRef TimerTM As System.Threading.Timer, ByVal TimeVal As Integer)

    '        Dim strMSG As String

    '        Try
    '            If (TimerTM Is Nothing) Then Return
    '            TimerTM.Change(TimeVal, TimeVal)
    '            Exit Sub

    '            ' トラップエラー発生時 
    '        Catch ex As Exception
    '            strMSG = "globals.TimerTM_Start() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '        End Try
    '    End Sub
    '#End Region

    '#Region "汎用タイマー停止(コールバックメソッド(TimerTM_Tick)の呼出しを停止する)"
    '    '''=========================================================================
    '    ''' <summary>汎用タイマー停止(コールバックメソッド(TimerTM_Tick)の呼出しを停止する)</summary>
    '    ''' <param name="TimerTM">(INP)タイマー</param>
    '    '''=========================================================================
    '    Public Sub TimerTM_Stop(ByRef TimerTM As System.Threading.Timer)

    '        Dim strMSG As String

    '        Try
    '            ' コールバックメソッドの呼出しを停止する
    '            If (TimerTM Is Nothing) Then Return
    '            TimerTM.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)
    '            Exit Sub

    '            ' トラップエラー発生時 
    '        Catch ex As Exception
    '            strMSG = "globals.TimerTM_Stop() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '        End Try
    '    End Sub
    '#End Region

    '#Region "汎用タイマーを破棄する"
    '    '''=========================================================================
    '    ''' <summary>汎用タイマーを破棄する</summary>
    '    ''' <param name="TimerTM">(I/O)タイマー</param>
    '    '''=========================================================================
    '    Public Sub TimerTM_Dispose(ByRef TimerTM As System.Threading.Timer)

    '        Dim strMSG As String

    '        Try
    '            ' コールバックメソッドの呼出しを停止する
    '            If (TimerTM Is Nothing) Then Return
    '            TimerTM.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)
    '            TimerTM.Dispose()                                           ' タイマーを破棄する
    '            Exit Sub

    '            ' トラップエラー発生時 
    '        Catch ex As Exception
    '            strMSG = "globals.TimerTM_Dispose() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '        End Try
    '    End Sub
    '#End Region

    '#Region "タイムアウトフラグを返す"
    '    '''=========================================================================
    '    ''' <summary>タイムアウトフラグを返す</summary>
    '    ''' <returns>Trur=タイムアウト, False=タイムアウトでない</returns>
    '    '''=========================================================================
    '    Public Function TimerTM_Sts() As Boolean

    '        Dim strMSG As String

    '        Try
    '            ' タイムアウトフラグを返す
    '            Return (bTmTimeOut)

    '            ' トラップエラー発生時 
    '        Catch ex As Exception
    '            strMSG = "globals.TimerTM_Sts() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '            Return (bTmTimeOut)
    '        End Try
    '    End Function
    '#End Region

    '#Region "タイマーイベント(指定タイマ間隔が経過した時に発生)"
    '    '''=========================================================================
    '    ''' <summary>タイマーイベント(指定タイマ間隔が経過した時に発生)</summary>
    '    ''' <param name="Sts">(INP)</param>
    '    '''=========================================================================
    '    Private Sub TimerTM_Tick(ByVal Sts As Object)

    '        Dim strMSG As String

    '        Try
    '            bTmTimeOut = True                                           ' タイムアウトフラグON
    '            Exit Sub

    '            ' トラップエラー発生時 
    '        Catch ex As Exception
    '            strMSG = "globals.TimerTM_Tick() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '        End Try
    '    End Sub
    '#End Region

    '    '========================================================================================
    '    '   分布図表示関連処理
    '    '========================================================================================
    '    '' '' ''#Region "分布図表示"
    '    '' '' ''    '''=========================================================================
    '    '' '' ''    '''<summary>分布図表示</summary>
    '    '' '' ''    '''<remarks></remarks>
    '    '' '' ''    '''=========================================================================
    '    '' '' ''    Public Sub picGraphAccumulationRedraw()

    '    '' '' ''        Dim iCnt As Short 'ｶｳﾝﾀ
    '    '' '' ''        Dim lMax As Integer
    '    '' '' ''        Dim lScale As Integer
    '    '' '' ''        Dim lScaleMax As Integer
    '    '' '' ''        Dim dblGraphDiv As Double
    '    '' '' ''        Dim dblGraphTop As Double
    '    '' '' ''        Dim digL As Integer
    '    '' '' ''        Dim digH As Integer
    '    '' '' ''        Dim digSW As Integer

    '    '' '' ''        lMax = 0
    '    '' '' ''        If (bFgDispGrp) Then
    '    '' '' ''            Form1.lblGraphAccumulationTitle.Text = MSG_TRIM_04
    '    '' '' ''            Form1.lblMinValue.Text = dblMinIT.ToString("0.000") ' 最小値
    '    '' '' ''            Form1.lblMaxValue.Text = dblMaxIT.ToString("0.000") ' 最大値

    '    '' '' ''            For iCnt = 0 To (MAX_SCALE_RNUM - 1)

    '    '' '' ''                glRegistNum(iCnt) = glRegistNumIT(iCnt)

    '    '' '' ''                If lMax < glRegistNum(iCnt) Then
    '    '' '' ''                    lMax = glRegistNum(iCnt)
    '    '' '' ''                End If

    '    '' '' ''                gDistRegNumLblAry(iCnt).Text = CStr(glRegistNum(iCnt)) ' 分布グラフ抵抗数

    '    '' '' ''            Next
    '    '' '' ''        Else

    '    '' '' ''            Form1.lblGraphAccumulationTitle.Text = MSG_TRIM_05
    '    '' '' ''            Form1.lblMinValue.Text = dblMinFT.ToString("0.000") ' 最小値
    '    '' '' ''            Form1.lblMaxValue.Text = dblMaxFT.ToString("0.000") ' 最大値

    '    '' '' ''            For iCnt = 0 To (MAX_SCALE_RNUM - 1)

    '    '' '' ''                glRegistNum(iCnt) = glRegistNumFT(iCnt)

    '    '' '' ''                If lMax < glRegistNum(iCnt) Then
    '    '' '' ''                    lMax = glRegistNum(iCnt)
    '    '' '' ''                End If

    '    '' '' ''                gDistRegNumLblAry(iCnt).Text = CStr(glRegistNum(iCnt)) ' 分布グラフ抵抗数

    '    '' '' ''            Next
    '    '' '' ''        End If

    '    '' '' ''        Form1.lblGoodChip.Text = CStr(lOkChip)                        ' OK数
    '    '' '' ''        Form1.lblNgChip.Text = CStr(lNgChip)                          ' NG数

    '    '' '' ''        ' 誤差ﾃﾞｰﾀがある(IT)
    '    '' '' ''        Call Form1.GetMoveMode(digL, digH, digSW)
    '    '' '' ''        If ITNx_cnt >= 0 Then
    '    '' '' ''            If (digL = 0) Then                                 ' x0モード ?
    '    '' '' ''                ' 平均値取得
    '    '' '' ''                dblAverageIT = Form1.Utility1.GetAverage(ITNx, ITNx_cnt + 1)
    '    '' '' ''                ' 標準偏差の取得
    '    '' '' ''                dblDeviationIT = Form1.Utility1.GetDeviation(ITNx, ITNx_cnt + 1, dblAverageIT)
    '    '' '' ''            End If
    '    '' '' ''        End If

    '    '' '' ''        ' 誤差ﾃﾞｰﾀがある(FT)
    '    '' '' ''        If FTNx_cnt >= 0 Then
    '    '' '' ''            ' 平均値取得
    '    '' '' ''            dblAverageFT = Form1.Utility1.GetAverage(FTNx, FTNx_cnt + 1)
    '    '' '' ''            ' 標準偏差の取得
    '    '' '' ''            dblDeviationFT = Form1.Utility1.GetDeviation(FTNx, FTNx_cnt + 1, dblAverageFT)
    '    '' '' ''        End If

    '    '' '' ''        If (bFgDispGrp) Then
    '    '' '' ''            Form1.lblDeviationValue.Text = dblDeviationIT.ToString("0.000000") ' 標準偏差(IT)
    '    '' '' ''        Else
    '    '' '' ''            Form1.lblDeviationValue.Text = dblDeviationFT.ToString("0.000000") ' 標準偏差(FT)
    '    '' '' ''        End If

    '    '' '' ''        If (bFgDispGrp) Then
    '    '' '' ''            dblAverage = dblAverageIT
    '    '' '' ''        Else
    '    '' '' ''            dblAverage = dblAverageFT
    '    '' '' ''        End If
    '    '' '' ''        Form1.lblAverageValue.Text = dblAverage.ToString("0.000")             ' 平均値

    '    '' '' ''        lScaleMax = 0 ' オートスケーリング
    '    '' '' ''        lScale = 100
    '    '' '' ''        Do
    '    '' '' ''            If (lScale > lMax) Then
    '    '' '' ''                lScaleMax = lScale
    '    '' '' ''            ElseIf ((lScale * 2) > lMax) Then
    '    '' '' ''                lScaleMax = (lScale * 2)
    '    '' '' ''            ElseIf ((lScale * 5) > lMax) Then
    '    '' '' ''                lScaleMax = (lScale * 5)
    '    '' '' ''            End If
    '    '' '' ''            lScale = lScale * 10
    '    '' '' ''        Loop While (0 = lScaleMax) And (MAX_SCALE_NUM > lScale)
    '    '' '' ''        If (0 = lScaleMax) Then
    '    '' '' ''            lScaleMax = MAX_SCALE_NUM + 1
    '    '' '' ''        End If

    '    '' '' ''        If (bFgDispGrp) Then
    '    '' '' ''            If ((0 >= typResistorInfoArray(1).dblInitTest_LowLimit) And (0 <= typResistorInfoArray(1).dblInitTest_HighLimit)) Then
    '    '' '' ''                dblGraphDiv = (typResistorInfoArray(1).dblInitTest_HighLimit * 1.5 - typResistorInfoArray(1).dblInitTest_LowLimit * 1.5) / 10
    '    '' '' ''                dblGraphTop = typResistorInfoArray(1).dblInitTest_HighLimit * 1.5
    '    '' '' ''            ElseIf ((0 >= typResistorInfoArray(1).dblInitTest_LowLimit) And (0 > typResistorInfoArray(1).dblInitTest_HighLimit)) Then
    '    '' '' ''                dblGraphDiv = (typResistorInfoArray(1).dblInitTest_HighLimit / 1.5 - typResistorInfoArray(1).dblInitTest_LowLimit * 1.5) / 10
    '    '' '' ''                dblGraphTop = typResistorInfoArray(1).dblInitTest_HighLimit * 1.5
    '    '' '' ''            ElseIf ((0 < typResistorInfoArray(1).dblInitTest_LowLimit) And (0 <= typResistorInfoArray(1).dblInitTest_HighLimit)) Then
    '    '' '' ''                dblGraphDiv = (typResistorInfoArray(1).dblInitTest_HighLimit * 1.5 - typResistorInfoArray(1).dblInitTest_LowLimit / 1.5) / 10
    '    '' '' ''                dblGraphTop = typResistorInfoArray(1).dblInitTest_HighLimit * 1.5
    '    '' '' ''            Else
    '    '' '' ''                dblGraphDiv = 0.3
    '    '' '' ''                dblGraphTop = 1.5
    '    '' '' ''            End If
    '    '' '' ''        Else
    '    '' '' ''            If ((0 >= typResistorInfoArray(1).dblFinalTest_LowLimit) And (0 <= typResistorInfoArray(1).dblFinalTest_HighLimit)) Then
    '    '' '' ''                dblGraphDiv = (typResistorInfoArray(1).dblFinalTest_HighLimit * 1.5 - typResistorInfoArray(1).dblFinalTest_LowLimit * 1.5) / 10
    '    '' '' ''                dblGraphTop = typResistorInfoArray(1).dblFinalTest_HighLimit * 1.5
    '    '' '' ''            ElseIf ((0 >= typResistorInfoArray(1).dblFinalTest_LowLimit) And (0 > typResistorInfoArray(1).dblFinalTest_HighLimit)) Then
    '    '' '' ''                dblGraphDiv = (typResistorInfoArray(1).dblFinalTest_HighLimit / 1.5 - typResistorInfoArray(1).dblFinalTest_LowLimit * 1.5) / 10
    '    '' '' ''                dblGraphTop = typResistorInfoArray(1).dblFinalTest_HighLimit * 1.5
    '    '' '' ''            ElseIf ((0 < typResistorInfoArray(1).dblFinalTest_LowLimit) And (0 <= typResistorInfoArray(1).dblFinalTest_HighLimit)) Then
    '    '' '' ''                dblGraphDiv = (typResistorInfoArray(1).dblFinalTest_HighLimit * 1.5 - typResistorInfoArray(1).dblFinalTest_LowLimit / 1.5) / 10
    '    '' '' ''                dblGraphTop = typResistorInfoArray(1).dblFinalTest_HighLimit * 1.5
    '    '' '' ''            Else
    '    '' '' ''                dblGraphDiv = 0.3
    '    '' '' ''                dblGraphTop = 1.5
    '    '' '' ''            End If
    '    '' '' ''        End If

    '    '' '' ''        gDistGrpPerLblAry(0).Text = "～" & dblGraphTop.ToString("0.00")
    '    '' '' ''        For iCnt = 1 To 11
    '    '' '' ''            gDistGrpPerLblAry(iCnt).Text = (dblGraphTop - (dblGraphDiv * (iCnt - 1))).ToString("0.00") & "～"
    '    '' '' ''        Next

    '    '' '' ''        picGraphAccumulationDrawSubLine()
    '    '' '' ''        picGraphAccumulationDrawLine(lScaleMax)
    '    '' '' ''        picGraphAccumulationPrintRegistNum()        ' 分布グラフに抵抗数を設定する
    '    '' '' ''    End Sub
    '    '' '' ''#End Region

    '#Region "分布図表示サブ"
    '    '''=========================================================================
    '    '''<summary>分布図表示サブ</summary>
    '    '''<remarks></remarks>
    '    '''=========================================================================
    '    Public Sub picGraphAccumulationDrawSubLine()
    '        'Dim i As Short

    '        '      'UPGRADE_ISSUE: PictureBox メソッド picGraphAccumulation.Line はアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"' をクリックしてください。
    '        'picGraphAccumulation.Line (56, 16) - (56, 112), RGB(0, 255, 0)
    '        '      'UPGRADE_ISSUE: PictureBox メソッド picGraphAccumulation.Line はアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"' をクリックしてください。
    '        'picGraphAccumulation.Line (56, 112) - (288, 112), RGB(0, 255, 0)
    '        '      For i = 0 To 10
    '        '          'UPGRADE_ISSUE: PictureBox メソッド picGraphAccumulation.Line はアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"' をクリックしてください。
    '        '	picGraphAccumulation.Line (56, 24 + (i * 8)) - (288, 24 + (i * 8)), RGB(0, 0, 128)
    '        '      Next
    '        '      'UPGRADE_ISSUE: PictureBox メソッド picGraphAccumulation.Line はアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"' をクリックしてください。
    '        'picGraphAccumulation.Line (172, 112) - (172, 116), RGB(0, 255, 0)

    '    End Sub
    '#End Region

    '    '' '' ''#Region "分布図表示サブ"
    '    '' '' ''    '''=========================================================================
    '    '' '' ''    '''<summary>分布図表示サブ</summary>
    '    '' '' ''    '''<remarks></remarks>
    '    '' '' ''    '''=========================================================================
    '    '' '' ''    Public Sub picGraphAccumulationDrawLine(ByRef lScaleMax As Integer)
    '    '' '' ''        Dim i As Short
    '    '' '' ''        Dim X As Short

    '    '' '' ''        For i = 0 To 11
    '    '' '' ''            X = CShort((glRegistNum(i) * 232) \ lScaleMax) ' 分布グラフ抵抗数
    '    '' '' ''            If (232 < X) Then
    '    '' '' ''                X = 232
    '    '' '' ''            End If
    '    '' '' ''            '         'UPGRADE_ISSUE: PictureBox メソッド picGraphAccumulation.Line はアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"' をクリックしてください。
    '    '' '' ''            'picGraphAccumulation.Line (56, 18 + (i * 8)) - (288, 22 + (i * 8)), RGB(0, 0, 0), BF
    '    '' '' ''            '         'UPGRADE_ISSUE: PictureBox メソッド picGraphAccumulation.Line はアップグレードされませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"' をクリックしてください。
    '    '' '' ''            'picGraphAccumulation.Line (56, 18 + (i * 8)) - (56 + X, 22 + (i * 8)), RGB(0, 255, 255), BF
    '    '' '' ''        Next
    '    '' '' ''        Form1.lblRegistUnit.Text = CStr(lScaleMax \ 2)
    '    '' '' ''    End Sub
    '    '' '' ''#End Region

    '    '' '' ''#Region "分布グラフに抵抗数を設定する"
    '    '' '' ''    '''=========================================================================
    '    '' '' ''    '''<summary>分布グラフに抵抗数を設定する</summary>
    '    '' '' ''    '''<remarks></remarks>
    '    '' '' ''    '''=========================================================================
    '    '' '' ''    Public Sub picGraphAccumulationPrintRegistNum()
    '    '' '' ''        Dim i As Short

    '    '' '' ''        For i = 0 To (MAX_SCALE_RNUM - 1)
    '    '' '' ''            gDistRegNumLblAry(i).Text = CStr(glRegistNum(i))  ' 分布グラフ抵抗数
    '    '' '' ''        Next

    '    '' '' ''    End Sub
    '    '' '' ''#End Region

    '    '' '' ''#Region "グラフクリック時処理"
    '    '' '' ''    '''=========================================================================
    '    '' '' ''    '''<summary>グラフクリック時処理</summary>
    '    '' '' ''    '''<remarks></remarks>
    '    '' '' ''    '''=========================================================================
    '    '' '' ''    Public Sub lblGraphClick_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblGraphClick.Click
    '    '' '' ''        ' グラフクリック
    '    '' '' ''        frmDistribution.Show()
    '    '' '' ''    End Sub
    '    '' '' ''#End Region


    End Module