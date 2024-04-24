'==============================================================================
'   Description : Trimming Program
'
'   Copyright(C): TOWA LASERFRONT CORP. 2018
'
'==============================================================================
Option Strict Off
Option Explicit On

Imports System.IO

Imports LaserFront.Trimmer.DllVideo.VideoLibrary
Imports LaserFront.Trimmer.DefWin32Fnc
Imports TrimClassLibrary
Imports LaserFront.Trimmer.DllLaserTeach.ctl_LaserTeach
Imports LaserFront.Trimmer.DllTeach

Module UserBas

#Region "定数定義"
    '-------------------------------------------------------------------------------
    '   DLL定義
    '-------------------------------------------------------------------------------
    '----- WIN32 API -----
    'V2.1.0.0④    Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
    'V2.1.0.0④    Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer

    '-------------------------------------------------------------------------------
    '   使用ファイル
    '   ①トリミングデータファイル(xxxxxx.TXT ← xxxxxxは任意の文字列)
    '       ・例 TRM_DATA.TXT(通常はC:\TRIMDATA\DATAフォルダ下)
    '   ②ロット情報ファイル (C:\TRIMDATA\DATA\Lotproduct.DAT)
    '   ③ユーザ定義ファイル (C:\TRIM\EDIT_DEF_User.ini)
    '   ④ﾃﾞｰﾀ最小値・最大値定義ファイル (C:\TRIM\EDIT_DEF_UserSample.ini)
    '   ⑤ログファイル(トリミング結果ログ)
    '     "LOG_" + ｢ロット番号｣ + ".LOG"
    '-------------------------------------------------------------------------------
    '----- データファイル名 -----
    Public Const cTRIMFILEPATH As String = "C:\TRIMDATA\DATA\"
    Public Const cLOGFILEPATH As String = "C:\TRIMDATA\LOG\"
    Public Const cLOT_FNAME As String = "C:\TRIMDATA\DATA\Lotproduct.DAT"
    'V2.2.2.0①     Public Const cDEF_FNAME As String = "C:\TRIM\DefFunc_UserSL432R0050.ini"
    'V2.2.2.0① Public Const cEDITDEF_FNAME As String = "C:\TRIM\EDIT_DEF_UserSL432R0050.ini"
    Public Const cDEF_FNAME As String = "C:\TRIM\DefFunc_UserSL432R.ini"            'V2.2.2.0① 
    Public Const cEDITDEF_FNAME As String = "C:\TRIM\EDIT_DEF_UserSL432R.ini"       'V2.2.2.0① 
    'Public Const cTEMPLATPATH As String = "C:\TRIM"                    ' Video.OCX用ﾃﾝﾌﾟﾚｰﾄﾌｧｲﾙの保存場所
    Public Const cTEMPLATPATH As String = "C:\TRIM\VIDEO" '             ' Video.OCX用ﾃﾝﾌﾟﾚｰﾄﾌｧｲﾙの保存場所

    '-------------------------------------------------------------------------------
    '   定数データ定義
    '------------------------------------------------------------------------------
    '-----------------------------------------------------------------------------
    Public Const EXTEQU As Short = 3                    ' 抵抗データでのＯＮ／ＯＦＦ最大外部機器数
    Public Const DEV_TIMER As Short = 1000          ' ZWAIT分割用         (1 SEC単位)
    '----- ｷｬﾌﾟｼｮﾝ(Program title) -----
    Public Const cAPPcTITLE As String = "SL432R"
    Public Const cAPPcTITLEcS As String = "SL432R"

    '----- 最大値 -----
    Public Const MAXRNO As Short = 50                   ' MAX抵抗数
    Public Const MAXCTN As Short = 200                  ' MAXカット数 'V1.0.4.3①５０から１００へ拡張　'V2.0.0.0③ 2017/11/20 最大カット数を１００から２００へ変更
    Public Const MAXSCTN As Short = 10                  ' MAXｻｰﾍﾟﾝﾀｲﾝｶｯﾄ数
    Public Const MAXCND As Short = 4                    ' 加工条件数（※FL用）
    Public Const MAXIDX As Short = 5                    ' MAXｲﾝﾃﾞｯｸｽｶｯﾄ数
    Public Const MAXCNO As Short = 200                  ' MAXカット総数(20抵抗×10カット分)
    Public Const MAXGNO As Short = 10                   ' MAXGPIB制御(MAX10台)
    Public Const MAXRGN As Short = MAXRNO               ' MAXパターン登録数
    Public Const MAXBLKX As Short = 50                  ' MAXブロックＸ数
    Public Const MAXBLKY As Short = 50                  ' MAXブロックＹ数
    Public Const MAXBLK As Short = 2500                 ' MAXブロック数(X×Y)
    Public Const MAX_LCUT As Short = 7                  ' Ｌカットのカット数（Ｌ１～Ｌ７）V1.0.4.3③
    Public Const MAX_RETRACECUT As Short = 10           ' リトレースカットのカット数 'V2.0.0.0⑦

    '----- ﾀｲﾑｱｳﾄ値 -----
    Const GTMOUT As Short = 100                         ' GPIB TIMEOUT(×0.1s = 10秒)

    '----- ﾎﾟｰｽﾞﾀｲﾑ -----
    Const PROB_ON_TIM As Short = 300                    ' ﾌﾟﾛｰﾌﾞON後のﾎﾟｰｽﾞ (msec)
    Const PROB_OFF_TIM As Short = 0                     ' ﾌﾟﾛｰﾌﾞOFF前のﾎﾟｰｽﾞ(msec)
    Const REL_ON_TIM As Short = 10                      ' ﾘﾚｰON後のﾎﾟｰｽﾞ    (msec)
    Const REL_OFF_TIM As Short = 0                      ' ﾘﾚｰOFF後のﾎﾟｰｽﾞ   (msec)

    '----- 判定(表示用) -----
    Public Const JG_SP As String = " SKIP"              ' 初期値
    Public Const JG_OK As String = "   OK"              ' 正常
    Public Const JG_IH As String = "IT-HI"              ' 初期判定ｴﾗｰ(ITHI)
    Public Const JG_IL As String = "IT-LO"              ' 初期判定ｴﾗｰ(ITLO)
    Public Const JG_IO As String = "IT-OPEN"            ' 初期判定ｴﾗｰ(IT OPEN)
    Public Const JG_FH As String = "FT-HI"              ' 終了判定ｴﾗｰ(FTHI)
    Public Const JG_FL As String = "FT-LO"              ' 終了判定ｴﾗｰ(FTLO)
    Public Const JG_FO As String = "FT-OPEN"            ' 終了判定ｴﾗｰ(FT OPEN)
    Public Const JG_ER As String = "ERROR"              ' ｴﾗｰ発生(電圧設定等)
    Public Const JG_RS As String = "RESET"              ' RESET指定
    Public Const JG_PT As String = "NG-PT"              ' ﾊﾟﾀｰﾝ認識ｴﾗｰ
    Public Const JG_VA As String = "VA-NG"              ' 再測定変化量エラー'V2.0.0.0②
    Public Const JG_STD As String = "STD-NG"            ' スタンダード抵抗測定エラー 'V2.0.0.0⑮
    Public Const JG_CUTVA As String = "CUT-VA"          ' カット毎の抵抗値変化量判定ＮＧ 'V2.1.0.0①
    ''' <summary>
    ''' 判定結果列挙子定義
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum eJudge As Integer
        JG_SP
        JG_OK
        JG_IH
        JG_IL
        JG_IO
        JG_FH
        JG_FL
        JG_FO
        JG_ER
        JG_RS
        JG_PT
        JG_VA                                           ' 再測定変化量エラー'V2.0.0.0②
        JG_STD                                          ' スタンダード抵抗測定エラー 'V2.0.0.0⑮
        JG_CUTVA                                        ' カット毎の抵抗値変化量判定ＮＧ 'V2.1.0.0①
    End Enum

    ''' <summary>
    ''' 状態表示列挙子定義
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum eDispMode As Integer
        DISP_MODE_INIT
        DISP_MODE_TRIM
        DISP_MODE_REMEAS
        DISP_MODE_RESULT
        DISP_MODE_CONTRO_LMEASURE_RESULT
    End Enum

    '----- その他 -----
    Public Const Z0 As Short = 0
    Public Const Z1 As Short = 1
    Public Const Z2 As Short = 2
    Public Const Z3 As Short = 3
    Public Const Z4 As Short = 4
    Public Const ZOPT As Short = 1                      ' ﾗﾝﾌﾟ照明(0:標準, 1:LED照明)

    '----- カット系関数戻値 -----
    Public Const CUT_END_NOM As Integer = 1             ' 目標値を超えたので終了
    Public Const CUT_END_LMT As Integer = 2             ' 指定移動量までカットしたので終了

    '-------------------------------------------------------------------------------
    ' アッテネータテーブル、温度センサー情報テーブル関連
    '-------------------------------------------------------------------------------
    'V2.1.0.0②↓
    Public Const cLASERPOWER_HEADER As String = "LASERPOWER-"               ' アッテネータテーブルファイルヘッダ
    Public Const cLASERPOWER_PATH As String = "C:\TRIMDATA\LASERPOWER\"     ' アッテネータテーブルファイルフォルダ
    'V2.1.0.0②↑

    'V2.1.0.0③↓
    Public Const MAX_TEMPERATURE As Short = 99                              ' 温度センサー情報テーブル数 'V2.1.0.0③
    Public Const cTEMPERATURE_HEADER As String = "TEMPERATURE-"             ' 温度センサー情報テーブルファイルヘッダ
    Public Const cTEMPERATURE_PATH As String = "C:\TRIMDATA\TEMPERATURE\"   ' 温度センサー情報テーブルファイルフォルダ
    'V2.1.0.0③↑

    'V2.2.0.0⑮↓
    Public Const cPROBEDATA_PATH As String = "C:\TRIMDATA\PROBEDATA\"     ' プローブテーブルファイルフォルダ
    Public Const cPROBEDATA_FILE As String = "PROBEDATA.csv"            ' プローブテーブルファイル
    'V2.2.0.0⑮↑

    'V2.1.0.0②↓DllLaserTeachで宣言 Imports LaserFront.Trimmer.DllLaserTeach.ctl_LaserTeach
    'Public Structure stATTENUATOR_TABLE         ' アッテネータテーブル
    '    Dim No As Integer                       ' 番号
    '    Dim Power As Double                     ' パワー設定
    '    Dim PowerUnit As String                 ' パワー単位
    '    Dim Limit As Double                     ' 範囲
    '    Dim LimitUnit As String                 ' 範囲単位
    '    Dim Rate As Double                      ' 減衰率
    '    Dim RateUnit As String                  ' 減衰率単位
    '    Dim Rotation As Integer                 ' 回転量
    '    Dim FixAtt As Integer                   ' 固定アッテネータ
    '    Dim Comment As String                   ' コメント
    'End Structure
    'V2.1.0.0②↑

    'V2.1.0.0③↓
    Public Structure stTEMPERATURE_TABLE        ' 温度センサー情報テーブル
        Dim No As Integer                       ' 番号
        Dim Title As String                     ' 元素記号
        Dim dTemperatura0 As Double             ' ０℃
        Dim dDaihyouAlpha As Double             ' 代表α値
        Dim dDaihyouBeta As Double              ' 代表β値
        Dim Comment As String                   ' コメント
    End Structure
    'V2.1.0.0③↑

    'V2.2.0.0⑮ ↓
    Public Structure stPROBEDATA_TABLE        ' プローブデータテーブル
        Dim No As Integer                       ' 番号
        Dim ProbeOn As Double                   ' プローブ接触位置
        Dim ProbeOff As Double                  ' プローブ待機位置              'V2.2.0.0⑳
        Dim dTableOffsetX As Double             ' テーブルオフセットX
        Dim dTableOffsetY As Double             ' テーブルオフセットY
        Dim dBPOffsetX As Double                ' BPオフセットX
        Dim dBPOffsetY As Double                ' BPオフセットY
        'V2.2.1.6②↓
        Dim iPP30 As Short                      ' 位置補正モード(0:自動補正モード, 1:手動補正モード, 2:自動+微調)
        Dim iPP31 As Short                      ' 位置補正方法(0:補正なし, 1:補正あり)
        Dim fpp34_x As Double                   ' 補正ポジションオフセットx
        Dim fpp34_y As Double                   ' 補正ポジションオフセットy
        Dim fTheta As Double                    ' θ軸角度
        Dim iPP38 As Short                      ' パターングループ番号
        Dim iPP37_1 As Short                    ' パターン1 テンプレート番号
        Dim fpp32_x As Double                   ' パターン1座標x
        Dim fpp32_y As Double                   ' パターン1座標y
        Dim iPP37_2 As Short                    ' パターン2 テンプレート番号
        Dim fpp33_x As Double                   ' パターン2座標x
        Dim fpp33_y As Double                   ' パターン2座標y
        'V2.2.1.6②↑
        Dim Comment As String                   ' コメント 
    End Structure
    'V2.2.0.0⑮ ↑

    Public Const PROBE_DATA_MAX = 30            ' プローブデータの最大数      'V2.2.1.0①

    '-------------------------------------------------------------------------------
    '   機能選択定義テーブル
    '-------------------------------------------------------------------------------
    '----- 機能選択定義テーブルのｲﾝﾃﾞｯｸｽ -----
    Public Const F_LOAD As Short = 0                    ' LOAD(F1)ボタン
    Public Const F_SAVE As Short = 1                    ' SAVE(F2)ボタン
    Public Const F_EDIT As Short = 2                    ' EDIT(F3)ボタン
    Public Const F_LASER As Short = 3                   ' LASER(F5)ボタン
    Public Const F_LOTCHG As Short = 4                  ' ﾛｯﾄ切替(S-F6)ボタン(特注)
    Public Const F_PROBE As Short = 5                   ' PROBE(F7)ボタン
    Public Const F_TEACH As Short = 6                   ' TEACH(F8)ボタン
    Public Const F_CUTPOS As Short = 7                  ' CUTPOS(S-F8)ボタン
    Public Const F_RECOG As Short = 8                   ' RECOG(F9)ボタン(未使用)
    Public Const F_TX As Short = 9                     ' ＴＸボタン
    Public Const F_TY As Short = 10                     ' ＴＹボタン
    Public Const F_TY2 As Short = 11                    ' TY2ボタン
    'Public Const F_MSTCHK As Short = 9                  ' ﾏｽﾀﾁｪｯｸ(F4)ボタン(特注)
    Public Const MAX_FNCNO As Short = 11                ' 機能選択定義テーブルのｲﾝﾃﾞｯｸｽ数

    '----- 機能選択定義テーブル形式定義 (ユーザ定義ファイル(EDIT_DEF_User.ini)より設定する) -----
    Public Structure FNC_DEF
        Dim iDEF As Short                               ' 機能選択定義(-1:非表示, 0:選択不可, 1:選択可, )
        Dim iPAS As Short                               ' パスワード指定(0:パスワードなし, 1:パスワードあり)
        Dim sCMD As String                              ' ｺﾏﾝﾄﾞ(キー名)
    End Structure
    Public stFNC(MAX_FNCNO) As FNC_DEF                  ' 機能選択定義テーブル

    '-------------------------------------------------------------------------------
    '   フラグ他
    '-------------------------------------------------------------------------------
    Public FlgCan As Short                              ' Cancel Flag
    Public FlgUpd As Short                              ' データ更新 Flag
    Public FlgUpdGPIB As Short                          ' GPIBデータ更新Flag
    Public FlgGPIB As Boolean                           ' GPIB初期化Flag
    Public pbLoadFlg As Boolean                         ' データロードフラグ
    Public gbInitialized As Boolean                     ' True=原点復帰済  , False=原点復帰未
    Public gflgResetStart As Short                      ' True=初期設定済み, False=初期設定済みでない
    Public fStartTrim As Boolean                        ' スタートTRIMフラグ

    '----- ローダ入出力関連(ｵﾌﾟｼｮﾝ) -----
    Public giHostMode As Short                          ' ﾛｰﾀﾞﾓｰﾄﾞ(0:手動ﾓｰﾄﾞ, 1:自動ﾓｰﾄﾞ)
    Public Const cHOSTcMODEcAUTO As Short = 1           '  1:自動ﾓｰﾄﾞ
    Public Const cHOSTcMODEcMANUAL As Short = 0         '  0:手動ﾓｰﾄﾞ
    Public gbHostConnected As Boolean                   ' ホスト接続状態(True=接続(ﾛｰﾀﾞ有), False=未接続(ﾛｰﾀﾞ無))
    Public giHostRun As Short                           ' ﾛｰﾀﾞ動作中(0:停止, 1:動作中)

    '----- ファイル番号 -----
    Public fNum As Short                                ' ﾄﾘﾐﾝｸﾞﾃﾞｰﾀﾌｧｲﾙ番号
    Public Logno As Short                               ' ﾛｸﾞﾌｧｲﾙ番号

    '----- ｱﾌﾟﾘﾓｰﾄﾞ ----- (注)OcxSystem定義と一致させる必要有り
    Public giAppMode As Short                           ' ｱﾌﾟﾘﾓｰﾄﾞ

    Public Const APP_MODE_IDLE As Short = 0             ' トリマ装置アイドル中
    Public Const APP_MODE_LOAD As Short = 1             ' ファイルロード
    Public Const APP_MODE_SAVE As Short = 2             ' ファイルセーブ
    Public Const APP_MODE_EDIT As Short = 3             ' データ編集     
    Public Const APP_MODE_LASER As Short = 5            ' レーザー調整  
    Public Const APP_MODE_LOTCHG As Short = 6           ' ロット切替    
    Public Const APP_MODE_DATASET As Short = 6          ' データ設定
    Public Const APP_MODE_PROBE As Short = 7            ' プローブ      
    Public Const APP_MODE_TEACH As Short = 8            ' ティーチング  
    Public Const APP_MODE_RECOG As Short = 9            ' パターン登録(θ補正用)  
    Public Const APP_MODE_EXIT As Short = 10            ' 終了 　　　　 
    Public Const APP_MODE_TRIM As Short = 11            ' トリミング中
    Public Const APP_MODE_TRIM_AUTO As Short = 111      ' 自動運転のトリミング中
    Public Const APP_MODE_CUTPOS As Short = 12          ' カット位置補正
    Public Const APP_MODE_PROBE2 As Short = 13          ' プローブ2     
    Public Const APP_MODE_LOGGING As Short = 14         ' ロギング
    Public Const APP_MODE_FINEADJ As Short = 16         ' 一時停止画面
    Public Const APP_MODE_CARIB_REC As Short = 17       ' 画像登録(キャリブレーション補正用)【外部カメラ】
    Public Const APP_MODE_CUTREVIDE As Short = 18       ' カット位置補正【外部カメラ】
    Public Const APP_MODE_AUTO As Short = 50            ' V1.2.0.0④自動運転
    Public Const APP_MODE_BLOCK_RECOG As Short = 70     ' ブロック内の２点補正機能
    Public Const APP_MODE_VACCUME_CHECK As Short = 71   ' AbsVaccume()で使用、バキュームチェックモード
    Public Const APP_MODE_LOTCHANGE As Short = 72       ' ロット切替
    Public Const APP_MODE_LOTNO As Short = APP_MODE_LOTCHANGE           ' ロット番号設定中

    Public Const APP_MODE_TX As Short = 41              ' TXティーチング
    Public Const APP_MODE_TY As Short = 42              ' TYティーチング
    '----- frmResetの処理モード -----
    Public gMode As Short                               '  0 : 原点復帰他

    ' ﾄﾘﾐﾝｸﾞ用ポーズ時間1～3(1:最初の抵抗用, 2:偶数抵抗用, 3:奇数抵抗用)
    Public glWTimeT(3) As Integer                       ' ﾄﾘﾐﾝｸﾞ用ポーズ時間ms1～3(1:2000ms, 2:   0ms, 3:   0ms)
    Public glWTimeM(3) As Integer                       ' 測定用ポーズ時間ms1～3  (1:5000ms, 2:4000ms, 3: 300ms)

    Public bDebugLogOut As Boolean = False              ' デバッグログの出力有無
    Public bNgCutDebugLogOut As Boolean = False         ' ＮＧカット用デバッグログの出力有無    'V1.2.0.2
    Public bCutVariationDebugLogOut As Boolean = False  ' カット毎の抵抗値変化量判定機能用デバッグログの出力有無    'V2.1.0.0①
    Public giAutoOperationDebugLogOut As Integer = 0    ' ロット処理系のデバッグログ出力用      ''V2.2.1.3②

    Public bRelayBoard As Boolean = True               ' 低熱起電力リレーボード２ 'V2.0.0.0⑬
    Public bPowerOnOffUse As Boolean = True            ' 外部機器電源ON,OFF制御   'V2.0.0.0②
    Public giBlueCrossDisable As Integer = 0           ' ティーチング画面で水色クロスラインを非表示にする      'V2.2.0.0②
    Public giMouseClickMove As Integer = 0             ' 一時停止画面で画面クリックしたときに移動する動作の有効無効      'V2.2.0.0④
    Public giLoaderType As Integer = 0                 ' TLF製ローダ対応    'V2.2.0.0⑤
    Public giCutStop As Integer = 0                    ' カット毎停止機能   'V2.2.0.0⑥
    Public giClcleStop As Integer = 0                  ' カット毎停止機能   'V2.2.0.0⑦
    Public gisupplyMgNum As Integer                    ' 処理中供給マガジン番号
    Public gisupplyMgStepNum As Integer                ' 処理中供給マガジン段数
    Public gistoreMgNum As Integer                     ' 処理中収納マガジン番号
    Public gistoreMgStepNum As Integer                 ' 処理中収納マガジン段数（次置く段数）
    Public giTxtLogType As Integer = 1                 ' 画面ログ出力タイプ    'V2.2.0.0⑤
    Public giTablePosUpd As Integer = 0                ' パターン登録画面でパターン登録座標を変更可能とする 
    Public giLaserOffMode As Integer = 0               ' レーザOFFモード状態
    Public giRecogPointCorrLine As Integer = 0         ' カット位置補正基準点指定   ' V2.2.1.2①

    '----- ローカル変数 -----
    Private gLastDScanMode As Integer = -1
    Private gintPRH As Integer = -1
    Private gintPRL As Integer = -1
    Private gintPRG As Integer = -1

    Private giMultiMeter As Integer = -1                ' マルチメータの設定レンジ
    Private gisCutPosExecute As Boolean                 ' カット補正有無し（抵抗データの設定（stPTN(i).PtnFlg）から求める）
    Private glCutPosTimes As Integer = 1                ' ０：補正しない　１：毎回補正　２以上：指定回数おきに補正を実施
    Private glCutPosCounter As Integer                  ' 補正回数カウンター
    Private gisCutPosExecuteAutoNG As Boolean           ' カット補正自動判定有無し（抵抗データの設定（stPTN(i).PtnFlg）から求める）

    Public ObjGazou As Process = Nothing               ' Gazou Processオブジェクト

#End Region

#Region "トリミングデータ形式定義"
    '-------------------------------------------------------------------------------
    '   トリミングデータ形式定義
    '-------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------
    '   ユーザ・データ
    '-------------------------------------------------------------------------------
    Public Const MAX_RES_USER As Short = 5             ' ユーザ定義の抵抗数最大値

    Public Structure USER_DATA
        Dim iTrimType As Integer                        ' 製品種別
        Dim sLotNumber As String                        ' ロット番号
        Dim sOperator As String                         ' オペレータ名
        Dim sPatternNo As String                        ' パターンＮｏ．
        Dim sProgramNo As String                        ' プログラムＮｏ．
        Dim iTrimSpeed As Integer                       ' トリミング速度
        Dim iLotChange As Integer                       ' ロット終了条件
        Dim lLotEndSL As Long                           ' ロット処理枚数
        Dim lCutHosei As Long                           ' カット位置補正頻度
        Dim lPrintRes As Long                           ' ロット終了時印刷素子数
        Dim iTempResUnit As Integer                     ' 温度センサー 抵抗レンジ 1:Ω, 2:KΩ
        Dim iTempTemp As Integer                        ' 参照温度	１：０℃ または ２：２５℃ 'V2.0.0.0⑪温度を直接指定に変更
        Dim dStandardRes0 As Double                     ' 標準抵抗値０℃	0.01～100M　'V2.0.0.0⑪未使用化
        Dim dStandardRes25 As Double                    ' 標準抵抗値２５℃	0.01～100M　'V2.0.0.0⑪未使用化
        'V2.0.0.0⑪↓
        Dim dTemperatura0 As Double                     ' ０℃　　　'V2.0.0.0⑪
        Dim dDaihyouAlpha As Double                     ' 代表α値　'V2.0.0.0⑪
        Dim dDaihyouBeta As Double                      ' 代表β値　'V2.0.0.0⑪
        Dim dAlpha As Double                            ' α値　　　'V2.0.0.0⑪
        Dim dBeta As Double                             ' β値　　　'V2.0.0.0⑪
        'V2.0.0.0⑪↑
        Dim iTempSensorInfNoDaihyou As Integer          ' 代表β値温度センサー情報一元管理選択番号 V2.1.0.0③  
        Dim iTempSensorInfNoStd As Integer              ' STD温度センサー情報一元管理選択番号 V2.1.0.0③ 
        Dim dResTempCoff As Double                      ' 抵抗温度係数
        Dim dFinalLimitHigh As Double                   ' ファイナルリミット　Hight[%]
        Dim dFinalLimitLow As Double                    ' ファイナルリミット　Lo[%]
        Dim dRelativeHigh As Double                     ' 相対値リミット　Hight[%]
        Dim dRelativeLow As Double                      ' 相対値リミット　Lo[%]
        ' 以下は、抵抗数分（１～５）
        <VBFixedArray(MAX_RES_USER)> Dim iResUnit() As Integer          ' 温度センサー 抵抗レンジ 1:Ω, 2:KΩ
        <VBFixedArray(MAX_RES_USER)> Dim dNomCalcCoff() As Double       ' 補正値（ノミナル値算出係数）
        <VBFixedArray(MAX_RES_USER)> Dim dTargetCoff() As Double        ' 目標値算出係数
        <VBFixedArray(MAX_RES_USER)> Dim dTargetCoffJudge() As Double   ' 判定用目標値算出係数 'V2.1.0.0③
        <VBFixedArray(MAX_RES_USER)> Dim iChangeSpeed() As Integer      ' 測定速度を変更するカットNo.
        <VBFixedArray(MAX_RES_USER)> Dim dItVal() As Double             ' [結果]ＩＴ測定値
        <VBFixedArray(MAX_RES_USER)> Dim dFtVal() As Double             ' [結果]ＦＴ測定値
        <VBFixedArray(MAX_RES_USER)> Dim dDev() As Double               ' [結果]ＤＥＶ値

        'V2.0.0.0②↓
        Dim dRated As Double                            ' 定格
        Dim dMagnification As Double                   ' 定格電圧の倍率


        Dim dResNumber As Integer                      ' 抵抗個数
        Dim dCurrentLimit As Double                    ' 電流制限
        Dim dAppliedSecond As Double                   ' 印加秒数
        Dim dVariation As Double                      ' 変化量
        'V2.0.0.0②↑
        Dim intClampVacume As Short                     'V2.0.0.0⑬ クランプと吸着の有り無し
        Dim NgJudgeRate As Double                       'V2.0.0.1③　トリミングNG信号を出力するＮＧの比率


        'この構造体のインスタンスを初期化するには、"Initialize" を呼び出さなければなりません。
        Public Sub Initialize()
            ReDim iResUnit(MAX_RES_USER)
            ReDim dNomCalcCoff(MAX_RES_USER)
            ReDim dTargetCoff(MAX_RES_USER)
            ReDim dTargetCoffJudge(MAX_RES_USER)     ' 目標値算出係数 'V2.1.0.0③
            ReDim iChangeSpeed(MAX_RES_USER)
            ReDim dItVal(MAX_RES_USER)
            ReDim dFtVal(MAX_RES_USER)
            ReDim dDev(MAX_RES_USER)
        End Sub
    End Structure
    Public stUserData As USER_DATA                          ' ユーザ・データ

    Private gdNOMforStatistical(MAX_RES_USER) As Double  'V2.0.0.0⑨

    '-------------------------------------------------------------------------------
    '   プレートデータ
    '-------------------------------------------------------------------------------
    Public Structure PLATE_DATA
        Dim Pnx As Short                                ' プレート数x = 1
        Dim Pny As Short                                ' プレート数y = 1
        Dim Pivx As Double                              ' プレートインターバルx(mm) = 0
        Dim Pivy As Double                              ' プレートインターバルy(mm) = 0
        Dim BNX As Short                                ' ブロック数x = 1
        Dim BNY As Short                                ' ブロック数y = 1
        Dim zsx As Double                               ' ブロック(抵抗)サイズx(mm)
        Dim zsy As Double                               ' ブロック(抵抗)サイズy(mm)
        Dim ADJX As Double                              ' アジャスト位置X(mm)
        Dim ADJY As Double                              ' アジャスト位置Y(mm)
        Dim z_xoff As Double                            ' トリムポジションオフセットX(mm)
        Dim z_yoff As Double                            ' トリムポジションオフセットY(mm)
        Dim Z_ZOFF As Double                            ' Z PROBE OFF位置(mm)
        Dim Z_ZON As Double                             ' Z PROBE ON 位置(mm)
        Dim Z2_ZOFF As Double                           ' Z2 PROBE OFF位置(mm)
        Dim Z2_ZON As Double                            ' Z2 PROBE ON 位置(mm)
        Dim BPOX As Double                              ' BP Offset X(mm)
        Dim BPOY As Double                              ' BP Offset Y(mm)
        Dim PrbRetry As Short                           ' プローブリトライ(1:有, 0:無)
        Dim RCount As Short                             ' 抵抗数
        Dim GCount As Short                             ' GPIB制御数
        Dim PtnCount As Short                           ' パターン登録数
        Dim TeachBlockX As Short                        ' ティーチングブロックＸ ###1040
        Dim TeachBlockY As Short                        ' ティーチングブロックＹ ###1040
        Dim StageSpeedY As Long                         ' ステージ・スピードＹ   ###1040
        Dim dblChipSizeXDir As Double                   'V1.2.0.0① チップサイズサイズx(mm)
        Dim dblChipSizeYDir As Double                   'V1.2.0.0① チップサイズサイズy(mm)
        Dim dblStepOffsetXDir As Double                 ' ステップオフセット量X
        Dim dblStepOffsetYDir As Double                 ' ステップオフセット量Y
        Dim DistributionResNo As Integer                'V2.0.0.0⑨ 分布図の表示抵抗番号統計処理用ダミー
        'Dim PtnFlg As Short                            ' パターン認識(0:無し, 1:有り, 2:手動) ← 抵抗毎に持つ
        Dim dblStdMagnification As Double               ' デジタルカメラ倍率   'V2.2.0.0②
        Dim ProbNo As Integer                           ' ﾌﾟﾛｰﾌﾞNo  V2.2.0.0⑮
    End Structure
    Public stPLT As PLATE_DATA                          ' プレートデータ

    '-------------------------------------------------------------------------------
    '   パワー制御用データ(パワー制御結果取得用)
    '-------------------------------------------------------------------------------
    Public Structure POWER_DATA
        Dim intQR As Short                              ' Qレート (x100Hz)(0.1KHz)
        Dim XOFF As Double                              ' XY TABLE Offset X(mm)　※未使用
        Dim YOFF As Double                              ' XY TABLE Offset Y(mm)　※未使用
        Dim dblFullPower As Double                      ' MAXパワー[W]
        Dim dblspecPower As Double                      ' 設定パワー[W]
        Dim dblMeasPower As Double                      ' 測定したパワー[W]
        Dim dblRotPar As Double                         ' 減衰率(%)
        Dim dblRotAtt As Double                         ' ロータリーアッテネータの回転量(0-FFF)
        Dim iFixAtt As Short                            ' 固定アッテネータのON/OFF(0:OFF,1:ON)
        Dim iErrAtt As Short                            ' エラー値
        '                                               '   0:正常
        '                                               '   1:設定不可(10回リトライしてもNG)
        '                                               '   2:最大ﾊﾟﾜｰ < 指定ﾊﾟﾜｰ
        '                                               '   3:最大ﾊﾟﾜｰ * 最大減衰率 > 指定ﾊﾟﾜｰ
        '                                               '   4:USBｵｰﾌﾟﾝｴﾗｰ
        '                                               '   5:USB入力ｴﾗｰ
        '                                               '  10:手動設定
        '                                               ' ※「10:手動設定」の場合は、
        '                                               ' ・ロータリーアッテネータの回転量(0-FFF)と
        '                                               ' ・固定アッテネータのON/OFF(0:OFF,1:ON)が返る｡
        '                                               '  MAXパワー､指定パワー､測定したパワーは0が返る｡
        Dim iTrimAtt As Short                           ' ＡＴＴデータをトリミングデータから取得する時１従来通りの時は０
        Dim intPowerAdjustMode As Short                 ' ﾊﾟﾜｰ調整ﾓｰﾄﾞ
        Dim dblPowerAdjustTarget As Double              ' 調整目標ﾊﾟﾜｰ
        Dim dblPowerAdjustQRate As Double               ' ﾊﾟﾜｰ調整Qﾚｰﾄ
        Dim dblPowerAdjustToleLevel As Double           ' ﾊﾟﾜｰ調整許容範囲
        Dim iAttNo As Integer                           ' アッテネータNo.（0:指定無）'V2.1.0.0②
    End Structure
    Public stLASER As POWER_DATA                        ' パワー制御用データ

    '-------------------------------------------------------------------------------
    '   カットデータ
    '-------------------------------------------------------------------------------
    Public Structure Cut_Info                               ' カットデータ情報
        Dim intCUT As Short                                 ' ｶｯﾄ方法(1:ﾄﾗｯｷﾝｸﾞ, 2:ｲﾝﾃﾞｯｸｽ(STｶｯﾄ/Lｶｯﾄのみ))
        Dim intCTYP As Short                                ' ｶｯﾄ形状(1:ｽﾄﾚｰﾄ, 2:Lｶｯﾄ, 3:ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ)
        Dim intNum As Short                                 ' ｶｯﾄ本数(ｻｰﾍﾟﾝﾀｲﾝｶｯﾄのみ)
        Dim dblSTX As Double                                ' ﾄﾘﾐﾝｸﾞ ｽﾀｰﾄ点 X
        Dim dblSTY As Double                                ' ﾄﾘﾐﾝｸﾞ ｽﾀｰﾄ点 Y
        Dim dblSX2 As Double                                ' ﾄﾘﾐﾝｸﾞ ｽﾀｰﾄ点2 X(ｻｰﾍﾟﾝﾀｲﾝｶｯﾄのみ)　'V1.0.4.3③リトレースのオフセットに使用
        Dim dblSY2 As Double                                ' ﾄﾘﾐﾝｸﾞ ｽﾀｰﾄ点2 Y(ｻｰﾍﾟﾝﾀｲﾝｶｯﾄのみ)　'V1.0.4.3③リトレースのオフセットに使用
        Dim dblCOF As Double                                ' ｶｯﾄｵﾌ(%)→目標値に対するｵﾌｾｯﾄ値(目標値×(1＋ｶｯﾄｵﾌ/100))
        '                                                   ' 例) -1%なら本来の目標値の99%となる
        Dim intTMM As Short                                 ' モード(0:高速(コンパレータ非積分モード), 1:高精度(積分モード))　※インデックス時使用する。
        Dim intMType As Short                               ' 測定種別(0=内部測定, 1=外部測定)
        '                                                   ' ｶｯﾄ方法 = ｲﾝﾃﾞｯｸｽﾄﾘﾐﾝｸﾞ時有効
        Dim intQF1 As Short                                 ' Qレート(0.1KHz)
        Dim dblV1 As Double                                 ' ﾄﾘﾑ速度(mm/s)
        Dim intQF2 As Short                                 ' V1.0.4.3③ストレートカット・リトレースのQレート(0.1KHz)に使用
        Dim dblV2 As Double                                 ' V1.0.4.3③ストレートカット・リトレースのトリム速度(mm/s)に使用
        Dim dblDL2 As Double                                ' 第2のｶｯﾄ長(ﾘﾐｯﾄｶｯﾄ量mm(Lﾀｰﾝ前))
        Dim dblDL3 As Double                                ' 第3のｶｯﾄ長(Lｶｯﾄ時のﾘﾐｯﾄｶｯﾄ量mm(Lﾀｰﾝ後), ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ時の移動ﾋﾟｯﾁ)
        Dim intANG As Short                                 ' ｶｯﾄ方向1(90°単位　0°～360°) 0:+X, 180:-X, 90:+Y, 270:-Y
        Dim intANG2 As Short                                ' ｶｯﾄ方向2(90°単位　0°～360°) Lｶｯﾄ時のLﾀｰﾝ後の移動方向, ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ時の移動方向
        Dim dblLTP As Double                                ' Lﾀｰﾝ ﾎﾟｲﾝﾄ(%)
        'V2.1.0.0①↓ カット毎の抵抗値変化量判定機能追加
        Dim iVariationRepeat As Integer                     ' リピート有無
        Dim iVariation As Integer                           ' 判定有無
        Dim dRateOfUp As Double                             ' 上昇率
        Dim dVariationLow As Double                         ' 下限値
        Dim dVariationHi As Double                          ' 上限値
        'V2.1.0.0①↑
        'V2.2.0.0②↓Uカット用パラメータ 
        Dim dUCutL1 As Double                               ' L1
        Dim dUCutL2 As Double                               ' L2
        Dim intUCutQF1 As Short                             ' L1用Qレート
        Dim dblUCutV1 As Double                             ' UカットL1時リム速度(mm/s)に使用
        Dim intUCutANG As Short                             ' Uカット時ｶｯﾄ方向1(90°単位　0°～360°) 0:+X, 180:-X, 90:+Y, 270:-Y
        Dim dblUCutTurnP As Double                          ' Ujカット時ターンポイント 
        Dim intUCutTurnDir As Short                         ' Ujカット時ターン方向
        Dim dblUCutR1 As Double                             ' UカットR1指定
        Dim dblUCutR2 As Double                             ' UカットR2指定
        'V2.2.0.0②↑

        <VBFixedArray(MAXCND)> Dim intCND() As Short        ' 加工条件番号
        <VBFixedArray(MAXIDX)> Dim intIXN() As Short        ' ｲﾝﾃﾞｯｸｽｶｯﾄ数1～5  ※0指定で最終とみなす
        <VBFixedArray(MAXIDX)> Dim dblDL1() As Double       ' ｶｯﾄ長1～5(ｲﾝﾃﾞｯｸｽﾋﾟｯﾁ1～5)
        <VBFixedArray(MAXIDX)> Dim lngPAU() As Integer      ' ピッチ間ポーズ時間1～5(ms)
        <VBFixedArray(MAXIDX)> Dim dblDEV() As Double       ' 誤差(%)1～5(ｲﾝﾃﾞｯｸｽﾋﾟｯﾁ1～5) 大→小
        <VBFixedArray(MAXIDX)> Dim intIXMType() As Short    ' 測定機器0～5(0:内部測定　1～:外部機器)
        <VBFixedArray(MAXIDX)> Dim intIXTMM() As Short      ' 測定モード(0:高速　1:高精度)
        Dim cFormat As String                               ' ###1042① 文字データ

        'V2.2.1.7① ↓
        Dim cMarkFix As String                               ' 印字固定部
        Dim cMarkStartNum As String                          ' 開始番号
        Dim intMarkRepeatCnt As Short                        ' 重複回数
        'V2.2.1.7① ↑

        'V1.0.4.3③ ADD ↓
        <VBFixedArray(MAX_LCUT)> Dim dCutLen() As Double    ' カット長１～７　リターン時も使用
        <VBFixedArray(MAX_LCUT)> Dim dQRate() As Double     ' Ｑレート１～７　リターン時も使用
        <VBFixedArray(MAX_LCUT)> Dim dSpeed() As Double     ' 速度１～７
        <VBFixedArray(MAX_LCUT)> Dim dAngle() As Double     ' 角度１～７
        <VBFixedArray(MAX_LCUT)> Dim dTurnPoint() As Double ' ターンポイント１～６
        'V1.0.4.3③ ADD ↑
        'V2.0.0.0⑦ ADD ↓
        Dim intRetraceCnt As Short                          ' リトレースカット本数
        <VBFixedArray(MAX_RETRACECUT)> Dim dblRetraceOffX() As Double   ' リトレースのオフセットＸ
        <VBFixedArray(MAX_RETRACECUT)> Dim dblRetraceOffY() As Double   ' リトレースのオフセットＹ
        <VBFixedArray(MAX_RETRACECUT)> Dim dblRetraceQrate() As Double  ' ストレートカット・リトレースのQレート(0.1KHz)に使用
        <VBFixedArray(MAX_RETRACECUT)> Dim dblRetraceSpeed() As Double  ' ストレートカット・リトレースのトリム速度(mm/s)に使用
        'V2.0.0.0⑦ ADD ↑



        'この構造体のインスタンスを初期化するには、"Initialize" を呼び出さなければなりません。
        Public Sub Initialize()
            ReDim intCND(MAXCND)
            ReDim intIXN(MAXIDX)
            ReDim dblDL1(MAXIDX)
            ReDim lngPAU(MAXIDX)
            ReDim dblDEV(MAXIDX)
            ReDim intIXMType(MAXIDX)
            ReDim intIXTMM(MAXIDX)
            'V1.0.4.3③ ADD ↓
            ReDim dCutLen(MAX_LCUT)
            ReDim dQRate(MAX_LCUT)
            ReDim dSpeed(MAX_LCUT)
            ReDim dAngle(MAX_LCUT)
            ReDim dTurnPoint(MAX_LCUT)
            'V1.0.4.3③ ADD ↑
            'V2.0.0.0⑦ ADD ↓
            ReDim dblRetraceOffX(MAX_RETRACECUT)        ' カット長１～７　リターン時も使用
            ReDim dblRetraceOffY(MAX_RETRACECUT)        ' Ｑレート１～７　リターン時も使用
            ReDim dblRetraceQrate(MAX_RETRACECUT)       ' ストレートカット・リトレースのQレート(0.1KHz)に使用
            ReDim dblRetraceSpeed(MAX_RETRACECUT)       ' ストレートカット・リトレースのトリム速度(mm/s)に使用
            'V2.0.0.0⑦ ADD ↑
        End Sub
    End Structure

    'V2.2.0.0⑬↓
    Public Structure CutInfoTeachIF
        Dim intidno As Short 'インデックス
        Dim intrno As Short '抵抗番号
        Dim intcno As Short 'カットＮｏ．
        Dim intctype As Short 'カットタイプ
        Dim intcutdir As Short 'カット方向
        Dim dblx As Double 'カットスタート位置Ｘ
        Dim dbly As Double 'カットスタート位置Ｙ
        Dim dblL1 As Double 'カット長1
        Dim dblR1 As Double 'R1
        Dim dblTpt As Double 'ターンポイント
        Dim dblL2 As Double 'カット長２
        Dim dblR2 As Double 'R2
        Dim dblL3 As Double 'カット長３
        Dim intidx As Short 'インッデクス
        Dim intMode As Short '測定モード
        Dim intAngl As Short '角度
        Dim dblPT As Double 'ﾋﾟｯﾁ
        Dim intStp As Short 'ｽﾃｯﾌﾟ方向
        Dim intCnt As Short '本数
        Dim dblZom As Double '倍率
        Dim intClen As Short '文字列長
        Dim iTurnDir As Short 'L2ターン方向 (1:CW, 2:CCW)
        Dim dblIxLimit As Double    ' インデックスカット時のリミット長      ''V6.1.2.0➀
        Dim dblDstFrm2ndCut As Double   'フックカット＋スキャンカット時の第２カットからの距離 'V6.5.1.0①
        Dim dblL4 As Double 'カット長4
        Dim dblL5 As Double 'カット長5
        Dim dblL6 As Double 'カット長6
        Dim dblL7 As Double 'カット長7

        Dim dblspd1 As Double       ' L1カット時速度
        Dim dblspd2 As Double       ' L2カット時速度
        Dim dblspd3 As Double       ' L2カット時速度
        Dim dblspd4 As Double       ' L2カット時速度
        Dim dblspd5 As Double       ' L2カット時速度
        Dim dblspd6 As Double       ' L2カット時速度
        Dim dblspd7 As Double       ' L2カット時速度
        Dim dblQrate1 As Double     ' L1カット時Qレート
        Dim dblQrate2 As Double     ' L2カット時Qレート
        Dim dblQrate3 As Double     ' L2カット時Qレート
        Dim dblQrate4 As Double     ' L2カット時Qレート
        Dim dblQrate5 As Double     ' L2カット時Qレート
        Dim dblQrate6 As Double     ' L2カット時Qレート
        Dim dblQrate7 As Double     ' L2カット時Qレート
        Dim dAngle1 As Double       ' L6カット時角度
        Dim dAngle2 As Double       ' L6カット時角度
        Dim dAngle3 As Double       ' L6カット時角度
        Dim dAngle4 As Double       ' L6カット時角度
        Dim dAngle5 As Double       ' L6カット時角度
        Dim dAngle6 As Double       ' L6カット時角度
        Dim dAngle7 As Double       ' L6カット時角度
    End Structure
    'V2.2.0.0⑬↑

    '----- ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ情報形式定義 -----
    Public Structure Sp_Cut_Info                        ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ情報
        <VBFixedArray(MAXSCTN)> Dim dblSTX() As Double  ' ﾄﾘﾐﾝｸﾞ ｽﾀｰﾄ点 X
        <VBFixedArray(MAXSCTN)> Dim dblSTY() As Double  ' ﾄﾘﾐﾝｸﾞ ｽﾀｰﾄ点 Y

        'この構造体のインスタンスを初期化するには、"Initialize" を呼び出さなければなりません。
        Public Sub Initialize()
            ReDim dblSTX(MAXSCTN)
            ReDim dblSTY(MAXSCTN)
        End Sub

    End Structure

    '-------------------------------------------------------------------------------
    '   抵抗データ
    '-------------------------------------------------------------------------------
    Public Structure Reg_Info                           ' 抵抗データ情報
        Dim strRNO As String                            ' 抵抗名
        Dim strTANI As String                           ' 単位("V","Ω" 等)
        Dim intSLP As Short                             ' 電圧変化スロープ(1:+V, 2:-V, 4:抵抗, 5:電圧測定のみ, 6:抵抗測定のみ 7:NGﾏｰｷﾝｸﾞ)
        Dim lngRel As Integer                           ' リレービット
        Dim dblNOM As Double                            ' トリミング目標値
        Dim dblITL As Double                            ' 初期判定下限値 (ITLO)
        Dim dblITH As Double                            ' 初期判定上限値 (ITHI)
        Dim dblFTL As Double                            ' 終了判定下限値 (FTLO)
        Dim dblFTH As Double                            ' 終了判定上限値 (FTHI)
        Dim intMode As Short                            ' 判定モード(0:比率(%), 1:数値(絶対値))
        Dim intMeasMode As Short                        ' 測定モード(0:なし, 1:ITのみ 2:FTのみ 3:IT,FT両方)
        Dim intTMM1 As Short                            ' モード(0:高速(コンパレータ非積分モード), 1:高精度(積分モード))
        Dim intPRH As Short                             ' ハイ側プローブ番号(High Probe No.)
        Dim intPRL As Short                             ' ロー側プローブ番号(Low Probe No.)
        Dim intPRG As Short                             ' ガードプローブ番号(Gaude probe No.)
        Dim intMType As Short                           ' 測定種別(0=内部測定, 1=外部測定)
        '                                               ' x0モードのIT/FT時有効(電圧ﾄﾘﾐﾝｸﾞ)
        '                                               ' x3モードの電圧測定時有効
        Dim intTNN As Short                             ' カット数(1～5)
        Dim bPattern As Boolean                         'V1.2.0.0③ カット位置補正の判定 True：OK False:NG


        '配列STCUTで各要素を初期化する必要があります。
        <VBFixedArray(MAXCTN)> Dim STCUT() As Cut_Info ' カットデータ情報

        <VBFixedArray(EXTEQU)> Dim intOnExtEqu() As Short        ' ＯＮ機器１～３
        <VBFixedArray(EXTEQU)> Dim intOffExtEqu() As Short       ' ＯＦＦ機器１～３
        Dim intReMeas As Short                          '再測定回数
        Dim intReMeas_Time As Short                     '測定－再測定までのポーズ時間（ｍｓ）
        Dim intITReMeas As Short                        'V2.0.0.0⑧ イニシャル抵抗再測定回数(IT測定回数)
        Dim intFTReMeas As Short                        'V2.0.0.0⑧ ファイナル抵抗再測定回数(FT測定回数)
        Dim intCircuitNo As Short                       'V2.0.0.0⑩ サーキット番号

        'この構造体のインスタンスを初期化するには、"Initialize" を呼び出さなければなりません。
        Public Sub Initialize()
            ReDim STCUT(MAXCTN)
            ReDim intOnExtEqu(EXTEQU)                   ' ＯＮ機器１～３
            ReDim intOffExtEqu(EXTEQU)                  ' ＯＦＦ機器１～３
        End Sub
    End Structure
    '配列 stREG で各要素を初期化する必要があります。 
    Public stREG(MAXRNO) As Reg_Info                    ' 抵抗データ情報

    '-------------------------------------------------------------------------------
    '   GPIB設定用データ
    '-------------------------------------------------------------------------------
    Public Structure GPIB_DATA
        Dim strGNAM As String                           ' 名称
        Dim intGAD As Short                             ' GPIB ADRESS
        Dim intDLM As Short                             ' DELIMITER(0:CRLF, 1:CR, 2:LF, 3:なし)
        'V2.0.0.0④        Dim strCCMD As String                           ' 設定コマンド
        Dim strCCMD1 As String                          ' 設定コマンド'V2.0.0.0④
        Dim strCCMD2 As String                          ' 設定コマンド'V2.0.0.0④
        Dim strCCMD3 As String                          ' 設定コマンド'V2.0.0.0④
        Dim strCON As String                            ' ＯＮコマンド
        Dim strCOFF As String                           ' ＯＦＦコマンド
        Dim lngPOWON As Integer                         ' ON後のﾎﾟｰｽﾞ時間(ms)
        Dim lngPOWOFF As Integer                        ' OFF後のﾎﾟｰｽﾞ時間(ms)
        Dim strCTRG As String                           ' トリガーコマンド
    End Structure
    Public stGPIB(MAXGNO) As GPIB_DATA                  ' GPIB設定用データ
    Public DMM As Short                                 ' ﾃﾞｼﾞﾀﾙﾏﾙﾁﾒｰﾀ Index

    '-------------------------------------------------------------------------------
    '   パターン登録データ(カット位置補正用)
    '-------------------------------------------------------------------------------
    Public Structure Ptn_Info                           ' パターン登録情報
        Dim PtnFlg As Short                             ' パターン認識(0:無し,1:有り, 2:手動)   'V1.0.4.3⑥　[3:自動ＮＧ判定あり]　追加
        Dim intGRP As Short                             ' パターン登録ｸﾞﾙｰﾌﾟ番号(1-999)
        Dim intPTN As Short                             ' パターン登録番号(1-50)
        Dim dblPosX As Double                           ' パターン位置X(補正位置ティーチング用)
        Dim dblPosY As Double                           ' パターン位置Y(補正位置ティーチング用)
        Dim dblDRX As Double                            ' ズレ量保存ワークX
        Dim dblDRY As Double                            ' ズレ量保存ワークY
    End Structure
    Public stPTN(MAXRGN) As Ptn_Info                    ' パターン登録情報
    Public gTblPtn(MAXRGN) As Short                     ' パターン認識結果(1 ORG) 0:OK, 1:NG
    Private gcPtnCorrval(MAXRGN) As String              ' パターンマッチ相関値他情報( "NONE","SAME",MANUAL")

    '----- ﾊﾟﾀｰﾝ登録データ(θ調整用) -----
    Public dblCorrectX As Double                        ' θ補正時のXYﾃｰﾌﾞﾙずれ量X(mm) ※ThetaCorrection()で設定
    Public dblCorrectY As Double                        ' θ補正時のXYﾃｰﾌﾞﾙずれ量Y(mm)
    Public giTemplateGroup As Short                     ' ﾃﾝﾌﾟﾚｰﾄｸﾞﾙｰﾌﾟ番号(1～999)

    '-------------------------------------------------------------------------------
    '   θ調整関連(ｵﾌﾟｼｮﾝ)
    '-------------------------------------------------------------------------------
    Public Structure Theta_Info                         ' θ調整情報
        Dim iFlg As Short                               ' 0:θﾏﾆｭｱﾙ調整を無効とする, 1:θﾏﾆｭｱﾙ調整を有効とする
        Dim iPP30 As Short                              ' 位置補正モード(0:自動補正モード, 1:手動補正モード, 2:自動+微調)
        Dim iPP31 As Short                              ' 位置補正方法(0:補正なし, 1:補正あり)
        Dim fPP53Min As Double                          ' θ回転最小角度
        Dim fPP53Max As Double                          ' θ回転最大角度
        Dim fTheta As Double                            ' θ軸角度
        Dim fpp32_x As Double                           ' パターン1座標x
        Dim fpp32_y As Double                           ' パターン1座標y
        Dim fpp33_x As Double                           ' パターン2座標x
        Dim fpp33_y As Double                           ' パターン2座標y
        Dim iPP37_1 As Short                            ' パターン1 テンプレート番号
        Dim iPP37_2 As Short                            ' パターン2 テンプレート番号
        Dim iPP38 As Short                              ' パターングループ番号
        Dim fpp34_x As Double                           ' 補正ポジションオフセットx
        Dim fpp34_y As Double                           ' 補正ポジションオフセットy
    End Structure
    Public stThta As Theta_Info                         ' θ調整情報

    Public stResult As Theta_Cor_Info          ' θ補正結果 ※｢Theta_Cor_Info｣はVideo.OCXで定義
    'Public Type Theta_Cor_Info
    '    fTheta          As Double                      ' θ角度(°)
    '    fPosx           As Double                      ' トリム位置X(mm)
    '    fPosy           As Double                      ' トリム位置Y(mm)
    '    fPos1x          As Double                      ' 補正位置1X(mm)
    '    fPos1y          As Double                      ' 補正位置1Y(mm)
    '    fPos2x          As Double                      ' 補正位置2X(mm)
    '    fPos2y          As Double                      ' 補正位置2Y(mm)
    '    fCorx           As Double                      ' トリム位置Xのずれ量(mm)
    '    fCory           As Double                      ' トリム位置Yのずれ量(mm)
    '    fCor1x          As Double                      ' 補正位置1Xのずれ量(mm)
    '    fCor1y          As Double                      ' 補正位置1Yのずれ量(mm)
    '    fCor2x          As Double                      ' 補正位置2Xのずれ量(mm)
    '    fCor2y          As Double                      ' 補正位置2Yのずれ量(mm)
    '    fCorV1          As Double                      ' 補正位置1の一致度(閾値)
    '    fCorV2          As Double                      ' 補正位置2の一致度(閾値)
    'End Type
#End Region

#Region "トリミング結果表示用データ他"

    '-------------------------------------------------------------------------------
    '   トリミング結果表示用データ (frmInfo画面)
    '-------------------------------------------------------------------------------
    '----- ロット情報 -----
    Public Structure RESULT_PARAM                   ' ロット情報形式定義
        ' 基板単位
        Dim StartTime As DateTime                   ' 基板スタート時間
        Dim EndTime As DateTime                     ' 基板エンド時間
        Dim BlockCounter As Integer                 ' ブロックカウンター
        Dim BlockCntX As Integer                    ' ブロックカウンターＸ
        Dim BlockCntY As Integer                    ' ブロックカウンターＹ
        Dim PlateCntX As Integer                    ' プレートカウンターＸ
        Dim PlateCntY As Integer                    ' プレートカウンターＹ
        Dim TrimCounter As Integer                  ' ﾄﾘﾐﾝｸﾞ数(ﾜｰｸ投入数)
        Dim OK_Counter As Integer                   ' OK数
        Dim NG_Counter As Integer                   ' NG数
        Dim ITHigh As Integer                       ' 初期測定上限値異常
        Dim ITLow As Integer                        ' 初期測定下限値異常
        Dim ITOpen As Integer                       ' 測定値異常
        Dim FTHigh As Integer                       ' 最終測定上限値異常
        Dim FTLow As Integer                        ' 最終測定下限値異常
        Dim FTOpen As Integer                       ' 測定値異常
        Dim Pattern As Integer                      ' カット位置補正の判定 'V1.2.0.0③
        Dim VaNG As Integer                         ' "VA-NG" ' 再測定変化量エラー'V2.0.0.0②
        Dim StdNg As Integer                        ' スタンダード抵抗測定エラー 'V2.0.0.0⑮
        Dim ValLow As Integer                       ' カット後上昇率：Low異常     'V2.2.0.029 
        Dim ValHigh As Integer                      ' カット後上昇率：High異常    'V2.2.0.029 

        ' ロット通算
        Dim LotStart As DateTime                    ' ロット開始時間
        Dim LotEnd As DateTime                      ' ロット終了時間
        Dim LotCounter As Integer                   ' 処理ロット数
        Dim PlateCounter As Integer                 ' 基板カウンター
        Dim Total_TrimCounter As Integer            ' 抵抗トータル処理数
        Dim Total_OK_Counter As Integer             ' OK数
        Dim Total_NG_Counter As Integer             ' NG数
        Dim Total_ITHigh As Integer                 ' 初期測定上限値異常
        Dim Total_ITLow As Integer                  ' 初期測定下限値異常
        Dim Total_ITOpen As Integer                 ' 測定値異常
        Dim Total_FTHigh As Integer                 ' 最終測定上限値異常
        Dim Total_FTLow As Integer                  ' 最終測定下限値異常
        Dim Total_FTOpen As Integer                 ' 測定値異常
        Dim Total_Pattern As Integer                ' カット位置補正の判定 'V1.2.0.0③
        Dim Total_VaNG As Integer                   ' "VA-NG" ' 再測定変化量エラー'V2.0.0.0②
        Dim Total_StdNg As Integer                  ' スタンダード抵抗測定エラー 'V2.0.0.0⑮
        Dim LotPrint As Boolean                     ' ロット印刷の有無
        ' その他
        Dim Probe_Counter As Long                   ' プローブON回数

        Dim Total_ValLow As Integer                 ' カット後上昇率：Low異常     'V2.2.0.029 
        Dim Total_ValHigh As Integer                ' カット後上昇率：High異常    'V2.2.0.029 

    End Structure
    Public stCounter As RESULT_PARAM                        ' 表示用データ定義

    Public Enum COUNTER
        NONE
        INITIALIZE
        PRODUCT_INIT
        PROBE_INIT
        ALLDATA_DISP
        LOT_UP
        PROBE_UP
        COUNTUP
        OKNG_UP
        INITIAL_DISP
        SKIP
    End Enum

    '-------------------------------------------------------------------------------
    '  トリミング関連データ他
    '-------------------------------------------------------------------------------
    '----- Digital SW -----
    Public DGH As Short                                 ' Digital SW(Hight)
    Public DGL As Short                                 ' Digital SW(Low)
    Public DGSW As Short                                ' Digital SW

    '----- 測定値表示用 -----
    Public dblNM(2) As Double                           ' 目標値(1:IT値/2:FT値)
    Public dblVX(2) As Double                           ' 測定値(1:IT値/2:FT値)
    Public strJUG(MAXRNO) As String                     ' IT/FT判定("OK   ","FT-HI","FT-LO"他)
    '                                                   ' ※strJUG(0)は1ﾌﾞﾛｯｸの判定結果

    '----- トリミング関連 -----
    Public dblLN(MAX_LCUT, MAXCTN) As Double            ' ｶｯﾄ量mm(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後) * ｶｯﾄ数分 V1.0.4.3⑦２をMAX_LCUTへ変更
    Public dblML(MAX_LCUT) As Double                    ' ﾘﾐｯﾄｶｯﾄ量mm(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)　　　 V1.0.4.3⑦２をMAX_LCUTへ変更
    Public LTFlg As Short                               ' Lﾀｰﾝﾌﾗｸﾞ(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
    Public LTAng(MAX_LCUT) As Double                    ' ｶｯﾄ方向(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)　　　　　 V1.0.4.3⑦２をMAX_LCUTへ変更、ShortからDoubleに変更
    Public LTP As Double                                ' Lﾀｰﾝﾎﾟｲﾝﾄ(%)

    '----- トリミングデータ -----
    Public gsDataFileName As String                     ' データファイル名

    '----- メッセージデータ -----
    Public TTL_Msg(10) As String                        ' タイトル メッセージ(0 ORG)
    Public prf As Short                                 ' 印刷(0:しない, 1:する) ※D_DREAD2()で設定

    '-------------------------------------------------------------------------------
    '   ログ関連データ
    '-------------------------------------------------------------------------------
    Public gsLogFileName As String                      ' ログファイル名を保持
    Public gsLogBlock As String                         ' ﾌﾞﾛｯｸ番号(X,Y)



    '-------------------------------------------------------------------------------
    '   LEDデータ
    '-------------------------------------------------------------------------------
    Const LED_ONOFF As Short = &H8                      ' ■■　LED ON/OFF Bit

    '-------------------------------------------------------------------------------
    '   そのほか
    '-------------------------------------------------------------------------------
    Public Const FUNC_OK As Short = 1                          ' ■■　OK
    Public Const FUNC_NG As Short = 0                          ' ■■　NG

    'V2.2.0.0⑯↓
    Public Const MAXBlock = 5

    ' 基板１枚から複数抵抗値を取得するための構造体 
    Public Structure BLOCK_DATA

        Dim DataNo As Integer                               ' No：指定番号＝オープンカット時のカット本数
        Dim gBlockCnt As Integer                            ' ブロック数 
        Dim gProcCnt As Integer                             ' 処理数カウンター 

        <VBFixedArray(MAXBlock)> Dim dblNominal() As Double ' 1ブロック内抵抗数 
        <VBFixedArray(MAXBlock)> Dim iUnit() As Integer   ' 単位 
        <VBFixedArray(MAXBlock)> Dim dblCorr() As Double    ' 補正値 

        Public Sub Initialize()
            ReDim dblNominal(MAXBlock)
            ReDim iUnit(MAXBlock)
            ReDim dblCorr(MAXBlock)
        End Sub

    End Structure

    Public Structure MULTI_BLOCK_DATA

        Dim gStepRpt As Integer                                     ' 並び方向 : 0:Y方向＝列、1:X方向＝行
        Dim gMultiBlock As Integer                                  ' 複数抵抗値指定：0:通常1種類、1:マルチ

        <VBFixedArray(MAXBlock)> Dim BLOCK_DATA() As BLOCK_DATA ' 1ブロック用データ 

        Public Sub Initialize()
            ReDim BLOCK_DATA(MAXBlock)
        End Sub

    End Structure

    Public Structure SAVE_BLOCK_DATA

        Dim dblNominal As Double                            ' No：指定番号＝オープンカット時のカット本数
        Dim dblCorr As Double                               ' ブロック数 

    End Structure

    ' 最大５種類の抵抗値を作成
    Public stMultiBlock As MULTI_BLOCK_DATA                 ' 基板１枚から複数抵抗値を作成するために使用する構造体 
    Public gMultiBlockNo As Integer                         ' マルチブロックの番号 
    Public stDefaultBlock(MAXRNO) As SAVE_BLOCK_DATA        ' 指定ブロックが存在しなかった場合のデフォルト値 
    Public stExecBlkData As BLOCK_DATA                      ' 実行中の複数抵抗値処理の内容保存用
    Public beforeExecBlkDataNo As Integer                   ' 前回実行したマルチブロックNo       'V2.2.0.033

    'V2.2.0.0⑯ 
    ' 集計データ保存用 
    Structure TOTAL_DATA_MULTI

        <VBFixedArray(MAX_RES_USER)> Dim gITNx_cnt() As Integer     ' IT 算出用ﾜｰｸ数
        <VBFixedArray(MAX_RES_USER)> Dim gITNg_cnt() As Integer     ' IT NG数記録
        <VBFixedArray(MAX_RES_USER)> Dim gFTNx_cnt() As Integer     ' FT 算出用ﾜｰｸ数
        <VBFixedArray(MAX_RES_USER)> Dim gFTNg_cnt() As Integer     ' FT NG数記録
        <VBFixedArray(MAX_RES_USER)> Dim dblAverage() As Double     ' 平均値
        <VBFixedArray(MAX_RES_USER)> Dim dblDeviationIT() As Double ' 標準偏差(IT)
        <VBFixedArray(MAX_RES_USER)> Dim dblDeviationFT() As Double ' 標準偏差(FT)
        <VBFixedArray(MAX_RES_USER)> Dim dblAverageIT() As Double   ' IT平均値
        <VBFixedArray(MAX_RES_USER)> Dim dblAverageFT() As Double   ' FT平均値
        <VBFixedArray(MAX_RES_USER)> Dim TotalIT() As Double        ' IT 合計
        <VBFixedArray(MAX_RES_USER)> Dim TotalFT() As Double        ' FT 合計
        <VBFixedArray(MAX_RES_USER)> Dim TotalSum2IT() As Double    ' IT２乗和
        <VBFixedArray(MAX_RES_USER)> Dim TotalSum2FT() As Double    ' FT２乗和
        <VBFixedArray(MAX_RES_USER)> Dim dblMinIT() As Double       ' IT最小値ﾌｧｲﾅﾙ
        <VBFixedArray(MAX_RES_USER)> Dim dblMaxIT() As Double       ' IT最大値ﾌｧｲﾅﾙ
        <VBFixedArray(MAX_RES_USER)> Dim dblMinFT() As Double       ' FT最小値ﾌｧｲﾅﾙ
        <VBFixedArray(MAX_RES_USER)> Dim dblMaxFT() As Double       ' FT最大値ﾌｧｲﾅﾙ
        <VBFixedArray(MAX_RES_USER)> Dim TrimCounter() As Double    ' トリミング数カウンター
        <VBFixedArray(MAX_RES_USER)> Dim Total_TrimCounter() As Double ' トリミング数カウンター


        Public stCounter1 As RESULT_PARAM                        ' 表示用データ定義

        'この構造体のインスタンスを初期化するには、"Initialize" を呼び出さなければなりません。
        Public Sub Initialize()
            ReDim gITNx_cnt(MAX_RES_USER)                       ' IT 算出用ﾜｰｸ数
            ReDim gITNg_cnt(MAX_RES_USER)                       ' IT NG数記録
            ReDim gFTNx_cnt(MAX_RES_USER)                       ' FT 算出用ﾜｰｸ数
            ReDim gFTNg_cnt(MAX_RES_USER)                       ' FT NG数記録
            ReDim dblAverage(MAX_RES_USER)                      ' 平均値
            ReDim dblDeviationIT(MAX_RES_USER)                  ' 標準偏差(IT)
            ReDim dblDeviationFT(MAX_RES_USER)                  ' 標準偏差(FT)
            ReDim dblAverageIT(MAX_RES_USER)                    ' IT平均値
            ReDim dblAverageFT(MAX_RES_USER)                    ' FT平均値

            ReDim TotalIT(MAX_RES_USER)                         ' IT 合計
            ReDim TotalFT(MAX_RES_USER)                         ' FT 合計
            ReDim TotalSum2IT(MAX_RES_USER)                     ' IT２乗和 
            ReDim TotalSum2FT(MAX_RES_USER)                     ' FT２乗和 

            ReDim dblMinIT(MAX_RES_USER)                        ' IT最小値ﾌｧｲﾅﾙ
            ReDim dblMaxIT(MAX_RES_USER)                        ' IT最大値ﾌｧｲﾅﾙ
            ReDim dblMinFT(MAX_RES_USER)                        ' FT最小値ﾌｧｲﾅﾙ
            ReDim dblMaxFT(MAX_RES_USER)                        ' FT最大値ﾌｧｲﾅﾙ
            ReDim TrimCounter(MAX_RES_USER)                     ' トリミング数カウンター
            ReDim Total_TrimCounter(MAX_RES_USER)               ' トリミング数カウンタートータル
        End Sub

    End Structure
    ' 複数抵抗値取得用の集計データ保存用 
    Public stToTalDataMulti(MAX_RES_USER) As TOTAL_DATA_MULTI       ' stToTalDataMultiは配列番号１から使用


    ''' <summary>
    ''' Multi動作する、しない
    ''' </summary>
    Public Enum MULTI_MODE
        NONE = 0
        EXEC_MULTI
    End Enum

    'V2.2.0.0⑯↑

    Public sStrTrig As String = ""                              ' GPIB送信コマンド：HIOKI用で使用   'V2.2.1.4① 
    Public gLastsetNomx As Double = 0.0                         ' 最後に設定した時の目標値  'V2.2.1.4① 

    'V2.2.1.7③↓
    Structure MARKING_ALARMLIST
        Dim AlarmTrimData As String                             ' アラームになった基板処理時のロット＝トリミングデータ名
        Dim LotCount As Integer                                 ' アラームになった基板のロット内枚数
    End Structure
    Public LotMarkingAlarmCnt As Integer                        ' 現在の自動運転中にマーキングエラーとなった基板数

    Public gMarkAlarmList(10) As MARKING_ALARMLIST

    ''' <summary>
    ''' マーク印字ログ出力用
    ''' </summary>
    Structure LogMarkPrint
        Dim sOperator As String                                 ' 作業者名
        Dim sDate As String                                     ' 日付
        Dim sLaserLot As String                                 ' レーザーロット番号
        Dim sInitialNo As String                                ' 開始番号 
        Dim sEndNo As String                                    ' 終了番号
        Dim sAutoOpeStartTime As String                         ' 自動運転開始時間
        Dim sAutoOpeEndTime As String                           ' 自動運転終了時間 
    End Structure

    Public gLogMarkPrint As LogMarkPrint                    ' マーク印字ログ保存用

    'V2.2.1.7③↑

#End Region

#Region "ユーザプログラム向けカット共通パラメータ"
    '-----------------------------------------------------------------------
    '   ユーザプログラム向けカットパラメータの構造体
    '   各カットのパラメータ(長さ、スピード、Qレート)保存領域
    '-----------------------------------------------------------------------
    Dim cutCmnPrm As CUT_COMMON_PRM

#End Region

    '=========================================================================
    '   メソッド定義
    '=========================================================================
    '=========================================================================
    '   初期設定処理
    '=========================================================================
#Region "カットパラメータを初期化する"
    '''=========================================================================
    ''' <summary>カットパラメータを初期化する</summary>
    ''' <param name="pstCutCmnPrm">(I/O)カットパラメータ</param> 
    '''=========================================================================
    Public Sub InitCutParam(ByRef pstCutCmnPrm As CUT_COMMON_PRM)

        Dim strMSG As String

        Try
            ' カットパラメータを初期化する(カット情報構造体)
            pstCutCmnPrm.CutInfo.srtMoveMode = 1                    ' 動作モード（0:トリミング、1:ティーチング、2:強制カット）
            pstCutCmnPrm.CutInfo.srtCutMode = 0                     ' カットモード(0:ノーマル、1:リターン、2:リトレース、3:斜め）
            pstCutCmnPrm.CutInfo.dblTarget = 0.0#                   ' 目標値
            pstCutCmnPrm.CutInfo.srtSlope = 0                       ' 4:抵抗測定＋スロープ(4->0)
            pstCutCmnPrm.CutInfo.srtMeasType = 0                    ' 測定タイプ(0:高速(3回)、1:高精度(2000回)
            pstCutCmnPrm.CutInfo.dblAngle = 0.0#                    ' カット角度
            pstCutCmnPrm.CutInfo.dblLTP = 0.0#                      ' Lターンポイント
            pstCutCmnPrm.CutInfo.srtLTDIR = 0                       ' Lターン後の方向
            pstCutCmnPrm.CutInfo.dblRADI = 0.0#                     ' R部回転半径（Uカットで使用）
            ' For Hook Or UCut
            pstCutCmnPrm.CutInfo.dblRADI2 = 0.0#                    ' R2部回転半径（Uカットで使用）
            pstCutCmnPrm.CutInfo.srtHkOrUType = 0                   ' HookCut(3)かUカット（3以外）の指定。
            ' For Index
            pstCutCmnPrm.CutInfo.srtIdxScnCnt = 0                   ' インデックス/スキャンカット数(1～32767)
            pstCutCmnPrm.CutInfo.srtIdxMeasMode = 0                 ' インデックス測定モード（0:抵抗、1:電圧、2:外部）
            ' For EdgeSense
            pstCutCmnPrm.CutInfo.dblEsPoint = 0.0#                  ' エッジセンスポイント
            pstCutCmnPrm.CutInfo.dblRdrJdgVal = 0.0#                ' ラダー内部判定変化量
            pstCutCmnPrm.CutInfo.dblMinJdgVal = 0.0#                ' ラダーカット後最低許容変化量
            pstCutCmnPrm.CutInfo.srtEsAftCutCnt = 0                 ' ラダー切抜け後のカット回数（測定回数）
            pstCutCmnPrm.CutInfo.srtMinOvrNgCnt = 0                 ' ラダー抜出し後、最低変化量の連続Over許容数
            pstCutCmnPrm.CutInfo.srtMinOvrNgMode = 0                ' 連続Over時のNG処理（0:NG判定未実施, 1:NG判定実施。ラダー中切り, 2:NG判定未実施。ラダー切上げ）
            ' For Scan
            pstCutCmnPrm.CutInfo.dblStepPitch = 0.0#                ' ステップ移動ピッチ
            pstCutCmnPrm.CutInfo.srtStepDir = 0                     ' ステップ方向

            ' カットパラメータを初期化する(加工設定構造体)
            pstCutCmnPrm.CutCond.CutLen.dblL1 = 0.0#                  ' カット長(Line1用)
            pstCutCmnPrm.CutCond.CutLen.dblL2 = 0.0#                  ' カット長(Line2用)
            pstCutCmnPrm.CutCond.CutLen.dblL3 = 0.0#                  ' カット長(Line3用)
            pstCutCmnPrm.CutCond.CutLen.dblL4 = 0.0#                  ' カット長(Line4用)

            pstCutCmnPrm.CutCond.SpdOwd.dblL1 = 0.0#                  ' カットスピード（往路）(Line1用)
            pstCutCmnPrm.CutCond.SpdOwd.dblL2 = 0.0#                  ' カットスピード（往路）(Line2用)
            pstCutCmnPrm.CutCond.SpdOwd.dblL3 = 0.0#                  ' カットスピード（往路）(Line3用)
            pstCutCmnPrm.CutCond.SpdOwd.dblL4 = 0.0#                  ' カットスピード（往路）(Line4用)

            pstCutCmnPrm.CutCond.SpdRet.dblL1 = 0.0#                  ' カットスピード（復路）(Line1用)
            pstCutCmnPrm.CutCond.SpdRet.dblL2 = 0.0#                  ' カットスピード（復路）(Line2用)
            pstCutCmnPrm.CutCond.SpdRet.dblL3 = 0.0#                  ' カットスピード（復路）(Line3用)
            pstCutCmnPrm.CutCond.SpdRet.dblL4 = 0.0#                  ' カットスピード（復路）(Line4用)

            pstCutCmnPrm.CutCond.QRateOwd.dblL1 = 0.0#                ' カットQレート（往路）(Line1用)
            pstCutCmnPrm.CutCond.QRateOwd.dblL2 = 0.0#                ' カットQレート（往路）(Line2用)
            pstCutCmnPrm.CutCond.QRateOwd.dblL3 = 0.0#                ' カットQレート（往路）(Line3用)
            pstCutCmnPrm.CutCond.QRateOwd.dblL4 = 0.0#                ' カットQレート（往路）(Line4用)

            pstCutCmnPrm.CutCond.QRateRet.dblL1 = 0.0#                ' カットQレート（復路）(Line1用)
            pstCutCmnPrm.CutCond.QRateRet.dblL2 = 0.0#                ' カットQレート（復路）(Line2用)
            pstCutCmnPrm.CutCond.QRateRet.dblL3 = 0.0#                ' カットQレート（復路）(Line3用)
            pstCutCmnPrm.CutCond.QRateRet.dblL4 = 0.0#                ' カットQレート（復路）(Line4用)

            pstCutCmnPrm.CutCond.CondOwd.srtL1 = 0                  ' カット条件番号（往路）(Line1用)
            pstCutCmnPrm.CutCond.CondOwd.srtL2 = 0                  ' カット条件番号（往路）(Line2用)
            pstCutCmnPrm.CutCond.CondOwd.srtL3 = 0                  ' カット条件番号（往路）(Line3用)
            pstCutCmnPrm.CutCond.CondOwd.srtL4 = 0                  ' カット条件番号（往路）(Line4用)

            pstCutCmnPrm.CutCond.CondRet.srtL1 = 0                  ' カット条件番号（復路）(Line1用)
            pstCutCmnPrm.CutCond.CondRet.srtL2 = 0                  ' カット条件番号（復路）(Line2用)
            pstCutCmnPrm.CutCond.CondRet.srtL3 = 0                  ' カット条件番号（復路）(Line3用)
            pstCutCmnPrm.CutCond.CondRet.srtL4 = 0                  ' カット条件番号（復路）(Line4用)

            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.InitCutParam() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "構造体の初期化"
    '''=========================================================================
    ''' <summary>構造体の初期化</summary>
    ''' <remarks>配列を使用している構造体のインスタンスを初期化するには"Initialize"を呼び出さなければならない</remarks>
    '''=========================================================================
    Public Sub Init_Struct()

        Dim i As Integer
        Dim j As Integer
        Dim strMSG As String

        Try
            ' ユーザ情報の初期化
            stUserData.Initialize()

            ' 抵抗データ/カットデータ構造体の初期化
            For i = 0 To MAXRNO
                stREG(i).Initialize()
                For j = 0 To MAXCTN
                    stREG(i).STCUT(j).Initialize()
                Next j
            Next i

            ' トリミング要求/応答データ構造体の初期化
            stSCMD.Initialize()                         ' 要求データ(コマンド)
            stSRES.Initialize()                         ' 応答データ(コマンド)
            stTGPI.prmGPIB.Initialize()                 ' GPIBデータ(トリミング要求データ)
            stTCUT.prmCut.Initialize()                  ' カットデータ(トリミング要求データ)
            stCutL.c.Initialize()                       ' L cutパラメータ(トリミング要求データ)
            stCutHK.c.Initialize()                      ' HOOK cutパラメータ(トリミング要求データ)
            stCutMK.c.Initialize()                      ' Letter Markingパラメータ(トリミング要求データ)

            ' トリミング結果データ構造体の初期化
            stResultWd.Initialize()                     ' WORD型データ用
            stResultDd.Initialize()                     ' Double型データ用)

#If cOSCILLATORcFLcUSE Then
            ' トリマー加工条件構造体(FL用)初期化
            stCND.Initialize()
#End If

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Init_Struct() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "初期設定処理"
    '''=========================================================================
    ''' <summary>初期設定処理</summary>
    ''' <returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Public Function UserVal() As Short

        Dim r As Short                                                  ' 関数Return値
        Dim strMSG As String

        Try
            ' 変数初期値設定
            UserVal = 0                                                 ' Retern値 = Normal
            stPLT.z_xoff = 0 : stPLT.z_yoff = 0                         ' XY Table Offset(mm)
            stPLT.Z_ZOFF = 0.0# : stPLT.Z_ZON = 0.0#                    ' Z PROBE OFF/ON OFFSET(mm)
            'stPLT.Z2_ZOFF = 0#: stPLT.Z2_ZON = 0#                      ' Z2 PROBE OFF/ON OFFSET(mm)
            stPLT.Pnx = 1 : stPLT.Pny = 1                               ' プレート数x ,y
            stPLT.Pivx = 0.0# : stPLT.Pivy = 0.0#                       ' プレートインターバルx(mm) = 0
            stPLT.ADJX = 0 : stPLT.ADJY = 0                             ' Adjust Point(mm)
            giTemplateGroup = -1                                        ' パターンテンプレートグループナンバー
            FlgUpd = TriState.False                                     ' データ更新 Flag OFF

            ' トリミングデータ設定
            r = rData_load()                                            ' データファイルリード
            If (r <> 0) Then                                            ' データファイル　ロードエラー ?
                UserVal = 1                                             ' Retern値 = エラー
                Exit Function
            End If

#If cOFFLINEcDEBUG = 0 Then
            ' システム変数設定(プローブON/OFF位置他)
            r = ObjSys.EX_PROP_SET(gSysPrm, stPLT.Z_ZON, stPLT.Z_ZOFF, gSysPrm.stDEV.gfTrimX, gSysPrm.stDEV.gfTrimY, gSysPrm.stDEV.gfSmaxX, gSysPrm.stDEV.gfSmaxY)
            If (r <> 0) Then                                            ' システム変数設定エラー ?
                UserVal = 2                                             ' Retern値 = エラー
                Exit Function
            End If
#End If

            ' ログファイル名を設定する ("C:\TRIMDATA\LOG\""LOG_yyyymmdd" + ".LOG")
            Call SetLogFileName(gsLogFileName)

            'V2.2.0.0②↓
            If LaserFront.Trimmer.DllVideo.VideoLibrary.IsDigitalCamera Then
                ObjVdo.StdMagnification = CDec(stPLT.dblStdMagnification)         ' 内部カメラ表示倍率を設定 
            End If
            'V2.2.0.0②↑

            ' GPIB初期化
            r = GPIB_Init()                                            ' GPIB初期化
            If (r <> 0) Then                                           ' GPIB初期化エラー ?
                'UserVal = 3                                            ' Retern値 = エラー
            End If

            ' 生産数,良品数,表示(frmInfo画面)
            Call Disp_frmInfo(COUNTER.ALLDATA_DISP, COUNTER.NONE)

            ' ｼｸﾞﾅﾙﾀﾜｰ制御(On=なし,Off=全ビット)
            Call ObjSys.SetSignalTower(0, &HFFFFS)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.UserVal() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Function
#End Region

#Region "偏差を求める"
    '''=========================================================================
    ''' <summary>偏差を求める</summary>
    ''' <param name="a"></param>
    ''' <param name="b"></param>
    ''' <returns>FNDEV(a, b) = a - b</returns>
    '''=========================================================================
    Public Function FNDEV(ByRef a As Double, ByRef b As Double) As Double

        Dim strMSG As String

        Try
            FNDEV = a - b

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.FNDEV() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Function
#End Region

#Region "偏差(%)を求める"
    '''=========================================================================
    ''' <summary>偏差(%からppmへ変更)を求める</summary>
    ''' <param name="a"></param>
    ''' <param name="b"></param>
    ''' <returns>FNDEVP(a:実測値, b:目標値) = (a:実測値 / b:目標値 - 1) * 100</returns>
    '''=========================================================================
    Public Function FNDEVP(ByRef a As Double, ByRef b As Double) As Double

        Dim strMSG As String

        Try
            ' PPMへ変更            FNDEVP = (a / b - 1.0#) * 100.0#
            FNDEVP = (a / b - 1.0#) * 10.0 ^ 6
            Exit Function

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.FNDEVP() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            FNDEVP = 0
        End Try
    End Function
#End Region

#Region "機能選択定義テーブル設定"
    '''=========================================================================
    '''<summary>機能選択定義テーブル設定</summary>
    '''<returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Public Function GetFncDefParameter() As Short

        Dim i As Short                                                  ' Counter
        Dim sPath As String                                             ' ﾌｧｲﾙ名
        Dim sSect As String                                             ' ｾｸｼｮﾝ名
        Dim strMSG As String

        Try
            ' 初期処理
            GetFncDefParameter = 0                                      ' Return値 = 正常
            sPath = cDEF_FNAME                                          ' ユーザ定義ファイル名

            ' ｷｰ名を設定する
            stFNC(F_LOAD).sCMD = "LOAD"                                 ' LOADボタン
            stFNC(F_SAVE).sCMD = "SAVE"                                 ' SAVEボタン
            stFNC(F_EDIT).sCMD = "EDIT"                                 ' EDITボタン
            stFNC(F_LASER).sCMD = "LASER"                               ' LASERボタン
            stFNC(F_LOTCHG).sCMD = "LOTCHG"                             ' ﾛｯﾄ切替ボタン
            stFNC(F_PROBE).sCMD = "PROBE"                               ' PROBE(ﾌﾟﾛｰﾌﾞ)ボタン
            stFNC(F_TEACH).sCMD = "TEACH"                               ' TEACH(ﾃｨｰﾁﾝｸﾞ)ボタン
            stFNC(F_CUTPOS).sCMD = "CUTPOS"                             ' CUTPOS(ｶｯﾄ位置補正)ボタン
            stFNC(F_RECOG).sCMD = "RECOG"                               ' RECOG(画像登録)ボタン
            stFNC(F_TX).sCMD = "TX"                                     ' TXボタン
            stFNC(F_TY).sCMD = "TY"                                     ' TYボタン
            'stFNC(F_MSTCHK).sCMD = "MSTCHK"                             ' ﾏｽﾀﾁｪｯｸ(F4)ボタン

            ' ｾｸｼｮﾝ名を設定する
            '    #If (cFncMode = 0) Then                                ' エンジニアモード ?
            sSect = "FUNCDEF" ' ｾｸｼｮﾝ名 = "FUNCDEF"
            '    #Else
            '        sSect = "FUNCDEF_OPERATOR"                         ' ｾｸｼｮﾝ名 = "FUNCDEF_OPERATOR"
            '    #End If

            ' 機能選択定義テーブルを設定する(0:(選択不可), 1:(選択可))
            For i = 0 To (MAX_FNCNO - 1)                                ' 定義数分繰返す
                stFNC(i).iDEF = GetPrivateProfileInt(sSect, stFNC(i).sCMD, 0, sPath)
            Next i

            Exit Function

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.GetFncDefParameter() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Return値 = ﾄﾗｯﾌﾟｴﾗｰ発生
        End Try
    End Function
#End Region

#Region "パスワード機能選択定義テーブル設定"
    '''=========================================================================
    '''<summary>パスワード機能選択定義テーブル設定</summary>
    '''<returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Public Function GetPasFuncDefParameter() As Short

        Dim i As Short                                                  ' Counter
        Dim sPath As String                                             ' ﾌｧｲﾙ名
        Dim sSect As String                                             ' ｾｸｼｮﾝ名
        Dim strMSG As String

        Try
            ' 初期処理
            GetPasFuncDefParameter = 0                                  ' Return値 = 正常
            sPath = cDEF_FNAME                                          ' ユーザ定義ファイル名

            ' ｷｰ名を設定する
            stFNC(F_LOAD).sCMD = "LOAD"                                 ' LOADボタン
            stFNC(F_SAVE).sCMD = "SAVE"                                 ' SAVEボタン
            stFNC(F_EDIT).sCMD = "EDIT"                                 ' EDITボタン
            stFNC(F_LASER).sCMD = "LASER"                               ' LASERボタン
            stFNC(F_LOTCHG).sCMD = "LOTCHG"                             ' ﾛｯﾄ切替ボタン
            stFNC(F_PROBE).sCMD = "PROBE"                               ' PROBE(ﾌﾟﾛｰﾌﾞ)ボタン
            stFNC(F_TEACH).sCMD = "TEACH"                               ' TEACH(ﾃｨｰﾁﾝｸﾞ)ボタン
            stFNC(F_CUTPOS).sCMD = "CUTPOS"                             ' CUTPOS(ｶｯﾄ位置補正)ボタン
            stFNC(F_RECOG).sCMD = "RECOG"                               ' RECOG(画像登録)ボタン
            stFNC(F_TX).sCMD = "TX"                                     ' TXボタン
            stFNC(F_TY).sCMD = "TY"                                     ' TYボタン
            'stFNC(F_MSTCHK).sCMD = "MSTCHK"                             ' ﾏｽﾀﾁｪｯｸ(F4)ボタン

            ' ｾｸｼｮﾝ名を設定する
            sSect = "PASSWORD"                                          ' ｾｸｼｮﾝ名 = "PASSWORD"

            ' 機能選択定義テーブルを設定する(0:(選択不可), 1:(選択可))
            For i = 0 To (MAX_FNCNO - 1)                                ' 定義数分繰返す
                stFNC(i).iPAS = GetPrivateProfileInt(sSect, stFNC(i).sCMD, 0, sPath)
            Next i
            Exit Function

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.GetPasFuncDefParameter() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Return値 = ﾄﾗｯﾌﾟｴﾗｰ発生
        End Try
    End Function
#End Region

#Region "パスワードチェック"
    '''=========================================================================
    '''<summary>パスワードチェック</summary>
    '''<returns>True=正常, False=エラー</returns>
    '''=========================================================================
    Public Function Func_Password(ByRef IntIndexNo As Short) As Boolean

        Dim r As Short
        Dim strMSG As String

        Try
            Func_Password = True
            If (stFNC(IntIndexNo).iPAS = 1) Then
                r = ObjPas.ShowDialog((gSysPrm.stTMN.giMsgTyp), (gSysPrm.stSYP.gstrPassword))
                If (r <> 1) Then
                    Func_Password = False                               ' ﾊﾟｽﾜｰﾄﾞ入力ｴﾗｰならEXIT
                End If
            End If

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.Func_Password() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (False)                                          ' Return値 = ﾄﾗｯﾌﾟｴﾗｰ発生
        End Try
    End Function
#End Region

#Region "トリミングデータファイル名"
    Public Function GetTrimmingDataFileName() As String
        Try
            Return (gsDataFileName)
        Catch ex As Exception
            MsgBox("GetTrimmingDataFileName() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (gsDataFileName)
        End Try
    End Function
    Public Sub SetTrimmingDataFileName(ByVal TrimmingFileName As String)
        Try
            gsDataFileName = TrimmingFileName
        Catch ex As Exception
            MsgBox("SetTrimmingDataFileName() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Sub
#End Region
    '=========================================================================
    '   データファイルロード/セーブ処理
    '=========================================================================
#Region "データファイルロード"
    '''=========================================================================
    '''<summary>データファイルロード</summary>
    '''<returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Public Function rData_load() As Short

        Dim a As String                                 ' データ入力域
        Dim b(20) As Short                              ' データ入力域
        Dim c(5) As String                              ' データ入力域
        Dim X As Double                                 ' データ入力域
        Dim y As Double                                 ' データ入力域
        Dim z(2) As Short                               ' データ入力域
        Dim i As Short                                  ' Counter
        Dim j As Short                                  ' Counter
        Dim k As Short                                  ' Counter
        Dim strMSG As String                            ' データ入力域
        Dim VersionNumber As Integer = 0                ' 旧バージョンファイルの判定
        ' 'V2.2.0.038       Dim stPROBEDATA(11) As stPROBEDATA_TABLE        ' プローブデータ定義 V2.2.0.0⑮ 
        Dim stPROBEDATA(PROBE_DATA_MAX) As stPROBEDATA_TABLE        ' プローブデータ定義  'V2.2.0.038　'V2.2.1.0①
        Dim MaxNo As Integer                            ' プローブデータのテーブルデータ数 V2.2.0.0⑮ 

        '--------------------------------------------------------------------------
        '   データファイルよりシステムデータを設定する
        '--------------------------------------------------------------------------
        rData_load = 1                                      ' Return値 = エラー
        a = ""
        On Error GoTo rData_load_FileNone
        fNum = FreeFile()                                   ' Data File No.
        FileOpen(fNum, gsDataFileName, OpenMode.Input)      ' データファイル オープン

        ' 各種オフセットデータ等を設定する
        On Error GoTo STP_SERR
        'V2.0.0.0        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
        'V2.0.0.0
        Input(fNum, a) : Input(fNum, a)    ' 2行空読み
        If a.IndexOf("[ DATA VERSION 3 ]") >= 0 Then        ' ●システムデータ [ DATA VERSION 3 ] の文字を挿入する。
            VersionNumber = 3
        ElseIf a.IndexOf("[ DATA VERSION 4 ]") >= 0 Then    ' ●システムデータ [ DATA VERSION 4 ] の文字を挿入する。'V2.0.0.1③　トリミングNG信号を出力するＮＧの比率を追加
            VersionNumber = 4
        ElseIf a.IndexOf("[ DATA VERSION 5 ]") >= 0 Then    'V2.1.0.0①カット毎の抵抗値変化量判定機能②レーザーパワーキャリブレーション機能③温度センサー情報の一元管理機能
            VersionNumber = 5                               'V2.1.0.0①～③
        ElseIf a.IndexOf("[ DATA VERSION 6 ]") >= 0 Then    'V2.1.0.0①カット毎の抵抗値変化量判定機能②レーザーパワーキャリブレーション機能③温度センサー情報の一元管理機能'V2.2.0.0②
            VersionNumber = 6                               'V2.1.0.0①～③ 'V2.2.0.0②
        ElseIf a.IndexOf("[ DATA VERSION 7 ]") >= 0 Then    'V2.2.0.034 カット毎の抵抗値変化量判定機能②レーザーパワーキャリブレーション機能③温度センサー情報の一元管理機能
            VersionNumber = 7                               'V2.2.0.034
        ElseIf a.IndexOf("[ DATA VERSION 8 ]") >= 0 Then    'V2.2.1.7① "マーク印字"機能追加
            VersionNumber = 8                               'V2.2.1.7①
        End If
        Input(fNum, a)
        'V2.0.0.0
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
        Input(fNum, stPLT.Z_ZOFF)
        Input(fNum, stPLT.Z_ZON)                            ' Z-Probe Off/On Offset (Off/On(mm))
        'V2.2.0.0⑮ ↓
        'V2.2.2.0②         '        If VersionNumber >= 6 Then      ' ﾌﾟﾛｰﾌﾞNo
        If VersionNumber >= 7 Then      ' ﾌﾟﾛｰﾌﾞNo      'V2.2.2.0② 
            Input(fNum, stPLT.ProbNo)                       ' Probe No 
        Else
            ' ファイルバージョンが6未満の場合はプローブマスター化していないためトリミングデータそのものを使用して０とする
            stPLT.ProbNo = 0

            'V2.2.2.0⑥  '#0128のみ旧バージョンであったらプローブをすこし上げた値にする
            If giLoaderType <> 0 Then
                stPLT.Z_ZON = 10.0              'V2.2.0.0⑲
                stPLT.Z_ZOFF = 8.0              'V2.2.0.0⑲
            End If
            'V2.2.2.0⑥

        End If
        'V2.2.0.0⑮ ↑
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
        Input(fNum, stPLT.z_xoff)
        Input(fNum, stPLT.z_yoff)                           ' trim position offset x,y(mm)
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
        Input(fNum, stPLT.BPOX)
        Input(fNum, stPLT.BPOY)                             ' BP offset x,y(mm)
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
        Input(fNum, stPLT.ADJX)
        Input(fNum, stPLT.ADJY)                             ' アジャスト位置x,y(mm)
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
        Input(fNum, stPLT.BNX)
        Input(fNum, stPLT.BNY)                              ' Block数X,Y
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
        Input(fNum, stPLT.zsx)
        Input(fNum, stPLT.zsy)                              ' Block SizeX,Y(mm)
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
        Input(fNum, stPLT.PrbRetry)                         ' プローブリトライ(1:有 0:無)
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
        Input(fNum, stLASER.intQR)
        Input(fNum, stLASER.dblspecPower)                   ' ﾚｰｻﾞ調整Qﾚｰﾄa(0.1KHz), 指定ﾊﾟﾜｰ(W)
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
        Input(fNum, glWTimeT(1))
        Input(fNum, glWTimeT(2))
        Input(fNum, glWTimeT(3))                            ' ﾄﾘﾐﾝｸﾞ用ポーズ時間ms1～3(1:2000ms, 2:   0ms, 3:   0ms)
        Input(fNum, glWTimeM(1))
        Input(fNum, glWTimeM(2))
        Input(fNum, glWTimeM(3))                            ' 測定用ポーズ時間ms1～3  (1:5000ms, 2:4000ms, 3: 300ms)
        Input(fNum, a)                                      ' 1行空読み
        If (a = "/") Then                                   ' ###1040旧データの場合は、ここで終了する。
            stLASER.iTrimAtt = 0                            ' ###1040③保存無し
            stLASER.dblRotPar = gSysPrm.stRAT.gfAttRate     ' ###1040③ 減衰率(%)
            stLASER.dblRotAtt = gSysPrm.stRAT.giAttRot      ' ###1040③ ロータリーアッテネータの回転量(0-FFF)
            stLASER.iFixAtt = gSysPrm.stRAT.giAttFix        ' ###1040③ 固定アッテネータのON/OFF(0:OFF,1:ON)
            stPLT.TeachBlockX = 1                           ' ###1040① ティーチング・ブロックX
            stPLT.TeachBlockY = 1                           ' ###1040① ティーチング・ブロックY
            'V2.0.0.1②            stPLT.StageSpeedY = 25000                       ' ###1040④ Ｙ軸スピード
            stPLT.StageSpeedY = SETAXISSPDY_DEFALT          ' V2.0.0.1② Ｙ軸スピード
            stPLT.dblChipSizeXDir = 1.0                          ' V1.2.0.0①チップサイズサイズx(mm)
            stPLT.dblChipSizeYDir = 1.0                          ' V1.2.0.0①チップサイズサイズy(mm)
            GoTo PLATE_END
        Else
            Input(fNum, a) : Input(fNum, a)                 ' 2行空読み
        End If
        Input(fNum, stLASER.iTrimAtt)                       ' ###1040③ アッテネータトリムデータへの保存
        Input(fNum, stLASER.dblRotPar)                      ' ###1040③ 減衰率(%)
        Input(fNum, stLASER.dblRotAtt)                      ' ###1040③ ロータリーアッテネータの回転量(0-FFF)
        Input(fNum, stLASER.iFixAtt)                        ' ###1040③ 固定アッテネータのON/OFF(0:OFF,1:ON)
        'V2.1.0.0②↓
        If VersionNumber >= 5 Then
            Input(fNum, stLASER.iAttNo)                     ' アッテネータNo.（0:指定無）'V2.1.0.0②
        Else
            stLASER.iAttNo=0
        End If
        'V2.1.0.0②↑
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' ###1040 3行空読み
        Input(fNum, stPLT.TeachBlockX)                      ' ###1040① ティーチング・ブロックX
        Input(fNum, stPLT.TeachBlockY)                      ' ###1040① ティーチング・ブロックY
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' ###1040 3行空読み
        Input(fNum, stPLT.StageSpeedY)                      ' ###1040④ Ｙ軸スピード
        Input(fNum, a)                                      ' ###1040 1行空読み
        'V1.2.0.0①↓
        If (a = "/") Then                                   ' ###1040旧データの場合は、ここで終了する。
            stPLT.dblChipSizeXDir = 1.0                          ' チップサイズサイズx(mm)
            stPLT.dblChipSizeYDir = 1.0                          ' チップサイズサイズy(mm)
            GoTo PLATE_END
        Else
            Input(fNum, a) : Input(fNum, a)                 ' 2行空読み
            Input(fNum, stPLT.dblChipSizeXDir)                   ' チップサイズサイズx(mm)
            Input(fNum, stPLT.dblChipSizeYDir)                   ' チップサイズサイズy(mm)
        End If
        'V2.0.0.0①↓
        If VersionNumber >= 3 Then
            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
            Input(fNum, stPLT.dblStepOffsetXDir)                ' ステップオフセット量X
            Input(fNum, stPLT.dblStepOffsetYDir)                ' ステップオフセット量Y
        Else
            stPLT.dblStepOffsetXDir = 0.0                       ' ステップオフセット量X
            stPLT.dblStepOffsetYDir = 0.0                       ' ステップオフセット量Y
            stPLT.StageSpeedY = stPLT.StageSpeedY / 2           ' V2.0.0.1②１#0005と#0050は、速度が異なるので１／２する。
        End If
        'V2.0.0.0①↑
        'V2.2.0.0②↓
        'V2.2.2.0②        If VersionNumber >= 6 Then
        If VersionNumber >= 7 Then      'V2.2.2.0②
            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
            Input(fNum, stPLT.dblStdMagnification)              ' デジタルカメラ倍率
        Else
            stPLT.dblStdMagnification = 1.0                     ' デジタルカメラ倍率 ：デフォルト
        End If
        'V2.2.0.0②↑

        Input(fNum, a)                                      ' 1行空読み
        'V1.2.0.0①↑
PLATE_END:
        If (a <> "/") Then
STP_SERR:
            strMSG = "データエラー (システムデータ) !!" & vbCrLf
            Call Z_PRINT(strMSG)
            GoTo rData_load_END
        End If

        'V2.1.0.0②↓
        If stLASER.iAttNo > 0 Then
            If UserSub.LaserCalibrationAttenuatorDataGet(stLASER.iAttNo, stLASER.dblRotPar, stLASER.iFixAtt, stLASER.dblRotAtt) Then
                Z_PRINT("アッテネータテーブルからNO=[" & stLASER.iAttNo.ToString & "][" & stLASER.dblRotPar.ToString & "%]の情報を取得しました。")
            Else
                Z_PRINT("アッテネータテーブルからの情報取得がエラーになりました。NO=[" & stLASER.iAttNo.ToString & "]")
                GoTo STP_SERR
            End If
        End If
        'V2.1.0.0②↑

        'V2.1.0.0②        If stLASER.iTrimAtt = 1 Then                                        ' ###1040⑥
        If stLASER.iTrimAtt = 1 OrElse stLASER.iAttNo > 0 Then                  'V2.1.0.0②
            gSysPrm.stRAT.gfAttRate = stLASER.dblRotPar                     ' ###1040⑥ 減衰率(%)
            gSysPrm.stRAT.giAttRot = stLASER.dblRotAtt                      ' ###1040⑥ ロータリーアッテネータの回転量(0-FFF)
            gSysPrm.stRAT.giAttFix = stLASER.iFixAtt                        ' ###1040⑥ 固定アッテネータ(0:OFF, 1:ON)
            Call DllSysPrmSysParam_definst.PutSysPrm_ROT_ATT(gSysPrm.stRAT) ' ###1040⑥
            Call Form1.SetATTRateToScreen(False)                            'V2.0.0.0⑮
        End If                                                              ' ###1040⑥

        '--------------------------------------------------------------------------
        '   データファイルより抵抗データを設定する
        '--------------------------------------------------------------------------
        On Error GoTo STP_RERR
        Input(fNum, a)                                      ' 1行空読み
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
        Input(fNum, stPLT.RCount)                           ' 抵抗数
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み

        ' 抵抗数分、以下のデータを設定する
        '  抵抗No., 抵抗名, ｽﾛｰﾌﾟ(+:1 -:2), リレービット, HP, LP, GP
        For i = 1 To stPLT.RCount
            If (i > MAXRNO) Then GoTo STP_RERR
            Input(fNum, z(0))                                               '抵抗Ｎｏ
            Input(fNum, stREG(i).strRNO)                                    '抵抗名
            Input(fNum, stREG(i).intSLP)                                    'スロープ
            Input(fNum, c(0))                                               'リレービット
            Input(fNum, stREG(i).intPRH)                                    'ＨＩ側プローブ
            Input(fNum, stREG(i).intPRL)                                    'ＬＯ側プローブ
            Input(fNum, stREG(i).intPRG)                                    '
            stREG(i).lngRel = Val(CStr(System.Math.Abs(CDbl("&H" & c(0))))) ' リレービット(HEX)
        Next i

        ' 抵抗数分、以下のデータを設定する
        If VersionNumber <= 2 Then                                           'V2.0.0.0⑧　旧バージョンの時
            '  抵抗No., 目標値, 単位, 判定(0:% 1:絶対値), ' 精度(0:高速 1:高精度),初期判定下限値, 初期判定上限値, 終了判定下限値, 終了判定上限値, ｶｯﾄ数, 測定
            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
            For i = 1 To stPLT.RCount
                If (i > MAXRNO) Then Exit For
                Input(fNum, z(0))                                               '抵抗Ｎｏ
                Input(fNum, stREG(i).dblNOM)                                    '目標値
                Input(fNum, stREG(i).strTANI)                                   '単位
                Input(fNum, stREG(i).intMode)                                   '判定
                Input(fNum, stREG(i).intTMM1)                                   '精度
                Input(fNum, stREG(i).dblITL)                                    '初期判定下限値
                Input(fNum, stREG(i).dblITH)                                    '初期判定上限値
                Input(fNum, stREG(i).dblFTL)                                    '終了判定下限値
                Input(fNum, stREG(i).dblFTH)                                    '終了判定上限値
                Input(fNum, stREG(i).intTNN)                                    'ｶｯﾄ数
                Input(fNum, stREG(i).intMType)                                  '測定機器
                'V2.0.0.0⑧↓
                If UserModule.IsMarking(i) Then                                 '測定モードマーキングはなし(0:なし, 1:ITのみ 2:FTのみ 3:IT,FT両方)
                    stREG(i).intMeasMode = MEAS_JUDGE_NONE
                ElseIf UserModule.IsMeasureOnly(i) Then                         '測定モード測定のみはＦＴのみ(0:なし, 1:ITのみ 2:FTのみ 3:IT,FT両方)
                    stREG(i).intMeasMode = MEAS_JUDGE_FT
                Else
                    stREG(i).intMeasMode = MEAS_JUDGE_BOTH
                End If
                stREG(i).intReMeas = 0
                stREG(i).intReMeas_Time = 0
                If stREG(i).intMType = 1 Then    ' 外部測定器
                    stREG(i).intITReMeas = 5
                Else
                    stREG(i).intITReMeas = 2
                End If
                If stREG(i).intMType = 1 And stREG(i).intSLP = SLP_RMES Then    ' 外部測定器
                    stREG(i).intFTReMeas = 5
                Else
                    stREG(i).intFTReMeas = 2
                End If
                'V2.0.0.0⑧↑
                'V2.0.0.0②↓
                For j = 1 To EXTEQU
                    stREG(i).intOnExtEqu(j) = 0                         ' ＯＮ機器
                Next
                For j = 1 To EXTEQU
                    stREG(i).intOffExtEqu(j) = 0                          ' ＯＦＦ機器
                Next
                'V2.0.0.0②↑
            Next i
            'V2.0.0.0⑧↓
        Else
            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
            For i = 1 To stPLT.RCount
                If (i > MAXRNO) Then Exit For
                Input(fNum, z(0))                                               '抵抗Ｎｏ
                Input(fNum, stREG(i).dblNOM)                                    '目標値
                Input(fNum, stREG(i).strTANI)                                   '単位
                Input(fNum, stREG(i).intMode)                                   '判定
                Input(fNum, stREG(i).intMeasMode)                               '測定モード(0:なし, 1:ITのみ 2:FTのみ 3:IT,FT両方)
                Input(fNum, stREG(i).dblITL)                                    '初期判定下限値
                Input(fNum, stREG(i).dblITH)                                    '初期判定上限値
                Input(fNum, stREG(i).dblFTL)                                    '終了判定下限値
                Input(fNum, stREG(i).dblFTH)                                    '終了判定上限値
                Input(fNum, stREG(i).intTNN)                                    'ｶｯﾄ数
                Input(fNum, stREG(i).intMType)                                  '測定機器
                Input(fNum, stREG(i).intTMM1)                                   '精度
                Input(fNum, stREG(i).intReMeas)                                 '再測定回数（0:再測定無 1:再測定回数）
                Input(fNum, stREG(i).intReMeas_Time)                            '再測定前ポーズ時間
            Next i
            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
            For i = 1 To stPLT.RCount
                If (i > MAXRNO) Then Exit For
                Input(fNum, z(0))                                               '抵抗Ｎｏ
                Input(fNum, stREG(i).intITReMeas)                               'イニシャル抵抗再測定回数(IT測定回数)
                Input(fNum, stREG(i).intFTReMeas)                               'ファイナル抵抗再測定回数(FT測定回数)
                Input(fNum, stREG(i).intCircuitNo)                              'サーキット番号
            Next i
            'V2.0.0.0②↓
            ' 抵抗数分、以下のデータを設定する
            ' 抵抗Ｎｏ ＯＮ機器１ ＯＮ機器２ ＯＮ機器３ ＯＦＦ機器１ ＯＦＦ機器２ ＯＦＦ機器３
            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
            For i = 1 To stPLT.RCount
                If (i > MAXRNO) Then Exit For
                Input(fNum, z(0))                                               '抵抗Ｎｏ
                For j = 1 To EXTEQU
                    Input(fNum, stREG(i).intOnExtEqu(j))                        ' ＯＮ機器
                Next
                For j = 1 To EXTEQU
                    Input(fNum, stREG(i).intOffExtEqu(j))                       ' ＯＦＦ機器
                Next
            Next i
            'V2.0.0.0②↑
        End If                                                              'V2.0.0.0⑧
        'V2.0.0.0⑧↑

        Input(fNum, a)                                      ' 1行空読み
        If (a <> "/") Then
STP_RERR:
            strMSG = "データエラー (抵抗データ) !!" & vbCrLf
            Call Z_PRINT(strMSG)
            GoTo rData_load_END
        End If

        '--------------------------------------------------------------------------
        '   データファイルよりカットデータを設定する
        '--------------------------------------------------------------------------
        On Error GoTo STP_CERR
        Input(fNum, a)                                      ' 1行空読み
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み

        ' 抵抗数分、カットデータを設定する
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
        For i = 1 To stPLT.RCount
            ' 抵抗内カット数分、以下のカットデータを設定する
            '  抵抗No., ｶｯﾄNo., ｶｯﾄ方法, ｶｯﾄ形状, 本数, STX(mm), STY(mm), STX2(mm), STY2(mm),
            '  CUTOFF(%), MD, Qﾚｰﾄ(.1K), 速度(mm/sec), DL1(mm), DL2(mm)
            For j = 1 To stREG(i).intTNN
                If (j > MAXCTN) Then GoTo STP_CERR
                Input(fNum, z(0))                               '抵抗Ｎｏ
                Input(fNum, z(1))                               'カットＮｏ
                Input(fNum, stREG(i).STCUT(j).intCUT)           'カット方法
                Input(fNum, stREG(i).STCUT(j).intCTYP)          'カット形状
                Input(fNum, stREG(i).STCUT(j).intNum)           '本数（サーペンタイン用）
                Input(fNum, stREG(i).STCUT(j).dblSTX)           'スタートＰ＿Ｘ(mm)
                Input(fNum, stREG(i).STCUT(j).dblSTY)           'スタートＰ＿Ｙ(mm)
                Input(fNum, stREG(i).STCUT(j).dblSX2)           'スタートＰ＿Ｘ(mm)　（サーペンタイン用）
                Input(fNum, stREG(i).STCUT(j).dblSY2)           'スタートＰ＿Ｙ(mm)　（サーペンタイン用）
                Input(fNum, stREG(i).STCUT(j).dblCOF)           'カットオフ
                Input(fNum, stREG(i).STCUT(j).intTMM)           '測定モード
                Input(fNum, stREG(i).STCUT(j).intQF1)           'Ｑレート
                Input(fNum, stREG(i).STCUT(j).dblV1)            '速度
                Input(fNum, stREG(i).STCUT(j).dblDL2)           'カット長１
                Input(fNum, stREG(i).STCUT(j).dblDL3)           'カット長２
            Next j
        Next i

        ' 抵抗数分、カットデータを設定する
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
        For i = 1 To stPLT.RCount
            ' 抵抗内カット数分、以下のカットデータを設定する
            '  抵抗No., ｶｯﾄNo., ANG1(°), ANG2(°), LTP(%), 測定機器(0=内部測定, 1～=外部測定)
            For j = 1 To stREG(i).intTNN
                If (j > MAXCTN) Then GoTo STP_CERR
                Input(fNum, z(0))                               '抵抗Ｎｏ
                Input(fNum, z(1))                               'カットＮｏ
                Input(fNum, stREG(i).STCUT(j).intANG)           'カット方向１
                Input(fNum, stREG(i).STCUT(j).intANG2)          'カット方向２
                Input(fNum, stREG(i).STCUT(j).dblLTP)           'ＬターンＰ
                Input(fNum, stREG(i).STCUT(j).intMType)         '測定機器
            Next j
        Next i

        ' 抵抗数分、カットデータを設定する
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
        For i = 1 To stPLT.RCount
            ' 抵抗内カット数分、以下のカットデータを設定する
            '  抵抗No., ｶｯﾄNo., IX回数1, IX回数2, IX回数3, IX回数4, IX回数5
            '                   ﾋﾟｯﾁ1(mm), ﾋﾟｯﾁ2(mm), ﾋﾟｯﾁ3(mm), ﾋﾟｯﾁ4(mm), ﾋﾟｯﾁ5(mm)
            For j = 1 To stREG(i).intTNN
                If (j > MAXCTN) Then GoTo STP_CERR
                Input(fNum, z(0))                               '抵抗Ｎｏ
                Input(fNum, z(1))                               'カットＮｏ
                For k = 1 To MAXIDX
                    Input(fNum, stREG(i).STCUT(j).intIXN(k))        'IX回数１
                    'Input(fNum, stREG(i).STCUT(j).intIXN(2))        'IX回数２
                    'Input(fNum, stREG(i).STCUT(j).intIXN(3))        'IX回数３
                    'Input(fNum, stREG(i).STCUT(j).intIXN(4))        'IX回数４
                    'Input(fNum, stREG(i).STCUT(j).intIXN(5))        'IX回数５
                Next k

                For k = 1 To MAXIDX
                    Input(fNum, stREG(i).STCUT(j).dblDL1(k))        'ﾋﾟｯﾁ１
                    'Input(fNum, stREG(i).STCUT(j).dblDL1(2))        'ﾋﾟｯﾁ２
                    'Input(fNum, stREG(i).STCUT(j).dblDL1(3))        'ﾋﾟｯﾁ３
                    'Input(fNum, stREG(i).STCUT(j).dblDL1(4))        'ﾋﾟｯﾁ４
                    'Input(fNum, stREG(i).STCUT(j).dblDL1(5))        'ﾋﾟｯﾁ５
                Next k
            Next j
        Next i

        ' 抵抗数分、カットデータを設定する
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
        For i = 1 To stPLT.RCount
            ' 抵抗内カット数分、以下のカットデータを設定する
            '  抵抗No., ｶｯﾄNo., PAU1(ms), PAU2(ms), PAU3(ms), PAU4(ms), PAU5(ms)
            '                   誤差1(%), 誤差2(%), 誤差3(%), 誤差4(%), 誤差5(%)
            For j = 1 To stREG(i).intTNN
                If (j > MAXCTN) Then GoTo STP_CERR
                Input(fNum, z(0))                               '抵抗Ｎｏ
                Input(fNum, z(1))                               'カットＮｏ

                For k = 1 To MAXIDX
                    Input(fNum, stREG(i).STCUT(j).lngPAU(k))        'PAU1-5(ms)
                    'Input(fNum, stREG(i).STCUT(j).lngPAU(2))        'PAU2(ms)
                    'Input(fNum, stREG(i).STCUT(j).lngPAU(3))        'PAU3(ms)
                    'Input(fNum, stREG(i).STCUT(j).lngPAU(4))        'PAU4(ms)
                    'Input(fNum, stREG(i).STCUT(j).lngPAU(5))        'PAU5(ms)
                Next k

                For k = 1 To MAXIDX
                    Input(fNum, stREG(i).STCUT(j).dblDEV(k))        '誤差1-5(%)
                    'Input(fNum, stREG(i).STCUT(j).dblDEV(2))        '誤差2(%)
                    'Input(fNum, stREG(i).STCUT(j).dblDEV(3))        '誤差3(%)
                    'Input(fNum, stREG(i).STCUT(j).dblDEV(4))        '誤差4(%)
                    'Input(fNum, stREG(i).STCUT(j).dblDEV(5))        '誤差5(%)
                Next k
            Next j
        Next i


        ' 抵抗数分、カットデータを設定する
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
        For i = 1 To stPLT.RCount
            ' 抵抗内カット数分、以下のカットデータを設定する
            '  抵抗No., ｶｯﾄNo., 測定機器1, 測定機器2, 測定機器3, 測定機器4, 測定機器5
            '                   測定モード1, 測定モード2, 測定モード3, 測定モード4, 測定モード5
            For j = 1 To stREG(i).intTNN
                If (j > MAXCTN) Then GoTo STP_CERR
                Input(fNum, z(0))                               '抵抗No.
                Input(fNum, z(1))                               'ｶｯﾄNo.

                For k = 1 To MAXIDX
                    Input(fNum, stREG(i).STCUT(j).intIXMType(k))    '測定機器1-5
                    'Input(fNum, stREG(i).STCUT(j).intIXMType(2))    '測定機器2
                    'Input(fNum, stREG(i).STCUT(j).intIXMType(3))    '測定機器3
                    'Input(fNum, stREG(i).STCUT(j).intIXMType(4))    '測定機器4
                    'Input(fNum, stREG(i).STCUT(j).intIXMType(5))    '測定機器5
                Next k

                For k = 1 To MAXIDX
                    Input(fNum, stREG(i).STCUT(j).intIXTMM(k))      '測定モード1-5
                    'Input(fNum, stREG(i).STCUT(j).intIXTMM(2))      '測定モード2
                    'Input(fNum, stREG(i).STCUT(j).intIXTMM(3))      '測定モード3
                    'Input(fNum, stREG(i).STCUT(j).intIXTMM(4))      '測定モード4
                    'Input(fNum, stREG(i).STCUT(j).intIXTMM(5))      '測定モード5
                Next k
            Next j
        Next i


        ' 抵抗数分、カットデータを設定する
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
        For i = 1 To stPLT.RCount
            ' 抵抗内カット数分、以下のカットデータを設定する
            '  抵抗No., ｶｯﾄNo., 加工条件No.1, 加工条件No.2, 加工条件No.3, 加工条件No.4
            For j = 1 To stREG(i).intTNN
                If (j > MAXCTN) Then GoTo STP_CERR
                Input(fNum, z(0))                               '抵抗No.
                Input(fNum, z(1))                               'ｶｯﾄNo.
                For k = 1 To MAXCND
                    Input(fNum, stREG(i).STCUT(j).intCND(k))    '加工条件1-5
                Next k
            Next j
        Next i

        Input(fNum, a)                                      ' 1行空読み
        ' V1.0.4.3③ ADD START ↓
        ' 初期化
        For i = 1 To stPLT.RCount
            For j = 1 To stREG(i).intTNN
                If (j > MAXCTN) Then GoTo STP_CERR
                For k = 1 To MAX_LCUT
                    stREG(i).STCUT(j).dCutLen(k) = 0.001        'カット長1-7
                    stREG(i).STCUT(j).dQRate(k) = 0.1           'Ｑレート1-7
                    stREG(i).STCUT(j).dSpeed(k) = 0.1           '速度1-7
                    stREG(i).STCUT(j).dAngle(k) = 0             '角度1-7
                    stREG(i).STCUT(j).dTurnPoint(k) = 0.0       'ターンポイント1-7
                    stREG(i).STCUT(j).intQF2 = 0.1              'リトレース・Ｑレート
                    stREG(i).STCUT(j).dblV2 = 0.1               'リトレース・速度
                    stREG(i).STCUT(j).cFormat = ""              'マーキング文字
                    'V2.2.1.7① ↓
                    stREG(i).STCUT(j).cMarkFix = ""              '印字固定部
                    stREG(i).STCUT(j).cMarkStartNum = ""         '開始番号
                    stREG(i).STCUT(j).intMarkRepeatCnt = 0       '重複回数

                    'If (VersionNumber >= 8) Then
                    '    stREG(i).STCUT(j).cMarkFix = ""              '印字固定部
                    '    stREG(i).STCUT(j).cMarkStartNum = ""         '開始番号
                    '    stREG(i).STCUT(j).intMarkRepeatCnt = 0       '重複回数
                    'End If
                    'V2.2.1.7① ↑
                Next k
            Next j
        Next i
        If a.IndexOf("- L CUT PARAMETER -") >= 0 Then       ' Ｌカットパラメータ追加先頭行に"- L CUT PARAMETER -"の文字を挿入する。
            If VersionNumber < 2 Then                               'V2.0.0.0⑦
                VersionNumber = 2                                   'V2.0.0.0⑦２番目のバージョン（１は無し）
            End If                                                  'V2.0.0.0⑦
            ' Ｌカットのカット長
            Input(fNum, a) : Input(fNum, a)    ' 2行空読み
            For i = 1 To stPLT.RCount
                For j = 1 To stREG(i).intTNN
                    If (j > MAXCTN) Then GoTo STP_CERR
                    Input(fNum, z(0))                               '抵抗No.
                    Input(fNum, z(1))                               'ｶｯﾄNo.
                    For k = 1 To MAX_LCUT
                        Input(fNum, stREG(i).STCUT(j).dCutLen(k))    'カット長1-7
                    Next k
                Next j
            Next i

            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
            ' ＬカットのＱレート
            For i = 1 To stPLT.RCount
                For j = 1 To stREG(i).intTNN
                    If (j > MAXCTN) Then GoTo STP_CERR
                    Input(fNum, z(0))                               '抵抗No.
                    Input(fNum, z(1))                               'ｶｯﾄNo.
                    For k = 1 To MAX_LCUT
                        Input(fNum, stREG(i).STCUT(j).dQRate(k))    'Ｑレート1-7
                    Next k
                Next j
            Next i

            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
            ' Ｌカットの速度
            For i = 1 To stPLT.RCount
                For j = 1 To stREG(i).intTNN
                    If (j > MAXCTN) Then GoTo STP_CERR
                    Input(fNum, z(0))                               '抵抗No.
                    Input(fNum, z(1))                               'ｶｯﾄNo.
                    For k = 1 To MAX_LCUT
                        Input(fNum, stREG(i).STCUT(j).dSpeed(k))    '速度1-7
                    Next k
                Next j
            Next i

            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
            ' Ｌカットの角度
            For i = 1 To stPLT.RCount
                For j = 1 To stREG(i).intTNN
                    If (j > MAXCTN) Then GoTo STP_CERR
                    Input(fNum, z(0))                               '抵抗No.
                    Input(fNum, z(1))                               'ｶｯﾄNo.
                    For k = 1 To MAX_LCUT
                        Input(fNum, stREG(i).STCUT(j).dAngle(k))    '角度1-7
                    Next k
                Next j
            Next i

            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
            ' Ｌカットの角度
            For i = 1 To stPLT.RCount
                For j = 1 To stREG(i).intTNN
                    If (j > MAXCTN) Then GoTo STP_CERR
                    Input(fNum, z(0))                               '抵抗No.
                    Input(fNum, z(1))                               'ｶｯﾄNo.
                    For k = 1 To MAX_LCUT - 1                       ' ターンポイントは、カット数より１つ少ない
                        Input(fNum, stREG(i).STCUT(j).dTurnPoint(k))    'ターンポイント1-7
                    Next k
                Next j
            Next i


            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
            ' ストレートカット・リトレースのＱレート、速度、マーキングの文字
            For i = 1 To stPLT.RCount
                For j = 1 To stREG(i).intTNN
                    If (j > MAXCTN) Then GoTo STP_CERR
                    Input(fNum, z(0))                               '抵抗No.
                    Input(fNum, z(1))                               'ｶｯﾄNo.
                    Input(fNum, stREG(i).STCUT(j).intQF2)           'Ｑレート
                    Input(fNum, stREG(i).STCUT(j).dblV2)            '速度
                    Input(fNum, stREG(i).STCUT(j).cFormat)    'ターンポイント1-7
                    'V2.2.1.7① ↓
                    'V2.2.2.0②                     If (VersionNumber >= 8) Then
                    If (VersionNumber >= 8) OrElse (VersionNumber = 6) Then            'V2.2.2.0② 
                        Input(fNum, stREG(i).STCUT(j).cMarkFix)              '印字固定部
                        Input(fNum, stREG(i).STCUT(j).cMarkStartNum)         '開始番号
                        Input(fNum, stREG(i).STCUT(j).intMarkRepeatCnt)      '重複回数
                    End If
                    'V2.2.1.7① ↑
                Next j
            Next i

            Input(fNum, a)                                      ' 1行空読み
            'V2.0.0.0⑦            If (a <> "/") Then
            'V2.0.0.0⑦                GoTo STP_CERR
            'V2.0.0.0⑦            End If
            ' V1.0.4.3③ ADD END ↑
            'V2.0.0.0⑦ ADD START↓

            ' 初期化
            For i = 1 To stPLT.RCount
                For j = 1 To stREG(i).intTNN
                    If (j > MAXCTN) Then GoTo STP_CERR
                    For k = 1 To MAX_RETRACECUT
                        stREG(i).STCUT(j).dblRetraceOffX(k) = 0.0       'リトレースのオフセットＸ
                        stREG(i).STCUT(j).dblRetraceOffY(k) = 0.0       'リトレースのオフセットＹ
                        stREG(i).STCUT(j).dblRetraceQrate(k) = 10.0     'リトレース・Ｑレート
                        stREG(i).STCUT(j).dblRetraceSpeed(k) = 10.0     'リトレース・速度
                        If (CNS_CUTP_ST_TR = stREG(i).STCUT(j).intCTYP) And k = 1 Then
                            stREG(i).STCUT(j).dblRetraceOffX(k) = stREG(i).STCUT(j).dblSX2      'リトレースのオフセットＸ
                            stREG(i).STCUT(j).dblRetraceOffY(k) = stREG(i).STCUT(j).dblSY2      'リトレースのオフセットＹ
                            stREG(i).STCUT(j).dblRetraceQrate(k) = stREG(i).STCUT(j).intQF2     'リトレース・Ｑレート
                            stREG(i).STCUT(j).dblRetraceSpeed(k) = stREG(i).STCUT(j).dblV2      'リトレース・速度
                        End If
                    Next k
                    stREG(i).STCUT(j).intRetraceCnt = 1
                Next j
            Next i
            If a.IndexOf("- RETRACE CUT PARAMETER -") >= 0 Then       ' リトレースカットパラメータ追加先頭行に"- RETRACE CUT PARAMETER -"の文字を挿入する。

                '  Dim intRetraceCnt As Short                          ' リトレースカット本数
                Input(fNum, a) : Input(fNum, a)    ' 2行空読み
                For i = 1 To stPLT.RCount
                    For j = 1 To stREG(i).intTNN
                        If (j > MAXCTN) Then GoTo STP_CERR
                        Input(fNum, z(0))                               '抵抗No.
                        Input(fNum, z(1))                               'ｶｯﾄNo.
                        Input(fNum, stREG(i).STCUT(j).intRetraceCnt)    'リトレースカット本数
                    Next j
                Next i

                ' リトレースのオフセットＸ
                Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
                For i = 1 To stPLT.RCount
                    For j = 1 To stREG(i).intTNN
                        If (j > MAXCTN) Then GoTo STP_CERR
                        Input(fNum, z(0))                               '抵抗No.
                        Input(fNum, z(1))                               'ｶｯﾄNo.
                        For k = 1 To MAX_RETRACECUT
                            Input(fNum, stREG(i).STCUT(j).dblRetraceOffX(k))    'リトレースのオフセットＸ
                        Next k
                    Next j
                Next i

                Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
                ' リトレースのオフセットＹ
                For i = 1 To stPLT.RCount
                    For j = 1 To stREG(i).intTNN
                        If (j > MAXCTN) Then GoTo STP_CERR
                        Input(fNum, z(0))                               '抵抗No.
                        Input(fNum, z(1))                               'ｶｯﾄNo.
                        For k = 1 To MAX_RETRACECUT
                            Input(fNum, stREG(i).STCUT(j).dblRetraceOffY(k))    'リトレースのオフセットＹ
                        Next k
                    Next j
                Next i

                Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
                ' ストレートカット・リトレースのQレート
                For i = 1 To stPLT.RCount
                    For j = 1 To stREG(i).intTNN
                        If (j > MAXCTN) Then GoTo STP_CERR
                        Input(fNum, z(0))                               '抵抗No.
                        Input(fNum, z(1))                               'ｶｯﾄNo.
                        For k = 1 To MAX_RETRACECUT
                            Input(fNum, stREG(i).STCUT(j).dblRetraceQrate(k))    'ストレートカット・リトレースのQレート
                        Next k
                    Next j
                Next i

                Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
                ' ストレートカット・リトレースのトリム速度
                For i = 1 To stPLT.RCount
                    For j = 1 To stREG(i).intTNN
                        If (j > MAXCTN) Then GoTo STP_CERR
                        Input(fNum, z(0))                               '抵抗No.
                        Input(fNum, z(1))                               'ｶｯﾄNo.
                        For k = 1 To MAX_RETRACECUT
                            Input(fNum, stREG(i).STCUT(j).dblRetraceSpeed(k))    'ストレートカット・リトレースのトリム速度
                        Next k
                    Next j
                Next i
                Input(fNum, a)                                      ' 1行空読み
                    'V2.1.0.0①                If (a <> "/") Then
                    If (VersionNumber < 5 AndAlso a <> "/") Then        'V2.1.0.0①
                        GoTo STP_CERR
                    End If
                ElseIf (a <> "/") Then
                    GoTo STP_CERR
            End If
            'V2.0.0.0⑦ ADD END  ↑
            'V1.0.4.3③            If (a <> "/") Then
        ElseIf (a <> "/") Then
            GoTo STP_CERR
        End If   'V2.1.0.0①

        'V2.1.0.0①↓ カット毎の抵抗値変化量判定機能追加
        If VersionNumber >= 5 Then
            Input(fNum, a) : Input(fNum, a)    ' ２行空読み１行目は、"- RETRACE CUT PARAMETER -"の最後に読み込み済み
            For i = 1 To stPLT.RCount
                For j = 1 To stREG(i).intTNN
                    If (j > MAXCTN) Then GoTo STP_CERR
                    Input(fNum, z(0))                               '抵抗No.
                    Input(fNum, z(1))                               'ｶｯﾄNo.
                    Input(fNum, stREG(i).STCUT(j).iVariationRepeat) ' リピート有無
                    Input(fNum, stREG(i).STCUT(j).iVariation)       ' 判定有無
                    Input(fNum, stREG(i).STCUT(j).dRateOfUp)        ' 上昇率
                    Input(fNum, stREG(i).STCUT(j).dVariationLow)    ' 下限値
                    Input(fNum, stREG(i).STCUT(j).dVariationHi)     ' 上限値
                Next j
            Next i
        End If

        'V2.2.0.0②↓
        ' Uカットパラメータの追加 
        'V2.2.2.0②         If VersionNumber >= 6 Then
        If VersionNumber >= 7 Then        'V2.2.2.0② 
            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み１行目
            For i = 1 To stPLT.RCount
                For j = 1 To stREG(i).intTNN
                    If (j > MAXCTN) Then GoTo STP_CERR
                    Input(fNum, z(0))                               '抵抗No.
                    Input(fNum, z(1))                               'ｶｯﾄNo.
                    Input(fNum, stREG(i).STCUT(j).dUCutL1)          ' L1
                    Input(fNum, stREG(i).STCUT(j).dUCutL2)          ' L2
                    Input(fNum, stREG(i).STCUT(j).intUCutQF1)       ' Qレート1
                    Input(fNum, stREG(i).STCUT(j).dblUCutV1)        ' 速度1
                    Input(fNum, stREG(i).STCUT(j).intUCutANG)       ' 角度
                    Input(fNum, stREG(i).STCUT(j).dblUCutTurnP)     ' ターンポイント
                    Input(fNum, stREG(i).STCUT(j).intUCutTurnDir)   ' ターン方向
                    Input(fNum, stREG(i).STCUT(j).dblUCutR1)        ' R1
                    Input(fNum, stREG(i).STCUT(j).dblUCutR2)        ' R2
                Next j
            Next i
        Else
            For i = 1 To stPLT.RCount
                For j = 1 To stREG(i).intTNN
                    If (j > MAXCTN) Then GoTo STP_CERR
                    stREG(i).STCUT(j).dUCutL1 = 0.0          ' L1
                    stREG(i).STCUT(j).dUCutL2 = 0.0          ' L2
                    stREG(i).STCUT(j).intUCutQF1 = 0.1       ' Qレート
                    stREG(i).STCUT(j).dblUCutV1 = 0.1        ' 速度
                    stREG(i).STCUT(j).intUCutANG = 0         ' 角度
                    stREG(i).STCUT(j).dblUCutTurnP = 0       ' ターンポイント
                    stREG(i).STCUT(j).intUCutTurnDir = 1     ' CW
                    stREG(i).STCUT(j).dblUCutR1 = 0          ' R1
                    stREG(i).STCUT(j).dblUCutR2 = 0          ' R2

                Next j
            Next i
        End If
        'V2.2.0.0②↑

        If VersionNumber >= 5 Then
            Input(fNum, a)                              ' 1行空読み
        End If
        If (a <> "/") Then
            'V2.1.0.0①↑
STP_CERR:
            strMSG = "データエラー (カットデータ) !!" & vbCrLf
            Call Z_PRINT(strMSG)
            GoTo rData_load_END
        End If

        '--------------------------------------------------------------------------
        '   データファイルよりパターン登録データ(カット位置補正用)を設定する
        '--------------------------------------------------------------------------
        On Error GoTo STP_PERR
        Input(fNum, a) ' 1行空読み
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
        Input(fNum, stPLT.PtnCount)                         ' ﾊﾟﾀｰﾝ登録数
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み

        ' ﾊﾟﾀｰﾝ登録数分、以下のデータを設定する
        For i = 1 To stPLT.PtnCount
            If (i > MAXRGN) Then GoTo STP_PERR
            Input(fNum, stPTN(i).intGRP)
            If stPTN(i).intGRP = 0 Then
                stPTN(i).intGRP = 1
            End If
            Input(fNum, stPTN(i).intPTN)
            If stPTN(i).intPTN = 0 Then
                stPTN(i).intPTN = 1
            End If
            Input(fNum, stPTN(i).dblPosX)
            Input(fNum, stPTN(i).dblPosY)
            Input(fNum, stPTN(i).PtnFlg)                    ' ﾊﾟﾀｰﾝ認識(1:有り, 0:無し, 2:手動)

            stPTN(i).dblDRX = 0                             ' ズレ量保存ワークX
            stPTN(i).dblDRY = 0                             ' ズレ量保存ワークY
        Next i

        Input(fNum, a)                                      ' 1行空読み
        If (a <> "/") Then
STP_PERR:
            strMSG = "データエラー (パターン登録データ【カット位置補正用】) !!" & vbCrLf
            Call Z_PRINT(strMSG)
            GoTo rData_load_END
        End If

        '--------------------------------------------------------------------------
        '   データファイルよりパターン登録データ(ＸＹθ補正用)を設定する
        '--------------------------------------------------------------------------
        stThta.iPP30 = 1                                    ' 補正モード=手動
        stThta.iPP31 = 0                                    ' 補正なし

        On Error GoTo STP_P2ERR
        Input(fNum, a) ' 1行空読み
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)        ' 3行空読み
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)        ' 3行空読み
        Input(fNum, stThta.iPP30)                               ' 補正モード(0:自動,1:手動)
        Input(fNum, stThta.iPP31)                               ' 補正方法(0:補正なし, 1:補正あり)
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)        ' 3行空読み
        Input(fNum, stThta.iPP38)                               ' ｸﾞﾙｰﾌﾟ番号1
        Input(fNum, stThta.iPP37_1)                             ' ﾊﾟﾀｰﾝ番号1
        Input(fNum, stThta.fpp32_x)                             ' 登録位置1X(mm)
        Input(fNum, stThta.fpp32_y)                             ' 登録位置1Y(mm)
        Input(fNum, stThta.iPP38)                               ' ｸﾞﾙｰﾌﾟ番号2
        Input(fNum, stThta.iPP37_2)                             ' ﾊﾟﾀｰﾝ番号2
        Input(fNum, stThta.fpp33_x)                             ' 登録位置2X(mm)
        Input(fNum, stThta.fpp33_y)                             ' 登録位置2Y(mm)
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)        ' 3行空読み
        Input(fNum, stThta.fTheta)                              ' θ軸角度
        Input(fNum, stThta.fPP53Min)                            ' 最小角度
        Input(fNum, stThta.fPP53Max)                            ' 最大角度
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)        ' 3行空読み
        Input(fNum, stThta.fpp34_x)                             ' 補正ポジションオフセットX
        Input(fNum, stThta.fpp34_y)                             ' 補正ポジションオフセットY

        Input(fNum, a)                                      ' 1行空読み
        If (a <> "/") Then
STP_P2ERR:
            strMSG = "データエラー (パターン登録データ【ＸＹθ補正用】) !!" + vbCrLf
            Call Z_PRINT(strMSG)
            GoTo rData_load_END
        End If

        '--------------------------------------------------------------------------
        '   データファイルよりＧＰＩＢデータを設定する
        '--------------------------------------------------------------------------
        On Error GoTo STP_GERR
        Input(fNum, a)                                      ' 1行空読み
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
        Input(fNum, stPLT.GCount)                           ' 制御数
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み

        ' GPIB制御数分、以下のデータを設定する
        For i = 1 To stPLT.GCount
            If (i > MAXGNO) Then GoTo STP_GERR
            Input(fNum, z(0))
            Input(fNum, stGPIB(i).strGNAM)
            Input(fNum, stGPIB(i).intGAD)
            Input(fNum, stGPIB(i).intDLM)
            'V2.0.0.0④            Input(fNum, stGPIB(i).strCCMD)
            'V2.0.0.0④↓
            If VersionNumber <= 2 Then
                Input(fNum, stGPIB(i).strCCMD1)
                stGPIB(i).strCCMD2 = ""
                stGPIB(i).strCCMD3 = ""
            Else
                Input(fNum, stGPIB(i).strCCMD1)
                Input(fNum, stGPIB(i).strCCMD2)
                Input(fNum, stGPIB(i).strCCMD3)
            End If
            'V2.0.0.0④↑
            Input(fNum, stGPIB(i).strCTRG)
        Next i

        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
        ' GPIB制御数分、以下のデータを設定する
        For i = 1 To stPLT.GCount
            If (i > MAXGNO) Then GoTo STP_GERR
            Input(fNum, z(0))
            Input(fNum, stGPIB(i).strCON)
            Input(fNum, stGPIB(i).lngPOWON)
            Input(fNum, stGPIB(i).strCOFF)
            Input(fNum, stGPIB(i).lngPOWOFF)
        Next i

        Input(fNum, a)                                      ' 1行空読み
        If (a <> "/") Then
STP_GERR:
            strMSG = "データエラー (ＧＰＩＢデータ) !!" & vbCrLf
            Call Z_PRINT(strMSG)
            GoTo rData_load_END
        End If

        '--------------------------------------------------------------------------
        '   データファイルよりメッセージを設定する
        '--------------------------------------------------------------------------
        On Error GoTo STP_MERR
        Input(fNum, a)                                      ' 1行空読み
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み
        Input(fNum, a)                                      ' 1行空読み
        ' タイトルを設定する
        For i = 0 To 9
            Input(fNum, TTL_Msg(i))
        Next i

        TTL_Msg(0) = "トリミング"
        TTL_Msg(1) = "測定"
        TTL_Msg(2) = "カット実行"
        TTL_Msg(3) = "ステップ＆リピート"
        TTL_Msg(4) = "測定マーキングモード"
        TTL_Msg(5) = "電源モード"
        TTL_Msg(6) = "測定値変動測定"

        Input(fNum, a)                                      ' 1行空読み
        If (a <> "/") Then
STP_MERR:
            strMSG = "データエラー (メッセージデータ) !!" & vbCrLf
            Call Z_PRINT(strMSG)
            GoTo rData_load_END
        End If

        '--------------------------------------------------------------------------
        '   データファイルよりユーザデータを設定する
        '--------------------------------------------------------------------------
        On Error GoTo STP_USERERR
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　●ユーザデータタイトル
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル
        Input(fNum, stUserData.iTrimType)                   ' 製品種別

        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル
        Input(fNum, stUserData.sLotNumber)                  ' ロット番号
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル
        Input(fNum, stUserData.sOperator)                   ' オペレータ名
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル
        Input(fNum, stUserData.sPatternNo)                  ' パターンＮｏ．
        Input(fNum, stUserData.sProgramNo)                  ' プログラムＮｏ．
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル
        Input(fNum, stUserData.iTrimSpeed)                  ' トリミング速度 1:高速、 2:高精度、3:設定値
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル
        Input(fNum, stUserData.iLotChange)                  ' ロット終了条件 0:終了条件判定無し 1:枚数 2:ローダー信号 3:両方
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル
        Input(fNum, stUserData.lLotEndSL)                   ' ロット処理枚数
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル
        Input(fNum, stUserData.lCutHosei)                   ' カット位置補正頻度
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル
        Input(fNum, stUserData.lPrintRes)                   ' ロット終了時印刷素子数

        'V2.0.0.0⑩↓
        If VersionNumber <= 2 Then
            For i = 1 To stPLT.RCount
                If (i > MAXRNO) Then Exit For
                If UserSub.IsTrimType3() Then
                    stREG(i).intCircuitNo = i                                       'チップ抵抗はサーキット番号は抵抗番号で初期化
                Else
                    stREG(i).intCircuitNo = 1                                       'サーキット番号１で初期化
                End If
            Next i
        End If
        'V2.0.0.0⑩↑

        If VersionNumber <= 2 Then                              'V2.0.0.0⑪
            ' 温度センサー 
            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル
            Input(fNum, stUserData.iTempResUnit)                ' 抵抗レンジ 1:Ω, 2:KΩ
            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル
            Input(fNum, stUserData.iTempTemp)                   ' 参照温度 １：０℃ または ２：２５℃

            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル
            Input(fNum, stUserData.dStandardRes0)               ' 標準抵抗値   ０℃
            Input(fNum, stUserData.dStandardRes25)              ' 標準抵抗値 ２５℃
            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル
            Input(fNum, stUserData.dResTempCoff)                ' 抵抗温度係数")を変更するカットNo.
            'V2.0.0.0⑪↓
            If stUserData.iTempTemp = 1 Then
                stUserData.iTempTemp = 0
            Else
                stUserData.iTempTemp = 25
            End If
            stUserData.intClampVacume = 1                       ' クランプと吸着有り
            stUserData.dTemperatura0 = 0.0                      ' ０℃
            stUserData.dDaihyouAlpha = 0.0                      ' 代表α値
            stUserData.dDaihyouBeta = 0.0                       ' 代表β値
            stUserData.dAlpha = 0.0                             ' α値
            stUserData.dBeta = 0.0                              ' β値
        Else
            'V2.0.0.0⑭↓
            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル
            Input(fNum, stUserData.intClampVacume)              ' クランプと吸着の有無
            'V2.0.0.0⑭↑
            ' 温度センサー 
            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル
            Input(fNum, stUserData.iTempResUnit)                ' 抵抗レンジ 1:Ω, 2:KΩ
            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル
            Input(fNum, stUserData.iTempTemp)                   ' 参照温度

            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル
            Input(fNum, stUserData.dTemperatura0)               ' ０℃
            Input(fNum, stUserData.dDaihyouAlpha)               ' 代表α値
            Input(fNum, stUserData.dDaihyouBeta)                ' 代表β値
            Input(fNum, stUserData.dAlpha)                      ' α値
            Input(fNum, stUserData.dBeta)                       ' β値
            'V2.1.0.0③↓
            If VersionNumber >= 5 Then
                Input(fNum, stUserData.iTempSensorInfNoDaihyou)     ' 代表温度センサー情報一元管理選択番号
                If stUserData.iTempSensorInfNoDaihyou > 0 Then
                    Dim dDummy As Double
                    If Not TemperatureTableDataGet(stUserData.iTempSensorInfNoDaihyou, dDummy, stUserData.dDaihyouAlpha, stUserData.dDaihyouBeta, True) Then
                        Call Z_PRINT("温度センサー情報取得エラーNo=[" & stUserData.iTempSensorInfNoDaihyou.ToString("0") & "]")
                        GoTo rData_load_END
                    End If
                End If
                Input(fNum, stUserData.iTempSensorInfNoStd)        ' STD温度センサー情報一元管理選択番号
                If stUserData.iTempSensorInfNoStd > 0 Then
                    If Not TemperatureTableDataGet(stUserData.iTempSensorInfNoStd, stUserData.dTemperatura0, stUserData.dAlpha, stUserData.dBeta, True) Then
                        Call Z_PRINT("温度センサー情報取得エラーNo=[" & stUserData.iTempSensorInfNoStd.ToString("0") & "]")
                        GoTo rData_load_END
                    End If
                End If
            Else
                stUserData.iTempSensorInfNoDaihyou = 0              ' 未使用
                stUserData.iTempSensorInfNoStd = 0                  ' 未使用
            End If
            'V2.1.0.0③↑
            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル
            Input(fNum, stUserData.dResTempCoff)                ' 抵抗温度係数")を変更するカットNo.
        End If
        'V2.0.0.0⑪↑

        ' 抵抗 
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル
        Input(fNum, stUserData.dFinalLimitLow)              ' ファイナルリミット Low[%]
        Input(fNum, stUserData.dFinalLimitHigh)             ' ファイナルリミット Hight[%]
        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル
        Input(fNum, stUserData.dRelativeLow)                ' 相対値リミット Low[%]
        Input(fNum, stUserData.dRelativeHigh)               ' 相対値リミット Hight[%]

        Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル
        'V2.0.0.0⑩↓データエラー (ユーザデータ) !!対策
        Dim RCnt As Short
        If VersionNumber <= 2 Then
            RCnt = GetRCountExceptMeasureOldVersion()
        Else
            RCnt = GetRCountExceptMeasure()
        End If
        'V2.0.0.0⑩↑
        ' MAX_RES_USER
        For i = 1 To MAX_RES_USER
            'V2.0.0.0⑩            If i <= GetRCountExceptMeasure() Then
            If i <= RCnt Then                                       'V2.0.0.0⑩
                Input(fNum, stUserData.iResUnit(i))                 ' 抵抗レンジ 1:Ω, 2:KΩ
                Input(fNum, stUserData.dNomCalcCoff(i))             ' 補正値（ノミナル値算出係数）
                If VersionNumber <= 2 Then                                                      'V2.0.0.0⑫
                    stUserData.dNomCalcCoff(i) = (stUserData.dNomCalcCoff(i) - 1.0) * 1000000.0 'V2.0.0.0⑫ 補正値の項目をppm入力に変更
                End If                                                                          'V2.0.0.0⑫
                Input(fNum, stUserData.dTargetCoff(i))              ' 目標値算出係数
                'V2.1.0.0①↓
                If VersionNumber >= 5 Then
                    Input(fNum, stUserData.dTargetCoffJudge(i))     ' 判定用目標値算出係数
                Else
                    stUserData.dTargetCoffJudge(i) = stUserData.dTargetCoff(i)
                End If
                'V2.1.0.0①↑
                Input(fNum, stUserData.iChangeSpeed(i))             ' 測定速度を変更するカットNo.
            Else
                stUserData.iResUnit(i) = 1
                stUserData.dNomCalcCoff(i) = 0.0
                stUserData.dTargetCoff(i) = 0.0
                stUserData.iChangeSpeed(i) = 0.0
            End If
        Next i

        'V2.0.0.0②↓
        If VersionNumber >= 3 Then
            ' 電圧
            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル
            Input(fNum, stUserData.dRated)                      ' 定格
            Input(fNum, stUserData.dMagnification)              ' 定格電圧の倍率
            Input(fNum, stUserData.dResNumber)                  ' 抵抗個数
            Input(fNum, stUserData.dCurrentLimit)               ' 電流制限
            Input(fNum, stUserData.dAppliedSecond)              ' 印加秒数
            Input(fNum, stUserData.dVariation)                  ' 変化量
        Else
            stUserData.dRated = 0.0                             ' 定格
            stUserData.dMagnification = 1.0                     ' 定格電圧の倍率
            stUserData.dResNumber = 1                           ' 抵抗個数
            stUserData.dCurrentLimit = 0.01                     ' 電流制限
            stUserData.dAppliedSecond = 1.0                     ' 印加秒数
            stUserData.dVariation = 100                         ' 変化量
        End If
        'V2.0.0.0②↑

        'V2.0.0.1③↓
        If VersionNumber >= 4 Then              'V2.0.0.1③　トリミングNG信号を出力するＮＧの比率
            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル
            Input(fNum, stUserData.NgJudgeRate)
        Else
            stUserData.NgJudgeRate = 100.0
        End If
        'V2.0.0.1③↑

        'V2.2.0.034 ↓
        If VersionNumber >= 7 Then
            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル
            Input(fNum, stMultiBlock.gMultiBlock)               '複数抵抗値機能有効／無効
            Input(fNum, stMultiBlock.gStepRpt)                  '複数並び方向：0:列、1:行
            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル

            For l As Integer = 0 To 4
                Input(fNum, stMultiBlock.BLOCK_DATA(l).gBlockCnt)   ' ブロック1数
            Next
            Input(fNum, a) : Input(fNum, a) : Input(fNum, a)    ' 3行空読み　項目タイトル

            For l As Integer = 0 To 4
                For m As Integer = 0 To 4
                    Input(fNum, stMultiBlock.BLOCK_DATA(l).dblNominal(m))    ' 抵抗値
                    Input(fNum, stMultiBlock.BLOCK_DATA(l).iUnit(m))         ' 単位
                    Input(fNum, stMultiBlock.BLOCK_DATA(l).dblCorr(m))       ' 補正値
                Next m
            Next l

        Else

            'V2.2.0.033 ↓ 
            stMultiBlock.gMultiBlock = 0
            stMultiBlock.Initialize()
            For p As Integer = 0 To 5
                stMultiBlock.BLOCK_DATA(p).DataNo = p + 1           ' DataNo
                stMultiBlock.BLOCK_DATA(p).Initialize()
                stMultiBlock.BLOCK_DATA(p).gBlockCnt = 0            ' ブロック数
            Next
            'V2.2.0.033 ↑

        End If
        'V2.2.0.034 ↑

        Input(fNum, a)                                      ' 1行空読み
        If (a <> "/") Then
STP_USERERR:
            strMSG = "データエラー (ユーザデータ) !!" & vbCrLf
            Call Z_PRINT(strMSG)
            GoTo rData_load_END
        End If
        '--------------------------------------------------------------------------
        If VersionNumber < 5 Then                               'V2.1.0.0①初期値設定の時
            Call UserSub.CutVariationDataInitialize()           'V2.1.0.0① カット毎の抵抗値変化量判定機能初期データの設定
        End If                                                  'V2.1.0.0①
        rData_load = 0                                      ' Return値 = 正常

rData_load_END:
        FileClose(fNum)                                     ' データファイル クローズ

        'V2.2.0.0⑮ ↓
        If rData_load = 0 Then
            'V2.2.2.0②             If VersionNumber >= 6 Then      ' ﾌﾟﾛｰﾌﾞNo 
            If VersionNumber >= 7 Then      ' ﾌﾟﾛｰﾌﾞNo      'V2.2.2.0② 
                If stPLT.ProbNo <> 0 Then       ' 指定のプローブデータを読込み設定する 
                    ConvProbeData(stPLT.ProbNo)
                End If
            End If

            If giLoaderType <> 0 Then   'クランプ吸着動作設定
                ObjSys.setClampVaccumConfig(stUserData.intClampVacume - 1)
            End If

            ''V2.2.0.0⑯↓ 
            'stMultiBlock.gMultiBlock = 0
            'stMultiBlock.Initialize()
            'For i = 0 To 5
            '    stMultiBlock.BLOCK_DATA(i).DataNo = i + 1           ' DataNo
            '    stMultiBlock.BLOCK_DATA(i).Initialize()
            '    stMultiBlock.BLOCK_DATA(i).gBlockCnt = 0            ' ブロック数
            'Next
            ''V2.2.0.0⑯↑

        End If
        'V2.2.0.0⑮ ↑

rData_load_FileNone:

    End Function
#End Region

#Region "データファイルセーブ"
    '''=========================================================================
    '''<summary>データファイルセーブ</summary>
    '''<returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Public Function rData_save(ByRef DataPath As String) As Short

        On Error GoTo STP_END
        Dim a As String                                 ' データ出力域
        Dim b As String                                 ' データ出力域
        Dim c(17) As String                             ' データ出力域
        Dim d As String                                 ' データ出力域             'V2.2.0.0⑮ 
        Dim i As Short                                  ' Counter
        Dim j As Short                                  ' Counter
        Dim k As Short                                  ' Counter

        rData_save = 1                                  ' 異常終了

        '--------------------------------------------------------------------------
        '   システムデータを設定する
        '--------------------------------------------------------------------------
        fNum = FreeFile() ' Data File No.
        FileOpen(fNum, DataPath, OpenMode.Output)           ' データファイル Open

        a = "'======================================="
        PrintLine(fNum, a)
        'V2.0.0.1③        a = "' ●システムデータ [ DATA VERSION 3 ] "
        'V2.1.0.0   a = "' ●システムデータ [ DATA VERSION 4 ] "
        'V2.2.0.0② a = "' ●システムデータ [ DATA VERSION 5 ] "        'V2.1.0.0
        'V2.2.0.034 a = "' ●システムデータ [ DATA VERSION 6 ] "        'V2.2.0.0②
        'a = "' ●システムデータ [ DATA VERSION 7 ] "        'V2.2.0.034
        a = "' ●システムデータ [ DATA VERSION 8 ] "        'V2.2.1.7①
        PrintLine(fNum, a)
        a = "'======================================="
        PrintLine(fNum, a)

        ' 各種オフセットデータ等を設定する
        a = "'-------------------------------------"
        PrintLine(fNum, a)
        a = "' Z-Probe Off/On Offset (Off/On(mm)) ProbeNo"
        PrintLine(fNum, a)
        a = "'-------------------------------------"
        PrintLine(fNum, a)
        a = stPLT.Z_ZOFF.ToString("###0.0###")
        b = stPLT.Z_ZON.ToString("###0.0###")
        'V2.2.0.0⑮ ↓
        d = stPLT.ProbNo.ToString("0")
        PrintLine(fNum, a, b, d)                            ' Z-Probe Off/On Offset (Off/On(mm))
        'V2.2.0.0⑮ ↓

        '    a = stPLT.Z2_ZOFF.ToString("###0.0###")
        '    b = stPLT.Z2_ZON.ToString("###0.0###")
        '    Print #fNum, a, b                              ' Z2-Probe Off/On Offset (Off/On(mm))

        a = "'---------------------------------"
        PrintLine(fNum, a)
        a = "' ﾃｰﾌﾞﾙ ﾎﾟｼﾞｼｮﾝ ｵﾌｾｯﾄ (XY(mm))"
        PrintLine(fNum, a)
        a = "'---------------------------------"
        PrintLine(fNum, a)
        a = stPLT.z_xoff.ToString("###0.0###")
        b = stPLT.z_yoff.ToString("###0.0###")
        PrintLine(fNum, a, b)                               ' trim position offset x,y(mm)

        a = "'--------------------------------"
        PrintLine(fNum, a)
        a = "' ﾋﾞｰﾑ ﾎﾟｼﾞｼｮﾝ ｵﾌｾｯﾄ (xy(mm))"
        PrintLine(fNum, a)
        a = "'--------------------------------"
        PrintLine(fNum, a)
        a = stPLT.BPOX.ToString("###0.0###")
        b = stPLT.BPOY.ToString("###0.0###")
        PrintLine(fNum, a, b)                               ' BP offset x,y(mm)

        a = "'--------------------------------"
        PrintLine(fNum, a)
        a = "' ｱｼﾞｬｽﾄ ﾎﾟｼﾞｼｮﾝ (xy(mm))"
        PrintLine(fNum, a)
        a = "'--------------------------------"
        PrintLine(fNum, a)
        a = stPLT.ADJX.ToString("###0.0###")
        b = stPLT.ADJY.ToString("###0.0###")
        PrintLine(fNum, a, b)                               ' BP offset x,y(mm)

        a = "'--------------------------------"
        PrintLine(fNum, a)
        a = "' ブロック数 (XY)"
        PrintLine(fNum, a)
        a = "'--------------------------------"
        PrintLine(fNum, a)
        a = stPLT.BNX.ToString("0")
        b = stPLT.BNY.ToString("0")
        PrintLine(fNum, a, b)                               ' ブロック数x,y

        a = "'---------------------"
        PrintLine(fNum, a)
        a = "' ﾌﾞﾛｯｸｻｲｽﾞ(XY(mm))"
        PrintLine(fNum, a)
        a = "'---------------------"
        PrintLine(fNum, a)
        a = stPLT.zsx.ToString("0.0000")
        b = stPLT.zsy.ToString("0.0000")
        PrintLine(fNum, a, b)                               ' ブロックサイズ X,Y(mm)

        a = "'-------------------------------"
        PrintLine(fNum, a)
        a = "' プローブリトライ(1:有 0:無)"
        PrintLine(fNum, a)
        a = "'-------------------------------"
        PrintLine(fNum, a)
        a = stPLT.PrbRetry.ToString("#0")
        PrintLine(fNum, a)                                  ' プローブリトライ(1:有 0:無)

        a = "'-----------------------------------"
        PrintLine(fNum, a)
        a = "' ﾚｰｻﾞ調整Qﾚｰﾄ(0.1KHz)　指定ﾊﾟﾜｰ(W)"
        PrintLine(fNum, a)
        a = "'-----------------------------------"
        PrintLine(fNum, a)
        c(0) = stLASER.intQR.ToString("#0")
        c(1) = stLASER.dblspecPower.ToString("#0.00")
        PrintLine(fNum, c(0), c(1))                         ' Qレート (x100Hz)(0.1KHz), 指定ﾊﾟﾜｰ(W)

        a = "'------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "' ﾄﾘﾐﾝｸﾞ用/測定用ﾎﾟｰｽﾞ時間ms(1:最初の抵抗用 2:偶数抵抗用 3:奇数抵抗用)"
        PrintLine(fNum, a)
        a = "'------------------------------------------------------------------------"
        PrintLine(fNum, a)
        c(1) = glWTimeT(1).ToString("0")
        c(2) = glWTimeT(2).ToString("0")
        c(3) = glWTimeT(3).ToString("0")
        PrintLine(fNum, c(1), c(2), c(3))                   ' ﾄﾘﾐﾝｸﾞ用ポーズ時間ms1～3(1:2000ms, 2:   0ms, 3:   0ms)
        c(1) = glWTimeM(1).ToString("0")
        c(2) = glWTimeM(2).ToString("0")
        c(3) = glWTimeM(3).ToString("0")
        PrintLine(fNum, c(1), c(2), c(3))                   ' 測定用ポーズ時間ms1～3  (1:2000ms, 2:   0ms, 3:   0ms)
        ' ###1040 ADD START
        'V2.1.0.0②        a = "'------------------------------------------------------------------------"
        'V2.1.0.0②        PrintLine(fNum, a)
        'V2.1.0.0②        a = "' アッテネータ減衰率（保存、減衰率、回転量、固定アッテネータ(0:OFF 1:ON)）"
        'V2.1.0.0②        PrintLine(fNum, a)
        'V2.1.0.0②        a = "'------------------------------------------------------------------------"
        'V2.1.0.0②↓
        a = "'----------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "' アッテネータ減衰率（保存、減衰率、回転量、固定アッテネータ(0:OFF 1:ON)）アッテネータNo.（0:指定無）"
        PrintLine(fNum, a)
        a = "'----------------------------------------------------------------------------------------------------"
        'V2.1.0.0②↑
        PrintLine(fNum, a)
        c(1) = stLASER.iTrimAtt.ToString("0")
        c(2) = stLASER.dblRotPar.ToString("0.00")
        c(3) = stLASER.dblRotAtt.ToString("0")
        c(4) = stLASER.iFixAtt.ToString("0")
        c(5) = stLASER.iAttNo.ToString("0")             'V2.1.0.0②
        PrintLine(fNum, c(1), c(2), c(3), c(4), c(5))   'V2.1.0.0②　 c(5)追加

        a = "'------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "' ティーチング・ブロック(XY)"
        PrintLine(fNum, a)
        a = "'------------------------------------------------------------------------"
        PrintLine(fNum, a)
        c(1) = stPLT.TeachBlockX.ToString("0")
        c(2) = stPLT.TeachBlockY.ToString("0")
        PrintLine(fNum, c(1), c(2))

        a = "'------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "' Ｙ軸スピード"
        PrintLine(fNum, a)
        a = "'------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = stPLT.StageSpeedY.ToString("0")
        PrintLine(fNum, a)

        ' ###1040 ADD END
        'V1.2.0.0①↓
        a = "'---------------------"
        PrintLine(fNum, a)
        a = "' チップサイズ(XY(mm))"
        PrintLine(fNum, a)
        a = "'---------------------"
        PrintLine(fNum, a)
        a = stPLT.dblChipSizeXDir.ToString("0.0000")
        b = stPLT.dblChipSizeYDir.ToString("0.0000")
        PrintLine(fNum, a, b)                               ' チップサイズ X,Y(mm)
        'V1.2.0.0①↑
        'V2.0.0.0①↓
        a = "'-------------------------------------"
        PrintLine(fNum, a)
        a = "' ステップオフセット量(XY)"
        PrintLine(fNum, a)
        a = "'-------------------------------------"
        PrintLine(fNum, a)
        c(1) = stPLT.dblStepOffsetXDir.ToString("0.0000")
        c(2) = stPLT.dblStepOffsetYDir.ToString("0.0000")
        PrintLine(fNum, c(1), c(2))
        'V2.0.0.0①↑
        'V2.2.0.0②↓
        a = "'-------------------------------------"
        PrintLine(fNum, a)
        a = "' デジタルカメラ倍率  (0.5～2.0)"
        PrintLine(fNum, a)
        a = "'-------------------------------------"
        PrintLine(fNum, a)
        a = stPLT.dblStdMagnification.ToString("0.0")
        PrintLine(fNum, a)                               'デジタルカメラ倍率
        'V2.2.0.0②↑
        a = "/"
        PrintLine(fNum, a)
        a = "'"
        PrintLine(fNum, a)

        '--------------------------------------------------------------------------
        '   抵抗データを設定する
        '--------------------------------------------------------------------------
        a = "'======================================="
        PrintLine(fNum, a)
        a = "' ●抵抗データ"
        PrintLine(fNum, a)
        a = "'======================================="
        PrintLine(fNum, a)

        a = "'-------------------"
        PrintLine(fNum, a)
        a = "'抵抗数"
        PrintLine(fNum, a)
        a = "'-------------------"
        PrintLine(fNum, a)
        c(0) = stPLT.RCount.ToString("0")
        PrintLine(fNum, c(0))                               ' 抵抗数

        ' 抵抗数分、以下を設定する
        a = "'----------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "'抵抗No.    抵抗名 ｽﾛｰﾌﾟ(+:1 -:2 抵抗:4) ﾘﾚｰﾋﾞｯﾄ(Hex)   HP            LP            GP"
        PrintLine(fNum, a)
        a = "'----------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.RCount                           ' 抵抗数分繰り返す
            c(0) = i.ToString("0")                       ' 抵抗No.
            c(1) = """" & stREG(i).strRNO & """"            ' 抵抗名
            c(2) = stREG(i).intSLP.ToString("0")         ' ｽﾛｰﾌﾟ
            c(3) = """" & Hex(stREG(i).lngRel) & """"       ' リレービット(HEX)
            c(4) = stREG(i).intPRH.ToString("0")         ' HP
            c(5) = stREG(i).intPRL.ToString("0")         ' LP
            c(6) = stREG(i).intPRG.ToString("0")         ' GP
            PrintLine(fNum, c(0), c(1), c(2), c(3), c(4), c(5), c(6))
        Next i

        a = "'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "'抵抗No.      目標値        単位  判定(0:% 1:絶対値)  測定（IT/FT有無） ITLO         ITHI           FTLO         FTHI      カット数  測定(0:内部1:外部)  精度(0:高速1:高精度)再測定回数　再測定前待機時間"
        PrintLine(fNum, a)
        a = "'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.RCount                                   ' 抵抗数分繰り返す
            c(0) = i.ToString("0")                                  ' 抵抗No.
            ''V2.2.0.0⑫            c(1) = stREG(i).dblNOM.ToString("#0.0000#")             ' 目標値
            c(1) = stREG(i).dblNOM.ToString("#0.000000#")             ' 目標値     'V2.2.0.0⑫
            c(2) = """" & stREG(i).strTANI & """"                   ' 単位
            c(3) = stREG(i).intMode.ToString("0")                   ' 判定(0:% 1:絶対値)
            c(4) = stREG(i).intMeasMode.ToString("0")               ' 測定モード(0:なし, 1:ITのみ 2:FTのみ 3:IT,FT両方)
            c(5) = stREG(i).dblITL.ToString("#0.0000#")             ' 初期判定下限値
            c(6) = stREG(i).dblITH.ToString("#0.0000#")             ' 初期判定上限値
            c(7) = stREG(i).dblFTL.ToString("#0.0000#")             ' 終了判定下限値
            c(8) = stREG(i).dblFTH.ToString("#0.0000#")             ' 終了判定上限値
            c(9) = stREG(i).intTNN.ToString("0")                    ' ｶｯﾄ数
            c(10) = stREG(i).intMType.ToString("0")                 ' 測定
            c(11) = stREG(i).intTMM1.ToString("0")                  ' 精度(0:高速 1:高精度)
            c(12) = stREG(i).intReMeas.ToString("0")                ' 再測定回数（0:再測定無 1:再測定回数）
            c(13) = stREG(i).intReMeas_Time.ToString("0")           ' 再測定前ポーズ時間
            PrintLine(fNum, c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), c(11), c(12), c(13))
        Next i

        a = "'-----------------------------------------------------"
        PrintLine(fNum, a)
        a = "'抵抗No.    　IT測定回数　  FT測定回数　  サーキット"
        PrintLine(fNum, a)
        a = "'-----------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.RCount                                   ' 抵抗数分繰り返す
            c(0) = i.ToString("0")                                  ' 抵抗No.
            c(1) = stREG(i).intITReMeas.ToString("0")              ' イニシャル抵抗再測定回数(IT測定回数)
            c(2) = stREG(i).intFTReMeas.ToString("0")              ' ファイナル抵抗再測定回数(FT測定回数)
            c(3) = stREG(i).intCircuitNo.ToString("0")             ' サーキット番号
            PrintLine(fNum, c(0), c(1), c(2), c(3))
        Next i
        'V2.0.0.0②↓
        a = "'------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "'抵抗No. 　ＯＮ機器１    ＯＮ機器２    ＯＮ機器３  ＯＦＦ機器１  ＯＦＦ機器２  ＯＦＦ機器３"
        PrintLine(fNum, a)
        a = "'------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.RCount                                       ' 抵抗数分繰り返す
            c(0) = i.ToString("0")                                      ' 抵抗No.

            For j = 1 To EXTEQU
                c(j) = stREG(i).intOnExtEqu(j).ToString("0")            ' ＯＮ機器
            Next
            For j = 1 To EXTEQU
                c(j + EXTEQU) = stREG(i).intOffExtEqu(j).ToString("0")  ' ＯＦＦ機器
            Next
            PrintLine(fNum, c(0), c(1), c(2), c(3), c(4), c(5), c(6))
        Next i
        'V2.0.0.0②↑
        a = "/"
        PrintLine(fNum, a)
        a = "'"
        PrintLine(fNum, a)

        '--------------------------------------------------------------------------
        '   カットデータを設定する
        '--------------------------------------------------------------------------
        a = "'======================================="
        PrintLine(fNum, a)
        a = "' ●カットデータ"
        PrintLine(fNum, a)
        a = "'======================================="
        PrintLine(fNum, a)

        ' カット数分、以下を設定する
        a = "'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "'抵抗No.     ｶｯﾄNo.     ｶｯﾄ方法        ｶｯﾄ形状        本数             STX(mm)       STY(mm)      STX2(mm)      STY2(mm)       CUTOFF(%)   MD       Qﾚｰﾄ(.1K)      速度(mm/sec)      DL1(mm)       DL2(mm)"
        PrintLine(fNum, a)
        a = "'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.RCount                           ' 抵抗数分繰り返す
            For j = 1 To stREG(i).intTNN                    ' 抵抗のカット数分繰り返す
                c(0) = i.ToString("0")                   ' 抵抗NO.
                c(1) = j.ToString("0")                   ' カットNO.
                c(2) = stREG(i).STCUT(j).intCUT.ToString("0")    ' ｶｯﾄ方法
                c(3) = stREG(i).STCUT(j).intCTYP.ToString("0")   ' ｶｯﾄ形状
                c(4) = stREG(i).STCUT(j).intNum.ToString("0")    ' 本数
                c(5) = stREG(i).STCUT(j).dblSTX.ToString("##0.0000#") ' STX(mm)
                c(6) = stREG(i).STCUT(j).dblSTY.ToString("##0.0000#") ' STY(mm)
                c(7) = stREG(i).STCUT(j).dblSX2.ToString("##0.0000#") ' STX2(mm)
                c(8) = stREG(i).STCUT(j).dblSY2.ToString("##0.0000#") ' STY2(mm)
                c(9) = stREG(i).STCUT(j).dblCOF.ToString("##0.0000#") ' CUTOFF(%)
                c(10) = stREG(i).STCUT(j).intTMM.ToString("0")   ' MD
                c(11) = stREG(i).STCUT(j).intQF1.ToString("0")   ' Qﾚｰﾄ(.1K)
                c(12) = stREG(i).STCUT(j).dblV1.ToString("##0.0") ' 速度

                c(13) = stREG(i).STCUT(j).dblDL2.ToString("###0.0###") ' DL1(mm)
                c(14) = stREG(i).STCUT(j).dblDL3.ToString("###0.0###") ' DL2(mm)
                PrintLine(fNum, c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), c(11), c(12), c(13), c(14))
            Next j
        Next i

        a = "'------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "'抵抗No.     ｶｯﾄNo.       ANG1(°)      ANG2(°)        LTP(%)      測定(0:内部 1:外部) "
        PrintLine(fNum, a)
        a = "'------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.RCount                           ' 抵抗数分繰り返す
            For j = 1 To stREG(i).intTNN                    ' カット数分繰り返す
                c(0) = i.ToString("0")                                          ' 抵抗NO.
                c(1) = j.ToString("0")                                          ' カットNO.
                c(2) = stREG(i).STCUT(j).intANG.ToString("0")                   ' ANG1(°)
                c(3) = stREG(i).STCUT(j).intANG2.ToString("0")                  ' ANG2(°)
                c(4) = stREG(i).STCUT(j).dblLTP.ToString("###0.0###")           ' LTP(%)
                c(5) = stREG(i).STCUT(j).intMType.ToString("0")                 ' 測定(0:内部 1:外部)
                PrintLine(fNum, c(0), c(1), c(2), c(3), c(4), c(5))
            Next j
        Next i

        a = "'---------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "'抵抗No.     ｶｯﾄNo.       IX回数1       IX回数2       IX回数3       IX回数4       IX回数5      ﾋﾟｯﾁ1(mm)     ﾋﾟｯﾁ2(mm)     ﾋﾟｯﾁ3(mm)     ﾋﾟｯﾁ4(mm)     ﾋﾟｯﾁ5(mm) "
        PrintLine(fNum, a)
        a = "'---------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.RCount                           ' 抵抗数分繰り返す
            For j = 1 To stREG(i).intTNN                    ' カット数分繰り返す
                c(0) = i.ToString("0")                                                  ' 抵抗NO.
                c(1) = j.ToString("0")                                                  ' カットNO.
                c(2) = stREG(i).STCUT(j).intIXN(1).ToString("##0")            ' IX回数1
                c(3) = stREG(i).STCUT(j).intIXN(2).ToString("##0")            ' IX回数2
                c(4) = stREG(i).STCUT(j).intIXN(3).ToString("##0")            ' IX回数3
                c(5) = stREG(i).STCUT(j).intIXN(4).ToString("##0")            ' IX回数4
                c(6) = stREG(i).STCUT(j).intIXN(5).ToString("##0")            ' IX回数5
                c(7) = stREG(i).STCUT(j).dblDL1(1).ToString("###0.00##")     ' ﾋﾟｯﾁ1(mm)
                c(8) = stREG(i).STCUT(j).dblDL1(2).ToString("###0.00##")     ' ﾋﾟｯﾁ2(mm)
                c(9) = stREG(i).STCUT(j).dblDL1(3).ToString("###0.00##")     ' ﾋﾟｯﾁ3(mm)
                c(10) = stREG(i).STCUT(j).dblDL1(4).ToString("###0.00##")    ' ﾋﾟｯﾁ4(mm)
                c(11) = stREG(i).STCUT(j).dblDL1(5).ToString("###0.00##")    ' ﾋﾟｯﾁ5(mm)
                PrintLine(fNum, c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), c(11))
            Next j
        Next i

        a = "'---------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "'抵抗No.     ｶｯﾄNo.       PAU1(ms)      PAU2(ms)      PAU3(ms)      PAU4(ms)      PAU5(ms)       誤差1(%)      誤差2(%)      誤差3(%)      誤差4(%)      誤差5(%)"
        PrintLine(fNum, a)
        a = "'---------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.RCount                           ' 抵抗数分繰り返す
            For j = 1 To stREG(i).intTNN                    ' カット数分繰り返す
                c(0) = i.ToString("0")                                          ' 抵抗NO.
                c(1) = j.ToString("0")                                          ' カットNO.
                c(2) = stREG(i).STCUT(j).lngPAU(1).ToString("0")                ' ピッチ間ポーズ時間1
                c(3) = stREG(i).STCUT(j).lngPAU(2).ToString("0")                ' ピッチ間ポーズ時間2
                c(4) = stREG(i).STCUT(j).lngPAU(3).ToString("0")                ' ピッチ間ポーズ時間3
                c(5) = stREG(i).STCUT(j).lngPAU(4).ToString("0")                ' ピッチ間ポーズ時間4
                c(6) = stREG(i).STCUT(j).lngPAU(5).ToString("0")                ' ピッチ間ポーズ時間5
                c(7) = stREG(i).STCUT(j).dblDEV(1).ToString("0.0000")           ' 誤差1(%)
                c(8) = stREG(i).STCUT(j).dblDEV(2).ToString("0.0000")           ' 誤差2(%)
                c(9) = stREG(i).STCUT(j).dblDEV(3).ToString("0.0000")           ' 誤差3(%)
                c(10) = stREG(i).STCUT(j).dblDEV(4).ToString("0.0000")          ' 誤差4(%)
                c(11) = stREG(i).STCUT(j).dblDEV(5).ToString("0.0000")          ' 誤差5(%)
                PrintLine(fNum, c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), c(11))
            Next j
        Next i

        a = "'---------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "'抵抗No.     ｶｯﾄNo.       測定機器1      測定機器2      測定機器3      測定機器4      測定機器5       測定モード1      測定モード2      測定モード3      測定モード4      測定モード5"
        PrintLine(fNum, a)
        a = "'---------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.RCount                           ' 抵抗数分繰り返す
            For j = 1 To stREG(i).intTNN                    ' カット数分繰り返す
                c(0) = i.ToString("0")                                          ' 抵抗NO.
                c(1) = j.ToString("0")                                          ' カットNO.
                c(2) = stREG(i).STCUT(j).intIXMType(1).ToString("0")                ' 測定機器1
                c(3) = stREG(i).STCUT(j).intIXMType(2).ToString("0")                ' 測定機器2
                c(4) = stREG(i).STCUT(j).intIXMType(3).ToString("0")                ' 測定機器3
                c(5) = stREG(i).STCUT(j).intIXMType(4).ToString("0")                ' 測定機器4
                c(6) = stREG(i).STCUT(j).intIXMType(5).ToString("0")                ' 測定機器5


                c(7) = stREG(i).STCUT(j).intIXTMM(1).ToString("0")                 ' 測定モード1
                c(8) = stREG(i).STCUT(j).intIXTMM(2).ToString("0")                 ' 測定モード2
                c(9) = stREG(i).STCUT(j).intIXTMM(3).ToString("0")                 ' 測定モード3
                c(10) = stREG(i).STCUT(j).intIXTMM(4).ToString("0")                ' 測定モード4
                c(11) = stREG(i).STCUT(j).intIXTMM(5).ToString("0")                ' 測定モード5

                PrintLine(fNum, c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10), c(11))
            Next j
        Next i


        a = "'------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "'抵抗No.     ｶｯﾄNo.       条件No.1      条件No.2      条件No.3      条件No.4"
        PrintLine(fNum, a)
        a = "'------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.RCount                                       ' 抵抗数分繰り返す
            For j = 1 To stREG(i).intTNN                                ' カット数分繰り返す
                c(0) = i.ToString("0")                                  ' 抵抗NO.
                c(1) = j.ToString("0")                                  ' カットNO.
                c(2) = stREG(i).STCUT(j).intCND(1).ToString("0")        ' 加工条件No.1～4
                c(3) = stREG(i).STCUT(j).intCND(2).ToString("0")
                c(4) = stREG(i).STCUT(j).intCND(3).ToString("0")
                c(5) = stREG(i).STCUT(j).intCND(4).ToString("0")
                PrintLine(fNum, c(0), c(1), c(2), c(3), c(4), c(5))
            Next j
        Next i

        'V1.0.4.3③ ADD ↓
        a = "'- L CUT PARAMETER ------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "'抵抗No.     ｶｯﾄNo.       LDL1(mm)       LDL2(mm)      LDL3(mm)      LDL4(mm)      LDL5(mm)      LDL6(mm)      LDL7(mm)  "
        PrintLine(fNum, a)
        a = "'------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.RCount                                                   ' 抵抗数分繰り返す
            For j = 1 To stREG(i).intTNN                                            ' カット数分繰り返す
                a = i.ToString("0")                                                 ' 抵抗NO.
                b = j.ToString("0")                                                 ' カットNO.
                For k = 1 To MAX_LCUT
                    c(k) = stREG(i).STCUT(j).dCutLen(k).ToString("###0.0###")       ' カット長
                Next k
                PrintLine(fNum, a, b, c(1), c(2), c(3), c(4), c(5), c(6), c(7))
            Next j
        Next i

        a = "'------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "'抵抗No.     ｶｯﾄNo.       LQﾚｰﾄ1         LQﾚｰﾄ2        LQﾚｰﾄ3        LQﾚｰﾄ4        LQﾚｰﾄ5        LQﾚｰﾄ6        LQﾚｰﾄ7 "
        PrintLine(fNum, a)
        a = "'------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.RCount                                               ' 抵抗数分繰り返す
            For j = 1 To stREG(i).intTNN                                        ' カット数分繰り返す
                a = i.ToString("0")                                             ' 抵抗NO.
                b = j.ToString("0")                                             ' カットNO.
                For k = 1 To MAX_LCUT
                    c(k) = stREG(i).STCUT(j).dQRate(k).ToString("0")            ' Ｑレート
                Next k
                PrintLine(fNum, a, b, c(1), c(2), c(3), c(4), c(5), c(6), c(7))
            Next j
        Next i

        a = "'------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "'抵抗No.     ｶｯﾄNo.       L速度1         L速度2        L速度3        L速度4        L速度5        L速度6        L速度7 "
        PrintLine(fNum, a)
        a = "'------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.RCount                                               ' 抵抗数分繰り返す
            For j = 1 To stREG(i).intTNN                                        ' カット数分繰り返す
                a = i.ToString("0")                                             ' 抵抗NO.
                b = j.ToString("0")                                             ' カットNO.
                For k = 1 To MAX_LCUT
                    c(k) = stREG(i).STCUT(j).dSpeed(k).ToString("##0.0")        ' 速度
                Next k
                PrintLine(fNum, a, b, c(1), c(2), c(3), c(4), c(5), c(6), c(7))
            Next j
        Next i

        a = "'------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "'抵抗No.     ｶｯﾄNo.       LANG1          LANG2         LANG3         LANG4         LANG5         LANG6         LANG7 "
        PrintLine(fNum, a)
        a = "'------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.RCount                                               ' 抵抗数分繰り返す
            For j = 1 To stREG(i).intTNN                                        ' カット数分繰り返す
                a = i.ToString("0")                                             ' 抵抗NO.
                b = j.ToString("0")                                             ' カットNO.
                For k = 1 To MAX_LCUT
                    c(k) = stREG(i).STCUT(j).dAngle(k).ToString("0")            ' 角度
                Next k
                PrintLine(fNum, a, b, c(1), c(2), c(3), c(4), c(5), c(6), c(7))
            Next j
        Next i

        a = "'------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "'抵抗No.     ｶｯﾄNo.        LTP1          LTP2          LTP3          LTP4          LTP5          LTP6"
        PrintLine(fNum, a)
        a = "'------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.RCount                                                   ' 抵抗数分繰り返す
            For j = 1 To stREG(i).intTNN                                            ' カット数分繰り返す
                a = i.ToString("0")                                                 ' 抵抗NO.
                b = j.ToString("0")                                                 ' カットNO.
                For k = 1 To MAX_LCUT - 1                                           ' ターンポイントは、カット数より１つ少ない
                    c(k) = stREG(i).STCUT(j).dTurnPoint(k).ToString("##0.0000#")    ' ターンポイント
                Next k
                PrintLine(fNum, a, b, c(1), c(2), c(3), c(4), c(5), c(6))
            Next j
        Next i

        a = "'------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        'a = "'抵抗No.     ｶｯﾄNo.       リトレースQﾚｰﾄ(.1K)      リトレース速度(mm/sec)      文字" 'V2.2.1.7①
        a = "'抵抗No.     ｶｯﾄNo.       リトレースQﾚｰﾄ(.1K)      リトレース速度(mm/sec)      文字      印字固定部      開始番号      重複回数" 'V2.2.1.7①
        PrintLine(fNum, a)
        a = "'------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.RCount                                           ' 抵抗数分繰り返す
            For j = 1 To stREG(i).intTNN                                    ' カット数分繰り返す
                a = i.ToString("0")                                         ' 抵抗NO.
                b = j.ToString("0")                                         ' カットNO.
                c(0) = stREG(i).STCUT(j).intQF2.ToString("0")               ' Qﾚｰﾄ(.1K)
                c(1) = stREG(i).STCUT(j).dblV2.ToString("##0.0")            ' 速度
                c(2) = """" & stREG(i).STCUT(j).cFormat & """"
                c(3) = """" & stREG(i).STCUT(j).cMarkFix & """"              'V2.2.1.7①
                c(4) = """" & stREG(i).STCUT(j).cMarkStartNum & """"         'V2.2.1.7①
                c(5) = stREG(i).STCUT(j).intMarkRepeatCnt.ToString("0")      'V2.2.1.7①

                'V2.2.1.7① ↓
                'PrintLine(fNum, a, b, c(0), c(1), c(2))
                PrintLine(fNum, a, b, c(0), c(1), c(2), c(3), c(4), c(5))
                'V2.2.1.7① ↑
            Next j
        Next i
        'V1.0.4.3③ ADD ↑
        'V2.0.0.0⑦ ADD ↓
        a = "'- RETRACE CUT PARAMETER -----------------"
        PrintLine(fNum, a)
        a = "'抵抗No.     ｶｯﾄNo.         リトレース本数"
        PrintLine(fNum, a)
        a = "'-----------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.RCount                                                       ' 抵抗数分繰り返す
            For j = 1 To stREG(i).intTNN                                                ' カット数分繰り返す
                a = i.ToString("0")                                                     ' 抵抗NO.
                b = j.ToString("0")                                                     ' カットNO.
                c(1) = stREG(i).STCUT(j).intRetraceCnt.ToString("0")    ' リトレースのオフセットＸ
                PrintLine(fNum, a, b, c(1))
            Next j
        Next i

        a = "'---------------------------------------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "'抵抗No.     ｶｯﾄNo.         OFFX1         OFFX2         OFFX3         OFFX4         OFFX5         OFFX6         OFFX7         OFFX8         OFFX9         OFFX10"
        PrintLine(fNum, a)
        a = "'---------------------------------------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.RCount                                                       ' 抵抗数分繰り返す
            For j = 1 To stREG(i).intTNN                                                ' カット数分繰り返す
                a = i.ToString("0")                                                     ' 抵抗NO.
                b = j.ToString("0")                                                     ' カットNO.
                For k = 1 To MAX_RETRACECUT
                    c(k) = stREG(i).STCUT(j).dblRetraceOffX(k).ToString("##0.0000#")    ' リトレースのオフセットＸ
                Next k
                PrintLine(fNum, a, b, c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10))
            Next j
        Next i

        a = "'---------------------------------------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "'抵抗No.     ｶｯﾄNo.         OFFY1         OFFY2         OFFY3         OFFY4         OFFY5         OFFY6         OFFY7         OFFY8         OFFY9         OFFY10"
        PrintLine(fNum, a)
        a = "'---------------------------------------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.RCount                                                   ' 抵抗数分繰り返す
            For j = 1 To stREG(i).intTNN                                            ' カット数分繰り返す
                a = i.ToString("0")                                                 ' 抵抗NO.
                b = j.ToString("0")                                                 ' カットNO.
                For k = 1 To MAX_RETRACECUT
                    c(k) = stREG(i).STCUT(j).dblRetraceOffY(k).ToString("##0.0000#")    ' リトレースのオフセットＹ
                Next k
                PrintLine(fNum, a, b, c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10))
            Next j
        Next i

        a = "'----------------------------------------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "'抵抗No.     ｶｯﾄNo.         QRATE1        QRATE2        QRATE3        QRATE4        QRATE5        QRATE6        QRATE7        QRATE8        QRATE9        QRATE10"
        PrintLine(fNum, a)
        a = "'----------------------------------------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.RCount                                                   ' 抵抗数分繰り返す
            For j = 1 To stREG(i).intTNN                                            ' カット数分繰り返す
                a = i.ToString("0")                                                 ' 抵抗NO.
                b = j.ToString("0")                                                 ' カットNO.
                For k = 1 To MAX_RETRACECUT
                    c(k) = stREG(i).STCUT(j).dblRetraceQrate(k).ToString("0")     ' ストレートカット・リトレースのQレート(0.1KHz)に使用
                Next k
                PrintLine(fNum, a, b, c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10))
            Next j
        Next i

        a = "'---------------------------------------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "'抵抗No.     ｶｯﾄNo.         V1            V2            V3            V4            V5            V6            V7            V8            V9            V10   "
        PrintLine(fNum, a)
        a = "'---------------------------------------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.RCount                                                   ' 抵抗数分繰り返す
            For j = 1 To stREG(i).intTNN                                            ' カット数分繰り返す
                a = i.ToString("0")                                                 ' 抵抗NO.
                b = j.ToString("0")                                                 ' カットNO.
                For k = 1 To MAX_RETRACECUT
                    c(k) = stREG(i).STCUT(j).dblRetraceSpeed(k).ToString("##0.0")     ' ストレートカット・リトレースのトリム速度(mm/s)に使用
                Next k
                PrintLine(fNum, a, b, c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9), c(10))
            Next j
        Next i
        'V2.0.0.0⑦ ADD ↑

        'V2.1.0.0①↓
        a = "'- CUT VARIATIONREPEAT PARAMETER ------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "'抵抗No.     ｶｯﾄNo.         リピート有無　判定有無　    上昇率　      下限値　      上限値   "
        PrintLine(fNum, a)
        a = "'--------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.RCount                                                   ' 抵抗数分繰り返す
            For j = 1 To stREG(i).intTNN                                            ' カット数分繰り返す
                a = i.ToString("0")                                                 ' 抵抗NO.
                b = j.ToString("0")                                                 ' カットNO.

                c(1) = stREG(i).STCUT(j).iVariationRepeat.ToString("0") ' リピート有無
                c(2) = stREG(i).STCUT(j).iVariation.ToString("0")       ' 判定有無
                c(3) = stREG(i).STCUT(j).dRateOfUp.ToString("0.0000")        ' 上昇率
                c(4) = stREG(i).STCUT(j).dVariationLow.ToString("0.0000")    ' 下限値
                c(5) = stREG(i).STCUT(j).dVariationHi.ToString("0.0000")     ' 上限値
                PrintLine(fNum, a, b, c(1), c(2), c(3), c(4), c(5))
            Next j
        Next i
        'V2.1.0.0①↑

        'V2.2.0.0②↓　Uカット用パラメータの追加
        a = "'U CUT PARAMETER ------------------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "'抵抗No.     ｶｯﾄNo.         L1            L2          Qレート1       速度1         角度       ターンポイント  ターン方向       R1            R2   "
        PrintLine(fNum, a)
        a = "'----------------------------------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.RCount                                                   ' 抵抗数分繰り返す
            For j = 1 To stREG(i).intTNN                                            ' カット数分繰り返す
                a = i.ToString("0")                                                 ' 抵抗NO.
                b = j.ToString("0")                                                 ' カットNO.
                c(1) = stREG(i).STCUT(j).dUCutL1.ToString("###0.0###")      ' Uカット時L1カット長
                c(2) = stREG(i).STCUT(j).dUCutL2.ToString("###0.0###")      ' Uカット時L2カット長
                c(3) = stREG(i).STCUT(j).intUCutQF1.ToString("0")           ' Ｑレート1
                c(4) = stREG(i).STCUT(j).dblUCutV1.ToString("##0.0")        ' 速度1
                c(5) = stREG(i).STCUT(j).intUCutANG.ToString("0")           ' 角度
                c(6) = stREG(i).STCUT(j).dblUCutTurnP.ToString("###0.0###") ' LTP(%)
                c(7) = stREG(i).STCUT(j).intUCutTurnDir.ToString("0")       ' ターン方向
                c(8) = stREG(i).STCUT(j).dblUCutR1.ToString("###0.0###")   ' Uカット時半径R1 
                c(9) = stREG(i).STCUT(j).dblUCutR2.ToString("###0.0###")   ' Uカット時半径R2 
                PrintLine(fNum, a, b, c(1), c(2), c(3), c(4), c(5), c(6), c(7), c(8), c(9))
            Next j
        Next i
        'V2.2.0.0②↑

        a = "/"
        PrintLine(fNum, a)
        a = "'"
        PrintLine(fNum, a)

        '--------------------------------------------------------------------------
        '   パターン登録データ(カット位置補正用)を設定する
        '--------------------------------------------------------------------------
        a = "'========================================="
        PrintLine(fNum, a)
        a = "' ●パターン登録データ(カット位置補正用)"
        PrintLine(fNum, a)
        a = "'========================================="
        PrintLine(fNum, a)

        a = "'--------------------------------------"
        PrintLine(fNum, a)
        a = "'登録数"
        PrintLine(fNum, a)
        a = "'--------------------------------------"
        PrintLine(fNum, a)
        a = stPLT.RCount.ToString("0")
        PrintLine(fNum, a)                                           ' ﾊﾟﾀｰﾝ登録数（パターン登録数＝抵抗数の為）

        ' ﾊﾟﾀｰﾝ登録数分、以下を設定する
        a = "'---------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "'ｸﾞﾙｰﾌﾟ番号 ﾊﾟﾀｰﾝ番号 登録位置X(mm) 登録位置Y(mm) ﾊﾟﾀｰﾝ認識(0:無 1:有 2:手動)"
        PrintLine(fNum, a)
        a = "'---------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.RCount                         ' ﾊﾟﾀｰﾝ登録数分繰り返す
            c(0) = stPTN(i).intGRP.ToString("0")              ' GRP NO.
            c(1) = stPTN(i).intPTN.ToString("0")              ' PTN NO.
            c(2) = stPTN(i).dblPosX.ToString("###0.000#")     ' POS_X
            c(3) = stPTN(i).dblPosY.ToString("###0.000#")     ' POS_Y
            b = stPTN(i).PtnFlg.ToString("0")                           ' ﾊﾟﾀｰﾝ認識(1:有 0:無 2:手動)
            PrintLine(fNum, c(0), c(1), c(2), c(3), b)
        Next i

        a = "/"
        PrintLine(fNum, a)
        a = "'"
        PrintLine(fNum, a)

        '--------------------------------------------------------------------------
        '   パターン登録データ(ＸＹθ補正用)を設定する
        '--------------------------------------------------------------------------
        a = "'========================================="
        PrintLine(fNum, a)
        a = "' ●パターン登録データ(ＸＹθ補正用)"
        PrintLine(fNum, a)
        a = "'========================================="
        PrintLine(fNum, a)

        a = "'--------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "'補正モード(0:自動/1:手動)　補正方法(0:補正なし 1:補正あり)"
        PrintLine(fNum, a)
        a = "'--------------------------------------------------------------"
        PrintLine(fNum, a)
        a = stThta.iPP30.ToString("0")
        b = stThta.iPP31.ToString("0")
        PrintLine(fNum, a, b)                                   ' 補正モード,補正方法

        a = "'----------------------------------------------------"
        PrintLine(fNum, a)
        a = "'ｸﾞﾙｰﾌﾟ番号 ﾊﾟﾀｰﾝ番号   登録位置X(mm) 登録位置Y(mm)"
        PrintLine(fNum, a)
        a = "'----------------------------------------------------"
        PrintLine(fNum, a)
        c(0) = stThta.iPP38.ToString("0")
        c(1) = stThta.iPP37_1.ToString("0")
        c(2) = stThta.fpp32_x.ToString("###0.000#")
        c(3) = stThta.fpp32_y.ToString("###0.000#")
        PrintLine(fNum, c(0), c(1), c(2), c(3))
        c(0) = stThta.iPP38.ToString("0")
        c(1) = stThta.iPP37_2.ToString("0")
        c(2) = stThta.fpp33_x.ToString("###0.000#")
        c(3) = stThta.fpp33_y.ToString("###0.000#")
        PrintLine(fNum, c(0), c(1), c(2), c(3))

        a = "'---------------------------------------"
        PrintLine(fNum, a)
        a = "' θ軸角度      最小角度      最大角度"
        PrintLine(fNum, a)
        a = "'---------------------------------------"
        PrintLine(fNum, a)
        c(0) = stThta.fTheta.ToString("#0.0###")
        c(1) = stThta.fPP53Min.ToString("#0.0###")
        c(2) = stThta.fPP53Max.ToString("#0.0###")
        PrintLine(fNum, c(0), c(1), c(2))                   ' θ軸角度,最小角度,最大角度

        a = "'-----------------------------"
        PrintLine(fNum, a)
        a = "' 補正ﾎﾟｼﾞｼｮﾝｵﾌｾｯﾄXY"
        PrintLine(fNum, a)
        a = "'-----------------------------"
        PrintLine(fNum, a)
        c(0) = stThta.fpp34_x.ToString("#0.000#")
        c(1) = stThta.fpp34_y.ToString("#0.000#")
        PrintLine(fNum, c(0), c(1))                             ' 補正ﾎﾟｼﾞｼｮﾝｵﾌｾｯﾄXY

        a = "/"
        PrintLine(fNum, a)
        a = "'"
        PrintLine(fNum, a)

        '--------------------------------------------------------------------------
        '   ＧＰＩＢデータを設定する
        '--------------------------------------------------------------------------
        a = "'======================================="
        PrintLine(fNum, a)
        a = "' ●ＧＰＩＢデータ"
        PrintLine(fNum, a)
        a = "'======================================="
        PrintLine(fNum, a)

        a = "'---------"
        PrintLine(fNum, a)
        a = "'制御数"
        PrintLine(fNum, a)
        a = "'---------"
        PrintLine(fNum, a)
        a = stPLT.GCount.ToString("0")
        PrintLine(fNum, a)                                  ' 制御数

        ' 制御数分、以下を設定する
        a = "'----------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "'番号          名称        ｱﾄﾞﾚｽ         ﾃﾞﾘﾐﾀ          設定コマンド1-2-3　　　　トリガーコマンド"
        PrintLine(fNum, a)
        a = "'----------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.GCount                           ' 制御数分繰り返す
            c(0) = i.ToString("0")                          ' 番号
            c(1) = """" & stGPIB(i).strGNAM & """"          ' 名称
            c(2) = stGPIB(i).intGAD.ToString("0")           ' ｱﾄﾞﾚｽ
            c(3) = stGPIB(i).intDLM.ToString("0")           ' ﾃﾞﾘﾐﾀ
            'V2.0.0.0④↓
            c(4) = """" & stGPIB(i).strCCMD1 & """"          ' 設定コマンド
            c(5) = """" & stGPIB(i).strCCMD2 & """"          ' 設定コマンド
            c(6) = """" & stGPIB(i).strCCMD3 & """"          ' 設定コマンド
            c(7) = """" & stGPIB(i).strCTRG & """"          ' トリガーコマンド
            'V2.0.0.0④↑
            'V2.0.0.0④            c(4) = """" & stGPIB(i).strCCMD & """"          ' 設定コマンド
            'V2.0.0.0④            c(5) = """" & stGPIB(i).strCTRG & """"          ' トリガーコマンド
            PrintLine(fNum, c(0), c(1), c(2), c(3), c(4), c(5), c(6), c(7))
        Next i

        a = "'----------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        a = "'番号         ＯＮコマンド          ONﾎﾟｰｽﾞ時間(ms)     ＯＦＦコマンド        OFFﾎﾟｰｽﾞ時間(ms)"
        PrintLine(fNum, a)
        a = "'----------------------------------------------------------------------------------------------------------------------------------"
        PrintLine(fNum, a)
        For i = 1 To stPLT.GCount                           ' 制御数分繰り返す
            c(0) = i.ToString("0") ' 番号
            c(1) = """" & stGPIB(i).strCON & """"           ' ＯＮコマンド
            c(2) = stGPIB(i).lngPOWON.ToString("0")         ' ONﾎﾟｰｽﾞ時間(ms)
            c(3) = """" & stGPIB(i).strCOFF & """"          ' ＯＦＦコマンド
            c(4) = stGPIB(i).lngPOWOFF.ToString("0")        ' OFFﾎﾟｰｽﾞ時間(ms)
            PrintLine(fNum, c(0), c(1), c(2), c(3), c(4))
        Next i

        a = "/"
        PrintLine(fNum, a)
        a = "'"
        PrintLine(fNum, a)

        '--------------------------------------------------------------------------
        '   メッセージデータを設定する
        '--------------------------------------------------------------------------
        a = "'======================================="
        PrintLine(fNum, a)
        a = "' ●メッセージ"
        PrintLine(fNum, a)
        a = "'======================================="
        PrintLine(fNum, a)

        ' タイトルを設定する
        a = "' タイトルx0-x9"
        PrintLine(fNum, a)
        For i = 0 To 9
            a = """" & TTL_Msg(i) & """"
            PrintLine(fNum, a)
        Next i

        a = "/"
        PrintLine(fNum, a)

        '--------------------------------------------------------------------------
        '   ユーザデータを設定する
        '--------------------------------------------------------------------------
        PrintLine(fNum, "'=======================================")
        PrintLine(fNum, "' ●ユーザデータ")
        PrintLine(fNum, "'=======================================")

        PrintLine(fNum, "'-------------------------------------")
        PrintLine(fNum, "' 製品種別")
        PrintLine(fNum, "'-------------------------------------")
        PrintLine(fNum, stUserData.iTrimType.ToString("0"))

        PrintLine(fNum, "'-------------------------------------")
        PrintLine(fNum, "' ロット番号")
        PrintLine(fNum, "'-------------------------------------")
        PrintLine(fNum, """" & stUserData.sLotNumber & """")

        PrintLine(fNum, "'-------------------------------------")
        PrintLine(fNum, "' オペレータ名")
        PrintLine(fNum, "'-------------------------------------")
        PrintLine(fNum, """" & stUserData.sOperator & """")

        PrintLine(fNum, "'-------------------------------------")
        PrintLine(fNum, "' パターンＮｏ．      プログラムＮｏ．")
        PrintLine(fNum, "'-------------------------------------")
        PrintLine(fNum, """" & stUserData.sPatternNo & """", """" & stUserData.sProgramNo & """")

        PrintLine(fNum, "'-------------------------------------")
        PrintLine(fNum, "' トリミング速度 1:高速、 2:高精度、3:設定値")
        PrintLine(fNum, "'-------------------------------------")
        PrintLine(fNum, stUserData.iTrimSpeed.ToString("0"))

        PrintLine(fNum, "'---------------------------------------------------------------")
        PrintLine(fNum, "' ロット終了条件 0:終了条件判定無し 1:枚数 2:ローダー信号 3:両方")
        PrintLine(fNum, "'---------------------------------------------------------------")
        PrintLine(fNum, stUserData.iLotChange.ToString("0"))

        PrintLine(fNum, "'-------------------------------------")
        PrintLine(fNum, "' ロット処理枚数")
        PrintLine(fNum, "'-------------------------------------")
        PrintLine(fNum, stUserData.lLotEndSL.ToString("0"))

        PrintLine(fNum, "'-------------------------------------")
        PrintLine(fNum, "' カット位置補正頻度")
        PrintLine(fNum, "'-------------------------------------")
        PrintLine(fNum, stUserData.lCutHosei.ToString("0"))

        PrintLine(fNum, "'-------------------------------------")
        PrintLine(fNum, "' ロット終了時印刷素子数")
        PrintLine(fNum, "'-------------------------------------")
        PrintLine(fNum, stUserData.lPrintRes.ToString("0"))

        'V2.0.0.0⑭↓
        PrintLine(fNum, "'----------------------------------------------")
        PrintLine(fNum, "' 1:クランプ吸着両方 2:クランプのみ 3:吸着のみ")
        PrintLine(fNum, "'----------------------------------------------")
        PrintLine(fNum, stUserData.intClampVacume.ToString("0"))
        'V2.0.0.0⑭↑

        PrintLine(fNum, "'-------------------------------------")
        PrintLine(fNum, "' 温度センサー 抵抗レンジ 1:Ω 2:KΩ")
        PrintLine(fNum, "'-------------------------------------")
        PrintLine(fNum, stUserData.iTempResUnit.ToString("0"))

        PrintLine(fNum, "'-------------------------------------")
        PrintLine(fNum, "' 参照温度 １：０℃ または ２：２５℃")
        PrintLine(fNum, "'-------------------------------------")
        PrintLine(fNum, stUserData.iTempTemp.ToString("0"))

        'V2.0.0.0⑪        PrintLine(fNum, "'-------------------------------------")
        'V2.0.0.0⑪        PrintLine(fNum, "' スタンダード抵抗値　０℃　２５℃　　")
        'V2.0.0.0⑪        PrintLine(fNum, "'-------------------------------------")
        'V2.0.0.0⑪        PrintLine(fNum, stUserData.dStandardRes0.ToString("#0.000#"), stUserData.dStandardRes25.ToString("#0.000#"))

        'V2.1.0.0③        PrintLine(fNum, "'-------------------------------------")
        'V2.1.0.0③        PrintLine(fNum, "' スタンダード抵抗値　０℃　代表α値　代表β値　α値　β値　")
        'V2.1.0.0③        PrintLine(fNum, "'-------------------------------------")
        'V2.1.0.0③        PrintLine(fNum, stUserData.dTemperatura0.ToString("0.0000000"), stUserData.dDaihyouAlpha.ToString("0.0000000"), stUserData.dDaihyouBeta.ToString("0.0000000"), stUserData.dAlpha.ToString("0.0000000"), stUserData.dBeta.ToString("0.0000000"))
        'V2.1.0.0③↓
        PrintLine(fNum, "'-----------------------------------------------------------------------------------------")
        PrintLine(fNum, "' STD抵抗値０℃　代表α値　 代表β値　    α値　        β値　        代表No.     　STDNo.")
        PrintLine(fNum, "'-----------------------------------------------------------------------------------------")
        PrintLine(fNum, stUserData.dTemperatura0.ToString("0.0000000"), stUserData.dDaihyouAlpha.ToString("0.0000000"), stUserData.dDaihyouBeta.ToString("0.0000000"), stUserData.dAlpha.ToString("0.0000000"), stUserData.dBeta.ToString("0.0000000"), stUserData.iTempSensorInfNoDaihyou.ToString("0"), stUserData.iTempSensorInfNoStd.ToString("0"))
        'V2.1.0.0③↑

        PrintLine(fNum, "'-------------------------------------")
        PrintLine(fNum, "' 抵抗温度係数")
        PrintLine(fNum, "'-------------------------------------")
        PrintLine(fNum, stUserData.dResTempCoff.ToString("#0.00000#"))

        PrintLine(fNum, "'-------------------------------------------------------")
        PrintLine(fNum, "' ファイナルリミット Low[%]  ファイナルリミット Hight[%]")
        PrintLine(fNum, "'-------------------------------------------------------")
        PrintLine(fNum, stUserData.dFinalLimitLow.ToString("#0.000#"), stUserData.dFinalLimitHigh.ToString("#0.000#"))

        PrintLine(fNum, "'-----------------------------------------------")
        PrintLine(fNum, "' 相対値リミット Low[%]  相対値リミット Hight[%]")
        PrintLine(fNum, "'-----------------------------------------------")
        PrintLine(fNum, stUserData.dRelativeLow.ToString("#0.000#"), stUserData.dRelativeHigh.ToString("#0.000#"))

        'V2.1.0.0①        PrintLine(fNum, "'--------------------------------------------------------------------------------------------------")
        'V2.1.0.0①        PrintLine(fNum, "' 抵抗レンジ 1:Ω 2:KΩ、補正値（ノミナル値算出係数）、目標値算出係数、測定速度を変更するカットNo.")
        'V2.1.0.0①        PrintLine(fNum, "'--------------------------------------------------------------------------------------------------")
        PrintLine(fNum, "'-----------------------------------------------------------------------------------------------------------------------")
        PrintLine(fNum, "' 抵抗レンジ 1:Ω 2:KΩ、補正値（ノミナル値算出係数）、目標値算出係数、判定用目標値算出係数、測定速度を変更するカットNo.")
        PrintLine(fNum, "'-----------------------------------------------------------------------------------------------------------------------")

        j = GetRCountExceptMeasure()
        If j > MAX_RES_USER Then
            j = MAX_RES_USER
        End If
        For i = 1 To j
            'V2.1.0.0①            PrintLine(fNum, stUserData.iResUnit(i).ToString("0"), stUserData.dNomCalcCoff(i).ToString("#0.00000#"), stUserData.dTargetCoff(i).ToString("#0.0#"), stUserData.iChangeSpeed(i).ToString("0"))
            PrintLine(fNum, stUserData.iResUnit(i).ToString("0"), stUserData.dNomCalcCoff(i).ToString("#0.00000#"), stUserData.dTargetCoff(i).ToString("#0.0#"), stUserData.dTargetCoffJudge(i).ToString("#0.0#"), stUserData.iChangeSpeed(i).ToString("0"))    'V2.1.0.0①
        Next i

        'V2.0.0.0②↓
        PrintLine(fNum, "'--------------------------------------------------------------------------")
        PrintLine(fNum, "' 定格   定格電圧の倍率   抵抗個数 　　  電流制限 　　  印加秒数     変化量")
        PrintLine(fNum, "'--------------------------------------------------------------------------")
        PrintLine(fNum, stUserData.dRated.ToString("0.000"), stUserData.dMagnification.ToString("0.00"), stUserData.dResNumber.ToString("0"), stUserData.dCurrentLimit.ToString("0.00"), stUserData.dAppliedSecond.ToString("0.00"), stUserData.dVariation.ToString("0"))
        'V2.0.0.0②↑

        'V2.0.0.1③↓
        PrintLine(fNum, "'-------------------------------------------------")
        PrintLine(fNum, "' トリミング不良信号(BIT1)出力する１基板のＮＧ比率")
        PrintLine(fNum, "'-------------------------------------------------")
        PrintLine(fNum, stUserData.NgJudgeRate.ToString("#0.00#"))
        'V2.0.0.1③↑

        'V2.2.0.034 ↓
        PrintLine(fNum, "'-------------------------------------------------")
        PrintLine(fNum, "' 複数抵抗用データ(0:通常、1:複数抵抗値)　列/行　")
        PrintLine(fNum, "'-------------------------------------------------")
        PrintLine(fNum, stMultiBlock.gMultiBlock.ToString())
        PrintLine(fNum, stMultiBlock.gStepRpt.ToString())

        PrintLine(fNum, "'-------------------------------------------------")
        PrintLine(fNum, "' 複数ブロックデータ(ブロック数1-5)")
        PrintLine(fNum, "'-------------------------------------------------")
        PrintLine(fNum, stMultiBlock.BLOCK_DATA(0).gBlockCnt.ToString("0"), stMultiBlock.BLOCK_DATA(1).gBlockCnt.ToString("0"), stMultiBlock.BLOCK_DATA(2).gBlockCnt.ToString("0"), stMultiBlock.BLOCK_DATA(3).gBlockCnt.ToString("0"), stMultiBlock.BLOCK_DATA(4).gBlockCnt.ToString("0"))
        PrintLine(fNum, "'-------------------------------------------------")
        PrintLine(fNum, "' 複数ブロックデータ(R1抵抗値 R1単位 R1補正値)")
        PrintLine(fNum, "'-------------------------------------------------")
        For l As Integer = 0 To 4
            For m As Integer = 0 To 4
                PrintLine(fNum, stMultiBlock.BLOCK_DATA(l).dblNominal(m), stMultiBlock.BLOCK_DATA(l).iUnit(m), stMultiBlock.BLOCK_DATA(l).dblCorr(m))
            Next m
        Next
        'V2.2.0.034 ↑


        a = "/"
        PrintLine(fNum, a)

        '--------------------------------------------------------------------------
        '   ファイルクローズ
        '--------------------------------------------------------------------------
        FileClose(fNum)

        rData_save = 0              ' 正常終了

        ' プローブデータファイルを更新    'V2.2.0.0⑮ 
        If stPLT.ProbNo <> 0 Then
            UpdateProbeData(stPLT.ProbNo)
        End If

        Exit Function
STP_END:

        Call Z_PRINT(Err.Description & vbCrLf)

    End Function
#End Region

#Region "ロット情報入力"
    '''=========================================================================
    ''' <summary>ロット情報入力</summary>
    ''' <returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Public Function GetLotInf() As Short

        Dim strMSG As String

        '---------------------------------------------------------------------------
        '   ロット情報をリードする
        '---------------------------------------------------------------------------
        Try

#If cGETcLOTINF Then
            ' 生産数
            stCounter.TrimCounter = GetPrivateProfileInt("LOTINF", "PRODUCT", 0, cLOT_FNAME)
            ' 良品数
            stCounter.OK_Counter = GetPrivateProfileInt("LOTINF", "GOOD", 0, cLOT_FNAME)
#End If
            ' ﾌﾟﾛｰﾌﾞON回数
            stCounter.Probe_Counter = GetPrivateProfileInt("PROBE", "COUNT", 0, cLOT_FNAME)
            ' ﾄﾘﾐﾝｸﾞﾃﾞｰﾀﾌｧｲﾙ名
            gsDataFileName = GetPrivateProfileString_S("DATA", "FNAME", cLOT_FNAME, "")

            Return (cFRS_NORMAL) ' Return値 = 正常

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "GetLotInf() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (-1)                                                 ' Return値 = エラー
        End Try
    End Function
#End Region
#Region "ロット情報出力"
    '''=========================================================================
    '''<summary>ロット情報出力</summary>
    '''<returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Public Function PutLotInf() As Short

        Dim s As String
        Dim r As Integer
        Dim strWAK As String
        Dim strMSG As String

        Try
            '---------------------------------------------------------------------------
            '   ロット情報をライトする
            '---------------------------------------------------------------------------
            ' ロット番号
            r = WritePrivateProfileString("LOTINF", "LOTNUM", stUserData.sLotNumber, cLOT_FNAME)
            ' 生産数
            s = CStr(stCounter.TrimCounter)
            r = WritePrivateProfileString("LOTINF", "PRODUCT", s, cLOT_FNAME)
            ' 良品数
            s = CStr(stCounter.OK_Counter)
            r = WritePrivateProfileString("LOTINF", "GOOD", s, cLOT_FNAME)
            ' ﾌﾟﾛｰﾌﾞON回数
            s = CStr(stCounter.Probe_Counter)
            r = WritePrivateProfileString("PROBE", "COUNT", s, cLOT_FNAME)
            ' ﾄﾘﾐﾝｸﾞﾃﾞｰﾀﾌｧｲﾙ名
            strWAK = gsDataFileName                         ' ファイルパス名が長いと穴が切られて戻ってくるのでワークに退避する 
            r = WritePrivateProfileString("DATA", "FNAME", strWAK, cLOT_FNAME)

            Return (cFRS_NORMAL)                            ' Return値 = 正常

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "ロット情報ファイルライトエラー !!" + ex.Message
            Call Z_PRINT(strMSG & vbCrLf)
            Return (cERR_TRAP)                              ' Return値 = ﾄﾗｯﾌﾟｴﾗｰ発生
        End Try
    End Function
#End Region


#Region "■■ データファイル名(xxx.TXT)からiniファイル名(xxx.INI)を返す ■■"
    '''=========================================================================
    ''' <summary>"データファイル名(xxx.TXT)からiniファイル名(xxx.INI)を返す</summary>
    ''' <param name="gsTxtFileName">(INP)データファイル名</param>
    ''' <param name="gsDatFileName">(OUT)iniファイル名</param>
    '''=========================================================================
    Public Sub Make_Filename_Ini(ByRef gsTxtFileName As String, ByRef gsDatFileName As String)

        Dim strpos As Short
        Dim strMSG As String

        Try
            strpos = InStr(gsTxtFileName, ".")
            gsDatFileName = gsTxtFileName.Substring(0, strpos - 1)
            gsDatFileName = gsDatFileName + ".INI"

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.Make_Filename_Ini() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

    '=========================================================================
    '   トリミング実行制御処理
    '=========================================================================
#Region "ユーザープログラム開始"
    '''=========================================================================
    '''<summary>ユーザープログラム開始</summary>
    '''<returns>cFRS_NORMAL   = 正常
    '''         cFRS_ERR_RST  = RESET SW押下
    '''         cFRS_TRIM_NG  = トリミングNG
    '''         cFRS_ERR_TRIM = トリマエラー
    '''         cFRS_ERR_PTN  = パターン認識エラー
    '''         その他
    ''' </returns>
    '''=========================================================================
    Public Function User() As Short

        Dim r As Short                                                  ' 関数戻値
        Dim iRtn As Short = cFRS_NORMAL                                 ' 関数戻値
        Dim strMSG As String

        Try
            '---------------------------------------------------------------------------
            '   トリミング前処理
            '---------------------------------------------------------------------------
            Form1.Refresh()                                             'V2.1.0.0⑤
            UserSub.CutVariationJudgeExecuteCheck()                     'V2.1.0.0① トリミングデータ単位での抵抗値変化量判定有無チェック

            ' ボタンを無効にする
            Call Form1.Btn_Enb_OnOff(0)                                 ' frmMain画面ボタン非活性化
            Call Form1.SBtn_Enb_OnOff(0)                                ' frmInfo画面ボタン非活性化
            'Call Disp_Result(2, 0)                                     ' OK/NG表示域 = 処理中(黄色表示)
            Form1.BtnADJ.Focus()                                        'V2.0.0.0⑥　ADJボタンにフォーカス移動

            ' ログ画面表示クリア基板枚数を超えたらログ画面をクリアする
            gDspCounter = gDspCounter + 1                               ' ログ画面表示基板枚数カウンタ更新
            If (gDspCounter > gDspClsCount) Then                        ' カウンタ > ログ画面表示クリア基板枚数
                    Z_CLS()                                             ' ログ画面クリア               ###lstLog
                gDspCounter = 1                                         ' ログ画面表示基板枚数カウンタ再設定
            End If

            Call SetLogFileName(gsLogFileName)                          ' ログファイル名設定

            ' セーフティチェック
            r = SafetyCheck()                                           ' セーフティチェック
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
                Return (r)                                              ' Return値設定
            End If

            UserBas.PlateStartSetting()                                 ' 基板処理スタート時間保存

            Call BSIZE(stPLT.zsx, stPLT.zsy)                            ' ブロックサイズ設定

            ' XYﾃｰﾌﾞﾙをトリム位置に移動する
            Call ObjSys.EX_BPOFF(gSysPrm, stPLT.BPOX, stPLT.BPOY)
            r = Move_Trimposition()                                     ' θ補正＆XYﾃｰﾌﾞﾙトリム位置移動
            If (r <> cFRS_NORMAL) Then                                  ' 補正エラー ?
                SetLotMarkAlarm(gsDataFileName, MarkingCount)           ' エラーとなった基板の情報を保存   'V2.2.1.7③
                If (r = cFRS_ERR_PTN Or r = cFRS_ERR_RST) Then          ' ﾊﾟﾀｰﾝ認識ｴﾗｰ ? ###1033 キャンセル追加(cFRS_ERR_RST)
                    stCounter.PlateCounter = stCounter.PlateCounter - 1 ' 基板処理カウンター減算 ###1033
                    iRtn = cFRS_ERR_PTN                                 ' Return値 = ﾊﾟﾀｰﾝ認識ｴﾗｰ
                    GoTo STP_SBACK
                End If
                Return (r)                                              ' Return値設定(非常停止等)
            End If

            'V2.2.1.1⑤↓
            ' 目標値が書き換えられた可能性があるので再度GPIB設定コマンドを送信する
            'V2.2.1.4① GPIB_Init()
            'V2.2.1.1⑤↑
            gLastsetNomx = 0.0   'V2.2.1.4①

            '---------------------------------------------------------------------------
            '   トリミングを実行する
            '---------------------------------------------------------------------------
            ' 画像表示プログラムを起動する
            'V2.2.0.0① r = Exec_GazouProc(ObjGazou, DISPGAZOU_PATH, DISPGAZOU_WRK, INTERNAL_CAMERA)

            Call SETAXISSPDY(stPLT.StageSpeedY)                         ' ###1040④ Ｙ軸ステージ速度の変更機能追加

            iRtn = STRP()                                               ' ステップ＆リピート処理

            Call SETAXISSPDY(SETAXISSPDY_DEFALT)                        ' ###1040④ Ｙ軸ステージ速度を元に戻す。'V2.0.0.0⑮Ｙ軸ステージ速度を元に戻す。25000から15000へ変更

            ' 'V2.2.1.7⑤           If UserSub.IsTRIM_MODE_ITTRFT() And UserSub.IsSpecialTrimType() Then    ' ユーザプログラム特殊処理
            If (UserSub.IsTRIM_MODE_ITTRFT() And UserSub.IsSpecialTrimType()) Or ((DGL = TRIM_MODE_CUT) And UserSub.IsTrimType5()) Then    ' ユーザプログラム特殊処理  'V2.2.1.7⑤
                Call UserSub.SetStartCheckStatus(False)                         ' 設定画面の確認無効化
            End If

            UserSub.VariationMesStartDataReset()                                'V2.0.0.0② 測定値変動検出機能開始ブロック位置初期化

            Call PutLotInf()                                            ' ロット情報出力
            ' 画像表示プログラムを終了する
            'V2.2.0.0① End_GazouProc(ObjGazou)

            ' トリミングNG/ﾊﾟﾀｰﾝ認識ｴﾗｰ/トリマエラー以外のエラーなら終了
            If (iRtn < cFRS_NORMAL) Then                                ' エラー ?
                ' ﾄﾘﾐﾝｸﾞNG/ﾊﾟﾀｰﾝ認識/トリマエラー ?
                If (iRtn = cFRS_TRIM_NG) Or (iRtn = cFRS_ERR_TRIM) Or (iRtn = cFRS_ERR_PTN) Then
                Else                                                    ' その他のｱﾌﾟﾘ終了ﾚﾍﾞﾙｴﾗｰなら終了
                    'Call Disp_Result(1, iRtn)                           ' トリミング結果表示(OK/NG表示)
                    Return (iRtn)                                       ' Return値設定(非常停止等)
                End If
            End If

            '---------------------------------------------------------------------------
            '   トリミング後処理
            '---------------------------------------------------------------------------
STP_SBACK:
            Call ObjSys.Ilum_Ctrl(gSysPrm, Z0, ZOPT)                     ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ消灯(制御有時有効)
            r = Prob_Off()                                              ' Z2/ZﾌﾟﾛｰﾌﾞをOFF位置に移動
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
                Return (r)                                              ' Return値設定
            End If

STP_END:
            'Call Disp_Result(1, iRtn)                                   ' トリミング結果表示(OK/NG表示)
            'If (prf = 1) Then                                           ' 印刷 ?
            '    'Call ObjPrt.Z_LPRINT_ENDDOC()                          ' プリンタ 印字終了（残り排出）
            'End If

            ' ボタンを有効にする
            'V2.0.0.0⑮            Call Form1.SBtn_Enb_OnOff(1)                                ' frmInfo画面ボタン活性化
            'V2.0.0.0⑮            Call Form1.Btn_Enb_OnOff(1)                                 ' frmMain画面ボタン活性化
            Return (iRtn)                                               ' Return値設定

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.User() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cFRS_TRIM_NG)                                       ' Return値 = トリミングNG
        End Try
    End Function
#End Region

#Region "ステップ＆リピート処理"
    '''=========================================================================
    '''<summary>ステップ＆リピート処理</summary>
    '''<returns>cFRS_NORMAL   = 正常
    '''         cFRS_ERR_RST  = RESET SW押下
    '''         cFRS_TRIM_NG  = トリミングNG
    '''         cFRS_ERR_TRIM = トリマエラー
    '''         cFRS_ERR_PTN  = パターン認識エラー
    '''         その他
    ''' </returns>
    '''=========================================================================
    Private Function STRP() As Short

        Dim strMSG As String                                            ' メッセージ表示用域
        Dim strX As String                                              ' メッセージ編集域
        Dim strY As String                                              ' メッセージ編集域
        Dim XV As Short                                                 ' Ｘブロック数
        Dim YV As Short                                                 ' Ｙブロック数
        Dim YV1 As Short
        Dim r As Short                                                  ' 戻り値
        Dim NGflg As Short                                              ' トリミングNGﾌﾗｸﾞ
        Dim dlbOffX As Double = 0
        Dim dlbOffY As Double = 0
        Dim iRtn As Short = cFRS_NORMAL                                 ' 関数戻値
        Dim StepOffX As Double, StepOffY As Double                      ' ステップオフセット用変数

        Try
            ' 初期処理
            STRP = cFRS_NORMAL                                          ' Return値 = 正常
            NGflg = 0                                                   ' トリミングNGﾌﾗｸﾞ初期化
            Call DScanModeReset()                                       ' 測定器の外部内部切り替え状態記録変数の初期化
            Call Z_PRINT(vbCrLf)                                        ' CrLf表示

            beforeExecBlkDataNo = 0                                     ' V2.2.0.033

            ' Ｘブロック数*Ｘプレート数分処理する(X方向←)
            For XV = 1 To stPLT.BNX * stPLT.Pnx
                ' XYテーブル補正(ﾌﾟﾚｰﾄｲﾝﾀｰﾊﾞﾙx分)
                '   現Xﾌﾞﾛｯｸ番号/Xﾌﾞﾛｯｸ数の余り1  で 現Xﾌﾞﾛｯｸ番号 > 1 の時　又は
                '   Xﾌﾞﾛｯｸ数=1で現Xﾌﾞﾛｯｸ番号 > 1 ならXYテーブル補正を行う
                If (((XV Mod stPLT.BNX) = 1) And (YV > 1)) Or ((stPLT.BNX = 1) And (XV > 1)) Then
                    dlbOffX = stPLT.Pivx
                End If
                stCounter.PlateCntX = Int(XV / stPLT.BNX) + 1

                ' Ｙブロック数*Ｙプレート数分処理する(Y方向↓)
                For YV = 1 To stPLT.BNY * stPLT.Pny
                    System.Windows.Forms.Application.DoEvents()
                    ' 奇数Xブロック時
                    If ((XV Mod 2) = 1) Then
                        ' 現Yﾌﾞﾛｯｸ番号/Yﾌﾞﾛｯｸ数の余り1  で 現Yﾌﾞﾛｯｸ番号 > 1 の時 XYテーブル補正(ﾌﾟﾚｰﾄｲﾝﾀｰﾊﾞﾙY分)を行う
                        If ((YV Mod stPLT.BNY) = 1) And (YV > 1) Then dlbOffY = stPLT.Pivy
                        YV1 = YV                                        ' 現Yﾌﾞﾛｯｸ番号 = +現Yﾌﾞﾛｯｸ番号
                        stCounter.PlateCntY = Int(YV / stPLT.BNY) + 1
                        ' 偶数Xブロック時
                    Else
                        ' 現Yﾌﾞﾛｯｸ番号/Yﾌﾞﾛｯｸ数の余り0で
                        ' 現Yﾌﾞﾛｯｸ番号 < Yﾌﾞﾛｯｸ数*ﾌﾟﾚｰﾄ数=1 ならXYテーブル補正(-ﾌﾟﾚｰﾄｲﾝﾀｰﾊﾞﾙ-Y分)を行う
                        If ((YV Mod stPLT.BNY) = 1) And (YV > 1) Then dlbOffY = -stPLT.Pivy
                        ' 現Yﾌﾞﾛｯｸ番号 =Yﾌﾞﾛｯｸ数 * Yﾌﾟﾚｰﾄ数 + 1 -現Yﾌﾞﾛｯｸ番
                        YV1 = stPLT.BNY * stPLT.Pny + 1 - YV
                        stCounter.PlateCntY = Int(YV1 / stPLT.BNY) + 1
                    End If

                    'V2.0.0.0②↓
                    If DGL = TRIM_VARIATION_MEAS Then ' 測定値変動測定
                        If UserSub.bVariationMesStep Then
                            If ((XV - 1) Mod stPLT.BNX + 1) = UserSub.gVariationMeasBlockXStartNo And ((YV1 - 1) Mod stPLT.BNY + 1) = UserSub.gVariationMeasBlockYStartNo Then
                                UserSub.bVariationMesStep = False
                            Else
                                Continue For
                            End If
                        End If
                    End If
                    'V2.0.0.0②↑
                    UserBas.InitForStepMove()                               ' 移動状態確認フラグリセット
                    ' XYテーブル指定ブロック移動
                    Call BSIZE(stPLT.zsx, stPLT.zsy)
#If TXTY_USE Then
                    If stPLT.BNY > 1 Then
                        StepOffX = stPLT.dblStepOffsetXDir / (stPLT.BNY - 1) * (YV1 - 1)
                    Else
                        StepOffX = 0.0
                    End If
                    If stPLT.BNX > 1 Then
                        StepOffY = stPLT.dblStepOffsetYDir / (stPLT.BNX - 1) * (XV - 1)
                    Else
                        StepOffY = 0.0
                    End If

                    r = TSTEP(CShort(XV), CShort(YV1), dlbOffX + StepOffX, dlbOffY + StepOffY)
#Else
                    r = TSTEP(CShort(XV), CShort(YV1), dlbOffX, dlbOffY)
#End If
                    r = ObjSys.EX_ZGETSRVSIGNAL(gSysPrm, r, 0)

                    If (r <> cFRS_NORMAL) Then                          ' エラー ?
                        STRP = r                                        ' Return値設定(非常停止等)
                        Exit Function
                    End If

                    ' ログ画面に処理中のﾌﾟﾚｰﾄ番号とﾌﾞﾛｯｸ番号を表示する
                    stCounter.BlockCntX = (XV - 1) Mod stPLT.BNX + 1
                    stCounter.BlockCntY = (YV1 - 1) Mod stPLT.BNY + 1
                    strX = stCounter.BlockCntX.ToString("000")
                    strY = stCounter.BlockCntY.ToString("000")
                    strMSG = "■BLOCK(" & strX & "," & strY & ")"       ' ﾌﾞﾛｯｸ番号表示
                    Call Z_PRINT(strMSG & vbCrLf)
                    If (prf = 1) Then                                   ' 印刷 ?
                        'Call ObjPrt.Z_LPRINT(strMSG)
                    End If

                    ' ログ出力用ﾌﾞﾛｯｸ番号(X,Y)
                    gsLogBlock = "       " & strX & "-" & strY & ","

                    'V2.2.0.0⑯↓            'V2.2.2.0③ 
                    stExecBlkData.DataNo = 0
                    'If stMultiBlock.gMultiBlock <> 0 Then
                    '    ' 複数抵抗値取得データの場合データを展開する 
                    '    r = ApplyMultiData(stCounter.BlockCntX, stCounter.BlockCntY)
                    '    'V2.2.0.033↓
                    '    If stExecBlkData.DataNo <> beforeExecBlkDataNo AndAlso stExecBlkData.DataNo <> 0 Then
                    '        UserSub.ResetlResCounterForPrinter()            ' 印刷素子カウンターのリセット 
                    '        beforeExecBlkDataNo = stExecBlkData.DataNo
                    '    End If
                    '    'V2.2.0.033↑
                    'End If
                    'V2.2.0.0⑯↑            'V2.2.2.0③ 

                    stCounter.BlockCounter = stCounter.BlockCounter + 1 ' ログに出力する為に先にカウントアップする。
                    ' トリミング処理
                    r = Form_()                                         ' トリミング処理

                    UserBas.GetPosForStepMove(XV, YV)                   ' ステップ移動していた場合は位置を変更する。

                    If stMultiBlock.gMultiBlock <> 0 Then
                        ' 複数抵抗値取得データの場合データを展開する 
                        RestoreMultiData(stCounter.BlockCntX, stCounter.BlockCntY)
                    End If

                    ' トリミング結果表示(OK/NG表示)
                    If (strJUG(0) = JG_OK) Or (strJUG(0) = JG_RS) Then
                        iRtn = cFRS_NORMAL
                    Else
                        'V2.0.0.1③                        iRtn = cFRS_ERR_TRIM
                        iRtn = cFRS_TRIM_NG                 'V2.0.0.1③
                    End If

                    Call Disp_Result(1, iRtn)                                   ' トリミング結果表示(OK/NG表示)

                    If (r = cFRS_ERR_RST) Then                          ' RESETｷｰ押下なら終了
                        STRP = r                                        ' 戻り値設定
                        Call Disp_Result(0, 0)                                      ' トリミング結果表示(OK/NG表示)
                        Exit Function
                    End If

                    ' トリミングNG以外のエラーなら終了
                    If (r <> cFRS_NORMAL) Then                          ' エラー ?
                        ' ﾄﾘﾐﾝｸﾞNGなら続行
                        If (r = cFRS_TRIM_NG) Or (r = cFRS_ERR_PTN) Then
                            NGflg = 1                                   ' トリミングNGﾌﾗｸﾞON
                        Else                                            ' その他のｱﾌﾟﾘ終了ﾚﾍﾞﾙｴﾗｰなら終了
                            STRP = r                                    ' Return値設定(非常停止等)
                            Exit Function
                        End If
                    End If
                Next YV                                                 ' 次 Y方向ブロック処理へ
            Next XV                                                     ' 次 X方向ブロック処理へ

STRP_END:

            If (NGflg = 1) Then                                         ' トリミングNG ?
                Return (cFRS_TRIM_NG)                                   ' Return値 = トリミングNG
            End If
            Return (cFRS_NORMAL)                                        ' Return値 = 正常

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.STRP() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cFRS_TRIM_NG)                                       ' Return値 = トリミングNG
        End Try
    End Function
#End Region

#Region "トリミングフォーム処理"
    '''=========================================================================
    '''<summary>トリミングフォーム処理</summary>
    '''<returns>cFRS_NORMAL   = 正常
    '''         cFRS_ERR_RST  = RESET SW押下
    '''         cFRS_TRIM_NG  = トリミングNG
    '''         cFRS_ERR_TRIM = トリマエラー
    '''         cFRS_ERR_PTN  = パターン認識エラー
    '''         その他
    ''' </returns>
    '''=========================================================================
    Private Function Form_() As Short

        Dim intRtCd As Short = cFRS_NORMAL                              ' 戻値
        Dim r As Integer                                                ' 関数戻値
        Dim Sts As Integer
        Dim strMSG As String                                            ' メッセージ表示用域
        Dim xPos As Double                                              ' BP現在値X(ｸﾛｽﾗｲﾝ補正用)
        Dim yPos As Double                                              ' BP現在値Y(ｸﾛｽﾗｲﾝ補正用)
        'Dim ObjGazou As Process = Nothing                               ' Processオブジェクト
        Dim LdIn As Integer

        Try
            ' 初期処理
            If (DGL = TRIM_MODE_STPRPT) Then                            ' ｽﾃｯﾌﾟ＆ﾘﾋﾟｰﾄﾓｰﾄﾞ ?
                Call ObjSys.Ilum_Ctrl(gSysPrm, Z1, ZOPT)                ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ点灯(制御有時有効)
            End If

            ' BPｵﾌｾｯﾄ設定(無いとダメ)
            Call ObjSys.EX_BPOFF(gSysPrm, stPLT.BPOX, stPLT.BPOY)
            ' 最初の抵抗内カット開始位置にBP最高速移動(絶対値)
            ' ###1040②            Call ObjSys.EX_MOVE(gSysPrm, stREG(1).STCUT(1).dblSTX, stREG(1).STCUT(1).dblSTY, 1)
            Call ObjSys.EX_MOVE(gSysPrm, 0, 0, 1)                       ' ###1040② BPオフセットへ移動

            ' ADJボタンがONならSTART(1)/RESET(3)押下待ちとする
            Call HALT_SWCHECK(Sts)                                      ' HALT SWチェック  
            If (Sts = cSTS_HALTSW_ON) Or Form1.GetBtnADJStatus() Then
                If (Sts = cSTS_HALTSW_ON) Then Form1.BtnADJ.PerformClick()
                Call ZCONRST()                                          ' ラッチクリア
                Form1.BtnADJ.Focus()                                    ' ###005
                Form1.StepMoveButtonOn()
                Call ZGETBPPOS(xPos, yPos)                              ' BP現在位置取得
                ObjCrossLine.CrossLineDispXY(xPos, yPos)                ' クロスライン表示
                giAppMode = APP_MODE_FINEADJ                            ' ###1040②
                gbExitFlg = False                                       ' ###1040②

                Form1.ChkLoaderInfoDisp(0)                              'V2.2.0.0⑤

                If giLoaderType = 1 Then
                    'V2.2.0.0① ↓
                    '倍率変更バーの表示
                    SetMagnifyBar(True)                                     ' V2.2.0.0①
                    'V2.2.0.0① ↑

                    ' ローダ停止まで待つ
                    r = ObjLoader.WaitLoaderStop()
                    If (r <> cFRS_NORMAL) Then
                        Return (cFRS_ERR_RST)                           ' Return値 = RESETｷｰ
                    End If

                    Call ZCONRST()                                              ' コンソールキーラッチ解除

                    ' 電磁ロック(観音扉右側ロック)を解除する
                    r = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_OFF)
                    If (r = cFRS_TO_EXLOCK) Then                                ' 「前面扉ロック解除タイムアウト」なら戻り値を「RESET」にする
                        r = cFRS_ERR_RST
                        Return (r)
                    End If
                    Call COVERCHK_ONOFF(1)                         ' 「固定カバー開チェックなし」にする
                    r = COVERCHK_ONOFF(1)
                    System.Windows.Forms.Application.DoEvents()
                End If

                Form1.Refresh()
                gObjADJ = New frmFineAdjust()                           ' ###1040②
                ' 一時停止用微調整フォームの表示                        ' ###1040②
                Call gObjADJ.SetInitialData(gSysPrm, 0, 0, 1, 1)        ' ###1040②
                Call gObjADJ.Focus()                                    ' ###1040②
                Call gObjADJ.Show()                                     ' ###1040②
                'V2.2.0.0①  End_GazouProc(ObjGazou)                                 ' 画像表示プログラムを終了する
                Call ObjSys.Ilum_Ctrl(gSysPrm, Z1, ZOPT)                ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ点灯(制御有時有効)
                Do
                    ' セーフティチェック
                    r = SafetyCheck()                                   ' セーフティチェック
                    If (r <> cFRS_NORMAL) Then                          ' エラー ?
                        Form1.StepMoveButtonOff()
                        If (gObjADJ Is Nothing = False) Then
                            Call gObjADJ.Sub_StopTimer()
                            Call gObjADJ.Close()                                '  オブジェクト開放
                            Call gObjADJ.Dispose()                              '  リソース開放
                            gObjADJ = Nothing
                        End If
                        Return (r)                                      ' Return値設定
                    End If

                    ' ローダアラーム/非常停止チェック
                    If giLoaderType = 1 Then        '@@@888 
                        r = ObjLoader.GetLoaderIO(LdIn)                                   ' ローダ
                        If ((LdIn And clsLoaderIf.LINP_NO_ALM_RESTART) <> clsLoaderIf.LINP_NO_ALM_RESTART) Then
                            Form1.StepMoveButtonOff()
                            If (gObjADJ Is Nothing = False) Then
                                Call gObjADJ.Sub_StopTimer()
                                Call gObjADJ.Close()                                '  オブジェクト開放
                                Call gObjADJ.Dispose()                              '  リソース開放
                                gObjADJ = Nothing
                            End If
                            r = cFRS_ERR_LDR                              ' Return値 = ローダアラーム検出
                            GoTo STP_ERR_LDR                                    ' ローダアラーム表示へ
                        End If
                    End If

                    ' START(1)/RESET(3)押下チェック
                    r = STARTRESET_SWCHECK(1, Sts)
                    If (Sts = cFRS_ERR_START) Then
                        Exit Do                                         ' STARTキー押下なら抜ける
                    End If
                    If (Sts = cFRS_ERR_RST) Then                        ' RESETキー押下ならEXIT 
                        Call ZCONRST()                                  ' ラッチクリア
                        Form1.StepMoveButtonOff()
                        ObjCrossLine.CrossLineOff()                             ' クロスライン非表示
                        If (gObjADJ Is Nothing = False) Then
                            Call gObjADJ.Sub_StopTimer()
                            Call gObjADJ.Close()                                '  オブジェクト開放
                            Call gObjADJ.Dispose()                              '  リソース開放
                            gObjADJ = Nothing
                        End If
                        Return (cFRS_ERR_RST)                           ' Return値 = RESETｷｰ
                    End If

                    System.Threading.Thread.Sleep(100)                  ' Wait(ms)
                    System.Windows.Forms.Application.DoEvents()
                    ' ###1040②                Loop While (1)
                Loop While (Not gbExitFlg)
ADJ_START:
                Call ObjSys.Ilum_Ctrl(gSysPrm, Z0, ZOPT)                ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ消灯(制御有時有効)
                Form1.StepMoveButtonOff()

                If giLoaderType = 1 Then
                    Call COVERLATCH_CLEAR()                             ' カバーラッチクリア
                    Call COVERCHK_ONOFF(0)                              '「固定カバー開チェックなし」にする

                    r = WaitCoverClose()                                ' 筐体カバーの閉を待つ

                    ' 電磁ロック(観音扉右側ロック)を解除する
                    r = ObjLoader.EL_Lock_OnOff(ObjLoader.EX_LOK_MD_ON)
                    If (r = cFRS_TO_EXLOCK) Then                                ' 「前面扉ロック解除タイムアウト」なら戻り値を「RESET」にする
                        r = cFRS_ERR_RST
                        Return (r)
                    End If
                End If

                ObjCrossLine.CrossLineOff()                             ' クロスライン非表示
                'V2.2.0.0① Exec_GazouProc(ObjGazou, DISPGAZOU_PATH, DISPGAZOU_WRK, INTERNAL_CAMERA)  ' 画像表示プログラムを起動する

                Form1.ChkLoaderInfoDisp(1)                              'V2.2.0.0⑤

                '倍率変更バーの非表示
                SetMagnifyBar(False)                                         ' V2.2.0.0①

                giAppMode = APP_MODE_TRIM                               ' ###1040②
                If (gObjADJ Is Nothing = False) Then                    ' ###1040②
                    r = gObjADJ.GetReturnVal()                              ' ###1040②
                    Call gObjADJ.Sub_StopTimer()                        ' ###1040②
                    Call gObjADJ.Close()                                ' ###1040② オブジェクト開放
                    Call gObjADJ.Dispose()                              ' ###1040② リソース開放
                    gObjADJ = Nothing                                   ' ###1040②
                End If                                                  ' ###1040②
                If (r = cFRS_ERR_RST) Then                              ' ###1040②
                    Return (cFRS_ERR_RST)                               ' ###1040②
                End If                                                  ' ###1040②
            End If
            Call ZCONRST()                                              ' ラッチクリア
            Call Disp_Result(2, 0)                                      ' OK/NG表示域 = 処理中(黄色表示)

            ' セーフティチェック
            r = SafetyCheck()                                           ' セーフティチェック
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
                Return (r)                                              ' Return値設定
            End If

            If (DGL <> TRIM_MODE_STPRPT) Then                           ' ｽﾃｯﾌﾟ＆ﾘﾋﾟｰﾄﾓｰﾄﾞ以外 ?
                Call ObjSys.Ilum_Ctrl(gSysPrm, Z0, ZOPT)                ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ消灯(制御有時有効)
            End If

            'V2.2.2.0③ ↓
            If stMultiBlock.gMultiBlock <> 0 Then
                ' 複数抵抗値取得データの場合データを展開する 
                r = ApplyMultiData(stCounter.BlockCntX, stCounter.BlockCntY)
                'V2.2.0.033↓
                If stExecBlkData.DataNo <> beforeExecBlkDataNo AndAlso stExecBlkData.DataNo <> 0 Then
                    UserSub.ResetlResCounterForPrinter()            ' 印刷素子カウンターのリセット 
                    beforeExecBlkDataNo = stExecBlkData.DataNo
                End If
                'V2.2.0.033↑
            End If
            'V2.2.2.0③ ↑

            DScanModeReset()

            ''V2.0.0.0⑰↓
            'V2.2.0.0①  Exec_GazouProc(ObjGazou, DISPGAZOU_PATH, DISPGAZOU_WRK, INTERNAL_CAMERA)  ' 画像表示プログラムを起動する
            ''V2.0.0.0⑰↑
            '' 画像表示プログラムを起動する
            'r = Exec_GazouProc(ObjGazou, DISPGAZOU_PATH, DISPGAZOU_WRK, 0)

            Form1.BtnADJ.Focus()                                        'V2.0.0.0⑥　ADJボタンにフォーカス移動

            ' トリミングフォーム処理
            Select Case (DGL)
                ' SW = x0ﾓｰﾄﾞ(ﾄﾘﾐﾝｸﾞﾓｰﾄﾞ) 'V1.0.4.3⑩測定マーキングモード・ファイナル測定のみ(TRIM_MODE_MEAS_MARK)追加 'V2.0.0.0②測定値変動測定(TRIM_VARIATION_MEAS)追加
                Case TRIM_MODE_ITTRFT, TRIM_MODE_MEAS_MARK, TRIM_VARIATION_MEAS
                    intRtCd = Trim_()                                   ' トリミング実行

                Case TRIM_MODE_MEAS                                     ' SW = x3ﾓｰﾄﾞ(測定ﾓｰﾄﾞ)
                    intRtCd = Meas()                                    ' 測定

                Case TRIM_MODE_CUT                                      ' SW = x5ﾓｰﾄﾞ(ｶｯﾄﾁｪｯｸﾓｰﾄﾞ)
                    intRtCd = CUT_CHK()                                 ' カットチェック

                Case TRIM_MODE_STPRPT                                   ' SW = x6ﾓｰﾄﾞ(ｽﾃｯﾌﾟ＆ﾘﾋﾟｰﾄ)
                    intRtCd = STPRP()                                   ' ＸＹテーブルチェック処理

                Case TRIM_MODE_POWER                                    ' 電源モード'V2.0.0.0②
                    intRtCd = Power()                                   ' 電源モード'V2.0.0.0②

                Case Else                                               ' 未使用SW ?
                    strMSG = "ＤＧ－ＳＷ　ＥＲＲＯＲ"
                    Call Z_PRINT(strMSG & vbCrLf)
                    intRtCd = cFRS_TRIM_NG                              ' Return値 = トリミングNG
                    GoTo STP_END
            End Select

            ' 後処理
            If (DGL <> TRIM_MODE_STPRPT) Then                           ' ｽﾃｯﾌﾟ＆ﾘﾋﾟｰﾄﾓｰﾄﾞ以外 ?
                Call ObjSys.Ilum_Ctrl(gSysPrm, Z0, ZOPT)                 ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ消灯(制御有時有効)
            End If

STP_END:

            Form1.KeyPreview = True                                 ' ファンクションキーを受ける様にする。
            Form1.Activate()

            Return (intRtCd)                                            ' Return値設定


STP_ERR_LDR:
            Dim AlarmKind As Integer
            Dim rtnCode As Integer

            AlarmKind = cGMODE_LDR_ALARM
            rtnCode = ObjSys.Sub_CallFormLoaderAlarm(AlarmKind, ObjPlcIf)
            Call ObjSys.W_RESET()                                      ' アラームリセット信号送出
            Call ObjSys.W_START()                                      ' スタート信号送出
            If rtnCode = cFRS_ERR_START Then
                GoTo ADJ_START
            End If

            Return (rtnCode)                                            ' Return値設定


            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.Form_() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            intRtCd = cFRS_TRIM_NG                                      ' Return値 = トリミングNG
            GoTo STP_END
        End Try
    End Function
#End Region

    '=========================================================================
    '   ステップ＆リピート処理
    '=========================================================================
#Region "ＸＹテーブルチェック処理"
    '''=========================================================================
    '''<summary>ＸＹテーブルチェック処理(ｽﾃｯﾌﾟ＆ﾘﾋﾟｰﾄ)</summary>
    '''<returns>cFRS_NORMAL   = 正常
    '''         cFRS_ERR_RST  = RESET SW押下
    '''         cFRS_TRIM_NG  = トリミングNG
    '''         cFRS_ERR_TRIM = トリマエラー
    '''         cFRS_ERR_PTN  = パターン認識エラー
    '''         その他
    ''' </returns>
    '''=========================================================================
    Private Function STPRP() As Short

        Dim r As Short
        Dim X As Double
        Dim y As Double
        Dim strMSG As String

        Try
            ' ＸＹテーブルチェック処理
            STPRP = cFRS_NORMAL                                         ' Return値 = 正常
            Call ObjSys.Ilum_Ctrl(gSysPrm, Z1, ZOPT)                     ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ点灯(制御有時有効)
            Call Disp_Init()                                            ' 見出し表示(ログ画面)/印刷
            X = stREG(1).STCUT(1).dblSTX                                ' 第一抵抗の第一カット位置X
            y = stREG(1).STCUT(1).dblSTY                                ' 第一抵抗の第一カット位置Y
            r = ObjSys.EX_MOVE(gSysPrm, X, y, 1)                         ' BP移動(第一抵抗の第一カット位置)(絶対値)
            Return (r)                                                  ' (注)エラー時のメッセージは表示済み

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "STPRP() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cFRS_TRIM_NG)
        End Try
    End Function
#End Region

    '=========================================================================
    '   トリミング処理
    '=========================================================================
#Region "トリミング処理"
    '''=========================================================================
    ''' <summary>トリミング処理</summary>
    ''' <returns>cFRS_NORMAL   = 正常
    '''         cFRS_ERR_RST  = RESET SW押下
    '''         cFRS_TRIM_NG  = トリミングNG
    '''         cFRS_ERR_TRIM = トリマエラー
    '''         cFRS_ERR_PTN  = パターン認識エラー
    '''         その他
    ''' </returns>
    ''' <remarks>１ブロック分のトリミングを行う</remarks>
    '''=========================================================================
    Private Function Trim_() As Short

        Dim rtn As Short = cFRS_NORMAL                                  ' 関数戻値
        Dim r As Short                                                  ' 関数戻値
        Dim i As Short                                                  ' Index
        Dim rn As Short                                                 ' 抵抗番号
        Dim dblMx As Double = 0.0                                       ' 測定値(V)
        Dim FlgRetry As Short                                           ' ﾌﾟﾛｰﾌﾞﾘﾄﾗｲ ﾌﾗｸﾞ
        Dim strMSG As String                                            ' メッセージ編集域
        Dim FinalJudgeNG As Boolean = False                             ' 最終判定がNGの時 True
        Dim BlockJudgeNG As Boolean = False                             ' ブロックがNGの時 True
        Dim JudgeNG As Boolean = False                                  ' スキップ(bSkip)判定のリセット前の保持用
        Dim NetworkSkip As Boolean = False                              ' ネットワーク抵抗の時のエラースキップ
        Dim MesTime As Integer
        Dim Retry_Cnt As Integer
        Dim Judge As Integer                                            ' 判定結果'V2.0.0.0⑨
        Dim bStdJudgeNG As Boolean = False                              ' V2.0.0.0⑮ スタンダード測定の判定保存
        Dim bCircuit As Boolean = False                                 'V2.0.0.0⑨サーキットかの判定フラグ

        Try
            '---------------------------------------------------------------------------
            '   初期処理
            '---------------------------------------------------------------------------
            If UserSub.IsTrimType3() And UserBas.GetRCountExceptMeasure() > 1 Then
                bCircuit = True
            End If

            ' 作業域初期化
            For i = 0 To MAXRNO                                         ' 抵抗数分繰返す
                strJUG(i) = JG_SP                                       ' 判定初期化
            Next i
            strJUG(0) = JG_OK                                           ' 判定(1ﾌﾞﾛｯｸ) = "OK"

            'V2.2.1.7③↓
            ' マーク印字モードはトリミングでは処理しない 
            If UserSub.IsTrimType5() Then
                Call Z_PRINT("製品種別マーク印字はx2モードで実行してください。 " & vbCrLf)
                Return (cFRS_ERR_RST)
            End If
            'V2.2.1.7③↑

            UserSub.InitResPatternmatchResult()                         'V1.2.0.0③パターン認識結果格納領域の初期化（初期状態はＯＫ）

            ' セーフティチェック
            r = SafetyCheck()                                           ' セーフティチェック
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
                strJUG(0) = JG_ER                                       ' 判定(1ﾌﾞﾛｯｸ) = "ERROR"
                rtn = r                                                 ' Return値 = セーフティチェックエラー
                GoTo Trim_EXT
            End If

            ' パターン認識処理（カット位置補正用）
            r = Ptn_Match_Exe()                                         ' パターン認識実行
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
                If (r = cFRS_ERR_PTN) Then                              ' パターン認識エラー ?
                    strJUG(0) = JG_PT                                   ' 判定(1ﾌﾞﾛｯｸ) = ﾊﾟﾀｰﾝ認識ｴﾗｰ
                    rtn = cFRS_ERR_PTN                                  ' Return値 = パターン認識エラー
                    GoTo Trim_EXT
                ElseIf (r = cFRS_ERR_RST) Then                          ' キャンセル ?
                    strJUG(0) = JG_RS                                   ' 判定(1ﾌﾞﾛｯｸ) = REST SW押下
                    rtn = cFRS_ERR_RST                                  ' Return値 = REST SW押下
                    GoTo Trim_EXT
                Else
                    strJUG(0) = JG_ER                                   ' 判定(1ﾌﾞﾛｯｸ) = "ERROR"
                    rtn = r                                             ' Return値設定
                    GoTo Trim_EXT
                End If
            End If

            ' ﾎﾟｰｽﾞ付きﾌﾟﾛｰﾌﾞON(ZﾌﾟﾛｰﾌﾞをON位置(Z.ZON)に移動)
            r = Prob_On()                                               ' ﾌﾟﾛｰﾌﾞON
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
                strJUG(0) = JG_ER                                       ' 判定(1ﾌﾞﾛｯｸ) = "ERROR"
                rtn = r                                                 ' Return値設定
                GoTo Trim_EXT
            End If

            ' 見出し表示(ログ画面)/印刷
            Call Disp_Init()

            'V2.0.0.0②↓
            If DGL = TRIM_VARIATION_MEAS Then ' 測定値変動測定
                Call UserSub.SetTarrgetOnVariationMeas()
            End If
            'V2.0.0.0②↑

            '---------------------------------------------------------------------------
            '   トリミング実行処理
            '---------------------------------------------------------------------------
            ' 抵抗数分以下の処理を行う
            For rn = 1 To stPLT.RCount                                  ' 抵抗数分繰返す

                ' 作業域初期化
                dblMx = 0.0                                             ' 測定値
                dblNM(1) = stREG(rn).dblNOM                             ' IT目標値
                dblNM(2) = stREG(rn).dblNOM                             ' FT目標値
                dblVX(1) = 0.0#                                         ' IT測定値
                dblVX(2) = 0.0#                                         ' FT測定値
                FlgRetry = 0                                            ' ﾌﾟﾛｰﾌﾞﾘﾄﾗｲ ﾌﾗｸﾞ初期化
                UserSub.VariationCutNGCutNoReset()                      'V2.1.0.0⑤カット毎の抵抗値変化量判定エラーカット番号初期化

                ' セーフティチェック
                r = SafetyCheck()                                       ' セーフティチェック
                If (r <> cFRS_NORMAL) Then                              ' エラー ?
                    strJUG(rn) = JG_ER                                  ' 判定 = "ERROR"
                    strJUG(0) = JG_ER                                   ' 判定(1ﾌﾞﾛｯｸ) = "ERROR"
                    rtn = r                                             ' Return値 = セーフティチェックエラー
                    GoTo Trim_EXT
                End If

                If IsCutResistor(rn) And NetworkSkip Then               ' ネットワーク抵抗でＮＧ発生時は、以降のトリミングを行わない。
                    GoTo Trim_DSP
                End If

                'V1.2.0.0②↓
                If Not stREG(rn).bPattern Then                          'V1.2.0.0③ カット位置補正の判定 True：OK False:NG
                    strJUG(rn) = JG_PT                                  ' 判定 = "NG-PT" ﾊﾟﾀｰﾝ認識ｴﾗｰ
                    strJUG(0) = JG_PT                                   ' 判定(1ﾌﾞﾛｯｸ) ﾊﾟﾀｰﾝ認識ｴﾗｰ
                    FinalJudgeNG = True
                    DebugLogOut("パターン認識エラー(FinalJudgeNG = True)抵抗[" & rn.ToString & "]")
                    UserSub.NgJudgeSet()                                ' 素子毎のＮＧ判定
                    GoTo Trim_DSP
                End If
                'V1.2.0.0②↑

                ' パターン認識NGの抵抗はSKIP
                ' 特殊処理（パターン認識ＮＧの場合は、前回の補正値を使用する）
                'If (gTblPtn(rn) = 1) Then                               ' パターン認識NG ?
                '    strJUG(rn) = JG_PT                                  ' 判定 = ﾊﾟﾀｰﾝ認識ｴﾗｰ
                '    strJUG(0) = JG_PT                                   ' 判定(1ﾌﾞﾛｯｸ) = ﾊﾟﾀｰﾝ認識ｴﾗｰ
                '    rtn = cFRS_ERR_PTN                                  ' Return値 = パターン認識エラー
                '    GoTo Trim_DSP
                'End If


                If UserSub.IsSpecialTrimType And IsCutResistor(rn) Then     ' トリミング抵抗の時
                    UserSub.CalcTargeResistancetValue(rn)
                    dblNM(2) = UserSub.GetTRV()                             ' 目標値
                End If

                'V2.0.0.0②↓
                If bPowerOnOffUse Then
                    If FUNC_OK = Func_V_On_Judge(rn) Then                   '   電圧ON有り？
                        r = Func_V_On_Ex(rn)                                '   電圧ON
                        If (FUNC_NG = r) Then
                            strJUG(rn) = JG_ER                                  ' 判定 = ｴﾗｰ発生(電圧設定等)
                            strJUG(0) = JG_ER                                   ' 判定(1ﾌﾞﾛｯｸ) = ｴﾗｰ発生(電圧設定等)
                            rtn = cFRS_ERR_TRIM                                 ' トリマエラー
                            GoTo Trim_EXT
                        End If
                    End If
                End If
                'V2.0.0.0②↑

                '-----------------------------------------------------------------------
                '   初期値を測定する(IT)
                '-----------------------------------------------------------------------
                'V2.0.0.0                If DGL <> TRIM_MODE_MEAS_MARK And ((stREG(rn).intSLP <> SLP_VMES And stREG(rn).intSLP <> SLP_RMES And stREG(rn).intSLP <> SLP_NG_MARK And stREG(rn).intSLP <> SLP_OK_MARK)) Then   ' 5:電圧測定のみ, 6:抵抗測定のみ 7:NGﾏｰｷﾝｸﾞ でない場合にIT測定を行う。'V1.0.4.3⑤ ＯＫマーキング(SLP_OK_MARK)追加
                If DGL <> TRIM_VARIATION_MEAS And DGL <> TRIM_MODE_MEAS_MARK And IsMeasureMode(rn, MEAS_JUDGE_IT) Then    'V2.0.0.0

                    Call DScanModeSet(rn, 0, 0)                             ' DCスキャナに接続する測定器を切替る 


STP_RETRY:
                    ' 抵抗測定/電圧測定(内部/外部測定器)
                    'V2.0.0.0⑧                    If stREG(rn).intMType = 1 Then                          ' 外部測定器
                    'V2.0.0.0⑧                        MesTime = gGpibMultiMeterCount
                    'V2.0.0.0⑧                    Else
                    'V2.0.0.0⑧                        'MesTime = 1            '20130418
                    'V2.0.0.0⑧                        MesTime = 2             '20130418
                    'V2.0.0.0⑧                    End If
                    MesTime = stREG(rn).intITReMeas                             'V2.0.0.0⑧

                    For i = 1 To MesTime
                        If UserSub.IsSpecialTrimType And IsCutResistor(rn) Then
                            r = V_R_MEAS(stREG(rn).intSLP, stREG(rn).intMType, dblMx, rn, UserSub.GetTRV())
                        Else
                            r = V_R_MEAS(stREG(rn).intSLP, stREG(rn).intMType, dblMx, rn, stREG(rn).dblNOM)
                        End If
                    Next
                    If (r <> cFRS_NORMAL) Then                               ' エラー ?
                        strJUG(rn) = JG_IO                                  ' 判定 = "IT-OPEN"
                        strJUG(0) = JG_IO                                   ' 判定(1ﾌﾞﾛｯｸ) = "IT-OPEN"
                        FinalJudgeNG = True
                        DebugLogOut("IT(1)TRIM NG(FinalJudgeNG = True)抵抗[" & rn.ToString & "]")
                        UserSub.NgJudgeSet()                                ' 素子毎のＮＧ判定
                        rtn = cFRS_TRIM_NG                                  ' Return値 = トリミングNG
                        GoTo Trim_DSP
                    Else
                        dblVX(1) = dblMx                                    ' IT測定値
                        Call UserSub.SetInitialResValue(dblMx)              ' 初期測定値の保存
                    End If

                    ' 目標値判定処理(IT)
                    If UserSub.IsSpecialTrimType And IsCutResistor(rn) Then
                        strJUG(rn) = Test_ItFt(0, stREG(rn).intMode, dblMx, UserSub.GetTRV(), stREG(rn).dblITL, stREG(rn).dblITH, Judge)    'V2.0.0.0⑨Judge追加
                    Else
                        strJUG(rn) = Test_ItFt(0, stREG(rn).intMode, dblMx, stREG(rn).dblNOM, stREG(rn).dblITL, stREG(rn).dblITH, Judge)    'V2.0.0.0⑨Judge追加
                    End If
                    If (strJUG(rn) <> JG_OK) Then                           ' IT-NG ?
                        ' IT-NG時でﾌﾟﾛｰﾌﾞﾘﾄﾗｲ指定の場合は一度だけﾘﾄﾗｲする
                        If (stPLT.PrbRetry = 1) And (FlgRetry = 0) Then
                            FlgRetry = 1                                    ' ﾌﾟﾛｰﾌﾞﾘﾄﾗｲ ﾌﾗｸﾞON
                            r = Probe_Retry(rn)                             ' ﾌﾟﾛｰﾌﾞﾘﾄﾗｲ処理
                            GoTo STP_RETRY                                  ' 一度だけﾘﾄﾗｲする
                        End If
                        strJUG(0) = strJUG(rn)                              ' 判定(1ﾌﾞﾛｯｸ)設定
                        FinalJudgeNG = True
                        DebugLogOut("IT(2)TRIM NG(FinalJudgeNG = True)抵抗[" & rn.ToString & "]")
                        UserSub.NgJudgeSet()                                ' 素子毎のＮＧ判定
                        rtn = cFRS_TRIM_NG                                  ' Return値 = トリミングNG
                        GoTo Trim_DSP
                    End If

                    '' '' FT判定を行いFT範囲内ならﾄﾘﾐﾝｸﾞ処理をｽｷｯﾌﾟする
                    ' ''strMSG = Test_ItFt(1, stREG(rn).intMode, dblMx, stREG(rn).dblNOM, stREG(rn).dblFTL, stREG(rn).dblFTH)
                    ' ''If (strMSG = JG_OK) Then
                    ' ''    dblVX(1) = dblMx                                    ' IT測定値
                    ' ''    dblVX(2) = dblMx                                    ' FT測定値
                    ' ''    GoTo Trim_DSP
                    ' ''End If

                End If
                '-----------------------------------------------------------------------
                '   トリミングを行う
                '-----------------------------------------------------------------------
                ' 測定マーキングモードは、ＯＫまたはＮＧマーキングのみカットします。
                ' カット位置補正の「ＮＧ判定あり」が設定されている場合は、補正ＮＧの場合は、ＮＧまたはＯＫのマーキングのみ実施します。
                ' チップ抵抗モードのOK,NGマーキングは、後のMarkingForChipMode()で処理するのでここでは行わない。
                'V2.0.0.0② 測定値変動測定(TRIM_VARIATION_MEAS)追加
                'V2.0.0.0②                If (DGL <> TRIM_MODE_MEAS_MARK And stREG(rn).intSLP < SLP_VMES) Or (Not UserSub.IsTrimType3 And ((stREG(rn).intSLP = SLP_NG_MARK And FinalJudgeNG) Or (stREG(rn).intSLP = SLP_OK_MARK And Not FinalJudgeNG))) Then  ' 5:電圧測定のみ, 6:抵抗測定のみ でない場合にトリミングを行う。
                If ((DGL <> TRIM_VARIATION_MEAS And DGL <> TRIM_MODE_MEAS_MARK And IsCutResistor(rn)) Or _
                    (Not UserSub.IsTrimType3 And Not UserSub.IsTrimType4 And _
                     ((stREG(rn).intSLP = SLP_NG_MARK And FinalJudgeNG) Or (stREG(rn).intSLP = SLP_OK_MARK And Not FinalJudgeNG)))) Then
                    If (stREG(rn).intSLP = SLP_NG_MARK And FinalJudgeNG) Then
                        DebugLogOut("NG MARK TRIM NG(FinalJudgeNG = True)抵抗[" & rn.ToString & "]")
                    End If
                    r = VTrim_One(rn, dblMx)                                ' 1抵抗分トリミングを行う
                    If (r = cFRS_ERR_RST) Then                              ' RESET SW押下 ?
                        strJUG(rn) = JG_RS                                  ' 判定 = "RESET"
                        strJUG(0) = JG_RS                                   ' 判定(1ﾌﾞﾛｯｸ) = "RESET"
                        rtn = cFRS_ERR_RST                                  ' Return値 = RESET SW押下
                        GoTo Trim_EXT
                    End If
                    If (r < cFRS_NORMAL) Then                               ' エラー ?
                        If (r <> cFRS_TRIM_NG) Then                          ' ﾄﾘﾐﾝｸﾞNGなら続行 その他のｱﾌﾟﾘ終了ﾚﾍﾞﾙｴﾗｰなら終了
                            strJUG(rn) = JG_ER                              ' 判定 = "ERROR"
                            strJUG(0) = JG_ER                               ' 判定(1ﾌﾞﾛｯｸ) = "ERROR"
                            rtn = r                                         ' Return値設定(非常停止等)
                            GoTo Trim_EXT
                        End If
                    End If

                End If

                Retry_Cnt = stREG(rn).intReMeas                            ' 再測定の回数

                '-----------------------------------------------------------------------
                '   ﾄﾘﾐﾝｸﾞ後の値を測定する(FT)
                '-----------------------------------------------------------------------
                'V2.0.0.0                If (stREG(rn).intSLP <> SLP_NG_MARK And stREG(rn).intSLP <> SLP_OK_MARK) Then  ' 7:NGﾏｰｷﾝｸﾞ でない場合にFT測定を行う。'V1.0.4.3⑤ ＯＫマーキング(SLP_OK_MARK)追加 'V1.2.0.0③ カット位置補正の判定stREG(rn).bPattern追加
                If IsMeasureMode(rn, MEAS_JUDGE_FT) Then        'V2.0.0.0
FT_MEAS:
                    Call DScanModeSet(rn, 0, 0)                             ' ＤＣスキャナに接続する測定器を切り替えてスキャナをセットする。

                    'V2.0.0.0⑧                    If UserSub.IsTrimType1 And stREG(rn).intMType = 1 And stREG(rn).intSLP = SLP_RMES Then      ' 温度センサーで外部測定器で基準抵抗の場合は５回測定
                    'V2.0.0.0⑧                        MesTime = gGpibMultiMeterCount
                    'V2.0.0.0⑧                    Else
                    'V2.0.0.0⑧                        'MesTime = 1                '20130418
                    'V2.0.0.0⑧                        MesTime = 2                   '20130418
                    'V2.0.0.0⑧                    End If
                    MesTime = stREG(rn).intFTReMeas                                 'V2.0.0.0⑧

                    For i = 1 To MesTime
                        ' 抵抗測定/電圧測定(内部/外部測定器)
                        If UserSub.IsSpecialTrimType And IsCutResistor(rn) Then
                            r = V_R_MEAS(stREG(rn).intSLP, stREG(rn).intMType, dblMx, rn, UserSub.GetTRV())
                        Else
                            r = V_R_MEAS(stREG(rn).intSLP, stREG(rn).intMType, dblMx, rn, stREG(rn).dblNOM)
                        End If
                    Next

                    If (r <> cFRS_NORMAL) Then                               ' エラー ?
                        strJUG(rn) = JG_FO                                  ' 判定 = "FT-OPEN"
                        strJUG(0) = JG_FO                                   ' 判定(1ﾌﾞﾛｯｸ) = "FT-OPEN"
                        FinalJudgeNG = True
                        DebugLogOut("FT(1) TRIM NG(FinalJudgeNG = True)抵抗[" & rn.ToString & "]")
                        UserSub.NgJudgeSet()                                ' 素子毎のＮＧ判定
                        rtn = cFRS_TRIM_NG                                  ' Return値 = トリミングNG
                        GoTo Trim_DSP
                    Else
                        dblVX(2) = dblMx                                    ' FT測定値
                        If UserSub.IsTrimType1() Or UserSub.IsTrimType4() Then  ' 温度センサー'V2.0.0.0①sTrimType4()追加
                            If stREG(rn).intSLP = SLP_RMES Then             ' 抵抗測定のみ
                                bStdJudgeNG = False                         'V2.0.0.0⑮
                                Call UserSub.SetStandardResValue(dblMx)     ' 標準抵抗
                            End If
                        End If
                    End If

                    'V2.0.0.0②↓
                    If DGL = TRIM_VARIATION_MEAS Then   ' 測定値変動測定でＦＴの時、ＦＴの判定より変化量の判定を優先する。
                        If Not UserSub.VariationMeasJudge(rn, dblMx) Then
                            strJUG(rn) = JG_VA
                            Judge = eJudge.JG_VA
                            FinalJudgeNG = True
                            DebugLogOut("TRIM_VARIATION_MEAS TRIM NG(FinalJudgeNG = True)抵抗[" & rn.ToString & "]")
                            UserSub.NgJudgeSet()                                ' 素子毎のＮＧ判定
                            strJUG(0) = strJUG(rn)                              ' 判定(1ﾌﾞﾛｯｸ)設定
                            rtn = cFRS_TRIM_NG                                  ' Return値 = トリミングNG
                            GoTo Trim_DSP
                        End If
                    End If
                    'V2.0.0.0②↑

                    ' 目標値判定処理(FT)
                    If UserSub.IsSpecialTrimType And IsCutResistor(rn) Then
                        strJUG(rn) = Test_ItFt(1, stREG(rn).intMode, dblMx, UserSub.GetTRV(), stREG(rn).dblFTL, stREG(rn).dblFTH, Judge)    'V2.0.0.0⑨Judge追加
                    Else
                        strJUG(rn) = Test_ItFt(1, stREG(rn).intMode, dblMx, stREG(rn).dblNOM, stREG(rn).dblFTL, stREG(rn).dblFTH, Judge)    'V2.0.0.0⑨Judge追加
                    End If
                    If (strJUG(rn) <> JG_OK) Then                           ' FT-NG ?

                        If Retry_Cnt > 0 Then

                            Call Disp_Result(eDispMode.DISP_MODE_REMEAS, 0) '再測定中表示

                            Func_Wait(stREG(rn).intReMeas_Time)
                            Retry_Cnt = Retry_Cnt - 1
                            'V2.0.0.0②ログを出すと再測定で合わなくなる。                            Call Disp_Final(rn)                                 ' ﾄﾘﾐﾝｸﾞ結果表示/ログ出力
                            GoTo FT_MEAS
                        End If
                        FinalJudgeNG = True
                        DebugLogOut("FT(2)TRIM NG (FinalJudgeNG = True)抵抗[" & rn.ToString & "]")
                        UserSub.NgJudgeSet()                                ' 素子毎のＮＧ判定
                        strJUG(0) = strJUG(rn)                              ' 判定(1ﾌﾞﾛｯｸ)設定

                        'V2.0.0.0⑮↓
                        If UserSub.IsTrimType1() Or UserSub.IsTrimType4() Then  ' 温度センサー
                            If stREG(rn).intSLP = SLP_RMES Then                 ' 抵抗測定のみ
                                strJUG(0) = strJUG(rn)                          ' 判定(1ﾌﾞﾛｯｸ)設定
                                bStdJudgeNG = True                              ' 標準抵抗測定NG
                            End If
                        End If
                        'V2.0.0.0⑮↑

                        rtn = cFRS_TRIM_NG                                  ' Return値 = トリミングNG
                        GoTo Trim_DSP
                        'V2.0.0.0⑮↓
                    Else        ' FT結果ＯＫの時
                        If bStdJudgeNG And (UserSub.IsTrimType1() Or UserSub.IsTrimType4()) Then  ' 標準抵抗測定NG
                            strJUG(0) = JG_STD
                            strJUG(rn) = JG_STD
                        End If
                        'V2.0.0.0⑮↑
                    End If

Trim_DSP:
                    'V2.0.0.0②↓
                    If bPowerOnOffUse Then
                        If FUNC_OK = Func_V_Off_Judge(rn) Then                      '   電圧OFF有り？                                               ' DC電源装置 電圧OFF
                            r = Func_V_Off_Ex(rn)                                   '   電圧OFF
                            If (FUNC_NG = r) Then
                                strJUG(rn) = JG_ER                                  ' 判定 = ｴﾗｰ発生(電圧設定等)
                                strJUG(0) = JG_ER                                   ' 判定(1ﾌﾞﾛｯｸ) = ｴﾗｰ発生(電圧設定等)
                                rtn = cFRS_ERR_TRIM                                 ' トリマエラー
                                GoTo Trim_EXT
                            End If
                        End If
                    End If
                    'V2.0.0.0②↑

                    'V2.1.0.0①↓
                    '-----------------------------------------------------------------------
                    '   カット毎の抵抗値変化量判定
                    '-----------------------------------------------------------------------
                    If UserSub.CutVariationFinalJudgeNG() Then                  ' ＮＧの場合
                        strJUG(0) = JG_CUTVA
                        strJUG(rn) = JG_CUTVA
                        UserSub.NgJudgeSet()                                    ' 素子毎のＮＧ判定
                        FinalJudgeNG = True                                     ' この抵抗の判定をNGとする　'V2.2.1.10① 
                    End If
                    'V2.1.0.0①↑
                    '-----------------------------------------------------------------------
                    '   測定値表示処理
                    '-----------------------------------------------------------------------
                    Call Disp_Final(rn)                                         ' ﾄﾘﾐﾝｸﾞ結果表示/ログ出力
                    If IsCutResistor(rn) Then                                   ' トリミング抵抗の時
                        If NetworkSkip Then
                            Call Disp_frmInfo(COUNTER.COUNTUP, COUNTER.SKIP)        ' ﾄﾘﾐﾝｸﾞ結果表示(frmInfo画面 OK/NGｶｳﾝﾄｱｯﾌﾟ)
                        Else
                            Call Disp_frmInfo(COUNTER.COUNTUP, COUNTER.OKNG_UP, rn)     ' ﾄﾘﾐﾝｸﾞ結果表示のみ 'V1.2.0.0② rn　追加
                        End If
                    End If

                    '-----------------------------------------------------------------------
                    '   ネットワークの場合は、ＮＧ後スキップする。
                    '-----------------------------------------------------------------------
                    'V2.0.0.0⑩                    If FinalJudgeNG And UserSub.IsTrimType2() And IsCutResistor(rn) And UserBas.GetRCountExceptMeasure() > 1 Then
                    'V2.0.0.0⑩ 現状は、チップ抵抗のみサーキット処理の対象
                    If FinalJudgeNG And (UserSub.IsTrimType2() Or UserSub.IsTrimType3()) And IsCutResistor(rn) And UserBas.GetRCountExceptMeasure() > 1 Then
                        NetworkSkip = True
                        UserSub.SkipSet()
                    End If

                    JudgeNG = UserSub.SkipGet()         ' FinalJudge()で初期化されるのでここで保持する
                    '-----------------------------------------------------------------------
                    '   ファイナル判定
                    '-----------------------------------------------------------------------
                    If IsCutResistor(rn) Then                                       ' トリミング抵抗の時
                        UserSub.DevCalculation(rn, dblMx)                           ' 偏差計算
                        If UserSub.IsSpecialTrimType() Then                         ' ユーザ指定トリミングモード
                            If Not UserSub.FinalJudge(rn) Then
                                FinalJudgeNG = True
                                DebugLogOut("FinalJudge TRIM NG(FinalJudgeNG = True)抵抗[" & rn.ToString & "]")
                                strJUG(0) = JG_FH
                                rtn = cFRS_TRIM_NG                                  ' Return値 = トリミングNG
                            End If
                        End If
                    End If
                End If

                'V2.0.0.0⑨↓統計値保存
                '相対値判定をFinalJudge()で行っていてstrJUG(rn)にセットするのでFinalJudge()の後で処理する。
                If (Not gObjFrmDistribute Is Nothing And UserModule.IsCutResistor(rn) And strJUG(rn) = JG_OK) Then
                    If UserSub.IsTrimType3() And UserBas.GetRCountExceptMeasure() > 1 Then
                        If UserSub.IsCheckCircuitEnd(rn) Then
                            If Not JudgeNG Then
                                For cno As Short = 1 To UserBas.GetRCountExceptMeasure()
                                    Call gObjFrmDistribute.StatisticalDataSave(FINAL_TEST, cno, gdNOMforStatistical(cno), eJudge.JG_OK)   ' OK,NGをw分けないで全て保存
                                Next
                            End If
                        End If
                    Else
                        Call gObjFrmDistribute.StatisticalDataSave(FINAL_TEST, UserSub.GetResNumberInCircuit(rn), gdNOMforStatistical(GetResNumberInCircuit(rn)), eJudge.JG_OK)   ' OK,NGをw分けないで全て保存
                    End If
                End If
                'V2.0.0.0⑨↑

                If FinalJudgeNG Then        'V2.0.0.0⑩
                    BlockJudgeNG = True     'V2.0.0.0⑩
                End If                      'V2.0.0.0⑩
                'V2.1.0.4①↓
                ' 判定が変化率エラーの場合には、NGマーキングをしたいのでフラグをONする
                If strJUG(rn) = JG_CUTVA Then
                    FinalJudgeNG = True
                End If
                'V2.1.0.4①↑
                'V1.2.0.0②↓
                'V2.0.0.0①                If UserSub.IsTrimType3 And (stREG(rn).intSLP <> SLP_NG_MARK And stREG(rn).intSLP <> SLP_OK_MARK) Then
                If UserModule.IsCutResistor(rn) And ((UserSub.IsTrimType3() And UserSub.IsCheckCircuitEnd(rn)) Or UserSub.IsTrimType4()) Then           'V2.0.0.0① 'V2.0.0.0⑩
                    r = UserSub.MarkingForChipMode(rn, Not FinalJudgeNG)                ' チップ抵抗モードのＯＫまたはＮＧマーキング処理
                    FinalJudgeNG = False                                            ' １抵抗毎に判定なのでリセットする。
                    'V2.0.0.0⑩↓
                    NetworkSkip = False                                             ' サーキットリセット
                    If (UserSub.IsTrimType3() And Not UserSub.IsCheckLastCircuit(rn)) Then
                        For i = 0 To stPLT.RCount                                       ' 抵抗数分繰返す
                            strJUG(i) = JG_SP                                           ' 判定初期化
                        Next i
                        strJUG(0) = JG_OK                                               ' 判定(1ﾌﾞﾛｯｸ) = "OK"
                    End If
                    'V2.0.0.0⑩↑
                End If
                'V1.2.0.0②↑

                If (Form1.System1.Sys_Err_Chk_EX(gSysPrm, APP_MODE_LOTCHG) <> cFRS_NORMAL) Then ' 非常停止等 ?
                    Call Form1.AppEndDataSave()                                       ' ｿﾌﾄ強制終了時のﾃﾞｰﾀ保存確認
                    Call Form1.AplicationForcedEnding()                               ' ｿﾌﾄ強制終了処理
                    End
                End If

Trim_Next:
            Next rn                                                     ' 次抵抗へ
            Call Disp_frmInfo(COUNTER.COUNTUP, COUNTER.SKIP)     ' ﾄﾘﾐﾝｸﾞ結果表示のみ

            '---------------------------------------------------------------------------
            '   終了処理
            '---------------------------------------------------------------------------
Trim_EXT:
            Call DSCAN(Z0, Z0, Z0)                                      ' DCスキャナオフ
            r = V_Off()                                                 ' DC電源装置 電圧OFF
            r = Prob_Off()                                              ' Z2/ZﾌﾟﾛｰﾌﾞをOFF位置(Z.ZOFF)に移動
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
                rtn = r                                                 ' Return値設定
            End If

            If BlockJudgeNG Then                                        'V2.0.0.0⑩
                strJUG(0) = JG_ER                                       'V2.0.0.0⑩    ' 判定(1ﾌﾞﾛｯｸ) = "ERROR"
            End If                                                      'V2.0.0.0⑩

            Return (rtn)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.Trim_() TRAP ERROR = " + ex.Message
            Call Z_PRINT(strMSG & vbCrLf)
            strJUG(0) = JG_ER                                           ' 判定(1ﾌﾞﾛｯｸ) = "ERROR"
            Return (cFRS_TRIM_NG)                                       ' Return値 = トリミングNG
        End Try
    End Function
#End Region

#Region "プローブリトライ処理(オプション)"
    '''=========================================================================
    ''' <summary>プローブリトライ処理(オプション)</summary>
    ''' <param name="rn">(INP) 抵抗番号(1 ORG)</param>
    ''' <returns>cFRS_NORMAL   = 正常
    '''         その他
    ''' </returns>
    '''=========================================================================
    Private Function Probe_Retry(ByRef rn As Short) As Short

        Dim r As Short
        Dim strMSG As String

        Try
            r = Prob_Off()                                              ' Z2/ZﾌﾟﾛｰﾌﾞをOFF位置に移動
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
                Return (r)                                              ' Return値設定
            End If

            r = Prob_On()                                               ' ﾌﾟﾛｰﾌﾞON(ﾎﾟｰｽﾞ付)
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
                Return (r)                                              ' Return値設定
            End If
            Call Disp_frmInfo(COUNTER.PROBE_UP, COUNTER.NONE)              ' ﾌﾟﾛｰﾌﾞON回数ｶｳﾝﾄｱｯﾌﾟ

            ' DCスキャナのプローブ番号指定
            'V1.0.4.3⑨            Call DSCAN(stREG(rn).intPRH, stREG(rn).intPRL, stREG(rn).intPRG)
            Call DSCAN(UserSub.ConvtChannel(stREG(rn).intPRH), UserSub.ConvtChannel(stREG(rn).intPRL), UserSub.ConvtChannel(stREG(rn).intPRG))    'V1.0.4.3⑨ ConvtChannel()追加
            Call System.Threading.Thread.Sleep(10)                      ' Wait(ms)
            Return (cFRS_NORMAL)                                        ' Return値設定

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Probe_Retry() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Return値設定
        End Try
    End Function
#End Region

#Region "測定値判定処理(IT/FT)"
    '''=========================================================================
    '''<summary>測定値判定処理(IT/FT)</summary>
    '''<param name="KD">    (INP) 種別(0:IT, 1:FT)</param>
    '''<param name="MD">    (INP) ﾓｰﾄﾞ(0:%, 1:数値(絶対値))</param>
    '''<param name="dblMx"> (INP) 測定値</param>
    '''<param name="dblNOM">(INP) 目標値</param>
    '''<param name="Lo">    (INP) Lowﾘﾐｯﾄ値 (% or 数値)</param>
    '''<param name="Hi">    (INP) Highﾘﾐｯﾄ値(% or 数値)</param>
    '''<returns>判定文字列("OK   "等)</returns>
    '''=========================================================================
    Public Function Test_ItFt(ByVal KD As Short, ByVal MD As Short, ByVal dblMx As Double, ByVal dblNOM As Double, ByVal Lo As Double, ByVal Hi As Double, ByRef Judge As Integer) As String

        Dim wkLo As Double                                  ' ﾜｰｸ
        Dim wkHi As Double                                  '
        Dim wkECF As Short                                  '
        Dim strMSG As String                                ' メッセージ表示用域
        Dim strRtn As String = ""                           ' メッセージ表示用域

        Try
            ' Lowﾘﾐｯﾄ値/Highﾘﾐｯﾄ値を設定する
            If (MD = 0) Then                                ' ﾓｰﾄﾞ = % ?
                If (dblNOM = 0.0#) Then                     ' 目標値 = 0 ?
                    wkLo = Lo * 0.01                        ' Lowﾘﾐｯﾄ値
                    wkHi = Hi * 0.01                        ' Highﾘﾐｯﾄ値
                Else
                    wkLo = dblNOM + (System.Math.Abs(dblNOM) * Lo * 0.01) ' Lowﾘﾐｯﾄ値  (LOW = (NOM*(100+Lo)/100))
                    wkHi = dblNOM + (System.Math.Abs(dblNOM) * Hi * 0.01) ' Highﾘﾐｯﾄ値 (HIGH= (NOM*(100+Hi)/100))
                End If
            Else                                            ' ﾓｰﾄﾞ = 絶対値
                'wkLo = Lo                                  ' Lowﾘﾐｯﾄ値
                'wkHi = Hi                                  ' Highﾘﾐｯﾄ値
                wkLo = dblNOM + Lo                          ' Lowﾘﾐｯﾄ値         '2013.03.16
                wkHi = dblNOM + Hi                          ' Highﾘﾐｯﾄ値        '2013.03.16
            End If

            Call DebugLogOut(String.Format("{0} 測定値={1} 目標値={2} Lo={3} High={4}", IIf(KD = 0, "IT", "FT"), dblMx, dblNOM, Lo, Hi))

            ' 少数桁5桁に桁数を合わせる
            wkLo = Double.Parse(wkLo.ToString(TARGET_DIGIT_DEFINE))       'V2.0.0.0⑤ "0.00000"からTARGET_DIGIT_DEFINEへ変更
            wkHi = Double.Parse(wkHi.ToString(TARGET_DIGIT_DEFINE))       'V2.0.0.0⑤ "0.00000"からTARGET_DIGIT_DEFINEへ変更
            dblMx = Double.Parse(dblMx.ToString(TARGET_DIGIT_DEFINE))     'V2.0.0.0⑤ "0.00000"からTARGET_DIGIT_DEFINEへ変更

            ' 測定値判定処理(IT/FT)
            If (dblMx < wkLo) Then                          ' 測定値 < Lowﾘﾐｯﾄ値 ?
                wkECF = 2
            ElseIf (dblMx > 100000000.0) Then               ' 測定値 > Highﾘﾐｯﾄ値 OPEN
                wkECF = 4
            ElseIf (dblMx > wkHi) Then                      ' 測定値 > Highﾘﾐｯﾄ値 ?
                wkECF = 3
            Else                                            '  Lowﾘﾐｯﾄ値 <= 測定値 <= Highﾘﾐｯﾄ値
                wkECF = 1
            End If

            ' 判定文字列を設定する
            Select Case (wkECF)
                Case 1 ' OK ?
                    strRtn = JG_OK                          ' Return値 = "OK   "
                    Judge = eJudge.JG_OK
                Case 2 ' 測定値 < Lowﾘﾐｯﾄ値 ?
                    Select Case (KD)
                        Case 0 ' IT
                            strRtn = JG_IL                  ' Return値 = "IT-LO"
                            Judge = eJudge.JG_IL
                        Case 1 ' FT
                            strRtn = JG_FL                  ' Return値 = "FT-LO"
                            Judge = eJudge.JG_FL
                    End Select
                Case 3 ' 測定値 > Highﾘﾐｯﾄ値 ?
                    Select Case (KD)
                        Case 0 ' IT
                            strRtn = JG_IH                  ' Return値 = "IT-HI"
                            Judge = eJudge.JG_IH
                        Case 1 ' FT
                            strRtn = JG_FH                  ' Return値 = "FT-HI"
                            Judge = eJudge.JG_FH
                    End Select
                Case 4  ' OPEN
                    Select Case (KD)
                        Case 0 ' IT
                            strRtn = JG_IO                  ' Return値 = "IT-HO"
                            Judge = eJudge.JG_IO
                        Case 1 ' FT
                            strRtn = JG_FO                  ' Return値 = "FT-HO"
                            Judge = eJudge.JG_FO
                    End Select
            End Select
            Return (strRtn)                                 ' Return値設定

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.Test_ItFt() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (strRtn)                                 ' Return値
        End Try
    End Function
#End Region

#Region "直線カットトリミング処理(ﾄﾘﾐﾝｸﾞﾓｰﾄﾞ)"
    '''=========================================================================
    '''<summary>直線カットトリミング処理(x0ﾓｰﾄﾞ(ﾄﾘﾐﾝｸﾞﾓｰﾄﾞ))</summary>
    '''<param name="rn">(INP) 抵抗番号(1 ORG)</param>
    '''<param name="Mx">(INP) 初期測定値(Lカット時参照)</param>
    '''<returns> 0 = 正常
    '''          3 = RESET SW押下
    '''          1 = 目標値を超えたので終了
    '''          2 = 指定移動量までカットしたので終了
    ''' </returns>
    '''=========================================================================
    Public Function VTrim_One(ByRef rn As Short, ByRef Mx As Double) As Short

        Dim cn As Short                                 ' Index
        Dim r As Short                                  ' 関数戻値
        Dim NOM As Double                               ' 目標値(V)
        Dim NOMx As Double                              ' 目標値（カットオフ）
        Dim dblQrate As Double
        Dim dblQrate2 As Double
        Dim strMSG As String                                            ' メッセージ表示用域
        Dim CutMode As Short                            'V1.0.4.3⑧

        Try
            '---------------------------------------------------------------------------
            '   初期処理
            '---------------------------------------------------------------------------
            VTrim_One = cFRS_NORMAL                         ' Return値 = 正常
            NOM = stREG(rn).dblNOM                          ' 目標値
            'V2.0.0.0⑱↓インデックスでも参照しているが念のためここでも設定する。
            If UserSub.IsSpecialTrimType Then
                NOM = UserSub.GetTRV()                      ' 目標値
            End If
            'V2.0.0.0⑱↑

            ' 最初のカット位置へBPを移動する(カット位置補正あり)
            Call STRXY(rn, stREG(rn).STCUT(1).dblSTX, stREG(rn).STCUT(1).dblSTY)

            '' ADV/HALT/RESET待ち(ADJ ON時) ※抵抗毎に停止
            'r = Form1.System1.HALT2(3)                         ' ADV(1)/HALT(2)/RESET(3)待ち
            'If (r = cFRS_ERR_RST) Then                      ' RESET SW押下 ?
            '    VTrim_One = cFRS_ERR_RST                    ' Return値 = RESET SW押下
            '    Exit Function
            'End If
            'If (r < cFRS_NORMAL) Then                       ' エラー ?
            '    VTrim_One = r                               ' Return値設定
            '    Exit Function
            'End If

            UserSub.CutVariationInitialize(rn)           'V2.1.0.0①判定用目標値算出係数をカット前の測定値として保存

            '---------------------------------------------------------------------------
            '   1抵抗分電圧トリミングまたは抵抗トリミングを行う
            '---------------------------------------------------------------------------
            For cn = 1 To stREG(rn).intTNN                  ' カット数分繰返す
                UserSub.CutVariationInitByCut()             'V2.1.0.0① 抵抗値変化量判定・カット後の変化量未計算状態にする。
                InitCutParam(cutCmnPrm)
                dblQrate = stREG(rn).STCUT(cn).intQF1
                dblQrate = dblQrate / 10.0
                dblQrate2 = stREG(rn).STCUT(cn).intQF2
                dblQrate2 = dblQrate2 / 10.0

                ' Q-RATE設定/BP移動(カット位置補正あり)
                'Call QRATE(stREG(rn).STCUT(cn).intQF1)      ' Q-RATE設定
                'Call QRATE(dblQrate)      ' Q-RATE設定
                If (stREG(rn).STCUT(cn).intCUT <> CNS_CUTM_NON_POS_IX) Then   ' ｶｯﾄ方法 = ﾎﾟｼﾞｼｮﾆﾝｸﾞ無しｲﾝﾃﾞｯｸｽ以外 ?
                    Call STRXY(rn, stREG(rn).STCUT(cn).dblSTX, stREG(rn).STCUT(cn).dblSTY)
                End If

                MoveStop()              'V2.2.0.0⑥ 
                '-----------------------------------------------------------------------
                '   ＮＧマーク
                '-----------------------------------------------------------------------
                If (stREG(rn).intSLP = SLP_NG_MARK Or stREG(rn).intSLP = SLP_OK_MARK) Then                                              'V1.0.4.3⑤ ＯＫマーキング(SLP_OK_MARK)追加
                    NgCutDebugLogOut("抵抗[" & rn.ToString & "]カット[" & cn.ToString & "]種別[" & stREG(rn).intSLP.ToString & "] X=[" & stREG(rn).STCUT(cn).dblSTX.ToString & "] Y=[" & stREG(rn).STCUT(cn).dblSTY.ToString & "] カット長 = [" & stREG(rn).STCUT(cn).dblDL2.ToString & "]")    'V1.2.0.2ＮＧカットのログ
                    '###1042① カット形状
                    Select Case (stREG(rn).STCUT(cn).intCTYP)                                                                           '###1042①
                        Case CNS_CUTP_ST, CNS_CUTP_ST_TR                                                                                                '###1042① STカット
                            '                    r = TrimSt(FORCE_MODE, 0, 0, stREG(rn).intSLP, stREG(rn).STCUT(cn).intANG, stREG(rn).STCUT(cn).dblDL2, stREG(rn).STCUT(cn).dblV1, stREG(rn).STCUT(cn).dblV1, dblQrate, dblQrate, 0, 0)
                            r = TrimSt(FORCE_MODE, 0, 0, SLP_RTRM, stREG(rn).STCUT(cn).intANG, stREG(rn).STCUT(cn).dblDL2, stREG(rn).STCUT(cn).dblV1, stREG(rn).STCUT(cn).dblV1, dblQrate, dblQrate, 0, 0)

                            If stREG(rn).STCUT(cn).intCTYP = CNS_CUTP_ST_TR Then        'V2.0.0.0⑦ リトレースカット
                                r = CUT_RETRACE(rn, cn, stREG(rn).STCUT(cn).dblDL2)     'V2.0.0.0⑦
                            End If                                                      'V2.0.0.0⑦

                        Case CNS_CUTP_L                                                                                                 '###1042① Lカット
                            r = TRM_L6(FORCE_MODE, rn, cn, stREG(rn).dblNOM)        'V1.0.4.3⑦ 斜めLｶｯﾄ電圧/抵抗ﾄﾘﾐﾝｸﾞ
                        Case CNS_CUTP_M                                                                                                 '###1042① 文字マーキング
                            r = TrimMK(stREG(rn).STCUT(cn).cFormat, _
                                    stREG(rn).STCUT(cn).dblSTX, stREG(rn).STCUT(cn).dblSTY, _
                                        stREG(rn).STCUT(cn).dblDL2, _
                                        stREG(rn).STCUT(cn).dblV1, stREG(rn).STCUT(cn).intANG, dblQrate, 0, 2)        '###1042①
                        Case Else                                                                                                       '###1042①
                            '                    r = TrimSt(FORCE_MODE, 0, 0, stREG(rn).intSLP, stREG(rn).STCUT(cn).intANG, stREG(rn).STCUT(cn).dblDL2, stREG(rn).STCUT(cn).dblV1, stREG(rn).STCUT(cn).dblV1, dblQrate, dblQrate, 0, 0)
                            r = TrimSt(FORCE_MODE, 0, 0, SLP_RTRM, stREG(rn).STCUT(cn).intANG, stREG(rn).STCUT(cn).dblDL2, stREG(rn).STCUT(cn).dblV1, stREG(rn).STCUT(cn).dblV1, dblQrate, dblQrate, 0, 0)
                    End Select                                                                                                          '###1042①
                    r = cFRS_NORMAL     ' ＮＧマークはリターンを見ない

                    'V2.2.0.0⑯ ↓
                    ' 複数抵抗値取得の場合のOKマーキングでは、カット数は５登録されているが、登録番号までのカットまでしか行わない 
                    'V2.2.0.0⑯↓
                    If (stMultiBlock.gMultiBlock <> 0) AndAlso (stREG(rn).intSLP = SLP_OK_MARK) Then
                        'OKマーキングとしての連番を調べ、マルチブロック番号以降は処理しない。
                        'Dim OkMarkNo As Integer = UserModule.GetOkMarkingResNo(rn)
                        If cn >= stExecBlkData.DataNo Then
                            Exit For
                        End If
                    End If
                    'V2.2.0.0⑯ ↑


                    '-----------------------------------------------------------------------
                    '   ｶｯﾄ方法 = トラッキングトリミング時(内部測定器のみ)
                    '-----------------------------------------------------------------------
                ElseIf (stREG(rn).STCUT(cn).intCUT = CNS_CUTM_TR) Then    ' ｶｯﾄ方法 = ﾄﾗｯｷﾝｸﾞ ?

                    Call DScanModeSet(rn, 0, 0)                             ' DCスキャナに接続する測定器を切替る 

                    ' カット形状
                    Select Case (stREG(rn).STCUT(cn).intCTYP)
                        Case CNS_CUTP_ST, CNS_CUTP_ST_TR                    ' STカット(斜め直線カット電圧/抵抗トリミング)
                            'V1.0.4.3⑧↓
                            If stREG(rn).STCUT(cn).intCTYP = CNS_CUTP_ST Then
                                CutMode = CNS_CUTP_NORMAL
                            Else
                                CutMode = CNS_CUTP_ST_TR
                            End If
                            NOMx = Func_CalNomForCutOff(rn, cn, NOM)         'V1.1.0.1③カットオフによる目標値
                            'V1.0.4.3⑧↑
                            r = TrimSt(TRIM_MODE, CutMode, NOMx, stREG(rn).intSLP, stREG(rn).STCUT(cn).intANG, stREG(rn).STCUT(cn).dblDL2, stREG(rn).STCUT(cn).dblV1, stREG(rn).STCUT(cn).dblV2, dblQrate, dblQrate2, 0, 0, stREG(rn).STCUT(cn).dblSX2, stREG(rn).STCUT(cn).dblSY2)

                            If stREG(rn).STCUT(cn).intCTYP = CNS_CUTP_ST_TR Then        'V2.0.0.0⑦ ストレート・リトレースカット本数１０本化
                                Dim Len As Double                                       'V2.0.0.0⑦
                                GET_CUT_LENGTH(Len)                                     'V2.0.0.0⑦
                                r = CUT_RETRACE(rn, cn, Len)                            'V2.0.0.0⑦
                            End If                                                      'V2.0.0.0⑦

                        Case CNS_CUTP_L                     ' Lカット(斜めＬカット電圧/抵抗トリミング)
                            'V1.1.0.1③                            NOMx = NOM * System.Math.Abs(1 + (stREG(rn).STCUT(cn).dblCOF * 0.01))
                            NOMx = Func_CalNomForCutOff(rn, cn, NOM)         'V1.1.0.1③カットオフによる目標値
                            r = TRM_L6(TRIM_MODE, rn, cn, NOMx)        ' 斜めLｶｯﾄ電圧/抵抗ﾄﾘﾐﾝｸﾞ

                        Case CNS_CUTP_SP ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ
                            r = TRM_SPT(rn, cn, NOM)        ' SPｶｯﾄ電圧ﾄﾘﾐﾝｸﾞ(斜め直線ｶｯﾄ電圧/抵抗ﾄﾘﾐﾝｸﾞ)

                        'V2.2.0.0② ↓
                        Case CNS_CUTP_U
                            r = TRM_VH(rn, cn, NOM)

                            'V2.2.0.0② ↑

                        Case Else
                            GoTo STP_ERR                    ' STカット/Lカット以外はエラー
                    End Select

                    '-----------------------------------------------------------------------
                    '   ｶｯﾄ方法 = ｲﾝﾃﾞｯｸｽﾄﾘﾐﾝｸﾞ, (ﾎﾟｼﾞｼｮﾆﾝｸﾞ無しｲﾝﾃﾞｯｸｽﾄﾘﾐﾝｸﾞ時(ｵﾌﾟｼｮﾝ))
                    '-----------------------------------------------------------------------
                ElseIf (stREG(rn).STCUT(cn).intCUT = CNS_CUTM_IX) Then

                    ' カット形状
                    Select Case (stREG(rn).STCUT(cn).intCTYP)
                        Case CNS_CUTP_ST, CNS_CUTP_L, CNS_CUTP_ST_TR         ' STカット/Lカット
                            r = TRM_IX(rn, cn, Mx)          ' 斜めｲﾝﾃﾞｯｸｽｶｯﾄﾄﾘﾐﾝｸﾞ
                            If r = 5 Then
                                r = cFRS_NORMAL
                                Exit For
                            End If

                        Case CNS_CUTP_U             'Uカット   V2.2.0.0② 
                            r = TRM_IX_U(rn, cn, Mx)


                        Case CNS_CUTP_SP ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ
                            r = TRM_IX_SPT(rn, cn, NOM)     ' 斜めｲﾝﾃﾞｯｸｽｶｯﾄﾄﾘﾐﾝｸﾞ

                        Case Else
                            GoTo STP_ERR                    ' STカット/Lカット以外はエラー
                    End Select
                End If

                ' 終了判定
                If Not UserSub.IsSpecialTrimType Then
                    If (r = 1) Then Exit For '                  ' ﾄﾘﾐﾝｸﾞ終了(目標値を超えた) ?
                End If

                If (r < cFRS_NORMAL) Then Exit For '        ' その他のｴﾗｰ ?

                'V2.1.0.0①↓カット毎の抵抗値変化量判定
                If UserSub.CutVariationJudge(rn, cn) = False Then
                    r = cFRS_NORMAL
                    Exit For
                End If
                'V2.1.0.0①↑

                MoveStop()              'V2.2.0.0⑥ 'V2.2.0.027

            Next cn                                         ' 次カットへ

            '---------------------------------------------------------------------------
            '   終了処理
            '---------------------------------------------------------------------------

            'V2.1.0.0①↓カット毎の抵抗値変化量判定、ループ内で判定されなかった場合を考慮、判定済みはスキップされる。
            UserSub.CutVariationJudge(rn, cn)
            'V2.1.0.0①↑

            ' エラーならトリミングNG
            If (r = 99) Then r = cFRS_TRIM_NG
            Return (r)                                                  ' Return値設定

STP_ERR:
            strMSG = "VTrim_One() カット形状エラー"
            Call Z_PRINT(strMSG & vbCrLf)     ' ログ画面に表示
            Return (cFRS_TRIM_NG)                                       ' Return値 = トリミングNG

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.STPRP() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cFRS_TRIM_NG)                                       ' Return値 = トリミングNG
        End Try
    End Function
#End Region

#Region "Lﾀｰﾝ後のｶｯﾄ方向(1:時計方向,2:反時計方向)を返す"
    '''=========================================================================
    '''<summary>Lﾀｰﾝ後のｶｯﾄ方向(1:時計方向,2:反時計方向)を返す</summary>
    '''<param name="ANG1">(INP) Lﾀｰﾝ前のｶｯﾄ方向(90°単位の0～360°)</param>
    '''<param name="ANG2">(INP) Lﾀｰﾝ後移動方向 (90°単位の0～360°)</param>
    '''<returns>Lﾀｰﾝ後のｶｯﾄ方向(1:時計方向,2:反時計方向)</returns>
    '''=========================================================================
    Private Function Get_Cut_Dir(ByRef ANG1 As Short, ByRef ANG2 As Short) As Short

        Dim r As Short

        r = 1                                           ' Return値 = 1(時計方向)
        Select Case (ANG1)
            Case 0 ' →
                Select Case (ANG2)
                    Case 90 : r = 2                     ' →↑
                    Case 270 : r = 1                    ' →↓
                End Select
            Case 90 ' ↑
                Select Case (ANG2)
                    Case 0 : r = 1                      ' ↑→
                    Case 180 : r = 2                    ' ←↑
                End Select
            Case 180 ' ←
                Select Case (ANG2)
                    Case 90 : r = 1                     ' ↑←
                    Case 270 : r = 2                    ' ↓←
                End Select
            Case 270 ' ↓
                Select Case (ANG2)
                    Case 180 : r = 1                    ' ←↓
                    Case 0 : r = 2                      ' ↓→
                End Select
        End Select

        Get_Cut_Dir = r                                 ' Return値 = Lﾀｰﾝ後のｶｯﾄ方向(1:時計方向,2:反時計方向)

    End Function
#End Region

#Region "斜めＬカット電圧/抵抗トリミング"
    '''=========================================================================
    '''<summary>斜めＬカット電圧/抵抗トリミング</summary>
    '''<param name="rn"> (INP) 抵抗データIndex　(1 org)</param>
    '''<param name="cn"> (INP) カットデータIndex(1 org)</param>
    '''<param name="NOM">(INP) 目標値</param>
    '''<returns> 0 = 正常
    '''          1 = 目標値を超えたので終了
    '''          2 = Lターン後の指定移動量までカットしたので終了
    '''         99 = その他エラー
    ''' </returns>
    '''=========================================================================
    Public Function TRM_VL(ByVal MoveMode As Short, ByRef rn As Short, ByRef cn As Short, ByRef NOM As Double) As Integer

        Dim iDir As Short                               ' Lﾀｰﾝ後移動方向(1:時計方向,2:反時計方向)
        Dim CutLen(2) As Double
        Dim SpdOwd(2) As Double
        Dim SpdRet(2) As Double
        Dim QRateOwd(2) As Double
        Dim QRateRet(2) As Double
        Dim CondOwd(2) As Short
        Dim CondRet(2) As Short
        Dim dblQrate As Double


        If stREG(rn).intSLP <> SLP_VTRIMPLS And stREG(rn).intSLP <> SLP_VTRIMMNS And stREG(rn).intSLP <> SLP_RTRM Then
            TRM_VL = 0
            Exit Function
        End If

        CutLen(0) = stREG(rn).STCUT(cn).dblDL2
        CutLen(1) = stREG(rn).STCUT(cn).dblDL3

        SpdOwd(0) = stREG(rn).STCUT(cn).dblV1
        SpdOwd(1) = stREG(rn).STCUT(cn).dblV1
        SpdRet(0) = stREG(rn).STCUT(cn).dblV1
        SpdRet(1) = stREG(rn).STCUT(cn).dblV1

        dblQrate = stREG(rn).STCUT(cn).intQF1
        dblQrate = dblQrate / 10.0

        QRateOwd(0) = dblQrate
        QRateOwd(1) = dblQrate
        QRateRet(0) = dblQrate
        QRateRet(1) = dblQrate

        CondOwd(0) = stREG(rn).STCUT(cn).intCND(1)
        CondOwd(1) = stREG(rn).STCUT(cn).intCND(2)

        CondRet(0) = stREG(rn).STCUT(cn).intCND(3)
        CondRet(1) = stREG(rn).STCUT(cn).intCND(4)

        ' Lﾀｰﾝ後移動方向(1:時計方向,2:反時計方向)を求める
        iDir = Get_Cut_Dir(stREG(rn).STCUT(cn).intANG, stREG(rn).STCUT(cn).intANG2)

        TRM_VL = TrimL(MoveMode, 0, NOM, stREG(rn).intSLP, stREG(rn).STCUT(cn).intTMM, stREG(rn).STCUT(cn).intANG, stREG(rn).STCUT(cn).dblLTP, iDir, _
                    CutLen, SpdOwd, SpdRet, QRateOwd, QRateRet, CondOwd, CondRet)


    End Function
#End Region

#Region "６点ターンポイントＬカット電圧/抵抗トリミング"
    ''' <summary>
    ''' ６点ターンポイントＬカット電圧/抵抗トリミング
    ''' </summary>
    ''' <param name="MoveMode">動作モード(０：トリム　１：ティーチング　２：強制カット)</param>
    ''' <param name="rn">(INP) 抵抗データIndex　(1 org)idou</param>
    ''' <param name="cn">(INP) カットデータIndex(1 org)</param>
    ''' <param name="NOM">(INP) 目標値</param>
    ''' <returns> 0 = 正常
    '''            1 = 目標値を超えたので終了
    '''            2 = Lターン後の指定移動量までカットしたので終了
    '''           99 = その他エラー
    ''' </returns>
    ''' <remarks></remarks>
    Public Function TRM_L6(ByVal MoveMode As Short, ByRef rn As Short, ByRef cn As Short, ByRef NOM As Double) As Integer

        Try
            Dim rslt As Integer
            Dim cutCmnPrm As CUT_COMMON_PRM_L6

            If stREG(rn).intSLP <> SLP_VTRIMPLS And stREG(rn).intSLP <> SLP_VTRIMMNS And stREG(rn).intSLP <> SLP_RTRM Then
                TRM_L6 = cFRS_NORMAL
                Exit Function
            End If

            'rslt = TRM_VL(MoveMode, rn, cn, NOM)

            cutCmnPrm.CutInfo.srtMoveMode = MoveMode                    ' 動作モード(0:トリミング、1:ティーチング、2:強制カット)
            cutCmnPrm.CutInfo.srtCutMode = 0                            ' カットモード(0:ノーマル、1:リターン、2:リトレース、4:斜め)
            cutCmnPrm.CutInfo.dblTarget = NOM                           ' 目標値(カット時は0を設定)
            cutCmnPrm.CutInfo.srtSlope = stREG(rn).intSLP               ' スロープ設定(1:電圧測定＋スロープ、2:電圧測定－スロープ、4:抵抗測定＋スロープ、5:抵抗測定－スロープ)
            cutCmnPrm.CutInfo.srtMeasType = stREG(rn).STCUT(cn).intTMM  ' 測定タイプ(0:高速(3回)、1:高精度(2000回)、2:（IDXのみ）外部機器、3:測定無し、5～:指定回数測定）

            cutCmnPrm.CutCond.dCutLen_1 = stREG(rn).STCUT(cn).dCutLen(1)
            cutCmnPrm.CutCond.dCutLen_2 = stREG(rn).STCUT(cn).dCutLen(2)
            cutCmnPrm.CutCond.dCutLen_3 = stREG(rn).STCUT(cn).dCutLen(3)
            cutCmnPrm.CutCond.dCutLen_4 = stREG(rn).STCUT(cn).dCutLen(4)
            cutCmnPrm.CutCond.dCutLen_5 = stREG(rn).STCUT(cn).dCutLen(5)
            cutCmnPrm.CutCond.dCutLen_6 = stREG(rn).STCUT(cn).dCutLen(6)
            cutCmnPrm.CutCond.dCutLen_7 = stREG(rn).STCUT(cn).dCutLen(7)

            cutCmnPrm.CutCond.dQRate_1 = stREG(rn).STCUT(cn).dQRate(1) / 10.0
            cutCmnPrm.CutCond.dQRate_2 = stREG(rn).STCUT(cn).dQRate(2) / 10.0
            cutCmnPrm.CutCond.dQRate_3 = stREG(rn).STCUT(cn).dQRate(3) / 10.0
            cutCmnPrm.CutCond.dQRate_4 = stREG(rn).STCUT(cn).dQRate(4) / 10.0
            cutCmnPrm.CutCond.dQRate_5 = stREG(rn).STCUT(cn).dQRate(5) / 10.0
            cutCmnPrm.CutCond.dQRate_6 = stREG(rn).STCUT(cn).dQRate(6) / 10.0
            cutCmnPrm.CutCond.dQRate_7 = stREG(rn).STCUT(cn).dQRate(7) / 10.0

            cutCmnPrm.CutCond.dSpeed_1 = stREG(rn).STCUT(cn).dSpeed(1)
            cutCmnPrm.CutCond.dSpeed_2 = stREG(rn).STCUT(cn).dSpeed(2)
            cutCmnPrm.CutCond.dSpeed_3 = stREG(rn).STCUT(cn).dSpeed(3)
            cutCmnPrm.CutCond.dSpeed_4 = stREG(rn).STCUT(cn).dSpeed(4)
            cutCmnPrm.CutCond.dSpeed_5 = stREG(rn).STCUT(cn).dSpeed(5)
            cutCmnPrm.CutCond.dSpeed_6 = stREG(rn).STCUT(cn).dSpeed(6)
            cutCmnPrm.CutCond.dSpeed_7 = stREG(rn).STCUT(cn).dSpeed(7)

            cutCmnPrm.CutCond.dAngle_1 = stREG(rn).STCUT(cn).dAngle(1)
            cutCmnPrm.CutCond.dAngle_2 = stREG(rn).STCUT(cn).dAngle(2)
            cutCmnPrm.CutCond.dAngle_3 = stREG(rn).STCUT(cn).dAngle(3)
            cutCmnPrm.CutCond.dAngle_4 = stREG(rn).STCUT(cn).dAngle(4)
            cutCmnPrm.CutCond.dAngle_5 = stREG(rn).STCUT(cn).dAngle(5)
            cutCmnPrm.CutCond.dAngle_6 = stREG(rn).STCUT(cn).dAngle(6)
            cutCmnPrm.CutCond.dAngle_7 = stREG(rn).STCUT(cn).dAngle(7)

            cutCmnPrm.CutCond.dTurnPoint_1 = stREG(rn).STCUT(cn).dTurnPoint(1)
            cutCmnPrm.CutCond.dTurnPoint_2 = stREG(rn).STCUT(cn).dTurnPoint(2)
            cutCmnPrm.CutCond.dTurnPoint_3 = stREG(rn).STCUT(cn).dTurnPoint(3)
            cutCmnPrm.CutCond.dTurnPoint_4 = stREG(rn).STCUT(cn).dTurnPoint(4)
            cutCmnPrm.CutCond.dTurnPoint_5 = stREG(rn).STCUT(cn).dTurnPoint(5)
            cutCmnPrm.CutCond.dTurnPoint_6 = stREG(rn).STCUT(cn).dTurnPoint(6)

            'cutCmnPrm.CutInfo.srtMoveMode = FORCE_MODE                    ' 動作モード(0:トリミング、1:ティーチング、2:強制カット)
            'rslt = TRIM_L6(cutCmnPrm)
            'cutCmnPrm.CutInfo.srtMoveMode = MoveMode                    ' 動作モード(0:トリミング、1:ティーチング、2:強制カット)
            rslt = TRIM_L6(cutCmnPrm)
            Return (rslt)

        Catch ex As Exception
            MsgBox("User.TRM_L6() TRAP ERROR = " + ex.Message)
            Return (cERR_TRAP)                                       ' トラップエラー発生
        End Try
    End Function
#End Region

#Region "斜めフックカット電圧トリミング"
    '''=========================================================================
    '''<summary>斜めフックカット電圧トリミング</summary>
    '''<param name="rn"> (INP) 抵抗データIndex　(1 org)</param>
    '''<param name="cn"> (INP) カットデータIndex(1 org)</param>
    '''<returns> 0 = 正常
    '''          1 = 目標値を超えたので終了
    '''          2 = Lターン後の指定移動量までカットしたので終了
    '''         99 = その他エラー
    ''' </returns>
    '''=========================================================================
    Public Function TRM_VH(ByRef rn As Short, ByRef cn As Short, ByRef NOM As Double) As Short

        '    ' 斜めﾌｯｸｶｯﾄ電圧ﾄﾘﾐﾝｸﾞ(SLP,目標値,モード,ﾀｰﾝﾎﾟｲﾝﾄ,Lﾀｰﾝ後移動方向,移動量1,移動量2,移動量3,速度,開始角度)
        '    TRM_VH = VHTRIM2(stREG(i).intSLP, NOM, stREG(rn).STCUT.intTMM(j), stREG(rn).STCUT.dblLTP(j), stREG(rn).STCUT.intDIR(j), _
        ''                     gdblDL1(j, 1), stREG(rn).STCUT.dblDL2(j), stREG(rn).STCUT.dblDL3(j), stREG(rn).STCUT.dblV1(j), stREG(rn).STCUT.intANG(j))
        '斜めﾌｯｸｶｯﾄ電圧ﾄﾘﾐﾝｸﾞ(SLP, 目標値, モード, ﾀｰﾝﾎﾟｲﾝﾄ, Lﾀｰﾝ後移動方向, 移動量1, 移動量2, 移動量3, 速度, 開始角度)
        'TRM_VH = VUTRIM2(stREG(rn).intSLP, NOM, stREG(rn).STCUT(cn).intTMM, stREG(rn).STCUT(cn).dblUCutTurnP, stREG(rn).STCUT(cn).intUCutTurnDir,
        'stREG(rn).STCUT(cn).dUCutL1, stREG(rn).STCUT(cn).dUCutL2, 0.0, stREG(rn).STCUT(cn).dblUCutV1, stREG(rn).STCUT(cn).intUCutANG)

        TRM_VH = VUTRIM3(stREG(rn).intSLP, NOM, stREG(rn).STCUT(cn).intTMM, stREG(rn).STCUT(cn).dblUCutTurnP, stREG(rn).STCUT(cn).intUCutTurnDir,
                             stREG(rn).STCUT(cn).dUCutL1, stREG(rn).STCUT(cn).dUCutL2, stREG(rn).STCUT(cn).dblUCutR1, stREG(rn).STCUT(cn).dblUCutR2, stREG(rn).STCUT(cn).dblUCutV1, stREG(rn).STCUT(cn).intUCutANG)

    End Function
#End Region

#Region "サーペンタインカット トリミング x0(ﾄﾘﾐﾝｸﾞﾓｰﾄﾞ)"
    '''=========================================================================
    '''<summary>サーペンタインカット トリミング x0(ﾄﾘﾐﾝｸﾞﾓｰﾄﾞ)</summary>
    '''<param name="rn"> (INP) 抵抗データIndex　(1 org)</param>
    '''<param name="cn"> (INP) カットデータIndex(1 org)</param>
    '''<param name="NOM">(INP) 目標値</param>
    '''<returns> 0 = 正常
    '''          1 = 正常(目標値を超えたので終了)
    '''          3 = RESET SW押下
    ''' </returns>
    '''=========================================================================
    Private Function TRM_SPT(ByRef rn As Short, ByRef cn As Short, ByRef NOM As Double) As Short

        Dim r As Short                                  ' 戻り値
        Dim i As Short                                  ' Index
        Dim NOMx As Double                              ' 目標値（カットオフ）
        'Dim SLP As Short                                ' スロープ
        Dim ANG As Short                                ' ｶｯﾄ方向(90°単位　0°～360°)
        Dim stSPC As Sp_Cut_Info                        ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ情報
        Dim strMSG As String                            ' メッセージ編集域
        Dim dblQrate As Double

        dblQrate = stREG(rn).STCUT(cn).intQF1
        dblQrate = dblQrate / 10.0

        ' 初期処理
        TRM_SPT = cFRS_NORMAL                           ' Return値 = 正常
        strMSG = ""
        stSPC.dblSTX = New Double(MAXSCTN) {}           ' Sp_Cut_Info構造体初期化 
        stSPC.dblSTY = New Double(MAXSCTN) {}

        Call Set_SpCut_Info(rn, cn, stSPC)              ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ情報設定
        'SLP = stREG(rn).intSLP                          ' 電圧変化スロープ = 1(＋), 2(－)
        ANG = stREG(rn).STCUT(cn).intANG                ' カット方向
        ' ｶｯﾄｵﾌ(%)→目標値に対するｵﾌｾｯﾄ値(目標値×(1＋ｶｯﾄｵﾌ/100))
        NOMx = NOM * System.Math.Abs(1 + (stREG(rn).STCUT(cn).dblCOF * 0.01))

#If (cCND = 1) Then                                     ' 条件出しﾓｰﾄﾞ ?
        strMSG = "抵抗番号=" + Format(rn, "0") + ",ｶｯﾄ番号=" + Format(cn, "0") + ",目標値(V)=" + Format(NOM, "0.0####")
        Call Z_PRINT(strMSG + vbCrLf)
#End If
        ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ実行(STｶｯﾄをｶｯﾄ本数分実行)
        For i = 1 To stREG(rn).STCUT(cn).intNum         ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ本数分繰返す

            ' セーフティチェック
            r = SafetyCheck()                           ' セーフティチェック
            If (r <> cFRS_NORMAL) Then                  ' エラー ?
                TRM_SPT = r                             ' Return値を設定する
                Exit Function
            End If

#If (cCND = 1) Then                                     ' 条件出しﾓｰﾄﾞ ?
            strMSG = "ｶｯﾄ長=" + Format(stREG(rn).STCUT(cn).dblDL2, "#0.0####") + ",目標値(V)=" + Format(NOMx, "0.0####") + ",ｶｯﾄ本数=" + Format(i, "0") + "/" + Format(stREG(rn).STCUT(cn).intNum, "0")
            Call Z_PRINT(strMSG + vbCrLf)
#End If
            ' BP絶対値移動(ｶｯﾄ位置 + PTN補正値)(絶対値)
            r = ObjSys.EX_MOVE(gSysPrm, stSPC.dblSTX(i) + stPTN(rn).dblDRX, stSPC.dblSTY(i) + stPTN(rn).dblDRY, 1)
            If (r <> 0) Then                            ' ｴﾗｰ ?
                TRM_SPT = cFRS_ERR_TRIM                 ' Return値 =トリマエラー
                Exit Function
            End If

            MoveStop()              'V2.2.0.0⑥ 

            ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ実行(斜め直線ｶｯﾄ電圧ﾄﾘﾐﾝｸﾞ)
            'r = VTRIM2(SLP, NOMx, stREG(rn).STCUT(cn).dblDL2, stREG(rn).STCUT(cn).dblV1, ANG)
            r = TrimSt(TRIM_MODE, 0, NOMx, stREG(rn).intSLP, ANG, stREG(rn).STCUT(cn).dblDL2, stREG(rn).STCUT(cn).dblV1, stREG(rn).STCUT(cn).dblV1, dblQrate, dblQrate, 0, 0)

            ' 終了判定を行う
            If (r = 99) Then                            ' エラー ?(1:目標値を超えた,2:指定移動量までカット,99:エラー)
                TRM_SPT = cFRS_TRIM_NG                  ' Return値 =トリミングNG
                Exit Function
            ElseIf (r = 1) Then                         ' 目標値を超えたので終了 ?
                TRM_SPT = 1                             ' Return値 = 1(目標値を超えたので終了)
                Exit For
            End If

            Call Cnv_ANG(ANG)                           ' カット方向を反対方向に変換
        Next i                                          ' 次カットへ

    End Function
#End Region

#Region "ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ(ｲﾝﾃﾞｯｸｽｶｯﾄﾄﾘﾐﾝｸﾞ) x0(ﾄﾘﾐﾝｸﾞﾓｰﾄﾞ)"
    '''=========================================================================
    '''<summary>ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ(ｲﾝﾃﾞｯｸｽｶｯﾄﾄﾘﾐﾝｸﾞ) x0(ﾄﾘﾐﾝｸﾞﾓｰﾄﾞ)</summary>
    '''<param name="rn"> (INP) 抵抗データIndex　(1 org)</param>
    '''<param name="cn"> (INP) カットデータIndex(1 org)</param>
    '''<param name="NOM">(INP) 目標値</param>
    '''<returns> 0 = 正常
    '''          1 = 正常(目標値を超えたので終了)
    '''          3 = RESET SW押下
    '''          上記以外その他エラー
    ''' </returns>
    '''=========================================================================
    Private Function TRM_IX_SPT(ByRef rn As Short, ByRef cn As Short, ByRef NOM As Double) As Short

        Dim r As Short                                  ' 戻り値
        Dim i As Short                                  ' Index
        Dim j As Short                                  ' Index
        Dim IDX As Short                                ' ｲﾝﾃﾞｯｸｽｶｯﾄ1～5(ﾋﾟｯﾁ大,中,小)
        Dim count As Short                              ' ｲﾝﾃﾞｯｸｽｶｯﾄ数
        Dim ln As Double                                ' 現在のｶｯﾄ長
        Dim dblLN As Double                             ' ｶｯﾄ量
        Dim CutL As Double                              ' 最大ｶｯﾄ長
        Dim NOMx As Double                              ' 目標値（カットオフ）
        Dim ANG As Short                                ' ｶｯﾄ方向(90°単位　0°～360°)
        Dim stSPC As Sp_Cut_Info                        ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ情報
        Dim strMSG As String                            ' メッセージ編集域
        Dim wkL1 As Double                              ' 作業域
        Dim wkL2 As Double                              ' 作業域
        Dim dblMx As Double                             ' 作業域

        If stREG(rn).intSLP > 4 Then
            TRM_IX_SPT = cFRS_NORMAL
            Exit Function
        End If


        ' 初期処理
        TRM_IX_SPT = cFRS_NORMAL                        ' Return値 = 正常
        strMSG = ""
        stSPC.dblSTX = New Double(MAXSCTN) {}           ' Sp_Cut_Info構造体初期化 
        stSPC.dblSTY = New Double(MAXSCTN) {}
        Call Set_SpCut_Info(rn, cn, stSPC)              ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ情報設定
        ANG = stREG(rn).STCUT(cn).intANG                ' カット方向
        ' ｶｯﾄｵﾌ(%)→目標値に対するｵﾌｾｯﾄ値(目標値×(1＋ｶｯﾄｵﾌ/100))
        NOMx = NOM * System.Math.Abs(1 + (stREG(rn).STCUT(cn).dblCOF * 0.01))
        CutL = stREG(rn).STCUT(cn).dblDL2               ' ﾘﾐｯﾄｶｯﾄ量mm

        ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ実行(STｶｯﾄをｶｯﾄ本数分実行)
        For i = 1 To stREG(rn).STCUT(cn).intNum         ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ本数分繰返す

            ' BP絶対値移動(ｶｯﾄ位置 + PTN補正値)(絶対値)
            r = ObjSys.EX_MOVE(gSysPrm, stSPC.dblSTX(i) + stPTN(rn).dblDRX, stSPC.dblSTY(i) + stPTN(rn).dblDRY, 1)
            If (r <> 0) Then                            ' ｴﾗｰ ?
                TRM_IX_SPT = cFRS_ERR_TRIM              ' Return値 =トリマエラー
                Exit Function
            End If
            dblLN = 0.0#                                ' ｶｯﾄ量初期化

            For IDX = 1 To MAXIDX                       ' ｲﾝﾃﾞｯｸｽｶｯﾄ1～5(ﾋﾟｯﾁ大,中,小)分繰返す
STP_CHG_PIT:
                count = stREG(rn).STCUT(cn).intIXN(IDX) ' count = ｲﾝﾃﾞｯｸｽｶｯﾄ数
                ln = stREG(rn).STCUT(cn).dblDL1(IDX)    ' カット長を設定する(ﾋﾟｯﾁ大,中,小)
                For j = 1 To count                      ' ｲﾝﾃﾞｯｸｽｶｯﾄ数分繰返す

                    ' セーフティチェック
                    r = SafetyCheck()                   ' セーフティチェック
                    If (r <> cFRS_NORMAL) Then          ' エラー ?
                        TRM_IX_SPT = r                  ' Return値を設定する
                        Exit Function
                    End If

                    Call DScanModeSet(rn, cn, IDX)              ' ＤＣスキャナに接続する測定器を切り替えてスキャナをセットする。

                    MoveStop()              'V2.2.0.0⑥ 

                    ' 抵抗測定/電圧測定(内部/外部測定器)
                    r = V_R_MEAS(stREG(rn).intSLP, stREG(rn).STCUT(cn).intMType, dblMx, rn, NOMx)
                    If (r <> cFRS_NORMAL) Then                               ' エラー
                        GoTo TRM_IX_ERR
                    End If

                    ' 目標値を超えたか調べる
                    If (stREG(rn).intSLP = SLP_VTRIMPLS) Or (stREG(rn).intSLP = SLP_RTRM) Then ' +ｽﾛｰﾌﾟ/抵抗 ?
                        If (dblMx >= NOMx) Then         ' 測定値 >= 目標値なら次へ
                            TRM_IX_SPT = 1              ' Return値 = 1(目標値を超えたので終了)
                            Exit Function
                        End If
                    Else                                ' -ｽﾛｰﾌﾟ ?
                        If (dblMx <= NOMx) Then         ' 測定値 <= 目標値なら次へ
                            TRM_IX_SPT = 1              ' Return値 = 1(目標値を超えたので終了)
                            Exit Function
                        End If
                    End If

                    ' ｲﾝﾃﾞｯｸｽｶｯﾄﾋﾟｯﾁ(1:大,2～4:中,5:小)をチェックする
                    r = Get_Idx_Pitch(rn, cn, IDX, NOMx, dblMx)
                    If (r <> IDX) Then                  ' ｶｯﾄﾋﾟｯﾁ変更 ?
                        IDX = r                         ' ｲﾝﾃﾞｯｸｽｶｯﾄﾋﾟｯﾁを変更する
                        GoTo STP_CHG_PIT
                    End If

                    ' 次のｶｯﾄで最大ｶｯﾄ量を超える ? (※下記のようにしないと正しい比較ができない)
                    wkL1 = CDbl((dblLN + ln).ToString("#0.0000"))
                    wkL2 = CDbl(CutL.ToString("#0.0000"))
                    If (wkL1 > wkL2) Then ' 最大ｶｯﾄ量を超える ?
                        'ln = wkL2 - dblLN(LTFlg, cn)        ' Ln = 残りのｶｯﾄ量(←だめ※下記のようにしないとln=0とならない場合あり)
                        ln = CDbl((wkL2 - dblLN).ToString("#0.0000"))
                        If (ln <= 0) Then               ' 最大ｶｯﾄ量までカット ?
                            GoTo TRM_SPT_NEXT
                            Exit Function
                        End If
                    End If

                    ' 斜め直線ｶｯﾄ(ﾎﾟｼﾞｼｮﾆﾝｸﾞなし)
                    r = CUT2(ln, stREG(rn).STCUT(cn).dblV1, ANG)
                    Call Check_ERR_LSR_STATUS_STANBY(r)                      ' レーザアラーム８３３エラー時のプログラム終了処理
                    Call ZWAIT(stREG(rn).STCUT(cn).lngPAU(IDX)) ' ｲﾝﾃﾞｯｸｽﾋﾟｯﾁ間ﾎﾟｰｽﾞ(ms)
                    If (r <> 0) Then GoTo TRM_IX_ERR ' エラー ?
                    dblLN = dblLN + ln                  ' ｶｯﾄ済量mmを退避

#If (cCND = 1) Then ' 条件出しﾓｰﾄﾞ ?
                    strMSG = "　 抵抗番号=" + Format(rn, "0") + ",ｶｯﾄ番号=" + Format(cn, "0")
                    strMSG = strMSG + "　   ｶｯﾄ長=" + Format(ln, "#0.0####") + ",目標値=" + Format(NOMx, "0.0####") + ",測定値=" + Format(dblMx, "0.0####") + ",ｶｯﾄ量=" + Format(dblLN, "#0.0####")
                    Call Z_PRINT(strMSG + vbCrLf)
#End If

TRM_IX_NEXT:
                Next j                                  ' 次ｶｯﾄへ
            Next IDX                                    ' 次ﾋﾟｯﾁへ
TRM_SPT_NEXT:
            Call Cnv_ANG(ANG)                           ' カット方向を反対方向に変換
        Next i                                          ' 次ｻｰﾍﾟﾝﾀｲﾝｶｯﾄへ
        TRM_IX_SPT = 2                                  ' Return値 = 2(指定移動量までカット)
        Exit Function

TRM_IX_ERR:
        TRM_IX_SPT = cFRS_TRIM_NG                       ' Return値 = トリミングNG

    End Function
#End Region

#Region "ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ情報を設定する"
    '''=========================================================================
    '''<summary>ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ情報を設定する</summary>
    '''<param name="rn">    (INP) 抵抗データIndex　(1 org)</param>
    '''<param name="cn">    (INP) カットデータIndex(1 org)</param>
    '''<param name="pstSPC">(OUT) ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ情報</param>
    '''=========================================================================
    Public Sub Set_SpCut_Info(ByRef rn As Short, ByRef cn As Short, ByRef pstSPC As Sp_Cut_Info)

        Dim i As Short
        Dim Flg(2) As Short                             ' 移動ﾋﾟｯﾁﾌﾗｸﾞ
        Dim Pit As Double                               ' 移動ﾋﾟｯﾁ
        Dim STX(2) As Double                            ' ｶｯﾄ開始座標X
        Dim STY(2) As Double                            ' ｶｯﾄ開始座標y
        Dim ofx As Double                               ' ｽﾃｯﾌﾟｵﾌｾｯﾄX(ﾜｰｸ)
        Dim ofy As Double                               ' ｽﾃｯﾌﾟｵﾌｾｯﾄY(ﾜｰｸ)
        Dim cin(2) As Integer                           ' 矢印SW

        ' 初期処理
        Call Cnv_Arrow(stREG(rn).STCUT(cn).intANG, cin(1))  ' カット方向(角度)を矢印SWに変換
        Call Cnv_Arrow(stREG(rn).STCUT(cn).intANG2, cin(2)) ' ステップ方向(角度)を矢印SWに変換
        Pit = 0.0#                                      ' 移動ﾋﾟｯﾁ
        Flg(1) = 0                                      ' ﾌﾗｸﾞ初期化
        Flg(2) = 0                                      '
        STX(1) = stREG(rn).STCUT(cn).dblSTX             ' 1本目(奇数)の開始座標X
        STY(1) = stREG(rn).STCUT(cn).dblSTY             ' 1本目(奇数)の開始座標Y
        STX(2) = stREG(rn).STCUT(cn).dblSX2             ' 2本目(偶数)の開始座標X
        STY(2) = stREG(rn).STCUT(cn).dblSY2             ' 2本目(偶数)の開始座標Y

        ' カット開始座標XYを設定する
        For i = 1 To stREG(rn).STCUT(cn).intNum         ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ本数分繰返す
            ' 偶数本目のカット開始座標XYを設定する
            If (i Mod 2 = 0) Then                       ' 偶数本目 ?
                If (Flg(2) = 0) Then                    ' 移動ﾋﾟｯﾁ設定
                    Pit = 0.0#
                    Flg(2) = 1
                Else
                    Pit = stREG(rn).STCUT(cn).dblDL3 * 2
                End If

                ' ｽﾃｯﾌﾟ方向がY方向ならX座標は変わらない
                If (stREG(rn).STCUT(cn).intANG2 = 90) Or (stREG(rn).STCUT(cn).intANG2 = 270) Then
                    pstSPC.dblSTX(i) = stREG(rn).STCUT(cn).dblSX2
                    ' BP座標の設定(移動ﾋﾟｯﾁ分移動後のBP位置→STY(2))
                    Call ObjUtl.GetBPmovePitch(cin(2), ofx, ofy, Pit, STX(2), STY(2), (gSysPrm.stDEV.giBpDirXy))
                    pstSPC.dblSTY(i) = STY(2)           ' Y座標設定

                    ' ｽﾃｯﾌﾟ方向がX方向ならY座標は変わらない
                Else
                    pstSPC.dblSTY(i) = stREG(rn).STCUT(cn).dblSY2
                    ' BP座標の設定(移動ﾋﾟｯﾁ分移動後のBP位置→STX(2))
                    Call ObjUtl.GetBPmovePitch(cin(2), ofx, ofy, Pit, STX(2), STY(2), (gSysPrm.stDEV.giBpDirXy))
                    pstSPC.dblSTX(i) = STX(2)           ' X座標設定
                End If

                ' 奇数本目のカット開始座標XYを設定する
            Else
                If (Flg(1) = 0) Then                    ' 移動ﾋﾟｯﾁ設定
                    Pit = 0
                    Flg(1) = 1
                Else
                    Pit = stREG(rn).STCUT(cn).dblDL3 * 2
                End If

                ' ｽﾃｯﾌﾟ方向がY方向ならX座標は変わらない
                If (stREG(rn).STCUT(cn).intANG2 = 90) Or (stREG(rn).STCUT(cn).intANG2 = 270) Then
                    pstSPC.dblSTX(i) = stREG(rn).STCUT(cn).dblSTX
                    ' BP座標の設定(移動ﾋﾟｯﾁ分移動後のBP位置→STY(1))
                    Call ObjUtl.GetBPmovePitch(cin(2), ofx, ofy, Pit, STX(1), STY(1), (gSysPrm.stDEV.giBpDirXy))
                    pstSPC.dblSTY(i) = STY(1)           ' Y座標設定

                    ' ｽﾃｯﾌﾟ方向がX方向ならY座標は変わらない
                Else
                    pstSPC.dblSTY(i) = stREG(rn).STCUT(cn).dblSTY
                    ' BP座標の設定(移動ﾋﾟｯﾁ分移動後のBP位置→STX(1))
                    Call ObjUtl.GetBPmovePitch(cin(2), ofx, ofy, Pit, STX(1), STY(1), (gSysPrm.stDEV.giBpDirXy))
                    pstSPC.dblSTX(i) = STX(1)           ' X座標設定
                End If
            End If

        Next i

    End Sub
#End Region

#Region "ｲﾝﾃﾞｯｸｽｶｯﾄﾋﾟｯﾁ(1:大,2～4:中,5:小)を返す"
    '''=========================================================================
    '''<summary>ｲﾝﾃﾞｯｸｽｶｯﾄﾋﾟｯﾁ(1:大,2～4:中,5:小)を返す</summary>
    '''<param name="rn"> (INP) 抵抗データIndex　(1 org)</param>
    '''<param name="cn"> (INP) カットデータIndex(1 org)</param>
    '''<param name="IDX">(INP) 現在のﾋﾟｯﾁ</param>
    '''<param name="NOM">(INP) 目標値</param>
    '''<param name="Mx"> (INP) 測定値</param>
    '''<returns> ｲﾝﾃﾞｯｸｽｶｯﾄﾋﾟｯﾁ(1:大,2～4:中,5:小)</returns>
    '''=========================================================================
    Public Function Get_Idx_Pitch(ByRef rn As Short, ByRef cn As Short, ByRef IDX As Short, ByRef NOM As Double, ByRef Mx As Double) As Short

        Dim DEV As Double
        Dim i As Short

        ' 初期処理
        Get_Idx_Pitch = IDX                             ' Return値 = 現在のｲﾝﾃﾞｯｸｽｶｯﾄﾋﾟｯﾁ(1:大 → 5:小)
        DEV = System.Math.Abs((Mx / NOM) - 1) * 100     ' DEV(誤差%) = 測定値 - 目標値(絶対値)

        ' ｲﾝﾃﾞｯｸｽｶｯﾄﾋﾟｯﾁ(1:大 → 5:小) の内、誤差以上のものを返す
        For i = 1 To MAXIDX                             ' ｲﾝﾃﾞｯｸｽｶｯﾄﾋﾟｯﾁ数分繰り返す
            If (stREG(rn).STCUT(cn).dblDL1(i) = 0.0#) Then Exit Function
            Get_Idx_Pitch = i                           ' Return値 = ｲﾝﾃﾞｯｸｽｶｯﾄﾋﾟｯﾁ(1:大,2:中,3:小)
            If (DEV >= stREG(rn).STCUT(cn).dblDEV(i)) Then ' 誤差% > ?
                If (i < IDX) Then Get_Idx_Pitch = IDX ' 現在のｲﾝﾃﾞｯｸｽｶｯﾄﾋﾟｯﾁより大なら現在のｲﾝﾃﾞｯｸｽｶｯﾄﾋﾟｯﾁを返す
                Exit Function
            End If
        Next i

    End Function
#End Region

#Region "カット方向を反対方向に変換して返す(角度(0～360))"
    '''=========================================================================
    '''<summary>カット方向を反対方向に変換して返す</summary>
    '''<param name="ANG">(I/O) 角度(0～360)</param>
    '''=========================================================================
    Public Sub Cnv_ANG(ByRef ANG As Short)

        ANG = ANG + 180                         ' 角度 = 反対方向
        If (ANG >= 360) Then ANG = ANG - 360 ' 角度 = -360～360

    End Sub
#End Region

#Region "カット方向を反対方向に変換して返す(カット方向(1:180°, 2: 90°, 3:0°, 4:270°))"
    '''=========================================================================
    '''<summary>カット方向を反対方向に変換して返す</summary>
    '''<param name="ANG">(I/O) カット方向(1:180°, 2: 90°, 3:0°, 4:270°)</param>
    '''=========================================================================
    Public Sub Cnv_ANG2(ByVal ANG As Short, ByRef ANG1 As Short)

        Select Case ANG
            Case 180
                ANG1 = 0 ' 180°→ 0°
            Case 90
                ANG1 = 270 '  90°→ 270°
            Case 0
                ANG1 = 180 '   0°→ 180°
            Case 270
                ANG1 = 90 ' 270°→ 90°
        End Select

        ''Select Case ANG
        ''    Case 1
        ''        ANG = 3 ' 180°→ 0°
        ''    Case 2
        ''        ANG = 4 '  90°→ 270°
        ''    Case 3
        ''        ANG = 1 '   0°→ 180°
        ''    Case 4
        ''        ANG = 2 ' 270°→ 90°
        ''End Select

    End Sub
#End Region

#Region "カット方向を矢印SWに変換して返す"
    '''=========================================================================
    '''<summary>カット方向を矢印SWに変換して返す</summary>
    '''<param name="ANG">  (INP) 角度(0～360)</param>
    '''<param name="Arrow">(OUT) 矢印SW</param>
    '''=========================================================================
    Public Sub Cnv_Arrow(ByRef ANG As Short, ByRef Arrow As Integer)

        If (ANG = 180) Then     ' 180°
            Arrow = &H400S      ' ←
        ElseIf (ANG = 90) Then  ' 90°
            Arrow = &H800S      ' ↑
        ElseIf (ANG = 270) Then ' 270°
            Arrow = &H1000S     ' ↓
        Else                    ' 0°(360°)
            Arrow = &H200S      ' →
        End If

    End Sub
#End Region

#Region "インデックスカットトリミング x0(ﾄﾘﾐﾝｸﾞﾓｰﾄﾞ)"
    '''=========================================================================
    '''<summary>インデックスカットトリミング x0(ﾄﾘﾐﾝｸﾞﾓｰﾄﾞ)</summary>
    '''<param name="rn"> (INP) 抵抗データIndex　(1 org)</param>
    '''<param name="cn"> (INP) カット番号　　　 (1 org)</param>
    '''<param name="Mx"> (INP) 初期測定値</param>
    '''<returns> 0 = 正常
    '''          1 = 正常(目標値を超えたので終了)
    '''          2 = 指定移動量までカットしたので終了
    '''          5 = TRV目標値を超えたので終了
    '''          上記以外 = その他エラー
    ''' </returns>
    '''=========================================================================
    Public Function TRM_IX(ByRef rn As Short, ByRef cn As Short, ByRef Mx As Double) As Short

        Dim i As Short                                  ' ﾙｰﾌﾟ回数
        Dim j As Short                                  ' ﾙｰﾌﾟ回数
        Dim IDX As Short                                ' ｲﾝﾃﾞｯｸｽｶｯﾄ1～3(ﾋﾟｯﾁ大,中,小)
        Dim count As Short                              ' ｲﾝﾃﾞｯｸｽｶｯﾄ数
        Dim r As Short                                  ' 関数戻値
        Dim CutL As Double                              ' 最大ｶｯﾄ長
        Dim ln As Double                                ' 現在のｶｯﾄ長
        Dim NOM(MAX_LCUT) As Double                     ' 目標値(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後) V1.0.4.3⑦２をMAX_LCUTへ変更
        Dim NOMx As Double                              ' 目標値
        Dim VX(3) As Double                             ' 作業域
        Dim strMSG As String                            ' メッセージ編集域
        Dim wkL1 As Double                              ' 作業域
        Dim wkL2 As Double                              ' 作業域
        Dim dblMx As Double                             ' 作業域
        Dim dblQrate As Double
        Dim shSLP As Short
        Dim CutLSum As Double                           ' V1.0.4.3⑧ 最大ｶｯﾄ長積算値
        Dim dblQRateL(MAX_LCUT) As Double               ' V1.0.4.3⑦　Ｑレート
        Dim dblSpeedL(MAX_LCUT) As Double               ' V1.0.4.3⑦　速度
        Dim dblSpeed As Double                          ' V1.0.4.3⑦　速度
        Dim SaveIDX As Short                            'V2.1.0.0⑤
        Try

            '---------------------------------------------------------------------------
            '   初期処理
            '---------------------------------------------------------------------------
            CutLSum = 0.0                                   ' V1.0.4.3⑧
            TRM_IX = cFRS_NORMAL                            ' Return値 = 正常
            strMSG = ""
            LTFlg = 1                                       ' Lﾀｰﾝﾌﾗｸﾞ(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
            LTAng(1) = stREG(rn).STCUT(cn).intANG           ' ANG(1) = Lﾀｰﾝ前のｶｯﾄ方向
            LTAng(2) = stREG(rn).STCUT(cn).intANG2          ' ANG(2) = Lﾀｰﾝ後のｶｯﾄ方向
            dblML(1) = stREG(rn).STCUT(cn).dblDL2           ' ﾘﾐｯﾄｶｯﾄ量mm(1:Lﾀｰﾝ前)
            dblML(2) = stREG(rn).STCUT(cn).dblDL3           ' ﾘﾐｯﾄｶｯﾄ量mm(2:Lﾀｰﾝ後)
            LTP = stREG(rn).STCUT(cn).dblLTP                ' Lﾀｰﾝﾎﾟｲﾝﾄ(%)
            NOM(1) = stREG(rn).dblNOM                       ' 目標値
            ' ｶｯﾄｵﾌ(%)→目標値に対するｵﾌｾｯﾄ値(目標値×(1＋ｶｯﾄｵﾌ/100))
            NOMx = Func_CalNomForCutOff(rn, cn, NOM(1))          'カットオフによる目標値
            NOM(1) = NOMx                                   ' Lﾀｰﾝ前目標値 = 目標値
            NOM(2) = NOMx                                   ' Lﾀｰﾝ後目標値 = 目標値

            ' Lﾀｰﾝ ﾎﾟｲﾝﾄ情報を設定する(Lｶｯﾄ時)
            'V1.0.4.3⑦            If (stREG(rn).STCUT(cn).intCTYP = CNS_CUTP_L) And ((LTP <> 0.0#) And (LTP < 100.0#)) Then ' LｶｯﾄでLﾀｰﾝ ﾎﾟｲﾝﾄ指定あり ?
            'V1.0.4.3⑦            NOM(1) = Mx + (NOM(2) - Mx) * (LTP * 0.01)  ' Lﾀｰﾝ前目標値設定(初期値＋(目標値-初期値)×Lﾀｰﾝﾎﾟｲﾝﾄ/100)
            'V1.0.4.3⑦            End If

            ' Lﾀｰﾝ前ﾘﾐｯﾄｶｯﾄ量mmと目標値を設定する
            CutL = dblML(LTFlg)                             ' ﾘﾐｯﾄｶｯﾄ量mm(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
            NOMx = NOM(LTFlg)                               ' 目標値(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)

            '-----------------------------------------------------------------------
            '   ユーザプログラム特殊処理 カット毎の目標値を求める。START
            '-----------------------------------------------------------------------
            If UserSub.IsSpecialTrimType Then
                NOMx = UserSub.GetTargeResistancetValue(rn, cn)
                If UserSub.IsTrimType2 Or UserSub.IsTrimType3() Then    'V1.0.4.3④IsTrimType3()追加
                    ' G15A-15A.BAS : 14580       IF TRM1#(CN1%)<=.5# THEN GOTO *NEXT.CT1
                    '###1032                    If NOMx <= 0.5# Then
                    If NOMx <= 0.005# Then
                        Return (cFRS_NORMAL)
                    End If
                End If
            End If

            '-----------------------------------------------------------------------
            '   ユーザプログラム特殊処理 カット毎の目標値を求める。END
            '-----------------------------------------------------------------------

            ' Qレート設定
            If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then
                dblQrate = stREG(rn).STCUT(cn).intQF1
                dblQrate = dblQrate / 10.0
            Else
#If cOSCILLATORcFLcUSE Then
                ' FL時は加工条件番号テーブルからQレートを設定する(カットスピードはデータから設定)
                IDX = stREG(rn).STCUT(cn).intCND(CUT_CND_L1)
                dblQrate = stCND.Freq(IDX)
#End If
            End If

            ' V1.0.4.3⑦↓
            '-----------------------------------------------------------------------
            '   Ｌカット種別追加（６点ターンポイント仕様）
            '-----------------------------------------------------------------------
            If (stREG(rn).STCUT(cn).intCTYP = CNS_CUTP_L) Then
                For i = 1 To MAX_LCUT
                    NOM(i) = Mx + (NOMx - Mx) * (stREG(rn).STCUT(cn).dTurnPoint(i) * 0.01)  ' Lﾀｰﾝ前目標値設定(初期値＋(目標値-初期値)×Lﾀｰﾝﾎﾟｲﾝﾄ/100)
                    dblML(i) = stREG(rn).STCUT(cn).dCutLen(i)
                    LTAng(i) = stREG(rn).STCUT(cn).dAngle(i)
                    dblQRateL(i) = stREG(rn).STCUT(cn).dQRate(i)
                    dblSpeedL(i) = stREG(rn).STCUT(cn).dSpeed(i)
                Next
                NOM(MAX_LCUT) = NOMx
                CutL = dblML(1)                             ' ﾘﾐｯﾄｶｯﾄ量mm(Ｌ１カット)
                NOMx = NOM(1)                               ' 目標値(Ｌ１カット)
                dblQrate = dblQRateL(1) / 10.0
                dblSpeed = dblSpeedL(1)
            Else
                dblSpeed = stREG(rn).STCUT(cn).dblV1
            End If
            ' V1.0.4.3⑦↑

            ' ｶｯﾄ量初期化
            For i = 1 To MAXCTN                             ' MAXカット数分繰返す
                For j = 1 To MAX_LCUT                           ' MAXカット数分繰返す
                    dblLN(j, i) = 0.0#                          ' ｶｯﾄ量初期化(1:Lﾀｰﾝ前)
                    'V1.0.4.3⑦                    dblLN(1, i) = 0.0#                          ' ｶｯﾄ量初期化(1:Lﾀｰﾝ前)
                    'V1.0.4.3⑦                    dblLN(2, i) = 0.0#                          ' ｶｯﾄ量初期化(2:Lﾀｰﾝ後)
                Next j
            Next i

            '---------------------------------------------------------------------------
            '   ｲﾝﾃﾞｯｸｽｶｯﾄでLｶｯﾄ/STｶｯﾄを行う
            '---------------------------------------------------------------------------
            For IDX = 1 To MAXIDX                           ' ｲﾝﾃﾞｯｸｽｶｯﾄ1～5(ﾋﾟｯﾁ大,中,小)分繰返す
STP_CHG_PIT:
                count = stREG(rn).STCUT(cn).intIXN(IDX)     ' count = ｲﾝﾃﾞｯｸｽｶｯﾄ数
                ln = stREG(rn).STCUT(cn).dblDL1(IDX)        ' カット長を設定する(ﾋﾟｯﾁ大,中,小)
                For i = 1 To count                          ' ｲﾝﾃﾞｯｸｽｶｯﾄ数分繰返す
#If (cCND = 1) Then ' 条件出しﾓｰﾄﾞ ?
                    strMSG = "　 抵抗番号=" + Format(rn, "0") + ",ｶｯﾄ番号=" + Format(cn, "0")
                    strMSG = strMSG + ",目標値(Lﾀｰﾝ前,後)=" + Format(NOM(1), "0.0####") + "," + Format(NOM(2), "0.0####") + vbCrLf
                    strMSG = strMSG + "　   ｶｯﾄ長=" + Format(ln, "#0.0####") + ",目標値=" + Format(NOMx, "0.0####") + ",LTFlg=" + Format(LTFlg, "0") + ",ｶｯﾄ量(Lﾀｰﾝ前,後)=" + Format(dblLN(1, cn), "#0.0####") + "," + Format(dblLN(2, cn), "#0.0####")
                    Call Z_PRINT(strMSG + vbCrLf)
#End If

                    MoveStop()              'V2.2.0.0⑥ 

                    'V2.1.0.0⑤↓
                    If IsCutVariationJudgeExecute() AndAlso UserSub.IsCutMeasureBefore() Then
                        dblMx = UserSub.CutVariationMeasureBeforeGet()
                    Else
                        'V2.1.0.0⑤↑

                        Call UserSub.ChangeMeasureSpeed(rn, cn, IDX)     ' 測定速度の変更（特注処理）

                        Call DScanModeSet(rn, cn, IDX)              ' ＤＣスキャナに接続する測定器を切り替えてスキャナをセットする。

                        ' 電圧(外部/内部)/抵抗測定(外部/内部)を行う
                        ' 測定レンジの目標値を最終目標値にする。2013.3.28                    r = V_R_MEAS(stREG(rn).intSLP, stREG(rn).STCUT(cn).intIXMType(IDX), dblMx, rn, NOMx)
                        If UserSub.IsSpecialTrimType Then
                            r = V_R_MEAS(stREG(rn).intSLP, stREG(rn).STCUT(cn).intIXMType(IDX), dblMx, rn, UserSub.GetTRV())
                        Else
                            r = V_R_MEAS(stREG(rn).intSLP, stREG(rn).STCUT(cn).intIXMType(IDX), dblMx, rn, NOMx)
                        End If
                        If (r <> cFRS_NORMAL) Then                          ' エラー
                            Call UserSub.ResoreMeasureSpeed(rn, cn, IDX)         ' 測定速度の変更を元に戻す（特注処理）
                            Return (cFRS_TRIM_NG)                           ' Return値 = トリミングNG
                        End If

                        Call UserSub.ResoreMeasureSpeed(rn, cn, IDX)     ' 測定速度の変更を元に戻す（特注処理）
                    End If       'V2.1.0.0⑤
                    SaveIDX = IDX                                       'V2.1.0.0⑤

                    If bDebugLogOut Then
                        DebugLogOut("抵抗[" & rn.ToString & "]カット[" & cn.ToString & "] IX測定値=[" & dblMx.ToString & "]")
                    End If

                    If UserSub.IsSpecialTrimType And dblMx >= UserSub.GetTRV() Then
                        If bDebugLogOut Then
                            DebugLogOut("TRV目標値到達 抵抗[" & rn.ToString & "]カット[" & cn.ToString & "] IX測定値=[" & dblMx.ToString & "]  TRV=[" & UserSub.GetTRV().ToString & "]")
                        End If
                        UserSub.CutVariationMeasureAfterSet(dblMx) 'V2.1.0.0①カット後の測定値保存
                        Return (5)
                    End If

#If (cCND = 1) Then ' 条件出しﾓｰﾄﾞ ?
                    strMSG = "■測定値=" + dblMx.ToString("#0.0###")
                    Call Z_PRINT(strMSG + vbCrLf)
#End If
                    ' 目標値を超えたか調べる
                    If (stREG(rn).intSLP = SLP_VTRIMPLS) Or (stREG(rn).intSLP = SLP_RTRM) Then ' +ｽﾛｰﾌﾟ/抵抗 ?
                        If (dblMx >= NOMx) Then             ' 測定値 >= 目標値なら次へ
                            TRM_IX = 1                      ' Return値 = 1(目標値を超えたので終了)
                            '-------------------------
                            '   Lﾀｰﾝ処理(Lｶｯﾄ時)
                            '-------------------------
                            ' Lｶｯﾄ以外ならEXIT
                            If (stREG(rn).STCUT(cn).intCTYP <> CNS_CUTP_L) Then
                                GoTo TRM_IX_EXIT
                            End If
                            If (LTFlg >= MAX_LCUT) Then
                                GoTo TRM_IX_EXIT                            ' Lﾀｰﾝ後ならEXIT
                            End If
                            TRM_IX = 0                                      ' Return値 = 正常
                            LTFlg = LTFlg + 1                               ' Lﾀｰﾝﾌﾗｸﾞ = 次のカット(Lﾀｰﾝ後)
                            CutL = dblML(LTFlg)                             ' ﾘﾐｯﾄｶｯﾄ量mm(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
                            NOMx = NOM(LTFlg)                               ' 目標値(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
                            dblQrate = dblQRateL(LTFlg) / 10.0
                            dblSpeed = dblSpeedL(LTFlg)
                            ln = stREG(rn).STCUT(cn).dblDL1(IDX)            ' カット長を設定する(ﾋﾟｯﾁ大,中,小)
                            GoTo TRM_IX_NEXT                               ' Lｶｯﾄを行う
                            '-------------------------
                            '   Lﾀｰﾝ処理(Lｶｯﾄ時)
                            '-------------------------


                        End If
                    Else                                    ' -ｽﾛｰﾌﾟ ?
                        If (dblMx <= NOMx) Then             ' 測定値 <= 目標値なら次へ
                            TRM_IX = 1                      ' Return値 = 1(目標値を超えたので終了)
                            '-------------------------
                            '   Lﾀｰﾝ処理(Lｶｯﾄ時)
                            '-------------------------
                            ' Lｶｯﾄ以外ならEXIT
                            If (stREG(rn).STCUT(cn).intCTYP <> CNS_CUTP_L) Then
                                GoTo TRM_IX_EXIT
                            End If
                            If (LTFlg >= MAX_LCUT) Then
                                GoTo TRM_IX_EXIT                            ' Lﾀｰﾝ後ならEXIT
                            End If
                            TRM_IX = 0                                      ' Return値 = 正常
                            LTFlg = LTFlg + 1                               ' Lﾀｰﾝﾌﾗｸﾞ = 次のカット(Lﾀｰﾝ後)
                            CutL = dblML(LTFlg)                             ' ﾘﾐｯﾄｶｯﾄ量mm(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
                            NOMx = NOM(LTFlg)                               ' 目標値(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
                            dblQrate = dblQRateL(LTFlg) / 10.0
                            dblSpeed = dblSpeedL(LTFlg)
                            ln = stREG(rn).STCUT(cn).dblDL1(IDX)            ' カット長を設定する(ﾋﾟｯﾁ大,中,小)
                            GoTo TRM_IX_NEXT                                ' Lｶｯﾄを行う
                            '-------------------------
                            '   Lﾀｰﾝ処理(Lｶｯﾄ時)
                            '-------------------------

                        End If
                    End If

                    ' ｲﾝﾃﾞｯｸｽｶｯﾄﾋﾟｯﾁ(1:大,2-4:中,5:小)をチェックする
                    r = Get_Idx_Pitch(rn, cn, IDX, NOMx, dblMx) ' 目標値との誤差によりﾋﾟｯﾁを変更する
                    If (r <> IDX) Then                      ' ｶｯﾄﾋﾟｯﾁ変更 ?
                        IDX = r                             ' ｲﾝﾃﾞｯｸｽｶｯﾄﾋﾟｯﾁを変更する
                        GoTo STP_CHG_PIT
                    End If

                    ' 次のｶｯﾄで最大ｶｯﾄ量を超える ? (※下記のようにしないと正しい比較ができない)
                    wkL1 = CDbl((dblLN(LTFlg, cn) + ln).ToString("#0.0000"))
                    wkL2 = CDbl(CutL.ToString("#0.0000"))
                    If (wkL1 > wkL2) Then                   ' 最大ｶｯﾄ量を超える ?
                        ' ln = 残りのｶｯﾄ量(下記のようにしないとln=0とならない場合あり)
                        ln = CDbl((wkL2 - dblLN(LTFlg, cn)).ToString("#0.0000"))
                        If (ln <= 0) Then                   ' 最大ｶｯﾄ量までカット ?
                            TRM_IX = 2                      ' Return値 = 2(指定移動量までカット)
                            '-------------------------
                            '   Lﾀｰﾝ処理(Lｶｯﾄ時)
                            '-------------------------
                            ' Lｶｯﾄ以外ならEXIT
                            ' Lｶｯﾄ以外ならEXIT
                            If (stREG(rn).STCUT(cn).intCTYP <> CNS_CUTP_L) Then
                                GoTo TRM_IX_EXIT
                            End If
                            If (LTFlg >= MAX_LCUT) Then
                                GoTo TRM_IX_EXIT                            ' Lﾀｰﾝ後ならEXIT
                            End If
                            TRM_IX = 0                                      ' Return値 = 正常
                            LTFlg = LTFlg + 1                               ' Lﾀｰﾝﾌﾗｸﾞ = 次のカット(Lﾀｰﾝ後)
                            CutL = dblML(LTFlg)                             ' ﾘﾐｯﾄｶｯﾄ量mm(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
                            NOMx = NOM(LTFlg)                               ' 目標値(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
                            ln = stREG(rn).STCUT(cn).dblDL1(IDX)            ' カット長を設定する(ﾋﾟｯﾁ大,中,小)
                            GoTo TRM_IX_NEXT                               ' Lｶｯﾄを行う  ※※GoTo TRM_IX_NEXTがエラーとなるので修正必要
                            '-------------------------
                            '   Lﾀｰﾝ処理(Lｶｯﾄ時)
                            '-------------------------


                        End If
                    End If

                    ' 斜め直線ｶｯﾄ(ﾎﾟｼﾞｼｮﾆﾝｸﾞなし)
                    If (stREG(rn).intSLP = SLP_ATRIMPLS) Then
                        shSLP = 1
                    ElseIf (stREG(rn).intSLP = SLP_ATRIMMNS) Then
                        shSLP = 2
                    Else
                        shSLP = stREG(rn).intSLP
                    End If
                    r = TrimSt(FORCE_MODE, 0, 0, shSLP, LTAng(LTFlg), ln, dblSpeed, dblSpeed, _
                                  dblQrate, dblQrate, stREG(rn).STCUT(cn).intCND(CUT_CND_L1), stREG(rn).STCUT(cn).intCND(CUT_CND_L1))
                    'r = TrimSt(FORCE_MODE, 0, 0, shSLP, LTAng(LTFlg), ln, stREG(rn).STCUT(cn).dblV1, stREG(rn).STCUT(cn).dblV1, _
                    '              dblQrate, dblQrate, stREG(rn).STCUT(cn).intCND(CUT_CND_L1), stREG(rn).STCUT(cn).intCND(CUT_CND_L1))
                    Call ZWAIT(stREG(rn).STCUT(cn).lngPAU(IDX)) ' ｲﾝﾃﾞｯｸｽﾋﾟｯﾁ間ﾎﾟｰｽﾞ(ms)
                    If (r <> 0) And (r <> 2) Then
                        Return (cFRS_TRIM_NG)                           ' Return値 = トリミングNG
                    End If
                    UserSub.CutVariationCutSet()                        'V2.1.0.0⑤カットが有った事を記録する。
                    dblLN(LTFlg, cn) = dblLN(LTFlg, cn) + ln    ' ｶｯﾄ済量mmを退避(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
                    CutLSum = CutLSum + ln                      'V1.0.4.3⑧ 積算値

TRM_IX_NEXT:
                Next i                                      ' 次ｶｯﾄへ

                If count > 0 Then                   'V2.1.0.0⑤ タクトアップの為、カット数が０でもIDX回行う為
                    ' セーフティチェック
                    r = SafetyCheck()                           ' セーフティチェック
                    If (r <> 0) Then                            ' エラー ?
                        TRM_IX = r                              ' Return値 = セーフティチェックエラー
                        Exit Function
                    End If
                End If                              'V2.1.0.0⑤ 
            Next IDX                                        ' 次ﾋﾟｯﾁへ

TRM_IX_EXIT:
            'V2.1.0.0⑤↓測定しないで抜けるパターン有り
            If IsCutVariationJudgeExecute() AndAlso UserSub.IsNotCutMeasureAfter() = True Then
                Call UserSub.ChangeMeasureSpeed(rn, cn, SaveIDX)     ' 測定速度の変更（特注処理）

                Call DScanModeSet(rn, cn, SaveIDX)              ' ＤＣスキャナに接続する測定器を切り替えてスキャナをセットする。

                ' 電圧(外部/内部)/抵抗測定(外部/内部)を行う
                If UserSub.IsSpecialTrimType Then
                    r = V_R_MEAS(stREG(rn).intSLP, stREG(rn).STCUT(cn).intIXMType(SaveIDX), dblMx, rn, UserSub.GetTRV())
                Else
                    r = V_R_MEAS(stREG(rn).intSLP, stREG(rn).STCUT(cn).intIXMType(SaveIDX), dblMx, rn, NOMx)
                End If
                If (r <> cFRS_NORMAL) Then                          ' エラー
                    Call UserSub.ResoreMeasureSpeed(rn, cn, SaveIDX)         ' 測定速度の変更を元に戻す（特注処理）
                    Call Z_PRINT("インデックスカット追加測定時エラー抵抗[" & rn.ToString & "]カット[" & cn.ToString & "]" & vbCrLf)
                    CutVariationDebugLogOut("インデックスカット追加測定時エラー 抵抗[" & rn.ToString & "]カット[" & cn.ToString & "]IX[" & SaveIDX.ToString & "]")
                    Return (cFRS_TRIM_NG)                           ' Return値 = トリミングNG
                End If

                Call UserSub.ResoreMeasureSpeed(rn, cn, SaveIDX)     ' 測定速度の変更を元に戻す（特注処理）

                CutVariationDebugLogOut("抵抗[" & rn.ToString & "]カット[" & cn.ToString & "]IX[" & SaveIDX.ToString & "] IX測定値=[" & dblMx.ToString & "]")

                UserSub.CutVariationMeasureAfterSet(dblMx)
            End If
            'V2.1.0.0⑤↑

            If (stREG(rn).STCUT(cn).intCTYP = CNS_CUTP_ST_TR) Then
                'V2.0.0.0⑦↓
                r = CUT_RETRACE(rn, cn, CutLSum)
                Return (r)
                'V2.0.0.0⑦↑
                'V2.0.0.0⑦                ' 斜め直線ｶｯﾄ(ﾎﾟｼﾞｼｮﾆﾝｸﾞなし)
                'V2.0.0.0⑦                If (stREG(rn).intSLP = SLP_ATRIMPLS) Then
                'V2.0.0.0⑦                    shSLP = 1
                'V2.0.0.0⑦                ElseIf (stREG(rn).intSLP = SLP_ATRIMMNS) Then
                'V2.0.0.0⑦                    shSLP = 2
                'V2.0.0.0⑦                Else
                'V2.0.0.0⑦                    shSLP = stREG(rn).intSLP
                'V2.0.0.0⑦                End If
                'V2.0.0.0⑦                dblQrate = stREG(rn).STCUT(cn).intQF2
                'V2.0.0.0⑦                dblQrate = dblQrate / 10.0
                'V2.0.0.0⑦                Call STRXY(rn, stREG(rn).STCUT(cn).dblSTX + stREG(rn).STCUT(cn).dblSX2, stREG(rn).STCUT(cn).dblSTY + stREG(rn).STCUT(cn).dblSY2)
                'V2.0.0.0⑦                r = TrimSt(FORCE_MODE, CNS_CUTP_NORMAL, 0, shSLP, LTAng(LTFlg), CutLSum, stREG(rn).STCUT(cn).dblV2, stREG(rn).STCUT(cn).dblV2, _
                'V2.0.0.0⑦                              dblQrate, dblQrate, stREG(rn).STCUT(cn).intCND(CUT_CND_L1), stREG(rn).STCUT(cn).intCND(CUT_CND_L1))
                'V2.0.0.0⑦                If (r <> 0) And (r <> 2) Then
                'V2.0.0.0⑦                    Return (cFRS_TRIM_NG)                           ' Return値 = トリミングNG
                'V2.0.0.0⑦                End If
            End If

            Exit Function

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.TRM_IX() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cFRS_TRIM_NG)                           ' Return値 = トリミングNG
        End Try
    End Function
#End Region

    '==========================================================================
    '   測定処理
    '==========================================================================
#Region "測定処理"
    ''' <summary>
    ''' 測定器の外部内部切り替え状態記録変数を初期化して再度設定する様にする。
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub DScanModeReset()
        gLastDScanMode = -999
        gintPRH = -999
        gintPRL = -999
        gintPRG = -999
    End Sub

    '''=========================================================================
    ''' <summary>
    ''' DCスキャナに接続する測定器の切り替えとスキャナーセットリセットを必ず行う。
    ''' </summary>
    ''' <param name="rn">抵抗番号</param>
    ''' <param name="cn">カット番号、０以外の指定の時はCUTは抵抗の設定を行う</param>
    ''' <param name="idx">インデックスカット番号１～MAXIDX</param>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub DScanModeResetSet(ByVal rn As Integer, ByVal cn As Integer, ByVal idx As Integer)
        Try
            DScanModeReset()
            DScanModeSet(rn, cn, idx)
        Catch ex As Exception
            Call Z_PRINT("UserBas.DScanModeSet() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub

    '''=========================================================================
    ''' <summary>
    ''' DCスキャナに接続する測定器の切り替えとスキャナーセット
    ''' </summary>
    ''' <param name="rn">抵抗番号</param>
    ''' <param name="cn">カット番号、０以外の指定の時はCUTは抵抗の設定を行う</param>
    ''' <param name="idx">インデックスカット番号１～MAXIDX</param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub DScanModeSet(ByVal rn As Integer, ByVal cn As Integer, ByVal idx As Integer)
        '-----------------------------------------------'
        ' DCスキャナに接続する測定器を切替る    '
        '-----------------------------------------------'
        Try

            If cn = 0 Then
                If (stREG(rn).intMType = 0) Then                            ' 内部測定
                    If (stREG(rn).intSLP = SLP_RTRM Or stREG(rn).intSLP = SLP_RMES) Then  ' 抵抗測定
                        If gLastDScanMode <> 1 Then
                            Call MFSET("R")
                            Call DScanModeReset()
                            gLastDScanMode = 1
                        End If
                    Else
                        If gLastDScanMode <> 2 Then
                            Call MFSET("V")                                 ' 電圧測定
                            Call DScanModeReset()
                            gLastDScanMode = 2
                        End If
                    End If
                Else
                    If gLastDScanMode <> 3 Then
                        Call MFSET("X")                                     ' 外部機器測定
                        Call DScanModeReset()
                        gLastDScanMode = 3
                    End If
                End If
            Else
                If (stREG(rn).STCUT(cn).intIXMType(idx) = 0) Then           ' 電圧内部測定
                    If (stREG(rn).intSLP = SLP_RTRM Or stREG(rn).intSLP = SLP_RMES) Then  ' 抵抗測定
                        If gLastDScanMode <> 1 Then
                            Call MFSET("R")
                            Call DScanModeReset()
                            gLastDScanMode = 1
                        End If
                    Else
                        If gLastDScanMode <> 2 Then
                            Call MFSET("V")                                 ' 電圧測定
                            Call DScanModeReset()
                            gLastDScanMode = 2
                        End If
                    End If
                Else
                    If gLastDScanMode <> 3 Then
                        Call MFSET("X")                                     ' 外部機器測定
                        Call DScanModeReset()
                        gLastDScanMode = 3
                    End If
                End If
            End If

            ' モード切替えたら必ずDCスキャナのプローブ番号を設定し直す。
            If gintPRH <> stREG(rn).intPRH Or gintPRL <> stREG(rn).intPRL Or gintPRG <> stREG(rn).intPRG Then
                'V1.0.4.3⑨                Call DSCAN(stREG(rn).intPRH, stREG(rn).intPRL, stREG(rn).intPRG)
                Call DSCAN(UserSub.ConvtChannel(stREG(rn).intPRH), UserSub.ConvtChannel(stREG(rn).intPRL), UserSub.ConvtChannel(stREG(rn).intPRG))  'V1.0.4.3⑨
                gintPRH = stREG(rn).intPRH
                gintPRL = stREG(rn).intPRL
                gintPRG = stREG(rn).intPRG
                'Call System.Threading.Thread.Sleep(10)                  ' Wait(ms)
                Call System.Threading.Thread.Sleep(200)                  ' Wait(ms) 20130418
            End If

        Catch ex As Exception
            Call Z_PRINT("UserBas.DScanModeSet() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Sub
    '''=========================================================================
    '''<summary>測定処理</summary>
    '''<returns>cFRS_NORMAL   = 正常
    '''         cFRS_ERR_RST  = RESET SW押下
    '''         cFRS_ERR_TRIM = トリマエラー
    '''         その他のエラー
    '''</returns>
    '''=========================================================================
    Public Function Meas() As Short

        Dim rn As Short                                                 ' 抵抗番号
        Dim rtn As Short = cFRS_NORMAL                                  ' 戻値
        Dim r As Short                                                  ' 戻値
        Dim dblMx As Double = 0.0                                       ' 測定値(V)
        Dim strMSG As String

        Try
            ' 初期処理
            Call Disp_Init()                                            ' 見出し表示(ログ画面)/印刷

            ' ﾎﾟｰｽﾞ付きﾌﾟﾛｰﾌﾞON(ZﾌﾟﾛｰﾌﾞをON位置(Z.ZON)に移動)
            r = Prob_On()                                               ' ﾌﾟﾛｰﾌﾞON
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
                rtn = r                                                 ' Return値設定
                GoTo Meas_End
            End If


            ' 測定処理(DigH-SWにより抵抗測定か電圧測定のどちらかのみ行う)
            For rn = 1 To stPLT.RCount                                  ' 抵抗数分繰返す
                'V2.0.0.0                If (stREG(rn).intSLP = SLP_NG_MARK Or stREG(rn).intSLP = SLP_OK_MARK) Then                ' 7:NGﾏｰｷﾝｸﾞ の時測定しない。'V1.0.4.3⑤ ＯＫマーキング(SLP_OK_MARK)追加
                'V2.0.0.0                    GoTo Meas_Next
                'V2.0.0.0                End If
                If UserModule.IsMeasureResistor(rn) Then                    'V2.0.0.0

                    dblNM(1) = 0.0#                                         ' 目標値
                    dblVX(1) = 0.0#                                         ' 測定値

                    MoveStop()              'V2.2.0.0⑥ 

                    ' セーフティチェック
                    r = SafetyCheck()                                       ' セーフティチェック
                    If (r <> cFRS_NORMAL) Then                              ' エラー ?
                        rtn = r                                             ' Return値 = セーフティチェックエラー
                        GoTo Meas_End
                    End If

                    'V2.0.0.0②↓
                    If bPowerOnOffUse Then
                        If FUNC_OK = Func_V_On_Judge(rn) Then               '   ＯＮ機器有りかのチェック
                            r = Func_V_On_Ex(rn)                            '   全ての対象のＯＮ機器をＯＮする。
                            If (FUNC_NG = r) Then
                                GoTo Meas_End
                            End If
                        End If
                    End If
                    'V2.0.0.0②↑

                    dblNM(1) = stREG(rn).dblNOM                             ' 目標値設定
                    If UserSub.IsSpecialTrimType And IsCutResistor(rn) Then ' トリミング抵抗の時
                        UserSub.CalcTargeResistancetValue(rn)               ' 目標値を求める。
                        dblNM(1) = UserSub.GetTRV()                         ' 目標値設定
                    End If

                    Call DScanModeSet(rn, 0, 0)                             ' DCスキャナに接続する測定器を切替る 

                    ' 電圧測定前のﾎﾟｰｽﾞ
                    If (rn = 1) Then                                        ' 最初の抵抗 ?
                        Call ZWAIT(glWTimeM(1))                             ' Wait(ms)
                    End If
                    If (rn Mod 2 = 0) Then                                  ' 偶数抵抗 ?
                        Call ZWAIT(glWTimeM(2))                             ' Wait(ms)
                    Else                                                    ' 奇数抵抗
                        Call ZWAIT(glWTimeM(3))                             ' Wait(ms)
                    End If

                    Dim MesTime As Integer
                    'V2.0.0.0⑧                If stREG(rn).intMType = 1 Then                          ' 外部測定器
                    'V2.0.0.0⑧                    MesTime = gGpibMultiMeterCount
                    'V2.0.0.0⑧                Else
                    'V2.0.0.0⑧                    MesTime = 1
                    'V2.0.0.0⑧                End If
                    MesTime = stREG(rn).intFTReMeas                                 'V2.0.0.0⑧

                    For i As Integer = 1 To MesTime
                        ' 抵抗測定/電圧測定(内部/外部測定器)
                        r = V_R_MEAS(stREG(rn).intSLP, stREG(rn).intMType, dblVX(1), rn, dblNM(1))
                    Next

                    If UserSub.IsTrimType1() Or UserSub.IsTrimType4() Then  ' 温度センサー'V2.0.0.0①sTrimType4()追加
                        If stREG(rn).intSLP = SLP_RMES Then             ' 抵抗測定のみ
                            Call UserSub.SetStandardResValue(dblVX(1))  ' 標準抵抗
                        End If
                    End If

                    ' 測定値表示
                    Call Disp_Final(rn)                                     ' 測定値表示

                    If (Form1.System1.Sys_Err_Chk_EX(gSysPrm, APP_MODE_LOTCHG) <> cFRS_NORMAL) Then ' 非常停止等 ?
                        Call Form1.AppEndDataSave()                                       ' ｿﾌﾄ強制終了時のﾃﾞｰﾀ保存確認
                        Call Form1.AplicationForcedEnding()                               ' ｿﾌﾄ強制終了処理
                        End
                    End If

                    'V2.0.0.0②↓
                    If bPowerOnOffUse Then
                        If FUNC_OK = Func_V_Off_Judge(rn) Then                  '   電圧OFF有り？                                               ' DC電源装置 電圧OFF
                            r = Func_V_Off_Ex(rn)                               '   電圧OFF
                            If (FUNC_NG = r) Then
                                GoTo Meas_End
                            End If
                        End If
                    End If
                    'V2.0.0.0②↑

                End If                                                  'V2.0.0.0

Meas_Next:
            Next rn                                                     ' 次抵抗へ

            ' 終了処理
Meas_End:
            'Call Rel_Off(RelBit)                                        ' ﾘﾚｰOFF
            'Call DSCAN(Z0, Z0, Z0)                                      ' DCスキャナオフ
            r = V_Off()                                                 ' DC電源装置 電圧OFF
            r = Prob_Off()                                              ' Z2/ZﾌﾟﾛｰﾌﾞをOFF位置に移動
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
                rtn = r                                                 ' Return値設定(非常停止等)
            End If

            Return (rtn)                                                ' (注)エラー時のメッセージは表示済み

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.Meas() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cFRS_TRIM_NG)
        End Try
    End Function

    'V2.0.0.0②↓
    '''=========================================================================
    '''<summary>測定処理</summary>
    '''<returns>cFRS_NORMAL   = 正常
    '''         cFRS_ERR_RST  = RESET SW押下
    '''         cFRS_ERR_TRIM = トリマエラー
    '''         その他のエラー
    '''</returns>
    '''=========================================================================
    Public Function Power() As Short
        Dim rn As Short                                                 ' 抵抗番号
        Dim rtn As Short = cFRS_NORMAL                                  ' 戻値
        Dim r As Short                                                  ' 戻値
        Dim dblMx As Double = 0.0                                       ' 測定値
        Dim NowTime As Long
        Dim dMaxCurrent As Double = Double.MaxValue                     ' 最大印加電流
        Dim Gno As Short                                                ' 電源の番号
        Dim strMSG As String
        Dim iRnoNo As Short = -1

        Try
            Dim StopWatch As New System.Diagnostics.Stopwatch
            ' 初期処理
            Dim lAppliedSecond As Long = Long.Parse(stUserData.dAppliedSecond * 1000.0)
            Dim dAddVlot As Double

            strJUG(0) = JG_OK

            Call Z_PRINT("=== 電圧印加 ===" & vbCrLf)

            'V2.2.1.7③↓
            ' マーク印字モードはトリミングでは処理しない 
            If UserSub.IsTrimType5() Then
                Call Z_PRINT("製品種別マーク印字はx2モードで実行してください。 " & vbCrLf)
                Return (cFRS_ERR_RST)
            End If
            'V2.2.1.7③↑

            For rn = 1 To stPLT.RCount Step 1
                If IsCutResistor(rn) Then
                    '①	最大印加電流値の計算
                    '定格電圧算出（Ｖ）＝√定格電力（Ｗ）×入力抵抗値（Ω）
                    '　　※　入力抵抗値は、データ設定コマンドにて設定された設定抵抗値です。
                    '印加電圧＝定格電圧×倍率
                    '最大印加電流＝印加電圧／抵抗値×個数×電流制限（倍）

                    '計算例）
                    '定格電力 = 0.125(W)
                    '入力抵抗値（設定抵抗値）=1000Ω
                    '倍率 = 3.5
                    '電流制限 = 1.40倍
                    '定格電圧算出（Ｖ）＝ √（0.125*1000）＝ 11.180
                    '印加電圧 = 11.18 * 3.5 = 39.131
                    '最大印加電流 = 39.131 / 1000 * 3 * 1.4 = 0.162(A)
                    If iRnoNo < 0 Then
                        iRnoNo = rn
                    End If
                    dAddVlot = Math.Sqrt(stUserData.dRated * stREG(rn).dblNOM) * stUserData.dMagnification
                    DebugLogOut("印加電圧[" & dAddVlot.ToString & "]= Sqrt(" & stUserData.dRated.ToString & " * " & stREG(rn).dblNOM.ToString & ") * " & stUserData.dMagnification.ToString)

                    dMaxCurrent = dAddVlot / stREG(rn).dblNOM * stUserData.dResNumber * stUserData.dCurrentLimit
                    DebugLogOut("最大印加電流[" & dMaxCurrent.ToString & "]= " & dAddVlot.ToString & " / " & stREG(rn).dblNOM.ToString & " * " & stUserData.dResNumber.ToString & " * " & stUserData.dCurrentLimit.ToString)
                    Exit For
                End If
            Next

            If dMaxCurrent = Double.MaxValue Then
                Call Z_PRINT("最大印加電流算出対象抵抗値が有りませんでした" & vbCrLf)
                Return (cFRS_ERR_RST)
            Else
                'Call Z_PRINT("最大印加電流=[" & dMaxCurrent.ToString & "]" & vbCrLf)
            End If

            Gno = Func_V_On_Number(rn)
            If Gno = 0 Then
                Call Z_PRINT("電源の設定が有りませんでした" & vbCrLf)
                Return (cFRS_ERR_RST)
            Else
                'r = GPIB_Cmd_Send(Gno, "")
            End If


            ' 電流源のＯＮコマンドを書き換える。
            Dim sCmd As String
            strMSG = stGPIB(Gno).strCON
            Dim sPos As Integer, ePos As Integer
            Dim bStart As Boolean

            sCmd = "VOLT"
            sPos = strMSG.IndexOf(sCmd)
            bStart = False
            If sPos >= 0 Then
                For ePos = sPos To strMSG.Length - 1
                    If (Char.IsNumber(strMSG(ePos))) Or strMSG(ePos) = "." Or strMSG(ePos) = "-" Then
                        If Not bStart Then
                            sPos = ePos     ' 数字の始まり
                            bStart = True
                        End If
                    Else
                        If bStart Then
                            Exit For        ' ePos は、数字の終わり
                        End If
                    End If
                Next
                strMSG = strMSG.Substring(0, sPos) + dAddVlot.ToString("0.0000") + strMSG.Substring(ePos)
            Else
                Call Z_PRINT("GPIBのONコマンドにVOLTの設定がありません。" & vbCrLf)
                Return (cFRS_ERR_RST)
            End If

            sCmd = "CURR"
            sPos = strMSG.IndexOf(sCmd)
            bStart = False
            If sPos >= 0 Then
                For ePos = sPos To strMSG.Length - 1
                    If (Char.IsNumber(strMSG(ePos))) Or strMSG(ePos) = "." Or strMSG(ePos) = "-" Then
                        If Not bStart Then
                            sPos = ePos     ' 数字の始まり
                            bStart = True
                        End If
                    Else
                        If bStart Then
                            Exit For        ' ePos は、数字の終わり
                        End If
                    End If
                Next
                strMSG = strMSG.Substring(0, sPos) + dMaxCurrent.ToString("0.0000") + strMSG.Substring(ePos)
            Else
                Call Z_PRINT("GPIBのONコマンドにCURRの設定がありません。" & vbCrLf)
                Return (cFRS_ERR_RST)
            End If

            stGPIB(Gno).strCON = strMSG


            ' ﾎﾟｰｽﾞ付きﾌﾟﾛｰﾌﾞON(ZﾌﾟﾛｰﾌﾞをON位置(Z.ZON)に移動)
            r = Prob_On()                                               ' ﾌﾟﾛｰﾌﾞON
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
                rtn = r                                                 ' Return値設定
                GoTo POWER_End
            End If

            If FUNC_OK = Func_V_On_Judge(iRnoNo) Then               '   ＯＮ機器有りかのチェック
                r = Func_V_On_Ex(iRnoNo)                            '   全ての対象のＯＮ機器をＯＮする。
                If (FUNC_NG = r) Then
                    GoTo POWER_End
                End If
            End If

            Call Z_PRINT("印加電圧=[" & dAddVlot.ToString("0.0000") & "] 印加電流=[" & dMaxCurrent.ToString("0.0000") & "]" & vbCrLf)

            StopWatch.Start()
            Do
                NowTime = StopWatch.ElapsedMilliseconds

                ' セーフティチェック
                r = SafetyCheck()                                       ' セーフティチェック
                If (r <> cFRS_NORMAL) Then                              ' エラー ?
                    rtn = r                                             ' Return値 = セーフティチェックエラー
                    GoTo POWER_End
                End If

                rtn = MEAS_DMM(Gno, dblMx)
                If rtn <> cFRS_NORMAL Or dblMx > dMaxCurrent Then
                    r = V_Off()                                                 ' DC電源装置 電圧OFF
                    r = Prob_Off()                                              ' Z2/ZﾌﾟﾛｰﾌﾞをOFF位置に移動
                    Buzzer()
                    If rtn <> cFRS_NORMAL Then
                        strMSG = "モニター表示電流値が読み取れませんでした。"
                    Else
                        strMSG = "最大印加電流エラー[" & dblMx.ToString("0.000") & "] > [" & dMaxCurrent.ToString("0.000") & "]"
                    End If
                    FrmMessageDisp(ObjSys, cGMODE_MSG_DSP, cFRS_ERR_START, True, _
                            strMSG, "処理を中断します", "", System.Drawing.Color.Red, System.Drawing.Color.Red, System.Drawing.Color.Blue)
                    rtn = cFRS_ERR_RST
                    GoTo POWER_End
                End If

                If (Form1.System1.Sys_Err_Chk_EX(gSysPrm, APP_MODE_LOTCHG) <> cFRS_NORMAL) Then ' 非常停止等 ?
                    Call Form1.AppEndDataSave()                                       ' ｿﾌﾄ強制終了時のﾃﾞｰﾀ保存確認
                    Call Form1.AplicationForcedEnding()                               ' ｿﾌﾄ強制終了処理
                    End
                End If

            Loop While (NowTime < lAppliedSecond)

            If FUNC_OK = Func_V_Off_Judge(rn) Then                  '   電圧OFF有り？                                               ' DC電源装置 電圧OFF
                r = Func_V_Off_Ex(rn)                               '   電圧OFF
                If (FUNC_NG = r) Then
                    GoTo POWER_End
                End If
            End If

            ' 終了処理
POWER_End:
            StopWatch.Stop()
            r = V_Off()                                                 ' DC電源装置 電圧OFF
            r = Prob_Off()                                              ' Z2/ZﾌﾟﾛｰﾌﾞをOFF位置に移動
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
                rtn = r                                                 ' Return値設定(非常停止等)
            End If

            Return (rtn)                                                ' (注)エラー時のメッセージは表示済み

            ' トラップエラー発生時 
        Catch ex As Exception
            MsgBox("User.Power() TRAP ERROR = " + ex.Message)
            Return (cFRS_TRIM_NG)
        End Try
    End Function
    'V2.0.0.0②↑
#End Region

#Region "電圧測定/抵抗測定処理(内部測定器/外部測定器)"
    '''=========================================================================
    ''' <summary>
    ''' MEASUREコマンドのエラー情報出力
    ''' </summary>
    ''' <param name="Code">
    ''' MEASUREの戻り値
    ''' 0	 : 正常 (ERROR_SUCCESS)
    ''' 1	 : 設定ずみ (ERR_ALREADY_SET)
    ''' 220 : 各軸のリミット検出 (ERR_AXS_LIM)
    ''' 309 : 不正ポインタ (ERR_SYS_BADPOINTER)
    ''' 501 : 測定レンジ設定エラー：指定レンジ設定タイプなし (ERR_MEAS_RANGESET_TYPE)
    ''' 502 : 測定レンジ設定エラー：対象レンジなし (ERR_MEAS_SETRNG_NO)
    ''' 503 : 測定レンジ設定エラー：最小レンジ以下 (ERR_MEAS_SETRNG_LO)
    ''' 504 : 測定レンジ設定エラー：最大レンジ以上 (ERR_MEAS_SETRNG_HI)
    ''' 514 : 測定スキャナ設定エラー：不正スキャナ番号 (ERR_MEAS_SCANNER)
    ''' 515 : 測定スキャナ設定エラー：最小スキャナ番号以下 (ERR_MEAS_SCANNER_LO)
    ''' 520 : 測定範囲外：ショート-抵抗(0x6666)/電圧(0x3333)以下 (ERR_MEAS_SPAN_SHORT)
    ''' 521 : 測定範囲外：オーバー-抵抗(0xCCCC)/電圧(0x6666)以上 (ERR_MEAS_SPAN_OVER)
    ''' 523 : 測定回数指定エラー (ERR_MEAS_COUNT)
    ''' 524 : 測定時の設定モード（Mfsetモード）エラー (ERR_MEAS_SETMODE)
    ''' 525 : オートレンジ測定:定電流測定範囲オーバー（差電流領域） (ERR_MEAS_AUTORNG_OVER)
    ''' 528 : スキャナ設定完了タイムアウト (ERR_MEAS_SCANSET_TIMEOUT)
    ''' 803 : Qレートパラメータエラー：最小Qレート以下 (ERR_LSR_PARAM_QSW_LO)
    ''' 804 : Qレートパラメータエラー：最小Qレート以上 (ERR_LSR_PARAM_QSW_HI)
    ''' </param>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub MeasureError(ByVal Code As Short)
        Dim sMessage As String = "NO ERROR"
        Select Case (Code)
            Case ERR_AXS_LIM
                sMessage = "各軸のリミット検出"
            Case ERR_SYS_BADPOINTER
                sMessage = "不正ポインタ"
            Case ERR_MEAS_RANGESET_TYPE
                sMessage = "測定レンジ設定エラー：指定レンジ設定タイプなし"
            Case ERR_MEAS_SETRNG_NO
                sMessage = "測定レンジ設定エラー：対象レンジなし"
            Case ERR_MEAS_SETRNG_LO
                sMessage = "測定レンジ設定エラー：最小レンジ以下"
            Case ERR_MEAS_SETRNG_HI
                sMessage = "測定レンジ設定エラー：最大レンジ以上"
            Case ERR_MEAS_SCANNER
                sMessage = "測定スキャナ設定エラー：不正スキャナ番号"
            Case ERR_MEAS_SCANNER_LO
                sMessage = "測定スキャナ設定エラー：最小スキャナ番号以下"
            Case ERR_MEAS_SPAN_SHORT
                'sMessage = "測定範囲外：ショート-抵抗(0x6666)/電圧(0x3333)以下"
            Case ERR_MEAS_SPAN_OVER
                'sMessage = "測定範囲外：オーバー-抵抗(0xCCCC)/電圧(0x6666)以上"
            Case ERR_MEAS_COUNT
                sMessage = "測定回数指定エラー"
            Case ERR_MEAS_SETMODE
                sMessage = "測定時の設定モード（Mfsetモード）エラー"
            Case ERR_MEAS_AUTORNG_OVER
                sMessage = "オートレンジ測定:定電流測定範囲オーバー（差電流領域）"
            Case ERR_MEAS_SCANSET_TIMEOUT
                sMessage = "スキャナ設定完了タイムアウト"
            Case ERR_LSR_PARAM_QSW_LO
                sMessage = "Qレートパラメータエラー：最小Qレート以下"
            Case ERR_LSR_PARAM_QSW_HI
                sMessage = "Qレートパラメータエラー：最小Qレート以上"
            Case Else
                sMessage = "不明エラー = [" & Code.ToString & "]"
        End Select

        If sMessage <> "NO ERROR" Then
            Call Form1.System1.OperationLogging(gSysPrm, MSG_OPLOG_TRIMST, "内部測定エラー:" & sMessage)
            Call Z_PRINT("測定エラー [" & sMessage.ToString & "]" & vbCrLf)
        End If
    End Sub
    '''=========================================================================
    '''<summary>電圧測定/抵抗測定処理(内部測定器/外部測定器)</summary>
    '''<param name="V_R_FLG">   (INP) ｽﾛｰﾌﾟ(1:+V, 2:-V, 4:抵抗)</param>
    '''<param name="Inmes">     (INP) 内部/外部種別(0=内部測定器, 1以降=外部測定器番号)</param>
    '''<param name="dblMxVal">  (OUT) 測定値(V/Ω)</param>
    '''<param name="rn">        (INP) 抵抗番号(1～)</param>
    ''' <param name="TargetVal">(INP)目標値</param>
    '''<returns>cFRS_NORMAL   = 正常
    '''         cFRS_ERR_TRIM = 測定エラー(トリマエラー)
    '''</returns>
    '''=========================================================================
    Public Function V_R_MEAS(ByVal V_R_FLG As Short, ByVal Inmes As Short, ByRef dblMxVal As Double, ByVal rn As Short, ByVal TargetVal As Double) As Short

        Dim rtn As Short = cFRS_NORMAL                                  ' 戻値
        Dim strMSG As String                                            ' メッセージ編集域

        dblMxVal = Double.MaxValue

        Try
            ' 電圧測定を行う(内部測定器/外部測定器)
            If (V_R_FLG <> 4 And V_R_FLG <> 6) Then                     ' 測定モード = 電圧測定
                If (Inmes = 0) Then                                     ' 内部測定器
                    rtn = MEASURE(MEAS_MODE_VOLTAGE, MEAS_RNGSET_FIX_TAR, stREG(rn).intTMM1, TargetVal, 0, dblMxVal)       ' 抵抗測定(内部測定器)
                    If rtn <> cFRS_NORMAL Then
                        MeasureError(rtn)
                        rtn = cFRS_NORMAL                              ' エラーにしない。
                    End If
                Else
                    DMM = Inmes                                         ' 外部測定器番号
                    rtn = MEAS_DMM(DMM, dblMxVal)                         ' 電圧測定(外部測定器(ﾃﾞｼﾞﾀﾙﾏﾙﾁﾒｰﾀによる測定))
                    If rtn <> cFRS_NORMAL Then
                        GoTo STP_ERR
                    End If
                End If

                ' 抵抗測定を行う(内部測定器/外部測定器)
            Else
                If (Inmes = 0) Then                                     ' 内部測定器
                    rtn = MEASURE(MEAS_MODE_RESISTOR, MEAS_RNGSET_FIX_TAR, stREG(rn).intTMM1, TargetVal, 0, dblMxVal)       ' 抵抗測定(内部測定器)
                    If rtn <> cFRS_NORMAL Then
                        MeasureError(rtn)
                        rtn = cFRS_NORMAL                              ' エラーにしない。
                    End If
                Else
                    If UserSub.IsSpecialTrimType() Then
                        Call Change_Range_DMM(Inmes, TargetVal)
                    End If
                    DMM = Inmes                                         ' 外部測定器番号
                    rtn = MEAS_DMM(DMM, dblMxVal)                       ' 抵抗測定(外部測定器(ﾃﾞｼﾞﾀﾙﾏﾙﾁﾒｰﾀによる測定))
                    If rtn <> cFRS_NORMAL Then
                        GoTo STP_ERR
                    End If
                End If
            End If

            Return (cFRS_NORMAL)

STP_ERR:    ' 測定エラー
            Return (cFRS_ERR_TRIM)                                      ' Return値 = トリマエラー

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "UserBas.V_R_MEAS() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cFRS_TRIM_NG)
        End Try
    End Function
#End Region

#Region "電圧測定(内部測定器)"
    '''=========================================================================
    '''<summary>電圧測定(内部測定器)</summary>
    '''<param name="rn">   (INP) 抵抗番号</param>
    '''<param name="dblMx">(OUT) 測定値</param>
    '''<returns>cFRS_NORMAL   = 正常
    '''         上記以外      = エラー
    '''</returns>
    '''=========================================================================
    Public Function V_MEAS(ByRef rn As Short, ByRef dblMx As Double) As Short

        Dim rtn As Short = cFRS_NORMAL                                  ' 戻値
        Dim strMSG As String                                            ' メッセージ編集域
        Dim r As Short


        Try

#If cOFFLINEcDEBUG Then                                 ' DEBUG
            V_MEAS = 0                                      ' Return値 = 正常
            dblMx = 35                                      ' 測定値(V)
            Exit Function
#End If
            dblMx = 0
            r = VMEAS(Z1, Z0, dblMx)                        ' 電圧測定(内部測定器)

            'r = MFSET("V")
            'r = MEASURE(1, 0, 0, 0, 0, dblMx)               ' ' レンジ設定タイプ（0:オートレンジ、1:固定レンジ-目標値指定、2:固定レンジ-レンジ番号指定
            r = MEASURE(1, 1, 0, 3, 0, dblMx)               ' 抵抗測定(内部測定器)
            'r = MEASURE(1, 2, 0, 0, 14, dblMx)               ' 抵抗測定(内部測定器)
            'If (r <> 0) Then                                ' エラー ?
            '    V_MEAS = 1                                  ' Return値 = 測定エラー
            'End If

#If (cCND = 1) Then                                     ' 条件出しﾓｰﾄﾞ ?
            strMSG = "■電圧測定値(V)=" + dblMx.ToString("#0.0###")
            Call Z_PRINT(strMSG + vbCrLf)
#End If
            Return (rtn)

STP_ERR:
            strMSG = "電圧測定エラー(VMEAS) !! "
            Call Z_PRINT(strMSG & vbCrLf)
            Return (1)                                      ' Return値 = 測定エラー

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.V_MEAS() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cFRS_TRIM_NG)
        End Try
    End Function
#End Region

#Region "抵抗測定(内部測定器)"
    '''=========================================================================
    '''<summary>抵抗測定(内部測定器)</summary>
    '''<param name="dblMx">(OUT) 測定値(Ω)</param>
    '''<returns>cFRS_NORMAL   = 正常
    '''         上記以外      = エラー
    '''</returns>
    '''=========================================================================
    Public Function R_MEAS(ByRef dblMx As Double) As Short

        On Error GoTo STP_ERR
        Dim r As Integer
        Dim strMSG As String                            ' メッセージ編集域

#If cOFFLINEcDEBUG Then                                 ' DEBUG
        R_MEAS = 0                                      ' Return値 = 正常
        dblMx = 5.0#                                    ' 測定値
        Exit Function
#End If

        dblMx = 0
        'r = RMEAS(Z1, Z0, dblMx)                        ' 抵抗測定(内部測定器)
        'If (r <> 0) Then                                ' エラー ?
        '    R_MEAS = 1                                  ' Return値 = 測定エラー
        'End If

        'r = MFSET_EX("R", 1000)
        'r = MFSET("R")

        'Call MSCAN(ph, pl, ag1, ag2, ag3, ag4, ag5)     ' スキャナー番号設定
        'r = MEASURE(0, 1, 0, 1000, 0, dblMx) ' 1:固定レンジ-目標値指定

        ' 抵抗測定(内部測定器)レンジ設定タイプ（0:オートレンジ)
        r = MEASURE(0, 0, 0, 0, 0, dblMx)               ' 抵抗測定(内部測定器)
        'If (r <> 0) Then                                ' エラー ?
        '    R_MEAS = 1                                  ' Return値 = 測定エラー
        'End If


#If (cCND = 1) Then                                     ' 条件出しﾓｰﾄﾞ ?
        strMSG = "■抵抗測定値(V)=" + dblMx.ToString("#0.0###")
        Call Z_PRINT(strMSG + vbCrLf)
#End If
        Exit Function

STP_ERR:
        strMSG = "抵抗測定エラー(RMEAS) !! "
        Call Z_PRINT(strMSG & vbCrLf)
        R_MEAS = 1                                      ' Return値 = 測定エラー

    End Function
#End Region

#Region "電圧測定(ﾃﾞｼﾞﾀﾙﾏﾙﾁﾒｰﾀによる測定)"
    '''=========================================================================
    ''' <summary>電圧測定(ﾃﾞｼﾞﾀﾙﾏﾙﾁﾒｰﾀによる測定)</summary>
    ''' <param name="Gno">(INP) GPIBﾃﾞｰﾀｲﾝﾃﾞｯｸｽ(1～)</param>
    ''' <param name="VX"> (OUT) 測定値(V)</param>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          上記以外      = エラー
    ''' </returns>
    '''=========================================================================
    Private Function MEAS_DMM(ByRef Gno As Short, ByRef VX As Double) As Short

        Dim rtn As Short = cFRS_NORMAL                                  ' 戻値
        Dim r As Integer
        Dim rFlg As Short
        Dim strMSG As String                                            ' メッセージ編集域
        Dim dblVX As Double                                             ' 入力電圧値

        Try
            ' 初期処理
            rFlg = 0
            VX = 0                                                      ' 電圧値初期化
            If (Gno = -1) Then Return (cFRS_NORMAL) '                   ' TRIGｺﾏﾝﾄﾞ指定なしならNOP

            ' 測定コマンド送信
STP_RETRY:

            'V2.2.1.4①↓
            If sStrTrig <> "" Then
                strMSG = sStrTrig                                ' トリガーコマンド送信("READ?"/"MEASure:VOLTage:DC?"等)
            Else
                strMSG = stGPIB(Gno).strCTRG                                ' トリガーコマンド送信("READ?"/"MEASure:VOLTage:DC?"等)
            End If
            'V2.2.1.4①↑
            r = ObjGpib.Gpib_Send(strMSG, gDevId, stGPIB(Gno).intGAD, stGPIB(Gno).intDLM, gEOI)
            If (r <> cFRS_NORMAL) Then GoTo MEAS_DMM_Err
            Call ZWAIT(10)

            ' 電圧値入力→ dblVx
            r = ObjGpib.Gpib_RVal(dblVX, gDevId, stGPIB(Gno).intGAD, stGPIB(Gno).intDLM, gEOI)
            'r = ObjGpib.Gpib_Recv(strMSG, gDevId, stGPIB(Gno).intGAD, stGPIB(Gno).intDLM, gEOI)
            If (r <> cFRS_NORMAL) Then GoTo MEAS_DMM_Err

            VX = dblVX                                                  ' 電圧値設定[V]
#If (cCND = 1) Then                                                     ' DEBUG MODE ?
            strMSG = "■測定値=" + dblVX.ToString("#0.00000")
            Call Z_PRINT(strMSG + vbCrLf)
#End If
            Return (rtn)                                                ' 正常リターン

MEAS_DMM_Err:
            rFlg = rFlg + 1
            If (rFlg <= 2) Then GoTo STP_RETRY
            strMSG = " GP-IB通信タイムアウトエラー(マルチメーター) !! "
            Call Z_PRINT(strMSG & vbCrLf)
            Return (1)                                                  ' Return値 = 測定エラー

STP_ERR:
            strMSG = "測定エラー(MEAS_DMM) !! "
            Call Z_PRINT(strMSG & vbCrLf)
            Return (1)                                                  ' Return値 = 測定エラー

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.MEAS_DMM() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (1)                                                  ' Return値 = 測定エラー

        Finally
            sStrTrig = ""       'V2.2.1.4①
        End Try

    End Function
#End Region

    '==========================================================================
    '   ＧＰＩＢ制御処理
    '==========================================================================
#Region "GPIB初期設定処理"
    '''=========================================================================
    '''<summary>GPIB初期設定処理(GP-IB n台)</summary>
    '''<returns>cFRS_NORMAL   = 正常
    '''         上記以外      = その他エラー
    '''</returns>
    '''=========================================================================
    Public Function GPIB_Init() As Short

        Dim strMSG As String                                            ' メッセージ編集域
        Dim strDBG As String = ""                                       ' メッセージ編集域(ﾃﾞﾊﾞｯｸﾞ用)
        Dim i As Short                                                  ' Index
        Dim WkGno As Integer = -1
        Dim r As Integer                                                ' 関数戻値

        Try
            ' 初期処理
            GPIB_Init = cFRS_NORMAL                                     ' Return値 = 正常
            DMM = -1                                                    ' ﾃﾞｼﾞﾀﾙﾏﾙﾁﾒｰﾀ Index

            ' ＧＰＩＢ初期化
            If (stPLT.GCount <= 0) Then Exit Function '                 ' GPIB制御なしならNOP
            If (FlgGPIB = False) Then                                   ' GPIB初期化Flag OFF ?
                r = ObjGpib.Gpib_Init(gstrDeviceName, gDevId)           ' GPIB初期化(デバイスＩＤ取得)
                FlgGPIB = True                                          ' GPIB初期化Flag ON
            End If

            ' ＧＰＩＢ初期設定
            strMSG = "=== GPIB 初期化 ===" & vbCrLf
            Call Z_PRINT(strMSG)
            For i = 1 To stPLT.GCount                                   ' GPIB制御数分繰返す
                strMSG = "  初期化(ﾃﾞﾊﾞｲｽ番号=" & i.ToString("0") & ", 名称=" & stGPIB(i).strGNAM & ") ... "
                Call Z_PRINT(strMSG)

                ' TRIGｺﾏﾝﾄﾞ指定ありのもの(最初に見つかったもの)をﾃﾞｼﾞﾀﾙﾏﾙﾁﾒｰﾀ(ＲＥＡＤ系)としてIndexを設定する
                If (DMM = -1) Then                                      ' ﾃﾞｼﾞﾀﾙﾏﾙﾁﾒｰﾀ Index
                    If (stGPIB(i).strCTRG <> "") Then                   ' TRIGｺﾏﾝﾄﾞ指定あり？
                        DMM = i                                         ' ﾃﾞｼﾞﾀﾙﾏﾙﾁﾒｰﾀ Index設定
                    End If
                End If

#If (cDBG = 1) Then                                                     ' ﾃﾞﾊﾞｯｸﾞ ?
                Call Msg_Disp(vbCrLf, 0)
#End If
                'V2.0.0.0④ strCCMDからstrCCMD1,strCCMD2,strCCMD3へ変更
                ' 設定コマンド送信(設定コマンドなし又はGPIBアドレスが同じものは送信しない)
                If (stGPIB(i).strCCMD1 = "") Or (stGPIB(i).intGAD = WkGno) Then
                    strMSG = "OK" & vbCrLf
                    Call Z_PRINT(strMSG)
                Else
                    WkGno = stGPIB(i).intGAD                            ' GPIBアドレス退避

                    'V2.2.1.4①　ChangeHIOKICommand(stGPIB(i).strCCMD1, i, stREG(1).dblNOM)      'V2.2.1.1⑤

                    strMSG = stGPIB(i).strCCMD1                          ' 設定コマンド
#If (cDBG = 1) Then                                                     ' ﾃﾞﾊﾞｯｸﾞ ?
                    strDBG = "(DBG) GPIB 送信データ→" + """" + strMSG + """"
                    Call Msg_Disp(strDBG, 0)
#End If
                    ' 文字列変換

                    r = ObjGpib.Gpib_Send(strMSG, gDevId, stGPIB(i).intGAD, stGPIB(i).intDLM, gEOI)
                    If (r <> cFRS_NORMAL) Then GoTo GPIB_ERR

                    'V2.0.0.0④↓
                    If stGPIB(i).strCCMD2 <> "" Then
                        'V2.2.1.4①　ChangeHIOKICommand(stGPIB(i).strCCMD2, i, stREG(1).dblNOM)      'V2.2.1.1⑤
                        r = ObjGpib.Gpib_Send(stGPIB(i).strCCMD2, gDevId, stGPIB(i).intGAD, stGPIB(i).intDLM, gEOI)
                        If (r <> cFRS_NORMAL) Then GoTo GPIB_ERR
                    End If

                    If stGPIB(i).strCCMD3 <> "" Then
                        'V2.2.1.4①　ChangeHIOKICommand(stGPIB(i).strCCMD3, i, stREG(1).dblNOM)      'V2.2.1.1⑤
                        r = ObjGpib.Gpib_Send(stGPIB(i).strCCMD3, gDevId, stGPIB(i).intGAD, stGPIB(i).intDLM, gEOI)
                        If (r <> cFRS_NORMAL) Then GoTo GPIB_ERR
                    End If
                    'V2.0.0.0④↑

                    strMSG = "OK" & vbCrLf
                    Call Z_PRINT(strMSG)
                End If
            Next i

            Return (cFRS_NORMAL)

            ' GPIB初期化エラー
GPIB_ERR:
            strMSG = "GPIB 初期化エラー"
            Call Z_PRINT(strMSG & vbCrLf)
            Return (1)                                                  ' Return値 = エラー

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "NG = " + ex.Message + vbCrLf
            MsgBox(strMSG)
            Call Z_PRINT(strMSG)
            Return (1)                                                  ' Return値 = エラー
        End Try
    End Function
#End Region

#Region "DC電源装置 電圧ON処理"
    '''=========================================================================
    '''<summary>DC電源装置 電圧ON処理(+nV)</summary>
    '''<returns>cFRS_NORMAL   = 正常
    '''         上記以外      = その他エラー
    '''</returns>
    '''=========================================================================
    Public Function V_On() As Short

        Dim strMSG As String                                            ' メッセージ編集域
        Dim i As Short                                                  ' Index
        Dim r As Short

        Try
            For i = 1 To stPLT.GCount                                   ' GPIB制御数分繰返す
                r = GPIB_On(i)                                          ' 電圧ONｺﾏﾝﾄﾞ送信(ﾜｰｸ印加用)
                If (r <> 0) Then GoTo STP_ERR
            Next i

            Return (cFRS_NORMAL)                                        ' 正常リターン

STP_ERR:
            strMSG = "ﾜｰｸ印加用電圧ONエラー(" & stGPIB(i).strGNAM & ") !! "
            Call Msg_Disp(strMSG, 1)                                    ' ﾒｯｾｰｼﾞ表示(ﾛｸﾞ画面)/印字/ブザーON
            Return (1)                                                  ' Return値 = エラー

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.V_On() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (1)
        End Try
    End Function
#End Region

#Region "DC電源装置 電圧OFF処理"
    '''=========================================================================
    '''<summary>DC電源装置 電圧OFF処理(+nV)</summary>
    '''<returns>cFRS_NORMAL   = 正常
    '''         上記以外      = その他エラー
    '''</returns>
    '''=========================================================================
    Public Function V_Off() As Short

        Dim strMSG As String                                            ' メッセージ編集域
        Dim i As Short                                                  ' Index
        Dim r As Short

        Try
            For i = 1 To stPLT.GCount                                   ' GPIB制御数分繰返す
                r = GPIB_Off(i)                                         ' 電圧OFFｺﾏﾝﾄﾞ送信(ﾜｰｸ印加用)
                If (r <> 0) Then GoTo STP_ERR
            Next i
            Return (cFRS_NORMAL)                                        ' 正常リターン

STP_ERR:
            strMSG = "ﾜｰｸ印加用電圧OFFエラー(" & stGPIB(i).strGNAM & ") !! "
            Call Msg_Disp(strMSG, 1)                                    ' ﾒｯｾｰｼﾞ表示(ﾛｸﾞ画面)/印字/ブザーON
            Return (1)                                                  ' Return値 = エラー

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.V_Off() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (1)                                                  ' Return値 = エラー
        End Try
    End Function
#End Region

#Region "■■　DC電源装置 電圧ON処理(GP-IB機器番号指定)　■■"
    '''=========================================================================
    ''' <summary>DC電源装置 電圧ON処理(+nV)</summary>
    ''' <param name="Gno">(INP)GP-IB機器番号</param>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          上記以外      = その他エラー
    ''' </returns>
    '''=========================================================================
    Public Function V_On_Ex(ByVal Gno As Short) As Short

        Dim strMSG As String                                            ' メッセージ編集域
        Dim i As Short                                                  ' Index
        Dim r As Short

        Try
            r = GPIB_On(Gno)                                            ' 電圧ONｺﾏﾝﾄﾞ送信(ﾜｰｸ印加用)
            If (r <> 0) Then GoTo STP_ERR
            Return (cFRS_NORMAL)                                        ' 正常リターン

STP_ERR:
            strMSG = "ﾜｰｸ印加用電圧ONエラー(" & stGPIB(i).strGNAM & ") !! "
            Call Msg_Disp(strMSG, 1)                                    ' ﾒｯｾｰｼﾞ表示(ﾛｸﾞ画面)/印字/ブザーON
            Return (1)                                                  ' Return値 = エラー

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.V_On_Ex() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (1)
        End Try
    End Function
#End Region

#Region "■■　DC電源装置 電圧OFF処理(GP-IB機器番号指定)　■■"
    '''=========================================================================
    ''' <summary>DC電源装置 電圧OFF処理(+nV)</summary>
    ''' <param name="Gno">(INP)GP-IB機器番号</param>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          上記以外      = その他エラー
    ''' </returns>
    '''=========================================================================
    Public Function V_Off_Ex(ByVal Gno As Short) As Short

        Dim strMSG As String                                            ' メッセージ編集域
        Dim i As Short                                                  ' Index
        Dim r As Short

        Try
            r = GPIB_Off(Gno)                                         ' 電圧OFFｺﾏﾝﾄﾞ送信(ﾜｰｸ印加用)
            If (r <> 0) Then GoTo STP_ERR
            Return (cFRS_NORMAL)                                        ' 正常リターン

STP_ERR:
            strMSG = "ﾜｰｸ印加用電圧OFFエラー(" & stGPIB(i).strGNAM & ") !! "
            Call Msg_Disp(strMSG, 1)                                    ' ﾒｯｾｰｼﾞ表示(ﾛｸﾞ画面)/印字/ブザーON
            Return (1)                                                  ' Return値 = エラー

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.V_Off_Ex() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (1)                                                  ' Return値 = エラー
        End Try
    End Function
#End Region

#Region "ＯＮコマンドをGP-IB機器へ送信する"
    '''=========================================================================
    '''<summary>ＯＮコマンドをGP-IB機器へ送信する</summary>
    '''<param name="gn">(INP) GPIB設定用データの添字(1 ORG)</param>
    '''<returns>cFRS_NORMAL   = 正常
    '''         上記以外      = その他エラー
    '''</returns>
    '''=========================================================================
    Public Function GPIB_On(ByRef gn As Short) As Short

        Dim strMSG As String                                            ' メッセージ編集域
        Dim strDBG As String = ""                                       ' メッセージ編集域(ﾃﾞﾊﾞｯｸﾞ用)
        Dim r As Short                                                  ' 戻り値

        Try
            ' ＯＮコマンドを設定する
            If (gn = 0) Then Exit Function
            strMSG = stGPIB(gn).strCON                                  ' ＯＮコマンド設定

            ' ＯＮコマンド無しならEXIT
            If (strMSG = "") Then Return (cFRS_NORMAL)

#If (cDBG = 1) Then                                                     ' ﾃﾞﾊﾞｯｸﾞ ?
            strDBG = "(DBG) GPIB 送信データ→" + """" + strMSG + """"
            Call Msg_Disp(strDBG, 0)
#End If

            ' ＯＮコマンド送信
            r = ObjGpib.Gpib_Send(strMSG, gDevId, stGPIB(gn).intGAD, stGPIB(gn).intDLM, gEOI)
            If (r <> cFRS_NORMAL) Then GoTo GPIB_On_ERR
            Call ZWAIT(stGPIB(gn).lngPOWON)                             ' ON後ﾎﾟｰｽﾞ(ms)
            Return (cFRS_NORMAL)                                        ' 正常リターン

GPIB_On_ERR:
            strMSG = " ＯＮコマンド送信エラー!! "
            Call Msg_Disp(strMSG, 1)                                    ' ﾒｯｾｰｼﾞ表示(ﾛｸﾞ画面)/印字/ブザーON
            Return (1)                                                  ' Return値 = エラー

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.GPIB_On() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (1)                                                  ' Return値 = エラー
        End Try
    End Function
#End Region

#Region "ＯＦＦコマンドをGP-IB機器へ送信する"
    '''=========================================================================
    '''<summary>ＯＦＦコマンドをGP-IB機器へ送信する</summary>
    '''<param name="gn">(INP) GPIB設定用データの添字(1 ORG)</param>
    '''<returns>cFRS_NORMAL   = 正常
    '''         上記以外      = その他エラー
    '''</returns>
    '''=========================================================================
    Public Function GPIB_Off(ByRef gn As Short) As Short

        Dim strMSG As String                                            ' メッセージ編集域
        Dim strDBG As String = ""                                       ' メッセージ編集域(ﾃﾞﾊﾞｯｸﾞ用)
        Dim r As Integer                                                ' 戻り値

        Try
            ' ＯＦＦコマンドを設定する
            Call ZWAIT(200)                                             ' Wait(ms)
            If (gn = 0) Then Return (cFRS_NORMAL)
            strMSG = stGPIB(gn).strCOFF                                 ' ＯＦＦコマンド設定

            ' ＯＦＦコマンド無しならEXIT
            If (strMSG = "") Then Return (cFRS_NORMAL)
#If (cDBG = 1) Then                                                     ' ﾃﾞﾊﾞｯｸﾞ ?
            strDBG = "(DBG) GPIB 送信データ→" + """" + strMSG + """"
            Call Msg_Disp(strDBG, 0)
#End If
            ' ＯＦＦコマンドを送信する
            r = ObjGpib.Gpib_Send(strMSG, gDevId, stGPIB(gn).intGAD, stGPIB(gn).intDLM, gEOI)
            If (r <> cFRS_NORMAL) Then GoTo GPIB_Off_ERR
            Call ZWAIT(stGPIB(gn).lngPOWOFF)                            ' OFF後ﾎﾟｰｽﾞ(ms)
            Return (cFRS_NORMAL)                                        ' 正常リターン

GPIB_Off_ERR:
            strMSG = " ＯＦＦコマンド送信エラー!! "
            Call Msg_Disp(strMSG, 1)                                    ' ﾒｯｾｰｼﾞ表示(ﾛｸﾞ画面)/印字/ブザーON
            Return (1)                                                  ' Return値 = エラー

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.GPIB_Off() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (1)                                                  ' Return値 = エラー
        End Try
    End Function
#End Region

#Region "コマンドをGP-IB機器へ送信する"
    '''=========================================================================
    '''<summary>ＯＦＦコマンドをGP-IB機器へ送信する</summary>
    '''<param name="gn">    (INP) GPIB設定用データの添字(1 ORG)</param>
    '''<param name="strDAT">(INP) 送信コマンド</param>
    '''<returns>cFRS_NORMAL   = 正常
    '''         上記以外      = その他エラー
    '''</returns>
    '''=========================================================================
    Public Function GPIB_Cmd_Send(ByRef gn As Short, ByRef strDAT As String) As Short

        Dim strMSG As String                                            ' メッセージ編集域
        Dim strDBG As String = ""                                       ' メッセージ編集域(ﾃﾞﾊﾞｯｸﾞ用)
        Dim r As Integer                                                ' 戻り値

        Try
            ' コマンドを設定する
            If (gn = 0) Then Return (cFRS_NORMAL)

            ' コマンド無しならEXIT
            If (strDAT = "") Then Return (cFRS_NORMAL)
#If (cDBG = 1) Then                                                     ' ﾃﾞﾊﾞｯｸﾞ ?
            strDBG = "(DBG) GPIB 送信データ→" + """" + strDAT + """"
            Call Msg_Disp(strDBG, 0)
#End If
            ' コマンド送信
            r = ObjGpib.Gpib_Send(strDAT, gDevId, stGPIB(gn).intGAD, stGPIB(gn).intDLM, gEOI)
            If (r <> cFRS_NORMAL) Then GoTo STP_ERR
            Return (cFRS_NORMAL)                                        ' 正常リターン

STP_ERR:
            strMSG = "設定コマンド送信エラー!! "
            Call Msg_Disp(strMSG, 1)                                    ' ﾒｯｾｰｼﾞ表示(ﾛｸﾞ画面)/印字/ブザーON
            Return (1)                                                  ' Return値 = エラー

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.GPIB_Cmd_Send() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (1)                                                  ' Return値 = エラー
        End Try
    End Function
#End Region

    '==========================================================================
    '   カッティングチェック処理
    '==========================================================================
#Region "カッティングチェック処理"
    '''=========================================================================
    '''<summary>カッティングチェック処理</summary>
    '''<returns>cFRS_NORMAL   = 正常
    '''         cFRS_ERR_RST  = RESET SW押下
    '''         cFRS_TRIM_NG  = トリミングNG
    '''         cFRS_ERR_PTN  = パターン認識エラー
    '''         上記以外      = その他エラー
    '''</returns>
    '''=========================================================================
    Public Function CUT_CHK() As Short

        Dim rtn As Short = cFRS_NORMAL                                  ' 戻値
        Dim rn As Short                                                 ' 抵抗番号
        Dim cn As Short                                                 ' カット番号
        Dim r As Short                                                  ' 戻値
        Dim Flg As Short                                                ' flg
        Dim dblQrate As Double
        Dim dblQrate2 As Double                                         ' （1:リターン, 2:リトレース用）
        Dim strMSG As String
        Dim CutMode As Short                                            'V1.0.4.3⑧

        Try
            '-------------------------------------------------------------------
            '   初期処理
            '-------------------------------------------------------------------
            ' セーフティチェック
            r = SafetyCheck()                                           ' セーフティチェック
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
                Return (r)                                              ' Return値 = セーフティチェックエラー
            End If

            ' パターン認識処理
            r = Ptn_Match_Exe()                                         ' パターン認識実行
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
                If (r = cFRS_ERR_PTN) Then                              ' パターン認識エラー
                    Return (r)                                          ' Return値設定
                ElseIf (r = cFRS_ERR_RST) Then                          ' キャンセル ?
                    Return (cFRS_ERR_RST)                               ' Return値 = REST SW押下
                Else
                    Return (r)                                          ' Return値設定
                End If
            End If

            '-------------------------------------------------------------------
            '   カッティングチェック処理
            '-------------------------------------------------------------------
            Call ObjSys.Ilum_Ctrl(gSysPrm, Z1, ZOPT)                    ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ点灯(制御有時有効)
            Call Disp_Init()                                            ' 見出し表示(ログ画面)/印刷

            For rn = 1 To stPLT.RCount                                  ' 抵抗数分繰返す

                If UserModule.IsCutResistorIncMarking(rn) Then          'V2.0.0.0⑮

                    'V2.0.0.0⑮                If Not IsCutResistorIncCharacter(rn) Then               ' トリミング抵抗以外の時
                    'V2.0.0.0⑮                    GoTo STP_NEXT
                    'V2.0.0.0⑮                End If
                    Flg = 1
                    ' カット数分繰返す
                    For cn = 1 To stREG(rn).intTNN
                        strMSG = "R = " & rn.ToString & " (" & stREG(rn).strRNO.ToString & ") " & ", C = " & cn.ToString & vbCrLf
                        Call Z_PRINT(strMSG)                                ' ログ画面に表示(抵抗No.,カットNo.)

                        InitCutParam(cutCmnPrm)
#If cOSCILLATORcFLcUSE Then
                        ' FL時は加工条件番号テーブルからQレートを設定する(カットスピードはデータから設定)
                        dblQrate = stCND.Freq(stREG(rn).STCUT(cn).intCND(CUT_CND_L1))
                        dblQrate2 = stCND.Freq(stREG(rn).STCUT(cn).intCND(CUT_CND_L2))
#Else
                        dblQrate = stREG(rn).STCUT(cn).intQF1
                        dblQrate = dblQrate / 10.0
                        dblQrate2 = stREG(rn).STCUT(cn).intQF2
                        dblQrate2 = dblQrate2 / 10.0
#End If
                        ' ＢＰ移動(カット位置補正あり)
                        'V2.0.0.0⑮                    If (stREG(rn).STCUT(cn).intCUT <> CNS_CUTM_NON_POS_IX) Then           ' ﾎﾟｼﾞｼｮﾆﾝｸﾞ無しIXｶｯﾄはBP移動なし
                        Call STRXY(rn, stREG(rn).STCUT(cn).dblSTX, stREG(rn).STCUT(cn).dblSTY)
                        Call ObjSys.WAIT(0.5)
                        'V2.0.0.0⑮                    End If

                        MoveStop()              'V2.2.0.0⑥ 

                        ' '' 抵抗毎に停止する(ADJ ON時)
                        ''If (Flg = 1) Then                       ' 抵抗毎に停止
                        ''    r = Form1.System1.HALT2(3)             ' ADV(1)/HALT(2)/RESET(3)待ち(ADJ ON時)
                        ''    If (r = cFRS_ERR_RST) Then          ' RESET SW押下 ?
                        ''        Return(r)                     ' Return値 = RESET SW押下
                        ''    End If
                        ''    If (r < cFRS_NORMAL) Then           ' エラー ?
                        ''        Return(r)                     ' Return値設定
                        ''    End If
                        ''End If
                        Flg = 0

                        ' セーフティチェック
                        r = SafetyCheck()                                   ' セーフティチェック
                        If (r <> cFRS_NORMAL) Then                          ' エラー ?
                            Return (r)                                      ' Return値 = セーフティチェックエラー
                        End If

                        ' パターン認識NGの抵抗はSKIP
                        If (gTblPtn(rn) = 1) Then                           ' パターン認識NG ?
                            strMSG = "ﾊﾟﾀｰﾝ認識NGの為SKIP(R=" & rn.ToString("00") & ")" & vbCrLf
                            Call Z_PRINT(strMSG)                            ' ログ画面に表示

                            SetLotMarkAlarm(gsDataFileName, MarkingCount)           ' エラーとなった基板の情報を保存   'V2.2.1.7③

                            GoTo STP_NEXT
                        End If

                        'V2.0.0.0⑮                    dblQrate = stREG(rn).STCUT(cn).intQF1
                        'V2.0.0.0⑮                    'dblQrate = dblQrate / 10.0

                        ' カット処理
                        'Call QRATE(stREG(rn).STCUT(cn).intQF1)             ' Q-RATE設定
                        Select Case (stREG(rn).STCUT(cn).intCTYP)           ' ｶｯﾄ形状(1:ｽﾄﾚｰﾄ, 2:Lｶｯﾄ)
                            Case CNS_CUTP_ST, CNS_CUTP_ST_TR ' STカット(斜め直線カット※ﾘﾐｯﾄ長分ｶｯﾄ)
                                'V1.0.4.3⑧↓
                                If stREG(rn).STCUT(cn).intCTYP = CNS_CUTP_ST Then
                                    CutMode = CNS_CUTP_NORMAL
                                Else
                                    CutMode = CNS_CUTP_ST_TR
                                End If
                                'V1.0.4.3⑧↑
                                'V1.0.4.3⑧                            rtn = TrimSt(FORCE_MODE, 0, 0, stREG(rn).intSLP, stREG(rn).STCUT(cn).intANG, stREG(rn).STCUT(cn).dblDL2, stREG(rn).STCUT(cn).dblV1, stREG(rn).STCUT(cn).dblV1, dblQrate, dblQrate, 0, 0)
                                'V2.0.0.0⑮                                rtn = TrimSt(FORCE_MODE, CutMode, 0, stREG(rn).intSLP, stREG(rn).STCUT(cn).intANG, stREG(rn).STCUT(cn).dblDL2, stREG(rn).STCUT(cn).dblV1, stREG(rn).STCUT(cn).dblV1, dblQrate, dblQrate, 0, 0, stREG(rn).STCUT(cn).dblSX2, stREG(rn).STCUT(cn).dblSY2)
                                rtn = TrimSt(FORCE_MODE, CutMode, 0, SLP_RTRM, stREG(rn).STCUT(cn).intANG, stREG(rn).STCUT(cn).dblDL2, stREG(rn).STCUT(cn).dblV1, stREG(rn).STCUT(cn).dblV2, dblQrate, dblQrate2, 0, 0, stREG(rn).STCUT(cn).dblSX2, stREG(rn).STCUT(cn).dblSY2)

                                If stREG(rn).STCUT(cn).intCTYP = CNS_CUTP_ST_TR Then        'V2.0.0.0⑦ リトレースカット
                                    MoveStop()              'V2.2.0.0⑥ 
                                    rtn = CUT_RETRACE(rn, cn, stREG(rn).STCUT(cn).dblDL2)   'V2.0.0.0⑦
                                End If                                                      'V2.0.0.0⑦

                            Case CNS_CUTP_L ' Lカット
                                rtn = CUT_L(rn, cn)                         ' 斜めLカット(直線カット) ※ﾘﾐｯﾄ長分ｶｯﾄ

                            Case CNS_CUTP_SP ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ
                                rtn = CUT_SPT(rn, cn)                       ' 斜めSPカット(直線カット) ※ﾘﾐｯﾄ長分ｶｯﾄ

                            Case CNS_CUTP_IX ' ｲﾝﾃﾞｯｸｽｶｯﾄ
                                rtn = CUT_IX2(rn, cn)                       ' 斜めIXカット(直線カット) ※ﾘﾐｯﾄ長分ｶｯﾄ
                                'V1.0.4.3⑤↓
                            Case CNS_CUTP_M '文字マーキング
                                'V2.2.1.7③↓
                                If stREG(rn).intSLP = SLP_MARK Then
                                    ' マーク印字
                                    Dim retry_marking As Integer = 0
                                    Dim MarkStr As String = ""
                                    If (stREG(rn).STCUT(cn).cMarkStartNum <> "") Then
                                        Dim len As Integer = stREG(rn).STCUT(cn).cMarkStartNum.Length
                                        MarkStr = (Integer.Parse(stREG(rn).STCUT(cn).cMarkStartNum) + MarkingCount - 1).ToString.PadLeft(len, "0"c)
                                        If MarkStr.Length > 6 Then
                                            ' 数字カウンタ部の文字数は最大６桁までなので、それを超えた場合マーク印字を行わない 
                                            SetLotMarkAlarm(gsDataFileName, MarkingCount)           ' エラーとなった基板の情報を保存 
                                            GoTo STP_NEXT
                                        End If
                                    End If
                                    MarkStr = stREG(rn).STCUT(cn).cMarkFix & MarkStr
RETRY_MARK:
                                    r = TrimMK(MarkStr,
                                        stREG(rn).STCUT(cn).dblSTX, stREG(rn).STCUT(cn).dblSTY,
                                            stREG(rn).STCUT(cn).dblDL2,
                                            stREG(rn).STCUT(cn).dblV1, stREG(rn).STCUT(cn).intANG, dblQrate, 0, FORCE_MODE)

                                    If retry_marking < stREG(rn).STCUT(cn).intMarkRepeatCnt Then
                                        retry_marking = retry_marking + 1
                                        Call STRXY(rn, stREG(rn).STCUT(cn).dblSTX, stREG(rn).STCUT(cn).dblSTY)
                                        Call ObjSys.WAIT(0.1)
                                        GoTo RETRY_MARK
                                    End If

                                Else
                                        ' OKマーキングまたはNGマーキング
                                        'V2.0.0.0⑮                                r = TrimMK(stREG(rn).STCUT(cn).cFormat, _
                                        'V2.0.0.0⑮                                        stREG(rn).STCUT(cn).dblSTX, stREG(rn).STCUT(cn).dblSTY, _
                                        'V2.0.0.0⑮                                            stREG(rn).STCUT(cn).dblDL2, _
                                        'V2.0.0.0⑮                                            stREG(rn).STCUT(cn).dblV1, stREG(rn).STCUT(cn).intANG, stREG(rn).STCUT(cn).intQF1, 0, 2)        '###1042①
                                        'V2.0.0.0⑮↓
                                        r = TrimMK(stREG(rn).STCUT(cn).cFormat,
                                        stREG(rn).STCUT(cn).dblSTX, stREG(rn).STCUT(cn).dblSTY,
                                            stREG(rn).STCUT(cn).dblDL2,
                                            stREG(rn).STCUT(cn).dblV1, stREG(rn).STCUT(cn).intANG, dblQrate, 0, FORCE_MODE)
                                    'V2.0.0.0⑮↑
                                    'V1.0.4.3⑤↑
                                End If
                                'V2.2.1.7③↑

                            Case Else
                                GoTo STP_TYPE_ERR                           ' STカット/Lカット/ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ以外はエラー
                        End Select

                        ' RESET SW押下又はエラーなら戻り
                        If (rtn <> cFRS_NORMAL) And (rtn <> 2) Then         ' RESET SW押下又はエラー ?
                            If (rtn <> cFRS_ERR_RST) Then                   ' RESET SW押下以外ならトリミングNG
                                Return (cFRS_TRIM_NG)                       ' 戻値 = トリミングNG
                            Else
                                Return (cFRS_ERR_RST)
                            End If
                        End If

                        If (Form1.System1.Sys_Err_Chk_EX(gSysPrm, APP_MODE_LOTCHG) <> cFRS_NORMAL) Then ' 非常停止等 ?
                            Call Form1.AppEndDataSave()                                       ' ｿﾌﾄ強制終了時のﾃﾞｰﾀ保存確認
                            Call Form1.AplicationForcedEnding()                               ' ｿﾌﾄ強制終了処理
                            End
                        End If
                        System.Windows.Forms.Application.DoEvents()


                    Next cn                                                 ' 次のカットへ
                End If                                              'V2.0.0.0⑮
STP_NEXT:
            Next rn                                                     ' 次の抵抗へ

            '-------------------------------------------------------------------
            '   終了処理
            '-------------------------------------------------------------------
            Call ObjSys.Ilum_Ctrl(gSysPrm, Z0, ZOPT)                    ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ消灯(制御有時有効)
            Return (cFRS_NORMAL)

STP_TYPE_ERR:
            strMSG = "カット形状エラー"
            Call Z_PRINT(strMSG & vbCrLf)                               ' ログ画面に表示
            Return (cFRS_TRIM_NG)                                       ' 戻値 = トリミングNG

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.CUT_CHK() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cFRS_TRIM_NG)
        End Try
    End Function
#End Region

#Region "カッティング(STカット)"
    '''=========================================================================
    '''<summary>カッティング(STカット)</summary>
    '''<param name="rn"> (INP) 抵抗番号(1 ORG)</param>
    '''<param name="cn"> (INP) ｶｯﾄ番号 (1 ORG)</param>
    '''<returns>cFRS_NORMAL   = 正常
    '''         cFRS_TRIM_NG  = トリミングNG
    '''</returns>
    '''=========================================================================
    Public Function CUT_ST(ByRef rn As Short, ByRef cn As Short) As Short

        Dim rtn As Short = cFRS_NORMAL                                  ' 戻値
        Dim r As Short                                                  ' 戻値
        Dim strMSG As String                                            ' メッセージ編集域

        Try
            ' 斜めカット(カット量,速度,角度)
            If (stREG(rn).STCUT(cn).dblDL2 = 0) Then Exit Function '    ' ｶｯﾄ長 = 0 なら NOP
            r = CUT2(stREG(rn).STCUT(cn).dblDL2, stREG(rn).STCUT(cn).dblV1, stREG(rn).STCUT(cn).intANG)
            If (r <> cFRS_NORMAL) Then
                Call Check_ERR_LSR_STATUS_STANBY(r)                     ' レーザアラーム８３３エラー時のプログラム終了処理
                rtn = cFRS_TRIM_NG                                      ' Return値 =トリミングNG
            End If
            Return (rtn)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.CUT_ST() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cFRS_TRIM_NG)
        End Try
    End Function
#End Region

#Region "カッティング(Lカット) (x5ﾓｰﾄﾞ)"
    '''=========================================================================
    '''<summary>カッティング(Lカット) (x5ﾓｰﾄﾞ)</summary>
    '''<param name="rn"> (INP) 抵抗番号(1 ORG)</param>
    '''<param name="cn"> (INP) ｶｯﾄ番号 (1 ORG)</param>
    '''<returns>cFRS_NORMAL   = 正常
    '''         cFRS_TRIM_NG  = トリミングNG
    '''</returns>
    '''=========================================================================
    Public Function CUT_L(ByRef rn As Short, ByRef cn As Short) As Short

        On Error GoTo STP_ERR
        Dim r As Short                                  ' 戻値
        Dim iDir As Short                               ' Lﾀｰﾝ後移動方向(1:時計方向,2:反時計方向)
        Dim strMSG As String                            ' メッセージ編集域
        Dim CutLen(2) As Double
        Dim SpdOwd(2) As Double
        Dim SpdRet(2) As Double
        Dim QRateOwd(2) As Double
        Dim QRateRet(2) As Double
        Dim CondOwd(2) As Short
        Dim CondRet(2) As Short
        Dim dblQrate As Double

        CutLen(0) = stREG(rn).STCUT(cn).dblDL2
        CutLen(1) = stREG(rn).STCUT(cn).dblDL3

        SpdOwd(0) = stREG(rn).STCUT(cn).dblV1
        SpdOwd(1) = stREG(rn).STCUT(cn).dblV1
        SpdRet(0) = stREG(rn).STCUT(cn).dblV1
        SpdRet(1) = stREG(rn).STCUT(cn).dblV1

        dblQrate = stREG(rn).STCUT(cn).intQF1
        dblQrate = dblQrate / 10.0

        QRateOwd(0) = dblQrate
        QRateOwd(1) = dblQrate
        QRateRet(0) = dblQrate
        QRateRet(1) = dblQrate

        CondOwd(0) = 0
        CondOwd(1) = 0

        CondRet(0) = 0
        CondRet(1) = 0


        CUT_L = cFRS_NORMAL                             ' Return値 = 正常
        If (stREG(rn).STCUT(cn).dblDL2 = 0) And (stREG(rn).STCUT(cn).dblDL3 = 0) Then
            Exit Function                               ' 第１,2のｶｯﾄ長 = 0 なら NOP
        End If

        ' Lﾀｰﾝ後移動方向(1:時計方向,2:反時計方向)を求める
        iDir = Get_Cut_Dir(stREG(rn).STCUT(cn).intANG, stREG(rn).STCUT(cn).intANG2)

        ' 斜めLカット(Lターン後方向(時計半時計),始めの移動量,Lターン後移動量,速度,始めの移動方向角度)
        'r = LCUT2(iDir, stREG(rn).STCUT(cn).dblDL2, stREG(rn).STCUT(cn).dblDL3, stREG(rn).STCUT(cn).dblV1, stREG(rn).STCUT(cn).intANG)

        'V1.0.4.3⑦        r = TrimL(FORCE_MODE, 0, 0, stREG(rn).intSLP, stREG(rn).STCUT(cn).intTMM, stREG(rn).STCUT(cn).intANG, stREG(rn).STCUT(cn).dblLTP, stREG(rn).STCUT(cn).intANG2, _
        'V1.0.4.3⑦                    CutLen, SpdOwd, SpdRet, QRateOwd, QRateRet, CondOwd, CondRet)

        r = TRM_L6(FORCE_MODE, rn, cn, stREG(rn).dblNOM)        'V1.0.4.3⑦ 斜めLｶｯﾄ電圧/抵抗ﾄﾘﾐﾝｸﾞ


        If (r <> cFRS_NORMAL) And (r <> 2) Then
            CUT_L = cFRS_TRIM_NG                        ' Return値 =トリミングNG
        End If
        Exit Function

STP_ERR:
        strMSG = "トラップエラー(CUT_L) !! "
        Call Z_PRINT(strMSG & vbCrLf)
        CUT_L = cERR_TRAP                               ' 戻値 = ﾄﾗｯﾌﾟｴﾗｰ発生

    End Function
#End Region

#Region "カッティング(ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ) (x5ﾓｰﾄﾞ)"
    '''=========================================================================
    '''<summary>カッティング(ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ) (x5ﾓｰﾄﾞ)</summary>
    '''<param name="rn"> (INP) 抵抗番号(1 ORG)</param>
    '''<param name="cn"> (INP) ｶｯﾄ番号 (1 ORG)</param>
    '''<returns>cFRS_NORMAL   = 正常
    '''         cFRS_TRIM_NG  = トリミングNG
    '''</returns>
    '''=========================================================================
    Public Function CUT_SPT(ByRef rn As Short, ByRef cn As Short) As Short

        On Error GoTo STP_ERR
        Dim i As Short                                  ' Index
        Dim r As Short                                  ' 戻値
        Dim ANG As Short                                ' ｶｯﾄ方向(90°単位　0°～360°)
        Dim strMSG As String                            ' メッセージ編集域
        Dim stSPC As Sp_Cut_Info                        ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ情報
        Dim dblQrate As Double

        ' 初期処理
        CUT_SPT = cFRS_NORMAL                           ' Return値 = 正常
        If (stREG(rn).STCUT(cn).dblDL2 = 0) Then        ' ｶｯﾄ長0ならNOP
            Exit Function
        End If
        stSPC.dblSTX = New Double(MAXSCTN) {}           ' Sp_Cut_Info構造体初期化 
        stSPC.dblSTY = New Double(MAXSCTN) {}
        Call Set_SpCut_Info(rn, cn, stSPC)              ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ情報設定
        ANG = stREG(rn).STCUT(cn).intANG                ' カット方向

        dblQrate = stREG(rn).STCUT(cn).intQF1
        dblQrate = dblQrate / 10.0


        ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ実行(STｶｯﾄをｶｯﾄ本数分実行)
        For i = 1 To stREG(rn).STCUT(cn).intNum         ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ本数分繰返す

            ' セーフティチェック
            r = SafetyCheck()                           ' セーフティチェック
            If (r <> cFRS_NORMAL) Then                  ' エラー ?
                CUT_SPT = r                             ' Return値を設定する
                Exit Function
            End If

            ' BP絶対値移動(ｶｯﾄ位置 + PTN補正値)(絶対値)
            r = ObjSys.EX_MOVE(gSysPrm, stSPC.dblSTX(i) + stPTN(rn).dblDRX, stSPC.dblSTY(i) + stPTN(rn).dblDRY, 1)
            If (r <> cFRS_NORMAL) Then                  ' ｴﾗｰ ?
                CUT_SPT = cFRS_ERR_TRIM                 ' Return値 =トリマエラー
                Exit Function
            End If

            MoveStop()              'V2.2.0.0⑥ 

            ' 斜めSTｶｯﾄ(カット量,速度,角度)
            r = TrimSt(2, 0, 0, stREG(rn).intSLP, ANG, stREG(rn).STCUT(cn).dblDL2, stREG(rn).STCUT(cn).dblV1, stREG(rn).STCUT(cn).dblV1, dblQrate, dblQrate, 0, 0)

            'r = CUT2(stREG(rn).STCUT(cn).dblDL2, stREG(rn).STCUT(cn).dblV1, ANG)
            If (r <> cFRS_NORMAL) And (r <> 2) Then
                CUT_SPT = cFRS_TRIM_NG                  ' Return値 =トリミングNG
            End If

            Call Cnv_ANG(ANG)                           ' カット方向を反対方向に変換
        Next i
        Exit Function

STP_ERR:
        strMSG = "トラップエラー(CUT_SPT) !! "
        Call Z_PRINT(strMSG & vbCrLf)
        CUT_SPT = cERR_TRAP                             ' 戻値 = ﾄﾗｯﾌﾟｴﾗｰ発生

    End Function
#End Region

#Region "カッティング(フックカット) (x5ﾓｰﾄﾞ)"
    '''=========================================================================
    '''<summary>カッティング(フックカット) (x5ﾓｰﾄﾞ)</summary>
    '''<param name="rn"> (INP) 抵抗番号(1 ORG)</param>
    '''<param name="cn"> (INP) ｶｯﾄ番号 (1 ORG)</param>
    '''<returns>cFRS_NORMAL   = 正常
    '''         cFRS_TRIM_NG  = トリミングNG
    '''</returns>
    '''=========================================================================
    Public Function CUT_HOOK(ByRef rn As Short, ByRef cn As Short) As Short

        On Error GoTo STP_ERR
        Dim r As Short                                  ' 戻値
        Dim strMSG As String                            ' メッセージ編集域

        CUT_HOOK = cFRS_NORMAL                          ' Return値 = 正常
        'If (stREG(rn).STCUT(cn).dblDL2 = 0) Then Exit Function ' 第１のｶｯﾄ長 = 0 ?

        '' 斜めフックカット(Lターン後方向,移動量,Lターン後移動量,フックカット移動量,速度,始めの角度)
        'CUT_HOOK = HCUT2(stREG(rn).STCUT(cn).intDIR, gdblDL1(cn, 1), stREG(rn).STCUT(cn).dblDL2, stREG(rn).STCUT(cn).dblDL3, stREG(rn).STCUT(cn).dblV1, stREG(rn).STCUT(cn).intANG)

        Exit Function

STP_ERR:
        strMSG = "トラップエラー(CUT_HOOK) !! "
        Call Z_PRINT(strMSG & vbCrLf)
        CUT_HOOK = cERR_TRAP                            ' 戻値 = ﾄﾗｯﾌﾟｴﾗｰ発生

    End Function
#End Region

#Region "カッティング(IXカット) (x5ﾓｰﾄﾞ)"
    '''=========================================================================
    '''<summary>カッティング(IXカット) (x5ﾓｰﾄﾞ)</summary>
    '''<param name="rn"> (INP) 抵抗番号(1 ORG)</param>
    '''<param name="cn"> (INP) ｶｯﾄ番号 (1 ORG)</param>
    '''<returns>cFRS_NORMAL   = 正常
    '''         cFRS_TRIM_NG  = トリミングNG
    '''</returns>
    '''=========================================================================
    Public Function CUT_IX2(ByRef rn As Short, ByRef cn As Short) As Integer

        Dim i As Short                                  ' ﾙｰﾌﾟ回数
        Dim IDX As Integer                                ' ｲﾝﾃﾞｯｸｽｶｯﾄ1～3(ﾋﾟｯﾁ大,中,小)
        Dim count As Short                              ' ｲﾝﾃﾞｯｸｽｶｯﾄ数
        Dim r As Integer                                  ' 関数戻値
        Dim CutL As Double                              ' 最大ｶｯﾄ長
        Dim ln As Double                                ' 現在のｶｯﾄ長
        Dim NOM(2) As Double                            ' 目標値(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
        Dim NOMx As Double                              ' 目標値
        Dim VX(3) As Double                             ' 作業域
        Dim strMSG As String                            ' メッセージ編集域
        Dim wkL1 As Double                              ' 作業域
        Dim wkL2 As Double                              ' 作業域
        'Dim dblMx As Double                             ' 作業域
        Dim dblQrate As Double

        Try

            '---------------------------------------------------------------------------
            '   初期処理
            '---------------------------------------------------------------------------
            CUT_IX2 = cFRS_NORMAL                            ' Return値 = 正常
            strMSG = ""
            LTFlg = 1                                       ' Lﾀｰﾝﾌﾗｸﾞ(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
            LTAng(1) = stREG(rn).STCUT(cn).intANG           ' ANG(1) = Lﾀｰﾝ前のｶｯﾄ方向
            LTAng(2) = stREG(rn).STCUT(cn).intANG2          ' ANG(2) = Lﾀｰﾝ後のｶｯﾄ方向
            dblML(1) = stREG(rn).STCUT(cn).dblDL2           ' ﾘﾐｯﾄｶｯﾄ量mm(1:Lﾀｰﾝ前)
            dblML(2) = stREG(rn).STCUT(cn).dblDL3           ' ﾘﾐｯﾄｶｯﾄ量mm(2:Lﾀｰﾝ後)
            LTP = stREG(rn).STCUT(cn).dblLTP                ' Lﾀｰﾝﾎﾟｲﾝﾄ(%)
            'NOM(1) = stREG(rn).dblNOM                       ' 目標値
            ' ｶｯﾄｵﾌ(%)→目標値に対するｵﾌｾｯﾄ値(目標値×(1＋ｶｯﾄｵﾌ/100))
            'NOMx = Func_CalNomForCutOff(rn, cn, NOM(1))          'カットオフによる目標値

            'NOM(1) = NOMx                                   ' Lﾀｰﾝ前目標値 = 目標値
            'NOM(2) = NOMx                                   ' Lﾀｰﾝ後目標値 = 目標値

            '' Lﾀｰﾝ ﾎﾟｲﾝﾄ情報を設定する(Lｶｯﾄ時)
            'If (stREG(rn).STCUT(cn).intCTYP = CNS_CUTP_L) And ((LTP <> 0.0#) And (LTP < 100.0#)) Then ' LｶｯﾄでLﾀｰﾝ ﾎﾟｲﾝﾄ指定あり ?
            '    NOM(1) = Mx + (NOM(2) - Mx) * (LTP * 0.01)  ' Lﾀｰﾝ前目標値設定(初期値＋(目標値-初期値)×Lﾀｰﾝﾎﾟｲﾝﾄ/100)
            'End If

            '' Lﾀｰﾝ前ﾘﾐｯﾄｶｯﾄ量mmと目標値を設定する
            CutL = dblML(LTFlg)                             ' ﾘﾐｯﾄｶｯﾄ量mm(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
            'NOMx = NOM(LTFlg)                               ' 目標値(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)

            ' Qレート設定
            If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then
                dblQrate = stREG(rn).STCUT(cn).intQF1
                dblQrate = dblQrate / 10.0
            Else
#If cOSCILLATORcFLcUSE Then
                ' FL時は加工条件番号テーブルからQレートを設定する(カットスピードはデータから設定)
                IDX = stREG(rn).STCUT(cn).intCND(CUT_CND_L1)
                dblQrate = stCND.Freq(IDX)
#End If
            End If

            ' ｶｯﾄ量初期化
            For i = 1 To MAXCTN                             ' MAXカット数分繰返す
                dblLN(1, i) = 0.0#                          ' ｶｯﾄ量初期化(1:Lﾀｰﾝ前)
                dblLN(2, i) = 0.0#                          ' ｶｯﾄ量初期化(2:Lﾀｰﾝ後)
            Next i

            '---------------------------------------------------------------------------
            '   ｲﾝﾃﾞｯｸｽｶｯﾄでLｶｯﾄ/STｶｯﾄを行う
            '---------------------------------------------------------------------------
            For IDX = 1 To MAXIDX                           ' ｲﾝﾃﾞｯｸｽｶｯﾄ1～5(ﾋﾟｯﾁ大,中,小)分繰返す
STP_CHG_PIT:
                count = stREG(rn).STCUT(cn).intIXN(IDX)     ' count = ｲﾝﾃﾞｯｸｽｶｯﾄ数
                ln = stREG(rn).STCUT(cn).dblDL1(IDX)        ' カット長を設定する(ﾋﾟｯﾁ大,中,小)
                For i = 1 To count                          ' ｲﾝﾃﾞｯｸｽｶｯﾄ数分繰返す


#If (cCND = 1) Then ' 条件出しﾓｰﾄﾞ ?
                    strMSG = "　 抵抗番号=" + Format(rn, "0") + ",ｶｯﾄ番号=" + Format(cn, "0")
                    strMSG = strMSG + ",目標値(L1,L2)=" + Format(NOM(1), "0.0####") + "," + Format(NOM(2), "0.0####") + vbCrLf
                    strMSG = strMSG + "　   ｶｯﾄ長=" + Format(ln, "#0.0####") + ",目標値=" + Format(NOMx, "0.0####") + ",LTFlg=" + Format(LTFlg, "0") + ",ｶｯﾄ量(Lﾀｰﾝ前,後)=" + Format(dblLN(1, cn), "#0.0####") + "," + Format(dblLN(2, cn), "#0.0####")
                    Call Z_PRINT(strMSG + vbCrLf)
#End If
                    'Call DScanModeSet(rn, cn, IDX)              ' ＤＣスキャナに接続する測定器を切り替えてスキャナをセットする。

                    ' 電圧(外部/内部)/抵抗測定(内部)を行う
                    'r = V_R_MEAS(stREG(rn).intSLP, stREG(rn).STCUT(cn).intMType, dblMx, rn, stREG(rn).dblNOM)
                    'r = V_R_MEAS(stREG(rn).intSLP, stREG(rn).STCUT(cn).intIXMType(IDX), dblMx, rn, stREG(rn).dblNOM)

#If (cCND = 1) Then ' 条件出しﾓｰﾄﾞ ?
                    strMSG = "■測定値=" + ObjUtl.sFormat(dblMx, "#0.0###", 7)
                    Call Z_PRINT(strMSG + vbCrLf)
#End If
                    ' 目標値を超えたか調べる
                    'If (stREG(rn).intSLP = SLP_VTRIMPLS) Or (stREG(rn).intSLP = SLP_RTRM) Or (stREG(rn).intSLP = SLP_ATRIMPLS) Then ' +ｽﾛｰﾌﾟ/抵抗 ?
                    '    If (dblMx >= NOMx) Then             ' 測定値 >= 目標値なら次へ
                    '        CUT_IX2 = 1                      ' Return値 = 1(目標値を超えたので終了)
                    '        '-------------------------
                    '        '   Lﾀｰﾝ処理(Lｶｯﾄ時)
                    '        '-------------------------
                    '        ' Lｶｯﾄ以外ならEXIT
                    '        If (stREG(rn).STCUT(cn).intCTYP <> CNS_CUTP_L) Then Exit Function
                    '        If (LTFlg >= 2) Then Exit Function ' Lﾀｰﾝ後ならEXIT
                    '        CUT_IX2 = 0                                      ' Return値 = 正常
                    '        LTFlg = 2                                       ' Lﾀｰﾝﾌﾗｸﾞ = 2(Lﾀｰﾝ後)
                    '        CutL = dblML(LTFlg)                             ' ﾘﾐｯﾄｶｯﾄ量mm(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
                    '        NOMx = NOM(LTFlg)                               ' 目標値(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
                    '        ln = stREG(rn).STCUT(cn).dblDL1(IDX)            ' カット長を設定する(ﾋﾟｯﾁ大,中,小)
                    '        GoTo TRM_IX_NEXT                               ' Lｶｯﾄを行う  ※※GoTo TRM_IX_NEXTがエラーとなるので修正必要
                    '        '-------------------------
                    '        '   Lﾀｰﾝ処理(Lｶｯﾄ時)
                    '        '-------------------------


                    '    End If
                    'Else                                    ' -ｽﾛｰﾌﾟ ?
                    '    If (dblMx <= NOMx) Then             ' 測定値 <= 目標値なら次へ
                    '        CUT_IX2 = 1                      ' Return値 = 1(目標値を超えたので終了)
                    '        '-------------------------
                    '        '   Lﾀｰﾝ処理(Lｶｯﾄ時)
                    '        '-------------------------
                    '        ' Lｶｯﾄ以外ならEXIT
                    '        If (stREG(rn).STCUT(cn).intCTYP <> CNS_CUTP_L) Then Exit Function

                    '        If (LTFlg >= 2) Then Exit Function ' Lﾀｰﾝ後ならEXIT
                    '        CUT_IX2 = 0                                      ' Return値 = 正常
                    '        LTFlg = 2                                       ' Lﾀｰﾝﾌﾗｸﾞ = 2(Lﾀｰﾝ後)
                    '        CutL = dblML(LTFlg)                             ' ﾘﾐｯﾄｶｯﾄ量mm(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
                    '        NOMx = NOM(LTFlg)                               ' 目標値(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
                    '        ln = stREG(rn).STCUT(cn).dblDL1(IDX)            ' カット長を設定する(ﾋﾟｯﾁ大,中,小)
                    '        GoTo TRM_IX_NEXT                               ' Lｶｯﾄを行う  ※※GoTo TRM_IX_NEXTがエラーとなるので修正必要
                    '        '-------------------------
                    '        '   Lﾀｰﾝ処理(Lｶｯﾄ時)
                    '        '-------------------------

                    '    End If
                    'End If

                    '' ｲﾝﾃﾞｯｸｽｶｯﾄﾋﾟｯﾁ(1:大,2-4:中,5:小)をチェックする
                    'r = Get_Idx_Pitch(rn, cn, IDX, NOMx, dblMx) ' 目標値との誤差によりﾋﾟｯﾁを変更する
                    'If (r <> IDX) Then                      ' ｶｯﾄﾋﾟｯﾁ変更 ?
                    '    IDX = r                             ' ｲﾝﾃﾞｯｸｽｶｯﾄﾋﾟｯﾁを変更する
                    '    GoTo STP_CHG_PIT
                    'End If

                    MoveStop()              'V2.2.0.0⑥ 

                    ' 次のｶｯﾄで最大ｶｯﾄ量を超える ? (※下記のようにしないと正しい比較ができない)
                    wkL1 = CDbl((dblLN(LTFlg, cn) + ln).ToString("#0.0000"))
                    wkL2 = CDbl(CutL.ToString("#0.0000"))
                    If (wkL1 > wkL2) Then                   ' 最大ｶｯﾄ量を超える ?
                        ' ln = 残りのｶｯﾄ量(下記のようにしないとln=0とならない場合あり)
                        ln = CDbl((wkL2 - dblLN(LTFlg, cn)).ToString("#0.0000"))
                        If (ln <= 0) Then                   ' 最大ｶｯﾄ量までカット ?
                            CUT_IX2 = 2                      ' Return値 = 2(指定移動量までカット)
                            '-------------------------
                            '   Lﾀｰﾝ処理(Lｶｯﾄ時)
                            '-------------------------
                            ' Lｶｯﾄ以外ならEXIT
                            If (stREG(rn).STCUT(cn).intCTYP <> CNS_CUTP_L) Then Exit Function
                            If (LTFlg >= 2) Then Exit Function ' Lﾀｰﾝ後ならEXIT
                            CUT_IX2 = 0                                      ' Return値 = 正常
                            LTFlg = 2                                       ' Lﾀｰﾝﾌﾗｸﾞ = 2(Lﾀｰﾝ後)
                            CutL = dblML(LTFlg)                             ' ﾘﾐｯﾄｶｯﾄ量mm(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
                            NOMx = NOM(LTFlg)                               ' 目標値(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
                            ln = stREG(rn).STCUT(cn).dblDL1(IDX)            ' カット長を設定する(ﾋﾟｯﾁ大,中,小)
                            GoTo TRM_IX_NEXT                               ' Lｶｯﾄを行う  ※※GoTo TRM_IX_NEXTがエラーとなるので修正必要
                            '-------------------------
                            '   Lﾀｰﾝ処理(Lｶｯﾄ時)
                            '-------------------------


                        End If
                    End If

                    ' 斜め直線ｶｯﾄ(ﾎﾟｼﾞｼｮﾆﾝｸﾞなし)
                    Dim shSLP As Short
                    If (stREG(rn).intSLP = SLP_ATRIMPLS) Then
                        shSLP = 1
                    ElseIf (stREG(rn).intSLP = SLP_ATRIMMNS) Then
                        shSLP = 2
                    Else
                        shSLP = stREG(rn).intSLP
                    End If
                    r = TrimSt(FORCE_MODE, 0, 0, shSLP, LTAng(LTFlg), ln, stREG(rn).STCUT(cn).dblV1, stREG(rn).STCUT(cn).dblV1, _
                                  dblQrate, dblQrate, stREG(rn).STCUT(cn).intCND(CUT_CND_L1), stREG(rn).STCUT(cn).intCND(CUT_CND_L1))
                    Call DebugLogOut("CUT=[" & r.ToString("0") & "]IDX=[" & IDX.ToString("0") & "]SLP=[" & shSLP.ToString("0") & "]LEN=[" & ln.ToString("0.0000") & "]V=[" & stREG(rn).STCUT(cn).dblV1.ToString("0.0000") & "]Qrate=[" & dblQrate.ToString("0.0000") & "]")

                    Call ZWAIT(stREG(rn).STCUT(cn).lngPAU(IDX)) ' ｲﾝﾃﾞｯｸｽﾋﾟｯﾁ間ﾎﾟｰｽﾞ(ms)
                    If (r <> 0) And (r <> 2) Then
                        Return (cFRS_TRIM_NG)                           ' Return値 = トリミングNG
                    End If
                    dblLN(LTFlg, cn) = dblLN(LTFlg, cn) + ln ' ｶｯﾄ済量mmを退避(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)

TRM_IX_NEXT:
                Next i                                      ' 次ｶｯﾄへ

                ' セーフティチェック
                r = SafetyCheck()                           ' セーフティチェック
                If (r <> 0) Then                            ' エラー ?
                    CUT_IX2 = r                              ' Return値 = セーフティチェックエラー
                    Exit Function
                End If
            Next IDX                                        ' 次ﾋﾟｯﾁへ
            Exit Function

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.CUT_IX2() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cFRS_TRIM_NG)                           ' Return値 = トリミングNG
        End Try
    End Function
#End Region

    '==========================================================================
    '   トリミング実行
    '==========================================================================
#Region "ストレートカット(斜め直線カット電圧/抵抗トリミング)"
    '''=========================================================================
    ''' <summary>レーザーカット(ストレートカット)</summary>
    ''' <param name="MoveMode">動作モード(0:トリミング、1:ティーチング、2:強制カット)</param>
    ''' <param name="CutMode">カットモード(0:ノーマル、1:リターン、2:リトレース、4:斜め)</param>
    ''' <param name="Target">目標値(カット時は0を設定)</param>
    ''' <param name="Slope">スロープ設定(1:電圧測定＋スロープ、2:電圧測定－スロープ、4:抵抗測定＋スロープ、5:抵抗測定－スロープ)</param>
    ''' <param name="Angle">カット角度</param>
    ''' <param name="CutLen">カット長：Line1～4用のパラメータ(1～4:Line1～4)</param>
    ''' <param name="SpdOwd">カットスピード(往路)：Line1～4用のパラメータ(1～4:Line1～4)</param>
    ''' <param name="SpdRet">カットスピード(復路)：Line1～4用のパラメータ(1～4:Line1～4)</param>
    ''' <param name="QRateOwd">カットQレート(往路)：Line1～4用のパラメータ(1～4:Line1～4)</param>
    ''' <param name="QRateRet">カットQレート(復路)：Line1～4用のパラメータ(1～4:Line1～4)</param>
    ''' <param name="CondOwd">加工条件番号(往路)：Line1～4用のパラメータ(1～4:Line1～4)</param>
    ''' <param name="CondRet">加工条件番号(復路)：Line1～4用のパラメータ(1～4:Line1～4)</param>
    ''' <param name="OffsetX">リトレース時オフセットＸ</param>
    ''' <param name="OffsetY">リトレース時オフセットＹ</param>
    ''' <returns>cFRS_NORMAL(0)  = 正常, 1 = 目標値を超えたので終了, 2 = 指定移動量までカットしたので終了
    ''' </returns>
    '''=========================================================================
    Public Function TrimSt(ByVal MoveMode As Short, ByVal CutMode As Short, ByVal Target As Double, ByVal Slope As Short, ByVal Angle As Double, ByVal CutLen As Double, ByVal SpdOwd As Double, _
    ByVal SpdRet As Double, ByVal QRateOwd As Double, ByVal QRateRet As Double, ByVal CondOwd As Short, ByVal CondRet As Short, _
                Optional ByVal OffsetX As Double = 0.0, Optional ByVal OffsetY As Double = 0.0) As Integer

        Dim rslt As Integer
        Dim strMSG As String                                            ' メッセージ編集域

        Try
            If Slope <> SLP_VTRIMPLS And Slope <> SLP_VTRIMMNS And Slope <> SLP_RTRM Then
                TrimSt = cFRS_NORMAL
                Exit Function
            End If

            If CutLen <= 0.0 Then               'V1.2.0.1
                TrimSt = cFRS_NORMAL            'V1.2.0.1
                Exit Function                   'V1.2.0.1
            End If                              'V1.2.0.1

            If CutMode = CUT_MODE_RETRACE Then  'V2.0.0.0⑦
                CutMode = CUT_MODE_NORMAL       'V2.0.0.0⑦
            End If                              'V2.0.0.0⑦

            cutCmnPrm.CutInfo.srtMoveMode = MoveMode                    ' 動作モード(0:トリミング、1:ティーチング、2:強制カット)
            cutCmnPrm.CutInfo.srtCutMode = CutMode                      ' カットモード(0:ノーマル、1:リターン、2:リトレース、4:斜め)
            cutCmnPrm.CutInfo.dblTarget = Target                        ' 目標値(カット時は0を設定)
            cutCmnPrm.CutInfo.srtSlope = Slope                          ' スロープ設定(1:電圧測定＋スロープ、2:電圧測定－スロープ、4:抵抗測定＋スロープ、5:抵抗測定－スロープ)
            cutCmnPrm.CutInfo.dblAngle = Angle                          ' カット角度

            cutCmnPrm.CutCond.CutLen.dblL1 = CutLen                     ' カット長：Line1用のパラメータ
            cutCmnPrm.CutCond.SpdOwd.dblL1 = SpdOwd                     ' カットスピード(往路)：Line1用のパラメータ
            cutCmnPrm.CutCond.SpdRet.dblL1 = SpdRet                     ' カットスピード(復路)：Line1用のパラメータ
            cutCmnPrm.CutCond.QRateOwd.dblL1 = QRateOwd                 ' カットQレート(往路)：Line1用のパラメータ
            cutCmnPrm.CutCond.QRateRet.dblL1 = QRateRet                 ' カットQレート(復路)：Line1用のパラメータ
            cutCmnPrm.CutCond.CondOwd.srtL1 = CondOwd                   ' 加工条件番号(往路)：Line1用のパラメータ
            cutCmnPrm.CutCond.CondRet.srtL1 = CondRet                   ' 加工条件番号(復路)：Line1用のパラメータ

            If CutMode = CUT_MODE_RETRACE Then
                cutCmnPrm.CutCond.CutLen.dblL2 = OffsetX                ' 'V1.0.4.3⑧リトレース時オフセットＸ
                cutCmnPrm.CutCond.CutLen.dblL3 = OffsetY                ' 'V1.0.4.3⑧リトレース時オフセットＹ
            Else
                cutCmnPrm.CutCond.CutLen.dblL2 = 0.0                    ' 'V1.0.4.3⑧リトレース時オフセットＸ
                cutCmnPrm.CutCond.CutLen.dblL3 = 0.0                    ' 'V1.0.4.3⑧リトレース時オフセットＹ
            End If

            ' ストレートカットを実行する
            rslt = TRIM_ST(cutCmnPrm)                                   ' 戻り値(1 = 目標値を超えたので終了, 2 = 指定移動量までカットしたので終了)
            Call Check_ERR_LSR_STATUS_STANBY(rslt)                      ' レーザアラーム８３３エラー時のプログラム終了処理
            Return (rslt)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.TrimSt() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Return値 = ﾄﾗｯﾌﾟｴﾗｰ発生
        End Try
    End Function
#End Region

#Region "Ｌカット(斜めＬカット電圧/抵抗トリミング)"
    '''=========================================================================
    ''' <summary>Ｌカット(斜めＬカット電圧/抵抗トリミング)</summary>
    ''' <param name="MoveMode">動作モード(0:トリミング、1:ティーチング、2:強制カット)</param>
    ''' <param name="CutMode">カットモード(0:ノーマル、1:リターン、2:リトレース、4:斜め)</param>
    ''' <param name="Target">目標値(カット時は0を設定)</param>
    ''' <param name="Slope">スロープ設定(1:電圧測定＋スロープ、2:電圧測定－スロープ、4:抵抗測定＋スロープ、5:抵抗測定－スロープ)</param>
    ''' <param name="Angle">カット角度</param>
    ''' <param name="LTurnPoint">Lターンポイント</param>
    ''' <param name="LTurnDir">Lターン後の方向</param>
    ''' <param name="CutLen">カット長：Line1～4用のパラメータ(1～4:Line1～4)</param>
    ''' <param name="SpdOwd">カットスピード(往路)：Line1～4用のパラメータ(1～4:Line1～4)</param>
    ''' <param name="SpdRet">カットスピード(復路)：Line1～4用のパラメータ(1～4:Line1～4)</param>
    ''' <param name="QRateOwd">カットQレート(往路)：Line1～4用のパラメータ(1～4:Line1～4)</param>
    ''' <param name="QRateRet">カットQレート(復路)：Line1～4用のパラメータ(1～4:Line1～4)</param>
    ''' <param name="CondOwd">カット条件番号(往路)：Line1～4用のパラメータ(1～4:Line1～4)</param>
    ''' <param name="CondRet">カット条件番号(復路)：Line1～4用のパラメータ(1～4:Line1～4)</param>
    '''=========================================================================
    Public Function TrimL(ByVal MoveMode As Short, ByVal CutMode As Short, ByVal Target As Double, ByVal Slope As Short, ByVal MeasType As Short, ByVal Angle As Double, ByVal LTurnPoint As Double, ByVal LTurnDir As Short,
    ByVal CutLen As Double(), ByVal SpdOwd As Double(), ByVal SpdRet As Double(), ByVal QRateOwd As Double(), ByVal QRateRet As Double(), ByVal CondOwd As Short(), ByVal CondRet As Short(), Optional ByVal Radius As Double = 0.0) As Integer

        Dim i As Integer
        Dim rslt As Integer
        Dim strMSG As String                                            ' メッセージ編集域

        Try
            cutCmnPrm.CutInfo.srtMoveMode = MoveMode                    ' 動作モード(0:トリミング、1:ティーチング、2:強制カット)
            cutCmnPrm.CutInfo.srtCutMode = CutMode                      ' カットモード(0:ノーマル、1:リターン、2:リトレース、4:斜め)
            cutCmnPrm.CutInfo.dblTarget = Target                        ' 目標値(カット時は0を設定)
            cutCmnPrm.CutInfo.srtSlope = Slope                          ' スロープ設定(1:電圧測定＋スロープ、2:電圧測定－スロープ、4:抵抗測定＋スロープ、5:抵抗測定－スロープ)
            cutCmnPrm.CutInfo.srtMeasType = MeasType                    ' 測定タイプ(0:高速(3回)、1:高精度(2000回)、2:（IDXのみ）外部機器、3:測定無し、5～:指定回数測定）
            cutCmnPrm.CutInfo.dblAngle = Angle                          ' カット角度
            cutCmnPrm.CutInfo.dblLTP = LTurnPoint                       ' Lターンポイント
            cutCmnPrm.CutInfo.srtLTDIR = LTurnDir                       ' Lターン後の方向
            cutCmnPrm.CutInfo.dblRADI = Radius                          ' Ｒ半径       'V2.2.0.0②

            i = 0
            cutCmnPrm.CutCond.CutLen.dblL1 = CutLen(i)                  ' カット長：Line1用のパラメータ
            cutCmnPrm.CutCond.SpdOwd.dblL1 = SpdOwd(i)                  ' カットスピード(往路)：Line1用のパラメータ
            cutCmnPrm.CutCond.SpdRet.dblL1 = SpdRet(i)                  ' カットスピード(復路)：Line1用のパラメータ
            cutCmnPrm.CutCond.QRateOwd.dblL1 = QRateOwd(i)              ' カットQレート(往路)：Line1用のパラメータ
            cutCmnPrm.CutCond.QRateRet.dblL1 = QRateRet(i)              ' カットQレート(復路)：Line1用のパラメータ
            cutCmnPrm.CutCond.CondOwd.srtL1 = CondOwd(i)                ' カット条件番号(往路)：Line1用のパラメータ
            cutCmnPrm.CutCond.CondRet.srtL1 = CondRet(i)                ' カット条件番号(復路)：Line1用のパラメータ

            i = 1
            cutCmnPrm.CutCond.CutLen.dblL2 = CutLen(i)                  ' カット長：Line2用のパラメータ
            cutCmnPrm.CutCond.SpdOwd.dblL2 = SpdOwd(i)                  ' カットスピード(往路)：Line2用のパラメータ
            cutCmnPrm.CutCond.SpdRet.dblL2 = SpdRet(i)                  ' カットスピード(復路)：Line2用のパラメータ
            cutCmnPrm.CutCond.QRateOwd.dblL2 = QRateOwd(i)              ' カットQレート(往路)：Line2用のパラメータ
            cutCmnPrm.CutCond.QRateRet.dblL2 = QRateRet(i)              ' カットQレート(復路)：Line2用のパラメータ
            cutCmnPrm.CutCond.CondOwd.srtL2 = CondOwd(i)                ' カット条件番号(往路)：Line2用のパラメータ
            cutCmnPrm.CutCond.CondRet.srtL2 = CondRet(i)                ' カット条件番号(復路)：Line2用のパラメータ

            'V2.2.0.0②rslt = TRIM_L(cutCmnPrm)
            rslt = TRIM_LWithR(cutCmnPrm)                               ' R付きLカット   V2.2.0.0②
            Call Check_ERR_LSR_STATUS_STANBY(rslt)                      ' レーザアラーム８３３エラー時のプログラム終了処理

            Return (rslt)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.TrimL() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Return値 = ﾄﾗｯﾌﾟｴﾗｰ発生
        End Try
    End Function
#End Region

#Region "インデックスカット(斜めインデックスカット電圧/抵抗トリミング)"
    '''=========================================================================
    ''' <summary>インデックスカット(斜めインデックスカット電圧/抵抗トリミング)</summary>
    ''' <param name="MoveMode">動作モード(0:トリミング、1:ティーチング、2:強制カット)</param>
    ''' <param name="CutMode">カットモード(0:ノーマル、1:リターン、2:リトレース、4:斜め)</param>
    ''' <param name="Target">目標値(カット時は0を設定)</param>
    ''' <param name="Slope">スロープ設定(1:電圧測定＋スロープ、2:電圧測定－スロープ、4:抵抗測定＋スロープ、5:抵抗測定－スロープ)</param>
    ''' <param name="MeasType"></param>
    ''' <param name="Angle">カット角度</param>
    ''' <param name="CutLen">カット長：Line1～4用のパラメータ(1～4:Line1～4)</param>
    ''' <param name="IndexScanCount"></param>
    ''' <param name="IndexMeasMode"></param>
    ''' <param name="SpdOwd">カットスピード(往路)：Line1～4用のパラメータ(1～4:Line1～4)</param>
    ''' <param name="SpdRet">カットスピード(復路)：Line1～4用のパラメータ(1～4:Line1～4)</param>
    ''' <param name="QRateOwd">カットQレート(往路)：Line1～4用のパラメータ(1～4:Line1～4)</param>
    ''' <param name="QRateRet">カットQレート(復路)：Line1～4用のパラメータ(1～4:Line1～4)</param>
    ''' <param name="CondOwd">カット条件番号(往路)：Line1～4用のパラメータ(1～4:Line1～4)</param>
    ''' <param name="CondRet">カット条件番号(復路)：Line1～4用のパラメータ(1～4:Line1～4)</param>
    '''=========================================================================
    Public Function TrimIx(ByVal MoveMode As Short, ByVal CutMode As Short, ByVal Target As Double, ByVal Slope As Short, ByVal MeasType As Short, _
    ByVal Angle As Double, ByVal CutLen As Double, ByVal IndexScanCount As Short, ByVal IndexMeasMode As Short, _
    ByVal SpdOwd As Double, ByVal SpdRet As Double, ByVal QRateOwd As Double, ByVal QRateRet As Double, ByVal CondOwd As Short, ByVal CondRet As Short) As Integer

        Dim rslt As Integer
        Dim strMSG As String                                            ' メッセージ編集域

        Try
            cutCmnPrm.CutInfo.srtMoveMode = MoveMode                    ' 動作モード(0:トリミング、1:ティーチング、2:強制カット)
            cutCmnPrm.CutInfo.srtCutMode = CutMode                      ' カットモード(0:ノーマル、1:リターン、2:リトレース、4:斜め)
            cutCmnPrm.CutInfo.dblTarget = Target                        ' 目標値(カット時は0を設定)
            cutCmnPrm.CutInfo.srtSlope = Slope                          ' スロープ設定(1:電圧測定＋スロープ、2:電圧測定－スロープ、4:抵抗測定＋スロープ、5:抵抗測定－スロープ)
            cutCmnPrm.CutInfo.srtMeasType = MeasType                    ' 測定タイプ(0:高速(3回)、1:高精度(2000回)、2:（IDXのみ）外部機器、3:測定無し、5～:指定回数測定）
            cutCmnPrm.CutInfo.dblAngle = Angle                          ' カット角度
            cutCmnPrm.CutInfo.srtIdxScnCnt = IndexScanCount             ' インデックス/スキャンカット数(1～32767)
            cutCmnPrm.CutInfo.srtIdxMeasMode = IndexMeasMode            ' インデックス測定モード(0:抵抗、1:電圧、2:外部)

            cutCmnPrm.CutCond.CutLen.dblL1 = CutLen                     ' カット長：Line1用のパラメータ
            cutCmnPrm.CutCond.SpdOwd.dblL1 = SpdOwd                     ' カットスピード(往路)：Line1用のパラメータ
            cutCmnPrm.CutCond.SpdRet.dblL1 = SpdRet                     ' カットスピード(復路)：Line1用のパラメータ
            cutCmnPrm.CutCond.QRateOwd.dblL1 = QRateOwd                 ' カットQレート(往路)：Line1用のパラメータ
            cutCmnPrm.CutCond.QRateRet.dblL1 = QRateRet                 ' カットQレート(復路)：Line1用のパラメータ
            cutCmnPrm.CutCond.CondOwd.srtL1 = CondOwd                   ' カット条件番号(往路)：Line1用のパラメータ
            cutCmnPrm.CutCond.CondRet.srtL1 = CondRet                   ' カット条件番号(復路)：Line1用のパラメータ

            rslt = TRIM_IX(cutCmnPrm)
            Call Check_ERR_LSR_STATUS_STANBY(rslt)                      ' レーザアラーム８３３エラー時のプログラム終了処理

            Return (rslt)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.TrimIx() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Return値 = ﾄﾗｯﾌﾟｴﾗｰ発生
        End Try
    End Function
#End Region

    'V2.0.0.0⑦↓
#Region "リトレースカット"
    ''' <summary>
    ''' ストレート・リトレースカット本数１０本化
    ''' </summary>
    ''' <param name="rn">抵抗番号</param>
    ''' <param name="cn">カット番号</param>
    ''' <param name="Len">カット長</param>
    ''' <returns>cFRS_NORMAL(0)  = 正常, その他異常（1 = 目標値を超えたので終了）</returns>
    ''' <remarks></remarks>
    Private Function CUT_RETRACE(ByRef rn As Short, ByRef cn As Short, ByRef Len As Double) As Short
        Try
            Dim dPosX As Double, dPosY As Double, dQrate As Double, dSpeed As Double
            Dim rtn As Integer

            If Len <= 0.0 Then
                Return (cFRS_NORMAL)
            End If

            If stREG(rn).STCUT(cn).intCTYP <> CNS_CUTP_ST_TR Then
                Z_PRINT("リトレースカットで無いカットからリトレースカットが呼ばれましたRES=[" & rn.ToString & "]CUT=[" & cn.ToString & "]")
                Return (cFRS_NORMAL)
            End If

            dPosX = stREG(rn).STCUT(cn).dblSTX
            dPosY = stREG(rn).STCUT(cn).dblSTY

            For i As Short = 1 To stREG(rn).STCUT(cn).intRetraceCnt
                dPosX = dPosX + stREG(rn).STCUT(cn).dblRetraceOffX(i)
                dPosY = dPosY + stREG(rn).STCUT(cn).dblRetraceOffY(i)
                dQrate = stREG(rn).STCUT(cn).dblRetraceQrate(i) / 10.0
                dSpeed = stREG(rn).STCUT(cn).dblRetraceSpeed(i)
                Call STRXY(rn, dPosX, dPosY)                            ' カット位置移動
                rtn = TrimSt(FORCE_MODE, CUT_MODE_NORMAL, 0, SLP_RTRM, stREG(rn).STCUT(cn).intANG, Len, dSpeed, dSpeed, dQrate, dQrate, 0, 0)
                If (rtn <> 0) And (rtn <> 2) Then
                    Z_PRINT("CUT_RETRACE ERROR RETURN =[" & rtn.ToString & "] RES=[" & rn.ToString & "]CUT=[" & cn.ToString & "]")
                    Return (cFRS_TRIM_NG)                           ' Return値 = トリミングNG
                End If
            Next

            Return (cFRS_NORMAL)

        Catch ex As Exception
            MsgBox("User.CUT_RETRACE() TRAP ERROR = " + ex.Message)
            Return (cERR_TRAP)                                          ' Return値 = ﾄﾗｯﾌﾟｴﾗｰ発生
        End Try
    End Function

#End Region
    'V2.0.0.0⑦↑

    '==========================================================================
    '   トリミング結果表示/ログ出力処理
    '==========================================================================
#Region "タイトルメッセージ表示(ログ画面)/印刷"
    '''=========================================================================
    '''<summary>タイトルメッセージ表示(ログ画面)/印刷</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Sub Disp_Init()

        Dim strMSG As String

        Try
            ' ﾄﾘﾐﾝｸﾞﾓｰﾄﾞ表示((例)' === 測定モード ===")
            strMSG = "=== " & TTL_Msg(DGL) & " ==="                     ' ﾄﾘﾐﾝｸﾞﾓｰﾄﾞ
            Call Z_PRINT(strMSG & vbCrLf)                               ' ログ画面に表示
            ' ﾄﾘﾐﾝｸﾞﾓｰﾄﾞ,測定ﾓｰﾄﾞ以外はReturn
            If (Not UserSub.IsTRIM_MODE_ITTRFT()) And (DGL <> TRIM_MODE_MEAS) Then Exit Sub
            If (prf = 1) Then                                           ' 印刷?
                'Call ObjPrt.Z_LPRINT(strMSG)                           ' タイトル印刷
            End If

            ' タイトル1表示
            'If (DGL <> TRIM_MODE_MEAS) Then                             ' ﾄﾘﾐﾝｸﾞﾓｰﾄﾞ ?
            '    strMSG = "       【イニシャル】                   【ファイナル】"
            '    '       strMSG = "              【イニシャル】                  【ファイナル】"
            '    Call Z_PRINT(strMSG & vbCrLf) ' ログ画面に表示
            '    If (prf = 1) Then                                       ' 印刷?
            '        'Call ObjPrt.Z_LPRINT(strMSG)                       ' タイトル印刷
            '    End If
            'End If

            '' '' '' タイトル2表示
            '' '' '' 測定ﾓｰﾄﾞ時
            ' '' ''If (DGL = TRIM_MODE_MEAS) Or (DGL = TRIM_MODE_ITTRFT) Then
            ' '' ''    strMSG = "抵抗    目標値     測定値       誤差"
            ' '' ''    'Else ' ﾄﾘﾐﾝｸﾞﾓｰﾄﾞ
            ' '' ''    '    '       strMSG = "抵抗    目標値[ V] 測定値[ V] 誤差[ V]  目標値[ V] 測定値[ V] 誤差[ V] 判定"
            ' '' ''    '    strMSG = "抵抗    目標値     測定値      誤差    　目標値     測定値      誤差  　 判定"
            ' '' ''End If

            ' '' ''Call Z_PRINT(strMSG & vbCrLf)                               ' ログ画面に表示
            ' '' ''If (prf = 1) Then                                           ' 印刷?
            ' '' ''    'Call ObjPrt.Z_LPRINT(strMSG)                           ' タイトル印刷
            ' '' ''End If

            '' '' '' タイトル3表示
            '' '' '' 測定ﾓｰﾄﾞ時
            ' '' ''If (DGL = TRIM_MODE_MEAS) Or (DGL = TRIM_MODE_ITTRFT) Then
            ' '' ''    strMSG = "--------------------------------------"
            ' '' ''    'Else
            ' '' ''    '    strMSG = "-----------------------------------------------------------------------------"
            ' '' ''End If
            ' '' ''Call Z_PRINT(strMSG & vbCrLf)                               ' ログ画面に表示
            ' '' ''If (prf = 1) Then                                           ' 印刷?
            ' '' ''    'Call ObjPrt.Z_LPRINT(strMSG)                           ' タイトル印刷
            ' '' ''End If
            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.Disp_Init() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "測定値表示(ログ画面)/印刷/ログ出力"
    '''=========================================================================
    '''<summary>測定値表示(ログ画面)/印刷/ログ出力</summary>
    '''<param name="rn"> (INP) 抵抗番号(1 ORG)</param>
    '''=========================================================================
    Sub Disp_Final(ByRef rn As Short)

        Dim i As Short                                                  ' Index
        Dim strMSG As String = ""                                       ' メッセージ編集域
        Dim strLOG As String                                            ' ログ編集域
        Dim strTANI(2) As String                                        ' 単位("V","Ω" 等)
        Dim wkNOM(2) As Double                                          ' 目標値(1:目標照度,2:目標値(V))
        Dim wkVx(2) As Double                                           ' 測定値
        Dim WKdev(2) As Double                                          ' 誤差(1:IT, 2:FT)

        Try
            '---------------------------------------------------------------------------
            '   トリミング結果をログ画面に出力する
            '---------------------------------------------------------------------------
            If Not UserSub.IsSpecialTrimType Or DGL = TRIM_MODE_MEAS Then

                ' ■タイトル2表示
                If (rn = 1) Then
                    ' 測定ﾓｰﾄﾞ時
                    If (DGL = TRIM_MODE_MEAS) Or UserSub.IsTRIM_MODE_ITTRFT() Then
                        '                        strMSG = "抵抗    目標値     測定値       誤差"
                        strMSG = "抵抗       目標値          測定値           誤差"
                        'Else ' ﾄﾘﾐﾝｸﾞﾓｰﾄﾞ
                        '    '       strMSG = "抵抗    目標値[ V] 測定値[ V] 誤差[ V]  目標値[ V] 測定値[ V] 誤差[ V] 判定"
                        '    strMSG = "抵抗    目標値     測定値      誤差    　目標値     測定値      誤差  　 判定"
                    End If
                    If UserSub.IsTRIM_MODE_ITTRFT() Then
                        strMSG = strMSG + "  　 判定"
                    End If

                    Call Z_PRINT(strMSG & vbCrLf)                               ' ログ画面に表示
                    If (prf = 1) Then                                           ' 印刷?
                        'Call ObjPrt.Z_LPRINT(strMSG)                           ' タイトル印刷
                    End If

                    ' タイトル3表示
                    ' 測定ﾓｰﾄﾞ時
                    If (DGL = TRIM_MODE_MEAS) Then
                        '                        strMSG = "--------------------------------------"
                        strMSG = "-------------------------------------------------------"
                    Else
                        strMSG = "------------------------------------------------"
                    End If
                    Call Z_PRINT(strMSG & vbCrLf)                               ' ログ画面に表示
                    If (prf = 1) Then                                           ' 印刷?
                        'Call ObjPrt.Z_LPRINT(strMSG)                           ' タイトル印刷
                    End If
                End If
            End If

            ' 抵抗名/単位を設定する
            strMSG = stREG(rn).strRNO.PadRight(11)
            strTANI(1) = stREG(rn).strTANI.PadRight(2)
            If (stREG(rn).intMode = 0) Then                             ' 判定モード = 0(比率(%)) ?
                strTANI(2) = "ppm"
            Else
                strTANI(2) = strTANI(1)                                 ' 絶対値なら単位を設定
            End If

            ' 目標値/測定値/誤差を設定する(IT/FT)
            For i = 1 To 2                                              ' IT/FT
                wkNOM(i) = dblNM(i)                                     ' 目標値(1:IT,2:FT)
                wkVx(i) = dblVX(i)                                      ' 測定値(1:IT,2:FT)
                If (stREG(rn).intMode = 0) Then                         ' 判定モード = 0(比率(%)) ?
                    WKdev(i) = FNDEVP(wkVx(i), wkNOM(i))                ' 誤差 = (測定値 / 目標値 - 1) * 100
                Else
                    WKdev(i) = wkVx(i) - wkNOM(i)                       ' 誤差1(絶対値) = 測定値 - 目標値
                End If
            Next i

            If Not UserSub.IsSpecialTrimType Or DGL = TRIM_MODE_MEAS Then
                ' 測定値表示(ログ画面)
                '                If (DGL = TRIM_MODE_MEAS) Or UserSub.IsTRIM_MODE_ITTRFT() Then
                If UserSub.IsTRIM_MODE_ITTRFT() Then
                    'V2.0.0.0⑤                    strMSG = strMSG & wkNOM(2).ToString("#0.00000") & strTANI(1) & " "
                    'V2.0.0.0⑤                    strMSG = strMSG & UserSub.ChangeOverFlow(wkVx(2).ToString("#0.00000")) & strTANI(1) & "  "
                    strMSG = strMSG & wkNOM(2).ToString(TARGET_DIGIT_DEFINE) & strTANI(1) & " "                                 'V2.0.0.0⑤
                    strMSG = strMSG & UserSub.ChangeOverFlow(wkVx(2).ToString(TARGET_DIGIT_DEFINE)) & strTANI(1) & "  "         'V2.0.0.0⑤
                    strMSG = strMSG & UserSub.ChangeOverFlow(WKdev(2).ToString("#0.0")) & strTANI(2)
                    If UserSub.IsTRIM_MODE_ITTRFT() Then
                        strMSG = strMSG & " " & strJUG(rn)
                    End If
                Else
                    '  "抵抗 目標値 測定値 誤差"
                    'V2.0.0.0⑤                    strMSG = strMSG & wkNOM(1).ToString("#0.00000") & strTANI(1) & " "
                    'V2.0.0.0⑤                    strMSG = strMSG & UserSub.ChangeOverFlow(wkVx(1).ToString("#0.00000")) & strTANI(1) & "  "
                    strMSG = strMSG & wkNOM(1).ToString(TARGET_DIGIT_DEFINE) & strTANI(1) & " "                                 'V2.0.0.0⑤
                    strMSG = strMSG & UserSub.ChangeOverFlow(wkVx(1).ToString(TARGET_DIGIT_DEFINE)) & strTANI(1) & "  "         'V2.0.0.0⑤
                    strMSG = strMSG & UserSub.ChangeOverFlow(WKdev(1).ToString("#0.0")) & strTANI(2)
                    If UserSub.IsTRIM_MODE_ITTRFT() Then
                        strMSG = strMSG & " " & strJUG(rn)
                    End If
                    'Else
                    '    '  "抵抗 目標値 測定値 誤差  目標値 測定値 誤差 判定"
                    '    For i = 1 To 2                              ' IT/FT
                    '        strMSG = strMSG & wkNOM(i).ToString("#0.000") & strTANI(1) & " "
                    '        strMSG = strMSG & wkVx(i).ToString("#0.000") & strTANI(1) & " "
                    '        strMSG = strMSG & WKdev(i).ToString("#0.000") & strTANI(2) & " "
                    '    Next i
                    '    strMSG = strMSG & " " & strJUG(rn)
                End If
                Call Z_PRINT(strMSG & vbCrLf)                               ' 測定値表示(ログ画面/印刷)
            End If
            'If (prf = 1) Then Call ObjPrt.Z_LPRINT(strMSG) ' 印刷

            '---------------------------------------------------------------------------
            '   トリミング結果をログファイルに出力する(ﾄﾘﾐﾝｸﾞﾓｰﾄﾞ時)
            '---------------------------------------------------------------------------
            strLOG = IO.Path.GetFileNameWithoutExtension(gsDataFileName) & ","              ' データ名
            strLOG = strLOG & stUserData.sLotNumber & ","                                   ' ロット番号
            strLOG = strLOG & stCounter.PlateCounter.ToString() & ","                       ' 基板番号
            'strLOG = strLOG & stCounter.BlockCounter.ToString() & ","                      ' ブロック番号
            strLOG = strLOG & stCounter.BlockCntX.ToString() & ","                          ' ブロックＸ
            strLOG = strLOG & stCounter.BlockCntY.ToString() & ","                          ' ブロックＹ
            strLOG = strLOG & stREG(rn).strRNO & ","                                        ' 抵抗名
            If UserSub.IsTrimType1() Or UserSub.IsTrimType4() Then                        ' 温度センサータイプの時目標値は変動する。'V2.0.0.0①sTrimType4()追加
                'V2.0.0.0⑤                strLOG = strLOG & UserSub.GetTRV().ToString("0.00000") & ","                ' 目標値
                strLOG = strLOG & UserSub.GetTRV().ToString(TARGET_DIGIT_DEFINE) & ","                ' 目標値 'V2.0.0.0⑤
            Else
                'V2.0.0.0⑤                strLOG = strLOG & stREG(rn).dblNOM.ToString("0.00000") & ","                ' 目標値
                strLOG = strLOG & stREG(rn).dblNOM.ToString(TARGET_DIGIT_DEFINE) & ","                ' 目標値 'V2.0.0.0⑤
            End If
            strLOG = strLOG & strTANI(1).Trim.ToString & ","                                ' 単位
            'V2.0.0.0⑤            strLOG = strLOG & UserSub.ChangeOverFlow(dblVX(1).ToString("0.00000")) & ","                            ' ＩＴ値
            strLOG = strLOG & UserSub.ChangeOverFlow(dblVX(1).ToString(TARGET_DIGIT_DEFINE)) & ","                            ' ＩＴ値 'V2.0.0.0⑤
            strLOG = strLOG & strTANI(1).Trim & ","                                         ' 単位
            'V2.0.0.0⑤            strLOG = strLOG & UserSub.ChangeOverFlow(dblVX(2).ToString("0.00000")) & ","                            ' ＦＴ値
            strLOG = strLOG & UserSub.ChangeOverFlow(dblVX(2).ToString(TARGET_DIGIT_DEFINE)) & ","                            ' ＦＴ値 'V2.0.0.0⑤
            strLOG = strLOG & strTANI(1).Trim & ","                                         ' 単位
            strLOG = strLOG & UserSub.ChangeOverFlow(WKdev(2).ToString("0.0")) & ","                                ' 判定
            strLOG = strLOG & strTANI(2).Trim                                               ' 単位
            If UserSub.IsTRIM_MODE_ITTRFT() Then
                strLOG = strLOG & "," & strJUG(rn).Trim                                         ' 判定
                If gisCutPosExecute Then                                                        ' カット位置補正有
                    strLOG = strLOG & "," & stPTN(rn).dblDRX.ToString("0.0000") & ","           ' 補正値Ｘ
                    strLOG = strLOG & stPTN(rn).dblDRY.ToString("0.0000") & ","                 ' 補正値Ｙ
                    strLOG = strLOG & gcPtnCorrval(rn).Replace("SAME", "")                      ' 一致度
                    'V2.1.0.0①                ElseIf DGL = TRIM_VARIATION_MEAS Then
                ElseIf DGL = TRIM_VARIATION_MEAS OrElse UserSub.IsCutVariationJudgeExecute() Then
                    strLOG = strLOG & ",,,"                 ' ,補正値Ｘ,補正値Ｙ, 一致度
                End If
            End If
            'V2.0.0.0②↓
            If DGL = TRIM_VARIATION_MEAS Then ' 測定値変動測定でＦＴの時
                strLOG = strLOG & "," & UserSub.ChangeOverFlow(dMeasVariationNOM(rn).ToString(TARGET_DIGIT_DEFINE)) & ","     ' トリミング後ＦＴ値
                If dMeasVariationNOM(rn) = 0.0 Or dblVX(2) > 10 ^ 37 Then
                    strLOG = strLOG & ",ppm"                         ' 変化量(ppm)
                Else
                    Dim dData As Double
                    dData = Math.Round(dMeasVariationDev(rn), 1)
                    strLOG = strLOG & UserSub.ChangeOverFlow(dData.ToString()) & ",ppm"                         ' 変化量(ppm)
                End If
            End If
            'V2.0.0.0②↑

            'V2.0.0.0⑨↓
            gdNOMforStatistical(GetResNumberInCircuit(rn)) = WKdev(2)  ' サーキット統計用FT値保存
            'V2.0.0.0⑨↑

            'V2.1.0.0①↓
            If UserSub.IsCutVariationJudgeExecute() Then
                If UserSub.CutVariationFinalJudgeNG() Then
                    strLOG = strLOG & "," & UserSub.CutVariationCutNoGet().ToString("0") & "," & UserSub.CutVariationRateGet().ToString("0.0000") & ""  '",カット番号,上昇率"
                Else
                    strLOG = strLOG & ",,"  '",カット番号,上昇率"
                End If
            End If
            'V2.1.0.0①↑

            Call Save_Log(strLOG)


            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.Disp_Final() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "メッセージ表示(ログ画面)/印刷"
    '''=========================================================================
    '''<summary>メッセージ表示(ログ画面)/印刷</summary>
    '''<param name="msg">(INP) 表示メッセージ(CRLF無し)</param>
    '''<param name="Bp"> (INP) Beep音(0:鳴らさない, 1:鳴らす)</param>
    '''=========================================================================
    Sub Msg_Disp(ByRef msg As String, ByRef Bp As Short)

        Dim a As Integer
        Dim s As String
        Dim strMSG As String                                            ' メッセージ編集域

        Try
            ' メッセージ表示(ログ画面)/印刷

            If (msg = "") Then Exit Sub
            a = &H67S
            s = msg & vbCrLf                                            ' エラー表示(ログ画面)
            Call Z_PRINT(s)
            If (prf = 1) Then
                s = msg
                'Call ObjPrt.Z_LPRINT(s)                                ' 印刷
            End If

            ' Beep音鳴動
            a = &HE1S
            If (Bp = 1) Then
                Call Beep()                                             ' Beep音
            End If
            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.Msg_Disp() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "トリミング結果表示(OK/NG)"
    '''=========================================================================
    '''<summary>トリミング結果表示(OK/NG)</summary>
    '''<param name="MD"> (INP) 処理モード(0=ｸﾘｱ, 1=結果表示, 2:ﾄﾘﾐﾝｸﾞ中)</param>
    '''<param name="sts">(INP) cFRS_NORMAL=正常, その他=エラー</param>
    '''<remarks>stsはMD=1の時有効</remarks>
    '''=========================================================================
    Public Sub Disp_Result(ByRef MD As Short, ByRef sts As Short)

        Dim strMSG As String                                            ' メッセージ編集域

        Try
            ' 「OK/NG」表示ｸﾘｱ
            If (MD = 0) Then                                            ' ﾓｰﾄﾞ = 「OK/NG」表示ｸﾘｱ ?
                ObjMain.LblSTS.BackColor = System.Drawing.Color.White   ' 背景色 = 白
                ObjMain.LblSTS.Text = ""                                ' ｷｬﾌﾟｼｮﾝｸﾘｱ

                ' ﾄﾘﾐﾝｸﾞ中
            ElseIf (MD = 2) Then                                        ' ﾓｰﾄﾞ = ﾄﾘﾐﾝｸﾞ中 ?
                ObjMain.LblSTS.BackColor = System.Drawing.Color.Yellow  ' 背景色 = 黄色
                ObjMain.LblSTS.Text = "処理中"

                ' トリミング結果(OK/NG)を表示する
            Else                                                        ' ﾓｰﾄﾞ = 結果表示
                If (Not UserSub.IsTRIM_MODE_ITTRFT() Or DGL = TRIM_MODE_POWER) Then              ' ﾄﾘﾐﾝｸﾞﾓｰﾄﾞ以外は背景色 = 白'V2.0.0.0② TRIM_MODE_POWER追加
                    ObjMain.LblSTS.BackColor = System.Drawing.Color.White ' 背景色 = 白
                    ObjMain.LblSTS.Text = ""                            ' ｷｬﾌﾟｼｮﾝｸﾘｱ
                    GoTo STP_END                                        ' Return
                End If

                Select Case (sts)
                    Case cFRS_NORMAL                                    ' 正常
                        ObjMain.LblSTS.BackColor = System.Drawing.Color.Lime ' 背景色 = 緑
                        ObjMain.LblSTS.Text = "ＯＫ"
                    Case Else                                           ' NG
                        ObjMain.LblSTS.BackColor = System.Drawing.Color.Red  ' 背景色 = 赤
                        ObjMain.LblSTS.Text = "ＮＧ"
                End Select

            End If

STP_END:
            ObjMain.LblSTS.Refresh()
            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.Disp_Result() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "トリミング結果表示(frmInfo画面)"
    ''' <summary>
    ''' ＮＧカウンターの更新
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Set_NG_Counter()
        stCounter.NG_Counter = stCounter.ITHigh + stCounter.ITLow + stCounter.ITOpen + stCounter.FTHigh + stCounter.FTLow + stCounter.FTOpen + stCounter.Pattern + stCounter.VaNG + stCounter.StdNg 'V1.2.0.0③ Pattern追加'V2.0.0.0②再測定変化量エラー追加'V2.0.0.0⑮スタンダード抵抗測定エラー
        stCounter.Total_NG_Counter = stCounter.Total_ITHigh + stCounter.Total_ITLow + stCounter.Total_ITOpen + stCounter.Total_FTHigh + stCounter.Total_FTLow + stCounter.Total_FTOpen + stCounter.Total_Pattern + stCounter.Total_VaNG + stCounter.Total_StdNg 'V1.2.0.0③ Pattern追加'V1.2.0.0③ Pattern追加'V2.0.0.0②再測定変化量エラー追加'V2.0.0.0⑮スタンダード抵抗測定エラー

        ' 'V2.2.0.0⑯↓
        If stMultiBlock.gMultiBlock <> 0 Then
            gObjFrmDistribute.CalcNgCounter()
        End If
        ' 'V2.2.0.0⑯↑

    End Sub
    '''=========================================================================
    '''<summary>トリミング結果表示(frmInfo画面)</summary>
    '''<param name="MD"> (INP) 処理モード
    '''                        0=全初期化, 1=ﾛｯﾄ番号ｶｳﾝﾄｱｯﾌﾟ, 2=ﾌﾟﾛｰﾌﾞON回数ｶｳﾝﾄｱｯﾌﾟ,3=ﾄﾘﾐﾝｸﾞ結果ｶｳﾝﾄｱｯﾌﾟ
    '''                       10=生産数初期化, 20=ﾌﾟﾛｰﾌﾞON回数初期化, 30=表示</param>
    '''<param name="Rlt">(INP) 結果(1:OK/NG, 2:ﾄﾘﾐﾝｸﾞ中, 3:ﾊﾟﾀｰﾝ認識NG, 10:ﾄﾘﾐﾝｸﾞ前(全ﾜｰｸ))</param>
    ''' <param name="rn">(INP) 処理抵抗番号</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    '''=========================================================================
    ''' V1.2.0.0②    Public Function Disp_frmInfo(ByRef MD As COUNTER, ByRef Rlt As Short) As Short
    Public Function Disp_frmInfo(ByRef MD As COUNTER, ByRef Rlt As Short, Optional ByVal rn As Integer = 0) As Short

        Dim strMSG As String

        Try
            '---------------------------------------------------------------------------
            '   処理モードによってトリミング数/OK数/不良率/ﾌﾟﾛｰﾌﾞON回数を設定する
            '---------------------------------------------------------------------------
            Select Case (MD)
                Case 0 ' 全初期化(未使用)
                    stCounter.TrimCounter = 0                               ' ｼﾘｱﾙ番号(ﾄﾘﾐﾝｸﾞ数)
                    stCounter.OK_Counter = 0                               ' OK枚数
                    stCounter.Probe_Counter = 0                               ' ﾌﾟﾛｰﾌﾞON回数(未使用)

                    'V2.2.0.0⑯↓
                    If stMultiBlock.gMultiBlock <> 0 Then
                        gObjFrmDistribute.ClearMultiLotCountData()
                    End If
                    'V2.2.0.0⑯↑

                Case COUNTER.PRODUCT_INIT ' 生産数初期化（ロット切り替え時）
                    ' 基板単位
                    stCounter.TrimCounter = 0                   ' ﾄﾘﾐﾝｸﾞ数(ﾜｰｸ投入数)
                    stCounter.OK_Counter = 0                    ' OK数
                    stCounter.NG_Counter = 0                    ' NG数
                    stCounter.ITHigh = 0                        ' 初期測定上限値異常
                    stCounter.ITLow = 0                         ' 初期測定下限値異常
                    stCounter.ITOpen = 0                        ' 測定値異常
                    stCounter.FTHigh = 0                        ' 最終測定上限値異常
                    stCounter.FTLow = 0                         ' 最終測定下限値異常
                    stCounter.FTOpen = 0                        ' 測定値異常
                    stCounter.Pattern = 0                       ' カット位置補正の判定 'V1.2.0.0③
                    stCounter.VaNG = 0                          ' 再測定変化量エラーV2.0.0.0②
                    stCounter.StdNg = 0                         ' スタンダード抵抗測定エラー 'V2.0.0.0⑮
                    stCounter.ValLow = 0                        ' カット毎上昇率変化Low異常        'V2.2.0.029
                    stCounter.ValHigh = 0                       ' カット毎上昇率変化High異常       'V2.2.0.029
                    ' ロット通算
                    stCounter.PlateCounter = 0                  ' 基板カウンター
                    stCounter.Total_TrimCounter = 0             ' 抵抗トータル処理数
                    stCounter.Total_OK_Counter = 0              ' OK数
                    stCounter.Total_NG_Counter = 0              ' NG数
                    stCounter.Total_ITHigh = 0                  ' 初期測定上限値異常
                    stCounter.Total_ITLow = 0                   ' 初期測定下限値異常
                    stCounter.Total_ITOpen = 0                  ' 測定値異常
                    stCounter.Total_FTHigh = 0                  ' 最終測定上限値異常
                    stCounter.Total_FTLow = 0                   ' 最終測定下限値異常
                    stCounter.Total_FTOpen = 0                  ' 測定値異常
                    stCounter.Total_Pattern = 0                 ' カット位置補正の判定 'V1.2.0.0③
                    stCounter.Total_VaNG = 0                    ' 再測定変化量エラーV2.0.0.0②
                    stCounter.Total_StdNg = 0                   ' スタンダード抵抗測定エラー 'V2.0.0.0⑮
                    stCounter.Total_ValLow = 0                  ' カット毎上昇率変化Low異常        'V2.2.0.029
                    stCounter.Total_ValHigh = 0                 ' カット毎上昇率変化High異常       'V2.2.0.029

                    'V2.2.0.0⑯↓
                    gObjFrmDistribute.ClearMultiLotCountData()
                    'V2.2.0.0⑯↑


                    gObjFrmDistribute.ClearCounter()            ' 分布図データ初期化 'V2.0.0.0⑨
                Case COUNTER.PROBE_INIT  ' ﾌﾟﾛｰﾌﾞ回数初期化
                    stCounter.Probe_Counter = 0                               ' ﾌﾟﾛｰﾌﾞ回数(未使用)

                Case COUNTER.ALLDATA_DISP ' トリミング数/OK数/不良率/ﾌﾟﾛｰﾌﾞON回数表示

                Case COUNTER.PROBE_UP ' ﾌﾟﾛｰﾌﾞON回数ｶｳﾝﾄｱｯﾌﾟ(未使用)
                    If (stCounter.Probe_Counter < Long.MaxValue) Then            ' < Long型の最大値 ?
                        stCounter.Probe_Counter = stCounter.Probe_Counter + 1
                    End If

                Case COUNTER.COUNTUP ' ﾄﾘﾐﾝｸﾞ結果ｶｳﾝﾄｱｯﾌﾟ (1=OK, 2=NG)
                    Select Case Rlt
                        Case COUNTER.OKNG_UP ' OK/NG
                            If UserSub.IsTRIM_MODE_ITTRFT() Then        ' ﾄﾘﾐﾝｸﾞﾓｰﾄﾞ(x0) ?
                                'V1.2.0.0②↓
                                Dim strJudge As String
                                'V2.0.0.0⑩                                If Not UserSub.IsTrimType2 Then
                                'V2.0.0.0⑩                                    strJudge = strJUG(rn)
                                'V2.0.0.0⑩                                Else
                                'V2.0.0.0⑩                                    strJudge = strJUG(0)
                                'V2.0.0.0⑩                                End If
                                'V2.0.0.0⑩↓
                                If UserSub.IsTrimType2() Or (UserSub.IsTrimType3() And UserBas.GetRCountExceptMeasure() > 1) Then         'V2.0.0.0⑩
                                    'サーキットの場合スキップ抵抗は呼ばれない
                                    strJudge = strJUG(0)
                                Else
                                    'サーキット以外の場合
                                    strJudge = strJUG(rn)
                                End If
                                'V2.0.0.0⑩↑
                                Select Case strJudge
                                    'V1.2.0.0②↑
                                    'V1.2.0.0②                                Select Case strJUG(0)
                                    Case JG_OK  ' ﾄﾘﾐﾝｸﾞOK ?
                                        'V2.0.0.0⑩                                        If Not UserSub.IsTrimType2 Then
                                        If UserSub.IsTrimType1() Or (UserSub.IsTrimType3() And UserBas.GetRCountExceptMeasure() = 1) Or UserSub.IsTrimType4() Then     'V2.0.0.0⑩
                                            stCounter.OK_Counter = stCounter.OK_Counter + 1
                                            stCounter.Total_OK_Counter = stCounter.Total_OK_Counter + 1
                                            ' 'V2.2.0.0⑯↓
                                            If stMultiBlock.gMultiBlock <> 0 Then
                                                gObjFrmDistribute.SetOkCounterMulti()
                                            End If
                                            ' 'V2.2.0.0⑯↑
                                        End If
                                    Case JG_IH  ' "IT-HI"               ' 初期判定ｴﾗｰ(ITHI)
                                        stCounter.ITHigh = stCounter.ITHigh + 1
                                        stCounter.Total_ITHigh = stCounter.Total_ITHigh + 1
                                        ' 'V2.2.0.0⑯↓
                                        If stMultiBlock.gMultiBlock <> 0 Then
                                            gObjFrmDistribute.SetITHICounterMulti()
                                        End If
                                            ' 'V2.2.0.0⑯↑
                                    Case JG_IL   ' "IT-LO"              ' 初期判定ｴﾗｰ(ITLO)
                                        stCounter.ITLow = stCounter.ITLow + 1
                                        stCounter.Total_ITLow = stCounter.Total_ITLow + 1
                                        ' 'V2.2.0.0⑯↓
                                        If stMultiBlock.gMultiBlock <> 0 Then
                                            gObjFrmDistribute.SetITLowCounterMulti()
                                        End If
                                            ' 'V2.2.0.0⑯↑
                                    Case JG_IO  ' "IT-OPEN"              ' 終了判定ｴﾗｰ(ITOPEN)
                                        stCounter.ITOpen = stCounter.ITOpen + 1
                                        stCounter.Total_ITOpen = stCounter.Total_ITOpen + 1
                                        ' 'V2.2.0.0⑯↓
                                        If stMultiBlock.gMultiBlock <> 0 Then
                                            gObjFrmDistribute.SetITOpenCounterMulti()
                                        End If
                                            ' 'V2.2.0.0⑯↑
                                    Case JG_FH   ' "FT-HI"              ' 終了判定ｴﾗｰ(FTHI)
                                        stCounter.FTHigh = stCounter.FTHigh + 1
                                        stCounter.Total_FTHigh = stCounter.Total_FTHigh + 1
                                        ' 'V2.2.0.0⑯↓
                                        If stMultiBlock.gMultiBlock <> 0 Then
                                            gObjFrmDistribute.SetFTHighCounterMulti()
                                        End If
                                            ' 'V2.2.0.0⑯↑
                                    Case JG_FL  ' "FT-LO"              ' 終了判定ｴﾗｰ(FTLO)
                                        stCounter.FTLow = stCounter.FTLow + 1
                                        stCounter.Total_FTLow = stCounter.Total_FTLow + 1
                                        ' 'V2.2.0.0⑯↓
                                        If stMultiBlock.gMultiBlock <> 0 Then
                                            gObjFrmDistribute.SetFTLOCounterMulti()
                                        End If
                                            ' 'V2.2.0.0⑯↑
                                    Case JG_FO  ' "IT-OPEN"              ' 終了判定ｴﾗｰ(FTOPEN)
                                        stCounter.FTOpen = stCounter.FTOpen + 1
                                        stCounter.Total_FTOpen = stCounter.Total_FTOpen + 1
                                        ' 'V2.2.0.0⑯↓
                                        If stMultiBlock.gMultiBlock <> 0 Then
                                            gObjFrmDistribute.SetFTOpenCounterMulti()
                                        End If
                                            ' 'V2.2.0.0⑯↑
                                        'V1.2.0.0③↓
                                    Case JG_PT
                                        stCounter.Pattern = stCounter.Pattern + 1
                                        stCounter.Total_Pattern = stCounter.Total_Pattern + 1
                                        'V1.2.0.0③↑
                                        ' 'V2.2.0.0⑯↓
                                        If stMultiBlock.gMultiBlock <> 0 Then
                                            gObjFrmDistribute.SetPatternCounterMulti()
                                        End If
                                            ' 'V2.2.0.0⑯↑
                                        'V2.0.0.0②↓
                                    Case JG_VA      ' 変動量
                                        stCounter.VaNG = stCounter.VaNG + 1
                                        stCounter.Total_VaNG = stCounter.Total_VaNG + 1
                                        'V2.0.0.0②↑
                                        ' 'V2.2.0.0⑯↓
                                        If stMultiBlock.gMultiBlock <> 0 Then
                                            gObjFrmDistribute.SetVaNGCounterMulti()
                                        End If
                                            ' 'V2.2.0.0⑯↑

                                        'V2.0.0.0⑮↓
                                    Case JG_STD     ' 温度センサースタンダード測定NG
                                        stCounter.StdNg = stCounter.StdNg + 1
                                        stCounter.Total_StdNg = stCounter.Total_StdNg + 1
                                        'V2.0.0.0⑮↑
                                        ' 'V2.2.0.0⑯↓
                                        If stMultiBlock.gMultiBlock <> 0 Then
                                            gObjFrmDistribute.SetStdNgCounterMulti()
                                        End If
                                            ' 'V2.2.0.0⑯↑
                                        'V2.1.0.0⑤↓
                                    Case JG_CUTVA     ' カット毎の抵抗値変化量判定ＮＧ
                                        If UserSub.GetVariationNGHiorLow() Then
                                            stCounter.FTLow = stCounter.FTLow + 1
                                            stCounter.Total_FTLow = stCounter.Total_FTLow + 1
                                            ' 'V2.2.0.0⑯↓
                                            If stMultiBlock.gMultiBlock <> 0 Then
                                                gObjFrmDistribute.SetFTLOCounterMulti()
                                            End If
                                            ' 'V2.2.0.0⑯↑
                                        Else
                                            stCounter.FTHigh = stCounter.FTHigh + 1
                                            stCounter.Total_FTHigh = stCounter.Total_FTHigh + 1
                                            ' 'V2.2.0.0⑯↓
                                            If stMultiBlock.gMultiBlock <> 0 Then
                                                gObjFrmDistribute.SetFTHighCounterMulti()
                                            End If
                                            ' 'V2.2.0.0⑯↑
                                        End If
                                        'V2.1.0.0⑤↑
                                        'V2.2.0.029↓
                                        ' 上昇率判定NGの場合
                                        If UserSub.GetVariationNGHiorLow() Then
                                            stCounter.ValLow = stCounter.ValLow + 1                            ' カット毎上昇率変化Low異常        'V2.2.0.029
                                            stCounter.Total_ValLow = stCounter.Total_ValLow + 1                ' カット毎上昇率変化Low異常        'V2.2.0.029
                                            ' 'V2.2.0.0⑯↓
                                            If stMultiBlock.gMultiBlock <> 0 Then
                                                gObjFrmDistribute.SetValLowCounterMulti()
                                            End If
                                            ' 'V2.2.0.0⑯↑
                                        Else
                                            stCounter.ValHigh = stCounter.ValHigh + 1                          ' カット毎上昇率変化High異常      'V2.2.0.029
                                            stCounter.Total_ValHigh = stCounter.Total_ValHigh + 1              ' カット毎上昇率変化High異常      'V2.2.0.029
                                            ' 'V2.2.0.0⑯↓
                                            If stMultiBlock.gMultiBlock <> 0 Then
                                                gObjFrmDistribute.SetValHighCounterMulti()
                                            End If
                                            ' 'V2.2.0.0⑯↑
                                        End If
                                        'V2.2.0.029↑

                                    Case Else
                                End Select
                                'V1.0.4.3⑪                                stCounter.NG_Counter = stCounter.ITHigh + stCounter.ITLow + stCounter.ITOpen + stCounter.FTHigh + stCounter.FTLow + stCounter.Total_FTOpen
                                'V1.2.0.0③                                stCounter.NG_Counter = stCounter.ITHigh + stCounter.ITLow + stCounter.ITOpen + stCounter.FTHigh + stCounter.FTLow + stCounter.FTOpen + stCounter.Pattern 'V1.2.0.0③ Pattern追加
                                'V1.2.0.0③                                stCounter.Total_NG_Counter = stCounter.Total_ITHigh + stCounter.Total_ITLow + stCounter.Total_ITOpen + stCounter.Total_FTHigh + stCounter.Total_FTLow + stCounter.Total_FTOpen + stCounter.Total_Pattern 'V1.2.0.0③ Pattern追加
                                Call Set_NG_Counter()                       'V1.2.0.0③
                            End If
                        Case COUNTER.INITIAL_DISP ' ﾄﾘﾐﾝｸﾞ前(全ﾜｰｸ))
                            'For i = 1 To MAXWNO
                            '    ObjMain.LblWk(i).BackColor = &HFFFFFF  ' ﾊﾞｯｸｶﾗｰ = 白
                            'Next i
                        Case COUNTER.SKIP
                            ' カウントしない
                    End Select
            End Select

            '---------------------------------------------------------------------------
            '   ロット番号/トリミング数/OK数/不良率/ﾌﾟﾛｰﾌﾞON回数表示(frmInfo画面)
            '---------------------------------------------------------------------------

            ObjMain.LblN_0.Text = stCounter.Total_TrimCounter.ToString("###,##0")     ' トリミング数表示
            ObjMain.LblN_1.Text = stCounter.Total_OK_Counter.ToString("###,##0")      ' OK数表示
            ObjMain.LblN_3.Text = stCounter.Total_NG_Counter.ToString("###,##0")      ' NG数表示

            If (stCounter.Total_TrimCounter = 0) Then                                           ' トリミング数 = 0
                ObjMain.LblN_2.Text = "0.00"                        ' OK率表示
                ObjMain.LblN_4.Text = "0.00"                        ' NG率表示
            Else
                ObjMain.LblN_2.Text = (stCounter.Total_OK_Counter / stCounter.Total_TrimCounter * 100).ToString("#0.00")
                ObjMain.LblN_4.Text = (stCounter.Total_NG_Counter / stCounter.Total_TrimCounter * 100).ToString("#0.00")
            End If

            ObjMain.LblITLONG.Text = stCounter.Total_ITLow.ToString
            ObjMain.LblITHING.Text = stCounter.Total_ITHigh.ToString
            ObjMain.LblITOPENNG.Text = stCounter.Total_ITOpen.ToString

            ObjMain.LblFTLONG.Text = stCounter.Total_FTLow.ToString
            ObjMain.LblFTHING.Text = stCounter.Total_FTHigh.ToString
            ObjMain.LblFTOPENNG.Text = stCounter.Total_FTOpen.ToString

            ObjMain.LblVALHING.Text = "(" & stCounter.Total_ValHigh.ToString & ")"          ' 上昇率Hi-NG  V2.2.0.029
            ObjMain.LblVALLONG.Text = "(" & stCounter.Total_ValLow.ToString & ")"           ' 上昇率Lo-NG  V2.2.0.029

            Call Form1.StatisticalDataDisp()        'V2.0.0.0⑨

            Return (cFRS_NORMAL)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.Disp_frmInfo() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cFRS_NORMAL)
        End Try
    End Function
#End Region

#Region "ログ出力"
    '''=========================================================================
    ''' <summary>ログ出力</summary>
    ''' <param name="strDAT"> (INP) ログデータ</param>
    ''' <remarks>トリミング処理時の測定値等のデータをセーブする</remarks>
    '''=========================================================================
    Public Sub Save_Log(ByRef strDAT As String)

        Dim strLOG As String                                            ' データ編集域
        Dim bFileExist As Boolean                                       'True:ファイルあり, False:ファイルなし
        Dim WS As IO.StreamWriter
        Dim FileName As String                                          'V2.0.0.0②gsLogFileNameをFileNameへ変更

        Try

            ' 初期処理
            If (Not UserSub.IsTRIM_MODE_ITTRFT() And DGL <> TRIM_MODE_MEAS) Then  ' トリミングと測定のモード以外はNOP
                Exit Sub
            End If


            'V2.0.0.0②↓
            If DGL = TRIM_VARIATION_MEAS Then ' 測定値変動測定の時
                FileName = gsLogFileName.Replace(".CSV", "-R.CSV")
            Else
                FileName = gsLogFileName
            End If
            'V2.0.0.0②↑

            If (FileName = "") Then
                Call Z_PRINT("UserBas.Save_Log() ログファイル名が設定されていません。")
                Exit Sub '                    ' ﾛｸﾞﾌｧｲﾙ名なしならNOP
            End If


            If IO.File.Exists(FileName) Then
                bFileExist = True                                       ' ファイル有り
            Else
                bFileExist = False
            End If

            WS = New IO.StreamWriter(FileName, True, System.Text.Encoding.GetEncoding("Shift-JIS"))
            If Not bFileExist Then                                      ' ファイル無い場合はヘッダ情報出力
                'V2.1.0.0①                If gisCutPosExecute Or DGL = TRIM_VARIATION_MEAS Then           ' カット位置補正有　'V2.0.0.0②TRIM_VARIATION_MEAS追加
                If gisCutPosExecute OrElse DGL = TRIM_VARIATION_MEAS OrElse UserSub.IsCutVariationJudgeExecute() Then           ' カット位置補正有　'V2.0.0.0②TRIM_VARIATION_MEAS追加 'V2.1.0.0①カット毎の抵抗値変化量判定追加
                    '                    strLOG = "日付,時間,データ名,ロット番号,基板番号,ブロック番号, 抵抗名,目標値,単位,ＩＴ値,単位,ＦＴ値,単位,誤差,単位,判定,補正値Ｘ，Ｙ，一致度"
                    strLOG = "日付,時間,データ名,ロット番号,基板番号,Ｘ,Ｙ, 抵抗名,目標値,単位,ＩＴ値,単位,ＦＴ値,単位,誤差,単位,判定,補正値Ｘ,Ｙ,一致度"
                Else
                    'strLOG = "日付,時間,データ名,ロット番号,基板番号,ブロック番号, 抵抗名,目標値,単位,ＩＴ値,単位,ＦＴ値,単位,誤差,単位,判定"
                    strLOG = "日付,時間,データ名,ロット番号,基板番号,Ｘ,Ｙ, 抵抗名,目標値,単位,ＩＴ値,単位,ＦＴ値,単位,誤差,単位,判定"
                End If

                'V2.1.0.0①↓
                If UserSub.IsCutVariationJudgeExecute() Then
                    strLOG = strLOG & ",カット番号,上昇率"
                End If
                'V2.1.0.0①↑

                'V2.0.0.0②↓
                If DGL = TRIM_VARIATION_MEAS Then ' 測定値変動測定でＦＴの時
                    strLOG = strLOG & ",トリミング後ＦＴ値,変化量,単位"
                End If
                'V2.0.0.0②↑
                WS.WriteLine(strLOG)
            End If

            ' ログデータ出力

            strLOG = DateTime.Now().ToString("yyyy/MM/dd,HH:mm:ss") & "," & strDAT                   ' 日付,時間を付加する。

            WS.WriteLine(strLOG)

            WS.Close()

        Catch ex As Exception
            Call Z_PRINT("UserBas.Save_Log() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub
#End Region

    '==========================================================================
    '   各コマンド処理
    '==========================================================================
#Region "レーザー調整処理"
    '''=========================================================================
    ''' <summary>レーザー調整処理</summary>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          上記以外      = エラー
    ''' </returns>
    '''=========================================================================
    Public Function User_LaserTeach(Optional ByVal bCalibration As Boolean = False) As Short        'V2.1.0.0② bCalibration 追加

        Dim r As Short
        Dim strMSG As String                                            ' データ編集域
        Dim iAttRot, iAttFix As Integer                                 ' ###1040③

        Try
            '---------------------------------------------------------------------------
            '   レーザー調整前処理
            '---------------------------------------------------------------------------
            ' パワー測定値表示ならﾊﾟﾜｰ調整位置へ移動する (ﾊﾟﾜｰﾒｰﾀ付き載物台の場合)
            ' (RMCTRL2 >=2 で 測定値表示の場合
            'If (gSysPrm.stRMC.giRmCtrl2 >= 2) And (gSysPrm.stRMC.giPMonLow = 1) Then
            '    ' ﾃｰﾌﾞﾙ移動(ﾊﾟﾜｰ調整位置)
            '    r = ObjSys.EX_SMOVE2(gSysPrm, gSysPrm.stRA2.gfATTTableOffsetX, gSysPrm.stRA2.gfATTTableOffsetY)
            '    If (r <> cFRS_NORMAL) Then                              ' ｴﾗｰ ?
            '        Return (r)                                          ' (注)軸ﾘﾐｯﾄ/ﾀｲﾑｱｳﾄｴﾗｰﾒｯｾｰｼﾞは表示済み
            '    End If
            '    ' BP移動(ﾊﾟﾜｰ調整位置)
            '    Call BSIZE(0, 0)                                        ' ﾌﾞﾛｯｸｻｲｽﾞ/BPｵﾌｾｯﾄ(ﾊﾟﾜｰ調整位置)設定
            '    r = ObjSys.EX_BPOFF(gSysPrm, gSysPrm.stRA2.gfATTBpOffsetX, gSysPrm.stRA2.gfATTBpOffsetY)
            '    If (r <> cFRS_NORMAL) Then                              ' ｴﾗｰ ?
            '        Return (r)                                          ' (注)軸ﾘﾐｯﾄ/ﾀｲﾑｱｳﾄｴﾗｰﾒｯｾｰｼﾞは表示済み
            '    End If
            '    r = ObjSys.EX_MOVE(gSysPrm, 10, 0, 1)                   ' BP移動(ﾊﾟﾜｰ調整位置)(絶対値)
            '    If (r <> cFRS_NORMAL) Then                              ' ｴﾗｰ ?
            '        Return (r)                                          ' (注)軸ﾘﾐｯﾄ/ﾀｲﾑｱｳﾄｴﾗｰﾒｯｾｰｼﾞは表示済み
            '    End If
            'Else
            If (gSysPrm.stIOC.giPM_Tp = 1) Then
                ' XYテーブルをﾊﾟﾜｰ調整位置へ移動する
                r = ObjSys.EX_SMOVE2(gSysPrm, gSysPrm.stRA2.gfATTTableOffsetX, gSysPrm.stRA2.gfATTTableOffsetY)
                If (r <> cFRS_NORMAL) Then                          ' エラー ?
                    Return (r)
                End If
                ' ブロックサイズ設定しfθセンタへ
                If (gSysPrm.stDEV.giBpSize = 6060) Then
                    r = ObjSys.EX_BSIZE(gSysPrm, 60.0, 20.0)
                ElseIf (gSysPrm.stDEV.giBpSize = 90) Then
                    r = ObjSys.EX_BSIZE(gSysPrm, 90.0, 90.0)
                    'V2.2.0.0⑧ ↓
                ElseIf (gSysPrm.stDEV.giBpSize = 40) Then
                    r = ObjSys.EX_BSIZE(gSysPrm, 40.0, 40.0)
                    'V2.2.0.0⑧ ↑
                Else
                    r = ObjSys.EX_BSIZE(gSysPrm, 80.0, 80.0)
                End If
                If (r <> cFRS_NORMAL) Then
                    Return (r)
                End If
                r = ObjSys.EX_MOVE(gSysPrm, gSysPrm.stRA2.gfATTBpOffsetX, gSysPrm.stRA2.gfATTBpOffsetY, 1)
                If (r <> cFRS_NORMAL) Then
                    Return (r)
                End If
            Else
                ' ﾊﾟｰﾂﾊﾝﾄﾞﾗをトリム位置に移動
                Call BSIZE(stPLT.zsx, stPLT.zsy)                        ' ブロックサイズ設定
                r = ObjSys.EX_START(gSysPrm, stPLT.z_xoff, stPLT.z_yoff, 0)
                If (r <> cFRS_NORMAL) Then                              ' ｴﾗｰ ?
                    Return (r)                                          ' (注)軸ﾘﾐｯﾄ/ﾀｲﾑｱｳﾄｴﾗｰﾒｯｾｰｼﾞは表示済み
                End If
                ' BP OFFSET値設定
                r = ObjSys.EX_BPOFF(gSysPrm, stPLT.BPOX, stPLT.BPOY)
                If (r <> cFRS_NORMAL) Then                              ' ｴﾗｰ ?
                    Return (r)                                          ' (注)軸ﾘﾐｯﾄ/ﾀｲﾑｱｳﾄｴﾗｰﾒｯｾｰｼﾞは表示済み
                End If
                r = ObjSys.EX_MOVE(gSysPrm, 10, 0, 1)                   ' BP移動(絶対値)
                If (r <> cFRS_NORMAL) Then                              ' ｴﾗｰ ?
                    Return (r)                                          ' (注)軸ﾘﾐｯﾄ/ﾀｲﾑｱｳﾄｴﾗｰﾒｯｾｰｼﾞは表示済み
                End If
            End If

            ' 初期設定(OcxSystemｵﾌﾞｼﾞｪｸﾄ, OcxUtilityｵﾌﾞｼﾞｪｸﾄ, Qレート[KHz](ｵﾌﾟｼｮﾝ), 設定パワー[W](ｵﾌﾟｼｮﾝ), 処理ﾓｰﾄﾞ(0=標準)(ｵﾌﾟｼｮﾝ), 実行ﾓｰﾄﾞ(0:手動,1:自動)(ｵﾌﾟｼｮﾝ))
            'V2.0.0.0            r = Form1.Ctl_LaserTeach1.SetUp((ObjSys), (Form1.Utility1), stLASER.intQR / 10.0#, stLASER.dblspecPower, 0, 0)
            r = Form1.Ctl_LaserTeach1.SetUp((ObjSys), (Form1.Utility1), stLASER.intQR / 10.0#, stLASER.dblspecPower, 0, 0, KND_USER)    'V2.0.0.0
            If (r <> cFRS_NORMAL) Then                                  ' ｴﾗｰ ?
                Return (r)
            End If

            iAttRot = gSysPrm.stRAT.giAttRot                            ' ###1040③
            iAttFix = gSysPrm.stRAT.giAttFix                            ' ###1040③

            '---------------------------------------------------------------------------
            '   レーザー調整処理
            '---------------------------------------------------------------------------
            Call ATTRESET()                                             'V2.1.0.0⑥
            'V2.1.0.0②↓
            If bCalibration Then
                '  アッテネータテーブルから全データ取得
                Dim stData(MAX_ATTENUATOR) As stATTENUATOR_TABLE
                Dim MaxNo As Integer = 0
                If Not UserSub.LaserCalibrationAllDataGet(MaxNo, stData) Then
                    Call Z_PRINT("アッテネータテーブルからデータを取得出来ませんでした。")
                    Return (cFRS_ERR_RST)
                End If
                ' 'V2.2.0.0①                r = Form1.Ctl_LaserTeach1.LaserCalibration(MaxNo, stData)
                r = Form1.Ctl_LaserTeach1.LaserCalibration(MaxNo, stData, 655, 70)  'V2.2.0.0①
                If r = cFRS_NORMAL Then
                    UserSub.LaserCalibrationAllDataWrite(stData)
                    Call Z_PRINT("アッテネータテーブルを更新しました。")
                End If
            Else
                'V2.1.0.0②↑
                'V2.2.0.0①　r = Form1.Ctl_LaserTeach1.LaserProc
                r = Form1.Ctl_LaserTeach1.LaserProc(655, 70)             'V2.2.0.0①
            End If                                                      'V2.1.0.0②
            If (r <> cFRS_NORMAL) Then                                  ' ｴﾗｰ ?
                Return (r)                                              ' (注)軸ﾘﾐｯﾄ/ﾀｲﾑｱｳﾄｴﾗｰﾒｯｾｰｼﾞは表示済み
            End If

            '---------------------------------------------------------------------------
            '   レーザー調整後処理
            '---------------------------------------------------------------------------
            ' RMCTRL2 >=2 の場合、 減衰率をシスパラより再表示("減衰率 = 99.9%")
            Call DllSysPrmSysParam_definst.GetSystemParameter(gSysPrm)   ' システム設定ファイルリード

            If iAttRot = gSysPrm.stRAT.giAttRot And iAttFix = gSysPrm.stRAT.giAttFix Then
                Return (cFRS_NORMAL)                                    ' 変更無しの場合 Return値 = 正常
            End If
            Call Form1.SetATTRateToScreen(False)                        ' ###1040③

            ' 測定値表示(メイン画面)
            ' RMCTRL2 >=3 で 測定値表示の場合に表示を更新する
            If (gSysPrm.stRMC.giRmCtrl2 >= 3) And (gSysPrm.stRMC.giPMonHi = 1) Then
                r = Form1.Ctl_LaserTeach1.GetMesPower(gSysPrm.stRAT.gfMesPower) ' 測定値取得
                If (r = cFRS_NORMAL) Then
                    If (gSysPrm.stTMN.giMsgTyp = 0) Then
                        strMSG = "レーザーパワー設定値　"
                        strMSG = strMSG & gSysPrm.stRAT.gfMesPower.ToString("##0.00") & "W"
                    Else
                        strMSG = "Laser Power "
                        strMSG = strMSG & gSysPrm.stRAT.gfMesPower.ToString("##0.00") & "W"
                    End If
                    Form1.LblMes.Text = strMSG                          ' 測定パワー[W]の表示
                End If
            End If

            Return (cFRS_NORMAL)                                        ' Return値 = 正常

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.User_LaserTeach() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Return値 = トラップエラー
        End Try
    End Function
#End Region

#Region "プローブティーチィング処理"
    '''=========================================================================
    ''' <summary>プローブティーチィング処理</summary>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          上記以外      = エラー
    ''' </returns>
    ''' <remarks>プローブとＸＹテーブルの位置をティーチングで設定する</remarks>
    '''=========================================================================
    Public Function User_ProbeTeaching() As Short

        Dim i As Short                                                  ' Work Index
        Dim rnotbl(MAXRNO) As Short                                     ' 抵抗名Table(Resistance Name)
        Dim pr4tbl(MAXRNO, 7) As Short                                  ' Probe No. Table
        Dim pr10tbl(MAXRNO) As Short                                    ' Slp Table
        Dim pr9tbl(MAXRNO) As Double                                    ' Nominal value Table
        Dim pr12tbl(MAXRNO, 2) As Double                                ' Final(initial) Test limit high,low Table(%)
        Dim cutspxtbl(MAXRNO) As Double                                 ' Cut start point x Table
        Dim cutspytbl(MAXRNO) As Double                                 ' Cut start point y Table
        Dim r As Short
        Dim W_bpox As Double                                            ' Beem Position X OFFSET(mm)
        Dim W_bpoy As Double                                            ' Beem Position Y OFFSET(mm)
        Dim W_Xoff As Double                                            ' Trim Position Offset X(mm)
        Dim W_Yoff As Double                                            ' Trim Position Offset Y(mm)
        Dim W_XCor As Double                                            ' ずれ量X X(mm)
        Dim W_YCor As Double                                            ' ずれ量Y Y(mm)
        '                                                               ' プローブ接触位置確認用ﾃｰﾌﾞﾙ
        Dim DataHI(MAXRNO, 2) As Double                                 '  Hの座標（左側の配列=抵抗番号,右側の配列=1:X座標、2:Y座標)
        Dim DataLO(MAXRNO, 2) As Double                                 '  Lの座標（左側の配列=抵抗番号,右側の配列=1:X座標、2:Y座標)
        Dim parModules As MainModules
        parModules = New MainModules
        Dim strMSG As String                                            ' メッセージ編集域
        Dim gCntRegData As Integer                          ' プローブコマンドに必要な情報を渡す(抵抗数)  'V2.2.2.0⑤ 

        Try
            '--------------------------------------------------------------------------
            '   初期処理
            '--------------------------------------------------------------------------
            W_bpox = stPLT.BPOX                                         ' Beem Position X OFFSET
            W_bpoy = stPLT.BPOY                                         ' Beem Position Y OFFSET
            W_Xoff = stPLT.z_xoff                                       ' Trim Position Offset X(mm)
            W_Yoff = stPLT.z_yoff                                       ' Trim Position Offset Y(mm)
            Call BSIZE(stPLT.zsx, stPLT.zsy)                            ' ブロックサイズ設定
            r = Move_Trimposition()                                     ' θ補正(ｵﾌﾟｼｮﾝ) & XYﾃｰﾌﾞﾙﾄﾘﾑ位置移動
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
                Return (r)                                              ' Return値設定
            End If

            Call ObjSys.Ilum_Ctrl(gSysPrm, Z1, ZOPT)                    ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ点灯(制御有時有効)

            ' 抵抗数分設定する
            For i = 1 To stPLT.RCount
                If (stREG(i).intPRH <> 0 And stREG(i).intSLP <> SLP_NG_MARK And stREG(i).intSLP <> SLP_OK_MARK) Then ' NGマーキングは除外する。'V1.0.4.3⑤ ＯＫマーキング(SLP_OK_MARK)追加
                    rnotbl(i) = i                                       ' 抵抗番号
                    pr4tbl(i, 1) = UserSub.ConvtChannel(stREG(i).intPRH) ' H  'V1.0.4.3⑨ ConvtChannel()追加
                    pr4tbl(i, 2) = UserSub.ConvtChannel(stREG(i).intPRL) ' L  'V1.0.4.3⑨ ConvtChannel()追加
                    pr4tbl(i, 3) = UserSub.ConvtChannel(stREG(i).intPRG) ' G1 'V1.0.4.3⑨ ConvtChannel()追加
                    pr4tbl(i, 4) = 0                                    ' G2
                    pr4tbl(i, 5) = 0                                    ' G3
                    pr4tbl(i, 6) = 0                                    ' G4
                    pr4tbl(i, 7) = 0                                    ' G5
                    pr9tbl(i) = stREG(i).dblNOM                         ' 目標値(Nominal value)
                    pr10tbl(i) = stREG(i).intSLP                        ' 電圧変化ｽﾛｰﾌﾟ(1:+  2:-  4:R)
                    If pr10tbl(i) = 5 Then                              ' 5:電圧測定のみ, 6:抵抗測定のみ)
                        pr10tbl(i) = 1
                    ElseIf pr10tbl(i) = 6 Then
                        pr10tbl(i) = 4
                    End If
                    'pr12tbl(i, 1) = 0                                  ' Test%(H) = 目標値*(1 + dblTol/100)
                    'pr12tbl(i, 2) = 0                                  ' Test%(L) = 目標値*(1 - .4)
                    'pr10tbl(i) = 0                                     ' 
                    'V2.0.0.0                    pr12tbl(i, 1) = 0                                   ' 
                    'V2.0.0.0                    pr12tbl(i, 2) = 0                                   ' 
                    'V2.0.0.0↓
                    ' Lowﾘﾐｯﾄ値/Highﾘﾐｯﾄ値を設定する
                    If (stREG(i).intMode = 0) Then                                      ' ﾓｰﾄﾞ = % ?
                        If (stREG(i).dblNOM = 0.0#) Then                                ' 目標値 = 0 ?
                            pr12tbl(i, 2) = stREG(i).dblITL * 0.01                      ' Lowﾘﾐｯﾄ値
                            pr12tbl(i, 1) = stREG(i).dblITH * 0.01                      ' Highﾘﾐｯﾄ値
                        Else
                            pr12tbl(i, 2) = stREG(i).dblNOM + (System.Math.Abs(stREG(i).dblNOM) * stREG(i).dblITL * 0.01) ' Lowﾘﾐｯﾄ値  (LOW = (NOM*(100+Lo)/100))
                            pr12tbl(i, 1) = stREG(i).dblNOM + (System.Math.Abs(stREG(i).dblNOM) * stREG(i).dblITH * 0.01) ' Highﾘﾐｯﾄ値 (HIGH= (NOM*(100+Hi)/100))
                        End If
                    Else                                                                ' ﾓｰﾄﾞ = 絶対値
                        pr12tbl(i, 2) = stREG(i).dblNOM + stREG(i).dblITL               ' Lowﾘﾐｯﾄ値
                        pr12tbl(i, 1) = stREG(i).dblNOM + stREG(i).dblITH               ' Highﾘﾐｯﾄ値
                    End If
                    'V2.0.0.0↑
                    cutspxtbl(i) = stREG(i).STCUT(1).dblSTX             ' Cut start point x
                    cutspytbl(i) = stREG(i).STCUT(1).dblSTY             ' Cut start point y

                    gCntRegData = gCntRegData + 1  ''V2.2.2.0⑤

                End If
            Next

            '--------------------------------------------------------------------------
            ' プローブティーチィングコントロールにデータを渡す
            ' ※パラメータ説明
            '   intrn() As Integer        : 抵抗番号
            '   dblBpx As Double          : BPオフセットＸ
            '   dblBpy As Double          : BPオフセットＹ
            '   dblBLsizex As Double      : ブロックサイズＸ
            '   dblBLsizey As Double      : ブロックサイズＹ
            '   dblStgOfx As Double       : XYテーブルオフセットＸ
            '   dblStgOfy As Double       : XYテーブルオフセットＹ
            '   dblZoff1 As Double        : Z OFF1(コンタクト)
            '   dblZoff2 As Double        : Z OFF2(ステップ位置)
            '   intpr4() As Integer       : プローブ番号
            '   intpr10() As Integer      : 電圧変化ｽﾛｰﾌﾟ
            '   dblpr9() As Double        : トリミング目標値
            '   dblpr12() As Double       : ファイナルテスト
            '   inttrmd As Integer        : トリムモード
            '   dblcpx() As Double        : カットスタートポイントＸ
            '   dblcpy() As Double        : カットスタートポイントＹ
            '   intfrtop As Integer       : グラフ表示位置(Dsplay Pos.)
            '   intfrleft As Integer      : グラフ表示位置(Dsplay Pos.)
            '   CorrectTrimPosX As Double : ﾄﾘﾑﾎﾟｼﾞｼｮﾝ補正値X(XYθ補正時のずれ量X)　(ｵﾌﾟｼｮﾝ ﾃｰﾌﾞﾙ移動機能有効時使用)
            '   CorrectTrimPosY As Double : ﾄﾘﾑﾎﾟｼﾞｼｮﾝ補正値Y(XYθ補正時のずれ量Y)  (同上)
            '   BNM_X As Integer          : ブロック数X
            '   BNM_Y As Integer          : ブロック数Y
            '   dblZ2off1 As Double       : Z2 OFF1(コンタクト)   (Z2有時有効)
            '   dblZ2off2 As Double       : Z2 OFF2(ステップ位置) (同上)
            '--------------------------------------------------------------------------
            ' 'V6.0.1.023　プローブに必要な情報を渡す
            ObjPrb.SetTrimmingDataName(gsDataFileName, gCntRegData)       'V2.2.2.0⑤

            ObjPrb.SetMainObject(parModules)                            ' 親モジュールのメソッドを設定する
            r = ObjPrb.Setup(rnotbl, W_bpox, W_bpoy, stPLT.zsx, stPLT.zsy, stPLT.z_xoff, stPLT.z_yoff, stPLT.Z_ZON, stPLT.Z_ZOFF, pr4tbl, pr10tbl, pr9tbl, pr12tbl, 0, cutspxtbl, cutspytbl, 7500, 105, dblCorrectX, dblCorrectY, stPLT.BNX, stPLT.BNY, stPLT.Z2_ZON, stPLT.Z2_ZOFF)
            '--------------------------------------------------------------------------
            ' θマニュアル調整　初期化(手動補正モードで、補正なしの時有効)
            '   注）Setup()の後にCallする
            '--------------------------------------------------------------------------
            ' 補正モードが手動で補正なし?
            If (stThta.iPP30 = 1) And (stThta.iPP31 = 0) Then
                stThta.iFlg = 1                                         ' θﾏﾆｭｱﾙ調整を有効とする
            Else
                stThta.iFlg = 0                                         ' θﾏﾆｭｱﾙ調整を無効とする
            End If

            ' θマニュアル調整　初期化
            r = ObjPrb.SetupTheta(stThta.iFlg, stThta.iPP30, stThta.iPP31, stThta.fPP53Min, stThta.fPP53Max, stThta.fTheta)

            '--------------------------------------------------------------------------
            '   プローブ接触位置確認機能(ｵﾌﾟｼｮﾝ)前処理
            '--------------------------------------------------------------------------
            ' 抵抗数分設定する
            For i = 1 To stPLT.RCount
                DataHI(i, 1) = i + 1.0#                                 ' Hの座標X
                DataHI(i, 2) = i + 2.0#                                 ' Hの座標Y
                DataLO(i, 1) = i + 3.0#                                 ' Lの座標X
                DataLO(i, 2) = i + 4.0#                                 ' Lの座標Y
            Next i

            ' プローブ接触位置確認機能初期設定(ｵﾌﾟｼｮﾝ)
            r = ObjPrb.SetupPrbPosChk(stPLT.RCount, stPLT.ADJX, stPLT.ADJY, 200, 9800, DataHI, DataLO)
            If (r <> cFRS_NORMAL) Then                                  ' ｴﾗｰ ?
                Return (r)
            End If

            '--------------------------------------------------------------------------
            '   プローブ調整処理
            '--------------------------------------------------------------------------
            ObjPrb.Visible = True
            r = ObjPrb.START()                                          ' プローブ調整(Probe Teaching)
            ObjPrb.Visible = False
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
                Return (r)
            End If
            Console.WriteLine("Probe.setup()  XOFF=" & stPLT.z_xoff & " YOFF=" & stPLT.z_yoff & " ZON=" & stPLT.Z_ZON)

            '--------------------------------------------------------------------------
            '   プローブ調整結果を取得する
            '--------------------------------------------------------------------------
            ' プローブ調整結果取得(ﾄﾘﾑ位置/Z位置更新)
            r = ObjPrb.GetResult(stPLT.z_xoff, stPLT.z_yoff, stPLT.Z_ZON, stPLT.Z2_ZON)
            If (r = 0) Then                                             ' OK戻り ?
                If (giAppMode = APP_MODE_PROBE Or giAppMode = APP_MODE_TEACH Or giAppMode = APP_MODE_CUTPOS) And (stPLT.TeachBlockX > 1 Or stPLT.TeachBlockY > 1) Then  ' ###1040①
                    Call ChangeTeachBlockPosition(stPLT.z_xoff, stPLT.z_yoff)                                                                                           ' ###1040①
                End If                                                                                                                                                  ' ###1040①
                FlgUpd = TriState.True                                  ' データ更新 Flag ON
                Console.WriteLine("Probe.GetResult()  XOFF=" & stPLT.z_xoff & " YOFF=" & stPLT.z_yoff & " ZON=" & stPLT.Z_ZON)
                ' PROBE2ｺﾏﾝﾄﾞならXY移動分BPもずらす(Teachｺﾏﾝﾄﾞがない為)
                If (giAppMode = APP_MODE_PROBE2) Then                   ' PROBE2ｺﾏﾝﾄﾞ ?
                    W_XCor = W_Xoff - (stPLT.z_xoff - dblCorrectX)      ' W_Xoff = XY移動分 X
                    W_YCor = W_Yoff - (stPLT.z_yoff - dblCorrectY)      ' W_Yoff = XY移動分 Y
                    stPLT.BPOX = W_bpox - W_XCor                        ' BP Offset X(mm)±ずれ量X
                    stPLT.BPOY = W_bpoy + W_YCor                        ' BP Offset Y(mm)±ずれ量Y
                End If
                '###1030 ADD START
                Dim bCutPosExecute As Boolean = False   ' カット位置補正の有り無しチェック
                For i = 1 To stPLT.PtnCount
                    If stPTN(i).PtnFlg <> CUT_PATTERN_NONE Then
                        bCutPosExecute = True
                    End If
                Next
                If bCutPosExecute Then
                    ' ###1030①プローブコマンド時にステージオフセットをずらした量だけＢＰオフセットに加算する。但し、カット位置補正有りの時のみ。
                    W_XCor = (stPLT.z_xoff - dblCorrectX) - W_Xoff      ' W_Xoff = XY移動分 X
                    W_YCor = (stPLT.z_yoff - dblCorrectY) - W_Yoff      ' W_Yoff = XY移動分 Y
                    stPLT.BPOX = W_bpox - W_XCor                        ' BP Offset X(mm)±ずれ量X
                    stPLT.BPOY = W_bpoy - W_YCor                        ' BP Offset Y(mm)±ずれ量Y
                End If
                '###1030 ADD END

                ' ﾄﾘﾑ位置更新(XYﾃｰﾌﾞﾙ補正分を引く)
                stPLT.z_xoff = stPLT.z_xoff - dblCorrectX               ' Trim Position Offset Y(mm)
                stPLT.z_yoff = stPLT.z_yoff - dblCorrectY               ' Trim Position Offset X(mm)
                ' システム変数設定(プローブON/OFF位置他)
                Call PROP_SET(stPLT.Z_ZON, stPLT.Z_ZOFF, gSysPrm.stDEV.gfTrimX, gSysPrm.stDEV.gfTrimY, gSysPrm.stDEV.gfSmaxX, gSysPrm.stDEV.gfSmaxY)
            End If

            ' θマニュアル調整結果取得
            r = ObjPrb.GetResultTheta(stThta.fTheta)
            If (r = 0) Then                                             ' OK戻り ?
                FlgUpd = TriState.True                                  ' データ更新 Flag ON
                Console.WriteLine("Probe.GetResultTheta()  ANG=" & stThta.fTheta)
            End If

            ' プローブティーチング結果(プローブ接触位置)を取得する(ｵﾌﾟｼｮﾝ)
            r = ObjPrb.GetResultPrbPosChk(DataHI, DataLO)
            If (r = 0) Then                                             ' OK戻り ?
                ' ここで結果を反映する
                FlgUpd = TriState.True                                  ' データ更新 Flag ON

            End If

            ' データファイルを更新する
            If (FlgUpd = TriState.True) Then                            ' データ更新 Flag ON ?
                'Call rData_save(gsDataFileName)                        ' ファイルへセーブ
            End If

            Return (cFRS_NORMAL)                                        ' Return値 = 正常

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.User_ProbeTeaching() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Return値 = トラップエラー
        End Try
    End Function
#End Region

#Region "スタートポジション ティーチング(TEACH(F8))処理"
    '''=========================================================================
    ''' <summary>スタートポジション ティーチング(TEACH(F8))処理</summary>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          上記以外      = エラー
    ''' </returns>
    ''' <remarks>BPオフセット値とトリミングスタート点をティーチングで設定する</remarks>
    '''=========================================================================
    Public Function User_teaching() As Integer

        Dim k As Integer                                                ' Work Index
        Dim t As Integer                                                ' Work Index
        Dim CutNum As Integer                                           ' カット総数
        Dim rn As Integer                                               ' 抵抗番号
        Dim cn As Integer                                               ' カット番号
        Dim r As Integer                                                ' Return Value From Function
        Dim dblTmpBpX As Double                                         ' Work
        Dim dblTmpBpY As Double                                         ' Work
        Dim dirL1 As Short                                              ' カット方向(1:-X←,2:+Y↑, 3:+X:→, 4:-Y↓)
        Dim dirL2 As Short
        Dim strMSG As String                                            ' メッセージ表示用域

        Dim parModules As MainModules
        parModules = New MainModules

        Try
            '--------------------------------------------------------------------------
            '   初期設定処理
            '--------------------------------------------------------------------------
            User_teaching = 0                                           ' Return値 = Normal
            Call BSIZE(stPLT.zsx, stPLT.zsy)                            ' ブロックサイズ設定
            'Call System1.EX_BPOFF(SysPrm, BPOX, BPOY)' BPｵﾌｾｯﾄ設定
            r = Move_Trimposition()                                     ' θ補正(ｵﾌﾟｼｮﾝ) & XYﾃｰﾌﾞﾙﾄﾘﾑ位置移動
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
                Return (r)                                              ' Return値設定
            End If
            ' BPｵﾌｾｯﾄ設定
            r = ObjSys.EX_BPOFF(gSysPrm, stPLT.BPOX, stPLT.BPOY)
            If (r <> cFRS_NORMAL) Then
                Return (r)                                              ' Return値設定
            End If

            ' パターン認識処理
            giTemplateGroup = -1                                        ' ﾃﾝﾌﾟﾚｰﾄｸﾞﾙｰﾌﾟ番号設定するため初期化
            r = Ptn_Match_Exe()                                         ' パターン認識実行
            If (r <> cFRS_NORMAL) Then
                Return (r)                                              ' Return値設定
            End If
            Call ObjSys.Ilum_Ctrl(gSysPrm, Z1, ZOPT)                     ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ点灯(制御有時有効)

            ' 配列の要素数をカット総数とする
            Call Get_Cut_Num(CutNum)                                    ' カット総数を取得する
            'ReDim testtbl(CutNum, 1)
            'ReDim sRName(CutNum)
            'ReDim dblStartPos(2, CutNum)

            Dim testtbl(CutNum, 1) As Short                             ' 抵抗番号,カット番号テーブル
            Dim sRName(CutNum) As String                                ' 抵抗番号
            Dim dblStartPos(2, CutNum) As Double                        ' 開始位置テーブル(Start Pos Table)

            ' 抵抗番号とカット番号をテーブルにセット
            t = 0                                                       ' ｶｯﾄ総数ｲﾝﾃﾞｯｸｽ
            For rn = 1 To stPLT.RCount                                  ' 抵抗数分設定する
                If UserModule.IsCutResistorIncMarking(stREG, rn) Then
                    k = stREG(rn).intTNN                                    ' k = 抵抗内カット数
                    r = Get_Cut_Num_Spt(rn)                                 ' r = 抵抗内ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ数
                    k = k + r                                               ' k = 抵抗内カット数 + 抵抗内ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ数

                    For cn = 1 To k                                         ' 抵抗内カット数分設定する
                        t = t + 1                                           ' ｶｯﾄ総数 += 1
                        sRName(t) = stREG(rn).strRNO                        ' 抵抗名(Resistance Name)
                        testtbl(t, 0) = CShort(rn)                                  ' 抵抗番号(Resistance No.)
                        testtbl(t, 1) = CShort(cn)                                  ' カット番号(Cut No.)
                    Next cn
                End If
            Next rn

            ' 'V2.2.0.0③↓
            If giBlueCrossDisable <> 0 Then
                ObjVdo.SetCrossLineVisible(False) '水色クロスライン非表示
            End If
            ' 'V2.2.0.0③↑

            ' クロスライン表示用  ###232
            r = ObjTch.SetCrossLineObject(gparModules)
            If (r <> cFRS_NORMAL) Then
                MsgBox("User_teaching() SetCrossLineObject ERROR")
            End If

            '--------------------------------------------------------------------------
            ' ティーチングコントロールにデータを渡す
            ' (Set Param. For Teaching Control)
            ' ※パラメータ説明(Param.))
            '   ①piPosNum() As Integer  : 抵抗番号カット番号(Cut No.)
            '   ②dblBpx As Double       : BPオフセットＸ(BP Offset)
            '   ③dblBpy As Double       : BPオフセットＹ
            '   ④dblBLsizex As Double   : ブロックサイズＸ(Block Size)
            '   ⑤dblBLsizey As Double   : ブロックサイズＹ
            '   ⑥sRN() As String        : 抵抗名(Resistance Name)
            '   ⑦BpDirXy As Integer     : BP 方向(0:XY NOM, 1:X REV, 2:Y REV, 3:XY REV)
            '--------------------------------------------------------------------------
            ' ティーチングＯＣＸに画像表示プログラムの表示位置を渡す
            ObjTch.dispXPos = FORM_X + ObjVdo.Location.X
            ObjTch.dispYPos = FORM_Y + ObjVdo.Location.Y


            ' ティーチングDLLにデータを渡す
            Call ObjTch.Setup(testtbl, stPLT.BPOX, stPLT.BPOY, stPLT.zsx, stPLT.zsy, sRName, gSysPrm.stDEV.giBpDirXy)

            Call Form1.VideoLibrary1.VideoStop()      'V2.2.0.0⑭

            '' カットトレース前処理(描画用)
            'Call ObjTch.IniCutTrace(ObjVdo.gpbxGazou, Form1.Picture2, Form1.Picture1)

            ' カット毎のstx, styデータをコントロールに渡す
            t = 1                                                       ' ｶｯﾄ総数ｲﾝﾃﾞｯｸｽ
            For rn = 1 To stPLT.RCount                                  ' 抵抗数分繰返す
                If UserModule.IsCutResistorIncMarking(stREG, rn) Then
                    k = stREG(rn).intTNN                                    ' k = 抵抗内カット数
                    'r = Get_Cut_Num_Spt(rn)                                ' 抵抗内ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ数
                    'k = k + r                                              ' k = 抵抗内カット数 + 抵抗内ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ数
                    For cn = 1 To k                                         ' 抵抗内カット数分繰返す
                        ' 角度→方向変換
                        't = t + 1                                          ' t = Table Index
                        '' ''Angle = stREG(kd, rn).STCUT(cn).intANG + 180
                        '' ''If Angle >= 360 Then Angle = Angle - 360
                        '' ''If Angle < 0 Then Angle = Angle + 360
                        '' ''dirL1 = (Angle / 90) + 1

                        ' '' '' カットトレースのためのカット方向を設定する
                        '' ''Call Cnv_Cut_Ang(rn, cn, dirL1, dirL2)
                        'V2.2.0.0②↓
                        ' UCutの場合方向を計算する 
                        Call Cnv_Cut_Dir(rn, cn, dirL1, dirL2)
                        'V2.2.0.0②↑

                        ' カットトレース用データを設定する
                        Call Sub_Cut_Setup(rn, cn, dirL1, dirL2, testtbl, t)

                    Next cn
                End If
            Next rn

            ''V2.2.0.0②↓
            '' クロスライン表示用
            'r = ObjTch.SetCrossLineObject(gparModules)
            'If r <> cFRS_NORMAL Then
            '    MsgBox("User.User_teaching() SetCrossLineObject ERROR")
            'End If
            ''V2.2.0.0②↑

            '--------------------------------------------------------------------------
            '   ティーチングコントロール表示
            '--------------------------------------------------------------------------
            ObjTch.ZOFF = stPLT.Z_ZOFF                                  ' Z PROBE OFF OFFSET(mm)
            ObjTch.ZON = stPLT.Z_ZON                                    ' Z PROBE ON OFFSET(mm)

            ' マーキング位置の表示のため、ビデオライブラリの描画オブジェクトを渡す
            ' 親モジュールのメソッドを設定する。
            ObjTch.SetMainObject(parModules)


            ObjTch.Visible = True
            ' ティーチング処理を実行する
            r = ObjTch.START()
            ' ティーチングのコントロールを非表示にする
            ObjTch.Visible = False

            '--------------------------------------------------------------------------
            '   ティーチング結果取得
            '--------------------------------------------------------------------------
            If (r = cFRS_NORMAL) Then                                   ' ティーチング処理正常終了 ?
                If ObjTch.Getresult(dblTmpBpX, dblTmpBpY, dblStartPos) = 0 Then
                    ' ビームポジションオフセット値更新
                    stPLT.BPOX = dblTmpBpX
                    stPLT.BPOY = dblTmpBpY

                    ' 抵抗数分stx,styを設定する
                    t = 0                                               ' t = Table ｲﾝﾃﾞｯｸｽ
                    For rn = 1 To stPLT.RCount                          ' 抵抗数分繰返す
                        If UserModule.IsCutResistorIncMarking(stREG, rn) Then

                            k = stREG(rn).intTNN                            ' k = 抵抗内カット数
                            'r = Get_Cut_Num_Spt(rn)                        ' r = 抵抗内ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ数
                            'k = k + r                                      ' k = 抵抗内カット数 + 抵抗内ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ数

                            For cn = 1 To k                                 ' 抵抗内カット数分設定する
                                t = t + 1                                   ' Table ｲﾝﾃﾞｯｸｽ += 1
                                ' ﾄﾘﾐﾝｸﾞｽﾀｰﾄ点 XY更新
                                stREG(rn).STCUT(cn).dblSTX = dblStartPos(1, t) - stPTN(rn).dblDRX
                                stREG(rn).STCUT(cn).dblSTY = dblStartPos(2, t) - stPTN(rn).dblDRY
                                ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄならﾄﾘﾐﾝｸﾞｽﾀｰﾄ点2 XY更新
                                If (stREG(rn).STCUT(cn).intCTYP = CNS_CUTP_SP) Then
                                    t = t + 1                               ' Table ｲﾝﾃﾞｯｸｽ += 1
                                    stREG(rn).STCUT(cn).dblSX2 = dblStartPos(1, t) - stPTN(rn).dblDRX
                                    stREG(rn).STCUT(cn).dblSY2 = dblStartPos(2, t) - stPTN(rn).dblDRY
                                    Call Get_Cut_Pitch(rn, cn)              ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄﾋﾟｯﾁ設定(stREG(kd, rn).STCUT(cn).dblDL3)
                                    'cn = cn + 1
                                End If
                            Next cn
                        End If
                    Next rn
                    FlgUpd = True                                       ' データ更新 Flag ON
                    'Call rData_save(gsDataFileName)                    ' ファイルへセーブ
                End If
            End If

            ObjCrossLine.CrossLineOff()                                 ' クロスラインの非表示

            ' 'V2.2.0.0③↓
            If giBlueCrossDisable <> 0 Then
                ObjVdo.SetCrossLineVisible(True) '水色クロスライン表示
            End If
            ' 'V2.2.0.0③↑

            'V2.2.0.0③            Return (cFRS_NORMAL)                                        ' Return値 = 正常
            Return (r)                                        ' Return値 ='V2.2.0.0③ 

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.User_teaching() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Return値 = トラップエラー
        End Try
    End Function
#End Region

#Region "カット総数を取得する"
    '''=========================================================================
    ''' <summary>カット総数を取得する</summary>
    ''' <param name="Num"> (OUT) カット総数</param>
    '''=========================================================================
    Private Sub Get_Cut_Num(ByRef Num As Short)

        Dim i As Short
        Dim Ct As Short
        Dim Ct2 As Short
        Dim strMSG As String                                            ' メッセージ編集域

        Try
            Ct = 0
            For i = 1 To stPLT.RCount                                   ' 抵抗数分繰返す
                Ct2 = Get_Cut_Num_Spt(i)                                ' 抵抗内ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ数
                Ct = stREG(i).intTNN + Ct + Ct2                         ' カット数 += 抵抗内カット数 + ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ数
            Next i
            Num = Ct
            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.Get_Cut_Num() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "抵抗内のｻｰﾍﾟﾝﾀｲﾝｶｯﾄ数を取得する"
    '''=========================================================================
    ''' <summary>抵抗内のｻｰﾍﾟﾝﾀｲﾝｶｯﾄ数を取得する</summary>
    ''' <param name="rn"> (INP) カット抵抗番号</param>
    ''' <returns>ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ数</returns>
    '''=========================================================================
    Private Function Get_Cut_Num_Spt(ByRef rn As Short) As Short

        Dim cn As Short
        Dim Ct As Short
        Dim strMSG As String                                            ' メッセージ編集域

        Try
            Ct = 0                                                      ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ数 = 0
            For cn = 1 To stREG(rn).intTNN                              ' 抵抗内カット数分繰返す
                If (stREG(rn).STCUT(cn).intCTYP = CNS_CUTP_SP) Then     ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ ?
                    Ct = Ct + 1                                         ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ数 += 1
                End If
            Next cn
            Return (Ct)                                                 ' 戻値 = ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ数

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.Get_Cut_Num_Spt() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (0)                                                  ' Return値 =  ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ数
        End Try
    End Function
#End Region

#Region "ｻｰﾍﾟﾝﾀｲﾝｶｯﾄﾋﾟｯﾁを取得する"
    '''=========================================================================
    ''' <summary>ｻｰﾍﾟﾝﾀｲﾝｶｯﾄﾋﾟｯﾁを取得する</summary>
    ''' <param name="rn"> (INP) 抵抗番号</param>
    ''' <param name="cn"> (INP) カット番号</param>
    ''' <remarks>ﾋﾟｯﾁは.dblDL3に設定する</remarks>
    '''=========================================================================
    Private Sub Get_Cut_Pitch(ByRef rn As Short, ByRef cn As Short)

        Dim Pit As Double                                               ' ｶｯﾄ開始座標X
        Dim strMSG As String                                            ' メッセージ編集域

        Try
            ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄﾋﾟｯﾁを求める
            If (stREG(rn).STCUT(cn).intANG2 = 90) Or (stREG(rn).STCUT(cn).intANG2 = 270) Then ' ｽﾃｯﾌﾟ方向がY方向 ?
                Pit = stREG(rn).STCUT(cn).dblSY2 - stREG(rn).STCUT(cn).dblSTY
            Else
                Pit = stREG(rn).STCUT(cn).dblSX2 - stREG(rn).STCUT(cn).dblSTX
            End If

            stREG(rn).STCUT(cn).dblDL3 = System.Math.Abs(Pit)
            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.Get_Cut_Pitch() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "カットトレースのためのカット方向を設定する"
    '''=========================================================================
    ''' <summary>カットトレースのためのカット方向を設定する</summary>
    ''' <param name="rn">   (INP) 抵抗番号</param>
    ''' <param name="cn">   (INP) カット番号</param>
    ''' <param name="dirL1">(I/O) カット方向1</param>
    ''' <param name="dirL2">(I/O) カット方向2</param>
    ''' <remarks>・STｶｯﾄ/IDXｶｯﾄ時(dirL2は返さない)
    '''            入力(dirL1) = カット方向(1:180°, 2:270°, 3:0°, 4:90°)
    '''            出力(dirL1) = カット方向(1:180°, 2: 90°, 3:0°, 4:270°)
    '''         ・ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ時
    '''            入力(dirL1) = カット方向(1:180°, 2:270°, 3:0°, 4:90°)
    '''            出力(dirL1) = カット方向(1:180°, 2: 90°, 3:0°, 4:270°)
    '''            出力(dirL2) = ｽﾃｯﾌﾟ方向 (1:180°, 2: 90°, 3:0°, 4:270°)
    '''         ・L ｶｯﾄ/ HOOK ｶｯﾄ時(dirL2は返さない)
    '''            入力(dirL1) = カット方向(1:180°, 2:270°, 3:0°, 4:90°)
    '''            出力(dirL1) = カット方向(1:-X-Y(↓←), 2:+Y-X(←↑), 3:+X+Y(→↑), 4:-Y+X(↓→),
    '''                                     5:-X+Y(↑←), 6:+Y+X(↑→), 7:+X-Y(→↓), 8:-Y-X (←↓))
    ''' </remarks>
    '''=========================================================================
    Private Sub Cnv_Cut_Ang(ByRef rn As Short, ByRef cn As Short, ByRef dirL1 As Short, ByRef dirL2 As Short)

        Dim strMSG As String                                            ' メッセージ編集域

        Try
            ' STｶｯﾄ/IDXｶｯﾄ時
            If (stREG(rn).STCUT(cn).intCTYP = CNS_CUTP_ST) Or (stREG(rn).STCUT(cn).intCTYP = CNS_CUTP_IX) Then
                Select Case dirL1
                    Case 1
                        dirL1 = 1   ' 180°(-X←)
                    Case 4
                        dirL1 = 2   '  90°(+Y↑)
                    Case 3
                        dirL1 = 3   '   0°(+X→)
                    Case 2
                        dirL1 = 4   ' 270°(-Y↓)
                    Case Else
                        dirL1 = 1   ' 180°(-X←)
                End Select
            End If

            ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ時
            If (stREG(rn).STCUT(cn).intCTYP = CNS_CUTP_SP) Then
                Select Case dirL1                                   ' 最初のｶｯﾄ方向
                    Case 1
                        dirL1 = 1                                   ' 180°(-X←)
                        Select Case (stREG(rn).STCUT(cn).intANG2)   ' ｽﾃｯﾌﾟ方向
                            Case 90 : dirL2 = 2                     ' ｽﾃｯﾌﾟ方向 = 2(↑)
                            Case 270 : dirL2 = 4                    ' ｽﾃｯﾌﾟ方向 = 4(↓)
                        End Select
                    Case 4
                        dirL1 = 2                                   '  90°(+Y↑)
                        Select Case (stREG(rn).STCUT(cn).intANG2)   ' ｽﾃｯﾌﾟ方向
                            Case 0 : dirL2 = 3                      ' ｽﾃｯﾌﾟ方向 = 3(→)
                            Case 180 : dirL2 = 1                    ' ｽﾃｯﾌﾟ方向 = 1(←)
                        End Select
                    Case 3
                        dirL1 = 3                                   '   0°(+X→)
                        Select Case (stREG(rn).STCUT(cn).intANG2)   ' ｽﾃｯﾌﾟ方向
                            Case 90 : dirL2 = 2                     ' ｽﾃｯﾌﾟ方向 = 2(↑)
                            Case 270 : dirL2 = 4                    ' ｽﾃｯﾌﾟ方向 = 4(↓)
                        End Select
                    Case 2
                        dirL1 = 4                                   ' 270°(-Y↓)
                        Select Case (stREG(rn).STCUT(cn).intANG2)   ' ｽﾃｯﾌﾟ方向
                            Case 0 : dirL2 = 3                      ' ｽﾃｯﾌﾟ方向 = 3(→)
                            Case 180 : dirL2 = 1                    ' ｽﾃｯﾌﾟ方向 = 1(←)
                        End Select
                    Case Else
                        Call ObjSys.TrmMsgBox(gSysPrm, CStr(CDbl("Cut Direction Error Direction = ") + dirL1), MsgBoxStyle.OkOnly, My.Application.Info.Title)
                End Select
            End If

            If (stREG(rn).STCUT(cn).intCTYP = CNS_CUTP_L) Then               ' L CUT
                Select Case dirL1
                    Case 1 ' 始めの移動方向 = 180°(-X←) ?
                        Select Case (stREG(rn).STCUT(cn).intANG2)   ' Lターン後の移動方向
                            Case 90 : dirL1 = 5                     ' カット方向 = 5:-X+Y (↑←)
                            Case 270 : dirL1 = 1                    ' カット方向 = 1:-X-Y (↓←)
                        End Select

                    Case 4 ' 始めの移動方向 =  90°(+Y↑) ?
                        Select Case (stREG(rn).STCUT(cn).intANG2)   ' Lターン後の移動方向
                            Case 0 : dirL1 = 6                      ' カット方向 = 6:+Y+X (↑→)
                            Case 180 : dirL1 = 2                    ' カット方向 = 2:+Y-X (←↑)
                        End Select

                    Case 3 ' 始めの移動方向 =   0°(+X→) ?
                        Select Case (stREG(rn).STCUT(cn).intANG2)   ' Lターン後の移動方向
                            Case 270 : dirL1 = 7                    ' カット方向 = 7:+X-Y (→↓)
                            Case 90 : dirL1 = 3                     ' カット方向 = 3:+X+Y (→↑)
                        End Select

                    Case 2 ' 始めの移動方向 = 270°(-Y↓) ?
                        Select Case (stREG(rn).STCUT(cn).intANG2)   ' Lターン後の移動方向
                            Case 180 : dirL1 = 8                    ' カット方向 = 8:-Y-X (←↓)
                            Case 0 : dirL1 = 4                      ' カット方向 = 4:-Y+X (↓→)
                        End Select

                    Case Else ' 始めの移動方向 = 180°(-X←)
                        Select Case (stREG(rn).STCUT(cn).intANG2)   ' Lターン後の移動方向
                            Case 90 : dirL1 = 5                     ' カット方向 = 5:-X+Y (↑←)
                            Case 270 : dirL1 = 1                    ' カット方向 = 1:-X-Y (↓←)
                        End Select
                End Select
            End If

            'If (stREG(kd, rn).STCUT.intCTYP(cn) = 17) Then             ' U CUT
            '    Select Case dirL1
            '        Case 1                                  ' 始めの移動方向 = 180°(-X←) ?
            '            If (stREG(kd, rn).STCUT.intDIR(cn) = 1) Then       ' Lターン後の移動方向 = 時計方向 ?
            '                '                    dirL1 = 5                       ' カット方向 = 5:-X+Y (↑←)
            '                dirL1 = 1                       ' カット方向 = 1:-X-Y (↓←)
            '            Else
            '                dirL1 = 1                       ' カット方向 = 1:-X-Y (↓←)
            '            End If
            '        Case 4                                  ' 始めの移動方向 =  90°(+Y↑) ?
            '            If (stREG(kd, rn).STCUT.intDIR(cn) = 1) Then       ' Lターン後の移動方向 = 時計方向 ?
            '                '                    dirL1 = 8                       ' カット方向 = 8:+Y+X (↑→)
            '                dirL1 = 4                       ' カット方向 = 4:+Y-X (←↑)
            '            Else
            '                dirL1 = 4                       ' カット方向 = 4:+Y-X (←↑)
            '            End If
            '        Case 3                                  ' 始めの移動方向 =   0°(+X→) ?
            '            If (stREG(kd, rn).STCUT.intDIR(cn) = 1) Then       ' Lターン後の移動方向 = 時計方向 ?
            '                '                    dirL1 = 6                       ' カット方向 = 6:+X-Y (→↓)
            '                dirL1 = 2                       ' カット方向 = 2:+X+Y (→↑)
            '            Else
            '                dirL1 = 2                       ' カット方向 = 2:+X+Y (→↑)
            '            End If
            '        Case 2                                  ' 始めの移動方向 = 270°(-Y↓) ?
            '            If (stREG(kd, rn).STCUT.intDIR(cn) = 1) Then       ' Lターン後の移動方向 = 時計方向 ?
            '                '                    dirL1 = 7                       ' カット方向 = 7:-Y-X (←↓)
            '                dirL1 = 3                       ' カット方向 = 3:-Y+X (↓→)
            '            Else
            '                dirL1 = 3                       ' カット方向 = 3:-Y+X (↓→)
            '            End If
            '        Case Else                               ' 始めの移動方向 = 180°(-X←)
            '            If (stREG(kd, rn).STCUT.intDIR(cn) = 1) Then       ' Lターン後の移動方向 = 時計方向 ?
            '                '                    dirL1 = 5                       ' カット方向 = 5:-X+Y (↑←)
            '                dirL1 = 1                       ' カット方向 = 1:-X-Y (↓←)
            '            Else
            '                dirL1 = 1                       ' カット方向 = 1:-X-Y (↓←)
            '            End If
            '    End Select
            'End If

            '' SCAN CUT (ｶｯﾄ方向 1:-X, 2:+X, 3:-Y, 4:+Y)/ｽﾃｯﾌﾟ方向(1:+X, 2:-X, 3:+Y, 4:-Y)
            'If (stREG(kd, rn).STCUT.intCTYP(cn) = 5) Then              ' SCAN CUT ?
            '    Select Case dirL1
            '        Case 1                                  ' 始めの移動方向 = 180°(-X←) ?
            '            dirL1 = 1                           ' カット方向 = 1:-X(←)
            '            If (stREG(kd, rn).STCUT.intDIR(cn) = 1) Then       ' ｽﾃｯﾌﾟ方向 = 時計方向 ?
            '                dirL2 = 3                       ' ｽﾃｯﾌﾟ方向 = 3:+Y(↑)
            '            Else
            '                dirL2 = 4                       ' ｽﾃｯﾌﾟ方向 = 4:-Y(↓)
            '            End If
            '        Case 4                                  ' 始めの移動方向 =  90°(+Y↑) ?
            '            dirL1 = 4                           ' カット方向 = 4:+Y(↑)
            '            If (stREG(kd, rn).STCUT.intDIR(cn) = 1) Then       ' ｽﾃｯﾌﾟ方向 = 時計方向 ?
            '                dirL2 = 1                       ' ｽﾃｯﾌﾟ方向 = 1:+X(→)
            '            Else
            '                dirL2 = 2                       ' ｽﾃｯﾌﾟ方向 = 2:-X(←)
            '            End If
            '        Case 3                                  ' 始めの移動方向 =   0°(+X→) ?
            '            dirL1 = 2                           ' カット方向 = 2:+X(→)
            '            If (stREG(kd, rn).STCUT.intDIR(cn) = 1) Then       ' ｽﾃｯﾌﾟ方向 = 時計方向 ?
            '                dirL2 = 4                       ' ｽﾃｯﾌﾟ方向 = 4:-Y(↓)
            '            Else
            '                dirL2 = 3                       ' ｽﾃｯﾌﾟ方向 = 3:+Y(↑)
            '            End If
            '        Case 2                                  ' 始めの移動方向 = 270°(-Y↓) ?
            '            dirL1 = 3                           ' カット方向 = 3:-Y(↓)
            '            If (stREG(kd, rn).STCUT.intDIR(cn) = 1) Then       ' ｽﾃｯﾌﾟ方向 = 時計方向 ?
            '                dirL2 = 2                       ' ｽﾃｯﾌﾟ方向 = 2:-X(←)
            '            Else
            '                dirL2 = 1                       ' ｽﾃｯﾌﾟ方向 = 1:+X(→)
            '            End If
            '    End Select
            'End If
            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.Cnv_Cut_Ang() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "各カットトレースのためセットアップ処理"
    '''=========================================================================
    ''' <summary>各カットトレースのためセットアップ処理</summary>
    ''' <param name="rn">   (INP) 抵抗番号</param>
    ''' <param name="cn">   (INP) カット番号</param>
    ''' <param name="dirL1">(I/O) カット方向1</param>
    ''' <param name="dirL2">(I/O) カット方向2</param>
    ''' <param name="testtbl"></param>
    ''' <param name="t"></param>
    '''=========================================================================
    Private Sub Sub_Cut_Setup(ByRef rn As Short, ByRef cn As Short, ByRef dirL1 As Short, ByRef dirL2 As Short, ByRef testtbl(,) As Short, ByRef t As Short)

        Dim strMSG As String                                ' メッセージ編集域
        Dim CutLen As Double = 0.0
        Dim dirL1_SP As Short
        'V2.0.0.6①        Dim sTurnDir As Short
        Dim tmpCutInfo As Teaching.CutInfo
        Dim qrate As Integer
        Dim spd As Double

        Try

            Select Case stREG(rn).STCUT(cn).intCTYP ' ｶｯﾄ形状(1:ｽﾄﾚｰﾄ, 2:Lｶｯﾄ, 3:ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ)
                Case CNS_CUTP_ST, CNS_CUTP_ST_TR ' ストレートカット(カット方向 1:-X←, 2:+Y↑, 3:+X→ ,4:-Y↓)
                    'Call ObjTch.SetupCutSST(testtbl(t, 0), testtbl(t, 1), stREG(rn).STCUT(cn).dblSTX + stPTN(rn).dblDRX, _
                    '        stREG(rn).STCUT(cn).dblSTY + stPTN(rn).dblDRY, stREG(rn).STCUT(cn).dblDL2, stREG(rn).STCUT(cn).intANG, TrimDef.CNS_CUTP_ST)

                    If stREG(rn).STCUT(cn).intCUT = 1 Then
                        ' トラッキング
                        qrate = stREG(rn).STCUT(cn).intQF1
                        spd = stREG(rn).STCUT(cn).dblV1
                    Else
                        'インデックス
                        qrate = stREG(rn).STCUT(cn).intQF1
                        spd = stREG(rn).STCUT(cn).dblV1

                    End If

                    Call ObjTch.SetupCutSST_SP(testtbl(t, 0), testtbl(t, 1), stREG(rn).STCUT(cn).dblSTX + stPTN(rn).dblDRX,
                            stREG(rn).STCUT(cn).dblSTY + stPTN(rn).dblDRY, stREG(rn).STCUT(cn).dblDL2, stREG(rn).STCUT(cn).intANG, qrate, spd, TrimDef.CNS_CUTP_ST)

                Case CNS_CUTP_L 'Ｌカット(カット方向 1:-X-Y, 2:+Y-X, 3:+X+Y ,4:-Y+X, 5:-X+Y, 6:+Y+X, 7:+X-Y ,8:-Y-X)


                    ' Lﾀｰﾝ後移動方向(1:時計方向,2:反時計方向)を求める
                    'V2.0.0.6①                    sTurnDir = Get_Cut_Dir(stREG(rn).STCUT(cn).intANG, stREG(rn).STCUT(cn).intANG2)

                    'If stREG(rn).STCUT(cn).intANG < stREG(rn).STCUT(cn).intANG2 Then
                    '    sTurnDir = CCW
                    'Else
                    '    sTurnDir = CW
                    'End If


                    'Call ObjTch.SetupCutSL(testtbl(t, 0), testtbl(t, 1), stREG(rn).STCUT(cn).dblSTX + stPTN(rn).dblDRX, _
                    '        stREG(rn).STCUT(cn).dblSTY + stPTN(rn).dblDRY, stREG(rn).STCUT(cn).dblDL2, 0.0#, stREG(rn).STCUT(cn).dblDL3, _
                    ''            stREG(rn).STCUT(cn).intANG, stREG(rn).STCUT(cn).intANG2)
                    'Call ObjTch.SetupCutSL(testtbl(t, 0), testtbl(t, 1), stREG(rn).STCUT(cn).dblSTX + stPTN(rn).dblDRX, _
                    '        stREG(rn).STCUT(cn).dblSTY + stPTN(rn).dblDRY, stREG(rn).STCUT(cn).dblDL2, 0.0#, stREG(rn).STCUT(cn).dblDL3, _
                    '            stREG(rn).STCUT(cn).intANG, stREG(rn).STCUT(cn).intANG2, TrimDef.CNS_CUTP_L)    'V2.0.0.6①
                    ''V2.0.0.6①                                stREG(rn).STCUT(cn).intANG, sTurnDir, TrimDef.CNS_CUTP_L)

                    tmpCutInfo.dAngle1 = stREG(rn).STCUT(cn).dAngle(1)
                    tmpCutInfo.dAngle2 = stREG(rn).STCUT(cn).dAngle(2)
                    tmpCutInfo.dAngle3 = stREG(rn).STCUT(cn).dAngle(3)
                    tmpCutInfo.dAngle4 = stREG(rn).STCUT(cn).dAngle(4)
                    tmpCutInfo.dAngle5 = stREG(rn).STCUT(cn).dAngle(5)
                    tmpCutInfo.dAngle6 = stREG(rn).STCUT(cn).dAngle(6)
                    tmpCutInfo.dAngle7 = stREG(rn).STCUT(cn).dAngle(7)

                    tmpCutInfo.dblL1 = stREG(rn).STCUT(cn).dCutLen(1)
                    tmpCutInfo.dblL2 = stREG(rn).STCUT(cn).dCutLen(2)
                    tmpCutInfo.dblL3 = stREG(rn).STCUT(cn).dCutLen(3)
                    tmpCutInfo.dblL4 = stREG(rn).STCUT(cn).dCutLen(4)
                    tmpCutInfo.dblL5 = stREG(rn).STCUT(cn).dCutLen(5)
                    tmpCutInfo.dblL6 = stREG(rn).STCUT(cn).dCutLen(6)
                    tmpCutInfo.dblL7 = stREG(rn).STCUT(cn).dCutLen(7)

                    tmpCutInfo.dblQrate1 = stREG(rn).STCUT(cn).dQRate(1)
                    tmpCutInfo.dblQrate2 = stREG(rn).STCUT(cn).dQRate(2)
                    tmpCutInfo.dblQrate3 = stREG(rn).STCUT(cn).dQRate(3)
                    tmpCutInfo.dblQrate4 = stREG(rn).STCUT(cn).dQRate(4)
                    tmpCutInfo.dblQrate5 = stREG(rn).STCUT(cn).dQRate(5)
                    tmpCutInfo.dblQrate6 = stREG(rn).STCUT(cn).dQRate(6)
                    tmpCutInfo.dblQrate7 = stREG(rn).STCUT(cn).dQRate(7)

                    tmpCutInfo.dblspd1 = stREG(rn).STCUT(cn).dSpeed(1)
                    tmpCutInfo.dblspd2 = stREG(rn).STCUT(cn).dSpeed(2)
                    tmpCutInfo.dblspd3 = stREG(rn).STCUT(cn).dSpeed(3)
                    tmpCutInfo.dblspd4 = stREG(rn).STCUT(cn).dSpeed(4)
                    tmpCutInfo.dblspd5 = stREG(rn).STCUT(cn).dSpeed(5)
                    tmpCutInfo.dblspd6 = stREG(rn).STCUT(cn).dSpeed(6)
                    tmpCutInfo.dblspd7 = stREG(rn).STCUT(cn).dSpeed(7)


                    Call ObjTch.SetupCutSL6(testtbl(t, 0), testtbl(t, 1), stREG(rn).STCUT(cn).dblSTX + stPTN(rn).dblDRX,
                            stREG(rn).STCUT(cn).dblSTY + stPTN(rn).dblDRY, tmpCutInfo, TrimDef.CNS_CUTP_L)

                Case CNS_CUTP_SP ' ｻｰﾍﾟﾝﾀｲﾝｶｯﾄ(カット方向 1:-X←, 2:+Y↑, 3:+X→ ,4:-Y↓)
                    Call ObjTch.SetupCutSST(testtbl(t, 0), testtbl(t, 1), stREG(rn).STCUT(cn).dblSTX + stPTN(rn).dblDRX, _
                            stREG(rn).STCUT(cn).dblSTY + stPTN(rn).dblDRY, stREG(rn).STCUT(cn).dblDL2, stREG(rn).STCUT(cn).intANG, TrimDef.CNS_CUTP_ST)
                    t = t + 1
                    Call Cnv_ANG2(stREG(rn).STCUT(cn).intANG, dirL1_SP) ' カット方向を反対方向に変換

                    Call ObjTch.SetupCutSST(testtbl(t, 0), testtbl(t, 1), stREG(rn).STCUT(cn).dblSX2 + stPTN(rn).dblDRX, _
                            stREG(rn).STCUT(cn).dblSY2 + stPTN(rn).dblDRY, stREG(rn).STCUT(cn).dblDL2, dirL1_SP, TrimDef.CNS_CUTP_ST)


                Case CNS_CUTP_IX ' インデックスカット(カット方向 1:-X←, 2:+Y↑, 3:+X→ ,4:-Y↓)
                    ' SetupCutIX(抵抗番号, ｶｯﾄ番号, ｶｯﾄ方向, ｽﾀｰﾄ位置X, ｽﾀｰﾄ位置Y, ｶｯﾄ長1, IDX回数, 測定ﾓｰﾄﾞ(未使用))
                    'Call ObjTch.SetupCutIX(testtbl(t, 0), testtbl(t, 1), dirL1, stREG(kd, rn).STCUT(cn).dblSTX, stREG(kd, rn).STCUT(cn).dblSTY, gdblDL1(i, 1), stREG(kd, rn).STCUT.intIXN(cn), stREG(kd, rn).STCUT.intTMM(cn)) '

                    ' SetupCutIX(抵抗番号, ｶｯﾄ番号, ｶｯﾄ方向, ｽﾀｰﾄ位置X, ｽﾀｰﾄ位置Y, ｶｯﾄ長1, IDX回数, 測定ﾓｰﾄﾞ(未使用))
                    'Call ObjTch.SetupCutIX(testtbl(t, 0), testtbl(t, 1), dirL1, stREG(kd, rn).STCUT(cn).dblSTX, stREG(kd, rn).STCUT(cn).dblSTY, _
                    'stREG(kd, rn).STCUT(cn).dblDL1(1), stREG(kd, rn).STCUT(cn).intIXN(1), stREG(kd, rn).STCUT(cn).intTMM)

                    ''V2.2.0.0① Call ObjTch.SetupCutIX(testtbl(t, 0), testtbl(t, 1), dirL1, stREG(rn).STCUT(cn).dblSTX, stREG(rn).STCUT(cn).dblSTY,
                    ''V2.2.0.0① stREG(rn).STCUT(cn).dblDL1(1), 3, stREG(rn).STCUT(cn).intTMM)

                    Call ObjTch.SetupCutIX(testtbl(t, 0), testtbl(t, 1), dirL1, stREG(rn).STCUT(cn).dblSTX, stREG(rn).STCUT(cn).dblSTY,
                    stREG(rn).STCUT(cn).dblDL1(1), 3, stREG(rn).STCUT(cn).intTMM, 0, 0, 0)                       'V2.2.0.0①


                Case CNS_CUTP_M ' 文字マーキング
                    ' SetupCutM(抵抗番号, ｶｯﾄ番号, 方向方向(1:-X, 2:+X, 3:-Y, 4:+Y), ｽﾀｰﾄ位置X, ｽﾀｰﾄ位置Y, 倍率, 文字列長)
                    ' 'V2.2.0.0⑬Call ObjTch.SetupCutMStr(testtbl(t, 0), testtbl(t, 1), stREG(rn).STCUT(cn).intANG, stREG(rn).STCUT(cn).dblSTX, stREG(rn).STCUT(cn).dblSTY, stREG(rn).STCUT(cn).dblDL2, 3, stREG(rn).STCUT(cn).cFormat)
                    'V2.2.1.7③ ↓
                    Dim MartStr As String = ""
　                    If stREG(rn).intSLP = SLP_MARK Then
                        'マーク印字の場合は、固定長＋開始番号をそのまま
                        MartStr = stREG(rn).STCUT(cn).cMarkFix & stREG(rn).STCUT(cn).cMarkStartNum
                    Else
                        'マーク印字でない場合は文字
                        MartStr = stREG(rn).STCUT(cn).cFormat
                    End If
                    Call ObjTch.SetupCutMStr(testtbl(t, 0), testtbl(t, 1), stREG(rn).STCUT(cn).intANG, stREG(rn).STCUT(cn).dblSTX, stREG(rn).STCUT(cn).dblSTY, stREG(rn).STCUT(cn).dblDL2, 3, MartStr, stREG(rn).STCUT(cn).intQF1, stREG(rn).STCUT(cn).dblV1)      'V2.2.0.0⑬
                    ' Call ObjTch.SetupCutMStr(testtbl(t, 0), testtbl(t, 1), stREG(rn).STCUT(cn).intANG, stREG(rn).STCUT(cn).dblSTX, stREG(rn).STCUT(cn).dblSTY, stREG(rn).STCUT(cn).dblDL2, 3, stREG(rn).STCUT(cn).cFormat, stREG(rn).STCUT(cn).intQF1, stREG(rn).STCUT(cn).dblV1)      'V2.2.0.0⑬
                    'V2.2.1.7③ ↑
                    '    cn = cn + 1

                    'Case 3 ' フックカット(カット方向 1:-X-Y, 2:+Y-X, 3:+X+Y ,4:-Y+X, 5:-X+Y, 6:+Y+X, 7:+X-Y ,8:-Y-X)
                    '    ' SetupCutHK(抵抗番号,ｶｯﾄ番号,ｶｯﾄ方向,ｽﾀｰﾄ位置X, ｽﾀｰﾄ位置Y,ｶｯﾄ長1,R1半径(未使用), ﾀｰﾝﾎﾟｲﾝﾄ(未使用), ｶｯﾄ長2, r2(-1固定), ﾌｯｸｶｯﾄ移動量)
                    '    '                Call ObjTch.SetupCutHK(testtbl(t, 0), testtbl(t, 1), dirL1, stREG(kd, rn).STCUT(cn).dblSTX, stREG(kd, rn).STCUT(cn).dblSTY, gdblDL1(cn), 0#, 100#, stREG(kd, rn).STCUT(cn).dblDL2, -1, dl3(cn))
                    '    '            Call ObjTch.SetupCutHK(testtbl(t, 0), testtbl(t, 1), dirL1, stREG(kd, rn).STCUT(cn).dblSTX, stREG(kd, rn).STCUT(cn).dblSTY, gdblDL1(i, 1), 0#, 100#, stREG(kd, rn).STCUT(cn).dblDL2, -1, stREG(kd, rn).STCUT(cn).dblDL3)

                    'Case 4 ' インデックスカット(カット方向 1:-X←, 2:+Y↑, 3:+X→ ,4:-Y↓)
                    '    ' SetupCutIX(抵抗番号, ｶｯﾄ番号, ｶｯﾄ方向, ｽﾀｰﾄ位置X, ｽﾀｰﾄ位置Y, ｶｯﾄ長1, IDX回数, 測定ﾓｰﾄﾞ(未使用))
                    '    Call ObjTch.SetupCutIX(testtbl(t, 0), testtbl(t, 1), dirL1, stREG(kd, rn).STCUT(cn).dblSTX, stREG(kd, rn).STCUT(cn).dblSTY, gdblDL1(i, 1), stREG(kd, rn).STCUT.intIXN(cn), stREG(kd, rn).STCUT.intTMM(cn)) '

                    'Case 5 ' スキャンカット(ｶｯﾄ方向 1:-X, 2:+X, 3:-Y, 4:+Y)/ｽﾃｯﾌﾟ方向(1:+X, 2:-X, 3:+Y, 4:-Y)
                    '    '            SetupCutSC(抵抗番号, ｶｯﾄ番号, ｶｯﾄ方向, ｽﾀｰﾄ位置X, ｽﾀｰﾄ位置Y, ｶｯﾄ長1, ﾋﾟｯﾁ, ｽﾃｯﾌﾟ方向, 本数)
                    '    Call ObjTch.SetupCutSC(testtbl(t, 0), testtbl(t, 1), dirL1, stREG(kd, rn).STCUT(cn).dblSTX, stREG(kd, rn).STCUT(cn).dblSTY, gdblDL1(i, 1), stREG(kd, rn).STCUT(cn).dblDL2, dirL2, stREG(kd, rn).STCUT.intIXN(cn))
                    'V2.2.0.0② ↓
                Case CNS_CUTP_U
                    ' SetupCutU(抵抗番号, ｶｯﾄ番号, ｶｯﾄ方向, ｽﾀｰﾄ位置X, ｽﾀｰﾄ位置Y, ｶｯﾄ長1, R1半径, ｶｯﾄ長2, Lﾀｰﾝ後移動方向(未使用))
                    Call ObjTch.SetupCutU(testtbl(t, 0), testtbl(t, 1), dirL1, stREG(rn).STCUT(cn).dblSTX, stREG(rn).STCUT(cn).dblSTY, stREG(rn).STCUT(cn).dUCutL1, stREG(rn).STCUT(cn).dblUCutR1, stREG(rn).STCUT(cn).dUCutL2, 0, stREG(rn).STCUT(cn).dblUCutR2, stREG(rn).STCUT(cn).dblUCutV1, stREG(rn).STCUT(cn).intUCutQF1)
                    'V2.2.0.0② ↑
                    'Case 17 ' Uカット(カット方向 1:-X-Y, 2:+X+Y, 3:-Y+X ,4:+Y-X)
                    '    ' SetupCutU(抵抗番号, ｶｯﾄ番号, ｶｯﾄ方向, ｽﾀｰﾄ位置X, ｽﾀｰﾄ位置Y, ｶｯﾄ長1, R1半径, ｶｯﾄ長2, Lﾀｰﾝ後移動方向(未使用))
                    '    Call ObjTch.SetupCutU(testtbl(t, 0), testtbl(t, 1), dirL1, stREG(kd, rn).STCUT(cn).dblSTX, stREG(kd, rn).STCUT(cn).dblSTY, gdblDL1(i, 1), 0.0#, stREG(kd, rn).STCUT(cn).dblDL2, stREG(kd, rn).STCUT.intDIR(cn))

                    'Case 35 'Z(NO CUT)
                    '    ' SetupCutZ(抵抗番号, ｶｯﾄ番号, ｽﾀｰﾄ位置X, ｽﾀｰﾄ位置Y)
                    '    Call ObjTch.SetupCutZ(testtbl(t, 0), testtbl(t, 1), stREG(kd, rn).STCUT(cn).dblSTX, stREG(kd, rn).STCUT(cn).dblSTY)

                    'Case 36 ' 円弧カット
                    '    dirA1 = stREG(kd, rn).STCUT.intANG(cn)                                  ' 始めの移動角度
                    '    ' SetupCutCir(抵抗番号, ｶｯﾄ番号, ｽﾀｰﾄ位置X, ｽﾀｰﾄ位置Y,円弧部の半径,円弧の角度, 始めの移動角度)
                    '    '            Call ObjTch.SetupCutCir(testtbl(t, 0), testtbl(t, 1), stREG(kd, rn).STCUT(cn).dblSTX, stREG(kd, rn).STCUT(cn).dblSTY, 1, 45, -270)
                    '    Call ObjTch.SetupCutCir(testtbl(t, 0), testtbl(t, 1), stREG(kd, rn).STCUT(cn).dblSTX, stREG(kd, rn).STCUT(cn).dblSTY, gdblDL1(i, 1), stREG(kd, rn).STCUT(cn).dblDL2, dirA1)

                Case Else
                    Call ObjSys.TrmMsgBox(gSysPrm, "Cut Type Error Type = " & Str(stREG(rn).STCUT(cn).intCTYP), MsgBoxStyle.OkOnly, My.Application.Info.Title)
            End Select
            t = t + 1
            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.Sub_Cut_Setup() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "カット位置補正のためのパターン登録処理"
    '''=========================================================================
    ''' <summary>カット位置補正のためのパターン登録処理</summary>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          上記以外      = エラー
    ''' </returns>
    '''=========================================================================
    Public Function User_CutpositionTeach() As Short

        Dim sRName() As String                                          ' 抵抗名
        'Dim dcpcor() As Double                                         ' 位置X,Y
        Dim iGrpNum(MAXRGN) As Short                                    ' ﾊﾟﾀｰﾝｸﾞﾙｰﾌﾟ番号
        Dim iPtnNum(MAXRGN) As Short                                    ' パターン番号
        Dim i As Short                                                  ' COUNTER
        Dim j As Short                                                  ' COUNTER
        Dim r As Integer                                                ' 関数RETURN値
        Dim s As String
        Dim strMSG As String                                            ' メッセージ編集域

        Try
            '---------------------------------------------------------------------------
            '   初期設定処理
            '---------------------------------------------------------------------------
            ChDir(My.Application.Info.DirectoryPath)

            ' XYテーブル移動
            Call BSIZE(stPLT.zsx, stPLT.zsy)                            ' ブロックサイズ設定
            r = Move_Trimposition()                                     ' θ補正(ｵﾌﾟｼｮﾝ) & XYﾃｰﾌﾞﾙﾄﾘﾑ位置移動
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
                Return (r)                                              ' Return値設定
            End If

            Call ObjSys.Ilum_Ctrl(gSysPrm, Z1, ZOPT)                    ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ点灯(制御有時有効)
            ObjVdo.frmLeft = ObjMain.Text2.Left                         ' フォームの表示位置
            ObjVdo.frmTop = ObjMain.Text2.Top
            ObjVdo.PATTERNGROUP = giTemplateGroup                       ' パターンテンプレートグループナンバー = 1
            ObjVdo.RNASTMPNUM = True
            ObjVdo.ZON = stPLT.Z_ZON
            ObjVdo.ZOFF = stPLT.Z_ZOFF

            j = stPLT.PtnCount                                          ' 配列要素数をパターン登録数とする
            ReDim sRName(j)
            Dim dcpcor(2, j) As Double                                  ' 位置X,Y

            ' 名前と位置を設定する(カットデータ)
            For i = 1 To stPLT.PtnCount                                 ' パターン登録数分設定する
                'sRName(i) = "PTN" & Str(i)                             ' 名前("PTNx") ###777
                sRName(i) = stREG(i).strRNO                             ' 抵抗名称     ###777
                dcpcor(1, i) = stPTN(i).dblPosX                         ' パターン位置X
                dcpcor(2, i) = stPTN(i).dblPosY                         ' パターン位置Y
                iGrpNum(i) = stPTN(i).intGRP                            ' ﾊﾟﾀｰﾝｸﾞﾙｰﾌﾟ番号
                iPtnNum(i) = stPTN(i).intPTN                            ' パターン番号
            Next i

            ' ビデオライブラリー用パラメータ設定
            ObjVdo.pp32_x = 0.0#                                        ' CorStgPos1X
            ObjVdo.pp32_y = 0.0#                                        ' CorStgPos1Y
            ObjVdo.pp34_x = stPLT.BPOX                                  ' Bp Offset X
            ObjVdo.pp34_y = stPLT.BPOY                                  ' Bp Offset Y
            'ObjVdo.pfTrim_x = Form1.trimmer.LOADPOSX                   ' 画面表示位置x
            'ObjVdo.pfTrim_y = Form1.trimmer.LOADPOSY                   ' 画面表示位置y
            ObjVdo.pfTrim_x = gSysPrm.stDEV.gfTrimX                     ' トリム位置x
            ObjVdo.pfTrim_y = gSysPrm.stDEV.gfTrimY                     ' トリム位置y
            ObjVdo.pfStgOffX = stPLT.z_xoff                             ' Trim Position Offset Y(mm)
            ObjVdo.pfStgOffY = stPLT.z_yoff                             ' Trim Position Offset Y(mm)
            ObjVdo.pfBlock_x = stPLT.zsx                                ' Block Size x(mm)
            ObjVdo.pfBlock_y = stPLT.zsy                                ' Block Size y(mm)
            ObjVdo.zwaitpos = stPLT.Z_ZOFF                              ' Z PROBE OFF OFFSET(mm)

            ' テンプレートグループ選択/テンプレート番号設定
            Call ObjVdo.SelectTemplateGroup(giTemplateGroup)
            r = ObjVdo.SetTemplateNum_EX(stPLT.PtnCount, iPtnNum, iGrpNum)

            'Call BSIZE(0, 0)                                           ' ブロックサイズ = 0
            'Call System1.EX_BPOFF(SysPrm, 0, 0)                        ' BPオフセット   = 0

            'V2.2.0.026↓
            If giRecogPointCorrLine <> 0 Then
                Dim xPos As Double
                Dim yPos As Double
                Call ZGETBPPOS(xPos, yPos)                              ' BP現在位置取得
                ObjCrossLine.CrossLineDispXY(xPos, yPos)                ' クロスライン表示
            End If
            'V2.2.0.026↑

            '---------------------------------------------------------------------------
            '   補正位置ティーチング処理
            ' 　(Rame, 位置XY, ﾌﾞﾛｯｸｻｲｽﾞx, ﾌﾞﾛｯｸｻｲｽﾞy, BpｵﾌｾｯﾄX, BpｵﾌｾｯﾄY)
            '---------------------------------------------------------------------------
            r = ObjVdo.CutPosTeach(sRName, dcpcor, stPLT.zsx, stPLT.zsy, stPLT.BPOX, stPLT.BPOY)

            'V2.2.0.026↓
            If ObjCrossLine Is Nothing = False Then
                ObjCrossLine.CrossLineOff()                             ' クロスライン非表示
            End If
            'V2.2.0.026↑

            If (r <= cFRS_VIDEO_INI) And (r >= cFRS_MVC_10) Then        ' Video.OCXエラー ?
                Select Case r
                    Case cFRS_VIDEO_INI
                        s = "VIDEOLIB: Not initialized."                ' "初期化が行われていません。"
                    Case cFRS_VIDEO_FRM
                        s = "VIDEOLIB: Form Display Now"                ' "フォームの表示中です。"
                    Case cFRS_VIDEO_PRP
                        s = "VIDEOLIB: Invalid property value."         ' "プロパティ値が不正です"
                    Case cFRS_VIDEO_UXP
                        s = "VIDEOLIB: Unexpected error"                ' "予期せぬエラー"
                    Case Else
                        s = "VIDEOLIB: Unexpected error 2"
                End Select
                Call ObjSys.TrmMsgBox(gSysPrm, s, MsgBoxStyle.OkOnly, My.Application.Info.Title)

            ElseIf (r < cFRS_NORMAL) Then                               ' 非常停止等エラー ?
                Return (r)
            End If

            ' パターン位置XYの補正値を取得する
            r = ObjVdo.Getresult(dcpcor)                                ' 結果取得
            If (r = cFRS_NORMAL) Then                                   ' 正常終了 ?
                For i = 1 To stPLT.PtnCount                             ' パターン登録数分設定する
                    stPTN(i).dblPosX = dcpcor(1, i)                     ' パターン位置X設定
                    stPTN(i).dblPosY = dcpcor(2, i)                     ' パターン位置Y設定
                Next

                ' データファイルへセーブ
                FlgUpd = TriState.True                                  ' データ更新 Flag ON
                'Call rData_save(gsDataFileName)                        ' ファイル自動更新(ｾｰﾌﾞ(F2)が無い時)
            End If

            '---------------------------------------------------------------------------
            '   終了処理
            '---------------------------------------------------------------------------
            Call BSIZE(stPLT.zsx, stPLT.zsy)                            ' ブロックサイズ/BP オフセット設定
            Call ObjSys.EX_BPOFF(gSysPrm, stPLT.BPOX, stPLT.BPOY)
            Call ObjSys.EX_PROBOFF(gSysPrm)                             ' ZプローブをOFF位置に移動
            Call ObjSys.EX_SBACK(gSysPrm)                               ' ﾊﾟｰﾂﾊﾝﾄﾞﾗをロード位置に戻す
            Return (cFRS_NORMAL)                                        ' Return値 = 正常

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.User_CutpositionTeach() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Return値 = トラップエラー
        End Try
    End Function
#End Region

#Region "カット位置補正処理(BP補正)"
    '''=========================================================================
    ''' <summary>カット位置補正処理(BP補正)</summary>
    ''' <param name="i"> (INP) パターン登録データの添字(1 ORG)</param>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          上記以外      = エラー
    ''' </returns>
    ''' <remarks>パターンマッチィングを行いズレ量X,YをstPTN(i).dblDRX/DRYに保存する</remarks>
    '''=========================================================================
    Public Function CutPosCorrect(ByRef i As Short) As Short

        Dim crx As Double                                               ' ずれ量X
        Dim cry As Double                                               ' ずれ量Y
        Dim r As Integer                                                ' 関数値
        Dim rtn As Integer = cFRS_NORMAL                                ' 関数値
        Dim nTemplateNum As Short                                       ' テンプレート番号
        Dim corrval As Double                                           ' 相関値
        Dim Thresh As Double                                            ' 閾値
        Dim strMSG As String                                            ' メッセージ表示用域

        Try
            '---------------------------------------------------------------------------
            '   初期設定処理
            '---------------------------------------------------------------------------
            ChDir(My.Application.Info.DirectoryPath)
            'If (stPLT.PtnFlg = 0) Then                                  ' パターン認識無し ?
            '    stPTN(i).dblDRX = 0                                     ' ズレ量X保存
            '    stPTN(i).dblDRY = 0                                     ' ズレ量Y保存
            '    Return (cFRS_NORMAL)                                    ' Return値 = 正常
            'End If

            ' ビデオライブラリー用パラメータ設定
            ObjVdo.pp32_x = 0.0#
            ObjVdo.pp32_y = 0.0#
            ObjVdo.pp34_x = stPLT.BPOX                                  ' Bp Offset X
            ObjVdo.pp34_y = stPLT.BPOY                                  ' Bp Offset Y
            ObjVdo.pfTrim_x = gSysPrm.stDEV.gfTrimX                     ' 画面表示位置x
            ObjVdo.pfTrim_y = gSysPrm.stDEV.gfTrimY                     ' 画面表示位置y
            ObjVdo.pfBlock_x = stPLT.zsx                                ' Block Size x(mm)
            ObjVdo.pfBlock_y = stPLT.zsy                                ' Block Size y(mm)
            ObjVdo.zwaitpos = stPLT.Z_ZOFF                              ' Z PROBE OFF OFFSET(mm)

            ' ﾃﾝﾌﾟﾚｰﾄｸﾞﾙｰﾌﾟ番号設定(毎回やると遅くなる)
            ' 毎回やらないと「パターンが見つかりません」のエラーになる2013/03/10            
            'If (giTemplateGroup <> stPTN(i).intGRP) Then
            giTemplateGroup = stPTN(i).intGRP                       ' ﾃﾝﾌﾟﾚｰﾄｸﾞﾙｰﾌﾟ番号設定
            ObjVdo.PATTERNGROUP = giTemplateGroup
            ObjVdo.SelectTemplateGroup(giTemplateGroup)
            'End If

            '---------------------------------------------------------------------------
            '   パターンマッチィング処理
            '---------------------------------------------------------------------------
            ObjVdo.VideoStart()                                         ' ビデオライブラリースタート
            nTemplateNum = stPTN(i).intPTN                              ' ﾃﾝﾌﾟﾚｰﾄ番号(1～50)
            Thresh = DllSysPrmSysParam_definst.GetPtnMatchThresh(giTemplateGroup, nTemplateNum) ' 閾値設定

            ' 目印位置XYへBP最高速移動(絶対値)
            r = ObjSys.EX_MOVE(gSysPrm, stPTN(i).dblPosX, stPTN(i).dblPosY, 1)
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
                Return (r)
            End If

            ' @@@888 Call ZWAIT(100)                                             ' WAIT(Sec)

            ' パターンマッチィングを行う
            Call ObjVdo.PatternDisp(False)                              ' パターン枠非表示
            'r = ObjVdo.PatternMatching(nTemplateNum, crx, cry, corrval)
            r = ObjVdo.PatternMatching_EX(nTemplateNum, 0, True, crx, cry, corrval)

            ' パターンマッチィング正常時
            If (r = 0) And (corrval > Thresh) Then                      ' パターンマッチィング ?
                Call ObjVdo.PatternDisp(True)
                'crx = crx / 1000.0#                                     ' crx = ずれ量x
                ''crx = -crx / 1000#                                     ' crx = ずれ量x(※VIDEOで左上ｺｰﾅｰ対応なしなのか?反転しないとダメ)
                'cry = -cry / 1000.0#                                    ' cry = ずれ量y

                ' ズレ量表示
                'strMSG = "カット位置補正値 : "
                strMSG = "カット位置補正値" & i.ToString("00") & " : "
                strMSG = strMSG & "dX=" & crx.ToString("#0.0000") & "mm" & ", dY=" & cry.ToString("#0.0000") & "mm"
                strMSG = strMSG & " 閾値=" & Thresh.ToString("0.00") & ", 一致度=" & corrval.ToString("##0.00% ") & vbCrLf
                Call Z_PRINT(strMSG)
                gcPtnCorrval(i) = corrval.ToString("0.00")
                ' ズレ量保存
                stPTN(i).dblDRX = crx                                   ' ズレ量X
                stPTN(i).dblDRY = cry                                   ' ズレ量Y

                ' パターンマッチィングエラー時
            Else
                'For ib = 1 To 20                                        ' ビープ音
                '    Call Beep()
                'Next ib
                strMSG = "カット位置補正 : パターンが見つかりません。" & vbCrLf & "                 (ｸﾞﾙｰﾌﾟ番号="
                strMSG = strMSG & stPTN(i).intGRP.ToString & ", ﾃﾝﾌﾟﾚｰﾄ番号=" & nTemplateNum.ToString & " 閾値=" & corrval.ToString("0.000") & ")" & vbCrLf
                Call Z_PRINT(strMSG)
                rtn = cFRS_ERR_PTN                                      ' RETURN値 = パターンマッチィングエラー
                Call BSIZE(stPLT.zsx, stPLT.zsy)                        ' エラーの時に０に変えられる為再度設定する。
                Call ObjSys.EX_BPOFF(gSysPrm, stPLT.BPOX, stPLT.BPOY)   ' エラーの時に０に変えられる為再度設定する。
            End If

            '---------------------------------------------------------------------------
            '   終了処理
            '---------------------------------------------------------------------------
            ObjVdo.PatternDisp((False))                                 ' パターン表示
            If (gSysPrm.stDEV.giCutPic = 0) Then                        ' VGAボードあり?
                ObjVdo.VideoStop()                                      ' ビデオライブラリーストップ
            End If
            ObjCrossLine.CrossLineOff()                             ' クロスライン非表示　' 'V2.2.1.3③
            ObjMain.Refresh()                                           ' メイン画面 Refresh
            Return (rtn)                                                ' Return値設定

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.CutPosCorrect() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Return値 = トラップエラー
        End Try
    End Function
#End Region

#Region "パターン認識実行(カット位置補正)"
    '''=========================================================================
    ''' <summary>パターン認識実行</summary>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          上記以外      = エラー
    ''' </returns>
    '''=========================================================================
    Public Function Ptn_Match_Exe() As Short

        Dim i As Short                                                  ' Index
        Dim r As Short                                                  ' 関数戻値
        Dim rtn As Short = cFRS_NORMAL                                  ' 関数戻値
        Dim Flg As Short
        Dim strMSG As String = ""                                       ' メッセージ編集域
        Dim bPatternMatch As Boolean = True                             ' V1.0.4.3⑥

        Try

            ' 初期処理
            Flg = 0
            If UserSub.IsSpecialTrimType And giAppMode = APP_MODE_TRIM Then
                If glCutPosTimes = 0 Then                               ' カット位置補正の実施頻度を指定 ０：補正しない １：毎回補正 ２以上：指定回数おきに補正を実施
                    If gisCutPosExecuteAutoNG Then                      ' V1.0.4.3⑥　自動ＮＧ判定ありの時は、補正を実施する。
                        bPatternMatch = False                           ' V1.0.4.3⑥　通常のカット位置補正は実施しない。
                    Else
                        Exit Function
                    End If
                ElseIf glCutPosCounter < glCutPosTimes Then
                    If glCutPosCounter = 1 Then
                        For i = 1 To stPLT.PtnCount                                 ' パターン登録数分繰返す
                            gcPtnCorrval(i) = "SAME"
                        Next
                    End If
                    glCutPosCounter = glCutPosCounter + 1           ' 補正回数カウンタ１カウントアップ
                    If gisCutPosExecuteAutoNG Then                      ' V1.0.4.3⑥　自動ＮＧ判定ありの時は、補正を実施する。
                        bPatternMatch = False                           ' V1.0.4.3⑥　通常のカット位置補正は実施しない。
                    Else
                        Exit Function
                    End If
                End If
            End If

            For i = 1 To MAXRGN                                         ' パターン登録数分繰返す
                gTblPtn(i) = 0                                          ' パターン認識結果 = 0(OK)
                If giAppMode <> APP_MODE_TRIM Then
                    stPTN(i).dblDRX = 0.0                               ' ズレ量X
                    stPTN(i).dblDRY = 0.0                               ' ズレ量Y
                End If
            Next i

            ' パターン認識処理
            Call ObjSys.Ilum_Ctrl(gSysPrm, Z1, ZOPT)                    ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ点灯(制御有時有効)
            Form1.Refresh()                                             ' パターン認識時画面表示のため
            '@@@888 TEST Call ObjSys.WAIT(0.9)                                       ' 同上
            For i = 1 To stPLT.PtnCount                                 ' パターン登録数分繰返す
                ' V1.0.4.3⑥                If (stPTN(i).PtnFlg <> CUT_PATTERN_NONE) Then           ' パターン認識有(自動)又は手動の時？ ###777
                If ((bPatternMatch = True And stPTN(i).PtnFlg <> CUT_PATTERN_NONE) Or stPTN(i).PtnFlg = CUT_PATTERN_AUTO_NG) Then           ' パターン認識有(自動)又は手動の時？ ###777

                    'V2.2.1.1⑥↓
                    'カット位置補正を実行する必要があるかチェック 
                    If CompareCutCorrData(i) = False Then
                        strMSG = "既に実行済の為、前回の値を使用します。: パターン番号=" & stPTN(i).intGRP.ToString & ", テンプレート番号=" & stPTN(i).intPTN.ToString & vbCrLf
                        Call Z_PRINT(strMSG)
                        Continue For
                    End If
                    'V2.2.1.1⑥↑

                    r = CutPosCorrect(i)                                ' ﾊﾟﾀｰﾝ認識(ｶｯﾄ位置補正値設定)
                    If (r = cFRS_NORMAL) Then                           ' ﾊﾟﾀｰﾝ認識正常 ?

                    ElseIf (r = cFRS_ERR_RST) Then                      ' RESET SW押下 ?
                        Return (cFRS_ERR_RST)                           ' Return値 = RESET SW押下

                        ' 手動ﾃｨｰﾁﾝｸﾞを実行する(ﾊﾟﾀｰﾝ認識ｴﾗｰで手動ﾃｨｰﾁﾝｸﾞ指定時)
                    ElseIf (r = cFRS_ERR_PTN) Then                      ' ﾊﾟﾀｰﾝ認識ｴﾗｰ ?
                        If (stPTN(i).PtnFlg = CUT_PATTERN_MANUAL) Then  ' 手動ﾃｨｰﾁﾝｸﾞ指定 ?
                            r = ObjMTC.SetCrossLineObject(gparModules)
                            If r <> cFRS_NORMAL Then
                                MsgBox("User.Ptn_Match_Exe() SetCrossLineObject ERROR")
                            End If
                            'ObjCrossLine.CrossLineDispXY(0, 0)                ' クロスライン表示
                            'V2.2.0.0① End_GazouProc(ObjGazou)                                 ' 画像表示プログラムを終了する
                            r = Manual_Teach(i)                         ' 手動ティーチング処理
                            ObjCrossLine.CrossLineOff()                             ' クロスライン非表示
                            If (r <> cFRS_NORMAL) Then                  ' 手動ティーチングキャンセル ?
                                gcPtnCorrval(i) = "NG"
                                If UserSub.IsSpecialTrimType And rtn <> cFRS_ERR_RST Then       ' パターンマッチエラーとしないで継続
                                    rtn = cFRS_NORMAL                                           ' Return値 = 正常
                                Else
                                    Return (r)                              ' Return値設定
                                End If
                            End If
                            gcPtnCorrval(i) = "MANUAL"
                        Else
                            Flg = 1
                            gTblPtn(i) = 1                              ' パターン認識結果 = 1(NG)
                            gcPtnCorrval(i) = "NG"
                            ' V1.0.4.3⑥ 自動ＮＧ判定ＮＧ↓
                            If stPTN(i).PtnFlg = CUT_PATTERN_AUTO_NG Then
                                stREG(i).bPattern = False
                            End If
                            ' V1.0.4.3⑥ 自動ＮＧ判定ＮＧ↑
                        End If

                        ' 非常停止等エラー
                    Else
                        If UserSub.IsSpecialTrimType And rtn <> cFRS_ERR_RST Then       ' パターンマッチエラーとしない。
                            gcPtnCorrval(i) = "NG"
                            rtn = cFRS_NORMAL                           ' パターンマッチエラーとしないで継続
                        Else
                            Return (r)                                  ' Return値設定
                        End If
                    End If
                ElseIf (stPTN(i).PtnFlg = CUT_PATTERN_NONE) Then
                    gcPtnCorrval(i) = "NONE"
                    stPTN(i).dblDRX = 0.0                               ' 無しの時は、補正量０とする。
                    stPTN(i).dblDRY = 0.0
                End If
            Next i

        ' ブロック内でパターン認識有りと無しが混在していた場合は、有りの補正情報を無しへ展開する。
        If bPatternMatch And UserSub.IsSpecialTrimType Then         'V1.0.4.3⑥ bPatternMatch追加
            Dim sSavePos As Short = 0
            Dim bNotPtn As Boolean = False
            For i = 1 To stPLT.RCount                               ' パターン登録数分繰返す
                If stPTN(i).PtnFlg = CUT_PATTERN_NONE Then
                    bNotPtn = True                                  ' パターン認識無し
                ElseIf stPTN(i).PtnFlg <> CUT_PATTERN_AUTO_NG Then
                    sSavePos = i                                    ' パターン認識あり
                End If
            Next i
            If sSavePos <> 0 And bNotPtn Then
                For i = 1 To stPLT.RCount                           ' パターン登録数分繰返す
                    If stPTN(i).PtnFlg = CUT_PATTERN_NONE Then
                        stPTN(i).dblDRX = stPTN(sSavePos).dblDRX
                        stPTN(i).dblDRY = stPTN(sSavePos).dblDRY
                        gcPtnCorrval(i) = "SAME"
                    End If
                Next i
            End If
        End If

        ' 後処理
        If (bPatternMatch = True And Flg = 0) Then
            glCutPosCounter = 1                                     ' 成功回数１回目
            rtn = cFRS_NORMAL                                       ' Return値 = 正常
        Else
            If UserSub.IsSpecialTrimType Then
                rtn = cFRS_NORMAL                                       ' パターンマッチエラーとしない。
            Else
                rtn = cFRS_ERR_PTN                                      ' Return値 = パターン認識エラー
            End If
        End If

STP_END:
        Call ObjSys.Ilum_Ctrl(gSysPrm, Z0, ZOPT)                    ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ消灯(制御有時有効)
        Return (rtn)                                                ' Return値設定

        ' トラップエラー発生時 
        Catch ex As Exception
        strMSG = "User.Ptn_Match_Exe() TRAP ERROR = " + ex.Message
        MsgBox(strMSG)
        Return (cFRS_ERR_PTN)                                       ' Return値 = パターン認識エラー
        End Try
    End Function
#End Region
#Region "カット位置補正値保存（カウンター使用時）"
    '''=========================================================================
    ''' <summary>
    ''' カット位置補正頻度の設定　
    ''' </summary>
    ''' <param name="lTimes">０：補正しない １：毎回補正 ２以上：指定回数おきに補正を実施</param>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub SetCutPosTimes(ByVal lTimes As Long)
        glCutPosTimes = lTimes
        System.Diagnostics.Debug.WriteLine("カット位置補正頻度 = " & Val(glCutPosTimes))
    End Sub
#End Region
#Region "カット位置補正情報の取得（カウンター使用時）"
    '''=========================================================================
    ''' <summary>
    ''' カット位置補正情報の取得
    ''' </summary>
    ''' <param name="dPosX"></param>
    ''' <param name="dPosY"></param>
    ''' <returns>補正直後の場合 TRUE</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function GetCutPosData(ByVal rn As Short, ByRef dPosX As Double, ByRef dPosY As Double) As Boolean

        dPosX = stPTN(rn).dblDRX               ' 前回カット位置補正量Ｘ
        dPosY = stPTN(rn).dblDRY               ' 前回カット位置補正量Ｙ

        If glCutPosCounter = 1 Then
            GetCutPosData = True
        Else
            GetCutPosData = False
        End If

    End Function
#End Region

#Region "手動ティーチング処理"
    '''=========================================================================
    ''' <summary>手動ティーチング処理</summary>
    ''' <param name="i"> (INP) パターン登録データの添字(1 ORG)</param>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          cFRS_ERR_RST  = キャンセル選択
    ''' </returns>
    '''=========================================================================
    Public Function Manual_Teach(ByVal i As Integer) As Integer

        'Dim frmMT As System.Windows.Forms.Form          ' Form
        Dim r As Short                                                  ' 戻値
        Dim rtn As Integer = cFRS_NORMAL                                ' 戻値
        Dim strMSG As String                                            ' ﾒｯｾｰｼﾞ編集域

        Try
            If (giAppMode = APP_MODE_TRIM) Then ' トリミング中
                'V2.2.0.0① End_GazouProc(ObjGazou)                                 ' 画像表示プログラムを終了する
            End If
            ' ボタン等を非表示にする
            ObjMain.cmdHelp.Visible = False                               ' Versionボタン非表示 
            ObjMain.Grpcmds.Visible = False                               ' コマンドボタングループボックス非表示
            ObjMain.GrpMode.Visible = False                               ' ディジタルSWグループボックス非表示
            ObjMain.frmInfo.Visible = False                               ' 結果表示域非表示

            strMSG = "手動でパターン位置を合わせてください。(パターン番号 = " & i.ToString("##0") & ")"
            Call Z_PRINT(strMSG & vbCrLf)

            ' 手動ティーチング前処理
            Call ObjMTC.Setup(gSysPrm, stPLT.zsx, stPLT.zsy, stPLT.BPOX, stPLT.BPOY, Form1.Text2.Left, Form1.Text2.Top)

            ' 手動ティーチング処理
            r = ObjMTC.ManualTeach_Renamed(stPTN(i).dblPosX, stPTN(i).dblPosY, stPTN(i).dblDRX, stPTN(i).dblDRY)
            If (r <> cFRS_ERR_START) Then                               ' OK 以外 ?
                If (r = cFRS_ERR_RST) Then
                    rtn = cFRS_ERR_RST                                  ' Return値 = キャンセル選択
                Else
                    rtn = r
                End If
            End If

            ' ズレ量表示
            If (r = 1) Then                                 ' キャンセルでない ?
                strMSG = "カット位置補正値" & i.ToString("00") & " : "
                strMSG = strMSG & "dX=" & stPTN(i).dblDRX.ToString("#0.0000").PadLeft(7) & "mm" & ", dY=" & stPTN(i).dblDRY.ToString("#0.0000").PadLeft(7) & "mm" & vbCrLf
                Call Z_PRINT(strMSG)
            End If

            ' ボタン等を表示する
            ObjMain.cmdHelp.Visible = True                                ' Versionボタン表示 
            ObjMain.Grpcmds.Visible = True                                ' コマンドボタングループボックス表示
            ObjMain.GrpMode.Visible = True                                ' ディジタルSWグループボックス表示
            ObjMain.frmInfo.Visible = True                                ' 結果表示域表示

            If (giAppMode = APP_MODE_TRIM) Then ' トリミング中
                'V2.2.0.0① Exec_GazouProc(ObjGazou, DISPGAZOU_PATH, DISPGAZOU_WRK, INTERNAL_CAMERA)  ' 画像表示プログラムを起動する
            End If

            Return (rtn)                                                ' Return値設定

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.Manual_Teach() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Return値 = トラップエラー
        End Try
    End Function
#End Region

#Region "θ補正の為のパターン登録(RECOG)処理"
    '''=========================================================================
    ''' <summary>θ補正の為のパターン登録</summary>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          上記以外      = エラー
    ''' </returns>
    '''=========================================================================
    Public Function User_PatternTeach() As Integer

        Dim s As String
        Dim r As Integer
        Dim strMSG As String                                        ' メッセージ編集域

        Try
            '---------------------------------------------------------------------------
            '   初期設定処理
            '---------------------------------------------------------------------------
            gbInitialized = False
            ChDir(My.Application.Info.DirectoryPath)
            Call BSIZE(stPLT.zsx, stPLT.zsy)                        ' ブロックサイズ設定
            Call InitThetaCorrection()                              ' パターン登録(RECOG)コントロール初期値設定

            ' フォーム表示位置
            Form1.VideoLibrary1.frmLeft = Form1.Text2.Left          ' フォーム表示位置
            Form1.VideoLibrary1.frmTop = Form1.Text2.Top            '

            '---------------------------------------------------------------------------
            '   θ補正の為のパターン登録(XYﾃｰﾌﾞﾙﾄﾘﾑ位置移動も行う)
            '---------------------------------------------------------------------------
            r = Form1.VideoLibrary1.PatternRegist(giAppMode)

            ' データの更新を行う
            If (r = cFRS_NORMAL Or r = cFRS_ERR_START) Then                               ' 正常終了 ?
                ' トリミングデータ(プレートデータ)を更新する
                stThta.fpp34_x = Form1.VideoLibrary1.pp34_x         ' 補正ポジションオフセットx更新
                stThta.fpp34_y = Form1.VideoLibrary1.pp34_y         ' 補正ポジションオフセットy更新

                'V2.2.0.028 ↓
                If giTablePosUpd = 1 Then
                    stThta.fpp32_x = Form1.VideoLibrary1.pp32_x        ' 補正ポジションオフセットx更新
                    stThta.fpp32_y = Form1.VideoLibrary1.pp32_y        ' 補正ポジションオフセットy更新
                    stThta.fpp33_x = Form1.VideoLibrary1.PP33X         ' 補正ポジションオフセットx更新
                    stThta.fpp33_y = Form1.VideoLibrary1.pp32_y         ' 補正ポジションオフセットy更新
                End If
                'V2.2.0.028 ↑

                ' Video.OCXエラー ?
            ElseIf (r <= cFRS_VIDEO_INI) And (r >= cFRS_MVC_10) Then
                Select Case r
                    Case cFRS_VIDEO_INI
                        s = "VIDEOLIB: Not initialized."            ' "初期化が行われていません。"
                    Case cFRS_VIDEO_FRM
                        s = "VIDEOLIB: Form Display Now"            ' "フォームの表示中です。"
                    Case cFRS_VIDEO_PRP
                        s = "VIDEOLIB: Invalid property value."     ' "プロパティ値が不正です"
                    Case cFRS_VIDEO_UXP
                        s = "VIDEOLIB: Unexpected error"            ' "予期せぬエラー"
                    Case Else
                        s = "VIDEOLIB: Unexpected error 2"
                End Select
                Call ObjSys.TrmMsgBox(gSysPrm, s, MsgBoxStyle.OkOnly, My.Application.Info.Title)

                ' その他のエラー 
            Else
                Return (r)                                          ' Return値設定
            End If

            Return (cFRS_NORMAL)                                    ' Return値 = 正常

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.User_PatternTeach() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                      ' Return値 = トラップエラー
        End Try
    End Function
#End Region

#Region "パターン登録初期設定処理"
    '''=========================================================================
    ''' <summary>パターン登録初期設定処理</summary>
    ''' <remarks>※未テストのため参考程度とする</remarks>
    '''=========================================================================
    Function InitThetaCorrection() As Integer

        Dim strMSG As String                                ' メッセージ編集域

        Try
            ObjVdo.pfTrim_x = stPLT.z_xoff                  ' Parts Handler Load Pos. X(mm)
            ObjVdo.pfTrim_y = stPLT.z_yoff                  ' Parts Handler Load Pos. Y(mm)
            ObjVdo.pfBlock_x = stPLT.zsx                    ' Block Size X
            ObjVdo.pfBlock_y = stPLT.zsy                    ' Block Size Y
            ObjVdo.pfBpOff_x = stPLT.BPOX                   ' BP OFF X
            ObjVdo.pfBpOff_y = stPLT.BPOY                   ' BP OFF Y
            ObjVdo.zwaitpos = stPLT.Z_ZOFF                  ' Z PROBE OFF OFFSET

            ObjVdo.ThetaRCenterX = gSysPrm.stDEV.gfRot_X1    ' 回転中心座標 X mm
            ObjVdo.ThetaRCenterY = gSysPrm.stDEV.gfRot_Y1    ' 回転中心座標 Y mm
            ObjVdo.PP18 = 0                                 ' Z待機位置
            ObjVdo.PP30 = stThta.iPP30                      ' 補正モード(0:自動,1:手動, 2:自動+微調)
            ObjVdo.PP31 = stThta.iPP31                      ' 補正方法(0:なし,1:1回のみ,2:毎回）※PP30=0のときは無効
            ObjVdo.PP53 = stThta.fTheta                     ' θ軸角度

            ' 手動補正モードで補正あり ?
            If (stThta.iPP30 = 1) Then
                ObjVdo.PP31 = 2                             ' 手動補正時の動作 = 毎回
            End If
            ObjVdo.pp32_x = stThta.fpp32_x                  ' パターン1座標x
            ObjVdo.pp32_y = stThta.fpp32_y                  ' パターン1座標y
            ObjVdo.PP33X = stThta.fpp33_x                   ' パターン2座標x
            ObjVdo.PP33Y = stThta.fpp33_y                   ' パターン2座標y
            If (gSysPrm.stDEV.giEXCAM = 1) Then              ' 外部ｶﾒﾗ ?
                ObjVdo.pp34_x = 0.0#                        ' 手動で補正あり時 pp34_x,y分
                ObjVdo.pp34_y = 0.0#                        ' ずれるので0にする
                ObjVdo.pp36_x = gSysPrm.stGRV.gfEXCAM_PixelX ' Xピクセル分解能 um
                ObjVdo.pp36_y = gSysPrm.stGRV.gfEXCAM_PixelY ' Yピクセル分解能 um

            Else
                ObjVdo.pp34_x = stThta.fpp34_x              ' 補正ポジションオフセットx
                ObjVdo.pp34_y = stThta.fpp34_y              ' 補正ポジションオフセットy
                ObjVdo.pp36_x = gSysPrm.stGRV.gfPixelX       ' Xピクセル分解能 um
                ObjVdo.pp36_y = gSysPrm.stGRV.gfPixelY       ' Yピクセル分解能 um
            End If
            'ObjVdo.PP35 = 1                                ' Debug用(1:on,0:off)
            ObjVdo.PP35 = 0                                 ' Debug用(1:on,0:off)
            ObjVdo.PP37_1 = stThta.iPP37_1                  ' パターン1 テンプレート番号
            ObjVdo.PP37_2 = stThta.iPP37_2                  ' パターン2 テンプレート番号
            ObjVdo.PP52 = stThta.iPP38                      ' POS1 パターン グループ番号
            ObjVdo.PP52_1 = stThta.iPP38                    ' POS2 パターン グループ番号
            ObjVdo.RNASTMPNUM = False
            ObjVdo.frmLeft = ObjMain.Text2.Left             ' 表示位置(Left)
            ObjVdo.frmTop = ObjMain.Text2.Top               ' 表示位置(Top)
            Return (cFRS_NORMAL)                            ' Return値 = 正常

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.InitThetaCorrection() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                              ' Return値 = トラップエラー
        End Try
    End Function
#End Region

#Region "θ補正&XYﾃｰﾌﾞﾙﾄﾘﾑ位置移動処理"
    '''=========================================================================
    ''' <summary>θ補正とXYﾃｰﾌﾞﾙﾄﾘﾑ位置移動処理</summary>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          cFRS_ERR_RST   = Cancel(RESETｷｰ)
    '''          cFRS_ERR_PTN  = パターン認識エラー
    '''          上記以外      = その他エラー
    ''' </returns>
    '''=========================================================================
    Public Function Move_Trimposition() As Short

        Dim rtn As Integer = cFRS_NORMAL                                ' Return値
        Dim Thresh1 As Double
        Dim Thresh2 As Double
        'Dim fExCmX As Double                                           ' 外部ｶﾒﾗｵﾌｾｯﾄX(mm)
        'Dim fExCmY As Double                                           ' 外部ｶﾒﾗｵﾌｾｯﾄY(mm)
        Dim strMSG As String                                            ' Display Message
        Dim r As Short

        Try
            ' 初期処理
            dblCorrectX = 0                                             ' θ補正時のXYﾃｰﾌﾞﾙずれ量X,Y(mm)
            dblCorrectY = 0

#If cOFFLINEcDEBUG Then
            Call Z_PRINT("Move_Trimposition() NOP" & vbCrLf)
            Return (cFRS_NORMAL)
#End If
            r = UserBas.Prob_Off()                                          ' プローブ待機位置移動
            If (r <> cFRS_NORMAL) Then                                  ' エラーならエラーリターン(メッセージ表示済み)
                Return (r)
            End If

            ' θ補正処理
            If (gSysPrm.stDEV.giTheta = 0) Then GoTo STP_START '        ' θ無しなら補正しないでXYﾃｰﾌﾞﾙをﾄﾘﾑ位置に移動
            If (stThta.iPP30 = 1) And (stThta.iPP31 = 0) Then           ' 手動補正モードで補正なし ?
                Call ROUND4(stThta.fTheta)                              ' θ回転
                GoTo STP_START
            End If

            Call ObjSys.Ilum_Ctrl(gSysPrm, Z1, ZOPT)                    ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ点灯(制御有時有効)
            r = ThetaCorrection(dblCorrectX, dblCorrectY)               ' θ補正実行
            If r = cFRS_ERR_RST Then                                    ' ###1033 キャンセル
                Return (r)                                              ' ###1033 キャンセル
            End If                                                     ' ###1033 キャンセル
            If (r <> cFRS_NORMAL) Then                                  ' ERROR ?
                ' パターン認識エラー ?
                If (r >= cFRS_MVC_10) And (r <= cFRS_VIDEO_PTN) Then
                    Call Beep()                                         ' Beep音
                    If (gSysPrm.stTMN.giMsgTyp = 0) Then
                        strMSG = "パターン認識エラー(" + r.ToString("0") + ")"
                    Else
                        strMSG = "VIDEOLIB: Pattern Matching Error"
                    End If
                    Call Z_PRINT(strMSG & vbCrLf)
                    rtn = cFRS_ERR_PTN                                  ' パターン認識エラー
                Else
                    rtn = r
                End If
                Call ObjSys.Ilum_Ctrl(gSysPrm, Z0, ZOPT)                ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ消灯(制御有時有効)
                Return (rtn)
            End If

            ' 閾値取得
            Thresh1 = DllSysPrmSysParam_definst.GetPtnMatchThresh(stThta.iPP38, stThta.iPP37_1)
            Thresh2 = DllSysPrmSysParam_definst.GetPtnMatchThresh(stThta.iPP38, stThta.iPP37_2)

            ' θ補正結果取得
            Call ObjVdo.GetThetaResult(stResult)
            If (gSysPrm.stDEV.giTheta = 0) Then
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    strMSG = "トリム位置X,Y=" & stResult.fPosx.ToString("0.0000").PadLeft(9) & ","
                    strMSG = strMSG & stResult.fPosy.ToString("0.0000").PadLeft(9) & "  "
                    strMSG = strMSG & " ずれ量X,Y    =" & stResult.fCorx.ToString("0.0000").PadLeft(9) & ","
                    strMSG = strMSG & stResult.fCory.ToString("0.0000").PadLeft(9) & vbCrLf
                    strMSG = strMSG & "補正位置1X,Y =" & stResult.fPos1x.ToString("0.0000").PadLeft(9) & ","
                    strMSG = strMSG & stResult.fPos1y.ToString("0.0000").PadLeft(9) & "  "
                    strMSG = strMSG & " ずれ量1X,Y   =" & stResult.fCor1x.ToString("0.0000").PadLeft(9) & ","
                    strMSG = strMSG & stResult.fCor1y.ToString("0.0000").PadLeft(9) & vbCrLf
                    If (stThta.iPP30 = 0) Then                              ' 自動補正モード ?
                        strMSG = strMSG & "  一致度POS1   =" & stResult.fCorV1.ToString("0.0000").PadLeft(9) & vbCrLf
                    End If
                Else
                    ' θ補正表示情報設定(英語)
                    strMSG = "  Trim PositionXY=" & stResult.fPosx.ToString("0.0000").PadLeft(9) & ","
                    strMSG = strMSG & stResult.fPosy.ToString("0.0000").PadLeft(9) & "  "
                    strMSG = strMSG & " Distance=" & stResult.fCorx.ToString("0.0000").PadLeft(9) & ","
                    strMSG = strMSG & stResult.fCory.ToString("0.0000").PadLeft(9) & vbCrLf
                    strMSG = strMSG & "  Correct position1=" & stResult.fPos1x.ToString("0.0000").PadLeft(9) & ","
                    strMSG = strMSG & stResult.fPos1y.ToString("0.0000").PadLeft(9) & "  "
                    strMSG = strMSG & " Distance1=" & stResult.fCor1x.ToString("0.0000").PadLeft(9) & ","
                    strMSG = strMSG & stResult.fCor1y.ToString("0.0000").PadLeft(9) & vbCrLf
                    If (stThta.iPP30 = 0) Then                              ' 自動補正モード ?
                        strMSG = strMSG & "  Correlation coefficient1=" & stResult.fCorV1.ToString("0.0000").PadLeft(9) & vbCrLf
                    End If
                End If
            Else
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    ' θ補正表示情報設定(日本語)
                    strMSG = "θ角度 = " & stResult.fTheta.ToString("0.0000").PadLeft(7) & "°" & vbCrLf
                    If (stThta.iPP30 = 2) Then                              ' 自動+微調の場合
                        strMSG = "θ角度 = " & stResult.fTheta.ToString("0.0000").PadLeft(7) & "°"
                        strMSG = strMSG & "+ " & stThta.fTheta.ToString("0.0000").PadLeft(7) & "°" & vbCrLf
                    End If
                    strMSG = strMSG & "トリム位置X,Y=" & stResult.fPosx.ToString("0.0000").PadLeft(9) & ","
                    strMSG = strMSG & stResult.fPosy.ToString("0.0000").PadLeft(9) & "  "
                    strMSG = strMSG & " ずれ量X,Y    =" & stResult.fCorx.ToString("0.0000").PadLeft(9) & ","
                    strMSG = strMSG & stResult.fCory.ToString("0.0000").PadLeft(9) & vbCrLf
                    strMSG = strMSG & "補正位置1X,Y =" & stResult.fPos1x.ToString("0.0000").PadLeft(9) & ","
                    strMSG = strMSG & stResult.fPos1y.ToString("0.0000").PadLeft(9) & "  "
                    strMSG = strMSG & " ずれ量1X,Y   =" & stResult.fCor1x.ToString("0.0000").PadLeft(9) & ","
                    strMSG = strMSG & stResult.fCor1y.ToString("0.0000").PadLeft(9) & vbCrLf
                    strMSG = strMSG & "補正位置2X,Y =" & stResult.fPos2x.ToString("0.0000").PadLeft(9) & ","
                    strMSG = strMSG & stResult.fPos2y.ToString("0.0000").PadLeft(9) & "  "
                    strMSG = strMSG & " ずれ量2X,Y   =" & stResult.fCor2x.ToString("0.0000").PadLeft(9) & ","
                    strMSG = strMSG & stResult.fCor2y.ToString("0.0000").PadLeft(9) & vbCrLf
                    If (stThta.iPP30 = 0) Then                              ' 自動補正モード ?
                        strMSG = strMSG & "  一致度POS1   =" & stResult.fCorV1.ToString("0.0000").PadLeft(9) & ","
                        strMSG = strMSG & " 一致度POS2   =" & stResult.fCorV2.ToString("0.0000").PadLeft(9) & vbCrLf
                    End If
                Else
                    ' θ補正表示情報設定(英語)
                    strMSG = "Theta = " & stResult.fTheta.ToString("0.0000").PadLeft(7) & "degree" & vbCrLf
                    If (stThta.iPP30 = 2) Then                              ' 自動+微調の場合
                        strMSG = "Theta= " & stResult.fTheta.ToString("0.0000").PadLeft(7) & "degree"
                        strMSG = strMSG & " + " & stThta.fTheta.ToString("0.0000").PadLeft(7) & "degree" & vbCrLf
                    End If
                    strMSG = strMSG & "  Trim PositionXY=" & stResult.fPosx.ToString("0.0000").PadLeft(9) & ","
                    strMSG = strMSG & stResult.fPosy.ToString("0.0000").PadLeft(9) & "  "
                    strMSG = strMSG & " Distance=" & stResult.fCorx.ToString("0.0000").PadLeft(9) & ","
                    strMSG = strMSG & stResult.fCory.ToString("0.0000").PadLeft(9) & vbCrLf
                    strMSG = strMSG & "  Correct position1=" & stResult.fPos1x.ToString("0.0000").PadLeft(9) & ","
                    strMSG = strMSG & stResult.fPos1y.ToString("0.0000").PadLeft(9) & "  "
                    strMSG = strMSG & " Distance1=" & stResult.fCor1x.ToString("0.0000").PadLeft(9) & ","
                    strMSG = strMSG & stResult.fCor1y.ToString("0.0000").PadLeft(9) & vbCrLf
                    strMSG = strMSG & "  Correct position2=" & stResult.fPos2x.ToString("0.0000").PadLeft(9) & ","
                    strMSG = strMSG & stResult.fPos2y.ToString("0.0000").PadLeft(9) & "  "
                    strMSG = strMSG & " Distance2=" & stResult.fCor2x.ToString("0.0000").PadLeft(9) & ","
                    strMSG = strMSG & stResult.fCor2y.ToString("0.0000").PadLeft(9) & vbCrLf
                    If (stThta.iPP30 = 0) Then                              ' 自動補正モード ?
                        strMSG = strMSG & "  Correlation coefficient1=" & stResult.fCorV1.ToString("0.0000").PadLeft(9) & ","
                        strMSG = strMSG & " Correlation coefficient2=" & stResult.fCorV2.ToString("0.0000").PadLeft(9) & vbCrLf
                    End If
                End If
            End If

            ' θ補正情報表示
            Call Z_PRINT(strMSG)

            If stThta.iPP30 <> 1 Then   ' ###1033
                ' POS1の閾値のチェックを行う
                If (Thresh1 > stResult.fCorV1) Then
                    Call Beep()                                             ' Beep音
                    If (gSysPrm.stTMN.giMsgTyp = 0) Then
                        strMSG = "パターン認識エラー (POS1閾値) [" & Thresh1.ToString("0.000") & "] > [" & stResult.fCorV1.ToString("0.00") & "]"
                    Else
                        strMSG = "Pattern Matching Error(POS1 THRESH) [" & Thresh1.ToString("0.000") & "] > [" & stResult.fCorV1.ToString("0.00") & "]"
                    End If
                    Call Z_PRINT(strMSG & vbCrLf)
                    Call ObjSys.Ilum_Ctrl(gSysPrm, Z0, ZOPT)                ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ消灯(制御有時有効)
                    Return (cFRS_ERR_PTN)                                   ' パターン認識エラー
                End If

                ' POS2の閾値のチェックを行う
                If (Thresh2 > stResult.fCorV2) Then
                    Call Beep()                                             ' Beep音
                    If (gSysPrm.stTMN.giMsgTyp = 0) Then
                        strMSG = "パターン認識エラー (POS2閾値)"
                    Else
                        strMSG = "Pattern Matching Error(POS2 THRESH)"
                    End If
                    Call Z_PRINT(strMSG & vbCrLf)
                    Call ObjSys.Ilum_Ctrl(gSysPrm, Z0, ZOPT)                ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ消灯(制御有時有効)
                    Return (cFRS_ERR_PTN)                                   ' パターン認識エラー
                End If

            End If                      ' ###1033

            ' パーツハンドラをトリム位置に移動する
STP_START:
            Call ObjSys.Ilum_Ctrl(gSysPrm, Z0, ZOPT)                    ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ消灯(制御有時有効)

            ' 外部カメラオフセット値を設定する
            'If (gSysPrm.stDEV.giEXCAM_Usr = 0) Then                    ' 外部カメラなし ?
            '    fExCmX = 0.0#                                          ' 外部ｶﾒﾗｵﾌｾｯﾄX = 0
            '    fExCmY = 0.0#                                          ' 外部ｶﾒﾗｵﾌｾｯﾄY = 0
            'Else
            '    fExCmX = SysPrm.stDEV.gfExCmX                          ' 外部ｶﾒﾗｵﾌｾｯﾄX設定
            '    fExCmY = SysPrm.stDEV.gfExCmY                          ' 外部ｶﾒﾗｵﾌｾｯﾄY設定
            'End If
            Call BSIZE(stPLT.zsx, stPLT.zsy)                            ' ブロックサイズ/BPｵﾌｾｯﾄ設定
            Call ObjSys.EX_BPOFF(gSysPrm, stPLT.BPOX, stPLT.BPOY)

            ' ZをZOFF位置へ移動する(EX_STARTのZOFF位置をZ_ZOFFとする)
            'r = PROBOFF_EX(stPLT.Z_ZOFF)                                ' Z OFFE
            'r = ObjMain.System1.EX_ZGETSRVSIGNAL(gSysPrm, CShort(r), 0)
            'If (r <> cFRS_NORMAL) Then                                  ' エラーならエラーリターン(メッセージ表示済み)
            '    Return (r)
            'End If

            ' パーツハンドラをトリム位置に移動(MD=原点位置ﾁｪｯｸ無し)
            r = ObjSys.EX_START(gSysPrm, stPLT.z_xoff + dblCorrectX, stPLT.z_yoff + dblCorrectY, 0)
            'r = System1.EX_START(gSysPrm, stPLT.z_xoff + fExCmX + dblCorrectX, stPLT.z_yoff + fExCmY + dblCorrectY, 0)
            If (r <> cFRS_NORMAL) Then ' ERROR ?
                rtn = r                                                 ' Return値設定
            End If

            If (giAppMode = APP_MODE_PROBE Or giAppMode = APP_MODE_TEACH Or giAppMode = APP_MODE_CUTPOS) And (stPLT.TeachBlockX > 1 Or stPLT.TeachBlockY > 1) Then  ' ###1040①
                r = ObjSys.EX_TSTEP(gSysPrm, stPLT.TeachBlockX, stPLT.TeachBlockY)                                                                                  ' ###1040①
                If (r <> cFRS_NORMAL) Then                                                                                                                          ' ###1040①
                    rtn = r                                                                                                                                         ' ###1040①
                End If                                                                                                                                              ' ###1040①
            End If                                                                                                                                                  ' ###1040①

            'Call frmMain.VideoLibrary1.ChangeCamera(0)                 ' 内部ｶﾒﾗに切替える(外部ｶﾒﾗがあれば)
            Return (rtn)                                                ' Return値設定

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.Move_Trimposition() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Return値 = トラップエラー
        End Try
    End Function
#End Region

#Region "θ補正実行"
    '''=========================================================================
    ''' <summary>θ補正実行</summary>
    ''' <param name="dblCorrectX">(OUT) XYテーブル補正値X</param>
    ''' <param name="dblCorrectY">(OUT) XYテーブル補正値Y</param>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          cFRS_ERR_PTN  = パターン認識エラー
    '''          上記以外      = その他エラー
    ''' </returns>
    '''=========================================================================
    Private Function ThetaCorrection(ByRef dblCorrectX As Double, ByRef dblCorrectY As Double) As Integer

        Dim r As Integer, rLast As Integer                  ' ###1033 ADD rLast
        Dim strMSG As String                                ' メッセージ編集域

        rLast = 0                                           ' ###1033
        Try
            ' 初期処理
            Call InitThetaCorrection()                      ' パターン登録初期値設定

            ' カメラ切替
            If (gSysPrm.stDEV.giCutPic = 0) Then             ' VGAボードあり?
                ObjVdo.VideoStop()                          ' ビデオストップ/スタート(内部/外部ｶﾒﾗ)
                Call ObjVdo.VideoStart2(gSysPrm.stDEV.giEXCAM)
            Else
                If (gSysPrm.stDEV.giEXCAM_Usr = 1) Then    ' 外部カメラを使用？
                    r = ObjVdo.ChangeCamera(EXTERNAL_CAMERA)            ' カメラ切替(外部ｶﾒﾗ)
                    If r <> 0 Then
                        MsgBox("User.ThetaCorrection() 外部カメラ切り替えエラー = [" & r.ToString & "]")
                    End If
                End If
            End If

            ' θ補正処理
            ObjVdo.frmTop = Form1.Text2.Location.Y          ' 補正画面表示位置設定
            ObjVdo.frmLeft = Form1.Text2.Location.X
            r = ObjVdo.CorrectTheta(giAppMode)              ' θ補正

            ' XYテーブル補正値(θ補正時のXYﾃｰﾌﾞﾙずれ量)取得
            If (r = 0) Then
                dblCorrectX = ObjVdo.CorrectTrimPosX
                dblCorrectY = ObjVdo.CorrectTrimPosY
            Else
                dblCorrectX = 0
                dblCorrectY = 0
                rLast = r                                   ' ###1033
            End If

            ' 後処理
            If (gSysPrm.stDEV.giCutPic = 0) Then             ' VGAボードあり?
                ObjVdo.VideoStop()                          ' ビデオライブラリストップ
                ObjMain.Refresh()
            Else
                If (gSysPrm.stDEV.giEXCAM_Usr = 1) Then             ' 外部カメラを使用？
                    r = ObjVdo.ChangeCamera(INTERNAL_CAMERA)       ' カメラ切替(内部ｶﾒﾗ)
                    If r <> 0 Then
                        MsgBox("User.ThetaCorrection() 外部カメラ切り替えエラー = [" & r.ToString & "]")
                    End If
                End If
            End If

            If rLast = 0 Then                               ' ###1033
                rLast = r                                   ' ###1033
            End If                                          ' ###1033
            ' ###1033Return (r)                                      ' Return値設定
            Return (rLast)                                  ' Return値設定

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.ThetaCorrection() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                              ' Return値 = トラップエラー
        End Try
    End Function
#End Region

    '==========================================================================
    '   共通関数
    '==========================================================================
#Region "トリミングデータの変更時の処理"
    ''' <summary>
    ''' トリミングデータの変更時の処理 ###1041①
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub TrimmingDataChange()
        Try
            Dim Rtn As Integer
            Rtn = SETZOFFPOS(stPLT.Z_ZOFF)    ' INTIME内部の待機位置を変更する。
            If Rtn <> cFRS_NORMAL Then
                Call MsgBox_Exception("SETZOFFPOS Ｚ軸待機位置変更が異常終了しました。=[" & Rtn.ToString & "]", Form1)
            Else
                'Rtn = Prob_Off()
                'If Rtn <> cFRS_NORMAL Then
                '    Call MsgBox_Exception("Ｚ軸待機位置移動が異常終了しました。=[" & Rtn.ToString & "]", Form1)
                'End If
            End If

            'V2.0.0.0⑨↓
            With Form1.CCmb_DistributeResList   ' 分布図表示の抵抗番号
                If Form1.CCmb_DistributeResList.Enabled = True Then
                    .Items.Clear()
                    Dim RCount As Integer = UserBas.GetRCountExceptMeasure()
                    Dim iCnt As Integer = 0
                    For i As Integer = 1 To stPLT.RCount Step 1
                        If IsCutResistor(i) Then
                            .Items.Add(stREG(i).strRNO) ' 総抵抗数分繰り返す
                            iCnt = iCnt + 1
                            If iCnt >= RCount Then
                                Exit For
                            End If
                        End If
                    Next i
                    If iCnt > 0 Then
                        stPLT.DistributionResNo = 1
                        .SelectedIndex = stPLT.DistributionResNo - 1
                    End If
                End If
            End With
            'V2.0.0.0⑨↑

            'V2.1.0.0②↓
            If stLASER.iAttNo > 0 Then
                If UserSub.LaserCalibrationAttenuatorDataGet(stLASER.iAttNo, stLASER.dblRotPar, stLASER.iFixAtt, stLASER.dblRotAtt) Then
                    Z_PRINT("アッテネータテーブルからNO=[" & stLASER.iAttNo.ToString & "][" & stLASER.dblRotPar.ToString & "%]の情報を取得しました。")
                    gSysPrm.stRAT.gfAttRate = stLASER.dblRotPar                     ' ###1040⑥ 減衰率(%)
                    gSysPrm.stRAT.giAttRot = stLASER.dblRotAtt                      ' ###1040⑥ ロータリーアッテネータの回転量(0-FFF)
                    gSysPrm.stRAT.giAttFix = stLASER.iFixAtt                        ' ###1040⑥ 固定アッテネータ(0:OFF, 1:ON)
                    Call DllSysPrmSysParam_definst.PutSysPrm_ROT_ATT(gSysPrm.stRAT) ' ###1040⑥
                    Call Form1.SetATTRateToScreen(False)                            'V2.0.0.0⑮
                    'V2.1.0.0⑥カバー開のエラー                    Call Form1.SetATTRateToScreen(True)           'V2.1.0.0⑥ トリミングデータでのＡＴＴ減衰率の設定
                Else
                    Z_PRINT("アッテネータテーブルからの情報取得がエラーになりました。NO=[" & stLASER.iAttNo.ToString & "]")
                    Call MsgBox_Exception("アッテネータテーブルからの情報取得がエラーになりました。NO=[" & stLASER.iAttNo.ToString & "]", Form1)
                End If
            End If
            UserSub.LaserCalibrationSet(POWER_CHECK_LOT)                'V2.1.0.0② レーザパワーモニタリング実行有無設定
            'V2.1.0.0②↑

        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub
#End Region

#Region "ＢＰ移動(カット位置補正あり)"
    '''=========================================================================
    ''' <summary>ＢＰ移動(カット位置補正あり)</summary>
    ''' <param name="i">   (INP) パターン登録データの添字(1 ORG)</param>
    ''' <param name="STX"> (INP) カット位置X</param>
    ''' <param name="STY"> (INP) カット位置Y</param>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          その他        = エラー
    ''' </returns>
    '''=========================================================================
    Public Function STRXY(ByRef i As Short, ByRef STX As Double, ByRef STY As Double) As Short

        Dim r As Short
        Dim POSX As Double                                              ' カット位置X
        Dim POSY As Double                                              ' カット位置Y
        Dim strMSG As String                                            ' メッセージ表示用域

        Try

            If UserSub.IsSpecialTrimType And glCutPosTimes = 0 Then         ' ０：補正しない
                POSX = STX                                                  ' POSX = カット位置X + 補正値X
                POSY = STY                                                  ' POSY = カット位置Y + 補正値Y
            Else
                POSX = STX + stPTN(i).dblDRX                                ' POSX = カット位置X + 補正値X
                POSY = STY + stPTN(i).dblDRY                                ' POSY = カット位置Y + 補正値Y
            End If

            r = ObjSys.EX_MOVE(gSysPrm, POSX, POSY, 1)                  ' BP移動(絶対値)

            If (DGL = TRIM_MODE_CUT) Then                               ' sw = X5(ｶｯﾃｨﾝｸﾞ) ?
                strMSG = "STX=" & POSX.ToString("###0.0###") & ", STY=" & POSY.ToString("###0.0###") & "    DRX=" & stPTN(i).dblDRX.ToString("###0.0###") & ", DRY=" & stPTN(i).dblDRY.ToString("###0.0###") & vbCrLf
                Call Z_PRINT(strMSG)
            End If
            Return (r)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.STRXY() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Return値 = トラップエラー
        End Try
    End Function
#End Region

#Region "ﾎﾟｰｽﾞ付きﾌﾟﾛｰﾌﾞON"
    '''=========================================================================
    ''' <summary>ﾎﾟｰｽﾞ付きﾌﾟﾛｰﾌﾞON</summary>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          上記以外      = エラー
    ''' </returns>
    '''=========================================================================
    Public Function Prob_On() As Short

        Dim r As Short
        Dim strMSG As String                                    ' メッセージ編集域

        Try
            Prob_On = cFRS_NORMAL                               ' Return値 = 正常
            r = EX_ZMOVE(stPLT.Z_ZON)                           ' ZをON位置に移動
            If (r <> cFRS_NORMAL) Then                          ' エラー ?
                Prob_On = r                                     ' Return値設定
                Exit Function
            End If
            'r = EX_ZMOVE2(stPLT.Z2_ZON)                         ' Z2ﾌﾟﾛｰﾌﾞをON位置に移動
            'If (r <> cFRS_NORMAL) Then                          ' エラー ?
            '    Prob_On = r                                     ' Return値設定
            '    Exit Function
            'End If
            Call ZWAIT(PROB_ON_TIM)                             ' WAIT(msec)
            Exit Function

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Prob_On() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Prob_On = cERR_TRAP                                 ' Retuen値 = ﾄﾗｯﾌﾟｴﾗｰ発生
        End Try
    End Function
#End Region

#Region "ﾌﾟﾛｰﾌﾞOFF"
    '''=========================================================================
    ''' <summary>ﾌﾟﾛｰﾌﾞOFF</summary>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          上記以外      = エラー
    ''' </returns>
    '''=========================================================================
    Public Function Prob_Off() As Integer

        Dim r As Integer
        Dim strMSG As String                                            ' メッセージ編集域

        Try
            'r = EX_ZMOVE2(0.0)                                          ' Z2ﾌﾟﾛｰﾌﾞをOFF位置に移動
            'If (r <> cFRS_NORMAL) Then                                  ' エラー ?
            '    Prob_Off = r                                            ' Return値設定
            '    Exit Function
            'End If
            ' ZをZOFF位置へ移動する(EX_STARTのZOFF位置をZ_ZOFFとする)INTIMEにOFF位置を記憶させる為こちらを使用する
            r = PROBOFF_EX(stPLT.Z_ZOFF)                                ' Z OFFE
            r = ObjMain.System1.EX_ZGETSRVSIGNAL(gSysPrm, CShort(r), 0)
            If (r <> cFRS_NORMAL) Then                                  ' エラーならエラーリターン(メッセージ表示済み)
                Return (r)
            End If
            'r = EX_ZMOVE(stPLT.Z_ZOFF)                                  ' ZﾌﾟﾛｰﾌﾞをOFF位置(Z.ZOFF)に移動
            'If (r <> cFRS_NORMAL) Then                                  ' エラー ?
            '    Return (r)                                              ' Return値設定
            'End If
            Return (r)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Prob_Off() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Retuen値 = ﾄﾗｯﾌﾟｴﾗｰ発生
        End Try
    End Function
#End Region

#Region "Ｚ軸待機位置変更とﾌﾟﾛｰﾌﾞOFF"
    '''=========================================================================
    ''' <summary>Ｚ軸待機位置変更とﾌﾟﾛｰﾌﾞOFF ###1041①</summary>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          上記以外      = エラー
    ''' </returns>
    '''=========================================================================
    Public Function SetZOff_Prob_Off() As Short

        Dim iRtn As Integer

        Try
            'V2.0.0.0⑯↓
            iRtn = SETZOFFPOS(-1)                         ' INTIME内部の待機位置を変更する。
            'iRtn = SETZOFFPOS(stPLT.Z_ZOFF)              ' INTIME内部の待機位置を変更する。
            'V2.0.0.0⑯↑
            If iRtn <> cFRS_NORMAL Then
                Call Z_PRINT("SETZOFFPOS Ｚ軸待機位置(" & stPLT.Z_ZOFF.ToString & ")変更が異常終了しました。=[" & iRtn.ToString & "]")
                Return (cERR_TRAP)                                          ' Retuen値 = ﾄﾗｯﾌﾟｴﾗｰ発生
            Else
                iRtn = Prob_Off()
                If iRtn <> cFRS_NORMAL Then
                    Call Z_PRINT("Prob_Off() Ｚ軸待機位置移動が異常終了しました。=[" & iRtn.ToString & "]")
                    Return (cERR_TRAP)                                          ' Retuen値 = ﾄﾗｯﾌﾟｴﾗｰ発生
                End If
            End If
            Return (cFRS_NORMAL)
            ' トラップエラー発生時 
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
            Return (cERR_TRAP)                                          ' Retuen値 = ﾄﾗｯﾌﾟｴﾗｰ発生
        End Try
    End Function
#End Region

#Region "Z移動(絶対)"
    '''=========================================================================
    ''' <summary>Z移動(絶対)</summary>
    ''' <param name="z"></param>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          上記以外      = エラー
    ''' </returns>
    '''=========================================================================
    Public Function EX_ZMOVE(ByVal z As Double) As Integer

        Dim r As Integer
        Dim opt As Integer = 0                                          ' Z軸原点復帰ﾁｪｯｸﾌﾗｸﾞ(0:無, 1:Z軸 2:Z軸)
        Dim MD As Integer = 1                                           ' 絶対移動
        Dim strMSG As String

        Try
            ' ZなしならNOP
            If (gSysPrm.stDEV.giPrbTyp = 0) Then
                Return (cFRS_NORMAL)
            End If

            ' Z移動
            r = ZZMOVE(z, MD)
            r = ObjSys.EX_ZGETSRVSIGNAL(gSysPrm, r, opt)
            Return (r)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "EX_ZMOVE() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Retuen値 = ﾄﾗｯﾌﾟｴﾗｰ発生
        End Try
    End Function
#End Region

#Region "Z2移動(絶対)"
    '''=========================================================================
    ''' <summary>Z2移動(絶対)</summary>
    ''' <param name="z">z(INP) : 移動量</param>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          上記以外      = エラー
    ''' </returns>
    '''=========================================================================
    Public Function EX_ZMOVE2(ByVal z As Double) As Integer

        Dim r As Integer
        Dim opt As Integer = 0                                          ' Z軸原点復帰ﾁｪｯｸﾌﾗｸﾞ(0:無, 1:Z軸 2:Z軸)
        Dim MD As Integer = 1                                           ' 絶対移動
        Dim strMSG As String

        Try
            ' 下方ﾌﾟﾛｰﾌﾞなしならNOP
            If ((gSysPrm.stDEV.giPrbTyp And 2) = 0) Then
                Return (cFRS_NORMAL)
            End If

            ' Z2移動
            r = ZZMOVE2(z, MD)
            r = ObjSys.EX_ZGETSRVSIGNAL(gSysPrm, r, opt)
            Return (r)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "EX_ZMOVE2() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Retuen値 = ﾄﾗｯﾌﾟｴﾗｰ発生
        End Try
    End Function
#End Region

#Region "EXTERNAL バックライトＬＥＤ制御ＯＮ"
    '''=========================================================================
    ''' <summary>
    ''' バックライトＬＥＤ制御ＯＮ
    ''' </summary>
    ''' <remarks>特殊処理の制御、通常の制御ILUM_CTRL_USERは０にする。</remarks>
    '''=========================================================================
    Public Sub BackLight_On()

        Try
            Call EXTOUT1(glLedBit, 0)               ' LED OFF(OnBit, OffBit)
            Call ZWAIT(REL_ON_TIM)                  ' Wait(ms)
            ' トラップエラー発生時 
        Catch ex As Exception
            MsgBox("User.BackLight_On() TRAP ERROR = " + ex.Message)
        End Try
    End Sub
#End Region

#Region "EXTERNAL バックライトＬＥＤ制御ＯＦＦ"
    '''=========================================================================
    ''' <summary>
    ''' バックライトＬＥＤ制御ＯＦＦ
    ''' </summary>
    ''' <remarks>特殊処理の制御、通常の制御ILUM_CTRL_USERは０にする。</remarks>
    '''=========================================================================
    Public Sub BackLight_Off()

        Try
            Call EXTOUT1(0, glLedBit)               ' LED OFF(OnBit, OffBit)
            Call ZWAIT(REL_ON_TIM)                  ' Wait(ms)
            ' トラップエラー発生時 
        Catch ex As Exception
            MsgBox("User.BackLight_Off() TRAP ERROR = " + ex.Message)
        End Try
    End Sub
#End Region

#Region "セーフティチェック(トリミング中)"
    '''=========================================================================
    ''' <summary>セーフティチェック(トリミング中)</summary>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          上記以外      = エラー
    ''' </returns>
    '''=========================================================================
    Public Function SafetyCheck() As Short

        Dim strMSG As String                                            ' メッセージ編集域
        Dim r As Short

        Try
            ' 非常停止等チェック(トリミング中)
            r = ObjSys.Sys_Err_Chk_EX(gSysPrm, APP_MODE_TRIM)
            If (r <> cFRS_NORMAL) Then                                  ' 非常停止等 ?
                Return (r)                                              ' Return値 = 非常停止等
            End If
            Return (cFRS_NORMAL)                                        ' Return値 = 正常

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.SafetyCheck() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Return値 = トラップエラー
        End Try
    End Function
#End Region

#Region "ログファイル名を設定する"
    '''=========================================================================
    ''' <summary>ログファイル名を設定する</summary>
    ''' <param name="FPath">(OUT)ログファイル名</param>
    ''' <remarks>・ログファイル名は起動時、データロード時、ロット切替時に設定される。
    '''          ・拡張子が.LOGの場合シスパラの操作ログの保存日数が過ぎると
    '''            操作ログと共に自動的に削除される。
    '''           </remarks>
    '''=========================================================================
    Public Sub SetLogFileName(ByRef FPath As String)

        Dim strDAT As String
        Dim strMSG As String                                            ' メッセージ編集域

        Try
            '' ログファイル名を設定する ("C:\TRIMDATA\LOG\""LOG_" + ｢ロット番号｣ + ".LOG")
            'FPath = cLOGFILEPATH & "LOG_" & stRLT.strLOT & ".LOG"

            ' ログファイル名を設定する ("C:\TRIMDATA\LOG\""年（２桁）+月（２桁）+日（２桁）+ロットナンバー．ＣＳＶ")
            strMSG = DateTime.Now.ToString()                            ' "yyyy/MM/dd HH:mm:ss"
            'strDAT = strMSG.Substring(0, 4) + strMSG.Substring(5, 2) + strMSG.Substring(8, 2)  ' yyyymmdd
            strDAT = strMSG.Substring(2, 2) + strMSG.Substring(5, 2) + strMSG.Substring(8, 2)   ' yymmdd
            FPath = cLOGFILEPATH & strDAT & stUserData.sLotNumber & ".CSV"

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.SetLogFileName() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

    '==========================================================================
    '   スクリーンキーボードの起動/終了処理
    '==========================================================================
#Region "スクリーンキーボードの起動処理"
    '''=========================================================================
    ''' <summary>スクリーンキーボードの起動処理</summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub StartSoftwareKeyBoard(ByRef ps As Process)

        Dim strMsg As String

        Try
            'ps = New Process
            ps.StartInfo.FileName = "osk.exe"
            ps.Start()

        Catch ex As Exception
            strMsg = "User.StartSoftwareKeyBoard() TRAP ERROR = " + ex.Message
            MsgBox(strMsg)
        End Try
    End Sub
#End Region

#Region "スクリーンキーボードの終了処理"
    '''=========================================================================
    ''' <summary>スクリーンキーボードの終了処理</summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub EndSoftwareKeyBoard(ByRef ps As Process)

        Dim strMsg As String

        Try
            If (ps.HasExited <> True) Then
                ps.Kill()
            End If

        Catch ex As Exception
            strMsg = "User.EndSoftwareKeyBoard() TRAP ERROR = " + ex.Message
            MsgBox(strMsg)
        End Try
    End Sub
#End Region

    '==========================================================================
    '   XYテーブルのステップ移動処理
    '==========================================================================
#Region "XYテーブルのステップ移動処理"
    '''=========================================================================
    ''' <summary>
    ''' XYテーブルのステップ移動処理"
    ''' </summary>
    ''' <param name="direct">Forward 前進：正、Backword 後進：負</param>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub StepMove(ByVal direct As Integer)

        Dim XPos As Integer, YPos As Integer
        Dim r As Integer                                                ' 戻り値
#If TXTY_USE Then
        Dim StepOffX As Double, StepOffY As Double                      ' ステップオフセット用変数
#End If

        Try

            XPos = stCounter.BlockCntX
            YPos = stCounter.BlockCntY

            If direct > 0 Then      ' Forward　前進
                If (XPos Mod 2) = 1 Then    ' 奇数列
                    YPos = YPos + 1
                    If YPos > stPLT.BNY * stPLT.Pny Then
                        YPos = stPLT.BNY * stPLT.Pny
                        XPos = XPos + 1
                        If XPos > stPLT.BNX * stPLT.Pnx Then
                            Exit Sub    ' 最後まで前進
                        End If
                    End If
                Else
                    YPos = YPos - 1
                    If YPos < 1 Then
                        YPos = 1
                        XPos = XPos + 1
                        If XPos > stPLT.BNX * stPLT.Pnx Then
                            Exit Sub    ' 最後まで前進
                        End If
                    End If

                End If
            Else                    ' Backword 後進
                If (XPos Mod 2) = 1 Then    ' 奇数列
                    YPos = YPos - 1
                    If YPos < 1 Then
                        YPos = 1
                        XPos = XPos - 1
                        If XPos < 1 Then
                            Exit Sub    ' 最初まで後進
                        End If
                    End If
                Else
                    YPos = YPos + 1
                    If YPos > stPLT.BNY * stPLT.Pny Then
                        YPos = stPLT.BNY * stPLT.Pny
                        XPos = XPos - 1
                    End If
                End If
            End If

#If TXTY_USE Then
            If stPLT.BNY > 1 Then
                StepOffX = stPLT.dblStepOffsetXDir / (stPLT.BNY - 1) * (YPos - 1)
            Else
                StepOffX = 0.0
            End If
            If stPLT.BNX > 1 Then
                StepOffY = stPLT.dblStepOffsetYDir / (stPLT.BNX - 1) * (XPos - 1)
            Else
                StepOffY = 0.0
            End If
            r = TSTEP(CShort(XPos), CShort(YPos), Int((XPos - 1) / stPLT.BNX) * stPLT.Pivx + StepOffX, Int((YPos - 1) / stPLT.BNY) * stPLT.Pivy + StepOffY) ' V1.0.1.0②
#Else
            r = TSTEP(CShort(XPos), CShort(YPos), Int((XPos - 1) / stPLT.BNX) * stPLT.Pivx, Int((YPos - 1) / stPLT.BNY) * stPLT.Pivy)
#End If
            r = ObjSys.EX_ZGETSRVSIGNAL(gSysPrm, CShort(r), 0)
            If (r = cFRS_NORMAL) Then
                Form1.Refresh()                                             ' パターン認識時画面表示のため
                stCounter.PlateCntX = CInt(Int((XPos - 1) / stPLT.BNX) + 1)
                stCounter.PlateCntY = CInt(Int((YPos - 1) / stPLT.BNY) + 1)
                stCounter.BlockCntX = (XPos - 1) Mod stPLT.BNX + 1
                stCounter.BlockCntY = (YPos - 1) Mod stPLT.BNY + 1
                Call Z_PRINT("■BLOCK(" & stCounter.BlockCntX.ToString("000") & "," & stCounter.BlockCntY.ToString("000") & ")" & vbCrLf)
                MoveFlagOfStepMove = True
            Else                         ' エラー
                Exit Sub
            End If

        Catch ex As Exception
            MsgBox("User.StepMove() TRAP ERROR = " + ex.Message)
        End Try
    End Sub

    Private MoveFlagOfStepMove As Boolean = False

    '''=========================================================================
    ''' <summary>
    ''' ステップ移動の有り無し確認フラグの初期化
    ''' </summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub InitForStepMove()
        MoveFlagOfStepMove = False
    End Sub

    '''=========================================================================
    ''' <summary>
    ''' ステップ移動後の位置取得
    ''' </summary>
    ''' <param name="xPos"></param>
    ''' <param name="yPos"></param>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub GetPosForStepMove(ByRef XPos As Integer, ByRef YPos As Integer)
        If MoveFlagOfStepMove Then
            XPos = (stCounter.PlateCntX - 1) * stPLT.BNX + stCounter.BlockCntX
            If ((stCounter.BlockCntX Mod 2) = 1) Then   ' 奇数Xブロック時
                YPos = (stCounter.PlateCntY - 1) * stPLT.BNY + stCounter.BlockCntY
            Else                                        ' 偶数Xブロック時
                YPos = stPLT.BNY * stPLT.Pny + 1 - stCounter.BlockCntY
            End If
        End If
    End Sub

#End Region

    'V2.0.0.0↓
#Region "■■　外部電源ＯＮ実行　■■"
    '''=========================================================================
    ''' <param name="Rn">     抵抗番号</param>
    ''' <returns> FUNC_OK = 外部電源ＯＮ（正常動作）
    '''           cERR_TRAP = エラー
    ''' </returns>
    '''=========================================================================
    Private Function Func_V_On_Ex(ByVal Rn As Short) As Short
        Dim i As Short
        Dim r As Short

        Func_V_On_Ex = FUNC_OK

        ' DC電源装置 電圧ON                 
        For i = 1 To EXTEQU Step 1
            If 0 <> stREG(Rn).intOnExtEqu(i) Then
                Call DebugLogOut(String.Format("外部電源ON {0}={1}", stREG(Rn).intOnExtEqu(i), stGPIB(stREG(Rn).intOnExtEqu(i)).strCON))
                r = V_On_Ex(stREG(Rn).intOnExtEqu(i))  ' 電圧ON

                If (r <> 0) Then                                        ' エラー ?
                    Func_V_On_Ex = FUNC_NG
                End If

            End If
        Next i
    End Function
#End Region

#Region "■■　外部電源ＯＦＦ実行　■■"
    '''=========================================================================
    ''' <param name="Rn">     抵抗番号</param>
    ''' <returns> FUNC_OK = ＧＮＤＯＦＦ（正常動作）
    '''           cERR_TRAP = エラー
    ''' </returns>
    '''=========================================================================
    Private Function Func_V_Off_Ex(ByVal Rn As Short) As Short
        Dim i As Short
        Dim r As Short

        Func_V_Off_Ex = FUNC_OK

        ' DC電源装置 電圧ON                 
        For i = 1 To EXTEQU Step 1
            If 0 <> stREG(Rn).intOffExtEqu(i) Then
                Call DebugLogOut(String.Format("外部電源OFF {0}={1}", stREG(Rn).intOffExtEqu(i), stGPIB(stREG(Rn).intOffExtEqu(i)).strCOFF))
                r = V_Off_Ex(stREG(Rn).intOffExtEqu(i))  ' 電圧Off
                If (r <> 0) Then                                        ' エラー ?
                    Func_V_Off_Ex = FUNC_NG
                End If
            End If
        Next i
    End Function
#End Region

#Region "■■　外部電源ＯＮ設定　有／無判定処理　■■"
    '''=========================================================================
    ''' <param name="Rn">     抵抗番号</param>
    ''' <returns> FUNC_OK = 外部電源ＯＮ設定あり
    '''           FUNC_NG = 外部電源ＯＮ設定なし
    ''' </returns>
    '''=========================================================================
    Private Function Func_V_On_Judge(ByVal Rn As Short) As Short
        Dim i As Short

        Func_V_On_Judge = FUNC_NG

        For i = 1 To EXTEQU Step 1
            If 0 <> stREG(Rn).intOnExtEqu(i) Then
                Func_V_On_Judge = FUNC_OK
                Exit For
            End If
        Next i

    End Function
#End Region

#Region "■■　外部電源ＯＮ機器番号取得処理　■■"
    '''=========================================================================
    ''' <param name="Rn">     抵抗番号</param>
    ''' <returns> FUNC_OK = 外部電源ＯＮ設定あり
    '''           FUNC_NG = 外部電源ＯＮ設定なし
    ''' </returns>
    '''=========================================================================
    Private Function Func_V_On_Number(ByVal Rn As Short) As Short
        Dim i As Short

        Func_V_On_Number = 0

        For i = 1 To EXTEQU Step 1
            If 0 <> stREG(Rn).intOnExtEqu(i) Then
                Func_V_On_Number = stREG(Rn).intOnExtEqu(i)
                Return (Func_V_On_Number)
            End If
        Next i

    End Function
#End Region

#Region "■■　外部電源ＯＦＦ設定　有／無判定処理　■■"
    '''=========================================================================
    ''' <param name="Rn">     抵抗番号</param>
    ''' <returns> FUNC_OK = 外部電源ＯＦＦ設定あり
    '''           FUNC_NG = 外部電源ＯＦＦ設定なし
    ''' </returns>
    '''=========================================================================
    Private Function Func_V_Off_Judge(ByVal Rn As Short) As Short
        Dim i As Short

        Func_V_Off_Judge = FUNC_NG

        For i = 1 To EXTEQU Step 1
            If 0 <> stREG(Rn).intOffExtEqu(i) Then
                Func_V_Off_Judge = FUNC_OK
                Exit For
            End If
        Next i
    End Function
#End Region
    'V2.0.0.0↑

#Region "基板処理終了時ブザーＯＮ"
    Public Sub Buzzer()
        For i As Integer = 1 To 3

            Form1.System1.SetSignalTower(SIGOUT_BZ1_ON Or SIGOUT_RED_BLK, 0)
            System.Threading.Thread.Sleep(1500)
            Form1.System1.SetSignalTower(0, SIGOUT_BZ1_ON Or SIGOUT_RED_BLK)
            System.Threading.Thread.Sleep(1000)
        Next
    End Sub
#End Region

    '==========================================================================
    '   下層処理ファンクション
    '==========================================================================

#Region "トリミング開始時の初期化処理"
    '===============================================================================
    '【機　能】 トリミング開始時の初期化処理
    '【引　数】 無し\
    '【戻り値】 無し
    '===============================================================================
    Public Sub UserParameterInitialize()
        giMultiMeter = -1                                   ' 測定レンジの初期化
        If UserSub.IsSpecialTrimType Then
            Call SetCutPosTimes(stUserData.lCutHosei)       ' カット位置補正頻度
        Else
            Call SetCutPosTimes(1)                          ' 毎回補正
        End If
        glCutPosCounter = Integer.MaxValue                  ' 補正カウンターリセット
        For i As Integer = 1 To MAXRGN
            gcPtnCorrval(i) = "NONE"                        ' パターンマッチ相関値他情報( "NONE","SAME",MANUAL")
        Next
    End Sub
#End Region

#Region "ロット開始処理"
    '''===============================================================================
    ''' <summary>
    ''' ロットスタートの時の処理を行う。
    ''' </summary>
    ''' <returns>True:ロットスタート可 False:ロットスタート不可</returns>
    ''' <remarks></remarks>
    '''===============================================================================
    Public Function LotStartSetting() As Boolean
        Try

            If UserSub.IsTRIM_MODE_ITTRFT() And UserSub.IsSpecialTrimType() Then    ' ユーザプログラム特殊処理
                If Not UserSub.StandardResistanceMeasure() Then                 ' 標準抵抗値チェック
                    Return (False)
                End If
            End If

            If UserSub.GetStartCheckStatus() Then                       ' ロットスタート初回のみ実施項目
                'V2.0.0.2①                If UserSub.IsTRIM_MODE_ITTRFT() And UserSub.IsSpecialTrimType() Then    ' ユーザプログラム特殊処理
                ' 'V2.2.1.7⑤                If UserSub.IsTRIM_MODE_ITTRFT() Then    ' ユーザプログラム特殊処理 'V2.0.0.2①IsSpecialTrimTypeチェック外す
                If UserSub.IsTRIM_MODE_ITTRFT() Or ((DGL = TRIM_MODE_CUT) And UserSub.IsTrimType5()) Then    ' ユーザプログラム特殊処理 'V2.0.0.2①IsSpecialTrimTypeチェック外す 'V2.2.1.7⑤ 
                    Dim Rtn As Short
                    Dim fLotInf As New FormEdit.frmLotInfoInput(False)
                    fLotInf.ShowDialog()
                    Rtn = fLotInf.sGetReturn()
                    fLotInf.Dispose()
                    If Rtn = cFRS_ERR_RST Then                          ' キャンセルリターン
                        Call Form1.System1.OperationLogging(gSysPrm, "データ確認キャンセル", "MANUAL")   'V2.0.0.2①
                        Return (False)
                    Else
                        Call Form1.System1.OperationLogging(gSysPrm, "ロット初期化処理実行1", "MANUAL")   'V2.0.0.2①
                        UserBas.stCounter.LotStart = DateTime.Now()     ' 設定データ確認時間（ロットスタート時間保存）
                        Call UserParameterInitialize()                  ' パラメータの初期化
                        Call UserSub.MakePrintFileHeader()              ' 印刷データヘッダー情報作成
                        ObjMain.LblLOT.Text = stUserData.sLotNumber
                        UserSub.ResetlResCounterForPrinter()            ' 印刷素子カウンターのリセット
                        stCounter.PlateCounter = 0                      ' 基板カウンター
                        Call UserBas.Disp_frmInfo(COUNTER.PRODUCT_INIT, COUNTER.NONE)

                        UserBas.stCounter.LotPrint = False              ' ロット終了時の印刷実行済みでTrue
                        If stCounter.LotCounter = 0 Then                ' 最初は、ロット交換が無いのでここでカウントアップする。
                            stCounter.LotCounter = 1
                        End If
                        'V2.2.0.034↓
                        'カウンターのクリア
                        If stMultiBlock.gMultiBlock <> 0 Then
                            For i As Integer = 0 To 5
                                stMultiBlock.BLOCK_DATA(i).gProcCnt = 0
                            Next i
                        End If
                        'V2.2.0.034↑

                    End If
                End If
                'V1.1.0.1①↓
            Else
                If frmAutoObj.gbFgAutoOperation = True Then
                    Call Form1.System1.OperationLogging(gSysPrm, "ロット初期化処理実行2", "MANUAL")   'V2.0.0.2①
                    UserBas.stCounter.LotStart = DateTime.Now()     ' 設定データ確認時間（ロットスタート時間保存）
                    Call UserParameterInitialize()                  ' パラメータの初期化
                    Call UserSub.MakePrintFileHeader()              ' 印刷データヘッダー情報作成
                    ObjMain.LblLOT.Text = stUserData.sLotNumber
                    UserSub.ResetlResCounterForPrinter()            ' 印刷素子カウンターのリセット
                    stCounter.PlateCounter = 0                      ' 基板カウンター
                    Call UserBas.Disp_frmInfo(COUNTER.PRODUCT_INIT, COUNTER.NONE)

                    UserBas.stCounter.LotPrint = False              ' ロット終了時の印刷実行済みでTrue
                    'V2.2.0.034↓
                    'カウンターのクリア
                    If stMultiBlock.gMultiBlock <> 0 Then
                        For i As Integer = 0 To 5
                            stMultiBlock.BLOCK_DATA(i).gProcCnt = 0
                        Next i
                    End If
                    'V2.2.0.034↑

                End If
                'V1.1.0.1①↑
            End If

            'V2.0.0.0②↓
            If DGL = TRIM_VARIATION_MEAS Then ' 測定値変動測定	'V2.0.0.1①
                If UserSub.bVariationMesStep Then ' 測定値変動検出機能位置し有効
                    stCounter.PlateCounter = UserSub.gVariationMeasPlateStartNo - 1     ' 基板カウンター
                End If
            End If                                              'V2.0.0.1①
            'V2.0.0.0②↑

            Return (True)

        Catch ex As Exception
            Call Z_PRINT("UserSub.LotStartSetting() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Function
#End Region

    'V2.0.0.0↓
#Region "■■　ＷＡＩＴ処理（セーフティチェック機能付）　■■"
    '''=========================================================================
    ''' <param name="W_TIME"> ＷＡＩＴ時間（セーフティチェック機能付）</param>
    ''' <returns> cFRS_NORMAL=正常
    '''           Not cFRS_NORMAL=異常
    ''' </returns>
    '''=========================================================================
    Public Function Func_Wait(ByVal W_TIME As Integer) As Integer
        Dim lngLoopCntr As Long
        Dim ww_time As Integer
        Dim i As Long
        Dim r As Integer

        Func_Wait = cFRS_NORMAL

        lngLoopCntr = W_TIME \ DEV_TIMER
        ww_time = W_TIME Mod DEV_TIMER


        For i = 1 To lngLoopCntr
            ZWAIT(DEV_TIMER)

            ' セーフティチェック
            r = SafetyCheck()                                   ' セーフティチェック
            If (r <> cFRS_NORMAL) Then                          ' エラー ?
                Func_Wait = r
                Exit For
            End If

        Next i

        ZWAIT(ww_time)

    End Function
#End Region
    'V2.0.0.0↑

#Region "基板処理開始処理"
    '''===============================================================================
    ''' <summary>
    ''' 基板処理開始時の初期化処理
    ''' </summary>
    ''' <remarks></remarks>
    '''===============================================================================
    Public Sub PlateStartSetting()

        gisCutPosExecuteAutoNG = False                          ' V1.0.4.3⑥カット補正自動判定有無し（抵抗データの設定（stPTN(i).PtnFlg）から求める）
        gisCutPosExecute = False                                ' カット位置補正の有り無しチェック→補正値と一致度をログに出力する為
        For i As Integer = 1 To stPLT.PtnCount
            If stPTN(i).PtnFlg <> CUT_PATTERN_NONE Then
                gisCutPosExecute = True
            End If
            If stPTN(i).PtnFlg = CUT_PATTERN_AUTO_NG Then       ' V1.0.4.3⑥
                gisCutPosExecuteAutoNG = True                   ' V1.0.4.3⑥
            End If                                              ' V1.0.4.3⑥
        Next

        UserBas.stCounter.BlockCounter = 0                      ' ブロックカウンター
        UserBas.stCounter.TrimCounter = 0                       ' ﾄﾘﾐﾝｸﾞ数(ﾜｰｸ投入数)
        UserBas.stCounter.OK_Counter = 0                        ' OK数
        UserBas.stCounter.NG_Counter = 0                        ' NG数
        UserBas.stCounter.ITHigh = 0                            ' 初期測定上限値異常
        UserBas.stCounter.ITLow = 0                             ' 初期測定下限値異常
        UserBas.stCounter.ITOpen = 0                            ' 測定値異常
        UserBas.stCounter.FTHigh = 0                            ' 最終測定上限値異常
        UserBas.stCounter.FTLow = 0                             ' 最終測定下限値異常
        UserBas.stCounter.FTOpen = 0                            ' 測定値異常
        UserBas.stCounter.Pattern = 0                           ' カット位置補正の判定 'V1.2.0.0③
        UserBas.stCounter.VaNG = 0                              ' 再測定変化量エラー追加　V2.0.0.0②
        UserBas.stCounter.StartTime = DateTime.Now()            ' 基板処理スタート時間保存

        UserBas.stCounter.ValHigh = 0                            ' 最終測定上限値異常    'V2.2.0.029
        UserBas.stCounter.ValLow = 0                             ' 最終測定下限値異常    'V2.2.0.029

        gObjFrmDistribute.ClearMultiCountPlateData()            '複数抵抗値用の基板カウンタクリア

        Call Disp_frmInfo(COUNTER.ALLDATA_DISP, COUNTER.NONE)   ' 生産数,良品数,表示(frmInfo画面)'V2.2.0.0⑯

        stCounter.PlateCounter = stCounter.PlateCounter + 1     ' 基板処理カウンター

        Call UserParameterInitialize()                          ' パラメータの初期化
        Call UserSub.NgJudgeReset()                             ' 素子毎の判定リセット

        For i As Integer = 1 To stPLT.PtnCount                      ' パターン登録数分初期化設定する
            stPTN(i).dblDRX = 0.0                                   ' ズレ量X
            stPTN(i).dblDRY = 0.0                                   ' ズレ量Y
        Next i



        Z_CLS()                                                 ' ログ画面クリア
    End Sub

#End Region

#Region "■■　カットオフによる目標値　■■"
    '''=========================================================================
    ''' <param name="RNo"> 抵抗番号</param>
    ''' <param name="Cno"> カット番号</param>
    ''' <param name="wNom"> 目標値</param>
    ''' <returns> カットオフによる目標値
    ''' </returns>
    '''=========================================================================
    Public Function Func_CalNomForCutOff(ByVal RNo As Short, ByVal Cno As Short, ByVal wNom As Double) As Double
        Dim dNom As Double
        Try
            Func_CalNomForCutOff = 0.0

            If (0 = stREG(RNo).intMode) Then
                dNom = wNom * System.Math.Abs(1 + (stREG(RNo).STCUT(Cno).dblCOF * 0.01))       '比率
            Else
                dNom = wNom + stREG(RNo).STCUT(Cno).dblCOF                                     '数値（絶対値）
            End If

            Func_CalNomForCutOff = dNom

        Catch ex As Exception
            Call Z_PRINT("User.Func_CalNomForCutOff() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Function
#End Region

#Region "■■ 抵抗名検索 ■■"
    '''=========================================================================
    ''' <summary>
    ''' 抵抗名検索
    ''' </summary>
    ''' <param name="rn">抵抗番号</param>
    ''' <param name="Code">検索文字列</param>
    ''' <returns>True = 一致, False = 不一致</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function CheckResistorName(ByVal rn As Integer, ByVal Code As String) As Boolean
        Try
            If Code = Microsoft.VisualBasic.Strings.Left(stREG(rn).strRNO, Len(Code)) Then
                CheckResistorName = True
            Else
                CheckResistorName = False
            End If
        Catch ex As Exception
            Call Z_PRINT("User.CheckResistorName() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Function
#End Region

#Region "■■ カット抵抗判定 ■■"
#If False Then
    '''=========================================================================
    ''' <summary>
    ''' カット抵抗判定
    ''' </summary>
    ''' <param name="rn">抵抗番号</param>
    ''' <returns>True = カットありの抵抗, False = 測定のみの抵抗</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsCutResistor(ByVal rn As Integer) As Boolean
        Try
            If stREG(rn).intSLP = SLP_VTRIMPLS Or stREG(rn).intSLP = SLP_VTRIMMNS Or stREG(rn).intSLP = SLP_RTRM Then
                IsCutResistor = True
            Else
                IsCutResistor = False
            End If
        Catch ex As Exception
            Call Z_PRINT("User.IsCutResistor() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function
#End If
#End Region

#Region "■■ カット抵抗判定(文字マーク含む) ■■"
    'V2.0.0.0⑮　IsCutResistorIncMarking　へ変更
    'V1.0.4.3⑤↓
    'V2.0.0.0⑮    '''=========================================================================
    'V2.0.0.0⑮    ''' <summary>
    'V2.0.0.0⑮    ''' カット抵抗判定(文字マーク含む) 
    'V2.0.0.0⑮    ''' </summary>
    'V2.0.0.0⑮    ''' <param name="rn">抵抗番号</param>
    'V2.0.0.0⑮    ''' <returns>True = カットありの抵抗, False = 測定のみの抵抗</returns>
    'V2.0.0.0⑮    ''' <remarks></remarks>
    'V2.0.0.0⑮    '''=========================================================================
    'V2.0.0.0⑮    Public Function IsCutResistorIncCharacter(ByVal rn As Integer) As Boolean
    'V2.0.0.0⑮        Try
    'V2.0.0.0⑮            If stREG(rn).intSLP = SLP_VTRIMPLS Or stREG(rn).intSLP = SLP_VTRIMMNS Or stREG(rn).intSLP = SLP_RTRM Or stREG(rn).intSLP = SLP_OK_MARK Or stREG(rn).intSLP = SLP_NG_MARK Then
    'V2.0.0.0⑮                IsCutResistorIncCharacter = True
    'V2.0.0.0⑮            Else
    'V2.0.0.0⑮                IsCutResistorIncCharacter = False
    'V2.0.0.0⑮            End If
    'V2.0.0.0⑮        Catch ex As Exception
    'V2.0.0.0⑮            Call Z_PRINT("User.IsCutResistorIncCharacter() TRAP ERROR = " & ex.Message & vbCrLf)
    'V2.0.0.0⑮        End Try
    'V2.0.0.0⑮    End Function
    'V1.0.4.3⑤↑
#End Region

#Region "■■ ｎ番目のカット抵抗番号取得 ■■"
    '''=========================================================================
    ''' <summary>
    ''' ｎ番目のカット抵抗番号取得
    ''' </summary>
    ''' <param name="rn">抵抗番号</param>
    ''' <returns>True = カットありの抵抗, False = 測定のみの抵抗</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function GetCutResistorNo(ByVal rn As Integer) As Integer
        Dim iResNo As Integer, iCnt As Integer
        Try
            iCnt = 0
            For iResNo = 1 To MAXRNO Step 1
                If IsCutResistor(iResNo) Then           ' カット有（測定のみでない）抵抗の場合
                    iCnt = iCnt + 1
                    If iCnt = rn Then                   'カットする抵抗のみの求める順番と同じ
                        GetCutResistorNo = iResNo
                        Exit Function
                    End If
                End If
                If iResNo > stPLT.RCount Then           ' 登録数を超えたら
                    GetCutResistorNo = iResNo
                    Call Z_PRINT("GetCutResistorNo 検索抵抗番号が登録数を超えています = " & rn.ToString & vbCrLf)
                    Exit Function
                End If
            Next
        Catch ex As Exception
            Call Z_PRINT("User.IsCutResistor() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function
#End Region

#Region "■■ 現在の抵抗番号が、カットのみの抵抗の何番目かの取得 ■■"
    '''=========================================================================
    ''' <summary>
    ''' 現在の抵抗番号が、カットのみの抵抗の何番目かの取得（GetCutResistorNoの逆）
    ''' </summary>
    ''' <param name="rn">現在の抵抗番号</param>
    ''' <returns>True = カットありの抵抗, False = 測定のみの抵抗</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function GetResistorNo(ByVal rn As Integer) As Integer
        Dim iResNo As Integer, iCnt As Integer
        Try
            'V2.0.0.0⑩↓
            If UserSub.IsTrimType3() Or UserSub.IsTrimType4() Then
                iCnt = UserSub.GetResNumberInCircuit(rn)
                Return (iCnt)
            End If
            'V2.0.0.0⑩↑

            iCnt = 0
            For iResNo = 1 To rn Step 1
                If IsCutResistor(iResNo) Then           ' カット有（測定のみでない）抵抗の場合
                    iCnt = iCnt + 1
                End If
            Next
            GetResistorNo = iCnt
        Catch ex As Exception
            Call Z_PRINT("User.GetResistorNo() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function
#End Region

#Region "■■ 測定とＮＧマーキングを除いた番号での抵抗目標値の取得 ■■"

    '''=========================================================================
    ''' <summary>
    ''' 測定とＮＧマーキングを除いた番号での抵抗目標値の取得
    ''' </summary>
    ''' <param name="rn">測定のみを除いた番号での抵抗番号</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function GetNominalResistance(ByVal rn As Integer) As Double
        Try
            GetNominalResistance = stREG(UserBas.GetCutResistorNo(rn)).dblNOM ' 目標抵抗値
        Catch ex As Exception
            Call Z_PRINT("User.GetNominalResistance() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function
#End Region

#Region "■■ 測定のみを除いた抵抗数の取得 ■■"
    '''=========================================================================
    ''' <summary>
    ''' 測定のみを除いた抵抗数の取得
    ''' </summary>
    ''' <returns>抵抗数</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function GetRCountExceptMeasure() As Short
        Try
            GetRCountExceptMeasure = 0
            'V2.0.0.0⑩↓
            If UserSub.IsTrimType3() Then
                GetRCountExceptMeasure = UserSub.CircuitResistorCount()
                Return (GetRCountExceptMeasure)
            End If
            'V2.0.0.0⑩↑

            'V2.0.0.0①↓
            If UserSub.IsTrimType1() Or UserSub.IsTrimType4() Then  ' 温度センサーの時は１固定
                GetRCountExceptMeasure = 1
                Return (GetRCountExceptMeasure)
            End If
            'V2.0.0.0①↑
            For iResNo As Integer = 1 To stPLT.RCount Step 1
                If IsCutResistor(iResNo) Then           ' カット有（測定のみでない）抵抗の場合
                    GetRCountExceptMeasure = GetRCountExceptMeasure + 1
                End If
            Next

        Catch ex As Exception
            Call Z_PRINT("User.GetRCountExceptMeasure() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function
#End Region

    'V2.0.0.0⑩↓
#Region "■■ 測定のみを除いた抵抗数の取得 ■■"
    '''=========================================================================
    ''' <summary>
    ''' 測定のみを除いた抵抗数の取得(サーキット投入前のバージョン）
    ''' </summary>
    ''' <returns>抵抗数</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function GetRCountExceptMeasureOldVersion() As Short
        Try
            GetRCountExceptMeasureOldVersion = 0
            For iResNo As Integer = 1 To stPLT.RCount Step 1
                If IsCutResistor(iResNo) Then           ' カット有（測定のみでない）抵抗の場合
                    GetRCountExceptMeasureOldVersion = GetRCountExceptMeasureOldVersion + 1
                End If
            Next

        Catch ex As Exception
            Call Z_PRINT("User.GetRCountExceptMeasureOldVersion() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function
#End Region
    'V2.0.0.0⑩↑

#Region "外部測定器のレンジ設定変更（ADVANTEST　R6581)"
    '''=========================================================================
    ''' <summary>
    ''' 外部測定器のレンジ設定変更（ADVANTEST　R6581)
    ''' </summary>
    ''' <param name="Inmes">(IN)内部/外部種別(0=内部測定器, 1以降=外部測定器番号)</param>
    ''' <param name="dNOMx">(IN)目標値</param>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub Change_Range_DMM(ByVal Inmes As Short, ByVal dNOMx As Double)
        Dim Gno As Short
        Dim Pos As Integer = -1
        Dim iMultiMeter As Integer

        Try
            sStrTrig = ""   'V2.2.1.4①
            If Inmes < 1 Then      ' 内部測定器の場合
                Exit Sub
            Else
                Gno = Inmes
            End If

            'V2.2.1.1⑤ ↓'V2.2.1.4①
            sStrTrig = ""
            Dim result As Boolean = stGPIB(Gno).strGNAM.ToUpper().Contains("HIOKI")
            Dim bChangeFlg As Boolean = False

            If result = True Then
                Dim SearchStr As String = "RES:RANG:AUTO OFF"
                Dim PutStr As String = ":RES:RANG"
                ' HIOKIの場合、設定コマンドの中に「 "RES:RANG:AUTO OFF"」があったら、目標値を送ってレンジの設定を行う。
                Dim cmd1 = stGPIB(Gno).strCCMD1.ToUpper().IndexOf(SearchStr)
                Dim cmd2 = stGPIB(Gno).strCCMD2.ToUpper().IndexOf(SearchStr)
                Dim cmd3 = stGPIB(Gno).strCCMD3.ToUpper().IndexOf(SearchStr)

                If cmd1 >= 0 OrElse cmd2 >= 0 OrElse cmd3 >= 0 Then

                    If gLastsetNomx <> dNOMx Then
                        bChangeFlg = True
                        gLastsetNomx = dNOMx
                    End If

                    Change_HIOKI_CMD(Gno, dNOMx, bChangeFlg)

                    'Dim arr() As String = stGPIB(Gno).strCTRG.Split(";")
                    'For i As Integer = 0 To arr.Length - 1
                    '    Pos = arr(i).ToUpper.IndexOf(PutStr)
                    '    If Pos >= 0 Then
                    '        ' 存在したらレンジ目標値指定に置き換える
                    '        arr(i) = PutStr & " " & CInt(dNOMx).ToString()
                    '        Exit For
                    '    End If
                    'Next
                    'Dim TrigStr As String = ""
                    '' 置き換えた文字列を結合する
                    'For i As Integer = 0 To arr.Length - 1
                    '    If TrigStr.Trim <> "" Then
                    '        TrigStr = TrigStr & ";"
                    '    End If
                    '    TrigStr = TrigStr & arr(i)
                    'Next
                End If

            Else
                'V2.2.1.1⑤ ↑

                ' レンジテーブル
                ' R1  　－
                ' R2     10Ω
                ' R3    100Ω
                ' R4   1000Ω　
                ' R5    10KΩ
                ' R6   100KΩ
                ' R7  1000KΩ 
                ' R8    10MΩ
                ' R9   100MΩ
                ' R10 1000MΩ
                ' R20  　－
                If dNOMx <= 11.5 Then
                    iMultiMeter = 2
                ElseIf dNOMx <= 115.0 Then
                    iMultiMeter = 3
                ElseIf dNOMx <= 1150.0 Then
                    iMultiMeter = 4
                ElseIf dNOMx <= 11500.0 Then
                    iMultiMeter = 5
                ElseIf dNOMx <= 115000.0 Then
                    iMultiMeter = 6
                ElseIf dNOMx <= 1150000.0 Then
                    iMultiMeter = 7
                ElseIf dNOMx <= 11500000.0 Then
                    iMultiMeter = 8
                ElseIf dNOMx <= 115000000.0 Then
                    iMultiMeter = 9
                    'ElseIf dNOMx <= 1000000000.0 Then　２桁はサポートしていません。
                    '    iMultiMeter = 10
                Else
                    iMultiMeter = 0 ' AUTOレンジ
                End If

                '###1031            If iMultiMeter = giMultiMeter Then                      ' 前回と同じレンジなので送信しない。
                '###1031                Return
                '###1031            Else
                Pos = stGPIB(Gno).strCTRG.IndexOf("R")
                If Pos >= 0 And Char.IsNumber(stGPIB(Gno).strCTRG.Substring(Pos + 1, 1)) Then    ' 数字の場合のみ置換
                    Mid(stGPIB(Gno).strCTRG, Pos + 2) = iMultiMeter.ToString
                    giMultiMeter = iMultiMeter
                End If
                '###1031            End If

            End If


        Catch ex As Exception
            Call Z_PRINT("User.Change_Range_DMM() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub
#End Region

#Region "デバッグ用のログ出力"
    Public Sub DebugLogOut(ByVal strLOG As String)
        Dim WS As IO.StreamWriter

        Try
            If bDebugLogOut Then
                WS = IO.File.AppendText(gsLogFileName.Replace(".CSV", ".LOG"))
                WS.WriteLine(DateTime.Now().ToString("yyyy/MM/dd,HH:mm:ss") & "," & strLOG)
                WS.Close()
            End If
        Catch ex As Exception
            Call Z_PRINT("UserBas.DebugLogOut() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub
#End Region

    'V1.2.0.2↓
#Region "デバッグ用のログ出力"
    Public Sub NgCutDebugLogOut(ByVal strLOG As String)
        Dim WS As IO.StreamWriter

        Try
            If bNgCutDebugLogOut Then
                WS = IO.File.AppendText(gsLogFileName.Replace(".CSV", "_NGCUT.LOG"))
                WS.WriteLine(DateTime.Now().ToString("yyyy/MM/dd,HH:mm:ss") & "," & strLOG)
                WS.Close()
            End If
        Catch ex As Exception
            Call Z_PRINT("UserBas.NgCutDebugLogOut() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub
#End Region
    'V1.2.0.2↑

    'V2.1.0.0①↓
#Region "カット毎の抵抗値変化量判定機能デバッグ用のログ出力"
    Public Sub CutVariationDebugLogOut(ByVal strLOG As String)
        Dim WS As IO.StreamWriter

        Try
            If bCutVariationDebugLogOut Then
                WS = IO.File.AppendText(gsLogFileName.Replace(".CSV", "_CUTVA.LOG"))
                WS.WriteLine(DateTime.Now().ToString("yyyy/MM/dd,HH:mm:ss") & "," & strLOG)
                WS.Close()
            End If
        Catch ex As Exception
            Call Z_PRINT("UserBas.CutVariationDebugLogOut() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub
#End Region
    'V2.1.0.0①↑

#Region "ロット処理関連のデバッグ用のログ出力"
    Public Sub AutoOperationDebugLogOut(ByVal strLOG As String)
        Dim WS As IO.StreamWriter

        Try
            If giAutoOperationDebugLogOut Then
                If gsLogFileName Is Nothing = False Then
                    WS = IO.File.AppendText(gsLogFileName.Replace(".CSV", "_OPE.LOG"))
                    WS.WriteLine(DateTime.Now().ToString("yyyy/MM/dd,HH:mm:ss") & "," & strLOG)
                    WS.Close()
                End If
            End If
        Catch ex As Exception
            Call Z_PRINT("UserBas.AutoOperationDebugLogOut() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub
#End Region


    ' ###1040の修正追加START

#Region "ﾋﾞｰﾑﾎﾟｼﾞｼｮﾅのオフセット値の取得"
    '''=========================================================================
    '''<summary>ﾋﾞｰﾑﾎﾟｼﾞｼｮﾅのオフセット値の取得</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub ChangeTeachBlockPosition(ByRef z_xoff As Double, ByRef z_yoff As Double)
        z_xoff = z_xoff - (stPLT.TeachBlockX - 1) * stPLT.zsx
        z_yoff = z_yoff - (stPLT.TeachBlockY - 1) * stPLT.zsy
    End Sub
#End Region

#Region "ﾋﾞｰﾑﾎﾟｼﾞｼｮﾅのオフセット値の取得"
    '''=========================================================================
    '''<summary>ﾋﾞｰﾑﾎﾟｼﾞｼｮﾅのオフセット値の取得</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub GetBpOffset(ByRef BpOffX As Double, ByRef BpOffY As Double)
        BpOffX = stPLT.BPOX
        BpOffY = stPLT.BPOY
    End Sub
#End Region

#Region "ブロックサイズの取得"
    '''=========================================================================
    '''<summary>ブロックサイズの取得</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub GetBlockSize(ByRef BsizeX As Double, ByRef BsizeY As Double)

        BsizeX = stPLT.zsx
        BsizeY = stPLT.zsy
    End Sub
#End Region

#Region "ﾋﾞｰﾑﾎﾟｼﾞｼｮﾅのオフセット値の設定"
    '''=========================================================================
    '''<summary>ﾋﾞｰﾑﾎﾟｼﾞｼｮﾅのオフセット値の設定</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub SetBpOffset(ByVal BpOffX As Double, ByVal BpOffY As Double)
        Dim strMSG As String

        Try

            With stPLT
                If .BPOX <> 0 And BpOffX = 0 And .BPOY <> 0 And BpOffY = 0 Then       ' ###270
                    strMSG = "SetBpOffset() : BP Offset Update 0 :BpOffX= " + CStr(BpOffX) + " :BpOffY= " + CStr(BpOffY) + " :dblBpOffSetXDir=" + CStr(.BPOX) + " :dblBpOffSetYDir=" + CStr(.BPOY)
                    MsgBox(strMSG)
                    Exit Sub
                End If
                .BPOX = BpOffX
                .BPOY = BpOffY

                'INTRIM側のオフセット値も更新する
                Call ObjSys.EX_BPOFF(gSysPrm, .BPOX, .BPOY)

            End With
        Catch ex As Exception
            strMSG = "UserBas.SetBpOffset() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "プレート内のブロック数を取得する"

    '''=========================================================================
    '''<summary>プレート内のブロック数を取得する</summary>
    '''=========================================================================
    Public Function GetBlockCnt() As Integer
        Try
            With stPLT
                GetBlockCnt = .BNX * .BNY

                Exit Function
            End With

        Catch ex As Exception
            Dim strMSG As String
            strMSG = "UserBas.GetBlockCnt() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try

    End Function
#End Region

#Region "プレート数を取得する"
    '''=========================================================================
    '''<summary>プレート数を取得する</summary>
    '''=========================================================================
    Public Function GetPlateCnt() As Integer
        Try
            With stPLT
                GetPlateCnt = .Pnx * .Pny

                Exit Function
            End With
        Catch ex As Exception
            Dim strMSG As String
            strMSG = "UserBas.GetPlateCnt() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Function
#End Region

#Region "指定ブロック番号のステージ位置を取得"
    '''=========================================================================
    ''' <summary>指定されたブロック番号の位置情報を取得する。</summary>
    ''' <param name="curPltNo">   (INP)現在のプレート番号を設定</param>
    ''' <param name="curBlkNo">   (INP)現在のブロック番号を設定</param>
    ''' <param name="stgx">       (OUT)ステージ位置X</param>
    ''' <param name="stgy">       (OUT)ステージ位置Y</param>
    ''' <param name="dispPltPosX">(OUT)プレート番号X</param>
    ''' <param name="dispPltPosY">(OUT)プレート番号Y</param>
    ''' <param name="dispBlkPosX">(OUT)ブロック番号X</param>
    ''' <param name="dispBlkPosY">(OUT)ブロック番号Y</param>
    ''' <returns>トリミングの最終ブロックの場合BLOCK_END＝1、
    '''          最終プレート最終ブロックの場合PLATE_BLOCK_END=2を返す。</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function GetTargetStagePos(ByVal curPltNo As Integer, ByVal curBlkNo As Integer, _
                                        ByRef stgx As Double, ByRef stgy As Double, _
                                        ByRef dispPltPosX As Integer, ByRef dispPltPosY As Integer, _
                                        ByRef dispBlkPosX As Integer, ByRef dispBlkPosY As Integer) As Integer
        Dim strMsg As String
        Dim workStgPosX As Double
        Dim workStgPosY As Double
        Dim totalPlateCnt As Integer
        Dim totalBlockCnt As Integer

        GetTargetStagePos = cFRS_NORMAL
        Try
            ' 基準位置、ステップ方向により次のブロック位置を取得する。
            With stPLT
                ' 最終ポジションの判定
                totalPlateCnt = .Pnx * .Pny
                totalBlockCnt = .BNX * .BNY             '###115 千鳥対応

                ' パラメータチェック
                If ((curPltNo < 0) Or (curBlkNo < 0)) Then
                    GetTargetStagePos = -1 * ERR_CMD_PRM
                    Exit Function
                End If

                If (curPltNo > totalPlateCnt) Then
                    GetTargetStagePos = PLATE_BLOCK_END
                    Exit Function
                ElseIf (curBlkNo > totalBlockCnt) Then
                    GetTargetStagePos = BLOCK_END
                    Exit Function
                End If

                ' データ取得対象は"0"オリジンのため、ここで一つ減算する。
                curPltNo = curPltNo - 1
                curBlkNo = curBlkNo - 1

                ' ステップ&(リピート方向)
                'If (.intDirStepRepeat = STEP_RPT_Y) _
                '    Or (.intDirStepRepeat = STEP_RPT_CHIPXSTPY) Then
                '    ' Y方向
                '    Call GetBlockPos_StpY(curPltNo, curBlkNo, gSysPrm.stDEV.giBpDirXy, _
                '                    .intPlateCntXDir, .intPlateCntYDir, _
                '                    .intBlockCntXDir, .intBlockCntYDir, workStgPosX, workStgPosY, _
                '                    dispPltPosX, dispPltPosY, dispBlkPosX, dispBlkPosY)
                'ElseIf (.intDirStepRepeat = STEP_RPT_X) _
                '       Or (.intDirStepRepeat = STEP_RPT_CHIPYSTPX) Then
                '    ' X方向
                '    Call GetBlockPos_StpX(curPltNo, curBlkNo, gSysPrm.stDEV.giBpDirXy, _
                '                    .intPlateCntXDir, .intPlateCntYDir, _
                '                    .intBlockCntXDir, .intBlockCntYDir, workStgPosX, workStgPosY, _
                '                    dispPltPosX, dispPltPosY, dispBlkPosX, dispBlkPosY)
                'Else
                '    ' ステップ&リピートなし
                '    '----- ###169↓ -----
                '    ' ステップ&リピートなしでも表示用ブロック数を更新するため下記をCallする
                '    Call GetBlockPos_StpY(curPltNo, curBlkNo, gSysPrm.stDEV.giBpDirXy, _
                '                    .intPlateCntXDir, .intPlateCntYDir, _
                '                    .intBlockCntXDir, .intBlockCntYDir, workStgPosX, workStgPosY, _
                '                    dispPltPosX, dispPltPosY, dispBlkPosX, dispBlkPosY)

                '    workStgPosX = 0.0                                   ' ステージ位置X,Yは0に再設定
                '    workStgPosY = 0.0

                '    'dispPltPosX = 1
                '    'dispPltPosY = 1
                '    'dispBlkPosX = 1
                '    'dispBlkPosY = 1
                '    '----- ###169↑ -----
                'End If

                stgx = workStgPosX
                stgy = workStgPosY
            End With

        Catch ex As Exception
            strMsg = "basTrimming.GetTargetStagePos() TRAP ERROR = " + ex.Message
            MsgBox(strMsg)
        End Try
    End Function
#End Region

#Region "ログ表示の向けに、指定ブロックXYからステージグループ番号、ブロック番号を取得する"
    '''=========================================================================
    '''<summary>指定ブロックXYのポジション情報から、</summary>
    ''' <param name="curBlkNoX">(INP)現在ブロックのX位置（行位置）</param>
    ''' <param name="curBlkNoY">(INP)現在ブロックのY位置（列位置）</param>
    ''' <param name="stgGrpNoX">(OUT)ステージグループ番号X</param>
    ''' <param name="stgGrpNoY">(OUT)ステージグループ番号Y</param>
    ''' <param name="blockNoX"> (OUT)ステージグループ番号を加味したブロック番号X</param>
    ''' <param name="blockNoY"> (OUT)ステージグループ番号を加味したブロック番号Y</param>
    '''<returns></returns>
    '''<remarks></remarks>
    '''  プレートの並び　　プレート内部の並び
    '''  ____ ____ ____      ____ ____
    ''' | ⑨ | ⑧ | ① |　　| ⑧ | ① |
    ''' |____|____|____|    |____|____|
    ''' | ⑩ | ⑦ | ② |    | ⑦ | ② |    
    ''' |____|____|____|    |____|____|
    ''' | ⑪ | ⑥ | ③ |    | ⑥ | ③ |
    ''' |____|____|____|    |____|____|
    ''' | ⑫ | ⑤ | ④ |    | ⑤ | ④ |
    ''' |____|____|____|    |____|____|
    ''' 
    '''=========================================================================
    Public Function GetDisplayPosInfo(ByVal curBlkNoX As Integer, ByVal curBlkNoY As Integer, _
                ByRef stgGrpNoX As Integer, ByRef stgGrpNoY As Integer, _
                ByRef blockNoX As Integer, ByRef blockNoY As Integer) As Boolean

        Dim strMSG As String

        GetDisplayPosInfo = True
        Try

            ''----- ###165↓ -----
            'With typPlateInfo

            '    If (.intResistDir = 0) Then                             ' 抵抗(ﾁｯﾌﾟ)並び方向 = X方向の場合
            '        ' ステージグループ番号X = 1
            '        stgGrpNoX = 1
            '        ' ステージグループ番号Y = ブロック番号Y / ステージグループ内ブロック数
            '        If ((curBlkNoY Mod .intBlkCntInStgGrpY) <> 0) Then  ' 余り有り ? 
            '            stgGrpNoY = curBlkNoY \ .intBlkCntInStgGrpY + 1
            '        Else
            '            stgGrpNoY = curBlkNoY \ .intBlkCntInStgGrpY
            '        End If

            '        ' ブロック番号X,Y
            '        blockNoX = curBlkNoX
            '        blockNoY = curBlkNoY

            '    Else                                                    ' 抵抗(ﾁｯﾌﾟ)並び方向 = Y方向の場合
            '        ' ステージグループ番号Y = 1
            '        stgGrpNoY = 1
            '        ' ステージグループ番号X = ブロック番号X / ステージグループ内ブロック数
            '        If ((curBlkNoY Mod .intBlkCntInStgGrpY) <> 0) Then  ' 余り有り ? 
            '            stgGrpNoX = curBlkNoX \ .intBlkCntInStgGrpX + 1
            '        Else
            '            stgGrpNoX = curBlkNoX \ .intBlkCntInStgGrpX
            '        End If

            '        ' ブロック番号X,Y
            '        blockNoX = curBlkNoX
            '        blockNoY = curBlkNoY
            '    End If

            'End With

            'With typPlateInfo
            '    'ステージグループ間隔を加味したブロック位置の算出
            '    '   →チップ方向ステップがある場合のステップのカウント方法を別途検討が必要
            '    ' X方向
            '    'If (.intBlkCntInStgGrpX <> 0) Then
            '    If (.dblStgGrpItvX <> 0) Then
            '        '   ステージグループ間隔が設定されている場合。
            '        '   →「現在ブロック/ステージグループ数」がステージグループの番号
            '        '   　「現在ブロック/ステージグループ数の余り」がステージグループ内部のブロック番号
            '        'X方向チップステップがある場合
            '        '   →チップステップのカウントも加算する。
            '        '   　ステージグループは、「現在のブロック/（ステージグループ内ブロック数＊チップステップ数）」
            '        '   　ブロック数は、「現在のブロック/チップステップ数」
            '        '   　チップステップ数は、「(現在のブロック/チップステップ数)の余り」
            '        stgGrpNoX = (curBlkNoX + 1) \ .intBlkCntInStgGrpX
            '        blockNoX = curBlkNoX Mod .intBlkCntInStgGrpX
            '        If (blockNoX = 0) Then
            '            blockNoX = .intBlkCntInStgGrpX
            '        End If
            '    Else
            '        '   ステージグループ間隔が設定されていない場合
            '        '   →現在ブロック=ステージグループ番号
            '        '   　現在ブロック=ブロック番号
            '        stgGrpNoX = curBlkNoX
            '        blockNoX = curBlkNoX
            '    End If

            '    'Y方向
            '    'If (.intBlkCntInStgGrpY <> 0) Then
            '    If (.dblStgGrpItvY <> 0) Then
            '        '   ステージグループ間隔が設定されている場合。
            '        '   →「現在ブロック/ステージグループ数」がステージグループの番号
            '        '   　「現在ブロック/ステージグループ数の余り」がステージグループ内部のブロック番号
            '        stgGrpNoY = (curBlkNoY + 1) \ .intBlkCntInStgGrpY
            '        blockNoY = curBlkNoY Mod .intBlkCntInStgGrpY
            '        If (blockNoY = 0) Then
            '            blockNoY = .intBlkCntInStgGrpY
            '        End If
            '    Else
            '        '   ステージグループ間隔が設定されていない場合
            '        '   →現在ブロック=ステージグループ番号
            '        '   　現在ブロック=ブロック番号
            '        stgGrpNoY = curBlkNoY
            '        blockNoY = curBlkNoY
            '    End If
            'End With
            '----- ###165↑ -----

        Catch ex As Exception
            strMSG = "DataAccess.GetDisplayPosInfo() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            GetDisplayPosInfo = False
        End Try
    End Function
#End Region

#Region "例外発生時に表示するﾒｯｾｰｼﾞﾎﾞｯｸｽ"
    ''' <summary>例外発生時に表示するﾒｯｾｰｼﾞﾎﾞｯｸｽ</summary>
    Public Sub MsgBox_Exception(ByRef exMsg As String, ByRef Obj As Object)
        Dim st As New StackTrace
        Dim msg As String
        Try
            ' GetFrame(0)=GetMethod, GetFrame(1)=CallerMethod
            msg = st.GetFrame(1).GetMethod.Name & "() TRAP ERROR = " & exMsg
            Call MsgBox(Obj.Name & "." & msg, DirectCast( _
                        MsgBoxStyle.OkOnly + _
                        MsgBoxStyle.Critical, MsgBoxStyle), _
                        My.Application.Info.Title)
        Catch ex As Exception
            Call MsgBox(Obj.Name & "." & "MsgBox_Exception() TRAP ERROR = " & ex.Message, _
                        DirectCast( _
                        MsgBoxStyle.OkOnly + _
                        MsgBoxStyle.Critical, MsgBoxStyle), _
                        My.Application.Info.Title)
        End Try

    End Sub
#End Region

    '#Region "指定ブロックのステージ位置を取得-(未使用)"
    '    '''=========================================================================
    '    '''<summary>指定ブロックの座標位置を取得する</summary>
    '    '''<param name="plateNo"> (INP) プレート番号</param>
    '    '''<param name="blockNo">(INP)プレート内のブロック番号</param>
    '    '''<param name="stgx">(OUT)ステージX座標</param>
    '    '''<param name="stgy">(OUT)ステージY座標</param>
    '    '''<returns>最終プレート、最終ブロックの場合TRUEを返す。</returns>
    '    '''<remarks></remarks>
    '    '''  プレートの並び　　プレート内部の並び
    '    '''  ____ ____ ____      ____ ____
    '    ''' | ⑨ | ⑧ | ① |　　| ⑧ | ① |
    '    ''' |____|____|____|    |____|____|
    '    ''' | ⑩ | ⑦ | ② |    | ⑦ | ② |    
    '    ''' |____|____|____|    |____|____|
    '    ''' | ⑪ | ⑥ | ③ |    | ⑥ | ③ |
    '    ''' |____|____|____|    |____|____|
    '    ''' | ⑫ | ⑤ | ④ |    | ⑤ | ④ |
    '    ''' |____|____|____|    |____|____|
    '    ''' 
    '    '''=========================================================================
    '    Public Function GetTargetStagePos(ByVal plateNo As Integer, ByVal blockNo As Integer, ByRef stgx As Double, ByRef stgy As Double) As Boolean
    '        Dim intPlateXCnt As Integer
    '        Dim intPlateYCnt As Integer
    '        Dim intLastPlateNo As Integer
    '        Dim dblWorkBaseStgPosX As Double
    '        Dim dblWorkBaseStgPosY As Double
    '        Dim strMSG As String

    '        GetTargetStagePos = False

    '        Try

    '            'プレートによるベースの位置座標を取得する。
    '            With typPlateInfo
    '                'プレート間隔の計算
    '                intPlateXCnt = plateNo / .intPlateCntXDir
    '                If (intPlateXCnt Mod 2) = 0 Then
    '                    intPlateYCnt = plateNo Mod .intPlateCntYDir
    '                Else
    '                    intPlateYCnt = (.intPlateCntYDir + 1) - (plateNo Mod .intPlateCntYDir)
    '                End If

    '                'プレートのベースポジション
    '                dblWorkBaseStgPosX = (.dblPlateSizeX * intPlateXCnt)
    '                dblWorkBaseStgPosY = (.dblPlateSizeY * intPlateYCnt)

    '                'ブロック座標を取得する
    '                stgx = dblWorkBaseStgPosX + typGrpInfoArray(blockNo).dblStgPosX
    '                stgy = dblWorkBaseStgPosY + typGrpInfoArray(blockNo).dblStgPosY

    '                '最終プレートの判定
    '                intLastPlateNo = .intPlateCntXDir * .intPlateCntYDir
    '                If (plateNo = intLastPlateNo) Then
    '                    GetTargetStagePos = True
    '                End If
    '            End With

    '            ' トラップエラー発生時 
    '        Catch ex As Exception
    '            strMSG = "DataAccess.GetTargetStagePos() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '            GetTargetStagePos = 0
    '        End Try

    '    End Function
    '#End Region

#Region "開始ログ表示処理"
    '''=========================================================================
    ''' <summary>開始ログ表示</summary>
    ''' <param name="plateNoX"></param>
    ''' <param name="plateNoY"></param>
    ''' <param name="StageGrpNoX"></param>
    ''' <param name="StageGrpNoY"></param>
    ''' <param name="blockNoX"></param>
    ''' <param name="blockNoY"></param>
    ''' <remarks>処理開始時の実行箇所の情報を表示する。</remarks>
    '''=========================================================================
    Public Sub DisplayStartLog(ByVal plateNoX As Integer, ByVal plateNoY As Integer, _
                ByVal StageGrpNoX As Integer, ByVal StageGrpNoY As Integer, _
                ByVal blockNoX As Integer, ByVal blockNoY As Integer)
        'Dim bDispLogWrite As Boolean
        'Dim strLOG As String
        Dim strMsg As String
        'Dim digL As Integer
        'Dim digH As Integer
        'Dim digSW As Integer


        Try
            'bDispLogWrite = False

            '' ﾃﾞｼﾞSWの2桁目をﾁｪｯｸする。
            'Call Form1.GetMoveMode(digL, digH, digSW)
            'Select Case digH
            '    'Select Case gDigH
            '    'Case 0, 1
            '    Case 0                                          '###217
            '        'Case 0, 1
            '        '    ' ﾃﾞｼﾞSWの1桁目をﾁｪｯｸする。
            '        '    Select Case digL
            '        '        'Select Case gDigL
            '        '        Case 0, 1, 2
            '        '            ' 表示しない
            '        '            Exit Sub
            '        '        Case Else
            '        '            ' ログ画面に文字列を表示する
            '        '            bDispLogWrite = True
            '        '    End Select
            '    Case Else
            '        ' ログ画面に文字列を表示する
            '        bDispLogWrite = True
            'End Select

            '' ログ出力文字列の構築と出力
            'If (bDispLogWrite = True) Then
            '    strLOG = "--- Plate X=" + plateNoX.ToString("000") + " Y=" + plateNoY.ToString("000")
            '    strLOG = strLOG + "--- StageGroup X= " & StageGrpNoX.ToString("000") + " Y= " & StageGrpNoY.ToString("000")
            '    strLOG = strLOG + " Block X=" + blockNoX.ToString("000") + " Y=" + blockNoY.ToString("000")

            '    ''''処理の検討が必要。
            '    'If StepDir = 1 Then
            '    '    strLOG = strLOG & " Block X=" & (LogBlkData + XY(1) + 1).ToString("000") & " Y=" & (XY(2) + 1).ToString("000") '(TXT仕様変更)
            '    'Else
            '    '    strLOG = strLOG & " Block X=" & (XY(1) + 1).ToString("000") & " Y=" & (LogBlkData + XY(2) + 1).ToString("000") '(TXT仕様変更)
            '    'End If
            '    'strLOG = strLOG & vbCrLf
            '    strLOG = strLOG

            '    ' ログ画面に文字列を表示する
            '    Call Form1.Z_PRINT(strLOG)
            'End If

            ' トラップエラー発生時 
        Catch ex As Exception
            strMsg = "basTrimming.DisplayStartLog() TRAP ERROR = " + ex.Message
            MsgBox(strMsg)
        End Try
    End Sub
#End Region
    ' ###1040の修正追加END

#Region "レーザアラーム８３３エラー時の処理"
    ''' <summary>
    ''' レーザアラーム８３３エラー時のプログラム終了処理
    ''' </summary>
    ''' <param name="iRtn"></param>
    ''' <remarks></remarks>
    Public Sub Check_ERR_LSR_STATUS_STANBY(ByRef iRtn As Integer)
        Try
            If (iRtn = ERR_LSR_STATUS_STANBY OrElse iRtn = ERR_LSR_STATUS_OSCERR) Then           ' 833:LASER IS NOT READY ' 850:Error occured,:ES:LD Alarm  V2.1.0.3①
                iRtn = ObjSys.EX_ZGETSRVSIGNAL(gSysPrm, iRtn, 0)
                Call V_Off()                                                      ' DC電源装置 電圧OFF    V2.1.0.3①
                Call Form1.AppEndDataSave()                                       ' ｿﾌﾄ強制終了時のﾃﾞｰﾀ保存確認
                Call Form1.AplicationForcedEnding()                               ' ｿﾌﾄ強制終了処理
            End If
        Catch ex As Exception
            MsgBox("DisplayDistribute() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Sub

#End Region


#Region "カットトレースのためのカット方向を設定する"
    '''=========================================================================
    '''<summary>カットトレースのためのカット方向を設定する</summary>
    '''<param name="rn">   (INP)抵抗番号</param>
    '''<param name="cn">   (INP)カット番号</param>
    '''<param name="dirL1">(I/O)カット方向1</param>
    '''<param name="dirL2">(I/O)カット方向2</param>
    '''<remarks>・STｶｯﾄ/IDXｶｯﾄ時(dirL2は返さない)
    '''           入力        = カット角度(0°, 90°, 180°, 270°)
    '''           出力(dirL1) = カット方向(3:+X(0°), 2:+Y(90°), 1:-X(180°), 4:-Y(270°))
    '''         ・L ｶｯﾄ/ HOOK ｶｯﾄ時(dirL2は返さない)
    '''           入力        = カット角度(0°, 90°, 180°, 270°)
    '''                       = Lﾀｰﾝ方向(1:CW, 2:CCW)
    '''           出力(dirL1) = カット方向 3:+X+Y(→↑), 4:-Y+X(↓→), 1:-X-Y(↓←) ,2:+Y-X(←↑),
    '''                                    7:+X-Y(→↓), 8:-Y-X(←↓), 5:-X+Y(↑←) ,6:+Y+X(↑→))
    '''         ・スキャンカット時
    '''           入力        = カット角度(0°, 90°, 180°, 270°)
    '''                       = ｽﾃｯﾌﾟ方向(0:0°, 1:90°, 2:180°, 3:270)
    '''           出力(dirL1) = カット方向(1:-X, 2:+X, 3:-Y, 4:+Y)
    '''           出力(dirL2) = ｽﾃｯﾌﾟ方向 (1:+X, 2:-X, 3:+Y, 4:-Y)
    ''' </remarks>
    '''=========================================================================
    Private Sub Cnv_Cut_Dir(ByRef rn As Short, ByRef cn As Short, ByRef dirL1 As Short, ByRef dirL2 As Short)

        Dim strMSG As String                                ' メッセージ編集域

        Try

            ' Lｶｯﾄ/HOOKｶｯﾄ/Uｶｯﾄ時
            If (stREG(rn).STCUT(cn).intCTYP = CNS_CUTP_L) Or
               (stREG(rn).STCUT(cn).intCTYP = CNS_CUTP_U) Then
                Select Case (stREG(rn).STCUT(cn).intUCutANG)
                    Case 0                                          ' カット方向(1:+X(0°)) ?
                        If (stREG(rn).STCUT(cn).intUCutTurnDir = 1) Then
                            dirL1 = 7                               ' +X-Y(→↓)
                        Else
                            dirL1 = 3                               ' +X+Y(→↑)
                        End If
                    Case 90                                         ' カット方向(2:+Y(90°)) ?
                        If (stREG(rn).STCUT(cn).intUCutTurnDir = 1) Then
                            dirL1 = 6                               ' +Y+X(↑→))
                        Else
                            dirL1 = 2                               ' +Y-X(↑←)
                        End If
                    Case 180                                        ' カット方向(3:-X(180°)) ?
                        If (stREG(rn).STCUT(cn).intUCutTurnDir = 1) Then
                            dirL1 = 5                               ' -X+Y(←↑)
                        Else
                            dirL1 = 1                               ' -X-Y(←↓)
                        End If
                    Case Else                                       ' カット方向(4:-Y(270°)) ?
                        If (stREG(rn).STCUT(cn).intUCutTurnDir = 1) Then
                            dirL1 = 8                               ' -Y-X(↓←)
                        Else
                            dirL1 = 4                               ' -Y+X(↓→)
                        End If
                End Select
            End If

            Exit Sub

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "i-TKY.Cnv_Cut_Dir() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try

    End Sub
#End Region


#Region "Ｕカット－インデックスカットトリミング x0(ﾄﾘﾐﾝｸﾞﾓｰﾄﾞ)"
    '''=========================================================================
    '''<summary>Ｕカット－インデックスカットトリミング x0(ﾄﾘﾐﾝｸﾞﾓｰﾄﾞ)</summary>
    '''<param name="rn"> (INP) 抵抗データIndex　(1 org)</param>
    '''<param name="cn"> (INP) カット番号　　　 (1 org)</param>
    '''<param name="Mx"> (INP) 初期測定値</param>
    '''<returns> 0 = 正常
    '''          1 = 正常(目標値を超えたので終了)
    '''          2 = 指定移動量までカットしたので終了
    '''          5 = TRV目標値を超えたので終了
    '''          上記以外 = その他エラー
    ''' </returns>
    '''=========================================================================
    Public Function TRM_IX_U(ByRef rn As Short, ByRef cn As Short, ByRef Mx As Double) As Short

        Dim i As Short                                  ' ﾙｰﾌﾟ回数
        Dim j As Short                                  ' ﾙｰﾌﾟ回数
        Dim IDX As Short                                ' ｲﾝﾃﾞｯｸｽｶｯﾄ1～3(ﾋﾟｯﾁ大,中,小)
        Dim count As Short                              ' ｲﾝﾃﾞｯｸｽｶｯﾄ数
        Dim r As Short                                  ' 関数戻値
        Dim CutL As Double                              ' 最大ｶｯﾄ長
        Dim ln As Double                                ' 現在のｶｯﾄ長
        Dim NOM(MAX_LCUT) As Double                     ' 目標値(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後) V1.0.4.3⑦２をMAX_LCUTへ変更
        Dim NOMx As Double                              ' 目標値
        Dim VX(3) As Double                             ' 作業域
        Dim strMSG As String                            ' メッセージ編集域
        Dim wkL1 As Double                              ' 作業域
        Dim wkL2 As Double                              ' 作業域
        Dim dblMx As Double                             ' 作業域
        Dim dblQrate As Double
        Dim shSLP As Short
        Dim CutLSum As Double                           ' V1.0.4.3⑧ 最大ｶｯﾄ長積算値
        Dim dblQRateL(MAX_LCUT) As Double               ' V1.0.4.3⑦　Ｑレート
        Dim dblSpeedL(MAX_LCUT) As Double               ' V1.0.4.3⑦　速度
        Dim dblSpeed As Double                          ' V1.0.4.3⑦　速度
        Dim SaveIDX As Short                            'V2.1.0.0⑤
        Dim ExecUCut As Boolean                         ' Uカットを実行するかどうかの指定 
        Dim L1Length As Double                          ' L3のために、L1のカット長を保存 
        Dim L3Angle As Double                           ' L3カット角度
        Dim L2Angle As Double                           ' R2用角度 

        Try

            '---------------------------------------------------------------------------
            '   初期処理
            '---------------------------------------------------------------------------
            CutLSum = 0.0                                   ' V1.0.4.3⑧
            L1Length = 0.0                                  ' L1カット長
            TRM_IX_U = cFRS_NORMAL                          ' Return値 = 正常
            strMSG = ""
            LTFlg = 1                                       ' Lﾀｰﾝﾌﾗｸﾞ(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
            LTAng(1) = stREG(rn).STCUT(cn).intANG           ' ANG(1) = Lﾀｰﾝ前のｶｯﾄ方向
            LTAng(2) = stREG(rn).STCUT(cn).intANG2          ' ANG(2) = Lﾀｰﾝ後のｶｯﾄ方向
            dblML(1) = stREG(rn).STCUT(cn).dblDL2           ' ﾘﾐｯﾄｶｯﾄ量mm(1:Lﾀｰﾝ前)
            dblML(2) = stREG(rn).STCUT(cn).dblDL3           ' ﾘﾐｯﾄｶｯﾄ量mm(2:Lﾀｰﾝ後)
            LTP = stREG(rn).STCUT(cn).dblLTP                ' Lﾀｰﾝﾎﾟｲﾝﾄ(%)
            NOM(1) = stREG(rn).dblNOM                       ' 目標値
            ' ｶｯﾄｵﾌ(%)→目標値に対するｵﾌｾｯﾄ値(目標値×(1＋ｶｯﾄｵﾌ/100))
            NOMx = Func_CalNomForCutOff(rn, cn, NOM(1))          'カットオフによる目標値
            NOM(1) = NOMx                                   ' Lﾀｰﾝ前目標値 = 目標値
            NOM(2) = NOMx                                   ' Lﾀｰﾝ後目標値 = 目標値

            ' Lﾀｰﾝ前ﾘﾐｯﾄｶｯﾄ量mmと目標値を設定する
            CutL = dblML(LTFlg)                             ' ﾘﾐｯﾄｶｯﾄ量mm(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
            NOMx = NOM(LTFlg)                               ' 目標値(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)

            '-----------------------------------------------------------------------
            '   ユーザプログラム特殊処理 カット毎の目標値を求める。START
            '-----------------------------------------------------------------------
            If UserSub.IsSpecialTrimType Then
                NOMx = UserSub.GetTargeResistancetValue(rn, cn)
                If UserSub.IsTrimType2() Or UserSub.IsTrimType3() Then    'V1.0.4.3④IsTrimType3()追加
                    ' G15A-15A.BAS : 14580       IF TRM1#(CN1%)<=.5# THEN GOTO *NEXT.CT1
                    '###1032                    If NOMx <= 0.5# Then
                    If NOMx <= 0.005# Then
                        Return (cFRS_NORMAL)
                    End If
                End If
            End If

            '-----------------------------------------------------------------------
            '   ユーザプログラム特殊処理 カット毎の目標値を求める。END
            '-----------------------------------------------------------------------

            ' Qレート設定
            If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then
                dblQrate = stREG(rn).STCUT(cn).intQF1
                dblQrate = dblQrate / 10.0
            Else
#If cOSCILLATORcFLcUSE Then
                ' FL時は加工条件番号テーブルからQレートを設定する(カットスピードはデータから設定)
                IDX = stREG(rn).STCUT(cn).intCND(CUT_CND_L1)
                dblQrate = stCND.Freq(IDX)
#End If
            End If

            '-----------------------------------------------------------------------
            '   Ｕカット
            '-----------------------------------------------------------------------
            If (stREG(rn).STCUT(cn).intCTYP = CNS_CUTP_U) Then

                NOM(1) = Mx + (NOMx - Mx) * (stREG(rn).STCUT(cn).dblUCutTurnP * 0.01)  ' ﾀｰﾝ前目標値設定(初期値＋(目標値-初期値)×ﾀｰﾝﾎﾟｲﾝﾄ/100)
                NOM(2) = NOMx
                'dblML(1) = stREG(rn).STCUT(cn).dUCutL1
                'dblML(2) = stREG(rn).STCUT(cn).dUCutL2
                dblML(1) = stREG(rn).STCUT(cn).dUCutL1 - stREG(rn).STCUT(cn).dblUCutR1
                dblML(2) = stREG(rn).STCUT(cn).dUCutL2 - stREG(rn).STCUT(cn).dblUCutR1 - stREG(rn).STCUT(cn).dblUCutR2
                LTAng(1) = stREG(rn).STCUT(cn).intUCutANG
                L3Angle = stREG(rn).STCUT(cn).intUCutANG - 180
                If L3Angle < 0 Then
                    L3Angle = L3Angle + 360
                End If
                If L3Angle >= 360 Then
                    L3Angle = L3Angle - 360
                End If

                If stREG(rn).STCUT(cn).intUCutTurnDir = 1 Then
                    ' CW
                    LTAng(2) = stREG(rn).STCUT(cn).intUCutANG - 90
                Else
                    ' CCW
                    LTAng(2) = stREG(rn).STCUT(cn).intUCutANG + 90
                End If
                If LTAng(2) < 0 Then
                    LTAng(2) = LTAng(2) + 360
                End If
                If LTAng(2) >= 360 Then
                    LTAng(2) = LTAng(2) - 360
                End If
                dblQRateL(1) = stREG(rn).STCUT(cn).intUCutQF1
                dblSpeedL(1) = stREG(rn).STCUT(cn).dblUCutV1

                NOM(MAX_LCUT) = NOMx
                CutL = dblML(1)                             ' ﾘﾐｯﾄｶｯﾄ量mm(Ｌ１カット)
                NOMx = NOM(1)                               ' 目標値(Ｌ１カット)
                dblQrate = dblQRateL(1) / 10.0
                dblSpeed = dblSpeedL(1)
            Else
                dblSpeed = stREG(rn).STCUT(cn).dblV1
            End If
            ' V1.0.4.3⑦↑

            ' ｶｯﾄ量初期化
            For i = 1 To MAXCTN                             ' MAXカット数分繰返す
                For j = 1 To MAX_LCUT                           ' MAXカット数分繰返す
                    dblLN(j, i) = 0.0#                          ' ｶｯﾄ量初期化(1:Lﾀｰﾝ前)
                    'V1.0.4.3⑦                    dblLN(1, i) = 0.0#                          ' ｶｯﾄ量初期化(1:Lﾀｰﾝ前)
                    'V1.0.4.3⑦                    dblLN(2, i) = 0.0#                          ' ｶｯﾄ量初期化(2:Lﾀｰﾝ後)
                Next j
            Next i

            '---------------------------------------------------------------------------
            '   ｲﾝﾃﾞｯｸｽｶｯﾄでLｶｯﾄ/STｶｯﾄを行う
            '---------------------------------------------------------------------------
            For IDX = 1 To MAXIDX                           ' ｲﾝﾃﾞｯｸｽｶｯﾄ1～5(ﾋﾟｯﾁ大,中,小)分繰返す
STP_CHG_PIT:
                count = stREG(rn).STCUT(cn).intIXN(IDX)     ' count = ｲﾝﾃﾞｯｸｽｶｯﾄ数
                ln = stREG(rn).STCUT(cn).dblDL1(IDX)        ' カット長を設定する(ﾋﾟｯﾁ大,中,小)
                For i = 1 To count                          ' ｲﾝﾃﾞｯｸｽｶｯﾄ数分繰返す
#If (cCND = 1) Then ' 条件出しﾓｰﾄﾞ ?
                    strMSG = "　 抵抗番号=" + Format(rn, "0") + ",ｶｯﾄ番号=" + Format(cn, "0")
                    strMSG = strMSG + ",目標値(Lﾀｰﾝ前,後)=" + Format(NOM(1), "0.0####") + "," + Format(NOM(2), "0.0####") + vbCrLf
                    strMSG = strMSG + "　   ｶｯﾄ長=" + Format(ln, "#0.0####") + ",目標値=" + Format(NOMx, "0.0####") + ",LTFlg=" + Format(LTFlg, "0") + ",ｶｯﾄ量(Lﾀｰﾝ前,後)=" + Format(dblLN(1, cn), "#0.0####") + "," + Format(dblLN(2, cn), "#0.0####")
                    Call Z_PRINT(strMSG + vbCrLf)
#End If
                    ExecUCut = False

                    'V2.1.0.0⑤↓
                    If IsCutVariationJudgeExecute() AndAlso UserSub.IsCutMeasureBefore() Then
                        dblMx = UserSub.CutVariationMeasureBeforeGet()
                    Else
                        'V2.1.0.0⑤↑

                        Call UserSub.ChangeMeasureSpeed(rn, cn, IDX)     ' 測定速度の変更（特注処理）

                        Call DScanModeSet(rn, cn, IDX)              ' ＤＣスキャナに接続する測定器を切り替えてスキャナをセットする。

                        ' 電圧(外部/内部)/抵抗測定(外部/内部)を行う
                        ' 測定レンジの目標値を最終目標値にする。2013.3.28                    r = V_R_MEAS(stREG(rn).intSLP, stREG(rn).STCUT(cn).intIXMType(IDX), dblMx, rn, NOMx)
                        If UserSub.IsSpecialTrimType Then
                            r = V_R_MEAS(stREG(rn).intSLP, stREG(rn).STCUT(cn).intIXMType(IDX), dblMx, rn, UserSub.GetTRV())
                        Else
                            r = V_R_MEAS(stREG(rn).intSLP, stREG(rn).STCUT(cn).intIXMType(IDX), dblMx, rn, NOMx)
                        End If
                        If (r <> cFRS_NORMAL) Then                          ' エラー
                            Call UserSub.ResoreMeasureSpeed(rn, cn, IDX)         ' 測定速度の変更を元に戻す（特注処理）
                            Return (cFRS_TRIM_NG)                           ' Return値 = トリミングNG
                        End If

                        Call UserSub.ResoreMeasureSpeed(rn, cn, IDX)     ' 測定速度の変更を元に戻す（特注処理）
                    End If       'V2.1.0.0⑤
                    SaveIDX = IDX                                       'V2.1.0.0⑤

                    If bDebugLogOut Then
                        DebugLogOut("抵抗[" & rn.ToString & "]カット[" & cn.ToString & "] IX測定値=[" & dblMx.ToString & "]")
                    End If

                    If UserSub.IsSpecialTrimType And dblMx >= UserSub.GetTRV() Then
                        If bDebugLogOut Then
                            DebugLogOut("TRV目標値到達 抵抗[" & rn.ToString & "]カット[" & cn.ToString & "] IX測定値=[" & dblMx.ToString & "]  TRV=[" & UserSub.GetTRV().ToString & "]")
                        End If
                        UserSub.CutVariationMeasureAfterSet(dblMx) 'V2.1.0.0①カット後の測定値保存
                        Return (5)
                    End If

                    MoveStop()              'V2.2.0.0⑥ 

#If (cCND = 1) Then ' 条件出しﾓｰﾄﾞ ?
                    strMSG = "■測定値=" + dblMx.ToString("#0.0###")
                    Call Z_PRINT(strMSG + vbCrLf)
#End If
                    ' 目標値を超えたか調べる
                    If (stREG(rn).intSLP = SLP_VTRIMPLS) Or (stREG(rn).intSLP = SLP_RTRM) Then ' +ｽﾛｰﾌﾟ/抵抗 ?
                        If (dblMx >= NOMx) Then             ' 測定値 >= 目標値なら次へ
                            TRM_IX_U = 1                      ' Return値 = 1(目標値を超えたので終了)
                            '-------------------------
                            '   Lﾀｰﾝ処理(Lｶｯﾄ時)
                            '-------------------------
                            ' Lｶｯﾄ以外ならEXIT
                            If (stREG(rn).STCUT(cn).intCTYP <> CNS_CUTP_U) Then
                                GoTo TRM_IX_EXIT
                            End If
                            If (LTFlg >= 2) Then
                                GoTo TRM_IX_EXIT                            ' Lﾀｰﾝ後ならEXIT
                            End If
                            TRM_IX_U = 0                                    ' Return値 = 正常
                            LTFlg = LTFlg + 1                               ' ﾀｰﾝﾌﾗｸﾞ = 次のカット(ﾀｰﾝ後)
                            CutL = dblML(LTFlg)                             ' ﾘﾐｯﾄｶｯﾄ量mm(1:ﾀｰﾝ前,2:ﾀｰﾝ後)
                            NOMx = NOM(LTFlg)                               ' 目標値(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
                            dblQrate = dblQRateL(LTFlg) / 10.0
                            dblSpeed = dblSpeedL(LTFlg)
                            ln = stREG(rn).STCUT(cn).dblDL1(IDX)            ' カット長を設定する(ﾋﾟｯﾁ大,中,小)
                            ExecUCut = True
                            GoTo TRM_R1_CUT                                 ' R1ｶｯﾄを行う
                            '-------------------------
                            '   Lﾀｰﾝ処理(Lｶｯﾄ時)
                            '-------------------------


                        End If
                    Else                                    ' -ｽﾛｰﾌﾟ ?
                        If (dblMx <= NOMx) Then             ' 測定値 <= 目標値なら次へ
                            TRM_IX_U = 1                      ' Return値 = 1(目標値を超えたので終了)
                            '-------------------------
                            '   Lﾀｰﾝ処理(Lｶｯﾄ時)
                            '-------------------------
                            ' Lｶｯﾄ以外ならEXIT
                            If (stREG(rn).STCUT(cn).intCTYP <> CNS_CUTP_U) Then
                                GoTo TRM_IX_EXIT
                            End If
                            If (LTFlg >= 2) Then
                                ExecUCut = True
                                GoTo TRM_IX_EXIT                            ' Lﾀｰﾝ後ならEXIT
                            End If
                            TRM_IX_U = 0                                    ' Return値 = 正常
                            LTFlg = LTFlg + 1                               ' Lﾀｰﾝﾌﾗｸﾞ = 次のカット(Lﾀｰﾝ後)
                            CutL = dblML(LTFlg)                             ' ﾘﾐｯﾄｶｯﾄ量mm(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
                            NOMx = NOM(LTFlg)                               ' 目標値(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
                            dblQrate = dblQRateL(LTFlg) / 10.0
                            dblSpeed = dblSpeedL(LTFlg)
                            ln = stREG(rn).STCUT(cn).dblDL1(IDX)            ' カット長を設定する(ﾋﾟｯﾁ大,中,小)
                            ExecUCut = True
                            GoTo TRM_R1_CUT                                 ' R1ｶｯﾄを行う
                            '-------------------------
                            '   Lﾀｰﾝ処理(Lｶｯﾄ時)
                            '-------------------------

                        End If
                    End If

                    ' ｲﾝﾃﾞｯｸｽｶｯﾄﾋﾟｯﾁ(1:大,2-4:中,5:小)をチェックする
                    r = Get_Idx_Pitch(rn, cn, IDX, NOMx, dblMx) ' 目標値との誤差によりﾋﾟｯﾁを変更する
                    If (r <> IDX) Then                      ' ｶｯﾄﾋﾟｯﾁ変更 ?
                        IDX = r                             ' ｲﾝﾃﾞｯｸｽｶｯﾄﾋﾟｯﾁを変更する
                        GoTo STP_CHG_PIT
                    End If

                    ' 次のｶｯﾄで最大ｶｯﾄ量を超える ? (※下記のようにしないと正しい比較ができない)
                    wkL1 = CDbl((dblLN(LTFlg, cn) + ln).ToString("#0.0000"))
                    wkL2 = CDbl(CutL.ToString("#0.0000"))
                    If (wkL1 > wkL2) Then                   ' 最大ｶｯﾄ量を超える ?
                        ' ln = 残りのｶｯﾄ量(下記のようにしないとln=0とならない場合あり)
                        ln = CDbl((wkL2 - dblLN(LTFlg, cn)).ToString("#0.0000"))
                        If (ln <= 0) Then                   ' 最大ｶｯﾄ量までカット ?
                            TRM_IX_U = 2                      ' Return値 = 2(指定移動量までカット)
                            '-------------------------
                            '   Lﾀｰﾝ処理(Lｶｯﾄ時)
                            '-------------------------
                            ' Lｶｯﾄ以外ならEXIT
                            ' Lｶｯﾄ以外ならEXIT
                            If (stREG(rn).STCUT(cn).intCTYP <> CNS_CUTP_U) Then
                                GoTo TRM_IX_EXIT
                            End If
                            If (LTFlg >= 2) Then
                                ExecUCut = True
                                GoTo TRM_IX_EXIT                            ' Lﾀｰﾝ後ならEXIT
                            End If
                            TRM_IX_U = 0                                      ' Return値 = 正常
                            LTFlg = LTFlg + 1                               ' Lﾀｰﾝﾌﾗｸﾞ = 次のカット(Lﾀｰﾝ後)
                            CutL = dblML(LTFlg)                             ' ﾘﾐｯﾄｶｯﾄ量mm(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
                            NOMx = NOM(LTFlg)                               ' 目標値(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
                            ln = stREG(rn).STCUT(cn).dblDL1(IDX)            ' カット長を設定する(ﾋﾟｯﾁ大,中,小)
                            ExecUCut = True
                            GoTo TRM_R1_CUT                                 ' R1ｶｯﾄを行う
                            '-------------------------
                            '   Lﾀｰﾝ処理(Lｶｯﾄ時)
                            '-------------------------


                        End If
                    End If

                    ' 斜め直線ｶｯﾄ(ﾎﾟｼﾞｼｮﾆﾝｸﾞなし)
                    If (stREG(rn).intSLP = SLP_ATRIMPLS) Then
                        shSLP = 1
                    ElseIf (stREG(rn).intSLP = SLP_ATRIMMNS) Then
                        shSLP = 2
                    Else
                        shSLP = stREG(rn).intSLP
                    End If
                    r = TrimSt(FORCE_MODE, 0, 0, shSLP, LTAng(LTFlg), ln, dblSpeed, dblSpeed,
                                  dblQrate, dblQrate, stREG(rn).STCUT(cn).intCND(CUT_CND_L1), stREG(rn).STCUT(cn).intCND(CUT_CND_L1))
                    'r = TrimSt(FORCE_MODE, 0, 0, shSLP, LTAng(LTFlg), ln, stREG(rn).STCUT(cn).dblV1, stREG(rn).STCUT(cn).dblV1, _
                    '              dblQrate, dblQrate, stREG(rn).STCUT(cn).intCND(CUT_CND_L1), stREG(rn).STCUT(cn).intCND(CUT_CND_L1))
                    Call ZWAIT(stREG(rn).STCUT(cn).lngPAU(IDX)) ' ｲﾝﾃﾞｯｸｽﾋﾟｯﾁ間ﾎﾟｰｽﾞ(ms)
                    If (r <> 0) And (r <> 2) Then
                        Z_PRINT("CUT_RETRACE ERROR RETURN =[" & r.ToString & "] RES=[" & rn.ToString & "]CUT=[" & cn.ToString & "]")
                        Return (cFRS_TRIM_NG)                           ' Return値 = トリミングNG
                    End If
                    UserSub.CutVariationCutSet()                        'V2.1.0.0⑤カットが有った事を記録する。
                    dblLN(LTFlg, cn) = dblLN(LTFlg, cn) + ln    ' ｶｯﾄ済量mmを退避(1:Lﾀｰﾝ前,2:Lﾀｰﾝ後)
                    CutLSum = CutLSum + ln                      'V1.0.4.3⑧ 積算値


TRM_R1_CUT:
                    ' R1カットの実行 
                    If ExecUCut = True Then
                        L1Length = dblLN(LTFlg - 1, cn)             'L1のカット済長保存 
                        If stREG(rn).STCUT(cn).dblUCutR1 > 0 Then
                            r = TRM_Circle(rn, cn, stREG(rn).STCUT(cn).dblUCutR1, stREG(rn).STCUT(cn).intUCutANG)
                            If (r <> 0) And (r <> 2) Then
                                Z_PRINT("CUT_RETRACE ERROR RETURN =[" & r.ToString & "] RES=[" & rn.ToString & "]CUT=[" & cn.ToString & "]")
                                Return (cFRS_TRIM_NG)                           ' Return値 = トリミングNG
                            End If
                        End If
                    End If

TRM_IX_NEXT:
                Next i                                      ' 次ｶｯﾄへ

                If count > 0 Then                   'V2.1.0.0⑤ タクトアップの為、カット数が０でもIDX回行う為
                    ' セーフティチェック
                    r = SafetyCheck()                           ' セーフティチェック
                    If (r <> 0) Then                            ' エラー ?
                        TRM_IX_U = r                              ' Return値 = セーフティチェックエラー
                        Exit Function
                    End If
                End If                              'V2.1.0.0⑤ 
            Next IDX                                        ' 次ﾋﾟｯﾁへ

TRM_IX_EXIT:

            ' ここにきて、フラグが立っていない場合は、L1のトリミングで距離リミットに到達していなくて、切りあがっていない場合：R1カットの実行 
            If ExecUCut = False Then
                L1Length = dblLN(1, cn)             'L1のカット済長保存 
                ExecUCut = True
                If stREG(rn).STCUT(cn).dblUCutR1 > 0 Then
                    r = TRM_Circle(rn, cn, stREG(rn).STCUT(cn).dblUCutR1, stREG(rn).STCUT(cn).intUCutANG)
                    If (r <> 0) And (r <> 2) Then
                        Z_PRINT("CUT_RETRACE ERROR RETURN =[" & r.ToString & "] RES=[" & rn.ToString & "]CUT=[" & cn.ToString & "]")
                        Return (cFRS_TRIM_NG)                           ' Return値 = トリミングNG
                    End If
                End If
            End If


            ' R2 + L3 カットの実行 
            If ExecUCut = True Then
                MoveStop()              'V2.2.0.0⑥ 

                ' R2カットの実行
                If stREG(rn).STCUT(cn).dblUCutR2 > 0 Then

                    L2Angle = LTAng(2)
                    r = TRM_Circle(rn, cn, stREG(rn).STCUT(cn).dblUCutR2, L2Angle)
                    If (r <> 0) And (r <> 2) Then
                        Z_PRINT("CUT_RETRACE ERROR RETURN =[" & r.ToString & "] RES=[" & rn.ToString & "]CUT=[" & cn.ToString & "]")
                        Return (cFRS_TRIM_NG)                           ' Return値 = トリミングNG
                    End If

                    MoveStop()              'V2.2.0.0⑥ 

                End If
                L1Length = L1Length + (stREG(rn).STCUT(cn).dblUCutR1 - stREG(rn).STCUT(cn).dblUCutR2)
                ' L3カットの実行
                r = TrimSt(FORCE_MODE, 0, 0, shSLP, L3Angle, L1Length, dblSpeed, dblSpeed,
                                  dblQrate, dblQrate, stREG(rn).STCUT(cn).intCND(CUT_CND_L1), stREG(rn).STCUT(cn).intCND(CUT_CND_L1))
            End If

            'V2.1.0.0⑤↓測定しないで抜けるパターン有り
            If IsCutVariationJudgeExecute() AndAlso UserSub.IsNotCutMeasureAfter() = True Then
                Call UserSub.ChangeMeasureSpeed(rn, cn, SaveIDX)     ' 測定速度の変更（特注処理）

                Call DScanModeSet(rn, cn, SaveIDX)              ' ＤＣスキャナに接続する測定器を切り替えてスキャナをセットする。

                ' 電圧(外部/内部)/抵抗測定(外部/内部)を行う
                If UserSub.IsSpecialTrimType Then
                    r = V_R_MEAS(stREG(rn).intSLP, stREG(rn).STCUT(cn).intIXMType(SaveIDX), dblMx, rn, UserSub.GetTRV())
                Else
                    r = V_R_MEAS(stREG(rn).intSLP, stREG(rn).STCUT(cn).intIXMType(SaveIDX), dblMx, rn, NOMx)
                End If
                If (r <> cFRS_NORMAL) Then                          ' エラー
                    Call UserSub.ResoreMeasureSpeed(rn, cn, SaveIDX)         ' 測定速度の変更を元に戻す（特注処理）
                    Call Z_PRINT("インデックスカット追加測定時エラー抵抗[" & rn.ToString & "]カット[" & cn.ToString & "]" & vbCrLf)
                    CutVariationDebugLogOut("インデックスカット追加測定時エラー 抵抗[" & rn.ToString & "]カット[" & cn.ToString & "]IX[" & SaveIDX.ToString & "]")
                    Return (cFRS_TRIM_NG)                           ' Return値 = トリミングNG
                End If

                Call UserSub.ResoreMeasureSpeed(rn, cn, SaveIDX)     ' 測定速度の変更を元に戻す（特注処理）

                CutVariationDebugLogOut("抵抗[" & rn.ToString & "]カット[" & cn.ToString & "]IX[" & SaveIDX.ToString & "] IX測定値=[" & dblMx.ToString & "]")

                UserSub.CutVariationMeasureAfterSet(dblMx)
            End If
            'V2.1.0.0⑤↑

            Exit Function

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "User.TRM_IX() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cFRS_TRIM_NG)                           ' Return値 = トリミングNG
        End Try
    End Function
#End Region

#Region "斜めＬカット電圧/抵抗トリミングを使用してＲカットをする"
    '''=========================================================================
    '''<summary>斜めＬカット電圧/抵抗トリミングを使用してＲカット</summary>
    '''<param name="rn"> (INP) 抵抗データIndex　(1 org)</param>
    '''<param name="cn"> (INP) カットデータIndex(1 org)</param>
    '''<param name="Radius">(INP) 半径</param>
    '''<param name="Angle">(INP) 角度</param>
    '''<returns> 0 = 正常
    '''          1 = 目標値を超えたので終了
    '''          2 = Lターン後の指定移動量までカットしたので終了
    '''         99 = その他エラー
    ''' </returns>
    '''=========================================================================
    Public Function TRM_Circle(ByRef rn As Short, ByRef cn As Short, ByRef Radius As Double, ByVal Angle As Double) As Integer

        ' Dim iDir As Short                               ' Lﾀｰﾝ後移動方向(1:時計方向,2:反時計方向)
        Dim CutLen(2) As Double
        Dim SpdOwd(2) As Double
        Dim SpdRet(2) As Double
        Dim QRateOwd(2) As Double
        Dim QRateRet(2) As Double
        Dim CondOwd(2) As Short
        Dim CondRet(2) As Short
        Dim dblQrate As Double

        CutLen(0) = Radius
        CutLen(1) = Radius

        SpdOwd(0) = stREG(rn).STCUT(cn).dblUCutV1
        SpdOwd(1) = stREG(rn).STCUT(cn).dblUCutV1
        SpdRet(0) = stREG(rn).STCUT(cn).dblUCutV1
        SpdRet(1) = stREG(rn).STCUT(cn).dblUCutV1

        dblQrate = stREG(rn).STCUT(cn).intUCutQF1
        dblQrate = dblQrate / 10.0

        QRateOwd(0) = dblQrate
        QRateOwd(1) = dblQrate
        QRateRet(0) = dblQrate
        QRateRet(1) = dblQrate

        CondOwd(0) = stREG(rn).STCUT(cn).intCND(1)
        CondOwd(1) = stREG(rn).STCUT(cn).intCND(2)

        CondRet(0) = stREG(rn).STCUT(cn).intCND(3)
        CondRet(1) = stREG(rn).STCUT(cn).intCND(4)

        ' Lﾀｰﾝ後移動方向(1:時計方向,2:反時計方向)を求める
        ' iDir = Get_Cut_Dir(stREG(rn).STCUT(cn).intANG, stREG(rn).STCUT(cn).intANG2)

        TRM_Circle = TrimL(FORCE_MODE, CUT_MODE_NORMAL, 0, SLP_RTRM, stREG(rn).STCUT(cn).intTMM, Angle, stREG(rn).STCUT(cn).dblLTP, stREG(rn).STCUT(cn).intUCutTurnDir,
                    CutLen, SpdOwd, SpdRet, QRateOwd, QRateRet, CondOwd, CondRet, Radius)


    End Function
#End Region

#Region " 画像表示倍率バーの表示、非表示を設定する"

    ''' <summary>
    ''' 画像表示倍率バーの表示、非表示を設定する    'V2.2.0.0① 
    ''' </summary>
    ''' <param name="flg"></param>
    Public Sub SetMagnifyBar(ByVal flg As Boolean)

        Try

            If flg = True Then
                Form1.VideoLibrary1.SetTrackBarVisible(True)
                Form1.BtnStartPosSet.Visible = False
            Else
                Form1.VideoLibrary1.SetTrackBarVisible(False)
                Form1.BtnStartPosSet.Visible = True
            End If

        Catch ex As Exception

        End Try

    End Sub
#End Region

#Region " 画像表示倍率バーの表示、非表示を設定する"
    ''' <summary>
    ''' 画像表示倍率バーの表示、非表示を設定する   'V2.2.0.0⑥
    ''' </summary>
    ''' <returns></returns>
    Public Function MoveStop() As Integer
        Dim r As Integer
        Dim sts As Integer
        Dim xPos As Double
        Dim yPos As Double


        Try
            MoveStop = 0

            ' 一時停止を行う
            If giCutStop <> 0 Then
                System.Windows.Forms.Application.DoEvents()

                If Form1.JudgeStop() Then

                    Call ZGETBPPOS(xPos, yPos)                              ' BP現在位置取得
                    ObjCrossLine.CrossLineDispXY(xPos, yPos)                ' クロスライン表示

                    r = ObjSys.WaitStart(sts)           ' ADV(1)/HALT(2)/RESET(3)待ち(ADJ ON時)
                    If sts = 1 Then
                        Form1.btnCutStop.BackColor = SystemColors.Control
                    End If
                    If (r = cFRS_ERR_RST) Then          ' RESET SW押下 ?
                        Return (r)                      ' Return値 = RESET SW押下
                    End If
                    If (r < cFRS_NORMAL) Then           ' エラー ?
                        Return (r)                      ' Return値設定
                    End If
                    ObjCrossLine.CrossLineOff()                             ' クロスライン非表示
                End If
            End If

        Catch ex As Exception

        End Try

    End Function
#End Region

    ''' <summary>
    ''' 筐体カバーが閉じるのを待つ 
    ''' </summary>
    ''' <returns></returns>
    Public Function WaitCoverClose() As Integer
        Dim r As Integer


        Try

            Do
                System.Threading.Thread.Sleep(100)                      ' Wait(ms)
                System.Windows.Forms.Application.DoEvents()

                ' 筐体カバー閉を確認する
                ' r = ObjSys.CoverCheck(0, False)                       ' 筐体カバー閉チェック(RESETキー無効指定, 原点復帰処理中以外)
                r = ObjSys.Sub_CoverCheck(gSysPrm, 0, False)
                If (r = cFRS_NORMAL) Then Exit Do
                If (r <> ERR_OPN_CVR) Then Return (r) '                 ' 非常停止等のエラー

                ' "筐体カバーを閉じて","","STARTキーを押すか、OKボタンを押して下さい。"
                r = ObjLoader.Sub_CallFrmMsgDisp(Form1.System1, cGMODE_MSG_DSP, cFRS_ERR_START, True,
                                My.Resources.MSG_SPRASH36, "", My.Resources.MSG_frmLimit_07, System.Drawing.Color.Blue, System.Drawing.Color.Blue, System.Drawing.Color.Blue)
                ' 非常停止等のエラーならアプリ強制終了へ(エラーメッセージは表示済み)
                'If (r < cFRS_NORMAL) Then Return (RtnCode)             ' ###193
                If (r < cFRS_NORMAL) Then Return (r) '                  ' ###193
                Call COVERLATCH_CLEAR()                                 ' カバー開ラッチのクリア

            Loop While (1)

        Catch ex As Exception

        End Try



    End Function


    ''' <summary>
    ''' ブロック番号から複数抵抗値指定処理実行時のデータを取得する 
    ''' </summary>
    ''' <param name="blkx:現在のブロック番号X"></param>
    ''' <param name="blky:現在のブロック番号Y"></param>
    ''' <param name="blkData"></param>
    ''' <returns></returns>
    Public Function GetMultiBlockdata(ByVal blkx As Integer, ByVal blky As Integer, ByRef blkData As BLOCK_DATA) As Integer
        Dim Multicnt As Integer = 0
        Dim Maxblock As Integer = 0

        Try
            GetMultiBlockdata = cFRS_ERR_TRIM

            '行列の方向
            If stMultiBlock.gStepRpt = 0 Then
                ' X方向の列数が指定されている

                ' X方向のブロック番号がどの指定抵抗値にいるか判定する
                For Multicnt = 0 To 4
                    If (stMultiBlock.BLOCK_DATA(Multicnt).gBlockCnt) <> 0 Then
                        ' 先頭ブロック列と最終ブロック列の範囲かチェック
                        Maxblock = Maxblock + stMultiBlock.BLOCK_DATA(Multicnt).gBlockCnt
                        If (blkx <= Maxblock) Then

                            BlockDataCopy(blkData, Multicnt)
                            GetMultiBlockdata = cFRS_NORMAL

                            Exit For
                        End If
                    End If
                Next

            Else
                ' Y方向の行数が指定されている

                ' X方向のブロック番号がどの指定抵抗値にいるか判定する
                For Multicnt = 0 To 4
                    If (stMultiBlock.BLOCK_DATA(Multicnt).gBlockCnt) <> 0 Then
                        ' 先頭ブロック列と最終ブロック列の範囲かチェック
                        Maxblock = Maxblock + stMultiBlock.BLOCK_DATA(Multicnt).gBlockCnt
                        If (blky <= Maxblock) Then

                            BlockDataCopy(blkData, Multicnt)
                            GetMultiBlockdata = cFRS_NORMAL
                            Exit For

                        End If
                    End If
                Next Multicnt

            End If

        Catch ex As Exception

            MsgBox("GetMultiBlockdata() Err = " & ex.Message)

        End Try

    End Function


    ''' <summary>
    ''' マルチデータ指定番号のマルチブロックデータをコピー 
    ''' </summary>
    ''' <param name="blkData"></param>
    ''' <param name="Multicnt"></param>
    ''' <returns></returns>
    Public Function BlockDataCopy(ByRef blkData As BLOCK_DATA, ByVal Multicnt As Integer) As Integer
        Dim rn As Integer

        Try

            blkData.DataNo = Multicnt + 1

            ' 抵抗数分コピー
            For rn = 0 To MAX_RES_USER - 1

                ' ブロック数 
                blkData.gBlockCnt = stMultiBlock.BLOCK_DATA(Multicnt).gBlockCnt
                ' 目標抵抗値
                blkData.dblNominal(rn) = stMultiBlock.BLOCK_DATA(Multicnt).dblNominal(rn)
                ' 単位
                blkData.iUnit(rn) = stMultiBlock.BLOCK_DATA(Multicnt).iUnit(rn)
                ' 温度ｾﾝｻｰ画面、抵抗ﾄﾘﾐﾝｸﾞ画面のどちらか
                If (UserSub.IsTrimType1() Or UserSub.IsTrimType4()) Then
                    ' 温度センサーの場合は抵抗値と単位のみ
                Else
                    ' 補正値 
                    blkData.dblCorr(rn) = stMultiBlock.BLOCK_DATA(Multicnt).dblCorr(rn)
                End If


            Next rn


        Catch ex As Exception

            MsgBox("BlockDataCopy() Err = " & ex.Message)

        End Try


    End Function


    'V2.2.0.0⑯↓
    ''' <summary>
    ''' 複数抵抗値データの適用
    ''' </summary>
    ''' <param name="BlockCntX"></param>
    ''' <param name="BlockCntY"></param>
    ''' <returns></returns>
    Public Function ApplyMultiData(ByVal BlockCntX As Integer, ByVal BlockCntY As Integer) As Integer
        Dim r As Integer
        Dim rn As Integer

        Try

            stExecBlkData.DataNo = 0
            stExecBlkData.Initialize()

            ' 更新前のデータを保存しておく
            For rn = 1 To stPLT.RCount                                  ' 抵抗数分繰返す
                If UserModule.GetOkMarkingResNo(rn) Then
                    'OKマーキングデータは保存しない
                Else
                    ' 抵抗値の保存 
                    stDefaultBlock(rn).dblNominal = stREG(rn).dblNOM

                    ' 温度ｾﾝｻｰ画面、抵抗ﾄﾘﾐﾝｸﾞ画面のどちらか
                    If (UserSub.IsTrimType1() Or UserSub.IsTrimType4()) Then
                        ' 温度センサーの場合は抵抗値と単位のみ
                    Else
                        ' 補正値の保存 
                        stDefaultBlock(rn).dblCorr = stUserData.dNomCalcCoff(rn)
                    End If

                End If
            Next rn

            ' 指定ブロックで適用されるブロックデータの取得
            r = GetMultiBlockdata(stCounter.BlockCntX, stCounter.BlockCntY, stExecBlkData)
            If r = cFRS_NORMAL Then
                Dim iUserResCnt As Integer
                iUserResCnt = UserBas.GetRCountExceptMeasure()
                If iUserResCnt > MAX_RES_USER Then
                    iUserResCnt = MAX_RES_USER
                End If

                '' 現状のブロックデータを書き換える 
                Dim Cnt As Integer

                For rn = 1 To stPLT.RCount Step 1
                    If UserModule.IsCutResistor(rn) Then
                        If iUserResCnt = 1 Then
                            Cnt = 1
                        Else
                            Cnt = UserSub.GetResNumberInCircuit(rn)
                        End If
                        stREG(rn).dblNOM = stExecBlkData.dblNominal(Cnt - 1) ' 目標抵抗値
                        ' 温度ｾﾝｻｰ画面、抵抗ﾄﾘﾐﾝｸﾞ画面のどちらか
                        If (UserSub.IsTrimType1() Or UserSub.IsTrimType4()) Then
                            ' 温度センサーの場合は補正値なし
                        Else
                            stUserData.dNomCalcCoff(rn) = stExecBlkData.dblCorr(Cnt - 1)
                        End If

                    End If
                Next rn

            Else
                Z_PRINT("ブロックデータが取得できなかったため、初期データを使用します。 [X=" & stCounter.BlockCntX.ToString & "], [Y=" & stCounter.BlockCntY & "]")
            End If


        Catch ex As Exception

            MsgBox("ApplyMultiData() Err = " & ex.Message)

        End Try

    End Function

    ''' <summary>
    ''' 複数抵抗値取得用のデータを元に戻す 
    ''' </summary>
    ''' <param name="BlockCntX"></param>
    ''' <param name="BlockCntY"></param>
    ''' <returns></returns>
    Public Function RestoreMultiData(ByVal BlockCntX As Integer, ByVal BlockCntY As Integer) As Integer
        Dim rn As Integer

        Try

            ' 現状のブロックデータを書き換える 
            For rn = 1 To stPLT.RCount                                  ' 抵抗数分繰返す

                If UserModule.GetOkMarkingResNo(rn) Then
                    'OKマーキングデータは保存しない
                Else

                    ' 抵抗値の保存 
                    stREG(rn).dblNOM = stDefaultBlock(rn).dblNominal

                    ' 温度ｾﾝｻｰ画面、抵抗ﾄﾘﾐﾝｸﾞ画面のどちらか
                    If (UserSub.IsTrimType1() Or UserSub.IsTrimType4()) Then
                        ' 温度センサーの場合は補正値なし
                    Else
                        ' 補正値の保存 
                        stUserData.dNomCalcCoff(rn) = stDefaultBlock(rn).dblCorr

                    End If

                End If

            Next rn


        Catch ex As Exception
            MsgBox("RestoreMultiData() Err = " & ex.Message)
        End Try


    End Function


    ''' <summary>
    ''' カット位置補正のデータ(認識位置、テンプレート番号、パターン番号)が同じかチェックする。
    ''' </summary>
    ''' <param name="Rn:抵抗番号"></param>
    ''' <returns>戻り値：True：補正を実行する、False：補正をしない</returns>
    Public Function CompareCutCorrData(ByVal Rn As Integer) As Boolean

        Try

            ' 第１抵抗は必ず行う
            If Rn = 1 Then
                Return (True)
            End If

            ' カット位置補正の位置が同じかどうかチェックする 
            If stPTN(Rn).dblPosX <> stPTN(Rn - 1).dblPosX OrElse stPTN(Rn).dblPosY <> stPTN(Rn - 1).dblPosY Then
                Return (True)
            End If

            ' テンプレート番号、パターン番号が同じかチェックする
            If stPTN(Rn).intGRP <> stPTN(Rn - 1).intGRP OrElse stPTN(Rn).intPTN <> stPTN(Rn - 1).intPTN Then
                Return (True)
            End If

            ' カット位置補正は実行しないので補正値は前の実行結果を使用する 
            ' ズレ量保存
            stPTN(Rn).dblDRX = stPTN(Rn - 1).dblDRX                                   ' ズレ量X
            stPTN(Rn).dblDRY = stPTN(Rn - 1).dblDRY                                   ' ズレ量Y

            Return (False)

        Catch ex As Exception
            MsgBox("CompareCutCorrData() Err = " & ex.Message)
        End Try


    End Function

    ''' <summary>
    ''' 機器名に[HIOKI]を含んでいる場合、指定の文字列に”RES:RANG***”があれば"RES:RANG 目標値"に置き換える 
    ''' </summary>
    ''' <param name="strcmd">コマンド文字列</param>
    Public Sub ChangeHIOKICommand(ByRef strcmd As String, ByVal Gno As Integer, ByVal dNOMx As Double)

        Dim Pos As Integer

        Try

            Dim result As Boolean = stGPIB(Gno).strGNAM.ToUpper().Contains("HIOKI")
            If result = False Then
                Return
            End If

            Dim SearchStr As String = "RES:RANG"
            Dim PutStr As String = ":RES:RANG"
            ' HIOKIの場合、設定コマンドの中に「 "RES:RANG"」があったら、設定の目標値を書き換える 
            Dim cmd = strcmd.ToUpper().IndexOf(SearchStr)

            If cmd >= 0 Then
                Dim arr() As String = strcmd.Split(";")
                For i As Integer = 0 To arr.Length - 1
                    Pos = arr(i).ToUpper.IndexOf(PutStr)
                    If Pos >= 0 Then
                        ' 存在したらレンジ目標値指定に置き換える
                        ''V2.2.1.3①↓
                        ''V2.2.1.3①　arr(i) = PutStr & " " & CInt(dNOMx).ToString()
                        arr(i) = PutStr & " " & CDbl(dNOMx).ToString("0.#######")
                        ''V2.2.1.3①↑
                        Exit For
                    End If
                Next
                Dim TrigStr As String = ""
                ' 置き換えた文字列を結合する
                For i As Integer = 0 To arr.Length - 1
                    If TrigStr.Trim <> "" Then
                        TrigStr = TrigStr & ";"
                    End If
                    TrigStr = TrigStr & arr(i)
                Next

                strcmd = TrigStr

            End If

        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>
    ''' HIOKI製のコマンドを編集 
    ''' </summary>
    ''' <returns></returns>
    Public Function Change_HIOKI_CMD(ByVal Gno As Integer, ByVal dNOMx As Double, ByVal bChangeFlg As Boolean) As Integer
        Dim Pos As Integer
        Dim SearchStr As String = "RES:RANG"
        Dim PutStr As String = ":RES:RANG"
        Dim TrigStr As String = ""

        Try

            Dim arr() As String = stGPIB(Gno).strCTRG.Split(";")
            For i As Integer = 0 To arr.Length - 1
                Pos = arr(i).ToUpper.IndexOf(PutStr)
                If Pos >= 0 Then
                    If bChangeFlg = True Then
                        ' 存在したらレンジ目標値指定に置き換える
                        arr(i) = PutStr & " " & CDbl(dNOMx).ToString("0.#######")
                    Else
                        arr(i) = ""
                    End If
                End If
                If TrigStr.Trim <> "" Then
                    TrigStr = TrigStr & ";"
                End If
                TrigStr = TrigStr & arr(i)
            Next
            sStrTrig = TrigStr

        Catch ex As Exception

        End Try

    End Function


    'V2.2.1.7③↓
    ''' <summary>
    ''' マーク印字自動運転中に発生したアラームのリストを保存 
    ''' </summary>
    ''' <param name="strTrimDataName"></param>
    ''' <param name="alarmcnt"></param>
    ''' <returns></returns>
    Public Function SetLotMarkAlarm(ByVal strTrimDataName As String, ByVal PlateNo As Integer) As Integer

        Try

            LotMarkingAlarmCnt = LotMarkingAlarmCnt + 1
            ReDim Preserve gMarkAlarmList(LotMarkingAlarmCnt)

            gMarkAlarmList(LotMarkingAlarmCnt).AlarmTrimData = System.IO.Path.GetFileName(strTrimDataName)

            gMarkAlarmList(LotMarkingAlarmCnt).LotCount = PlateNo

        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' マーク印字用ログファイル出力 
    ''' </summary>
    ''' <returns></returns>
    Public Function WriteLogMarkPrint() As Integer

        Dim sMarkPrintLogPath As String = "C:\TRIMDATA\MARKLOG\"
        Dim sMarkPrintLogFileName As String = "MARK_YYYYMM.LOG"
        Dim sHeaderData As String

        Try

            'V2.2.1.7⑥ ↓
            ' マーク印字で無ければアラーム表示しない。 
            If UserSub.IsTrimType5() <> True Then
                Return cFRS_NORMAL
            End If
            'V2.2.1.7⑥ ↑

            sMarkPrintLogFileName = "MARK_" & Now.ToString("yyyyMM") & ".LOG"

            Dim sPath As String = sMarkPrintLogPath & sMarkPrintLogFileName

            ' 自動運転終了時間のセット 
            SetAutoOpeEndTime()

            sHeaderData = ""
            If File.Exists(sPath) = False Then
                ' ファイルが無い場合、ヘッダー出力する 
                sHeaderData = "作業者,日付,依頼No(ロットNo),開始番号,終了番号,自動運転開始時間,自動運転終了時間"
            End If

            Dim EndNum As String = ""

            If (stREG(1).STCUT(1).cMarkStartNum <> "") Then
                Dim len As Integer = stREG(1).STCUT(1).cMarkStartNum.Length
                EndNum = (Integer.Parse(stREG(1).STCUT(1).cMarkStartNum) + MarkingCount - 1).ToString.PadLeft(len, "0"c)
            End If

            Using WSR As New System.IO.StreamWriter(sPath, True, System.Text.Encoding.GetEncoding("Shift-JIS"))  ' 第２引数 上書きは、False
                If sHeaderData <> "" Then
                    WSR.WriteLine(sHeaderData)                          ' ヘッダ出力
                End If

                WSR.WriteLine(stUserData.sOperator.ToString & "," & gLogMarkPrint.sDate & "," & stUserData.sLotNumber & "," & stREG(1).STCUT(1).cMarkStartNum & "," & EndNum & "," & gLogMarkPrint.sAutoOpeStartTime & "," & gLogMarkPrint.sAutoOpeEndTime)

            End Using


        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' 自動運転開始時間、日付の設定
    ''' </summary>
    Public Sub SetAutoOpeStartTime()

        Try

            gLogMarkPrint.sDate = Now.ToString("yyyy/MM/dd")

            gLogMarkPrint.sAutoOpeStartTime = Now.ToString("yyyy/MM/dd HH:mm:ss")

        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>
    ''' 自動運転終了時間 
    ''' </summary>
    Public Sub SetAutoOpeEndTime()

        Try

            gLogMarkPrint.sAutoOpeEndTime = Now.ToString("yyyy/MM/dd HH:mm:ss")

        Catch ex As Exception

        End Try

    End Sub

    'V2.2.1.7③↑

End Module

'=============================== END OF FILE ===============================
