'===============================================================================
'   Description : ブロック単位のトリミング処理
'
'   Copyright(C) Laser Front 2010
'
'===============================================================================
Option Strict Off
Option Explicit On
Module Trimming
#Region "グローバル定数/変数の定義"
    '-------------------------------------------------------------------------------
    '   定数定義
    '-------------------------------------------------------------------------------
    '----- 最大値/最小値 -----
    Public Const cMAXcMARKINGcSTRLEN As Integer = 18        ' マーキング文字列最大長(byte)
    Public Const cCNDNUM As Integer = 4                     ' 1ｶｯﾄの最大加工条件数(FL用)
    Public Const cResultMax As Integer = 256                ' トリミング結果データの最大配列数
    Public Const cResultAry As Integer = 999                ' トリミング結果データの最大数

    '----- 入出力 -----
    Public Const INP_MAX As Integer = 5                     ' 軸Signal状態の数
    Public Const INP_ICSLSS As Integer = 0                  ' [0]:コンソールSWセンス
    Public Const INP_IITLKS As Integer = 1                  ' [1]:インターロック関係SWセンス
    Public Const INP_AUTLODL As Integer = 2                 ' [2]:オートローダLO
    Public Const INP_AUTLODH As Integer = 3                 ' [3]:オートローダHI
    Public Const INP_ATTNATE As Integer = 4                 ' [4]:固定アッテネータ

    Public Const OUT_MAX As Integer = 4                     ' 軸Signal状態の数
    Public Const OUT_OCSLLN As Integer = 0                  ' [0]:コンソール制御
    Public Const OUT_OSYSCTL As Integer = 1                 ' [1]:サーボパワー
    Public Const OUT_AUTLODL As Integer = 2                 ' [2]:オートローダLO
    Public Const OUT_AUTLODH As Integer = 3                 ' [3]:オートローダHI
    Public Const OUT_SIGNALT As Integer = 4                 ' [4]:シグナルタワー(未使用)
    Public Const OUT_Z2CONT As Integer = 6                  ' [5]:Z2サーボパワー

    '----- トリミング要求データのデータタイプ -----
    Public Const DATTYPE_PLATE As UShort = 1                ' プレートデータ
    Public Const DATTYPE_REGI As UShort = 2                 ' 抵抗データ
    Public Const DATTYPE_CUT As UShort = 3                  ' カットデータ
    Public Const DATTYPE_PARAM As UShort = 4                ' カットパラメータ
    Public Const DATTYPE_GPIB As UShort = 8                 ' GPIB設定用データ

    '-------------------------------------------------------------------------------
    '   カットタイプ別パラメータ形式定義(VB→INtime)
    '-------------------------------------------------------------------------------
    '----- ST cut -----
    Public Structure PRM_CUT_ST                             ' ST cutパラメータ形式定義
        Dim DIR As UShort                                   ' カット方向(1:+X, 2:-X, 3:+Y, 4:-Y)
        Dim MODE As UShort                                  ' 動作モード(0:NOM, 1:リターン, 2:リトレース, 3:斜め)
        Dim angle As UShort                                 ' 斜めカット角度(0～359)
        Dim Length As Double                                ' 最大カッティング長(0.0001～20.0000(mm))
        Dim spd2 As Double                                  ' 復路スピード(mm/s)
        Dim qrate2 As Double                                ' リターン/リトレースのQrate2(KHz)
        'Dim chenge As Double                                ' 切り替えポイント(0.0～100.0%)(SL436K用)
    End Structure

    '----- L cut -----
    Public Structure PRM_CUT_L                              ' L cutパラメータ形式定義
        Dim DIR As UShort                                   ' カット方向(1:+X, 2:-X, 3:+Y, 4:-Y)
        Dim tdir As UShort                                  ' Lターン方向(1:+X, 2:-X, 3:+Y, 4:-Y)
        Dim MODE As UShort                                  ' 動作モード(0:NOM, 1:リターン, 2:リトレース, 3:斜め)
        Dim angle As UShort                                 ' 斜めカット角度(0～359)
        Dim turn As Double                                  ' Lターンポイント(0.0～100.0(%))
        Dim L1 As Double                                    ' L1 最大カッティング長(0.0001～20.0000(mm))
        Dim L2 As Double                                    ' L2 最大カッティング長(0.0001～20.0000(mm))
        Dim r As Double                                     ' ターンの円弧半径(mm)
        Dim spd2 As Double                                  ' 復路スピード(mm/s)
        Dim qrate2 As Double                                ' Qrate2(KHz)
        <VBFixedArray(1)> Dim qrate3() As Double            ' Qrate3(KHz) FL時のﾘﾀｰﾝ/ﾘﾄﾚｰｽ時のQrate
        '                                                   ' Qrate3[0]: 加工条件番号3のQﾚｰﾄを設定(Lﾀｰﾝ前のQrate)
        '                                                   ' Qrate3[1]: 加工条件番号4のQﾚｰﾄを設定(Lﾀｰﾝ後のQrate)
        <VBFixedArray(2)> Dim spd3() As Double              ' FL時のLﾀｰﾝ後/ﾘﾀｰﾝ/ﾘﾄﾚｰｽ時のスピード(mm/s)
        '                                                   ' Spd3[0]: Lﾀｰﾝ後のｽﾋﾟｰﾄﾞ
        '                                                   ' Spd3[1]: ﾘﾀｰﾝ/ﾘﾄﾚｰｽ時のLﾀｰﾝ前のｽﾋﾟｰﾄﾞ
        '                                                   ' Spd3[2]: ﾘﾀｰﾝ/ﾘﾄﾚｰｽ時のLﾀｰﾝ後のｽﾋﾟｰﾄﾞ
        ' この構造体を初期化するには、"Initialize" を呼び出さなければなりません。 
        Public Sub Initialize()
            ReDim qrate3(1)
            ReDim spd3(2)
        End Sub
    End Structure

    '----- HOOK cut -----
    Public Structure PRM_CUT_HOOK                           ' HOOK cutパラメータ形式定義
        Dim DIR As UShort                                   ' カット方向(1:+X, 2:-X, 3:+Y, 4:-Y)
        Dim tdir As UShort                                  ' Lターン方向(1:+X, 2:-X, 3:+Y, 4:-Y)
        Dim turn As Double                                  ' Lターンポイント(0.0～100.0(%))
        Dim L1 As Double                                    ' L1 最大カッティング長(0.0001～20.0000(mm))
        Dim r1 As Double                                    ' ターン1の円弧半径(mm)
        Dim L2 As Double                                    ' L2 最大カッティング長(0.0001～20.0000(mm))
        Dim r2 As Double                                    ' ターン2の円弧半径(mm)
        Dim L3 As Double                                    ' L3 最大カッティング長(0.00001～20.0000(mm))
        <VBFixedArray(1)> Dim qrate2() As Double            ' Qrate2(KHz) FL時のL2/L3のQrate
        '                                                   ' Qrate2[0]: 加工条件番号2のQﾚｰﾄを設定(L2のQrate)
        '                                                   ' Qrate2[1]: 加工条件番号3のQﾚｰﾄを設定(L3のQrate)
        <VBFixedArray(1)> Dim spd2() As Double              ' FL時のL2/L3のスピード(mm/s)
        '                                                   ' Spd2[0]: L2のｽﾋﾟｰﾄﾞ
        '                                                   ' Spd2[1]: L3のｽﾋﾟｰﾄﾞ
        ' この構造体を初期化するには、"Initialize" を呼び出さなければなりません。 
        Public Sub Initialize()
            ReDim qrate2(1)
            ReDim spd2(1)
        End Sub
    End Structure

    '----- INDEX cut -----
    Public Structure PRM_CUT_INDEX                          ' INDEX cutパラメータ形式定義
        Dim DIR As UShort                                   ' カット方向(1:+X, 2:-X, 3:+Y, 4:-Y)
        Dim maxindex As UShort                              ' インデックス数(1～32767)
        Dim measure As UShort                               ' 測定モード(0:高速, 1:高精度)
        Dim Length As Double                                ' インデックス長(0.0001～20.0000(mm))
    End Structure

    '----- SCAN cut -----
    Public Structure PRM_CUT_SCAN                           ' SCAN cutパラメータ形式定義
        Dim DIR As UShort                                   ' カット方向(1:+X, 2:-X, 3:+Y, 4:-Y)
        Dim sdir As UShort                                  ' ステップ方向(1:+X, 2:-X, 3:+Y, 4:-Y)
        Dim lines As UShort                                 ' 本数(1～n)
        Dim Length As Double                                ' カッティング長(0.0001～20.0000(mm))
        Dim pitch As Double                                 ' ピッチ(0.0001～20.0000(mm))
    End Structure

    '----- Letter Marking -----
    Public Structure PRM_CUT_MARKING                        ' Letter Markingパラメータ形式定義
        '                                                   ' 文字
        <VBFixedArray(cMAXcMARKINGcSTRLEN - 1)> Dim str() As Byte
        Dim magnify As Double                               ' 倍率(１～999)
        Dim DIR As UShort                                   ' 文字の向き(1:0, 2:90, 3:180, 4:270)
        ' この構造体を初期化するには、"Initialize" を呼び出さなければなりません。 
        Public Sub Initialize()
            ReDim str(cMAXcMARKINGcSTRLEN - 1)
        End Sub
    End Structure

    '----- C cut -----
    Public Structure PRM_CUT_C                              ' C cutパラメータ形式定義
        Dim DIR As UShort                                   ' カット方向(1:CW, 2:CCW)
        Dim angle As UShort                                 ' カット角度(0～359)
        Dim count As UShort                                 ' 回数
        Dim st_r As Double                                  ' 円弧半径 (mm)
        Dim pitch As Double                                 ' ピッチ
    End Structure

    '----- ES cut -----
    Public Structure PRM_CUT_ES                             ' ES cutパラメータ形式定義
        Dim DIR As UShort                                   ' カット方向  1:+X, 2:-X, 3:+Y, 4:-Y
        Dim L1 As Double                                    ' L1 最大カッティング長(0.0001～20.0000(mm))
        Dim EsPoint As Double                               ' ESﾎﾟｲﾝﾄ(-99.9999～0.0000%))
        Dim ESchangerate As Double                          ' ES判定変化率(0.0～100.0%))
        Dim EScutlen As Double                              ' ES後ｶｯﾄ長(0.0001～20.0000(mm))
    End Structure

    '----- ES2 cut -----
    Public Structure PRM_CUT_ES2                            ' ES2 cutパラメータ形式定義
        Dim DIR As UShort                                   ' カット方向(1:+X, 2:-X, 3:+Y, 4:-Y)
        Dim L1 As Double                                    ' L1 最大カッティング長(0.0001～20.0000(mm))
        Dim EsPoint As Double                               ' ESﾎﾟｲﾝﾄ(-99.9999～0.0000%)
        Dim ESWide As Double                                ' ES判定変化率(0.0～100.0%)
        Dim ESWide2 As Double                               ' ES後変化率(0.0～100.0%)
        Dim EScount As UShort                               ' ES後確認回数(0～20)
    End Structure

    '----- UCUTパラメータ(1要素) -----
    Public Structure UCUT_PARAM_EL                          ' UCUTパラメータ(1要素)形式定義
        Dim RATIO As Double                                 ' 目標値に対する初期値の差(%)
        Dim LTP As Double                                   ' Lターンポイント(0.0～100.0%)
        Dim LTP2 As Double                                  ' Lターンポイント2(0.0～100.0%)
        Dim L1 As Double                                    ' L1 最大カッティング長(0.0001～20.0000mm)
        Dim L2 As Double                                    ' L2 最大カッティング長(0.0001～20.0000mm)
        Dim r As Double                                     ' 円弧半径 (mm)
        Dim V As Double                                     ' 速度(mm/s)
        Dim NOM As Double                                   ' 目標値
        Dim Flg As Boolean                                  ' データ有効(未使用)
    End Structure

    '----- UCUTパラメータ -----
    Public Structure S_UCUTPARAM_EL                         ' UCUTパラメータ形式定義
        Dim RNO As UShort
        Dim NOM As Double
        Dim PRM_UNIT As UCUT_PARAM_EL
    End Structure

    '----- UCUTパラメータテーブル(1抵抗分) -----
    Public Structure S_UCUTPARAM                            ' UCUTパラメータ形式定義
        <VBFixedArray(19)> Dim EL() As S_UCUTPARAM_EL       ' UCUTパラメータ 
        ' この構造体を初期化するには、"Initialize" を呼び出さなければなりません。 
        Public Sub Initialize()
            ReDim EL(19)
        End Sub
    End Structure

    '-------------------------------------------------------------------------------
    '   カットデータ形式定義(VB→INtime)
    '-------------------------------------------------------------------------------
    Public Structure PRM_CUT_DATA                           ' カットデータ形式定義
        Dim CP1 As UShort                                   ' カット番号 1-20
        Dim CP2 As UShort                                   ' 定電流印加後測定遅延時間(0-32767msec) 
        Dim CP3 As UShort                                   ' カット形状(1:st, 2:L, 3:HK, 4:IX 他)
        Dim cp4_x As Double                                 ' カットスタート座標X(-80.0000～+80.0000)
        Dim cp4_y As Double                                 ' カットスタート座標Y(-80.0000～+80.0000)
        Dim CP5 As Double                                   ' カットスピード(0.1～409.0mm/s)
        Dim CP6 As Double                                   ' レーザーQスイッチレート(0.1～50.0KHz) ※FL時は加工条件番号1のQﾚｰﾄを設定
        Dim CP7 As Double                                   ' カットオフ %(-99.999 ～ +999.999)
        Dim CP71 As Double                                  ' カットデータ平均化率(0.0～100.0, 0%)(未使用)
        <VBFixedArray(cCNDNUM - 1)> Dim CP72() As Byte      ' 加工条件番号1～4(FL用) 
        'Dim CP50 As UShort                                  ' パルス幅制御(0:無し 1:有り)(SL436K用)
        'Dim CP51 As Double                                  ' パルス幅時間(SL436K用)
        'Dim CP52 As Double                                  ' LSwパルス幅時間(外部シャッタ)(SL436K用)
        Dim dummy As PRM_CUT_HOOK                           ' カットパラメータ(union) ※union定義ないので最大のものを指定

        ' この構造体を初期化するには、"Initialize" を呼び出さなければなりません。 
        Public Sub Initialize()
            ReDim CP72(cCNDNUM - 1)
        End Sub
    End Structure

    '-------------------------------------------------------------------------------
    '   抵抗データ形式定義(VB→INtime)
    '-------------------------------------------------------------------------------
    Public Structure PRM_REGISTER                           ' 抵抗データ形式定義
        Dim PR1a As UShort                                  ' 抵抗番号(1-999=トリミング, 1000-9999=マーキング)
        Dim PR2a As UShort                                  ' 判定測定(0:高速, 1:高精度
        Dim PR3a As UShort                                  ' サーキット(抵抗が属するサーキット番号)
        Dim PR4_ha As UShort                                ' ハイ側プローブ番号
        Dim PR4_la As UShort                                ' ロー側プローブ番号
        Dim PR4_g1a As UShort                               ' 第1アクティブガード番号
        Dim PR4_g2a As UShort                               ' 第2アクティブガード番号
        Dim PR4_g3a As UShort                               ' 第3アクティブガード番号
        Dim PR4_g4a As UShort                               ' 第4アクティブガード番号
        Dim PR4_g5a As UShort                               ' 第5アクティブガード番号
        Dim PR5a As UInteger                                ' External bits
        Dim PR6a As UShort                                  ' ポーズタイム(External bits出力後のウェイト) (msec)
        Dim PR7a As UShort                                  ' 目標値指定(0:絶対値, 1:レシオ, 2:計算式)
        Dim PR8a As UShort                                  ' ベース抵抗No.(レシオ時の基準抵抗番号)
        Dim PR9a As Double                                  ' トリミング目標値(ohm)
        Dim PR10a As UShort                                 ' 電圧変化スロープ(0:+スロープ, 1:-スロープ) ※ﾌﾟﾚｰﾄﾃﾞｰﾀの測定モード=電圧の場合有効
        Dim PR11_Ha As Double                               ' IT Limit H(-99.99～9999.99%)
        Dim PR11_La As Double                               ' IT Limit L(-99.99～9999.99%)
        Dim PR12_Ha As Double                               ' FT Limit H(-99.99～9999.99%)
        Dim PR12_La As Double                               ' FT Limit L(-99.99～9999.99%)
        Dim PR13a As UShort                                 ' カット数(1～20)
        Dim PR14 As UShort                                  ' カット位置補正フラグ(0:補正しない, 1:補正する)
        Dim PR14_Ha As Double                               ' イニシャルOKテストHIGHリミット(SL436K用)
        Dim PR14_La As Double                               ' イニシャルOKテストLOWリミット (SL436K用)
        Dim fCutMag As Double                               ' 切上げ倍率(CHIPのみ)
        Dim pCutData As UInteger                            ' カットデータポインタ(INTIME側で使用)
    End Structure

    '-------------------------------------------------------------------------------
    '   プレートデータ形式定義(VB→INtime)
    '-------------------------------------------------------------------------------
    Public Structure TRIM_PLATE_DATA                        ' プレートデータ形式定義
        Dim wCircuitCnt As UShort                           ' サーキット数
        Dim wRegistCnt As UShort                            ' 抵抗数
        Dim wTrimMode As UShort                             ' 測定モード(0:抵抗, 1:電圧)
        Dim wDelayTrim As UShort                            ' ディレイトリム(0=なし, 1=ﾃﾞｨﾚｲﾄﾘﾑを実行する, 2=ﾃﾞｨﾚｲﾄﾘﾑ2を実行する)
        Dim fBPOffsetX As Double                            ' BPオフセットX(mm)
        Dim fBPOffsetY As Double                            ' BPオフセットY(mm)
        Dim fAdjustOffsetX As Double                        ' アジャスト位置X(mm)
        Dim fAdjustOffsetY As Double                        ' アジャスト位置Y(mm)
        Dim fNgCriterion As Double                          ' NG判定基準(%)
        Dim fZStepPos As Double                             ' Z軸ｽﾃｯﾌﾟ&ﾘﾋﾟｰﾄ位置
        Dim fZTrimPos As Double                             ' Z軸ｺﾝﾀｸﾄ位置
        Dim fReProbingX As Double                           ' 再ﾌﾟﾛｰﾋﾞﾝｸﾞX移動量
        Dim fReProbingY As Double                           ' 再ﾌﾟﾛｰﾋﾞﾝｸﾞY移動量
        Dim wReProbingCnt As UShort                         ' 再ﾌﾟﾛｰﾋﾞﾝｸﾞ回数
        Dim wInitialOK As UShort                            ' ｲﾆｼｬﾙOKﾃｽﾄ有無(0:無し 1:有り))(SL436K用)
        Dim wNGMark As UShort                               ' NGﾏｰｷﾝｸﾞする/しない)(SL436K用)
        Dim w4Terminal As UShort                            ' 4端子ｵｰﾌﾟﾝﾁｪｯｸする/しない)(SL436K用)
        Dim wLogMode As UShort                              ' ﾛｷﾞﾝｸﾞﾓｰﾄﾞ
        '                                                   ' 0:しない, 1:INITIAL TEST, 2:FINAL TEST, 3:INITIAL + FINAL)	
    End Structure

    '-------------------------------------------------------------------------------
    '   GPIB設定用データ形式定義(VB→INtime)
    '-------------------------------------------------------------------------------
    Public Structure TRIM_PLATE_GPIB                        ' GPIB設定用データ形式定義
        Dim wGPIBmode As UShort                             ' GP-IB制御(0:しない 1:する)
        Dim wDelim As UShort                                ' ﾃﾞﾘﾐﾀ(0:CR+LF 1:CR 2:LF 3:NONE)
        Dim wTimeout As UShort                              ' ﾀｲﾑｱｳﾄ(0～1000)(100ms単位)
        Dim wAddress As UShort                              ' 機器ｱﾄﾞﾚｽ(0～30)
        <VBFixedArray(39)> Dim strI() As Byte               ' 初期化ｺﾏﾝﾄﾞ(MAX40byte)
        <VBFixedArray(9)> Dim strT() As Byte                ' ﾄﾘｶﾞｺﾏﾝﾄﾞ(10byte)
        <VBFixedArray(5)> Dim wReserve() As Byte            ' 予備(6byte)  
        Dim wMeasurementMode As UShort                      ' 測定モード(0:絶対, 1:偏差) 

        ' この構造体を初期化するには、"Initialize" を呼び出さなければなりません。 
        Public Sub Initialize()
            ReDim strI(39)
            ReDim strT(9)
            ReDim wReserve(5)
        End Sub
    End Structure

    '-------------------------------------------------------------------------------
    '   トリミング要求データ形式定義(VB→INtime)
    '-------------------------------------------------------------------------------
    '----- トリミング要求データ(プレートデータ) -----
    Public Structure TRIM_DAT_PLATE                         ' トリミング要求データ(プレートデータ)形式定義
        Dim cmdNo As UInteger                               ' コマンドNo.(DllTrimFncで設定するので未使用)
        Dim type As UShort                                  ' データタイプ(1:プレートデータ)
        Dim index_reg As UShort                             ' 抵抗データ・インデックス
        Dim index_cut As UShort                             ' カットデータ・インデックス
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET種別(0:TKY, 1:CHIP, 2:NET)
        Dim prmPlate As TRIM_PLATE_DATA                     ' プレートデータ
    End Structure

    '----- トリミング要求データ(GPIB設定用データ) -----
    Public Structure TRIM_DAT_GPIB                          ' トリミング要求データ(GPIB設定用データ)形式定義
        Dim cmdNo As UInteger                               ' コマンドNo.(DllTrimFncで設定するので未使用)
        Dim type As UShort                                  ' データタイプ(8:GPIBデータ)
        Dim index_reg As UShort                             ' 抵抗データ・インデックス
        Dim index_cut As UShort                             ' カットデータ・インデックス
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET種別(0:TKY, 1:CHIP, 2:NET)
        Dim prmGPIB As TRIM_PLATE_GPIB                      ' GPIB設定用データ
    End Structure

    '----- トリミング要求データ(抵抗データ) -----
    Public Structure TRIM_DAT_REGI                          ' トリミング要求データ(抵抗データ)形式定義
        Dim cmdNo As UInteger                               ' コマンドNo.(DllTrimFncで設定するので未使用)
        Dim type As UShort                                  ' データタイプ(2:抵抗データ)
        Dim index_reg As UShort                             ' 抵抗データ・インデックス
        Dim index_cut As UShort                             ' カットデータ・インデックス
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET種別(0:TKY, 1:CHIP, 2:NET)
        Dim prmReg As PRM_REGISTER                          ' 抵抗データ
    End Structure

    '----- トリミング要求データ(カットデータ) -----
    Public Structure TRIM_DAT_CUT                           ' トリミング要求データ(カットデータ)形式定義
        Dim cmdNo As UInteger                               ' コマンドNo.(DllTrimFncで設定するので未使用)
        Dim type As UShort                                  ' データタイプ(3:カットデータ)
        Dim index_reg As UShort                             ' 抵抗データ・インデックス
        Dim index_cut As UShort                             ' カットデータ・インデックス
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET種別(0:TKY, 1:CHIP, 2:NET)
        Dim prmCut As PRM_CUT_DATA                          ' カットデータ
    End Structure

    '----- トリミング要求データ(ST cut/ST cut2パラメータ) -----
    Public Structure TRIM_DAT_CUT_ST                        ' トリミング要求データ(カットパラメータ)形式定義
        Dim cmdNo As UInteger                               ' コマンドNo.(DllTrimFncで設定するので未使用)
        Dim type As UShort                                  ' データタイプ(4:カットパラメータ)
        Dim index_reg As UShort                             ' 抵抗データ・インデックス
        Dim index_cut As UShort                             ' カットデータ・インデックス
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET種別(0:TKY, 1:CHIP, 2:NET)
        Dim c As PRM_CUT_ST                                 ' ST cutパラメータ
    End Structure

    '----- トリミング要求データ(L cutパラメータ) -----
    Public Structure TRIM_DAT_CUT_L                         ' トリミング要求データ(カットパラメータ)形式定義
        Dim cmdNo As UInteger                               ' コマンドNo.(DllTrimFncで設定するので未使用)
        Dim type As UShort                                  ' データタイプ(4:カットパラメータ)
        Dim index_reg As UShort                             ' 抵抗データ・インデックス
        Dim index_cut As UShort                             ' カットデータ・インデックス
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET種別(0:TKY, 1:CHIP, 2:NET)
        Dim c As PRM_CUT_L                                  ' L cutパラメータ
    End Structure

    '----- トリミング要求データ(HOOK cutパラメータ) -----
    Public Structure TRIM_DAT_CUT_HOOK                      ' トリミング要求データ(カットパラメータ)形式定義
        Dim cmdNo As UInteger                               ' コマンドNo.(DllTrimFncで設定するので未使用)
        Dim type As UShort                                  ' データタイプ(4:カットパラメータ)
        Dim index_reg As UShort                             ' 抵抗データ・インデックス
        Dim index_cut As UShort                             ' カットデータ・インデックス
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET種別(0:TKY, 1:CHIP, 2:NET)
        Dim c As PRM_CUT_HOOK                               ' HOOK cutパラメータ
    End Structure

    '----- トリミング要求データ(INDEX cutパラメータ) -----
    Public Structure TRIM_DAT_CUT_INDEX                     ' トリミング要求データ(カットパラメータ)形式定義
        Dim cmdNo As UInteger                               ' コマンドNo.(DllTrimFncで設定するので未使用)
        Dim type As UShort                                  ' データタイプ(4:カットパラメータ)
        Dim index_reg As UShort                             ' 抵抗データ・インデックス
        Dim index_cut As UShort                             ' カットデータ・インデックス
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET種別(0:TKY, 1:CHIP, 2:NET)
        Dim c As PRM_CUT_INDEX                              ' INDEX cutパラメータ
    End Structure

    '----- トリミング要求データ(SCAN cutパラメータ) -----
    Public Structure TRIM_DAT_CUT_SCAN                     ' トリミング要求データ(カットパラメータ)形式定義
        Dim cmdNo As UInteger                               ' コマンドNo.(DllTrimFncで設定するので未使用)
        Dim type As UShort                                  ' データタイプ(4:カットパラメータ)
        Dim index_reg As UShort                             ' 抵抗データ・インデックス
        Dim index_cut As UShort                             ' カットデータ・インデックス
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET種別(0:TKY, 1:CHIP, 2:NET)
        Dim c As PRM_CUT_SCAN                               ' SCAN cutパラメータ
    End Structure

    '----- トリミング要求データ(Letter Markingパラメータ) -----
    Public Structure TRIM_DAT_CUT_MARKING                   ' トリミング要求データ(カットパラメータ)形式定義
        Dim cmdNo As UInteger                               ' コマンドNo.(DllTrimFncで設定するので未使用)
        Dim type As UShort                                  ' データタイプ(4:カットパラメータ)
        Dim index_reg As UShort                             ' 抵抗データ・インデックス
        Dim index_cut As UShort                             ' カットデータ・インデックス
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET種別(0:TKY, 1:CHIP, 2:NET)
        Dim c As PRM_CUT_MARKING                            ' Letter Markingパラメータ
    End Structure

    '----- トリミング要求データ(C cutパラメータ) -----
    Public Structure TRIM_DAT_CUT_C                         ' トリミング要求データ(カットパラメータ)形式定義
        Dim cmdNo As UInteger                               ' コマンドNo.(DllTrimFncで設定するので未使用)
        Dim type As UShort                                  ' データタイプ(4:カットパラメータ)
        Dim index_reg As UShort                             ' 抵抗データ・インデックス
        Dim index_cut As UShort                             ' カットデータ・インデックス
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET種別(0:TKY, 1:CHIP, 2:NET)
        Dim c As PRM_CUT_C                                  ' C cutパラメータ
    End Structure

    '----- トリミング要求データ(ES cutパラメータ) -----
    Public Structure TRIM_DAT_CUT_ES                        ' トリミング要求データ(カットパラメータ)形式定義
        Dim cmdNo As UInteger                               ' コマンドNo.(DllTrimFncで設定するので未使用)
        Dim type As UShort                                  ' データタイプ(4:カットパラメータ)
        Dim index_reg As UShort                             ' 抵抗データ・インデックス
        Dim index_cut As UShort                             ' カットデータ・インデックス
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET種別(0:TKY, 1:CHIP, 2:NET)
        Dim c As PRM_CUT_ES                                 ' ES cutパラメータ
    End Structure

    '----- トリミング要求データ(ES2 cutパラメータ) -----
    Public Structure TRIM_DAT_CUT_ES2                       ' トリミング要求データ(カットパラメータ)形式定義
        Dim cmdNo As UInteger                               ' コマンドNo.(DllTrimFncで設定するので未使用)
        Dim type As UShort                                  ' データタイプ(4:カットパラメータ)
        Dim index_reg As UShort                             ' 抵抗データ・インデックス
        Dim index_cut As UShort                             ' カットデータ・インデックス
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET種別(0:TKY, 1:CHIP, 2:NET)
        Dim c As PRM_CUT_ES2                                ' ES2 cutパラメータ
    End Structure

    '----- トリミング要求データ(Z cut(NOP)パラメータ) -----
    Public Structure TRIM_DAT_CUT_Z                         ' トリミング要求データ(カットパラメータ)形式定義
        Dim cmdNo As UInteger                               ' コマンドNo.(DllTrimFncで設定するので未使用)
        Dim type As UShort                                  ' データタイプ(4:カットパラメータ)
        Dim index_reg As UShort                             ' 抵抗データ・インデックス
        Dim index_cut As UShort                             ' カットデータ・インデックス
        Dim TkyKnd As UShort                                ' TKY/CHIP/NET種別(0:TKY, 1:CHIP, 2:NET)
    End Structure

    '-------------------------------------------------------------------------------
    '   応答データ(トリミング結果データ)形式定義(VB←INtime)
    '-------------------------------------------------------------------------------
    '----- トリミング結果データ(WORD型データ用) -----
    Public Structure TRIM_RESULT_WORD                       ' トリミング結果データ(WORD型データ用)形式定義
        Dim wTxSize As UShort                               ' 転送サイズ(DllTrimFncで設定する)
        <VBFixedArray(cResultMax - 1)> Dim wd() As UShort   ' 結果(wd[0]～wd[255])
        ' この構造体を使用するには"Initialize"を呼び出さなければならない。 
        Public Sub Initialize()
            ReDim wd(cResultMax - 1)
        End Sub
    End Structure

    '----- トリミング結果データ(Double型データ用) -----
    Public Structure TRIM_RESULT_Double                     ' トリミング結果データ(WORD型データ用)形式定義
        Dim wTxSize As UShort                               ' 転送サイズ(DllTrimFncで設定する)
        <VBFixedArray(cResultMax - 1)> Dim dd() As Double   ' 結果(dd[0]～dd[255])
        ' この構造体を使用するには"Initialize"を呼び出さなければならない。 
        Public Sub Initialize()
            ReDim dd(cResultMax - 1)
        End Sub
    End Structure

    '-------------------------------------------------------------------------------
    '   要求/応答データ(コマンド)形式定義(VB←→INtime)
    '-------------------------------------------------------------------------------
    '----- 要求データ(VB→INtime) -----
    Public Structure S_CMD_DAT                              ' 要求データ形式定義
        Dim cmdNo As UInteger                               ' コマンドNo.(DllTrimFncで設定するので未使用)
        <VBFixedArray(9)> Dim dbPara() As Double            ' double 型パラメータ(dbPara(0-9))
        <VBFixedArray(9)> Dim dwPara() As Integer           ' long	 型パラメータ(dbPara(0-9))
        Dim flgTrim As UInteger                             ' 0:ﾄﾘﾐﾝｸﾞ中でない, 1:ﾄﾘﾐﾝｸﾞ中(IRQ0割込禁止) 

        ' この構造体を初期化するには、"Initialize" を呼び出さなければなりません。 
        Public Sub Initialize()
            ReDim dbPara(9)
            ReDim dwPara(9)
        End Sub
    End Structure

    '----- レシオモード２計算式データ(VB→INtime) -----
    Public Structure S_RATIO2EXP                            ' レシオモード２計算式データ形式定義
        Dim cmdNo As UInteger                               ' コマンドNo.(DllTrimFncで設定するので未使用)
        Dim RNO As UInteger                                 ' 抵抗番号
        Dim strExp As String                                ' 計算式文字列
    End Structure

    '----- 応答データ(コマンド)(VB←INtime) -----
    Public Structure S_RES_DAT                              ' 応答データ形式定義
        Dim status As Integer                               ' 0:成功, 0以外:不成功 (※符号なしなのに-1を設定している?)
        Dim dwerrno As Integer                              ' エラー番号(0:正常)
        <VBFixedArray(3)> Dim signal() As UInteger          ' 軸ステータス
        '                                                   ' [0]:X軸
        '                                                   ' [1]:Y軸
        '                                                   ' [2]:Z軸
        '                                                   ' [3]:θ軸
        '                                                   ' I/O入力状態
        <VBFixedArray(INP_MAX - 1)> Dim in_dat() As UInteger
        '                                                   ' [0]:コンソールSWセンス
        '                                                   ' [1]:インターロック関係SWセンス
        '                                                   ' [2]:オートローダLO
        '                                                   ' [3]:オートローダHI
        '                                                   ' [4]:固定アッテネータ
        '                                                   ' I/O出力状態
        <VBFixedArray(OUT_MAX - 1)> Dim outdat() As UInteger
        '                                                   ' [0]:コンソール制御
        '                                                   ' [1]:サーボパワー
        '                                                   ' [2]:オートローダLO
        '                                                   ' [3]:オートローダHI
        '                                                   ' [4]:シグナルタワー(未使用)
        <VBFixedArray(3)> Dim wData() As UInteger           ' TKY戻値
        <VBFixedArray(4)> Dim pos() As Double               ' 現在位置
        '                                                   ' [0]:X軸
        '                                                   ' [1]:Y軸
        '                                                   ' [2]:Z軸
        '                                                   ' [3]:BPX
        '                                                   ' [4]:BPY
        Dim fData As Double                                 ' 戻値(測定値等)

        ' この構造体を初期化するには、"Initialize" を呼び出さなければなりません。 
        Public Sub Initialize()
            ReDim signal(3)
            ReDim in_dat(INP_MAX - 1)
            ReDim outdat(OUT_MAX - 1)
            ReDim pos(4)
        End Sub
    End Structure

    '-------------------------------------------------------------------------------
    '   要求/応答データ定義
    '-------------------------------------------------------------------------------
    Public stSCMD As S_CMD_DAT                              ' 要求データ(コマンド)(VB→INtime)
    Public stSRES As S_RES_DAT                              ' 応答データ(コマンド)(VB←INtime)

    '----- トリミング要求データ(VB→INtime) -----
    Public stTPLT As TRIM_DAT_PLATE                         ' プレートデータ
    Public stTGPI As TRIM_DAT_GPIB                          ' GPIB設定データ
    Public stTREG As TRIM_DAT_REGI                          ' 抵抗データ
    Public stTCUT As TRIM_DAT_CUT                           ' カットデータ
    '                                                       ' カットパラメータ 
    Public stCutST As TRIM_DAT_CUT_ST                       ' ST cutパラメータ
    Public stCutL As TRIM_DAT_CUT_L                         ' L cutパラメータ
    Public stCutHK As TRIM_DAT_CUT_HOOK                     ' HOOK cutパラメータ
    Public stCutIX As TRIM_DAT_CUT_INDEX                    ' INDEX cutパラメータ
    Public stCutSC As TRIM_DAT_CUT_SCAN                     ' SCAN cutパラメータ
    Public stCutMK As TRIM_DAT_CUT_MARKING                  ' Letter Markingパラメータ
    Public stCutC As TRIM_DAT_CUT_C                         ' C cutパラメータ
    Public stCutES As TRIM_DAT_CUT_ES                       ' ES cutパラメータ
    Public stCutE2 As TRIM_DAT_CUT_ES2                      ' ES2 cutパラメータ
    Public stCutZ As TRIM_DAT_CUT_Z                         ' Z cut(NOP)パラメータ

    '----- トリミング結果データ -----
    Public stResultWd As TRIM_RESULT_WORD                   ' トリミング結果データ(WORD型データ用)
    Public stResultDd As TRIM_RESULT_Double                 ' トリミング結果データ(Double型データ用)

    Public gwTrimResult(cResultAry - 1) As UShort           ' 結果(gwTrimResult[0]～gwTrimResult[999])
    Public gfInitialTest(cResultAry - 1) As Double          ' IT測定値(gwTrimResult[0]～gwTrimResult[999])


#End Region

End Module
