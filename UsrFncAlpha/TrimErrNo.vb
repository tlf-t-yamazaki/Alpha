'===============================================================================
'   Description  : エラーコードの定義
'
'   Copyright(C) : TOWA LASERFRONT CORP. 2011
'
'===============================================================================
Option Strict Off
Option Explicit On
Module TrimErrNo
#Region "エラーコードの定義"
    '===============================================================================
    '   軸定義
    '===============================================================================
    '----- 座標軸番号 -----
    Public Const AXIS_X As Short = 0                                    ' X軸
    Public Const AXIS_Y As Short = 1                                    ' Y軸
    Public Const AXIS_Z As Short = 2                                    ' Z軸
    Public Const AXIS_T As Short = 3                                    ' Theta軸
    Public Const AXIS_Z2 As Short = 4                                   ' Z2軸
    Public Const AXIS_XY As Short = 5                                   ' XY軸同時指定

    '----- 座標軸ビット指定 -----
    Public Const STATUS_BIT_AXIS_X As Short = &H1                       ' X軸
    Public Const STATUS_BIT_AXIS_Y As Short = &H2                       ' Y軸
    Public Const STATUS_BIT_AXIS_Z As Short = &H4                       ' Z軸
    Public Const STATUS_BIT_AXIS_T As Short = &H8                       ' Theta軸
    Public Const STATUS_BIT_AXIS_Z2 As Short = &H10                     ' Z2軸

    '----- 全軸のステータスビット(GET_ALLAXIS_STATUS関数のステータスビット) -----
    Public Const BIT_AXIS_X_SERVO_ALM As UShort = &H1000                ' X軸サーボアラーム
    Public Const BIT_AXIS_Y_SERVO_ALM As UShort = &H2000                ' Y軸サーボアラーム
    Public Const BIT_AXIS_Z_SERVO_ALM As UShort = &H4000                ' Z軸サーボアラーム
    Public Const BIT_AXIS_T_SERVO_ALM As UShort = &H8000                ' Theta軸サーボアラーム
    Public Const BIT_AXIS_X_PLUS_LIMIT As UShort = &H1                  ' X軸＋リミット
    Public Const BIT_AXIS_Y_PLUS_LIMIT As UShort = &H2                  ' Y軸＋リミット
    Public Const BIT_AXIS_Z_PLUS_LIMIT As UShort = &H4                  ' Z軸＋リミット
    Public Const BIT_AXIS_T_PLUS_LIMIT As UShort = &H8                  ' Theta軸＋リミット
    Public Const BIT_AXIS_X_MINUS_LIMIT As UShort = &H10                ' X軸－リミット
    Public Const BIT_AXIS_Y_MINUS_LIMIT As UShort = &H20                ' Y軸－リミット
    Public Const BIT_AXIS_Z_MINUS_LIMIT As UShort = &H40                ' Z軸－リミット
    Public Const BIT_AXIS_T_MINUS_LIMIT As UShort = &H80                ' Theta軸－リミット
#End Region

#Region "i-TKY系エラーコードの定義"
    '===============================================================================
    '   i-TKY系/ユーザプロのエラーコード（戻値）定義(0番～100番)
    '   ※OcxSystemと同じ定義とする事
    '===============================================================================
    Public Const cFRS_NORMAL As Short = 0                               ' 正常
    Public Const cFRS_ERR_ADV As Short = 1                              ' OK(ADV(STARTｷｰ))← START/RESET待ち時
    Public Const cFRS_ERR_START As Short = 1                            ' OK(STARTｷｰ)     ← START/RESET待ち時
    Public Const cFRS_ERR_RST As Short = 3                              ' Cancel(RESETｷｰ) ← START/RESET待ち時
    Public Const cFRS_ERR_HALT As Short = 2                             ' HALTｷｰ
    Public Const cFRS_ERR_Z As Short = 4                                ' Zｷｰ
    Public Const cFRS_TxTy As Short = 5                                 ' TX2/TY2押下
    Public Const cFRS_ERR_LOTEND As Short = 6                           ' LOTEND用   

    Public Const cFRS_ERR_CVR As Short = -1                             ' 筐体カバー開検出
    Public Const cFRS_ERR_SCVR As Short = -2                            ' スライドカバー開検出
    Public Const cFRS_ERR_LATCH As Short = -3                           ' カバー開ラッチ検出

    Public Const cFRS_ERR_EMG As Short = -11                            ' 非常停止
    Public Const cFRS_ERR_DUST As Short = -12                           ' 集塵機異常検出
    Public Const cFRS_ERR_AIR As Short = -13                            ' エアー圧エラー検出
    Public Const cFRS_ERR_MVC As Short = -14                            ' ﾏｽﾀｰﾊﾞﾙﾌﾞ回路状態エラー検出
    Public Const cFRS_ERR_HW As Short = -15                             ' ハードウェアエラー検出
    Public Const cFRS_ERR_LDR As Short = -16                            ' ローダアラーム検出          
    Public Const cFRS_ERR_LDR1 As Short = -17                           ' ローダアラーム検出(全停止異常)
    Public Const cFRS_ERR_LDR2 As Short = -18                           ' ローダアラーム検出(サイクル停止) 
    Public Const cFRS_ERR_LDR3 As Short = -19                           ' ローダアラーム検出(軽故障)
    Public Const cFRS_ERR_LDRTO As Short = -20                          ' ローダ通信タイムアウト          

    '----- IO制御タイムアウト(OcxSystemで使用) -----
    Public Const cFRS_TO_SCVR_CL As Short = -21                         ' タイムアウト(スライドカバー閉待ち)
    Public Const cFRS_TO_SCVR_OP As Short = -22                         ' タイムアウト(スライドカバー開待ち)
    Public Const cFRS_TO_SCVR_ON As Short = -23                         ' タイムアウト(ｽﾗｲﾄﾞｶﾊﾞｰｽﾄｯﾊﾟｰ行待ち)
    Public Const cFRS_TO_SCVR_OFF As Short = -24                        ' タイムアウト(ｽﾗｲﾄﾞｶﾊﾞｰｽﾄｯﾊﾟｰ戻待ち)
    Public Const cFRS_TO_CLAMP_ON As Short = -25                        ' タイムアウト(クランプＯＮ)
    Public Const cFRS_TO_CLAMP_OFF As Short = -26                       ' タイムアウト(クランプＯＦＦ)
    Public Const cFRS_TO_PM_DW As Short = -27                           ' タイムアウト(パワーメータ下降移動)
    Public Const cFRS_TO_PM_UP As Short = -28                           ' タイムアウト(パワーメータ上昇移動)
    Public Const cFRS_TO_PM_FW As Short = -29                           ' タイムアウト(パワーメータ測定端移動)
    Public Const cFRS_TO_PM_BK As Short = -30                           ' タイムアウト(パワーメータ待機端移動)

    '----- ファイル入出力エラー -----
    Public Const cFRS_FIOERR_INP As Short = -31                         ' ファイル入力エラー
    Public Const cFRS_FIOERR_OUT As Short = -32                         ' ファイル出力エラー
    Public Const cFRS_COM_ERR As Short = -33                            ' 通信エラー(PLC)
    Public Const cFRS_COM_FL_ERR As Short = -34                         ' 通信エラー(FL)
    Public Const cFRS_FL_ERR_INP As Short = -35                         ' ＦＬ側に加工条件の設定がない
    Public Const cFRS_FL_ERR_SET As Short = -36                         ' 加工条件番号設定エラー

    '----- Video.OCXのエラー -----
    Public Const cFRS_VIDEO_PTN As Short = -40                          ' パターン認識エラー
    Public Const cFRS_VIDEO_PT1 As Short = -41                          ' パターン認識エラー(補正位置1)
    Public Const cFRS_VIDEO_PT2 As Short = -42                          ' パターン認識エラー(補正位置2)
    Public Const cFRS_VIDEO_COM As Short = -43                          ' 通信エラー(CV3000)
    Public Const cFRS_VIDEO_INI As Short = -44                          ' 初期化が行われていません
    Public Const cFRS_VIDEO_IN2 As Short = -45                          ' 初期化済み
    Public Const cFRS_VIDEO_FRM As Short = -46                          ' フォーム表示中
    Public Const cFRS_VIDEO_PRP As Short = -47                          ' プロパティ値不正
    Public Const cFRS_VIDEO_GRP As Short = -48                          ' ﾃﾝﾌﾟﾚｰﾄｸﾞﾙｰﾌﾟ番号ｴﾗｰ
    Public Const cFRS_VIDEO_MXT As Short = -49                          ' テンプレート数 > MAX
    Public Const cFRS_VIDEO_UXP As Short = -50                          ' 予期せぬエラー
    Public Const cFRS_VIDEO_UX2 As Short = -51                          ' 予期せぬエラー2
    Public Const cFRS_MVC_UTL As Short = -52                            ' MvcUtil エラー
    Public Const cFRS_MVC_PT2 As Short = -53                            ' MvcPt2 エラー
    Public Const cFRS_MVC_10 As Short = -54                             ' Mvc10 エラー

    '----- その他のエラー -----
    Public Const cFRS_ERR_SHUTTER As Short = -60                        ' 外部シャッターエラー検出
    Public Const cFRS_ERR_PLC As Short = -61                            ' PLCステータス異常
    Public Const cFRS_TO_EXLOCK As Short = -62                          ' 電磁ロックタイムアウト   'V2.2.0.0⑤

    '----- 画面処理用の戻り値 -----                                     ' -80番代と-90番代はUsrpro,TKY等の画面処理で使用
    Public Const cFRS_ERR_HING As Short = -80                           ' 連続NG-HIGHｴエラー発生
    Public Const cFRS_ERR_REPROB As Short = -81                         ' 再プロービング失敗
    Public Const cFRS_FNG_DATA As Short = -85                           ' データ未ロード
    Public Const cFRS_FNG_CMD As Short = -86                            ' 他コマンド実行中
    Public Const cFRS_FNG_PASS As Short = -87                           ' ﾊﾟｽﾜｰﾄﾞ入力ｴﾗｰ
    Public Const cFRS_FNG_CPOS As Short = -88                           ' カット位置補正対象の抵抗がない

    Public Const cFRS_TRIM_NG As Short = -90                            ' トリミングNG
    Public Const cFRS_ERR_TRIM As Short = -91                           ' トリマエラー
    Public Const cFRS_ERR_PTN As Short = -92                            ' パターン認識エラー(パターンが見つからなかった)
    Public Const cFRS_ERR_PT2 As Short = -93                            ' パターン認識エラー(閾値エラー)

    '----- その他のエラー -----
    Public Const cERR_END As Short = -999                               ' エラー発生
    Public Const cERR_TRAP As Short = -99                               ' トラップエラー発生
    Public Const cFRS_VACCUME_ERROR As Short = -100                     ' 吸着エラー

#End Region

#Region "INtime側エラーコード定義"
    '===============================================================================
    '   INtime側エラーコード定義(100番以降)
    '   (※)OcxSystemからはINtime側エラー(100番以降)は下記のエラーコード× -1 で返す
    '===============================================================================
    '-----( エラーステータス )-----
    '  トリムブロック処理時のエラー (LastErrorに設定し、dwerrnoでWindows側へ返信)
    Public Const ERROR_SUCCESS As Integer = 0                           'ユーザファンクション正常リターン値

    Public Const ERR_CLEAR As Short = 0                                 ' リミット状態クリア（リミットエラーなど）
    Public Const ERR_ALREADY_SET As Short = 1
    Public Const ERR_TERMINATE As Short = 2                             ' ホストより中断終了
    Public Const ERR_RESET As Short = 3                                 ' リセット終了
    Public Const ERR_CONTACT As Short = 5                               ' コンタクトエラー

    ' ソフトウェアエラー
    'Public Const ERR_EMGSWCH As Short = 102                            ' 非常停止
    'Public Const ERR_SRV_ALM As Short = 103                            ' サーボアラーム
    'Public Const ERR_AXS_LIM As Short = 104                            ' XYZ軸リミット検出

    'ステージエラー
    '  タイムアウトエラーについては、ベース+対称軸ビット(X:0x01,Y:0x02,Z:0x04,T:0x08,Z2:0x10)
    '  のエラーコードで返す。
    'Public Const ERR_TIMEOUT_BASE As Short = 100                       ' タイムアウトエラー(ベース番号)
    ' X軸タイムアウトAs Short =101
    ' Y軸タイムアウトAs Short =102
    ' Z軸タイムアウトAs Short =104
    ' Theta軸タイムアウトAs Short =108
    ' Z2軸タイムアウトAs Short =116
    ' 複数軸はorの値
    ' X,Y軸タイムアウトAs Short =103

    'Public Const ERR_TIMEOUT_X As Short = 101                          ' X軸タイムアウト
    'Public Const ERR_TIMEOUT_Y As Short = 102                          ' Y軸タイムアウト
    'Public Const ERR_TIMEOUT_Z As Short = 103                          ' Z軸タイムアウト
    'Public Const ERR_TIMEOUT_T As Short = 104                          ' θ軸タイムアウト
    'Public Const ERR_SOFTLIMIT_X As Short = 105                        ' X軸ソフトリミット
    'Public Const ERR_SOFTLIMIT_Y As Short = 106                        ' Y軸ソフトリミット
    'Public Const ERR_SOFTLIMIT_Z As Short = 107                        ' Z軸ソフトリミット
    'Public Const ERR_TIMEOUT_ATT As Short = 108                        ' ATT ← タイムアウト下に移動
    'Public Const ERR_TIMEOUT_Z2 As Short = 109                         ' Z2軸タイムアウト
    'Public Const ERR_BP_XLIMIT As Short = 110                          ' BP X軸可動範囲オーバー（未使用）
    'Public Const ERR_BP_YLIMIT As Short = 111                          ' BP Y軸可動範囲オーバー(未使用）
    'Public Const ERR_BP_MOVE_TIMEOUT As Short = 112                    ' BP タイムアウト

    '----- ソフトリミットエラー(OcxSystemで使用) -----
    Public Const ERR_SOFTLIMIT_X As Short = 105                         ' X軸ソフトリミット
    Public Const ERR_SOFTLIMIT_Y As Short = 106                         ' Y軸ソフトリミット
    Public Const ERR_SOFTLIMIT_Z As Short = 107                         ' Z軸ソフトリミット
    Public Const ERR_SOFTLIMIT_Z2 As Short = 113                        ' Z2軸ソフトリミット
    Public Const ERR_BP_XLIMIT As Short = 110                           ' BP X軸ソフトリミットエラー
    Public Const ERR_BP_YLIMIT As Short = 111                           ' BP Y軸ソフトリミットエラー

    Public Const ERR_INTIME_BASE As Short = 100                         ' INtime側エラーベース番号(INtime側エラーコードの最小値)
    ' 200にしたいがERR_TIMEOUT_ATT（108）がいきている?
    '----- ハードウェアエラー -----
    Public Const ERR_EMGSWCH As Short = 201                             ' 非常停止
    Public Const ERR_SRV_ALM As Short = 202                             ' サーボアラーム
    'Public Const ERR_AXS_LIM As Short = 203                            ' XYZ軸リミット検出 （実際には未使用）
    Public Const ERR_AXS_LIM_X As Short = 203                           ' X軸リミット検出
    Public Const ERR_AXS_LIM_Y As Short = 204                           ' Y軸リミット検出
    Public Const ERR_AXS_LIM_Z As Short = 205                           ' Z軸リミット検出
    Public Const ERR_AXS_LIM_T As Short = 206                           ' θ軸リミット検出
    Public Const ERR_AXS_LIM_ATT As Short = 207                         ' ATT軸リミット検出
    Public Const ERR_AXS_LIM_Z2 As Short = 208                          ' Z2軸リミット検出

    Public Const ERR_OPN_CVR As Short = 209                             ' 筐体カバー開検出
    Public Const ERR_OPN_SCVR As Short = 210                            ' スライドカバー開検出
    Public Const ERR_OPN_CVRLTC As Short = 211                          ' カバー開ラッチ検出
    Public Const ERR_AXS_LIM As Short = 220                             ' 各軸のリミット検出(リミット到達中)

    '----- ロジカルエラー -----
    Public Const ERR_CMD_NOTSPT As Short = 301                          ' 未サポートコマンド
    Public Const ERR_CMD_PRM As Short = 302                             ' パラメータエラー
    Public Const ERR_CMD_LIM_L As Short = 303                           ' パラメータ下限値エラー
    Public Const ERR_CMD_LIM_U As Short = 304                           ' パラメータ上限値エラー
    Public Const ERR_RT2WIN_SEND As Short = 305                         ' INTime→Windows送信エラー
    Public Const ERR_RT2WIN_RECV As Short = 306                         ' INTime→Windows受信エラー
    Public Const ERR_WIN2RT_SEND As Short = 307                         ' Windows→INTime送信エラー
    Public Const ERR_WIN2RT_RECV As Short = 308                         ' Windows→INTime受信エラー
    Public Const ERR_SYS_BADPOINTER As Short = 309                      ' 不正ポインタ
    Public Const ERR_SYS_FREE_MEMORY As Short = 310                     ' メモリ領域の開放エラー
    Public Const ERR_SYS_ALLOC_MEMORY As Short = 311                    ' メモリ領域の確保エラー
    Public Const ERR_CALC_OVERFLOW As Short = 320                       ' オーバーフロー
    Public Const ERR_INTIME_NOTMOVE As Short = 350                      ' INTRIMが起動していない

    '----- ハードウェアIO系エラーコード：400番台 -----
    Public Const ERR_CSLLAMP_SETNO As Short = 401                       ' コンソール系ランプ番号指定エラー
    Public Const ERR_SIGTWRLAMP_SETNO As Short = 402                    ' シグナルタワーランプ番号指定エラー
    Public Const ERR_SIGTWRLAMP_SETMODENO As Short = 402                ' シグナルタワーランプ点灯/点滅モード指定エラー
    Public Const ERR_BIT_ONOFF As Short = 403                           ' ビットオン/オフ指定エラー
    Public Const ERR_SETPRM_WAITSWNON As Short = 404                    ' 入力待ち対象スイッチの指定が全てなしになっている。

    '----- 測定系エラーコード：500番台 -----
    Public Const ERR_MEAS_RANGESET_TYPE As Short = 501                  ' 測定レンジ設定エラー：指定レンジ設定タイプなし
    Public Const ERR_MEAS_SETRNG_NO As Short = 502                      ' 測定レンジ設定エラー：対象レンジなし
    Public Const ERR_MEAS_SETRNG_LO As Short = 503                      ' 測定レンジ設定エラー：最小レンジ以下
    Public Const ERR_MEAS_SETRNG_HI As Short = 504                      ' 測定レンジ設定エラー：最大レンジ以上
    Public Const ERR_MEAS_RNG_NOTSET As Short = 505                     ' 測定レンジ設定エラー：レンジ未設定
    Public Const ERR_MEAS_FAIL As Short = 506                           ' 測定エラー
    Public Const ERR_MEAS_FAST_R As Short = 507                         ' 高速度抵抗測定エラー
    Public Const ERR_MEAS_HIGHPRECI_R As Short = 508                    ' 高精度抵抗測定エラー
    Public Const ERR_MEAS_FAST_V As Short = 509                         ' 高速度電圧測定エラー
    Public Const ERR_MEAS_HIGHPRECI_V As Short = 510                    ' 高精度電圧測定エラー
    Public Const ERR_MEAS_TARGET As Short = 511                         ' 測定目標値設定エラー
    Public Const ERR_MEAS_TARGET_LO As Short = 512                      ' 測定目標値設定エラー：最小目標値以下
    Public Const ERR_MEAS_TARGET_HI As Short = 513                      ' 測定目標値設定エラー：最大目標値以下
    Public Const ERR_MEAS_SCANNER As Short = 514                        ' 測定スキャナ設定エラー：不正スキャナ番号
    Public Const ERR_MEAS_SCANNER_LO As Short = 515                     ' 測定スキャナ設定エラー：最小スキャナ番号以下
    Public Const ERR_MEAS_SCANNER_HI As Short = 516                     ' 測定スキャナ設定エラー：最大スキャナ番号以上
    Public Const ERR_MEAS_RNG_SHORT As Short = 517                      ' 測定値：レンジショート（0.01以下）
    Public Const ERR_MEAS_RNG_OVER As Short = 518                       ' 測定値：レンジオーバー（67000000.0以上）
    Public Const ERR_MEAS_SPAN As Short = 519                           ' 測定範囲外
    Public Const ERR_MEAS_SPAN_SHORT As Short = 520                     ' 測定範囲外：ショート-抵抗(0x6666)/電圧(0x3333)以下
    Public Const ERR_MEAS_SPAN_OVER As Short = 521                      ' 測定範囲外：オーバー-抵抗(0xCCCC)/電圧(0x6666)以上
    Public Const ERR_MEAS_INVALID_SLOPE As Short = 522                  ' スロープ設定エラー
    Public Const ERR_MEAS_COUNT As Short = 523                          ' 測定回数指定エラー
    Public Const ERR_MEAS_SETMODE As Short = 524                        ' 測定時の設定モード（Mfsetモード）エラー

    Public Const ERR_MEAS_AUTORNG_OVER As Short = 525                   ' オートレンジ測定:定電流測定範囲オーバー（差電流領域）
    Public Const ERR_MEAS_K2VAL_SHORT As Short = 526                    ' 差電流測定:K2設定値0.4未満
    Public Const ERR_MEAS_K2VAL_OVER As Short = 527                     ' 差電流測定:K2設定値0.8以上
    Public Const ERR_MEAS_SCANSET_TIMEOUT As Short = 528                ' スキャナ設定完了タイムアウト
    Public Const ERR_MEAS_RANGESET_TIMEOUT As Short = 529               ' レンジ設定回数をオーバした場合のタイムアウト

    '----- BP系エラーコード：600番台 -----
    Public Const ERR_BP_MAXLINEANUM_LO As Short = 601                   ' リニアリティ補正：最小番号以下
    Public Const ERR_BP_MAXLINEANUM_HI As Short = 602                   ' リニアリティ補正：最大番号以上
    Public Const ERR_BP_LOGICALNUM_LO As Short = 603                    ' 象限設定：最小象限値番号以下
    Public Const ERR_BP_LOGICALNUM_HI As Short = 604                    ' 象限設定：最大象限値番号以下
    Public Const ERR_BP_LIMITOVER As Short = 605                        ' BP移動距離設定：リミットオーバー
    Public Const ERR_BP_HARD_LIMITOVER_LO As Short = 606                ' BP移動距離設定：リミットオーバー（最小可動範囲以下）
    Public Const ERR_BP_HARD_LIMITOVER_HI As Short = 607                ' BP移動距離設定：リミットオーバー（最大可動範囲以上）
    Public Const ERR_BP_SOFT_LIMITOVER As Short = 608                   ' ソフト可動範囲オーバー
    Public Const ERR_BP_BSIZE_OVER As Short = 609                       ' ブロックサイズ設定オーバー（ソフト可動範囲オーバー）
    Public Const ERR_BP_SIZESET As Short = 610                          ' BPサイズ設定エラー
    Public Const ERR_BP_MOVE_TIMEOUT As Short = 611                     ' BP タイムアウト
    Public Const ERR_BP_GRV_ALARM_X As Short = 620                      ' ガルバノアラームX
    Public Const ERR_BP_GRV_ALARM_Y As Short = 621                      ' ガルバノアラームY
    Public Const ERR_BP_GRVMOVE_HARDERR As Short = 622                  ' ポジショニング動作時のガルバノ動作異常（指令値と現在値不一致）
    Public Const ERR_BP_GRVSET_AXIS As Short = 630                      ' ポジショニング動作時の座標設定or座標取得エラー
    Public Const ERR_BP_GRVSET_MOVEMODE As Short = 631                  ' ガルバノモード設定時の動作モード設定値不正（ポジショニングAs Short =0、トリミングAs Short =1以外の値が設定されている）
    Public Const ERR_BP_GRVSET_SPEEDSHORT As Short = 632                ' ガルバノ速度指定：最小スピード未満
    Public Const ERR_BP_GRVSET_SPEEDOVER As Short = 633                 ' ガルバノ速度指定：最大スピードオーバー

    '----- トリミング/カット系エラーコード:700番台 -----
    Public Const ERR_CUT_NOT_SUPPORT As Short = 701                     ' 未サポートカット形状
    Public Const ERR_CUT_PARAM_LEN As Short = 702                       ' カットパラメータエラー：カット長設定エラー
    Public Const ERR_CUT_PARAM_LEN_LO As Short = 703                    ' カットパラメータエラー：カット長最小値以下
    Public Const ERR_CUT_PARAM_LEN_HI As Short = 704                    ' カットパラメータエラー：カット長最大値以上
    Public Const ERR_CUT_PARAM_CORR As Short = 706                      ' カットパラメータエラー：パラメータ相関エラー
    Public Const ERR_CUT_PARAM_SPD As Short = 707                       ' カットパラメータエラー：スピード設定エラー
    Public Const ERR_CUT_PARAM_SPD_LO As Short = 708                    ' カットパラメータエラー：スピード設定オーバー
    Public Const ERR_CUT_PARAM_SPD_HI As Short = 709                    ' カットパラメータエラー：スピード設定ショート
    Public Const ERR_CUT_PARAM_DIR As Short = 710                       ' カットパラメータエラー：カット方向
    Public Const ERR_CUT_PARAM_CUTCNT As Short = 711                    ' カットパラメータエラー：カット本数
    Public Const ERR_CUT_PARAM_CHRSIZE As Short = 712                   ' カットパラメータエラー：文字サイズ指定エラー
    Public Const ERR_CUT_PARAM_CUTANGLE As Short = 713                  ' カットパラメータエラー：角度指定エラー
    Public Const ERR_CUT_PARAM_CHARSET As Short = 714                   ' カットパラメータエラー：文字列指定エラー
    Public Const ERR_CUT_PARAM_STRLEN As Short = 715                    ' カットパラメータエラー：マーキング文字列長指定エラー
    Public Const ERR_CUT_PARAM_NOHAR As Short = 716                     ' カットパラメータエラー：指定文字なし
    Public Const ERR_CUT_PARAM_NORES As Short = 717                     ' カットパラメータエラー：抵抗番号不正
    Public Const ERR_CUT_PARAM_CUTMODE As Short = 718                   ' カットパラメータエラー：カットモード不正（CUT_MODE_NORMAL(1)～CUT_MODE_NANAME(4)以外指定）
    Public Const ERR_CUT_PARAM_PARCENT As Short = 719                   ' カットパラメータエラー：ポイント設定
    Public Const ERR_CUT_PARAM_BP As Short = 720

    Public Const ERR_CUT_RATIOPRM_BASERNO As Short = 730                ' レシオパラメータエラー：計算用ベース抵抗番号指定エラー
    Public Const ERR_CUT_RATIOPRM_BASER_NG As Short = 731               ' レシオパラメータエラー：計算用ベース抵抗判定NG
    Public Const ERR_CUT_RATIOPRM_MODENOT2 As Short = 732               ' レシオパラメータエラー：計算用ベース抵抗がRatioモード2でない
    Public Const ERR_CUT_RATIOPRM_CALCFORM As Short = 733               ' レシオパラメータエラー：計算用式が不正
    Public Const ERR_CUT_RATIOPRM_CANTSET As Short = 734                ' レシオパラメータエラー：データの登録が出来ない
    Public Const ERR_CUT_RATIOPRM_GRPNO As Short = 735                  ' レシオパラメータエラー：グループ番号指定エラー

    Public Const ERR_CUT_UCUTPRM_NOPRMRES As Short = 740                ' Uカットパラメータエラー：抵抗情報パラメータエリアの設定なし
    Public Const ERR_CUT_UCUTPRM_NOPRMCUT As Short = 741                ' Uカットパラメータエラー：カット情報パラメータエリアの設定なし
    Public Const ERR_CUT_UCUTPRM_NORES As Short = 742                   ' Uカットパラメータエラー：対象抵抗なし
    Public Const ERR_CUT_UCUTPRM_NODATA As Short = 743                  ' Uカットパラメータエラー：パラメータデータが設定されていない
    Public Const ERR_CUT_UCUTPRM_NOMODE As Short = 744                  ' Uカットパラメータエラー：指定モードなし
    Public Const ERR_CUT_UCUTPRM_TBLIDX As Short = 745                  ' Uカットパラメータエラー：テーブルのインデックスが不正

    Public Const ERR_CUT_CIRCUIT_NO As Short = 750                      ' サーキットパラメータエラー：サーキット番号指定エラー

    Public Const ERR_TRIMRESULT_TESTNO As Short = 760                   ' トリミング結果：取得対象結果番号指定エラー（テスト実施数より大きい）
    Public Const ERR_TRIMRESULT_RESNO As Short = 761                    ' トリミング結果：取得対象抵抗開始番号指定エラー（登録抵抗数より大きい開始番号）
    Public Const ERR_TRIMRESULT_TOTALRESNO As Short = 762               ' トリミング結果：取得対象抵抗番号指定エラー（登録抵抗数より大きい）
    Public Const ERR_TRIMRESULT_CUTNO As Short = 763                    ' トリミング結果：取得対象カット開始番号指定エラー（最大カット数より大きい開始番号）
    Public Const ERR_TRIMRESULT_TOTALCUTNO As Short = 764               ' トリミング結果：取得対象カット番号指定エラー（最大カット数より大きい開始番号）
    Public Const ERR_TRIMRESULT_ESRESCUTNO As Short = 765               ' トリミング結果：取得対象抵抗番号orカット番号指定エラー（最大カット数より大きい開始番号）

    Public Const ERR_TRIM_COTACTCHK As Short = 770                      ' コンタクトチェックエラー
    Public Const ERR_TRIM_OPNSRTCHK As Short = 771                      ' オープンショートチェックエラー
    Public Const ERR_TRIM_CIR_NORES As Short = 772                      ' サーキットトリミング-次対象抵抗なし

    Public Const ERR_MARKVFONT_ANGLE As Short = 780                     ' フォント設定：角度指定エラー
    Public Const ERR_MARKVFONT_MAXPOINT As Short = 781                  ' フォント設定：ポイント数オーバー

    '----- レーザ系エラーコード:800番台 -----
    Public Const ERR_LSR_PARAM_ONOFF As Short = 801                     ' レーザON/OFFパラメータエラー
    Public Const ERR_LSR_PARAM_QSW As Short = 802                       ' Qレートパラメータエラー
    Public Const ERR_LSR_PARAM_QSW_LO As Short = 803                    ' Qレートパラメータエラー：最小Qレート以下
    Public Const ERR_LSR_PARAM_QSW_HI As Short = 804                    ' Qレートパラメータエラー：最小Qレート以上
    Public Const ERR_LSR_PARAM_FLCUTCND As Short = 830                  ' FL用加工条件番号設定エラー
    Public Const ERR_LSR_PARAM_FLPRM As Short = 831                     ' FL用パラメータエラー（制御種別指定エラー）
    Public Const ERR_LSR_PARAM_FLTIMEOUT As Short = 832                 ' FL用設定タイムアウト
    Public Const ERR_LSR_STATUS_STANBY As Short = 833                   ' FL:Stanby error,ES:POWER ON READY error
    Public Const ERR_LSR_STATUS_OSCERR As Short = 850                   ' FL:Error occured,:ES:LD Alarm

    '----- ステージ系エラーコード: -----
    Public Const ERR_STG_STATUS As Short = 901                          ' ステージ系でステータスエラー発生
    Public Const ERR_STG_NOT_EXIST As Short = 902                       ' 対象ステージ(X/Y/Z/Z2/Theta)存在せず。
    Public Const ERR_STG_ORG_NONCOMPLETION As Short = 903               ' 原点復帰未完了
    'Public Const ERR_SOFTLIMIT_X As Short = 105                        ' X軸ソフトリミット
    'Public Const ERR_SOFTLIMIT_Y As Short = 106                        ' Y軸ソフトリミット
    'Public Const ERR_SOFTLIMIT_Z As Short = 107                        ' Z軸ソフトリミット
    Public Const ERR_STG_MOVESEQUENCE As Short = 905                    ' ステージ動作シーケンス設定エラー
    Public Const ERR_STG_THETANOTEXIST As Short = 906                   ' θ載物台なし
    Public Const ERR_STG_INIT_X As Short = 910                          ' X軸原点復帰エラー
    Public Const ERR_STG_INIT_Y As Short = 911                          ' Y軸原点復帰エラー
    Public Const ERR_STG_INIT_Z As Short = 912                          ' Z軸原点復帰エラー
    Public Const ERR_STG_INIT_T As Short = 913                          ' Theta軸原点復帰エラー
    Public Const ERR_STG_INIT_Z2 As Short = 914                         ' Z2軸原点復帰エラー
    Public Const ERR_STG_MOVE_X As Short = 915                          ' X軸ステージ動作異常：指令値 !As Short = 移動位置
    Public Const ERR_STG_MOVE_Y As Short = 916                          ' Y軸ステージ動作異常：指令値 !As Short = 移動位置
    Public Const ERR_STG_MOVE_Z As Short = 917                          ' Z軸ステージ動作異常：指令値 !As Short = 移動位置
    Public Const ERR_STG_MOVE_T As Short = 918                          ' Theta軸ステージ動作異常：指令値 !As Short = 移動位置
    Public Const ERR_STG_MOVE_Z2 As Short = 919                         ' Z2軸ステージ動作異常：指令値 !As Short = 移動位置

    Public Const ERR_STG_SOFTLMT_PLUS As Short = 930                    ' プラスリミットオーバー
    Public Const ERR_STG_SOFTLMT_MINUS As Short = 931                   ' マイナスリミットオーバー

    '----- タイムアウトエラー -----
    '  タイムアウトエラーについては、ベース+対称軸ビット(X:0x01,Y:0x02,Z:0x04,T:0x08,Z2:0x10)
    '  のエラーコードで返す。
    Public Const ERR_TIMEOUT_BASE As Short = 950
    ' タイムアウトエラー(ベース番号)
    ' X軸タイムアウトAs Short =951
    ' Y軸タイムアウトAs Short =952
    ' Z軸タイムアウトAs Short =954
    ' Theta軸タイムアウトAs Short =958
    ' Z2軸タイムアウトAs Short =966
    ' 複数軸はorの値
    ' X,Y軸タイムアウトAs Short =953

    '----- 軸系エラーベース番号 -----
    Public Const ERR_AXIS_X_BASE As Short = 1000                        ' X軸系エラーベース番号
    Public Const ERR_AXIS_Y_BASE As Short = 2000                        ' Y軸系エラーベース番号
    Public Const ERR_AXIS_Z_BASE As Short = 3000                        ' Z軸系エラーベース番号
    Public Const ERR_AXIS_T_BASE As Short = 4000                        ' Theta軸系エラーベース番号
    Public Const ERR_AXIS_Z2_BASE As Short = 5000                       ' Z2軸系エラーベース番号

    '----- タイムアウト -----
    Public Const ERR_TIMEOUT_AXIS_X As Short = ERR_TIMEOUT_BASE + STATUS_BIT_AXIS_X                      ' X軸タイムアウト
    Public Const ERR_TIMEOUT_AXIS_Y As Short = ERR_TIMEOUT_BASE + STATUS_BIT_AXIS_Y                      ' Y軸タイムアウト
    Public Const ERR_TIMEOUT_AXIS_Z As Short = ERR_TIMEOUT_BASE + STATUS_BIT_AXIS_Z                      ' Z軸タイムアウト
    Public Const ERR_TIMEOUT_AXIS_T As Short = ERR_TIMEOUT_BASE + STATUS_BIT_AXIS_T                      ' θ軸タイムアウト
    Public Const ERR_TIMEOUT_AXIS_Z2 As Short = ERR_TIMEOUT_BASE + STATUS_BIT_AXIS_Z2                    ' Z2軸タイムアウト
    Public Const ERR_TIMEOUT_AXIS_XY As Short = ERR_TIMEOUT_BASE + STATUS_BIT_AXIS_X + STATUS_BIT_AXIS_Y ' XY軸タイムアウト
    Public Const ERR_TIMEOUT_ATT As Short = 108                                                          ' ロータリアッテネータタイムアウト

    '----- サーボアラーム -----
    Public Const ERR_AXIS_X_SERVO_ALM As Short = ERR_AXIS_X_BASE + 0                                     ' X軸サーボアラーム
    Public Const ERR_AXIS_Y_SERVO_ALM As Short = ERR_AXIS_Y_BASE + 0                                     ' Y軸サーボアラーム
    Public Const ERR_AXIS_Z_SERVO_ALM As Short = ERR_AXIS_Z_BASE + 0                                     ' Z軸サーボアラーム
    Public Const ERR_AXIS_T_SERVO_ALM As Short = ERR_AXIS_T_BASE + 0                                     ' θ軸サーボアラーム

    '----- 軸プラスリミットオーバー -----
    Public Const ERR_STG_SOFTLMT_PLUS_AXIS_X As Short = ERR_AXIS_X_BASE + ERR_STG_SOFTLMT_PLUS           ' X軸プラスリミットオーバー
    Public Const ERR_STG_SOFTLMT_PLUS_AXIS_Y As Short = ERR_AXIS_Y_BASE + ERR_STG_SOFTLMT_PLUS           ' Y軸プラスリミットオーバー
    Public Const ERR_STG_SOFTLMT_PLUS_AXIS_Z As Short = ERR_AXIS_Z_BASE + ERR_STG_SOFTLMT_PLUS           ' Z軸プラスリミットオーバー
    Public Const ERR_STG_SOFTLMT_PLUS_AXIS_T As Short = ERR_AXIS_T_BASE + ERR_STG_SOFTLMT_PLUS           ' θ軸プラスリミットオーバー
    Public Const ERR_STG_SOFTLMT_PLUS_AXIS_Z2 As Short = ERR_AXIS_Z2_BASE + ERR_STG_SOFTLMT_PLUS         ' Z2軸プラスリミットオーバー

    '----- 軸マイナスリミットオーバー -----
    Public Const ERR_STG_SOFTLMT_MINUS_AXIS_X As Short = ERR_AXIS_X_BASE + ERR_STG_SOFTLMT_MINUS         ' X軸マイナスリミットオーバー
    Public Const ERR_STG_SOFTLMT_MINUS_AXIS_Y As Short = ERR_AXIS_Y_BASE + ERR_STG_SOFTLMT_MINUS         ' Y軸マイナスリミットオーバー
    Public Const ERR_STG_SOFTLMT_MINUS_AXIS_Z As Short = ERR_AXIS_Z_BASE + ERR_STG_SOFTLMT_MINUS         ' Z軸マイナスリミットオーバー
    Public Const ERR_STG_SOFTLMT_MINUS_AXIS_T As Short = ERR_AXIS_T_BASE + ERR_STG_SOFTLMT_MINUS         ' θ軸マイナスリミットオーバー
    Public Const ERR_STG_SOFTLMT_MINUS_AXIS_Z2 As Short = ERR_AXIS_Z2_BASE + ERR_STG_SOFTLMT_MINUS       ' Z2軸マイナスリミットオーバー

    '----- IO系エラーコード: -----
    Public Const ERR_IO_PCINOTFOUND As Short = 10000                    ' PCIボードが検出できず
    Public Const ERR_IO_NOTGET_RTCCMOSDATA As Short = 10001             ' RTCﾃﾞｰﾀ読み込み失敗

    '----- GPIB系エラーコード -----
    Public Const ERR_GPIB_PARAM As Short = 11001                        ' GPIBパラメータエラー
    Public Const ERR_GPIB_TCPSOCKET As Short = 11002                    ' GPIB-TCP/IP送信エラー
    Public Const ERR_GPIB_EXEC As Short = 11003                         ' GPIBコマンド実行エラー

#End Region

End Module