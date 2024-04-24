
Imports LaserFront.Trimmer.DllSystem
Imports LaserFront.Trimmer.DefWin32Fnc

Public Class clsLoaderIf

    ' TLF製ローダ用のI/O - I/F
    Public Const LARM_ARM1 As UShort = &H100                        ' B8 : 軽故障(続行可能)※アラーム状態解除後スタート(W109.01)で続行
    Public Const LARM_ARM2 As UShort = &H200                        ' B9 : サイクル停止    ※リセット(W109.03)後原点復帰してアイドル状態
    Public Const LARM_ARM3 As UShort = &H400                        ' B10: 全停止異常      ※同上

    Public EXTOUT_EX_LOK_ON As UShort = &H20                        ' B5 : EXTBIT：電磁ロック(扉ロック)


    '----- 出力データ(トリマ  → ローダ) -----
    Public Const LOUT_REDY As UShort = &H1                              ' B0 : トリマ部レディ(0=Not Ready, 1=Ready) ※BITを反転して出力する									
    Public Const LOUT_AUTO As UShort = &H2                              ' B1 : ローダモード切替え(0=手動, 1=自動(マガジンチェック))									
    Public Const LOUT_STOP As UShort = &H4                              ' B2 : トリマ部停止中(0=動作中,1=停止中)									
    Public Const LOUT_SUPLY As UShort = &H8                             ' B3 : 基板要求(0=基板要求無,1=基板要求)※連続運転開始									
    Public Const LOUT_TRM_NG As UShort = &H10                           ' B4 : トリミングＮＧ  (0:正常, 1:ＮＧ)									
    Public Const LOUT_INTLOK_DISABLE As UShort = &H20                   ' B5 : インターロック解除(0:インターロック中, 1:インターロック解除中(全部/一部)) 							
    Public Const LOUT_REQ_COLECT As UShort = &H40                       ' B6 : 基板回収要求(0=要求無,1=要求有)									
    Public Const LOUT_PROC_CONTINUE As UShort = &H80                    ' B7 : 動作継続信号(0:なし, 1:継続実行)									
    Public Const LOUT_ORG_BACK As UShort = &H100                        ' B8 : ローダ原点復帰要求(0=原点復帰未要求, 1=原点復帰要求)									
    Public Const LOUT_STB As UShort = &H200                             ' B9 : 品種データ送信(STB)									
    Public Const LOUT_NG_DISCHRAGE As UShort = &H400                    ' B10: ＮＧ基板排出要求(0=ＮＧ排出要求無, 1=ＮＧ排出要求)									
    Public Const LOUT_DISCHRAGE As UShort = &H800                       ' B11: 供給位置決完了(0=完了でない, 1=完了)									
    Public Const LOUT_STS_RUN As UShort = &H1000                        ' B12: トリマ運転中(0:一時停止, 1:運転中)									
    Public Const LOUT_EMPTY_OPE As UShort = &H2000                      ' B13: 空運転中									
    Public Const LOUT_UPPER_CLAMP As UShort = &H2000                    ' B13: 基板押えクランプ 									
    Public Const LOUT_CYCL_STOP As UShort = &H2000                      ' B13: サイクル停止要求(0=要求無,1=要求)
    Public Const LOUT_VACCUME As UShort = &H4000                        ' B14: 吸着(手動モード時有効)									
    Public Const LOUT_CLAMP As UShort = &H8000                          ' B15: 載物台クランプ開閉(0=開, 1=閉)(手動モード時有効)									

    '----- 入力データ(ローダ  → トリマ) -----				
    Public Const LINP_READY As UShort = &H1                             ' B0 : ローダ部レディ(0=Not Ready, 1=Ready)				
    Public Const LINP_AUTO As UShort = &H2                              ' B1 : ローダモード切替え(0=手動, 1=自動)				
    Public Const LINP_STOP As UShort = &H4                              ' B2 : ローダ部停止中(0=基板交換中, 1=停止中)				
    Public Const LINP_TRM_START As UShort = &H8                         ' B3 : トリミングスタート要求(0=スタート非要求, 1=スタート要求)				
    Public Const LINP_NGTRAY_OUT_COMP As UShort = &H10                  ' B4 : NGトレイへの排出完了信号				
    Public Const LINP_TRM_LOTCHANGE_START As UShort = &H10              ' B4 : ロット切替え信号
    Public Const LIN_CYCL_STOP As UShort = &H20                         ' B5 : サイクル停止応答(0=応答無,1=応答)
    Public Const LINP_cHSTcPROCSTART As UShort = &H20                   ' B5 : 加工スタート信号：(0=応答無,1=応答)
    Public Const LINP_cHSTcVACERROR As UShort = &H40                    ' B6 : 吸着エラー：0=正常、1=エラー発生 
    Public Const LINP_HST_MOVESUPPLY As Short = &H40S                    ' B6 : 供給位置移動指示 			
    Public Const LINP_MAGAZINE_SHIFT As UShort = &H40                   ' B6 : マガジンシフト完了：0:シフトしていない、1:シフトしている				
    Public Const LINP_NO_ALM_RESTART As UShort = &H80                   ' B7 : ローダ部正常(0=アラーム発生, 1=正常)				
    Public Const LINP_ORG_BACK As UShort = &H100                        ' B8 : ローダ原点復帰完了(0=原点復帰未完了, 1=原点復帰完了)				
    Public Const LINP_LOT_CHG As UShort = &H200                         ' B9 : ロット切換要求(0=ロット切換非要求, 1=ロット切換要求(満杯))				
    Public Const LINP_END_MAGAZINE As UShort = &H400                    ' B10: マガジン終了(0=マガジン非終了, 1=マガジン終了)				
    Public Const LINP_END_ALL_MAGAZINE As UShort = &H800                ' B11: 全マガジン終了(0=全マガジン非終了, 1=全マガジン終了)				
    Public Const LINP_NG_FULL As UShort = &H1000                        ' B12: NG排出満杯(0=NG排出未満杯, 1=NG排出満杯(完了))				
    Public Const LINP_DISCHRAGE As UShort = &H2000                      ' B13: 排出ピック完了(0=完了でない, 1=完了)				
    Public Const LINP_2PIECES As UShort = &H4000                        ' B14: ２枚取り検出(0=２枚取り未検出, 1=２枚取り検出)				
    Public Const LINP_WBREAK As UShort = &H8000                         ' B15: 基板割れ検出(0=基板割れ未検出, 1=基板割れ検出)				

    Public giOPLDTimeOutFlg As Integer                              ' ローダ通信タイムアウト検出(0=検出無し, 1=検出あり)
    Public giOPLDTimeOut As Integer                                 ' ローダ通信タイムアウト時間(msec)
    Public giOPLDTimeOutCounter As Integer                          ' ローダ通信タイムアウトリトライカウンター 
    Public giOPLDTimeOutExtCounter As Integer                       ' ローダ通信タイムアウト延長リトライカウンター 
    Public gbOPLDTimeOutExt As Boolean = False                      ' ローダ通信タイムアウト延長ビット確認フラグ 

    Public gLdWDate As UShort = 0                                       ' ローダ部送信データ(モニタ用)
    Public gLdRDate As UShort = 0                                       ' ローダ部受信データ(モニタ用)

    Private bFgTimeOut As Boolean                                       ' ローダ通信タイムアウトフラグ
    Private iBefData(7) As Integer                                      ' アラーム情報退避域

    '電磁ロック用
    Public EX_LOK_STS As Integer = &H216A                           ' 電磁ロックステータスアドレス
    Public EX_LOK_TOUT As Integer = (10 * 1000)                     ' 電磁ロックステータスタイムアウト値(msec)
    '----- EL_Lock_OnOff関数のモード -----
    Public EX_LOK_MD_ON As Integer = 1                              ' 電磁ロックON
    Public EX_LOK_MD_OFF As Integer = 0                             ' 電磁ロックOFF
    Public EXTINP_EX_LOK_ON As Integer                              ' 電磁ロック用

    Public Const LALARM_COUNT As Integer = 128                      ' 最大アラーム数
    Public bFgLoaderAlarmFRM As Boolean                             ' ローダアラーム画面表示中 

    Public Const MAXWORK_KND As Integer = 10                         ' 基板品種の数
    Private gfBordTableOutPosX(0 To MAXWORK_KND - 1) As Double       ' ローダ基板テーブル排出位置X
    Private gfBordTableOutPosY(0 To MAXWORK_KND - 1) As Double       ' ローダ基板テーブル排出位置Y
    Private gfBordTableInPosX(0 To MAXWORK_KND - 1) As Double        ' ローダ基板テーブル供給位置X
    Private gfBordTableInPosY(0 To MAXWORK_KND - 1) As Double        ' ローダ基板テーブル供給位置Y
    Private gdExechangeTheta As Double                               ' 基板交換時θ載物台角度
    Private bFgCyclStp As Boolean                                    ' サイクル停止する／しない 

    Public Const COVER_CHECK_OFF As Integer = 1                      ' 固定カバーチェックを行わない  
    Public Const COVER_CHECK_ON As Integer = 0                       ' 固定カバーチェックを行う      

    Public m_lTrimResult As Integer = cFRS_NORMAL                   ' 基板単位のトリミング結果 
    '                                                               ' cFRS_NORMAL (正常)
    '                                                               ' cFRS_TRIM_NG(トリミングNG)
    '                                                               ' cFRS_ERR_PTN(パターン認識エラー) ※なし
    Public gbIniFlg As Integer = 0                                   ' 初期フラグ(0=初回, 1= トリミング中, 2=終了) 
    Public Const LOADER_PARAMPATH As String = "C:\TRIM\LOADER.INI"  ' ローダパラメータファイル
    Private giLotAbort As Integer = 0                               ' ロット中断用フラグ 
    Private giLotChangeFlg As Integer = 0                           ' ロット切り替え実行フラフ   'V2.2.1.1⑧

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    Public Sub New()
        Dim strKEY As String
        Dim strSEC As String



        Try

            gdExechangeTheta = Double.Parse(GetPrivateProfileString_S("DEVICE_CONST", "CLAMP_LESS_THETA", SYSPARAMPATH, "0.0"))

            ' ローダとの搬送に使用する載物台位置の取得 
            strSEC = "SYSTEM"
            strKEY = "EJECTPOS1X"
            gfBordTableOutPosX(0) = Val(GetPrivateProfileString_S(strSEC, strKEY, LOADER_PARAMPATH, "0.0000"))          ' ローダ基板テーブル排出位置X
            strKEY = "EJECTPOS1Y"
            gfBordTableOutPosY(0) = Val(GetPrivateProfileString_S(strSEC, strKEY, LOADER_PARAMPATH, "0.0000"))          ' ローダ基板テーブル排出位置Y

            strKEY = "INSERTPOS1X"
            gfBordTableInPosX(0) = Val(GetPrivateProfileString_S(strSEC, strKEY, LOADER_PARAMPATH, "0.0000"))           ' ローダ基板テーブル供給位置X
            strKEY = "INSERTPOS1Y"
            gfBordTableInPosY(0) = Val(GetPrivateProfileString_S(strSEC, strKEY, LOADER_PARAMPATH, "0.0000"))           ' ローダ基板テーブル供給位置Y

            giOPLDTimeOutFlg = Val(GetPrivateProfileString_S("DEVICE_CONST", "LOADER_TIMEOUT_CHECK", SYSPARAMPATH, "1"))  ' ローダ通信タイムアウト検出(0=検出無し, 1=検出あり)
            giOPLDTimeOut = Val(GetPrivateProfileString_S("DEVICE_CONST", "LOADER_TIMEOUT", SYSPARAMPATH, "180000"))      ' ローダ通信タイムアウト時間(msec)


        Catch ex As Exception

        End Try

    End Sub

    '===============================================================================
    '   共通関数
    '===============================================================================

#Region "ローダからのトリミングスタート待ち処理(トリミング実行時用)"
    '''=========================================================================
    ''' <summary>ローダからのトリミングスタート待ち処理(トリミング実行時用)</summary>
    ''' <remarks>ローダへ基板要求を送信し、ローダからのトリミングスタート信号を待つ</remarks>
    ''' <param name="ObjSys">       (INP)OcxSystemオブジェクト</param>
    ''' <param name="bFgAutoMode">  (INP)ローダ自動モードフラグ</param>
    ''' <param name="iTrimResult">  (INP)トリミング結果(前回)
    '''                                   cFRS_NORMAL   = 正常
    '''                                   cFRS_TRIM_NG  = トリミングNG
    '''                                   cFRS_ERR_PTN  = パターン認識エラー</param>
    ''' <param name="bFgMagagin">   (OUT)マガジン終了フラグ</param>
    ''' <param name="bFgAllMagagin">(OUT)全マガジン終了フラグ</param>
    ''' <param name="bFgLot">       (OUT)ロット切替え要求フラグ</param>
    ''' <param name="bIniFlg">      (INP)初期フラグ(0=初回, 1=トリミング中,
    '''                                             2=全マガジン終了(未使用), 3=最終基板の取出)
    '''                                                                                        </param>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          cFRS_ERR_LDR1 = ローダアラーム検出(全停止異常)
    '''          cFRS_ERR_LDR2 = ローダアラーム検出(サイクル停止)
    '''          cFRS_ERR_LDR3 = ローダアラーム検出(軽故障(続行可能))
    '''          cFRS_ERR_RST  = RESETキー押下
    '''          cFRS_ERR_CVR  = 筐体カバー開検出
    '''          cFRS_ERR_SCVR = スライドカバー開検出
    '''          cFRS_ERR_EMG  = 非常停止 他</returns>
    '''=========================================================================
    Public Function Loader_WaitTrimStart(ByVal ObjSys As SystemNET, ByVal bFgAutoMode As Boolean, ByVal iTrimResult As Integer,
                                         ByRef bFgMagagin As Boolean, ByRef bFgAllMagagin As Boolean, ByRef bFgLot As Boolean, ByVal bIniFlg As Integer) As Integer

        Dim Idx As Integer
        Dim r As Integer
        Dim rtnCode As Integer = cFRS_NORMAL
        Dim OnBit As UShort
        Dim OffBit As UShort
        Dim WaitBit As UShort
        Dim strMSG As String
        Dim RetKey As Integer

        Try
            AutoOperationDebugLogOut("Loader_WaitTrimStart() - Start")

            ' ローダ無効またはローダ手動モードならNOP
            If ((bFgAutoMode = False)) Then
                Return (cFRS_NORMAL)
            End If


            '-------------------------------------------------------------------
            '   初回以外なら基板排出処理を行う
            '-------------------------------------------------------------------
            If (bIniFlg <> 0) Then                                      ' 初回以外 ?
                rtnCode = Loader_WaitDischarge(ObjSys, bFgAutoMode, iTrimResult, bFgMagagin, bFgAllMagagin, bFgLot)
                If (rtnCode = cFRS_ERR_START) Then                      ' サイクル停止で基板なし続行指定 ?
                    GoTo STP_010                                        '
                End If
                If (rtnCode <> cFRS_NORMAL) Then                        ' エラー ? (※エラー発生時のメッセージは表示済)
                    Return (rtnCode)
                End If
                If (bIniFlg = 2) Then                                   ' 全マガジン終了なら終了
                    Return (rtnCode)
                End If
            End If
STP_010:
            ' 前回のトリミングスタート要求Bit/排出ピック完了BitのOffを待つ
            If (bIniFlg = 0) Then                                       ' 初回 ?
                WaitBit = LINP_TRM_START
                OnBit = LOUT_SUPLY + LOUT_STOP                          ' ローダ出力BIT = 基板要求+トリマ部停止中
                OffBit = LOUT_DISCHRAGE + LOUT_PROC_CONTINUE            ' OffBit = 供給位置決完了
            Else
                WaitBit = LINP_TRM_START
                OnBit = LOUT_DISCHRAGE + LOUT_STOP                      ' ローダ出力BIT = 供給位置決完了+トリマ部停止中
                OffBit = LOUT_SUPLY + LOUT_PROC_CONTINUE                ' OffBit = 基板要求
            End If
STP_RETRY:

            rtnCode = Sub_WaitLoaderData(ObjSys, WaitBit, False, bFgMagagin, bFgAllMagagin, bFgLot, RetKey)
            If (rtnCode <> cFRS_NORMAL) Then                            ' エラー ? (※エラー発生時のメッセージは表示済)
                If ((rtnCode = cFRS_ERR_LDR3) Or (rtnCode = cFRS_ERR_LDR2)) AndAlso (RetKey = cFRS_ERR_START) Then ' ###196
                    Call W_RESET()                                      ' アラームリセット信号送出
                    Call W_START()                                      ' スタート信号送出
                    GoTo STP_RETRY
                End If
                Return (rtnCode)
            End If

            '-------------------------------------------------------------------
            '   テーブルを基板供給位置に移動する
            '-------------------------------------------------------------------
            If (bIniFlg <> 3) Then                                      ' 最終基板の取出ならNOP 
                Idx = 0               ' Idx = 基板品種番号 - 1
                '基板品種は1を使用
                r = SMOVE2(gfBordTableInPosX(Idx), gfBordTableInPosY(Idx))
                If (r <> cFRS_NORMAL) Then                                  ' エラー ?
                    rtnCode = Form1.System1.EX_ZGETSRVSIGNAL(gSysPrm, r, 0) ' メッセージ表示
                    Return (rtnCode)
                End If
            End If


            '-------------------------------------------------------------------
            '   初回    →基板要求信号を送信し、トリミングスタート要求を待つ
            '   初回以外→供給位置決完了信号を送信し、トリミングスタート要求を待つ
            '-------------------------------------------------------------------
            ' ローダへ基板要求信号(初回)または供給位置決完了(初回以外)を送信する (トリミングNG, パターン認識エラーは基板要求信号と同時に出力する)
            OffBit = OffBit + LOUT_NG_DISCHRAGE + LOUT_TRM_NG           ' 「ＮＧ基板排出要求」/ [トリミングNG](BITをOFFする) 
            If (iTrimResult <> cFRS_NORMAL) Then                        ' 基板単位のトリミング結果(前回)が正常でなければ
                OnBit = OnBit + LOUT_NG_DISCHRAGE                       ' 「ＮＧ基板排出要求」BITをONする
            End If
            If (iTrimResult = cFRS_TRIM_NG) Then                        ' 基板単位のトリミング結果(前回) = トリミングNG ?
                OnBit = OnBit + LOUT_TRM_NG
            ElseIf (iTrimResult = cFRS_ERR_PTN) Then                    ' 基板単位のトリミング結果(前回) = パターン認識エラー ?
                OnBit = OnBit + LOUT_TRM_NG
            End If
            OffBit = OffBit And Not LOUT_SUPLY                          ' 基板要求(連続運転開始)はOFFしない
            Call Sub_ATLDSET(OnBit, OffBit)                             ' ローダ出力(ON=基板要求または供給位置決完了+ﾄﾘﾏ停止中+他, OFF=供給位置決完了または基板要求)

            ' ローダからのトリミングスタート要求を待つ
STP_RETRY2:
            rtnCode = Sub_WaitLoaderData(ObjSys, LINP_TRM_START, True, bFgMagagin, bFgAllMagagin, bFgLot, RetKey)
            'If (rtnCode = cFRS_ERR_LDR3) Then                           ' 軽故障(続行可能) ? ###196 ###073
            If ((rtnCode = cFRS_ERR_LDR3) Or (rtnCode = cFRS_ERR_LDR2)) AndAlso (RetKey = cFRS_ERR_START) Then ' ###196
                Call W_RESET()                                          ' アラームリセット信号送出
                Call W_START()                                          ' スタート信号送出
                GoTo STP_RETRY2
            End If
            If (rtnCode = cFRS_NORMAL) Then                             ' 正常 ? (※エラー発生時のメッセージは表示済)
                ' ローダへトリマ動作中(ﾄﾘﾏ停止OFF)を送信する
                OnBit = OnBit And Not LOUT_SUPLY                        ' 基板要求(連続運転開始)はOFFしない
                Call Sub_ATLDSET(&H0, OnBit)                            ' ローダ出力(ON=なし, OFF=基板要求+ﾄﾘﾏ停止中+他)
            End If

            Return (rtnCode)

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "LoaderIOFor436.Loader_WaitTrimStart() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)
        End Try
    End Function
#End Region


#Region "ローダからの応答データ待ち"
    '''=========================================================================
    ''' <summary>ローダからの応答データ待ち</summary>
    ''' <param name="ObjSys">       (INP)OcxSystemオブジェクト</param>
    ''' <param name="WaitData">     (INP)応答待ちするデータ</param>
    ''' <param name="OnOff">        (INP)True=On待ち, False=Off待ち</param>
    ''' <param name="bFgMagagin">   (OUT)マガジン終了フラグ</param>
    ''' <param name="bFgAllMagagin">(OUT)全マガジン終了フラグ</param>
    ''' <param name="bFgLot">       (OUT)ロット切替え要求フラグ</param>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          cFRS_ERR_LDR1 = ローダアラーム検出(全停止異常)
    '''          cFRS_ERR_LDR2 = ローダアラーム検出(サイクル停止)
    '''          cFRS_ERR_LDR3 = ローダアラーム検出(軽故障(続行可能))
    '''          cFRS_ERR_RST  = RESETキー押下
    '''          cFRS_ERR_CVR  = 筐体カバー開検出
    '''          cFRS_ERR_SCVR = スライドカバー開検出
    '''          cFRS_ERR_EMG  = 非常停止 他
    ''' </returns>
    '''=========================================================================
    Private Function Sub_WaitLoaderData(ByVal ObjSys As SystemNET, ByVal WaitData As UShort, ByVal OnOff As Boolean,
                                        ByRef bFgMagagin As Boolean, ByRef bFgAllMagagin As Boolean, ByRef bFgLot As Boolean, ByRef RetStat As Integer) As Integer

        Dim TimerRS As System.Threading.Timer = Nothing
        Dim LdIn As UShort
        Dim WaitBit As UShort
        Dim rtnCode As Integer = cFRS_NORMAL
        Dim r As Integer
        Dim strMSG As String
        Dim BreakFirst As Integer
        Dim TwoTakeFirst As Integer
        Dim bFlgWbrk As Boolean = False                                 ' タイマーリセットフラグ 
        Dim bFlg2Pce As Boolean = False                                 ' タイマーリセットフラグ 
        Dim iCnt As Integer
        Dim strLoaderAlarm(LALARM_COUNT) As String
        Dim strLoaderAlarmInfo(LALARM_COUNT) As String
        Dim strLoaderAlarmExec(LALARM_COUNT) As String
        Dim AlarmKind As Integer



        Try
            ' 応答待ちするデータはOn/Off待ち ?
            If (OnOff = True) Then
                WaitBit = WaitData                                      ' On待ち
            Else
                WaitBit = 0                                             ' Off待ち
            End If
            BreakFirst = 0
            TwoTakeFirst = 0

            AutoOperationDebugLogOut("Sub_WaitLoaderData() Start")       ''V2.2.1.3②

            ' ローダ通信タイムアウトチェック用タイマーオブジェクトの作成(TimerRS_TickをX msec間隔で実行する)
            Sub_SetTimeoutTimer(TimerRS)

            ' ローダからの応答データを待つ
            Do
                ' ローダアラーム/非常停止チェック
                r = GetLoaderIO(LdIn)                                   ' ローダ
                If ((LdIn And LINP_NO_ALM_RESTART) <> LINP_NO_ALM_RESTART) Then
                    rtnCode = cFRS_ERR_LDR                              ' Return値 = ローダアラーム検出
                    GoTo STP_ERR_LDR                                    ' ローダアラーム表示へ
                End If

                ' ローダ通信タイムアウトチェック
                If (bFgTimeOut = True) Then                             ' タイムアウト ?
                    ' コールバックメソッドの呼出しを停止する
                    TimerRS.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)
                    rtnCode = cFRS_ERR_LDRTO                            ' Return値 = ローダ通信タイムアウト
                    GoTo STP_ERR_LDR                                    ' エラーメッセージ表示へ
                End If

                ' 非常停止等チェック(トリマ装置アイドル中)
                r = Form1.System1.Sys_Err_Chk_EX(gSysPrm, giAppMode)           ' 非常停止/ｶﾊﾞｰ/ｴｱｰ圧/集塵機/ﾏｽﾀｰﾊﾞﾙﾌﾞﾁｪｯｸ
                If (r <> cFRS_NORMAL) Then                                  ' 非常停止等検出 ?
                    rtnCode = cFRS_ERR_EMG                            ' Return値 = ローダ通信タイムアウト
                    GoTo STP_ERR_LDR                                    ' エラーメッセージ表示へ
                End If


                ' 筐体カバー閉/スライドカバー閉/非常停止チェック
                r = ObjSys.Sys_Err_Chk_EX(gSysPrm, giAppMode)           ' 非常停止/ｶﾊﾞｰ/ｴｱｰ圧/集塵機/ﾏｽﾀｰﾊﾞﾙﾌﾞﾁｪｯｸ
                If (r <> cFRS_NORMAL) Then                                  ' 非常停止等検出 ?
                    GoTo STP_END
                End If

                ' トリミング要求ON待ちのときにオールマガジン終了で交換中信号がOFFした場合には、最後の基板が異常終了したと認識する。
                If (WaitData = LINP_TRM_START) And (OnOff = True) Then      ' トリミングスタート要求待ち ?
                    If (((LdIn And LINP_STOP) = LINP_STOP) And ((LdIn And LINP_END_ALL_MAGAZINE) = LINP_END_ALL_MAGAZINE)) Then
                        rtnCode = cFRS_ERR_LOTEND                       ' Return値設定(cFRS_ERR_RST/cFRS_ERR_EMG)
                        bFgAllMagagin = True
                        For iCnt = 1 To 3
                            ' ローダアラーム/非常停止チェック
                            r = GetLoaderIO(LdIn)                           ' ローダ入力
                            If ((LdIn And LINP_NO_ALM_RESTART) <> LINP_NO_ALM_RESTART) Then
                                Console.WriteLine("Sub_WaitLoaderData() ローダアラーム(LdIn=%f)", LdIn)
                                rtnCode = cFRS_ERR_LDR                      ' Return値 = ローダアラーム検出
                                GoTo STP_ERR_LDR                            ' ローダアラーム表示へ
                            End If
                            Call System.Threading.Thread.Sleep(100)
                        Next
                        GoTo STP_END
                    End If

                    '-----------------------------------------------------------
                    '「トリミングスタート要求待ち」のときに「基板割れ検出」又は
                    '「２枚取り検出」なら ローダ通信タイムアウトタイマーを生成し直す
                    '-----------------------------------------------------------
                    If ((bFlgWbrk = True) Or (bFlg2Pce = True)) Then    ' タイマーリセットFLG ON ?
                        ' コールバックメソッドの呼出しを停止する
                        TimerRS.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)
                        TimerRS.Dispose()                               ' タイマーを破棄する
                        Sub_SetTimeoutTimer(TimerRS)                    ' タイマー生成
                    End If

                End If

                ' 
                ' スタート信号OFF待ちの時に前マガジン終了がONしていたらロット終了になる 
                If giLoaderType = 1 Then
                    If (WaitData = LINP_TRM_START) And (OnOff = False) Then      ' トリミングスタート要求のOFF待ち

                        If (((LdIn And LINP_STOP) = LINP_STOP) And ((LdIn And LINP_END_ALL_MAGAZINE) = LINP_END_ALL_MAGAZINE)) Then
                            rtnCode = cFRS_ERR_LOTEND                       ' Return値設定(cFRS_ERR_RST/cFRS_ERR_EMG)
                            bFgAllMagagin = True
                            Exit Do
                        End If

                    End If
                End If

                System.Windows.Forms.Application.DoEvents()
                Call System.Threading.Thread.Sleep(1)                   ' Wait(msec)
            Loop While ((LdIn And WaitData) <> WaitBit)                 ' 応答データ待ち

            ' マガジン終了, 全マガジン終了, ロット切替えチェック
            If (WaitData = LINP_TRM_START) And (OnOff = True) Then      ' トリミングスタート要求待ち ?
                rtnCode = LoaderBitCheck(bFgMagagin, bFgAllMagagin, bFgLot)
            End If

            ' 終了処理
STP_END:
            ' コールバックメソッドの呼出しを停止する
            TimerRS.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)
            TimerRS.Dispose()                                           ' タイマーを破棄する
            AutoOperationDebugLogOut("Sub_WaitLoaderData() - STP_END - TimerRS.Dispose()")       ''V2.2.1.3②
            Return (rtnCode)

            ' ローダエラー発生時
STP_ERR_LDR:
            ' コールバックメソッドの呼出しを停止する
            TimerRS.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)
            TimerRS.Dispose()                                           ' タイマーを破棄する
            AutoOperationDebugLogOut("Sub_WaitLoaderData() - STP_ERR_LDR - TimerRS.Dispose()")       ''V2.2.1.3②

            If (rtnCode = cFRS_ERR_LDRTO) Then                          ' ローダ通信タイムアウト ?

                AutoOperationDebugLogOut("Sub_WaitLoaderData() - rtnCode = cFRS_ERR_LDRTO")       ''V2.2.1.3②

                ' rtnCode = Sub_CallFrmRset(ObjSys, cGMODE_LDR_TMOUT)     ' エラーメッセージ表示
                rtnCode = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_TMOUT)
                RetStat = cFRS_ERR_RST       ' ボタンを押した内容：
                rtnCode = cFRS_ERR_LDR1     ' アラームレベル: 
            Else
                ' ローダアラームメッセージ作成 & ローダアラーム画面表示
                ' rtnCode = Loader_AlarmCheck(ObjSys, True, AlarmCount, strLoaderAlarm, strLoaderAlarmInfo, strLoaderAlarmExec)
                AlarmKind = cGMODE_LDR_ALARM
                rtnCode = ObjSys.Sub_CallFormLoaderAlarm(AlarmKind, ObjPlcIf)
                RetStat = rtnCode       ' ボタンを押した内容：
                rtnCode = AlarmKind     ' アラームレベル: 
            End If

            If AlarmKind <> cFRS_ERR_LDR3 Then
                Call Sub_ATLDSET(&H0, LOUT_AUTO)        ' ローダ手動モード切替え(ローダ出力(ON=なし, OFF=自動))
            End If

            Return (rtnCode)

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "LoaderIOFor436.Sub_WaitLoaderData() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            AutoOperationDebugLogOut("Sub_WaitLoaderData() - cERR_TRAP")       ''V2.2.1.3②

            Return (cERR_TRAP)
        End Try
    End Function
#End Region


#Region "ローダからの基板排出完了待ち処理(トリミング実行時用)"
    '''=========================================================================
    ''' <summary>ローダからの基板排出完了待ち処理(トリミング実行時用)</summary>
    ''' <remarks>ローダへ基板排出要求を送信し、ローダからの基板排出完了信号を待つ</remarks>
    ''' <param name="ObjSys">       (INP)OcxSystemオブジェクト</param>
    ''' <param name="bFgAutoMode">  (INP)ローダ自動モードフラグ</param>
    ''' <param name="iTrimResult">  (INP)トリミング結果(前回)
    '''                                   cFRS_NORMAL   = 正常
    '''                                   cFRS_TRIM_NG  = トリミングNG
    '''                                   cFRS_ERR_PTN  = パターン認識エラー</param>
    ''' <param name="bFgMagagin">   (OUT)マガジン終了フラグ</param>
    ''' <param name="bFgAllMagagin">(OUT)全マガジン終了フラグ</param>
    ''' <param name="bFgLot">       (OUT)ロット切替え要求フラグ</param>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          cFRS_ERR_START= 正常(サイクル停止で基板なし続行指定) 
    '''          cFRS_ERR_LDR1 = ローダアラーム検出(全停止異常)
    '''          cFRS_ERR_LDR2 = ローダアラーム検出(サイクル停止)
    '''          cFRS_ERR_LDR3 = ローダアラーム検出(軽故障(続行可能))
    '''          cFRS_ERR_RST  = RESETキー押下
    '''          cFRS_ERR_CVR  = 筐体カバー開検出
    '''          cFRS_ERR_SCVR = スライドカバー開検出
    '''          cFRS_ERR_EMG  = 非常停止 他</returns>
    '''=========================================================================
    Public Function Loader_WaitDischarge(ByVal ObjSys As SystemNET, ByVal bFgAutoMode As Boolean, ByVal iTrimResult As Integer,
                                         ByRef bFgMagagin As Boolean, ByRef bFgAllMagagin As Boolean, ByRef bFgLot As Boolean) As Integer

        Dim Idx As Integer
        Dim r As Integer
        Dim rtnCode As Integer = cFRS_NORMAL
        Dim OnBit As UShort
        Dim OffBit As UShort
        Dim strMSG As String
        Dim RetKey As Integer

        Try
            ' ローダ無効またはローダ手動モードならNOP
            If (bFgAutoMode = False) Then
                Return (cFRS_NORMAL)
            End If

            ' 前回の排出ピック完了BitのOffを待つ
STP_RETRY:
            rtnCode = Sub_WaitLoaderData(ObjSys, LINP_DISCHRAGE, False, bFgMagagin, bFgAllMagagin, bFgLot, RetKey)
            If (rtnCode <> cFRS_NORMAL) Then                            ' エラー ? (※エラー発生時のメッセージは表示済)
                If rtnCode = cFRS_ERR_EMG Then
                    Return (rtnCode)
                End If
                If ((rtnCode = cFRS_ERR_LDR3) Or (rtnCode = cFRS_ERR_LDR2)) AndAlso (RetKey = cFRS_ERR_START) Then
                    Call W_RESET()                                      ' アラームリセット信号送出
                    Call W_START()                                      ' スタート信号送出
                    GoTo STP_RETRY
                End If
                Return (rtnCode)
            End If

            '-------------------------------------------------------------------
            '   テーブルを基板排出位置に移動する
            '-------------------------------------------------------------------
            Idx = 1                   ' Idx = 基板品種番号 - 1

            If (0 <> gSysPrm.stDEV.giTheta) Then
                'SL436Rでθ有の場合には、指定角度に回転する
                ROUND4(gdExechangeTheta)
            End If

            r = MoveGlassOutPos()
            'r = SMOVE2(gfBordTableOutPosX(Idx), gfBordTableOutPosY(Idx))
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?
                rtnCode = Form1.System1.EX_ZGETSRVSIGNAL(gSysPrm, r, 0) ' メッセージ表示
                Return (rtnCode)
            End If

            ' ロットの中断を確認
            If giLotAbort <> 0 Then
                Return cFRS_ERR_RST
            End If

            giLotAbort = 0
            '-------------------------------------------------------------------
            '   サイクル停止処理
            '-------------------------------------------------------------------
            If (Form1.JudgeCycleStop() = 1) Then                                 ' サイクル停止フラグ ON ?
                r = CycleStop_Proc(ObjSys)                              ' サイクル停止処理
                Call LAMP_CTRL(LAMP_HALT, False)                        ' サイクル停止処理が終わったらHALTランプは消灯する
                bFgCyclStp = False              ' サイクル停止フラグOFF
                If (r < cFRS_NORMAL) Then Return (r) '                  ' 非常停止等のエラーならエラー戻り(エラーメッセージは表示済み)
                If ((r = cFRS_ERR_START) Or (r = cFRS_ERR_RST)) Then    ' 基板なし続行またはCancel(RESETキー押下) ?
                    Return (r)                                          ' Return値 = cFRS_ERR_START(基板なし続行), cFRS_ERR_RST(Cancel(RESETキー押下))
                End If
            End If                                                      ' r = cFRS_NORMAL(基板あり続行)なら処理続行

            '-------------------------------------------------------------------
            '   基板要求信号(初回以外)を送信し、排出ピック完了を待つ
            '------------------------------------------------------------------
            ' 基板要求信号を送信する前にローダ部停止中BitのONを待つ
STP_RETRY2:
            rtnCode = Sub_WaitLoaderData(ObjSys, LINP_STOP, True, bFgMagagin, bFgAllMagagin, bFgLot, RetKey)
            If (rtnCode <> cFRS_NORMAL) Then                            ' エラー ? (※エラー発生時のメッセージは表示済)
                'If (rtnCode = cFRS_ERR_LDR3) Then                       ' 軽故障(続行可能) ? ###196 ###073
                If ((rtnCode = cFRS_ERR_LDR3) Or (rtnCode = cFRS_ERR_LDR2)) AndAlso (RetKey = cFRS_ERR_START) Then ' ###196
                    Call W_RESET()                                      ' アラームリセット信号送出
                    Call W_START()                                      ' スタート信号送出
                    GoTo STP_RETRY2
                End If
                Return (rtnCode)
            End If

            ' ローダへ基板要求信号(初回以外)を送信する (トリミングNG, パターン認識エラーは基板要求信号と同時に出力する)

            If (bFgAllMagagin = True) Then
                OnBit = LOUT_REQ_COLECT + LOUT_STOP                     ' ローダ出力BIT = 基板回収要求+トリマ部停止中
                OffBit = LOUT_SUPLY
            Else
                OnBit = LOUT_SUPLY + LOUT_REQ_COLECT + LOUT_STOP        ' ローダ出力BIT = 基板要求+基板回収要求+トリマ部停止中
                OffBit = 0
            End If
            If (iTrimResult <> cFRS_NORMAL) Then                        ' 基板単位のトリミング結果(前回)が正常でなければ
                OnBit = OnBit + LOUT_NG_DISCHRAGE                       ' 「ＮＧ基板排出要求」BITをONする
            Else

            End If
            If (iTrimResult = cFRS_TRIM_NG) Then                        ' 基板単位のトリミング結果(前回) = トリミングNG ?
                OnBit = OnBit + LOUT_TRM_NG
            ElseIf (iTrimResult = cFRS_ERR_PTN) Then                    ' 基板単位のトリミング結果(前回) = パターン認識エラー ?
                'OnBit = OnBit + LOUT_PTN_NG                            ' ###070
                OnBit = OnBit + LOUT_TRM_NG                             ' ###070
            End If
            Call Sub_ATLDSET(OnBit, OffBit)                             ' ローダ出力(ON=基板要求+ﾄﾘﾏ停止中+他, OFF=なし)

            ' ローダからの排出ピック完了を待つ
STP_RETRY3:
            rtnCode = Sub_WaitLoaderData(ObjSys, LINP_DISCHRAGE, True, bFgMagagin, bFgAllMagagin, bFgLot, RetKey)
            'If (rtnCode = cFRS_ERR_LDR3) Then                           ' 軽故障(続行可能) ? ###196 ###073
            If ((rtnCode = cFRS_ERR_LDR3) Or (rtnCode = cFRS_ERR_LDR2)) AndAlso (RetKey = cFRS_ERR_START) Then ' ###196
                Call W_RESET()                                          ' アラームリセット信号送出
                Call W_START()                                          ' スタート信号送出
                GoTo STP_RETRY3
            End If
            If (rtnCode = cFRS_NORMAL) Then                             ' 正常 ? (※エラー発生時のメッセージは表示済)
                ' ローダへトリマ動作中(ﾄﾘﾏ停止OFF)を送信する
                OnBit = OnBit And Not LOUT_SUPLY                        ' 基板要求(連続運転開始)はOFFしない
                Call Sub_ATLDSET(&H0, OnBit)                            ' ローダ出力(ON=なし, OFF=基板要求+ﾄﾘﾏ停止中+他)
            End If
            Return (rtnCode)

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "LoaderIOFor436.Loader_WaitDischarge() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)
        End Try
    End Function
#End Region


#Region "ローダ通信タイムアウトチェック用タイマー作成"
    '''=========================================================================
    ''' <summary>ローダ通信タイムアウトチェック用タイマー作成</summary>
    ''' <param name="TimerRS">(OUT)タイマー</param>
    '''=========================================================================
    Private Sub Sub_SetTimeoutTimer(ByRef TimerRS As System.Threading.Timer)

        Dim TimeVal As Integer
        Dim strMSG As String

        Try
            ' タイマー値を設定する
            If (giOPLDTimeOutFlg = 0) Then                              ' ローダ通信タイムアウト検出無し ?
                TimeVal = System.Threading.Timeout.Infinite             ' タイマー値 = なし
            Else
                TimeVal = giOPLDTimeOut                                 ' タイマー値 = ローダ通信タイムアウト時間(msec)
                giOPLDTimeOutExtCounter = 0                                ' リトライカウンター初期化
            End If

            ' ローダ通信タイムアウトチェック用タイマーオブジェクトの作成(TimerRS_TickをTimeVal msec間隔で実行する)
            bFgTimeOut = False                                          ' ローダ通信タイムアウトフラグOFF
            TimerRS = New System.Threading.Timer(New System.Threading.TimerCallback(AddressOf TimerRS_Tick), Nothing, TimeVal, TimeVal)

            AutoOperationDebugLogOut("Sub_SetTimeoutTimer() - TimerRS Start")       ''V2.2.1.3②

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "LoaderIOFor436.Sub_SetTimeoutTimer() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region



#Region "オートローダからデータを入力する"
    '''=========================================================================
    ''' <summary>オートローダからデータを入力する</summary>
    ''' <param name="LdIn">(OUT)入力データ</param>
    ''' <returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Public Function GetLoaderIO(ByRef LdIn As UShort) As Integer

        Dim r As Integer
        Dim iData As Integer
        Dim strMSG As String

        Try
            ' オートローダ入力
            r = ZATLDGET(iData)
            LdIn = iData                                                ' ローダ部受信データ設定
            gLdRDate = iData                                            ' ローダ部受信データ設定(モニタ用)

            Call IoMonitor(gLdRDate, 0)                                   ' IOﾓﾆﾀ表示

            Return (r)

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "LoaderIOFor436.GetLoaderIO() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)
        End Try
    End Function
#End Region


#Region "タイマーイベント(指定タイマ間隔が経過した時に発生)"
    '''=========================================================================
    ''' <summary>タイマーイベント(指定タイマ間隔が経過した時に発生)</summary>
    ''' <param name="Sts">(INP)</param>
    '''=========================================================================
    Private Sub TimerRS_Tick(ByVal Sts As Object)

        Dim strMSG As String

        Try

            bFgTimeOut = True                                           ' ローダ通信タイムアウトフラグON

            Exit Sub

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "LoaderIOFor436.TimerRS_Tick() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region


#Region "ローダからのマガジン終了,全マガジン終了,ロット切替え要求チェック"
    '''=========================================================================
    ''' <summary>ローダからのマガジン終了,全マガジン終了,ロット切替え要求チェック</summary>
    ''' <param name="bFgMagagin">   (OUT)マガジン終了フラグ</param>
    ''' <param name="bFgAllMagagin">(OUT)全マガジン終了フラグ</param>
    ''' <param name="bFgLot">       (OUT)ロット切替え要求フラグ</param>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          cFRS_ERR_LDR1 = ローダアラーム検出(全停止異常)
    '''          cFRS_ERR_LDR2 = ローダアラーム検出(サイクル停止)
    '''          cFRS_ERR_LDR3 = ローダアラーム検出(軽故障(続行可能))
    '''          cFRS_ERR_RST  = RESETキー押下
    '''          cFRS_ERR_CVR  = 筐体カバー開検出
    '''          cFRS_ERR_SCVR = スライドカバー開検出
    '''          cFRS_ERR_EMG  = 非常停止 他</returns>
    '''=========================================================================
    Private Function LoaderBitCheck(ByRef bFgMagagin As Boolean, ByRef bFgAllMagagin As Boolean, ByRef bFgLot As Boolean) As Integer

        Dim r As Integer
        Dim LdIn As UShort
        Dim strMSG As String
        Dim RetKey As Integer

        Try
            ' 全マガジン終了チェック
            r = GetLoaderIO(LdIn)                                       ' ローダ入力
            If (LdIn And LINP_END_ALL_MAGAZINE) Then                    ' 全マガジン終了？
                bFgAllMagagin = True
            Else
                bFgAllMagagin = False
            End If

            ' マガジン終了チェック
            If (bFgMagagin = True) Then                                 ' マガジン終了フラグON ?
                ' 前回のマガジン終了BitのOffを待つ
STP_RETRY:
                r = Sub_WaitLoaderData(Form1.System1, LINP_END_MAGAZINE, False, bFgMagagin, bFgAllMagagin, bFgLot, RetKey)
                If (r <> cFRS_NORMAL) Then                              ' エラー ? (※エラー発生時のメッセージは表示済)
                    'If (r = cFRS_ERR_LDR3) Then                         ' 軽故障(続行可能) ? 
                    If (r = cFRS_ERR_LDR3) Or (r = cFRS_ERR_LDR2) Then  ' 
                        Call W_RESET()                                  ' アラームリセット信号送出
                        Call W_START()                                  ' スタート信号送出
                        GoTo STP_RETRY
                    End If
                    Return (r)
                End If
            End If
            r = GetLoaderIO(LdIn)                                       ' ローダ入力
            If (LdIn And LINP_END_MAGAZINE) Then                        ' マガジン終了？
                bFgMagagin = True
            Else
                bFgMagagin = False
            End If

            ' ロット切替え要求チェック
            If (bFgLot = True) Then                                     ' ロット切替え要求フラグON ?
                ' 前回のロット切替えBitのOffを待つ
STP_RETRY2:
                r = Sub_WaitLoaderData(Form1.System1, LINP_LOT_CHG, False, bFgMagagin, bFgAllMagagin, bFgLot, RetKey)
                If (r <> cFRS_NORMAL) Then                              ' エラー ? (※エラー発生時のメッセージは表示済)
                    'If (r = cFRS_ERR_LDR3) Then                         ' 軽故障(続行可能) 
                    If (r = cFRS_ERR_LDR3) Or (r = cFRS_ERR_LDR2) Then  ' 
                        Call W_RESET()                                  ' アラームリセット信号送出
                        Call W_START()                                  ' スタート信号送出
                        GoTo STP_RETRY2
                    End If
                    Return (r)
                End If
            End If
            r = GetLoaderIO(LdIn)                                       ' ローダ入力
            If (LdIn And LINP_LOT_CHG) Then                             ' ロット切替え要求 ？
                bFgLot = True
            Else
                bFgLot = False
            End If

            Return (cFRS_NORMAL)

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "LoaderIOFor436.LoaderBitCheck() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)
        End Try
    End Function
#End Region


#Region "スタート信号送出(オートローダ通信)"
    '''=========================================================================
    ''' <summary>スタート信号送出(オートローダ通信) </summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub W_START()

        Dim strMSG As String

        Try
            ' スタート信号送出
            'Call m_PlcIf.WritePlcWR(LOFS_W109, 0)
            'Call m_PlcIf.WritePlcWR(LOFS_W109, LDDV_ARM_START)
            ObjSys.W_START()

            ' トラップエラー発生時(エラーメッセージはex.Messageに設定される)
        Catch ex As Exception
            strMSG = ex.Message + "(W_START)"
            'MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "アラームリセット信号送出(オートローダ通信)"
    '''=========================================================================
    ''' <summary>アラームリセット信号送出(オートローダ通信) </summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub W_RESET()

        Dim strMSG As String

        Try

            'Call m_PlcIf.WritePlcWR(LOFS_W109, 0)
            'Call m_PlcIf.WritePlcWR(LOFS_W109, LDDV_ARM_RESET)
            ObjSys.W_RESET()

            ' トラップエラー発生時(エラーメッセージはex.Messageに設定される)
        Catch ex As Exception
            strMSG = ex.Message + "(W_RESET)"
            'MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "FrmReset実行サブルーチン"
    '''=========================================================================
    ''' <summary>FrmReset実行サブルーチン</summary>
    ''' <param name="ObjSys">(INP)OcxSystemオブジェクト</param>
    ''' <param name="gMode"> (INP)処理モード</param>
    ''' <returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Public Function Sub_CallFrmRset(ByVal ObjSys As SystemNET, ByVal gMode As Integer) As Integer

        Dim r As Integer
        'Dim objForm As frmReset = Nothing
        Dim strMSG As String

        ' DllSystemにアラームメッセージ表示を作成して呼び出す

        Try

            '' FrmReset画面表示(処理モードに対応する処理を行う)
            'objForm = New frmReset()
            'Call objForm.ShowDialog(Nothing, gMode, ObjSys)
            'r = objForm.sGetReturn()                                    ' Return値取得

            '' オブジェクト開放
            'If (objForm Is Nothing = False) Then
            '    Call objForm.Close()                                    ' オブジェクト開放
            '    Call objForm.Dispose()                                  ' リソース開放
            'End If

            '' 原点復帰時はクランプを開状態にする 
            'Select Case (gMode)
            '    Case cGMODE_ORG, cGMODE_LDR_ORG
            '        If (r = cFRS_NORMAL) Then                               ' ###163
            '            Call W_CLMP_ONOFF(0)                                ' クランプＯＦＦ(開)

            '        End If
            '    Case cGMODE_LDR_END                                         'V4.12.2.4② 自動運転終了時追加
            '        If (r = cFRS_ERR_START) Then
            '            Call W_CLMP_ONOFF(0)                                ' クランプＯＦＦ(開)
            '        End If

            'End Select

            Return (r)                                                  ' Return(エラー時のメッセージは表示済)

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "Sub_CallFrmRset() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)
        End Try
    End Function
#End Region

#Region "FormLoaderAlarm実行サブルーチン"
    '''=========================================================================
    ''' <summary>FormLoaderAlarm実行サブルーチン</summary>
    ''' <param name="ObjSys">            (INP)OcxSystemオブジェクト</param>
    ''' <param name="AlarmKind">         (INP)アラーム種類(全停止異常, サイクル停止, 軽故障, アラーム無し)</param>
    ''' <param name="AlarmCount">        (INP)発生アラーム数</param>
    ''' <param name="strLoaderAlarm">    (INP)アラーム文字列</param>
    ''' <param name="strLoaderAlarmInfo">(INP)アラーム情報1(※未使用)</param>
    ''' <param name="strLoaderAlarmExec">(INP)アラーム情報(対策)</param>
    ''' <returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    '''
    Private Function Sub_CallFormLoaderAlarm(ByVal ObjSys As SystemNET, ByRef AlarmKind As Integer, ByVal AlarmCount As Integer,
                                    ByRef strLoaderAlarm() As String, ByRef strLoaderAlarmInfo() As String, ByRef strLoaderAlarmExec() As String) As Integer 'V6.0.0.0⑱

        Dim rtn As Integer = cFRS_NORMAL
        Dim r As Integer = cFRS_NORMAL
        Dim svAppMode As Integer = 0
        Dim strMSG As String

        Try
            ' ローダアラーム画面はDllSystemで用意したものを使用する

            '' ローダ無効NならOP
            'If (giLoaderType = 0) Then
            '    Return (cFRS_NORMAL)
            'End If

            '' アプリモードを「ローダアラーム表示」にする ###088
            'svAppMode = giAppMode
            'giAppMode = APP_MODE_LDR_ALRM
            ''@@@888 Call COVERCHK_ONOFF(COVER_CHECK_OFF)                        ' 「固定カバー開チェックなし」にする

            ' 電磁ロック(観音扉右側ロック)を解除する
            r = EL_Lock_OnOff(EX_LOK_MD_OFF)
            If (r = cFRS_TO_EXLOCK) Then                                ' 「前面扉ロック解除タイムアウト」なら戻り値を「RESET」にする
                r = cFRS_ERR_RST
                Return (r)
            End If
            If (r < cFRS_NORMAL) Then                                   ' 異常終了レベルのエラー ?
                Return (r)
            End If

            '' シグナルタワー制御(On=異常, Off=全ﾋﾞｯﾄ) ###191
            'Select Case (gSysPrm.stIOC.giSignalTower)
            '    Case SIGTOWR_NORMAL                                     ' 標準(赤点滅)
            '        'V5.0.0.9⑭ ↓　V6.0.3.0⑧(ローム殿仕様は赤点滅＋ブザーＯＮ)
            '        ' Call Form1.System1.SetSignalTower(SIGOUT_RED_BLK Or SIGOUT_BZ1_ON, &HFFFF)
            '        Call Form1.System1.SetSignalTowerCtrl(Form1.System1.SIGNAL_ALARM)
            '        'V5.0.0.9⑭ ↑　V6.0.3.0⑧

            '    Case SIGTOWR_SPCIAL                                     ' 特注(赤点滅+ブザー１)
            '        'r = Form1.System1.SetSignalTower(EXTOUT_RED_BLK Or EXTOUT_BZ1_ON, &HFFFF)
            'End Select

            '' アラームメッセージ表示中ON信号送出
            'Call W_ALM_DSP()

            'If giLotStopRemove = 1 Then
            '    SaveLotProcTime()
            'End If

            '' ローダアラーム画面を表示する
            'bFgLoaderAlarmFRM = True                                    ' ローダアラーム画面表示中ON

            'Dim objForm As New FormLoaderAlarm()
            'objForm.SetAlarmLevel(AlarmKind)
            'Call objForm.ShowDialog(Nothing, ObjSys, AlarmKind, AlarmCount, strLoaderAlarm, strLoaderAlarmInfo, strLoaderAlarmExec)
            'rtn = objForm.sGetReturn()                                  ' Return値取得
            'AlarmKind = objForm.GetAlarmLevel()

            '' オブジェクト開放
            'If (objForm Is Nothing = False) Then
            '    Call objForm.Close()                                    ' オブジェクト開放
            '    Call objForm.Dispose()                                  ' リソース開放
            'End If
            'bFgLoaderAlarmFRM = False                                   ' ローダアラーム画面表示中OFF

            '' アラームメッセージ表示中OFF信号送出
            '' Call W_ALM_DSP()                                            ' V1.18.0.0⑭

            '' シグナルタワー制御(On=0, Off=異常) ###191
            'Select Case (gSysPrm.stIOC.giSignalTower)
            '    Case SIGTOWR_NORMAL                                     ' 標準
            '        'V5.0.0.9⑭ ↓ V6.0.3.0⑧
            '        ' Call Form1.System1.SetSignalTower(0, SIGOUT_RED_BLK Or SIGOUT_BZ1_ON)
            '        Call Form1.System1.SetSignalTowerCtrl(Form1.System1.SIGNAL_ALL_OFF)
            '        'V5.0.0.9⑭ ↑ V6.0.3.0⑧

            '    Case SIGTOWR_SPCIAL                                     ' 特注(赤点滅+ブザー１)
            '        'r = Form1.System1.SetSignalTower(EXTOUT_RED_BLK Or EXTOUT_BZ1_ON, &HFFFF)
            'End Select

            '' 自動運転中(一時停止中以外)はシグナルタワー制御(自動運転中(緑点灯))を行う
            'If ((bFgAutoMode = True) And (gObjADJ Is Nothing = True)) Then
            '    ' シグナルタワー制御(On=自動運転中(緑点灯),Off=全ﾋﾞｯﾄ)
            '    'V5.0.0.9⑭ ↑ V6.0.3.0⑧
            '    ' Call Form1.System1.SetSignalTower(SIGOUT_GRN_ON, &HFFFF)
            '    Call Form1.System1.SetSignalTowerCtrl(Form1.System1.SIGNAL_OPERATION)
            '    'V5.0.0.9⑭ ↑ V6.0.3.0⑧

            'End If


            '' 筐体カバー閉を確認する
            'r = frmReset.Sub_CoverCheck()
            'If (r < cFRS_NORMAL) Then
            '    Return (r)
            'End If

            'giAppMode = svAppMode
            'Call COVERLATCH_CLEAR()                                     ' カバー開ラッチのクリア 
            'Call COVERCHK_ONOFF(COVER_CHECK_ON)                         ' 「固定カバー開チェックあり」にする

            '' 電磁ロック(観音扉右側ロック)する
            'If (giAppMode = APP_MODE_FINEADJ) Then
            '    ' 一時停止画面なら電磁ロック(観音扉右側ロック)を解除する
            '    r = EL_Lock_OnOff(EX_LOK_MD_OFF)
            'Else
            '    r = EL_Lock_OnOff(EX_LOK_MD_ON)
            'End If
            'If (r = cFRS_TO_EXLOCK) Then                                ' 「前面扉ロックタイムアウト」なら戻り値を「RESET」にする
            '    r = cFRS_ERR_RST
            '    Return (r)
            'End If
            'If (r < cFRS_NORMAL) Then                                   ' 異常終了レベルのエラー ?
            '    Return (r)
            'End If

            Return (rtn)                                                ' Return(エラー時のメッセージは表示済) 

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "Sub_CallFormLoaderAlarm() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)
        End Try
    End Function
#End Region


#Region "電磁ロック(観音扉右側ロック)をロックまたは解除する"
    '''=========================================================================
    ''' <summary>電磁ロック(観音扉右側ロック)をロックまたは解除する</summary>
    ''' <param name="Md">(INP)モード(0(EX_LOK_MD_OFF)=ロック解除,
    '''                               1(EX_LOK_MD_ON) =ロック)</param>
    ''' <returns>cFRS_NORMAL     = 正常
    '''          cFRS_TO_EXLOCK  = 前面扉ロックタイムアウト
    '''          上記以外      = その他のエラー
    ''' </returns>
    '''=========================================================================
    Public Function EL_Lock_OnOff(ByVal Md As Integer) As Integer

        Dim TimerLock As System.Threading.Timer = Nothing
        Dim sw As Long = 0
        Dim InterlockSts As Integer = 0
        Dim Sts As Integer = 0
        Dim r As Integer = cFRS_NORMAL
        Dim bTOut As Boolean
        Dim strTOUT As String
        Dim strMSG As String

        Try
            '-------------------------------------------------------------------
            '   初期処理
            '-------------------------------------------------------------------
            If (giLoaderType = 0) Then
                Return (cFRS_NORMAL)
            End If

            ' インターロック解除ならNOP
            r = INTERLOCK_CHECK(InterlockSts, sw)                       ' インターロック状態取得
            If (InterlockSts <> INTERLOCK_STS_DISABLE_NO) Then          ' インターロック中でない ?
                Return (cFRS_NORMAL)                                    ' Return値 = 正常
            End If

            ' タイムアウトチェック用タイマーオブジェクトの作成(TimerLock_TickをX msec間隔で実行する)
            TimerTM_Create(TimerLock, EX_LOK_TOUT)

STP_RETRY:
            '-------------------------------------------------------------------
            '   電磁ロック(観音扉右側ロック)をロックまたは解除する
            '-------------------------------------------------------------------
            If (Md = EX_LOK_MD_ON) Then                                 ' ロックモード ?
                Call EXTOUT1(EXTOUT_EX_LOK_ON, 0)                       ' 電磁ロック(観音扉右側ロック)をロックする
                strTOUT = My.Resources.MSG_DOORLOCK_TIMEOUT                                  '  "前面扉ロックタイムアウト"
            Else
                Call EXTOUT1(0, EXTOUT_EX_LOK_ON)                       ' 電磁ロック(観音扉右側ロック)を解除する
                strTOUT = My.Resources.MSG_DOORLOCKRELEASE_TIMEOUT                                  '  "前面扉ロック解除タイムアウト"
            End If

            '-------------------------------------------------------------------
            '   電磁ロックがロックまたは解除されたかチェックする
            '-------------------------------------------------------------------
            Do
                System.Threading.Thread.Sleep(100)                      ' Wait(ms)
                System.Windows.Forms.Application.DoEvents()

                ' 電磁ロックステータス取得
                Call INP16(EX_LOK_STS, Sts)

                ' 電磁ロック(観音扉右側ロック)をロックまたは解除されたかチェックする
                If (Md = EX_LOK_MD_ON) Then                             ' ロックモード ?
                    If ((Sts And EXTINP_EX_LOK_ON) = EXTINP_EX_LOK_ON) Then
                        Exit Do                                         ' 電磁ロックならExit
                    End If
                Else
                    If ((Sts And EXTINP_EX_LOK_ON) = 0) Then
                        Exit Do                                         ' 電磁ロック解除ならExit
                    End If
                End If

                '-------------------------------------------------------------------
                '   タイムアウトチェック
                '-------------------------------------------------------------------
                bTOut = TimerTM_Sts()
                If (bTOut = True) Then                                  ' タイムアウト ?
                    ' コールバックメソッドの呼出しを停止する
                    TimerTM_Stop(TimerLock)

                    ' ランプ制御
                    Call LAMP_CTRL(LAMP_START, True)                    ' STARTランプON
                    Call LAMP_CTRL(LAMP_RESET, True)                    ' RESETランプON
                    Call ZCONRST()                                      ' コンソールキーラッチ解除

                    '  "前面扉ロック(or 解除)タイムアウト" "STARTキー：処理続行，RESETキー：処理終了"
                    r = Sub_CallFrmMsgDisp(Form1.System1, cGMODE_MSG_DSP, cFRS_ERR_START + cFRS_ERR_RST, True,
                            strTOUT, My.Resources.MSG_SPRASH35, "", System.Drawing.Color.Blue, System.Drawing.Color.Black, System.Drawing.Color.Black)
                    ' 非常停止等のエラーならエラー戻り(エラーメッセージは表示済み)
                    If (r < cFRS_NORMAL) Then Exit Do
                    If (r = cFRS_ERR_RST) Then
                        Call ZCONRST()                                  ' コンソールキーラッチ解除
                        r = cFRS_TO_EXLOCK                              ' Return値 =  電磁ロックタイムアウト
                        Exit Do
                    End If

                    ' ランプ制御
                    Call LAMP_CTRL(LAMP_START, False)                   ' STARTランプOFF
                    Call LAMP_CTRL(LAMP_RESET, False)                   ' RESETランプOFF
                    Call ZCONRST()                                      ' コンソールキーラッチ解除

                    ' タイマー開始
                    Call TimerTM_Start(TimerLock, EX_LOK_TOUT)
                    GoTo STP_RETRY                                      ' リトライへ

                End If

            Loop While (1)

            '-------------------------------------------------------------------
            '   後処理
            '-------------------------------------------------------------------
            ' タイマーを破棄する
            TimerTM_Dispose(TimerLock)

            Return (r)

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "LoaderIOFor436.EL_Lock_OnOff() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)
        End Try
    End Function

#End Region

#Region "タイムアウトフラグを返す"
    '''=========================================================================
    ''' <summary>タイムアウトフラグを返す</summary>
    ''' <returns>Trur=タイムアウト, False=タイムアウトでない</returns>
    '''=========================================================================
    Public Function TimerTM_Sts() As Boolean

        Dim strMSG As String

        Try
            ' タイムアウトフラグを返す
            Return (bTmTimeOut)

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "globals.TimerTM_Sts() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (bTmTimeOut)
        End Try
    End Function
#End Region


#Region "汎用タイマー開始"
    '''=========================================================================
    ''' <summary>汎用タイマー開始</summary>
    ''' <param name="TimerTM">(INP)タイマー</param>
    '''=========================================================================
    Public Sub TimerTM_Start(ByRef TimerTM As System.Threading.Timer, ByVal TimeVal As Integer)

        Dim strMSG As String

        Try
            If (TimerTM Is Nothing) Then Return
            bTmTimeOut = False                                          ' タイムアウトフラグOFF V1.18.0.1⑧
            TimerTM.Change(TimeVal, TimeVal)
            Exit Sub

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "globals.TimerTM_Start() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "汎用タイマー停止(コールバックメソッド(TimerTM_Tick)の呼出しを停止する)"
    '''=========================================================================
    ''' <summary>汎用タイマー停止(コールバックメソッド(TimerTM_Tick)の呼出しを停止する)</summary>
    ''' <param name="TimerTM">(INP)タイマー</param>
    '''=========================================================================
    Public Sub TimerTM_Stop(ByRef TimerTM As System.Threading.Timer)

        Dim strMSG As String

        Try
            ' コールバックメソッドの呼出しを停止する
            If (TimerTM Is Nothing) Then Return
            TimerTM.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)
            Exit Sub

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "globals.TimerTM_Stop() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region


#Region "ローダアラームレベルの読込み"

    ''' <summary>
    ''' アラームレベルの読込み
    ''' </summary>
    ''' <param name="lData"></param>
    ''' <returns></returns>
    Public Function ReadAlarmLevel(ByVal lData As Long) As Integer                                 ' ローダアラーム状態取得

        Try




        Catch ex As Exception

        End Try


    End Function

#End Region


#Region "汎用タイマー"
    '===========================================================================
    '   汎用タイマー
    '===========================================================================
    Private bTmTimeOut As Boolean                                       ' タイムアウトフラグ

#Region "汎用タイマー生成"
    '''=========================================================================
    ''' <summary>汎用タイマー生成</summary>
    ''' <param name="TimerTM">(I/O)タイマー</param>
    ''' <param name="TimeVal">(INP)タイムアウト値(msec)</param>
    ''' <remarks>タイマー生成した場合はTimerTM_DisposeをCallしてタイマーを破棄する事</remarks>
    '''=========================================================================
    Public Sub TimerTM_Create(ByRef TimerTM As System.Threading.Timer, ByVal TimeVal As Integer)

        Dim strMSG As String

        Try
            ' タイムアウトチェック用タイマーオブジェクトの作成(TimerTM_TickをTimeVal msec間隔で実行する)
            bTmTimeOut = False                                          ' タイムアウトフラグOFF

            If (TimeVal = 0) Then
                TimerTM = New System.Threading.Timer(New System.Threading.TimerCallback(AddressOf TimerTM_Tick), Nothing, System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)
            Else
                TimerTM = New System.Threading.Timer(New System.Threading.TimerCallback(AddressOf TimerTM_Tick), Nothing, TimeVal, TimeVal)
            End If

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "globals.TimerTM_Create() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "タイマーイベント(指定タイマ間隔が経過した時に発生)"
    '''=========================================================================
    ''' <summary>タイマーイベント(指定タイマ間隔が経過した時に発生)</summary>
    ''' <param name="Sts">(INP)</param>
    '''=========================================================================
    Private Sub TimerTM_Tick(ByVal Sts As Object)

        Dim strMSG As String

        Try
            bTmTimeOut = True                                           ' タイムアウトフラグON
            Exit Sub

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "globals.TimerTM_Tick() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "汎用タイマーを破棄する"
    '''=========================================================================
    ''' <summary>汎用タイマーを破棄する</summary>
    ''' <param name="TimerTM">(I/O)タイマー</param>
    '''=========================================================================
    Public Sub TimerTM_Dispose(ByRef TimerTM As System.Threading.Timer)

        Dim strMSG As String

        Try
            ' コールバックメソッドの呼出しを停止する
            If (TimerTM Is Nothing) Then Return
            TimerTM.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)
            TimerTM.Dispose()                                           ' タイマーを破棄する
            Exit Sub

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "globals.TimerTM_Dispose() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "アラーム情報退避域クリア"
    '''=========================================================================
    ''' <summary>アラーム情報退避域クリア</summary>
    '''=========================================================================
    Public Sub ClearBefAlarm()

        Dim Len As Integer
        Dim i As Integer
        Dim strMSG As String

        Try

            ' アラーム情報退避域を初期化する
            Len = iBefData.Length
            For i = 0 To (Len - 1)
                iBefData(i) = 0
            Next

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "LoaderIOFor436.ClearBefAlarm() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#End Region


#Region "FrmResetを使用して指定のメッセージを表示する"
    '''=========================================================================
    ''' <summary>FrmResetを使用して指定のメッセージを表示する ###089</summary>
    ''' <param name="ObjSys">(INP)OcxSystemオブジェクト</param>
    ''' <param name="gMode"> (INP)処理モード</param>
    ''' <param name="Md">    (INP)cFRS_ERR_START                = STARTキー押下待ち
    '''                           cFRS_ERR_RST                  = RESETキー押下待ち
    '''                           cFRS_ERR_START + cFRS_ERR_RST = START/RESETキー押下待ち</param>
    ''' <param name="BtnDsp">(INP)ボタン表示する/しない</param>
    ''' <param name="Msg1">  (INP)表示メッセージ１</param>
    ''' <param name="Msg2">  (INP)表示メッセージ２</param>
    ''' <param name="MSG3">  (INP)表示メッセージ３</param>
    ''' <param name="Col1">  (INP)メッセージ色１</param>
    ''' <param name="Col2">  (INP)メッセージ色２</param>
    ''' <param name="Col3">  (INP)メッセージ色３</param>
    ''' <returns>cFRS_ERR_START = OKボタン(STARTキー)押下
    '''          cFRS_ERR_RST   = Cancelボタン(RESETキー)押下
    '''          上記以外       = エラー</returns>
    '''=========================================================================
    Public Function Sub_CallFrmMsgDisp(ByVal ObjSys As SystemNET, ByVal gMode As Integer, ByVal Md As Integer, ByVal BtnDsp As Boolean,
                                       ByVal Msg1 As String, ByVal Msg2 As String, ByVal Msg3 As String,
                                       ByVal Col1 As Color, ByVal Col2 As Color, ByVal Col3 As Color) As Integer

        Dim r As Integer
        Dim ColAry(3) As Color
        Dim MsgAry(3) As String
        Dim strMSG As String

        Try
            ' パラメータ設定
            'MsgAry(0) = Msg1
            'MsgAry(1) = Msg2
            'MsgAry(2) = Msg3
            'ColAry(0) = Col1
            'ColAry(1) = Col2
            'ColAry(2) = Col3

            r = ObjSys.Form_MsgDispBtnEx(Md, BtnDsp, Msg1, Msg2, Msg3, Col1, Col2, Col3, "OK", "Cancel", 1)

            ' DllSystemの共通I/Fを呼び出す 
            '' FrmReset画面表示(指定のメッセージを表示する)
            'Dim objForm As New frmReset()   'V6.0.0.0⑪
            'Call objForm.ShowDialog(Nothing, gMode, ObjSys, MsgAry, ColAry, Md, BtnDsp)
            'r = objForm.sGetReturn()                                    ' Return値取得

            '' オブジェクト開放
            'If (objForm Is Nothing = False) Then
            '    Call objForm.Close()                                    ' オブジェクト開放
            '    Call objForm.Dispose()                                  ' リソース開放
            'End If

            Return (r)                                                  ' Return(エラー時のメッセージは表示済)

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "LoaderIOFor436.Sub_CallFrmMsgDisp() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)
        End Try
    End Function
#End Region

#Region "サイクル停止処理"
    '''=========================================================================
    ''' <summary>サイクル停止処理</summary>
    ''' <param name="ObjSys">       (INP)OcxSystemオブジェクト</param>
    ''' <returns>cFRS_NORMAL     = 正常
    '''          cFRS_ERR_RST    = Cancel(RESETキー押下)
    '''          cFRS_TO_EXLOCK  = 前面扉ロックタイムアウト
    '''          上記以外        = その他のエラー
    ''' </returns>
    '''=========================================================================
    Public Function CycleStop_Proc(ByVal ObjSys As SystemNET) As Integer

        Dim RtnCode As Integer = cFRS_NORMAL
        Dim r As Integer = cFRS_NORMAL
        Dim strMSG As String
        Dim gObjMSG As FrmWait = Nothing

        Try
            '-------------------------------------------------------------------
            '   初期処理
            '-------------------------------------------------------------------

            ' サイクル停止要求信号を送出してローダからのサイクル停止応答を待つ
            If IsNothing(gObjMSG) = True Then
                gObjMSG = New FrmWait()
                Call gObjMSG.Show(Form1)
            End If


            r = W_CycleStop(ObjSys, 1)
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?(メッセージは表示済)
                If IsNothing(gObjMSG) <> True Then
                    gObjMSG.MsgClose()
                    gObjMSG = Nothing
                End If
                RtnCode = r                                             ' Return値設定
                GoTo STP_END                                            ' 処理終了へ
            End If

            Call Form1.System1.SetSignalTowerCtrl(Form1.System1.SIGNAL_CYCLE_STOP)

            Call LAMP_CTRL(LAMP_HALT, False)                            ' HALTランプOFF
            Call ZCONRST()                                              ' コンソールキーラッチ解除
            ' Call W_CLMP_ONOFF(0)                                        ' クランプOFF 'V5.0.0.4① １回目ＴＫＹがＩ／ＯでＯＮすると２回目ＰＬＣでＯＦＦ出来ないのでＯＦＦしておく。
            Call Form1.System1.ClampVacume_Ctrl(gSysPrm, 0, giAppMode, 0)
            '-------------------------------------------------------------------
            '   電磁ロック(観音扉右側ロック)を解除する
            '-------------------------------------------------------------------

STP_RETRY:
            RtnCode = cFRS_NORMAL
            Call COVERCHK_ONOFF(COVER_CHECK_OFF)                        '「固定カバー開チェックなし」にする
            r = EL_Lock_OnOff(EX_LOK_MD_OFF)                            ' 電磁ロック(観音扉右側ロック)を解除する
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?(メッセージは表示済)
                RtnCode = r                                             ' Return値設定
                GoTo STP_END                                            ' 処理終了へ
            End If
            Call ZCONRST()

            If IsNothing(gObjMSG) <> True Then
                gObjMSG.MsgClose()
                gObjMSG = Nothing
            End If

            '-------------------------------------------------------------------
            '   「サイクル停止中」メッセージを表示して「STARTキー」「RESETキー」
            '    の押下待ち(この間に基板を取り出して確認が可能)
            '-------------------------------------------------------------------
            Dim md As Short = giAppMode
            giAppMode = APP_MODE_IDLE

            ' メッセージ表示(STARTキー, RESETキー押下待ち)
            ' "サイクル停止中", "OKボタン押下で自動運転を続行します", "Cancelボタン押下で自動運転を終了します"
            r = Sub_CallFrmMsgDisp(ObjSys, cGMODE_MSG_DSP, cFRS_ERR_START + cFRS_ERR_RST, True,
                    My.Resources.MSG_LOADER_48, My.Resources.MSG_LOADER_45, My.Resources.MSG_LOADER_46, System.Drawing.Color.Blue, System.Drawing.Color.Blue, System.Drawing.Color.Blue)

            giAppMode = md
            If (r < cFRS_NORMAL) Then                                   ' 非常停止等のエラーならエラー戻り(エラーメッセージは表示済み)
                RtnCode = r                                             ' Return値設定
                GoTo STP_END                                            ' 処理終了へ
            ElseIf (r = cFRS_ERR_RST) Then                              ' Cancel(RESETキー押下) ?
                RtnCode = cFRS_ERR_RST                                  ' Return値 = Cancel(RESETキー押下)
            End If

            '-------------------------------------------------------------------
            '   電磁ロック(観音扉右側ロック)をロックする
            '-------------------------------------------------------------------
            ' 筐体カバー閉を確認する
            r = ObjSys.Sub_CoverCheck(gSysPrm, 0, False)
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?(メッセージは表示済)
                RtnCode = r                                             ' Return値設定
                GoTo STP_END                                            ' 処理終了へ
            End If

            ' 「固定カバー開チェックあり」にする
            Call COVERLATCH_CLEAR()                                     ' カバー開ラッチのクリア
            Call COVERCHK_ONOFF(COVER_CHECK_ON)                         '「固定カバー開チェックあり」にする

            Call W_RESET()                                              ' アラームリセット信号送出'
            Call W_START()                                              ' スタート信号送出'

            ' 電磁ロック(観音扉右側ロック)をロックする
            r = EL_Lock_OnOff(EX_LOK_MD_ON)                             ' 電磁ロック
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?(メッセージは表示済)
                RtnCode = r                                             ' Return値設定
                GoTo STP_END                                            ' 処理終了へ
            End If

            '-------------------------------------------------------------------
            '   「STARTキー押下時」は基板ありを確認
            '   「RESETキー押下時」は基板なしを確認
            '-------------------------------------------------------------------
            If (RtnCode = cFRS_ERR_RST) Then                            ' RESETキー押下 ?
STP_010_PLATENON_CHECK:
                ' 載物台に基板がない事を確認(載物台に基板がある場合は、取り除かれるまで待つ)
                r = SubstrateNottingCheck(Form1.System1, APP_MODE_FINEADJ)     ' ※ 一時停止画面モードでCall
                If (r < cFRS_NORMAL) Then                               ' 非常停止等のエラーならエラー戻り(エラーメッセージは表示済み)
                    RtnCode = r                                         ' Return値設定
                    GoTo STP_END                                        ' 処理終了へ
                End If
            Else                                                        ' STARTキー押下時
STP_010_PLATEEXIST_CHECK:
                ' 載物台の基板あり/なしを確認
                RtnCode = Sub_SubstrateExistCheck(ObjSys)              ' RtnCode = cFRS_NORMAL(基板あり続行), cFRS_ERR_START(基板なし続行), cFRS_ERR_RST(基板なしでCancel(RESETキー押下)
                If (RtnCode < cFRS_NORMAL) Then                         ' 非常停止等のエラーならエラー戻り(エラーメッセージは表示済み)
                    GoTo STP_END                                        ' 処理終了へ
                End If
                '「基板なし続行」指定なら載物台に基板がない事を確認へ
                If (RtnCode = cFRS_ERR_START) Then
                    GoTo STP_010_PLATENON_CHECK
                End If
                ' 割欠検出なら再度頭からチェック
                If (RtnCode = cFRS_ERR_HALT) Then
                    GoTo STP_RETRY
                End If
            End If

            '-------------------------------------------------------------------
            '   後処理
            '-------------------------------------------------------------------
STP_END:
            ' サイクル停止要求信号をOFFする
            r = W_CycleStop(ObjSys, 0)
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?(メッセージは表示済)
                RtnCode = r                                             ' Return値設定
            End If

            ' シグナルタワー制御(On=自動運転中(緑点灯),Off=全ﾋﾞｯﾄ)
            Call Form1.System1.SetSignalTowerCtrl(Form1.System1.SIGNAL_OPERATION)

            ' PLC側と同期を取るため吸着OFF(I/O)を行う
            Call ZABSVACCUME(0)                                         ' バキュームの制御(吸着OFF)

            Return (RtnCode)

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "LoaderIOFor436.CycleStop_Proc() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            RtnCode = cERR_TRAP
            GoTo STP_END
        End Try
    End Function
#End Region

#Region "載物台に基板がある事をチェックする(サイクル停止中の続行指定時 )"
    '''=========================================================================
    ''' <summary>載物台に基板がある事をチェックする(サイクル停止中の続行指定時用)</summary>
    ''' <param name="ObjSys"> (INP)OcxSystemオブジェクト</param>
    ''' <returns>cFRS_NORMAL    = 正常(基板あり続行)
    '''          cFRS_ERR_START = 正常(基板なし続行)
    '''          cFRS_ERR_RST   = 基板なしでCancel(RESETキー押下)
    '''          cFRS_ERR_HALT  = 割欠検出
    '''          上記以外=エラー</returns>
    ''' <remarks>ローム殿特注(SL436R/SL436S)</remarks>
    '''=========================================================================
    Public Function Sub_SubstrateExistCheck(ByVal ObjSys As SystemNET) As Integer

        Dim lData As Long = 0
        Dim r As Integer = cFRS_NORMAL
        Dim rtn As Integer = cFRS_NORMAL
        Dim strMSG As String = ""
        Dim strMS2 As String = ""
        Dim strMS3 As String = ""
        Dim bFlg As Boolean = True

        Try
            ' 載物台に基板がある事をチェックする
            If (gSysPrm.stIOC.giClamp = 1) Then
                ' 載物台クランプON   
                r = Form1.System1.ClampCtrl(gSysPrm, 1, 0)
                If (r <> cFRS_NORMAL) Then

                End If

                System.Threading.Thread.Sleep(gSysPrm.stIOC.glClampWait) ' Wait(ms)
                Call ZABSVACCUME(1)                                     ' (クランプOFFで基板がづれるのをふせぐため)
                r = Form1.System1.ClampCtrl(gSysPrm, 0, 0)
                If (r <> cFRS_NORMAL) Then

                End If
            End If

            System.Threading.Thread.Sleep(500)                          ' Wait(ms) ※200msだとワーク有が検出されない場合がある
            ' 吸着状態の取得   : 

            ' 基板がある場合はクランプOFFしない
            If (gSysPrm.stIOC.giClamp = 1) Then
                '                                     ' クランプON
                r = Form1.System1.ClampCtrl(gSysPrm, 1, 0)

                ' If (lData And 1) Then                         ' 載物台に基板有 ?
                If ObjSys.getStageVaccumDisp() Then
                    '                                                   ' クランプOFFしない
                Else
                    System.Threading.Thread.Sleep(gSysPrm.stIOC.glClampWait) ' Wait(ms)
                    ' Call W_CLMP_ONOFF(0)                                ' クランプOFF
                    r = Form1.System1.ClampCtrl(gSysPrm, 0, 0)
                    If (r <> cFRS_NORMAL) Then
                        '                        r = W_Read(LOFS_W44, lData)                             ' 物理入力状態取得(B14C2)
                    End If

                End If

            End If

            ' 基板がある場合は吸着OFFしない
            ' If (lData And 1) Then                         ' 載物台に基板有 ?
            If ObjSys.getStageVaccumDisp() Then

            Else
                Call ZABSVACCUME(0)                                     ' バキュームの制御(吸着OFF)
            End If

            ' 「固定カバー開チェックなし」にする
            Call COVERCHK_ONOFF(COVER_CHECK_OFF)
            r = EL_Lock_OnOff(EX_LOK_MD_OFF)                            ' 電磁ロック(観音扉右側ロック)を解除する
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?(メッセージは表示済)
                Return (r)
            End If

            ' メッセージ表示
            ' If (lData And 1) Then                         ' 載物台に基板有 ?
            If ObjSys.getStageVaccumDisp() Then
                rtn = cFRS_NORMAL                                       ' Return値 = 正常(基板あり続行)

            Else                                                        ' 載物台に基板がない場合

                If frmAutoObj.gbFgAutoOperation = True Then

                    ' メッセージ表示(STARTキー, RESETキー押下待ち)
                    ' "載物台に基板がありません", "OKボタン押下で自動運転を続行します", "Cancelボタン押下で自動運転を終了します"
                    r = Sub_CallFrmMsgDisp(ObjSys, cGMODE_MSG_DSP, cFRS_ERR_START + cFRS_ERR_RST, True,
                            My.Resources.MSG_LOADER_42, My.Resources.MSG_LOADER_45, My.Resources.MSG_LOADER_46, System.Drawing.Color.Blue, System.Drawing.Color.Blue, System.Drawing.Color.Blue)
                Else
                    ' メッセージ表示(STARTキー,キー押下待ち)
                    ' "載物台に基板がありません"
                    r = Sub_CallFrmMsgDisp(ObjSys, cGMODE_MSG_DSP, cFRS_ERR_START, True,
                            My.Resources.MSG_LOADER_42, "", "", System.Drawing.Color.Blue, System.Drawing.Color.Blue, System.Drawing.Color.Blue)
                    r = cFRS_ERR_RST
                    ' 「固定カバー開チェックあり」にする
                    Call COVERLATCH_CLEAR()                                     ' カバー開ラッチのクリア
                    Call COVERCHK_ONOFF(COVER_CHECK_ON)                         '「固定カバー開チェックあり」にする
                    Call ZCONRST()                                              ' コンソールキーラッチ解除

                    Return (r)
                End If

                If (r < cFRS_NORMAL) Then Return (r) '                  ' 非常停止等のエラーならエラー戻り(エラーメッセージは表示済み)

                If (r = cFRS_ERR_RST) Then
                    rtn = cFRS_ERR_RST                                  ' Return値 = 基板なしでCancel(RESETキー押下)
                Else
                    rtn = cFRS_ERR_START                                ' Return値 = 正常(基板なし続行)
                End If
            End If

            ' 筐体カバー閉を確認する
            r = ObjSys.Sub_CoverCheck(gSysPrm, 0, False)

            ' 「固定カバー開チェックあり」にする
            Call COVERLATCH_CLEAR()                                     ' カバー開ラッチのクリア
            Call COVERCHK_ONOFF(COVER_CHECK_ON)                         '「固定カバー開チェックあり」にする
            Call ZCONRST()                                              ' コンソールキーラッチ解除

            ' 電磁ロック(観音扉右側ロック)をロックする
            r = EL_Lock_OnOff(EX_LOK_MD_ON)                             ' 電磁ロック
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?(メッセージは表示済)
                rtn = r                                                 ' Return値設定
            End If

            Return (rtn)

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "LoaderIOFor436.Sub_SubstrateExistCheck() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)
        End Try
    End Function
#End Region


#Region "サイクル停止信号を送出してローダからのサイクル停止応答を待つ(オートローダシリアル通信)"
    '''=========================================================================
    ''' <summary>サイクル停止要求信号を送出してローダからのサイクル停止応答を待つ</summary>
    ''' <param name="ObjSys">(INP)OcxSystemオブジェクト</param>
    ''' <param name="OnOff"> (INP)1=サイクル停止信号ON送出, 0=サイクル停止信号OFF送出</param>
    ''' <returns>cFRS_NORMAL    = 正常(基板あり続行)
    '''          cFRS_ERR_START = 正常(基板なし続行)
    '''          cFRS_ERR_RST   = Cancel(RESETキー押下)
    '''          cFRS_ERR_LDR0  = ローダ通信タイムアウト他</returns>
    ''' <remarks>ローム殿特注(SL436R/SL436S)</remarks>
    '''=========================================================================
    Public Function W_CycleStop(ByVal ObjSys As SystemNET, ByVal OnOff As Integer) As Integer

        Dim TimerLock As System.Threading.Timer = Nothing
        Dim rtnCode As Integer = cFRS_NORMAL
        Dim r As Integer = cFRS_NORMAL
        Dim TimeVal As Integer
        Dim bTOut As Boolean
        Dim lData As Long = 0
        Dim LdIn As UShort = 0
        Dim strMSG As String

        Try
#If cPLCcOFFLINEcDEBUG Then
            Return (cFRS_NORMAL)
#End If
            '-------------------------------------------------------------------
            '   初期処理
            '-------------------------------------------------------------------

RETRY_CYCLE:
            rtnCode = cFRS_NORMAL
            '-------------------------------------------------------------------
            '   サイクル停止ON/OFF信号送信(SL436S/SL436R)
            '-------------------------------------------------------------------
            ' タイマー値を設定する
            If (giOPLDTimeOutFlg = 0) Then                              ' ローダ通信タイムアウト検出無し ?
                TimeVal = System.Threading.Timeout.Infinite             ' タイマー値 = なし
            Else
                TimeVal = giOPLDTimeOut                                 ' タイマー値 = ローダ通信タイムアウト時間(msec)
            End If

            If (OnOff = 1) Then                                     ' サイクル停止要求BIT ON ?
                ' サイクル停止要求BIT ON
                Call Sub_ATLDSET(LOUT_CYCL_STOP, &H0)               ' ローダ出力(ON=サイクル停止要求, OFF=なし)
            Else
                ' サイクル停止要求BIT OFF
                Call Sub_ATLDSET(&H0, LOUT_CYCL_STOP)               ' ローダ出力(ON=なし, OFF=サイクル停止要求)
            End If

            ' サイクル停止OFFなら応答待ちしない
            If (OnOff = 0) Then
                Return (cFRS_NORMAL)
            End If


            '-------------------------------------------------------------------
            '   タイムアウトチェック用タイマーオブジェクトの作成
            '   (TimerLock_TickをX msec間隔で実行する)
            '-------------------------------------------------------------------
            TimerTM_Create(TimerLock, TimeVal)

            '-------------------------------------------------------------------
            '   サイクル停止応答待ち(SL436S/SL436R)
            '-------------------------------------------------------------------
            Do
                System.Threading.Thread.Sleep(10)                       ' Wait(ms)
                System.Windows.Forms.Application.DoEvents()

                ' サイクル停止応答待ち
                ' ローダアラーム/非常停止チェック
                r = GetLoaderIO(LdIn)                               ' ローダI/O入力
                If ((LdIn And LINP_NO_ALM_RESTART) <> LINP_NO_ALM_RESTART) Then
                    rtnCode = cFRS_ERR_LDR                          ' Return値 = ローダアラーム検出
                    GoTo STP_ERR_LDR                                ' ローダアラーム表示へ
                End If
                ' サイクル停止応答かチェックする(SL436S時)
                If ((LdIn And LIN_CYCL_STOP) = LIN_CYCL_STOP) Then
                    Exit Do                                         ' サイクル停止応答ならExit
                End If

                '-------------------------------------------------------------------
                '   タイムアウトチェック
                '-------------------------------------------------------------------
                bTOut = TimerTM_Sts()
                If (bTOut = True) Then                                  ' タイムアウト ?
                    rtnCode = cFRS_ERR_LDRTO                            ' Retuen値 = ローダ通信タイムアウト
                    Exit Do
                End If

            Loop While (1)

            '-------------------------------------------------------------------
            '   後処理
            '-------------------------------------------------------------------
STP_END:
            TimerTM_Stop(TimerLock)                                     ' コールバックメソッドの呼出しを停止する
            TimerTM_Dispose(TimerLock)                                  ' タイマーを破棄する

            If (rtnCode = cFRS_ERR_LDRTO) Then                          ' ローダ通信タイムアウト ?
                AutoOperationDebugLogOut("W_CycleStop() - rtnCode = cFRS_ERR_LDRTO")       ''V2.2.1.3②
                rtnCode = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_TMOUT)

            End If

            Return (rtnCode)

            ' ローダエラー発生時
STP_ERR_LDR:
            TimerTM_Stop(TimerLock)                                     ' コールバックメソッドの呼出しを停止する
            TimerTM_Dispose(TimerLock)                                  ' タイマーを破棄する
            If (rtnCode = cFRS_ERR_LDRTO) Then                          ' ローダ通信タイムアウト ?
                ' rtnCode = Sub_CallFrmRset(ObjSys, cGMODE_LDR_TMOUT)     ' エラーメッセージ表示
                ' rtnCode = ObjSys.Sub_CallFormLoaderAlarm(cGMODE_LDR_TMOUT, ObjPlcIf)
                AutoOperationDebugLogOut("W_CycleStop() STP_ERR_LDR - rtnCode = cFRS_ERR_LDRTO")       ''V2.2.1.3②

                rtnCode = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_TMOUT)
            Else
                ' ローダアラームメッセージ作成 & ローダアラーム画面表示
                rtnCode = ObjSys.Sub_CallFormLoaderAlarm(cGMODE_LDR_ALARM, ObjPlcIf)
                If (rtnCode = cFRS_ERR_LDR3) Or (rtnCode = cFRS_ERR_LDR2) Then ' 
                    Call W_RESET()                                      ' アラームリセット信号送出
                    Call W_START()                                      ' スタート信号送出
                    GoTo RETRY_CYCLE
                End If
            End If
            Call Sub_ATLDSET(&H0, LOUT_AUTO)        ' ローダ手動モード切替え(ローダ出力(ON=なし, OFF=自動))

            Return (rtnCode)

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "LoaderIOFor436.W_CycleStop() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            rtnCode = cERR_TRAP
            GoTo STP_END
        End Try
    End Function
#End Region

#Region "載物台の基板なしをチェックする"
    '''=========================================================================
    ''' <summary>原点復帰時の載物台の基板なしをチェックする</summary>
    ''' <param name="ObjSys">(INP)OcxSystemオブジェクト</param>
    ''' <param name="Mode">   (INP)APP_MODE_LOADERINIT = ローダ原点復帰時
    '''                            APP_MODE_AUTO　　　 = 自動運転時
    '''                            APP_MODE_FINEADJ 　 = 一時停止(サイクル停止でCancel指定時) 
    ''' </param>2
    ''' <returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Public Function SubstrateNottingCheck(ByVal ObjSys As SystemNET, ByVal Mode As Integer) As Integer

        Dim r As Integer = cFRS_NORMAL
        Dim rtn As Integer = cFRS_NORMAL
        Dim strMSG As String = ""

        Try
            ' 載物台に基板がある場合、取り除かれるまで待つ(SL436R時)
            If (gSysPrm.stTMN.gsKeimei <> MACHINE_TYPE_SL436) Then Return (cFRS_NORMAL)

            Do
                System.Threading.Thread.Sleep(300)                      ' Wait(ms)
                System.Windows.Forms.Application.DoEvents()

                ' 載物台に基板がない事をチェックする
                Call ZCONRST()                                          ' コンソールキーラッチ解除
                r = Sub_SubstrateExistCheckForCycle(ObjSys)                    ' 基板なしチェック
                If (r = cFRS_NORMAL) Then Exit Do '                     ' 基板なしならループを抜ける
                If (r < cFRS_NORMAL) Then                               ' 異常終了レベルのエラー ?
                    Return (r)                                          ' 呼び出し元でアプリケーション強制終了へ
                End If
                '                                                       ' 基板有(r = cFRS_ERR_RST)なら取り除かれるまで待つ
            Loop While (1)

            Return (rtn)

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "LoaderIOFor436.SubstrateCheck() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)
        End Try
    End Function
#End Region

#Region "基板排出位置への移動"
    ''' <summary>
    ''' 基板排出位置への移動
    ''' </summary>
    ''' <returns></returns>
    Public Function MoveGlassOutPos() As Integer
        Dim r As Integer
        Dim rtnCode As Integer = cFRS_NORMAL

        Try

            '基板品種は1を使用
            r = SMOVE2(ObjLoader.gfBordTableOutPosX(0), ObjLoader.gfBordTableOutPosY(0))
            If (r <> cFRS_NORMAL) Then
                rtnCode = Form1.System1.EX_ZGETSRVSIGNAL(gSysPrm, r, 0) ' メッセージ表示
                Return (rtnCode)

            End If

            Return r

        Catch ex As Exception

        End Try

    End Function

#End Region

#Region "基板1枚の搬送処理"
    ''' <summary>
    ''' 基板1枚の搬送処理
    ''' </summary>
    ''' <returns></returns>
    Public Function LoaderGlassHandlingProc(ByVal objSys As SystemNET) As Integer

        Dim r As Integer
        Dim bFgLot As Boolean = False                                   ' ロット切替え要求フラグ
        Dim bFgMagagin As Boolean = False                               ' マガジン終了フラグﾞ
        Dim bFgAllMagagin As Boolean = False                            ' 全マガジン終了フラグﾞ


        Try

            Call Sub_ATLDSET(LOUT_STS_RUN Or LOUT_STOP, LOUT_REDY)                        ' ローダ出力(ON=なし, OFF=トリマ停止中) 

            AutoOperationDebugLogOut("LoaderGlassHandlingProc() - Start")

            ' ローダへ基板要求(基板投入/交換要求信号)を送信し、トリミングスタート信号を待つ
            r = Loader_WaitTrimStart(objSys, frmAutoObj.gbFgAutoOperation, m_lTrimResult, bFgMagagin, bFgAllMagagin, bFgLot, gbIniFlg)

            AutoOperationDebugLogOut("LoaderGlassHandlingProc() - End")

            gbIniFlg = 1
            Return r

        Catch ex As Exception

        End Try

    End Function

#End Region



#Region "オートローダ原点復帰実行サブルーチン"
    '''=========================================================================
    ''' <summary>オートローダ原点復帰実行サブルーチン</summary>
    ''' <param name="igMode">(INP)処理モード</param>
    ''' '                         ※下記を想定
    '''                             cGMODE_ORG     = 原点復帰
    '''                             cGMODE_LDR_ORG = ローダ原点復帰
    ''' <returns>cFRS_NORMAL   = 正常
    '''          cFRS_ERR_LDR  = ローダアラーム検出
    '''          cFRS_ERR_LDR1 = ローダアラーム検出(全停止異常)
    '''          cFRS_ERR_LDR2 = ローダアラーム検出(サイクル停止)
    '''          cFRS_ERR_LDR3 = ローダアラーム検出(軽故障(続行可能))
    ''' </returns>
    '''=========================================================================
    Public Function Sub_Loader_OrgBack(ByVal igMode As Integer) As Integer

        Dim TimerRS As System.Threading.Timer = Nothing
        Dim rtnCode As Integer = cFRS_NORMAL
        Dim r As Integer                                                ' ###163
        Dim LdIn As Integer
        'V6.0.0.0⑱        Dim objForm As Object = Nothing
        Dim strMSG As String            'V6.3.2.0⑧
        Dim WaitKey As Integer          'V6.3.2.0⑧
        Dim RetKey As Integer

        Try

            '' シグナルタワー制御(On=原点復帰中,Off=全ﾋﾞｯﾄ)
            'Select Case (gSysPrm.stIOC.giSignalTower)
            '    Case SIGTOWR_NORMAL                                 ' 標準(原点復帰中(緑点滅))
            '        Call Form1.System1.SetSignalTowerCtrl(Form1.System1.SIGNAL_ZRN)

            '    Case SIGTOWR_SPCIAL                                 ' 特注(原点復帰中(黄色点滅))

            'End Select

            ' ダミーでシリアル通信実施しPLCとの通信状態をチェックする


            ' ローダ出力(On=トリマ部レディ+トリマ正常, Off=左記以外)
            Call Sub_ATLDSET(0, LOUT_REDY)

            ' ローダ原点復帰処理
            Call Sub_ATLDSET(LOUT_ORG_BACK, &H0)                        ' ローダ出力(On=ローダ原点復帰要求, Off=0)

            ' 前回のローダ原点復帰完了BitのOffを待つ
STP_RETRY:
            rtnCode = Sub_WaitLoaderData(Form1.System1, LINP_ORG_BACK, False, False, False, False, RetKey)

            If (rtnCode <> cFRS_NORMAL) Then                            ' エラー ? (※エラー発生時のメッセージは表示済)
                'If (rtnCode = cFRS_ERR_LDR3) Then                       ' 軽故障(続行可能) ? ###196 ###073
                If (rtnCode = cFRS_ERR_LDR3) Or (rtnCode = cFRS_ERR_LDR2) Then  ' ###196
                    Call W_RESET()                                      ' アラームリセット信号送出
                    Call W_START()                                      ' スタート信号送出
                    GoTo STP_RETRY
                End If
                ' トリマ運転中(原点復帰後は原則ONのまま、一時停止はローダの基板交換動作中に停止したい場合に使用(使用しない?))
                Call Sub_ATLDSET(LOUT_STS_RUN, LOUT_ORG_BACK)               ' ローダ出力(On=トリマ運転中, Off=ローダ原点復帰要求)
                Return (rtnCode)
            End If

            AutoOperationDebugLogOut("Sub_Loader_OrgBack() - 1")       ''V2.2.1.3②

            ' ローダ通信タイムアウトチェック用タイマーオブジェクトの作成(TimerRS_TickをX msec間隔で実行する)
            Sub_SetTimeoutTimer(TimerRS)

            ' ローダの原点復帰完了を待つ
            Do
                ' ローダアラーム/非常停止チェック
                ' Call GetLoaderIO(LdIn)                                  ' ローダ入力
                ' オートローダ入力
                r = ZATLDGET(LdIn)

                If ((LdIn And LINP_NO_ALM_RESTART) <> LINP_NO_ALM_RESTART) Then
                    'Return (cFRS_ERR_LDR)                              ' ###163 Return値 = ローダアラーム検出
                    rtnCode = cFRS_ERR_LDR                              ' ###163
                    GoTo STP_END                                        ' ###163
                End If

                ' ローダ通信タイムアウトチェック
                If (bFgTimeOut = True) Then                             ' タイムアウト ?
                    ' コールバックメソッドの呼出しを停止する
                    TimerRS.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)
                    rtnCode = cFRS_ERR_LDRTO                            ' Return値 = ローダ通信タイムアウト
                    GoTo STP_END
                End If

                ' 非常停止等チェック(トリマ装置アイドル中)
                r = Form1.System1.Sys_Err_Chk_EX(gSysPrm, giAppMode)           ' 非常停止/ｶﾊﾞｰ/ｴｱｰ圧/集塵機/ﾏｽﾀｰﾊﾞﾙﾌﾞﾁｪｯｸ
                If (r <> cFRS_NORMAL) Then                                  ' 非常停止等検出 ?
                    rtnCode = cFRS_ERR_EMG                            ' Return値 = ローダ通信タイムアウト
                    GoTo STP_END                                    ' エラーメッセージ表示へ
                End If

                '---------------------------------------------------------------------------k
                '   非常停止チェック
                '---------------------------------------------------------------------------
                r = Form1.System1.EmergencySwCheck()
                If r <> cFRS_NORMAL Then ' 非常停止 ?
                    rtnCode = cFRS_ERR_EMG                            ' Return値 = ローダ通信タイムアウト
                    GoTo STP_END                                    ' エラーメッセージ表示へ
                End If

                ''V4.1.0.0⑩
                ' インターロック状態の表示およびローダへ通知(SL436R) ###162
                r = Form1.DispInterLockSts()
                'V4.1.0.0⑩

                System.Windows.Forms.Application.DoEvents()
                Call System.Threading.Thread.Sleep(1)                   ' Wait(msec)

                ' 'V2.2.1.1⑦↓
                ' ローダ側の装置と干渉する部分の原点復帰が完了のフラグをチェックする
                If giLoaderType <> 0 Then
                    If ((LdIn And LINP_ORG_BACK) = LINP_ORG_BACK) Then       ' 原点復帰完了待ち
                        Exit Do
                    End If
                    ' 原点復帰可能のフラグがONしているかチェック
                    If (Form1.System1.getTrimOriginPossibleStatus()) Then
                        Exit Do
                    End If
                End If
                ' 'V2.2.1.1⑦↑

            Loop While ((LdIn And LINP_ORG_BACK) <> LINP_ORG_BACK)      ' 原点復帰完了待ち


            ' 終了処理
STP_END:

            ' 'V2.2.1.1⑦↓
            ' ローダ側の装置と干渉する部分の原点復帰が完了のフラグをチェックする
            If giLoaderType <> 0 Then
                ' TLF製ローダの場合は完全に原点復帰が終了してからフラグを立てる。 
            Else
                ''V4.1.0.0⑨
                ' トリマ運転中(原点復帰後は原則ONのまま、一時停止はローダの基板交換動作中に停止したい場合に使用(使用しない?))
                ' Call Sub_ATLDSET(LOUT_STS_RUN, LOUT_ORG_BACK)               ' ローダ出力(On=トリマ運転中, Off=ローダ原点復帰要求)
                Call Form1.System1.Z_ATLDSET(LOUT_STS_RUN, LOUT_ORG_BACK)
                ''V4.1.0.0⑨
            End If
            ' 'V2.2.1.1⑦↑

            ' コールバックメソッドの呼出しを停止する
            If (IsNothing(TimerRS) = False) Then                        ' ###173
                AutoOperationDebugLogOut("Sub_Loader_OrgBack() - IsNothing(TimerRS) = False")       ''V2.2.1.3②
                TimerRS.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)
                TimerRS.Dispose()                                       ' タイマーを破棄する
            End If

            'ローダ原点復帰タイムアウトの場合には、
            If rtnCode = cFRS_ERR_LDRTO Then
                AutoOperationDebugLogOut("Sub_Loader_OrgBack() - rtnCode = cFRS_ERR_LDRTO")       ''V2.2.1.3②
                Return (rtnCode)
            ElseIf rtnCode = cFRS_ERR_LDR Then
                ' "ローダ原点復帰未完了","",""
                r = Sub_CallFrmMsgDisp(Form1.System1, cGMODE_MSG_DSP, WaitKey, True,
                    My.Resources.MSG_LDALARM_11, "", "", System.Drawing.Color.Red, System.Drawing.Color.Red, System.Drawing.Color.Red)
                Return (rtnCode)
            ElseIf rtnCode = cFRS_ERR_EMG Then

                Return (rtnCode)
            End If

        ' ローダエラーならステージを原点に戻す(残基板取り除くため)
        If (rtnCode <> cFRS_NORMAL) Then

            ' XYZθ軸初期化
            r = Form1.System1.EX_SYSINIT(gSysPrm, stPLT.Z_ZOFF, stPLT.Z_ZON)
            If (r <> cFRS_NORMAL) Then                              ' エラー ? (※メッセージは表示済)
                Return (r)
            End If

        End If

        ' シグナルタワー制御(On=レディ(手動),Off=全ﾋﾞｯﾄ) ###007
        Select Case (gSysPrm.stIOC.giSignalTower)
            Case SIGTOWR_NORMAL                                     ' 標準(無点灯)
                'V5.0.0.9⑭ ↓ V6.0.3.0⑧
                ' Call Form1.System1.SetSignalTower(0, &HEFFF)
                Call Form1.System1.SetSignalTowerCtrl(Form1.System1.SIGNAL_ALL_OFF)
                    'V5.0.0.9⑭ ↑ V6.0.3.0⑧

            Case SIGTOWR_SPCIAL                                     ' 特注(黄色点灯)
                'Call Form1.System1.SetSignalTower(EXTOUT_YLW_ON, &HEFFF)
        End Select

        Return (rtnCode)

        ' トラップエラー発生時
        Catch ex As Exception
        strMSG = "LoaderIOFor436.Sub_Loader_OrgBack() TRAP ERROR = " + ex.Message
        MsgBox(strMSG)
        Return (cERR_TRAP)
        End Try
    End Function
#End Region


#Region "自動運転終了メッセージ表示およびシグナルタワー制御(全マガジン終了,自動運転OFF)"
    '''=========================================================================
    ''' <summary>自動運転終了メッセージ表示(自動運転時)および
    '''          シグナルタワー制御(全マガジン終了,自動運転OFF)</summary>
    ''' <param name="ObjSys">       (INP)OcxSystemオブジェクト</param>
    ''' <param name="bFgAutoMode">  (OUT)ローダ自動モードフラグ</param>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          cFRS_ERR_EMG  = 非常停止</returns>
    '''=========================================================================
    Public Function Loader_EndAutoDrive(ByVal ObjSys As SystemNET) As Integer

        Dim rtnCode As Integer = cFRS_NORMAL
        Dim strMSG As String

        Try

            ChkProcFile()

            Call Form1.System1.SetSignalTowerCtrl(Form1.System1.SIGNAL_OPERATION_END)

            ' 自動運転終了メッセージ表示(自動運転時)

            FrmMessageDisp(ObjSys, cGMODE_LDR_END, cFRS_ERR_START, True,
                             "自動運転終了", My.Resources.MSG_frmLimit_07, "", System.Drawing.Color.Black, System.Drawing.Color.Black, System.Drawing.Color.Black)

            Call Form1.System1.SetSignalTowerCtrl(Form1.System1.SIGNAL_ALL_OFF)

            ' 終了処理
STP_END:
            Return (rtnCode)

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "LoaderIOFor436.Loader_EndAutoDrive() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)
        End Try
    End Function
#End Region


#Region "載物台に基板があったら取り除く "
    '''=========================================================================
    ''' <summary>載物台に基板があったら取り除く</summary>
    ''' <param name="ObjSys"> (INP)OcxSystemオブジェクト</param>
    ''' <returns>cFRS_NORMAL    = 正常(基板あり続行)
    '''          cFRS_ERR_START = 正常(基板なし続行)
    '''          cFRS_ERR_RST   = 基板なしでCancel(RESETキー押下)
    '''          cFRS_ERR_HALT  = 割欠検出
    '''          上記以外=エラー</returns>
    ''' <remarks>ローム殿特注(SL436R/SL436S)</remarks>
    '''=========================================================================
    Public Function Sub_SubstrateNothingCheck(ByVal ObjSys As SystemNET) As Integer

        Dim lData As Long = 0
        Dim r As Integer = cFRS_NORMAL
        Dim rtn As Integer = cFRS_NORMAL
        Dim strMSG As String = ""
        Dim strMS2 As String = ""
        Dim strMS3 As String = ""
        Dim bFlg As Boolean = True

        Try
            ' 載物台に基板がある事をチェックする
            If (gSysPrm.stIOC.giClamp = 1) Then
                ' 載物台クランプON      Call W_CLMP_ONOFF(1)                                    ' クランプON
                r = Form1.System1.ClampCtrl(gSysPrm, 1, 0)
                If (r <> cFRS_NORMAL) Then

                End If

                System.Threading.Thread.Sleep(gSysPrm.stIOC.glClampWait) ' Wait(ms)
                Call ZABSVACCUME(1)                                     ' (クランプOFFで基板がづれるのをふせぐため)
                r = Form1.System1.ClampCtrl(gSysPrm, 0, 0)
                If (r <> cFRS_NORMAL) Then

                End If
            End If

            System.Threading.Thread.Sleep(500)                          ' Wait(ms) ※200msだとワーク有が検出されない場合がある
            ' 吸着状態の取得   

            ' 基板がある場合はクランプOFFしない
            If (gSysPrm.stIOC.giClamp = 1) Then
                '     ' クランプON
                r = Form1.System1.ClampCtrl(gSysPrm, 1, 0)

                ' If (lData And 1) Then                         ' 載物台に基板有 ?
                If ObjSys.getStageVaccumDisp() Then
                    '                                                   ' クランプOFFしない
                Else
                    System.Threading.Thread.Sleep(gSysPrm.stIOC.glClampWait) ' Wait(ms)
                    ' Call W_CLMP_ONOFF(0)                                ' クランプOFF
                    r = Form1.System1.ClampCtrl(gSysPrm, 0, 0)
                    If (r <> cFRS_NORMAL) Then
                        '                        r = W_Read(LOFS_W44, lData)                             ' 物理入力状態取得(B14C2)
                    End If

                End If

            End If

            ' 基板がある場合は吸着OFFしない
            ' If (lData And 1) Then                         ' 載物台に基板有 ?
            If ObjSys.getStageVaccumDisp() Then

            Else
                Call ZABSVACCUME(0)                                     ' バキュームの制御(吸着OFF)
            End If

            ' 「固定カバー開チェックなし」にする
            Call COVERCHK_ONOFF(COVER_CHECK_OFF)

            ' メッセージ表示
            If ObjSys.getStageVaccumDisp() Then
                ' メッセージ表示(STARTキー, RESETキー押下待ち)
                ' "載物台の基板を取り除いてください",
                r = Sub_CallFrmMsgDisp(ObjSys, cGMODE_MSG_DSP, cFRS_ERR_START, True,
                        "", My.Resources.MSG_LOADER_50, "", System.Drawing.Color.Blue, System.Drawing.Color.Blue, System.Drawing.Color.Blue)
                If (r < cFRS_NORMAL) Then Return (r) '                  ' 非常停止等のエラーならエラー戻り(エラーメッセージは表示済み)
                If (r = cFRS_ERR_RST) Then
                    rtn = cFRS_ERR_RST                                  ' Return値 = 基板なしでCancel(RESETキー押下)
                Else
                    rtn = cFRS_ERR_START                                ' Return値 = 正常(基板なし続行)
                End If

            Else                                                        ' 載物台に基板がない場合
                rtn = cFRS_NORMAL                                       ' Return値 = 正常(基板あり続行)

            End If

            ' 筐体カバー閉を確認する
            r = ObjSys.Sub_CoverCheck(gSysPrm, 0, False)

            r = Form1.System1.ClampCtrl(gSysPrm, 0, 0)
            Call ZABSVACCUME(0)                                     ' バキュームの制御(吸着OFF)

            ' 「固定カバー開チェックあり」にする
            Call COVERLATCH_CLEAR()                                     ' カバー開ラッチのクリア
            Call COVERCHK_ONOFF(COVER_CHECK_ON)                         '「固定カバー開チェックあり」にする
            Call ZCONRST()                                              ' コンソールキーラッチ解除

            Return (rtn)

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "LoaderIOFor436.Sub_SubstrateExistCheck() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)
        End Try
    End Function
#End Region


#Region "PLCからタクト関連、作業マガジン、段数を読込み "
    ''' <summary>
    ''' PLCからタクト関連、作業マガジン、段数を読込み
    ''' </summary>
    ''' <returns></returns>
    Public Function DispLoaderInfo() As Integer

        ' 基板交換時間書込み
        Dim Trimtime As Integer
        Dim SupplyMag As Integer = 0
        Dim SupplySlot As Integer = 0
        Dim StoreMag As Integer = 0
        Dim StoreSlot As Integer = 0


        ObjSys.Sub_GetProcessTime(gitacktTime, gichangePlateTime, Trimtime)
        gichangePlateTime = gitacktTime - Trimtime
        ObjSys.Sub_SetChangePlateTime(gichangePlateTime)

        objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_TACT, gitacktTime)
        objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_EXCHANGE, gichangePlateTime)
        objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_TRIMMING, Trimtime)

        ObjSys.Sub_GetNowProcessMgInfo(SupplyMag, SupplySlot, StoreMag, StoreSlot)
        objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_SUPPLY_MAGAGINE, SupplyMag)
        objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_SUPPLY_SLOT, SupplySlot)
        ''V2.2.0.037　objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_STORE_MAGAGINE, StoreMag)
        ''V2.2.0.037　objLoaderInfo.UpdateLoaderInfo(frmLoaderInfo.LoaderDispMode.DISP_STORE_SLOT, StoreSlot)

    End Function

#End Region

#Region "ロット中断のフラグをONする "
    ''' <summary>
    ''' ロット中断のフラグをONする
    ''' </summary>
    ''' <param name="mode"></param>
    ''' <returns></returns>
    Public Function SetLotAbort(ByVal mode As Integer) As Integer
        giLotAbort = mode
    End Function
#End Region


#Region "ロット中断フラグの状態を取得する  "
    ''' <summary>
    ''' ロット中断フラグの状態を取得する 
    ''' </summary>
    ''' <returns></returns>
    Public Function GetLotAbort() As Integer
        GetLotAbort = giLotAbort
    End Function
#End Region

#Region "載物台に基板がない事をチェックする"     'V2.2.0.0⑦
    '''=========================================================================
    ''' <summary>載物台に基板がない事をチェックする</summary>
    ''' <param name="ObjSys"> (INP)OcxSystemオブジェクト</param>
    ''' <returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Public Function Sub_SubstrateExistCheckForCycle(ByVal ObjSys As SystemNET) As Integer

        Dim lData As Long = 0
        Dim lBit As Long = 0
        Dim r As Integer = cFRS_NORMAL
        Dim rtn As Integer = cFRS_NORMAL
        Dim strMSG As String = ""
        Dim strMS2 As String = ""
        Dim strMS3 As String = ""
        Dim bFlg As Boolean = True
        Dim WaitKey As Integer       'V6.3.2.0⑧
        Dim vac As Integer

        Try

            ' クランプレス時は念のためクランプOFFする
            If (gSysPrm.stIOC.giClamp = 2) Then
                r = Form1.System1.ClampCtrl(gSysPrm, 0, 0)
                If (r <> cFRS_NORMAL) Then

                End If
            End If

            ' 載物台に基板がない事をチェックする
            '----- V1.16.0.0⑫↓ -----
            If (gSysPrm.stIOC.giClamp = 1) Then
                ' Call W_CLMP_ONOFF(1)                                    ' クランプON
                r = Form1.System1.ClampCtrl(gSysPrm, 1, 0)
                If (r <> cFRS_NORMAL) Then

                End If
                System.Threading.Thread.Sleep(gSysPrm.stIOC.glClampWait) ' Wait(ms)
                ''V6.2.1.0②　クランプONは上のみでしている                Call W_CLMP_ONOFF(0)                                    ' クランプOFF
            End If

            Call ZABSVACCUME(1)                                         ' バキュームの制御(吸着ON)
            System.Threading.Thread.Sleep(500)                          ' Wait(ms) ※200msだとワーク有が検出されない場合がある

            vac = ObjSys.getStageVaccumDisp()

            If (gSysPrm.stIOC.giClamp = 1) Then
                r = Form1.System1.ClampCtrl(gSysPrm, 0, 0)
            End If

            Call ZABSVACCUME(0)                                         ' バキュームの制御(吸着OFF)

            ' ワーク有ならメッセージ表示
            '            If (lData And lBit) Then                                    ' 載物台に基板有 ? V2.0.0.0⑥
            If vac <> 0 Then

                ' 「固定カバー開チェックなし」にする
                Call COVERCHK_ONOFF(COVER_CHECK_OFF)
                ' 電磁ロック(観音扉右側ロック)を解除する(一時停止(サイクル停止でCancel指定時)) ローム殿特注(SL436R/SL436S)
                r = EL_Lock_OnOff(EX_LOK_MD_OFF)                        ' 電磁ロック(観音扉右側ロック)を解除する
                If (r <> cFRS_NORMAL) Then                              ' エラー ?(メッセージは表示済)
                    rtn = r                                         ' Return値設定
                    Return (rtn)
                End If

                ' モードによって表示メッセージを設定する
                strMSG = My.Resources.MSG_LOADER_36                              ' "載物台の基板を取り除いて下さい"
                strMS2 = ""                                         ' ""
                strMS3 = ""                                         ' ""

                WaitKey = cFRS_ERR_START

                ' メッセージ表示(STARTキー押下待ち)
                r = Sub_CallFrmMsgDisp(ObjSys, cGMODE_MSG_DSP, WaitKey, True,
                        strMSG, strMS2, strMS3, System.Drawing.Color.Blue, System.Drawing.Color.Blue, System.Drawing.Color.Blue)
                'V6.3.2.0⑧↑
                If (r < cFRS_NORMAL) Then Return (r) '                  ' 非常停止等のエラーならエラー戻り(エラーメッセージは表示済み)
                rtn = cFRS_ERR_RST                                      ' Return値 = Cancel(RESETｷｰ)
            End If

            ' 筐体カバー閉を確認する
            r = ObjSys.Sub_CoverCheck(gSysPrm, 0, False)

            ' 「固定カバー開チェックあり」にする
            Call ZCONRST()                                              ' コンソールキーラッチ解除
            Call COVERLATCH_CLEAR()                                     ' カバー開ラッチのクリア
            Call COVERCHK_ONOFF(COVER_CHECK_ON)                         ' 「固定カバー開チェックあり」にする


            ' 電磁ロック(観音扉右側ロック)をロックする
            r = EL_Lock_OnOff(EX_LOK_MD_ON)                             ' 電磁ロック
            If (r <> cFRS_NORMAL) Then                                  ' エラー ?(メッセージは表示済)
                rtn = r                                                 ' Return値設定
            End If

            Return (rtn)

            ' トラップエラー発生時
        Catch ex As Exception
            strMSG = "LoaderIOFor436.Sub_SubstrateExistCheckForCycle() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)
        End Try
    End Function
#End Region
    ''' <summary>
    ''' ローダの停止を待つ
    ''' </summary>
    ''' <returns></returns>
    Public Function WaitLoaderStop() As Integer
        Dim r As Integer
        Dim LdIn As Integer
        Dim TimerRS As System.Threading.Timer = Nothing
        Dim rtnCode As Integer
        Dim AlarmKind As Integer


        Try

            AutoOperationDebugLogOut("WaitLoaderStop() - start ")       ''V2.2.1.3②

            ' ローダ通信タイムアウトチェック用タイマーオブジェクトの作成(TimerRS_TickをX msec間隔で実行する)
            Sub_SetTimeoutTimer(TimerRS)

            ' ローダからの応答データを待つ
            Do
                ' ローダアラーム/非常停止チェック
                r = GetLoaderIO(LdIn)                                   ' ローダ
                If ((LdIn And LINP_NO_ALM_RESTART) <> LINP_NO_ALM_RESTART) Then
                    rtnCode = cFRS_ERR_LDR                              ' Return値 = ローダアラーム検出
                    GoTo STP_ERR_LDR                                    ' ローダアラーム表示へ
                End If

                ' ローダ通信タイムアウトチェック
                If (bFgTimeOut = True) Then                             ' タイムアウト ?
                    ' コールバックメソッドの呼出しを停止する
                    TimerRS.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)
                    rtnCode = cFRS_ERR_LDRTO                            ' Return値 = ローダ通信タイムアウト
                    GoTo STP_ERR_LDR                                    ' エラーメッセージ表示へ
                End If

                ' 非常停止等チェック(トリマ装置アイドル中)
                r = Form1.System1.Sys_Err_Chk_EX(gSysPrm, giAppMode)           ' 非常停止/ｶﾊﾞｰ/ｴｱｰ圧/集塵機/ﾏｽﾀｰﾊﾞﾙﾌﾞﾁｪｯｸ
                If (r <> cFRS_NORMAL) Then                                  ' 非常停止等検出 ?
                    rtnCode = cFRS_ERR_EMG                            ' Return値 = ローダ通信タイムアウト
                    GoTo STP_ERR_LDR                                    ' エラーメッセージ表示へ
                End If

                ' ローダの停止を待つ
                r = ObjSys.getLoaderActStatus()
                If (r = 0) Then
                    Exit Do
                End If

                System.Windows.Forms.Application.DoEvents()
                Call System.Threading.Thread.Sleep(1)                   ' Wait(msec)
            Loop While (True)                 ' 応答データ待ち


STP_END:

            TimerTM_Stop(TimerRS)                                     ' コールバックメソッドの呼出しを停止する
            TimerTM_Dispose(TimerRS)                                  ' タイマーを破棄する
            AutoOperationDebugLogOut("WaitLoaderStop() TimerTM_Dispose normal")       ''V2.2.1.3②
            Return cFRS_NORMAL

STP_ERR_LDR:
            TimerTM_Stop(TimerRS)                                     ' コールバックメソッドの呼出しを停止する
            TimerTM_Dispose(TimerRS)                                  ' タイマーを破棄する
            AutoOperationDebugLogOut("WaitLoaderStop() TimerTM_Dispose STP_ERR_LDR:")       ''V2.2.1.3②
            If (rtnCode = cFRS_ERR_LDRTO) Then                          ' ローダ通信タイムアウト ?
                ' rtnCode = Sub_CallFrmRset(ObjSys, cGMODE_LDR_TMOUT)     ' エラーメッセージ表示
                ' rtnCode = ObjSys.Sub_CallFormLoaderAlarm(cGMODE_LDR_TMOUT, ObjPlcIf)
                AutoOperationDebugLogOut("WaitLoaderStop() STP_ERR_LDR - rtnCode = cFRS_ERR_LDRTO")       ''V2.2.1.3②

                rtnCode = ObjSys.Sub_CallFrmRsetLoader(ObjPlcIf, cGMODE_LDR_TMOUT)
            Else
                ' ローダアラームメッセージ作成 & ローダアラーム画面表示
                rtnCode = ObjSys.Sub_CallFormLoaderAlarm(AlarmKind, ObjPlcIf)
                rtnCode = AlarmKind     ' アラームレベル: 


                If (rtnCode = cFRS_ERR_LDR3) Or (rtnCode = cFRS_ERR_LDR2) Then ' 
                    Call W_RESET()                                      ' アラームリセット信号送出
                    Call W_START()                                      ' スタート信号送出
                End If
            End If
            Call Sub_ATLDSET(&H0, LOUT_AUTO)        ' ローダ手動モード切替え(ローダ出力(ON=なし, OFF=自動))

            Return rtnCode

        Catch ex As Exception

            AutoOperationDebugLogOut("WaitLoaderStop() Catch ex As Exception")       ''V2.2.1.3②

        End Try


    End Function


    ''' <summary>
    ''' ロットで指定したファイルが全て処理されたかチェックする 
    ''' </summary>
    ''' <returns></returns>
    Public Function ChkProcFile() As Integer
        Dim i As Integer

        Try

            Dim Datano As Integer = frmAutoObj.GetNowLotDataNo() + 1

            If frmAutoObj.giAutoDataFileNum > Datano Then

                Z_PRINT("指定したファイルが処理されずにロットが終了しました。")

                For i = Datano To frmAutoObj.giAutoDataFileNum - 1

                    Z_PRINT("ファイル名：" & (frmAutoObj.gsAutoDataFileFullPath(i)))

                Next i

            End If

        Catch ex As Exception

        End Try


    End Function

    ' 'V2.2.1.1⑦↓
    ''' <summary>
    '''     ' ローダ側の装置と干渉する部分の原点復帰が完了のフラグをチェックする 
    ''' </summary>
    ''' <returns></returns>
    Public Function WaitLoaderOrigin() As Integer
        Dim r As Integer
        Dim LdIn As Integer
        Dim TimerRS As System.Threading.Timer = Nothing
        Dim rtnCode As Integer
        Dim WaitKey As Integer = cFRS_ERR_START

        Try

            AutoOperationDebugLogOut("WaitLoaderOrigin() - start ")       ''V2.2.1.3②

            ' ローダ通信タイムアウトチェック用タイマーオブジェクトの作成(TimerRS_TickをX msec間隔で実行する)
            Sub_SetTimeoutTimer(TimerRS)

            ' ローダの原点復帰完了を待つ
            Do
                ' ローダアラーム/非常停止チェック
                ' Call GetLoaderIO(LdIn)                                  ' ローダ入力
                ' オートローダ入力
                r = ZATLDGET(LdIn)

                If ((LdIn And LINP_NO_ALM_RESTART) <> LINP_NO_ALM_RESTART) Then
                    'Return (cFRS_ERR_LDR)                              ' ###163 Return値 = ローダアラーム検出
                    rtnCode = cFRS_ERR_LDR                              ' ###163
                    GoTo STP_END                                        ' ###163
                End If

                ' ローダ通信タイムアウトチェック
                If (bFgTimeOut = True) Then                             ' タイムアウト ?
                    ' コールバックメソッドの呼出しを停止する
                    TimerRS.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)
                    rtnCode = cFRS_ERR_LDRTO                            ' Return値 = ローダ通信タイムアウト
                    GoTo STP_END
                End If

                ' 非常停止等チェック(トリマ装置アイドル中)
                r = Form1.System1.Sys_Err_Chk_EX(gSysPrm, giAppMode)           ' 非常停止/ｶﾊﾞｰ/ｴｱｰ圧/集塵機/ﾏｽﾀｰﾊﾞﾙﾌﾞﾁｪｯｸ
                If (r <> cFRS_NORMAL) Then                                  ' 非常停止等検出 ?
                    rtnCode = cFRS_ERR_EMG                            ' Return値 = ローダ通信タイムアウト
                    GoTo STP_END                                    ' エラーメッセージ表示へ
                End If

                '---------------------------------------------------------------------------k
                '   非常停止チェック
                '---------------------------------------------------------------------------
                r = Form1.System1.EmergencySwCheck()
                If r <> cFRS_NORMAL Then ' 非常停止 ?
                    rtnCode = cFRS_ERR_EMG                            ' Return値 = ローダ通信タイムアウト
                    GoTo STP_END                                    ' エラーメッセージ表示へ
                End If

                ''V4.1.0.0⑩
                ' インターロック状態の表示およびローダへ通知(SL436R) ###162
                r = Form1.DispInterLockSts()
                'V4.1.0.0⑩

                System.Windows.Forms.Application.DoEvents()
                Call System.Threading.Thread.Sleep(1)                   ' Wait(msec)

            Loop While ((LdIn And LINP_ORG_BACK) <> LINP_ORG_BACK)      ' 原点復帰完了待ち


STP_END:
            ' 'V2.2.1.1⑦↓
            ' ローダ側の装置と干渉する部分の原点復帰が完了のフラグをチェックする
            If giLoaderType <> 0 Then
                ' TLF製ローダの場合は完全に原点復帰が終了してからフラグを立てる。 
                ''V4.1.0.0⑨
                ' トリマ運転中(原点復帰後は原則ONのまま、一時停止はローダの基板交換動作中に停止したい場合に使用(使用しない?))
                ' Call Sub_ATLDSET(LOUT_STS_RUN, LOUT_ORG_BACK)               ' ローダ出力(On=トリマ運転中, Off=ローダ原点復帰要求)
                Call Form1.System1.Z_ATLDSET(LOUT_STS_RUN, LOUT_ORG_BACK)
                ''V4.1.0.0⑨
            End If
            ' 'V2.2.1.1⑦↑

            ' コールバックメソッドの呼出しを停止する
            If (IsNothing(TimerRS) = False) Then
                AutoOperationDebugLogOut("WaitLoaderOrigin() - IsNothing(TimerRS) = False ")       ''V2.2.1.3②
                TimerRS.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)
                TimerRS.Dispose()                                       ' タイマーを破棄する
            End If

            'ローダ原点復帰タイムアウトの場合には、
            If rtnCode = cFRS_ERR_LDRTO Then
                Return (rtnCode)
            ElseIf rtnCode = cFRS_ERR_LDR Then
                ' "ローダ原点復帰未完了","",""
                r = Sub_CallFrmMsgDisp(Form1.System1, cGMODE_MSG_DSP, WaitKey, True,
                    My.Resources.MSG_LDALARM_11, "", "", System.Drawing.Color.Red, System.Drawing.Color.Red, System.Drawing.Color.Red)
                Return (rtnCode)
            ElseIf rtnCode = cFRS_ERR_EMG Then

                Return (rtnCode)
            End If


        Catch ex As Exception
            AutoOperationDebugLogOut("WaitLoaderOrigin() - Catch ex As Exception")       ''V2.2.1.3②

        End Try

    End Function
    ' 'V2.2.1.1⑦↑

    ''' <summary>
    ''' ロット切り替え信号の設定   'V2.2.1.1⑧
    ''' </summary>
    ''' <param name="count"></param>
    ''' <returns></returns>
    Public Function SetLotChangeFlg(ByVal count As Integer) As Integer

        Try

            giLotChangeFlg = count                                    'ロット切り替え実行フラフ  


        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' ロット切り替え信号の取得   'V2.2.1.1⑧
    ''' </summary>
    ''' <param name="count"></param>
    ''' <returns></returns>
    Public Function GetLotChangeFlg() As Integer

        Try

            Return giLotChangeFlg                                   'ロット切り替え実行フラフ  


        Catch ex As Exception

        End Try

    End Function

End Class
