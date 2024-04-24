'==============================================================================
'   Description : ユーザプログラム用固有ファンクション
'
'　 2012/11/16 First Written by N.Arata(OLFT)
'
'==============================================================================
Option Strict Off
Option Explicit On

Imports System
Imports System.Drawing.Printing
Imports System.IO
Imports System.Text
Imports LaserFront.Trimmer.DefWin32Fnc
Imports LaserFront.Trimmer.DllLaserTeach.ctl_LaserTeach

Module UserSub
    Private bStartCheck As Boolean          ' データ設定確認が必要な時　Trueとする。
    Private dInitialResValue As Double      ' 初期測定値
    Private dStandardResValue As Double     ' 標準抵抗測定値
    Private lResCounterForPrinter As Long   ' 印刷用素子カウンタ

    Private intTMM_Save As Integer          ' 保存用　モード(0:高速(コンパレータ非積分モード), 1:高精度(積分モード))
    Private intMType_Save As Integer        ' 保存用　測定種別(0=内部測定, 1=外部測定)
    Private dTRV As Double                  ' 目標抵抗値
    Private bOkJudge As Boolean             ' 素子単位のNG判定
    Private bSkip As Boolean                ' ネットワーク抵抗のスキップ
    Private sResistorPrintData(MAX_RES_USER) As String    ' ネットワーク抵抗の時の出力データ

    Public Printer As New cPrintDocument    ' ﾌﾟﾘﾝﾀｰｵﾌﾞｼﾞｪｸﾄ
    '===============================================================================
    ' 印刷用データ領域
    '===============================================================================
    Private Const cTRIM_PRINT_DATA_HEAD As String = "C:\TRIMDATA\PRINTDATA\TRIM_PRINT_DATA_HEAD.TXT"
    Private Const cTRIM_PRINT_DATA_RES As String = "C:\TRIMDATA\PRINTDATA\TRIM_PRINT_DATA_RES.TXT"
    Private Const cTRIM_PRINT_DATA_PLATE As String = "C:\TRIMDATA\PRINTDATA\TRIM_PRINT_DATA_PLATE.TXT"
    Private Const cTRIM_PRINT_DATA_END As String = "C:\TRIMDATA\PRINTDATA\TRIM_PRINT_DATA_END.TXT"

    '''===============================================================================
    ''' <summary>
    ''' 印刷用素子カウンタのリセット
    ''' </summary>
    ''' <remarks></remarks>
    '''===============================================================================
    Public Sub ResetlResCounterForPrinter()
        lResCounterForPrinter = 0
    End Sub

    '===============================================================================
    '【機　能】 抵抗温度係数の算出
    '【引　数】 スタンダード抵抗値０℃、スタンダード抵抗値２５℃
    '【戻り値】 抵抗温度係数
    '===============================================================================
    Public Function GetResTempCoff(ByVal dStandardRes0 As Double, ByVal dStandardRes25 As Double) As Double
        GetResTempCoff = (dStandardRes25 - dStandardRes0) / dStandardRes0 * 10.0 ^ 6 / 25.0
    End Function

    '===============================================================================
    '【機　能】 ユーザ設定画面確認
    '【引　数】 true , false
    '【戻り値】 無し
    '===============================================================================
    Public Sub SetStartCheckStatus(ByVal bCheck As Boolean)
        bStartCheck = bCheck
    End Sub
    Public Function GetStartCheckStatus() As Boolean
        Return (bStartCheck)
    End Function

    '===============================================================================
    '【機　能】 標準抵抗値の抵抗値算出
    '【引　数】 抵抗番号,カット番号
    '【戻り値】 目標抵抗値
    '===============================================================================
    Public Function CalcStandardResistanceValue() As Double

        'V2.0.0.0⑪        CalcStandardResistanceValue = stUserData.dStandardRes25

        'V2.0.0.0⑪↓
        'STD抵抗値(25℃)=STD(0℃)抵抗値×(1+α×25+β×25^2)
        Dim dAlpha As Double = stUserData.dAlpha / 10.0 ^ 6
        Dim dBeta As Double = stUserData.dBeta / 10.0 ^ 6
        Dim dTemp As Double = 25.0
        CalcStandardResistanceValue = stUserData.dTemperatura0 * (1 + dAlpha * dTemp + dBeta * dTemp ^ 2)
        DebugLogOut("STD抵抗値(25℃)[" & CalcStandardResistanceValue.ToString & "]=STD(0℃)抵抗値[" & stUserData.dTemperatura0.ToString & "]*(1+α[" & dAlpha.ToString & "]*[" & dTemp.ToString & "]+β[" & dBeta.ToString & "]*[" & dTemp.ToString & "]^2)")
        'V2.0.0.0⑪↑

        ' A0023NI.BAS の プログラムの場合
        ' 21730  IF RTP%=2 THEN NT#=25:MSRV#=SRV# ELSE NT#=0:MSRV#=SRV#*(1+SNTC#*NST#)
        'If stUserData.iTempTemp = 2 Then
        '    CalcStandardResistanceValue = stUserData.dStandardRes0
        'Else
        '    CalcStandardResistanceValue = stUserData.dStandardRes0 * (1.0 + stUserData.dResTempCoff * 25.0)
        'End If
    End Function

    '''===============================================================================
    ''' <summary>
    ''' スタンダード抵抗のチェック
    ''' </summary>
    ''' <returns>True:正常 False:スタンダード抵抗測定値異常</returns>
    ''' <remarks></remarks>
    '''===============================================================================
    Public Function StandardResistanceMeasure() As Boolean
        Dim rn As Integer = 1
        Dim Rtn As Short
        Dim dblMx As Double
        Dim strJUG As String
        Dim Judge As Integer                                            ' 判定結果'V2.0.0.0⑨

        Try

            If Not UserSub.IsTrimType1() And Not UserSub.IsTrimType4() Then
                Return (True)
            End If

            If stREG(rn).intSLP <> SLP_RMES Then       ' 測定
                Return (True)
            End If
            'V2.0.0.0①↓
            For i As Short = 1 To stPLT.RCount
                If stREG(i).intSLP = SLP_RMES Then                      ' 抵抗測定のみ
                    stREG(i).dblNOM = CalcStandardResistanceValue()     ' スタンダード抵抗測定全てに設定する。
                End If
            Next
            'V2.0.0.0①↑
            'V2.0.0.0①            stREG(rn).dblNOM = CalcStandardResistanceValue()
            Call DScanModeResetSet(rn, 0, 0)                             ' DCスキャナに接続する測定器を切替る 
            Rtn = V_R_MEAS(stREG(rn).intSLP, stREG(rn).intMType, dblMx, rn, stREG(rn).dblNOM)
            If (Rtn <> cFRS_NORMAL) Then
                Call Z_PRINT("スタンダード抵抗が測定できません" & vbCrLf)
                Return (False)
            Else
                ' 目標値判定処理(FT)
                strJUG = Test_ItFt(1, stREG(rn).intMode, dblMx, stREG(rn).dblNOM, stREG(rn).dblITL, stREG(rn).dblITH, Judge)    'V2.0.0.0⑨Judge追加
                If (strJUG <> JG_OK) Then                           ' FT-NG ?
                    Call Z_PRINT("スタンダード抵抗を確認してください 測定値 ＝ " & dblMx.ToString("0.00000") & "Ω" & vbCrLf)
                    Return (False)
                Else
                    Return (True)
                End If
            End If

        Catch ex As Exception
            Call Z_PRINT("UserSub.StandardResistanceMeasure() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Function

    '''===============================================================================
    ''' <summary>
    ''' 目標値の取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    '''===============================================================================
    Public Function GetTRV() As Double
        Return (dTRV)
    End Function

    '===============================================================================
    '【機　能】 目標値算出
    '【引　数】 抵抗番号,カット番号
    '【戻り値】 目標抵抗値
    '===============================================================================
    Public Sub CalcTargeResistancetValue(ByVal rn As Integer)

        Try
            If IsTrimType1() Or UserSub.IsTrimType4() Then  'V2.0.0.0①sTrimType4()追加

#If OLD_CALURATION Then 'V2.0.0.0⑪
                ' 温度センサーの場合　　　　　　　：　TRV　＝　スタンダード実測値 ／ スタンダード（０℃or２５℃）
                If stUserData.iTempTemp = 1 Then        ' 参照温度	１：０℃
                    dTRV = dStandardResValue / stUserData.dStandardRes0 * stREG(rn).dblNOM
                    DebugLogOut("TRV:" & dTRV.ToString & " = " & dStandardResValue.ToString & " / " & stUserData.dStandardRes0.ToString & " * " & stREG(rn).dblNOM.ToString)
                Else                                    ' 参照温度	２：２５℃
                    dTRV = dStandardResValue / stUserData.dStandardRes25 * stREG(rn).dblNOM
                    DebugLogOut("TRV:" & dTRV.ToString & " = " & dStandardResValue.ToString & " / " & stUserData.dStandardRes25.ToString & " * " & stREG(rn).dblNOM.ToString)
                End If
#Else
                Dim dAlpha As Double = stUserData.dAlpha / 10.0 ^ 6
                Dim dBeta As Double = stUserData.dBeta / 10.0 ^ 6
                Dim dDaihyouAlpha As Double = stUserData.dDaihyouAlpha / 10.0 ^ 6
                Dim dDaihyouBeta As Double = stUserData.dDaihyouBeta / 10.0 ^ 6

                'ステージ温度換算計算式=(-α+SQRT(α^2-4*β*(1-STD実測値/STD0℃抵抗値)))/(2*β)
                Dim dStageTempConv As Double = (-1.0 * dAlpha + Math.Sqrt(dAlpha ^ 2 - 4.0 * dBeta * (1.0 - dStandardResValue / stUserData.dTemperatura0))) / (2 * dBeta)
                DebugLogOut("ステージ温度[" & dStageTempConv.ToString & "]= (-1.0 * " & dAlpha.ToString & " + Sqrt(" & dAlpha.ToString & " ^ 2 - 4.0 * " & dBeta.ToString & " * (1.0 - " & dStandardResValue.ToString & " / " & stUserData.dTemperatura0.ToString & "))) / (2 * " & dBeta.ToString & ")")

                'センサー計算式(＝トリミング時の目標値) = (設定抵抗値 / (1 + α * 設定温度 + β * 設定温度 ^ 2)) * (1 + α * ステージ温度 + β * ステージ温度 ^ 2)
                dTRV = (stREG(rn).dblNOM / (1 + dDaihyouAlpha * stUserData.iTempTemp + dDaihyouBeta * stUserData.iTempTemp ^ 2)) * (1 + dDaihyouAlpha * dStageTempConv + dDaihyouBeta * dStageTempConv ^ 2)
                DebugLogOut("TRV[" & dTRV.ToString & "] = (" & stREG(rn).dblNOM.ToString & " / (1 + " & dDaihyouAlpha.ToString & " * " & stUserData.iTempTemp.ToString & " + " & dDaihyouBeta.ToString & " * " & stUserData.iTempTemp.ToString & "^ 2)) * (1 + " & dDaihyouAlpha.ToString & " * " & dStageTempConv.ToString & " + " & dDaihyouBeta.ToString & "*" & dStageTempConv.ToString & " ^ 2)")
#End If


            ElseIf IsTrimType2() Or IsTrimType3() Then  'V1.0.4.3④IsTrimType3()追加

                'V1.2.0.0②↓
                Dim ResCnt As Short
                'V2.0.0.0⑩                If UserSub.IsTrimType3() Then
                'V2.0.0.0⑩                ResCnt = 1                  ' チップ抵抗モードは１番目だけ使用する。
                'V2.0.0.0⑩            Else
                ResCnt = rn
                'V2.0.0.0⑩            End If
                'V1.2.0.0②↑

                ' 高精度薄膜抵抗トリミングの場合　：　TRV　＝　目標抵抗値 × 補正値
                'V2.0.0.0⑫                dTRV = stREG(rn).dblNOM * stUserData.dNomCalcCoff(ResCnt)
                dTRV = stREG(rn).dblNOM * (stUserData.dNomCalcCoff(UserSub.GetResNumberInCircuit(ResCnt)) / 1000000.0 + 1.0)                   'V2.0.0.0⑫ 補正値の項目をppm入力に変更 'V2.0.0.0⑩サーキット対応
                DebugLogOut("TRV:" & dTRV.ToString & " = " & stREG(rn).dblNOM.ToString & "* (" & stUserData.dNomCalcCoff(UserSub.GetResNumberInCircuit(ResCnt)).ToString & ") / 1000000.0 + 1.0)")
                'V1.2.0.0②                dTRV = stREG(rn).dblNOM * stUserData.dNomCalcCoff(rn)
                'V1.2.0.0②                DebugLogOut("TRV:" & dTRV.ToString & " = " & stREG(rn).dblNOM.ToString & " * " & stUserData.dNomCalcCoff(rn))

            Else
                Call Z_PRINT("UserSub.CalcTargeResistancetValue() ERROR 標準トリミングで呼ばれました = " & vbCrLf)
            End If

        Catch ex As Exception
            Call Z_PRINT("UserSub.CalcTargeResistancetValue() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub

    '===============================================================================
    '【機　能】 各カットの目標値算出
    '【引　数】 抵抗番号,カット番号
    '【戻り値】 目標抵抗値
    '===============================================================================
    Public Function GetTargeResistancetValue(ByVal rn As Integer, ByVal cn As Integer) As Double

        Try
            If IsTrimType1() Or IsTrimType4() Then

                ' 目標値（TRM)　＝　TRV　－　（カット毎のオフセット値　×　初期測定値　／　目標値算出係数　）
                GetTargeResistancetValue = dTRV - (stREG(rn).STCUT(cn).dblCOF * dInitialResValue / stUserData.dTargetCoff(UserBas.GetResistorNo(rn)))
                DebugLogOut("抵抗[" & rn.ToString & "]カット[" & cn.ToString & "]目標値:" & GetTargeResistancetValue.ToString & " = " & dTRV.ToString & " - (" & stREG(rn).STCUT(cn).dblCOF.ToString & " * " & dInitialResValue.ToString & " / " & stUserData.dTargetCoff(UserBas.GetResistorNo(rn)).ToString & ")")

            ElseIf IsTrimType2() Or IsTrimType3() Then  'V1.0.4.3④IsTrimType3()追加

                'V1.2.0.0②↓
                Dim ResCnt As Short
                'V2.0.0.0⑩                If UserSub.IsTrimType3() Then
                'V2.0.0.0⑩                    ResCnt = 1                  ' チップ抵抗モードは１番目だけ使用する。
                'V2.0.0.0⑩                Else
                ResCnt = rn
                'V2.0.0.0⑩                End If
                'V1.2.0.0②↑

                ' 目標値（TRM)　＝　TRV　－　（カット毎のオフセット値　×　初期測定値　／　目標値算出係数　）
                'V2.0.0.0⑩                GetTargeResistancetValue = dTRV - (stREG(rn).STCUT(cn).dblCOF * dInitialResValue / stUserData.dTargetCoff(ResCnt))
                GetTargeResistancetValue = dTRV - (stREG(rn).STCUT(cn).dblCOF * dInitialResValue / stUserData.dTargetCoff(UserSub.GetResNumberInCircuit(ResCnt)))       'V2.0.0.0⑩ UserSub.GetResNumberInCircuit(ResCnt)追加
                DebugLogOut("抵抗[" & rn.ToString & "]カット[" & cn.ToString & "]目標値:" & GetTargeResistancetValue.ToString & " = " & dTRV.ToString & " - (" & stREG(rn).STCUT(cn).dblCOF.ToString & " * " & dInitialResValue.ToString & " / " & stUserData.dTargetCoff(UserSub.GetResNumberInCircuit(ResCnt)).ToString & ")")
                'V1.2.0.0②                GetTargeResistancetValue = dTRV - (stREG(rn).STCUT(cn).dblCOF * dInitialResValue / stUserData.dTargetCoff(rn))
                'V1.2.0.0②                DebugLogOut("抵抗[" & rn.ToString & "]カット[" & cn.ToString & "]目標値:" & GetTargeResistancetValue.ToString & " = " & dTRV.ToString & " - (" & stREG(rn).STCUT(cn).dblCOF.ToString & " * " & dInitialResValue.ToString & " / " & stUserData.dTargetCoff(rn).ToString & ")")

            Else
                Call Z_PRINT("UserSub.GetTargeResistancetValue() ERROR 標準トリミングで呼ばれました = " & vbCrLf)
            End If

        Catch ex As Exception
            Call Z_PRINT("UserSub.GetTargeResistancetValue() TRAP ERROR = " & ex.Message & vbCrLf)
            GetTargeResistancetValue = -9999.999
        End Try
    End Function

    '===============================================================================
    '【機　能】 抵抗測定、高速、高精度測定の変更
    '【引　数】 抵抗番号,カット番号
    '【戻り値】 無し
    '===============================================================================
    Public Sub ChangeMeasureSpeed(ByVal rn As Integer, ByVal cn As Integer, ByVal idx As Short)

        Try
            intTMM_Save = stREG(rn).STCUT(cn).intIXTMM(idx)         ' 測定モード(0:高速　1:高精度)
            intMType_Save = stREG(rn).STCUT(cn).intIXMType(idx)     ' 測定機器0～5(0:内部測定　1～:外部機器)

            If stUserData.iTrimSpeed = 1 Then                       ' 高速
                stREG(rn).STCUT(cn).intIXTMM(idx) = 0               ' モード(0:高速(コンパレータ非積分モード), 1:高精度(積分モード))　※インデックス時使用する。
                stREG(rn).STCUT(cn).intIXMType(idx) = 0             ' 測定種別(0=内部測定, 1=外部測定)
            ElseIf stUserData.iTrimSpeed = 2 Then                   ' 高精度
                'V1.2.0.0②↓
                Dim ResCnt As Short
                'V2.0.0.0⑩                If UserSub.IsTrimType3() Then
                'V2.0.0.0⑩                    ResCnt = 1                  ' チップ抵抗モードは１番目だけ使用する。
                'V2.0.0.0⑩                Else
                ResCnt = rn
                'V2.0.0.0⑩                End If
                If cn < stUserData.iChangeSpeed(GetResistorNo(ResCnt)) Then
                    'V1.2.0.0②↑
                    'V1.2.0.0②                    If cn < stUserData.iChangeSpeed(GetResistorNo(rn)) Then
                    stREG(rn).STCUT(cn).intIXTMM(idx) = 0           ' モード(0:高速(コンパレータ非積分モード), 1:高精度(積分モード))　※インデックス時使用する。
                    stREG(rn).STCUT(cn).intIXMType(idx) = 0         ' 測定種別(0=内部測定, 1=外部測定)
                End If
            Else
                Exit Sub
            End If
        Catch ex As Exception
            Call Z_PRINT("UserSub.ChangeMeasureSpeed() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Sub
    '===============================================================================
    '【機　能】 抵抗測定、高速、高精度測定モードの復元
    '【引　数】 抵抗番号,カット番号
    '【戻り値】 無し
    '===============================================================================
    Public Sub ResoreMeasureSpeed(ByVal rn As Integer, ByVal cn As Integer, ByVal idx As Short)
        Try
            stREG(rn).STCUT(cn).intIXTMM(idx) = intTMM_Save        ' モードの保存
            stREG(rn).STCUT(cn).intIXMType(idx) = intMType_Save    ' 測定種別の保存
        Catch ex As Exception
            Call Z_PRINT("UserSub.ResoreMeasureSpeed() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub

    '===============================================================================
    '【機　能】 温度センサータイプかを判断する
    '【引　数】 無し
    '【戻り値】 True = 一致, False = 不一致
    '===============================================================================
    Public Function IsTrimType1() As Boolean
        If stUserData.iTrimType = 1 Then
            IsTrimType1 = True
        Else
            IsTrimType1 = False
        End If
    End Function
    '===============================================================================
    '【機　能】 抵抗トリミングタイプかを判断する
    '【引　数】 無し
    '【戻り値】 True = 一致, False = 不一致
    '===============================================================================
    Public Function IsTrimType2() As Boolean
        If stUserData.iTrimType = 2 Then
            IsTrimType2 = True
        Else
            IsTrimType2 = False
        End If
    End Function
    'V1.0.4.3④ ADD START
    '===============================================================================
    '【機　能】 チップ抵抗トリミングタイプかを判断する
    '【引　数】 無し
    '【戻り値】 True = 一致, False = 不一致
    '===============================================================================
    Public Function IsTrimType3() As Boolean
        If stUserData.iTrimType = 3 Then
            IsTrimType3 = True
        Else
            IsTrimType3 = False
        End If
    End Function
    'V1.0.4.3④ ADD END
    'V2.0.0.0① ADD START
    '===============================================================================
    '【機　能】 チップ温度センサータイプかを判断する
    '【引　数】 無し
    '【戻り値】 True = 一致, False = 不一致
    '===============================================================================
    Public Function IsTrimType4() As Boolean
        If stUserData.iTrimType = 4 Then
            IsTrimType4 = True
        Else
            IsTrimType4 = False
        End If
    End Function

    'V2.2.1.7① ↓
    '===============================================================================
    '【機　能】 マーク印字タイプかを判断する
    '【引　数】 無し
    '【戻り値】 True = 一致, False = 不一致
    '===============================================================================
    Public Function IsTrimType5() As Boolean
        If stUserData.iTrimType = 5 Then
            IsTrimType5 = True
        Else
            IsTrimType5 = False
        End If
    End Function
    'V2.2.1.7① ↑

    ' ''' <summary>
    ' ''' チップタイプかを判断する
    ' ''' </summary>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Function IsCircuitTrimType() As Boolean
    '    '20171207佐々木様と相談してチップ抵抗トリミングタイプのみに適用　If stUserData.iTrimType = 3 Or stUserData.iTrimType = 4 Then
    '    If stUserData.iTrimType = 3 Then
    '        IsCircuitTrimType = True
    '    Else
    '        IsCircuitTrimType = False
    '    End If
    'End Function
    'V2.0.0.0① ADD END
    '===============================================================================
    '【機　能】 特殊処理のトリミングタイプかを判断する
    '【引　数】 無し
    '【戻り値】 True = 一致, False = 不一致
    '===============================================================================
    Public Function IsSpecialTrimType() As Boolean
        If stUserData.iTrimType <> 0 Then
            IsSpecialTrimType = True
        Else
            IsSpecialTrimType = False
        End If
    End Function
    '===============================================================================
    '【機　能】 初期測定値の保存
    '【引　数】 初期測定値
    '【戻り値】 無し
    '===============================================================================
    Public Sub SetInitialResValue(ByVal dVal As Double)
        dInitialResValue = dVal
    End Sub

    'V2.1.0.0⑤↓
    ''' <summary>
    ''' 初期測定値の取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetInitialResValue() As Double
        Return (dInitialResValue)
    End Function
    'V2.1.0.0⑤↑
    '===============================================================================
    '【機　能】 標準抵抗測定値の保存
    '【引　数】 標準抵抗測定値
    '【戻り値】 無し
    '===============================================================================
    Public Sub SetStandardResValue(ByVal dVal As Double)
        dStandardResValue = dVal
    End Sub

    '===============================================================================
    '【機　能】 ファイナル測定後の処理
    '【引　数】 抵抗番号,ファイナルテスト測定値
    '【戻り値】 無し
    '===============================================================================
    Public Sub DevCalculation(ByVal rn As Integer, ByVal dFtVal As Double)
        Dim iResNo As Integer
        Try


            iResNo = GetResistorNo(rn)      ' トリミングデータ上の抵抗番号からカットする抵抗番号を求める。（測定のみを除外する。）

            'V1.2.0.0②↓
            'V2.0.0.0⑩            If UserSub.IsTrimType3() Then
            'V2.0.0.0⑩                iResNo = 1
            'V2.0.0.0⑩            End If
            'V1.2.0.0②↑

            If iResNo > MAX_RES_USER Then
                Return
            End If

            stUserData.dFtVal(iResNo) = dFtVal

            '14710      DEV1#=FIX((R.FT1#-NRV1#)/NRV1#*1000000#)
            '14832      DEV2#=FIX((R.FT2#-NRV2#)/NRV2#*1000000#)

            If stREG(rn).dblNOM = 0.0 Then
                Call Z_PRINT("UserSub.DevCalculation() 目標値が０です。計算が出来ません" & vbCrLf)
                Exit Sub
            End If
            ' トリミング誤差　＝　（　トリミング値　－　スタンダード実測値に対してのトリミング目標値　）／スタンダード実測値に対してのトリミング目標値　* 10^6
            'V2.0.0.0⑱            stUserData.dDev(iResNo) = FNDEVP(stUserData.dFtVal(iResNo), stREG(rn).dblNOM)
            stUserData.dDev(iResNo) = FNDEVP(stUserData.dFtVal(iResNo), UserSub.GetTRV())

        Catch ex As Exception
            Call Z_PRINT("UserSub.DevCalculation() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub

    '''===============================================================================
    ''' <summary>
    ''' ロット交換の可否確認
    ''' </summary>
    ''' <param name="HostMode">ローダモード(cHOSTcMODEcMANUAL:手動 cHOSTcMODEcAUTO:自動)</param>
    ''' <param name="Start">cHSTcTRMCMD：トリミングスタート　cHSTcLOTCHANGE：ロット交換スタート</param>
    ''' <returns>0：ロット交換無し 1:枚数到達のロット交換 2:切り替え信号のロット交換</returns>
    ''' <remarks></remarks>
    '''===============================================================================
'V2.2.1.1⑧'    Public Function IsLotChange(ByVal HostMode As Integer, ByVal Start As Short, ByVal fStartTrim As Boolean) As Integer
    Public Function IsLotChange(ByVal HostMode As Integer, ByVal Start As Short, ByVal fStartTrim As Boolean, ByRef lotcnt As Integer) As Integer   'V2.2.1.1⑧


        Dim bPrint As Boolean = False
        Dim LdIDat As Integer

        IsLotChange = 0

        'If Start <> cHSTcTRMCMD And Start <> cHSTcLOTCHANGE Then
        '    Exit Function
        'End If


        Select Case (stUserData.iLotChange) ' ロット終了条件 0:終了条件判定無し 1:枚数 2:ローダー信号 3:両方
            Case 0
                IsLotChange = 0
            Case 1
                If stCounter.PlateCounter >= stUserData.lLotEndSL Then      ' 処理基板数に到達
                    If fStartTrim Then
                        IsLotChange = 1
                    End If
                    If Not UserBas.stCounter.LotPrint Then                  'V1.2.0.3
                        bPrint = True
                    End If                                                  'V1.2.0.3
                End If
            Case 2
                'V1.2.0.0④↓
                If giLoaderType = 1 Then
                    ObjSys.Z_ATLDGET(LdIDat)                                        ' ローダー入力
                    If LdIDat And clsLoaderIf.LINP_LOT_CHG Then
                        LdIDat = cHSTcLOTCHANGE
                    End If
                    Start = LdIDat And cHSTcLOTCHANGE
                End If

                If fStartTrim And HostMode = cHOSTcMODEcAUTO And Start = cHSTcLOTCHANGE Then
                    IsLotChange = 2
                    bPrint = True
                End If
            Case 3
                'V1.2.0.0④↓
                If giLoaderType = 1 Then
                    ObjSys.Z_ATLDGET(LdIDat)                                        ' ローダー入力
                    If LdIDat And clsLoaderIf.LINP_LOT_CHG Then

                        LdIDat = cHSTcLOTCHANGE
                    End If
                    Start = LdIDat And cHSTcLOTCHANGE
                End If

                If fStartTrim And HostMode = cHOSTcMODEcAUTO And Start = cHSTcLOTCHANGE Then
                    IsLotChange = 2
                    bPrint = True
                Else
                    If stCounter.PlateCounter >= stUserData.lLotEndSL Then      ' 処理基板数に到達
                        If fStartTrim Then
                            IsLotChange = 1
                        End If
                        If Not UserBas.stCounter.LotPrint Then                  'V1.2.0.3
                            bPrint = True
                        End If                                                  'V1.2.0.3
                    End If
                End If
        End Select

        ''V2.2.1.1⑧ ↓
        'フラグ以外の条件でロット切り替えを行う場合は、その分ロット切り替え回数を加算する 
        If IsLotChange = 1 Then
            If lotcnt > 0 Then  'ロット切り替えフラグがONしていた場合実行する 
                lotcnt = lotcnt + 1
            End If
        End If

        'ロット切り替えフラグがONしていた場合実行する 
        If lotcnt > 0 Then
            IsLotChange = 1

            Call Z_PRINT("ロット切り替えフラグにより、ロット切り替えを行いました。" & lotcnt.ToString)

            If Not UserBas.stCounter.LotPrint Then
                bPrint = True
            End If
        End If
        ''V2.2.1.1⑧ ↑

        'V1.2.0.0⑥        If bPrint And Not UserBas.stCounter.LotPrint Then
        If bPrint Then
            Call UserSub.LotEnd()                           ' ロット終了時のデータ出力
            Call Printer.Print(False)                       ' ロット情報印刷
            UserBas.stCounter.LotPrint = True               ' ロット終了時の印刷実行済みでTrue
        End If

    End Function

    '===============================================================================
    '【機　能】 印刷ヘッダ情報ファイルの作成、既存ファイルの削除も行う
    '【引　数】 無し
    '【戻り値】 無し
    '===============================================================================
    Public Sub MakePrintFileHeader()
        Dim WS As IO.StreamWriter
        Dim sData As String


        Try
            UserBas.stCounter.LotPrint = False              'V1.2.0.3 念の為追加
            ' 印刷データを削除する。
            If IO.File.Exists(cTRIM_PRINT_DATA_HEAD) = True Then
                IO.File.Delete(cTRIM_PRINT_DATA_HEAD)
            End If

            If IO.File.Exists(cTRIM_PRINT_DATA_RES) = True Then
                IO.File.Delete(cTRIM_PRINT_DATA_RES)
            End If
            If IO.File.Exists(cTRIM_PRINT_DATA_PLATE) = True Then
                IO.File.Delete(cTRIM_PRINT_DATA_PLATE)
            End If
            If IO.File.Exists(cTRIM_PRINT_DATA_END) = True Then
                IO.File.Delete(cTRIM_PRINT_DATA_END)
            End If

            ' ヘッダー情報を出力する。
            WS = New IO.StreamWriter(cTRIM_PRINT_DATA_HEAD, True, System.Text.Encoding.GetEncoding("Shift-JIS"))
            WS.WriteLine("───────────────────────────────────────────────")
            WS.WriteLine("日付  " & DateTime.Now.ToString("yyyy/MM/dd"))
            WS.WriteLine("ロットＮｏ．      ＝ " & stUserData.sLotNumber.PadRight(20) & "オペレータ名     ＝ " & stUserData.sOperator)
            WS.WriteLine("パターンＮｏ．    ＝ " & stUserData.sPatternNo.PadRight(20) & "プログラムＮｏ． ＝ " & stUserData.sProgramNo)

            If UserSub.IsTrimType1() Or UserSub.IsTrimType4() Then       ' 温度センサー'V2.0.0.0①sTrimType4()追加
                WS.WriteLine("Ｒ１:設定抵抗値             ＝ " & stREG(UserBas.GetCutResistorNo(1)).dblNOM.ToString(TARGET_DIGIT_DEFINE).PadLeft(15) & " [ohm]")
#If OLD_CALURATION Then 'V2.0.0.0⑪
                WS.WriteLine("スタンダード抵抗値（０℃）  ＝ " & stUserData.dStandardRes0.ToString("0.00000").PadLeft(15) & " [ohm]")
                WS.WriteLine("スタンダード抵抗値（２５℃）＝ " & stUserData.dStandardRes25.ToString("0.00000").PadLeft(15) & " [ohm]")
                If stUserData.iTempTemp = 1 Then    ' 参照温度	１：０℃ または ２：２５℃
                    WS.WriteLine("参照温度　　　　　　　　　　＝０℃")
                ElseIf stUserData.iTempTemp = 2 Then
                    WS.WriteLine("参照温度　　　　　　　　　　＝２５℃")
                End If
#Else
                'V2.0.0.4①                Dim dStdResValue As Double = stUserData.dTemperatura0 * (1.0 + stUserData.dAlpha * stUserData.iTempTemp + stUserData.dBeta * stUserData.iTempTemp ^ 2)
                'V2.0.0.4①↓
                Dim dAlpha As Double = stUserData.dAlpha / 10.0 ^ 6
                Dim dBeta As Double = stUserData.dBeta / 10.0 ^ 6
                Dim dStdResValue As Double = stUserData.dTemperatura0 * (1.0 + dAlpha * stUserData.iTempTemp + dBeta * stUserData.iTempTemp ^ 2)
                'V2.0.0.4①↑
                WS.WriteLine("STD抵抗値（" & stUserData.iTempTemp.ToString("0") & "℃）  ＝ " & dStdResValue.ToString(TARGET_DIGIT_DEFINE).PadLeft(15) & " [ohm]")
#End If
                WS.WriteLine("ファイナルテストリミット[ppm] High ＝ " & stUserData.dFinalLimitHigh.ToString("0.0").PadLeft(10) & "  Low　＝ " & stUserData.dFinalLimitLow.ToString("0.0").PadLeft(10))
            Else                        ' 抵抗トリミング
                If stUserData.iTrimSpeed = 1 Then
                    sData = "トリミングモード     ＝ 高速度モード"
                ElseIf stUserData.iTrimSpeed = 2 Then
                    sData = "トリミングモード     ＝ 高精度モード"
                Else
                    sData = "トリミングモード     ＝ 設定値"
                End If
                WS.WriteLine(sData)

                Dim Rcnt As Integer = UserBas.GetRCountExceptMeasure()
                'V2.2.0.0⑯↓
                If stMultiBlock.gMultiBlock <> 0 Then

                    ' マルチブロックで設定されている分実行する 
                    For blk As Integer = 0 To stMultiBlock.BLOCK_DATA.Length - 2
                        ' @'V2.2.0.033 If stMultiBlock.BLOCK_DATA(0).gBlockCnt <> 0 Then
                        If stMultiBlock.BLOCK_DATA(blk).gBlockCnt <> 0 Then     'V2.2.0.033
                            WS.WriteLine("MBNo： " & stMultiBlock.BLOCK_DATA(blk).DataNo.ToString)
                            'V2.2.0.033   WS.WriteLine("マルチブロックNo： " & stMultiBlock.BLOCK_DATA(blk).DataNo.ToString)

                            For rn As Integer = 1 To Rcnt
                                WS.WriteLine("R" & rn.ToString & ":設定抵抗値     ＝ " & stMultiBlock.BLOCK_DATA(blk).dblNominal(UserBas.GetCutResistorNo(rn) - 1).ToString(TARGET_DIGIT_DEFINE).PadLeft(15) & " [ohm]")
                            Next rn
                            For rn As Integer = 1 To Rcnt
                                WS.WriteLine("R" & rn.ToString & ":補正値         ＝ " & stMultiBlock.BLOCK_DATA(blk).dblCorr(rn - 1).ToString("0.000").PadLeft(15) & " [ppm]") 'V2.0.0.0⑫補正値の項目をppm入力に変更
                            Next rn
                        End If
                    Next blk
                Else
                    For rn As Integer = 1 To Rcnt
                        WS.WriteLine("R" & rn.ToString & ":設定抵抗値     ＝ " & stREG(UserBas.GetCutResistorNo(rn)).dblNOM.ToString(TARGET_DIGIT_DEFINE).PadLeft(15) & " [ohm]")
                        'V2.0.0.0⑩                    If IsTrimType3() Then           'V1.2.0.0②
                        'V2.0.0.0⑩                        Exit For                    'V1.2.0.0②
                        'V2.0.0.0⑩                    End If                          'V1.2.0.0②
                    Next
                    For rn As Integer = 1 To Rcnt
                        WS.WriteLine("R" & rn.ToString & ":補正値         ＝ " & stUserData.dNomCalcCoff(rn).ToString("0.000").PadLeft(15) & " [ppm]") 'V2.0.0.0⑫補正値の項目をppm入力に変更
                        'V2.0.0.0⑩                    If IsTrimType3() Then           'V1.2.0.0②
                        'V2.0.0.0⑩                        Exit For                    'V1.2.0.0②
                        'V2.0.0.0⑩                    End If                          'V1.2.0.0②
                    Next
                End If
                'V2.2.0.0⑯↑
                WS.WriteLine("ファイナルテストリミット[ppm] High ＝ " & stUserData.dFinalLimitHigh.ToString("0.0").PadLeft(10) & "  Low　＝ " & stUserData.dFinalLimitLow.ToString("0.0").PadLeft(10))
                'V2.0.0.0⑩                If IsTrimType2() Then               'V1.2.0.0②
                WS.WriteLine("相対値リミット[ppm]              ＝ " & stUserData.dRelativeHigh.ToString("0.000").PadLeft(10))
                'V2.0.0.0⑩            End If                              'V1.2.0.0②

            End If

            WS.WriteLine("───────────────────────────────────────────────")

            WS.Close()

        Catch ex As Exception
            Call Z_PRINT("UserSub.MakePrintFileHeader() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Sub
    '===============================================================================
    '【機　能】 素子単位のファイル又はログの出力
    '【引　数】 IO.StreamWriter、文字データ
    '【戻り値】 無し
    '===============================================================================
    Private Sub ResistorDataOutPut(ByVal WS As IO.StreamWriter, ByVal bPrint As Boolean, ByVal sMessage As String)
        Dim printcnt As Integer = 0
        Dim blkcnt As Integer = 0

        'V2.2.0.033↓
        If stMultiBlock.gMultiBlock <> 0 Then
            For cnt As Integer = 0 To 4
                If (stMultiBlock.BLOCK_DATA(cnt).gBlockCnt) <> 0 Then
                    blkcnt = blkcnt + 1
                End If
            Next

            printcnt = stUserData.lPrintRes \ blkcnt


            If stMultiBlock.BLOCK_DATA(stExecBlkData.DataNo).gProcCnt < printcnt And bPrint Then ' ロット終了時印刷素子数に達していない場合は、ファイルへの出力    'V2.2.0.033
                WS.WriteLine(sMessage)
            End If

        Else
            printcnt = stUserData.lPrintRes
            If lResCounterForPrinter < printcnt And bPrint Then ' ロット終了時印刷素子数に達していない場合は、ファイルへの出力    'V2.2.0.033
                WS.WriteLine(sMessage)
            End If

        End If


        'V2.2.0.033        If lResCounterForPrinter < stUserData.lPrintRes And bPrint Then ' ロット終了時印刷素子数に達していない場合は、ファイルへの出力
        '        If lResCounterForPrinter < printcnt And bPrint Then ' ロット終了時印刷素子数に達していない場合は、ファイルへの出力    'V2.2.0.033

        Call Z_PRINT(sMessage.Replace(vbTab, " ") & vbCrLf)         ' ログ出力エリアへの出力

    End Sub
    '''=============================================================================
    ''' <summary>
    ''' 素子単位の判定NG化
    ''' </summary>
    ''' <remarks></remarks>
    '''=============================================================================
    Public Sub NgJudgeSet()
        bOkJudge = False
    End Sub

    Public Sub SkipSet()
        bSkip = True
    End Sub

    ''' <summary>
    ''' サーキットスキップは、True
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SkipGet() As Boolean
        Return (bSkip)
    End Function
    '''=============================================================================
    ''' <summary>
    ''' 素子単位の判定初期化
    ''' </summary>
    ''' <remarks></remarks>
    '''=============================================================================
    Public Sub NgJudgeReset()
        For i As Integer = 1 To MAX_RES_USER
            sResistorPrintData(i) = ""
        Next
        bOkJudge = True
        bSkip = False
    End Sub

    '===============================================================================
    '【機　能】 全抵抗終了時の判定
    '【引　数】 抵抗番号
    '【戻り値】 無し
    '===============================================================================
    Public Function FinalJudge(ByVal rn As Integer) As Boolean
        Dim WS As IO.StreamWriter
        Dim dDev As Double
        Dim sJudge As String
        Dim iResCnt As Integer
        Dim bHeaderPrint As Boolean
        Dim iCnt As Integer

        ' 14840       DEV#=DEV1#-DEV2#
        ' 14841       IF ECF%<>1 THEN FTOV3#=FTOV3#+1#: GOTO *TRIM.NG
        ' 14842       CALL TEST%(DEV#,SRV#,Z2,STLO#,STHI#)
        ' 14843       IF ECF%=2 THEN FTLO3#=FTLO3#+1#:  GOTO *TRIM.NG
        ' 14844       IF ECF%=3 THEN FTHI3#=FTHI3#+1#:  GOTO *TRIM.NG

        'V2.0.0.0⑤ 全ての"0.00000"をTARGET_DIGIT_DEFINEへ変更
        'V2.0.0.0⑤　PadLeft(13)をPadLeft(15)へ変更

        Try
            If IO.File.Exists(cTRIM_PRINT_DATA_RES) Then
                bHeaderPrint = False
            Else
                bHeaderPrint = True
            End If

            FinalJudge = True


            iResCnt = GetRCountExceptMeasure()
            If iResCnt > MAX_RES_USER Then
                iResCnt = MAX_RES_USER
            End If

            'V1.2.0.0②↓
            'V2.0.0.0⑩            If UserSub.IsTrimType3() Then
            'V2.0.0.0⑩                iResCnt = 1
            'V2.0.0.0⑩            End If
            'V1.2.0.0②↑

            WS = New IO.StreamWriter(cTRIM_PRINT_DATA_RES, True, System.Text.Encoding.GetEncoding("Shift-JIS"))     ' 抵抗データ印刷データ
            If bHeaderPrint Then
                If UserSub.IsTrimType1 Or UserSub.IsTrimType4() Then    'V2.0.0.0①sTrimType4()追加
                    ResistorDataOutPut(WS, True, "No" & vbTab & "X" & vbTab & "Y" & vbTab & "抵抗名" & vbTab & "目標抵抗値 イニシャル測定値 ファイナル測定値       誤差    判定")
                Else
                    'V2.0.0.0②↓
                    If DGL = TRIM_VARIATION_MEAS Then ' 測定値変動測定
                        ResistorDataOutPut(WS, True, "No" & vbTab & "X" & vbTab & "Y" & vbTab & "抵抗名" & vbTab & "トリミング後ＦＴ値 ファイナル測定値   誤差   判定")
                    Else
                        'V2.0.0.0②↑
                        'V2.2.0.033↓
                        If stMultiBlock.gMultiBlock <> 0 Then
                            ResistorDataOutPut(WS, True, "No" & vbTab & "X" & vbTab & "Y" & vbTab & "抵抗名       " & vbTab & "イニシャル測定値   ファイナル測定値   誤差   判定")
                        Else
                            ResistorDataOutPut(WS, True, "No" & vbTab & "X" & vbTab & "Y" & vbTab & "抵抗名" & vbTab & "イニシャル測定値   ファイナル測定値   誤差   判定")
                        End If

                    End If                      'V2.0.0.0②
                End If
            End If
            If (stREG(rn).intMode = 0) Then                         ' 判定モード = 0(比率(ppm)) ?
                dDev = FNDEVP(dblVX(2), dblNM(2))                ' 誤差 = (測定値 / 目標値 - 1) * 100
            Else
                dDev = dblVX(2) - dblNM(2)                       ' 誤差1(絶対値) = 測定値 - 目標値
            End If

            'V2.2.0.033↓
            Dim addStr As String = ""
            If stMultiBlock.gMultiBlock <> 0 Then
                addStr = " MBNo:" & stExecBlkData.DataNo.ToString()
            Else
                addStr = ""
            End If
            'V2.2.0.033↑

            If UserSub.IsTrimType1 Or iResCnt = 1 Then                                     ' １素子１抵抗の時　２抵抗以上は、最後に出力

                If Not stREG(rn).bPattern Then                  'V1.2.0.0③ カット位置補正の判定 True：OK False:NG
                    'V2.2.0.033                    sResistorPrintData(1) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & "   " & vbTab & "カット位置補正 自動ＮＧ判定 = ＮＧ"
                    sResistorPrintData(1) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & addStr & "   " & vbTab & "カット位置補正 自動ＮＧ判定 = ＮＧ"       'V2.2.0.033 
                ElseIf UserSub.IsTrimType1 Or UserSub.IsTrimType4() Then
                    'V2.2.0.033  sResistorPrintData(1) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & vbTab & dblNM(2).ToString(TARGET_DIGIT_DEFINE).PadLeft(15) & " " & UserSub.ChangeOverFlow(dblVX(1).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dblVX(2).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dDev.ToString("0.0")).PadLeft(15) & " " & strJUG(rn)
                    sResistorPrintData(1) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & addStr & vbTab & dblNM(2).ToString(TARGET_DIGIT_DEFINE).PadLeft(15) & " " & UserSub.ChangeOverFlow(dblVX(1).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dblVX(2).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dDev.ToString("0.0")).PadLeft(15) & " " & strJUG(rn)    'V2.2.0.033 
                Else
                    'V2.0.0.0②↓
                    If DGL = TRIM_VARIATION_MEAS Then ' 測定値変動測定
                        'V2.2.0.033  sResistorPrintData(1) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & vbTab & UserSub.ChangeOverFlow(dMeasVariationNOM(rn).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dblVX(2).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dDev.ToString("0.0")).PadLeft(15) & "  " & strJUG(rn)
                        sResistorPrintData(1) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & addStr & vbTab & UserSub.ChangeOverFlow(dMeasVariationNOM(rn).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dblVX(2).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dDev.ToString("0.0")).PadLeft(15) & "  " & strJUG(rn)
                    Else
                        'V2.0.0.0②↑
                        'V2.2.0.033 sResistorPrintData(1) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & vbTab & UserSub.ChangeOverFlow(dblVX(1).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dblVX(2).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dDev.ToString("0.0")).PadLeft(15) & "  " & strJUG(rn)
                        sResistorPrintData(1) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & addStr & vbTab & UserSub.ChangeOverFlow(dblVX(1).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dblVX(2).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dDev.ToString("0.0")).PadLeft(15) & "  " & strJUG(rn)
                    End If                      'V2.0.0.0②
                End If
                If strJUG(rn) = JG_OK Then
                    ResistorDataOutPut(WS, True, sResistorPrintData(1))
                    lResCounterForPrinter = lResCounterForPrinter + 1                       ' ＯＫのみ印刷してカウントアップする。
                    If UserSub.IsTrimType2 Then
                        stCounter.OK_Counter = stCounter.OK_Counter + 1
                        stCounter.Total_OK_Counter = stCounter.Total_OK_Counter + 1

                        'V2.2.0.0⑯↓
                        If stMultiBlock.gMultiBlock <> 0 Then
                            stMultiBlock.BLOCK_DATA(stExecBlkData.DataNo).gProcCnt = stMultiBlock.BLOCK_DATA(stExecBlkData.DataNo).gProcCnt + 1
                            gObjFrmDistribute.SetOkCounterMulti()
                        End If
                        'V2.2.0.0⑯↑

                    End If
                Else
                    ResistorDataOutPut(WS, False, sResistorPrintData(1))
                End If
            Else
                iCnt = GetResistorNo(rn)
                If iCnt <= MAX_RES_USER Then
                    'V2.0.0.0②↓
                    If DGL = TRIM_VARIATION_MEAS Then ' 測定値変動測定
                        'V2.2.0.033 sResistorPrintData(iCnt) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & vbTab & UserSub.ChangeOverFlow(dMeasVariationNOM(rn).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dblVX(2).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dDev.ToString("0.0")).PadLeft(15) & "  " & strJUG(rn)
                        sResistorPrintData(iCnt) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & addStr & vbTab & UserSub.ChangeOverFlow(dMeasVariationNOM(rn).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dblVX(2).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dDev.ToString("0.0")).PadLeft(15) & "  " & strJUG(rn)
                    Else
                        'V2.0.0.0②↑
                        'V2.2.0.033 sResistorPrintData(iCnt) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & vbTab & UserSub.ChangeOverFlow(dblVX(1).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dblVX(2).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dDev.ToString("0.0")).PadLeft(15) & "  " & strJUG(rn)
                        sResistorPrintData(iCnt) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & addStr & vbTab & UserSub.ChangeOverFlow(dblVX(1).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dblVX(2).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dDev.ToString("0.0")).PadLeft(15) & "  " & strJUG(rn)
                    End If                      'V2.0.0.0②
                    'V1.2.0.0③↓
                    If Not stREG(rn).bPattern Then                  'カット位置補正の判定 True：OK False:NG
                        'V2.2.0.033 sResistorPrintData(iCnt) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & "   " & vbTab & "カット位置補正 自動ＮＧ判定 = ＮＧ"
                        sResistorPrintData(iCnt) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & addStr & "   " & vbTab & "カット位置補正 自動ＮＧ判定 = ＮＧ"
                    End If
                    'V1.2.0.0③↑
                End If
            End If

            'V2.0.0.0⑩            If UserSub.IsTrimType2 Then
            If UserSub.IsTrimType2 Or UserSub.IsTrimType3() Then        'V2.0.0.0⑩チップ抵抗モード追加
                If iResCnt > 1 And GetResistorNo(rn) = iResCnt Then     ' 抵抗数１以上の時

                    Dim sMessage As String = ""
                    Dim largest As Double = Double.MinValue
                    Dim smallest As Double = Double.MaxValue
                    Dim i As Integer

                    If bSkip Then
                        dDev = -1000000.0
                    Else
                        For i = 1 To iResCnt Step 1
                            largest = Math.Max(largest, stUserData.dDev(i))
                            smallest = Math.Min(smallest, stUserData.dDev(i))
                        Next

                        ' 相対値　＝　最大値　－　最小値

                        dDev = Math.Abs(largest - smallest)
                        DebugLogOut("DEV =" & dDev.ToString(TARGET_DIGIT_DEFINE) & " L= " & largest.ToString(TARGET_DIGIT_DEFINE) & " M= " & smallest.ToString(TARGET_DIGIT_DEFINE))
                    End If


                    If dDev <= stUserData.dRelativeHigh And bOkJudge Then
                        stCounter.OK_Counter = stCounter.OK_Counter + 1
                        stCounter.Total_OK_Counter = stCounter.Total_OK_Counter + 1

                        'V2.2.0.0⑯↓
                        If stMultiBlock.gMultiBlock <> 0 Then
                            gObjFrmDistribute.SetOkCounterMulti()
                        End If
                        'V2.2.0.0⑯↑

                        sJudge = "OK"
                        For i = 1 To iResCnt Step 1
                            ResistorDataOutPut(WS, True, sResistorPrintData(i))
                        Next
                        sMessage = sMessage & "相対値[ppm] " & UserSub.ChangeOverFlow(dDev.ToString("0.0")) & vbTab & sJudge
                        ResistorDataOutPut(WS, True, sMessage)
                        lResCounterForPrinter = lResCounterForPrinter + 1                       ' ＯＫのみ印刷してカウントアップする。
                        'V2.2.0.0⑯↓
                        If stMultiBlock.gMultiBlock <> 0 Then
                            stMultiBlock.BLOCK_DATA(stExecBlkData.DataNo).gProcCnt = stMultiBlock.BLOCK_DATA(stExecBlkData.DataNo).gProcCnt + 1
                        End If
                    Else
                        FinalJudge = False
                        sJudge = "NG"
                        For i = 1 To iResCnt Step 1
                            ResistorDataOutPut(WS, False, sResistorPrintData(i))
                        Next
                        sMessage = sMessage & "相対値[ppm] " & UserSub.ChangeOverFlow(dDev.ToString("0.0")) & vbTab & sJudge
                        ResistorDataOutPut(WS, False, sMessage)
                        If Not bSkip Then
                            stCounter.FTHigh = stCounter.FTHigh + 1
                            stCounter.Total_FTHigh = stCounter.Total_FTHigh + 1

                            ' 'V2.2.0.0⑯↓
                            If stMultiBlock.gMultiBlock <> 0 Then
                                gObjFrmDistribute.SetFTHighCounterMulti()
                            End If
                            ' 'V2.2.0.0⑯↑

                        End If
                        strJUG(rn) = JG_FH
                    End If
                    NgJudgeReset()
                End If
            End If
            WS.Close()

            'V1.2.0.0②            If GetResistorNo(rn) = iResCnt Then
            'V2.0.0.0⑩            If GetResistorNo(rn) = iResCnt Or UserSub.IsTrimType3() Then    'V1.2.0.0② チップ抵抗モードは抵抗単位でカウントする。
            If GetResistorNo(rn) = iResCnt Then         'V2.0.0.0⑩
                stCounter.TrimCounter = stCounter.TrimCounter + 1 ' ﾄﾘﾐﾝｸﾞ数ｶｳﾝﾄｱｯﾌﾟ
                stCounter.Total_TrimCounter = stCounter.Total_TrimCounter + 1

                ' 'V2.2.0.0⑯↓
                If stMultiBlock.gMultiBlock <> 0 Then
                    gObjFrmDistribute.SetTrimCounterMulti()
                End If
                ' 'V2.2.0.0⑯↑

            End If

            Call Set_NG_Counter()                       'V1.2.0.0③ NGカウンターの更新

        Catch ex As Exception
            Call Z_PRINT("UserSub.FinalJudge() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Function

    '===============================================================================
    '【機　能】 基板終了時の印刷処理を行う
    '【引　数】 無し
    '【戻り値】 無し
    '===============================================================================
    Public Sub SubstrateEnd()
        Dim WS As IO.StreamWriter

        Try
            If (stCounter.PlateCounter = 0) Then
                Return
            End If
            '###1030③            UserBas.stCounter.EndTime = DateTime.Now()              ' 基板処理終了時間保存

            WS = New IO.StreamWriter(cTRIM_PRINT_DATA_PLATE, True, System.Text.Encoding.GetEncoding("Shift-JIS"))   ' 基板データ印刷
            If stCounter.PlateCounter = 1 Then
                WS.WriteLine("───────────────────────────────────────────────")
            End If

            If stMultiBlock.gMultiBlock <> 0 Then

                WS.WriteLine("No." & stCounter.PlateCounter.ToString & "  Start = " & stCounter.StartTime.ToString("HH:mm:ss") & " End = " & stCounter.EndTime.ToString("HH:mm:ss"))

                ' マルチブロックで設定されている分実行する 
                For blk As Integer = 1 To stMultiBlock.BLOCK_DATA.Length - 1

                    'V2.2.0.033 If stMultiBlock.BLOCK_DATA(0).gBlockCnt <> 0 Then
                    If stMultiBlock.BLOCK_DATA(blk - 1).gBlockCnt <> 0 Then       'V2.2.0.033

                        With stToTalDataMulti(blk)

                            'V2.2.0.033 WS.WriteLine("マルチブロックNo： " & stMultiBlock.BLOCK_DATA(blk - 1).DataNo.ToString)
                            WS.WriteLine("MBNo： " & stMultiBlock.BLOCK_DATA(blk - 1).DataNo.ToString)

                            WS.WriteLine("  Element = " & .stCounter1.TrimCounter.ToString & " Good = " & .stCounter1.OK_Counter.ToString & " pcs  Reject = " & .stCounter1.NG_Counter.ToString & " pcs")
                            WS.WriteLine("Initilal Low = " & .stCounter1.ITLow.ToString.PadRight(20) & "High = " & .stCounter1.ITHigh.ToString.PadRight(20) & "Open = " & .stCounter1.ITOpen.ToString)
                            WS.WriteLine("On Trim Low  = " & .stCounter1.FTLow.ToString.PadRight(10) & " ( " & .stCounter1.ValLow.ToString.PadRight(4) & " ) " & "High = " & .stCounter1.FTHigh.ToString.PadRight(10) & " ( " & .stCounter1.ValHigh.ToString.PadRight(4) & " ) " & "Open = " & .stCounter1.FTOpen.ToString)

                        End With

                    End If

                Next blk

            Else
                WS.WriteLine("No." & stCounter.PlateCounter.ToString & "  Start = " & stCounter.StartTime.ToString("HH:mm:ss") & " End = " & stCounter.EndTime.ToString("HH:mm:ss") & "  Element = " & stCounter.TrimCounter.ToString & " Good = " & stCounter.OK_Counter.ToString & " pcs  Reject = " & stCounter.NG_Counter.ToString & " pcs")
                ''V2.2.0.0⑯            WS.WriteLine("Initilal Low = " & stCounter.ITLow.ToString & "  High =" & stCounter.ITHigh.ToString & "  Open = " & stCounter.ITOpen.ToString & "         On Trim Low = " & stCounter.FTLow.ToString & "  High = " & stCounter.FTHigh.ToString & "  Open = " & stCounter.FTOpen.ToString)
                WS.WriteLine("Initilal Low = " & stCounter.ITLow.ToString.PadRight(20) & "High = " & stCounter.ITHigh.ToString.PadRight(20) & "Open = " & stCounter.ITOpen.ToString)       'V2.2.0.0⑯
                WS.WriteLine("On Trim Low  = " & stCounter.FTLow.ToString.PadRight(10) & " ( " & stCounter.ValLow.ToString.PadRight(4) & " ) " & "High = " & stCounter.FTHigh.ToString.PadRight(10) & " ( " & stCounter.ValHigh.ToString.PadRight(4) & " ) " & "Open = " & stCounter.FTOpen.ToString)      'V2.2.0.0⑯
            End If

            WS.Close()
        Catch ex As Exception
            Call Z_PRINT("UserSub.SubstrateEnd() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Sub

    'V2.0.0.0⑨↓
    ''' <summary>
    ''' 統計データの出力
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub StatisticalPrintDataOut()
        Try
            Dim WS As IO.StreamWriter
            Dim iResCnt As Integer
            Dim JudgeMode As Integer = FINAL_TEST
            Dim dMin As Double, dMax As Double, dAve As Double, dDev As Double
            Dim No As Integer = 0


            iResCnt = GetRCountExceptMeasure()
            If iResCnt > MAX_RES_USER Then
                iResCnt = MAX_RES_USER
            End If

            WS = New IO.StreamWriter(cTRIM_PRINT_DATA_RES, True, System.Text.Encoding.GetEncoding("Shift-JIS"))     ' 抵抗データ印刷データ
            WS.WriteLine("───────────────────────────────────────────────")
            WS.WriteLine("抵抗名　    最小　        最大　        平均　        標準偏差")

            'V2.2.0.0⑯↓
            If stMultiBlock.gMultiBlock <> 0 Then
                ' 複数抵抗値対応時

                ' マルチブロックで設定されている分実行する 
                For blk As Integer = 1 To stMultiBlock.BLOCK_DATA.Length - 1

                    If stMultiBlock.BLOCK_DATA(blk - 1).gBlockCnt <> 0 Then       'V2.2.0.033

                        'V2.2.0.033 WS.WriteLine("マルチブロックNo： " & stMultiBlock.BLOCK_DATA(blk - 1).DataNo.ToString)
                        WS.WriteLine("MBNo： " & stMultiBlock.BLOCK_DATA(blk - 1).DataNo.ToString)

                        'V2.2.0.033 For rn As Integer = 1 To MAX_RES_USER
                        For rn As Integer = 1 To stPLT.RCount   'V2.2.0.033
                            If UserModule.IsCutResistor(rn) Then
                                Call gObjFrmDistribute.StatisticalDataGetMulti(JudgeMode, rn, dMin, dMax, dAve, dDev, blk)
                                WS.WriteLine(stREG(rn).strRNO & vbTab & UserSub.ChangeOverFlow(dMin.ToString(TARGET_DIGIT_DEFINE)).PadLeft(13) & " " & UserSub.ChangeOverFlow(dMax.ToString(TARGET_DIGIT_DEFINE)).PadLeft(13) & " " & UserSub.ChangeOverFlow(dAve.ToString(TARGET_DIGIT_DEFINE)).PadLeft(13) & " " & UserSub.ChangeOverFlow(dDev.ToString(TARGET_DIGIT_DEFINE)).PadLeft(13))
                            End If
                        Next rn

                    End If

                Next blk

            Else

                For rn As Integer = 1 To stPLT.RCount
                    If UserModule.IsCutResistor(rn) Then
                        No = No + 1
                        Call gObjFrmDistribute.StatisticalDataGet(JudgeMode, No, dMin, dMax, dAve, dDev)
                        WS.WriteLine(stREG(rn).strRNO & vbTab & UserSub.ChangeOverFlow(dMin.ToString(TARGET_DIGIT_DEFINE)).PadLeft(13) & " " & UserSub.ChangeOverFlow(dMax.ToString(TARGET_DIGIT_DEFINE)).PadLeft(13) & " " & UserSub.ChangeOverFlow(dAve.ToString(TARGET_DIGIT_DEFINE)).PadLeft(13) & " " & UserSub.ChangeOverFlow(dDev.ToString(TARGET_DIGIT_DEFINE)).PadLeft(13))
                        If No >= iResCnt Then
                            Exit For
                        End If
                    End If
                Next

            End If
            'V2.2.0.0⑯↑

            WS.Close()

        Catch ex As Exception
            Call Z_PRINT("UserSub.StatisticalPrintDataOut() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub
    'V2.0.0.0⑨↑
    '===============================================================================
    '【機　能】 ロット終了時の印刷処理を行う
    '【引　数】 無し
    '【戻り値】 無し
    '===============================================================================
    Public Sub LotEnd()
        Dim WS As IO.StreamWriter

        Try
            UserBas.stCounter.LotEnd = DateTime.Now()           ' ロット終了時間

            Call StatisticalPrintDataOut()                      ' 統計データ出力'V2.0.0.0⑨

            If IO.File.Exists(cTRIM_PRINT_DATA_END) = True Then ' 毎回ファイルを削除する。
                IO.File.Delete(cTRIM_PRINT_DATA_END)
            End If
            WS = New IO.StreamWriter(cTRIM_PRINT_DATA_END, True, System.Text.Encoding.GetEncoding("Shift-JIS"))   ' ロット終了時データ印刷
            WS.WriteLine("───────────────────────────────────────────────")
            If stMultiBlock.gMultiBlock <> 0 Then

                'WS.WriteLine("Substrate = " & stCounter.PlateCounter.ToString())

                '' マルチブロックで設定されている分実行する 
                'For blk As Integer = 1 To stMultiBlock.BLOCK_DATA.Length - 1

                '    ' 'V2.2.0.033 If stMultiBlock.BLOCK_DATA(0).gBlockCnt <> 0 Then
                '    If stMultiBlock.BLOCK_DATA(blk - 1).gBlockCnt <> 0 Then       'V2.2.0.033

                '        With stToTalDataMulti(blk)

                '            ' 'V2.2.0.033 WS.WriteLine("マルチブロックNo： " & stMultiBlock.BLOCK_DATA(blk - 1).DataNo.ToString)
                '            WS.WriteLine("MBNo： " & stMultiBlock.BLOCK_DATA(blk - 1).DataNo.ToString)

                '            WS.WriteLine("   Element = " & .stCounter1.Total_TrimCounter.ToString & " pcs  Good = " & .stCounter1.Total_OK_Counter.ToString & " pcs  Reject = " & .stCounter1.Total_NG_Counter.ToString & " pcs")
                '            WS.WriteLine("Initilal Low =  " & .stCounter1.Total_ITLow.ToString & "  High = " & .stCounter1.Total_ITHigh.ToString & " Open = " & .stCounter1.Total_ITOpen.ToString)
                '            WS.WriteLine("On Trim Low =  " & .stCounter1.Total_FTLow.ToString & " ( " & .stCounter1.Total_ValLow.ToString & " ) " & "  High = " & .stCounter1.Total_FTHigh.ToString & " ( " & .stCounter1.Total_ValHigh.ToString & " ) " & " Open = " & .stCounter1.Total_FTOpen.ToString)

                '        End With

                '    End If

                'Next blk

            Else

                WS.WriteLine("Substrate = " & stCounter.PlateCounter.ToString & "   Element = " & stCounter.Total_TrimCounter.ToString & " pcs  Good = " & stCounter.Total_OK_Counter.ToString & " pcs  Reject = " & stCounter.Total_NG_Counter.ToString & " pcs")
                'V2.2.0.029           WS.WriteLine("Initilal Low =  " & stCounter.Total_ITLow.ToString & "  High = " & stCounter.Total_ITHigh.ToString & " Open = " & stCounter.Total_ITOpen.ToString & "         On Trim Low =  " & stCounter.Total_FTLow.ToString & "  High = " & stCounter.Total_FTHigh.ToString & " Open = " & stCounter.Total_FTOpen.ToString)
                'V2.2.1.1① WS.WriteLine("Initilal Low =  " & stCounter.Total_ITLow.ToString & "  High = " & stCounter.Total_ITHigh.ToString & " Open = " & stCounter.Total_ITOpen.ToString)            'V2.2.0.029
                'V2.2.1.1① WS.WriteLine("On Trim Low =  " & stCounter.Total_FTLow.ToString & " ( " & stCounter.Total_ValLow.ToString & " ) " & "  High = " & stCounter.Total_FTHigh.ToString & " ( " & stCounter.Total_ValHigh.ToString & " ) " & " Open = " & stCounter.Total_FTOpen.ToString)     'V2.2.0.029
                WS.WriteLine("Initilal Low = " & stCounter.Total_ITLow.ToString.PadRight(20) & "High = " & stCounter.Total_ITHigh.ToString.PadRight(20) & "Open = " & stCounter.Total_ITOpen.ToString.PadRight(10))            'V2.2.1.1①
                WS.WriteLine("On Trim Low  = " & stCounter.Total_FTLow.ToString.PadRight(10) & " ( " & stCounter.Total_ValLow.ToString.PadRight(4) & " ) " & "High = " & stCounter.Total_FTHigh.ToString.PadRight(10) & " ( " & stCounter.Total_ValHigh.ToString.PadRight(4) & " ) " & "Open = " & stCounter.Total_FTOpen.ToString.PadRight(10))     'V2.2.1.1①


            End If
            WS.WriteLine("───────────────────────────────────────────────")
            WS.WriteLine("設定データ確認時間：　" & stCounter.LotStart.ToString("HH:mm:ss") & " 　終了時間： " & stCounter.LotEnd.ToString("HH:mm:ss") & " 　経過時間：　" & stCounter.LotEnd.Subtract(stCounter.LotStart).Hours.ToString("00") & ":" & stCounter.LotEnd.Subtract(stCounter.LotStart).Minutes.ToString("00") & ":" & stCounter.LotEnd.Subtract(stCounter.LotStart).Seconds.ToString("00"))
            WS.WriteLine("───────────────────────────────────────────────")
            WS.Close()


            UserSub.VariationMesStartDataReset()                'V2.0.0.0② 測定値変動検出機能開始ブロック位置初期化

            WriteLogMarkPrint()         ' ロットのログ内容をファイル出力               V2.2.1.7③

        Catch ex As Exception
            Call Z_PRINT("UserSub.LotEnd() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Sub

    Public Function ChangeOverFlow(ByVal sNum As String) As String
        Dim iPos As Integer

        iPos = sNum.IndexOf(".")
        If iPos > 9 Or (iPos = -1 And sNum.Length > 9) Then         ' １００Ｍ以上の桁または小数点が無くて文字列９以上の場合は、０にする。
            'V2.0.0.0⑤            Return ("0.00000")
            Return (TARGET_DIGIT_DEFINE)                            'V2.0.0.0⑤
        Else
            Return (sNum)
        End If

    End Function

    'V1.0.4.3⑨↓
    ''' <summary>
    ''' 増設リレーボード対応、チャンネル変換　ｃｈ７～１６⇒ｃｈ３３～４２
    ''' </summary>
    ''' <param name="ProbeChannel"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ConvtChannel(ByRef ProbeChannel As Short) As Short
        Try
            'V2.0.0.0⑬↓
            If bRelayBoard Then
                If 9 <= ProbeChannel And ProbeChannel <= 18 Then
                    ConvtChannel = 24 + ProbeChannel
                Else
                    ConvtChannel = ProbeChannel
                End If
            Else
                'V2.0.0.0⑬↑
                If 7 <= ProbeChannel And ProbeChannel <= 16 Then
                    ConvtChannel = 26 + ProbeChannel
                Else
                    ConvtChannel = ProbeChannel
                End If
            End If                                      'V2.0.0.0⑬
        Catch ex As Exception
            Call Z_PRINT("UserSub.ConvertChannel() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function
    'V1.0.4.3⑨↑
    'V1.0.4.3⑩↓
    ''' <summary>
    ''' 測定マーキングモード・ファイナル測定のみの場合もTRIM_MODE_ITTRFTと判定する。
    ''' </summary>
    ''' <returns>True:TRIM_MODE_ITTRFT　False:TRIM_MODE_ITTRFTとTRIM_MODE_MEAS_MARK以外</returns>
    ''' <remarks></remarks>
    Public Function IsTRIM_MODE_ITTRFT() As Boolean
        Try
            'V2.0.0.0②            If (DGL = TRIM_MODE_ITTRFT Or DGL = TRIM_MODE_MEAS_MARK) Then
            If (DGL = TRIM_MODE_ITTRFT Or DGL = TRIM_MODE_MEAS_MARK Or DGL = TRIM_MODE_POWER Or DGL = TRIM_VARIATION_MEAS) Then
                Return (True)
            Else
                Return (False)
            End If

        Catch ex As Exception

        End Try
    End Function
    'V1.0.4.3⑩↑

#Region "印刷処理ｸﾗｽ"
    '===============================================================================
    '【機　能】 印刷処理をおこなうｸﾗｽ
    '【仕　様】 Print ﾒｿｯﾄﾞを呼び出して印刷をおこなう
    '===============================================================================
    Public Class cPrintDocument
        Private Const FILE_ENCODING As String = "shift_jis"
        'Private Const FILE_ENCODING As String = "utf-8"
        Private ReadOnly FONT_SIZE As Font = New Font("ＭＳ ゴシック", 9.0!)
        Private ReadOnly FILEPATH_ARRAY As String() = {cTRIM_PRINT_DATA_HEAD, _
                                                       cTRIM_PRINT_DATA_RES, _
                                                       cTRIM_PRINT_DATA_PLATE, _
                                                       cTRIM_PRINT_DATA_END}
        Private Const MSG_YESNO As String = "ロット情報の印刷を実行します。"
        Private Const MSG_FILE_NOTHING As String = " が見つかりません。"
        Private Const MARGIN_LEFT As Integer = 100      ' 印刷ﾏｰｼﾞﾝ左(1/100 ｲﾝﾁ単位)ﾃﾞﾌｫﾙﾄは100
        Private Const MARGIN_RIGHT As Integer = 100     ' 印刷ﾏｰｼﾞﾝ右(1/100 ｲﾝﾁ単位)ﾃﾞﾌｫﾙﾄは100
        Private Const MARGIN_TOP As Integer = 10        ' 印刷ﾏｰｼﾞﾝ上(1/100 ｲﾝﾁ単位)ﾃﾞﾌｫﾙﾄは100
        Private Const MARGIN_BOTTOM As Integer = 10     ' 印刷ﾏｰｼﾞﾝ下(1/100 ｲﾝﾁ単位)ﾃﾞﾌｫﾙﾄは100

        Private Const DEF_TEXT_BUF As Integer = 1024    ' すべての文字数ﾊﾞｯﾌｧ(足りなくなれば拡張される)
        Private Const DEF_LINE_BUF As Integer = 128     ' 一行の文字数ﾊﾞｯﾌｧ(足りなくなれば拡張される)
        Private m_PrintTextBuf As StringBuilder         ' 印刷する文字すべてを格納する
        Private m_BufIndex As Integer                   ' 現在の文字位置

        Private Const PRINTDEFAULT_DIR = "C:\TRIMDATA\PRINTDEFAULT"            ' V2.2.0.024
        Private Const PRINTLOG_DIR = "C:\TRIMDATA\PRINTLOG"            ' V2.2.2.0⑦ 

        ''' <summary>印刷をおこなう</summary>
        ''' <param name="confirmMsgBox">Trueの場合確認ﾒｯｾｰｼﾞを表示する</param>
        Public Sub Print(Optional ByVal confirmMsgBox As Boolean = False)

            If (True = confirmMsgBox) Then ' 確認ﾒｯｾｰｼﾞの表示
                If (MsgBoxResult.No = MsgBox(MSG_YESNO, DirectCast( _
                                             MsgBoxStyle.YesNo + _
                                             MsgBoxStyle.Question, MsgBoxStyle), _
                                             My.Application.Info.Title)) _
                                             Then Exit Sub
            End If

            Try
                'V2.2.1.7④↓
                ' マーク印字モードは印刷しない 
                If UserSub.IsTrimType5() Then
                    Return
                End If
                'V2.2.1.7④↑

                Using printer As New PrintDocument
                    m_PrintTextBuf = New StringBuilder(DEF_TEXT_BUF)
                    m_BufIndex = 0

                    For Each path As String In FILEPATH_ARRAY
                        ''V2.2.0.033↓
                        'If (stMultiBlock.gMultiBlock <> 0) AndAlso (path = cTRIM_PRINT_DATA_END) Then
                        '    Continue For
                        'End If
                        ''V2.2.0.033↑

                        If (True = File.Exists(path)) Then
                            ' ﾌｧｲﾙが存在する場合
                            Using sr As New StreamReader(path, System.Text.Encoding.GetEncoding(FILE_ENCODING))
                                While (-1 < sr.Peek()) ' 使用できる文字がなくなるまで継続
                                    m_PrintTextBuf.Append(sr.ReadLine() & vbLf) ' 一行づつ追加する
                                End While
                            End Using
                            'V2.2.2.0⑦↓
                            Dim orgfilename As String = IO.Path.GetFileName(path)
                            Dim writefolder As String = PRINTLOG_DIR & "\" & DateTime.Now.ToString("yyyyMM") & "\"
                            '                            If (False = System.IO.File.Exists(writefolder)) Then
                            If (False = IO.Directory.Exists(writefolder)) Then
                                MkDir(writefolder)
                            End If
                            FileCopy(path, writefolder & "\" & DateTime.Now.ToString("yyyyMMdd_hhmmss_") & stUserData.sLotNumber.Trim() & "_" & orgfilename)
                            'V2.2.2.0⑦↑

                        Else
                            ' ﾌｧｲﾙが存在しない旨を追加する
                            m_PrintTextBuf.Append(vbLf & path & MSG_FILE_NOTHING & vbLf & vbLf)
                        End If
                    Next

                    With printer.DefaultPageSettings
                        .Margins = New Margins(MARGIN_LEFT, MARGIN_RIGHT, MARGIN_TOP, MARGIN_BOTTOM)
                        .Landscape = False ' 用紙の向き(縦)
                    End With
                    AddHandler printer.PrintPage, AddressOf m_Printer_PrintPage

                    'V2.2.0.024↓
                    If printer.PrinterSettings.PrinterName = "Microsoft Print to PDF" Then
                        printer.PrinterSettings.PrintToFile = True
                        ' PDFの出力先とファイル名を指定
                        printer.PrinterSettings.PrintFileName = PRINTDEFAULT_DIR & "\" & System.IO.Path.GetFileNameWithoutExtension(gsDataFileName) & "_" & Now.ToString("yyyyMMddHHmmss") & ".pdf"
                    End If
                    'V2.2.0.024↑
                    Call printer.Print() ' m_Printer_PrintPage() ｲﾍﾞﾝﾄが発生する



                End Using

            Catch ex As Exception
                Call MsgBox(ex.ToString())
            End Try

        End Sub

        ''' <summary>文字を書き出す</summary>
        ''' <remarks>一行ごとに先頭座標を指定して描画する</remarks>
        Private Sub m_Printer_PrintPage(ByVal sender As Object, ByVal e As PrintPageEventArgs)
            Dim x As Integer = e.MarginBounds.Left ' 印刷開始初期位置
            Dim y As Integer = e.MarginBounds.Top  ' 印刷開始初期位置

            Try
                ' 現在のﾍﾟｰｼﾞにおさまる かつ 文字をすべて書き出していない場合継続
                While ((y + FONT_SIZE.Height) < e.MarginBounds.Bottom) AndAlso _
                        (m_BufIndex < m_PrintTextBuf.Length)
                    Dim lineBuf As New StringBuilder(DEF_LINE_BUF) ' 行ﾊﾞｯﾌｧ

                    While True
                        If (m_PrintTextBuf.Length <= m_BufIndex) OrElse _
                            (vbLf = m_PrintTextBuf.Chars(m_BufIndex)) Then
                            ' 文字をすべて書き出した または 改行ｺｰﾄﾞの場合
                            m_BufIndex += 1
                            Exit While
                        End If

                        lineBuf.Append(m_PrintTextBuf.Chars(m_BufIndex)) ' 一文字追加する
                        If ((e.MarginBounds.Width) < _
                            (e.Graphics.MeasureString(lineBuf.ToString(), FONT_SIZE).Width)) Then
                            ' 印刷幅におさまらない場合、一文字削除する
                            lineBuf.Remove(lineBuf.Length - 1, 1)
                            Exit While
                        End If

                        m_BufIndex += 1
                    End While

                    ' 一行分書き出す
                    'Debug.Print(lineBuf.ToString())
                    e.Graphics.DrawString(lineBuf.ToString(), FONT_SIZE, Brushes.Black, x, y)
                    y += FONT_SIZE.GetHeight(e.Graphics) ' 次の行の印刷位置へ
                End While

                If (m_PrintTextBuf.Length <= m_BufIndex) Then
                    ' 文字をすべて書き出した場合
                    e.HasMorePages = False
                    'Debug.Print((m_PrintTextBuf.Capacity).ToString()) ' DEF_TEXT_BUF のｻｲｽﾞ調整
                    m_PrintTextBuf = Nothing
                    m_BufIndex = 0
                Else
                    ' この設定により再度 m_Printer_PrintPage() ｲﾍﾞﾝﾄが発生する
                    e.HasMorePages = True ' 次のﾍﾟｰｼﾞへ
                End If

            Catch ex As Exception
                Call MsgBox(ex.ToString())
            End Try
        End Sub

    End Class
#End Region
    ''' <summary>
    ''' 自動運転時のオフセットパラメータ反映処理（テーブル位置、ビーム位置、プローブ接触位置）
    ''' </summary>
    ''' <param name="AutoDataFileFullPath"></param>
    ''' <param name="iAutoDataFileNum"></param>
    ''' <returns>正常終了：cFRS_NORMAL　異常終了：データの番号</returns>
    ''' <remarks></remarks>
    Public Function SetOffSetDataToAutoOperationData(ByVal AutoDataFileFullPath() As String, ByVal iAutoDataFileNum As Short) As Short

        Dim r As Short
        Dim stPLT_Local As PLATE_DATA                          ' プレートデータ

        If iAutoDataFileNum <= 0 Then
            Return (True)
        End If

        gsDataFileName = AutoDataFileFullPath(0)                     ' データファイル名設定

        r = rData_load()                                            ' データファイルリード
        If (r <> 0) Then                                            ' データファイル　ロードエラー
            Return (1)
        End If

        stPLT_Local.z_xoff = stPLT.z_xoff                           ' テーブル位置オフセット　トリムポジションオフセットX(mm)
        stPLT_Local.z_yoff = stPLT.z_yoff                           ' テーブル位置オフセット　トリムポジションオフセットY(mm)

        stPLT_Local.BPOX = stPLT.BPOX                               ' ビーム位置オフセット　BP Offset X(mm)
        stPLT_Local.BPOY = stPLT.BPOY                               ' ビーム位置オフセット　BP Offset Y(mm)

        stPLT_Local.Z_ZON = stPLT.Z_ZON                             ' Z PROBE ON 位置(mm)

        For Cnt As Integer = 1 To (iAutoDataFileNum - 1)
            gsDataFileName = AutoDataFileFullPath(Cnt)
            r = rData_load()                                            ' データファイルリード
            If (r <> 0) Then                                            ' データファイル　ロードエラー
                Return (Cnt + 1)
            End If
            stPLT.z_xoff = stPLT_Local.z_xoff       ' テーブル位置オフセット　トリムポジションオフセットX(mm)
            stPLT.z_yoff = stPLT_Local.z_yoff       ' テーブル位置オフセット　トリムポジションオフセットY(mm)
            stPLT.BPOX = stPLT_Local.BPOX           ' ビーム位置オフセット　BP Offset X(mm)
            stPLT.BPOY = stPLT_Local.BPOY           ' ビーム位置オフセット　BP Offset Y(mm)
            stPLT.Z_ZON = stPLT_Local.Z_ZON         ' Z PROBE ON 位置(mm)

            If rData_save(gsDataFileName) <> cFRS_NORMAL Then       ' データファイルセーブ
                Return (Cnt + 1)
            End If
        Next

        Return (cFRS_NORMAL)

    End Function

    'V1.2.0.0③↓
    ''' <summary>
    ''' パターン認識結果格納領域の初期化（初期状態はＯＫ）
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub InitResPatternmatchResult()
        Try
            Dim iResNo As Integer
            For iResNo = 1 To MAXRNO Step 1
                stREG(iResNo).bPattern = True
            Next
        Catch ex As Exception
            Call Z_PRINT("UserSub.InitResPatternmatchResult() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub
    'V1.2.0.0③↑
    'V1.2.0.0②↓
    Public Function MarkingForChipMode(ByVal rn As Short, ByVal Result As Boolean) As Short
        Try
            Dim rNo As Integer
            Dim cn As Integer
            Dim dOffSetX As Double, dOffSetY As Double
            Dim Rtn As Short = cFRS_NORMAL              ' 関数戻値

            If Not UserSub.IsTrimType3() And Not UserSub.IsTrimType4() Then     ' 温度センサー'V2.0.0.0①sTrimType4()追加
                Return (Rtn)
            End If

            ' 第１抵抗第１カットのスタート座標から現在の抵抗の第１カットのスタート座標までの距離を求める。
            dOffSetX = stREG(rn).STCUT(1).dblSTX - stREG(1).STCUT(1).dblSTX
            dOffSetY = stREG(rn).STCUT(1).dblSTY - stREG(1).STCUT(1).dblSTY

            'V2.0.0.0⑩↓
            If UserSub.IsTrimType3() Then
                rn = UserSub.GetTopResNoinCircuit(rn)
            End If
            'V2.0.0.0⑩↑

            'V1.2.0.2↓
            For rNo = 1 To stPLT.RCount Step 1
                If IsCutResistor(rNo) Then
                    dOffSetX = stREG(rn).STCUT(1).dblSTX - stREG(rNo).STCUT(1).dblSTX
                    dOffSetY = stREG(rn).STCUT(1).dblSTY - stREG(rNo).STCUT(1).dblSTY
                    Exit For
                End If
            Next
            'V1.2.0.2↑

            'V1.2.0.2            For rNo = 1 To MAXRNO Step 1
            For rNo = 1 To stPLT.RCount Step 1                  'V1.2.0.2
                If Result Then                                  ' ＯＫ判定の時
                    If stREG(rNo).intSLP <> SLP_OK_MARK Then    ' ＯＫマーク以外はスキップ
                        Continue For
                    End If
                Else                                            ' ＮＧ判定の時
                    If stREG(rNo).intSLP <> SLP_NG_MARK Then    ' ＮＧマーク以外はスキップ
                        Continue For
                    End If
                End If
                ' カット位置を現在の抵抗の位置に合わせてオフセットさせる。
                For cn = 1 To stREG(rNo).intTNN Step 1
                    stREG(rNo).STCUT(cn).dblSTX = stREG(rNo).STCUT(cn).dblSTX + dOffSetX
                    stREG(rNo).STCUT(cn).dblSTY = stREG(rNo).STCUT(cn).dblSTY + dOffSetY
                Next
                Rtn = VTrim_One(rNo, stREG(rNo).dblNOM)          ' 1抵抗分トリミングを行う
                ' カット位置を元に戻す。
                For cn = 1 To stREG(rNo).intTNN Step 1
                    stREG(rNo).STCUT(cn).dblSTX = stREG(rNo).STCUT(cn).dblSTX - dOffSetX
                    stREG(rNo).STCUT(cn).dblSTY = stREG(rNo).STCUT(cn).dblSTY - dOffSetY
                Next
            Next

            Return (Rtn)

        Catch ex As Exception
            Call Z_PRINT("UserSub.MarkingForChipMode() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function
    'V1.2.0.0②↑

    'V2.0.0.0②↓
    ''' <summary>
    ''' ログデータからFT値の取得
    ''' </summary>
    ''' <param name="sLotNumber">ロット番号</param>
    ''' <param name="PlateNumber">基板番号</param>
    ''' <param name="BlockX">ブロックＸ番号</param>
    ''' <param name="BlockY">ブロックＹ番号</param>
    ''' <param name="ResCounter">検査して求めたデータ数</param>
    ''' <param name="Target">ＦＴ値（配列）</param>
    ''' <returns></returns>
    ''' <remarks></remarks>    
    Public Function GetTargerDataFromLogFile(ByVal sLotNumber As String, ByVal PlateNumber As Integer, ByVal BlockX As Integer, ByVal BlockY As Integer, ByRef ResCounter As Integer, ByRef Target() As Double) As Boolean

        Dim sPath As String
        Dim sInpData As String
        Dim splt() As String
        Dim ITEM_LOT_NUM As Integer = 3
        Dim ITEM_PLATE_NUM As Integer = 4
        Dim ITEM_BLOCKX_NUM As Integer = 5
        Dim ITEM_BLOCKY_NUM As Integer = 6
        Dim ITEM_FT_NUM As Integer = 12
        Dim itemcnt As Integer
        Dim sLot As String
        Dim nPlate As Integer
        Dim nBlockX As Integer
        Dim nBlockY As Integer
        Dim SearchDir As String
        Dim SearchFile As String

        Try

            SearchDir = "C:\TRIMDATA\LOG"                                   'サーチするフォルダ指定
            SearchFile = "*" + sLotNumber.Trim() + ".CSV"                   'サーチするファイルの検索キー(ロット番号が含まれているCSVファイル)

            '条件が一致するファイル名の取得実行 「検索対象は指定フォルダのみ」サブフォルダは除外.サブを含める場合は最後の引数を「SearchOption.AllDirectories」に変更
            Dim files() As String = System.IO.Directory.GetFiles(SearchDir, SearchFile, SearchOption.TopDirectoryOnly)

            ' 取得したすべてのファイルを最終書き込み日時順でソートする)
            Array.Sort(Of String)(files, AddressOf CompareLastWriteTime)
            If (files.Length = 0) Then
                Return (False)
            End If
            sPath = files(files.Length - 1)
            Using sr As New StreamReader(sPath, Encoding.GetEncoding("Shift_JIS"))

                'タイトル読み飛ばし
                sInpData = sr.ReadLine()

                itemcnt = 0
                '最終行まで１行ごとにファイル読込み
                Do While (False = sr.EndOfStream)
                    sInpData = sr.ReadLine()                                ' １行読込み
                    splt = sInpData.Split(","c)                             ' カンマ区切りで分割

                    sLot = splt(ITEM_LOT_NUM)                               ' ロット番号の取得 
                    nPlate = splt(ITEM_PLATE_NUM)                           ' 基板番号の取得 
                    nBlockX = splt(ITEM_BLOCKX_NUM)                         ' Block番号Xの取得 
                    nBlockY = splt(ITEM_BLOCKY_NUM)                         ' Block番号Yの取得 

                    'ロット番号、基板番号、Block番号X、Yが一致するデータのみ抽出 
                    If ((sLot = sLotNumber) AndAlso (nPlate = PlateNumber) AndAlso (nBlockX = BlockX) AndAlso (nBlockY = BlockY)) Then
                        Target(itemcnt) = splt(ITEM_FT_NUM)              ' FT結果の取得 
                        itemcnt = itemcnt + 1
                        '検出した項目数の設定
                        ResCounter = itemcnt
                    Else
                        If ResCounter > 0 Then
                            Exit Do
                        End If
                        itemcnt = 0
                    End If
                Loop
            End Using

            Return (True)
        Catch ex As Exception
            Call Z_PRINT("UserSub.GetTargerDataFromLogFile() TRAP ERROR = " & ex.Message & vbCrLf)
            Return (False)
        End Try
    End Function

    ' 二つのファイルの最終書き込み日時を取得して比較するメソッド
    Private Function CompareLastWriteTime(ByVal fileX As String, ByVal fileY As String) As Integer
        Return DateTime.Compare(File.GetLastWriteTime(fileX), File.GetLastWriteTime(fileY))
    End Function

    ' 測定値変動検出機能
    Public bVariationMesStep As Boolean = True
    Public gVariationMeasPlateStartNo As Integer = 1
    Public gVariationMeasBlockXStartNo As Integer = 1
    Public gVariationMeasBlockYStartNo As Integer = 1
    Public dMeasVariationNOM(MAXRNO) As Double                  ' トリミング後ＦＴ値
    Public dMeasVariationDev(MAX_RES_USER) As Double            ' 変化量

    ''' <summary>
    ''' 測定値変動検出機能開始ブロック位置初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub VariationMesStartDataReset()
        Try
            gVariationMeasPlateStartNo = 1
            gVariationMeasBlockXStartNo = 1
            gVariationMeasBlockYStartNo = 1
            'V2.0.0.1①            bVariationMesStep = True
            bVariationMesStep = False       'V2.0.0.1①
        Catch ex As Exception
            Call Z_PRINT("UserSub.VariationMesStartDataReset() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub
    Public Function SetTarrgetOnVariationMeas() As Boolean
        Try
            Dim bRtn As Boolean
            Dim ResCounter As Integer
            Dim Target(MAXRNO) As Double
            Dim rno As Integer

            For rn As Short = 1 To stPLT.RCount
                dMeasVariationNOM(rn) = 0.0
            Next

            bRtn = GetTargerDataFromLogFile(stUserData.sLotNumber, stCounter.PlateCounter, stCounter.BlockCntX, stCounter.BlockCntY, ResCounter, Target)
            If Not bRtn Then
                Call DebugLogOut("測定値変動検出 目標値設定エラー LOT=[" & stUserData.sLotNumber & "] PLATE=[" & stCounter.PlateCounter.ToString & "]X=[" & stCounter.BlockCntX.ToString & "]Y=[" & stCounter.BlockCntY.ToString & "]")
                Return (False)
            End If
            Dim Rcnt As Integer = UserBas.GetRCountExceptMeasure()

            rno = 0
            For rn As Integer = 1 To stPLT.RCount
                If UserModule.IsCutResistor(rn) Then
                    If rno < ResCounter Then
                        dMeasVariationNOM(rn) = Target(rno)
                        rno = rno + 1
                    Else
                        Call Z_PRINT("測定値変動検出 目標値データが有りません LOT=[" & stUserData.sLotNumber & "] PLATE=[" & stCounter.PlateCounter.ToString & "]X=[" & stCounter.BlockCntX.ToString & "]Y=[" & stCounter.BlockCntY.ToString & "]RES=[" & rn.ToString & "]")
                        Call DebugLogOut("測定値変動検出 目標値設定エラー LOT=[" & stUserData.sLotNumber & "] PLATE=[" & stCounter.PlateCounter.ToString & "]X=[" & stCounter.BlockCntX.ToString & "]Y=[" & stCounter.BlockCntY.ToString & "]]RES=[" & rn.ToString & "]")
                    End If
                End If
            Next
            Return (True)
        Catch ex As Exception
            Call Z_PRINT("UserSub.SetTarrgetOnVariationMeas() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function

    Public Function VariationMeasJudge(ByVal rn As Short, ByVal dblMx As Double) As Boolean
        Try
            dMeasVariationDev(rn) = 0.0

            If dMeasVariationNOM(rn) = 0.0 Then
                Return (False)
            End If

            ' トリミング誤差絶対値　＝　（　トリミング値　－　トリミング時のＴＦ値　）／トリミング時のＴＦ値　* 10^6
            dMeasVariationDev(rn) = FNDEVP(dblMx, dMeasVariationNOM(rn))

            If dMeasVariationDev(rn) > Math.Abs(stUserData.dVariation) Then
                Return (False)
            Else
                Return (True)
            End If

        Catch ex As Exception
            Call Z_PRINT("UserSub.VariationMeasJudge() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function

    'V2.0.0.0②↑

    'V2.0.0.0⑩↓
    ''' <summary>
    ''' サーキット総数のカウント
    ''' </summary>
    ''' <param name="stPlate"></param>
    ''' <param name="stRegData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCircuitSum(ByVal stPlate As PLATE_DATA, ByRef stRegData As Reg_Info()) As Integer
        Try
            Dim iResCnt As Integer = 0
            Dim iCircuit As Integer = -1

            For rn As Integer = 1 To stPlate.RCount
                If UserModule.IsCutResistor(stRegData, rn) Then
                    If iCircuit <> stRegData(rn).intCircuitNo Then
                        iResCnt = iResCnt + 1
                    End If
                    iCircuit = stRegData(rn).intCircuitNo
                End If
            Next
            Return (iResCnt)

        Catch ex As Exception
            Call Z_PRINT("UserSub.GetCircuitSum() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function
    ''' <summary>
    ''' 抵抗番号からサーキット番号を取得する。
    ''' </summary>
    ''' <param name="stPlate"></param>
    ''' <param name="stRegData"></param>
    ''' <param name="rno"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCircuitNoFromResNo(ByVal stPlate As PLATE_DATA, ByRef stRegData As Reg_Info(), ByRef rno As Short) As Integer
        Try
            Dim iCircuitNo As Integer = 0
            Dim iCircuit As Integer = -1

            For rn As Integer = 1 To stPlate.RCount
                If UserModule.IsCutResistor(stRegData, rn) Then
                    If iCircuit <> stRegData(rn).intCircuitNo Then
                        iCircuitNo = iCircuitNo + 1
                    End If
                    iCircuit = stRegData(rn).intCircuitNo
                    If rn = rno Then
                        Return (iCircuitNo)
                    End If
                End If
            Next
        Catch ex As Exception
            Call Z_PRINT("UserSub.GetCircuitNoFromResNo() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function
    ''' <summary>
    ''' サーキット内の抵抗数のカウント
    ''' </summary>
    ''' <param name="stRegData">抵抗データ構造体</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CircuitResistorCount(ByVal stPlate As PLATE_DATA, ByRef stRegData As Reg_Info()) As Integer
        Try
            Dim iResCnt As Integer = 0
            Dim iCircuit As Integer = -1

            For rn As Integer = 1 To stPlate.RCount
                If UserModule.IsCutResistor(stRegData, rn) Then
                    If iCircuit < 0 Then
                        iCircuit = stRegData(rn).intCircuitNo
                    End If
                    If stRegData(rn).intCircuitNo <> iCircuit Then
                        Return (iResCnt)
                    End If
                    iResCnt = iResCnt + 1
                End If
            Next
            Return (iResCnt)

        Catch ex As Exception
            Call Z_PRINT("UserSub.CircuitResistorCount() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function

    ''' <summary>
    ''' サーキット内の抵抗数のカウント
    ''' </summary>
    ''' <returns>同一サーキット内抵抗数</returns>
    ''' <remarks></remarks>
    Public Function CircuitResistorCount() As Integer
        Try
            Return (CircuitResistorCount(stPLT, stREG))

        Catch ex As Exception
            Call Z_PRINT("UserSub.CircuitResistorCount() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function

    ''' <summary>
    ''' サーキットの最後の抵抗かをチェックする。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsCheckCircuitEnd(ByVal rn As Short) As Boolean
        Try
            Dim iCircuit As Integer = stREG(rn).intCircuitNo

            For i As Integer = (rn + 1) To stPLT.RCount
                If UserModule.IsCutResistor(i) Then
                    If iCircuit = stREG(i).intCircuitNo Then    ' 後に同じ抵抗が出て来たら最後では無い
                        Return (False)
                    Else
                        Return (True)
                    End If
                End If
            Next

            Return (True)

        Catch ex As Exception
            Call Z_PRINT("UserSub.IsCheckCircuitEnd() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Function
    ''' <summary>
    ''' 現在の抵抗番号が最後のサーキットかを判定する。
    ''' </summary>
    ''' <param name="rn">抵抗番号</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsCheckLastCircuit(ByVal rn As Short) As Boolean
        Try
            Dim CircuitCnt As Integer = GetCircuitSum(stPLT, stREG)
            Dim CircuitNO As Integer = GetCircuitNoFromResNo(stPLT, stREG, rn)

            If CircuitNO = CircuitCnt Then
                Return (True)
            Else
                Return (False)
            End If

        Catch ex As Exception
            Call Z_PRINT("UserSub.IsCheckLastCircuit() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function
    ''' <summary>
    ''' 同じサーキット番号の先頭の抵抗番号を求める。
    ''' </summary>
    ''' <param name="rn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetTopResNoinCircuit(ByVal rn As Short) As Short
        Try
            Dim iCircuit As Integer = stREG(rn).intCircuitNo
            Dim TopResNo As Short = rn
            Dim ResNo As Short

            For ResNo = TopResNo To 1 Step -1
                If UserModule.IsCutResistor(ResNo) Then
                    If iCircuit = stREG(ResNo).intCircuitNo Then    ' 前に同じサーキット番号が出て来たら
                        TopResNo = ResNo
                    End If
                End If
            Next

            Return (TopResNo)

        Catch ex As Exception
            Call Z_PRINT("UserSub.GetTopResNoinCircuit() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function

    ''' <summary>
    ''' サーキット内で何番目の抵抗かを求める
    ''' </summary>
    ''' <param name="stRegData"></param>
    ''' <param name="rn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetResNumberInCircuit(ByRef stRegData As Reg_Info(), ByVal rn As Short) As Integer
        Try

            Dim iCircuit As Integer = stRegData(rn).intCircuitNo
            Dim iNumber As Integer = 0

            If UserSub.IsTrimType1() Or UserSub.IsTrimType4() Then  ' 温度センサーの時は１固定
                Return (1)
            End If

            For i As Integer = 1 To rn
                If UserModule.IsCutResistor(stRegData, i) Then
                    If iCircuit = stRegData(i).intCircuitNo Then    ' 同じサーキット番号
                        iNumber = iNumber + 1
                    End If
                End If
            Next

            Return (iNumber)

        Catch ex As Exception
            Call Z_PRINT("UserSub.GetResNumberInCircuit() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function

    ''' <summary>
    ''' サーキット内で何番目の抵抗かを求める
    ''' </summary>
    ''' <param name="rn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetResNumberInCircuit(ByVal rn As Short) As Integer
        Try

            Return (GetResNumberInCircuit(stREG, rn))

        Catch ex As Exception
            Call Z_PRINT("UserSub.GetResNumberInCircuit() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function

    ''' <summary>
    ''' サーキット内の順番の抵抗番号を求める
    ''' </summary>
    ''' <param name="stPlate"></param>
    ''' <param name="stRegData"></param>
    ''' <param name="Circuit"></param>
    ''' <param name="No"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRNumByCircuit(ByVal stPlate As PLATE_DATA, ByRef stRegData As Reg_Info(), ByVal Circuit As Short, ByVal No As Short) As Integer
        Try
            Dim iResCnt As Integer = 0
            Dim iCircuit As Integer = 0

            For rno As Integer = 1 To stPlate.RCount
                If UserModule.IsCutResistor(stRegData, rno) Then
                    If iCircuit <> stRegData(rno).intCircuitNo Then
                        iCircuit = iCircuit + 1
                    End If
                    If iCircuit = Circuit Then
                        iResCnt = iResCnt + 1
                        If iResCnt = No Then
                            Return (rno)
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            Call Z_PRINT("UserSub.GetRNumByCircuit() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function
    Public Function GetRNumByCircuit(ByVal Circuit As Short, ByVal No As Short) As Integer
        Try
            Return (GetRNumByCircuit(stPLT, stREG, Circuit, No))
        Catch ex As Exception
            Call Z_PRINT("UserSub.GetRNumByCircuit() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function
    'V2.0.0.0⑩↑

    'V2.0.0.0⑭↓
    ''' <summary>
    ''' クランプ吸着変更
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ClampVacumeChange()
        Try
            Dim r As Short

            If giLoaderType <> 0 Then   'クランプ吸着動作設定
                ObjSys.setClampVaccumConfig(stUserData.intClampVacume - 1)
            End If

            Select Case (stUserData.intClampVacume)
                Case CLAMP_VACCUME_USE
                Case CLAMP_ONLY_USE
                    Call Form1.System1.AbsVaccume(gSysPrm, 0, giAppMode, 0)
                    Call Form1.System1.Adsorption(gSysPrm, 0)
                Case VACCUME_ONLY_USE
                    r = Form1.System1.ClampCtrl(gSysPrm, 0, 0, False)                 ' クランプ/吸着OFF
                    If (r = cFRS_NORMAL) Then
                        'Call Sub_ATLDSET(COM_STS_CLAMP_ON, 0)                       ' ローダー出力(ON=載物台ｸﾗﾝﾌﾟ開,OFF=なし)
                        'gbClampOpen = False
                    Else
                        Call Z_PRINT("クランプ開エラーが発生しました。。" & vbCrLf)
                    End If
                    'If giLoaderType = 1 Then
                    '    Call Form1.System1.AbsVaccume(gSysPrm, 1, giAppMode, 0)
                    'End If
                Case Else
                    Throw New Exception("Case " & stUserData.intClampVacume & ": Nothing")
            End Select
        Catch ex As Exception
            Call Z_PRINT("UserSub.ClampVacumeChange() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub
    'V2.0.0.0⑭↑

    'V2.0.0.1③↓
#Region "基板処理枚数からの判定"
    ''' <summary>
    ''' 基板処理枚数からの判定(ＮＧ数比率判定）
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function PlateNGJudgeByCounter() As Boolean
        Dim dDrate As Double
        Try
            If UserBas.stCounter.TrimCounter = 0 Then
                Return (False)
            End If

            dDrate = UserBas.stCounter.NG_Counter / UserBas.stCounter.TrimCounter * 100.0

            If dDrate >= stUserData.NgJudgeRate Then
                DebugLogOut("１基板ＮＧ判定[" & dDrate.ToString & "]=[" & UserBas.stCounter.NG_Counter.ToString & "]/[" & UserBas.stCounter.TrimCounter.ToString & "] * 100.0 >= [" & stUserData.NgJudgeRate.ToString & "]")
                Z_PRINT("基板ＮＧ判定 比率[" & dDrate.ToString & "]=[" & UserBas.stCounter.NG_Counter.ToString & "]/[" & UserBas.stCounter.TrimCounter.ToString & "] * 100.0 >= [" & stUserData.NgJudgeRate.ToString & "]")
                Return (True)
            Else
                DebugLogOut("１基板ＯＫ判定[" & dDrate.ToString & "]=[" & UserBas.stCounter.NG_Counter.ToString & "]/[" & UserBas.stCounter.TrimCounter.ToString & "] * 100.0 < [" & stUserData.NgJudgeRate.ToString & "]")
                Return (False)
            End If

        Catch ex As Exception
            MsgBox("PlateNGJudgeByCounter() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Function
#End Region
    'V2.0.0.1③↑

#Region "'V2.1.0.0①②③ -------------2019/9/20 機能追加"
    'V2.1.0.0①↓
#Region "カット毎の抵抗値変化量判定機能"
    ''' <summary>
    ''' 変数定義
    ''' </summary>
    ''' <remarks></remarks>
    Private CutMeasureBefore As Double = Double.MinValue    ' カット前測定値
    Private CutMeasureAfter As Double = Double.MinValue     ' カット後測定値
    Private bVariationDone As Boolean = False               ' カット後の変化量計算済み=True,未計算=False
    Private bBeforeMeasureReadDone As Boolean = False       ' タクトアップの為にカット前測定値次のカットの初期測定値に使用'V2.1.0.0⑤
    Private iVariationCutNGCutNo As Integer = 0             ' カット毎の抵抗値変化量判定エラーカット番号、初期値０
    Private dSavedVariationRate As Double                   ' カット毎の抵抗値変化量保存用
    Private bCutVariationJudgeExecute As Boolean = False    ' カット毎の抵抗値変化量判定有り
    Private bVariationNGHiorLow As Boolean = True           ' カット毎の抵抗値変化量ＮＧ種別ＬＯ：True　ＨＩ：False'V2.1.0.0⑤
    Private bCutVariationCutDone As Boolean = False         ' カット毎の抵抗値変化量カット有りでTrue'V2.1.0.0⑤

    ''' <summary>
    ''' カット毎の抵抗値変化量判定機能初期データの設定
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CutVariationDataInitialize()
        Try
            Dim dBefore As Double
            Dim dAfter As Double
            Dim dTargetCoffJudge As Double

            For rn As Integer = 1 To stPLT.RCount
                dTargetCoffJudge = stUserData.dTargetCoffJudge(UserSub.GetResNumberInCircuit(rn))
                For CutNo As Short = 1 To stREG(rn).intTNN
                    If CutNo = 1 Then
                        dBefore = dTargetCoffJudge
                    Else
                        dBefore = dAfter
                    End If
                    dAfter = dBefore + stREG(rn).STCUT(CutNo).dblCOF

                    stREG(rn).STCUT(CutNo).iVariationRepeat = 0             ' リピート有無
                    stREG(rn).STCUT(CutNo).iVariation = 0                   ' 判定有無
                    If UserModule.IsCutResistor(stREG, rn) Then
                        stREG(rn).STCUT(CutNo).dRateOfUp = (dAfter - dBefore) / dTargetCoffJudge * 100      ' 上昇率
                    Else
                        stREG(rn).STCUT(CutNo).dRateOfUp = 0.0                                              ' 上昇率
                    End If
                    stREG(rn).STCUT(CutNo).dVariationLow = -1.0             ' 下限値
                    stREG(rn).STCUT(CutNo).dVariationHi = 1.0               ' 上限値
                Next
            Next
        Catch ex As Exception
            MsgBox("CutVariationDataInitialize() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' カット毎の抵抗値変化量判定のリピート有への全コピー処理
    ''' </summary>
    ''' <param name="stPlate">プレートデータ構造体</param>
    ''' <param name="stRegData">抵抗データ構造体</param>
    ''' <param name="ResNo">コピー元抵抗番号</param>
    ''' <param name="CutNo">コピー元カット番号</param>
    ''' <remarks></remarks>
    Public Sub CutVariationDataCopy(ByVal stPlate As PLATE_DATA, ByRef stRegData As Reg_Info(), ByVal ResNo As Integer, ByVal CutNo As Integer)
        Try
            Dim OrderNo As Integer

            OrderNo = UserSub.GetResNumberInCircuit(stRegData, ResNo)               ' サーキット内の抵抗の順番

            For rn As Integer = 1 To stPlate.RCount
                If rn = ResNo Then                                                  ' コピー元はスキップする。
                    Continue For
                End If
                If UserModule.IsCutResistor(stRegData, rn) Then
                    If OrderNo = UserSub.GetResNumberInCircuit(stRegData, rn) Then                                      ' サーキット内の同じ抵抗順番
                        stRegData(rn).STCUT(CutNo).iVariationRepeat = stRegData(ResNo).STCUT(CutNo).iVariationRepeat    ' リピート有無
                        stRegData(rn).STCUT(CutNo).iVariation = stRegData(ResNo).STCUT(CutNo).iVariation                ' 判定有無
                        stRegData(rn).STCUT(CutNo).dRateOfUp = stRegData(ResNo).STCUT(CutNo).dRateOfUp                  ' 上昇率
                        stRegData(rn).STCUT(CutNo).dVariationLow = stRegData(ResNo).STCUT(CutNo).dVariationLow          ' 下限値
                        stRegData(rn).STCUT(CutNo).dVariationHi = stRegData(ResNo).STCUT(CutNo).dVariationHi            ' 上限値
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox("CutVariationDataCopy() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' トリミングデータ単位での抵抗値変化量判定有無確認
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsCutVariationJudgeExecute() As Boolean
        Return (bCutVariationJudgeExecute)
    End Function

    ''' <summary>
    ''' トリミングデータ単位での抵抗値変化量判定有無チェック
    ''' </summary>
    ''' <remarks></remarks>
    Public Function CutVariationJudgeExecuteCheck() As Boolean
        Try
            bCutVariationJudgeExecute = False

            If (DGL = TRIM_MODE_ITTRFT) AndAlso IsSpecialTrimType() Then
                For ResNo As Integer = 1 To stPLT.RCount
                    If IsCutResistor(ResNo) Then
                        For CutNo As Integer = 1 To stREG(ResNo).intTNN
                            If stREG(ResNo).STCUT(CutNo).iVariation = 1 Then
                                bCutVariationJudgeExecute = True
                                Exit For
                            End If
                        Next
                    End If
                Next
            End If
            Return (bCutVariationJudgeExecute)
        Catch ex As Exception
            MsgBox("CutVariationJudgeExecuteCheck() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' 抵抗値変化量判定・判定用目標値算出係数からのカット前抵抗値の保存
    ''' </summary>
    ''' <param name="ResNo"></param>
    ''' <remarks>１抵抗カット開始前の初期化処理</remarks>
    Public Sub CutVariationInitialize(ByVal ResNo As Integer)
        Try

            iVariationCutNGCutNo = 0                ' カット毎の抵抗値変化量判定エラーカット番号初期化

            If IsCutVariationJudgeExecute() AndAlso IsCutResistor(ResNo) Then
                Call CutVariationDebugLogOut("抵抗値変化量判定初期化RES=[" & ResNo.ToString("0") & "]")
                bVariationDone = False              ' カット後の変化量計算済み=True,未計算=False
                ' 判定用目標値算出係数のカット前抵抗値保存
                CutVariationMeasureBeforeSet(UserSub.GetInitialResValue())
            End If

        Catch ex As Exception
            MsgBox("CutVariationInitialize() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try

    End Sub

    'V2.1.0.0⑤↓
    ''' <summary>
    ''' 抵抗値変化量判定・カット有りにする。
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CutVariationCutSet()
        bCutVariationCutDone = True
        Call CutVariationDebugLogOut("カット毎抵抗値変化量カット有りにセット")
    End Sub

    Private Function CutVariationCutDone() As Boolean
        Return (bCutVariationCutDone)
    End Function
    'V2.1.0.0⑤↑
    ''' <summary>
    ''' 抵抗値変化量判定・カット後の変化量未計算状態にする。
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CutVariationInitByCut()
        bVariationDone = False                          ' カット後の変化量計算済み=True,未計算=False
        CutMeasureAfter = Double.MinValue               'V2.1.0.0⑤ カット後測定値
        bCutVariationCutDone = False                   'V2.1.0.0⑤
        Call CutVariationDebugLogOut("カット毎抵抗値変化量カット有り初期化")
    End Sub

    ''' <summary>
    ''' 抵抗値変化量判定結果取得
    ''' </summary>
    ''' <returns>OK:False,NG:True</returns>
    ''' <remarks></remarks>
    Public Function CutVariationFinalJudgeNG() As Boolean
        If iVariationCutNGCutNo > 0 Then
            Return (True)
        Else
            Return (False)
        End If
    End Function

    ''' <summary>
    ''' 抵抗値変化量判定用カット前抵抗値を保存する。
    ''' </summary>
    ''' <param name="dMeasure">抵抗値</param>
    ''' <remarks></remarks>
    Private Sub CutVariationMeasureBeforeSet(ByVal dMeasure As Double)
        CutMeasureBefore = dMeasure
        If CutMeasureBefore = Double.MinValue Then
            Call CutVariationDebugLogOut("カット前抵抗値初期化")
        Else
            Call CutVariationDebugLogOut("カット前抵抗値保存値=[" & CutMeasureBefore.ToString & "]")
            bBeforeMeasureReadDone = True
        End If
    End Sub

    'V2.1.0.0⑤↓
    ''' <summary>
    ''' カット前抵抗値の取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CutVariationMeasureBeforeGet() As Double
        Return (CutMeasureBefore)
    End Function

    ''' <summary>
    ''' 抵抗値変化量判定用カット前抵抗値保存かの確認・一度読み出したらオフする。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsCutMeasureBefore() As Boolean
        If bBeforeMeasureReadDone Then
            bBeforeMeasureReadDone = False
            Return (True)
        Else
            Return (False)
        End If
    End Function
    'V2.1.0.0⑤↑

    ''' <summary>
    ''' 抵抗値変化量判定用カット後抵抗値保存
    ''' </summary>
    ''' <param name="dMeasure">抵抗値</param>
    ''' <remarks></remarks>
    Public Sub CutVariationMeasureAfterSet(ByVal dMeasure As Double)
        CutMeasureAfter = dMeasure
        If CutMeasureAfter = Double.MinValue Then
            Call CutVariationDebugLogOut("カット後抵抗値初期化")
        Else
            Call CutVariationDebugLogOut("カット後抵抗値保存値=[" & CutMeasureAfter.ToString & "]")
        End If
    End Sub

    'V2.1.0.0⑤↓
    ''' <summary>
    ''' 抵抗値変化量判定用カット後抵抗値未保存かの確認
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsNotCutMeasureAfter() As Boolean
        If CutMeasureAfter = Double.MinValue Then
            Return (True)
        Else
            Return (False)
        End If
    End Function

    ''' <summary>
    ''' カット毎の抵抗値変化量判定エラーカット番号初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub VariationCutNGCutNoReset()
        iVariationCutNGCutNo = 0
    End Sub
    'V2.1.0.0⑤↑

    ''' <summary>
    ''' カット毎の抵抗値変化量判定
    ''' </summary>
    ''' <param name="ResNo">抵抗番号</param>
    ''' <param name="CutNo">カット番号</param>
    ''' <returns>OK:True,NG:False</returns>
    ''' <remarks></remarks>
    Public Function CutVariationJudge(ByVal ResNo As Integer, ByVal CutNo As Integer) As Boolean
        Try
            Dim sRtn As Short

            If IsCutVariationJudgeExecute() AndAlso IsCutResistor(ResNo) AndAlso (bVariationDone = False) Then

                iVariationCutNGCutNo = 0

                If IsNotCutMeasureAfter() = True Then   ' カット後抵抗値未測定
                    If UserSub.IsSpecialTrimType Then
                        sRtn = V_R_MEAS(stREG(ResNo).intSLP, stREG(ResNo).intMType, CutMeasureAfter, ResNo, UserSub.GetTRV())
                    Else
                        sRtn = V_R_MEAS(stREG(ResNo).intSLP, stREG(ResNo).intMType, CutMeasureAfter, ResNo, stREG(ResNo).dblNOM)
                    End If
                    If sRtn = cFRS_NORMAL Then
                        Call CutVariationDebugLogOut("カット後抵抗値未測定時測定(CutVariationJudge) RES=[" & ResNo.ToString & "] CUT=[" & CutNo.ToString & "] 測定値=[" & CutMeasureAfter.ToString & "]")
                    Else
                        Call Z_PRINT("カット毎の抵抗値変化量判定 RES=[" & ResNo.ToString & "] CUT=[" & CutNo.ToString & "] 測定エラー=[" & sRtn.ToString & "]")
                        Call DebugLogOut("カット毎の抵抗値変化量判定(JudgeVariationByCut) RES=[" & ResNo.ToString & "] CUT=[" & CutNo.ToString & "] 測定エラー=[" & sRtn.ToString & "]")
                        Call CutVariationDebugLogOut("カット毎の抵抗値変化量判定(JudgeVariationByCut) RES=[" & ResNo.ToString & "] CUT=[" & CutNo.ToString & "] 測定エラー=[" & sRtn.ToString & "]")
                        iVariationCutNGCutNo = CutNo
                        Return (False)
                    End If

                End If

                ' 判定が無くても判定有りのカットの為に計算は必要。

                Dim dTargetCoffJudge As Double = UserSub.GetInitialResValue()
                '上昇率　＝　(カット後抵抗値　－カット前抵抗値)/判定用目標値算出係数　×　１００
                dSavedVariationRate = (CutMeasureAfter - CutVariationMeasureBeforeGet()) / dTargetCoffJudge * 100
                CutVariationDebugLogOut("上昇率[" & dSavedVariationRate.ToString & "] = (カット後抵抗値[" & CutMeasureAfter.ToString & "]-カット前抵抗値[" & CutVariationMeasureBeforeGet().ToString & "])/初期測定値[" & dTargetCoffJudge.ToString & "]*100")

                Dim dLo As Double = stREG(ResNo).STCUT(CutNo).dRateOfUp + stREG(ResNo).STCUT(CutNo).dVariationLow
                Dim dHi As Double = stREG(ResNo).STCUT(CutNo).dRateOfUp + stREG(ResNo).STCUT(CutNo).dVariationHi

                CutVariationMeasureBeforeSet(CutMeasureAfter)               ' 次のカットの為にカット前抵抗値に現在のカット後抵抗値を保存して入れ替える。
                CutVariationMeasureAfterSet(Double.MinValue)                ' 初期化状態にする。
                bVariationDone = True                                       ' カット後の変化量計算済み状態にする。

                ' カット毎の抵抗値変化量判定が有りの時で、上下限値から外れている時は、カット番号を保存してエラーリターンする。
                If CutVariationCutDone() AndAlso stREG(ResNo).STCUT(CutNo).iVariation = 1 AndAlso (dSavedVariationRate < dLo OrElse dHi < dSavedVariationRate) Then
                    DebugLogOut("上昇率判定NG 抵抗NO=[" & ResNo.ToString & "] カットNO=[" & CutNo.ToString & "] 下限[" & dLo.ToString & "] <= 上昇率[" & dSavedVariationRate.ToString & "] <= 上限[" & dHi.ToString & "]")
                    CutVariationDebugLogOut("上昇率判定NG 抵抗NO=[" & ResNo.ToString & "] カットNO=[" & CutNo.ToString & "] 下限[" & dLo.ToString & "] <= 上昇率[" & dSavedVariationRate.ToString & "] <= 上限[" & dHi.ToString & "]")
                    iVariationCutNGCutNo = CutNo
                    'V2.1.0.0⑤↓
                    If dSavedVariationRate < dLo Then
                        bVariationNGHiorLow = True
                    Else
                        bVariationNGHiorLow = False
                    End If
                    'V2.1.0.0⑤↑
                    Return (False)
                Else
                    If Not CutVariationCutDone() Then
                        CutVariationDebugLogOut("上昇率判定無(カット無し) 抵抗NO=[" & ResNo.ToString & "] カットNO=[" & CutNo.ToString & "] 下限[" & dLo.ToString & "] <= 上昇率[" & dSavedVariationRate.ToString & "] <= 上限[" & dHi.ToString & "]")
                    ElseIf stREG(ResNo).STCUT(CutNo).iVariation = 1 Then
                        CutVariationDebugLogOut("上昇率判定OK 抵抗NO=[" & ResNo.ToString & "] カットNO=[" & CutNo.ToString & "] 下限[" & dLo.ToString & "] <= 上昇率[" & dSavedVariationRate.ToString & "] <= 上限[" & dHi.ToString & "]")
                    Else
                        CutVariationDebugLogOut("上昇率判定無 抵抗NO=[" & ResNo.ToString & "] カットNO=[" & CutNo.ToString & "] 下限[" & dLo.ToString & "] <= 上昇率[" & dSavedVariationRate.ToString & "] <= 上限[" & dHi.ToString & "]")
                    End If
                    Return (True)
                End If
            End If
            Return (True)

        Catch ex As Exception
            MsgBox("CutVariationJudge() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (False)
        End Try

    End Function

    'V2.1.0.0⑤↓
    ''' <summary>
    ''' カット毎の抵抗値変化量判定ＮＧ時のHI,LO種別判断
    ''' </summary>
    ''' <returns>True:Lo,False：HI</returns>
    ''' <remarks></remarks>
    Public Function GetVariationNGHiorLow() As Boolean
        Return (bVariationNGHiorLow)
    End Function
    'V2.1.0.0⑤↑

    ''' <summary>
    ''' カット毎の抵抗値変化量判定・ＮＧ時のカット番号取得
    ''' </summary>
    ''' <returns>ＮＧ時のカット番号</returns>
    ''' <remarks></remarks>
    Public Function CutVariationCutNoGet() As Double
        Return (iVariationCutNGCutNo)
    End Function

    ''' <summary>
    ''' カット毎の抵抗値変化量判定・ＮＧ時の上昇率取得
    ''' </summary>
    ''' <returns>ＮＧ時の上昇率</returns>
    ''' <remarks></remarks>
    Public Function CutVariationRateGet() As Double
        Return (dSavedVariationRate)
    End Function
#End Region
    'V2.1.0.0①↑

    'V2.1.0.0②③↓
#Region "アッテネータテーブル、温度センサー情報テーブル関連"
    ''' <summary>
    ''' １番新しいファイルを取得する。
    ''' </summary>
    ''' <param name="sPath">フォルダ</param>
    ''' <param name="sHeader">ファイル名のヘッダ"",""</param>
    ''' <param name="sGetFileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetNewestFile(ByVal sPath As String, ByVal sHeader As String, ByRef sGetFileName As String) As Boolean
        Try
            Dim sFileName As String
            Dim sExtension As String
            Dim sFileList As String() = System.IO.Directory.GetFiles(sPath)
            Dim Year As Integer, Month As Integer, Day As Integer
            Dim Today As DateTime = System.DateTime.Today

            If sFileList.Length = 0 Then
                Call Z_PRINT("フォルダ[" & sPath & "]にファイルが存在しません。")
                Return (False)
            End If

            Array.Sort(sFileList)
            sGetFileName = sFileList(sFileList.Length - 1)
            ' 妥当性チェック
            For i As Integer = sFileList.Length - 1 To 0 Step -1
                sExtension = System.IO.Path.GetExtension(sFileList(i))                                                      ' 拡張子取得
                If sExtension.Equals(".CSV", StringComparison.OrdinalIgnoreCase) Then                                       ' 拡張子が一致するファイルが対象
                    sFileName = System.IO.Path.GetFileNameWithoutExtension(sFileList(i))                                    ' 拡張子を除くファイル名
                    If sFileName.Length = (sHeader.Length + 8) Then                                                         ' YYYYMMDDの８文字
                        If sFileName.Substring(0, sHeader.Length).Equals(sHeader, StringComparison.OrdinalIgnoreCase) Then  ' ファイル名のタイトル部が一致する
                            Year = Integer.Parse(sFileName.Substring(sHeader.Length, 4))                                    ' 年
                            Month = Integer.Parse(sFileName.Substring(sHeader.Length + 4, 2))                               ' 月
                            Day = Integer.Parse(sFileName.Substring(sHeader.Length + 6, 2))                                 ' 日
                            Dim FileDate As New DateTime(Year, Month, Day)                                                  ' ファイルの年月日をDateTime型に変換
                            If FileDate.Date.CompareTo(Today.Date) <= 0 Then                                                ' 今日までの日付を対象とする。
                                sGetFileName = sFileList(i)                                                                 ' もし、未来の日付が有ってその日になると対象に入ってしまう。
                                Return (True)
                            End If
                        End If
                    End If
                End If
            Next
            Return (False)
        Catch ex As Exception
            MsgBox("GetNewestFile() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (False)
        End Try
    End Function
#End Region
    'V2.1.0.0②③↑

    'V2.1.0.0②↓
#Region "レーザパワーモニタリング"
    ''' <summary>
    ''' レーザパワーモニタリングモード管理変数
    ''' </summary>
    ''' <remarks></remarks>
    Public Const POWER_CHECK_NONE As Short = 0
    Public Const POWER_CHECK_START As Short = 1
    Public Const POWER_CHECK_LOT As Short = 2
    Private gbLaserCaribrarionUse As Boolean = False
    Private bLaserCalibrationExecute As Boolean = False
    Private giLaserCalibrationMode As Integer = 0

    ''' <summary>
    ''' レーザパワーモニタリング使用有り無し
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsLaserCaribrarionUse() As Boolean
        Return (gbLaserCaribrarionUse)
    End Function

    ''' <summary>
    ''' レーザパワーモニタリングモード取得
    ''' </summary>
    ''' <returns>0:POWER CHECK なし,1:POWER CHECK 自動運転開始時,2:POWER CHECK ロット毎</returns>
    ''' <remarks></remarks>
    Public Function LaserCalibrationModeGet() As Integer
        Return (giLaserCalibrationMode)
    End Function

    ''' <summary>
    ''' レーザパワーモニタリング実行モード設定
    ''' </summary>
    ''' <param name="Mode">POWER_CHECK_NONE,POWER_CHECK_START,POWER_CHECK_LOT</param>
    ''' <remarks></remarks>
    Public Sub LaserCalibrationModeSet(ByVal Mode As Integer)
        Try
            If IsLaserCaribrarionUse() Then
                If Mode = POWER_CHECK_NONE Then                 ' なしに変更したら実行フラグも無しにする。
                    bLaserCalibrationExecute = False
                End If
                If giLaserCalibrationMode <> POWER_CHECK_LOT AndAlso Mode = POWER_CHECK_LOT Then
                    If pbLoadFlg = True Then
                        bLaserCalibrationExecute = True         ' ロット毎に変更したらロード済みなら実行有りにする。
                    End If
                End If
                giLaserCalibrationMode = Mode
                WritePrivateProfileString("LASER", "LASER_CALIBRATION_MODE", giLaserCalibrationMode.ToString("0"), USER_SYSPARAMPATH)
            End If
        Catch ex As Exception
            MsgBox("LaserCalibrationModeSet() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Sub
    Public Sub LaserCalibrationModeUpdate()
        Try
            Select Case (giLaserCalibrationMode)
                Case POWER_CHECK_NONE
                    Form1.ButtonLaserCalibration.Text = "POWER CHECK なし"
                    Form1.ButtonLaserCalibration.BackColor = SystemColors.Control
                Case POWER_CHECK_START
                    Form1.ButtonLaserCalibration.Text = "POWER CHECK 自動運転開始時"
                    Form1.ButtonLaserCalibration.BackColor = System.Drawing.Color.LightSkyBlue
                Case POWER_CHECK_LOT
                    Form1.ButtonLaserCalibration.Text = "POWER CHECK ロット毎"
                    Form1.ButtonLaserCalibration.BackColor = System.Drawing.Color.LightPink
            End Select

            Form1.ButtonLaserCalibration.Refresh()

        Catch ex As Exception
            MsgBox("LaserCalibrationModeUpdate() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' ユーザプログラム起動時レーザパワーモニタリングモード設定
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub LaserCalibrationModeLoad()
        Try

            If gSysPrm.stIOC.giPM_Tp = 1 Then   ' パワーメータの装着タイプ(0:なし(手置き), 1:ステージ設置タイプ, 2:ステージ外設置タイプ)
                gbLaserCaribrarionUse = True

                Dim LaserCalibrationMode As Integer = Integer.Parse(GetPrivateProfileString_S("LASER", "LASER_CALIBRATION_MODE", USER_SYSPARAMPATH, "0"))

                UserSub.LaserCalibrationModeSet(LaserCalibrationMode)

                UserSub.LaserCalibrationModeUpdate()
            Else
                gbLaserCaribrarionUse = False

                Form1.cmdLaserCalibration.Enabled = False
                Form1.cmdLaserCalibration.Visible = False

                Form1.ButtonLaserCalibration.Enabled = False
                Form1.ButtonLaserCalibration.Visible = False

                Form1.cmdLaserTeach.Width = Form1.CmdTx.Width
            End If

        Catch ex As Exception
            MsgBox("LaserCalibrationModeLoad() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' レーザパワーモニタリング実行有無設定
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub LaserCalibrationSet(ByVal Mode As Integer)
        Try
            If IsLaserCaribrarionUse() Then
                Select Case (Mode)
                    Case POWER_CHECK_NONE
                        bLaserCalibrationExecute = False
                    Case POWER_CHECK_START
                        If LaserCalibrationModeGet() = POWER_CHECK_START Then
                            bLaserCalibrationExecute = True
                        End If
                    Case POWER_CHECK_LOT
                        If LaserCalibrationModeGet() = POWER_CHECK_LOT Then
                            bLaserCalibrationExecute = True
                        End If
                End Select
            Else
                bLaserCalibrationExecute = False
            End If
        Catch ex As Exception
            MsgBox("LaserCalibrationSet() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' レーザパワーモニタリング実行有無
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>ファンクション内で呼ばれたら実行無しに変更する</remarks>
    Public Function LaserCalibrationExecute() As Boolean
        Try
            Dim bRtn As Boolean = False

            If IsLaserCaribrarionUse() Then
                If DGL = TRIM_MODE_ITTRFT OrElse DGL = TRIM_MODE_CUT OrElse DGL = TRIM_MODE_MEAS_MARK OrElse DGL = TRIM_VARIATION_MEAS Then
                    bRtn = bLaserCalibrationExecute
                    bLaserCalibrationExecute = False
                End If
            End If

            Return (bRtn)
        Catch ex As Exception
            MsgBox("LaserCalibrationExecute() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try

    End Function

    ''' <summary>
    ''' アッテネータテーブルからデータ取得
    ''' </summary>
    ''' <param name="No">番号</param>
    ''' <param name="stData">データ格納構造体</param>
    ''' <param name="MaxNo">番号に99以上を設定した時最大番号</param>
    ''' <returns>True:検索番号に一致した時のみ</returns>
    ''' <remarks></remarks>
    Private Function LaserCalibrationAttenuatorTableGet(ByVal No As Integer, ByRef stData As stATTENUATOR_TABLE, ByRef MaxNo As Integer) As Boolean
        Try
            Dim sFolder As String = vbNullString
            Dim sData As String
            Dim mData() As String
            Dim TableNo As Integer
            Dim bHeader As Boolean = True
            Dim dData As Double

            MaxNo = 0

            If Not GetNewestFile(cLASERPOWER_PATH, cLASERPOWER_HEADER, sFolder) Then
                Call Z_PRINT("アッテネータテーブルファイルを取得出来ませんでした。")
                Return (False)
            End If

            If IO.File.Exists(sFolder) = True Then  ' ファイルが有る。
                Using sr As New System.IO.StreamReader(sFolder, System.Text.Encoding.GetEncoding("Shift-JIS"))
                    Do While Not sr.EndOfStream
                        sData = sr.ReadLine
                        If bHeader Then                                     ' １行目はタイトル行
                            bHeader = False
                            Continue Do
                        End If
                        mData = sData.Split(",")                            ' 文字列を','で分割して取出す
                        TableNo = Integer.Parse(mData(0))

                        If TableNo > MaxNo Then
                            MaxNo = TableNo
                        End If

                        If TableNo = No Then                                            ' 番号一致
                            stData.No = TableNo
                            stData.Power = mData(1).Trim                                ' パワー設定
                            If Not Double.TryParse(stData.Power, dData) Then
                                Call Z_PRINT("[" & No.ToString & "]番目[" & stData.Power & "]パワー設定値が数値に変換できません。")
                                Return (False)
                            End If
                            stData.PowerUnit = mData(2).Trim                            ' パワー単位
                            stData.Limit = mData(3).Trim                                ' 範囲
                            If Not Double.TryParse(stData.Limit, dData) Then
                                Call Z_PRINT("[" & No.ToString & "]番目[" & stData.Limit & "]パワー設定範囲が数値に変換できません。")
                                Return (False)
                            End If
                            stData.LimitUnit = mData(4).Trim                            ' 範囲単位

                            If Not Double.TryParse(mData(5), stData.Rate) Then          ' 減衰率
                                Call Z_PRINT("[" & No.ToString & "]番目[" & mData(5) & "]減衰率が数値に変換できません。")
                                Return (False)
                            End If
                            stData.RateUnit = mData(6).Trim                             ' 減衰率単位

                            If Not Integer.TryParse(mData(7), stData.Rotation) Then     ' 回転量
                                Call Z_PRINT("[" & No.ToString & "]番目[" & mData(7) & "]回転量が数値に変換できません。")
                                Return (False)
                            End If

                            If Not Integer.TryParse(mData(8), stData.FixAtt) Then     ' 固定アッテネータ
                                Call Z_PRINT("[" & No.ToString & "]番目[" & mData(8) & "]固定アッテネータが数値に変換できません。")
                                Return (False)
                            End If

                            stData.Comment = ""
                            For i As Integer = 9 To mData.Length - 1
                                stData.Comment = stData.Comment & "," & mData(i)                   ' コメント
                            Next
                            Return (True)
                        End If
                    Loop
                End Using
            End If

            Return (False)

        Catch ex As Exception
            MsgBox("LaserCalibrationAttenuatorTableGet() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (False)
        End Try
    End Function

    ''' <summary>
    ''' アッテネータテーブル内の最大番号を取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LaserCalibrationMaxNumberGet() As Integer
        Try

            Dim stAttenuatorTable As stATTENUATOR_TABLE
            Dim MaxNo As Integer

            stAttenuatorTable.No = -1           ' 番号
            stAttenuatorTable.Power = ""        ' パワー設定
            stAttenuatorTable.PowerUnit = ""    ' パワー単位
            stAttenuatorTable.Limit = ""        ' 範囲
            stAttenuatorTable.LimitUnit = ""    ' 範囲単位
            stAttenuatorTable.Rate = -1.0       ' 減衰率
            stAttenuatorTable.RateUnit = ""     ' 減衰率単位
            stAttenuatorTable.Rotation = -1     ' 回転量
            stAttenuatorTable.FixAtt = -1       ' 固定アッテネータ
            stAttenuatorTable.Comment = ""      ' コメント

            Call LaserCalibrationAttenuatorTableGet(MAX_ATTENUATOR + 1, stAttenuatorTable, MaxNo)

            If MaxNo > MAX_ATTENUATOR Then
                MaxNo = MAX_ATTENUATOR
            End If

            Return (MaxNo)

        Catch ex As Exception
            MsgBox("LaserCalibrationMaxNumberGetByAttenuatorTable() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (0)
        End Try
    End Function

    ''' <summary>
    ''' フルパワー値の取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LaserCalibrationFullPowerGet(ByRef FullPowerTarget As Double, ByRef FullPowerLimit As Double) As Boolean
        Try
            Dim stAttenuatorTable As stATTENUATOR_TABLE
            Dim MaxNo As Integer

            stAttenuatorTable.No = -1           ' 番号
            stAttenuatorTable.Power = ""        ' パワー設定
            stAttenuatorTable.PowerUnit = ""    ' パワー単位
            stAttenuatorTable.Limit = ""        ' 範囲
            stAttenuatorTable.LimitUnit = ""    ' 範囲単位
            stAttenuatorTable.Rate = -1.0       ' 減衰率
            stAttenuatorTable.RateUnit = ""     ' 減衰率単位
            stAttenuatorTable.Rotation = -1     ' 回転量
            stAttenuatorTable.FixAtt = -1       ' 固定アッテネータ
            stAttenuatorTable.Comment = ""      ' コメント

            If LaserCalibrationAttenuatorTableGet(0, stAttenuatorTable, MaxNo) Then

                If Not Double.TryParse(stAttenuatorTable.Power.Trim, FullPowerTarget) Then
                    Call Z_PRINT("[0]番目フルパワー[" & stAttenuatorTable.Power & "]パワー設定値が数値に変換できません。")
                    Return (False)
                Else
                    If stAttenuatorTable.PowerUnit.IndexOf("mW") >= 0 OrElse stAttenuatorTable.PowerUnit.IndexOf("ｍＷ") >= 0 Then
                        FullPowerTarget = FullPowerTarget / 1000.0
                    End If
                End If

                If Not Double.TryParse(stAttenuatorTable.Limit.Trim, FullPowerLimit) Then
                    Call Z_PRINT("[[0]番目[" & stAttenuatorTable.Limit & "]パワー設定範囲が数値に変換できません。")
                    Return (False)
                Else
                    If stAttenuatorTable.LimitUnit.IndexOf("mW") >= 0 OrElse stAttenuatorTable.LimitUnit.IndexOf("ｍＷ") >= 0 Then
                        FullPowerLimit = FullPowerLimit / 1000.0
                    End If
                End If

                Return (True)
            Else
                Return (False)
            End If

        Catch ex As Exception
            MsgBox("LaserCalibrationFullPowerGet() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (False)
        End Try
    End Function

    ''' <summary>
    ''' アッテネータテーブルから設定データの取得
    ''' </summary>
    ''' <param name="No">番号</param>
    ''' <param name="dblRotPar">減衰率</param>
    ''' <param name="iFixAtt">固定アッテネータ</param>
    ''' <param name="dblRotAtt">回転量</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LaserCalibrationAttenuatorDataGet(ByVal No As Integer, ByRef dblRotPar As Double, ByRef iFixAtt As Double, ByRef dblRotAtt As Double) As Boolean
        Try
            Dim stAttenuatorTable As stATTENUATOR_TABLE
            Dim MaxNo As Integer

            If No < 0 Or MAX_ATTENUATOR < No Then
                Return (False)
            End If

            stAttenuatorTable.No = -1           ' 番号
            stAttenuatorTable.Power = ""        ' パワー設定
            stAttenuatorTable.PowerUnit = ""    ' パワー単位
            stAttenuatorTable.Limit = ""        ' 範囲
            stAttenuatorTable.LimitUnit = ""    ' 範囲単位
            stAttenuatorTable.Rate = -1.0       ' 減衰率
            stAttenuatorTable.RateUnit = ""     ' 減衰率単位
            stAttenuatorTable.Rotation = -1     ' 回転量
            stAttenuatorTable.FixAtt = -1       ' 固定アッテネータ
            stAttenuatorTable.Comment = ""      ' コメント

            If LaserCalibrationAttenuatorTableGet(No, stAttenuatorTable, MaxNo) Then
                dblRotPar = stAttenuatorTable.Rate
                iFixAtt = stAttenuatorTable.FixAtt
                dblRotAtt = stAttenuatorTable.Rotation
                Return (True)
            Else
                Return (False)
            End If

        Catch ex As Exception
            MsgBox("LaserCalibrationAttenuatorDataGet() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (False)
        End Try
    End Function

    ''' <summary>
    '''  アッテネータテーブルから全データ取得
    ''' </summary>
    ''' <param name="stData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LaserCalibrationAllDataGet(ByRef MaxNo As Integer, ByRef stData() As stATTENUATOR_TABLE) As Boolean
        Try
            Dim bRtn As Boolean = False
            Dim iTemp As Integer

            MaxNo = LaserCalibrationMaxNumberGet()

            If MaxNo <= 0 Then
                Return (False)
            End If

            For No As Integer = 0 To MaxNo
                If LaserCalibrationAttenuatorTableGet(No, stData(No), iTemp) Then
                    bRtn = True
                Else
                    Z_PRINT("アッテネータテーブルからの情報取得がエラーになりました。NO=[" & No.ToString & "]")
                    Return (False)
                End If
            Next

            Return (bRtn)

        Catch ex As Exception
            MsgBox("LaserCalibrationAllDataGet() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (False)
        End Try
    End Function

    Public Function LaserCalibrationAllDataWrite(ByRef stData() As stATTENUATOR_TABLE) As Boolean
        Try
            Dim sFolder As String = vbNullString
            Dim sHeaderData As String = vbNullString
            Dim sFileName As String

            If Not GetNewestFile(cLASERPOWER_PATH, cLASERPOWER_HEADER, sFolder) Then
                Call Z_PRINT("アッテネータテーブルファイルを取得出来ませんでした。")
                Return (False)
            End If

            Dim MaxNo As Integer = LaserCalibrationMaxNumberGet()

            If IO.File.Exists(sFolder) = True Then  ' ファイルが有る。
                Using sr As New System.IO.StreamReader(sFolder, System.Text.Encoding.GetEncoding("Shift-JIS"))
                    Do While Not sr.EndOfStream
                        sHeaderData = sr.ReadLine
                        Exit Do
                    Loop
                End Using
            End If

            sFileName = cLASERPOWER_PATH & cLASERPOWER_HEADER & DateTime.Now().ToString("yyyyMMdd") & ".CSV"

            Using WSR As New System.IO.StreamWriter(sFileName, False, System.Text.Encoding.GetEncoding("Shift-JIS"))  ' 第２引数 上書きは、False
                WSR.WriteLine(sHeaderData)                          ' ヘッダ出力

                'Public Structure stATTENUATOR_TABLE         ' アッテネータテーブル
                '            Dim No As Integer                       ' 番号
                '            Dim Power As String                     ' パワー設定
                '            Dim PowerUnit As String                 ' パワー単位
                '            Dim Limit As String                     ' 範囲
                '            Dim LimitUnit As String                 ' 範囲単位
                '            Dim Rate As Double                      ' 減衰率
                '            Dim RateUnit As String                  ' 減衰率単位
                '            Dim Rotation As Integer                 ' 回転量
                '            Dim FixAtt As Integer                   ' 固定アッテネータ
                '            Dim Comment As String                   ' コメント
                'End Structure

                For No As Integer = 0 To MaxNo
                    WSR.WriteLine(stData(No).No.ToString("0") & "," & stData(No).Power & "," & stData(No).PowerUnit & "," & stData(No).Limit & "," & stData(No).LimitUnit & "," & stData(No).Rate.ToString("0.00") & "," & stData(No).RateUnit & "," & stData(No).Rotation.ToString("0") & "," & stData(No).FixAtt.ToString("0") & stData(No).Comment)
                Next
            End Using

            Return (True)

        Catch ex As Exception
            MsgBox("LaserCalibrationAllDataWrite() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (False)
        End Try
    End Function

#End Region             'レーザパワーモニタリング
    'V2.1.0.0②↑

    'V2.1.0.0③↓
#Region "温度センサー情報テーブル"
    ''' <param name="stData">取得データ構造体</param>
    ''' <param name="MaxNo">情報テーブル内最大番号</param>
    ''' <returns>True:検索番号に一致した時のみ</returns>
    ''' <remarks></remarks>
    Private Function TemperatureTableGet(ByVal No As Integer, ByRef stData As stTEMPERATURE_TABLE, ByRef MaxNo As Integer) As Boolean
        Try
            Dim sFolder As String = vbNullString
            Dim sData As String
            Dim mData() As String
            Dim TableNo As Integer
            Dim bHeader As Boolean = True
            Dim dData As Double

            MaxNo = 0

            If Not GetNewestFile(cTEMPERATURE_PATH, cTEMPERATURE_HEADER, sFolder) Then
                Call Z_PRINT("温度センサー情報テーブルファイルを取得出来ませんでした。")
                Return (False)
            End If

            If IO.File.Exists(sFolder) = True Then  ' ファイルが有る。
                Using sr As New System.IO.StreamReader(sFolder, System.Text.Encoding.GetEncoding("Shift-JIS"))
                    Do While Not sr.EndOfStream
                        sData = sr.ReadLine
                        If bHeader Then                                     ' １行目はタイトル行
                            bHeader = False
                            Continue Do
                        End If
                        mData = sData.Split(",")                            ' 文字列を','で分割して取出す
                        TableNo = Integer.Parse(mData(0))

                        If TableNo > MaxNo Then
                            MaxNo = TableNo
                        End If

                        If TableNo = No Then                                ' 番号一致
                            stData.No = TableNo                             ' 番号
                            stData.Title = mData(1)                         ' 元素記号
                            If mData(2) = vbNullString Or mData(3) = vbNullString Or mData(4) = vbNullString Then
                                Call Z_PRINT("[" & No.ToString & "]番目でデータが存在しない項目があります。")
                                Return (False)
                            End If

                            If Double.TryParse(mData(2), dData) Then
                                stData.dTemperatura0 = Double.Parse(mData(2))   ' ０℃
                            Else
                                Call Z_PRINT("[" & No.ToString & "]番目０℃データ[" & mData(2) & "]が数値に変換できません。")
                                Return (False)
                            End If

                            If Double.TryParse(mData(3), dData) Then
                                stData.dDaihyouAlpha = Double.Parse(mData(3))   ' 代表α値
                            Else
                                Call Z_PRINT("[" & No.ToString & "]番目α値データ[" & mData(3) & "]が数値に変換できません。")
                                Return (False)
                            End If


                            If Double.TryParse(mData(4), dData) Then
                                stData.dDaihyouBeta = Double.Parse(mData(4))    ' 代表β値
                            Else
                                Call Z_PRINT("[" & No.ToString & "]番目β値データ[" & mData(4) & "]が数値に変換できません。")
                                Return (False)
                            End If

                            For i As Integer = 5 To mData.Length - 1
                                stData.Comment = "," & mData(i)                   ' コメント
                            Next
                            Return (True)
                        End If
                    Loop
                End Using
            End If

            Return (False)

        Catch ex As Exception
            MsgBox("TemperatureTableGet() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (False)
        End Try
    End Function

    ''' <summary>
    ''' 温度センサー情報テーブル内の最大番号を取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function TemperatureTableMaxNumberGet() As Integer
        Try

            Dim stTemperatureTable As stTEMPERATURE_TABLE
            Dim MaxNo As Integer

            stTemperatureTable.No = -1                       ' 番号
            stTemperatureTable.Title = ""                    ' 元素記号
            stTemperatureTable.dTemperatura0 = -1.0             ' ０℃
            stTemperatureTable.dDaihyouAlpha = -1.0             ' 代表α値
            stTemperatureTable.dDaihyouBeta = -1.0              ' 代表β値
            stTemperatureTable.Comment = ""                    ' コメント


            Call TemperatureTableGet(MAX_TEMPERATURE + 1, stTemperatureTable, MaxNo)

            If MaxNo > MAX_TEMPERATURE Then
                MaxNo = MAX_TEMPERATURE
            End If

            Return (MaxNo)

        Catch ex As Exception
            MsgBox("TemperatureTableMaxNumberGet() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (0)
        End Try
    End Function

    ''' <summary>
    ''' 温度センサー情報テーブルから情報取得
    ''' </summary>
    ''' <param name="No">検索対象番号</param>
    ''' <param name="dTemperatura0">０℃</param>
    ''' <param name="dDaihyouAlpha">代表α値</param>
    ''' <param name="dDaihyouBeta">代表β値</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function TemperatureTableDataGet(ByVal No As Integer, ByRef dTemperatura0 As Double, ByRef dDaihyouAlpha As Double, ByRef dDaihyouBeta As Double, Optional bLimitCheck As Boolean = False) As Boolean
        Try
            Dim stTemperatureTable As stTEMPERATURE_TABLE
            Dim MaxNo As Integer
            Dim Min As Double, Max As Double

            If No < 1 Or MAX_TEMPERATURE < No Then
                Return (False)
            End If

            stTemperatureTable.No = -1                      ' 番号
            stTemperatureTable.Title = ""                   ' 元素記号
            stTemperatureTable.dTemperatura0 = -1.0         ' ０℃
            stTemperatureTable.dDaihyouAlpha = -1.0         ' 代表α値
            stTemperatureTable.dDaihyouBeta = -1.0          ' 代表β値
            stTemperatureTable.Comment = ""                 ' コメント


            If TemperatureTableGet(No, stTemperatureTable, MaxNo) Then
                dTemperatura0 = stTemperatureTable.dTemperatura0    ' ０℃
                dDaihyouAlpha = stTemperatureTable.dDaihyouAlpha    ' 代表α値
                dDaihyouBeta = stTemperatureTable.dDaihyouBeta      ' 代表β値
                ' 範囲チェック
                If bLimitCheck Then
                    ' ０℃
                    Min = Double.Parse(GetPrivateProfileString_S("USER_VALUE", "011_MIN", cEDITDEF_FNAME, "0.0000001"))
                    Max = Double.Parse(GetPrivateProfileString_S("USER_VALUE", "011_MAX", cEDITDEF_FNAME, "100000000.0000000"))
                    If dTemperatura0 < Min Or Max < dTemperatura0 Then
                        Z_PRINT("温度センサー情報テーブルNo.=[" & No.ToString("0") & "]０℃上下限値エラー=[" & dTemperatura0.ToString("0.00000000") & "]")
                        Return (False)
                    End If
                    ' α値(ppm/℃)
                    Min = Double.Parse(GetPrivateProfileString_S("USER_VALUE", "012_MIN", cEDITDEF_FNAME, "-9999.0000000"))
                    Max = Double.Parse(GetPrivateProfileString_S("USER_VALUE", "012_MAX", cEDITDEF_FNAME, "9999.0000000"))
                    If dDaihyouAlpha < Min Or Max < dDaihyouAlpha Then
                        Z_PRINT("温度センサー情報テーブルNo.=[" & No.ToString("0") & "]α値上下限値エラー=[" & dDaihyouAlpha.ToString("0.00000000") & "]")
                        Return (False)
                    End If
                    ' β値(ppm/℃)
                    Min = Double.Parse(GetPrivateProfileString_S("USER_VALUE", "013_MIN", cEDITDEF_FNAME, "-9999.0000000"))
                    Max = Double.Parse(GetPrivateProfileString_S("USER_VALUE", "013_MAX", cEDITDEF_FNAME, "9999.0000000"))
                    If dDaihyouBeta < Min Or Max < dDaihyouBeta Then
                        Z_PRINT("温度センサー情報テーブルNo.=[" & No.ToString("0") & "]β値上下限値エラー=[" & dDaihyouBeta.ToString("0.00000000") & "]")
                        Return (False)
                    End If
                End If

                Return (True)
            Else
                Return (False)
            End If

        Catch ex As Exception
            MsgBox("TemperatureTableDataGet() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (False)
        End Try
    End Function
#End Region             ' 温度センサー情報テーブル
    'V2.1.0.0③↑
#End Region         'V2.1.0.0①②③

#Region "プローブマスターデータを読込み指定Noのプローブデータを設定する"    'V2.2.0.0⑮
    ''' <summary>
    ''' プローブデータを設定する 
    ''' </summary>
    ''' <returns></returns>
    Public Function ConvProbeData(ByVal ProbNo As Integer) As Integer
        Dim Ret As Boolean
        Dim stlocalProbeData As stPROBEDATA_TABLE
        Dim MaxNo As Integer

        Try
            ConvProbeData = cFRS_NORMAL

            If ProbNo = 0 Then
                Exit Function
            End If

            stlocalProbeData.No = 0
            stlocalProbeData.ProbeOn = 0.0
            stlocalProbeData.ProbeOff = 0.0                 'V2.2.0.0⑳
            stlocalProbeData.dTableOffsetX = 0.0
            stlocalProbeData.dTableOffsetY = 0.0
            stlocalProbeData.dBPOffsetX = 0.0
            stlocalProbeData.dBPOffsetY = 0.0
            stlocalProbeData.Comment = ""

            'θ補正関係はデータがない場合もあるので初期化はしない      'V2.2.1.6②
            'V2.2.1.6②↓
            stlocalProbeData.iPP30 = stThta.iPP30                   ' 位置補正モード：
            stlocalProbeData.iPP31 = stThta.iPP31                   ' 位置補正方法：
            stlocalProbeData.fpp34_x = stThta.fpp34_x               ' 補正ポジションオフセットX：
            stlocalProbeData.fpp34_y = stThta.fpp34_y               ' 補正ポジションオフセットY：
            stlocalProbeData.fTheta = stThta.fTheta                 ' 角度
            stlocalProbeData.iPP38 = stThta.iPP38                   ' グループ番号：
            stlocalProbeData.iPP37_1 = stThta.iPP37_1               ' パターン番号1：
            stlocalProbeData.fpp32_x = stThta.fpp32_x               ' パターン座標1X：
            stlocalProbeData.fpp32_y = stThta.fpp32_y               ' パターン座標1Y：
            stlocalProbeData.iPP37_2 = stThta.iPP37_2               ' パターン番号2：
            stlocalProbeData.fpp33_x = stThta.fpp33_x               ' パターン座標2X：
            stlocalProbeData.fpp33_y = stThta.fpp33_y               ' パターン座標2Y：
            'V2.2.1.6②↑

            ' PROBEDATAファイルを読込んで、その内容に更新する 
            Ret = ReadProbeCsv(ProbNo, stlocalProbeData, MaxNo)
            If Ret = True Then

                'V2.2.1.6②↓
                Z_PRINT("プローブテーブルからNO=[" & ProbNo.ToString() & "]の情報を取得しました。")
                Z_PRINT("　プローブON位置=[" & stlocalProbeData.ProbeOn.ToString("#0.000#") & "] ")
                Z_PRINT("　プローブOFF位置=[" & stlocalProbeData.ProbeOff.ToString("#0.000#") & "] ")          'V2.2.0.0⑳
                Z_PRINT("　テーブル位置オフセットX=[" & stlocalProbeData.dTableOffsetX.ToString("#0.000000#") & "] ")
                Z_PRINT("　テーブル位置オフセットY=[" & stlocalProbeData.dTableOffsetY.ToString("#0.000000#") & "] ")
                Z_PRINT("　BPオフセットX=[" & stlocalProbeData.dBPOffsetX.ToString("#0.000#") & "] ")
                Z_PRINT("　BPオフセットX=[" & stlocalProbeData.dBPOffsetY.ToString("#0.000#") & "] ")

                'V2.2.1.6②↓
                Z_PRINT("  位置補正モード=[" & stlocalProbeData.iPP30.ToString() & "]の情報を取得しました。")
                Z_PRINT("  位置補正方法=[" & stlocalProbeData.iPP31.ToString() & "]の情報を取得しました。")
                Z_PRINT("　補正ポジションオフセットX=[" & stlocalProbeData.fpp34_x.ToString("#0.000#") & "] ")
                Z_PRINT("　補正ポジションオフセットY=[" & stlocalProbeData.fpp34_y.ToString("#0.000#") & "] ")
                Z_PRINT("　角度=[" & stlocalProbeData.fTheta.ToString("#0.000#") & "] ")
                Z_PRINT("　グループ番号=[" & stlocalProbeData.iPP38.ToString() & "] ")
                Z_PRINT("　パターン番号1=[" & stlocalProbeData.iPP37_1.ToString() & "] ")
                Z_PRINT("　パターン座標1X=[" & stlocalProbeData.fpp32_x.ToString("#0.000#") & "] ")
                Z_PRINT("　パターン座標1Y=[" & stlocalProbeData.fpp32_y.ToString("#0.000#") & "] ")
                Z_PRINT("　パターン番号2=[" & stlocalProbeData.iPP37_2.ToString() & "] ")
                Z_PRINT("　パターン座標2X=[" & stlocalProbeData.fpp33_x.ToString("#0.000#") & "] ")
                Z_PRINT("　パターン座標2Y=[" & stlocalProbeData.fpp33_y.ToString("#0.000#") & "] ")
                'V2.2.1.6②↑

                Z_PRINT("　コメント=[" & stlocalProbeData.Comment & "] ")
            Else
                Z_PRINT("プローブテーブルからNO=[" & ProbNo.ToString & "]の情報を取得時にエラーが発生しました。")
                ConvProbeData = cFRS_FIOERR_INP

            End If
            stPLT.z_xoff = stlocalProbeData.dTableOffsetX       ' テーブル位置オフセット　トリムポジションオフセットX(mm)
            stPLT.z_yoff = stlocalProbeData.dTableOffsetY       ' テーブル位置オフセット　トリムポジションオフセットY(mm)
            stPLT.BPOX = stlocalProbeData.dBPOffsetX            ' ビーム位置オフセット　BP Offset X(mm)
            stPLT.BPOY = stlocalProbeData.dBPOffsetY            ' ビーム位置オフセット　BP Offset Y(mm)
            stPLT.Z_ZON = stlocalProbeData.ProbeOn              ' Z PROBE ON 位置(mm)
            stPLT.Z_ZOFF = stlocalProbeData.ProbeOff            ' Z PROBE OFF 位置(mm)            ' V2.2.0.0⑳
            'V2.2.1.6②↓
            stThta.iPP30 = stlocalProbeData.iPP30               ' 位置補正モード：
            stThta.iPP31 = stlocalProbeData.iPP31               ' 位置補正方法：
            stThta.fpp34_x = stlocalProbeData.fpp34_x           ' 補正ポジションオフセットX：
            stThta.fpp34_y = stlocalProbeData.fpp34_y           ' 補正ポジションオフセットY：
            stThta.fTheta = stlocalProbeData.fTheta             ' 角度
            stThta.iPP38 = stlocalProbeData.iPP38               ' グループ番号：
            stThta.iPP37_1 = stlocalProbeData.iPP37_1           ' パターン番号1：
            stThta.fpp32_x = stlocalProbeData.fpp32_x           ' パターン座標1X：
            stThta.fpp32_y = stlocalProbeData.fpp32_y           ' パターン座標1Y：
            stThta.iPP37_2 = stlocalProbeData.iPP37_2           ' パターン番号2：
            stThta.fpp33_x = stlocalProbeData.fpp33_x           ' パターン座標2X：
            stThta.fpp33_y = stlocalProbeData.fpp33_y           ' パターン座標2Y：
            'V2.2.1.6②↑

            If stPLT.Z_ZON < stPLT.Z_ZOFF Then
                Z_PRINT("プローブON位置がプローブ待機位置よりも低く設定されています。")
                Z_PRINT(" [ON位置=" & stPLT.Z_ZON & "],[待機位置=" & stPLT.Z_ZOFF.ToString & "]")
            End If

        Catch ex As Exception

        End Try

    End Function

#End Region


#Region "現在のプローブデータファイルを読込んで指定Noのプローブデータを更新して、プローブデータファイルを書き込む"    'V2.2.0.0⑮
    ''' <summary>
    ''' 現在のプローブデータファイルを読込んで指定Noのプローブデータを更新して、プローブデータファイルを書き込む
    ''' </summary>
    ''' <returns></returns>
    Public Function UpdateProbeData(ByVal ProbNo As Integer) As Integer
        Dim stlocalProbeData(PROBE_DATA_MAX) As stPROBEDATA_TABLE       'V2.2.0.038　'V2.2.1.0①
        ''V2.2.1.0①　Dim stlocalProbeData(11) As stPROBEDATA_TABLE　
        Dim Maxno As Integer
        Dim Header As String = ""

        Try

            ' プローブデータを全て読込む
            ReadAllProbeCsv(stlocalProbeData, Maxno, Header)

            ' 指定のプローブNoのデータを更新する 
            stlocalProbeData(ProbNo).dTableOffsetX = stPLT.z_xoff      ' テーブル位置オフセット　トリムポジションオフセットX(mm)
            stlocalProbeData(ProbNo).dTableOffsetY = stPLT.z_yoff      ' テーブル位置オフセット　トリムポジションオフセットY(mm)
            stlocalProbeData(ProbNo).dBPOffsetX = stPLT.BPOX           ' ビーム位置オフセット　BP Offset X(mm)
            stlocalProbeData(ProbNo).dBPOffsetY = stPLT.BPOY           ' ビーム位置オフセット　BP Offset Y(mm)
            stlocalProbeData(ProbNo).ProbeOn = stPLT.Z_ZON             ' Z PROBE ON 位置(mm)
            stlocalProbeData(ProbNo).ProbeOff = stPLT.Z_ZOFF           ' Z PROBE OFF 位置(mm)                'V2.2.0.0⑳ 

            stlocalProbeData(ProbNo).iPP30 = stThta.iPP30              ' 位置補正モード：
            stlocalProbeData(ProbNo).iPP31 = stThta.iPP31              ' 位置補正方法：
            stlocalProbeData(ProbNo).fpp34_x = stThta.fpp34_x          ' 補正ポジションオフセットX：
            stlocalProbeData(ProbNo).fpp34_y = stThta.fpp34_y          ' 補正ポジションオフセットY：
            stlocalProbeData(ProbNo).fTheta = stThta.fTheta            ' 角度：
            stlocalProbeData(ProbNo).iPP38 = stThta.iPP38              ' グループ番号
            stlocalProbeData(ProbNo).iPP37_1 = stThta.iPP37_1          ' パターン番号
            stlocalProbeData(ProbNo).fpp32_x = stThta.fpp32_x          ' パターン座標1X：
            stlocalProbeData(ProbNo).fpp32_y = stThta.fpp32_y          ' パターン座標1Y：
            stlocalProbeData(ProbNo).iPP37_2 = stThta.iPP37_2          ' パターン番号
            stlocalProbeData(ProbNo).fpp33_x = stThta.fpp33_x          ' パターン座標1X：
            stlocalProbeData(ProbNo).fpp33_y = stThta.fpp33_y          ' パターン座標1Y：

            'プローブデータを書き込む
            WriteAllProbeCsv(stlocalProbeData, Maxno, Header)

        Catch ex As Exception

        End Try

    End Function

#End Region


#Region "指定したNoのプローブデータの読込み"    'V2.2.0.0⑮
    ''' <summary>
    ''' 指定したNoのプローブデータの読込み
    ''' </summary>
    ''' <returns></returns>
    Public Function ReadProbeCsv(ByVal No As Integer, ByRef stData As stPROBEDATA_TABLE, ByRef MaxNo As Integer) As Boolean
        Dim sFolder As String = ""
        Dim sData As String = ""
        Dim bHeader As Boolean = True
        Dim mData() As String
        Dim TableNo As Integer
        Dim dData As Double
        Dim Ret As Boolean = False

        Try

            sFolder = cPROBEDATA_PATH & cPROBEDATA_FILE

            If IO.File.Exists(sFolder) = True Then  ' ファイルが有る。
                Using sr As New System.IO.StreamReader(sFolder, System.Text.Encoding.GetEncoding("Shift-JIS"))
                    Do While Not sr.EndOfStream
                        sData = sr.ReadLine
                        If bHeader Then                                     ' １行目はタイトル行
                            bHeader = False
                            Continue Do
                        End If
                        mData = sData.Split(",")                            ' 文字列を','で分割して取出す
                        TableNo = Integer.Parse(mData(0))

                        If TableNo > MaxNo Then
                            MaxNo = TableNo
                        End If

                        If TableNo = No Then                                            ' 番号一致
                            stData.No = TableNo

                            If Double.TryParse(mData(1), dData) Then
                                stData.ProbeOn = Double.Parse(mData(1))                                ' プローブ接触位置
                            Else
                                Call Z_PRINT("[" & No.ToString & "]番目のデータ[" & mData(1) & "]が数値に変換できません。")
                                Return (False)
                            End If

                            'V2.2.0.021↓
                            If Double.TryParse(mData(2), dData) Then
                                stData.ProbeOff = Double.Parse(mData(2))                                ' プローブ待機位置
                            Else
                                Call Z_PRINT("[" & No.ToString & "]番目のデータ[" & mData(2) & "]が数値に変換できません。")
                                Return (False)
                            End If
                            'V2.2.0.021↑

                            If Double.TryParse(mData(3), dData) Then
                                stData.dTableOffsetX = Double.Parse(mData(3))                          ' テーブルオフセットＸ
                            Else
                                Call Z_PRINT("[" & No.ToString & "]番目のデータ[" & mData(3) & "]が数値に変換できません。")
                                Return (False)
                            End If

                            If Double.TryParse(mData(4), dData) Then
                                stData.dTableOffsetY = Double.Parse(mData(4))                          ' テーブルオフセットＹ
                            Else
                                Call Z_PRINT("[" & No.ToString & "]番目のデータ[" & mData(4) & "]が数値に変換できません。")
                                Return (False)
                            End If

                            If Double.TryParse(mData(5), dData) Then
                                stData.dBPOffsetX = Double.Parse(mData(5))                          ' BPオフセットＸ
                            Else
                                Call Z_PRINT("[" & No.ToString & "]番目のデータ[" & mData(5) & "]が数値に変換できません。")
                                Return (False)
                            End If

                            If Double.TryParse(mData(6), dData) Then
                                stData.dBPOffsetY = Double.Parse(mData(6))                          ' BPオフセットＹ
                            Else
                                Call Z_PRINT("[" & No.ToString & "]番目のデータ[" & mData(6) & "]が数値に変換できません。")
                                Return (False)
                            End If

                            'V2.2.1.6② ↓
                            If mData.Length > 8 Then
                                ' 新しいデータの場合、θ補正関係のパラメータが含まれるので20分割になる
                                If Short.TryParse(mData(7), dData) Then
                                    stData.iPP30 = Short.Parse(mData(7))                          ' 位置補正モード
                                Else
                                    Call Z_PRINT("[" & No.ToString & "]番目のデータ[" & mData(7) & "]が数値に変換できません。")
                                    Return (False)
                                End If
                                If Short.TryParse(mData(8), dData) Then
                                    stData.iPP31 = Short.Parse(mData(8))                          ' 位置補正方法
                                Else
                                    Call Z_PRINT("[" & No.ToString & "]番目のデータ[" & mData(8) & "]が数値に変換できません。")
                                    Return (False)
                                End If
                                If Double.TryParse(mData(9), dData) Then
                                    stData.fpp34_x = Double.Parse(mData(9))                          ' 補正ポジションオフセットX
                                Else
                                    Call Z_PRINT("[" & No.ToString & "]番目のデータ[" & mData(9) & "]が数値に変換できません。")
                                    Return (False)
                                End If
                                If Double.TryParse(mData(10), dData) Then
                                    stData.fpp34_y = Double.Parse(mData(10))                          ' 補正ポジションオフセットY
                                Else
                                    Call Z_PRINT("[" & No.ToString & "]番目のデータ[" & mData(10) & "]が数値に変換できません。")
                                    Return (False)
                                End If
                                If Double.TryParse(mData(11), dData) Then
                                    stData.fTheta = Double.Parse(mData(11))                          ' 角度
                                Else
                                    Call Z_PRINT("[" & No.ToString & "]番目のデータ[" & mData(11) & "]が数値に変換できません。")
                                    Return (False)
                                End If
                                If Short.TryParse(mData(12), dData) Then
                                    stData.iPP38 = Short.Parse(mData(12))                          ' グループ番号
                                Else
                                    Call Z_PRINT("[" & No.ToString & "]番目のデータ[" & mData(12) & "]が数値に変換できません。")
                                    Return (False)
                                End If
                                If Short.TryParse(mData(13), dData) Then
                                    stData.iPP37_1 = Short.Parse(mData(13))                          ' パターン番号1
                                Else
                                    Call Z_PRINT("[" & No.ToString & "]番目のデータ[" & mData(13) & "]が数値に変換できません。")
                                    Return (False)
                                End If
                                If Double.TryParse(mData(14), dData) Then
                                    stData.fpp32_x = Double.Parse(mData(14))                          ' パターン座標1X
                                Else
                                    Call Z_PRINT("[" & No.ToString & "]番目のデータ[" & mData(14) & "]が数値に変換できません。")
                                    Return (False)
                                End If
                                If Double.TryParse(mData(15), dData) Then
                                    stData.fpp32_y = Double.Parse(mData(15))                          ' パターン座標1Y
                                Else
                                    Call Z_PRINT("[" & No.ToString & "]番目のデータ[" & mData(15) & "]が数値に変換できません。")
                                    Return (False)
                                End If
                                If Short.TryParse(mData(16), dData) Then
                                    stData.iPP37_2 = Short.Parse(mData(16))                          ' パターン番号2
                                Else
                                    Call Z_PRINT("[" & No.ToString & "]番目のデータ[" & mData(16) & "]が数値に変換できません。")
                                    Return (False)
                                End If
                                If Double.TryParse(mData(17), dData) Then
                                    stData.fpp33_x = Double.Parse(mData(17))                          ' パターン座標2X
                                Else
                                    Call Z_PRINT("[" & No.ToString & "]番目のデータ[" & mData(17) & "]が数値に変換できません。")
                                    Return (False)
                                End If
                                If Double.TryParse(mData(18), dData) Then
                                    stData.fpp33_y = Double.Parse(mData(18))                          ' パターン座標2Y
                                Else
                                    Call Z_PRINT("[" & No.ToString & "]番目のデータ[" & mData(18) & "]が数値に変換できません。")
                                    Return (False)
                                End If

                                stData.Comment = mData(19)                   ' コメント
                            Else
                                ' 古いデータの場合、θ補正関係のパラメータがないので8分割になる
                                stData.Comment = mData(7)                   ' コメント
                            End If
                            'V2.2.1.6② ↑

                            Ret = True
                            Exit Do
                        End If
                    Loop
                End Using
            End If

            Return Ret

        Catch ex As Exception
            MsgBox("ReadProbeCsv() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (False)
        End Try
    End Function

#End Region


#Region "PROBEDATAファイルの内容を全て読込む"   'V2.2.0.0⑮

    ''' <summary>
    ''' PROBEDATAファイルの内容を全て読込む
    ''' </summary>
    ''' <param name="stData"></param>
    ''' <param name="MaxNo"></param>
    ''' <returns></returns>
    Public Function ReadAllProbeCsv(ByRef stData() As stPROBEDATA_TABLE, ByRef MaxNo As Integer, ByRef header As String) As Boolean

        Dim sFolder As String = ""
        Dim sData As String = ""
        Dim bHeader As Boolean = True
        Dim mData() As String
        Dim TableNo As Integer
        Dim dData As Double
        Dim Ret As Boolean = False

        Try

            sFolder = cPROBEDATA_PATH & cPROBEDATA_FILE

            If IO.File.Exists(sFolder) = True Then  ' ファイルが有る。
                Using sr As New System.IO.StreamReader(sFolder, System.Text.Encoding.GetEncoding("Shift-JIS"))
                    Do While Not sr.EndOfStream
                        sData = sr.ReadLine
                        If bHeader Then                                     ' １行目はタイトル行
                            bHeader = False
                            header = sData
                            Continue Do
                        End If
                        mData = sData.Split(",")                            ' 文字列を','で分割して取出す
                        TableNo = Integer.Parse(mData(0))

                        If TableNo > MaxNo Then
                            MaxNo = TableNo
                        End If

                        stData(TableNo).No = TableNo

                        If Double.TryParse(mData(1), dData) Then
                            stData(TableNo).ProbeOn = Double.Parse(mData(1))                                ' プローブ接触位置
                        Else
                            Call Z_PRINT("[" & TableNo.ToString & "]番目のデータ[" & mData(1) & "]が数値に変換できません。")
                            Return (False)
                        End If

                        'V2.2.0.0⑳↓
                        If Double.TryParse(mData(2), dData) Then
                            stData(TableNo).ProbeOff = Double.Parse(mData(2))                                ' プローブ待機位置
                        Else
                            Call Z_PRINT("[" & TableNo.ToString & "]番目のデータ[" & mData(2) & "]が数値に変換できません。")
                            Return (False)
                        End If
                        'V2.2.0.0⑳↑

                        If Double.TryParse(mData(3), dData) Then
                            stData(TableNo).dTableOffsetX = Double.Parse(mData(3))                          ' テーブルオフセットＸ
                        Else
                            Call Z_PRINT("[" & TableNo.ToString & "]番目のデータ[" & mData(3) & "]が数値に変換できません。")
                            Return (False)
                        End If

                        If Double.TryParse(mData(4), dData) Then
                            stData(TableNo).dTableOffsetY = Double.Parse(mData(4))                          ' テーブルオフセットＹ
                        Else
                            Call Z_PRINT("[" & TableNo.ToString & "]番目のデータ[" & mData(4) & "]が数値に変換できません。")
                            Return (False)
                        End If

                        If Double.TryParse(mData(5), dData) Then
                            stData(TableNo).dBPOffsetX = Double.Parse(mData(5))                          ' BPオフセットＸ
                        Else
                            Call Z_PRINT("[" & TableNo.ToString & "]番目のデータ[" & mData(5) & "]が数値に変換できません。")
                            Return (False)
                        End If

                        If Double.TryParse(mData(6), dData) Then
                            stData(TableNo).dBPOffsetY = Double.Parse(mData(6))                          ' BPオフセットＹ
                        Else
                            Call Z_PRINT("[" & TableNo.ToString & "]番目のデータ[" & mData(6) & "]が数値に変換できません。")
                            Return (False)
                        End If

                        'V2.2.1.6② ↓
                        If mData.Length > 8 Then
                            ' 新しいデータの場合、θ補正関係のパラメータが含まれるので20分割になる
                            If Short.TryParse(mData(7), dData) Then
                                stData(TableNo).iPP30 = Short.Parse(mData(7))                          ' 位置補正モード
                            Else
                                Call Z_PRINT("[" & TableNo.ToString & "]番目のデータ[" & mData(7) & "]が数値に変換できません。")
                                Return (False)
                            End If
                            If Short.TryParse(mData(8), dData) Then
                                stData(TableNo).iPP31 = Short.Parse(mData(8))                          ' 位置補正方法
                            Else
                                Call Z_PRINT("[" & TableNo.ToString & "]番目のデータ[" & mData(8) & "]が数値に変換できません。")
                                Return (False)
                            End If
                            If Double.TryParse(mData(9), dData) Then
                                stData(TableNo).fpp34_x = Double.Parse(mData(9))                          ' 補正ポジションオフセットX
                            Else
                                Call Z_PRINT("[" & TableNo.ToString & "]番目のデータ[" & mData(9) & "]が数値に変換できません。")
                                Return (False)
                            End If
                            If Double.TryParse(mData(10), dData) Then
                                stData(TableNo).fpp34_y = Double.Parse(mData(10))                          ' 補正ポジションオフセットY
                            Else
                                Call Z_PRINT("[" & TableNo.ToString & "]番目のデータ[" & mData(10) & "]が数値に変換できません。")
                                Return (False)
                            End If
                            If Double.TryParse(mData(11), dData) Then
                                stData(TableNo).fTheta = Double.Parse(mData(11))                          ' 角度
                            Else
                                Call Z_PRINT("[" & TableNo.ToString & "]番目のデータ[" & mData(11) & "]が数値に変換できません。")
                                Return (False)
                            End If
                            If Short.TryParse(mData(12), dData) Then
                                stData(TableNo).iPP38 = Short.Parse(mData(12))                          ' グループ番号
                            Else
                                Call Z_PRINT("[" & TableNo.ToString & "]番目のデータ[" & mData(12) & "]が数値に変換できません。")
                                Return (False)
                            End If
                            If Short.TryParse(mData(13), dData) Then
                                stData(TableNo).iPP37_1 = Short.Parse(mData(13))                          ' パターン番号1
                            Else
                                Call Z_PRINT("[" & TableNo.ToString & "]番目のデータ[" & mData(13) & "]が数値に変換できません。")
                                Return (False)
                            End If
                            If Double.TryParse(mData(14), dData) Then
                                stData(TableNo).fpp32_x = Double.Parse(mData(14))                          ' パターン座標1X
                            Else
                                Call Z_PRINT("[" & TableNo.ToString & "]番目のデータ[" & mData(14) & "]が数値に変換できません。")
                                Return (False)
                            End If
                            If Double.TryParse(mData(15), dData) Then
                                stData(TableNo).fpp32_y = Double.Parse(mData(15))                          ' パターン座標1Y
                            Else
                                Call Z_PRINT("[" & TableNo.ToString & "]番目のデータ[" & mData(15) & "]が数値に変換できません。")
                                Return (False)
                            End If
                            If Short.TryParse(mData(16), dData) Then
                                stData(TableNo).iPP37_2 = Short.Parse(mData(16))                          ' パターン番号2
                            Else
                                Call Z_PRINT("[" & TableNo.ToString & "]番目のデータ[" & mData(16) & "]が数値に変換できません。")
                                Return (False)
                            End If
                            If Double.TryParse(mData(17), dData) Then
                                stData(TableNo).fpp33_x = Double.Parse(mData(17))                          ' パターン座標2X
                            Else
                                Call Z_PRINT("[" & TableNo.ToString & "]番目のデータ[" & mData(17) & "]が数値に変換できません。")
                                Return (False)
                            End If
                            If Double.TryParse(mData(18), dData) Then
                                stData(TableNo).fpp33_y = Double.Parse(mData(18))                          ' パターン座標2Y
                            Else
                                Call Z_PRINT("[" & TableNo.ToString & "]番目のデータ[" & mData(18) & "]が数値に変換できません。")
                                Return (False)
                            End If
                            stData(TableNo).Comment = mData(19)                   ' コメント
                        Else
                            ' 古いデータの場合、θ補正関係のパラメータがないので8分割になる
                            stData(TableNo).Comment = mData(7)                   ' コメント
                        End If
                        'V2.2.1.6② ↑

                        Ret = True
                    Loop
                End Using
            End If

            Return Ret

        Catch ex As Exception
            MsgBox("ReadAllProbeCsv() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (False)
        End Try
    End Function

#End Region

#Region "プローブデータを書き込む"
    ''' <summary>
    '''     'プローブデータを書き込む    'V2.2.0.0⑮
    ''' </summary>
    ''' <param name="stlocalProbeData"></param>
    ''' <param name="Maxno"></param>
    ''' <returns></returns>
    Public Function WriteAllProbeCsv(ByRef stData() As stPROBEDATA_TABLE, ByVal Maxno As Integer, ByVal sHeaderData As String) As Integer
        Dim sFileName As String

        Try

            sFileName = cPROBEDATA_PATH & cPROBEDATA_FILE

            Using WSR As New System.IO.StreamWriter(sFileName, False, System.Text.Encoding.GetEncoding("Shift-JIS"))  ' 第２引数 上書きは、False
                WSR.WriteLine(sHeaderData)                          ' ヘッダ出力

                For No As Integer = 1 To Maxno
                    'V2.2.0.0⑳                    WSR.WriteLine(stData(No).No.ToString & "," & stData(No).ProbeOn.ToString("0.000") & "," & stData(No).dTableOffsetX.ToString("0.000") & "," & stData(No).dTableOffsetY.ToString("0.000") & "," & stData(No).dBPOffsetX.ToString("0.000") & "," & stData(No).dBPOffsetY.ToString("0.000") & "," & stData(No).Comment)
                    WSR.WriteLine(stData(No).No.ToString & "," & stData(No).ProbeOn.ToString("0.000") & "," & stData(No).ProbeOff.ToString("0.000") & "," _
                                  & stData(No).dTableOffsetX.ToString("0.000") & "," & stData(No).dTableOffsetY.ToString("0.000") & "," _
                                  & stData(No).dBPOffsetX.ToString("0.000") & "," & stData(No).dBPOffsetY.ToString("0.000") & "," _
                                  & stData(No).iPP30.ToString() & "," & stData(No).iPP31.ToString() & "," & stData(No).fpp34_x.ToString("0.000") & "," _
                                  & stData(No).fpp34_y.ToString("0.000") & "," & stData(No).fTheta.ToString("0.000") & "," _
                                  & stData(No).iPP38.ToString() & "," & stData(No).iPP37_1.ToString() & "," & stData(No).fpp32_x.ToString("0.000") & "," & stData(No).fpp32_y.ToString("0.000") & "," _
                                  & stData(No).iPP37_2.ToString() & "," & stData(No).fpp33_x.ToString("0.000") & "," & stData(No).fpp33_y.ToString("0.000") & "," _
                                  & stData(No).Comment)     'V2.2.0.0⑳
                Next
            End Using

        Catch ex As Exception

        End Try


    End Function

#End Region


End Module

'=============================== END OF FILE ===============================

