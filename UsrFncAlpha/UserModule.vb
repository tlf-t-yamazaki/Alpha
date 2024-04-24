'==============================================================================
'   Description : ユーザプログラム用固有ファンクション
'
'　 2012/11/16 First Written by N.Arata(OLFT)
'
'==============================================================================
Option Strict Off
Option Explicit On

Imports System.Threading.Thread
Imports System.Runtime.InteropServices
Imports LaserFront.Trimmer.DllSysPrm.SysParam
Imports LaserFront.Trimmer.DefWin32Fnc
Imports UsrFunc.My.Resources
Imports LaserFront.Trimmer.DllJog                                       'V2.2.0.0①

Module UserModule

#Region "抵抗種別判定"

#Region "カット抵抗判定(トリミング有り、マーキング無し）"
    '''=========================================================================
    ''' <summary>
    ''' カット抵抗判定
    ''' </summary>
    ''' <param name="stRegData">抵抗データ構造体</param>
    ''' <param name="rn">抵抗番号</param>
    ''' <returns>True = カットありの抵抗, False = 測定のみの抵抗</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsCutResistor(ByRef stRegData As Reg_Info(), ByVal rn As Integer) As Boolean
        Try
            If stRegData(rn).intSLP = SLP_VTRIMPLS Or stRegData(rn).intSLP = SLP_VTRIMMNS Or stRegData(rn).intSLP = SLP_RTRM Or stRegData(rn).intSLP = SLP_ATRIMPLS Or stRegData(rn).intSLP = SLP_ATRIMMNS Or stRegData(rn).intSLP = SLP_MARK Then 'V2.2.1.7①
                IsCutResistor = True
            Else
                IsCutResistor = False
            End If
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function

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
            IsCutResistor = IsCutResistor(stREG, rn)
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region

#Region "マーキングを含めたカット抵抗判定(トリミング有り、マーキング有り）"
    '''=========================================================================
    ''' <summary>
    ''' マーキングを含めたカット抵抗判定
    ''' </summary>
    ''' <param name="stRegData">抵抗データ構造体</param>
    ''' <param name="rn">抵抗番号</param>
    ''' <returns>True = カットありの抵抗, False = 測定のみの抵抗</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsCutResistorIncMarking(ByRef stRegData As Reg_Info(), ByVal rn As Integer) As Boolean
        Try
            If stRegData(rn).intSLP = SLP_VTRIMPLS Or stRegData(rn).intSLP = SLP_VTRIMMNS Or stRegData(rn).intSLP = SLP_RTRM Or stRegData(rn).intSLP = SLP_ATRIMPLS Or stRegData(rn).intSLP = SLP_ATRIMMNS Or stRegData(rn).intSLP = SLP_NG_MARK Or stRegData(rn).intSLP = SLP_OK_MARK Or stRegData(rn).intSLP = SLP_MARK Then 'V2.2.1.7①
                IsCutResistorIncMarking = True
            Else
                IsCutResistorIncMarking = False
            End If
        Catch ex As Exception
            Call Z_PRINT("UserModule.IsCutResistorIncMarking() TRAP ERROR = " & ex.Message & vbCrLf)
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function

    '''=========================================================================
    ''' <summary>
    ''' マーキングを含めたカット抵抗判定
    ''' </summary>
    ''' <param name="rn">抵抗番号</param>
    ''' <returns>True = カットありの抵抗, False = 測定のみの抵抗</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsCutResistorIncMarking(ByVal rn As Integer) As Boolean
        Try
            IsCutResistorIncMarking = IsCutResistorIncMarking(stREG, rn)
        Catch ex As Exception
            Call Z_PRINT("UserModule.IsCutResistorIncMarking() TRAP ERROR = " & ex.Message & vbCrLf)
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region

#Region "マーキングカット抵抗判定(マーキング有りのみ）"
    '''=========================================================================
    ''' <summary>
    ''' マーキングを含めたカット抵抗判定
    ''' </summary>
    ''' <param name="stRegData">抵抗データ構造体</param>
    ''' <param name="rn">抵抗番号</param>
    ''' <returns>True = カットありの抵抗, False = 測定のみの抵抗</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsMarking(ByRef stRegData As Reg_Info(), ByVal rn As Integer) As Boolean
        Try
            If stRegData(rn).intSLP = SLP_NG_MARK Or stRegData(rn).intSLP = SLP_OK_MARK Or stRegData(rn).intSLP = SLP_MARK Then 'V2.2.1.7①
                IsMarking = True
            Else
                IsMarking = False
            End If
        Catch ex As Exception
            Call Z_PRINT("UserModule.IsMarking() TRAP ERROR = " & ex.Message & vbCrLf)
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function

    '''=========================================================================
    ''' <summary>
    ''' マーキングを含めたカット抵抗判定
    ''' </summary>
    ''' <param name="stRegData">抵抗データ構造体</param>
    ''' <param name="rn">抵抗番号</param>
    ''' <returns>True = カットありの抵抗, False = 測定のみの抵抗</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsOkMarking(ByRef stRegData As Reg_Info(), ByVal rn As Integer) As Boolean
        Try
            If stRegData(rn).intSLP = SLP_OK_MARK Then
                IsOkMarking = True
            Else
                IsOkMarking = False
            End If
        Catch ex As Exception
            Call Z_PRINT("UserModule.IsMarking() TRAP ERROR = " & ex.Message & vbCrLf)
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
    '''=========================================================================
    ''' <summary>
    ''' マーキングを含めたカット抵抗判定
    ''' </summary>
    ''' <param name="rn">抵抗番号</param>
    ''' <returns>True = カットありの抵抗, False = 測定のみの抵抗</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsMarking(ByVal rn As Integer) As Boolean
        Try
            IsMarking = IsMarking(stREG, rn)
        Catch ex As Exception
            Call Z_PRINT("UserModule.IsMarking() TRAP ERROR = " & ex.Message & vbCrLf)
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region

#Region "指定抵抗番号以降のマーキングカットの抵抗番号を返す。"
    ''' <summary>
    ''' 指定抵抗番号以降のマーキングカットの抵抗番号取得
    ''' </summary>
    ''' <param name="rn">抵抗番号</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetMarkingResNo(ByVal rn As Integer) As Integer
        GetMarkingResNo = 0
        Try
            Dim iResNo As Integer
            For iResNo = rn + 1 To stPLT.RCount Step 1
                If UserModule.IsMarking(iResNo) Then    'マーキングデータ
                    GetMarkingResNo = iResNo
                    Return (GetMarkingResNo)
                End If
            Next
        Catch ex As Exception
            Call Z_PRINT("UserModule.IsCutResistor() TRAP ERROR = " & ex.Message & vbCrLf)
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region

#Region "OKマーキングとしての連番を返す。"
    ''' <summary>
    ''' OKマーキングとしての連番を返す
    ''' </summary>
    ''' <param name="rn">抵抗番号</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetOkMarkingResNo(ByVal rn As Integer) As Integer
        GetOkMarkingResNo = 0
        Try
            Dim iResNo As Integer
            Dim OKCount As Integer = 0
            For iResNo = 1 To stPLT.RCount Step 1
                If UserModule.IsOkMarking(stREG, iResNo) Then       'OKマーキングデータ
                    OKCount = OKCount + 1
                    If rn = iResNo Then
                        GetOkMarkingResNo = OKCount
                        Return (GetOkMarkingResNo)
                    End If
                End If
            Next
        Catch ex As Exception
            Call Z_PRINT("UserModule.IsCutResistor() TRAP ERROR = " & ex.Message & vbCrLf)
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region

#Region "測定のみかの判定"
    '''=========================================================================
    ''' <summary>
    ''' 測定判定
    ''' </summary>
    ''' <param name="stRegData">抵抗データ構造体</param>
    ''' <param name="rn">抵抗番号</param>
    ''' <returns>True = 測定のみの抵抗, False = カットありの抵抗</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsMeasureOnly(ByRef stRegData As Reg_Info(), ByVal rn As Short) As Boolean
        Try
            If stRegData(rn).intSLP = SLP_VMES Or stRegData(rn).intSLP = SLP_AMES Or stRegData(rn).intSLP = SLP_RMES Then
                IsMeasureOnly = True
            Else
                IsMeasureOnly = False
            End If
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function

    '''=========================================================================
    ''' <summary>
    ''' 測定判定
    ''' </summary>
    ''' <param name="rn">抵抗番号</param>
    ''' <returns>True = 測定のみの抵抗, False = カットありの抵抗</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsMeasureOnly(ByVal rn As Short) As Boolean
        Try
            IsMeasureOnly = IsMeasureOnly(stREG, rn)
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region

#Region "測定かの判定"
    '''=========================================================================
    ''' <summary>
    ''' 測定判定
    ''' </summary>
    ''' <param name="stRegData">抵抗データ構造体</param>
    ''' <param name="rn">抵抗番号</param>
    ''' <param name="MeasMode">測定モード(0:なし, 1:ITのみ 2:FTのみ 3:IT,FT両方)</param>
    ''' <returns>True:測定有り False:測定無し</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsMeasureMode(ByRef stRegData As Reg_Info(), ByVal rn As Short, ByVal MeasMode As Short) As Boolean
        Try
            IsMeasureMode = False

            If IsMarking(rn) Then                   ' マーキング抵抗は測定対象外
                Return (False)
            End If

            If MeasMode = MEAS_JUDGE_IT Then
                If stREG(rn).intMeasMode = MEAS_JUDGE_IT Or stREG(rn).intMeasMode = MEAS_JUDGE_BOTH Then
                    IsMeasureMode = True
                End If
            ElseIf MeasMode = MEAS_JUDGE_FT Then
                If stREG(rn).intMeasMode = MEAS_JUDGE_FT Or stREG(rn).intMeasMode = MEAS_JUDGE_BOTH Then
                    IsMeasureMode = True
                End If
            Else
                Call Z_PRINT("CheckMeasureMode:指定されたモードが正しく有りません=[" & MeasMode.ToString() & "]")
                IsMeasureMode = False
            End If
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function

    '''=========================================================================
    ''' <summary>
    ''' 測定判定
    ''' </summary>
    ''' <param name="rn">抵抗番号</param>
    ''' <param name="MeasMode">測定モード(0:なし, 1:ITのみ 2:FTのみ 3:IT,FT両方)</param>
    ''' <returns>True:測定有り False:測定無し</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsMeasureMode(ByVal rn As Short, ByVal MeasMode As Short) As Boolean
        Try
            IsMeasureMode = IsMeasureMode(stREG, rn, MeasMode)
            'V2.1.0.0⑤↓抵抗値変化量判定の場合は、必ず初期測定が必要
            If MeasMode = MEAS_JUDGE_IT AndAlso IsMeasureMode = False Then
                If UserSub.IsCutVariationJudgeExecute() AndAlso UserModule.IsCutResistor(stREG, rn) Then
                    IsMeasureMode = True
                End If
            End If
            'V2.1.0.0⑤↑
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region

#Region "×５測定モード対象の抵抗かを判定"
    ''' <summary>
    ''' ×５測定モード対象の抵抗かを判定する。
    ''' </summary>
    ''' <param name="rn"></param>
    ''' <returns>True:対象 False:測定無し</returns>
    ''' <remarks></remarks>
    Public Function IsMeasureResistor(ByVal rn As Short) As Boolean
        Try
            If IsMarking(rn) Then       ' マーキング抵抗は測定対象外
                Return (False)
            End If
            If stREG(rn).intMeasMode = MEAS_JUDGE_IT Or stREG(rn).intMeasMode = MEAS_JUDGE_FT Or stREG(rn).intMeasMode = MEAS_JUDGE_BOTH Then
                Return (True)
            Else
                Return (False)
            End If
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region

#Region "指定した抵抗番号と抵抗名が一致するかのチェック"
    '''=========================================================================
    ''' <summary>
    ''' 抵抗番号と抵抗名が完全一致するかチェックする
    ''' </summary>
    ''' <param name="iResNo">抵抗番号</param>
    ''' <param name="strRNO">抵抗名</param>
    ''' <returns>True：一致 False:不一致</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsResistorByName(ByVal iResNo As Integer, ByVal strRNO As String) As Boolean
        Try
            If stREG(iResNo).strRNO.Trim.CompareTo(strRNO.Trim) = 0 Then
                Return (True)
            Else
                Return (False)
            End If
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
    '''=========================================================================
    ''' <summary>
    ''' 抵抗番号と抵抗名が前方部分一致するかチェックする
    ''' </summary>
    ''' <param name="iResNo">抵抗番号</param>
    ''' <param name="strRNO">抵抗名</param>
    ''' <returns>True：一致 False:不一致</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsResistorByNameStartsWith(ByVal iResNo As Integer, ByVal strRNO As String) As Boolean
        Try
            If stREG(iResNo).strRNO.Trim.StartsWith(strRNO.Trim) Then
                Return (True)
            Else
                Return (False)
            End If
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function

#End Region

#Region "抵抗名から該当する抵抗番号を取得する"
    '''=========================================================================
    ''' <summary>
    ''' 抵抗名から該当する抵抗番号を取得する
    ''' </summary>
    ''' <param name="strRNO">抵抗番号（文字列）</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function GetResistorNoByName(ByVal strRNO As String) As Integer
        GetResistorNoByName = 0
        Try
            Dim iResNo As Integer

            For iResNo = 1 To stPLT.RCount Step 1
                If IsResistorByName(iResNo, strRNO) Then
                    Return (iResNo)
                End If
            Next

        Catch ex As Exception
            Call Z_PRINT("UserModule.GetResistorNoByName() TRAP ERROR = " & ex.Message & vbCrLf)
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region

#End Region

#Region "■■ 測定のみを除いた抵抗数の取得 ■■"
#If False Then
    ''' <summary>
    ''' 測定のみを除いた抵抗数の取得
    ''' </summary>
    ''' <param name="stPlate">プレートデータ構造体</param>
    ''' <param name="stRegData">抵抗データ構造体</param>
    ''' <returns>抵抗数</returns>
    ''' <remarks></remarks>

    Public Function GetRCountExceptMeasure(ByVal stPlate As PLATE_DATA, ByRef stRegData As Reg_Info()) As Short
        Try
            GetRCountExceptMeasure = 0
            For iResNo As Integer = 1 To stPlate.RCount Step 1
                If IsCutResistor(stRegData, iResNo) Then           ' カット有（測定のみでない）抵抗の場合
                    GetRCountExceptMeasure = GetRCountExceptMeasure + 1
                End If
            Next

        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End If
#End Region

#Region "抵抗データのコピー処理"
    ''' <summary>
    ''' 抵抗データ１レコードのコピー処理
    ''' </summary>
    ''' <param name="ToRes">コピー先</param>
    ''' <param name="FromRes">コピー元</param>
    ''' <remarks></remarks>
    Public Sub CopyResistorData(ByRef ToRes As Reg_Info, ByRef FromRes As Reg_Info)

        Try

            ToRes = FromRes
            ToRes.STCUT = DirectCast(ToRes.STCUT.Clone(), Cut_Info())
            ToRes.intOnExtEqu = DirectCast(ToRes.intOnExtEqu.Clone(), Short())      ' ＯＮ機器１～３
            ToRes.intOffExtEqu = DirectCast(ToRes.intOffExtEqu.Clone(), Short())    ' ＯＦＦ機器１～３

            For i As Integer = 1 To MAXCTN
                ToRes.STCUT(i).intCND = DirectCast(ToRes.STCUT(i).intCND.Clone(), Short())
                ToRes.STCUT(i).intIXN = DirectCast(ToRes.STCUT(i).intIXN.Clone(), Short())
                ToRes.STCUT(i).dblDL1 = DirectCast(ToRes.STCUT(i).dblDL1.Clone(), Double())
                ToRes.STCUT(i).lngPAU = DirectCast(ToRes.STCUT(i).lngPAU.Clone(), Integer())
                ToRes.STCUT(i).dblDEV = DirectCast(ToRes.STCUT(i).dblDEV.Clone(), Double())
                ToRes.STCUT(i).intIXMType = DirectCast(ToRes.STCUT(i).intIXMType.Clone(), Short())
                ToRes.STCUT(i).intIXTMM = DirectCast(ToRes.STCUT(i).intIXTMM.Clone(), Short())
            Next

            For i As Integer = 1 To MAX_LCUT
                ToRes.STCUT(i).dCutLen = DirectCast(ToRes.STCUT(i).dCutLen.Clone(), Double())
                ToRes.STCUT(i).dQRate = DirectCast(ToRes.STCUT(i).dQRate.Clone(), Double())
                ToRes.STCUT(i).dSpeed = DirectCast(ToRes.STCUT(i).dSpeed.Clone(), Double())
                ToRes.STCUT(i).dAngle = DirectCast(ToRes.STCUT(i).dAngle.Clone(), Double())
                ToRes.STCUT(i).dTurnPoint = DirectCast(ToRes.STCUT(i).dTurnPoint.Clone(), Double())
            Next

            'V2.0.0.2②↓
            For i As Integer = 1 To MAX_RETRACECUT
                ToRes.STCUT(i).dblRetraceOffX = DirectCast(ToRes.STCUT(i).dblRetraceOffX.Clone(), Double())
                ToRes.STCUT(i).dblRetraceOffY = DirectCast(ToRes.STCUT(i).dblRetraceOffY.Clone(), Double())
                ToRes.STCUT(i).dblRetraceQrate = DirectCast(ToRes.STCUT(i).dblRetraceQrate.Clone(), Double())
                ToRes.STCUT(i).dblRetraceSpeed = DirectCast(ToRes.STCUT(i).dblRetraceSpeed.Clone(), Double())
            Next
            'V2.0.0.2②↑

        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try

    End Sub

    ''' <summary>
    ''' ユーザデータのコピー処理
    ''' </summary>
    ''' <param name="ToUser">コピー先</param>
    ''' <param name="FromUser">コピー元</param>
    ''' <remarks></remarks>
    Public Sub CopyUserData(ByRef ToUser As USER_DATA, ByRef FromUser As USER_DATA)
        Try
            ToUser.Initialize()
            For i As Integer = 0 To 1                                      ' MAXBLKX
                For j As Integer = 0 To 1                                  ' MAXBLKY
                Next j
            Next i

            ToUser = FromUser
            ToUser.iResUnit = DirectCast(FromUser.iResUnit.Clone(), Integer())
            ToUser.dNomCalcCoff = DirectCast(FromUser.dNomCalcCoff.Clone(), Double())
            ToUser.dTargetCoff = DirectCast(FromUser.dTargetCoff.Clone(), Double())
            ToUser.iChangeSpeed = DirectCast(FromUser.iChangeSpeed.Clone(), Integer())
            ToUser.dItVal = DirectCast(FromUser.dItVal.Clone(), Double())
            ToUser.dFtVal = DirectCast(FromUser.dFtVal.Clone(), Double())
            ToUser.dDev = DirectCast(FromUser.dDev.Clone(), Double())

        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub

    ''' <summary>
    ''' 全ての抵抗データのコピー処理（全配列）
    ''' </summary>
    ''' <param name="stPlate">プレートデータ構造体</param>
    ''' <param name="ToRes">コピー先</param>
    ''' <param name="FromRes">コピー元</param>
    ''' <remarks></remarks>
    Public Sub CopyResistorDataArray(ByVal stPlate As PLATE_DATA, ByRef ToRes As Reg_Info(), ByRef FromRes As Reg_Info())

        Try
            For i As Integer = 1 To MAXRNO Step 1
                ToRes(i).Initialize()
                Call CopyResistorData(ToRes(i), FromRes(i))
            Next

        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try

    End Sub
#End Region

#Region "指定した位置から後方へカットデータをコピーする。"
#If False Then
    ''' <summary>
    ''' 指定した位置から指定した位置までカットデータをコピーする。
    ''' </summary>
    ''' <param name="stPlate">プレートデータ構造体</param>
    ''' <param name="stRegData">抵抗データ構造体</param>
    ''' <param name="Res">コピー元の抵抗番号</param>
    ''' <param name="FromCut">コピー元のカット番号</param>
    ''' <param name="ToCut">コピーをする最後のカット番号</param>
    ''' <param name="bForce">強制コピーモード</param>
    ''' <param name="bCut">カット方法のコピー有無</param>
    ''' <param name="bCTYP">カット形状のコピー有無</param>
    ''' <param name="bLen">カット長のコピー有無</param>
    ''' <param name="bANG">カット方向のコピー有無</param>
    ''' <param name="bSpeed">カット速度のコピー有無</param>
    ''' <param name="bCutCnd">カット条件</param>
    ''' <remarks>全ての抵抗データを対象とする。</remarks>
    Public Sub CutDataCopy(ByVal stPlate As PLATE_DATA, ByRef stRegData As Reg_Info(), ByVal Res As Integer, ByVal FromCut As Integer, ByVal ToCut As Integer, ByVal bForce As Boolean, ByVal bCut As Boolean, ByVal bCTYP As Boolean, ByVal bLen As Boolean, ByVal bANG As Boolean, ByVal bSpeed As Boolean, ByVal bCutCnd As Boolean)

        Try

            Dim iFrom As Integer

            If ToCut > MAXCTN Then
                ToCut = MAXCTN
            End If

            For iR As Integer = Res To stPlate.RCount Step 1                ' 抵抗は指定抵抗番号から後の全ての番号が対象
                If IsCutResistor(stRegData, iR) Then                        ' カット有（測定のみでない）抵抗の場合
                    iFrom = FromCut
                    If stRegData(iR).intTNN < iFrom Then                    ' もしカットデータ数が足りない時は、足りない番号からコピーする。
                        iFrom = stRegData(iR).intTNN + 1
                    End If
                    For iC As Integer = iFrom To ToCut Step 1               ' カットは、コピー元のカット番号の次からへ元をコピーする。
                        If iR = Res And iC = FromCut Then                   ' コピー元と先が同じ場合はスキップ
                            Continue For
                        End If
                        If stRegData(iR).intTNN < iC Or bForce Then             ' カットデータが新規の場合または強制コピーモードの時は全てをコピーする。
                            stRegData(iR).STCUT(iC) = stRegData(Res).STCUT(FromCut)
                            stRegData(iR).STCUT(iC).intCND = DirectCast(stRegData(iR).STCUT(iC).intCND.Clone(), Short())
                            stRegData(iR).STCUT(iC).intIXN = DirectCast(stRegData(iR).STCUT(iC).intIXN.Clone(), Short())
                            stRegData(iR).STCUT(iC).dblDL1 = DirectCast(stRegData(iR).STCUT(iC).dblDL1.Clone(), Double())
                            stRegData(iR).STCUT(iC).lngPAU = DirectCast(stRegData(iR).STCUT(iC).lngPAU.Clone(), Integer())
                            stRegData(iR).STCUT(iC).dblDEV = DirectCast(stRegData(iR).STCUT(iC).dblDEV.Clone(), Double())
                            stRegData(iR).STCUT(iC).intIXMType = DirectCast(stRegData(iR).STCUT(iC).intIXMType.Clone(), Short())
                            stRegData(iR).STCUT(iC).intIXTMM = DirectCast(stRegData(iR).STCUT(iC).intIXTMM.Clone(), Short())
                        Else                                                ' 既存のカットデータの場合は指定された項目だけコピーする。
                            If bCut Then                                    ' カット方法
                                stRegData(iR).STCUT(iC).intCUT = stRegData(Res).STCUT(FromCut).intCUT
                            End If
                            If bCTYP Then                                   ' カット形状
                                stRegData(iR).STCUT(iC).intCTYP = stRegData(Res).STCUT(FromCut).intCTYP
                            End If
                            If bLen Then                                    ' カット長
                                stRegData(iR).STCUT(iC).dblDL2 = stRegData(Res).STCUT(FromCut).dblDL2
                                stRegData(iR).STCUT(iC).dblDL3 = stRegData(Res).STCUT(FromCut).dblDL3
                                For iX As Integer = 1 To MAXIDX Step 1
                                    stRegData(iR).STCUT(iC).dblDL1(iX) = stRegData(iR).STCUT(iC).dblDL1(iX)
                                Next iX
                            End If
                            If bANG Then                                    ' カット方向
                                stRegData(iR).STCUT(iC).intANG = stRegData(Res).STCUT(FromCut).intANG
                                stRegData(iR).STCUT(iC).intANG2 = stRegData(Res).STCUT(FromCut).intANG2
                            End If
                            If bSpeed Then                                  ' カット速度
                                stRegData(iR).STCUT(iC).dblV1 = stRegData(Res).STCUT(FromCut).dblV1
                            End If
                            If bCutCnd Then                                 ' カット条件
                                For i As Integer = 1 To MAXCND
                                    stRegData(iR).STCUT(iC).intCND(i) = stRegData(Res).STCUT(FromCut).intCND(i)
                                Next
                            End If
                        End If
                    Next iC
                    If stRegData(iR).intTNN < ToCut Then
                        stRegData(iR).intTNN = ToCut                            ' 全ての抵抗データでカット数を合わせる。    
                    End If
                End If
            Next iR
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub
#End If
#End Region

#Region "指定した位置から後方へ抵抗データをコピーする。"
    ''' <summary>
    ''' 抵抗データのコピー
    ''' </summary>
    ''' <param name="stRegData">抵抗データ構造体</param>
    ''' <param name="FromRes">コピー元の抵抗番号</param>
    ''' <param name="ToRes">コピーをする最後の抵抗番号</param>
    ''' <remarks>プローブデータとカットデータはコピーしない</remarks>
    Public Sub ResistorDataCopy(ByRef stRegData As Reg_Info(), ByVal FromRes As Integer, ByVal ToRes As Integer)

        Try
            Dim iFrom As Integer, iNo As Integer
            Dim sResNumber As String = ""
            Dim sResName As String = ""

            iFrom = FromRes + 1

            For i As Integer = 0 To stRegData(FromRes).strRNO.Length - 1
                If Char.IsNumber(stRegData(FromRes).strRNO.Chars(i)) Then
                    sResNumber = sResNumber + stRegData(FromRes).strRNO.Chars(i)
                Else
                    sResName = sResName + stRegData(FromRes).strRNO.Chars(i)
                End If
            Next i
            If sResName.Equals(String.Empty) Then   ' 文字列は空です
                sResName = "R"
            End If
            If sResNumber.Equals(String.Empty) Then ' 文字列は空です
                iNo = FromRes
            Else
                iNo = Integer.Parse(sResNumber)
            End If

            For iR As Integer = iFrom To ToRes Step 1                       ' 抵抗は指定抵抗番号から後の全ての番号が対象
                iNo = iNo + 1
                stRegData(iR).strRNO = sResName & iNo.ToString("0")     ' 抵抗名
                stRegData(iR).strTANI = stRegData(FromRes).strTANI      ' 単位("V","Ω" 等)
                stRegData(iR).intSLP = stRegData(FromRes).intSLP        ' 電圧変化スロープ(1:+V, 2:-V, 4:抵抗, 5:電圧測定のみ, 6:抵抗測定のみ 7:NGﾏｰｷﾝｸﾞ)
                stRegData(iR).lngRel = stRegData(FromRes).lngRel        ' リレービット
                stRegData(iR).dblNOM = stRegData(FromRes).dblNOM        ' トリミング目標値
                stRegData(iR).dblITL = stRegData(FromRes).dblITL        ' 初期判定下限値 (ITLO)
                stRegData(iR).dblITH = stRegData(FromRes).dblITH        ' 初期判定上限値 (ITHI)
                stRegData(iR).dblFTL = stRegData(FromRes).dblFTL        ' 終了判定下限値 (FTLO)
                stRegData(iR).dblFTH = stRegData(FromRes).dblFTH        ' 終了判定上限値 (FTHI)
                stRegData(iR).intMode = stRegData(FromRes).intMode      ' 判定モード(0:比率(%), 1:数値(絶対値))
                stRegData(iR).intTMM1 = stRegData(FromRes).intTMM1      ' モード(0:高速(コンパレータ非積分モード), 1:高精度(積分モード))
                'stRegData(iR).intPRH = stRegData(FromRes).intPRH        ' ハイ側プローブ番号(High Probe No.)
                'stRegData(iR).intPRL = stRegData(FromRes).intPRL        ' ロー側プローブ番号(Low Probe No.)
                'stRegData(iR).intPRG = stRegData(FromRes).intPRG        ' ガードプローブ番号(Gaude probe No.)
                stRegData(iR).intMType = stRegData(FromRes).intMType    ' 測定種別(0=内部測定, 1=外部測定)
            Next iR
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub
#End Region

#Region "指定した位置から後方へプローブ番号を増減する。"
    Public Sub ProbeNumberIncDec(ByRef stRegData As Reg_Info(), ByVal FromRes As Integer, ByVal ToRes As Integer, ByVal PRHIncDec As Integer, ByVal PRLIncDec As Integer)

        Try

            Dim iFrom As Integer
            Dim iDiffHI As Integer, iDiffLO As Integer

            iFrom = FromRes + 1
            For iR As Integer = iFrom To ToRes Step 1                       ' 抵抗は指定抵抗番号から後の全ての番号が対象
                iDiffHI = PRHIncDec * (iR - FromRes)
                iDiffLO = PRLIncDec * (iR - FromRes)
                stRegData(iR).intPRH = stRegData(FromRes).intPRH + iDiffHI  ' ハイ側プローブ番号(High Probe No.)
                stRegData(iR).intPRL = stRegData(FromRes).intPRL + iDiffLO  ' ロー側プローブ番号(Low Probe No.)
            Next iR
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub
#End Region

#Region "ＴＸ，ＴＹ補正前のデータ設定"
    ''' <summary>
    ''' ＴＸ，ＴＹ補正前のデータ設定
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SetTrimDataForTXTY()

        Try
            Dim ResNo As Integer = 0
            For rn As Integer = 1 To stPLT.RCount
                If UserModule.IsCutResistor(rn) Then
                    ResNo = ResNo + 1
                    typResistorInfoArray(ResNo).intResNo = ResNo
                    typResistorInfoArray(ResNo).intCutCount = stREG(rn).intTNN
                    typResistorInfoArray(ResNo).Initialize()
                    For cn As Integer = 1 To stREG(rn).intTNN
                        typResistorInfoArray(ResNo).ArrCut(cn).intCutNo = cn
                        typResistorInfoArray(ResNo).ArrCut(cn).dblStartPointX = stREG(rn).STCUT(cn).dblSTX
                        typResistorInfoArray(ResNo).ArrCut(cn).dblStartPointY = stREG(rn).STCUT(cn).dblSTY
                        typResistorInfoArray(ResNo).ArrCut(cn).dblTeachPointX = stREG(rn).STCUT(cn).dblSTX
                        typResistorInfoArray(ResNo).ArrCut(cn).dblTeachPointY = stREG(rn).STCUT(cn).dblSTY
                    Next cn
                End If
            Next rn

            typPlateInfo.intBlockCntXDir = stPLT.BNX                ' ブロック数X
            typPlateInfo.intBlockCntYDir = stPLT.BNY                ' ブロック数Y
            typPlateInfo.dblBlockSizeXDir = stPLT.zsx               ' ブロック(抵抗)サイズx(mm)
            typPlateInfo.dblBlockSizeYDir = stPLT.zsy               ' ブロック(抵抗)サイズy(mm)
            typPlateInfo.dblTableOffsetXDir = stPLT.z_xoff          ' テーブル位置オフセットX→トリムポジションオフセットX(mm)
            typPlateInfo.dblTableOffsetYDir = stPLT.z_yoff          ' テーブル位置オフセットY→トリムポジションオフセットY(mm)
            typPlateInfo.dblBpOffSetXDir = stPLT.BPOX               ' BP Offset X(mm)
            typPlateInfo.dblBpOffSetYDir = stPLT.BPOY               ' BP Offset Y(mm)
            'V2.0.0.0①            typPlateInfo.intResistDir = 0                           ' 抵抗並び方向０はＸ方向、１はＹ方向
            typPlateInfo.intResistDir = Integer.Parse(GetPrivateProfileString_S("USER", "TXTY_DIRECTION", SYSPARAMPATH, "1"))   'V2.0.0.0①
            typPlateInfo.intResistCntInBlock = ResNo                ' 1ブロック内抵抗数=1グループ内抵抗数=抵抗数
            typPlateInfo.intResistCntInGroup = ResNo                ' 1ブロック内抵抗数=1グループ内抵抗数=抵抗数
            If UserSub.IsTrimType3() Then
                typPlateInfo.intResistCntInGroup = UserSub.GetCircuitSum(stPLT, stREG)
            End If
            typPlateInfo.intGroupCntInBlockXBp = 1                  ' ブロック内ＢＰグループ数(サーキット数)
            typPlateInfo.intGroupCntInBlockYStage = 1               ' ブロック内ステージグループ数
            typPlateInfo.dblChipSizeXDir = stPLT.dblChipSizeXDir    ' チップサイズX
            typPlateInfo.dblChipSizeXDir = stPLT.zsx                ' チップサイズX　V2.0.0.0①チップサイズが０なので、ブロックサイズを設定する。
            typPlateInfo.dblChipSizeYDir = stPLT.dblChipSizeYDir    ' チップサイズY
            typPlateInfo.dblStepOffsetXDir = 0                      ' ステップオフセット量X
            typPlateInfo.dblStepOffsetYDir = 0                      ' ステップオフセット量Y
            typPlateInfo.dblBpGrpItv = 0                            ' BPグループ間隔（以前のCHIPのグループ間隔）
            typPlateInfo.dblStgGrpItvX = 0                          ' X方向のステージグループ間隔（以前のＣＨＩＰのステップ間インターバル）
            typPlateInfo.dblStgGrpItvY = 0                          ' Y方向のステージグループ間隔（以前のＣＨＩＰのステップ間インターバル）
            typPlateInfo.intBlkCntInStgGrpX = stPLT.BNX             ' X方向のステージグループ内ブロック数
            typPlateInfo.intBlkCntInStgGrpY = stPLT.BNY             ' Y方向のステージグループ内ブロック数

            gfCorrectPosX = dblCorrectX                             ' ＸＹΘ補正Ｘ
            gfCorrectPosY = dblCorrectY                             ' ＸＹΘ補正Ｙ

        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try

    End Sub
#End Region

#Region "ＴＸ，ＴＹ補正後のデータ設定"
    ''' <summary>
    ''' ＴＸ，ＴＹ補正後のデータ設定
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SetTrimDataFromTXTY()
        Try
            'If stPLT.dblChipSizeXDir <> 0.0 Then
            '    stPLT.dblChipSizeXDir = typPlateInfo.dblChipSizeXDir    ' ステップサイズX
            'End If
            'stPLT.zsx = typPlateInfo.dblBlockSizeXDir                   ' ブロック(抵抗)サイズx(mm)
            'stPLT.zsy = typPlateInfo.dblBlockSizeYDir                   ' ブロック(抵抗)サイズy(mm)
            If giAppMode = APP_MODE_TY Then
                stPLT.zsx = typPlateInfo.dblChipSizeXDir                    ' チップサイズ×ブロック(抵抗)数
                stPLT.dblStepOffsetXDir = typPlateInfo.dblStepOffsetXDir    ' ステップオフセット量X
                stPLT.dblStepOffsetYDir = typPlateInfo.dblStepOffsetYDir    ' ステップオフセット量Y
            Else
                If stPLT.dblChipSizeYDir <> 0.0 Then
                    stPLT.dblChipSizeYDir = typPlateInfo.dblChipSizeYDir    ' チップサイズY
                End If
            End If

            Dim ResNo As Integer = 1
            For rn As Integer = 1 To stPLT.RCount
                If UserModule.IsCutResistor(rn) Then
                    For cn As Integer = 1 To stREG(rn).intTNN
                        stREG(rn).STCUT(cn).dblSTX = typResistorInfoArray(ResNo).ArrCut(cn).dblStartPointX
                        stREG(rn).STCUT(cn).dblSTY = typResistorInfoArray(ResNo).ArrCut(cn).dblStartPointY
                    Next cn
                    ResNo = ResNo + 1
                End If
            Next rn
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try

    End Sub
#End Region

#Region "スタートポジション ティーチング(TEACH(F8))処理"
    '''=========================================================================
    ''' <summary>スタートポジション ティーチング(TEACH(F8))処理</summary>
    ''' <returns>cFRS_NORMAL   = 正常
    '''          上記以外      = エラー
    ''' </returns>
    ''' <remarks>BPオフセット値とトリミングスタート点をティーチングで設定する</remarks>
    '''=========================================================================
    Public Function User_TxTyTeach() As Short

        Dim r As Short                                                  ' Return Value From Function

        Try
            '--------------------------------------------------------------------------
            '   初期設定処理
            '--------------------------------------------------------------------------
            User_TxTyTeach = 0                                           ' Return値 = Normal
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
            'V2.0.0.0⑮            giTemplateGroup = -1                                        ' ﾃﾝﾌﾟﾚｰﾄｸﾞﾙｰﾌﾟ番号設定するため初期化
            'V2.0.0.0⑮            r = Ptn_Match_Exe()                                         ' パターン認識実行
            'V2.0.0.0⑮            If (r <> cFRS_NORMAL) Then
            'V2.0.0.0⑮                Return (r)                                              ' Return値設定
            'V2.0.0.0⑮            End If
            'V2.0.0.0⑮            Call ObjSys.Ilum_Ctrl(gSysPrm, Z1, ZOPT)                     ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ点灯(制御有時有効)


            ' クロスライン表示用
            r = ObjTch.SetCrossLineObject(gparModules)
            If r <> cFRS_NORMAL Then
                MsgBox("User.User_TxTyTeach() SetCrossLineObject ERROR")
            End If

            '--------------------------------------------------------------------------
            '   ＴＸ、ＴＹ用プレートデータ設定
            '--------------------------------------------------------------------------
            SetTrimDataForTXTY()

            '--------------------------------------------------------------------------
            '   ティーチングコントロール表示
            '--------------------------------------------------------------------------
            Dim TxTyObj As frmTxTyTeach = New frmTxTyTeach()

            TryCast(TxTyObj, Form).Show(Form1)                             'V6.0.0.0⑪

            User_TxTyTeach = TxTyObj.Execute()                            'V6.0.0.0⑬

            'TxTyObj.ShowDialog()
            r = TxTyObj.sGetReturn()                         ' Return値 = コマンド終了結果

            '--------------------------------------------------------------------------
            '   ティーチング結果取得
            '--------------------------------------------------------------------------
            If (r = cFRS_ERR_START Or r = cFRS_TxTy) Then               ' ティーチング処理正常終了
                SetTrimDataFromTXTY()
            End If

            ObjCrossLine.CrossLineOff()                                 ' クロスラインの非表示

            Return (r)                                                  ' Return値 = 正常

            ' トラップエラー発生時 
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
            Return (cERR_TRAP)                                          ' Return値 = トラップエラー
        End Try
    End Function
#End Region

#Region "ＪＯＧ操作画面処理用共通関数"
    '========================================================================================
    '   ＪＯＧ操作画面処理用共通関数
    '========================================================================================
#Region "ジョグ操作用変数定義"
    '========================================================================================
    '   ジョグ操作用変数定義(ＴＸ/ＴＹティーチング他共通)
    '========================================================================================
    '-------------------------------------------------------------------------------
    '   ジョグ操作用定義
    '-------------------------------------------------------------------------------
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

        Dim TenKey() As Button                                  ' V2.2.0.0①
        Dim KeyDown As Keys                                     ' V2.2.0.0①

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

    '    '----- 入力モード -----
    Public Const MD2_BUTN As Integer = 0                    ' 画面ボタン入力

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

        Try
            ' 移動ピッチスライダー初期設定
            If (stJOG.Md = MODE_BP) Then                            ' モード = 1(BP移動) ?
                dblTchMoval(IDX_PIT) = gSysPrm.stSYP.gBpPIT         ' BP用ﾋﾟｯﾁ設定
                dblTchMoval(IDX_HPT) = gSysPrm.stSYP.gBpHighPIT
                dblTchMoval(IDX_PAU) = gSysPrm.stSYP.gPitPause
                'V2.2.1.1③↓
                If gSysPrm.stDEV.giBpSize = 40 Then
                    dblTchMoval(3) = 1
                End If
                'V2.2.1.1③↑
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
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub
#End Region

#Region "BP/XYテーブルのJOG操作(Do Loopなし)"
    '''=========================================================================
    '''<summary>BP/XYテーブルのJOG操作</summary>
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
    Public Function JogEzMove_Ex(ByRef stJOG As JOG_PARAM, ByVal SysPrm As SYSPARAM_PARAM, _
                         ByRef TBarLowPitch As System.Windows.Forms.TrackBar, _
                         ByRef TBarHiPitch As System.Windows.Forms.TrackBar, _
                         ByRef TBarPause As System.Windows.Forms.TrackBar, _
                         ByRef LblTchMoval0 As System.Windows.Forms.Label, _
                         ByRef LblTchMoval1 As System.Windows.Forms.Label, _
                         ByRef LblTchMoval2 As System.Windows.Forms.Label, _
                         ByRef dblTchMoval() As Double) As Integer

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
               (giAppMode <> APP_MODE_FINEADJ) Then
                '(giAppMode <> APP_MODE_TRIM) Then                      
                stJOG.cgX = 0.0# : stJOG.cgY = 0.0#                     ' 移動量X,Y
            End If

            'If (giAppMode = APP_MODE_TRIM) Then                        
            If (giAppMode = APP_MODE_FINEADJ) Then
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
                        r = cFRS_NORMAL                                 ' Retuen値 = 正常
                    End If

                    ' 現在の位置を表示する
                    Call DispPosition(stJOG, 1)
                    Call Form1.System1.WAIT(dblTchMoval(IDX_PAU))       ' Wait(sec)

                    InpKey = CType(CtrlJog.MouseClickLocation.Clear(InpKey), UShort)    'V2.2.0.0① 
                    stJOG.KeyDown = Keys.None                                           'V2.2.0.0① 

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
            Call MsgBox_Exception(ex.Message, Form1)
            Return (cERR_TRAP)                                          ' Return値 = 例外エラー 
        End Try
    End Function
#End Region

#Region "BP/XYテーブルのJOG操作"
    '''=========================================================================
    '''<summary>BP/XYテーブルのJOG操作</summary>
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
    Public Function JogEzMove(ByRef stJOG As JOG_PARAM, ByVal SysPrm As SYSPARAM_PARAM,
                         ByRef TBarLowPitch As System.Windows.Forms.TrackBar,
                         ByRef TBarHiPitch As System.Windows.Forms.TrackBar,
                         ByRef TBarPause As System.Windows.Forms.TrackBar,
                         ByRef LblTchMoval0 As System.Windows.Forms.Label,
                         ByRef LblTchMoval1 As System.Windows.Forms.Label,
                         ByRef LblTchMoval2 As System.Windows.Forms.Label,
                         ByRef dblTchMoval() As Double,
                         ByVal commonMethods As ICommonMethods) As Integer      ''V2.2.0.0①   引数 ICommonMethods 追加

        Dim r As Short

        Try
            '---------------------------------------------------------------------------
            '   初期処理
            '---------------------------------------------------------------------------
            X = 0.0 : Y = 0.0                                   ' 移動ﾋﾟｯﾁX,Y
            mvx = stJOG.PosX : mvy = stJOG.PosY                 ' BP or ﾃｰﾌﾞﾙ位置X,Y
            mvxBk = stJOG.PosX : mvyBk = stJOG.PosY
            ' キャリブレーション実行/カット位置補正【外部カメラ】時 ※相対座標を表示するためクリアしない
            ' トリミング時の一時停止画面もクリアしない
            If (giAppMode <> APP_MODE_CARIB_REC) And (giAppMode <> APP_MODE_CUTREVIDE) And
               (giAppMode <> APP_MODE_FINEADJ) Then
                stJOG.cgX = 0.0# : stJOG.cgY = 0.0#             ' 移動量X,Y
            End If
            stJOG.Flg = -1
            InpKey = 0
            Call Init_Proc(stJOG, TBarLowPitch, TBarHiPitch, TBarPause, LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval)

            ' 現在の位置を表示する(ﾃｷｽﾄﾎﾞｯｸｽの背景色を処理中(黄色)に設定する)
            Call DispPosition(stJOG, 1)
            'Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
            'Call Me.Focus()                                     ' フォーカスを設定する(テンキー入力のため)
            ''                                                   ' KeyPreviewプロパティをTrueにすると全てのキーイベントをまずフォームが受け取るようになる。

            ' メインフォームにJOG制御関数を設定する      'V2.2.0.0① 
            Form1.SetActiveJogMethod(AddressOf commonMethods.JogKeyDown,
                                              AddressOf commonMethods.JogKeyUp,
                                              AddressOf commonMethods.MoveToCenter)

            '---------------------------------------------------------------------------
            '   コンソールボタン又はコンソールキーからのキー入力処理を行う
            '---------------------------------------------------------------------------
            Do
                ' システムエラーチェック
                r = Form1.System1.SysErrChk_ForVBNET(giAppMode)
                If (r <> cFRS_NORMAL) Then GoTo STP_END

                ' メッセージポンプ
                '  →VB.NETはマルチスレッド対応なので、本来はイベントの開放などでなく、
                '    スレッドを生成してコーディングをするのが正しい。
                '    スレッドでなくても、最低でタイマーを利用する。
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(10)               ' CPU使用率を下げるためスリープ


                ' コンソールボタン又はコンソールキーからのキー入力
                Call ReadConsoleSw(stJOG, cin)                  ' キー入力

                '-----------------------------------------------------------------------
                '   入力キーチェック
                '-----------------------------------------------------------------------
                If (cin And CONSOLE_SW_RESET) Then              ' RESET SW ?
                    ' RESET SW押下時
                    If (stJOG.Opt And CONSOLE_SW_RESET) Then    ' RESETキー有効 ?
                        r = cFRS_ERR_RST                        ' Return値 = Cancel(RESETｷｰ)
                        Exit Do
                    End If

                    ' HALT SW押下時
                ElseIf (cin And CONSOLE_SW_HALT) Then           ' HALT SW ?
                    If (stJOG.Opt And CONSOLE_SW_HALT) Then     ' オプション(0:HALTキー無効, 1:HALTキー有効)
                        r = cFRS_ERR_HALT                       ' Return値 = HALTｷｰ
                        Exit Do
                    End If

                    ' START SW押下時
                ElseIf (cin And CONSOLE_SW_START) Then          ' START SW ?
                    If (stJOG.Opt And CONSOLE_SW_START) Then    ' STARTキー有効 ?
                        'stJOG.PosX = mvx                       ' 位置X,Y更新
                        'stJOG.PosY = mvy
                        r = cFRS_ERR_START                      ' Return値 = OK(STARTｷｰ) 
                        Exit Do
                    End If

                    ' Z SWがONからOFF(又はOFFからON)に切替わった時
                ElseIf (stJOG.bZ <> bZ) Then
                    If (stJOG.Opt And CONSOLE_SW_ZSW) Then      ' Zキー有効 ?
                        r = cFRS_ERR_Z                          ' Return値 = ZｷｰON/OFF
                        stJOG.bZ = bZ                           ' ON/OFF
                        Exit Do
                    End If

                    ' 矢印SW押下時
                ElseIf cin And &H1E00US Then                    ' 矢印SW
                    '「キー入力待ちモード」なら何もしない
                    If (stJOG.Md = MODE_KEY) Then

                    Else
                        If cin And &H100US Then                     ' HI SW ? 
                            mPIT = dblTchMoval(IDX_HPT)             ' mPIT = 移動高速ﾋﾟｯﾁ
                        Else
                            mPIT = dblTchMoval(IDX_PIT)             ' mPIT = 移動通常ﾋﾟｯﾁ
                        End If

                        ' XYテーブル絶対値移動(ソフトリミットチェック有り)
                        r = cFRS_NORMAL
                        If (stJOG.Md = MODE_STG) Then                ' モード = XYテーブル移動 ?
                            ' XYテーブル絶対値移動
                            r = Sub_XYtableMove(SysPrm, Form1.System1, Form1.Utility1, stJOG)
                            If (r <> cFRS_NORMAL) Then              ' ｴﾗｰ ?
                                If (Form1.System1.IsSoftLimitXY(r) = False) Then
                                    GoTo STP_END                    ' ｿﾌﾄﾘﾐｯﾄｴﾗｰ以外はｴﾗｰﾘﾀｰﾝ
                                End If
                            End If

                            '  モード = BP移動の場合
                        ElseIf (stJOG.Md = MODE_BP) Then
                            ' BP絶対値移動
                            r = Sub_BPmove(SysPrm, Form1.System1, Form1.Utility1, stJOG)
                            If (r <> cFRS_NORMAL) Then              ' BP移動エラー ?
                                If (Form1.System1.IsSoftLimitBP(r) = False) Then
                                    GoTo STP_END                    ' ｿﾌﾄﾘﾐｯﾄｴﾗｰ以外はｴﾗｰﾘﾀｰﾝ
                                End If
                            End If
                        End If

                        ' ソフトリミットエラーの場合は HI SW以外はOFFする
                        If (r <> cFRS_NORMAL) Then                  ' ｴﾗｰ ?
                            If (stJOG.BtnHI.BackColor = System.Drawing.Color.Yellow) Then
                                InpKey = cBIT_HI                    ' HI SW ON
                            Else
                                InpKey = 0                          ' HI SW以外はOFF
                            End If
                        End If

                        ' 現在の位置を表示する
                        Call DispPosition(stJOG, 1)
                        Call Form1.System1.WAIT(SysPrm.stSYP.gPitPause)    ' Wait(sec)
                    End If
                    InpKey = CType(CtrlJog.MouseClickLocation.Clear(InpKey), UShort)    'V2.2.0.0① 
                    stJOG.KeyDown = Keys.None                                           'V2.2.0.0① 

                End If

            Loop While (stJOG.Flg = -1)

            '---------------------------------------------------------------------------
            '   終了処理
            '---------------------------------------------------------------------------
            ' 座標表示用ﾃｷｽﾄﾎﾞｯｸｽの背景色を白色に設定する
            Call DispPosition(stJOG, 0)

            ' 親画面からOK/Cancelﾎﾞﾀﾝ押下 ?
            If (stJOG.Flg <> -1) Then
                r = stJOG.Flg
            End If

            ' OK(STARTｷｰ)なら位置X,Y更新
            If (r = cFRS_ERR_START) Then                            ' OK(STARTｷｰ) ?
                stJOG.PosX = mvx                                    ' 位置X,Y更新
                stJOG.PosY = mvy
            End If

STP_END:
            Call ZCONRST()                                          ' ｺﾝｿｰﾙｷｰﾗｯﾁ解除 
            Return (r)                                              ' Return値設定 

            ' トラップエラー発生時
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
            Return (cERR_TRAP)                                      ' Return値 = 例外エラー 

        Finally
            Form1.SetActiveJogMethod(Nothing, Nothing, Nothing)    'V6.0.0.0⑪

        End Try
    End Function
#End Region

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
            Call MsgBox_Exception(ex.Message, Form1)
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

        Dim xPos As Double = 0.0
        Dim yPos As Double = 0.0
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
                    ' トリミング時の一時停止画面表示中なら補正クロスラインを表示する
                    If (giAppMode = APP_MODE_FINEADJ) Or (giAppMode = APP_MODE_TX) Then
                        'xPos = Double.Parse(stJOG.TextX.Text)
                        'yPos = Double.Parse(stJOG.TextY.Text)
                        Call ZGETBPPOS(xPos, yPos)
                        ObjCrossLine.CrossLineDispXY(xPos, yPos)
                    End If
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

#Region "BP絶対値移動(ソフトリミットチェック有り)"
    '''=========================================================================
    ''' <summary>BP絶対値移動(ソフトリミットチェック有り)</summary>
    ''' <param name="SysPrm">(INP)システムパラメータ</param>
    ''' <param name="ObjSys">(INP)OcxSystemオブジェク</param>
    ''' <param name="ObjUtl">(INP)OcxUtilityオブジェク</param>
    ''' <param name="stJOG"> (I/O)JOG操作用パラメータ</param>
    ''' <returns>0=正常, 0以外:エラー</returns>
    '''=========================================================================
    Private Function Sub_BPmove(ByVal SysPrm As SYSPARAM_PARAM, ByVal ObjSys As Object, ByVal ObjUtl As Object, ByRef stJOG As JOG_PARAM) As Integer

        Dim r As Integer
        Dim strMSG As String

        Try
            ' BP移動量の算出(→X,Y)
            mvxBk = mvx                                             ' 現在の位置退避
            mvyBk = mvy
            'V2.2.0.0①↓
            If ((cin And CtrlJog.MouseClickLocation.Move) = &H0) Then           'V6.0.0.0⑧
                Call ObjUtl.GetBPmovePitch(cin, X, Y, mPIT, mvx, mvy, SysPrm.stDEV.giBpDirXy)
            Else
                'V6.0.0.0⑧              ↓
                Dim dirX As Double = 0.0
                Dim dirY As Double = 0.0
                Dim tmpX As Double = 0.0
                Dim tmpY As Double = 0.0
                ObjUtl.GetBPmovePitch(cin, dirX, dirY, 1.0, tmpX, tmpY, SysPrm.stDEV.giBpDirXy)   ' 符号を取得

                X = Math.Abs(CtrlJog.MouseClickLocation.DistanceX) * Math.Sign(dirX)
                Y = Math.Abs(CtrlJog.MouseClickLocation.DistanceY) * Math.Sign(dirY)
                mvx -= X
                mvy -= Y
                'V6.0.0.0⑧              ↑
            End If
            'V2.2.0.0①↑

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
    Private Function Sub_XYtableMove(ByVal SysPrm As SYSPARAM_PARAM, ByVal ObjSys As Object, ByVal ObjUtl As Object, ByRef stJOG As JOG_PARAM) As Integer

        Dim r As Integer
        Dim strMSG As String

        Try
            ' XYテーブル移動量の算出(→X,Y)
            mvxBk = X                                               ' 現在の位置退避
            mvyBk = Y
            'V2.2.0.0① ↓
            If ((cin And CtrlJog.MouseClickLocation.Move) = &H0) Then
                Call TrimClassCommon.GetXYmovePitch(cin, X, Y, mPIT, giStageYDir)
            Else
                Dim dirX As Double = 0.0
                Dim dirY As Double = 0.0
                TrimClassCommon.GetXYmovePitch(cin, dirX, dirY, 1.0, giStageYDir)   ' 符号を取得

                X = -(Math.Abs(CtrlJog.MouseClickLocation.DistanceX) * Math.Sign(dirX)) 'V6.0.0.0-24 -() 追加
                Y = -(Math.Abs(CtrlJog.MouseClickLocation.DistanceY) * Math.Sign(dirY)) 'V6.0.0.0-24 -() 追加
            End If
            'V2.2.0.0① ↑

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

#Region "ボタン押下時処理(ＪＯＧ操作画面)"
    '========================================================================================
    '   ボタン押下時処理(ＪＯＧ操作画面)
    '========================================================================================
#Region "HALTボタン押下時処理"
    '''=========================================================================
    '''<summary>HALTボタン押下時処理</summary>
    '''=========================================================================
    Public Sub SubBtnHALT_Click()
        InpKey = CONSOLE_SW_HALT
    End Sub
#End Region

#Region "STARTボタン押下時処理"
    '''=========================================================================
    '''<summary>STARTボタン押下時処理</summary>
    '''=========================================================================
    Public Sub SubBtnSTART_Click()
        InpKey = CONSOLE_SW_START
    End Sub
#End Region

#Region "RESETボタン押下時処理"
    '''=========================================================================
    '''<summary>RESETボタン押下時処理</summary>
    '''=========================================================================
    Public Sub SubBtnRESET_Click()
        InpKey = CONSOLE_SW_RESET
    End Sub
#End Region

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
    Public Sub GetInpKey(ByRef IKey As UShort)
        IKey = InpKey
    End Sub
#End Region

#Region "InpKeyを設定する"
    '''=========================================================================
    '''<summary>InpKeyを設定する</summary>
    '''<param name="IKey">(INP)InpKey</param>
    '''=========================================================================
    Public Sub PutInpKey(ByVal IKey As UShort)
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

#End Region

#Region "ＪＯＧ操作画面処理用トラックバー処理"
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

        'V2.2.1.1③ ↓
        ' BP最小分解能によって最小値を変更する
        If dblTchMoval(3) <> 0 Then
            minval = 1                                     ' 目盛2～ 
        End If
        'V2.2.1.1③ ↑

        TBarLowPitch.TickFrequency = dblTchMoval(IDX_PIT) * 1000        ' 0.001mm単位
        TBarLowPitch.Maximum = 100                                      ' 目盛1(or 2)～100(0.001m～0.1mm)
        TBarLowPitch.Minimum = minval
        TBarLowPitch.Value = dblTchMoval(IDX_PIT) * 1000        ' 0.001mm単位

        ' HIGHﾋﾟｯﾁがが範囲外なら範囲内に変更する
        If (dblTchMoval(IDX_HPT) < cHPT_LO) Then dblTchMoval(IDX_HPT) = cHPT_LO
        If (dblTchMoval(IDX_HPT) > cHPT_HI) Then dblTchMoval(IDX_HPT) = cHPT_HI

        ' HIGHﾋﾟｯﾁの目盛を設定する
        TBarHiPitch.TickFrequency = dblTchMoval(IDX_HPT) * 100          ' 0.01mm単位
        TBarHiPitch.Maximum = 500                                       ' 目盛1～100(0.01m～5.00mm)
        TBarHiPitch.Minimum = 1
        TBarHiPitch.Value = dblTchMoval(IDX_HPT) * 100          ' 0.01mm単位

        ' Pause Timeが範囲外なら範囲内に変更する
        If (dblTchMoval(IDX_PAU) < cPAU_LO) Then dblTchMoval(IDX_PAU) = cPAU_LO
        If (dblTchMoval(IDX_PAU) > cPAU_HI) Then dblTchMoval(IDX_PAU) = cPAU_HI

        ' Pause Timeの目盛を設定する
        TBarPause.TickFrequency = dblTchMoval(IDX_PAU) * 20             ' 0.5秒単位
        TBarPause.Maximum = 20                                          ' 目盛1～20(0.05秒～1.00秒)
        TBarPause.Minimum = 1
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

#End Region

#Region "ＪＯＧ操作画面処理用テンキー入力処理"
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

#End Region

#End Region

#Region "グループ数,ブロック数,チップ数(抵抗数),チップサイズを取得する(ＴＸ/ＴＹティーチング用)"
    '''=========================================================================
    ''' <summary>グループ数,ブロック数,チップ数(抵抗数),チップサイズを取得する</summary>
    ''' <param name="AppMode">  (INP)モード</param>
    ''' <param name="Gn">       (OUT)グループ数</param>
    ''' <param name="RnBn">     (OUT)チップ数(ＴＸティーチング時)または
    '''                              ブロック数(ＴＹティーチング時)</param>
    ''' <param name="DblChipSz">(OUT)チップサイズ</param>
    ''' <returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Public Function GetChipNumAndSize(ByVal AppMode As Short, ByRef Gn As Short, ByRef RnBn As Short, ByRef DblChipSz As Double) As Short

        Dim ChipNum As Short                                        ' チップ数(抵抗数)
        Dim ChipSzX As Double                                       ' チップサイズX
        Dim ChipSzY As Double                                       ' チップサイズY
        Dim strMSG As String

        Try
            ' 前処理(CHIP/NET共通)
            ChipNum = typPlateInfo.intResistCntInGroup              ' チップ数(抵抗数) = 1グループ内(1サーキット内)抵抗数
            ChipSzX = typPlateInfo.dblChipSizeXDir                  ' チップサイズX,Y
            ChipSzY = typPlateInfo.dblChipSizeYDir

            ' プレートデータからグループ数, ブロック数, チップ数(抵抗数), チップサイズを取得する
            If (AppMode = APP_MODE_TX) Then
                '----- ＴＸティーチング時 -----
                ' チップ数(抵抗数)を返す
                RnBn = ChipNum                                      ' 1グループ内(1サーキット内)抵抗数をセット
                ' グループ数を返す
                Gn = typPlateInfo.intGroupCntInBlockXBp             ' ＢＰグループ数(サーキット数)をセット
                ' チップサイズを返す
                If (typPlateInfo.intResistDir = 0) Then             ' チップ並びはX方向 ?
                    DblChipSz = System.Math.Abs(ChipSzX)
                Else
                    'V2.0.0.0①                    DblChipSz = System.Math.Abs(ChipSzY)
                    DblChipSz = ChipSzY             'V2.0.0.0①
                End If

            Else
                '----- ＴＹティーチング時 -----
                ' グループ数を返す
                Gn = typPlateInfo.intGroupCntInBlockYStage          ' ブロック内Stageグループ数をセット
                ' ブロック数とチップサイズを返す
                If (typPlateInfo.intResistDir = 0) Then             ' チップ並びはX方向 ?
                    RnBn = typPlateInfo.intBlockCntYDir             ' ブロック数Yをセット
                    DblChipSz = System.Math.Abs(ChipSzY)            ' チップサイズYをセット
                Else
                    RnBn = typPlateInfo.intBlockCntXDir             ' ブロック数Xをセット
                    DblChipSz = System.Math.Abs(ChipSzX)            ' チップサイズXをセット
                End If
            End If

            strMSG = "GetChipNumAndSize() Gn=" + Gn.ToString("0") + ", RnBn=" + RnBn.ToString("0") + ", ChipSZ=" + DblChipSz.ToString("0.00000")
            Console.WriteLine(strMSG)
            Return (cFRS_NORMAL)                                    ' Return値 = 正常

            ' トラップエラー発生時
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
            Return (cERR_TRAP)                                      ' Return値 = 例外エラー
        End Try
    End Function
#End Region

#Region "ブロックサイズを算出する【CHIP/NET用】"
    '''=========================================================================
    '''<summary>ブロックサイズを算出する【CHIP/NET用】</summary>
    '''<param name="dblBSX">(OUT) ブロックサイズX</param>
    '''<param name="dblBSY">(OUT) ブロックサイズY</param>
    '''=========================================================================
    Public Sub CalcBlockSize(ByRef dblBSX As Double, ByRef dblBSY As Double)

        Dim i As Integer
        Dim intChipNum As Integer
        Dim intGNx As Integer
        Dim intGNY As Integer
        Dim dData As Double = 0.0

        Try
            ' CHIP/NET時 
            ' グループ数X,Y
            intGNx = typPlateInfo.intGroupCntInBlockXBp                 ' ＢＰグループ数(サーキット数)
            'V2.0.0.0⑮            intGNY = typPlateInfo.intGroupCntInBlockXBp
            intGNY = typPlateInfo.intGroupCntInBlockYStage

            ' グループ内抵抗数             
            intChipNum = typPlateInfo.intResistCntInGroup

            ' ブロックサイズX,Yを求める
            If (typPlateInfo.intResistDir = 0) Then                     ' 抵抗(ﾁｯﾌﾟ)並び方向(0:X, 1:Y)
                ' 抵抗(ﾁｯﾌﾟ)並び方向 = X方向の場合
                If (intGNx = 1) Then
                    ' 1グループ(1サーキット)の場合
                    dData = typPlateInfo.dblChipSizeXDir * intChipNum   ' Data = チップサイズX * チップ数

                Else
                    ' 複数グループ(複数サーキット)の場合
                    For i = 1 To intGNx
                        If (i = intGNx) Then                            ' 最終グループ ?
                            ' Data = Data + (チップサイズX * グループ内(サーキット内)抵抗数)
                            dData = dData + (typPlateInfo.dblChipSizeXDir * typPlateInfo.intResistCntInGroup)
                        Else
                            ' Data = Data + (チップサイズX * グループ内(サーキット内)抵抗数 + ＢＰグループ(サーキット)間隔)
                            dData = dData + (typPlateInfo.dblChipSizeXDir * typPlateInfo.intResistCntInGroup + typPlateInfo.dblBpGrpItv)
                        End If
                    Next i
                End If

                ' ブロックサイズX,Yを返す
                dblBSX = dData                                          ' ブロックサイズX = 計算値
                dblBSY = typPlateInfo.dblChipSizeYDir                   ' ブロックサイズY = チップサイズY

            Else
                ' 抵抗(ﾁｯﾌﾟ)並び方向 = Y方向の場合
                If (intGNY = 1) Then
                    ' 1グループ(1サーキット)の場合
                    dData = typPlateInfo.dblChipSizeYDir * intChipNum   ' Data = チップサイズY * チップ数

                Else
                    ' 複数グループ(複数サーキット)の場合
                    For i = 1 To intGNY
                        If (i = intGNY) Then                            ' 最終グループ ?
                            ' Data = Data + (チップサイズY * グループ内(サーキット内)抵抗数)
                            dData = dData + (typPlateInfo.dblChipSizeYDir * typPlateInfo.intResistCntInGroup)
                        Else
                            ' Data = Data + (チップサイズY * グループ内(サーキット内)抵抗数 + ＢＰグループ(サーキット)間隔)
                            dData = dData + (typPlateInfo.dblChipSizeYDir * typPlateInfo.intResistCntInGroup + typPlateInfo.dblBpGrpItv)
                        End If
                    Next i

                End If

                ' ブロックサイズX,Yを返す
                dblBSX = typPlateInfo.dblChipSizeXDir                   ' ブロックサイズX = チップサイズX
                dblBSY = dData                                          ' ブロックサイズY = 計算値

                'V2.0.0.0①↓ステップ位置
                If (giAppMode = APP_MODE_TY) Then
                    dblBSY = Math.Abs(dData)                                          ' ブロックサイズY = 計算値
                End If
                'V2.0.0.0①↑
            End If

            ' トラップエラー発生時 
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try

    End Sub

#End Region

#Region "BPをﾌﾞﾛｯｸの右上に移動する"
    '''=========================================================================
    '''<summary>BPをﾌﾞﾛｯｸの右上に移動する</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub BpMoveOrigin_Ex()
        Try
            Dim dblBpOffsX As Double
            Dim dblBpOffsY As Double
            Dim dblBSX As Double
            Dim dblBSY As Double

            ' ﾌﾞﾛｯｸｻｲｽﾞ取得
            Call CalcBlockSize(dblBSX, dblBSY)
            ' BP位置ｵﾌｾｯﾄX,Y設定
            dblBpOffsX = typPlateInfo.dblBpOffSetXDir
            dblBpOffsY = typPlateInfo.dblBpOffSetYDir
            ' BPをﾌﾞﾛｯｸの右上に移動する(BSIZE()後BPOFF()実行)
            Call Form1.System1.BpMoveOrigin(gSysPrm, dblBSX, dblBSY, dblBpOffsX, dblBpOffsY)
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub
#End Region

#Region "指定抵抗番号、ｶｯﾄ番号のｽﾀｰﾄﾎﾟｲﾝﾄを返す"
    '''=========================================================================
    '''<summary>指定抵抗番号、ｶｯﾄ番号のｽﾀｰﾄﾎﾟｲﾝﾄを返す</summary>
    '''<param name="intRegNo">(INP) 抵抗番号</param>
    '''<param name="intCutNo">(INP) ｶｯﾄ番号</param>
    '''<param name="dblX"    >(OUT) ｽﾀｰﾄﾎﾟｲﾝﾄX</param>
    '''<param name="dblY"    >(OUT) ｽﾀｰﾄﾎﾟｲﾝﾄY</param>
    '''<returns>TRUE:ﾃﾞｰﾀあり, FALSE:ﾃﾞｰﾀなし</returns>
    '''=========================================================================
    Public Function GetCutStartPoint(ByRef intRegNo As Short, ByRef intCutNo As Short, ByRef dblX As Double, ByRef dblY As Double) As Boolean
        Try
            Dim bRetc As Boolean
            Dim i As Short
            Dim j As Short

            bRetc = False
            For i = 1 To MAXRNO
                If (intRegNo = typResistorInfoArray(i).intResNo) Then                       ' 抵抗番号一致
                    For j = 1 To MaxCntCut
                        If (intCutNo = typResistorInfoArray(i).ArrCut(j).intCutNo) Then     ' ｶｯﾄ番号一致
                            dblX = typResistorInfoArray(i).ArrCut(j).dblStartPointX         ' ｽﾀｰﾄﾎﾟｲﾝﾄ
                            dblY = typResistorInfoArray(i).ArrCut(j).dblStartPointY
                            bRetc = True
                            GetCutStartPoint = bRetc
                            Exit Function
                        End If
                    Next
                End If
            Next
            GetCutStartPoint = bRetc
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region

#Region "自動レーザパワー調整処理"
    '''=========================================================================
    ''' <summary>
    ''' 自動レーザパワー調整処理
    ''' </summary>
    ''' <param name="bPowerMonitoring">True:フルパワー測定</param>
    ''' <returns>cFRS_NORMAL  = 正常,cFRS_ERR_RST = Cancel(RESETｷｰ),上記以外 　　= 非常停止検出等のエラー</returns>
    ''' <remarks>自動レーザパワーの調整処理実行</remarks>
    '''=========================================================================
    Public Function AutoLaserPowerADJ(Optional ByVal bPowerMonitoring As Boolean = False) As Short

        Dim r As Integer

        Try
            Dim strMsg As String

            With stLASER
                ' パワーメータのデータ取得タイプがステージ設置タイプでなく「Ｉ／Ｏ読取り」/「ＵＳＢ」でなければNOP(そのまま抜ける)

                If (gSysPrm.stIOC.giPM_Tp <> 1 Or gSysPrm.stIOC.giPM_DataTp = PM_DTTYPE_NONE) Then
                    Return (cFRS_NORMAL)
                End If

                ' パワー調整実行フラグ
                If Not bPowerMonitoring Then
                    If (.intPowerAdjustMode <> 1) Then
                        ' ﾊﾟﾜｰ調整を実行しない場合はそのまま抜ける
                        Return (cFRS_NORMAL)
                    End If
                End If

                ' Zを原点へ移動
                r = EX_ZMOVE(0)
                If (r <> cFRS_NORMAL) Then                              ' エラー ?(メッセージは表示済み) 
                    Return (r)                                          ' Return値設定 
                End If

                '---------------------------------------------------------------
                '   自動パワー調整実行
                '---------------------------------------------------------------
                If (gSysPrm.stRAT.giOsc_Res = OSCILLATOR_FL) Then
#If cOSCILLATORcFLcUSE Then
                    Dim iCurr As Integer
                    Dim iCurrOfs As Integer
                    Dim dMeasPower As Double
                    Dim dFullPower As Double
                    Dim AdjustTarget As Double
                    Dim AdjustLevel As Double
                    Dim CndNum As Integer
                    '-----------------------------------------------------------
                    '   FL時
                    '-----------------------------------------------------------

                    ' パワー調整する加工条件番号配列に有効/無効を設定する
                    r = SetAutoPowerCndNumAry(stPWR)

                    ' カットに使用する加工条件番号のパワー調整を行う 
                    For CndNum = 0 To (MAX_BANK_NUM - 1)
                        If (stPWR.CndNumAry(CndNum) = 1) Then               ' 加工条件は有効 ?
                            AdjustTarget = stPWR.AdjustTargetAry(CndNum)    ' 目標パワー値(W)
                            AdjustLevel = stPWR.AdjustLevelAry(CndNum)      ' 調整許容範囲(±W)

                            ' メッセージ表示("パワー調整開始"+ " 加工条件番号xx")
                            strMsg = MSG_AUTOPOWER_01 + " " + MSG_AUTOPOWER_02 + CndNum.ToString("00")
                            Call Z_PRINT(strMsg)

                            ' パワー調整を行う
                            r = Form1.System1.Form_FLAutoLaser(gSysPrm, CndNum, AdjustTarget, AdjustLevel, iCurr, iCurrOfs, dMeasPower, dFullPower)
                            If (r < cFRS_NORMAL) Then
                                ' エラーメッセージ表示
                                r = Form1.System1.Form_AxisErrMsgDisp(System.Math.Abs(r))
                                Return (r)
                            End If

                            ' 調整結果をメイン画面に表示する
                            If (r = cFRS_NORMAL) Then                   ' 正常終了 ? 
                                ' メッセージ表示("レーザパワー設定値"+" = xx.xxW, " + "電流値=" + "xxxmA")
                                strMsg = MSG_AUTOPOWER_03 + "= " + dMeasPower.ToString("0.00") + "W, "
                                strMsg = strMsg + MSG_AUTOPOWER_04 + "= " + iCurr.ToString("0") + "mA"
                                Call Z_PRINT(strMsg)
                                stCND.Curr(CndNum) = iCurr              ' 電流値設定
                            Else
                                ' メッセージ表示("パワー調整未完了")
                                strMsg = MSG_AUTOPOWER_05
                                Call Z_PRINT(strMsg)
                                Exit For                                ' 処理終了
                            End If
                        End If
                    Next CndNum

#End If
                Else
                    '-----------------------------------------------------------
                    '   FL以外の場合
                    '-----------------------------------------------------------
                    r = Form1.System1.Form_AutoLaser(gSysPrm, .dblPowerAdjustQRate,
                                        .dblPowerAdjustTarget, .dblPowerAdjustToleLevel, bPowerMonitoring)

                    If (r = cFRS_NORMAL) Then                           ' 正常終了 ? 
                        ' メッセージ表示("パワー調整正常終了")
                        strMsg = MSG_AUTOPOWER_06
                        Call Z_PRINT(strMsg)
                    Else
                        ' メッセージ表示("パワー調整未完了")
                        strMsg = MSG_AUTOPOWER_05                       ' "パワー調整未完了"
                        If (bPowerMonitoring = True) Then               ' レーザパワーのモニタリング ?
                            strMsg = MSG_165                            ' "レーザパワーモニタリング異常終了"
                        End If
                        Call Z_PRINT(strMsg)
                    End If
                End If

                System.Windows.Forms.Application.DoEvents()
            End With
            Return (r)                                                  ' Return値設定 

            ' トラップエラー発生時 
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
            Return (cERR_TRAP)                                          ' Return値 = トラップエラー発生
        End Try
    End Function
#End Region

#Region "パワー調整する加工条件番号配列に有効/無効を設定する"
#If cOSCILLATORcFLcUSE Then
    '''=========================================================================
    ''' <summary>パワー調整する加工条件番号配列に有効/無効を設定する</summary>
    ''' <param name="stPWR">(OUT)FL用パワー調整情報
    '''                              ※配列は0オリジン</param>
    ''' <remarks>自動レーザパワーの調整処理実行用</remarks>
    ''' <returns>cFRS_NORMAL  = 正常
    '''          上記以外 　　= エラー</returns> 
    '''=========================================================================
    Private Function SetAutoPowerCndNumAry(ByRef stPWR As POWER_ADJUST_INFO) As Short

        Dim Rn As Integer
        Dim Cn As Integer
        Dim CndNum As Integer
        Dim CutType As Integer

        Try
            '------------------------------------------------------------------
            '   初期処理
            '------------------------------------------------------------------
            ' FLでなければNOP
            If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then Return (cFRS_NORMAL)

            ' トリマー加工条件構造体(FL用)初期化 
            stPWR.Initialize()

            '------------------------------------------------------------------
            '   加工条件番号配列を設定する
            '------------------------------------------------------------------
            For Rn = 1 To stPLT.RCount              ' １ブロック内抵抗数分チェックする 
                If UserModule.IsCutResistorIncMarking(Rn) Then
                    For Cn = 1 To stREG(Rn).intTNN      ' 抵抗内カット数分チェックする
                        ' カットタイプ取得
                        CutType = stREG(Rn).STCUT(Cn).intCTYP

                        ' 加工条件1は全カット無条件に設定する
                        CndNum = stREG(Rn).STCUT(Cn).intCND(CUT_CND_L1)
                        If (stPWR.CndNumAry(CndNum) = 0) Then               ' 無効 ? 
                            stPWR.CndNumAry(CndNum) = 1                     ' 有効に設定
                            ' 目標パワー値(W)と調整許容範囲(±W)を設定する
                            stPWR.AdjustTargetAry(CndNum) = stCND.dblPowerAdjustTarget(CndNum)
                            stPWR.AdjustLevelAry(CndNum) = stCND.dblPowerAdjustToleLevel(CndNum)
                        End If

                        ' 加工条件2はLカット, 斜めLカット, Lカット(ﾘﾀｰﾝ/ﾘﾄﾚｰｽ), 斜めLカット(ﾘﾀｰﾝ/ﾘﾄﾚｰｽ)
                        ' HOOKカット, Uカット時に設定する
                        If (CutType = CNS_CUTP_L) Or (CutType = CNS_CUTP_NL) Or _
                           (CutType = CNS_CUTP_Lr) Or (CutType = CNS_CUTP_Lt) Or _
                           (CutType = CNS_CUTP_NLr) Or (CutType = CNS_CUTP_NLt) Or _
                           (CutType = CNS_CUTP_HK) Or (CutType = CNS_CUTP_U) Then
                            ' 加工条件2
                            CndNum = stREG(Rn).STCUT(Cn).intCND(CUT_CND_L2)
                            If (stPWR.CndNumAry(CndNum) = 0) Then           ' 無効 ? 
                                stPWR.CndNumAry(CndNum) = 1                 ' 有効に設定
                                ' 目標パワー値(W)と調整許容範囲(±W)を設定する
                                stPWR.AdjustTargetAry(CndNum) = stCND.dblPowerAdjustTarget(CndNum)
                                stPWR.AdjustLevelAry(CndNum) = stCND.dblPowerAdjustToleLevel(CndNum)
                            End If
                        End If

                        ' 加工条件3はHOOKカット, Uカット時に設定する
                        If (CutType = CNS_CUTP_HK) Or (CutType = CNS_CUTP_U) Then
                            ' 加工条件3
                            CndNum = stREG(Rn).STCUT(Cn).intCND(CUT_CND_L3)
                            If (stPWR.CndNumAry(CndNum) = 0) Then           ' 無効 ? 
                                stPWR.CndNumAry(CndNum) = 1                 ' 有効に設定
                                ' 目標パワー値(W)と調整許容範囲(±W)を設定する
                                stPWR.AdjustTargetAry(CndNum) = stCND.dblPowerAdjustTarget(CndNum)
                                stPWR.AdjustLevelAry(CndNum) = stCND.dblPowerAdjustToleLevel(CndNum)
                            End If
                        End If

                        ' 加工条件4は現状は未使用(予備)

                        ' 加工条件5～8はリターン/リトレース用 
                        ' 加工条件5(STカット(ﾘﾀｰﾝ/ﾘﾄﾚｰｽ), 斜めSTカット(ﾘﾀｰﾝ/ﾘﾄﾚｰｽ)時
                        If (CutType = CNS_CUTP_STr) Or (CutType = CNS_CUTP_STt) Or _
                           (CutType = CNS_CUTP_NSTr) Or (CutType = CNS_CUTP_NSTt) Then
                            ' 加工条件5の条件番号をカットデータの加工条件2より設定する
                            CndNum = stREG(Rn).STCUT(Cn).intCND(CUT_CND_L1)
                            If (stPWR.CndNumAry(CndNum) = 0) Then           ' 無効 ? 
                                stPWR.CndNumAry(CndNum) = 1                 ' 有効に設定
                                ' 目標パワー値(W)と調整許容範囲(±W)を設定する
                                stPWR.AdjustTargetAry(CndNum) = stCND.dblPowerAdjustTarget(CndNum)
                                stPWR.AdjustLevelAry(CndNum) = stCND.dblPowerAdjustToleLevel(CndNum)
                            End If
                        End If

                        ' 加工条件5,6(Lカット(ﾘﾀｰﾝ/ﾘﾄﾚｰｽ), 斜めLカット(ﾘﾀｰﾝ/ﾘﾄﾚｰｽ)時
                        If (CutType = CNS_CUTP_Lr) Or (CutType = CNS_CUTP_Lt) Or _
                           (CutType = CNS_CUTP_NLr) Or (CutType = CNS_CUTP_NLt) Then
                            ' 加工条件5の条件番号をカットデータの加工条件3より設定する
                            CndNum = stREG(Rn).STCUT(Cn).intCND(CUT_CND_L2)
                            If (stPWR.CndNumAry(CndNum) = 0) Then           ' 無効 ? 
                                stPWR.CndNumAry(CndNum) = 1                 ' 有効に設定
                                ' 目標パワー値(W)と調整許容範囲(±W)を設定する
                                stPWR.AdjustTargetAry(CndNum) = stCND.dblPowerAdjustTarget(CndNum)
                                stPWR.AdjustLevelAry(CndNum) = stCND.dblPowerAdjustToleLevel(CndNum)
                            End If
                            ' 加工条件6の条件番号をカットデータの加工条件4より設定する
                            CndNum = stREG(Rn).STCUT(Cn).intCND(CUT_CND_L3)
                            If (stPWR.CndNumAry(CndNum) = 0) Then           ' 無効 ? 
                                stPWR.CndNumAry(CndNum) = 1                 ' 有効に設定
                                ' 目標パワー値(W)と調整許容範囲(±W)を設定する
                                stPWR.AdjustTargetAry(CndNum) = stCND.dblPowerAdjustTarget(CndNum)
                                stPWR.AdjustLevelAry(CndNum) = stCND.dblPowerAdjustToleLevel(CndNum)
                            End If
                        End If

                        ' 加工条件7,8は現状は未使用(予備)


                    Next Cn
                End If
            Next Rn

            Return (cFRS_NORMAL)                                        ' Return値設定 

            ' トラップエラー発生時 
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
            Return (cERR_TRAP)                                          ' Return値 = トラップエラー発生
        End Try
    End Function
#End If
#End Region

    '#Region "DispGazou.exe処理"
    '=========================================================================
    '   画像表示プログラムの起動処理
    '=========================================================================

#Region "DispgazouにWindowメッセージを送信する"

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Public Function FindWindow( _
         ByVal lpClassName As String, _
         ByVal lpWindowName As String) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Public Function SendNotifyMessage( _
                           ByVal hWnd As IntPtr, _
                           ByVal wMsg As Int32, _
                           ByVal wParam As Int32, _
                           ByVal lParam As Int32) As Integer
    End Function

    Private Const WM_APP As Int32 = &H8000
    '    '''=========================================================================
    '    ''' <summary>DispgazouにWindowメッセージを送信する</summary>
    '    ''' <param name="No">(INP)メッセージ番号</param>
    '    ''' <returns></returns>
    '    ''' <remarks></remarks>
    '    '''=========================================================================
    '    Public Function SendMsgToDispGazou(ByRef ObjProc As Process, ByVal No As Integer) As Integer

    '        Dim result As Integer = cFRS_NORMAL
    '        Dim Cnt As Integer = 0
    '        Dim hWnd As Int32
    '        Try
    'SND_MSG_RETRY_START:
    '            '相手のウィンドウハンドルを取得します
    '            hWnd = FindWindow(Nothing, "DispGazou") 'V4.3.0.0③
    '            If hWnd = 0 Then
    '                'ハンドルが取得できなかった
    '                Cnt = Cnt + 1
    '                If Cnt < 100 Then
    '                    If (Cnt Mod 10) = 0 Then
    '                        Call FinalEnd_GazouProc(ObjProc)                                'DispGazou強制終了
    '                        'V2.2.0.0① Execute_GazouProc(ObjProc, DISPGAZOU_PATH, DISPGAZOU_WRK, INTERNAL_CAMERA)    '再起動
    '                    End If
    '                    System.Threading.Thread.Sleep(10)
    '                    GoTo SND_MSG_RETRY_START
    '                Else
    '                    MessageBox.Show("相手Windowのハンドルが取得できません")
    '                End If
    '            End If

    '            result = SendNotifyMessage(hWnd, WM_APP, 0, No)

    '            Return result
    '        Catch ex As Exception
    '            MsgBox("UserModule.SendMsgToDispGazou() TRAP ERROR = " + ex.Message)
    '        End Try
    '    End Function
    '#End Region

    '#Region "画像表示プログラムを起動する"
    '    '''=========================================================================
    '    ''' <summary>画像表示プログラムを起動する</summary>
    '    ''' <param name="ObjProc"> (OUT)Processｵﾌﾞｼﾞｪｸﾄ</param>
    '    ''' <param name="strFName">(INP)起動プログラム名</param>
    '    ''' <param name="Camera">  (INP)カメラ番号(0-3)</param> 
    '    ''' <returns>0 = 正常, 0以外 = エラー</returns>
    '    '''=========================================================================
    '    Public Function Execute_GazouProc(ByRef ObjProc As Process, ByRef strFName As String, ByRef strWrk As String, ByVal Camera As Integer) As Integer

    '        Dim strARG As String                                        ' 引数() 

    '        Dim dispXPos As Integer
    '        Dim dispYPos As Integer
    '        Dim Cnt As Integer = 0

    '        Try
    '            TrimClassCommon.ForceEndProcess(DISPGAZOU_PATH)       ' プロセスを強制終了する。

    '            ' 表示位置設定
    '            dispXPos = FORM_X + Form1.VideoLibrary1.Location.X
    '            dispYPos = FORM_Y + Form1.VideoLibrary1.Location.Y

    '            ' ｺﾏﾝﾄﾞﾗｲﾝ引数設定
    '            strARG = Camera.ToString("0") + " "                     ' args[0] :カメラ番号(0-3)"
    '            'strARG = "0 "                                           ' args[0] :カメラ番号(0-3)"
    '            strARG = strARG + "1 "                                  ' args[1] :(0=ボタン表示する, 1=ボタン表示しない)
    '            strARG = strARG + dispXPos.ToString("0") + " "          ' args[2] :フォームの表示位置X
    '            strARG = strARG + dispYPos.ToString("0")                ' args[3] :フォームの表示位置Y
    '            strARG = strARG + " 1"                                  ' args[4] :(0=メッセージ制御無し, 1=メッセージ制御有り)
    '            strARG = strARG + " 1"                                  ' args[5] :(0=シンプルトリマ用サイズ小画面, 1=通常画面サイズ)

    '            ' プロセスの起動
    '            ObjProc = New Process                                   ' Processｵﾌﾞｼﾞｪｸﾄを生成する 
    '            ObjProc.StartInfo.FileName = strFName                   ' プロセス名 
    '            ObjProc.StartInfo.Arguments = strARG                    ' ｺﾏﾝﾄﾞﾗｲﾝ引数設定
    '            ObjProc.StartInfo.WorkingDirectory = strWrk             ' 作業フォルダ
    '            ObjProc.Start()                                         ' プロセス起動

    '            ' チャネルを登録
    '            'ChannelServices.RegisterChannel(ipcChnl, False)
    'IPC_RETRY_START:  ' サーバ（DispGazou)側を停止して直ぐに起動した後だとポートに書き込めないエラーになる。対処方法が判らないので再試行する。
    '            Try
    '                'refObj.CallServer("STOP")
    '                'V2.0.0.3①                System.Threading.Thread.Sleep(2000)
    '                System.Threading.Thread.Sleep(500)
    '                SendMsgToDispGazou(ObjProc, 2)       'STOP 'V4.0.0.0-87
    '            Catch ex As Exception
    '                Cnt = Cnt + 1
    '                If Cnt < 100 Then
    '                    If (Cnt Mod 10) = 0 Then
    '                        Call FinalEnd_GazouProc(ObjProc)                                'DispGazou強制終了
    '                        'V2.2.0.0①  Execute_GazouProc(ObjProc, DISPGAZOU_PATH, DISPGAZOU_WRK, Camera)    '再起動
    '                    End If
    '                    System.Threading.Thread.Sleep(10)
    '                    GoTo IPC_RETRY_START
    '                Else
    '                    MsgBox("UserModule.Execute_GazouProc() TRAP ERROR = " + ex.Message)
    '                End If
    '            End Try


    '            ' トラップエラー発生時 
    '        Catch ex As Exception
    '            MsgBox("UserModule.Execute_GazouProc() TRAP ERROR = " + ex.Message)
    '            Return (cERR_TRAP)
    '        End Try
    '    End Function
    '#End Region

    '#Region "画像表示プログラムを強制終了する"
    '    '''=========================================================================
    '    '''<summary>画像表示プログラムを強制終了する</summary>
    '    '''<param name="ObjProc"> (OUT)Processｵﾌﾞｼﾞｪｸﾄ</param>
    '    '''<returns>0 = 正常, 0以外 = エラー</returns>
    '    '''=========================================================================
    '    Public Function FinalEnd_GazouProc(ByRef ObjProc As Process) As Integer
    '        Try
    '            TrimClassCommon.ForceEndProcess(DISPGAZOU_PATH)       ' ダメ押しでプロセスを強制終了する。

    '            Return (cFRS_NORMAL)

    '            ' トラップエラー発生時 
    '        Catch ex As Exception
    '            '            MsgBox("basTrimming.FinalEnd_GazouProc() TRAP ERROR = " + ex.Message)
    '            Return (cERR_TRAP)
    '        End Try
    '    End Function
    '#End Region

    '#Region "画像表示プログラムを起動する"
    '    '''=========================================================================
    '    ''' <summary>
    '    ''' 画像表示プログラムを起動する
    '    ''' </summary>
    '    ''' <param name="ObjProc"> (OUT)Processｵﾌﾞｼﾞｪｸﾄ</param>
    '    ''' <param name="strFName">(INP)起動プログラム名</param>
    '    ''' <param name="strWrk">(INP)作業フォルダ</param>
    '    ''' <param name="Camera">(INP)カメラ番号(0-3)</param>
    '    ''' <returns></returns>
    '    ''' <remarks></remarks>
    '    '''=========================================================================
    '    Public Function Exec_GazouProc(ByRef ObjProc As Process, ByRef strFName As String, ByRef strWrk As String, ByVal Camera As Integer) As Integer

    '        Dim Cnt As Integer = 0
    '        ' Dim result As Integer

    '        Try
    '            'If Form1.GetDistributeOnOffStatus() Then
    '            '    Return (cFRS_NORMAL)
    '            'End If
    '            ' VideoOcx表示を停止
    '            Call Form1.VideoLibrary1.VideoStop()

    'IPC_RETRY_START:  ' サーバ（DispGazou)側を停止して直ぐに起動した後だとポートに書き込めないエラーになる。対処方法が判らないので再試行する。
    '            Try
    '                SendMsgToDispGazou(ObjProc, 5)                           ' START_NORMAL
    '            Catch ex As Exception
    '                Cnt = Cnt + 1
    '                If Cnt < 100 Then
    '                    If (Cnt Mod 10) = 0 Then
    '                        Call FinalEnd_GazouProc(ObjProc)                                'DispGazou強制終了
    '                        System.Threading.Thread.Sleep(100)
    '                        'V2.2.0.0①                         Execute_GazouProc(ObjProc, strFName, strWrk, Camera)    '再起動
    '                    End If
    '                    System.Threading.Thread.Sleep(10)
    '                    GoTo IPC_RETRY_START
    '                Else
    '                    MsgBox("UserModule.Exec_GazouProc() TRAP ERROR = " + ex.Message)
    '                End If
    '            End Try

    '            Return (cFRS_NORMAL)
    '            ' トラップエラー発生時 
    '        Catch ex As Exception
    '            MsgBox("UserModule.Exec_GazouProc() TRAP ERROR = " + ex.Message)
    '            Return (cERR_TRAP)
    '        End Try
    '    End Function
    '#End Region

    '#Region "画像表示プログラムを強制終了する"
    '    '''=========================================================================
    '    '''<summary>画像表示プログラムを強制終了する</summary>
    '    '''<param name="ObjProc"> (OUT)Processｵﾌﾞｼﾞｪｸﾄ</param>
    '    '''<returns>0 = 正常, 0以外 = エラー</returns>
    '    '''=========================================================================
    '    Public Function End_GazouProc(ByRef ObjProc As Process) As Integer

    '        Dim Cnt As Integer = 0
    '        '        Dim result As Integer 

    '        Try

    'IPC_RETRY_START:  ' サーバ（DispGazou)側を停止して直ぐに起動した後だとポートに書き込めないエラーになる。対処方法が判らないので再試行する。
    '            Try

    '                SendMsgToDispGazou(ObjProc, 2)       'STOP 

    '            Catch ex As Exception
    '                Cnt = Cnt + 1
    '                If Cnt < 100 Then
    '                    If (Cnt Mod 10) = 0 Then
    '                        Call FinalEnd_GazouProc(ObjProc)                                'DispGazou強制終了
    '                        'V2.2.0.0①                         Execute_GazouProc(ObjProc, DISPGAZOU_PATH, DISPGAZOU_WRK, INTERNAL_CAMERA)    '再起動
    '                    End If
    '                    System.Threading.Thread.Sleep(10)
    '                    GoTo IPC_RETRY_START
    '                Else
    '                    MsgBox("UserModule.End_GazouProc() TRAP ERROR = " + ex.Message)
    '                End If
    '            End Try

    '            Call Form1.VideoLibrary1.VideoStart()

    '            ' 画面を更新
    '            Call Form1.Refresh()

    '            Return (cFRS_NORMAL)

    '            ' トラップエラー発生時 
    '        Catch ex As Exception
    '            MsgBox("UserModule.End_GazouProc() TRAP ERROR = " + ex.Message)
    '            Return (cERR_TRAP)
    '        End Try
    '    End Function
    '#End Region

#End Region

#Region "ビームポジショナのエイジング"
    ''' <summary>
    ''' ビームポジショナのエイジング・ガルバを最大まで動かしてエイジングする。
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub BeamPositionerAging()
        Try
            Dim r As Integer

            Call BSIZE(80, 80)                                                                      '   ブロックサイズ０設定
            r = ObjSys.EX_MOVE(gSysPrm, 0, 0, 1)                                                    '   BP(0,0)へ移動
            If (r < cFRS_NORMAL) Then                                                               ' 
                Call Z_PRINT("ＢＰ移動エラーが発生しました。EX_MOVE(gSysPrm, 0, 0, 1)" + vbCrLf)    ' 
            End If                                                                                  ' 
            r = ObjSys.EX_MOVE(gSysPrm, 80, 80, 1)                                                  '   BP(80,80)へ移動
            If (r < cFRS_NORMAL) Then                                                               ' 
                Call Z_PRINT("ＢＰ移動エラーが発生しました。EX_MOVE(gSysPrm, 80, 80, 1)" + vbCrLf)  ' 
            End If                                                                                  ' 
            r = ObjSys.EX_MOVE(gSysPrm, 40, 40, 1)                                                  '   BP(80,80)へ移動
            If (r < cFRS_NORMAL) Then                                                               ' 
                Call Z_PRINT("ＢＰ移動エラーが発生しました。EX_MOVE(gSysPrm, 40, 40, 1)" + vbCrLf)  ' 
            End If                                                                                  ' 
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub
#End Region

#Region "ブロック内の２点補正によるＸＹ補正の実行可否"
    Private bBlockXYCorrection As Boolean = False
    ''' <summary>
    ''' シスパラの設定からのブロック内の２点補正によるＸＹ補正の有り設定
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SetBlockXYCorrectionOn()
        Try
            bBlockXYCorrection = True
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub
    ''' <summary>
    ''' ブロック内の２点補正によるＸＹ補正を使用の有無を取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsBlockXYCorrectionUse() As Boolean
        Try
            If Not bBlockXYCorrection Then
                Return (False)
            End If

            ' θ補正処理
            If (gSysPrm.stDEV.giTheta > 0) Then                         ' ＸＹθ有りなら実行しない
                Return (False)
            End If

            Return (True)

        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
    ''' <summary>
    ''' ブロック内の２点補正によるＸＹ補正の実行可否
    ''' </summary>
    ''' <param name="AppMode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsBlockXYCorrection(ByRef AppMode As Short) As Boolean
        Try
            If Not bBlockXYCorrection Then
                Return (False)
            End If

            ' θ補正処理
            If (gSysPrm.stDEV.giTheta > 0) Then                         ' ＸＹθ有りなら実行しない
                Return (False)
            End If

            If stThta.iPP31 = 0 Then                                    ' 補正なしなら実行しない
                Return (False)
            End If

            If AppMode <> APP_MODE_TRIM Then                            ' トリミング時以外は実行しない
                Return (False)
            End If

            Return (True)

        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region

#Region "ブロック内の２点補正によるＸＹ補正"
    ''' <summary>
    ''' ブロック内の２点補正によるＸＹ補正
    ''' </summary>
    ''' <param name="AppMode">トリミングモード</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function BlockXYCorrection(ByRef AppMode As Short) As Integer
        Dim rtn As Integer = cFRS_NORMAL                                ' Return値
        Dim Thresh1 As Double = 0
        Dim Thresh2 As Double = 0
        Dim strMSG As String                                            ' Display Message
        Dim r As Integer = 0

        Try
            If Not IsBlockXYCorrection(AppMode) Then                    ' ブロック内の２点補正によるＸＹ補正の実行可否
                Return (cFRS_NORMAL)
            End If

            Call ObjSys.Ilum_Ctrl(gSysPrm, Z1, ZOPT)                    ' ｲﾙﾐﾈｰｼｮﾝﾗﾝﾌﾟ点灯(制御有時有効)

            Call InitThetaCorrection()                                  ' パターン登録初期値設定

            ' カメラ切替
            If (gSysPrm.stDEV.giCutPic = 0) Then                        ' VGAボードあり
                ObjVdo.VideoStop()                                      ' ビデオストップ/スタート(内部/外部ｶﾒﾗ)
                Call ObjVdo.VideoStart2(gSysPrm.stDEV.giEXCAM)
            Else
                If (gSysPrm.stDEV.giEXCAM_Usr = 1) Then                 ' 外部カメラを使用？
                    Call ObjVdo.ChangeCamera(EXTERNAL_CAMERA)            ' カメラ切替(外部ｶﾒﾗ)
                End If
            End If
#If cLEDcILLUMINATION Then
            If stUserData.iLEDIllumination = ELD_USE_ONLY Then          '「使用時のみＯＮ」
                UserSub.LEDLight_On()
            End If
#End If

            ' θ補正処理
            ObjVdo.frmTop = Form1.Text2.Location.Y                      ' 補正画面表示位置設定
            ObjVdo.frmLeft = Form1.Text2.Location.X
            r = ObjVdo.CorrectTheta(APP_MODE_BLOCK_RECOG)               ' θ補正

            ' XYテーブル補正値(θ補正時のXYﾃｰﾌﾞﾙずれ量)取得
            If (r = 0) Then
                dblCorrectX = ObjVdo.CorrectTrimPosX
                dblCorrectY = ObjVdo.CorrectTrimPosY
            Else
                dblCorrectX = 0
                dblCorrectY = 0
            End If

            ' 後処理
            If (gSysPrm.stDEV.giCutPic = 0) Then                        ' VGAボードあり?
                ObjVdo.VideoStop()                                      ' ビデオライブラリストップ
                ObjMain.Refresh()
            Else
                If (gSysPrm.stDEV.giEXCAM_Usr = 1) Then                 ' 外部カメラを使用？
                    Call ObjVdo.ChangeCamera(INTERNAL_CAMERA)           ' カメラ切替(内部ｶﾒﾗ)
                End If
            End If
#If cLEDcILLUMINATION Then
            If stUserData.iLEDIllumination = ELD_USE_ONLY Then          '「使用時のみＯＮ」
                UserSub.LEDLight_Off()
            End If
#End If
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
            If (gSysPrm.stTMN.giMsgTyp = 0) Then
                strMSG = "補正位置1X,Y =" & stThta.fpp32_x.ToString("0.0000").PadLeft(9) & ","
                strMSG = strMSG & stThta.fpp32_y.ToString("0.0000").PadLeft(9) & "  "
                strMSG = strMSG & " ずれ量1X,Y   =" & stResult.fCor1x.ToString("0.0000").PadLeft(9) & ","
                strMSG = strMSG & stResult.fCor1y.ToString("0.0000").PadLeft(9) & vbCrLf
                strMSG = strMSG & "補正位置2X,Y =" & stThta.fpp33_x.ToString("0.0000").PadLeft(9) & ","
                strMSG = strMSG & stThta.fpp33_y.ToString("0.0000").PadLeft(9) & "  "
                strMSG = strMSG & " ずれ量2X,Y   =" & stResult.fCor2x.ToString("0.0000").PadLeft(9) & ","
                strMSG = strMSG & stResult.fCor2y.ToString("0.0000").PadLeft(9) & vbCrLf
                'If (stThta.iPP30 = 0) Then                              ' 自動補正モード ?
                strMSG = strMSG & "  一致度POS1   =" & stResult.fCorV1.ToString("0.0000").PadLeft(9) & ","
                strMSG = strMSG & " 一致度POS2   =" & stResult.fCorV2.ToString("0.0000").PadLeft(9) & vbCrLf
                'End If
            Else
                strMSG = "  Correct position1=" & stThta.fpp32_x.ToString("0.0000").PadLeft(9) & ","
                strMSG = strMSG & stThta.fpp32_y.ToString("0.0000").PadLeft(9) & "  "
                strMSG = strMSG & " Distance1=" & stResult.fCor1x.ToString("0.0000").PadLeft(9) & ","
                strMSG = strMSG & stResult.fCor1y.ToString("0.0000").PadLeft(9) & vbCrLf
                strMSG = strMSG & "  Correct position2=" & stThta.fpp33_x.ToString("0.0000").PadLeft(9) & ","
                strMSG = strMSG & stThta.fpp33_y.ToString("0.0000").PadLeft(9) & "  "
                strMSG = strMSG & " Distance2=" & stResult.fCor2x.ToString("0.0000").PadLeft(9) & ","
                strMSG = strMSG & stResult.fCor2y.ToString("0.0000").PadLeft(9) & vbCrLf
                'If (stThta.iPP30 = 0) Then                              ' 自動補正モード ?
                strMSG = strMSG & "  Correlation coefficient1=" & stResult.fCorV1.ToString("0.0000").PadLeft(9) & ","
                strMSG = strMSG & " Correlation coefficient2=" & stResult.fCorV2.ToString("0.0000").PadLeft(9) & vbCrLf
                'End If
            End If

            ' θ補正情報表示
            Call Z_PRINT(strMSG)

            ' POS1の閾値のチェックを行う
            If (Thresh1 > stResult.fCorV1) Then
                Call Beep()                                             ' Beep音
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    strMSG = "パターン認識エラー (POS1閾値)"
                Else
                    strMSG = "Pattern Matching Error(POS1 THRESH)"
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
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region

#Region "座標計算モジュール"

    ''' <summary>
    ''' ２点間の距離を求める
    ''' </summary>
    ''' <param name="x"></param>
    ''' <param name="y"></param>
    ''' <param name="x2"></param>
    ''' <param name="y2"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDistance(ByRef x As Double, ByRef y As Double, ByRef x2 As Double, ByRef y2 As Double) As Double
        Try
            Dim distance As Double = Math.Sqrt((x2 - x) * (x2 - x) + (y2 - y) * (y2 - y))
            Return (distance)
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function

    ''' <summary>
    ''' ２点間の角度（ラジアン）を求める
    ''' </summary>
    ''' <param name="x"></param>
    ''' <param name="y"></param>
    ''' <param name="x2"></param>
    ''' <param name="y2"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRadian(ByRef x As Double, ByRef y As Double, ByRef x2 As Double, ByRef y2 As Double) As Double
        Try
            Dim radian As Double = Math.Atan2(y2 - y, x2 - x)
            Return (radian)
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function

    ''' <summary>
    ''' ２点間の角度（度）を求める
    ''' </summary>
    ''' <param name="x"></param>
    ''' <param name="y"></param>
    ''' <param name="x2"></param>
    ''' <param name="y2"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDegree(ByRef x As Double, ByRef y As Double, ByRef x2 As Double, ByRef y2 As Double) As Double
        Try
            Dim radian As Double = GetRadian(x, y, x2, y2)
            Dim degree As Double = radian * 180D / Math.PI
            Return (degree)
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function

    ''' <summary>
    ''' 角度と距離から座標を求める
    ''' </summary>
    ''' <param name="degree"></param>
    ''' <param name="distance"></param>
    ''' <param name="x"></param>
    ''' <param name="y"></param>
    ''' <remarks></remarks>
    Public Sub GetXYfromDegree(ByVal degree As Double, ByVal distance As Double, ByRef x As Double, ByRef y As Double)
        Try
            'θ(degree)にMath.PI / 180を掛けているのはdegreeをradianに変換している
            'radiusは半径である、距離とは半径を求めているのと同じ
            Dim radian As Double = degree * Math.PI / 180
            x = Math.Cos(radian) * distance
            y = Math.Sin(radian) * distance
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub

#Region "回転座標を求める"
    Public Sub XYRotation(ByVal x1 As Double, ByVal y1 As Double, ByVal angle As Double, ByRef x As Double, ByRef y As Double)
        Try
            Dim degrees As Double = Math.PI / 180 * angle
            x = x1 * Math.Cos(degrees) - y1 * Math.Sin(degrees)
            y = x1 * Math.Sin(degrees) + y1 * Math.Cos(degrees)
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub
#End Region

#End Region

#Region "２点の座標データからのずれを補正"
    Sub HoseiCalc(ByVal bStageMode As Boolean, ByVal A1 As Double, ByVal B1 As Double, ByVal A2 As Double, ByVal B2 As Double, ByVal C1 As Double, ByVal D1 As Double, ByVal C2 As Double, ByVal D2 As Double, ByVal x1 As Double, ByVal y1 As Double, ByRef x As Double, ByRef y As Double)
        Try
            ' (x - a) / (a' - a) = (X - A)/(A'-A)
            ' (y - b) / (b' - b) = (Y - B)/(B'-B)
            ' x = a + (X - A) / (A'-A) * (a' - a)
            ' y = b + (Y - B) / (B'-B) * (b' - b)
            ' (X,Y) → (x,y)
            ' (a,b) → (A1,B1) (a',b') → (A2,B2) 
            ' (A,B) → (C1,D1) (A',B') → (C2,D2) 
            ' 元の２点 (A1,B1) (A2,B2)
            ' パターン認識で得られた２点  (C1,D1) (C2,D2) 
            ' 求めたい座標(x1,y1)
            ' 補正された座標(x,y)

            'If (A1 = A2) Or (B1 = B2) Then
            '    ' ４５度回転させて計算してから元に戻す
            '    Dim RA1 As Double, RB1 As Double, RA2 As Double, RB2 As Double, RC1 As Double, RD1 As Double, RC2 As Double, RD2 As Double, Rx1 As Double, Ry1 As Double
            '    XYRotation(A1, B1, 45, RA1, RB1)
            '    XYRotation(A2, B2, 45, RA2, RB2)
            '    XYRotation(C1, D1, 45, RC1, RD1)
            '    XYRotation(C2, D2, 45, RC2, RD2)
            '    XYRotation(x1, y1, 45, Rx1, Ry1)
            '    x = (Rx1 - RA1) / (RA2 - RA1) * (RC2 - RC1) + RC1
            '    y = (Ry1 - RB1) / (RB2 - RB1) * (RD2 - RD1) + RD1
            '    XYRotation(x, y, -45, x, y)
            'Else
            '    x = (x1 - A1) / (A2 - A1) * (C2 - C1) + C1
            '    y = (y1 - B1) / (B2 - B1) * (D2 - D1) + D1
            'End If

            'x = Math.Round(x, 6)
            'y = Math.Round(y, 6)

            '最初の座標を原点に移動する。
            Dim X1O1 As Double = 0.0
            Dim Y1O1 As Double = 0.0
            Dim X2O1 As Double = A2 - A1
            Dim Y2O1 As Double = B2 - B1

            Dim X1O2 As Double = 0.0
            Dim Y1O2 As Double = 0.0
            Dim X2O2 As Double = C2 - C1
            Dim Y2O2 As Double = D2 - D1

            ' 第１座標のずれ量
            'Dim diffx As Double = C1 - A1
            'Dim diffy As Double = D1 - B1

            ' 座標のずれ量センター
            Dim diffx2 As Double = C1 - A1
            Dim diffy2 As Double = D1 - B1
            Dim diffx As Double = (C1 + C2) / 2 - (A1 + A2) / 2
            Dim diffy As Double = (D1 + D2) / 2 - (B1 + B2) / 2

            '距離の比率を求める
            Dim distance1 As Double, distance2 As Double, Rate As Double, distance3 As Double
            distance1 = GetDistance(X1O1, Y1O1, X2O1, Y2O1)
            distance2 = GetDistance(X1O2, Y1O2, X2O2, Y2O2)
            distance3 = GetDistance(0, 0, x1, y1)
            Rate = distance2 / distance1

            '角度を求める
            Dim degree1 As Double, degree2 As Double, diffdegree As Double
            degree1 = GetDegree(X1O1, Y1O1, X2O1, Y2O1)
            degree2 = GetDegree(X1O2, Y1O2, X2O2, Y2O2)
            diffdegree = degree2 - degree1

            ' 距離と角度から座標を求める。
            Dim dX1 As Double, dY1 As Double, dX2 As Double, dY2 As Double
            GetXYfromDegree(diffdegree, distance3, dX2, dY2)

            XYRotation(x1, y1, diffdegree, dX1, dY1)    ' 参考コード

            If bStageMode Then
                'x = dX1 + diffx
                'y = dY1 + diffy
                x = dX1 + diffx2
                y = dY1 + diffy2
            Else
                x = dX1
                y = dY1
            End If

            Call DebugLogOut(String.Format("HoseiCalc (A1,B1)=({0},{1})(A2,B2)=({2},{3})(C1,D1)=({4},{5})(C2,D2)=({6},{7})(x1,x1)=({8},{9})(x,y)=({10},{11})", A1, B1, A2, B2, C1, D1, C2, D2, x1, y1, x, y))

        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub
#End Region

#Region "２点の座標データから求める座標を求める"
    Sub HoseiCoordinate(ByVal x1 As Double, ByVal y1 As Double, ByRef x As Double, ByRef y As Double, Optional ByVal bStageMode As Boolean = True)
        Try
            HoseiCalc(bStageMode, stThta.fpp32_x, stThta.fpp32_y, stThta.fpp33_x, stThta.fpp33_y, stThta.fpp32_x + stResult.fCor1x, stThta.fpp32_y + stResult.fCor1y, stThta.fpp33_x + stResult.fCor2x, stThta.fpp33_y + stResult.fCor2y, x1, y1, x, y)
            Call DebugLogOut(String.Format("HoseiCoordinate (X,Y)=({0},{1})", x, y))
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub

#End Region

#Region "２点の座標データから求める座標のずれ量を求める"
    Sub HoseiCoordinateDelta(ByVal x1 As Double, ByVal y1 As Double, ByRef x As Double, ByRef y As Double)
        Try
            Dim x2 As Double, y2 As Double
            HoseiCoordinate(x1, y1, x2, y2)
            x = x2 - x1
            y = y2 - y1
            Call DebugLogOut(String.Format("HoseiCoordinateDelta (⊿X,⊿Y)=({0},{1})", x, y))
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub

#End Region

#Region "ステージステップオフセットを求める"
    Sub GetTstepOffset(ByVal BlockX As Integer, ByVal BlockY As Integer, ByRef dlbOffX As Double, ByRef dlbOffY As Double)
        Try
            Dim dXpos As Double, dYpos As Double

            dXpos = stPLT.zsx * (BlockX - 1)
            dYpos = stPLT.zsy * (BlockY - 1)
            HoseiCoordinateDelta(dXpos, dYpos, dlbOffX, dlbOffY)
            Call DebugLogOut(String.Format("GetTstepOffset BLOCK({0},{1})(X,Y)=({2},{3})(⊿X,⊿Y)=({4},{5})", BlockX, BlockY, dXpos, dYpos, dlbOffX, dlbOffY))

        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try

    End Sub

#End Region

#Region "メッセージ表示（コンソールキー対応）"
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
    Public Function FrmMessageDisp(ByVal ObjSys As Object, ByVal gMode As Integer, ByVal Md As Integer, ByVal BtnDsp As Boolean, _
                                       ByVal Msg1 As String, ByVal Msg2 As String, ByVal Msg3 As String, _
                                       ByVal Col1 As Object, ByVal Col2 As Object, ByVal Col3 As Object) As Integer

        Dim r As Integer
        Dim objForm As Object = Nothing
        Dim ColAry(3) As Object
        Dim MsgAry(3) As String

        Try
            ' パラメータ設定
            MsgAry(0) = Msg1
            MsgAry(1) = Msg2
            MsgAry(2) = Msg3
            ColAry(0) = Col1
            ColAry(1) = Col2
            ColAry(2) = Col3

            ' frmMessage画面表示(指定のメッセージを表示する)
            objForm = New frmMessage()
            Call objForm.ShowDialog(Nothing, gMode, ObjSys, MsgAry, ColAry, Md, BtnDsp)
            r = objForm.sGetReturn()                                    ' Return値取得

            ' オブジェクト開放
            If (objForm Is Nothing = False) Then
                Call objForm.Close()                                    ' オブジェクト開放
                Call objForm.Dispose()                                  ' リソース開放
            End If

            Return (r)                                                  ' Return(エラー時のメッセージは表示済) 

            ' トラップエラー発生時 
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region


#Region "カメラ画像表示PictureBoxクリック位置をJOG経由で画像センターに移動する"
    ''' <summary>カメラ画像表示PictureBoxクリック位置をJOG経由で画像センターに移動する</summary>
    ''' <param name="distanceX"></param>
    ''' <param name="distanceY"></param>
    ''' <param name="stJOG">'V6.0.0.0-23</param>
    ''' <remarks>'V6.0.0.0⑧</remarks>
    Public Sub MoveToCenter(ByVal distanceX As Decimal, ByVal distanceY As Decimal, ByRef stJOG As JOG_PARAM)
        stJOG.KeyDown = Keys.Execute                                    'V6.0.0.0-23
        InpKey = (InpKey Or CtrlJog.MouseClickLocation.GetInpKey(distanceX, distanceY))
    End Sub
#End Region


End Module

'=============================== END OF FILE ===============================
