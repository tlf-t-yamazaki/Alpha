
Imports TrimClassLibrary                'V6.0.0.0①  'V2.2.0.0①

'V2.2.0.0① 'Public Class MainModules
Public Class MainModules
    Implements IMainModules           'V6.0.0.0①    'V2.2.0.0①

#Region "アプリケーション種別を返す(OCX用)"
    '''=========================================================================
    ''' <summary>アプリケーション種別を返す</summary>
    ''' <param name="AppKind">9=ユーザプロ</param>
    '''=========================================================================
    Public Sub GetAppKind(ByRef AppKind As Short) Implements IMainModules.GetAppKind    'V6.0.0.0①  'V2.2.0.0①
        AppKind = KND_USER
    End Sub
#End Region

#Region "抵抗(チップ)並び方向を返す(OCX用)"
    '''=========================================================================
    ''' <summary>抵抗(チップ)並び方向を返す</summary>
    ''' <param name="ResistDir">0=X方向, 1=Y方向</param>
    '''=========================================================================
    Public Sub GetResistDir(ByRef ResistDir As Short) Implements IMainModules.GetResistDir  'V6.0.0.0① 'V2.2.0.0①
        ResistDir = 0
    End Sub
#End Region

#Region "プレート内ブロックのX方向、Y方向の開始位置算出(OCX用)"
    '''=========================================================================
    ''' <summary>プレート内ブロックのX方向、Y方向の開始位置算出</summary>
    ''' <returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Public Function Call_CalcBlockXYStartPos() As Integer Implements IMainModules.Call_CalcBlockXYStartPos  'V6.0.0.0①  'V2.2.0.0①

        Dim r As Integer = cFRS_NORMAL

        'r = CalcBlockXYStartPos()
        Return (r)

    End Function
#End Region

#Region "指定ブロックXYからステージ位置XYを取得しテーブル移動する(OCX用)"
    '''=========================================================================
    ''' <summary>指定ブロックXYからステージ位置XYを取得しテーブル移動する</summary>
    ''' <param name="xBlockNo">(INP)ブロック番号X</param>
    ''' <param name="yBlockNo">(INP)ブロック番号Y</param>
    ''' <param name="OffSetX"> (INP)オフセットX</param>
    ''' <param name="OffSetY"> (INP)オフセットY</param>
    ''' <param name="stgx">    (OUT)ステージ位置X</param>
    ''' <param name="stgy">    (OUT)ステージ位置Y</param>
    ''' <returns>0=正常, 0以外=エラー</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function Call_GetTargetStagePosByXY(ByVal xBlockNo As Integer, ByVal yBlockNo As Integer,
                                               ByVal OffSetX As Double, ByVal OffSetY As Double,
                                               ByRef stgx As Double, ByRef stgy As Double) As Integer _
                                               Implements IMainModules.Call_GetTargetStagePosByXY       'V6.0.0.0①

        Dim r As Integer

        If giAppMode = APP_MODE_PROBE And (stPLT.TeachBlockX > 1 Or stPLT.TeachBlockY > 1) Then ' ###1040① Move_Trimposition()で移動しているのでここでは移動しない。
            Return (cFRS_NORMAL)                                                                ' ###1040①
        End If                                                                                  ' ###1040①

        ' XYテーブル指定ブロック移動
        r = TSTEP(xBlockNo, yBlockNo, OffSetX, OffSetY)
        r = ObjSys.EX_ZGETSRVSIGNAL(gSysPrm, r, 0)

        Return (r)

    End Function
#End Region

#Region "(Teaching向け)メイン画面上のクロスラインの表示位置を変更する"
    '''=========================================================================
    '''<summary>メイン画面上のクロスラインの表示位置を変更する</summary>
    '''=========================================================================
    Public Sub SetCrossLinePos(ByVal xPos As Integer, ByVal yPos As Integer)

        ' クロスライン位置を設定する 
        'Form1.Picture1.Top = xPos + Form1.VideoLibrary1.Location.Y
        'Form1.Picture2.Left = yPos + Form1.VideoLibrary1.Location.X
        ' クロスライン位置を設定する 
        ObjVdo.SetCorrCrossCenter(yPos, xPos)

        ' 画面の再描画
        'Form1.Refresh()
    End Sub
#End Region

#Region "(Teaching向け)マーキングエリア表示"
    '''=========================================================================
    '''<summary>メイン画面上のマーキングエリアの四角を表示/非表示する</summary>
    '''=========================================================================
    Public Sub DisplayMarkingArea(ByVal bDisp As Boolean, ByVal xPos As Integer, ByVal yPos As Integer,
                                        ByVal width As Integer, ByVal height As Integer)

        ObjVdo.SetMarkingArea(bDisp, xPos, yPos, width, height)

    End Sub
#End Region

#Region "(Teaching-Jog向け)クロスライン位置移動表示"
    Public Sub DispCrossLine(ByVal xPos As Double, ByVal yPos As Double) Implements IMainModules.DispCrossLine
        ObjCrossLine.CrossLineDispXY(xPos, yPos)
        ''クロスライン補正処理を呼び出す
        'gstCLC.x = xPos                    ' BP位置X(mm)
        'gstCLC.y = yPos                    ' BP位置Y(mm)
        'Call CrossLineCorrect(gstCLC)       ' 補正クロスライン表示
    End Sub
#End Region

#Region "補正クロスライン位置取得"
    Public Sub GetCorrCrossLinePixel(ByVal bpx As Double, ByVal bpy As Double, ByRef xPos As Integer, ByRef yPos As Integer) Implements IMainModules.GetCorrCrossLinePixel

        ObjCrossLine.GetCorrCrossPixel(bpx, bpy, xPos, yPos)

    End Sub

#End Region
#Region "ビデオのスタートストップ処理"
    '===========================================================================
    ' ビデオスタート停止処理
    '===========================================================================
    ''' <summary>
    ''' ビデオのスタート処理
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub VideoStart() Implements IMainModules.VideoStart  'V6.0.0.0① 'V2.2.0.0①
        Try
            ''Call ObjVdo.VideoStart()
            Call ObjVdo.VideoStart()        ''V2.2.0.0⑭
        Catch ex As Exception
            MsgBox("MainModules.VideoStart() TRAP ERROR = " + ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' ビデオのストップ処理
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub VideoStop() Implements IMainModules.VideoStop  'V6.0.0.0① 'V2.2.0.0①
        Try
            ''Call ObjVdo.VideoStop()
            Call ObjVdo.VideoStop() 'V2.2.0.0⑭
        Catch ex As Exception
            MsgBox("MainModules.VideoStop() TRAP ERROR = " + ex.Message)
        End Try
    End Sub
#End Region

#Region "I/F合わせるためのダミー関数"   'V2.2.0.0①
    Public Sub Call_SetAlmStartTime(strDAT As String, ErrCode As Short) Implements IMainModules.Call_SetAlmStartTime
        Throw New NotImplementedException()
    End Sub

    Public Sub Call_GetVacumeStatus(ByRef Sts As Integer) Implements IMainModules.Call_GetVacumeStatus
        '@@@888
        ' 吸着センサの状態を取得する関数呼び出し 
        Sts = 1
        '@@@888        Throw New NotImplementedException()
    End Sub

    Public Sub Call_SetAlmEndTime() Implements IMainModules.Call_SetAlmEndTime
        Throw New NotImplementedException()
    End Sub

    Public Sub CrossLineDispOff() Implements IMainModules.CrossLineDispOff
        ObjVdo.SetCorrCrossVisible(False)
    End Sub

    Public Sub SetVideoTrackBar(visible As Boolean, enabled As Boolean) Implements IMainModules.SetVideoTrackBar
        ObjVdo.SetTrackBar(visible, enabled)
    End Sub

    Public Sub SetActiveJogMethod(keyDown As Action(Of KeyEventArgs),
                                  keyUp As Action(Of KeyEventArgs),
                                  moveToCenter As Action(Of Decimal, Decimal)) Implements IMainModules.SetActiveJogMethod
        ObjMain.SetActiveJogMethod(keyDown, keyUp, moveToCenter)

    End Sub

    Public Function Call_Loader_AlarmCheck_ManualMode() As Integer Implements IMainModules.Call_Loader_AlarmCheck_ManualMode
        Throw New NotImplementedException()
    End Function

    Public Sub sub_MGStopJog() Implements IMainModules.sub_MGStopJog
        ''V2.2.0.0⑤        Throw New NotImplementedException()
    End Sub

    Public Sub SetAPBPrm_Step(StepCount As Short, ProbePitch As Double, StepCount2 As Short, ProbePitch2 As Double) Implements IMainModules.SetAPBPrm_Step
        Throw New NotImplementedException()
    End Sub

    Public Function Sub_CallFrmMatrix() As Integer Implements IMainModules.Sub_CallFrmMatrix
        Throw New NotImplementedException()
    End Function

    Public Function Sub_ProbeCheck(Plt As Integer, BlkX As Integer, BlkY As Integer, Limit As Double, ByRef strLOG As String, ByRef strDSP As String) As Integer Implements IMainModules.Sub_ProbeCheck
        Throw New NotImplementedException()
    End Function

    Public Function CalcStartPointPlate(ByRef OrgStartPoint(,) As Double, ByRef StartPoint(,) As Double, ByVal PlateX As Integer, ByVal PlateY As Integer) As Boolean Implements IMainModules.CalcStartPointPlate
        Throw New NotImplementedException()
    End Function

    Public Function Call_StagePosByPltBlkNo(ByVal xPlateNo As Integer, ByVal yPlateNo As Integer, ByVal xBlockNo As Integer, ByVal yBlockNo As Integer,
                                            ByVal stgOffSetX As Double, ByVal stgOffSetY As Double, ByRef stgx As Double, ByRef stgy As Double) As Integer _
     Implements IMainModules.Call_StagePosByPltBlkNo
        Throw New NotImplementedException()
    End Function

    Public Function GetPlateBlockNumber(ByVal plateX As Integer, ByVal plateY As Integer, ByVal blockX As Integer, ByVal blockY As Integer, ByRef plateNumber As Integer, ByRef blockNumber As Integer) As Boolean _
    Implements IMainModules.GetPlateBlockNumber
        Throw New NotImplementedException()
    End Function

    Public Function GetPlateBPOffset(ByVal plateX As Integer, ByVal plateY As Integer, ByRef OffsetX As Double, ByRef OffsetY As Double) As Boolean _
    Implements IMainModules.GetPlateBPOffset
        Throw New NotImplementedException()
    End Function

    Public Function GetPlateStepOffset(ByVal plateX As Integer, ByVal plateY As Integer, ByRef OffsetX As Double, ByRef OffsetY As Double) As Boolean _
    Implements IMainModules.GetPlateStepOffset
        Throw New NotImplementedException()
    End Function

    Public Function GetPlateTableOffset(ByVal plateX As Integer, ByVal plateY As Integer, ByRef OffsetX As Double, ByRef OffsetY As Double) As Boolean _
    Implements IMainModules.GetPlateTableOffset
        Throw New NotImplementedException()
    End Function

    Public Function GetPlateTXOffset(ByVal plateX As Integer, ByVal plateY As Integer, ByRef OffsetX As Double, ByRef OffsetY As Double) As Boolean _
     Implements IMainModules.GetPlateTXOffset
        Throw New NotImplementedException()
    End Function

    Public Function ReverseCalcStartPointPlate(ByRef OrgStartPoint(,) As Double, ByRef StartPoint(,) As Double, ByVal PlateX As Integer, ByVal PlateY As Integer) As Boolean _
        Implements IMainModules.ReverseCalcStartPointPlate
        Throw New NotImplementedException()
    End Function

    Public Function SetPlateBPOffset(ByVal plateX As Integer, ByVal plateY As Integer, ByVal OffsetX As Double, ByVal OffsetY As Double) As Boolean _
        Implements IMainModules.SetPlateBPOffset
        Throw New NotImplementedException()
    End Function

    Public Function SetPlateOffsetData(ByVal plateX As Integer, ByVal plateY As Integer) As Boolean _
     Implements IMainModules.SetPlateOffsetData
        Throw New NotImplementedException()
    End Function

    Public Function SetStartPointPlate(ByRef OrgStartPoint(,) As Double, ByRef StartPoint(,) As Double, ByVal PlateX As Integer, ByVal PlateY As Integer) As Boolean _
    Implements IMainModules.SetStartPointPlate
        Throw New NotImplementedException()
    End Function

    Public Function SetTrimParameterPlate(ByVal ChipSize(,) As Double, ByVal PlateNo As Integer) As Boolean _
             Implements IMainModules.SetTrimParameterPlate
        Throw New NotImplementedException()
    End Function

    Public Function SetTrimParamToGlobalAreaPlate(ByVal ChipSize(,) As Double, ByVal PlateNo As Integer) As Boolean _
     Implements IMainModules.SetTrimParamToGlobalAreaPlate
        Throw New NotImplementedException()
    End Function

    Public Function MainGpibUnitOff() As Integer _
        Implements IMainModules.MainGpibUnitOff
        'Throw New NotImplementedException()                            
    End Function
    Public Function Call_GetCHTheta(Channel As String, ByRef angle As Double) As Integer _
        Implements IMainModules.Call_GetCHTheta
        Throw New NotImplementedException()
    End Function

#Region "θ補正"
    '''=========================================================================
    '''<summary>θ補正</summary>
    '''=========================================================================
    Public Function Call_DoCorrectPos(dispCurPltNoX As Integer, dispCurPltNoY As Integer,
                                 Optional ByVal posz As Double = -9999,
                                 Optional ByVal highsensortpos As Double = -9999) As Integer Implements IMainModules.Call_DoCorrectPos

    End Function
#End Region

    Public Sub CrossLineDispOn() Implements IMainModules.CrossLineDispOn      'V6.0.0.0①

    End Sub


    '''=========================================================================
    ''' <summary>クロスラインOFFSET設定</summary>
    ''' <param name="xOffset">(INP)OFFSET X位置(pixel)</param>
    ''' <param name="yOffset">(INP)OFFSET Y位置(pixel)</param>
    '''=========================================================================
    Public Sub DispCrossOffset(ByVal xOffset As Integer, ByVal yOffset As Integer) _
        Implements IMainModules.DispCrossOffset                           'V6.0.0.0①

    End Sub

    ''' <summary>
    ''' ローダ原点復帰の呼び戻し用 
    ''' </summary>
    ''' <param name="mode"></param>
    Public Function Call_Sub_Loader_OrgBack(ByVal mode As Integer) As Integer _
        Implements IMainModules.Call_Sub_Loader_OrgBack                           'V6.0.0.0①

        Dim r As Integer

        Try

            r = ObjLoader.Sub_Loader_OrgBack(mode)

            Call_Sub_Loader_OrgBack = r
        Catch ex As Exception

        End Try

    End Function


    ''' <summary>
    ''' ローダ原点復帰の完了を待つ       'V2.2.1.1⑦
    ''' </summary>
    ''' <returns></returns>
    Public Function Call_Sub_WaitLoaderOrigin() As Integer _
                Implements IMainModules.Call_Sub_WaitLoaderOrigin
        Dim r As Integer = cFRS_NORMAL

        Try

            r = ObjLoader.WaitLoaderOrigin()

            Call_Sub_WaitLoaderOrigin = r
            Return r

        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' ロット切り替え信号の設定を行う        'V2.2.1.1⑦
    ''' </summary>
    ''' <param name="count">ロット切り替え回数</param>
    ''' <returns></returns>
    Public Function Call_Sub_LotChangeFlgSet(ByVal count As Integer) As Integer _
                Implements IMainModules.Call_Sub_LotChangeFlgSet
        Dim r As Integer = cFRS_NORMAL

        Try

            r = ObjLoader.SetLotChangeFlg(count)

            Return r

        Catch ex As Exception

        End Try

    End Function


    ''' <summary>
    ''' ロット切り替え信号の設定を行う        'V2.2.1.1⑦
    ''' </summary>
    ''' <param name="count">ロット切り替え回数</param>
    ''' <returns></returns>
    Public Function Call_Sub_LotChangeFlgGet() As Integer _
                Implements IMainModules.Call_Sub_LotChangeFlgGet

        Dim r As Integer = cFRS_NORMAL

        Try

            r = ObjLoader.GetLotChangeFlg()

            Return r

        Catch ex As Exception

        End Try

    End Function


#End Region

End Class

