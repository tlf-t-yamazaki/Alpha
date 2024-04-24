Public Class frmMessage

#Region "【変数定義】"
    ''===========================================================================
    ''   変数定義
    ''===========================================================================
    Public giTrimErr As Short = 0                           ' ﾄﾘﾏｰ ｴﾗｰ ﾌﾗｸﾞ ※ｴﾗｰ時はｸﾗﾝﾌﾟｸﾗﾝﾌﾟOFF時ﾄﾘﾏ動作中OFFをﾛｰﾀﾞｰに送信しない
    '                                                       ' B0 : 吸着ｴﾗｰ(EXIT)
    '                                                       ' B1 : その他ｴﾗｰ
    '                                                       ' B2 : 集塵機ｱﾗｰﾑ検出
    '                                                       ' B3 : 軸ﾘﾐｯﾄ､軸ｴﾗｰ､軸ﾀｲﾑｱｳﾄ
    '                                                       ' B4 : 非常停止
    '                                                       ' B5 : ｴｱｰ圧ｴﾗｰ
    ''-------------------------------------------------------------------------------
    ''   ランプ ON/OFF制御用ランプ番号(コンソール制御)
    ''-------------------------------------------------------------------------------
    'Public Const LAMP_START As Short = 0                            ' STARTランプ
    'Public Const LAMP_RESET As Short = 1                            ' RESETランプ
    'Public Const LAMP_PRB As Short = 2                              ' Zランプ
    'Public Const LAMP_HALT As Short = 5                             ' HALTランプ
    'Public Const LATCH_CLR As Short = 11                            ' B11 : カバー開ラッチクリア

    ''-------------------------------------------------------------------------------
    ''   インターロック状態
    ''-------------------------------------------------------------------------------
    ''----- インターロック状態 -----
    'Public Const INTERLOCK_STS_DISABLE_NO As Integer = 0            ' インターロック状態（解除なし）
    'Public Const INTERLOCK_STS_DISABLE_PART As Integer = 1          ' インターロック一部解除（ステージ動作可能）
    'Public Const INTERLOCK_STS_DISABLE_FULL As Integer = 2          ' インターロック全解除

    ''----- アクチュエータ入力ビット -----
    'Public Const BIT_SLIDE_COVER_OPEN As UShort = &H1               ' スライドカバー開(=1)
    'Public Const BIT_SLIDE_COVER_CLOSE As UShort = &H2              ' スライドカバー閉(=1)
    'Public Const BIT_SLIDE_COVER_MOVING As UShort = &H4             ' スライドカバー動作中(=1)
    'Public Const BIT_SOURCE_AIR_CHECK As UShort = &H8               ' 供給元エアー：0/1=異常/正常
    'Public Const BIT_MAIN_COVER_OPENCLOSE As UShort = &H10          ' 固定カバー：0/1=開/閉
    'Public Const BIT_COVER_OPEN_RATCH As UShort = &H20              ' カバー開ラッチ
    'Public Const BIT_INTERLOCK_NO1_RELEASE As UShort = &H100        ' インターロック解除1：0/1=無効/有効
    'Public Const BIT_INTERLOCK_NO2_RELEASE As UShort = &H200        ' インターロック解除2：0/1=無効/有効
    'Public Const BIT_EMERGENCY_STATUS_ONOFF As UShort = &H400       ' 非常停止状態：0/1=異常/正常 (※非常停止はH/Wが落ちるので返って来ない)

    ''-------------------------------------------------------------------------------
    ''   その他
    ''-------------------------------------------------------------------------------
    '----- Formの幅/高さ -----
    Private Const WIDTH_NOMAL As Integer = 570                      ' Formの幅
    Private Const WEIGHT_NOMAL As Integer = 203                     ' Formの高さ(通常モード)
    ''Private Const WEIGHT_LDALM As Integer = 460                    ' Formの高さ(ローダアラームモード)
    'Private Const WEIGHT_LDALM As Integer = 203 + 129 + 460         ' Formの高さ(ローダアラームモード) ###161
    'Private Const WEIGHT_LDALM2 As Integer = 203 + 129 + 2          ' Formの高さ(解除ボタン表示モード) ###161

    Private stSzNML As System.Drawing.Size = New System.Drawing.Size(WIDTH_NOMAL, WEIGHT_NOMAL)
    'Private stSzLDE As System.Drawing.Size = New System.Drawing.Size(WIDTH_NOMAL, WEIGHT_LDALM)
    'Private stSzLDE2 As System.Drawing.Size = New System.Drawing.Size(WIDTH_NOMAL, WEIGHT_LDALM2) ' ###161

    ''----- Cancelボタン表示位置 ----- 
    Private LocBtnLeft As System.Drawing.Point = New System.Drawing.Point(219 - 130, 163)
    Private LocBtnRight As System.Drawing.Point = New System.Drawing.Point(219 + 130, 163)
    Private LocBtnCenter As System.Drawing.Point = New System.Drawing.Point(219, 162)


    '----- 変数定義 -----
    Private mExitFlag As Integer                                    ' 終了フラグ
    Private gMode As Integer                                        ' 処理モード
    Private ObjSys As Object                                        ' OcxSystemオブジェクト

    ''----- ローダアラーム情報 -----
    'Private AlarmCount As Integer
    'Private strLoaderAlarm(LALARM_COUNT) As String                  ' アラーム文字列
    'Private strLoaderAlarmInfo(LALARM_COUNT) As String              ' アラーム情報1
    'Private strLoaderAlarmExec(LALARM_COUNT) As String              ' アラーム情報(対策)

    '----- 指定メッセージ表示用 -----
    Private Const MSGARY_NO As Integer = 3                          ' 表示メッセージの最大数
    Private DspWaitKey As Integer                                   ' WaitKey 
    Private DspBtnDsp As Boolean                                    ' ボタン表示する/しない
    Private strMsgAry(MSGARY_NO) As String                          ' 表示メッセージ１－３
    Private ColColAry(MSGARY_NO) As Object                          ' メッセージ色１－３

    'V2.2.0.0⑱↓
    Private buzzerStop As Boolean                                   '　ブザーを止める用  'V2.2.0.0⑱
    Private TimerLotEnd As System.Threading.Timer
    'V2.2.0.0⑱↑


#End Region

#Region "終了結果を返す"
    '''=========================================================================
    ''' <summary>終了結果を返す</summary>
    ''' <returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Public ReadOnly Property sGetReturn() As Integer
        Get
            Return (mExitFlag)
        End Get
    End Property
#End Region

#Region "ShowDialogメソッドに独自の引数を追加する(指定メッセージ表示用)"
    '''=========================================================================
    ''' <summary>ShowDialogメソッドに独自の引数を追加する(指定メッセージ表示用) ###089</summary>
    ''' <param name="Owner">    (INP)未使用</param>
    ''' <param name="iGmode">   (INP)処理モード</param>
    ''' <param name="ObjSystem">(INP)OcxSystemオブジェクト</param>
    ''' <param name="MsgAry">   (INP)表示メッセージ１－３</param>
    ''' <param name="ColAry">   (INP)メッセージ色１－３</param>
    ''' <param name="Md">       (INP)cFRS_ERR_START                = STARTキー押下待ち
    '''                              cFRS_ERR_RST                  = RESETキー押下待ち
    '''                              cFRS_ERR_START + cFRS_ERR_RST = START/RESETキー押下待ち</param>
    ''' <param name="BtnDsp">   (INP)ボタン表示する/しない</param>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Overloads Sub ShowDialog(ByVal Owner As IWin32Window, ByVal iGmode As Integer, ByVal ObjSystem As Object, _
                                    ByVal MsgAry() As String, ByVal ColAry() As Object, ByVal Md As Integer, ByVal BtnDsp As Boolean)

        Dim Idx As Integer

        Try
            ' 初期処理
            Console.WriteLine("パラメータ = " + gMode.ToString)
            BtnCancel.Location = LocBtnCenter
            BtnOK.Location = LocBtnCenter
            mExitFlag = -1                                              ' 終了フラグ = 初期化
            gMode = iGmode                                              ' 処理モード
            ObjSys = ObjSystem                                          ' OcxSystemオブジェクト
            LblCaption.Text = ""
            Label1.Text = ""
            Label2.Text = ""

            ' パラメータを取得する
            For Idx = 0 To (MSGARY_NO - 1)
                strMsgAry(Idx) = ""
                ColColAry(Idx) = System.Drawing.SystemColors.ControlText
            Next Idx
            DspWaitKey = Md                                             ' WaitKey
            DspBtnDsp = BtnDsp                                          ' ボタン表示する/しない

            ' 表示メッセージ１－３
            For Idx = 0 To (MsgAry.Length - 1)
                If (Idx > (MSGARY_NO - 1)) Then Exit For
                strMsgAry(Idx) = MsgAry(Idx)
            Next Idx

            ' メッセージ色１－３
            For Idx = 0 To (ColAry.Length - 1)
                If (Idx > (MSGARY_NO - 1)) Then Exit For
                ColColAry(Idx) = ColAry(Idx)
            Next Idx

            ' 画面表示
            Me.Size = stSzNML                                           ' Formの幅/高さを通常モード用にする
            Me.ShowDialog()                                             ' 画面表示
            Me.BringToFront()                                           ' 最前面に表示 
            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            MsgBox("frmMessage.ShowDialog() TRAP ERROR = " + ex.Message)
        End Try
    End Sub
#End Region

#Region "フォームが表示された時の処理"
    '''=========================================================================
    ''' <summary>フォームが表示された時の処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub frmMessage_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

        Dim r As Integer = cFRS_NORMAL
        Dim strMSG As String

        Try
            ' 画面処理メイン
            r = frmMessage_Main(gMode)
            mExitFlag = r                                           ' mExitFlagに戻り値を設定する 

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "frmMessage.frmMessage_Shown() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
        Me.Close()                                                  ' フォームを閉じる
    End Sub
#End Region

#Region "Cancel(or OK)ボタン押下時処理"
    '''=========================================================================
    ''' <summary>Cancel(or OK)ボタン押下時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click

        Try
            If (BtnCancel.Text = "Cancel") Then
                mExitFlag = cFRS_ERR_RST
            Else
                mExitFlag = cFRS_ERR_START
            End If

            ' トラップエラー発生時 
        Catch ex As Exception
            MsgBox("frmMessage.BtnCancel_Click() TRAP ERROR = " + ex.Message)
        End Try

    End Sub
#End Region

#Region "OKボタン押下時処理"
    '''=========================================================================
    ''' <summary>OKボタン押下時処理 ###073</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Try
            mExitFlag = cFRS_ERR_START

            ' トラップエラー発生時 
        Catch ex As Exception
            MsgBox("frmMessage.BtnOK_Click() TRAP ERROR = " + ex.Message)
        End Try
    End Sub
#End Region

#Region "画面処理メイン"
    '''=========================================================================
    ''' <summary>画面処理メイン</summary>
    ''' <param name="gMode">(INP)処理モード</param>
    ''' <returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Private Function frmMessage_Main(ByVal gMode As Integer) As Integer

        Dim r As Integer = cFRS_NORMAL
        Dim strMSG As String

        Try
            '-------------------------------------------------------------------
            '   処理モードに対応する処理を行う
            '-------------------------------------------------------------------
            Select Case gMode
                '   '-----------------------------------------------------------
                '   '   原点復帰処理
                '   '-----------------------------------------------------------
                Case cGMODE_ORG                                                 ' 原点復帰処理
                    '-----------------------------------------------------------
                    '   非常停止メッセージ表示
                    '-----------------------------------------------------------
                Case cGMODE_EMG
STP_EMERGENCY:
                    giTrimErr = giTrimErr Or &H10                   ' ﾄﾘﾏｰ ｴﾗｰ ﾌﾗｸﾞ(非常停止)
                    r = ObjSys.SetSignalTower(0, &HFFFF)            ' ｼｸﾞﾅﾙﾀﾜｰ制御(On=0, Off=全ﾋﾞｯﾄ)
                    Call EXTOUT1(0, &HFFFF)                         ' EXTBIT (On=0, Off=全ビット)
                    Call EXTOUT2(0, &HFFFF)                         ' EXTBIT2(On=0, Off=全ビット)
                    r = Sub_DispEmergencyMsg()                      ' 非常停止メッセージ表示

                    '-----------------------------------------------------------
                    '   集塵機異常メッセージ表示
                    '-----------------------------------------------------------
                Case cGMODE_ERR_DUST
STP_ARMDUST:
                    Call LAMP_CTRL(LAMP_START, False)               ' STARTﾗﾝﾌﾟ消灯
                    ' メッセージ表示
                    ' "集塵機異常が発生しました", "RESETキーを押すとプログラムを終了します", ""
                    Call Sub_SetMessage("Cancelボタン押下でプログラムを終了します", "RESETキーを押すとプログラムを終了します", "", System.Drawing.Color.Red, System.Drawing.Color.Red, System.Drawing.SystemColors.ControlText)
                    Me.Refresh()

                    ' メッセージ表示してRESETキー押下待ち
                    r = Sub_WaitStartRestKey(cFRS_ERR_RST)
                    If (r = cFRS_ERR_EMG) Then GoTo STP_EMERGENCY
                    r = cFRS_ERR_DUST                                           ' Return値 = 集塵機異常検出

                    '-----------------------------------------------------------
                    '   エアー圧低下検出メッセージ表示
                    '-----------------------------------------------------------
                Case cGMODE_ERR_AIR
STP_AIRVALVE:
                    Call LAMP_CTRL(LAMP_START, False)                   ' STARTﾗﾝﾌﾟ消灯
                    ' メッセージ表示
                    ' "エアー圧低下検出", "RESETキーを押すとプログラムを終了します", ""
                    Call Sub_SetMessage("エアー圧低下検出", "RESETキーを押すとプログラムを終了します", "", System.Drawing.Color.Red, System.Drawing.Color.Red, System.Drawing.SystemColors.ControlText)
                    Me.Refresh()

                    ' RESETキー押下待ち
                    r = Sub_WaitStartRestKey(cFRS_ERR_RST)
                    If (r = cFRS_ERR_EMG) Then GoTo STP_EMERGENCY
                    r = cFRS_ERR_AIR                                            ' Return値 = エアー圧エラー検出

                    '-----------------------------------------------------------
                    '   自動運転開始(STARTｷｰ押下待ち)
                    '-----------------------------------------------------------
                Case cGMODE_LDR_START
                    ' メッセージ表示
                    ' "STARTキーを押すと自動運転を開始します", "", ""
                    Call Sub_SetMessage("STARTキーを押すと自動運転を開始します", "", "", System.Drawing.SystemColors.ControlText, System.Drawing.SystemColors.ControlText, System.Drawing.SystemColors.ControlText)
                    Me.Refresh()

                    ' STARTキー押下待ち
                    r = Sub_WaitStartRestKey(cFRS_ERR_START + cFRS_ERR_RST)
                    If (r = cFRS_ERR_EMG) Then GoTo STP_EMERGENCY
                    '                                                           ' Return値 = cFRS_ERR_START(STARTキー押下)/cFRS_ERR_RST(RESETキー押下)

                    '-----------------------------------------------------------
                    '   自動運転終了(STARTｷｰ押下待ち)
                    '-----------------------------------------------------------
                Case cGMODE_LDR_END
                    ' メッセージ表示
                    BtnCancel.Visible = True                                    ' Cancel(OK)ボタン表示
                    BtnCancel.Text = "OK"
                    ' "自動運転終了", "STARTキーを押すか、OKボタンを押して下さい。", ""
                    Call Sub_SetMessage("自動運転終了", "STARTキーを押すか、OKボタンを押して下さい。", "", System.Drawing.SystemColors.ControlText, System.Drawing.SystemColors.ControlText, System.Drawing.SystemColors.ControlText)
                    Me.Refresh()

                    'V2.2.0.0⑱↓
                    buzzerStop = False
                    TimerLotEnd = New System.Threading.Timer(New System.Threading.TimerCallback(AddressOf TimerLotEnd_Tick), Nothing, 5000, 5000)
                    'V2.2.0.0⑱↑

                    ' STARTキー押下待ち
                    r = Sub_WaitStartRestKey(cFRS_ERR_START)
                    BtnCancel.Visible = False                                   ' Cancel(OK)ボタン非表示
                    BtnCancel.Text = "Cancel"
                    If (r = cFRS_ERR_EMG) Then GoTo STP_EMERGENCY
                    '                                                           ' Return値 = cFRS_ERR_START(STARTキー押下)
                    '-----------------------------------------------------------
                    '   ステージを原点に戻す(残基板取り除くため)
                    '-----------------------------------------------------------
                Case cGMODE_LDR_STAGE_ORG
                    ' "ステージ原点移動中", "", ""
                    Call Sub_SetMessage("ステージ原点移動中", "", "", System.Drawing.SystemColors.ControlText, System.Drawing.SystemColors.ControlText, System.Drawing.SystemColors.ControlText)
                    Me.Refresh()
                    r = Sub_XY_OrgBack()
                    Select Case (r)
                        Case cFRS_ERR_EMG
                            GoTo STP_EMERGENCY                                  ' 非常停止メッセージ表示
                    End Select

                    '-----------------------------------------------------------
                    '   トリミング中のｽﾗｲﾄﾞｶﾊﾞｰ開/筐体ｶﾊﾞｰ開メッセージ表示(STARTｷｰ押下待ち)
                    '-----------------------------------------------------------
                Case cGMODE_SCVR_OPN, cGMODE_CVR_OPN
                    r = Sub_CvrOpen(gMode)                                      ' ｽﾗｲﾄﾞｶﾊﾞｰ開/筐体ｶﾊﾞｰ開ﾒｯｾｰｼﾞ表示
                    If (r = cFRS_ERR_EMG) Then                                  ' 非常停止 ?
                        GoTo STP_EMERGENCY                                      ' 非常停止メッセージ表示へ
                    End If                                                      ' ※ﾄﾘﾐﾝｸﾞ中の筐体ｶﾊﾞｰ開は原点復帰は行わない

                    '-----------------------------------------------------------
                    '   指定メッセージ表示(STARTキー/RESETキー押下待ち)
                    '-----------------------------------------------------------
                Case cGMODE_MSG_DSP
                    ' OK/Cancelボタン表示設定
                    BtnCancel.Visible = False                                   ' Cancelボタン非表示
                    BtnOK.Visible = False                                       ' OKボタン非表示
                    If (DspBtnDsp = True) Then                                  ' ボタン表示する ?
                        If (DspWaitKey = (cFRS_ERR_START + cFRS_ERR_RST)) Then  ' OK/Cancelボタンを表示する ?
                            BtnCancel.Location = LocBtnRight                     ' Cancelボタン表示位置を右にずらす 
                            BtnCancel.Visible = True                            ' Cancelボタン表示
                            BtnOK.Location = LocBtnLeft                     ' Cancelボタン表示位置を右にずらす 
                            BtnOK.Visible = True                                ' OKボタン表示
                        End If
                        If (DspWaitKey = cFRS_ERR_RST) Then                     ' Cancelボタンを表示する ?
                            BtnCancel.Text = "Cancel"
                            BtnCancel.Visible = True                            ' Cancelボタン表示
                        End If
                        If (DspWaitKey = cFRS_ERR_START) Then                   ' OKボタンを表示する ?
                            BtnCancel.Text = "OK"
                            BtnCancel.Visible = True                            ' Cancel(OK)ボタン表示
                        End If
                    End If

                    ' 指定メッセージを表示する
                    Call Sub_SetMessage(strMsgAry(0), strMsgAry(1), strMsgAry(2), ColColAry(0), ColColAry(1), ColColAry(2))
                    Me.Refresh()

                    ' STARTキー/RESETキー押下待ち
                    r = Sub_WaitStartRestKey(DspWaitKey)
                    BtnCancel.Visible = False                                   ' Cancelボタン非表示
                    BtnOK.Visible = False                                       ' OKボタン非表示
                    BtnCancel.Text = "Cancel"
                    If (r = cFRS_ERR_EMG) Then GoTo STP_EMERGENCY
                Case Else
STP_INTRIM:
                    ' INtime側エラー時
                    If (gMode >= ERR_INTIME_BASE) Then                          ' INtime側エラー ?
                        '   ' ソフトリミットエラーの場合
                        If (ObjSys.IsSoftLimitCode(gMode)) Then                 ' ソフトリミットエラー
                            'r = Sub_ErrSoftLimit(gMode, giTrimErr)              ' メッセージ表示&STARTキー押下待ち

                            ' ソフトリミットエラー以外の場合
                        Else
                            ' シグナルタワー3色制御あり(特注) ?
                            If (gSysPrm.stIOC.giSigTwr2Flag = 1) Then
                                ' シグナルタワー３色制御(赤点滅) (EXTOUT(OnBit, OffBit))
                                Call ObjSys.EXTIO_Out_Sub(gSysPrm.stIOC.glSigTwr2_Out_Adr, gSysPrm.stIOC.glSigTwr2_Red_Blnk, _
                                                           gSysPrm.stIOC.glSigTwr2_Red_On Or gSysPrm.stIOC.glSigTwr2_Yellow_On Or gSysPrm.stIOC.glSigTwr2_Yellow_Blnk)
                                ' ブザー制御あり(特注) ?
                                If (gSysPrm.stIOC.giBuzerCtrlFlag = 1) Then
                                    ' ブザー音2(ピ～ピッピ) (EXTOUT(OnBit, OffBit))
                                    Call ObjSys.EXTIO_Out_Sub(gSysPrm.stIOC.glBuzerCtrl_Out_Adr, gSysPrm.stIOC.glBuzerCtrl_Out2, gSysPrm.stIOC.glBuzerCtrl_Out1)
                                End If
                            End If

                            ' メッセージ表示 & STARTキー押下待ち
                            'r = Sub_ErrAxis(System.Math.Abs(r), giTrimErr)
                        End If

                    End If

                    ' エラーならメッセージを表示してエラーリターン
                    r = Form1.System1.EX_ZGETSRVSIGNAL(gSysPrm, r, 0)       ' エラーならメッセージを表示する
                    If (r = cFRS_ERR_EMG) Then                                  ' 非常停止検出 ?
                        GoTo STP_EMERGENCY                                      ' 非常停止メッセージ表示へ
                    End If

                    r = -1 * gMode                                              ' Return値 = = gMode(-xxxで戻る)

            End Select

            ' 終了処理
            Return (r)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "FrmReset.frmMessage_Main() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Function
#End Region

#Region "非常停止メッセージ表示"
    '''=========================================================================
    ''' <summary>非常停止メッセージ表示</summary>
    ''' <returns></returns>
    '''=========================================================================
    Private Function Sub_DispEmergencyMsg() As Integer

        Dim strMSG As String

        Try


            ' メッセージ表示
            Call Sub_SetMessage("非常停止しました", "装置内の基板を取り除いてください", "Cancelボタン押下でプログラムを終了します", System.Drawing.Color.Red, System.Drawing.Color.Red, System.Drawing.Color.Red)
            Me.Refresh()
            mExitFlag = cFRS_NORMAL                                     ' ExitFlg = 初期化
            BtnCancel.Visible = True                                    ' Cancelボタン表示

            ' Cancelボタンの押下を待つ
            Do
                System.Windows.Forms.Application.DoEvents()             ' メッセージポンプ
                Call System.Threading.Thread.Sleep(1)                   ' Wait(msec)
            Loop While (mExitFlag = cFRS_NORMAL)
            BtnCancel.Visible = False                                   ' Cancelボタン非表示

            Return (cFRS_ERR_EMG)                                       ' Retuen値 = 非常停止

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "FrmReset.Sub_DispEmergencyMsg() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)                                          ' Retuen値 = ﾄﾗｯﾌﾟｴﾗｰ発生
        End Try
    End Function
#End Region

#Region "フォームに表示するメッセージを設定する"
    '''=========================================================================
    ''' <summary>フォームに表示するメッセージを設定する</summary>
    ''' <param name="strMSG1">(INP)LblCaptionに表示する文字列</param>
    ''' <param name="strMSG2">(INP)Label1に表示する文字列</param>
    ''' <param name="strMSG3">(INP)Label2に表示する文字列</param>
    ''' <param name="Color1"> (INP)LblCaptionの文字の色</param>
    ''' <param name="Color2"> (INP)Label1の文字の色</param>
    ''' <param name="Color3"> (INP)Label2の文字の色</param>
    '''=========================================================================
    Private Sub Sub_SetMessage(ByVal strMSG1 As String, ByVal strMSG2 As String, ByVal strMSG3 As String, _
                                 ByVal Color1 As Object, ByVal Color2 As Object, ByVal Color3 As Object)

        Try
            ' メッセージ設定
            LblCaption.ForeColor = Color1
            Label1.ForeColor = Color2
            Label2.ForeColor = Color3
            LblCaption.Text = strMSG1
            Label1.Text = strMSG2
            Label2.Text = strMSG3
            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            MsgBox("FrmReset.Sub_SetMessage() TRAP ERROR = " + ex.Message)
        End Try
    End Sub
#End Region

#Region "Start/Resetキー押下待ちｻﾌﾞﾙｰﾁﾝ"
    '''=========================================================================
    ''' <summary>Start/Resetキー押下待ちｻﾌﾞﾙｰﾁﾝ</summary>
    ''' <param name="Md">(INP)cFRS_ERR_START                = STARTキー押下待ち
    '''                       cFRS_ERR_RST                  = RESETキー押下待ち
    '''                       cFRS_ERR_START + cFRS_ERR_RST = START/RESETキー押下待ち
    ''' </param>
    ''' <returns>cFRS_ERR_START = STARTキー押下
    '''          cFRS_ERR_RST   = RESETキー押下
    '''          上記以外=エラー
    ''' </returns>
    '''=========================================================================
    Private Function Sub_WaitStartRestKey(ByVal Md As Integer) As Integer

        Dim sts As Long = 0
        Dim r As Long = 0

        Try
            ' パラメータチェック
            If (Md = 0) Then
                Return (-1 * ERR_CMD_PRM)                           ' パラメータエラー
            End If

#If cOFFLINEcDEBUG Then                                             ' OffLineﾃﾞﾊﾞｯｸﾞON ?(↓FormResetが最前面表示なので下記のようにしないとMsgBoxが最前面表示されない)
            Dim Dr As System.Windows.Forms.DialogResult
            Dr = MessageBox.Show("START SW CHECK", "Debug", MessageBoxButtons.OKCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
            If (Dr = System.Windows.Forms.DialogResult.OK) Then
                mExitFlag = cFRS_ERR_START                          ' Return値 = STARTキー押下
            Else
                mExitFlag = cFRS_ERR_RST                            ' Return値 = RESETキー押下
            End If
            Return (mExitFlag)
#End If

            ' START/RESETキー押下待ち(Ok/Cancelボタンも有効)
            Call ZCONRST()                                          ' コンソールキーラッチ解除
            mExitFlag = 0
            Do
                r = STARTRESET_SWCHECK(False, sts)                  ' START/RESET SW押下チェック
                If (sts = cFRS_ERR_RST) And ((Md = cFRS_ERR_RST) Or (Md = cFRS_ERR_START + cFRS_ERR_RST)) Then
                    mExitFlag = cFRS_ERR_RST                        ' ExitFlag = Cancel(RESETキー)
                ElseIf (sts = cFRS_ERR_START) And ((Md = cFRS_ERR_START) Or (Md = cFRS_ERR_START + cFRS_ERR_RST)) Then
                    mExitFlag = cFRS_ERR_START                      ' ExitFlag = OK(STARTキー)
                End If

                System.Windows.Forms.Application.DoEvents()         ' メッセージポンプ
                Call System.Threading.Thread.Sleep(1)               ' Wait(msec)
                If ObjSys.EmergencySwCheck() Then                   ' 非常停止 ?
                    mExitFlag = cFRS_ERR_EMG                        ' Return値 = 非常停止検出
                End If
                Me.BringToFront()                                   ' 最前面に表示 

                'V2.2.0.0⑱↓
                If buzzerStop = True Then
                    Form1.System1.SetSignalTowerCtrl(Form1.System1.SIGNAL_BUZZER_OFF)
                    buzzerStop = False
                End If
                'V2.2.0.0⑱↑

            Loop While (mExitFlag = 0)

            Return (mExitFlag)

            ' トラップエラー発生時 
        Catch ex As Exception
            MsgBox("FrmReset.Sub_WaitRestKey() TRAP ERROR = " + ex.Message)
            Return (cERR_TRAP)                                      ' Retuen値 = ﾄﾗｯﾌﾟｴﾗｰ発生
        End Try
    End Function
#End Region

#Region "ステージを原点に戻す(残基板取り除くため)"
    '''=========================================================================
    ''' <summary>ステージを原点に戻す(残基板取り除くため)</summary>
    ''' <returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Private Function Sub_XY_OrgBack() As Integer

        Dim r As Integer
        Dim rtnCode As Integer = cFRS_NORMAL


        Try
            ' ステージを原点に戻す(XYZθ軸初期化)
            r = Form1.System1.EX_SYSINIT(gSysPrm, stPLT.Z_ZOFF, stPLT.Z_ZON)
            If (r <> cFRS_NORMAL) Then
                Return (r)
            End If

            Return (rtnCode)

            ' トラップエラー発生時 
        Catch ex As Exception
            MsgBox("FrmReset.Sub_XY_OrgBack() TRAP ERROR = " + ex.Message)
            Return (cERR_TRAP)
        End Try
    End Function
#End Region

#Region "トリミング中のｽﾗｲﾄﾞｶﾊﾞｰ開/筐体ｶﾊﾞｰ開メッセージ表示処理"
    '''=========================================================================
    ''' <summary>トリミング中のｽﾗｲﾄﾞｶﾊﾞｰ開/筐体ｶﾊﾞｰ開メッセージ表示</summary>
    ''' <param name="Mode">(INP)処理モード</param>
    ''' <returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Private Function Sub_CvrOpen(ByVal Mode As Integer) As Integer

        Dim r As Integer
        Dim rtnCode As Integer

        Try
            ' シグナルタワー制御(On=異常+ﾌﾞｻﾞｰ1, Off=全ﾋﾞｯﾄ) ###007
            Select Case (gSysPrm.stIOC.giSignalTower)
                Case SIGTOWR_NORMAL                                     ' 標準(赤点滅+ブザー１)
                    Call Form1.System1.SetSignalTower(SIGOUT_RED_BLK Or SIGOUT_BZ1_ON, &HFFFF)
                Case SIGTOWR_SPCIAL                                     ' 特注(赤点滅+ブザー１)
                    'r = Form1.System1.SetSignalTower(EXTOUT_RED_BLK Or EXTOUT_BZ1_ON, &HFFFF)
            End Select

            Call LAMP_CTRL(LAMP_START, False)                           ' STARTﾗﾝﾌﾟOFF

            ' メッセージ表示
            If (gMode = cGMODE_CVR_OPN) Then                            ' 筐体ｶﾊﾞｰ開 ?
                Call Sub_SetMessage("筐体カバーが開きました", "RESETキーを押すとプログラムを終了します", "", System.Drawing.Color.Red, System.Drawing.Color.Red, System.Drawing.SystemColors.ControlText)
                rtnCode = cFRS_ERR_CVR                                  ' Return値 = 筐体カバー開検出
            Else
                Call Sub_SetMessage("スライドカバーが開きました", "RESETキーを押すとプログラムを終了します", "", System.Drawing.Color.Red, System.Drawing.Color.Red, System.Drawing.SystemColors.ControlText)
                rtnCode = cFRS_ERR_SCVR                                 ' Return値 = スライドカバー開検出
            End If
            Me.Refresh()


            ' メッセージ表示してRESETｷｰ押下待ち
            r = Sub_WaitStartRestKey(cFRS_ERR_RST)                      ' RESETｷｰ押下待ち

            If (r = cFRS_ERR_EMG) Then                                  ' 非常停止検出 ?
                rtnCode = cFRS_ERR_EMG
                GoTo STP_END                                            ' 非常停止メッセージ表示へ
            End If

STP_END:
            Return (rtnCode)

            ' トラップエラー発生時 
        Catch ex As Exception
            MsgBox("FrmReset.Sub_CvrOpen() TRAP ERROR = " + ex.Message)
            Return (cERR_TRAP)                                      ' Retuen値 = ﾄﾗｯﾌﾟｴﾗｰ発生
        End Try
    End Function
#End Region

    ''' <summary>
    ''' ロットエンド時ブザーを終了するタイマー
    ''' </summary>
    Private Sub TimerLotEnd_Tick()

        Try

            buzzerStop = True
            TimerLotEnd.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)
            TimerLotEnd.Dispose()                                           ' タイマーを破棄する


        Catch ex As Exception

        End Try

    End Sub


End Class