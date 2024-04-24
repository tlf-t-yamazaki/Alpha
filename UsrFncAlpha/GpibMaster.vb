'==============================================================================
'   Description : ＣＯＮＴＥＣ　ＧＰＩＢ通信用クラス
'
'   Copyright(C): TOWA LASERFRONT CORP. 2018
'
'==============================================================================
Option Strict Off
Option Explicit On 

Public Class GpibMaster
    '==========================================================================
    '   定数/変数定義
    '==========================================================================
#Region "ＣＯＮＴＥＣ　ＧＰＩＢ通信用ＤＬＬ定義"
    '--------------------------------------------------------------------------
    '   ＣＯＮＴＥＣ　ＧＰＩＢ通信用ＤＬＬ定義
    '--------------------------------------------------------------------------
    ' Common
    Declare Function GpibInit Lib "CGPIB.DLL" (ByVal strDeviceName As String, ByRef shDevId As Short) As Integer
    Declare Function GpibExit Lib "CGPIB.DLL" (ByVal shDevId As Short) As Integer
    Declare Function GpibGetErrorString Lib "CGPIB.DLL" (ByVal intErrorCode As Integer, ByVal strErrorString As System.Text.StringBuilder) As Integer

    ' Initialization
    Declare Function GpibSetEquipment Lib "CGPIB.DLL" (ByVal shDevId As Short, ByRef shEqpId As Short, ByVal shPrmAddr As Short, ByVal shScdAddr As Short, ByVal shDelim As Short) As Integer
    Declare Function GpibGetEquipment Lib "CGPIB.DLL" (ByVal shEqpId As Short, ByRef shPrmAddr As Short, ByRef shScdAddr As Short, ByRef shDelim As Short) As Integer
    Declare Function GpibSendIFC Lib "CGPIB.DLL" (ByVal shId As Short, ByVal shIfcTime As Short) As Integer
    Declare Function GpibChangeREN Lib "CGPIB.DLL" (ByVal shId As Short, ByVal shEnable As Short) As Integer
    Declare Function GpibEnableRemote Lib "CGPIB.DLL" (ByVal shId As Short) As Integer

    ' Send and Reception
    Declare Function GpibSendData Lib "CGPIB.DLL" (ByVal shId As Short, ByRef intSendLen As Integer, ByVal strSendBuf As System.Text.StringBuilder) As Integer
    Declare Function GpibRecData Lib "CGPIB.DLL" (ByVal shId As Short, ByRef intRecLen As Integer, ByVal strRecBuf As System.Text.StringBuilder) As Integer
    Declare Function GpibSetDelim Lib "CGPIB.DLL" (ByVal shId As Short, ByVal shDelim As Short, ByVal shEoi As Short, ByVal shEos As Short) As Integer
    Declare Function GpibGetDelim Lib "CGPIB.DLL" (ByVal shId As Short, ByRef shDelim As Short, ByRef shEoi As Short, ByRef shEos As Short) As Integer
    Declare Function GpibSetTimeOut Lib "CGPIB.DLL" (ByVal shId As Short, ByVal intTimeOut As Integer) As Integer
    Declare Function GpibGetTimeOut Lib "CGPIB.DLL" (ByVal shId As Short, ByRef intTimeOut As Integer) As Integer
    Declare Function GpibSetEscape Lib "CGPIB.DLL" (ByVal shId As Short, ByVal shEnable As Short, ByVal shKeyType As Short, ByVal intKeyCode As Integer) As Integer
    Declare Function GpibGetEscape Lib "CGPIB.DLL" (ByVal shId As Short, ByRef shEnable As Short, ByVal shKeyType As Short, ByRef intKeyCode As Integer) As Integer
    Declare Function GpibSetSlowMode Lib "CGPIB.DLL" (ByVal shId As Short, ByVal shSlowTime As Short) As Integer
    Declare Function GpibGetSlowMode Lib "CGPIB.DLL" (ByVal shId As Short, ByRef shSlowTime As Short) As Integer
    Declare Function GpibSetSmoothMode Lib "CGPIB.DLL" (ByVal shId As Short, ByVal shMode As Short) As Integer
    Declare Function GpibGetSmoothMode Lib "CGPIB.DLL" (ByVal shId As Short, ByRef shMode As Short) As Integer
    Declare Function GpibSetAddrInfo Lib "CGPIB.DLL" (ByVal shId As Short, ByVal shTalker As Short, ByRef shListenerArray As Short) As Integer
    Declare Function GpibGetAddrInfo Lib "CGPIB.DLL" (ByVal shId As Short, ByRef shTalker As Short, ByRef shListenerArray As Short) As Integer

    ' Serial poll
    Declare Function GpibSPoll Lib "CGPIB.DLL" (ByVal shEqpId As Short, ByRef shStb As Short, ByRef shSrq As Short) As Integer
    Declare Function GpibSPollAll Lib "CGPIB.DLL" (ByVal shId As Short, ByRef shAddrArray As Short, ByRef shStbArray As Short, ByRef shSrqArray As Short) As Integer

    ' Parallel poll
    Declare Function GpibPPoll Lib "CGPIB.DLL" (ByVal shDevId As Short, ByRef shPpr As Short) As Integer
    Declare Function GpibSetPPoll Lib "CGPIB.DLL" (ByVal shDevId As Short, ByRef shAddrArray As Short, ByRef shDataLineArray As Short, ByRef shPolarityArray As Short) As Integer
    Declare Function GpibGetPPoll Lib "CGPIB.DLL" (ByVal shDevId As Short, ByRef shAddrArray As Short, ByRef shDataLineArray As Short, ByRef shPolarityArray As Short) As Integer
    Declare Function GpibResetPPoll Lib "CGPIB.DLL" (ByVal shDevId As Short, ByRef shAddrArray As Short) As Integer

    ' SRQ response
    Declare Function GpibSetPPollResponse Lib "CGPIB.DLL" (ByVal shDevId As Short, ByVal shResponse As Short) As Integer
    Declare Function GpibGetPPollResponse Lib "CGPIB.DLL" (ByVal shDevId As Short, ByRef shResponse As Short) As Integer
    Declare Function GpibSetIst Lib "CGPIB.DLL" (ByVal shDevId As Short, ByVal shIst As Short) As Integer
    Declare Function GpibGetIst Lib "CGPIB.DLL" (ByVal shDevId As Short, ByRef shIst As Short) As Integer
    Declare Function GpibSendSRQ Lib "CGPIB.DLL" (ByVal shDevId As Short, ByVal shSrqSend As Short, ByVal shStb As Short) As Integer
    Declare Function GpibCheckSPoll Lib "CGPIB.DLL" (ByVal shDevId As Short, ByRef shSPoll As Short, ByRef shStb As Short) As Integer

    ' Control
    Declare Function GpibSendCommands Lib "CGPIB.DLL" (ByVal shDevId As Short, ByRef shCmdArray As Short) As Integer
    Declare Function GpibSendTrigger Lib "CGPIB.DLL" (ByVal shId As Short) As Integer
    Declare Function GpibSendDeviceClear Lib "CGPIB.DLL" (ByVal shId As Short) As Integer
    Declare Function GpibChangeLocal Lib "CGPIB.DLL" (ByVal shId As Short) As Integer
    Declare Function GpibSendLocalLockout Lib "CGPIB.DLL" (ByVal shDevId As Short) As Integer

    ' Status
    Declare Function GpibSetStatus Lib "CGPIB.DLL" (ByVal shId As Short, ByVal shSelect As Short, ByVal intData As Integer) As Integer
    Declare Function GpibGetStatus Lib "CGPIB.DLL" (ByVal shId As Short, ByVal shSelect As Short, ByRef intData As Integer) As Integer
    Declare Function GpibReadLines Lib "CGPIB.DLL" (ByVal shId As Short, ByVal shSelect As Short, ByRef shLineStatus As Short) As Integer
    Declare Function GpibFindListener Lib "CGPIB.DLL" (ByVal shDevId As Short, ByVal shPrmAddr As Short, ByVal shScdAddr As Short, ByRef shArraySize As Short, ByRef shAddrArray As Short) As Integer
    Declare Function GpibGetSignal Lib "CGPIB.DLL" (ByVal shDevId As Short, ByRef intwParam As Integer, ByRef intlParam As Integer) As Integer

    'Event
    Declare Function GpibSetNotifySignal Lib "CGPIB.DLL" (ByVal shDevId As Short, ByVal hWnd As Integer, ByVal intNotifySignalMask As Integer) As Integer
    Declare Function GpibGetNotifySignal Lib "CGPIB.DLL" (ByVal shDevId As Short, ByRef hWnd As Integer, ByRef intNotifySignalMask As Integer) As Integer
    Declare Function GpibSetNotifyMessage Lib "CGPIB.DLL" (ByVal shDevId As Short, ByVal intMessage As Integer) As Integer
    Declare Function GpibGetNotifyMessage Lib "CGPIB.DLL" (ByVal shDevId As Short, ByRef intMessage As Integer) As Integer
#End Region

    '=========================================================================
    '   メソッド定義
    '=========================================================================
#Region "初期化処理"
    '''=========================================================================
    ''' <summary>初期化処理</summary>
    ''' <param name="strDeviceName">(INP)デバイス名(例)"GPIB000")</param>
    ''' <param name="mintDevId">    (OUT)デバイスID</param>
    ''' <returns>0 = 正常, 0以外 = エラー</returns>
    ''' <remarks>デバイス名からデバイスIDを取得する</remarks>
    '''=========================================================================
    Public Function Gpib_Init(ByVal strDeviceName As String, ByRef mintDevId As Short) As Integer

        Dim strTmp As System.Text.StringBuilder
        Dim lngData As Integer
        Dim mlngRet As Integer
        Dim strMSG As String

        Try
            ' 初期化関数実行
            mlngRet = GpibInit(strDeviceName, mintDevId)                ' 初期化
            If (mlngRet <> 0) Then                                      ' エラー ?
                strTmp = New System.Text.StringBuilder("", 256)
                Call GpibGetErrorString(mlngRet, strTmp)                ' メッセージ取得 
                MsgBox("Gpib_Init:" & strTmp.ToString)

            Else ' 正常にデバイスIDを取得できた場合
                strTmp = New System.Text.StringBuilder("GpibInit : デバイスID = " & mintDevId)
                ' 機器アドレスの取得
                mlngRet = GpibGetStatus(mintDevId, &H8S, lngData)
                ' プロパティでマスタに設定されているか確認
                mlngRet = GpibGetStatus(mintDevId, &HAS, lngData)
                ' マスタに設定されているか確認
                If (mlngRet = 0) And (lngData = 1) Then
                    'Call MsgBox("マスタに設定されていません。プロパティでマスタに設定してください。", MsgBoxStyle.OkOnly)
                End If
            End If

            ' IFCを送出
            GpibSendIFC(mintDevId, 1)
            Return (mlngRet)                                            ' デバイスIDを返す

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "GpibMaster.Gpib_Init() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (-1)                                                 ' Return値 = ﾄﾗｯﾌﾟｴﾗｰ発生
        End Try
    End Function
#End Region

#Region "終了処理"
    '''=========================================================================
    ''' <summary>終了処理</summary>
    ''' <param name="mintDevId">(INP)デバイスID</param>
    ''' <returns>0 = 正常, 0以外 = エラー</returns>
    ''' <remarks>デバイスIDを無効にする</remarks>
    '''=========================================================================
    Public Function Gpib_Term(ByVal mintDevId As Short) As Integer

        Dim strTmp As System.Text.StringBuilder
        Dim mlngRet As Integer
        Dim strMSG As String

        Try
            ' 終了関数実行
            If (mintDevId < 0) Then Return (0) '                        ' デバイスIDなしならNOP
            mlngRet = GpibExit(mintDevId)                               ' 終了
            If (mlngRet <> 0) Then                                      ' エラー ?
                strTmp = New System.Text.StringBuilder("", 256)
                Call GpibGetErrorString(mlngRet, strTmp)
                MsgBox(strTmp.ToString)
            End If
            Return (mlngRet)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "GpibMaster.Gpib_Term() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (-1)                                                 ' Return値 = ﾄﾗｯﾌﾟｴﾗｰ発生
        End Try
    End Function
#End Region

#Region "データ送信処理"
    '''=========================================================================
    ''' <summary>データ送信処理</summary>
    ''' <param name="strSendDat">(INP)送信データ</param>
    ''' <param name="mintDevId"> (INP)デバイスID</param>
    ''' <param name="Addr">      (INP)GPIBアドレス</param>
    ''' <param name="Delim">     (INP)デリミタ設定(0:使用しない, 1:CR+LF, 2:CR, 3:LF)</param>
    ''' <param name="Eoi">       (INP)EOI(0:出力しない, 0以外:出力する)</param>
    ''' <remarks>スレーブ機器にデータを送信する</remarks>
    '''=========================================================================
    Public Function Gpib_Send(ByVal strSendDat As String, ByVal mintDevId As Short, ByVal Addr As Short, ByVal Delim As Short, ByVal Eoi As Short) As Integer

        Dim strTmp As System.Text.StringBuilder
        Dim strSendBuf As System.Text.StringBuilder
        Dim lngSendLen As Integer
        Dim mlngRet As Integer
        Dim EqpId As Short
        Dim strMSG As String

        Try
            ' 相手機器の設定
            mlngRet = funcSetParam(mintDevId, Addr, Delim, Eoi, EqpId)
            If (mlngRet <> 0) Then                                      ' エラー ?
                Return (mlngRet)
            End If

            ' 送信データチェック
            strSendBuf = New System.Text.StringBuilder(strSendDat)
            lngSendLen = strSendBuf.Length
            If (lngSendLen = 0) Then
                strTmp = New System.Text.StringBuilder("Gpib_Send : 送信データがありません")
                'MsgBox(strTmp.ToString)
                Return (1)
            End If

            ' 送信実行
            mlngRet = GpibSendData(EqpId, lngSendLen, strSendBuf)
            If (mlngRet <> 0) Then                                      ' エラー ?
                strTmp = New System.Text.StringBuilder("", 256)
                Call GpibGetErrorString(mlngRet, strTmp)
                'MsgBox(strTmp.ToString)
            End If
            Return (mlngRet)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "GpibMaster.Gpib_Send() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (-1)                                                 ' Return値 = ﾄﾗｯﾌﾟｴﾗｰ発生
        End Try
    End Function
#End Region

#Region "データ受信処理"
    '''=========================================================================
    ''' <summary>データ受信処理</summary>
    ''' <param name="strRecDat">(OUT)受信データ</param>
    ''' <param name="mintDevId">(INP)デバイスID</param>
    ''' <param name="Addr">     (INP)GPIBアドレス</param>
    ''' <param name="Delim">    (INP)デリミタ設定(0:使用しない, 1:CR+LF, 2:CR, 3:LF)</param>
    ''' <param name="Eoi">      (INP)EOI(0:出力しない, 0以外:出力する)</param>
    ''' <remarks>スレーブ機器からデータを受信する</remarks>
    '''=========================================================================
    Public Function Gpib_Recv(ByRef strRecDat As String, ByVal mintDevId As Short, ByVal Addr As Short, ByVal Delim As Short, ByVal Eoi As Short) As Integer

        Dim strTmp As System.Text.StringBuilder
        Dim lngRecLen As Integer
        Dim strRecBuf As System.Text.StringBuilder
        Dim mlngRet As Integer
        Dim EqpId As Short
        Dim strDAT As String
        Dim strMSG As String

        Try
            mlngRet = funcSetParam(mintDevId, Addr, Delim, Eoi, EqpId)
            If (mlngRet <> 0) Then                                      ' エラー ?
                Return (mlngRet)
            End If

            ' 受信実行
            lngRecLen = 256
            strRecBuf = New System.Text.StringBuilder("", 256)
            mlngRet = GpibRecData(EqpId, lngRecLen, strRecBuf)
            strRecBuf.Length = lngRecLen

            ' 戻り値処理
            If (mlngRet <> 0) Then                                      ' エラー ?
                strTmp = New System.Text.StringBuilder("", 256)
                Call GpibGetErrorString(mlngRet, strTmp)
            Else
                '表示
                strRecBuf.Length = lngRecLen
                'strRecDat = VB.Left(strRecBuf.ToString, lngRecLen) '受信データ設定
                strDAT = strRecBuf.ToString                             ' 受信データ取得
                strRecDat = strDAT.Substring(0, lngRecLen)              ' 受信データを返す
                strTmp = New System.Text.StringBuilder("「データの受信」正常終了")
            End If
            Return (mlngRet)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "GpibMaster.Gpib_Recv() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (-1)                                                 ' Return値 = ﾄﾗｯﾌﾟｴﾗｰ発生
        End Try
    End Function
#End Region

#Region "データ受信処理(Double型データで返す)"
    '''=========================================================================
    ''' <summary>データ受信処理</summary>
    ''' <param name="dblRecDat">(OUT)受信データ</param>
    ''' <param name="mintDevId">(INP)デバイスID</param>
    ''' <param name="Addr">     (INP)GPIBアドレス</param>
    ''' <param name="Delim">    (INP)デリミタ設定(0:使用しない, 1:CR+LF, 2:CR, 3:LF)</param>
    ''' <param name="Eoi">      (INP)EOI(0:出力しない, 0以外:出力する)</param>
    ''' <remarks>スレーブ機器からデータを受信する</remarks>
    '''=========================================================================
    Public Function Gpib_RVal(ByRef dblRecDat As Double, ByVal mintDevId As Short, ByVal Addr As Short, ByVal Delim As Short, ByVal Eoi As Short) As Integer

        Dim strTmp As System.Text.StringBuilder
        Dim lngRecLen As Integer
        Dim strRecBuf As System.Text.StringBuilder
        Dim mlngRet As Integer
        Dim EqpId As Short
        Dim strDAT As String
        Dim strMSG As String

        Try
            dblRecDat = 0.0
            mlngRet = funcSetParam(mintDevId, Addr, Delim, Eoi, EqpId)
            If (mlngRet <> 0) Then                                      ' エラー ?
                Return (mlngRet)
            End If

            ' 受信実行
            lngRecLen = 256
            strRecBuf = New System.Text.StringBuilder("", 256)
            mlngRet = GpibRecData(EqpId, lngRecLen, strRecBuf)
            strRecBuf.Length = lngRecLen

            ' 戻り値処理
            If (mlngRet <> 0) Then                                      ' エラー ?
                strTmp = New System.Text.StringBuilder("", 256)
                Call GpibGetErrorString(mlngRet, strTmp)
            Else
                ' 表示
                If (lngRecLen <= 0) Then

                Else

                    strRecBuf.Length = lngRecLen
                    strDAT = strRecBuf.ToString                         ' 受信データ取得
                    strMSG = strDAT.Substring(0, lngRecLen)             ' 

                    dblRecDat = Double.Parse(strMSG)                    ' 受信データをDouble型データで返す 
                End If
                strTmp = New System.Text.StringBuilder("「データの受信」正常終了")
            End If
            Return (mlngRet)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "GpibMaster.Gpib_RVal() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (-1)                                                 ' Return値 = ﾄﾗｯﾌﾟｴﾗｰ発生
        End Try
    End Function
#End Region

#Region "相手機器の設定"
    '''=========================================================================
    ''' <summary>相手機器の設定</summary>
    ''' <param name="mintDevId">(INP)デバイスID</param>
    ''' <param name="Addr">     (INP)GPIBアドレス</param>
    ''' <param name="Delim">    (INP)デリミタ設定(0:使用しない, 1:CR+LF, 2:CR, 3:LF)</param>
    ''' <param name="Eoi">      (INP)EOI(0:出力しない, 0以外:出力する)</param>
    ''' <param name="EqpId">    (OUT)イクウィップメントID</param>
    ''' <returns>0 = 正常, 0以外 = エラー</returns>
    ''' <remarks>イクウィップメントID・デリミタ・EOIの設定を行います</remarks>
    '''=========================================================================
    Private Function funcSetParam(ByVal mintDevId As Short, ByVal Addr As Short, ByVal Delim As Short, ByVal Eoi As Short, ByRef EqpId As Short) As Integer

        Dim strTmp As System.Text.StringBuilder
        Dim mlngRet As Integer
        Dim strMSG As String

        Try
            ' 機器アドレスのイクウィップメントID(mintSelectEqpId)を取得
            mlngRet = GpibSetEquipment(mintDevId, EqpId, Addr, 0, Delim)
            If (mlngRet <> 0) Then                                      ' エラー ?
                strTmp = New System.Text.StringBuilder("", 256)
                Call GpibGetErrorString(mlngRet, strTmp)
                'MsgBox(strTmp.ToString)
                Return (mlngRet)
            End If

            ' デリミタ・EOIの設定
            mlngRet = GpibSetDelim(EqpId, Delim, Eoi, 0)
            If (mlngRet <> 0) Then                                      ' エラー ?
                strTmp = New System.Text.StringBuilder("", 256)
                Call GpibGetErrorString(mlngRet, strTmp)
                'MsgBox(strTmp.ToString)
            End If
            Return (mlngRet)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "GpibMaster.funcSetParam() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (-1)                                                 ' Return値 = ﾄﾗｯﾌﾟｴﾗｰ発生
        End Try
    End Function
#End Region

End Class
