'===============================================================================
'   Description  : トリマー加工条件をFL側よりRS232Cで送受信する
'                  (C#で作成された｢DllFLCom.dll｣を使用)
'                   ※対象機器　FL用FPGA設定用
'                     通信方式       : 調歩同期式半二重通信
'                     ボーレート     : 38,400BPS
'                     キャラクター長 : 8 Bit
'                     パリティ       : なし
'                     ストップ       : 1 BIT
'                     デリミタコード : CR
'
'                   送受信データ形式
'                   1. 送信データ形式(PC　←→ FL)
'                   -------------------------------------------
'                   |コマンド名(2) | データ(4)        | CR(1) |
'                   |A～PのASCII   |0～9,a～fのASCII  |       |
'                   -------------------------------------------
'                   2. 応答要求データ形式(PC　→ FL)
'                   -----------------------------------
'                   |コマンド名(2) | ﾘｰﾄﾞ(1) |　CR(1) |
'                   |A～PのASCII   |  r      |　      |
'                   -----------------------------------
'   Copyright(C) : TOWA LASERFRONT CORP. 2011
'
'   Remarks      : 下記DLL(C#)を使用(「参照の追加」で追加)
'          　　　　DllFLCom.dll, DllSerialIO.dll, DllCndXMLIO.dll
'
'===============================================================================
#If cOSCILLATORcFLcUSE Then
Option Strict Off
Option Explicit On
Module Rs232c
#Region "変数定義"
    '===========================================================================
    '   変数定義
    '===========================================================================
    '----- ポート情報 -----
    Public Structure ComInfo
        Dim PortName As String                                  ' ＲＳ２３２Ｃポート番号
        Dim BaudRate As Long                                    ' 伝送スピード
        Dim Parity As Integer                                   ' パリティ(0:None, 1:Odd, 2:Even, 3:Mark, 4:Space)
        Dim DataBits As Integer                                 ' データー長
        Dim StopBits As Integer                                 ' ストップビット長
    End Structure

    '----- クラスオブジェクト -----
    Private ObjPortInfo As DllSerialIO.PortInformation          ' シリアルポート情報オブジェク(DllSerialIO.dll)
    Private ObjFLCom As DllFLCom.FLComIO                        ' トリマー加工条件送受信オブジェクト(DllFLCom.dll)

    '----- バージョン情報クラス -----
    Private ObjPortCVer As DllSerialIO.VersionInformation       ' DllSerialIO.dll
    Private ObjFLXMLVer As DllCndXMLIO.VersionInformation       ' DllCndXMLIO.dll
    Private ObjFLComVer As DllFLCom.VersionInformation          ' DllFLCom.dll

    '----- トリマー加工条件 -----
    Public Const MAX_BANK_NUM As Integer = 32                   ' 最大加工条件数(0-31)
    Public Const MAX_STEG_NUM As Integer = 20                   ' STEG波形最大値(1-20)
    Public Const MAX_CURR_VAL As Integer = 8500                 ' 最大電流値(mA)
    Public Const MIN_CURR_VAL As Integer = 1                    ' 最小電流値(mA)
    Public Const MAX_FREQ_VAL As Integer = 100                  ' 最大周波数(KHz)
    Public Const MIN_FREQ_VAL As Integer = 1                    ' 最小周波数(KHz)

    Public Structure TrimCondInfo                               ' トリマー加工条件形式定義
        Dim Curr() As Integer                                   ' 電流値(mA)
        Dim Freq() As Double                                    ' 周波数(KHz)
        Dim Steg() As Integer                                   ' STEG波形
        Dim dblPowerAdjustTarget() As Double                    ' 調整目標パワー1～n(0～29)
        Dim dblPowerAdjustToleLevel() As Double                 ' パワー調整許容範囲1～n(0～29)

        ' 構造体の初期化
        Public Sub Initialize()
            ReDim Curr(MAX_BANK_NUM - 1)                        ' 配列(0-31) 
            ReDim Freq(MAX_BANK_NUM - 1)
            ReDim Steg(MAX_BANK_NUM - 1)
            ReDim dblPowerAdjustTarget(MAX_BANK_NUM - 1)        ' 調整目標パワー1～n(0～29)
            ReDim dblPowerAdjustToleLevel(MAX_BANK_NUM - 1)     ' パワー調整許容範囲1～n(0～29)
        End Sub
    End Structure

    '---------------------------------------------------------------------------
    '   エラーコード(C#で作成されたdllが返しているもの)
    '---------------------------------------------------------------------------
    Public Enum SerialErrorCode
        '----- 1-18はDllSerialIOで使用 -----
        rRS_OK = 0                                              '  0:正常
        rRS_ReadTimeout                                         '  1:リードタイムアウト
        rRS_WriteTimeout                                        '  2:ライトタイムアウト
        rRS_RespomseTimeout                                     '  3:応答タイムアウト
        rRS_FailOpen                                            '  4:ｼﾘｱﾙﾎﾟｰﾄｵｰﾌﾟﾝ失敗
        rRS_FailClose                                           '  5:ｼﾘｱﾙﾎﾟｰﾄｸﾛｰｽﾞ失敗
        rRS_FailInit                                            '  6:ｼﾘｱﾙﾎﾟｰﾄ初期化失敗
        rRS_SerialErrorFrame                                    '  7:H/Wでﾌﾚｰﾑｴﾗｰ検出
        rRS_SerialErrorOverrun                                  '  8:文字ﾊﾞｯﾌｧのｵｰﾊﾞｰﾗﾝ発生
        rRS_SerialErrorRXOver                                   '  9:入力ﾊﾞｯﾌｧのｵｰﾊﾞｰﾌﾛｰ発生
        rRS_SerialErrorRXParity                                 ' 10:H/Wでﾊﾟﾘﾃｨｴﾗｰ発生
        rRS_SerialErrorTXFull                                   ' 11:ｱﾌﾟﾘｹｰｼｮﾝは文字を送信しようとしたが出力ﾊﾞｯﾌｧが一杯
        rRS_InvalidSerialProtInfo                               ' 12:シリアルポート情報不正
        rRS_InvalidValue                                        ' 13:無効なデータ
        rRS_FailSerialRead                                      ' 14:ｼﾘｱﾙﾎﾟｰﾄからの読込失敗
        rRS_FailSerialWrite                                     ' 15:ｼﾘｱﾙﾎﾟｰﾄへの書込失敗
        rRS_NotOpen                                             ' 16:ｼﾘｱﾙﾎﾟｰﾄがｵｰﾌﾟﾝしていない
        rRS_Exception                                           ' 17:例外

        '----- 以降は当関数で使用 -----
        rRS_FLCND_NONE = 101                                    ' 101:加工条件の設定なし
        rRS_FLCND_XMLNONE = 102                                 ' 102:加工条件ファイルが存在しない
        rRS_FLCND_XMLREADERR = 103                              ' 103:加工条件ファイルリードエラー
        rRS_FLCND_XMLWRITERR = 104                              ' 104:加工条件ファイルライトエラー
        rRS_FLCND_SNDERR = 105                                  ' 105:加工条件送信エラー
        rRS_FLCND_RCVERR = 106                                  ' 106:加工条件受信エラー

        '----- 以降はDllFLComで使用 -----
        rRS_CndNum = 900                                        ' 900:加工条件番号エラー
        rRS_Trap = 999                                          ' 999:トラップエラー発生

    End Enum

    '----- その他 -----
    Private Rs_Flag As Integer                                  ' ﾌﾗｸﾞ(0:未ｵｰﾌﾟﾝ, 1:ｵｰﾌﾟﾝ済)

#End Region

#Region "RS232C用メソッド"
#Region "FL用加工条件ファイルをリードしてFL側へ加工条件を送信する"
    '''=========================================================================
    ''' <summary>FL用加工条件ファイルをリードしてFL側へ加工条件を送信する</summary>
    ''' <param name="stCND">   (OUT)トリマー加工条件構造体</param>
    ''' <param name="DatFName">(INP)データファイル名</param>
    ''' <param name="CndFName">(OUT)ＦＬ用加工条件ファイル名</param>
    ''' <returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Public Function SendTrimCondInfToFL(ByVal stCND As TrimCondInfo, ByRef DatFName As String, ByRef CndFName As String) As Integer

        Dim r As Integer
        Dim strMSG As String

        Try
            ' FL(ﾌｧｲﾊﾞｰﾚｰｻﾞ) でなければ NOP
            If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then
                Return (SerialErrorCode.rRS_OK)
            End If

            ' ＦＬ用加工条件ファイル名を取得する(ファイル存在チェックあり)
            r = GetFLCndFileName(DatFName, CndFName, True)
            If (r <> SerialErrorCode.rRS_OK) Then
                Return (r)
            End If

            ' ＦＬ用加工条件を読込む
            r = ReadFLCndFile(stCND, CndFName)
            If (r <> SerialErrorCode.rRS_OK) Then
                Return (r)
            End If

            ' FL側へ加工条件を送信する
            r = TrimCondInfoSnd(stCND)
            If (r <> SerialErrorCode.rRS_OK) Then
                Return (r)
            End If
            Return (SerialErrorCode.rRS_OK)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Rs232c.SendTrimCondInfToFL() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)
        End Try
    End Function
#End Region

#Region "FL側から現在の加工条件を受信してFL用加工条件ファイルをライトする"
    '''=========================================================================
    ''' <summary>FL側から現在の加工条件を受信してFL用加工条件ファイルをライトする</summary>
    ''' <param name="stCND">   (OUT)トリマー加工条件構造体</param>
    ''' <param name="DatFName">(INP)データファイル名</param>
    ''' <param name="CndFName">(OUT)ＦＬ用加工条件ファイル名</param>
    ''' <returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Public Function RcvTrimCondInfToFL(ByVal stCND As TrimCondInfo, ByVal DatFName As String, ByRef CndFName As String) As Integer

        Dim r As Integer
        Dim strMSG As String

        Try
            ' FL(ﾌｧｲﾊﾞｰﾚｰｻﾞ) でなければ NOP
            If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then
                Return (SerialErrorCode.rRS_OK)
            End If

            ' FL側から現在の加工条件を受信する
            r = TrimCondInfoRcv(stCND)
            If (r <> SerialErrorCode.rRS_OK) Then
                Return (r)
            End If

            ' ＦＬ用加工条件ファイル名を取得する(ファイル存在チェックなし)
            r = GetFLCndFileName(DatFName, CndFName, False)
            If (r <> SerialErrorCode.rRS_OK) Then
                Return (r)
            End If

            ' ＦＬ用加工条件ファイルを書き込む
            r = WriteFLCndFile(stCND, CndFName)
            If (r <> SerialErrorCode.rRS_OK) Then
                Return (r)
            End If

            Return (SerialErrorCode.rRS_OK)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Rs232c.RcvTrimCondInfToFL() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)
        End Try
    End Function
#End Region

#Region "FL側へ加工条件を送信する"
    '''=========================================================================
    ''' <summary>FL側へ加工条件を送信する</summary>
    ''' <param name="stCND">(INP)トリマー加工条件構造体</param>
    ''' <returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Public Function TrimCondInfoSnd(ByVal stCND As TrimCondInfo) As Integer

        Dim wkCND As TrimCondInfo
        Dim r, i As Integer
        Dim strMSG As String

        Try
            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            '   FL側へ加工条件を送信する
            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            ' FL(ﾌｧｲﾊﾞｰﾚｰｻﾞ) でなければ NOP
            If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then
                Return (SerialErrorCode.rRS_OK)
            End If
#If cOFFLINEcDEBUG Then                                 ' DEBUG
            Return (SerialErrorCode.rRS_OK)
#End If
            ' ポート情報を設定する
            stCOM.PortName = "COM3"                             ' ポート番号
            stCOM.BaudRate = 38400                              ' Speed
            stCOM.Parity = 0                                    ' パリティ(0:None)
            stCOM.DataBits = 8                                  ' データ長 = 8 Bit
            stCOM.StopBits = 1                                  ' Stop Bit = 1 Bit

            ' ポートオープン
            r = Rs232c_Open(stCOM)
            If (r <> SerialErrorCode.rRS_OK) Then
                ' "シリアルポートＯＰＥＮエラー"
                strMSG = MSG_136 + "(" + stCOM.PortName + ")"
                Call MsgBox(strMSG, vbOKOnly)
                Return (r)
            End If

            ' シリアルポートへトリマー加工条件を一括で送信する
            r = RsSendBankALL(stCND.Curr, stCND.Freq, stCND.Steg)
            If (r <> SerialErrorCode.rRS_OK) Then
                ' "シリアルポート送信エラー"
                strMSG = MSG_138 + "(" + stCOM.PortName + ")"
                Call MsgBox(strMSG, vbOKOnly)
            End If

            ' シリアルポートからトリマー加工条件を一括で受信する
            wkCND = Nothing
            wkCND.Initialize()
            r = RsReceiveBankALL(wkCND.Curr, wkCND.Freq, wkCND.Steg, 10000)
            If (r <> SerialErrorCode.rRS_OK) Then
                ' "シリアルポート受信エラー"
                strMSG = MSG_139 + "(" + stCOM.PortName + ")"
                Call MsgBox(strMSG, vbOKOnly)
                Rs232c_Close()
                Return (r)
            End If

            ' ポートクローズ
            Rs232c_Close()

            ' 送信したトリマー加工条件が正しく設定されているか確認する
            For i = 0 To (MAX_BANK_NUM - 1)                     ' 最大加工条件数分確認する
                If (wkCND.Curr(i) <> stCND.Curr(i)) Then
                    Return (SerialErrorCode.rRS_FLCND_SNDERR)
                End If
                If (wkCND.Steg(i) <> stCND.Steg(i)) Then
                    Return (SerialErrorCode.rRS_FLCND_SNDERR)
                End If
            Next i

            Return (SerialErrorCode.rRS_OK)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Rs232c.TrimCondInfoSnd() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (SerialErrorCode.rRS_Trap)
        End Try
    End Function
#End Region

#Region "FL側から現在の加工条件を受信する"
    '''=========================================================================
    ''' <summary>FL側から現在の加工条件を受信する</summary>
    ''' <param name="stCND">(OUT)トリマー加工条件構造体</param>
    ''' <returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Public Function TrimCondInfoRcv(ByVal stCND As TrimCondInfo) As Integer

        Dim r As Integer
        Dim strMSG As String

        Try
            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            '   FL側から現在の加工条件を取得する
            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            ' FL(ﾌｧｲﾊﾞｰﾚｰｻﾞ) でなければ NOP
            If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then
                Return (SerialErrorCode.rRS_OK)
            End If

            ' ポート情報を設定する
            stCOM.PortName = "COM3"                             ' ポート番号
            stCOM.BaudRate = 38400                              ' Speed
            stCOM.Parity = 0                                    ' パリティ(0:None)
            stCOM.DataBits = 8                                  ' データ長 = 8 Bit
            stCOM.StopBits = 1                                  ' Stop Bit = 1 Bit

            ' ポートオープン
            r = Rs232c_Open(stCOM)
            If (r <> SerialErrorCode.rRS_OK) Then
                ' "シリアルポートＯＰＥＮエラー"
                strMSG = MSG_136 + "(" + stCOM.PortName + ")"
                Call MsgBox(strMSG, vbOKOnly)
                Return (r)
            End If

            ' シリアルポートからトリマー加工条件を一括で受信する
            r = RsReceiveBankALL(stCND.Curr, stCND.Freq, stCND.Steg, cTIMEOUT)
            If (r <> SerialErrorCode.rRS_OK) And (r <> SerialErrorCode.rRS_FLCND_NONE) Then
                ' "シリアルポート受信エラー"
                strMSG = MSG_139 + "(" + stCOM.PortName + ")"
                Call MsgBox(strMSG, vbOKOnly)
            End If

            ' ポートクローズ
            Rs232c_Close()
            Return (r)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Rs232c.TrimCondInfoRcv() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (SerialErrorCode.rRS_Trap)
        End Try
    End Function
#End Region

#Region "ＦＬ用加工条件ファイルを読込む"
    '''=========================================================================
    ''' <summary>ＦＬ用加工条件ファイルを読込む</summary>
    ''' <param name="stCND">(OUT)トリマー加工条件構造体</param>
    ''' <param name="FName">(INP)加工条件ファイル名</param>
    ''' <returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Public Function ReadFLCndFile(ByVal stCND As TrimCondInfo, ByVal FName As String) As Integer

        Dim r As Boolean
        Dim strMSG As String

        Try
            ' FL(ﾌｧｲﾊﾞｰﾚｰｻﾞ) でなければ NOP
            If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then
                Return (SerialErrorCode.rRS_OK)
            End If

            ' ＦＬ用加工条件ファイルを加工条件構造体に読込む()
            Dim ObjFLXMLIO As DllCndXMLIO.CndXMLIO = New DllCndXMLIO.CndXMLIO           ' 加工条件ファイルＩＯオブジェクト
            r = ObjFLXMLIO.Read_CndXMLFile(FName, stCND.Curr, stCND.Freq, stCND.Steg)
            If (r <> True) Then
                Return (SerialErrorCode.rRS_FLCND_XMLREADERR)
            End If
            Return (SerialErrorCode.rRS_OK)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Rs232c.ReadFLCndFile() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (SerialErrorCode.rRS_Trap)
        End Try
    End Function
#End Region

#Region "ＦＬ用加工条件ファイルを書き込む"
    '''=========================================================================
    ''' <summary>ＦＬ用加工条件ファイルを書き込む</summary>
    ''' <param name="stCND">(INP)トリマー加工条件構造体</param>
    ''' <param name="FName">(INP)加工条件ファイル名</param>
    ''' <returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Public Function WriteFLCndFile(ByVal stCND As TrimCondInfo, ByVal FName As String) As Integer

        Dim r As Boolean
        Dim strMSG As String

        Try
            ' FL(ﾌｧｲﾊﾞｰﾚｰｻﾞ) でなければ NOP
            If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then
                Return (SerialErrorCode.rRS_OK)
            End If

            ' ＦＬ用加工条件ファイルを書き込む
            Dim ObjFLXMLIO As DllCndXMLIO.CndXMLIO = New DllCndXMLIO.CndXMLIO           ' 加工条件ファイルＩＯオブジェクト
            r = ObjFLXMLIO.Write_CndXMLFile(FName, stCND.Curr, stCND.Freq, stCND.Steg)
            If (r <> True) Then
                Return (SerialErrorCode.rRS_FLCND_XMLWRITERR)
            End If
            Return (SerialErrorCode.rRS_OK)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Rs232c.WriteFLCndFile() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (SerialErrorCode.rRS_Trap)
        End Try
    End Function
#End Region

#Region "ＦＬ用加工条件ファイル名を取得する"
    '''=========================================================================
    ''' <summary>ＦＬ用加工条件ファイル名を取得する</summary>
    ''' <param name="InpFName">(INP)データファイル名</param>
    ''' <param name="OutFName">(OUT)ＦＬ用加工条件ファイル名</param>
    ''' <param name="Flg">     (INP)加工条件ファイルの存在チェックの有無</param>
    ''' <returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Public Function GetFLCndFileName(ByRef InpFName As String, ByRef OutFName As String, ByVal Flg As Boolean) As Integer

        Dim len As Integer
        Dim strMSG As String

        Try
            ' FL(ﾌｧｲﾊﾞｰﾚｰｻﾞ) でなければ NOP
            If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then
                Return (SerialErrorCode.rRS_OK)
            End If

            ' データファイル名の拡張子を".xml"に変換して返す
            len = InpFName.Length
            OutFName = InpFName.Substring(0, len - 4)
            OutFName = OutFName + ".xml"

            ' データファイルの存在チェック
            If (Flg = True) Then
                If (System.IO.File.Exists(OutFName) = False) Then
                    Return (SerialErrorCode.rRS_FLCND_XMLNONE)
                End If
            End If
            Return (SerialErrorCode.rRS_OK)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Rs232c.GetFLCndFileName() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (SerialErrorCode.rRS_Trap)
        End Try
    End Function
#End Region

#Region "FL側からエラー情報を受信する"
    '''=========================================================================
    ''' <summary>FL側からエラー情報を受信する</summary>
    ''' <returns>0=正常, 0以外=エラー</returns>
    '''=========================================================================
    Public Function ReceiveErrInfo(ByRef ErrInf As Integer) As Integer

        Dim r As Integer
        Dim strMSG As String

        Try
            ' FL(ﾌｧｲﾊﾞｰﾚｰｻﾞ) でなければ NOP
            ErrInf = 0
            If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then
                Return (SerialErrorCode.rRS_OK)
            End If

            ' ポート情報を設定する
            stCOM.PortName = "COM3"                             ' ポート番号
            stCOM.BaudRate = 38400                              ' Speed
            stCOM.Parity = 0                                    ' パリティ(0:None)
            stCOM.DataBits = 8                                  ' データ長 = 8 Bit
            stCOM.StopBits = 1                                  ' Stop Bit = 1 Bit

            ' ポートオープン
            r = Rs232c_Open(stCOM)
            If (r <> SerialErrorCode.rRS_OK) Then
                ' "シリアルポートＯＰＥＮエラー"
                strMSG = MSG_136 + "(" + stCOM.PortName + ")"
                Call MsgBox(strMSG, vbOKOnly)
                Return (r)
            End If

            ' FL側からエラー情報を受信する
            r = RsReceiveErrInfo(ErrInf, cTIMEOUT)
            If (r <> SerialErrorCode.rRS_OK) Then
                ' "シリアルポート受信エラー"
                strMSG = MSG_139 + "(" + stCOM.PortName + ")"
                Call MsgBox(strMSG, vbOKOnly)
            End If

            ' ポートクローズ
            Rs232c_Close()

            ' INtime側にFLのエラー情報を送信する(ログ出力用)
            Call SET_FL_ERRLOG(ErrInf)

            Return (r)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Rs232c.ReceiveErrInfo() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (SerialErrorCode.rRS_Trap)
        End Try
    End Function
#End Region

#Region "RS232Cポートのオープン"
    '''=========================================================================
    '''<summary>RS232Cポートのオープン</summary>
    '''<param name="pstCom">(INP) ポート情報</param>
    '''<returns>0=正常, 0以外=エラー</returns> 
    '''=========================================================================
    Public Function Rs232c_Open(ByVal pstCom As ComInfo) As Integer

        Dim r As Integer
        Dim strMSG As String

        Try
            ' 初期処理
            r = SerialErrorCode.rRS_OK                          ' Return値 = 正常
            ObjPortInfo = New DllSerialIO.PortInformation       ' ﾎﾟｰﾄ情報ｵﾌﾞｼﾞｪｸﾄ生成
            ObjFLCom = New DllFLCom.FLComIO                     ' トリマー加工条件送受信ｵﾌﾞｼﾞｪｸﾄ生成

            ' ポート情報を設定する
            ObjPortInfo.PortName = pstCom.PortName              ' ポート番号
            ObjPortInfo.BaudRate = pstCom.BaudRate              ' Speed
            ObjPortInfo.Parity = pstCom.Parity                  ' Parity(0:None, 1:Odd, 2:Even, 3:Mark, 4:Space)
            ObjPortInfo.DataBits = pstCom.DataBits              ' Char Data
            ObjPortInfo.StopBits = pstCom.StopBits              ' Stop Bit

            ' ポートオープン
            r = ObjFLCom.Serial_Open(ObjPortInfo)               ' ポートオープン
            Rs_Flag = 1                                         ' ﾌﾗｸﾞ=1(ｵｰﾌﾟﾝ済)
            Return (r)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Rs232c.Rs232c_Open() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (SerialErrorCode.rRS_Trap)
        End Try
    End Function
#End Region

#Region "RS232Cポートのクローズ"
    '''=========================================================================
    '''<summary>RS232Cポートのクローズ</summary>
    '''<returns>0=正常, 0以外=エラー</returns> 
    '''=========================================================================
    Public Function Rs232c_Close() As Integer

        Dim r As Integer
        Dim strMSG As String

        Try
            ' ポートクローズ
            r = SerialErrorCode.rRS_OK                          ' Return値 = 正常
            If (Rs_Flag = 0) Then Exit Function ' ﾌﾗｸﾞ=0(未ｵｰﾌﾟﾝ)ならNOP
            r = ObjFLCom.Serial_Close()                         ' ポートクローズ
            Rs_Flag = 0                                         ' ﾌﾗｸﾞ=0(未ｵｰﾌﾟﾝ)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Rs232c.Rs232c_Close() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            r = SerialErrorCode.rRS_Trap
        End Try

        ObjPortInfo = Nothing                                   ' ﾎﾟｰﾄ情報ｵﾌﾞｼﾞｪｸﾄ解放
        ObjFLCom = Nothing                                      ' トリマー加工条件送受信ｵﾌﾞｼﾞｪｸﾄ解放
        Return (r)

    End Function
#End Region

#Region "シリアルポートへトリマー加工条件を個別に送信する"
    '''=========================================================================
    '''<summary>シリアルポートへトリマー加工条件を個別に送信する</summary>
    '''<param name="CndNum">(INP) 条件番号</param>
    '''<param name="Curr">  (INP) 電流値</param>
    '''<param name="Freq">  (INP) 周波数</param> 
    '''<param name="Steg">  (INP) STEG波形</param>  
    '''<returns>0=正常, 0以外=エラー</returns> 
    '''=========================================================================
    Public Function RsSendBankInfo(ByVal CndNum As Integer, ByVal Curr As Integer, ByVal Freq As Double, ByVal Steg As Integer) As Integer

        Dim wkFreq As Integer
        Dim r As Integer
        Dim dblWK As Double
        Dim strMSG As String

        Try
            ' 初期処理
            If (Rs_Flag = 0) Then                               ' ﾌﾗｸﾞ=0(未ｵｰﾌﾟﾝ) ?
                Return (SerialErrorCode.rRS_NotOpen)            ' Return値 = ｼﾘｱﾙﾎﾟｰﾄがｵｰﾌﾟﾝしていない
            End If

            ' 周波数(KHz)を繰返し間隔に変換する
            ' 繰返し間隔N = 繰返し時間/200ns 例)10KHzの場合は、500(0x01f4)を送信する。
            If (Freq <= 0) Then
                wkFreq = 0
            Else
                dblWK = 1000000 / 200 / Freq
                wkFreq = dblWK
            End If

            ' トリマー加工条件送信
            r = ObjFLCom.Serial_SendBankInfo(CndNum, Curr, wkFreq, Steg)
            Return (r)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Rs232c.RsSendBankInfo() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (SerialErrorCode.rRS_Trap)
        End Try
    End Function
#End Region

#Region "シリアルポートへトリマー加工条件を一括で送信する"
    '''=========================================================================
    '''<summary>シリアルポートへトリマー加工条件を一括で送信する</summary>
    '''<param name="Curr">  (INP) 電流値</param>
    '''<param name="Freq">  (INP) 周波数</param> 
    '''<param name="Steg">  (INP) STEG波形</param>  
    '''<returns>0=正常, 0以外=エラー</returns> 
    '''=========================================================================
    Public Function RsSendBankALL(ByVal Curr() As Integer, ByVal Freq() As Double, ByVal Steg() As Integer) As Integer

        Dim i As Integer
        Dim r As Integer
        Dim strMSG As String

        Try
            ' 初期処理
            If (Rs_Flag = 0) Then                               ' ﾌﾗｸﾞ=0(未ｵｰﾌﾟﾝ) ?
                Return (SerialErrorCode.rRS_NotOpen)            ' Return値 = ｼﾘｱﾙﾎﾟｰﾄがｵｰﾌﾟﾝしていない
            End If

            ' シリアルポートへトリマー加工条件を送信する
            For i = 0 To (MAX_BANK_NUM - 1)                     ' 最大加工条件数分(0-31)送信する
                r = RsSendBankInfo(i, Curr(i), Freq(i), Steg(i))
                If (r <> SerialErrorCode.rRS_OK) Then           ' 送信エラー ?
                    Exit For
                End If
            Next i
            Return (r)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Rs232c.RsSendBankALL() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (SerialErrorCode.rRS_Trap)
        End Try
    End Function
#End Region

#Region "シリアルポートからトリマー加工条件を個別に受信する"
    '''=========================================================================
    '''<summary>シリアルポートからトリマー加工条件を個別に受信する</summary>
    '''<param name="CndNum"> (INP) 条件番号</param>
    '''<param name="Curr">   (OUT) 電流値</param>
    '''<param name="Freq">   (OUT) 周波数</param> 
    '''<param name="Steg">   (OUT) STEG波形</param>  
    '''<param name="TimeOut">(INP) 応答待タイマ値(ms)</param>  
    '''<returns>0=正常, 0以外=エラー</returns> 
    '''=========================================================================
    Public Function RsReceiveBankInfo(ByVal CndNum As Integer, ByRef Curr As Integer, ByRef Freq As Double, ByRef Steg As Integer, ByVal TimeOut As Integer) As Integer

        Dim wkFreq As Integer
        Dim r As Integer
        Dim dblWK As Double
        Dim dblWK2 As Double
        Dim strMSG As String

        Try
            ' 初期処理
            If (Rs_Flag = 0) Then                               ' ﾌﾗｸﾞ=0(未ｵｰﾌﾟﾝ) ?
                Return (SerialErrorCode.rRS_NotOpen)            ' Return値 = ｼﾘｱﾙﾎﾟｰﾄがｵｰﾌﾟﾝしていない
            End If

            ' トリマー加工条件受信
            r = ObjFLCom.Serial_ReceiveBankInfo(CndNum, Curr, wkFreq, Steg, TimeOut)
            If (r <> SerialErrorCode.rRS_OK) Then               ' 受信エラー ?
                Return (r)                                      ' Return値設定
            End If

            ' 繰返し間隔を周波数(KHz)に変換する
            If (wkFreq <= 0) Then
                Freq = 0
            Else
                ' Freq = 1000000 / (wkFreq * 200)
                dblWK = wkFreq * 200
                dblWK2 = 1000000
                Freq = dblWK2 / dblWK
                dblWK2 = Fix(10 * Freq)                         ' 小数点２位以下を切り捨てる
                Freq = dblWK2 / 10                              ' 
            End If
            Return (SerialErrorCode.rRS_OK)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Rs232c.RsReceiveBankInfo() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (SerialErrorCode.rRS_Trap)
        End Try
    End Function
#End Region

#Region "シリアルポートからトリマー加工条件を一括で受信する"
    '''=========================================================================
    '''<summary>シリアルポートからトリマー加工条件を一括で受信する</summary>
    '''<param name="Curr">   (OUT) 電流値</param>
    '''<param name="Freq">   (OUT) 周波数</param> 
    '''<param name="Steg">   (OUT) STEG波形</param> 
    '''<param name="TimeOut">(INP) 応答待タイマ値(ms)</param>  
    '''<returns>0=正常, 0以外=エラー</returns> 
    '''=========================================================================
    Public Function RsReceiveBankALL(ByRef Curr() As Integer, ByRef Freq() As Double, ByRef Steg() As Integer, ByVal TimeOut As Integer) As Integer

        Dim i As Integer
        Dim r As Integer
        Dim strMSG As String

        Try
            ' 初期処理
            If (Rs_Flag = 0) Then                               ' ﾌﾗｸﾞ=0(未ｵｰﾌﾟﾝ) ?
                Return (SerialErrorCode.rRS_NotOpen)            ' Return値 = ｼﾘｱﾙﾎﾟｰﾄがｵｰﾌﾟﾝしていない
            End If

            ' シリアルポートからトリマー加工条件を一括で受信する
            For i = 0 To (MAX_BANK_NUM - 1)                     ' 最大加工条件数分受信する
                r = RsReceiveBankInfo(i, Curr(i), Freq(i), Steg(i), TimeOut)
                If (r <> SerialErrorCode.rRS_OK) Then           ' エラー ?
                    Return (r)
                End If
            Next i

            ' FL側の設定があるか確認する
            For i = 0 To (MAX_BANK_NUM - 1)                     ' 最大加工条件数分受信する
                If (Curr(i) <> 0) Then                          ' 電流値設定0以外が見つかったらFL側の設定があると判断する
                    Return (SerialErrorCode.rRS_OK)             ' Return値 = 正常
                End If
            Next i
            Return (SerialErrorCode.rRS_FLCND_NONE)             ' Return値 = 加工条件の設定なし

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Rs232c.RsReceiveBankALL() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (SerialErrorCode.rRS_Trap)
        End Try
    End Function
#End Region

#Region "シリアルポートからエラー情報を受信する"
    '''=========================================================================
    '''<summary>シリアルポートからエラー情報を受信する</summary>
    '''<param name="ErrInf"> (OUT) エラー情報</param>
    '''<param name="TimeOut">(INP) 応答待タイマ値(ms)</param>  
    '''<returns>0=正常, 0以外=エラー</returns> 
    '''=========================================================================
    Public Function RsReceiveErrInfo(ByRef ErrInf As Integer, ByVal TimeOut As Integer) As Integer

        Dim r As Integer
        Dim strMSG As String

        Try
            ' 初期処理
            If (Rs_Flag = 0) Then                               ' ﾌﾗｸﾞ=0(未ｵｰﾌﾟﾝ) ?
                Return (SerialErrorCode.rRS_NotOpen)            ' Return値 = ｼﾘｱﾙﾎﾟｰﾄがｵｰﾌﾟﾝしていない
            End If

            ' エラー情報受信
            r = ObjFLCom.Serial_ReceiveErrInfo(ErrInf, TimeOut)
            If (r <> SerialErrorCode.rRS_OK) Then               ' 受信エラー ?
                Return (r)                                      ' Return値設定
            End If
            Return (SerialErrorCode.rRS_OK)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "Rs232c.RsReceiveErrInfo() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (SerialErrorCode.rRS_Trap)
        End Try
    End Function
#End Region

    '#Region "バージョンの取得"
    '    '''=========================================================================
    '    '''<summary>バージョンの取得</summary>
    '    '''<param name="strVER">(OUT) strVER (0) = DllSerialIO.dllのバージョン
    '    '''                           strVER (1) = DllCndXMLIO.dllのバージョン
    '    '''                           strVER (2) = DllFLCom.dllのバージョン</param>
    '    '''=========================================================================
    '    Public Sub Rs232c_GetVersion(ByRef strVER() As String)

    '        Dim iMajor As Integer                                   ' Major Version
    '        Dim iMinor As Integer                                   ' Minor Version
    '        Dim iBNum As Integer                                    ' Build Number
    '        Dim iRev As Integer                                     ' Revision
    '        Dim strMSG As String

    '        Try
    '            ' バージョン情報クラスオブジェクト生成
    '            ObjPortCVer = New DllSerialIO.VersionInformation
    '            ObjFLXMLVer = New DllCndXMLIO.VersionInformation
    '            ObjFLComVer = New DllFLCom.VersionInformation

    '            ' バージョンの取得("Vx.x.x.x"の形式で返す)
    '            Call ObjPortCVer.GetVersion(iMajor, iMinor, iBNum, iRev)
    '            strVER(0) = "V" + Format(iMajor, "0") + "." + Format(iMinor, "0") + "." + Format(iBNum, "0") + "." + Format(iRev, "0")

    '            Call ObjFLXMLVer.GetVersion(iMajor, iMinor, iBNum, iRev)
    '            strVER(1) = "V" + Format(iMajor, "0") + "." + Format(iMinor, "0") + "." + Format(iBNum, "0") + "." + Format(iRev, "0")

    '            Call ObjFLComVer.GetVersion(iMajor, iMinor, iBNum, iRev)
    '            strVER(2) = "V" + Format(iMajor, "0") + "." + Format(iMinor, "0") + "." + Format(iBNum, "0") + "." + Format(iRev, "0")

    '            ' バージョン情報クラスオブジェクト開放
    '            ObjPortCVer = Nothing
    '            ObjFLXMLVer = Nothing
    '            ObjFLComVer = Nothing

    '            ' トラップエラー発生時 
    '        Catch ex As Exception
    '            strMSG = "Rs232c.Rs232c_GetVersion() TRAP ERROR = " + ex.Message
    '            MsgBox(strMSG)
    '        End Try
    '    End Sub
    '#End Region

#End Region
End Module
#End If                 ' cOSCILLATORcFLcUSE
