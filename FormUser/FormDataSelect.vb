'===============================================================================
'   Description  : データ選択画面処理(自動運転用)
'
'   Copyright(C) : TOWA LASERFRONT CORP. 2011
'
'===============================================================================
Imports System.IO

Public Class FormDataSelect
#Region "【変数定義】"
    '===========================================================================
    '   変数定義
    '===========================================================================
    Public giTrimNGMode As Integer                      ' トリミング不良信号を連続自動運転中止通知に使用(0:標準(機能なし), 1:機能あり)　tky.ini "DEVICE_CONST", "AUTO_OPERATION_TRM_NG"

    Private Const DATA_DIR_PATH As String = "C:\TRIMDATA\DATA"          ' データファイルフォルダ(既定値)
    'Private Const DATA_ENTRY_PATH As String = DATA_DIR_PATH & "\ENTRYLOT"   ' 登録済みﾃﾞｰﾀﾌｧｲﾙﾌｫﾙﾀﾞ
    Private Const ENTRY_PATH As String = "C:\TRIMDATA\ENTRYLOT\"
    Private Const ENTRY_TMP_FILE As String = "SAVE_ENTRY.TMP"


    '----- 連続運転用(SL436R用) -----
    Public gbFgAutoOperation As Boolean                     ' 自動運転フラグ(True:自動運転中, False:自動運転中でない) 
    Public gsAutoDataFileFullPath() As String               ' 連続運転登録データファイル名配列
    Public giAutoDataFileNum As Short                       ' 連続運転登録データファイル数
    'Public Const MODE_MAGAZINE As Short = 0                 ' マガジンモード
    'Public Const MODE_LOT As Short = 1                      ' ロットモード
    'Public Const MODE_ENDLESS As Short = 2                  ' エンドレスモード

    '----- 変数定義 -----
    Private mExitFlag As Integer                            ' 終了フラグ
    Private m_mainEdit As Form1                             ' ﾒｲﾝ画面への参照
    '    Private gsAutoDataFileFullPath() As String         ' 連続運転登録ﾘｽﾄﾌﾙﾊﾟｽ文字列配列
    Private sLogFileName As String
    Private sPlateDataFileName As String
    'Private sSaveL7 As String, sSaveL4 As String, sSaveL14 As String, sSaveL21 As String, sSaveL22 As String, sSaveL17 As String, sSaveL5 As String, sSaveL15 As String, sSaveL23 As String, sSaveL6 As String, sSaveL29 As String, sSaveL30 As String, sSaveL31 As String, sSaveL32 As String, sSaveL33 As String
    Private InitiallNgCount As Long
    Private NowlNgCount As Long
    Private bLotChange As Boolean
    Private NowExecuteLotNo As Integer
    Private AutoOpeCancel As Boolean
    Private CancelReason As Integer
    Private Const NO_MORE_ENTRY As Integer = 1
#End Region

#Region "ｺﾝｽﾄﾗｸﾀ"
    Friend Sub New(ByRef mainEdit As Form1)

        ' この呼び出しは、Windows フォーム デザイナで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        m_mainEdit = mainEdit ' ﾒｲﾝ画面への参照を設定

    End Sub
#End Region

#Region "【メソッド定義】"
#Region "終了結果を返す"
    '''=========================================================================
    ''' <summary>終了結果を返す</summary>
    ''' <returns>cFRS_ERR_START = OKボタン押下
    '''          cFRS_ERR_RST   = Cancelボタン押下</returns>
    '''=========================================================================
    Public Function sGetReturn() As Integer

        Dim strMSG As String

        Try
            Return (mExitFlag)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "FormDataSelect.sGetReturn() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Function
#End Region

#Region "ShowDialogメソッドに独自の引数を追加する"
    '''=========================================================================
    ''' <summary>ShowDialogメソッドに独自の引数を追加する</summary>
    ''' <param name="Owner">(INP)未使用</param>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Overloads Sub ShowDialog(ByVal Owner As IWin32Window)

        Dim strMSG As String

        Try
            ' 初期処理
            mExitFlag = -1                                              ' 終了フラグ = 初期化

            ' 画面表示
            Me.ShowDialog()                                             ' 画面表示
            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "FormDataSelect.ShowDialog() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "Form_Load時処理"
    '''=========================================================================
    ''' <summary>Form_Load時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub FormDataSelect_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim strMSG As String

        Try
            ' フォーム名 
            Me.Text = MSG_AUTO_14                                       ' "データ登録"

            ' ラベル名・ボタン名 
            LblDataFile.Text = MSG_AUTO_05                              ' "データファイル"
            LblListList.Text = MSG_AUTO_06                              ' "登録済みデータファイル"
            BtnUp.Text = MSG_AUTO_07                                    ' "リストの1つ上へ"
            BtnDown.Text = MSG_AUTO_08                                  ' "リストの1つ下へ"
            BtnDelete.Text = MSG_AUTO_09                                ' "リストから削除"
            BtnClear.Text = MSG_AUTO_10                                 ' "リストをクリア"
            BtnSelect.Text = MSG_AUTO_11                                ' "↓登録↓"
            BtnOK.Text = MSG_AUTO_12                                    ' "OK"
            BtnCancel.Text = MSG_AUTO_13                                ' "キャンセル"

            ' リストボックスクリア
            Call ListList.Items.Clear()                                 ' 「登録済みデータファイル」リストボックスクリア

            ' 「データファイル」リストボックスに日付付きファイル名を表示する
            DrvListBox.Drive = "C:"                                    ' ドライブ 
            DirListBox.Path = DATA_DIR_PATH                             ' ディレクトリリストボックス既定値
            MakeFileList()                                              ' ←通常はDirListBox_Change()イベントが発生するので不要だが
            '                                                           ' カレントが"C:\TRIMDATA\DATA"だと発生しないので必要

            ' 登録済みﾃﾞｰﾀﾌｧｲﾙﾌｫﾙﾀﾞの有無を確認する
            If (False = Directory.Exists(ENTRY_PATH)) Then
                Directory.CreateDirectory(ENTRY_PATH)              ' ﾌｫﾙﾀﾞが存在しなければ作成する
            End If

            Call LoadPlateDataFileFullPath()

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "FormDataSelect.FormDataSelect_Load() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

    '========================================================================================
    '   ボタン押下時の処理
    '========================================================================================
#Region "ﾃﾞｰﾀ設定ﾎﾞﾀﾝ・編集ﾎﾞﾀﾝ押下時処理"
    ''' <summary>ﾃﾞｰﾀ設定ﾎﾞﾀﾝ</summary>
    Private Sub cmdLotInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLotInfo.Click
        Call LoadAndEditData(0)

    End Sub

    ''' <summary>編集ﾎﾞﾀﾝ</summary>
    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        ' パスワード入力

        If (Func_Password(F_EDIT) <> True) Then         ' パスワード入力ｴﾗｰならEXIT
            Return
        End If

        Call LoadAndEditData(1)

            MakeFileList()              ' DirListBoxで選択されているﾌｫﾙﾀﾞに保存された場合、ﾘｽﾄの更新が必要   'V4.7.0.0⑬
    End Sub

    ''' <summary>選択中の登録済みﾃﾞｰﾀﾌｧｲﾙを読み込んでﾃﾞｰﾀ設定画面または編集画面を開く</summary>
    ''' <param name="button">0=ﾃﾞｰﾀ設定ﾎﾞﾀﾝ,1=編集ﾎﾞﾀﾝ</param>
    Private Sub LoadAndEditData(ByVal button As Integer)
        Dim rslt As Short
        Dim s As String
        Dim r As Short

        ' 登録済みのﾃﾞｰﾀﾌｧｲﾙがなければNOP
        If (ListList.Items.Count < 1) Then Exit Sub
        Try
            '-----------------------------------------------------------------------
            '   初期処理
            '-----------------------------------------------------------------------
            giAppMode = APP_MODE_LOAD                       ' ｱﾌﾟﾘﾓｰﾄﾞ = ファイルロード(F1)

            ' パスワード入力(オプション)
            rslt = Func_Password(F_LOAD)
            If (rslt <> True) Then
                Exit Try                                    ' ﾊﾟｽﾜｰﾄﾞ入力ｴﾗｰならEXIT
            End If

            ' ﾃﾞｰﾀﾌｧｲﾙ名設定
            With ListList
                gsDataFileName = (ENTRY_PATH & .Items(.SelectedIndex))
            End With

            ' 旧設定の装置の電圧をOFFする
            r = V_Off()                                     ' DC電源装置 電圧OFF処理

            ' トリミングデータ設定
            r = UserVal()                                   ' データ初期設定
            If (r <> 0) Then                                ' エラー ?
                pbLoadFlg = False                           ' データロード済フラグ = False
                s = "Data load Error : " & gsDataFileName & vbCrLf
                Me.LblFullPath.Text = s
                Call Z_PRINT(s)
            Else
                Call Z_CLS()                                ' データロードでログ画面クリア              ###lstLog
                gDspCounter = 0                             ' ログ画面表示基板枚数カウンタクリア
                pbLoadFlg = True                            ' データロード済フラグ = True
                s = "Data loaded : " & gsDataFileName & vbCrLf
                Call Z_PRINT(s)

                Call m_mainEdit.System1.OperationLogging( _
                        gSysPrm, MSG_OPLOG_FUNC01, "File='" & gsDataFileName & "' MANUAL")

                'V2.1.0.0④↓
                If gsDataFileName.Length > 60 Then
                    m_mainEdit.LblDataFileName.Text = gsDataFileName
                Else
                    'V2.1.0.0④↑
                    ' ファイルパス名の表示
                    If (gSysPrm.stTMN.giMsgTyp = 0) Then
                        m_mainEdit.LblDataFileName.Text = "データファイル名 " & gsDataFileName
                    Else
                        m_mainEdit.LblDataFileName.Text = "File name " & gsDataFileName
                    End If
                End If          'V2.1.0.0④
                '-----------------------------------------------------------------------
                '   FL側へ加工条件を送信する(FL時で加工条件ファイルがある場合)
                '-----------------------------------------------------------------------
                Call m_mainEdit.SendFlParam(gsDataFileName)

                '###1040⑥                Call m_mainEdit.SetATTRateToScreen(True)    ' ###1040③ アッテネータの設定
            End If

            '-----------------------------------------------------------------------
            '   ﾛｰﾄﾞ終了処理
            '-----------------------------------------------------------------------
            ChDrive("C")                                    ' ChDriveしないと次起動時FDドライブを見に行って,
            ChDir(My.Application.Info.DirectoryPath)        ' "MVCutil.dllがない"となり起動できなくなる

            ' ======================================================================
            '   ﾃﾞｰﾀ設定画面・編集画面呼び出し
            ' ======================================================================
            ' ﾃﾞｰﾀﾛｰﾄﾞﾁｪｯｸ (ﾄﾘﾐﾝｸﾞﾃﾞｰﾀ初期設定:UserVal() のｴﾗｰﾁｪｯｸ)
            If pbLoadFlg = False Then
                s = MSG_DataNotLoad                         ' ﾃﾞｰﾀ未ﾛｰﾄﾞ
                Call Z_PRINT(s)
                Call Beep()
                Exit Try
            End If

            If (0 = button) Then
                ' ﾃﾞｰﾀ設定画面
                giAppMode = APP_MODE_LOTNO                  ' ｱﾌﾟﾘﾓｰﾄﾞ = ロット番号設定中
                ' データ編集
                Call m_mainEdit.System1.OperationLogging(gSysPrm, MSG_OPLOG_LOTSET, "")

                Dim fLotInf As New FormEdit.frmLotInfoInput()
                fLotInf.ShowDialog(Me)
                fLotInf.Dispose()
            Else
                ' 編集画面
                giAppMode = APP_MODE_EDIT                   ' ｱﾌﾟﾘﾓｰﾄﾞ = 編集画面表示
                Call m_mainEdit.System1.OperationLogging(gSysPrm, MSG_OPLOG_FUNC03, "")

                FlgUpdGPIB = 0                              ' GPIBデータ更新Flag Off
                Dim fForm As New FormEdit.frmEdit           ' frmｵﾌﾞｼﾞｪｸﾄ生成
                fForm.ShowDialog()                          ' データ編集
                fForm.Dispose()                             ' frmｵﾌﾞｼﾞｪｸﾄ開放

                ' GPIBデータ更新ならGPIB初期化を行う
                If (FlgUpdGPIB = 1) Then
                    Call GPIB_Init()
                End If
            End If

            If (True = FlgUpd) Then
                '-----------------------------------------------------------------------
                '   データファイルをセーブする
                '-----------------------------------------------------------------------
                If rData_save(gsDataFileName) <> 0 Then       ' データファイルセーブ
                    Exit Try
                Else
                    Call Z_PRINT("Data saved : " & gsDataFileName & vbCrLf)
                End If

                '-----------------------------------------------------------------------
                '   操作ログ等を出力する
                '-----------------------------------------------------------------------
                Call m_mainEdit.System1.OperationLogging( _
                    gSysPrm, MSG_OPLOG_FUNC02, "File='" & gsDataFileName & "' MANUAL")

                FlgUpd = Convert.ToInt16(TriState.False)    ' データ更新 Flag OFF
            End If

            ChDrive("C")                                    ' ChDriveしないと次起動時FDドライブを見に行って,"MVCutil.dllがない"となり起動できなくなる
            ChDir(My.Application.Info.DirectoryPath)

            ' トラップエラー発生時
        Catch ex As Exception
            MsgBox("LoadAndEditData() TRAP ERROR = " + ex.Message)
        Finally
            Call ZCONRST()                                  ' ｺﾝｿｰﾙｷｰ ﾗｯﾁ解除
            giAppMode = APP_MODE_LOTCHG                     ' ｱﾌﾟﾘﾓｰﾄﾞ = ロット切替
        End Try

    End Sub
#End Region

#Region "OKボタン押下時処理"
    '''=========================================================================
    ''' <summary>OKボタン押下時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOK.Click

        Dim Idx As Integer
        Dim strMSG As String = ""
        Dim sRtn As Short

        Try
            gbFgAutoOperation = False

            ' 選択リスト1以上有りかチェックする ?
            If (ListList.Items.Count < 1) Then
                '"データファイルを選択してください。"
                Call MsgBox(MSG_AUTO_18, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

#If cOSCILLATORcFLcUSE Then
        Dim r As Integer
        Dim strDAT As String
            ' 選択データに対応する加工条件ファイルが存在するかチェックする(FL時)
            If (gSysPrm.stRAT.giOsc_Res = OSCILLATOR_FL) Then
                For Idx = 0 To ListList.Items.Count - 1
                    strDAT = (ENTRY_PATH & ListList.Items(Idx))
                    r = GetFLCndFileName(strDAT, strMSG, True)              ' 存在チェック 
                    If (r <> SerialErrorCode.rRS_OK) Then                   ' 加工条件ファイルが存在しない ?
                        ' "加工条件ファイルが存在しません。(加工条件ファイル名)"
                        strMSG = MSG_AUTO_20 + "(" + strMSG + ")"
                        Call MsgBox(strMSG, MsgBoxStyle.OkOnly, "")
                        ListList.SelectedIndex = Idx
                        Call ListList_SelectedIndexChanged(sender, e)                      ' データファイル名をフルパスでラベルテキストボックスに設定する
                        Exit Sub
                    End If
                Next Idx
            End If
#End If

            ' 連続運転用のデータファイル数とデータファイル名配列をグローバル領域に設定する
            giAutoDataFileNum = ListList.Items.Count                    ' データファイル数
            ReDim gsAutoDataFileFullPath(giAutoDataFileNum - 1)
            For Idx = 0 To giAutoDataFileNum - 1                        ' データファイル名
                gsAutoDataFileFullPath(Idx) = (ENTRY_PATH & ListList.Items(Idx))
            Next

            If OffSetCheckBox.Checked = True Then
                sRtn = UserSub.SetOffSetDataToAutoOperationData(gsAutoDataFileFullPath, giAutoDataFileNum)
                If sRtn <> cFRS_NORMAL Then
                    Call Form1.System1.TrmMsgBox(gSysPrm, "オフセットパラメータ自動反映処理" & vbCrLf & "ファイル[" & gsAutoDataFileFullPath(sRtn - 1) & "]の処理でエラーが発生しました。", MsgBoxStyle.OkOnly, My.Application.Info.Title)
                    Return
                End If
            End If

            If Not Form1.TrimDataLoad(gsAutoDataFileFullPath(0)) Then
                'V2.1.0.0④                Call Z_PRINT("自動運転時トリミングデータファイルＬＯＡＤエラー = " & vbCrLf)
                Form1.AutoRunnningDisp.BackColor = System.Drawing.Color.Yellow  ' 背景色 = 黄色 'V2.0.0.2③
                Form1.AutoRunnningDisp.Text = "自動運転解除中"                                  'V2.0.0.2③
                'V2.1.0.0④↓
                Call Z_PRINT("自動運転時トリミングデータファイルＬＯＡＤエラー" & vbCrLf & "= [" & gsAutoDataFileFullPath(0) & "]")
                Call Form1.System1.TrmMsgBox(gSysPrm, "データファイルＬＯＡＤエラー" & vbCrLf & "ファイル[" & gsAutoDataFileFullPath(0) & "]の処理でエラーが発生しました。", MsgBoxStyle.OkOnly, My.Application.Info.Title)
                Return
                'V2.1.0.0④↑
            Else
                Call InitialAutoOperation()
                gbFgAutoOperation = True
                Form1.AutoRunnningDisp.BackColor = System.Drawing.Color.Lime ' 背景色 = 緑
                Form1.AutoRunnningDisp.Text = "自動運転中"
                UserSub.LaserCalibrationSet(POWER_CHECK_START)          'V2.1.0.0② レーザパワーモニタリング実行有無設定
            End If

            mExitFlag = cFRS_ERR_START                                  ' Return値 = OKボタン押下 

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "FormDataSelect.BtnOK_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
        Me.Close()                                                      ' フォームを閉じる
    End Sub
#End Region

#Region "Cancelボタン押下時処理"
    '''=========================================================================
    ''' <summary>Cancelボタン押下時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click

        Dim strMSG As String

        Try
            mExitFlag = cFRS_ERR_RST                                    ' Return値 = Cancelボタン押下

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "FormDataSelect.BtnCancel_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
        Me.Close()                                                      ' フォームを閉じる
    End Sub
#End Region

#Region "「リストの１つ上へ」ボタン押下時処理"
    '''=========================================================================
    ''' <summary>「リストの１つ上へ」ボタン押下時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnUp.Click

        Dim Idx As Integer
        Dim strMSG As String

        Try
            Idx = ListList.SelectedIndex
            ' 先頭が選択されている場合NOP
            If (Idx <= 0) Then Exit Sub
            Call SwapList(Idx, (Idx - 1))       ' ﾘｽﾄを入れ替え
            ListList.SelectedIndex = (Idx - 1)  ' １つ上のｲﾝﾃﾞｯｸｽを指定する

            ' データファイル名をフルパスでラベルテキストボックスに設定する
            Call ListList_SelectedIndexChanged(sender, e)
            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "FormDataSelect.BtnUp_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "「リストの１つ下へ」ボタン押下時処理"
    '''=========================================================================
    ''' <summary>「リストの１つ下へ」ボタン押下時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDown.Click

        Dim Idx As Integer
        Dim strMSG As String

        Try
            Idx = ListList.SelectedIndex
            ' 最後が選択されている場合NOP
            If ((Idx + 1) >= ListList.Items.Count) Then Exit Sub
            Call SwapList(Idx, (Idx + 1))       ' ﾘｽﾄを入れ替え
            ListList.SelectedIndex = (Idx + 1)  ' １つ下のｲﾝﾃﾞｯｸｽを指定する

            ' データファイル名をフルパスでラベルテキストボックスに設定する
            Call ListList_SelectedIndexChanged(sender, e)
            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "FormDataSelect.BtnDown_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "登録済みﾃﾞｰﾀﾌｧｲﾙの項目を入れ替える"
    ''' <summary>登録済みﾃﾞｰﾀﾌｧｲﾙの項目を入れ替える</summary>
    ''' <param name="iSrc">元位置</param>
    ''' <param name="iDst">移動先位置</param>
    Private Sub SwapList(ByVal iSrc As Integer, ByVal iDst As Integer)
        Dim tmpStr As String
        tmpStr = ListList.Items(iSrc)
        ListList.Items.RemoveAt(iSrc)
        ListList.Items.Insert(iDst, tmpStr)

    End Sub
#End Region

#Region "「リストから削除」ボタン押下時処理"
    '''=========================================================================
    ''' <summary>「リストから削除」ボタン押下時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDelete.Click

        Dim Idx As Integer
        Dim strMSG As String

        Try
            ' 「登録済みデータファイル」リストボックスから1つ削除する
            Idx = ListList.SelectedIndex
            If (Idx < 0) Then Exit Sub
            File.Delete(ENTRY_PATH & ListList.Items(Idx))    ' 選択されているﾌｧｲﾙを削除する
            ListList.Items.RemoveAt(Idx)                                '「登録済みデータファイル」リストボックスから1項目削除(※Remove()は文字列指定)

            ' データファイル名をフルパスでラベルテキストボックスに設定する(削除の１つ前のデータを選択状態とする)
            If (0 <= Idx) Then
                Idx = (Idx - 1)
                ' ﾘｽﾄの先頭が削除された場合に他のﾃﾞｰﾀがあれば選択する
                If (Idx < 0) AndAlso (0 < ListList.Items.Count) Then Idx = 0
                ListList.SelectedIndex = Idx    ' ｲﾍﾞﾝﾄにより登録済みﾃﾞｰﾀの選択中ﾌｧｲﾙﾌﾙﾊﾟｽを再表示
            Else
                ' 登録済みﾃﾞｰﾀﾌｧｲﾙがなくなった場合
                Call ListList_SelectedIndexChanged(sender, e)  ' 登録済みﾃﾞｰﾀの選択中ﾌｧｲﾙﾌﾙﾊﾟｽを再表示
            End If

            Call DirListBox_Change(sender, e)   ' ﾃﾞｨﾚｸﾄﾘﾂﾘｰを再表示

            ' エンドレスモード処理
            Call DspEndless()

            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "FormDataSelect.BtnDelete_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "「リストをクリア」ボタン押下時処理"
    '''=========================================================================
    ''' <summary>「リストをクリア」ボタン押下時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>「登録済みデータファイル」リストボックスから全て削除</remarks>
    '''=========================================================================
    Private Sub BtnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClear.Click

        Dim r As Integer
        Dim strMSG As String
        Try
            If (ListList.Items.Count < 1) Then
                ' 登録済みﾃﾞｰﾀﾌｧｲﾙﾘｽﾄに項目がない場合、ENTRYLOT ﾌｫﾙﾀﾞ内のﾌｧｲﾙをすべて削除する
                For Each tmpFile As String In (Directory.GetFiles(ENTRY_PATH))
                    File.Delete(tmpFile)
                Next
                Exit Sub
            Else
                ' 登録済みﾃﾞｰﾀﾌｧｲﾙﾘｽﾄに項目がある場合、削除確認ﾒｯｾｰｼﾞを表示する
                ' "登録リストを全て削除します。" & vbCrLf & "よろしいですか？"
                strMSG = MSG_AUTO_15 & vbCrLf & MSG_AUTO_16
                r = MsgBox(strMSG, MsgBoxStyle.OkCancel, "")
                If (r <> MsgBoxResult.Ok) Then Exit Sub ' ｷｬﾝｾﾙ

                ' ENTRYLOT ﾌｫﾙﾀﾞ内のﾌｧｲﾙをすべて削除する
                For Each tmpFile As String In (Directory.GetFiles(ENTRY_PATH))
                    File.Delete(tmpFile)
                Next

                ' 「登録済みデータファイル」リストボックスクリア
                Call ListList.Items.Clear() '「登録済みデータファイル」リストボックスクリア

                ' データファイル名をフルパスでラベルテキストボックスに設定する(クリアする)
                Call ListList_SelectedIndexChanged(sender, e)
                Call DirListBox_Change(sender, e) ' ﾃﾞｨﾚｸﾄﾘﾂﾘｰを再表示する

                ' エンドレスモード処理
                Call DspEndless()

                Exit Sub
            End If

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "FormDataSelect.BtnClear_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "登録ボタン押下時処理"
    '''=========================================================================
    ''' <summary>登録ボタン押下時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSelect.Click

        Dim Idx As Integer
        Dim Sz As Integer
        Dim Pos As Integer
        Dim strDAT As String
        Dim strMSG As String
        Try
            '「データファイル」リストボックスインデックス無効ならNOP
            Idx = ListFile.SelectedIndex
            If (Idx < 0) Then Exit Sub
            ' エンドレスモードで選択リスト1以上有りならNOP

            ' 指定のデータファイル名を「登録済みデータファイル」リストボックスに追加する
            strDAT = ListFile.Items(Idx)                                ' 日付時刻付きファイル名なのでファイル名のみ取り出す 
            Sz = strDAT.Length
            Pos = strDAT.LastIndexOf(" ")
            If (Pos = -1) Then Exit Sub
            strDAT = strDAT.Substring(Pos + 1, Sz - Pos - 1)
            Idx = ListList.Items.Count

            Dim sFromFilePath As String = ""
            Dim sCopyFilePath As String = ""
            If (False = CopyEntryFileToWorkFolder(sFromFilePath, sCopyFilePath, strDAT)) Then ' 選択ﾌｧｲﾙをｺﾋﾟｰする
                ' TODO: ｴﾗｰﾒｯｾｰｼﾞ
                MsgBox((sFromFilePath & vbCrLf & vbTab & "↓" & vbCrLf & sCopyFilePath), _
                    DirectCast((MsgBoxStyle.Critical + MsgBoxStyle.OkOnly), MsgBoxStyle))
            Else
                Call ListList.Items.Add(strDAT)
                ListList.SelectedIndex = Idx

                ' データファイル名をフルパスでラベルテキストボックスに設定する
                Call ListList_SelectedIndexChanged(sender, e)

            End If

            ' エンドレスモード処理
            Call DspEndless()
            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "FormDataSelect.BtnSelect_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "登録ﾌｧｲﾙをENTRYLOTﾌｫﾙﾀﾞにｺﾋﾟｰする"
    ''' <summary>登録ﾌｧｲﾙをENTRYLOTﾌｫﾙﾀﾞにｺﾋﾟｰする</summary>
    ''' <param name="sFromFilePath">IN="",OUT=ｺﾋﾟｰ元ﾌﾙﾊﾟｽ</param>
    ''' <param name="sCopyFilePath">IN="",OUT=ｺﾋﾟｰ先ﾌﾙﾊﾟｽ</param>
    ''' <param name="sCopyFile">IN=ｺﾋﾟｰするﾌｧｲﾙ名.拡張子,OUT=ｺﾋﾟｰしたﾌｧｲﾙ名.拡張子</param>
    ''' <returns>True=成功,False=失敗</returns>
    ''' <remarks>ﾌｧｲﾙ名に名に_01,_02と連番を付加する</remarks>
    Private Function CopyEntryFileToWorkFolder(ByRef sFromFilePath As String, _
                ByRef sCopyFilePath As String, ByRef sCopyFile As String) As Boolean
        Dim sTmpFile As String          ' ﾌｧｲﾙ名
        Dim sExtended As String         ' 拡張子

        CopyEntryFileToWorkFolder = False
        Try
            sTmpFile = sCopyFile.Split(".")(0)
            sExtended = "." & sCopyFile.Split(".")(1)

            For i As Integer = 0 To 99 Step 1
                ' 連番を追加したﾌｧｲﾙ名の作成
                sCopyFilePath = (ENTRY_PATH & sTmpFile & "_" & i.ToString("00") & sExtended)
                Debug.Print(sCopyFilePath)
                ' 同名ﾌｧｲﾙの存在確認
                If (False = File.Exists(sCopyFilePath)) Then
                    ' 存在しなければﾌｧｲﾙをｺﾋﾟｰ
                    sFromFilePath = (FileLstBox.Path & "\" & sTmpFile & sExtended)
                    File.Copy(sFromFilePath, sCopyFilePath)
                    If (File.Exists(sCopyFilePath)) Then
                        sCopyFile = (sTmpFile & "_" & i.ToString("00") & sExtended)
                        CopyEntryFileToWorkFolder = True
                    End If
                    Exit Function
                End If
            Next i

        Catch ex As Exception
            Dim strMSG As String = _
                "FormDataSelect.CopyEntryFileToWorkFolder() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try

    End Function
#End Region

    '========================================================================================
    '   リストボックスのクリックイベント処理
    '========================================================================================
#Region "「データファイル」リストボックスダブルクリックイベント処理"
    '''=========================================================================
    ''' <summary>「データファイル」リストボックスダブルクリックイベント処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub ListFile_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListFile.DoubleClick
        Dim strMSG As String

        Try
            ' 登録ボタン押下時処理へ
            Call BtnSelect_Click(sender, e)
            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "FormDataSelect.ListFile_DoubleClick() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "「登録済みデータファイル」ﾘｽﾄのｲﾝﾃﾞｯｸｽ変更ｲﾍﾞﾝﾄ"
    '''=========================================================================
    ''' <summary>「登録済みデータファイル」ﾘｽﾄのｲﾝﾃﾞｯｸｽ変更ｲﾍﾞﾝﾄ</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub ListList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                Handles ListList.SelectedIndexChanged
        Dim Idx As Integer
        Dim strMSG As String
        Try
            ' 「登録済みデータファイル」リストボックスで選択されたデータファイル名をフルパスでラベルテキストボックスに設定する
            Idx = ListList.SelectedIndex
            If (Idx < 0) Then
                LblFullPath.Text = ""
            Else
                LblFullPath.Text = (ENTRY_PATH & ListList.Items(Idx))
            End If
            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "FormDataSelect.ListList_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "「ドライブリストボックス」の SelectedIndexChanged 処理"
    '''=========================================================================
    ''' <summary>「ドライブリストボックス」の SelectedIndexChanged 処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>DriveListBoxは非標準コントロールなのでツールボックスに追加する必要有り</remarks>
    '''=========================================================================
    Private Sub DrvListBox_SelectedIndexChanged( _
        ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DrvListBox.SelectedIndexChanged

        Try
            ' ディレクトリリストボックスの選択ドライブを変更する
            Dim tmpDrv As String = DrvListBox.Drive
            If (0 = (String.Compare(tmpDrv, "C:", True))) Then tmpDrv = DATA_DIR_PATH
            DirListBox.Path = tmpDrv
            Call DirListBox_Change(sender, e)   ' ﾃﾞｨﾚｸﾄﾘﾂﾘｰを再表示する
            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            Dim strMSG As String = ex.Message
            MsgBox(strMSG)
            DrvListBox.Drive = "C:"
        End Try
    End Sub
#End Region

#Region "「ディレクトリリストボックス」の変更時処理"
    '''=========================================================================
    ''' <summary>「ディレクトリリストボックス」の変更時処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>DirListBoxは非標準コントロールなのでツールボックスに追加する必要有り</remarks>
    '''=========================================================================
    Private Sub DirListBox_Change(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DirListBox.Change

        Dim strMSG As String

        Try
            ' 選択ディレクトリを変更する(FileLstBoxは作業用のDummy)
            FileLstBox.Path = DirListBox.Path

            ' 「データファイル」リストボックスに日付時刻付きファイル名を表示する
            MakeFileList()
            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "FormDataSelect.DirListBox_Change() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub

    ''' <summary>ﾃﾞｨﾚｸﾄﾘﾎﾞｯｸｽｸﾘｯｸ時の処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>この処理により DirListBox_Change ｲﾍﾞﾝﾄが発生し、ﾃﾞｨﾚｸﾄﾘﾂﾘｰを再表示する</remarks>
    Private Sub DirListBox_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DirListBox.Click
        Try
            With DirectCast(sender, VB6.DirListBox)
                If (.Path <> .DirList(.DirListIndex)) Then
                    .Path = .DirList(.DirListIndex) ' 選択したﾃﾞｨﾚｸﾄﾘをﾊﾟｽに設定する
                End If
            End With
        Catch ex As Exception
            Dim strMSG As String = "FormDataSelect.DirListBox_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

    '========================================================================================
    '   動作モードラジオボタン変更時の処理
    '========================================================================================
#Region "マガジンモード選択処理"
    '''=========================================================================
    ''' <summary>マガジンモード選択処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnMdMagazine_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call DspEndless()
    End Sub
#End Region

#Region "ロットモード選択処理"
    '''=========================================================================
    ''' <summary>ロットモード選択処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnMdLot_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call DspEndless()
    End Sub
#End Region

#Region "エンドレスモード選択処理"
    '''=========================================================================
    ''' <summary>エンドレスモード選択処理</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnMdEndless_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call DspEndless()
    End Sub
#End Region

    '========================================================================================
    '   共通関数定義
    '========================================================================================
#Region "「データファイル」リストボックスに日付時刻付きファイル名を表示する"
    '''=========================================================================
    ''' <summary>「データファイル」リストボックスに日付時刻付きファイル名を表示する</summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub MakeFileList()

        Dim Count As Integer
        Dim i As Integer
        Dim Sz As Integer
        Dim strWK As String
        Dim strDAT As String
        Dim strMSG As String

        Try
            ' 「データファイル」リストボックスに日付時刻付きファイル名を表示する
            Call ListFile.Items.Clear()                                                 '「データファイル」リストボックスクリア
            FileLstBox.Refresh()                                                        ' ファイルリストを更新する  'V4.7.0.0⑬
            Count = FileLstBox.Items.Count                                              ' ファイルの数 
            For i = 0 To (Count - 1)
                ' ファイル拡張子を設定
                strWK = ".txt"

                ' 対象の拡張子でなければSKIP
                strDAT = FileLstBox.Items(i)
                Sz = strDAT.Length
                If (Sz < 4) Then GoTo STP_NEXT
                strDAT = strDAT.Substring(Sz - 4, 4)                                    ' 拡張子を取り出す
                If (String.Compare(strDAT, strWK, True)) Then GoTo STP_NEXT '           ' 対象の拡張子でなければSKIP(大文字、小文字を区別しない)

                ' 日付時刻付きファイルリスト作成
                Dim tmpFile As String = FileLstBox.Path & "\" & FileLstBox.Items(i)
                If (False = (File.Exists(tmpFile))) Then Continue For ' ﾌｧｲﾙの存在確認
                strDAT = FileDateTime(tmpFile)
                Dim Dt As DateTime = DateTime.Parse(strDAT)
                strDAT = Dt.ToString("yyyy/MM/dd HH:mm:ss") + " " + FileLstBox.Items(i) ' 日付時刻の長さを合わせる 
                Call ListFile.Items.Add(strDAT)                                         ' 日付時刻付きファイル名を表示する
STP_NEXT:
            Next i
            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "FormDataSelect.MakeFileList() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "エンドレスモード処理"
    '''=========================================================================
    ''' <summary>エンドレスモード処理</summary>
    ''' <remarks>エンドレスモード時はデータファイルを１つしか選択できない</remarks>
    '''=========================================================================
    Private Sub DspEndless()

        Dim strMSG As String

        Try
            ' エンドレスモードで選択リスト1以上有りなら下記のボタン等を非活性化にする
            'If (BtnMdEndless.Checked = True) And (ListList.Items.Count >= 1) Then
            '    ListFile.Enabled = False                                ' データファイルリストボックス非活性化
            '    BtnSelect.Enabled = False                               ' 登録ボタン非活性化 
            '    BtnUp.Enabled = False                                   '「リストの１つ上へ」ボタン非活性化
            '    BtnDown.Enabled = False                                 '「リストの１つ下へ」ボタン非活性化
            'Else
            ListFile.Enabled = True                                 ' データファイルリストボックス活性化
            BtnSelect.Enabled = True                                ' 登録ボタン活性化 
            BtnUp.Enabled = True                                    '「リストの１つ上へ」ボタン活性化
            BtnDown.Enabled = True                                  '「リストの１つ下へ」ボタン活性化
            'End If

            Exit Sub

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "FormDataSelect.DspEndless() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "自動運転用ファンクション"
    '=========================================================================
    '【機　能】プレートデータファイル名の設定
    '【引　数】0:設定 1:設定無し
    '【戻り値】連続自動運転のログファイル名を生成する。
    '=========================================================================
    Public Function PlateDataFileName(ByVal mode As Integer, ByVal sName As String) As String

        If mode = 0 Then
            sPlateDataFileName = sName
        End If

        PlateDataFileName = sPlateDataFileName

    End Function

    '=========================================================================
    '【機　能】プレートデータ・ファイルを削除する。
    '【引　数】無し
    '【戻り値】無し
    '=========================================================================
    Public Sub SavePlateDataFileDelete()
        Dim sFolder As String

        Try

            sFolder = ENTRY_PATH & ENTRY_TMP_FILE

            If IO.File.Exists(sFolder) = True Then  ' ファイルが有れば削除する。
                IO.File.Delete(sFolder)
            End If
        Catch ex As Exception
            Call Z_PRINT("FormDataSelect.SavePlateDataFileDelete() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Sub
    '=========================================================================
    '【機　能】途中で終了した時にプレートデータを保存する。
    '【引　数】プレートデータ配列、スタート(0 Origin)、終了
    '【戻り値】無し
    '=========================================================================
    Public Sub SavePlateDataFileFullPath(ByRef sPath() As String, ByVal iStart As Integer, ByVal iEnd As Integer)
        Dim sFolder As String
        Dim iFileNo As Integer
        Dim WS As IO.StreamWriter

        Try

            sFolder = ENTRY_PATH & ENTRY_TMP_FILE

            Call SavePlateDataFileDelete()

            WS = New IO.StreamWriter(sFolder, True, System.Text.Encoding.GetEncoding("Shift-JIS"))

            For iFileNo = iStart To iEnd
                WS.WriteLine(sPath(iFileNo))
            Next

            WS.Close()

        Catch ex As Exception
            Call Z_PRINT("FormDataSelect.SavePlateDataFileFullPath() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Sub
    '=========================================================================
    '【機　能】途中で終了した時に保存されたプレートデータをロードする。
    '【引　数】無し
    '【戻り値】無し
    '=========================================================================
    Private Sub LoadPlateDataFileFullPath() ' private V4.7.0.0⑬
        'Public Sub LoadPlateDataFileFullPath()
        Dim sFolder As String
        Dim sPathData As String = ""

        sFolder = ENTRY_PATH & ENTRY_TMP_FILE

        If IO.File.Exists(sFolder) = True Then  ' ファイルが読み取る。
            Using sr As New System.IO.StreamReader(sFolder, System.Text.Encoding.GetEncoding("Shift-JIS"))
                Do While Not sr.EndOfStream
                    Dim sPath As String = sr.ReadLine
                    If sPath <> "" Then
                        'V1.1.0.1②                        ListList.Items.Add(sPath)
                        ListList.Items.Add(IO.Path.GetFileName(sPath))           'V1.1.0.1②
                    Else
                        MsgBox("ファイルが存在しませんでした =" & sPathData, vbOKOnly Or vbExclamation Or vbSystemModal Or vbMsgBoxSetForeground, "Warning")
                    End If
                Loop
            End Using
        End If

    End Sub

    '=========================================================================
    '【機　能】連続自動運転用連続トリミングＮＧ枚数カウンター
    '【第１引数】0:初期化、1:NGカウント設定 その他：現在カウンターの取得
    '【第２引数】ＮＧカウンター値
    '【戻り値】ＮＧカウンター値
    '=========================================================================
    Private Function NGCountData(ByVal mode As Integer, ByVal lNgCount As Long) As Long

        If mode = 0 Then
            InitiallNgCount = lNgCount
            NowlNgCount = 0
        ElseIf mode = 1 Then
            NowlNgCount = lNgCount - InitiallNgCount
        End If

        NGCountData = NowlNgCount
        Debug.Print("InitiallNgCount=" & InitiallNgCount & "NowlNgCount=" & NowlNgCount)
    End Function
    '=========================================================================
    '【機　能】連続トリミングＮＧ枚数カウンター設定初期化
    '【引　数】ＮＧカウンター値
    '【戻り値】無し
    '=========================================================================
    Public Sub InitNGCountForContinueAuto(ByVal lNgCount As Long)
        Call NGCountData(0, lNgCount)
    End Sub

    '=========================================================================
    '【機　能】連続トリミングＮＧ枚数カウンター設定
    '【引　数】ＮＧカウンター値
    '【戻り値】無し
    '=========================================================================
    Public Sub SetNGCountForContinueAuto(ByVal lNgCount As Long)
        Call NGCountData(1, lNgCount)
    End Sub

    '=========================================================================
    '【機　能】ロット切り替え判定初期化
    '【引　数】無し
    '【戻り値】無し
    '=========================================================================
    Public Sub InitLotChangeJudge()
        bLotChange = False
    End Sub

    '=========================================================================
    '【機　能】ロット切り替え判定セット
    '【引　数】無し
    '【戻り値】無し
    '=========================================================================
    Public Sub SetLotChangeJudge()
        bLotChange = True
    End Sub
    '=========================================================================
    '【機　能】ロット切り替え判定取得
    '【引　数】無し
    '【戻り値】無し
    '=========================================================================
    Public Function GetLotChangeJudge() As Boolean
        GetLotChangeJudge = bLotChange
        If bLotChange Then
            bLotChange = False
        End If
    End Function
    '=========================================================================
    '【機　能】連続自動運転モードの初期化
    '【引　数】無し
    '【戻り値】無し
    '=========================================================================
    Public Sub InitialAutoOperation()
        AutoOpeCancel = False
        NowExecuteLotNo = 0
        CancelReason = 0

        MarkingCount = 0                ' マーキング用カウンタクリア	            V2.2.1.7③
        LotMarkingAlarmCnt = 0          ' マーキング実行時アラーム数カウンタクリア	            V2.2.1.7③

        Call PlateDataFileName(0, gsAutoDataFileFullPath(0))  ' プレートデータファイル名を保存
        Call InitLotChangeJudge()
        Call Form1.System1.AutoLoaderFlgReset()                 'V1.2.0.0④ オートローダーフラグリセット
    End Sub

    '=========================================================================
    '【機　能】ロット切り替え判定セット
    '【引　数】無し
    '【戻り値】無し
    '=========================================================================
    Public Function GetAutoOpeCancelStatus() As Boolean
        If gbFgAutoOperation = True Then
            GetAutoOpeCancelStatus = AutoOpeCancel
        Else
            GetAutoOpeCancelStatus = False
        End If
    End Function

    '=========================================================================
    '【機　能】ロット切り替え処理可否チェック
    '【引　数】無し
    '【戻り値】無し
    '=========================================================================
    Public Function LotChangeExecuteCheck() As Boolean

        If NowExecuteLotNo + 1 >= giAutoDataFileNum Then
            LotChangeExecuteCheck = False
        Else
            LotChangeExecuteCheck = True
        End If
    End Function

    '=========================================================================
    '【機　能】ロット切り替え処理
    '【引　数】無し
    '【戻り値】無し
    '=========================================================================
    Public Function LotChangeExecute() As Boolean
        Try
            If LotChangeExecuteCheck() Then
                NowExecuteLotNo = NowExecuteLotNo + 1
                MarkingCount = 0             ' マーキング用カウンタクリア	V2.2.1.7③
                SetAutoOpeStartTime()          ' V2.2.1.7③
                'V2.1.0.0④                Call Form1.TrimDataLoad(gsAutoDataFileFullPath(NowExecuteLotNo))
                'V2.1.0.0④↓
                If Not Form1.TrimDataLoad(gsAutoDataFileFullPath(NowExecuteLotNo)) Then
                    Call Z_PRINT("自動運転時トリミングデータファイルＬＯＡＤエラー" & vbCrLf & "= [" & gsAutoDataFileFullPath(NowExecuteLotNo) & "]")
                    AutoOpeCancel = True
                    LotChangeExecute = False
                    Exit Function
                Else
                    'V2.1.0.0④↑
                    Call PlateDataFileName(0, gsAutoDataFileFullPath(NowExecuteLotNo))  ' プレートデータファイル名を保存
                    LotChangeExecute = True
                End If                      'V2.1.0.0④
            Else
                Call Z_PRINT("ロット切り替え信号を受けましたが、次のエントリーが有りません。" & vbCrLf)
                CancelReason = NO_MORE_ENTRY
                AutoOpeCancel = True
                LotChangeExecute = False
            End If
        Catch ex As Exception
            Call Z_PRINT("FormDataSelect.LotChangeExecute() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function


    '=========================================================================
    '【機　能】連続自動運転終了処理
    '【引　数】無し
    '【戻り値】無し
    '=========================================================================
    Public Sub AutoOperationEnd()

        If gbFgAutoOperation = False Then
            Exit Sub
        End If

        If giLoaderType = 0 Then            'V2.2.1.1②条件追加
            Call Sub_ATLDSET(COM_STS_LOT_END, 0)    'V1.2.0.0④ ローダー出力(ON=ロット終了,OFF=なし)
        End If

        Call Form1.System1.AutoLoaderFlgReset() 'V1.2.0.0④ オートローダーフラグリセット

        Call SavePlateDataFileDelete()

        NowExecuteLotNo = NowExecuteLotNo + 1
        MarkingCount = 0                ' マーキング用カウンタクリア	            V2.2.1.7③
        Form1.DispMarkAlarmList()       ' マーク印字のエラーリストを画面に表示        V2.2.1.7③
        LotMarkingAlarmCnt = 0          ' マーキング実行時アラーム数カウンタクリア	            V2.2.1.7③

        If AutoOpeCancel Or (NowExecuteLotNo < giAutoDataFileNum) Then
            If AutoOpeCancel Then
                NowExecuteLotNo = NowExecuteLotNo - 1   ' 現在のプレートデータから保存する。
            End If
            If NowExecuteLotNo < giAutoDataFileNum Then
                Call SavePlateDataFileFullPath(gsAutoDataFileFullPath, NowExecuteLotNo, giAutoDataFileNum - 1)
            End If
        End If

        Call UserSub.SetStartCheckStatus(True)          'V1.2.0.0④ 設定画面の確認有効化
        gbFgAutoOperation = False
        Form1.AutoRunnningDisp.BackColor = System.Drawing.Color.Yellow  ' 背景色 = 黄色 'V2.0.0.2③
        Form1.AutoRunnningDisp.Text = "自動運転解除中"                                  'V2.0.0.2③

    End Sub

#End Region

    ''' <summary>
    ''' 現在実行中の登録ファイルNoを返す 
    ''' </summary>
    ''' <returns></returns>
    Public Function GetNowLotDataNo() As Integer
        Try

            Return NowExecuteLotNo

        Catch ex As Exception

        End Try

    End Function


    ''' <summary>
    ''' AutoOpeCancel のフラグを設定する 
    ''' </summary>
    ''' <param name="mode"></param>
    Public Sub SetAutoOpeCancel(ByVal mode As Boolean)
        Try
            AutoOpeCancel = mode
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' クランプボタン処理　'V2.2.1.1⑨
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnClamp_Click(sender As Object, e As EventArgs) Handles btnClamp.Click
        Dim r As Integer

        Try

            btnClamp.BackColor = Color.Yellow
            btnClamp.Enabled = False

            ' 載物台クランプON   
            r = Form1.System1.ClampCtrl(gSysPrm, 1, 0)
            If (r <> cFRS_NORMAL) Then

            End If

            System.Threading.Thread.Sleep(500)

            ' 載物台クランプOFF 
            r = Form1.System1.ClampCtrl(gSysPrm, 0, 0)
            If (r <> cFRS_NORMAL) Then

            End If



        Catch ex As Exception
        Finally
            btnClamp.Enabled = True
            btnClamp.BackColor = SystemColors.ButtonFace

        End Try

    End Sub

#End Region
End Class

'=============================== END OF FILE ===============================