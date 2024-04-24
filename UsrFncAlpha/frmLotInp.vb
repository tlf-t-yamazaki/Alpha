Option Strict Off
Option Explicit On
Friend Class frmLotInp
	Inherits System.Windows.Forms.Form
	'==============================================================================
	'
	'   DESCRIPTION:    ロット番号表示 & 入力
	'
	'==============================================================================
	'-------------------------------------------------------------------------------
	'   内部変数定義
	'-------------------------------------------------------------------------------
	Private Const MAX_LOT_LEN As Short = 64 ' MAXロット番号文字列数
	Private mExitFlag As Short ' 結果(0:初期, 1:OK(ADVｷｰ), 3:Cancel(RESETｷｰ))
	
	'===============================================================================
	'【機　能】 OK/Cancel結果取得
	'【引　数】 なし
	'【戻り値】 結果 = 1:OK(ADVｷｰ), 3:Cancel(RESETｷｰ)
	'===============================================================================
	Public Function GetResult() As Short
		
		GetResult = mExitFlag
		
	End Function
	
	'===============================================================================
	'【機　能】 Cancelボタン押下時処理
	'【引　数】 なし
	'【戻り値】 なし
	'===============================================================================
	Private Sub CmndCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmndCancel.Click
		
		mExitFlag = cFRS_ERR_RST ' ExitFlag = 3:Cancel(RESETｷｰ))
		
	End Sub
	
	'===============================================================================
	'【機　能】 OKボタン押下時処理
	'【引　数】 なし
	'【戻り値】 なし
	'===============================================================================
	Private Sub CmndOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmndOK.Click
		
		Dim strMSG As String
0:      'Dim strDATE As String ' 現在日時("YYYY/MM/DD HH:MM")
        'Dim i As Short
		Dim r As Short
		
        If (TextLOT.Text = "") Then                                     ' 入力なし ?
            TextLOT.Focus()                                             ' フォーカス設定
            Exit Sub
        End If
		
		' 確認ﾒｯｾｰｼﾞ表示
		strMSG = "ロット番号を切り替えます。よろしいですか？"
        r = Form1.System1.TrmMsgBox(gSysPrm, strMSG, MsgBoxStyle.OkCancel, My.Application.Info.Title)
        If (r = cFRS_ERR_RST) Then Exit Sub ' Cancel(RESETｷｰ) ならEXIT

        ' データ範囲チェック
        r = Data_Check()
        If (r <> 0) Then Exit Sub

        ' ロット情報設定
        stUserData.sLotNumber = TextLOT.Text
        Call Disp_frmInfo(COUNTER.PRODUCT_INIT, COUNTER.NONE)                                        ' 生産数初期化(frmInfo画面も再表示)

        ' ログファイル名を設定する ("C:\TRIMDATA\LOG\""LOG_yyyymmdd" + ".LOG")
        Call SetLogFileName(gsLogFileName)

        mExitFlag = cFRS_ERR_ADV                                        ' ExitFlag = 1:OK(ADVｷｰ)

    End Sub

    '===============================================================================
    '【機　能】 データ範囲チェック処理
    '【引　数】 なし
    '【戻り値】 結果 = 0:OK, 0以外:ｴﾗｰ
    '===============================================================================
    Private Function Data_Check() As Short

        On Error GoTo STP_TRAP
        Dim strMSG As String
        Dim iLen As Short

        Data_Check = cFRS_NORMAL ' Return値 = 正常
        iLen = Len(TextLOT.Text)
        If (iLen > MAX_LOT_LEN) Then GoTo STP_ERR ' データ範囲チェック

        Exit Function

STP_ERR:
        strMSG = "ロット番号は" & MAX_LOT_LEN.ToString("0") & "文字以内で指定して下さい"
        Call Form1.System1.TrmMsgBox(gSysPrm, strMSG, MsgBoxStyle.OkOnly, My.Application.Info.Title)
        Data_Check = 1 ' Return値 = データ範囲チェックエラー
        TextLOT.Focus() ' フォーカス設定
        Exit Function

STP_TRAP:
        Data_Check = cERR_TRAP ' Return値 = ﾄﾗｯﾌﾟｴﾗｰ発生

    End Function

    '===============================================================================
    '【機　能】 Form_Activate時処理
    '【引　数】 なし
    '【戻り値】 なし
    '===============================================================================
    'UPGRADE_WARNING: Form イベント frmLotInp.Activate には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
    Private Sub frmLotInp_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        'Dim i As Short
        Dim r As Short
        Dim st As Integer

        ' ロット番号設定
        mExitFlag = 0
        TextLOT.Text = stUserData.sLotNumber ' ロット番号
        TextLOT.Focus() ' フォーカス設定
        TextLOT.SelectionStart = 0 ' ﾃｷｽﾄの選択範囲(先頭文字の直前から)
        TextLOT.SelectionLength = Len(TextLOT.Text) ' 選択範囲の文字数

        ' OK/Calcel入力待ち
        Call ZCONRST() ' コンソールキーラッチ解除
        Do
            System.Windows.Forms.Application.DoEvents() ' メッセージポンプ

            ' 非常停止等チェック
            'r = form1.System1.Sys_Err_Chk(gSysPrm, APP_MODE_LOTCHG, Form1)
            r = Form1.System1.Sys_Err_Chk_EX(gSysPrm, APP_MODE_LOTCHG)
            If (r <> cFRS_NORMAL) Then ' 非常停止等 ?
                mExitFlag = r
                Exit Do
            End If

            ' コンソール入力
            Call ZINPSTS(1, st) ' コンソール入力
            If st And &H4S Then ' ADV キーが押されているか？
                Call ZCONRST() ' コンソールキーラッチ解除
                Call CmndOK_Click(CmndOk, New System.EventArgs()) ' OKボタン押下時処理

            ElseIf st And &H8S Then  ' RESETｷｰ ?
                Call ZCONRST() ' コンソールキーラッチ解除
                Call CmndCancel_Click(CmndCancel, New System.EventArgs()) ' Cancelボタン押下時処理
            End If

        Loop While (mExitFlag = 0)

        Call ZCONRST() ' コンソールキーラッチ解除
        Me.Close()

    End Sub
End Class