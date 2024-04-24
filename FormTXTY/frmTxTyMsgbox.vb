'===============================================================================
'   Description  : ＴＸ/ＴＹティーチング画面終了確認処理
'
'   Copyright(C) : TOWA LASERFRONT CORP. 2011
'
'===============================================================================
Option Strict Off
Option Explicit On
Friend Class frmTxTyMsgbox
	Inherits System.Windows.Forms.Form
#Region "プライベート変数定義"
    '===========================================================================
    '   変数定義
    '===========================================================================
    Private mEndFlag As Short
#End Region

#Region "終了結果を返す"
    '''=========================================================================
    '''<summary>終了結果を返す</summary>
    '''<returns>cFRS_ERR_START=OK(STARTｷｰ)
    '''         cFRS_ERR_RST  =Cancel(RESETｷｰ)
    '''         cFRS_TxTy     =TX2/TY2押下</returns>
    '''=========================================================================
    Public ReadOnly Property sGetReturn() As Short
        Get
            sGetReturn = mEndFlag
        End Get
    End Property
#End Region

#Region "Cancelボタン押下時処理"
    '''=========================================================================
    '''<summary>Cancelボタン押下時処理</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub cmdCAN_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCAN.Click
        mEndFlag = cFRS_ERR_RST                             ' Cancel(RESETｷｰ)
    End Sub
#End Region

#Region "OKボタン押下時処理"
    '''=========================================================================
    '''<summary>OKボタン押下時処理</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
        mEndFlag = cFRS_ERR_START                                       ' OK(STARTｷｰ) 
    End Sub
#End Region

#Region "TX2ボタン押下時処理"
    '''=========================================================================
    '''<summary>TX2ボタン押下時処理</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub cmdOKTxTy_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOKTxTy.Click
        mEndFlag = cFRS_TxTy                                ' TX2(Teach)/TY2押下
    End Sub
#End Region

#Region "フォームInitialize時処理"
    '''=========================================================================
    '''<summary>フォームInitialize時処理</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub Form_Initialize_Renamed()

        mEndFlag = -1
        Me.Text = MSG_CLOSE_LABEL01                         ' "画面終了確認"
        Me.cmdOK.Text = MSG_CLOSE_LABEL02                   ' "はい(&Y)"
        Me.cmdCAN.Text = MSG_CLOSE_LABEL03                  ' "いいえ(&N)"
        Me.cmdOKTxTy.Text = MSG_EXECUTE_TXTYLABEL           ' "TX2(&T)" or "TY2(&T)"
        Me.Label1.Text = MSG_105                            ' "前の画面に戻ります。よろしいですか？"

        ' TY2ボタンはメイン画面のTY2ボタンが有効時に表示する
        If (giAppMode = APP_MODE_TY) Then                   ' TY2コマンド ? 
            If (stFNC(F_TY2).iDEF = 1) Then           ' TY2ボタンが有効 ?
                Me.cmdOKTxTy.Visible = True
            Else
                Me.cmdOKTxTy.Visible = False                ' OKボタンをTY2ボタンの表示位置へ移動 
                Me.cmdOK.Location = Me.cmdOKTxTy.Location
            End If
        End If

    End Sub
#End Region

#Region "フォームActivate時処理"
    '''=========================================================================
    '''<summary>フォームActivate時処理</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub frmTxTyMsgbox_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        Dim cin As Integer
        Dim r As Integer = 0
        Dim strMSG As String

        Try
            If (mEndFlag <> -1) Then Exit Sub
            mEndFlag = 0
            Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
            Call ZCONRST()                                          ' ｺﾝｿｰﾙｷｰﾗｯﾁ解除

            Do
                System.Windows.Forms.Application.DoEvents()         ' メッセージポンプ
#If cOFFLINEcDEBUG = 0 Then
                r = STARTRESET_SWCHECK(False, cin)                  ' コンソール入力(監視なしモード)
#Else
                cin = 0
#End If
                If (cin = cFRS_ERR_RST) Then                        ' RESET キー押下？
                    Call cmdCAN_Click(cmdCAN, New System.EventArgs())
                ElseIf (cin = cFRS_ERR_START) Then                  ' START キーが押されているか？
                    Call cmdOK_Click(cmdOK, New System.EventArgs())
                End If
                Call ZCONRST()                                      ' ｺﾝｿｰﾙｷｰﾗｯﾁ解除
                System.Windows.Forms.Application.DoEvents()         ' メッセージポンプ
                Me.Refresh()
            Loop While (mEndFlag = 0)

            ' トラップエラー発生時 
        Catch ex As Exception
            strMSG = "frmTxTyMsgbox.frmTxTyMsgbox_Activated() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try

        Call ZCONRST()                                              ' ｺﾝｿｰﾙｷｰﾗｯﾁ解除
        Me.Close()

    End Sub
#End Region

#Region "キー押下時処理"
    '''=========================================================================
    '''<summary>キー押下時処理</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Private Sub frmTxTyMsgbox_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        If UCase(Chr(KeyAscii)) = UCase("y") Then
            Call cmdOK_Click(cmdOK, New System.EventArgs())
        End If

        If UCase(Chr(KeyAscii)) = UCase("n") Then
            Call cmdCAN_Click(cmdCAN, New System.EventArgs())
        End If

        KeyAscii = 0
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If

    End Sub
#End Region

End Class