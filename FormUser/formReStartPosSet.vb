Public Class formReStartPosSet


    ''' <summary>
    ''' 再測定開始位置画面でＯＫを押したときの処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnOK_Click(sender As System.Object, e As System.EventArgs) Handles BtnOK.Click

        Dim iPlate As Integer
        Dim iBlockX As Integer
        Dim iBlockY As Integer

        '基板番号のチェック 
        If Integer.TryParse(txtPlateNo.Text.ToString(), iPlate) Then
            ' 
        Else
            If (gSysPrm.stTMN.giMsgTyp = 0) Then
                Call MsgBox("基板番号に数値を設定してください。")
            Else
                Call MsgBox("Please set PlateNo numeric value. ")
            End If
            Return
        End If

        'BlockXのチェック 
        If Integer.TryParse(txtBlockX.Text.ToString(), iBlockX) Then
            ' 入力範囲チェック
            If (iBlockX <= 0) OrElse (stPLT.BNX < iBlockX) Then
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    Call MsgBox("ブロックＸには１から" + stPLT.BNX.ToString() + "の値を入力してください。")
                Else
                    Call MsgBox("Please input BlockX from 1 to " + stPLT.BNX.ToString())
                End If
                Return
            End If
        Else
            If (gSysPrm.stTMN.giMsgTyp = 0) Then
                Call MsgBox("ブロックＸに数値を設定してください。")
            Else
                Call MsgBox("Please input BlockX numeric value.")
            End If

            Return
        End If

        'BlockYのチェック 
        If Integer.TryParse(txtBlockY.Text.ToString(), iBlockY) Then
            ' 入力範囲チェック
            If (iBlockY <= 0) OrElse (stPLT.BNY < iBlockY) Then
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    Call MsgBox("ブロックＹには１から" + stPLT.BNY.ToString() + "の値を入力してください。")
                Else
                    Call MsgBox("Please input BlockY from 1 to " + stPLT.BNY.ToString())
                End If
                Return
            End If
        Else
            If (gSysPrm.stTMN.giMsgTyp = 0) Then
                Call MsgBox("ブロックＹに数値を設定してください。")
            Else
                Call MsgBox("Please input BlockY numeric value.")
            End If

            Return
        End If

        UserSub.gVariationMeasPlateStartNo = iPlate
        UserSub.gVariationMeasBlockXStartNo = iBlockX
        UserSub.gVariationMeasBlockYStartNo = iBlockY
        UserSub.bVariationMesStep = True

        Close()

    End Sub

    Private Sub BtnCancel_Click(sender As System.Object, e As System.EventArgs) Handles BtnCancel.Click
        Close()
    End Sub

    Private Sub formReStartPosSet_Shown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Shown
        txtPlateNo.Text = UserSub.gVariationMeasPlateStartNo.ToString("0")
        txtBlockX.Text = UserSub.gVariationMeasBlockXStartNo.ToString("0")
        txtBlockY.Text = UserSub.gVariationMeasBlockYStartNo.ToString("0")
    End Sub
End Class