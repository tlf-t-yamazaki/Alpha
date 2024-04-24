Public Class frmLoaderInfo

    ' 更新項目指定用 
    Enum LoaderDispMode

        DISP_TACT = 0
        DISP_EXCHANGE
        DISP_TRIMMING
        DISP_SUPPLY_MAGAGINE
        DISP_SUPPLY_SLOT
        DISP_STORE_MAGAGINE
        DISP_STORE_SLOT

    End Enum

    Public saveLoaderInfoDisp As Integer            ' 画面の表示状態保存用    


#Region "フォームの表示"
    ''' <summary>
    ''' フォームの表示
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub frmLoaderInfo_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        Try

            Me.Top = Form1.btnLoaderInfo.Top - 220
            Me.Left = Form1.btnLoaderInfo.Left + (Form1.btnLoaderInfo.Width / 2)


        Catch ex As Exception

        End Try

    End Sub
#End Region

#Region "ローダ情報画面の内容を更新する"
    ''' <summary>
    ''' ローダ情報画面の内容を更新する 
    ''' </summary>
    ''' <param name="mode"></param>
    ''' <param name="setval"></param>
    Public Sub UpdateLoaderInfo(ByVal mode As Integer, ByVal setval As integer)
        Dim tmpval As Double

        Try

            Select Case mode
                Case LoaderDispMode.DISP_TACT
                    tmpval = setval / 10.0
                    lblTact.Text = tmpval.ToString("0.0") + " S"

                Case LoaderDispMode.DISP_EXCHANGE
                    tmpval = setval / 10.0
                    lblExchange.Text = tmpval.ToString("0.0") + " S"

                Case LoaderDispMode.DISP_TRIMMING
                    tmpval = setval / 10.0
                    lblTrimming.Text = tmpval.ToString("0.0") + " S"

                Case LoaderDispMode.DISP_SUPPLY_MAGAGINE
                    lblSupplyMag.Text = setval.ToString("0")

                Case LoaderDispMode.DISP_SUPPLY_SLOT    'トリミング枚数として使用する
                    lblSupplySlot.Text = setval.ToString("0")

                Case LoaderDispMode.DISP_STORE_MAGAGINE
                    'lblStoreMag.Text = setval.ToString("0")

                Case LoaderDispMode.DISP_STORE_SLOT
                    'lblStoreSlot.Text = setval.ToString("0")

                Case Else

            End Select

        Catch ex As Exception

        End Try

    End Sub
#End Region

End Class