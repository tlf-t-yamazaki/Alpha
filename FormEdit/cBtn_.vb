Option Explicit On
Option Strict On

Imports System
Imports System.Drawing
Imports System.Windows.Forms

Namespace FormEdit
    Public Class cBtn_
        Inherits System.Windows.Forms.Button

#Region "宣言"
        Private m_lblToolTip As Label   ' 編集ﾌｫｰﾑのﾂｰﾙﾁｯﾌﾟを参照する
#End Region

#Region "初期化"
        ''' <summary>ｺﾝｽﾄﾗｸﾀ</summary>
        Public Sub New()
            ' この呼び出しは、Windows フォーム デザイナで必要です。
            InitializeComponent()

            ' InitializeComponent() 呼び出しの後で初期化を追加します。
            Call InitControl()
        End Sub

        ''' <summary>初期化</summary>
        Private Sub InitControl()
            With Me
                .Font = New Font("MS UI Gothic", 12.0!, FontStyle.Regular)
                .Size = New Size(100, 50)
            End With
        End Sub
#End Region

#Region "値の設定"
        ''' <summary>ﾂｰﾙﾁｯﾌﾟへの参照を設定</summary>
        ''' <param name="toolTip">LblToolTip</param>
        Friend Sub SetLblToolTip(ByRef toolTip As Label)
            m_lblToolTip = toolTip
        End Sub
#End Region

        ''' <summary>編集画面のﾂｰﾙﾁｯﾌﾟﾃｷｽﾄを変更する</summary>
        Private Sub ChangeToolTipText()
            If (Not m_lblToolTip Is Nothing) Then
                m_lblToolTip.Text = Me.Text
            End If
        End Sub

        ''' <summary>編集画面のﾂｰﾙﾁｯﾌﾟﾃｷｽﾄを削除する</summary>
        Private Sub RemoveToolTipText()
            If (Not m_lblToolTip Is Nothing) Then
                m_lblToolTip.Text = ""
            End If
        End Sub

#Region "ｲﾍﾞﾝﾄ"
        ''' <summary>Enter ｲﾍﾞﾝﾄ</summary>
        ''' <param name="e"></param>
        Protected Overrides Sub OnEnter(ByVal e As System.EventArgs)
            MyBase.OnEnter(e)
            Call ChangeToolTipText()    ' ﾂｰﾙﾁｯﾌﾟﾃｷｽﾄを変更
        End Sub

        ''' <summary>Leave ｲﾍﾞﾝﾄ</summary>
        ''' <param name="e"></param>
        Protected Overrides Sub OnLeave(ByVal e As System.EventArgs)
            MyBase.OnLeave(e)
            Call RemoveToolTipText()    ' ﾂｰﾙﾁｯﾌﾟﾃｷｽﾄを削除
        End Sub

        ''' <summary>ﾎﾞﾀﾝ上でのｶｰｿﾙｷｰ入力を有効にする</summary>
        ''' <param name="e"></param>
        Protected Overrides Sub OnPreviewKeyDown(ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs)
            MyBase.OnPreviewKeyDown(e)

            Dim KeyCode As Integer = e.KeyCode
            Select Case (KeyCode)
                Case Keys.Left, Keys.Up, Keys.Right, Keys.Down ' 有効にするｷｰ
                    e.IsInputKey = True
                Case Else
                    Exit Sub
            End Select
        End Sub

        ''' <summary>Paint ｲﾍﾞﾝﾄ (自動生成)</summary>
        ''' <param name="e"></param>
        Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
            MyBase.OnPaint(e)

            'カスタム描画コードをここに追加します。
        End Sub
#End Region

    End Class
End Namespace
