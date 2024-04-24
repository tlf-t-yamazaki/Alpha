Option Explicit On
Option Strict On

Imports System
Imports System.Drawing
Imports System.Windows.Forms

Namespace FormEdit
    Public Class cCmb_
        Inherits System.Windows.Forms.ComboBox

#Region "宣言"
        Public Const COL_LBLUE As Integer = &HFFFF80        ' 背景色薄いﾌﾞﾙｰ

        Private m_strMsg As String = "????"                 ' ｴﾗｰﾒｯｾｰｼﾞに表示するﾃｷｽﾄ
        Private m_strTip As String = "????"                 ' ﾂｰﾙﾁｯﾌﾟに表示するﾃｷｽﾄ

        Private m_lblToolTip As Label                       ' 編集ﾌｫｰﾑのﾂｰﾙﾁｯﾌﾟを参照する
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
                .Font = New Font("ＭＳ ゴシック", 14.25!, FontStyle.Bold)
                .Size = New Size(80, 27)
            End With
        End Sub
#End Region

#Region "値の設定"
        ''' <summary>ﾂｰﾙﾁｯﾌﾟに表示する文字列を設定</summary>
        ''' <param name="strTip">LblToolTip</param>
        Friend Sub SetStrTip(ByRef strTip As String)
            m_strTip = strTip
        End Sub

        ''' <summary>ﾂｰﾙﾁｯﾌﾟへの参照を設定</summary>
        ''' <param name="toolTip">LblToolTip</param>
        Friend Sub SetLblToolTip(ByRef toolTip As Label)
            m_lblToolTip = toolTip
        End Sub
#End Region

        ''' <summary>編集画面のﾂｰﾙﾁｯﾌﾟﾃｷｽﾄを変更する</summary>
        Private Sub ChangeToolTipText()
            If (Not m_lblToolTip Is Nothing) Then
                m_lblToolTip.Text = m_strTip
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
            Me.BackColor = ColorTranslator.FromOle(COL_LBLUE)
            Me.SelectAll()
            Call ChangeToolTipText()    ' ﾂｰﾙﾁｯﾌﾟﾃｷｽﾄを変更
        End Sub

        ''' <summary>Leave ｲﾍﾞﾝﾄ</summary>
        ''' <param name="e"></param>
        Protected Overrides Sub OnLeave(ByVal e As System.EventArgs)
            MyBase.OnLeave(e)
            Me.BackColor = Color.Empty
            Call RemoveToolTipText()    ' ﾂｰﾙﾁｯﾌﾟﾃｷｽﾄを削除
        End Sub

        ''' <summary>MouseDown ｲﾍﾞﾝﾄ</summary>
        ''' <param name="e"></param>
        Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
            MyBase.OnMouseDown(e)
            Me.BackColor = ColorTranslator.FromOle(COL_LBLUE)
            Me.SelectAll()
            Call ChangeToolTipText()    ' ﾂｰﾙﾁｯﾌﾟﾃｷｽﾄを変更
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
