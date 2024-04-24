Option Explicit On
Option Strict On

Imports System
Imports System.Drawing
Imports System.Windows.Forms

Namespace FormEdit
    Public Class cTxt_
        Inherits System.Windows.Forms.TextBox

#Region "宣言"
        Const COL_LBLUE As Integer = &HFFFF80               ' 背景色薄いﾌﾞﾙｰ

        Private m_bMdownFlg As Boolean = False              ' MouseDownﾌﾗｸﾞ(1回目のみ処理するために使用)

        Private m_strMinVal As String = ""                  ' 入力値下限値または文字列数下限値
        Private m_strMaxVal As String = ""                  ' 入力値上限値または文字列数上限値
        Private m_strFormat As String = ""                  ' 文字列入力書式

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

                .AutoSize = False
                .Font = New Font("ＭＳ ゴシック", 14.25!, FontStyle.Bold)
                '.ImeMode = ImeMode.Disable
                .ShortcutsEnabled = False
                ' 漢字入力の場合は、プロパティでLeftにしてRightにしない。
                If .ImeMode = Windows.Forms.ImeMode.Disable Then
                    .TextAlign = HorizontalAlignment.Right
                End If
                .Size = New Size(100, 26)
            End With
        End Sub
#End Region

#Region "値の設定"
        ''' <summary>ﾎﾞｯｸｽで使用する上下限値を設定する</summary>
        ''' <param name="strMin">下限値(文字列)</param>
        ''' <param name="strMax">上限値(文字列)</param>
        Friend Sub SetMinMax(ByRef strMin As String, ByRef strMax As String)
            m_strMinVal = strMin
            m_strMaxVal = strMax
            Call SetFormat()
        End Sub

        ''' <summary>ﾎﾞｯｸｽで使用する書式を設定する</summary>
        Private Sub SetFormat()
            Dim digMin As Integer
            Dim digMax As Integer

            ' 最小値から小数点以下の桁数を算出
            If (m_strMinVal.IndexOf(".") < 0) Then
                digMin = 0 ' 少数点が含まれない場合
            Else
                If (m_strMinVal.LastIndexOf(".") < m_strMinVal.Length) Then
                    ' m_strMinVal                           : 12.345
                    ' m_strMinVal.Split("."c)(0)            : 12
                    ' m_strMinVal.Split("."c)(1)            : 345 (少数点が文字列の末尾にある場合Nothingとなる)
                    ' (m_strMinVal.Split("."c)(1)).Length   : 3
                    digMin = (m_strMinVal.Split("."c)(1)).Length
                Else
                    digMin = 1 ' 少数点が文字列の末尾にある場合、少数点以下一桁とみなす
                End If
            End If

            ' 最大値から小数点以下の桁数を算出
            If (m_strMaxVal.IndexOf(".") < 0) Then
                digMax = 0 ' 少数点が含まれない場合
            Else
                If (m_strMaxVal.LastIndexOf(".") < m_strMaxVal.Length) Then
                    ' m_strMaxVal                           : 12.345
                    ' m_strMaxVal.Split("."c)(0)            : 12
                    ' m_strMaxVal.Split("."c)(1)            : 345 (少数点が文字列の末尾にある場合Nothingとなる)
                    ' (m_strMaxVal.Split("."c)(1)).Length   : 3
                    digMax = (m_strMaxVal.Split("."c)(1)).Length
                Else
                    digMax = 1 ' 少数点が文字列の末尾にある場合、少数点以下一桁とみなす
                End If
            End If

            ' 最小値と最大値で桁数が異なった場合に桁数が多いほうを使用する
            If (0 = digMin) AndAlso (0 = digMax) Then
                m_strFormat = "0" ' 整数
            ElseIf (digMin <= digMax) Then
                m_strFormat = ("0." & New String("0"c, digMax))
            Else
                m_strFormat = ("0." & New String("0"c, digMin))
            End If

        End Sub

        ''' <summary>ｴﾗｰﾒｯｾｰｼﾞに表示する文字列を設定</summary>
        ''' <param name="strMsg"></param>
        Friend Sub SetStrMsg(ByRef strMsg As String)
            m_strMsg = strMsg
        End Sub

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

#Region "値の取得"
        ''' <summary>上限値文字列を返す</summary>
        ''' <returns>上限値文字列</returns>
        Friend Function GetMaxVal() As String
            GetMaxVal = m_strMaxVal
        End Function

        ''' <summary>下限値文字列を返す</summary>
        ''' <returns>下限値文字列</returns>
        Friend Function GetMinVal() As String
            GetMinVal = m_strMinVal
        End Function

        ''' <summary>文字列の書式を返す(少数点以下の表示桁数で使用する)</summary>
        ''' <returns>文字列書式</returns>
        Friend Function GetStrFormat() As String
            GetStrFormat = m_strFormat
        End Function

        ''' <summary>ｴﾗｰﾒｯｾｰｼﾞで使用するﾃｷｽﾄを返す</summary>
        ''' <returns>ｴﾗｰﾒｯｾｰｼﾞ使用文字列</returns>
        Friend Function GetStrMsg() As String
            GetStrMsg = m_strMsg
        End Function

        ''' <summary>ﾂｰﾙﾁｯﾌﾟに表示するﾃｷｽﾄを返す</summary>
        ''' <returns>ﾂｰﾙﾁｯﾌﾟ表示文字列</returns>
        Friend Function GetStrTip() As String
            GetStrTip = m_strTip
        End Function
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
            m_bMdownFlg = False
        End Sub

        ''' <summary>MouseDown ｲﾍﾞﾝﾄ</summary>
        ''' <param name="e"></param>
        Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
            MyBase.OnMouseDown(e)
            Me.BackColor = ColorTranslator.FromOle(COL_LBLUE)
            If (False = m_bMdownFlg) Then
                Me.SelectAll()
                m_bMdownFlg = True
            End If
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
