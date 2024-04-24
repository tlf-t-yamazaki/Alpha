Option Explicit On
Option Strict On

Imports System
Imports System.Drawing
Imports System.Windows.Forms

Namespace FormEdit
    Public Class cLbl_
        Inherits System.Windows.Forms.Label

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
                .Margin = New Padding(3, 10, 3, 0)
                .TextAlign = ContentAlignment.BottomLeft
                .Size = New Size(120, 26)
            End With
        End Sub
#End Region

#Region "ｲﾍﾞﾝﾄ"
        ''' <summary>Paint ｲﾍﾞﾝﾄ (自動生成)</summary>
        ''' <param name="e"></param>
        Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
            MyBase.OnPaint(e)

            'カスタム描画コードをここに追加します。
        End Sub
#End Region

    End Class
End Namespace
