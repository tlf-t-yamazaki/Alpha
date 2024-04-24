Option Explicit On
Option Strict On

Imports System
Imports System.Drawing
Imports System.Windows.Forms

Namespace FormEdit
    Public Class cPnl_
        Inherits System.Windows.Forms.Panel

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
                .Size = New Size(1236, 760)
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
