Namespace FormEdit
    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
    Partial Class cLbl_
        Inherits System.Windows.Forms.Label

        ' AutoSizeを強制的にFalseにする
        <System.ComponentModel.DefaultValue(False)> _
        Public Overrides Property AutoSize() As Boolean
            Get
                Return MyBase.AutoSize
            End Get
            Set(ByVal value As Boolean)
                MyBase.AutoSize = False
            End Set
        End Property

        'Control は、コンポーネント一覧に後処理を実行するために、dispose をオーバーライドします。
        <System.Diagnostics.DebuggerNonUserCode()> _
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            Try
                If disposing AndAlso components IsNot Nothing Then
                    components.Dispose()
                End If
            Finally
                MyBase.Dispose(disposing)
            End Try
        End Sub

        'コントロール デザイナで必要です。
        Private components As System.ComponentModel.IContainer

        ' メモ: 以下のプロシージャはコンポーネント デザイナで必要です。
        ' コンポーネント デザイナを使って変更できます。
        ' コード エディタを使って変更しないでください。
        <System.Diagnostics.DebuggerStepThrough()> _
        Private Sub InitializeComponent()
            components = New System.ComponentModel.Container()
        End Sub

    End Class
End Namespace
