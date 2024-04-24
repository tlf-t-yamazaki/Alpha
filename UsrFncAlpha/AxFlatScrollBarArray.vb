'UPGRADE_WARNING: ActiveX コントロール配列を含むフォームを表示するには、プロジェクト全体をコンパイルする必要があります。

Imports System.ComponentModel

<ProvideProperty("Index",GetType(AxMSComCtl2.AxFlatScrollBar))> Public Class AxFlatScrollBarArray
	Inherits Microsoft.VisualBasic.Compatibility.VB6.BaseOcxArray
	Implements IExtenderProvider

	Public Sub New()
		MyBase.New()
	End Sub

	Public Sub New(ByVal Container As IContainer)
		MyBase.New(Container)
	End Sub

	Public Shadows Event [Change] (ByVal sender As System.Object, ByVal e As System.EventArgs)
	Public Shadows Event [Scroll] (ByVal sender As System.Object, ByVal e As System.EventArgs)

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Function CanExtend(ByVal target As Object) As Boolean Implements IExtenderProvider.CanExtend
		If TypeOf target Is AxMSComCtl2.AxFlatScrollBar Then
			Return BaseCanExtend(target)
		End If
	End Function

	Public Function GetIndex(ByVal o As AxMSComCtl2.AxFlatScrollBar) As Short
		Return BaseGetIndex(o)
	End Function

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Sub SetIndex(ByVal o As AxMSComCtl2.AxFlatScrollBar, ByVal Index As Short)
		BaseSetIndex(o, Index)
	End Sub

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Function ShouldSerializeIndex(ByVal o As AxMSComCtl2.AxFlatScrollBar) As Boolean
		Return BaseShouldSerializeIndex(o)
	End Function

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Sub ResetIndex(ByVal o As AxMSComCtl2.AxFlatScrollBar)
		BaseResetIndex(o)
	End Sub

	Default Public ReadOnly Property Item(ByVal Index As Short) As AxMSComCtl2.AxFlatScrollBar
		Get
			Item = CType(BaseGetItem(Index), AxMSComCtl2.AxFlatScrollBar)
		End Get
	End Property

	Protected Overrides Function GetControlInstanceType() As System.Type
		Return GetType(AxMSComCtl2.AxFlatScrollBar)
	End Function

	Protected Overrides Sub HookUpControlEvents(ByVal o As Object)
		Dim ctl As AxMSComCtl2.AxFlatScrollBar = CType(o, AxMSComCtl2.AxFlatScrollBar)
		MyBase.HookUpControlEvents(o)
		If Not ChangeEvent Is Nothing Then
			AddHandler ctl.Change, New System.EventHandler(AddressOf HandleChange)
		End If
		If Not ScrollEvent Is Nothing Then
			AddHandler ctl.Scroll, New System.EventHandler(AddressOf HandleScroll)
		End If
	End Sub

	Private Sub HandleChange (ByVal sender As System.Object, ByVal e As System.EventArgs) 
		RaiseEvent [Change] (sender, e)
	End Sub

	Private Sub HandleScroll (ByVal sender As System.Object, ByVal e As System.EventArgs) 
		RaiseEvent [Scroll] (sender, e)
	End Sub

End Class

