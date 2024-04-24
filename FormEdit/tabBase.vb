Option Explicit On
Option Strict Off

Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Collections.Generic

Namespace FormEdit
    Friend Class tabBase
        Inherits UserControl

#Region "宣言"
        Private tabName As String           ' ﾒｲﾝ編集画面でのﾀﾌﾞ表示名(ﾌﾟﾛﾊﾟﾃｨTAB_NAMEから使用する)
        Protected Friend Property TAB_NAME() As String
            Get
                Return tabName
            End Get
            Protected Set(ByVal value As String)
                tabName = value
            End Set
        End Property

        Private firstControl As Control     ' ﾌｫｰｶｽ設定で使用する(ﾌﾟﾛﾊﾟﾃｨFIRST_CONTROLから使用する)
        Protected Friend Property FIRST_CONTROL() As Control
            Get
                Return firstControl
            End Get
            Protected Set(ByVal value As Control)
                firstControl = value
            End Set
        End Property

        Protected m_MainEdit As frmEdit     ' ﾒｲﾝ編集画面への参照
        Protected m_TabIdx As Integer       ' ﾒｲﾝﾀﾌﾞｺﾝﾄﾛｰﾙ上のｲﾝﾃﾞｯｸｽ
        Protected m_sPath As String = cEDITDEF_FNAME ' "C:\TRIM\EDIT_DEF_User.ini"
        Protected m_CheckFlg As Boolean     ' ﾃﾞｰﾀﾁｪｯｸ中ﾌﾗｸﾞ

        Protected Property m_ResNo() As Integer ' 処理中の抵抗番号(各ﾀﾌﾞ共通)
            Get
                Return m_MainEdit.giRNO
            End Get
            Set(ByVal value As Integer) ' ｸﾞﾛｰﾊﾞﾙへの設定はここでおこなう
                With m_MainEdit
                    .giRNO = value
                    .giCNO = 1 ' 抵抗番号を変更した場合はｶｯﾄ番号を1にする
                End With
            End Set
        End Property

        Protected Property m_CutNo() As Integer ' 処理中のｶｯﾄ番号(各ﾀﾌﾞ共通)
            Get
                Return m_MainEdit.giCNO
            End Get
            Set(ByVal value As Integer) ' ｸﾞﾛｰﾊﾞﾙへの設定はここでおこなう
                m_MainEdit.giCNO = value
            End Set
        End Property

        Protected Property m_GpibNo() As Integer ' 処理中のGP-IB登録番号(各ﾀﾌﾞ共通)
            Get
                Return m_MainEdit.giGNO
            End Get
            Set(ByVal value As Integer) ' ｸﾞﾛｰﾊﾞﾙへの設定はここでおこなう
                m_MainEdit.giGNO = value
            End Set
        End Property
#End Region

        'V2.0.0.0↓
#Region "コンボボックスデータ構造体"
        ''' <summary>
        ''' コンボボックスデータ構造体
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure ComboDataStruct
            ''' <summary>
            ''' 名称
            ''' </summary>
            ''' <remarks></remarks>
            Private strName As String

            ''' <summary>
            ''' 値
            ''' </summary>
            ''' <remarks></remarks>
            Private nValue As Integer

            ''' <summary>
            ''' 名称のプロパティ
            ''' </summary>
            ''' <value>名称</value>
            ''' <returns>文字列</returns>
            ''' <remarks></remarks>
            Public Property Name() As String
                Get
                    Return Me.strName
                End Get
                Set(ByVal value As String)
                    Me.strName = value
                End Set
            End Property

            ''' <summary>
            ''' 値のプロパティ
            ''' </summary>
            ''' <value>値</value>
            ''' <returns>数値</returns>
            ''' <remarks></remarks>
            Public Property Value() As Integer
                Get
                    Return Me.nValue
                End Get
                Set(ByVal value As Integer)
                    Me.nValue = value
                End Set
            End Property

            ''' <summary>
            ''' データ設定
            ''' </summary>
            ''' <param name="strName">名称</param>
            ''' <param name="nValue">値</param>
            ''' <remarks></remarks>
            Public Sub SetData(ByVal strName As String, ByVal nValue As Integer)
                Me.Name = strName
                Me.Value = nValue
            End Sub
        End Structure
#End Region
        'V2.0.0.0↑

#Region "ｺﾝｽﾄﾗｸﾀ"
        ''' <summary>ｺﾝｽﾄﾗｸﾀ</summary>
        Protected Sub New()
            'Friend Sub New(ByRef mainEdit As frmEdit, ByVal tabIdx As Integer)
            ' この呼び出しは、Windows フォーム デザイナで必要です。
            InitializeComponent()

            ' InitializeComponent() 呼び出しの後で初期化を追加します。
            'Call initControl(mainEdit, tabIdx)
        End Sub
#End Region

#Region "初期化処理"
        ''' <summary>ｺﾝﾄﾛｰﾙ初期化処理</summary>
        ''' <param name="mainEdit">ﾒｲﾝ編集画面への参照</param>
        ''' <param name="tabIdx">ﾒｲﾝﾀﾌﾞｺﾝﾄﾛｰﾙ上のｲﾝﾃﾞｯｸｽ</param>
        Protected Overridable Sub InitAllControl(ByRef mainEdit As frmEdit, ByVal tabIdx As Integer)
            'Dim GrpArray() As cGrp_     ' ｸﾞﾙｰﾌﾟﾎﾞｯｸｽの表示設定で使用する
            'Dim LblArray() As cLbl_     ' ﾗﾍﾞﾙへの表示設定で使用する

            'm_TabIdx = tabIdx           ' ﾒｲﾝ編集画面ﾀﾌﾞｺﾝﾄﾛｰﾙ上でのｲﾝﾃﾞｯｸｽ
            'm_MainEdit = mainEdit       ' ﾒｲﾝ編集画面への参照を設定

            'Try

            'Catch ex As Exception
            '    Call MsgBox_Exception(ex)
            'End Try

        End Sub

        ''' <summary>各ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のｺﾝﾄﾛｰﾙに設定をおこなう</summary>
        ''' <param name="ctlArray">各ｸﾞﾙｰﾌﾟﾎﾞｯｸｽごとのｺﾝﾄﾛｰﾙ配列(1次元配列用)</param>
        Protected Sub SetControlData(ByRef ctlArray() As Control)
            Dim txtTag As Integer = 0
            Dim cmbTag As Integer = 0
            Try
                For i As Integer = 0 To (ctlArray.Length - 1) Step 1
                    If (TypeOf ctlArray(i) Is cTxt_) Then ' ﾃｷｽﾄﾎﾞｯｸｽの場合
                        Dim cText As cTxt_ = DirectCast(ctlArray(i), cTxt_)
                        With cText
                            .Tag = txtTag
                            txtTag = (txtTag + 1)
                            ' MeによりOverridesされたinitTextBoxを呼び出す(Meが必要)
                            Call Me.InitTextBox(cText) ' 処理でTagを使用する

                            ' 規定値の場合は数値入力用ﾃｷｽﾄﾎﾞｯｸｽ　かつ　漢字入力無し（Disable）の時
                            If (Short.MaxValue = .MaxLength And .ImeMode = Windows.Forms.ImeMode.Disable) Then
                                AddHandler .TextChanged, AddressOf cTxt_TextChanged ' 入力ﾁｪｯｸをおこなう
                            End If
                        End With
                        Continue For

                    ElseIf (TypeOf ctlArray(i) Is cCmb_) Then ' ｺﾝﾎﾞﾎﾞｯｸｽの場合
                        Dim cCombo As cCmb_ = DirectCast(ctlArray(i), cCmb_)
                        With cCombo
                            ' ﾃﾞｻﾞｲﾝﾓｰﾄﾞ上でDropDownListに設定するとTextが空白となり
                            ' 名前やｲﾝﾃﾞｯｸｽの確認が手間なためここで設定する
                            .DropDownStyle = ComboBoxStyle.DropDownList
                            .Tag = cmbTag
                            cmbTag = (cmbTag + 1)
                            ' MeによりOverridesされたinitComboBoxを呼び出す(Meが必要)
                            Call Me.InitComboBox(cCombo) ' 処理でTagを使用する
                            ' ｲﾍﾞﾝﾄを設定 Me.cCmb_SelectedIndexChanged(Meが必要)
                            AddHandler .SelectedIndexChanged, AddressOf Me.cCmb_SelectedIndexChanged
                        End With
                        Continue For

                    Else
                        Continue For
                    End If
                Next i

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

        ''' <summary>各ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のｺﾝﾄﾛｰﾙに設定をおこなう</summary>
        ''' <param name="ctlArray">各ｸﾞﾙｰﾌﾟﾎﾞｯｸｽごとのｺﾝﾄﾛｰﾙ配列(2次元配列用)</param>
        Protected Sub SetControlData(ByRef ctlArray(,) As Control)
            Dim txtTag As Integer = 0
            Dim cmbTag As Integer = 0
            Try
                For i As Integer = 0 To (ctlArray.GetLength(0) - 1) Step 1
                    For j As Integer = 0 To (ctlArray.GetLength(1) - 1) Step 1
                        If (TypeOf ctlArray(i, j) Is cTxt_) Then ' ﾃｷｽﾄﾎﾞｯｸｽの場合
                            Dim cText As cTxt_ = DirectCast(ctlArray(i, j), cTxt_)
                            With cText
                                .Tag = txtTag
                                txtTag = (txtTag + 1)
                                ' MeによりOverridesされたinitTextBoxを呼び出す(Meが必要)
                                Call Me.InitTextBox(cText) ' 処理でTagを使用する

                                ' 規定値の場合は数値入力用ﾃｷｽﾄﾎﾞｯｸｽ　かつ　漢字入力無し（Disable）の時
                                If (Short.MaxValue = .MaxLength And .ImeMode = Windows.Forms.ImeMode.Disable) Then
                                    AddHandler .TextChanged, AddressOf cTxt_TextChanged ' 入力ﾁｪｯｸをおこなう
                                End If
                            End With
                            Continue For

                        ElseIf (TypeOf ctlArray(i, j) Is cCmb_) Then ' ｺﾝﾎﾞﾎﾞｯｸｽの場合
                            Dim cCombo As cCmb_ = DirectCast(ctlArray(i, j), cCmb_)
                            With cCombo
                                ' ﾃﾞｻﾞｲﾝﾓｰﾄﾞ上でDropDownListに設定するとTextが空白となり
                                ' 名前やｲﾝﾃﾞｯｸｽの確認が手間なためここで設定する
                                .DropDownStyle = ComboBoxStyle.DropDownList
                                .Tag = cmbTag
                                cmbTag = (cmbTag + 1)
                                ' MeによりOverridesされたinitComboBoxを呼び出す(Meが必要)
                                Call Me.InitComboBox(cCombo) ' 処理でTagを使用する
                                ' ｲﾍﾞﾝﾄを設定 Me.cCmb_SelectedIndexChanged(Meが必要)
                                AddHandler .SelectedIndexChanged, AddressOf Me.cCmb_SelectedIndexChanged
                            End With
                            Continue For

                        Else
                            Continue For
                        End If
                    Next j
                    txtTag = 0
                    cmbTag = 0
                Next i

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

        ''' <summary>ｶｰｿﾙｷｰによるﾌｫｰｶｽ移動をおこなうように設定する</summary>
        ''' <param name="ctlArray">ﾌｫｰｶｽ移動をおこなうすべてのｺﾝﾄﾛｰﾙ配列</param>
        Protected Sub SetTabIndex(ByRef ctlArray() As Control)
            Try
                For i As Integer = 0 To (ctlArray.Length - 1) Step 1
                    With ctlArray(i)
                        .TabIndex = i ' ﾀﾌﾞｷｰとｶｰｿﾙｷｰによるﾌｫｰｶｽ移動用
                        ' ｲﾍﾞﾝﾄを設定
                        AddHandler .KeyDown, AddressOf ctlArray_KeyDown
                    End With
                Next i

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#Region "初期化時にｺﾝﾎﾞﾎﾞｯｸｽの設定をおこなう"
        ''' <summary>初期化時にｺﾝﾎﾞﾎﾞｯｸｽのﾘｽﾄ･ﾒｯｾｰｼﾞ設定をおこなう</summary>
        ''' <param name="cCombo">設定をおこなうｺﾝﾎﾞﾎﾞｯｸｽ</param>
        Protected Overridable Sub InitComboBox(ByRef cCombo As cCmb_)
            'Try

            'Catch ex As Exception
            '    Call MsgBox_Exception(ex)
            'End Try

        End Sub
#End Region

#Region "初期化時にﾃｷｽﾄﾎﾞｯｸｽの設定をおこなう"
        ''' <summary>初期化時にﾃｷｽﾄﾎﾞｯｸｽの上下限値･ﾒｯｾｰｼﾞ設定をおこなう</summary>
        ''' <param name="cTextBox">設定をおこなうﾃｷｽﾄﾎﾞｯｸｽ</param>
        Protected Overridable Sub InitTextBox(ByRef cTextBox As cTxt_)
            'Try

            'Catch ex As Exception
            '    Call MsgBox_Exception(ex)
            'End Try

        End Sub
#End Region

#Region "ﾃｷｽﾄﾎﾞｯｸｽに値を設定する"
        ''' <summary>ﾃｷｽﾄﾎﾞｯｸｽに値を設定する</summary>
        Protected Overridable Sub SetDataToText()
            'Try

            'Catch ex As Exception
            '    Call MsgBox_Exception(ex)
            'End Try

        End Sub
#End Region

#Region "すべてのﾃｷｽﾄﾎﾞｯｸｽのﾃﾞｰﾀﾁｪｯｸをおこなう"
        ''' <summary>すべてのﾃｷｽﾄﾎﾞｯｸｽのﾃﾞｰﾀﾁｪｯｸをおこなう</summary>
        ''' <returns>0=正常, 1=ｴﾗｰ</returns>
        Friend Overridable Function CheckAllTextData() As Integer
            Dim ret As Integer
            Try
                ret = 0
            Catch ex As Exception
                ret = 1
            Finally
                CheckAllTextData = ret
            End Try

        End Function
#End Region

#Region "各ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のﾃｷｽﾄﾎﾞｯｸｽをCheckTextData()に送る"
        ''' <summary>ｸﾞﾙｰﾌﾟﾎﾞｯｸｽごとのｺﾝﾄﾛｰﾙ配列</summary>
        ''' <param name="ctlArray"></param>
        ''' <returns>0=正常, 1=ｴﾗｰ</returns>
        Protected Function CheckControlData(ByRef ctlArray() As Control) As Integer
            Dim ret As Integer = 0
            Try
                For Each ctl As Control In ctlArray
                    ' ｺﾝﾄﾛｰﾙが非表示ではない、または無効ではない、かつﾃｷｽﾄﾎﾞｯｸｽの場合
                    If ((False <> ctl.Visible) OrElse (False <> ctl.Enabled)) _
                        AndAlso (TypeOf ctl Is cTxt_) Then
                        ret = Me.CheckTextData(DirectCast(ctl, cTxt_))
                        If (ret <> 0) Then Exit For
                    End If
                Next

            Catch ex As Exception
                Call MsgBox_Exception(ex)
                ret = 1
            Finally
                CheckControlData = ret
            End Try

        End Function

        ''' <summary>各ｸﾞﾙｰﾌﾟﾎﾞｯｸｽ内のﾃｷｽﾄﾎﾞｯｸｽをCheckTextData()に送る</summary>
        ''' <param name="ctlArray">ｸﾞﾙｰﾌﾟﾎﾞｯｸｽごとのｺﾝﾄﾛｰﾙ配列(2次元)</param>
        ''' <returns>0=正常, 1=ｴﾗｰ</returns>
        Protected Function CheckControlData(ByRef ctlArray(,) As Control) As Integer
            Dim ret As Integer = 0
            Try
                For Each ctl As Control In ctlArray
                    ' ｺﾝﾄﾛｰﾙが非表示ではない、または無効ではない、かつﾃｷｽﾄﾎﾞｯｸｽの場合
                    If ((False <> ctl.Visible) OrElse (False <> ctl.Enabled)) _
                        AndAlso (TypeOf ctl Is cTxt_) Then
                        ret = CheckTextData(DirectCast(ctl, cTxt_))
                        If (ret <> 0) Then Exit For
                    End If
                Next

            Catch ex As Exception
                Call MsgBox_Exception(ex)
                ret = 1
            Finally
                CheckControlData = ret
            End Try

        End Function
#End Region

#Region "ﾃｷｽﾄﾎﾞｯｸｽのﾃﾞｰﾀﾁｪｯｸ関数を呼び出す"
        ''' <summary>ﾃｷｽﾄﾎﾞｯｸｽのﾃﾞｰﾀﾁｪｯｸ関数を呼び出す</summary>
        ''' <param name="cTextBox">ﾁｪｯｸするﾃｷｽﾄﾎﾞｯｸｽ</param>
        ''' <returns>0=正常, 1=ｴﾗｰ</returns>
        Protected Overridable Function CheckTextData(ByRef cTextBox As cTxt_) As Integer
            Dim ret As Integer
            Try
                ret = 0
            Catch ex As Exception
                ret = 1
            Finally
                CheckTextData = ret
            End Try

        End Function
#End Region

#Region "NotOverridable"
#Region "Doubleﾃﾞｰﾀﾁｪｯｸ"
        ''' <summary>ﾃﾞｰﾀﾁｪｯｸ処理(Double型)</summary>
        ''' <param name="cTextBox">ﾁｪｯｸするﾃｷｽﾄﾎﾞｯｸｽ</param>
        ''' <param name="setVar">設定する変数への参照</param>
        ''' <returns>0=正常, 1=ｴﾗｰ</returns>
        Protected Function CheckDoubleData(ByRef cTextBox As cTxt_, ByRef setVar As Double) As Integer
            Dim strMsg As String ' ﾒｯｾｰｼﾞ編集域
            Dim MSG As String
            Dim MIN As Double
            Dim MAX As Double
            Dim ret As Integer
            Try
                With cTextBox
                    MSG = .GetStrMsg()
                    MIN = Double.Parse(.GetMinVal())
                    MAX = Double.Parse(.GetMaxVal())

                    ' ﾃﾞｰﾀ範囲ﾁｪｯｸ
                    If (.Text <> "") Then
                        Dim dTmp As Double
                        If (True = Double.TryParse(.Text, dTmp)) Then
                            If (MIN <= dTmp) And (dTmp <= MAX) Then
                                ' 小数点以下の桁数をﾌｫｰﾏｯﾄする(四捨五入される)
                                .Text = dTmp.ToString(.GetStrFormat())
                                setVar = Double.Parse(.Text)
                                ret = 0
                                Exit Try ' ﾁｪｯｸOK
                            End If
                        End If
                    End If

                    ' 範囲ﾁｪｯｸｴﾗｰ
                    If (gSysPrm.stTMN.giMsgTyp = 0) Then
                        strMsg = .GetStrTip() & "(" & MSG & ")"
                    Else
                        strMsg = MIN.ToString() & "～" & MAX.ToString() & "(" & MSG & ")"
                    End If
                    Call MsgBox_CheckErr(cTextBox, strMsg, setVar.ToString(.GetStrFormat()))
                    ret = 1
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex, cTextBox.Name)
                ret = 1
            Finally
                CheckDoubleData = ret
            End Try

        End Function
#End Region

#Region "Shortﾃﾞｰﾀﾁｪｯｸ"
        ''' <summary>ﾃﾞｰﾀﾁｪｯｸ処理(Short型)</summary>
        ''' <param name="cTextBox">ﾁｪｯｸするﾃｷｽﾄﾎﾞｯｸｽ</param>
        ''' <param name="setVar">設定する変数への参照</param>
        ''' <returns>0=正常, 1=ｴﾗｰ</returns>
        Protected Function CheckShortData(ByRef cTextBox As cTxt_, ByRef setVar As Short) As Integer
            Dim strMsg As String ' ﾒｯｾｰｼﾞ編集域
            Dim MSG As String
            Dim MIN As Short
            Dim MAX As Short
            Dim ret As Integer
            Try
                With cTextBox
                    MSG = .GetStrMsg()
                    MIN = Short.Parse(.GetMinVal())
                    MAX = Short.Parse(.GetMaxVal())
                    ' 整数ﾁｪｯｸ
                    ret = (.Text).IndexOf(".")
                    If (0 < ret) Then
                        If (gSysPrm.stTMN.giMsgTyp = 0) Then
                            strMsg = "整数で指定して下さい" & "(" & MSG & ")"
                        Else
                            strMsg = "Please specify it by the integer." & "(" & MSG & ")"
                        End If
                        ret = 1
                        Exit Try
                    End If

                    ' ﾃﾞｰﾀ範囲ﾁｪｯｸ
                    If (.Text <> "") Then
                        Dim sTmp As Short = Short.Parse(.Text)
                        If (MIN <= sTmp) And (sTmp <= MAX) Then
                            .Text = sTmp.ToString(.GetStrFormat())
                            setVar = Short.Parse(.Text)
                            ret = 0
                            Exit Try ' ﾁｪｯｸOK
                        End If
                    End If

                    ' 範囲ﾁｪｯｸｴﾗｰ
                    If (0 = gSysPrm.stTMN.giMsgTyp) Then
                        strMsg = .GetStrTip() & "(" & MSG & ")"
                    Else
                        strMsg = MIN.ToString() & "～" & MAX.ToString() & "(" & MSG & ")"
                    End If

                    Call MsgBox_CheckErr(cTextBox, strMsg, setVar.ToString())
                    ret = 1
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex, cTextBox.Name)
                ret = 1
            Finally
                CheckShortData = ret
            End Try

        End Function
#End Region

#Region "Integerﾃﾞｰﾀﾁｪｯｸ"
        ''' <summary>ﾃﾞｰﾀﾁｪｯｸ処理(Integer型:VB6=Long)</summary>
        ''' <param name="cTextBox">ﾁｪｯｸするﾃｷｽﾄﾎﾞｯｸｽ</param>
        ''' <param name="setVar">設定する変数への参照</param>
        ''' <returns>0=正常, 1=ｴﾗｰ</returns>
        Protected Function CheckIntData(ByRef cTextBox As cTxt_, ByRef setVar As Integer) As Integer
            Dim strMsg As String ' ﾒｯｾｰｼﾞ編集域
            Dim MSG As String
            Dim MIN As Integer
            Dim MAX As Integer
            Dim ret As Integer
            Try
                With cTextBox
                    MSG = .GetStrMsg()
                    MIN = Integer.Parse(.GetMinVal())
                    MAX = Integer.Parse(.GetMaxVal())
                    ' 整数ﾁｪｯｸ
                    ret = (.Text).IndexOf(".")
                    If (0 < ret) Then
                        If (gSysPrm.stTMN.giMsgTyp = 0) Then
                            strMsg = "整数で指定して下さい" & "(" & MSG & ")"
                        Else
                            strMsg = "Please specify it by the integer." & "(" & MSG & ")"
                        End If
                        ret = 1
                        Exit Try
                    End If

                    ' ﾃﾞｰﾀ範囲ﾁｪｯｸ
                    If (.Text <> "") Then
                        Dim iTmp As Integer = Integer.Parse(.Text)
                        If (MIN <= iTmp) And (iTmp <= MAX) Then
                            .Text = iTmp.ToString(.GetStrFormat())
                            setVar = Integer.Parse(.Text)
                            ret = 0
                            Exit Try ' ﾁｪｯｸOK
                        End If
                    End If

                    ' 範囲ﾁｪｯｸｴﾗｰ
                    If (gSysPrm.stTMN.giMsgTyp = 0) Then
                        strMsg = .GetStrTip() & "(" & MSG & ")"
                    Else
                        strMsg = MIN.ToString() & "～" & MAX.ToString() & "(" & MSG & ")"
                    End If

                    Call MsgBox_CheckErr(cTextBox, strMsg, setVar.ToString())
                    ret = 1
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex, cTextBox.Name)
                ret = 1
            Finally
                CheckIntData = ret
            End Try

        End Function
#End Region

#Region "Longﾃﾞｰﾀﾁｪｯｸ"
        ''' <summary>ﾃﾞｰﾀﾁｪｯｸ処理(Long型:VB6=Long)</summary>
        ''' <param name="cTextBox">ﾁｪｯｸするﾃｷｽﾄﾎﾞｯｸｽ</param>
        ''' <param name="setVar">設定する変数への参照</param>
        ''' <returns>0=正常, 1=ｴﾗｰ</returns>
        Protected Function CheckLongData(ByRef cTextBox As cTxt_, ByRef setVar As Long) As Long
            Dim strMsg As String ' ﾒｯｾｰｼﾞ編集域
            Dim MSG As String
            Dim MIN As Long
            Dim MAX As Long
            Dim ret As Long
            Try
                With cTextBox
                    MSG = .GetStrMsg()
                    MIN = Long.Parse(.GetMinVal())
                    MAX = Long.Parse(.GetMaxVal())
                    ' 整数ﾁｪｯｸ
                    ret = (.Text).IndexOf(".")
                    If (0 < ret) Then
                        If (gSysPrm.stTMN.giMsgTyp = 0) Then
                            strMsg = "整数で指定して下さい" & "(" & MSG & ")"
                        Else
                            strMsg = "Please specify it by the integer." & "(" & MSG & ")"
                        End If
                        ret = 1
                        Exit Try
                    End If

                    ' ﾃﾞｰﾀ範囲ﾁｪｯｸ
                    If (.Text <> "") Then
                        Dim LTmp As Long = Long.Parse(.Text)
                        If (MIN <= LTmp) And (LTmp <= MAX) Then
                            .Text = LTmp.ToString(.GetStrFormat())
                            setVar = Long.Parse(.Text)
                            ret = 0
                            Exit Try ' ﾁｪｯｸOK
                        End If
                    End If

                    ' 範囲ﾁｪｯｸｴﾗｰ
                    If (gSysPrm.stTMN.giMsgTyp = 0) Then
                        strMsg = .GetStrTip() & "(" & MSG & ")"
                    Else
                        strMsg = MIN.ToString() & "～" & MAX.ToString() & "(" & MSG & ")"
                    End If

                    Call MsgBox_CheckErr(cTextBox, strMsg, setVar.ToString())
                    ret = 1
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex, cTextBox.Name)
                ret = 1
            Finally
                CheckLongData = ret
            End Try

        End Function
#End Region

#Region "Stringﾃﾞｰﾀﾁｪｯｸ"
        ''' <summary>ﾃﾞｰﾀﾁｪｯｸ処理(String型)</summary>
        ''' <param name="cTextBox">ﾁｪｯｸするﾃｷｽﾄﾎﾞｯｸｽ</param>
        ''' <param name="setVar">設定する変数への参照</param>
        ''' <returns>0=正常, 1=ｴﾗｰ</returns>
        Protected Function CheckStrData(ByRef cTextBox As cTxt_, ByRef setVar As String) As Integer
            Dim strMsg As String ' ﾒｯｾｰｼﾞ編集域
            Dim MSG As String
            Dim MIN As Integer
            Dim MAX As Integer
            Dim len As Integer
            Dim ret As Integer
            Try
                With cTextBox
                    MSG = .GetStrMsg()
                    MIN = Integer.Parse(.GetMinVal())
                    MAX = Integer.Parse(.GetMaxVal())
                    len = .Text.Length
                    ' ﾃﾞｰﾀ範囲ﾁｪｯｸ
                    If (MIN <= len) And (len <= MAX) Then
                        setVar = cTextBox.Text
                        ret = 0
                        Exit Try ' ﾁｪｯｸOK
                    End If

                    ' 範囲ﾁｪｯｸｴﾗｰ
                    If (gSysPrm.stTMN.giMsgTyp = 0) Then
                        strMsg = .GetStrTip() & "(" & MSG & ")"
                    Else
                        strMsg = MIN.ToString("0") & "～" & MAX.ToString("0") & "(" & MSG & ")"
                    End If

                    Call MsgBox_CheckErr(cTextBox, strMsg, setVar.ToString())
                    ret = 1
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex, cTextBox.Name)
                ret = 1
            Finally
                CheckStrData = ret
            End Try

        End Function
#End Region

#Region "HEXﾃﾞｰﾀﾁｪｯｸ"
        ''' <summary>ﾃﾞｰﾀﾁｪｯｸ処理(16進文字列)</summary>
        ''' <param name="cTextBox">ﾁｪｯｸするﾃｷｽﾄﾎﾞｯｸｽ</param>
        ''' <param name="setVar">設定する変数への参照</param>
        ''' <returns>0=正常, 1=ｴﾗｰ</returns>
        Protected Function CheckHexData(ByRef cTextBox As cTxt_, ByRef setVar As Integer) As Integer
            Dim strMsg As String ' ﾒｯｾｰｼﾞ編集域
            Dim MSG As String
            Dim MIN As Integer
            Dim MAX As Integer
            Dim ret As Integer
            Try
                With cTextBox
                    MSG = .GetStrMsg()
                    MIN = Integer.Parse(.GetMinVal())
                    MAX = Integer.Parse(.GetMaxVal())
                    ' ﾃﾞｰﾀ範囲ﾁｪｯｸ
                    If (.Text <> "") Then
                        ret = ObjUtl.Chk_Hex_Char(.Text) ' 16進文字列かどうかを調べる
                        If (0 = ret) Then
                            Dim iTmp As Integer
                            ' ﾃﾞｰﾀ範囲ﾁｪｯｸ
                            iTmp = Math.Abs(Integer.Parse(.Text, Globalization.NumberStyles.HexNumber))
                            If (MIN <= iTmp) And (iTmp <= MAX) Then
                                setVar = iTmp
                                ret = 0
                                Exit Try ' ﾁｪｯｸOK
                            End If
                        End If
                    End If

                    ' 範囲ﾁｪｯｸｴﾗｰ
                    If (0 = gSysPrm.stTMN.giMsgTyp) Then
                        strMsg = (MIN.ToString("X")).ToUpper & "(Hex)～" & _
                                    (MAX.ToString("X")).ToUpper & _
                                    "(Hex)の範囲で指定して下さい" & "(" & MSG & ")"
                    Else
                        strMsg = MIN.ToString("0") & "(Hex) to " & MAX.ToString("0") & "(Hex) (" & MSG & ")"
                    End If

                    Call MsgBox_CheckErr(cTextBox, strMsg, setVar.ToString("X")) ' 16進数文字列を設定する
                    ret = 1
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex, cTextBox.Name)
                ret = 1
            Finally
                CheckHexData = ret
            End Try

        End Function
#End Region

#Region "ﾃｷｽﾄﾎﾞｯｸｽ入力ﾁｪｯｸ"
        ''' <summary>ﾃｷｽﾄﾎﾞｯｸｽを監視して数字以外が入力されたら削除する関数</summary>
        ''' <param name="cTextBox">対象のﾃｷｽﾄﾎﾞｯｸｽ</param>
        Protected Sub CheckTextBoxKeyDown(ByVal cTextBox As cTxt_)
            Try
                With cTextBox
                    '入力文字を１文字づつチェック
                    For i As Integer = 1 To Len(.Text)
                        '最初の１文字目以外に－が入っていたら削除
                        If Mid(.Text, i, 1) = Chr(45) And i <> 1 Then
                            '- を文字列から抜き取りテキストボックスに代入
                            .Text = Mid(.Text, 1, i - 1) & Mid(.Text, i + 1, Len(.Text) - i)
                            Beep()
                            'カーソル位置をテキストの末尾へ
                            .SelectionStart = Len(.Text)
                            Exit Sub
                        End If
                        '0～9 - . 及び　Enter Tab 等の制御文字以外を削除
                        If Mid(.Text, i, 1) >= Chr(32) And Mid(.Text, i, 1) < Chr(45) Or Mid(.Text, i, 1) > Chr(57) Or Mid(.Text, i, 1) = Chr(47) Then
                            '以下上記同様の処理
                            .Text = Mid(.Text, 1, i - 1) & Mid(.Text, i + 1, Len(.Text) - i)
                            Beep()
                            .SelectionStart = Len(.Text)
                            Exit Sub
                        End If
                    Next i

                    If (0 <= .Text.IndexOf(".")) Then ' 小数点がある場合
                        If (0 = .Text.IndexOf(".")) Then
                            ' 文字の最初に入力された場合
                            .Text = .Text.Remove(0, 1) ' 入力された小数点を削除
                            Beep()
                            .SelectionStart = .Text.Length
                        ElseIf ("0" = .GetStrFormat()) Then
                            ' 整数入力用ﾃｷｽﾄﾎﾞｯｸｽの場合
                            .Text = .Text.Substring(0, .Text.IndexOf(".")) ' 小数点以下の文字列を削除
                            Beep()
                            .SelectionStart = .Text.Length
                        Else
                            ' DO NOTHING
                        End If
                    End If
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#Region "ﾃﾞｰﾀﾁｪｯｸｴﾗｰ時に表示するﾒｯｾｰｼﾞﾎﾞｯｸｽ"
        ''' <summary>ﾃｷｽﾄﾎﾞｯｸｽのﾃﾞｰﾀﾁｪｯｸｴﾗｰ時に表示するﾒｯｾｰｼﾞﾎﾞｯｸｽ</summary>
        ''' <param name="cTextBox">ﾁｪｯｸ中のﾃｷｽﾄﾎﾞｯｸｽ</param>
        ''' <param name="strMsg">表示するｴﾗｰﾒｯｾｰｼﾞ</param>
        ''' <param name="text">ﾃｷｽﾄﾎﾞｯｸｽのTextを再設定する場合に使用</param>
        Protected Sub MsgBox_CheckErr(ByRef cTextBox As cTxt_, ByRef strMsg As String, Optional ByRef text As String = "")
            Try
                With cTextBox
                    If ("" <> text) Then .Text = text ' 変更前の状態に戻す
                    .Select() ' ﾌｫｰｶｽｾｯﾄ
                    .SelectAll() ' 入力文字全選択
                    .BackColor = Color.Yellow ' Enterｲﾍﾞﾝﾄが色を変えるためSelect()よりも後におこなう
                End With
                Call MsgBox(strMsg, DirectCast( _
                            MsgBoxStyle.OkOnly + _
                            MsgBoxStyle.Information, MsgBoxStyle), _
                            My.Application.Info.Title)
            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

        ''' <summary>ｺﾝﾎﾞﾎﾞｯｸｽのﾃﾞｰﾀﾁｪｯｸｴﾗｰ時に表示するﾒｯｾｰｼﾞﾎﾞｯｸｽ</summary>
        ''' <param name="cCombo">ﾁｪｯｸ中のｺﾝﾎﾞﾎﾞｯｸｽ</param>
        ''' <param name="strMsg">表示するｴﾗｰﾒｯｾｰｼﾞ</param>
        Protected Sub MsgBox_CheckErr(ByRef cCombo As cCmb_, ByVal strMsg As String)
            Try
                With cCombo
                    .Select() ' ﾌｫｰｶｽｾｯﾄ
                    .BackColor = Color.Yellow ' Enterｲﾍﾞﾝﾄが色を変えるためSelect()よりも後におこなう
                    Call MsgBox(strMsg, DirectCast( _
                                MsgBoxStyle.OkOnly + _
                                MsgBoxStyle.Information, MsgBoxStyle), _
                                My.Application.Info.Title)
                    .BackColor = ColorTranslator.FromOle(cCmb_.COL_LBLUE) ' 背景色をもとに戻す
                End With
            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try
        End Sub
#End Region

#Region "例外発生時に表示するﾒｯｾｰｼﾞﾎﾞｯｸｽ"
        ''' <summary>例外発生時に表示するﾒｯｾｰｼﾞﾎﾞｯｸｽ</summary>
        ''' <param name="ex">発生した例外</param>
        ''' <param name="addMsg">追加表示するﾒｯｾｰｼﾞ</param>
        Protected Sub MsgBox_Exception( _
                ByVal ex As Exception, Optional ByVal addMsg As String = "")
            Dim st As New StackTrace
            Dim msg As String
            Try
                ' GetFrame(0)=GetMethod, GetFrame(1)=CallerMethod
                msg = (st.GetFrame(1).GetMethod.Name & "() TRAP ERROR = " & ex.Message)
                If ("" <> addMsg) Then
                    msg &= ("  --->  " & addMsg)
                End If
#If DEBUG Then
                msg &= (vbCrLf & vbCrLf & ex.StackTrace)
#End If
                Call MsgBox(Me.Name & "." & msg, DirectCast( _
                            MsgBoxStyle.OkOnly + _
                            MsgBoxStyle.Critical, MsgBoxStyle), _
                            My.Application.Info.Title)
            Catch e As Exception
                Call MsgBox(Me.Name & "." & "MsgBox_Exception() TRAP ERROR = " & e.Message, _
                            DirectCast( _
                            MsgBoxStyle.OkOnly + _
                            MsgBoxStyle.Critical, MsgBoxStyle), _
                            My.Application.Info.Title)
            End Try

        End Sub
#End Region

#Region "追加ﾎﾞﾀﾝｸﾘｯｸ時に表示するﾒｯｾｰｼﾞﾎﾞｯｸｽ"
        ''' <summary>追加ﾎﾞﾀﾝｸﾘｯｸ時に表示するﾒｯｾｰｼﾞﾎﾞｯｸｽ</summary>
        ''' <param name="DspMSG">表示ﾀｲﾄﾙ("抵抗ﾃﾞｰﾀ等)</param>
        ''' <param name="Opt">ｵﾌﾟｼｮﾝﾎﾞﾀﾝ(0=前に追加 ,1=後に追加)</param>
        ''' <returns>1:OK(ADVｷｰ), 3:Cancel(RESETｷｰ)</returns>
        Protected Function MsgBox_AddClick(ByRef DspMSG As String, ByRef Opt As Short) As Integer
            Dim strMsg(4) As String ' ﾒｯｾｰｼﾞﾎﾞｯｸｽのｷｬﾌﾟｼｮﾝ表示用
            Dim ret As Integer
            Try
                ' ﾒｯｾｰｼﾞﾎﾞｯｸｽのｷｬﾌﾟｼｮﾝ設定
                strMsg(1) = cAPPcTITLE
                strMsg(2) = DspMSG & "を追加します。" & vbCrLf
                strMsg(2) = strMsg(2) & "よろしいですか？"
                strMsg(3) = "前に追加"
                strMsg(4) = "後に追加"

                ' 確認ﾒｯｾｰｼﾞを表示
                'ret = System1.TrmMsgBox2(SysPrm, strMSG, 1, Opt) ' ﾒｯｾｰｼﾞ表示
                ret = ObjSys.TrmMsgBox2(gSysPrm, strMsg, 1, Opt) ' ﾒｯｾｰｼﾞ表示
                MsgBox_AddClick = ret ' Return値設定

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Function
#End Region

#Region "ｲﾍﾞﾝﾄを発生させずにｺﾝﾎﾞﾎﾞｯｸｽのｲﾝﾃﾞｯｸｽを変更する"
        ''' <summary>ｲﾍﾞﾝﾄを発生させずにｺﾝﾎﾞﾎﾞｯｸｽのｲﾝﾃﾞｯｸｽを変更する</summary>
        ''' <param name="cCombo">ｲﾝﾃﾞｯｸｽを変更するｺﾝﾎﾞﾎﾞｯｸｽ</param>
        ''' <param name="index">設定するｲﾝﾃﾞｯｸｽ番号</param>
        Protected Sub NoEventIndexChange(ByRef cCombo As cCmb_, ByVal index As Integer)
            With cCombo
                If (0 = cCombo.Items.Count) Then Exit Sub ' 起動時
                Try
                    RemoveHandler .SelectedIndexChanged, AddressOf cCmb_SelectedIndexChanged
                    .SelectedIndex = index
                Catch ex As Exception
                    Call MsgBox_Exception(ex, Me.Name & "." & .Name & " :  " & .Text)
                Finally
                    Try
                        AddHandler .SelectedIndexChanged, AddressOf cCmb_SelectedIndexChanged
                    Catch ex As Exception
                        Call MsgBox_Exception(ex, Me.Name & "." & .Name & " : AddHandler")
                    End Try
                End Try
            End With

        End Sub
#End Region
#End Region

#Region "相関ﾁｪｯｸ"
        ''' <summary>相関ﾁｪｯｸ処理</summary>
        ''' <returns>0 = 正常, 1 = ｴﾗｰ</returns>
        Protected Overridable Function CheckRelation() As Integer
            Dim ret As Integer
            Try
                ret = 0
            Catch ex As Exception
                Call MsgBox_Exception(ex)
                ret = 1
            Finally
                CheckRelation = ret
            End Try

        End Function
#End Region

#Region "追加･削除ﾎﾞﾀﾝ関連処理"
        ''' <summary>指定の抵抗ﾃﾞｰﾀを初期化する(複数の派生ｸﾗｽで使用する)</summary>
        ''' <param name="rn">抵抗番号(1 ORG)</param>
        Protected Sub InitResData(ByRef rn As Integer)
            Try
                With m_MainEdit
                    .W_REG(rn).strRNO = ""              ' 抵抗名
                    .W_REG(rn).strTANI = "V "           ' 単位("V","Ω" 等)
                    .W_REG(rn).intSLP = SLP_VTRIMPLS    ' 電圧変化ｽﾛｰﾌﾟ(1:+V, 2:-V, 4:抵抗)
                    .W_REG(rn).lngRel = 0               ' ﾘﾚｰﾋﾞｯﾄ
                    .W_REG(rn).dblNOM = 0.0#            ' ﾄﾘﾐﾝｸﾞ 目標値
                    .W_REG(rn).dblITL = 0.0#            ' 初期判定下限値 (ITLO)
                    .W_REG(rn).dblITH = 0.0#            ' 初期判定上限値 (ITHI)
                    .W_REG(rn).dblFTL = 0.0#            ' 終了判定下限値 (FTLO)
                    .W_REG(rn).dblFTH = 0.0#            ' 終了判定上限値 (FTHI)
                    .W_REG(rn).intMode = 0              ' 判定モード(0:比率(%), 1:数値(絶対値))
                    .W_REG(rn).intPRH = 0               ' HI側ﾌﾟﾛｰﾌﾞ番号
                    .W_REG(rn).intPRL = 0               ' LO側ﾌﾟﾛｰﾌﾞ番号
                    .W_REG(rn).intPRG = 0               ' ｶﾞｰﾄﾞﾌﾟﾛｰﾌﾞ番号
                    .W_REG(rn).intMType = 0             ' 測定種別(0=内部測定, 1=外部測定)
                    .W_REG(rn).intTNN = 1               ' 抵抗内ｶｯﾄ数
                    Call InitCutData(rn)                ' ｶｯﾄﾃﾞｰﾀ初期化
                    Call InitCutCorrData(rn)            ' ｶｯﾄ位置補正ﾃﾞｰﾀ初期化
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

        ''' <summary>指定抵抗の指定ｶｯﾄﾃﾞｰﾀを初期化する(複数の派生ｸﾗｽで使用する)</summary>
        ''' <param name="rn">抵抗番号(1 ORG)</param>
        ''' <param name="cn">ｶｯﾄ番号 (1 ORG) ※0のときは指定抵抗の全ｶｯﾄﾃﾞｰﾀを初期化する</param>
        Protected Sub InitCutData(ByVal rn As Integer, Optional ByVal cn As Integer = 0)
            Dim cNo As Integer
            Dim Num As Integer
            Try
                If (cn = 0) Then ' 全ｶｯﾄﾃﾞｰﾀを初期化
                    cNo = 1 ' 1～
                    Num = MAXCTN ' MAXｶｯﾄ数
                Else ' 指定ｶｯﾄﾃﾞｰﾀを初期化
                    cNo = cn
                    Num = cn
                End If

                With m_MainEdit
                    For cNo = cNo To Num Step 1 ' 指定数分繰返す
                        .W_REG(rn).STCUT(cNo).intCUT = 2      ' ｶｯﾄ方法(1:ﾄﾗｯｷﾝｸﾞ, 2:ｲﾝﾃﾞｯｸｽ, 3:ﾎﾟｼﾞｼｮﾆﾝｸﾞ無しｲﾝﾃﾞｯｸｽ)
                        .W_REG(rn).STCUT(cNo).intCTYP = CNS_CUTP_ST     ' ｶｯﾄ形状(1:ｽﾄﾚｰﾄ, 2:Lｶｯﾄ)
                        .W_REG(rn).STCUT(cNo).dblSTX = 0.0#   ' ﾄﾘﾐﾝｸﾞ ｽﾀｰﾄ点 X
                        .W_REG(rn).STCUT(cNo).dblSTY = 0.0#   ' ﾄﾘﾐﾝｸﾞ ｽﾀｰﾄ点 Y
                        .W_REG(rn).STCUT(cNo).dblSX2 = 0.0#   ' ﾄﾘﾐﾝｸﾞ ｽﾀｰﾄ点2 X
                        .W_REG(rn).STCUT(cNo).dblSY2 = 0.0#   ' ﾄﾘﾐﾝｸﾞ ｽﾀｰﾄ点2 Y
                        .W_REG(rn).STCUT(cNo).dblDL2 = 0.001#   ' 第2のｶｯﾄ長(ﾘﾐｯﾄｶｯﾄ量mm(Lﾀｰﾝ前)) V1.1.0.0① 0.0から0.001へ変更
                        .W_REG(rn).STCUT(cNo).dblDL3 = 0.0#   ' 第3のｶｯﾄ長(ﾘﾐｯﾄｶｯﾄ量mm(Lﾀｰﾝ後))
                        .W_REG(rn).STCUT(cNo).intANG = 0      ' ｶｯﾄ方向1
                        .W_REG(rn).STCUT(cNo).intANG2 = 0     ' ｶｯﾄ方向2
                        .W_REG(rn).STCUT(cNo).intQF1 = 10     ' Qﾚｰﾄ(0.1KHz)
                        .W_REG(rn).STCUT(cNo).dblV1 = 10.0#    ' ﾄﾘﾑ速度(mm/s) V1.1.0.0① 0.0から10.0に変更
                        .W_REG(rn).STCUT(cNo).dblCOF = 0.0#   ' ｶｯﾄｵﾌ(%)
                        .W_REG(rn).STCUT(cNo).dblLTP = 0.0#   ' Lﾀｰﾝ ﾎﾟｲﾝﾄ(%)
                        .W_REG(rn).STCUT(cNo).intTMM = 1      ' ﾓｰﾄﾞ(0:高速(ｺﾝﾊﾟﾚｰﾀ非積分ﾓｰﾄﾞ), 1:高精度(積分ﾓｰﾄﾞ))
                        .W_REG(rn).STCUT(cNo).intMType = 0    ' 外部／内部測定器
                        .W_REG(rn).STCUT(cNo).cFormat = ""    ' ###1042① 文字データ
                        .W_REG(rn).STCUT(cNo).cMarkFix = ""    ' 印字固定部 'V2.2.1.7①
                        .W_REG(rn).STCUT(cNo).cMarkStartNum = ""    ' 開始番号 'V2.2.1.7①
                        .W_REG(rn).STCUT(cNo).intMarkRepeatCnt = 0    ' 重複回数 'V2.2.1.7①

                        'V2.1.0.0①↓ カット毎の抵抗値変化量判定機能追加
                        .W_REG(rn).STCUT(cNo).iVariationRepeat = 0      ' リピート有無
                        .W_REG(rn).STCUT(cNo).iVariation = 0            ' 判定有無
                        .W_REG(rn).STCUT(cNo).dRateOfUp = 0.0           ' 上昇率
                        .W_REG(rn).STCUT(cNo).dVariationLow = -1.0      ' 下限値
                        .W_REG(rn).STCUT(cNo).dVariationHi = 1.0        ' 上限値
                        'V2.1.0.0①↑
                        ' ｲﾝﾃﾞｯｸｽｶｯﾄ情報設定
                        For ix As Integer = 1 To MAXIDX Step 1       ' MAXｲﾝﾃﾞｯｸｽｶｯﾄ数分繰返す
                            .W_REG(rn).STCUT(cNo).intIXN(ix) = 0       ' ｲﾝﾃﾞｯｸｽｶｯﾄ数1-5
                            .W_REG(rn).STCUT(cNo).dblDL1(ix) = 0.0#    ' ｶｯﾄ長1-5
                            .W_REG(rn).STCUT(cNo).lngPAU(ix) = 0       ' ﾋﾟｯﾁ間ﾎﾟｰｽﾞ時間1-5
                            .W_REG(rn).STCUT(cNo).dblDEV(ix) = 0.0#    ' 誤差1-5(%)
                            .W_REG(rn).STCUT(cNo).intIXMType(ix) = 0   ' 測定機器
                            .W_REG(rn).STCUT(cNo).intIXTMM(ix) = 1     ' 測定ﾓｰﾄﾞ
                        Next ix

                        'V1.0.4.3③ ADD ↓
                        For i As Integer = 1 To MAX_LCUT Step 1         ' MAXｽｶｯﾄ数分繰返す
                            .W_REG(rn).STCUT(cNo).dCutLen(i) = 0.001#   ' カット長
                            .W_REG(rn).STCUT(cNo).dQRate(i) = 10        ' Ｑレート
                            .W_REG(rn).STCUT(cNo).dSpeed(i) = 10.0#     ' 速度
                            .W_REG(rn).STCUT(cNo).dAngle(i) = 0         ' 方向（角度）
                            .W_REG(rn).STCUT(cNo).dTurnPoint(i) = 0.0#  ' Ｌターンポイント
                        Next i
                        'V1.0.4.3③ ADD ↑

                        'V2.0.0.0⑦ ADD ↓
                        .W_REG(rn).STCUT(cNo).intRetraceCnt = 0                 ' リトレースカット本数
                        For i As Integer = 1 To MAX_RETRACECUT Step 1           ' MAXｽｶｯﾄ数分繰返す
                            .W_REG(rn).STCUT(cNo).dblRetraceOffX(i) = 0.0       ' リトレースのオフセットＸ
                            .W_REG(rn).STCUT(cNo).dblRetraceOffY(i) = 0.0       ' リトレースのオフセットＹ
                            .W_REG(rn).STCUT(cNo).dblRetraceQrate(i) = 10       ' ストレートカット・リトレースのQレート(0.1KHz)に使用
                            .W_REG(rn).STCUT(cNo).dblRetraceSpeed(i) = 10.0     ' ストレートカット・リトレースのトリム速度(mm/s)に使用
                        Next i
                        'V2.0.0.0⑦ ADD ↑

                        'V2.2.0.0②↓
                        'Uカットパラメータの追加
                        .W_REG(rn).STCUT(cn).dUCutL1 = 0.0          ' L1
                        .W_REG(rn).STCUT(cn).dUCutL2 = 0.0          ' L2
                        .W_REG(rn).STCUT(cn).intUCutQF1 = 0.1       ' Qレート
                        .W_REG(rn).STCUT(cn).dblUCutV1 = 0.1        ' 速度
                        .W_REG(rn).STCUT(cn).intUCutANG = 0         ' 角度
                        .W_REG(rn).STCUT(cn).dblUCutTurnP = 0       ' ターンポイント
                        .W_REG(rn).STCUT(cn).intUCutTurnDir = 1       ' ターン方向
                        .W_REG(rn).STCUT(cn).dblUCutR1 = 0          ' R1
                        .W_REG(rn).STCUT(cn).dblUCutR2 = 0          ' R2
                        'V2.2.0.0②↑

                        ' FL加工条件
                        For fl As Integer = 1 To MAXCND Step 1
                            .W_REG(rn).STCUT(cNo).intCND(fl) = 0       ' FL設定No.
                        Next fl
                    Next cNo
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

        ''' <summary>ｶｯﾄ位置補正ﾃﾞｰﾀを初期化する</summary>
        ''' <param name="rn">抵抗番号</param>
        Private Sub InitCutCorrData(ByVal rn As Integer)
            Try
                With m_MainEdit.W_PTN(rn)
                    .PtnFlg = 0         ' 補正実行(0:なし, 1:自動, 2:自動+手動)
                    .intGRP = 1         ' ﾊﾟﾀｰﾝ登録ｸﾞﾙｰﾌﾟ番号(1-50)
                    .intPTN = 1         ' ﾊﾟﾀｰﾝ登録番号(1-50)
                    .dblPosX = 0.0#     ' ﾊﾟﾀｰﾝ位置X(補正位置ﾃｨｰﾁﾝｸﾞ用)
                    .dblPosY = 0.0#     ' ﾊﾟﾀｰﾝ位置Y(補正位置ﾃｨｰﾁﾝｸﾞ用)
                    .dblDRX = 0.0#      ' ずれ量保存ﾜｰｸX
                    .dblDRY = 0.0#      ' ずれ量保存ﾜｰｸY
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

        ''' <summary>指定登録番号のGP-IBﾃﾞｰﾀを初期化する(複数の派生ｸﾗｽで使用する)</summary>
        ''' <param name="no">登録番号</param>
        Protected Sub InitGpibData(ByVal no As Integer)
            Try
                With m_MainEdit
                    .W_GPIB(no).strGNAM = ""    ' 機器名
                    .W_GPIB(no).intGAD = 0      ' ｱﾄﾞﾚｽ
                    .W_GPIB(no).intDLM = 0      ' ﾃﾞﾘﾐﾀ(0:CRLF, 1:CR, 2:LF, 3:なし)
                    'V2.0.0.0④                    .W_GPIB(no).strCCMD = ""    ' 設定ｺﾏﾝﾄﾞ
                    .W_GPIB(no).strCCMD1 = ""    ' 設定ｺﾏﾝﾄﾞ 'V2.0.0.0④
                    .W_GPIB(no).strCCMD2 = ""    ' 設定ｺﾏﾝﾄﾞ 'V2.0.0.0④
                    .W_GPIB(no).strCCMD3 = ""    ' 設定ｺﾏﾝﾄﾞ 'V2.0.0.0④
                    .W_GPIB(no).strCON = ""     ' ONｺﾏﾝﾄﾞ
                    .W_GPIB(no).lngPOWON = 0    ' ON後のﾎﾟｰｽﾞ時間(ms)
                    .W_GPIB(no).strCOFF = ""    ' OFFｺﾏﾝﾄﾞ
                    .W_GPIB(no).lngPOWOFF = 0   ' OFF後のﾎﾟｰｽﾞ時間(ms)
                    .W_GPIB(no).strCTRG = ""    ' ﾄﾘｶﾞｰｺﾏﾝﾄﾞ
                End With

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

#Region "ｲﾍﾞﾝﾄ"
#Region "NotOverridable"
        ''' <summary>ImeModeの切替が有効な(ひらがな、漢字などを入力する)ﾃｷｽﾄﾎﾞｯｸｽにこのｲﾍﾞﾝﾄを設定する</summary>
        ''' <remarks>文字入力後にﾏｳｽ操作で別のｺﾝﾄﾛｰﾙに移動しようとした場合にﾁｪｯｸをおこない、変数に設定する</remarks>
        Protected Sub cTxt_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs)
            ' TODO: ImeModeの切替をおこなえる(ひらがな、漢字などを入力する)ﾃｷｽﾄﾎﾞｯｸｽにこのｲﾍﾞﾝﾄを設定する
            If (0 <> Me.CheckTextData(DirectCast(sender, cTxt_))) Then
                e.Cancel = True
            End If
        End Sub

        ''' <summary>ﾃｷｽﾄﾎﾞｯｸｽ･ｺﾝﾎﾞﾎﾞｯｸｽでｷｰを押した時の処理</summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Protected Sub ctlArray_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
            Dim KeyCode As Integer = e.KeyCode
            Dim Forward As Boolean = True
            Try
                If (TypeOf sender Is cTxt_) Then ' ﾃｷｽﾄﾎﾞｯｸｽの場合
                    Select Case (KeyCode)
                        Case Keys.Return, Keys.Down ' Returnｷｰ, ↓ｷｰ
                            If (KeyCode = Keys.Return) Then
                                Dim ret As Integer
                                ' Returnｷｰの場合、値のﾁｪｯｸをおこなう
                                ret = Me.CheckTextData(DirectCast(sender, cTxt_))
                                If (ret <> 0) Then Exit Sub
                            End If
                            Forward = True
                        Case Keys.Up ' ↑ｷｰ
                            Forward = False
                        Case Else ' Return,↓,↑ではない場合
                            Exit Sub
                    End Select

                ElseIf (TypeOf sender Is cCmb_) Then ' ｺﾝﾎﾞﾎﾞｯｸｽの場合
                    Select Case (KeyCode)
                        Case Keys.Return, Keys.Right ' Returnｷｰ, →ｷｰ
                            Forward = True
                        Case Keys.Left  ' ←ｷｰ
                            Forward = False
                        Case Else ' Return,→,←ではない場合
                            Exit Sub
                    End Select
                    e.Handled = True ' 移動時にｲﾝﾃﾞｯｸｽが変更されないようにする

                ElseIf (TypeOf sender Is cBtn_) OrElse _
                        (TypeOf sender Is cChk_) OrElse _
                        (TypeOf sender Is cRBtn_) Then ' ﾎﾞﾀﾝ,ﾁｪｯｸﾎﾞｯｸｽ,ﾗｼﾞｵﾎﾞﾀﾝの場合
                    Select Case (KeyCode)
                        Case Keys.Enter, Keys.Return ' Enterｷｰ, Returnｷｰ
                            Exit Sub ' 動作はｲﾍﾞﾝﾄでおこなう
                        Case Keys.Right, Keys.Down ' →ｷｰ, ↓ｷｰ
                            Forward = True
                        Case Keys.Left, Keys.Up ' ↑ｷｰ, ←ｷｰ
                            Forward = False
                        Case Else ' その他のｷｰ
                            Exit Sub
                    End Select

                Else
                    Exit Sub
                End If
                ' 次のｺﾝﾄﾛｰﾙにﾌｫｰｶｽをあてる
                Me.SelectNextControl(DirectCast(sender, Control), Forward, True, True, True)

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

        ''' <summary>ﾃｷｽﾄﾎﾞｯｸｽの内容が変更された時の処理</summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Protected Sub cTxt_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
            Try
                ' 入力文字ﾁｪｯｸをおこなう
                Call CheckTextBoxKeyDown(DirectCast(sender, cTxt_))

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub

        ''' <summary>ｺﾝﾄﾛｰﾙがﾚｲｱｳﾄされる時の処理</summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Protected Sub tabBase_Layout(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LayoutEventArgs) Handles MyBase.Layout
            Try
                If (False = m_CheckFlg) Then    ' ﾃﾞｰﾀﾁｪｯｸ中でなければ
                    Call Me.SetDataToText()     ' ﾃｷｽﾄﾎﾞｯｸｽに値を設定する
                End If

            Catch ex As Exception
                Call MsgBox_Exception(ex)
            End Try

        End Sub
#End Region

        ''' <summary>ｺﾝﾎﾞﾎﾞｯｸｽのｲﾝﾃﾞｯｸｽが変更された時の処理</summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Protected Overridable Sub cCmb_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
            'Try

            'Catch ex As Exception
            '    Call MsgBox_Exception(ex)
            'End Try

        End Sub
#End Region

        'V2.0.0.0↓
#Region "コンボボックスリストデータのテキスト名から値を取得"
        ''' <summary>
        ''' コンボボックスリストデータのテキスト名から値を取得
        ''' </summary>
        ''' <param name="strName">コンボボックスのテキスト名</param>
        ''' <param name="lstData">使用するリストデータ</param>
        ''' <returns>値</returns>
        ''' <remarks></remarks>
        Protected Function GetComboBoxName2Value(ByVal strName As String, ByVal lstData As List(Of ComboDataStruct)) As Integer
            Dim i As Integer

            GetComboBoxName2Value = 0

            For i = 0 To lstData.Count - 1 Step 1
                If strName = lstData(i).Name Then
                    GetComboBoxName2Value = lstData(i).Value
                    Exit For
                End If
            Next i

        End Function
#End Region

#Region "コンボボックスリストデータの値からインデックスを取得"
        ''' <summary>
        ''' コンボボックスリストデータの値からインデックスを取得
        ''' </summary>
        ''' <param name="nValue">値</param>
        ''' <param name="lstData">使用するリストデータ</param>
        ''' <returns>インデックス</returns>
        ''' <remarks></remarks>
        Protected Function GetComboBoxValue2Index(ByVal nValue As Integer, ByVal lstData As List(Of ComboDataStruct)) As Integer
            Dim i As Integer

            GetComboBoxValue2Index = 0

            For i = 0 To lstData.Count - 1 Step 1
                If lstData(i).Value = nValue Then
                    GetComboBoxValue2Index = i
                    Exit For
                End If
            Next i

        End Function
#End Region
        'V2.0.0.0↑

    End Class
End Namespace
