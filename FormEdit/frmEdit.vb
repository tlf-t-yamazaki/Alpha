'==============================================================================
'
'   DESCRIPTION:    パラメータ編集画面処理('10.07.22 A.W)
'
'==============================================================================
Option Explicit On
Option Strict On

Imports System
Imports System.Drawing
Imports System.Windows.Forms

#Const _PTN_TAB = True          ' ﾊﾟﾀｰﾝﾀﾌﾞ ｺﾒﾝﾄｱｳﾄで非表示
#Const _GPIB_TAB = True         ' GP-IBﾀﾌﾞ ｺﾒﾝﾄｱｳﾄで非表示

'V1.0.4.3③使用しないので関連ソースを削除 #Const _ARRAY_DATA = False      ' ﾌﾟﾚｰﾄﾃﾞｰﾀが配列かどうか(frmEdit.vbのみ)

Namespace FormEdit
    Friend Class frmEdit
        Inherits System.Windows.Forms.Form

#Region "定数定義"
        '---------------------------------------------------------------------------
        '   定数定義
        '---------------------------------------------------------------------------
        Private m_tabUser As tabUser                            ' ユーザタブ
        Private m_tabUser2 As tabUser2                          ' ユーザタブ2 'V2.0.0.0②
        Private m_tabSystem As tabSystem                        ' ｼｽﾃﾑﾀﾌﾞ
        Private m_tabResistor As tabResistor                    ' 抵抗ﾀﾌﾞ
        Private m_tabCut As tabCut                              ' ｶｯﾄﾀﾌﾞ

#If _PTN_TAB Then
        Private m_tabPattern As tabPattern                      ' ﾊﾟﾀｰﾝ登録ﾀﾌﾞ
#End If

#If _GPIB_TAB Then
        Private m_tabGPIB As tabGpib                            ' GP-IBﾀﾌﾞ
#End If

        Private Const TAB_COUNT As Integer = (7 - 1)            ' (ﾀﾌﾞの枚数 - 1)'V2.0.0.0②
        'V2.0.0.0②#If _PTN_TAB And _GPIB_TAB Then
        'V2.0.0.0②        Private Const TAB_COUNT As Integer = (6 - 1)            ' (ﾀﾌﾞの枚数 - 1)
        'V2.0.0.0②#ElseIf _PTN_TAB Or _GPIB_TAB Then
        'V2.0.0.0②        Private Const TAB_COUNT As Integer = (5 - 1)            ' (ﾀﾌﾞの枚数 - 1)
        'V2.0.0.0②#Else
        'V2.0.0.0②        Private Const TAB_COUNT As Integer = (4 - 1)            ' (ﾀﾌﾞの枚数 - 1)
        'V2.0.0.0②#End If

        '---------------------------------------------------------------------------
        '   編集用ﾃﾞｰﾀ域
        '---------------------------------------------------------------------------
        Friend W_stUserData As USER_DATA                        ' ユーザデータ
        Friend W_PLT As PLATE_DATA                              ' ﾌﾟﾚｰﾄﾃﾞｰﾀ [ 1 ]ORG
        Friend W_LASER As POWER_DATA                            ' ﾊﾟﾜｰ制御用ﾃﾞｰﾀ
        Friend W_REG(MAXRNO) As Reg_Info                        ' 抵抗ﾃﾞｰﾀ [ 1 ]ORG

#If cOSCILLATORcFLcUSE Then
        Friend W_FLCND As TrimCondInfo                          ' FL加工条件 [ 0 ]ORG
#End If
        Friend W_PTN(MAXRGN) As Ptn_Info                        ' ﾊﾟﾀｰﾝ登録ﾃﾞｰﾀ(ｶｯﾄ位置補正用) [ 1 ]ORG
        Friend W_THE As Theta_Info                              ' ﾊﾟﾀｰﾝ登録ﾃﾞｰﾀ(XYθ補正用) [ 1 ]ORG
        Friend W_GPIB(MAXGNO) As GPIB_DATA                      ' GP-IBﾃﾞｰﾀ [ 1 ]ORG

        '---------------------------------------------------------------------------
        '   その他
        '---------------------------------------------------------------------------
        Private flgChk As Boolean               ' ﾃﾞｰﾀﾁｪｯｸ中ﾌﾗｸﾞ(False:ﾁｪｯｸ中でない, True:ﾁｪｯｸ中)
        Private flgClose As Boolean             ' FormClosingｲﾍﾞﾝﾄで使用する(True:閉じる, False:閉じない)
        Private exitflg As Integer              ' 編集画面を抜けるときのボタン：'V2.2.1.6①

        Friend giRNO As Integer                 ' 各ﾀﾌﾞ共有処理中抵抗番号(1 ORG)
        Friend giCNO As Integer                 ' 各ﾀﾌﾞ共有処理中ｶｯﾄ番号 (1 ORG)
        Friend giGNO As Integer                 ' 処理中GPIB番号         (1 ORG)

        Private procHandle1 As Process          ' ｿﾌﾄｳｪｱｷｰﾎﾞｰﾄﾞの起動･終了で使用
        Private strProc As String = "OSK"       ' ｿﾌﾄｳｪｱｷｰﾎﾞｰﾄﾞの起動･終了で使用
#End Region

#Region "ﾌｫｰﾑの初期化"
        ''' <summary>ﾌｫｰﾑ初期化処理</summary>
        Private Sub Form_Initialize_Renamed()
            '---------------------------------------------------------------------------
            '   ﾃﾞｰﾀを編集用ﾃﾞｰﾀ域に設定する
            '---------------------------------------------------------------------------
            flgChk = False      ' ﾃﾞｰﾀﾁｪｯｸ中ﾌﾗｸﾞ = False:ﾁｪｯｸ中でない
            flgClose = False    ' FormClosingｲﾍﾞﾝﾄで使用する = False:閉じない

            giRNO = 1 ' 処理中抵抗番号     (1 ORG)
            giCNO = 1 ' 処理中ｶｯﾄ番号      (1 ORG)
            giGNO = 1 ' 処理中GP-IB登録番号(1 ORG)

            Try
                ' ユーザデータ
                Call ReadUserData()

                ' ｼｽﾃﾑﾃﾞｰﾀ
                Call ReadPlateData()

                ' ﾊﾟﾜｰ調整ﾃﾞｰﾀ
                W_LASER = stLASER

                ' 抵抗/ｶｯﾄﾃﾞｰﾀ
                Call ReadResistorData()

#If cOSCILLATORcFLcUSE Then
                ' FL加工条件ﾃﾞｰﾀ(表示のみ)
                Call ReadFlConditionData()
#End If

                ' ﾊﾟﾀｰﾝ登録ﾃﾞｰﾀ(ｶｯﾄ位置補正用)
                W_PTN = DirectCast(stPTN.Clone, Ptn_Info())

                ' θ補正ﾃﾞｰﾀ(XYθ補正用)
                W_THE = stThta

                ' GP-IBﾃﾞｰﾀ
                W_GPIB = DirectCast(stGPIB.Clone, GPIB_DATA())

                ' 各ﾀﾌﾞを配置する
                Call LayoutTab()

            Catch ex As Exception
                Call MsgBox_Exception(ex.Message)
            End Try

        End Sub
#End Region

#Region "メインフォームからユーザデータを読み込む"

        '''=========================================================================
        ''' <summary>メインフォームからユーザデータを読み込む</summary>
        '''=========================================================================
        Private Sub ReadUserData()
            'V2.0.0.0⑪↓
            Try
                W_stUserData = stUserData
                W_stUserData.iResUnit = DirectCast(stUserData.iResUnit.Clone(), Integer())              ' 温度センサー 抵抗レンジ 1:Ω, 2:KΩ
                W_stUserData.dNomCalcCoff = DirectCast(stUserData.dNomCalcCoff.Clone(), Double())       ' 温度センサー 抵抗レンジ 1:Ω, 2:KΩ
                W_stUserData.dTargetCoff = DirectCast(stUserData.dTargetCoff.Clone(), Double())         ' 温度センサー 抵抗レンジ 1:Ω, 2:KΩ
                W_stUserData.dTargetCoffJudge = DirectCast(stUserData.dTargetCoffJudge.Clone(), Double())   ' 目標値算出係数 'V2.1.0.0③
                W_stUserData.iChangeSpeed = DirectCast(stUserData.iChangeSpeed.Clone(), Integer())      ' 温度センサー 抵抗レンジ 1:Ω, 2:KΩ
            Catch ex As Exception
                Call MsgBox_Exception(ex.Message)
            End Try
            'V2.0.0.0⑪↑
            'With W_stUserData
            '    .iTrimType = stUserData.iTrimType                       ' 製品種別
            '    .sLotNumber = stUserData.sLotNumber                     ' ロット番号
            '    .sOperator = stUserData.sOperator                       ' オペレータ名
            '    .sPatternNo = stUserData.sPatternNo                     ' パターンＮｏ．
            '    .sProgramNo = stUserData.sProgramNo                     ' プログラムＮｏ．
            '    .iTrimSpeed = stUserData.iTrimSpeed                     ' トリミング速度
            '    .iLotChange = stUserData.iLotChange                     ' ロット終了条件
            '    .lLotEndSL = stUserData.lLotEndSL                       ' ロット処理枚数
            '    .lCutHosei = stUserData.lCutHosei                       ' カット位置補正頻度
            '    .lPrintRes = stUserData.lPrintRes                       ' ロット終了時印刷素子数
            '    .iTempResUnit = stUserData.iTempResUnit                 ' 温度センサー 抵抗レンジ 1:Ω, 2:KΩ
            '    .iTempTemp = stUserData.iTempTemp                       ' 参照温度	１：０℃ または ２：２５℃
            '    .dStandardRes0 = stUserData.dStandardRes0               ' 標準抵抗値 ０℃	0.01～100M
            '    .dStandardRes25 = stUserData.dStandardRes25             ' 標準抵抗値 ２５℃	0.01～100M
            '    .dResTempCoff = stUserData.dResTempCoff                 ' 抵抗温度係数
            '    .dFinalLimitHigh = stUserData.dFinalLimitHigh           ' ファイナルリミット　Hight[%]
            '    .dFinalLimitLow = stUserData.dFinalLimitLow             ' ファイナルリミット　Lo[%]
            '    .dRelativeHigh = stUserData.dRelativeHigh               ' 相対値リミット　Hight[%]
            '    .dRelativeLow = stUserData.dRelativeLow                 ' 相対値リミット　Lo[%]
            '    .Initialize()
            '    For rn As Integer = 1 To MAX_RES_USER Step 1
            '        .iResUnit(rn) = stUserData.iResUnit(rn)             ' 温度センサー 抵抗レンジ 1:Ω, 2:KΩ
            '        .dNomCalcCoff(rn) = stUserData.dNomCalcCoff(rn)     ' 補正値（ノミナル値算出係数）
            '        .dTargetCoff(rn) = stUserData.dTargetCoff(rn)       ' 目標値算出係数
            '        .iChangeSpeed(rn) = stUserData.iChangeSpeed(rn)     ' 測定速度を変更するカットNo.
            '    Next rn
            'End With ' W_stUserData


        End Sub
#End Region

#Region "ﾒｲﾝﾌｫｰﾑからのﾃﾞｰﾀ読み込み"

#Region "ﾒｲﾝﾌｫｰﾑからﾌﾟﾚｰﾄﾃﾞｰﾀを読み込む"
        ''' <summary>ﾒｲﾝﾌｫｰﾑからﾌﾟﾚｰﾄﾃﾞｰﾀを読み込む</summary>
        Private Sub ReadPlateData()
            Try
                W_PLT = stPLT
            Catch ex As Exception
                Call Z_PRINT("ReadPlateData() TRAP ERROR = " & ex.Message & vbCrLf)
            End Try
        End Sub
#End Region

#Region "ﾒｲﾝﾌｫｰﾑから抵抗ﾃﾞｰﾀを読み込む"
        ''' <summary>ﾒｲﾝﾌｫｰﾑから抵抗ﾃﾞｰﾀを読み込む</summary>
        Private Sub ReadResistorData()
            Try
                Call CopyResistorDataArray(stPLT, W_REG, stREG)
            Catch ex As Exception
                Call Z_PRINT("ReadResistorData() TRAP ERROR = " & ex.Message & vbCrLf)
            End Try
        End Sub
#End Region


#If cOSCILLATORcFLcUSE Then
#Region "ﾒｲﾝﾌｫｰﾑからFL加工条件を読み込む(表示のみ)"
        ''' <summary>ﾒｲﾝﾌｫｰﾑからFL加工条件を読み込む(ｶｯﾄﾀﾌﾞでも使用する)</summary>

        Friend Sub ReadFlConditionData()
            With W_FLCND
                For i As Integer = 0 To (MAX_BANK_NUM - 1) Step 1
                    .Curr = DirectCast(stCND.Curr.Clone, Integer()) ' 電流値
                    .Freq = DirectCast(stCND.Freq.Clone, Double())  ' Qﾚｰﾄ
                    .Steg = DirectCast(stCND.Steg.Clone, Integer()) ' STEG本数
                Next i
            End With

        End Sub
#End Region
#End If
#End Region

#Region "各ﾀﾌﾞをﾚｲｱｳﾄする"
        ''' <summary>各ﾀﾌﾞをﾚｲｱｳﾄする</summary>
        Private Sub LayoutTab()
            Dim tabPages() As TabPage = New TabPage(TAB_COUNT) {}
            Try
                For i As Integer = 0 To (tabPages.Length - 1) Step 1
                    tabPages(i) = New TabPage
                    Select Case (i)
                        Case 0 ' ユーザタブ
                            m_tabUser = New tabUser(Me, i)
                            tabPages(i).Controls.Add(Me.m_tabUser)
                            tabPages(i).Text = m_tabUser.TAB_NAME
                        Case 1 ' ｼｽﾃﾑﾀﾌﾞ
                            m_tabSystem = New tabSystem(Me, i)
                            tabPages(i).Controls.Add(Me.m_tabSystem)
                            tabPages(i).Text = m_tabSystem.TAB_NAME
                        Case 2 ' 抵抗ﾀﾌﾞ
                            m_tabResistor = New tabResistor(Me, i)
                            tabPages(i).Controls.Add(Me.m_tabResistor)
                            tabPages(i).Text = m_tabResistor.TAB_NAME
                        Case 3 ' ｶｯﾄﾀﾌﾞ
                            m_tabCut = New tabCut(Me, i)
                            tabPages(i).Controls.Add(Me.m_tabCut)
                            tabPages(i).Text = m_tabCut.TAB_NAME
#If _PTN_TAB Then
                        Case 4 ' ﾊﾟﾀｰﾝ登録ﾀﾌﾞ
                            m_tabPattern = New tabPattern(Me, i)
                            tabPages(i).Controls.Add(Me.m_tabPattern)
                            tabPages(i).Text = m_tabPattern.TAB_NAME
#End If

#If _PTN_TAB AndAlso _GPIB_TAB Then
                        Case 5 ' GP-IBﾀﾌﾞ
                            m_tabGPIB = New tabGpib(Me, i)
                            tabPages(i).Controls.Add(Me.m_tabGPIB)
                            tabPages(i).Text = m_tabGPIB.TAB_NAME
#ElseIf _GPIB_TAB Then
                        Case 4 ' GP-IBﾀﾌﾞ
                            m_tabGPIB = New tabGpib(Me, i)
                            tabPages(i).Controls.Add(Me.m_tabGPIB)
                            tabPages(i).Text = m_tabGPIB.TAB_NAME
#End If
                        Case 6 ' ユーザタブ                                      'V2.0.0.0②
                            m_tabUser2 = New tabUser2(Me, i)                    'V2.0.0.0②
                            tabPages(i).Controls.Add(Me.m_tabUser2)             'V2.0.0.0②
                            tabPages(i).Text = m_tabUser2.TAB_NAME              'V2.0.0.0②
                        Case Else
                            Throw New Exception("Case " & i & ": Nothing")
                    End Select
                    MTab.TabPages.Add(tabPages(i))
                Next i

            Catch ex As Exception
                Call MsgBox_Exception(ex.Message)
            End Try

        End Sub
#End Region

#Region "ﾃﾞｰﾀ更新"
        '===============================================================================
        '【機　能】トリミングデータ更新処理
        '【引　数】なし
        '【戻り値】なし
        '===============================================================================
        Private Sub DataUpdate()
            Try
                ' ユーザデータ更新
                Call WriteUserData()

                ' ｼｽﾃﾑﾃﾞｰﾀ更新
                Call WritePlateData()

                ' ﾊﾟﾜｰ調整ﾃﾞｰﾀ更新
                '###1040③                stLASER.intQR = W_LASER.intQR ' Qﾚｰﾄ (x100Hz)(0.1KHz)
                '###1040③                stLASER.dblspecPower = W_LASER.dblspecPower ' 設定ﾊﾟﾜｰ[W]
                stLASER = W_LASER           '###1040③

                ' 抵抗･ｶｯﾄﾃﾞｰﾀ更新
                Call WriteResistorData()

                ' ﾊﾟﾀｰﾝ登録ﾃﾞｰﾀ(ｶｯﾄ位置補正用)更新
                stPTN = DirectCast(W_PTN.Clone, Ptn_Info())

                ' θ補正ﾃﾞｰﾀ(XYθ補正用)更新
                stThta = W_THE

                ' TODO: この処理をおこなう必要があるのか確認する
                '' GPIB更新なら旧設定の装置の電圧をOFFする
                'If (FlgUpdGPIB = 1) Then
                '    '        r = V_Off()                                    ' DC電源装置 電圧OFF処理
                'End If

                ' GP-IBﾃﾞｰﾀ
                stGPIB = DirectCast(W_GPIB.Clone, GPIB_DATA())

            Catch ex As Exception
                Call MsgBox_Exception(ex.Message)
            End Try

        End Sub
#End Region

#Region "メインフォームのユーザデータに書き込む"
        '''=========================================================================
        ''' <summary>メインフォームのユーザデータに書き込む</summary>
        '''=========================================================================
        Private Sub WriteUserData()
            'V2.0.0.0⑪↓
            Try
                stUserData = W_stUserData
                stUserData.iResUnit = DirectCast(W_stUserData.iResUnit.Clone(), Integer())              ' 温度センサー 抵抗レンジ 1:Ω, 2:KΩ
                stUserData.dNomCalcCoff = DirectCast(W_stUserData.dNomCalcCoff.Clone(), Double())       ' 温度センサー 抵抗レンジ 1:Ω, 2:KΩ
                stUserData.dTargetCoff = DirectCast(W_stUserData.dTargetCoff.Clone(), Double())         ' 温度センサー 抵抗レンジ 1:Ω, 2:KΩ
                stUserData.dTargetCoffJudge = DirectCast(W_stUserData.dTargetCoffJudge.Clone(), Double())   ' 目標値算出係数 'V2.1.0.0③
                stUserData.iChangeSpeed = DirectCast(W_stUserData.iChangeSpeed.Clone(), Integer())      ' 温度センサー 抵抗レンジ 1:Ω, 2:KΩ
            Catch ex As Exception
                Call MsgBox_Exception(ex.Message)
            End Try
            'V2.0.0.0⑪↑

            'With stUserData
            '    .iTrimType = W_stUserData.iTrimType                   ' 製品種別
            '    .sLotNumber = W_stUserData.sLotNumber                 ' ロット番号
            '    .sOperator = W_stUserData.sOperator                   ' オペレータ名
            '    .sPatternNo = W_stUserData.sPatternNo                 ' パターンＮｏ．
            '    .sProgramNo = W_stUserData.sProgramNo                 ' プログラムＮｏ．
            '    .iTrimSpeed = W_stUserData.iTrimSpeed                 ' トリミング速度
            '    .iLotChange = W_stUserData.iLotChange                 ' ロット終了条件
            '    .lLotEndSL = W_stUserData.lLotEndSL                   ' ロット処理枚数
            '    .lCutHosei = W_stUserData.lCutHosei                   ' カット位置補正頻度
            '    .lPrintRes = W_stUserData.lPrintRes                   ' ロット終了時印刷素子数
            '    .iTempResUnit = W_stUserData.iTempResUnit             ' 温度センサー 抵抗レンジ 1:Ω, 2:KΩ
            '    .iTempTemp = W_stUserData.iTempTemp                   ' 参照温度	１：０℃ または ２：２５℃
            '    .dStandardRes0 = W_stUserData.dStandardRes0           ' 標準抵抗値	０℃ 0.01～100M
            '    .dStandardRes25 = W_stUserData.dStandardRes25         ' 標準抵抗値	２５℃ 0.01～100M
            '    .dResTempCoff = W_stUserData.dResTempCoff             ' 抵抗温度係数
            '    .dFinalLimitHigh = W_stUserData.dFinalLimitHigh       ' ファイナルリミット　Hight[%]
            '    .dFinalLimitLow = W_stUserData.dFinalLimitLow         ' ファイナルリミット　Lo[%]
            '    .dRelativeHigh = W_stUserData.dRelativeHigh           ' 相対値リミット　Hight[%]
            '    .dRelativeLow = W_stUserData.dRelativeLow             ' 相対値リミット　Lo[%]
            '    .intClampVacume = W_stUserData.intClampVacume       'V2.0.0.0⑬ クランプと吸着の有り無し
            '    .Initialize()
            '    For rn As Integer = 1 To MAX_RES_USER Step 1
            '        .iResUnit(rn) = W_stUserData.iResUnit(rn)           ' 温度センサー 抵抗レンジ 1:Ω, 2:KΩ
            '        .dNomCalcCoff(rn) = W_stUserData.dNomCalcCoff(rn)   ' 補正値（ノミナル値算出係数）
            '        .dTargetCoff(rn) = W_stUserData.dTargetCoff(rn)     ' 目標値算出係数
            '        .iChangeSpeed(rn) = W_stUserData.iChangeSpeed(rn)   ' 測定速度を変更するカットNo.
            '    Next rn


            'End With ' W_stUserData


        End Sub
#End Region

#Region "ﾒｲﾝﾌｫｰﾑへのﾃﾞｰﾀ書き込み"

#Region "ﾒｲﾝﾌｫｰﾑのﾌﾟﾚｰﾄﾃﾞｰﾀに書き込む"
        ''' <summary>ﾒｲﾝﾌｫｰﾑのﾌﾟﾚｰﾄﾃﾞｰﾀに書き込む</summary>
        Private Sub WritePlateData()
            Try
                stPLT = W_PLT
            Catch ex As Exception
                Call MsgBox_Exception(ex.Message)
            End Try
        End Sub
#End Region

#Region "ﾒｲﾝﾌｫｰﾑの抵抗ﾃﾞｰﾀに書き込む"
        ''' <summary>ﾒｲﾝﾌｫｰﾑの抵抗ﾃﾞｰﾀに書き込む</summary>
        Private Sub WriteResistorData()
            Try
                Call CopyResistorDataArray(W_PLT, stREG, W_REG)
            Catch ex As Exception
                Call MsgBox_Exception(ex.Message)
            End Try
        End Sub
#End Region

#End Region

#Region "例外発生時に表示するﾒｯｾｰｼﾞﾎﾞｯｸｽ"
        ''' <summary>例外発生時に表示するﾒｯｾｰｼﾞﾎﾞｯｸｽ</summary>
        Protected Sub MsgBox_Exception(ByRef exMsg As String)
            Dim st As New StackTrace
            Dim msg As String
            Try
                ' GetFrame(0)=GetMethod, GetFrame(1)=CallerMethod
                msg = st.GetFrame(1).GetMethod.Name & "() TRAP ERROR = " & exMsg
                Call MsgBox(Me.Name & "." & msg, DirectCast( _
                            MsgBoxStyle.OkOnly + _
                            MsgBoxStyle.Critical, MsgBoxStyle), _
                            My.Application.Info.Title)
            Catch ex As Exception
                Call MsgBox(Me.Name & "." & "MsgBox_Exception() TRAP ERROR = " & ex.Message, _
                            DirectCast( _
                            MsgBoxStyle.OkOnly + _
                            MsgBoxStyle.Critical, MsgBoxStyle), _
                            My.Application.Info.Title)
            End Try

        End Sub
#End Region

#Region "ｿﾌﾄｳｪｱｷｰﾎﾞｰﾄﾞ起動"
        Private Sub CmndKey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmndKey.Click
            Try
                ' すでに起動中ならNOP((注)アプリ名に拡張子は含めない)
                If Process.GetProcessesByName(strProc).Length >= 1 Then
                    Exit Sub
                End If

                Call StartSoftwareKeyBoard(procHandle1)      ' ソフトウェアキーボードを起動する

            Catch ex As Exception
                Call MsgBox_Exception(ex.Message)
            End Try

        End Sub
#End Region

#Region "データ編集終了時のボタン内容"
        ''' <summary>
        ''' データ編集終了時のボタン内容　'V2.2.1.6①
        ''' </summary>
        ''' <returns></returns>
        Public Function GetResult() As Integer

            Try

                GetResult = exitflg

            Catch ex As Exception

            End Try


        End Function
#End Region

#Region "ｲﾍﾞﾝﾄ"
#Region "ﾌｫｰﾑﾛｰﾄﾞ"
        '===============================================================================
        '【機　能】 Form Load時処理
        '【引　数】 なし
        '【戻り値】 なし
        '===============================================================================
        Private Sub frmEdit_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
            Try
                procHandle1 = New Process

                ' ﾃﾞｰﾀﾌｧｲﾙ名他設定
                LblFPATH.Text = gsDataFileName
                LblGuid.Text = "データ確定：ＥＮＴＥＲキー" & vbCrLf & _
                                "【テキストボックス】次項目移動：↓ キー,  前項目移動：↑ キー" & vbCrLf & _
                                "【 コンボボックス 】次項目移動：→ キー,  前項目移動：← キー,  項目選択：↑↓ キー"

                MTab.SelectedIndex = 0 ' ﾀﾌﾞ番号 = ｼｽﾃﾑ

            Catch ex As Exception
                Call MsgBox_Exception(ex.Message)
            End Try

        End Sub
#End Region

#Region "ﾀﾌﾞ選択変更"
        '===============================================================================
        '【機　能】 タブクリック時の処理
        '【引　数】 PreviousTab(INP) : 前タブ番号(0 ORG)
        '【戻り値】 なし
        '===============================================================================
        Private Sub MTab_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MTab.SelectedIndexChanged
            '---------------------------------------------------------------------------
            '   ｸﾘｯｸされたﾀﾌﾞに対応するﾃﾞｰﾀを画面項目に設定する
            '---------------------------------------------------------------------------
            If (False = flgChk) Then ' ﾃﾞｰﾀﾁｪｯｸ以外 ?
                Try
                    Select Case (MTab.SelectedIndex) ' ﾀﾌﾞ番号
                        Case 0 ' ユーザタブ
                            m_tabUser.FIRST_CONTROL.Select()      ' ﾌﾞﾛｯｸｻｲｽﾞX
                        Case 1 ' ｼｽﾃﾑﾀﾌﾞ
                            m_tabSystem.FIRST_CONTROL.Select()      ' ﾌﾞﾛｯｸｻｲｽﾞX
                        Case 2 ' 抵抗ﾀﾌﾞ
                            m_tabResistor.FIRST_CONTROL.Select()    ' 抵抗番号
                        Case 3 ' ｶｯﾄﾀﾌﾞ
                            m_tabCut.FIRST_CONTROL.Select()         ' 抵抗番号
#If _PTN_TAB Then
                        Case 4 ' ﾊﾟﾀｰﾝ登録ﾀﾌﾞ
                            m_tabPattern.FIRST_CONTROL.Select()     ' ﾊﾟﾀｰﾝ認識
#End If

#If _PTN_TAB AndAlso _GPIB_TAB Then
                        Case 5 ' GP-IBﾀﾌﾞ
                            m_tabGPIB.FIRST_CONTROL.Select()        ' 登録番号
#ElseIf _GPIB_TAB Then
                        Case 4 ' GP-IBﾀﾌﾞ
                            m_tabGPIB.FIRST_CONTROL.Select()        ' 登録番号
#End If
                        Case 6 ' ユーザタブ2                         ' V2.0.0.0②
                            m_tabUser2.FIRST_CONTROL.Select()        ' 電圧 V2.0.0.0②
                        Case Else
                            Throw New Exception("Case " & MTab.SelectedIndex & ": Nothing")
                    End Select

                Catch ex As Exception
                    Call MsgBox_Exception(ex.Message)
                End Try
            End If

        End Sub
#End Region

#Region "OKﾎﾞﾀﾝ押下時処理"
        '''=========================================================================
        '''<summary>ＯＫボタン押下時処理</summary>
        '''<remarks></remarks>
        '''=========================================================================
        Private Sub CmndOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmndOK.Click
            Dim ret As Integer = 1
            Try
                'Cursor.Current = Cursors.WaitCursor             ' ｶｰｿﾙを砂時計にする
                Me.Enabled = False
                flgChk = True                                   ' ﾃﾞｰﾀﾁｪｯｸ中ﾌﾗｸﾞ = 1(ﾁｪｯｸ中)

                exitflg = DialogResult.OK                       'V2.2.1.6①
                ' すでに起動中ならNOP((注)ｱﾌﾟﾘ名に拡張子は含めない)
                If (1 <= Process.GetProcessesByName(strProc).Length) Then
                    Call EndSoftwareKeyBoard(procHandle1)        ' ｿﾌﾄｳｪｱｷｰﾎﾞｰﾄﾞを終了する
                End If

                '--------------------------------------------------------------------------
                '   確認ﾒｯｾｰｼﾞを表示する
                '--------------------------------------------------------------------------
                Dim strMsg As String = "トリミングデータを更新します。よろしいですか？"
                ret = MsgBox(strMsg, DirectCast( _
                            MsgBoxStyle.OkCancel + _
                            MsgBoxStyle.Information, MsgBoxStyle), _
                            My.Application.Info.Title)
                If (ret = MsgBoxResult.Cancel) Then ' Cancel(RESETｷｰ) ?
                    flgClose = False
                    Exit Sub
                End If

                '--------------------------------------------------------------------------
                '   全ﾀﾌﾞの全項目のﾃﾞｰﾀをﾁｪｯｸする
                '--------------------------------------------------------------------------
                ' ユーザデータチェック
                ret = m_tabUser.CheckAllTextData()
                If (0 <> ret) Then Exit Try

                ' ｼｽﾃﾑﾃﾞｰﾀﾁｪｯｸ
                ret = m_tabSystem.CheckAllTextData()
                If (0 <> ret) Then Exit Try

                ' 抵抗ﾃﾞｰﾀﾁｪｯｸ
                ret = m_tabResistor.CheckAllTextData()
                If (0 <> ret) Then Exit Try

                ' ｶｯﾄﾃﾞｰﾀﾁｪｯｸ
                ret = m_tabCut.CheckAllTextData()
                If (0 <> ret) Then Exit Try

#If _PTN_TAB Then
                ' ﾊﾟﾀｰﾝ登録ﾃﾞｰﾀﾁｪｯｸ
                ret = m_tabPattern.CheckAllTextData()
                If (0 <> ret) Then Exit Try
#End If

#If _GPIB_TAB Then
                ' GP-IBﾃﾞｰﾀﾁｪｯｸ
                ret = m_tabGPIB.CheckAllTextData()
                If (0 <> ret) Then Exit Try
#End If
                'V2.0.0.0②↓
                ret = m_tabUser2.CheckAllTextData()
                If (0 <> ret) Then Exit Try
                'V2.0.0.0②↑

                '--------------------------------------------------------------------------
                '   ﾃﾞｰﾀ更新処理
                '--------------------------------------------------------------------------
                Call DataUpdate()                               ' トリミングデータ更新
                FlgUpd = Convert.ToInt16(TriState.True)         ' データ更新 Flag ON

            Catch ex As Exception
                Call MsgBox_Exception(ex.Message)
                flgClose = False
            Finally
                flgChk = False                                  ' ﾃﾞｰﾀﾁｪｯｸ中ﾌﾗｸﾞ = 0(ﾁｪｯｸ中でない)
                Me.Enabled = True
                'Cursor.Current = Cursors.Default                ' ｶｰｿﾙを矢印に戻す
            End Try

            If (0 = ret) Then
                flgClose = True
                Me.Close()
            End If

        End Sub
#End Region

#Region "Cancelﾎﾞﾀﾝ押下時処理"
        '''=========================================================================
        '''<summary>Ｃａｎｃｅｌボタン押下時処理</summary>
        '''<remarks></remarks>
        '''=========================================================================
        Private Sub CmndCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmndCancel.Click
            Dim strMsg As String
            Dim ret As Integer
            Try
                Me.Enabled = False

                exitflg = DialogResult.Cancel                        'V2.2.1.6①
                ' すでに起動中ならNOP((注)アプリ名に拡張子は含めない)
                If (1 <= Process.GetProcessesByName(strProc).Length) Then
                    Call EndSoftwareKeyBoard(procHandle1)        ' ソフトウェアキーボードを終了する
                End If

                ' 確認ﾒｯｾｰｼﾞを表示する
                strMsg = "編集中のデータを破棄してよろしいですか？"
                ret = MsgBox(strMsg, DirectCast( _
                            MsgBoxStyle.OkCancel + _
                            MsgBoxStyle.Exclamation + _
                            MsgBoxStyle.DefaultButton2, MsgBoxStyle), _
                        My.Application.Info.Title)
                If (ret = MsgBoxResult.Cancel) Then ' Cancel(RESETｷｰ) ?
                    flgClose = False
                    Exit Sub
                End If

            Catch ex As Exception
                Call MsgBox_Exception(ex.Message)
                flgClose = False
            Finally
                Me.Enabled = True
            End Try

            FlgCan = Convert.ToInt16(TriState.True)
            flgClose = True
            Me.Close()

        End Sub
#End Region

#Region "ﾌｫｰﾑが閉じられる時の処理"
        ''' <summary>ﾌｫｰﾑが閉じられる時の処理</summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Private Sub frmEdit_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
            e.Cancel = (Not flgClose) ' 意図せずﾌｫｰﾑが閉じられるのを回避する
        End Sub
#End Region
#End Region

    End Class
End Namespace
