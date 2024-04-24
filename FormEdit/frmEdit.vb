'==============================================================================
'
'   DESCRIPTION:    �p�����[�^�ҏW��ʏ���('10.07.22 A.W)
'
'==============================================================================
Option Explicit On
Option Strict On

Imports System
Imports System.Drawing
Imports System.Windows.Forms

#Const _PTN_TAB = True          ' �������� ���ı�ĂŔ�\��
#Const _GPIB_TAB = True         ' GP-IB��� ���ı�ĂŔ�\��

'V1.0.4.3�B�g�p���Ȃ��̂Ŋ֘A�\�[�X���폜 #Const _ARRAY_DATA = False      ' ��ڰ��ް����z�񂩂ǂ���(frmEdit.vb�̂�)

Namespace FormEdit
    Friend Class frmEdit
        Inherits System.Windows.Forms.Form

#Region "�萔��`"
        '---------------------------------------------------------------------------
        '   �萔��`
        '---------------------------------------------------------------------------
        Private m_tabUser As tabUser                            ' ���[�U�^�u
        Private m_tabUser2 As tabUser2                          ' ���[�U�^�u2 'V2.0.0.0�A
        Private m_tabSystem As tabSystem                        ' �������
        Private m_tabResistor As tabResistor                    ' ��R���
        Private m_tabCut As tabCut                              ' ������

#If _PTN_TAB Then
        Private m_tabPattern As tabPattern                      ' ����ݓo�^���
#End If

#If _GPIB_TAB Then
        Private m_tabGPIB As tabGpib                            ' GP-IB���
#End If

        Private Const TAB_COUNT As Integer = (7 - 1)            ' (��ނ̖��� - 1)'V2.0.0.0�A
        'V2.0.0.0�A#If _PTN_TAB And _GPIB_TAB Then
        'V2.0.0.0�A        Private Const TAB_COUNT As Integer = (6 - 1)            ' (��ނ̖��� - 1)
        'V2.0.0.0�A#ElseIf _PTN_TAB Or _GPIB_TAB Then
        'V2.0.0.0�A        Private Const TAB_COUNT As Integer = (5 - 1)            ' (��ނ̖��� - 1)
        'V2.0.0.0�A#Else
        'V2.0.0.0�A        Private Const TAB_COUNT As Integer = (4 - 1)            ' (��ނ̖��� - 1)
        'V2.0.0.0�A#End If

        '---------------------------------------------------------------------------
        '   �ҏW�p�ް���
        '---------------------------------------------------------------------------
        Friend W_stUserData As USER_DATA                        ' ���[�U�f�[�^
        Friend W_PLT As PLATE_DATA                              ' ��ڰ��ް� [ 1 ]ORG
        Friend W_LASER As POWER_DATA                            ' ��ܰ����p�ް�
        Friend W_REG(MAXRNO) As Reg_Info                        ' ��R�ް� [ 1 ]ORG

#If cOSCILLATORcFLcUSE Then
        Friend W_FLCND As TrimCondInfo                          ' FL���H���� [ 0 ]ORG
#End If
        Friend W_PTN(MAXRGN) As Ptn_Info                        ' ����ݓo�^�ް�(��Ĉʒu�␳�p) [ 1 ]ORG
        Friend W_THE As Theta_Info                              ' ����ݓo�^�ް�(XY�ƕ␳�p) [ 1 ]ORG
        Friend W_GPIB(MAXGNO) As GPIB_DATA                      ' GP-IB�ް� [ 1 ]ORG

        '---------------------------------------------------------------------------
        '   ���̑�
        '---------------------------------------------------------------------------
        Private flgChk As Boolean               ' �ް��������׸�(False:�������łȂ�, True:������)
        Private flgClose As Boolean             ' FormClosing����ĂŎg�p����(True:����, False:���Ȃ�)
        Private exitflg As Integer              ' �ҏW��ʂ𔲂���Ƃ��̃{�^���F'V2.2.1.6�@

        Friend giRNO As Integer                 ' �e��ދ��L��������R�ԍ�(1 ORG)
        Friend giCNO As Integer                 ' �e��ދ��L��������Ĕԍ� (1 ORG)
        Friend giGNO As Integer                 ' ������GPIB�ԍ�         (1 ORG)

        Private procHandle1 As Process          ' ��ĳ�����ް�ނ̋N����I���Ŏg�p
        Private strProc As String = "OSK"       ' ��ĳ�����ް�ނ̋N����I���Ŏg�p
#End Region

#Region "̫�т̏�����"
        ''' <summary>̫�я���������</summary>
        Private Sub Form_Initialize_Renamed()
            '---------------------------------------------------------------------------
            '   �ް���ҏW�p�ް���ɐݒ肷��
            '---------------------------------------------------------------------------
            flgChk = False      ' �ް��������׸� = False:�������łȂ�
            flgClose = False    ' FormClosing����ĂŎg�p���� = False:���Ȃ�

            giRNO = 1 ' ��������R�ԍ�     (1 ORG)
            giCNO = 1 ' ��������Ĕԍ�      (1 ORG)
            giGNO = 1 ' ������GP-IB�o�^�ԍ�(1 ORG)

            Try
                ' ���[�U�f�[�^
                Call ReadUserData()

                ' �����ް�
                Call ReadPlateData()

                ' ��ܰ�����ް�
                W_LASER = stLASER

                ' ��R/����ް�
                Call ReadResistorData()

#If cOSCILLATORcFLcUSE Then
                ' FL���H�����ް�(�\���̂�)
                Call ReadFlConditionData()
#End If

                ' ����ݓo�^�ް�(��Ĉʒu�␳�p)
                W_PTN = DirectCast(stPTN.Clone, Ptn_Info())

                ' �ƕ␳�ް�(XY�ƕ␳�p)
                W_THE = stThta

                ' GP-IB�ް�
                W_GPIB = DirectCast(stGPIB.Clone, GPIB_DATA())

                ' �e��ނ�z�u����
                Call LayoutTab()

            Catch ex As Exception
                Call MsgBox_Exception(ex.Message)
            End Try

        End Sub
#End Region

#Region "���C���t�H�[�����烆�[�U�f�[�^��ǂݍ���"

        '''=========================================================================
        ''' <summary>���C���t�H�[�����烆�[�U�f�[�^��ǂݍ���</summary>
        '''=========================================================================
        Private Sub ReadUserData()
            'V2.0.0.0�J��
            Try
                W_stUserData = stUserData
                W_stUserData.iResUnit = DirectCast(stUserData.iResUnit.Clone(), Integer())              ' ���x�Z���T�[ ��R�����W 1:��, 2:K��
                W_stUserData.dNomCalcCoff = DirectCast(stUserData.dNomCalcCoff.Clone(), Double())       ' ���x�Z���T�[ ��R�����W 1:��, 2:K��
                W_stUserData.dTargetCoff = DirectCast(stUserData.dTargetCoff.Clone(), Double())         ' ���x�Z���T�[ ��R�����W 1:��, 2:K��
                W_stUserData.dTargetCoffJudge = DirectCast(stUserData.dTargetCoffJudge.Clone(), Double())   ' �ڕW�l�Z�o�W�� 'V2.1.0.0�B
                W_stUserData.iChangeSpeed = DirectCast(stUserData.iChangeSpeed.Clone(), Integer())      ' ���x�Z���T�[ ��R�����W 1:��, 2:K��
            Catch ex As Exception
                Call MsgBox_Exception(ex.Message)
            End Try
            'V2.0.0.0�J��
            'With W_stUserData
            '    .iTrimType = stUserData.iTrimType                       ' ���i���
            '    .sLotNumber = stUserData.sLotNumber                     ' ���b�g�ԍ�
            '    .sOperator = stUserData.sOperator                       ' �I�y���[�^��
            '    .sPatternNo = stUserData.sPatternNo                     ' �p�^�[���m���D
            '    .sProgramNo = stUserData.sProgramNo                     ' �v���O�����m���D
            '    .iTrimSpeed = stUserData.iTrimSpeed                     ' �g���~���O���x
            '    .iLotChange = stUserData.iLotChange                     ' ���b�g�I������
            '    .lLotEndSL = stUserData.lLotEndSL                       ' ���b�g��������
            '    .lCutHosei = stUserData.lCutHosei                       ' �J�b�g�ʒu�␳�p�x
            '    .lPrintRes = stUserData.lPrintRes                       ' ���b�g�I��������f�q��
            '    .iTempResUnit = stUserData.iTempResUnit                 ' ���x�Z���T�[ ��R�����W 1:��, 2:K��
            '    .iTempTemp = stUserData.iTempTemp                       ' �Q�Ɖ��x	�P�F�O�� �܂��� �Q�F�Q�T��
            '    .dStandardRes0 = stUserData.dStandardRes0               ' �W����R�l �O��	0.01�`100M
            '    .dStandardRes25 = stUserData.dStandardRes25             ' �W����R�l �Q�T��	0.01�`100M
            '    .dResTempCoff = stUserData.dResTempCoff                 ' ��R���x�W��
            '    .dFinalLimitHigh = stUserData.dFinalLimitHigh           ' �t�@�C�i�����~�b�g�@Hight[%]
            '    .dFinalLimitLow = stUserData.dFinalLimitLow             ' �t�@�C�i�����~�b�g�@Lo[%]
            '    .dRelativeHigh = stUserData.dRelativeHigh               ' ���Βl���~�b�g�@Hight[%]
            '    .dRelativeLow = stUserData.dRelativeLow                 ' ���Βl���~�b�g�@Lo[%]
            '    .Initialize()
            '    For rn As Integer = 1 To MAX_RES_USER Step 1
            '        .iResUnit(rn) = stUserData.iResUnit(rn)             ' ���x�Z���T�[ ��R�����W 1:��, 2:K��
            '        .dNomCalcCoff(rn) = stUserData.dNomCalcCoff(rn)     ' �␳�l�i�m�~�i���l�Z�o�W���j
            '        .dTargetCoff(rn) = stUserData.dTargetCoff(rn)       ' �ڕW�l�Z�o�W��
            '        .iChangeSpeed(rn) = stUserData.iChangeSpeed(rn)     ' ���葬�x��ύX����J�b�gNo.
            '    Next rn
            'End With ' W_stUserData


        End Sub
#End Region

#Region "Ҳ�̫�т�����ް��ǂݍ���"

#Region "Ҳ�̫�т�����ڰ��ް���ǂݍ���"
        ''' <summary>Ҳ�̫�т�����ڰ��ް���ǂݍ���</summary>
        Private Sub ReadPlateData()
            Try
                W_PLT = stPLT
            Catch ex As Exception
                Call Z_PRINT("ReadPlateData() TRAP ERROR = " & ex.Message & vbCrLf)
            End Try
        End Sub
#End Region

#Region "Ҳ�̫�т����R�ް���ǂݍ���"
        ''' <summary>Ҳ�̫�т����R�ް���ǂݍ���</summary>
        Private Sub ReadResistorData()
            Try
                Call CopyResistorDataArray(stPLT, W_REG, stREG)
            Catch ex As Exception
                Call Z_PRINT("ReadResistorData() TRAP ERROR = " & ex.Message & vbCrLf)
            End Try
        End Sub
#End Region


#If cOSCILLATORcFLcUSE Then
#Region "Ҳ�̫�т���FL���H������ǂݍ���(�\���̂�)"
        ''' <summary>Ҳ�̫�т���FL���H������ǂݍ���(�����ނł��g�p����)</summary>

        Friend Sub ReadFlConditionData()
            With W_FLCND
                For i As Integer = 0 To (MAX_BANK_NUM - 1) Step 1
                    .Curr = DirectCast(stCND.Curr.Clone, Integer()) ' �d���l
                    .Freq = DirectCast(stCND.Freq.Clone, Double())  ' Qڰ�
                    .Steg = DirectCast(stCND.Steg.Clone, Integer()) ' STEG�{��
                Next i
            End With

        End Sub
#End Region
#End If
#End Region

#Region "�e��ނ�ڲ��Ă���"
        ''' <summary>�e��ނ�ڲ��Ă���</summary>
        Private Sub LayoutTab()
            Dim tabPages() As TabPage = New TabPage(TAB_COUNT) {}
            Try
                For i As Integer = 0 To (tabPages.Length - 1) Step 1
                    tabPages(i) = New TabPage
                    Select Case (i)
                        Case 0 ' ���[�U�^�u
                            m_tabUser = New tabUser(Me, i)
                            tabPages(i).Controls.Add(Me.m_tabUser)
                            tabPages(i).Text = m_tabUser.TAB_NAME
                        Case 1 ' �������
                            m_tabSystem = New tabSystem(Me, i)
                            tabPages(i).Controls.Add(Me.m_tabSystem)
                            tabPages(i).Text = m_tabSystem.TAB_NAME
                        Case 2 ' ��R���
                            m_tabResistor = New tabResistor(Me, i)
                            tabPages(i).Controls.Add(Me.m_tabResistor)
                            tabPages(i).Text = m_tabResistor.TAB_NAME
                        Case 3 ' ������
                            m_tabCut = New tabCut(Me, i)
                            tabPages(i).Controls.Add(Me.m_tabCut)
                            tabPages(i).Text = m_tabCut.TAB_NAME
#If _PTN_TAB Then
                        Case 4 ' ����ݓo�^���
                            m_tabPattern = New tabPattern(Me, i)
                            tabPages(i).Controls.Add(Me.m_tabPattern)
                            tabPages(i).Text = m_tabPattern.TAB_NAME
#End If

#If _PTN_TAB AndAlso _GPIB_TAB Then
                        Case 5 ' GP-IB���
                            m_tabGPIB = New tabGpib(Me, i)
                            tabPages(i).Controls.Add(Me.m_tabGPIB)
                            tabPages(i).Text = m_tabGPIB.TAB_NAME
#ElseIf _GPIB_TAB Then
                        Case 4 ' GP-IB���
                            m_tabGPIB = New tabGpib(Me, i)
                            tabPages(i).Controls.Add(Me.m_tabGPIB)
                            tabPages(i).Text = m_tabGPIB.TAB_NAME
#End If
                        Case 6 ' ���[�U�^�u                                      'V2.0.0.0�A
                            m_tabUser2 = New tabUser2(Me, i)                    'V2.0.0.0�A
                            tabPages(i).Controls.Add(Me.m_tabUser2)             'V2.0.0.0�A
                            tabPages(i).Text = m_tabUser2.TAB_NAME              'V2.0.0.0�A
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

#Region "�ް��X�V"
        '===============================================================================
        '�y�@�@�\�z�g���~���O�f�[�^�X�V����
        '�y���@���z�Ȃ�
        '�y�߂�l�z�Ȃ�
        '===============================================================================
        Private Sub DataUpdate()
            Try
                ' ���[�U�f�[�^�X�V
                Call WriteUserData()

                ' �����ް��X�V
                Call WritePlateData()

                ' ��ܰ�����ް��X�V
                '###1040�B                stLASER.intQR = W_LASER.intQR ' Qڰ� (x100Hz)(0.1KHz)
                '###1040�B                stLASER.dblspecPower = W_LASER.dblspecPower ' �ݒ���ܰ[W]
                stLASER = W_LASER           '###1040�B

                ' ��R�����ް��X�V
                Call WriteResistorData()

                ' ����ݓo�^�ް�(��Ĉʒu�␳�p)�X�V
                stPTN = DirectCast(W_PTN.Clone, Ptn_Info())

                ' �ƕ␳�ް�(XY�ƕ␳�p)�X�V
                stThta = W_THE

                ' TODO: ���̏����������Ȃ��K�v������̂��m�F����
                '' GPIB�X�V�Ȃ狌�ݒ�̑��u�̓d����OFF����
                'If (FlgUpdGPIB = 1) Then
                '    '        r = V_Off()                                    ' DC�d�����u �d��OFF����
                'End If

                ' GP-IB�ް�
                stGPIB = DirectCast(W_GPIB.Clone, GPIB_DATA())

            Catch ex As Exception
                Call MsgBox_Exception(ex.Message)
            End Try

        End Sub
#End Region

#Region "���C���t�H�[���̃��[�U�f�[�^�ɏ�������"
        '''=========================================================================
        ''' <summary>���C���t�H�[���̃��[�U�f�[�^�ɏ�������</summary>
        '''=========================================================================
        Private Sub WriteUserData()
            'V2.0.0.0�J��
            Try
                stUserData = W_stUserData
                stUserData.iResUnit = DirectCast(W_stUserData.iResUnit.Clone(), Integer())              ' ���x�Z���T�[ ��R�����W 1:��, 2:K��
                stUserData.dNomCalcCoff = DirectCast(W_stUserData.dNomCalcCoff.Clone(), Double())       ' ���x�Z���T�[ ��R�����W 1:��, 2:K��
                stUserData.dTargetCoff = DirectCast(W_stUserData.dTargetCoff.Clone(), Double())         ' ���x�Z���T�[ ��R�����W 1:��, 2:K��
                stUserData.dTargetCoffJudge = DirectCast(W_stUserData.dTargetCoffJudge.Clone(), Double())   ' �ڕW�l�Z�o�W�� 'V2.1.0.0�B
                stUserData.iChangeSpeed = DirectCast(W_stUserData.iChangeSpeed.Clone(), Integer())      ' ���x�Z���T�[ ��R�����W 1:��, 2:K��
            Catch ex As Exception
                Call MsgBox_Exception(ex.Message)
            End Try
            'V2.0.0.0�J��

            'With stUserData
            '    .iTrimType = W_stUserData.iTrimType                   ' ���i���
            '    .sLotNumber = W_stUserData.sLotNumber                 ' ���b�g�ԍ�
            '    .sOperator = W_stUserData.sOperator                   ' �I�y���[�^��
            '    .sPatternNo = W_stUserData.sPatternNo                 ' �p�^�[���m���D
            '    .sProgramNo = W_stUserData.sProgramNo                 ' �v���O�����m���D
            '    .iTrimSpeed = W_stUserData.iTrimSpeed                 ' �g���~���O���x
            '    .iLotChange = W_stUserData.iLotChange                 ' ���b�g�I������
            '    .lLotEndSL = W_stUserData.lLotEndSL                   ' ���b�g��������
            '    .lCutHosei = W_stUserData.lCutHosei                   ' �J�b�g�ʒu�␳�p�x
            '    .lPrintRes = W_stUserData.lPrintRes                   ' ���b�g�I��������f�q��
            '    .iTempResUnit = W_stUserData.iTempResUnit             ' ���x�Z���T�[ ��R�����W 1:��, 2:K��
            '    .iTempTemp = W_stUserData.iTempTemp                   ' �Q�Ɖ��x	�P�F�O�� �܂��� �Q�F�Q�T��
            '    .dStandardRes0 = W_stUserData.dStandardRes0           ' �W����R�l	�O�� 0.01�`100M
            '    .dStandardRes25 = W_stUserData.dStandardRes25         ' �W����R�l	�Q�T�� 0.01�`100M
            '    .dResTempCoff = W_stUserData.dResTempCoff             ' ��R���x�W��
            '    .dFinalLimitHigh = W_stUserData.dFinalLimitHigh       ' �t�@�C�i�����~�b�g�@Hight[%]
            '    .dFinalLimitLow = W_stUserData.dFinalLimitLow         ' �t�@�C�i�����~�b�g�@Lo[%]
            '    .dRelativeHigh = W_stUserData.dRelativeHigh           ' ���Βl���~�b�g�@Hight[%]
            '    .dRelativeLow = W_stUserData.dRelativeLow             ' ���Βl���~�b�g�@Lo[%]
            '    .intClampVacume = W_stUserData.intClampVacume       'V2.0.0.0�L �N�����v�Ƌz���̗L�薳��
            '    .Initialize()
            '    For rn As Integer = 1 To MAX_RES_USER Step 1
            '        .iResUnit(rn) = W_stUserData.iResUnit(rn)           ' ���x�Z���T�[ ��R�����W 1:��, 2:K��
            '        .dNomCalcCoff(rn) = W_stUserData.dNomCalcCoff(rn)   ' �␳�l�i�m�~�i���l�Z�o�W���j
            '        .dTargetCoff(rn) = W_stUserData.dTargetCoff(rn)     ' �ڕW�l�Z�o�W��
            '        .iChangeSpeed(rn) = W_stUserData.iChangeSpeed(rn)   ' ���葬�x��ύX����J�b�gNo.
            '    Next rn


            'End With ' W_stUserData


        End Sub
#End Region

#Region "Ҳ�̫�тւ��ް���������"

#Region "Ҳ�̫�т���ڰ��ް��ɏ�������"
        ''' <summary>Ҳ�̫�т���ڰ��ް��ɏ�������</summary>
        Private Sub WritePlateData()
            Try
                stPLT = W_PLT
            Catch ex As Exception
                Call MsgBox_Exception(ex.Message)
            End Try
        End Sub
#End Region

#Region "Ҳ�̫�т̒�R�ް��ɏ�������"
        ''' <summary>Ҳ�̫�т̒�R�ް��ɏ�������</summary>
        Private Sub WriteResistorData()
            Try
                Call CopyResistorDataArray(W_PLT, stREG, W_REG)
            Catch ex As Exception
                Call MsgBox_Exception(ex.Message)
            End Try
        End Sub
#End Region

#End Region

#Region "��O�������ɕ\������ү�����ޯ��"
        ''' <summary>��O�������ɕ\������ү�����ޯ��</summary>
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

#Region "��ĳ�����ް�ދN��"
        Private Sub CmndKey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmndKey.Click
            Try
                ' ���łɋN�����Ȃ�NOP((��)�A�v�����Ɋg���q�͊܂߂Ȃ�)
                If Process.GetProcessesByName(strProc).Length >= 1 Then
                    Exit Sub
                End If

                Call StartSoftwareKeyBoard(procHandle1)      ' �\�t�g�E�F�A�L�[�{�[�h���N������

            Catch ex As Exception
                Call MsgBox_Exception(ex.Message)
            End Try

        End Sub
#End Region

#Region "�f�[�^�ҏW�I�����̃{�^�����e"
        ''' <summary>
        ''' �f�[�^�ҏW�I�����̃{�^�����e�@'V2.2.1.6�@
        ''' </summary>
        ''' <returns></returns>
        Public Function GetResult() As Integer

            Try

                GetResult = exitflg

            Catch ex As Exception

            End Try


        End Function
#End Region

#Region "�����"
#Region "̫��۰��"
        '===============================================================================
        '�y�@�@�\�z Form Load������
        '�y���@���z �Ȃ�
        '�y�߂�l�z �Ȃ�
        '===============================================================================
        Private Sub frmEdit_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
            Try
                procHandle1 = New Process

                ' �ް�̧�ٖ����ݒ�
                LblFPATH.Text = gsDataFileName
                LblGuid.Text = "�f�[�^�m��F�d�m�s�d�q�L�[" & vbCrLf & _
                                "�y�e�L�X�g�{�b�N�X�z�����ڈړ��F�� �L�[,  �O���ڈړ��F�� �L�[" & vbCrLf & _
                                "�y �R���{�{�b�N�X �z�����ڈړ��F�� �L�[,  �O���ڈړ��F�� �L�[,  ���ڑI���F���� �L�["

                MTab.SelectedIndex = 0 ' ��ޔԍ� = ����

            Catch ex As Exception
                Call MsgBox_Exception(ex.Message)
            End Try

        End Sub
#End Region

#Region "��ޑI��ύX"
        '===============================================================================
        '�y�@�@�\�z �^�u�N���b�N���̏���
        '�y���@���z PreviousTab(INP) : �O�^�u�ԍ�(0 ORG)
        '�y�߂�l�z �Ȃ�
        '===============================================================================
        Private Sub MTab_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MTab.SelectedIndexChanged
            '---------------------------------------------------------------------------
            '   �د����ꂽ��ނɑΉ������ް�����ʍ��ڂɐݒ肷��
            '---------------------------------------------------------------------------
            If (False = flgChk) Then ' �ް������ȊO ?
                Try
                    Select Case (MTab.SelectedIndex) ' ��ޔԍ�
                        Case 0 ' ���[�U�^�u
                            m_tabUser.FIRST_CONTROL.Select()      ' ��ۯ�����X
                        Case 1 ' �������
                            m_tabSystem.FIRST_CONTROL.Select()      ' ��ۯ�����X
                        Case 2 ' ��R���
                            m_tabResistor.FIRST_CONTROL.Select()    ' ��R�ԍ�
                        Case 3 ' ������
                            m_tabCut.FIRST_CONTROL.Select()         ' ��R�ԍ�
#If _PTN_TAB Then
                        Case 4 ' ����ݓo�^���
                            m_tabPattern.FIRST_CONTROL.Select()     ' ����ݔF��
#End If

#If _PTN_TAB AndAlso _GPIB_TAB Then
                        Case 5 ' GP-IB���
                            m_tabGPIB.FIRST_CONTROL.Select()        ' �o�^�ԍ�
#ElseIf _GPIB_TAB Then
                        Case 4 ' GP-IB���
                            m_tabGPIB.FIRST_CONTROL.Select()        ' �o�^�ԍ�
#End If
                        Case 6 ' ���[�U�^�u2                         ' V2.0.0.0�A
                            m_tabUser2.FIRST_CONTROL.Select()        ' �d�� V2.0.0.0�A
                        Case Else
                            Throw New Exception("Case " & MTab.SelectedIndex & ": Nothing")
                    End Select

                Catch ex As Exception
                    Call MsgBox_Exception(ex.Message)
                End Try
            End If

        End Sub
#End Region

#Region "OK���݉���������"
        '''=========================================================================
        '''<summary>�n�j�{�^������������</summary>
        '''<remarks></remarks>
        '''=========================================================================
        Private Sub CmndOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmndOK.Click
            Dim ret As Integer = 1
            Try
                'Cursor.Current = Cursors.WaitCursor             ' ���ق������v�ɂ���
                Me.Enabled = False
                flgChk = True                                   ' �ް��������׸� = 1(������)

                exitflg = DialogResult.OK                       'V2.2.1.6�@
                ' ���łɋN�����Ȃ�NOP((��)���ؖ��Ɋg���q�͊܂߂Ȃ�)
                If (1 <= Process.GetProcessesByName(strProc).Length) Then
                    Call EndSoftwareKeyBoard(procHandle1)        ' ��ĳ�����ް�ނ��I������
                End If

                '--------------------------------------------------------------------------
                '   �m�Fү���ނ�\������
                '--------------------------------------------------------------------------
                Dim strMsg As String = "�g���~���O�f�[�^���X�V���܂��B��낵���ł����H"
                ret = MsgBox(strMsg, DirectCast( _
                            MsgBoxStyle.OkCancel + _
                            MsgBoxStyle.Information, MsgBoxStyle), _
                            My.Application.Info.Title)
                If (ret = MsgBoxResult.Cancel) Then ' Cancel(RESET��) ?
                    flgClose = False
                    Exit Sub
                End If

                '--------------------------------------------------------------------------
                '   �S��ނ̑S���ڂ��ް�����������
                '--------------------------------------------------------------------------
                ' ���[�U�f�[�^�`�F�b�N
                ret = m_tabUser.CheckAllTextData()
                If (0 <> ret) Then Exit Try

                ' �����ް�����
                ret = m_tabSystem.CheckAllTextData()
                If (0 <> ret) Then Exit Try

                ' ��R�ް�����
                ret = m_tabResistor.CheckAllTextData()
                If (0 <> ret) Then Exit Try

                ' ����ް�����
                ret = m_tabCut.CheckAllTextData()
                If (0 <> ret) Then Exit Try

#If _PTN_TAB Then
                ' ����ݓo�^�ް�����
                ret = m_tabPattern.CheckAllTextData()
                If (0 <> ret) Then Exit Try
#End If

#If _GPIB_TAB Then
                ' GP-IB�ް�����
                ret = m_tabGPIB.CheckAllTextData()
                If (0 <> ret) Then Exit Try
#End If
                'V2.0.0.0�A��
                ret = m_tabUser2.CheckAllTextData()
                If (0 <> ret) Then Exit Try
                'V2.0.0.0�A��

                '--------------------------------------------------------------------------
                '   �ް��X�V����
                '--------------------------------------------------------------------------
                Call DataUpdate()                               ' �g���~���O�f�[�^�X�V
                FlgUpd = Convert.ToInt16(TriState.True)         ' �f�[�^�X�V Flag ON

            Catch ex As Exception
                Call MsgBox_Exception(ex.Message)
                flgClose = False
            Finally
                flgChk = False                                  ' �ް��������׸� = 0(�������łȂ�)
                Me.Enabled = True
                'Cursor.Current = Cursors.Default                ' ���ق���ɖ߂�
            End Try

            If (0 = ret) Then
                flgClose = True
                Me.Close()
            End If

        End Sub
#End Region

#Region "Cancel���݉���������"
        '''=========================================================================
        '''<summary>�b�����������{�^������������</summary>
        '''<remarks></remarks>
        '''=========================================================================
        Private Sub CmndCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmndCancel.Click
            Dim strMsg As String
            Dim ret As Integer
            Try
                Me.Enabled = False

                exitflg = DialogResult.Cancel                        'V2.2.1.6�@
                ' ���łɋN�����Ȃ�NOP((��)�A�v�����Ɋg���q�͊܂߂Ȃ�)
                If (1 <= Process.GetProcessesByName(strProc).Length) Then
                    Call EndSoftwareKeyBoard(procHandle1)        ' �\�t�g�E�F�A�L�[�{�[�h���I������
                End If

                ' �m�Fү���ނ�\������
                strMsg = "�ҏW���̃f�[�^��j�����Ă�낵���ł����H"
                ret = MsgBox(strMsg, DirectCast( _
                            MsgBoxStyle.OkCancel + _
                            MsgBoxStyle.Exclamation + _
                            MsgBoxStyle.DefaultButton2, MsgBoxStyle), _
                        My.Application.Info.Title)
                If (ret = MsgBoxResult.Cancel) Then ' Cancel(RESET��) ?
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

#Region "̫�т������鎞�̏���"
        ''' <summary>̫�т������鎞�̏���</summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Private Sub frmEdit_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
            e.Cancel = (Not flgClose) ' �Ӑ}����̫�т�������̂��������
        End Sub
#End Region
#End Region

    End Class
End Namespace
