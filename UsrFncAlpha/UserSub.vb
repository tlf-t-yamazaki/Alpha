'==============================================================================
'   Description : ���[�U�v���O�����p�ŗL�t�@���N�V����
'
'�@ 2012/11/16 First Written by N.Arata(OLFT)
'
'==============================================================================
Option Strict Off
Option Explicit On

Imports System
Imports System.Drawing.Printing
Imports System.IO
Imports System.Text
Imports LaserFront.Trimmer.DefWin32Fnc
Imports LaserFront.Trimmer.DllLaserTeach.ctl_LaserTeach

Module UserSub
    Private bStartCheck As Boolean          ' �f�[�^�ݒ�m�F���K�v�Ȏ��@True�Ƃ���B
    Private dInitialResValue As Double      ' ��������l
    Private dStandardResValue As Double     ' �W����R����l
    Private lResCounterForPrinter As Long   ' ����p�f�q�J�E���^

    Private intTMM_Save As Integer          ' �ۑ��p�@���[�h(0:����(�R���p���[�^��ϕ����[�h), 1:�����x(�ϕ����[�h))
    Private intMType_Save As Integer        ' �ۑ��p�@������(0=��������, 1=�O������)
    Private dTRV As Double                  ' �ڕW��R�l
    Private bOkJudge As Boolean             ' �f�q�P�ʂ�NG����
    Private bSkip As Boolean                ' �l�b�g���[�N��R�̃X�L�b�v
    Private sResistorPrintData(MAX_RES_USER) As String    ' �l�b�g���[�N��R�̎��̏o�̓f�[�^

    Public Printer As New cPrintDocument    ' ��������޼ު��
    '===============================================================================
    ' ����p�f�[�^�̈�
    '===============================================================================
    Private Const cTRIM_PRINT_DATA_HEAD As String = "C:\TRIMDATA\PRINTDATA\TRIM_PRINT_DATA_HEAD.TXT"
    Private Const cTRIM_PRINT_DATA_RES As String = "C:\TRIMDATA\PRINTDATA\TRIM_PRINT_DATA_RES.TXT"
    Private Const cTRIM_PRINT_DATA_PLATE As String = "C:\TRIMDATA\PRINTDATA\TRIM_PRINT_DATA_PLATE.TXT"
    Private Const cTRIM_PRINT_DATA_END As String = "C:\TRIMDATA\PRINTDATA\TRIM_PRINT_DATA_END.TXT"

    '''===============================================================================
    ''' <summary>
    ''' ����p�f�q�J�E���^�̃��Z�b�g
    ''' </summary>
    ''' <remarks></remarks>
    '''===============================================================================
    Public Sub ResetlResCounterForPrinter()
        lResCounterForPrinter = 0
    End Sub

    '===============================================================================
    '�y�@�@�\�z ��R���x�W���̎Z�o
    '�y���@���z �X�^���_�[�h��R�l�O���A�X�^���_�[�h��R�l�Q�T��
    '�y�߂�l�z ��R���x�W��
    '===============================================================================
    Public Function GetResTempCoff(ByVal dStandardRes0 As Double, ByVal dStandardRes25 As Double) As Double
        GetResTempCoff = (dStandardRes25 - dStandardRes0) / dStandardRes0 * 10.0 ^ 6 / 25.0
    End Function

    '===============================================================================
    '�y�@�@�\�z ���[�U�ݒ��ʊm�F
    '�y���@���z true , false
    '�y�߂�l�z ����
    '===============================================================================
    Public Sub SetStartCheckStatus(ByVal bCheck As Boolean)
        bStartCheck = bCheck
    End Sub
    Public Function GetStartCheckStatus() As Boolean
        Return (bStartCheck)
    End Function

    '===============================================================================
    '�y�@�@�\�z �W����R�l�̒�R�l�Z�o
    '�y���@���z ��R�ԍ�,�J�b�g�ԍ�
    '�y�߂�l�z �ڕW��R�l
    '===============================================================================
    Public Function CalcStandardResistanceValue() As Double

        'V2.0.0.0�J        CalcStandardResistanceValue = stUserData.dStandardRes25

        'V2.0.0.0�J��
        'STD��R�l(25��)=STD(0��)��R�l�~(1+���~25+���~25^2)
        Dim dAlpha As Double = stUserData.dAlpha / 10.0 ^ 6
        Dim dBeta As Double = stUserData.dBeta / 10.0 ^ 6
        Dim dTemp As Double = 25.0
        CalcStandardResistanceValue = stUserData.dTemperatura0 * (1 + dAlpha * dTemp + dBeta * dTemp ^ 2)
        DebugLogOut("STD��R�l(25��)[" & CalcStandardResistanceValue.ToString & "]=STD(0��)��R�l[" & stUserData.dTemperatura0.ToString & "]*(1+��[" & dAlpha.ToString & "]*[" & dTemp.ToString & "]+��[" & dBeta.ToString & "]*[" & dTemp.ToString & "]^2)")
        'V2.0.0.0�J��

        ' A0023NI.BAS �� �v���O�����̏ꍇ
        ' 21730  IF RTP%=2 THEN NT#=25:MSRV#=SRV# ELSE NT#=0:MSRV#=SRV#*(1+SNTC#*NST#)
        'If stUserData.iTempTemp = 2 Then
        '    CalcStandardResistanceValue = stUserData.dStandardRes0
        'Else
        '    CalcStandardResistanceValue = stUserData.dStandardRes0 * (1.0 + stUserData.dResTempCoff * 25.0)
        'End If
    End Function

    '''===============================================================================
    ''' <summary>
    ''' �X�^���_�[�h��R�̃`�F�b�N
    ''' </summary>
    ''' <returns>True:���� False:�X�^���_�[�h��R����l�ُ�</returns>
    ''' <remarks></remarks>
    '''===============================================================================
    Public Function StandardResistanceMeasure() As Boolean
        Dim rn As Integer = 1
        Dim Rtn As Short
        Dim dblMx As Double
        Dim strJUG As String
        Dim Judge As Integer                                            ' ���茋��'V2.0.0.0�H

        Try

            If Not UserSub.IsTrimType1() And Not UserSub.IsTrimType4() Then
                Return (True)
            End If

            If stREG(rn).intSLP <> SLP_RMES Then       ' ����
                Return (True)
            End If
            'V2.0.0.0�@��
            For i As Short = 1 To stPLT.RCount
                If stREG(i).intSLP = SLP_RMES Then                      ' ��R����̂�
                    stREG(i).dblNOM = CalcStandardResistanceValue()     ' �X�^���_�[�h��R����S�Ăɐݒ肷��B
                End If
            Next
            'V2.0.0.0�@��
            'V2.0.0.0�@            stREG(rn).dblNOM = CalcStandardResistanceValue()
            Call DScanModeResetSet(rn, 0, 0)                             ' DC�X�L���i�ɐڑ����鑪����ؑւ� 
            Rtn = V_R_MEAS(stREG(rn).intSLP, stREG(rn).intMType, dblMx, rn, stREG(rn).dblNOM)
            If (Rtn <> cFRS_NORMAL) Then
                Call Z_PRINT("�X�^���_�[�h��R������ł��܂���" & vbCrLf)
                Return (False)
            Else
                ' �ڕW�l���菈��(FT)
                strJUG = Test_ItFt(1, stREG(rn).intMode, dblMx, stREG(rn).dblNOM, stREG(rn).dblITL, stREG(rn).dblITH, Judge)    'V2.0.0.0�HJudge�ǉ�
                If (strJUG <> JG_OK) Then                           ' FT-NG ?
                    Call Z_PRINT("�X�^���_�[�h��R���m�F���Ă������� ����l �� " & dblMx.ToString("0.00000") & "��" & vbCrLf)
                    Return (False)
                Else
                    Return (True)
                End If
            End If

        Catch ex As Exception
            Call Z_PRINT("UserSub.StandardResistanceMeasure() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Function

    '''===============================================================================
    ''' <summary>
    ''' �ڕW�l�̎擾
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    '''===============================================================================
    Public Function GetTRV() As Double
        Return (dTRV)
    End Function

    '===============================================================================
    '�y�@�@�\�z �ڕW�l�Z�o
    '�y���@���z ��R�ԍ�,�J�b�g�ԍ�
    '�y�߂�l�z �ڕW��R�l
    '===============================================================================
    Public Sub CalcTargeResistancetValue(ByVal rn As Integer)

        Try
            If IsTrimType1() Or UserSub.IsTrimType4() Then  'V2.0.0.0�@sTrimType4()�ǉ�

#If OLD_CALURATION Then 'V2.0.0.0�J
                ' ���x�Z���T�[�̏ꍇ�@�@�@�@�@�@�@�F�@TRV�@���@�X�^���_�[�h�����l �^ �X�^���_�[�h�i�O��or�Q�T���j
                If stUserData.iTempTemp = 1 Then        ' �Q�Ɖ��x	�P�F�O��
                    dTRV = dStandardResValue / stUserData.dStandardRes0 * stREG(rn).dblNOM
                    DebugLogOut("TRV:" & dTRV.ToString & " = " & dStandardResValue.ToString & " / " & stUserData.dStandardRes0.ToString & " * " & stREG(rn).dblNOM.ToString)
                Else                                    ' �Q�Ɖ��x	�Q�F�Q�T��
                    dTRV = dStandardResValue / stUserData.dStandardRes25 * stREG(rn).dblNOM
                    DebugLogOut("TRV:" & dTRV.ToString & " = " & dStandardResValue.ToString & " / " & stUserData.dStandardRes25.ToString & " * " & stREG(rn).dblNOM.ToString)
                End If
#Else
                Dim dAlpha As Double = stUserData.dAlpha / 10.0 ^ 6
                Dim dBeta As Double = stUserData.dBeta / 10.0 ^ 6
                Dim dDaihyouAlpha As Double = stUserData.dDaihyouAlpha / 10.0 ^ 6
                Dim dDaihyouBeta As Double = stUserData.dDaihyouBeta / 10.0 ^ 6

                '�X�e�[�W���x���Z�v�Z��=(-��+SQRT(��^2-4*��*(1-STD�����l/STD0����R�l)))/(2*��)
                Dim dStageTempConv As Double = (-1.0 * dAlpha + Math.Sqrt(dAlpha ^ 2 - 4.0 * dBeta * (1.0 - dStandardResValue / stUserData.dTemperatura0))) / (2 * dBeta)
                DebugLogOut("�X�e�[�W���x[" & dStageTempConv.ToString & "]= (-1.0 * " & dAlpha.ToString & " + Sqrt(" & dAlpha.ToString & " ^ 2 - 4.0 * " & dBeta.ToString & " * (1.0 - " & dStandardResValue.ToString & " / " & stUserData.dTemperatura0.ToString & "))) / (2 * " & dBeta.ToString & ")")

                '�Z���T�[�v�Z��(���g���~���O���̖ڕW�l) = (�ݒ��R�l / (1 + �� * �ݒ艷�x + �� * �ݒ艷�x ^ 2)) * (1 + �� * �X�e�[�W���x + �� * �X�e�[�W���x ^ 2)
                dTRV = (stREG(rn).dblNOM / (1 + dDaihyouAlpha * stUserData.iTempTemp + dDaihyouBeta * stUserData.iTempTemp ^ 2)) * (1 + dDaihyouAlpha * dStageTempConv + dDaihyouBeta * dStageTempConv ^ 2)
                DebugLogOut("TRV[" & dTRV.ToString & "] = (" & stREG(rn).dblNOM.ToString & " / (1 + " & dDaihyouAlpha.ToString & " * " & stUserData.iTempTemp.ToString & " + " & dDaihyouBeta.ToString & " * " & stUserData.iTempTemp.ToString & "^ 2)) * (1 + " & dDaihyouAlpha.ToString & " * " & dStageTempConv.ToString & " + " & dDaihyouBeta.ToString & "*" & dStageTempConv.ToString & " ^ 2)")
#End If


            ElseIf IsTrimType2() Or IsTrimType3() Then  'V1.0.4.3�CIsTrimType3()�ǉ�

                'V1.2.0.0�A��
                Dim ResCnt As Short
                'V2.0.0.0�I                If UserSub.IsTrimType3() Then
                'V2.0.0.0�I                ResCnt = 1                  ' �`�b�v��R���[�h�͂P�Ԗڂ����g�p����B
                'V2.0.0.0�I            Else
                ResCnt = rn
                'V2.0.0.0�I            End If
                'V1.2.0.0�A��

                ' �����x������R�g���~���O�̏ꍇ�@�F�@TRV�@���@�ڕW��R�l �~ �␳�l
                'V2.0.0.0�K                dTRV = stREG(rn).dblNOM * stUserData.dNomCalcCoff(ResCnt)
                dTRV = stREG(rn).dblNOM * (stUserData.dNomCalcCoff(UserSub.GetResNumberInCircuit(ResCnt)) / 1000000.0 + 1.0)                   'V2.0.0.0�K �␳�l�̍��ڂ�ppm���͂ɕύX 'V2.0.0.0�I�T�[�L�b�g�Ή�
                DebugLogOut("TRV:" & dTRV.ToString & " = " & stREG(rn).dblNOM.ToString & "* (" & stUserData.dNomCalcCoff(UserSub.GetResNumberInCircuit(ResCnt)).ToString & ") / 1000000.0 + 1.0)")
                'V1.2.0.0�A                dTRV = stREG(rn).dblNOM * stUserData.dNomCalcCoff(rn)
                'V1.2.0.0�A                DebugLogOut("TRV:" & dTRV.ToString & " = " & stREG(rn).dblNOM.ToString & " * " & stUserData.dNomCalcCoff(rn))

            Else
                Call Z_PRINT("UserSub.CalcTargeResistancetValue() ERROR �W���g���~���O�ŌĂ΂�܂��� = " & vbCrLf)
            End If

        Catch ex As Exception
            Call Z_PRINT("UserSub.CalcTargeResistancetValue() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub

    '===============================================================================
    '�y�@�@�\�z �e�J�b�g�̖ڕW�l�Z�o
    '�y���@���z ��R�ԍ�,�J�b�g�ԍ�
    '�y�߂�l�z �ڕW��R�l
    '===============================================================================
    Public Function GetTargeResistancetValue(ByVal rn As Integer, ByVal cn As Integer) As Double

        Try
            If IsTrimType1() Or IsTrimType4() Then

                ' �ڕW�l�iTRM)�@���@TRV�@�|�@�i�J�b�g���̃I�t�Z�b�g�l�@�~�@��������l�@�^�@�ڕW�l�Z�o�W���@�j
                GetTargeResistancetValue = dTRV - (stREG(rn).STCUT(cn).dblCOF * dInitialResValue / stUserData.dTargetCoff(UserBas.GetResistorNo(rn)))
                DebugLogOut("��R[" & rn.ToString & "]�J�b�g[" & cn.ToString & "]�ڕW�l:" & GetTargeResistancetValue.ToString & " = " & dTRV.ToString & " - (" & stREG(rn).STCUT(cn).dblCOF.ToString & " * " & dInitialResValue.ToString & " / " & stUserData.dTargetCoff(UserBas.GetResistorNo(rn)).ToString & ")")

            ElseIf IsTrimType2() Or IsTrimType3() Then  'V1.0.4.3�CIsTrimType3()�ǉ�

                'V1.2.0.0�A��
                Dim ResCnt As Short
                'V2.0.0.0�I                If UserSub.IsTrimType3() Then
                'V2.0.0.0�I                    ResCnt = 1                  ' �`�b�v��R���[�h�͂P�Ԗڂ����g�p����B
                'V2.0.0.0�I                Else
                ResCnt = rn
                'V2.0.0.0�I                End If
                'V1.2.0.0�A��

                ' �ڕW�l�iTRM)�@���@TRV�@�|�@�i�J�b�g���̃I�t�Z�b�g�l�@�~�@��������l�@�^�@�ڕW�l�Z�o�W���@�j
                'V2.0.0.0�I                GetTargeResistancetValue = dTRV - (stREG(rn).STCUT(cn).dblCOF * dInitialResValue / stUserData.dTargetCoff(ResCnt))
                GetTargeResistancetValue = dTRV - (stREG(rn).STCUT(cn).dblCOF * dInitialResValue / stUserData.dTargetCoff(UserSub.GetResNumberInCircuit(ResCnt)))       'V2.0.0.0�I UserSub.GetResNumberInCircuit(ResCnt)�ǉ�
                DebugLogOut("��R[" & rn.ToString & "]�J�b�g[" & cn.ToString & "]�ڕW�l:" & GetTargeResistancetValue.ToString & " = " & dTRV.ToString & " - (" & stREG(rn).STCUT(cn).dblCOF.ToString & " * " & dInitialResValue.ToString & " / " & stUserData.dTargetCoff(UserSub.GetResNumberInCircuit(ResCnt)).ToString & ")")
                'V1.2.0.0�A                GetTargeResistancetValue = dTRV - (stREG(rn).STCUT(cn).dblCOF * dInitialResValue / stUserData.dTargetCoff(rn))
                'V1.2.0.0�A                DebugLogOut("��R[" & rn.ToString & "]�J�b�g[" & cn.ToString & "]�ڕW�l:" & GetTargeResistancetValue.ToString & " = " & dTRV.ToString & " - (" & stREG(rn).STCUT(cn).dblCOF.ToString & " * " & dInitialResValue.ToString & " / " & stUserData.dTargetCoff(rn).ToString & ")")

            Else
                Call Z_PRINT("UserSub.GetTargeResistancetValue() ERROR �W���g���~���O�ŌĂ΂�܂��� = " & vbCrLf)
            End If

        Catch ex As Exception
            Call Z_PRINT("UserSub.GetTargeResistancetValue() TRAP ERROR = " & ex.Message & vbCrLf)
            GetTargeResistancetValue = -9999.999
        End Try
    End Function

    '===============================================================================
    '�y�@�@�\�z ��R����A�����A�����x����̕ύX
    '�y���@���z ��R�ԍ�,�J�b�g�ԍ�
    '�y�߂�l�z ����
    '===============================================================================
    Public Sub ChangeMeasureSpeed(ByVal rn As Integer, ByVal cn As Integer, ByVal idx As Short)

        Try
            intTMM_Save = stREG(rn).STCUT(cn).intIXTMM(idx)         ' ���胂�[�h(0:�����@1:�����x)
            intMType_Save = stREG(rn).STCUT(cn).intIXMType(idx)     ' ����@��0�`5(0:��������@1�`:�O���@��)

            If stUserData.iTrimSpeed = 1 Then                       ' ����
                stREG(rn).STCUT(cn).intIXTMM(idx) = 0               ' ���[�h(0:����(�R���p���[�^��ϕ����[�h), 1:�����x(�ϕ����[�h))�@���C���f�b�N�X���g�p����B
                stREG(rn).STCUT(cn).intIXMType(idx) = 0             ' ������(0=��������, 1=�O������)
            ElseIf stUserData.iTrimSpeed = 2 Then                   ' �����x
                'V1.2.0.0�A��
                Dim ResCnt As Short
                'V2.0.0.0�I                If UserSub.IsTrimType3() Then
                'V2.0.0.0�I                    ResCnt = 1                  ' �`�b�v��R���[�h�͂P�Ԗڂ����g�p����B
                'V2.0.0.0�I                Else
                ResCnt = rn
                'V2.0.0.0�I                End If
                If cn < stUserData.iChangeSpeed(GetResistorNo(ResCnt)) Then
                    'V1.2.0.0�A��
                    'V1.2.0.0�A                    If cn < stUserData.iChangeSpeed(GetResistorNo(rn)) Then
                    stREG(rn).STCUT(cn).intIXTMM(idx) = 0           ' ���[�h(0:����(�R���p���[�^��ϕ����[�h), 1:�����x(�ϕ����[�h))�@���C���f�b�N�X���g�p����B
                    stREG(rn).STCUT(cn).intIXMType(idx) = 0         ' ������(0=��������, 1=�O������)
                End If
            Else
                Exit Sub
            End If
        Catch ex As Exception
            Call Z_PRINT("UserSub.ChangeMeasureSpeed() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Sub
    '===============================================================================
    '�y�@�@�\�z ��R����A�����A�����x���胂�[�h�̕���
    '�y���@���z ��R�ԍ�,�J�b�g�ԍ�
    '�y�߂�l�z ����
    '===============================================================================
    Public Sub ResoreMeasureSpeed(ByVal rn As Integer, ByVal cn As Integer, ByVal idx As Short)
        Try
            stREG(rn).STCUT(cn).intIXTMM(idx) = intTMM_Save        ' ���[�h�̕ۑ�
            stREG(rn).STCUT(cn).intIXMType(idx) = intMType_Save    ' �����ʂ̕ۑ�
        Catch ex As Exception
            Call Z_PRINT("UserSub.ResoreMeasureSpeed() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub

    '===============================================================================
    '�y�@�@�\�z ���x�Z���T�[�^�C�v���𔻒f����
    '�y���@���z ����
    '�y�߂�l�z True = ��v, False = �s��v
    '===============================================================================
    Public Function IsTrimType1() As Boolean
        If stUserData.iTrimType = 1 Then
            IsTrimType1 = True
        Else
            IsTrimType1 = False
        End If
    End Function
    '===============================================================================
    '�y�@�@�\�z ��R�g���~���O�^�C�v���𔻒f����
    '�y���@���z ����
    '�y�߂�l�z True = ��v, False = �s��v
    '===============================================================================
    Public Function IsTrimType2() As Boolean
        If stUserData.iTrimType = 2 Then
            IsTrimType2 = True
        Else
            IsTrimType2 = False
        End If
    End Function
    'V1.0.4.3�C ADD START
    '===============================================================================
    '�y�@�@�\�z �`�b�v��R�g���~���O�^�C�v���𔻒f����
    '�y���@���z ����
    '�y�߂�l�z True = ��v, False = �s��v
    '===============================================================================
    Public Function IsTrimType3() As Boolean
        If stUserData.iTrimType = 3 Then
            IsTrimType3 = True
        Else
            IsTrimType3 = False
        End If
    End Function
    'V1.0.4.3�C ADD END
    'V2.0.0.0�@ ADD START
    '===============================================================================
    '�y�@�@�\�z �`�b�v���x�Z���T�[�^�C�v���𔻒f����
    '�y���@���z ����
    '�y�߂�l�z True = ��v, False = �s��v
    '===============================================================================
    Public Function IsTrimType4() As Boolean
        If stUserData.iTrimType = 4 Then
            IsTrimType4 = True
        Else
            IsTrimType4 = False
        End If
    End Function

    'V2.2.1.7�@ ��
    '===============================================================================
    '�y�@�@�\�z �}�[�N�󎚃^�C�v���𔻒f����
    '�y���@���z ����
    '�y�߂�l�z True = ��v, False = �s��v
    '===============================================================================
    Public Function IsTrimType5() As Boolean
        If stUserData.iTrimType = 5 Then
            IsTrimType5 = True
        Else
            IsTrimType5 = False
        End If
    End Function
    'V2.2.1.7�@ ��

    ' ''' <summary>
    ' ''' �`�b�v�^�C�v���𔻒f����
    ' ''' </summary>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Function IsCircuitTrimType() As Boolean
    '    '20171207���X�ؗl�Ƒ��k���ă`�b�v��R�g���~���O�^�C�v�݂̂ɓK�p�@If stUserData.iTrimType = 3 Or stUserData.iTrimType = 4 Then
    '    If stUserData.iTrimType = 3 Then
    '        IsCircuitTrimType = True
    '    Else
    '        IsCircuitTrimType = False
    '    End If
    'End Function
    'V2.0.0.0�@ ADD END
    '===============================================================================
    '�y�@�@�\�z ���ꏈ���̃g���~���O�^�C�v���𔻒f����
    '�y���@���z ����
    '�y�߂�l�z True = ��v, False = �s��v
    '===============================================================================
    Public Function IsSpecialTrimType() As Boolean
        If stUserData.iTrimType <> 0 Then
            IsSpecialTrimType = True
        Else
            IsSpecialTrimType = False
        End If
    End Function
    '===============================================================================
    '�y�@�@�\�z ��������l�̕ۑ�
    '�y���@���z ��������l
    '�y�߂�l�z ����
    '===============================================================================
    Public Sub SetInitialResValue(ByVal dVal As Double)
        dInitialResValue = dVal
    End Sub

    'V2.1.0.0�D��
    ''' <summary>
    ''' ��������l�̎擾
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetInitialResValue() As Double
        Return (dInitialResValue)
    End Function
    'V2.1.0.0�D��
    '===============================================================================
    '�y�@�@�\�z �W����R����l�̕ۑ�
    '�y���@���z �W����R����l
    '�y�߂�l�z ����
    '===============================================================================
    Public Sub SetStandardResValue(ByVal dVal As Double)
        dStandardResValue = dVal
    End Sub

    '===============================================================================
    '�y�@�@�\�z �t�@�C�i�������̏���
    '�y���@���z ��R�ԍ�,�t�@�C�i���e�X�g����l
    '�y�߂�l�z ����
    '===============================================================================
    Public Sub DevCalculation(ByVal rn As Integer, ByVal dFtVal As Double)
        Dim iResNo As Integer
        Try


            iResNo = GetResistorNo(rn)      ' �g���~���O�f�[�^��̒�R�ԍ�����J�b�g�����R�ԍ������߂�B�i����݂̂����O����B�j

            'V1.2.0.0�A��
            'V2.0.0.0�I            If UserSub.IsTrimType3() Then
            'V2.0.0.0�I                iResNo = 1
            'V2.0.0.0�I            End If
            'V1.2.0.0�A��

            If iResNo > MAX_RES_USER Then
                Return
            End If

            stUserData.dFtVal(iResNo) = dFtVal

            '14710      DEV1#=FIX((R.FT1#-NRV1#)/NRV1#*1000000#)
            '14832      DEV2#=FIX((R.FT2#-NRV2#)/NRV2#*1000000#)

            If stREG(rn).dblNOM = 0.0 Then
                Call Z_PRINT("UserSub.DevCalculation() �ڕW�l���O�ł��B�v�Z���o���܂���" & vbCrLf)
                Exit Sub
            End If
            ' �g���~���O�덷�@���@�i�@�g���~���O�l�@�|�@�X�^���_�[�h�����l�ɑ΂��Ẵg���~���O�ڕW�l�@�j�^�X�^���_�[�h�����l�ɑ΂��Ẵg���~���O�ڕW�l�@* 10^6
            'V2.0.0.0�Q            stUserData.dDev(iResNo) = FNDEVP(stUserData.dFtVal(iResNo), stREG(rn).dblNOM)
            stUserData.dDev(iResNo) = FNDEVP(stUserData.dFtVal(iResNo), UserSub.GetTRV())

        Catch ex As Exception
            Call Z_PRINT("UserSub.DevCalculation() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub

    '''===============================================================================
    ''' <summary>
    ''' ���b�g�����̉ۊm�F
    ''' </summary>
    ''' <param name="HostMode">���[�_���[�h(cHOSTcMODEcMANUAL:�蓮 cHOSTcMODEcAUTO:����)</param>
    ''' <param name="Start">cHSTcTRMCMD�F�g���~���O�X�^�[�g�@cHSTcLOTCHANGE�F���b�g�����X�^�[�g</param>
    ''' <returns>0�F���b�g�������� 1:�������B�̃��b�g���� 2:�؂�ւ��M���̃��b�g����</returns>
    ''' <remarks></remarks>
    '''===============================================================================
'V2.2.1.1�G'    Public Function IsLotChange(ByVal HostMode As Integer, ByVal Start As Short, ByVal fStartTrim As Boolean) As Integer
    Public Function IsLotChange(ByVal HostMode As Integer, ByVal Start As Short, ByVal fStartTrim As Boolean, ByRef lotcnt As Integer) As Integer   'V2.2.1.1�G


        Dim bPrint As Boolean = False
        Dim LdIDat As Integer

        IsLotChange = 0

        'If Start <> cHSTcTRMCMD And Start <> cHSTcLOTCHANGE Then
        '    Exit Function
        'End If


        Select Case (stUserData.iLotChange) ' ���b�g�I������ 0:�I���������薳�� 1:���� 2:���[�_�[�M�� 3:����
            Case 0
                IsLotChange = 0
            Case 1
                If stCounter.PlateCounter >= stUserData.lLotEndSL Then      ' ��������ɓ��B
                    If fStartTrim Then
                        IsLotChange = 1
                    End If
                    If Not UserBas.stCounter.LotPrint Then                  'V1.2.0.3
                        bPrint = True
                    End If                                                  'V1.2.0.3
                End If
            Case 2
                'V1.2.0.0�C��
                If giLoaderType = 1 Then
                    ObjSys.Z_ATLDGET(LdIDat)                                        ' ���[�_�[����
                    If LdIDat And clsLoaderIf.LINP_LOT_CHG Then
                        LdIDat = cHSTcLOTCHANGE
                    End If
                    Start = LdIDat And cHSTcLOTCHANGE
                End If

                If fStartTrim And HostMode = cHOSTcMODEcAUTO And Start = cHSTcLOTCHANGE Then
                    IsLotChange = 2
                    bPrint = True
                End If
            Case 3
                'V1.2.0.0�C��
                If giLoaderType = 1 Then
                    ObjSys.Z_ATLDGET(LdIDat)                                        ' ���[�_�[����
                    If LdIDat And clsLoaderIf.LINP_LOT_CHG Then

                        LdIDat = cHSTcLOTCHANGE
                    End If
                    Start = LdIDat And cHSTcLOTCHANGE
                End If

                If fStartTrim And HostMode = cHOSTcMODEcAUTO And Start = cHSTcLOTCHANGE Then
                    IsLotChange = 2
                    bPrint = True
                Else
                    If stCounter.PlateCounter >= stUserData.lLotEndSL Then      ' ��������ɓ��B
                        If fStartTrim Then
                            IsLotChange = 1
                        End If
                        If Not UserBas.stCounter.LotPrint Then                  'V1.2.0.3
                            bPrint = True
                        End If                                                  'V1.2.0.3
                    End If
                End If
        End Select

        ''V2.2.1.1�G ��
        '�t���O�ȊO�̏����Ń��b�g�؂�ւ����s���ꍇ�́A���̕����b�g�؂�ւ��񐔂����Z���� 
        If IsLotChange = 1 Then
            If lotcnt > 0 Then  '���b�g�؂�ւ��t���O��ON���Ă����ꍇ���s���� 
                lotcnt = lotcnt + 1
            End If
        End If

        '���b�g�؂�ւ��t���O��ON���Ă����ꍇ���s���� 
        If lotcnt > 0 Then
            IsLotChange = 1

            Call Z_PRINT("���b�g�؂�ւ��t���O�ɂ��A���b�g�؂�ւ����s���܂����B" & lotcnt.ToString)

            If Not UserBas.stCounter.LotPrint Then
                bPrint = True
            End If
        End If
        ''V2.2.1.1�G ��

        'V1.2.0.0�E        If bPrint And Not UserBas.stCounter.LotPrint Then
        If bPrint Then
            Call UserSub.LotEnd()                           ' ���b�g�I�����̃f�[�^�o��
            Call Printer.Print(False)                       ' ���b�g�����
            UserBas.stCounter.LotPrint = True               ' ���b�g�I�����̈�����s�ς݂�True
        End If

    End Function

    '===============================================================================
    '�y�@�@�\�z ����w�b�_���t�@�C���̍쐬�A�����t�@�C���̍폜���s��
    '�y���@���z ����
    '�y�߂�l�z ����
    '===============================================================================
    Public Sub MakePrintFileHeader()
        Dim WS As IO.StreamWriter
        Dim sData As String


        Try
            UserBas.stCounter.LotPrint = False              'V1.2.0.3 �O�̈גǉ�
            ' ����f�[�^���폜����B
            If IO.File.Exists(cTRIM_PRINT_DATA_HEAD) = True Then
                IO.File.Delete(cTRIM_PRINT_DATA_HEAD)
            End If

            If IO.File.Exists(cTRIM_PRINT_DATA_RES) = True Then
                IO.File.Delete(cTRIM_PRINT_DATA_RES)
            End If
            If IO.File.Exists(cTRIM_PRINT_DATA_PLATE) = True Then
                IO.File.Delete(cTRIM_PRINT_DATA_PLATE)
            End If
            If IO.File.Exists(cTRIM_PRINT_DATA_END) = True Then
                IO.File.Delete(cTRIM_PRINT_DATA_END)
            End If

            ' �w�b�_�[�����o�͂���B
            WS = New IO.StreamWriter(cTRIM_PRINT_DATA_HEAD, True, System.Text.Encoding.GetEncoding("Shift-JIS"))
            WS.WriteLine("����������������������������������������������������������������������������������������������")
            WS.WriteLine("���t  " & DateTime.Now.ToString("yyyy/MM/dd"))
            WS.WriteLine("���b�g�m���D      �� " & stUserData.sLotNumber.PadRight(20) & "�I�y���[�^��     �� " & stUserData.sOperator)
            WS.WriteLine("�p�^�[���m���D    �� " & stUserData.sPatternNo.PadRight(20) & "�v���O�����m���D �� " & stUserData.sProgramNo)

            If UserSub.IsTrimType1() Or UserSub.IsTrimType4() Then       ' ���x�Z���T�['V2.0.0.0�@sTrimType4()�ǉ�
                WS.WriteLine("�q�P:�ݒ��R�l             �� " & stREG(UserBas.GetCutResistorNo(1)).dblNOM.ToString(TARGET_DIGIT_DEFINE).PadLeft(15) & " [ohm]")
#If OLD_CALURATION Then 'V2.0.0.0�J
                WS.WriteLine("�X�^���_�[�h��R�l�i�O���j  �� " & stUserData.dStandardRes0.ToString("0.00000").PadLeft(15) & " [ohm]")
                WS.WriteLine("�X�^���_�[�h��R�l�i�Q�T���j�� " & stUserData.dStandardRes25.ToString("0.00000").PadLeft(15) & " [ohm]")
                If stUserData.iTempTemp = 1 Then    ' �Q�Ɖ��x	�P�F�O�� �܂��� �Q�F�Q�T��
                    WS.WriteLine("�Q�Ɖ��x�@�@�@�@�@�@�@�@�@�@���O��")
                ElseIf stUserData.iTempTemp = 2 Then
                    WS.WriteLine("�Q�Ɖ��x�@�@�@�@�@�@�@�@�@�@���Q�T��")
                End If
#Else
                'V2.0.0.4�@                Dim dStdResValue As Double = stUserData.dTemperatura0 * (1.0 + stUserData.dAlpha * stUserData.iTempTemp + stUserData.dBeta * stUserData.iTempTemp ^ 2)
                'V2.0.0.4�@��
                Dim dAlpha As Double = stUserData.dAlpha / 10.0 ^ 6
                Dim dBeta As Double = stUserData.dBeta / 10.0 ^ 6
                Dim dStdResValue As Double = stUserData.dTemperatura0 * (1.0 + dAlpha * stUserData.iTempTemp + dBeta * stUserData.iTempTemp ^ 2)
                'V2.0.0.4�@��
                WS.WriteLine("STD��R�l�i" & stUserData.iTempTemp.ToString("0") & "���j  �� " & dStdResValue.ToString(TARGET_DIGIT_DEFINE).PadLeft(15) & " [ohm]")
#End If
                WS.WriteLine("�t�@�C�i���e�X�g���~�b�g[ppm] High �� " & stUserData.dFinalLimitHigh.ToString("0.0").PadLeft(10) & "  Low�@�� " & stUserData.dFinalLimitLow.ToString("0.0").PadLeft(10))
            Else                        ' ��R�g���~���O
                If stUserData.iTrimSpeed = 1 Then
                    sData = "�g���~���O���[�h     �� �����x���[�h"
                ElseIf stUserData.iTrimSpeed = 2 Then
                    sData = "�g���~���O���[�h     �� �����x���[�h"
                Else
                    sData = "�g���~���O���[�h     �� �ݒ�l"
                End If
                WS.WriteLine(sData)

                Dim Rcnt As Integer = UserBas.GetRCountExceptMeasure()
                'V2.2.0.0�O��
                If stMultiBlock.gMultiBlock <> 0 Then

                    ' �}���`�u���b�N�Őݒ肳��Ă��镪���s���� 
                    For blk As Integer = 0 To stMultiBlock.BLOCK_DATA.Length - 2
                        ' @'V2.2.0.033 If stMultiBlock.BLOCK_DATA(0).gBlockCnt <> 0 Then
                        If stMultiBlock.BLOCK_DATA(blk).gBlockCnt <> 0 Then     'V2.2.0.033
                            WS.WriteLine("MBNo�F " & stMultiBlock.BLOCK_DATA(blk).DataNo.ToString)
                            'V2.2.0.033   WS.WriteLine("�}���`�u���b�NNo�F " & stMultiBlock.BLOCK_DATA(blk).DataNo.ToString)

                            For rn As Integer = 1 To Rcnt
                                WS.WriteLine("R" & rn.ToString & ":�ݒ��R�l     �� " & stMultiBlock.BLOCK_DATA(blk).dblNominal(UserBas.GetCutResistorNo(rn) - 1).ToString(TARGET_DIGIT_DEFINE).PadLeft(15) & " [ohm]")
                            Next rn
                            For rn As Integer = 1 To Rcnt
                                WS.WriteLine("R" & rn.ToString & ":�␳�l         �� " & stMultiBlock.BLOCK_DATA(blk).dblCorr(rn - 1).ToString("0.000").PadLeft(15) & " [ppm]") 'V2.0.0.0�K�␳�l�̍��ڂ�ppm���͂ɕύX
                            Next rn
                        End If
                    Next blk
                Else
                    For rn As Integer = 1 To Rcnt
                        WS.WriteLine("R" & rn.ToString & ":�ݒ��R�l     �� " & stREG(UserBas.GetCutResistorNo(rn)).dblNOM.ToString(TARGET_DIGIT_DEFINE).PadLeft(15) & " [ohm]")
                        'V2.0.0.0�I                    If IsTrimType3() Then           'V1.2.0.0�A
                        'V2.0.0.0�I                        Exit For                    'V1.2.0.0�A
                        'V2.0.0.0�I                    End If                          'V1.2.0.0�A
                    Next
                    For rn As Integer = 1 To Rcnt
                        WS.WriteLine("R" & rn.ToString & ":�␳�l         �� " & stUserData.dNomCalcCoff(rn).ToString("0.000").PadLeft(15) & " [ppm]") 'V2.0.0.0�K�␳�l�̍��ڂ�ppm���͂ɕύX
                        'V2.0.0.0�I                    If IsTrimType3() Then           'V1.2.0.0�A
                        'V2.0.0.0�I                        Exit For                    'V1.2.0.0�A
                        'V2.0.0.0�I                    End If                          'V1.2.0.0�A
                    Next
                End If
                'V2.2.0.0�O��
                WS.WriteLine("�t�@�C�i���e�X�g���~�b�g[ppm] High �� " & stUserData.dFinalLimitHigh.ToString("0.0").PadLeft(10) & "  Low�@�� " & stUserData.dFinalLimitLow.ToString("0.0").PadLeft(10))
                'V2.0.0.0�I                If IsTrimType2() Then               'V1.2.0.0�A
                WS.WriteLine("���Βl���~�b�g[ppm]              �� " & stUserData.dRelativeHigh.ToString("0.000").PadLeft(10))
                'V2.0.0.0�I            End If                              'V1.2.0.0�A

            End If

            WS.WriteLine("����������������������������������������������������������������������������������������������")

            WS.Close()

        Catch ex As Exception
            Call Z_PRINT("UserSub.MakePrintFileHeader() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Sub
    '===============================================================================
    '�y�@�@�\�z �f�q�P�ʂ̃t�@�C�����̓��O�̏o��
    '�y���@���z IO.StreamWriter�A�����f�[�^
    '�y�߂�l�z ����
    '===============================================================================
    Private Sub ResistorDataOutPut(ByVal WS As IO.StreamWriter, ByVal bPrint As Boolean, ByVal sMessage As String)
        Dim printcnt As Integer = 0
        Dim blkcnt As Integer = 0

        'V2.2.0.033��
        If stMultiBlock.gMultiBlock <> 0 Then
            For cnt As Integer = 0 To 4
                If (stMultiBlock.BLOCK_DATA(cnt).gBlockCnt) <> 0 Then
                    blkcnt = blkcnt + 1
                End If
            Next

            printcnt = stUserData.lPrintRes \ blkcnt


            If stMultiBlock.BLOCK_DATA(stExecBlkData.DataNo).gProcCnt < printcnt And bPrint Then ' ���b�g�I��������f�q���ɒB���Ă��Ȃ��ꍇ�́A�t�@�C���ւ̏o��    'V2.2.0.033
                WS.WriteLine(sMessage)
            End If

        Else
            printcnt = stUserData.lPrintRes
            If lResCounterForPrinter < printcnt And bPrint Then ' ���b�g�I��������f�q���ɒB���Ă��Ȃ��ꍇ�́A�t�@�C���ւ̏o��    'V2.2.0.033
                WS.WriteLine(sMessage)
            End If

        End If


        'V2.2.0.033        If lResCounterForPrinter < stUserData.lPrintRes And bPrint Then ' ���b�g�I��������f�q���ɒB���Ă��Ȃ��ꍇ�́A�t�@�C���ւ̏o��
        '        If lResCounterForPrinter < printcnt And bPrint Then ' ���b�g�I��������f�q���ɒB���Ă��Ȃ��ꍇ�́A�t�@�C���ւ̏o��    'V2.2.0.033

        Call Z_PRINT(sMessage.Replace(vbTab, " ") & vbCrLf)         ' ���O�o�̓G���A�ւ̏o��

    End Sub
    '''=============================================================================
    ''' <summary>
    ''' �f�q�P�ʂ̔���NG��
    ''' </summary>
    ''' <remarks></remarks>
    '''=============================================================================
    Public Sub NgJudgeSet()
        bOkJudge = False
    End Sub

    Public Sub SkipSet()
        bSkip = True
    End Sub

    ''' <summary>
    ''' �T�[�L�b�g�X�L�b�v�́ATrue
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SkipGet() As Boolean
        Return (bSkip)
    End Function
    '''=============================================================================
    ''' <summary>
    ''' �f�q�P�ʂ̔��菉����
    ''' </summary>
    ''' <remarks></remarks>
    '''=============================================================================
    Public Sub NgJudgeReset()
        For i As Integer = 1 To MAX_RES_USER
            sResistorPrintData(i) = ""
        Next
        bOkJudge = True
        bSkip = False
    End Sub

    '===============================================================================
    '�y�@�@�\�z �S��R�I�����̔���
    '�y���@���z ��R�ԍ�
    '�y�߂�l�z ����
    '===============================================================================
    Public Function FinalJudge(ByVal rn As Integer) As Boolean
        Dim WS As IO.StreamWriter
        Dim dDev As Double
        Dim sJudge As String
        Dim iResCnt As Integer
        Dim bHeaderPrint As Boolean
        Dim iCnt As Integer

        ' 14840       DEV#=DEV1#-DEV2#
        ' 14841       IF ECF%<>1 THEN FTOV3#=FTOV3#+1#: GOTO *TRIM.NG
        ' 14842       CALL TEST%(DEV#,SRV#,Z2,STLO#,STHI#)
        ' 14843       IF ECF%=2 THEN FTLO3#=FTLO3#+1#:  GOTO *TRIM.NG
        ' 14844       IF ECF%=3 THEN FTHI3#=FTHI3#+1#:  GOTO *TRIM.NG

        'V2.0.0.0�D �S�Ă�"0.00000"��TARGET_DIGIT_DEFINE�֕ύX
        'V2.0.0.0�D�@PadLeft(13)��PadLeft(15)�֕ύX

        Try
            If IO.File.Exists(cTRIM_PRINT_DATA_RES) Then
                bHeaderPrint = False
            Else
                bHeaderPrint = True
            End If

            FinalJudge = True


            iResCnt = GetRCountExceptMeasure()
            If iResCnt > MAX_RES_USER Then
                iResCnt = MAX_RES_USER
            End If

            'V1.2.0.0�A��
            'V2.0.0.0�I            If UserSub.IsTrimType3() Then
            'V2.0.0.0�I                iResCnt = 1
            'V2.0.0.0�I            End If
            'V1.2.0.0�A��

            WS = New IO.StreamWriter(cTRIM_PRINT_DATA_RES, True, System.Text.Encoding.GetEncoding("Shift-JIS"))     ' ��R�f�[�^����f�[�^
            If bHeaderPrint Then
                If UserSub.IsTrimType1 Or UserSub.IsTrimType4() Then    'V2.0.0.0�@sTrimType4()�ǉ�
                    ResistorDataOutPut(WS, True, "No" & vbTab & "X" & vbTab & "Y" & vbTab & "��R��" & vbTab & "�ڕW��R�l �C�j�V��������l �t�@�C�i������l       �덷    ����")
                Else
                    'V2.0.0.0�A��
                    If DGL = TRIM_VARIATION_MEAS Then ' ����l�ϓ�����
                        ResistorDataOutPut(WS, True, "No" & vbTab & "X" & vbTab & "Y" & vbTab & "��R��" & vbTab & "�g���~���O��e�s�l �t�@�C�i������l   �덷   ����")
                    Else
                        'V2.0.0.0�A��
                        'V2.2.0.033��
                        If stMultiBlock.gMultiBlock <> 0 Then
                            ResistorDataOutPut(WS, True, "No" & vbTab & "X" & vbTab & "Y" & vbTab & "��R��       " & vbTab & "�C�j�V��������l   �t�@�C�i������l   �덷   ����")
                        Else
                            ResistorDataOutPut(WS, True, "No" & vbTab & "X" & vbTab & "Y" & vbTab & "��R��" & vbTab & "�C�j�V��������l   �t�@�C�i������l   �덷   ����")
                        End If

                    End If                      'V2.0.0.0�A
                End If
            End If
            If (stREG(rn).intMode = 0) Then                         ' ���胂�[�h = 0(�䗦(ppm)) ?
                dDev = FNDEVP(dblVX(2), dblNM(2))                ' �덷 = (����l / �ڕW�l - 1) * 100
            Else
                dDev = dblVX(2) - dblNM(2)                       ' �덷1(��Βl) = ����l - �ڕW�l
            End If

            'V2.2.0.033��
            Dim addStr As String = ""
            If stMultiBlock.gMultiBlock <> 0 Then
                addStr = " MBNo:" & stExecBlkData.DataNo.ToString()
            Else
                addStr = ""
            End If
            'V2.2.0.033��

            If UserSub.IsTrimType1 Or iResCnt = 1 Then                                     ' �P�f�q�P��R�̎��@�Q��R�ȏ�́A�Ō�ɏo��

                If Not stREG(rn).bPattern Then                  'V1.2.0.0�B �J�b�g�ʒu�␳�̔��� True�FOK False:NG
                    'V2.2.0.033                    sResistorPrintData(1) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & "   " & vbTab & "�J�b�g�ʒu�␳ �����m�f���� = �m�f"
                    sResistorPrintData(1) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & addStr & "   " & vbTab & "�J�b�g�ʒu�␳ �����m�f���� = �m�f"       'V2.2.0.033 
                ElseIf UserSub.IsTrimType1 Or UserSub.IsTrimType4() Then
                    'V2.2.0.033  sResistorPrintData(1) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & vbTab & dblNM(2).ToString(TARGET_DIGIT_DEFINE).PadLeft(15) & " " & UserSub.ChangeOverFlow(dblVX(1).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dblVX(2).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dDev.ToString("0.0")).PadLeft(15) & " " & strJUG(rn)
                    sResistorPrintData(1) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & addStr & vbTab & dblNM(2).ToString(TARGET_DIGIT_DEFINE).PadLeft(15) & " " & UserSub.ChangeOverFlow(dblVX(1).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dblVX(2).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dDev.ToString("0.0")).PadLeft(15) & " " & strJUG(rn)    'V2.2.0.033 
                Else
                    'V2.0.0.0�A��
                    If DGL = TRIM_VARIATION_MEAS Then ' ����l�ϓ�����
                        'V2.2.0.033  sResistorPrintData(1) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & vbTab & UserSub.ChangeOverFlow(dMeasVariationNOM(rn).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dblVX(2).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dDev.ToString("0.0")).PadLeft(15) & "  " & strJUG(rn)
                        sResistorPrintData(1) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & addStr & vbTab & UserSub.ChangeOverFlow(dMeasVariationNOM(rn).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dblVX(2).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dDev.ToString("0.0")).PadLeft(15) & "  " & strJUG(rn)
                    Else
                        'V2.0.0.0�A��
                        'V2.2.0.033 sResistorPrintData(1) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & vbTab & UserSub.ChangeOverFlow(dblVX(1).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dblVX(2).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dDev.ToString("0.0")).PadLeft(15) & "  " & strJUG(rn)
                        sResistorPrintData(1) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & addStr & vbTab & UserSub.ChangeOverFlow(dblVX(1).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dblVX(2).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dDev.ToString("0.0")).PadLeft(15) & "  " & strJUG(rn)
                    End If                      'V2.0.0.0�A
                End If
                If strJUG(rn) = JG_OK Then
                    ResistorDataOutPut(WS, True, sResistorPrintData(1))
                    lResCounterForPrinter = lResCounterForPrinter + 1                       ' �n�j�݈̂�����ăJ�E���g�A�b�v����B
                    If UserSub.IsTrimType2 Then
                        stCounter.OK_Counter = stCounter.OK_Counter + 1
                        stCounter.Total_OK_Counter = stCounter.Total_OK_Counter + 1

                        'V2.2.0.0�O��
                        If stMultiBlock.gMultiBlock <> 0 Then
                            stMultiBlock.BLOCK_DATA(stExecBlkData.DataNo).gProcCnt = stMultiBlock.BLOCK_DATA(stExecBlkData.DataNo).gProcCnt + 1
                            gObjFrmDistribute.SetOkCounterMulti()
                        End If
                        'V2.2.0.0�O��

                    End If
                Else
                    ResistorDataOutPut(WS, False, sResistorPrintData(1))
                End If
            Else
                iCnt = GetResistorNo(rn)
                If iCnt <= MAX_RES_USER Then
                    'V2.0.0.0�A��
                    If DGL = TRIM_VARIATION_MEAS Then ' ����l�ϓ�����
                        'V2.2.0.033 sResistorPrintData(iCnt) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & vbTab & UserSub.ChangeOverFlow(dMeasVariationNOM(rn).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dblVX(2).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dDev.ToString("0.0")).PadLeft(15) & "  " & strJUG(rn)
                        sResistorPrintData(iCnt) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & addStr & vbTab & UserSub.ChangeOverFlow(dMeasVariationNOM(rn).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dblVX(2).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dDev.ToString("0.0")).PadLeft(15) & "  " & strJUG(rn)
                    Else
                        'V2.0.0.0�A��
                        'V2.2.0.033 sResistorPrintData(iCnt) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & vbTab & UserSub.ChangeOverFlow(dblVX(1).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dblVX(2).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dDev.ToString("0.0")).PadLeft(15) & "  " & strJUG(rn)
                        sResistorPrintData(iCnt) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & addStr & vbTab & UserSub.ChangeOverFlow(dblVX(1).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dblVX(2).ToString(TARGET_DIGIT_DEFINE)).PadLeft(15) & " " & UserSub.ChangeOverFlow(dDev.ToString("0.0")).PadLeft(15) & "  " & strJUG(rn)
                    End If                      'V2.0.0.0�A
                    'V1.2.0.0�B��
                    If Not stREG(rn).bPattern Then                  '�J�b�g�ʒu�␳�̔��� True�FOK False:NG
                        'V2.2.0.033 sResistorPrintData(iCnt) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & "   " & vbTab & "�J�b�g�ʒu�␳ �����m�f���� = �m�f"
                        sResistorPrintData(iCnt) = stCounter.PlateCounter.ToString(0) & vbTab & stCounter.BlockCntX.ToString(0) & vbTab & stCounter.BlockCntY.ToString(0) & vbTab & stREG(rn).strRNO & addStr & "   " & vbTab & "�J�b�g�ʒu�␳ �����m�f���� = �m�f"
                    End If
                    'V1.2.0.0�B��
                End If
            End If

            'V2.0.0.0�I            If UserSub.IsTrimType2 Then
            If UserSub.IsTrimType2 Or UserSub.IsTrimType3() Then        'V2.0.0.0�I�`�b�v��R���[�h�ǉ�
                If iResCnt > 1 And GetResistorNo(rn) = iResCnt Then     ' ��R���P�ȏ�̎�

                    Dim sMessage As String = ""
                    Dim largest As Double = Double.MinValue
                    Dim smallest As Double = Double.MaxValue
                    Dim i As Integer

                    If bSkip Then
                        dDev = -1000000.0
                    Else
                        For i = 1 To iResCnt Step 1
                            largest = Math.Max(largest, stUserData.dDev(i))
                            smallest = Math.Min(smallest, stUserData.dDev(i))
                        Next

                        ' ���Βl�@���@�ő�l�@�|�@�ŏ��l

                        dDev = Math.Abs(largest - smallest)
                        DebugLogOut("DEV =" & dDev.ToString(TARGET_DIGIT_DEFINE) & " L= " & largest.ToString(TARGET_DIGIT_DEFINE) & " M= " & smallest.ToString(TARGET_DIGIT_DEFINE))
                    End If


                    If dDev <= stUserData.dRelativeHigh And bOkJudge Then
                        stCounter.OK_Counter = stCounter.OK_Counter + 1
                        stCounter.Total_OK_Counter = stCounter.Total_OK_Counter + 1

                        'V2.2.0.0�O��
                        If stMultiBlock.gMultiBlock <> 0 Then
                            gObjFrmDistribute.SetOkCounterMulti()
                        End If
                        'V2.2.0.0�O��

                        sJudge = "OK"
                        For i = 1 To iResCnt Step 1
                            ResistorDataOutPut(WS, True, sResistorPrintData(i))
                        Next
                        sMessage = sMessage & "���Βl[ppm] " & UserSub.ChangeOverFlow(dDev.ToString("0.0")) & vbTab & sJudge
                        ResistorDataOutPut(WS, True, sMessage)
                        lResCounterForPrinter = lResCounterForPrinter + 1                       ' �n�j�݈̂�����ăJ�E���g�A�b�v����B
                        'V2.2.0.0�O��
                        If stMultiBlock.gMultiBlock <> 0 Then
                            stMultiBlock.BLOCK_DATA(stExecBlkData.DataNo).gProcCnt = stMultiBlock.BLOCK_DATA(stExecBlkData.DataNo).gProcCnt + 1
                        End If
                    Else
                        FinalJudge = False
                        sJudge = "NG"
                        For i = 1 To iResCnt Step 1
                            ResistorDataOutPut(WS, False, sResistorPrintData(i))
                        Next
                        sMessage = sMessage & "���Βl[ppm] " & UserSub.ChangeOverFlow(dDev.ToString("0.0")) & vbTab & sJudge
                        ResistorDataOutPut(WS, False, sMessage)
                        If Not bSkip Then
                            stCounter.FTHigh = stCounter.FTHigh + 1
                            stCounter.Total_FTHigh = stCounter.Total_FTHigh + 1

                            ' 'V2.2.0.0�O��
                            If stMultiBlock.gMultiBlock <> 0 Then
                                gObjFrmDistribute.SetFTHighCounterMulti()
                            End If
                            ' 'V2.2.0.0�O��

                        End If
                        strJUG(rn) = JG_FH
                    End If
                    NgJudgeReset()
                End If
            End If
            WS.Close()

            'V1.2.0.0�A            If GetResistorNo(rn) = iResCnt Then
            'V2.0.0.0�I            If GetResistorNo(rn) = iResCnt Or UserSub.IsTrimType3() Then    'V1.2.0.0�A �`�b�v��R���[�h�͒�R�P�ʂŃJ�E���g����B
            If GetResistorNo(rn) = iResCnt Then         'V2.0.0.0�I
                stCounter.TrimCounter = stCounter.TrimCounter + 1 ' ���ݸސ����ı���
                stCounter.Total_TrimCounter = stCounter.Total_TrimCounter + 1

                ' 'V2.2.0.0�O��
                If stMultiBlock.gMultiBlock <> 0 Then
                    gObjFrmDistribute.SetTrimCounterMulti()
                End If
                ' 'V2.2.0.0�O��

            End If

            Call Set_NG_Counter()                       'V1.2.0.0�B NG�J�E���^�[�̍X�V

        Catch ex As Exception
            Call Z_PRINT("UserSub.FinalJudge() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Function

    '===============================================================================
    '�y�@�@�\�z ��I�����̈���������s��
    '�y���@���z ����
    '�y�߂�l�z ����
    '===============================================================================
    Public Sub SubstrateEnd()
        Dim WS As IO.StreamWriter

        Try
            If (stCounter.PlateCounter = 0) Then
                Return
            End If
            '###1030�B            UserBas.stCounter.EndTime = DateTime.Now()              ' ������I�����ԕۑ�

            WS = New IO.StreamWriter(cTRIM_PRINT_DATA_PLATE, True, System.Text.Encoding.GetEncoding("Shift-JIS"))   ' ��f�[�^���
            If stCounter.PlateCounter = 1 Then
                WS.WriteLine("����������������������������������������������������������������������������������������������")
            End If

            If stMultiBlock.gMultiBlock <> 0 Then

                WS.WriteLine("No." & stCounter.PlateCounter.ToString & "  Start = " & stCounter.StartTime.ToString("HH:mm:ss") & " End = " & stCounter.EndTime.ToString("HH:mm:ss"))

                ' �}���`�u���b�N�Őݒ肳��Ă��镪���s���� 
                For blk As Integer = 1 To stMultiBlock.BLOCK_DATA.Length - 1

                    'V2.2.0.033 If stMultiBlock.BLOCK_DATA(0).gBlockCnt <> 0 Then
                    If stMultiBlock.BLOCK_DATA(blk - 1).gBlockCnt <> 0 Then       'V2.2.0.033

                        With stToTalDataMulti(blk)

                            'V2.2.0.033 WS.WriteLine("�}���`�u���b�NNo�F " & stMultiBlock.BLOCK_DATA(blk - 1).DataNo.ToString)
                            WS.WriteLine("MBNo�F " & stMultiBlock.BLOCK_DATA(blk - 1).DataNo.ToString)

                            WS.WriteLine("  Element = " & .stCounter1.TrimCounter.ToString & " Good = " & .stCounter1.OK_Counter.ToString & " pcs  Reject = " & .stCounter1.NG_Counter.ToString & " pcs")
                            WS.WriteLine("Initilal Low = " & .stCounter1.ITLow.ToString.PadRight(20) & "High = " & .stCounter1.ITHigh.ToString.PadRight(20) & "Open = " & .stCounter1.ITOpen.ToString)
                            WS.WriteLine("On Trim Low  = " & .stCounter1.FTLow.ToString.PadRight(10) & " ( " & .stCounter1.ValLow.ToString.PadRight(4) & " ) " & "High = " & .stCounter1.FTHigh.ToString.PadRight(10) & " ( " & .stCounter1.ValHigh.ToString.PadRight(4) & " ) " & "Open = " & .stCounter1.FTOpen.ToString)

                        End With

                    End If

                Next blk

            Else
                WS.WriteLine("No." & stCounter.PlateCounter.ToString & "  Start = " & stCounter.StartTime.ToString("HH:mm:ss") & " End = " & stCounter.EndTime.ToString("HH:mm:ss") & "  Element = " & stCounter.TrimCounter.ToString & " Good = " & stCounter.OK_Counter.ToString & " pcs  Reject = " & stCounter.NG_Counter.ToString & " pcs")
                ''V2.2.0.0�O            WS.WriteLine("Initilal Low = " & stCounter.ITLow.ToString & "  High =" & stCounter.ITHigh.ToString & "  Open = " & stCounter.ITOpen.ToString & "         On Trim Low = " & stCounter.FTLow.ToString & "  High = " & stCounter.FTHigh.ToString & "  Open = " & stCounter.FTOpen.ToString)
                WS.WriteLine("Initilal Low = " & stCounter.ITLow.ToString.PadRight(20) & "High = " & stCounter.ITHigh.ToString.PadRight(20) & "Open = " & stCounter.ITOpen.ToString)       'V2.2.0.0�O
                WS.WriteLine("On Trim Low  = " & stCounter.FTLow.ToString.PadRight(10) & " ( " & stCounter.ValLow.ToString.PadRight(4) & " ) " & "High = " & stCounter.FTHigh.ToString.PadRight(10) & " ( " & stCounter.ValHigh.ToString.PadRight(4) & " ) " & "Open = " & stCounter.FTOpen.ToString)      'V2.2.0.0�O
            End If

            WS.Close()
        Catch ex As Exception
            Call Z_PRINT("UserSub.SubstrateEnd() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Sub

    'V2.0.0.0�H��
    ''' <summary>
    ''' ���v�f�[�^�̏o��
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub StatisticalPrintDataOut()
        Try
            Dim WS As IO.StreamWriter
            Dim iResCnt As Integer
            Dim JudgeMode As Integer = FINAL_TEST
            Dim dMin As Double, dMax As Double, dAve As Double, dDev As Double
            Dim No As Integer = 0


            iResCnt = GetRCountExceptMeasure()
            If iResCnt > MAX_RES_USER Then
                iResCnt = MAX_RES_USER
            End If

            WS = New IO.StreamWriter(cTRIM_PRINT_DATA_RES, True, System.Text.Encoding.GetEncoding("Shift-JIS"))     ' ��R�f�[�^����f�[�^
            WS.WriteLine("����������������������������������������������������������������������������������������������")
            WS.WriteLine("��R���@    �ŏ��@        �ő�@        ���ρ@        �W���΍�")

            'V2.2.0.0�O��
            If stMultiBlock.gMultiBlock <> 0 Then
                ' ������R�l�Ή���

                ' �}���`�u���b�N�Őݒ肳��Ă��镪���s���� 
                For blk As Integer = 1 To stMultiBlock.BLOCK_DATA.Length - 1

                    If stMultiBlock.BLOCK_DATA(blk - 1).gBlockCnt <> 0 Then       'V2.2.0.033

                        'V2.2.0.033 WS.WriteLine("�}���`�u���b�NNo�F " & stMultiBlock.BLOCK_DATA(blk - 1).DataNo.ToString)
                        WS.WriteLine("MBNo�F " & stMultiBlock.BLOCK_DATA(blk - 1).DataNo.ToString)

                        'V2.2.0.033 For rn As Integer = 1 To MAX_RES_USER
                        For rn As Integer = 1 To stPLT.RCount   'V2.2.0.033
                            If UserModule.IsCutResistor(rn) Then
                                Call gObjFrmDistribute.StatisticalDataGetMulti(JudgeMode, rn, dMin, dMax, dAve, dDev, blk)
                                WS.WriteLine(stREG(rn).strRNO & vbTab & UserSub.ChangeOverFlow(dMin.ToString(TARGET_DIGIT_DEFINE)).PadLeft(13) & " " & UserSub.ChangeOverFlow(dMax.ToString(TARGET_DIGIT_DEFINE)).PadLeft(13) & " " & UserSub.ChangeOverFlow(dAve.ToString(TARGET_DIGIT_DEFINE)).PadLeft(13) & " " & UserSub.ChangeOverFlow(dDev.ToString(TARGET_DIGIT_DEFINE)).PadLeft(13))
                            End If
                        Next rn

                    End If

                Next blk

            Else

                For rn As Integer = 1 To stPLT.RCount
                    If UserModule.IsCutResistor(rn) Then
                        No = No + 1
                        Call gObjFrmDistribute.StatisticalDataGet(JudgeMode, No, dMin, dMax, dAve, dDev)
                        WS.WriteLine(stREG(rn).strRNO & vbTab & UserSub.ChangeOverFlow(dMin.ToString(TARGET_DIGIT_DEFINE)).PadLeft(13) & " " & UserSub.ChangeOverFlow(dMax.ToString(TARGET_DIGIT_DEFINE)).PadLeft(13) & " " & UserSub.ChangeOverFlow(dAve.ToString(TARGET_DIGIT_DEFINE)).PadLeft(13) & " " & UserSub.ChangeOverFlow(dDev.ToString(TARGET_DIGIT_DEFINE)).PadLeft(13))
                        If No >= iResCnt Then
                            Exit For
                        End If
                    End If
                Next

            End If
            'V2.2.0.0�O��

            WS.Close()

        Catch ex As Exception
            Call Z_PRINT("UserSub.StatisticalPrintDataOut() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub
    'V2.0.0.0�H��
    '===============================================================================
    '�y�@�@�\�z ���b�g�I�����̈���������s��
    '�y���@���z ����
    '�y�߂�l�z ����
    '===============================================================================
    Public Sub LotEnd()
        Dim WS As IO.StreamWriter

        Try
            UserBas.stCounter.LotEnd = DateTime.Now()           ' ���b�g�I������

            Call StatisticalPrintDataOut()                      ' ���v�f�[�^�o��'V2.0.0.0�H

            If IO.File.Exists(cTRIM_PRINT_DATA_END) = True Then ' ����t�@�C�����폜����B
                IO.File.Delete(cTRIM_PRINT_DATA_END)
            End If
            WS = New IO.StreamWriter(cTRIM_PRINT_DATA_END, True, System.Text.Encoding.GetEncoding("Shift-JIS"))   ' ���b�g�I�����f�[�^���
            WS.WriteLine("����������������������������������������������������������������������������������������������")
            If stMultiBlock.gMultiBlock <> 0 Then

                'WS.WriteLine("Substrate = " & stCounter.PlateCounter.ToString())

                '' �}���`�u���b�N�Őݒ肳��Ă��镪���s���� 
                'For blk As Integer = 1 To stMultiBlock.BLOCK_DATA.Length - 1

                '    ' 'V2.2.0.033 If stMultiBlock.BLOCK_DATA(0).gBlockCnt <> 0 Then
                '    If stMultiBlock.BLOCK_DATA(blk - 1).gBlockCnt <> 0 Then       'V2.2.0.033

                '        With stToTalDataMulti(blk)

                '            ' 'V2.2.0.033 WS.WriteLine("�}���`�u���b�NNo�F " & stMultiBlock.BLOCK_DATA(blk - 1).DataNo.ToString)
                '            WS.WriteLine("MBNo�F " & stMultiBlock.BLOCK_DATA(blk - 1).DataNo.ToString)

                '            WS.WriteLine("   Element = " & .stCounter1.Total_TrimCounter.ToString & " pcs  Good = " & .stCounter1.Total_OK_Counter.ToString & " pcs  Reject = " & .stCounter1.Total_NG_Counter.ToString & " pcs")
                '            WS.WriteLine("Initilal Low =  " & .stCounter1.Total_ITLow.ToString & "  High = " & .stCounter1.Total_ITHigh.ToString & " Open = " & .stCounter1.Total_ITOpen.ToString)
                '            WS.WriteLine("On Trim Low =  " & .stCounter1.Total_FTLow.ToString & " ( " & .stCounter1.Total_ValLow.ToString & " ) " & "  High = " & .stCounter1.Total_FTHigh.ToString & " ( " & .stCounter1.Total_ValHigh.ToString & " ) " & " Open = " & .stCounter1.Total_FTOpen.ToString)

                '        End With

                '    End If

                'Next blk

            Else

                WS.WriteLine("Substrate = " & stCounter.PlateCounter.ToString & "   Element = " & stCounter.Total_TrimCounter.ToString & " pcs  Good = " & stCounter.Total_OK_Counter.ToString & " pcs  Reject = " & stCounter.Total_NG_Counter.ToString & " pcs")
                'V2.2.0.029           WS.WriteLine("Initilal Low =  " & stCounter.Total_ITLow.ToString & "  High = " & stCounter.Total_ITHigh.ToString & " Open = " & stCounter.Total_ITOpen.ToString & "         On Trim Low =  " & stCounter.Total_FTLow.ToString & "  High = " & stCounter.Total_FTHigh.ToString & " Open = " & stCounter.Total_FTOpen.ToString)
                'V2.2.1.1�@ WS.WriteLine("Initilal Low =  " & stCounter.Total_ITLow.ToString & "  High = " & stCounter.Total_ITHigh.ToString & " Open = " & stCounter.Total_ITOpen.ToString)            'V2.2.0.029
                'V2.2.1.1�@ WS.WriteLine("On Trim Low =  " & stCounter.Total_FTLow.ToString & " ( " & stCounter.Total_ValLow.ToString & " ) " & "  High = " & stCounter.Total_FTHigh.ToString & " ( " & stCounter.Total_ValHigh.ToString & " ) " & " Open = " & stCounter.Total_FTOpen.ToString)     'V2.2.0.029
                WS.WriteLine("Initilal Low = " & stCounter.Total_ITLow.ToString.PadRight(20) & "High = " & stCounter.Total_ITHigh.ToString.PadRight(20) & "Open = " & stCounter.Total_ITOpen.ToString.PadRight(10))            'V2.2.1.1�@
                WS.WriteLine("On Trim Low  = " & stCounter.Total_FTLow.ToString.PadRight(10) & " ( " & stCounter.Total_ValLow.ToString.PadRight(4) & " ) " & "High = " & stCounter.Total_FTHigh.ToString.PadRight(10) & " ( " & stCounter.Total_ValHigh.ToString.PadRight(4) & " ) " & "Open = " & stCounter.Total_FTOpen.ToString.PadRight(10))     'V2.2.1.1�@


            End If
            WS.WriteLine("����������������������������������������������������������������������������������������������")
            WS.WriteLine("�ݒ�f�[�^�m�F���ԁF�@" & stCounter.LotStart.ToString("HH:mm:ss") & " �@�I�����ԁF " & stCounter.LotEnd.ToString("HH:mm:ss") & " �@�o�ߎ��ԁF�@" & stCounter.LotEnd.Subtract(stCounter.LotStart).Hours.ToString("00") & ":" & stCounter.LotEnd.Subtract(stCounter.LotStart).Minutes.ToString("00") & ":" & stCounter.LotEnd.Subtract(stCounter.LotStart).Seconds.ToString("00"))
            WS.WriteLine("����������������������������������������������������������������������������������������������")
            WS.Close()


            UserSub.VariationMesStartDataReset()                'V2.0.0.0�A ����l�ϓ����o�@�\�J�n�u���b�N�ʒu������

            WriteLogMarkPrint()         ' ���b�g�̃��O���e���t�@�C���o��               V2.2.1.7�B

        Catch ex As Exception
            Call Z_PRINT("UserSub.LotEnd() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Sub

    Public Function ChangeOverFlow(ByVal sNum As String) As String
        Dim iPos As Integer

        iPos = sNum.IndexOf(".")
        If iPos > 9 Or (iPos = -1 And sNum.Length > 9) Then         ' �P�O�O�l�ȏ�̌��܂��͏����_�������ĕ�����X�ȏ�̏ꍇ�́A�O�ɂ���B
            'V2.0.0.0�D            Return ("0.00000")
            Return (TARGET_DIGIT_DEFINE)                            'V2.0.0.0�D
        Else
            Return (sNum)
        End If

    End Function

    'V1.0.4.3�H��
    ''' <summary>
    ''' ���݃����[�{�[�h�Ή��A�`�����l���ϊ��@�����V�`�P�U�˂����R�R�`�S�Q
    ''' </summary>
    ''' <param name="ProbeChannel"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ConvtChannel(ByRef ProbeChannel As Short) As Short
        Try
            'V2.0.0.0�L��
            If bRelayBoard Then
                If 9 <= ProbeChannel And ProbeChannel <= 18 Then
                    ConvtChannel = 24 + ProbeChannel
                Else
                    ConvtChannel = ProbeChannel
                End If
            Else
                'V2.0.0.0�L��
                If 7 <= ProbeChannel And ProbeChannel <= 16 Then
                    ConvtChannel = 26 + ProbeChannel
                Else
                    ConvtChannel = ProbeChannel
                End If
            End If                                      'V2.0.0.0�L
        Catch ex As Exception
            Call Z_PRINT("UserSub.ConvertChannel() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function
    'V1.0.4.3�H��
    'V1.0.4.3�I��
    ''' <summary>
    ''' ����}�[�L���O���[�h�E�t�@�C�i������݂̂̏ꍇ��TRIM_MODE_ITTRFT�Ɣ��肷��B
    ''' </summary>
    ''' <returns>True:TRIM_MODE_ITTRFT�@False:TRIM_MODE_ITTRFT��TRIM_MODE_MEAS_MARK�ȊO</returns>
    ''' <remarks></remarks>
    Public Function IsTRIM_MODE_ITTRFT() As Boolean
        Try
            'V2.0.0.0�A            If (DGL = TRIM_MODE_ITTRFT Or DGL = TRIM_MODE_MEAS_MARK) Then
            If (DGL = TRIM_MODE_ITTRFT Or DGL = TRIM_MODE_MEAS_MARK Or DGL = TRIM_MODE_POWER Or DGL = TRIM_VARIATION_MEAS) Then
                Return (True)
            Else
                Return (False)
            End If

        Catch ex As Exception

        End Try
    End Function
    'V1.0.4.3�I��

#Region "��������׽"
    '===============================================================================
    '�y�@�@�\�z ��������������Ȃ��׽
    '�y�d�@�l�z Print ҿ��ނ��Ăяo���Ĉ���������Ȃ�
    '===============================================================================
    Public Class cPrintDocument
        Private Const FILE_ENCODING As String = "shift_jis"
        'Private Const FILE_ENCODING As String = "utf-8"
        Private ReadOnly FONT_SIZE As Font = New Font("�l�r �S�V�b�N", 9.0!)
        Private ReadOnly FILEPATH_ARRAY As String() = {cTRIM_PRINT_DATA_HEAD, _
                                                       cTRIM_PRINT_DATA_RES, _
                                                       cTRIM_PRINT_DATA_PLATE, _
                                                       cTRIM_PRINT_DATA_END}
        Private Const MSG_YESNO As String = "���b�g���̈�������s���܂��B"
        Private Const MSG_FILE_NOTHING As String = " ��������܂���B"
        Private Const MARGIN_LEFT As Integer = 100      ' ���ϰ��ݍ�(1/100 ����P��)��̫�Ă�100
        Private Const MARGIN_RIGHT As Integer = 100     ' ���ϰ��݉E(1/100 ����P��)��̫�Ă�100
        Private Const MARGIN_TOP As Integer = 10        ' ���ϰ��ݏ�(1/100 ����P��)��̫�Ă�100
        Private Const MARGIN_BOTTOM As Integer = 10     ' ���ϰ��݉�(1/100 ����P��)��̫�Ă�100

        Private Const DEF_TEXT_BUF As Integer = 1024    ' ���ׂĂ̕������ޯ̧(����Ȃ��Ȃ�Ίg�������)
        Private Const DEF_LINE_BUF As Integer = 128     ' ��s�̕������ޯ̧(����Ȃ��Ȃ�Ίg�������)
        Private m_PrintTextBuf As StringBuilder         ' ������镶�����ׂĂ��i�[����
        Private m_BufIndex As Integer                   ' ���݂̕����ʒu

        Private Const PRINTDEFAULT_DIR = "C:\TRIMDATA\PRINTDEFAULT"            ' V2.2.0.024
        Private Const PRINTLOG_DIR = "C:\TRIMDATA\PRINTLOG"            ' V2.2.2.0�F 

        ''' <summary>����������Ȃ�</summary>
        ''' <param name="confirmMsgBox">True�̏ꍇ�m�Fү���ނ�\������</param>
        Public Sub Print(Optional ByVal confirmMsgBox As Boolean = False)

            If (True = confirmMsgBox) Then ' �m�Fү���ނ̕\��
                If (MsgBoxResult.No = MsgBox(MSG_YESNO, DirectCast( _
                                             MsgBoxStyle.YesNo + _
                                             MsgBoxStyle.Question, MsgBoxStyle), _
                                             My.Application.Info.Title)) _
                                             Then Exit Sub
            End If

            Try
                'V2.2.1.7�C��
                ' �}�[�N�󎚃��[�h�͈�����Ȃ� 
                If UserSub.IsTrimType5() Then
                    Return
                End If
                'V2.2.1.7�C��

                Using printer As New PrintDocument
                    m_PrintTextBuf = New StringBuilder(DEF_TEXT_BUF)
                    m_BufIndex = 0

                    For Each path As String In FILEPATH_ARRAY
                        ''V2.2.0.033��
                        'If (stMultiBlock.gMultiBlock <> 0) AndAlso (path = cTRIM_PRINT_DATA_END) Then
                        '    Continue For
                        'End If
                        ''V2.2.0.033��

                        If (True = File.Exists(path)) Then
                            ' ̧�ق����݂���ꍇ
                            Using sr As New StreamReader(path, System.Text.Encoding.GetEncoding(FILE_ENCODING))
                                While (-1 < sr.Peek()) ' �g�p�ł��镶�����Ȃ��Ȃ�܂Ōp��
                                    m_PrintTextBuf.Append(sr.ReadLine() & vbLf) ' ��s�Âǉ�����
                                End While
                            End Using
                            'V2.2.2.0�F��
                            Dim orgfilename As String = IO.Path.GetFileName(path)
                            Dim writefolder As String = PRINTLOG_DIR & "\" & DateTime.Now.ToString("yyyyMM") & "\"
                            '                            If (False = System.IO.File.Exists(writefolder)) Then
                            If (False = IO.Directory.Exists(writefolder)) Then
                                MkDir(writefolder)
                            End If
                            FileCopy(path, writefolder & "\" & DateTime.Now.ToString("yyyyMMdd_hhmmss_") & stUserData.sLotNumber.Trim() & "_" & orgfilename)
                            'V2.2.2.0�F��

                        Else
                            ' ̧�ق����݂��Ȃ��|��ǉ�����
                            m_PrintTextBuf.Append(vbLf & path & MSG_FILE_NOTHING & vbLf & vbLf)
                        End If
                    Next

                    With printer.DefaultPageSettings
                        .Margins = New Margins(MARGIN_LEFT, MARGIN_RIGHT, MARGIN_TOP, MARGIN_BOTTOM)
                        .Landscape = False ' �p���̌���(�c)
                    End With
                    AddHandler printer.PrintPage, AddressOf m_Printer_PrintPage

                    'V2.2.0.024��
                    If printer.PrinterSettings.PrinterName = "Microsoft Print to PDF" Then
                        printer.PrinterSettings.PrintToFile = True
                        ' PDF�̏o�͐�ƃt�@�C�������w��
                        printer.PrinterSettings.PrintFileName = PRINTDEFAULT_DIR & "\" & System.IO.Path.GetFileNameWithoutExtension(gsDataFileName) & "_" & Now.ToString("yyyyMMddHHmmss") & ".pdf"
                    End If
                    'V2.2.0.024��
                    Call printer.Print() ' m_Printer_PrintPage() ����Ă���������



                End Using

            Catch ex As Exception
                Call MsgBox(ex.ToString())
            End Try

        End Sub

        ''' <summary>�����������o��</summary>
        ''' <remarks>��s���Ƃɐ擪���W���w�肵�ĕ`�悷��</remarks>
        Private Sub m_Printer_PrintPage(ByVal sender As Object, ByVal e As PrintPageEventArgs)
            Dim x As Integer = e.MarginBounds.Left ' ����J�n�����ʒu
            Dim y As Integer = e.MarginBounds.Top  ' ����J�n�����ʒu

            Try
                ' ���݂��߰�ނɂ����܂� ���� ���������ׂď����o���Ă��Ȃ��ꍇ�p��
                While ((y + FONT_SIZE.Height) < e.MarginBounds.Bottom) AndAlso _
                        (m_BufIndex < m_PrintTextBuf.Length)
                    Dim lineBuf As New StringBuilder(DEF_LINE_BUF) ' �s�ޯ̧

                    While True
                        If (m_PrintTextBuf.Length <= m_BufIndex) OrElse _
                            (vbLf = m_PrintTextBuf.Chars(m_BufIndex)) Then
                            ' ���������ׂď����o���� �܂��� ���s���ނ̏ꍇ
                            m_BufIndex += 1
                            Exit While
                        End If

                        lineBuf.Append(m_PrintTextBuf.Chars(m_BufIndex)) ' �ꕶ���ǉ�����
                        If ((e.MarginBounds.Width) < _
                            (e.Graphics.MeasureString(lineBuf.ToString(), FONT_SIZE).Width)) Then
                            ' ������ɂ����܂�Ȃ��ꍇ�A�ꕶ���폜����
                            lineBuf.Remove(lineBuf.Length - 1, 1)
                            Exit While
                        End If

                        m_BufIndex += 1
                    End While

                    ' ��s�������o��
                    'Debug.Print(lineBuf.ToString())
                    e.Graphics.DrawString(lineBuf.ToString(), FONT_SIZE, Brushes.Black, x, y)
                    y += FONT_SIZE.GetHeight(e.Graphics) ' ���̍s�̈���ʒu��
                End While

                If (m_PrintTextBuf.Length <= m_BufIndex) Then
                    ' ���������ׂď����o�����ꍇ
                    e.HasMorePages = False
                    'Debug.Print((m_PrintTextBuf.Capacity).ToString()) ' DEF_TEXT_BUF �̻��ޒ���
                    m_PrintTextBuf = Nothing
                    m_BufIndex = 0
                Else
                    ' ���̐ݒ�ɂ��ēx m_Printer_PrintPage() ����Ă���������
                    e.HasMorePages = True ' �����߰�ނ�
                End If

            Catch ex As Exception
                Call MsgBox(ex.ToString())
            End Try
        End Sub

    End Class
#End Region
    ''' <summary>
    ''' �����^�]���̃I�t�Z�b�g�p�����[�^���f�����i�e�[�u���ʒu�A�r�[���ʒu�A�v���[�u�ڐG�ʒu�j
    ''' </summary>
    ''' <param name="AutoDataFileFullPath"></param>
    ''' <param name="iAutoDataFileNum"></param>
    ''' <returns>����I���FcFRS_NORMAL�@�ُ�I���F�f�[�^�̔ԍ�</returns>
    ''' <remarks></remarks>
    Public Function SetOffSetDataToAutoOperationData(ByVal AutoDataFileFullPath() As String, ByVal iAutoDataFileNum As Short) As Short

        Dim r As Short
        Dim stPLT_Local As PLATE_DATA                          ' �v���[�g�f�[�^

        If iAutoDataFileNum <= 0 Then
            Return (True)
        End If

        gsDataFileName = AutoDataFileFullPath(0)                     ' �f�[�^�t�@�C�����ݒ�

        r = rData_load()                                            ' �f�[�^�t�@�C�����[�h
        If (r <> 0) Then                                            ' �f�[�^�t�@�C���@���[�h�G���[
            Return (1)
        End If

        stPLT_Local.z_xoff = stPLT.z_xoff                           ' �e�[�u���ʒu�I�t�Z�b�g�@�g�����|�W�V�����I�t�Z�b�gX(mm)
        stPLT_Local.z_yoff = stPLT.z_yoff                           ' �e�[�u���ʒu�I�t�Z�b�g�@�g�����|�W�V�����I�t�Z�b�gY(mm)

        stPLT_Local.BPOX = stPLT.BPOX                               ' �r�[���ʒu�I�t�Z�b�g�@BP Offset X(mm)
        stPLT_Local.BPOY = stPLT.BPOY                               ' �r�[���ʒu�I�t�Z�b�g�@BP Offset Y(mm)

        stPLT_Local.Z_ZON = stPLT.Z_ZON                             ' Z PROBE ON �ʒu(mm)

        For Cnt As Integer = 1 To (iAutoDataFileNum - 1)
            gsDataFileName = AutoDataFileFullPath(Cnt)
            r = rData_load()                                            ' �f�[�^�t�@�C�����[�h
            If (r <> 0) Then                                            ' �f�[�^�t�@�C���@���[�h�G���[
                Return (Cnt + 1)
            End If
            stPLT.z_xoff = stPLT_Local.z_xoff       ' �e�[�u���ʒu�I�t�Z�b�g�@�g�����|�W�V�����I�t�Z�b�gX(mm)
            stPLT.z_yoff = stPLT_Local.z_yoff       ' �e�[�u���ʒu�I�t�Z�b�g�@�g�����|�W�V�����I�t�Z�b�gY(mm)
            stPLT.BPOX = stPLT_Local.BPOX           ' �r�[���ʒu�I�t�Z�b�g�@BP Offset X(mm)
            stPLT.BPOY = stPLT_Local.BPOY           ' �r�[���ʒu�I�t�Z�b�g�@BP Offset Y(mm)
            stPLT.Z_ZON = stPLT_Local.Z_ZON         ' Z PROBE ON �ʒu(mm)

            If rData_save(gsDataFileName) <> cFRS_NORMAL Then       ' �f�[�^�t�@�C���Z�[�u
                Return (Cnt + 1)
            End If
        Next

        Return (cFRS_NORMAL)

    End Function

    'V1.2.0.0�B��
    ''' <summary>
    ''' �p�^�[���F�����ʊi�[�̈�̏������i������Ԃ͂n�j�j
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub InitResPatternmatchResult()
        Try
            Dim iResNo As Integer
            For iResNo = 1 To MAXRNO Step 1
                stREG(iResNo).bPattern = True
            Next
        Catch ex As Exception
            Call Z_PRINT("UserSub.InitResPatternmatchResult() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub
    'V1.2.0.0�B��
    'V1.2.0.0�A��
    Public Function MarkingForChipMode(ByVal rn As Short, ByVal Result As Boolean) As Short
        Try
            Dim rNo As Integer
            Dim cn As Integer
            Dim dOffSetX As Double, dOffSetY As Double
            Dim Rtn As Short = cFRS_NORMAL              ' �֐��ߒl

            If Not UserSub.IsTrimType3() And Not UserSub.IsTrimType4() Then     ' ���x�Z���T�['V2.0.0.0�@sTrimType4()�ǉ�
                Return (Rtn)
            End If

            ' ��P��R��P�J�b�g�̃X�^�[�g���W���猻�݂̒�R�̑�P�J�b�g�̃X�^�[�g���W�܂ł̋��������߂�B
            dOffSetX = stREG(rn).STCUT(1).dblSTX - stREG(1).STCUT(1).dblSTX
            dOffSetY = stREG(rn).STCUT(1).dblSTY - stREG(1).STCUT(1).dblSTY

            'V2.0.0.0�I��
            If UserSub.IsTrimType3() Then
                rn = UserSub.GetTopResNoinCircuit(rn)
            End If
            'V2.0.0.0�I��

            'V1.2.0.2��
            For rNo = 1 To stPLT.RCount Step 1
                If IsCutResistor(rNo) Then
                    dOffSetX = stREG(rn).STCUT(1).dblSTX - stREG(rNo).STCUT(1).dblSTX
                    dOffSetY = stREG(rn).STCUT(1).dblSTY - stREG(rNo).STCUT(1).dblSTY
                    Exit For
                End If
            Next
            'V1.2.0.2��

            'V1.2.0.2            For rNo = 1 To MAXRNO Step 1
            For rNo = 1 To stPLT.RCount Step 1                  'V1.2.0.2
                If Result Then                                  ' �n�j����̎�
                    If stREG(rNo).intSLP <> SLP_OK_MARK Then    ' �n�j�}�[�N�ȊO�̓X�L�b�v
                        Continue For
                    End If
                Else                                            ' �m�f����̎�
                    If stREG(rNo).intSLP <> SLP_NG_MARK Then    ' �m�f�}�[�N�ȊO�̓X�L�b�v
                        Continue For
                    End If
                End If
                ' �J�b�g�ʒu�����݂̒�R�̈ʒu�ɍ��킹�ăI�t�Z�b�g������B
                For cn = 1 To stREG(rNo).intTNN Step 1
                    stREG(rNo).STCUT(cn).dblSTX = stREG(rNo).STCUT(cn).dblSTX + dOffSetX
                    stREG(rNo).STCUT(cn).dblSTY = stREG(rNo).STCUT(cn).dblSTY + dOffSetY
                Next
                Rtn = VTrim_One(rNo, stREG(rNo).dblNOM)          ' 1��R���g���~���O���s��
                ' �J�b�g�ʒu�����ɖ߂��B
                For cn = 1 To stREG(rNo).intTNN Step 1
                    stREG(rNo).STCUT(cn).dblSTX = stREG(rNo).STCUT(cn).dblSTX - dOffSetX
                    stREG(rNo).STCUT(cn).dblSTY = stREG(rNo).STCUT(cn).dblSTY - dOffSetY
                Next
            Next

            Return (Rtn)

        Catch ex As Exception
            Call Z_PRINT("UserSub.MarkingForChipMode() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function
    'V1.2.0.0�A��

    'V2.0.0.0�A��
    ''' <summary>
    ''' ���O�f�[�^����FT�l�̎擾
    ''' </summary>
    ''' <param name="sLotNumber">���b�g�ԍ�</param>
    ''' <param name="PlateNumber">��ԍ�</param>
    ''' <param name="BlockX">�u���b�N�w�ԍ�</param>
    ''' <param name="BlockY">�u���b�N�x�ԍ�</param>
    ''' <param name="ResCounter">�������ċ��߂��f�[�^��</param>
    ''' <param name="Target">�e�s�l�i�z��j</param>
    ''' <returns></returns>
    ''' <remarks></remarks>    
    Public Function GetTargerDataFromLogFile(ByVal sLotNumber As String, ByVal PlateNumber As Integer, ByVal BlockX As Integer, ByVal BlockY As Integer, ByRef ResCounter As Integer, ByRef Target() As Double) As Boolean

        Dim sPath As String
        Dim sInpData As String
        Dim splt() As String
        Dim ITEM_LOT_NUM As Integer = 3
        Dim ITEM_PLATE_NUM As Integer = 4
        Dim ITEM_BLOCKX_NUM As Integer = 5
        Dim ITEM_BLOCKY_NUM As Integer = 6
        Dim ITEM_FT_NUM As Integer = 12
        Dim itemcnt As Integer
        Dim sLot As String
        Dim nPlate As Integer
        Dim nBlockX As Integer
        Dim nBlockY As Integer
        Dim SearchDir As String
        Dim SearchFile As String

        Try

            SearchDir = "C:\TRIMDATA\LOG"                                   '�T�[�`����t�H���_�w��
            SearchFile = "*" + sLotNumber.Trim() + ".CSV"                   '�T�[�`����t�@�C���̌����L�[(���b�g�ԍ����܂܂�Ă���CSV�t�@�C��)

            '��������v����t�@�C�����̎擾���s �u�����Ώۂ͎w��t�H���_�̂݁v�T�u�t�H���_�͏��O.�T�u���܂߂�ꍇ�͍Ō�̈������uSearchOption.AllDirectories�v�ɕύX
            Dim files() As String = System.IO.Directory.GetFiles(SearchDir, SearchFile, SearchOption.TopDirectoryOnly)

            ' �擾�������ׂẴt�@�C�����ŏI�������ݓ������Ń\�[�g����)
            Array.Sort(Of String)(files, AddressOf CompareLastWriteTime)
            If (files.Length = 0) Then
                Return (False)
            End If
            sPath = files(files.Length - 1)
            Using sr As New StreamReader(sPath, Encoding.GetEncoding("Shift_JIS"))

                '�^�C�g���ǂݔ�΂�
                sInpData = sr.ReadLine()

                itemcnt = 0
                '�ŏI�s�܂łP�s���ƂɃt�@�C���Ǎ���
                Do While (False = sr.EndOfStream)
                    sInpData = sr.ReadLine()                                ' �P�s�Ǎ���
                    splt = sInpData.Split(","c)                             ' �J���}��؂�ŕ���

                    sLot = splt(ITEM_LOT_NUM)                               ' ���b�g�ԍ��̎擾 
                    nPlate = splt(ITEM_PLATE_NUM)                           ' ��ԍ��̎擾 
                    nBlockX = splt(ITEM_BLOCKX_NUM)                         ' Block�ԍ�X�̎擾 
                    nBlockY = splt(ITEM_BLOCKY_NUM)                         ' Block�ԍ�Y�̎擾 

                    '���b�g�ԍ��A��ԍ��ABlock�ԍ�X�AY����v����f�[�^�̂ݒ��o 
                    If ((sLot = sLotNumber) AndAlso (nPlate = PlateNumber) AndAlso (nBlockX = BlockX) AndAlso (nBlockY = BlockY)) Then
                        Target(itemcnt) = splt(ITEM_FT_NUM)              ' FT���ʂ̎擾 
                        itemcnt = itemcnt + 1
                        '���o�������ڐ��̐ݒ�
                        ResCounter = itemcnt
                    Else
                        If ResCounter > 0 Then
                            Exit Do
                        End If
                        itemcnt = 0
                    End If
                Loop
            End Using

            Return (True)
        Catch ex As Exception
            Call Z_PRINT("UserSub.GetTargerDataFromLogFile() TRAP ERROR = " & ex.Message & vbCrLf)
            Return (False)
        End Try
    End Function

    ' ��̃t�@�C���̍ŏI�������ݓ������擾���Ĕ�r���郁�\�b�h
    Private Function CompareLastWriteTime(ByVal fileX As String, ByVal fileY As String) As Integer
        Return DateTime.Compare(File.GetLastWriteTime(fileX), File.GetLastWriteTime(fileY))
    End Function

    ' ����l�ϓ����o�@�\
    Public bVariationMesStep As Boolean = True
    Public gVariationMeasPlateStartNo As Integer = 1
    Public gVariationMeasBlockXStartNo As Integer = 1
    Public gVariationMeasBlockYStartNo As Integer = 1
    Public dMeasVariationNOM(MAXRNO) As Double                  ' �g���~���O��e�s�l
    Public dMeasVariationDev(MAX_RES_USER) As Double            ' �ω���

    ''' <summary>
    ''' ����l�ϓ����o�@�\�J�n�u���b�N�ʒu������
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub VariationMesStartDataReset()
        Try
            gVariationMeasPlateStartNo = 1
            gVariationMeasBlockXStartNo = 1
            gVariationMeasBlockYStartNo = 1
            'V2.0.0.1�@            bVariationMesStep = True
            bVariationMesStep = False       'V2.0.0.1�@
        Catch ex As Exception
            Call Z_PRINT("UserSub.VariationMesStartDataReset() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub
    Public Function SetTarrgetOnVariationMeas() As Boolean
        Try
            Dim bRtn As Boolean
            Dim ResCounter As Integer
            Dim Target(MAXRNO) As Double
            Dim rno As Integer

            For rn As Short = 1 To stPLT.RCount
                dMeasVariationNOM(rn) = 0.0
            Next

            bRtn = GetTargerDataFromLogFile(stUserData.sLotNumber, stCounter.PlateCounter, stCounter.BlockCntX, stCounter.BlockCntY, ResCounter, Target)
            If Not bRtn Then
                Call DebugLogOut("����l�ϓ����o �ڕW�l�ݒ�G���[ LOT=[" & stUserData.sLotNumber & "] PLATE=[" & stCounter.PlateCounter.ToString & "]X=[" & stCounter.BlockCntX.ToString & "]Y=[" & stCounter.BlockCntY.ToString & "]")
                Return (False)
            End If
            Dim Rcnt As Integer = UserBas.GetRCountExceptMeasure()

            rno = 0
            For rn As Integer = 1 To stPLT.RCount
                If UserModule.IsCutResistor(rn) Then
                    If rno < ResCounter Then
                        dMeasVariationNOM(rn) = Target(rno)
                        rno = rno + 1
                    Else
                        Call Z_PRINT("����l�ϓ����o �ڕW�l�f�[�^���L��܂��� LOT=[" & stUserData.sLotNumber & "] PLATE=[" & stCounter.PlateCounter.ToString & "]X=[" & stCounter.BlockCntX.ToString & "]Y=[" & stCounter.BlockCntY.ToString & "]RES=[" & rn.ToString & "]")
                        Call DebugLogOut("����l�ϓ����o �ڕW�l�ݒ�G���[ LOT=[" & stUserData.sLotNumber & "] PLATE=[" & stCounter.PlateCounter.ToString & "]X=[" & stCounter.BlockCntX.ToString & "]Y=[" & stCounter.BlockCntY.ToString & "]]RES=[" & rn.ToString & "]")
                    End If
                End If
            Next
            Return (True)
        Catch ex As Exception
            Call Z_PRINT("UserSub.SetTarrgetOnVariationMeas() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function

    Public Function VariationMeasJudge(ByVal rn As Short, ByVal dblMx As Double) As Boolean
        Try
            dMeasVariationDev(rn) = 0.0

            If dMeasVariationNOM(rn) = 0.0 Then
                Return (False)
            End If

            ' �g���~���O�덷��Βl�@���@�i�@�g���~���O�l�@�|�@�g���~���O���̂s�e�l�@�j�^�g���~���O���̂s�e�l�@* 10^6
            dMeasVariationDev(rn) = FNDEVP(dblMx, dMeasVariationNOM(rn))

            If dMeasVariationDev(rn) > Math.Abs(stUserData.dVariation) Then
                Return (False)
            Else
                Return (True)
            End If

        Catch ex As Exception
            Call Z_PRINT("UserSub.VariationMeasJudge() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function

    'V2.0.0.0�A��

    'V2.0.0.0�I��
    ''' <summary>
    ''' �T�[�L�b�g�����̃J�E���g
    ''' </summary>
    ''' <param name="stPlate"></param>
    ''' <param name="stRegData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCircuitSum(ByVal stPlate As PLATE_DATA, ByRef stRegData As Reg_Info()) As Integer
        Try
            Dim iResCnt As Integer = 0
            Dim iCircuit As Integer = -1

            For rn As Integer = 1 To stPlate.RCount
                If UserModule.IsCutResistor(stRegData, rn) Then
                    If iCircuit <> stRegData(rn).intCircuitNo Then
                        iResCnt = iResCnt + 1
                    End If
                    iCircuit = stRegData(rn).intCircuitNo
                End If
            Next
            Return (iResCnt)

        Catch ex As Exception
            Call Z_PRINT("UserSub.GetCircuitSum() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function
    ''' <summary>
    ''' ��R�ԍ�����T�[�L�b�g�ԍ����擾����B
    ''' </summary>
    ''' <param name="stPlate"></param>
    ''' <param name="stRegData"></param>
    ''' <param name="rno"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCircuitNoFromResNo(ByVal stPlate As PLATE_DATA, ByRef stRegData As Reg_Info(), ByRef rno As Short) As Integer
        Try
            Dim iCircuitNo As Integer = 0
            Dim iCircuit As Integer = -1

            For rn As Integer = 1 To stPlate.RCount
                If UserModule.IsCutResistor(stRegData, rn) Then
                    If iCircuit <> stRegData(rn).intCircuitNo Then
                        iCircuitNo = iCircuitNo + 1
                    End If
                    iCircuit = stRegData(rn).intCircuitNo
                    If rn = rno Then
                        Return (iCircuitNo)
                    End If
                End If
            Next
        Catch ex As Exception
            Call Z_PRINT("UserSub.GetCircuitNoFromResNo() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function
    ''' <summary>
    ''' �T�[�L�b�g���̒�R���̃J�E���g
    ''' </summary>
    ''' <param name="stRegData">��R�f�[�^�\����</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CircuitResistorCount(ByVal stPlate As PLATE_DATA, ByRef stRegData As Reg_Info()) As Integer
        Try
            Dim iResCnt As Integer = 0
            Dim iCircuit As Integer = -1

            For rn As Integer = 1 To stPlate.RCount
                If UserModule.IsCutResistor(stRegData, rn) Then
                    If iCircuit < 0 Then
                        iCircuit = stRegData(rn).intCircuitNo
                    End If
                    If stRegData(rn).intCircuitNo <> iCircuit Then
                        Return (iResCnt)
                    End If
                    iResCnt = iResCnt + 1
                End If
            Next
            Return (iResCnt)

        Catch ex As Exception
            Call Z_PRINT("UserSub.CircuitResistorCount() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function

    ''' <summary>
    ''' �T�[�L�b�g���̒�R���̃J�E���g
    ''' </summary>
    ''' <returns>����T�[�L�b�g����R��</returns>
    ''' <remarks></remarks>
    Public Function CircuitResistorCount() As Integer
        Try
            Return (CircuitResistorCount(stPLT, stREG))

        Catch ex As Exception
            Call Z_PRINT("UserSub.CircuitResistorCount() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function

    ''' <summary>
    ''' �T�[�L�b�g�̍Ō�̒�R�����`�F�b�N����B
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsCheckCircuitEnd(ByVal rn As Short) As Boolean
        Try
            Dim iCircuit As Integer = stREG(rn).intCircuitNo

            For i As Integer = (rn + 1) To stPLT.RCount
                If UserModule.IsCutResistor(i) Then
                    If iCircuit = stREG(i).intCircuitNo Then    ' ��ɓ�����R���o�ė�����Ō�ł͖���
                        Return (False)
                    Else
                        Return (True)
                    End If
                End If
            Next

            Return (True)

        Catch ex As Exception
            Call Z_PRINT("UserSub.IsCheckCircuitEnd() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Function
    ''' <summary>
    ''' ���݂̒�R�ԍ����Ō�̃T�[�L�b�g���𔻒肷��B
    ''' </summary>
    ''' <param name="rn">��R�ԍ�</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsCheckLastCircuit(ByVal rn As Short) As Boolean
        Try
            Dim CircuitCnt As Integer = GetCircuitSum(stPLT, stREG)
            Dim CircuitNO As Integer = GetCircuitNoFromResNo(stPLT, stREG, rn)

            If CircuitNO = CircuitCnt Then
                Return (True)
            Else
                Return (False)
            End If

        Catch ex As Exception
            Call Z_PRINT("UserSub.IsCheckLastCircuit() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function
    ''' <summary>
    ''' �����T�[�L�b�g�ԍ��̐擪�̒�R�ԍ������߂�B
    ''' </summary>
    ''' <param name="rn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetTopResNoinCircuit(ByVal rn As Short) As Short
        Try
            Dim iCircuit As Integer = stREG(rn).intCircuitNo
            Dim TopResNo As Short = rn
            Dim ResNo As Short

            For ResNo = TopResNo To 1 Step -1
                If UserModule.IsCutResistor(ResNo) Then
                    If iCircuit = stREG(ResNo).intCircuitNo Then    ' �O�ɓ����T�[�L�b�g�ԍ����o�ė�����
                        TopResNo = ResNo
                    End If
                End If
            Next

            Return (TopResNo)

        Catch ex As Exception
            Call Z_PRINT("UserSub.GetTopResNoinCircuit() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function

    ''' <summary>
    ''' �T�[�L�b�g���ŉ��Ԗڂ̒�R�������߂�
    ''' </summary>
    ''' <param name="stRegData"></param>
    ''' <param name="rn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetResNumberInCircuit(ByRef stRegData As Reg_Info(), ByVal rn As Short) As Integer
        Try

            Dim iCircuit As Integer = stRegData(rn).intCircuitNo
            Dim iNumber As Integer = 0

            If UserSub.IsTrimType1() Or UserSub.IsTrimType4() Then  ' ���x�Z���T�[�̎��͂P�Œ�
                Return (1)
            End If

            For i As Integer = 1 To rn
                If UserModule.IsCutResistor(stRegData, i) Then
                    If iCircuit = stRegData(i).intCircuitNo Then    ' �����T�[�L�b�g�ԍ�
                        iNumber = iNumber + 1
                    End If
                End If
            Next

            Return (iNumber)

        Catch ex As Exception
            Call Z_PRINT("UserSub.GetResNumberInCircuit() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function

    ''' <summary>
    ''' �T�[�L�b�g���ŉ��Ԗڂ̒�R�������߂�
    ''' </summary>
    ''' <param name="rn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetResNumberInCircuit(ByVal rn As Short) As Integer
        Try

            Return (GetResNumberInCircuit(stREG, rn))

        Catch ex As Exception
            Call Z_PRINT("UserSub.GetResNumberInCircuit() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function

    ''' <summary>
    ''' �T�[�L�b�g���̏��Ԃ̒�R�ԍ������߂�
    ''' </summary>
    ''' <param name="stPlate"></param>
    ''' <param name="stRegData"></param>
    ''' <param name="Circuit"></param>
    ''' <param name="No"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRNumByCircuit(ByVal stPlate As PLATE_DATA, ByRef stRegData As Reg_Info(), ByVal Circuit As Short, ByVal No As Short) As Integer
        Try
            Dim iResCnt As Integer = 0
            Dim iCircuit As Integer = 0

            For rno As Integer = 1 To stPlate.RCount
                If UserModule.IsCutResistor(stRegData, rno) Then
                    If iCircuit <> stRegData(rno).intCircuitNo Then
                        iCircuit = iCircuit + 1
                    End If
                    If iCircuit = Circuit Then
                        iResCnt = iResCnt + 1
                        If iResCnt = No Then
                            Return (rno)
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            Call Z_PRINT("UserSub.GetRNumByCircuit() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function
    Public Function GetRNumByCircuit(ByVal Circuit As Short, ByVal No As Short) As Integer
        Try
            Return (GetRNumByCircuit(stPLT, stREG, Circuit, No))
        Catch ex As Exception
            Call Z_PRINT("UserSub.GetRNumByCircuit() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function
    'V2.0.0.0�I��

    'V2.0.0.0�M��
    ''' <summary>
    ''' �N�����v�z���ύX
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ClampVacumeChange()
        Try
            Dim r As Short

            If giLoaderType <> 0 Then   '�N�����v�z������ݒ�
                ObjSys.setClampVaccumConfig(stUserData.intClampVacume - 1)
            End If

            Select Case (stUserData.intClampVacume)
                Case CLAMP_VACCUME_USE
                Case CLAMP_ONLY_USE
                    Call Form1.System1.AbsVaccume(gSysPrm, 0, giAppMode, 0)
                    Call Form1.System1.Adsorption(gSysPrm, 0)
                Case VACCUME_ONLY_USE
                    r = Form1.System1.ClampCtrl(gSysPrm, 0, 0, False)                 ' �N�����v/�z��OFF
                    If (r = cFRS_NORMAL) Then
                        'Call Sub_ATLDSET(COM_STS_CLAMP_ON, 0)                       ' ���[�_�[�o��(ON=�ڕ������ߊJ,OFF=�Ȃ�)
                        'gbClampOpen = False
                    Else
                        Call Z_PRINT("�N�����v�J�G���[���������܂����B�B" & vbCrLf)
                    End If
                    'If giLoaderType = 1 Then
                    '    Call Form1.System1.AbsVaccume(gSysPrm, 1, giAppMode, 0)
                    'End If
                Case Else
                    Throw New Exception("Case " & stUserData.intClampVacume & ": Nothing")
            End Select
        Catch ex As Exception
            Call Z_PRINT("UserSub.ClampVacumeChange() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Sub
    'V2.0.0.0�M��

    'V2.0.0.1�B��
#Region "�������������̔���"
    ''' <summary>
    ''' �������������̔���(�m�f���䗦����j
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function PlateNGJudgeByCounter() As Boolean
        Dim dDrate As Double
        Try
            If UserBas.stCounter.TrimCounter = 0 Then
                Return (False)
            End If

            dDrate = UserBas.stCounter.NG_Counter / UserBas.stCounter.TrimCounter * 100.0

            If dDrate >= stUserData.NgJudgeRate Then
                DebugLogOut("�P��m�f����[" & dDrate.ToString & "]=[" & UserBas.stCounter.NG_Counter.ToString & "]/[" & UserBas.stCounter.TrimCounter.ToString & "] * 100.0 >= [" & stUserData.NgJudgeRate.ToString & "]")
                Z_PRINT("��m�f���� �䗦[" & dDrate.ToString & "]=[" & UserBas.stCounter.NG_Counter.ToString & "]/[" & UserBas.stCounter.TrimCounter.ToString & "] * 100.0 >= [" & stUserData.NgJudgeRate.ToString & "]")
                Return (True)
            Else
                DebugLogOut("�P��n�j����[" & dDrate.ToString & "]=[" & UserBas.stCounter.NG_Counter.ToString & "]/[" & UserBas.stCounter.TrimCounter.ToString & "] * 100.0 < [" & stUserData.NgJudgeRate.ToString & "]")
                Return (False)
            End If

        Catch ex As Exception
            MsgBox("PlateNGJudgeByCounter() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Function
#End Region
    'V2.0.0.1�B��

#Region "'V2.1.0.0�@�A�B -------------2019/9/20 �@�\�ǉ�"
    'V2.1.0.0�@��
#Region "�J�b�g���̒�R�l�ω��ʔ���@�\"
    ''' <summary>
    ''' �ϐ���`
    ''' </summary>
    ''' <remarks></remarks>
    Private CutMeasureBefore As Double = Double.MinValue    ' �J�b�g�O����l
    Private CutMeasureAfter As Double = Double.MinValue     ' �J�b�g�㑪��l
    Private bVariationDone As Boolean = False               ' �J�b�g��̕ω��ʌv�Z�ς�=True,���v�Z=False
    Private bBeforeMeasureReadDone As Boolean = False       ' �^�N�g�A�b�v�ׂ̈ɃJ�b�g�O����l���̃J�b�g�̏�������l�Ɏg�p'V2.1.0.0�D
    Private iVariationCutNGCutNo As Integer = 0             ' �J�b�g���̒�R�l�ω��ʔ���G���[�J�b�g�ԍ��A�����l�O
    Private dSavedVariationRate As Double                   ' �J�b�g���̒�R�l�ω��ʕۑ��p
    Private bCutVariationJudgeExecute As Boolean = False    ' �J�b�g���̒�R�l�ω��ʔ���L��
    Private bVariationNGHiorLow As Boolean = True           ' �J�b�g���̒�R�l�ω��ʂm�f��ʂk�n�FTrue�@�g�h�FFalse'V2.1.0.0�D
    Private bCutVariationCutDone As Boolean = False         ' �J�b�g���̒�R�l�ω��ʃJ�b�g�L���True'V2.1.0.0�D

    ''' <summary>
    ''' �J�b�g���̒�R�l�ω��ʔ���@�\�����f�[�^�̐ݒ�
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CutVariationDataInitialize()
        Try
            Dim dBefore As Double
            Dim dAfter As Double
            Dim dTargetCoffJudge As Double

            For rn As Integer = 1 To stPLT.RCount
                dTargetCoffJudge = stUserData.dTargetCoffJudge(UserSub.GetResNumberInCircuit(rn))
                For CutNo As Short = 1 To stREG(rn).intTNN
                    If CutNo = 1 Then
                        dBefore = dTargetCoffJudge
                    Else
                        dBefore = dAfter
                    End If
                    dAfter = dBefore + stREG(rn).STCUT(CutNo).dblCOF

                    stREG(rn).STCUT(CutNo).iVariationRepeat = 0             ' ���s�[�g�L��
                    stREG(rn).STCUT(CutNo).iVariation = 0                   ' ����L��
                    If UserModule.IsCutResistor(stREG, rn) Then
                        stREG(rn).STCUT(CutNo).dRateOfUp = (dAfter - dBefore) / dTargetCoffJudge * 100      ' �㏸��
                    Else
                        stREG(rn).STCUT(CutNo).dRateOfUp = 0.0                                              ' �㏸��
                    End If
                    stREG(rn).STCUT(CutNo).dVariationLow = -1.0             ' �����l
                    stREG(rn).STCUT(CutNo).dVariationHi = 1.0               ' ����l
                Next
            Next
        Catch ex As Exception
            MsgBox("CutVariationDataInitialize() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' �J�b�g���̒�R�l�ω��ʔ���̃��s�[�g�L�ւ̑S�R�s�[����
    ''' </summary>
    ''' <param name="stPlate">�v���[�g�f�[�^�\����</param>
    ''' <param name="stRegData">��R�f�[�^�\����</param>
    ''' <param name="ResNo">�R�s�[����R�ԍ�</param>
    ''' <param name="CutNo">�R�s�[���J�b�g�ԍ�</param>
    ''' <remarks></remarks>
    Public Sub CutVariationDataCopy(ByVal stPlate As PLATE_DATA, ByRef stRegData As Reg_Info(), ByVal ResNo As Integer, ByVal CutNo As Integer)
        Try
            Dim OrderNo As Integer

            OrderNo = UserSub.GetResNumberInCircuit(stRegData, ResNo)               ' �T�[�L�b�g���̒�R�̏���

            For rn As Integer = 1 To stPlate.RCount
                If rn = ResNo Then                                                  ' �R�s�[���̓X�L�b�v����B
                    Continue For
                End If
                If UserModule.IsCutResistor(stRegData, rn) Then
                    If OrderNo = UserSub.GetResNumberInCircuit(stRegData, rn) Then                                      ' �T�[�L�b�g���̓�����R����
                        stRegData(rn).STCUT(CutNo).iVariationRepeat = stRegData(ResNo).STCUT(CutNo).iVariationRepeat    ' ���s�[�g�L��
                        stRegData(rn).STCUT(CutNo).iVariation = stRegData(ResNo).STCUT(CutNo).iVariation                ' ����L��
                        stRegData(rn).STCUT(CutNo).dRateOfUp = stRegData(ResNo).STCUT(CutNo).dRateOfUp                  ' �㏸��
                        stRegData(rn).STCUT(CutNo).dVariationLow = stRegData(ResNo).STCUT(CutNo).dVariationLow          ' �����l
                        stRegData(rn).STCUT(CutNo).dVariationHi = stRegData(ResNo).STCUT(CutNo).dVariationHi            ' ����l
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox("CutVariationDataCopy() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' �g���~���O�f�[�^�P�ʂł̒�R�l�ω��ʔ���L���m�F
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsCutVariationJudgeExecute() As Boolean
        Return (bCutVariationJudgeExecute)
    End Function

    ''' <summary>
    ''' �g���~���O�f�[�^�P�ʂł̒�R�l�ω��ʔ���L���`�F�b�N
    ''' </summary>
    ''' <remarks></remarks>
    Public Function CutVariationJudgeExecuteCheck() As Boolean
        Try
            bCutVariationJudgeExecute = False

            If (DGL = TRIM_MODE_ITTRFT) AndAlso IsSpecialTrimType() Then
                For ResNo As Integer = 1 To stPLT.RCount
                    If IsCutResistor(ResNo) Then
                        For CutNo As Integer = 1 To stREG(ResNo).intTNN
                            If stREG(ResNo).STCUT(CutNo).iVariation = 1 Then
                                bCutVariationJudgeExecute = True
                                Exit For
                            End If
                        Next
                    End If
                Next
            End If
            Return (bCutVariationJudgeExecute)
        Catch ex As Exception
            MsgBox("CutVariationJudgeExecuteCheck() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' ��R�l�ω��ʔ���E����p�ڕW�l�Z�o�W������̃J�b�g�O��R�l�̕ۑ�
    ''' </summary>
    ''' <param name="ResNo"></param>
    ''' <remarks>�P��R�J�b�g�J�n�O�̏���������</remarks>
    Public Sub CutVariationInitialize(ByVal ResNo As Integer)
        Try

            iVariationCutNGCutNo = 0                ' �J�b�g���̒�R�l�ω��ʔ���G���[�J�b�g�ԍ�������

            If IsCutVariationJudgeExecute() AndAlso IsCutResistor(ResNo) Then
                Call CutVariationDebugLogOut("��R�l�ω��ʔ��菉����RES=[" & ResNo.ToString("0") & "]")
                bVariationDone = False              ' �J�b�g��̕ω��ʌv�Z�ς�=True,���v�Z=False
                ' ����p�ڕW�l�Z�o�W���̃J�b�g�O��R�l�ۑ�
                CutVariationMeasureBeforeSet(UserSub.GetInitialResValue())
            End If

        Catch ex As Exception
            MsgBox("CutVariationInitialize() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try

    End Sub

    'V2.1.0.0�D��
    ''' <summary>
    ''' ��R�l�ω��ʔ���E�J�b�g�L��ɂ���B
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CutVariationCutSet()
        bCutVariationCutDone = True
        Call CutVariationDebugLogOut("�J�b�g����R�l�ω��ʃJ�b�g�L��ɃZ�b�g")
    End Sub

    Private Function CutVariationCutDone() As Boolean
        Return (bCutVariationCutDone)
    End Function
    'V2.1.0.0�D��
    ''' <summary>
    ''' ��R�l�ω��ʔ���E�J�b�g��̕ω��ʖ��v�Z��Ԃɂ���B
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CutVariationInitByCut()
        bVariationDone = False                          ' �J�b�g��̕ω��ʌv�Z�ς�=True,���v�Z=False
        CutMeasureAfter = Double.MinValue               'V2.1.0.0�D �J�b�g�㑪��l
        bCutVariationCutDone = False                   'V2.1.0.0�D
        Call CutVariationDebugLogOut("�J�b�g����R�l�ω��ʃJ�b�g�L�菉����")
    End Sub

    ''' <summary>
    ''' ��R�l�ω��ʔ��茋�ʎ擾
    ''' </summary>
    ''' <returns>OK:False,NG:True</returns>
    ''' <remarks></remarks>
    Public Function CutVariationFinalJudgeNG() As Boolean
        If iVariationCutNGCutNo > 0 Then
            Return (True)
        Else
            Return (False)
        End If
    End Function

    ''' <summary>
    ''' ��R�l�ω��ʔ���p�J�b�g�O��R�l��ۑ�����B
    ''' </summary>
    ''' <param name="dMeasure">��R�l</param>
    ''' <remarks></remarks>
    Private Sub CutVariationMeasureBeforeSet(ByVal dMeasure As Double)
        CutMeasureBefore = dMeasure
        If CutMeasureBefore = Double.MinValue Then
            Call CutVariationDebugLogOut("�J�b�g�O��R�l������")
        Else
            Call CutVariationDebugLogOut("�J�b�g�O��R�l�ۑ��l=[" & CutMeasureBefore.ToString & "]")
            bBeforeMeasureReadDone = True
        End If
    End Sub

    'V2.1.0.0�D��
    ''' <summary>
    ''' �J�b�g�O��R�l�̎擾
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CutVariationMeasureBeforeGet() As Double
        Return (CutMeasureBefore)
    End Function

    ''' <summary>
    ''' ��R�l�ω��ʔ���p�J�b�g�O��R�l�ۑ����̊m�F�E��x�ǂݏo������I�t����B
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsCutMeasureBefore() As Boolean
        If bBeforeMeasureReadDone Then
            bBeforeMeasureReadDone = False
            Return (True)
        Else
            Return (False)
        End If
    End Function
    'V2.1.0.0�D��

    ''' <summary>
    ''' ��R�l�ω��ʔ���p�J�b�g���R�l�ۑ�
    ''' </summary>
    ''' <param name="dMeasure">��R�l</param>
    ''' <remarks></remarks>
    Public Sub CutVariationMeasureAfterSet(ByVal dMeasure As Double)
        CutMeasureAfter = dMeasure
        If CutMeasureAfter = Double.MinValue Then
            Call CutVariationDebugLogOut("�J�b�g���R�l������")
        Else
            Call CutVariationDebugLogOut("�J�b�g���R�l�ۑ��l=[" & CutMeasureAfter.ToString & "]")
        End If
    End Sub

    'V2.1.0.0�D��
    ''' <summary>
    ''' ��R�l�ω��ʔ���p�J�b�g���R�l���ۑ����̊m�F
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsNotCutMeasureAfter() As Boolean
        If CutMeasureAfter = Double.MinValue Then
            Return (True)
        Else
            Return (False)
        End If
    End Function

    ''' <summary>
    ''' �J�b�g���̒�R�l�ω��ʔ���G���[�J�b�g�ԍ�������
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub VariationCutNGCutNoReset()
        iVariationCutNGCutNo = 0
    End Sub
    'V2.1.0.0�D��

    ''' <summary>
    ''' �J�b�g���̒�R�l�ω��ʔ���
    ''' </summary>
    ''' <param name="ResNo">��R�ԍ�</param>
    ''' <param name="CutNo">�J�b�g�ԍ�</param>
    ''' <returns>OK:True,NG:False</returns>
    ''' <remarks></remarks>
    Public Function CutVariationJudge(ByVal ResNo As Integer, ByVal CutNo As Integer) As Boolean
        Try
            Dim sRtn As Short

            If IsCutVariationJudgeExecute() AndAlso IsCutResistor(ResNo) AndAlso (bVariationDone = False) Then

                iVariationCutNGCutNo = 0

                If IsNotCutMeasureAfter() = True Then   ' �J�b�g���R�l������
                    If UserSub.IsSpecialTrimType Then
                        sRtn = V_R_MEAS(stREG(ResNo).intSLP, stREG(ResNo).intMType, CutMeasureAfter, ResNo, UserSub.GetTRV())
                    Else
                        sRtn = V_R_MEAS(stREG(ResNo).intSLP, stREG(ResNo).intMType, CutMeasureAfter, ResNo, stREG(ResNo).dblNOM)
                    End If
                    If sRtn = cFRS_NORMAL Then
                        Call CutVariationDebugLogOut("�J�b�g���R�l�����莞����(CutVariationJudge) RES=[" & ResNo.ToString & "] CUT=[" & CutNo.ToString & "] ����l=[" & CutMeasureAfter.ToString & "]")
                    Else
                        Call Z_PRINT("�J�b�g���̒�R�l�ω��ʔ��� RES=[" & ResNo.ToString & "] CUT=[" & CutNo.ToString & "] ����G���[=[" & sRtn.ToString & "]")
                        Call DebugLogOut("�J�b�g���̒�R�l�ω��ʔ���(JudgeVariationByCut) RES=[" & ResNo.ToString & "] CUT=[" & CutNo.ToString & "] ����G���[=[" & sRtn.ToString & "]")
                        Call CutVariationDebugLogOut("�J�b�g���̒�R�l�ω��ʔ���(JudgeVariationByCut) RES=[" & ResNo.ToString & "] CUT=[" & CutNo.ToString & "] ����G���[=[" & sRtn.ToString & "]")
                        iVariationCutNGCutNo = CutNo
                        Return (False)
                    End If

                End If

                ' ���肪�����Ă�����L��̃J�b�g�ׂ̈Ɍv�Z�͕K�v�B

                Dim dTargetCoffJudge As Double = UserSub.GetInitialResValue()
                '�㏸���@���@(�J�b�g���R�l�@�|�J�b�g�O��R�l)/����p�ڕW�l�Z�o�W���@�~�@�P�O�O
                dSavedVariationRate = (CutMeasureAfter - CutVariationMeasureBeforeGet()) / dTargetCoffJudge * 100
                CutVariationDebugLogOut("�㏸��[" & dSavedVariationRate.ToString & "] = (�J�b�g���R�l[" & CutMeasureAfter.ToString & "]-�J�b�g�O��R�l[" & CutVariationMeasureBeforeGet().ToString & "])/��������l[" & dTargetCoffJudge.ToString & "]*100")

                Dim dLo As Double = stREG(ResNo).STCUT(CutNo).dRateOfUp + stREG(ResNo).STCUT(CutNo).dVariationLow
                Dim dHi As Double = stREG(ResNo).STCUT(CutNo).dRateOfUp + stREG(ResNo).STCUT(CutNo).dVariationHi

                CutVariationMeasureBeforeSet(CutMeasureAfter)               ' ���̃J�b�g�ׂ̈ɃJ�b�g�O��R�l�Ɍ��݂̃J�b�g���R�l��ۑ����ē���ւ���B
                CutVariationMeasureAfterSet(Double.MinValue)                ' ��������Ԃɂ���B
                bVariationDone = True                                       ' �J�b�g��̕ω��ʌv�Z�ςݏ�Ԃɂ���B

                ' �J�b�g���̒�R�l�ω��ʔ��肪�L��̎��ŁA�㉺���l����O��Ă��鎞�́A�J�b�g�ԍ���ۑ����ăG���[���^�[������B
                If CutVariationCutDone() AndAlso stREG(ResNo).STCUT(CutNo).iVariation = 1 AndAlso (dSavedVariationRate < dLo OrElse dHi < dSavedVariationRate) Then
                    DebugLogOut("�㏸������NG ��RNO=[" & ResNo.ToString & "] �J�b�gNO=[" & CutNo.ToString & "] ����[" & dLo.ToString & "] <= �㏸��[" & dSavedVariationRate.ToString & "] <= ���[" & dHi.ToString & "]")
                    CutVariationDebugLogOut("�㏸������NG ��RNO=[" & ResNo.ToString & "] �J�b�gNO=[" & CutNo.ToString & "] ����[" & dLo.ToString & "] <= �㏸��[" & dSavedVariationRate.ToString & "] <= ���[" & dHi.ToString & "]")
                    iVariationCutNGCutNo = CutNo
                    'V2.1.0.0�D��
                    If dSavedVariationRate < dLo Then
                        bVariationNGHiorLow = True
                    Else
                        bVariationNGHiorLow = False
                    End If
                    'V2.1.0.0�D��
                    Return (False)
                Else
                    If Not CutVariationCutDone() Then
                        CutVariationDebugLogOut("�㏸�����薳(�J�b�g����) ��RNO=[" & ResNo.ToString & "] �J�b�gNO=[" & CutNo.ToString & "] ����[" & dLo.ToString & "] <= �㏸��[" & dSavedVariationRate.ToString & "] <= ���[" & dHi.ToString & "]")
                    ElseIf stREG(ResNo).STCUT(CutNo).iVariation = 1 Then
                        CutVariationDebugLogOut("�㏸������OK ��RNO=[" & ResNo.ToString & "] �J�b�gNO=[" & CutNo.ToString & "] ����[" & dLo.ToString & "] <= �㏸��[" & dSavedVariationRate.ToString & "] <= ���[" & dHi.ToString & "]")
                    Else
                        CutVariationDebugLogOut("�㏸�����薳 ��RNO=[" & ResNo.ToString & "] �J�b�gNO=[" & CutNo.ToString & "] ����[" & dLo.ToString & "] <= �㏸��[" & dSavedVariationRate.ToString & "] <= ���[" & dHi.ToString & "]")
                    End If
                    Return (True)
                End If
            End If
            Return (True)

        Catch ex As Exception
            MsgBox("CutVariationJudge() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (False)
        End Try

    End Function

    'V2.1.0.0�D��
    ''' <summary>
    ''' �J�b�g���̒�R�l�ω��ʔ���m�f����HI,LO��ʔ��f
    ''' </summary>
    ''' <returns>True:Lo,False�FHI</returns>
    ''' <remarks></remarks>
    Public Function GetVariationNGHiorLow() As Boolean
        Return (bVariationNGHiorLow)
    End Function
    'V2.1.0.0�D��

    ''' <summary>
    ''' �J�b�g���̒�R�l�ω��ʔ���E�m�f���̃J�b�g�ԍ��擾
    ''' </summary>
    ''' <returns>�m�f���̃J�b�g�ԍ�</returns>
    ''' <remarks></remarks>
    Public Function CutVariationCutNoGet() As Double
        Return (iVariationCutNGCutNo)
    End Function

    ''' <summary>
    ''' �J�b�g���̒�R�l�ω��ʔ���E�m�f���̏㏸���擾
    ''' </summary>
    ''' <returns>�m�f���̏㏸��</returns>
    ''' <remarks></remarks>
    Public Function CutVariationRateGet() As Double
        Return (dSavedVariationRate)
    End Function
#End Region
    'V2.1.0.0�@��

    'V2.1.0.0�A�B��
#Region "�A�b�e�l�[�^�e�[�u���A���x�Z���T�[���e�[�u���֘A"
    ''' <summary>
    ''' �P�ԐV�����t�@�C�����擾����B
    ''' </summary>
    ''' <param name="sPath">�t�H���_</param>
    ''' <param name="sHeader">�t�@�C�����̃w�b�_"",""</param>
    ''' <param name="sGetFileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetNewestFile(ByVal sPath As String, ByVal sHeader As String, ByRef sGetFileName As String) As Boolean
        Try
            Dim sFileName As String
            Dim sExtension As String
            Dim sFileList As String() = System.IO.Directory.GetFiles(sPath)
            Dim Year As Integer, Month As Integer, Day As Integer
            Dim Today As DateTime = System.DateTime.Today

            If sFileList.Length = 0 Then
                Call Z_PRINT("�t�H���_[" & sPath & "]�Ƀt�@�C�������݂��܂���B")
                Return (False)
            End If

            Array.Sort(sFileList)
            sGetFileName = sFileList(sFileList.Length - 1)
            ' �Ó����`�F�b�N
            For i As Integer = sFileList.Length - 1 To 0 Step -1
                sExtension = System.IO.Path.GetExtension(sFileList(i))                                                      ' �g���q�擾
                If sExtension.Equals(".CSV", StringComparison.OrdinalIgnoreCase) Then                                       ' �g���q����v����t�@�C�����Ώ�
                    sFileName = System.IO.Path.GetFileNameWithoutExtension(sFileList(i))                                    ' �g���q�������t�@�C����
                    If sFileName.Length = (sHeader.Length + 8) Then                                                         ' YYYYMMDD�̂W����
                        If sFileName.Substring(0, sHeader.Length).Equals(sHeader, StringComparison.OrdinalIgnoreCase) Then  ' �t�@�C�����̃^�C�g��������v����
                            Year = Integer.Parse(sFileName.Substring(sHeader.Length, 4))                                    ' �N
                            Month = Integer.Parse(sFileName.Substring(sHeader.Length + 4, 2))                               ' ��
                            Day = Integer.Parse(sFileName.Substring(sHeader.Length + 6, 2))                                 ' ��
                            Dim FileDate As New DateTime(Year, Month, Day)                                                  ' �t�@�C���̔N������DateTime�^�ɕϊ�
                            If FileDate.Date.CompareTo(Today.Date) <= 0 Then                                                ' �����܂ł̓��t��ΏۂƂ���B
                                sGetFileName = sFileList(i)                                                                 ' �����A�����̓��t���L���Ă��̓��ɂȂ�ƑΏۂɓ����Ă��܂��B
                                Return (True)
                            End If
                        End If
                    End If
                End If
            Next
            Return (False)
        Catch ex As Exception
            MsgBox("GetNewestFile() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (False)
        End Try
    End Function
#End Region
    'V2.1.0.0�A�B��

    'V2.1.0.0�A��
#Region "���[�U�p���[���j�^�����O"
    ''' <summary>
    ''' ���[�U�p���[���j�^�����O���[�h�Ǘ��ϐ�
    ''' </summary>
    ''' <remarks></remarks>
    Public Const POWER_CHECK_NONE As Short = 0
    Public Const POWER_CHECK_START As Short = 1
    Public Const POWER_CHECK_LOT As Short = 2
    Private gbLaserCaribrarionUse As Boolean = False
    Private bLaserCalibrationExecute As Boolean = False
    Private giLaserCalibrationMode As Integer = 0

    ''' <summary>
    ''' ���[�U�p���[���j�^�����O�g�p�L�薳��
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsLaserCaribrarionUse() As Boolean
        Return (gbLaserCaribrarionUse)
    End Function

    ''' <summary>
    ''' ���[�U�p���[���j�^�����O���[�h�擾
    ''' </summary>
    ''' <returns>0:POWER CHECK �Ȃ�,1:POWER CHECK �����^�]�J�n��,2:POWER CHECK ���b�g��</returns>
    ''' <remarks></remarks>
    Public Function LaserCalibrationModeGet() As Integer
        Return (giLaserCalibrationMode)
    End Function

    ''' <summary>
    ''' ���[�U�p���[���j�^�����O���s���[�h�ݒ�
    ''' </summary>
    ''' <param name="Mode">POWER_CHECK_NONE,POWER_CHECK_START,POWER_CHECK_LOT</param>
    ''' <remarks></remarks>
    Public Sub LaserCalibrationModeSet(ByVal Mode As Integer)
        Try
            If IsLaserCaribrarionUse() Then
                If Mode = POWER_CHECK_NONE Then                 ' �Ȃ��ɕύX��������s�t���O�������ɂ���B
                    bLaserCalibrationExecute = False
                End If
                If giLaserCalibrationMode <> POWER_CHECK_LOT AndAlso Mode = POWER_CHECK_LOT Then
                    If pbLoadFlg = True Then
                        bLaserCalibrationExecute = True         ' ���b�g���ɕύX�����烍�[�h�ς݂Ȃ���s�L��ɂ���B
                    End If
                End If
                giLaserCalibrationMode = Mode
                WritePrivateProfileString("LASER", "LASER_CALIBRATION_MODE", giLaserCalibrationMode.ToString("0"), USER_SYSPARAMPATH)
            End If
        Catch ex As Exception
            MsgBox("LaserCalibrationModeSet() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Sub
    Public Sub LaserCalibrationModeUpdate()
        Try
            Select Case (giLaserCalibrationMode)
                Case POWER_CHECK_NONE
                    Form1.ButtonLaserCalibration.Text = "POWER CHECK �Ȃ�"
                    Form1.ButtonLaserCalibration.BackColor = SystemColors.Control
                Case POWER_CHECK_START
                    Form1.ButtonLaserCalibration.Text = "POWER CHECK �����^�]�J�n��"
                    Form1.ButtonLaserCalibration.BackColor = System.Drawing.Color.LightSkyBlue
                Case POWER_CHECK_LOT
                    Form1.ButtonLaserCalibration.Text = "POWER CHECK ���b�g��"
                    Form1.ButtonLaserCalibration.BackColor = System.Drawing.Color.LightPink
            End Select

            Form1.ButtonLaserCalibration.Refresh()

        Catch ex As Exception
            MsgBox("LaserCalibrationModeUpdate() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' ���[�U�v���O�����N�������[�U�p���[���j�^�����O���[�h�ݒ�
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub LaserCalibrationModeLoad()
        Try

            If gSysPrm.stIOC.giPM_Tp = 1 Then   ' �p���[���[�^�̑����^�C�v(0:�Ȃ�(��u��), 1:�X�e�[�W�ݒu�^�C�v, 2:�X�e�[�W�O�ݒu�^�C�v)
                gbLaserCaribrarionUse = True

                Dim LaserCalibrationMode As Integer = Integer.Parse(GetPrivateProfileString_S("LASER", "LASER_CALIBRATION_MODE", USER_SYSPARAMPATH, "0"))

                UserSub.LaserCalibrationModeSet(LaserCalibrationMode)

                UserSub.LaserCalibrationModeUpdate()
            Else
                gbLaserCaribrarionUse = False

                Form1.cmdLaserCalibration.Enabled = False
                Form1.cmdLaserCalibration.Visible = False

                Form1.ButtonLaserCalibration.Enabled = False
                Form1.ButtonLaserCalibration.Visible = False

                Form1.cmdLaserTeach.Width = Form1.CmdTx.Width
            End If

        Catch ex As Exception
            MsgBox("LaserCalibrationModeLoad() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' ���[�U�p���[���j�^�����O���s�L���ݒ�
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub LaserCalibrationSet(ByVal Mode As Integer)
        Try
            If IsLaserCaribrarionUse() Then
                Select Case (Mode)
                    Case POWER_CHECK_NONE
                        bLaserCalibrationExecute = False
                    Case POWER_CHECK_START
                        If LaserCalibrationModeGet() = POWER_CHECK_START Then
                            bLaserCalibrationExecute = True
                        End If
                    Case POWER_CHECK_LOT
                        If LaserCalibrationModeGet() = POWER_CHECK_LOT Then
                            bLaserCalibrationExecute = True
                        End If
                End Select
            Else
                bLaserCalibrationExecute = False
            End If
        Catch ex As Exception
            MsgBox("LaserCalibrationSet() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' ���[�U�p���[���j�^�����O���s�L��
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>�t�@���N�V�������ŌĂ΂ꂽ����s�����ɕύX����</remarks>
    Public Function LaserCalibrationExecute() As Boolean
        Try
            Dim bRtn As Boolean = False

            If IsLaserCaribrarionUse() Then
                If DGL = TRIM_MODE_ITTRFT OrElse DGL = TRIM_MODE_CUT OrElse DGL = TRIM_MODE_MEAS_MARK OrElse DGL = TRIM_VARIATION_MEAS Then
                    bRtn = bLaserCalibrationExecute
                    bLaserCalibrationExecute = False
                End If
            End If

            Return (bRtn)
        Catch ex As Exception
            MsgBox("LaserCalibrationExecute() Execption error." & vbCrLf & " error msg = " & ex.Message)
        End Try

    End Function

    ''' <summary>
    ''' �A�b�e�l�[�^�e�[�u������f�[�^�擾
    ''' </summary>
    ''' <param name="No">�ԍ�</param>
    ''' <param name="stData">�f�[�^�i�[�\����</param>
    ''' <param name="MaxNo">�ԍ���99�ȏ��ݒ肵�����ő�ԍ�</param>
    ''' <returns>True:�����ԍ��Ɉ�v�������̂�</returns>
    ''' <remarks></remarks>
    Private Function LaserCalibrationAttenuatorTableGet(ByVal No As Integer, ByRef stData As stATTENUATOR_TABLE, ByRef MaxNo As Integer) As Boolean
        Try
            Dim sFolder As String = vbNullString
            Dim sData As String
            Dim mData() As String
            Dim TableNo As Integer
            Dim bHeader As Boolean = True
            Dim dData As Double

            MaxNo = 0

            If Not GetNewestFile(cLASERPOWER_PATH, cLASERPOWER_HEADER, sFolder) Then
                Call Z_PRINT("�A�b�e�l�[�^�e�[�u���t�@�C�����擾�o���܂���ł����B")
                Return (False)
            End If

            If IO.File.Exists(sFolder) = True Then  ' �t�@�C�����L��B
                Using sr As New System.IO.StreamReader(sFolder, System.Text.Encoding.GetEncoding("Shift-JIS"))
                    Do While Not sr.EndOfStream
                        sData = sr.ReadLine
                        If bHeader Then                                     ' �P�s�ڂ̓^�C�g���s
                            bHeader = False
                            Continue Do
                        End If
                        mData = sData.Split(",")                            ' �������','�ŕ������Ď�o��
                        TableNo = Integer.Parse(mData(0))

                        If TableNo > MaxNo Then
                            MaxNo = TableNo
                        End If

                        If TableNo = No Then                                            ' �ԍ���v
                            stData.No = TableNo
                            stData.Power = mData(1).Trim                                ' �p���[�ݒ�
                            If Not Double.TryParse(stData.Power, dData) Then
                                Call Z_PRINT("[" & No.ToString & "]�Ԗ�[" & stData.Power & "]�p���[�ݒ�l�����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If
                            stData.PowerUnit = mData(2).Trim                            ' �p���[�P��
                            stData.Limit = mData(3).Trim                                ' �͈�
                            If Not Double.TryParse(stData.Limit, dData) Then
                                Call Z_PRINT("[" & No.ToString & "]�Ԗ�[" & stData.Limit & "]�p���[�ݒ�͈͂����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If
                            stData.LimitUnit = mData(4).Trim                            ' �͈͒P��

                            If Not Double.TryParse(mData(5), stData.Rate) Then          ' ������
                                Call Z_PRINT("[" & No.ToString & "]�Ԗ�[" & mData(5) & "]�����������l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If
                            stData.RateUnit = mData(6).Trim                             ' �������P��

                            If Not Integer.TryParse(mData(7), stData.Rotation) Then     ' ��]��
                                Call Z_PRINT("[" & No.ToString & "]�Ԗ�[" & mData(7) & "]��]�ʂ����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If

                            If Not Integer.TryParse(mData(8), stData.FixAtt) Then     ' �Œ�A�b�e�l�[�^
                                Call Z_PRINT("[" & No.ToString & "]�Ԗ�[" & mData(8) & "]�Œ�A�b�e�l�[�^�����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If

                            stData.Comment = ""
                            For i As Integer = 9 To mData.Length - 1
                                stData.Comment = stData.Comment & "," & mData(i)                   ' �R�����g
                            Next
                            Return (True)
                        End If
                    Loop
                End Using
            End If

            Return (False)

        Catch ex As Exception
            MsgBox("LaserCalibrationAttenuatorTableGet() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (False)
        End Try
    End Function

    ''' <summary>
    ''' �A�b�e�l�[�^�e�[�u�����̍ő�ԍ����擾
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LaserCalibrationMaxNumberGet() As Integer
        Try

            Dim stAttenuatorTable As stATTENUATOR_TABLE
            Dim MaxNo As Integer

            stAttenuatorTable.No = -1           ' �ԍ�
            stAttenuatorTable.Power = ""        ' �p���[�ݒ�
            stAttenuatorTable.PowerUnit = ""    ' �p���[�P��
            stAttenuatorTable.Limit = ""        ' �͈�
            stAttenuatorTable.LimitUnit = ""    ' �͈͒P��
            stAttenuatorTable.Rate = -1.0       ' ������
            stAttenuatorTable.RateUnit = ""     ' �������P��
            stAttenuatorTable.Rotation = -1     ' ��]��
            stAttenuatorTable.FixAtt = -1       ' �Œ�A�b�e�l�[�^
            stAttenuatorTable.Comment = ""      ' �R�����g

            Call LaserCalibrationAttenuatorTableGet(MAX_ATTENUATOR + 1, stAttenuatorTable, MaxNo)

            If MaxNo > MAX_ATTENUATOR Then
                MaxNo = MAX_ATTENUATOR
            End If

            Return (MaxNo)

        Catch ex As Exception
            MsgBox("LaserCalibrationMaxNumberGetByAttenuatorTable() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (0)
        End Try
    End Function

    ''' <summary>
    ''' �t���p���[�l�̎擾
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LaserCalibrationFullPowerGet(ByRef FullPowerTarget As Double, ByRef FullPowerLimit As Double) As Boolean
        Try
            Dim stAttenuatorTable As stATTENUATOR_TABLE
            Dim MaxNo As Integer

            stAttenuatorTable.No = -1           ' �ԍ�
            stAttenuatorTable.Power = ""        ' �p���[�ݒ�
            stAttenuatorTable.PowerUnit = ""    ' �p���[�P��
            stAttenuatorTable.Limit = ""        ' �͈�
            stAttenuatorTable.LimitUnit = ""    ' �͈͒P��
            stAttenuatorTable.Rate = -1.0       ' ������
            stAttenuatorTable.RateUnit = ""     ' �������P��
            stAttenuatorTable.Rotation = -1     ' ��]��
            stAttenuatorTable.FixAtt = -1       ' �Œ�A�b�e�l�[�^
            stAttenuatorTable.Comment = ""      ' �R�����g

            If LaserCalibrationAttenuatorTableGet(0, stAttenuatorTable, MaxNo) Then

                If Not Double.TryParse(stAttenuatorTable.Power.Trim, FullPowerTarget) Then
                    Call Z_PRINT("[0]�Ԗڃt���p���[[" & stAttenuatorTable.Power & "]�p���[�ݒ�l�����l�ɕϊ��ł��܂���B")
                    Return (False)
                Else
                    If stAttenuatorTable.PowerUnit.IndexOf("mW") >= 0 OrElse stAttenuatorTable.PowerUnit.IndexOf("���v") >= 0 Then
                        FullPowerTarget = FullPowerTarget / 1000.0
                    End If
                End If

                If Not Double.TryParse(stAttenuatorTable.Limit.Trim, FullPowerLimit) Then
                    Call Z_PRINT("[[0]�Ԗ�[" & stAttenuatorTable.Limit & "]�p���[�ݒ�͈͂����l�ɕϊ��ł��܂���B")
                    Return (False)
                Else
                    If stAttenuatorTable.LimitUnit.IndexOf("mW") >= 0 OrElse stAttenuatorTable.LimitUnit.IndexOf("���v") >= 0 Then
                        FullPowerLimit = FullPowerLimit / 1000.0
                    End If
                End If

                Return (True)
            Else
                Return (False)
            End If

        Catch ex As Exception
            MsgBox("LaserCalibrationFullPowerGet() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (False)
        End Try
    End Function

    ''' <summary>
    ''' �A�b�e�l�[�^�e�[�u������ݒ�f�[�^�̎擾
    ''' </summary>
    ''' <param name="No">�ԍ�</param>
    ''' <param name="dblRotPar">������</param>
    ''' <param name="iFixAtt">�Œ�A�b�e�l�[�^</param>
    ''' <param name="dblRotAtt">��]��</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LaserCalibrationAttenuatorDataGet(ByVal No As Integer, ByRef dblRotPar As Double, ByRef iFixAtt As Double, ByRef dblRotAtt As Double) As Boolean
        Try
            Dim stAttenuatorTable As stATTENUATOR_TABLE
            Dim MaxNo As Integer

            If No < 0 Or MAX_ATTENUATOR < No Then
                Return (False)
            End If

            stAttenuatorTable.No = -1           ' �ԍ�
            stAttenuatorTable.Power = ""        ' �p���[�ݒ�
            stAttenuatorTable.PowerUnit = ""    ' �p���[�P��
            stAttenuatorTable.Limit = ""        ' �͈�
            stAttenuatorTable.LimitUnit = ""    ' �͈͒P��
            stAttenuatorTable.Rate = -1.0       ' ������
            stAttenuatorTable.RateUnit = ""     ' �������P��
            stAttenuatorTable.Rotation = -1     ' ��]��
            stAttenuatorTable.FixAtt = -1       ' �Œ�A�b�e�l�[�^
            stAttenuatorTable.Comment = ""      ' �R�����g

            If LaserCalibrationAttenuatorTableGet(No, stAttenuatorTable, MaxNo) Then
                dblRotPar = stAttenuatorTable.Rate
                iFixAtt = stAttenuatorTable.FixAtt
                dblRotAtt = stAttenuatorTable.Rotation
                Return (True)
            Else
                Return (False)
            End If

        Catch ex As Exception
            MsgBox("LaserCalibrationAttenuatorDataGet() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (False)
        End Try
    End Function

    ''' <summary>
    '''  �A�b�e�l�[�^�e�[�u������S�f�[�^�擾
    ''' </summary>
    ''' <param name="stData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LaserCalibrationAllDataGet(ByRef MaxNo As Integer, ByRef stData() As stATTENUATOR_TABLE) As Boolean
        Try
            Dim bRtn As Boolean = False
            Dim iTemp As Integer

            MaxNo = LaserCalibrationMaxNumberGet()

            If MaxNo <= 0 Then
                Return (False)
            End If

            For No As Integer = 0 To MaxNo
                If LaserCalibrationAttenuatorTableGet(No, stData(No), iTemp) Then
                    bRtn = True
                Else
                    Z_PRINT("�A�b�e�l�[�^�e�[�u������̏��擾���G���[�ɂȂ�܂����BNO=[" & No.ToString & "]")
                    Return (False)
                End If
            Next

            Return (bRtn)

        Catch ex As Exception
            MsgBox("LaserCalibrationAllDataGet() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (False)
        End Try
    End Function

    Public Function LaserCalibrationAllDataWrite(ByRef stData() As stATTENUATOR_TABLE) As Boolean
        Try
            Dim sFolder As String = vbNullString
            Dim sHeaderData As String = vbNullString
            Dim sFileName As String

            If Not GetNewestFile(cLASERPOWER_PATH, cLASERPOWER_HEADER, sFolder) Then
                Call Z_PRINT("�A�b�e�l�[�^�e�[�u���t�@�C�����擾�o���܂���ł����B")
                Return (False)
            End If

            Dim MaxNo As Integer = LaserCalibrationMaxNumberGet()

            If IO.File.Exists(sFolder) = True Then  ' �t�@�C�����L��B
                Using sr As New System.IO.StreamReader(sFolder, System.Text.Encoding.GetEncoding("Shift-JIS"))
                    Do While Not sr.EndOfStream
                        sHeaderData = sr.ReadLine
                        Exit Do
                    Loop
                End Using
            End If

            sFileName = cLASERPOWER_PATH & cLASERPOWER_HEADER & DateTime.Now().ToString("yyyyMMdd") & ".CSV"

            Using WSR As New System.IO.StreamWriter(sFileName, False, System.Text.Encoding.GetEncoding("Shift-JIS"))  ' ��Q���� �㏑���́AFalse
                WSR.WriteLine(sHeaderData)                          ' �w�b�_�o��

                'Public Structure stATTENUATOR_TABLE         ' �A�b�e�l�[�^�e�[�u��
                '            Dim No As Integer                       ' �ԍ�
                '            Dim Power As String                     ' �p���[�ݒ�
                '            Dim PowerUnit As String                 ' �p���[�P��
                '            Dim Limit As String                     ' �͈�
                '            Dim LimitUnit As String                 ' �͈͒P��
                '            Dim Rate As Double                      ' ������
                '            Dim RateUnit As String                  ' �������P��
                '            Dim Rotation As Integer                 ' ��]��
                '            Dim FixAtt As Integer                   ' �Œ�A�b�e�l�[�^
                '            Dim Comment As String                   ' �R�����g
                'End Structure

                For No As Integer = 0 To MaxNo
                    WSR.WriteLine(stData(No).No.ToString("0") & "," & stData(No).Power & "," & stData(No).PowerUnit & "," & stData(No).Limit & "," & stData(No).LimitUnit & "," & stData(No).Rate.ToString("0.00") & "," & stData(No).RateUnit & "," & stData(No).Rotation.ToString("0") & "," & stData(No).FixAtt.ToString("0") & stData(No).Comment)
                Next
            End Using

            Return (True)

        Catch ex As Exception
            MsgBox("LaserCalibrationAllDataWrite() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (False)
        End Try
    End Function

#End Region             '���[�U�p���[���j�^�����O
    'V2.1.0.0�A��

    'V2.1.0.0�B��
#Region "���x�Z���T�[���e�[�u��"
    ''' <param name="stData">�擾�f�[�^�\����</param>
    ''' <param name="MaxNo">���e�[�u�����ő�ԍ�</param>
    ''' <returns>True:�����ԍ��Ɉ�v�������̂�</returns>
    ''' <remarks></remarks>
    Private Function TemperatureTableGet(ByVal No As Integer, ByRef stData As stTEMPERATURE_TABLE, ByRef MaxNo As Integer) As Boolean
        Try
            Dim sFolder As String = vbNullString
            Dim sData As String
            Dim mData() As String
            Dim TableNo As Integer
            Dim bHeader As Boolean = True
            Dim dData As Double

            MaxNo = 0

            If Not GetNewestFile(cTEMPERATURE_PATH, cTEMPERATURE_HEADER, sFolder) Then
                Call Z_PRINT("���x�Z���T�[���e�[�u���t�@�C�����擾�o���܂���ł����B")
                Return (False)
            End If

            If IO.File.Exists(sFolder) = True Then  ' �t�@�C�����L��B
                Using sr As New System.IO.StreamReader(sFolder, System.Text.Encoding.GetEncoding("Shift-JIS"))
                    Do While Not sr.EndOfStream
                        sData = sr.ReadLine
                        If bHeader Then                                     ' �P�s�ڂ̓^�C�g���s
                            bHeader = False
                            Continue Do
                        End If
                        mData = sData.Split(",")                            ' �������','�ŕ������Ď�o��
                        TableNo = Integer.Parse(mData(0))

                        If TableNo > MaxNo Then
                            MaxNo = TableNo
                        End If

                        If TableNo = No Then                                ' �ԍ���v
                            stData.No = TableNo                             ' �ԍ�
                            stData.Title = mData(1)                         ' ���f�L��
                            If mData(2) = vbNullString Or mData(3) = vbNullString Or mData(4) = vbNullString Then
                                Call Z_PRINT("[" & No.ToString & "]�ԖڂŃf�[�^�����݂��Ȃ����ڂ�����܂��B")
                                Return (False)
                            End If

                            If Double.TryParse(mData(2), dData) Then
                                stData.dTemperatura0 = Double.Parse(mData(2))   ' �O��
                            Else
                                Call Z_PRINT("[" & No.ToString & "]�ԖڂO���f�[�^[" & mData(2) & "]�����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If

                            If Double.TryParse(mData(3), dData) Then
                                stData.dDaihyouAlpha = Double.Parse(mData(3))   ' ��\���l
                            Else
                                Call Z_PRINT("[" & No.ToString & "]�Ԗڃ��l�f�[�^[" & mData(3) & "]�����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If


                            If Double.TryParse(mData(4), dData) Then
                                stData.dDaihyouBeta = Double.Parse(mData(4))    ' ��\���l
                            Else
                                Call Z_PRINT("[" & No.ToString & "]�Ԗڃ��l�f�[�^[" & mData(4) & "]�����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If

                            For i As Integer = 5 To mData.Length - 1
                                stData.Comment = "," & mData(i)                   ' �R�����g
                            Next
                            Return (True)
                        End If
                    Loop
                End Using
            End If

            Return (False)

        Catch ex As Exception
            MsgBox("TemperatureTableGet() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (False)
        End Try
    End Function

    ''' <summary>
    ''' ���x�Z���T�[���e�[�u�����̍ő�ԍ����擾
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function TemperatureTableMaxNumberGet() As Integer
        Try

            Dim stTemperatureTable As stTEMPERATURE_TABLE
            Dim MaxNo As Integer

            stTemperatureTable.No = -1                       ' �ԍ�
            stTemperatureTable.Title = ""                    ' ���f�L��
            stTemperatureTable.dTemperatura0 = -1.0             ' �O��
            stTemperatureTable.dDaihyouAlpha = -1.0             ' ��\���l
            stTemperatureTable.dDaihyouBeta = -1.0              ' ��\���l
            stTemperatureTable.Comment = ""                    ' �R�����g


            Call TemperatureTableGet(MAX_TEMPERATURE + 1, stTemperatureTable, MaxNo)

            If MaxNo > MAX_TEMPERATURE Then
                MaxNo = MAX_TEMPERATURE
            End If

            Return (MaxNo)

        Catch ex As Exception
            MsgBox("TemperatureTableMaxNumberGet() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (0)
        End Try
    End Function

    ''' <summary>
    ''' ���x�Z���T�[���e�[�u��������擾
    ''' </summary>
    ''' <param name="No">�����Ώ۔ԍ�</param>
    ''' <param name="dTemperatura0">�O��</param>
    ''' <param name="dDaihyouAlpha">��\���l</param>
    ''' <param name="dDaihyouBeta">��\���l</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function TemperatureTableDataGet(ByVal No As Integer, ByRef dTemperatura0 As Double, ByRef dDaihyouAlpha As Double, ByRef dDaihyouBeta As Double, Optional bLimitCheck As Boolean = False) As Boolean
        Try
            Dim stTemperatureTable As stTEMPERATURE_TABLE
            Dim MaxNo As Integer
            Dim Min As Double, Max As Double

            If No < 1 Or MAX_TEMPERATURE < No Then
                Return (False)
            End If

            stTemperatureTable.No = -1                      ' �ԍ�
            stTemperatureTable.Title = ""                   ' ���f�L��
            stTemperatureTable.dTemperatura0 = -1.0         ' �O��
            stTemperatureTable.dDaihyouAlpha = -1.0         ' ��\���l
            stTemperatureTable.dDaihyouBeta = -1.0          ' ��\���l
            stTemperatureTable.Comment = ""                 ' �R�����g


            If TemperatureTableGet(No, stTemperatureTable, MaxNo) Then
                dTemperatura0 = stTemperatureTable.dTemperatura0    ' �O��
                dDaihyouAlpha = stTemperatureTable.dDaihyouAlpha    ' ��\���l
                dDaihyouBeta = stTemperatureTable.dDaihyouBeta      ' ��\���l
                ' �͈̓`�F�b�N
                If bLimitCheck Then
                    ' �O��
                    Min = Double.Parse(GetPrivateProfileString_S("USER_VALUE", "011_MIN", cEDITDEF_FNAME, "0.0000001"))
                    Max = Double.Parse(GetPrivateProfileString_S("USER_VALUE", "011_MAX", cEDITDEF_FNAME, "100000000.0000000"))
                    If dTemperatura0 < Min Or Max < dTemperatura0 Then
                        Z_PRINT("���x�Z���T�[���e�[�u��No.=[" & No.ToString("0") & "]�O���㉺���l�G���[=[" & dTemperatura0.ToString("0.00000000") & "]")
                        Return (False)
                    End If
                    ' ���l(ppm/��)
                    Min = Double.Parse(GetPrivateProfileString_S("USER_VALUE", "012_MIN", cEDITDEF_FNAME, "-9999.0000000"))
                    Max = Double.Parse(GetPrivateProfileString_S("USER_VALUE", "012_MAX", cEDITDEF_FNAME, "9999.0000000"))
                    If dDaihyouAlpha < Min Or Max < dDaihyouAlpha Then
                        Z_PRINT("���x�Z���T�[���e�[�u��No.=[" & No.ToString("0") & "]���l�㉺���l�G���[=[" & dDaihyouAlpha.ToString("0.00000000") & "]")
                        Return (False)
                    End If
                    ' ���l(ppm/��)
                    Min = Double.Parse(GetPrivateProfileString_S("USER_VALUE", "013_MIN", cEDITDEF_FNAME, "-9999.0000000"))
                    Max = Double.Parse(GetPrivateProfileString_S("USER_VALUE", "013_MAX", cEDITDEF_FNAME, "9999.0000000"))
                    If dDaihyouBeta < Min Or Max < dDaihyouBeta Then
                        Z_PRINT("���x�Z���T�[���e�[�u��No.=[" & No.ToString("0") & "]���l�㉺���l�G���[=[" & dDaihyouBeta.ToString("0.00000000") & "]")
                        Return (False)
                    End If
                End If

                Return (True)
            Else
                Return (False)
            End If

        Catch ex As Exception
            MsgBox("TemperatureTableDataGet() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (False)
        End Try
    End Function
#End Region             ' ���x�Z���T�[���e�[�u��
    'V2.1.0.0�B��
#End Region         'V2.1.0.0�@�A�B

#Region "�v���[�u�}�X�^�[�f�[�^��Ǎ��ݎw��No�̃v���[�u�f�[�^��ݒ肷��"    'V2.2.0.0�N
    ''' <summary>
    ''' �v���[�u�f�[�^��ݒ肷�� 
    ''' </summary>
    ''' <returns></returns>
    Public Function ConvProbeData(ByVal ProbNo As Integer) As Integer
        Dim Ret As Boolean
        Dim stlocalProbeData As stPROBEDATA_TABLE
        Dim MaxNo As Integer

        Try
            ConvProbeData = cFRS_NORMAL

            If ProbNo = 0 Then
                Exit Function
            End If

            stlocalProbeData.No = 0
            stlocalProbeData.ProbeOn = 0.0
            stlocalProbeData.ProbeOff = 0.0                 'V2.2.0.0�S
            stlocalProbeData.dTableOffsetX = 0.0
            stlocalProbeData.dTableOffsetY = 0.0
            stlocalProbeData.dBPOffsetX = 0.0
            stlocalProbeData.dBPOffsetY = 0.0
            stlocalProbeData.Comment = ""

            '�ƕ␳�֌W�̓f�[�^���Ȃ��ꍇ������̂ŏ������͂��Ȃ�      'V2.2.1.6�A
            'V2.2.1.6�A��
            stlocalProbeData.iPP30 = stThta.iPP30                   ' �ʒu�␳���[�h�F
            stlocalProbeData.iPP31 = stThta.iPP31                   ' �ʒu�␳���@�F
            stlocalProbeData.fpp34_x = stThta.fpp34_x               ' �␳�|�W�V�����I�t�Z�b�gX�F
            stlocalProbeData.fpp34_y = stThta.fpp34_y               ' �␳�|�W�V�����I�t�Z�b�gY�F
            stlocalProbeData.fTheta = stThta.fTheta                 ' �p�x
            stlocalProbeData.iPP38 = stThta.iPP38                   ' �O���[�v�ԍ��F
            stlocalProbeData.iPP37_1 = stThta.iPP37_1               ' �p�^�[���ԍ�1�F
            stlocalProbeData.fpp32_x = stThta.fpp32_x               ' �p�^�[�����W1X�F
            stlocalProbeData.fpp32_y = stThta.fpp32_y               ' �p�^�[�����W1Y�F
            stlocalProbeData.iPP37_2 = stThta.iPP37_2               ' �p�^�[���ԍ�2�F
            stlocalProbeData.fpp33_x = stThta.fpp33_x               ' �p�^�[�����W2X�F
            stlocalProbeData.fpp33_y = stThta.fpp33_y               ' �p�^�[�����W2Y�F
            'V2.2.1.6�A��

            ' PROBEDATA�t�@�C����Ǎ���ŁA���̓��e�ɍX�V���� 
            Ret = ReadProbeCsv(ProbNo, stlocalProbeData, MaxNo)
            If Ret = True Then

                'V2.2.1.6�A��
                Z_PRINT("�v���[�u�e�[�u������NO=[" & ProbNo.ToString() & "]�̏����擾���܂����B")
                Z_PRINT("�@�v���[�uON�ʒu=[" & stlocalProbeData.ProbeOn.ToString("#0.000#") & "] ")
                Z_PRINT("�@�v���[�uOFF�ʒu=[" & stlocalProbeData.ProbeOff.ToString("#0.000#") & "] ")          'V2.2.0.0�S
                Z_PRINT("�@�e�[�u���ʒu�I�t�Z�b�gX=[" & stlocalProbeData.dTableOffsetX.ToString("#0.000000#") & "] ")
                Z_PRINT("�@�e�[�u���ʒu�I�t�Z�b�gY=[" & stlocalProbeData.dTableOffsetY.ToString("#0.000000#") & "] ")
                Z_PRINT("�@BP�I�t�Z�b�gX=[" & stlocalProbeData.dBPOffsetX.ToString("#0.000#") & "] ")
                Z_PRINT("�@BP�I�t�Z�b�gX=[" & stlocalProbeData.dBPOffsetY.ToString("#0.000#") & "] ")

                'V2.2.1.6�A��
                Z_PRINT("  �ʒu�␳���[�h=[" & stlocalProbeData.iPP30.ToString() & "]�̏����擾���܂����B")
                Z_PRINT("  �ʒu�␳���@=[" & stlocalProbeData.iPP31.ToString() & "]�̏����擾���܂����B")
                Z_PRINT("�@�␳�|�W�V�����I�t�Z�b�gX=[" & stlocalProbeData.fpp34_x.ToString("#0.000#") & "] ")
                Z_PRINT("�@�␳�|�W�V�����I�t�Z�b�gY=[" & stlocalProbeData.fpp34_y.ToString("#0.000#") & "] ")
                Z_PRINT("�@�p�x=[" & stlocalProbeData.fTheta.ToString("#0.000#") & "] ")
                Z_PRINT("�@�O���[�v�ԍ�=[" & stlocalProbeData.iPP38.ToString() & "] ")
                Z_PRINT("�@�p�^�[���ԍ�1=[" & stlocalProbeData.iPP37_1.ToString() & "] ")
                Z_PRINT("�@�p�^�[�����W1X=[" & stlocalProbeData.fpp32_x.ToString("#0.000#") & "] ")
                Z_PRINT("�@�p�^�[�����W1Y=[" & stlocalProbeData.fpp32_y.ToString("#0.000#") & "] ")
                Z_PRINT("�@�p�^�[���ԍ�2=[" & stlocalProbeData.iPP37_2.ToString() & "] ")
                Z_PRINT("�@�p�^�[�����W2X=[" & stlocalProbeData.fpp33_x.ToString("#0.000#") & "] ")
                Z_PRINT("�@�p�^�[�����W2Y=[" & stlocalProbeData.fpp33_y.ToString("#0.000#") & "] ")
                'V2.2.1.6�A��

                Z_PRINT("�@�R�����g=[" & stlocalProbeData.Comment & "] ")
            Else
                Z_PRINT("�v���[�u�e�[�u������NO=[" & ProbNo.ToString & "]�̏����擾���ɃG���[���������܂����B")
                ConvProbeData = cFRS_FIOERR_INP

            End If
            stPLT.z_xoff = stlocalProbeData.dTableOffsetX       ' �e�[�u���ʒu�I�t�Z�b�g�@�g�����|�W�V�����I�t�Z�b�gX(mm)
            stPLT.z_yoff = stlocalProbeData.dTableOffsetY       ' �e�[�u���ʒu�I�t�Z�b�g�@�g�����|�W�V�����I�t�Z�b�gY(mm)
            stPLT.BPOX = stlocalProbeData.dBPOffsetX            ' �r�[���ʒu�I�t�Z�b�g�@BP Offset X(mm)
            stPLT.BPOY = stlocalProbeData.dBPOffsetY            ' �r�[���ʒu�I�t�Z�b�g�@BP Offset Y(mm)
            stPLT.Z_ZON = stlocalProbeData.ProbeOn              ' Z PROBE ON �ʒu(mm)
            stPLT.Z_ZOFF = stlocalProbeData.ProbeOff            ' Z PROBE OFF �ʒu(mm)            ' V2.2.0.0�S
            'V2.2.1.6�A��
            stThta.iPP30 = stlocalProbeData.iPP30               ' �ʒu�␳���[�h�F
            stThta.iPP31 = stlocalProbeData.iPP31               ' �ʒu�␳���@�F
            stThta.fpp34_x = stlocalProbeData.fpp34_x           ' �␳�|�W�V�����I�t�Z�b�gX�F
            stThta.fpp34_y = stlocalProbeData.fpp34_y           ' �␳�|�W�V�����I�t�Z�b�gY�F
            stThta.fTheta = stlocalProbeData.fTheta             ' �p�x
            stThta.iPP38 = stlocalProbeData.iPP38               ' �O���[�v�ԍ��F
            stThta.iPP37_1 = stlocalProbeData.iPP37_1           ' �p�^�[���ԍ�1�F
            stThta.fpp32_x = stlocalProbeData.fpp32_x           ' �p�^�[�����W1X�F
            stThta.fpp32_y = stlocalProbeData.fpp32_y           ' �p�^�[�����W1Y�F
            stThta.iPP37_2 = stlocalProbeData.iPP37_2           ' �p�^�[���ԍ�2�F
            stThta.fpp33_x = stlocalProbeData.fpp33_x           ' �p�^�[�����W2X�F
            stThta.fpp33_y = stlocalProbeData.fpp33_y           ' �p�^�[�����W2Y�F
            'V2.2.1.6�A��

            If stPLT.Z_ZON < stPLT.Z_ZOFF Then
                Z_PRINT("�v���[�uON�ʒu���v���[�u�ҋ@�ʒu�����Ⴍ�ݒ肳��Ă��܂��B")
                Z_PRINT(" [ON�ʒu=" & stPLT.Z_ZON & "],[�ҋ@�ʒu=" & stPLT.Z_ZOFF.ToString & "]")
            End If

        Catch ex As Exception

        End Try

    End Function

#End Region


#Region "���݂̃v���[�u�f�[�^�t�@�C����Ǎ���Ŏw��No�̃v���[�u�f�[�^���X�V���āA�v���[�u�f�[�^�t�@�C������������"    'V2.2.0.0�N
    ''' <summary>
    ''' ���݂̃v���[�u�f�[�^�t�@�C����Ǎ���Ŏw��No�̃v���[�u�f�[�^���X�V���āA�v���[�u�f�[�^�t�@�C������������
    ''' </summary>
    ''' <returns></returns>
    Public Function UpdateProbeData(ByVal ProbNo As Integer) As Integer
        Dim stlocalProbeData(PROBE_DATA_MAX) As stPROBEDATA_TABLE       'V2.2.0.038�@'V2.2.1.0�@
        ''V2.2.1.0�@�@Dim stlocalProbeData(11) As stPROBEDATA_TABLE�@
        Dim Maxno As Integer
        Dim Header As String = ""

        Try

            ' �v���[�u�f�[�^��S�ēǍ���
            ReadAllProbeCsv(stlocalProbeData, Maxno, Header)

            ' �w��̃v���[�uNo�̃f�[�^���X�V���� 
            stlocalProbeData(ProbNo).dTableOffsetX = stPLT.z_xoff      ' �e�[�u���ʒu�I�t�Z�b�g�@�g�����|�W�V�����I�t�Z�b�gX(mm)
            stlocalProbeData(ProbNo).dTableOffsetY = stPLT.z_yoff      ' �e�[�u���ʒu�I�t�Z�b�g�@�g�����|�W�V�����I�t�Z�b�gY(mm)
            stlocalProbeData(ProbNo).dBPOffsetX = stPLT.BPOX           ' �r�[���ʒu�I�t�Z�b�g�@BP Offset X(mm)
            stlocalProbeData(ProbNo).dBPOffsetY = stPLT.BPOY           ' �r�[���ʒu�I�t�Z�b�g�@BP Offset Y(mm)
            stlocalProbeData(ProbNo).ProbeOn = stPLT.Z_ZON             ' Z PROBE ON �ʒu(mm)
            stlocalProbeData(ProbNo).ProbeOff = stPLT.Z_ZOFF           ' Z PROBE OFF �ʒu(mm)                'V2.2.0.0�S 

            stlocalProbeData(ProbNo).iPP30 = stThta.iPP30              ' �ʒu�␳���[�h�F
            stlocalProbeData(ProbNo).iPP31 = stThta.iPP31              ' �ʒu�␳���@�F
            stlocalProbeData(ProbNo).fpp34_x = stThta.fpp34_x          ' �␳�|�W�V�����I�t�Z�b�gX�F
            stlocalProbeData(ProbNo).fpp34_y = stThta.fpp34_y          ' �␳�|�W�V�����I�t�Z�b�gY�F
            stlocalProbeData(ProbNo).fTheta = stThta.fTheta            ' �p�x�F
            stlocalProbeData(ProbNo).iPP38 = stThta.iPP38              ' �O���[�v�ԍ�
            stlocalProbeData(ProbNo).iPP37_1 = stThta.iPP37_1          ' �p�^�[���ԍ�
            stlocalProbeData(ProbNo).fpp32_x = stThta.fpp32_x          ' �p�^�[�����W1X�F
            stlocalProbeData(ProbNo).fpp32_y = stThta.fpp32_y          ' �p�^�[�����W1Y�F
            stlocalProbeData(ProbNo).iPP37_2 = stThta.iPP37_2          ' �p�^�[���ԍ�
            stlocalProbeData(ProbNo).fpp33_x = stThta.fpp33_x          ' �p�^�[�����W1X�F
            stlocalProbeData(ProbNo).fpp33_y = stThta.fpp33_y          ' �p�^�[�����W1Y�F

            '�v���[�u�f�[�^����������
            WriteAllProbeCsv(stlocalProbeData, Maxno, Header)

        Catch ex As Exception

        End Try

    End Function

#End Region


#Region "�w�肵��No�̃v���[�u�f�[�^�̓Ǎ���"    'V2.2.0.0�N
    ''' <summary>
    ''' �w�肵��No�̃v���[�u�f�[�^�̓Ǎ���
    ''' </summary>
    ''' <returns></returns>
    Public Function ReadProbeCsv(ByVal No As Integer, ByRef stData As stPROBEDATA_TABLE, ByRef MaxNo As Integer) As Boolean
        Dim sFolder As String = ""
        Dim sData As String = ""
        Dim bHeader As Boolean = True
        Dim mData() As String
        Dim TableNo As Integer
        Dim dData As Double
        Dim Ret As Boolean = False

        Try

            sFolder = cPROBEDATA_PATH & cPROBEDATA_FILE

            If IO.File.Exists(sFolder) = True Then  ' �t�@�C�����L��B
                Using sr As New System.IO.StreamReader(sFolder, System.Text.Encoding.GetEncoding("Shift-JIS"))
                    Do While Not sr.EndOfStream
                        sData = sr.ReadLine
                        If bHeader Then                                     ' �P�s�ڂ̓^�C�g���s
                            bHeader = False
                            Continue Do
                        End If
                        mData = sData.Split(",")                            ' �������','�ŕ������Ď�o��
                        TableNo = Integer.Parse(mData(0))

                        If TableNo > MaxNo Then
                            MaxNo = TableNo
                        End If

                        If TableNo = No Then                                            ' �ԍ���v
                            stData.No = TableNo

                            If Double.TryParse(mData(1), dData) Then
                                stData.ProbeOn = Double.Parse(mData(1))                                ' �v���[�u�ڐG�ʒu
                            Else
                                Call Z_PRINT("[" & No.ToString & "]�Ԗڂ̃f�[�^[" & mData(1) & "]�����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If

                            'V2.2.0.021��
                            If Double.TryParse(mData(2), dData) Then
                                stData.ProbeOff = Double.Parse(mData(2))                                ' �v���[�u�ҋ@�ʒu
                            Else
                                Call Z_PRINT("[" & No.ToString & "]�Ԗڂ̃f�[�^[" & mData(2) & "]�����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If
                            'V2.2.0.021��

                            If Double.TryParse(mData(3), dData) Then
                                stData.dTableOffsetX = Double.Parse(mData(3))                          ' �e�[�u���I�t�Z�b�g�w
                            Else
                                Call Z_PRINT("[" & No.ToString & "]�Ԗڂ̃f�[�^[" & mData(3) & "]�����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If

                            If Double.TryParse(mData(4), dData) Then
                                stData.dTableOffsetY = Double.Parse(mData(4))                          ' �e�[�u���I�t�Z�b�g�x
                            Else
                                Call Z_PRINT("[" & No.ToString & "]�Ԗڂ̃f�[�^[" & mData(4) & "]�����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If

                            If Double.TryParse(mData(5), dData) Then
                                stData.dBPOffsetX = Double.Parse(mData(5))                          ' BP�I�t�Z�b�g�w
                            Else
                                Call Z_PRINT("[" & No.ToString & "]�Ԗڂ̃f�[�^[" & mData(5) & "]�����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If

                            If Double.TryParse(mData(6), dData) Then
                                stData.dBPOffsetY = Double.Parse(mData(6))                          ' BP�I�t�Z�b�g�x
                            Else
                                Call Z_PRINT("[" & No.ToString & "]�Ԗڂ̃f�[�^[" & mData(6) & "]�����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If

                            'V2.2.1.6�A ��
                            If mData.Length > 8 Then
                                ' �V�����f�[�^�̏ꍇ�A�ƕ␳�֌W�̃p�����[�^���܂܂��̂�20�����ɂȂ�
                                If Short.TryParse(mData(7), dData) Then
                                    stData.iPP30 = Short.Parse(mData(7))                          ' �ʒu�␳���[�h
                                Else
                                    Call Z_PRINT("[" & No.ToString & "]�Ԗڂ̃f�[�^[" & mData(7) & "]�����l�ɕϊ��ł��܂���B")
                                    Return (False)
                                End If
                                If Short.TryParse(mData(8), dData) Then
                                    stData.iPP31 = Short.Parse(mData(8))                          ' �ʒu�␳���@
                                Else
                                    Call Z_PRINT("[" & No.ToString & "]�Ԗڂ̃f�[�^[" & mData(8) & "]�����l�ɕϊ��ł��܂���B")
                                    Return (False)
                                End If
                                If Double.TryParse(mData(9), dData) Then
                                    stData.fpp34_x = Double.Parse(mData(9))                          ' �␳�|�W�V�����I�t�Z�b�gX
                                Else
                                    Call Z_PRINT("[" & No.ToString & "]�Ԗڂ̃f�[�^[" & mData(9) & "]�����l�ɕϊ��ł��܂���B")
                                    Return (False)
                                End If
                                If Double.TryParse(mData(10), dData) Then
                                    stData.fpp34_y = Double.Parse(mData(10))                          ' �␳�|�W�V�����I�t�Z�b�gY
                                Else
                                    Call Z_PRINT("[" & No.ToString & "]�Ԗڂ̃f�[�^[" & mData(10) & "]�����l�ɕϊ��ł��܂���B")
                                    Return (False)
                                End If
                                If Double.TryParse(mData(11), dData) Then
                                    stData.fTheta = Double.Parse(mData(11))                          ' �p�x
                                Else
                                    Call Z_PRINT("[" & No.ToString & "]�Ԗڂ̃f�[�^[" & mData(11) & "]�����l�ɕϊ��ł��܂���B")
                                    Return (False)
                                End If
                                If Short.TryParse(mData(12), dData) Then
                                    stData.iPP38 = Short.Parse(mData(12))                          ' �O���[�v�ԍ�
                                Else
                                    Call Z_PRINT("[" & No.ToString & "]�Ԗڂ̃f�[�^[" & mData(12) & "]�����l�ɕϊ��ł��܂���B")
                                    Return (False)
                                End If
                                If Short.TryParse(mData(13), dData) Then
                                    stData.iPP37_1 = Short.Parse(mData(13))                          ' �p�^�[���ԍ�1
                                Else
                                    Call Z_PRINT("[" & No.ToString & "]�Ԗڂ̃f�[�^[" & mData(13) & "]�����l�ɕϊ��ł��܂���B")
                                    Return (False)
                                End If
                                If Double.TryParse(mData(14), dData) Then
                                    stData.fpp32_x = Double.Parse(mData(14))                          ' �p�^�[�����W1X
                                Else
                                    Call Z_PRINT("[" & No.ToString & "]�Ԗڂ̃f�[�^[" & mData(14) & "]�����l�ɕϊ��ł��܂���B")
                                    Return (False)
                                End If
                                If Double.TryParse(mData(15), dData) Then
                                    stData.fpp32_y = Double.Parse(mData(15))                          ' �p�^�[�����W1Y
                                Else
                                    Call Z_PRINT("[" & No.ToString & "]�Ԗڂ̃f�[�^[" & mData(15) & "]�����l�ɕϊ��ł��܂���B")
                                    Return (False)
                                End If
                                If Short.TryParse(mData(16), dData) Then
                                    stData.iPP37_2 = Short.Parse(mData(16))                          ' �p�^�[���ԍ�2
                                Else
                                    Call Z_PRINT("[" & No.ToString & "]�Ԗڂ̃f�[�^[" & mData(16) & "]�����l�ɕϊ��ł��܂���B")
                                    Return (False)
                                End If
                                If Double.TryParse(mData(17), dData) Then
                                    stData.fpp33_x = Double.Parse(mData(17))                          ' �p�^�[�����W2X
                                Else
                                    Call Z_PRINT("[" & No.ToString & "]�Ԗڂ̃f�[�^[" & mData(17) & "]�����l�ɕϊ��ł��܂���B")
                                    Return (False)
                                End If
                                If Double.TryParse(mData(18), dData) Then
                                    stData.fpp33_y = Double.Parse(mData(18))                          ' �p�^�[�����W2Y
                                Else
                                    Call Z_PRINT("[" & No.ToString & "]�Ԗڂ̃f�[�^[" & mData(18) & "]�����l�ɕϊ��ł��܂���B")
                                    Return (False)
                                End If

                                stData.Comment = mData(19)                   ' �R�����g
                            Else
                                ' �Â��f�[�^�̏ꍇ�A�ƕ␳�֌W�̃p�����[�^���Ȃ��̂�8�����ɂȂ�
                                stData.Comment = mData(7)                   ' �R�����g
                            End If
                            'V2.2.1.6�A ��

                            Ret = True
                            Exit Do
                        End If
                    Loop
                End Using
            End If

            Return Ret

        Catch ex As Exception
            MsgBox("ReadProbeCsv() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (False)
        End Try
    End Function

#End Region


#Region "PROBEDATA�t�@�C���̓��e��S�ēǍ���"   'V2.2.0.0�N

    ''' <summary>
    ''' PROBEDATA�t�@�C���̓��e��S�ēǍ���
    ''' </summary>
    ''' <param name="stData"></param>
    ''' <param name="MaxNo"></param>
    ''' <returns></returns>
    Public Function ReadAllProbeCsv(ByRef stData() As stPROBEDATA_TABLE, ByRef MaxNo As Integer, ByRef header As String) As Boolean

        Dim sFolder As String = ""
        Dim sData As String = ""
        Dim bHeader As Boolean = True
        Dim mData() As String
        Dim TableNo As Integer
        Dim dData As Double
        Dim Ret As Boolean = False

        Try

            sFolder = cPROBEDATA_PATH & cPROBEDATA_FILE

            If IO.File.Exists(sFolder) = True Then  ' �t�@�C�����L��B
                Using sr As New System.IO.StreamReader(sFolder, System.Text.Encoding.GetEncoding("Shift-JIS"))
                    Do While Not sr.EndOfStream
                        sData = sr.ReadLine
                        If bHeader Then                                     ' �P�s�ڂ̓^�C�g���s
                            bHeader = False
                            header = sData
                            Continue Do
                        End If
                        mData = sData.Split(",")                            ' �������','�ŕ������Ď�o��
                        TableNo = Integer.Parse(mData(0))

                        If TableNo > MaxNo Then
                            MaxNo = TableNo
                        End If

                        stData(TableNo).No = TableNo

                        If Double.TryParse(mData(1), dData) Then
                            stData(TableNo).ProbeOn = Double.Parse(mData(1))                                ' �v���[�u�ڐG�ʒu
                        Else
                            Call Z_PRINT("[" & TableNo.ToString & "]�Ԗڂ̃f�[�^[" & mData(1) & "]�����l�ɕϊ��ł��܂���B")
                            Return (False)
                        End If

                        'V2.2.0.0�S��
                        If Double.TryParse(mData(2), dData) Then
                            stData(TableNo).ProbeOff = Double.Parse(mData(2))                                ' �v���[�u�ҋ@�ʒu
                        Else
                            Call Z_PRINT("[" & TableNo.ToString & "]�Ԗڂ̃f�[�^[" & mData(2) & "]�����l�ɕϊ��ł��܂���B")
                            Return (False)
                        End If
                        'V2.2.0.0�S��

                        If Double.TryParse(mData(3), dData) Then
                            stData(TableNo).dTableOffsetX = Double.Parse(mData(3))                          ' �e�[�u���I�t�Z�b�g�w
                        Else
                            Call Z_PRINT("[" & TableNo.ToString & "]�Ԗڂ̃f�[�^[" & mData(3) & "]�����l�ɕϊ��ł��܂���B")
                            Return (False)
                        End If

                        If Double.TryParse(mData(4), dData) Then
                            stData(TableNo).dTableOffsetY = Double.Parse(mData(4))                          ' �e�[�u���I�t�Z�b�g�x
                        Else
                            Call Z_PRINT("[" & TableNo.ToString & "]�Ԗڂ̃f�[�^[" & mData(4) & "]�����l�ɕϊ��ł��܂���B")
                            Return (False)
                        End If

                        If Double.TryParse(mData(5), dData) Then
                            stData(TableNo).dBPOffsetX = Double.Parse(mData(5))                          ' BP�I�t�Z�b�g�w
                        Else
                            Call Z_PRINT("[" & TableNo.ToString & "]�Ԗڂ̃f�[�^[" & mData(5) & "]�����l�ɕϊ��ł��܂���B")
                            Return (False)
                        End If

                        If Double.TryParse(mData(6), dData) Then
                            stData(TableNo).dBPOffsetY = Double.Parse(mData(6))                          ' BP�I�t�Z�b�g�x
                        Else
                            Call Z_PRINT("[" & TableNo.ToString & "]�Ԗڂ̃f�[�^[" & mData(6) & "]�����l�ɕϊ��ł��܂���B")
                            Return (False)
                        End If

                        'V2.2.1.6�A ��
                        If mData.Length > 8 Then
                            ' �V�����f�[�^�̏ꍇ�A�ƕ␳�֌W�̃p�����[�^���܂܂��̂�20�����ɂȂ�
                            If Short.TryParse(mData(7), dData) Then
                                stData(TableNo).iPP30 = Short.Parse(mData(7))                          ' �ʒu�␳���[�h
                            Else
                                Call Z_PRINT("[" & TableNo.ToString & "]�Ԗڂ̃f�[�^[" & mData(7) & "]�����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If
                            If Short.TryParse(mData(8), dData) Then
                                stData(TableNo).iPP31 = Short.Parse(mData(8))                          ' �ʒu�␳���@
                            Else
                                Call Z_PRINT("[" & TableNo.ToString & "]�Ԗڂ̃f�[�^[" & mData(8) & "]�����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If
                            If Double.TryParse(mData(9), dData) Then
                                stData(TableNo).fpp34_x = Double.Parse(mData(9))                          ' �␳�|�W�V�����I�t�Z�b�gX
                            Else
                                Call Z_PRINT("[" & TableNo.ToString & "]�Ԗڂ̃f�[�^[" & mData(9) & "]�����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If
                            If Double.TryParse(mData(10), dData) Then
                                stData(TableNo).fpp34_y = Double.Parse(mData(10))                          ' �␳�|�W�V�����I�t�Z�b�gY
                            Else
                                Call Z_PRINT("[" & TableNo.ToString & "]�Ԗڂ̃f�[�^[" & mData(10) & "]�����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If
                            If Double.TryParse(mData(11), dData) Then
                                stData(TableNo).fTheta = Double.Parse(mData(11))                          ' �p�x
                            Else
                                Call Z_PRINT("[" & TableNo.ToString & "]�Ԗڂ̃f�[�^[" & mData(11) & "]�����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If
                            If Short.TryParse(mData(12), dData) Then
                                stData(TableNo).iPP38 = Short.Parse(mData(12))                          ' �O���[�v�ԍ�
                            Else
                                Call Z_PRINT("[" & TableNo.ToString & "]�Ԗڂ̃f�[�^[" & mData(12) & "]�����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If
                            If Short.TryParse(mData(13), dData) Then
                                stData(TableNo).iPP37_1 = Short.Parse(mData(13))                          ' �p�^�[���ԍ�1
                            Else
                                Call Z_PRINT("[" & TableNo.ToString & "]�Ԗڂ̃f�[�^[" & mData(13) & "]�����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If
                            If Double.TryParse(mData(14), dData) Then
                                stData(TableNo).fpp32_x = Double.Parse(mData(14))                          ' �p�^�[�����W1X
                            Else
                                Call Z_PRINT("[" & TableNo.ToString & "]�Ԗڂ̃f�[�^[" & mData(14) & "]�����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If
                            If Double.TryParse(mData(15), dData) Then
                                stData(TableNo).fpp32_y = Double.Parse(mData(15))                          ' �p�^�[�����W1Y
                            Else
                                Call Z_PRINT("[" & TableNo.ToString & "]�Ԗڂ̃f�[�^[" & mData(15) & "]�����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If
                            If Short.TryParse(mData(16), dData) Then
                                stData(TableNo).iPP37_2 = Short.Parse(mData(16))                          ' �p�^�[���ԍ�2
                            Else
                                Call Z_PRINT("[" & TableNo.ToString & "]�Ԗڂ̃f�[�^[" & mData(16) & "]�����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If
                            If Double.TryParse(mData(17), dData) Then
                                stData(TableNo).fpp33_x = Double.Parse(mData(17))                          ' �p�^�[�����W2X
                            Else
                                Call Z_PRINT("[" & TableNo.ToString & "]�Ԗڂ̃f�[�^[" & mData(17) & "]�����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If
                            If Double.TryParse(mData(18), dData) Then
                                stData(TableNo).fpp33_y = Double.Parse(mData(18))                          ' �p�^�[�����W2Y
                            Else
                                Call Z_PRINT("[" & TableNo.ToString & "]�Ԗڂ̃f�[�^[" & mData(18) & "]�����l�ɕϊ��ł��܂���B")
                                Return (False)
                            End If
                            stData(TableNo).Comment = mData(19)                   ' �R�����g
                        Else
                            ' �Â��f�[�^�̏ꍇ�A�ƕ␳�֌W�̃p�����[�^���Ȃ��̂�8�����ɂȂ�
                            stData(TableNo).Comment = mData(7)                   ' �R�����g
                        End If
                        'V2.2.1.6�A ��

                        Ret = True
                    Loop
                End Using
            End If

            Return Ret

        Catch ex As Exception
            MsgBox("ReadAllProbeCsv() Execption error." & vbCrLf & " error msg = " & ex.Message)
            Return (False)
        End Try
    End Function

#End Region

#Region "�v���[�u�f�[�^����������"
    ''' <summary>
    '''     '�v���[�u�f�[�^����������    'V2.2.0.0�N
    ''' </summary>
    ''' <param name="stlocalProbeData"></param>
    ''' <param name="Maxno"></param>
    ''' <returns></returns>
    Public Function WriteAllProbeCsv(ByRef stData() As stPROBEDATA_TABLE, ByVal Maxno As Integer, ByVal sHeaderData As String) As Integer
        Dim sFileName As String

        Try

            sFileName = cPROBEDATA_PATH & cPROBEDATA_FILE

            Using WSR As New System.IO.StreamWriter(sFileName, False, System.Text.Encoding.GetEncoding("Shift-JIS"))  ' ��Q���� �㏑���́AFalse
                WSR.WriteLine(sHeaderData)                          ' �w�b�_�o��

                For No As Integer = 1 To Maxno
                    'V2.2.0.0�S                    WSR.WriteLine(stData(No).No.ToString & "," & stData(No).ProbeOn.ToString("0.000") & "," & stData(No).dTableOffsetX.ToString("0.000") & "," & stData(No).dTableOffsetY.ToString("0.000") & "," & stData(No).dBPOffsetX.ToString("0.000") & "," & stData(No).dBPOffsetY.ToString("0.000") & "," & stData(No).Comment)
                    WSR.WriteLine(stData(No).No.ToString & "," & stData(No).ProbeOn.ToString("0.000") & "," & stData(No).ProbeOff.ToString("0.000") & "," _
                                  & stData(No).dTableOffsetX.ToString("0.000") & "," & stData(No).dTableOffsetY.ToString("0.000") & "," _
                                  & stData(No).dBPOffsetX.ToString("0.000") & "," & stData(No).dBPOffsetY.ToString("0.000") & "," _
                                  & stData(No).iPP30.ToString() & "," & stData(No).iPP31.ToString() & "," & stData(No).fpp34_x.ToString("0.000") & "," _
                                  & stData(No).fpp34_y.ToString("0.000") & "," & stData(No).fTheta.ToString("0.000") & "," _
                                  & stData(No).iPP38.ToString() & "," & stData(No).iPP37_1.ToString() & "," & stData(No).fpp32_x.ToString("0.000") & "," & stData(No).fpp32_y.ToString("0.000") & "," _
                                  & stData(No).iPP37_2.ToString() & "," & stData(No).fpp33_x.ToString("0.000") & "," & stData(No).fpp33_y.ToString("0.000") & "," _
                                  & stData(No).Comment)     'V2.2.0.0�S
                Next
            End Using

        Catch ex As Exception

        End Try


    End Function

#End Region


End Module

'=============================== END OF FILE ===============================

