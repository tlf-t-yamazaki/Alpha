'==============================================================================
'   Description : ���[�U�v���O�����p�ŗL�t�@���N�V����
'
'�@ 2012/11/16 First Written by N.Arata(OLFT)
'
'==============================================================================
Option Strict Off
Option Explicit On

Imports System.Threading.Thread
Imports System.Runtime.InteropServices
Imports LaserFront.Trimmer.DllSysPrm.SysParam
Imports LaserFront.Trimmer.DefWin32Fnc
Imports UsrFunc.My.Resources
Imports LaserFront.Trimmer.DllJog                                       'V2.2.0.0�@

Module UserModule

#Region "��R��ʔ���"

#Region "�J�b�g��R����(�g���~���O�L��A�}�[�L���O�����j"
    '''=========================================================================
    ''' <summary>
    ''' �J�b�g��R����
    ''' </summary>
    ''' <param name="stRegData">��R�f�[�^�\����</param>
    ''' <param name="rn">��R�ԍ�</param>
    ''' <returns>True = �J�b�g����̒�R, False = ����݂̂̒�R</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsCutResistor(ByRef stRegData As Reg_Info(), ByVal rn As Integer) As Boolean
        Try
            If stRegData(rn).intSLP = SLP_VTRIMPLS Or stRegData(rn).intSLP = SLP_VTRIMMNS Or stRegData(rn).intSLP = SLP_RTRM Or stRegData(rn).intSLP = SLP_ATRIMPLS Or stRegData(rn).intSLP = SLP_ATRIMMNS Or stRegData(rn).intSLP = SLP_MARK Then 'V2.2.1.7�@
                IsCutResistor = True
            Else
                IsCutResistor = False
            End If
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function

    '''=========================================================================
    ''' <summary>
    ''' �J�b�g��R����
    ''' </summary>
    ''' <param name="rn">��R�ԍ�</param>
    ''' <returns>True = �J�b�g����̒�R, False = ����݂̂̒�R</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsCutResistor(ByVal rn As Integer) As Boolean
        Try
            IsCutResistor = IsCutResistor(stREG, rn)
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region

#Region "�}�[�L���O���܂߂��J�b�g��R����(�g���~���O�L��A�}�[�L���O�L��j"
    '''=========================================================================
    ''' <summary>
    ''' �}�[�L���O���܂߂��J�b�g��R����
    ''' </summary>
    ''' <param name="stRegData">��R�f�[�^�\����</param>
    ''' <param name="rn">��R�ԍ�</param>
    ''' <returns>True = �J�b�g����̒�R, False = ����݂̂̒�R</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsCutResistorIncMarking(ByRef stRegData As Reg_Info(), ByVal rn As Integer) As Boolean
        Try
            If stRegData(rn).intSLP = SLP_VTRIMPLS Or stRegData(rn).intSLP = SLP_VTRIMMNS Or stRegData(rn).intSLP = SLP_RTRM Or stRegData(rn).intSLP = SLP_ATRIMPLS Or stRegData(rn).intSLP = SLP_ATRIMMNS Or stRegData(rn).intSLP = SLP_NG_MARK Or stRegData(rn).intSLP = SLP_OK_MARK Or stRegData(rn).intSLP = SLP_MARK Then 'V2.2.1.7�@
                IsCutResistorIncMarking = True
            Else
                IsCutResistorIncMarking = False
            End If
        Catch ex As Exception
            Call Z_PRINT("UserModule.IsCutResistorIncMarking() TRAP ERROR = " & ex.Message & vbCrLf)
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function

    '''=========================================================================
    ''' <summary>
    ''' �}�[�L���O���܂߂��J�b�g��R����
    ''' </summary>
    ''' <param name="rn">��R�ԍ�</param>
    ''' <returns>True = �J�b�g����̒�R, False = ����݂̂̒�R</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsCutResistorIncMarking(ByVal rn As Integer) As Boolean
        Try
            IsCutResistorIncMarking = IsCutResistorIncMarking(stREG, rn)
        Catch ex As Exception
            Call Z_PRINT("UserModule.IsCutResistorIncMarking() TRAP ERROR = " & ex.Message & vbCrLf)
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region

#Region "�}�[�L���O�J�b�g��R����(�}�[�L���O�L��̂݁j"
    '''=========================================================================
    ''' <summary>
    ''' �}�[�L���O���܂߂��J�b�g��R����
    ''' </summary>
    ''' <param name="stRegData">��R�f�[�^�\����</param>
    ''' <param name="rn">��R�ԍ�</param>
    ''' <returns>True = �J�b�g����̒�R, False = ����݂̂̒�R</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsMarking(ByRef stRegData As Reg_Info(), ByVal rn As Integer) As Boolean
        Try
            If stRegData(rn).intSLP = SLP_NG_MARK Or stRegData(rn).intSLP = SLP_OK_MARK Or stRegData(rn).intSLP = SLP_MARK Then 'V2.2.1.7�@
                IsMarking = True
            Else
                IsMarking = False
            End If
        Catch ex As Exception
            Call Z_PRINT("UserModule.IsMarking() TRAP ERROR = " & ex.Message & vbCrLf)
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function

    '''=========================================================================
    ''' <summary>
    ''' �}�[�L���O���܂߂��J�b�g��R����
    ''' </summary>
    ''' <param name="stRegData">��R�f�[�^�\����</param>
    ''' <param name="rn">��R�ԍ�</param>
    ''' <returns>True = �J�b�g����̒�R, False = ����݂̂̒�R</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsOkMarking(ByRef stRegData As Reg_Info(), ByVal rn As Integer) As Boolean
        Try
            If stRegData(rn).intSLP = SLP_OK_MARK Then
                IsOkMarking = True
            Else
                IsOkMarking = False
            End If
        Catch ex As Exception
            Call Z_PRINT("UserModule.IsMarking() TRAP ERROR = " & ex.Message & vbCrLf)
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
    '''=========================================================================
    ''' <summary>
    ''' �}�[�L���O���܂߂��J�b�g��R����
    ''' </summary>
    ''' <param name="rn">��R�ԍ�</param>
    ''' <returns>True = �J�b�g����̒�R, False = ����݂̂̒�R</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsMarking(ByVal rn As Integer) As Boolean
        Try
            IsMarking = IsMarking(stREG, rn)
        Catch ex As Exception
            Call Z_PRINT("UserModule.IsMarking() TRAP ERROR = " & ex.Message & vbCrLf)
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region

#Region "�w���R�ԍ��ȍ~�̃}�[�L���O�J�b�g�̒�R�ԍ���Ԃ��B"
    ''' <summary>
    ''' �w���R�ԍ��ȍ~�̃}�[�L���O�J�b�g�̒�R�ԍ��擾
    ''' </summary>
    ''' <param name="rn">��R�ԍ�</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetMarkingResNo(ByVal rn As Integer) As Integer
        GetMarkingResNo = 0
        Try
            Dim iResNo As Integer
            For iResNo = rn + 1 To stPLT.RCount Step 1
                If UserModule.IsMarking(iResNo) Then    '�}�[�L���O�f�[�^
                    GetMarkingResNo = iResNo
                    Return (GetMarkingResNo)
                End If
            Next
        Catch ex As Exception
            Call Z_PRINT("UserModule.IsCutResistor() TRAP ERROR = " & ex.Message & vbCrLf)
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region

#Region "OK�}�[�L���O�Ƃ��Ă̘A�Ԃ�Ԃ��B"
    ''' <summary>
    ''' OK�}�[�L���O�Ƃ��Ă̘A�Ԃ�Ԃ�
    ''' </summary>
    ''' <param name="rn">��R�ԍ�</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetOkMarkingResNo(ByVal rn As Integer) As Integer
        GetOkMarkingResNo = 0
        Try
            Dim iResNo As Integer
            Dim OKCount As Integer = 0
            For iResNo = 1 To stPLT.RCount Step 1
                If UserModule.IsOkMarking(stREG, iResNo) Then       'OK�}�[�L���O�f�[�^
                    OKCount = OKCount + 1
                    If rn = iResNo Then
                        GetOkMarkingResNo = OKCount
                        Return (GetOkMarkingResNo)
                    End If
                End If
            Next
        Catch ex As Exception
            Call Z_PRINT("UserModule.IsCutResistor() TRAP ERROR = " & ex.Message & vbCrLf)
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region

#Region "����݂̂��̔���"
    '''=========================================================================
    ''' <summary>
    ''' ���蔻��
    ''' </summary>
    ''' <param name="stRegData">��R�f�[�^�\����</param>
    ''' <param name="rn">��R�ԍ�</param>
    ''' <returns>True = ����݂̂̒�R, False = �J�b�g����̒�R</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsMeasureOnly(ByRef stRegData As Reg_Info(), ByVal rn As Short) As Boolean
        Try
            If stRegData(rn).intSLP = SLP_VMES Or stRegData(rn).intSLP = SLP_AMES Or stRegData(rn).intSLP = SLP_RMES Then
                IsMeasureOnly = True
            Else
                IsMeasureOnly = False
            End If
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function

    '''=========================================================================
    ''' <summary>
    ''' ���蔻��
    ''' </summary>
    ''' <param name="rn">��R�ԍ�</param>
    ''' <returns>True = ����݂̂̒�R, False = �J�b�g����̒�R</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsMeasureOnly(ByVal rn As Short) As Boolean
        Try
            IsMeasureOnly = IsMeasureOnly(stREG, rn)
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region

#Region "���肩�̔���"
    '''=========================================================================
    ''' <summary>
    ''' ���蔻��
    ''' </summary>
    ''' <param name="stRegData">��R�f�[�^�\����</param>
    ''' <param name="rn">��R�ԍ�</param>
    ''' <param name="MeasMode">���胂�[�h(0:�Ȃ�, 1:IT�̂� 2:FT�̂� 3:IT,FT����)</param>
    ''' <returns>True:����L�� False:���薳��</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsMeasureMode(ByRef stRegData As Reg_Info(), ByVal rn As Short, ByVal MeasMode As Short) As Boolean
        Try
            IsMeasureMode = False

            If IsMarking(rn) Then                   ' �}�[�L���O��R�͑���ΏۊO
                Return (False)
            End If

            If MeasMode = MEAS_JUDGE_IT Then
                If stREG(rn).intMeasMode = MEAS_JUDGE_IT Or stREG(rn).intMeasMode = MEAS_JUDGE_BOTH Then
                    IsMeasureMode = True
                End If
            ElseIf MeasMode = MEAS_JUDGE_FT Then
                If stREG(rn).intMeasMode = MEAS_JUDGE_FT Or stREG(rn).intMeasMode = MEAS_JUDGE_BOTH Then
                    IsMeasureMode = True
                End If
            Else
                Call Z_PRINT("CheckMeasureMode:�w�肳�ꂽ���[�h���������L��܂���=[" & MeasMode.ToString() & "]")
                IsMeasureMode = False
            End If
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function

    '''=========================================================================
    ''' <summary>
    ''' ���蔻��
    ''' </summary>
    ''' <param name="rn">��R�ԍ�</param>
    ''' <param name="MeasMode">���胂�[�h(0:�Ȃ�, 1:IT�̂� 2:FT�̂� 3:IT,FT����)</param>
    ''' <returns>True:����L�� False:���薳��</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsMeasureMode(ByVal rn As Short, ByVal MeasMode As Short) As Boolean
        Try
            IsMeasureMode = IsMeasureMode(stREG, rn, MeasMode)
            'V2.1.0.0�D����R�l�ω��ʔ���̏ꍇ�́A�K���������肪�K�v
            If MeasMode = MEAS_JUDGE_IT AndAlso IsMeasureMode = False Then
                If UserSub.IsCutVariationJudgeExecute() AndAlso UserModule.IsCutResistor(stREG, rn) Then
                    IsMeasureMode = True
                End If
            End If
            'V2.1.0.0�D��
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region

#Region "�~�T���胂�[�h�Ώۂ̒�R���𔻒�"
    ''' <summary>
    ''' �~�T���胂�[�h�Ώۂ̒�R���𔻒肷��B
    ''' </summary>
    ''' <param name="rn"></param>
    ''' <returns>True:�Ώ� False:���薳��</returns>
    ''' <remarks></remarks>
    Public Function IsMeasureResistor(ByVal rn As Short) As Boolean
        Try
            If IsMarking(rn) Then       ' �}�[�L���O��R�͑���ΏۊO
                Return (False)
            End If
            If stREG(rn).intMeasMode = MEAS_JUDGE_IT Or stREG(rn).intMeasMode = MEAS_JUDGE_FT Or stREG(rn).intMeasMode = MEAS_JUDGE_BOTH Then
                Return (True)
            Else
                Return (False)
            End If
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region

#Region "�w�肵����R�ԍ��ƒ�R������v���邩�̃`�F�b�N"
    '''=========================================================================
    ''' <summary>
    ''' ��R�ԍ��ƒ�R�������S��v���邩�`�F�b�N����
    ''' </summary>
    ''' <param name="iResNo">��R�ԍ�</param>
    ''' <param name="strRNO">��R��</param>
    ''' <returns>True�F��v False:�s��v</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsResistorByName(ByVal iResNo As Integer, ByVal strRNO As String) As Boolean
        Try
            If stREG(iResNo).strRNO.Trim.CompareTo(strRNO.Trim) = 0 Then
                Return (True)
            Else
                Return (False)
            End If
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
    '''=========================================================================
    ''' <summary>
    ''' ��R�ԍ��ƒ�R�����O��������v���邩�`�F�b�N����
    ''' </summary>
    ''' <param name="iResNo">��R�ԍ�</param>
    ''' <param name="strRNO">��R��</param>
    ''' <returns>True�F��v False:�s��v</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function IsResistorByNameStartsWith(ByVal iResNo As Integer, ByVal strRNO As String) As Boolean
        Try
            If stREG(iResNo).strRNO.Trim.StartsWith(strRNO.Trim) Then
                Return (True)
            Else
                Return (False)
            End If
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function

#End Region

#Region "��R������Y�������R�ԍ����擾����"
    '''=========================================================================
    ''' <summary>
    ''' ��R������Y�������R�ԍ����擾����
    ''' </summary>
    ''' <param name="strRNO">��R�ԍ��i������j</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function GetResistorNoByName(ByVal strRNO As String) As Integer
        GetResistorNoByName = 0
        Try
            Dim iResNo As Integer

            For iResNo = 1 To stPLT.RCount Step 1
                If IsResistorByName(iResNo, strRNO) Then
                    Return (iResNo)
                End If
            Next

        Catch ex As Exception
            Call Z_PRINT("UserModule.GetResistorNoByName() TRAP ERROR = " & ex.Message & vbCrLf)
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region

#End Region

#Region "���� ����݂̂���������R���̎擾 ����"
#If False Then
    ''' <summary>
    ''' ����݂̂���������R���̎擾
    ''' </summary>
    ''' <param name="stPlate">�v���[�g�f�[�^�\����</param>
    ''' <param name="stRegData">��R�f�[�^�\����</param>
    ''' <returns>��R��</returns>
    ''' <remarks></remarks>

    Public Function GetRCountExceptMeasure(ByVal stPlate As PLATE_DATA, ByRef stRegData As Reg_Info()) As Short
        Try
            GetRCountExceptMeasure = 0
            For iResNo As Integer = 1 To stPlate.RCount Step 1
                If IsCutResistor(stRegData, iResNo) Then           ' �J�b�g�L�i����݂̂łȂ��j��R�̏ꍇ
                    GetRCountExceptMeasure = GetRCountExceptMeasure + 1
                End If
            Next

        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End If
#End Region

#Region "��R�f�[�^�̃R�s�[����"
    ''' <summary>
    ''' ��R�f�[�^�P���R�[�h�̃R�s�[����
    ''' </summary>
    ''' <param name="ToRes">�R�s�[��</param>
    ''' <param name="FromRes">�R�s�[��</param>
    ''' <remarks></remarks>
    Public Sub CopyResistorData(ByRef ToRes As Reg_Info, ByRef FromRes As Reg_Info)

        Try

            ToRes = FromRes
            ToRes.STCUT = DirectCast(ToRes.STCUT.Clone(), Cut_Info())
            ToRes.intOnExtEqu = DirectCast(ToRes.intOnExtEqu.Clone(), Short())      ' �n�m�@��P�`�R
            ToRes.intOffExtEqu = DirectCast(ToRes.intOffExtEqu.Clone(), Short())    ' �n�e�e�@��P�`�R

            For i As Integer = 1 To MAXCTN
                ToRes.STCUT(i).intCND = DirectCast(ToRes.STCUT(i).intCND.Clone(), Short())
                ToRes.STCUT(i).intIXN = DirectCast(ToRes.STCUT(i).intIXN.Clone(), Short())
                ToRes.STCUT(i).dblDL1 = DirectCast(ToRes.STCUT(i).dblDL1.Clone(), Double())
                ToRes.STCUT(i).lngPAU = DirectCast(ToRes.STCUT(i).lngPAU.Clone(), Integer())
                ToRes.STCUT(i).dblDEV = DirectCast(ToRes.STCUT(i).dblDEV.Clone(), Double())
                ToRes.STCUT(i).intIXMType = DirectCast(ToRes.STCUT(i).intIXMType.Clone(), Short())
                ToRes.STCUT(i).intIXTMM = DirectCast(ToRes.STCUT(i).intIXTMM.Clone(), Short())
            Next

            For i As Integer = 1 To MAX_LCUT
                ToRes.STCUT(i).dCutLen = DirectCast(ToRes.STCUT(i).dCutLen.Clone(), Double())
                ToRes.STCUT(i).dQRate = DirectCast(ToRes.STCUT(i).dQRate.Clone(), Double())
                ToRes.STCUT(i).dSpeed = DirectCast(ToRes.STCUT(i).dSpeed.Clone(), Double())
                ToRes.STCUT(i).dAngle = DirectCast(ToRes.STCUT(i).dAngle.Clone(), Double())
                ToRes.STCUT(i).dTurnPoint = DirectCast(ToRes.STCUT(i).dTurnPoint.Clone(), Double())
            Next

            'V2.0.0.2�A��
            For i As Integer = 1 To MAX_RETRACECUT
                ToRes.STCUT(i).dblRetraceOffX = DirectCast(ToRes.STCUT(i).dblRetraceOffX.Clone(), Double())
                ToRes.STCUT(i).dblRetraceOffY = DirectCast(ToRes.STCUT(i).dblRetraceOffY.Clone(), Double())
                ToRes.STCUT(i).dblRetraceQrate = DirectCast(ToRes.STCUT(i).dblRetraceQrate.Clone(), Double())
                ToRes.STCUT(i).dblRetraceSpeed = DirectCast(ToRes.STCUT(i).dblRetraceSpeed.Clone(), Double())
            Next
            'V2.0.0.2�A��

        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try

    End Sub

    ''' <summary>
    ''' ���[�U�f�[�^�̃R�s�[����
    ''' </summary>
    ''' <param name="ToUser">�R�s�[��</param>
    ''' <param name="FromUser">�R�s�[��</param>
    ''' <remarks></remarks>
    Public Sub CopyUserData(ByRef ToUser As USER_DATA, ByRef FromUser As USER_DATA)
        Try
            ToUser.Initialize()
            For i As Integer = 0 To 1                                      ' MAXBLKX
                For j As Integer = 0 To 1                                  ' MAXBLKY
                Next j
            Next i

            ToUser = FromUser
            ToUser.iResUnit = DirectCast(FromUser.iResUnit.Clone(), Integer())
            ToUser.dNomCalcCoff = DirectCast(FromUser.dNomCalcCoff.Clone(), Double())
            ToUser.dTargetCoff = DirectCast(FromUser.dTargetCoff.Clone(), Double())
            ToUser.iChangeSpeed = DirectCast(FromUser.iChangeSpeed.Clone(), Integer())
            ToUser.dItVal = DirectCast(FromUser.dItVal.Clone(), Double())
            ToUser.dFtVal = DirectCast(FromUser.dFtVal.Clone(), Double())
            ToUser.dDev = DirectCast(FromUser.dDev.Clone(), Double())

        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub

    ''' <summary>
    ''' �S�Ă̒�R�f�[�^�̃R�s�[�����i�S�z��j
    ''' </summary>
    ''' <param name="stPlate">�v���[�g�f�[�^�\����</param>
    ''' <param name="ToRes">�R�s�[��</param>
    ''' <param name="FromRes">�R�s�[��</param>
    ''' <remarks></remarks>
    Public Sub CopyResistorDataArray(ByVal stPlate As PLATE_DATA, ByRef ToRes As Reg_Info(), ByRef FromRes As Reg_Info())

        Try
            For i As Integer = 1 To MAXRNO Step 1
                ToRes(i).Initialize()
                Call CopyResistorData(ToRes(i), FromRes(i))
            Next

        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try

    End Sub
#End Region

#Region "�w�肵���ʒu�������փJ�b�g�f�[�^���R�s�[����B"
#If False Then
    ''' <summary>
    ''' �w�肵���ʒu����w�肵���ʒu�܂ŃJ�b�g�f�[�^���R�s�[����B
    ''' </summary>
    ''' <param name="stPlate">�v���[�g�f�[�^�\����</param>
    ''' <param name="stRegData">��R�f�[�^�\����</param>
    ''' <param name="Res">�R�s�[���̒�R�ԍ�</param>
    ''' <param name="FromCut">�R�s�[���̃J�b�g�ԍ�</param>
    ''' <param name="ToCut">�R�s�[������Ō�̃J�b�g�ԍ�</param>
    ''' <param name="bForce">�����R�s�[���[�h</param>
    ''' <param name="bCut">�J�b�g���@�̃R�s�[�L��</param>
    ''' <param name="bCTYP">�J�b�g�`��̃R�s�[�L��</param>
    ''' <param name="bLen">�J�b�g���̃R�s�[�L��</param>
    ''' <param name="bANG">�J�b�g�����̃R�s�[�L��</param>
    ''' <param name="bSpeed">�J�b�g���x�̃R�s�[�L��</param>
    ''' <param name="bCutCnd">�J�b�g����</param>
    ''' <remarks>�S�Ă̒�R�f�[�^��ΏۂƂ���B</remarks>
    Public Sub CutDataCopy(ByVal stPlate As PLATE_DATA, ByRef stRegData As Reg_Info(), ByVal Res As Integer, ByVal FromCut As Integer, ByVal ToCut As Integer, ByVal bForce As Boolean, ByVal bCut As Boolean, ByVal bCTYP As Boolean, ByVal bLen As Boolean, ByVal bANG As Boolean, ByVal bSpeed As Boolean, ByVal bCutCnd As Boolean)

        Try

            Dim iFrom As Integer

            If ToCut > MAXCTN Then
                ToCut = MAXCTN
            End If

            For iR As Integer = Res To stPlate.RCount Step 1                ' ��R�͎w���R�ԍ������̑S�Ă̔ԍ����Ώ�
                If IsCutResistor(stRegData, iR) Then                        ' �J�b�g�L�i����݂̂łȂ��j��R�̏ꍇ
                    iFrom = FromCut
                    If stRegData(iR).intTNN < iFrom Then                    ' �����J�b�g�f�[�^��������Ȃ����́A����Ȃ��ԍ�����R�s�[����B
                        iFrom = stRegData(iR).intTNN + 1
                    End If
                    For iC As Integer = iFrom To ToCut Step 1               ' �J�b�g�́A�R�s�[���̃J�b�g�ԍ��̎�����֌����R�s�[����B
                        If iR = Res And iC = FromCut Then                   ' �R�s�[���Ɛ悪�����ꍇ�̓X�L�b�v
                            Continue For
                        End If
                        If stRegData(iR).intTNN < iC Or bForce Then             ' �J�b�g�f�[�^���V�K�̏ꍇ�܂��͋����R�s�[���[�h�̎��͑S�Ă��R�s�[����B
                            stRegData(iR).STCUT(iC) = stRegData(Res).STCUT(FromCut)
                            stRegData(iR).STCUT(iC).intCND = DirectCast(stRegData(iR).STCUT(iC).intCND.Clone(), Short())
                            stRegData(iR).STCUT(iC).intIXN = DirectCast(stRegData(iR).STCUT(iC).intIXN.Clone(), Short())
                            stRegData(iR).STCUT(iC).dblDL1 = DirectCast(stRegData(iR).STCUT(iC).dblDL1.Clone(), Double())
                            stRegData(iR).STCUT(iC).lngPAU = DirectCast(stRegData(iR).STCUT(iC).lngPAU.Clone(), Integer())
                            stRegData(iR).STCUT(iC).dblDEV = DirectCast(stRegData(iR).STCUT(iC).dblDEV.Clone(), Double())
                            stRegData(iR).STCUT(iC).intIXMType = DirectCast(stRegData(iR).STCUT(iC).intIXMType.Clone(), Short())
                            stRegData(iR).STCUT(iC).intIXTMM = DirectCast(stRegData(iR).STCUT(iC).intIXTMM.Clone(), Short())
                        Else                                                ' �����̃J�b�g�f�[�^�̏ꍇ�͎w�肳�ꂽ���ڂ����R�s�[����B
                            If bCut Then                                    ' �J�b�g���@
                                stRegData(iR).STCUT(iC).intCUT = stRegData(Res).STCUT(FromCut).intCUT
                            End If
                            If bCTYP Then                                   ' �J�b�g�`��
                                stRegData(iR).STCUT(iC).intCTYP = stRegData(Res).STCUT(FromCut).intCTYP
                            End If
                            If bLen Then                                    ' �J�b�g��
                                stRegData(iR).STCUT(iC).dblDL2 = stRegData(Res).STCUT(FromCut).dblDL2
                                stRegData(iR).STCUT(iC).dblDL3 = stRegData(Res).STCUT(FromCut).dblDL3
                                For iX As Integer = 1 To MAXIDX Step 1
                                    stRegData(iR).STCUT(iC).dblDL1(iX) = stRegData(iR).STCUT(iC).dblDL1(iX)
                                Next iX
                            End If
                            If bANG Then                                    ' �J�b�g����
                                stRegData(iR).STCUT(iC).intANG = stRegData(Res).STCUT(FromCut).intANG
                                stRegData(iR).STCUT(iC).intANG2 = stRegData(Res).STCUT(FromCut).intANG2
                            End If
                            If bSpeed Then                                  ' �J�b�g���x
                                stRegData(iR).STCUT(iC).dblV1 = stRegData(Res).STCUT(FromCut).dblV1
                            End If
                            If bCutCnd Then                                 ' �J�b�g����
                                For i As Integer = 1 To MAXCND
                                    stRegData(iR).STCUT(iC).intCND(i) = stRegData(Res).STCUT(FromCut).intCND(i)
                                Next
                            End If
                        End If
                    Next iC
                    If stRegData(iR).intTNN < ToCut Then
                        stRegData(iR).intTNN = ToCut                            ' �S�Ă̒�R�f�[�^�ŃJ�b�g�������킹��B    
                    End If
                End If
            Next iR
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub
#End If
#End Region

#Region "�w�肵���ʒu�������֒�R�f�[�^���R�s�[����B"
    ''' <summary>
    ''' ��R�f�[�^�̃R�s�[
    ''' </summary>
    ''' <param name="stRegData">��R�f�[�^�\����</param>
    ''' <param name="FromRes">�R�s�[���̒�R�ԍ�</param>
    ''' <param name="ToRes">�R�s�[������Ō�̒�R�ԍ�</param>
    ''' <remarks>�v���[�u�f�[�^�ƃJ�b�g�f�[�^�̓R�s�[���Ȃ�</remarks>
    Public Sub ResistorDataCopy(ByRef stRegData As Reg_Info(), ByVal FromRes As Integer, ByVal ToRes As Integer)

        Try
            Dim iFrom As Integer, iNo As Integer
            Dim sResNumber As String = ""
            Dim sResName As String = ""

            iFrom = FromRes + 1

            For i As Integer = 0 To stRegData(FromRes).strRNO.Length - 1
                If Char.IsNumber(stRegData(FromRes).strRNO.Chars(i)) Then
                    sResNumber = sResNumber + stRegData(FromRes).strRNO.Chars(i)
                Else
                    sResName = sResName + stRegData(FromRes).strRNO.Chars(i)
                End If
            Next i
            If sResName.Equals(String.Empty) Then   ' ������͋�ł�
                sResName = "R"
            End If
            If sResNumber.Equals(String.Empty) Then ' ������͋�ł�
                iNo = FromRes
            Else
                iNo = Integer.Parse(sResNumber)
            End If

            For iR As Integer = iFrom To ToRes Step 1                       ' ��R�͎w���R�ԍ������̑S�Ă̔ԍ����Ώ�
                iNo = iNo + 1
                stRegData(iR).strRNO = sResName & iNo.ToString("0")     ' ��R��
                stRegData(iR).strTANI = stRegData(FromRes).strTANI      ' �P��("V","��" ��)
                stRegData(iR).intSLP = stRegData(FromRes).intSLP        ' �d���ω��X���[�v(1:+V, 2:-V, 4:��R, 5:�d������̂�, 6:��R����̂� 7:NGϰ�ݸ�)
                stRegData(iR).lngRel = stRegData(FromRes).lngRel        ' �����[�r�b�g
                stRegData(iR).dblNOM = stRegData(FromRes).dblNOM        ' �g���~���O�ڕW�l
                stRegData(iR).dblITL = stRegData(FromRes).dblITL        ' �������艺���l (ITLO)
                stRegData(iR).dblITH = stRegData(FromRes).dblITH        ' �����������l (ITHI)
                stRegData(iR).dblFTL = stRegData(FromRes).dblFTL        ' �I�����艺���l (FTLO)
                stRegData(iR).dblFTH = stRegData(FromRes).dblFTH        ' �I���������l (FTHI)
                stRegData(iR).intMode = stRegData(FromRes).intMode      ' ���胂�[�h(0:�䗦(%), 1:���l(��Βl))
                stRegData(iR).intTMM1 = stRegData(FromRes).intTMM1      ' ���[�h(0:����(�R���p���[�^��ϕ����[�h), 1:�����x(�ϕ����[�h))
                'stRegData(iR).intPRH = stRegData(FromRes).intPRH        ' �n�C���v���[�u�ԍ�(High Probe No.)
                'stRegData(iR).intPRL = stRegData(FromRes).intPRL        ' ���[���v���[�u�ԍ�(Low Probe No.)
                'stRegData(iR).intPRG = stRegData(FromRes).intPRG        ' �K�[�h�v���[�u�ԍ�(Gaude probe No.)
                stRegData(iR).intMType = stRegData(FromRes).intMType    ' ������(0=��������, 1=�O������)
            Next iR
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub
#End Region

#Region "�w�肵���ʒu�������փv���[�u�ԍ��𑝌�����B"
    Public Sub ProbeNumberIncDec(ByRef stRegData As Reg_Info(), ByVal FromRes As Integer, ByVal ToRes As Integer, ByVal PRHIncDec As Integer, ByVal PRLIncDec As Integer)

        Try

            Dim iFrom As Integer
            Dim iDiffHI As Integer, iDiffLO As Integer

            iFrom = FromRes + 1
            For iR As Integer = iFrom To ToRes Step 1                       ' ��R�͎w���R�ԍ������̑S�Ă̔ԍ����Ώ�
                iDiffHI = PRHIncDec * (iR - FromRes)
                iDiffLO = PRLIncDec * (iR - FromRes)
                stRegData(iR).intPRH = stRegData(FromRes).intPRH + iDiffHI  ' �n�C���v���[�u�ԍ�(High Probe No.)
                stRegData(iR).intPRL = stRegData(FromRes).intPRL + iDiffLO  ' ���[���v���[�u�ԍ�(Low Probe No.)
            Next iR
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub
#End Region

#Region "�s�w�C�s�x�␳�O�̃f�[�^�ݒ�"
    ''' <summary>
    ''' �s�w�C�s�x�␳�O�̃f�[�^�ݒ�
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SetTrimDataForTXTY()

        Try
            Dim ResNo As Integer = 0
            For rn As Integer = 1 To stPLT.RCount
                If UserModule.IsCutResistor(rn) Then
                    ResNo = ResNo + 1
                    typResistorInfoArray(ResNo).intResNo = ResNo
                    typResistorInfoArray(ResNo).intCutCount = stREG(rn).intTNN
                    typResistorInfoArray(ResNo).Initialize()
                    For cn As Integer = 1 To stREG(rn).intTNN
                        typResistorInfoArray(ResNo).ArrCut(cn).intCutNo = cn
                        typResistorInfoArray(ResNo).ArrCut(cn).dblStartPointX = stREG(rn).STCUT(cn).dblSTX
                        typResistorInfoArray(ResNo).ArrCut(cn).dblStartPointY = stREG(rn).STCUT(cn).dblSTY
                        typResistorInfoArray(ResNo).ArrCut(cn).dblTeachPointX = stREG(rn).STCUT(cn).dblSTX
                        typResistorInfoArray(ResNo).ArrCut(cn).dblTeachPointY = stREG(rn).STCUT(cn).dblSTY
                    Next cn
                End If
            Next rn

            typPlateInfo.intBlockCntXDir = stPLT.BNX                ' �u���b�N��X
            typPlateInfo.intBlockCntYDir = stPLT.BNY                ' �u���b�N��Y
            typPlateInfo.dblBlockSizeXDir = stPLT.zsx               ' �u���b�N(��R)�T�C�Yx(mm)
            typPlateInfo.dblBlockSizeYDir = stPLT.zsy               ' �u���b�N(��R)�T�C�Yy(mm)
            typPlateInfo.dblTableOffsetXDir = stPLT.z_xoff          ' �e�[�u���ʒu�I�t�Z�b�gX���g�����|�W�V�����I�t�Z�b�gX(mm)
            typPlateInfo.dblTableOffsetYDir = stPLT.z_yoff          ' �e�[�u���ʒu�I�t�Z�b�gY���g�����|�W�V�����I�t�Z�b�gY(mm)
            typPlateInfo.dblBpOffSetXDir = stPLT.BPOX               ' BP Offset X(mm)
            typPlateInfo.dblBpOffSetYDir = stPLT.BPOY               ' BP Offset Y(mm)
            'V2.0.0.0�@            typPlateInfo.intResistDir = 0                           ' ��R���ѕ����O�͂w�����A�P�͂x����
            typPlateInfo.intResistDir = Integer.Parse(GetPrivateProfileString_S("USER", "TXTY_DIRECTION", SYSPARAMPATH, "1"))   'V2.0.0.0�@
            typPlateInfo.intResistCntInBlock = ResNo                ' 1�u���b�N����R��=1�O���[�v����R��=��R��
            typPlateInfo.intResistCntInGroup = ResNo                ' 1�u���b�N����R��=1�O���[�v����R��=��R��
            If UserSub.IsTrimType3() Then
                typPlateInfo.intResistCntInGroup = UserSub.GetCircuitSum(stPLT, stREG)
            End If
            typPlateInfo.intGroupCntInBlockXBp = 1                  ' �u���b�N���a�o�O���[�v��(�T�[�L�b�g��)
            typPlateInfo.intGroupCntInBlockYStage = 1               ' �u���b�N���X�e�[�W�O���[�v��
            typPlateInfo.dblChipSizeXDir = stPLT.dblChipSizeXDir    ' �`�b�v�T�C�YX
            typPlateInfo.dblChipSizeXDir = stPLT.zsx                ' �`�b�v�T�C�YX�@V2.0.0.0�@�`�b�v�T�C�Y���O�Ȃ̂ŁA�u���b�N�T�C�Y��ݒ肷��B
            typPlateInfo.dblChipSizeYDir = stPLT.dblChipSizeYDir    ' �`�b�v�T�C�YY
            typPlateInfo.dblStepOffsetXDir = 0                      ' �X�e�b�v�I�t�Z�b�g��X
            typPlateInfo.dblStepOffsetYDir = 0                      ' �X�e�b�v�I�t�Z�b�g��Y
            typPlateInfo.dblBpGrpItv = 0                            ' BP�O���[�v�Ԋu�i�ȑO��CHIP�̃O���[�v�Ԋu�j
            typPlateInfo.dblStgGrpItvX = 0                          ' X�����̃X�e�[�W�O���[�v�Ԋu�i�ȑO�̂b�g�h�o�̃X�e�b�v�ԃC���^�[�o���j
            typPlateInfo.dblStgGrpItvY = 0                          ' Y�����̃X�e�[�W�O���[�v�Ԋu�i�ȑO�̂b�g�h�o�̃X�e�b�v�ԃC���^�[�o���j
            typPlateInfo.intBlkCntInStgGrpX = stPLT.BNX             ' X�����̃X�e�[�W�O���[�v���u���b�N��
            typPlateInfo.intBlkCntInStgGrpY = stPLT.BNY             ' Y�����̃X�e�[�W�O���[�v���u���b�N��

            gfCorrectPosX = dblCorrectX                             ' �w�x���␳�w
            gfCorrectPosY = dblCorrectY                             ' �w�x���␳�x

        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try

    End Sub
#End Region

#Region "�s�w�C�s�x�␳��̃f�[�^�ݒ�"
    ''' <summary>
    ''' �s�w�C�s�x�␳��̃f�[�^�ݒ�
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SetTrimDataFromTXTY()
        Try
            'If stPLT.dblChipSizeXDir <> 0.0 Then
            '    stPLT.dblChipSizeXDir = typPlateInfo.dblChipSizeXDir    ' �X�e�b�v�T�C�YX
            'End If
            'stPLT.zsx = typPlateInfo.dblBlockSizeXDir                   ' �u���b�N(��R)�T�C�Yx(mm)
            'stPLT.zsy = typPlateInfo.dblBlockSizeYDir                   ' �u���b�N(��R)�T�C�Yy(mm)
            If giAppMode = APP_MODE_TY Then
                stPLT.zsx = typPlateInfo.dblChipSizeXDir                    ' �`�b�v�T�C�Y�~�u���b�N(��R)��
                stPLT.dblStepOffsetXDir = typPlateInfo.dblStepOffsetXDir    ' �X�e�b�v�I�t�Z�b�g��X
                stPLT.dblStepOffsetYDir = typPlateInfo.dblStepOffsetYDir    ' �X�e�b�v�I�t�Z�b�g��Y
            Else
                If stPLT.dblChipSizeYDir <> 0.0 Then
                    stPLT.dblChipSizeYDir = typPlateInfo.dblChipSizeYDir    ' �`�b�v�T�C�YY
                End If
            End If

            Dim ResNo As Integer = 1
            For rn As Integer = 1 To stPLT.RCount
                If UserModule.IsCutResistor(rn) Then
                    For cn As Integer = 1 To stREG(rn).intTNN
                        stREG(rn).STCUT(cn).dblSTX = typResistorInfoArray(ResNo).ArrCut(cn).dblStartPointX
                        stREG(rn).STCUT(cn).dblSTY = typResistorInfoArray(ResNo).ArrCut(cn).dblStartPointY
                    Next cn
                    ResNo = ResNo + 1
                End If
            Next rn
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try

    End Sub
#End Region

#Region "�X�^�[�g�|�W�V���� �e�B�[�`���O(TEACH(F8))����"
    '''=========================================================================
    ''' <summary>�X�^�[�g�|�W�V���� �e�B�[�`���O(TEACH(F8))����</summary>
    ''' <returns>cFRS_NORMAL   = ����
    '''          ��L�ȊO      = �G���[
    ''' </returns>
    ''' <remarks>BP�I�t�Z�b�g�l�ƃg���~���O�X�^�[�g�_���e�B�[�`���O�Őݒ肷��</remarks>
    '''=========================================================================
    Public Function User_TxTyTeach() As Short

        Dim r As Short                                                  ' Return Value From Function

        Try
            '--------------------------------------------------------------------------
            '   �����ݒ菈��
            '--------------------------------------------------------------------------
            User_TxTyTeach = 0                                           ' Return�l = Normal
            Call BSIZE(stPLT.zsx, stPLT.zsy)                            ' �u���b�N�T�C�Y�ݒ�
            'Call System1.EX_BPOFF(SysPrm, BPOX, BPOY)' BP�̾�Đݒ�
            r = Move_Trimposition()                                     ' �ƕ␳(��߼��) & XYð�����шʒu�ړ�
            If (r <> cFRS_NORMAL) Then                                  ' �G���[ ?
                Return (r)                                              ' Return�l�ݒ�
            End If
            ' BP�̾�Đݒ�
            r = ObjSys.EX_BPOFF(gSysPrm, stPLT.BPOX, stPLT.BPOY)
            If (r <> cFRS_NORMAL) Then
                Return (r)                                              ' Return�l�ݒ�
            End If

            ' �p�^�[���F������
            'V2.0.0.0�N            giTemplateGroup = -1                                        ' ����ڰĸ�ٰ�ߔԍ��ݒ肷�邽�ߏ�����
            'V2.0.0.0�N            r = Ptn_Match_Exe()                                         ' �p�^�[���F�����s
            'V2.0.0.0�N            If (r <> cFRS_NORMAL) Then
            'V2.0.0.0�N                Return (r)                                              ' Return�l�ݒ�
            'V2.0.0.0�N            End If
            'V2.0.0.0�N            Call ObjSys.Ilum_Ctrl(gSysPrm, Z1, ZOPT)                     ' ���Ȱ������ߓ_��(����L���L��)


            ' �N���X���C���\���p
            r = ObjTch.SetCrossLineObject(gparModules)
            If r <> cFRS_NORMAL Then
                MsgBox("User.User_TxTyTeach() SetCrossLineObject ERROR")
            End If

            '--------------------------------------------------------------------------
            '   �s�w�A�s�x�p�v���[�g�f�[�^�ݒ�
            '--------------------------------------------------------------------------
            SetTrimDataForTXTY()

            '--------------------------------------------------------------------------
            '   �e�B�[�`���O�R���g���[���\��
            '--------------------------------------------------------------------------
            Dim TxTyObj As frmTxTyTeach = New frmTxTyTeach()

            TryCast(TxTyObj, Form).Show(Form1)                             'V6.0.0.0�J

            User_TxTyTeach = TxTyObj.Execute()                            'V6.0.0.0�L

            'TxTyObj.ShowDialog()
            r = TxTyObj.sGetReturn()                         ' Return�l = �R�}���h�I������

            '--------------------------------------------------------------------------
            '   �e�B�[�`���O���ʎ擾
            '--------------------------------------------------------------------------
            If (r = cFRS_ERR_START Or r = cFRS_TxTy) Then               ' �e�B�[�`���O��������I��
                SetTrimDataFromTXTY()
            End If

            ObjCrossLine.CrossLineOff()                                 ' �N���X���C���̔�\��

            Return (r)                                                  ' Return�l = ����

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
            Return (cERR_TRAP)                                          ' Return�l = �g���b�v�G���[
        End Try
    End Function
#End Region

#Region "�i�n�f�����ʏ����p���ʊ֐�"
    '========================================================================================
    '   �i�n�f�����ʏ����p���ʊ֐�
    '========================================================================================
#Region "�W���O����p�ϐ���`"
    '========================================================================================
    '   �W���O����p�ϐ���`(�s�w/�s�x�e�B�[�`���O������)
    '========================================================================================
    '-------------------------------------------------------------------------------
    '   �W���O����p��`
    '-------------------------------------------------------------------------------
    '    '----- JOG����p�p�����[�^�`����`(OcxJOG���g�p���Ȃ��ꍇ) -----
    Public Structure JOG_PARAM
        Dim Md As Short                                         ' �������[�h(0:XYð��وړ�, 1:BP�ړ�, 2:�L�[���͑҂����[�h)
        Dim Md2 As Short                                        ' ���̓��[�h(0:������ݓ���, 1:�ݿ�ٓ���)
        Dim Opt As UShort                                       ' �I�v�V����(�L�[�̗L��(1)/����(0)�w��)
        '                                                       '  BIT0:START�L�[
        '                                                       '  BIT1:RESET�L�[
        '                                                       '  BIT2:Z�L�[
        '                                                       '  BIT3:
        '                                                       '  BIT4:���g�p
        '                                                       '  BIT5:HALT�L�[
        '                                                       '  BIT6:���g�p
        '                                                       '  BIT7-15:���g�p
        Dim Flg As Short                                        ' �e��ʂ�OK/Cancel���݉����׸�(cFRS_ERR_ADV, cFRS_ERR_RST)
        Dim PosX As Double                                      ' BP or ð��� X�ʒu
        Dim PosY As Double                                      ' BP or ð��� Y�ʒu
        Dim BpOffX As Double                                    ' BP�̾��X 
        Dim BpOffY As Double                                    ' BP�̾��Y
        Dim BszX As Double                                      ' ��ۯ�����X 
        Dim BszY As Double                                      ' ��ۯ�����Y
        Dim TextX As Object                                     ' BP or ð��� X�ʒu�\���p÷���ޯ��
        Dim TextY As Object                                     ' BP or ð��� Y�ʒu�\���p÷���ޯ��
        Dim cgX As Double                                       ' �ړ���X 
        Dim cgY As Double                                       ' �ړ���Y
        Dim bZ As Boolean                                       ' Z�L�[  (True:ON, False:OFF)

        Dim BtnHI As Object                                     ' HI�{�^��
        Dim BtnZ As Object                                      ' Z�{�^��
        Dim BtnSTART As Object                                  ' START�{�^��
        Dim BtnHALT As Object                                   ' HALT�{�^��
        Dim BtnRESET As Object                                  ' RESET�{�^��
        Dim CurrentNo As Integer                                ' �������̍s�ԍ�(�O���b�h�\���p)

        Dim TenKey() As Button                                  ' V2.2.0.0�@
        Dim KeyDown As Keys                                     ' V2.2.0.0�@

    End Structure

    '    '----- ZINPSTS�֐�(�R���\�[������)�ߒl -----
    Public Const CONSOLE_SW_START As UShort = &H1           ' bit 0(01)  : START       0/1=������/����
    Public Const CONSOLE_SW_RESET As UShort = &H2           ' bit 1(02)  : RESET       0/1=������/����
    Public Const CONSOLE_SW_ZSW As UShort = &H4             ' bit 2(04)  : Z_ON/OFF_SW 0/1=������/����
    '    Public Const CONSOLE_SW_ZDOWN As UShort = &H8           ' bit 3(08)  : Z_DOWN      1=��ԃZ���X
    '    Public Const CONSOLE_SW_ZUP As UShort = &H10            ' bit 4(10)  : Z_UP        1=��ԃZ���X
    Public Const CONSOLE_SW_HALT As UShort = &H20           ' bit 5(20)  : HALT        0/1=������/����

    '    '----- �R���\�[���L�[SW -----
    '    'Public Const cBIT_ADV As UShort = &H1US                 ' START(ADV)�L�[
    '    'Public Const cBIT_HALT As UShort = &H2US                ' HALT�L�[
    '    'Public Const cBIT_RESET As UShort = &H8US               ' RESET�L�[
    '    'Public Const cBIT_Z As UShort = &H20US                  ' Z�L�[
    Public Const cBIT_HI As UShort = &H100US                ' HI�L�[

    '    '----- �������[�h��` -----
    Public Const MODE_STG As Integer = 0                    ' XY�e�[�u�����[�h
    Public Const MODE_BP As Integer = 1                     ' BP���[�h
    Public Const MODE_KEY As Integer = 2                    ' �L�[���͑҂����[�h

    '    '----- ���̓��[�h -----
    Public Const MD2_BUTN As Integer = 0                    ' ��ʃ{�^������

    '    '----- �s�b�`�ő�l/�ŏ��l -----
    Public Const cPT_LO As Double = 0.001                   ' �߯��ŏ��l(mm)
    Public Const cPT_HI As Double = 0.1                     ' �߯��ő�l(mm)
    Public Const cHPT_LO As Double = 0.01                   ' HIGH�߯��ŏ��l(mm)
    Public Const cHPT_HI As Double = 5.0#                   ' HIGH�߯��ő�l(mm)
    Public Const cPAU_LO As Double = 0.05                   ' �|�[�Y�ŏ��l(sec)
    Public Const cPAU_HI As Double = 1.0#                   ' �|�[�Y�ő�l(sec)

    '    '----- �Y���� -----
    Public Const IDX_PIT As Short = 0                       ' �߯�
    Public Const IDX_HPT As Short = 1                       ' HIGH�߯�
    Public Const IDX_PAU As Short = 2                       ' �|�[�Y

    '    '----- ���̑� -----
    '    'Private dblTchMoval(3) As Double                           ' �߯��ޔ���(0=�߯�, 1=HIGH�߯�, 2=Pause Time(Sec))
    Private InpKey As UShort                                    ' �ݿ�ٷ����͈� 
    Private cin As UShort                                       ' �ݿ�ٓ��͒l
    Private bZ As Boolean                                       ' Z�L�[ �ޔ��� (True:ON, False:OFF)
    Private bHI As Boolean                                      ' HI�L�[(True:ON, False:OFF)

    Private mPIT As Double                                      ' �ړ��߯�
    Private X As Double                                         ' �ړ��߯�(X)
    Private Y As Double                                         ' �ړ��߯�(Y)
    '    Private NOWXP As Double                                     ' BP���ݒlX(�۽ײݕ␳�p)
    '    Private NOWYP As Double                                     ' BP���ݒlY(�۽ײݕ␳�p)
    Private mvx As Double                                       ' BP/ð��ٓ��̈ʒuX
    Private mvy As Double                                       ' BP/ð��ٓ��̈ʒuY
    Private mvxBk As Double                                     ' BP/ð��ٓ��̈ʒuX(�ޔ�p)
    Private mvyBk As Double                                     ' BP/ð��ٓ��̈ʒuY(�ޔ�p)
#End Region

#Region "�����ݒ菈��"
    '''=========================================================================
    '''<summary>�����ݒ菈��</summary>
    '''<param name="stJOG">       (INP)JOG����p�p�����[�^</param>
    '''<param name="TBarLowPitch">(I/O)�X���C�_�[1(Low�߯�)</param>
    '''<param name="TBarHiPitch"> (I/O)�X���C�_�[2(HIGH�߯�)</param>
    '''<param name="TBarPause">   (I/O)�X���C�_�[3(Pause Time)</param>
    '''<param name="LblTchMoval0">(I/O)�ڐ�1(Low Pich Label)</param>
    '''<param name="LblTchMoval1">(I/O)�ڐ�2(Low�߯� Label)</param>
    '''<param name="LblTchMoval2">(I/O)�ڐ�3(HIGH�߯� Label)</param>
    '''<param name="dblTchMoval"> (I/O)�߯��ޔ���(0=�߯�, 1=HIGH�߯�, 2=Pause Time)</param>
    '''=========================================================================
    Public Sub JogEzInit(ByVal stJOG As JOG_PARAM, _
                         ByRef TBarLowPitch As System.Windows.Forms.TrackBar, _
                         ByRef TBarHiPitch As System.Windows.Forms.TrackBar, _
                         ByRef TBarPause As System.Windows.Forms.TrackBar, _
                         ByRef LblTchMoval0 As System.Windows.Forms.Label, _
                         ByRef LblTchMoval1 As System.Windows.Forms.Label, _
                         ByRef LblTchMoval2 As System.Windows.Forms.Label, _
                         ByRef dblTchMoval() As Double)

        Try
            ' �ړ��s�b�`�X���C�_�[�����ݒ�
            If (stJOG.Md = MODE_BP) Then                            ' ���[�h = 1(BP�ړ�) ?
                dblTchMoval(IDX_PIT) = gSysPrm.stSYP.gBpPIT         ' BP�p�߯��ݒ�
                dblTchMoval(IDX_HPT) = gSysPrm.stSYP.gBpHighPIT
                dblTchMoval(IDX_PAU) = gSysPrm.stSYP.gPitPause
                'V2.2.1.1�B��
                If gSysPrm.stDEV.giBpSize = 40 Then
                    dblTchMoval(3) = 1
                End If
                'V2.2.1.1�B��
            Else
                dblTchMoval(IDX_PIT) = gSysPrm.stSYP.gPIT           ' XY�e�[�u���p�߯��ݒ�
                dblTchMoval(IDX_HPT) = gSysPrm.stSYP.gStageHighPIT
                dblTchMoval(IDX_PAU) = gSysPrm.stSYP.gPitPause
            End If
            Call XyzBpMovingPitchInit(TBarLowPitch, TBarHiPitch, TBarPause, _
                                      LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval)

            Call Form1.System1.SetSysParam(gSysPrm)                 ' �V�X�e���p�����[�^�̐ݒ�(OcxSystem�p)

            InpKey = 0

            ' �g���b�v�G���[������
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub
#End Region

#Region "BP/XY�e�[�u����JOG����(Do Loop�Ȃ�)"
    '''=========================================================================
    '''<summary>BP/XY�e�[�u����JOG����</summary>
    '''<param name="stJOG">       (INP)JOG����p�p�����[�^</param>
    '''<param name="TBarLowPitch">(I/O)�X���C�_�[1(Low�߯�)</param>
    '''<param name="TBarHiPitch"> (I/O)�X���C�_�[2(HIGH�߯�)</param>
    '''<param name="TBarPause">   (I/O)�X���C�_�[3(Pause Time)</param>
    '''<param name="LblTchMoval0">(I/O)�ڐ�1(Low Pich Label)</param>
    '''<param name="LblTchMoval1">(I/O)�ڐ�2(Low�߯� Label)</param>
    '''<param name="LblTchMoval2">(I/O)�ڐ�3(HIGH�߯� Label)</param>
    '''<param name="dblTchMoval"> (I/O)�߯��ޔ���(0=�߯�, 1=HIGH�߯�, 2=Pause Time)</param>
    '''<returns>cFRS_ERR_ADV = OK(START��) 
    '''         cFRS_ERR_RST = Cancel(RESET��)
    '''         cFRS_ERR_HLT = HALT��
    '''         -1�ȉ�       = �G���[</returns>
    ''' <remarks>JogEzInit�֐���Call�ςł��邱��</remarks>
    '''=========================================================================
    Public Function JogEzMove_Ex(ByRef stJOG As JOG_PARAM, ByVal SysPrm As SYSPARAM_PARAM, _
                         ByRef TBarLowPitch As System.Windows.Forms.TrackBar, _
                         ByRef TBarHiPitch As System.Windows.Forms.TrackBar, _
                         ByRef TBarPause As System.Windows.Forms.TrackBar, _
                         ByRef LblTchMoval0 As System.Windows.Forms.Label, _
                         ByRef LblTchMoval1 As System.Windows.Forms.Label, _
                         ByRef LblTchMoval2 As System.Windows.Forms.Label, _
                         ByRef dblTchMoval() As Double) As Integer

        Dim r As Short

        Try
            '---------------------------------------------------------------------------
            '   ��������
            '---------------------------------------------------------------------------
            X = 0.0 : Y = 0.0                                           ' �ړ��߯�X,Y
            mvx = stJOG.PosX : mvy = stJOG.PosY                         ' BP or ð��وʒuX,Y
            mvxBk = stJOG.PosX : mvyBk = stJOG.PosY
            ' �L�����u���[�V�������s/�J�b�g�ʒu�␳�y�O���J�����z�� �����΍��W��\�����邽�߃N���A���Ȃ�
            ' �g���~���O���̈ꎞ��~��ʂ��N���A���Ȃ�
            If (giAppMode <> APP_MODE_CARIB_REC) And (giAppMode <> APP_MODE_CUTREVIDE) And _
               (giAppMode <> APP_MODE_FINEADJ) Then
                '(giAppMode <> APP_MODE_TRIM) Then                      
                stJOG.cgX = 0.0# : stJOG.cgY = 0.0#                     ' �ړ���X,Y
            End If

            'If (giAppMode = APP_MODE_TRIM) Then                        
            If (giAppMode = APP_MODE_FINEADJ) Then
                mvx = stJOG.cgX - stJOG.BpOffX : mvy = stJOG.cgY - stJOG.BpOffY
                mvxBk = mvx : mvyBk = mvy
            End If

            Call Init_Proc(stJOG, TBarLowPitch, TBarHiPitch, TBarPause, LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval)

            ' ���݂̈ʒu��\������(÷���ޯ���̔w�i�F��������(���F)�ɐݒ肷��)
            Call DispPosition(stJOG, 1)
            'Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
            'Call Me.Focus()                                            ' �t�H�[�J�X��ݒ肷��(�e���L�[���͂̂���)
            ''                                                          ' KeyPreview�v���p�e�B��True�ɂ���ƑS�ẴL�[�C�x���g���܂��t�H�[�����󂯎��悤�ɂȂ�B
            '---------------------------------------------------------------------------
            '   �R���\�[���{�^�����̓R���\�[���L�[����̃L�[���͏������s��
            '---------------------------------------------------------------------------
            ' �V�X�e���G���[�`�F�b�N
            r = Form1.System1.SysErrChk_ForVBNET(giAppMode)
            If (r <> cFRS_NORMAL) Then Return (r)

            ' ���b�Z�[�W�|���v
            System.Windows.Forms.Application.DoEvents()

            ' �R���\�[���{�^�����̓R���\�[���L�[����̃L�[����
            Call ReadConsoleSw(stJOG, cin)                              ' �L�[����

            '-----------------------------------------------------------------------
            '   ���̓L�[�`�F�b�N
            '-----------------------------------------------------------------------
            If (cin And CONSOLE_SW_RESET) Then                          ' RESET SW ?
                ' RESET SW������
                If (stJOG.Opt And CONSOLE_SW_RESET) Then                ' RESET�L�[�L�� ?
                    Return (cFRS_ERR_RST)                               ' Return�l = Cancel(RESET��)
                End If

                ' HALT SW������
            ElseIf (cin And CONSOLE_SW_HALT) Then                       ' HALT SW ?
                If (stJOG.Opt And CONSOLE_SW_HALT) Then                 ' �I�v�V����(0:HALT�L�[����, 1:HALT�L�[�L��)
                    r = cFRS_ERR_HALT                                   ' Return�l = HALT��
                    GoTo STP_END
                End If

                ' START SW������
            ElseIf (cin And CONSOLE_SW_START) Then                      ' START SW ?
                If (stJOG.Opt And CONSOLE_SW_START) Then                ' START�L�[�L�� ?
                    r = cFRS_ERR_START                                  ' Return�l = OK(START��) 
                    GoTo STP_END
                End If

                ' Z SW��ON����OFF(����OFF����ON)�ɐؑւ������
            ElseIf (stJOG.bZ <> bZ) Then
                If (stJOG.Opt And CONSOLE_SW_ZSW) Then                  ' Z�L�[�L�� ?
                    r = cFRS_ERR_Z                                      ' Return�l = Z��ON/OFF
                    stJOG.bZ = bZ                                       ' ON/OFF
                    GoTo STP_END
                End If

                ' ���SW������
            ElseIf cin And &H1E00US Then                                ' ���SW
                '�u�L�[���͑҂����[�h�v�Ȃ牽�����Ȃ�
                If (stJOG.Md = MODE_KEY) Then

                Else
                    If cin And &H100US Then                             ' HI SW ? 
                        mPIT = dblTchMoval(IDX_HPT)                     ' mPIT = �ړ������߯�
                    Else
                        mPIT = dblTchMoval(IDX_PIT)                     ' mPIT = �ړ��ʏ��߯�
                    End If

                    ' XY�e�[�u����Βl�ړ�(�\�t�g���~�b�g�`�F�b�N�L��)
                    r = cFRS_NORMAL
                    If (stJOG.Md = MODE_STG) Then                       ' ���[�h = XY�e�[�u���ړ� ?
                        ' XY�e�[�u����Βl�ړ�
                        r = Sub_XYtableMove(SysPrm, Form1.System1, Form1.Utility1, stJOG)
                        If (r <> cFRS_NORMAL) Then                      ' �װ ?
                            If (Form1.System1.IsSoftLimitXY(r) = False) Then
                                Return (r)                              ' ����ЯĴװ�ȊO�ʹװ����
                            End If
                        End If

                        '  ���[�h = BP�ړ��̏ꍇ
                    ElseIf (stJOG.Md = MODE_BP) Then
                        ' BP��Βl�ړ�
                        r = Sub_BPmove(SysPrm, Form1.System1, Form1.Utility1, stJOG)
                        If (r <> cFRS_NORMAL) Then                      ' BP�ړ��G���[ ?
                            If (Form1.System1.IsSoftLimitBP(r) = False) Then
                                Return (r)                              ' ����ЯĴװ�ȊO�ʹװ����
                            End If
                        End If
                    End If

                    ' �\�t�g���~�b�g�G���[�̏ꍇ�� HI SW�ȊO��OFF����
                    If (r <> cFRS_NORMAL) Then                          ' �װ ?
                        If (stJOG.BtnHI.BackColor = System.Drawing.Color.Yellow) Then
                            InpKey = cBIT_HI                            ' HI SW ON
                        Else
                            InpKey = 0                                  ' HI SW�ȊO��OFF
                        End If
                        r = cFRS_NORMAL                                 ' Retuen�l = ����
                    End If

                    ' ���݂̈ʒu��\������
                    Call DispPosition(stJOG, 1)
                    Call Form1.System1.WAIT(dblTchMoval(IDX_PAU))       ' Wait(sec)

                    InpKey = CType(CtrlJog.MouseClickLocation.Clear(InpKey), UShort)    'V2.2.0.0�@ 
                    stJOG.KeyDown = Keys.None                                           'V2.2.0.0�@ 

                End If

            End If

            '---------------------------------------------------------------------------
            '   �I������
            '---------------------------------------------------------------------------
STP_END:
            'stJOG.PosX = mvx                                            ' �ʒuX,Y�X�V
            'stJOG.PosY = mvy
            Return (r)                                                  ' Return�l�ݒ� 

            ' �g���b�v�G���[������
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
            Return (cERR_TRAP)                                          ' Return�l = ��O�G���[ 
        End Try
    End Function
#End Region

#Region "BP/XY�e�[�u����JOG����"
    '''=========================================================================
    '''<summary>BP/XY�e�[�u����JOG����</summary>
    '''<param name="stJOG">       (INP)JOG����p�p�����[�^</param>
    '''<param name="TBarLowPitch">(I/O)�X���C�_�[1(Low�߯�)</param>
    '''<param name="TBarHiPitch"> (I/O)�X���C�_�[2(HIGH�߯�)</param>
    '''<param name="TBarPause">   (I/O)�X���C�_�[3(Pause Time)</param>
    '''<param name="LblTchMoval0">(I/O)�ڐ�1(Low Pich Label)</param>
    '''<param name="LblTchMoval1">(I/O)�ڐ�2(Low�߯� Label)</param>
    '''<param name="LblTchMoval2">(I/O)�ڐ�3(HIGH�߯� Label)</param>
    '''<param name="dblTchMoval"> (I/O)�߯��ޔ���(0=�߯�, 1=HIGH�߯�, 2=Pause Time)</param>
    '''<returns>cFRS_ERR_ADV = OK(START��) 
    '''         cFRS_ERR_RST = Cancel(RESET��)
    '''         cFRS_ERR_HLT = HALT��
    '''         -1�ȉ�       = �G���[</returns>
    ''' <remarks>JogEzInit�֐���Call�ςł��邱��</remarks>
    '''=========================================================================
    Public Function JogEzMove(ByRef stJOG As JOG_PARAM, ByVal SysPrm As SYSPARAM_PARAM,
                         ByRef TBarLowPitch As System.Windows.Forms.TrackBar,
                         ByRef TBarHiPitch As System.Windows.Forms.TrackBar,
                         ByRef TBarPause As System.Windows.Forms.TrackBar,
                         ByRef LblTchMoval0 As System.Windows.Forms.Label,
                         ByRef LblTchMoval1 As System.Windows.Forms.Label,
                         ByRef LblTchMoval2 As System.Windows.Forms.Label,
                         ByRef dblTchMoval() As Double,
                         ByVal commonMethods As ICommonMethods) As Integer      ''V2.2.0.0�@   ���� ICommonMethods �ǉ�

        Dim r As Short

        Try
            '---------------------------------------------------------------------------
            '   ��������
            '---------------------------------------------------------------------------
            X = 0.0 : Y = 0.0                                   ' �ړ��߯�X,Y
            mvx = stJOG.PosX : mvy = stJOG.PosY                 ' BP or ð��وʒuX,Y
            mvxBk = stJOG.PosX : mvyBk = stJOG.PosY
            ' �L�����u���[�V�������s/�J�b�g�ʒu�␳�y�O���J�����z�� �����΍��W��\�����邽�߃N���A���Ȃ�
            ' �g���~���O���̈ꎞ��~��ʂ��N���A���Ȃ�
            If (giAppMode <> APP_MODE_CARIB_REC) And (giAppMode <> APP_MODE_CUTREVIDE) And
               (giAppMode <> APP_MODE_FINEADJ) Then
                stJOG.cgX = 0.0# : stJOG.cgY = 0.0#             ' �ړ���X,Y
            End If
            stJOG.Flg = -1
            InpKey = 0
            Call Init_Proc(stJOG, TBarLowPitch, TBarHiPitch, TBarPause, LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval)

            ' ���݂̈ʒu��\������(÷���ޯ���̔w�i�F��������(���F)�ɐݒ肷��)
            Call DispPosition(stJOG, 1)
            'Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
            'Call Me.Focus()                                     ' �t�H�[�J�X��ݒ肷��(�e���L�[���͂̂���)
            ''                                                   ' KeyPreview�v���p�e�B��True�ɂ���ƑS�ẴL�[�C�x���g���܂��t�H�[�����󂯎��悤�ɂȂ�B

            ' ���C���t�H�[����JOG����֐���ݒ肷��      'V2.2.0.0�@ 
            Form1.SetActiveJogMethod(AddressOf commonMethods.JogKeyDown,
                                              AddressOf commonMethods.JogKeyUp,
                                              AddressOf commonMethods.MoveToCenter)

            '---------------------------------------------------------------------------
            '   �R���\�[���{�^�����̓R���\�[���L�[����̃L�[���͏������s��
            '---------------------------------------------------------------------------
            Do
                ' �V�X�e���G���[�`�F�b�N
                r = Form1.System1.SysErrChk_ForVBNET(giAppMode)
                If (r <> cFRS_NORMAL) Then GoTo STP_END

                ' ���b�Z�[�W�|���v
                '  ��VB.NET�̓}���`�X���b�h�Ή��Ȃ̂ŁA�{���̓C�x���g�̊J���ȂǂłȂ��A
                '    �X���b�h�𐶐����ăR�[�f�B���O������̂��������B
                '    �X���b�h�łȂ��Ă��A�Œ�Ń^�C�}�[�𗘗p����B
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(10)               ' CPU�g�p���������邽�߃X���[�v


                ' �R���\�[���{�^�����̓R���\�[���L�[����̃L�[����
                Call ReadConsoleSw(stJOG, cin)                  ' �L�[����

                '-----------------------------------------------------------------------
                '   ���̓L�[�`�F�b�N
                '-----------------------------------------------------------------------
                If (cin And CONSOLE_SW_RESET) Then              ' RESET SW ?
                    ' RESET SW������
                    If (stJOG.Opt And CONSOLE_SW_RESET) Then    ' RESET�L�[�L�� ?
                        r = cFRS_ERR_RST                        ' Return�l = Cancel(RESET��)
                        Exit Do
                    End If

                    ' HALT SW������
                ElseIf (cin And CONSOLE_SW_HALT) Then           ' HALT SW ?
                    If (stJOG.Opt And CONSOLE_SW_HALT) Then     ' �I�v�V����(0:HALT�L�[����, 1:HALT�L�[�L��)
                        r = cFRS_ERR_HALT                       ' Return�l = HALT��
                        Exit Do
                    End If

                    ' START SW������
                ElseIf (cin And CONSOLE_SW_START) Then          ' START SW ?
                    If (stJOG.Opt And CONSOLE_SW_START) Then    ' START�L�[�L�� ?
                        'stJOG.PosX = mvx                       ' �ʒuX,Y�X�V
                        'stJOG.PosY = mvy
                        r = cFRS_ERR_START                      ' Return�l = OK(START��) 
                        Exit Do
                    End If

                    ' Z SW��ON����OFF(����OFF����ON)�ɐؑւ������
                ElseIf (stJOG.bZ <> bZ) Then
                    If (stJOG.Opt And CONSOLE_SW_ZSW) Then      ' Z�L�[�L�� ?
                        r = cFRS_ERR_Z                          ' Return�l = Z��ON/OFF
                        stJOG.bZ = bZ                           ' ON/OFF
                        Exit Do
                    End If

                    ' ���SW������
                ElseIf cin And &H1E00US Then                    ' ���SW
                    '�u�L�[���͑҂����[�h�v�Ȃ牽�����Ȃ�
                    If (stJOG.Md = MODE_KEY) Then

                    Else
                        If cin And &H100US Then                     ' HI SW ? 
                            mPIT = dblTchMoval(IDX_HPT)             ' mPIT = �ړ������߯�
                        Else
                            mPIT = dblTchMoval(IDX_PIT)             ' mPIT = �ړ��ʏ��߯�
                        End If

                        ' XY�e�[�u����Βl�ړ�(�\�t�g���~�b�g�`�F�b�N�L��)
                        r = cFRS_NORMAL
                        If (stJOG.Md = MODE_STG) Then                ' ���[�h = XY�e�[�u���ړ� ?
                            ' XY�e�[�u����Βl�ړ�
                            r = Sub_XYtableMove(SysPrm, Form1.System1, Form1.Utility1, stJOG)
                            If (r <> cFRS_NORMAL) Then              ' �װ ?
                                If (Form1.System1.IsSoftLimitXY(r) = False) Then
                                    GoTo STP_END                    ' ����ЯĴװ�ȊO�ʹװ����
                                End If
                            End If

                            '  ���[�h = BP�ړ��̏ꍇ
                        ElseIf (stJOG.Md = MODE_BP) Then
                            ' BP��Βl�ړ�
                            r = Sub_BPmove(SysPrm, Form1.System1, Form1.Utility1, stJOG)
                            If (r <> cFRS_NORMAL) Then              ' BP�ړ��G���[ ?
                                If (Form1.System1.IsSoftLimitBP(r) = False) Then
                                    GoTo STP_END                    ' ����ЯĴװ�ȊO�ʹװ����
                                End If
                            End If
                        End If

                        ' �\�t�g���~�b�g�G���[�̏ꍇ�� HI SW�ȊO��OFF����
                        If (r <> cFRS_NORMAL) Then                  ' �װ ?
                            If (stJOG.BtnHI.BackColor = System.Drawing.Color.Yellow) Then
                                InpKey = cBIT_HI                    ' HI SW ON
                            Else
                                InpKey = 0                          ' HI SW�ȊO��OFF
                            End If
                        End If

                        ' ���݂̈ʒu��\������
                        Call DispPosition(stJOG, 1)
                        Call Form1.System1.WAIT(SysPrm.stSYP.gPitPause)    ' Wait(sec)
                    End If
                    InpKey = CType(CtrlJog.MouseClickLocation.Clear(InpKey), UShort)    'V2.2.0.0�@ 
                    stJOG.KeyDown = Keys.None                                           'V2.2.0.0�@ 

                End If

            Loop While (stJOG.Flg = -1)

            '---------------------------------------------------------------------------
            '   �I������
            '---------------------------------------------------------------------------
            ' ���W�\���p÷���ޯ���̔w�i�F�𔒐F�ɐݒ肷��
            Call DispPosition(stJOG, 0)

            ' �e��ʂ���OK/Cancel���݉��� ?
            If (stJOG.Flg <> -1) Then
                r = stJOG.Flg
            End If

            ' OK(START��)�Ȃ�ʒuX,Y�X�V
            If (r = cFRS_ERR_START) Then                            ' OK(START��) ?
                stJOG.PosX = mvx                                    ' �ʒuX,Y�X�V
                stJOG.PosY = mvy
            End If

STP_END:
            Call ZCONRST()                                          ' �ݿ�ٷ�ׯ����� 
            Return (r)                                              ' Return�l�ݒ� 

            ' �g���b�v�G���[������
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
            Return (cERR_TRAP)                                      ' Return�l = ��O�G���[ 

        Finally
            Form1.SetActiveJogMethod(Nothing, Nothing, Nothing)    'V6.0.0.0�J

        End Try
    End Function
#End Region

#Region "�����ݒ菈��"
    '''=========================================================================
    '''<summary>�����ݒ菈��</summary>
    '''<param name="stJOG">       (INP)JOG����p�p�����[�^</param>
    '''<param name="TBarLowPitch">(I/O)�X���C�_�[1(Low�߯�)</param>
    '''<param name="TBarHiPitch"> (I/O)�X���C�_�[2(HIGH�߯�)</param>
    '''<param name="TBarPause">   (I/O)�X���C�_�[3(Pause Time)</param>
    '''<param name="LblTchMoval0">(I/O)�ڐ�1(Low Pich Label)</param>
    '''<param name="LblTchMoval1">(I/O)�ڐ�2(Low�߯� Label)</param>
    '''<param name="LblTchMoval2">(I/O)�ڐ�3(HIGH�߯� Label)</param>
    '''<param name="dblTchMoval"> (I/O)�߯��ޔ���(0=�߯�, 1=HIGH�߯�, 2=Pause Time)</param>
    '''=========================================================================
    Private Sub Init_Proc(ByVal stJOG As JOG_PARAM, _
                         ByRef TBarLowPitch As System.Windows.Forms.TrackBar, _
                         ByRef TBarHiPitch As System.Windows.Forms.TrackBar, _
                         ByRef TBarPause As System.Windows.Forms.TrackBar, _
                         ByRef LblTchMoval0 As System.Windows.Forms.Label, _
                         ByRef LblTchMoval1 As System.Windows.Forms.Label, _
                         ByRef LblTchMoval2 As System.Windows.Forms.Label, _
                         ByRef dblTchMoval() As Double)

        Try

            ' �ړ��s�b�`�X���C�_�[�ݒ�(�O��ݒ肵���l)
            Call XyzBpMovingPitchInit(TBarLowPitch, TBarHiPitch, TBarPause, _
                                      LblTchMoval0, LblTchMoval1, LblTchMoval2, dblTchMoval)

            ' �{�^���L��/�����ݒ�
            If (stJOG.Opt And CONSOLE_SW_HALT) Then                 ' HALT�L�[�L��/����
                stJOG.BtnHALT.Enabled = True
            Else
                stJOG.BtnHALT.Enabled = False
            End If
            If (stJOG.Opt And CONSOLE_SW_START) Then                ' START�L�[�L��/����
                stJOG.BtnSTART.Enabled = True
            Else
                stJOG.BtnSTART.Enabled = False
            End If
            If (stJOG.Opt And CONSOLE_SW_RESET) Then                ' RESET�L�[�L��/����
                stJOG.BtnRESET.Enabled = True
            Else
                stJOG.BtnRESET.Enabled = False
            End If
            If (stJOG.Opt And CONSOLE_SW_ZSW) Then                  ' Z�L�[�L��/����
                stJOG.BtnZ.Enabled = True
            Else
                stJOG.BtnZ.Enabled = False
            End If

            ' Z�L�[/HI�L�[��ԓ��ޔ�
            bZ = stJOG.bZ                                           ' Z�L�[�ޔ�
            If (bZ = False) Then                                    ' Z�{�^���̔w�i�F��ݒ�
                stJOG.BtnZ.BackColor = System.Drawing.SystemColors.Control ' �w�i�F = �D�F
                stJOG.BtnZ.Text = "Z Off"
            Else
                stJOG.BtnZ.BackColor = System.Drawing.Color.Yellow        ' �w�i�F = ���F
                stJOG.BtnZ.Text = "Z On"
            End If

            If (stJOG.BtnHI.BackColor = System.Drawing.Color.Yellow) Then ' HI�L�[��Ԏ擾
                bHI = True
                InpKey = InpKey Or cBIT_HI                          ' HI SW ON
            Else
                bHI = False
                InpKey = InpKey And Not cBIT_HI                     ' HI SW OFF
            End If

            ' �g���b�v�G���[������
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub
#End Region

#Region "��ʃ{�^�����̓R���\�[���L�[����̃L�[����"
    '''=========================================================================
    '''<summary>��ʃ{�^�����̓R���\�[���L�[����̃L�[����</summary>
    '''<param name="stJOG">(INP)JOG����p�p�����[�^</param>
    '''<param name="cin">  (OUT)�R���\�[�����͒l</param>
    '''=========================================================================
    Private Sub ReadConsoleSw(ByRef stJOG As JOG_PARAM, ByRef cin As UShort)

        Dim r As Integer
        Dim sw As Long
        Dim strMSG As String

        Try
            ' HALT�L�[���̓`�F�b�N
            r = HALT_SWCHECK(sw)
            If (sw <> 0) Then                                           ' HALT�L�[���� ?
                If (stJOG.Opt And CONSOLE_SW_HALT) Then                 ' HALT�L�[�L�� ?
                    cin = CONSOLE_SW_HALT
                    Exit Sub
                End If
            End If

            ' Z�L�[���̓`�F�b�N
            r = Z_SWCHECK(sw)                                           ' Z�X�C�b�`�̏�Ԃ��`�F�b�N����
            If (sw <> 0) Then                                           ' Z�L�[���� ?
                If (stJOG.Opt And CONSOLE_SW_ZSW) Then                  ' Z�L�[�L�� ?
                    Call SubBtnZ_Click(stJOG)
                    Exit Sub
                End If
            End If

            ' START/RESET�L�[���̓`�F�b�N
            r = STARTRESET_SWCHECK(False, sw)                           ' START/RESET�L�[�����`�F�b�N(�Ď��Ȃ����[�h)

            ' �R���\�[�����͒l�ɕϊ����Đݒ�
            If (sw = cFRS_ERR_START) Then                               ' START�L�[���� ?
                If (stJOG.Opt And CONSOLE_SW_START) Then                ' START�L�[�L�� ?
                    cin = CONSOLE_SW_START
                    Exit Sub
                End If
            ElseIf (sw = cFRS_ERR_RST) Then                             ' RESET�L�[���� ?
                If (stJOG.Opt And CONSOLE_SW_RESET) Then                ' RESET�L�[�L�� ?
                    cin = CONSOLE_SW_RESET
                    Exit Sub
                End If
                '    ElseIf (sw = CONSOLE_SW_ZSW) Then                          ' Z�L�[���� ?
                '        If (stJOG.opt And CONSOLE_SW_ZSW) Then                  ' Z�L�[�L�� ?
                '            cin = CONSOLE_SW_ZSW
                '        End If
            End If

            ' �u��ʃ{�^�����́v
            cin = InpKey                                                ' ��ʃ{�^������

            ' �g���b�v�G���[������
        Catch ex As Exception
            strMSG = "Globals.ReadConsoleSw() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "���W�\��"
    '''=========================================================================
    '''<summary>���W�\��</summary>
    '''<param name="stJOG">(INP)JOG����p�p�����[�^</param>
    '''<param name="Md">   (INP)0=�w�i�F�𔒐F�ɐݒ�, 1=�w�i�F��������(���F)�ɐݒ�</param>
    '''=========================================================================
    Private Sub DispPosition(ByVal stJOG As JOG_PARAM, ByVal MD As Integer)

        Dim xPos As Double = 0.0
        Dim yPos As Double = 0.0
        Dim strMSG As String

        Try
            '�u�L�[���͑҂����[�h�v�Ȃ�NOP
            If (stJOG.Md = MODE_KEY) Then Exit Sub

            ' �␳�ʒu�e�B�[�`���O�Ȃ�O���b�h�ɕ\������
            If (giAppMode = APP_MODE_CUTPOS) Then
                stJOG.TextX.set_TextMatrix(stJOG.CurrentNo, 2, (stJOG.PosX + stJOG.cgX).ToString("0.0000"))
                stJOG.TextX.set_TextMatrix(stJOG.CurrentNo, 3, (stJOG.PosY + stJOG.cgY).ToString("0.0000"))
                Exit Sub

                ' �J�b�g�ʒu�␳�y�O���J�����z�Ȃ�O���b�h�ɑ��΍��W��\������
            ElseIf (giAppMode = APP_MODE_CUTREVIDE) Then
                stJOG.TextX.set_TextMatrix(stJOG.CurrentNo, 3, (stJOG.cgX).ToString("0.0000"))    ' �����X
                stJOG.TextX.set_TextMatrix(stJOG.CurrentNo, 4, (stJOG.cgY).ToString("0.0000"))    ' �����Y
                Exit Sub
            End If

            ' �e�L�X�g�{�b�N�X�ɍ��W��\������
            If (MD = 0) Then
                ' �L�����u���[�V�������s���͔w�i�F���D�F�ɐݒ�
                If (giAppMode = APP_MODE_CARIB_REC) Then
                    ' �w�i�F���D�F�ɐݒ�
                    stJOG.TextX.BackColor = System.Drawing.SystemColors.Control
                    stJOG.TextY.BackColor = System.Drawing.SystemColors.Control
                Else
                    ' �w�i�F�𔒐F�ɐݒ�
                    stJOG.TextX.BackColor = System.Drawing.Color.White
                    stJOG.TextY.BackColor = System.Drawing.Color.White
                End If
            Else
                ' �L�����u���[�V�������s���͑��΍��W��\��
                If (giAppMode = APP_MODE_CARIB_REC) Then
                    stJOG.TextX.Text = stJOG.cgX.ToString("0.0000")
                    stJOG.TextY.Text = stJOG.cgY.ToString("0.0000")
                Else
                    ' ���̑��̃��[�h���͐�΍��W��\��
                    stJOG.TextX.Text = (stJOG.PosX + stJOG.cgX).ToString("0.0000")
                    stJOG.TextY.Text = (stJOG.PosY + stJOG.cgY).ToString("0.0000")
                    ' �g���~���O���̈ꎞ��~��ʕ\�����Ȃ�␳�N���X���C����\������
                    If (giAppMode = APP_MODE_FINEADJ) Or (giAppMode = APP_MODE_TX) Then
                        'xPos = Double.Parse(stJOG.TextX.Text)
                        'yPos = Double.Parse(stJOG.TextY.Text)
                        Call ZGETBPPOS(xPos, yPos)
                        ObjCrossLine.CrossLineDispXY(xPos, yPos)
                    End If
                End If
                ' �w�i�F��������(���F)�ɐݒ�
                stJOG.TextX.BackColor = System.Drawing.Color.Yellow
                stJOG.TextY.BackColor = System.Drawing.Color.Yellow
            End If

            stJOG.TextX.Refresh()
            stJOG.TextY.Refresh()

            ' �g���b�v�G���[������
        Catch ex As Exception
            strMSG = "Globals.DispPosition() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "BP��Βl�ړ�(�\�t�g���~�b�g�`�F�b�N�L��)"
    '''=========================================================================
    ''' <summary>BP��Βl�ړ�(�\�t�g���~�b�g�`�F�b�N�L��)</summary>
    ''' <param name="SysPrm">(INP)�V�X�e���p�����[�^</param>
    ''' <param name="ObjSys">(INP)OcxSystem�I�u�W�F�N</param>
    ''' <param name="ObjUtl">(INP)OcxUtility�I�u�W�F�N</param>
    ''' <param name="stJOG"> (I/O)JOG����p�p�����[�^</param>
    ''' <returns>0=����, 0�ȊO:�G���[</returns>
    '''=========================================================================
    Private Function Sub_BPmove(ByVal SysPrm As SYSPARAM_PARAM, ByVal ObjSys As Object, ByVal ObjUtl As Object, ByRef stJOG As JOG_PARAM) As Integer

        Dim r As Integer
        Dim strMSG As String

        Try
            ' BP�ړ��ʂ̎Z�o(��X,Y)
            mvxBk = mvx                                             ' ���݂̈ʒu�ޔ�
            mvyBk = mvy
            'V2.2.0.0�@��
            If ((cin And CtrlJog.MouseClickLocation.Move) = &H0) Then           'V6.0.0.0�G
                Call ObjUtl.GetBPmovePitch(cin, X, Y, mPIT, mvx, mvy, SysPrm.stDEV.giBpDirXy)
            Else
                'V6.0.0.0�G              ��
                Dim dirX As Double = 0.0
                Dim dirY As Double = 0.0
                Dim tmpX As Double = 0.0
                Dim tmpY As Double = 0.0
                ObjUtl.GetBPmovePitch(cin, dirX, dirY, 1.0, tmpX, tmpY, SysPrm.stDEV.giBpDirXy)   ' �������擾

                X = Math.Abs(CtrlJog.MouseClickLocation.DistanceX) * Math.Sign(dirX)
                Y = Math.Abs(CtrlJog.MouseClickLocation.DistanceY) * Math.Sign(dirY)
                mvx -= X
                mvy -= Y
                'V6.0.0.0�G              ��
            End If
            'V2.2.0.0�@��

            ' BP��Βl�ړ�(�\�t�g���~�b�g�`�F�b�N�L��)
            r = ObjSys.BPMOVE(SysPrm, stJOG.BpOffX, stJOG.BpOffY, stJOG.BszX, stJOG.BszY, mvx, mvy, 1)
            If (r <> cFRS_NORMAL) Then                              ' �װ�Ȃ�װ����(���b�Z�[�W�\���ς�)
                If (ObjSys.IsSoftLimitBP(r) = False) Then
                    GoTo STP_END                                    ' ����ЯĴװ�ȊO�ʹװ����
                End If
                mvx = mvxBk                                         ' BP����ЯĴװ����BP�ʒu��߂�
                mvy = mvyBk
                GoTo STP_END                                        ' BP����ЯĴװ
            End If

            stJOG.cgX = stJOG.cgX + (-1 * X)                        ' BP�ړ���X�X�V (���ړ��ʂ͔��]���Ă���̂�-1���|����)
            stJOG.cgY = stJOG.cgY + (-1 * Y)                        ' BP�ړ���Y�X�V

STP_END:
            Return (r)

            ' �g���b�v�G���[������
        Catch ex As Exception
            strMSG = "Globals.Sub_BPmove() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)
        End Try
    End Function
#End Region

#Region "XY�e�[�u����Βl�ړ�(�\�t�g���~�b�g�`�F�b�N�L��)"
    '''=========================================================================
    ''' <summary>XY�e�[�u����Βl�ړ�(�\�t�g���~�b�g�`�F�b�N�L��)</summary>
    ''' <param name="SysPrm">(INP)�V�X�e���p�����[�^</param>
    ''' <param name="ObjSys">(INP)OcxSystem�I�u�W�F�N</param>
    ''' <param name="ObjUtl">(INP)OcxUtility�I�u�W�F�N</param>
    ''' <param name="stJOG"> (I/O)JOG����p�p�����[�^</param>
    ''' <returns>0=����, 0�ȊO:�G���[</returns>
    '''=========================================================================
    Private Function Sub_XYtableMove(ByVal SysPrm As SYSPARAM_PARAM, ByVal ObjSys As Object, ByVal ObjUtl As Object, ByRef stJOG As JOG_PARAM) As Integer

        Dim r As Integer
        Dim strMSG As String

        Try
            ' XY�e�[�u���ړ��ʂ̎Z�o(��X,Y)
            mvxBk = X                                               ' ���݂̈ʒu�ޔ�
            mvyBk = Y
            'V2.2.0.0�@ ��
            If ((cin And CtrlJog.MouseClickLocation.Move) = &H0) Then
                Call TrimClassCommon.GetXYmovePitch(cin, X, Y, mPIT, giStageYDir)
            Else
                Dim dirX As Double = 0.0
                Dim dirY As Double = 0.0
                TrimClassCommon.GetXYmovePitch(cin, dirX, dirY, 1.0, giStageYDir)   ' �������擾

                X = -(Math.Abs(CtrlJog.MouseClickLocation.DistanceX) * Math.Sign(dirX)) 'V6.0.0.0-24 -() �ǉ�
                Y = -(Math.Abs(CtrlJog.MouseClickLocation.DistanceY) * Math.Sign(dirY)) 'V6.0.0.0-24 -() �ǉ�
            End If
            'V2.2.0.0�@ ��

            ' XY�e�[�u����Βl�ړ�(�\�t�g���~�b�g�`�F�b�N�L��)
            r = ObjSys.XYtableMove(SysPrm, mvx + X, mvy + Y)
            If (r <> cFRS_NORMAL) Then                              ' �װ�Ȃ�װ����(���b�Z�[�W�\���ς�)
                If (ObjSys.IsSoftLimitXY(r) = False) Then
                    GoTo STP_END                                    ' ����ЯĴװ�ȊO�ʹװ����
                End If
                X = mvxBk                                           ' ����ЯĴװ����X,Y�ʒu��߂�
                Y = mvyBk
                GoTo STP_END                                        ' ����ЯĴװ
            End If

            mvx = mvx + X                                           ' �e�[�u���ʒuX,Y�X�V(��΍��W)
            mvy = mvy + Y
            stJOG.cgX = stJOG.cgX + X                               ' �e�[�u���ړ���X,Y�X�V
            stJOG.cgY = stJOG.cgY + Y

STP_END:
            Return (r)

            ' �g���b�v�G���[������
        Catch ex As Exception
            strMSG = "Globals.Sub_XYtableMove() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
            Return (cERR_TRAP)
        End Try
    End Function
#End Region

#Region "�{�^������������(�i�n�f������)"
    '========================================================================================
    '   �{�^������������(�i�n�f������)
    '========================================================================================
#Region "HALT�{�^������������"
    '''=========================================================================
    '''<summary>HALT�{�^������������</summary>
    '''=========================================================================
    Public Sub SubBtnHALT_Click()
        InpKey = CONSOLE_SW_HALT
    End Sub
#End Region

#Region "START�{�^������������"
    '''=========================================================================
    '''<summary>START�{�^������������</summary>
    '''=========================================================================
    Public Sub SubBtnSTART_Click()
        InpKey = CONSOLE_SW_START
    End Sub
#End Region

#Region "RESET�{�^������������"
    '''=========================================================================
    '''<summary>RESET�{�^������������</summary>
    '''=========================================================================
    Public Sub SubBtnRESET_Click()
        InpKey = CONSOLE_SW_RESET
    End Sub
#End Region

#Region "Z�{�^������������"
    '''=========================================================================
    '''<summary>RESET�{�^������������</summary>
    '''<param name="stJOG">(INP)JOG����p�p�����[�^</param>
    '''=========================================================================
    Public Sub SubBtnZ_Click(ByVal stJOG As JOG_PARAM)

        Dim strMSG As String

        Try
            If (stJOG.BtnZ.BackColor = System.Drawing.Color.Yellow) Then    ' Z SW ON ?
                stJOG.BtnZ.BackColor = System.Drawing.SystemColors.Control
                stJOG.BtnZ.Text = "Z Off"
                InpKey = InpKey And Not CONSOLE_SW_ZSW                      ' Z SW OFF
                bZ = False                                                  ' Z�L�[�ޔ���
            Else
                stJOG.BtnZ.BackColor = System.Drawing.Color.Yellow
                stJOG.BtnZ.Text = "Z On"
                InpKey = InpKey Or CONSOLE_SW_ZSW                           ' Z SW ON
                bZ = True                                                   ' Z�L�[�ޔ���
            End If

            ' �g���b�v�G���[������
        Catch ex As Exception
            strMSG = "Globals.SubBtnZ_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "HI�{�^������������"
    '''=========================================================================
    '''<summary>HI�{�^������������</summary>
    '''<param name="stJOG">(INP)JOG����p�p�����[�^</param>
    '''=========================================================================
    Public Sub SubBtnHI_Click(ByVal stJOG As JOG_PARAM)

        ' �w�i�F��ؑւ���
        If (stJOG.BtnHI.BackColor = System.Drawing.Color.Yellow) Then   ' �w�i�F = ���F ?
            ' �w�i�F���f�t�H���g�ɂ���
            stJOG.BtnHI.BackColor = System.Drawing.SystemColors.Control
            InpKey = InpKey And Not cBIT_HI                             ' HI SW OFF
        Else
            ' �w�i�F�����F�ɂ���
            stJOG.BtnHI.BackColor = System.Drawing.Color.Yellow
            InpKey = InpKey Or cBIT_HI                                  ' HI SW ON
        End If

    End Sub
#End Region

#Region "InpKey���擾����"
    '''=========================================================================
    '''<summary>InpKey���擾����</summary>
    '''<param name="IKey">(OUT)InpKey</param>
    '''=========================================================================
    Public Sub GetInpKey(ByRef IKey As UShort)
        IKey = InpKey
    End Sub
#End Region

#Region "InpKey��ݒ肷��"
    '''=========================================================================
    '''<summary>InpKey��ݒ肷��</summary>
    '''<param name="IKey">(INP)InpKey</param>
    '''=========================================================================
    Public Sub PutInpKey(ByVal IKey As UShort)
        InpKey = IKey
    End Sub
#End Region

#Region "���{�^��������"
    '''=========================================================================
    '''<summary>���{�^��������</summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub SubBtnJOG_0_MouseDown()
        InpKey = InpKey Or &H1000US                         ' +Y ON
    End Sub
    Public Sub SubBtnJOG_0_MouseUp()
        InpKey = InpKey And Not &H1000US                    ' +Y OFF
    End Sub

    Public Sub SubBtnJOG_1_MouseDown()
        InpKey = InpKey Or &H800US                          ' -Y ON
    End Sub
    Public Sub SubBtnJOG_1_MouseUp()
        InpKey = InpKey And Not &H800US                     ' -Y OFF
    End Sub

    Public Sub SubBtnJOG_2_MouseDown()
        InpKey = InpKey Or &H400US                          ' +X ON
    End Sub
    Public Sub SubBtnJOG_2_MouseUp()
        InpKey = InpKey And Not &H400US                     ' +X OFF
    End Sub

    Public Sub SubBtnJOG_3_MouseDown()
        InpKey = InpKey Or &H200US                          ' -X ON
    End Sub
    Public Sub SubBtnJOG_3_MouseUp()
        InpKey = InpKey And Not &H200US                     ' -X OFF
    End Sub

    Public Sub SubBtnJOG_4_MouseDown()
        InpKey = InpKey Or &HA00US                          ' -X -Y ON
    End Sub
    Public Sub SubBtnJOG_4_MouseUp()
        InpKey = InpKey And Not &HA00US                     ' -X -Y OFF
    End Sub

    Public Sub SubBtnJOG_5_MouseDown()
        InpKey = InpKey Or &HC00US                          ' +X -Y ON
    End Sub
    Public Sub SubBtnJOG_5_MouseUp()
        InpKey = InpKey And Not &HC00US                     ' +X -Y OFF
    End Sub

    Public Sub SubBtnJOG_6_MouseDown()
        InpKey = InpKey Or &H1400US                         ' +X +Y ON
    End Sub
    Public Sub SubBtnJOG_6_MouseUp()
        InpKey = InpKey And Not &H1400US                    ' +X +Y OFF
    End Sub

    Public Sub SubBtnJOG_7_MouseDown()
        InpKey = InpKey Or &H1200US                         ' -X +Y ON
    End Sub
    Public Sub SubBtnJOG_7_MouseUp()
        InpKey = InpKey And Not &H1200US                    ' -X +Y OFF
    End Sub
#End Region

#End Region

#Region "�i�n�f�����ʏ����p�g���b�N�o�[����"
    '========================================================================================
    '   �i�n�f�����ʏ����p�g���b�N�o�[����
    '========================================================================================
#Region "�g���b�N�o�[�̃X���C�_�[��ʏ����l�\��"
    '''=========================================================================
    '''<summary>�g���b�N�o�[�̃X���C�_�[��ʏ����l�\��</summary>
    '''<param name="TBarLowPitch">(I/O)�X���C�_�[1(Low�߯�)</param>
    '''<param name="TBarHiPitch"> (I/O)�X���C�_�[2(HIGH�߯�)</param>
    '''<param name="TBarPause">   (I/O)�X���C�_�[3(Pause Time)</param>
    '''<param name="LblTchMoval0">(I/O)�ڐ�1(Low Pich Label)</param>
    '''<param name="LblTchMoval1">(I/O)�ڐ�2(Low�߯� Label)</param>
    '''<param name="LblTchMoval2">(I/O)�ڐ�3(HIGH�߯� Label)</param>
    '''<param name="dblTchMoval"> (I/O)�߯��ޔ���(0=�߯�, 1=HIGH�߯�, 2=Pause Time)</param>
    '''=========================================================================
    Public Sub XyzBpMovingPitchInit(ByRef TBarLowPitch As System.Windows.Forms.TrackBar, _
                                    ByRef TBarHiPitch As System.Windows.Forms.TrackBar, _
                                    ByRef TBarPause As System.Windows.Forms.TrackBar, _
                                    ByRef LblTchMoval0 As System.Windows.Forms.Label, _
                                    ByRef LblTchMoval1 As System.Windows.Forms.Label, _
                                    ByRef LblTchMoval2 As System.Windows.Forms.Label, _
                                    ByRef dblTchMoval() As Double)

        Dim minval As Short

        ' LOW�߯������͈͊O�Ȃ�͈͓��ɕύX����
        If (dblTchMoval(IDX_PIT) < cPT_LO) Then dblTchMoval(IDX_PIT) = cPT_LO
        If (dblTchMoval(IDX_PIT) > cPT_HI) Then dblTchMoval(IDX_PIT) = cPT_HI

        ' LOW�߯��̖ڐ���ݒ肷��
        If (dblTchMoval(IDX_PIT) < 0.002) Then                          ' ����\�ɂ��ŏ��ڐ���ݒ肷��
            minval = 1                                                  ' �ڐ�1�`
        Else
            minval = 2                                                  ' �ڐ�2�`
        End If

        'V2.2.1.1�B ��
        ' BP�ŏ�����\�ɂ���čŏ��l��ύX����
        If dblTchMoval(3) <> 0 Then
            minval = 1                                     ' �ڐ�2�` 
        End If
        'V2.2.1.1�B ��

        TBarLowPitch.TickFrequency = dblTchMoval(IDX_PIT) * 1000        ' 0.001mm�P��
        TBarLowPitch.Maximum = 100                                      ' �ڐ�1(or 2)�`100(0.001m�`0.1mm)
        TBarLowPitch.Minimum = minval
        TBarLowPitch.Value = dblTchMoval(IDX_PIT) * 1000        ' 0.001mm�P��

        ' HIGH�߯������͈͊O�Ȃ�͈͓��ɕύX����
        If (dblTchMoval(IDX_HPT) < cHPT_LO) Then dblTchMoval(IDX_HPT) = cHPT_LO
        If (dblTchMoval(IDX_HPT) > cHPT_HI) Then dblTchMoval(IDX_HPT) = cHPT_HI

        ' HIGH�߯��̖ڐ���ݒ肷��
        TBarHiPitch.TickFrequency = dblTchMoval(IDX_HPT) * 100          ' 0.01mm�P��
        TBarHiPitch.Maximum = 500                                       ' �ڐ�1�`100(0.01m�`5.00mm)
        TBarHiPitch.Minimum = 1
        TBarHiPitch.Value = dblTchMoval(IDX_HPT) * 100          ' 0.01mm�P��

        ' Pause Time���͈͊O�Ȃ�͈͓��ɕύX����
        If (dblTchMoval(IDX_PAU) < cPAU_LO) Then dblTchMoval(IDX_PAU) = cPAU_LO
        If (dblTchMoval(IDX_PAU) > cPAU_HI) Then dblTchMoval(IDX_PAU) = cPAU_HI

        ' Pause Time�̖ڐ���ݒ肷��
        TBarPause.TickFrequency = dblTchMoval(IDX_PAU) * 20             ' 0.5�b�P��
        TBarPause.Maximum = 20                                          ' �ڐ�1�`20(0.05�b�`1.00�b)
        TBarPause.Minimum = 1
        TBarPause.Value = dblTchMoval(IDX_PAU) * 20             ' 0.5�b�P��

        ' �ړ��s�b�`��\������
        LblTchMoval0.Text = dblTchMoval(IDX_PIT).ToString("0.0000")
        LblTchMoval1.Text = dblTchMoval(IDX_HPT).ToString("0.0000")
        LblTchMoval2.Text = dblTchMoval(IDX_PAU).ToString("0.0000")

    End Sub
#End Region

#Region "�g���b�N�o�[�̃X���C�_�[�ړ�����"
    '''=========================================================================
    '''<summary>�g���b�N�o�[�̃X���C�_�[�ړ�����</summary>
    '''<param name="Index">       (INP)0=LOW�߯�, 1=HIGH�߯�, 2=Pause</param>
    '''<param name="TBarLowPitch">(I/O)�X���C�_�[1(Low�߯�)</param>
    '''<param name="TBarHiPitch"> (I/O)�X���C�_�[2(HIGH�߯�)</param>
    '''<param name="TBarPause">   (I/O)�X���C�_�[3(Pause Time)</param>
    '''<param name="LblTchMoval0">(I/O)�ڐ�1(Low Pich Label)</param>
    '''<param name="LblTchMoval1">(I/O)�ڐ�2(Low�߯� Label)</param>
    '''<param name="LblTchMoval2">(I/O)�ڐ�3(HIGH�߯� Label)</param>
    '''<param name="dblTchMoval"> (I/O)�߯��ޔ���(0=�߯�, 1=HIGH�߯�, 2=Pause Time)</param>
    '''=========================================================================
    Public Sub SetSliderPitch(ByRef Index As Short, _
                              ByRef TBarLowPitch As System.Windows.Forms.TrackBar, _
                              ByRef TBarHiPitch As System.Windows.Forms.TrackBar, _
                              ByRef TBarPause As System.Windows.Forms.TrackBar, _
                              ByRef LblTchMoval0 As System.Windows.Forms.Label, _
                              ByRef LblTchMoval1 As System.Windows.Forms.Label, _
                              ByRef LblTchMoval2 As System.Windows.Forms.Label, _
                              ByRef dblTchMoval() As Double)

        Dim lVal As Integer

        ' BP�̈ړ��s�b�`����ݒ肷��
        Select Case Index
            Case IDX_PIT    ' LOW�߯�
                lVal = TBarLowPitch.Value                       ' �ײ�ޖڐ��l�擾
                dblTchMoval(Index) = 0.001 * lVal               ' LOW�߯��l�ύX
                LblTchMoval0.Text = dblTchMoval(Index).ToString("0.0000")
                LblTchMoval0.Refresh()

            Case IDX_HPT    ' HIGH�߯�
                lVal = TBarHiPitch.Value                        ' �ײ�ޖڐ��l�擾
                dblTchMoval(Index) = 0.01 * lVal                ' HIGH�߯��l�ύX
                LblTchMoval1.Text = dblTchMoval(Index).ToString("0.0000")
                LblTchMoval1.Refresh()

            Case IDX_PAU    ' Pause Time
                lVal = TBarPause.Value                          ' �ײ�ޖڐ��l�擾
                dblTchMoval(Index) = 0.05 * lVal                ' �ړ��s�b�`�Ԃ̃|�[�Y�l�ύX
                LblTchMoval2.Text = dblTchMoval(Index).ToString("0.0000")
                LblTchMoval2.Refresh()
        End Select

    End Sub
#End Region

#End Region

#Region "�i�n�f�����ʏ����p�e���L�[���͏���"
    '========================================================================================
    '   �i�n�f�����ʏ����p�e���L�[���͏���
    '========================================================================================
#Region "�e���L�[�_�E���T�u���[�`��"
    '''=========================================================================
    '''<summary>�e���L�[�_�E���T�u���[�`��</summary>
    ''' <param name="KeyCode">(INP)�L�[�R�[�h</param>
    '''=========================================================================
    Public Sub Sub_10KeyDown(ByVal KeyCode As Short)

        Dim strMSG As String

        Try
            ' Num Lock��
            Select Case (KeyCode)
                Case System.Windows.Forms.Keys.NumPad2                      ' ��  (KeyCode =  98(&H62)
                    InpKey = InpKey Or &H1000                               ' +Y ON(��)
                Case System.Windows.Forms.Keys.NumPad8                      ' ��  (KeyCode = 104(&H68)
                    InpKey = InpKey Or &H800                                ' -Y ON(��)
                Case System.Windows.Forms.Keys.NumPad4                      ' ��  (KeyCode = 100(&H64)
                    InpKey = InpKey Or &H400                                ' +X ON(��)
                Case System.Windows.Forms.Keys.NumPad6                      ' ��  (KeyCode = 102(&H66)
                    InpKey = InpKey Or &H200                                ' -X ON(��)
                Case System.Windows.Forms.Keys.NumPad9                      ' PgUp(KeyCode = 105(&H69)
                    InpKey = InpKey Or &HA00                                ' -X -Y ON
                Case System.Windows.Forms.Keys.NumPad7                      ' Home(KeyCode = 103(&H67))
                    InpKey = InpKey Or &HC00                                ' +X -Y ON
                Case System.Windows.Forms.Keys.NumPad1                      ' End(KeyCode =   97(&H61)
                    InpKey = InpKey Or &H1400                               ' +X +Y ON
                Case System.Windows.Forms.Keys.NumPad3                      ' PgDn(KeyCode =  99(&H63)
                    InpKey = InpKey Or &H1200                               ' -X +Y ON
                Case System.Windows.Forms.Keys.NumPad5                      ' 5�� (KeyCode = 101(&H65)
                    'Call BtnHI_Click(sender, e)                             ' HI�{�^�� ON/OFF
            End Select

            ' �g���b�v�G���[������
        Catch ex As Exception
            strMSG = "Globals.Sub_10KeyDown() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�e���L�[�A�b�v�T�u���[�`��"
    '''=========================================================================
    '''<summary>�e���L�[�A�b�v�T�u���[�`��</summary>
    ''' <param name="KeyCode">(INP)�L�[�R�[�h</param>
    '''=========================================================================
    Public Sub Sub_10KeyUp(ByVal KeyCode As Short)

        Dim strMSG As String

        Try
            ' Num Lock��
            Select Case (KeyCode)
                Case System.Windows.Forms.Keys.NumPad2                      ' ��  (KeyCode =  98(&H62)
                    InpKey = InpKey And Not &H1000                          ' +Y OFF
                Case System.Windows.Forms.Keys.NumPad8                      ' ��  (KeyCode = 104(&H68)
                    InpKey = InpKey And Not &H800                           ' -Y OFF
                Case System.Windows.Forms.Keys.NumPad4                      ' ��  (KeyCode = 100(&H64)
                    InpKey = InpKey And Not &H400                           ' +X OFF
                Case System.Windows.Forms.Keys.NumPad6                      ' ��  (KeyCode = 102(&H66)
                    InpKey = InpKey And Not &H200                           ' -X OFF
                Case System.Windows.Forms.Keys.NumPad9                      ' PgUp(KeyCode = 105(&H69)
                    InpKey = InpKey And Not &HA00                           ' -X -Y OFF
                Case System.Windows.Forms.Keys.NumPad7                      ' Home(KeyCode = 103(&H67))
                    InpKey = InpKey And Not &HC00                           ' +X -Y OFF
                Case System.Windows.Forms.Keys.NumPad1                      ' End(KeyCode =   97(&H61)
                    InpKey = InpKey And Not &H1400                          ' +X +Y OFF
                Case System.Windows.Forms.Keys.NumPad3                      ' PgDn(KeyCode =  99(&H63)
                    InpKey = InpKey And Not &H1200                          ' -X +Y OFF
            End Select

            ' �g���b�v�G���[������
        Catch ex As Exception
            strMSG = "Globals.Sub_10KeyUp() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "���̓L�[�R�[�h�̃N���A"
    '''=========================================================================
    ''' <summary>
    ''' ���̓L�[�R�[�h�̃N���A
    ''' </summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Sub ClearInpKey()

        Try
            InpKey = 0

            ' �g���b�v�G���[������
        Catch ex As Exception
            MsgBox("Globals.ClearInpKey() TRAP ERROR = " + ex.Message)
        End Try
    End Sub
#End Region

#End Region

#End Region

#Region "�O���[�v��,�u���b�N��,�`�b�v��(��R��),�`�b�v�T�C�Y���擾����(�s�w/�s�x�e�B�[�`���O�p)"
    '''=========================================================================
    ''' <summary>�O���[�v��,�u���b�N��,�`�b�v��(��R��),�`�b�v�T�C�Y���擾����</summary>
    ''' <param name="AppMode">  (INP)���[�h</param>
    ''' <param name="Gn">       (OUT)�O���[�v��</param>
    ''' <param name="RnBn">     (OUT)�`�b�v��(�s�w�e�B�[�`���O��)�܂���
    '''                              �u���b�N��(�s�x�e�B�[�`���O��)</param>
    ''' <param name="DblChipSz">(OUT)�`�b�v�T�C�Y</param>
    ''' <returns>0=����, 0�ȊO=�G���[</returns>
    '''=========================================================================
    Public Function GetChipNumAndSize(ByVal AppMode As Short, ByRef Gn As Short, ByRef RnBn As Short, ByRef DblChipSz As Double) As Short

        Dim ChipNum As Short                                        ' �`�b�v��(��R��)
        Dim ChipSzX As Double                                       ' �`�b�v�T�C�YX
        Dim ChipSzY As Double                                       ' �`�b�v�T�C�YY
        Dim strMSG As String

        Try
            ' �O����(CHIP/NET����)
            ChipNum = typPlateInfo.intResistCntInGroup              ' �`�b�v��(��R��) = 1�O���[�v��(1�T�[�L�b�g��)��R��
            ChipSzX = typPlateInfo.dblChipSizeXDir                  ' �`�b�v�T�C�YX,Y
            ChipSzY = typPlateInfo.dblChipSizeYDir

            ' �v���[�g�f�[�^����O���[�v��, �u���b�N��, �`�b�v��(��R��), �`�b�v�T�C�Y���擾����
            If (AppMode = APP_MODE_TX) Then
                '----- �s�w�e�B�[�`���O�� -----
                ' �`�b�v��(��R��)��Ԃ�
                RnBn = ChipNum                                      ' 1�O���[�v��(1�T�[�L�b�g��)��R�����Z�b�g
                ' �O���[�v����Ԃ�
                Gn = typPlateInfo.intGroupCntInBlockXBp             ' �a�o�O���[�v��(�T�[�L�b�g��)���Z�b�g
                ' �`�b�v�T�C�Y��Ԃ�
                If (typPlateInfo.intResistDir = 0) Then             ' �`�b�v���т�X���� ?
                    DblChipSz = System.Math.Abs(ChipSzX)
                Else
                    'V2.0.0.0�@                    DblChipSz = System.Math.Abs(ChipSzY)
                    DblChipSz = ChipSzY             'V2.0.0.0�@
                End If

            Else
                '----- �s�x�e�B�[�`���O�� -----
                ' �O���[�v����Ԃ�
                Gn = typPlateInfo.intGroupCntInBlockYStage          ' �u���b�N��Stage�O���[�v�����Z�b�g
                ' �u���b�N���ƃ`�b�v�T�C�Y��Ԃ�
                If (typPlateInfo.intResistDir = 0) Then             ' �`�b�v���т�X���� ?
                    RnBn = typPlateInfo.intBlockCntYDir             ' �u���b�N��Y���Z�b�g
                    DblChipSz = System.Math.Abs(ChipSzY)            ' �`�b�v�T�C�YY���Z�b�g
                Else
                    RnBn = typPlateInfo.intBlockCntXDir             ' �u���b�N��X���Z�b�g
                    DblChipSz = System.Math.Abs(ChipSzX)            ' �`�b�v�T�C�YX���Z�b�g
                End If
            End If

            strMSG = "GetChipNumAndSize() Gn=" + Gn.ToString("0") + ", RnBn=" + RnBn.ToString("0") + ", ChipSZ=" + DblChipSz.ToString("0.00000")
            Console.WriteLine(strMSG)
            Return (cFRS_NORMAL)                                    ' Return�l = ����

            ' �g���b�v�G���[������
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
            Return (cERR_TRAP)                                      ' Return�l = ��O�G���[
        End Try
    End Function
#End Region

#Region "�u���b�N�T�C�Y���Z�o����yCHIP/NET�p�z"
    '''=========================================================================
    '''<summary>�u���b�N�T�C�Y���Z�o����yCHIP/NET�p�z</summary>
    '''<param name="dblBSX">(OUT) �u���b�N�T�C�YX</param>
    '''<param name="dblBSY">(OUT) �u���b�N�T�C�YY</param>
    '''=========================================================================
    Public Sub CalcBlockSize(ByRef dblBSX As Double, ByRef dblBSY As Double)

        Dim i As Integer
        Dim intChipNum As Integer
        Dim intGNx As Integer
        Dim intGNY As Integer
        Dim dData As Double = 0.0

        Try
            ' CHIP/NET�� 
            ' �O���[�v��X,Y
            intGNx = typPlateInfo.intGroupCntInBlockXBp                 ' �a�o�O���[�v��(�T�[�L�b�g��)
            'V2.0.0.0�N            intGNY = typPlateInfo.intGroupCntInBlockXBp
            intGNY = typPlateInfo.intGroupCntInBlockYStage

            ' �O���[�v����R��             
            intChipNum = typPlateInfo.intResistCntInGroup

            ' �u���b�N�T�C�YX,Y�����߂�
            If (typPlateInfo.intResistDir = 0) Then                     ' ��R(����)���ѕ���(0:X, 1:Y)
                ' ��R(����)���ѕ��� = X�����̏ꍇ
                If (intGNx = 1) Then
                    ' 1�O���[�v(1�T�[�L�b�g)�̏ꍇ
                    dData = typPlateInfo.dblChipSizeXDir * intChipNum   ' Data = �`�b�v�T�C�YX * �`�b�v��

                Else
                    ' �����O���[�v(�����T�[�L�b�g)�̏ꍇ
                    For i = 1 To intGNx
                        If (i = intGNx) Then                            ' �ŏI�O���[�v ?
                            ' Data = Data + (�`�b�v�T�C�YX * �O���[�v��(�T�[�L�b�g��)��R��)
                            dData = dData + (typPlateInfo.dblChipSizeXDir * typPlateInfo.intResistCntInGroup)
                        Else
                            ' Data = Data + (�`�b�v�T�C�YX * �O���[�v��(�T�[�L�b�g��)��R�� + �a�o�O���[�v(�T�[�L�b�g)�Ԋu)
                            dData = dData + (typPlateInfo.dblChipSizeXDir * typPlateInfo.intResistCntInGroup + typPlateInfo.dblBpGrpItv)
                        End If
                    Next i
                End If

                ' �u���b�N�T�C�YX,Y��Ԃ�
                dblBSX = dData                                          ' �u���b�N�T�C�YX = �v�Z�l
                dblBSY = typPlateInfo.dblChipSizeYDir                   ' �u���b�N�T�C�YY = �`�b�v�T�C�YY

            Else
                ' ��R(����)���ѕ��� = Y�����̏ꍇ
                If (intGNY = 1) Then
                    ' 1�O���[�v(1�T�[�L�b�g)�̏ꍇ
                    dData = typPlateInfo.dblChipSizeYDir * intChipNum   ' Data = �`�b�v�T�C�YY * �`�b�v��

                Else
                    ' �����O���[�v(�����T�[�L�b�g)�̏ꍇ
                    For i = 1 To intGNY
                        If (i = intGNY) Then                            ' �ŏI�O���[�v ?
                            ' Data = Data + (�`�b�v�T�C�YY * �O���[�v��(�T�[�L�b�g��)��R��)
                            dData = dData + (typPlateInfo.dblChipSizeYDir * typPlateInfo.intResistCntInGroup)
                        Else
                            ' Data = Data + (�`�b�v�T�C�YY * �O���[�v��(�T�[�L�b�g��)��R�� + �a�o�O���[�v(�T�[�L�b�g)�Ԋu)
                            dData = dData + (typPlateInfo.dblChipSizeYDir * typPlateInfo.intResistCntInGroup + typPlateInfo.dblBpGrpItv)
                        End If
                    Next i

                End If

                ' �u���b�N�T�C�YX,Y��Ԃ�
                dblBSX = typPlateInfo.dblChipSizeXDir                   ' �u���b�N�T�C�YX = �`�b�v�T�C�YX
                dblBSY = dData                                          ' �u���b�N�T�C�YY = �v�Z�l

                'V2.0.0.0�@���X�e�b�v�ʒu
                If (giAppMode = APP_MODE_TY) Then
                    dblBSY = Math.Abs(dData)                                          ' �u���b�N�T�C�YY = �v�Z�l
                End If
                'V2.0.0.0�@��
            End If

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try

    End Sub

#End Region

#Region "BP����ۯ��̉E��Ɉړ�����"
    '''=========================================================================
    '''<summary>BP����ۯ��̉E��Ɉړ�����</summary>
    '''<remarks></remarks>
    '''=========================================================================
    Public Sub BpMoveOrigin_Ex()
        Try
            Dim dblBpOffsX As Double
            Dim dblBpOffsY As Double
            Dim dblBSX As Double
            Dim dblBSY As Double

            ' ��ۯ����ގ擾
            Call CalcBlockSize(dblBSX, dblBSY)
            ' BP�ʒu�̾��X,Y�ݒ�
            dblBpOffsX = typPlateInfo.dblBpOffSetXDir
            dblBpOffsY = typPlateInfo.dblBpOffSetYDir
            ' BP����ۯ��̉E��Ɉړ�����(BSIZE()��BPOFF()���s)
            Call Form1.System1.BpMoveOrigin(gSysPrm, dblBSX, dblBSY, dblBpOffsX, dblBpOffsY)
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub
#End Region

#Region "�w���R�ԍ��A��Ĕԍ��̽����߲�Ă�Ԃ�"
    '''=========================================================================
    '''<summary>�w���R�ԍ��A��Ĕԍ��̽����߲�Ă�Ԃ�</summary>
    '''<param name="intRegNo">(INP) ��R�ԍ�</param>
    '''<param name="intCutNo">(INP) ��Ĕԍ�</param>
    '''<param name="dblX"    >(OUT) �����߲��X</param>
    '''<param name="dblY"    >(OUT) �����߲��Y</param>
    '''<returns>TRUE:�ް�����, FALSE:�ް��Ȃ�</returns>
    '''=========================================================================
    Public Function GetCutStartPoint(ByRef intRegNo As Short, ByRef intCutNo As Short, ByRef dblX As Double, ByRef dblY As Double) As Boolean
        Try
            Dim bRetc As Boolean
            Dim i As Short
            Dim j As Short

            bRetc = False
            For i = 1 To MAXRNO
                If (intRegNo = typResistorInfoArray(i).intResNo) Then                       ' ��R�ԍ���v
                    For j = 1 To MaxCntCut
                        If (intCutNo = typResistorInfoArray(i).ArrCut(j).intCutNo) Then     ' ��Ĕԍ���v
                            dblX = typResistorInfoArray(i).ArrCut(j).dblStartPointX         ' �����߲��
                            dblY = typResistorInfoArray(i).ArrCut(j).dblStartPointY
                            bRetc = True
                            GetCutStartPoint = bRetc
                            Exit Function
                        End If
                    Next
                End If
            Next
            GetCutStartPoint = bRetc
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region

#Region "�������[�U�p���[��������"
    '''=========================================================================
    ''' <summary>
    ''' �������[�U�p���[��������
    ''' </summary>
    ''' <param name="bPowerMonitoring">True:�t���p���[����</param>
    ''' <returns>cFRS_NORMAL  = ����,cFRS_ERR_RST = Cancel(RESET��),��L�ȊO �@�@= ����~���o���̃G���[</returns>
    ''' <remarks>�������[�U�p���[�̒����������s</remarks>
    '''=========================================================================
    Public Function AutoLaserPowerADJ(Optional ByVal bPowerMonitoring As Boolean = False) As Short

        Dim r As Integer

        Try
            Dim strMsg As String

            With stLASER
                ' �p���[���[�^�̃f�[�^�擾�^�C�v���X�e�[�W�ݒu�^�C�v�łȂ��u�h�^�n�ǎ��v/�u�t�r�a�v�łȂ����NOP(���̂܂ܔ�����)

                If (gSysPrm.stIOC.giPM_Tp <> 1 Or gSysPrm.stIOC.giPM_DataTp = PM_DTTYPE_NONE) Then
                    Return (cFRS_NORMAL)
                End If

                ' �p���[�������s�t���O
                If Not bPowerMonitoring Then
                    If (.intPowerAdjustMode <> 1) Then
                        ' ��ܰ���������s���Ȃ��ꍇ�͂��̂܂ܔ�����
                        Return (cFRS_NORMAL)
                    End If
                End If

                ' Z�����_�ֈړ�
                r = EX_ZMOVE(0)
                If (r <> cFRS_NORMAL) Then                              ' �G���[ ?(���b�Z�[�W�͕\���ς�) 
                    Return (r)                                          ' Return�l�ݒ� 
                End If

                '---------------------------------------------------------------
                '   �����p���[�������s
                '---------------------------------------------------------------
                If (gSysPrm.stRAT.giOsc_Res = OSCILLATOR_FL) Then
#If cOSCILLATORcFLcUSE Then
                    Dim iCurr As Integer
                    Dim iCurrOfs As Integer
                    Dim dMeasPower As Double
                    Dim dFullPower As Double
                    Dim AdjustTarget As Double
                    Dim AdjustLevel As Double
                    Dim CndNum As Integer
                    '-----------------------------------------------------------
                    '   FL��
                    '-----------------------------------------------------------

                    ' �p���[����������H�����ԍ��z��ɗL��/������ݒ肷��
                    r = SetAutoPowerCndNumAry(stPWR)

                    ' �J�b�g�Ɏg�p������H�����ԍ��̃p���[�������s�� 
                    For CndNum = 0 To (MAX_BANK_NUM - 1)
                        If (stPWR.CndNumAry(CndNum) = 1) Then               ' ���H�����͗L�� ?
                            AdjustTarget = stPWR.AdjustTargetAry(CndNum)    ' �ڕW�p���[�l(W)
                            AdjustLevel = stPWR.AdjustLevelAry(CndNum)      ' �������e�͈�(�}W)

                            ' ���b�Z�[�W�\��("�p���[�����J�n"+ " ���H�����ԍ�xx")
                            strMsg = MSG_AUTOPOWER_01 + " " + MSG_AUTOPOWER_02 + CndNum.ToString("00")
                            Call Z_PRINT(strMsg)

                            ' �p���[�������s��
                            r = Form1.System1.Form_FLAutoLaser(gSysPrm, CndNum, AdjustTarget, AdjustLevel, iCurr, iCurrOfs, dMeasPower, dFullPower)
                            If (r < cFRS_NORMAL) Then
                                ' �G���[���b�Z�[�W�\��
                                r = Form1.System1.Form_AxisErrMsgDisp(System.Math.Abs(r))
                                Return (r)
                            End If

                            ' �������ʂ����C����ʂɕ\������
                            If (r = cFRS_NORMAL) Then                   ' ����I�� ? 
                                ' ���b�Z�[�W�\��("���[�U�p���[�ݒ�l"+" = xx.xxW, " + "�d���l=" + "xxxmA")
                                strMsg = MSG_AUTOPOWER_03 + "= " + dMeasPower.ToString("0.00") + "W, "
                                strMsg = strMsg + MSG_AUTOPOWER_04 + "= " + iCurr.ToString("0") + "mA"
                                Call Z_PRINT(strMsg)
                                stCND.Curr(CndNum) = iCurr              ' �d���l�ݒ�
                            Else
                                ' ���b�Z�[�W�\��("�p���[����������")
                                strMsg = MSG_AUTOPOWER_05
                                Call Z_PRINT(strMsg)
                                Exit For                                ' �����I��
                            End If
                        End If
                    Next CndNum

#End If
                Else
                    '-----------------------------------------------------------
                    '   FL�ȊO�̏ꍇ
                    '-----------------------------------------------------------
                    r = Form1.System1.Form_AutoLaser(gSysPrm, .dblPowerAdjustQRate,
                                        .dblPowerAdjustTarget, .dblPowerAdjustToleLevel, bPowerMonitoring)

                    If (r = cFRS_NORMAL) Then                           ' ����I�� ? 
                        ' ���b�Z�[�W�\��("�p���[��������I��")
                        strMsg = MSG_AUTOPOWER_06
                        Call Z_PRINT(strMsg)
                    Else
                        ' ���b�Z�[�W�\��("�p���[����������")
                        strMsg = MSG_AUTOPOWER_05                       ' "�p���[����������"
                        If (bPowerMonitoring = True) Then               ' ���[�U�p���[�̃��j�^�����O ?
                            strMsg = MSG_165                            ' "���[�U�p���[���j�^�����O�ُ�I��"
                        End If
                        Call Z_PRINT(strMsg)
                    End If
                End If

                System.Windows.Forms.Application.DoEvents()
            End With
            Return (r)                                                  ' Return�l�ݒ� 

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
            Return (cERR_TRAP)                                          ' Return�l = �g���b�v�G���[����
        End Try
    End Function
#End Region

#Region "�p���[����������H�����ԍ��z��ɗL��/������ݒ肷��"
#If cOSCILLATORcFLcUSE Then
    '''=========================================================================
    ''' <summary>�p���[����������H�����ԍ��z��ɗL��/������ݒ肷��</summary>
    ''' <param name="stPWR">(OUT)FL�p�p���[�������
    '''                              ���z���0�I���W��</param>
    ''' <remarks>�������[�U�p���[�̒����������s�p</remarks>
    ''' <returns>cFRS_NORMAL  = ����
    '''          ��L�ȊO �@�@= �G���[</returns> 
    '''=========================================================================
    Private Function SetAutoPowerCndNumAry(ByRef stPWR As POWER_ADJUST_INFO) As Short

        Dim Rn As Integer
        Dim Cn As Integer
        Dim CndNum As Integer
        Dim CutType As Integer

        Try
            '------------------------------------------------------------------
            '   ��������
            '------------------------------------------------------------------
            ' FL�łȂ����NOP
            If (gSysPrm.stRAT.giOsc_Res <> OSCILLATOR_FL) Then Return (cFRS_NORMAL)

            ' �g���}�[���H�����\����(FL�p)������ 
            stPWR.Initialize()

            '------------------------------------------------------------------
            '   ���H�����ԍ��z���ݒ肷��
            '------------------------------------------------------------------
            For Rn = 1 To stPLT.RCount              ' �P�u���b�N����R�����`�F�b�N���� 
                If UserModule.IsCutResistorIncMarking(Rn) Then
                    For Cn = 1 To stREG(Rn).intTNN      ' ��R���J�b�g�����`�F�b�N����
                        ' �J�b�g�^�C�v�擾
                        CutType = stREG(Rn).STCUT(Cn).intCTYP

                        ' ���H����1�͑S�J�b�g�������ɐݒ肷��
                        CndNum = stREG(Rn).STCUT(Cn).intCND(CUT_CND_L1)
                        If (stPWR.CndNumAry(CndNum) = 0) Then               ' ���� ? 
                            stPWR.CndNumAry(CndNum) = 1                     ' �L���ɐݒ�
                            ' �ڕW�p���[�l(W)�ƒ������e�͈�(�}W)��ݒ肷��
                            stPWR.AdjustTargetAry(CndNum) = stCND.dblPowerAdjustTarget(CndNum)
                            stPWR.AdjustLevelAry(CndNum) = stCND.dblPowerAdjustToleLevel(CndNum)
                        End If

                        ' ���H����2��L�J�b�g, �΂�L�J�b�g, L�J�b�g(����/��ڰ�), �΂�L�J�b�g(����/��ڰ�)
                        ' HOOK�J�b�g, U�J�b�g���ɐݒ肷��
                        If (CutType = CNS_CUTP_L) Or (CutType = CNS_CUTP_NL) Or _
                           (CutType = CNS_CUTP_Lr) Or (CutType = CNS_CUTP_Lt) Or _
                           (CutType = CNS_CUTP_NLr) Or (CutType = CNS_CUTP_NLt) Or _
                           (CutType = CNS_CUTP_HK) Or (CutType = CNS_CUTP_U) Then
                            ' ���H����2
                            CndNum = stREG(Rn).STCUT(Cn).intCND(CUT_CND_L2)
                            If (stPWR.CndNumAry(CndNum) = 0) Then           ' ���� ? 
                                stPWR.CndNumAry(CndNum) = 1                 ' �L���ɐݒ�
                                ' �ڕW�p���[�l(W)�ƒ������e�͈�(�}W)��ݒ肷��
                                stPWR.AdjustTargetAry(CndNum) = stCND.dblPowerAdjustTarget(CndNum)
                                stPWR.AdjustLevelAry(CndNum) = stCND.dblPowerAdjustToleLevel(CndNum)
                            End If
                        End If

                        ' ���H����3��HOOK�J�b�g, U�J�b�g���ɐݒ肷��
                        If (CutType = CNS_CUTP_HK) Or (CutType = CNS_CUTP_U) Then
                            ' ���H����3
                            CndNum = stREG(Rn).STCUT(Cn).intCND(CUT_CND_L3)
                            If (stPWR.CndNumAry(CndNum) = 0) Then           ' ���� ? 
                                stPWR.CndNumAry(CndNum) = 1                 ' �L���ɐݒ�
                                ' �ڕW�p���[�l(W)�ƒ������e�͈�(�}W)��ݒ肷��
                                stPWR.AdjustTargetAry(CndNum) = stCND.dblPowerAdjustTarget(CndNum)
                                stPWR.AdjustLevelAry(CndNum) = stCND.dblPowerAdjustToleLevel(CndNum)
                            End If
                        End If

                        ' ���H����4�͌���͖��g�p(�\��)

                        ' ���H����5�`8�̓��^�[��/���g���[�X�p 
                        ' ���H����5(ST�J�b�g(����/��ڰ�), �΂�ST�J�b�g(����/��ڰ�)��
                        If (CutType = CNS_CUTP_STr) Or (CutType = CNS_CUTP_STt) Or _
                           (CutType = CNS_CUTP_NSTr) Or (CutType = CNS_CUTP_NSTt) Then
                            ' ���H����5�̏����ԍ����J�b�g�f�[�^�̉��H����2���ݒ肷��
                            CndNum = stREG(Rn).STCUT(Cn).intCND(CUT_CND_L1)
                            If (stPWR.CndNumAry(CndNum) = 0) Then           ' ���� ? 
                                stPWR.CndNumAry(CndNum) = 1                 ' �L���ɐݒ�
                                ' �ڕW�p���[�l(W)�ƒ������e�͈�(�}W)��ݒ肷��
                                stPWR.AdjustTargetAry(CndNum) = stCND.dblPowerAdjustTarget(CndNum)
                                stPWR.AdjustLevelAry(CndNum) = stCND.dblPowerAdjustToleLevel(CndNum)
                            End If
                        End If

                        ' ���H����5,6(L�J�b�g(����/��ڰ�), �΂�L�J�b�g(����/��ڰ�)��
                        If (CutType = CNS_CUTP_Lr) Or (CutType = CNS_CUTP_Lt) Or _
                           (CutType = CNS_CUTP_NLr) Or (CutType = CNS_CUTP_NLt) Then
                            ' ���H����5�̏����ԍ����J�b�g�f�[�^�̉��H����3���ݒ肷��
                            CndNum = stREG(Rn).STCUT(Cn).intCND(CUT_CND_L2)
                            If (stPWR.CndNumAry(CndNum) = 0) Then           ' ���� ? 
                                stPWR.CndNumAry(CndNum) = 1                 ' �L���ɐݒ�
                                ' �ڕW�p���[�l(W)�ƒ������e�͈�(�}W)��ݒ肷��
                                stPWR.AdjustTargetAry(CndNum) = stCND.dblPowerAdjustTarget(CndNum)
                                stPWR.AdjustLevelAry(CndNum) = stCND.dblPowerAdjustToleLevel(CndNum)
                            End If
                            ' ���H����6�̏����ԍ����J�b�g�f�[�^�̉��H����4���ݒ肷��
                            CndNum = stREG(Rn).STCUT(Cn).intCND(CUT_CND_L3)
                            If (stPWR.CndNumAry(CndNum) = 0) Then           ' ���� ? 
                                stPWR.CndNumAry(CndNum) = 1                 ' �L���ɐݒ�
                                ' �ڕW�p���[�l(W)�ƒ������e�͈�(�}W)��ݒ肷��
                                stPWR.AdjustTargetAry(CndNum) = stCND.dblPowerAdjustTarget(CndNum)
                                stPWR.AdjustLevelAry(CndNum) = stCND.dblPowerAdjustToleLevel(CndNum)
                            End If
                        End If

                        ' ���H����7,8�͌���͖��g�p(�\��)


                    Next Cn
                End If
            Next Rn

            Return (cFRS_NORMAL)                                        ' Return�l�ݒ� 

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
            Return (cERR_TRAP)                                          ' Return�l = �g���b�v�G���[����
        End Try
    End Function
#End If
#End Region

    '#Region "DispGazou.exe����"
    '=========================================================================
    '   �摜�\���v���O�����̋N������
    '=========================================================================

#Region "Dispgazou��Window���b�Z�[�W�𑗐M����"

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Public Function FindWindow( _
         ByVal lpClassName As String, _
         ByVal lpWindowName As String) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Public Function SendNotifyMessage( _
                           ByVal hWnd As IntPtr, _
                           ByVal wMsg As Int32, _
                           ByVal wParam As Int32, _
                           ByVal lParam As Int32) As Integer
    End Function

    Private Const WM_APP As Int32 = &H8000
    '    '''=========================================================================
    '    ''' <summary>Dispgazou��Window���b�Z�[�W�𑗐M����</summary>
    '    ''' <param name="No">(INP)���b�Z�[�W�ԍ�</param>
    '    ''' <returns></returns>
    '    ''' <remarks></remarks>
    '    '''=========================================================================
    '    Public Function SendMsgToDispGazou(ByRef ObjProc As Process, ByVal No As Integer) As Integer

    '        Dim result As Integer = cFRS_NORMAL
    '        Dim Cnt As Integer = 0
    '        Dim hWnd As Int32
    '        Try
    'SND_MSG_RETRY_START:
    '            '����̃E�B���h�E�n���h�����擾���܂�
    '            hWnd = FindWindow(Nothing, "DispGazou") 'V4.3.0.0�B
    '            If hWnd = 0 Then
    '                '�n���h�����擾�ł��Ȃ�����
    '                Cnt = Cnt + 1
    '                If Cnt < 100 Then
    '                    If (Cnt Mod 10) = 0 Then
    '                        Call FinalEnd_GazouProc(ObjProc)                                'DispGazou�����I��
    '                        'V2.2.0.0�@ Execute_GazouProc(ObjProc, DISPGAZOU_PATH, DISPGAZOU_WRK, INTERNAL_CAMERA)    '�ċN��
    '                    End If
    '                    System.Threading.Thread.Sleep(10)
    '                    GoTo SND_MSG_RETRY_START
    '                Else
    '                    MessageBox.Show("����Window�̃n���h�����擾�ł��܂���")
    '                End If
    '            End If

    '            result = SendNotifyMessage(hWnd, WM_APP, 0, No)

    '            Return result
    '        Catch ex As Exception
    '            MsgBox("UserModule.SendMsgToDispGazou() TRAP ERROR = " + ex.Message)
    '        End Try
    '    End Function
    '#End Region

    '#Region "�摜�\���v���O�������N������"
    '    '''=========================================================================
    '    ''' <summary>�摜�\���v���O�������N������</summary>
    '    ''' <param name="ObjProc"> (OUT)Process��޼ު��</param>
    '    ''' <param name="strFName">(INP)�N���v���O������</param>
    '    ''' <param name="Camera">  (INP)�J�����ԍ�(0-3)</param> 
    '    ''' <returns>0 = ����, 0�ȊO = �G���[</returns>
    '    '''=========================================================================
    '    Public Function Execute_GazouProc(ByRef ObjProc As Process, ByRef strFName As String, ByRef strWrk As String, ByVal Camera As Integer) As Integer

    '        Dim strARG As String                                        ' ����() 

    '        Dim dispXPos As Integer
    '        Dim dispYPos As Integer
    '        Dim Cnt As Integer = 0

    '        Try
    '            TrimClassCommon.ForceEndProcess(DISPGAZOU_PATH)       ' �v���Z�X�������I������B

    '            ' �\���ʒu�ݒ�
    '            dispXPos = FORM_X + Form1.VideoLibrary1.Location.X
    '            dispYPos = FORM_Y + Form1.VideoLibrary1.Location.Y

    '            ' �����ײ݈����ݒ�
    '            strARG = Camera.ToString("0") + " "                     ' args[0] :�J�����ԍ�(0-3)"
    '            'strARG = "0 "                                           ' args[0] :�J�����ԍ�(0-3)"
    '            strARG = strARG + "1 "                                  ' args[1] :(0=�{�^���\������, 1=�{�^���\�����Ȃ�)
    '            strARG = strARG + dispXPos.ToString("0") + " "          ' args[2] :�t�H�[���̕\���ʒuX
    '            strARG = strARG + dispYPos.ToString("0")                ' args[3] :�t�H�[���̕\���ʒuY
    '            strARG = strARG + " 1"                                  ' args[4] :(0=���b�Z�[�W���䖳��, 1=���b�Z�[�W����L��)
    '            strARG = strARG + " 1"                                  ' args[5] :(0=�V���v���g���}�p�T�C�Y�����, 1=�ʏ��ʃT�C�Y)

    '            ' �v���Z�X�̋N��
    '            ObjProc = New Process                                   ' Process��޼ު�Ă𐶐����� 
    '            ObjProc.StartInfo.FileName = strFName                   ' �v���Z�X�� 
    '            ObjProc.StartInfo.Arguments = strARG                    ' �����ײ݈����ݒ�
    '            ObjProc.StartInfo.WorkingDirectory = strWrk             ' ��ƃt�H���_
    '            ObjProc.Start()                                         ' �v���Z�X�N��

    '            ' �`���l����o�^
    '            'ChannelServices.RegisterChannel(ipcChnl, False)
    'IPC_RETRY_START:  ' �T�[�o�iDispGazou)�����~���Ē����ɋN�������ゾ�ƃ|�[�g�ɏ������߂Ȃ��G���[�ɂȂ�B�Ώ����@������Ȃ��̂ōĎ��s����B
    '            Try
    '                'refObj.CallServer("STOP")
    '                'V2.0.0.3�@                System.Threading.Thread.Sleep(2000)
    '                System.Threading.Thread.Sleep(500)
    '                SendMsgToDispGazou(ObjProc, 2)       'STOP 'V4.0.0.0-87
    '            Catch ex As Exception
    '                Cnt = Cnt + 1
    '                If Cnt < 100 Then
    '                    If (Cnt Mod 10) = 0 Then
    '                        Call FinalEnd_GazouProc(ObjProc)                                'DispGazou�����I��
    '                        'V2.2.0.0�@  Execute_GazouProc(ObjProc, DISPGAZOU_PATH, DISPGAZOU_WRK, Camera)    '�ċN��
    '                    End If
    '                    System.Threading.Thread.Sleep(10)
    '                    GoTo IPC_RETRY_START
    '                Else
    '                    MsgBox("UserModule.Execute_GazouProc() TRAP ERROR = " + ex.Message)
    '                End If
    '            End Try


    '            ' �g���b�v�G���[������ 
    '        Catch ex As Exception
    '            MsgBox("UserModule.Execute_GazouProc() TRAP ERROR = " + ex.Message)
    '            Return (cERR_TRAP)
    '        End Try
    '    End Function
    '#End Region

    '#Region "�摜�\���v���O�����������I������"
    '    '''=========================================================================
    '    '''<summary>�摜�\���v���O�����������I������</summary>
    '    '''<param name="ObjProc"> (OUT)Process��޼ު��</param>
    '    '''<returns>0 = ����, 0�ȊO = �G���[</returns>
    '    '''=========================================================================
    '    Public Function FinalEnd_GazouProc(ByRef ObjProc As Process) As Integer
    '        Try
    '            TrimClassCommon.ForceEndProcess(DISPGAZOU_PATH)       ' �_�������Ńv���Z�X�������I������B

    '            Return (cFRS_NORMAL)

    '            ' �g���b�v�G���[������ 
    '        Catch ex As Exception
    '            '            MsgBox("basTrimming.FinalEnd_GazouProc() TRAP ERROR = " + ex.Message)
    '            Return (cERR_TRAP)
    '        End Try
    '    End Function
    '#End Region

    '#Region "�摜�\���v���O�������N������"
    '    '''=========================================================================
    '    ''' <summary>
    '    ''' �摜�\���v���O�������N������
    '    ''' </summary>
    '    ''' <param name="ObjProc"> (OUT)Process��޼ު��</param>
    '    ''' <param name="strFName">(INP)�N���v���O������</param>
    '    ''' <param name="strWrk">(INP)��ƃt�H���_</param>
    '    ''' <param name="Camera">(INP)�J�����ԍ�(0-3)</param>
    '    ''' <returns></returns>
    '    ''' <remarks></remarks>
    '    '''=========================================================================
    '    Public Function Exec_GazouProc(ByRef ObjProc As Process, ByRef strFName As String, ByRef strWrk As String, ByVal Camera As Integer) As Integer

    '        Dim Cnt As Integer = 0
    '        ' Dim result As Integer

    '        Try
    '            'If Form1.GetDistributeOnOffStatus() Then
    '            '    Return (cFRS_NORMAL)
    '            'End If
    '            ' VideoOcx�\�����~
    '            Call Form1.VideoLibrary1.VideoStop()

    'IPC_RETRY_START:  ' �T�[�o�iDispGazou)�����~���Ē����ɋN�������ゾ�ƃ|�[�g�ɏ������߂Ȃ��G���[�ɂȂ�B�Ώ����@������Ȃ��̂ōĎ��s����B
    '            Try
    '                SendMsgToDispGazou(ObjProc, 5)                           ' START_NORMAL
    '            Catch ex As Exception
    '                Cnt = Cnt + 1
    '                If Cnt < 100 Then
    '                    If (Cnt Mod 10) = 0 Then
    '                        Call FinalEnd_GazouProc(ObjProc)                                'DispGazou�����I��
    '                        System.Threading.Thread.Sleep(100)
    '                        'V2.2.0.0�@                         Execute_GazouProc(ObjProc, strFName, strWrk, Camera)    '�ċN��
    '                    End If
    '                    System.Threading.Thread.Sleep(10)
    '                    GoTo IPC_RETRY_START
    '                Else
    '                    MsgBox("UserModule.Exec_GazouProc() TRAP ERROR = " + ex.Message)
    '                End If
    '            End Try

    '            Return (cFRS_NORMAL)
    '            ' �g���b�v�G���[������ 
    '        Catch ex As Exception
    '            MsgBox("UserModule.Exec_GazouProc() TRAP ERROR = " + ex.Message)
    '            Return (cERR_TRAP)
    '        End Try
    '    End Function
    '#End Region

    '#Region "�摜�\���v���O�����������I������"
    '    '''=========================================================================
    '    '''<summary>�摜�\���v���O�����������I������</summary>
    '    '''<param name="ObjProc"> (OUT)Process��޼ު��</param>
    '    '''<returns>0 = ����, 0�ȊO = �G���[</returns>
    '    '''=========================================================================
    '    Public Function End_GazouProc(ByRef ObjProc As Process) As Integer

    '        Dim Cnt As Integer = 0
    '        '        Dim result As Integer 

    '        Try

    'IPC_RETRY_START:  ' �T�[�o�iDispGazou)�����~���Ē����ɋN�������ゾ�ƃ|�[�g�ɏ������߂Ȃ��G���[�ɂȂ�B�Ώ����@������Ȃ��̂ōĎ��s����B
    '            Try

    '                SendMsgToDispGazou(ObjProc, 2)       'STOP 

    '            Catch ex As Exception
    '                Cnt = Cnt + 1
    '                If Cnt < 100 Then
    '                    If (Cnt Mod 10) = 0 Then
    '                        Call FinalEnd_GazouProc(ObjProc)                                'DispGazou�����I��
    '                        'V2.2.0.0�@                         Execute_GazouProc(ObjProc, DISPGAZOU_PATH, DISPGAZOU_WRK, INTERNAL_CAMERA)    '�ċN��
    '                    End If
    '                    System.Threading.Thread.Sleep(10)
    '                    GoTo IPC_RETRY_START
    '                Else
    '                    MsgBox("UserModule.End_GazouProc() TRAP ERROR = " + ex.Message)
    '                End If
    '            End Try

    '            Call Form1.VideoLibrary1.VideoStart()

    '            ' ��ʂ��X�V
    '            Call Form1.Refresh()

    '            Return (cFRS_NORMAL)

    '            ' �g���b�v�G���[������ 
    '        Catch ex As Exception
    '            MsgBox("UserModule.End_GazouProc() TRAP ERROR = " + ex.Message)
    '            Return (cERR_TRAP)
    '        End Try
    '    End Function
    '#End Region

#End Region

#Region "�r�[���|�W�V���i�̃G�C�W���O"
    ''' <summary>
    ''' �r�[���|�W�V���i�̃G�C�W���O�E�K���o���ő�܂œ������ăG�C�W���O����B
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub BeamPositionerAging()
        Try
            Dim r As Integer

            Call BSIZE(80, 80)                                                                      '   �u���b�N�T�C�Y�O�ݒ�
            r = ObjSys.EX_MOVE(gSysPrm, 0, 0, 1)                                                    '   BP(0,0)�ֈړ�
            If (r < cFRS_NORMAL) Then                                                               ' 
                Call Z_PRINT("�a�o�ړ��G���[���������܂����BEX_MOVE(gSysPrm, 0, 0, 1)" + vbCrLf)    ' 
            End If                                                                                  ' 
            r = ObjSys.EX_MOVE(gSysPrm, 80, 80, 1)                                                  '   BP(80,80)�ֈړ�
            If (r < cFRS_NORMAL) Then                                                               ' 
                Call Z_PRINT("�a�o�ړ��G���[���������܂����BEX_MOVE(gSysPrm, 80, 80, 1)" + vbCrLf)  ' 
            End If                                                                                  ' 
            r = ObjSys.EX_MOVE(gSysPrm, 40, 40, 1)                                                  '   BP(80,80)�ֈړ�
            If (r < cFRS_NORMAL) Then                                                               ' 
                Call Z_PRINT("�a�o�ړ��G���[���������܂����BEX_MOVE(gSysPrm, 40, 40, 1)" + vbCrLf)  ' 
            End If                                                                                  ' 
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub
#End Region

#Region "�u���b�N���̂Q�_�␳�ɂ��w�x�␳�̎��s��"
    Private bBlockXYCorrection As Boolean = False
    ''' <summary>
    ''' �V�X�p���̐ݒ肩��̃u���b�N���̂Q�_�␳�ɂ��w�x�␳�̗L��ݒ�
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SetBlockXYCorrectionOn()
        Try
            bBlockXYCorrection = True
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub
    ''' <summary>
    ''' �u���b�N���̂Q�_�␳�ɂ��w�x�␳���g�p�̗L�����擾
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsBlockXYCorrectionUse() As Boolean
        Try
            If Not bBlockXYCorrection Then
                Return (False)
            End If

            ' �ƕ␳����
            If (gSysPrm.stDEV.giTheta > 0) Then                         ' �w�x�ƗL��Ȃ���s���Ȃ�
                Return (False)
            End If

            Return (True)

        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
    ''' <summary>
    ''' �u���b�N���̂Q�_�␳�ɂ��w�x�␳�̎��s��
    ''' </summary>
    ''' <param name="AppMode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsBlockXYCorrection(ByRef AppMode As Short) As Boolean
        Try
            If Not bBlockXYCorrection Then
                Return (False)
            End If

            ' �ƕ␳����
            If (gSysPrm.stDEV.giTheta > 0) Then                         ' �w�x�ƗL��Ȃ���s���Ȃ�
                Return (False)
            End If

            If stThta.iPP31 = 0 Then                                    ' �␳�Ȃ��Ȃ���s���Ȃ�
                Return (False)
            End If

            If AppMode <> APP_MODE_TRIM Then                            ' �g���~���O���ȊO�͎��s���Ȃ�
                Return (False)
            End If

            Return (True)

        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region

#Region "�u���b�N���̂Q�_�␳�ɂ��w�x�␳"
    ''' <summary>
    ''' �u���b�N���̂Q�_�␳�ɂ��w�x�␳
    ''' </summary>
    ''' <param name="AppMode">�g���~���O���[�h</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function BlockXYCorrection(ByRef AppMode As Short) As Integer
        Dim rtn As Integer = cFRS_NORMAL                                ' Return�l
        Dim Thresh1 As Double = 0
        Dim Thresh2 As Double = 0
        Dim strMSG As String                                            ' Display Message
        Dim r As Integer = 0

        Try
            If Not IsBlockXYCorrection(AppMode) Then                    ' �u���b�N���̂Q�_�␳�ɂ��w�x�␳�̎��s��
                Return (cFRS_NORMAL)
            End If

            Call ObjSys.Ilum_Ctrl(gSysPrm, Z1, ZOPT)                    ' ���Ȱ������ߓ_��(����L���L��)

            Call InitThetaCorrection()                                  ' �p�^�[���o�^�����l�ݒ�

            ' �J�����ؑ�
            If (gSysPrm.stDEV.giCutPic = 0) Then                        ' VGA�{�[�h����
                ObjVdo.VideoStop()                                      ' �r�f�I�X�g�b�v/�X�^�[�g(����/�O�����)
                Call ObjVdo.VideoStart2(gSysPrm.stDEV.giEXCAM)
            Else
                If (gSysPrm.stDEV.giEXCAM_Usr = 1) Then                 ' �O���J�������g�p�H
                    Call ObjVdo.ChangeCamera(EXTERNAL_CAMERA)            ' �J�����ؑ�(�O�����)
                End If
            End If
#If cLEDcILLUMINATION Then
            If stUserData.iLEDIllumination = ELD_USE_ONLY Then          '�u�g�p���݂̂n�m�v
                UserSub.LEDLight_On()
            End If
#End If

            ' �ƕ␳����
            ObjVdo.frmTop = Form1.Text2.Location.Y                      ' �␳��ʕ\���ʒu�ݒ�
            ObjVdo.frmLeft = Form1.Text2.Location.X
            r = ObjVdo.CorrectTheta(APP_MODE_BLOCK_RECOG)               ' �ƕ␳

            ' XY�e�[�u���␳�l(�ƕ␳����XYð��ق����)�擾
            If (r = 0) Then
                dblCorrectX = ObjVdo.CorrectTrimPosX
                dblCorrectY = ObjVdo.CorrectTrimPosY
            Else
                dblCorrectX = 0
                dblCorrectY = 0
            End If

            ' �㏈��
            If (gSysPrm.stDEV.giCutPic = 0) Then                        ' VGA�{�[�h����?
                ObjVdo.VideoStop()                                      ' �r�f�I���C�u�����X�g�b�v
                ObjMain.Refresh()
            Else
                If (gSysPrm.stDEV.giEXCAM_Usr = 1) Then                 ' �O���J�������g�p�H
                    Call ObjVdo.ChangeCamera(INTERNAL_CAMERA)           ' �J�����ؑ�(�������)
                End If
            End If
#If cLEDcILLUMINATION Then
            If stUserData.iLEDIllumination = ELD_USE_ONLY Then          '�u�g�p���݂̂n�m�v
                UserSub.LEDLight_Off()
            End If
#End If
            If (r <> cFRS_NORMAL) Then                                  ' ERROR ?
                ' �p�^�[���F���G���[ ?
                If (r >= cFRS_MVC_10) And (r <= cFRS_VIDEO_PTN) Then
                    Call Beep()                                         ' Beep��
                    If (gSysPrm.stTMN.giMsgTyp = 0) Then
                        strMSG = "�p�^�[���F���G���[(" + r.ToString("0") + ")"
                    Else
                        strMSG = "VIDEOLIB: Pattern Matching Error"
                    End If
                    Call Z_PRINT(strMSG & vbCrLf)
                    rtn = cFRS_ERR_PTN                                  ' �p�^�[���F���G���[
                Else
                    rtn = r
                End If
                Call ObjSys.Ilum_Ctrl(gSysPrm, Z0, ZOPT)                ' ���Ȱ������ߏ���(����L���L��)
                Return (rtn)
            End If

            ' 臒l�擾
            Thresh1 = DllSysPrmSysParam_definst.GetPtnMatchThresh(stThta.iPP38, stThta.iPP37_1)
            Thresh2 = DllSysPrmSysParam_definst.GetPtnMatchThresh(stThta.iPP38, stThta.iPP37_2)

            ' �ƕ␳���ʎ擾
            Call ObjVdo.GetThetaResult(stResult)
            If (gSysPrm.stTMN.giMsgTyp = 0) Then
                strMSG = "�␳�ʒu1X,Y =" & stThta.fpp32_x.ToString("0.0000").PadLeft(9) & ","
                strMSG = strMSG & stThta.fpp32_y.ToString("0.0000").PadLeft(9) & "  "
                strMSG = strMSG & " �����1X,Y   =" & stResult.fCor1x.ToString("0.0000").PadLeft(9) & ","
                strMSG = strMSG & stResult.fCor1y.ToString("0.0000").PadLeft(9) & vbCrLf
                strMSG = strMSG & "�␳�ʒu2X,Y =" & stThta.fpp33_x.ToString("0.0000").PadLeft(9) & ","
                strMSG = strMSG & stThta.fpp33_y.ToString("0.0000").PadLeft(9) & "  "
                strMSG = strMSG & " �����2X,Y   =" & stResult.fCor2x.ToString("0.0000").PadLeft(9) & ","
                strMSG = strMSG & stResult.fCor2y.ToString("0.0000").PadLeft(9) & vbCrLf
                'If (stThta.iPP30 = 0) Then                              ' �����␳���[�h ?
                strMSG = strMSG & "  ��v�xPOS1   =" & stResult.fCorV1.ToString("0.0000").PadLeft(9) & ","
                strMSG = strMSG & " ��v�xPOS2   =" & stResult.fCorV2.ToString("0.0000").PadLeft(9) & vbCrLf
                'End If
            Else
                strMSG = "  Correct position1=" & stThta.fpp32_x.ToString("0.0000").PadLeft(9) & ","
                strMSG = strMSG & stThta.fpp32_y.ToString("0.0000").PadLeft(9) & "  "
                strMSG = strMSG & " Distance1=" & stResult.fCor1x.ToString("0.0000").PadLeft(9) & ","
                strMSG = strMSG & stResult.fCor1y.ToString("0.0000").PadLeft(9) & vbCrLf
                strMSG = strMSG & "  Correct position2=" & stThta.fpp33_x.ToString("0.0000").PadLeft(9) & ","
                strMSG = strMSG & stThta.fpp33_y.ToString("0.0000").PadLeft(9) & "  "
                strMSG = strMSG & " Distance2=" & stResult.fCor2x.ToString("0.0000").PadLeft(9) & ","
                strMSG = strMSG & stResult.fCor2y.ToString("0.0000").PadLeft(9) & vbCrLf
                'If (stThta.iPP30 = 0) Then                              ' �����␳���[�h ?
                strMSG = strMSG & "  Correlation coefficient1=" & stResult.fCorV1.ToString("0.0000").PadLeft(9) & ","
                strMSG = strMSG & " Correlation coefficient2=" & stResult.fCorV2.ToString("0.0000").PadLeft(9) & vbCrLf
                'End If
            End If

            ' �ƕ␳���\��
            Call Z_PRINT(strMSG)

            ' POS1��臒l�̃`�F�b�N���s��
            If (Thresh1 > stResult.fCorV1) Then
                Call Beep()                                             ' Beep��
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    strMSG = "�p�^�[���F���G���[ (POS1臒l)"
                Else
                    strMSG = "Pattern Matching Error(POS1 THRESH)"
                End If
                Call Z_PRINT(strMSG & vbCrLf)
                Call ObjSys.Ilum_Ctrl(gSysPrm, Z0, ZOPT)                ' ���Ȱ������ߏ���(����L���L��)
                Return (cFRS_ERR_PTN)                                   ' �p�^�[���F���G���[
            End If

            ' POS2��臒l�̃`�F�b�N���s��
            If (Thresh2 > stResult.fCorV2) Then
                Call Beep()                                             ' Beep��
                If (gSysPrm.stTMN.giMsgTyp = 0) Then
                    strMSG = "�p�^�[���F���G���[ (POS2臒l)"
                Else
                    strMSG = "Pattern Matching Error(POS2 THRESH)"
                End If
                Call Z_PRINT(strMSG & vbCrLf)
                Call ObjSys.Ilum_Ctrl(gSysPrm, Z0, ZOPT)                ' ���Ȱ������ߏ���(����L���L��)
                Return (cFRS_ERR_PTN)                                   ' �p�^�[���F���G���[
            End If
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region

#Region "���W�v�Z���W���[��"

    ''' <summary>
    ''' �Q�_�Ԃ̋��������߂�
    ''' </summary>
    ''' <param name="x"></param>
    ''' <param name="y"></param>
    ''' <param name="x2"></param>
    ''' <param name="y2"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDistance(ByRef x As Double, ByRef y As Double, ByRef x2 As Double, ByRef y2 As Double) As Double
        Try
            Dim distance As Double = Math.Sqrt((x2 - x) * (x2 - x) + (y2 - y) * (y2 - y))
            Return (distance)
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function

    ''' <summary>
    ''' �Q�_�Ԃ̊p�x�i���W�A���j�����߂�
    ''' </summary>
    ''' <param name="x"></param>
    ''' <param name="y"></param>
    ''' <param name="x2"></param>
    ''' <param name="y2"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRadian(ByRef x As Double, ByRef y As Double, ByRef x2 As Double, ByRef y2 As Double) As Double
        Try
            Dim radian As Double = Math.Atan2(y2 - y, x2 - x)
            Return (radian)
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function

    ''' <summary>
    ''' �Q�_�Ԃ̊p�x�i�x�j�����߂�
    ''' </summary>
    ''' <param name="x"></param>
    ''' <param name="y"></param>
    ''' <param name="x2"></param>
    ''' <param name="y2"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDegree(ByRef x As Double, ByRef y As Double, ByRef x2 As Double, ByRef y2 As Double) As Double
        Try
            Dim radian As Double = GetRadian(x, y, x2, y2)
            Dim degree As Double = radian * 180D / Math.PI
            Return (degree)
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function

    ''' <summary>
    ''' �p�x�Ƌ���������W�����߂�
    ''' </summary>
    ''' <param name="degree"></param>
    ''' <param name="distance"></param>
    ''' <param name="x"></param>
    ''' <param name="y"></param>
    ''' <remarks></remarks>
    Public Sub GetXYfromDegree(ByVal degree As Double, ByVal distance As Double, ByRef x As Double, ByRef y As Double)
        Try
            '��(degree)��Math.PI / 180���|���Ă���̂�degree��radian�ɕϊ����Ă���
            'radius�͔��a�ł���A�����Ƃ͔��a�����߂Ă���̂Ɠ���
            Dim radian As Double = degree * Math.PI / 180
            x = Math.Cos(radian) * distance
            y = Math.Sin(radian) * distance
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub

#Region "��]���W�����߂�"
    Public Sub XYRotation(ByVal x1 As Double, ByVal y1 As Double, ByVal angle As Double, ByRef x As Double, ByRef y As Double)
        Try
            Dim degrees As Double = Math.PI / 180 * angle
            x = x1 * Math.Cos(degrees) - y1 * Math.Sin(degrees)
            y = x1 * Math.Sin(degrees) + y1 * Math.Cos(degrees)
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub
#End Region

#End Region

#Region "�Q�_�̍��W�f�[�^����̂����␳"
    Sub HoseiCalc(ByVal bStageMode As Boolean, ByVal A1 As Double, ByVal B1 As Double, ByVal A2 As Double, ByVal B2 As Double, ByVal C1 As Double, ByVal D1 As Double, ByVal C2 As Double, ByVal D2 As Double, ByVal x1 As Double, ByVal y1 As Double, ByRef x As Double, ByRef y As Double)
        Try
            ' (x - a) / (a' - a) = (X - A)/(A'-A)
            ' (y - b) / (b' - b) = (Y - B)/(B'-B)
            ' x = a + (X - A) / (A'-A) * (a' - a)
            ' y = b + (Y - B) / (B'-B) * (b' - b)
            ' (X,Y) �� (x,y)
            ' (a,b) �� (A1,B1) (a',b') �� (A2,B2) 
            ' (A,B) �� (C1,D1) (A',B') �� (C2,D2) 
            ' ���̂Q�_ (A1,B1) (A2,B2)
            ' �p�^�[���F���œ���ꂽ�Q�_  (C1,D1) (C2,D2) 
            ' ���߂������W(x1,y1)
            ' �␳���ꂽ���W(x,y)

            'If (A1 = A2) Or (B1 = B2) Then
            '    ' �S�T�x��]�����Čv�Z���Ă��猳�ɖ߂�
            '    Dim RA1 As Double, RB1 As Double, RA2 As Double, RB2 As Double, RC1 As Double, RD1 As Double, RC2 As Double, RD2 As Double, Rx1 As Double, Ry1 As Double
            '    XYRotation(A1, B1, 45, RA1, RB1)
            '    XYRotation(A2, B2, 45, RA2, RB2)
            '    XYRotation(C1, D1, 45, RC1, RD1)
            '    XYRotation(C2, D2, 45, RC2, RD2)
            '    XYRotation(x1, y1, 45, Rx1, Ry1)
            '    x = (Rx1 - RA1) / (RA2 - RA1) * (RC2 - RC1) + RC1
            '    y = (Ry1 - RB1) / (RB2 - RB1) * (RD2 - RD1) + RD1
            '    XYRotation(x, y, -45, x, y)
            'Else
            '    x = (x1 - A1) / (A2 - A1) * (C2 - C1) + C1
            '    y = (y1 - B1) / (B2 - B1) * (D2 - D1) + D1
            'End If

            'x = Math.Round(x, 6)
            'y = Math.Round(y, 6)

            '�ŏ��̍��W�����_�Ɉړ�����B
            Dim X1O1 As Double = 0.0
            Dim Y1O1 As Double = 0.0
            Dim X2O1 As Double = A2 - A1
            Dim Y2O1 As Double = B2 - B1

            Dim X1O2 As Double = 0.0
            Dim Y1O2 As Double = 0.0
            Dim X2O2 As Double = C2 - C1
            Dim Y2O2 As Double = D2 - D1

            ' ��P���W�̂����
            'Dim diffx As Double = C1 - A1
            'Dim diffy As Double = D1 - B1

            ' ���W�̂���ʃZ���^�[
            Dim diffx2 As Double = C1 - A1
            Dim diffy2 As Double = D1 - B1
            Dim diffx As Double = (C1 + C2) / 2 - (A1 + A2) / 2
            Dim diffy As Double = (D1 + D2) / 2 - (B1 + B2) / 2

            '�����̔䗦�����߂�
            Dim distance1 As Double, distance2 As Double, Rate As Double, distance3 As Double
            distance1 = GetDistance(X1O1, Y1O1, X2O1, Y2O1)
            distance2 = GetDistance(X1O2, Y1O2, X2O2, Y2O2)
            distance3 = GetDistance(0, 0, x1, y1)
            Rate = distance2 / distance1

            '�p�x�����߂�
            Dim degree1 As Double, degree2 As Double, diffdegree As Double
            degree1 = GetDegree(X1O1, Y1O1, X2O1, Y2O1)
            degree2 = GetDegree(X1O2, Y1O2, X2O2, Y2O2)
            diffdegree = degree2 - degree1

            ' �����Ɗp�x������W�����߂�B
            Dim dX1 As Double, dY1 As Double, dX2 As Double, dY2 As Double
            GetXYfromDegree(diffdegree, distance3, dX2, dY2)

            XYRotation(x1, y1, diffdegree, dX1, dY1)    ' �Q�l�R�[�h

            If bStageMode Then
                'x = dX1 + diffx
                'y = dY1 + diffy
                x = dX1 + diffx2
                y = dY1 + diffy2
            Else
                x = dX1
                y = dY1
            End If

            Call DebugLogOut(String.Format("HoseiCalc (A1,B1)=({0},{1})(A2,B2)=({2},{3})(C1,D1)=({4},{5})(C2,D2)=({6},{7})(x1,x1)=({8},{9})(x,y)=({10},{11})", A1, B1, A2, B2, C1, D1, C2, D2, x1, y1, x, y))

        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub
#End Region

#Region "�Q�_�̍��W�f�[�^���狁�߂���W�����߂�"
    Sub HoseiCoordinate(ByVal x1 As Double, ByVal y1 As Double, ByRef x As Double, ByRef y As Double, Optional ByVal bStageMode As Boolean = True)
        Try
            HoseiCalc(bStageMode, stThta.fpp32_x, stThta.fpp32_y, stThta.fpp33_x, stThta.fpp33_y, stThta.fpp32_x + stResult.fCor1x, stThta.fpp32_y + stResult.fCor1y, stThta.fpp33_x + stResult.fCor2x, stThta.fpp33_y + stResult.fCor2y, x1, y1, x, y)
            Call DebugLogOut(String.Format("HoseiCoordinate (X,Y)=({0},{1})", x, y))
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub

#End Region

#Region "�Q�_�̍��W�f�[�^���狁�߂���W�̂���ʂ����߂�"
    Sub HoseiCoordinateDelta(ByVal x1 As Double, ByVal y1 As Double, ByRef x As Double, ByRef y As Double)
        Try
            Dim x2 As Double, y2 As Double
            HoseiCoordinate(x1, y1, x2, y2)
            x = x2 - x1
            y = y2 - y1
            Call DebugLogOut(String.Format("HoseiCoordinateDelta (��X,��Y)=({0},{1})", x, y))
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Sub

#End Region

#Region "�X�e�[�W�X�e�b�v�I�t�Z�b�g�����߂�"
    Sub GetTstepOffset(ByVal BlockX As Integer, ByVal BlockY As Integer, ByRef dlbOffX As Double, ByRef dlbOffY As Double)
        Try
            Dim dXpos As Double, dYpos As Double

            dXpos = stPLT.zsx * (BlockX - 1)
            dYpos = stPLT.zsy * (BlockY - 1)
            HoseiCoordinateDelta(dXpos, dYpos, dlbOffX, dlbOffY)
            Call DebugLogOut(String.Format("GetTstepOffset BLOCK({0},{1})(X,Y)=({2},{3})(��X,��Y)=({4},{5})", BlockX, BlockY, dXpos, dYpos, dlbOffX, dlbOffY))

        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try

    End Sub

#End Region

#Region "���b�Z�[�W�\���i�R���\�[���L�[�Ή��j"
    '''=========================================================================
    ''' <summary>FrmReset���g�p���Ďw��̃��b�Z�[�W��\������ ###089</summary>
    ''' <param name="ObjSys">(INP)OcxSystem�I�u�W�F�N�g</param>
    ''' <param name="gMode"> (INP)�������[�h</param>
    ''' <param name="Md">    (INP)cFRS_ERR_START                = START�L�[�����҂�
    '''                           cFRS_ERR_RST                  = RESET�L�[�����҂�
    '''                           cFRS_ERR_START + cFRS_ERR_RST = START/RESET�L�[�����҂�</param>
    ''' <param name="BtnDsp">(INP)�{�^���\������/���Ȃ�</param>
    ''' <param name="Msg1">  (INP)�\�����b�Z�[�W�P</param>
    ''' <param name="Msg2">  (INP)�\�����b�Z�[�W�Q</param>
    ''' <param name="MSG3">  (INP)�\�����b�Z�[�W�R</param>
    ''' <param name="Col1">  (INP)���b�Z�[�W�F�P</param>
    ''' <param name="Col2">  (INP)���b�Z�[�W�F�Q</param>
    ''' <param name="Col3">  (INP)���b�Z�[�W�F�R</param>
    ''' <returns>cFRS_ERR_START = OK�{�^��(START�L�[)����
    '''          cFRS_ERR_RST   = Cancel�{�^��(RESET�L�[)����
    '''          ��L�ȊO       = �G���[</returns> 
    '''=========================================================================
    Public Function FrmMessageDisp(ByVal ObjSys As Object, ByVal gMode As Integer, ByVal Md As Integer, ByVal BtnDsp As Boolean, _
                                       ByVal Msg1 As String, ByVal Msg2 As String, ByVal Msg3 As String, _
                                       ByVal Col1 As Object, ByVal Col2 As Object, ByVal Col3 As Object) As Integer

        Dim r As Integer
        Dim objForm As Object = Nothing
        Dim ColAry(3) As Object
        Dim MsgAry(3) As String

        Try
            ' �p�����[�^�ݒ�
            MsgAry(0) = Msg1
            MsgAry(1) = Msg2
            MsgAry(2) = Msg3
            ColAry(0) = Col1
            ColAry(1) = Col2
            ColAry(2) = Col3

            ' frmMessage��ʕ\��(�w��̃��b�Z�[�W��\������)
            objForm = New frmMessage()
            Call objForm.ShowDialog(Nothing, gMode, ObjSys, MsgAry, ColAry, Md, BtnDsp)
            r = objForm.sGetReturn()                                    ' Return�l�擾

            ' �I�u�W�F�N�g�J��
            If (objForm Is Nothing = False) Then
                Call objForm.Close()                                    ' �I�u�W�F�N�g�J��
                Call objForm.Dispose()                                  ' ���\�[�X�J��
            End If

            Return (r)                                                  ' Return(�G���[���̃��b�Z�[�W�͕\����) 

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            Call MsgBox_Exception(ex.Message, Form1)
        End Try
    End Function
#End Region


#Region "�J�����摜�\��PictureBox�N���b�N�ʒu��JOG�o�R�ŉ摜�Z���^�[�Ɉړ�����"
    ''' <summary>�J�����摜�\��PictureBox�N���b�N�ʒu��JOG�o�R�ŉ摜�Z���^�[�Ɉړ�����</summary>
    ''' <param name="distanceX"></param>
    ''' <param name="distanceY"></param>
    ''' <param name="stJOG">'V6.0.0.0-23</param>
    ''' <remarks>'V6.0.0.0�G</remarks>
    Public Sub MoveToCenter(ByVal distanceX As Decimal, ByVal distanceY As Decimal, ByRef stJOG As JOG_PARAM)
        stJOG.KeyDown = Keys.Execute                                    'V6.0.0.0-23
        InpKey = (InpKey Or CtrlJog.MouseClickLocation.GetInpKey(distanceX, distanceY))
    End Sub
#End Region


End Module

'=============================== END OF FILE ===============================
