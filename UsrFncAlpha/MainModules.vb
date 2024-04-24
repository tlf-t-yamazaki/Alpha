
Imports TrimClassLibrary                'V6.0.0.0�@  'V2.2.0.0�@

'V2.2.0.0�@ 'Public Class MainModules
Public Class MainModules
    Implements IMainModules           'V6.0.0.0�@    'V2.2.0.0�@

#Region "�A�v���P�[�V������ʂ�Ԃ�(OCX�p)"
    '''=========================================================================
    ''' <summary>�A�v���P�[�V������ʂ�Ԃ�</summary>
    ''' <param name="AppKind">9=���[�U�v��</param>
    '''=========================================================================
    Public Sub GetAppKind(ByRef AppKind As Short) Implements IMainModules.GetAppKind    'V6.0.0.0�@  'V2.2.0.0�@
        AppKind = KND_USER
    End Sub
#End Region

#Region "��R(�`�b�v)���ѕ�����Ԃ�(OCX�p)"
    '''=========================================================================
    ''' <summary>��R(�`�b�v)���ѕ�����Ԃ�</summary>
    ''' <param name="ResistDir">0=X����, 1=Y����</param>
    '''=========================================================================
    Public Sub GetResistDir(ByRef ResistDir As Short) Implements IMainModules.GetResistDir  'V6.0.0.0�@ 'V2.2.0.0�@
        ResistDir = 0
    End Sub
#End Region

#Region "�v���[�g���u���b�N��X�����AY�����̊J�n�ʒu�Z�o(OCX�p)"
    '''=========================================================================
    ''' <summary>�v���[�g���u���b�N��X�����AY�����̊J�n�ʒu�Z�o</summary>
    ''' <returns>0=����, 0�ȊO=�G���[</returns>
    '''=========================================================================
    Public Function Call_CalcBlockXYStartPos() As Integer Implements IMainModules.Call_CalcBlockXYStartPos  'V6.0.0.0�@  'V2.2.0.0�@

        Dim r As Integer = cFRS_NORMAL

        'r = CalcBlockXYStartPos()
        Return (r)

    End Function
#End Region

#Region "�w��u���b�NXY����X�e�[�W�ʒuXY���擾���e�[�u���ړ�����(OCX�p)"
    '''=========================================================================
    ''' <summary>�w��u���b�NXY����X�e�[�W�ʒuXY���擾���e�[�u���ړ�����</summary>
    ''' <param name="xBlockNo">(INP)�u���b�N�ԍ�X</param>
    ''' <param name="yBlockNo">(INP)�u���b�N�ԍ�Y</param>
    ''' <param name="OffSetX"> (INP)�I�t�Z�b�gX</param>
    ''' <param name="OffSetY"> (INP)�I�t�Z�b�gY</param>
    ''' <param name="stgx">    (OUT)�X�e�[�W�ʒuX</param>
    ''' <param name="stgy">    (OUT)�X�e�[�W�ʒuY</param>
    ''' <returns>0=����, 0�ȊO=�G���[</returns>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Function Call_GetTargetStagePosByXY(ByVal xBlockNo As Integer, ByVal yBlockNo As Integer,
                                               ByVal OffSetX As Double, ByVal OffSetY As Double,
                                               ByRef stgx As Double, ByRef stgy As Double) As Integer _
                                               Implements IMainModules.Call_GetTargetStagePosByXY       'V6.0.0.0�@

        Dim r As Integer

        If giAppMode = APP_MODE_PROBE And (stPLT.TeachBlockX > 1 Or stPLT.TeachBlockY > 1) Then ' ###1040�@ Move_Trimposition()�ňړ����Ă���̂ł����ł͈ړ����Ȃ��B
            Return (cFRS_NORMAL)                                                                ' ###1040�@
        End If                                                                                  ' ###1040�@

        ' XY�e�[�u���w��u���b�N�ړ�
        r = TSTEP(xBlockNo, yBlockNo, OffSetX, OffSetY)
        r = ObjSys.EX_ZGETSRVSIGNAL(gSysPrm, r, 0)

        Return (r)

    End Function
#End Region

#Region "(Teaching����)���C����ʏ�̃N���X���C���̕\���ʒu��ύX����"
    '''=========================================================================
    '''<summary>���C����ʏ�̃N���X���C���̕\���ʒu��ύX����</summary>
    '''=========================================================================
    Public Sub SetCrossLinePos(ByVal xPos As Integer, ByVal yPos As Integer)

        ' �N���X���C���ʒu��ݒ肷�� 
        'Form1.Picture1.Top = xPos + Form1.VideoLibrary1.Location.Y
        'Form1.Picture2.Left = yPos + Form1.VideoLibrary1.Location.X
        ' �N���X���C���ʒu��ݒ肷�� 
        ObjVdo.SetCorrCrossCenter(yPos, xPos)

        ' ��ʂ̍ĕ`��
        'Form1.Refresh()
    End Sub
#End Region

#Region "(Teaching����)�}�[�L���O�G���A�\��"
    '''=========================================================================
    '''<summary>���C����ʏ�̃}�[�L���O�G���A�̎l�p��\��/��\������</summary>
    '''=========================================================================
    Public Sub DisplayMarkingArea(ByVal bDisp As Boolean, ByVal xPos As Integer, ByVal yPos As Integer,
                                        ByVal width As Integer, ByVal height As Integer)

        ObjVdo.SetMarkingArea(bDisp, xPos, yPos, width, height)

    End Sub
#End Region

#Region "(Teaching-Jog����)�N���X���C���ʒu�ړ��\��"
    Public Sub DispCrossLine(ByVal xPos As Double, ByVal yPos As Double) Implements IMainModules.DispCrossLine
        ObjCrossLine.CrossLineDispXY(xPos, yPos)
        ''�N���X���C���␳�������Ăяo��
        'gstCLC.x = xPos                    ' BP�ʒuX(mm)
        'gstCLC.y = yPos                    ' BP�ʒuY(mm)
        'Call CrossLineCorrect(gstCLC)       ' �␳�N���X���C���\��
    End Sub
#End Region

#Region "�␳�N���X���C���ʒu�擾"
    Public Sub GetCorrCrossLinePixel(ByVal bpx As Double, ByVal bpy As Double, ByRef xPos As Integer, ByRef yPos As Integer) Implements IMainModules.GetCorrCrossLinePixel

        ObjCrossLine.GetCorrCrossPixel(bpx, bpy, xPos, yPos)

    End Sub

#End Region
#Region "�r�f�I�̃X�^�[�g�X�g�b�v����"
    '===========================================================================
    ' �r�f�I�X�^�[�g��~����
    '===========================================================================
    ''' <summary>
    ''' �r�f�I�̃X�^�[�g����
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub VideoStart() Implements IMainModules.VideoStart  'V6.0.0.0�@ 'V2.2.0.0�@
        Try
            ''Call ObjVdo.VideoStart()
            Call ObjVdo.VideoStart()        ''V2.2.0.0�M
        Catch ex As Exception
            MsgBox("MainModules.VideoStart() TRAP ERROR = " + ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' �r�f�I�̃X�g�b�v����
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub VideoStop() Implements IMainModules.VideoStop  'V6.0.0.0�@ 'V2.2.0.0�@
        Try
            ''Call ObjVdo.VideoStop()
            Call ObjVdo.VideoStop() 'V2.2.0.0�M
        Catch ex As Exception
            MsgBox("MainModules.VideoStop() TRAP ERROR = " + ex.Message)
        End Try
    End Sub
#End Region

#Region "I/F���킹�邽�߂̃_�~�[�֐�"   'V2.2.0.0�@
    Public Sub Call_SetAlmStartTime(strDAT As String, ErrCode As Short) Implements IMainModules.Call_SetAlmStartTime
        Throw New NotImplementedException()
    End Sub

    Public Sub Call_GetVacumeStatus(ByRef Sts As Integer) Implements IMainModules.Call_GetVacumeStatus
        '@@@888
        ' �z���Z���T�̏�Ԃ��擾����֐��Ăяo�� 
        Sts = 1
        '@@@888        Throw New NotImplementedException()
    End Sub

    Public Sub Call_SetAlmEndTime() Implements IMainModules.Call_SetAlmEndTime
        Throw New NotImplementedException()
    End Sub

    Public Sub CrossLineDispOff() Implements IMainModules.CrossLineDispOff
        ObjVdo.SetCorrCrossVisible(False)
    End Sub

    Public Sub SetVideoTrackBar(visible As Boolean, enabled As Boolean) Implements IMainModules.SetVideoTrackBar
        ObjVdo.SetTrackBar(visible, enabled)
    End Sub

    Public Sub SetActiveJogMethod(keyDown As Action(Of KeyEventArgs),
                                  keyUp As Action(Of KeyEventArgs),
                                  moveToCenter As Action(Of Decimal, Decimal)) Implements IMainModules.SetActiveJogMethod
        ObjMain.SetActiveJogMethod(keyDown, keyUp, moveToCenter)

    End Sub

    Public Function Call_Loader_AlarmCheck_ManualMode() As Integer Implements IMainModules.Call_Loader_AlarmCheck_ManualMode
        Throw New NotImplementedException()
    End Function

    Public Sub sub_MGStopJog() Implements IMainModules.sub_MGStopJog
        ''V2.2.0.0�D        Throw New NotImplementedException()
    End Sub

    Public Sub SetAPBPrm_Step(StepCount As Short, ProbePitch As Double, StepCount2 As Short, ProbePitch2 As Double) Implements IMainModules.SetAPBPrm_Step
        Throw New NotImplementedException()
    End Sub

    Public Function Sub_CallFrmMatrix() As Integer Implements IMainModules.Sub_CallFrmMatrix
        Throw New NotImplementedException()
    End Function

    Public Function Sub_ProbeCheck(Plt As Integer, BlkX As Integer, BlkY As Integer, Limit As Double, ByRef strLOG As String, ByRef strDSP As String) As Integer Implements IMainModules.Sub_ProbeCheck
        Throw New NotImplementedException()
    End Function

    Public Function CalcStartPointPlate(ByRef OrgStartPoint(,) As Double, ByRef StartPoint(,) As Double, ByVal PlateX As Integer, ByVal PlateY As Integer) As Boolean Implements IMainModules.CalcStartPointPlate
        Throw New NotImplementedException()
    End Function

    Public Function Call_StagePosByPltBlkNo(ByVal xPlateNo As Integer, ByVal yPlateNo As Integer, ByVal xBlockNo As Integer, ByVal yBlockNo As Integer,
                                            ByVal stgOffSetX As Double, ByVal stgOffSetY As Double, ByRef stgx As Double, ByRef stgy As Double) As Integer _
     Implements IMainModules.Call_StagePosByPltBlkNo
        Throw New NotImplementedException()
    End Function

    Public Function GetPlateBlockNumber(ByVal plateX As Integer, ByVal plateY As Integer, ByVal blockX As Integer, ByVal blockY As Integer, ByRef plateNumber As Integer, ByRef blockNumber As Integer) As Boolean _
    Implements IMainModules.GetPlateBlockNumber
        Throw New NotImplementedException()
    End Function

    Public Function GetPlateBPOffset(ByVal plateX As Integer, ByVal plateY As Integer, ByRef OffsetX As Double, ByRef OffsetY As Double) As Boolean _
    Implements IMainModules.GetPlateBPOffset
        Throw New NotImplementedException()
    End Function

    Public Function GetPlateStepOffset(ByVal plateX As Integer, ByVal plateY As Integer, ByRef OffsetX As Double, ByRef OffsetY As Double) As Boolean _
    Implements IMainModules.GetPlateStepOffset
        Throw New NotImplementedException()
    End Function

    Public Function GetPlateTableOffset(ByVal plateX As Integer, ByVal plateY As Integer, ByRef OffsetX As Double, ByRef OffsetY As Double) As Boolean _
    Implements IMainModules.GetPlateTableOffset
        Throw New NotImplementedException()
    End Function

    Public Function GetPlateTXOffset(ByVal plateX As Integer, ByVal plateY As Integer, ByRef OffsetX As Double, ByRef OffsetY As Double) As Boolean _
     Implements IMainModules.GetPlateTXOffset
        Throw New NotImplementedException()
    End Function

    Public Function ReverseCalcStartPointPlate(ByRef OrgStartPoint(,) As Double, ByRef StartPoint(,) As Double, ByVal PlateX As Integer, ByVal PlateY As Integer) As Boolean _
        Implements IMainModules.ReverseCalcStartPointPlate
        Throw New NotImplementedException()
    End Function

    Public Function SetPlateBPOffset(ByVal plateX As Integer, ByVal plateY As Integer, ByVal OffsetX As Double, ByVal OffsetY As Double) As Boolean _
        Implements IMainModules.SetPlateBPOffset
        Throw New NotImplementedException()
    End Function

    Public Function SetPlateOffsetData(ByVal plateX As Integer, ByVal plateY As Integer) As Boolean _
     Implements IMainModules.SetPlateOffsetData
        Throw New NotImplementedException()
    End Function

    Public Function SetStartPointPlate(ByRef OrgStartPoint(,) As Double, ByRef StartPoint(,) As Double, ByVal PlateX As Integer, ByVal PlateY As Integer) As Boolean _
    Implements IMainModules.SetStartPointPlate
        Throw New NotImplementedException()
    End Function

    Public Function SetTrimParameterPlate(ByVal ChipSize(,) As Double, ByVal PlateNo As Integer) As Boolean _
             Implements IMainModules.SetTrimParameterPlate
        Throw New NotImplementedException()
    End Function

    Public Function SetTrimParamToGlobalAreaPlate(ByVal ChipSize(,) As Double, ByVal PlateNo As Integer) As Boolean _
     Implements IMainModules.SetTrimParamToGlobalAreaPlate
        Throw New NotImplementedException()
    End Function

    Public Function MainGpibUnitOff() As Integer _
        Implements IMainModules.MainGpibUnitOff
        'Throw New NotImplementedException()                            
    End Function
    Public Function Call_GetCHTheta(Channel As String, ByRef angle As Double) As Integer _
        Implements IMainModules.Call_GetCHTheta
        Throw New NotImplementedException()
    End Function

#Region "�ƕ␳"
    '''=========================================================================
    '''<summary>�ƕ␳</summary>
    '''=========================================================================
    Public Function Call_DoCorrectPos(dispCurPltNoX As Integer, dispCurPltNoY As Integer,
                                 Optional ByVal posz As Double = -9999,
                                 Optional ByVal highsensortpos As Double = -9999) As Integer Implements IMainModules.Call_DoCorrectPos

    End Function
#End Region

    Public Sub CrossLineDispOn() Implements IMainModules.CrossLineDispOn      'V6.0.0.0�@

    End Sub


    '''=========================================================================
    ''' <summary>�N���X���C��OFFSET�ݒ�</summary>
    ''' <param name="xOffset">(INP)OFFSET X�ʒu(pixel)</param>
    ''' <param name="yOffset">(INP)OFFSET Y�ʒu(pixel)</param>
    '''=========================================================================
    Public Sub DispCrossOffset(ByVal xOffset As Integer, ByVal yOffset As Integer) _
        Implements IMainModules.DispCrossOffset                           'V6.0.0.0�@

    End Sub

    ''' <summary>
    ''' ���[�_���_���A�̌Ăі߂��p 
    ''' </summary>
    ''' <param name="mode"></param>
    Public Function Call_Sub_Loader_OrgBack(ByVal mode As Integer) As Integer _
        Implements IMainModules.Call_Sub_Loader_OrgBack                           'V6.0.0.0�@

        Dim r As Integer

        Try

            r = ObjLoader.Sub_Loader_OrgBack(mode)

            Call_Sub_Loader_OrgBack = r
        Catch ex As Exception

        End Try

    End Function


    ''' <summary>
    ''' ���[�_���_���A�̊�����҂�       'V2.2.1.1�F
    ''' </summary>
    ''' <returns></returns>
    Public Function Call_Sub_WaitLoaderOrigin() As Integer _
                Implements IMainModules.Call_Sub_WaitLoaderOrigin
        Dim r As Integer = cFRS_NORMAL

        Try

            r = ObjLoader.WaitLoaderOrigin()

            Call_Sub_WaitLoaderOrigin = r
            Return r

        Catch ex As Exception

        End Try

    End Function

    ''' <summary>
    ''' ���b�g�؂�ւ��M���̐ݒ���s��        'V2.2.1.1�F
    ''' </summary>
    ''' <param name="count">���b�g�؂�ւ���</param>
    ''' <returns></returns>
    Public Function Call_Sub_LotChangeFlgSet(ByVal count As Integer) As Integer _
                Implements IMainModules.Call_Sub_LotChangeFlgSet
        Dim r As Integer = cFRS_NORMAL

        Try

            r = ObjLoader.SetLotChangeFlg(count)

            Return r

        Catch ex As Exception

        End Try

    End Function


    ''' <summary>
    ''' ���b�g�؂�ւ��M���̐ݒ���s��        'V2.2.1.1�F
    ''' </summary>
    ''' <param name="count">���b�g�؂�ւ���</param>
    ''' <returns></returns>
    Public Function Call_Sub_LotChangeFlgGet() As Integer _
                Implements IMainModules.Call_Sub_LotChangeFlgGet

        Dim r As Integer = cFRS_NORMAL

        Try

            r = ObjLoader.GetLotChangeFlg()

            Return r

        Catch ex As Exception

        End Try

    End Function


#End Region

End Class

