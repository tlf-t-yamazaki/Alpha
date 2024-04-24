'===============================================================================
'   Description  : �f�[�^�I����ʏ���(�����^�]�p)
'
'   Copyright(C) : TOWA LASERFRONT CORP. 2011
'
'===============================================================================
Imports System.IO

Public Class FormDataSelect
#Region "�y�ϐ���`�z"
    '===========================================================================
    '   �ϐ���`
    '===========================================================================
    Public giTrimNGMode As Integer                      ' �g���~���O�s�ǐM����A�������^�]���~�ʒm�Ɏg�p(0:�W��(�@�\�Ȃ�), 1:�@�\����)�@tky.ini "DEVICE_CONST", "AUTO_OPERATION_TRM_NG"

    Private Const DATA_DIR_PATH As String = "C:\TRIMDATA\DATA"          ' �f�[�^�t�@�C���t�H���_(����l)
    'Private Const DATA_ENTRY_PATH As String = DATA_DIR_PATH & "\ENTRYLOT"   ' �o�^�ς��ް�̧��̫���
    Private Const ENTRY_PATH As String = "C:\TRIMDATA\ENTRYLOT\"
    Private Const ENTRY_TMP_FILE As String = "SAVE_ENTRY.TMP"


    '----- �A���^�]�p(SL436R�p) -----
    Public gbFgAutoOperation As Boolean                     ' �����^�]�t���O(True:�����^�]��, False:�����^�]���łȂ�) 
    Public gsAutoDataFileFullPath() As String               ' �A���^�]�o�^�f�[�^�t�@�C�����z��
    Public giAutoDataFileNum As Short                       ' �A���^�]�o�^�f�[�^�t�@�C����
    'Public Const MODE_MAGAZINE As Short = 0                 ' �}�K�W�����[�h
    'Public Const MODE_LOT As Short = 1                      ' ���b�g���[�h
    'Public Const MODE_ENDLESS As Short = 2                  ' �G���h���X���[�h

    '----- �ϐ���` -----
    Private mExitFlag As Integer                            ' �I���t���O
    Private m_mainEdit As Form1                             ' Ҳ݉�ʂւ̎Q��
    '    Private gsAutoDataFileFullPath() As String         ' �A���^�]�o�^ؽ����߽������z��
    Private sLogFileName As String
    Private sPlateDataFileName As String
    'Private sSaveL7 As String, sSaveL4 As String, sSaveL14 As String, sSaveL21 As String, sSaveL22 As String, sSaveL17 As String, sSaveL5 As String, sSaveL15 As String, sSaveL23 As String, sSaveL6 As String, sSaveL29 As String, sSaveL30 As String, sSaveL31 As String, sSaveL32 As String, sSaveL33 As String
    Private InitiallNgCount As Long
    Private NowlNgCount As Long
    Private bLotChange As Boolean
    Private NowExecuteLotNo As Integer
    Private AutoOpeCancel As Boolean
    Private CancelReason As Integer
    Private Const NO_MORE_ENTRY As Integer = 1
#End Region

#Region "�ݽ�׸�"
    Friend Sub New(ByRef mainEdit As Form1)

        ' ���̌Ăяo���́AWindows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
        InitializeComponent()

        ' InitializeComponent() �Ăяo���̌�ŏ�������ǉ����܂��B
        m_mainEdit = mainEdit ' Ҳ݉�ʂւ̎Q�Ƃ�ݒ�

    End Sub
#End Region

#Region "�y���\�b�h��`�z"
#Region "�I�����ʂ�Ԃ�"
    '''=========================================================================
    ''' <summary>�I�����ʂ�Ԃ�</summary>
    ''' <returns>cFRS_ERR_START = OK�{�^������
    '''          cFRS_ERR_RST   = Cancel�{�^������</returns>
    '''=========================================================================
    Public Function sGetReturn() As Integer

        Dim strMSG As String

        Try
            Return (mExitFlag)

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "FormDataSelect.sGetReturn() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Function
#End Region

#Region "ShowDialog���\�b�h�ɓƎ��̈�����ǉ�����"
    '''=========================================================================
    ''' <summary>ShowDialog���\�b�h�ɓƎ��̈�����ǉ�����</summary>
    ''' <param name="Owner">(INP)���g�p</param>
    ''' <remarks></remarks>
    '''=========================================================================
    Public Overloads Sub ShowDialog(ByVal Owner As IWin32Window)

        Dim strMSG As String

        Try
            ' ��������
            mExitFlag = -1                                              ' �I���t���O = ������

            ' ��ʕ\��
            Me.ShowDialog()                                             ' ��ʕ\��
            Exit Sub

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "FormDataSelect.ShowDialog() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "Form_Load������"
    '''=========================================================================
    ''' <summary>Form_Load������</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub FormDataSelect_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim strMSG As String

        Try
            ' �t�H�[���� 
            Me.Text = MSG_AUTO_14                                       ' "�f�[�^�o�^"

            ' ���x�����E�{�^���� 
            LblDataFile.Text = MSG_AUTO_05                              ' "�f�[�^�t�@�C��"
            LblListList.Text = MSG_AUTO_06                              ' "�o�^�ς݃f�[�^�t�@�C��"
            BtnUp.Text = MSG_AUTO_07                                    ' "���X�g��1���"
            BtnDown.Text = MSG_AUTO_08                                  ' "���X�g��1����"
            BtnDelete.Text = MSG_AUTO_09                                ' "���X�g����폜"
            BtnClear.Text = MSG_AUTO_10                                 ' "���X�g���N���A"
            BtnSelect.Text = MSG_AUTO_11                                ' "���o�^��"
            BtnOK.Text = MSG_AUTO_12                                    ' "OK"
            BtnCancel.Text = MSG_AUTO_13                                ' "�L�����Z��"

            ' ���X�g�{�b�N�X�N���A
            Call ListList.Items.Clear()                                 ' �u�o�^�ς݃f�[�^�t�@�C���v���X�g�{�b�N�X�N���A

            ' �u�f�[�^�t�@�C���v���X�g�{�b�N�X�ɓ��t�t���t�@�C������\������
            DrvListBox.Drive = "C:"                                    ' �h���C�u 
            DirListBox.Path = DATA_DIR_PATH                             ' �f�B���N�g�����X�g�{�b�N�X����l
            MakeFileList()                                              ' ���ʏ��DirListBox_Change()�C�x���g����������̂ŕs�v����
            '                                                           ' �J�����g��"C:\TRIMDATA\DATA"���Ɣ������Ȃ��̂ŕK�v

            ' �o�^�ς��ް�̧��̫��ނ̗L�����m�F����
            If (False = Directory.Exists(ENTRY_PATH)) Then
                Directory.CreateDirectory(ENTRY_PATH)              ' ̫��ނ����݂��Ȃ���΍쐬����
            End If

            Call LoadPlateDataFileFullPath()

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "FormDataSelect.FormDataSelect_Load() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

    '========================================================================================
    '   �{�^���������̏���
    '========================================================================================
#Region "�ް��ݒ����݁E�ҏW���݉���������"
    ''' <summary>�ް��ݒ�����</summary>
    Private Sub cmdLotInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLotInfo.Click
        Call LoadAndEditData(0)

    End Sub

    ''' <summary>�ҏW����</summary>
    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click
        ' �p�X���[�h����

        If (Func_Password(F_EDIT) <> True) Then         ' �p�X���[�h���ʹװ�Ȃ�EXIT
            Return
        End If

        Call LoadAndEditData(1)

            MakeFileList()              ' DirListBox�őI������Ă���̫��ނɕۑ����ꂽ�ꍇ�AؽĂ̍X�V���K�v   'V4.7.0.0�L
    End Sub

    ''' <summary>�I�𒆂̓o�^�ς��ް�̧�ق�ǂݍ�����ް��ݒ��ʂ܂��͕ҏW��ʂ��J��</summary>
    ''' <param name="button">0=�ް��ݒ�����,1=�ҏW����</param>
    Private Sub LoadAndEditData(ByVal button As Integer)
        Dim rslt As Short
        Dim s As String
        Dim r As Short

        ' �o�^�ς݂��ް�̧�ق��Ȃ����NOP
        If (ListList.Items.Count < 1) Then Exit Sub
        Try
            '-----------------------------------------------------------------------
            '   ��������
            '-----------------------------------------------------------------------
            giAppMode = APP_MODE_LOAD                       ' ����Ӱ�� = �t�@�C�����[�h(F1)

            ' �p�X���[�h����(�I�v�V����)
            rslt = Func_Password(F_LOAD)
            If (rslt <> True) Then
                Exit Try                                    ' �߽ܰ�ޓ��ʹװ�Ȃ�EXIT
            End If

            ' �ް�̧�ٖ��ݒ�
            With ListList
                gsDataFileName = (ENTRY_PATH & .Items(.SelectedIndex))
            End With

            ' ���ݒ�̑��u�̓d����OFF����
            r = V_Off()                                     ' DC�d�����u �d��OFF����

            ' �g���~���O�f�[�^�ݒ�
            r = UserVal()                                   ' �f�[�^�����ݒ�
            If (r <> 0) Then                                ' �G���[ ?
                pbLoadFlg = False                           ' �f�[�^���[�h�σt���O = False
                s = "Data load Error : " & gsDataFileName & vbCrLf
                Me.LblFullPath.Text = s
                Call Z_PRINT(s)
            Else
                Call Z_CLS()                                ' �f�[�^���[�h�Ń��O��ʃN���A              ###lstLog
                gDspCounter = 0                             ' ���O��ʕ\��������J�E���^�N���A
                pbLoadFlg = True                            ' �f�[�^���[�h�σt���O = True
                s = "Data loaded : " & gsDataFileName & vbCrLf
                Call Z_PRINT(s)

                Call m_mainEdit.System1.OperationLogging( _
                        gSysPrm, MSG_OPLOG_FUNC01, "File='" & gsDataFileName & "' MANUAL")

                'V2.1.0.0�C��
                If gsDataFileName.Length > 60 Then
                    m_mainEdit.LblDataFileName.Text = gsDataFileName
                Else
                    'V2.1.0.0�C��
                    ' �t�@�C���p�X���̕\��
                    If (gSysPrm.stTMN.giMsgTyp = 0) Then
                        m_mainEdit.LblDataFileName.Text = "�f�[�^�t�@�C���� " & gsDataFileName
                    Else
                        m_mainEdit.LblDataFileName.Text = "File name " & gsDataFileName
                    End If
                End If          'V2.1.0.0�C
                '-----------------------------------------------------------------------
                '   FL���։��H�����𑗐M����(FL���ŉ��H�����t�@�C��������ꍇ)
                '-----------------------------------------------------------------------
                Call m_mainEdit.SendFlParam(gsDataFileName)

                '###1040�E                Call m_mainEdit.SetATTRateToScreen(True)    ' ###1040�B �A�b�e�l�[�^�̐ݒ�
            End If

            '-----------------------------------------------------------------------
            '   ۰�ޏI������
            '-----------------------------------------------------------------------
            ChDrive("C")                                    ' ChDrive���Ȃ��Ǝ��N����FD�h���C�u�����ɍs����,
            ChDir(My.Application.Info.DirectoryPath)        ' "MVCutil.dll���Ȃ�"�ƂȂ�N���ł��Ȃ��Ȃ�

            ' ======================================================================
            '   �ް��ݒ��ʁE�ҏW��ʌĂяo��
            ' ======================================================================
            ' �ް�۰������ (���ݸ��ް������ݒ�:UserVal() �̴װ����)
            If pbLoadFlg = False Then
                s = MSG_DataNotLoad                         ' �ް���۰��
                Call Z_PRINT(s)
                Call Beep()
                Exit Try
            End If

            If (0 = button) Then
                ' �ް��ݒ���
                giAppMode = APP_MODE_LOTNO                  ' ����Ӱ�� = ���b�g�ԍ��ݒ蒆
                ' �f�[�^�ҏW
                Call m_mainEdit.System1.OperationLogging(gSysPrm, MSG_OPLOG_LOTSET, "")

                Dim fLotInf As New FormEdit.frmLotInfoInput()
                fLotInf.ShowDialog(Me)
                fLotInf.Dispose()
            Else
                ' �ҏW���
                giAppMode = APP_MODE_EDIT                   ' ����Ӱ�� = �ҏW��ʕ\��
                Call m_mainEdit.System1.OperationLogging(gSysPrm, MSG_OPLOG_FUNC03, "")

                FlgUpdGPIB = 0                              ' GPIB�f�[�^�X�VFlag Off
                Dim fForm As New FormEdit.frmEdit           ' frm��޼ު�Đ���
                fForm.ShowDialog()                          ' �f�[�^�ҏW
                fForm.Dispose()                             ' frm��޼ު�ĊJ��

                ' GPIB�f�[�^�X�V�Ȃ�GPIB���������s��
                If (FlgUpdGPIB = 1) Then
                    Call GPIB_Init()
                End If
            End If

            If (True = FlgUpd) Then
                '-----------------------------------------------------------------------
                '   �f�[�^�t�@�C�����Z�[�u����
                '-----------------------------------------------------------------------
                If rData_save(gsDataFileName) <> 0 Then       ' �f�[�^�t�@�C���Z�[�u
                    Exit Try
                Else
                    Call Z_PRINT("Data saved : " & gsDataFileName & vbCrLf)
                End If

                '-----------------------------------------------------------------------
                '   ���샍�O�����o�͂���
                '-----------------------------------------------------------------------
                Call m_mainEdit.System1.OperationLogging( _
                    gSysPrm, MSG_OPLOG_FUNC02, "File='" & gsDataFileName & "' MANUAL")

                FlgUpd = Convert.ToInt16(TriState.False)    ' �f�[�^�X�V Flag OFF
            End If

            ChDrive("C")                                    ' ChDrive���Ȃ��Ǝ��N����FD�h���C�u�����ɍs����,"MVCutil.dll���Ȃ�"�ƂȂ�N���ł��Ȃ��Ȃ�
            ChDir(My.Application.Info.DirectoryPath)

            ' �g���b�v�G���[������
        Catch ex As Exception
            MsgBox("LoadAndEditData() TRAP ERROR = " + ex.Message)
        Finally
            Call ZCONRST()                                  ' �ݿ�ٷ� ׯ�����
            giAppMode = APP_MODE_LOTCHG                     ' ����Ӱ�� = ���b�g�ؑ�
        End Try

    End Sub
#End Region

#Region "OK�{�^������������"
    '''=========================================================================
    ''' <summary>OK�{�^������������</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOK.Click

        Dim Idx As Integer
        Dim strMSG As String = ""
        Dim sRtn As Short

        Try
            gbFgAutoOperation = False

            ' �I�����X�g1�ȏ�L�肩�`�F�b�N���� ?
            If (ListList.Items.Count < 1) Then
                '"�f�[�^�t�@�C����I�����Ă��������B"
                Call MsgBox(MSG_AUTO_18, MsgBoxStyle.OkOnly)
                Exit Sub
            End If

#If cOSCILLATORcFLcUSE Then
        Dim r As Integer
        Dim strDAT As String
            ' �I���f�[�^�ɑΉ�������H�����t�@�C�������݂��邩�`�F�b�N����(FL��)
            If (gSysPrm.stRAT.giOsc_Res = OSCILLATOR_FL) Then
                For Idx = 0 To ListList.Items.Count - 1
                    strDAT = (ENTRY_PATH & ListList.Items(Idx))
                    r = GetFLCndFileName(strDAT, strMSG, True)              ' ���݃`�F�b�N 
                    If (r <> SerialErrorCode.rRS_OK) Then                   ' ���H�����t�@�C�������݂��Ȃ� ?
                        ' "���H�����t�@�C�������݂��܂���B(���H�����t�@�C����)"
                        strMSG = MSG_AUTO_20 + "(" + strMSG + ")"
                        Call MsgBox(strMSG, MsgBoxStyle.OkOnly, "")
                        ListList.SelectedIndex = Idx
                        Call ListList_SelectedIndexChanged(sender, e)                      ' �f�[�^�t�@�C�������t���p�X�Ń��x���e�L�X�g�{�b�N�X�ɐݒ肷��
                        Exit Sub
                    End If
                Next Idx
            End If
#End If

            ' �A���^�]�p�̃f�[�^�t�@�C�����ƃf�[�^�t�@�C�����z����O���[�o���̈�ɐݒ肷��
            giAutoDataFileNum = ListList.Items.Count                    ' �f�[�^�t�@�C����
            ReDim gsAutoDataFileFullPath(giAutoDataFileNum - 1)
            For Idx = 0 To giAutoDataFileNum - 1                        ' �f�[�^�t�@�C����
                gsAutoDataFileFullPath(Idx) = (ENTRY_PATH & ListList.Items(Idx))
            Next

            If OffSetCheckBox.Checked = True Then
                sRtn = UserSub.SetOffSetDataToAutoOperationData(gsAutoDataFileFullPath, giAutoDataFileNum)
                If sRtn <> cFRS_NORMAL Then
                    Call Form1.System1.TrmMsgBox(gSysPrm, "�I�t�Z�b�g�p�����[�^�������f����" & vbCrLf & "�t�@�C��[" & gsAutoDataFileFullPath(sRtn - 1) & "]�̏����ŃG���[���������܂����B", MsgBoxStyle.OkOnly, My.Application.Info.Title)
                    Return
                End If
            End If

            If Not Form1.TrimDataLoad(gsAutoDataFileFullPath(0)) Then
                'V2.1.0.0�C                Call Z_PRINT("�����^�]���g���~���O�f�[�^�t�@�C���k�n�`�c�G���[ = " & vbCrLf)
                Form1.AutoRunnningDisp.BackColor = System.Drawing.Color.Yellow  ' �w�i�F = ���F 'V2.0.0.2�B
                Form1.AutoRunnningDisp.Text = "�����^�]������"                                  'V2.0.0.2�B
                'V2.1.0.0�C��
                Call Z_PRINT("�����^�]���g���~���O�f�[�^�t�@�C���k�n�`�c�G���[" & vbCrLf & "= [" & gsAutoDataFileFullPath(0) & "]")
                Call Form1.System1.TrmMsgBox(gSysPrm, "�f�[�^�t�@�C���k�n�`�c�G���[" & vbCrLf & "�t�@�C��[" & gsAutoDataFileFullPath(0) & "]�̏����ŃG���[���������܂����B", MsgBoxStyle.OkOnly, My.Application.Info.Title)
                Return
                'V2.1.0.0�C��
            Else
                Call InitialAutoOperation()
                gbFgAutoOperation = True
                Form1.AutoRunnningDisp.BackColor = System.Drawing.Color.Lime ' �w�i�F = ��
                Form1.AutoRunnningDisp.Text = "�����^�]��"
                UserSub.LaserCalibrationSet(POWER_CHECK_START)          'V2.1.0.0�A ���[�U�p���[���j�^�����O���s�L���ݒ�
            End If

            mExitFlag = cFRS_ERR_START                                  ' Return�l = OK�{�^������ 

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "FormDataSelect.BtnOK_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
        Me.Close()                                                      ' �t�H�[�������
    End Sub
#End Region

#Region "Cancel�{�^������������"
    '''=========================================================================
    ''' <summary>Cancel�{�^������������</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click

        Dim strMSG As String

        Try
            mExitFlag = cFRS_ERR_RST                                    ' Return�l = Cancel�{�^������

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "FormDataSelect.BtnCancel_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
        Me.Close()                                                      ' �t�H�[�������
    End Sub
#End Region

#Region "�u���X�g�̂P��ցv�{�^������������"
    '''=========================================================================
    ''' <summary>�u���X�g�̂P��ցv�{�^������������</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnUp.Click

        Dim Idx As Integer
        Dim strMSG As String

        Try
            Idx = ListList.SelectedIndex
            ' �擪���I������Ă���ꍇNOP
            If (Idx <= 0) Then Exit Sub
            Call SwapList(Idx, (Idx - 1))       ' ؽĂ����ւ�
            ListList.SelectedIndex = (Idx - 1)  ' �P��̲��ޯ�����w�肷��

            ' �f�[�^�t�@�C�������t���p�X�Ń��x���e�L�X�g�{�b�N�X�ɐݒ肷��
            Call ListList_SelectedIndexChanged(sender, e)
            Exit Sub

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "FormDataSelect.BtnUp_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�u���X�g�̂P���ցv�{�^������������"
    '''=========================================================================
    ''' <summary>�u���X�g�̂P���ցv�{�^������������</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDown.Click

        Dim Idx As Integer
        Dim strMSG As String

        Try
            Idx = ListList.SelectedIndex
            ' �Ōオ�I������Ă���ꍇNOP
            If ((Idx + 1) >= ListList.Items.Count) Then Exit Sub
            Call SwapList(Idx, (Idx + 1))       ' ؽĂ����ւ�
            ListList.SelectedIndex = (Idx + 1)  ' �P���̲��ޯ�����w�肷��

            ' �f�[�^�t�@�C�������t���p�X�Ń��x���e�L�X�g�{�b�N�X�ɐݒ肷��
            Call ListList_SelectedIndexChanged(sender, e)
            Exit Sub

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "FormDataSelect.BtnDown_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�o�^�ς��ް�̧�ق̍��ڂ����ւ���"
    ''' <summary>�o�^�ς��ް�̧�ق̍��ڂ����ւ���</summary>
    ''' <param name="iSrc">���ʒu</param>
    ''' <param name="iDst">�ړ���ʒu</param>
    Private Sub SwapList(ByVal iSrc As Integer, ByVal iDst As Integer)
        Dim tmpStr As String
        tmpStr = ListList.Items(iSrc)
        ListList.Items.RemoveAt(iSrc)
        ListList.Items.Insert(iDst, tmpStr)

    End Sub
#End Region

#Region "�u���X�g����폜�v�{�^������������"
    '''=========================================================================
    ''' <summary>�u���X�g����폜�v�{�^������������</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDelete.Click

        Dim Idx As Integer
        Dim strMSG As String

        Try
            ' �u�o�^�ς݃f�[�^�t�@�C���v���X�g�{�b�N�X����1�폜����
            Idx = ListList.SelectedIndex
            If (Idx < 0) Then Exit Sub
            File.Delete(ENTRY_PATH & ListList.Items(Idx))    ' �I������Ă���̧�ق��폜����
            ListList.Items.RemoveAt(Idx)                                '�u�o�^�ς݃f�[�^�t�@�C���v���X�g�{�b�N�X����1���ڍ폜(��Remove()�͕�����w��)

            ' �f�[�^�t�@�C�������t���p�X�Ń��x���e�L�X�g�{�b�N�X�ɐݒ肷��(�폜�̂P�O�̃f�[�^��I����ԂƂ���)
            If (0 <= Idx) Then
                Idx = (Idx - 1)
                ' ؽĂ̐擪���폜���ꂽ�ꍇ�ɑ����ް�������ΑI������
                If (Idx < 0) AndAlso (0 < ListList.Items.Count) Then Idx = 0
                ListList.SelectedIndex = Idx    ' ����Ăɂ��o�^�ς��ް��̑I��̧�����߽���ĕ\��
            Else
                ' �o�^�ς��ް�̧�ق��Ȃ��Ȃ����ꍇ
                Call ListList_SelectedIndexChanged(sender, e)  ' �o�^�ς��ް��̑I��̧�����߽���ĕ\��
            End If

            Call DirListBox_Change(sender, e)   ' �ިڸ���ذ���ĕ\��

            ' �G���h���X���[�h����
            Call DspEndless()

            Exit Sub

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "FormDataSelect.BtnDelete_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�u���X�g���N���A�v�{�^������������"
    '''=========================================================================
    ''' <summary>�u���X�g���N���A�v�{�^������������</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>�u�o�^�ς݃f�[�^�t�@�C���v���X�g�{�b�N�X����S�č폜</remarks>
    '''=========================================================================
    Private Sub BtnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClear.Click

        Dim r As Integer
        Dim strMSG As String
        Try
            If (ListList.Items.Count < 1) Then
                ' �o�^�ς��ް�̧��ؽĂɍ��ڂ��Ȃ��ꍇ�AENTRYLOT ̫��ޓ���̧�ق����ׂč폜����
                For Each tmpFile As String In (Directory.GetFiles(ENTRY_PATH))
                    File.Delete(tmpFile)
                Next
                Exit Sub
            Else
                ' �o�^�ς��ް�̧��ؽĂɍ��ڂ�����ꍇ�A�폜�m�Fү���ނ�\������
                ' "�o�^���X�g��S�č폜���܂��B" & vbCrLf & "��낵���ł����H"
                strMSG = MSG_AUTO_15 & vbCrLf & MSG_AUTO_16
                r = MsgBox(strMSG, MsgBoxStyle.OkCancel, "")
                If (r <> MsgBoxResult.Ok) Then Exit Sub ' ��ݾ�

                ' ENTRYLOT ̫��ޓ���̧�ق����ׂč폜����
                For Each tmpFile As String In (Directory.GetFiles(ENTRY_PATH))
                    File.Delete(tmpFile)
                Next

                ' �u�o�^�ς݃f�[�^�t�@�C���v���X�g�{�b�N�X�N���A
                Call ListList.Items.Clear() '�u�o�^�ς݃f�[�^�t�@�C���v���X�g�{�b�N�X�N���A

                ' �f�[�^�t�@�C�������t���p�X�Ń��x���e�L�X�g�{�b�N�X�ɐݒ肷��(�N���A����)
                Call ListList_SelectedIndexChanged(sender, e)
                Call DirListBox_Change(sender, e) ' �ިڸ���ذ���ĕ\������

                ' �G���h���X���[�h����
                Call DspEndless()

                Exit Sub
            End If

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "FormDataSelect.BtnClear_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�o�^�{�^������������"
    '''=========================================================================
    ''' <summary>�o�^�{�^������������</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSelect.Click

        Dim Idx As Integer
        Dim Sz As Integer
        Dim Pos As Integer
        Dim strDAT As String
        Dim strMSG As String
        Try
            '�u�f�[�^�t�@�C���v���X�g�{�b�N�X�C���f�b�N�X�����Ȃ�NOP
            Idx = ListFile.SelectedIndex
            If (Idx < 0) Then Exit Sub
            ' �G���h���X���[�h�őI�����X�g1�ȏ�L��Ȃ�NOP

            ' �w��̃f�[�^�t�@�C�������u�o�^�ς݃f�[�^�t�@�C���v���X�g�{�b�N�X�ɒǉ�����
            strDAT = ListFile.Items(Idx)                                ' ���t�����t���t�@�C�����Ȃ̂Ńt�@�C�����̂ݎ��o�� 
            Sz = strDAT.Length
            Pos = strDAT.LastIndexOf(" ")
            If (Pos = -1) Then Exit Sub
            strDAT = strDAT.Substring(Pos + 1, Sz - Pos - 1)
            Idx = ListList.Items.Count

            Dim sFromFilePath As String = ""
            Dim sCopyFilePath As String = ""
            If (False = CopyEntryFileToWorkFolder(sFromFilePath, sCopyFilePath, strDAT)) Then ' �I��̧�ق��߰����
                ' TODO: �װү����
                MsgBox((sFromFilePath & vbCrLf & vbTab & "��" & vbCrLf & sCopyFilePath), _
                    DirectCast((MsgBoxStyle.Critical + MsgBoxStyle.OkOnly), MsgBoxStyle))
            Else
                Call ListList.Items.Add(strDAT)
                ListList.SelectedIndex = Idx

                ' �f�[�^�t�@�C�������t���p�X�Ń��x���e�L�X�g�{�b�N�X�ɐݒ肷��
                Call ListList_SelectedIndexChanged(sender, e)

            End If

            ' �G���h���X���[�h����
            Call DspEndless()
            Exit Sub

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "FormDataSelect.BtnSelect_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�o�^̧�ق�ENTRYLOT̫��ނɺ�߰����"
    ''' <summary>�o�^̧�ق�ENTRYLOT̫��ނɺ�߰����</summary>
    ''' <param name="sFromFilePath">IN="",OUT=��߰�����߽</param>
    ''' <param name="sCopyFilePath">IN="",OUT=��߰�����߽</param>
    ''' <param name="sCopyFile">IN=��߰����̧�ٖ�.�g���q,OUT=��߰����̧�ٖ�.�g���q</param>
    ''' <returns>True=����,False=���s</returns>
    ''' <remarks>̧�ٖ��ɖ���_01,_02�ƘA�Ԃ�t������</remarks>
    Private Function CopyEntryFileToWorkFolder(ByRef sFromFilePath As String, _
                ByRef sCopyFilePath As String, ByRef sCopyFile As String) As Boolean
        Dim sTmpFile As String          ' ̧�ٖ�
        Dim sExtended As String         ' �g���q

        CopyEntryFileToWorkFolder = False
        Try
            sTmpFile = sCopyFile.Split(".")(0)
            sExtended = "." & sCopyFile.Split(".")(1)

            For i As Integer = 0 To 99 Step 1
                ' �A�Ԃ�ǉ�����̧�ٖ��̍쐬
                sCopyFilePath = (ENTRY_PATH & sTmpFile & "_" & i.ToString("00") & sExtended)
                Debug.Print(sCopyFilePath)
                ' ����̧�ق̑��݊m�F
                If (False = File.Exists(sCopyFilePath)) Then
                    ' ���݂��Ȃ����̧�ق��߰
                    sFromFilePath = (FileLstBox.Path & "\" & sTmpFile & sExtended)
                    File.Copy(sFromFilePath, sCopyFilePath)
                    If (File.Exists(sCopyFilePath)) Then
                        sCopyFile = (sTmpFile & "_" & i.ToString("00") & sExtended)
                        CopyEntryFileToWorkFolder = True
                    End If
                    Exit Function
                End If
            Next i

        Catch ex As Exception
            Dim strMSG As String = _
                "FormDataSelect.CopyEntryFileToWorkFolder() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try

    End Function
#End Region

    '========================================================================================
    '   ���X�g�{�b�N�X�̃N���b�N�C�x���g����
    '========================================================================================
#Region "�u�f�[�^�t�@�C���v���X�g�{�b�N�X�_�u���N���b�N�C�x���g����"
    '''=========================================================================
    ''' <summary>�u�f�[�^�t�@�C���v���X�g�{�b�N�X�_�u���N���b�N�C�x���g����</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub ListFile_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListFile.DoubleClick
        Dim strMSG As String

        Try
            ' �o�^�{�^��������������
            Call BtnSelect_Click(sender, e)
            Exit Sub

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "FormDataSelect.ListFile_DoubleClick() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�u�o�^�ς݃f�[�^�t�@�C���vؽĂ̲��ޯ���ύX�����"
    '''=========================================================================
    ''' <summary>�u�o�^�ς݃f�[�^�t�@�C���vؽĂ̲��ޯ���ύX�����</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub ListList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                Handles ListList.SelectedIndexChanged
        Dim Idx As Integer
        Dim strMSG As String
        Try
            ' �u�o�^�ς݃f�[�^�t�@�C���v���X�g�{�b�N�X�őI�����ꂽ�f�[�^�t�@�C�������t���p�X�Ń��x���e�L�X�g�{�b�N�X�ɐݒ肷��
            Idx = ListList.SelectedIndex
            If (Idx < 0) Then
                LblFullPath.Text = ""
            Else
                LblFullPath.Text = (ENTRY_PATH & ListList.Items(Idx))
            End If
            Exit Sub

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "FormDataSelect.ListList_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�u�h���C�u���X�g�{�b�N�X�v�� SelectedIndexChanged ����"
    '''=========================================================================
    ''' <summary>�u�h���C�u���X�g�{�b�N�X�v�� SelectedIndexChanged ����</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>DriveListBox�͔�W���R���g���[���Ȃ̂Ńc�[���{�b�N�X�ɒǉ�����K�v�L��</remarks>
    '''=========================================================================
    Private Sub DrvListBox_SelectedIndexChanged( _
        ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DrvListBox.SelectedIndexChanged

        Try
            ' �f�B���N�g�����X�g�{�b�N�X�̑I���h���C�u��ύX����
            Dim tmpDrv As String = DrvListBox.Drive
            If (0 = (String.Compare(tmpDrv, "C:", True))) Then tmpDrv = DATA_DIR_PATH
            DirListBox.Path = tmpDrv
            Call DirListBox_Change(sender, e)   ' �ިڸ���ذ���ĕ\������
            Exit Sub

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            Dim strMSG As String = ex.Message
            MsgBox(strMSG)
            DrvListBox.Drive = "C:"
        End Try
    End Sub
#End Region

#Region "�u�f�B���N�g�����X�g�{�b�N�X�v�̕ύX������"
    '''=========================================================================
    ''' <summary>�u�f�B���N�g�����X�g�{�b�N�X�v�̕ύX������</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>DirListBox�͔�W���R���g���[���Ȃ̂Ńc�[���{�b�N�X�ɒǉ�����K�v�L��</remarks>
    '''=========================================================================
    Private Sub DirListBox_Change(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DirListBox.Change

        Dim strMSG As String

        Try
            ' �I���f�B���N�g����ύX����(FileLstBox�͍�Ɨp��Dummy)
            FileLstBox.Path = DirListBox.Path

            ' �u�f�[�^�t�@�C���v���X�g�{�b�N�X�ɓ��t�����t���t�@�C������\������
            MakeFileList()
            Exit Sub

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "FormDataSelect.DirListBox_Change() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub

    ''' <summary>�ިڸ���ޯ���د����̏���</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>���̏����ɂ�� DirListBox_Change ����Ă��������A�ިڸ���ذ���ĕ\������</remarks>
    Private Sub DirListBox_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DirListBox.Click
        Try
            With DirectCast(sender, VB6.DirListBox)
                If (.Path <> .DirList(.DirListIndex)) Then
                    .Path = .DirList(.DirListIndex) ' �I�������ިڸ�؂��߽�ɐݒ肷��
                End If
            End With
        Catch ex As Exception
            Dim strMSG As String = "FormDataSelect.DirListBox_Click() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

    '========================================================================================
    '   ���샂�[�h���W�I�{�^���ύX���̏���
    '========================================================================================
#Region "�}�K�W�����[�h�I������"
    '''=========================================================================
    ''' <summary>�}�K�W�����[�h�I������</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnMdMagazine_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call DspEndless()
    End Sub
#End Region

#Region "���b�g���[�h�I������"
    '''=========================================================================
    ''' <summary>���b�g���[�h�I������</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnMdLot_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call DspEndless()
    End Sub
#End Region

#Region "�G���h���X���[�h�I������"
    '''=========================================================================
    ''' <summary>�G���h���X���[�h�I������</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    '''=========================================================================
    Private Sub BtnMdEndless_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call DspEndless()
    End Sub
#End Region

    '========================================================================================
    '   ���ʊ֐���`
    '========================================================================================
#Region "�u�f�[�^�t�@�C���v���X�g�{�b�N�X�ɓ��t�����t���t�@�C������\������"
    '''=========================================================================
    ''' <summary>�u�f�[�^�t�@�C���v���X�g�{�b�N�X�ɓ��t�����t���t�@�C������\������</summary>
    ''' <remarks></remarks>
    '''=========================================================================
    Private Sub MakeFileList()

        Dim Count As Integer
        Dim i As Integer
        Dim Sz As Integer
        Dim strWK As String
        Dim strDAT As String
        Dim strMSG As String

        Try
            ' �u�f�[�^�t�@�C���v���X�g�{�b�N�X�ɓ��t�����t���t�@�C������\������
            Call ListFile.Items.Clear()                                                 '�u�f�[�^�t�@�C���v���X�g�{�b�N�X�N���A
            FileLstBox.Refresh()                                                        ' �t�@�C�����X�g���X�V����  'V4.7.0.0�L
            Count = FileLstBox.Items.Count                                              ' �t�@�C���̐� 
            For i = 0 To (Count - 1)
                ' �t�@�C���g���q��ݒ�
                strWK = ".txt"

                ' �Ώۂ̊g���q�łȂ����SKIP
                strDAT = FileLstBox.Items(i)
                Sz = strDAT.Length
                If (Sz < 4) Then GoTo STP_NEXT
                strDAT = strDAT.Substring(Sz - 4, 4)                                    ' �g���q�����o��
                If (String.Compare(strDAT, strWK, True)) Then GoTo STP_NEXT '           ' �Ώۂ̊g���q�łȂ����SKIP(�啶���A����������ʂ��Ȃ�)

                ' ���t�����t���t�@�C�����X�g�쐬
                Dim tmpFile As String = FileLstBox.Path & "\" & FileLstBox.Items(i)
                If (False = (File.Exists(tmpFile))) Then Continue For ' ̧�ق̑��݊m�F
                strDAT = FileDateTime(tmpFile)
                Dim Dt As DateTime = DateTime.Parse(strDAT)
                strDAT = Dt.ToString("yyyy/MM/dd HH:mm:ss") + " " + FileLstBox.Items(i) ' ���t�����̒��������킹�� 
                Call ListFile.Items.Add(strDAT)                                         ' ���t�����t���t�@�C������\������
STP_NEXT:
            Next i
            Exit Sub

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "FormDataSelect.MakeFileList() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�G���h���X���[�h����"
    '''=========================================================================
    ''' <summary>�G���h���X���[�h����</summary>
    ''' <remarks>�G���h���X���[�h���̓f�[�^�t�@�C�����P�����I���ł��Ȃ�</remarks>
    '''=========================================================================
    Private Sub DspEndless()

        Dim strMSG As String

        Try
            ' �G���h���X���[�h�őI�����X�g1�ȏ�L��Ȃ牺�L�̃{�^������񊈐����ɂ���
            'If (BtnMdEndless.Checked = True) And (ListList.Items.Count >= 1) Then
            '    ListFile.Enabled = False                                ' �f�[�^�t�@�C�����X�g�{�b�N�X�񊈐���
            '    BtnSelect.Enabled = False                               ' �o�^�{�^���񊈐��� 
            '    BtnUp.Enabled = False                                   '�u���X�g�̂P��ցv�{�^���񊈐���
            '    BtnDown.Enabled = False                                 '�u���X�g�̂P���ցv�{�^���񊈐���
            'Else
            ListFile.Enabled = True                                 ' �f�[�^�t�@�C�����X�g�{�b�N�X������
            BtnSelect.Enabled = True                                ' �o�^�{�^�������� 
            BtnUp.Enabled = True                                    '�u���X�g�̂P��ցv�{�^��������
            BtnDown.Enabled = True                                  '�u���X�g�̂P���ցv�{�^��������
            'End If

            Exit Sub

            ' �g���b�v�G���[������ 
        Catch ex As Exception
            strMSG = "FormDataSelect.DspEndless() TRAP ERROR = " + ex.Message
            MsgBox(strMSG)
        End Try
    End Sub
#End Region

#Region "�����^�]�p�t�@���N�V����"
    '=========================================================================
    '�y�@�@�\�z�v���[�g�f�[�^�t�@�C�����̐ݒ�
    '�y���@���z0:�ݒ� 1:�ݒ薳��
    '�y�߂�l�z�A�������^�]�̃��O�t�@�C�����𐶐�����B
    '=========================================================================
    Public Function PlateDataFileName(ByVal mode As Integer, ByVal sName As String) As String

        If mode = 0 Then
            sPlateDataFileName = sName
        End If

        PlateDataFileName = sPlateDataFileName

    End Function

    '=========================================================================
    '�y�@�@�\�z�v���[�g�f�[�^�E�t�@�C�����폜����B
    '�y���@���z����
    '�y�߂�l�z����
    '=========================================================================
    Public Sub SavePlateDataFileDelete()
        Dim sFolder As String

        Try

            sFolder = ENTRY_PATH & ENTRY_TMP_FILE

            If IO.File.Exists(sFolder) = True Then  ' �t�@�C�����L��΍폜����B
                IO.File.Delete(sFolder)
            End If
        Catch ex As Exception
            Call Z_PRINT("FormDataSelect.SavePlateDataFileDelete() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Sub
    '=========================================================================
    '�y�@�@�\�z�r���ŏI���������Ƀv���[�g�f�[�^��ۑ�����B
    '�y���@���z�v���[�g�f�[�^�z��A�X�^�[�g(0 Origin)�A�I��
    '�y�߂�l�z����
    '=========================================================================
    Public Sub SavePlateDataFileFullPath(ByRef sPath() As String, ByVal iStart As Integer, ByVal iEnd As Integer)
        Dim sFolder As String
        Dim iFileNo As Integer
        Dim WS As IO.StreamWriter

        Try

            sFolder = ENTRY_PATH & ENTRY_TMP_FILE

            Call SavePlateDataFileDelete()

            WS = New IO.StreamWriter(sFolder, True, System.Text.Encoding.GetEncoding("Shift-JIS"))

            For iFileNo = iStart To iEnd
                WS.WriteLine(sPath(iFileNo))
            Next

            WS.Close()

        Catch ex As Exception
            Call Z_PRINT("FormDataSelect.SavePlateDataFileFullPath() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try

    End Sub
    '=========================================================================
    '�y�@�@�\�z�r���ŏI���������ɕۑ����ꂽ�v���[�g�f�[�^�����[�h����B
    '�y���@���z����
    '�y�߂�l�z����
    '=========================================================================
    Private Sub LoadPlateDataFileFullPath() ' private V4.7.0.0�L
        'Public Sub LoadPlateDataFileFullPath()
        Dim sFolder As String
        Dim sPathData As String = ""

        sFolder = ENTRY_PATH & ENTRY_TMP_FILE

        If IO.File.Exists(sFolder) = True Then  ' �t�@�C�����ǂݎ��B
            Using sr As New System.IO.StreamReader(sFolder, System.Text.Encoding.GetEncoding("Shift-JIS"))
                Do While Not sr.EndOfStream
                    Dim sPath As String = sr.ReadLine
                    If sPath <> "" Then
                        'V1.1.0.1�A                        ListList.Items.Add(sPath)
                        ListList.Items.Add(IO.Path.GetFileName(sPath))           'V1.1.0.1�A
                    Else
                        MsgBox("�t�@�C�������݂��܂���ł��� =" & sPathData, vbOKOnly Or vbExclamation Or vbSystemModal Or vbMsgBoxSetForeground, "Warning")
                    End If
                Loop
            End Using
        End If

    End Sub

    '=========================================================================
    '�y�@�@�\�z�A�������^�]�p�A���g���~���O�m�f�����J�E���^�[
    '�y��P�����z0:�������A1:NG�J�E���g�ݒ� ���̑��F���݃J�E���^�[�̎擾
    '�y��Q�����z�m�f�J�E���^�[�l
    '�y�߂�l�z�m�f�J�E���^�[�l
    '=========================================================================
    Private Function NGCountData(ByVal mode As Integer, ByVal lNgCount As Long) As Long

        If mode = 0 Then
            InitiallNgCount = lNgCount
            NowlNgCount = 0
        ElseIf mode = 1 Then
            NowlNgCount = lNgCount - InitiallNgCount
        End If

        NGCountData = NowlNgCount
        Debug.Print("InitiallNgCount=" & InitiallNgCount & "NowlNgCount=" & NowlNgCount)
    End Function
    '=========================================================================
    '�y�@�@�\�z�A���g���~���O�m�f�����J�E���^�[�ݒ菉����
    '�y���@���z�m�f�J�E���^�[�l
    '�y�߂�l�z����
    '=========================================================================
    Public Sub InitNGCountForContinueAuto(ByVal lNgCount As Long)
        Call NGCountData(0, lNgCount)
    End Sub

    '=========================================================================
    '�y�@�@�\�z�A���g���~���O�m�f�����J�E���^�[�ݒ�
    '�y���@���z�m�f�J�E���^�[�l
    '�y�߂�l�z����
    '=========================================================================
    Public Sub SetNGCountForContinueAuto(ByVal lNgCount As Long)
        Call NGCountData(1, lNgCount)
    End Sub

    '=========================================================================
    '�y�@�@�\�z���b�g�؂�ւ����菉����
    '�y���@���z����
    '�y�߂�l�z����
    '=========================================================================
    Public Sub InitLotChangeJudge()
        bLotChange = False
    End Sub

    '=========================================================================
    '�y�@�@�\�z���b�g�؂�ւ�����Z�b�g
    '�y���@���z����
    '�y�߂�l�z����
    '=========================================================================
    Public Sub SetLotChangeJudge()
        bLotChange = True
    End Sub
    '=========================================================================
    '�y�@�@�\�z���b�g�؂�ւ�����擾
    '�y���@���z����
    '�y�߂�l�z����
    '=========================================================================
    Public Function GetLotChangeJudge() As Boolean
        GetLotChangeJudge = bLotChange
        If bLotChange Then
            bLotChange = False
        End If
    End Function
    '=========================================================================
    '�y�@�@�\�z�A�������^�]���[�h�̏�����
    '�y���@���z����
    '�y�߂�l�z����
    '=========================================================================
    Public Sub InitialAutoOperation()
        AutoOpeCancel = False
        NowExecuteLotNo = 0
        CancelReason = 0

        MarkingCount = 0                ' �}�[�L���O�p�J�E���^�N���A	            V2.2.1.7�B
        LotMarkingAlarmCnt = 0          ' �}�[�L���O���s���A���[�����J�E���^�N���A	            V2.2.1.7�B

        Call PlateDataFileName(0, gsAutoDataFileFullPath(0))  ' �v���[�g�f�[�^�t�@�C������ۑ�
        Call InitLotChangeJudge()
        Call Form1.System1.AutoLoaderFlgReset()                 'V1.2.0.0�C �I�[�g���[�_�[�t���O���Z�b�g
    End Sub

    '=========================================================================
    '�y�@�@�\�z���b�g�؂�ւ�����Z�b�g
    '�y���@���z����
    '�y�߂�l�z����
    '=========================================================================
    Public Function GetAutoOpeCancelStatus() As Boolean
        If gbFgAutoOperation = True Then
            GetAutoOpeCancelStatus = AutoOpeCancel
        Else
            GetAutoOpeCancelStatus = False
        End If
    End Function

    '=========================================================================
    '�y�@�@�\�z���b�g�؂�ւ������ۃ`�F�b�N
    '�y���@���z����
    '�y�߂�l�z����
    '=========================================================================
    Public Function LotChangeExecuteCheck() As Boolean

        If NowExecuteLotNo + 1 >= giAutoDataFileNum Then
            LotChangeExecuteCheck = False
        Else
            LotChangeExecuteCheck = True
        End If
    End Function

    '=========================================================================
    '�y�@�@�\�z���b�g�؂�ւ�����
    '�y���@���z����
    '�y�߂�l�z����
    '=========================================================================
    Public Function LotChangeExecute() As Boolean
        Try
            If LotChangeExecuteCheck() Then
                NowExecuteLotNo = NowExecuteLotNo + 1
                MarkingCount = 0             ' �}�[�L���O�p�J�E���^�N���A	V2.2.1.7�B
                SetAutoOpeStartTime()          ' V2.2.1.7�B
                'V2.1.0.0�C                Call Form1.TrimDataLoad(gsAutoDataFileFullPath(NowExecuteLotNo))
                'V2.1.0.0�C��
                If Not Form1.TrimDataLoad(gsAutoDataFileFullPath(NowExecuteLotNo)) Then
                    Call Z_PRINT("�����^�]���g���~���O�f�[�^�t�@�C���k�n�`�c�G���[" & vbCrLf & "= [" & gsAutoDataFileFullPath(NowExecuteLotNo) & "]")
                    AutoOpeCancel = True
                    LotChangeExecute = False
                    Exit Function
                Else
                    'V2.1.0.0�C��
                    Call PlateDataFileName(0, gsAutoDataFileFullPath(NowExecuteLotNo))  ' �v���[�g�f�[�^�t�@�C������ۑ�
                    LotChangeExecute = True
                End If                      'V2.1.0.0�C
            Else
                Call Z_PRINT("���b�g�؂�ւ��M�����󂯂܂������A���̃G���g���[���L��܂���B" & vbCrLf)
                CancelReason = NO_MORE_ENTRY
                AutoOpeCancel = True
                LotChangeExecute = False
            End If
        Catch ex As Exception
            Call Z_PRINT("FormDataSelect.LotChangeExecute() TRAP ERROR = " & ex.Message & vbCrLf)
        End Try
    End Function


    '=========================================================================
    '�y�@�@�\�z�A�������^�]�I������
    '�y���@���z����
    '�y�߂�l�z����
    '=========================================================================
    Public Sub AutoOperationEnd()

        If gbFgAutoOperation = False Then
            Exit Sub
        End If

        If giLoaderType = 0 Then            'V2.2.1.1�A�����ǉ�
            Call Sub_ATLDSET(COM_STS_LOT_END, 0)    'V1.2.0.0�C ���[�_�[�o��(ON=���b�g�I��,OFF=�Ȃ�)
        End If

        Call Form1.System1.AutoLoaderFlgReset() 'V1.2.0.0�C �I�[�g���[�_�[�t���O���Z�b�g

        Call SavePlateDataFileDelete()

        NowExecuteLotNo = NowExecuteLotNo + 1
        MarkingCount = 0                ' �}�[�L���O�p�J�E���^�N���A	            V2.2.1.7�B
        Form1.DispMarkAlarmList()       ' �}�[�N�󎚂̃G���[���X�g����ʂɕ\��        V2.2.1.7�B
        LotMarkingAlarmCnt = 0          ' �}�[�L���O���s���A���[�����J�E���^�N���A	            V2.2.1.7�B

        If AutoOpeCancel Or (NowExecuteLotNo < giAutoDataFileNum) Then
            If AutoOpeCancel Then
                NowExecuteLotNo = NowExecuteLotNo - 1   ' ���݂̃v���[�g�f�[�^����ۑ�����B
            End If
            If NowExecuteLotNo < giAutoDataFileNum Then
                Call SavePlateDataFileFullPath(gsAutoDataFileFullPath, NowExecuteLotNo, giAutoDataFileNum - 1)
            End If
        End If

        Call UserSub.SetStartCheckStatus(True)          'V1.2.0.0�C �ݒ��ʂ̊m�F�L����
        gbFgAutoOperation = False
        Form1.AutoRunnningDisp.BackColor = System.Drawing.Color.Yellow  ' �w�i�F = ���F 'V2.0.0.2�B
        Form1.AutoRunnningDisp.Text = "�����^�]������"                                  'V2.0.0.2�B

    End Sub

#End Region

    ''' <summary>
    ''' ���ݎ��s���̓o�^�t�@�C��No��Ԃ� 
    ''' </summary>
    ''' <returns></returns>
    Public Function GetNowLotDataNo() As Integer
        Try

            Return NowExecuteLotNo

        Catch ex As Exception

        End Try

    End Function


    ''' <summary>
    ''' AutoOpeCancel �̃t���O��ݒ肷�� 
    ''' </summary>
    ''' <param name="mode"></param>
    Public Sub SetAutoOpeCancel(ByVal mode As Boolean)
        Try
            AutoOpeCancel = mode
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' �N�����v�{�^�������@'V2.2.1.1�H
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnClamp_Click(sender As Object, e As EventArgs) Handles btnClamp.Click
        Dim r As Integer

        Try

            btnClamp.BackColor = Color.Yellow
            btnClamp.Enabled = False

            ' �ڕ���N�����vON   
            r = Form1.System1.ClampCtrl(gSysPrm, 1, 0)
            If (r <> cFRS_NORMAL) Then

            End If

            System.Threading.Thread.Sleep(500)

            ' �ڕ���N�����vOFF 
            r = Form1.System1.ClampCtrl(gSysPrm, 0, 0)
            If (r <> cFRS_NORMAL) Then

            End If



        Catch ex As Exception
        Finally
            btnClamp.Enabled = True
            btnClamp.BackColor = SystemColors.ButtonFace

        End Try

    End Sub

#End Region
End Class

'=============================== END OF FILE ===============================