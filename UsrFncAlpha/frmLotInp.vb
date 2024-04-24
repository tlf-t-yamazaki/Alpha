Option Strict Off
Option Explicit On
Friend Class frmLotInp
	Inherits System.Windows.Forms.Form
	'==============================================================================
	'
	'   DESCRIPTION:    ���b�g�ԍ��\�� & ����
	'
	'==============================================================================
	'-------------------------------------------------------------------------------
	'   �����ϐ���`
	'-------------------------------------------------------------------------------
	Private Const MAX_LOT_LEN As Short = 64 ' MAX���b�g�ԍ�������
	Private mExitFlag As Short ' ����(0:����, 1:OK(ADV��), 3:Cancel(RESET��))
	
	'===============================================================================
	'�y�@�@�\�z OK/Cancel���ʎ擾
	'�y���@���z �Ȃ�
	'�y�߂�l�z ���� = 1:OK(ADV��), 3:Cancel(RESET��)
	'===============================================================================
	Public Function GetResult() As Short
		
		GetResult = mExitFlag
		
	End Function
	
	'===============================================================================
	'�y�@�@�\�z Cancel�{�^������������
	'�y���@���z �Ȃ�
	'�y�߂�l�z �Ȃ�
	'===============================================================================
	Private Sub CmndCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmndCancel.Click
		
		mExitFlag = cFRS_ERR_RST ' ExitFlag = 3:Cancel(RESET��))
		
	End Sub
	
	'===============================================================================
	'�y�@�@�\�z OK�{�^������������
	'�y���@���z �Ȃ�
	'�y�߂�l�z �Ȃ�
	'===============================================================================
	Private Sub CmndOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmndOK.Click
		
		Dim strMSG As String
0:      'Dim strDATE As String ' ���ݓ���("YYYY/MM/DD HH:MM")
        'Dim i As Short
		Dim r As Short
		
        If (TextLOT.Text = "") Then                                     ' ���͂Ȃ� ?
            TextLOT.Focus()                                             ' �t�H�[�J�X�ݒ�
            Exit Sub
        End If
		
		' �m�Fү���ޕ\��
		strMSG = "���b�g�ԍ���؂�ւ��܂��B��낵���ł����H"
        r = Form1.System1.TrmMsgBox(gSysPrm, strMSG, MsgBoxStyle.OkCancel, My.Application.Info.Title)
        If (r = cFRS_ERR_RST) Then Exit Sub ' Cancel(RESET��) �Ȃ�EXIT

        ' �f�[�^�͈̓`�F�b�N
        r = Data_Check()
        If (r <> 0) Then Exit Sub

        ' ���b�g���ݒ�
        stUserData.sLotNumber = TextLOT.Text
        Call Disp_frmInfo(COUNTER.PRODUCT_INIT, COUNTER.NONE)                                        ' ���Y��������(frmInfo��ʂ��ĕ\��)

        ' ���O�t�@�C������ݒ肷�� ("C:\TRIMDATA\LOG\""LOG_yyyymmdd" + ".LOG")
        Call SetLogFileName(gsLogFileName)

        mExitFlag = cFRS_ERR_ADV                                        ' ExitFlag = 1:OK(ADV��)

    End Sub

    '===============================================================================
    '�y�@�@�\�z �f�[�^�͈̓`�F�b�N����
    '�y���@���z �Ȃ�
    '�y�߂�l�z ���� = 0:OK, 0�ȊO:�װ
    '===============================================================================
    Private Function Data_Check() As Short

        On Error GoTo STP_TRAP
        Dim strMSG As String
        Dim iLen As Short

        Data_Check = cFRS_NORMAL ' Return�l = ����
        iLen = Len(TextLOT.Text)
        If (iLen > MAX_LOT_LEN) Then GoTo STP_ERR ' �f�[�^�͈̓`�F�b�N

        Exit Function

STP_ERR:
        strMSG = "���b�g�ԍ���" & MAX_LOT_LEN.ToString("0") & "�����ȓ��Ŏw�肵�ĉ�����"
        Call Form1.System1.TrmMsgBox(gSysPrm, strMSG, MsgBoxStyle.OkOnly, My.Application.Info.Title)
        Data_Check = 1 ' Return�l = �f�[�^�͈̓`�F�b�N�G���[
        TextLOT.Focus() ' �t�H�[�J�X�ݒ�
        Exit Function

STP_TRAP:
        Data_Check = cERR_TRAP ' Return�l = �ׯ�ߴװ����

    End Function

    '===============================================================================
    '�y�@�@�\�z Form_Activate������
    '�y���@���z �Ȃ�
    '�y�߂�l�z �Ȃ�
    '===============================================================================
    'UPGRADE_WARNING: Form �C�x���g frmLotInp.Activate �ɂ͐V�������삪�܂܂�܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' ���N���b�N���Ă��������B
    Private Sub frmLotInp_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        'Dim i As Short
        Dim r As Short
        Dim st As Integer

        ' ���b�g�ԍ��ݒ�
        mExitFlag = 0
        TextLOT.Text = stUserData.sLotNumber ' ���b�g�ԍ�
        TextLOT.Focus() ' �t�H�[�J�X�ݒ�
        TextLOT.SelectionStart = 0 ' ÷�Ă̑I��͈�(�擪�����̒��O����)
        TextLOT.SelectionLength = Len(TextLOT.Text) ' �I��͈͂̕�����

        ' OK/Calcel���͑҂�
        Call ZCONRST() ' �R���\�[���L�[���b�`����
        Do
            System.Windows.Forms.Application.DoEvents() ' ���b�Z�[�W�|���v

            ' ����~���`�F�b�N
            'r = form1.System1.Sys_Err_Chk(gSysPrm, APP_MODE_LOTCHG, Form1)
            r = Form1.System1.Sys_Err_Chk_EX(gSysPrm, APP_MODE_LOTCHG)
            If (r <> cFRS_NORMAL) Then ' ����~�� ?
                mExitFlag = r
                Exit Do
            End If

            ' �R���\�[������
            Call ZINPSTS(1, st) ' �R���\�[������
            If st And &H4S Then ' ADV �L�[��������Ă��邩�H
                Call ZCONRST() ' �R���\�[���L�[���b�`����
                Call CmndOK_Click(CmndOk, New System.EventArgs()) ' OK�{�^������������

            ElseIf st And &H8S Then  ' RESET�� ?
                Call ZCONRST() ' �R���\�[���L�[���b�`����
                Call CmndCancel_Click(CmndCancel, New System.EventArgs()) ' Cancel�{�^������������
            End If

        Loop While (mExitFlag = 0)

        Call ZCONRST() ' �R���\�[���L�[���b�`����
        Me.Close()

    End Sub
End Class