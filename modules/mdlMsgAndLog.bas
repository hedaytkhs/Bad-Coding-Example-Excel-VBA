Attribute VB_Name = "mdlMsgAndLog"

'****************************************************************************************
'
'    Application���A�\�����b�Z�[�W���̒萔�擾�p�̃��W���[��
'
'    �񋓌^�ɒ�`�������̂��I�[�g�R���v���[�g�𗘗p���ăv���O�����쐬���ȗ͉��y�щǐ������߂邽�߂ɗp����
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************
Option Explicit

Enum APP_CNST_ID
    AppName = 1
    ErrNumber = 2
    ErrDescription = 3
    msgTextFileNotSpecified = 4
    msgPreProcessFailed = 5
    
    msgImportCompleted = 6
    msgImportFailed = 7
    FileSelectTextFilter = 8
    FileSelectTitle = 9
    msgShowLogFile = 10
End Enum

Public Function GetAppCNST(CNST_ID As APP_CNST_ID) As String

    Dim sRet As String
    Select Case CNST_ID
        Case APP_CNST_ID.AppName: sRet = "EquationReport��Excel�ɃC���|�[�g����"
        Case APP_CNST_ID.ErrNumber: sRet = vbCrLf & "ErrNumber: "
        Case APP_CNST_ID.ErrDescription: sRet = vbCrLf & "ErrDescription: "
        Case APP_CNST_ID.msgTextFileNotSpecified: sRet = "�e�L�X�g�t�@�C�����w�肳��܂���ł���." & vbCrLf & "EquationReport�̃C���|�[�g�𒆎~���܂�."
        Case APP_CNST_ID.msgPreProcessFailed: sRet = "�O�����Ɏ��s���܂���." & vbCrLf & "EquationReport�̃C���|�[�g�𒆎~���܂�."
    
        Case APP_CNST_ID.msgImportCompleted: sRet = "EquationReport��Excel�ɃC���|�[�g���܂���."
        Case APP_CNST_ID.msgImportFailed: sRet = "EquationReport�̃C���|�[�g�Ɏ��s���܂���."
        Case APP_CNST_ID.FileSelectTextFilter: sRet = "Equation Report �e�L�X�g�t�@�C��,*.txt"
        Case APP_CNST_ID.FileSelectTitle: sRet = "Equation Report �e�L�X�g�t�@�C����I�����Ă�������"
        Case APP_CNST_ID.msgShowLogFile: sRet = "�G���[�E���O��\�����܂���?"
    End Select
    
    GetAppCNST = sRet
End Function

