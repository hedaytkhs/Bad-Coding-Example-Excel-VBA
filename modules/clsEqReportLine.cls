VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEqReportLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************
'
'    Equation Report�̗v�f�𑀍삷�邽�߂̃N���X
'
'    Equation Report�̍��ږ��A�l�Ȃǋ��ʗv�f���܂Ƃ߂�
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************
Option Explicit

Dim IsGroup_ As Boolean
Dim IsEquationReportItem_ As Boolean
Dim Text_ As String
Dim Value_ As String

Dim EquationReportItemID_ As EquationReportItem



Property Get IsGroup() As Boolean
    IsGroup = IsGroup_
End Property

Property Get IsEquationReportItem() As Boolean
    IsEquationReportItem = IsEquationReportItem_
End Property

Property Let Text(ByVal sText As String)
    Text_ = sText
    '���肷��
End Property

Property Get Value() As String
    Value = Value_
End Property

'���肷��
'Item�̎�ޕ����A��ԃt���O�AValue��ReportItem�̓��e���i�[����





'EquationReportItem�̏ꍇ�́AEquationReportItemID��ݒ�


'Group�̔���
' ":"�ŋ�؂��镶��������B���A:�̑O�̕������ ATA��������LRU�ł���B

'EquationReportItem�̔���
'vbTab & "="�ŋ�؂��镶��������B���A=�̑O�̕�����́AEquation Report Title�̂����ꂩ�̕�����ƈ�v����

'Help Text���J�n���ꂽ��t���O�𗧂Ă�
'Logic���J�n���ꂽ��t���O�𗧂Ă�
'


' "Possible Causes (LRU Short Name)"���܂܂��
' Value�����o��

'Possible Causes List�̊J�n�s�ł���

'Variable    Variable Details�̊J�n�s�ł���

' Variable�̓r���̍s�ł���ꍇ
' Variable��Variable Details�ɕ������� Variable��Variable Details���ꂼ����R���N�V������AddItem����


'�قȂ�Equation Name���o�Ă�����AclsFaultEquation�̌��݂̃C���X�^���X����Excel��1�s�ɓ]�L����

'clsFaultEquation�̎��̃C���X�^���X����������
' ����Equation Name�̒l���Z�b�g����





Private Sub Class_Initialize()
    IsGroup_ = False
    IsEquationReportItem_ = False
End Sub