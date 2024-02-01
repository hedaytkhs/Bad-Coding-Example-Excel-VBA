VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFaultEquation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************
'
'    Equation Report 1アイテム内の各要素を操作するためのクラス
'
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************


Option Explicit


Dim EquationReportItem As clsEquationReportItem
Public EquationReportItems As Collection
Dim EquationVariable As clsVariable
Public EquationVariables As Collection

'Public Function AddVariable(ByRef EquationVariable As tpEquationVariable) As Boolean
'    Set EquationVariable = New clsVariable
'
'    With EquationVariable
'        .Variable = EquationVariable.Variable
'        .VariableDetails = EquationVariable.VariableDetails
'    End With
'
'    EquationVariables.Add EquationVariable
'
'End Function


Private Sub Class_Initialize()
    Set EquationReportItems = New Collection
    Set EquationVariables = New Collection
    
    Set EquationReportItem = New clsEquationReportItem
    With EquationReportItem
        .ID = GroupATA
        .Name = GetEquationReportTitle(GroupATA)
        .Value = ""
    End With
    EquationReportItems.Add EquationReportItem

    Set EquationReportItem = New clsEquationReportItem
    With EquationReportItem
        .ID = GroupLRU
        .Name = GetEquationReportTitle(GroupLRU)
        .Value = ""
    End With
    EquationReportItems.Add EquationReportItem

    Set EquationReportItem = New clsEquationReportItem
    With EquationReportItem
        .ID = EquationName
        .Name = GetEquationReportTitle(EquationName)
        .Value = ""
    End With
    EquationReportItems.Add EquationReportItem

    Set EquationReportItem = New clsEquationReportItem
    With EquationReportItem
        .ID = Priority
        .Name = GetEquationReportTitle(Priority)
        .Value = ""
    End With
    EquationReportItems.Add EquationReportItem

    Set EquationReportItem = New clsEquationReportItem
    With EquationReportItem
        .ID = LRU
        .Name = GetEquationReportTitle(LRU)
        .Value = ""
    End With
    EquationReportItems.Add EquationReportItem

    Set EquationReportItem = New clsEquationReportItem
    With EquationReportItem
        .ID = ATAChapter
        .Name = GetEquationReportTitle(ATAChapter)
        .Value = ""
    End With
    EquationReportItems.Add EquationReportItem


    Set EquationReportItem = New clsEquationReportItem
    With EquationReportItem
        .ID = EquID
        .Name = GetEquationReportTitle(EquID)
        .Value = ""
    End With
    EquationReportItems.Add EquationReportItem

    Set EquationReportItem = New clsEquationReportItem
    With EquationReportItem
        .ID = CompID
        .Name = GetEquationReportTitle(CompID)
        .Value = ""
    End With
    EquationReportItems.Add EquationReportItem

    Set EquationReportItem = New clsEquationReportItem
    With EquationReportItem
        .ID = Status
        .Name = GetEquationReportTitle(Status)
        .Value = ""
    End With
    EquationReportItems.Add EquationReportItem

    Set EquationReportItem = New clsEquationReportItem
    With EquationReportItem
        .ID = FaultLogged
        .Name = GetEquationReportTitle(FaultLogged)
        .Value = ""
    End With
    EquationReportItems.Add EquationReportItem

    Set EquationReportItem = New clsEquationReportItem
    With EquationReportItem
        .ID = FaultCode
        .Name = GetEquationReportTitle(FaultCode)
        .Value = ""
    End With
    EquationReportItems.Add EquationReportItem

    Set EquationReportItem = New clsEquationReportItem
    With EquationReportItem
        .ID = FaultLevel
        .Name = GetEquationReportTitle(FaultLevel)
        .Value = ""
    End With
    EquationReportItems.Add EquationReportItem

    Set EquationReportItem = New clsEquationReportItem
    With EquationReportItem
        .ID = MaintMessage
        .Name = GetEquationReportTitle(MaintMessage)
        .Value = ""
    End With
    EquationReportItems.Add EquationReportItem

    Set EquationReportItem = New clsEquationReportItem
    With EquationReportItem
        .ID = HelpText
        .Name = GetEquationReportTitle(HelpText)
        .Value = ""
    End With
    EquationReportItems.Add EquationReportItem

    Set EquationReportItem = New clsEquationReportItem
    With EquationReportItem
        .ID = FallingEdge
        .Name = GetEquationReportTitle(FallingEdge)
        .Value = ""
    End With
    EquationReportItems.Add EquationReportItem

    Set EquationReportItem = New clsEquationReportItem
    With EquationReportItem
        .ID = Logic
        .Name = GetEquationReportTitle(Logic)
        .Value = ""
    End With
    EquationReportItems.Add EquationReportItem

    Set EquationReportItem = New clsEquationReportItem
    With EquationReportItem
        .ID = FDE
        .Name = GetEquationReportTitle(FDE)
        .Value = ""
    End With
    EquationReportItems.Add EquationReportItem
End Sub

Private Sub Class_Terminate()
    Set EquationReportItems = Nothing
End Sub


