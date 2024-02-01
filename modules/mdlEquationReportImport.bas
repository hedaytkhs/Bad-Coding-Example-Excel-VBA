Attribute VB_Name = "mdlEquationReportImport"
Option Explicit

'Enum EquationReportGroup
'    ATA = 1
'    LRU = 2
'End Enum

Enum EquationReportItem
    GroupATA = 1
    GroupLRU = 2
    EquationName = 3
    Priority = 4
    LRU = 5
    ATAChapter = 6
    EquID = 7

    CompID = 8
    Status = 9
    FaultLogged = 10
    FaultCode = 11
    FaultLevel = 12

    MaintMessage = 13
    RisingEdge = 14
    HelpText = 16
    FallingEdge = 15
    PossibleCausedList = 17

    Logic = 18
    FDE = 19
    PossibleCausedShortName = 10
    Variable = 21
    VariableDetails = 22
End Enum

Public Type tpEquationVariable
    Variable As String
    VariableDetails As String
End Type

Public Function GetEquationReportTitle(EquationReportID As EquationReportItem) As String

    Dim sRet As String
    Select Case EquationReportID
        Case EquationReportItem.GroupATA: sRet = "ATA"
        Case EquationReportItem.GroupLRU: sRet = "LRU"
        
        Case EquationReportItem.EquationName: sRet = "Equation Name"
        Case EquationReportItem.Priority: sRet = "Priority"
        Case EquationReportItem.LRU: sRet = "LRU"
        Case EquationReportItem.ATAChapter: sRet = "ATA Chapter"
        Case EquationReportItem.EquID: sRet = "Equ ID#"

        Case EquationReportItem.CompID: sRet = "Comp ID"
        Case EquationReportItem.Status: sRet = "Status"
        Case EquationReportItem.FaultLogged: sRet = "Fault Logged"
        Case EquationReportItem.FaultCode: sRet = "Fault Code"
        Case EquationReportItem.FaultLevel: sRet = "Fault Level"

        Case EquationReportItem.MaintMessage: sRet = "Maint Message"
        Case EquationReportItem.RisingEdge: sRet = "Rising Edge"
        Case EquationReportItem.HelpText: sRet = "Help Text"
        Case EquationReportItem.FallingEdge: sRet = "Falling Edge"
        Case EquationReportItem.PossibleCausedList: sRet = "Possible Causes List"

        Case EquationReportItem.Logic: sRet = "Logic"
        Case EquationReportItem.FDE: sRet = "FDEs"
        Case EquationReportItem.PossibleCausedShortName: sRet = "Possible Causes (LRU Short Name)"
        Case EquationReportItem.Variable: sRet = "Variable"
        Case EquationReportItem.VariableDetails: sRet = ""
    End Select
    
    GetEquationReportTitle = sRet
End Function
