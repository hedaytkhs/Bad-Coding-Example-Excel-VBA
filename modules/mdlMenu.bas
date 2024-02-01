Attribute VB_Name = "mdlMenu"
'****************************************************************************************
'
'    Equation Reportの要素を操作するためのクラス
'
'    Equation Reportの項目名、値など共通要素をまとめる
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************
Option Explicit


'---------------------------------------------------------------------------------
'Applicationの状態管理用のユーザー定義型
'---------------------------------------------------------------------------------
Public Type tpArgument
    EquationReportTextFilePath As String
    IsCancelled As Boolean
    IsCompleted As Boolean
    LogText As String
    ErrNumber As Long
    ErrDescription As String
    LogFilePath As String
    LogFileCreated As Boolean
    objTs As TextStream
End Type

'Applicationの状態管理用のユーザー定義型を各関数の引数、戻り値に使用する
'Main関数内でこのユーザー定義型の変数およびその初期値を設定し、各関数の引数、戻り値に必ず設定する
'なるべく引数で値を渡すようにし、変数のスコープをむやみに広げないようにするために使用する

Sub EquationReportをExcelにインポートする(ByRef myButton As IRibbonControl)
#If (DEBUG_MODE = 1) Then
    MsgBox "作成中です.", vbInformation + vbOKOnly, CNST_APP_NAME
#Else
    Call Main
#End If
End Sub

Private Sub Main()
    Dim iRet As Integer
    Dim ImportEquationReport As tpArgument
    With ImportEquationReport
        .IsCancelled = False
        .IsCompleted = False
        .LogText = "Function: Main()" & vbCrLf
        .EquationReportTextFilePath = ""
        .ErrNumber = 0
        .ErrDescription = ""
    End With
    
    ImportEquationReport = GetEquationReportTextFilePath(ImportEquationReport)
    If ImportEquationReport.IsCancelled Then
        'ログファイル作成
        ImportEquationReport.LogText = ImportEquationReport.LogText & vbCrLf & GetAppCNST(msgTextFileNotSpecified)
        iRet = MsgBox(GetAppCNST(msgTextFileNotSpecified) & vbCrLf & GetAppCNST(msgShowLogFile), vbExclamation + vbYesNoCancel, GetAppCNST(AppName))
        If iRet = vbYes Then
            'ログファイル表示
        End If
        Exit Sub
    End If
    
    
    ImportEquationReport = PreProcess(ImportEquationReport)
    If ImportEquationReport.IsCancelled Then
        'ログファイル作成
        ImportEquationReport.LogText = ImportEquationReport.LogText & vbCrLf & GetAppCNST(msgPreProcessFailed)
        iRet = MsgBox(GetAppCNST(msgPreProcessFailed) & vbCrLf & GetAppCNST(msgShowLogFile), vbExclamation + vbYesNoCancel, GetAppCNST(AppName))
        If iRet = vbYes Then
            'ログファイル表示
        End If
        Exit Sub
    End If
    
    ImportEquationReport = EquationReportImportToExcelSheet(ImportEquationReport)
    If ImportEquationReport.IsCompleted Then
        MsgBox GetAppCNST(msgImportCompleted), vbInformation + vbOKOnly, GetAppCNST(AppName)
    Else
        iRet = MsgBox(GetAppCNST(msgImportFailed) & vbCrLf & GetAppCNST(msgShowLogFile), vbExclamation + vbYesNoCancel, GetAppCNST(AppName))
        If iRet = vbYes Then
            'ログファイル表示
        End If
    End If
End Sub

Private Function PreProcess(ByRef AppArguments As tpArgument) As tpArgument
Const CNST_FUNCTION_NAME As String = "Function: PreProcess"
On Error GoTo errHandler
    
    With PreProcess
        .IsCancelled = False
        .IsCompleted = False
        .LogText = AppArguments.LogText & CNST_FUNCTION_NAME & vbCrLf
        .ErrNumber = CLng(Err.Number)
        .ErrDescription = Err.Description
    End With
    
    If True Then
    
    Else
    
    End If
    
    
    
    Exit Function
errHandler:
    With PreProcess
        .IsCancelled = True
        .IsCompleted = False
        .ErrNumber = CLng(Err.Number)
        .ErrDescription = Err.Description
        .LogText = AppArguments.LogText & CNST_FUNCTION_NAME & vbCrLf & GetAppCNST(ErrNumber) & Str(.ErrNumber) & GetAppCNST(ErrDescription) & ErrDescription
    End With
    MsgBox Err.Number & ":" & Err.Description
End Function

Private Function EquationReportImportToExcelSheet(ByRef AppArguments As tpArgument) As tpArgument
Const CNST_FUNCTION_NAME As String = "Function: EquationReportImportToExcelSheet"
On Error GoTo errHandler
    
    With EquationReportImportToExcelSheet
        .IsCancelled = False
        .LogText = AppArguments.LogText & CNST_FUNCTION_NAME & vbCrLf
    End With
    
    'インポート先のシートを指定
'    Set oSheet = ActiveSheet
    Dim XlsSheetEquationReport As New clsXlsSheetEquationReport
    
    Dim objEqReportLine As New clsEqReportLine
   
    
    '前処理後のテキストファイルを開く
    
    Dim fso As New FileSystemObject
    Dim trgTs As TextStream
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    Set trgTs = fso.CreateTextFile(AppArguments.EquationReportTextFilePath, ForReading)
    
    
'    objEqReportLine.Text = trgTs.ReadLine
    
    
    
    
    trgTs.Close
    
    With EquationReportImportToExcelSheet
        .IsCompleted = True
        .LogText = AppArguments.LogText & "Function: EquationReportImportToExcelSheet" & vbCrLf
    End With
    
    Set trgTs = Nothing
    Set fso = Nothing
    

    Exit Function
errHandler:
    With EquationReportImportToExcelSheet
        .IsCancelled = True
        .IsCompleted = False
        .LogText = AppArguments.LogText & CNST_FUNCTION_NAME & vbCrLf
        .ErrNumber = CLng(Err.Number)
        .ErrDescription = Err.Description
        .LogText = AppArguments.LogText & CNST_FUNCTION_NAME & vbCrLf & GetAppCNST(ErrNumber) & Str(.ErrNumber) & GetAppCNST(ErrDescription) & ErrDescription
    End With
    MsgBox Err.Number & ":" & Err.Description
End Function


Private Function GetEquationReportTextFilePath(ByRef AppArguments As tpArgument) As tpArgument
Const CNST_FUNCTION_NAME As String = "Function: GetEquationReportTextFilePath"
On Error GoTo errHandler
    Dim sFilePath As Variant
    
    sFilePath = _
        Application.GetOpenFilename( _
             FileFilter:=GetAppCNST(FileSelectTextFilter) _
             , FilterIndex:=1 _
           , Title:=GetAppCNST(FileSelectTitle) _
           , MultiSelect:=False _
            )
    
    With GetEquationReportTextFilePath
        .EquationReportTextFilePath = sFilePath
        If UCase(sFilePath) = "FALSE" Then
            .IsCancelled = True
        Else
            .IsCancelled = False
        End If
        .LogText = AppArguments.LogText & CNST_FUNCTION_NAME & vbCrLf & sFilePath & vbCrLf
    End With
    
    Exit Function
errHandler:
    MsgBox Err.Number & ":" & Err.Description
    With GetEquationReportTextFilePath
        .IsCancelled = True
        .IsCompleted = False
        .ErrNumber = CLng(Err.Number)
        .ErrDescription = Err.Description
        .LogText = AppArguments.LogText & CNST_FUNCTION_NAME & vbCrLf & GetAppCNST(ErrNumber) & Str(.ErrNumber) & GetAppCNST(ErrDescription) & ErrDescription
    End With
End Function




