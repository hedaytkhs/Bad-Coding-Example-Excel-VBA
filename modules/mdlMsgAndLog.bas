Attribute VB_Name = "mdlMsgAndLog"

'****************************************************************************************
'
'    Application名、表示メッセージ等の定数取得用のモジュール
'
'    列挙型に定義した名称をオートコンプリートを利用してプログラム作成を省力化及び可読性を高めるために用いる
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
        Case APP_CNST_ID.AppName: sRet = "EquationReportをExcelにインポートする"
        Case APP_CNST_ID.ErrNumber: sRet = vbCrLf & "ErrNumber: "
        Case APP_CNST_ID.ErrDescription: sRet = vbCrLf & "ErrDescription: "
        Case APP_CNST_ID.msgTextFileNotSpecified: sRet = "テキストファイルが指定されませんでした." & vbCrLf & "EquationReportのインポートを中止します."
        Case APP_CNST_ID.msgPreProcessFailed: sRet = "前処理に失敗しました." & vbCrLf & "EquationReportのインポートを中止します."
    
        Case APP_CNST_ID.msgImportCompleted: sRet = "EquationReportをExcelにインポートしました."
        Case APP_CNST_ID.msgImportFailed: sRet = "EquationReportのインポートに失敗しました."
        Case APP_CNST_ID.FileSelectTextFilter: sRet = "Equation Report テキストファイル,*.txt"
        Case APP_CNST_ID.FileSelectTitle: sRet = "Equation Report テキストファイルを選択してください"
        Case APP_CNST_ID.msgShowLogFile: sRet = "エラー・ログを表示しますか?"
    End Select
    
    GetAppCNST = sRet
End Function

