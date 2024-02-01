VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************
'
'    ログファイルを操作するためのクラス
'
'    ログファイルのファイル名の取得
'    ログの追記
'
'    MRJ Technical Publication Tool
'
'    Hideaki Takahashi
'
'****************************************************************************************
Option Explicit

Dim LogFileName_ As String
Dim Folder_ As String
Dim LogFilePath_ As String

Private Function GetLogFileName(sOriginalFileName As String) As String
    Dim wRet As String
    Dim sDate As String
    Dim sTime As String
    sDate = Replace(Format(Now, "Short Date"), "/", "")
    sTime = Replace(Format(Now, "Short Time"), ":", "")

    wRet = sOriginalFileName & "_" & sDate & "_" & sTime & ".log"
    
'    LogFileName_ = wRet
    GetLogFileName = wRet
End Function

Public Function CreateLogFile(ByRef AppArguments As tpArgument, ByVal sFolder As String, ByVal sOriginalFileName As String) As tpArgument
Const CNST_FUNCTION_NAME As String = "Function: CreateLogFile"
On Error GoTo errHandler
    
    LogFileName_ = GetLogFileName(sOriginalFileName)
    
    'FileSystemObject
    'ログファイルを作成
    
    'テキストストリームを開いて渡す
    
    With CreateLogFile
        .LogFileCreated = True
        .LogFilePath = wRet
    End With
    
    
    Exit Function
errHandler:
    With CreateLogFile
        .IsCancelled = True
        .IsCompleted = False
        .LogFileCreated = False
        .ErrNumber = CLng(Err.Number)
        .ErrDescription = Err.Description
        .LogText = AppArguments.LogText & CNST_FUNCTION_NAME & vbCrLf & GetAppCNST(ErrNumber) & Str(.ErrNumber) & GetAppCNST(ErrDescription) & ErrDescription
    End With
End Function

