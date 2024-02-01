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
    '判定する
End Property

Property Get Value() As String
    Value = Value_
End Property

'判定する
'Itemの種類分け、状態フラグ、ValueにReportItemの内容を格納する





'EquationReportItemの場合は、EquationReportItemIDを設定


'Groupの判定
' ":"で区切られる文字がある。かつ、:の前の文字列は ATAもしくはLRUである。

'EquationReportItemの判定
'vbTab & "="で区切られる文字がある。かつ、=の前の文字列は、Equation Report Titleのいずれかの文字列と一致する

'Help Textが開始されたらフラグを立てる
'Logicが開始されたらフラグを立てる
'


' "Possible Causes (LRU Short Name)"が含まれる
' Valueを取り出す

'Possible Causes Listの開始行である

'Variable    Variable Detailsの開始行である

' Variableの途中の行である場合
' VariableとVariable Detailsに分離して VariableとVariable DetailsそれぞれをコレクションにAddItemする


'異なるEquation Nameが出てきたら、clsFaultEquationの現在のインスタンスから､Excelの1行に転記する

'clsFaultEquationの次のインスタンスを準備する
' 次のEquation Nameの値をセットする





Private Sub Class_Initialize()
    IsGroup_ = False
    IsEquationReportItem_ = False
End Sub
