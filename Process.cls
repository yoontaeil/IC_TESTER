VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mvar_Expression As String       'Let으로 받은 문자열
Private mvar_TempExpression As String   '처리를 위한 임수 문자열
Dim Token As String, TypeOfToken As String
Dim ErrorFlag As Boolean

Public Function Execute(Optional Level As Integer = 1) As Double
    Dim MidVal As Double
    Dim FuncName As String, Arg1 As Double, Arg2 As Double, Arg3 As Double
        
    ErrorFlag = False
    On Error GoTo errExecute
    
    Call GetNextToken(Token, TypeOfToken)
    Select Case TypeOfToken
    Case "NUMERIC"
        MidVal = val(Token)
        Call GetNextToken(Token, TypeOfToken)
        
    Case "-"
        MidVal = -1 * Execute(1)
        
    Case "("
        MidVal = Execute(1)
        If Not Token = ")" Then
            err.Raise 1001, "Calculator", "수식에 이상이 발생되었습니다.": Exit Function
        End If
        Call GetNextToken(Token, TypeOfToken)
    End Select
    
    Do While True
        Select Case TypeOfToken
            Case "^": If Level <= 5 Then MidVal = MidVal ^ Execute(6): GoTo NextSymbol
            Case "*": If Level <= 4 Then MidVal = MidVal * Execute(5): GoTo NextSymbol
            Case "/": If Level <= 4 Then MidVal = MidVal / Execute(5): GoTo NextSymbol
            Case "\": If Level <= 3 Then MidVal = MidVal \ Execute(4): GoTo NextSymbol
            Case "mod": If Level <= 2 Then MidVal = MidVal Mod Execute(3): GoTo NextSymbol
            Case "+": If Level <= 1 Then MidVal = MidVal + Execute(2): GoTo NextSymbol
            Case "-": If Level <= 1 Then MidVal = MidVal - Execute(2): GoTo NextSymbol
        End Select
        Exit Do
NextSymbol:
    Loop
    
    If ErrorFlag = True Then
        Execute = -123.4567
    Else
        Execute = MidVal
    End If
    Exit Function
    
errExecute:
    'Err.Raise Err.Number, Err.Source, Err.Description
    ErrorFlag = True
    Execute = -123.4567
    Exit Function
End Function

Public Property Let Expression(vData As String)
    Dim Index As Long
    Dim OpenCNT As Integer, CloseCNT As Integer
    '// 괄호의 수를 확인한다.
    OpenCNT = 0: CloseCNT = 0
    For Index = 1 To Len(vData)
        If Mid(vData, Index, 1) = "(" Then OpenCNT = OpenCNT + 1
        If Mid(vData, Index, 1) = ")" Then CloseCNT = CloseCNT + 1
    Next
    
    If OpenCNT = CloseCNT Then
        mvar_Expression = vData: mvar_TempExpression = LCase(vData)
    Else
        MsgBox "Expression has error !", vbExclamation, "Process"
    End If
End Property

Public Property Get Expression() As String
    Expression = mvar_Expression
End Property

Private Sub GetNextToken(ByRef Token As String, ByRef TypeOfToken As String)
    
    On Error GoTo errGetNextToken
    Token = Space(0)
    TypeOfToken = Space(0)
    
    '// 앞뒤 여백을 없앤다.
    mvar_TempExpression = Trim(mvar_TempExpression)
    
    Select Case Left$(mvar_TempExpression, 1)
    Case "0" To "9", "."
        TypeOfToken = "NUMERIC"
        Token = Left$(mvar_TempExpression, 1)
        mvar_TempExpression = Mid$(mvar_TempExpression, 2)
        
        While (Left$(mvar_TempExpression, 1) >= "0" And Left$(mvar_TempExpression, 1) <= "9") Or Left$(mvar_TempExpression, 1) = "."
            Token = Token + Left$(mvar_TempExpression, 1)
            mvar_TempExpression = Mid$(mvar_TempExpression, 2)
        Wend
        
    Case "a" To "z"
        TypeOfToken = "FUNCTION"
        Token = Left$(mvar_TempExpression, 1)
        mvar_TempExpression = Mid$(mvar_TempExpression, 2)
        
        While Left$(mvar_TempExpression, 1) >= "a" And Left$(mvar_TempExpression, 1) <= "z"
            Token = Token + Left$(mvar_TempExpression, 1)
            mvar_TempExpression = Mid$(mvar_TempExpression, 2)
        Wend
    
    Case Else
        Token = Left$(mvar_TempExpression, 1): TypeOfToken = Token
        mvar_TempExpression = Mid$(mvar_TempExpression, 2)
        
    End Select
    
    If TypeOfToken = "NUMERIC" Then
        If IsNumeric(Token) = False Then
            err.Raise 1001, "Calculator", "수식에 이상이 발생되었습니다.": Exit Sub
        End If
    End If
    Exit Sub
    
errGetNextToken:
    err.Raise err.Number, err.Source, err.Description
    Resume Next
End Sub


