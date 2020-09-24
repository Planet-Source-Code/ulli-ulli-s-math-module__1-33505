VERSION 5.00
Begin VB.UserControl Evaluator 
   Appearance      =   0  '2D
   BackColor       =   &H80000005&
   CanGetFocus     =   0   'False
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   ForwardFocus    =   -1  'True
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   PropertyPages   =   "Evaluator.ctx":0000
   ScaleHeight     =   255
   ScaleWidth      =   420
   ToolboxBitmap   =   "Evaluator.ctx":0016
   Windowless      =   -1  'True
   Begin VB.Image img 
      Height          =   240
      Left            =   0
      Picture         =   "Evaluator.ctx":0110
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "Evaluator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'Constants
Private Const BracketsMismatch  As String = "206Brackets are not paired"
Private Const ExtraChars        As String = "207Extra characters found"
Private Const OperandMissing    As String = "208Missing operand or argument"
Private Const StackLow          As String = "209Not enough items on stack"
Private Const NoNesting         As String = "211Integral and differential calculations cannot be nested"
Private Const NoIntervals       As String = "213Number of intervals cannot be zero for Integration"
Private Const TooComplex        As String = "028Formula too complex or singularity in formula"
Private Const NoPrevResult      As String = "010No previous result available"
Private Const ValueMissing      As String = "214No value for token '"
Private Const TooFewPrimes      As String = "215Cannot return result - size of prime table is insufficient"
Private Const PropLocked        As String = "216Property is locked at runtime"
Private Const PropFormula       As String = "Formula"
Private Const PropPTS           As String = "PrimeTableSize"
Private Const PTS               As Long = 500 'default prime table size

'Property Variables
Private dblMyResult         As Double
Private dblMyPreviousResult As Double
Private strMyFormula        As String
Private lngMyUBPrime        As Long

'Working Variables
Private colStck             As Collection
Private varData1            As Variant
Private varData2            As Variant
Private dblTemp1            As Double
Private dblTemp2            As Double
Private dblIntegralVar      As Double
Private lngDepth            As Long
Private lngIfActiveDepth    As Long
Private lngPrecedence       As Long
Private lngPtr1             As Long
Private lngPtr2             As Long
Private lngPopPending       As Long
Private lngPrime()          As Long
Private lngXPrime           As Long
Private strWord             As String
Private strChar             As String
Private strDPFrom           As String
Private strDPTo             As String
Private bolNextUnary        As Boolean
Private bolIsBinary         As Boolean
Private bolResetDataPending As Boolean
Private bolIntgDiff         As Boolean
Private bolPreviousValid    As Boolean
Private bolPrimesPresent    As Boolean
Private bolSuccess          As Boolean
Private bolShowProgress     As Boolean

'Event Declarations
Public Event QueryToken(ByVal Token As String, Value As String)
Public Event Plot(ByVal x As Double, ByVal y As Double)

Private Function Compact(Fml As String) As String

  Dim lngPtr   As Long
  Dim lngBckt  As Long
  Dim str1     As String
  Dim str2     As String

    For lngPtr = 1 To Len(Fml)
        str1 = Mid$(Fml, lngPtr, 1)
        lngBckt = lngBckt + (str1 = "(") - (str1 = ")")       'count brackets
        Select Case Right$(str2, 1) & str1
          Case "--"
            Mid$(str2, Len(str2), 1) = "+"                    'replace -- by +
            str1 = ""
          Case "+-"
            Mid$(str2, Len(str2), 1) = "-"                    'toggle
            str1 = ""
          Case "++", "-+"
            str1 = ""                                         'ignore +
        End Select
        str1 = IIf(str1 = strDPFrom, strDPTo, str1)           'replace . by , (or vice versa)
        str1 = IIf(str1 = " ", "", str1)                      'strip spaces
        str2 = str2 & str1
    Next lngPtr
    If lngBckt <> 0 Then
        Err.Raise Val(BracketsMismatch), Ambient.DisplayName, Mid$(BracketsMismatch, 4)
    End If
    Do
        Select Case Left$(str2, 1)
          Case "("
            If FindMatchingBracket(str2) = Len(str2) Then     'strip outer brackets
                str2 = Mid$(str2, 2, Len(str2) - 2)
              Else 'NOT FINDMATCHINGBRACKET(STR2)...
                Exit Do '>---> Loop
            End If
          Case "+"                                            'strip leading +
            str2 = Mid$(str2, 2)
          Case Else
            Exit Do '>---> Loop
        End Select
    Loop
    Compact = str2

End Function

Private Function Compute(Formula As String) As Double

  '''''''''''''''''''''''''''''''''
  'This is where it's all happening
  '''''''''''''''''''''''''''''''''

  'recursive variables - we dont want too many, so some of them are 'misused'

  Dim strFormula     As String
  Dim strLeftPart    As String
  Dim strRitePart    As String
  Dim lngSplitAt     As Long
  Dim dblLeftResult  As Double
  Dim dblRiteResult  As Double
  Dim dblForVar      As Double
  Dim dblResult      As Double

    If lngDepth > 255 Then                 'limit calculation depth just in case
        Err.Raise Val(TooComplex), Ambient.DisplayName, Mid$(TooComplex, 4)
    End If
    lngDepth = lngDepth + 1
    strFormula = Compact(Formula & " ")
    If Len(strFormula) Or lngDepth > 1 Then
        bolNextUnary = True
        lngSplitAt = 0
        lngPrecedence = 0
        For lngPtr1 = 1 To Len(strFormula)
            bolIsBinary = True
            Select Case Mid$(strFormula, lngPtr1, 1)
              Case "("                       'skip to end of whatever is in the brackets
                lngPtr1 = lngPtr1 + FindMatchingBracket(Mid$(strFormula, lngPtr1)) - 1 ':( Modifies active For-Variable
                bolNextUnary = False
              Case "|", ";"
                If lngPrecedence <= 21 Then  'precedence is relative
                    lngPrecedence = 21
                    lngSplitAt = lngPtr1
                End If
                bolNextUnary = True
              Case "&"
                If lngPrecedence <= 20 Then
                    lngPrecedence = 20
                    lngSplitAt = lngPtr1
                End If
                bolNextUnary = True
              Case ":"                        'scale to 0 thru 1
                If lngPrecedence <= 19 Then
                    lngPrecedence = 19
                    lngSplitAt = lngPtr1
                End If
                bolNextUnary = True
              Case "<", "=", ">", "{", "}"   'less equal greater min max
                If lngPrecedence <= 18 Then
                    lngPrecedence = 18
                    lngSplitAt = lngPtr1
                End If
                bolNextUnary = True
              Case "+", "-"                   'plus minus
                If Not bolNextUnary Then
                    If lngPrecedence <= 17 Then
                        lngPrecedence = 17
                        lngSplitAt = lngPtr1
                    End If
                    bolNextUnary = True
                End If
              Case "e", "E"                   'scientific notation
                bolNextUnary = False
                If lngPtr1 > 1 Then
                    bolNextUnary = IsNumeric(Mid$(strFormula, lngPtr1 - 1, 1))
                End If
              Case "*", "/", "\", "]"         'multiply divide integerdivide modulo
                If lngPrecedence <= 16 Then
                    lngPrecedence = 16
                    lngSplitAt = lngPtr1
                End If
                bolNextUnary = True
              Case "^"                        'exponentiation
                If lngPrecedence <= 15 Then
                    lngPrecedence = 15
                    lngSplitAt = lngPtr1
                End If
                bolNextUnary = True
              Case "#", "["                   'logarithm root
                If lngPrecedence <= 14 Then
                    lngPrecedence = 14
                    lngSplitAt = lngPtr1
                End If
                bolNextUnary = True
              Case "!", "%", "°", "²", "³", "'", """"  'postoperator
                If lngPrecedence <= 1 Then
                    lngPrecedence = 1                  'highest precedence
                    lngSplitAt = lngPtr1
                    bolIsBinary = False
                End If
                bolNextUnary = False
              Case Else
                bolNextUnary = False
            End Select
        Next lngPtr1
        If lngSplitAt Then
            'binary operation or postoperator
            strLeftPart = Left$(strFormula, lngSplitAt - 1)
            strRitePart = Mid$(strFormula, lngSplitAt + 1)
            If bolIsBinary Then                        'binary operator
                If Len(strLeftPart) Then
                    dblLeftResult = Compute(strLeftPart)
                  Else 'LEN(STRLEFTPART) = FALSE
                    Err.Raise Val(OperandMissing), Ambient.DisplayName, Mid$(OperandMissing, 4)
                End If
                If Len(strRitePart) Then
                    dblRiteResult = Compute(strRitePart)
                  Else 'LEN(STRRITEPART) = FALSE
                    Err.Raise Val(OperandMissing), Ambient.DisplayName, Mid$(OperandMissing, 4)
                End If
              Else 'BOLISBINARY = FALSE
                If Len(strLeftPart) Then
                    dblLeftResult = Compute(strLeftPart)
                  Else 'LEN(STRLEFTPART) = FALSE
                    Err.Raise Val(OperandMissing), Ambient.DisplayName, Mid$(OperandMissing, 4)
                End If
                If Len(strRitePart) Then
                    Err.Raise Val(ExtraChars), Ambient.DisplayName, Mid$(ExtraChars, 4)
                End If
            End If
            Select Case Mid$(strFormula, lngSplitAt, 1)
              Case "+"
                dblResult = dblLeftResult + dblRiteResult
              Case "-"
                dblResult = dblLeftResult - dblRiteResult
              Case "*", "&"
                dblResult = dblLeftResult * dblRiteResult
              Case "/"
                dblResult = dblLeftResult / dblRiteResult
              Case "\"
                dblResult = dblLeftResult \ dblRiteResult
              Case "]"
                dblResult = dblLeftResult Mod dblRiteResult
              Case ":"
                If dblRiteResult < 0 Then
                    If dblLeftResult <= dblRiteResult Then
                        dblResult = 1
                      ElseIf dblLeftResult > 0 Then 'NOT DBLLEFTRESULT...
                        dblResult = 0
                      Else 'NOT DBLLEFTRESULT...
                        dblResult = Abs(dblLeftResult / dblRiteResult)
                    End If
                  Else 'NOT DBLRITERESULT...
                    If dblLeftResult >= dblRiteResult Then
                        dblResult = 1
                      ElseIf dblLeftResult < 0 Then 'NOT DBLLEFTRESULT...
                        dblResult = 0
                      Else 'NOT DBLLEFTRESULT...
                        dblResult = Abs(dblLeftResult / dblRiteResult)
                    End If
                End If
              Case "|"
                dblResult = 1 - (1 - dblLeftResult) * (1 - dblRiteResult)
              Case ";"
                dblResult = ((dblLeftResult = 0 And dblRiteResult <> 0) Or (dblLeftResult <> 0 And dblRiteResult = 0))
              Case "{"
                dblResult = IIf(dblLeftResult < dblRiteResult, dblLeftResult, dblRiteResult)
              Case "}"
                dblResult = IIf(dblLeftResult > dblRiteResult, dblLeftResult, dblRiteResult)
              Case "^"   'power        eg 3^4 = 81
                'get around bug in VB6 (odd roots of negative numbers)
                If dblLeftResult >= 0 Then
                    dblResult = dblLeftResult ^ dblRiteResult
                  Else 'NOT DBLLEFTRESULT...
                    dblResult = 1 / dblRiteResult
                    If dblResult = Int(dblResult) And dblResult And 1 Then
                        dblResult = -((-dblLeftResult) ^ dblRiteResult)
                      Else 'NOT DBLRESULT...
                        dblResult = dblLeftResult ^ dblRiteResult
                    End If
                End If
                lngPtr1 = 0
              Case "["   'n-th Root    eg 3[64 = 4    -:-   4^3=64
                'uses bug circumvention of exponentiation
                dblResult = Compute(strRitePart & "^(1/" & strLeftPart & ")")
              Case "#"   'n log x      eg 4#81 = 3    -:-   3^4=81
                dblResult = Log(dblRiteResult) / Log(dblLeftResult)
              Case "="
                dblResult = (dblLeftResult = dblRiteResult)
              Case "<"
                dblResult = (dblLeftResult < dblRiteResult)
              Case ">"
                dblResult = (dblLeftResult > dblRiteResult)
              Case "!", "%", "°", "²", "³", "'", """"            'postoperator
                Select Case Mid$(strFormula, lngSplitAt, 1)
                  Case "!"
                    dblResult = Compute("fac(" & strLeftPart & ")")
                  Case "%"
                    dblResult = dblLeftResult / 100
                  Case "²"
                    dblResult = dblLeftResult ^ 2
                  Case "³"
                    dblResult = dblLeftResult ^ 3
                  Case "°"
                    dblResult = dblLeftResult * Atn(1) / 45      'angle - degrees
                  Case "'"
                    dblResult = dblLeftResult * Atn(1) / 2700    'angle - minutes
                  Case Else  '"
                    dblResult = dblLeftResult * Atn(1) / 162000  'angle - seconds
                End Select
            End Select
          Else                                     'not binary operation nor postoperator'LNGSPLITAT = FALSE
            strWord = ""
            For lngPtr1 = 1 To Len(strFormula)
                strChar = Mid$(strFormula, lngPtr1, 1)
                If strChar = "(" Then              'found a bracket so it's a function
                    Exit For                       'function name (possibly -FuncName) is in strWord '>---> Next
                  Else 'NOT STRCHAR...
                    strWord = strWord & strChar    'continue
                End If
            Next lngPtr1
            If Left$(strWord, 1) = "-" Then        'something is negative like -->   -sin(30°)
                dblResult = -Compute(Mid$(strFormula, 2))
              ElseIf lngPtr1 > Len(strFormula) Then  'not a function'NOT LEFT$(STRWORD,...
                If IsNumeric(strWord) Then
                    dblResult = CDbl(strWord)       'convert to numeric
                  Else                              'may be a token or the variant of diff or intg..'ISNUMERIC(STRWORD) = FALSE
                    Select Case LCase(strWord)      '..or one of the constants or previous result
                      Case "x"                      'may be diff or intg
                        If bolIntgDiff Then         'yep - do not handle this as token
                            dblResult = dblIntegralVar
                          Else 'BOLINTGDIFF = FALSE
                            dblResult = Compute(TranslateToken(strWord, "0"))
                        End If
                      Case "pr"
                        If bolPreviousValid Then
                            dblResult = dblMyPreviousResult
                          Else 'BOLPREVIOUSVALID = FALSE
                            Err.Raise Val(NoPrevResult), Ambient.DisplayName, Mid$(NoPrevResult, 4)
                        End If
                      Case "pi"
                        dblResult = 4 * Atn(1)
                      Case "e"
                        dblResult = Exp(1)
                      Case Else
                        dblResult = Compute(TranslateToken(strWord, "0"))
                    End Select
                End If
              Else                                   'it is a function'NOT LNGPTR1...
                strWord = LCase$(strWord)            'go ahead and scan the function names
                Select Case strWord
                  Case "harccosec"
                    dblResult = Compute(Mid$(strFormula, 10))
                    dblResult = Log((Sgn(dblResult) * Sqr(dblResult * dblResult + 1) + 1) / dblResult)
                  Case "element", "stack"
                    dblResult = Int(Compute(Mid$(strFormula, Len(strWord) + 1)))
                    If dblResult < 1 Or dblResult > colStck.Count Then
                        Err.Raise Val(StackLow), Ambient.DisplayName, Mid$(StackLow, 4)
                      Else 'NOT DBLRESULT...
                        dblResult = colStck.Item(dblResult)
                    End If
                  Case "arccosec"
                    dblResult = Compute(Mid$(strFormula, 9))
                    dblResult = Atn(dblResult / Sqr(dblResult * dblResult - 1)) + (Sgn(dblResult) - 1) * 2 * Atn(1)
                  Case "harcsin"
                    dblResult = Compute(Mid$(strFormula, 8))
                    dblResult = Log(dblResult + Sqr(dblResult * dblResult + 1))
                  Case "harccos"
                    dblResult = Compute(Mid$(strFormula, 8))
                    dblResult = Log(dblResult + Sqr(dblResult * dblResult - 1))
                  Case "harctan"
                    dblResult = Compute(Mid$(strFormula, 8))
                    dblResult = Log((1 + dblResult) / (1 - dblResult)) / 2
                  Case "harccot"
                    dblResult = Compute(Mid$(strFormula, 8))
                    dblResult = Log((dblResult + 1) / (dblResult - 1)) / 2
                  Case "harcsec"
                    dblResult = Compute(Mid$(strFormula, 8))
                    dblResult = Log((Sqr(-dblResult * dblResult + 1) + 1) / dblResult)
                  Case "arcsin"
                    dblResult = Compute(Mid$(strFormula, 7))
                    dblResult = Atn(dblResult / Sqr(-dblResult * dblResult + 1))
                  Case "arccos"
                    dblResult = Compute(Mid$(strFormula, 7))
                    dblResult = Compute("ArcSin(-" & Format$(dblResult) & ") + 2 * Atn(1)")
                  Case "arcsec"
                    dblResult = Compute(Mid$(strFormula, 7))
                    dblResult = Atn(dblResult / Sqr(dblResult * dblResult - 1)) + Sgn(Sgn(dblResult) - 1) * 2 * Atn(1)
                  Case "arccot"
                    dblResult = Compute(Mid$(strFormula, 7))
                    dblResult = Atn(dblResult) + 2 * Atn(1)
                  Case "popall"
                    dblResult = Compute("pop()")
                    lngPopPending = lngPopPending Or 2
                  Case "bottom"
                    If colStck.Count = 0 Then
                        Err.Raise Val(StackLow), Ambient.DisplayName, Mid$(StackLow, 4)
                      Else 'NOT COLSTCK.COUNT...
                        dblResult = colStck.Item(colStck.Count)
                    End If
                  Case "cosec"
                    dblResult = 1 / Sin(Compute(Mid$(strFormula, 6)))
                  Case "notin"
                    dblResult = 1 - Compute(Mid$(strFormula, 6))
                  Case "prim" 'returns first prime >= argument
                    GeneratePrimes
                    dblResult = Abs(Compute(Mid$(strFormula, 5)))
                    If dblResult <= 2 Then
                        dblResult = 2
                      ElseIf dblResult <= 3 Then 'NOT DBLRESULT...
                        dblResult = 3
                      Else 'NOT DBLRESULT...
                        dblResult = Prime(Int(dblResult + 0.5))
                    End If
                  Case "fact" 'push all factors of number on stack and return number of factors
                    GeneratePrimes
                    dblRiteResult = Fix(Compute(Mid$(strFormula, 5)))
                    If dblRiteResult < 2 And dblRiteResult > -2 Then
                        dblResult = 0
                      Else 'NOT DBLRITERESULT...
                        dblResult = 1
                        For lngPtr1 = lngMyUBPrime To 0 Step -1
                            dblLeftResult = dblRiteResult / lngPrime(lngPtr1)
                            If dblLeftResult = Int(dblLeftResult) Then
                                If colStck.Count = 0 Then
                                    colStck.Add lngPrime(lngPtr1)
                                  Else 'NOT COLSTCK.COUNT...
                                    colStck.Add lngPrime(lngPtr1), , 1
                                End If
                                dblResult = dblResult + 1 'number of factors
                                If Abs(dblLeftResult) = 1 Then
                                    Exit For '>---> Next
                                End If
                                dblRiteResult = dblLeftResult
                                lngPtr1 = lngPtr1 + 1 'try same factor again':( Modifies active For-Variable
                            End If
                        Next lngPtr1
                        If colStck.Count = 0 Then
                            colStck.Add dblLeftResult
                          Else 'NOT COLSTCK.COUNT...
                            colStck.Add dblLeftResult, , 1
                        End If
                        If dblLeftResult <> Sgn(dblLeftResult) Then 'no success
                            Do 'clear all pushed items off the stack
                                dblResult = dblResult - 1
                                colStck.Remove 1
                            Loop Until dblResult = 0
                            Err.Raise Val(TooFewPrimes), Ambient.DisplayName, Mid$(TooFewPrimes, 4)
                        End If
                    End If
                  Case "hsin"
                    dblResult = Compute(Mid$(strFormula, 5))
                    dblResult = (Exp(dblResult) - Exp(-dblResult)) / 2
                  Case "hcos"
                    dblResult = Compute(Mid$(strFormula, 5))
                    dblResult = (Exp(dblResult) + Exp(-dblResult)) / 2
                  Case "htan"
                    dblResult = Compute(Mid$(strFormula, 5))
                    dblResult = (Exp(dblResult) - Exp(-dblResult)) / (Exp(dblResult) + Exp(-dblResult))
                  Case "hsec"
                    dblResult = Compute(Mid$(strFormula, 5))
                    dblResult = 2 / (Exp(dblResult) + Exp(-dblResult))
                  Case "hcot"
                    dblResult = Compute(Mid$(strFormula, 5))
                    dblResult = (Exp(dblResult) + Exp(-dblResult)) / (Exp(dblResult) - Exp(-dblResult))
                  Case "diff" 'differentiation
                    If bolIntgDiff Then 'cannot nest diff
                        Err.Raise Val(NoNesting), Ambient.DisplayName, Mid$(NoNesting, 4)
                      Else 'BOLINTGDIFF = FALSE
                        If colStck.Count < 1 Then
                            Err.Raise Val(StackLow), Ambient.DisplayName, Mid$(StackLow, 4)
                          Else 'NOT COLSTCK.COUNT...
                            bolIntgDiff = True
                            dblIntegralVar = colStck.Item(1) - (2# ^ -13)
                            dblLeftResult = Compute(Mid$(strFormula, 5))
                            dblIntegralVar = colStck.Item(1) + (2# ^ -13)
                            dblRiteResult = Compute(Mid$(strFormula, 5))
                            dblResult = dblRiteResult / (2# ^ -12) - dblLeftResult / (2# ^ -12)
                            bolIntgDiff = False
                        End If
                    End If
                  Case "intg" 'uses Simpson's parabolic approximation
                    If bolIntgDiff Then 'cannot nest intg
                        Err.Raise Val(NoNesting), Ambient.DisplayName, Mid$(NoNesting, 4)
                      Else 'BOLINTGDIFF = FALSE
                        If colStck.Count < 3 Then
                            Err.Raise Val(StackLow), Ambient.DisplayName, Mid$(StackLow, 4)
                          Else 'NOT COLSTCK.COUNT...
                            dblForVar = colStck.Item(3)                                       'intervals
                            If dblForVar = 0 Then
                                Err.Raise Val(NoIntervals), Ambient.DisplayName, Mid$(NoIntervals, 4)
                              Else 'NOT DBLFORVAR...
                                bolIntgDiff = True
                                dblIntegralVar = colStck.Item(1)                              'high limit
                                dblResult = Compute(Mid$(strFormula, 5))                      '(low limit value..
                                RaiseEvent Plot(dblIntegralVar, dblResult)                    'com first plot point
                                dblIntegralVar = colStck.Item(2)                              'low limit
                                dblResult = (dblResult + Compute(Mid$(strFormula, 5))) / 2    '..+ high limit value) / 2 - will be plotted later
                                dblLeftResult = dblIntegralVar                                'low limit
                                dblRiteResult = (colStck.Item(1) - dblLeftResult) / dblForVar 'interval size
                                'trapezoid - all points 'between'intervals
                                For dblForVar = dblForVar - 1 To 1 Step -1
                                    dblIntegralVar = dblLeftResult + dblRiteResult * dblForVar
                                    dblTemp1 = Compute(Mid$(strFormula, 5))
                                    RaiseEvent Plot(dblIntegralVar, dblTemp1)                 'com plot point
                                    dblResult = dblResult + dblTemp1
                                Next dblForVar
                                dblIntegralVar = colStck.Item(2)
                                RaiseEvent Plot(dblIntegralVar, Compute(Mid$(strFormula, 5))) 'com last plot point
                                'midpoints - all points in the middle of intervals
                                dblForVar = colStck.Item(3)                                   'intervals
                                dblLeftResult = dblLeftResult - dblRiteResult / 2             '<-- 1/2 interval left
                                For dblForVar = dblForVar To 1 Step -1
                                    dblIntegralVar = dblLeftResult + dblRiteResult * dblForVar
                                    dblResult = dblResult + Compute(Mid$(strFormula, 5)) * 2  'double for Simpson's parabolic approximation
                                Next dblForVar
                                'Simpson
                                dblResult = dblResult * Abs(dblRiteResult) / 3
                                bolIntgDiff = False
                            End If
                        End If
                    End If
                  Case "push"
                    dblResult = Compute(Mid$(strFormula, 5))
                    If colStck.Count = 0 Then
                        colStck.Add dblResult                       'first item on stack
                      Else 'NOT COLSTCK.COUNT...
                        colStck.Add dblResult, , 1                  'push on top
                    End If
                  Case "prod" 'multiply all values on stack
                    If colStck.Count = 0 Then
                        Err.Raise Val(StackLow), Ambient.DisplayName, Mid$(StackLow, 4)
                      Else 'NOT COLSTCK.COUNT...
                        dblResult = 1
                        For Each varData1 In colStck
                            dblResult = dblResult * varData1
                        Next varData1
                    End If
                  Case "pcor"
                    If (colStck.Count < 4) Or (colStck.Count And 1) Then
                        Err.Raise Val(StackLow), Ambient.DisplayName, Mid$(StackLow, 4)
                      Else 'NOT (COLSTCK.COUNT...
                        lngPtr1 = 0
                        dblTemp1 = 0
                        dblTemp2 = 0
                        For Each varData1 In colStck
                            lngPtr1 = lngPtr1 + 1
                            If lngPtr1 And 1 Then
                                dblMyResult = varData1
                                dblLeftResult = dblLeftResult + dblMyResult    'sigma x
                                dblTemp1 = dblTemp1 + dblMyResult ^ 2          'sigma x²
                              Else 'NOT LNGPTR1...
                                dblRiteResult = dblRiteResult + varData1       'sigma y
                                dblTemp2 = dblTemp2 + varData1 ^ 2             'sigma y²
                                dblResult = dblResult + dblMyResult * varData1 'sigma xy
                            End If
                        Next varData1
                        dblMyResult = Sqr((colStck.Count / 2 * dblTemp1 - dblLeftResult * dblLeftResult) * (colStck.Count / 2 * dblTemp2 - dblRiteResult * dblRiteResult))
                        If dblMyResult = 0 Then
                            dblResult = 1
                          Else 'NOT DBLMYRESULT...
                            dblResult = (colStck.Count / 2 * dblResult - dblLeftResult * dblRiteResult) / dblMyResult
                        End If
                    End If
                  Case "frac" 'fractional part
                    dblResult = Compute(Mid$(strFormula, 5))
                    dblResult = dblResult - Fix(dblResult)
                  Case "weeks", "weekday", "days", "hours", "minutes", "seconds"
                    lngPtr1 = FindMatchingBracket(Mid$(strFormula, Len(strWord) + 1)) - 2
                    strLeftPart = LCase$(Mid$(strFormula, Len(strWord) + 2, lngPtr1))
                    If strLeftPart = "now" Then
                        strLeftPart = Date & "@" & Time
                    End If
                    lngPtr1 = InStr(strLeftPart, "@")
                    If lngPtr1 Then
                        strRitePart = Mid$(strLeftPart, lngPtr1 + 1)
                        strLeftPart = Left$(strLeftPart, lngPtr1 - 1)
                        dblResult = CDbl(DateValue(strLeftPart) + TimeValue(strRitePart))
                      Else 'LNGPTR1 = FALSE
                        dblResult = CDbl(DateValue(strLeftPart))
                    End If
                    dblResult = dblResult + 115858
                    Select Case strWord
                      Case "weekday"
                        dblResult = Weekday(strLeftPart, vbMonday)
                      Case "weeks"
                        dblResult = dblResult / 7
                      Case "hours"
                        dblResult = dblResult * 24
                      Case "minutes"
                        dblResult = dblResult * 1440
                      Case "seconds"
                        dblResult = dblResult * 86400
                    End Select
                  Case "sec"
                    dblResult = 1 / Cos(Compute(Mid$(strFormula, 4)))
                  Case "cot"
                    'get around bug in VB - sign of tan() wrong in quadrant 2 and 4
                    dblResult = Compute(Mid$(strFormula, 4))
                    dblResult = Abs(Cos(dblResult) / Sin(dblResult)) * Sgn(Cos(dblResult))
                  Case "log"
                    dblResult = Log(Compute(Mid$(strFormula, 4))) / Log(10)
                  Case "sin"
                    dblResult = Sin(Compute(Mid$(strFormula, 4)))
                  Case "cos"
                    dblResult = Cos(Compute(Mid$(strFormula, 4)))
                  Case "tan"
                    'get around bug in VB6 - sign of tan() wrong in quadrant 2 and 4
                    dblResult = Compute(Mid$(strFormula, 4))
                    dblResult = Abs(Sin(dblResult) / Cos(dblResult)) * Sgn(Sin(dblResult))
                  Case "atn"
                    dblResult = Atn(Compute(Mid$(strFormula, 4)))
                  Case "pop", "top"
                    If colStck.Count = 0 Then
                        Err.Raise Val(StackLow), Ambient.DisplayName, Mid$(StackLow, 4)
                      Else 'NOT COLSTCK.COUNT...
                        dblResult = colStck.Item(1)
                        lngPopPending = lngPopPending Or IIf(strWord = "pop", 1, 0)
                    End If
                  Case "sum"
                    If colStck.Count = 0 Then
                        Err.Raise Val(StackLow), Ambient.DisplayName, Mid$(StackLow, 4)
                      Else 'NOT COLSTCK.COUNT...
                        For Each varData1 In colStck
                            dblResult = dblResult + varData1
                        Next varData1
                    End If
                  Case "avg"
                    dblResult = Compute("sum()") / colStck.Count
                  Case "min", "max"
                    If colStck.Count = 0 Then
                        Err.Raise Val(StackLow), Ambient.DisplayName, Mid$(StackLow, 4)
                      Else 'NOT COLSTCK.COUNT...
                        dblLeftResult = colStck.Item(1)
                        dblRiteResult = dblLeftResult
                        For Each varData1 In colStck
                            If varData1 < dblLeftResult Then
                                dblLeftResult = varData1
                            End If
                            If varData1 > dblRiteResult Then
                                dblRiteResult = varData1
                            End If
                        Next varData1
                        If strWord = "min" Then
                            dblResult = dblLeftResult
                          Else 'NOT STRWORD...
                            dblResult = dblRiteResult
                        End If
                    End If
                  Case "med" 'median
                    Select Case colStck.Count
                      Case 0
                        Err.Raise Val(StackLow), Ambient.DisplayName, Mid$(StackLow, 4)
                      Case 1
                        dblTemp1 = 1
                        dblTemp2 = 1
                      Case 2
                        dblTemp1 = 1
                        dblTemp2 = 2
                      Case Else
                        lngPtr1 = 0
                        For Each varData1 In colStck
                            dblLeftResult = 0
                            dblRiteResult = 0
                            lngPtr2 = 0
                            lngPtr1 = lngPtr1 + 1
                            For Each varData2 In colStck
                                lngPtr2 = lngPtr2 + 1
                                If lngPtr1 <> lngPtr2 Then
                                    If varData1 <= varData2 Then
                                        dblLeftResult = dblLeftResult + 1
                                    End If
                                    If varData1 >= varData2 Then
                                        dblRiteResult = dblRiteResult + 1
                                    End If
                                End If
                            Next varData2
                            If dblLeftResult = Int(colStck.Count / 2) Then
                                lngSplitAt = lngSplitAt Or 1
                                dblTemp1 = lngPtr1
                            End If
                            If dblRiteResult = Int(colStck.Count / 2) Then
                                lngSplitAt = lngSplitAt Or 2
                                dblTemp2 = lngPtr1
                            End If
                            If lngSplitAt = 3 Then
                                Exit For '>---> Next
                            End If
                        Next varData1
                    End Select
                    dblResult = (colStck.Item(dblTemp1) + colStck.Item(dblTemp2)) / 2
                  Case "lcd" 'largest common divider
                    If colStck.Count = 0 Then
                        Err.Raise Val(StackLow), Ambient.DisplayName, Mid$(StackLow, 4)
                      Else 'NOT COLSTCK.COUNT...
                        dblResult = Abs(Int(colStck.Item(1)))
                        For Each varData1 In colStck
                            If varData1 <> 0 Then
                                dblRiteResult = Abs(varData1)
                                dblResult = LCD(dblResult, dblRiteResult)
                            End If
                        Next varData1
                        If dblResult = 0 Then  'only zeros on stack
                            Err.Raise Val(StackLow), Ambient.DisplayName, Mid$(StackLow, 4)
                        End If
                    End If
                  Case "scm" 'smallest common multiple
                    dblResult = 1
                    For Each varData1 In colStck
                        dblRiteResult = Int(Abs(varData1))
                        If dblRiteResult <> 0 Then
                            dblTemp1 = dblResult / dblRiteResult
                            If dblTemp1 <> Int(dblTemp1) Then
                                dblResult = dblResult * dblRiteResult / LCD(dblResult, dblRiteResult)
                            End If
                        End If
                    Next varData1
                  Case "cor"
                    If (colStck.Count < 4) Or (colStck.Count And 1) Then
                        Err.Raise Val(StackLow), Ambient.DisplayName, Mid$(StackLow, 4)
                      Else 'NOT (COLSTCK.COUNT...
                        lngPtr1 = 0
                        dblTemp1 = 0
                        For Each varData1 In colStck
                            lngPtr1 = lngPtr1 + 1
                            If lngPtr1 And 1 Then
                                dblMyResult = varData1
                                dblLeftResult = dblLeftResult + dblMyResult     'sigma x
                                dblTemp1 = dblTemp1 + dblMyResult ^ 2           'sigma x²
                              Else 'NOT LNGPTR1...
                                dblRiteResult = dblRiteResult + varData1        'sigma y
                                dblResult = dblResult + dblMyResult * varData1  'sigma xy
                            End If
                        Next varData1
                        dblMyResult = colStck.Count / 2 * dblTemp1 - dblLeftResult * dblLeftResult
                        If dblMyResult = 0 Then
                            dblResult = 1
                          Else 'NOT DBLMYRESULT...
                            dblResult = (colStck.Count / 2 * dblResult - dblLeftResult * dblRiteResult) / dblMyResult
                        End If
                    End If
                  Case "sqr"
                    dblResult = Sqr(Compute(Mid$(strFormula, 4)))
                  Case "abs"
                    dblResult = Abs(Compute(Mid$(strFormula, 4)))
                  Case "neg"
                    dblResult = -(Compute(Mid$(strFormula, 4)))
                  Case "num"
                    dblResult = colStck.Count
                  Case "var"
                    If colStck.Count = 0 Then
                        Err.Raise Val(StackLow), Ambient.DisplayName, Mid$(StackLow, 4)
                      Else 'NOT COLSTCK.COUNT...
                        For Each varData1 In colStck
                            dblResult = dblResult + varData1 ^ 2
                        Next varData1
                        dblResult = Sqr((dblResult - (Compute("sum()²") / colStck.Count)) / colStck.Count)
                    End If
                  Case "vnc"
                    If colStck.Count < 2 Then
                        Err.Raise Val(StackLow), Ambient.DisplayName, Mid$(StackLow, 4)
                      Else 'NOT COLSTCK.COUNT...
                        For Each varData1 In colStck
                            dblResult = dblResult + varData1 ^ 2
                        Next varData1
                        dblResult = Sqr((dblResult - (Compute("sum()²") / colStck.Count)) / (colStck.Count - 1))
                    End If
                  Case "not"
                    If lngIfActiveDepth > lngDepth Then
                        dblResult = (Compute(Mid$(strFormula, 4)) = 0)       'True/False for If
                      Else 'NOT LNGIFACTIVEDEPTH...
                        dblResult = Not (Int(Compute(Mid$(strFormula, 4))))
                    End If
                  Case "int"
                    dblResult = Int(Compute(Mid$(strFormula, 4)))
                  Case "fix"
                    dblResult = Fix(Compute(Mid$(strFormula, 4)))
                  Case "fah"
                    dblResult = Compute(Mid$(strFormula, 4)) * 1.8 + 32
                  Case "cel"
                    dblResult = (Compute(Mid$(strFormula, 4)) - 32) / 1.8
                  Case "fac"
                    dblLeftResult = Int(Compute(Mid$(strFormula, 4)))
                    If dblLeftResult < 0 Then
                        Err.Raise 5, Ambient.DisplayName
                      Else 'NOT DBLLEFTRESULT...
                        dblResult = 1
                        For dblLeftResult = Int(dblLeftResult) To 2 Step -1
                            dblResult = dblResult * dblLeftResult
                        Next dblLeftResult
                    End If
                  Case "exp"
                    dblResult = Exp(Compute(Mid$(strFormula, 4)))
                  Case "deg"
                    dblResult = Compute(Mid$(strFormula, 4)) / Atn(1) * 45
                  Case "gra"
                    dblResult = Compute(Mid$(strFormula, 4)) / Atn(1) * 50
                  Case "arc", "rad"
                    dblResult = Compute(Mid$(strFormula, 4)) * Atn(1) / 45
                  Case "rnd"
                    dblResult = Compute(Mid$(strFormula, 4)) * Rnd
                  Case "if"
                    lngIfActiveDepth = lngDepth
                    If Compute(Mid$(strFormula, 3, FindMatchingBracket(Mid$(strFormula, 3)))) <> 0 Then
                        lngPtr1 = FindMatchingBracket(Mid$(strFormula, 3)) + 3
                        lngPtr2 = FindMatchingBracket(Mid$(strFormula, lngPtr1))
                      Else 'NOT COMPUTE(MID$(STRFORMULA,...
                        lngPtr1 = FindMatchingBracket(Mid$(strFormula, 3)) + 3
                        lngPtr2 = FindMatchingBracket(Mid$(strFormula, lngPtr1))
                        lngPtr1 = lngPtr1 + lngPtr2
                        lngPtr2 = FindMatchingBracket(Mid$(strFormula, lngPtr1))
                    End If
                    lngIfActiveDepth = 0
                    dblResult = Compute(Mid$(strFormula, lngPtr1, lngPtr2))
                  Case "ln"
                    dblResult = Log(Compute(Mid$(strFormula, 3)))
                  Case Else 'function name token
                    dblResult = Compute(TranslateToken(strWord, "") & Mid$(strFormula, Len(strWord) + 1))
                End Select
            End If
        End If
        Compute = dblResult
    End If
    lngDepth = lngDepth - 1

End Function

Public Function Evaluate() As Double

    lngDepth = 0
    bolIntgDiff = False
    dblMyResult = Compute(strMyFormula)
    Evaluate = dblMyResult

End Function

Private Function FindMatchingBracket(Formula As String) As Long

  Dim lngPtr   As Long
  Dim lngMtch  As Long

    Do
        lngPtr = lngPtr + 1
        Select Case Mid$(Formula, lngPtr, 1)
          Case "("
            lngMtch = lngMtch + 1
          Case ")"
            lngMtch = lngMtch - 1
        End Select
    Loop While lngMtch
    FindMatchingBracket = lngPtr

End Function

Public Property Get Formula() As String
Attribute Formula.VB_Description = "Returns/Sets the Formula for the Evaluator to compute."
Attribute Formula.VB_HelpID = 10000
Attribute Formula.VB_ProcData.VB_Invoke_Property = "MiniCalc"
Attribute Formula.VB_UserMemId = 0
Attribute Formula.VB_MemberFlags = "200"

    Formula = strMyFormula

End Property

Public Property Let Formula(ByVal nwFormula As String)

    strMyFormula = Trim$(nwFormula)
    If LCase$(Left$(strMyFormula, 4)) <> "push" Then
        dblMyPreviousResult = dblMyResult
    End If
    lngPopPending = 0
    Evaluate
    If lngPopPending And 2 Then            'was PopAll
        Set colStck = New Collection
      ElseIf lngPopPending And 1 Then      'was Pop'NOT LNGPOPPENDING...
        colStck.Remove 1
    End If
    If LCase$(Left$(strMyFormula, 4)) <> "push" Then
        bolPreviousValid = True
    End If
    PropertyChanged PropFormula

End Property

Private Sub GeneratePrimes()

    If Not bolPrimesPresent Then
        ReDim Preserve lngPrime(0 To lngMyUBPrime)
        bolShowProgress = True
        Prime 5
        bolPrimesPresent = True
    End If

End Sub

Private Function LCD(ByVal dblNumerator As Double, ByVal dblDenominator As Double) As Double

  'Largest Common Divisor - credit goes to Euklid in ancient Greece

    Do Until dblNumerator = 0
        dblTemp1 = dblDenominator Mod dblNumerator    'mod includes int()
        dblDenominator = dblNumerator
        dblNumerator = dblTemp1
    Loop
    LCD = dblDenominator

End Function

Public Property Get PreviousResult() As Double

    PreviousResult = dblMyPreviousResult

End Property

Private Function Prime(ByVal dblFrom As Double) As Double

  'find first prim > dblFrom

    If bolShowProgress Then
        fProgress.UseWholeBar = True
        fProgress.Progress = 0
        DoEvents
    End If

    'set dblFrom down to previous odd number
    dblFrom = Int(dblFrom / 2) * 2 - 1
    Do
        bolSuccess = False
        'next number to test
        dblFrom = dblFrom + 2
        lngPtr1 = 0
        'division loop
        Do
            'next divisor
            lngPtr1 = lngPtr1 + 1
            If lngPtr1 > lngMyUBPrime Then
                Err.Raise Val(TooFewPrimes), Ambient.DisplayName, Mid$(TooFewPrimes, 4)
            End If
            'divide value by next divisor
            dblTemp1 = dblFrom / lngPrime(lngPtr1)
            If dblTemp1 = Int(dblTemp1) Then
                'no remainder - not a prime
                Exit Do '>---> Loop
            End If
            If dblTemp1 < lngPrime(lngPtr1) Then
                'all necessary divisions done - this is a prime
                If lngXPrime < lngMyUBPrime Then
                    'still room in primes table
                    lngXPrime = lngXPrime + 1
                    lngPrime(lngXPrime) = dblFrom
                    If bolShowProgress Then
                        If dblTemp2 = Int(dblTemp2) Then
                            fProgress.Progress = lngXPrime / lngMyUBPrime * 100
                        End If
                    End If
                End If
                'signal success
                bolSuccess = True
                Exit Do '>---> Loop
            End If
        Loop
    Loop Until lngXPrime = lngMyUBPrime And bolSuccess
    If bolShowProgress Then
        fProgress.Progress = -1
        bolShowProgress = False
    End If
    Prime = dblFrom

End Function

Public Property Get PrimeTableSize() As Long
Attribute PrimeTableSize.VB_ProcData.VB_Invoke_Property = "MiniCalc"

    PrimeTableSize = lngMyUBPrime

End Property

Public Property Let PrimeTableSize(ByVal nwSize As Long)

    If nwSize <> lngMyUBPrime Then
        If Ambient.UserMode Then 'not in run mode
            Err.Raise Val(PropLocked), Ambient.DisplayName, Mid$(PropLocked, 4)
          Else 'AMBIENT.USERMODE = FALSE
            If nwSize < 10 Then
                nwSize = 10
            End If
            If nwSize > 500000 Then 'oops, more than fivehundred thousend primes?
                Err.Raise 380, Ambient.DisplayName
              Else 'NOT NWSIZE...
                Select Case nwSize
                  Case Is < lngMyUBPrime
                    ReDim Preserve lngPrime(0 To nwSize)
                    lngMyUBPrime = nwSize
                    If bolPrimesPresent Then
                        lngXPrime = nwSize
                    End If
                  Case Is > lngMyUBPrime
                    lngMyUBPrime = nwSize
                    bolPrimesPresent = False
                End Select
                PropertyChanged PropPTS
            End If
        End If
    End If

End Property

Public Property Get Result() As Double

    Result = dblMyResult

End Property

Private Function TranslateToken(strToken As String, strDefault As String) As String

  Dim strTokenValue    As String

    If Len(strToken) Then
        If Ambient.UserMode Then
            RaiseEvent QueryToken(strToken, strTokenValue)
          Else 'AMBIENT.USERMODE = FALSE
            strTokenValue = InputBox("Enter Value for Token '" & strToken & "' below:", Ambient.DisplayName & ": Query Token Value", strDefault)
        End If
        If Len(strTokenValue) = 0 Then
            Err.Raise Val(ValueMissing), Ambient.DisplayName, Mid$(ValueMissing, 4) & strToken & "'"
        End If
        TranslateToken = strTokenValue
      Else 'LEN(STRTOKEN) = FALSE
        Err.Raise Val(OperandMissing), Ambient.DisplayName, Mid$(OperandMissing, 4)
    End If

End Function

Private Sub UserControl_Initialize()

    dblMyResult = 0
    dblMyPreviousResult = 0
    bolPreviousValid = False
    strMyFormula = ""
    Set colStck = New Collection
    strDPFrom = Format$(0.1)        'adjust decimal point to International Setting
    If InStr(strDPFrom, ",") Then
        strDPFrom = "."
        strDPTo = ","
      Else 'NOT INSTR(STRDPFROM,...
        strDPFrom = ","
        strDPTo = "."
    End If
    ReDim lngPrime(0 To 1)
    lngXPrime = 1
    lngPrime(0) = 2
    lngPrime(1) = 3

End Sub

Private Sub UserControl_InitProperties()

    lngMyUBPrime = PTS

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        strMyFormula = .ReadProperty(PropFormula, " ")
        lngMyUBPrime = .ReadProperty(PropPTS, PTS)
    End With 'PROPBAG

End Sub

Private Sub UserControl_Resize()

    UserControl.Size img.Width, img.Height

End Sub

Private Sub UserControl_Terminate()

    Set colStck = Nothing

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty PropFormula, strMyFormula, " "
        .WriteProperty PropPTS, lngMyUBPrime, PTS
    End With 'PROPBAG

End Sub

':) Ulli's VB Code Formatter V2.11.3 (06.04.2002 10:46:36) 55 + 1026 = 1081 Lines
