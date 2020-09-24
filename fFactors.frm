VERSION 5.00
Object = "{7B2798AE-2D81-11D3-B079-BC5450D64B2E}#31.0#0"; "Evaluator.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Factorization Example"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   1305
      TabIndex        =   2
      Top             =   825
      Width           =   1020
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   495
      Left            =   3135
      TabIndex        =   1
      Top             =   2355
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1305
      TabIndex        =   0
      Text            =   "(2^17-1) ³"
      Top             =   225
      Width           =   1170
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "The first call will create a table having the above number of primes. Subsequent calls use the generated table."
      Height          =   1065
      Left            =   2805
      TabIndex        =   7
      Top             =   1020
      Width           =   1740
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Factors"
      Height          =   195
      Left            =   285
      TabIndex        =   6
      Top             =   855
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Left            =   2865
      TabIndex        =   5
      Top             =   600
      Width           =   105
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Primes Table Size"
      Height          =   195
      Left            =   2865
      TabIndex        =   4
      Top             =   270
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter number or formula to factorize"
      Height          =   585
      Left            =   285
      TabIndex        =   3
      Top             =   135
      Width           =   975
   End
   Begin EvaluatorOCX.Evaluator Evaluator1 
      Left            =   4065
      Top             =   615
      _ExtentX        =   661
      _ExtentY        =   423
      Formula         =   ""
      PrimeTableSize  =   100000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Factorization example using the Evaluator Control
'
'please note that the 2^17-1 (131071) in this example is a prime

Option Explicit

Private Sub Command1_Click()

  Dim i As Long
  Dim j As Long

    List1.Clear
    With Evaluator1
        On Error Resume Next
          .Formula = "Fact(" & Text1 & ")" 'returns the number of factors in .Result
          'and the factors on the Control's stack
          If Err Then
              MsgBox Err.Description
            Else 'ERR = FALSE
              j = .Result 'number of factors
              .Formula = "Pop()" 'first factor
              If .Result < 0 Then
                  List1.AddItem .Result 'add first factor to listbox if negative
              End If
              For i = 1 To j - 1 'remaining factors
                  .Formula = "Pop()"
                  List1.AddItem .Result
              Next i
          End If
        On Error GoTo 0
    End With 'EVALUATOR1

End Sub

Private Sub Form_Load()

    Label3 = Evaluator1.PrimeTableSize

End Sub

Private Sub Text1_GotFocus()

    Text1.SelStart = 0
    Text1.SelLength = 127

End Sub

':) Ulli's VB Code Formatter V2.11.3 (06.04.2002 10:00:40) 5 + 47 = 52 Lines
