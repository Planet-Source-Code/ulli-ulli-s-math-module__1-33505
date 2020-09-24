VERSION 5.00
Object = "{7B2798AE-2D81-11D3-B079-BC5450D64B2E}#33.0#0"; "Evaluator.ocx"
Begin VB.Form fIntegration 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Integration Example"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   495
      Left            =   285
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "for x = 0° .. 180°"
      Height          =   195
      Left            =   2115
      TabIndex        =   7
      Top             =   705
      Width           =   1140
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4125
      TabIndex        =   6
      Top             =   975
      Width           =   2490
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4125
      TabIndex        =   5
      Top             =   285
      Width           =   2490
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3840
      TabIndex        =   4
      Top             =   285
      Width           =   165
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Max function value at"
      Height          =   195
      Left            =   2385
      TabIndex        =   3
      Top             =   1185
      Width           =   1530
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Function Plot"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   285
      TabIndex        =   1
      Top             =   1905
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Intg((sin(x))²)-pi/2 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1815
      TabIndex        =   0
      Top             =   285
      Width           =   1920
   End
   Begin EvaluatorOCX.Evaluator Evaluator1 
      Left            =   285
      Top             =   855
      _ExtentX        =   661
      _ExtentY        =   423
      Formula         =   ""
      PrimeTableSize  =   10
   End
End
Attribute VB_Name = "fIntegration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Integration example using the Evaluator Control
'
'Integration computes the area a under a function graph in square units.
'
'This example computes the area under a squared sine half wave, ie the
'red area which is plotted...
'
'----------------------------------------------------------------
'For those of you who had advanced math at school:
'
'        _
'       |pi
' a  =  |  (sin(x))²
'      _|0
'
'----------------------------------------------------------------
'
'...and then subtracts pi/2 from (a) to verify that the result is correct.
'

Option Explicit

Private ymax As Double, xmax As Double

Private Sub Command1_Click()

    Cls
    Label2.Visible = False
    Label5 = ""
    Label6 = ""
    DoEvents
    ScaleMode = vbPixels
    Screen.MousePointer = vbHourglass
    With Evaluator1
        .Formula = "Push(2000)"              'many intervals (this is a very 'stubborn' function and we want to impress you)
        .Formula = "Push(0°)"                'low integration limit
        .Formula = "Push(180°)"              'high integration limit
        .Formula = Label1                    'go
        Label5 = .Result                     'display result
        .Formula = "PopAll()"                'clear stack
    End With 'EVALUATOR1
    Screen.MousePointer = vbNormal
    Label6 = "X = " & xmax & "  Y = " & ymax
    Label2.Visible = True

End Sub

Private Sub Evaluator1_Plot(ByVal X As Double, ByVal Y As Double)

    Line (10 + X * 100, 300 - Y * 150)-(10 + X * 100, 301), vbRed

    If Y > ymax Then
        ymax = Y
        xmax = X
    End If

End Sub

':) Ulli's VB Code Formatter V2.11.3 (06.04.2002 00:06:29) 15 + 34 = 49 Lines
Private Sub Form_Load()

End Sub
