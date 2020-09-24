VERSION 5.00
Begin VB.Form fProgress 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   330
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   510
   ControlBox      =   0   'False
   DrawMode        =   10  'Stift maskieren
   FillColor       =   &H8000000F&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   22
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   34
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "fProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public UseWholeBar      As Boolean

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Private Type POINT
    x                   As Long
    y                   As Long
End Type

Private Type RECT
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
End Type

Private WindowRect      As RECT
Private CursorPos       As POINT
Private hWndTray        As Long
Private PrintY          As Long
Private FColor          As Long
Private BColor          As Long
Private ThisPercent     As Long
Private PrevPercent     As Long
Private PrintPercent    As String

Public Property Let Progress(ByVal Percent As Single)

    If Percent < 0 Then
        If hWndTray Then
            hWndTray = 0
            SetCursorPos CursorPos.x, CursorPos.y
            Do Until ShowCursor(True) = 0
            Loop
        End If
        PrevPercent = -1
        Unload Me
      Else 'NOT PERCENT...
        If hWndTray = 0 Then
            GetCursorPos CursorPos 'save cursor pos
            Do While ShowCursor(False) = 0
            Loop
            hWndTray = FindWindow("Shell_TrayWnd", "") 'find tray
            If Not UseWholeBar Then
                hWndTray = FindWindowEx(hWndTray, 0, "TrayNotifyWnd", "") 'find notify window
            End If
            GetWindowRect hWndTray, WindowRect
            With WindowRect
                Width = (.Right - .Left - 2) * 15 'adjust my size
                Height = (.Bottom - .Top - 2) * 15
            End With 'WINDOWRECT
            PrintY = (ScaleHeight - TextHeight("A")) / 2 'vertical print pos
            SetParent hWnd, hWndTray 'tray is my parent
            ScaleWidth = 1000 'percent * 10
            FColor = ForeColor 'colors...
            If FColor < 0 Then
                FColor = GetSysColor(FColor And &H7FFFFFFF)
            End If
            BColor = BackColor
            If BColor < 0 Then
                BColor = GetSysColor(BColor And &H7FFFFFFF)
            End If
            BColor = Not (FColor Xor BColor)
            Show
        End If
        With Screen
            SetCursorPos .Width - 1 / .TwipsPerPixelX, .Height / .TwipsPerPixelY 'to make tray visible
        End With 'SCREEN
        PrintPercent = Int(Percent) & "%"
        ThisPercent = Percent * 10
        If ThisPercent <> PrevPercent Then
            Cls
            CurrentX = 500 - TextWidth(PrintPercent) / 2
            CurrentY = PrintY
            Print PrintPercent
            'uncomment one
            Line (0, 0)-(ThisPercent, ScaleHeight), BColor, BF
            'Line (0, 3)-(ThisPercent , ScaleHeight - 4), BColor, BF 'alternatively - leaves a small top and bottom margin
            Refresh
            PrevPercent = ThisPercent
        End If
    End If

End Property

Private Sub Form_Load()

    PrevPercent = -1

End Sub

':) Ulli's VB Code Formatter V2.11.3 (06.04.2002 10:32:29) 34 + 67 = 101 Lines
