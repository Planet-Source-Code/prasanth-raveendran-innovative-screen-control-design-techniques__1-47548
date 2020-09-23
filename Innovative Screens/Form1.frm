VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Process Indicator"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrRunner 
      Interval        =   10
      Left            =   2010
      Top             =   3435
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   1920
      Width           =   5715
      Begin VB.Shape shpF 
         BorderColor     =   &H00E0E0E0&
         DrawMode        =   9  'Not Mask Pen
         Height          =   225
         Index           =   0
         Left            =   3555
         Top             =   2175
         Width           =   315
      End
      Begin VB.Shape shpRunner 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         BorderStyle     =   0  'Transparent
         Height          =   105
         Left            =   3645
         Top             =   2235
         Width           =   105
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Process indicator: Use to show there is something is running on background. Vary lite weight. Only shape controls used."
      Height          =   780
      Left            =   1005
      TabIndex        =   1
      Top             =   450
      Width           =   3960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IntInitialLeftPos   As Integer  ': Left of the lest most element
Private IntInitialTopPos    As Integer  ': Top of the lest most element
Private BytNumberOfBoxes    As Byte     ': How many boxes
Private LngRunnerColor      As Long     ': Color for Running Object
Private LngBoxBorderColor   As Long     ': Box Border color
Private IntSpeed            As Integer  ': Speed Factor, Same as timer
Private LngI                As Long     ': Jsut Counter

Private Sub Form_Load()

': Initial Settings
BytNumberOfBoxes = 12                   ': How many Boxes ?
LngRunnerColor = RGB(0, 128, 0)         ': Backcolor of running control
IntSpeed = 500                          ': Speed, Same as Timer->Intervel Property
LngBoxBorderColor = RGB(192, 192, 192)  ': Box Border Color
IntInitialLeftPos = 1600                ': Initial Left Position
IntInitialTopPos = 300                  ': Top Position

tmrRunner.Interval = IntSpeed
shpRunner.BackColor = LngRunnerColor

shpF(0).BorderColor = LngBoxBorderColor


    For LngI = 1 To (BytNumberOfBoxes - 1)
        Load shpF(LngI)
        shpF(LngI).Visible = True
        shpF(LngI).BorderColor = LngBoxBorderColor
    Next LngI
    
    For LngI = shpF.LBound To shpF.UBound
        shpF(LngI).Move IntInitialLeftPos, IntInitialTopPos, 150, 150
        IntInitialLeftPos = IntInitialLeftPos + 180
    Next LngI
    LngI = 0
End Sub

Private Sub tmrRunner_timer()
    With shpF(LngI)
        shpRunner.Move .Left + 30, .Top + 30, .Width - 45, .Height - 45
    End With
    
    If LngI >= (BytNumberOfBoxes - 1) Then
        LngI = 0
    Else
        LngI = LngI + 1
    End If
End Sub
