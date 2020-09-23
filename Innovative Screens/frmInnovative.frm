VERSION 5.00
Begin VB.Form frmInnovative 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmInnovative.frx":0000
   ScaleHeight     =   3750
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Show Process Indicator"
      Height          =   375
      Left            =   300
      TabIndex        =   9
      Top             =   2745
      Width           =   2040
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6315
      Top             =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test"
      Height          =   300
      Left            =   5025
      TabIndex        =   5
      Top             =   1140
      Width           =   1065
   End
   Begin VB.PictureBox pcPanel 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0C0C0&
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
      Height          =   540
      Left            =   0
      Picture         =   "frmInnovative.frx":27F52
      ScaleHeight     =   540
      ScaleWidth      =   6375
      TabIndex        =   1
      Top             =   3210
      Width           =   6375
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Innovative Buttons. Add captions as per ur need ->"
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Top             =   165
         Width           =   3735
      End
      Begin VB.Image imgButton 
         Height          =   360
         Index           =   1
         Left            =   5160
         Top             =   90
         Width           =   1065
      End
      Begin VB.Image imgButton 
         Height          =   360
         Index           =   0
         Left            =   3960
         Top             =   90
         Width           =   1065
      End
      Begin VB.Image imgUp 
         Height          =   345
         Left            =   1965
         Picture         =   "frmInnovative.frx":28D2A
         Top             =   555
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Image imgDown 
         Height          =   345
         Left            =   1800
         Picture         =   "frmInnovative.frx":2A0D4
         Top             =   630
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   210
         TabIndex        =   2
         Top             =   165
         Width           =   45
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expecting ur comments to ramandy@rediffmail.com"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2475
      TabIndex        =   8
      Top             =   2835
      Width           =   3690
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Innovative Check Boxes, Just click check boxes ->"
      Height          =   195
      Left            =   285
      TabIndex        =   7
      Top             =   2175
      Width           =   3630
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   3
      Left            =   4770
      Picture         =   "frmInnovative.frx":2B47E
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   2
      Left            =   4530
      Picture         =   "frmInnovative.frx":2B5C8
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   1
      Left            =   4290
      Picture         =   "frmInnovative.frx":2B712
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   0
      Left            =   4050
      Picture         =   "frmInnovative.frx":2B85C
      Top             =   2160
      Width           =   240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Innovative Option Buttons, Just click any option ->"
      Height          =   195
      Left            =   285
      TabIndex        =   6
      Top             =   1830
      Width           =   3675
   End
   Begin VB.Image imgOptionChecked 
      Height          =   240
      Left            =   1350
      Picture         =   "frmInnovative.frx":2B9A6
      Top             =   3390
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgOption 
      Height          =   240
      Index           =   3
      Left            =   4770
      Picture         =   "frmInnovative.frx":2BAF0
      Top             =   1830
      Width           =   240
   End
   Begin VB.Image imgOption 
      Height          =   240
      Index           =   2
      Left            =   4530
      Picture         =   "frmInnovative.frx":2BC3A
      Top             =   1830
      Width           =   240
   End
   Begin VB.Image imgOption 
      Height          =   240
      Index           =   1
      Left            =   4290
      Picture         =   "frmInnovative.frx":2BD84
      Top             =   1830
      Width           =   240
   End
   Begin VB.Image imgOptionOff 
      Height          =   240
      Left            =   1635
      Picture         =   "frmInnovative.frx":2BECE
      Top             =   3420
      Width           =   240
   End
   Begin VB.Image ImgOptionOn 
      Height          =   240
      Left            =   1950
      Picture         =   "frmInnovative.frx":2C018
      Top             =   3450
      Width           =   240
   End
   Begin VB.Image imgOption 
      Height          =   240
      Index           =   0
      Left            =   4050
      Picture         =   "frmInnovative.frx":2C162
      Top             =   1830
      Width           =   240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Innovative Progress bar ->"
      Height          =   195
      Left            =   1185
      TabIndex        =   4
      Top             =   885
      Width           =   1965
   End
   Begin VB.Shape shpFiller 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   240
      Left            =   1260
      Top             =   1170
      Width           =   15
   End
   Begin VB.Shape shpOuter 
      BorderColor     =   &H00C0C0C0&
      Height          =   300
      Left            =   1200
      Top             =   1140
      Width           =   3690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Innovative form designing TekniX"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   165
      Width           =   3825
   End
End
Attribute VB_Name = "frmInnovative"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IntI As Integer

Private Sub Command1_Click()
With shpOuter
    shpFiller.Move .Left + 30, .Top + 30, 15, .Height - 45
    Timer1.Enabled = True
End With
End Sub

Private Sub Command2_Click()
Form1.Show
End Sub

Private Sub Form_Load()
    For IntI = imgButton.LBound To imgButton.UBound
        Set imgButton(IntI).Picture = imgUp.Picture
    Next IntI
End Sub

Private Sub imgButton_Click(Index As Integer)
    Unload Me
End Sub

Private Sub imgButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgButton(Index).Picture = imgDown.Picture
End Sub

Private Sub imgButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgButton(Index).Picture = imgUp.Picture
End Sub

Private Sub imgCheck_Click(Index As Integer)

    If imgCheck(Index).Picture = imgOptionChecked.Picture Then
        Set imgCheck(Index).Picture = imgOptionOff.Picture
    Else
       Set imgCheck(Index).Picture = imgOptionChecked.Picture
    End If

    ': Write code here to handle option check

End Sub

Private Sub imgOption_Click(Index As Integer)

    For IntI = imgOption.LBound To imgOption.UBound
        Set imgOption(IntI).Picture = imgOptionOff.Picture
    Next IntI
    Set imgOption(Index).Picture = ImgOptionOn.Picture
    
    ': Code here to track Option selected
    
End Sub

Private Sub Timer1_Timer()
    shpFiller.Width = shpFiller.Width + 15
    If shpFiller.Width > (shpOuter.Width - 45) Then Timer1.Enabled = False
End Sub
