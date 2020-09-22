VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   8760
      TabIndex        =   0
      Top             =   1200
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1320
      Top             =   2520
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   10
      Left            =   4080
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   9
      Left            =   3120
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   8
      Left            =   3600
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   7
      Left            =   4560
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   6
      Left            =   5040
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   5
      Left            =   6000
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   4
      Left            =   5520
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   6480
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   7440
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   6960
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   2640
      Top             =   2520
      Width           =   480
   End
   Begin VB.Line Line2 
      Index           =   10
      X1              =   240
      X2              =   8640
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line2 
      Index           =   9
      X1              =   0
      X2              =   8400
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line2 
      Index           =   8
      X1              =   240
      X2              =   8640
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line2 
      Index           =   7
      X1              =   0
      X2              =   8400
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line2 
      Index           =   6
      X1              =   0
      X2              =   8400
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line2 
      Index           =   5
      X1              =   0
      X2              =   8400
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line2 
      Index           =   4
      X1              =   480
      X2              =   8880
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line2 
      Index           =   3
      X1              =   1560
      X2              =   9960
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   120
      X2              =   8520
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   0
      X2              =   8400
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   360
      X2              =   8760
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      X1              =   8640
      X2              =   8640
      Y1              =   360
      Y2              =   2760
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public say As Integer
Private Sub Command1_Click()
    Timer1.Enabled = True
End Sub
Private Sub Form_Load()
    Randomize Timer
    Timer1.Enabled = False
    say = 0
    For i = 0 To 10
        Image1(i).Left = 200
        Image1(i).Top = 200 + (i + 1) * 200
        Line2(i).X1 = 200:  Line2(i).X2 = 8800
        Line2(i).Y1 = 170 + (i + 1) * 200
        Line2(i).Y2 = 170 + (i + 1) * 200
    Next i
End Sub

Private Sub Image1_Click(Index As Integer)

End Sub

Private Sub Timer1_Timer()
    say = say + 1
    For i = 0 To 10
        Image1(i).Left = Image1(i).Left + Rnd(100) * 100
        Image1(i).Picture = LoadPicture("gif\HORSE (" & Trim(Str((i + say) Mod 23)) & ").gif")
        If Image1(i).Left + 255 >= 8640 Then
            Timer1.Enabled = False
            MsgBox Str(i + 1) & ". win"
            Call Form_Load
        End If
    Next i
End Sub
