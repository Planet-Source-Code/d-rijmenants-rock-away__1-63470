VERSION 5.00
Begin VB.Form frmHighScores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " High Scores"
   ClientHeight    =   3360
   ClientLeft      =   9465
   ClientTop       =   4020
   ClientWidth     =   3345
   ControlBox      =   0   'False
   Icon            =   "frmHighScores.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   25
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   3015
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   31
      Top             =   240
      Width           =   1700
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   30
      Top             =   480
      Width           =   1700
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   29
      Top             =   720
      Width           =   1700
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   28
      Top             =   960
      Width           =   1700
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   27
      Top             =   1200
      Width           =   1700
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   26
      Top             =   1440
      Width           =   1700
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   25
      Top             =   1680
      Width           =   1700
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   480
      TabIndex        =   24
      Top             =   1920
      Width           =   1700
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   480
      TabIndex        =   23
      Top             =   2160
      Width           =   1700
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   480
      TabIndex        =   22
      Top             =   2400
      Width           =   1700
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   21
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   20
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   19
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   18
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   2160
      TabIndex        =   17
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   16
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   2160
      TabIndex        =   15
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   2160
      TabIndex        =   14
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   2160
      TabIndex        =   13
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   2160
      TabIndex        =   12
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblNmbr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "1."
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   11
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblNmbr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "2."
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   10
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblNmbr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "3."
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   9
      Top             =   720
      Width           =   375
   End
   Begin VB.Label lblNmbr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "4."
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lblNmbr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "5."
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   7
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblNmbr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "6."
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblNmbr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "7."
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   5
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lblNmbr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "8."
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   4
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblNmbr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "9."
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   3
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lblNmbr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "10."
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   0
      TabIndex        =   2
      Top             =   2400
      Width           =   375
   End
End
Attribute VB_Name = "frmHighScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Me.Hide
End Sub

Private Sub Form_Activate()
Dim k As Integer
For k = 0 To 9
If Scores(k, 0) <> "" Then
    Me.lblNmbr(k) = Trim(Str(k + 1)) & "."
    Me.lblName(k).Caption = Scores(k, 0)
    Me.lblTime(k).Caption = Scores(k, 1) & " sec"
    Else
    Me.lblNmbr(k) = ""
    Me.lblName(k).Caption = ""
    Me.lblTime(k).Caption = ""
End If
Next
End Sub
