VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Rock Away Help"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   25
      Left            =   0
      TabIndex        =   2
      Top             =   4440
      Width           =   5535
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Dim tmp As String
tmp = tmp & "Goal of Rock Away is to play until only one rock remains in the center of the gameboard. The goal is simple, but the game is hard!" & vbCrLf & vbCrLf
tmp = tmp & "To take away a rock, you must jump over it with another rock. This can only be done horizontally or vertically." & vbCrLf & vbCrLf
tmp = tmp & "Click with the mouse on the rock that you want to jump with, and click on the empty place across the target rock. Your rock will jump over the target rock to the empty place, and the target rock disappears." & vbCrLf & vbCrLf
tmp = tmp & "You can use the 'Undo' or 'Redo' menu to go back on your steps." & vbCrLf & vbCrLf
tmp = tmp & "Select 'New' in the 'Game' menu to reset the timer." & vbCrLf & vbCrLf
tmp = tmp & "In the 'options' menu you can enable or disable the sound, and select the gameboard color." & vbCrLf & vbCrLf
tmp = tmp & "Good Luck!" & vbCrLf & vbCrLf
tmp = tmp & "Rock Away is written by Dirk Rijmenants © 2005" & vbCrLf
Me.lblHelp.Caption = tmp
End Sub
