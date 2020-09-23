VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Image Area Selection Demo - stormdev@golden.net"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   Icon            =   "frmPaintForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Selection Mode"
      Height          =   1215
      Left            =   3000
      TabIndex        =   3
      Top             =   2520
      Width           =   2535
      Begin VB.OptionButton Tools 
         Caption         =   "Inverted Select - Masking"
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton Tools 
         Caption         =   "Normal Select"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      DrawStyle       =   1  'Dash
      Height          =   3060
      Left            =   120
      Picture         =   "frmPaintForm.frx":000C
      ScaleHeight     =   3000
      ScaleWidth      =   2700
      TabIndex        =   0
      Top             =   1560
      Width           =   2760
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "stormdev@golden.net"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3000
      TabIndex        =   8
      Top             =   4320
      Width           =   1560
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Votes and comments welcome but not expected."
      Height          =   435
      Left            =   3000
      TabIndex        =   7
      Top             =   3840
      Width           =   2565
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose a selection mode from below and click and drag in the image to display the selection rectangle."
      Height          =   975
      Left            =   3120
      TabIndex        =   6
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Area"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5520
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmPaintForm.frx":1A630
      Height          =   795
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5445
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This code is copyright(c) 2000, 2001 Stormdev Software Development.
' You are hereby granted rights to use/modify this code as you see fit,
' for commercial or personal use. The only stipulation is that I
' ask for some feedback regarding the code contained herein.
'
' Send feedback to: stormdev@golden.net
'      Code Author: Jonathan Roach
'  Purpose of Code: Demonstrate the creation of rubber band selection rectangles
'    Level of Code: Beginner
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const INVERSE = 6       ' DrawMode property - XOR
Const SOLID = 0         ' DrawStyle property
Const DOT = 2           ' DrawStyle property
Dim DrawBox As Boolean
Dim OldX As Single
Dim OldY As Single
Dim StartX As Single
Dim StartY As Single

Private Sub Form_Load()
Me.Show
Picture1.DrawStyle = DOT
DrawBox = True
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.Cls
' Store the initial start of the line to draw.
StartX = X
StartY = Y

' Make the last location equal the starting location
OldX = StartX
OldY = StartY
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If the button is pressed
If Button = 1 Then
' Erase the previous line
    DrawLine StartX, StartY, OldX, OldY
' Draw the new line.
    DrawLine StartX, StartY, X, Y
' Save the coordinates for the next call.
    OldX = X
    OldY = Y
End If
End Sub

Sub DrawLine(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single)
' Save the current mode so that you can reset it on
' exit from this sub routine. Not needed in the sample
' but would need it if you are not sure what the
' DrawMode was on entry to this procedure.
SavedMode% = Picture1.DrawMode

' Set to XOR
Picture1.DrawMode = INVERSE
' Draw a box or line
If DrawBox = True Then
    Picture1.Line (X1, Y1)-(X2, Y2), , B
Else
    Picture1.Line (X1, Y1)-(X2, Y2)
End If

' Reset the DrawMode
Picture1.DrawMode = SavedMode%
End Sub

Private Sub Tools_Click(Index As Integer)
' Selection of rubber band line mode
Select Case Index
Case 0 ' Normal Select Tool
        DrawBox = True
        Picture1.FillStyle = 1
Case 1 ' Inverted Select Tool
        DrawBox = True
        Picture1.FillStyle = 0
End Select
End Sub
