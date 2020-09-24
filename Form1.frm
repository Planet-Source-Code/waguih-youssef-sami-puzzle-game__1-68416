VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Waguih Puzzle"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   ScaleHeight     =   384
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   493
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "Options"
      ForeColor       =   &H8000000E&
      Height          =   1560
      Left            =   5775
      TabIndex        =   0
      Top             =   315
      Width           =   1395
      Begin VB.CommandButton cmdSolve 
         BackColor       =   &H8000000C&
         Caption         =   "Solve"
         Height          =   495
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   330
         Width           =   1080
      End
      Begin VB.CommandButton cmdScramble 
         BackColor       =   &H8000000C&
         Caption         =   "shuffle"
         Height          =   495
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   900
         Width           =   1080
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   405
      Top             =   4500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Click to Drag and Drop in the empty Square"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   645
      Left            =   390
      TabIndex        =   3
      Top             =   4470
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   1140
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   4530
      Width           =   2400
   End
   Begin VB.Shape Shape1 
      Height          =   3960
      Left            =   360
      Top             =   330
      Width           =   5265
   End
   Begin VB.Image Pic 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Index           =   11
      Left            =   4275
      Stretch         =   -1  'True
      Top             =   2955
      Width           =   1305
   End
   Begin VB.Image Pic 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Index           =   10
      Left            =   2985
      Stretch         =   -1  'True
      Top             =   2955
      Width           =   1305
   End
   Begin VB.Image Pic 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Index           =   9
      Left            =   1695
      Stretch         =   -1  'True
      Top             =   2955
      Width           =   1305
   End
   Begin VB.Image Pic 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Index           =   8
      Left            =   405
      Stretch         =   -1  'True
      Top             =   2955
      Width           =   1305
   End
   Begin VB.Image Pic 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Index           =   0
      Left            =   405
      Stretch         =   -1  'True
      Top             =   375
      Width           =   1305
   End
   Begin VB.Image Pic 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Index           =   1
      Left            =   1695
      Stretch         =   -1  'True
      Top             =   390
      Width           =   1305
   End
   Begin VB.Image Pic 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Index           =   2
      Left            =   2985
      Stretch         =   -1  'True
      Top             =   390
      Width           =   1305
   End
   Begin VB.Image Pic 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Index           =   3
      Left            =   4275
      Stretch         =   -1  'True
      Top             =   375
      Width           =   1305
   End
   Begin VB.Image Pic 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Index           =   7
      Left            =   4275
      Stretch         =   -1  'True
      Top             =   1665
      Width           =   1305
   End
   Begin VB.Image Pic 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Index           =   6
      Left            =   2985
      Stretch         =   -1  'True
      Top             =   1665
      Width           =   1305
   End
   Begin VB.Image Pic 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Index           =   5
      Left            =   1695
      Stretch         =   -1  'True
      Top             =   1665
      Width           =   1305
   End
   Begin VB.Image Pic 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Index           =   4
      Left            =   405
      Stretch         =   -1  'True
      Top             =   1665
      Width           =   1305
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim F, M, Z, OldNumber(12), NewNumber

Private Sub cmdScramble_Click()
Z = 1
Dim i
For i = 0 To 12
OldNumber(i) = 0
Next
Timer1_Timer
End Sub

Private Sub cmdSolve_Click()
Dim i
For i = 1 To 11
Pic(i).Picture = LoadPicture(App.Path & "\Pic" & i + 1 & ".jpg")
Next
End Sub

Private Sub Form_Load()
F = 0
M = 0
Z = 1
Image1.Picture = LoadPicture(App.Path & "\fountain.jpg")
Form1.Picture = LoadPicture(App.Path & "\marble.jpg")
Form1.Icon = LoadPicture(App.Path & "\cube.ico")

End Sub

Private Sub Form_Activate()
Dim i
For i = 0 To 12
OldNumber(i) = 0
Next

'*****************
Dim intX As Integer
    Dim intY As Integer
    Dim sngWidth As Single
    Dim sngHeight As Single
    
    sngWidth = 65 'Form1.Picture.Width
    sngHeight = 60 'Form1.Picture.Height
    For intX = 0 To Int(Form1.ScaleWidth / sngWidth)
        For intY = 0 To Int(Form1.ScaleHeight / sngHeight)
            PaintPicture Form1.Picture, intX * sngWidth, intY * sngHeight, sngWidth, sngHeight, 0, 0
        Next
    Next


End Sub

Private Sub Pic_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
M = Index
If Pic(M).Picture = LoadPicture("") Then
Pic(M).Picture = Pic(F).Picture
Pic(F).Picture = LoadPicture("")
M = 0
F = 0
End If
End Sub

Private Sub Pic_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
F = Index
Pic(F).Drag vbBeginDrag

End Sub

Private Sub Timer1_Timer()

Do While Z < 12

NewNumber = Int(12 * Rnd + 1)
10:
If NewNumber = OldNumber(1) Or NewNumber = OldNumber(2) _
Or NewNumber = OldNumber(3) Or NewNumber = OldNumber(4) _
Or NewNumber = OldNumber(5) Or NewNumber = OldNumber(6) _
Or NewNumber = OldNumber(7) Or NewNumber = OldNumber(8) _
Or NewNumber = OldNumber(9) Or NewNumber = OldNumber(10) _
Or NewNumber = OldNumber(11) Or NewNumber = OldNumber(12) Then
NewNumber = Int(12 * Rnd + 1)
GoTo 10:

Else
Pic(Z).Picture = LoadPicture(App.Path & "\Pic" & NewNumber & ".jpg")
Z = Z + 1

End If
OldNumber(Z) = NewNumber
Loop
End Sub


