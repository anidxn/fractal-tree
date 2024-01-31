VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "The Pythagoras Tree"
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   12060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Tree3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Tree2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tree1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.PictureBox PicCanvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H0000FF00&
      Height          =   8775
      Left            =   120
      ScaleHeight     =   583
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   783
      TabIndex        =   0
      Top             =   480
      Width           =   11775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lev As Integer
Dim nItr As Integer
Dim dtr As Single
Private Declare Function FloodFill Lib "gdi32" _
        (ByVal hdc As Long, _
         ByVal X As Long, _
         ByVal Y As Long, _
         ByVal crColor As Long) As Long


Private Function drawbox(ByVal Ax As Integer, ByVal Ay As Integer, ByVal iAng As Single, ByVal side As Single, left As Boolean)
    If side <= 3 Then
       Exit Function
    End If
    Dim Bx As Integer, By As Integer, Cx As Integer, Cy As Integer, Dx As Integer, Dy As Integer

    Bx = Ax + side * Cos(iAng * dtr)
    By = Ay - side * Sin(iAng * dtr)
    
    Cx = Ax + (side * (2 ^ 0.5)) * Cos((iAng + 45) * dtr)
    Cy = Ay - (side * (2 ^ 0.5)) * Sin((iAng + 45) * dtr)
    
    Dx = Ax + side * Cos((iAng + 90) * dtr)
    Dy = Ay - side * Sin((iAng + 90) * dtr)
    
    lev = lev + 1
    PicCanvas.FillColor = RGB(255 - 15 * lev, 255, 0)

    PicCanvas.Line (Ax, Ay)-(Bx, By)
    PicCanvas.Line (Bx, By)-(Cx, Cy)
    PicCanvas.Line (Cx, Cy)-(Dx, Dy)
    PicCanvas.Line (Dx, Dy)-(Ax, Ay)
    
    FloodFill PicCanvas.hdc, Ax + side * Cos((iAng + 45) * dtr), Ay - side * Sin((iAng + 45) * dtr), vbGreen
    'DoEvents
    
    If left = True Then
       Call drawbox(Dx, Dy, iAng + 45, side * (2 ^ 0.5) / 2, True) 'decreament side by sqrt(2)/2 of square
       Call drawbox(Cx, Cy, iAng + 45, side * (2 ^ 0.5) / 2, False)
    Else
       Call drawbox(Bx, By, iAng - 45, side * (2 ^ 0.5) / 2, False)
       Call drawbox(Cx, Cy, iAng - 45, side * (2 ^ 0.5) / 2, True)
    End If
    
    lev = lev - 1 'each level gets the same color
   
End Function

Private Sub Command1_Click()
PicCanvas.Cls
lev = 1 'for coloring
PicCanvas.FillStyle = 0
PicCanvas.FillColor = vbYellow
   Dim l As Single
    l = 100#
    PicCanvas.Line (300, 400)-(300 + l, 400 + l), vbGreen, B
    Call drawbox(300, 400, 45, l * (2 ^ 0.5) / 2, True)
    Call drawbox(300 + l, 400, 45, l * (2 ^ 0.5) / 2, False)

End Sub

Private Sub Command2_Click()

nItr = CInt(InputBox("Enter number of Iterations [max 15] : "))
If nItr > 15 Then nItr = 15
If nItr < 1 Then nItr = 1

PicCanvas.Cls

lev = 1 'for coloring
PicCanvas.FillStyle = 0
PicCanvas.FillColor = vbYellow

   Dim l As Single
    l = 20
    PicCanvas.Line (350, 300)-(350 + l, 300 + l), vbGreen, B
    Call drawSphere(350, 300, 60, l, True)
    Call drawSphere(350 + l, 300, 30, l, False)

End Sub



Private Sub Form_Load()
   dtr = 3.14159 / 180
   PicCanvas.ForeColor = vbGreen
End Sub
'---------------------Draws a spere with all box of equal length of side & angle 60 deg
Private Function drawSphere(ByVal Ax As Integer, ByVal Ay As Integer, ByVal iAng As Single, ByVal side As Single, left As Boolean)
    
    If lev = nItr Then
       Exit Function
    End If
    Dim Bx As Integer, By As Integer, Cx As Integer, Cy As Integer, Dx As Integer, Dy As Integer

    Bx = Ax + side * Cos(iAng * dtr)
    By = Ay - side * Sin(iAng * dtr)
    
    Cx = Ax + (side * (2 ^ 0.5)) * Cos((iAng + 45) * dtr)
    Cy = Ay - (side * (2 ^ 0.5)) * Sin((iAng + 45) * dtr)
    
    Dx = Ax + side * Cos((iAng + 90) * dtr)
    Dy = Ay - side * Sin((iAng + 90) * dtr)
    
    lev = lev + 1
    PicCanvas.FillColor = RGB(255, 255 - 10 * lev, 0)

    PicCanvas.Line (Ax, Ay)-(Bx, By)
    PicCanvas.Line (Bx, By)-(Cx, Cy)
    PicCanvas.Line (Cx, Cy)-(Dx, Dy)
    PicCanvas.Line (Dx, Dy)-(Ax, Ay)
    
    FloodFill PicCanvas.hdc, Ax + side * Cos((iAng + 45) * dtr), Ay - side * Sin((iAng + 45) * dtr), vbGreen
    DoEvents
    
    If left = True Then
       Call drawSphere(Dx, Dy, iAng + 60, side, True) 'left branch
       Call drawSphere(Cx, Cy, iAng + 30, side, False) 'right branch
    Else
       Call drawSphere(Bx, By, iAng - 60, side, False) 'right branch
       Call drawSphere(Cx, Cy, iAng - 30, side, True) 'left branch
    End If
    
    lev = lev - 1 'each level gets the same color
   
End Function


Private Sub Command3_Click()
PicCanvas.Cls
lev = 1 'for coloring
PicCanvas.FillStyle = 0
PicCanvas.FillColor = vbYellow
   Dim l As Single
    l = 100#
    PicCanvas.Line (200, 400)-(200 + l, 400 + l), vbGreen, B
    Call drawTree(200, 400, 60, l * 0.5, True) 'Left
    Call drawTree(200 + l, 400, 60, l * 0.866, False) 'Right
End Sub


Private Function drawTree(ByVal Ax As Integer, ByVal Ay As Integer, ByVal iAng As Single, ByVal side As Single, left As Boolean)
    If lev = 13 Then
       Exit Function
    End If
    Dim Bx As Integer, By As Integer, Cx As Integer, Cy As Integer, Dx As Integer, Dy As Integer

    Bx = Ax + side * Cos(iAng * dtr)
    By = Ay - side * Sin(iAng * dtr)
    
    Cx = Ax + (side * (2 ^ 0.5)) * Cos((iAng + 45) * dtr)
    Cy = Ay - (side * (2 ^ 0.5)) * Sin((iAng + 45) * dtr)
    
    Dx = Ax + side * Cos((iAng + 90) * dtr)
    Dy = Ay - side * Sin((iAng + 90) * dtr)

    lev = lev + 1
    PicCanvas.FillColor = RGB(255 - 8 * lev, 255, 0)

    PicCanvas.Line (Ax, Ay)-(Bx, By)
    PicCanvas.Line (Bx, By)-(Cx, Cy)
    PicCanvas.Line (Cx, Cy)-(Dx, Dy)
    PicCanvas.Line (Dx, Dy)-(Ax, Ay)
    
    FloodFill PicCanvas.hdc, Ax + side * Cos((iAng + 45) * dtr), Ay - side * Sin((iAng + 45) * dtr), vbGreen
    'DoEvents
    
    If left = True Then
       Call drawTree(Dx, Dy, iAng + 60, side * 0.5, True)   'decreament side by sqrt(2)/2 of square
       Call drawTree(Cx, Cy, iAng + 60, side * 0.866, False)
    Else
       Call drawTree(Cx, Cy, iAng - 30, side * 0.5, True) 'sin(30) =0.5
       Call drawTree(Bx, By, iAng - 30, side * 0.866, False)
    End If
    
    lev = lev - 1 'each level gets the same color
   
End Function

Private Sub Form_Unload(Cancel As Integer)
Set Form1 = Nothing
End Sub
