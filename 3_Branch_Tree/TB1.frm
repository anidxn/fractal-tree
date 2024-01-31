VERSION 5.00
Begin VB.Form T1F 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tree Branches"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "TB1.frx":0000
      Left            =   1800
      List            =   "TB1.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox PicCanvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C000&
      Height          =   6975
      Left            =   120
      ScaleHeight     =   6945
      ScaleWidth      =   9105
      TabIndex        =   0
      Top             =   480
      Width           =   9135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Developed By : Papia Das"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2520
      TabIndex        =   3
      Top             =   2640
      Width           =   3480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select an angle : "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1500
   End
End
Attribute VB_Name = "T1F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pi = 3.14159
Dim theta As Integer

Private Sub Combo1_Click()
 Select Case Combo1.ListIndex
     Case 0:
       theta = 60
     Case 1:
       theta = 90
     Case 2:
       theta = 110
     Case 3:
       theta = 120
     Case 4:
       theta = 130
 End Select
  PicCanvas.Cls
  PicCanvas.Line (4000, 4000)-(4000, 6000) 'draw STEM
  Call drawT2(4000, 4000, 90, 1000)
End Sub

Private Function drawT2(ByVal x1 As Integer, ByVal y1 As Integer, ByVal angB As Integer, ByVal l As Single)

    If l <= 10 Then
      Exit Function
    End If
    Dim newAng As Integer, xn As Integer, yn As Integer
    newAng = angB + theta
    xn = x1 + l * Cos(newAng * pi / 180)
    yn = y1 - l * Sin(newAng * pi / 180)
    'PicCanvas.ForeColor = vbGreen 'vbBlue
    PicCanvas.Line (x1, y1)-(xn, yn)
    Call drawT2(xn, yn, newAng, l * 0.5)
    
    newAng = angB
    xn = x1 + l * Cos(newAng * pi / 180)
    yn = y1 - l * Sin(newAng * pi / 180)
    'PicCanvas.ForeColor = vbGreen
    PicCanvas.Line (x1, y1)-(xn, yn)
    Call drawT2(xn, yn, newAng, l * 0.5)
    
    newAng = angB - theta
    xn = x1 + l * Cos(newAng * pi / 180)
    yn = y1 - l * Sin(newAng * pi / 180)
    'PicCanvas.ForeColor = vbGreen 'vbBlue
    PicCanvas.Line (x1, y1)-(xn, yn)
    Call drawT2(xn, yn, newAng, l * 0.5)
End Function

