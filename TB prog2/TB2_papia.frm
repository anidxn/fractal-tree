VERSION 5.00
Begin VB.Form T2F_papia 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tree Branches : Developed by Papia Das"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   12585
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
      ItemData        =   "TB2_papia.frx":0000
      Left            =   1800
      List            =   "TB2_papia.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox PicCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   8415
      Left            =   120
      ScaleHeight     =   557
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   821
      TabIndex        =   0
      Top             =   480
      Width           =   12375
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
      Left            =   4200
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
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   1500
   End
End
Attribute VB_Name = "T2F_papia"
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
       theta = 30
     Case 1:
       theta = 60
     Case 2:
       theta = 90
     Case 3:
       theta = 120
     Case 4:
       theta = 180
 End Select
  PicCanvas.Cls
  PicCanvas.Line (400, 350)-(400, 550) 'drawing STEM
  Call drawT(400, 350, 90, 100)
End Sub

Private Function drawT(ByVal x1 As Integer, ByVal y1 As Integer, ByVal angB As Integer, ByVal l As Single)

    If l <= 5 Then
      Exit Function
    End If
    Dim newAng As Integer, xn As Integer, yn As Integer
    newAng = angB + theta / 2
    xn = x1 + l * Cos(newAng * pi / 180)
    yn = y1 - l * Sin(newAng * pi / 180)
    PicCanvas.Line (x1, y1)-(xn, yn)
    Call drawT(xn, yn, newAng, l * 0.7)
    
    newAng = angB - theta / 2
    xn = x1 + l * Cos(newAng * pi / 180)
    yn = y1 - l * Sin(newAng * pi / 180)
    PicCanvas.Line (x1, y1)-(xn, yn)
    Call drawT(xn, yn, newAng, l * 0.7)
End Function
