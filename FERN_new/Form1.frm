VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10995
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   13365
   LinkTopic       =   "Form1"
   ScaleHeight     =   10995
   ScaleWidth      =   13365
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicCanvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   10455
      Left            =   120
      ScaleHeight     =   10425
      ScaleWidth      =   13065
      TabIndex        =   0
      Top             =   240
      Width           =   13095
   End
   Begin VB.Menu mchos 
      Caption         =   "Choose"
      Begin VB.Menu mBF 
         Caption         =   "Barnsley Fern"
      End
      Begin VB.Menu mTF 
         Caption         =   "Thelypteridaceae "
      End
      Begin VB.Menu mLF 
         Caption         =   "Leptosporangiate"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long, j As Integer
Dim num As Single
Dim x As Single, y As Single
Dim x1 As Single, y1 As Single

Dim m_Prob(0 To 3) As Single

Private Sub Form_Load()
    PicCanvas.Scale (-4, 10)-(4, 0)
End Sub

Private Sub Form_Resize()
If Form1.WindowState <> vbMinimized Then
   Dim hgt As Single
    hgt = ScaleHeight - PicCanvas.Top
    If hgt < 120 Then hgt = 120
    PicCanvas.Move 0, PicCanvas.Top, ScaleWidth, hgt
    PicCanvas.Scale (-4, 10.1)-(4, -0.1)
End If
End Sub

Private Sub mBF_Click() 'Barnsley Fern
Randomize
PicCanvas.Cls
    m_Prob(0) = 0.01
    m_Prob(1) = 0.85
    m_Prob(2) = 0.07
    m_Prob(3) = 0.07
    
    Screen.MousePointer = vbHourglass
    
        x = 1: y = 1
        For i = 1 To 200000 'draw upto 1 lakh pixels
            If i Mod 1000 = 0 Then DoEvents 'update picturebox after 1000 pixels

            'choose function using probability
            num = Rnd 'Rnd returns value from 0...1 & sum of all probability=1
            For j = 0 To 3
                num = num - m_Prob(j)
                If num <= 0 Then
                    Exit For
                End If
            Next j
            
            Select Case j
                   Case 0:
                     x1 = x * 0
                     y1 = y * 0.16
                     'PicCanvas.PSet (x1, y1), vbBlue
                   Case 1:
                     x1 = 0.85 * x + 0.04 * y
                     y1 = -0.04 * x + 0.85 * y + 1.6
                     'PicCanvas.PSet (x1, y1), vbGreen
                   Case 2:
                     x1 = 0.2 * x - 0.26 * y
                     y1 = 0.23 * x + 0.22 * y + 1.6
                     'PicCanvas.PSet (x1, y1), vbRed
                   Case 3:
                     x1 = -0.15 * x + 0.28 * y
                     y1 = 0.26 * x + 0.24 * y + 0.44
                     'PicCanvas.PSet (x1, y1), vbYellow
            End Select

            x = x1: y = y1
            PicCanvas.PSet (x, y), RGB(86, 255, 32) 'vbGreen
       Next i
    Screen.MousePointer = vbDefault
End Sub

Private Sub mLF_Click()
Randomize
PicCanvas.Cls
    m_Prob(0) = 0.02
    m_Prob(1) = 0.84
    m_Prob(2) = 0.07
    m_Prob(3) = 0.07
    
    Screen.MousePointer = vbHourglass
    
        x = 1: y = 1
        For i = 1 To 100000 'draw upto 1 lakh pixels
            If i Mod 1000 = 0 Then DoEvents 'update picturebox after 1000 pixels

            'choose finction using probability
            num = Rnd 'Rnd returns value from 0...1 & sum of all probability=1
            For j = 0 To 3
                num = num - m_Prob(j)
                If num <= 0 Then
                    Exit For
                End If
            Next j
            
            Select Case j
                   Case 0:
                     x1 = x * 0
                     y1 = y * 0.25 - 0.14
                   Case 1:
                     x1 = 0.85 * x + 0.02 * y
                     y1 = -0.02 * x + 0.83 * y + 1
                   Case 2:
                     x1 = 0.09 * x - 0.28 * y
                     y1 = 0.3 * x + 0.11 * y + 0.6
                   Case 3:
                     x1 = -0.09 * x + 0.28 * y
                     y1 = 0.3 * x + 0.09 * y + 0.7
            End Select

            x = x1: y = y1
            PicCanvas.PSet (x, y), vbGreen
       Next i
    Screen.MousePointer = vbDefault
End Sub

Private Sub mTF_Click()
Randomize
PicCanvas.Cls
    m_Prob(0) = 0.02
    m_Prob(1) = 0.84
    m_Prob(2) = 0.07
    m_Prob(3) = 0.07
    
    Screen.MousePointer = vbHourglass
    
        x = 1: y = 1
        For i = 1 To 100000 'draw upto 1 lakh pixels
            If i Mod 1000 = 0 Then DoEvents 'update picturebox after 1000 pixels

            'choose finction using probability
            num = Rnd 'Rnd returns value from 0...1 & sum of all probability=1
            For j = 0 To 3
                num = num - m_Prob(j)
                If num <= 0 Then
                    Exit For
                End If
            Next j
            
            Select Case j
                   Case 0:
                     x1 = x * 0
                     y1 = y * 0.25 - 0.4
                   Case 1:
                     x1 = 0.95 * x + 0.005 * y - 0.002
                     y1 = -0.005 * x + 0.93 * y + 0.5
                   Case 2:
                     x1 = 0.035 * x - 0.2 * y - 0.09
                     y1 = 0.16 * x + 0.04 * y + 0.02
                   Case 3:
                     x1 = -0.04 * x + 0.2 * y + 0.083
                     y1 = 0.16 * x + 0.04 * y + 0.12
            End Select

            x = x1: y = y1
            PicCanvas.PSet (x, y), vbGreen
       Next i
    Screen.MousePointer = vbDefault
End Sub

