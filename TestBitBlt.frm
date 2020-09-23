VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "How to combine foreground and background"
   ClientHeight    =   5832
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNextStep 
      Caption         =   "Next Step"
      Height          =   372
      Left            =   3240
      TabIndex        =   6
      Top             =   5400
      Width           =   1452
   End
   Begin VB.PictureBox Foreground 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      Height          =   2400
      Left            =   5400
      ScaleHeight     =   196
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   5
      Top             =   3000
      Width           =   2400
   End
   Begin VB.PictureBox ReverseMask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      Height          =   2400
      Left            =   2760
      ScaleHeight     =   196
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   4
      Top             =   3000
      Width           =   2400
   End
   Begin VB.PictureBox Background 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      Height          =   2400
      Left            =   2760
      Picture         =   "TestBitBlt.frx":0000
      ScaleHeight     =   196
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   3
      Top             =   120
      Width           =   2400
   End
   Begin VB.PictureBox Final 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      Height          =   2400
      Left            =   5400
      ScaleHeight     =   196
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   2
      Top             =   120
      Width           =   2400
   End
   Begin VB.PictureBox Mask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      Height          =   2400
      Left            =   120
      ScaleHeight     =   196
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   1
      Top             =   3000
      Width           =   2400
   End
   Begin VB.PictureBox Sprite 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      Height          =   2400
      Left            =   120
      Picture         =   "TestBitBlt.frx":38582
      ScaleHeight     =   196
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   0
      Top             =   120
      Width           =   2400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Dim i As Integer
Dim i2 As Integer
Dim Result
Dim Step As Byte


Private Sub cmdNextStep_Click()
Step = Step + 1
Select Case Step
Case 1: Step1
Case 2: Step2
Case 3: Step3
Case 4: Step4
Case 5: Step5
Case 6: Step6
Case 7: Step7
Case 8: Step8
Case 9: Step9
End Select
End Sub

Sub Step1()
'make a copy of the sprite
Caption = "Make a copy of the sprite"
Result = BitBlt(Mask.hDC, 0, 0, Mask.Width, Mask.Height, Sprite.hDC, 0, 0, vbSrcCopy)
Mask.Picture = Mask.Image
End Sub

Sub Step2()
cmdNextStep.Enabled = False
'white out the color which is to be transparent
Caption = "White out the areas of sprite which are not required"
Dim TransColor As Long
TransColor = Sprite.Point(5, 5)
Sprite.Picture = Sprite.Image

For i = 0 To 200
For i2 = 0 To 200
   If Sprite.Point(i, i2) = TransColor Then
     Mask.PSet (i, i2), vbWhite
   End If
Next
DoEvents
Next
cmdNextStep.Enabled = True
End Sub

Sub Step3()
cmdNextStep.Enabled = False
'now any non-white areas can be blacked out
Caption = "Now any non-white areas can be blacked out"
For i = 0 To 200
For i2 = 0 To 200
   If Mask.Point(i, i2) <> vbWhite Then
     Mask.PSet (i, i2), vbBlack
   End If
Next
DoEvents
Next
cmdNextStep.Enabled = True
End Sub

Sub Step4()
'Place the background on your final picture
Caption = "Copy the background onto the final picture"
BitBlt Final.hDC, 0, 0, Final.Width, Final.Height, Background.hDC, 0, 0, vbSrcCopy
Final.Picture = Final.Image
End Sub

Sub Step5()
'Now use mergepaint (this only copies the back bits) to white out the foreground on the final pic"
Caption = "Now use mergepaint (this only copies the back bits) to white out the foreground on the final pic"
BitBlt Final.hDC, 0, 0, Final.Width, Final.Height, Mask.hDC, 0, 0, vbMergePaint
Final.Picture = Final.Image
End Sub

Sub Step6()
'make a reverse mask using notcopy
Caption = "Now make a reverse mask using NotCopy"
BitBlt ReverseMask.hDC, 0, 0, Mask.Width, Mask.Height, Mask.hDC, 0, 0, vbNotSrcCopy
ReverseMask.Picture = ReverseMask.Image
End Sub

Sub Step7()
'make a copy of the sprite
Caption = "Make another copy of the sprite"
BitBlt Foreground.hDC, 0, 0, Mask.Width, Mask.Height, Sprite.hDC, 0, 0, vbSrcCopy
Foreground.Picture = Foreground.Image
End Sub

Sub Step8()
'now white out unwanted area
Caption = "Now white out the unwanted area, using mergepaint from the reverse mask"
BitBlt Foreground.hDC, 0, 0, Mask.Width, Mask.Height, ReverseMask.hDC, 0, 0, vbMergePaint
Foreground.Picture = Foreground.Image
End Sub

Sub Step9()
'copy the foreground to the final picture
Caption = "Copy the foreground to the final picture using And (does not copy white areas)"
BitBlt Final.hDC, 0, 0, Mask.Width, Mask.Height, Foreground.hDC, 0, 0, vbSrcAnd
Final.Picture = Final.Image
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

