VERSION 5.00
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17280
   DrawWidth       =   3
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   864
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1152
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrBullet 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1920
      Top             =   1800
   End
   Begin VB.PictureBox Sprite1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   885
      Left            =   0
      Picture         =   "frmGame.frx":0000
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   4
      Top             =   0
      Width           =   1200
      Visible         =   0   'False
   End
   Begin VB.PictureBox Mask1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   885
      Left            =   0
      Picture         =   "frmGame.frx":313E
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   3
      Top             =   0
      Width           =   1200
      Visible         =   0   'False
   End
   Begin VB.Timer tmrFade 
      Interval        =   100
      Left            =   1560
      Top             =   1800
   End
   Begin VB.Frame frmBar 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   11400
      Width           =   17295
      Begin VB.Label lblAdd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add house"
         BeginProperty Font 
            Name            =   "Netherworld"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   1440
         TabIndex        =   5
         Top             =   120
         Width           =   1170
      End
      Begin VB.Label lblHealth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Health : 100"
         BeginProperty Font 
            Name            =   "Netherworld"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Image Bullet 
      Height          =   60
      Left            =   2640
      Picture         =   "frmGame.frx":627C
      Top             =   4680
      Width           =   60
      Visible         =   0   'False
   End
   Begin VB.Image ImgChar 
      Height          =   480
      Left            =   840
      Picture         =   "frmGame.frx":62EE
      Top             =   2520
      Width           =   480
      Visible         =   0   'False
   End
   Begin VB.Label lblLoading 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Netherworld"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6480
      TabIndex        =   2
      Top             =   4800
      Width           =   2205
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bLeft As Boolean, Up As Boolean, Right As Boolean, Down As Boolean, ChrState As Integer, Direction As Integer, cBullet As Integer, fBlock As Boolean
Dim Color As Integer, Sound As Integer
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then bLeft = True: Direction = 3
If KeyCode = vbKeyRight Then Right = True: Direction = 4
If KeyCode = vbKeyUp Then Up = True: Direction = 1
If KeyCode = vbKeyDown Then Down = True: Direction = 2
If fBlock = True Then
ElseIf fBlock = False Then
If KeyCode = vbKeyControl Then
Fire
Bullet.Visible = True
fBlock = True
End If
End If
MoveChr
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 27
End
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then bLeft = False: Direction = 3
If KeyCode = vbKeyRight Then Right = False: Direction = 4
If KeyCode = vbKeyUp Then Up = False: Direction = 1
If KeyCode = vbKeyDown Then Down = False: Direction = 2
MoveChr
End Sub

Private Sub Form_Load()
bLeft = False: Right = False
Up = False
Down = False
ChrState = 1
Color = 1
Sound = 1
Direction = 1
Fade
End Sub
Function MoveChr()
If bLeft = True Then
If ChrState >= 15 Then
ChrState = 1
ImgChar.Picture = LoadPicture(App.Path & "\character\char01\left\run" & ChrState & ".ico")
ElseIf ChrState < 15 Then
ChrState = ChrState + 1
ImgChar.Picture = LoadPicture(App.Path & "\character\char01\left\run" & ChrState & ".ico")
End If
If CheckGround = True Then
ImgChar.Left = ImgChar.Left - 3
Else
End If
ElseIf Right = True Then
If ChrState >= 15 Then
ChrState = 1
ImgChar.Picture = LoadPicture(App.Path & "\character\char01\right\run" & ChrState & ".ico")
ElseIf ChrState < 15 Then
ChrState = ChrState + 1
ImgChar.Picture = LoadPicture(App.Path & "\character\char01\right\run" & ChrState & ".ico")
End If
If CheckGround = True Then
ImgChar.Left = ImgChar.Left + 3
Else
End If
ElseIf Up = True Then
If ChrState >= 15 Then
ChrState = 1
ImgChar.Picture = LoadPicture(App.Path & "\character\char01\up\run" & ChrState & ".ico")
ElseIf ChrState < 15 Then
ChrState = ChrState + 1
ImgChar.Picture = LoadPicture(App.Path & "\character\char01\up\run" & ChrState & ".ico")
If CheckGround = True Then
ImgChar.Top = ImgChar.Top - 3
Else
End If
End If
ElseIf Down = True Then
If ChrState >= 15 Then
ChrState = 1
ImgChar.Picture = LoadPicture(App.Path & "\character\char01\down\run" & ChrState & ".ico")
ElseIf ChrState < 15 Then
ChrState = ChrState + 1
ImgChar.Picture = LoadPicture(App.Path & "\character\char01\down\run" & ChrState & ".ico")
End If
If CheckGround = True Then
ImgChar.Top = ImgChar.Top + 3
Else
End If
End If
If ChrState = 7 Or ChrState = 14 Then
If Sound = 1 Then
sndPlaySound App.Path & "\sfx\foot1.wav", 1
Sound = 2
Else
sndPlaySound App.Path & "\sfx\foot2.wav", 1
Sound = 1
End If
End If
End Function
Function CheckGround() As Boolean
Dim Walk As Boolean, bNot As Boolean
If bLeft = True Then
If Me.Point(ImgChar.Left - 1, ImgChar.Top + ImgChar.Height / 2) = vbBlack Or Me.Point(ImgChar.Left - 1, ImgChar.Top) = RGB(136, 68, 0) Then
CheckGround = False
Else
CheckGround = True
End If
ElseIf Right = True Then
If Me.Point(ImgChar.Left + ImgChar.Width - 4, ImgChar.Top + ImgChar.Height / 2) = vbBlack Or Me.Point(ImgChar.Left - 4, ImgChar.Top) = RGB(136, 68, 0) Then
CheckGround = False
Else
CheckGround = True
End If
ElseIf Up = True Then
If Me.Point(ImgChar.Left + ImgChar.Width / 2, ImgChar.Top + 1) = vbBlack Or Me.Point(ImgChar.Left + ImgChar.Width / 2, ImgChar.Top - 1) = RGB(136, 68, 0) Then
CheckGround = False
Else
CheckGround = True
End If
ElseIf Down = True Then
If Me.Point(ImgChar.Left + ImgChar.Width / 2, ImgChar.Top + ImgChar.Height) = vbBlack Or Me.Point(ImgChar.Left + ImgChar.Width / 2, ImgChar.Top + ImgChar.Height) = RGB(136, 68, 0) Then
CheckGround = False
Else
CheckGround = True
End If
End If
End Function
Function Fade()
Me.BackColor = RGB(0, 255, 0)
Me.Line (0, 0)-(0, 864), vbBlack
Me.Line (0, 0)-(1152, 0), vbBlack
Me.Line (0, 759)-(1152, 759), vbBlack
Me.Line (1150, 860)-(1150, 0), vbBlack
BitBlt Me.hDC, 35, 45, 76, 55, Mask1.hDC, 0, 0, SRCAND
BitBlt Me.hDC, 35, 45, 76, 55, Sprite1.hDC, 0, 0, SRCINVERT
BitBlt Me.hDC, 205, 75, 76, 55, Mask1.hDC, 0, 0, SRCAND
BitBlt Me.hDC, 205, 75, 76, 55, Sprite1.hDC, 0, 0, SRCINVERT
BitBlt Me.hDC, 35, 305, 76, 55, Mask1.hDC, 0, 0, SRCAND
BitBlt Me.hDC, 35, 305, 76, 55, Sprite1.hDC, 0, 0, SRCINVERT
Me.Refresh
ImgChar.Visible = True
lblLoading.Visible = False
End Function
Function Door() As Boolean
Dim Walk As Boolean, bNot As Boolean
If bLeft = True Then
If Me.Point(ImgChar.Left, ImgChar.Top) = RGB(0, 0, 125) Then
Door = False
Else
Door = True
End If
ElseIf Right = True Then
If Me.Point(ImgChar.Left, ImgChar.Top) = RGB(0, 0, 125) Then
Door = False
Else
Door = True
End If
ElseIf Up = True Then
If Me.Point(ImgChar.Left, ImgChar.Top) = RGB(0, 0, 125) Then
Door = False
Else
Door = True
End If
ElseIf Down = True Then
If Me.Point(ImgChar.Left, ImgChar.Top) = RGB(0, 0, 125) Then
Door = False
Else
Door = True
End If
End If
End Function
Function Fire()
Select Case Direction
Case 1
i = Bullet.Top
Bullet.Top = ImgChar.Top
Bullet.Left = ImgChar.Left + ImgChar.Width / 2
Do
DoEvents
If Bullet.Top < 4 Or Bulletway = True Then
fBlock = False
Bullet.Visible = False
Exit Function
End If
Bullet.Top = Bullet.Top - 3
Loop

Case 2
Bullet.Top = ImgChar.Top
Bullet.Left = ImgChar.Left + ImgChar.Width / 2
Do Until Bullet.Top > Me.ScaleHeight Or Bulletway = True
DoEvents
If Bullet > Me.ScaleHeight Or Bulletway = True Then
fBlock = False
Bullet.Visible = False
Exit Function
End If
Bullet.Top = Bullet.Top + 3
Loop

Case 3
i = Bullet.Left
Bullet.Top = ImgChar.Top
Bullet.Left = ImgChar.Left + ImgChar.Width / 2
Do Until i < 4 Or Bulletway = True
DoEvents
If i < 0 Or Bulletway = True Then
fBlock = False
Bullet.Visible = False
Exit Function
End If
i = i - 1
Bullet.Left = Bullet.Left - 3
Loop
Case 4
i = Bullet.Left
Bullet.Top = ImgChar.Top
Bullet.Left = ImgChar.Left + ImgChar.Width / 2
Do Until i > Me.ScaleWidth Or Bulletway = True
DoEvents
If i > Me.ScaleWidth - 4 Or Bulletway = True Then
fBlock = False
Bullet.Visible = False
Exit Function
End If
i = i - 1
Bullet.Left = Bullet.Left + 3
Loop
End Select
End Function
Function Bulletway() As Boolean
If Me.Point(Bullet.Left, Bullet.Top) = vbBlack Then
Bulletway = True
fBlock = False
Bullet.Visible = False
sndPlaySound App.Path & "\sfx\impact.wav", 1
Else
Bulletway = False
End If
End Function

Private Sub lblAdd_Click()
Dim X As Integer, Y As Integer
X = CInt(Rnd * Me.ScaleWidth)
Y = CInt(Rnd * Me.ScaleHeight)
BitBlt Me.hDC, X, Y, 76, 55, Mask1.hDC, 0, 0, SRCAND
BitBlt Me.hDC, X, Y, 76, 55, Sprite1.hDC, 0, 0, SRCINVERT
Me.Refresh
End Sub
