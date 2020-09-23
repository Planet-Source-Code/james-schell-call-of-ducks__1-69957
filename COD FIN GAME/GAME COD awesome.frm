VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Game"
   ClientHeight    =   6570
   ClientLeft      =   3840
   ClientTop       =   2325
   ClientWidth     =   7665
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   Picture         =   "GAME COD awesome.frx":0000
   ScaleHeight     =   6570
   ScaleWidth      =   7665
   Begin VB.Frame Frame3 
      BackColor       =   &H0000FF00&
      Caption         =   "FINAL REVIEW"
      Height          =   375
      Left            =   120
      MousePointer    =   2  'Cross
      TabIndex        =   9
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MAIN MENU"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   2040
         TabIndex        =   18
         Top             =   5520
         Width           =   1515
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   4080
         TabIndex        =   17
         Top             =   4440
         Width           =   120
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   3600
         TabIndex        =   16
         Top             =   3480
         Width           =   120
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   3600
         TabIndex        =   15
         Top             =   2400
         Width           =   120
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   2640
         TabIndex        =   14
         Top             =   1320
         Width           =   120
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL SCORE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   1080
         TabIndex        =   13
         Top             =   4440
         Width           =   2835
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DIFFICULTY:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   1080
         TabIndex        =   12
         Top             =   3480
         Width           =   2340
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MULTIPLIER:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   1080
         TabIndex        =   11
         Top             =   2400
         Width           =   2370
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SCORE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   1080
         TabIndex        =   10
         Top             =   1320
         Width           =   1470
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FF00&
      Caption         =   "Difficulty"
      Height          =   255
      Left            =   -360
      TabIndex        =   3
      Top             =   4080
      Width           =   1455
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   4440
         TabIndex        =   19
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Veteran"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         TabIndex        =   6
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         TabIndex        =   5
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recruit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         TabIndex        =   4
         Top             =   1320
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0000FF00&
      Caption         =   "SCORE"
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   5640
      Width           =   1575
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   90
      End
   End
   Begin VB.Timer Timer18 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5040
      Top             =   3600
   End
   Begin VB.Timer Timer17 
      Interval        =   1
      Left            =   5040
      Top             =   1200
   End
   Begin VB.Timer Timer16 
      Interval        =   1
      Left            =   5520
      Top             =   720
   End
   Begin VB.Timer Timer15 
      Interval        =   1
      Left            =   5040
      Top             =   720
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5520
      Top             =   3120
   End
   Begin VB.Timer Timer13 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5040
      Top             =   3120
   End
   Begin VB.Timer Timer12 
      Interval        =   500
      Left            =   3720
      Top             =   1680
   End
   Begin VB.Timer Timer11 
      Interval        =   500
      Left            =   3240
      Top             =   1200
   End
   Begin VB.Timer Timer10 
      Interval        =   500
      Left            =   3720
      Top             =   1200
   End
   Begin VB.Timer Timer9 
      Interval        =   1
      Left            =   120
      Top             =   600
   End
   Begin VB.Timer Timer8 
      Interval        =   500
      Left            =   3720
      Top             =   720
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5040
      Top             =   2640
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5520
      Top             =   2640
   End
   Begin VB.Timer Timer5 
      Interval        =   1
      Left            =   5520
      Top             =   240
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   5040
      Top             =   240
   End
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   3240
      Top             =   720
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      ScaleHeight     =   255
      ScaleWidth      =   855
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.Image Image15 
      Height          =   495
      Left            =   4920
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Image Image14 
      Height          =   495
      Left            =   4920
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Image Image13 
      Height          =   495
      Left            =   4920
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   3360
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   6960
      TabIndex        =   1
      Top             =   5880
      Width           =   90
   End
   Begin VB.Image Image12 
      Height          =   975
      Left            =   6480
      Picture         =   "GAME COD awesome.frx":2BA1
      Top             =   2400
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Image Image11 
      Height          =   495
      Left            =   4920
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image Image10 
      Height          =   495
      Left            =   4920
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Image9 
      Height          =   735
      Left            =   6600
      Picture         =   "GAME COD awesome.frx":35CC
      Top             =   1560
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image Image8 
      Height          =   735
      Left            =   6600
      Picture         =   "GAME COD awesome.frx":3BF0
      Top             =   1680
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image Image7 
      Height          =   735
      Left            =   6600
      Picture         =   "GAME COD awesome.frx":42A7
      Top             =   960
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image Image6 
      Height          =   615
      Left            =   6600
      Picture         =   "GAME COD awesome.frx":48C9
      Top             =   960
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Image B0 
      Height          =   210
      Left            =   2400
      Picture         =   "GAME COD awesome.frx":4F63
      Top             =   120
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Image B1 
      Height          =   210
      Left            =   3360
      Picture         =   "GAME COD awesome.frx":5222
      Top             =   120
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Image B2 
      Height          =   210
      Left            =   2400
      Picture         =   "GAME COD awesome.frx":5540
      Top             =   360
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Image B3 
      Height          =   210
      Left            =   3360
      Picture         =   "GAME COD awesome.frx":58AB
      Top             =   360
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Image Image1 
      Height          =   1260
      Left            =   5880
      Picture         =   "GAME COD awesome.frx":5C5D
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   780
      Left            =   5760
      Picture         =   "GAME COD awesome.frx":669B
      Top             =   5280
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image Image5 
      Height          =   2535
      Left            =   0
      Picture         =   "GAME COD awesome.frx":6CB2
      Top             =   4200
      Width           =   7680
   End
   Begin VB.Image Image4 
      Height          =   1080
      Left            =   1200
      Picture         =   "GAME COD awesome.frx":C1D2
      Top             =   3240
      Width           =   1320
   End
   Begin VB.Image Image3 
      Height          =   4455
      Left            =   0
      Picture         =   "GAME COD awesome.frx":CEC5
      Top             =   0
      Width           =   7680
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'dims my sound files to play them
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
    Const SND_SYNC = &H0
    Const SND_ASYNC = &H1
    Const SND_NODEFAULT = &H2
    Const SND_LOOP = &H8
    Const SND_NOSTOP = &H10
Dim fly As Long
Dim fly2 As Long
Dim fly3 As Long
Dim fly4 As Long
Dim fly5 As Long
Dim a As Long
Dim time As Long
Dim p As Boolean
'makes a array of the varible
Dim flying(5) As Boolean
Dim diff As Long
Dim diffp As String
Dim miss%
Dim pts%
Dim ptsd%
'tells how much each duck is worth
Private Sub Pts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pts = pts + 10
End Sub
'if u miss then takes away time
Private Sub miss_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If a < 3 And Button = 1 Then time = time - miss
End Sub
'makes it display # of bullets and plays sound for the shot, empty, reload
Private Sub All_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If p = True Then Exit Sub
If a > 3 Then GoTo bob1
If Button = 1 Then Image2.Visible = True: soundfile$ = "COD\shotgun.wav": wFlags! = SND_ASYNC Or SND_NODEFAULT: HaHa = sndPlaySound(soundfile$, wFlags!):  a = a + 1
If a > 3 Then a = a + 1
bob1:
If Button = 1 And a > 3 Then soundfile$ = "COD\Empty.wav": wFlags! = SND_ASYNC Or SND_NODEFAULT: HaHa = sndPlaySound(soundfile$, wFlags!)
If Button = 2 And a > 0 Then a = 0: soundfile$ = "COD\reload.wav": wFlags! = SND_ASYNC Or SND_NODEFAULT: HaHa = sndPlaySound(soundfile$, wFlags!)
End Sub
'makes it so the flash does not stay on
Private Sub All_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Image2.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'resets all varilbes and the form3
If KeyCode = vbKeyR Then
    Unload Form3: Load Form3: Form3.Show
    pts = 0
    time = 1
    ptsd = 0
    fly = 0
    fly1 = 0
    fly2 = 0
    fly3 = 0
    fly4 = 0
    fly5 = 0
    miss = 0
    diff = 0
    diffp = ""
    a = 0
End If
'pauses the game
If KeyCode = vbKeyP And p = False Then
    p = True
    Call pause
Else
    Call unpause
End If
'exits the game
If KeyCode = vbKeyEscape Then End
End Sub

Private Sub Form_Load()
'randomizes the game
Randomize
'sets time to 1
time = 1
'sets frame1's property's
With Frame1
.Height = Form3.Height
.Width = Form3.Width
.Left = 0
.Top = 0
End With
'sets frame3's property's
With Frame3
.Height = Form3.Height
.Width = Form3.Width
.Left = 0
.Top = 0
End With
'tell where each image starts
Image13.Top = 0 - Image13.Height
Image13.Left = 6000
Label1.Left = Label1.Left - 10
Image11.Left = 0 - Image10.Width
Image11.Top = 4200
Image10.Left = Form3.Width
Image10.Top = 4200
Image14.Left = 0 - Image14.Width
Image14.Top = Int(Rnd * 3600) + 1
Image15.Left = Form3.Width + Image15.Width
Image15.Top = Int(Rnd * 3600) + 1
a = 0
'makes 0,1,2,3,4flying true at load
For X = 0 To 4
flying(X) = True
Next
'pauses the game so you can't start right away
p = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'calls the Private sub All_MouseMove
Call All_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'calls private sub All_MouseUp
Call All_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'make font 12 if mouse is on the frame
Label3.FontSize = 12
Label4.FontSize = 12
Label5.FontSize = 12
'clears label16 when it is not being used
Label16.Caption = ""
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'makes font 14 if on label
Label15.FontSize = 12
End Sub

Private Sub Image10_Click()
'if the pause is true then it exits this sub
If p = True Then Exit Sub
'kills the duck if u have ammo
If a > 3 Then GoTo jim
flying(0) = False
Image10.Picture = Image12.Picture
Timer4.Enabled = False: Timer6.Enabled = True: Timer3.Enabled = False
jim:
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'calls the Private sub All_MouseMove
Call All_MouseDown(Button, Shift, X, Y)
'calls the private sub Pts_MouseDown
Call Pts_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Image10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'calls private sub All_MouseUp
Call All_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Image11_Click()
'if the pause is true then it exits this sub
If p = True Then Exit Sub
'kills the duck if u have ammo
If a > 3 Then GoTo jim1
flying(1) = False
Image11.Picture = Image12.Picture
Timer5.Enabled = False: Timer8.Enabled = False: Timer7.Enabled = True
jim1:
End Sub

Private Sub Image11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'calls the Private sub All_MouseMove
Call All_MouseDown(Button, Shift, X, Y)
'calls the private sub Pts_MouseDown
Call Pts_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Image11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'calls private sub All_MouseUp
Call All_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Image13_Click()
'if the pause is true then it exits this sub
If p = True Then Exit Sub
'kills the duck if u have ammo
If a > 3 Then GoTo jim2
flying(2) = False
Image13.Picture = Image12.Picture
Timer17.Enabled = False: Timer18.Enabled = True: Timer11.Enabled = False
jim2:
End Sub

Private Sub Image13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'calls the Private sub All_MouseMove
Call All_MouseDown(Button, Shift, X, Y)
'calls the private sub Pts_MouseDown
Call Pts_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Image13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'calls private sub All_MouseUp
Call All_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Image14_Click()
'if the pause is true then it exits this sub
If p = True Then Exit Sub
'kills the duck if u have ammo
If a > 3 Then GoTo jim3
flying(3) = False
Image14.Picture = Image12.Picture
Timer15.Enabled = False: Timer13.Enabled = True: Timer10.Enabled = False
jim3:
End Sub

Private Sub Image14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'calls the Private sub All_MouseMove
Call All_MouseDown(Button, Shift, X, Y)
'calls the private sub Pts_MouseDown
Call Pts_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Image14_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'calls private sub All_MouseUp
Call All_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Image15_Click()
'if the pause is true then it exits this sub
If p = True Then Exit Sub
'kills the duck if u have ammo
If a > 3 Then GoTo jim4
flying(4) = False
Image15.Picture = Image12.Picture
Timer16.Enabled = False: Timer14.Enabled = True: Timer12.Enabled = False
jim4:
End Sub

Private Sub Image15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'calls the Private sub All_MouseMove
Call All_MouseDown(Button, Shift, X, Y)
'calls the private sub Pts_MouseDown
Call Pts_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Image15_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'calls private sub All_MouseUp
Call All_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'calls the Private sub All_MouseMove
Call All_MouseDown(Button, Shift, X, Y)
'calls the private sub Pts_MouseDown
Call miss_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'calls private sub All_MouseUp
Call All_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'calls the Private sub All_MouseMove
Call All_MouseDown(Button, Shift, X, Y)
'calls the private sub Pts_MouseDown
Call miss_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'calls private sub All_MouseUp
Call All_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'calls the Private sub All_MouseMove
Call All_MouseDown(Button, Shift, X, Y)
'calls the private sub Pts_MouseDown
Call miss_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'calls private sub All_MouseUp
Call All_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Label15_Click()
'resets varibles and loads form1
pts = 0
    time = 1
    ptsd = 0
    fly = 0
    fly1 = 0
    fly2 = 0
    fly3 = 0
    fly4 = 0
    fly5 = 0
    miss = 0
    diff = 0
    diffp = ""
    a = 0
    Unload Form3: Load Form1: Form1.Show
End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'makes font 14 if mouse is on label
Label15.FontSize = 14
End Sub

Private Sub Label3_Click()
'sets the setting for the recruit level
Frame1.Visible = False
diff = 45
time = 60
miss = 1
ptsd = 1
Call unpause
diffp = "Recruit"
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'makes font 14 if mouse is on label
Label3.FontSize = 14
'Gives a description for the recruit level
Label16.Caption = "So easy a Caveman could do it" & vbNewLine & "If you miss you lose a second for each shot missed."

End Sub

Private Sub Label4_Click()
'sets the setting for the normal level
Frame1.Visible = False
diff = 65
time = 45
miss = 2
ptsd = 2
Call unpause
diffp = "Normal"
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'makes font 14 if mouse is on label
Label4.FontSize = 14
'Gives a description for the normal level
Label16.Caption = "Avg. noob could do this" & vbNewLine & "If you miss you lose two seconds for each shot missed."
End Sub

Private Sub Label5_Click()
'sets the setting for the Veteran level
Frame1.Visible = False
diff = 105
time = 30
miss = 4
ptsd = 3
Call unpause
diffp = "Veteran"
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'makes font 14 if mouse is on label
Label5.FontSize = 14
'Gives a description for the verteran level
Label16.Caption = "You well not survive" & vbNewLine & "If you miss you lose four seconds for each shot missed."
End Sub

Private Sub Timer1_Timer()
'counts down time and displays the time in the label
time = time - 1
Label1.Caption = time
End Sub
'sets the flapping of the wings and image flash
Private Sub Timer10_Timer()
If flying(3) = False Then Exit Sub
fly4 = fly4 + 1
If (fly4 / 2) = Int(fly4 / 2) Then
    Image14.Picture = Image6.Picture
Else:
    Image14.Picture = Image7.Picture
End If
End Sub
'sets the flapping of the wings and image flash
Private Sub Timer11_Timer()
If flying(2) = False Then Exit Sub
fly3 = fly3 + 1
If (fly3 / 2) = Int(fly3 / 2) Then
    Image13.Picture = Image8.Picture
Else:
    Image13.Picture = Image9.Picture
End If
End Sub
'sets the flapping of the wings and image flash
Private Sub Timer12_Timer()
If flying(4) = False Then Exit Sub
fly5 = fly5 + 1
If (fly5 / 2) = Int(fly5 / 2) Then
    Image15.Picture = Image8.Picture
Else:
    Image15.Picture = Image9.Picture
End If
End Sub
'sends image14 down
Private Sub Timer13_Timer()
Image14.Top = Image14.Top + 45
End Sub

Private Sub Timer14_Timer()
'sends image15 down
Image15.Top = Image15.Top + 45
End Sub

Private Sub Timer15_Timer()
'moves the image
Image14.Left = Image14.Left + diff
End Sub

Private Sub Timer16_Timer()
'moves the image
Image15.Left = Image15.Left - diff
End Sub

Private Sub Timer17_Timer()
'moves the image
Image13.Left = Image13.Left - diff
Image13.Top = Image13.Top + diff
End Sub
'moves the image
Private Sub Timer18_Timer()
Image13.Top = Image13.Top + 45
End Sub

Private Sub Timer2_Timer()
'respawanes the ducks after they die
If Image10.Top + Image10.Height < 0 Then
    Image10.Top = 4200: Image10.Left = 0 - Image10.Width
End If
If Image11.Top + Image11.Height < 0 Then
    Image11.Top = 4200: Image11.Left = Form3.Width
End If
If Image14.Left > Form3.Width Then
    Image14.Left = 0 - Image14.Width
End If
If Image15.Left + Image15.Width < 0 Then
    Image15.Left = Form3.Width
End If
If Image13.Picture = Image12.Picture Then GoTo bob6
If Image13.Top > Image5.Top Then
     Image13.Top = 0 - Image13.Height: Image13.Left = 6000
End If
'sets windowstate and sets what picture is up if a = 0-3
bob6:
Form3.WindowState = 0
Select Case a
Case 0
Picture1.Picture = B3
Case 1
Picture1.Picture = B2
Case 2
Picture1.Picture = B1
Case 3
Picture1.Picture = B0
End Select

End Sub
'sets the flapping of the wings and image flash
Private Sub Timer3_Timer()
If flying(0) = False Then Exit Sub
fly = fly + 1
If (fly / 2) = Int(fly / 2) Then
    Image10.Picture = Image6.Picture
Else:
    Image10.Picture = Image7.Picture
End If
End Sub
'moves the image
Private Sub Timer4_Timer()
Image10.Left = Image10.Left + diff
Image10.Top = Image10.Top - diff
End Sub
'moves the image
Private Sub Timer5_Timer()
Image11.Left = Image11.Left - diff
Image11.Top = Image11.Top - diff
End Sub
'moves the image
Private Sub Timer6_Timer()
Image10.Top = Image10.Top + 45
End Sub
'moves the image
Private Sub Timer7_Timer()
Image11.Top = Image11.Top + 45
End Sub
'sets the flapping of the wings and image flash
Private Sub Timer8_Timer()
If flying(1) = False Then Exit Sub
fly2 = fly2 + 1
If (fly2 / 2) = Int(fly2 / 2) Then
    Image11.Picture = Image8.Picture
Else:
    Image11.Picture = Image9.Picture
End If
End Sub
'makes the game pasued by turning off timers and setting labels
Private Sub pause()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Timer10.Enabled = False
Timer11.Enabled = False
Timer12.Enabled = False
Timer13.Enabled = False
Timer14.Enabled = False
Timer15.Enabled = False
Timer16.Enabled = False
Timer17.Enabled = False
Timer18.Enabled = False
With Label2
.Caption = "PAUSE"
.Visible = True
End With
End Sub
'makes unpause if the game is paused and unsets the label
Private Sub unpause()
p = False
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
Timer10.Enabled = True
Timer11.Enabled = True
Timer12.Enabled = True
Timer15.Enabled = True
Timer16.Enabled = True
Timer17.Enabled = True
If flying(0) = False Then
    Timer6.Enabled = True
Else
    Timer6.Enabled = False
End If
If flying(1) = False Then
    Timer7.Enabled = True
Else
    Timer7.Enabled = False
End If
If flying(2) = False Then
    Timer13.Enabled = True
Else
    Timer13.Enabled = False
End If
If flying(3) = False Then
    Timer14.Enabled = True
Else
    Timer14.Enabled = False
End If
If flying(4) = False Then
    Timer18.Enabled = True
Else
    Timer18.Enabled = False
End If
Timer8.Enabled = True
Label2.Visible = False
End Sub

Private Sub Timer9_Timer()
'turns back on the wing flapping and turnes off the duck dieing
If Image11.Top > Form3.Height Then
    Image11.Left = 0 - Image10.Width: Image11.Top = 4200: Timer7.Enabled = False: Timer8.Enabled = True: flying(1) = True: Timer5.Enabled = True
End If
If Image10.Top > Form3.Height Then
    Image10.Left = Form3.Width: Image10.Top = 4200: Timer6.Enabled = False: Timer3.Enabled = True: flying(0) = True: Timer4.Enabled = True
End If
If Image14.Top > Form3.Height Then
    Image14.Left = 0 - Image14.Width: Image14.Top = Int(Rnd * 3600) + 1: Timer13.Enabled = False: Timer10.Enabled = True: flying(3) = True: Timer15.Enabled = True
End If
If Image15.Top > Form3.Height Then
    Image15.Left = Form3.Width + Image15.Width: Image15.Top = Int(Rnd * 3600) + 1: Timer14.Enabled = False: Timer12.Enabled = True: flying(4) = True: Timer16.Enabled = True
End If
If Image13.Top > Form3.Height Then
    Image13.Top = 0 - Image13.Height: Image13.Left = 6000: Timer18.Enabled = False: Timer11.Enabled = True: flying(2) = True: Timer13.Enabled = True
End If
'sets the final Score labels
Label13.Caption = diffp
Label11.Caption = pts
Label12.Caption = ptsd
Label14.Caption = pts * ptsd
'if paused and a duck is dead the duck well keep dieing
If p = True Then Call pause
If flying(0) = False Then Image10.Picture = Image12.Picture: Image10.Top = Image10.Top + 45
If flying(1) = False Then Image11.Picture = Image12.Picture: Image11.Top = Image11.Top + 45
If flying(2) = False Then Image13.Picture = Image12.Picture: Image13.Top = Image13.Top + 45
If flying(3) = False Then Image14.Picture = Image12.Picture: Image14.Top = Image14.Top + 45
If flying(4) = False Then Image15.Picture = Image12.Picture: Image15.Top = Image15.Top + 45
'shows your score while you play
Label6.Caption = pts
'makes it so when the time runs out then the game is over and the final score comes up
If time <= 0 Then
    Label2.Caption = "GAME OVER": Label2.Visible = True: Timer1.Enabled = False
End If
If Label2.Caption = "GAME OVER" Then Frame3.Visible = True
End Sub
