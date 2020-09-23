VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu"
   ClientHeight    =   7665
   ClientLeft      =   4335
   ClientTop       =   2265
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   6585
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3120
      Top             =   600
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2760
      Top             =   600
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1560
      Top             =   2400
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1560
      Top             =   2040
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   2400
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   2040
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   2400
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   120
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COD .1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   555
      Left            =   2595
      TabIndex        =   3
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "End"
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   1665
      TabIndex        =   2
      Top             =   2880
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   960
      TabIndex        =   1
      Top             =   2880
      Width           =   420
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "START"
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   540
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu help 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idk&
Dim the&
Dim a&
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'resets the form
If KeyCode = vbKeyR Then
    Unload Form1: Load Form1: Form1.Show: the = 0: a = 0: idk = 0
End If
'ends the program
If KeyCode = vbKeyEscape Then End
End Sub
Private Sub Form_Load()
'tells were the label is
With Label4
.Left = Form1.Width
.Top = 0
End With
'tells how big the form is
With Form1
.Width = 6825
.Height = 8505
End With
'tells what timers are on
Timer1.Enabled = True
'tells were the label is
With Label2
.FontSize = 12
.Left = Form1.Width / 2 - (Label2.Width / 2) + 50
.Top = -500
End With
'tells were the label is
With Label3
.FontSize = 12
.Left = Form1.Width / 2 - (Label3.Width / 2) + 50
.Top = -500
End With
'tells what the label proberties are
Label1.FontSize = 12
Label1.Left = Form1.Width / 2 - (Label1.Width / 2) + 50
Label1.Top = -500
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'tells what font size is when the mouse is on the form
Label1.FontSize = 12
Label2.FontSize = 12
Label3.FontSize = 12
End Sub

Private Sub help_Click()
'gives help
MsgBox " Click the start button to start  Click the About button to learn more about the game   Click the End button to end the program  Left Click to Shot Right Click to reload You only have 3 Shots before you have to reload During Game Press escape to end game", 288, "HELP"
End Sub

Private Sub Label1_Click()
'rests varbles and loads form
Unload Form1: Load Form3: Form3.Show
idk = 0
the = 0
a = 0
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'sets the font
If Timer2.Enabled = False And Timer3.Enabled = False Then
    Label1.FontSize = 16
End If
End Sub
Private Sub Label2_Click()
'rests varbles and loads form
Unload Form1: Load Form2: Form2.Show
idk = 0
the = 0
a = 0
End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'sets the font
If Timer4.Enabled = False And Timer5.Enabled = False Then
    Label2.FontSize = 16
End If
End Sub
Private Sub Label3_Click()
'ends program
End
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'sets the font
If Timer6.Enabled = False And Timer7.Enabled = False Then
    Label3.FontSize = 16
End If
End Sub
Private Sub Timer1_Timer()
'sets varaibles
the = Form1.Width / 2
a = Form1.Height / 2 - 250
'makes a circle
Circle (the, a), (idk), RGB(200, 50, 40)
'sets varable
idk = Val(idk) + 50
If idk > 6000 Then Timer1.Enabled = False
End Sub
Private Sub Timer2_Timer()
'moves label1
If Timer1.Enabled = False Then
    Label1.Top = Label1.Top + 100
End If
'wut happens if label is in place
If Label1.Top > Form1.Height + Label1.Height Then
    Timer2.Enabled = False: Timer3.Enabled = True
End If
End Sub
Private Sub Timer3_Timer()
'moves label
Label1.Top = Label1.Top - 100
'wut happens if label is in place
If Label1.Top < Form1.Height / 2 - (Label1.Height + 25) Then
    Timer3.Enabled = False: Timer4.Enabled = True
End If
End Sub
Private Sub Timer4_Timer()
'moves label
Label2.Top = Label2.Top + 100
'wut happens if label1 is in place
If Label2.Top > Form1.Height Then
    Timer5.Enabled = True: Timer4.Enabled = False
End If
End Sub
Private Sub Timer5_Timer()
'moves label
Label2.Top = Label2.Top - 100
'wut happens if label1 is in place
If Label2.Top < (Label1.Top + Label1.Height) + 500 Then
    Timer5.Enabled = False: Timer6.Enabled = True
End If
End Sub
Private Sub Timer6_Timer()
'MOVES LABEL
Label3.Top = Label3.Top + 100
'wut happens if label1 is in place
If Label3.Top > Form1.Height Then
    Timer7.Enabled = True: Timer6.Enabled = False
End If
End Sub
Private Sub Timer7_Timer()
'MOVES LABEL
Label3.Top = Label3.Top - 100
'wut happens if label1 is in place
If Label3.Top < (Label1.Top + Label1.Height) + 1250 Then
    Timer7.Enabled = False: Timer8.Enabled = True
End If
End Sub
Private Sub Timer8_Timer()
'moves label
Label4.Left = Label4.Left - 100
'wut happens if label1 is in place
If Label4.Left + Label4.Width < -120 Then
    Timer9.Enabled = True: Timer8.Enabled = False
End If
End Sub
Private Sub Timer9_Timer()
'moves label
Label4.Left = Label4.Left + 100
'wut happens if label1 is in place
If Label4.Left + Label4.Width > ((Form1.Width / 2) + (Label4.Width / 2)) Then
    Timer9.Enabled = False: Form1.AutoRedraw = True
End If
End Sub
