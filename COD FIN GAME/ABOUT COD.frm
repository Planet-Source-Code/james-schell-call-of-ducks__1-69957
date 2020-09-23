VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00404040&
   Caption         =   "About"
   ClientHeight    =   3060
   ClientLeft      =   3300
   ClientTop       =   3645
   ClientWidth     =   7755
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   7755
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   585
      Left            =   1035
      TabIndex        =   1
      Top             =   2400
      Width           =   2640
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'says what the label says
Label1.Caption = "YOU ARE DirtDiver" & vbNewLine & "EMBARKING TO TAKE OUT SPECIAL FORCES FROM" & vbNewLine & "DUCK LAND YOU MUST SAVE THE" & vbNewLine & "DUCK LAND AND DO IT IN RECORD TIME"
End Sub

Private Sub Label2_Click()
'loads form
Unload Form2: Load Form1: Form1.Show
End Sub
