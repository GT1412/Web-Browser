VERSION 5.00
Begin VB.Form frmsplash 
   Caption         =   "Form1"
   ClientHeight    =   9465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19860
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   20.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9465
   ScaleWidth      =   19860
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   5760
      Top             =   4080
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Miss.Sujata Sonule"
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   14880
      TabIndex        =   4
      Top             =   8400
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Miss.Komal Dhanorkar"
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   14880
      TabIndex        =   3
      Top             =   7440
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mr.Gaurav Talodhikar"
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   14760
      TabIndex        =   2
      Top             =   6480
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Devloped By:"
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   12960
      TabIndex        =   1
      Top             =   5520
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Web Browser"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2295
      Left            =   7080
      TabIndex        =   0
      Top             =   1560
      Width           =   6375
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   11880
      Left            =   -360
      Picture         =   "frmsplash.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20730
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    Unload Me
    frmlogin.Show
End Sub
