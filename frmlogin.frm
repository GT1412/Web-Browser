VERSION 5.00
Begin VB.Form frmlogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3450
   ClientLeft      =   7260
   ClientTop       =   4530
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4560
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4335
      Begin VB.CommandButton cmdexit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   4
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton cmdlogin 
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtpwd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "@"
         TabIndex        =   2
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtuser 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name :"
         BeginProperty Font 
            Name            =   "Nirmala UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   2535
         Left            =   0
         Picture         =   "frmlogin.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4335
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Login Details"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdlogin_Click()
If txtuser.Text = "a" And txtpwd.Text = "b" Then
            Unload Me
            frmmain.Show
            
    Else
        MsgBox " Enter Correct User name and Password", vbCritical + vbOKCancel
        
    End If

End Sub

