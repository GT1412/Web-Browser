VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmoffline 
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   9480
      Top             =   5280
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6720
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   9015
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   20055
      ExtentX         =   35375
      ExtentY         =   15901
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton cmdall 
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   0
      Top             =   10200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Saved Web Pages :"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmoffline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdall_Click()

    On Error Resume Next
    With CommonDialog1
        .DialogTitle = "Open File"
        .Filter = "Web page (*.htm;*.html) | *.htm;*.html|" & _
        "All Supported Picture formats|*.gif;*.tif;*.pcd;*.jpg;*.wmf;" & _
        "*.tga;*.jpeg;*.ras;*.png;*.eps;*.bmp;*.pcx|" & _
        "Text formats (*.txt;*.doc)|*.txt;*.doc|" & _
        "All files (*.*)|*.*|"
        .ShowOpen
        .Flags = 5
    WebBrowser1.Navigate2 .FileName
    End With
End Sub


Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
WebBrowser1.Silent = True
Timer1.Enabled = False
End Sub
