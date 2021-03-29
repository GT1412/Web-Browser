VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmonline 
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11280
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   15600
      Top             =   240
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\B.ScIT\Web-Browser\Browser.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\B.ScIT\Web-Browser\Browser.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from sites"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
      Left            =   18000
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdall 
      Caption         =   "List All"
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
      Left            =   6240
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin VB.ComboBox cmbcategory 
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   4
      Top             =   240
      Width           =   4575
   End
   Begin VB.ComboBox cmbwebsite 
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   4575
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   9375
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   20175
      ExtentX         =   35586
      ExtentY         =   16536
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
   Begin VB.Label Label2 
      Caption         =   "Category :"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Web Sites :"
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
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmonline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbcategory_Click()
cmbwebsite.Clear
 str = "SELECT WEB_SITE FROM sites where CATEGORY='" & cmbcategory.Text & "'"
 rs.Open str, cn, 1, 3
 While Not rs.EOF
     cmbwebsite.AddItem (rs.Fields("WEB_SITE"))
    rs.MoveNext
 Wend
 rs.Close
End Sub

Private Sub cmbwebsite_Click()
    WebBrowser1.Navigate "http://" & Trim(cmbwebsite.Text)

End Sub

Private Sub cmdall_Click()
cmbwebsite.Clear
 str = "SELECT distinct WEB_SITE FROM sites"
 rs.Open str, cn, 1, 3
 While Not rs.EOF
     cmbwebsite.AddItem (rs.Fields("WEB_SITE"))
    rs.MoveNext
 Wend
 rs.Close
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
     str = "SELECT distinct CATEGORY FROM sites"
 rs.Open str, cn, 1, 3
 While Not rs.EOF
    cmbcategory.AddItem (rs.Fields("CATEGORY"))
    
    rs.MoveNext
 Wend
 rs.Close
 str = "SELECT distinct WEB_SITE FROM sites"
 rs.Open str, cn, 1, 3
 While Not rs.EOF
     cmbwebsite.AddItem (rs.Fields("WEB_SITE"))
    rs.MoveNext
 Wend
 rs.Close
End Sub
