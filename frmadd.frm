VERSION 5.00
Begin VB.Form frmadd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8460
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmadd.frx":0000
   ScaleHeight     =   5025
   ScaleWidth      =   8460
   Begin VB.Frame Frame1 
      Caption         =   "Web Site Detail"
      Height          =   2415
      Left            =   960
      TabIndex        =   0
      Top             =   1200
      Width           =   6615
      Begin VB.TextBox txtwebsite 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox txtcategory 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   1
         Top             =   1080
         Width           =   3495
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
         Left            =   1920
         TabIndex        =   10
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label10 
         Caption         =   "Address :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Category :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   9
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add Website"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmadd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbwebsite_Click()
    str = "select * from sites where Web_site='" & cmbwebsite.Text & "'"
    rs.Open str, cn, 1, 3
      txtwebsite.Text = rs.Fields("WEB_SITE")
        txtcategory.Text = rs.Fields("CATEGORY")
        rs.Close
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmddelete_Click()
If cmddelete.Caption = "Delete" Then
        cmddelete.Caption = "Remove"
        cmdedit.Enabled = False
        cmdnew.Enabled = False
        txtwebsite.Visible = False
        cmbwebsite.Visible = True
        Call filldata
    Else
        
    str = "Delete from sites where Web_site='" & txtwebsite.Text & "'"
    cn.Execute (str)
        MsgBox " Website Detail is Deleted"
        Unload Me
    End If
End Sub

Private Sub cmdedit_Click()
    If cmdedit.Caption = "Edit" Then
        cmdedit.Caption = "Update"
        cmddelete.Enabled = False
        cmdnew.Enabled = False
        txtwebsite.Visible = False
        cmbwebsite.Visible = True
        Call filldata
    Else
        
    str = "select * from sites where Web_site='" & txtwebsite.Text & "'"
    rs.Open str, cn, 1, 3
        
        rs.Fields("WEB_SITE") = txtwebsite.Text
        rs.Fields("CATEGORY") = txtcategory.Text
        rs.Update
        rs.Close
        MsgBox " Website Detail is Updated"
        Unload Me
    End If
    
End Sub
Private Sub filldata()
    str = "select * from Sites"
    rs.Open str, cn, 1, 3
    While Not rs.EOF
        cmbwebsite.AddItem (rs.Fields("Web_site"))
        rs.MoveNext
    Wend
    rs.Close
End Sub
Private Sub cmdnew_Click()
    If cmdnew.Caption = "New" Then
        cmdnew.Caption = "Save"
        cmdedit.Enabled = False
        cmddelete.Enabled = False
    Else
    
        str = "select * from sites"
        rs.Open str, cn, 1, 3
        rs.AddNew
        rs.Fields("WEB_SITE") = txtwebsite.Text
        rs.Fields("CATEGORY") = txtcategory.Text
        rs.Update
        rs.Close
        MsgBox " New Website Address is Added"
        Unload Me
    
    End If
End Sub

