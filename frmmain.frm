VERSION 5.00
Begin VB.MDIForm frmmain 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   8535
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13995
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnubrowse 
      Caption         =   "Browse"
      Begin VB.Menu mnuonline 
         Caption         =   "Online"
      End
      Begin VB.Menu mnuoffline 
         Caption         =   "Offline"
      End
   End
   Begin VB.Menu mnusites 
      Caption         =   "Sites"
      Begin VB.Menu mnubrowser 
         Caption         =   "Browser"
      End
      Begin VB.Menu mnuadd 
         Caption         =   "Add"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "About Us"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub mnuadd_Click()
    frmadd.Show
End Sub

Private Sub mnubrowser_Click()
    frmbrowser.Show
End Sub

Private Sub mnuoffline_Click()
    frmoffline.Show
End Sub

Private Sub mnuonline_Click()
    frmonline.Show
End Sub
