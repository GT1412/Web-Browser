Attribute VB_Name = "Module1"
Public cn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public str As String

Sub main()
str = "Provider=Microsoft.Jet.oledb.4.0;Data Source=D:\B.ScIT\Web-Browser\Browser.mdb;Persist Security Info= False"
cn.Open str
frmsplash.Show
End Sub
Function CHECKTEXT(K As Integer)
Select Case K
        Case 65 To 90, 97 To 122, 8, 32
                 K = K
        Case Else
                 K = 0
End Select
CHECKTEXT = K
End Function
Function CHECKNUM(K As Integer)
Select Case K
        Case 48 To 57, 8
                 K = K
        Case Else
                 K = 0
End Select
CHECKNUM = K
End Function



