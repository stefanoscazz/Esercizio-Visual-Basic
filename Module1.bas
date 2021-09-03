Attribute VB_Name = "modConnection"
Public cn  As ADODB.Connection
Public rs As ADODB.Recordset


Sub main()
Set cn = New ADODB.Connection
cn.ConnectionString = "PROVIDER=Microsoft.ace.oledb.12.0;DATA SOURCE=UserManager.accdb;Jet OLEDB"
cn.Open

If cn.State = adStateOpen Then
FrmPrincipale.Show
End If


End Sub
