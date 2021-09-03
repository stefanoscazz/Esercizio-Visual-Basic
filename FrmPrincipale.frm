VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPrincipale 
   Caption         =   "Principale"
   ClientHeight    =   5895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   3735
      Left            =   0
      TabIndex        =   7
      Top             =   1680
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6588
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nome"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cognome"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Email"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Cap"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Via"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "DataDiNascita"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdRicerca 
      Caption         =   "Ricerca"
      Height          =   495
      Left            =   7440
      TabIndex        =   6
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtRicerca 
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton smdStampa 
      Caption         =   "Stampa"
      Height          =   855
      Left            =   4080
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdElimina 
      Caption         =   "Elimina"
      Height          =   855
      Left            =   2760
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdModifica 
      Caption         =   "Modifica"
      Height          =   855
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdAggiungi 
      Caption         =   "Aggiungi"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblRicerca 
      Caption         =   "Ricerca : "
      Height          =   255
      Left            =   5880
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "FrmPrincipale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
End Sub
Private Sub cmdAggiungi_Click()
   FrmAggiungi.Show
End Sub
Private Sub cmdElimina_Click()
Dim arr
Dim id As Integer
Dim sql As String
Dim conferma
Dim selectedIndex


'ELIMINARE UTENTE SELEZIONATO
conferma = MsgBox("sei sicuro di voler eliminare l'utente?", vbYesNo)
    If conferma = 6 Then
     id = ListView1.SelectedItem.Text
    sql = "DELETE FROM utenti WHERE id = " & id
    rs.Open sql, cn, adOpenDynamic, adLockBatchOptimistic
    ListView1.ListItems.Remove ListView1.SelectedItem.Index
    End If



End Sub

Private Sub cmdModifica_Click()

'APRIRE FORM MODIFICA
FrmModifica.Show

End Sub

Private Sub cmdRicerca_Click()
Dim sql As String
Dim sqlNumRecord As String
Dim list As ListItem
Dim numRecord As New ADODB.Recordset
Set numRecord = New ADODB.Recordset
arr = Split(txtRicerca.Text)


'STRINGA SQL RITORNA UTENTI
sql = "SELECT * FROM utenti WHERE Nome = '" & txtRicerca.Text & "' "
sql = sql & " OR Cognome = '" & txtRicerca.Text & "' OR Email = '" & txtRicerca.Text & "'"
sql = sql & " OR CAP = '" & txtRicerca.Text & "'  OR Via = '" & txtRicerca.Text & "'  "
If IsDate(txtRicerca.Text) Then
sql = sql & "OR DataDiNascita = #" & txtRicerca.Text & "#"
End If


'STRINGA SQL RITORNA NUMERO UTENTI CHE CORRISPONDONO AL VALORE IMMESSO NEL TEXTBOX
sqlNumRecord = "SELECT COUNT(*) from Utenti WHERE Nome = '" & txtRicerca.Text & "'"
sqlNumRecord = sqlNumRecord & " OR Cognome = '" & txtRicerca.Text & "' OR Email = '" & txtRicerca.Text & "'"
sqlNumRecord = sqlNumRecord & " OR CAP = '" & txtRicerca.Text & "'  OR Via = '" & txtRicerca.Text & "'  "
If IsDate(txtRicerca.Text) Then
sqlNumRecord = sqlNumRecord & " OR  DataDiNascita = #" & txtRicerca.Text & "#"
End If




If txtRicerca.Text = "" Then
MsgBox "Campo ricerca vuoto"
Else
numRecord.Open sqlNumRecord, cn, adOpenDynamic, adLockBatchOptimistic
ListView1.ListItems.Clear
rs.Open sql, cn, adOpenDynamic, adLockBatchOptimistic
For I = 1 To numRecord(0)
Set list = ListView1.ListItems.Add(, , rs!id)
 list.SubItems(1) = rs!nome
 list.SubItems(2) = rs!cognome
 list.SubItems(3) = rs!email
 list.SubItems(4) = rs!cap
 list.SubItems(5) = rs!via
 list.SubItems(6) = rs!DataDiNascita
rs.MoveNext
Next I
numRecord.Close
rs.Close
End If


End Sub

Private Sub Form_Load()
Dim numRecord As ADODB.Recordset
Dim sql As String
Dim list As ListItem
Set rs = New ADODB.Recordset
Set numRecord = New ADODB.Recordset
Dim data As Date


numRecord.Open "SELECT COUNT(*) from Utenti", cn, adOpenDynamic, adLockBatchOptimistic

sql = "select * from utenti"
rs.Open sql, cn, adOpenDynamic, adLockBatchOptimistic

 'LOOP CHE AGGIUNGE TUTTI GLI UTENTI DAL DB A LISTBOX
 For I = 1 To numRecord(0)
 Set list = ListView1.ListItems.Add(, , rs!id)
 list.SubItems(1) = rs!nome
 list.SubItems(2) = rs!cognome
 list.SubItems(3) = rs!email
 list.SubItems(4) = rs!cap
 list.SubItems(5) = rs!via
 list.SubItems(6) = rs!DataDiNascita
 rs.MoveNext
 Next I
rs.Close




End Sub
