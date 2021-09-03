VERSION 5.00
Begin VB.Form FrmModifica 
   Caption         =   "Modifica Utente"
   ClientHeight    =   5895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNome 
      Height          =   500
      Left            =   1000
      TabIndex        =   6
      Top             =   960
      Width           =   7000
   End
   Begin VB.TextBox txtCognome 
      Height          =   500
      Left            =   1000
      TabIndex        =   5
      Top             =   1920
      Width           =   7000
   End
   Begin VB.TextBox txtCap 
      Height          =   495
      Left            =   1000
      TabIndex        =   4
      Top             =   2920
      Width           =   2000
   End
   Begin VB.TextBox txtVia 
      Height          =   495
      Left            =   3100
      TabIndex        =   3
      Top             =   2920
      Width           =   4900
   End
   Begin VB.TextBox txtDataDiNascita 
      Height          =   495
      Left            =   1000
      TabIndex        =   2
      Top             =   3960
      Width           =   2000
   End
   Begin VB.TextBox txtEmail 
      Height          =   495
      Left            =   3100
      TabIndex        =   1
      Top             =   3960
      Width           =   4900
   End
   Begin VB.CommandButton cmdSalva 
      Caption         =   "Salva"
      Height          =   495
      Left            =   1000
      TabIndex        =   0
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label lblNome 
      Caption         =   "Nome : "
      Height          =   495
      Left            =   1080
      TabIndex        =   12
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblCognome 
      Caption         =   "Cognome :"
      Height          =   495
      Left            =   1080
      TabIndex        =   11
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblCap 
      Caption         =   "Cap :"
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblVia 
      Caption         =   "Via : "
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblDataDiNascita 
      Caption         =   "Data di nascita (GG/MM/AAAA) : "
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblEmail 
      Caption         =   "Email :"
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
End
Attribute VB_Name = "FrmModifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public id As Integer
Private Sub cmdSalva_Click()
Dim sql As String
Dim nome As String
Dim cognome As String
Dim data As String

Dim via As String
Dim cap As String
Dim email As String
Set rs = New ADODB.Recordset

nome = txtNome.Text
cognome = txtCognome.Text
cap = txtCap.Text
via = txtVia.Text
email = txtEmail.Text
data = txtDataDiNascita.Text
arr = Split(data, "/")
sql = "UPDATE utenti SET nome = '" & nome & "',cognome = '" & cognome & "',"
sql = sql & "cap = '" & cap & "',via = '" & via & "', dataDiNascita = '" & txtDataDiNascita.Text & "',"
sql = sql & "email = '" & email & "' WHERE id = " & id 'id variabile pubblica
If nome = "" Or cognome = "" Or data = "" Or via = "" Or cap = "" Or email = "" Or data = "" Then ' check di tutti i campi che non siano vuoti
MsgBox "compila tutti i campi"
ElseIf UBound(arr) <> 2 Then 'check se l'array ha tre elementi
MsgBox "formtao data non valido "
ElseIf Len(arr(0)) <> 2 And Len(arr(1)) <> 2 Or Len(arr(2)) <> 4 Then 'check formato della data
MsgBox "formtao data non valido"
Else
rs.Open sql, cn, adOpenDynamic, adLockBatchOptimistic
FrmPrincipale.Show
Unload FrmModifica
End If





End Sub

Private Sub Form_Load()
Dim selectedIndex As Integer
Dim arr
Dim list As ListItem


'selectedIndex = FrmPrincipale.ListView1.SelectedItem.Text 'indice utente selezionato
'arr = Split(FrmPrincipale.lstUtenti.list(selectedIndex), ",") 'ottengo array dall'utente selezionato
id = FrmPrincipale.ListView1.SelectedItem.Text
txtNome.Text = FrmPrincipale.ListView1.SelectedItem.SubItems(1)
txtCognome.Text = FrmPrincipale.ListView1.SelectedItem.SubItems(2)
txtCap.Text = FrmPrincipale.ListView1.SelectedItem.SubItems(4)
txtVia.Text = FrmPrincipale.ListView1.SelectedItem.SubItems(5)
txtDataDiNascita.Text = FrmPrincipale.ListView1.SelectedItem.SubItems(6)
txtEmail.Text = FrmPrincipale.ListView1.SelectedItem.SubItems(3)

Unload FrmPrincipale
  
 
End Sub


Private Sub Form_Unload(Cancel As Integer)
FrmPrincipale.Show
End Sub

