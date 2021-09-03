VERSION 5.00
Begin VB.Form FrmAggiungi 
   Caption         =   "Aggiungi Utente"
   ClientHeight    =   5910
   ClientLeft      =   1395
   ClientTop       =   1965
   ClientWidth     =   8760
   LinkTopic       =   "Form2"
   ScaleHeight     =   5910
   ScaleWidth      =   8760
   Begin VB.CommandButton cmdSalva 
      Caption         =   "Salva"
      Height          =   495
      Left            =   1000
      TabIndex        =   12
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox txtEmail 
      Height          =   495
      Left            =   3100
      TabIndex        =   11
      Top             =   3960
      Width           =   4900
   End
   Begin VB.TextBox txtDataDiNascita 
      Height          =   495
      Left            =   1000
      TabIndex        =   10
      Top             =   3960
      Width           =   2000
   End
   Begin VB.TextBox txtVia 
      Height          =   495
      Left            =   3100
      TabIndex        =   9
      Top             =   2920
      Width           =   4900
   End
   Begin VB.TextBox txtCap 
      Height          =   495
      Left            =   1000
      TabIndex        =   8
      Top             =   2920
      Width           =   2000
   End
   Begin VB.TextBox txtCognome 
      Height          =   500
      Left            =   1000
      TabIndex        =   2
      Top             =   1920
      Width           =   7000
   End
   Begin VB.TextBox txtNome 
      Height          =   500
      Left            =   1000
      TabIndex        =   0
      Top             =   960
      Width           =   7000
   End
   Begin VB.Label lblEmail 
      Caption         =   "Email :"
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblDataDiNascita 
      Caption         =   "Data di nascita (GG/MM/AAAA) : "
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblVia 
      Caption         =   "Via : "
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblCap 
      Caption         =   "Cap :"
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblCognome 
      Caption         =   "Cognome :"
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblNome 
      Caption         =   "Nome : "
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "FrmAggiungi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSalva_Click()

Dim sql As String
Dim nome As String
Dim cognome As String
Dim data As String
Dim cap As String
Dim email As String
Dim via As String
Dim arrLen As Integer

Set rs = New ADODB.Recordset

nome = txtNome.Text
cognome = txtCognome.Text
data = txtDataDiNascita.Text
arr = Split(data, "/")
cap = txtCap.Text
email = txtEmail.Text
via = txtVia.Text
arrLen = UBound(arr)
If nome = "" Or cognome = "" Or via = "" Or cap = "" Or email = "" Or data = "" Then 'check campi vuoti
MsgBox "Compila tutti i campi"
ElseIf arrLen <> 2 Then 'check array non ha tre elementi
MsgBox "formtao data non valido"
ElseIf Len(arr(0)) <> 2 And Len(arr(1)) <> 2 Or Len(arr(2)) <> 4 Then 'check formato data
MsgBox "formtao data non valido"
Else
sql = "insert into utenti(nome,cognome,cap,email,via,dataDiNascita)VALUES('" & nome & "','" & cognome & "', '" & cap & "','" & email & "','" & via & "','" & data & "')"
rs.Open sql, cn, adOpenDynamic, adLockBatchOptimistic
 FrmPrincipale.Show
 Unload FrmAggiungi
End If
End Sub



Private Sub Form_Load()
Unload FrmPrincipale
End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmPrincipale.Show
End Sub
