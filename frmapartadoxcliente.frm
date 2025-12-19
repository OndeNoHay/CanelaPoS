VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmapartadoxcliente 
   Caption         =   "Artículos Apartados por Cliente"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10875
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAnularApartado 
      Caption         =   "Anular Apartado"
      Height          =   375
      Left            =   7680
      TabIndex        =   3
      Top             =   5400
      Width           =   3015
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir sin Seleccionar para Pagar"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   5400
      Width           =   2655
   End
   Begin VB.CommandButton cmdseleccionar 
      Caption         =   "Seleccionar para pagar"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   2055
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmapartadoxcliente.frx":0000
      Height          =   5295
      Left            =   0
      OleObjectBlob   =   "frmapartadoxcliente.frx":0014
      TabIndex        =   0
      Top             =   0
      Width           =   10935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmapartadoxcliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub cmdseleccionar_Click()
On Error GoTo sehodio
DBGrid1.Col = 0
If DBGrid1.Text = "" Then Exit Sub
ArtApartParaPagar = DBGrid1.Text
IdVentaApartado = ArtApartParaPagar
Unload Me
Exit Sub
sehodio:
MsgBox ("Hay algún error")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo sehodio
If KeyCode = 13 Then
    DBGrid1.Col = 0
    If DBGrid1.Text = "" Then Exit Sub
    ArtApartParaPagar = DBGrid1.Text
    IdVentaApartado = ArtApartParaPagar
    Unload Me
End If
Exit Sub
sehodio:
MsgBox ("Hay algún error")
End Sub

Private Sub Form_Load()
Set Data1.Recordset = RsApartado
End Sub

Private Sub Form_Resize()
DBGrid1.Width = Me.Width
End Sub
