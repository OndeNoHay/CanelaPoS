VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmBuscaArticulos 
   Caption         =   "Busca Articulos Iguales"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10515
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FrmBuscaArticulos.frx":0000
      Height          =   4575
      Left            =   120
      OleObjectBlob   =   "FrmBuscaArticulos.frx":0014
      TabIndex        =   0
      Top             =   120
      Width           =   10335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4680
      Width           =   1215
   End
End
Attribute VB_Name = "FrmBuscaArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DBGrid1_Click()
    IdArtBuscado = DBGrid1.Text
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        IdArtBuscado = DBGrid1.Text
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    Set Data1.Recordset = RsBuscaArticulos

End Sub
