VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmVerTablas 
   Caption         =   "Ver Tablas"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdBorrarCampo 
      Caption         =   "Borrar Campo"
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Cmdshow 
      Caption         =   "Mostrar Tabla"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   840
      Width           =   3135
   End
   Begin VB.CommandButton CmdAddField 
      Caption         =   "Agregar campo"
      Height          =   375
      Left            =   9000
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "Eliminar Registro"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   480
      Width           =   3135
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Width           =   3135
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FrmVerTablas.frx":0000
      Height          =   7095
      Left            =   0
      OleObjectBlob   =   "FrmVerTablas.frx":0014
      TabIndex        =   1
      Top             =   1320
      Width           =   11415
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "FrmVerTablas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsTablas As Recordset
Dim Td As TableDef
Dim Registro As Long

Private Sub CmdAddField_Click()
On Error GoTo sehodio
    Dim Tabla As TableDef
    Set Tabla = bdtienda.TableDefs(List1.Text)
    Dim NuevoCampo As Variant
    Dim TipoCampo As Variant
    
    If List1.Text = "" Then Exit Sub
    NuevoCampo = InputBox("Nombre del campo para crear")
    TipoCampo = InputBox("Tipo de campo: dbText, dbDouble, dbCurrency, etc", "Tipo de Campo", "dbText")
    MsgBox TipoCampo
    If NuevoCampo = "" Then Exit Sub
    Select Case TipoCampo
        Case "dbText"
            Set fd = Tabla.CreateField(NuevoCampo, dbText, 255)
        Case "dbInteger"
            Set fd = Tabla.CreateField(NuevoCampo, dbInteger)
        Case "dbDouble"
            Set fd = Tabla.CreateField(NuevoCampo, dbDouble)
        Case "dbDate"
            Set fd = Tabla.CreateField(NuevoCampo, dbDate)
        Case "dbCurrency"
            Set fd = Tabla.CreateField(NuevoCampo, dbCurrency)
        Case "dbLong"
            Set fd = Tabla.CreateField(NuevoCampo, dbLong)
    End Select
    
    Tabla.Fields.Append fd
Exit Sub
sehodio:
    MsgBox (Err.Description & " - " & Err.Number)
End Sub

Private Sub cmdBorrar_Click()
Dim varBmk As Variant

    For Each varBmk In DBGrid1.SelBookmarks
        Data1.Recordset.Bookmark = varBmk
        Data1.Recordset.Delete
    Next

End Sub

Private Sub CmdBorrarCampo_Click()
On Error GoTo sehodio
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Set tdf = bdtienda.TableDefs(List1.Text)
    tdf.Fields.Delete InputBox("Qué campo desea borrar?")
    Set fld = Nothing
    Set tdf = Nothing
    Exit Sub
sehodio:
MsgBox (Err.Description)
End Sub

Private Sub Cmdshow_Click()
On Error Resume Next
Set RsTablas = bdtienda.OpenRecordset(List1.Text)
Set Data1.Recordset = RsTablas
End Sub

Private Sub DBGrid1_Click()
Registro = DBGrid1.Row
End Sub

Private Sub Form_Load()
    For Each Td In bdtienda.TableDefs
        List1.AddItem Td.Name
    Next Td
    If VerTabla <> "" Then
        Set RsTablas = bdtienda.OpenRecordset(VerTabla)
        Set Data1.Recordset = RsTablas
        VerTabla = ""
    End If

End Sub

Private Sub Form_Resize()
DBGrid1.Width = Me.Width - 100
End Sub

