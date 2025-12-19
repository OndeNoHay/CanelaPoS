VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmartidevol 
   BackColor       =   &H00FF0000&
   Caption         =   "Artículos para devolución"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdbuscar 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtbusca 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   5040
      Width           =   11295
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Articulos Devueltos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Articulos en Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.Data Data 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   4440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "rsarticulo"
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lbregistros 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Encontrados: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4920
         TabIndex        =   2
         Top             =   240
         Width           =   5175
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmartidevol.frx":0000
      Height          =   4335
      Left            =   120
      OleObjectBlob   =   "frmartidevol.frx":0013
      TabIndex        =   0
      Top             =   600
      Width           =   11175
   End
End
Attribute VB_Name = "frmartidevol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim msSortCol As String
Dim mbCtrlKey
Dim secsalir As Integer
Dim StrBusca As String
Dim VerEnVenta As Boolean

Private Sub cmdbuscar_Click()
    Dim dummy As String
    On Error GoTo sehodio
   ' MsgBox (RsArticulo.RecordCount)
    If VerEnVenta = True Then
        dummy = "select idart, codigo, tipo, precioventa, fechacompra, " _
        & "devuelto from articulos where idart like '*" & StrBusca & "*'" _
        & "or codigo like '*" & StrBusca & "*' or tipo like " _
        & "'*" & StrBusca & "*' and devuelto like false"
    Else
        dummy = "select idart, codigo, tipo, precioventa, fechacompra, " _
        & "devuelto from artidevueltos"
    End If
    Set RsArticulo = bdtienda.OpenRecordset(dummy)
    RsArticulo.MoveLast
    RsArticulo.MoveFirst
    DBGrid1.WrapCellPointer = True
    Set Data.Recordset = RsArticulo
    'RsArticulo.MoveLast
    'RsArticulo.MoveFirst
    'Data.DatabaseName = App.Path & "canela.mdb"
    'lbregistros.Caption = "Encontrados: " & RsArticulo.RecordCount
    IdArtSelec = 0
    
    DBGrid1.Refresh
    
    
    Exit Sub
sehodio:
    MsgBox (Err.Number & Chr(13) & Err.Description)
    MsgBox ("form_activate")


End Sub

Private Sub cmdcalcula_Click()
    On Error GoTo sehodio
    Dim preccompra As Integer
    preccompra = InputBox("¿Cual es el precio de compra?")
    Dim strprecio As String
    strprecio = "Los posibles precios de venta son:" & Chr(13)
    strprecio = strprecio & "50 %" & Chr(9) & preccompra * 1.5 & Chr(13)
    strprecio = strprecio & "60 %" & Chr(9) & preccompra * 1.6 & Chr(13)
    strprecio = strprecio & "70 %" & Chr(9) & preccompra * 1.7 & Chr(13)
    strprecio = strprecio & "80 %" & Chr(9) & preccompra * 1.8 & Chr(13)
    strprecio = strprecio & "90 %" & Chr(9) & preccompra * 1.9 & Chr(13)
    strprecio = strprecio & Chr(13)
    strprecio = strprecio & "100 %" & Chr(9) & preccompra * 2 & Chr(13)
    strprecio = strprecio & Chr(13)
    strprecio = strprecio & "110 %" & Chr(9) & preccompra * 2.1 & Chr(13)
    strprecio = strprecio & "120 %" & Chr(9) & preccompra * 2.2 & Chr(13)
    strprecio = strprecio & "130 %" & Chr(9) & preccompra * 2.3 & Chr(13)
    strprecio = strprecio & "140 %" & Chr(9) & preccompra * 2.4 & Chr(13)
    MsgBox (strprecio)
        
    Exit Sub
sehodio:
    MsgBox (Err.Number & Chr(13) & Err.Description)
    MsgBox ("cmdcalcula_click")

End Sub

Private Sub cmdpasar_Click()

End Sub

Private Sub DBGrid1_DblClick()
    On Error GoTo sehodio
    If ModoBusca = "articulos" Then
        IdArtSelec = DBGrid1.Text
        Set RsArticulo = bdtienda.OpenRecordset("Select idart, codigo, tipo, precioventa, color, talla from articulos where idart like " & IdArtSelec)
    ElseIf ModoBusca = "clientes" Then
        IdCliSelec = DBGrid1.Text
        Set RsArticulo = bdtienda.OpenRecordset("Select * from clientes where idcliente like " & IdCliSelec)
    End If
    Set Data.Recordset = RsArticulo
    Exit Sub
sehodio:
    MsgBox ("Debe hacer doble click sobre la primera columna")
    MsgBox (Err.Number & Chr(13) & Err.Description)
    MsgBox ("dbgrid1_dblclick")

End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo sehodio
    
    If Data.RecordsetType = vbRSTypeTable Then Exit Sub
    
    'comprueba el uso de la tecla ctrl para orden descendente
    If mbCtrlKey Then
        msSortCol = "[" & Data.Recordset(ColIndex).Name & "] desc"
        mbCtrlKey = 0 'actualiza
    Else
        msSortCol = "[" & Data.Recordset(ColIndex).Name & "]"
    End If
    OrdenaColumnas
    msSortCol = vbNullString 'actualiza

    Exit Sub
sehodio:
    MsgBox (Err.Number & Chr(13) & Err.Description)
    MsgBox ("dbgrid1_headclick")
    
End Sub


'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
''    On Error GoTo sehodio
''
''    Select Case KeyCode
''        Case vbKeyEscape
''            Venta.Show
''            Unload Me
''    End Select
''
''    Exit Sub
''sehodio:
''    MsgBox (Err.Number & Chr(13) & Err.Description)
''    MsgBox ("form_keydown")
'
'End Sub

Private Sub OrdenaColumnas()
    On Error GoTo SortErr

    Dim recRecordset1 As Recordset, recRecordset2 As Recordset
    Dim SortStr As String

    If Data.RecordsetType = vbRSTypeTable Then
        Beep
        MsgBox "Imposible ordenar un Recordset de tipo Table", 48
        Exit Sub
    End If

    Set recRecordset1 = Data.Recordset                        'copia el recordset
    
    If Len(msSortCol) = 0 Then
        SortStr = InputBox("Escriba la columna de orden:")
        If Len(SortStr) = 0 Then Exit Sub
    Else
        SortStr = msSortCol
    End If
    
    Screen.MousePointer = vbHourglass
    recRecordset1.Sort = SortStr
    
    'establece el orden
    Set recRecordset2 = recRecordset1.OpenRecordset(recRecordset1.Type)
    Set Data.Recordset = recRecordset2
    
    Screen.MousePointer = vbDefault
    Exit Sub

SortErr:
    Screen.MousePointer = vbDefault
    MsgBox "Error:" & Err & " " & Err.Description
    MsgBox ("ordenacolumnas")

End Sub


Public Sub PoneTodosLosCampos()
    On Error GoTo sehodio
    If AñadeArtic = True Then
        Set RsArticulo = bdtienda.OpenRecordset("select idart, codigo, tipo, preciocompra, precioventa, talla, extra, fechacompra from articulos")
        DBGrid1.WrapCellPointer = True
    End If
    Set Data.Recordset = RsArticulo
    IdArtSelec = 0
    Exit Sub
sehodio:
    MsgBox (Err.Number & Chr(13) & Err.Description)
    MsgBox ("form_activate")


End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdbuscar_Click
End Sub

Private Sub Form_Load()
    VerEnVenta = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
DBGrid1.Width = Me.Width - 300
DBGrid1.Height = Me.Height - 2000
Frame1.Top = Me.Height - 1300
Frame1.Width = Me.Width - 300
End Sub



Private Sub Option1_Click(Index As Integer)
If Index = 0 Then VerEnVenta = True Else: VerEnVenta = False

End Sub

Private Sub txtbusca_Change()
    StrBusca = txtbusca.Text
End Sub
