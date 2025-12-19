VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmarticulos 
   BackColor       =   &H00FF0000&
   Caption         =   "Resultado de artículos"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12480
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   12480
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11760
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   11640
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   5040
      Width           =   12375
      Begin VB.CommandButton CmdExport 
         BackColor       =   &H00FF8080&
         Caption         =   "Exportar CSV"
         Height          =   495
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdvertodo 
         BackColor       =   &H00FF8080&
         Caption         =   "Ver Todos"
         Height          =   495
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdcaja 
         BackColor       =   &H00FF8080&
         Caption         =   "Caja Diaria"
         Height          =   495
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdaddart 
         BackColor       =   &H00FF8080&
         Caption         =   "&AddArt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdllamar 
         BackColor       =   &H00FF8080&
         Caption         =   "Llamar Proveedor"
         Height          =   495
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdsalir 
         BackColor       =   &H00FF8080&
         Caption         =   "Sa&lir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11280
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.Data Data 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   1680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "rsarticulo"
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdpasar 
         BackColor       =   &H00FF8080&
         Caption         =   "&Seleccionar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4320
         Top             =   0
      End
      Begin VB.CommandButton cmdZoom 
         BackColor       =   &H00FF8080&
         Caption         =   "&Zoom"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   975
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
         Left            =   1920
         TabIndex        =   4
         Top             =   240
         Width           =   2775
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmarticulos.frx":0000
      Height          =   5055
      Left            =   0
      OleObjectBlob   =   "frmarticulos.frx":0013
      TabIndex        =   0
      Top             =   0
      Width           =   11175
   End
   Begin MSCommLib.MSComm Msc 
      Left            =   10920
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "frmarticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Dim msSortCol As String
Dim mbCtrlKey
Dim secsalir As Integer
Dim numproveedor As String


Private Sub cmdaddart_Click()
    FrmAddArt2.Show
End Sub



Private Sub cmdcaja_Click()
    FrmCajaDiaria.Show

End Sub



Private Sub CmdExport_Click()
    Screen.MousePointer = vbHourglass
    'ExportTOExcel CommonDialog1, adoPrimaryRS
    Dim Exportados As String
    CommonDialog1.Filter = "csv(*.csv)|*.csv"
    CommonDialog1.InitDir = App.Path
    CommonDialog1.FileName = "ListadoArticulos" & "-" & Day(Date) & "-" & Month(Date) & "-" & Year(Date)
    CommonDialog1.ShowSave
    If CommonDialog1.FileName = "" Then Exit Sub
    'locate is position of the last "\" length is the full path length
    
        
    Exportados = CommonDialog1.FileName
    If Exists(Exportados) Then Kill Exportados
    Dim RsExportados As Recordset
    Dim fecha As Date
    Dim Ssql As String
    fecha = InputBox("¿Desde qué fecha desea exportar?", "Fecha para exportar datos", Date)
    Ssql = "Select * from articulos  where fechacompra between #" & fecha & "# and #" & Date & "#"
    MsgBox Ssql
    Set RsExportados = bdtienda.OpenRecordset(Ssql)
    MsgBox (RsExportados.RecordCount)
    If ExportToCSV(RsExportados, Exportados) Then
        MsgBox "Los datos se exportaron correctamente"
    Else
        MsgBox "Los datos NO SE PUDIERON EXPORTAR!"
    End If
    RsExportados.Close
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdllamar_Click()
    On Error Resume Next
    Dim rsnumprov As Recordset
    numproveedor = Left(DBGrid1.Columns(1).Text, 3)

    Set rsnumprov = bdtienda.OpenRecordset("proveedor")
    Do Until rsnumprov.EOF
        If UCase(numproveedor) = UCase(rsnumprov!Codigo) Then
            numproveedor = rsnumprov!telefono
        End If
        rsnumprov.MoveNext
    Loop
    'MsgBox (numproveedor)
    If hablando = False Then
        Dim numtel As String
        numtel = numproveedor
        If numtel = "" Then Exit Sub
        If Len(numtel) = 6 Then numtel = "959" & numtel
        AbrePuerto
        Msc.Output = "ATDT" & numtel & ";" & vbCr
        hablando = True
        cmdllamar.Caption = "Colgar"
    Else
        Msc.PortOpen = False
        cmdllamar.Caption = "Llamar"
        hablando = False
    End If
End Sub
Private Sub AbrePuerto()
    On Error GoTo sehodio
    Msc.CommPort = "2"
    Msc.Settings = "28800,N,8,1"
    Msc.PortOpen = True
    Exit Sub
sehodio:
    MsgBox (Err.Number)
    MsgBox (Err.Description)
End Sub

Private Sub cmdpasar_Click()
    On Error GoTo sehodio
    Unload Me
    
    'MsgBox (IdArtSelec)
    If ModoBusca = "articulos" Then
        NumArtVend = NumArtVend + 1
        Venta.Show
        Venta.PoneArticulos
    ElseIf ModoBusca = "clientes" Then
        Venta.Show
        Venta.PoneClientes
        Venta.BuscaVentaApartada
        If VentaApartado = False Then
            Venta.cmdarticulo.Enabled = True
            Venta.cmdarticulo.SetFocus
        End If
    End If
    Venta.Show
    
    Exit Sub
sehodio:
    MsgBox (Err.Number & Chr(13) & Err.Description)
    MsgBox ("cmdpasar_click")

End Sub


Private Sub cmdsalir_Click()
Unload Me
End Sub


Private Sub cmdvertodo_Click()
    If cmdvertodo.Caption = "Ver Todos" Then
        Set RsArticulo = bdtienda.OpenRecordset("select idart, codigo, tipo, precioventa, fechacompra, vendido, apartado, foto, fotoweb from articulos order by idart")
        cmdvertodo.Caption = "Ver Disponibles"
    Else
        SqlArticulos = "Select idart, codigo, tipo, precioventa, color, talla, extra " _
        & "from articulos where vendido = false and apartado = false and(codigo " _
        & "like '*" & CodigoBusca & "*' or precioventa like '*" & CodigoBusca & "*' or " _
        & "talla like '*" & CodigoBusca & "*' or tipo like '*" & CodigoBusca & "*') order by codigo"
        Set RsArticulo = bdtienda.OpenRecordset(SqlArticulos)
        cmdvertodo.Caption = "Ver Todos"
    End If
    Set Data.Recordset = RsArticulo
    RsArticulo.MoveLast
    RsArticulo.MoveFirst
    'Data.DatabaseName = App.Path & "canela.mdb"
    lbregistros.Caption = "Encontrados: " & RsArticulo.RecordCount
    IdArtSelec = 0
    
    DBGrid1.Refresh
    
End Sub

Private Sub cmdZoom_Click()
    frmZoom.Show
'    IdArtSelec = DBGrid1.Text
End Sub

Private Sub Command1_Click()
With RsArticulo
    .MoveFirst
    Do While .EOF = False
        .Edit
        !vendido = True
        .Update
        
        .MoveNext
    Loop
End With
'Text1.Text = ""
End Sub


Private Sub DBGrid1_DblClick()
    On Error GoTo sehodio
    If ModoBusca = "articulos" Then
        IdArtSelec = DBGrid1.Text
        'numproveedor = Left(DBGrid1.Columns(2).Text, 3)
        'MsgBox (numproveedor)
        Set RsArticulo = bdtienda.OpenRecordset("Select idart, codigo, tipo, precioventa, color, talla from alaventa where idart like " & IdArtSelec)
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
    DBGrid1.Scroll 0, DBGrid1.ApproxCount
    AgrandaColumnas
    Exit Sub
sehodio:
    MsgBox (Err.Number & Chr(13) & Err.Description)
    MsgBox ("dbgrid1_headclick")
    
End Sub

Private Sub AgrandaColumnas()
    If VerTodos = True Then
        With DBGrid1
            .Refresh
            .Columns(0).Width = 500
            .Columns(1).Width = 1600
            .Columns(2).Width = 800
            .Columns(3).Width = 700
            .Columns(4).Width = 700
'            .Columns(6).Width = 0
'            .Columns(7).Width = 0
            .Columns(5).Width = 500
            .Columns(6).Width = 1600
            .Columns(7).Width = 300
            .Columns(8).Width = 300
            .Columns(9).Width = 1500
'            .Columns(13).Width = 0
'            .Columns(14).Width = 0
'            .Columns(11).Width = 100
'            .Columns(16).Width = 0
'            .Columns(17).Width = 0
            .Refresh
        End With
    Else
        With DBGrid1
            .Columns(0).Width = 700
            .Columns(1).Width = 1900
            .Columns(2).Width = 1200
            .Columns(3).Width = 800
            .Columns(4).Width = 800
            .Columns(5).Width = 500
            .Columns(6).Width = 1600
            .Refresh
        End With
    End If
End Sub

Private Sub Form_GotFocus()
    If VerTodos = True Then AgrandaColumnas

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo sehodio
    
    Select Case KeyCode
        Case 13
            If VerTodos = False Then
                cmdpasar_Click
            End If       'Venta.Show
        Case vbKeyEscape
            Unload Me
            Venta.Show
    End Select

    Exit Sub
sehodio:
    MsgBox (Err.Number & Chr(13) & Err.Description)
    MsgBox ("form_keydown")

End Sub

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


Private Sub Form_Load()
    RsArticulo.MoveLast
    RsArticulo.MoveFirst
   ' MsgBox (RsArticulo.RecordCount)
    Permiso = False
    On Error GoTo sehodio
    If AñadeArtic = True Then
        Set RsArticulo = bdtienda.OpenRecordset("select idart, codigo, tipo, preciocompra, precioventa, talla, extra, fechacompra from alaventa")
        DBGrid1.WrapCellPointer = True
        Timer1.Enabled = True
        cmdZoom.Visible = True
    End If
    Set Data.Recordset = RsArticulo
    RsArticulo.MoveLast
    RsArticulo.MoveFirst
    'Data.DatabaseName = App.Path & "canela.mdb"
    lbregistros.Caption = "Encontrados: " & RsArticulo.RecordCount
    IdArtSelec = 0
    
    DBGrid1.Refresh
    
    If ModoBusca <> "clientes" Then
        AgrandaColumnas
    End If
    Exit Sub
    Exit Sub
sehodio:
    MsgBox (Err.Number & Chr(13) & Err.Description)
    MsgBox ("form_activate")


End Sub
Public Sub PoneTodosLosCampos()
    On Error GoTo sehodio
    Dim dummy2 As String
    dummy2 = "Select idart, codigo, tipo, preciocompra, precioventa, talla, extra, vendido, apartado, fechacompra from articulos"
    If AñadeArtic = True Then
    
        Set RsArticulo = bdtienda.OpenRecordset(dummy2) '"select idart, codigo, tipo, preciocompra, precioventa, talla, extra, fechacompra from alaventa")
        DBGrid1.WrapCellPointer = True
        DBGrid1.AllowAddNew = True
    
        DBGrid1.AllowUpdate = True
        DBGrid1.AllowDelete = True
        Timer1.Enabled = True
        cmdZoom.Visible = True
        cmdllamar.Visible = True
        cmdaddart.Visible = True
        cmdcaja.Visible = True
    End If
    Set Data.Recordset = RsArticulo
    RsArticulo.MoveLast
    RsArticulo.MoveFirst
    'Data.DatabaseName = App.Path & "canela.mdb"
    lbregistros.Caption = "Encontrados: " & RsArticulo.RecordCount
    
    IdArtSelec = 0
    
    DBGrid1.Scroll 0, DBGrid1.ApproxCount
    
        AgrandaColumnas
        
    Exit Sub
sehodio:
    MsgBox (Err.Number & Chr(13) & Err.Description)
    MsgBox ("form_activate")


End Sub

Private Sub Form_Resize()
On Error Resume Next
DBGrid1.Width = Me.Width - 300
DBGrid1.Height = Me.Height - 1500
Frame1.Top = Me.Height - 1300
Frame1.Width = Me.Width - 300
End Sub

Private Sub Form_Unload(Cancel As Integer)
    AñadeArtic = False
    VerTodos = False
    secsalir = 0
End Sub

Private Sub lbregistros_DblClick()
    On Error GoTo sehodio
    VerTodos = False
    FrmPass.Show 1

    Exit Sub
sehodio:
    MsgBox (Err.Number & Chr(13) & Err.Description)
    MsgBox ("lbregistros_dblclick")

End Sub

Private Sub Timer1_Timer()
    On Error GoTo sehodio
    
    secsalir = secsalir + 1
    If secsalir >= 1000 Then
        AñadeArtic = False
        Unload Me
    End If

    Exit Sub
sehodio:
    MsgBox (Err.Number & Chr(13) & Err.Description)
    MsgBox ("timer1_timer")

End Sub
