VERSION 5.00
Begin VB.Form FrmInventario 
   BackColor       =   &H00FF8080&
   Caption         =   "Hacer Inventario"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5325
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   5325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Pone a cero inventario"
      Height          =   375
      Left            =   3000
      TabIndex        =   26
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pasa Inventario"
      Height          =   375
      Left            =   480
      TabIndex        =   25
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox txtbuscador2 
      Height          =   285
      Left            =   3840
      TabIndex        =   24
      Top             =   360
      Width           =   1335
   End
   Begin VB.CheckBox chkinventario 
      BackColor       =   &H00FF8080&
      Height          =   195
      Left            =   2280
      TabIndex        =   20
      Top             =   3840
      Width           =   255
   End
   Begin VB.TextBox txtBuscar 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.CheckBox chkapartado 
      BackColor       =   &H00FF8080&
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   3480
      Width           =   375
   End
   Begin VB.CheckBox chkvendido 
      BackColor       =   &H00FF8080&
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox txtfechacompra 
      Height          =   285
      Left            =   2280
      TabIndex        =   18
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtextra 
      Height          =   285
      Left            =   2280
      TabIndex        =   15
      Top             =   2760
      Width           =   2655
   End
   Begin VB.TextBox txtprecioventa 
      Height          =   285
      Left            =   2280
      TabIndex        =   14
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtpreciocompra 
      Height          =   285
      Left            =   2280
      TabIndex        =   13
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txttipo 
      Height          =   285
      Left            =   2280
      TabIndex        =   12
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtcodigo 
      Height          =   285
      Left            =   2280
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtidart 
      Height          =   285
      Left            =   2280
      TabIndex        =   10
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   12
      Left            =   3000
      TabIndex        =   23
      Top             =   360
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Idart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   11
      Left            =   360
      TabIndex        =   22
      Top             =   360
      Width           =   495
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5280
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inventario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   10
      Left            =   240
      TabIndex        =   21
      Top             =   3840
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   9
      Left            =   240
      TabIndex        =   19
      Top             =   0
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Compra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   8
      Left            =   240
      TabIndex        =   9
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apartado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   7
      Left            =   240
      TabIndex        =   8
      Top             =   3480
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   6
      Left            =   240
      TabIndex        =   7
      Top             =   3120
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Extra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   5
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Precio Venta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   2400
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Precio Compra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Idart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   495
   End
End
Attribute VB_Name = "FrmInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsInventario As Recordset
Dim Buscador As String
Dim Buscador2 As String

Private Sub Command1_Click()
Dim RsInvent As Recordset
Dim RsArtic As Recordset
Dim RsVendido As Recordset
Dim contador As Integer
Dim contador2 As Integer
Set RsInvent = bdtienda.OpenRecordset("inventario")
Set RsArtic = bdtienda.OpenRecordset("articulos")

    With RsArtic
        .MoveFirst
        Do Until .EOF = True
            If !fechacompra < #6/1/2005# Then
                .Edit
                RsArtic!vendido = True
                .Update
                contador = contador + 1
            End If
            .MoveNext
        Loop
    End With
    MsgBox (contador)
    contador = 0
On Error Resume Next
With RsInvent
    .MoveFirst
    Do While .EOF = False
        RsArtic.Index = "idart"
        RsArtic.Seek "=", !id
        RsArtic.Edit
        If RsArtic!fechacompra < #6/1/2005# Then
         If RsArtic!vendido = True Then
            contador = contador + 1
            RsArtic!vendido = False
         End If
        End If
        contador2 = contador2 + 1
        RsArtic.Update
'        .Edit
'        !com = "si"
'        .Update
        .MoveNext
    Loop
End With

MsgBox ("Total de prendas en inventario: " & contador2)
MsgBox ("Total de prendas marcadas como vendidas: " & contador)


End Sub

Private Sub Command2_Click()
    Dim RsAnular As Recordset
    Set RsAnular = bdtienda.OpenRecordset("select * from  articulos where inventario = true")
    With RsAnular
        .MoveFirst
        Do Until .EOF = True
            .Edit
            !inventario = False
            .Update
            .MoveNext
        Loop
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeySpace
        If chkinventario.Value = 1 Then chkinventario.Value = 0 Else chkinventario.Value = 1
        ActualizaDatos
    Case vbEnter
        ActualizaDatos
        Buscador = Trim(txtBuscar.Text)
        BuscaArticulo
        PoneDatosArticulos
    Case vbKeyLeft
        With RsInventario
            If .EOF = False Then
                .MovePrevious
                PoneDatosArticulos
            End If
        End With
    Case vbKeyRight
        With RsInventario
            If .EOF = False Then
                .MoveNext
                PoneDatosArticulos
            End If
        End With
            
    
End Select
End Sub
Private Sub ActualizaDatos()
    With RsInventario
        .Edit
'        !idart = txtidart
        !codigo = txtcodigo
        !tipo = txttipo
        !PrecioCompra = txtpreciocompra
        !PrecioVenta = txtprecioventa
        !extra = txtextra
        If chkvendido.Value = 1 Then
            !vendido = True
        Else
             !vendido = False
        End If
        If chkapartado.Value = 1 Then
            !apartado = True
        Else
            !apartado = False
        End If
        If chkinventario.Value = 1 Then
            !inventario = True
        Else
            !inventario = False
        End If
        
        !fechacompra = txtfechacompra
        .Update
    End With
End Sub
Private Sub BuscaArticulo()
    BorraControles
        If Buscador = "" Then
            Set RsInventario = bdtienda.OpenRecordset("select * from articulos where codigo like '*" & Buscador2 & "*'")
            With RsInventario
                If .EOF = False Then .MoveLast
                If .RecordCount > 1 Then
                    .MoveFirst
                    MsgBox ("Hay " & .RecordCount & " articulos." & Chr(13) & _
                    "Pulse flechas izquierda o derecha para verlos")
                End If
            End With
        Else
            Set RsInventario = bdtienda.OpenRecordset("articulos")
        
            With RsInventario
            .Index = "idart"
            .Seek "=", Buscador
            End With
        End If
    txtBuscar.SelStart = 0
    txtBuscar.SelLength = Len(txtBuscar.Text)
End Sub
Private Sub PoneDatosArticulos()
On Error GoTo sehodio
    With RsInventario
        txtidart = !idart
        txtcodigo = "" & !codigo
        txttipo = "" & !tipo
        txtpreciocompra = "" & !PrecioCompra
        txtprecioventa = "" & !PrecioVenta
        txtextra = "" & !extra
        If !vendido = True Then
            chkvendido.Value = 1
            PlayWave (App.Path & "\ringin.wav")
        Else
            chkvendido.Value = 0
        End If
        If !apartado = True Then
            chkapartado.Value = 1
            PlayWave (App.Path & "\ringin.wav")
        Else
            chkapartado.Value = 0
        End If
        If !inventario = True Then
            chkinventario.Value = 1
            PlayWave (App.Path & "\ringin.wav")
        Else
            chkinventario.Value = 0
        End If
        
        txtfechacompra = "" & !fechacompra
    End With
    Exit Sub
sehodio:
    PlayWave (App.Path & "\error.wav")
    MsgBox ("Hay algún problema. No se encuentra el artículo " & Buscador)
End Sub

Private Sub txtbuscador2_GotFocus()
    txtbuscador2.SelStart = 0
    txtbuscador2.SelLength = Len(txtbuscador2.Text)
    
End Sub

Private Sub txtbuscador2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Buscador2 = Trim(txtbuscador2.Text)
        BuscaArticulo
        PoneDatosArticulos
    End If

End Sub

Private Sub txtbuscador2_LostFocus()
    Buscador2 = ""
End Sub

Private Sub txtBuscar_GotFocus()
    txtBuscar.SelStart = 0
    txtBuscar.SelLength = Len(txtBuscar.Text)

End Sub

Private Sub txtBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Buscador = Trim(txtBuscar.Text)
        BuscaArticulo
        PoneDatosArticulos
    End If
    
End Sub
Private Sub BorraControles()
        txtidart = ""
        txtcodigo = ""
        txttipo = ""
        txtpreciocompra = ""
        txtprecioventa = ""
        txtextra = ""
            chkvendido = 0
            chkapartado.Value = 0
        txtfechacompra = ""
    
End Sub

Private Sub txtBuscar_LostFocus()
    Buscador = ""
End Sub
