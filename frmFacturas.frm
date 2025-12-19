VERSION 5.00
Begin VB.Form frmFacturas 
   Caption         =   "Facturas"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3510
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   3510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdverfacturas 
      Caption         =   "&Ver Facturas"
      Height          =   375
      Left            =   1080
      TabIndex        =   16
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Bus&car"
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdNueva 
      Caption         =   "&Nueva"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtPagado 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Text            =   "0"
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox txtTotal 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   "0"
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtFecha 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.ComboBox cmbProveedor 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtFactura 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin VB.OptionButton Optfact 
      Alignment       =   1  'Right Justify
      Caption         =   "Factura &B"
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.OptionButton Optfact 
      Alignment       =   1  'Right Justify
      Caption         =   "Factura &A"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.Label lbidfactura 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "IdFactura"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   14
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Pagado"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   13
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Total"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Proveedor"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Factura"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "frmFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsFacturas As Recordset
Dim Rsprovee As Recordset
Dim Actualizando As Boolean

Private Sub cmbProveedor_Click()
    cmbProveedor.Text = cmbProveedor.Text
End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo sehodio
    With RsFacturas
        .MoveFirst
        .Index = "identif"
        .Seek "=", txtFactura.Text
        'If .NoMatch = False Then Exit Sub
        
        lbidfactura = !idfactura
        txtFactura = !identif
        cmbProveedor.Text = !proveedor
        txtFecha = !Fecha
        If !legal = True Then Optfact(0).Value = 1 Else Optfact(0).Value = 0
        txtTotal = !total
        txtPagado = !pagado
        Actualizando = True
        Selecciona
    End With
    Exit Sub
sehodio:
MsgBox ("No se encuentran los datos")
Selecciona
End Sub
Private Sub Selecciona()
    txtFactura.SetFocus
    txtFactura.SelStart = 0
    txtFactura.SelLength = Len(txtFactura.Text)

End Sub
Private Sub Actualizar()
    With RsFacturas
        .Edit
        !identif = "" & txtFactura
        !proveedor = cmbProveedor.Text
        !Fecha = "" & txtFecha
        If Optfact(0).Value = True Then
            !legal = True
        Else
            !legal = False
        End If
        !total = 0 + txtTotal
        !pagado = 0 + txtPagado
        .Update
    End With
    Selecciona
End Sub
Private Sub cmdNueva_Click()
    If Actualizando = True Then
        Actualizar
        Actualizando = False
        BorraDatos
    Else
        AñadeFactura
    End If
    Selecciona
End Sub

Private Sub cmdverfacturas_Click()
    VerTabla = "Facturas"
    FrmVerTablas.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            SendKeys (vbTab)
    End Select
End Sub

Private Sub Form_Load()
    Set RsFacturas = bdtienda.OpenRecordset("facturas")
    Set Rsprovee = bdtienda.OpenRecordset("proveedor")
    With Rsprovee
        .MoveFirst
        Do Until .EOF = True
            cmbProveedor.AddItem !codigo
            .MoveNext
        Loop
    End With
    
    
End Sub
Private Sub AñadeFactura()
    If txtFactura = "" Then Exit Sub
    With RsFacturas
        .AddNew
        !identif = "" & txtFactura
        !proveedor = cmbProveedor.Text
        !Fecha = "" & txtFecha
        If Optfact(0).Value = True Then
            !legal = True
        Else
            !legal = False
        End If
        If InStr(1, txtTotal.Text, ".") Then txtTotal.Text = Replace(txtTotal.Text, ".", ",")
        !total = 0 + txtTotal
        If InStr(1, txtPagado.Text, ".") Then txtPagado.Text = Replace(txtPagado.Text, ".", ",")
        !pagado = 0 + txtPagado
        MsgBox ("IdFactura = " & !idfactura)
        .Update
    End With
    BorraDatos
    
End Sub
Private Sub BorraDatos()
    lbidfactura = ""
    txtFactura = ""
    cmbProveedor = ""
    txtFecha = ""
    Optfact(0).Value = True
    txtTotal = 0
    txtPagado = 0
    Selecciona
End Sub

Private Sub txtFactura_GotFocus()
    Selecciona
End Sub

Private Sub txtFecha_GotFocus()
    txtFecha.SelStart = 0
    txtFecha.SelLength = Len(txtFecha.Text)
End Sub

Private Sub txtPagado_GotFocus()
    txtPagado.SelStart = 0
    txtPagado.SelLength = Len(txtPagado.Text)
End Sub

Private Sub txtTotal_Change()
   ' txtPagado = txtTotal
End Sub

Private Sub txtTotal_GotFocus()
    txtTotal.SelStart = 0
    txtTotal.SelLength = Len(txtTotal.Text)
End Sub
