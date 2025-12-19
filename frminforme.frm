VERSION 5.00
Begin VB.Form frminforme 
   Caption         =   "Informe"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8205
   LinkTopic       =   "Form2"
   ScaleHeight     =   8490
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   7920
      Width           =   8175
      Begin VB.Label lbtotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3120
         TabIndex        =   10
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Total de prendas:"
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Establecer Condiciones de Búsqueda"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7935
      Begin VB.CommandButton cmdverdatos 
         Caption         =   "Ver Datos"
         Height          =   375
         Left            =   5280
         TabIndex        =   12
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3600
         TabIndex        =   9
         Top             =   210
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1920
         TabIndex        =   8
         Top             =   210
         Width           =   1335
      End
      Begin VB.CommandButton cmdprint 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   6840
         TabIndex        =   5
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "y"
         Height          =   255
         Left            =   3360
         TabIndex        =   7
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha de compra entre"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Label lbcantidad 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lbtipo 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo de Prenda"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frminforme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstipos As Recordset
Dim fechainicio, fechafin As Date
Dim tipos() As String
Dim cantidad() As Integer



Private Sub cmdprint_Click()
Dim x, y As Integer

Printer.ScaleMode = 6
x = 10
y = 10
Printer.Print "Informe de prendas en Canela a fecha " & Date
For i = 0 To lbtipo.Count - 1
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print lbtipo(i).Caption
    Printer.CurrentX = x + 30
    Printer.CurrentY = y
    Printer.Print lbcantidad(i).Caption
    y = y + 5
Next i
Printer.EndDoc
End Sub

Private Sub cmdverdatos_Click()
    On Error Resume Next
    BorraDatos
    PoneTipos
    PoneCantidad

End Sub
Private Sub BorraDatos()
    On Error Resume Next
    For i = 1 To lbtipo.Count
        Unload lbtipo(i)
        Unload lbcantidad(i)
    Next i
End Sub

Private Sub PoneTipos()
Set rstipos = bdtienda.OpenRecordset("select * from tipo order by orden")
Dim num As Integer
num = lbtipo.Count
With rstipos
Do Until .EOF
    Load lbtipo(num)
    Load lbcantidad(num)
    lbtipo(num) = !tipo
    lbtipo(num).Top = lbtipo(num - 1).Top + lbtipo(num - 1).Height
    lbtipo(num).Visible = True
    lbcantidad(num).Top = lbcantidad(num - 1).Top + lbcantidad(num - 1).Height
    lbcantidad(num).Visible = True
    ReDim Preserve tipos(num)
    tipos(num) = !tipo
    
    num = lbtipo.Count
    .MoveNext
Loop
ReDim cantidad(num)
End With
End Sub
Private Sub PoneCantidad()
    Dim rscantart As Recordset
    Dim contador As Integer
    Dim dummyrs As String
    Dim total As Integer
    contador = 1
    rstipos.MoveFirst
    Do Until rstipos.EOF
        dummyrs = "Select * from articulos where tipo = '" & rstipos!tipo & "' and" _
        & " vendido = false and fechacompra between #" & fechainicio & "# and #" & fechafin & "#"
        Set rscantart = bdtienda.OpenRecordset(dummyrs)
'        If rstipos!tipo = "conjunto" Then
'            Do Until rscantart.EOF
'                MsgBox (rscantart!fechacompra)
'                rscantart.MoveNext
'            Loop
'        End If
        If rscantart.EOF = False Then rscantart.MoveLast
        lbcantidad(contador) = rscantart.RecordCount
        cantidad(contador) = rscantart.RecordCount
        total = total + cantidad(contador)
        contador = contador + 1
        rstipos.MoveNext
    Loop
    lbtotal = total
End Sub

