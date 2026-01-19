VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form FrmEtiquetas 
   Caption         =   "Form1"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13140
   LinkTopic       =   "Form1"
   ScaleHeight     =   139.965
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   231.775
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox DBGrid1PB 
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5475
      ScaleWidth      =   13035
      TabIndex        =   26
      Top             =   1440
      Width           =   13095
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmetiquetasPS.frx":0000
         Height          =   5535
         Left            =   -120
         OleObjectBlob   =   "frmetiquetasPS.frx":0013
         TabIndex        =   27
         Top             =   0
         Width           =   13215
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   255
      Left            =   7320
      TabIndex        =   23
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5160
      TabIndex        =   22
      Text            =   "12345678901"
      Top             =   7200
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   3000
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   21
      Top             =   7080
      Width           =   855
   End
   Begin VB.Data Data 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "rsarticulo"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Intervalo Impresión"
      Height          =   1335
      Left            =   9600
      TabIndex        =   13
      Top             =   0
      Width           =   3495
      Begin VB.TextBox Txtultimo 
         Height          =   285
         Left            =   2280
         TabIndex        =   18
         Top             =   330
         Width           =   735
      End
      Begin VB.TextBox Txtprimero 
         Height          =   285
         Left            =   840
         TabIndex        =   17
         Top             =   330
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Imprime con logo"
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lbnumero 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   720
         Width           =   105
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   16
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tamaño Etiqueta"
      Height          =   1335
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   1935
      Begin VB.ComboBox Cmbalto 
         Height          =   315
         ItemData        =   "frmetiquetasPS.frx":09E6
         Left            =   840
         List            =   "frmetiquetasPS.frx":09E8
         TabIndex        =   11
         Text            =   "35"
         Top             =   840
         Width           =   735
      End
      Begin VB.ComboBox Cmbancho 
         Height          =   315
         Left            =   840
         TabIndex        =   10
         Text            =   "70"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Alto"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Ancho"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Condiciones Impresión"
      Height          =   1335
      Left            =   1920
      TabIndex        =   1
      Top             =   0
      Width           =   7695
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   4800
         TabIndex        =   24
         Text            =   "15"
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox chknum 
         Caption         =   "Número de Etiquetas:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox TxtNumEtiq 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   12
         Text            =   "27"
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Empezar a imprimir en:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox Cmbfila 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4560
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox cmbcolumna 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3120
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Margen Superior"
         Height          =   255
         Left            =   3240
         TabIndex        =   25
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Fila"
         Height          =   255
         Left            =   4200
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Columna"
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "imprime con código"
      Height          =   615
      Left            =   10800
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   8880
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   120
      Picture         =   "frmetiquetasPS.frx":09EA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   570
   End
End
Attribute VB_Name = "FrmEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Numetiqhor As Integer
Dim Numetiqver As Integer
Dim AltoEtiq, AnchoEtiq As Integer
Dim RsArtImpr As Recordset
Dim PasaPrimerNum As Boolean
Dim MargenSuperior As Integer

Private Sub Check1_Click()
If Check1.Value = 1 Then
    cmbcolumna.Enabled = True
    Cmbfila.Enabled = True
Else
    cmbcolumna.Enabled = False
    Cmbfila.Enabled = False
End If
End Sub

Private Sub chknum_Click()
If chknum.Value = 1 Then
    TxtNumEtiq.Enabled = True
Else
    TxtNumEtiq.Enabled = False
End If

End Sub

Private Sub Command1_Click()
ImprimeEtiquetas
End Sub
Private Sub ImprimeEtiquetas()
Dim Contahoriz, Contaverti As Integer
Dim NumImpresa As Integer
Dim PrimerArt, UltimoArt As Long
Dim NumParaImpr As Integer

On Error GoTo sehodio

PrimerArt = Val(TxtPrimero)
UltimoArt = Val(TxtUltimo)
'NumParaImpr = UltimoArt - PrimerArt

With RsArtImpr
    .MoveFirst
    Do Until !Idart = PrimerArt
        .MoveNext
    Loop
    

AnchoEtiq = Cmbancho
AltoEtiq = Cmbalto

Numetiqver = Int(297 / AltoEtiq)
Numetiqhor = Int(210 / AnchoEtiq)
Contahoriz = 0
Contaverti = 0

Dim x, Y As Integer
x = 2
Y = MargenSuperior

If Check1.Value = 1 Then
    Contahoriz = Val(cmbcolumna.Text) - 1
    Contaverti = Val(Cmbfila.Text) - 1
    For i = 1 To Contahoriz
        x = x + AnchoEtiq
    Next i
    For i = 1 To Contaverti
        Y = Y + AltoEtiq
    Next i
End If
'MsgBox (Printer.FontSize)
'Printer.PrintQuality = -1
Printer.ScaleMode = 6
Do Until Contaverti = Numetiqver
    Do Until Contahoriz = Numetiqhor
        'Printer.CurrentX = x
        'Printer.CurrentY = y
        Printer.FontSize = 10
        Printer.FontName = "IDAutomationHC39M"
        Printer.CurrentX = x + 22
        Printer.CurrentY = Y
        Dim Preciosindecimal As String
        Preciosindecimal = !PrecioCompra
        If InStr(1, Preciosindecimal, ",") > 0 Or InStr(1, Preciosindecimal, ".") > 0 Then
            Preciosindecimal = Replace(Preciosindecimal, ",", "")
        End If
        
        Printer.Print "*" & !Idart & Preciosindecimal & (Int(Rnd * 89) + 10) & "*" 'Contahoriz & Contaverti
        
        Printer.FontName = "Arial"
        Printer.FontSize = 8
        Printer.CurrentX = x + 42
        Printer.CurrentY = Y
        Printer.Print "Code: " & Left(!codigo, 3)
         
        Printer.FontName = "Arial"
        Printer.PaintPicture Image1.Picture, x, Y, 10, 10
        
        'Printer.CurrentX = x + 12
        'Printer.CurrentY = y + 5
        If !deposito = True Then
            Printer.FillColor = vbRed
            Printer.Circle (x + 14, Y + 5), 2, vbRed
        End If
        
        Printer.CurrentX = x
        Printer.CurrentY = Y + 12
        Printer.Print !Tipo & "  " & !Color & "  " & !Talla & "  " & !extra '"Etiqueta vertical nº" & Contaverti
        
        Printer.CurrentX = x
        Printer.CurrentY = Y + 15
        Printer.FontSize = 12
        Printer.Print "PVP: " & Format(!PrecioVenta, "#00.00#") & "€" '"Etiqueta horizontal nº " & Contahoriz
        
        Printer.CurrentX = x + 30
        Printer.CurrentY = Y + 15
        Printer.FontSize = 8
        Printer.Print Right(!codigo, Len(!codigo) - 4)
        
        If !preciorebaja > 0 Then
            Printer.CurrentX = x
            Printer.CurrentY = Y + 20
            Printer.FontSize = 12
            Printer.FontBold = True
            Printer.ForeColor = vbRed
            Printer.Print "Rebajado: " & Format(!preciorebaja, "#00.00#") & €
            Printer.ForeColor = vbBlack
            Printer.FontBold = False
        End If
        
        
        
        Contahoriz = Contahoriz + 1
        x = x + AnchoEtiq '(210 / Numetiqhor)
        NumImpresa = NumImpresa + 1
        If UltimoArt = !Idart Then
            Printer.EndDoc
            Exit Sub
        End If
        .MoveNext
    Loop
    
    Contahoriz = 0
    Contaverti = Contaverti + 1
    
    If Contaverti = Numetiqver Then 'And NumImpresa < Val(TxtNumEtiq) Then
        Printer.NewPage
        Contaverti = 0
        Y = MargenSuperior
        x = 5
    Else
'    If y >= 260 Then
'        Printer.NewPage
'        y = 10
'    End If
        Y = Y + AltoEtiq '(290 / Numetiqver)
        x = 5
    End If
Loop
End With
Printer.EndDoc
Exit Sub
sehodio:
    MsgBox (Err.Number & Chr(13) & Err.Description)
End Sub
Private Sub ImprimeCodigo()
Dim Contahoriz, Contaverti As Integer
Contahoriz = 0
Contaverti = 0
Dim x, Y As Integer
x = 5
Y = 10

Printer.ScaleMode = 6
Do Until Contaverti = Numetiqver
    Do Until Contahoriz = Numetiqhor
        Printer.FontName = "IDAutomationHC39M"
        Printer.CurrentX = x
        Printer.CurrentY = Y
        Printer.Print "*1234567890*" & Contahoriz & Contaverti
        Printer.FontName = "Arial"
        Printer.CurrentX = x
        Printer.CurrentY = Y + 10
        Printer.Print "Etiqueta horizontal nº " & Contahoriz
        Printer.CurrentX = x
        Printer.CurrentY = Y + 14
        Printer.Print "Etiqueta vertical nº" & Contaverti
        Contahoriz = Contahoriz + 1
        x = x + (210 / Numetiqhor)
    Loop
    Contahoriz = 0
    Contaverti = Contaverti + 1
    If Y >= 260 Then
        Printer.NewPage
        Y = 10
    End If
    Y = Y + (290 / Numetiqver)
    x = 5
Loop
Printer.EndDoc

End Sub

Private Sub Command2_Click()
Numetiqhor = 2
Numetiqver = 14
ImprimeCodigo
End Sub

Private Sub Command3_Click()
        Printer.PaintPicture Picture1.Picture, 5, 5, 10, 10

End Sub

Private Sub DBGrid1_Click()

    'PasaPrimerNum = Not PasaPrimerNum
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If PasaPrimerNum = True Then
    TxtPrimero = DBGrid1.Text
    PasaPrimerNum = False
Else
    TxtUltimo = DBGrid1.Text
    PasaPrimerNum = True
End If

End Sub

Private Sub Form_Load()
AnchoEtiq = Cmbancho
AltoEtiq = Cmbalto
Numetiqhor = 3
Numetiqver = 12
'MargenSuperior = 3
MargenSuperior = Int(Text2.Text)

For i = 1 To Numetiqhor
    cmbcolumna.AddItem i
Next i
For i = 1 To Numetiqver
    Cmbfila.AddItem i
Next i
Set RsArtImpr = bdtienda.OpenRecordset("Select idart, codigo, tipo, precioventa, preciorebaja, color, talla, extra, fechacompra, preciocompra, deposito from articulos where vendido = false order by idart")
Set Data.Recordset = RsArtImpr
DBGrid1.Columns(8).Visible = False
End Sub

Private Sub Text1_Change()
Call DrawBarcode(Text1, Picture1)
End Sub

Private Sub Text2_Change()
    On Error GoTo sehodio
    MargenSuperior = Int(Text2.Text)
    Exit Sub
sehodio:
    MsgBox "El margen no se ha podido fijar. Actualmente está en 3px"
    MargenSuperior = 3
End Sub

Private Sub Txtultimo_Change()
On Error Resume Next
    Dim canti As Integer
    canti = Val(TxtUltimo.Text) - Val(TxtPrimero.Text) + 1
    If canti <= 0 Then lbnumero = "": Exit Sub
    lbnumero = "Nº etiquetas: " & canti
End Sub
