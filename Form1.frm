VERSION 5.00
Begin VB.Form frmtotales 
   Caption         =   "Form1"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4755
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdbalance 
      Caption         =   "Balance"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "ingresos"
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "importe prendas por vender"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lbarticulos 
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Artículos"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lbresultado 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lbgasto 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lbingreso 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
End
Attribute VB_Name = "frmtotales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BuscaRegistro()
With RsArticulo
    .MoveFirst
    Do Until .EOF
        If !idart = Text1.Text Then
            .Edit
            !vendido = False
            .Update
        End If
        Label1 = !idart
        DoEvents
        .MoveNext
    Loop
    Text1.Text = ""
    Text1.SetFocus
End With
End Sub

Private Sub cmdbalance_Click()
 Set RsArticulo = bdtienda.OpenRecordset("Select * from articulos where vendido = false")
 On Error Resume Next
 With RsArticulo
    .MoveLast
    lbarticulos = RsArticulo.RecordCount
    .MoveFirst
    Do Until .EOF
    lbgasto = Val(lbgasto) + !preciocompra
    lbingreso = Val(lbingreso) + !precioventa
    lbresultado = Val(lbingreso) - Val(lbgasto)
    DoEvents
    
    .MoveNext
    Loop
End With
 
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        BuscaRegistro
End Select

End Sub

Private Sub Form_Load()
Set RsArticulo = bdtienda.OpenRecordset("articulos")
    MsgBox (RsArticulo.RecordCount)

End Sub

