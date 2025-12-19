VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmResultados 
   Caption         =   "Resultados"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.MonthView calendar2 
      Height          =   2370
      Left            =   4320
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   17367042
      CurrentDate     =   40156
   End
   Begin MSComCtl2.MonthView calendar1 
      Height          =   2370
      Left            =   1440
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   17367042
      CurrentDate     =   40156
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros de Consulta"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   6360
         TabIndex        =   5
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox txtfecha2 
         Height          =   285
         Left            =   4200
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtfecha1 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Final"
         Height          =   255
         Left            =   3120
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicio"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmResultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsResultado As Recordset
Dim DumSql As String
Dim Fecha1, Fecha2 As Date



Private Sub calendar1_DateClick(ByVal DateClicked As Date)
    txtfecha1.Text = calendar1.Value
    Fecha1 = calendar1.Value
    calendar1.Visible = False
End Sub



Private Sub calendar2_DateClick(ByVal DateClicked As Date)
    txtfecha2.Text = calendar2.Value
    Fecha2 = calendar2.Value
    calendar2.Visible = False
End Sub

Private Sub cmdConsultar_Click()
    Dim TotalVenta As Currency
    DumSql = "select * from venta where fecha between #" & Fecha1 & "# and #" & Fecha2 & "#"
    'MsgBox (DumSql)
    Set RsResultado = bdtienda.OpenRecordset(DumSql)
    MsgBox DumSql
    
    'MsgBox ("Totales " & RsResultado!totales)
    Do Until RsResultado.EOF
'        If RsResultado!Fecha > Fecha1 And RsResultado!Fecha < Fecha2 Then
            TotalVenta = TotalVenta + RsResultado!total
'        End If
        RsResultado.MoveNext
    Loop
'    MsgBox (RsResultado.RecordCount)
    MsgBox (TotalVenta)
End Sub

Private Sub MonthView2_DateClick(ByVal DateClicked As Date)

End Sub

Private Sub txtfecha1_Click()
    calendar1.Value = Now
    calendar1.Visible = True
End Sub

Private Sub txtfecha2_Click()
    calendar2.Value = Now
    calendar2.Visible = True
End Sub
