VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmPrinters 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Elegir Impresora"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   2985
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel LbPrinter 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmPrinters.frx":0000
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmPrinters.frx":007A
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   2775
   End
   Begin VB.ListBox LstPrinters 
      Height          =   645
      ItemData        =   "FrmPrinters.frx":00FA
      Left            =   120
      List            =   "FrmPrinters.frx":00FC
      TabIndex        =   0
      Top             =   960
      Width           =   2775
   End
End
Attribute VB_Name = "FrmPrinters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancelar_Click()
    Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Unload Me
End Sub
Private Sub Form_Load()
    'SkinMe Me
Dim i As Integer
    LbPrinter.Caption = Printer.DeviceName
    LstPrinters.Clear
    For i = 0 To Printers.Count - 1
        LstPrinters.AddItem Printers(i).DeviceName
    Next i
End Sub

Private Sub LstPrinters_Click()
    'CmdSelect.Enabled = (cboPrinter.ListIndex >= 0)
    If SelectPrinter(LstPrinters.Text) = False Then
        MsgBox "Impresora No Encontrada", vbCritical
    Else
        MsgBox "Impresora Actual:" & Chr(13) & LstPrinters.Text, vbInformation
        Unload Me
'        cmdprint.Enabled = True
    End If
End Sub
Private Function SelectPrinter(ByVal printer_name As String) As Boolean
Dim i As Integer
    'SelectPrinter = True
    For i = 0 To Printers.Count - 1
        If Printers(i).DeviceName = printer_name Then
            Set Printer = Printers(i)
            SelectPrinter = True
            Exit For
        End If
    Next i
End Function

