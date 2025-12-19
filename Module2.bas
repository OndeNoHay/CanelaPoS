Attribute VB_Name = "Module2"
   Option Explicit

      Private Type DOCINFO
          pDocName As String
          pOutputFile As String
          pDatatype As String
      End Type

      Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal _
         hPrinter As Long) As Long
      Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal _
         hPrinter As Long) As Long
      Private Declare Function EndPagePrinter Lib "winspool.drv" (ByVal _
         hPrinter As Long) As Long
      Private Declare Function OpenPrinter Lib "winspool.drv" Alias _
         "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, _
          ByVal pDefault As Long) As Long
      Private Declare Function StartDocPrinter Lib "winspool.drv" Alias _
         "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, _
         pDocInfo As DOCINFO) As Long
      Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal _
         hPrinter As Long) As Long
      Private Declare Function WritePrinter Lib "winspool.drv" (ByVal _
         hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, _
         pcWritten As Long) As Long

Public Sub ImprimeTicket(cadena As String)
  SelectPrinter "Axiohm A793 CLASS 7193 Full"
    Dim lhPrinter As Long
    Dim lReturn As Long
    Dim lpcWritten As Long
    Dim lDoc As Long
    Dim sWrittenData As String
    Dim MyDocInfo As DOCINFO
    lReturn = OpenPrinter(Printer.DeviceName, lhPrinter, 0)
    If lReturn = 0 Then
        MsgBox "The Printer Name you typed wasn't recognized."
        Exit Sub
    End If
    MyDocInfo.pDocName = "Ticket De Venta"
    MyDocInfo.pOutputFile = vbNullString
    MyDocInfo.pDatatype = vbNullString
    lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
    Call StartPagePrinter(lhPrinter)
    ImprimeAncho "        CANELA"
    If Header.Modo = "Apartado" Then ImprimeAncho "       APARTADO"
    If Header.Modo = "PRESTAMO" Then ImprimeAncho "       PRESTAMO"
    'MandaCodigo 27, 37, 0
    cadena = SustituyeAcentos(cadena)
    
  '  MandaCodigo 27, 33, 4 ' DOBLE ANCHO
    'MandaCodigo 27, 33, 5  'DOBLE ALTO
'    sWrittenData = "CANELA"
'    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, Len(sWrittenData), lpcWritten)
    sWrittenData = cadena
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
       Len(sWrittenData), lpcWritten)
    'MandaCodigo 13
    lReturn = EndPagePrinter(lhPrinter)
    lReturn = EndDocPrinter(lhPrinter)
    lReturn = ClosePrinter(lhPrinter)
    Printer.FontName = "IDAutomationHC39M"
    Printer.FontSize = 11
    Printer.FontBold = True
    Printer.Print " *" & Header.IdVenta & "*"
    'MandaCodigo 29, 72, 4
    'MandaCodigo 29, 104, 100
    'MandaCodigo 29, 107, 4, 1
    'Printer.EndDoc
    AbreCajon
End Sub
Sub AbreCajon()
On Error GoTo sehodio
    Open "COM1" For Output Access Write As #1

    Print #1, Chr(27) & Chr(112) & Chr(0); ""
    Close #1
Exit Sub
sehodio:
    MsgBox Err.Description & " - " & Err.Number & " - " & "No se ha abierto el cajon (no se encuentra)"
End Sub
Sub MandaCodigo(Codigo As Integer, Optional dum1, Optional dum2, Optional BarC)
        Dim intFF As Integer


        intFF = FreeFile


        Open "COM1" For Output As #intFF
        'or COM2 or whatever


        'replace those chars with whatever the printer is expecting
      
        Print #intFF, Chr$(27) & Chr$(33) & Chr$(5) & "canela" ' & " " & Chr$(dum1) & " " & Chr$(dum2)
        Close #intFF
End Sub
Public Sub ImprimeAncho(cadena As String)
On Error Resume Next
        Dim intFF As Integer
        intFF = FreeFile
        Open "COM1" For Output As #intFF
        Print #intFF, Chr$(18) & cadena ' & " " & Chr$(dum1) & " " & Chr$(dum2)
        'Print #intFF, Chr$(19) '& Cadena ' & " " & Chr$(dum1) & " " & Chr$(dum2)
        
        Close #intFF
End Sub
Public Function SustituyeAcentos(cadena) As String

    cadena = Replace(cadena, "á", Chr$(160))
    cadena = Replace(cadena, "é", Chr$(130))
    cadena = Replace(cadena, "í", Chr$(161))
    cadena = Replace(cadena, "ó", Chr$(162))
    cadena = Replace(cadena, "ú", Chr$(163))
    cadena = Replace(cadena, "€", Chr$(238))
    cadena = Replace(cadena, "º", Chr$(167))
    
    SustituyeAcentos = cadena
End Function

Public Function CentraTexto(cadena As String, anchura As Integer) As String

    Dim Diferencia As Integer
    Diferencia = (anchura - Len(cadena)) / 2
    Dim i As Integer
    For i = 0 To Diferencia
        CentraTexto = CentraTexto & " "
    Next i
    CentraTexto = CentraTexto & cadena
End Function
Public Sub ImprimeTexto(cadena As String)
  SelectPrinter "Axiohm A793 CLASS 7193 Full"
    Dim lhPrinter As Long
    Dim lReturn As Long
    Dim lpcWritten As Long
    Dim lDoc As Long
    Dim sWrittenData As String
    Dim MyDocInfo As DOCINFO
    lReturn = OpenPrinter(Printer.DeviceName, lhPrinter, 0)
    If lReturn = 0 Then
        MsgBox "The Printer Name you typed wasn't recognized."
        Exit Sub
    End If
    MyDocInfo.pDocName = "Listado de Caja"
    MyDocInfo.pOutputFile = vbNullString
    MyDocInfo.pDatatype = vbNullString
    lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
    Call StartPagePrinter(lhPrinter)
    ImprimeAncho "        CANELA"
    cadena = SustituyeAcentos(cadena)
    
    sWrittenData = cadena
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
       Len(sWrittenData), lpcWritten)
    lReturn = EndPagePrinter(lhPrinter)
    lReturn = EndDocPrinter(lhPrinter)
    lReturn = ClosePrinter(lhPrinter)

End Sub


