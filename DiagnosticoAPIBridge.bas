' ============================================================================
' SCRIPT DE DIAGNOSTICO - API BRIDGE PRESTASHOP
' Ejecutar desde VB6 para probar diferentes formatos de petición
' ============================================================================

Sub DiagnosticarAPIBridge()
    On Error Resume Next

    Dim xmlHttp As Object
    Dim url As String
    Dim codigo As String
    Dim response As String

    codigo = "2804389083757"  ' Código de prueba

    MsgBox "Iniciando diagnóstico del API Bridge..." & vbCrLf & _
           "Código a probar: " & codigo, vbInformation

    ' ========== PRUEBA 1: GET con action=search (actual) ==========
    Debug.Print "========================================"
    Debug.Print "PRUEBA 1: GET con action=search"
    Debug.Print "========================================"

    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    url = "https://www.canelamoda.es/api_bridge/bridge.php?action=search&code=" & codigo

    Debug.Print "URL: " & url

    xmlHttp.Open "GET", url, False
    xmlHttp.setRequestHeader "Accept", "application/json"
    xmlHttp.Send

    Debug.Print "Status: " & xmlHttp.Status
    Debug.Print "Response: " & Left(xmlHttp.responseText, 500)
    Debug.Print ""

    Set xmlHttp = Nothing


    ' ========== PRUEBA 2: GET con reference en lugar de code ==========
    Debug.Print "========================================"
    Debug.Print "PRUEBA 2: GET con reference= en lugar de code="
    Debug.Print "========================================"

    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    url = "https://www.canelamoda.es/api_bridge/bridge.php?action=search&reference=" & codigo

    Debug.Print "URL: " & url

    xmlHttp.Open "GET", url, False
    xmlHttp.setRequestHeader "Accept", "application/json"
    xmlHttp.Send

    Debug.Print "Status: " & xmlHttp.Status
    Debug.Print "Response: " & Left(xmlHttp.responseText, 500)
    Debug.Print ""

    Set xmlHttp = Nothing


    ' ========== PRUEBA 3: GET con ean13 en lugar de code ==========
    Debug.Print "========================================"
    Debug.Print "PRUEBA 3: GET con ean13= en lugar de code="
    Debug.Print "========================================"

    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    url = "https://www.canelamoda.es/api_bridge/bridge.php?action=search&ean13=" & codigo

    Debug.Print "URL: " & url

    xmlHttp.Open "GET", url, False
    xmlHttp.setRequestHeader "Accept", "application/json"
    xmlHttp.Send

    Debug.Print "Status: " & xmlHttp.Status
    Debug.Print "Response: " & Left(xmlHttp.responseText, 500)
    Debug.Print ""

    Set xmlHttp = Nothing


    ' ========== PRUEBA 4: POST con JSON ==========
    Debug.Print "========================================"
    Debug.Print "PRUEBA 4: POST con JSON"
    Debug.Print "========================================"

    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    url = "https://www.canelamoda.es/api_bridge/bridge.php"

    Dim jsonData As String
    jsonData = "{""action"":""search"",""code"":""" & codigo & """}"

    Debug.Print "URL: " & url
    Debug.Print "JSON: " & jsonData

    xmlHttp.Open "POST", url, False
    xmlHttp.setRequestHeader "Content-Type", "application/json"
    xmlHttp.setRequestHeader "Accept", "application/json"
    xmlHttp.Send jsonData

    Debug.Print "Status: " & xmlHttp.Status
    Debug.Print "Response: " & Left(xmlHttp.responseText, 500)
    Debug.Print ""

    Set xmlHttp = Nothing


    ' ========== PRUEBA 5: GET sin action (solo code) ==========
    Debug.Print "========================================"
    Debug.Print "PRUEBA 5: GET sin action, solo code="
    Debug.Print "========================================"

    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    url = "https://www.canelamoda.es/api_bridge/bridge.php?code=" & codigo

    Debug.Print "URL: " & url

    xmlHttp.Open "GET", url, False
    xmlHttp.setRequestHeader "Accept", "application/json"
    xmlHttp.Send

    Debug.Print "Status: " & xmlHttp.Status
    Debug.Print "Response: " & Left(xmlHttp.responseText, 500)
    Debug.Print ""

    Set xmlHttp = Nothing


    ' ========== PRUEBA 6: GET con método diferente (getProduct) ==========
    Debug.Print "========================================"
    Debug.Print "PRUEBA 6: GET con action=getProduct"
    Debug.Print "========================================"

    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    url = "https://www.canelamoda.es/api_bridge/bridge.php?action=getProduct&code=" & codigo

    Debug.Print "URL: " & url

    xmlHttp.Open "GET", url, False
    xmlHttp.setRequestHeader "Accept", "application/json"
    xmlHttp.Send

    Debug.Print "Status: " & xmlHttp.Status
    Debug.Print "Response: " & Left(xmlHttp.responseText, 500)
    Debug.Print ""

    Set xmlHttp = Nothing


    ' ========== PRUEBA 7: Probar test_bridge.html endpoint ==========
    Debug.Print "========================================"
    Debug.Print "PRUEBA 7: Verificar test_bridge.html"
    Debug.Print "========================================"

    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    url = "https://www.canelamoda.es/api_bridge/test_bridge.html"

    Debug.Print "URL: " & url

    xmlHttp.Open "GET", url, False
    xmlHttp.Send

    Debug.Print "Status: " & xmlHttp.Status
    Debug.Print "Response (primeros 200 chars): " & Left(xmlHttp.responseText, 200)
    Debug.Print ""

    Set xmlHttp = Nothing


    MsgBox "Diagnóstico completado. Revisa la ventana Immediate (Ctrl+G) para ver resultados.", vbInformation

End Sub
