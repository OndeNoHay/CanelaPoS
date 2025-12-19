VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{E5CEE37F-8CF8-489E-BFA0-8201CBD6AEE8}#1.0#0"; "PicFormat32.ocx"
Begin VB.Form FrmWebCam 
   BackColor       =   &H00FF8080&
   Caption         =   "Image Capture"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   615
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   625
   StartUpPosition =   3  'Windows Default
   Begin PicFormat32a.PicFormat32 PicFormat321 
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   7920
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin VB.CommandButton cmdBisuteria 
      Caption         =   "Bisuteria"
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdPlata 
      Caption         =   "Plata"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdropa 
      Caption         =   "Ropa"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   8040
      Width           =   1335
   End
   Begin VB.ListBox lstDevices 
      Height          =   450
      Left            =   1320
      TabIndex        =   5
      Top             =   0
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Guardar Foto (F2)"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   8760
      Width           =   1695
   End
   Begin VB.PictureBox picCapture 
      Height          =   7200
      Left            =   0
      ScaleHeight     =   7140
      ScaleWidth      =   9240
      TabIndex        =   3
      Top             =   480
      Width           =   9300
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   8760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Iniciar Cámara"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   8760
      Width           =   1815
   End
   Begin VB.Line Line2 
      X1              =   96
      X2              =   448
      Y1              =   576
      Y2              =   576
   End
   Begin VB.Line Line1 
      X1              =   104
      X2              =   456
      Y1              =   520
      Y2              =   520
   End
   Begin VB.Label Label1 
      Caption         =   "Available devices"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "FrmWebCam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'For further information on this Feel Free
'to contact me at brtiwari@yahoo.com
'and donot forget to vote me on PSC
Const WM_CAP As Integer = &H400

Const WM_CAP_DRIVER_CONNECT As Long = WM_CAP + 10
Const WM_CAP_DRIVER_DISCONNECT As Long = WM_CAP + 11
Const WM_CAP_EDIT_COPY As Long = WM_CAP + 30

Const WM_CAP_SET_PREVIEW As Long = WM_CAP + 50
Const WM_CAP_SET_PREVIEWRATE As Long = WM_CAP + 52
Const WM_CAP_SET_SCALE As Long = WM_CAP + 53
Const WS_CHILD As Long = &H40000000
Const WS_VISIBLE As Long = &H10000000
Const SWP_NOMOVE As Long = &H2
Const SWP_NOSIZE As Integer = 1
Const SWP_NOZORDER As Integer = &H4
Const HWND_BOTTOM As Integer = 1

Dim iDevice As Long  ' Current device ID
Dim hHwnd As Long ' Handle to preview window

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hndw As Long) As Boolean
Private Declare Function capCreateCaptureWindowA Lib "avicap32.dll" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Integer, ByVal hWndParent As Long, ByVal nID As Long) As Long
Private Declare Function capGetDriverDescriptionA Lib "avicap32.dll" (ByVal wDriver As Long, ByVal lpszName As String, ByVal cbName As Long, ByVal lpszVer As String, ByVal cbVer As Long) As Boolean

Dim Carpeta As String
Dim Tipo As String
Dim contador As Integer

Private Sub cmdBisuteria_Click()
    Carpeta = "Bisuteria\"
End Sub

Private Sub cmdPlata_Click()
    Carpeta = "Plata\"
End Sub

Private Sub cmdropa_Click()
    Carpeta = "Ropa\"
End Sub

Private Sub cmdSave_Click()
    Dim bm As Image
    
    ' Copy image to clipboard
    SendMessage hHwnd, WM_CAP_EDIT_COPY, 0, 0
    ClosePreviewWindow

    picCapture.Picture = Clipboard.GetData
    
'    CommonDialog1.CancelError = True
'    CommonDialog1.FileName = "Webcam1"
'    CommonDialog1.Filter = "Bitmap |*.bmp|JPEG |*.jpeg"
'
'    On Error GoTo NoSave
'    CommonDialog1.ShowSave
'    SavePicture picCapture.Image, App.Path & "\fotos\" & Carpeta & IdCliFoto & "-" & Hour(Time) & "-" & Minute(Time) & "-" & Second(Time) & ".bmp"
    Dim TempImg As String
    
    TempImg = App.Path & "\fotos\tempimg.bmp"
    
    On Error Resume Next
    
    Kill TempImg
    
    SavePicture picCapture.Image, TempImg
    
    DirFoto = "\fotos\" & Carpeta & IdCliFoto & "-" & Hour(Time) & "-" & Minute(Time) & "-" & Second(Time) & ".jpg"
    
    PicFormat321.SaveBmpToJpeg TempImg, App.Path & DirFoto, 65
    
    If FotoAddArt = True Then Unload Me
    
NoSave:
'    cmdStop.Enabled = False
'    'cmdSave.Enabled = False
'    OpenPreviewWindow
'    cmdStart.Enabled = True
End Sub

Private Sub cmdStart_Click()
    iDevice = lstDevices.ListIndex
    OpenPreviewWindow
End Sub

Private Sub cmdStop_Click()
    ClosePreviewWindow
    cmdStop.Enabled = False
    cmdSave.Enabled = False
    cmdStart.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbEnter
            cmdSave_Click
        Case vbKeyF2
            cmdSave_Click
    End Select
End Sub

Private Sub Form_Load()
    LoadDeviceList
    
    If lstDevices.ListCount > 0 Then
        lstDevices.Selected(0) = True
    Else
    cmdStart.Enabled = False
        lstDevices.AddItem ("No Device Available")
    End If
    
    cmdStop.Enabled = False
    cmdSave.Enabled = False
    cmdStart_Click
End Sub

Private Sub LoadDeviceList()
    Dim strName As String
    Dim strVer As String
    Dim iReturn As Boolean
    Dim x As Long
    
    x = 0
    strName = Space(100)
    strVer = Space(100)

    ' Load name of all available devices into lstDevices
    Do
        ' Get Driver name and version
        iReturn = capGetDriverDescriptionA(x, strName, 100, strVer, 100)
        ' If there was a device add device name to the list
        If iReturn Then lstDevices.AddItem Trim$(strName)
        x = x + 1
    Loop Until iReturn = False
End Sub

Private Sub OpenPreviewWindow()

    ' Open Preview window in picturebox
    hHwnd = capCreateCaptureWindowA(iDevice, WS_VISIBLE Or WS_CHILD, 0, 0, 640, 480, picCapture.hwnd, 0)

    ' Connect to device
    If SendMessage(hHwnd, WM_CAP_DRIVER_CONNECT, iDevice, 0) Then

        'Set the preview scale
        SendMessage hHwnd, WM_CAP_SET_SCALE, True, 0

        'Set the preview rate in milliseconds
        SendMessage hHwnd, WM_CAP_SET_PREVIEWRATE, 66, 0

        'Start previewing the image from the camera
        SendMessage hHwnd, WM_CAP_SET_PREVIEW, True, 0

        ' Resize window to fit in picturebox
        'SetWindowPos hHwnd, HWND_BOTTOM, 0, 0, picCapture.ScaleWidth, picCapture.ScaleHeight, SWP_NOMOVE Or SWP_NOZORDER

        cmdSave.Enabled = True
        cmdStop.Enabled = True
        cmdStart.Enabled = False
    Else

        ' Error connecting to device close window
        DestroyWindow hHwnd

        cmdSave.Enabled = False
    End If
 End Sub

Private Sub ClosePreviewWindow()
    ' Disconnect from device
    SendMessage hHwnd, WM_CAP_DRIVER_DISCONNECT, iDevice, 0

    ' close window
    DestroyWindow hHwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdStop.Enabled Then
        ClosePreviewWindow
    End If
End Sub

