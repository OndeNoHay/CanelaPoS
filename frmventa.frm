VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Venta 
   BackColor       =   &H00FF8080&
   Caption         =   "Venta"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12300
   Icon            =   "frmventa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   8940
   ScaleWidth      =   12300
   WindowState     =   2  'Maximized
   Begin VB.ComboBox ComboTallas 
      Height          =   315
      Index           =   1
      Left            =   5400
      TabIndex        =   73
      Text            =   "Combo1"
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton CmdTicketRegalo 
      Caption         =   "Imprime el Ticket Regalo"
      Height          =   615
      Left            =   8400
      TabIndex        =   71
      Top             =   8160
      Width           =   1455
   End
   Begin VB.TextBox TxtBusca 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   120
      TabIndex        =   68
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton CmdImprimeTicket 
      Caption         =   "Imprime el �ltimoTicket"
      Height          =   615
      Left            =   10080
      TabIndex        =   67
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   5280
      Top             =   3960
   End
   Begin VB.CommandButton cmdElimina 
      Caption         =   "Elimina"
      Height          =   375
      Index           =   1
      Left            =   9360
      TabIndex        =   65
      Top             =   2040
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   360
      Top             =   4800
   End
   Begin VB.CommandButton cmdgenerico 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Art�culo &Gen�rico"
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   3480
      Width           =   975
   End
   Begin VB.Frame frameCobrar 
      BackColor       =   &H00FF8080&
      Caption         =   "Cobrar"
      ForeColor       =   &H8000000E&
      Height          =   2055
      Left            =   7200
      TabIndex        =   53
      Top             =   6000
      Width           =   4575
      Begin VB.CommandButton cmdcobrar 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cobrar"
         Height          =   1455
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txttotal 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   600
         Width           =   2415
      End
      Begin VB.OptionButton OptTodo 
         BackColor       =   &H00FF8080&
         Caption         =   "Parte"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   57
         Top             =   1440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox TxtParte 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         TabIndex        =   56
         Text            =   "0"
         Top             =   1680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton OptTodo 
         BackColor       =   &H00FF8080&
         Caption         =   "Todo"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   55
         Top             =   1440
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkPago 
         BackColor       =   &H00FF8080&
         Caption         =   "Pago con Tarjeta"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Total a Pagar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lbparte 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   690
         TabIndex        =   59
         Top             =   1680
         Visible         =   0   'False
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmddevuelven 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Han devuelto una prenda"
      Height          =   855
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdllamar 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Llamar"
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   160
      Width           =   855
   End
   Begin VB.PictureBox Msc 
      Height          =   480
      Left            =   9600
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   72
      Top             =   1560
      Width           =   1200
   End
   Begin VB.CommandButton cmdventasxcliente 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ventas X Cliente"
      Height          =   255
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdvolver 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Volver"
      Height          =   735
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmddevuelto 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Devolver  Art�culos"
      Height          =   735
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   4320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Salir"
      Height          =   615
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "Apartado"
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   3720
      TabIndex        =   41
      Top             =   6000
      Width           =   3495
      Begin VB.CommandButton cmdprestamo 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Pr�stamo"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtpagoextra 
         Height          =   525
         Left            =   2520
         TabIndex        =   49
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmdpagoextra 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Pago E&xtra"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtentrega 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdapartado 
         BackColor       =   &H00FFC0C0&
         Caption         =   "A&partar"
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Pago Extra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   480
         Left            =   1680
         TabIndex        =   50
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Ya ha entregado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1560
         TabIndex        =   42
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Venta"
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   120
      TabIndex        =   40
      Top             =   6120
      Width           =   3615
      Begin MSComCtl2.DTPicker DtPicker1 
         Height          =   375
         Left            =   1920
         TabIndex        =   64
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   144506881
         CurrentDate     =   38197
      End
      Begin VB.CommandButton cmdventa 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Venta"
         Height          =   855
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox cmbpago 
         Height          =   315
         Left            =   1920
         TabIndex        =   28
         Text            =   "Forma de Pago"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label22 
         Caption         =   "Label22"
         Height          =   255
         Left            =   1320
         TabIndex        =   70
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label18 
         Caption         =   "Label18"
         Height          =   255
         Left            =   1920
         TabIndex        =   66
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Trabajo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1200
         Width           =   1695
      End
   End
   Begin VB.TextBox txtprefinal 
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   6960
      TabIndex        =   24
      Top             =   2040
      Width           =   975
   End
   Begin VB.ComboBox cmbdescu 
      Height          =   315
      Index           =   1
      ItemData        =   "frmventa.frx":0ECA
      Left            =   8160
      List            =   "frmventa.frx":0ECC
      TabIndex        =   25
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdborrar 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Borrar &Datos"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton CmbBorraArt 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Borrar &todos los Art�culos"
      Height          =   855
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdactualiza 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Actuali&za Datos"
      Height          =   255
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtidcliente 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmda�adir 
      BackColor       =   &H00FFC0C0&
      Caption         =   "A�adir Datos de &Nuevo Cliente"
      Height          =   495
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtfechanac 
      Height          =   285
      Left            =   7200
      TabIndex        =   6
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtlocalidad 
      Height          =   285
      Left            =   5160
      TabIndex        =   5
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtidarticulo 
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   32
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtprecio 
      Height          =   375
      Index           =   1
      Left            =   6000
      TabIndex        =   23
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txttalla 
      Height          =   375
      Index           =   1
      Left            =   5400
      TabIndex        =   22
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox txtcolor 
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   21
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox txttipo 
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   20
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtcodigo 
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   19
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdarticulo 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Buscar &Art�culo"
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtdireccion 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   960
      Width           =   3735
   End
   Begin VB.TextBox txttelefono 
      Height          =   285
      Left            =   7200
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtapellidos 
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Top             =   360
      Width           =   3735
   End
   Begin VB.TextBox txtnombre 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmdBuscaCliente 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Buscar Cliente"
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   210
      TabIndex        =   69
      Top             =   1680
      Width           =   645
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Final"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   7215
      TabIndex        =   39
      Top             =   1680
      Width           =   465
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "% Descuento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   8160
      TabIndex        =   38
      Top             =   1680
      Width           =   1200
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FF8080&
      Caption         =   "Cliente N�mero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   9360
      TabIndex        =   35
      Top             =   120
      Width           =   780
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Fecha Nacimiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   7200
      TabIndex        =   34
      Top             =   720
      Width           =   1635
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Localidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5160
      TabIndex        =   33
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Idart�culo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1395
      TabIndex        =   31
      Top             =   1680
      Width           =   825
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Talla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5400
      TabIndex        =   18
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Precio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6180
      TabIndex        =   17
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4800
      TabIndex        =   16
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4110
      TabIndex        =   15
      Top             =   1680
      Width           =   435
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "C�digo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2850
      TabIndex        =   14
      Top             =   1680
      Width           =   675
   End
   Begin VB.Line Line1 
      X1              =   239
      X2              =   9599
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Direcci�n"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1440
      TabIndex        =   13
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Tel�fono"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   7200
      TabIndex        =   12
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Apellidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3600
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1440
      TabIndex        =   10
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Venta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsdummy As Recordset
Dim dummyx As Currency
Dim cuentahoras As Integer
Dim dumIdCliente As Integer

Private Sub AbrePuerto()
'    On Error GoTo sehodio
    Msc.CommPort = "2"
    Msc.Settings = "28800,N,8,1"
    Msc.PortOpen = True
    Exit Sub
sehodio:
    MsgBox (Err.Number)
    MsgBox (Err.Description)
End Sub

Private Sub cmdElimina_Click(Index As Integer)
    If VentaApartado = True Then
        Desaparta txtidarticulo(Index)
    End If
    txtidarticulo(Index) = ""
    If NumArtVend = 1 Then NumArtVend = 0
    HaceSumaTotal
    
End Sub
Private Sub Desaparta(ByVal artic As Integer)
    With RsArtApartado
        .MoveFirst
        Do Until .EOF = True
            If !Idart = artic Then
                .Edit
                !apartado = False
                !vendido = False
                .Update
            End If
            .MoveNext
        Loop
    End With
    With RsDetalApartado
        .MoveFirst
        Do Until .EOF = True
            If !Idart = artic Then
                .Delete
                
            End If
            .MoveNext
        Loop
        .MoveFirst
        If .EOF = True Then RsApartado.Delete
    End With
    
End Sub

Private Sub CmdImprimeTicket_Click()
    HaceTicket
End Sub


Private Sub CmdTicketRegalo_Click()
    TicketRegalo = True
    HaceTicket
    
End Sub

Private Sub DtPicker1_Change()
    FechaTrabajo = HaceFecha(DTPicker1.Value)
    If FechaTrabajo <> Date Then Timer2.Enabled = False
    'MsgBox FechaTrabajo
End Sub

Private Sub dtpicker1_Click()
    FechaTrabajo = HaceFecha(DTPicker1.Value)
   ' MsgBox FechaTrabajo
    'Calendar1.Visible = False
    'txtfecha = FechaTrabajo
End Sub

Private Sub chkPago_Click()
    If chkPago.Value = 1 Then
        OptTodo(0).Visible = True
        OptTodo(0).Value = True
        OptTodo(1).Visible = True
        OptTodo(1).Value = False
        
    Else
        OptTodo(0).Visible = False
        OptTodo(0).Value = False
        OptTodo(1).Visible = False
        OptTodo(1).Value = False
        lbparte.Visible = False
        TxtParte.Visible = False
        
    End If
End Sub

Private Sub CmbBorraArt_Click()
    On Error Resume Next
    Dim dummy, i As Integer
    dummy = txtidarticulo.UBound
    If dummy > 1 Then
        For i = 2 To dummy
            Unload txtidarticulo(i)
            Unload txtcodigo(i)
            Unload txttipo(i)
            Unload txtprecio(i)
            Unload txtcolor(i)
            Unload txttalla(i)
            Unload cmbdescu(i)
            Unload txtprefinal(i)
            Unload cmdElimina(i)
        Next i
    End If
            txtidarticulo(1) = ""
            txtcodigo(1) = ""
            txttipo(1) = ""
            txtprecio(1) = ""
            txtcolor(1) = ""
            txttalla(1) = ""
            cmbdescu(1) = ""
            txtprefinal(1) = ""
   NumArtVend = 0
    txtTotal.Text = ""
    VentaApartado = False
    cmdapartado.Enabled = True
    
'inactivo porque la venta siguiente a un apartado toma el idventa anterior
'    IdCliente = 0
End Sub

Private Sub cmbdescu_Click(Index As Integer)
    On Error Resume Next
    txtprefinal(Index).Text = Format(txtprecio(Index).Text * (1 - (cmbdescu(Index).Text / 100)), "00.00")
'    HaceSumaTotal

End Sub

Private Sub cmbpago_Click()
    FormaPago = cmbpago.ItemData(cmbpago.ListIndex)
    MsgBox ("El pago con " & cmbpago.List(cmbpago.ListIndex) & " supone un descuento de " & cmbpago.ItemData(cmbpago.ListIndex) & " por ciento")
End Sub

Private Sub cmdactualiza_Click()
    Dim x
    If txtnombre = "" Then
        x = MsgBox("El cliente debe tener un nombre." & Chr(13) & "Proceso de registro anulado", vbExclamation)
        Exit Sub
    End If
    Set RsCliente = bdtienda.OpenRecordset("clientes")
    With RsCliente
        '.Index = "idcliente"
        '.Seek "=", txtidcliente
        .MoveFirst
        Do Until .EOF
        If !IdCliente = txtidcliente.Text Then
            .Edit
            !Nombre = txtnombre.Text
            !apellidos = txtapellidos
            If txttelefono <> "" Then !telefono = txttelefono
            !direccion = txtdireccion
            !localidad = txtlocalidad
            If txtfechanac <> "" Then !fechanac = txtfechanac
            .Update
            Exit Do
        End If
        .MoveNext
        Loop
    End With
    
End Sub

Private Sub cmda�adir_Click()
    Dim x
    If NuevoCliente = False Then
        x = MsgBox("Debe borrar los campos de datos y" & Chr(13) & "a�adir datos para el nuevo cliente", vbExclamation)
        Exit Sub
    End If
    If txtnombre = "" Then
        x = MsgBox("El cliente debe tener un nombre." & Chr(13) & "Proceso de registro anulado", vbExclamation)
        Exit Sub
    End If
    Set RsCliente = bdtienda.OpenRecordset("clientes")
    With RsCliente
        .AddNew
        !Nombre = txtnombre.Text
        !apellidos = txtapellidos
        If txttelefono = "" Then
            !telefono = 0
        Else
            !telefono = txttelefono
        End If
        !direccion = txtdireccion
        !localidad = txtlocalidad
        If txtfechanac <> "" Then
            !fechanac = txtfechanac
        End If
        .Update
        .MoveLast
        txtidcliente = !IdCliente
    End With
    If MsgBox("Cliente Creado." & Chr(13) & "�Desea a�adir su foto?", vbYesNo) = vbYes Then
        IdCliFoto = txtidcliente
        'FrmWebCam.Show
    End If
    
End Sub

Private Sub cmdapartado_Click()
    Modo = "Apartado"
    MarcaVenta
    VentaApartado = False
    frmalarma.Show 1
End Sub

Private Sub cmdarticulo_Click()
  On Error GoTo sehodio
    Dim idArtPrestaShop As Long

    ModoBusca = "articulos"
    If CodigoBusca = "" Then CodigoBusca = InputBox("Escriba el c�digo")

    If CodigoBusca <> "" Then
        ' ===== INTEGRACION PRESTASHOP: Intentar buscar en PrestaShop primero =====
        idArtPrestaShop = BuscarProductoPrestaShop(CodigoBusca)

        If idArtPrestaShop <> 0 Then
            ' Producto encontrado en PrestaShop y agregado a BD local
            SqlArticulos = "Select idart, codigo, tipo, precioventa, " _
            & " color, talla, extra from articulos where " _
            & " idart = " & idArtPrestaShop

            ' Log del SQL para debug
            ModuloLog.LogDebug "SQL generado para PrestaShop: " & SqlArticulos
        Else
            ' Si no esta en PrestaShop, buscar en BD local (comportamiento original)
'        SqlArticulos = "Select idart, idarticulo, codigo, tipo, precioventa, " _
'        & " color, talla, extra from articulos where vendido = false and apartado = false and" _
'        & " (idart = '" & CodigoBusca & "' or codigo like '*" & CodigoBusca & "*' or" _
'        & " idarticulo like '*" & CodigoBusca & "*') order by codigo" ' or idart like '*" & CodigoBusca & "*' or codigo like '*" & CodigoBusca & "*' order by codigo"
            SqlArticulos = "Select idart, codigo, tipo, precioventa, " _
            & " color, talla, extra from articulos where vendido = false and apartado = false and" _
            & " idart = " & CodigoBusca & " order by codigo" ' or idart like '*" & CodigoBusca & "*' or codigo like '*" & CodigoBusca & "*' order by codigo"
        End If
        ' ===== FIN INTEGRACION PRESTASHOP =====

'        SqlArticulos = "Select idart, idarticulo, codigo, tipo, precioventa, vendido, color, talla " _
'        & "from articulos where vendido = false and apartado = false and idart like " _
'        & "'*" & CodigoBusca & "*' or codigo like '*" & CodigoBusca & "*' or idarticulo like " _
'        & "'*" & CodigoBusca & "*' order by codigo" ' or idart like '*" & CodigoBusca & "*' or codigo like '*" & CodigoBusca & "*' order by codigo"
    Else: CodigoBusca = InputBox("Escriba alg�n dato para buscar")
        SqlArticulos = "Select idart, codigo, tipo, precioventa, color, talla, extra " _
        & "from articulos where vendido = false and apartado = false and(codigo " _
        & "like '*" & CodigoBusca & "*' or precioventa like '*" & CodigoBusca & "*' or " _
        & "talla like '*" & CodigoBusca & "*' or tipo like '*" & CodigoBusca & "*') order by codigo"
    End If

    ' ===== INTEGRACION PRESTASHOP: Cerrar recordset antes de abrirlo =====
    On Error Resume Next
    If Not RsArticulo Is Nothing Then
        RsArticulo.Close
        Set RsArticulo = Nothing
    End If
    On Error GoTo sehodio
    ' ===== FIN INTEGRACION PRESTASHOP =====

    'MsgBox (SqlArticulos)
    Set RsArticulo = bdtienda.OpenRecordset(SqlArticulos)

    ' ===== INTEGRACION PRESTASHOP: Log resultado de busqueda =====
    If idArtPrestaShop <> 0 Then
        If RsArticulo.EOF Then
            ModuloLog.LogError "Articulo PS no encontrado en BD despues de crearlo. SQL: " & SqlArticulos
        Else
            RsArticulo.MoveLast
            RsArticulo.MoveFirst
            ModuloLog.LogDebug "Articulo PS encontrado en BD. Records: " & RsArticulo.RecordCount
            ModuloLog.LogDebug "Datos articulo - idart: " & RsArticulo!Idart & " | tipo: " & RsArticulo!Tipo & " | precio: " & RsArticulo!PrecioVenta
        End If
    End If
    ' ===== FIN INTEGRACION PRESTASHOP =====
    If RsArticulo.EOF Then
        CodigoBusca = ""
        Exit Sub
    End If
    RsArticulo.MoveLast
    RsArticulo.MoveFirst  ' ===== INTEGRACION PRESTASHOP: Asegurar cursor en primera posicion =====
    If RsArticulo.RecordCount > 1 Then
        frmarticulos.Show
    Else
        NumArtVend = NumArtVend + 1
        ModuloLog.LogDebug "Llamando a PoneArticulos - NumArtVend: " & NumArtVend
        PoneArticulos
        ModuloLog.LogDebug "PoneArticulos ejecutado"

        ' ===== INTEGRACION PRESTASHOP: Popular ComboTallas si hay combinaciones =====
        If idArtPrestaShop <> 0 And HayProductoEnCache() Then
            Dim productoPS As ProductoPrestaShop
            Dim i As Integer

            productoPS = GetUltimoProductoEncontrado()

            If productoPS.TieneCombinaciones And productoPS.NumCombinaciones > 0 Then
                ' Popular ComboBox con tallas disponibles
                ComboTallas(NumArtVend).Clear
                For i = 1 To productoPS.NumCombinaciones
                    ComboTallas(NumArtVend).AddItem productoPS.Combinaciones(i).Talla
                Next i
                ComboTallas(NumArtVend).Enabled = True
                ComboTallas(NumArtVend).ListIndex = -1  ' Dejar vacio (usuario debe elegir)
                ModuloLog.LogDebug "ComboTallas populado con " & productoPS.NumCombinaciones & " tallas"
            Else
                ' Producto sin combinaciones - deshabilitar combo
                ComboTallas(NumArtVend).Clear
                ComboTallas(NumArtVend).Enabled = False
                ModuloLog.LogDebug "Producto sin combinaciones - ComboTallas deshabilitado"
            End If
        End If
        ' ===== FIN INTEGRACION PRESTASHOP =====
    End If
    CodigoBusca = ""
    Exit Sub
sehodio:
    MsgBox ("No se han encontrado datos")
End Sub

Private Sub cmdBorrar_Click()
    ' ===== INTEGRACION PRESTASHOP: Cancelar venta si hay articulos PS =====
    CancelarVenta
    ' ===== FIN INTEGRACION PRESTASHOP =====

    txtnombre = ""
    txtapellidos = ""
    txttelefono = ""
    txtdireccion = ""
    txtlocalidad = ""
    txtfechanac = ""
    txtidcliente = ""
    txtentrega = ""

    IdCliente = 0
    IdVenta = 0
    IdVentaApartado = 0
    Modo = "Venta"

    NuevoCliente = True
    CmbBorraArt_Click
    If VentaApartado = False Then
        cmdarticulo.Enabled = True
        cmdgenerico.Enabled = True
    End If

End Sub

Private Sub cmdBuscaCliente_Click()
'CmbBorraArt_Click
    txtentrega = ""
ModoBusca = "clientes"
    cmdapartado.Enabled = True
    
    CodigoBusca = InputBox("Escriba el C�digo de Cliente, nombre o apellido")
    If CodigoBusca <> "" Then
        SqlArticulos = "Select * from clientes where idcliente like '*" & CodigoBusca & "*' or nombre like '*" & CodigoBusca & "*' or apellidos like '*" & CodigoBusca & "*' order by idcliente"
    Else: CodigoBusca = InputBox("Escriba alg�n dato para buscar")
        SqlArticulos = "Select * from clientes where telefono like '*" & CodigoBusca & "*' or direccion like '*" & CodigoBusca & "*' or fechanac like '*" & CodigoBusca & "*' order by idcliente"
    End If
    Set RsArticulo = bdtienda.OpenRecordset(SqlArticulos)
    If RsArticulo.EOF Then Exit Sub
    RsArticulo.MoveLast
    
    frmarticulos.Show 1
    If VentaApartado = True Then
        cmdarticulo.Enabled = False
        cmdgenerico.Enabled = False
    Else
        cmdarticulo.Enabled = True
        cmdgenerico.Enabled = True
        
    End If
'    cmdapartado.Enabled = True
'    cmdpagoextra.Enabled = False
End Sub

Private Sub cmdcobrar_Click()
If NumArtVend <= 0 Then
    MsgBox "No hay art�culos a la venta"
    Exit Sub
End If
     Header.ACuenta = 0
     Header.fecha = Now
     Header.FormaPago = ""
     Header.IdCliente = 0
     Header.IdVenta = 0
     Header.Modo = ""
     Header.Nombre = ""
     Header.SumaTotal = 0
     Header.IVATotal = 0
     Header.Tarjeta = 0

SelectPrinter "Axiohm A793 CLASS 7193 Full"

PlayWave App.Path & "\ringin.wav"
If chkPago.Value = 1 Then
    If OptTodo(1).Value = 0 Then
        TxtParte.Text = 0
        PagoTarjeta = txtTotal.Text
        lbparte.Visible = False
        TxtParte.Visible = False
        Header.FormaPago = "Tarjeta"
    End If
End If
    If MsgBox("Total a pagar " & Chr(13) & "Efectivo: " & SumaTotal - PagoTarjeta - Val(txtentrega) & "�" _
        & Chr(13) & "Tarjeta: " & PagoTarjeta & "�", vbOKCancel, "ATENCI�N") = vbCancel Then Exit Sub
    
    Dim fechita As Date
    fechita = Format(FechaTrabajo, "Short Date")
    Header.fecha = FechaTrabajo
    If fechita <> Date Then
        FechaTrabajo = HaceFecha(fechita + Time)
    End If
    'Modo = "venta"
    If VentaApartado Then
        With RsApartado
            .Edit
            !ACuenta = 0
            !pagado = True
'**************OJO*******************
'Aqui duplicaba en la caja
'                If chkPago.Value = 1 Then
'                    If OptTodo(1).Value = 0 Then
'                        MueveCaja !IdVenta, FechaTrabajo, 0, PagoTarjeta
'                    Else
'                        MueveCaja !IdVenta, FechaTrabajo, !total - PagoTarjeta, PagoTarjeta
'
'                    End If
'
'                Else
'                        MueveCaja !IdVenta, FechaTrabajo, !total, PagoTarjeta
'
'                End If
            !total = dummyx
            .Update
        End With
        With RsArtApartado
            .MoveFirst
            If .EOF = True Then
                '.MoveFirst
                .Edit
                !apartado = False
                !vendido = True
                .Update
                .MoveNext
            Else
                Do Until .EOF
                .Edit
                !apartado = False
                !vendido = True
                .Update
                .MoveNext
                Loop
            End If
        End With
        
        Modo = "Venta"
        
        MarcaVenta
        CmbBorraArt_Click
        cmdBorrar_Click
        txtentrega = ""
        IdCliente = 0
        
    Else
        MarcaVenta
    End If
    
frmalarma.Show 1

    cmdBuscaCliente.SetFocus

    'cmdcobrar.Visible = False
    TxtParte.Visible = False
    OptTodo(0).Visible = False
    OptTodo(1).Visible = False
    chkPago.Value = False
    lbparte.Visible = False
    If VentaApartado = True Then
        cmdarticulo.Enabled = True
        cmdgenerico.Enabled = True
    End If
    Modo = "Venta"
End Sub
Private Sub ComprobarVenta()
Dim num As Integer
Dim RsVentaX As Recordset
Dim RsDetal As Recordset
Dim xx
Dim inicio As Integer
Dim totaleuros As Currency
'inicio = InputBox("inicio")
    Set RsVentaX = bdtienda.OpenRecordset("venta")
    With RsVentaX
        .MoveLast
        Set RsDetal = bdtienda.OpenRecordset("select * from detalleventa where idventa = " & IdVenta)
            If RsDetal.EOF = True Then
                xx = MsgBox("Atenci�n, NO SE HA REALIZADO LA VENTA CORRECTAMENTE" & Chr(13) & "DEBE BORRAR LA VENTA " & !IdVenta & " E INTENTARLO DE NUEVO", vbCritical)
            'Else
                'MsgBox ("Venta realizada correctamente")
            End If
            
        
    End With

End Sub
Private Sub cmddevuelto_Click()
    frmartidevol.Show
End Sub

Private Sub cmdfecha_Click()
    On Error GoTo sehodio
'    If txtfecha = "" Then Exit Sub
'    FechaTrabajo = txtfecha
    Exit Sub
sehodio:
    MsgBox ("Fecha no v�lida o formato incorrecto")
End Sub

Private Sub cmddevuelven_Click()
    Dim idcodigo As Integer
    Dim rsbuscar As Recordset
    Dim dumresult As String
    Dim tempidventa As Integer
    Dim importeventa As Currency
    On Error GoTo sehodio
    idcodigo = InputBox("�C�digo de la prenda?")
    Set rsbuscar = bdtienda.OpenRecordset("Select * from articulos where idart = " & idcodigo)
    If rsbuscar.EOF Then
        MsgBox ("no se ha encontrado el art�culo " & idcodigo)
        Exit Sub
    End If
    With rsbuscar
        dumresult = "Art�culo encontrado:" & Chr(13)
        dumresult = dumresult & !Idart
        dumresult = dumresult & Chr(13) & !Idart
        dumresult = dumresult & Chr(13) & !codigo
        dumresult = dumresult & Chr(13) & !Tipo
        dumresult = dumresult & Chr(13) & !PrecioVenta
        dumresult = dumresult & Chr(13) & !fechacompra
        dumresult = dumresult & Chr(13) & !fechaventa
        dumresult = dumresult & Chr(13) & !vendido
        MsgBox (dumresult)
        If MsgBox("�Desea dar el art�culo " & idcodigo & " como devuelto?", vbYesNo) = vbYes Then
            .Edit
            !vendido = False
            .Update
        End If
    End With
    
    Set RsDetalVenta = bdtienda.OpenRecordset("select * from detalleventa where idart = " & idcodigo)
    If RsDetalVenta.EOF = True Then
        MsgBox ("No se ha encontrado el detalleventa")
        Exit Sub
    Else
        tempidventa = RsDetalVenta!IdVenta
        importeventa = RsDetalVenta!PrecioFinal
        RsDetalVenta.Edit
        RsDetalVenta.Delete
'        RsDetalVenta.Update
        
    End If
    Dim RsVentaBorrar As Recordset
    Set RsVentaBorrar = bdtienda.OpenRecordset("select * from venta where idventa = " & tempidventa)
    With RsVentaBorrar
        If .EOF = True Then
            MsgBox ("No se ha encontrado la venta " & tempidventa)
        Else
        .Edit
        !total = !total - importeventa
        .Update
        End If
    End With
        Set RsArqueo = bdtienda.OpenRecordset("select * from arqueo where idventa = " & tempidventa)
    
    With RsArqueo
        If .EOF = True Then
            MsgBox ("No se ha encontrado el arqueo")
        Else
         .Edit
         !caja = !caja - importeventa
         .Update
        End If
    End With
    If MsgBox("�Desea hacer un vale de canje?", vbYesNo, "Vale") = vbYes Then
        HaceVale (importeventa)
    End If
    BlAlarmaQuitar = False
    frmalarma.Show 1
sehodio:
MsgBox (Err.Description)
End Sub

Private Sub cmdgenerico_Click()
    frmgenerico.Show 1
'    dumresp = InputBox("Precio del art�culo")

End Sub

Private Sub cmdllamar_Click()
'On Error Resume Next
If hablando = False Then
    Dim numtel As String
    numtel = txttelefono
    
    If numtel = "" Then Exit Sub
    If Len(numtel) = 6 Then numtel = "959" & numtel
    AbrePuerto
    Msc.Output = "ATDT" & numtel & ";" & vbCr
    hablando = True
    cmdllamar.Caption = "Colgar"
Else
    Msc.PortOpen = False
    cmdllamar.Caption = "Llamar"
    hablando = False
End If
End Sub
Private Sub cmdpagoextra_Click()
If MsgBox("�Desea a�adir un pago extra?", vbYesNo) = vbYes Then
     Modo = "Apartado"
    Dim x
    If NumArtVend = 0 Then
        MsgBox ("No hay ning�n art�culo en la venta")
        Exit Sub
    End If
    If Modo = "Apartado" Then
        If txtentrega.Text = "" Or txtnombre.Text = "" Or txtpagoextra = "" Then
            x = MsgBox("Para apartar es necesario: " & Chr(13) _
                & Chr(9) & "-Dar un nombre y tel�fono" & Chr(13) _
                & Chr(9) & "-Hacer una entrega en efectivo," & Chr(13) _
                & Chr(9) & "-O anotar un pago extra", vbCritical)
            Exit Sub
        End If
    End If
    Set RsVenta = bdtienda.OpenRecordset("select * from venta where idventa = " & RsApartado!IdVenta)
    
    With RsVenta
        .Edit
'        If IdCliente <> 0 Then
'            !IdCliente = IdCliente
'        Else
'            !IdCliente = 2
'        End If
        dummyx = !total + !ACuenta 'CCur(txttotal.Text) - CCur(txtpagoextra.Text)
        'dummyx = dummyx - !acuenta
        !total = dummyx - CCur(txtpagoextra.Text) - !ACuenta
        !ACuenta = CCur(txtentrega.Text) + CCur(txtpagoextra)
        !Descuento = DescuentoTotal
        .Update
        
        Header.Modo = "Apartado"
        Header.ACuenta = !ACuenta
        Header.SumaTotal = dummyx
        Header.IdCliente = txtidcliente
        Header.IdVenta = !IdVenta
        Header.fecha = FechaTrabajo
       ' MsgBox FechaTrabajo
        
        Header.Tarjeta = PagoTarjeta
        
        HaceTicket
        
        IdVenta = !IdVenta
        
        
        MueveCaja IdVenta, FechaTrabajo, txtpagoextra, PagoTarjeta, , InputBox("�Desea dal alg�n concepto a la entrada?")
        x = MsgBox(Chr(9) & "Total pagado " & Chr(9) & !ACuenta & " �" & Chr(13) _
            & Chr(9) & "Total por pagar " & Chr(9) & dummyx - !ACuenta & " �" & Chr(13) _
            & Chr(13) & Chr(9) & "Suma Total     " & Chr(9) & dummyx & " �", vbCritical)
    End With
        

    RsVenta.Close
    CmbBorraArt_Click
    cmdBorrar_Click
    txtpagoextra = ""
    txtentrega = ""
    dummyx = 0
    
End If
    VentaApartado = False

End Sub

Private Sub cmdprestamo_Click()

If IdCliente = 0 Then
    MsgBox ("Debe seleccionar un cliente para pr�stamo")
    Exit Sub
End If
If NumArtVend = 0 Then
    MsgBox ("Debe seleccionar alg�n art�culo para pr�stamo")
    Exit Sub
End If
Dim hacecomentario As String
hacecomentario = InputBox("�Desea anotar alg�n comentario?")
    Dim i As Integer
    Set RsPrestamo = bdtienda.OpenRecordset("prestamo")
    Dim SumaX As Currency
    With RsPrestamo
        For i = 1 To NumArtVend
        .AddNew
            !Idart = txtidarticulo(i)
            !IdCliente = IdCliente
            !fecha = FechaTrabajo
            !comentario = hacecomentario
        .Update
        ReDim Preserve Ticket(i)
        Ticket(i).Idart = txtidarticulo(i)
        Ticket(i).PrecioFinal = CCur(txtprefinal(i))
        Ticket(i).Descripcion = txtcodigo(i)
        SumaX = SumaX + CCur(txtprefinal(i))
        Next i
        Header.fecha = FechaTrabajo
        Header.Modo = "PRESTAMO"
        Header.IdCliente = IdCliente
        Header.SumaTotal = SumaX
        
    End With
    HaceTicket
    CmbBorraArt_Click
    cmdBorrar_Click
    txtentrega = ""
    txtpagoextra = ""
    IdCliente = 0
    BlAlarmaQuitar = True
    frmalarma.Show 1

End Sub


Private Sub cmdventa_Click()
If NumArtVend > 0 Then cmdcobrar.Visible = True
End Sub
Private Sub MarcaVenta()
    Erase Ticket
    Dim x
    Dim Y
    Dim SqlVenta As String
    If IdCliente = 0 Then
        If txtidcliente.Text <> "" Then IdCliente = txtidcliente
    End If
    If NumArtVend = 0 Then
        MsgBox ("No hay ning�n art�culo en la venta")
        Exit Sub
    End If
    
    If Modo = "Apartado" Then
        If IdCliente = 0 Then
            MsgBox ("Primero debe dar de alta al usuario")
            Exit Sub
        End If
        If txtentrega.Text = "" Or txtnombre.Text = "" Then
            x = MsgBox("Para apartar es necesario: " & Chr(13) _
                & Chr(9) & "-Hacer una entrega en efectivo" & Chr(13) _
                & Chr(9) & "-Dar un nombre y tel�fono", vbCritical)
            Exit Sub
        End If
    Else
        SqlVenta = "venta"
    End If
    
    If IdVentaApartado = 0 Then
        SqlVenta = "venta"
    Else
         SqlVenta = "Select * from venta where idventa = " & IdVentaApartado
        IdVentaApartado = 0
   
    End If
    
    Set RsVenta = bdtienda.OpenRecordset(SqlVenta)
    
    With RsVenta
        If .RecordCount = 0 Or .RecordCount > 1 Then
            .AddNew
        ElseIf .RecordCount = 1 Then
            .Edit
        End If
        'If Modo = "Venta" Then .AddNew Else .Edit
        If IdCliente <> 0 Then
            !IdCliente = IdCliente
        Else
            !IdCliente = 2
        End If
        Header.IdCliente = !IdCliente
        'Header.Nombre = !Nombre & " " & !apellido
        If chkPago.Value = 0 Then
            
            !total = txtTotal
        ElseIf chkPago.Value = 1 Then
            If OptTodo(0).Value = True Then
                !total = 0
            ElseIf OptTodo(1).Value = True Then
                !total = txtTotal
            End If
        End If
        
        Header.SumaTotal = !total
        Header.IVATotal = Header.SumaTotal * 0.21
        Header.BaseTotal = Header.SumaTotal * 0.79
        
        If Modo = "Venta" Then
            !pagado = True
            !Tarjeta = PagoTarjeta
            Header.Tarjeta = PagoTarjeta
           ' MsgBox FechaTrabajo
            
            MueveCaja !IdVenta, FechaTrabajo, !total, !Tarjeta
            Header.Modo = "Venta"
        ElseIf Modo = "Apartado" Then
            Y = CCur(txtTotal.Text) - CCur(txtentrega.Text)
            !total = Y
            !ACuenta = CCur(txtentrega.Text)
            Header.ACuenta = CCur(txtentrega.Text)
            Header.Modo = "Apartado"
           ' MsgBox FechaTrabajo
            
            MueveCaja !IdVenta, FechaTrabajo, !ACuenta, !Tarjeta
        End If
        !Descuento = DescuentoTotal
        IdVenta = !IdVenta
        Header.fecha = FechaTrabajo
        Header.IdVenta = IdVenta
        .Update
        '.MoveLast
        
    End With
    
    
    RsVenta.Close
    Dim dumporciento As Single
    Dim i As Integer
    Set RsDetalVenta = bdtienda.OpenRecordset("Select * from detalleventa where idventa = " & IdVenta & " order by idart")
        With RsDetalVenta
            ReDim Preserve Ticket(NumArtVend)
            For i = 1 To NumArtVend
                If txtidarticulo(i) <> "" Then
                    If .EOF = True Then
                        .AddNew
                    Else
                        .MoveFirst
                        Do Until .EOF = True
                            If !Idart = txtidarticulo(i).Text Then Exit Do
                            .MoveNext
                        Loop
                        .Edit
                    End If
                    !IdVenta = IdVenta
                    !IdCliente = IdCliente
                    !Idart = txtidarticulo(i).Text
                    !PrecioFinal = CCur(txtprefinal(i).Text)
                    .Update
                    Ticket(i).Idart = txtidarticulo(i).Text
                    Ticket(i).PrecioFinal = CCur(txtprefinal(i).Text)
                    Ticket(i).Descripcion = txttipo(i).Text '& " " & txtcodigo(i).Text
                    '.Update
                 '   Ticket(i).Descuento = CCur(cmbdescu(i).Text)
                End If
            Next i
        End With
    RsDetalVenta.Close
    HaceTicket
    
    ComprobarVenta

    MarcaVendido

    ' ===== INTEGRACION PRESTASHOP: Sincronizar stock despues de venta =====
    SincronizarStockVendido
    ' ===== FIN INTEGRACION PRESTASHOP =====

    CmbBorraArt_Click
    cmdBorrar_Click
    txtentrega = ""
    txtpagoextra = ""
    IdCliente = 0
    chkPago.Value = 0
End Sub

Private Sub cmdventasxcliente_Click()
    If IdCliente = 0 Then
        If txtidcliente.Text = 0 Or txtidcliente.Text = "" Then
            MsgBox ("Elige el cliente")
        Else
            frmventacliente.Show
        End If
    Else
            frmventacliente.Show
    End If
    
    
End Sub

Private Sub cmdvolver_Click()
Elige.Show
Unload Me
End Sub

Private Sub Command1_Click()
'Dim x
If MsgBox("�Desea salir?", vbYesNo) = vbYes Then
    End
End If
End Sub

Private Sub Form_Activate()
   ' PoneArticulos
    hablando = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim CtrlDown
    CtrlDown = (Shift And vbCtrlMask) > 0
    If KeyCode = 65 And CtrlDown Then
        AbrirCajon
    End If

End Sub

Private Sub Form_Load()
    Dim i As Integer

    ' ===== INTEGRACION PRESTASHOP: Inicializar sistema =====
    InicializarIntegracion
    ' ===== FIN INTEGRACION PRESTASHOP =====

    NuevoCliente = True
    NumArtVend = 0
    Set rsdummy = bdtienda.OpenRecordset("pago")
    Do Until rsdummy.EOF
        cmbpago.AddItem rsdummy!FormaPago
        cmbpago.ItemData(cmbpago.NewIndex) = rsdummy!Descuento
        rsdummy.MoveNext
    Loop
    For i = 5 To 100 Step 5
        cmbdescu(1).AddItem i
    Next i
    FechaTrabajo = HaceFecha(Now)
       ' MsgBox FechaTrabajo
    ' FechaTrabajo
    DTPicker1.Value = Now
    Me.Show
    txtbusca.SetFocus
    Modo = "Venta"
    'Calendar1.Value = Now
    'Set rsarticulo = bdtienda.OpenRecordset("articulos")
'    PoneArticulos
End Sub
Public Sub PoneArticulos()
    If RsArticulo.EOF Then MsgBox ("no hay art�culos")
    If NumArtVend > 1 Then
        A�adeControlesArticulos
    End If
        With RsArticulo
        txtidarticulo(NumArtVend).Text = !Idart
        txtcodigo(NumArtVend).Text = "" & !codigo
        txttipo(NumArtVend).Text = "" & !Tipo
        txtprecio(NumArtVend).Text = "" & !PrecioVenta
        txtcolor(NumArtVend).Text = "" & !Color
        txttalla(NumArtVend).Text = "" & !talla
        If cmbdescu(NumArtVend).Text <> "" Then
            txtprefinal(NumArtVend).Text = txtprecio(NumArtVend).Text * (1 - (cmbdescu(NumArtVend).Text * 100))
        Else
            txtprefinal(NumArtVend).Text = txtprecio(NumArtVend).Text
        End If
        End With
    HaceSumaTotal

End Sub
Public Sub PoneArticuloGenerico(ByVal Dumprecio As Currency)
    Set RsArticulo = bdtienda.OpenRecordset("articulos")
    If Dumprecio = 0 Or Dumprecio > 500 Then
        MsgBox ("Precio no v�lido. Int�ntelo de nuevo")
        Exit Sub
    End If
    
    'Dumtipo = InputBox("Tipo de art�culo (traje, vestido, etc.)")
    If Dumtipo = "" Then
        MsgBox ("Debe indicar el tipo de art�culo." & Chr(13) & "Art�culo no a�adido")
        
        NumArtVend = NumArtVend - 1
        Exit Sub
    End If
    'If RsArticulo.EOF Then MsgBox ("no hay art�culos")
    If NumArtVend > 1 Then
        A�adeControlesArticulos
    End If
        With RsArticulo
        .AddNew
        !PrecioVenta = Dumprecio
        !PrecioCompra = Dumprecio / 2
        If BlGoyse = True Then
            !Tipo = Dumtipo
            !codigo = GoyseCode
            BlGoyse = False
        Else
            !Tipo = Dumtipo
            !codigo = DumCode '"gen�rico"
        End If
        txtidarticulo(NumArtVend).Text = !Idart
        txtcodigo(NumArtVend).Text = DumCode '"" & !codigo
        txttipo(NumArtVend).Text = "" & !Tipo
        txtprecio(NumArtVend).Text = Dumprecio
        txtcolor(NumArtVend).Text = "" & !Color
        txttalla(NumArtVend).Text = "" & !talla
        If cmbdescu(NumArtVend).Text <> "" Then
            txtprefinal(NumArtVend).Text = txtprecio(NumArtVend).Text * (1 - (cmbdescu(NumArtVend).Text * 100))
        Else
            txtprefinal(NumArtVend).Text = txtprecio(NumArtVend).Text
        End If
        .Update
        End With
    HaceSumaTotal

End Sub
Public Sub PoneClientes()
    On Error GoTo sehodio
'    CmbBorraArt_Click
    If RsArticulo.EOF Then Exit Sub
    With RsArticulo
        txtidcliente = !IdCliente
        txtnombre = !Nombre
        txtapellidos = !apellidos
        txttelefono = "" & !telefono
        txtdireccion = !direccion
        txtfechanac = "" & !fechanac
        txtlocalidad = !localidad
    End With
    IdVentaApartado = 0
    NuevoCliente = False
    DevuelvePrestamo
    Exit Sub
sehodio:
    Exit Sub
End Sub
Private Sub A�adeControlesArticulos()
        Dim i As Integer
        Load txtidarticulo(NumArtVend)
        Load txtcodigo(NumArtVend)
        Load txttipo(NumArtVend)
        Load txtprecio(NumArtVend)
        Load txtcolor(NumArtVend)
        Load txttalla(NumArtVend)
        Load cmbdescu(NumArtVend)
        Load txtprefinal(NumArtVend)
        Load cmdElimina(NumArtVend)
        txtprefinal(NumArtVend).Text = ""
        
        txtidarticulo(NumArtVend).Top = txtidarticulo(NumArtVend - 1).Top + 400
        txtcodigo(NumArtVend).Top = txtcodigo(NumArtVend - 1).Top + 400
        txttipo(NumArtVend).Top = txttipo(NumArtVend - 1).Top + 400
        txtprecio(NumArtVend).Top = txtprecio(NumArtVend - 1).Top + 400
        txtcolor(NumArtVend).Top = txtcolor(NumArtVend - 1).Top + 400
        txttalla(NumArtVend).Top = txttalla(NumArtVend - 1).Top + 400
        cmbdescu(NumArtVend).Top = cmbdescu(NumArtVend - 1).Top + 400
        txtprefinal(NumArtVend).Top = txtprefinal(NumArtVend - 1).Top + 400
        cmdElimina(NumArtVend).Top = cmdElimina(NumArtVend - 1).Top + 400
        
        txtidarticulo(NumArtVend).Visible = True
        txtcodigo(NumArtVend).Visible = True
        txttipo(NumArtVend).Visible = True
        txtprecio(NumArtVend).Visible = True
        txtcolor(NumArtVend).Visible = True
        txttalla(NumArtVend).Visible = True
        cmbdescu(NumArtVend).Visible = True
        txtprefinal(NumArtVend).Visible = True
        cmdElimina(NumArtVend).Visible = True
        
        For i = 5 To 25 Step 5
            cmbdescu(NumArtVend).AddItem i
        Next i
        For i = 30 To 70 Step 10
            cmbdescu(NumArtVend).AddItem i
        Next i
        cmbdescu(NumArtVend).Text = ""

End Sub


Private Sub Frame3_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

Private Sub lbcambio_Click()

End Sub




Private Sub OptTodo_Click(Index As Integer)
If Index = 1 Then
    lbparte.Visible = True
    TxtParte.Visible = True
    TxtParte.Text = 0
    TxtParte.SelStart = 0
    TxtParte.SelLength = Len(TxtParte.Text)
    TxtParte.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
    cuentahoras = cuentahoras + 1
    Dim fechita As Date
    fechita = Format(FechaTrabajo, "Short Date")
    If cuentahoras >= 5 Then
'         If fechita <> Date Then
'            If MsgBox("Est� trabajando con fecha " & fechita & Chr(13) & "�Desea cambiarla a la fecha de hoy?", vbYesNo) = vbYes Then
'                FechaTrabajo = Now
'                DTPicker1.Value = FechaTrabajo
'                Timer2.Enabled = True
'            End If
'         End If
                Timer2.Enabled = True
        cuentahoras = 0
    End If
End Sub



Private Sub Timer2_Timer()
    FechaTrabajo = HaceFecha(Now)
    Label18 = FechaTrabajo
End Sub

Private Sub Timer3_Timer()
    Label22 = FechaTrabajo
End Sub

Private Sub TxtBusca_Click()
        txtbusca.SelStart = 0
        txtbusca.SelLength = Len(txtbusca.Text)

End Sub

Private Sub TxtBusca_GotFocus()
        txtbusca.SelStart = 0
        txtbusca.SelLength = Len(txtbusca.Text)

End Sub

Private Sub TxtBusca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(txtbusca.Text) < 5 Then Exit Sub
        CodigoBusca = Left(txtbusca.Text, 5)
        cmdarticulo_Click
        txtbusca.Text = ""
        txtbusca.SelStart = 0
        txtbusca.SelLength = Len(txtbusca.Text)
    End If
End Sub

Private Sub txtidarticulo_Click(Index As Integer)
   ' MarcaVendido
End Sub

Private Sub txtidcliente_Change()
    On Error Resume Next
    IdCliente = txtidcliente
End Sub

Public Sub MarcaVendido()
    Dim i As Integer
    Set RsArticulo = bdtienda.OpenRecordset("articulos", dbOpenTable)
    With RsArticulo
    .Index = "idart"
    For i = 1 To NumArtVend
        If txtidarticulo(i) <> "" Then
            .Seek "=", txtidarticulo(i)
            .Edit
            If Modo = "Venta" Then
                !vendido = True
            ElseIf Modo = "Apartado" Then
                !apartado = True
            End If
            
            .Update
        End If
            'MsgBox (!idart)
    Next i
    End With
        
End Sub

Private Sub HaceSumaTotal()
    Dim i As Integer
   On Error Resume Next
    DescuentoTotal = 0
    SumaTotal = 0
    PagoTarjeta = 0
    Dim dummy
    Dim dummx
    dummx = txtprefinal.Count
    For i = 1 To txtprefinal.Count
        If txtidarticulo(i) <> "" Then
            dummy = (txtprefinal(i).Text)
            SumaTotal = SumaTotal + dummy
            DescuentoTotal = DescuentoTotal + (txtprecio(i).Text - txtprefinal(i).Text)
        End If
    Next i
    If VentaApartado = True Then
        txtTotal.Text = CCur(SumaTotal) - CCur(txtentrega)
    Else
        txtTotal.Text = CCur(SumaTotal)
    End If

End Sub

Private Sub TxtParte_Change()
    If SumaTotal = 0 Then Exit Sub
    txtTotal = SumaTotal - Val(TxtParte) - Val(txtentrega)
    PagoTarjeta = Val(TxtParte)
End Sub

Private Sub txtprecio_Change(Index As Integer)
    'HaceSumaTotal
End Sub

Private Sub txtprefinal_Change(Index As Integer)
    HaceSumaTotal
End Sub
Public Sub BuscaVentaApartada()
    Dim x As Integer
    
    Set RsApartado = bdtienda.OpenRecordset("select * from venta where idcliente like " & IdCliente & " and pagado = false")
    If RsApartado.RecordCount > 0 Then
        CmbBorraArt_Click

        frmapartadoxcliente.Show 1
        If ArtApartParaPagar <> 0 Then
        'If MsgBox("Tiene articulos apartados desde el " & Chr(13) _
            & Chr(9) & RsApartado!Fecha & Chr(13) _
            & "�Quiere pagarlos?", vbYesNo) = vbYes Then
            Set RsApartado = bdtienda.OpenRecordset("select * from venta where" _
            & " idventa = " & ArtApartParaPagar & " and pagado = false")
            ArtApartParaPagar = 0
            
            On Error GoTo sehodio
            Set RsDetalApartado = bdtienda.OpenRecordset("select idart, preciofinal from detalleventa where idventa like " & RsApartado!IdVenta & " order by idart")
            If RsDetalApartado.EOF = True Then
                If MsgBox("La venta " & RsApartado!IdVenta & " no tiene art�culos asociados." & Chr(13) & "�Desea borrar esta venta?", vbOKCancel) = vbOK Then
                    
                    RsApartado.Delete
                    Exit Sub
                End If
            End If
            RsDetalApartado.MoveLast
            
            txtentrega.Text = RsApartado!ACuenta
            VentaApartado = True
            cmdapartado.Enabled = False
'            cmdapartado.Enabled = False
'            cmdpagoextra.Enabled = True
            
            ReDim IdArtApart(RsDetalApartado.RecordCount - 1)
            ReDim PreFinalApart(RsDetalApartado.RecordCount - 1)
            RsDetalApartado.MoveFirst
            For x = 0 To RsDetalApartado.RecordCount - 1
                IdArtApart(x) = RsDetalApartado!Idart
                PreFinalApart(x) = RsDetalApartado!PrecioFinal
                RsDetalApartado.MoveNext
            Next x
            A�adeArticulosApartados
        End If
    End If
    Exit Sub
sehodio:
MsgBox ("No se han encontrado art�culos de la venta " & RsApartado!IdVenta)
End Sub
Private Sub borraventa(ByVal IdVentaApartado As Integer)
End Sub
Private Sub A�adeArticulosApartados()
    cmdapartado.Enabled = False
    
    Dim dummy, x As Integer
    dummy = 0
    Dim dummystr As String
    If UBound(IdArtApart) > 0 Then
        For x = 1 To UBound(IdArtApart)
            dummystr = dummystr & " or idart like " & IdArtApart(x)
        Next x
    End If
    dummystr = "Select idart, codigo, tipo, apartado, vendido, precioventa, " _
        & "color, talla from articulos where apartado = true and idart like " _
        & IdArtApart(0) & dummystr & " order by idart"
    'MsgBox (dummystr)
    Set RsArtApartado = bdtienda.OpenRecordset(dummystr)
        If RsArtApartado.EOF = True Then
            If MsgBox("La venta " & RsApartado!IdVenta & " no tiene art�culos asociados." & Chr(13) & "�Desea borrar esta venta?", vbOKCancel) = vbOK Then
                
                RsApartado.Delete
                Exit Sub
            End If
        End If
    
    With RsArtApartado
        Do Until .EOF
          dummy = dummy + 1
          NumArtVend = dummy
          If dummy > 1 Then
              A�adeControlesArticulos
          End If
        
          txtidarticulo(dummy).Text = !Idart
          txtcodigo(dummy).Text = "" & !codigo
          txttipo(dummy).Text = "" & !Tipo
          txtprecio(dummy).Text = PreFinalApart(dummy - 1)
          txtcolor(dummy).Text = "" & !Color
          txttalla(dummy).Text = "" & !talla
          If cmbdescu(dummy).Text <> "" Then
              txtprefinal(dummy).Text = txtprecio(dummy).Text * (1 - (cmbdescu(dummy).Text * 100))
          Else
              txtprefinal(dummy).Text = txtprecio(dummy).Text
          End If
          ReDim Preserve Ticket(dummy)
          Ticket(dummy).Idart = !Idart
          Ticket(dummy).Descripcion = !Tipo
          Ticket(dummy).PrecioFinal = txtprefinal(dummy)
          .MoveNext
        Loop
    End With
    HaceSumaTotal
dummyx = Val(txtentrega) + Val(txtTotal)
'cmdapartado.Enabled = False
End Sub
