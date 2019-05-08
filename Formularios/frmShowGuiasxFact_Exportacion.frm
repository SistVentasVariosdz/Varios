VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "gridex20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmShowGuiasxFact_Prendas 
   Caption         =   "Autorización Facturas - Exportación"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14190
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   14190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPrecio 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Modificación"
      Height          =   2520
      Left            =   5580
      TabIndex        =   56
      Top             =   2160
      Visible         =   0   'False
      Width           =   3030
      Begin VB.TextBox txtImp_comision 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1590
         TabIndex        =   72
         Text            =   "0"
         Top             =   1260
         Width           =   1125
      End
      Begin VB.TextBox txtPorc_Descuento_Precio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1590
         TabIndex        =   58
         Text            =   "0"
         Top             =   390
         Width           =   540
      End
      Begin VB.TextBox txtPre_Unitario 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1590
         TabIndex        =   59
         Text            =   "0"
         Top             =   825
         Width           =   1125
      End
      Begin VB.CommandButton cmdAceptarPrecio 
         Caption         =   "Aceptar"
         Height          =   500
         Left            =   495
         TabIndex        =   60
         Top             =   1770
         Width           =   990
      End
      Begin VB.CommandButton cmdCancelarPrecio 
         Caption         =   "Cancelar"
         Height          =   500
         Left            =   1545
         TabIndex        =   62
         Top             =   1770
         Width           =   990
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Importe Comisión"
         Height          =   300
         Left            =   135
         TabIndex        =   73
         Top             =   1305
         Width           =   1485
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0FFC0&
         Caption         =   "% Descuento Precio"
         Height          =   435
         Left            =   135
         TabIndex        =   61
         Top             =   360
         Width           =   1410
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Precio Unitario"
         Height          =   315
         Left            =   135
         TabIndex        =   57
         Top             =   870
         Width           =   1485
      End
   End
   Begin VB.Frame fraDatosAdicionales 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Datos Adicionales"
      Height          =   8205
      Left            =   5880
      TabIndex        =   38
      Top             =   180
      Visible         =   0   'False
      Width           =   7875
      Begin VB.TextBox txtImp_Transporte_Pais_Destino 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6060
         TabIndex        =   76
         Text            =   "0"
         Top             =   3660
         Width           =   1125
      End
      Begin VB.TextBox txtImp_Desaduanaje 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1740
         TabIndex        =   75
         Text            =   "0"
         Top             =   3690
         Width           =   1125
      End
      Begin VB.TextBox txtPor_Comision 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6285
         MaxLength       =   10
         TabIndex        =   24
         Top             =   7185
         Width           =   930
      End
      Begin VB.TextBox txtRef_Embarque 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1800
         TabIndex        =   70
         Top             =   225
         Width           =   1830
      End
      Begin VB.TextBox txtCod_Class 
         Height          =   315
         Left            =   3930
         MaxLength       =   10
         TabIndex        =   23
         Top             =   7185
         Width           =   1125
      End
      Begin VB.TextBox txtCod_Vendor 
         Height          =   315
         Left            =   1755
         MaxLength       =   20
         TabIndex        =   22
         Top             =   7185
         Width           =   1620
      End
      Begin VB.TextBox txtPie_Pagina2 
         Height          =   885
         Left            =   1755
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   21
         Top             =   6105
         Width           =   5940
      End
      Begin VB.TextBox txtPie_Pagina1 
         Height          =   885
         Left            =   1740
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   20
         Top             =   5025
         Width           =   5940
      End
      Begin VB.TextBox txtCod_Embarque 
         Height          =   345
         Left            =   1740
         TabIndex        =   18
         Top             =   4095
         Width           =   585
      End
      Begin VB.TextBox txtDes_Embarque 
         Height          =   345
         Left            =   2385
         TabIndex        =   64
         Top             =   4095
         Width           =   4815
      End
      Begin VB.TextBox txtNom_Embarque 
         Height          =   315
         Left            =   1740
         TabIndex        =   19
         Top             =   4530
         Width           =   2340
      End
      Begin VB.TextBox txtDes_Termino_Venta 
         Height          =   345
         Left            =   2385
         TabIndex        =   54
         Top             =   2835
         Width           =   4815
      End
      Begin VB.TextBox txtCod_Termino_Venta 
         Height          =   345
         Left            =   1740
         TabIndex        =   14
         Top             =   2835
         Width           =   585
      End
      Begin VB.TextBox txtImp_Descuento 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6060
         TabIndex        =   17
         Text            =   "0"
         Top             =   3240
         Width           =   1125
      End
      Begin VB.TextBox txtImp_Flete 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1740
         TabIndex        =   15
         Text            =   "0"
         Top             =   3255
         Width           =   1125
      End
      Begin VB.TextBox txtImp_Seguro 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3780
         TabIndex        =   16
         Text            =   "0"
         Top             =   3240
         Width           =   1125
      End
      Begin VB.TextBox txtCod_CondVent 
         Height          =   285
         Left            =   1755
         TabIndex        =   12
         Top             =   2055
         Width           =   585
      End
      Begin VB.TextBox txtDes_CondVent 
         Height          =   285
         Left            =   2400
         TabIndex        =   48
         Top             =   2055
         Width           =   4815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancelar"
         Height          =   500
         Left            =   3945
         TabIndex        =   26
         Top             =   7575
         Width           =   990
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   500
         Left            =   2895
         TabIndex        =   25
         Top             =   7575
         Width           =   990
      End
      Begin VB.TextBox txtObservacion 
         Height          =   885
         Left            =   1755
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   675
         Width           =   5940
      End
      Begin VB.CommandButton cmdLugEnt 
         Caption         =   "..."
         Height          =   315
         Left            =   7320
         TabIndex        =   40
         ToolTipText     =   "Ver/Agregar Lugares de Entrega"
         Top             =   1695
         Width           =   375
      End
      Begin VB.TextBox txtLinea1 
         Height          =   285
         Left            =   2400
         TabIndex        =   39
         Top             =   1665
         Width           =   4815
      End
      Begin VB.TextBox txtSecuencia 
         Height          =   285
         Left            =   1755
         TabIndex        =   11
         Top             =   1665
         Width           =   585
      End
      Begin VB.TextBox txtCartaCredito 
         Height          =   315
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2430
         Width           =   2040
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Importe Transporte en Pais Destino"
         Height          =   315
         Left            =   3075
         TabIndex        =   78
         Top             =   3735
         Width           =   2715
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Importe Desaduanaje"
         Height          =   255
         Left            =   150
         TabIndex        =   77
         Top             =   3705
         Width           =   1605
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0FFFF&
         Caption         =   "% Comisión"
         Height          =   315
         Left            =   5325
         TabIndex        =   74
         Top             =   7245
         Width           =   900
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Número de Embarque"
         Height          =   255
         Left            =   180
         TabIndex        =   71
         Top             =   300
         Width           =   1545
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Class"
         Height          =   315
         Left            =   3450
         TabIndex        =   69
         Top             =   7230
         Width           =   420
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cod.Vendor"
         Height          =   255
         Left            =   165
         TabIndex        =   68
         Top             =   7260
         Width           =   1485
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Pie Factura 2:"
         Height          =   195
         Left            =   135
         TabIndex        =   67
         Top             =   6165
         Width           =   990
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Pie Factura 1:"
         Height          =   195
         Left            =   150
         TabIndex        =   66
         Top             =   5085
         Width           =   990
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Modo de Transporte"
         Height          =   315
         Left            =   150
         TabIndex        =   65
         Top             =   4155
         Width           =   1590
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Nombre Transporte"
         Height          =   270
         Left            =   150
         TabIndex        =   63
         Top             =   4590
         Width           =   1485
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Terminos de Ventas"
         Height          =   285
         Left            =   135
         TabIndex        =   55
         Top             =   2895
         Width           =   1590
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Importe Descuento"
         Height          =   495
         Left            =   5055
         TabIndex        =   53
         Top             =   3225
         Width           =   1080
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Importe Flete"
         Height          =   255
         Left            =   150
         TabIndex        =   51
         Top             =   3300
         Width           =   1485
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Importe Seguro"
         Height          =   465
         Left            =   3000
         TabIndex        =   50
         Top             =   3210
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Condic.de Ventas"
         Height          =   330
         Left            =   150
         TabIndex        =   49
         Top             =   2115
         Width           =   1590
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Observaciones :"
         Height          =   195
         Left            =   135
         TabIndex        =   43
         Top             =   735
         Width           =   1155
      End
      Begin VB.Label lbObservacion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1455
         TabIndex        =   47
         Top             =   525
         Width           =   45
      End
      Begin VB.Label lbCalidad 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   7335
         TabIndex        =   46
         Top             =   195
         Width           =   45
      End
      Begin VB.Label lbComb 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   6135
         TabIndex        =   45
         Top             =   225
         Width           =   45
      End
      Begin VB.Label lbDesTela 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   735
         TabIndex        =   44
         Top             =   195
         Width           =   45
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Lugar de Entrega"
         Height          =   330
         Left            =   150
         TabIndex        =   42
         Top             =   1710
         Width           =   1590
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Carta de Credito"
         Height          =   270
         Left            =   150
         TabIndex        =   41
         Top             =   2505
         Width           =   1485
      End
   End
   Begin GridEX20.GridEX GridEX4 
      Height          =   2055
      Left            =   4215
      TabIndex        =   35
      Top             =   4350
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3625
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ScrollToolTipColumn=   ""
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmShowGuiasxFact_Exportacion.frx":0000
      Column(2)       =   "frmShowGuiasxFact_Exportacion.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmShowGuiasxFact_Exportacion.frx":016C
      FormatStyle(2)  =   "frmShowGuiasxFact_Exportacion.frx":02A4
      FormatStyle(3)  =   "frmShowGuiasxFact_Exportacion.frx":0354
      FormatStyle(4)  =   "frmShowGuiasxFact_Exportacion.frx":0408
      FormatStyle(5)  =   "frmShowGuiasxFact_Exportacion.frx":04E0
      FormatStyle(6)  =   "frmShowGuiasxFact_Exportacion.frx":0598
      ImageCount      =   0
      PrinterProperties=   "frmShowGuiasxFact_Exportacion.frx":0678
   End
   Begin GridEX20.GridEX GridEX3 
      Height          =   2055
      Left            =   2850
      TabIndex        =   6
      Top             =   4350
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   3625
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ScrollToolTipColumn=   ""
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmShowGuiasxFact_Exportacion.frx":0850
      Column(2)       =   "frmShowGuiasxFact_Exportacion.frx":0918
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmShowGuiasxFact_Exportacion.frx":09BC
      FormatStyle(2)  =   "frmShowGuiasxFact_Exportacion.frx":0AF4
      FormatStyle(3)  =   "frmShowGuiasxFact_Exportacion.frx":0BA4
      FormatStyle(4)  =   "frmShowGuiasxFact_Exportacion.frx":0C58
      FormatStyle(5)  =   "frmShowGuiasxFact_Exportacion.frx":0D30
      FormatStyle(6)  =   "frmShowGuiasxFact_Exportacion.frx":0DE8
      ImageCount      =   0
      PrinterProperties=   "frmShowGuiasxFact_Exportacion.frx":0EC8
   End
   Begin GridEX20.GridEX GridEX2 
      Height          =   2055
      Left            =   90
      TabIndex        =   5
      Top             =   4350
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3625
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ScrollToolTipColumn=   ""
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmShowGuiasxFact_Exportacion.frx":10A0
      Column(2)       =   "frmShowGuiasxFact_Exportacion.frx":1168
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmShowGuiasxFact_Exportacion.frx":120C
      FormatStyle(2)  =   "frmShowGuiasxFact_Exportacion.frx":1344
      FormatStyle(3)  =   "frmShowGuiasxFact_Exportacion.frx":13F4
      FormatStyle(4)  =   "frmShowGuiasxFact_Exportacion.frx":14A8
      FormatStyle(5)  =   "frmShowGuiasxFact_Exportacion.frx":1580
      FormatStyle(6)  =   "frmShowGuiasxFact_Exportacion.frx":1638
      ImageCount      =   0
      PrinterProperties=   "frmShowGuiasxFact_Exportacion.frx":1718
   End
   Begin VB.Frame FraBuscar 
      Caption         =   "Argumentos de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   30
      TabIndex        =   7
      Top             =   30
      Width           =   14115
      Begin VB.CommandButton cmdBusCliente 
         Caption         =   "..."
         Height          =   285
         Left            =   8100
         TabIndex        =   79
         Tag             =   "..."
         Top             =   250
         Width           =   300
      End
      Begin VB.TextBox txtCod_TipoFact 
         Height          =   315
         Left            =   7485
         TabIndex        =   2
         Top             =   660
         Width           =   570
      End
      Begin VB.TextBox txtDes_TipoFact 
         Height          =   315
         Left            =   8085
         TabIndex        =   3
         Top             =   660
         Width           =   3075
      End
      Begin VB.TextBox txtNom_Cliente 
         Height          =   315
         Left            =   8445
         TabIndex        =   1
         Top             =   240
         Width           =   3075
      End
      Begin VB.TextBox txtAbr_Cliente 
         Height          =   315
         Left            =   7485
         TabIndex        =   0
         Top             =   240
         Width           =   570
      End
      Begin VB.CheckBox optTodos 
         Caption         =   "&Todos"
         Height          =   255
         Left            =   5610
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox Cbo_Almacen 
         Height          =   315
         Left            =   1410
         TabIndex        =   8
         Top             =   240
         Width           =   4080
      End
      Begin MSComCtl2.DTPicker dtpFecEmiIni 
         Height          =   315
         Left            =   1410
         TabIndex        =   27
         Top             =   675
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   95223809
         CurrentDate     =   37543
      End
      Begin MSComCtl2.DTPicker dtpFecEmiFin 
         Height          =   315
         Left            =   3450
         TabIndex        =   28
         Top             =   675
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   95223809
         CurrentDate     =   37543
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   570
         Left            =   12000
         TabIndex        =   4
         Top             =   285
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   1005
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~~~0~Verdadero~Falso~&Buscar~"
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1100
         ControlHeigth   =   550
         ControlSeparator=   50
      End
      Begin VB.Label Label10 
         Caption         =   "Tipo de Facturación"
         Height          =   420
         Left            =   6495
         TabIndex        =   52
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label Label9 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   6525
         TabIndex        =   36
         Top             =   285
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "Rango Fecha de Emisión:"
         Height          =   360
         Left            =   105
         TabIndex        =   30
         Top             =   645
         Width           =   1710
      End
      Begin VB.Label Label2 
         Caption         =   "Almacen"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   6150
      Left            =   0
      TabIndex        =   31
      Top             =   1215
      Width           =   14160
      _ExtentX        =   24977
      _ExtentY        =   10848
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmShowGuiasxFact_Exportacion.frx":18F0
      Column(2)       =   "frmShowGuiasxFact_Exportacion.frx":19B8
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmShowGuiasxFact_Exportacion.frx":1A5C
      FormatStyle(2)  =   "frmShowGuiasxFact_Exportacion.frx":1B94
      FormatStyle(3)  =   "frmShowGuiasxFact_Exportacion.frx":1C44
      FormatStyle(4)  =   "frmShowGuiasxFact_Exportacion.frx":1CF8
      FormatStyle(5)  =   "frmShowGuiasxFact_Exportacion.frx":1DD0
      FormatStyle(6)  =   "frmShowGuiasxFact_Exportacion.frx":1E88
      FormatStyle(7)  =   "frmShowGuiasxFact_Exportacion.frx":1F68
      FormatStyle(8)  =   "frmShowGuiasxFact_Exportacion.frx":2014
      ImageCount      =   0
      PrinterProperties=   "frmShowGuiasxFact_Exportacion.frx":20C4
   End
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   630
      Left            =   7575
      TabIndex        =   37
      Top             =   7515
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   1111
      Custom          =   $"frmShowGuiasxFact_Exportacion.frx":229C
      Orientacion     =   0
      Style           =   1
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1400
      ControlHeigth   =   600
      ControlSeparator=   50
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   6405
      Top             =   4935
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label lbRollos 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   8730
      TabIndex        =   34
      Top             =   6690
      Width           =   45
   End
   Begin VB.Label lbDes_Color 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   9690
      TabIndex        =   33
      Top             =   6690
      Width           =   45
   End
   Begin VB.Label lbGuia 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   9690
      TabIndex        =   32
      Top             =   6990
      Width           =   45
   End
End
Attribute VB_Name = "frmShowGuiasxFact_Prendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iRowAnterior As Long
Dim iColAnterior As Long
Dim bClickColSelec As Boolean
Dim bCargaGRid As Boolean
Dim bPuedeAutorizar  As Boolean
Dim sTipoDocAutorizar As String
Dim Doc As String
Dim strSQL As String
Public Codigo As String
Public Descripcion As String
Public TipoAdd As String
Dim sCod_TipoFact  As String

Dim sSer_Factura_Orig As String
Dim sNum_Factura_Orig As String
Dim Buscando As Integer
Private sww As Boolean

Private Sub DtFecVencimiento_Change()
  GridEX1.ClearFields
  dtpFecEmiIni.Value = ""
  dtpFecEmiFin.Value = ""
End Sub

Private Sub cmdAceptar_Click()
    GuardarDatos
End Sub

Private Sub cmdAceptarPrecio_Click()
    GrabarPrecio
End Sub

Private Sub cmdBusCliente_Click()
    Dim oTipo As New frmBusqGeneral
    Dim rs As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.SQuery = "SELECT Abr_Cliente as Código, nom_cliente as Descripción FROM TG_Cliente ORDER BY Abr_Cliente"
    oTipo.CARGAR_DATOS
    oTipo.Show 1
    If Codigo <> "" Then
        txtAbr_Cliente.Text = Codigo
        txtNom_Cliente.Text = Descripcion
        strSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
        txtAbr_Cliente.Tag = DevuelveCampo(strSQL, cCONNECT)
        
        SendKeys "{TAB}"
        Codigo = ""
    End If
    Set oTipo = Nothing
    Set rs = Nothing
End Sub

Private Sub cmdCancelarPrecio_Click()
    Me.fraPrecio.Visible = False
End Sub

Private Sub cmdLugEnt_Click()
    Load frmMantLugaresEntrega
    frmMantLugaresEntrega.sCod_Cliente = Me.txtAbr_Cliente.Tag
    frmMantLugaresEntrega.CARGA_GRID
    frmMantLugaresEntrega.Show vbModal
    Set frmMantLugaresEntrega = Nothing
End Sub

Private Sub Command1_Click()
    Me.fraDatosAdicionales.Visible = False
End Sub

Private Sub dtpFecEmiIni_Change()
  GridEX1.ClearFields
  'dtpFecEmiFin.Value = dtpFecEmiIni
End Sub

Private Sub Form_Load()
  sww = False
  dtpFecEmiIni.Value = Date
  'dtpFecEmiIni.Value = ""
  
  dtpFecEmiFin.Value = Date
  'dtpFecEmiFin.Value = ""
  
  FillAlmacen
  
'  FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name) & "/SALIR"
  
  iLanguage = CInt(GetSetting("Visuales", "Settings", "Language", "1"))
  
  If InStr(FunctButt1.FunctionsUser, "AUTORIZARPAGO") <> 0 Then
      bPuedeAutorizar = True
  End If
  
  Set GridEX2.ADORecordset = CargarRecordSetDesconectado("select Cod_CondVent,Des_CondVent as Descripcion from lg_condvent", cCONNECT)
    
  GridEX2.ColumnAutoResize = True
'  GridEX2.ClearFields
'  GridEX2.Rebind
  
  'GridEX2 will act as the drop down list
  'for column 'SupplierID' in GridEX1
  
  GridEX2.ActAsDropDown = True
  GridEX2.BoundColumnIndex = 1
  GridEX2.ReplaceColumnIndex = 2
   
  
  GridEX2.Columns("Cod_CondVent").Visible = False
  
  Set GridEX3.ADORecordset = CargarRecordSetDesconectado("select Cod_Moneda as cod_Moneda,Nom_Moneda as Descripcion from tg_moneda", cCONNECT)
    
  GridEX3.ColumnAutoResize = True

  GridEX3.ActAsDropDown = True
  GridEX3.BoundColumnIndex = 1
  GridEX3.ReplaceColumnIndex = 2
  
  GridEX3.Columns("Cod_Moneda").Visible = False
  
  GridEX4.ActAsDropDown = True
  GridEX4.BoundColumnIndex = 2
  GridEX4.ReplaceColumnIndex = 2
  
  Busca_AnexosCliente
End Sub

Private Sub Buscar()

On Error GoTo drDepurar

Dim ssql As String
Dim oGroup As GridEX20.JSGroup
Dim oFormat As JSFormatStyle

If txtCod_TipoFact = "" Then
    Aviso "Seleccione Tipo de Facturación", 2
End If

    sCod_TipoFact = txtCod_TipoFact.Text


If Left(Cbo_Almacen, 2) = "62" Then
  ssql = "Ventas_Muestra_Documentos_Pendientes_Facturar_Prendas '" & Left(Cbo_Almacen, 3) & "','" & dtpFecEmiIni & "','" & dtpFecEmiFin & "','" & IIf(optTodos, "*", "") & "','" & sCod_TipoFact & "','" & txtAbr_Cliente.Tag & "' ,'" & vusu & "'"
Else
  Exit Sub
End If

GridEX1.ClearFields

GridEX1.DefaultGroupMode = jgexDGMExpanded
bCargaGRid = False
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(ssql, cCONNECT)
  
Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Fac_Cli").Index, jgexSortAscending)

MuestraSubTotales
GridEX1.BackColorRowGroup = &H80000005

GridEX1.ColumnHeaderHeight = 500

GridEX1.Columns("SEL").ColumnType = jgexCheckBox
GridEX1.Columns("SEL").Visible = True
GridEX1.Columns("SEL").EditType = jgexEditCheckBox
GridEX1.Columns("SEL").Width = 500

GridEX1.Columns("Fecha").Width = 900
GridEX1.Columns("Ser_Factura").Width = 400
GridEX1.Columns("Num_Factura").Width = 1100
GridEX1.Columns("Cod_Cliente").Visible = False
GridEX1.Columns("nom_cliente").Width = 500
GridEX1.Columns("nro_Guia").Width = 1260
GridEX1.Columns("Ser_ParteSalida").Visible = False
GridEX1.Columns("Num_ParteSalida").Visible = False
GridEX1.Columns("Num_Packing").Width = 500
GridEX1.Columns("Cod_PurOrd").Width = 1000
GridEX1.Columns("Cod_PurOrd").Visible = False
GridEX1.Columns("Cod_PurOrd_Factura").Width = 1100

GridEX1.Columns("Cod_LotPurOrd").Visible = True
GridEX1.Columns("Cod_LotPurOrd").Width = 500
GridEX1.Columns("Cod_EstCli").Width = 1300
GridEX1.Columns("Cod_ColCli").Width = 680
GridEX1.Columns("Nom_ColCli").Width = 700
GridEX1.Columns("Cod_Talla").Width = 500
GridEX1.Columns("EstCli").Visible = False
GridEX1.Columns("Num_Prendas").Width = 580
GridEX1.Columns("Pre_Unitario").Width = 550
GridEX1.Columns("Imp_Comision").Width = 550
GridEX1.Columns("Imp_Comision").Caption = "Comisión"
GridEX1.Columns("Pre_Unitario_Org").Visible = False
GridEX1.Columns("MontoDespacho").Width = 845
GridEX1.Columns("MismoPrecio").Width = 300
GridEX1.Columns("Moneda").Width = 400
GridEX1.Columns("Cod_Moneda").Visible = False
GridEX1.Columns("Sel").Width = 390
GridEX1.Columns("Fac_Cli").Width = 1110
GridEX1.Columns("Gastos_Financieros").Width = 585
GridEX1.Columns("Otros").Width = 615
GridEX1.Columns("Motivo").Width = 1500
GridEX1.Columns("Observaciones").Width = 1500
GridEX1.Columns("cod_almacen").Width = 375
GridEX1.Columns("num_movstk").Width = 750
GridEX1.Columns("Num_Secuencia").Visible = False
GridEX1.Columns("Cod_CondVent").Width = 825
GridEX1.Columns("Condicion_Venta").Width = 705
GridEX1.Columns("COD_ANXO").Visible = False
GridEX1.Columns("DES_ANEXO").Width = 600
GridEX1.Columns("COD_condvent").Visible = False
GridEX1.Columns("COD_TIPANEX").Visible = False
GridEX1.Columns("DatosAdic").Width = 400

GridEX1.Columns("Fecha").Caption = "Fecha"
GridEX1.Columns("Ser_Factura").Caption = "Ser/Fact"
GridEX1.Columns("Num_Factura").Caption = "N/Fact"
GridEX1.Columns("Cod_Cliente").Caption = "Cliente"
GridEX1.Columns("nom_cliente").Caption = "Cliente"
GridEX1.Columns("nro_Guia").Caption = "NroGuia"
GridEX1.Columns("Ser_ParteSalida").Caption = "Ser/Parte"
GridEX1.Columns("Num_ParteSalida").Caption = "Num/Parte"
GridEX1.Columns("Cod_PurOrd").Caption = "Pur.Order"
GridEX1.Columns("Cod_PurOrd_Factura").Caption = "Pur.Order a Facturar"
GridEX1.Columns("Cod_LotPurOrd").Caption = "LotPurOrd"
GridEX1.Columns("Cod_EstCli").Caption = "EstCli"
GridEX1.Columns("Cod_ColCli").Caption = "Color"
GridEX1.Columns("Nom_ColCli").Caption = "Nombre Color"
GridEX1.Columns("Cod_Talla").Caption = "Talla"
GridEX1.Columns("EstCli").Caption = "EstCli"
GridEX1.Columns("Num_Prendas").Caption = "Prendas"
GridEX1.Columns("Pre_Unitario").Caption = "Precio Unitario"
GridEX1.Columns("MontoDespacho").Caption = "MontoDespacho"
GridEX1.Columns("MismoPrecio").Caption = "MismoPrecio"
GridEX1.Columns("Moneda").Caption = "Moneda"
GridEX1.Columns("Cod_Moneda").Caption = "Moneda"
GridEX1.Columns("Sel").Caption = "Sel"
GridEX1.Columns("Fac_Cli").Caption = "Fac_Cli"
GridEX1.Columns("Gastos_Financieros").Caption = "Gastos_Financieros"
GridEX1.Columns("Otros").Caption = "Otros"
GridEX1.Columns("Motivo").Caption = "Motivo"
GridEX1.Columns("Observaciones").Caption = "Observaciones"
GridEX1.Columns("cod_almacen").Caption = "almacen"
GridEX1.Columns("num_movstk").Caption = "Nro.Movstk"
GridEX1.Columns("Num_Secuencia").Caption = "Secuencia"
GridEX1.Columns("Cod_CondVent").Caption = "Cond.Vent"
GridEX1.Columns("Condicion_Venta").Caption = "Condicion.Venta"
GridEX1.Columns("DES_ANEXO").Caption = "Anexo"



With GridEX1.Columns("Condicion_Venta")
  .TextAlignment = jgexAlignLeft
  .EditType = jgexEditCombo
  Set .DropDownControl = GridEX2
End With

With GridEX1.Columns("moneda")
  .TextAlignment = jgexAlignLeft
  .EditType = jgexEditCombo
  Set .DropDownControl = GridEX3
End With

With GridEX1.Columns("Des_Anexo")
  .TextAlignment = jgexAlignLeft
  .EditType = jgexEditCombo
  Set .DropDownControl = GridEX4
End With

With GridEX1.Columns("Fecha")
  .EditType = jgexEditCalendarDropDown
End With

SetColores

GridEX1.DefaultGroupMode = jgexDGMCollapsed

If dtpFecEmiIni.Value <> "" Then
    GridEX1.DefaultGroupMode = jgexDGMExpanded
End If

If GridEX1.RowCount > 0 Then
    GridEX1.Row = 1
End If

GridEX1.ContinuousScroll = True

Buscando = 0

Exit Sub
Resume
drDepurar:
  errores err.Number
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Msg As Variant
    Select Case ActionName
    Case "BUSCAR"
        Buscando = 1
        Buscar
    End Select
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Msg As Variant
    Select Case ActionName
    Case "ESTILOSCLIENTE"
        EjecutaOpcionDLL oTablasEst, "ActEstCli", vper, vemp, Me, True
    Case "AUTORIZARPAGO"
        If GridEX1.RowCount = 0 Then Exit Sub
        Msg = MsgBox("¿Esta seguro de autorizar pago?", vbYesNo)
        If Msg = vbNo Then Exit Sub
        Autorizar
    Case "IMPRIMIR"
        If GridEX1.RowCount = 0 Then Exit Sub
            Imprimir
    Case "SALIR"
       Unload Me
    End Select
End Sub

Private Sub Imprimir()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String


    Ruta = vRuta & "\Rpt_Exportacion_PRECIOS.XLT"
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "reporte", GridEX1.ADORecordset
    
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler err, "GeneraReportes"
    Set oo = Nothing
End Sub

Private Sub GridEX1_AfterColEdit(ByVal ColIndex As Integer)
  If Left(Cbo_Almacen, 2) = "62" Then
        AfterColEdit_Prendas (ColIndex)
  End If
End Sub

'''jjj
Sub AfterColEdit_Prendas(ByVal ColIndex As Integer)
    Dim ssql As String
    On Error GoTo Error_Handler

    Dim oGroup As GridEX20.JSGroup


    Select Case ColIndex
        Case Is = GridEX1.Columns("Sel").Index
              ssql = "Ventas_Cambio_Estado_DocAlm_Prendas '$','$','$','$','$',$,'$',$,$,'$','$','$' ,'$','$','$','$',$,$,$,'$',$,'$','$','$','$','$','$','$','$','$',$,'$','$',$,$"
              ssql = VBsprintf(ssql, Left(Cbo_Almacen, 2), _
                               GridEX1.Value(GridEX1.Columns("num_movstk").Index), _
                               GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
                               GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
                               GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), _
                               GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
                               GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
                               GridEX1.Value(GridEX1.Columns("Gastos_Financieros").Index), _
                               GridEX1.Value(GridEX1.Columns("Otros").Index), sCod_TipoFact, _
                               GridEX1.Value(GridEX1.Columns("cod_tipanex").Index), _
                               GridEX1.Value(GridEX1.Columns("Cod_Anxo").Index), _
                               GridEX1.Value(GridEX1.Columns("Observaciones_Factura").Index), _
                               GridEX1.Value(GridEX1.Columns("Cod_LugEnt").Index), _
                               FixNulos(GridEX1.Value(GridEX1.Columns("Num_CartaCredito").Index), vbString), _
                               GridEX1.Value(GridEX1.Columns("Cod_cliente").Index), _
                               GridEX1.Value(GridEX1.Columns("Imp_Flete").Index), _
                               GridEX1.Value(GridEX1.Columns("Imp_Seguro").Index), GridEX1.Value(GridEX1.Columns("Imp_DESCUENTO").Index), GridEX1.Value(GridEX1.Columns("Cod_Termino_Venta").Index), GridEX1.Value(GridEX1.Columns("Num_Packing").Index), _
                               GridEX1.Value(GridEX1.Columns("cod_Embarque").Index), _
                               GridEX1.Value(GridEX1.Columns("Nom_Embarque").Index), _
                               GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index), _
                               GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index), IIf(GridEX1.Value(GridEX1.Columns("Sel").Index) = 0, "P", "A"), GridEX1.Value(GridEX1.Columns("COD_ESTCLI").Index), GridEX1.Value(GridEX1.Columns("Fecha").Index), GridEX1.Value(GridEX1.Columns("Cod_Vendor").Index), GridEX1.Value(GridEX1.Columns("Cod_Class").Index), GridEX1.Value(GridEX1.Columns("Num_Embarque").Index), vusu, GridEX1.Value(GridEX1.Columns("Por_Comision").Index), GridEX1.Value(GridEX1.Columns("imp_Desaduanaje").Index), GridEX1.Value(GridEX1.Columns("Imp_Transporte_Pais_Destino").Index))
        
                                   
                ExecuteCommandSQL cCONNECT, ssql
                SeleccionarOtrosReg GridEX1.Value(GridEX1.Columns("Sel").Index)
                
'          Case Is = GridEX1.Columns("Pre_Unitario").Index
'                GrabarPrecio
'                GridEX1.Value(GridEX1.Columns("sel").Index) = 0
'          Case Is = GridEX1.Columns("Num_Prendas").Index
'                GridEX1.Value(GridEX1.Columns("Monto Despacho").Index) = GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index) * GridEX1.Value(GridEX1.Columns("Kgs_a_Facturar").Index)
'                GridEX1.Value(GridEX1.Columns("sel").Index) = 0
          Case Is = GridEX1.Columns("Ser_Factura").Index
                GridEX1.Value(GridEX1.Columns("Fac_Cli").Index) = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) & "-" & RPad(GridEX1.Value(GridEX1.Columns("Num_Factura").Index), 13, " ") & "  " & GridEX1.Value(GridEX1.Columns("Nom_Cliente").Index)
                GridEX1.Groups.Clear
                Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Fac_Cli").Index, jgexSortAscending)
                GridEX1.Value(GridEX1.Columns("sel").Index) = 0
                
          Case Is = GridEX1.Columns("Num_Factura").Index
          
'                If GrabaDatosParaFacturaCambiada Then
'                    GridEX1.Value(GridEX1.Columns("Fac_Cli").Index) = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) & "-" & RPad(GridEX1.Value(GridEX1.Columns("Num_Factura").Index), 13, " ") & "  " & GridEX1.Value(GridEX1.Columns("Nom_Cliente").Index)
'                    GridEX1.Groups.Clear
'                    Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Fac_Cli").Index, jgexSortAscending)
'                    GridEX1.Value(GridEX1.Columns("sel").Index) = 0
'                End If
                
          'add para asignar factura por packing
          
          Call facturaxPacking(Trim(GridEX1.Value(GridEX1.Columns("num_factura").Index)), Trim(GridEX1.Value(GridEX1.Columns("num_packing").Index)))
             
             
          Case Is = GridEX1.Columns("Gastos_Financieros").Index
                Cambio_Importe "Gastos_Financieros"
                GridEX1.Value(GridEX1.Columns("sel").Index) = 0
                
'          Case Is = GridEX1.Columns("Otros").Index
'                Cambio_Importe "Otros"
'                GridEX1.Value(GridEX1.Columns("sel").Index) = 0
                
'          Case Is = GridEX1.Columns("DatosAdic").Index
          
          Case Is = GridEX1.Columns("Fecha").Index
                Cambio_Fecha GridEX1.Value(GridEX1.Columns("Fecha").Index)
                
          Case Is = GridEX1.Columns("Cod_PurOrd_Factura").Index
                Cambio_PO_Factura GridEX1.Value(GridEX1.Columns("Cod_PurOrd_Factura").Index)
      End Select
    Exit Sub
Resume

Error_Handler:
  errores err.Number
  If ColIndex = GridEX1.Columns("Sel").Index Then GridEX1.Value(GridEX1.Columns("sel").Index) = 0
End Sub
Private Sub facturaxPacking(Factura As String, num_packing As String)
Dim rs As ADODB.Recordset
Dim I As Integer
Dim filas As Integer
Dim ssql As String
Dim num_factura As String
Dim grupo As String
Dim oGroup As GridEX20.JSGroup
Dim serie As String, Nro_Factura As String, iPos As Integer, lvSw As Boolean
 
On Error GoTo errx

filas = GridEX1.RowCount
num_factura = "00000000"
num_factura = num_factura + Factura 'Trim(GridEX1.Value(GridEX1.Columns("Num_Factura").Index))
num_factura = Right(num_factura, 8)
    
Set rs = GridEX1.ADORecordset

rs.MoveFirst
Do While Not rs.EOF
 If Trim(rs.Fields("num_packing").Value) = Trim(num_packing) And Trim(rs.Fields("NUM_FACTURA").Value) = Trim(num_factura) Then
     Exit Sub
     Set rs = Nothing

 End If
 rs.MoveNext
Loop

GridEX1.Redraw = False
lvSw = True
  
I = 1
GridEX1.MoveFirst

Do While I <= filas
    
    If GridEX1.Value(GridEX1.Columns("num_packing").Index) = num_packing And Replace(num_factura, "0", " ") <> " " Then
    
    If lvSw Then iPos = GridEX1.Row
     lvSw = False
     

        
    ssql = "UP_MAN_TEMP_Ventas_NUEVO_DATO '$','$','$','$','$','$',$,'$','$','$','$','$','$','$',$,$,$,$"
            
    'GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _

           ssql = VBsprintf(ssql, vusu, Left(Cbo_Almacen, 2), _
            sSer_Factura_Orig, sNum_Factura_Orig, _
            GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
            num_factura, _
            GridEX1.Value(GridEX1.Columns("Num_Packing").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_cliente").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_PurOrd").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_LotPurOrd").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_Estcli").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_ColCli").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_Talla").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_PurOrd_Factura").Index), _
            GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
            GridEX1.Value(GridEX1.Columns("Imp_Descuento").Index), _
            GridEX1.Value(GridEX1.Columns("Pre_Unitario_Org").Index), _
            GridEX1.Value(GridEX1.Columns("Imp_Comision").Index))
            ExecuteCommandSQL cCONNECT, ssql
    
            'If GrabaDatosParaFacturaCambiada Then
            'GridEX1.Value(GridEX1.Columns("Fac_Cli").Index) = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) & "-" & RPad(GridEX1.Value(GridEX1.Columns("Num_Factura").Index), 13, " ") & "  " & GridEX1.Value(GridEX1.Columns("Nom_Cliente").Index)
            GridEX1.Value(GridEX1.Columns("Fac_Cli").Index) = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) & "-" & RPad(num_factura, 13, " ") & "  " & GridEX1.Value(GridEX1.Columns("Nom_Cliente").Index)
            GridEX1.Value(GridEX1.Columns("NUM_FACTURA").Index) = num_factura
            
            GridEX1.Groups.Clear
            Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Fac_Cli").Index, jgexSortAscending)
            'GridEX1.Value(GridEX1.Columns("sel").Index) = 0
            'End If
            
    End If
GridEX1.MoveNext

I = I + 1
If I = filas Then
    Exit Do
End If

Loop

GridEX1.Row = iPos
GridEX1.Redraw = True

Exit Sub
errx:
    errores err.Number
End Sub

Private Sub GridEX1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)

If Left(Cbo_Almacen, 2) = "62" Then
  Select Case ColIndex
    Case Is = GridEX1.Columns("Ser_Factura").Index
        sSer_Factura_Orig = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
        sNum_Factura_Orig = RPad(GridEX1.Value(GridEX1.Columns("Num_Factura").Index), 13, " ")
        Cancel = False
    Case Is = GridEX1.Columns("Num_Factura").Index
        sSer_Factura_Orig = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
        sNum_Factura_Orig = RPad(GridEX1.Value(GridEX1.Columns("Num_Factura").Index), 13, " ")
        Cancel = False
    Case Is = GridEX1.Columns("SEL").Index
      Cancel = False
    'Case Is = GridEX1.Columns("Pre_Unitario").Index
     ' Cancel = False
'      CargarPrecio
    Case Is = GridEX1.Columns("Condicion_Venta").Index
      Cancel = False
    Case Is = GridEX1.Columns("Moneda").Index
      Cancel = False
   Case Is = GridEX1.Columns("Gastos_Financieros").Index
      Cancel = False
   Case Is = GridEX1.Columns("Des_Anexo").Index
      Cancel = False
      Busca_AnexosCliente
      
   Case Is = GridEX1.Columns("DatosAdic").Index 'llll
      Cancel = False
      
      CargarDatos
      
   Case Is = GridEX1.Columns("Fecha").Index
      Cancel = False
   Case Is = GridEX1.Columns("Cod_PurOrd_Factura").Index
      Cancel = False
   Case Else
      Cancel = True
    End Select
End If
  
End Sub

Private Sub GridEX1_Click()
    'On Error Resume Next
    If GridEX1.RowCount = 0 Then Exit Sub
    If GridEX1.IsGroupItem(GridEX1.Row) = True Then Exit Sub
    If Not (GridEX1.Col) > 0 Then Exit Sub
    
    Dim ColIndex As Long
    Dim oRowData As JSRowData
    Dim SGRUPO As String
    Dim iRow As Long
    Dim I As Long
    Dim sCaptionGroup As String
    
    bCargaGRid = True
    
    ColIndex = GridEX1.Col
    If Trim(UCase(GridEX1.Columns(ColIndex).Key)) = "SEL" Then
        bClickColSelec = True
        SendKeys "{ENTER}"
    End If
End Sub

Private Sub GridEX1_DblClick()
    Dim I As Integer
    For I = 1 To GridEX1.Columns.Count
        Debug.Print GridEX1.Name & ".Columns(" & Chr(34) & GridEX1.Columns(I).Key & Chr(34) & ").width = " & CStr(GridEX1.Columns(I).Width)
    Next
    
    For I = 1 To GridEX1.Columns.Count
        Debug.Print GridEX1.Name & ".COLUMNS(" & Chr(34) & GridEX1.Columns(I).Key & Chr(34) & ").CAPTION = " & CStr(GridEX1.Columns(I).Caption)
    Next
    
End Sub

Private Sub GridEX1_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    Dim ocol As JSColumn
    Dim oRow As JSRowData
    Dim vCurrentRow As Variant
    Dim oRowGroup As JSRowData
    Dim sProveedor As String
    
    iColAnterior = LastCol
    iRowAnterior = LastRow
    
    If GridEX1.Row <> 0 Then
        Set oRow = GridEX1.GetRowData(GridEX1.Row)
    End If
      
    If GridEX1.RowCount > 0 Then
      On Error Resume Next
      'lbDesTela.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Tela").Index)), "", GridEX1.Value(GridEX1.Columns("Tela").Index))
      'lbComb.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Comb").Index)), "", GridEX1.Value(GridEX1.Columns("Comb").Index))
      'lbCalidad.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Calidad").Index)), "", GridEX1.Value(GridEX1.Columns("Calidad").Index))
      'lbRollos.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Numero_Rollos").Index)), "", GridEX1.Value(GridEX1.Columns("Numero_Rollos").Index))
      'If lbCod_Color.Visible Then lbDes_Color.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Color").Index)), "", GridEX1.Value(GridEX1.Columns("Color").Index))
      'lbGuia.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("nro_Guia").Index)), "", GridEX1.Value(GridEX1.Columns("nro_Guia").Index))
      'lbObservacion.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Observaciones").Index)), "", GridEX1.Value(GridEX1.Columns("Observaciones").Index))
    End If
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)

Dim strGroupCaption As String

If GridEX1.RowCount = 0 Then Exit Sub

If RowBuffer.RowType = jgexRowTypeGroupHeader Then
    strGroupCaption = RTrim(RowBuffer.GroupCaption) & " (" & RowBuffer.RecordCount & " Documentos " & "" & ") "
    RowBuffer.GroupCaption = strGroupCaption
End If


Dim fmtConDIA_Programado As JSFmtCondition
If Buscando = 1 Then
    Set fmtConDIA_Programado = GridEX1.FmtConditions.Add(GridEX1.Columns("MONTODESPACHO").Index, jgexEqual, 0)
    
    With fmtConDIA_Programado.FormatStyle
        .ForeColor = &H8000&
        .FontSize = 8
        .BackColor = &H80000018 'vbYellow
    End With
End If

End Sub

Private Sub MuestraSubTotales()
Dim colTemp As JSColumn

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("Moneda")
colTemp.AggregateFunction = jgexAggregateNone
colTemp.TotalRowPrefix = "SUB TOTAL "

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("Num_Prendas")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("MontoDespacho")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

End Sub

Private Sub SetColores()

Dim fmtCon As JSFmtCondition
Dim fmtCond2 As JSFmtCondition
Dim fmtCond3 As JSFmtCondition

Set fmtCon = GridEX1.FmtConditions.Add(GridEX1.Columns("SEL").Index, jgexEqual, -1)
    
    With GridEX1.FmtConditions
            .ApplyGroupCondition = True
            .ShowGroupConditionCount = True
            .GroupConditionCountTitle = "Documento(s) Autorizado(s)"
            Set fmtCon = .GroupCondition
    End With
    fmtCon.SetCondition GridEX1.Columns("SEL").Index, jgexEqual, -1
    fmtCon.FormatStyle.FontBold = True
    fmtCon.FormatStyle.BackColor = &HFFFFC0   '&HC0FFC0    ' &HC0E0FF    ' '&HC0FFFF
    
End Sub


Private Sub Autorizar()

On Error GoTo errorx
Dim ssql As String
Dim aMess(4), I As Integer


GridEX1.MoveFirst

For I = 0 To GridEX1.RowCount

  If GridEX1.Value(GridEX1.Columns("SEL").Index) Then
  
    If Left(Cbo_Almacen, 2) = "62" Then
  
      ssql = "Ventas_Cambio_Estado_DocAlm_Prendas '$','$','$','$','$',$,'$',$,$ ,'$','$','$','$','$','$','$',$,$,$,'$',$,'$','$','$','$','$','$','$','$','$',$,'$','$',$,$"
            
      ssql = VBsprintf(ssql, Left(Cbo_Almacen, 2), _
                       GridEX1.Value(GridEX1.Columns("num_movstk").Index), _
                       GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), _
                       GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
                       GridEX1.Value(GridEX1.Columns("Gastos_Financieros").Index), _
                       GridEX1.Value(GridEX1.Columns("Otros").Index), sCod_TipoFact, _
                       GridEX1.Value(GridEX1.Columns("cod_tipanex").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Anxo").Index), _
                       GridEX1.Value(GridEX1.Columns("Observaciones_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_LugEnt").Index), _
                       FixNulos(GridEX1.Value(GridEX1.Columns("Num_CartaCredito").Index), vbString), _
                       GridEX1.Value(GridEX1.Columns("Cod_cliente").Index), _
                       GridEX1.Value(GridEX1.Columns("Imp_Flete").Index), _
                       GridEX1.Value(GridEX1.Columns("Imp_Seguro").Index), GridEX1.Value(GridEX1.Columns("Imp_DESCUENTO").Index), GridEX1.Value(GridEX1.Columns("Cod_Termino_Venta").Index), GridEX1.Value(GridEX1.Columns("Num_Packing").Index), _
                       GridEX1.Value(GridEX1.Columns("cod_Embarque").Index), _
                       GridEX1.Value(GridEX1.Columns("Nom_Embarque").Index), _
                       GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index), _
                       GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index), IIf(GridEX1.Value(GridEX1.Columns("Sel").Index) = 0, "P", "A"), GridEX1.Value(GridEX1.Columns("COD_ESTCLI").Index), GridEX1.Value(GridEX1.Columns("Fecha").Index), GridEX1.Value(GridEX1.Columns("Cod_Vendor").Index), GridEX1.Value(GridEX1.Columns("Cod_Class").Index), GridEX1.Value(GridEX1.Columns("Num_Embarque").Index), vusu, GridEX1.Value(GridEX1.Columns("Por_Comision").Index), GridEX1.Value(GridEX1.Columns("imp_Desaduanaje").Index), GridEX1.Value(GridEX1.Columns("Imp_Transporte_Pais_Destino").Index))



      ExecuteCommandSQL cCONNECT, ssql
    End If
  End If

  GridEX1.MoveNext

Next I

If Left(Cbo_Almacen, 2) = "62" Then
  ExecuteCommandSQL cCONNECT, "Ventas_Genera_Docum_Autorizados_Prendas '" & vusu & "','" & Left(Cbo_Almacen, 2) & "'"
End If

Mensaje kMESSAGE_INF_PROCESS_SATISFACTO

Buscar
 
Exit Sub
Resume
errorx:
    errores err.Number
End Sub

Sub Cambio_Nro_Factura()

Dim serie As String, Nro_Factura As String, iPos, I As Integer, lvSw As Boolean

  GridEX1.Redraw = False

  lvSw = True
  
  Doc = GridEX1.Value(GridEX1.Columns("Cod_Doc").Index)
  serie = GridEX1.Value(GridEX1.Columns("Ser_Docum").Index)
  Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Docum_Ventas").Index)
  
  GridEX1.MoveFirst
  For I = 0 To GridEX1.RowCount
    If Doc = GridEX1.Value(GridEX1.Columns("Cod_Doc").Index) Then
      If lvSw Then iPos = GridEX1.Row
      lvSw = False
      GridEX1.Value(GridEX1.Columns("Ser_Docum").Index) = serie
      GridEX1.Value(GridEX1.Columns("Nro_Docum_Ventas").Index) = Nro_Factura
    End If
    GridEX1.MoveNext
  Next I
  
  GridEX1.Row = iPos
  
  GridEX1.Redraw = True
  
  SendKeys "{TAB}"
  
End Sub


Sub Cambio_Importe(Campo As String)
    Dim Fac_Cli As String, Importe As String, iPos, I As Integer, lvSw As Boolean

    GridEX1.Redraw = False
    lvSw = True
  
    Fac_Cli = GridEX1.Value(GridEX1.Columns("Fac_Cli").Index)
    Importe = GridEX1.Value(GridEX1.Columns(Campo).Index)
  
    GridEX1.MoveFirst
    For I = 0 To GridEX1.RowCount
      If Fac_Cli = GridEX1.Value(GridEX1.Columns("Fac_Cli").Index) Then
            If lvSw Then iPos = GridEX1.Row
            lvSw = False
            GridEX1.Value(GridEX1.Columns(Campo).Index) = Importe
      End If
      GridEX1.MoveNext
    Next I
    
    GridEX1.Row = iPos
    GridEX1.Redraw = True
End Sub

Private Sub GridEX2_Click()

Dim serie As String, Nro_Factura As String, iPos, I As Integer, lvSw As Boolean

  GridEX1.Redraw = False

  lvSw = True
  
  serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
  Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index)
  
  
  GridEX1.MoveFirst
  For I = 0 To GridEX1.RowCount
    If serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) And Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index) Then
      If lvSw Then iPos = GridEX1.Row
      lvSw = False
      GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index) = GridEX2.Value(GridEX2.Columns("Cod_CondVent").Index)
      GridEX1.Value(GridEX1.Columns("Condicion_Venta").Index) = GridEX2.Value(GridEX2.Columns("Descripcion").Index)
    End If
    GridEX1.MoveNext
  Next I
  
  GridEX1.Row = iPos
  
  GridEX1.Redraw = True
  
  SendKeys "{TAB}"
  
End Sub

Private Sub GridEX3_Click()

Dim serie As String, Nro_Factura As String, iPos, I As Integer, lvSw As Boolean

  GridEX1.Redraw = False
  
  serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
  Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index)
  lvSw = True
  GridEX1.MoveFirst
  For I = 0 To GridEX1.RowCount
    If serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) And Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index) Then
      If lvSw Then iPos = GridEX1.Row
      lvSw = False
      GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index) = GridEX3.Value(GridEX3.Columns("Cod_Moneda").Index)
      GridEX1.Value(GridEX1.Columns("Moneda").Index) = GridEX3.Value(GridEX3.Columns("Descripcion").Index)
    End If
    GridEX1.MoveNext
  Next I
  
  GridEX1.Row = iPos
  
  GridEX1.Redraw = True
  
  SendKeys "{TAB}"
  
End Sub


Private Sub FillAlmacen()

Dim rstAux As ADODB.Recordset
Dim strSQL As String
    
strSQL = "Ventas_Ayuda_Almacenes_Confecciones"
         
Set rstAux = CargarRecordSetDesconectado(strSQL, cCONNECT)
Cbo_Almacen.Clear
With rstAux
    If .RecordCount > 0 Then .MoveFirst
    Do Until .EOF
        Cbo_Almacen.AddItem !Cod_Almacen & " " & !Nom_Almacen
        .MoveNext
    Loop
    .Close
End With
If Cbo_Almacen.ListCount > 0 Then Cbo_Almacen.ListIndex = 0
Set rstAux = Nothing
    
End Sub



Private Sub GridEX4_Click()

Dim serie As String, Nro_Factura As String, iPos, I As Integer, lvSw As Boolean

  GridEX1.Redraw = False

  lvSw = True
  
  serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
  Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index)
  
  
  GridEX1.MoveFirst
  For I = 0 To GridEX1.RowCount
    If serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) And Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index) Then
    
      If lvSw Then iPos = GridEX1.Row
      
        lvSw = False
        GridEX1.Value(GridEX1.Columns("Cod_Anxo").Index) = GridEX4.Value(GridEX4.Columns("Cod_Anxo").Index)
        GridEX1.Value(GridEX1.Columns("Des_Anexo").Index) = GridEX4.Value(GridEX4.Columns("Des_Anexo").Index)
      
      If RTrim(FixNulos(GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index), vbString)) = "" Then
        GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index) = GridEX4.Value(GridEX4.Columns("Pie_Factura1").Index)
      End If
      If RTrim(FixNulos(GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index), vbString)) = "" Then
        GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index) = GridEX4.Value(GridEX4.Columns("Pie_Factura2").Index)
      End If
      
    End If
    GridEX1.MoveNext
  Next I
  
  GridEX1.Row = iPos
  
  GridEX1.Redraw = True
  
  
  SendKeys "{TAB}"
  
End Sub

Public Sub BuscaCliente(Opcion As String)
Dim rstAux As ADODB.Recordset

    strSQL = "SELECT Cod_Cliente, Abr_Cliente, Nom_Cliente FROM TG_CLIENTE WHERE "
    
    txtAbr_Cliente = Trim(txtAbr_Cliente)
    txtNom_Cliente = Trim(txtNom_Cliente)
    
    Select Case Opcion
    Case 1: strSQL = strSQL & "Abr_Cliente LIKE '%" & txtAbr_Cliente & "%'"
    Case 2: strSQL = strSQL & "Nom_Cliente LIKE '%" & txtNom_Cliente & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
    
    
    frmBusqGeneral3.gexLista.Columns("Cod_Cliente").Visible = False
    frmBusqGeneral3.gexLista.Columns("Abr_Cliente").Width = 570
    frmBusqGeneral3.gexLista.Columns("Nom_Cliente").Width = 2370
    
    frmBusqGeneral3.gexLista.Columns("Abr_Cliente").Caption = "Abrev."
    frmBusqGeneral3.gexLista.Columns("Nom_Cliente").Caption = "Cliente"
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtAbr_Cliente.Tag = ""
    txtAbr_Cliente = ""
    txtNom_Cliente = ""
    If Codigo <> "" Then
        
        txtAbr_Cliente = Descripcion
        txtNom_Cliente = TipoAdd
        txtAbr_Cliente.Tag = Codigo
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
        
    Codigo = ""
    Descripcion = ""
End Sub


Private Sub Text1_Change()

End Sub

Private Sub txtCod_Class_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtCod_Embarque_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaModoTransporte 1
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCod_Vendor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtImp_comision_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdAceptarPrecio.SetFocus
    End If
End Sub

Private Sub txtNom_embarque_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtAbr_Cliente_Change()
        txtAbr_Cliente.Tag = ""
    
End Sub

Private Sub TxtAbr_Cliente_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        KeyAscii = 0
'        BuscaCliente 1
'        SendKeys "{TAB}"
'    End If
    If KeyAscii = 13 Then
        If Trim(txtAbr_Cliente.Text) = "" Then
            cmdBusCliente_Click
        Else
            strSQL = "SELECT Nom_Cliente FROM TG_CLIENTE WHERE Abr_Cliente LIKE '" & Trim(txtAbr_Cliente.Text) & "%'"
            txtNom_Cliente.Text = DevuelveCampo(strSQL, cCONNECT)
            strSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
            txtAbr_Cliente.Tag = DevuelveCampo(strSQL, cCONNECT)

            SendKeys "{TAB}"


        End If
    End If
End Sub

Private Sub txtCartaCredito_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        BuscaCartaCredito 1
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtImp_Descuento_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       SendKeys "{TAB}"
    End If
End Sub

Private Sub txtImp_Flete_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub txtImp_Seguro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub TxtNom_Cliente_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        KeyAscii = 0
'        BuscaCliente 2
'        SendKeys "{TAB}"
'    End If

    If KeyAscii = 13 Then
        If Len(txtNom_Cliente) > 4 Then
            strSQL = "SELECT Abr_Cliente FROM TG_CLIENTE WHERE Nom_Cliente LIKE '" & Trim(txtNom_Cliente.Text) & "%'"
            txtNom_Cliente.Text = DevuelveCampo(strSQL, cCONNECT)
            strSQL = "SELECT Nom_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
            txtNom_Cliente.Text = DevuelveCampo(strSQL, cCONNECT)
            strSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtAbr_Cliente.Text) & "'"
            txtAbr_Cliente.Tag = DevuelveCampo(strSQL, cCONNECT)
            SendKeys "{TAB}"

        Else
            MsgBox ("El Texto Ingresado debe contar con un mínimo de 5 caracteres")
            txtNom_Cliente.SetFocus
        End If
    End If
End Sub

Private Sub CargarDatos()
    txtObservacion.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Observaciones_Factura").Index), vbString)
    txtSecuencia.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Cod_LugEnt").Index), vbLong)
    txtLinea1.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Des_LugEnt").Index), vbString)
    txtCod_CondVent.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), vbString)
    txtDes_CondVent.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Condicion_Venta").Index), vbString)
    txtCartaCredito.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Num_CartaCredito").Index), vbString)
    txtImp_Flete.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Imp_Flete").Index), vbDouble)
    txtImp_Seguro.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Imp_Seguro").Index), vbDouble)
    txtImp_Descuento.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Imp_Descuento").Index), vbDouble)
    txtCod_Termino_Venta = FixNulos(GridEX1.Value(GridEX1.Columns("Cod_Termino_Venta").Index), vbString)
    txtDes_Termino_Venta = FixNulos(GridEX1.Value(GridEX1.Columns("Des_Termino_Venta").Index), vbString)
    txtCod_Embarque.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Cod_Embarque").Index), vbString)
    txtDes_Embarque.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Des_Embarque").Index), vbString)
    txtNom_Embarque.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Nom_Embarque").Index), vbString)
    txtPie_Pagina1.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index), vbString)
    txtPie_Pagina2.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index), vbString)
    txtCod_Vendor.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Cod_Vendor").Index), vbString)
    txtCod_Class.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Cod_Class").Index), vbString)
    txtPor_Comision.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Por_Comision").Index), vbDouble)
    
    txtRef_Embarque.Text = FixNulos(DevuelveCampo("select ref_embarque FROM TG_EMBARQUE where num_embarque = '" & FixNulos(GridEX1.Value(GridEX1.Columns("Num_Embarque").Index), vbLong) & "'", cCONNECT), vbString)
    
    txtImp_Desaduanaje.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Imp_Desaduanaje").Index), vbDouble)
    txtImp_Transporte_Pais_Destino.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Imp_Transporte_Pais_Destino").Index), vbDouble)
    
    
    Me.fraDatosAdicionales.Visible = True
    Me.txtRef_Embarque.SetFocus
End Sub

Private Sub GuardarDatos()
On Error GoTo errx
Dim ssql As String

    GridEX1.Value(GridEX1.Columns("Observaciones_Factura").Index) = txtObservacion.Text
    GridEX1.Value(GridEX1.Columns("Cod_LugEnt").Index) = Val(txtSecuencia)
    GridEX1.Value(GridEX1.Columns("Des_LugEnt").Index) = txtLinea1
    GridEX1.Value(GridEX1.Columns("Num_CartaCredito").Index) = FixNulos(txtCartaCredito.Text, vbString)
    GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index) = txtCod_CondVent.Text
    GridEX1.Value(GridEX1.Columns("Condicion_Venta").Index) = txtDes_CondVent.Text
    GridEX1.Value(GridEX1.Columns("Imp_Flete").Index) = txtImp_Flete
    GridEX1.Value(GridEX1.Columns("Imp_Seguro").Index) = txtImp_Seguro.Text
    GridEX1.Value(GridEX1.Columns("Cod_Termino_Venta").Index) = txtCod_Termino_Venta.Text
    GridEX1.Value(GridEX1.Columns("Des_Termino_Venta").Index) = txtDes_Termino_Venta.Text
    GridEX1.Value(GridEX1.Columns("Imp_Descuento").Index) = txtImp_Descuento.Text
    GridEX1.Value(GridEX1.Columns("cod_Embarque").Index) = txtCod_Embarque.Text
    GridEX1.Value(GridEX1.Columns("Nom_Embarque").Index) = txtNom_Embarque.Text
    GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index) = txtPie_Pagina1.Text
    GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index) = txtPie_Pagina2.Text
    GridEX1.Value(GridEX1.Columns("Cod_Vendor").Index) = txtCod_Vendor.Text
    GridEX1.Value(GridEX1.Columns("Cod_Class").Index) = txtCod_Class.Text
    GridEX1.Value(GridEX1.Columns("Num_Embarque").Index) = FixNulos(DevuelveCampo("select num_embarque FROM TG_EMBARQUE where ref_embarque = '" & txtRef_Embarque.Text & "'", cCONNECT), vbLong)
    GridEX1.Value(GridEX1.Columns("Por_Comision").Index) = txtPor_Comision.Text
    GridEX1.Value(GridEX1.Columns("Imp_Desaduanaje").Index) = txtImp_Desaduanaje.Text
    GridEX1.Value(GridEX1.Columns("Imp_Transporte_Pais_Destino").Index) = txtImp_Transporte_Pais_Destino.Text

      ssql = "Ventas_Cambio_Estado_DocAlm_Prendas '$','$','$','$','$',$,'$',$,$,'$','$','$' ,'$','$','$','$',$,$,$,'$',$,'$','$','$','$','$','$','$','$','$',$,'$','$',$,$"
            
      ssql = VBsprintf(ssql, Left(Cbo_Almacen, 2), _
                       GridEX1.Value(GridEX1.Columns("num_movstk").Index), _
                       GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), _
                       GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
                       GridEX1.Value(GridEX1.Columns("Gastos_Financieros").Index), _
                       GridEX1.Value(GridEX1.Columns("Otros").Index), sCod_TipoFact, _
                       GridEX1.Value(GridEX1.Columns("cod_tipanex").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Anxo").Index), _
                       GridEX1.Value(GridEX1.Columns("Observaciones_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_LugEnt").Index), _
                       FixNulos(GridEX1.Value(GridEX1.Columns("Num_CartaCredito").Index), vbString), _
                       GridEX1.Value(GridEX1.Columns("Cod_cliente").Index), _
                       GridEX1.Value(GridEX1.Columns("Imp_Flete").Index), _
                       GridEX1.Value(GridEX1.Columns("Imp_Seguro").Index), _
                       GridEX1.Value(GridEX1.Columns("Imp_Descuento").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Termino_Venta").Index), _
                       GridEX1.Value(GridEX1.Columns("Num_Packing").Index), _
                       GridEX1.Value(GridEX1.Columns("cod_Embarque").Index), _
                       GridEX1.Value(GridEX1.Columns("Nom_Embarque").Index), _
                       GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index), _
                       GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index), IIf(GridEX1.Value(GridEX1.Columns("Sel").Index) = 0, "P", "A"), GridEX1.Value(GridEX1.Columns("COD_ESTCLI").Index), GridEX1.Value(GridEX1.Columns("Fecha").Index), GridEX1.Value(GridEX1.Columns("Cod_Vendor").Index), GridEX1.Value(GridEX1.Columns("Cod_Class").Index), GridEX1.Value(GridEX1.Columns("Num_Embarque").Index), vusu, GridEX1.Value(GridEX1.Columns("Por_comision").Index), GridEX1.Value(GridEX1.Columns("imp_Desaduanaje").Index), GridEX1.Value(GridEX1.Columns("Imp_Transporte_Pais_Destino").Index))

                           
    ExecuteCommandSQL cCONNECT, ssql

    DatosAdic_Click
    
    GridEX1.Value(GridEX1.Columns("Observaciones_Factura").Index) = txtObservacion.Text
    GridEX1.Value(GridEX1.Columns("Cod_LugEnt").Index) = Val(txtSecuencia)
    GridEX1.Value(GridEX1.Columns("Des_LugEnt").Index) = txtLinea1
    GridEX1.Value(GridEX1.Columns("Num_CartaCredito").Index) = FixNulos(txtCartaCredito.Text, vbString)
    GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index) = txtCod_CondVent.Text
    GridEX1.Value(GridEX1.Columns("Condicion_Venta").Index) = txtDes_CondVent.Text
    GridEX1.Value(GridEX1.Columns("Imp_Flete").Index) = txtImp_Flete
    GridEX1.Value(GridEX1.Columns("Imp_Seguro").Index) = txtImp_Seguro.Text
    GridEX1.Value(GridEX1.Columns("Cod_Termino_Venta").Index) = txtCod_Termino_Venta.Text
    GridEX1.Value(GridEX1.Columns("Des_Termino_Venta").Index) = txtDes_Termino_Venta.Text
    GridEX1.Value(GridEX1.Columns("Imp_Descuento").Index) = txtImp_Descuento.Text
    GridEX1.Value(GridEX1.Columns("cod_Embarque").Index) = txtCod_Embarque.Text
    GridEX1.Value(GridEX1.Columns("Nom_Embarque").Index) = txtNom_Embarque.Text
    GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index) = txtPie_Pagina1.Text
    GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index) = txtPie_Pagina2.Text
    GridEX1.Value(GridEX1.Columns("Cod_Vendor").Index) = txtCod_Vendor.Text
    GridEX1.Value(GridEX1.Columns("Cod_Class").Index) = txtCod_Class.Text
    GridEX1.Value(GridEX1.Columns("Num_Embarque").Index) = FixNulos(DevuelveCampo("select num_embarque FROM TG_EMBARQUE where ref_embarque = '" & txtRef_Embarque.Text & "'", cCONNECT), vbLong)
    GridEX1.Value(GridEX1.Columns("Por_Comision").Index) = txtPor_Comision.Text
    GridEX1.Value(GridEX1.Columns("Imp_Desaduanaje").Index) = txtImp_Desaduanaje.Text
    GridEX1.Value(GridEX1.Columns("Imp_Transporte_Pais_Destino").Index) = txtImp_Transporte_Pais_Destino.Text
    
    Me.fraDatosAdicionales.Visible = False
Exit Sub
errx:
    errores err.Number
End Sub

Private Sub DatosAdic_Click()

Dim serie As String, Nro_Factura As String, iPos, I As Integer, lvSw As Boolean

  GridEX1.Redraw = False

  lvSw = True
  
  serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
  Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index)
  
  
  GridEX1.MoveFirst
  For I = 0 To GridEX1.RowCount
    If serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) And Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index) Then
      If lvSw Then iPos = GridEX1.Row
      lvSw = False
        GridEX1.Value(GridEX1.Columns("Observaciones_Factura").Index) = txtObservacion.Text
        GridEX1.Value(GridEX1.Columns("Cod_LugEnt").Index) = Val(txtSecuencia)
        GridEX1.Value(GridEX1.Columns("Des_LugEnt").Index) = txtLinea1.Text
        GridEX1.Value(GridEX1.Columns("Num_CartaCredito").Index) = FixNulos(txtCartaCredito.Text, vbString)
        GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index) = txtCod_CondVent.Text
        GridEX1.Value(GridEX1.Columns("Condicion_Venta").Index) = txtDes_CondVent.Text
        GridEX1.Value(GridEX1.Columns("Cod_Termino_Venta").Index) = txtCod_Termino_Venta.Text
        GridEX1.Value(GridEX1.Columns("Des_Termino_Venta").Index) = txtDes_Termino_Venta.Text
        GridEX1.Value(GridEX1.Columns("Imp_Flete").Index) = txtImp_Flete.Text
        GridEX1.Value(GridEX1.Columns("Imp_Seguro").Index) = txtImp_Seguro.Text
        GridEX1.Value(GridEX1.Columns("Imp_Descuento").Index) = txtImp_Descuento.Text
        GridEX1.Value(GridEX1.Columns("Nom_Embarque").Index) = txtNom_Embarque.Text
        GridEX1.Value(GridEX1.Columns("cod_Embarque").Index) = txtCod_Embarque.Text
        GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index) = txtPie_Pagina1.Text
        GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index) = txtPie_Pagina2.Text
        GridEX1.Value(GridEX1.Columns("Cod_Vendor").Index) = txtCod_Vendor.Text
        GridEX1.Value(GridEX1.Columns("Cod_Class").Index) = txtCod_Class.Text
        GridEX1.Value(GridEX1.Columns("Num_Embarque").Index) = FixNulos(DevuelveCampo("select num_embarque FROM TG_EMBARQUE where ref_embarque = '" & txtRef_Embarque.Text & "'", cCONNECT), vbLong)
        GridEX1.Value(GridEX1.Columns("Por_Comision").Index) = txtPor_Comision.Text
        GridEX1.Value(GridEX1.Columns("Imp_Desaduanaje").Index) = txtImp_Desaduanaje.Text
        GridEX1.Value(GridEX1.Columns("Imp_Transporte_Pais_Destino").Index) = txtImp_Transporte_Pais_Destino.Text
    End If
    GridEX1.MoveNext
  Next I
  
  GridEX1.Row = iPos
  
  GridEX1.Redraw = True
    
  
End Sub


Private Sub txtobservacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtPie_Pagina1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtPie_Pagina2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtPorc_Descuento_Precio_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And txtPorc_Descuento_Precio > 0 Then
        txtPre_Unitario.Text = GridEX1.Value(GridEX1.Columns("Pre_Unitario_ORG").Index) - Round(GridEX1.Value(GridEX1.Columns("Pre_Unitario_ORG").Index) * (Val(txtPorc_Descuento_Precio) / 100), 2)
        cmdAceptarPrecio.SetFocus
    End If
End Sub

Private Sub txtPre_Unitario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtImp_comision.SetFocus
    End If
End Sub

Private Sub txtRef_Embarque_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaRef_Embarque 1
        SendKeys "{TAB}"
    End If

End Sub

Private Sub txtSecuencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        
        BuscaLugEnt 1
        SendKeys "{TAB}"
    End If
End Sub

Public Sub BuscaLugEnt(Opcion As String)
Dim rstAux As ADODB.Recordset
    strSQL = "SELECT Secuencia, RTRIM(Linea1) + ' ' + RTRIM(Linea2) + " & _
             "RTRIM(Linea3) AS Linea1 FROM TG_CLIENTE_LUGENT " & _
             "WHERE Cod_Cliente = '" & txtAbr_Cliente.Tag & "' AND "
    
    txtSecuencia = Trim(txtSecuencia)
    txtLinea1 = Trim(txtLinea1)
    
    Select Case Opcion
    Case 1: strSQL = strSQL & "CONVERT(varchar(8), Secuencia) like '%" & txtSecuencia & "%'"
    Case 2: strSQL = strSQL & "RTRIM(Linea1) + ' ' + RTRIM(Linea2) + " & _
             "RTRIM(Linea3) LIKE '%" & txtLinea1 & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
    
    frmBusqGeneral3.gexLista.Columns("Secuencia").Visible = False
    frmBusqGeneral3.gexLista.Columns("Secuencia").Width = 570
    frmBusqGeneral3.gexLista.Columns("Linea1").Width = 2370
    
    frmBusqGeneral3.gexLista.Columns("Secuencia").Caption = "Secuencia"
    frmBusqGeneral3.gexLista.Columns("Linea1").Caption = "Lug.Entr."
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtSecuencia = ""
    txtLinea1 = ""
    
    If Codigo <> "" Then
        txtSecuencia = Codigo
        txtLinea1 = Descripcion
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    
    Codigo = ""
    Descripcion = ""
End Sub




Private Sub txtCod_CondVent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaCondVent 1
        SendKeys "{TAB}"
    End If
End Sub

Public Sub BuscaCondVent(Opcion As String)
Dim rstAux As ADODB.Recordset

    strSQL = "SELECT Cod_CondVent, Des_CondVent FROM lg_condvent WHERE "
    
    txtCod_CondVent = Trim(txtCod_CondVent)
    txtDes_CondVent = Trim(txtDes_CondVent)
    
    Select Case Opcion
    Case 1: strSQL = strSQL & "Cod_condVent like '%" & txtCod_CondVent & "%'"
    Case 2: strSQL = strSQL & "Des_condVent LIKE '%" & txtDes_CondVent & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
    
    
    frmBusqGeneral3.gexLista.Columns("Cod_CondVent").Width = 700
    frmBusqGeneral3.gexLista.Columns("Des_CondVent").Width = 2000
    
    frmBusqGeneral3.gexLista.Columns("Cod_CondVent").Caption = "Cond.Vta"
    frmBusqGeneral3.gexLista.Columns("Des_condVent").Caption = "Descrip."
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtCod_CondVent = ""
    txtDes_CondVent = ""
    
    If Codigo <> "" Then
        txtCod_CondVent = Codigo
        txtDes_CondVent = Descripcion
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    
    Codigo = ""
    Descripcion = ""
End Sub


Private Function Busca_AnexosCliente()
  If GridEX1.RowCount > 0 Then
      Set GridEX4.ADORecordset = CargarRecordSetDesconectado("SM_TG_CLIENTE_ANEXOCONT '" & GridEX1.Value(GridEX1.Columns("COD_CLIENTE").Index) & "'", cCONNECT)
      GridEX4.ColumnAutoResize = True
    
      
      GridEX4.Columns("COD_CLIENTE").Visible = False
      GridEX4.Columns("COD_ANXO").Visible = False
      GridEX4.Columns("COD_TIPANEX").Visible = False
    
        GridEX4.ActAsDropDown = True
        GridEX4.BoundColumnIndex = 2
        GridEX4.ReplaceColumnIndex = 2
    
  End If
  
End Function

Private Sub txtCod_TipoFact_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaTipoFacturacion 1
        FunctButt1.SetFocus
    End If
End Sub


Public Sub BuscaTipoFacturacion(Opcion As String)
Dim rstAux As ADODB.Recordset

    strSQL = "SELECT Cod_TipoFact, Des_TipoFact FROM CN_TipoFactura_Venta WHERE "
    
    txtCod_TipoFact = Trim(txtCod_TipoFact)
    txtDes_TipoFact = Trim(txtDes_TipoFact)
    
    Select Case Opcion
    Case 1: strSQL = strSQL & "Cod_TipoFact LIKE '%" & txtCod_TipoFact & "%'"
    Case 2: strSQL = strSQL & "Des_TipoFact LIKE '%" & txtDes_TipoFact & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
    
    frmBusqGeneral3.gexLista.Columns("Cod_tipoFact").Width = 800
    frmBusqGeneral3.gexLista.Columns("Des_TipoFact").Width = 10000
    
    frmBusqGeneral3.gexLista.Columns("Cod_tipoFact").Caption = "Tipo de Facturación"
    frmBusqGeneral3.gexLista.Columns("des_tipoFact").Caption = "Descripción de Facturación "
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtCod_TipoFact.Tag = ""
    txtCod_TipoFact = ""
    txtDes_TipoFact = ""
    
    If Codigo <> "" Then
        
        txtCod_TipoFact = Codigo
        txtDes_TipoFact = Descripcion
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    
    Codigo = ""
    Descripcion = ""
End Sub


Private Sub SeleccionarOtrosReg(Valor As Variant)
    Dim serie As String, Nro_Factura As String, iPos, I As Integer, lvSw As Boolean
    Dim ssql As String
    
    GridEX1.Redraw = False

    lvSw = True
  
    serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
    Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index)
  
    GridEX1.MoveFirst
    For I = 0 To GridEX1.RowCount
        If serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) And _
           Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index) Then
           
            If lvSw Then iPos = GridEX1.Row
            lvSw = False
            GridEX1.Value(GridEX1.Columns("Sel").Index) = Valor
            ssql = "Ventas_Cambio_Estado_DocAlm_Prendas '$','$','$','$','$',$,'$',$,$,'$','$','$' ,'$','$','$','$',$,$,$,'$',$,'$','$','$','$','$','$','$','$','$',$,'$','$'"
                
            ssql = VBsprintf(ssql, Left(Cbo_Almacen, 2), _
                           GridEX1.Value(GridEX1.Columns("num_movstk").Index), _
                           GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
                           GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
                           GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), _
                           GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
                           GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
                           GridEX1.Value(GridEX1.Columns("Gastos_Financieros").Index), _
                           GridEX1.Value(GridEX1.Columns("Otros").Index), sCod_TipoFact, _
                           GridEX1.Value(GridEX1.Columns("cod_tipanex").Index), _
                           GridEX1.Value(GridEX1.Columns("Cod_Anxo").Index), _
                           GridEX1.Value(GridEX1.Columns("Observaciones_Factura").Index), _
                           GridEX1.Value(GridEX1.Columns("Cod_LugEnt").Index), _
                           FixNulos(GridEX1.Value(GridEX1.Columns("Num_CartaCredito").Index), vbString), _
                           GridEX1.Value(GridEX1.Columns("Cod_cliente").Index), _
                           GridEX1.Value(GridEX1.Columns("Imp_Flete").Index), _
                           GridEX1.Value(GridEX1.Columns("Imp_Seguro").Index), GridEX1.Value(GridEX1.Columns("Imp_DESCUENTO").Index), GridEX1.Value(GridEX1.Columns("Cod_Termino_Venta").Index), GridEX1.Value(GridEX1.Columns("Num_Packing").Index), _
                           GridEX1.Value(GridEX1.Columns("cod_Embarque").Index), _
                           GridEX1.Value(GridEX1.Columns("Nom_Embarque").Index), _
                           GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index), _
                           GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index), IIf(GridEX1.Value(GridEX1.Columns("Sel").Index) = 0, "P", "A"), GridEX1.Value(GridEX1.Columns("COD_ESTCLI").Index), GridEX1.Value(GridEX1.Columns("Fecha").Index), GridEX1.Value(GridEX1.Columns("Cod_Vendor").Index), GridEX1.Value(GridEX1.Columns("Cod_Class").Index), GridEX1.Value(GridEX1.Columns("Num_Embarque").Index), vusu, GridEX1.Value(GridEX1.Columns("Por_Comision").Index))
            ExecuteCommandSQL cCONNECT, ssql
        End If
        GridEX1.MoveNext
    Next I
    GridEX1.Row = iPos
    GridEX1.Redraw = True
End Sub


Private Sub txtCod_Termino_Venta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaTerminoVent 1
        SendKeys "{TAB}"
    End If
End Sub

Public Sub BuscaTerminoVent(Opcion As String)
Dim rstAux As ADODB.Recordset

    strSQL = "SELECT Cod_Termino_Venta, Des_Termino_Venta FROM CN_Termino_Venta WHERE "
    
    txtCod_Termino_Venta = Trim(txtCod_Termino_Venta)
    txtDes_Termino_Venta = Trim(txtDes_Termino_Venta)
    
    Select Case Opcion
    Case 1: strSQL = strSQL & "Cod_Termino_Venta like '%" & txtCod_Termino_Venta & "%'"
    Case 2: strSQL = strSQL & "Des_Termino_Venta LIKE '%" & txtDes_Termino_Venta & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
    
    frmBusqGeneral3.gexLista.Columns("Cod_Termino_Venta").Width = 700
    frmBusqGeneral3.gexLista.Columns("Des_Termino_Venta").Width = 2000
    
    frmBusqGeneral3.gexLista.Columns("Cod_Termino_Venta").Caption = "Termino.Venta"
    frmBusqGeneral3.gexLista.Columns("Des_Termino_Venta").Caption = "Descrip."
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtCod_Termino_Venta = ""
    txtDes_Termino_Venta = ""
    
    If Codigo <> "" Then
        txtCod_Termino_Venta = Codigo
        txtDes_Termino_Venta = Descripcion
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    
    Codigo = ""
    Descripcion = ""
End Sub



Private Sub CargarPrecio()
Dim ssql As String
On Error GoTo errx
    txtPre_Unitario = FixNulos(GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), vbDouble)
    txtPorc_Descuento_Precio = 100 - ((txtPre_Unitario * 100) / GridEX1.Value(GridEX1.Columns("Pre_Unitario_Org").Index))
    Me.fraPrecio.Visible = True
    Me.fraPrecio.Top = Me.fraDatosAdicionales.Top
Exit Sub
errx:
    errores err.Number
End Sub

Private Sub GrabarPrecio()
Dim ssql As String
On Error GoTo errx
    GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index) = txtPre_Unitario.Text
    GridEX1.Value(GridEX1.Columns("Imp_Comision").Index) = txtImp_comision.Text
    
    Me.fraPrecio.Visible = True

    ssql = "UP_MAN_TEMP_Ventas_Precios '$','$','$','$','$',$,'$','$','$','$','$','$',$,$,$,$"
            
    ssql = VBsprintf(ssql, "I", vusu, Left(Cbo_Almacen, 2), _
            GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
            GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
            GridEX1.Value(GridEX1.Columns("Num_Packing").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_cliente").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_PurOrd").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_LotPurOrd").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_Estcli").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_ColCli").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_Talla").Index), _
            GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
            GridEX1.Value(GridEX1.Columns("Pre_Unitario_Org").Index) - GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
            GridEX1.Value(GridEX1.Columns("Pre_Unitario_Org").Index), _
            GridEX1.Value(GridEX1.Columns("Imp_Comision").Index))
            
    ExecuteCommandSQL cCONNECT, ssql
    Me.fraPrecio.Visible = False
Exit Sub
errx:
    errores err.Number
End Sub




Private Sub txtCod_Emabarque_Venta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaModoTransporte 1
        SendKeys "{TAB}"
    End If
End Sub

Public Sub BuscaModoTransporte(Opcion As String)
Dim rstAux As ADODB.Recordset

    strSQL = "SELECT Cod_Embarque, Des_Embarque FROM TG_TIPEMB WHERE "
    
    txtCod_Embarque = Trim(txtCod_Embarque)
    txtDes_Embarque = Trim(txtDes_Embarque)
    
    Select Case Opcion
    Case 1: strSQL = strSQL & "Cod_Embarque like '%" & txtCod_Embarque & "%'"
    Case 2: strSQL = strSQL & "Des_Embarque LIKE '%" & txtDes_Embarque & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
    
    frmBusqGeneral3.gexLista.Columns("Cod_Embarque").Width = 700
    frmBusqGeneral3.gexLista.Columns("Des_Embarque").Width = 2000
    
    frmBusqGeneral3.gexLista.Columns("Cod_Embarque").Caption = "Embarque"
    frmBusqGeneral3.gexLista.Columns("Des_Embarque").Caption = "Descrip."
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtCod_Embarque = ""
    txtDes_Embarque = ""
    
    If Codigo <> "" Then
        txtCod_Embarque = Codigo
        txtDes_Embarque = Descripcion
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    
    Codigo = ""
    Descripcion = ""
End Sub

Public Function CargaValores(ByRef ObjTemp As Object) As Boolean
    ObjTemp.txtAbr_Cliente.Text = txtAbr_Cliente.Text
    ObjTemp.txtAbr_Cliente.Tag = txtAbr_Cliente.Tag
    ObjTemp.txtDes_Cliente.Text = txtNom_Cliente.Text
    'ObjTemp.txtCOD_TEMCLI.Text = gexLista.Value(gexLista.Columns("COD_TEMCLI").Index)
    'ObjTemp.CARGA_ESTCLI
End Function


Private Sub Cambio_Fecha(sFecha As String)
    Dim serie As String, Nro_Factura As String, iPos, I As Integer, lvSw As Boolean
    Dim ssql As String
    
    Dim xSerFactura As String
    Dim xNumFactura As String
    
    GridEX1.Redraw = False
    lvSw = True
  
    serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
    Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index)
  
    GridEX1.MoveFirst
    For I = 0 To GridEX1.RowCount
        xSerFactura = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
        xNumFactura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index)
        If serie = xSerFactura And Nro_Factura = xNumFactura Then
            If lvSw Then iPos = GridEX1.Row
            lvSw = False
            GridEX1.Value(GridEX1.Columns("Fecha").Index) = sFecha
        End If
        GridEX1.MoveNext
    Next I
    GridEX1.Row = iPos
    GridEX1.Redraw = True
End Sub

Public Sub BuscaRef_Embarque(Opcion As String)
Dim rstAux As ADODB.Recordset
Dim rsData As ADODB.Recordset

    'strSQL = "SELECT Ref_Embarque , Obs_Embarque FROM TG_EMBARQUE WHERE FLG_STATUS in ('T','F')  AND COD_TIPANEX = '" & GridEX1.Value(GridEX1.Columns("COD_TIPANEX").Index) & "' AND  COD_ANXO = '" & GridEX1.Value(GridEX1.Columns("COD_ANXO").Index) & "'   AND COD_CLIENTE = '" & GridEX1.Value(GridEX1.Columns("COD_CLIENTE").Index) & "' AND "
    strSQL = "SELECT Ref_Embarque , Obs_Embarque FROM TG_EMBARQUE WHERE FLG_STATUS in ('T','F') AND  COD_ANXO = '" & GridEX1.Value(GridEX1.Columns("COD_ANXO").Index) & "'   AND COD_CLIENTE = '" & GridEX1.Value(GridEX1.Columns("COD_CLIENTE").Index) & "' AND "
    
    txtRef_Embarque = Trim(txtRef_Embarque)
    
    Select Case Opcion
    Case 1: strSQL = strSQL & "Ref_Embarque like '%" & txtRef_Embarque & "%'"
    End Select
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
    
    frmBusqGeneral3.gexLista.Columns("Ref_Embarque").Width = 1700
    frmBusqGeneral3.gexLista.Columns("obs_Embarque").Width = 2000
    
    frmBusqGeneral3.gexLista.Columns("Ref_Embarque").Caption = "Número Embarque"
    frmBusqGeneral3.gexLista.Columns("Obs_Embarque").Caption = "Observaciones"
    
    If frmBusqGeneral3.gexLista.RowCount = 0 Then
        MsgBox "Embarque no existe", 1
        Exit Sub
    End If
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
    
    txtRef_Embarque = ""
    
    
    If Codigo <> "" Then
        txtRef_Embarque = Codigo
        If txtRef_Embarque <> "" Then
            strSQL = "TG_Embarques_Muestra '$','$','$','$','$','$','$'"
            strSQL = VBsprintf(strSQL, "3", 0, txtRef_Embarque, "", "", "", "")
            Set rsData = GetDataSet(cCONNECT, strSQL)
            If Not rsData Is Nothing Then
                Do While Not rsData.EOF
                    If RTrim(txtCod_Termino_Venta) = "" Then
                        txtCod_Termino_Venta = FixNulos(rsData("Cod_Termino_venta").Value, vbString)
                        txtDes_Termino_Venta = FixNulos(rsData("Des_Termino_Venta").Value, vbString)
                    End If
                    If RTrim(txtCod_Embarque.Text) = "" Then
                        txtCod_Embarque.Text = FixNulos(rsData("Cod_Embarque").Value, vbString)
                        txtDes_Embarque.Text = FixNulos(rsData("Des_Embarque").Value, vbString)
                    End If
                    If RTrim(txtNom_Embarque.Text) = "" Then
                        txtNom_Embarque.Text = FixNulos(rsData("Nom_Embarque").Value, vbString)
                    End If
                    
                    rsData.MoveNext
                Loop
                rsData.Close
            End If
            Set rsData = Nothing
            
        End If
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    
    Codigo = ""
    Descripcion = ""
End Sub



Private Sub Cambio_PO_Factura(sPO As String)
    Dim ssql As String
    On Error GoTo errx

    GridEX1.Value(GridEX1.Columns("Cod_PurOrd_Factura").Index) = sPO
    ssql = "UP_MAN_TEMP_Ventas_PurOrd_Factura '$','$','$','$','$',$,'$','$','$','$','$','$','$','$'"
            
    ssql = VBsprintf(ssql, "I", vusu, Left(Cbo_Almacen, 2), _
            GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
            GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
            GridEX1.Value(GridEX1.Columns("Num_Packing").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_cliente").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_PurOrd").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_LotPurOrd").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_Estcli").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_ColCli").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_Talla").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_PurOrd_Factura").Index), _
            GridEX1.Value(GridEX1.Columns("pre_unitario").Index))
            
    ExecuteCommandSQL cCONNECT, ssql

Exit Sub
errx:
    errores err.Number

End Sub



Private Function GrabaDatosParaFacturaCambiada() As Boolean
Dim ssql As String
Dim num_factura As String
On Error GoTo errx
    
    num_factura = "00000000"
    num_factura = num_factura + Trim(GridEX1.Value(GridEX1.Columns("Num_Factura").Index))
    num_factura = Right(num_factura, 8)
        
    ssql = "UP_MAN_TEMP_Ventas_NUEVO_DATO '$','$','$','$','$','$',$,'$','$','$','$','$','$','$',$,$,$,$"
            
    'GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _

    ssql = VBsprintf(ssql, vusu, Left(Cbo_Almacen, 2), _
            sSer_Factura_Orig, sNum_Factura_Orig, _
            GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
            num_factura, _
            GridEX1.Value(GridEX1.Columns("Num_Packing").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_cliente").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_PurOrd").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_LotPurOrd").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_Estcli").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_ColCli").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_Talla").Index), _
            GridEX1.Value(GridEX1.Columns("Cod_PurOrd_Factura").Index), _
            GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
            GridEX1.Value(GridEX1.Columns("Imp_Descuento").Index), _
            GridEX1.Value(GridEX1.Columns("Pre_Unitario_Org").Index), _
            GridEX1.Value(GridEX1.Columns("Imp_Comision").Index))
    ExecuteCommandSQL cCONNECT, ssql
    
    GrabaDatosParaFacturaCambiada = True
    
Exit Function
errx:
    errores err.Number
    GrabaDatosParaFacturaCambiada = False
End Function




Public Sub BuscaCartaCredito(Opcion As String)
Dim rstAux As ADODB.Recordset
    strSQL = "SELECT Num_CartaCredito , Fec_Emision " & _
             "FROM TG_Carta_Credito " & _
             "WHERE Cod_Cliente = '" & txtAbr_Cliente.Tag & "' AND "
    
    txtCartaCredito = Trim(txtCartaCredito)
        
    Select Case Opcion
    Case 1: strSQL = strSQL & "Num_CartaCredito like '%" & txtCartaCredito & "%'"
    End Select
    strSQL = strSQL & " AND FLG_STATUS IN ('B','F','T')"
    
    Set frmBusqGeneral3.oParent = Me
    frmBusqGeneral3.SQuery = strSQL
    frmBusqGeneral3.CARGAR_DATOS
    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
    
    frmBusqGeneral3.gexLista.Columns("Num_CartaCredito").Visible = True
    frmBusqGeneral3.gexLista.Columns("Num_CartaCredito").Width = 2000
    frmBusqGeneral3.gexLista.Columns("Fec_Emision").Width = 1500
    
    frmBusqGeneral3.gexLista.Columns("Num_CartaCredito").Caption = "Carta Credito"
    frmBusqGeneral3.gexLista.Columns("Fec_Emision").Caption = "Fec_Emision"
    
    If frmBusqGeneral3.gexLista.RowCount > 1 Then
        frmBusqGeneral3.Show vbModal
    Else
        frmBusqGeneral3.cmdAceptar.Value = True
    End If
       
    If Codigo <> "" Then
        txtCartaCredito = Codigo
    End If
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    
    Codigo = ""
    Descripcion = ""
End Sub


