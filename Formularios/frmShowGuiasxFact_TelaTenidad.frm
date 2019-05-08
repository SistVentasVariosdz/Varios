VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmShowGuiasxFact_TelaTenida 
   Caption         =   "Autorización de Pago de Documentos Tela Cruda / Teñida"
   ClientHeight    =   9270
   ClientLeft      =   270
   ClientTop       =   615
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   ScaleHeight     =   9270
   ScaleWidth      =   11745
   WindowState     =   2  'Maximized
   Begin VB.Frame fraDatosAdicionales 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Datos Adicionales"
      Height          =   7965
      Left            =   3120
      TabIndex        =   24
      Top             =   1320
      Visible         =   0   'False
      Width           =   7875
      Begin VB.TextBox txtCartaCredito 
         Height          =   315
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2070
         Width           =   2040
      End
      Begin VB.TextBox txtObservacion 
         Height          =   885
         Left            =   1755
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   675
         Width           =   5940
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   500
         Left            =   2895
         TabIndex        =   42
         Top             =   7335
         Width           =   990
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancelar"
         Height          =   500
         Left            =   3945
         TabIndex        =   43
         Top             =   7335
         Width           =   990
      End
      Begin VB.TextBox txtDes_CondVent 
         Height          =   285
         Left            =   2400
         TabIndex        =   44
         Top             =   1695
         Width           =   4815
      End
      Begin VB.TextBox txtCod_CondVent 
         Height          =   285
         Left            =   1755
         TabIndex        =   27
         Top             =   1695
         Width           =   585
      End
      Begin VB.TextBox txtImp_Seguro 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3780
         TabIndex        =   31
         Text            =   "0"
         Top             =   2880
         Width           =   1125
      End
      Begin VB.TextBox txtImp_Flete 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1740
         TabIndex        =   30
         Text            =   "0"
         Top             =   2895
         Width           =   1125
      End
      Begin VB.TextBox txtImp_Descuento 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6060
         TabIndex        =   32
         Text            =   "0"
         Top             =   2880
         Width           =   1125
      End
      Begin VB.TextBox txtCod_Termino_Venta 
         Height          =   345
         Left            =   1740
         TabIndex        =   29
         Top             =   2475
         Width           =   585
      End
      Begin VB.TextBox txtDes_Termino_Venta 
         Height          =   345
         Left            =   2385
         TabIndex        =   45
         Top             =   2475
         Width           =   4815
      End
      Begin VB.TextBox txtNom_Embarque 
         Height          =   315
         Left            =   1740
         TabIndex        =   36
         Top             =   4170
         Width           =   2340
      End
      Begin VB.TextBox txtDes_Embarque 
         Height          =   345
         Left            =   2385
         TabIndex        =   46
         Top             =   3735
         Width           =   4815
      End
      Begin VB.TextBox txtCod_Embarque 
         Height          =   345
         Left            =   1740
         TabIndex        =   35
         Top             =   3735
         Width           =   585
      End
      Begin VB.TextBox txtPie_Pagina1 
         Height          =   885
         Left            =   1740
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   37
         Top             =   4665
         Width           =   5940
      End
      Begin VB.TextBox txtPie_Pagina2 
         Height          =   885
         Left            =   1755
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   38
         Top             =   5745
         Width           =   5940
      End
      Begin VB.TextBox txtCod_Vendor 
         Height          =   315
         Left            =   1755
         MaxLength       =   20
         TabIndex        =   39
         Top             =   6825
         Width           =   1620
      End
      Begin VB.TextBox txtCod_Class 
         Height          =   315
         Left            =   3930
         MaxLength       =   10
         TabIndex        =   40
         Top             =   6825
         Width           =   1125
      End
      Begin VB.TextBox txtRef_Embarque 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1770
         TabIndex        =   25
         Top             =   225
         Width           =   1830
      End
      Begin VB.TextBox txtPor_Comision 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6285
         MaxLength       =   10
         TabIndex        =   41
         Top             =   6825
         Width           =   930
      End
      Begin VB.TextBox txtImp_Desaduanaje 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1740
         TabIndex        =   33
         Text            =   "0"
         Top             =   3330
         Width           =   1125
      End
      Begin VB.TextBox txtImp_Transporte_Pais_Destino 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6060
         TabIndex        =   34
         Text            =   "0"
         Top             =   3300
         Width           =   1125
      End
      Begin VB.Label Label30 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Carta de Credito"
         Height          =   270
         Left            =   150
         TabIndex        =   67
         Top             =   2145
         Width           =   1485
      End
      Begin VB.Label Label28 
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
         TabIndex        =   66
         Top             =   195
         Width           =   45
      End
      Begin VB.Label Label27 
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
         TabIndex        =   65
         Top             =   225
         Width           =   45
      End
      Begin VB.Label Label26 
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
         TabIndex        =   64
         Top             =   195
         Width           =   45
      End
      Begin VB.Label Label22 
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
         TabIndex        =   63
         Top             =   525
         Width           =   45
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Observaciones :"
         Height          =   195
         Left            =   135
         TabIndex        =   62
         Top             =   735
         Width           =   1155
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Condic.de Ventas"
         Height          =   330
         Left            =   150
         TabIndex        =   61
         Top             =   1755
         Width           =   1590
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Importe Seguro"
         Height          =   465
         Left            =   3000
         TabIndex        =   60
         Top             =   2850
         Width           =   975
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Importe Flete"
         Height          =   255
         Left            =   150
         TabIndex        =   59
         Top             =   2940
         Width           =   1485
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Importe Descuento"
         Height          =   495
         Left            =   5055
         TabIndex        =   58
         Top             =   2865
         Width           =   1080
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Terminos de Ventas"
         Height          =   285
         Left            =   135
         TabIndex        =   57
         Top             =   2535
         Width           =   1590
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Nombre Transporte"
         Height          =   270
         Left            =   150
         TabIndex        =   56
         Top             =   4230
         Width           =   1485
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Modo de Transporte"
         Height          =   315
         Left            =   150
         TabIndex        =   55
         Top             =   3795
         Width           =   1590
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Pie Factura 1:"
         Height          =   195
         Left            =   150
         TabIndex        =   54
         Top             =   4725
         Width           =   990
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Pie Factura 2:"
         Height          =   195
         Left            =   135
         TabIndex        =   53
         Top             =   5805
         Width           =   990
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cod.Vendor"
         Height          =   255
         Left            =   165
         TabIndex        =   52
         Top             =   6900
         Width           =   1485
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Class"
         Height          =   315
         Left            =   3450
         TabIndex        =   51
         Top             =   6870
         Width           =   420
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Número de Embarque"
         Height          =   255
         Left            =   180
         TabIndex        =   50
         Top             =   300
         Width           =   1545
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0FFFF&
         Caption         =   "% Comisión"
         Height          =   315
         Left            =   5325
         TabIndex        =   49
         Top             =   6885
         Width           =   900
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Importe Desaduanaje"
         Height          =   255
         Left            =   150
         TabIndex        =   48
         Top             =   3345
         Width           =   1605
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Importe Transporte en Pais Destino"
         Height          =   315
         Left            =   3075
         TabIndex        =   47
         Top             =   3375
         Width           =   2715
      End
   End
   Begin GridEX20.GridEX GridEX2 
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   4320
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
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmShowGuiasxFact_TelaTenidad.frx":0000
      FormatStyle(2)  =   "frmShowGuiasxFact_TelaTenidad.frx":0138
      FormatStyle(3)  =   "frmShowGuiasxFact_TelaTenidad.frx":01E8
      FormatStyle(4)  =   "frmShowGuiasxFact_TelaTenidad.frx":029C
      FormatStyle(5)  =   "frmShowGuiasxFact_TelaTenidad.frx":0374
      FormatStyle(6)  =   "frmShowGuiasxFact_TelaTenidad.frx":042C
      FormatStyle(7)  =   "frmShowGuiasxFact_TelaTenidad.frx":050C
      ImageCount      =   0
      PrinterProperties=   "frmShowGuiasxFact_TelaTenidad.frx":052C
   End
   Begin GridEX20.GridEX GridEX3 
      Height          =   2055
      Left            =   2880
      TabIndex        =   6
      Top             =   4320
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
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmShowGuiasxFact_TelaTenidad.frx":0704
      FormatStyle(2)  =   "frmShowGuiasxFact_TelaTenidad.frx":083C
      FormatStyle(3)  =   "frmShowGuiasxFact_TelaTenidad.frx":08EC
      FormatStyle(4)  =   "frmShowGuiasxFact_TelaTenidad.frx":09A0
      FormatStyle(5)  =   "frmShowGuiasxFact_TelaTenidad.frx":0A78
      FormatStyle(6)  =   "frmShowGuiasxFact_TelaTenidad.frx":0B30
      FormatStyle(7)  =   "frmShowGuiasxFact_TelaTenidad.frx":0C10
      ImageCount      =   0
      PrinterProperties=   "frmShowGuiasxFact_TelaTenidad.frx":0C30
   End
   Begin VB.Frame FraBuscar 
      BackColor       =   &H00C0FFFF&
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
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   11520
      Begin VB.ComboBox Cbo_Almacen 
         Height          =   315
         Left            =   1920
         TabIndex        =   9
         Top             =   240
         Width           =   3135
      End
      Begin VB.CheckBox optTodos 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Todos"
         Height          =   255
         Left            =   5160
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpFecEmiIni 
         Height          =   315
         Left            =   1950
         TabIndex        =   2
         Top             =   675
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   72024065
         CurrentDate     =   37543
      End
      Begin MSComCtl2.DTPicker dtpFecEmiFin 
         Height          =   315
         Left            =   3990
         TabIndex        =   4
         Top             =   675
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   72024065
         CurrentDate     =   37543
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   525
         Left            =   7320
         TabIndex        =   68
         Top             =   240
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   926
         Custom          =   $"frmShowGuiasxFact_TelaTenidad.frx":0E08
         Orientacion     =   0
         Style           =   1
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1000
         ControlHeigth   =   500
         ControlSeparator=   40
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Almacen"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Rango Fecha de Emisión:"
         Height          =   360
         Left            =   90
         TabIndex        =   3
         Top             =   705
         Width           =   2355
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5340
      Left            =   60
      TabIndex        =   0
      Top             =   1185
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   9419
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
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmShowGuiasxFact_TelaTenidad.frx":0EF1
      FormatStyle(2)  =   "frmShowGuiasxFact_TelaTenidad.frx":1029
      FormatStyle(3)  =   "frmShowGuiasxFact_TelaTenidad.frx":10D9
      FormatStyle(4)  =   "frmShowGuiasxFact_TelaTenidad.frx":118D
      FormatStyle(5)  =   "frmShowGuiasxFact_TelaTenidad.frx":1265
      FormatStyle(6)  =   "frmShowGuiasxFact_TelaTenidad.frx":131D
      FormatStyle(7)  =   "frmShowGuiasxFact_TelaTenidad.frx":13FD
      ImageCount      =   0
      PrinterProperties=   "frmShowGuiasxFact_TelaTenidad.frx":141D
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
      Left            =   9720
      TabIndex        =   23
      Top             =   6960
      Width           =   45
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Guia :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9000
      TabIndex        =   22
      Top             =   6960
      Width           =   510
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Observacion :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   21
      Top             =   6960
      Width           =   1245
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
      Left            =   1440
      TabIndex        =   20
      Top             =   6960
      Width           =   45
   End
   Begin VB.Label lbCod_Color 
      AutoSize        =   -1  'True
      Caption         =   "Color :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9000
      TabIndex        =   19
      Top             =   6660
      Width           =   570
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
      Left            =   9720
      TabIndex        =   18
      Top             =   6660
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Calidad :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6360
      TabIndex        =   17
      Top             =   6660
      Width           =   795
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
      Left            =   7320
      TabIndex        =   16
      Top             =   6630
      Width           =   45
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
      Left            =   8760
      TabIndex        =   15
      Top             =   6660
      Width           =   45
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Nro Rollos :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7560
      TabIndex        =   14
      Top             =   6660
      Width           =   1050
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
      Left            =   6120
      TabIndex        =   13
      Top             =   6660
      Width           =   45
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Comb :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5400
      TabIndex        =   12
      Top             =   6660
      Width           =   630
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
      Left            =   720
      TabIndex        =   11
      Top             =   6630
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tela :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   6630
      Width           =   510
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   6435
      Top             =   4905
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmShowGuiasxFact_TelaTenida"
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

Dim strsql As String
Public CODIGO As String
Public descripcion As String
Public TipoAdd As String
Dim sCod_TipoFact  As String

Dim sSer_Factura_Orig As String
Dim sNum_Factura_Orig As String

Private Sub Form_Resize()
    GridEX1.Width = Me.Width - 300
End Sub
Private Sub DtFecVencimiento_Change()
  GridEX1.ClearFields
  dtpFecEmiIni.Value = ""
  dtpFecEmiFin.Value = ""
End Sub

Private Sub CmdAceptar_Click()
    GuardarDatos
End Sub

Private Sub GuardarDatos()
On Error GoTo errx
Dim sSQL As String

    GridEX1.Value(GridEX1.Columns("Observaciones_Factura").Index) = txtObservacion.Text
    GridEX1.Value(GridEX1.Columns("Num_CartaCredito").Index) = FixNulos(txtCartaCredito.Text, vbString)
    GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index) = txtCod_CondVent.Text
    GridEX1.Value(GridEX1.Columns("Condicion_Venta").Index) = txtDes_CondVent.Text
    GridEX1.Value(GridEX1.Columns("Imp_Flete").Index) = txtImp_Flete
    GridEX1.Value(GridEX1.Columns("Imp_Seguro").Index) = txtImp_Seguro.Text
    GridEX1.Value(GridEX1.Columns("Cod_Termino_Venta").Index) = txtCod_Termino_Venta.Text
   ' GridEX1.Value(GridEX1.Columns("Des_Termino_Venta").Index) = txtDes_Termino_Venta.Text
    GridEX1.Value(GridEX1.Columns("Imp_Descuento").Index) = txtImp_Descuento.Text
    GridEX1.Value(GridEX1.Columns("cod_Embarque").Index) = txtCod_Embarque.Text
    GridEX1.Value(GridEX1.Columns("Nom_Embarque").Index) = txtNom_Embarque.Text
    GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index) = txtPie_Pagina1.Text
    GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index) = txtPie_Pagina2.Text
    GridEX1.Value(GridEX1.Columns("Cod_Vendor").Index) = txtCod_Vendor.Text
    GridEX1.Value(GridEX1.Columns("Cod_Class").Index) = txtCod_Class.Text
    GridEX1.Value(GridEX1.Columns("Num_Embarque").Index) = FixNulos(DevuelveCampo("select num_embarque FROM TG_EMBARQUE where ref_embarque = '" & txtRef_Embarque.Text & "'", cConnect), vbLong)
    GridEX1.Value(GridEX1.Columns("Por_Comision").Index) = txtPor_Comision.Text
    GridEX1.Value(GridEX1.Columns("Imp_Desaduanaje").Index) = txtImp_Desaduanaje.Text
    GridEX1.Value(GridEX1.Columns("Imp_Transporte_Pais_Destino").Index) = txtImp_Transporte_Pais_Destino.Text

      
       sSQL = "Ventas_Cambio_Estado_DocAlm_Prendas '$','$','$','$','$',$,'$',$ ,'$','$','$','$','$','$',$,$,$,'$','$','$','$','$','$','$','$','$','$','$','$',$,$"
        
            
                 sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
                       GridEX1.Value(GridEX1.Columns("num_movstk").Index), _
                       GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), _
                       GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
                       GridEX1.Value(GridEX1.Columns("Otros").Index), sCod_TipoFact, _
                       GridEX1.Value(GridEX1.Columns("cod_tipanex").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Anxo").Index), _
                       GridEX1.Value(GridEX1.Columns("Observaciones_Factura").Index), _
                       FixNulos(GridEX1.Value(GridEX1.Columns("Num_CartaCredito").Index), vbString), _
                       GridEX1.Value(GridEX1.Columns("cliente").Index), _
                       GridEX1.Value(GridEX1.Columns("Imp_Flete").Index), _
                       GridEX1.Value(GridEX1.Columns("Imp_Seguro").Index), _
                       GridEX1.Value(GridEX1.Columns("Imp_Descuento").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Termino_Venta").Index), _
                       GridEX1.Value(GridEX1.Columns("cod_Embarque").Index), _
                       GridEX1.Value(GridEX1.Columns("Nom_Embarque").Index), _
                       GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index), _
                       GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index), _
                       IIf(GridEX1.Value(GridEX1.Columns("Sel").Index) = 0, "P", "A"), _
                       GridEX1.Value(GridEX1.Columns("Fecha").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Vendor").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Class").Index), GridEX1.Value(GridEX1.Columns("Num_Embarque").Index), vusu, GridEX1.Value(GridEX1.Columns("Por_comision").Index), GridEX1.Value(GridEX1.Columns("imp_Desaduanaje").Index), GridEX1.Value(GridEX1.Columns("Imp_Transporte_Pais_Destino").Index))
                           
    ExecuteCommandSQL cConnect, sSQL

    DatosAdic_Click
    
    GridEX1.Value(GridEX1.Columns("Observaciones_Factura").Index) = txtObservacion.Text
'    GridEX1.Value(GridEX1.Columns("Cod_LugEnt").Index) = Val(txtSecuencia)
'    GridEX1.Value(GridEX1.Columns("Des_LugEnt").Index) = txtLinea1
    GridEX1.Value(GridEX1.Columns("Num_CartaCredito").Index) = FixNulos(txtCartaCredito.Text, vbString)
    GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index) = txtCod_CondVent.Text
    GridEX1.Value(GridEX1.Columns("Condicion_Venta").Index) = txtDes_CondVent.Text
    GridEX1.Value(GridEX1.Columns("Imp_Flete").Index) = txtImp_Flete
    GridEX1.Value(GridEX1.Columns("Imp_Seguro").Index) = txtImp_Seguro.Text
    GridEX1.Value(GridEX1.Columns("Cod_Termino_Venta").Index) = txtCod_Termino_Venta.Text
'    GridEX1.Value(GridEX1.Columns("Des_Termino_Venta").Index) = txtDes_Termino_Venta.Text
    GridEX1.Value(GridEX1.Columns("Imp_Descuento").Index) = txtImp_Descuento.Text
    GridEX1.Value(GridEX1.Columns("cod_Embarque").Index) = txtCod_Embarque.Text
    GridEX1.Value(GridEX1.Columns("Nom_Embarque").Index) = txtNom_Embarque.Text
    GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index) = txtPie_Pagina1.Text
    GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index) = txtPie_Pagina2.Text
    GridEX1.Value(GridEX1.Columns("Cod_Vendor").Index) = txtCod_Vendor.Text
    GridEX1.Value(GridEX1.Columns("Cod_Class").Index) = txtCod_Class.Text
    GridEX1.Value(GridEX1.Columns("Num_Embarque").Index) = FixNulos(DevuelveCampo("select num_embarque FROM TG_EMBARQUE where ref_embarque = '" & txtRef_Embarque.Text & "'", cConnect), vbLong)
    GridEX1.Value(GridEX1.Columns("Por_Comision").Index) = txtPor_Comision.Text
    GridEX1.Value(GridEX1.Columns("Imp_Desaduanaje").Index) = txtImp_Desaduanaje.Text
    GridEX1.Value(GridEX1.Columns("Imp_Transporte_Pais_Destino").Index) = txtImp_Transporte_Pais_Destino.Text
    
    Me.fraDatosAdicionales.Visible = False
Exit Sub
errx:
    errores err.Number
End Sub


Private Sub cmdLugEnt_Click()
'    Load frmMantLugaresEntrega
'    frmMantLugaresEntrega.sCod_Cliente = GridEX1.Value(GridEX1.Columns("Cliente").Index)
'    frmMantLugaresEntrega.CARGA_GRID
'    frmMantLugaresEntrega.Show vbModal
'    Set frmMantLugaresEntrega = Nothing
End Sub

Private Sub Command1_Click()
    Me.fraDatosAdicionales.Visible = False
End Sub

Private Sub dtpFecEmiIni_Change()
  GridEX1.ClearFields
  If Trim(dtpFecEmiIni.Value) <> "" Then
    dtpFecEmiFin.Value = dtpFecEmiIni
  End If
End Sub

Private Sub Form_Load()

  dtpFecEmiIni.Value = Date
  dtpFecEmiIni.Value = ""
  
  dtpFecEmiFin.Value = Date
  dtpFecEmiFin.Value = ""
  
  FillAlmacen
  
'  FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name) & "/SALIR"
  
  iLanguage = CInt(GetSetting("Visuales", "Settings", "Language", "1"))
  
  If InStr(FunctButt1.FunctionsUser, "AUTORIZARPAGO") <> 0 Then
      bPuedeAutorizar = True
  End If
  
  Set GridEX2.ADORecordset = CargarRecordSetDesconectado("select Cod_CondVent,Des_CondVent as Descripcion from lg_condvent", cConnect)
    
  GridEX2.ColumnAutoResize = True
'  GridEX2.ClearFields
'  GridEX2.Rebind
  
  'GridEX2 will act as the drop down list
  'for column 'SupplierID' in GridEX1
  
  GridEX2.ActAsDropDown = True
  GridEX2.BoundColumnIndex = 1
  GridEX2.ReplaceColumnIndex = 2
   
  
  GridEX2.Columns("Cod_CondVent").Visible = False
  
  Set GridEX3.ADORecordset = CargarRecordSetDesconectado("select Cod_Moneda as cod_Moneda,Nom_Moneda as Descripcion from tg_moneda", cConnect)
    
  GridEX3.ColumnAutoResize = True

  GridEX3.ActAsDropDown = True
  GridEX3.BoundColumnIndex = 1
  GridEX3.ReplaceColumnIndex = 2
  
  GridEX3.Columns("Cod_Moneda").Visible = False


End Sub

Private Sub BUSCAR()

On Error GoTo drDepurar

Dim sSQL As String
Dim oGroup As GridEX20.JSGroup
Dim oFormat As JSFormatStyle

If Left(Cbo_Almacen, 2) = "31" Then
  sSQL = "VENTAS_MUESTRA_DOCUMENTOS_PENDIENTES_FACTURAR_TELA_TENIDA_EX '" & Left(Cbo_Almacen, 3) & "','" & dtpFecEmiIni & "','" & dtpFecEmiFin & "','" & IIf(optTodos, "*", "") & "'"
  lbCod_Color.Visible = True
  lbDes_Color.Visible = True
ElseIf Left(Cbo_Almacen, 2) = "T1" Then
  sSQL = "Ventas_Muestra_Documentos_Pendientes_Facturar_Tela_Cruda '" & Left(Cbo_Almacen, 3) & "','" & dtpFecEmiIni & "','" & dtpFecEmiFin & "','" & IIf(optTodos, "*", "") & "'"
  lbCod_Color.Visible = False
  lbDes_Color.Visible = False
ElseIf Left(Cbo_Almacen, 2) = "T8" Or Left(Cbo_Almacen, 2) = "T7" Then
  sSQL = "Ventas_Muestra_Documentos_Pendientes_Facturar_Tejeduria '" & Left(Cbo_Almacen, 3) & "','" & dtpFecEmiIni & "','" & dtpFecEmiFin & "','" & IIf(optTodos, "*", "") & "'"
  lbCod_Color.Visible = False
  lbDes_Color.Visible = False
Else
  Exit Sub
End If

GridEX1.ClearFields

GridEX1.DefaultGroupMode = jgexDGMExpanded
bCargaGRid = False
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cConnect)
  
Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Fac_Cli").Index, jgexSortAscending)

MuestraSubTotales
GridEX1.BackColorRowGroup = &H80000005

GridEX1.ColumnHeaderHeight = 500

GridEX1.Columns("OC").Width = 1015
GridEX1.Columns("fecha").Width = 975
GridEX1.Columns("Motivo").Width = 3000
GridEX1.Columns("observaciones").Width = 2000
GridEX1.Columns("cliente").Width = 0
GridEX1.Columns("nro_Guia").Width = 1125
GridEX1.Columns("Cod_Tela").Width = 825
GridEX1.Columns("Tela").Width = 2500
GridEX1.Columns("Comb").Width = 900
GridEX1.Columns("moneda").Width = 900
GridEX1.Columns("Pre_Unitario").Width = 840
GridEX1.Columns("Kgs_Movimiento").Caption = "Kgs"
GridEX1.Columns("Kgs_Movimiento").Width = 765
GridEX1.Columns("Numero_Rollos").Width = 500
GridEX1.Columns("Calidad").Width = 500
GridEX1.Columns("monto despacho").Width = 855
GridEX1.Columns("SEL").Width = 450
'GridEX1.Columns("SEL2").Width = 450
GridEX1.Columns("Fac_Cli").Width = 0
GridEX1.Columns("Gastos Financieros").Width = 900
GridEX1.Columns("otros").Width = 810
GridEX1.Columns("Numero_Rollos").Caption = "Rollos"
GridEX1.Columns("Numero_Rollos").Width = 500
'GridEX1.Columns("DatosAdic").Width = 400
GridEX1.Columns("Ser_Factura").Width = 500
GridEX1.Columns("Num_Factura").Width = 900
GridEX1.Columns("Und").Width = 405

GridEX1.Columns("Ser_Parte_Salida").Visible = False
GridEX1.Columns("Numero_Parte_Salida").Visible = False
GridEX1.Columns("Nom_Cliente").Visible = False
GridEX1.Columns("Cliente").Visible = False
GridEX1.Columns("COD_CONDVENT").Visible = False
GridEX1.Columns("Cod_Moneda").Visible = False
'GridEX1.Columns("Num_movstk").Visible = False
GridEX1.Columns("SER_ORDCOMP").Visible = False
GridEX1.Columns("SEC_ORDCOMP").Visible = False
GridEX1.Columns("COD_ORDCOMP").Visible = False

GridEX1.Columns("Ser_Factura").Caption = "Serie"
GridEX1.Columns("Num_Factura").Caption = "Nro Factura"

GridEX1.Columns("Kgs_Movimiento").Format = "#######0.00"
GridEX1.Columns("Pre_Unitario").Format = "#######0.0000"
GridEX1.Columns("Pre_Unitario").Caption = "Precio"

GridEX1.Columns("monto despacho").Format = "#######0.00"

GridEX1.Columns("SEL").ColumnType = jgexCheckBox
GridEX1.Columns("SEL").Visible = True
GridEX1.Columns("SEL").EditType = jgexEditCheckBox
GridEX1.Columns("SEL").Width = 500

'
'GridEX1.Columns("SEL2").ColumnType = jgexCheckBox
'GridEX1.Columns("SEL2").Visible = True
'GridEX1.Columns("SEL2").EditType = jgexEditCheckBox
'GridEX1.Columns("SEL2").Width = 500

If Left(Cbo_Almacen, 2) = "TT" Then
  GridEX1.Columns("Kgs_a_Facturar").Width = 975
  GridEX1.Columns("Kgs_a_Facturar").Caption = "Kgs a Facturar"
  GridEX1.Columns("Kgs_a_Facturar").Format = "#######0.00"
  GridEX1.Columns("Kgs_a_Facturar").EditType = jgexEditNone
End If

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

SetColores

GridEX1.DefaultGroupMode = jgexDGMCollapsed

If dtpFecEmiIni.Value <> "" Then
    GridEX1.DefaultGroupMode = jgexDGMExpanded
End If

GridEX1.ContinuousScroll = True

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
      BUSCAR
    Case "AUTORIZARPAGO"
        If GridEX1.RowCount = 0 Then Exit Sub
        Msg = MsgBox("¿Esta seguro de autorizar pago?", vbYesNo)
        If Msg = vbNo Then Exit Sub
        Autorizar
    Case "SALIR"
       Unload Me
    End Select
End Sub

Private Sub GridEX1_AfterColEdit(ByVal ColIndex As Integer)

  If Left(Cbo_Almacen, 2) = "TT" Then
    AfterColEdit_Tenido (ColIndex)
    'AfterColEdit_Prendas (ColIndex) '/*mp*/
  ElseIf Left(Cbo_Almacen, 2) = "T1" Then
    AfterColEdit_Crudo (ColIndex)
End If
                           
End Sub

Sub AfterColEdit_Prendas(ByVal ColIndex As Integer)

Dim sSQL As String
On Error GoTo Error_Handler

Dim oGroup As GridEX20.JSGroup


Select Case ColIndex
  Case Is = GridEX1.Columns("SEL2").Index
    
      sSQL = "Ventas_Cambio_Estado_DocAlm_Prendas '$','$','$','$','$',$,'$',$ ,'$','$','$','$','$','$',$,$,$,'$','$','$','$','$','$','$','$','$','$','$','$',$,$"
        

                 sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
                       GridEX1.Value(GridEX1.Columns("num_movstk").Index), _
                       GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), _
                       GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
                       GridEX1.Value(GridEX1.Columns("Otros").Index), sCod_TipoFact, _
                       GridEX1.Value(GridEX1.Columns("cod_tipanex").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Anxo").Index), _
                       GridEX1.Value(GridEX1.Columns("Observaciones_Factura").Index), _
                       FixNulos(GridEX1.Value(GridEX1.Columns("Num_CartaCredito").Index), vbString), _
                       GridEX1.Value(GridEX1.Columns("cliente").Index), _
                       GridEX1.Value(GridEX1.Columns("Imp_Flete").Index), _
                       GridEX1.Value(GridEX1.Columns("Imp_Seguro").Index), _
                       GridEX1.Value(GridEX1.Columns("Imp_Descuento").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Termino_Venta").Index), _
                       GridEX1.Value(GridEX1.Columns("cod_Embarque").Index), _
                       GridEX1.Value(GridEX1.Columns("Nom_Embarque").Index), _
                       GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index), _
                       GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index), _
                       IIf(GridEX1.Value(GridEX1.Columns("Sel").Index) = 0, "P", "A"), _
                       GridEX1.Value(GridEX1.Columns("Fecha").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Vendor").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Class").Index), GridEX1.Value(GridEX1.Columns("Num_Embarque").Index), vusu, GridEX1.Value(GridEX1.Columns("Por_comision").Index), GridEX1.Value(GridEX1.Columns("imp_Desaduanaje").Index), GridEX1.Value(GridEX1.Columns("Imp_Transporte_Pais_Destino").Index))
      
                           
    ExecuteCommandSQL cConnect, sSQL
    SeleccionarOtrosReg GridEX1.Value(GridEX1.Columns("Sel2").Index)
  End Select
Exit Sub

Resume

Error_Handler:

  errores err.Number
   
  If ColIndex = GridEX1.Columns("Sel2").Index Then
     GridEX1.Value(GridEX1.Columns("sel2").Index) = 0
  End If
End Sub


Private Sub SeleccionarOtrosReg(Valor As Variant)
Dim Serie As String, Nro_Factura As String, iPos, i As Integer, lvSW As Boolean
Dim sSQL As String
  GridEX1.Redraw = False

  lvSW = True
  
  Serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
  Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index)
  
  
  GridEX1.MoveFirst
  For i = 0 To GridEX1.RowCount
    If Serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) And Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index) Then
      If lvSW Then iPos = GridEX1.Row
      lvSW = False
        GridEX1.Value(GridEX1.Columns("Sel2").Index) = Valor
      sSQL = "Ventas_Cambio_Estado_DocAlm_Prendas '$','$','$','$','$',$,'$',$ ,'$','$','$','$','$','$',$,$,$,'$','$','$','$','$','$','$','$','$','$','$','$',$,$"
        

                 sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
                       GridEX1.Value(GridEX1.Columns("num_movstk").Index), _
                       GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), _
                       GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
                       GridEX1.Value(GridEX1.Columns("Otros").Index), sCod_TipoFact, _
                       GridEX1.Value(GridEX1.Columns("cod_tipanex").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Anxo").Index), _
                       GridEX1.Value(GridEX1.Columns("Observaciones_Factura").Index), _
                       FixNulos(GridEX1.Value(GridEX1.Columns("Num_CartaCredito").Index), vbString), _
                       GridEX1.Value(GridEX1.Columns("cliente").Index), _
                       GridEX1.Value(GridEX1.Columns("Imp_Flete").Index), _
                       GridEX1.Value(GridEX1.Columns("Imp_Seguro").Index), _
                       GridEX1.Value(GridEX1.Columns("Imp_Descuento").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Termino_Venta").Index), _
                       GridEX1.Value(GridEX1.Columns("cod_Embarque").Index), _
                       GridEX1.Value(GridEX1.Columns("Nom_Embarque").Index), _
                       GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index), _
                       GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index), _
                       IIf(GridEX1.Value(GridEX1.Columns("Sel").Index) = 0, "P", "A"), _
                       GridEX1.Value(GridEX1.Columns("Fecha").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Vendor").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Class").Index), GridEX1.Value(GridEX1.Columns("Num_Embarque").Index), vusu, GridEX1.Value(GridEX1.Columns("Por_comision").Index), GridEX1.Value(GridEX1.Columns("imp_Desaduanaje").Index), GridEX1.Value(GridEX1.Columns("Imp_Transporte_Pais_Destino").Index))
         
         
         
         '1
      ExecuteCommandSQL cConnect, sSQL

    End If
    GridEX1.MoveNext
  Next i
  
  GridEX1.Row = iPos
  
  GridEX1.Redraw = True

End Sub


Sub AfterColEdit_Tenido(ByVal ColIndex As Integer)

Dim sSQL As String
On Error GoTo Error_Handler

Dim oGroup As GridEX20.JSGroup
Select Case ColIndex

  Case Is = GridEX1.Columns("Sel").Index
    
      sSQL = "Ventas_Cambio_Estado_DocAlm_Tela_Tenida '$','$','$','$','$',$,'$',$,$,'$','$','$',$,'$','$'"
      
      sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
                       GridEX1.Value(GridEX1.Columns("num_movstk").Index), _
                       GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), _
                       GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
                       GridEX1.Value(GridEX1.Columns("Gastos Financieros").Index), _
                       GridEX1.Value(GridEX1.Columns("Otros").Index), _
                       GridEX1.Value(GridEX1.Columns("Ser_ordcomp").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_ordcomp").Index), _
                       GridEX1.Value(GridEX1.Columns("Sec_OrdComp").Index), _
                       GridEX1.Value(GridEX1.Columns("Kgs_a_Facturar").Index), _
                       GridEX1.Value(GridEX1.Columns("Num_Secuencia").Index), _
                       GridEX1.Value(GridEX1.Columns("Und").Index))
                           
    ExecuteCommandSQL cConnect, sSQL
    
  Case Is = GridEX1.Columns("Pre_Unitario").Index
        sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
                       GridEX1.Value(GridEX1.Columns("num_movstk").Index), _
                       GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), _
                       GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
                       GridEX1.Value(GridEX1.Columns("Gastos Financieros").Index), _
                       GridEX1.Value(GridEX1.Columns("Otros").Index), _
                       GridEX1.Value(GridEX1.Columns("Ser_ordcomp").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_ordcomp").Index), _
                       GridEX1.Value(GridEX1.Columns("Sec_OrdComp").Index), _
                       GridEX1.Value(GridEX1.Columns("Kgs_a_Facturar").Index), _
                       GridEX1.Value(GridEX1.Columns("Num_Secuencia").Index), _
                       GridEX1.Value(GridEX1.Columns("Und").Index))
  
  
    GridEX1.Value(GridEX1.Columns("Monto Despacho").Index) = GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index) * GridEX1.Value(GridEX1.Columns("Kgs_a_Facturar").Index)
    GridEX1.Value(GridEX1.Columns("sel").Index) = 0
    
    ExecuteCommandSQL cConnect, sSQL
    
  Case Is = GridEX1.Columns("Kgs_a_Facturar").Index
    GridEX1.Value(GridEX1.Columns("Monto Despacho").Index) = GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index) * GridEX1.Value(GridEX1.Columns("Kgs_a_Facturar").Index)
    GridEX1.Value(GridEX1.Columns("sel").Index) = 0
  Case Is = GridEX1.Columns("Ser_Factura").Index
    GridEX1.Value(GridEX1.Columns("Fac_Cli").Index) = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) & "-" & GridEX1.Value(GridEX1.Columns("Num_Factura").Index) & "  " & GridEX1.Value(GridEX1.Columns("Nom_Cliente").Index)
    GridEX1.Groups.Clear
    Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Fac_Cli").Index, jgexSortAscending)
    GridEX1.Value(GridEX1.Columns("sel").Index) = 0
  Case Is = GridEX1.Columns("Num_Factura").Index
    If Trim(GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)) = "" Then GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) = "001"
    GridEX1.Value(GridEX1.Columns("Fac_Cli").Index) = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) & "-" & GridEX1.Value(GridEX1.Columns("Num_Factura").Index) & "  " & GridEX1.Value(GridEX1.Columns("Nom_Cliente").Index)
    GridEX1.Groups.Clear
    Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Fac_Cli").Index, jgexSortAscending)
    GridEX1.Value(GridEX1.Columns("sel").Index) = 0
  Case Is = GridEX1.Columns("Gastos Financieros").Index
    Cambio_Importe "Gastos Financieros"
    GridEX1.Value(GridEX1.Columns("sel").Index) = 0
  Case Is = GridEX1.Columns("Otros").Index
    Cambio_Importe "Otros"
    GridEX1.Value(GridEX1.Columns("sel").Index) = 0
  End Select
Exit Sub

Resume

Error_Handler:

  errores err.Number
   
  If ColIndex = GridEX1.Columns("Sel").Index Then
     GridEX1.Value(GridEX1.Columns("sel").Index) = 0
  End If
End Sub

Sub AfterColEdit_Crudo(ByVal ColIndex As Integer)

Dim sSQL As String
On Error GoTo Error_Handler

Dim oGroup As GridEX20.JSGroup

Select Case ColIndex
  Case Is = GridEX1.Columns("Sel").Index
      sSQL = "Ventas_Cambio_Estado_DocAlm_Tela_Cruda '$','$','$','$','$',$ , '$' , $ ,$,'$','$','$','$'"
      
      sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
                       GridEX1.Value(GridEX1.Columns("num_movstk").Index), _
                       GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), _
                       GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
                       GridEX1.Value(GridEX1.Columns("Gastos Financieros").Index), _
                       GridEX1.Value(GridEX1.Columns("Otros").Index), _
                       GridEX1.Value(GridEX1.Columns("Ser_ordcomp").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_ordcomp").Index), _
                       GridEX1.Value(GridEX1.Columns("Sec_OrdComp").Index), _
                       GridEX1.Value(GridEX1.Columns("Und").Index))
                       
    ExecuteCommandSQL cConnect, sSQL
  Case Is = GridEX1.Columns("Pre_Unitario").Index
    GridEX1.Value(GridEX1.Columns("Monto Despacho").Index) = GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index) * GridEX1.Value(GridEX1.Columns("Kgs_Movimiento").Index)
    GridEX1.Value(GridEX1.Columns("sel").Index) = 0
  Case Is = GridEX1.Columns("Ser_Factura").Index
    GridEX1.Value(GridEX1.Columns("Fac_Cli").Index) = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) & "-" & GridEX1.Value(GridEX1.Columns("Num_Factura").Index) & "  " & GridEX1.Value(GridEX1.Columns("Nom_Cliente").Index)
    GridEX1.Groups.Clear
    Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Fac_Cli").Index, jgexSortAscending)
    GridEX1.Value(GridEX1.Columns("sel").Index) = 0
  Case Is = GridEX1.Columns("Num_Factura").Index
    If Trim(GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)) = "" Then GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) = "001"
    GridEX1.Value(GridEX1.Columns("Fac_Cli").Index) = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) & "-" & GridEX1.Value(GridEX1.Columns("Num_Factura").Index) & "  " & GridEX1.Value(GridEX1.Columns("Nom_Cliente").Index)
    GridEX1.Groups.Clear
    Set oGroup = GridEX1.Groups.Add(GridEX1.Columns("Fac_Cli").Index, jgexSortAscending)
    GridEX1.Value(GridEX1.Columns("sel").Index) = 0
  Case Is = GridEX1.Columns("Gastos Financieros").Index
    Cambio_Importe "Gastos Financieros"
    GridEX1.Value(GridEX1.Columns("sel").Index) = 0
  'Case Is = GridEX1.Columns("DatosAdic").Index
  Case Is = GridEX1.Columns("Otros").Index
    Cambio_Importe "Otros"
    GridEX1.Value(GridEX1.Columns("sel").Index) = 0
  End Select
Exit Sub

Resume

Error_Handler:

  errores err.Number
   
  If ColIndex = GridEX1.Columns("Sel").Index Then
     GridEX1.Value(GridEX1.Columns("sel").Index) = 0
  End If
End Sub

Private Sub CargarDatos()
    txtObservacion.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Observaciones_Factura").Index), vbString)
'    txtSecuencia.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Cod_LugEnt").Index), vbLong)
'    txtLinea1.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Des_LugEnt").Index), vbString)
    txtCod_CondVent.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), vbString)
    txtDes_CondVent.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Condicion_Venta").Index), vbString)
    txtCartaCredito.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Num_CartaCredito").Index), vbString)
    txtImp_Flete.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Imp_Flete").Index), vbDouble)
    txtImp_Seguro.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Imp_Seguro").Index), vbDouble)
    txtImp_Descuento.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Imp_Descuento").Index), vbDouble)
    txtCod_Termino_Venta = FixNulos(GridEX1.Value(GridEX1.Columns("Cod_Termino_Venta").Index), vbString)
    txtDes_Termino_Venta = DevuelveCampo("select Des_Termino_Venta FROM CN_Termino_Venta where Cod_Termino_Venta = '" & FixNulos(GridEX1.Value(GridEX1.Columns("Cod_Termino_Venta").Index), vbString) & "'", cConnect)
    txtCod_Embarque.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Cod_Embarque").Index), vbString)
    txtDes_Embarque.Text = DevuelveCampo("select Des_Embarque FROM TG_TIPEMB where Cod_Embarque = '" & FixNulos(GridEX1.Value(GridEX1.Columns("Cod_Embarque").Index), vbString) & "'", cConnect)
    txtNom_Embarque.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Nom_Embarque").Index), vbString)
    txtPie_Pagina1.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index), vbString)
    txtPie_Pagina2.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index), vbString)
    txtCod_Vendor.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Cod_Vendor").Index), vbString)
    txtCod_Class.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Cod_Class").Index), vbString)
    txtPor_Comision.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Por_Comision").Index), vbDouble)
    
    txtRef_Embarque.Text = FixNulos(DevuelveCampo("select ref_embarque FROM TG_EMBARQUE where num_embarque = '" & FixNulos(GridEX1.Value(GridEX1.Columns("Num_Embarque").Index), vbLong) & "'", cConnect), vbString)
    
    txtImp_Desaduanaje.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Imp_Desaduanaje").Index), vbDouble)
    txtImp_Transporte_Pais_Destino.Text = FixNulos(GridEX1.Value(GridEX1.Columns("Imp_Transporte_Pais_Destino").Index), vbDouble)
    
    
    Me.fraDatosAdicionales.Visible = True
    Me.txtRef_Embarque.SetFocus
End Sub



Private Sub DatosAdic_Click()

Dim Serie As String, Nro_Factura As String, iPos, i As Integer, lvSW As Boolean

  GridEX1.Redraw = False

  lvSW = True
  
  Serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
  Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index)
  
  
  GridEX1.MoveFirst
  For i = 0 To GridEX1.RowCount
    If Serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) And Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index) Then
      If lvSW Then iPos = GridEX1.Row
      lvSW = False
        GridEX1.Value(GridEX1.Columns("Observaciones_Factura").Index) = txtObservacion.Text
        GridEX1.Value(GridEX1.Columns("Num_CartaCredito").Index) = FixNulos(txtCartaCredito.Text, vbString)
        GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index) = txtCod_CondVent.Text
        GridEX1.Value(GridEX1.Columns("Condicion_Venta").Index) = txtDes_CondVent.Text
        GridEX1.Value(GridEX1.Columns("Cod_Termino_Venta").Index) = txtCod_Termino_Venta.Text
'        GridEX1.Value(GridEX1.Columns("Des_Termino_Venta").Index) = txtDes_Termino_Venta.Text
        GridEX1.Value(GridEX1.Columns("Imp_Flete").Index) = txtImp_Flete.Text
        GridEX1.Value(GridEX1.Columns("Imp_Seguro").Index) = txtImp_Seguro.Text
        GridEX1.Value(GridEX1.Columns("Imp_Descuento").Index) = txtImp_Descuento.Text
        GridEX1.Value(GridEX1.Columns("Nom_Embarque").Index) = txtNom_Embarque.Text
        GridEX1.Value(GridEX1.Columns("cod_Embarque").Index) = txtCod_Embarque.Text
        GridEX1.Value(GridEX1.Columns("Pie_Factura1").Index) = txtPie_Pagina1.Text
        GridEX1.Value(GridEX1.Columns("Pie_Factura2").Index) = txtPie_Pagina2.Text
        GridEX1.Value(GridEX1.Columns("Cod_Vendor").Index) = txtCod_Vendor.Text
        GridEX1.Value(GridEX1.Columns("Cod_Class").Index) = txtCod_Class.Text
        GridEX1.Value(GridEX1.Columns("Num_Embarque").Index) = FixNulos(DevuelveCampo("select num_embarque FROM TG_EMBARQUE where ref_embarque = '" & txtRef_Embarque.Text & "'", cConnect), vbLong)
        GridEX1.Value(GridEX1.Columns("Por_Comision").Index) = txtPor_Comision.Text
        GridEX1.Value(GridEX1.Columns("Imp_Desaduanaje").Index) = txtImp_Desaduanaje.Text
        GridEX1.Value(GridEX1.Columns("Imp_Transporte_Pais_Destino").Index) = txtImp_Transporte_Pais_Destino.Text
    End If
    GridEX1.MoveNext
  Next i
  
  GridEX1.Row = iPos
  
  GridEX1.Redraw = True
    
  
End Sub


Private Sub GridEX1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)

If Left(Cbo_Almacen, 2) = "TT" Then
  Select Case ColIndex
    Case Is = GridEX1.Columns("Ser_Factura").Index
      If Trim(GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)) = "" Then GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) = ""
      Cancel = False
    Case Is = GridEX1.Columns("Num_Factura").Index
      If Trim(GridEX1.Value(GridEX1.Columns("Num_Factura").Index)) = "" Then GridEX1.Value(GridEX1.Columns("Num_Factura").Index) = ""
      Cancel = False
    Case Is = GridEX1.Columns("SEL").Index
      Cancel = False
'    Case Is = GridEX1.Columns("SEL2").Index
'      Cancel = False
    Case Is = GridEX1.Columns("Pre_Unitario").Index
      Cancel = False
    Case Is = GridEX1.Columns("Condicion_Venta").Index
      Cancel = False
    Case Is = GridEX1.Columns("Moneda").Index
      Cancel = False
   Case Is = GridEX1.Columns("Gastos Financieros").Index
      Cancel = False
   Case Is = GridEX1.Columns("Otros").Index
      Cancel = False
   Case Is = GridEX1.Columns("Kgs_a_Facturar").Index
      Cancel = False
   Case Is = GridEX1.Columns("Und").Index
      Cancel = False
'   Case Is = GridEX1.Columns("DatosAdic").Index
'      Cancel = False
'      CargarDatos
   Case Else
      Cancel = True
   End Select
Else
  Select Case ColIndex
    Case Is = GridEX1.Columns("Ser_Factura").Index
      If Trim(GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)) = "" Then GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) = ""
      Cancel = False
    Case Is = GridEX1.Columns("Num_Factura").Index
      If Trim(GridEX1.Value(GridEX1.Columns("Num_Factura").Index)) = "" Then GridEX1.Value(GridEX1.Columns("Num_Factura").Index) = ""
      Cancel = False
    Case Is = GridEX1.Columns("SEL").Index
      Cancel = False
'    Case Is = GridEX1.Columns("SEL2").Index
'      Cancel = False
    Case Is = GridEX1.Columns("Pre_Unitario").Index
      Cancel = False
    Case Is = GridEX1.Columns("Condicion_Venta").Index
      Cancel = False
    Case Is = GridEX1.Columns("Moneda").Index
      Cancel = False
   Case Is = GridEX1.Columns("Gastos Financieros").Index
      Cancel = False
   Case Is = GridEX1.Columns("Und").Index
      Cancel = False
   Case Else
      Cancel = True
    End Select
End If
  
End Sub

Private Sub GridEX1_Click()

'On Error Resume Next
    Dim ColIndex As Long
    Dim oRowData As JSRowData
    Dim SGRUPO As String
    Dim iRow As Long
    Dim i As Long
    Dim sCaptionGroup As String
    
    bCargaGRid = True
    
        If GridEX1.RowCount > 0 Then
        ColIndex = GridEX1.Col
        
        If Not GridEX1.IsGroupItem(GridEX1.Row) Then
            If UCase(GridEX1.Columns(ColIndex).Key) = "SEL" Then
                bClickColSelec = True
                SendKeys "{ENTER}"
            End If
'             If UCase(GridEX1.Columns(ColIndex).Key) = "SEL2" Then
'                bClickColSelec = True
'                SendKeys "{ENTER}"
'            End If
        Else
            If GridEX1.IsGroupItem(GridEX1.Row) Then
            End If
        End If
    End If
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
      lbDesTela.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Tela").Index)), "", GridEX1.Value(GridEX1.Columns("Tela").Index))
      lbComb.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Comb").Index)), "", GridEX1.Value(GridEX1.Columns("Comb").Index))
      lbCalidad.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Calidad").Index)), "", GridEX1.Value(GridEX1.Columns("Calidad").Index))
      lbRollos.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Numero_Rollos").Index)), "", GridEX1.Value(GridEX1.Columns("Numero_Rollos").Index))
      If lbCod_Color.Visible Then lbDes_Color.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Color").Index)), "", GridEX1.Value(GridEX1.Columns("Color").Index))
      lbGuia.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("nro_Guia").Index)), "", GridEX1.Value(GridEX1.Columns("nro_Guia").Index))
      lbObservacion.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Observaciones").Index)), "", GridEX1.Value(GridEX1.Columns("Observaciones").Index))
    End If
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)

Dim strGroupCaption As String

If RowBuffer.RowType = jgexRowTypeGroupHeader Then
    strGroupCaption = RTrim(RowBuffer.GroupCaption) & " (" & RowBuffer.RecordCount & " Documentos " & "" & ") "
    RowBuffer.GroupCaption = strGroupCaption
End If

End Sub

Private Sub MuestraSubTotales()
Dim colTemp As JSColumn

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("Moneda")
colTemp.AggregateFunction = jgexAggregateNone
colTemp.TotalRowPrefix = "SUB TOTAL "

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("Kgs_Movimiento")
colTemp.AggregateFunction = jgexSum
colTemp.TotalRowPrefix = ""

GridEX1.GroupFooterStyle = jgexTotalsGroupFooter
Set colTemp = GridEX1.Columns("Monto Despacho")
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
Dim sSQL As String
Dim aMess(4), i As Integer

GridEX1.MoveFirst

For i = 0 To GridEX1.RowCount

  If GridEX1.Value(GridEX1.Columns("SEL").Index) Then
  
    If Left(Cbo_Almacen, 2) = "31" Then
                       
     sSQL = "Ventas_Cambio_Estado_DocAlm_Tela_Tenida '$','$','$','$','$',$,'$',$,$,'$','$','$',$,'$','$'"
      
      sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
                       GridEX1.Value(GridEX1.Columns("num_movstk").Index), _
                       GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), _
                       GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
                       GridEX1.Value(GridEX1.Columns("Gastos Financieros").Index), _
                       GridEX1.Value(GridEX1.Columns("Otros").Index), _
                       GridEX1.Value(GridEX1.Columns("Ser_ordcomp").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_ordcomp").Index), _
                       GridEX1.Value(GridEX1.Columns("Sec_OrdComp").Index), _
                       GridEX1.Value(GridEX1.Columns("Kgs_a_Facturar").Index), _
                       GridEX1.Value(GridEX1.Columns("Num_Secuencia").Index), _
                       GridEX1.Value(GridEX1.Columns("Und").Index))


      ExecuteCommandSQL cConnect, sSQL
      
    ElseIf Left(Cbo_Almacen, 2) = "T1" Then

        sSQL = "Ventas_Cambio_Estado_DocAlm_Tela_Cruda '$','$','$','$','$',$ , '$' , $ ,$,'$','$','$','$'"
      
        sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
                         GridEX1.Value(GridEX1.Columns("num_movstk").Index), _
                         GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
                         GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
                         GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), _
                         GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
                         GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
                         GridEX1.Value(GridEX1.Columns("Gastos Financieros").Index), _
                         GridEX1.Value(GridEX1.Columns("Otros").Index), _
                         GridEX1.Value(GridEX1.Columns("Ser_ordcomp").Index), _
                         GridEX1.Value(GridEX1.Columns("Cod_ordcomp").Index), _
                         GridEX1.Value(GridEX1.Columns("Sec_OrdComp").Index), _
                         GridEX1.Value(GridEX1.Columns("Und").Index))
        ExecuteCommandSQL cConnect, sSQL
        
        
    
    
        ElseIf Left(Cbo_Almacen, 2) = "T8" Or Left(Cbo_Almacen, 2) = "T7" Then
                       
        sSQL = "Ventas_Cambio_Estado_DocAlm_Tejeduria '$','$','$','$','$',$,'$',$,$,'$','$','$',$,'$','$'"
      
        sSQL = VBsprintf(sSQL, Left(Cbo_Almacen, 2), _
                       GridEX1.Value(GridEX1.Columns("num_movstk").Index), _
                       GridEX1.Value(GridEX1.Columns("Ser_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Num_Factura").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index), _
                       GridEX1.Value(GridEX1.Columns("Pre_Unitario").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index), _
                       GridEX1.Value(GridEX1.Columns("Gastos Financieros").Index), _
                       GridEX1.Value(GridEX1.Columns("Otros").Index), _
                       GridEX1.Value(GridEX1.Columns("Ser_ordcomp").Index), _
                       GridEX1.Value(GridEX1.Columns("Cod_ordcomp").Index), _
                       GridEX1.Value(GridEX1.Columns("Sec_OrdComp").Index), _
                       GridEX1.Value(GridEX1.Columns("Kgs_a_Facturar").Index), _
                       GridEX1.Value(GridEX1.Columns("Num_Secuencia").Index), _
                       GridEX1.Value(GridEX1.Columns("Und").Index))


        ExecuteCommandSQL cConnect, sSQL
      
    
    End If
    
  End If

  GridEX1.MoveNext

Next i

If Left(Cbo_Almacen, 2) = "31" Then
  ExecuteCommandSQL cConnect, "Ventas_Genera_Docum_Autorizados_Tela_Tenida '" & vusu & "','" & Left(Cbo_Almacen, 2) & "'"
ElseIf Left(Cbo_Almacen, 2) = "T1" Then
  ExecuteCommandSQL cConnect, "Ventas_Genera_Docum_Autorizados_Tela_Cruda '" & vusu & "','" & Left(Cbo_Almacen, 2) & "'"
ElseIf Left(Cbo_Almacen, 2) = "T8" Or Left(Cbo_Almacen, 2) = "T7" Then
  ExecuteCommandSQL cConnect, "Ventas_Genera_Docum_Autorizados_Tejeduria '" & vusu & "','" & Left(Cbo_Almacen, 2) & "'"
  
  
End If

Mensaje kMESSAGE_INF_PROCESS_SATISFACTO

BUSCAR
 
Exit Sub
Resume
errorx:
    ErrorHandler err, "Autoriza Documentos"
End Sub

Sub Cambio_Nro_Factura()

Dim Serie As String, Nro_Factura As String, iPos, i As Integer, lvSW As Boolean

  GridEX1.Redraw = False

  lvSW = True
  
  Doc = GridEX1.Value(GridEX1.Columns("Cod_Doc").Index)
  Serie = GridEX1.Value(GridEX1.Columns("Ser_Docum").Index)
  Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Docum_Ventas").Index)
  
  GridEX1.MoveFirst
  For i = 0 To GridEX1.RowCount
    If Doc = GridEX1.Value(GridEX1.Columns("Cod_Doc").Index) Then
      If lvSW Then iPos = GridEX1.Row
      lvSW = False
      GridEX1.Value(GridEX1.Columns("Ser_Docum").Index) = Serie
      GridEX1.Value(GridEX1.Columns("Nro_Docum_Ventas").Index) = Nro_Factura
    End If
    GridEX1.MoveNext
  Next i
  
  GridEX1.Row = iPos
  
  GridEX1.Redraw = True
  
  SendKeys "{TAB}"
  
End Sub

Sub Cambio_Importe(Campo As String)

Dim Fac_Cli As String, Importe As String, iPos, i As Integer, lvSW As Boolean

  GridEX1.Redraw = False

  lvSW = True
  
  Fac_Cli = GridEX1.Value(GridEX1.Columns("Fac_Cli").Index)
  Importe = GridEX1.Value(GridEX1.Columns(Campo).Index)
  
  GridEX1.MoveFirst
  For i = 0 To GridEX1.RowCount
    If Fac_Cli = GridEX1.Value(GridEX1.Columns("Fac_Cli").Index) Then
      If lvSW Then iPos = GridEX1.Row
      lvSW = False
      GridEX1.Value(GridEX1.Columns(Campo).Index) = Importe
    End If
    GridEX1.MoveNext
  Next i
  
  GridEX1.Row = iPos
  
  GridEX1.Redraw = True
  
End Sub

Private Sub GridEX2_Click()

Dim Serie As String, Nro_Factura As String, iPos, i As Integer, lvSW As Boolean

  GridEX1.Redraw = False

  lvSW = True
  
  Serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
  Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index)
  
  
  GridEX1.MoveFirst
  For i = 0 To GridEX1.RowCount
    If Serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) And Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index) Then
      If lvSW Then iPos = GridEX1.Row
      lvSW = False
      GridEX1.Value(GridEX1.Columns("Cod_CondVent").Index) = GridEX2.Value(GridEX2.Columns("Cod_CondVent").Index)
      GridEX1.Value(GridEX1.Columns("Condicion_Venta").Index) = GridEX2.Value(GridEX2.Columns("Descripcion").Index)
    End If
    GridEX1.MoveNext
  Next i
  
  GridEX1.Row = iPos
  
  GridEX1.Redraw = True
  
  SendKeys "{TAB}"
  
End Sub

Private Sub GridEX3_Click()

Dim Serie As String, Nro_Factura As String, iPos, i As Integer, lvSW As Boolean

  GridEX1.Redraw = False
  
  Serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index)
  Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index)
  lvSW = True
  GridEX1.MoveFirst
  For i = 0 To GridEX1.RowCount
    If Serie = GridEX1.Value(GridEX1.Columns("Ser_Factura").Index) And Nro_Factura = GridEX1.Value(GridEX1.Columns("Num_Factura").Index) Then
      If lvSW Then iPos = GridEX1.Row
      lvSW = False
      GridEX1.Value(GridEX1.Columns("Cod_Moneda").Index) = GridEX3.Value(GridEX3.Columns("Cod_Moneda").Index)
      GridEX1.Value(GridEX1.Columns("Moneda").Index) = GridEX3.Value(GridEX3.Columns("Descripcion").Index)
    End If
    GridEX1.MoveNext
  Next i
  
  GridEX1.Row = iPos
  
  GridEX1.Redraw = True
  
  SendKeys "{TAB}"
  
End Sub


Private Sub FillAlmacen()

Dim rstAux As ADODB.Recordset
Dim strsql As String
    
strsql = "Ventas_Ayuda_Almacenes_Tela"
         
Set rstAux = CargarRecordSetDesconectado(strsql, cConnect)
Cbo_Almacen.Clear
With rstAux
    If .RecordCount > 0 Then .MoveFirst
    Do Until .EOF
        Cbo_Almacen.AddItem !Cod_almacen & " " & !nom_almacen
        .MoveNext
    Loop
    .Close
End With
If Cbo_Almacen.ListCount > 0 Then Cbo_Almacen.ListIndex = 0
Set rstAux = Nothing
    
End Sub

'Private Sub txtCartaCredito_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        BuscaCartaCredito 1
'        SendKeys "{TAB}"
'    End If
'End Sub




'Public Sub BuscaCartaCredito(opcion As String)
'Dim rstAux As ADODB.Recordset
'    strSQL = "SELECT Num_CartaCredito , Fec_Emision " & _
'             "FROM TG_Carta_Credito " & _
'             "WHERE Cod_Cliente = '" & GridEX1.Value(GridEX1.Columns("CLIENTE").Index) & "' AND "
'
'    txtCartaCredito = Trim(txtCartaCredito)
'
'    Select Case opcion
'    Case 1: strSQL = strSQL & "Num_CartaCredito like '%" & txtCartaCredito & "%'"
'    End Select
'    strSQL = strSQL & " AND FLG_STATUS IN ('B','F','T')"
'
'    Set frmBusqGeneral3.oParent = Me
'    frmBusqGeneral3.sQuery = strSQL
'    frmBusqGeneral3.CARGAR_DATOS
'    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
'
'    frmBusqGeneral3.gexLista.Columns("Num_CartaCredito").Visible = True
'    frmBusqGeneral3.gexLista.Columns("Num_CartaCredito").Width = 2000
'    frmBusqGeneral3.gexLista.Columns("Fec_Emision").Width = 1500
'
'    frmBusqGeneral3.gexLista.Columns("Num_CartaCredito").Caption = "Carta Credito"
'    frmBusqGeneral3.gexLista.Columns("Fec_Emision").Caption = "Fec_Emision"
'
'    If frmBusqGeneral3.gexLista.RowCount > 1 Then
'        frmBusqGeneral3.Show vbModal
'    Else
'        frmBusqGeneral3.cmdAceptar.Value = True
'    End If
'
'    If codigo <> "" Then
'        txtCartaCredito = codigo
'    End If
'    Unload frmBusqGeneral3
'    Set frmBusqGeneral3 = Nothing
'
'    codigo = ""
'    Descripcion = ""
'End Sub




Private Sub txtCod_Class_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtCod_CondVent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaCondVent 1
        SendKeys "{TAB}"
    End If
End Sub

Public Sub BuscaCondVent(Opcion As String)
'Dim rstAux As ADODB.Recordset
'
'    strSQL = "SELECT Cod_CondVent, Des_CondVent FROM lg_condvent WHERE "
'
'    txtCod_CondVent = Trim(txtCod_CondVent)
'    txtDes_CondVent = Trim(txtDes_CondVent)
'
'    Select Case opcion
'    Case 1: strSQL = strSQL & "Cod_condVent like '%" & txtCod_CondVent & "%'"
'    Case 2: strSQL = strSQL & "Des_condVent LIKE '%" & txtDes_CondVent & "%'"
'    End Select
'
'    Set frmBusqGeneral3.oParent = Me
'    frmBusqGeneral3.sQuery = strSQL
'    frmBusqGeneral3.CARGAR_DATOS
'    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
'
'
'    frmBusqGeneral3.gexLista.Columns("Cod_CondVent").Width = 700
'    frmBusqGeneral3.gexLista.Columns("Des_CondVent").Width = 2000
'
'    frmBusqGeneral3.gexLista.Columns("Cod_CondVent").Caption = "Cond.Vta"
'    frmBusqGeneral3.gexLista.Columns("Des_condVent").Caption = "Descrip."
'
'    If frmBusqGeneral3.gexLista.RowCount > 1 Then
'        frmBusqGeneral3.Show vbModal
'    Else
'        frmBusqGeneral3.cmdAceptar.Value = True
'    End If
'
'    txtCod_CondVent = ""
'    txtDes_CondVent = ""
'
'    If codigo <> "" Then
'        txtCod_CondVent = codigo
'        txtDes_CondVent = Descripcion
'    End If
'    Unload frmBusqGeneral3
'    Set frmBusqGeneral3 = Nothing
'
'    codigo = ""
'    Descripcion = ""
End Sub


Public Sub BuscaLugEnt(Opcion As String)
'Dim rstAux As ADODB.Recordset
'    strSQL = "SELECT Secuencia, RTRIM(Linea1) + ' ' + RTRIM(Linea2) + " & _
'             "RTRIM(Linea3) AS Linea1 FROM TG_CLIENTE_LUGENT " & _
'             "WHERE Cod_Cliente = '" & GridEX1.Value(GridEX1.Columns("CLIENTE").Index) & "' AND "
'
'    txtSecuencia = Trim(txtSecuencia)
'    txtLinea1 = Trim(txtLinea1)
'
'    Select Case opcion
'    Case 1: strSQL = strSQL & "CONVERT(varchar(8), Secuencia) like '%" & txtSecuencia & "%'"
'    Case 2: strSQL = strSQL & "RTRIM(Linea1) + ' ' + RTRIM(Linea2) + " & _
'             "RTRIM(Linea3) LIKE '%" & txtLinea1 & "%'"
'    End Select
'
'    Set frmBusqGeneral3.oParent = Me
'    frmBusqGeneral3.SQuery = strSQL
'    frmBusqGeneral3.Cargar_Datos
'    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
'
'    frmBusqGeneral3.gexLista.Columns("Secuencia").Visible = False
'    frmBusqGeneral3.gexLista.Columns("Secuencia").Width = 570
'    frmBusqGeneral3.gexLista.Columns("Linea1").Width = 2370
'
'    frmBusqGeneral3.gexLista.Columns("Secuencia").Caption = "Secuencia"
'    frmBusqGeneral3.gexLista.Columns("Linea1").Caption = "Lug.Entr."
'
'    If frmBusqGeneral3.gexLista.RowCount > 1 Then
'        frmBusqGeneral3.Show vbModal
'    Else
'        frmBusqGeneral3.cmdAceptar.Value = True
'    End If
'
'    txtSecuencia = ""
'    txtLinea1 = ""
'
'    If codigo <> "" Then
'        txtSecuencia = codigo
'        txtLinea1 = Descripcion
'    End If
'    Unload frmBusqGeneral3
'    Set frmBusqGeneral3 = Nothing
'
'    codigo = ""
'    Descripcion = ""
End Sub

Private Sub txtCod_Embarque_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaModoTransporte 1
        SendKeys "{TAB}"
    End If
End Sub

Public Sub BuscaModoTransporte(Opcion As String)
'Dim rstAux As ADODB.Recordset
'
'    strSQL = "SELECT Cod_Embarque, Des_Embarque FROM TG_TIPEMB WHERE "
'
'    txtCod_Embarque = Trim(txtCod_Embarque)
'    txtDes_Embarque = Trim(txtDes_Embarque)
'
'    Select Case opcion
'    Case 1: strSQL = strSQL & "Cod_Embarque like '%" & txtCod_Embarque & "%'"
'    Case 2: strSQL = strSQL & "Des_Embarque LIKE '%" & txtDes_Embarque & "%'"
'    End Select
'
'    Set frmBusqGeneral3.oParent = Me
'    frmBusqGeneral3.sQuery = strSQL
'    frmBusqGeneral3.CARGAR_DATOS
'    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
'
'    frmBusqGeneral3.gexLista.Columns("Cod_Embarque").Width = 700
'    frmBusqGeneral3.gexLista.Columns("Des_Embarque").Width = 2000
'
'    frmBusqGeneral3.gexLista.Columns("Cod_Embarque").Caption = "Embarque"
'    frmBusqGeneral3.gexLista.Columns("Des_Embarque").Caption = "Descrip."
'
'    If frmBusqGeneral3.gexLista.RowCount > 1 Then
'        frmBusqGeneral3.Show vbModal
'    Else
'        frmBusqGeneral3.cmdAceptar.Value = True
'    End If
'
'    txtCod_Embarque = ""
'    txtDes_Embarque = ""
'
'    If codigo <> "" Then
'        txtCod_Embarque = codigo
'        txtDes_Embarque = Descripcion
'    End If
'    Unload frmBusqGeneral3
'    Set frmBusqGeneral3 = Nothing
'
'    codigo = ""
'    Descripcion = ""
End Sub


Private Sub txtCod_Termino_Venta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaTerminoVent 1
        SendKeys "{TAB}"
    End If
End Sub


Public Sub BuscaTerminoVent(Opcion As String)
'Dim rstAux As ADODB.Recordset
'
'    strSQL = "SELECT Cod_Termino_Venta, Des_Termino_Venta FROM CN_Termino_Venta WHERE "
'
'    txtCod_Termino_Venta = Trim(txtCod_Termino_Venta)
'    txtDes_Termino_Venta = Trim(txtDes_Termino_Venta)
'
'    Select Case opcion
'    Case 1: strSQL = strSQL & "Cod_Termino_Venta like '%" & txtCod_Termino_Venta & "%'"
'    Case 2: strSQL = strSQL & "Des_Termino_Venta LIKE '%" & txtDes_Termino_Venta & "%'"
'    End Select
'
'    Set frmBusqGeneral3.oParent = Me
'    frmBusqGeneral3.sQuery = strSQL
'    frmBusqGeneral3.CARGAR_DATOS
'    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
'
'    frmBusqGeneral3.gexLista.Columns("Cod_Termino_Venta").Width = 700
'    frmBusqGeneral3.gexLista.Columns("Des_Termino_Venta").Width = 2000
'
'    frmBusqGeneral3.gexLista.Columns("Cod_Termino_Venta").Caption = "Termino.Venta"
'    frmBusqGeneral3.gexLista.Columns("Des_Termino_Venta").Caption = "Descrip."
'
'    If frmBusqGeneral3.gexLista.RowCount > 1 Then
'        frmBusqGeneral3.Show vbModal
'    Else
'        frmBusqGeneral3.cmdAceptar.Value = True
'    End If
'
'    txtCod_Termino_Venta = ""
'    txtDes_Termino_Venta = ""
'
'    If codigo <> "" Then
'        txtCod_Termino_Venta = codigo
'        txtDes_Termino_Venta = Descripcion
'    End If
'    Unload frmBusqGeneral3
'    Set frmBusqGeneral3 = Nothing
'
'    codigo = ""
'    Descripcion = ""
End Sub


Private Sub txtCod_Vendor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtImp_Desaduanaje_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
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

Private Sub txtImp_Transporte_Pais_Destino_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtNom_embarque_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
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

Private Sub txtPor_Comision_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtRef_Embarque_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaRef_Embarque 1
        SendKeys "{TAB}"
    End If

End Sub

Public Sub BuscaRef_Embarque(Opcion As String)
'Dim rstAux As ADODB.Recordset
'Dim rsData As ADODB.Recordset
'
'    'strSQL = "SELECT Ref_Embarque , Obs_Embarque FROM TG_EMBARQUE WHERE FLG_STATUS in ('T','F') AND COD_TIPANEX = '" & GridEX1.Value(GridEX1.Columns("COD_TIPANEX").Index) & "' AND  COD_ANXO = '" & GridEX1.Value(GridEX1.Columns("COD_ANXO").Index) & "' AND COD_CLIENTE = '" & GridEX1.Value(GridEX1.Columns("CLIENTE").Index) & "' AND "
'
'    strSQL = "VN_MUESTRA_EMBARQUES_TEXTILES_VIGENTES '" & GridEX1.Value(GridEX1.Columns("CLIENTE").Index) & "'"
'
'    txtRef_Embarque = Trim(txtRef_Embarque)
'
''    Select Case opcion
''    Case 1: strSQL = strSQL & "Ref_Embarque like '%" & txtRef_Embarque & "%'"
''    End Select
'
'    Set frmBusqGeneral3.oParent = Me
'    frmBusqGeneral3.sQuery = strSQL
'    frmBusqGeneral3.CARGAR_DATOS
'    Set rstAux = frmBusqGeneral3.gexLista.ADORecordset
'
'    frmBusqGeneral3.gexLista.Columns("Ref_Embarque").Width = 1700
'    frmBusqGeneral3.gexLista.Columns("NUM_EMBARQUE").Width = 2000
'    frmBusqGeneral3.gexLista.Columns("ANO").Width = 900
'    frmBusqGeneral3.gexLista.Columns("MES").Width = 900
'
'    frmBusqGeneral3.gexLista.Columns("NUM_EMBARQUE").Caption = "Número Embarque"
'    frmBusqGeneral3.gexLista.Columns("Ref_Embarque").Caption = "Ref. Embarque"
'    frmBusqGeneral3.gexLista.Columns("ANO").Caption = "Año"
'    frmBusqGeneral3.gexLista.Columns("MES").Caption = "Mes"
'
'    If frmBusqGeneral3.gexLista.RowCount = 0 Then
'        MsgBox "Embarque no existe", 1
'                Exit Sub
'    End If
'
'    If frmBusqGeneral3.gexLista.RowCount > 1 Then
'        frmBusqGeneral3.Show vbModal
'    Else
'        frmBusqGeneral3.cmdAceptar.Value = True
'    End If
'
'    txtRef_Embarque = ""
'
'
'    If codigo <> "" Then
'        txtRef_Embarque = codigo
'        If txtRef_Embarque <> "" Then
'            strSQL = "TG_Embarques_Muestra '$','$','$','$','$','$','$'"
'            strSQL = VBsprintf(strSQL, "3", 0, txtRef_Embarque, "", "", "", "")
'            Set rsData = GetDataSet(cConnect, strSQL)
'            If Not rsData Is Nothing Then
'                Do While Not rsData.EOF
'                    If RTrim(txtCod_Termino_Venta) = "" Then
'                        txtCod_Termino_Venta = FixNulos(rsData("Cod_Termino_venta").Value, vbString)
'                        txtDes_Termino_Venta = FixNulos(rsData("Des_Termino_Venta").Value, vbString)
'                    End If
'                    If RTrim(txtCod_Embarque.Text) = "" Then
'                        txtCod_Embarque.Text = FixNulos(rsData("Cod_Embarque").Value, vbString)
'                        txtDes_Embarque.Text = FixNulos(rsData("Des_Embarque").Value, vbString)
'                    End If
'                    If RTrim(txtNom_Embarque.Text) = "" Then
'                        txtNom_Embarque.Text = FixNulos(rsData("Nom_Embarque").Value, vbString)
'                    End If
'
'                    rsData.MoveNext
'                Loop
'                rsData.Close
'            End If
'            Set rsData = Nothing
'
'        End If
'    End If
'    Unload frmBusqGeneral3
'    Set frmBusqGeneral3 = Nothing
'
'    codigo = ""
'    Descripcion = ""
End Sub



Private Sub txtSecuencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        
        BuscaLugEnt 1
        SendKeys "{TAB}"
    End If
End Sub
