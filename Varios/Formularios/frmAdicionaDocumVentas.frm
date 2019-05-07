VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdicionaDocumVentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VENTA DE TELA ACABADA"
   ClientHeight    =   8895
   ClientLeft      =   405
   ClientTop       =   690
   ClientWidth     =   17595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   17595
   Begin VB.Frame fraSelGuias 
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00000000&
      Height          =   5400
      Left            =   960
      TabIndex        =   76
      Top             =   2040
      Width           =   14535
      Begin VB.CommandButton CmdCerrarGuias 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13440
         TabIndex        =   82
         Top             =   5000
         Width           =   975
      End
      Begin VB.CommandButton cmdAsigna 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   81
         Top             =   5000
         Width           =   975
      End
      Begin VB.TextBox txtNumeroGuia 
         BackColor       =   &H00C0FFFF&
         Height          =   350
         Left            =   6360
         TabIndex        =   80
         Top             =   120
         Width           =   2055
      End
      Begin VB.CommandButton cmdDesAsigna 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   79
         Top             =   5000
         Width           =   975
      End
      Begin VB.ComboBox cboAlmacen 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   150
         Width           =   3975
      End
      Begin VB.TextBox txtSerieGuia 
         BackColor       =   &H00C0FFFF&
         Height          =   350
         Left            =   5280
         TabIndex        =   77
         Top             =   120
         Width           =   1095
      End
      Begin GridEX20.GridEX grxListaGuiaPendientes 
         Height          =   4455
         Left            =   45
         TabIndex        =   83
         Top             =   480
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   7858
         Version         =   "2.0"
         AllowRowSizing  =   -1  'True
         AutomaticSort   =   -1  'True
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         GridLineStyle   =   2
         HideSelection   =   2
         MethodHoldFields=   -1  'True
         GroupByBoxInfoText=   "Arrastra la cabecera de una columna aquí para agruparlo por esa misma columna"
         GroupByBoxVisible=   0   'False
         HeaderFontName  =   "Verdana"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   6.75
         HeaderFontWeight=   700
         FontName        =   "Tahoma"
         ColumnHeaderHeight=   270
         IntProp1        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAdicionaDocumVentas.frx":0000
         Column(2)       =   "frmAdicionaDocumVentas.frx":00C8
         FormatStylesCount=   8
         FormatStyle(1)  =   "frmAdicionaDocumVentas.frx":016C
         FormatStyle(2)  =   "frmAdicionaDocumVentas.frx":0294
         FormatStyle(3)  =   "frmAdicionaDocumVentas.frx":0344
         FormatStyle(4)  =   "frmAdicionaDocumVentas.frx":03F8
         FormatStyle(5)  =   "frmAdicionaDocumVentas.frx":04D0
         FormatStyle(6)  =   "frmAdicionaDocumVentas.frx":0588
         FormatStyle(7)  =   "frmAdicionaDocumVentas.frx":0668
         FormatStyle(8)  =   "frmAdicionaDocumVentas.frx":06F8
         ImageCount      =   0
         PrinterProperties=   "frmAdicionaDocumVentas.frx":080C
      End
      Begin GridEX20.GridEX grxListaGuiasSeleccionadas 
         Height          =   4455
         Left            =   8520
         TabIndex        =   84
         Top             =   480
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   7858
         Version         =   "2.0"
         AllowRowSizing  =   -1  'True
         AutomaticSort   =   -1  'True
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         GridLineStyle   =   2
         HideSelection   =   2
         MethodHoldFields=   -1  'True
         GroupByBoxInfoText=   "Arrastra la cabecera de una columna aquí para agruparlo por esa misma columna"
         GroupByBoxVisible=   0   'False
         HeaderFontName  =   "Verdana"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   6.75
         HeaderFontWeight=   700
         FontName        =   "Tahoma"
         ColumnHeaderHeight=   270
         IntProp1        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAdicionaDocumVentas.frx":09E4
         Column(2)       =   "frmAdicionaDocumVentas.frx":0AAC
         FormatStylesCount=   8
         FormatStyle(1)  =   "frmAdicionaDocumVentas.frx":0B50
         FormatStyle(2)  =   "frmAdicionaDocumVentas.frx":0C78
         FormatStyle(3)  =   "frmAdicionaDocumVentas.frx":0D28
         FormatStyle(4)  =   "frmAdicionaDocumVentas.frx":0DDC
         FormatStyle(5)  =   "frmAdicionaDocumVentas.frx":0EB4
         FormatStyle(6)  =   "frmAdicionaDocumVentas.frx":0F6C
         FormatStyle(7)  =   "frmAdicionaDocumVentas.frx":104C
         FormatStyle(8)  =   "frmAdicionaDocumVentas.frx":10DC
         ImageCount      =   0
         PrinterProperties=   "frmAdicionaDocumVentas.frx":11F0
      End
      Begin VB.Label Label14 
         BackColor       =   &H0080C0FF&
         Caption         =   "<ENTER> para buscar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   13560
         TabIndex        =   87
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "ALMACEN:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   75
         TabIndex        =   86
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080C0FF&
         Caption         =   "GR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   85
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.CheckBox chkImpresionDirecta 
      Caption         =   "IMP. DIRECTA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9120
      TabIndex        =   74
      Top             =   8520
      Width           =   1455
   End
   Begin VB.TextBox txtCodigo_Producto 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   350
      Left            =   840
      TabIndex        =   63
      Top             =   8400
      Width           =   3375
   End
   Begin VB.Frame fraUbicacion 
      Height          =   615
      Left            =   0
      TabIndex        =   54
      Top             =   360
      Width           =   17415
      Begin VB.TextBox txtCod_Almacen 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   15240
         MaxLength       =   4
         TabIndex        =   72
         Top             =   240
         Width           =   465
      End
      Begin VB.TextBox txtDes_Almacen 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   15720
         TabIndex        =   71
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtCod_Vendedor 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   11640
         TabIndex        =   68
         Top             =   240
         Width           =   705
      End
      Begin VB.TextBox txtCod_Caja 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   8400
         TabIndex        =   67
         Top             =   240
         Width           =   465
      End
      Begin VB.TextBox txtCod_Tienda 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   4800
         TabIndex        =   66
         Top             =   240
         Width           =   465
      End
      Begin VB.TextBox txtCod_Fabrica 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1005
         TabIndex        =   65
         Top             =   240
         Width           =   465
      End
      Begin VB.TextBox txtDes_Tienda 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   5250
         TabIndex        =   61
         Top             =   240
         Width           =   2625
      End
      Begin VB.TextBox txtDes_Fabrica 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1485
         TabIndex        =   59
         Top             =   240
         Width           =   2625
      End
      Begin VB.TextBox txtDes_Caja 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   8805
         TabIndex        =   58
         Top             =   240
         Width           =   1905
      End
      Begin VB.TextBox txtDes_Vendedor 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   12285
         TabIndex        =   55
         Top             =   240
         Width           =   2505
      End
      Begin VB.Label Label9 
         Caption         =   "ALM."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   14800
         TabIndex        =   73
         Top             =   255
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "TIENDA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4125
         TabIndex        =   62
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "EMPRESA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "CAJA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7965
         TabIndex        =   57
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "VENDEDOR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   10725
         TabIndex        =   56
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.Frame frMain 
      Height          =   1080
      Left            =   0
      TabIndex        =   27
      Top             =   960
      Width           =   17415
      Begin VB.OptionButton Option1 
         Caption         =   "TODOS"
         Height          =   195
         Index           =   2
         Left            =   16320
         TabIndex        =   92
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "PARTIDA"
         Height          =   195
         Index           =   1
         Left            =   15120
         TabIndex        =   91
         Top             =   840
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "IND."
         Height          =   195
         Index           =   0
         Left            =   14280
         TabIndex        =   90
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtiva 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   12960
         Locked          =   -1  'True
         TabIndex        =   88
         Top             =   705
         Width           =   975
      End
      Begin VB.TextBox TxtTipo_Cambio 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   11640
         TabIndex        =   69
         Top             =   705
         Width           =   855
      End
      Begin VB.Frame frReferencia 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         TabIndex        =   41
         Top             =   5040
         Visible         =   0   'False
         Width           =   7815
      End
      Begin VB.TextBox txtCod_TipAne 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1005
         MaxLength       =   4
         TabIndex        =   40
         Text            =   "C"
         Top             =   420
         Width           =   465
      End
      Begin VB.TextBox txtDes_ConPag 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1485
         TabIndex        =   39
         Top             =   705
         Width           =   4425
      End
      Begin VB.TextBox txtCod_ConPag 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1005
         MaxLength       =   4
         TabIndex        =   38
         Top             =   705
         Width           =   465
      End
      Begin VB.TextBox txtDes_Moneda 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   9040
         TabIndex        =   37
         Top             =   120
         Width           =   1650
      End
      Begin VB.TextBox txtCod_Moneda 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   8445
         MaxLength       =   4
         TabIndex        =   36
         Top             =   120
         Width           =   600
      End
      Begin VB.TextBox txtDes_TipAne 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1485
         TabIndex        =   35
         Top             =   420
         Width           =   4425
      End
      Begin VB.TextBox txtNum_Docum 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   5850
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   33
         Top             =   120
         Width           =   2020
      End
      Begin VB.TextBox txtDes_TipDoc 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1485
         TabIndex        =   32
         Top             =   120
         Width           =   2625
      End
      Begin VB.TextBox txtCod_TipDoc 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1005
         MaxLength       =   4
         TabIndex        =   31
         Top             =   120
         Width           =   465
      End
      Begin VB.TextBox txtSer_Docum 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   4770
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   30
         Top             =   120
         Width           =   1080
      End
      Begin VB.TextBox txtDes_TipVenta 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   12240
         TabIndex        =   28
         Top             =   120
         Width           =   3855
      End
      Begin MSComCtl2.DTPicker dtpFec_Emision 
         Height          =   285
         Left            =   11640
         TabIndex        =   42
         Top             =   405
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   128
         Format          =   70254593
         CurrentDate     =   38182
      End
      Begin MSComCtl2.DTPicker dtpFec_Registro 
         Height          =   285
         Left            =   14760
         TabIndex        =   43
         Top             =   405
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   128
         Format          =   70254593
         CurrentDate     =   38182
      End
      Begin VB.TextBox txtCod_TipVenta 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   11640
         MaxLength       =   4
         TabIndex        =   29
         Top             =   120
         Width           =   600
      End
      Begin VB.TextBox txtNum_Ruc 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   6480
         MaxLength       =   11
         TabIndex        =   34
         Top             =   420
         Width           =   4220
      End
      Begin VB.Label Label15 
         Caption         =   "IVA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12600
         TabIndex        =   89
         Top             =   735
         Width           =   375
      End
      Begin VB.Label Label27 
         Caption         =   "T./C"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11280
         TabIndex        =   70
         Top             =   735
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "F. PAGO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "MON"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7920
         TabIndex        =   52
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "EMISION"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10920
         TabIndex        =   51
         Top             =   405
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "REGISTRO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13920
         TabIndex        =   50
         Top             =   405
         Width           =   855
      End
      Begin VB.Label Label28 
         Caption         =   "R.U.C"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   49
         Top             =   420
         Width           =   375
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "CLIENTE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   360
         TabIndex        =   48
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label5 
         Caption         =   "Numero :"
         Height          =   255
         Left            =   9390
         TabIndex        =   47
         Top             =   375
         Width           =   15
      End
      Begin VB.Label Label12 
         Caption         =   "DOCUMENTO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   40
         TabIndex        =   46
         Top             =   135
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "TIPO VENTA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10680
         TabIndex        =   45
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "S/N"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4200
         TabIndex        =   44
         Top             =   120
         Width           =   285
      End
   End
   Begin VB.TextBox txtTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "R E G I S T R O    D E   V E N T A S"
      Top             =   0
      Width           =   17415
   End
   Begin VB.TextBox txt_total 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   16200
      TabIndex        =   11
      Top             =   8400
      Width           =   1215
   End
   Begin VB.TextBox txt_igv 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   14520
      TabIndex        =   9
      Top             =   8400
      Width           =   1095
   End
   Begin VB.TextBox txt_subtotal 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   12960
      TabIndex        =   7
      Top             =   8400
      Width           =   1215
   End
   Begin VB.TextBox txt_descto 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   11085
      TabIndex        =   5
      Top             =   8400
      Width           =   975
   End
   Begin VB.Frame FraProductos 
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00000000&
      Height          =   5400
      Left            =   960
      TabIndex        =   3
      Top             =   2040
      Width           =   14535
      Begin VB.CheckBox chkTodos 
         BackColor       =   &H0080C0FF&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   13560
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdCerrarBusProductos 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13440
         TabIndex        =   23
         Top             =   5000
         Width           =   975
      End
      Begin VB.CommandButton cmdBusAgregarTelas 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12480
         TabIndex        =   22
         Top             =   5000
         Width           =   975
      End
      Begin VB.CommandButton cmdBusLimpiarCaja 
         BackColor       =   &H0080C0FF&
         Caption         =   "Borrar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   50
         TabIndex        =   21
         Top             =   120
         Width           =   800
      End
      Begin VB.TextBox txtBus_Des_Color 
         BackColor       =   &H00C0FFFF&
         Height          =   350
         Left            =   9960
         TabIndex        =   19
         Top             =   120
         Width           =   2535
      End
      Begin VB.TextBox txtBus_Cod_ordtra 
         BackColor       =   &H00C0FFFF&
         Height          =   350
         Left            =   4440
         TabIndex        =   17
         Top             =   120
         Width           =   1815
      End
      Begin VB.TextBox txtDescripcion_Producto 
         BackColor       =   &H00C0FFFF&
         Height          =   350
         Left            =   6720
         TabIndex        =   15
         Top             =   120
         Width           =   2655
      End
      Begin VB.TextBox txtBus_Codigo_RolloTinto 
         BackColor       =   &H00C0FFFF&
         Height          =   350
         Left            =   1440
         TabIndex        =   13
         Top             =   120
         Width           =   2295
      End
      Begin GridEX20.GridEX GrxProductos 
         Height          =   4450
         Left            =   45
         TabIndex        =   4
         Top             =   480
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   7858
         Version         =   "2.0"
         AllowRowSizing  =   -1  'True
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         GridLineStyle   =   2
         HideSelection   =   2
         MethodHoldFields=   -1  'True
         GroupByBoxInfoText=   "Arrastra la cabecera de una columna aquí para agruparlo por esa misma columna"
         GroupByBoxVisible=   0   'False
         HeaderFontName  =   "Verdana"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   6.75
         HeaderFontWeight=   700
         FontName        =   "Tahoma"
         ColumnHeaderHeight=   270
         IntProp1        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAdicionaDocumVentas.frx":13C8
         Column(2)       =   "frmAdicionaDocumVentas.frx":1490
         FormatStylesCount=   8
         FormatStyle(1)  =   "frmAdicionaDocumVentas.frx":1534
         FormatStyle(2)  =   "frmAdicionaDocumVentas.frx":165C
         FormatStyle(3)  =   "frmAdicionaDocumVentas.frx":170C
         FormatStyle(4)  =   "frmAdicionaDocumVentas.frx":17C0
         FormatStyle(5)  =   "frmAdicionaDocumVentas.frx":1898
         FormatStyle(6)  =   "frmAdicionaDocumVentas.frx":1950
         FormatStyle(7)  =   "frmAdicionaDocumVentas.frx":1A30
         FormatStyle(8)  =   "frmAdicionaDocumVentas.frx":1AC0
         ImageCount      =   0
         PrinterProperties=   "frmAdicionaDocumVentas.frx":1BD4
      End
      Begin VB.Label Label42 
         BackColor       =   &H0080C0FF&
         Caption         =   "<ENTER> para buscar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   12600
         TabIndex        =   24
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label41 
         BackColor       =   &H0080C0FF&
         Caption         =   "COLOR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9360
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label40 
         BackColor       =   &H0080C0FF&
         Caption         =   "PARTIDA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label34 
         BackColor       =   &H0080C0FF&
         Caption         =   "TELA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label39 
         BackColor       =   &H0080C0FF&
         Caption         =   "CODIGO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   17415
      Begin GridEX20.GridEX grxDatos 
         Height          =   5955
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   17175
         _ExtentX        =   30295
         _ExtentY        =   10504
         Version         =   "2.0"
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         GridLineStyle   =   2
         HideSelection   =   2
         MethodHoldFields=   -1  'True
         GroupByBoxInfoText=   "Arrastra la cabecera de una columna aquí para agruparlo por esa misma columna"
         GroupByBoxVisible=   0   'False
         HeaderFontName  =   "Verdana"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   6.75
         HeaderFontWeight=   700
         FontName        =   "Tahoma"
         ColumnHeaderHeight=   270
         IntProp1        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAdicionaDocumVentas.frx":1DAC
         Column(2)       =   "frmAdicionaDocumVentas.frx":1E74
         FormatStylesCount=   9
         FormatStyle(1)  =   "frmAdicionaDocumVentas.frx":1F18
         FormatStyle(2)  =   "frmAdicionaDocumVentas.frx":2040
         FormatStyle(3)  =   "frmAdicionaDocumVentas.frx":20F0
         FormatStyle(4)  =   "frmAdicionaDocumVentas.frx":21A4
         FormatStyle(5)  =   "frmAdicionaDocumVentas.frx":227C
         FormatStyle(6)  =   "frmAdicionaDocumVentas.frx":2334
         FormatStyle(7)  =   "frmAdicionaDocumVentas.frx":2414
         FormatStyle(8)  =   "frmAdicionaDocumVentas.frx":24A4
         FormatStyle(9)  =   "frmAdicionaDocumVentas.frx":25DC
         ImageCount      =   0
         PrinterProperties=   "frmAdicionaDocumVentas.frx":26F0
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   6720
      TabIndex        =   0
      Top             =   8280
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   900
      Custom          =   $"frmAdicionaDocumVentas.frx":28C8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   12
   End
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   510
      Left            =   4320
      TabIndex        =   75
      Top             =   8280
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   900
      Custom          =   $"frmAdicionaDocumVentas.frx":295B
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   12
   End
   Begin VB.Label Label33 
      Caption         =   "CODIGO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   64
      Top             =   8400
      Width           =   615
   End
   Begin VB.Label Label38 
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15600
      TabIndex        =   12
      Top             =   8520
      Width           =   735
   End
   Begin VB.Label Label37 
      Caption         =   "IGV"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14160
      TabIndex        =   10
      Top             =   8520
      Width           =   375
   End
   Begin VB.Label Label36 
      Caption         =   "SUBTOTAL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12120
      TabIndex        =   8
      Top             =   8520
      Width           =   855
   End
   Begin VB.Label Label35 
      Caption         =   "KILOS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10605
      TabIndex        =   6
      Top             =   8520
      Width           =   615
   End
End
Attribute VB_Name = "frmAdicionaDocumVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CODIGO As String, Descripcion As String, StrOption As String, strNum_Corre As String, strCod_Anxo As String
Public rsFactura As New ADODB.Recordset
Dim StrSQL As String
Dim bClickColSelec As Boolean
Dim errorx As String
Dim rstAux As ADODB.Recordset
Dim sTit As String
Public flg_Tiene_guias_asignadas As String
Public fila_seleccionada As Double
Private rsDocResumen As New ADODB.Recordset
Private ind_aplica_precio As Integer

Private Declare Function GetSystemMenu Lib "user32" _
    (ByVal hwnd As Long, _
     ByVal bRevert As Long) As Long

Private Declare Function RemoveMenu Lib "user32" _
    (ByVal hMenu As Long, _
     ByVal nPosition As Long, _
     ByVal wFlags As Long) As Long
     
Private Const MF_BYPOSITION = &H400&
Public iva As Double
Public Function DisableCloseButton(frm As Form) As Boolean

'PURPOSE: Removes X button from a form
'EXAMPLE: DisableCloseButton Me
'RETURNS: True if successful, false otherwise
'NOTES:   Also removes Exit Item from
'         Control Box Menu

    Dim lHndSysMenu As Long
    Dim lAns1 As Long, lAns2 As Long
    
    lHndSysMenu = GetSystemMenu(frm.hwnd, 0)
    'remove close button
    lAns1 = RemoveMenu(lHndSysMenu, 6, MF_BYPOSITION)
   'Remove seperator bar
    lAns2 = RemoveMenu(lHndSysMenu, 5, MF_BYPOSITION)
    'Return True if both calls were successful
    DisableCloseButton = (lAns1 <> 0 And lAns2 <> 0)

End Function


Sub Busca_Opcion(strCampo1 As String, strCampo2 As String, strTabla As String, txtCod As TextBox, txtDes As TextBox, Opcion As Integer)
On Error GoTo fin
Dim rstAux As ADODB.Recordset
    StrSQL = "Select " & strCampo1 & " AS Cod," & strCampo2 & " as Descripcion from " & strTabla

    txtCod = Trim(txtCod)
    txtDes = Trim(txtDes)
    Select Case Opcion
    Case 1: StrSQL = StrSQL & strCampo1 & " like '%" & txtCod & "%'"
    Case 2: StrSQL = StrSQL & strCampo2 & " like '%" & txtDes & "%'"
    End Select
    txtCod = ""
    txtDes = ""
    fila_seleccionada = 0
    
    With frmBusqGeneral
        Set .oParent = Me
        .SQuery = StrSQL
        .Cargar_Datos
        
        CODIGO = ".."
        Set rstAux = .gexList.ADORecordset
        'If rstAux.RecordCount > 1 Then
        .Show vbModal
        
        If fila_seleccionada > 0 And rstAux.RecordCount > 0 Then
            rstAux.AbsolutePosition = fila_seleccionada
            txtCod = Trim(rstAux!cod)
            txtDes = Trim(rstAux!Descripcion)
            'Select Case Opcion
            'Case 1: SendKeys "{TAB}": SendKeys "{TAB}"
            'Case 2: SendKeys "{TAB}"
            'End Select
        Else
            txtCod = ""
            txtDes = ""
            SendKeys "{TAB}"
        End If
        
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento (" & Opcion & ")"
End Sub

Private Sub chkTodos_Click()
On Error GoTo fin
     If GrxProductos.RowCount = 0 Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    Dim Valor As Boolean
    Dim i As Long

    GrxProductos.Update
    Set rs = GrxProductos.ADORecordset
    rs.MoveFirst
    Do While Not rs.EOF
        
    If chktodos.Value = Checked Then
        If rs("stock") > 0 Then
            rs("cant") = rs("stock")
            rs("total") = rs("stock") * rs("precio")
        End If
    Else
            rs("cant") = 0
    End If
        rs.MoveNext
    Loop
   
    rs.MoveFirst
    rs.Update
    Set GrxProductos.ADORecordset = rs
    Call ConfiguraGrilla_productos
Exit Sub
Resume
fin:
On Error Resume Next
Set rs = Nothing
MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
"Edicionar Producto "

End Sub

Private Sub CmdCerrarGuias_Click()
fraSelGuias.Visible = False
    flg_Tiene_guias_asignadas = "N"
If DevuelveCampo(" select count(*) from tx_movistk where ser_docum_ventas<>'' AND  ser_docum_ventas='" & Trim(txtSer_Docum.Text) & "' AND num_docum_ventas <>'' and num_docum_ventas='" & Trim(txtNum_Docum.Text) & "' ", cConnect) > 0 Then
    flg_Tiene_guias_asignadas = "S"
End If
Call adicionarProductoDesdeDetalleGuia

End Sub

Private Sub cmdDesAsigna_Click()
On Error GoTo fin
If grxListaGuiasSeleccionadas.RowCount = 0 Then Exit Sub
StrSQL = "CN_ASIGNA_GUIA_FACTURA '" & grxListaGuiasSeleccionadas.Value(grxListaGuiasSeleccionadas.Columns("cod_almacen").Index) & "','" & grxListaGuiasSeleccionadas.Value(grxListaGuiasSeleccionadas.Columns("num_movstk").Index) & "','',''"
Call ExecuteCommandSQL(cConnect, StrSQL)

Call buscalistaGuiasPendientes
Call buscalistaGuiasSeleccionadas
sTit = "Importante"

Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, sTit
    
End Sub
Private Sub cmdAsigna_Click()
On Error GoTo fin
If grxListaGuiaPendientes.RowCount = 0 Then Exit Sub

    StrSQL = "CN_ASIGNA_GUIA_FACTURA '" & grxListaGuiaPendientes.Value(grxListaGuiaPendientes.Columns("cod_almacen").Index) & "','" & grxListaGuiaPendientes.Value(grxListaGuiaPendientes.Columns("num_movstk").Index) & "','" & Trim(txtSer_Docum.Text) & "','" & Trim(txtNum_Docum.Text) & "'"
    Call ExecuteCommandSQL(cConnect, StrSQL)
    
    Call buscalistaGuiasPendientes
    Call buscalistaGuiasSeleccionadas

Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, sTit

End Sub

Private Sub cmdBusAgregarTelas_Click()
'SendKeys "{ENTER}"
If GrxProductos.RowCount <= 0 Then Exit Sub
Call adicionarProductoMasivo

If grxDatos.RowCount <= 0 Then
    Call buscaDetalle_factura
End If
FraProductos.Visible = False
Set GrxProductos.ADORecordset = Nothing
End Sub

Private Sub cmdBusLimpiarCaja_Click()
Call limpiarCajasBusqueda

End Sub
'Private Sub cmdBusquedaProductos_Click()
'FraProductos.Visible = True
'limpiarCajasBusqueda
'End Sub
Private Sub limpiarCajasBusqueda()
    txtBus_Codigo_RolloTinto.Text = ""
    txtBus_Cod_ordtra.Text = ""
    txtBus_Des_Color.Text = ""
    txtDescripcion_Producto.Text = ""
End Sub
Private Sub cmdCerrarBusProductos_Click()
FraProductos.Visible = False
Set GrxProductos.ADORecordset = Nothing
End Sub

Private Sub dtpFec_Emision_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

    'txtiva.Text = DevuelveCampo("SELECT PORC_IGV  FROM TG_IGV where ano= '" & Year(dtpFec_Emision) & "' and mes= '" & Format(Month(dtpFec_Emision), "00") & "'", cConnect)
End Sub

Private Sub dtpFec_Emision_Change()
    txtiva.Text = DevuelveCampo("SELECT PORC_IGV  FROM TG_IGV where ano= '" & Year(dtpFec_Emision) & "' and mes= '" & Format(Month(dtpFec_Emision), "00") & "'", cConnect)
    TxtTipo_Cambio.Text = DevuelveCampo("select isnull(Tipo_Venta,0) from cn_tipocambio where fecha = '" & dtpFec_Emision & "'", cConnect)
End Sub

Private Sub Form_Load()
  If Not IsNumeric(txtiva.Text) Then
     txtiva.Text = 0
    End If
    iva = 1 + (txtiva.Text / 100#)
    Call DisableCloseButton(Me)
    flg_Tiene_guias_asignadas = "N"
    FraProductos.Visible = False
    fraSelGuias.Visible = False
    dtpFec_Emision.Value = Date
    dtpFec_Registro.Value = Date
    ind_aplica_precio = 1
    Call buscaDetalle_factura
    Call obtieneDatosIniciales
    Call FillAlmacen
    'txtCod_TipDoc.SetFocus
    txtiva.Text = DevuelveCampo("SELECT PORC_IGV  FROM TG_IGV where ano= '" & Year(dtpFec_Emision) & "' and mes= '" & Format(Month(dtpFec_Emision), "00") & "'", cConnect)
    TxtTipo_Cambio.Text = DevuelveCampo("select isnull(Tipo_Venta,0) from cn_tipocambio where fecha = '" & dtpFec_Emision & "'", cConnect)
  
    If CDbl(txtiva.Text) <= 0 Then
        Call MsgBox("Ingrese el porcentaje del impuesto sobre el valor agregado (iva) ", vbCritical, "Importante")
        'Unload Me
    End If
    
     iva = 1 + (txtiva.Text / 100#)
            
    If Not IsNumeric(TxtTipo_Cambio.Text) Then
      TxtTipo_Cambio.Text = 0
    End If
    
    If CDbl(TxtTipo_Cambio.Text) <= 0 Then
        Call MsgBox("Ingrese el Tipo Cambio Para la fecha", vbCritical, "Importante")
        'Unload Me
    End If
    
End Sub
Private Sub FillAlmacen()
On Error GoTo fin
Dim sTit As String
    
    sTit = "Cargar Almacenes"
    StrSQL = " TI_MUESTRA_ALMACENES_TELA_TENIDA_ROLLO  '" & vusu & "'"
    
    Set rstAux = CargarRecordSetDesconectado(StrSQL, cConnect)
    cboAlmacen.Clear
    With rstAux
        If .RecordCount > 0 Then .MoveFirst
        Do Until .EOF
            cboAlmacen.AddItem !Cod_almacen & " " & !nom_almacen
            .MoveNext
        Loop
        .Close
    End With
    If cboAlmacen.ListCount > 0 Then cboAlmacen.ListIndex = 0
    Set rstAux = Nothing
    
Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, sTit
End Sub

Private Sub obtieneDatosIniciales()
Dim StrSQL As String
Dim pc As String
Dim auxset As ADODB.Recordset
pc = ComputerName
StrSQL = "CN_MUESTRA_CAJAS_VENDEDOR_ACCESO '" & pc & "'"
 Set auxset = Nothing
 Set auxset = CargarRecordSetDesconectado(StrSQL, cConnect)
 If auxset.RecordCount > 0 Then
    TxtCod_Fabrica.Text = auxset("cod_Fabrica")
    txtDes_Fabrica.Text = auxset("nom_fabrica")
    txtCod_Tienda.Text = auxset("cod_tienda")
    txtDes_Tienda.Text = auxset("des_tienda")
    txtCod_Caja.Text = auxset("cod_caja")
    txtDes_Caja.Text = auxset("des_caja")
    txtCod_Vendedor.Text = auxset("cod_vendedor")
    txtDes_Vendedor.Text = auxset("des_vendedor")
    txtCod_Almacen.Text = auxset("cod_almacen")
    txtDes_Almacen.Text = auxset("nom_almacen")
Else
    Call MsgBox("La PC no Tiene una Caja Asignada", vbExclamation, "Importante")

End If
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo dprDepurar
Select Case ActionName
Case Is = "GRABAR"
  If grxDatos.RowCount <= 0 Then Exit Sub
          
         If validaDatosIniciales = True Then
         
            If MsgBox("¡¡¡Esta apunto de confirmar en caja el documento de venta!!!:" & Chr(13) & Chr(10) & ":::::> " & Trim(txtDes_TipDoc.Text) & " " & txtSer_Docum & "-" & txtNum_Docum & Chr(13) & Chr(10) & "¿Son los datos correctos?", vbInformation + vbYesNo, "CONFIRMAR") = vbYes Then
          
            If flg_Tiene_guias_asignadas = "N" Then
              If GuardaDetalleVentas = True Then
                  Call obtieneDatosIniciales
                  Call estadoInicialVentana
                  Call buscaDetalle_factura
              End If
            End If
            
            If flg_Tiene_guias_asignadas = "S" Then
              If GuardaDetalleVentasGuias = True Then
                  Call obtieneDatosIniciales
                  Call estadoInicialVentana
                  Call buscaDetalle_factura
              End If
            End If
          
         End If
         
        End If
       
Case Is = "CANCELAR"
  If grxDatos.RowCount > 0 Then
  
  If MsgBox("¡...Al cancelar esta operacion se eliminaran los datos registrados...! " & Chr(13) & Chr(10) & " ¿Esta Seguro de proseguir? ", vbInformation + vbYesNo, "CONFIRMAR") = vbYes Then
    If flg_Tiene_guias_asignadas = "S" Then
      Call EliminaGuiasAsigandas
      End If
    Unload Me
  End If
  
  Else
    Unload Me
    
  End If
  
End Select

Exit Sub

Resume
dprDepurar:
errores err.Number
End Sub
Private Sub EliminaGuiasAsigandas()
On Error GoTo fin
Dim rsguiaAsig As New ADODB.Recordset

If grxListaGuiasSeleccionadas.RowCount <= 0 Then Exit Sub
  grxListaGuiasSeleccionadas.Update
  
Set rsguiaAsig = grxListaGuiasSeleccionadas.ADORecordset
  
rsguiaAsig.MoveFirst
Do While Not rsguiaAsig.EOF

StrSQL = "CN_ASIGNA_GUIA_FACTURA '" & rsguiaAsig("cod_almacen") & "','" & rsguiaAsig("num_movstk") & "','',''"
Call ExecuteCommandSQL(cConnect, StrSQL)

rsguiaAsig.MoveNext
Loop

sTit = "Importante"

Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, sTit
    
End Sub
Private Sub estadoInicialVentana()
'''generar el sgte numero de documento
'''limpiar y txt, grilla
txtDes_TipAne.Text = ""
txtNum_Ruc.Text = ""
txtDes_TipAne.Tag = ""
txtNum_Ruc.Tag = ""
txtNum_Docum.Text = DevuelveCampo("SELECT COR_NUMACTU FROM CN_VENTAS_CAJAS_DOCUMENTOS WHERE COD_FABRICA='" & TxtCod_Fabrica.Text & "' AND  COD_TIENDA='" & Trim(txtCod_Tienda.Text) & "' AND COD_CAJA='" & txtCod_Caja.Text & "' AND COD_TIPDOC='" & Trim(txtCod_TipDoc.Text) & "' AND COR_DOCSERIE ='" & txtSer_Docum.Text & "' ", cConnect)

Set grxDatos.ADORecordset = Nothing
Set GrxProductos.ADORecordset = Nothing

End Sub
'''''***********************************GUARDA EL DETALLE DE LA FACTURA DESDE LA BUSQUEDA O CON LECTORA DE BARRAS, GENERA MOVIMIENTO DE ALMACEN ****************************
Private Function GuardaDetalleVentas() As Boolean
On Error GoTo ErrDetMov
Dim sErr As String, cntAux As New ADODB.Connection, sTit As String, _
    snum_movstk As String, strNum_Corre  As String
Dim Kilos_Tenidos As Double
Dim RollosTeñidos As Double
Dim rstAux As New ADODB.Recordset

  GuardaDetalleVentas = False

    '''txtCod_OrdTra_Tinto = Trim(txtCod_OrdTra_Tinto)
    Kilos_Tenidos = 0
    RollosTeñidos = 0
    
    If grxDatos.RowCount = 0 Then
        MsgBox "Se debe especificar al menos un detalle", vbExclamation + vbOKCancel, sTit
        Exit Function
    End If
    
   sTit = "Guardar Detalle de Ventas"
    
    cntAux.Open cConnect
    cntAux.BeginTrans

    '''CABECERA VENTAS
    StrSQL = "VENTAS_UP_MAN_ROLLOS 'I','','" & TxtCod_Fabrica.Text & "','" & Trim(txtCod_Tienda.Text) & "','" & Trim(txtCod_Caja.Text) & "','" & Trim(txtCod_Vendedor.Text) & "','" & Trim(txtCod_Almacen.Text) & "','" & txtCod_TipDoc & "','" & txtSer_Docum & "','" _
            & txtNum_Docum & "','C','" & Trim(txtNum_Ruc.Tag) & "','" & txtCod_ConPag & "','" & txtCod_TipVenta.Text & "','" & Format(dtpFec_Emision.Value, "dd/mm/yyyy") & "','" _
            & Format(dtpFec_Registro.Value, "dd/mm/yyyy") & "','" & txtCod_Moneda & "','" _
            & vusu & "',''," _
            & TxtTipo_Cambio.Text & ",'','','N','N','S'"
            
    Set rstAux = cntAux.Execute(StrSQL, adExecuteNoRecords)
    strNum_Corre = rstAux!Num_Corre
    rstAux.Close

    '''CABECERA MOVIMIENTO
    StrSQL = "EXEC TI_UP_MAN_TX_MOVISTK_TELA_TENIDA_CABECERA_ROLLOS 'I', '" & _
             Trim(txtCod_Almacen.Text) & "', '', '" & Format(dtpFec_Registro.Value, _
             "dd/mm/yyyy") & "', '', '' ,'SVD','', '', '" & txtDes_TipAne.Tag & _
             "', '', '', 'movimiento de venta directo', '" & vusu & "', '" & _
             0 & "', '" & 0 & "','',''"

    Set rstAux = cntAux.Execute(StrSQL, adExecuteNoRecords)
    snum_movstk = rstAux!num_movstk
    rstAux.Close
    
    Set rstAux = grxDatos.ADORecordset
    With rstAux
        If .RecordCount > 0 Then .MoveFirst
        Do Until .EOF
    '''DETALLE MOVIMIENTO DE SALIDA DE ALMACEN
             StrSQL = "EXEC TI_UP_MAN_TX_MOVISTK_TELA_TENIDA_PESADAS_ROLLO_VENTAS_DIRECTA 'I', '" & _
             Trim(txtCod_Almacen.Text) & "', '" & snum_movstk & "', '', '" & _
             !codigoRollo & "'," & !Stock & "," & !cant & ",0, " & _
             Trim(!rollos) & ",'" & vusu & "',0"
             cntAux.Execute StrSQL, adExecuteNoRecords
    
    '''DETALLE VENTAS falta strCod_Anxo
            StrSQL = "VENTAS_UP_MAN_DETALLE_ROLLO 'I','" & strNum_Corre & "','','D','" & Trim(!codigoRollo) & "','" & _
            Trim(!Cod_Tela) & "','','" & !und & "'," & !rollos & "," & !Stock & "," & !cant & "," _
            & !precio & "," & !Total & ",0,'','',0,'" & Trim(txtCod_Almacen.Text) & "','" & !OT & "','" & vusu & "'"
            cntAux.Execute StrSQL, adExecuteNoRecords
            .MoveNext
        Loop
    End With
    
    '''ASOCIA FACTURA CON MOVIMIENTO DE ALMACEN
    StrSQL = "CN_VENTAS_CAJAS_RELACIONA_FACTURA_GUIA 'U','" & strNum_Corre & "','" & Trim(txtSer_Docum.Text) & "','" & Trim(txtNum_Docum.Text) & "','" & Trim(txtCod_Almacen.Text) & "','" & snum_movstk & "'"
    cntAux.Execute StrSQL, adExecuteNoRecords

    cntAux.CommitTrans
    cntAux.Close
    Set cntAux = Nothing
    
    '''IMPRIME DOCUMENTO
    Call Preliminar_Docum_Ventas(strNum_Corre)
    GuardaDetalleVentas = True
    'Unload Me
Exit Function
ErrDetMov:
    GuardaDetalleVentas = False
    sErr = err.Description
    cntAux.RollbackTrans
    cntAux.Close
    Set cntAux = Nothing
    MsgBox sErr, vbCritical + vbOKOnly, sTit
End Function

'''''***********************************GUARDA EL DETALLE DE LA FACTURA DESDE EL DETALLE DE LA GUIA****************************
Private Function GuardaDetalleVentasGuias() As Boolean
On Error GoTo ErrDetMov
Dim sErr As String, cntAux As New ADODB.Connection, sTit As String, _
    snum_movstk As String, strNum_Corre  As String
Dim Kilos_Tenidos As Double
Dim RollosTeñidos As Double
Dim rstAux As New ADODB.Recordset

  GuardaDetalleVentasGuias = False

    '''txtCod_OrdTra_Tinto = Trim(txtCod_OrdTra_Tinto)
    Kilos_Tenidos = 0
    RollosTeñidos = 0
    
    If grxDatos.RowCount = 0 Then
        MsgBox "Se debe especificar al menos un detalle", vbExclamation + vbOKCancel, sTit
        Exit Function
    End If
    
   sTit = "Guardar Detalle de Ventas"
    
    cntAux.Open cConnect
    cntAux.BeginTrans

    '''CABECERA VENTAS
    StrSQL = "VENTAS_UP_MAN_ROLLOS 'I','','" & TxtCod_Fabrica.Text & "','" & Trim(txtCod_Tienda.Text) & "','" & Trim(txtCod_Caja.Text) & "','" & Trim(txtCod_Vendedor.Text) & "','" & Trim(txtCod_Almacen.Text) & "','" & txtCod_TipDoc & "','" & txtSer_Docum & "','" _
            & txtNum_Docum & "','C','" & Trim(txtNum_Ruc.Tag) & "','" & txtCod_ConPag & "','" & txtCod_TipVenta.Text & "','" & Format(dtpFec_Emision.Value, "dd/mm/yyyy") & "','" _
            & Format(dtpFec_Registro.Value, "dd/mm/yyyy") & "','" & txtCod_Moneda & "','" _
            & vusu & "',''," _
            & TxtTipo_Cambio.Text & ",'','','N','N','N'"
            
    Set rstAux = cntAux.Execute(StrSQL, adExecuteNoRecords)
    strNum_Corre = rstAux!Num_Corre
    rstAux.Close

    '''CABECERA MOVIMIENTO
'    STRSQL = "EXEC TI_UP_MAN_TX_MOVISTK_TELA_TENIDA_CABECERA_ROLLOS 'I', '" & _
'             Trim(txtCod_Almacen.Text) & "', '', '" & Format(dtpFec_Registro.Value, _
'             "dd/mm/yyyy") & "', '', '' ,'SVD','', '', '" & txtDes_TipAne.Tag & _
'             "', '', '', 'movimiento de venta directo', '" & vusu & "', '" & _
'             0 & "', '" & 0 & "','',''"

'    Set rstAux = cntAux.Execute(STRSQL, adExecuteNoRecords)
'    sNum_MovStk = rstAux!num_movstk
'    rstAux.Close
    
    Set rstAux = grxDatos.ADORecordset
    With rstAux
        If .RecordCount > 0 Then .MoveFirst
        Do Until .EOF
'    '''DETALLE MOVIMIENTO DE SALIDA DE ALMACEN
'             STRSQL = "EXEC TI_UP_MAN_TX_MOVISTK_TELA_TENIDA_PESADAS_ROLLO_VENTAS_DIRECTA 'I', '" & _
'             Trim(txtCod_Almacen.Text) & "', '" & sNum_MovStk & "', '', '" & _
'             !codigorollo & "'," & !Stock & "," & !cant & ",0, " & _
'             Trim(!rollos) & ",'" & vusu & "',0"
'             cntAux.Execute STRSQL, adExecuteNoRecords
    
    '''DETALLE VENTAS falta strCod_Anxo
            StrSQL = "VENTAS_UP_MAN_DETALLE_ROLLO 'I','" & strNum_Corre & "','','D','" & Trim(!codigoRollo) & "','" & _
            Trim(!Cod_Tela) & "','','" & !und & "'," & !rollos & "," & !Stock & "," & !cant & "," _
            & !precio & "," & !Total & ",0,'','',0,'" & Trim(txtCod_Almacen.Text) & "','" & !OT & "','" & vusu & "'"
            cntAux.Execute StrSQL, adExecuteNoRecords
            .MoveNext
        Loop
    End With
    
    '''ASOCIA FACTURA CON MOVIMIENTO DE ALMACEN
    StrSQL = "CN_VENTAS_CAJAS_RELACIONA_FACTURA_GUIA 'U','" & strNum_Corre & "','" & Trim(txtSer_Docum.Text) & "','" & Trim(txtNum_Docum.Text) & "','" & Trim(txtCod_Almacen.Text) & "',''"
    cntAux.Execute StrSQL, adExecuteNoRecords

    cntAux.CommitTrans
    cntAux.Close
    Set cntAux = Nothing
    
    '''IMPRIME DOCUMENTO
    Call Preliminar_Docum_Ventas(strNum_Corre)
    GuardaDetalleVentasGuias = True
    'Unload Me
Exit Function
ErrDetMov:
    GuardaDetalleVentasGuias = False
    sErr = err.Description
    cntAux.RollbackTrans
    cntAux.Close
    Set cntAux = Nothing
    MsgBox sErr, vbCritical + vbOKOnly, sTit
End Function

Private Sub Preliminar_Docum_Ventas(Num_Corre As String)
On Error GoTo SALTO_ERROR

Dim sSQL As String, rs As New ADODB.Recordset
Dim imp_total As Double

Dim aMess(4), i As Integer
'  ssql = "Ventas_Actualiza_Datos_Impresion '$' , '$' , '$' , '$', '$' "
'  ssql = VBsprintf(ssql, _
'  GridEX1.Value(GridEX1.Columns("Num_Corre").Index), _
'  Format(GridEX1.Value(GridEX1.Columns("EMISION").Index), "dd/mm/yyyy"), _
'  IIf(GridEX1.Value(GridEX1.Columns("Retencion").Index), "S", "N"), _
'  GridEX1.Value(GridEX1.Columns("Glosa").Index), "N")
'  ExecuteCommandSQL cConnect, ssql
 
imp_total = DevuelveCampo("SELECT IMP_TOTAL FROM CN_VENTAS where num_corre='" & Num_Corre & "'", cConnect)
 
If Imprimir_FACTURA(Num_Corre, imp_total, Trim(txtCod_TipDoc.Text), Trim(txtSer_Docum.Text)) = False Then
   MsgBox "Problemas de Impresion con el Documento Nro " & txtNum_Docum.Text, vbInformation, "ERROR"
   'Buscar
   Exit Sub
End If
    
Exit Sub
SALTO_ERROR:
MsgBox err.Description, vbCritical, Me.Caption
    
End Sub
   
Public Function Imprimir_FACTURA(lvNumCorre As String, dbImp_Total As Double, strCod_Cod As String, Serie As String) As Boolean

Dim Rs2 As New ADODB.Recordset, rs3 As New ADODB.Recordset, StrSQL As String, scnt As Integer
scnt = 0
With rsFactura
 
    Select Case strCod_Cod
    
    Case Is = "FA" 'llll
        StrSQL = "VENTAS_EMITE_FACTURA_VENTAS_DETA_ROLLO '" & lvNumCorre & "','" & UCase(EnLetras(Trim(CStr(dbImp_Total)))) & "'"
        Set rsFactura = CargarRecordSetDesconectado(StrSQL, cConnect)
        
        If rsFactura.RecordCount > 0 Then
            Call Factura_sa("FA", Serie)
            scnt = 2
        Else
           Call MsgBox("La Factura no Tiene Detalle", vbInformation, "Mensaje")
           Imprimir_FACTURA = False
           Exit Function
        End If
'    Case Is = "NC"
'      strSQL = "Ventas_Emite_Documento_Abono '" & lvNumCorre & "','" & UCase(EnLetras(Trim(CStr(dbImp_Total)))) & "'"
'      Set rsFactura = CargarRecordSetDesconectado(strSQL, cConnect)
'      Call Factura_sa("NC", Serie)
'    Case Is = "ND"
'      strSQL = "Ventas_Emite_Documento_Abono '" & lvNumCorre & "','" & UCase(EnLetras(Trim(CStr(dbImp_Total)))) & "'"
'      Set rsFactura = CargarRecordSetDesconectado(strSQL, cConnect)
'      Call Factura_sa("ND", Serie)
      
    Case Is = "BV"
        
        'STRSQL = " Ventas_Emite_Documento_Abono '" & lvNumCorre & "','" & UCase(EnLetras(Trim(CStr(dbImp_Total)))) & "'"
        'Set rsFactura = CargarRecordSetDesconectado(STRSQL, cConnect)
        StrSQL = "VENTAS_EMITE_FACTURA_VENTAS_DETA_ROLLO '" & lvNumCorre & "','" & UCase(EnLetras(Trim(CStr(dbImp_Total)))) & "'"
        Set rsFactura = CargarRecordSetDesconectado(StrSQL, cConnect)
        
        Call Factura_sa("BV", Serie)
    Case Else
      MsgBox "No se ha Definido un Formato de Impresion para este tipo de documento", vbInformation, "ERROR"
       Imprimir_FACTURA = False
      Exit Function
    End Select
    
    'If rsFactura.RecordCount = 0 Then
    '  Imprimir_FACTURA = False
    '  Exit Function
    'End If
    
End With

Imprimir_FACTURA = True

End Function

Sub Factura_sa(tipo As String, Serie As String)
On Error GoTo ErrorImpresion
Dim oo As Object, lvSql As String, lvRuta As String

    Set oo = CreateObject("excel.application")
    
    If tipo = "FA" Then
        'If chkImpresionDirecta.Value = Checked Then
            oo.Workbooks.Open vRuta & "\Factura_Tela_Acabada_Rollo_Directa.XLT"
        'Else
        '    oo.Workbooks.Open vRuta & "\Factura_Tela_Acabada_Rollo.XLT"
        'End If
    End If

    If tipo = "ND" Then
        oo.Workbooks.Open vRuta & "\Abono_Textil.XLT"
    End If
    If tipo = "NC" Then
        oo.Workbooks.Open vRuta & "\Credito_Textil.XLT"
    End If

    If tipo = "BV" Then
        oo.Workbooks.Open vRuta & "\Impresion_Boleta.XLT"
    End If
    
    oo.DisplayAlerts = False
    
    If chkImpresionDirecta.Value = 1 Then
        oo.Visible = False
    Else
        oo.Visible = True
    End If
            
    oo.Run "Reporte", rsFactura, IIf(chkImpresionDirecta.Value = 1, 1, 0), cConnect
 
    If chkImpresionDirecta.Value = 1 Then
        oo.Workbooks.Close
    End If
    
    Set oo = Nothing
        
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    MsgBox "Hubo error en la impresion de La Factura " & err.Description, vbCritical, "Impresion"
End Sub
Sub Cambio_FR()
     'Imp_Gastos_Finacieros.Text = 0
     'Imp_Otros.Text = 0
     'Imp_Flete.Text = 0
     'txtPeso_Bruto.Text = 0
     'txtShip_Date.Text = ""
     'txtPeso_Neto.Text = 0
     'chkFlete.Value = 0
     'chkSeguro.Value = 0
     'frOtros.Visible = False
     'frExportacion.Visible = False
     'frReferencia.Visible = False
     'If txtCod_TipDoc = "NC" Or txtCod_TipDoc = "ND" Then
       'frReferencia.Visible = True
     'End If
     
     'If chkExportacion Then
     '  frExportacion.Visible = True
     'Else
     '  frOtros.Visible = True
     'End If
  
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

Select Case ActionName
    Case "GUIAS"
        If FraProductos.Visible = False And ((grxDatos.RowCount = 0 And flg_Tiene_guias_asignadas = "N") Or (flg_Tiene_guias_asignadas = "S")) Then
            Call FillAlmacen
            Call buscalistaGuiasPendientes
            Call buscalistaGuiasSeleccionadas
            fraSelGuias.Visible = True
        End If
    Case "AYUDA"
        
        If fraSelGuias.Visible = False And flg_Tiene_guias_asignadas = "N" Then
            FraProductos.Visible = True
            limpiarCajasBusqueda
        End If
End Select


End Sub

''''******************************HABILITA LA EDICION SOLO DE ALGUNAS COLUMNAS LAS TIENEN CANCEL=FALSE***********************
Private Sub GrxProductos_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
  Select Case ColIndex
    Case Is = GrxProductos.Columns("CANT").Index
      Cancel = False
    'Case Is = GrxProductos.Columns("ROLLOS").Index
    '  Cancel = False
    'Case Is = GrxProductos.Columns("PRECIO").Index
    '  Cancel = False
    'Case Is = GrxProductos.Columns("SEL").Index
    '  Cancel = False
    Case Else
      Cancel = True
  End Select
End Sub
Private Sub grxDatos_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
  Select Case ColIndex
    Case Is = grxDatos.Columns("CANT").Index
      Cancel = False
    'Case Is = grxDatos.Columns("ROLLOS").Index
    '  Cancel = False
    Case Is = grxDatos.Columns("PRECIO").Index
      Cancel = False
    Case Is = grxDatos.Columns("ELI").Index
      Cancel = False
    Case Else
      Cancel = True
  End Select
End Sub
'''******************************* ADICIONA ARTICULOS CON DOBLE CLICK *******************************************
Private Sub GrxProductos_DblClick()
    'adicionarProducto
    'FraProductos.Visible = False
End Sub
Private Function validaDatosIniciales() As Boolean
    validaDatosIniciales = True
         
    'If fraUbicacion.Enabled = False Then
    
        If Trim(TxtCod_Fabrica.Text) = "" Then
           Call MsgBox("Ingrese Una Empresa valida", vbInformation + vbOKOnly, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
        
        If Trim(txtCod_Tienda.Text) = "" Then
           Call MsgBox("Ingrese una Tienda valida", vbInformation + vbOKOnly, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
        
        If Trim(txtCod_Caja.Text) = "" Then
           Call MsgBox("Ingrese una caja valida", vbInformation + vbOKOnly, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
        
        If Trim(txtCod_Vendedor.Text) = "" Then
           Call MsgBox("el codigo del vendedor no es valido", vbInformation + vbOKOnly, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
        
        If Trim(txtCod_Almacen.Text) = "" Then
           Call MsgBox("El Codigo del Almacen no es valido", vbInformation + vbOKOnly, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
       
        If DevuelveCampo("SELECT COUNT(*) FROM CN_VENTAS_CAJAS_ALMACEN WHERE COD_FABRICA='" & Trim(TxtCod_Fabrica.Text) & "'", cConnect) <= 0 Then
           Call MsgBox("El Codigo Empresa no Valida", vbInformation + vbOKOnly, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
         
        If DevuelveCampo("SELECT COUNT(*) FROM CN_VENTAS_CAJAS_ALMACEN WHERE COD_FABRICA='" & Trim(TxtCod_Fabrica.Text) & "' and  cod_tienda='" & Trim(txtCod_Tienda.Text) & "'", cConnect) <= 0 Then
           Call MsgBox("El Codigo de Tienda no valida", vbInformation + vbOKOnly, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
         
        If DevuelveCampo("SELECT COUNT(*) FROM CN_VENTAS_CAJAS_ALMACEN WHERE COD_FABRICA='" & Trim(TxtCod_Fabrica.Text) & "' and  cod_tienda='" & Trim(txtCod_Tienda.Text) & "'and cod_caja = '" & Trim(txtCod_Caja.Text) & "'", cConnect) <= 0 Then
           Call MsgBox("El Codigo de caja no es valido ", vbInformation + vbOKOnly, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
         
        If DevuelveCampo("SELECT COUNT(*) FROM CN_VENTAS_CAJAS_ALMACEN WHERE COD_FABRICA='" & Trim(TxtCod_Fabrica.Text) & "' and  cod_tienda='" & Trim(txtCod_Tienda.Text) & "'and cod_caja = '" & Trim(txtCod_Caja.Text) & "' and cod_almacen= '" & Trim(txtCod_Almacen.Text) & "'", cConnect) <= 0 Then
           Call MsgBox("El Codigo de Caja no es valido ", vbInformation + vbOKOnly, "Mensaje")
           validaDatosIniciales = False
           Exit Function
           
        End If
        
        If Trim(txtCod_Moneda.Text) = "" Then
           Call MsgBox("Codigo de Moneda no valida", vbInformation + vbOKOnly, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
                   
        If DevuelveCampo(" select count(*) from cn_ventas where cod_tipdoc='" & Trim(txtCod_TipDoc.Text) & "' and  ser_docum='" & Trim(txtSer_Docum.Text) & "' and num_docum_ventas='" & Trim(txtNum_Docum.Text) & "' ", cConnect) > 0 Then
             MsgBox " El numero de documento ya Existe en el sistema, sirvase revisar...", vbInformation + vbOKOnly, "Importante"
             validaDatosIniciales = False
             Exit Function
        End If
        
        If DevuelveCampo(" select count(*) from cn_ventas_anulados where cod_tipdoc='" & Trim(txtCod_TipDoc.Text) & "' and  SER_DOCUM='" & Trim(txtSer_Docum.Text) & "' and num_docum_ventas='" & Trim(txtNum_Docum.Text) & "' ", cConnect) > 0 Then
             MsgBox " El numero de documento ya Existe en el sistema, sirvase revisar...", vbInformation + vbOKOnly, "Importante"
             validaDatosIniciales = False
             Exit Function
        End If
        
        If Trim(txtDes_TipAne.Text) = "" Then
           Call MsgBox("sivarse ingresar un cliente valido ", vbInformation + vbOKOnly, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
         
    'End If
End Function
'''******************************* ADICIONA LISTA ARTICULOS CUYA CANTIDAD SEA MAYOR A 0*******************************************
Private Sub adicionarProductoMasivo()
Dim RSAUX As ADODB.Recordset
Dim rslista As ADODB.Recordset
Dim i As Integer
On Error GoTo fin

If validaDatosIniciales = False Then
    Exit Sub
End If

GrxProductos.Refresh
GrxProductos.Update

Set RSAUX = grxDatos.ADORecordset

Set rslista = GrxProductos.ADORecordset
If rslista.RecordCount <= 0 Then Exit Sub

rslista.Update
'RSAUX.Update
rslista.MoveFirst
i = 1
Do While i <= rslista.RecordCount
If rslista!cant > 0 Then

    RSAUX.AddNew
    RSAUX!OT = rslista!OT
    RSAUX!codigoRollo = rslista!codigoRollo
    RSAUX!Cod_Tela = rslista!Cod_Tela
    RSAUX!Tela = rslista!Tela
    RSAUX!cod_Color = rslista!cod_Color
    RSAUX!color = rslista!color
    RSAUX!calidad = rslista!calidad
    RSAUX!rollos = rslista!rollos
    RSAUX!und = rslista!und
    RSAUX!Stock = rslista!Stock
    RSAUX!cant = rslista!cant
    RSAUX!precio = rslista!precio
    RSAUX!DEL = "X"
    RSAUX!Total = RSAUX!precio * RSAUX!cant
    RSAUX.Update
    
End If
rslista.MoveNext
i = i + 1
Loop

Set grxDatos.ADORecordset = RSAUX
Set rslista = Nothing

If grxDatos.RowCount >= 1 Then
    fraUbicacion.Enabled = False
Else
    fraUbicacion.Enabled = True
End If

Call Total_documento
Call ConfiguraGrilla_Detalle

Exit Sub
Resume
fin:
On Error Resume Next
Set RSAUX = Nothing
MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
"Edicionar Producto "
    
End Sub
'''******************************* ADICIONA LISTA ARTICULOS DESDE EL DETALLE DE LA GUIA 0*******************************************
Private Sub adicionarProductoDesdeDetalleGuia()
Dim RSAUX As ADODB.Recordset
Dim rslista As ADODB.Recordset
Dim i As Integer
On Error GoTo fin
''' volvemos a llenar el detalle
Call buscaDetalle_factura
Set RSAUX = grxDatos.ADORecordset

''' detalle de las guias
StrSQL = "CN_MUESTRA_DETALLE_GUIA_VENTA '" & Trim(txtSer_Docum.Text) & "','" & Trim(txtNum_Docum) & "'"
Set rslista = Nothing
Set rslista = CargarRecordSetDesconectado(StrSQL, cConnect)

If rslista.RecordCount <= 0 Then Exit Sub

If validaDatosIniciales = False Then
    Exit Sub
End If

grxDatos.Refresh
grxDatos.Update



rslista.Update
'RSAUX.Update
rslista.MoveFirst
i = 1
Do While i <= rslista.RecordCount
If rslista!cant > 0 Then

    RSAUX.AddNew
    RSAUX!OT = rslista!OT
    RSAUX!codigoRollo = rslista!codigoRollo
    RSAUX!Cod_Tela = rslista!Cod_Tela
    RSAUX!Tela = rslista!Tela
    RSAUX!cod_Color = rslista!cod_Color
    RSAUX!color = rslista!color
    RSAUX!calidad = rslista!calidad
    RSAUX!rollos = rslista!rollos
    RSAUX!und = rslista!und
    RSAUX!Stock = rslista!Stock
    RSAUX!cant = rslista!cant
    RSAUX!precio = rslista!precio
    RSAUX!DEL = "X"
    RSAUX!Total = RSAUX!precio * RSAUX!cant
    RSAUX.Update
    
End If
rslista.MoveNext
i = i + 1
Loop

Set grxDatos.ADORecordset = RSAUX
Set rslista = Nothing

If grxDatos.RowCount >= 1 Then
    fraUbicacion.Enabled = False
Else
    fraUbicacion.Enabled = True
End If

'Call Total_documento
Call ConfiguraGrilla_Detalle

Exit Sub
Resume
fin:
On Error Resume Next
Set RSAUX = Nothing
MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
"Edicionar Producto "
    
End Sub
'''******************************* ADICIONA LISTA ARTICULOS *******************************************
Private Sub adicionarProducto()
Dim RSAUX As ADODB.Recordset
On Error GoTo fin

Set RSAUX = grxDatos.ADORecordset
RSAUX.AddNew

RSAUX!OT = IIf(IsNull(GrxProductos.Value(GrxProductos.Columns("OT").Index)), "", GrxProductos.Value(GrxProductos.Columns("OT").Index))
RSAUX!codigoRollo = IIf(IsNull(GrxProductos.Value(GrxProductos.Columns("codigorollo").Index)), "", GrxProductos.Value(GrxProductos.Columns("codigorollo").Index))
RSAUX!Cod_Tela = IIf(IsNull(GrxProductos.Value(GrxProductos.Columns("cod_tela").Index)), "", GrxProductos.Value(GrxProductos.Columns("cod_tela").Index))
RSAUX!Tela = IIf(IsNull(GrxProductos.Value(GrxProductos.Columns("TELA").Index)), "", GrxProductos.Value(GrxProductos.Columns("TELA").Index))
RSAUX!cod_Color = IIf(IsNull(GrxProductos.Value(GrxProductos.Columns("cod_Color").Index)), "", GrxProductos.Value(GrxProductos.Columns("cod_color").Index))
RSAUX!color = IIf(IsNull(GrxProductos.Value(GrxProductos.Columns("COLOR").Index)), "", GrxProductos.Value(GrxProductos.Columns("COLOR").Index))
RSAUX!calidad = IIf(IsNull(GrxProductos.Value(GrxProductos.Columns("calidad").Index)), "", GrxProductos.Value(GrxProductos.Columns("calidad").Index))
RSAUX!rollos = IIf(IsNull(GrxProductos.Value(GrxProductos.Columns("rollos").Index)), "", GrxProductos.Value(GrxProductos.Columns("rollos").Index))
RSAUX!und = IIf(IsNull(GrxProductos.Value(GrxProductos.Columns("und").Index)), "", GrxProductos.Value(GrxProductos.Columns("und").Index))
RSAUX!cant = IIf(IsNull(GrxProductos.Value(GrxProductos.Columns("cant").Index)), "", GrxProductos.Value(GrxProductos.Columns("cant").Index))
RSAUX!Stock = IIf(IsNull(GrxProductos.Value(GrxProductos.Columns("stock").Index)), "", GrxProductos.Value(GrxProductos.Columns("stock").Index))
RSAUX!precio = IIf(IsNull(GrxProductos.Value(GrxProductos.Columns("PRECIO").Index)), "", GrxProductos.Value(GrxProductos.Columns("PRECIO").Index))
RSAUX!DEL = "X"
RSAUX!Total = RSAUX!precio * RSAUX!cant
RSAUX.Update
Set grxDatos.ADORecordset = RSAUX

'Call Total_documento
Call ConfiguraGrilla_Detalle

Exit Sub
Resume
fin:
On Error Resume Next
Set RSAUX = Nothing
MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
"Edicionar Producto "
    
End Sub

Private Sub Option2_Click()

End Sub

Private Sub Option1_Click(Index As Integer)
ind_aplica_precio = Index
End Sub

Private Sub txtCodigo_Producto_Change()
  If Len(Trim(txtCodigo_Producto.Text)) = 9 And flg_Tiene_guias_asignadas = "N" Then
    Call AdicionaProductoDirecto(1)
    txtCodigo_Producto.Text = ""
    SendKeys "{TAB}"
  End If
End Sub
Private Sub AdicionaProductoDirecto(Opcion As String)

    Dim StrSQL As String
    Dim sCodCentroCosto As String
    Dim rsetAux As ADODB.Recordset
    Dim rsetbusqueda As ADODB.Recordset
    Dim nrofilas As Integer
    
    On Error GoTo fin
    
    If validaDatosIniciales = False Then
        Exit Sub
    End If
    
    StrSQL = "TX_MUESTRA_ROLLOS_VENTA '" & Opcion & "','" & Trim(txtCod_Almacen.Text) & "','" & Trim(txtCodigo_Producto.Text) & "','" & Trim(txtBus_Cod_ordtra.Text) & "','" & Trim(txtDescripcion_Producto.Text) & "','" & Trim(txtBus_Des_Color.Text) & "'"
    Set rsetbusqueda = Nothing
    Set rsetbusqueda = CargarRecordSetDesconectado(StrSQL, cConnect)
    If rsetbusqueda.RecordCount <= 0 Then Exit Sub
    
    Set rsetAux = grxDatos.ADORecordset
    rsetAux.AddNew
    
    rsetAux!OT = rsetbusqueda!OT
    rsetAux!codigoRollo = rsetbusqueda!codigoRollo
    rsetAux!Cod_Tela = rsetbusqueda!Cod_Tela
    rsetAux!Tela = rsetbusqueda!Tela
    rsetAux!cod_Color = rsetbusqueda!cod_Color
    rsetAux!color = rsetbusqueda!color
    rsetAux!calidad = rsetbusqueda!calidad
    rsetAux!rollos = rsetbusqueda!rollos
    rsetAux!und = rsetbusqueda!und
    rsetAux!Stock = rsetbusqueda!Stock
    rsetAux!cant = rsetbusqueda!cant
    rsetAux!precio = rsetbusqueda!precio
    rsetAux!DEL = "X"
    rsetAux!Total = rsetAux!precio * rsetAux!cant
    rsetAux.Update
    Set grxDatos.ADORecordset = rsetAux
    
    If grxDatos.RowCount >= 1 Then
        fraUbicacion.Enabled = False
    Else
        fraUbicacion.Enabled = True
    End If
    
   ' Call Total_documento
    Call ConfiguraGrilla_Detalle
    
    Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub


'''************************************************************ELIMINA ARTICULO DEL DETALLE DE LA FACTURA****************************
Private Sub EliminaProducto()
    If grxDatos.RowCount = 0 Then Exit Sub
    Dim i As Integer
    Dim rstAux  As ADODB.Recordset
    grxDatos.Update
    Set rstAux = grxDatos.ADORecordset
    'rstAux.AbsolutePosition = grxDatos.RowIndex(grxDatos.Row)
    'rstAux.Delete
    'rstAux.Update
    rstAux.MoveFirst
    i = 1
    Do While i <= rstAux.RecordCount
        
        If rstAux("ELI").Value = True Then
          rstAux.AbsolutePosition = grxDatos.RowIndex(grxDatos.Row)
          rstAux.Delete
        Else
          rstAux("ELI") = 0
        End If
        rstAux.MoveNext
        i = i + 1
    Loop
    'rstAux.Update
    Set grxDatos.ADORecordset = rstAux
    
    If grxDatos.RowCount >= 1 Then
       fraUbicacion.Enabled = False
    Else
       fraUbicacion.Enabled = True
    End If
    'Call Total_documento
    'Call TotalizaDocumentoVenta
    'Call ConfiguraGrilla_Detalle
    
End Sub
''''*************************************************************SUMA TOTALES DE DOCUMENTO*********************************
Private Sub Total_documento()
On Error GoTo ErrCal
    Dim Total As Double
    Dim ColIndex As Long
    Dim totalkilos As Double
    Dim merma As Double
    Dim mermavar As Variant
    
    Dim i As Integer
    Total = 0
    totalkilos = 0
    'grxDatos.Update
    i = 1
    
    If grxDatos.RowCount >= 0 Then
    
            If grxDatos.RowCount > 0 Then
                'grxDatos.Update
            End If
            grxDatos.Refresh
            grxDatos.MoveFirst
            ColIndex = grxDatos.Col
            
            Do While i <= grxDatos.RowCount
               
              'If Not grxDatos.IsGroupItem(grxDatos.Row) = True And ColIndex > 0 Then
              If Not grxDatos.IsGroupItem(i) = True And ColIndex > 0 Then
              'If Trim(grxDatos.Value(grxDatos.Columns("codigorollo").Index)) <> "" Then
               
                Total = Total + grxDatos.Value(grxDatos.Columns("total").Index)
                totalkilos = totalkilos + grxDatos.Value(grxDatos.Columns("cant").Index)
                
              End If
              
                If i < grxDatos.RowCount Then
                    grxDatos.MoveNext
                End If
                i = i + 1
            Loop
            txt_total.Text = Total
            txt_descto.Text = totalkilos
            txt_subtotal.Text = Format(Total / iva, "####.00")
            txt_igv.Text = Format(Total - (Total / iva), "####.00")
            
            
     Else
            txt_total.Text = Total
            txt_descto.Text = totalkilos
            txt_subtotal.Text = Format(Total / iva, "####.00")
            txt_igv.Text = Format(Total - (Total / iva), "####.00")

     End If
     Exit Sub
ErrCal:
    MsgBox err.Description, vbCritical + vbOKOnly, "Cargar Calidades"
End Sub
Private Sub TotalizaDocumentoVenta()
On Error GoTo ErrCal

    Dim Total As Double
    Dim ColIndex As Long
    Dim totalkilos As Double
    Dim rds As New ADODB.Recordset
    
    Dim i As Integer
    Total = 0
    totalkilos = 0
    i = 1
    grxDatos.Groups.Clear
    grxDatos.Refresh
    grxDatos.Update
    
    
    If grxDatos.RowCount > 0 Then
            
            Set rds = grxDatos.ADORecordset
            rds.Update
            grxDatos.MoveFirst

            Do While Not rds.EOF
                   
                   Total = Total + rds("total").Value
                   totalkilos = totalkilos + rds("cant").Value
                   rds.MoveNext
            Loop
            txt_total.Text = Total
            txt_descto.Text = totalkilos
            
            If iva = 0 Then
                txt_subtotal.Text = 0
                txt_igv.Text = 0
            Else
                
                txt_subtotal.Text = Format(Total / iva, "####.00")
                txt_igv.Text = Format(Total - (Total / iva), "####.00")
                
            End If
            
     Else
            txt_total.Text = Total
            txt_descto.Text = totalkilos
            
            If iva = 0 Then
                txt_subtotal.Text = 0
                txt_igv.Text = 0
            Else
                txt_subtotal.Text = Format(Total / iva, "####.00")
                txt_igv.Text = Format(Total - (Total / iva), "####.00")
            
            End If

     End If
     Exit Sub
ErrCal:
    MsgBox err.Description, vbCritical + vbOKOnly, "Cargar Calidades"
End Sub

Private Sub Total_documento_actual()
On Error GoTo ErrCal
    Dim Total As Double
    Dim ColIndex As Long
    Dim totalkilos As Double
    Dim merma As Double
    Dim mermavar As Variant
    Dim rds As New ADODB.Recordset
    
    Dim i As Integer
    Total = 0
    totalkilos = 0
    'grxDatos.Update
    i = 1
    grxDatos.Update
    grxDatos.Groups.Clear
    
    If grxDatos.RowCount > 0 Then
    
            If grxDatos.RowCount > 0 Then
                'grxDatos.Update
            End If
            
            grxDatos.Refresh
            grxDatos.MoveFirst
            ColIndex = grxDatos.Col
            grxDatos.Update
                     
            
            Set rds = grxDatos.ADORecordset
            rds.Update
            If rds.RecordCount <= 0 Then Exit Sub
            Do While i <= rds.RecordCount
                
                If rds.BOF And rds.EOF Then
                    Exit Do
                Else
                
                   Total = Total + rds("total").Value
                   totalkilos = totalkilos + rds("cant").Value
                
                    If i < rds.RecordCount Then
                         rds.MoveNext
                    End If
                    i = i + 1
                
                End If
            
            Loop
            txt_total.Text = Total
            txt_descto.Text = totalkilos
            
            If iva = 0 Then
                txt_subtotal.Text = 0
                txt_igv.Text = 0
            Else
                txt_subtotal.Text = Format(Total / iva, "####.00")
                txt_igv.Text = Format(Total - (Total / iva), "####.00")
            End If
            
     Else
            txt_total.Text = Total
            txt_descto.Text = totalkilos
            
            If iva = 0 Then
                txt_subtotal.Text = 0
                txt_igv.Text = 0
            Else
                txt_subtotal.Text = Format(Total / iva, "####.00")
                txt_igv.Text = Format(Total - (Total / iva), "####.00")
            
            End If

     End If
     Exit Sub
ErrCal:
    MsgBox err.Description, vbCritical + vbOKOnly, "Cargar Calidades"
End Sub

'''*******************EVENTOS POR COLUMNA **********************************************************
Private Sub grxDatos_AfterColEdit(ByVal ColIndex As Integer)
  If grxDatos.RowCount <= 0 Then Exit Sub
  AfterColEdit_DETALLE_FACTURA (ColIndex)
End Sub

Sub AfterColEdit_DETALLE_FACTURA(ByVal ColIndex As Integer)

Dim sSQL As String
On Error GoTo Error_Handler

Dim oGroup As GridEX20.JSGroup
Select Case ColIndex

  Case Is = grxDatos.Columns("ELI").Index
        If flg_Tiene_guias_asignadas = "N" Then
             Call EliminaProducto
             Call TotalizaDocumentoVenta
             Call ConfiguraGrilla_Detalle
        End If
  Case Is = grxDatos.Columns("PRECIO").Index
  
        If IsNumeric(grxDatos.Value(grxDatos.Columns("PRECIO").Index)) = False Or grxDatos.Value(grxDatos.Columns("PRECIO").Index) = "" Then
            grxDatos.Value(grxDatos.Columns("PRECIO").Index) = 0
        End If
        
        If ind_aplica_precio = 0 Then
             grxDatos.Value(grxDatos.Columns("TOTAL").Index) = grxDatos.Value(grxDatos.Columns("PRECIO").Index) * grxDatos.Value(grxDatos.Columns("CANT").Index)
        Else
           If IsNumeric(grxDatos.Value(grxDatos.Columns("PRECIO").Index)) = True And grxDatos.Value(grxDatos.Columns("precio").Index) > 0 Then
                Call aplica_precio_tipo(grxDatos.Value(grxDatos.Columns("OT").Index), grxDatos.Value(grxDatos.Columns("precio").Index))
           End If
        End If
        Call TotalizaDocumentoVenta
        Call ConfiguraGrilla_Detalle
    
  Case Is = grxDatos.Columns("CANT").Index
  
        If IsNumeric(grxDatos.Value(grxDatos.Columns("CANT").Index)) = False Or grxDatos.Value(grxDatos.Columns("CANT").Index) = "" Then
             grxDatos.Value(grxDatos.Columns("CANT").Index) = 0
        End If
         
        grxDatos.Value(grxDatos.Columns("TOTAL").Index) = grxDatos.Value(grxDatos.Columns("PRECIO").Index) * grxDatos.Value(grxDatos.Columns("CANT").Index)
        Call TotalizaDocumentoVenta
        Call ConfiguraGrilla_Detalle
    
  End Select
  
Exit Sub

Resume
Error_Handler:
errores err.Number
End Sub
Private Sub aplica_precio_tipo(partida As String, precio As Double)
Dim rsPrecios As New ADODB.Recordset
On Error GoTo fin

grxDatos.Update
Set rsPrecios = grxDatos.ADORecordset
rsPrecios.Update
rsPrecios.MoveFirst

If ind_aplica_precio = 1 Then
    Do While Not rsPrecios.EOF

       If Trim(partida) = Trim(rsPrecios!OT) Then
            rsPrecios!precio = precio
            rsPrecios!Total = precio * rsPrecios!cant
       End If
       rsPrecios.MoveNext
    Loop
End If

If ind_aplica_precio = 2 Then
    Do While Not rsPrecios.EOF
            rsPrecios!precio = precio
            rsPrecios!Total = precio * rsPrecios!cant
            rsPrecios.MoveNext
    Loop
End If

Set grxDatos.ADORecordset = rsPrecios
grxDatos.Update

Exit Sub
fin:
MsgBox "Incovenientes para Asignar Precios", vbInformation + vbOKOnly, "Mensaje"
End Sub
Private Sub GrxProductos_AfterColEdit(ByVal ColIndex As Integer)
  AfterColEdit_PRODUCTOS (ColIndex)
End Sub

Sub AfterColEdit_PRODUCTOS(ByVal ColIndex As Integer)
Dim sSQL As String
On Error GoTo Error_Handler

Dim oGroup As GridEX20.JSGroup
Select Case ColIndex

  'Case Is = GrxProductos.Columns("SEL").Index
  '  Call adicionarProducto
  'Case Is = GrxProductos.Columns("PRECIO").Index
   ' If IsNumeric(GrxProductos.Value(GrxProductos.Columns("PRECIO").Index)) = False Or GrxProductos.Value(GrxProductos.Columns("PRECIO").Index) = "" Then
   '     GrxProductos.Value(GrxProductos.Columns("PRECIO").Index) = 0
   ' End If
    'GrxProductos.Value(GrxProductos.Columns("TOTAL").Index) = GrxProductos.Value(GrxProductos.Columns("PRECIO").Index) * GrxProductos.Value(GrxProductos.Columns("CANT").Index)
    'GridEX1.Value(GridEX1.Columns("sel").Index) = 0
  Case Is = GrxProductos.Columns("CANT").Index
    If IsNumeric(GrxProductos.Value(GrxProductos.Columns("CANT").Index)) = False Or GrxProductos.Value(GrxProductos.Columns("CANT").Index) = "" Then
        GrxProductos.Value(GrxProductos.Columns("CANT").Index) = 0
    End If
    GrxProductos.Value(GrxProductos.Columns("TOTAL").Index) = GrxProductos.Value(GrxProductos.Columns("PRECIO").Index) * GrxProductos.Value(GrxProductos.Columns("CANT").Index)
    'GridEX1.Value(GridEX1.Columns("sel").Index) = 0
  End Select
Exit Sub

Resume
Error_Handler:
errores err.Number
End Sub
'''***************************************evento click de las grillas  **********************************

Private Sub grxDatos_Click()
   On Error GoTo fin
    Dim ColIndex As Long
    Dim oRowData As JSRowData
    Dim SGRUPO As String
    Dim iRow As Long
    Dim i As Long
    Dim sCaptionGroup As String
        If grxDatos.RowCount > 0 Then
        ColIndex = grxDatos.Col
       
            If UCase(grxDatos.Columns(ColIndex).Key) = "ELI" Then
                bClickColSelec = True
                SendKeys "{ENTER}"
'            ElseIf UCase(grxDatos.Columns(ColIndex).Key) = "CANT" Then
'                If IsNumeric(grxDatos.Value(grxDatos.Columns("CANT").Index)) = False Then
'                    grxDatos.Value(grxDatos.Columns("CANT").Index) = 0
'                End If
            End If
    End If
Exit Sub
fin:
MsgBox "Inconvenientes", vbInformation + vbOKOnly, "Mensaje"
    
End Sub
Private Sub GrxProductos_Click()

    Dim ColIndex As Long
    Dim oRowData As JSRowData
    Dim SGRUPO As String
    Dim iRow As Long
    Dim i As Long
    Dim sCaptionGroup As String
    
        If GrxProductos.RowCount > 0 Then
        ColIndex = GrxProductos.Col
       
            'If UCase(GrxProductos.Columns(ColIndex).Key) = "SEL" Then
             '   bClickColSelec = True
             '   SendKeys "{ENTER}"
            'End If
    End If
End Sub
'''*******************************************************************************************
Private Sub txtBus_Cod_ordtra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call buscarProductos(4)
End If

End Sub

Private Sub txtBus_Codigo_RolloTinto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call buscarProductos(1)
End If
End Sub

Private Sub txtBus_Des_Color_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call buscarProductos(5)
End If

End Sub

Private Sub txtCod_ConPag_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  Call Busca_Opcion("Cod_CondVent", "Des_CondVent", "Lg_CondVent where ", txtCod_ConPag, txtDes_ConPag, 1)
  If Trim(txtDes_ConPag.Text) <> "" Then
    txtCodigo_Producto.SetFocus
  Else
    txtCod_ConPag.SetFocus
  End If
  
  End If
End Sub
Private Sub txtCod_Moneda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  Call Busca_Opcion("Cod_Moneda", "Nom_Moneda", "TG_Moneda where ", txtCod_Moneda, txtDes_Moneda, 1)
  
  If Trim(txtDes_Moneda.Text) <> "" Then
     txtCod_TipVenta.SetFocus
  Else
     txtCod_Moneda.SetFocus
  End If
  
  End If
End Sub

Private Sub txtCod_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion("Cod_TipAnex", "Des_Tipanex", "CN_TipoAnexoContable where ", txtCod_TipAne, txtDes_TipAne, 1)
End Sub
Private Sub txtCod_TipDoc_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
    'Call Busca_Opcion("Cod_TipDoc", "Des_TipDoc", "CN_TiposDocum where Flg_Doc_Ventas = '*' and ", txtCod_TipDoc, txtDes_TipDoc, 1)
    'Cambio_FR
    'If txtCod_TipDoc = "BV" Then txtCod_TipAne = ""
    Call buscaDocumentos(1)
    
    If Trim(txtDes_TipDoc.Text) <> "" Then
      txtCod_Moneda.SetFocus
    Else
      txtCod_TipDoc.SetFocus
    End If
    
  End If
  
End Sub

Private Sub txtCod_TipDoc_LostFocus()
  Cambio_FR
End Sub

Private Sub txtCod_TipVenta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  Call Busca_Opcion("Cod_Tipo_Venta", "Descripcion", "Cn_Tipos_Venta where ", txtCod_TipVenta, txtDes_TipVenta, 1)
  
  If Trim(txtDes_TipVenta.Text) <> "" Then
    txtDes_TipAne.SetFocus
  Else
     txtCod_TipVenta.SetFocus
  End If
  
  End If
  
'    If gfVerificar_ExisteRegistroTabla("Cn_Ventas_Motivos_Notas_Abonos", "Cod_TipDoc ='" & txtCod_TipDoc & "'", cCONNECT) = eNoExiste Then
End Sub
Private Sub txtDescripcion_Producto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call buscarProductos(3)
    End If
    
End Sub
''''*************************************************************BUSQUEDA DE PRODUCTOS *********************************
Private Sub buscarProductos(Opcion As String)

Dim StrSQL As String
Dim sCodCentroCosto As String
Dim nrofilas As Integer
Dim k, l As Long
Dim rsproductos  As New ADODB.Recordset
On Error GoTo fin
   
    StrSQL = "TX_MUESTRA_ROLLOS_VENTA '" & Opcion & "','" & Trim(txtCod_Almacen.Text) & "','" & Trim(txtBus_Codigo_RolloTinto.Text) & "','" & Trim(txtBus_Cod_ordtra.Text) & "','" & Trim(txtDescripcion_Producto.Text) & "','" & Trim(txtBus_Des_Color.Text) & "'"
    
    Set GrxProductos.ADORecordset = Nothing
    Set GrxProductos.ADORecordset = CargarRecordSetDesconectado(StrSQL, cConnect)
    If GrxProductos.RowCount <= 0 Then Exit Sub

    GrxProductos.Update
    Set rsproductos = GrxProductos.ADORecordset
    rsproductos.Update
    rsproductos.MoveFirst
    Do While Not rsproductos.EOF
       rsproductos!Stock = rsproductos!Stock - SumaTotalRollo(Trim(rsproductos!codigoRollo))
       rsproductos.MoveNext
    Loop

    Set GrxProductos.ADORecordset = rsproductos
    GrxProductos.Update
   
    Call eliminaRolloCeroNegativo
   
'    nrofilas = GrxProductos.RowCount
'    If nrofilas > 0 Then
'            nrofilas = 15
'    Else
'       FraProductos.Visible = True
'    End If

    Call ConfiguraGrilla_productos
    Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
Private Sub eliminaRolloCeroNegativo()
    Dim rsproductos As New ADODB.Recordset
    Dim u As Long
    Dim neg As String
On Error GoTo fin
    
    GrxProductos.Update
    Set rsproductos = GrxProductos.ADORecordset
    rsproductos.MoveFirst
    u = 1
    neg = "N"
    Do While Not rsproductos.EOF
        If rsproductos!Stock <= 0 Then
           neg = "S"
           rsproductos.AbsolutePosition = u
           rsproductos.Delete
           Exit Do
        End If
        rsproductos.MoveNext
        u = u + 1
    Loop
    If neg = "S" Then
        eliminaRolloCeroNegativo
    End If
   Set GrxProductos.ADORecordset = rsproductos
   
Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
Private Function SumaTotalRollo(codigoRollo As String) As Double
On Error GoTo fin
    Dim rssum As New ADODB.Recordset
    Dim pesorollo  As Double
    If grxDatos.RowCount <= 0 Then Exit Function
    pesorollo = 0
    grxDatos.Update
    Set rssum = Nothing
    Set rssum = grxDatos.ADORecordset
    rssum.Update
    rssum.MoveFirst
    Do While Not rssum.EOF
        If Trim(codigoRollo) = Trim(rssum!codigoRollo) Then
             pesorollo = pesorollo + rssum!cant
        End If
        rssum.MoveNext
    Loop
    rssum.MoveFirst
    rssum.Update
      
    SumaTotalRollo = pesorollo
Exit Function
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Function
''''*******************************************CONfigura GRILLA PRODUCTOS*********************************
Private Sub ConfiguraGrilla_productos()
    Dim C As Integer
    Dim colTemp As JSColumn
    Dim fmtCon  As JSFmtCondition

    On Error GoTo fin
    
    With GrxProductos
        
        For C = 1 To .Columns.Count
            .Columns(C).HeaderAlignment = jgexAlignCenter
            .Columns(C).TextAlignment = jgexAlignLeft
            .Columns(C).Visible = False
            
        Next C
 
        .Columns("OT").Width = 700
        .Columns("OT").Visible = True
        .Columns("codigorollo").Width = 1150
        .Columns("codigorollo").Visible = True
        .Columns("codigorollo").Caption = "CODIGO"
        .Columns("TELA").Width = 5500
        .Columns("TELA").Visible = True
        .Columns("COLOR").Width = 2000
        .Columns("COLOR").Visible = True
        .Columns("CALIDAD").Width = 500
        .Columns("CALIDAD").Visible = True
        
        .Columns("rollos").Width = 500
        .Columns("rollos").Visible = True
        .Columns("rollos").Caption = "ROL."
        .Columns("UND").Width = 500
        .Columns("UND").Visible = True
        
        .Columns("CALIDAD").Caption = "CAL"
        .Columns("STOCK").Width = 1000
        .Columns("STOCK").Visible = True
        .Columns("STOCK").Caption = "STOCK"
        .Columns("STOCK").TextAlignment = jgexAlignRight
        .Columns("CANT").Width = 1000
        .Columns("CANT").Visible = True
        .Columns("CANT").TextAlignment = jgexAlignRight
        
        .Columns("total").Width = 1000
        .Columns("total").Visible = True
        .Columns("TOTAL").TextAlignment = jgexAlignRight
        
        
        .Columns("PRECIO").Width = 1000
        .Columns("PRECIO").Visible = True
        .Columns("PRECIO").TextAlignment = jgexAlignRight
        
        .Columns("sel").Width = 800
        .Columns("Sel").Visible = False
        
        If .RowCount > 0 Then
            .Row = -1
            .Col = .Columns.Count - 1
        End If
        '.SetFocus
    End With
    
    Dim oGroup01 As GridEX20.JSGroup
    Dim oGroup02 As GridEX20.JSGroup
    Dim valorcant   As JSColumn
    Dim valorStock   As JSColumn
    
      With GrxProductos
            
        Set oGroup01 = .Groups.Add(.Columns("OT").Index, jgexSortAscending)
        .DefaultGroupMode = jgexDGMExpanded
        .BackColorRowGroup = RGB(239, 235, 222)
           
           .GroupFooterStyle = jgexTotalsGroupFooter
           Set valorcant = .Columns("CANT")
           
           With valorcant
               .AggregateFunction = jgexSum
               .TotalRowPrefix = "Total: "
               .TextAlignment = jgexAlignRight
           End With
           
           .GroupFooterStyle = jgexTotalsGroupFooter
           Set valorStock = .Columns("STOCK")
           
           With valorStock
               .AggregateFunction = jgexSum
               .TotalRowPrefix = "T: "
               .TextAlignment = jgexAlignRight
           End With

        End With
        
    Set fmtCon = GrxProductos.FmtConditions.Add(GrxProductos.Columns("CANT").Index, jgexGreaterThan, 0)
    fmtCon.FormatStyle.BackColor = &H80FFFF   ' &HFFFF00

    SetColores
    
    Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
Private Sub SetColores()
        GrxProductos.Columns("CANT").CellStyle = "Color_Cantidad"
        'GrxProductos.Columns("ROLLOS").CellStyle = "Color_Cantidad"
End Sub

''''*******************************************BUSCA DETALLE INICIAL DE DOCUMENTO*********************************
Private Sub buscaDetalle_factura()

    Dim StrSQL As String
    Dim sCodCentroCosto As String
    Dim nrofilas As Integer
    
    On Error GoTo fin
   
    StrSQL = "EXEC CN_MUESTRA_TELAS_DETALLE_FACTURA 'xx','" & Trim(txtCodigo_Producto.Text) & "','" & Trim(txtDescripcion_Producto.Text) & "'"
    
    Set grxDatos.ADORecordset = Nothing
    Set grxDatos.ADORecordset = CargarRecordSetDesconectado(StrSQL, cConnect)
    
    Call ConfiguraGrilla_Detalle
    Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
''''*******************************************CONFIGURA DETALLE DE DOCUMENTO*********************************
Private Sub ConfiguraGrilla_Detalle()
    Dim C As Integer
    On Error GoTo fin
    
    'Call ConfiguraGrilla_DetalleSinGrupos
    
       ' If grxDatos.RowCount > 0 Then
       '     Call Total_documento
       ' End If
        
    With grxDatos
        
        For C = 1 To .Columns.Count
            .Columns(C).HeaderAlignment = jgexAlignCenter
            .Columns(C).TextAlignment = jgexAlignLeft
            .Columns(C).Visible = False
        Next C
        
        .Columns("OT").Width = 800
        .Columns("OT").Visible = True
        .Columns("CODIGOROLLO").Width = 1500
        .Columns("CODIGOROLLO").Visible = True
        
        .Columns("TELA").Width = 5500
        .Columns("TELA").Visible = True
        .Columns("COLOR").Width = 2000
        .Columns("COLOR").Visible = True
        .Columns("CALIDAD").Width = 500
        .Columns("CALIDAD").Visible = True
        .Columns("ROLLOS").Width = 800
        .Columns("ROLLOS").Visible = True
        
        .Columns("UND").Width = 500
        .Columns("UND").Visible = True
        
        .Columns("CALIDAD").Caption = "CAL"
        
        .Columns("stock").Width = 1000
        .Columns("stock").Visible = True
        .Columns("stock").Caption = "STOCK"
        .Columns("stock").TextAlignment = jgexAlignRight
        
        .Columns("CANT").Width = 1000
        .Columns("CANT").Visible = True
        .Columns("CANT").Caption = "CANT"
        .Columns("cant").TextAlignment = jgexAlignRight
        
        .Columns("PRECIO").Width = 1000
        .Columns("PRECIO").Visible = True
        .Columns("PRECIO").TextAlignment = jgexAlignRight
        
        .Columns("TOTAL").Width = 1000
        .Columns("TOTAL").Visible = True
        .Columns("TOTAL").TextAlignment = jgexAlignRight
        
        .Columns("ELI").Width = 250
        .Columns("ELI").Visible = True
        .Columns("ELI").Caption = ""
        
        .Columns("DEL").Width = 250
        .Columns("DEL").Visible = True
        .Columns("DEL").Caption = "X"
        .Columns("DEL").TextAlignment = jgexAlignCenter
        SetColorDetalle
        
    End With
    
    Dim oGroup01 As GridEX20.JSGroup
    Dim oGroup02 As GridEX20.JSGroup
    Dim oGroup03 As GridEX20.JSGroup
    
    Dim valorcant    As JSColumn
    Dim valorStock   As JSColumn
    Dim ItemTotal    As JSColumn
    
      With grxDatos
            
        Set oGroup01 = .Groups.Add(.Columns("OT").Index, jgexSortAscending)
         
        .DefaultGroupMode = jgexDGMExpanded
        .BackColorRowGroup = RGB(239, 235, 222)
           
           .GroupFooterStyle = jgexTotalsGroupFooter
           Set valorcant = .Columns("CANT")
           Set valorStock = .Columns("STOCK")
           Set ItemTotal = .Columns("TOTAL")

           With valorcant
               .AggregateFunction = jgexSum
               .TotalRowPrefix = "T: "
               .TextAlignment = jgexAlignRight
           End With

           With valorStock
               .AggregateFunction = jgexSum
               .TotalRowPrefix = "T: "
               .TextAlignment = jgexAlignRight
           End With

           With ItemTotal
               .AggregateFunction = jgexSum
               .TotalRowPrefix = "T: "
               .TextAlignment = jgexAlignRight
           End With
'
'           If .RowCount > 0 Then
'                .Row = -1
'                .Col = .Columns.Count - 1
'           End If
        End With

    Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
Private Sub SetColorDetalle()
        'grxDatos.Columns("ROLLOS").CellStyle = "estilo_cantidad"
        grxDatos.Columns("CANT").CellStyle = "estilo_cantidad"
        grxDatos.Columns("PRECIO").CellStyle = "estilo_cantidad"
        grxDatos.Columns("ELI").CellStyle = "estilo_eliminar"
        grxDatos.Columns("DEL").CellStyle = "estilo_eliminar"
End Sub
Private Sub txtDes_ConPag_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  
    Call Busca_Opcion("Cod_CondVent", "Des_CondVent", "Lg_CondVent where ", txtCod_ConPag, txtDes_ConPag, 2)
    If Trim(txtDes_ConPag.Text) <> "" Then
      txtCodigo_Producto.SetFocus
    Else
      txtDes_ConPag.SetFocus
    End If
  End If
End Sub

Private Sub txtDes_Moneda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_Moneda", "Nom_Moneda", "TG_Moneda where ", txtCod_Moneda, txtDes_Moneda, 2)
    If Trim(txtDes_Moneda.Text) <> "" Then
       txtCod_TipVenta.SetFocus
    Else
       txtDes_Moneda.SetFocus
    End If
  
  End If
End Sub

Private Sub txtDes_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        Load FrmBusqClientes
        Set FrmBusqClientes.oParent = Me
        FrmBusqClientes.txtDescripcion_Cliente.Text = txtDes_TipAne.Text
        FrmBusqClientes.txtRuc_Cliente.Text = "" 'txtNum_Ruc.Text
        FrmBusqClientes.txtTip_Anex.Text = "C"
        
        Call FrmBusqClientes.Busca_Opcion_AnexoContable("2", "C", txtNum_Ruc.Text, txtDes_TipAne.Text)
        FrmBusqClientes.Show 1
        'FrmBusqClientes.txtDescripcion_Cliente.SetFocus
        Set FrmBusqClientes = Nothing
        
        If Trim(txtNum_Ruc.Text) <> "" Then
           txtCod_ConPag.SetFocus
        Else
           txtDes_TipAne.SetFocus
        End If
        
  End If
End Sub
Private Sub txtDes_TipDoc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    'Call Busca_Opcion("Cod_TipDoc", "Des_TipDoc", "CN_TiposDocum where Flg_Doc_Ventas = '*' and ", txtCod_TipDoc, txtDes_TipDoc, 2)
    'Cambio_FR
    Call buscaDocumentos(2)
    If Trim(txtDes_TipDoc.Text) <> "" Then
        txtCod_Moneda.SetFocus
    Else
        txtDes_Moneda.SetFocus
    End If
    
  End If
End Sub
Private Sub txtCod_Vendedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    buscaVendedor (1)
    
    If txtDes_Vendedor.Text <> "" Then
       txtCod_Almacen.SetFocus
    Else
       txtCod_Vendedor.SetFocus
    End If
    
End If
End Sub
Private Sub txtDes_Vendedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    buscaVendedor (2)
    If txtDes_Vendedor.Text <> "" Then
       txtCod_Almacen.SetFocus
    Else
       txtDes_Vendedor.SetFocus
    End If
    
End If
End Sub

Public Sub buscaVendedor(sOpcion As String)
On Error GoTo fin
Dim iCol As Long
Dim rstAux As New ADODB.Recordset
Dim Opcion As String
  StrSQL = "CN_MUESTRA_VENDEDOR_CAJAS '" & sOpcion & "','" & Trim(txtCod_Tienda.Text) & "','" & Trim(txtCod_Caja.Text) & "','" & Trim(txtCod_Vendedor.Text) & "','" & Trim(txtDes_Vendedor.Text) & "'"

    With frmBusqGeneralOperario
        Set .oParent = Me
        .SQuery = StrSQL
        .Cargar_Datos
        CODIGO = ".."
        Set rstAux = .DGridLista.ADORecordset
        
        .DGridLista.Columns("Codigo").Caption = "Codigo"
        .DGridLista.Columns("Codigo").Width = 900
        .DGridLista.Columns("nombre").Caption = "Nombre"
        .DGridLista.Columns("nombre").Width = 1500
        
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If CODIGO <> "" And rstAux.RecordCount > 0 Then
            txtCod_Vendedor = Trim(rstAux!CODIGO)
            txtCod_Vendedor.Tag = Left(Trim(rstAux!CODIGO), 1)
            txtDes_Vendedor.Text = Trim(rstAux!Nombre)
            txtDes_Vendedor.Tag = Right(Trim(rstAux!CODIGO), 4)
        End If
    End With
    Unload frmBusqGeneralOperario
    Set frmBusqGeneralOperario = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
fin:
On Error Resume Next
    Unload frmBusqGeneralOperario
    Set frmBusqGeneralOperario = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Vendedor(" & Opcion & ")"
End Sub
Public Sub BuscaCliente(sOpcion As String)
On Error GoTo fin
Dim iCol As Long
Dim rstAux As New ADODB.Recordset
Dim Opcion As String

  StrSQL = "CN_MUESTRA_VENDEDOR_CAJAS '" & sOpcion & "','" & Trim(txtCod_Tienda.Text) & "','" & Trim(txtCod_Caja.Text) & "','" & Trim(txtCod_Vendedor.Text) & "','" & Trim(txtDes_Vendedor.Text) & "'"

    With frmBusqGeneralOperario
        Set .oParent = Me
        .SQuery = StrSQL
        .Cargar_Datos
        CODIGO = ".."
        Set rstAux = .DGridLista.ADORecordset
        
        .DGridLista.Columns("Codigo").Caption = "Codigo"
        .DGridLista.Columns("Codigo").Width = 900
        .DGridLista.Columns("nombre").Caption = "Nombre"
        .DGridLista.Columns("nombre").Width = 1500
        
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If CODIGO <> "" And rstAux.RecordCount > 0 Then
            txtCod_Vendedor = Trim(rstAux!CODIGO)
            txtCod_Vendedor.Tag = Left(Trim(rstAux!CODIGO), 1)
            txtDes_Vendedor.Text = Trim(rstAux!Nombre)
            txtDes_Vendedor.Tag = Right(Trim(rstAux!CODIGO), 4)
        End If
    End With
    Unload frmBusqGeneralOperario
    Set frmBusqGeneralOperario = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
fin:
On Error Resume Next
    Unload frmBusqGeneralOperario
    Set frmBusqGeneralOperario = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Vendedor(" & Opcion & ")"
End Sub

Private Sub txtCod_Almacen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        buscaAlmacen (1)
        If txtDes_Almacen.Text <> "" Then
           txtCod_TipDoc.SetFocus
        Else
           txtCod_Almacen.SetFocus
        End If
        
    End If
End Sub
Private Sub txtDES_Almacen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        buscaAlmacen (2)
        If txtDes_Almacen.Text <> "" Then
           txtCod_TipDoc.SetFocus
        Else
           txtCod_Almacen.SetFocus
        End If
    End If
End Sub
Public Sub buscaDocumentos(sOpcion As String)
On Error GoTo fin
Dim iCol As Long
Dim rstAux As New ADODB.Recordset
Dim Opcion As String
  fila_seleccionada = 0
  StrSQL = "CN_MUESTRA_VENTAS_CAJAS_DOCUMENTOS  '" & sOpcion & "','" & Trim(txtCod_Tienda.Text) & "','" & Trim(txtCod_Caja.Text) & "','" & Trim(txtCod_TipDoc.Text) & "','" & Trim(txtDes_TipDoc.Text) & "'"
  With frmBusqGeneral
        Set .oParent = Me
        .SQuery = StrSQL
        .Cargar_Datos
        CODIGO = ".."
        Set rstAux = .gexList.ADORecordset
        
        .gexList.Columns("Cod_TipDoc").Caption = "Codigo"
        .gexList.Columns("Cod_TipDoc").Width = 1000
        .gexList.Columns("DES_TIPDOC").Caption = "Almacen"
        .gexList.Columns("DES_TIPDOC").Width = 4000
        
        If rstAux.RecordCount > 1 Then .Show vbModal
        If fila_seleccionada > 0 And rstAux.RecordCount > 0 Then
            rstAux.AbsolutePosition = fila_seleccionada
            txtCod_TipDoc.Text = Trim(rstAux!Cod_TipDoc)
            txtDes_TipDoc.Text = Trim(rstAux!Des_TipDoc)
            txtSer_Docum.Text = Trim(rstAux!Serie)
            txtNum_Docum.Text = Trim(rstAux!Nroactual)
         Else
            txtCod_TipDoc.Text = ""
            txtDes_TipDoc.Text = ""
            txtSer_Docum.Text = ""
            txtNum_Docum.Text = ""
        End If
        
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ",No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Documento(" & Opcion & ")"
End Sub
Public Sub buscaAlmacen(sOpcion As String)
On Error GoTo fin
Dim iCol As Long
Dim rstAux As New ADODB.Recordset
Dim Opcion As String
  StrSQL = "CN_MUESTRA_VENTAS_CAJAS_ALMACEN  '" & sOpcion & "','" & Trim(txtCod_Tienda.Text) & "','" & Trim(txtCod_Caja.Text) & "','" & Trim(txtCod_Almacen.Text) & "','" & Trim(txtDes_Almacen.Text) & "'"
  With frmBusqGeneral
        Set .oParent = Me
        .SQuery = StrSQL
        .Cargar_Datos
        CODIGO = ".."
        Set rstAux = .gexList.ADORecordset
        
        .gexList.Columns("cod_almacen").Caption = "Codigo"
        .gexList.Columns("cod_almacen").Width = 1000
        .gexList.Columns("nom_almacen").Caption = "Almacen"
        .gexList.Columns("nom_almacen").Width = 4000
        
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If CODIGO <> "" And rstAux.RecordCount > 0 Then
            txtCod_Almacen.Text = Trim(rstAux!Cod_almacen)
            txtDes_Almacen.Text = Trim(rstAux!nom_almacen)
        End If
        
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ",No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de almcen(" & Opcion & ")"
End Sub

Private Sub txtCod_fabrica_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  Call Busca_Opcion("cod_fabrica", "nom_fabrica", "tg_empresa where ", TxtCod_Fabrica, txtDes_Fabrica, 1)
    If Trim(txtDes_Fabrica.Text) <> "" Then
       txtCod_Tienda.SetFocus
    Else
       TxtCod_Fabrica.SetFocus
    End If
    
  End If
  
End Sub
Private Sub txtdes_fabrica_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        Call Busca_Opcion("cod_fabrica", "nom_fabrica", "tg_empresa where ", TxtCod_Fabrica, txtDes_Fabrica, 2)
        If Trim(txtDes_Fabrica.Text) <> "" Then
           txtCod_Tienda.SetFocus
        Else
           txtDes_Fabrica.SetFocus
        End If
  End If
End Sub
Private Sub txtcod_tienda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        Call Busca_Opcion("cod_tienda", "des_tienda", "cn_ventas_tiendas where ", txtCod_Tienda, txtDes_Tienda, 1)
        If Trim(txtDes_Tienda.Text) <> "" Then
          txtCod_Caja.SetFocus
        Else
          txtCod_Tienda.SetFocus
        End If
  End If
End Sub
Private Sub txtDES_tienda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        Call Busca_Opcion("cod_tienda", "des_tienda", "cn_ventas_tiendas where ", txtCod_Tienda, txtDes_Tienda, 2)
        If Trim(txtDes_Tienda.Text) <> "" Then
            txtCod_Caja.SetFocus
        Else
            txtDes_Tienda.SetFocus
        End If
  End If
End Sub
Private Sub txtcod_caja_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        Call Busca_Opcion("COD_CAJA", " 'CAJA NRO ' + COD_CAJA ", "CN_VENTAS_CAJAS where cod_tienda= '" & txtCod_Tienda.Text & "' and ", txtCod_Caja, txtDes_Caja, 1)
        If Trim(txtDes_Tienda.Text) <> "" Then
             txtCod_Vendedor.SetFocus
        Else
             txtCod_Caja.SetFocus
        End If
   End If
End Sub
Private Sub txtDES_caja_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     Call Busca_Opcion("COD_CAJA", " 'CAJA NRO ' + COD_CAJA ", "CN_VENTAS_CAJAS where cod_tienda= '" & txtCod_Tienda.Text & "' and ", txtCod_Caja, txtDes_Caja, 2)
     If Trim(txtDes_Tienda.Text) <> "" Then
          txtCod_Vendedor.SetFocus
     Else
          txtDes_Vendedor.SetFocus
     End If
  End If
  
End Sub

Private Sub txtDes_TipVenta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  Call Busca_Opcion("Cod_Tipo_Venta", "Descripcion", "Cn_Tipos_Venta where ", txtCod_TipVenta, txtDes_TipVenta, 2)
  
  If Trim(txtDes_TipVenta.Text) <> "" Then
     txtDes_TipAne.SetFocus
  Else
     txtDes_TipVenta.SetFocus
  End If
  
  End If
End Sub


Private Sub txtNum_Docum_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    'txtNum_Docum.Text = Format(txtNum_Docum.Text, "00000000")
    SendKeys "{TAB}"
  End If
  
  If KeyAscii >= 48 And KeyAscii <= 57 _
    Or KeyAscii = 8 Or KeyAscii = 13 Then Else KeyAscii = 0
End Sub

Private Sub txtNum_Docum_LostFocus()
  txtNum_Docum = Format(txtNum_Docum, "00000000")
End Sub
Private Sub txtNum_Ruc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        Load FrmBusqClientes
        Set FrmBusqClientes.oParent = Me
        FrmBusqClientes.txtDescripcion_Cliente.Text = "" 'txtDes_TipAne.Text
        FrmBusqClientes.txtRuc_Cliente.Text = txtNum_Ruc.Text
        FrmBusqClientes.txtTip_Anex.Text = "C"
        txtDes_TipAne.Text = ""
        
        Call FrmBusqClientes.Busca_Opcion_AnexoContable("1", "C", txtNum_Ruc.Text, txtDes_TipAne.Text)
        FrmBusqClientes.Show 1
        'FrmBusqClientes.txtRuc_Cliente.SetFocus
        
        'txtDes_TipAne.Text = FrmBusqClientes.codigo
        'txtNum_Ruc.Text = FrmBusqClientes.Descripcion
      
       If Trim(txtNum_Ruc.Text) <> "" Then
          txtCod_ConPag.SetFocus
       Else
          txtNum_Ruc.SetFocus
       End If
        Set FrmBusqClientes = Nothing
  End If
 
End Sub

Private Sub txtSer_Docum_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
    'txtSer_Docum.Text = Format(txtSer_Docum, "000")
  End If
  If KeyAscii >= 48 And KeyAscii <= 57 _
    Or KeyAscii = 8 Or KeyAscii = 13 Then Else KeyAscii = 0
    
End Sub

Private Sub txtSer_Docum_LostFocus()
  txtSer_Docum = Format(txtSer_Docum, "000")
End Sub

Private Sub txtSerieGuia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Call buscalistaGuiasPendientes
      Call buscalistaGuiasSeleccionadas
    End If
End Sub

Private Sub txtNumeroGuia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      buscalistaGuiasPendientes
      Call buscalistaGuiasSeleccionadas
    End If
End Sub
''''*******************************************BUSCA guias pendientes*********************************
Private Sub buscalistaGuiasPendientes()

    Dim StrSQL As String
    Dim sCodCentroCosto As String
    Dim nrofilas As Integer
    
    On Error GoTo fin
   
    txtSerieGuia.Text = Format(txtSerieGuia, "000")
    txtNumeroGuia.Text = Format(txtNumeroGuia, "00000000")
   
    StrSQL = "EXEC CN_MUESTRA_GUIAS_PENDIENTES_FACTURACION '" & Left(cboAlmacen, 2) & "','" & Trim(txtSerieGuia.Text) & "','" & Trim(txtNumeroGuia.Text) & "','" & Trim(txtNum_Ruc.Tag) & "'"
    
    Set grxListaGuiaPendientes.ADORecordset = Nothing
    Set grxListaGuiaPendientes.ADORecordset = CargarRecordSetDesconectado(StrSQL, cConnect)
    
    Call ConfiguraGrillaListaGuiasPendientes
    Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
''''*******************************************CONFIGURA detalle guias pendientes *********************************
Private Sub ConfiguraGrillaListaGuiasPendientes()
    Dim C As Integer
    On Error GoTo fin
    
    With grxListaGuiaPendientes
        
        For C = 1 To .Columns.Count
            .Columns(C).HeaderAlignment = jgexAlignCenter
            .Columns(C).TextAlignment = jgexAlignLeft
            .Columns(C).Visible = False
        Next C
        
    
        .Columns("cod_almacen").Width = 800
        .Columns("cod_almacen").Visible = True
        .Columns("cod_almacen").Caption = "Almacen"
        
        .Columns("num_movstk").Width = 1200
        .Columns("num_movstk").Visible = True
        .Columns("num_movstk").Caption = "Mov"
        
        .Columns("fec_movstk").Width = 1000
        .Columns("fec_movstk").Visible = True
        .Columns("fec_movstk").Caption = "Fec Mov"
        
        .Columns("ser_guia").Visible = True
        .Columns("ser_guia").Width = 800
        .Columns("ser_guia").Caption = "Serie"
        
        .Columns("numero_guia").Width = 1500
        .Columns("numero_guia").Visible = True
        .Columns("numero_guia").Caption = "Numero"
        
        .Columns("cod_usuario").Visible = True
        .Columns("cod_usuario").Width = 1000
        .Columns("cod_usuario").Caption = "Usuario"
        
        
        If .RowCount > 0 Then
            .Row = -1
            .Col = .Columns.Count - 1
        End If
       
    End With
    Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
''''*******************************************BUSCA guias seleccionadas*********************************
Private Sub buscalistaGuiasSeleccionadas()

    Dim StrSQL As String
    Dim sCodCentroCosto As String
    Dim nrofilas As Integer
    
    On Error GoTo fin
    
    StrSQL = "EXEC CN_MUESTRA_GUIAS_ASOCIADAS_FACTURAS '','" & Trim(txtSer_Docum.Text) & "','" & Trim(txtNum_Docum.Text) & "','" & Trim(txtNum_Ruc.Tag) & "'"
    
    Set grxListaGuiasSeleccionadas.ADORecordset = Nothing
    Set grxListaGuiasSeleccionadas.ADORecordset = CargarRecordSetDesconectado(StrSQL, cConnect)
    
    Call ConfiguraGrillaListaGuiasSeleccionadas
    Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
''''*******************************************CONFIGURA DETALLE de guias Seleccionas*********************************
Private Sub ConfiguraGrillaListaGuiasSeleccionadas()
    Dim C As Integer
    On Error GoTo fin
    
    With grxListaGuiasSeleccionadas
        
        For C = 1 To .Columns.Count
            .Columns(C).HeaderAlignment = jgexAlignCenter
            .Columns(C).TextAlignment = jgexAlignLeft
            .Columns(C).Visible = False
        Next C
        
        .Columns("cod_almacen").Width = 800
        .Columns("cod_almacen").Visible = True
        .Columns("cod_almacen").Caption = "Almacen"
        
        .Columns("num_movstk").Width = 1200
        .Columns("num_movstk").Visible = True
        .Columns("num_movstk").Caption = "Mov"
        
        .Columns("fec_movstk").Width = 1000
        .Columns("fec_movstk").Visible = True
        .Columns("fec_movstk").Caption = "Fec Mov"
        
        .Columns("ser_guia").Visible = True
        .Columns("ser_guia").Width = 800
        .Columns("ser_guia").Caption = "Serie"
        
        .Columns("numero_guia").Width = 1500
        .Columns("numero_guia").Visible = True
        .Columns("numero_guia").Caption = "Numero"
        
        .Columns("cod_usuario").Visible = True
        .Columns("cod_usuario").Width = 1000
        .Columns("cod_usuario").Caption = "Usuario"

        
        If .RowCount > 0 Then
            .Row = -1
            .Col = .Columns.Count - 1
        End If
       
    End With
    Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub


Private Sub TxtTipo_Cambio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

