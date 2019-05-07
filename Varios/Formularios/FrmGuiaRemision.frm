VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmGuiaRemision 
   Caption         =   "G U I A   D E   R E M I S I O N"
   ClientHeight    =   8955
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   15570
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTotalKilos 
      Appearance      =   0  'Flat
      Height          =   350
      Left            =   13200
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton cmdAnularGuia 
      Caption         =   "ANULAR"
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
      Left            =   8400
      TabIndex        =   53
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Frame FraProductos 
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00000000&
      Height          =   5400
      Left            =   960
      TabIndex        =   34
      Top             =   2040
      Width           =   14535
      Begin VB.TextBox txtBus_Codigo_RolloTinto 
         BackColor       =   &H00C0FFFF&
         Height          =   350
         Left            =   1440
         TabIndex        =   42
         Top             =   120
         Width           =   2295
      End
      Begin VB.TextBox txtDescripcion_Producto 
         BackColor       =   &H00C0FFFF&
         Height          =   350
         Left            =   6720
         TabIndex        =   41
         Top             =   120
         Width           =   2655
      End
      Begin VB.TextBox txtBus_Cod_ordtra 
         BackColor       =   &H00C0FFFF&
         Height          =   350
         Left            =   4440
         TabIndex        =   40
         Top             =   120
         Width           =   1815
      End
      Begin VB.TextBox txtBus_Des_Color 
         BackColor       =   &H00C0FFFF&
         Height          =   350
         Left            =   9960
         TabIndex        =   39
         Top             =   120
         Width           =   2535
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
         TabIndex        =   38
         Top             =   120
         Width           =   800
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
         TabIndex        =   37
         Top             =   5000
         Width           =   975
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
         TabIndex        =   36
         Top             =   5000
         Width           =   975
      End
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
         TabIndex        =   35
         Top             =   240
         Width           =   855
      End
      Begin GridEX20.GridEX GrxProductos 
         Height          =   4450
         Left            =   45
         TabIndex        =   43
         Top             =   480
         Width           =   14415
         _ExtentX        =   25426
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
         Column(1)       =   "FrmGuiaRemision.frx":0000
         Column(2)       =   "FrmGuiaRemision.frx":00C8
         FormatStylesCount=   8
         FormatStyle(1)  =   "FrmGuiaRemision.frx":016C
         FormatStyle(2)  =   "FrmGuiaRemision.frx":0294
         FormatStyle(3)  =   "FrmGuiaRemision.frx":0344
         FormatStyle(4)  =   "FrmGuiaRemision.frx":03F8
         FormatStyle(5)  =   "FrmGuiaRemision.frx":04D0
         FormatStyle(6)  =   "FrmGuiaRemision.frx":0588
         FormatStyle(7)  =   "FrmGuiaRemision.frx":0668
         FormatStyle(8)  =   "FrmGuiaRemision.frx":06F8
         ImageCount      =   0
         PrinterProperties=   "FrmGuiaRemision.frx":080C
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
         TabIndex        =   48
         Top             =   240
         Width           =   615
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
         TabIndex        =   47
         Top             =   240
         Width           =   495
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
         TabIndex        =   46
         Top             =   240
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
         TabIndex        =   45
         Top             =   240
         Width           =   615
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
         TabIndex        =   44
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   0
      TabIndex        =   49
      Top             =   2040
      Width           =   15375
      Begin GridEX20.GridEX grxDatos 
         Height          =   5955
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   10504
         Version         =   "2.0"
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
         Column(1)       =   "FrmGuiaRemision.frx":09E4
         Column(2)       =   "FrmGuiaRemision.frx":0AAC
         FormatStylesCount=   9
         FormatStyle(1)  =   "FrmGuiaRemision.frx":0B50
         FormatStyle(2)  =   "FrmGuiaRemision.frx":0C78
         FormatStyle(3)  =   "FrmGuiaRemision.frx":0D28
         FormatStyle(4)  =   "FrmGuiaRemision.frx":0DDC
         FormatStyle(5)  =   "FrmGuiaRemision.frx":0EB4
         FormatStyle(6)  =   "FrmGuiaRemision.frx":0F6C
         FormatStyle(7)  =   "FrmGuiaRemision.frx":104C
         FormatStyle(8)  =   "FrmGuiaRemision.frx":10DC
         FormatStyle(9)  =   "FrmGuiaRemision.frx":1214
         ImageCount      =   0
         PrinterProperties=   "FrmGuiaRemision.frx":1328
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
      TabIndex        =   33
      Text            =   "G U I A   D E   R E M I S I O N"
      Top             =   0
      Width           =   15495
   End
   Begin VB.Frame frMain 
      Height          =   1080
      Left            =   0
      TabIndex        =   10
      Top             =   960
      Width           =   15375
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   14520
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   750
         Width           =   735
      End
      Begin VB.CommandButton cmdGenPartida 
         BackColor       =   &H00C0C0C0&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   14520
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   11280
         MaxLength       =   11
         TabIndex        =   65
         Top             =   750
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   12120
         MaxLength       =   11
         TabIndex        =   63
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   11280
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   62
         Top             =   480
         Width           =   825
      End
      Begin VB.TextBox txtNum_Ruc 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   6480
         MaxLength       =   11
         TabIndex        =   23
         Top             =   420
         Width           =   3495
      End
      Begin VB.TextBox txtSer_Docum 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   4530
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   20
         Top             =   120
         Width           =   1080
      End
      Begin VB.TextBox txtCod_TipDoc 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1005
         MaxLength       =   4
         TabIndex        =   19
         Top             =   120
         Width           =   465
      End
      Begin VB.TextBox txtDes_TipDoc 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1485
         TabIndex        =   18
         Top             =   120
         Width           =   2625
      End
      Begin VB.TextBox txtNum_Docum 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   5610
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   17
         Top             =   120
         Width           =   2020
      End
      Begin VB.TextBox txtDes_TipAne 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1485
         TabIndex        =   16
         Top             =   420
         Width           =   4425
      End
      Begin VB.TextBox txtCod_Moneda 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   8085
         MaxLength       =   4
         TabIndex        =   15
         Top             =   120
         Width           =   600
      End
      Begin VB.TextBox txtDes_Moneda 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   8685
         TabIndex        =   14
         Top             =   120
         Width           =   1650
      End
      Begin VB.TextBox txtLug_Entrega 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1005
         TabIndex        =   13
         Top             =   705
         Width           =   8970
      End
      Begin VB.TextBox txtCod_TipAne 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1005
         MaxLength       =   4
         TabIndex        =   12
         Text            =   "C"
         Top             =   420
         Width           =   465
      End
      Begin VB.Frame frReferencia 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Top             =   5040
         Visible         =   0   'False
         Width           =   7815
      End
      Begin MSComCtl2.DTPicker dtpFec_Emision 
         Height          =   285
         Left            =   11280
         TabIndex        =   21
         Top             =   165
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
         Format          =   71368705
         CurrentDate     =   38182
      End
      Begin MSComCtl2.DTPicker dtpFec_Registro 
         Height          =   285
         Left            =   13680
         TabIndex        =   22
         Top             =   165
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   71368705
         CurrentDate     =   38182
      End
      Begin VB.Label Label10 
         Caption         =   "PLACA"
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
         Left            =   10699
         TabIndex        =   64
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "TRANSPORTISTA"
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
         Left            =   9960
         TabIndex        =   61
         Top             =   525
         Width           =   1335
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
         TabIndex        =   32
         Top             =   120
         Width           =   285
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
         Left            =   45
         TabIndex        =   31
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Numero :"
         Height          =   255
         Left            =   9390
         TabIndex        =   30
         Top             =   375
         Width           =   15
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
         TabIndex        =   29
         Top             =   480
         Width           =   630
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
         TabIndex        =   28
         Top             =   420
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "TRASLADO"
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
         Left            =   12720
         TabIndex        =   27
         Top             =   165
         Width           =   855
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
         Left            =   10560
         TabIndex        =   26
         Top             =   240
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
         Left            =   7680
         TabIndex        =   25
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "DIRECCION"
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
         Left            =   120
         TabIndex        =   24
         Top             =   800
         Width           =   855
      End
   End
   Begin VB.Frame fraUbicacion 
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   15375
      Begin VB.OptionButton Option1 
         Caption         =   "NUEVO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   9600
         TabIndex        =   57
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "BUSCAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   10920
         TabIndex        =   56
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtSerieGuiaExis 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   12690
         MaxLength       =   3
         TabIndex        =   55
         Top             =   240
         Width           =   840
      End
      Begin VB.TextBox txtNumeroGuiaExis 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   13530
         MaxLength       =   8
         TabIndex        =   54
         Top             =   240
         Width           =   1785
      End
      Begin VB.TextBox txtDes_Fabrica 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   4365
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "SALIDA ROLLO VENTAS CLIENTES"
         Top             =   240
         Width           =   3330
      End
      Begin VB.TextBox txtCod_Fabrica 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   3885
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "SVR"
         Top             =   240
         Width           =   465
      End
      Begin VB.TextBox txtDes_Almacen 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtCod_Almacen 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   480
         MaxLength       =   4
         TabIndex        =   4
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label1 
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
         Left            =   12360
         TabIndex        =   58
         Top             =   240
         Width           =   285
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "TRANSACCION"
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
         Left            =   2640
         TabIndex        =   9
         Top             =   255
         Width           =   1110
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
         Left            =   45
         TabIndex        =   8
         Top             =   255
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdBusquedaProductos 
      Caption         =   "AYUDA"
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
      Left            =   4320
      TabIndex        =   2
      Top             =   8520
      Width           =   1335
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
      TabIndex        =   1
      Top             =   8520
      Width           =   3375
   End
   Begin VB.CheckBox chkImpresionDirecta 
      Caption         =   "IMPRESION DIRECTA"
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
      Left            =   9840
      TabIndex        =   0
      Top             =   8520
      Width           =   1935
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   5760
      TabIndex        =   51
      Top             =   8400
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmGuiaRemision.frx":1500
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Label Label2 
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
      Left            =   12600
      TabIndex        =   60
      Top             =   8640
      Width           =   615
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   15000
      Top             =   8400
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
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
      Height          =   255
      Left            =   120
      TabIndex        =   52
      Top             =   8520
      Width           =   615
   End
End
Attribute VB_Name = "FrmGuiaRemision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CODIGO As String, Descripcion As String, StrOption As String, strNum_Corre As String, strCod_Anxo As String
Public rsFactura As New ADODB.Recordset
Dim strSQL As String
Dim bClickColSelec As Boolean
Dim errorx As String
Public rstAux As ADODB.Recordset
Dim sTit As String
Public flg_Tiene_guias_asignadas As String
Public fila_seleccionada As Double

Public vnum_ruc As String
Public vdes_tipanex As String
Public vcod_tipanex As String
Public vCod_Cliente_Tex As String
Public vlug_entrega As String
Public indice As Integer

'.txtNum_Ruc.Text = DGridLista.Value(DGridLista.Columns(1).Index)
'.txtDes_TipAne.Text = Trim(DGridLista.Value(DGridLista.Columns(2).Index))
'.txtNum_Ruc.Tag = Trim(DGridLista.Value(DGridLista.Columns(4).Index))
'.txtDes_TipAne.Tag = Trim(DGridLista.Value(DGridLista.Columns(5).Index))
'.txtLug_Entrega.Text = Trim(DGridLista.Value(DGridLista.Columns(6).Index))
        
Private Declare Function GetSystemMenu Lib "user32" _
    (ByVal hwnd As Long, _
     ByVal bRevert As Long) As Long

Private Declare Function RemoveMenu Lib "user32" _
    (ByVal hMenu As Long, _
     ByVal nPosition As Long, _
     ByVal wFlags As Long) As Long
     
Private Const MF_BYPOSITION = &H400&

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
    strSQL = "Select " & strCampo1 & " AS Cod," & strCampo2 & " as Descripcion from " & strTabla

    txtCod = Trim(txtCod)
    txtDes = Trim(txtDes)
    Select Case Opcion
    Case 1: strSQL = strSQL & strCampo1 & " like '%" & txtCod & "%'"
    Case 2: strSQL = strSQL & strCampo2 & " like '%" & txtDes & "%'"
    End Select
    txtCod = ""
    txtDes = ""
    fila_seleccionada = 0
    
    With frmBusqGeneral
        Set .oParent = Me
        .SQuery = strSQL
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

Private Sub cmdAnularGuia_Click()
 If grxDatos.RowCount <= 0 Then Exit Sub
 
 If MsgBox("¡¡¡Esta apunto de Anular en el almacen el documento de salida!!!:" & Chr(13) & Chr(10) & ":::::> " & Trim(txtDes_TipDoc.Text) & " " & txtSerieGuiaExis & "-" & txtNumeroGuiaExis & Chr(13) & Chr(10) & "¿Son los datos correctos?", vbYesNo, "CONFIRMAR") = vbYes Then
    Call AnulaGuiacliente
    Call MuestraDetalleGuiaExiste
    habilitaframe (indice)
End If
 
 
End Sub
Private Sub AnulaGuiacliente()
On Error GoTo fin

Dim rsset As New ADODB.Recordset
Set rsset = grxDatos.ADORecordset

ExecuteCommandSQL cConnect, "VENTAS_MAN_ANULA_GUIA_REMISION'" & rsset!Cod_almacen & "','" & rsset!num_movstk & "','" & vusu & "'"
           
Exit Sub
fin:
'On Error Resume Next
rsset.Close
Set rsset = Nothing
MsgBox err.Description & ",No se puede Continuar", vbExclamation + vbOKOnly, _
"Anula Guia "
    
    
End Sub



Private Sub cmdBusquedaProductos_Click()
    FraProductos.Visible = True
    limpiarCajasBusqueda
End Sub

Private Sub limpiarCajasBusqueda()
    txtBus_Codigo_RolloTinto.Text = ""
    txtBus_Cod_ordtra.Text = ""
    txtBus_Des_Color.Text = ""
    txtDescripcion_Producto.Text = ""
End Sub

Private Sub cmdGenPartida_Click()
frmChoferes.Show 1
End Sub

Private Sub Command1_Click()
frmTransporte.Show 1
End Sub

Private Sub Form_Load()

    Call DisableCloseButton(Me)
    flg_Tiene_guias_asignadas = "N"
    FraProductos.Visible = False
    dtpFec_Emision.Value = Date
    dtpFec_Registro.Value = Date
    Call buscaDetalle_factura
    Call obtieneDatosIniciales
    txtCod_Almacen.Text = "40"
    txtDes_Almacen.Text = "Almacen Telas Terminadas"
    indice = 0
    Call habilitaframe(0)
    
'TxtTipo_Cambio.Text = DevuelveCampo("select isnull(Tipo_Venta,0) from cn_tipocambio where fecha = '" & dtpFec_Emision & "'", cConnect)
'If Not IsNumeric(TxtTipo_Cambio.Text) Then
'TxtTipo_Cambio.Text = 0
'End If
'If CDbl(TxtTipo_Cambio.Text) <= 0 Then
'  Call MsgBox("Ingrese el Tipo Cambio Para la fecha", vbCritical, "Importante")
'  'Unload Me
'End If
    
End Sub
Private Sub obtieneDatosIniciales()
Dim strSQL As String
Dim pc As String
Dim auxset As ADODB.Recordset
pc = ComputerName
'STRSQL = "CN_MUESTRA_CAJAS_VENDEDOR_ACCESO '" & pc & "'"
' Set auxset = Nothing
' Set auxset = CargarRecordSetDesconectado(STRSQL, cConnect)
' If auxset.RecordCount > 0 Then
'    txtCod_Fabrica.Text = auxset("cod_Fabrica")
'    txtDes_Fabrica.Text = auxset("nom_fabrica")
'    txtCod_Tienda.Text = auxset("cod_tienda")
'    txtDes_Tienda.Text = auxset("des_tienda")
'    txtCod_Caja.Text = auxset("cod_caja")
'    txtDes_Caja.Text = auxset("des_caja")
'    txtCod_Vendedor.Text = auxset("cod_vendedor")
'    txtDes_Vendedor.Text = auxset("des_vendedor")
'    txtCod_Almacen.Text = auxset("cod_almacen")
'    txtDes_Almacen.Text = auxset("nom_almacen")
'Else
'    Call MsgBox("La PC no Tiene una Caja Asignada", vbExclamation, "Importante")
'
'End If

End Sub

Private Sub Option1_Click(Index As Integer)
indice = Index
Call habilitaframe(Index)
End Sub
Private Sub habilitaframe(Opcion As Integer)

txtCod_TipDoc.Text = ""
txtDes_TipDoc.Text = ""
txtSer_Docum.Text = ""
txtNum_Docum.Text = ""
txtCod_Moneda.Text = ""
txtDes_Moneda.Text = ""
txtDes_TipAne.Text = ""
txtNum_Ruc.Text = ""
txtLug_Entrega.Text = ""
txtSerieGuiaExis.Text = ""
txtNumeroGuiaExis.Text = ""

Set grxDatos.ADORecordset = Nothing
Set GrxProductos.ADORecordset = Nothing

If Opcion = 0 Then
    
    frMain.Enabled = True
    cmdBusquedaProductos.Enabled = True
    txtCodigo_Producto.Enabled = True
    txtSerieGuiaExis.Enabled = False
    txtNumeroGuiaExis.Enabled = False
    grxDatos.Enabled = True
    grxDatos.AllowEdit = True
    cmdAnularGuia.Enabled = False
End If

If Opcion = 1 Then
    
    frMain.Enabled = False
    cmdBusquedaProductos.Enabled = False
    txtCodigo_Producto.Enabled = False
    txtSerieGuiaExis.Enabled = True
    txtNumeroGuiaExis.Enabled = True
    grxDatos.AllowEdit = False
    cmdAnularGuia.Enabled = True
End If

Call obtieneDatosIniciales
Call estadoInicialVentana
Call buscaDetalle_factura

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
           txtDes_Almacen.SetFocus
        End If
    End If
End Sub
Public Sub buscaAlmacen(sOpcion As String)
On Error GoTo fin
Dim iCol As Long
Dim rstAux As New ADODB.Recordset
Dim Opcion As String
  strSQL = "CN_MUESTRA_ALMACEN  '" & sOpcion & "','" & txtCod_Almacen.Text & "','" & txtDes_Almacen.Text & "','" & vusu & "'"
  With frmBusqGeneral
        Set .oParent = Me
        .SQuery = strSQL
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

Private Sub txtCod_TipDoc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call buscaDocumentos(1)
    
    If Trim(txtDes_TipDoc.Text) <> "" Then
      txtCod_Moneda.SetFocus
    Else
      txtCod_TipDoc.SetFocus
    End If
    
  End If
  
End Sub

Private Sub txtDes_TipDoc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call buscaDocumentos(2)
    
    If Trim(txtDes_TipDoc.Text) <> "" Then
        txtCod_Moneda.SetFocus
    Else
        txtDes_Moneda.SetFocus
    End If
    
  End If
End Sub
Public Sub buscaDocumentos(sOpcion As String)
On Error GoTo fin
Dim iCol As Long
Dim rstAux As New ADODB.Recordset
Dim Opcion As String
  fila_seleccionada = 0
  strSQL = "CN_MUESTRA_VENTAS_DOCUMENTOS_GUIAS  '" & sOpcion & "','" & Trim(txtCod_TipDoc.Text) & "','" & Trim(txtDes_TipDoc.Text) & "'"
  With frmBusqGeneral
        Set .oParent = Me
        .SQuery = strSQL
        .Cargar_Datos
        CODIGO = ".."
        Set rstAux = .gexList.ADORecordset
        
        .gexList.Columns("Cod_TipDoc").Caption = "Codigo"
        .gexList.Columns("Cod_TipDoc").Width = 1000
        .gexList.Columns("DES_TIPDOC").Caption = "Almacen"
        .gexList.Columns("DES_TIPDOC").Width = 4000
        
        'If rstAux.RecordCount > 1 Then
        .Show vbModal
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

Private Sub txtCod_Moneda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  Call Busca_Opcion("Cod_Moneda", "Nom_Moneda", "TG_Moneda where ", txtCod_Moneda, txtDes_Moneda, 1)
  
  If Trim(txtDes_Moneda.Text) <> "" Then
     txtDes_TipAne.SetFocus
  Else
     txtCod_Moneda.SetFocus
  End If
  
  End If
End Sub
Private Sub txtDes_Moneda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_Moneda", "Nom_Moneda", "TG_Moneda where ", txtCod_Moneda, txtDes_Moneda, 2)
    If Trim(txtDes_Moneda.Text) <> "" Then
       txtDes_TipAne.SetFocus
    Else
       txtDes_Moneda.SetFocus
    End If
  
  End If
End Sub
Private Sub txtDes_TipAne_KeyPress(KeyAscii As Integer)
'  If KeyAscii = 13 Then
'        Load FrmBusqClientesGuia
'        Set FrmBusqClientesGuia.oParent = Me
'        FrmBusqClientesGuia.txtDescripcion_Cliente.Text = txtDes_TipAne.Text
'        FrmBusqClientesGuia.txtRuc_Cliente.Text = "" 'txtNum_Ruc.Text
'        FrmBusqClientesGuia.txtTip_Anex.Text = "C"
'
'        Call FrmBusqClientesGuia.Busca_Opcion_AnexoContable("2", "C", txtNum_Ruc.Text, txtDes_TipAne.Text)
'        FrmBusqClientesGuia.Show 1
'        'FrmBusqClientes.txtDescripcion_Cliente.SetFocus
'        Set FrmBusqClientesGuia = Nothing
'  End If

    If KeyAscii = 13 Then
        BuscaAnexosContable (2)
        'If txtDes_Almacen.Text <> "" Then
           'txtCod_TipDoc.SetFocus
        'Else
           'txtCod_Almacen.SetFocus
        'End If
    End If
    
End Sub
Private Sub txtNum_Ruc_KeyPress(KeyAscii As Integer)
'
'  If KeyAscii = 13 Then
'        Load FrmBusqClientesGuia
'        Set FrmBusqClientesGuia.oParent = Me
'        FrmBusqClientesGuia.txtDescripcion_Cliente.Text = "" 'txtDes_TipAne.Text
'        FrmBusqClientesGuia.txtRuc_Cliente.Text = txtNum_Ruc.Text
'        FrmBusqClientesGuia.txtTip_Anex.Text = "C"
'        txtDes_TipAne.Text = ""
'
'        Call FrmBusqClientesGuia.Busca_Opcion_AnexoContable("1", "C", txtNum_Ruc.Text, txtDes_TipAne.Text)
'        FrmBusqClientesGuia.Show 1
'        Set FrmBusqClientesGuia = Nothing
'  End If
  
If KeyAscii = 13 Then
    BuscaAnexosContable (1)
End If
  
End Sub
Private Sub grxDatos_Click()

    Dim ColIndex As Long
    Dim oRowData As JSRowData
    Dim SGRUPO As String
    Dim iRow As Long
    Dim i As Long
    Dim sCaptionGroup As String
        If grxDatos.RowCount > 0 Then
        ColIndex = grxDatos.Col
       
       If grxDatos.AllowEdit = True Then
        If UCase(grxDatos.Columns(ColIndex).Key) = "ELI" Then
            bClickColSelec = True
            SendKeys "{ENTER}"
        End If
       End If
    End If
End Sub
Private Sub grxDatos_AfterColEdit(ByVal ColIndex As Integer)
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
   End If
   'Call Total_documento
  Case Is = grxDatos.Columns("CANT").Index
  
     If IsNumeric(grxDatos.Value(grxDatos.Columns("CANT").Index)) = False Or grxDatos.Value(grxDatos.Columns("CANT").Index) = "" Then
         grxDatos.Value(grxDatos.Columns("CANT").Index) = 0
     End If
    grxDatos.Value(grxDatos.Columns("TOTAL").Index) = grxDatos.Value(grxDatos.Columns("PRECIO").Index) * grxDatos.Value(grxDatos.Columns("CANT").Index)
    'Call Total_documento
    Call ConfiguraGrilla_Detalle
  End Select
Exit Sub

Resume
Error_Handler:
errores err.Number
End Sub

Public Sub Busca_Opcion_AnexoContable(sTipo As String, txttipo As String, ruc As String, txtDes As String)
On Error GoTo fin

Dim rstAux As Object, strSQL As String
Set rstAux = CreateObject("ADODB.Recordset")
    strSQL = "CN_MUESTRA_ANEXOS_CLIENTES '" & sTipo & "','" & txttipo & "','" & ruc & "','" & txtDes & "'"
    
    
    With FrmBusqClientesGuia
        .SQuery = strSQL
        .Cargar_Datos
        
        CODIGO = ""
        .DGridLista.Columns("Cod").Visible = False
        .DGridLista.Columns("Tipo").Width = 800
        .DGridLista.Columns("Nombre").Width = 4075
        .DGridLista.Columns("RUC").Width = 1200
        .DGridLista.Columns("LUG_ENTREGA").Width = 800
        Set rstAux = .DGridLista.ADORecordset
    
    End With
    
Exit Sub
fin:
On Error Resume Next
    Unload FrmBusqClientesGuia
    Set FrmBusqClientesGuia = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento "
End Sub

Public Sub BuscaAnexosContable(sOpcion As String)
On Error GoTo fin
Dim iCol As Long
Dim rstAux As New ADODB.Recordset
Dim Opcion As String
  'STRSQL = "CN_MUESTRA_VENTAS_CAJAS_ALMACEN  '" & sOpcion & "','" & Trim(txtCod_Tienda.Text) & "','" & Trim(txtCod_Caja.Text) & "','" & Trim(txtCod_Almacen.Text) & "','" & Trim(txtDes_Almacen.Text) & "'"
  strSQL = "CN_MUESTRA_ANEXOS_CLIENTES '" & sOpcion & "','C','" & Trim(txtNum_Ruc.Text) & "','" & Trim(txtDes_TipAne.Text) & "'"
  fila_seleccionada = 0
  With FrmBusqClientesGuia
        Set .oParent = Me
        .SQuery = strSQL
        .Cargar_Datos
        CODIGO = ".."
        Set rstAux = .DGridLista.ADORecordset
        
         CODIGO = ""
        .DGridLista.Columns("Cod").Visible = False
        .DGridLista.Columns("Tipo").Width = 800
        .DGridLista.Columns("Nombre").Width = 4075
        .DGridLista.Columns("RUC").Width = 1200
        .DGridLista.Columns("LUG_ENTREGA").Width = 800
        Set rstAux = .DGridLista.ADORecordset
    
        '.DGridLista.SetFocus
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        'If codigo <> "" And rstAux.RecordCount > 0 Then
        '    rstAux.AbsolutePosition = fila_seleccionada
        '    txtNum_Ruc.Text = Trim(rstAux!ruc)
        '    txtDes_TipAne.Text = Trim(rstAux!Nombre)
        '    txtNum_Ruc.Tag = Trim(rstAux!Cod)
        '    txtDes_TipAne.Tag = Trim(rstAux!cod_cliente_tex)
        '    txtLug_Entrega.Text = Trim(rstAux!LUG_ENTREGA)
        '
        'End If
    
    txtNum_Ruc.Text = vnum_ruc
    txtDes_TipAne.Text = vdes_tipanex
    txtNum_Ruc.Tag = vcod_tipanex
    txtDes_TipAne.Tag = vCod_Cliente_Tex
    txtLug_Entrega.Text = vlug_entrega
    
    End With
    Unload FrmBusqClientesGuia
    Set FrmBusqClientesGuia = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
fin:
On Error Resume Next
    Unload FrmBusqClientesGuia
    Set FrmBusqClientesGuia = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ",No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de almcen(" & Opcion & ")"
End Sub

Private Sub txtBus_Codigo_RolloTinto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call buscarProductos(1)
End If
End Sub

Private Sub txtBus_Cod_ordtra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call buscarProductos(4)
End If

End Sub

Private Sub txtDescripcion_Producto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call buscarProductos(3)
    End If
    
End Sub
Private Sub txtBus_Des_Color_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call buscarProductos(5)
End If

End Sub

''''*************************************************************BUSQUEDA DE PRODUCTOS *********************************
Private Sub buscarProductos(Opcion As String)

    Dim strSQL As String
    Dim sCodCentroCosto As String
    Dim nrofilas As Integer
    Dim rsproductos As New ADODB.Recordset
    
    On Error GoTo fin
   
    strSQL = "TX_MUESTRA_ROLLOS_VENTA '" & Opcion & "','" & Trim(txtCod_Almacen.Text) & "','" & Trim(txtBus_Codigo_RolloTinto.Text) & "','" & Trim(txtBus_Cod_ordtra.Text) & "','" & Trim(txtDescripcion_Producto.Text) & "','" & Trim(txtBus_Des_Color.Text) & "'"
    
    Set GrxProductos.ADORecordset = Nothing
    Set GrxProductos.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    
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
    
    
    nrofilas = GrxProductos.RowCount
    If nrofilas > 0 Then
            nrofilas = 15
    Else
       FraProductos.Visible = True
    End If
        Call ConfiguraGrilla_productos
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
        .Columns("CANT").Width = 1000
        .Columns("CANT").Visible = True
        
        
        .Columns("total").Width = 1000
        .Columns("total").Visible = False
        
        .Columns("PRECIO").Width = 1000
        .Columns("PRECIO").Visible = False
        
        .Columns("sel").Width = 800
        .Columns("Sel").Visible = False
        
        If .RowCount > 0 Then
            .Row = -1
            .Col = .Columns.Count - 1
        End If
        '.SetFocus
    End With
    
    Set fmtCon = GrxProductos.FmtConditions.Add(GrxProductos.Columns("CANT").Index, jgexGreaterThan, 0)
    fmtCon.FormatStyle.BackColor = &H80FFFF   ' &HFFFF00

    Dim oGroup01 As GridEX20.JSGroup
    Dim oGroup02 As GridEX20.JSGroup
    Dim oGroup03 As GridEX20.JSGroup
    
    Dim valorcant    As JSColumn
    Dim valorStock   As JSColumn
    
      With GrxProductos
            
        Set oGroup01 = .Groups.Add(.Columns("OT").Index, jgexSortAscending)
        '.DefaultGroupMode = jgexDGMCollapsed
        .DefaultGroupMode = jgexDGMExpanded
        .BackColorRowGroup = RGB(239, 235, 222)
           
           .GroupFooterStyle = jgexTotalsGroupFooter
           Set valorcant = .Columns("CANT")
           Set valorStock = .Columns("STOCK")

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

           If .RowCount > 0 Then
                .Row = -1
                .Col = .Columns.Count - 1
           End If
        End With
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

    Dim strSQL As String
    Dim sCodCentroCosto As String
    Dim nrofilas As Integer
    
    On Error GoTo fin
   
    strSQL = "EXEC CN_MUESTRA_TELAS_DETALLE_FACTURA 'xx','" & Trim(txtCodigo_Producto.Text) & "','" & Trim(txtDescripcion_Producto.Text) & "'"
    
    Set grxDatos.ADORecordset = Nothing
    Set grxDatos.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    
    Call ConfiguraGrilla_Detalle
    Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
''''*******************************************CONFIGURA DETALLE DE DOCUMENTO*********************************
Private Sub ConfiguraGrilla_Detalle()
    Dim C As Integer
    On Error GoTo fin
    
    Call ConfiguraGrilla_DetalleSinGrupos
'    If grxDatos.RowCount > 0 Then
'        Dim rstemp As New ADODB.Recordset
'        Set rstemp = grxDatos.ADORecordset
'        grxDatos.ADORecordset = rstemp
'        Call Total_documento
'    End If
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
        
        .Columns("CANT").Width = 1000
        .Columns("CANT").Visible = True
        .Columns("CANT").Caption = "CANT"
        
      '  .Columns("PRECIO").Width = 1000
      '  .Columns("PRECIO").Visible = False
        
     '   .Columns("TOTAL").Width = 1000
     '   .Columns("TOTAL").Visible = False
        
        .Columns("ELI").Width = 250
        .Columns("ELI").Visible = True
        .Columns("ELI").Caption = ""

        .Columns("DEL").Width = 250
        .Columns("DEL").Visible = True
        .Columns("DEL").Caption = "X"
        .Columns("DEL").TextAlignment = jgexAlignCenter
        
        SetColorDetalle
        'DEL
        'With GrxProductos.Columns("OK")
        '    .TextAlignment = jgexAlignLeft
        '    .EditType = jgexEditDropDown
        '   Set .DropDownControl = GrxProductos
        'End With
        
    Dim oGroup01 As GridEX20.JSGroup
    Dim oGroup02 As GridEX20.JSGroup
    Dim oGroup03 As GridEX20.JSGroup
    
    Dim valorcant    As JSColumn
    Dim valorStock   As JSColumn
    Dim ItemTotal    As JSColumn
    Dim rsdatos  As New ADODB.Recordset
    
      With grxDatos
            
        Set oGroup01 = .Groups.Add(.Columns("OT").Index, jgexSortAscending)
        .DefaultGroupMode = jgexDGMExpanded
        .BackColorRowGroup = RGB(239, 235, 222)
           
           .GroupFooterStyle = jgexTotalsGroupFooter
           Set valorcant = .Columns("CANT")
           Set valorStock = .Columns("STOCK")
           
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

        End With
        
        If .RowCount > 0 Then
            .Row = -1
            .Col = .Columns.Count - 1
        End If
    End With
    
    Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
Private Sub SetColorDetalle()
        'grxDatos.Columns("ROLLOS").CellStyle = "estilo_cantidad"
        grxDatos.Columns("CANT").CellStyle = "estilo_cantidad"
        'grxDatos.Columns("PRECIO").CellStyle = "estilo_cantidad"
        grxDatos.Columns("ELI").CellStyle = "estilo_eliminar"
        grxDatos.Columns("DEL").CellStyle = "estilo_eliminar"
End Sub

Private Sub ConfiguraGrilla_DetalleSinGrupos()
    Dim C As Integer
    On Error GoTo fin
    
    Dim rsdatos As New ADODB.Recordset
    Set rsdatos = grxDatos.ADORecordset
    Set grxDatos.ADORecordset = rsdatos
    
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
        
        .Columns("CANT").Width = 1000
        .Columns("CANT").Visible = True
        .Columns("CANT").Caption = "CANT"
        
      '  .Columns("PRECIO").Width = 1000
      '  .Columns("PRECIO").Visible = False
        
     '   .Columns("TOTAL").Width = 1000
     '   .Columns("TOTAL").Visible = False
        
        .Columns("ELI").Width = 250
        .Columns("ELI").Visible = True
        .Columns("ELI").Caption = ""

        .Columns("DEL").Width = 250
        .Columns("DEL").Visible = True
        .Columns("DEL").Caption = "X"
        .Columns("DEL").TextAlignment = jgexAlignCenter
        End With
        SetColorDetalle
        Call Total_documento
    Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
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


Private Sub cmdCerrarBusProductos_Click()
FraProductos.Visible = False
Set GrxProductos.ADORecordset = Nothing
End Sub

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
Private Sub txtCodigo_Producto_Change()
  If Len(Trim(txtCodigo_Producto.Text)) = 9 Then
    Call AdicionaProductoDirecto(1)
    txtCodigo_Producto.Text = ""
    SendKeys "{TAB}"
  End If
End Sub
Private Sub AdicionaProductoDirecto(Opcion As String)

    Dim strSQL As String
    Dim sCodCentroCosto As String
    Dim rsetAux As ADODB.Recordset
    Dim rsetbusqueda As ADODB.Recordset
    Dim nrofilas As Integer
    
    On Error GoTo fin
    
    If validaDatosIniciales = False Then
        Exit Sub
    End If
    
    strSQL = "TX_MUESTRA_ROLLOS_VENTA '" & Opcion & "','" & Trim(txtCod_Almacen.Text) & "','" & Trim(txtCodigo_Producto.Text) & "','" & Trim(txtBus_Cod_ordtra.Text) & "','" & Trim(txtDescripcion_Producto.Text) & "','" & Trim(txtBus_Des_Color.Text) & "'"
    Set rsetbusqueda = Nothing
    Set rsetbusqueda = CargarRecordSetDesconectado(strSQL, cConnect)
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
    
    'Call Total_documento
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
    Call ConfiguraGrilla_Detalle
    
End Sub
''''*************************************************************SUMA TOTALES DE DOCUMENTO*********************************
Private Sub Total_documento()
On Error GoTo ErrCal
    Dim Total As Double
    Dim merma As Double
    Dim mermavar As Variant
    Dim i As Integer
    Total = 0
    'grxDatos.Update
    i = 1
    
    
    If grxDatos.RowCount >= 0 Then
    
            If grxDatos.RowCount > 0 Then
                'grxDatos.Update
            End If
            grxDatos.Refresh
            grxDatos.MoveFirst
            
            Do While i <= grxDatos.RowCount
                Total = Total + grxDatos.Value(grxDatos.Columns("CANT").Index)
                
                If i < grxDatos.RowCount Then
                    grxDatos.MoveNext
                End If
                i = i + 1
            Loop
            txtTotalKilos.Text = Total
            'txt_subtotal.Text = Format(Total / 1.18, "####.00")
            'txt_igv.Text = Format(Total - (Total / 1.18), "####.00")
            
            
     Else
            txtTotalKilos.Text = Total
            'txt_subtotal.Text = Format(Total / 1.18, "####.00")
            'txt_igv.Text = Format(Total - (Total / 1.18), "####.00")

     End If
     Exit Sub
ErrCal:
    MsgBox err.Description, vbCritical + vbOKOnly, "Cargar Calidades"

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
'    Case Is = grxDatos.Columns("PRECIO").Index
'      Cancel = False
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
    
        If Trim(txtCod_Almacen.Text) = "" Then
           Call MsgBox("El Codigo del Almacen no es valido", vbCritical + vbOKOnly, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
         
        If Trim(txtCod_Fabrica.Text) = "" Then
           Call MsgBox("Ingrese Una Empresa valida", vbInformation + vbOKOnly, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
        
        If Trim(txtCod_Almacen.Text) = "" Then
           Call MsgBox("El Codigo del Almacen no es valido", vbInformation + vbOKOnly, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
        
        If DevuelveCampo(" select count(*) from tx_movistk where  ser_guia ='" & Trim(txtSer_Docum.Text) & "' and numero_guia='" & Trim(txtNum_Docum.Text) & "' ", cConnect) > 0 Then
             MsgBox " la guia yan existe en el almacen, sirvase revisar...", vbInformation + vbOKOnly, "Importante"
             validaDatosIniciales = False
             Exit Function
        End If
        
        If DevuelveCampo(" select count(*) from tx_guias_Remision  where  ser_guia ='" & Trim(txtSer_Docum.Text) & "' and numero_guia='" & Trim(txtNum_Docum.Text) & "' ", cConnect) > 0 Then
             MsgBox " La guia se encuentra Anulada, sirvase revisar...", vbInformation + vbOKOnly, "Importante"
             validaDatosIniciales = False
             Exit Function
        End If
        
        If Trim(txtDes_TipAne.Text) = "" Then
           Call MsgBox("sivarse ingresar un cliente valido ", vbInformation + vbOKOnly, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
        
        
   'end if
End Function
Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo dprDepurar

    Select Case ActionName
    Case Is = "GRABAR"
      
      If grxDatos.RowCount <= 0 Then Exit Sub
      
       If indice = 0 Then
            If validaDatosIniciales = True Then
            
                  If MsgBox("¡¡¡Esta apunto de confirmar en el almacen el documento de salida!!!:" & Chr(13) & Chr(10) & ":::::> " & Trim(txtDes_TipDoc.Text) & " " & txtSer_Docum & "-" & txtNum_Docum & Chr(13) & Chr(10) & "¿Son los datos correctos?", vbInformation + vbOKCancel, "CONFIRMAR") = vbOK Then
                      If GuardaDetalleVentas = True Then
                          Call obtieneDatosIniciales
                          Call estadoInicialVentana
                          Call buscaDetalle_factura
                      End If
                  End If
            End If
       End If
       
       '''esta opcion es para cuando buscamos un guia existente
       If indice = 1 Then
        
        Dim rsrecord  As New ADODB.Recordset
         Set rsrecord = grxDatos.ADORecordset
         'Call Preliminar_Guia(sNum_MovStk, txtCod_Almacen.Text, txtCod_TipDoc.Text, txtSer_Docum.Text)
          Call Preliminar_Guia(rsrecord("num_movstk"), rsrecord("COD_ALMACEN"), "GR", txtSerieGuiaExis.Text)
       End If
    Case Is = "CANCELAR"
      
      If grxDatos.RowCount > 0 Then
      
             If indice = 0 Then
                If MsgBox("¡...Al cancelar esta operacion se eliminaran los datos registrados...! " & Chr(13) & Chr(10) & " ¿Esta Seguro de proseguir? ", vbYesNo, "CONFIRMAR") = vbYes Then
                  Unload Me
                End If
             Else
             
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
Private Sub estadoInicialVentana()
'''generar el sgte numero de documento
'''limpiar y txt, grilla
txtDes_TipAne.Text = ""
txtNum_Ruc.Text = ""
txtDes_TipAne.Tag = ""
txtNum_Ruc.Tag = ""
'001 001 01  GR  003             00000001
txtNum_Docum.Text = DevuelveCampo("SELECT COR_NUMACTU FROM CN_VENTAS_CAJAS_DOCUMENTOS WHERE COD_FABRICA='001' AND  COD_TIENDA='001' AND COD_CAJA='001' AND COD_TIPDOC='" & Trim(txtCod_TipDoc.Text) & "' AND COR_DOCSERIE ='" & txtSer_Docum.Text & "' ", cConnect)

Set grxDatos.ADORecordset = Nothing
Set GrxProductos.ADORecordset = Nothing

End Sub
'''''***********************************GUARDA EL DETALLE DE LA FACTURA DESDE LA BUSQUEDA O CON LECTORA DE BARRAS, GENERA MOVIMIENTO DE ALMACEN ****************************
Private Function GuardaDetalleVentas() As Boolean
On Error GoTo ErrDetMov
Dim sErr As String, cntAux As New ADODB.Connection, sTit As String, _
    sNum_MovStk As String, strNum_Corre  As String
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
'    STRSQL = "VENTAS_UP_MAN_ROLLOS 'I','','" & txtCod_Fabrica.Text & "','" & Trim(txtCod_Tienda.Text) & "','" & Trim(txtCod_Caja.Text) & "','" & Trim(txtCod_Vendedor.Text) & "','" & Trim(txtCod_Almacen.Text) & "','" & txtCod_TipDoc & "','" & txtSer_Docum & "','" _
'            & txtNum_Docum & "','C','" & Trim(txtNum_Ruc.Tag) & "','" & txtCod_ConPag & "','" & txtCod_TipVenta.Text & "','" & Format(dtpFec_Emision.Value, "dd/mm/yyyy") & "','" _
'            & Format(dtpFec_Registro.Value, "dd/mm/yyyy") & "','" & txtCod_Moneda & "','" _
'            & vusu & "',''," _
'            & TxtTipo_Cambio.Text & ",'','','N','N','S'"
'
'    Set rstAux = cntAux.Execute(STRSQL, adExecuteNoRecords)
'    strNum_Corre = rstAux!Num_Corre
'    rstAux.Close

    '''CABECERA MOVIMIENTO
    strSQL = "EXEC TI_UP_MAN_TX_MOVISTK_TELA_TENIDA_CABECERA_ROLLOS 'I', '" & _
             Trim(txtCod_Almacen.Text) & "', '', '" & Format(dtpFec_Registro.Value, _
             "dd/mm/yyyy") & "', '', '' ,'SVD','', '', '" & txtDes_TipAne.Tag & _
             "', '', '', 'movimiento de venta directo', '" & vusu & "', '" & _
             0 & "', '" & 0 & "','',''"

    Set rstAux = cntAux.Execute(strSQL, adExecuteNoRecords)
    sNum_MovStk = rstAux!num_movstk
    rstAux.Close
    
    Set rstAux = grxDatos.ADORecordset
    With rstAux
        If .RecordCount > 0 Then .MoveFirst
        Do Until .EOF
    '''DETALLE MOVIMIENTO DE SALIDA DE ALMACEN
             strSQL = "EXEC TI_UP_MAN_TX_MOVISTK_TELA_TENIDA_PESADAS_ROLLO_VENTAS_DIRECTA 'I', '" & _
             Trim(txtCod_Almacen.Text) & "', '" & sNum_MovStk & "', '', '" & _
             !codigoRollo & "'," & !Stock & "," & !cant & ",0, " & _
             Trim(!rollos) & ",'" & vusu & "',0"
             cntAux.Execute strSQL, adExecuteNoRecords
    
'    '''DETALLE VENTAS falta strCod_Anxo
'            strSQL = "VENTAS_UP_MAN_DETALLE_ROLLO 'I','" & strNum_Corre & "','','D','" & Trim(!codigorollo) & "','" & _
'            Trim(!cod_tela) & "','','" & !und & "'," & !rollos & "," & !Stock & "," & !cant & "," _
'            & !precio & "," & !Total & ",0,'','',0,'" & Trim(txtCod_Almacen.Text) & "','" & !OT & "','" & vusu & "'"
'            cntAux.Execute strSQL, adExecuteNoRecords
            
        .MoveNext
        Loop
    End With
    
    '''Guarda Guia Remision en memoria
    strSQL = "TX_GUARDA_MOVIMIENTO_GUIA '" & Trim(txtCod_Almacen.Text) & "','" & sNum_MovStk & "','" & Trim(txtSer_Docum.Text) & "','" & Trim(txtNum_Docum.Text) & "','" & Format(dtpFec_Emision.Value, "DD/MM/YYYY") & "','" & Format(dtpFec_Registro.Value, "DD/MM/YYYY") & "'"
    cntAux.Execute strSQL, adExecuteNoRecords

    cntAux.CommitTrans
    cntAux.Close
    Set cntAux = Nothing
    
    '''IMPRIME DOCUMENTO
    Call Preliminar_Guia(sNum_MovStk, txtCod_Almacen.Text, txtCod_TipDoc.Text, txtSer_Docum.Text)
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
Private Sub Preliminar_Guia(num_movstk As String, Cod_almacen As String, cod_tipodoc As String, ser_docum As String)
On Error GoTo SALTO_ERROR
Dim sSQL As String, rs As New ADODB.Recordset

Dim aMess(4), i As Integer
 
If Imprimir_guia(Cod_almacen, num_movstk, cod_tipodoc, ser_docum) = False Then
   MsgBox "Problemas de Impresion con el Documento Nro " & txtNum_Docum.Text, vbInformation, "ERROR"
   'Buscar
   Exit Sub
End If
    
Exit Sub
SALTO_ERROR:
MsgBox err.Description, vbCritical, Me.Caption
    
End Sub
   
Public Function Imprimir_guia(Cod_almacen As String, num_movstk As String, strCod_Cod As String, Serie As String) As Boolean

Dim Rs2 As New ADODB.Recordset, rs3 As New ADODB.Recordset, strSQL As String, scnt As Integer
scnt = 0
With rsFactura
 
    Select Case strCod_Cod
    
    Case Is = "GR" 'llll
    strSQL = "TI_MUESTRA_DATOS_GUIA '" & Cod_almacen & "','" & num_movstk & "'"
    Set rsFactura = CargarRecordSetDesconectado(strSQL, cConnect)
        If rsFactura.RecordCount > 0 Then
            Call guia_sa("GR", Serie)
            scnt = 2
        Else
           Call MsgBox("La Guia no Tiene Detalle", vbInformation, "Mensaje")
           Imprimir_guia = False
           Exit Function
        End If
        
    Case Else
      MsgBox "No se ha Definido un Formato de Impresion para este tipo de documento", vbInformation, "ERROR"
       Imprimir_guia = False
      Exit Function
    End Select
    
End With

Imprimir_guia = True

End Function
Sub guia_sa(tipo As String, Serie As String)
On Error GoTo ErrorImpresion
Dim oo As Object, lvSql As String, lvRuta As String

    Set oo = CreateObject("excel.application")
    
    If tipo = "GR" Then
            oo.Workbooks.Open vRuta & "\GRemisionTelaTenidaRollo.XLT"
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
        
    If chkTodos.Value = Checked Then
        If rs("stock") > 0 Then
            rs("cant") = rs("stock")
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
'''******************************* MUESTRA DETALLE DE LA GUIA 0*******************************************
Private Sub MuestraDetalleGuiaExiste()
Dim RSAUX As ADODB.Recordset
Dim rslista As ADODB.Recordset
Dim i As Integer
On Error GoTo fin


strSQL = "TX_MUESTRA_CABECERA_GUIA '" & txtCod_Almacen.Text & "','" & Trim(txtSerieGuiaExis.Text) & "','" & Trim(txtNumeroGuiaExis.Text) & "'"
Set rslista = Nothing
Set rslista = CargarRecordSetDesconectado(strSQL, cConnect)

If rslista.RecordCount <= 0 Then Exit Sub

txtCod_TipDoc.Text = rslista("TIPODOCUMENTO")
txtDes_TipDoc.Text = rslista("TIPODOCUMENTO")
txtSer_Docum.Text = txtSerieGuiaExis.Text
txtNum_Docum.Text = txtNumeroGuiaExis.Text
txtCod_Moneda.Text = rslista("COD_MONEDA")
txtDes_Moneda.Text = rslista("DES_MONEDA")
txtDes_TipAne.Text = rslista("NOM_CLIENTE")
txtNum_Ruc.Text = rslista("NUM_RUC")
txtLug_Entrega.Text = rslista("LUG_ENTREGA")


'''' detalle de las guias
strSQL = "CN_MUESTRA_DETALLE_GUIA_REMISION '" & Trim(txtSerieGuiaExis.Text) & "','" & Trim(txtNumeroGuiaExis.Text) & "'"
Set rslista = Nothing
Set grxDatos.ADORecordset = Nothing

Set rslista = CargarRecordSetDesconectado(strSQL, cConnect)
If rslista.RecordCount <= 0 Then Exit Sub
Set grxDatos.ADORecordset = rslista

Call ConfiguraGrilla_Detalle

Exit Sub
Resume
fin:
On Error Resume Next
Set rslista = Nothing
MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
"Edicionar Producto "
    
End Sub
Private Sub txtNum_Docum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtNum_Docum.Text = Format(txtNum_Docum, "00000000")
    End If
End Sub

Private Sub txtNumeroGuiaExis_LostFocus()
 txtNumeroGuiaExis.Text = Format(txtNumeroGuiaExis, "00000000")
End Sub

Private Sub txtSer_Docum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtSer_Docum.Text = Format(txtSer_Docum, "000")
    End If
End Sub
Private Sub txtNumeroGuiaExis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtNumeroGuiaExis.Text = Format(txtNumeroGuiaExis, "00000000")
        Call MuestraDetalleGuiaExiste
    End If
End Sub
Private Sub txtSerieGuiaExis_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       txtSerieGuiaExis.Text = Format(txtSerieGuiaExis, "000")
        Call MuestraDetalleGuiaExiste
        txtNumeroGuiaExis.SetFocus
    End If
End Sub

Private Sub txtSerieGuiaExis_LostFocus()
   txtSerieGuiaExis.Text = Format(txtSerieGuiaExis, "000")
End Sub
