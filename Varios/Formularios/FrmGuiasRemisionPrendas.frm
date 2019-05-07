VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmGuiasRemisionPrendas 
   Caption         =   "Form1"
   ClientHeight    =   8865
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16770
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   16770
   StartUpPosition =   3  'Windows Default
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
      Height          =   495
      Left            =   8040
      TabIndex        =   69
      Top             =   8280
      Width           =   1335
   End
   Begin VB.Frame fraUbicacion 
      Height          =   615
      Left            =   0
      TabIndex        =   48
      Top             =   360
      Width           =   16575
      Begin VB.CommandButton cmdBuscarGuia 
         BackColor       =   &H00C0C0C0&
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
         Height          =   390
         Left            =   15360
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   180
         Width           =   855
      End
      Begin VB.TextBox txtCod_Almacen 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   480
         MaxLength       =   4
         TabIndex        =   56
         Top             =   240
         Width           =   465
      End
      Begin VB.TextBox txtDes_Almacen 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   960
         TabIndex        =   55
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtCod_Fabrica 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   3885
         Locked          =   -1  'True
         TabIndex        =   54
         Text            =   "SV6"
         Top             =   240
         Width           =   465
      End
      Begin VB.TextBox txtDes_Fabrica 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   4365
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "Salidas por Ventas"
         Top             =   240
         Width           =   3330
      End
      Begin VB.TextBox txtNumeroGuiaExis 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   13530
         MaxLength       =   8
         TabIndex        =   52
         Top             =   240
         Width           =   1785
      End
      Begin VB.TextBox txtSerieGuiaExis 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   12690
         MaxLength       =   3
         TabIndex        =   51
         Top             =   240
         Width           =   840
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
         TabIndex        =   50
         Top             =   240
         Width           =   1095
      End
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
         TabIndex        =   49
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
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
         TabIndex        =   59
         Top             =   255
         Width           =   375
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
         TabIndex        =   58
         Top             =   255
         Width           =   1110
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
         TabIndex        =   57
         Top             =   240
         Width           =   285
      End
   End
   Begin VB.Frame FraProductos 
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00000000&
      Height          =   5520
      Left            =   960
      TabIndex        =   0
      Top             =   2040
      Width           =   15015
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
         Left            =   14160
         TabIndex        =   10
         Top             =   240
         Width           =   735
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
         Left            =   13920
         TabIndex        =   9
         Top             =   5115
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
         Left            =   12960
         TabIndex        =   8
         Top             =   5115
         Width           =   975
      End
      Begin VB.CommandButton cmdBusLimpiarCaja 
         BackColor       =   &H0080C0FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   50
         TabIndex        =   7
         Top             =   120
         Width           =   420
      End
      Begin VB.TextBox txtCodigo_Barra_Bus 
         BackColor       =   &H00C0FFFF&
         Height          =   350
         Left            =   10320
         TabIndex        =   6
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtDes_Estcli_Bus 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   350
         Left            =   3600
         TabIndex        =   5
         Top             =   120
         Width           =   3255
      End
      Begin VB.TextBox txtDes_Present_Bus 
         BackColor       =   &H00C0FFFF&
         Height          =   350
         Left            =   7440
         TabIndex        =   4
         Top             =   120
         Width           =   2295
      End
      Begin VB.TextBox TxtCod_Estcli_Bus 
         BackColor       =   &H00C0FFFF&
         Height          =   350
         Left            =   2160
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox txtCod_Ordpro_Bus 
         BackColor       =   &H00C0FFFF&
         Height          =   350
         Left            =   12120
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin VB.ComboBox cboTipoProducto 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   120
         Width           =   1320
      End
      Begin GridEX20.GridEX GrxProductos 
         Height          =   4575
         Left            =   45
         TabIndex        =   11
         Top             =   480
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   8070
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
         Column(1)       =   "FrmGuiasRemisionPrendas.frx":0000
         Column(2)       =   "FrmGuiasRemisionPrendas.frx":00C8
         FormatStylesCount=   8
         FormatStyle(1)  =   "FrmGuiasRemisionPrendas.frx":016C
         FormatStyle(2)  =   "FrmGuiasRemisionPrendas.frx":0294
         FormatStyle(3)  =   "FrmGuiasRemisionPrendas.frx":0344
         FormatStyle(4)  =   "FrmGuiasRemisionPrendas.frx":03F8
         FormatStyle(5)  =   "FrmGuiasRemisionPrendas.frx":04D0
         FormatStyle(6)  =   "FrmGuiasRemisionPrendas.frx":0588
         FormatStyle(7)  =   "FrmGuiasRemisionPrendas.frx":0668
         FormatStyle(8)  =   "FrmGuiasRemisionPrendas.frx":06F8
         ImageCount      =   0
         PrinterProperties=   "FrmGuiasRemisionPrendas.frx":080C
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
         Left            =   13200
         TabIndex        =   16
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label41 
         BackColor       =   &H0080C0FF&
         Caption         =   "BARRA"
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
         Left            =   9720
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label34 
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
         Left            =   6840
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label39 
         BackColor       =   &H0080C0FF&
         Caption         =   "COD"
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
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label16 
         BackColor       =   &H0080C0FF&
         Caption         =   "OP"
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
         Left            =   11880
         TabIndex        =   12
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
      Left            =   9600
      TabIndex        =   43
      Top             =   8400
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
      TabIndex        =   42
      Top             =   8400
      Width           =   3375
   End
   Begin VB.Frame frMain 
      Height          =   1080
      Left            =   0
      TabIndex        =   21
      Top             =   960
      Width           =   16575
      Begin VB.TextBox txtLug_Entrega 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1005
         TabIndex        =   67
         Top             =   720
         Width           =   4890
      End
      Begin VB.TextBox txtcod_transportista 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   12240
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   64
         Top             =   720
         Width           =   825
      End
      Begin VB.TextBox txtdes_transportista 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   13080
         MaxLength       =   11
         TabIndex        =   63
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtCod_PlacaVehiculo 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   6480
         MaxLength       =   11
         TabIndex        =   62
         Top             =   720
         Width           =   3135
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
         Left            =   15480
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   720
         Width           =   735
      End
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
         Left            =   9600
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   720
         Width           =   735
      End
      Begin VB.Frame frReferencia 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         TabIndex        =   31
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
         TabIndex        =   30
         Text            =   "C"
         Top             =   420
         Width           =   465
      End
      Begin VB.TextBox txtDes_Moneda 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   9040
         TabIndex        =   29
         Top             =   120
         Width           =   1650
      End
      Begin VB.TextBox txtCod_Moneda 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   8445
         MaxLength       =   4
         TabIndex        =   28
         Top             =   120
         Width           =   600
      End
      Begin VB.TextBox txtDes_TipAne 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1485
         TabIndex        =   27
         Top             =   420
         Width           =   4425
      End
      Begin VB.TextBox txtNum_Docum 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   5850
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   26
         Top             =   120
         Width           =   2020
      End
      Begin VB.TextBox txtDes_TipDoc 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1485
         TabIndex        =   25
         Top             =   120
         Width           =   2625
      End
      Begin VB.TextBox txtCod_TipDoc 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1005
         MaxLength       =   4
         TabIndex        =   24
         Top             =   120
         Width           =   465
      End
      Begin VB.TextBox txtSer_Docum 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   4770
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   23
         Top             =   120
         Width           =   1080
      End
      Begin VB.TextBox txtNum_Ruc 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   6480
         MaxLength       =   11
         TabIndex        =   22
         Top             =   420
         Width           =   4220
      End
      Begin MSComCtl2.DTPicker dtpFec_Emision 
         Height          =   285
         Left            =   11640
         TabIndex        =   32
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
         Format          =   72613889
         CurrentDate     =   38182
      End
      Begin MSComCtl2.DTPicker dtpFec_Registro 
         Height          =   285
         Left            =   14760
         TabIndex        =   33
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
         Format          =   72613889
         CurrentDate     =   38182
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
         TabIndex        =   68
         Top             =   810
         Width           =   855
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
         Left            =   10800
         TabIndex        =   66
         Top             =   765
         Width           =   1335
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
         Left            =   5880
         TabIndex        =   65
         Top             =   720
         Width           =   615
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
         TabIndex        =   41
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
         TabIndex        =   40
         Top             =   405
         Width           =   735
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
         Left            =   13920
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label5 
         Caption         =   "Numero :"
         Height          =   255
         Left            =   9390
         TabIndex        =   36
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
         TabIndex        =   35
         Top             =   135
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
         TabIndex        =   34
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
      TabIndex        =   20
      Text            =   "G U I A    D E   R E M I S I O N"
      Top             =   0
      Width           =   16575
   End
   Begin VB.TextBox txt_descto 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   13920
      TabIndex        =   19
      Top             =   8400
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   0
      TabIndex        =   17
      Top             =   2040
      Width           =   16575
      Begin GridEX20.GridEX grxDatos 
         Height          =   5955
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   16335
         _ExtentX        =   28813
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
         Column(1)       =   "FrmGuiasRemisionPrendas.frx":09E4
         Column(2)       =   "FrmGuiasRemisionPrendas.frx":0AAC
         FormatStylesCount=   9
         FormatStyle(1)  =   "FrmGuiasRemisionPrendas.frx":0B50
         FormatStyle(2)  =   "FrmGuiasRemisionPrendas.frx":0C78
         FormatStyle(3)  =   "FrmGuiasRemisionPrendas.frx":0D28
         FormatStyle(4)  =   "FrmGuiasRemisionPrendas.frx":0DDC
         FormatStyle(5)  =   "FrmGuiasRemisionPrendas.frx":0EB4
         FormatStyle(6)  =   "FrmGuiasRemisionPrendas.frx":0F6C
         FormatStyle(7)  =   "FrmGuiasRemisionPrendas.frx":104C
         FormatStyle(8)  =   "FrmGuiasRemisionPrendas.frx":10DC
         FormatStyle(9)  =   "FrmGuiasRemisionPrendas.frx":1214
         ImageCount      =   0
         PrinterProperties=   "FrmGuiasRemisionPrendas.frx":1328
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   5640
      TabIndex        =   44
      Top             =   8280
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   900
      Custom          =   $"FrmGuiasRemisionPrendas.frx":1500
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
      TabIndex        =   45
      Top             =   8280
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   900
      Custom          =   "0~0~AYUDA~Verdadero~Verdadero~&AYUDA~0~0~1~~0~Falso~Falso~&AYUDA~"
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
      TabIndex        =   47
      Top             =   8400
      Width           =   615
   End
   Begin VB.Label Label35 
      Caption         =   "UNID."
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
      Left            =   13320
      TabIndex        =   46
      Top             =   8520
      Width           =   615
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   11160
      Top             =   8280
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmGuiasRemisionPrendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Public CODIGO As String, Descripcion As String, strOption As String, strNum_Corre As String, strCod_Anxo As String
Public rsFactura As New ADODB.Recordset
Dim strSQL As String
Dim bClickColSelec As Boolean
Dim Errorx As String
Dim rstAux As ADODB.Recordset
Dim stit As String
Public flg_Tiene_guias_asignadas As String
Public fila_seleccionada As Double
Private rsDocResumen As New ADODB.Recordset
Private Declare Function GetSystemMenu Lib "user32" _
    (ByVal hwnd As Long, _
     ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" _
    (ByVal hMenu As Long, _
     ByVal nPosition As Long, _
     ByVal wFlags As Long) As Long
     
Private Const MF_BYPOSITION = &H400&
Public iva As Double
Public prnPrinter As Object
Private indiceMedioPago As Integer
Private indiceTipo_Impresion As Integer
Private indice As Integer
Dim Contador As Double
Public Function DisableCloseButton(frm As Form) As Boolean

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
        .sQuery = strSQL
        .Cargar_Datos
        
        CODIGO = ".."
        Set rstAux = .gexList.ADORecordset
        'If rstAux.RecordCount > 1 Then
        .Show vbModal
        
        If fila_seleccionada > 0 And rstAux.RecordCount > 0 Then
            rstAux.AbsolutePosition = fila_seleccionada
            txtCod = Trim(rstAux!Cod)
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
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento (" & Opcion & ")"
End Sub



Private Sub Check2_Click()
Call muestraventasvarias
End Sub

Private Sub chkTodos_Click()
On Error GoTo fin
    If GrxProductos.RowCount = 0 Then Exit Sub
    Dim rs As New ADODB.Recordset
    Dim valor As Boolean
    Dim I As Long

    GrxProductos.Update
    Set rs = GrxProductos.ADORecordset
    rs.MoveFirst
    Do While Not rs.EOF
        If chkTodos.Value = Checked Then
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
MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
"Edicionar Producto "

End Sub

Private Sub cmdAnularGuia_Click()
On Error GoTo fin
 If grxDatos.RowCount <= 0 Then Exit Sub
 
 If MsgBox("¡¡¡Esta apunto de Anular en el almacen el documento de salida!!!:" & Chr(13) & Chr(10) & ":::::> " & Trim(txtDes_TipDoc.Text) & " " & txtSerieGuiaExis & "-" & txtNumeroGuiaExis & Chr(13) & Chr(10) & "¿Son los datos correctos?", vbYesNo, "CONFIRMAR") = vbYes Then
    If AnulaGuiacliente = True Then
        Call MuestraDetalleGuiaExiste
        Call habilitaframe(indice)
    End If
End If
 
Exit Sub
fin:
MsgBox Err.Description & ",No se puede Continuar", vbExclamation + vbOKOnly, _
"Anula Guia "
 
End Sub
Private Function AnulaGuiacliente() As Boolean
On Error GoTo fin

Dim rsset As New ADODB.Recordset
Set rsset = grxDatos.ADORecordset

AnulaGuiacliente = False

ExecuteCommandSQL cConnect, "VENTAS_MAN_ANULA_GUIA_REMISION_PRENDAS '" & rsset!Cod_almacen & "','" & rsset!num_movstk & "','" & vusu & "'"
           
AnulaGuiacliente = True

Exit Function
fin:
rsset.Close
AnulaGuiacliente = False
Set rsset = Nothing
MsgBox Err.Description & ",No se puede Continuar", vbExclamation + vbOKOnly, _
"Anula Guia "

End Function

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

Private Sub cmdBuscarGuia_Click()
    Call MuestraDetalleGuiaExiste
End Sub
Private Sub cmdBusLimpiarCaja_Click()
   Call limpiarCajasBusqueda
End Sub
'Private Sub cmdBusquedaProductos_Click()
'FraProductos.Visible = True
'limpiarCajasBusqueda
'End Sub
Private Sub limpiarCajasBusqueda()
    TxtCod_Estcli_Bus.Text = ""
    txtDes_Estcli_Bus.Text = ""
    txtDes_Present_Bus.Text = ""
    txtCodigo_Barra_Bus.Text = ""
    txtCod_Ordpro_Bus.Text = ""
End Sub

Private Sub cmdCerrarBusProductos_Click()
FraProductos.Visible = False
Set GrxProductos.ADORecordset = Nothing
End Sub
'Private Sub cmdMedioPagoAgregar_Click()
'
'If Not IsNumeric(txtMedioPagoImporte.Text) Then
' Call MsgBox("Ingrese una Cantidad Valida", vbCritical, "Mensaje")
' Call iniciofraMedioPago
' Exit Sub
'End If
'AdicionaMedioPago
'End Sub

Private Sub cmdVentasVarios_Click()
Call muestraventasvarias
End Sub
Private Sub muestraventasvarias()
On Error GoTo fin
Dim stit As String
Dim strCadena As String
stit = "Ventas Varios"
Dim rsventasvarios As New ADODB.Recordset

If (Contador Mod 2) > 0 Then

    strCadena = "CN_MUESTRA_DATOS_VENTAS_VARIOS '','001','01'"
    Set rsventasvarios = CargarRecordSetDesconectado(strCadena, cConnect)
        
    txtCod_TipDoc.Text = Trim(rsventasvarios!Cod_TipDoc)
    txtDes_TipDoc.Text = Trim(rsventasvarios!DES_TIPDOC)
    txtSer_Docum.Text = Trim(rsventasvarios!COR_DOCSERIE)
    txtNum_Docum.Text = Trim(rsventasvarios!COR_NUMACTU)

    txtNum_ruc.Text = Trim(rsventasvarios!Num_Ruc)
    txtDes_TipAne.Text = Trim(rsventasvarios!DES_ANEXO)
    txtNum_ruc.Tag = Trim(rsventasvarios!Cod_Anxo)
    txtDes_TipAne.Tag = Trim(rsventasvarios!COD_CLIENTE)
    
    'txtCod_TipVenta.Text = Trim(rsventasvarios!Cod_Tipo_Venta)
    'txtDes_TipVenta.Text = Trim(rsventasvarios!DESCRIPCION_TIPO_VENTA)
    'txtCod_ConPag.Text = Trim(rsventasvarios!Cod_CondVent)
    'txtDes_ConPag.Text = Trim(rsventasvarios!Des_CondVent)
    txtCod_Moneda.Text = Trim(rsventasvarios!Cod_Moneda)
    txtDes_Moneda.Text = Trim(rsventasvarios!Nom_Moneda)
            
Else

    txtCod_TipDoc.Text = ""
    txtDes_TipDoc.Text = ""
    txtSer_Docum.Text = ""
    txtNum_Docum.Text = ""
    txtNum_ruc.Text = ""
    txtDes_TipAne.Text = ""
    txtNum_ruc.Tag = ""
    txtDes_TipAne.Tag = ""
    'txtCod_TipVenta.Text = ""
    'txtDes_TipVenta.Text = ""
    'txtCod_ConPag.Text = ""
    'txtDes_ConPag.Text = ""
    txtCod_Moneda.Text = ""
    txtDes_Moneda.Text = ""

End If
Contador = Contador + 1
Exit Sub
fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, stit
End Sub

Private Sub dtpFec_Emision_Change()
    'txtiva.Text = DevuelveCampo("SELECT PORC_IGV  FROM TG_IGV where ano= '" & Year(dtpFec_Emision) & "' and mes= '" & Format(Month(dtpFec_Emision), "00") & "'", cConnect)
    'TxtTipo_Cambio.Text = DevuelveCampo("select isnull(Tipo_Venta,0) from cn_tipocambio where fecha = '" & dtpFec_Emision & "'", cConnect)
End Sub

Private Sub Form_Load()
   
    Contador = 1
    indice = 0
    Call DisableCloseButton(Me)
    
    FraProductos.Visible = False
    dtpFec_Emision.Value = Date
    dtpFec_Registro.Value = Date
    habilitaframe (indice)
    Call buscaDetalle_factura
    Call obtieneDatosIniciales
    Call FillTipoProducto
    
End Sub
'''************************************************************ELIMINA ARTICULO DEL DETALLE DE LA FACTURA****************************
Private Sub EliminaProducto()
    If grxDatos.RowCount = 0 Then Exit Sub
    Dim I As Integer
    Dim rstAux  As ADODB.Recordset
    grxDatos.Update
    Set rstAux = grxDatos.ADORecordset
    rstAux.MoveFirst
    I = 1
    Do While I <= rstAux.RecordCount
        
        If rstAux("ELI").Value = True Then
          rstAux.AbsolutePosition = grxDatos.RowIndex(grxDatos.Row)
          rstAux.Delete
        Else
          rstAux("ELI") = 0
        End If
        rstAux.MoveNext
        I = I + 1
    Loop
    'rstAux.Update
    Set grxDatos.ADORecordset = rstAux
    
    If grxDatos.RowCount >= 1 Then
       fraUbicacion.Enabled = False
    Else
       fraUbicacion.Enabled = True
    End If
    Call ConfiguraGrilla_Detalle
End Sub

Private Sub FillTipoProducto()
On Error GoTo fin
Dim stit As String
    
    stit = "carga tipo Producto"
    strSQL = " LG_MUESTRA_TIPO_PRODUCTO "
    
    Set rstAux = CargarRecordSetDesconectado(strSQL, cConnect)
    cboTipoProducto.Clear
    With rstAux
        If .RecordCount > 0 Then .MoveFirst
        Do Until .EOF
            cboTipoProducto.AddItem !tipo_producto & " " & !Descripcion
            .MoveNext
        Loop
        .Close
    End With
    If cboTipoProducto.ListCount > 0 Then cboTipoProducto.ListIndex = 0
    Set rstAux = Nothing
    
Exit Sub
fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, stit
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo dprDepurar
Select Case ActionName
Case Is = "GRABAR"

 If indice = 0 Then
 
  If grxDatos.RowCount <= 0 Then Exit Sub
         If validaDatosIniciales = True Then
             If MsgBox("¡¡¡Esta apunto de confirmar en caja el documento de venta!!!:" & Chr(13) & Chr(10) & ":::::> " & Trim(txtDes_TipDoc.Text) & " " & txtSer_Docum & "-" & txtNum_Docum & Chr(13) & Chr(10) & "¿Son los datos correctos?", vbYesNo, "CONFIRMAR") = vbYes Then
                If GuardaDetalleVentas = True Then
                    Call obtieneDatosIniciales
                    Call estadoInicialVentana
                    Call habilitaframe(indice)
                    Call buscaDetalle_factura
                  End If
              End If
           End If
 Else
     
    If grxDatos.RowCount <= 0 Then Exit Sub
    Call Preliminar_Docum_Ventas(grxDatos.Value(grxDatos.Columns("num_movstk").Index), grxDatos.Value(grxDatos.Columns("cod_almacen").Index))

 End If
 
  
Case Is = "CANCELAR"
 If grxDatos.RowCount > 0 Then
  If MsgBox("¡...Al cancelar esta operacion se eliminaran los datos registrados...! " & Chr(13) & Chr(10) & " ¿Esta Seguro de proseguir? ", vbYesNo, "CONFIRMAR") = vbYes Then
    Unload Me
  End If
 Else
  Unload Me
   'Call imprimebixolon270("000000000051", "000")
 End If
  
End Select

Exit Sub

Resume
dprDepurar:
errores Err.Number
End Sub
Private Function cantidadValida() As Boolean
On Error GoTo fin
Dim rxValidar  As New ADODB.Recordset

cantidadValida = True

If grxDatos.RowCount <= 0 Then Exit Function
  grxDatos.Update
  
Set rxValidar = grxDatos.ADORecordset
rxValidar.MoveFirst
Do While Not rxValidar.EOF
    If rxValidar("cant") <= 0 Then
        cantidadValida = False
        Exit Do
    End If
rxValidar.MoveNext
Loop

Exit Function
fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, stit
End Function

Private Sub estadoInicialVentana()
'''generar el sgte numero de documento
'''limpiar y txt, grilla
txtDes_TipAne.Text = ""
txtNum_ruc.Text = ""
txtDes_TipAne.Tag = ""
txtNum_ruc.Tag = ""
txtNum_Docum.Text = DevuelveCampo("SELECT COR_NUMACTU FROM CN_VENTAS_CAJAS_DOCUMENTOS WHERE COD_FABRICA='" & Txtcod_Fabrica.Text & "' AND  COD_TIENDA='001' AND COD_CAJA='01' AND COD_TIPDOC='" & Trim(txtCod_TipDoc.Text) & "' AND COR_DOCSERIE ='" & txtSer_Docum.Text & "' ", cConnect)

Set grxDatos.ADORecordset = Nothing
Set GrxProductos.ADORecordset = Nothing

End Sub
'''''***********************************GUARDA EL DETALLE DE LA FACTURA DESDE LA BUSQUEDA O CON LECTORA DE BARRAS, GENERA MOVIMIENTO DE ALMACEN ****************************
Private Function GuardaDetalleVentas() As Boolean
On Error GoTo ErrDetMov
Dim sErr As String, cntAux As New ADODB.Connection, stit As String, _
    sNum_MovStk As String, strNum_Corre  As String
Dim Kilos_Tenidos As Double
Dim RollosTeñidos As Double
Dim rstAux As New ADODB.Recordset

  GuardaDetalleVentas = False

    '''txtCod_OrdTra_Tinto = Trim(txtCod_OrdTra_Tinto)
    Kilos_Tenidos = 0
    RollosTeñidos = 0
    
    If grxDatos.RowCount = 0 Then
        MsgBox "Se debe especificar al menos un detalle", vbExclamation + vbOKCancel, stit
        Exit Function
    End If
    
   stit = "Guardar Detalle de Ventas"
    
    cntAux.Open cConnect
    cntAux.BeginTrans

    '''CABECERA VENTAS
'    StrSql = " VENTAS_UP_MAN_PRENDAS_OTROS 'I','','" & txtCod_Fabrica.Text & "','" & Trim(txtCod_Tienda.Text) & "','" & Trim(txtCod_Caja.Text) & "','" & Trim(txtCod_Vendedor.Text) & "','" & Trim(txtCod_Almacen.Text) & "','" & txtCod_TipDoc & "','" & txtSer_Docum & "','" _
'            & txtNum_Docum & "','C','" & Trim(txtNum_Ruc.Tag) & "','" & txtCod_ConPag & "','" & txtCod_TipVenta.Text & "','" & Format(dtpFec_Emision.Value, "dd/mm/yyyy") & "','" _
'            & Format(dtpFec_Registro.Value, "dd/mm/yyyy") & "','" & txtCod_Moneda & "','" _
'            & vusu & "',''," _
'            & TxtTipo_Cambio.Text & ",'','','N','N','S'," & txtMedioPagoTotalPagoMN.Text & "," & txtMedioPagoVueltoMN.Text & ""
'
'    Set rstAux = cntAux.Execute(StrSql, adExecuteNoRecords)
'    strNum_Corre = rstAux!Num_Corre
'    rstAux.Close

    '''CABECERA MOVIMIENTO
    strSQL = "EXEC TI_UP_MAN_LG_MOVISTK_PRENDAS_OTROS 'I', '" & _
             Trim(txtCod_Almacen.Text) & "', '', '" & Format(dtpFec_Registro.Value, _
             "dd/mm/yyyy") & "', '' ,'" & Trim(Txtcod_Fabrica.Text) & "','', '" & txtDes_TipAne.Tag & "','','" & vusu & "','" & Trim(txtSer_Docum.Text) & "','" & Trim(txtNum_Docum.Text) & "'"

    Set rstAux = cntAux.Execute(strSQL, adExecuteNoRecords)
    sNum_MovStk = rstAux!num_movstk
    rstAux.Close
    
    Set rstAux = grxDatos.ADORecordset
    With rstAux
        If .RecordCount > 0 Then .MoveFirst
        Do Until .EOF
    '''DETALLE MOVIMIENTO DE SALIDA DE ALMACEN
             strSQL = "EXEC LG_UP_MAN_LG_MOVISTKITEM_PRENDAS_OTROS 'I', '" & _
             Trim(txtCod_Almacen.Text) & "','" & sNum_MovStk & "','" & Now() & "', '', '" & _
             Trim(!COD_ITEM) & "','" & Trim(!cod_comb) & "','" & Trim(!cod_Color) & "','" & Trim(!cod_estcli) & "','" & Trim(!cod_ordpro) & "'," & Trim(!cod_present) & ",'" & Trim(!cod_talla) & "','" & Trim(!codigo_barra) & "', '" & _
             Trim(!tipo_producto) & "'," & !cant & ",'" & vusu & "'"
             cntAux.Execute strSQL, adExecuteNoRecords
 
    '''DETALLE VENTAS falta strCod_Anxo
'            StrSql = "CN_VENTAS_ITEMS_PRENTAS_OTROS 'I','" & strNum_Corre & "','','" & Trim(!COD_ITEM) & "','" & Trim(!cod_comb) & "','" & Trim(!cod_Color) & "','" & _
'            Trim(!COD_CLIENTE) & "','" & Trim(!cod_purord) & "','" & Trim(!cod_lotpurord) & "','" & Trim(!cod_colcli) & "','" & Trim(!cod_estcli) & "','" & Trim(!cod_ordpro) & "'," _
'            & Trim(!cod_present) & ",'" & Trim(!cod_talla) & "','" & Trim(!codigo_barra) & "','" & !tipo_producto & "'," & !cant & "," & !precio & ", " & !Total & " ,'" & _
'            Trim(!des_estcli) & "','" & Trim(!des_present) & "','" & Trim(!des_comb) & "',0,'','',0,'" & vusu & "'"
'            cntAux.Execute StrSql, adExecuteNoRecords
            .MoveNext

        Loop
    End With
    
    '''ASOCIA FACTURA CON MOVIMIENTO DE ALMACEN
    'StrSql = "CN_VENTAS_CAJAS_RELACIONA_FACTURA_GUIA_PRENDAS 'U','" & strNum_Corre & "','" & Trim(txtSer_Docum.Text) & "','" & Trim(txtNum_Docum.Text) & "','" & Trim(txtCod_Almacen.Text) & "','" & sNum_MovStk & "'"
    'cntAux.Execute StrSql, adExecuteNoRecords

    cntAux.CommitTrans
    cntAux.Close
    Set cntAux = Nothing
    
    '''IMPRIME DOCUMENTO
    Call Preliminar_Docum_Ventas(sNum_MovStk, Trim(txtCod_Almacen.Text))
    
    GuardaDetalleVentas = True
    'Unload Me
Exit Function
ErrDetMov:
    GuardaDetalleVentas = False
    sErr = Err.Description
    cntAux.RollbackTrans
    cntAux.Close
    Set cntAux = Nothing
    MsgBox sErr, vbCritical + vbOKOnly, stit
End Function
'''''***********************************GUARDA EL DETALLE DE LA FACTURA DESDE EL DETALLE DE LA GUIA****************************
Private Function GuardaDetalleVentasGuias() As Boolean
On Error GoTo ErrDetMov
Dim sErr As String, cntAux As New ADODB.Connection, stit As String, _
    sNum_MovStk As String, strNum_Corre  As String
Dim Kilos_Tenidos As Double
Dim RollosTeñidos As Double
Dim rstAux As New ADODB.Recordset

  GuardaDetalleVentasGuias = False

    '''txtCod_OrdTra_Tinto = Trim(txtCod_OrdTra_Tinto)
    Kilos_Tenidos = 0
    RollosTeñidos = 0
    
    If grxDatos.RowCount = 0 Then
        MsgBox "Se debe especificar al menos un detalle", vbExclamation + vbOKCancel, stit
        Exit Function
    End If
    
   stit = "Guardar Detalle de Ventas"
    
    cntAux.Open cConnect
    cntAux.BeginTrans

    '''CABECERA VENTAS
'    StrSql = "VENTAS_UP_MAN_ROLLOS 'I','','" & txtCod_Fabrica.Text & "','" & Trim(txtCod_Tienda.Text) & "','" & Trim(txtCod_Caja.Text) & "','" & Trim(txtCod_Vendedor.Text) & "','" & Trim(txtCod_Almacen.Text) & "','" & txtCod_TipDoc & "','" & txtSer_Docum & "','" _
'            & txtNum_Docum & "','C','" & Trim(txtNum_Ruc.Tag) & "','" & txtCod_ConPag & "','" & txtCod_TipVenta.Text & "','" & Format(dtpFec_Emision.Value, "dd/mm/yyyy") & "','" _
'            & Format(dtpFec_Registro.Value, "dd/mm/yyyy") & "','" & txtCod_Moneda & "','" _
'            & vusu & "',''," _
'            & TxtTipo_Cambio.Text & ",'','','N','N','N'"
'
'    Set rstAux = cntAux.Execute(StrSql, adExecuteNoRecords)
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
'    '''DETALLE MOVIMIENTO DE SALIDA DE ALMACEN
'             STRSQL = "EXEC TI_UP_MAN_TX_MOVISTK_TELA_TENIDA_PESADAS_ROLLO_VENTAS_DIRECTA 'I', '" & _
'             Trim(txtCod_Almacen.Text) & "', '" & sNum_MovStk & "', '', '" & _
'             !codigorollo & "'," & !Stock & "," & !cant & ",0, " & _
'             Trim(!rollos) & ",'" & vusu & "',0"
'             cntAux.Execute STRSQL, adExecuteNoRecords
    
    '''DETALLE VENTAS falta strCod_Anxo
            strSQL = "VENTAS_UP_MAN_DETALLE_ROLLO 'I','" & strNum_Corre & "','','D','" & Trim(!codigoRollo) & "','" & _
            Trim(!cod_tela) & "','','" & !und & "'," & !rollos & "," & !Stock & "," & !cant & "," _
            & !precio & "," & !Total & ",0,'','',0,'" & Trim(txtCod_Almacen.Text) & "','" & !OT & "','" & vusu & "'"
            cntAux.Execute strSQL, adExecuteNoRecords
            .MoveNext
        Loop
    End With
    
    '''ASOCIA FACTURA CON MOVIMIENTO DE ALMACEN
    strSQL = "CN_VENTAS_CAJAS_RELACIONA_FACTURA_GUIA 'U','" & strNum_Corre & "','" & Trim(txtSer_Docum.Text) & "','" & Trim(txtNum_Docum.Text) & "','" & Trim(txtCod_Almacen.Text) & "',''"
    cntAux.Execute strSQL, adExecuteNoRecords

    cntAux.CommitTrans
    cntAux.Close
    Set cntAux = Nothing
    
    '''IMPRIME DOCUMENTO
    Call Preliminar_Docum_Ventas(sNum_MovStk, Trim(txtCod_Almacen.Text))
    GuardaDetalleVentasGuias = True
    'Unload Me
Exit Function
ErrDetMov:
    GuardaDetalleVentasGuias = False
    sErr = Err.Description
    cntAux.RollbackTrans
    cntAux.Close
    Set cntAux = Nothing
    MsgBox sErr, vbCritical + vbOKOnly, stit
End Function

Private Sub Preliminar_Docum_Ventas(snun_movstk As String, sCod_Almacen As String)
On Error GoTo SALTO_ERROR
Dim sSQL As String, rs As New ADODB.Recordset
Dim imp_total As Double

Dim aMess(4), I As Integer

If Imprimir_FACTURA(sCod_Almacen, snun_movstk, imp_total, Trim(txtCod_TipDoc.Text), Trim(txtSer_Docum.Text)) = False Then
   MsgBox "Problemas de Impresion con el Documento Nro " & txtNum_Docum.Text, vbInformation, "ERROR"
   'Buscar
   Exit Sub
End If
    
Exit Sub
SALTO_ERROR:
MsgBox Err.Description, vbCritical, Me.Caption
    
End Sub
   
Public Function Imprimir_FACTURA(vcod_almacen As String, vnum_movstk As String, dbImp_Total As Double, strCod_Cod As String, Serie As String) As Boolean

Dim Rs2 As New ADODB.Recordset, rs3 As New ADODB.Recordset, strSQL As String, scnt As Integer
scnt = 0
With rsFactura
 
    Select Case strCod_Cod
    
    Case Is = "GR" 'llll
        
        strSQL = "CN_MUESTRA_IMPRESION_DOCUMENTO_PRENDA_OTROS_GUIA_CLIENTE  '" & vcod_almacen & "' ,'" & vnum_movstk & "','" & UCase(EnLetras(Trim(CStr(dbImp_Total)))) & "'"
        Set rsFactura = CargarRecordSetDesconectado(strSQL, cConnect)
        
        If rsFactura.RecordCount > 0 Then
            Call Factura_sa("GR", Serie)
            scnt = 2
        Else
           Call MsgBox("La Factura no Tiene Detalle", vbInformation, "Mensaje")
           Imprimir_FACTURA = False
           Exit Function
        End If
        
    Case Else
      MsgBox "No se ha Definido un Formato de Impresion para este tipo de documento", vbInformation, "ERROR"
       Imprimir_FACTURA = False
      Exit Function
    End Select
End With
Imprimir_FACTURA = True
End Function
Sub Factura_sa(Tipo As String, Serie As String)
On Error GoTo ErrorImpresion
Dim oo As Object, lvSql As String, lvRuta As String

    Set oo = CreateObject("excel.application")
    
    If Tipo = "GR" Then
            oo.Workbooks.Open vRuta & "\guia_prendas_otros.XLT"
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
    MsgBox "Hubo error en la impresion de La Factura " & Err.Description, vbCritical, "Impresion"
End Sub
Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "AYUDA"
        
        'If fraSelGuias.Visible = False And flg_Tiene_guias_asignadas = "N" Then
         If indice = 0 Then
            FraProductos.Visible = True
            limpiarCajasBusqueda
         End If
        'End If
End Select


End Sub

''''******************************HABILITA LA EDICION SOLO DE ALGUNAS COLUMNAS LAS TIENEN CANCEL=FALSE***********************
'Private Sub grxMedioPagos_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
'  Select Case ColIndex
'    Case Is = grxMedioPagos.Columns("ELI").Index
'      Cancel = False
'    Case Else
'      Cancel = True
'  End Select
'End Sub

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
    Dim I As Integer
    
    If fraUbicacion.Enabled = False Then
    
        If Trim(txtCod_Almacen.Text) = "" Then
           Call MsgBox("El Codigo del Almacen no es valido", vbCritical, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
       
        If txtCod_TipDoc.Text = "" Then
           Call MsgBox("Sirvase a Ingresar un tipo de documento Valido", vbCritical, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
        
        If txtSer_Docum.Text = "" Then
           Call MsgBox("Sirvase a Ingresar una serie de documento valido", vbCritical, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
        If txtNum_Docum.Text = "" Then
           Call MsgBox("Sirvase a Ingresar un Numero de documento Valido", vbCritical, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
        If txtCod_Moneda.Text = "" Then
           Call MsgBox("Sirvase a Ingresar una Moneda Valida", vbCritical, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
         
        If txtNum_ruc.Text = "" Or txtDes_TipAne.Text = "" Then
           Call MsgBox("Sirvase a Ingresar un cliente Valido", vbCritical, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If

        
        If cantidadValida = False Then
           Call MsgBox("Sirvase a ingresar una Cantidad Valida ", vbCritical, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
      
    End If
End Function
'''******************************* ADICIONA LISTA ARTICULOS CUYA CANTIDAD SEA MAYOR A 0*******************************************
Private Sub adicionarProductoMasivo()
Dim RSAUX As ADODB.Recordset
Dim rslista As ADODB.Recordset
Dim I As Integer
On Error GoTo fin

'If validaDatosIniciales = False Then
'    Exit Sub
'End If

GrxProductos.Refresh
GrxProductos.Update

Set RSAUX = grxDatos.ADORecordset

Set rslista = GrxProductos.ADORecordset
If rslista.RecordCount <= 0 Then Exit Sub

rslista.Update
'RSAUX.Update
rslista.MoveFirst
I = 1
Do While I <= rslista.RecordCount
If rslista!cant > 0 Then

    RSAUX.AddNew
    RSAUX!COD_CLIENTE = rslista!COD_CLIENTE
    RSAUX!nom_cliente = rslista!nom_cliente
    RSAUX!cod_temcli = rslista!cod_temcli
    RSAUX!nom_temcli = rslista!nom_temcli
    RSAUX!cod_purord = rslista!cod_purord
    
    RSAUX!cod_lotpurord = rslista!cod_lotpurord
    RSAUX!cod_colcli = rslista!cod_colcli
    
    RSAUX!cod_ordpro = rslista!cod_ordpro
    RSAUX!codigo_barra = rslista!codigo_barra
    RSAUX!cod_estcli = rslista!cod_estcli
    RSAUX!des_estcli = rslista!des_estcli
    RSAUX!cod_present = rslista!cod_present
    RSAUX!des_present = rslista!des_present
    RSAUX!cod_talla = rslista!cod_talla
    RSAUX!Stock = rslista!Stock
    RSAUX!cant = rslista!cant
    RSAUX!precio = rslista!precio
    RSAUX!Total = RSAUX!precio * RSAUX!cant
    RSAUX!DEL = "X"
    RSAUX!cod_Color = rslista!cod_Color
    RSAUX!cod_comb = rslista!cod_comb
    RSAUX!tipo_producto = rslista!tipo_producto
    RSAUX!COD_ITEM = rslista!COD_ITEM

    RSAUX.Update
    

End If
rslista.MoveNext
I = I + 1
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
MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
"Edicionar Producto "
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
txtNum_ruc.Text = ""
txtLug_Entrega.Text = ""
txtSerieGuiaExis.Text = ""
txtNumeroGuiaExis.Text = ""

Set grxDatos.ADORecordset = Nothing
Set GrxProductos.ADORecordset = Nothing

If Opcion = 0 Then
    
    frMain.Enabled = True
    'cmdBusquedaProductos.Enabled = True
    txtCodigo_Producto.Enabled = True
    txtSerieGuiaExis.Enabled = False
    txtNumeroGuiaExis.Enabled = False
    grxDatos.Enabled = True
    grxDatos.AllowEdit = True
    cmdAnularGuia.Enabled = False
    cmdBuscarGuia.Enabled = False
End If

If Opcion = 1 Then
    frMain.Enabled = False
    'cmdBusquedaProductos.Enabled = False
    txtCodigo_Producto.Enabled = False
    txtSerieGuiaExis.Enabled = True
    txtNumeroGuiaExis.Enabled = True
    grxDatos.AllowEdit = False
    cmdAnularGuia.Enabled = True
    cmdBuscarGuia.Enabled = True
End If

Call obtieneDatosIniciales
Call estadoInicialVentana
Call buscaDetalle_factura

End Sub
Private Sub obtieneDatosIniciales()
    Dim strSQL As String
    Dim pc As String
    Dim auxset As ADODB.Recordset
    pc = ComputerName
End Sub
Private Sub TxtCod_Estcli_Bus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Left(cboTipoProducto, 2) = "01" Then
        buscarProductos (0)
    Else
        buscarProductosOtros (0)
    End If
End If
End Sub

Private Sub txtCod_Ordpro_Bus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Left(cboTipoProducto, 2) = "01" Then
            buscarProductos (5)
       Else
            Call MsgBox("Opcion Otros no tiene ordenes de produccion", vbApplicationModal, "Mensaje")
       End If
    End If
End Sub
Private Sub txtCodigo_Barra_Bus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Left(cboTipoProducto, 2) = "01" Then
        buscarProductos (3)
    Else
        buscarProductosOtros (3)
    End If
End If
End Sub

Private Sub txtCodigo_Producto_Change()
  If Len(Trim(txtCodigo_Producto.Text)) = 13 Then
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
    
    'If validaDatosIniciales = False Then
    '    Exit Sub
    'End If

Opcion = "4"
If IsNumeric(Left(Trim(txtCodigo_Producto.Text), 2)) Then
'''codigo de prenda cod_ordpro/cod_present/cod_talla--321450001000M
strSQL = "EXEC CF_MUESTRA_PRENDAS_MOV_TIENDA_ITEMS_FACTURA        '" & Opcion & _
                                                    "','" & Trim(txtCod_Almacen.Text) & _
                                                    "','','','" & Trim(txtDes_Present_Bus.Text) & _
                                                    "','','" & Trim(TxtCod_Estcli_Bus.Text) & _
                                                    "','" & Trim(txtDes_Estcli_Bus.Text) & _
                                                    "','" & Trim(txtCod_Ordpro_Bus.Text) & _
                                                    "','','" & Trim(txtCodigo_Producto.Text) & "'"
                                                    
Else
'''codigo de otros articulo cod_item/correlativo--pr00000100001
        strSQL = "EXEC CF_MUESTRA_ITEMS_TIENDA_FACTURA        '" & Opcion & _
                                                    "','" & Trim(txtCod_Almacen.Text) & _
                                                    "','" & Trim(TxtCod_Estcli_Bus.Text) & _
                                                    "','" & Trim(txtDes_Estcli_Bus.Text) & _
                                                    "','" & Trim(txtDes_Present_Bus.Text) & _
                                                    "','" & Trim(txtCodigo_Barra_Bus.Text) & "'"

End If
    'strSQL = "TX_MUESTRA_ROLLOS_VENTA '" & Opcion & "','" & Trim(txtCod_Almacen.Text) & "','" & Trim(txtCodigo_Producto.Text) & "','" & Trim(txtBus_Cod_ordtra.Text) & "','" & Trim(txtDescripcion_Producto.Text) & "','" & Trim(txtBus_Des_Color.Text) & "'"
    Set rsetbusqueda = Nothing
    Set rsetbusqueda = CargarRecordSetDesconectado(strSQL, cConnect)
    If rsetbusqueda.RecordCount <= 0 Then
       Call MsgBox("Articulo No existe o no hay Stocks", vbCritical, "Mensaje")
       Exit Sub
    End If
    Set rsetAux = grxDatos.ADORecordset
 
    rsetAux.AddNew
    rsetAux!COD_CLIENTE = rsetbusqueda!COD_CLIENTE
    rsetAux!nom_cliente = rsetbusqueda!nom_cliente
    rsetAux!cod_temcli = rsetbusqueda!cod_temcli
    rsetAux!nom_temcli = rsetbusqueda!nom_temcli
    rsetAux!cod_purord = rsetbusqueda!cod_purord
    
    rsetAux!cod_colcli = rsetbusqueda!cod_colcli
    rsetAux!cod_lotpurord = rsetbusqueda!cod_lotpurord
    
    rsetAux!cod_ordpro = rsetbusqueda!cod_ordpro
    rsetAux!codigo_barra = rsetbusqueda!codigo_barra
    rsetAux!cod_estcli = rsetbusqueda!cod_estcli
    rsetAux!des_estcli = rsetbusqueda!des_estcli
    rsetAux!cod_present = rsetbusqueda!cod_present
    rsetAux!des_present = rsetbusqueda!des_present
    rsetAux!cod_talla = rsetbusqueda!cod_talla
    rsetAux!Stock = rsetbusqueda!Stock
    rsetAux!cant = rsetbusqueda!cant
    rsetAux!precio = rsetbusqueda!precio
    rsetAux!Total = rsetbusqueda!precio * rsetbusqueda!cant
    rsetAux!DEL = "X"
    rsetAux!cod_Color = rsetbusqueda!cod_Color
    rsetAux!cod_comb = rsetbusqueda!cod_comb
    rsetAux!tipo_producto = rsetbusqueda!tipo_producto
    rsetAux!COD_ITEM = rsetbusqueda!COD_ITEM
    
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
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
Private Sub Total_documento()
On Error GoTo ErrCal
    Dim Total As Double
    Dim ColIndex As Long
    Dim totalkilos As Double
    Dim merma As Double
    Dim mermavar As Variant
    Dim rds As New ADODB.Recordset
    
    Dim I As Integer
    Total = 0
    totalkilos = 0
    'grxDatos.Update
    I = 1
    grxDatos.Update
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
            Do While I <= rds.RecordCount
                
                Total = Total + rds("total").Value
                totalkilos = totalkilos + rds("cant").Value
              
                If I < rds.RecordCount Then
                    rds.MoveNext
                End If
                I = I + 1
            Loop
            txt_descto.Text = totalkilos
            
            
     Else
            txt_descto.Text = totalkilos
     
     End If

   Exit Sub
ErrCal:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Cargar Calidades"
End Sub

'''*******************EVENTOS POR COLUMNA **********************************************************
Private Sub grxDatos_AfterColEdit(ByVal ColIndex As Integer)
  AfterColEdit_DETALLE_FACTURA (ColIndex)
End Sub
Sub AfterColEdit_DETALLE_FACTURA(ByVal ColIndex As Integer)
Dim sSQL As String
On Error GoTo Error_Handler

Dim oGroup As GridEX20.JSGroup
Select Case ColIndex

  Case Is = grxDatos.Columns("ELI").Index
        Call EliminaProducto
   'Call Total_documento
  Case Is = grxDatos.Columns("PRECIO").Index
  
    If IsNumeric(grxDatos.Value(grxDatos.Columns("PRECIO").Index)) = False Or grxDatos.Value(grxDatos.Columns("PRECIO").Index) = "" Then
        grxDatos.Value(grxDatos.Columns("PRECIO").Index) = 0
    End If
    grxDatos.Value(grxDatos.Columns("TOTAL").Index) = grxDatos.Value(grxDatos.Columns("PRECIO").Index) * grxDatos.Value(grxDatos.Columns("CANT").Index)
    Call Total_documento
    'GridEX1.Value(GridEX1.Columns("sel").Index) = 0
    
  Case Is = grxDatos.Columns("CANT").Index
  
     If IsNumeric(grxDatos.Value(grxDatos.Columns("CANT").Index)) = False Or grxDatos.Value(grxDatos.Columns("CANT").Index) = "" Then
         grxDatos.Value(grxDatos.Columns("CANT").Index) = 0
     End If
    grxDatos.Value(grxDatos.Columns("TOTAL").Index) = grxDatos.Value(grxDatos.Columns("PRECIO").Index) * grxDatos.Value(grxDatos.Columns("CANT").Index)
    Call Total_documento
    'GridEX1.Value(GridEX1.Columns("sel").Index) = 0
  End Select
Exit Sub

Resume
Error_Handler:
errores Err.Number
End Sub
Private Sub GrxProductos_AfterColEdit(ByVal ColIndex As Integer)
On Error GoTo fin
  Dim a As Integer
  AfterColEdit_PRODUCTOS (ColIndex)

  Exit Sub
fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub

Sub AfterColEdit_PRODUCTOS(ByVal ColIndex As Integer)
Dim sSQL As String
Dim saldo As Double
'On Error GoTo Error_Handler
On Error GoTo fin

Dim oGroup As GridEX20.JSGroup
Select Case ColIndex

  Case Is = GrxProductos.Columns("CANT").Index
    If IsNumeric(GrxProductos.Value(GrxProductos.Columns("CANT").Index)) = False Or GrxProductos.Value(GrxProductos.Columns("CANT").Index) = "" Then
        GrxProductos.Value(GrxProductos.Columns("CANT").Index) = 0
    End If
    GrxProductos.Value(GrxProductos.Columns("TOTAL").Index) = GrxProductos.Value(GrxProductos.Columns("PRECIO").Index) * CDbl(GrxProductos.Value(GrxProductos.Columns("CANT").Index))

    
  End Select
Exit Sub
fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
Private Sub grxDatos_Click()
    Dim ColIndex As Long
    Dim oRowData As JSRowData
    Dim SGRUPO As String
    Dim iRow As Long
    Dim I As Long
    Dim sCaptionGroup As String
        If grxDatos.RowCount > 0 Then
        ColIndex = grxDatos.Col
         If ColIndex = 0 Then Exit Sub
         
            If UCase(grxDatos.Columns(ColIndex).Key) = "ELI" Then
                bClickColSelec = True
                SendKeys "{ENTER}"
'            ElseIf UCase(grxDatos.Columns(ColIndex).Key) = "CANT" Then
'                If IsNumeric(grxDatos.Value(grxDatos.Columns("CANT").Index)) = False Then
'                    grxDatos.Value(grxDatos.Columns("CANT").Index) = 0
'                End If
            End If
    End If
End Sub
Private Sub GrxProductos_Click()

    Dim ColIndex As Long
    Dim oRowData As JSRowData
    Dim SGRUPO As String
    Dim iRow As Long
    Dim I As Long
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
Private Sub txtDescripcion_Producto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call buscarProductos(3)
    End If
    
End Sub
''''*************************************************************BUSQUEDA DE PRODUCTOS *********************************
Private Sub buscarProductos(Opcion As String)

Dim strSQL As String
Dim sCodCentroCosto As String
Dim nrofilas As Integer
Dim k, l As Long
Dim rsproductos  As New ADODB.Recordset
On Error GoTo fin

If txtCod_Almacen.Text = "" Then
Call MsgBox("Sirvase ingresar un almacen", vbCritical, "Mensaje")
Exit Sub
End If
   
strSQL = "EXEC CF_MUESTRA_PRENDAS_MOV_TIENDA_ITEMS_FACTURA        '" & Opcion & _
                                                    "','" & Trim(txtCod_Almacen.Text) & _
                                                    "','','','" & Trim(txtDes_Present_Bus.Text) & _
                                                    "','','" & Trim(TxtCod_Estcli_Bus.Text) & _
                                                    "','" & Trim(txtDes_Estcli_Bus.Text) & _
                                                    "','" & Trim(txtCod_Ordpro_Bus.Text) & _
                                                    "','','" & Trim(txtCodigo_Barra_Bus.Text) & "'"
                                                    
    Set GrxProductos.ADORecordset = Nothing
    Set GrxProductos.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    If GrxProductos.RowCount <= 0 Then
     Call ConfiguraGrilla_productos
    Exit Sub
    
    End If

    GrxProductos.Update
    Set rsproductos = GrxProductos.ADORecordset
    rsproductos.Update
    rsproductos.MoveFirst
    Do While Not rsproductos.EOF
       rsproductos!Stock = rsproductos!Stock - SumaTotalRollo(Trim(rsproductos!codigo_barra))
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
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
''''*************************************************************BUSQUEDA DE PRODUCTOS *********************************
Private Sub buscarProductosOtros(Opcion As String)

Dim strSQL As String
Dim sCodCentroCosto As String
Dim nrofilas As Integer
Dim k, l As Long
Dim rsproductos  As New ADODB.Recordset
On Error GoTo fin
   
strSQL = "EXEC CF_MUESTRA_ITEMS_TIENDA_FACTURA        '" & Opcion & _
                                                    "','" & Trim(txtCod_Almacen.Text) & _
                                                    "','" & Trim(TxtCod_Estcli_Bus.Text) & _
                                                    "','" & Trim(txtDes_Estcli_Bus.Text) & _
                                                    "','" & Trim(txtDes_Present_Bus.Text) & _
                                                    "','" & Trim(txtCodigo_Barra_Bus.Text) & "'"
                                                    
    Set GrxProductos.ADORecordset = Nothing
    Set GrxProductos.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    If GrxProductos.RowCount <= 0 Then
        Call ConfiguraGrilla_productos
        Exit Sub
    End If
    GrxProductos.Update
    Set rsproductos = GrxProductos.ADORecordset
    rsproductos.Update
    rsproductos.MoveFirst
    Do While Not rsproductos.EOF
       rsproductos!Stock = rsproductos!Stock - SumaTotalRollo(Trim(rsproductos!codigo_barra))
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
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
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
   Set rsproductos = Nothing
Exit Sub
fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
Private Function SumaTotalRollo(codigo_barra As String) As Double
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
        If Trim(codigo_barra) = Trim(rssum!codigo_barra) Then
             pesorollo = pesorollo + rssum!cant
        End If
        rssum.MoveNext
    Loop
    rssum.MoveFirst
    rssum.Update
      
    SumaTotalRollo = pesorollo
Exit Function
fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
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
        
        With .Columns("NOM_TEMCLI")
             .Visible = False
             .Width = 1000
             .Caption = "MARCA"
             .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("COD_PURORD")
            .Visible = False
            .Width = 1500
            .Caption = "PO"
            .TextAlignment = jgexAlignLeft
        End With
                        
        With .Columns("CODIGO_BARRA")
            .Visible = True
            .Width = 1500
            .Caption = "CODIGO"
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("cod_estcli")
            .Visible = True
            .Width = 1500
            .Caption = "COD-EST"
            .TextAlignment = jgexAlignLeft
        End With
        
        
        With .Columns("DES_ESTCLI")
            .Visible = True
            .Width = 3000
            .Caption = "ESTILO"
            .TextAlignment = jgexAlignLeft
        End With
                        
        With .Columns("Cod_OrdPro")
            .Visible = True
            .Width = 800
            .Caption = "OP"
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("Cod_Present")
            .Visible = False
            .Width = 500
            .Caption = "COD_PRESENT"
            .TextAlignment = jgexAlignLeft
        End With
        
        'Presentacion
        With .Columns("DES_PRESENT")
            .Visible = True
            .Width = 2000
            .Caption = "PRESENTACION"
            .TextAlignment = jgexAlignLeft
        End With
        With .Columns("Cod_Talla")
            .Visible = True
            .Width = 600
            .Caption = "TALLA"
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("STOCK")
            .Visible = True
            .Width = 1300
            .Caption = "PRENDAS"
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("CANT")
            .Visible = True
            .Width = 1300
            .Caption = "PRENDASTRANF"
            .TextAlignment = jgexAlignLeft
        End With
        With .Columns("PRECIO")
            .Visible = False
            .Width = 1300
            .Caption = "PRECIO"
            .TextAlignment = jgexAlignLeft
        End With
        With .Columns("TOTAL")
            .Visible = False
            .Width = 1300
            .Caption = "TOTAL"
            .TextAlignment = jgexAlignLeft
        End With

        SetColorDetalle
    End With
    
    Dim saldo As Double
    Set fmtCon = GrxProductos.FmtConditions.Add(GrxProductos.Columns("CANT").Index, jgexGreaterThan, 0)
    fmtCon.FormatStyle.BackColor = &H80FFFF   ' &HFFFF00
    SetColores
    
    Exit Sub
fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
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
   
    strSQL = "EXEC CN_MUESTRA_TELAS_DETALLE_FACTURA_PRENDAS 'x','',''"
    
    Set grxDatos.ADORecordset = Nothing
    Set grxDatos.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    
    Call ConfiguraGrilla_Detalle
    Exit Sub
fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
''''*******************************************CONFIGURA DETALLE DE DOCUMENTO*********************************
Private Sub ConfiguraGrilla_Detalle()
    Dim C As Integer
    On Error GoTo fin
    
    With grxDatos
        
        For C = 1 To .Columns.Count
            .Columns(C).HeaderAlignment = jgexAlignCenter
            .Columns(C).TextAlignment = jgexAlignLeft
            .Columns(C).Visible = False
        Next C
        
        With .Columns("NOM_TEMCLI")
             .Visible = False
             .Width = 1000
             .Caption = "MARCA"
             .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("COD_PURORD")
            .Visible = False
            .Width = 1500
            .Caption = "PO"
            .TextAlignment = jgexAlignLeft
        End With
                        
        With .Columns("CODIGO_BARRA")
            .Visible = True
            .Width = 1500
            .Caption = "CODIGO"
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("COD_ESTCLI")
            .Visible = True
            .Width = 1500
            .Caption = "COD-EST"
            .TextAlignment = jgexAlignLeft
        End With
        
        
        With .Columns("DES_ESTCLI")
            .Visible = True
            .Width = 3000
            .Caption = "ESTILO"
            .TextAlignment = jgexAlignLeft
        End With
                        
        With .Columns("Cod_OrdPro")
            .Visible = True
            .Width = 800
            .Caption = "OP"
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("Cod_Present")
            .Visible = False
            .Width = 500
            .Caption = "COD_PRESENT"
            .TextAlignment = jgexAlignLeft
        End With
        
        'Presentacion
        With .Columns("DES_PRESENT")
            .Visible = True
            .Width = 2000
            .Caption = "PRESENTACION"
            .TextAlignment = jgexAlignLeft
        End With
        With .Columns("Cod_Talla")
            .Visible = True
            .Width = 600
            .Caption = "TALLA"
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("STOCK")
            .Visible = True
            .Width = 1300
            .Caption = "PRENDAS"
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("CANT")
            .Visible = True
            .Width = 1300
            .Caption = "PRENDASTRANF"
            .TextAlignment = jgexAlignLeft
        End With
        With .Columns("PRECIO")
            .Visible = False
            .Width = 1300
            .Caption = "PRECIO"
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("TOTAL")
            .Visible = False
            .Width = 1300
            .Caption = "TOTAL"
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("DEL")
            .Visible = True
            .Width = 500
            .Caption = "ELI"
            .TextAlignment = jgexAlignLeft
        End With
        With .Columns("ELI")
            .Visible = True
            .Width = 500
            .Caption = ""
            .TextAlignment = jgexAlignLeft
        End With
        'SetColorDetalle
    End With
    
 SetColorDetalle
 Call Total_documento

    Exit Sub
fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
Private Sub ConfiguraGrilla_DetalleSinGrupos()
    Dim C As Integer
    On Error GoTo fin
    
 'SetColorDetalle
 'Call Total_documento
        
    Exit Sub
fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub

Private Sub SetColorDetalle()
        'grxDatos.Columns("ROLLOS").CellStyle = "estilo_cantidad"
        grxDatos.Columns("CANT").CellStyle = "estilo_cantidad"
        grxDatos.Columns("PRECIO").CellStyle = "estilo_cantidad"
        grxDatos.Columns("ELI").CellStyle = "estilo_eliminar"
        grxDatos.Columns("DEL").CellStyle = "estilo_eliminar"
End Sub
Private Sub txtDes_Estcli_Bus_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        If Left(cboTipoProducto, 2) = "01" Then
            buscarProductos (1)
        Else
            buscarProductosOtros (1)
        End If
   End If
End Sub

'Private Sub txtDes_Moneda_KeyPress(KeyAscii As Integer)
'  If KeyAscii = 13 Then
'    Call Busca_Opcion("Cod_Moneda", "Nom_Moneda", "TG_Moneda where ", txtCod_Moneda, txtDes_Moneda, 2)
'    If Trim(txtDes_Moneda.Text) <> "" Then
'       txtCod_TipVenta.SetFocus
'    Else
'       txtDes_Moneda.SetFocus
'    End If
'  End If
'End Sub
Private Sub txtDes_Present_Bus_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        If Left(cboTipoProducto, 2) = "01" Then
            buscarProductos (2)
        Else
            buscarProductosOtros (2)
        End If
   End If
    
End Sub

Private Sub txtDes_TipAne_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        Load FrmBusqClientesPrendas
        
        Set FrmBusqClientesPrendas.oParent = Me
        FrmBusqClientesPrendas.txtDescripcion_Cliente.Text = txtDes_TipAne.Text
        FrmBusqClientesPrendas.txtRuc_Cliente.Text = "" 'txtNum_Ruc.Text
        FrmBusqClientesPrendas.txtTip_Anex.Text = "C"
        
        Call FrmBusqClientesPrendas.Busca_Opcion_AnexoContable("2", "C", txtNum_ruc.Text, txtDes_TipAne.Text)
        FrmBusqClientesPrendas.Show 1
        'FrmBusqClientes.txtDescripcion_Cliente.SetFocus
        Set FrmBusqClientesPrendas = Nothing
        
        If Trim(txtNum_ruc.Text) <> "" Then
           txtCod_PlacaVehiculo.SetFocus
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
Private Sub txtCod_Almacen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        buscaAlmacen (1)
        If txtDes_Almacen.Text <> "" Then
          
          If indice = 0 Then
            txtCod_TipDoc.SetFocus
          Else
             txtNumeroGuiaExis.SetFocus
          End If
          
        Else
           txtCod_Almacen.SetFocus
        End If
        
    End If
End Sub
Private Sub txtDES_Almacen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        buscaAlmacen (2)
        
        If txtDes_Almacen.Text <> "" Then
           If indice = 0 Then
                txtCod_TipDoc.SetFocus
           Else
                txtNumeroGuiaExis.SetFocus
           End If
           
           
        Else
           txtCod_Almacen.SetFocus
        End If
        
    End If
End Sub
Public Sub buscaDocumentos(sopcion As String)
On Error GoTo fin
Dim iCol As Long
Dim rstAux As New ADODB.Recordset
Dim Opcion As String
  fila_seleccionada = 0
  strSQL = "CN_MUESTRA_VENTAS_CAJAS_DOCUMENTOS_PRENDAS  '" & sopcion & "','','','" & Trim(txtCod_TipDoc.Text) & "','" & Trim(txtDes_TipDoc.Text) & "'"
  With frmBusqGeneral
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos
        CODIGO = ".."
        Set rstAux = .gexList.ADORecordset
        
        .gexList.Columns("Cod_TipDoc").Caption = "Codigo"
        .gexList.Columns("Cod_TipDoc").Width = 1000
        .gexList.Columns("DES_TIPDOC").Caption = "Almacen"
        .gexList.Columns("DES_TIPDOC").Width = 4000
        
        If rstAux.RecordCount > 0 Then .Show vbModal
        If fila_seleccionada > 0 And rstAux.RecordCount > 0 Then
            rstAux.AbsolutePosition = fila_seleccionada
            txtCod_TipDoc.Text = Trim(rstAux!Cod_TipDoc)
            txtDes_TipDoc.Text = Trim(rstAux!DES_TIPDOC)
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
    MsgBox Err.Description & ",No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Documento(" & Opcion & ")"
End Sub
Public Sub buscaAlmacen(sopcion As String)
On Error GoTo fin
Dim iCol As Long
Dim rstAux As New ADODB.Recordset
Dim Opcion As String
  strSQL = "CN_MUESTRA_ALMACEN_GUIA_PRENDAS  '" & sopcion & "','" & Trim(txtCod_Almacen.Text) & "','" & Trim(txtDes_Almacen.Text) & "'"
  With frmBusqGeneral
        Set .oParent = Me
        .sQuery = strSQL
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
    MsgBox Err.Description & ",No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de almcen(" & Opcion & ")"
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
Private Sub txtNum_ruc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        Load FrmBusqClientesPrendas
        Set FrmBusqClientesPrendas.oParent = Me
        FrmBusqClientesPrendas.txtDescripcion_Cliente.Text = "" 'txtDes_TipAne.Text
        FrmBusqClientesPrendas.txtRuc_Cliente.Text = txtNum_ruc.Text
        FrmBusqClientesPrendas.txtTip_Anex.Text = "C"
        txtDes_TipAne.Text = ""
        
        Call FrmBusqClientesPrendas.Busca_Opcion_AnexoContable("1", "C", txtNum_ruc.Text, txtDes_TipAne.Text)
        FrmBusqClientesPrendas.Show 1
        'FrmBusqClientes.txtRuc_Cliente.SetFocus
        
        'txtDes_TipAne.Text = FrmBusqClientes.codigo
        'txtNum_Ruc.Text = FrmBusqClientes.Descripcion
      
       'If Trim(txtNum_Ruc.Text) <> "" Then
       '   txtCod_ConPag.SetFocus
       'Else
          txtNum_ruc.SetFocus
       'End If
        Set FrmBusqClientesPrendas = Nothing
  End If
 
End Sub
Private Sub txtNumeroGuiaExis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtNumeroGuiaExis.Text = Format(txtNumeroGuiaExis, "00000000")
        'Call MuestraDetalleGuiaExiste
        If indice = 1 Then
           cmdBuscarGuia.SetFocus
        End If
        
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
Private Sub txtSerieGuiaExis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtSerieGuiaExis.Text = Format(txtSerieGuiaExis, "000")
        'Call MuestraDetalleGuiaExiste
        txtNumeroGuiaExis.SetFocus
    End If
End Sub
'''******************************* MUESTRA DETALLE DE LA GUIA 0*******************************************
Private Sub MuestraDetalleGuiaExiste()
Dim RSAUX As ADODB.Recordset
Dim rslista As ADODB.Recordset
Dim I As Integer
On Error GoTo fin

If txtCod_Almacen.Text = "" Then
    Call MsgBox("sirvase ingresar un almacen", vbCritical, "Mensaje")
    txtCod_Almacen.SetFocus
    Exit Sub
End If

If txtSerieGuiaExis.Text = "" Then
    Call MsgBox("sirvase ingresar la serie de la guia", vbCritical, "Mensaje")
    txtSerieGuiaExis.SetFocus
    Exit Sub
End If

If txtNumeroGuiaExis.Text = "" Then
    Call MsgBox("sirvase ingresar el numero de la guia", vbCritical, "Mensaje")
    txtNumeroGuiaExis.SetFocus
    Exit Sub
End If

strSQL = "TX_MUESTRA_CABECERA_GUIA_prendas '" & txtCod_Almacen.Text & "','" & Trim(txtSerieGuiaExis.Text) & "','" & Trim(txtNumeroGuiaExis.Text) & "'"
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
txtNum_ruc.Text = rslista("NUM_RUC")
txtLug_Entrega.Text = rslista("LUG_ENTREGA")


'''' detalle de las guias
strSQL = "CN_MUESTRA_DETALLE_GUIA_REMISION_PRENDAS '" & Trim(txtCod_Almacen.Text) & "', '" & Trim(txtSerieGuiaExis.Text) & "','" & Trim(txtNumeroGuiaExis.Text) & "'"
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
MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
"Edicionar Producto "
    
End Sub

