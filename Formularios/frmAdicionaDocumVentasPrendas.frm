VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdicionaDocumVentasPrendas 
   Caption         =   "VENTAS DE PRENDAS Y OTROS"
   ClientHeight    =   8790
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   17415
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   17415
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraVendedor 
      Height          =   615
      Left            =   13000
      TabIndex        =   139
      Top             =   360
      Width           =   4455
      Begin VB.TextBox txtCod_Vendedor 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   915
         TabIndex        =   141
         Top             =   240
         Width           =   825
      End
      Begin VB.TextBox txtDes_Vendedor 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1680
         TabIndex        =   140
         Top             =   240
         Width           =   2625
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
         Left            =   120
         TabIndex        =   142
         Top             =   240
         Width           =   945
      End
   End
   Begin VB.Frame fraUbicacion 
      Height          =   615
      Left            =   10320
      TabIndex        =   14
      Top             =   360
      Width           =   2655
      Begin VB.TextBox txtDes_Almacen 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   960
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtCod_Almacen 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   480
         MaxLength       =   4
         TabIndex        =   15
         Top             =   240
         Width           =   465
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
         TabIndex        =   17
         Top             =   255
         Width           =   375
      End
   End
   Begin VB.Frame FraDatosCaja 
      ClipControls    =   0   'False
      Enabled         =   0   'False
      Height          =   615
      Left            =   0
      TabIndex        =   129
      Top             =   360
      Width           =   10280
      Begin VB.TextBox txtCod_Caja 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   7875
         TabIndex        =   137
         Top             =   240
         Width           =   465
      End
      Begin VB.TextBox txtDes_Caja 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   8280
         TabIndex        =   136
         Top             =   240
         Width           =   1905
      End
      Begin VB.TextBox txtCod_Tienda 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   4395
         TabIndex        =   134
         Top             =   240
         Width           =   465
      End
      Begin VB.TextBox txtDes_Tienda 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   4845
         TabIndex        =   133
         Top             =   240
         Width           =   2505
      End
      Begin VB.TextBox txtCod_Fabrica 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   885
         TabIndex        =   131
         Top             =   240
         Width           =   465
      End
      Begin VB.TextBox txtDes_Fabrica 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1365
         TabIndex        =   130
         Top             =   240
         Width           =   2265
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
         Left            =   7440
         TabIndex        =   138
         Top             =   240
         Width           =   405
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
         Left            =   3720
         TabIndex        =   135
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
         Left            =   0
         TabIndex        =   132
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame FraMediosPagos 
      Caption         =   "MEDIOS DE PAGO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6825
      Left            =   5160
      TabIndex        =   80
      Top             =   2040
      Width           =   12255
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   9000
         TabIndex        =   126
         Top             =   4680
         Width           =   3135
         Begin VB.OptionButton optTipo_Impresion 
            Caption         =   "ETIQUETERA"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   128
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optTipo_Impresion 
            Caption         =   "F. NORMAL"
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   127
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmdMedioPagoAgregar 
         Caption         =   "AGREGAR"
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
         Left            =   5640
         TabIndex        =   124
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text22 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   11160
         TabIndex        =   123
         Text            =   "S/."
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox Text21 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   9000
         TabIndex        =   122
         Text            =   "MONEDA NACIONAL"
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox txtMedioPagoImporteMN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   10560
         TabIndex        =   121
         Top             =   4200
         Width           =   1575
      End
      Begin VB.TextBox Text29 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   9000
         TabIndex        =   120
         Text            =   "IMPORTE"
         Top             =   4200
         Width           =   1575
      End
      Begin VB.TextBox txtMedioPagoIGVMN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   10560
         TabIndex        =   119
         Top             =   3840
         Width           =   1575
      End
      Begin VB.TextBox Text27 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   9000
         TabIndex        =   118
         Text            =   "I.G.V"
         Top             =   3840
         Width           =   1575
      End
      Begin VB.TextBox txtMedioPagoDsctoMN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   10560
         TabIndex        =   117
         Text            =   "0.00"
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox Text25 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   9000
         TabIndex        =   116
         Text            =   "DESCUENTO"
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox txtMedioPagoSubtotalMN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   10560
         TabIndex        =   115
         Top             =   3120
         Width           =   1575
      End
      Begin VB.TextBox Text23 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   9000
         TabIndex        =   114
         Text            =   "SUBTOTAL"
         Top             =   3120
         Width           =   1575
      End
      Begin VB.TextBox txtMedioPagoImporteME 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   10560
         TabIndex        =   113
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox Text19 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   9000
         TabIndex        =   112
         Text            =   "IMPORTE"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtMedioPagoIGVME 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   10560
         TabIndex        =   111
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox Text17 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   9000
         TabIndex        =   110
         Text            =   "I.G.V"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtMedioPagoDsctoME 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   10560
         TabIndex        =   109
         Text            =   "0.00"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   9000
         TabIndex        =   108
         Text            =   "DESCUENTO"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtMedioPagoSubtotalME 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   10560
         TabIndex        =   107
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   9000
         TabIndex        =   106
         Text            =   "SUBTOTAL"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   11160
         TabIndex        =   105
         Text            =   "US$"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   9000
         TabIndex        =   104
         Text            =   "MONEDA EXTRANJERA"
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtMedioPagoVueltoMN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   405
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   103
         Top             =   5770
         Width           =   2655
      End
      Begin VB.TextBox txtMedioPagoTotalPagoMN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   405
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   102
         Top             =   5350
         Width           =   2655
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   101
         Text            =   "VUELTO MN"
         Top             =   5770
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   100
         Text            =   "TOTAL PAGOS"
         Top             =   5350
         Width           =   1695
      End
      Begin VB.TextBox txtMedioPagoVueltoME 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   405
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   99
         Top             =   5770
         Width           =   2655
      End
      Begin VB.TextBox txtMedioPagoTotalPagoME 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   405
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   98
         Top             =   5350
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   97
         Text            =   "VUELTO ME"
         Top             =   5770
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   96
         Text            =   "TOTAL PAGOS"
         Top             =   5350
         Width           =   1695
      End
      Begin VB.TextBox txtMedioPagoDocumento 
         BackColor       =   &H00C0FFFF&
         Height          =   350
         Left            =   5760
         TabIndex        =   94
         Top             =   600
         Width           =   2415
      End
      Begin VB.ComboBox cboMedioPagoMoneda 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   5760
         Style           =   2  'Dropdown List
         TabIndex        =   92
         Top             =   240
         Width           =   2400
      End
      Begin VB.TextBox txtMedioPagoImporte 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   350
         Left            =   1680
         TabIndex        =   89
         Top             =   960
         Width           =   3135
      End
      Begin VB.OptionButton Option1 
         Caption         =   "TARJETAS DE CREDITO"
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
         Index           =   1
         Left            =   3240
         TabIndex        =   88
         Top             =   120
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "CONTADO/EFECTIVO"
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
         Index           =   0
         Left            =   1680
         TabIndex        =   87
         Top             =   120
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.ComboBox cboMedioPago 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   84
         Top             =   600
         Width           =   3120
      End
      Begin VB.CommandButton cmdImprimirDocumento 
         Caption         =   "IMPRIMIR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9360
         TabIndex        =   83
         Top             =   5520
         Width           =   1215
      End
      Begin VB.CommandButton cmdCerrarMediosPagos 
         Caption         =   "REGRESAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10560
         TabIndex        =   82
         Top             =   5520
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
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
         TabIndex        =   81
         Top             =   240
         Width           =   735
      End
      Begin GridEX20.GridEX grxMedioPagos 
         Height          =   3975
         Left            =   120
         TabIndex        =   85
         Top             =   1320
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   7011
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
         Column(1)       =   "frmAdicionaDocumVentasPrendas.frx":0000
         Column(2)       =   "frmAdicionaDocumVentasPrendas.frx":00C8
         FormatStylesCount=   8
         FormatStyle(1)  =   "frmAdicionaDocumVentasPrendas.frx":016C
         FormatStyle(2)  =   "frmAdicionaDocumVentasPrendas.frx":0294
         FormatStyle(3)  =   "frmAdicionaDocumVentasPrendas.frx":0344
         FormatStyle(4)  =   "frmAdicionaDocumVentasPrendas.frx":03F8
         FormatStyle(5)  =   "frmAdicionaDocumVentasPrendas.frx":04D0
         FormatStyle(6)  =   "frmAdicionaDocumVentasPrendas.frx":0588
         FormatStyle(7)  =   "frmAdicionaDocumVentasPrendas.frx":0668
         FormatStyle(8)  =   "frmAdicionaDocumVentasPrendas.frx":06F8
         ImageCount      =   0
         PrinterProperties=   "frmAdicionaDocumVentasPrendas.frx":080C
      End
      Begin VB.Label Label20 
         Caption         =   "DOCUM."
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
         Left            =   4920
         TabIndex        =   95
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "MONEDA"
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
         Left            =   4920
         TabIndex        =   93
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "IMPORTE DE PAGO"
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
         TabIndex        =   91
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label17 
         Caption         =   "MEDIOS DE PAGO"
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
         TabIndex        =   90
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label21 
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
         TabIndex        =   86
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame FraProductos 
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00000000&
      Height          =   5520
      Left            =   960
      TabIndex        =   54
      Top             =   2040
      Width           =   15015
      Begin VB.ComboBox cboTipoProducto 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   120
         Width           =   1320
      End
      Begin VB.TextBox txtCod_Ordpro_Bus 
         BackColor       =   &H00C0FFFF&
         Height          =   350
         Left            =   12120
         TabIndex        =   78
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox TxtCod_Estcli_Bus 
         BackColor       =   &H00C0FFFF&
         Height          =   350
         Left            =   2160
         TabIndex        =   62
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox txtDes_Present_Bus 
         BackColor       =   &H00C0FFFF&
         Height          =   350
         Left            =   7440
         TabIndex        =   61
         Top             =   120
         Width           =   2295
      End
      Begin VB.TextBox txtDes_Estcli_Bus 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   350
         Left            =   3600
         TabIndex        =   60
         Top             =   120
         Width           =   3255
      End
      Begin VB.TextBox txtCodigo_Barra_Bus 
         BackColor       =   &H00C0FFFF&
         Height          =   350
         Left            =   10320
         TabIndex        =   59
         Top             =   120
         Width           =   1575
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
         TabIndex        =   58
         Top             =   120
         Width           =   420
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
         TabIndex        =   57
         Top             =   5115
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
         Left            =   13920
         TabIndex        =   56
         Top             =   5115
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
         Left            =   14160
         TabIndex        =   55
         Top             =   240
         Width           =   735
      End
      Begin GridEX20.GridEX GrxProductos 
         Height          =   4575
         Left            =   45
         TabIndex        =   63
         Top             =   480
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   8070
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
         Column(1)       =   "frmAdicionaDocumVentasPrendas.frx":09E4
         Column(2)       =   "frmAdicionaDocumVentasPrendas.frx":0AAC
         FormatStylesCount=   8
         FormatStyle(1)  =   "frmAdicionaDocumVentasPrendas.frx":0B50
         FormatStyle(2)  =   "frmAdicionaDocumVentasPrendas.frx":0C78
         FormatStyle(3)  =   "frmAdicionaDocumVentasPrendas.frx":0D28
         FormatStyle(4)  =   "frmAdicionaDocumVentasPrendas.frx":0DDC
         FormatStyle(5)  =   "frmAdicionaDocumVentasPrendas.frx":0EB4
         FormatStyle(6)  =   "frmAdicionaDocumVentasPrendas.frx":0F6C
         FormatStyle(7)  =   "frmAdicionaDocumVentasPrendas.frx":104C
         FormatStyle(8)  =   "frmAdicionaDocumVentasPrendas.frx":10DC
         ImageCount      =   0
         PrinterProperties=   "frmAdicionaDocumVentasPrendas.frx":11F0
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
         TabIndex        =   77
         Top             =   240
         Width           =   255
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
         TabIndex        =   67
         Top             =   240
         Width           =   375
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
         TabIndex        =   66
         Top             =   240
         Width           =   615
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
         TabIndex        =   65
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
         Left            =   13200
         TabIndex        =   64
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame fraSelGuias 
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00000000&
      Height          =   5400
      Left            =   960
      TabIndex        =   0
      Top             =   2040
      Width           =   14535
      Begin VB.TextBox txtSerieGuia 
         BackColor       =   &H00C0FFFF&
         Height          =   350
         Left            =   5280
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
      Begin VB.ComboBox cboAlmacen 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   150
         Width           =   3975
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
         TabIndex        =   4
         Top             =   5000
         Width           =   975
      End
      Begin VB.TextBox txtNumeroGuia 
         BackColor       =   &H00C0FFFF&
         Height          =   350
         Left            =   6360
         TabIndex        =   3
         Top             =   120
         Width           =   2055
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
         TabIndex        =   2
         Top             =   5000
         Width           =   975
      End
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
         TabIndex        =   1
         Top             =   5000
         Width           =   975
      End
      Begin GridEX20.GridEX grxListaGuiaPendientes 
         Height          =   4455
         Left            =   45
         TabIndex        =   7
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
         Column(1)       =   "frmAdicionaDocumVentasPrendas.frx":13C8
         Column(2)       =   "frmAdicionaDocumVentasPrendas.frx":1490
         FormatStylesCount=   8
         FormatStyle(1)  =   "frmAdicionaDocumVentasPrendas.frx":1534
         FormatStyle(2)  =   "frmAdicionaDocumVentasPrendas.frx":165C
         FormatStyle(3)  =   "frmAdicionaDocumVentasPrendas.frx":170C
         FormatStyle(4)  =   "frmAdicionaDocumVentasPrendas.frx":17C0
         FormatStyle(5)  =   "frmAdicionaDocumVentasPrendas.frx":1898
         FormatStyle(6)  =   "frmAdicionaDocumVentasPrendas.frx":1950
         FormatStyle(7)  =   "frmAdicionaDocumVentasPrendas.frx":1A30
         FormatStyle(8)  =   "frmAdicionaDocumVentasPrendas.frx":1AC0
         ImageCount      =   0
         PrinterProperties=   "frmAdicionaDocumVentasPrendas.frx":1BD4
      End
      Begin GridEX20.GridEX grxListaGuiasSeleccionadas 
         Height          =   4455
         Left            =   8520
         TabIndex        =   8
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
         Column(1)       =   "frmAdicionaDocumVentasPrendas.frx":1DAC
         Column(2)       =   "frmAdicionaDocumVentasPrendas.frx":1E74
         FormatStylesCount=   8
         FormatStyle(1)  =   "frmAdicionaDocumVentasPrendas.frx":1F18
         FormatStyle(2)  =   "frmAdicionaDocumVentasPrendas.frx":2040
         FormatStyle(3)  =   "frmAdicionaDocumVentasPrendas.frx":20F0
         FormatStyle(4)  =   "frmAdicionaDocumVentasPrendas.frx":21A4
         FormatStyle(5)  =   "frmAdicionaDocumVentasPrendas.frx":227C
         FormatStyle(6)  =   "frmAdicionaDocumVentasPrendas.frx":2334
         FormatStyle(7)  =   "frmAdicionaDocumVentasPrendas.frx":2414
         FormatStyle(8)  =   "frmAdicionaDocumVentasPrendas.frx":24A4
         ImageCount      =   0
         PrinterProperties=   "frmAdicionaDocumVentasPrendas.frx":25B8
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
         TabIndex        =   11
         Top             =   240
         Width           =   255
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
         TabIndex        =   10
         Top             =   240
         Width           =   795
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
         TabIndex        =   9
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   0
      TabIndex        =   68
      Top             =   2040
      Width           =   17415
      Begin GridEX20.GridEX grxDatos 
         Height          =   5955
         Left            =   120
         TabIndex        =   69
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
         Column(1)       =   "frmAdicionaDocumVentasPrendas.frx":2790
         Column(2)       =   "frmAdicionaDocumVentasPrendas.frx":2858
         FormatStylesCount=   9
         FormatStyle(1)  =   "frmAdicionaDocumVentasPrendas.frx":28FC
         FormatStyle(2)  =   "frmAdicionaDocumVentasPrendas.frx":2A24
         FormatStyle(3)  =   "frmAdicionaDocumVentasPrendas.frx":2AD4
         FormatStyle(4)  =   "frmAdicionaDocumVentasPrendas.frx":2B88
         FormatStyle(5)  =   "frmAdicionaDocumVentasPrendas.frx":2C60
         FormatStyle(6)  =   "frmAdicionaDocumVentasPrendas.frx":2D18
         FormatStyle(7)  =   "frmAdicionaDocumVentasPrendas.frx":2DF8
         FormatStyle(8)  =   "frmAdicionaDocumVentasPrendas.frx":2E88
         FormatStyle(9)  =   "frmAdicionaDocumVentasPrendas.frx":2FC0
         ImageCount      =   0
         PrinterProperties=   "frmAdicionaDocumVentasPrendas.frx":30D4
      End
   End
   Begin VB.TextBox txt_descto 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   11085
      TabIndex        =   53
      Top             =   8400
      Width           =   975
   End
   Begin VB.TextBox txt_subtotal 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   12960
      TabIndex        =   52
      Top             =   8400
      Width           =   1215
   End
   Begin VB.TextBox txt_igv 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   14520
      TabIndex        =   51
      Top             =   8400
      Width           =   1095
   End
   Begin VB.TextBox txt_total 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   16200
      TabIndex        =   50
      Top             =   8400
      Width           =   1215
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
      TabIndex        =   49
      Text            =   "R E G I S T R O    D E   V E N T A S"
      Top             =   0
      Width           =   17415
   End
   Begin VB.Frame frMain 
      Height          =   1080
      Left            =   0
      TabIndex        =   18
      Top             =   960
      Width           =   17415
      Begin VB.CheckBox Check2 
         Caption         =   "VENTAS VARIOS"
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
         Left            =   6480
         TabIndex        =   125
         Top             =   760
         Width           =   1695
      End
      Begin VB.TextBox txtNum_Ruc 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   6480
         MaxLength       =   11
         TabIndex        =   36
         Top             =   420
         Width           =   4220
      End
      Begin VB.TextBox txtCod_TipVenta 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   11640
         MaxLength       =   4
         TabIndex        =   35
         Top             =   120
         Width           =   600
      End
      Begin VB.TextBox txtDes_TipVenta 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   12240
         TabIndex        =   32
         Top             =   120
         Width           =   3855
      End
      Begin VB.TextBox txtSer_Docum 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   4770
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   31
         Top             =   120
         Width           =   1080
      End
      Begin VB.TextBox txtCod_TipDoc 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1005
         MaxLength       =   4
         TabIndex        =   30
         Top             =   120
         Width           =   465
      End
      Begin VB.TextBox txtDes_TipDoc 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1485
         TabIndex        =   29
         Top             =   120
         Width           =   2625
      End
      Begin VB.TextBox txtNum_Docum 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   5850
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   28
         Top             =   120
         Width           =   2020
      End
      Begin VB.TextBox txtDes_TipAne 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1485
         TabIndex        =   27
         Top             =   420
         Width           =   4425
      End
      Begin VB.TextBox txtCod_Moneda 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   8445
         MaxLength       =   4
         TabIndex        =   26
         Top             =   120
         Width           =   600
      End
      Begin VB.TextBox txtDes_Moneda 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   9040
         TabIndex        =   25
         Top             =   120
         Width           =   1650
      End
      Begin VB.TextBox txtCod_ConPag 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1005
         MaxLength       =   4
         TabIndex        =   24
         Top             =   705
         Width           =   465
      End
      Begin VB.TextBox txtDes_ConPag 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1485
         TabIndex        =   23
         Top             =   705
         Width           =   4425
      End
      Begin VB.TextBox txtCod_TipAne 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1005
         MaxLength       =   4
         TabIndex        =   22
         Text            =   "C"
         Top             =   420
         Width           =   465
      End
      Begin VB.Frame frReferencia 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         TabIndex        =   21
         Top             =   5040
         Visible         =   0   'False
         Width           =   7815
      End
      Begin VB.TextBox TxtTipo_Cambio 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   11640
         TabIndex        =   20
         Top             =   705
         Width           =   855
      End
      Begin VB.TextBox txtiva 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   12960
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   705
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpFec_Emision 
         Height          =   285
         Left            =   11640
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
         Format          =   82182145
         CurrentDate     =   38182
      End
      Begin MSComCtl2.DTPicker dtpFec_Registro 
         Height          =   285
         Left            =   14760
         TabIndex        =   34
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
         Format          =   82182145
         CurrentDate     =   38182
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
         TabIndex        =   48
         Top             =   120
         Width           =   285
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
         TabIndex        =   47
         Top             =   120
         Width           =   975
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
      Begin VB.Label Label5 
         Caption         =   "Numero :"
         Height          =   255
         Left            =   9390
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
         Top             =   420
         Width           =   375
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
         TabIndex        =   42
         Top             =   405
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
         Left            =   10920
         TabIndex        =   41
         Top             =   405
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
         TabIndex        =   40
         Top             =   120
         Width           =   495
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
         TabIndex        =   39
         Top             =   720
         Width           =   735
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
         TabIndex        =   38
         Top             =   735
         Width           =   375
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
         TabIndex        =   37
         Top             =   735
         Width           =   375
      End
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
      TabIndex        =   13
      Top             =   8400
      Width           =   3375
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
      TabIndex        =   12
      Top             =   8520
      Width           =   1455
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   6720
      TabIndex        =   70
      Top             =   8280
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   900
      Custom          =   $"frmAdicionaDocumVentasPrendas.frx":32AC
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
      TabIndex        =   71
      Top             =   8280
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   900
      Custom          =   $"frmAdicionaDocumVentasPrendas.frx":334D
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   12
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   12120
      Top             =   8640
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
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
      Left            =   10605
      TabIndex        =   76
      Top             =   8520
      Width           =   615
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
      TabIndex        =   75
      Top             =   8520
      Width           =   855
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
      TabIndex        =   74
      Top             =   8520
      Width           =   375
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
      TabIndex        =   73
      Top             =   8520
      Width           =   735
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
      TabIndex        =   72
      Top             =   8400
      Width           =   615
   End
End
Attribute VB_Name = "frmAdicionaDocumVentasPrendas"
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
Dim Contador As Double
Private estado_caja  As String
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

Private Sub cboMedioPago_Change()
    txtMedioPagoDocumento.Enabled = False
    If Right(cboMedioPago, 1) = "S" Then
      txtMedioPagoDocumento.Enabled = True
    End If
End Sub
Private Sub cboMedioPago_Click()
    txtMedioPagoDocumento.Enabled = False
    If Right(cboMedioPago, 1) = "S" Then
      txtMedioPagoDocumento.Enabled = True
    End If
End Sub

Private Sub Check2_Click()
Call muestraventasvarias
End Sub

Private Sub chkTodos_Click()
On Error GoTo fin
    If GrxProductos.RowCount = 0 Then Exit Sub
    Dim Rs As New ADODB.Recordset
    Dim Valor As Boolean
    Dim i As Long

    GrxProductos.Update
    Set Rs = GrxProductos.ADORecordset
    Rs.MoveFirst
    Do While Not Rs.EOF
        If chkTodos.Value = Checked Then
            If Rs("stock") > 0 Then
                Rs("cant") = Rs("stock")
                Rs("total") = Rs("stock") * Rs("precio")
            End If
        Else
                Rs("cant") = 0
        End If
                Rs.MoveNext
    Loop
    Rs.MoveFirst
    Rs.Update
    
    Set GrxProductos.ADORecordset = Rs
    Call ConfiguraGrilla_productos
    
Exit Sub
Resume
fin:
On Error Resume Next
Set Rs = Nothing
MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
"Edicionar Producto "

End Sub

Private Sub CmdCerrarGuias_Click()
fraSelGuias.Visible = False
    flg_Tiene_guias_asignadas = "N"
If DevuelveCampo(" select count(*) from lg_movistk  where ser_docum_ventas<>'' AND  ser_docum_ventas='" & Trim(txtSer_Docum.Text) & "' AND num_docum_ventas <>'' and num_docum_ventas='" & Trim(txtNum_Docum.Text) & "' ", cConnect) > 0 Then
    flg_Tiene_guias_asignadas = "S"
End If
Call adicionarProductoDesdeDetalleGuia

End Sub

Private Sub cmdCerrarMediosPagos_Click()
      
   habilitaframe (True)
   'FraMediosPagos.Visible = False
End Sub

Private Sub cmdDesAsigna_Click()
On Error GoTo fin
If grxListaGuiasSeleccionadas.RowCount = 0 Then Exit Sub
StrSQL = "CN_ASIGNA_GUIA_FACTURA_PRENDAS '" & grxListaGuiasSeleccionadas.Value(grxListaGuiasSeleccionadas.Columns("cod_almacen").Index) & "','" & grxListaGuiasSeleccionadas.Value(grxListaGuiasSeleccionadas.Columns("num_movstk").Index) & "','',''"
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

    StrSQL = "CN_ASIGNA_GUIA_FACTURA_PRENDAS '" & grxListaGuiaPendientes.Value(grxListaGuiaPendientes.Columns("cod_almacen").Index) & "','" & grxListaGuiaPendientes.Value(grxListaGuiaPendientes.Columns("num_movstk").Index) & "','" & Trim(txtSer_Docum.Text) & "','" & Trim(txtNum_Docum.Text) & "'"
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

Private Sub cmdImprimirDocumento_Click()
  If grxDatos.RowCount <= 0 Then Exit Sub
         If validaDatosIniciales = True And validaImporteMedioPago = True Then
             
             If MsgBox("¡¡¡Esta apunto de confirmar en caja el documento de venta!!!:" & Chr(13) & Chr(10) & ":::::> " & Trim(txtDes_TipDoc.Text) & " " & txtSer_Docum & "-" & txtNum_Docum & Chr(13) & Chr(10) & "¿Son los datos correctos?", vbYesNo, "CONFIRMAR") = vbYes Then
              If flg_Tiene_guias_asignadas = "N" Then
                If GuardaDetalleVentas = True Then
                    Call obtieneDatosIniciales
                    Call estadoInicialVentana
                    Call buscaDetalle_factura
                    Call obtieneDatosInicialesMediosPago
                    Call iniciofraMedioPago
                    FraMediosPagos.Visible = False
                    Frame1.Enabled = True
                End If
              End If
              If flg_Tiene_guias_asignadas = "S" Then
                If GuardaDetalleVentasDesdeDetalleGuiaPrenas = True Then
                    Call obtieneDatosIniciales
                    Call estadoInicialVentana
                    Call buscaDetalle_factura
                    Call obtieneDatosInicialesMediosPago
                    Call iniciofraMedioPago
                    FraMediosPagos.Visible = False
                    Frame1.Enabled = True
                End If
              End If
              
         End If
   End If
End Sub

Private Sub cmdMedioPagoAgregar_Click()

If Not IsNumeric(txtMedioPagoImporte.Text) Then
 Call MsgBox("Ingrese una Cantidad Valida", vbCritical, "Mensaje")
 Call iniciofraMedioPago
 Exit Sub
End If
AdicionaMedioPago
End Sub

Private Sub cmdVentasVarios_Click()
Call muestraventasvarias
End Sub
Private Sub muestraventasvarias()
On Error GoTo fin
Dim sTit As String
Dim strCadena As String
sTit = "Ventas Varios"
Dim rsventasvarios As New ADODB.Recordset

If (Contador Mod 2) > 0 Then

    strCadena = "CN_MUESTRA_DATOS_VENTAS_VARIOS '','" & Trim(txtCod_Tienda.Text) & "','" & Trim(txtCod_Caja.Text) & "'"
    Set rsventasvarios = CargarRecordSetDesconectado(strCadena, cConnect)
        
    txtCod_TipDoc.Text = Trim(rsventasvarios!Cod_TipDoc)
    txtDes_TipDoc.Text = Trim(rsventasvarios!Des_TipDoc)
    txtSer_Docum.Text = Trim(rsventasvarios!COR_DOCSERIE)
    txtNum_Docum.Text = Trim(rsventasvarios!COR_NUMACTU)

    txtNum_ruc.Text = Trim(rsventasvarios!Num_Ruc)
    txtDes_TipAne.Text = Trim(rsventasvarios!Des_Anexo)
    txtNum_ruc.Tag = Trim(rsventasvarios!Cod_Anxo)
    txtDes_TipAne.Tag = Trim(rsventasvarios!cod_cliente)
    
    txtCod_TipVenta.Text = Trim(rsventasvarios!Cod_Tipo_Venta)
    txtDes_TipVenta.Text = Trim(rsventasvarios!DESCRIPCION_TIPO_VENTA)
    txtCod_ConPag.Text = Trim(rsventasvarios!Cod_CondVent)
    txtDes_ConPag.Text = Trim(rsventasvarios!Des_CondVent)
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
    txtCod_TipVenta.Text = ""
    txtDes_TipVenta.Text = ""
    txtCod_ConPag.Text = ""
    txtDes_ConPag.Text = ""
    txtCod_Moneda.Text = ""
    txtDes_Moneda.Text = ""

End If
Contador = Contador + 1
Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, sTit
End Sub

Private Sub dtpFec_Emision_Change()
    txtiva.Text = DevuelveCampo("SELECT PORC_IGV  FROM TG_IGV where ano= '" & Year(dtpFec_Emision) & "' and mes= '" & Format(Month(dtpFec_Emision), "00") & "'", cConnect)
    TxtTipo_Cambio.Text = DevuelveCampo("select isnull(Tipo_Venta,0) from cn_tipocambio where fecha = '" & dtpFec_Emision & "'", cConnect)
    dtpFec_Registro = dtpFec_Emision
End Sub

Private Sub Form_Load()
On Error GoTo fin
   
StrSQL = "EXEC SM_MUESTRA_ESTADO_CN_VENTAS_CAJAS_FECHA '" & ComputerName & "','" & usuario_windows & "'"
estado_caja = DevuelveCampo(StrSQL, cConnect)

          Contador = 1
          If Not IsNumeric(txtiva.Text) Then
           txtiva.Text = 0
          End If
          If Not IsNumeric(TxtTipo_Cambio) Then
            TxtTipo_Cambio.Text = 0
          End If
          iva = 1 + (txtiva.Text / 100#)
          Call DisableCloseButton(Me)
          flg_Tiene_guias_asignadas = "N"
          FraProductos.Visible = False
          fraSelGuias.Visible = False
          dtpFec_Emision.Value = Date
          dtpFec_Registro.Value = Date
          Call buscaDetalle_factura
          Call obtieneDatosIniciales
          Call FillAlmacen
          Call FillTipoProducto
          Call FillMedioPago(0)
          Call FillMoneda
          indiceMedioPago = 0
          indiceTipo_Impresion = 0
          Call obtieneDatosInicialesMediosPago
          FraMediosPagos.Visible = False
          
          'txtCod_TipDoc.SetFocus
          txtiva.Text = DevuelveCampo("SELECT PORC_IGV  FROM TG_IGV where ano= '" & Year(dtpFec_Emision) & "' and mes= '" & Format(Month(dtpFec_Emision), "00") & "'", cConnect)
          TxtTipo_Cambio.Text = DevuelveCampo("select isnull(Tipo_Venta,0) from cn_tipocambio where fecha = '" & dtpFec_Emision & "'", cConnect)
        
          dtpFec_Emision.Enabled = False
          dtpFec_Registro.Enabled = False
           
          StrSQL = "select dbo.SS_REVISA_PERMISO('" & vusu & "','EDIT_FEC_REGIS_VENTA')"
          If DevuelveCampo(StrSQL, cConnect) = 1 Then
             dtpFec_Emision.Enabled = True
             dtpFec_Registro.Enabled = True
          End If
          
          If CDbl(txtiva.Text) <= 0 Then
              Call MsgBox("Ingrese el porcentaje del impuesto sobre el valor agregado (iva) ", vbCritical, "Importante")
              'Unload Me
          End If
          
          iva = 1 + (txtiva.Text / 100#)
                  
          If Not IsNumeric(TxtTipo_Cambio.Text) Then
            TxtTipo_Cambio.Text = 0
          End If
          txtMedioPagoDocumento.Enabled = False
          
          If CDbl(TxtTipo_Cambio.Text) <= 0 Then
              Call MsgBox("Ingrese el Tipo Cambio Para la fecha", vbCritical, "Importante")
              'Unload Me
          End If
          
        If estado_caja = "A" Then
                  Check2.Value = 1
        Else
        
            FraDatosCaja.Enabled = False
            fraUbicacion.Enabled = False
            frMain.Enabled = False
            Frame1.Enabled = False
            FraVendedor.Enabled = False
            FunctButt2.Visible = False
            txtCodigo_Producto.Enabled = False
        
            If estado_caja = "P" Then
                MsgBox "¡¡¡ADVERTENCIA!!!" & Chr(13) & "¡...No se puede realizar ventas, la caja no se ha aperturado...!" & Chr(13) & " Utilice la opcion-->ventas/apertura Caja", vbCritical + vbOKOnly, "IMPORTANTE"
            ElseIf estado_caja = "C" Then
                MsgBox "¡¡¡ADVERTENCIA!!!" & Chr(13) & "¡...No se puede realizar ventas, la caja se encuentra cerrada...!" & Chr(13) & " Utilice la opcion -->ventas/cierre Caja/ deshacercierre", vbCritical + vbOKOnly, "IMPORTANTE"
            Else
                MsgBox "¡¡¡ADVERTENCIA!!!" & Chr(13) & "¡...No se puede realizar ventas, el estado de la caja no lo permite...!", vbCritical + vbOKOnly, "IMPORTANTE"
            End If
        
        End If
        
Exit Sub
fin:
MsgBox "No se Puede Continuar", err.Description + vbInformation + vbOKOnly, "IMPORTANTE"

End Sub
Private Sub FillTipoProducto()
On Error GoTo fin
Dim sTit As String
    
    sTit = "carga tipo Producto"
    StrSQL = " LG_MUESTRA_TIPO_PRODUCTO "
    
    Set rstAux = CargarRecordSetDesconectado(StrSQL, cConnect)
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
    MsgBox err.Description, vbCritical + vbOKOnly, sTit
End Sub
Private Sub FillMedioPago(tipo_pago As Integer)
On Error GoTo fin
Dim sTit As String
Dim tipPago As String
    tipPago = "T"
    If tipo_pago = 0 Then
        tipPago = "E"
    End If
    sTit = "Carga Medios Pagos"
    StrSQL = " CN_MUESTRA_MEDIOS_PAGO '" & tipPago & "'"
    Set rstAux = CargarRecordSetDesconectado(StrSQL, cConnect)
    cboMedioPago.Clear
    With rstAux
        If .RecordCount > 0 Then .MoveFirst
        Do Until .EOF
            cboMedioPago.AddItem !cod_medpag & " " & !des_medpago & Space(50) & !FLG_EXIGEDOC
            .MoveNext
        Loop
        .Close
    End With
    If cboMedioPago.ListCount > 0 Then cboMedioPago.ListIndex = 0
    Set rstAux = Nothing
Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, sTit
End Sub
Private Sub FillMoneda()
On Error GoTo fin
Dim sTit As String
    sTit = "CARGA MONEDAS"
    StrSQL = " CN_MUESTRA_MONEDAS "
    Set rstAux = CargarRecordSetDesconectado(StrSQL, cConnect)
    cboMedioPagoMoneda.Clear
    With rstAux
        If .RecordCount > 0 Then .MoveFirst
        Do Until .EOF
            cboMedioPagoMoneda.AddItem !Moneda
            .MoveNext
        Loop
        .Close
    End With
    If cboMedioPagoMoneda.ListCount > 0 Then cboMedioPagoMoneda.ListIndex = 0
    Set rstAux = Nothing
Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, sTit
End Sub

Private Sub FillAlmacen()
On Error GoTo fin
Dim sTit As String
    
    sTit = "Cargar Almacenes"
    StrSQL = " LG_MUESTRA_ALMACENES_PRODUCTOS_TERMINADOS  '" & vusu & "'"
    
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
StrSQL = "CN_MUESTRA_CAJAS_VENDEDOR_ACCESO '" & pc & "','" & usuario_windows & "'"
 Set auxset = Nothing
 Set auxset = CargarRecordSetDesconectado(StrSQL, cConnect)
 If auxset.RecordCount > 0 Then
    Txtcod_Fabrica.Text = auxset("cod_Fabrica")
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

Private Sub obtieneDatosInicialesMediosPago()
    Dim StrSQL As String
    StrSQL = "CN_UPMAN_CN_VENTAS_MEDIO_PAGO 'V','','','','',0,0,0"
    Set grxMedioPagos.ADORecordset = Nothing
    Set grxMedioPagos.ADORecordset = CargarRecordSetDesconectado(StrSQL, cConnect)
    Call ConfiguraGrillaMedioPago

End Sub
Private Sub habilitaframe(Valor As Boolean)
    FraDatosCaja.Enabled = Valor
    fraUbicacion.Enabled = Valor
    frMain.Enabled = Valor
    Frame1.Enabled = Valor
    FraMediosPagos.Visible = Not Valor
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo dprDepurar
Select Case ActionName
Case Is = "GRABAR"
  '''cartado
  If grxDatos.RowCount > 0 Then
   habilitaframe (False)
   'FraMediosPagos.Visible = True
   Call iniciofraMedioPago
  End If
Case Is = "CANCELAR"
 If grxDatos.RowCount > 0 Then
  If MsgBox("¡...Al cancelar esta operacion se eliminaran los datos registrados...! " & Chr(13) & Chr(10) & " ¿Esta Seguro de proseguir? ", vbYesNo, "CONFIRMAR") = vbYes Then
    If flg_Tiene_guias_asignadas = "S" Then
      Call EliminaGuiasAsigandas
      End If
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
errores err.Number
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
    MsgBox err.Description, vbCritical + vbOKOnly, sTit
End Function

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

If txtNum_ruc.Tag <> "0001" Then
    txtDes_TipAne.Text = ""
    txtNum_ruc.Text = ""
    txtDes_TipAne.Tag = ""
    txtNum_ruc.Tag = ""
End If

txtNum_Docum.Text = DevuelveCampo("SELECT COR_NUMACTU FROM CN_VENTAS_CAJAS_DOCUMENTOS WHERE COD_FABRICA='" & Txtcod_Fabrica.Text & "' AND  COD_TIENDA='" & Trim(txtCod_Tienda.Text) & "' AND COD_CAJA='" & txtCod_Caja.Text & "' AND COD_TIPDOC='" & Trim(txtCod_TipDoc.Text) & "' AND COR_DOCSERIE ='" & txtSer_Docum.Text & "' ", cConnect)

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
    StrSQL = " VENTAS_UP_MAN_PRENDAS_OTROS 'I','','" & Txtcod_Fabrica.Text & "','" & Trim(txtCod_Tienda.Text) & "','" & Trim(txtCod_Caja.Text) & "','" & Trim(txtCod_Vendedor.Text) & "','" & Trim(txtCod_Almacen.Text) & "','" & txtCod_TipDoc & "','" & txtSer_Docum & "','" _
            & txtNum_Docum & "','C','" & Trim(txtNum_ruc.Tag) & "','" & txtCod_ConPag & "','" & txtCod_TipVenta.Text & "','" & Format(dtpFec_Emision, "dd/mm/yyyy") & "','" _
            & Format(dtpFec_Registro, "dd/mm/yyyy") & "','" & txtCod_Moneda & "','" _
            & vusu & "',''," _
            & TxtTipo_Cambio.Text & ",'','','N','N','S'," & txtMedioPagoTotalPagoMN.Text & "," & txtMedioPagoVueltoMN.Text & ""
            
    Set rstAux = cntAux.Execute(StrSQL, adExecuteNoRecords)
    strNum_Corre = rstAux!Num_Corre
    rstAux.Close

    '''CABECERA MOVIMIENTO
    StrSQL = "EXEC TI_UP_MAN_LG_MOVISTK_PRENDAS_OTROS 'I', '" & _
             Trim(txtCod_Almacen.Text) & "', '', '" & Format(dtpFec_Registro.Value, _
             "dd/mm/yyyy") & "', '' ,'SV7','', '" & txtDes_TipAne.Tag & "','','" & vusu & "','',''"

    Set rstAux = cntAux.Execute(StrSQL, adExecuteNoRecords)
    sNum_MovStk = rstAux!num_movstk
    rstAux.Close
    
    Set rstAux = grxDatos.ADORecordset
    With rstAux
        If .RecordCount > 0 Then .MoveFirst
        Do Until .EOF
    '''DETALLE MOVIMIENTO DE SALIDA DE ALMACEN
             StrSQL = "EXEC LG_UP_MAN_LG_MOVISTKITEM_PRENDAS_OTROS 'I', '" & _
             Trim(txtCod_Almacen.Text) & "','" & sNum_MovStk & "','" & Now() & "', '', '" & _
             Trim(!COD_ITEM) & "','" & Trim(!Cod_Comb) & "','" & Trim(!cod_Color) & "','" & Trim(!cod_estcli) & "','" & Trim(!cod_ordpro) & "'," & Trim(!cod_present) & ",'" & Trim(!cod_talla) & "','" & Trim(!codigo_barra) & "', '" & _
             Trim(!tipo_producto) & "'," & !cant & ",'" & vusu & "'"
             cntAux.Execute StrSQL, adExecuteNoRecords
 
    '''DETALLE VENTAS falta strCod_Anxo
            StrSQL = "CN_VENTAS_ITEMS_PRENTAS_OTROS 'I','" & strNum_Corre & "','','" & Trim(!COD_ITEM) & "','" & Trim(!Cod_Comb) & "','" & Trim(!cod_Color) & "','" & _
            Trim(!cod_cliente) & "','" & Trim(!cod_purord) & "','" & Trim(!cod_lotpurord) & "','" & Trim(!cod_colcli) & "','" & Trim(!cod_estcli) & "','" & Trim(!cod_ordpro) & "'," _
            & Trim(!cod_present) & ",'" & Trim(!cod_talla) & "','" & Trim(!codigo_barra) & "','" & !tipo_producto & "'," & !cant & "," & !precio & ", " & !Total & " ,'" & _
            Trim(!des_estcli) & "','" & Trim(!des_present) & "','" & Trim(!Des_Comb) & "',0,'','',0,'" & vusu & "'"
            cntAux.Execute StrSQL, adExecuteNoRecords
            .MoveNext

        Loop
    End With
    
    '''ASOCIA FACTURA CON MOVIMIENTO DE ALMACEN
    StrSQL = "CN_VENTAS_CAJAS_RELACIONA_FACTURA_GUIA_PRENDAS 'U','" & strNum_Corre & "','" & Trim(txtSer_Docum.Text) & "','" & Trim(txtNum_Docum.Text) & "','" & Trim(txtCod_Almacen.Text) & "','" & sNum_MovStk & "'"
    cntAux.Execute StrSQL, adExecuteNoRecords

    '''GUARDA LOS IMPORTES DE LOS MEDIOS DE PAGO
    Set rstAux = Nothing
    grxMedioPagos.Update
    Set rstAux = grxMedioPagos.ADORecordset
    rstAux.Update
    
    With rstAux
    .MoveFirst
    Do Until .EOF
        
        StrSQL = "CN_UPMAN_CN_VENTAS_MEDIO_PAGO 'I','" & strNum_Corre & "','" & !cod_medpag & "','" & !Cod_Moneda & "','" & Trim(!DOC_MEDPAG) & "'," & !IMP_MEDPAGO & "," & !TIP_CAMBIO & ",'" & !IMP_TOTALMEDPAG & "'"
        cntAux.Execute StrSQL, adExecuteNoRecords
        
    .MoveNext
    Loop
    End With
    
    cntAux.CommitTrans
    cntAux.Close
    Set cntAux = Nothing
    
    '''IMPRIME DOCUMENTO
'    If indiceTipo_Impresion = 0 Then
'        'Call imprimeTicket(strNum_Corre, "000")
'        Call imprimebixolon270(strNum_Corre, "000")
'    Else
'        Call Preliminar_Docum_Ventas(strNum_Corre)
'    End If
    
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

Private Function GuardaDetalleVentasDesdeDetalleGuiaPrenas() As Boolean
On Error GoTo ErrDetMov
Dim sErr As String, cntAux As New ADODB.Connection, sTit As String, _
    sNum_MovStk As String, strNum_Corre  As String
Dim Kilos_Tenidos As Double
Dim RollosTeñidos As Double
Dim rstAux As New ADODB.Recordset

  GuardaDetalleVentasDesdeDetalleGuiaPrenas = False

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
    StrSQL = " VENTAS_UP_MAN_PRENDAS_OTROS 'I','','" & Txtcod_Fabrica.Text & "','" & Trim(txtCod_Tienda.Text) & "','" & Trim(txtCod_Caja.Text) & "','" & Trim(txtCod_Vendedor.Text) & "','" & Trim(txtCod_Almacen.Text) & "','" & txtCod_TipDoc & "','" & txtSer_Docum & "','" _
            & txtNum_Docum & "','C','" & Trim(txtNum_ruc.Tag) & "','" & txtCod_ConPag & "','" & txtCod_TipVenta.Text & "','" & Format(dtpFec_Emision.Value, "dd/mm/yyyy") & "','" _
            & Format(dtpFec_Registro.Value, "dd/mm/yyyy") & "','" & txtCod_Moneda & "','" _
            & vusu & "',''," _
            & TxtTipo_Cambio.Text & ",'','','N','N','N'," & txtMedioPagoTotalPagoMN.Text & "," & txtMedioPagoVueltoMN.Text & ""
            
    Set rstAux = cntAux.Execute(StrSQL, adExecuteNoRecords)
    strNum_Corre = rstAux!Num_Corre
    rstAux.Close

'''CABECERA MOVIMIENTO
'    StrSql = "EXEC TI_UP_MAN_LG_MOVISTK_PRENDAS_OTROS 'I', '" & _
'             Trim(txtCod_Almacen.Text) & "', '', '" & Format(dtpFec_Registro.Value, _
'             "dd/mm/yyyy") & "', '' ,'SV7','', '" & txtDes_TipAne.Tag & "','','" & vusu & "','',''"
'
'    Set rstAux = cntAux.Execute(StrSql, adExecuteNoRecords)
'    sNum_MovStk = rstAux!num_movstk
'    rstAux.Close
'
    Set rstAux = grxDatos.ADORecordset
    With rstAux
        If .RecordCount > 0 Then .MoveFirst
        Do Until .EOF
    '''DETALLE MOVIMIENTO DE SALIDA DE ALMACEN
'             StrSql = "EXEC LG_UP_MAN_LG_MOVISTKITEM_PRENDAS_OTROS 'I', '" & _
'             Trim(txtCod_Almacen.Text) & "','" & sNum_MovStk & "','" & Now() & "', '', '" & _
'             Trim(!COD_ITEM) & "','" & Trim(!cod_comb) & "','" & Trim(!cod_Color) & "','" & Trim(!cod_estcli) & "','" & Trim(!cod_ordpro) & "'," & Trim(!cod_present) & ",'" & Trim(!cod_talla) & "','" & Trim(!codigo_barra) & "', '" & _
'             Trim(!tipo_producto) & "'," & !cant & ",'" & vusu & "'"
'             cntAux.Execute StrSql, adExecuteNoRecords
 
    '''DETALLE VENTAS falta strCod_Anxo
            StrSQL = "CN_VENTAS_ITEMS_PRENTAS_OTROS 'I','" & strNum_Corre & "','','" & Trim(!COD_ITEM) & "','" & Trim(!Cod_Comb) & "','" & Trim(!cod_Color) & "','" & _
            Trim(!cod_cliente) & "','" & Trim(!cod_purord) & "','" & Trim(!cod_lotpurord) & "','" & Trim(!cod_colcli) & "','" & Trim(!cod_estcli) & "','" & Trim(!cod_ordpro) & "'," _
            & Trim(!cod_present) & ",'" & Trim(!cod_talla) & "','" & Trim(!codigo_barra) & "','" & !tipo_producto & "'," & !cant & "," & !precio & ", " & !Total & " ,'" & _
            Trim(!des_estcli) & "','" & Trim(!des_present) & "','" & Trim(!Des_Comb) & "',0,'','',0,'" & vusu & "'"
            cntAux.Execute StrSQL, adExecuteNoRecords
            .MoveNext
        Loop
    End With
    
    '''ASOCIA FACTURA CON MOVIMIENTO DE ALMACEN
    StrSQL = "CN_VENTAS_CAJAS_RELACIONA_FACTURA_GUIA_PRENDAS 'U','" & strNum_Corre & "','" & Trim(txtSer_Docum.Text) & "','" & Trim(txtNum_Docum.Text) & "','" & Trim(txtCod_Almacen.Text) & "','" & sNum_MovStk & "'"
    cntAux.Execute StrSQL, adExecuteNoRecords

    '''GUARDA LOS IMPORTES DE LOS MEDIOS DE PAGO
    Set rstAux = Nothing
    grxMedioPagos.Update
    Set rstAux = grxMedioPagos.ADORecordset
    rstAux.Update
    
    With rstAux
    .MoveFirst
    Do Until .EOF
        
        StrSQL = "CN_UPMAN_CN_VENTAS_MEDIO_PAGO 'I','" & strNum_Corre & "','" & !cod_medpag & "','" & !Cod_Moneda & "','" & Trim(!DOC_MEDPAG) & "'," & !IMP_MEDPAGO & "," & !TIP_CAMBIO & ",'" & !IMP_TOTALMEDPAG & "'"
        cntAux.Execute StrSQL, adExecuteNoRecords
        
    .MoveNext
    Loop
    End With
    
    cntAux.CommitTrans
    cntAux.Close
    Set cntAux = Nothing
    
    '''IMPRIME DOCUMENTO
    If indiceTipo_Impresion = 0 Then
        'Call imprimeTicket(strNum_Corre, "000")
        Call imprimebixolon270(strNum_Corre, "000")
    Else
        Call Preliminar_Docum_Ventas(strNum_Corre)
    End If
    GuardaDetalleVentasDesdeDetalleGuiaPrenas = True
    
Exit Function
ErrDetMov:
    GuardaDetalleVentasDesdeDetalleGuiaPrenas = False
    sErr = err.Description
    cntAux.RollbackTrans
    cntAux.Close
    Set cntAux = Nothing
    MsgBox sErr, vbCritical + vbOKOnly, sTit
End Function
Private Sub Preliminar_Docum_Ventas(Num_Corre As String)
On Error GoTo SALTO_ERROR

Dim sSQL As String, Rs As New ADODB.Recordset
Dim imp_total As Double
Dim aMess(4), i As Integer

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
        StrSQL = "CN_MUESTRA_IMPRESION_DOCUMENTO_PRENDA_OTROS '" & lvNumCorre & "','" & UCase(EnLetras(Trim(CStr(dbImp_Total)))) & "'"
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
        
        StrSQL = "CN_MUESTRA_IMPRESION_DOCUMENTO_PRENDA_OTROS '" & lvNumCorre & "','" & UCase(EnLetras(Trim(CStr(txt_total)))) & "'"
        Set rsFactura = CargarRecordSetDesconectado(StrSQL, cConnect)
        Call Factura_sa("BV", Serie)
    Case Is = "TK"
        Call imprimeTicket(lvNumCorre, Serie)
        
    Case Else
      MsgBox "No se ha Definido un Formato de Impresion para este tipo de documento", vbInformation, "ERROR"
       Imprimir_FACTURA = False
      Exit Function
    End Select
End With
Imprimir_FACTURA = True
End Function
'''*****************************************************************************************************************************
''' imprime tk para bixolon 270srp, impresion directa
''' con esta forma de impresion, en el ancho entra 33 caracteres mas o menos por que depende del ancho  de la letra (ejm i y la w)
''' para este tipo de letra los tamaños de las columnas para cada letra varia.
''' para letras como courier el ancho de letra es fija
''' cuando se agrega el driver de la bixolon la impresora debe tener el nombre "BIXOLON SRP-270", no se debe cambiar
'''*****************************************************************************************************************************
Private Sub imprimebixolon270(lvNumCorre As String, Serie As String)
Dim rxsTique As ADODB.Recordset
Dim rxaux As ADODB.Recordset
 Dim SImp_Total As Double
 Dim simp_igv As Double
 Dim simp_neto As Double
 Dim simp_efectivo As Double
 Dim simp_vuelto As Double
 Dim linea1 As String
 Dim linea2 As String
 Dim linea3 As String
 Dim linea4 As String
 Dim linea5 As String
 Dim linea6 As String
 Dim simp_letras As String
 Dim nro_articulos As Double
Dim StrSQL As String
Dim Slinea  As String
Dim oPrint As clsPrintFile
Dim E As Integer
Dim snum_autorizacion_sunat As String
Dim spacio_totales_7  As String
Dim C As Integer
Dim usuario As String
Dim nom_impresora As String
usuario = usuario_windows

StrSQL = "CN_MUESTRA_IMPRESION_DOCUMENTO_PRENDA_OTROS '" & lvNumCorre & "','" & UCase(EnLetras(Trim(CStr(txt_total)))) & "'"
Set rxsTique = CargarRecordSetDesconectado(StrSQL, cConnect)
Set rxaux = rxsTique

If rxsTique.RecordCount = 0 Then
    Call MsgBox("Documento No Tiene Detalle", vbCritical, "Mensaje")
    Exit Sub
End If

For Each prnPrinter In Printers

C = InStr(prnPrinter.DeviceName, "(") - 1
If C = -1 Then
 C = Len(prnPrinter.DeviceName)
End If
nom_impresora = prnPrinter.DeviceName
If (UCase(Trim(Mid(prnPrinter.DeviceName, 1, C))) = UCase("BIXOLON SRP-270-MASTERSEVEN-PC") And UCase(usuario) = "FACTIENDA") Then
    
    Set Printer = prnPrinter
    
    Printer.FontName = "FontB2x2[Ext.]"
    Printer.FontSize = 15
    Printer.FontBold = True
    Printer.Print Spc(3); rxsTique!NOM_FABRICA
    
    Printer.FontName = "FontA1x1[Ext.]"
    Printer.FontSize = 9
    Printer.Print Spc(10); "RUC:" + rxsTique!Num_Ruc_FABRICA
    
    Printer.FontName = "FontA1x1[Ext.]"
    Printer.FontSize = 9
    Printer.Print rxsTique!DIRECCION
    
    Printer.FontName = "FontA1x1[Ext.]"
    Printer.FontSize = 9
    Printer.Print Spc(4); "GALERIA EL REY DE GAMARRA"
    
    Printer.FontName = "FontA1x1[Ext.]"
    Printer.FontSize = 9
    Printer.Print Spc(8); "LA VICTORIA-LIMA-PERU"

    Printer.FontName = "FontA1x1[Ext.]"
    Printer.FontSize = 9
    Printer.Print Spc(7); "TELEFONO.: (13)" & rxsTique!TELEFONO

    E = 36 - Len(Trim(rxsTique!Des_TipDoc))
    E = E / 2

    Printer.FontName = "FontA1x1[Ext.]"
    Printer.FontSize = 9
    Printer.Print Spc(E); UCase(Trim(rxsTique!Des_TipDoc))

    Printer.FontName = "Arial"
    Printer.FontSize = 9
    Printer.Print " "

    Printer.FontName = "FontA1x1"
    Printer.FontSize = 13
    Printer.FontBold = True
    Printer.Print Space(14) + Trim(rxsTique!ser_docum) + "-" + Trim(rxsTique!num_docum_ventas)
    
    Printer.FontName = "Arial"
    Printer.FontSize = 9
    Printer.Print " "
    '''--------------------datos
    
    Printer.FontName = "FontA1x1[Ext.]"
    Printer.FontSize = 9
    Printer.Print "FECHA DE EMISION: " & Format(Now(), "DD/MM/YYYY")
    Printer.Print "HORA            : " & Format(Now(), "HH:MM:SS")
    Printer.Print "LOCAL           : " & rxsTique!DES_TIENDA
    Printer.Print "CAJA No         : " & rxsTique!cod_caja
    Printer.Print "TRANSACCION No  : " & rxsTique!Num_Corre
    Printer.Print "TIPO MONEDA     : " & rxsTique!Nom_Moneda
    Printer.Print "VENDEDOR        : " & Mid(Trim(rxsTique!VENDEDOR), 1, 13)
    Printer.Print "S/M             : " & rxsTique!SERIE_MAQUINA

    '''--------------------lineas del final
'''Totales
    SImp_Total = rxsTique!imp_total
    simp_igv = rxsTique!imp_igv
    simp_neto = rxsTique!imp_neto
    simp_efectivo = rxsTique!EFECTIVO
    simp_vuelto = rxsTique!VUELTO
    simp_letras = rxsTique!IMPORTE_LETRAS
    nro_articulos = rxsTique!nro_articulos
    
    linea1 = rxsTique!lineafinal1
    linea2 = rxsTique!lineafinal2
    linea3 = rxsTique!lineafinal3
    linea4 = rxsTique!lineafinal4
    linea5 = rxsTique!lineafinal5
    linea6 = rxsTique!lineafinal6
    snum_autorizacion_sunat = rxsTique!NUM_AUTORIZACION_SUNAT

    spacio_totales_7 = "       "
    
    Printer.Print "---------------------------------"
    Printer.Print "CODIGO  DESCRIPCION              "
    Printer.Print "          CANT     PRECIO  TOTAL "
    Printer.Print "---------------------------------"
    
    rxsTique.Update
    Dim i As Integer
    i = 1
    rxsTique.MoveFirst
    Do While i <= rxsTique.RecordCount
    
        Slinea = Mid(Trim(rxsTique!COD_ITEM), 1, 10) + "  " + Mid(Trim(rxsTique!Descripcion), 1, 23) '+ " " + Format(Trim(rxsTique!cantidad), "###.00") + " " + Format(Trim(rxsTique!IMP_UNITARIO), "###.00") + " " + Format(Trim(rxsTique!TOTAL_PAR), "####.00")
        Printer.Print Slinea
        Slinea = Space(7) + Right(spacio_totales_7 + Format(Trim(rxsTique!cantidad), "###.00"), 7) + " X  " + Right(spacio_totales_7 + Format(Trim(rxsTique!IMP_UNITARIO), "###.00"), 7) + " " + Right(spacio_totales_7 + Format(Trim(rxsTique!TOTAL_PAR), "####.00"), 7)
        Printer.Print Slinea
        rxsTique.MoveNext
        i = i + 1
    Loop

    Printer.Print ""
    Slinea = Space(7) + "SUB TOTAL          " & Right(spacio_totales_7 + Format(simp_neto, "####.00"), 7)
    Printer.Print Slinea
    
    Slinea = Space(7) + "IGV                " & Right(spacio_totales_7 + Format(simp_igv, "####.00"), 7)
    Printer.Print Slinea
    
    Slinea = Space(7) + "TOTAL              " & Right(spacio_totales_7 + Format(SImp_Total, "####.00"), 7)
    Printer.Print Slinea
        
    'Slinea = simp_letras
    'Printer.Print Slinea
    
    Slinea = Space(7) + "EFECTIVO           " & Right(spacio_totales_7 + Format(simp_efectivo, "####.00"), 7)
    Printer.Print Slinea
    
    Slinea = Space(7) + "VUELTO             " & Right(spacio_totales_7 + Format(simp_vuelto, "####.00"), 7)
    Printer.Print Slinea
    
    Printer.Print "---------------------------------"
    
    'Slinea = linea1
    'Printer.Print Slinea
    
    Slinea = linea2
    'Printer.Print Slinea
    
    Slinea = linea3
    'Printer.Print Slinea
    
    Slinea = linea1
    'Printer.Print Slinea
    
    Slinea = linea4
    Printer.Print Slinea
    
    'Slinea = linea5
    'Printer.Print Slinea
    Slinea = linea6
    Printer.Print Slinea
    
    Printer.Print "*********************************"
    Printer.Print "AUTORIZADO MEDIANTE RESOLUCION "
    Slinea = "NRO. " + Trim(snum_autorizacion_sunat) + "/SUNAT"
    Printer.Print Slinea
    Printer.Print "*********************************"

    '''lineas de codigo para cortar el ticket en la bixolon
    Printer.FontSize = 9
    Printer.FontName = "FontControl"
    Printer.Print "G"

    'Use special-function character to cut the paper
    'P: Partial cut
    'g: Partial cut without paper feeding
    
    Printer.EndDoc
    
    Exit For
End If

Next

End Sub

'''*****************************************************************************************************************************
''' imprime un archivo txt en tamaño tk para bixolon 270srp
'''*****************************************************************************************************************************
Private Sub imprimeTicket(lvNumCorre As String, Serie As String)
Dim rxsTique As ADODB.Recordset
Dim rxaux As ADODB.Recordset
 Dim SImp_Total As Double
 Dim simp_igv As Double
 Dim simp_neto As Double
 Dim simp_efectivo As Double
 Dim simp_vuelto As Double
 Dim linea1 As String
 Dim linea2 As String
 Dim linea3 As String
 Dim linea4 As String
 Dim linea5 As String
 Dim linea6 As String
 Dim simp_letras As String
 Dim nro_articulos As Double
Dim StrSQL As String
Dim Slinea  As String
Dim oPrint As clsPrintFile
Dim E As Integer
Dim snum_autorizacion_sunat As String
Dim spacio_totales_7  As String

StrSQL = "CN_MUESTRA_IMPRESION_DOCUMENTO_PRENDA_OTROS '" & lvNumCorre & "','" & UCase(EnLetras(Trim(CStr(txt_total)))) & "'"
Set rxsTique = CargarRecordSetDesconectado(StrSQL, cConnect)
Set rxaux = rxsTique

If rxsTique.RecordCount = 0 Then
    Call MsgBox("Documento No Tiene Detalle", vbCritical, "Mensaje")
End If

Set oPrint = New clsPrintFile

Close #1
Open "c:\DOCUMENTOVENTA_TK.txt" For Output As #1
    
Plin Chr(15)
Slinea = rxsTique!DES_FABRICA
Plin Slinea
Slinea = Space(12) + rxsTique!NOM_FABRICA
Plin Slinea
Slinea = Space(12) + "RUC:" + rxsTique!Num_Ruc_FABRICA
Plin Slinea
Slinea = rxsTique!DIRECCION
Plin Slinea
Slinea = "GALERIA EL REY DE GAMARRA-LA VICTORIA" 'rxsTique!RES_FABRICA
Plin Slinea
Slinea = Space(15) + "LIMA-PERU " 'rxsTique!RES_FABRICA
Plin Slinea
Slinea = Space(12) + "TELEFONO.: " & rxsTique!TELEFONO
Plin Slinea

E = 38 - Len(Trim(rxsTique!Des_TipDoc))
E = E / 2
Slinea = Space(E) + Trim(rxsTique!Des_TipDoc) '"TICKET N°" & rxsTique!ser_docum + " " + rxsTique!num_docum_ventas
Plin Slinea
Plin " "

Slinea = Space(14) + Trim(rxsTique!ser_docum) + "-" + Trim(rxsTique!num_docum_ventas)
Plin Slinea
Plin " "
''------------------DATOS
Slinea = "FECHA DE EMISION : " & Format(Now(), "DD/MM/YYYY")
Plin Slinea
Slinea = "HORA             : " & Format(Now(), "HH:MM:SS")
Plin Slinea
Slinea = "LOCAL            : " & rxsTique!DES_TIENDA
Plin Slinea

Slinea = "CAJA No          : " & rxsTique!cod_caja
Plin Slinea

Slinea = "TRANSACCION No   : " & rxsTique!Num_Corre
Plin Slinea

Slinea = "TIPO MONEDA      : " & rxsTique!Nom_Moneda
Plin Slinea

Slinea = "VENDEDOR         : " & Mid(Trim(rxsTique!VENDEDOR), 1, 19)
Plin Slinea

'Slinea = "N/A              : " & rxsTique!NUM_AUTORIZACION_SUNAT
'Plin Slinea

Slinea = "S/M              : " & rxsTique!SERIE_MAQUINA
Plin Slinea


'''Totales
SImp_Total = rxsTique!imp_total
simp_igv = rxsTique!imp_igv
simp_neto = rxsTique!imp_neto
simp_efectivo = rxsTique!EFECTIVO
simp_vuelto = rxsTique!VUELTO
simp_letras = rxsTique!IMPORTE_LETRAS
nro_articulos = rxsTique!nro_articulos

linea1 = rxsTique!lineafinal1
linea2 = rxsTique!lineafinal2
linea3 = rxsTique!lineafinal3
linea4 = rxsTique!lineafinal4
linea5 = rxsTique!lineafinal5
linea6 = rxsTique!lineafinal6
snum_autorizacion_sunat = rxsTique!NUM_AUTORIZACION_SUNAT

'Trim(rs.Fields("COD_FAMITEM").Value)
Slinea = " "
Slinea = " "
spacio_totales_7 = Space(7)

Slinea = "----------------------------------------"
Plin Slinea
Slinea = "CODIGO  DESCRIPCION  CANT  PRECIO TOTAL"
Plin Slinea
Slinea = "----------------------------------------"
Plin Slinea

rxsTique.Update
Dim i As Integer
i = 1
rxsTique.MoveFirst
Do While i <= rxsTique.RecordCount

    Slinea = Mid(Trim(rxsTique!COD_ITEM), 1, 10) + "  " + Mid(Trim(rxsTique!Descripcion), 1, 28) '+ " " + Format(Trim(rxsTique!cantidad), "###.00") + " " + Format(Trim(rxsTique!IMP_UNITARIO), "###.00") + " " + Format(Trim(rxsTique!TOTAL_PAR), "####.00")
    Plin Slinea
    Slinea = Space(17) + Right(spacio_totales_7 + Format(Trim(rxsTique!cantidad), "###.00"), 7) + " " + Right(spacio_totales_7 + Format(Trim(rxsTique!IMP_UNITARIO), "###.00"), 7) + " " + Right(spacio_totales_7 + Format(Trim(rxsTique!TOTAL_PAR), "####.00"), 7)
    Plin Slinea
    rxsTique.MoveNext
    i = i + 1
    
Loop

'rxaux.Update
'rxaux.MoveFirst
'000  0000  0000.00

Plin ""
Slinea = Space(15) + "TOTAL         S/. " & Right(spacio_totales_7 + Format(simp_neto, "####.00"), 7)
Plin Slinea

Slinea = Space(15) + "IGV           S/. " & Right(spacio_totales_7 + Format(simp_igv, "####.00"), 7)
Plin Slinea

Slinea = Space(15) + "TOTAL A PAGAR S/. " & Right(spacio_totales_7 + Format(SImp_Total, "####.00"), 7)
Plin Slinea

Slinea = simp_letras
Plin Slinea

Slinea = Space(15) + "EFECTIVO      S/. " & Right(spacio_totales_7 + Format(simp_efectivo, "####.00"), 7)
Plin Slinea

Slinea = Space(15) + "VUELTO        S/. " & Right(spacio_totales_7 + Format(simp_vuelto, "####.00"), 7)
Plin Slinea

Plin "----------------------------------------"
Slinea = linea1
Plin Slinea

Slinea = linea2
Plin Slinea

Slinea = linea3
Plin Slinea

Slinea = linea4
Plin Slinea

Slinea = linea5
Plin Slinea

Slinea = linea6
Plin Slinea
Plin "****************************************"
Plin "AUTORIZADO MEDIANTE RESOLUCION NRO."
Slinea = Trim(snum_autorizacion_sunat) + "/SUNAT"
Plin Slinea
Plin "****************************************"

Close #1
oPrint.SendPrint "c:\DOCUMENTOVENTA_TK.txt"
Set oPrint = Nothing

Call CortaBixolon270

End Sub
Private Sub CortaBixolon270()

For Each prnPrinter In Printers
    If prnPrinter.DeviceName = "BIXOLON SRP-270" Then
        Set Printer = prnPrinter
        
        Printer.FontSize = 7
        Printer.FontName = "FontControl"
        Printer.Print "G"
        Printer.EndDoc
        Exit For
    End If
Next

End Sub

Sub Factura_sa(Tipo As String, Serie As String)
On Error GoTo ErrorImpresion
Dim oo As Object, lvSql As String, lvRuta As String

    Set oo = CreateObject("excel.application")
    
    If Tipo = "FA" Then
            oo.workbooks.Open vRuta & "\Factura_prendas_otros.XLT"
    End If

'    If Tipo = "ND" Then
'        oo.Workbooks.Open vRuta & "\Abono_Textil.XLT"
'    End If
'    If Tipo = "NC" Then
'        oo.Workbooks.Open vRuta & "\Credito_Textil.XLT"
'    End If

    If Tipo = "BV" Then
        oo.workbooks.Open vRuta & "\Boleta_prendas_otros.XLT"
    End If
    
    oo.DisplayAlerts = False
    
    If chkImpresionDirecta.Value = 1 Then
        oo.Visible = False
    Else
        oo.Visible = True
    End If
            
    oo.run "Reporte", rsFactura, IIf(chkImpresionDirecta.Value = 1, 1, 0), cConnect
 
    If chkImpresionDirecta.Value = 1 Then
        oo.workbooks.Close
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
Private Sub grxMedioPagos_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
  Select Case ColIndex
    Case Is = grxMedioPagos.Columns("ELI").Index
      Cancel = False
    Case Else
      Cancel = True
  End Select
End Sub

Private Sub GrxProductos_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
  Select Case ColIndex
    Case Is = GrxProductos.Columns("CANT").Index
      Cancel = False
    Case Else
      Cancel = True
  End Select
End Sub
Private Sub grxDatos_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
  Select Case ColIndex
    Case Is = grxDatos.Columns("CANT").Index
      Cancel = False
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
    Dim i As Integer
On Error GoTo fin
    
    If fraUbicacion.Enabled = False Then
    
        If Trim(Txtcod_Fabrica.Text) = "" Then
           Call MsgBox("Ingrese Una Empresa valida", vbCritical, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
        
        If Trim(txtCod_Tienda.Text) = "" Then
           Call MsgBox("Ingrese una Tienda valida", vbCritical, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
        
        If Trim(txtCod_Caja.Text) = "" Then
           Call MsgBox("Ingrese una caja valida", vbCritical, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
        
        If Trim(txtCod_Vendedor.Text) = "" Then
           Call MsgBox("el codigo del vendedor no es valido", vbCritical, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
        
        If Trim(txtCod_Almacen.Text) = "" Then
           Call MsgBox("El Codigo del Almacen no es valido", vbCritical, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
       
        If DevuelveCampo("SELECT COUNT(*) FROM CN_VENTAS_CAJAS_ALMACEN WHERE COD_FABRICA='" & Trim(Txtcod_Fabrica.Text) & "'", cConnect) <= 0 Then
           Call MsgBox("El Codigo Empresa no Valida", vbCritical, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
         
        If DevuelveCampo("SELECT COUNT(*) FROM CN_VENTAS_CAJAS_ALMACEN WHERE COD_FABRICA='" & Trim(Txtcod_Fabrica.Text) & "' and  cod_tienda='" & Trim(txtCod_Tienda.Text) & "'", cConnect) <= 0 Then
           Call MsgBox("El Codigo de Tienda no valida", vbCritical, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
         
        If DevuelveCampo("SELECT COUNT(*) FROM CN_VENTAS_CAJAS_ALMACEN WHERE COD_FABRICA='" & Trim(Txtcod_Fabrica.Text) & "' and  cod_tienda='" & Trim(txtCod_Tienda.Text) & "'and cod_caja = '" & Trim(txtCod_Caja.Text) & "'", cConnect) <= 0 Then
           Call MsgBox("El Codigo de caja no es valido ", vbCritical, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
         
        If DevuelveCampo("SELECT COUNT(*) FROM CN_VENTAS_CAJAS_ALMACEN WHERE COD_FABRICA='" & Trim(Txtcod_Fabrica.Text) & "' and  cod_tienda='" & Trim(txtCod_Tienda.Text) & "'and cod_caja = '" & Trim(txtCod_Caja.Text) & "' and cod_almacen= '" & Trim(txtCod_Almacen.Text) & "'", cConnect) <= 0 Then
           Call MsgBox("El Codigo de Caja no es valido ", vbCritical, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
                
        If TxtTipo_Cambio.Text = "" Or CDbl(TxtTipo_Cambio.Text) = 0 Then
           Call MsgBox("Ingrese El Tipo Cambio", vbCritical, "Mensaje")
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
         
        If txtCod_TipVenta.Text = "" Then
           Call MsgBox("Sirvase a Ingresar un tipo venta Valido ", vbCritical, "Mensaje")
           validaDatosIniciales = False
           Exit Function
           
        End If

        If txtNum_ruc.Text = "" Or txtDes_TipAne.Text = "" Then
           Call MsgBox("Sirvase a Ingresar un cliente Valido", vbCritical, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If

        If txtCod_ConPag.Text = "" Then
           Call MsgBox("Sirvase a Ingresar una condicion de venta", vbCritical, "Mensaje")
           validaDatosIniciales = False
           Exit Function
        End If
        
        If cantidadValida = False Then
           Call MsgBox("Sirvase a ingresar una Cantidad Valida ", vbCritical, "Mensaje")
           validaDatosIniciales = False
           Exit Function
           
        End If
    End If
    
Exit Function
fin:
MsgBox err.Description & ", No se puede Continuar, se presento un inconveniente", vbExclamation + vbOKOnly, "Edicionar Medios de Pago "
   
End Function
Private Function validaImporteMedioPago() As Boolean
On Error GoTo fin

validaImporteMedioPago = True

        If grxMedioPagos.RowCount <= 0 Then
           Call MsgBox("No se ha Registrado Ningun Pago", vbCritical, "Mensaje")
           validaImporteMedioPago = False
           Exit Function
        Else
        
        '' valida el total pagado al final
        Dim rxaux As New ADODB.Recordset
        Dim totalval As Double
        grxMedioPagos.Update
        Set rxaux = grxMedioPagos.ADORecordset
        rxaux.Update

        With rxaux
            .MoveFirst
            Do Until .EOF
                totalval = totalval + !IMP_TOTALMEDPAG
                .MoveNext
            Loop
            If totalval < CDbl(txt_total.Text) Then
               Call MsgBox("El Importe de los pagos es Menor que el Total", vbCritical, "Mensaje")
               validaImporteMedioPago = False
               Exit Function
            End If
        End With

       End If

Exit Function
fin:
MsgBox err.Description & ", No se puede Continuar, se presento un inconveniente", vbExclamation + vbOKOnly, "Edicionar Medios de Pago "
End Function

'''******************************* ADICIONA LISTA ARTICULOS CUYA CANTIDAD SEA MAYOR A 0*******************************************
Private Sub adicionarProductoMasivo()
Dim RSAUX As ADODB.Recordset
Dim rslista As ADODB.Recordset
Dim i As Integer
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
i = 1
Do While i <= rslista.RecordCount
If rslista!cant > 0 Then

    RSAUX.AddNew
    RSAUX!cod_cliente = rslista!cod_cliente
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
    RSAUX!Cod_Comb = rslista!Cod_Comb
    RSAUX!tipo_producto = rslista!tipo_producto
    RSAUX!COD_ITEM = rslista!COD_ITEM
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
''''******************************* ADICIONA LISTA ARTICULOS DESDE EL DETALLE DE LA GUIA 0*******************************************
Private Sub adicionarProductoDesdeDetalleGuia()
Dim RSAUX As ADODB.Recordset
Dim rslista As ADODB.Recordset
Dim i As Integer
On Error GoTo fin
'''' volvemos a llenar el detalle
Call buscaDetalle_factura
Set RSAUX = grxDatos.ADORecordset

'''' detalle de las guias
StrSQL = "CN_MUESTRA_DETALLE_GUIA_VENTA_PRENDAS '" & Trim(txtSer_Docum.Text) & "','" & Trim(txtNum_Docum) & "'"
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
    RSAUX!cod_cliente = rslista!cod_cliente
    RSAUX!nom_cliente = rslista!nom_cliente
    RSAUX!cod_temcli = rslista!cod_temcli
    RSAUX!nom_temcli = rslista!nom_temcli
    RSAUX!cod_purord = rslista!cod_purord
    
    RSAUX!cod_colcli = rslista!cod_colcli
    RSAUX!cod_lotpurord = rslista!cod_lotpurord
    
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
    RSAUX!Cod_Comb = rslista!Cod_Comb
    RSAUX!tipo_producto = rslista!tipo_producto
    RSAUX!COD_ITEM = rslista!COD_ITEM
    
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
'On Error Resume Next
Set RSAUX = Nothing
MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
"Edicionar Producto "
    
End Sub
'''******************************* ADICIONA LISTA ARTICULOS *******************************************
Private Sub adicionarProducto()
Dim RSAUX As ADODB.Recordset
On Error GoTo fin

'Set RSAUX = grxDatos.ADORecordset
'RSAUX.AddNew
'
'RSAUX!OT = IIf(IsNull(GrxProductos.Value(GrxProductos.Columns("OT").Index)), "", GrxProductos.Value(GrxProductos.Columns("OT").Index))
'RSAUX!codigoRollo = IIf(IsNull(GrxProductos.Value(GrxProductos.Columns("codigorollo").Index)), "", GrxProductos.Value(GrxProductos.Columns("codigorollo").Index))
'RSAUX!cod_tela = IIf(IsNull(GrxProductos.Value(GrxProductos.Columns("cod_tela").Index)), "", GrxProductos.Value(GrxProductos.Columns("cod_tela").Index))
'RSAUX!TELA = IIf(IsNull(GrxProductos.Value(GrxProductos.Columns("TELA").Index)), "", GrxProductos.Value(GrxProductos.Columns("TELA").Index))
'RSAUX!cod_color = IIf(IsNull(GrxProductos.Value(GrxProductos.Columns("cod_Color").Index)), "", GrxProductos.Value(GrxProductos.Columns("cod_color").Index))
'RSAUX!color = IIf(IsNull(GrxProductos.Value(GrxProductos.Columns("COLOR").Index)), "", GrxProductos.Value(GrxProductos.Columns("COLOR").Index))
'RSAUX!calidad = IIf(IsNull(GrxProductos.Value(GrxProductos.Columns("calidad").Index)), "", GrxProductos.Value(GrxProductos.Columns("calidad").Index))
'RSAUX!rollos = IIf(IsNull(GrxProductos.Value(GrxProductos.Columns("rollos").Index)), "", GrxProductos.Value(GrxProductos.Columns("rollos").Index))
'RSAUX!und = IIf(IsNull(GrxProductos.Value(GrxProductos.Columns("und").Index)), "", GrxProductos.Value(GrxProductos.Columns("und").Index))
'RSAUX!cant = IIf(IsNull(GrxProductos.Value(GrxProductos.Columns("cant").Index)), "", GrxProductos.Value(GrxProductos.Columns("cant").Index))
'RSAUX!Stock = IIf(IsNull(GrxProductos.Value(GrxProductos.Columns("stock").Index)), "", GrxProductos.Value(GrxProductos.Columns("stock").Index))
'RSAUX!precio = IIf(IsNull(GrxProductos.Value(GrxProductos.Columns("PRECIO").Index)), "", GrxProductos.Value(GrxProductos.Columns("PRECIO").Index))
'RSAUX!DEL = "X"
'RSAUX!Total = RSAUX!precio * RSAUX!cant
'RSAUX.Update
'Set grxDatos.ADORecordset = RSAUX

'''Call Total_documento
'''Call ConfiguraGrilla_Detalle

Exit Sub
Resume
fin:
On Error Resume Next
Set RSAUX = Nothing
MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
"Edicionar Producto "
    
End Sub

Private Sub Option1_Click(Index As Integer)
indiceMedioPago = Index

Call FillMedioPago(indiceMedioPago)

End Sub

Private Sub optTipo_Impresion_Click(Index As Integer)
indiceTipo_Impresion = Index

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
  If Len(Trim(txtCodigo_Producto.Text)) = 13 And flg_Tiene_guias_asignadas = "N" Then
    Call AdicionaProductoDirecto(1)
    txtCodigo_Producto.Text = ""
    txtCodigo_Producto.SetFocus
    'SendKeys "{TAB}"
  End If
End Sub
Private Sub AdicionaProductoDirecto(Opcion As String)

    Dim StrSQL As String
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
'''codigo barra de prenda cod_ordpro/cod_present/cod_talla--321450001000M
StrSQL = "EXEC CF_MUESTRA_PRENDAS_MOV_TIENDA_ITEMS_FACTURA        '" & Opcion & _
                                                    "','" & Trim(txtCod_Almacen.Text) & _
                                                    "','','','" & Trim(txtDes_Present_Bus.Text) & _
                                                    "','','" & Trim(TxtCod_Estcli_Bus.Text) & _
                                                    "','" & Trim(txtDes_Estcli_Bus.Text) & _
                                                    "','" & Trim(txtCod_Ordpro_Bus.Text) & _
                                                    "','','" & Trim(txtCodigo_Producto.Text) & "'"
                                                    
Else
'''codigo barra de otros articulo cod_item/correlativo--pr00000100001
        StrSQL = "EXEC CF_MUESTRA_ITEMS_TIENDA_FACTURA        '" & Opcion & _
                                                    "','" & Trim(txtCod_Almacen.Text) & _
                                                    "','" & Trim(TxtCod_Estcli_Bus.Text) & _
                                                    "','" & Trim(txtDes_Estcli_Bus.Text) & _
                                                    "','" & Trim(txtDes_Present_Bus.Text) & _
                                                    "','" & Trim(txtCodigo_Barra_Bus.Text) & "'"

End If
    'strSQL = "TX_MUESTRA_ROLLOS_VENTA '" & Opcion & "','" & Trim(txtCod_Almacen.Text) & "','" & Trim(txtCodigo_Producto.Text) & "','" & Trim(txtBus_Cod_ordtra.Text) & "','" & Trim(txtDescripcion_Producto.Text) & "','" & Trim(txtBus_Des_Color.Text) & "'"
    Set rsetbusqueda = Nothing
    Set rsetbusqueda = CargarRecordSetDesconectado(StrSQL, cConnect)
    If rsetbusqueda.RecordCount <= 0 Then
       Call MsgBox("Articulo No existe o no hay Stocks", vbCritical, "Mensaje")
       Exit Sub
    End If
    Set rsetAux = grxDatos.ADORecordset
 
    rsetAux.AddNew
    rsetAux!cod_cliente = rsetbusqueda!cod_cliente
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
    rsetAux!Cod_Comb = rsetbusqueda!Cod_Comb
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
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
Private Sub iniciofraMedioPago()

If grxMedioPagos.RowCount = 0 Then
    txtMedioPagoImporte.Text = txt_total.Text
    txtMedioPagoTotalPagoME.Text = Format(0, "##0.00")
    txtMedioPagoVueltoME.Text = Format(0, "##0.00")
    txtMedioPagoTotalPagoMN.Text = Format(0, "##0.00")
    txtMedioPagoVueltoMN.Text = Format(0, "##0.00")
Else
    Call TotalPagosMonedas
    Call ConfiguraGrillaMedioPago
End If

End Sub
Private Function validaAddMediosPago() As Boolean
On Error GoTo fin
     Dim rxRecord As New ADODB.Recordset
     Dim i As Integer
     Dim Total As Double
     validaAddMediosPago = True
     If grxMedioPagos.RowCount = 0 Then Exit Function
     
     Set rxRecord = Nothing
     grxMedioPagos.Update
     Set rxRecord = grxMedioPagos.ADORecordset
     rxRecord.Update
     
     i = 1
    rxRecord.MoveFirst
    Do While i <= rxRecord.RecordCount
        Total = Total + rxRecord!IMP_TOTALMEDPAG
        If Total >= txt_total.Text Then
            Call MsgBox("el Importe Pagado ya es mayor", vbCritical, "Mensaje")
            validaAddMediosPago = False
        End If
        i = i + 1
      rxRecord.MoveNext
    Loop
     
     Set rxRecord = Nothing
     grxMedioPagos.Update
     Set rxRecord = grxMedioPagos.ADORecordset
     rxRecord.Update
     
     i = 1
    rxRecord.MoveFirst
    Do While i <= rxRecord.RecordCount
      If Left(cboMedioPago, 2) = rxRecord!cod_medpag And Right(cboMedioPagoMoneda, 3) = rxRecord!Cod_Moneda Then
        Call MsgBox("Medio de pago ya existe", vbCritical, "Mensaje")
        validaAddMediosPago = False
        Exit Do
      End If
      i = i + 1
      rxRecord.MoveNext
    Loop
    
    Exit Function
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Function

Private Sub AdicionaMedioPago()
    Dim StrSQL As String
    Dim tipocambio As Double
    Dim rxRecosertMedPag As ADODB.Recordset
    On Error GoTo fin

    txtMedioPagoTotalPagoMN.Text = 0
    txtMedioPagoVueltoMN.Text = 0
    txtMedioPagoTotalPagoME.Text = 0
    txtMedioPagoVueltoME.Text = 0
       
    If validaAddMediosPago = False Then
        Call iniciofraMedioPago
        Exit Sub
    End If

    grxMedioPagos.Update
    Set rxRecosertMedPag = grxMedioPagos.ADORecordset
    tipocambio = TxtTipo_Cambio.Text
    If Right(cboMedioPagoMoneda, 3) = "SOL" Then
     tipocambio = 1
    End If
    
    rxRecosertMedPag.AddNew
    rxRecosertMedPag!Num_Corre = "falta"
    rxRecosertMedPag!cod_medpag = Left(cboMedioPago, 2)
    rxRecosertMedPag!des_medpago = Left(cboMedioPago, Len(cboMedioPagoMoneda) - 2)
    rxRecosertMedPag!Cod_Moneda = Right(cboMedioPagoMoneda, 3)
    rxRecosertMedPag!Nom_Moneda = Left(cboMedioPagoMoneda, Len(cboMedioPagoMoneda) - 3)
    rxRecosertMedPag!DOC_MEDPAG = txtMedioPagoDocumento.Text
    rxRecosertMedPag!IMP_MEDPAGO = txtMedioPagoImporte.Text
    rxRecosertMedPag!IMP_TOTALMEDPAG = tipocambio * txtMedioPagoImporte.Text
    rxRecosertMedPag!TIP_CAMBIO = tipocambio
    rxRecosertMedPag.Update
    
    Set grxMedioPagos.ADORecordset = rxRecosertMedPag
    Call TotalPagosMonedas
    Call ConfiguraGrillaMedioPago
    Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
Private Sub TotalPagosMonedas()
On Error GoTo ErrCal
    Dim Total As Double
    Dim rxmedioPago As New ADODB.Recordset
    Dim i As Integer
    
    If grxMedioPagos.RowCount = 0 Then
    Exit Sub
    Call iniciofraMedioPago
    End If
    
    grxMedioPagos.Update
    Set rxmedioPago = grxMedioPagos.ADORecordset
    rxmedioPago.Update
    i = 1
    rxmedioPago.MoveFirst
    Do While i <= rxmedioPago.RecordCount
        Total = Total + rxmedioPago("IMP_TOTALMEDPAG").Value
        rxmedioPago.MoveNext
        i = i + 1
    Loop
    
    'Right(cboMedioPagoMoneda, 2)
    'TxtTipo_Cambio.Text
    '"##0.00000"
    txtMedioPagoTotalPagoMN.Text = Format(Total, "##0.00")
    txtMedioPagoVueltoMN.Text = Format((Total - txt_total.Text), "##0.00")
    
    txtMedioPagoTotalPagoME.Text = Format(0, "##0.00")
    txtMedioPagoVueltoME.Text = Format(0, "##0.00")
    
    If TxtTipo_Cambio.Text > 0 Then
        txtMedioPagoTotalPagoME = Format((Total / TxtTipo_Cambio.Text), "##0.00")
        txtMedioPagoVueltoME = Format((Total / TxtTipo_Cambio.Text) - (txt_total.Text / TxtTipo_Cambio), "##0.00")
    End If
    
    txtMedioPagoImporte.Text = Format(0, "##0.00")
    If txt_total.Text - Total > 0 Then
        txtMedioPagoImporte.Text = Format((txt_total.Text - Total), "##0.00")
    End If
    
Exit Sub
ErrCal:
    MsgBox err.Description, vbCritical + vbOKOnly, "Cargar Calidades"
End Sub

Private Sub ConfiguraGrillaMedioPago()
    Dim C As Integer
    Dim colTemp As JSColumn
    Dim fmtCon  As JSFmtCondition
    On Error GoTo fin
    With grxMedioPagos
         For C = 1 To .Columns.Count
            .Columns(C).HeaderAlignment = jgexAlignCenter
            .Columns(C).TextAlignment = jgexAlignLeft
            .Columns(C).Visible = False
        Next C
        With .Columns("ELI")
             .Visible = True
             .Width = 1000
             .Caption = "X"
             .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("des_medpago")
             .Visible = True
             .Width = 1500
             .Caption = "MEDIO PAGO"
             .TextAlignment = jgexAlignLeft
        End With
        With .Columns("cod_moneda")
            .Visible = True
            .Width = 800
            .Caption = "COD"
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("nom_moneda")
            .Visible = True
            .Width = 1000
            .Caption = "MONEDA"
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("DOC_MEDPAG")
            .Visible = True
            .Width = 1300
            .Caption = "DOCUMENTO"
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("IMP_MEDPAGO")
            .Visible = True
            .Width = 1000
            .Caption = "IMPORTE"
            .TextAlignment = jgexAlignLeft
        End With
        With .Columns("tip_cambio")
            .Visible = True
            .Width = 800
            .Caption = "TC"
            .TextAlignment = jgexAlignLeft
        End With
        With .Columns("IMP_TOTALMEDPAG")
            .Visible = True
            .Width = 1200
            .Caption = "TOTAL"
            .TextAlignment = jgexAlignLeft
        End With
        'IMP_TOTALMEDPAG
    End With
        'GrxProductos.Columns("IMP_MEDPAGO").CellStyle = "Color_Cantidad"
    Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
'''************************************************************ELIMINA MEDIO DE PAGO****************************
Private Sub EliminaMediosPagos()

    If grxMedioPagos.RowCount = 0 Then Exit Sub
    
    Dim i As Integer
    Dim rstAux  As ADODB.Recordset
    grxMedioPagos.Update
    Set rstAux = grxMedioPagos.ADORecordset
    rstAux.MoveFirst
    i = 1
    Do While i <= rstAux.RecordCount
        If rstAux("ELI").Value = True Then
          rstAux.AbsolutePosition = grxMedioPagos.RowIndex(grxMedioPagos.Row)
          rstAux.Delete
        Else
          rstAux("ELI") = 0
        End If
        rstAux.MoveNext
        i = i + 1
    Loop
    Set grxMedioPagos.ADORecordset = rstAux
    Call TotalPagosMonedas
    Call ConfiguraGrillaMedioPago
    
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
       Set grxMedioPagos.ADORecordset = Nothing
       Call iniciofraMedioPago
       Call obtieneDatosInicialesMediosPago
    End If
    'Call Total_documento
    Call ConfiguraGrilla_Detalle
    
End Sub
''''*************************************************************SUMA TOTALES DE DOCUMENTO*********************************
Private Sub Total_documentoxx()
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
               
              If Not grxDatos.IsGroupItem(grxDatos.Row) = True And ColIndex > 0 Then
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

Private Sub Total_documento()
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
                
                Total = Total + rds("total").Value
                totalkilos = totalkilos + rds("cant").Value
              
                If i < rds.RecordCount Then
                    rds.MoveNext
                End If
                i = i + 1
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

    txtMedioPagoSubtotalMN.Text = CDbl(txt_subtotal.Text)
    txtMedioPagoIGVMN.Text = CDbl(txt_igv.Text)
    txtMedioPagoImporteMN.Text = CDbl(txt_total.Text)
         
    If TxtTipo_Cambio.Text > 0 Then
        txtMedioPagoSubtotalME.Text = Format((txt_subtotal.Text / TxtTipo_Cambio.Text), "####.00")
        txtMedioPagoIGVME.Text = Format((txt_igv.Text / TxtTipo_Cambio), "####.00")
        txtMedioPagoImporteME.Text = Format((txt_total.Text / TxtTipo_Cambio), "####.00")
    Else
        txtMedioPagoSubtotalME.Text = Format(0, "####.00")
        txtMedioPagoIGVME.Text = Format(0, "####.00")
        txtMedioPagoImporteME.Text = Format(0, "####.00")
    End If
    
   Exit Sub
ErrCal:
    MsgBox err.Description, vbCritical + vbOKOnly, "Cargar Calidades"
End Sub
'''*******************EVENTOS POR COLUMNA **********************************************************
Private Sub grxMedioPagos_AfterColEdit(ByVal ColIndex As Integer)
  AfterColEdit_MedioPago (ColIndex)
End Sub
Sub AfterColEdit_MedioPago(ByVal ColIndex As Integer)
Dim sSQL As String
On Error GoTo Error_Handler
Dim oGroup As GridEX20.JSGroup

Select Case ColIndex
Case Is = grxMedioPagos.Columns("ELI").Index
        Call EliminaMediosPagos
        Call iniciofraMedioPago
End Select
Exit Sub

Resume
Error_Handler:
errores err.Number
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
errores err.Number
End Sub
Private Sub GrxProductos_AfterColEdit(ByVal ColIndex As Integer)
On Error GoTo fin
  Dim a As Integer
  AfterColEdit_PRODUCTOS (ColIndex)

  Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub

Sub AfterColEdit_PRODUCTOS(ByVal ColIndex As Integer)
Dim sSQL As String
Dim saldo As Double
'On Error GoTo Error_Handler
On Error GoTo fin

Dim oGroup As GridEX20.JSGroup
Select Case ColIndex

  ' Case Is = GrxProductos.Columns("SEL").Index
  '  Call adicionarProducto
  ' Case Is = GrxProductos.Columns("PRECIO").Index
  ' If IsNumeric(GrxProductos.Value(GrxProductos.Columns("PRECIO").Index)) = False Or GrxProductos.Value(GrxProductos.Columns("PRECIO").Index) = "" Then
  '     GrxProductos.Value(GrxProductos.Columns("PRECIO").Index) = 0
  ' End If
  ' GrxProductos.Value(GrxProductos.Columns("TOTAL").Index) = GrxProductos.Value(GrxProductos.Columns("PRECIO").Index) * GrxProductos.Value(GrxProductos.Columns("CANT").Index)
  ' GridEX1.Value(GridEX1.Columns("sel").Index) = 0
    
  Case Is = GrxProductos.Columns("CANT").Index
    If IsNumeric(GrxProductos.Value(GrxProductos.Columns("CANT").Index)) = False Or GrxProductos.Value(GrxProductos.Columns("CANT").Index) = "" Then
        GrxProductos.Value(GrxProductos.Columns("CANT").Index) = 0
    End If
    GrxProductos.Value(GrxProductos.Columns("TOTAL").Index) = GrxProductos.Value(GrxProductos.Columns("PRECIO").Index) * CDbl(GrxProductos.Value(GrxProductos.Columns("CANT").Index))

    'GridEX1.Value(GridEX1.Columns("sel").Index) = 0
    
  End Select
Exit Sub

'Resume
'Error_Handler:
'errores Err.Number
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
'''***************************************evento click de las grillas  **********************************
Private Sub grxMedioPagos_Click()
    Dim ColIndex As Long
    Dim oRowData As JSRowData
    Dim SGRUPO As String
    Dim iRow As Long
    Dim i As Long
    Dim sCaptionGroup As String
        If grxMedioPagos.RowCount > 0 Then
         ColIndex = grxMedioPagos.Col
         If ColIndex = 0 Then Exit Sub
            If UCase(grxMedioPagos.Columns(ColIndex).Key) = "ELI" Then
                bClickColSelec = True
                SendKeys "{ENTER}"
            End If
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
   
StrSQL = "EXEC CF_MUESTRA_PRENDAS_MOV_TIENDA_ITEMS_FACTURA        '" & Opcion & _
                                                    "','" & Trim(txtCod_Almacen.Text) & _
                                                    "','','','" & Trim(txtDes_Present_Bus.Text) & _
                                                    "','','" & Trim(TxtCod_Estcli_Bus.Text) & _
                                                    "','" & Trim(txtDes_Estcli_Bus.Text) & _
                                                    "','" & Trim(txtCod_Ordpro_Bus.Text) & _
                                                    "','','" & Trim(txtCodigo_Barra_Bus.Text) & "'"
                                                    
    Set GrxProductos.ADORecordset = Nothing
    Set GrxProductos.ADORecordset = CargarRecordSetDesconectado(StrSQL, cConnect)
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
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
''''*************************************************************BUSQUEDA DE PRODUCTOS *********************************
Private Sub buscarProductosOtros(Opcion As String)

Dim StrSQL As String
Dim sCodCentroCosto As String
Dim nrofilas As Integer
Dim k, l As Long
Dim rsproductos  As New ADODB.Recordset
On Error GoTo fin
   
StrSQL = "EXEC CF_MUESTRA_ITEMS_TIENDA_FACTURA        '" & Opcion & _
                                                    "','" & Trim(txtCod_Almacen.Text) & _
                                                    "','" & Trim(TxtCod_Estcli_Bus.Text) & _
                                                    "','" & Trim(txtDes_Estcli_Bus.Text) & _
                                                    "','" & Trim(txtDes_Present_Bus.Text) & _
                                                    "','" & Trim(txtCodigo_Barra_Bus.Text) & "'"
                                                    
    Set GrxProductos.ADORecordset = Nothing
    Set GrxProductos.ADORecordset = CargarRecordSetDesconectado(StrSQL, cConnect)
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
   Set rsproductos = Nothing
Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
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
            .Visible = True
            .Width = 1300
            .Caption = "PRECIO"
            .TextAlignment = jgexAlignLeft
        End With
        With .Columns("TOTAL")
            .Visible = True
            .Width = 1300
            .Caption = "TOTAL"
            .TextAlignment = jgexAlignLeft
        End With

        SetColorDetalle
    End With
    
'    Dim oGroup01 As GridEX20.JSGroup
'    Dim oGroup02 As GridEX20.JSGroup
'    Dim valorcant   As JSColumn
'
'      With GrxProductos
'
'        Set oGroup01 = .Groups.Add(.Columns("OT").Index, jgexSortAscending)
'        .DefaultGroupMode = jgexDGMExpanded
'        .BackColorRowGroup = RGB(239, 235, 222)
'
'           .GroupFooterStyle = jgexTotalsGroupFooter
'           Set valorcant = .Columns("CANT")
'
'           With valorcant
'               .AggregateFunction = jgexSum
'               .TotalRowPrefix = "Total: "
'               .TextAlignment = jgexAlignRight
'           End With
'
'        End With
    Dim saldo As Double
    Set fmtCon = GrxProductos.FmtConditions.Add(GrxProductos.Columns("CANT").Index, jgexGreaterThan, 0)
    
    'If saldo > 0 Then
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
   
    StrSQL = "EXEC CN_MUESTRA_TELAS_DETALLE_FACTURA_prendas 'x','',''"
    
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
            .Visible = True
            .Width = 1300
            .Caption = "PRECIO"
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("TOTAL")
            .Visible = True
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
    
 Call SetColorDetalle
 Call Total_documento

    Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub
Private Sub ConfiguraGrilla_DetalleSinGrupos()
    Dim C As Integer
    On Error GoTo fin
    
 'SetColorDetalle
 'Call Total_documento
        
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
Private Sub txtDes_Estcli_Bus_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        If Left(cboTipoProducto, 2) = "01" Then
            buscarProductos (1)
        Else
            buscarProductosOtros (1)
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
    
'    If txtDes_Vendedor.Text <> "" Then
'       'txtCod_Almacen.SetFocus
'       txtCod_TipDoc.SetFocus
'    Else
'       txtCod_Vendedor.SetFocus
'    End If
    
End If
End Sub
Private Sub txtDes_Vendedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    buscaVendedor (2)
'    If txtDes_Vendedor.Text <> "" Then
'       'txtCod_Almacen.SetFocus
'       txtCod_TipDoc.SetFocus
'    Else
'       txtDes_Vendedor.SetFocus
    ''End If
    
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
  Call Busca_Opcion("cod_fabrica", "nom_fabrica", "tg_empresa where ", Txtcod_Fabrica, txtDes_Fabrica, 1)
    If Trim(txtDes_Fabrica.Text) <> "" Then
       txtCod_Tienda.SetFocus
    Else
       Txtcod_Fabrica.SetFocus
    End If
    
  End If
  
End Sub
Private Sub txtdes_fabrica_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        Call Busca_Opcion("cod_fabrica", "nom_fabrica", "tg_empresa where ", Txtcod_Fabrica, txtDes_Fabrica, 2)
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


Private Sub txtMedioPagoImporte_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    cmdMedioPagoAgregar.SetFocus
Else
    Call SoloNumeros(txtMedioPagoImporte, KeyAscii, True, 2)
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
      
       If Trim(txtNum_ruc.Text) <> "" Then
          txtCod_ConPag.SetFocus
       Else
          txtNum_ruc.SetFocus
       End If
        Set FrmBusqClientesPrendas = Nothing
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
   
    StrSQL = "EXEC CN_MUESTRA_GUIAS_PENDIENTES_FACTURACION_PRENDAS '" & Left(cboAlmacen, 2) & "','" & Trim(txtSerieGuia.Text) & "','" & Trim(txtNumeroGuia.Text) & "','" & Trim(txtNum_ruc.Tag) & "'"
    
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
    
    StrSQL = "EXEC CN_MUESTRA_GUIAS_ASOCIADAS_FACTURAS_PRENDAS '','" & Trim(txtSer_Docum.Text) & "','" & Trim(txtNum_Docum.Text) & "','" & Trim(txtNum_ruc.Tag) & "'"
    
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



