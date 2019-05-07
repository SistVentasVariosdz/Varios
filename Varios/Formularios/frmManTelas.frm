VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "mantenimientos.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManTelas 
   Caption         =   "Telas"
   ClientHeight    =   8640
   ClientLeft      =   1590
   ClientTop       =   2025
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   13305
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPrevious 
      Height          =   495
      Left            =   3255
      Picture         =   "frmManTelas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   81
      ToolTipText     =   "Anterior"
      Top             =   8040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdNext 
      Height          =   495
      Left            =   3855
      Picture         =   "frmManTelas.frx":0172
      Style           =   1  'Graphical
      TabIndex        =   175
      ToolTipText     =   "Siguiente"
      Top             =   8040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdFirst 
      Height          =   495
      Left            =   2760
      Picture         =   "frmManTelas.frx":02E4
      Style           =   1  'Graphical
      TabIndex        =   174
      ToolTipText     =   "Primero"
      Top             =   8040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdLast 
      Height          =   495
      Left            =   4335
      Picture         =   "frmManTelas.frx":0456
      Style           =   1  'Graphical
      TabIndex        =   173
      ToolTipText     =   "Ultimo"
      Top             =   8040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame FraComercial 
      Caption         =   "Comercial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3600
      TabIndex        =   138
      Top             =   1920
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox TxtGramaje_Comercial 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   945
         TabIndex        =   140
         Text            =   "0"
         Top             =   360
         Width           =   1110
      End
      Begin VB.TextBox TxtAncho_Comercial 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         TabIndex        =   139
         Text            =   "0"
         Top             =   345
         Width           =   1110
      End
      Begin FunctionsButtons.FunctButt FunctButt3 
         Height          =   510
         Left            =   1200
         TabIndex        =   171
         Top             =   840
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"frmManTelas.frx":05C8
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "Gramaje"
         Height          =   195
         Left            =   240
         TabIndex        =   142
         Top             =   480
         Width           =   585
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "Ancho"
         Height          =   195
         Left            =   2640
         TabIndex        =   141
         Top             =   480
         Width           =   465
      End
   End
   Begin VB.Frame FraLista 
      Caption         =   "Lista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   30
      TabIndex        =   82
      Top             =   1125
      Width           =   11910
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2310
         Left            =   0
         TabIndex        =   176
         Top             =   120
         Width           =   11805
         _ExtentX        =   20823
         _ExtentY        =   4075
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Cod_Tela"
            Caption         =   "Código"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Des_Tela"
            Caption         =   "Descripción"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Cod_UniMed"
            Caption         =   "U.M.Textil"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Cod_UniMedcnf"
            Caption         =   "U.M.Conf"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Acu_Porcentaje"
            Caption         =   "% Acumulado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3720.189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column04 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraBuscar 
      Caption         =   "Buscar Por:"
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11925
      Begin VB.Frame Frafamilia 
         Height          =   600
         Left            =   105
         TabIndex        =   1
         Top             =   375
         Width           =   9465
         Begin VB.TextBox txtdes_famgruite 
            Height          =   285
            Left            =   5460
            TabIndex        =   9
            Top             =   200
            Width           =   1575
         End
         Begin VB.TextBox txtdes_familia 
            Height          =   285
            Left            =   2190
            TabIndex        =   5
            Top             =   200
            Width           =   1575
         End
         Begin VB.CommandButton cmdBusFamItem 
            Caption         =   "..."
            Height          =   300
            Left            =   1875
            TabIndex        =   4
            Tag             =   "..."
            Top             =   200
            Width           =   360
         End
         Begin VB.TextBox txtfamilia 
            Height          =   285
            Left            =   1425
            MaxLength       =   2
            TabIndex        =   3
            Top             =   200
            Width           =   525
         End
         Begin VB.CommandButton cmdBusgrupo 
            Caption         =   "..."
            Height          =   300
            Left            =   5145
            TabIndex        =   8
            Top             =   200
            Width           =   360
         End
         Begin VB.TextBox txtgrupo 
            Height          =   285
            Left            =   4410
            TabIndex        =   7
            Top             =   200
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Familia de Tela"
            Height          =   195
            Left            =   150
            TabIndex        =   2
            Top             =   250
            Width           =   1065
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Grupo"
            Height          =   195
            Left            =   3915
            TabIndex        =   6
            Top             =   225
            Width           =   435
         End
      End
      Begin VB.Frame Fratela 
         Height          =   600
         Left            =   105
         TabIndex        =   85
         Top             =   360
         Width           =   7230
         Begin VB.CommandButton cmdBusTela 
            Caption         =   "..."
            Height          =   330
            Left            =   2560
            TabIndex        =   100
            Tag             =   "..."
            Top             =   240
            Width           =   360
         End
         Begin VB.TextBox txtdes_tela 
            Height          =   285
            Left            =   2910
            TabIndex        =   88
            Top             =   270
            Width           =   3720
         End
         Begin VB.TextBox txtcod_tela 
            Height          =   285
            Left            =   1590
            MaxLength       =   8
            TabIndex        =   87
            Top             =   270
            Width           =   1005
         End
         Begin VB.Label Label2 
            Caption         =   "Tela"
            Height          =   240
            Left            =   360
            TabIndex        =   89
            Top             =   330
            Width           =   690
         End
      End
      Begin VB.Frame fraoptions 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   345
         Left            =   1440
         TabIndex        =   84
         Top             =   120
         Width           =   6135
         Begin VB.OptionButton optFlg_Operatividad 
            Caption         =   "Telas NO Operativas"
            Height          =   195
            Left            =   4080
            TabIndex        =   122
            Top             =   5
            Width           =   1785
         End
         Begin VB.OptionButton optfamtela 
            Caption         =   "Familia de Tela"
            Height          =   195
            Left            =   0
            TabIndex        =   11
            Top             =   5
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optcliente 
            Caption         =   "Cliente"
            Height          =   195
            Left            =   2880
            TabIndex        =   13
            Top             =   5
            Width           =   1410
         End
         Begin VB.OptionButton opttela 
            Caption         =   "Tela"
            Height          =   195
            Left            =   1800
            TabIndex        =   12
            Top             =   5
            Width           =   1425
         End
      End
      Begin FunctionsButtons.FunctButt FunctBuscar 
         Height          =   495
         Left            =   9960
         TabIndex        =   10
         Top             =   420
         Width           =   1260
         _ExtentX        =   2117
         _ExtentY        =   873
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&BUSCAR~0~0~1~~0~Falso~Falso~&BUSCAR~"
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Frame Fracliente 
         Height          =   600
         Left            =   105
         TabIndex        =   86
         Top             =   375
         Width           =   7305
         Begin VB.CommandButton cmdBusCliente 
            Caption         =   "..."
            Height          =   330
            Left            =   1650
            TabIndex        =   95
            Tag             =   "..."
            Top             =   240
            Width           =   360
         End
         Begin VB.TextBox txtcliente 
            Height          =   285
            Left            =   930
            MaxLength       =   5
            TabIndex        =   94
            Top             =   270
            Width           =   765
         End
         Begin VB.TextBox txttemporada 
            Height          =   285
            Left            =   4665
            TabIndex        =   93
            Top             =   270
            Width           =   735
         End
         Begin VB.CommandButton cmdBusTemporada 
            Caption         =   "..."
            Height          =   330
            Left            =   5370
            TabIndex        =   92
            Top             =   240
            Width           =   360
         End
         Begin VB.TextBox txtNom_TemCli 
            Height          =   285
            Left            =   5730
            TabIndex        =   91
            Top             =   270
            Width           =   1455
         End
         Begin VB.TextBox txtNom_Cliente 
            Height          =   285
            Left            =   2010
            TabIndex        =   90
            Top             =   270
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "Cliente"
            Height          =   255
            Left            =   60
            TabIndex        =   97
            Top             =   270
            Width           =   540
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Temporada"
            Height          =   195
            Left            =   3810
            TabIndex        =   96
            Top             =   300
            Width           =   810
         End
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt2 
      Height          =   450
      Left            =   0
      TabIndex        =   164
      Top             =   360
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   794
      Custom          =   "0~0~TWILL~Verdadero~Verdadero~&Mts. Twill x Hora~0~0~1~~0~Falso~Falso~&Mts. Twill x Hora~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1100
      ControlHeigth   =   430
      ControlSeparator=   110
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   4530
      Left            =   12000
      TabIndex        =   169
      Top             =   0
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   7990
      Custom          =   $"frmManTelas.frx":065E
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1200
      ControlHeigth   =   450
      ControlSeparator=   0
   End
   Begin FunctionsButtons.FunctButt FunctCambios 
      Height          =   3960
      Left            =   12000
      TabIndex        =   170
      Top             =   240
      Visible         =   0   'False
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   6985
      Custom          =   $"frmManTelas.frx":0A3D
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1250
      ControlHeigth   =   420
      ControlSeparator=   20
   End
   Begin VB.Frame Fradetalle 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4245
      Left            =   0
      TabIndex        =   83
      Tag             =   "Detail"
      Top             =   3720
      Width           =   13245
      Begin VB.Frame fraUnidadesMedida 
         Caption         =   "Datos Técnicos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   120
         TabIndex        =   57
         Top             =   2520
         Width           =   13005
         Begin VB.TextBox txtCunSas 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   12600
            TabIndex        =   163
            Text            =   "0"
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtColumnas 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   9240
            TabIndex        =   161
            Text            =   "0"
            Top             =   1200
            Width           =   975
         End
         Begin VB.OptionButton OptCentimetros 
            Caption         =   "Ctms."
            Height          =   195
            Left            =   7320
            TabIndex        =   157
            Top             =   1200
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton OptPulgadas 
            Caption         =   "Pulg."
            Height          =   195
            Left            =   6600
            TabIndex        =   158
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox TxtAnchoLavado 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6480
            TabIndex        =   155
            Text            =   "0"
            Top             =   840
            Width           =   630
         End
         Begin VB.TextBox TxtTipoCorte 
            Height          =   285
            Left            =   8400
            MaxLength       =   3
            TabIndex        =   153
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox TxtDes_TipoCorte 
            Height          =   285
            Left            =   8880
            TabIndex        =   152
            Top             =   480
            Width           =   3825
         End
         Begin VB.TextBox TxtGram_Comercial 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4320
            TabIndex        =   144
            Text            =   "0"
            Top             =   350
            Width           =   1005
         End
         Begin VB.TextBox TxtAnc_Comercial 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4320
            TabIndex        =   143
            Text            =   "0"
            Top             =   690
            Width           =   1005
         End
         Begin VB.TextBox txtAncho_Acab_Abierto 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6465
            Locked          =   -1  'True
            TabIndex        =   132
            Text            =   "0"
            Top             =   480
            Width           =   630
         End
         Begin VB.TextBox TxtMts_Twill_x_Hora 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   10755
            TabIndex        =   80
            Text            =   "0"
            Top             =   840
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.TextBox TxtGramDesLavado 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7170
            TabIndex        =   120
            Text            =   "0"
            Top             =   480
            Width           =   1005
         End
         Begin VB.Frame Frame3 
            Caption         =   "Encogimientos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   105
            TabIndex        =   115
            Top             =   1005
            Width           =   5265
            Begin VB.TextBox txtEncog_Ancho 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   3870
               TabIndex        =   117
               Text            =   "0"
               Top             =   210
               Width           =   1230
            End
            Begin VB.TextBox txtEncog_Largo 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   1155
               TabIndex        =   116
               Text            =   "0"
               Top             =   210
               Width           =   1215
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "Largo :"
               Height          =   195
               Left            =   210
               TabIndex        =   119
               Top             =   255
               Width           =   495
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "Ancho :"
               Height          =   195
               Left            =   2940
               TabIndex        =   118
               Top             =   255
               Width           =   555
            End
         End
         Begin VB.TextBox txtFactor_Ajuste_Explosion 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8640
            TabIndex        =   79
            Text            =   "0"
            Top             =   840
            Width           =   750
         End
         Begin VB.TextBox TxtRevirado 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   11040
            TabIndex        =   106
            Text            =   "0"
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtNum_Lavadas 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   13725
            TabIndex        =   72
            Text            =   "0"
            Top             =   945
            Visible         =   0   'False
            Width           =   1230
         End
         Begin VB.TextBox txtAncho_Acab 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2985
            TabIndex        =   69
            Text            =   "0"
            Top             =   690
            Width           =   1005
         End
         Begin VB.TextBox txtGramaje_Acab 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2985
            TabIndex        =   66
            Text            =   "0"
            Top             =   350
            Width           =   1005
         End
         Begin VB.TextBox txtAncho_Crudo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   780
            TabIndex        =   63
            Text            =   "0"
            Top             =   690
            Width           =   1110
         End
         Begin VB.TextBox txtGramaje_Crudo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   780
            TabIndex        =   60
            Text            =   "0"
            Top             =   350
            Width           =   1110
         End
         Begin VB.TextBox txtEncog_Largo_Vap 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   13245
            TabIndex        =   75
            Text            =   "0"
            Top             =   405
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtEncog_Ancho_Vap 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   13245
            TabIndex        =   77
            Text            =   "0"
            Top             =   720
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Frame FraGradoDoblez 
            Caption         =   "Grado de Linea doblez"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   960
            Left            =   12840
            TabIndex        =   108
            Top             =   480
            Visible         =   0   'False
            Width           =   1995
            Begin VB.OptionButton OptNinguna 
               Caption         =   "Ninguna"
               Height          =   210
               Left            =   120
               TabIndex        =   111
               Top             =   210
               Width           =   1425
            End
            Begin VB.OptionButton OptMangas 
               Caption         =   "Solo mangas"
               Height          =   210
               Left            =   120
               TabIndex        =   110
               Top             =   420
               Width           =   1425
            End
            Begin VB.OptionButton OptAmbas 
               Caption         =   "Mangas y/o espaldas"
               Height          =   210
               Left            =   120
               TabIndex        =   109
               Top             =   630
               Width           =   1845
            End
         End
         Begin VB.Frame FraInclinacion 
            Caption         =   "Inclinación de Trama"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   960
            Left            =   12720
            TabIndex        =   112
            Top             =   480
            Visible         =   0   'False
            Width           =   2100
            Begin VB.OptionButton Opt1 
               Caption         =   "1% - 3 % (de 1º a 2.5º)"
               Height          =   210
               Left            =   185
               TabIndex        =   114
               Top             =   630
               Width           =   1950
            End
            Begin VB.OptionButton Opt0 
               Caption         =   "0% - 1% (de 0º a 1º)"
               Height          =   210
               Left            =   185
               TabIndex        =   113
               Top             =   315
               Width           =   1740
            End
         End
         Begin VB.Label Label60 
            Caption         =   "Revirado"
            Height          =   255
            Left            =   10320
            TabIndex        =   166
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label59 
            Caption         =   "CunSas :"
            Height          =   255
            Left            =   0
            TabIndex        =   165
            Top             =   0
            Width           =   735
         End
         Begin VB.Label TxtxCunSas 
            Caption         =   "CunSas :"
            Height          =   255
            Left            =   12240
            TabIndex        =   162
            Top             =   960
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label58 
            Caption         =   "Columnas :"
            Height          =   255
            Left            =   8280
            TabIndex        =   160
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Medida:"
            Height          =   195
            Left            =   5520
            TabIndex        =   159
            Top             =   1200
            Width           =   930
         End
         Begin VB.Label Label56 
            Caption         =   "Ancho Lavado"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   5685
            TabIndex        =   156
            Top             =   720
            Width           =   780
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Familia Tela Corte:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8400
            TabIndex        =   154
            Top             =   240
            Width           =   2040
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            Caption         =   "Comercial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4425
            TabIndex        =   147
            Top             =   180
            Width           =   840
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "gr."
            Height          =   195
            Index           =   1
            Left            =   5370
            TabIndex        =   146
            Top             =   345
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "mts."
            Height          =   195
            Left            =   5370
            TabIndex        =   145
            Top             =   705
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label Label46 
            Caption         =   "Ancho Abierto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   405
            Left            =   5640
            TabIndex        =   133
            Top             =   360
            Width           =   780
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "Mts. Twill x Hora"
            Height          =   195
            Left            =   9480
            TabIndex        =   123
            Top             =   855
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            Caption         =   "Gramaje Despues de Lavado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5640
            TabIndex        =   121
            Top             =   120
            Width           =   2820
         End
         Begin VB.Label lblFactor_Ajuste_Explosion 
            AutoSize        =   -1  'True
            Caption         =   "Fact. Explosión:"
            Height          =   195
            Left            =   7320
            TabIndex        =   78
            Top             =   840
            Width           =   1125
         End
         Begin VB.Label Label40 
            Caption         =   "%"
            Height          =   255
            Left            =   12600
            TabIndex        =   107
            Top             =   840
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label39 
            Caption         =   "Revirado"
            Height          =   255
            Left            =   13065
            TabIndex        =   105
            Top             =   930
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "mts."
            Height          =   195
            Left            =   4035
            TabIndex        =   104
            Top             =   705
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "mts."
            Height          =   195
            Left            =   1935
            TabIndex        =   103
            Top             =   705
            Width           =   285
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "# Lavados :"
            Height          =   195
            Left            =   12675
            TabIndex        =   71
            Top             =   1050
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "gr."
            Height          =   195
            Index           =   2
            Left            =   4035
            TabIndex        =   67
            Top             =   345
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "gr."
            Height          =   195
            Index           =   0
            Left            =   1935
            TabIndex        =   61
            Top             =   360
            Width           =   180
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Encogimiento Lavado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   12600
            TabIndex        =   70
            Top             =   240
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Largo :"
            Height          =   195
            Left            =   12600
            TabIndex        =   74
            Top             =   450
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Ancho :"
            Height          =   195
            Left            =   12600
            TabIndex        =   76
            Top             =   765
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Crudo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1065
            TabIndex        =   58
            Top             =   180
            Width           =   510
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Acabados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3090
            TabIndex        =   64
            Top             =   180
            Width           =   855
         End
         Begin VB.Label Label22 
            Caption         =   "Ancho Tubular"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   2250
            TabIndex        =   68
            Top             =   600
            Width           =   780
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Gramaje"
            Height          =   195
            Left            =   2250
            TabIndex        =   65
            Top             =   345
            Width           =   585
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Ancho"
            Height          =   195
            Left            =   75
            TabIndex        =   62
            Top             =   690
            Width           =   465
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Gramaje"
            Height          =   195
            Left            =   75
            TabIndex        =   59
            Top             =   345
            Width           =   585
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Encogimiento Vaporizado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   12480
            TabIndex        =   73
            Top             =   480
            Visible         =   0   'False
            Width           =   2160
         End
      End
      Begin VB.TextBox TxtSinRevision 
         Height          =   285
         Left            =   11520
         TabIndex        =   151
         Text            =   " SIN REVISION"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Enabled         =   0   'False
         Height          =   735
         Left            =   11520
         TabIndex        =   148
         Top             =   1800
         Width           =   1575
         Begin MSComCtl2.DTPicker DTPUltRevision 
            Height          =   255
            Left            =   0
            TabIndex        =   149
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   72351745
            CurrentDate     =   38794
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            Caption         =   "Fec. Ult. Revision"
            Height          =   195
            Left            =   120
            TabIndex        =   150
            Top             =   120
            Width           =   1260
         End
      End
      Begin VB.TextBox TxtSufijo 
         Height          =   285
         Left            =   10680
         MaxLength       =   1
         TabIndex        =   128
         Top             =   2205
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox TxtDes_TelaOriginal 
         Height          =   285
         Left            =   5520
         MaxLength       =   50
         TabIndex        =   127
         Top             =   2205
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.TextBox TxtCod_TelaOriginal 
         Height          =   285
         Left            =   4200
         MaxLength       =   8
         TabIndex        =   125
         Top             =   2205
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox ChkTelaProcesoAdicional 
         Caption         =   "Tela Actual con Proceso Adicional"
         Height          =   195
         Left            =   120
         TabIndex        =   124
         Top             =   2325
         Width           =   2775
      End
      Begin VB.Frame Frame1 
         Caption         =   "Comentario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   49
         Top             =   2760
         Visible         =   0   'False
         Width           =   12690
         Begin VB.TextBox TxtRendimiento 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   9000
            MaxLength       =   8
            TabIndex        =   136
            Top             =   760
            Width           =   1125
         End
         Begin VB.TextBox TxtCod_OrdTra 
            Height          =   285
            Left            =   1920
            MaxLength       =   8
            TabIndex        =   55
            Top             =   760
            Width           =   1125
         End
         Begin VB.TextBox TxtCod_OrdTra_Tejeduria 
            Height          =   285
            Left            =   5130
            MaxLength       =   8
            TabIndex        =   56
            Top             =   760
            Width           =   1125
         End
         Begin VB.ComboBox cboCombo 
            Height          =   315
            Left            =   7440
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   440
            Width           =   4215
         End
         Begin VB.TextBox txtDes_Tel_Origen 
            Height          =   285
            Left            =   3120
            TabIndex        =   53
            Top             =   440
            Width           =   3165
         End
         Begin VB.TextBox txtCod_Tel_Origen 
            Height          =   285
            Left            =   1890
            MaxLength       =   8
            TabIndex        =   52
            Top             =   440
            Width           =   1125
         End
         Begin VB.TextBox txtComentario 
            Height          =   285
            Left            =   1050
            MaxLength       =   500
            TabIndex        =   51
            Top             =   145
            Width           =   11475
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "Rendimiento (Mts./Kgs.)"
            Height          =   195
            Left            =   7080
            TabIndex        =   137
            Top             =   800
            Width           =   1710
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "Partida Tintoreria"
            Height          =   195
            Left            =   480
            TabIndex        =   135
            Top             =   800
            Width           =   1200
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "O/T Tejeduria"
            Height          =   195
            Left            =   3960
            TabIndex        =   134
            Top             =   800
            Width           =   1005
         End
         Begin VB.Label Label45 
            Caption         =   "Combinacion"
            Height          =   240
            Left            =   6360
            TabIndex        =   131
            Top             =   525
            Width           =   960
         End
         Begin VB.Label Label44 
            Caption         =   "Tela Desarrollo Origen"
            Height          =   240
            Left            =   120
            TabIndex        =   130
            Top             =   500
            Width           =   1680
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Comentario:"
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   200
            Width           =   840
         End
      End
      Begin VB.Frame fraDatosGenerales 
         Caption         =   "Datos Específicos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   120
         TabIndex        =   98
         Top             =   50
         Width           =   6255
         Begin VB.TextBox txtDes_Tela_Comercial 
            Height          =   285
            Left            =   960
            MaxLength       =   150
            TabIndex        =   167
            Top             =   1780
            Width           =   5175
         End
         Begin VB.ComboBox cboTip_Ancho 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1490
            Width           =   1695
         End
         Begin VB.ComboBox cboCod_TipRaya 
            Height          =   315
            Left            =   3600
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1170
            Width           =   2055
         End
         Begin VB.ComboBox cboCodTip_Tela 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1170
            Width           =   1695
         End
         Begin VB.TextBox txtCod_CtaCont 
            Height          =   285
            Left            =   3600
            MaxLength       =   14
            TabIndex        =   31
            Top             =   1490
            Width           =   1695
         End
         Begin VB.ComboBox cboCod_GruTela 
            Height          =   315
            Left            =   3600
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   540
            Width           =   2055
         End
         Begin VB.ComboBox cboCod_UniMedcnf 
            Height          =   315
            Left            =   3600
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   860
            Width           =   2055
         End
         Begin VB.ComboBox cboCod_UniMed 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   860
            Width           =   1695
         End
         Begin VB.ComboBox cboCod_FamTela 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   540
            Width           =   1695
         End
         Begin VB.TextBox txtdestela 
            Height          =   285
            Left            =   960
            MaxLength       =   150
            TabIndex        =   15
            Top             =   240
            Width           =   5175
         End
         Begin VB.TextBox txtcodtela 
            Height          =   285
            Left            =   960
            MaxLength       =   8
            TabIndex        =   99
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            Caption         =   "N-Comercial"
            Height          =   195
            Left            =   30
            TabIndex        =   168
            Top             =   1815
            Width           =   855
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "T. Ancho"
            Height          =   195
            Left            =   30
            TabIndex        =   28
            Top             =   1545
            Width           =   660
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "T. Raya"
            Height          =   195
            Left            =   2760
            TabIndex        =   26
            Top             =   1230
            Width           =   570
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "T. Tela"
            Height          =   195
            Left            =   30
            TabIndex        =   24
            Top             =   1230
            Width           =   510
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "C.Contab"
            Height          =   195
            Left            =   2760
            TabIndex        =   30
            Top             =   1540
            Width           =   660
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "U.M.Cnf"
            Height          =   195
            Left            =   2760
            TabIndex        =   22
            Top             =   910
            Width           =   585
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "U.M.Textil"
            Height          =   195
            Left            =   30
            TabIndex        =   20
            Top             =   915
            Width           =   720
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Grupo :"
            Height          =   195
            Left            =   2760
            TabIndex        =   18
            Top             =   590
            Width           =   525
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Familia"
            Height          =   195
            Left            =   30
            TabIndex        =   16
            Top             =   585
            Width           =   480
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Tela"
            Height          =   195
            Left            =   30
            TabIndex        =   14
            Top             =   315
            Width           =   315
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos Numéricos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Left            =   6480
         TabIndex        =   32
         Top             =   0
         Width           =   5385
         Begin VB.ComboBox cboFlg_Operatividad 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3780
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   1110
            Width           =   1485
         End
         Begin VB.TextBox txtPeso 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   4200
            TabIndex        =   102
            Text            =   "0"
            Top             =   1500
            Width           =   1005
         End
         Begin VB.TextBox txtLongMalla 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   48
            Text            =   "0"
            Top             =   1545
            Width           =   855
         End
         Begin VB.TextBox txtRapport 
            Height          =   285
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   46
            Text            =   " "
            Top             =   1230
            Width           =   1920
         End
         Begin VB.ComboBox cboCod_Galga 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   880
            Width           =   1935
         End
         Begin VB.TextBox txtNum_Rpm 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4200
            TabIndex        =   40
            Text            =   "0"
            Top             =   560
            Width           =   975
         End
         Begin VB.TextBox txtNum_Aguja 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            TabIndex        =   38
            Text            =   "0"
            Top             =   560
            Width           =   855
         End
         Begin VB.TextBox txtNum_Alimentadores 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4200
            TabIndex        =   36
            Text            =   "0"
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtdiametro 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            TabIndex        =   34
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Operatividad:"
            Height          =   195
            Left            =   3780
            TabIndex        =   43
            Top             =   885
            Width           =   945
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Peso (kgs) x Und. Cnf"
            Height          =   195
            Left            =   2400
            TabIndex        =   101
            Top             =   1605
            Width           =   1545
         End
         Begin VB.Label LblLongMalla 
            AutoSize        =   -1  'True
            Caption         =   "Long. de Malla :"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   1590
            Width           =   1140
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Rapport"
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   1260
            Width           =   570
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Galga"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   930
            Width           =   420
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Número R.P.M."
            Height          =   195
            Left            =   2520
            TabIndex        =   39
            Top             =   615
            Width           =   1095
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Número. Agujas"
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   610
            Width           =   1125
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Número Alimentadores"
            Height          =   195
            Left            =   2520
            TabIndex        =   35
            Top             =   285
            Width           =   1590
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Diámetro Galga"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   290
            Width           =   1095
         End
      End
      Begin VB.Label LblSufijo 
         AutoSize        =   -1  'True
         Caption         =   "Sufijo"
         Height          =   195
         Left            =   10200
         TabIndex        =   129
         Top             =   2235
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label LblTelaOriginal 
         AutoSize        =   -1  'True
         Caption         =   "Tela Original"
         Height          =   195
         Left            =   3240
         TabIndex        =   126
         Top             =   2235
         Visible         =   0   'False
         Width           =   1005
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   525
      Left            =   5040
      TabIndex        =   172
      Top             =   8040
      Width           =   3555
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmManTelas.frx":0DB4
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   0
      Top             =   3150
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmManTelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public CODIGO, Descripcion As String, TipoAdd As String, tipoAdd2 As String
Dim Opcion As Integer
Dim sTipo As String
Dim StrSQL As String
Dim rsgrid As ADODB.Recordset
Dim varCod_Tela As String, sFlg_Factor_Ajuste_Explosion As String

'Para Seleccion de familias
Public varCadena_Familias As String
Public varCancelImpresion As Integer

Dim doblez, inclinacion As String
Dim iColum As Long

Private Sub cbogrupo_Click()
    Call CARGA_GRID
End Sub

Private Sub cboCod_FamTela_Click()

    Dim varCod_TipFamTela As String
    'Combo Grupo Item
    cboCod_GruTela.Clear
    'SELECT Cod_GruTela as Código, Des_GruTela as Descripción FROM TX_GRUTELA
    StrSQL = "SELECT Des_GruTela + space(100) + Cod_GruTela FROM TX_GRUTELA WHERE Cod_FamTela='" & Right(cboCod_FamTela.Text, 2) & "'"
    Call LlenaCombo(cboCod_GruTela, StrSQL, cConnect)
    
    StrSQL = "SELECT ISNULL(Cod_TipFamTela,'') FROM Tx_FamTela WHERE Cod_FamTela = '" & Right(cboCod_FamTela.Text, 2) & "'"
    varCod_TipFamTela = DevuelveCampo(StrSQL, cConnect)
    
    If varCod_TipFamTela = "N" Then
        StrSQL = "SELECT Cod_UniMedTex FROM TG_CONTROL"
        Call BuscaCombo(DevuelveCampo(StrSQL, cConnect), 2, Me.cboCod_UniMed)
        
        StrSQL = "SELECT Cod_UniMedCnfNor FROM TG_CONTROL"
        Call BuscaCombo(DevuelveCampo(StrSQL, cConnect), 2, Me.cboCod_UniMedcnf)
    Else
        StrSQL = "SELECT Cod_UniMedTex FROM TG_CONTROL"
        Call BuscaCombo(DevuelveCampo(StrSQL, cConnect), 2, Me.cboCod_UniMed)
        
        StrSQL = "SELECT Cod_UniMedCnfRec  FROM TG_CONTROL "
        Call BuscaCombo(DevuelveCampo(StrSQL, cConnect), 2, Me.cboCod_UniMedcnf)
    End If
    
    If UCase(Mid(cboCod_FamTela, 1, 2)) = "TW" Then
        TxtMts_Twill_x_Hora.Enabled = True
    Else
        TxtMts_Twill_x_Hora.Enabled = False
    End If
    
End Sub




Private Sub cboCodTip_Tela_Click()
    If Right(cboCodTip_Tela, 1) <> "L" Then
        cboCod_TipRaya.ListIndex = -1
    End If
End Sub

Private Sub cboCombo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboFlg_Operatividad_Click()
    cboFlg_Operatividad.ForeColor = cboCod_TipRaya.ForeColor
    If Left(cboFlg_Operatividad, 1) = "N" Then cboFlg_Operatividad.ForeColor = vbRed
End Sub

Private Sub ChkTelaProcesoAdicional_Click()
If ChkTelaProcesoAdicional Then
    LblTelaOriginal.Visible = True
    TxtCod_TelaOriginal.Visible = True
    TxtDes_TelaOriginal.Visible = True
    LblSufijo.Visible = True
    TxtSufijo.Visible = True
Else
    LblTelaOriginal.Visible = False
    TxtCod_TelaOriginal.Visible = False
    TxtDes_TelaOriginal.Visible = False
    LblSufijo.Visible = False
    TxtSufijo.Visible = False
    TxtCod_TelaOriginal.Text = ""
    TxtDes_TelaOriginal.Text = ""
    TxtSufijo.Text = ""
End If
End Sub

Private Sub cmdBusTela_Click()
    Dim StrSQL As String
    If Trim(txtcod_tela.Text) <> "" Then
        StrSQL = "SELECT Cod_Tela as Código, Des_Tela as Descripción FROM TX_TELA WHERE Cod_Tela='" & txtcod_tela.Text & "'"
    Else
        If Len(Trim(txtdes_tela.Text)) < 5 Then
            Call MsgBox("El Texto Ingresado debe contar con un mínimo de 5 caracteres", vbExclamation)
            Exit Sub
        Else
            StrSQL = "SELECT Cod_Tela as Código, Des_Tela as Descripción FROM TX_TELA WHERE Des_Tela LIKE '" & Trim(txtdes_tela.Text) & "%'"
        End If
    End If
    
    Dim oTipo As New frmBusqGeneral
    Dim RS As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.SQuery = StrSQL
    oTipo.Cargar_Datos
    oTipo.Show 1
    If CODIGO <> "" Then
        txtcod_tela.Text = CODIGO
        txtdes_tela.Text = Descripcion
        FunctBuscar.SetFocus
    End If
    Set oTipo = Nothing
    Set RS = Nothing
End Sub

Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If rsgrid.State <> 1 Then
    Exit Sub
End If
If Not rsgrid.EOF And Not rsgrid.BOF Then
    Call CargaDatos
End If
End Sub

Private Sub cmdBusCliente_Click()
    Dim oTipo As New frmBusqGeneral
    Dim RS As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.SQuery = "SELECT Abr_Cliente as Código, nom_cliente as Descripción FROM TG_Cliente order by 1"
    oTipo.Cargar_Datos
    oTipo.Show 1
    If CODIGO <> "" Then
        txtcliente.Text = CODIGO
        txtNom_Cliente.Text = Descripcion
        CODIGO = ""
    End If
    Set oTipo = Nothing
    Set RS = Nothing
End Sub

Private Sub cmdBusFamItem_Click()
    Dim oTipo As New frmBusqGeneral
    Dim RS As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.SQuery = "SELECT Cod_FamTela as Código, Des_FamTela as Descripción FROM TX_FAMTELA"
    oTipo.Cargar_Datos
    oTipo.Show 1
    If CODIGO <> "" Then
        txtfamilia.Text = CODIGO
        txtdes_familia.Text = Descripcion
        
        txtgrupo.Enabled = True
        cmdBusgrupo.Enabled = True
    End If
    Set oTipo = Nothing
    Set RS = Nothing
End Sub

Private Sub cmdBusgrupo_Click()
    Dim oTipo As New frmBusqGeneral
    Dim RS As New ADODB.Recordset
    Set oTipo.oParent = Me
    oTipo.SQuery = "SELECT Cod_GruTela as Código, Des_GruTela as Descripción FROM TX_GRUTELA WHERE Cod_FamTela='" & Trim(txtfamilia.Text) & "'"
    oTipo.Cargar_Datos
    oTipo.Show 1
    If CODIGO <> "" Then
        txtgrupo.Text = CODIGO
        txtdes_famgruite = Descripcion
        CODIGO = ""
    End If
    Set oTipo = Nothing
    Set RS = Nothing
End Sub

Private Sub cmdBusTemporada_Click()
    Dim oTipo As New frmBusqGeneral
    Dim RS As New ADODB.Recordset
    Set oTipo.oParent = Me
    StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtcliente.Text) & "'"
    oTipo.SQuery = "SELECT  Cod_TemCli as Código, Nom_TemCli as Descripción FROM TG_TemCli WHERE Cod_Cliente='" & DevuelveCampo(StrSQL, cConnect) & "'"
    oTipo.Cargar_Datos
    oTipo.Show 1
    If CODIGO <> "" Then
        txttemporada.Text = CODIGO
        txtNom_TemCli.Text = Descripcion
        CODIGO = ""
    End If
    Set oTipo = Nothing
    Set RS = Nothing
End Sub

Private Sub cmdFirst_Click()
    If Not rsgrid.BOF Then
        rsgrid.MoveFirst
    End If
End Sub

Private Sub cmdLast_Click()
    If Not rsgrid.EOF Then
        rsgrid.MoveLast
    End If
End Sub

Private Sub cmdNext_Click()
    If Not rsgrid.EOF Then
        rsgrid.MoveNext
        If rsgrid.EOF Then
            rsgrid.MoveLast
        End If
    End If
End Sub

Private Sub cmdPrevious_Click()
    If Not rsgrid.BOF Then
        rsgrid.MovePrevious
        If rsgrid.BOF Then
            rsgrid.MoveFirst
        End If
    End If
End Sub


Private Sub Form_Activate()
    Dim varCodFamTela As String
    varCodFamTela = Right(cboCod_FamTela.Text, 2)
    'Llena Familia de Telas
    If Opcion = 1 And varCodFamTela <> "" Then
        StrSQL = "SELECT des_famtela + space(100) + cod_famtela  FROM TX_FamTela"
        Call LlenaCombo(cboCod_FamTela, StrSQL, cConnect)
        Call BuscaCombo(varCodFamTela, 2, cboCod_FamTela)
    End If
    
End Sub

Private Sub Form_Load()
    Call FormSet(Me)
    FormateaGrid Me.DGridLista
    Opcion = 1
    Call CargaCombos
    StrSQL = "SELECT TOP 1 flg_factor_ajuste_explosion FROM TG_CONTROL"
    sFlg_Factor_Ajuste_Explosion = DevuelveCampo(StrSQL, cConnect)
    Call CARGA_GRID
    INHABILITA_DATOS
   ' Me.FunctTemporada.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    'Me.FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    'Me.FunctCambios.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    'FunctButt2.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    'MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    
    
    Me.OptPulgadas.Value = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub


Private Sub FunctBuscar_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    txtfamilia.Text = UCase(txtfamilia.Text)
    Call CARGA_GRID
End Sub

'Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
'Dim vCod_Cliente As Variant
'    Select Case ActionName
'        Case "COMPOSICION"
''            If Not rsgrid.EOF Then
''                Load frmMantHilosTel
''                frmMantHilosTel.Codigo_tela = rsgrid("Cod_Tela")
''                frmMantHilosTel.txtdes_tela = rsgrid("Des_Tela")
''                frmMantHilosTel.CARGA_GRID
''                frmMantHilosTel.Show 1
''            Else
''                MsgBox ("Debe seleccionar una Tela para acceder a esta opcion")
''            End If
'        Case "PROCESOS"
''            If Not rsgrid.EOF Then
''                Load frmMantTelaPro
''                frmMantTelaPro.Codigo_tela = rsgrid("Cod_Tela")
''                frmMantTelaPro.txtdes_tela = rsgrid("Des_Tela")
''                frmMantTelaPro.CARGA_GRID
''                frmMantTelaPro.Show 1
''            Else
''                MsgBox ("Debe seleccionar una Tela para acceder a esta opcion")
''            End If
'        Case "COMBINACIONES"
'
''            If Right(cboCodTip_Tela, 1) = "S" Then
''                Call MsgBox("El Tipo de Tela seleccionado no permite acceder a esta opción. Sirvase verificar", vbCritical)
''                Exit Sub
''            End If
'
''            If Not rsgrid.EOF Then
''                strSQL = "SELECT Cod_Cliente FROM TG_Cliente " & _
''                         "WHERE Abr_Cliente = '" & txtcliente & "'"
''                vCod_Cliente = DevuelveCampo(strSQL, cConnect)
''                Load frmMantTelaComb
''                frmMantTelaComb.Caption = "COMBINACIONES DE TELA:" & rsgrid("Cod_Tela") & " " & rsgrid("Des_Tela")
''                frmMantTelaComb.Codigo_tela = rsgrid("Cod_Tela")
''                frmMantTelaComb.txtdes_tela = rsgrid("Des_Tela")
''                'frmMantTelaComb.sCod_FamTela = rsgrid("Cod_FAMTELA")
''                frmMantTelaComb.sCod_FamTela = UCase(Mid(cboCod_FamTela, 1, 2))
''                frmMantTelaComb.scod_Cliente = IIf(IsNull(vCod_Cliente), "", vCod_Cliente)
''                frmMantTelaComb.sCod_Temcli = txttemporada.Text
''                frmMantTelaComb.CARGA_GRID
''                If rsgrid("Cod_famTela") = "DE" Then
''                    frmMantTelaComb.fraDE.Visible = True
''                Else
''                    frmMantTelaComb.fraDE.Visible = False
''                    frmMantTelaComb.fraGnrl.Top = 500
''                End If
''                frmMantTelaComb.Show 1
''            Else
''                MsgBox ("Debe seleccionar una Tela para acceder a esta opcion")
''            End If
'        Case "IMPRESION"
''            If opcion <> 3 Then
''                Call MsgBox("Esta opción solo es permitida si la busqueda es por cliente", vbCritical)
''                Exit Sub
''            End If
'
'''
'''            If Not rsgrid.EOF Then
'''
'''                strSQL = "SELECT COD_FAMTELA FROM TX_TELA WHERE COD_TELA = '" & rsgrid("Cod_Tela") & "'"
'''                If DevuelveCampo(strSQL, cConnect) = "DE" Then
'''                  strSQL = "SELECT FLG_STATUS_DESARROLLO FROM TX_TELA WHERE COD_TELA = '" & rsgrid("Cod_Tela") & "'"
'''                  If DevuelveCampo(strSQL, cConnect) = "P" Then
'''                    If MsgBox("desea cambiar de estado a ENVIADO", vbYesNo, "ADVERTENCIA") = vbYes Then
'''                      strSQL = "TX_DESARROLLO_CAMBIA_ESTADO_TELA_A_ENVIADO '" & rsgrid("Cod_Tela") & "','" & ComputerName & "','" & vusu & "'"
'''                      ExecuteCommandSQL cConnect, strSQL
'''                    End If
'''                  End If
'''                End If
'''                'Esta sentecia es para obtener el Codigo de Cliente
'''
'''                strSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtcliente.Text) & "'"
'''
'''                Dim oo As Object
'''                On Error GoTo AceptarErr
'''                Set oo = CreateObject("excel.application")
'''                oo.workbooks.Open vRuta & "\RptTelas.xlt"
'''                oo.Visible = True
'''                oo.run "Reporte", txtcodtela.Text, DevuelveCampo(strSQL, cConnect), txttemporada.Text, cConnect
'''                Screen.MousePointer = vbNormal
'''                oo.Visible = True
'''                Set oo = Nothing
'''                'MsgBox ("Aqui ira el reporte")
'''            Else
'''                MsgBox ("Debe seleccionar un item para acceder a esta opcion")
'''            End If
'''
'
'
'        Case "MEDIDA"
'            Load FrmMantMed
'            FrmMantMed.Cod_Item = rsgrid("Cod_Tela")
'            FrmMantMed.Tipo_Item = "T"
'            FrmMantMed.Datos "V", FrmMantMed.Tipo_Item
'            FrmMantMed.Show 1
'            Set FrmMantMed = Nothing
'        Case "PROC"
'            Load frmProcesos
'            frmProcesos.Show 1
'            Set frmProcesos = Nothing
'            'EjecutaOpcion1 "mnuProceso", vper, vemp
'        Case "IMPMAESTRO"
'            Load frmSelecFamilias
'            Set frmSelecFamilias.oParent = Me
'            frmSelecFamilias.CARGA_FAMILIAS
'            frmSelecFamilias.Show 1
'            If Me.varCancelImpresion = 0 Then ReporteMaestro
'        Case "MERMASESP"
'            If Not rsgrid.EOF Then
'                Load frmMerma
'                frmMerma.varCod_Tela = rsgrid("Cod_Tela")
'                frmMerma.TxtTela.Text = Trim(rsgrid("Cod_Tela")) & "-" & Trim(rsgrid("Des_Tela"))
'                strSQL = DevuelveCampo("select merma_tejeduria from tx_telas_mermas where cod_tela='" & rsgrid("Cod_Tela") & "'", cConnect)
'                If Trim(strSQL) <> "" Then frmMerma.TxtMer_Tejeduria.Text = strSQL
'                strSQL = DevuelveCampo("select merma_tinto from tx_telas_mermas where cod_tela='" & rsgrid("Cod_Tela") & "'", cConnect)
'                If Trim(strSQL) <> "" Then frmMerma.TxtMer_Tintoreria.Text = strSQL
'                frmMerma.Show 1
'            End If
'        Case "RAPPORT"
'            frmShowTX_Rapport.Show vbModal
'        Case "STATUS"
'            If Not rsgrid.EOF Then
'                Call CAMBIA_STATUS
'            End If
'    End Select
'
'Exit Sub
'AceptarErr:
'    ErrorHandler err, "Aceptar"
'    Screen.MousePointer = vbNormal
'    Set oo = Nothing
'End Sub

Sub ReporteMaestro()
On Error GoTo hand
    Dim oo As Object
    Dim StrSQL As String
    Dim RutaLogo As String
    
    StrSQL = "SELECT  Ruta_Logo FROM SEGURIDAD..SEG_EMPRESAS WHERE COD_EMPRESA ='" & vemp & "'"
    RutaLogo = DevuelveCampo(StrSQL, cConnect)

    Screen.MousePointer = 11
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\MaestroTelas.xlt"
    'oo.workbooks.Open "C:\Archivos de programa\Gestion de Pedidos\MaestroTelas.xlt"
    oo.Visible = True
    oo.run "Reporte", cConnect, varCadena_Familias, RutaLogo
    Screen.MousePointer = vbNormal
    oo.Visible = True
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler err, "ReporteMaestro"
    Screen.MousePointer = vbNormal
    Set oo = Nothing
End Sub
'
'Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
'If UCase(Mid(cboCod_FamTela, 1, 2)) = "TW" Then
'    Load FrmTela_MtsTwillxHora
'    FrmTela_MtsTwillxHora.sCod_Tela = Trim(rsgrid("Cod_Tela"))
'    FrmTela_MtsTwillxHora.LblCod_Tela = Trim(rsgrid("Cod_Tela"))
'    FrmTela_MtsTwillxHora.LblDes_Tela = Trim(rsgrid("Des_Tela"))
'    FrmTela_MtsTwillxHora.TxtTwill.Text = Trim(rsgrid("Mts_Twill_x_Hora"))
'    FrmTela_MtsTwillxHora.Show 1
'    Set FrmTela_MtsTwillxHora = Nothing
'    Call CARGA_GRID
'    Me.CargaDatos
'Else
'    MsgBox "Opcion válida sólo para Twill", vbCritical
'End If
'End Sub

Private Sub FunctButt3_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    Call Modifica_Comercial
Case "CANCELAR"
    FraComercial.Visible = False
    FraLista.Enabled = True
End Select
End Sub

'Private Sub FunctButt4_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
'Select Case ActionName
'Case "ACEPTAR"
'    Call Modifica_TipoCorte
'Case "CANCELAR"
'    FraFamTipoCorte.Visible = False
'    FraLista.Enabled = True
'End Select
'End Sub

'Private Sub FunctCambios_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
'Dim Estado As String
'Select Case ActionName
'    Case "NOOPERATIVAS"
'        Call NO_OPERATIVAS
'    Case "MODIFICARDATOS"
'        If Not rsgrid.EOF And Not rsgrid.BOF Then
'            strSQL = "UP_VALIDA_MODIFICACION_DATOS_TELA '" & rsgrid("cod_tela") & "','" & vusu & "'"
'            Estado = DevuelveCampo(strSQL, cConnect)
'            If Estado = "" Then
'            Else
'                If Estado = "P" Then
'                    Load FrmCambiosTelas
'                    With FrmCambiosTelas
'                    '.aCod_Famtela = codF_amtela
'                    Call guarda_variables
'                    .TxtGramaje = Trim(rsgrid("Gramaje_Acab"))
'                    .TxtAncho = Trim(rsgrid("Ancho_Acab"))
'                    .txtcod_tela = Trim(rsgrid("Cod_Tela"))
'                    .txtdes_tela = Trim(rsgrid("des_tela"))
'                    .Show 1
'                    CARGA_GRID
'                    End With
'                    Set FrmCambiosTelas = Nothing
'                End If
'            End If
'        End If
'    Case "AUTORIZAR"
'        If Not rsgrid.EOF And Not rsgrid.BOF Then
'            Load FrmSolicitud_CambiosTelas
'            With FrmSolicitud_CambiosTelas
'                .txtcod_tela = Trim(rsgrid("Cod_Tela"))
'                .txtdes_tela = Trim(rsgrid("des_tela"))
'                .Show 1
'                CARGA_GRID
'            End With
'            Set FrmSolicitud_CambiosTelas = Nothing
'        End If
'    Case "MUESTRAAUTORIZACION"
'        If Not rsgrid.EOF And Not rsgrid.BOF Then
'            Load FrmMuestraHistoricoCambioTelas
'            FrmMuestraHistoricoCambioTelas.txtcod_tela = rsgrid("cod_tela")
'            FrmMuestraHistoricoCambioTelas.txtdes_tela = rsgrid("Des_tela")
'            FrmMuestraHistoricoCambioTelas.CARGA_GRID
'            FrmMuestraHistoricoCambioTelas.Show 1
'            Set FrmMuestraHistoricoCambioTelas = Nothing
'        End If
'    Case "BITACORA"
'        If Not rsgrid.EOF And Not rsgrid.BOF Then
'            Load FrmTelaBitacora
'            FrmTelaBitacora.txtcod_tela = rsgrid("cod_tela")
'            FrmTelaBitacora.txtdes_tela = rsgrid("Des_tela")
'            FrmTelaBitacora.CARGA_GRID
'            FrmTelaBitacora.Show 1
'            Set FrmTelaBitacora = Nothing
'        End If
'    Case "DATOSCRUDO"
'        If Not rsgrid.EOF And Not rsgrid.BOF Then
'            Load FrmModDatosCrudo
'            FrmModDatosCrudo.sCod_Tela = rsgrid("cod_tela")
'            FrmModDatosCrudo.TxtGramajeCRudo = CDbl(rsgrid("Gramaje_Crudo"))
'            FrmModDatosCrudo.TxtAnchoCrudo = CDbl(rsgrid("Ancho_Crudo"))
'            FrmModDatosCrudo.Show vbModal
'            Set FrmModDatosCrudo = Nothing
'            CARGA_GRID
'        End If
'    Case "POSTTENIDO"
'        If Not rsgrid.EOF And Not rsgrid.BOF Then
'            If rsgrid("cod_famtela") = "DE" Then
'                MsgBox "Debe ingresar los datos por Combinacion", vbCritical
'                Exit Sub
'            End If
'            Load FrmManTelasDatTec
'            FrmManTelasDatTec.sCod_Tela = rsgrid("cod_tela")
'            FrmManTelasDatTec.sDes_tela = rsgrid("des_tela")
'            FrmManTelasDatTec.sFamite = rsgrid("cod_famtela")
'            FrmManTelasDatTec.CARGA_DATOS
'            FrmManTelasDatTec.Show 1
'            Set FrmManTelasDatTec = Nothing
'        End If
'    Case "PRUEBA"
'        If Not rsgrid.EOF And Not rsgrid.BOF Then
'            If rsgrid("cod_famtela") = "DE" Then
'                MsgBox "Debe ingresar los datos por Combinacion", vbCritical
'                Exit Sub
'            End If
'            Load FrmManTelasDatTecAdd
'            FrmManTelasDatTecAdd.sCod_Tela = rsgrid("cod_tela")
'            FrmManTelasDatTecAdd.sFamite = rsgrid("cod_famtela")
'            FrmManTelasDatTecAdd.CARGA_DATOS
'            FrmManTelasDatTecAdd.Show 1
'            Set FrmManTelasDatTecAdd = Nothing
'        End If
'    Case "ALTERNATIVAS"
'        Load FrmShowAlternativasPesoAncho
'        FrmShowAlternativasPesoAncho.vCod_Tela = rsgrid("cod_tela")
'        FrmShowAlternativasPesoAncho.CARGA_GRID
'        FrmShowAlternativasPesoAncho.Show vbModal
'        Set FrmShowAlternativasPesoAncho = Nothing
'    Case "COMERCIAL"
'        FraComercial.Visible = True
'        TxtGramaje_Comercial.Text = rsgrid("Gramaje_Comercial")
'        TxtAncho_Comercial.Text = rsgrid("Ancho_Comercial")
'        TxtGramaje_Comercial.SetFocus
'        FraLista.Enabled = False
'    Case "REVISION"
'        Load FrmVerRevisionesTelas
'        FrmVerRevisionesTelas.vCod_Tela = Trim(rsgrid("Cod_Tela"))
'        FrmVerRevisionesTelas.vDes_Tela = Trim(rsgrid("Des_Tela"))
'        FrmVerRevisionesTelas.CARGA_GRID
'        FrmVerRevisionesTelas.Show vbModal
'        Set FrmVerRevisionesTelas = Nothing
'        Call CARGA_GRID
'    Case "RELACIONADOS"
'        Load FrmVerArticulosRelacionados
'        FrmVerArticulosRelacionados.vCod_Tela = Trim(rsgrid("Cod_Tela"))
'        FrmVerArticulosRelacionados.vDes_Tela = Trim(rsgrid("Des_Tela"))
'        FrmVerArticulosRelacionados.CARGA_GRID
'        FrmVerArticulosRelacionados.Show vbModal
'        Set FrmVerArticulosRelacionados = Nothing
'    Case "RUTAS"
'        Load FrmShowTelas
'        FrmShowTelas.vCod_Tela = rsgrid("cod_tela")
'        FrmShowTelas.vDes_Tela = rsgrid("des_tela")
'        FrmShowTelas.vCOD_ORDTRA = TxtCod_OrdTra.Text
'        FrmShowTelas.vFamite = rsgrid("cod_famtela")
'        FrmShowTelas.CARGA_GRID
'        FrmShowTelas.Show vbModal
'        Set FrmShowTelas = Nothing
'End Select
'End Sub
'
'Private Sub FunctTemporada_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
'Dim rstAux As ADODB.Recordset
'    Select Case ActionName
'        Case "TEMPORADA"
'            If Not rsgrid.EOF Then
'                Load frmMantTelaTemCli
'                frmMantTelaTemCli.Codigo_tela = rsgrid("Cod_Tela")
'                frmMantTelaTemCli.CARGA_DATOS
'                'frmMantTelaTemCli.CARGA_COMBOS
'                frmMantTelaTemCli.Show 1
'            Else
'                MsgBox ("Debe seleccionar un item para acceder a esta opcion")
'            End If
'        Case "TICKET"
'        Dim CodHilado As String
'            If Not rsgrid.EOF Then
'                Set rstAux = New ADODB.Recordset
'                rstAux.ActiveConnection = cConnect
'                rstAux.CursorType = adOpenForwardOnly
'                rstAux.CursorLocation = adUseClient
'                rstAux.LockType = adLockReadOnly
'
'                strSQL = "EXEC SM_TRAE_DATOS_STICKER_TELA '" & rsgrid("Cod_Tela") & "'"
'                rstAux.Open strSQL
'
'                Load frmTicket
'                With frmTicket
'                    .TxtCodigo = rsgrid("Cod_Tela")
'                    .vCod_FamTela = rsgrid("Cod_FamTela")
'                    If .vCod_FamTela = "DE" Then
'                        .FraComb.Visible = True
'                        '.txtPeso2 = Trim(rsgrid("gramaje_despues_lavado"))
'                    Else
'                        .Frame2.Visible = True
'                        .TxtPeso2 = Trim(rsgrid("Gramaje_despuesLavado"))
'                    End If
'
'                    .TxtDescripcion = DevuelveCampo("SELECT des_famtela FROM TX_FamTela where cod_famtela='" & rsgrid("Cod_FamTela") & "'", cConnect)
'                    .txtComposicion = rstAux!Composicion
'                    .TxtHilado = rstAux!Yarn
'                    .TxtEncogAncho = Trim(rsgrid("Encog_Ancho"))
'                    .TxtEncogLargo = Trim(rsgrid("Encog_Largo"))
'                    .TxtMetodoTen = ""
'                    .TxtProcesoLav = ""
'                    .TxtcolorWay = ""
'                    .TxtGalga = DevuelveCampo("SELECT rtrim(Des_Galga) FROM TX_GALGA where cod_galga='" & rsgrid("Cod_Galga") & "'", cConnect)
'                    .TxtDiamGalga = Trim(rsgrid("diametro"))
'                    .TxtAncho = Trim(rsgrid("Ancho_Comercial"))
'                    .TxtAnchoPulg = CDbl(Trim(rsgrid("Ancho_Comercial"))) / 0.0254
'                    'CodHilado = DevuelveCampo("SELECT cod_hiltel FROM Tx_HilosTel WHERE Cod_Tela = '" & rsgrid("Cod_Tela") & "' AND num_secuencia='01'", cCONNECT)
'                    '.TxtHilado = DevuelveCampo("select des_hiltel from it_hilado where cod_hiltel='" & CodHilado & "'", cCONNECT)
'                    .txtPeso = Trim(rsgrid("Gramaje_Comercial"))
'                    .Show 1
'                End With
'                Set frmTicket = Nothing
'            End If
'        Case "ESTCOMP"
'            If rsgrid.EOF Then Exit Sub
'
'            frmEstCompTela.LblCod_Tela = rsgrid("Cod_Tela")
'            frmEstCompTela.LblDes_Tela = rsgrid("Des_Tela")
'            frmEstCompTela.SM_ESTILOS_COMPONENTES_POR_TELA
'            frmEstCompTela.Show vbModal
'        Case "CONDTRABAJO"
'            If rsgrid.EOF Then Exit Sub
'
'                Set rstAux = New ADODB.Recordset
'                rstAux.ActiveConnection = cConnect
'                rstAux.CursorType = adOpenForwardOnly
'                rstAux.CursorLocation = adUseClient
'                rstAux.LockType = adLockReadOnly
'                rstAux.Open "UP_MAN_TX_TELA 'I','" & rsgrid("COD_TELA") & "',0,0,0,'','',0,0,0,0,'','',0,0", cConnect, 3, 3
'
'                Load FrmCondicionesTrabajoxArticulo
'                With FrmCondicionesTrabajoxArticulo
'                    .txtFamTela = rsgrid("cod_famtela") 'rstAux("COD_FAMTELA") 'cboCod_FamTela.Text
'                    .TxtArticulo = rstAux("Cod_tela")
'                    .TxtDescripcion = rstAux("DES_TELA") 'txtdestela
'                    .TxtDensEsta = txtGramaje_Acab.Text
'                    .TxtAncho = Me.txtEncog_Ancho.Text
'                    .TxtLargo = Me.txtEncog_Largo.Text
'                    .TxtDensReq = IIf(IsNull(rstAux("Densidad_Requerida")), 0, rstAux("Densidad_Requerida"))
'                    .TxtAnchoEsta = IIf(IsNull(rstAux("ANCHO_ACAB")), 0, rstAux("ANCHO_ACAB")) 'txtAncho_Acab.Text
'                    .TxtLargo1 = IIf(IsNull(rstAux("enc_largo_esperado_desde")), 0, rstAux("enc_largo_esperado_desde"))
'                    .TxtLargo2 = IIf(IsNull(rstAux("enc_largo_esperado_hasta")), 0, rstAux("enc_largo_esperado_hasta"))
'                    .TxtAncho1 = IIf(IsNull(rstAux("enc_ancho_esperado_desde")), 0, rstAux("enc_ancho_esperado_desde"))
'                    .TxtAncho2 = IIf(IsNull(rstAux("enc_ancho_esperado_hasta")), 0, rstAux("enc_ancho_esperado_hasta"))
'                    .txtPorcentaje = IIf(IsNull(rstAux("por_merma_manufactura")), 0, rstAux("por_merma_manufactura"))
'                    .TxtDesPorcentaje = DevuelveCampo("select descripcion from tx_porcentajes where por_merma_manufactura='" & rsgrid("por_merma_manufactura") & "'", cConnect)
'                    .TxtRevirado = rstAux("Revirado")
'                    If rstAux("Condicion_de_llegada_de_tela") = "T" Then
'                        .OptTubular = True
'                    ElseIf rstAux("condicion_de_llegada_de_tela") = "A" Then
'                        .OptAbierta = True
'                    End If
'
'                    If rstAux("Tipo_de_Tejido") = "S" Then
'                        .OptSinSentido = True
'                    ElseIf rstAux("Tipo_de_tejido") = "C" Then
'                        .OptConSentido = True
'                    End If
'                    If rstAux("Merma_orillos") = "S" Then
'                        .OptSinOrillos = True
'                    ElseIf rstAux("Merma_orillos") = "A" Then
'                        .OptOrillosAguja = True
'                    ElseIf rstAux("merma_orillos") = "E" Then
'                        .OptOrillosEngomados = True
'                    End If
'
'                    If IIf(IsNull(rstAux("Grado_Linea_doblez")), " ", rstAux("grado_linea_doblez")) = "N" Or IIf(IsNull(rsgrid("Grado_Linea_doblez")), " ", rstAux("grado_linea_doblez")) = " " Then
'                        .OptNinguna = True
'                    ElseIf rstAux("Grado_linea_doblez") = "M" Then
'                        .OptMangas = True
'                    ElseIf rstAux("Grado_linea_doblez") = "A" Then
'                        .OptAmbas = True
'                    End If
'
'                    If IIf(IsNull(rstAux("inclinacion_trama")), "0", rstAux("inclinacion_trama")) = "0" Or IIf(IsNull(rsgrid("inclinacion_trama")), "0", rsgrid("inclinacion_trama")) = " " Then
'                        .Opt0 = True
'                    ElseIf rstAux("inclinacion_trama") = "1" Then
'                        .Opt1 = True
'                    End If
'                    Set rstAux = Nothing
'                    .Show 1
'                        Call CARGA_GRID
'                    Me.CargaDatos
'                End With
'        Case "NPS"
'            Load FrmNpsDondeUsaTela
'            FrmNpsDondeUsaTela.txtcod_tela.Text = rsgrid("Cod_Tela")
'            FrmNpsDondeUsaTela.txtdes_tela.Text = rsgrid("Des_Tela")
'            FrmNpsDondeUsaTela.CARGA_GRID
'            FrmNpsDondeUsaTela.Show 1
'            Set FrmNpsDondeUsaTela = Nothing
'        Case "HOJARUTA"
'            If rsgrid.EOF Then Exit Sub
'            Call Hoja_Ruta
'        Case "SECUENCIA"
'        If rsgrid.EOF Then Exit Sub
'        If UCase(rsgrid("Cod_FamTela")) <> "DE" Then
'            Load FrmManTela_Procesos_Textil
'            FrmManTela_Procesos_Textil.vCod_Tela = rsgrid("Cod_Tela")
'            FrmManTela_Procesos_Textil.txtcod_tela.Text = rsgrid("Cod_Tela")
'            FrmManTela_Procesos_Textil.txtdes_tela.Text = rsgrid("Des_Tela")
'            FrmManTela_Procesos_Textil.CARGA_GRID
'            FrmManTela_Procesos_Textil.Show vbModal
'            Set FrmManTela_Procesos_Textil = Nothing
'        Else
'            MsgBox "Debe ingresar por Combinaciones", vbCritical
'            Exit Sub
'            'Load FrmManTelaComb_Procesos_Textil
'            'FrmManTelaComb_Procesos_Textil.vCod_Tela = rsgrid("Cod_Tela")
'            'FrmManTelaComb_Procesos_Textil.vCod_Comb = Me.sCod_Comb
'            'FrmManTelaComb_Procesos_Textil.txtcod_tela.Text = rsgrid("Cod_Tela")
'            'FrmManTelaComb_Procesos_Textil.txtdes_tela.Text = rsgrid("Des_Tela")
'            'FrmManTelaComb_Procesos_Textil.TxtCod_Comb.Text = Me.sCod_Comb
'            'FrmManTelaComb_Procesos_Textil.TxtDes_Comb.Text = Me.sDes_Comb
'
'            'FrmManTelaComb_Procesos_Textil.CARGA_GRID
'            'FrmManTelaComb_Procesos_Textil.Show vbModal
'            'Set FrmManTelaComb_Procesos_Textil = Nothing
'        End If
'
'    End Select
'End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Eliminar As Integer
On Error GoTo fin

    Select Case ActionName
        Case "ADICIONAR"
            sTipo = "I"
            FraBuscar.Enabled = False
            LIMPIAR_DATOS
            HABILITA_DATOS
            'txtGramaje_Acab.Enabled = True
            'txtAncho_Acab.Enabled = True
            BuscaCombo "N", 1, cboFlg_Operatividad
            txtcodtela.Enabled = False
            If txtdestela.Enabled = True Then
                txtdestela.SetFocus
            End If
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
            varCod_Tela = ""
        Case "MODIFICAR"
            sTipo = "U"
            FraBuscar.Enabled = False
            HABILITA_DATOS
            txtcodtela.Enabled = False
            cboCod_FamTela.Enabled = False
            If txtdestela.Enabled = True Then
                txtdestela.SetFocus
            End If
            varCod_Tela = rsgrid("Cod_Tela")
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
            'varCod_Tela = rsgrid("Cod_Tela").Value
        Case "ELIMINAR"
            Eliminar = MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?.", vbInformation + vbYesNo, "Telas")
            If Eliminar = vbYes Then
                sTipo = "D"
                Call ELIMINAR_DATOS
                Call CARGA_GRID
            End If
        Case "GRABAR"
            If VALIDA_DATOS Then
                'varCod_Tela = rsgrid("Cod_Tela").Value
                If Not SALVAR_DATOS Then Exit Sub
                
                CargaDatos
                HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                DGridLista.Enabled = True
                'fraoptions.Enabled = False
                FraBuscar.Enabled = True
                
                Call CARGA_GRID
                If sTipo = "U" Then
                    Call BuscaCampo(rsgrid, "Cod_Tela", varCod_Tela)
                End If
                sTipo = ""
            End If
        Case "DESHACER"
            INHABILITA_DATOS
            sTipo = ""
            LIMPIAR_DATOS
            CargaDatos
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            DGridLista.Enabled = True
            'fraoptions.Enabled = False
            FraBuscar.Enabled = True
        Case "SALIR"
            sTipo = ""
            Unload Me
    End Select
    Exit Sub
fin:
MsgBox "Inconvenientes para realizar cambios " + err.Description, vbCritical + vbOKOnly, "Mensaje"
    
End Sub

Private Sub optcliente_Click()
                    
    txtcliente.Text = ""
    txtNom_Cliente.Text = ""
    txttemporada.Text = ""
    txtNom_TemCli.Text = ""
    Frafamilia.Visible = False
    Fratela.Visible = False
    Fracliente.Visible = True

    Opcion = 3
    txtcliente.SetFocus
    'HabilitaMant Me.FunctButt1, "IMPRESION/DESARROLLO/LISTADO/PROCESOS/COMPOSICION/COMBINACIONES"
    Call CARGA_GRID
End Sub

Private Sub optfamtela_Click()
    txtcliente = "": txttemporada = ""
    txtfamilia.Text = ""
    txtdes_familia.Text = ""
    txtgrupo.Text = ""
    txtdes_famgruite.Text = ""
    
    txtgrupo.Enabled = False
    cmdBusgrupo.Enabled = False
    
    Frafamilia.Visible = True
    Fratela.Visible = False
    Fracliente.Visible = False
    
    Opcion = 1
    txtfamilia.SetFocus
    'HabilitaMant Me.FunctButt1, "DESARROLLO/LISTADO/PROCESOS/COMPOSICION/COMBINACIONES"
   
    Call CARGA_GRID
End Sub

Private Sub optFlg_Operatividad_Click()
    txtcod_tela.Text = ""
    txtdes_tela.Text = ""
    txtcliente = "": txttemporada = ""
    Frafamilia.Visible = False
    Fratela.Visible = False
    Fracliente.Visible = False
    Opcion = 4
End Sub

Private Sub OptTela_Click()
    txtcod_tela.Text = ""
    txtdes_tela.Text = ""
    txtcliente = "": txttemporada = ""
    Frafamilia.Visible = False
    Fratela.Visible = True
    Fracliente.Visible = False

    Opcion = 2
   
    'HabilitaMant Me.FunctButt1, "DESARROLLO/LISTADO/PROCESOS/COMPOSICION/COMBINACIONES"
    txtcod_tela.SetFocus
    Call CARGA_GRID
End Sub



Private Sub txtAncho_Acab_KeyPress(KeyAscii As Integer)
    SoloNumeros txtAncho_Acab, KeyAscii, True, 2, 4
End Sub

Private Sub txtAncho_Acab_LostFocus()
    If txtAncho_Acab.Text = "" Then
        txtAncho_Acab.Text = 0
    End If
    txtAncho_Acab_Abierto.Text = txtAncho_Acab.Text * 2
End Sub

Private Sub TxtAncho_Comercial_GotFocus()
SelectionText TxtAncho_Comercial
End Sub

Private Sub TxtAncho_Comercial_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FunctButt3.SetFocus
Else
    Call SoloNumeros(TxtAncho_Comercial, KeyAscii, True, 2)
End If
End Sub

Private Sub txtAncho_Crudo_KeyPress(KeyAscii As Integer)
    SoloNumeros txtAncho_Crudo, KeyAscii, True, 2, 4
End Sub

Private Sub txtAncho_Crudo_LostFocus()
    If Trim(txtAncho_Crudo.Text) = "" Then
        txtAncho_Crudo.Text = 0
    End If
End Sub



Private Sub txtcliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtcliente.Text) = "" Then
            cmdBusCliente_Click
        Else
            StrSQL = "SELECT Nom_Cliente FROM TG_CLIENTE WHERE Abr_Cliente LIKE '" & Trim(txtcliente.Text) & "%'"
            txtNom_Cliente.Text = DevuelveCampo(StrSQL, cConnect)
            txttemporada.Enabled = True
            txtNom_TemCli.Enabled = True
            txttemporada.SetFocus
        End If
    End If
End Sub

Private Sub TxtCod_OrdTra_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_Tel_Origen) <> "" Then
            TxtCod_OrdTra.Text = Format(TxtCod_OrdTra, "00000")
            SendKeys "{TAB}"
        End If
    End If
End Sub

Private Sub TxtCod_OrdTra_Tejeduria_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(txtCod_Tel_Origen) <> "" Then
        TxtCod_OrdTra_Tejeduria.Text = Format(TxtCod_OrdTra_Tejeduria, "00000")
        SendKeys "{TAB}"
    End If
End If
End Sub

Private Sub txtCod_Tel_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
      BuscaTela 1
      FillCombo
      cboCombo.SetFocus
  End If
End Sub
Public Sub FillCombo()
On Error GoTo fin
Dim rstAux As ADODB.Recordset
    
    StrSQL = "SELECT Cod_Comb, Des_Comb FROM TX_TELACOMB " & _
             "WHERE Cod_Tela = '" & txtCod_Tel_Origen & "'"
    Set rstAux = CargarRecordSetDesconectado(StrSQL, cConnect)
    
    cboCombo.Clear
    With rstAux
    If .RecordCount > 0 Then .MoveFirst
    Do Until .EOF
        cboCombo.AddItem !Cod_Comb & " " & !Des_Comb
        .MoveNext
    Loop
    .Close
    End With
    Set rstAux = Nothing
Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, "Cargar Combinaciones"
End Sub

Private Sub BuscaTela(Opcion As Integer)

On Error GoTo fin

Dim rstAux As ADODB.Recordset
    
    StrSQL = "SELECT Cod_Tela as Codigo, Des_Tela as Descripcion FROM TX_TELA WHERE "
    
        
    txtCod_Tel_Origen = Trim(txtCod_Tel_Origen)
    txtCod_Tel_Origen.Text = CompletaCodigo(Trim(txtCod_Tel_Origen.Text), 8, 2)
    txtDes_Tel_Origen = Trim(txtDes_Tel_Origen)
    Select Case Opcion
    Case 1: StrSQL = StrSQL & "Cod_Tela like '%" & txtCod_Tel_Origen & "%'"
    Case 2: StrSQL = StrSQL & "Des_Tela like '%" & txtDes_Tel_Origen & "%'"
    End Select
    txtCod_Tel_Origen = ""
    txtDes_Tel_Origen = ""
    With frmBusqGeneral3
        Set .oParent = Me
        .SQuery = StrSQL
        .Cargar_Datos

'        .DGridLista.Columns("").Width = 1000
        
        CODIGO = ".."
        
        'Set rstAux = .DGridLista.ADORecordset
        Set rstAux = .gexLista.ADORecordset
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If CODIGO <> "" And rstAux.RecordCount > 0 Then
            txtCod_Tel_Origen = Trim(rstAux!CODIGO)
            txtDes_Tel_Origen = Trim(rstAux!Descripcion)
            SendKeys "{TAB}"
        Else
            SendKeys "{TAB}"
        End If
        
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    
Exit Sub
Resume
fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda "
End Sub


Private Sub TxtCod_TelaOriginal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxtCod_TelaOriginal.Text) = "" Then
        MsgBox ("Sirvase ingresar un codigo de Item")
    Else
        TxtCod_TelaOriginal.Text = CompletaCodigo(Trim(TxtCod_TelaOriginal.Text), 8, 2)
        StrSQL = "SELECT Des_Tela FROM TX_TELA WHERE Cod_Tela='" & TxtCod_TelaOriginal.Text & "'"
        TxtDes_TelaOriginal.Text = DevuelveCampo(StrSQL, cConnect)
        TxtSufijo.SetFocus
    End If
End If
End Sub


Private Sub txtColumnas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(Me.txtColumnas, KeyAscii, False, 2)
End If
End Sub

Private Sub txtCunSas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(Me.txtCunSas, KeyAscii, False, 2)
End If
End Sub

Private Sub txtdes_famgruite_KeyPress(KeyAscii As Integer)
    Dim StrSQL As String
    If KeyAscii = 13 Then
        If Len(Trim(txtdes_famgruite.Text)) < 5 Then
            Call MsgBox("El Texto Ingresado debe contar con un mínimo de 5 caracteres", vbInformation)
        Else
            StrSQL = "SELECT Cod_Gruitem FROM LG_FamGruIte WHERE Cod_Famitem='" & Trim(txtfamilia.Text) & "' AND des_famgruite LIKE '" & Trim(txtgrupo.Text) & "%'"
            txtgrupo.Text = DevuelveCampo(StrSQL, cConnect)
        End If
    End If

End Sub
Private Sub txtdes_familia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtgrupo.Text = ""
        If Len(Trim(txtdes_familia.Text)) < 5 Then
            MsgBox ("El Texto Ingresado debe contar con un mínimo de 5 caracteres")
        Else
            StrSQL = "SELECT Cod_FamTela FROM TX_FAMTELA WHERE  Des_FamTela LIKE '" & txtdes_familia.Text & "%'"
            txtfamilia.Text = DevuelveCampo(StrSQL, cConnect)
            
            txtgrupo.Enabled = True
            cmdBusgrupo.Enabled = True
            
        End If
    End If
End Sub

Private Sub txtDes_Tel_Origen_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
      BuscaTela 2
      FillCombo
      cboCombo.SetFocus
  End If
End Sub

Private Sub txtdes_tela_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtdes_tela.Text) = "" Then
             MsgBox ("Sirvase ingresar una Descripcion del Item")
        Else
            'Esta consulta es para obtener el Codigo de Cliente
            'strSQL = "SELECT Cod_Tela FROM TX_TELA WHERE Des_Tela LIKE '" & Trim(TxtDes_Tela.Text) & "%'"
            'TxtCod_Tela.Text = DevuelveCampo(strSQL, cCONNECT)
            cmdBusTela_Click
        End If
        Call CARGA_GRID
    End If
End Sub

Private Sub TxtDes_TipoCorte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_FamiliaTipoCorte(2)
End If
End Sub

Private Sub txtdiametro_KeyPress(KeyAscii As Integer)
    SoloNumeros txtdiametro, KeyAscii, False, 0, 4
End Sub

Private Sub txtdiametro_LostFocus()
    If txtdiametro.Text = "" Then
        txtdiametro.Text = 0
    End If
End Sub

Private Sub txtEncog_Ancho_KeyPress(KeyAscii As Integer)
    If KeyAscii = 45 And Val(txtEncog_Ancho.Text) <> 0 Then
        If Mid(txtEncog_Ancho.Text, 1, 1) <> "-" Then
            If Val(txtEncog_Ancho.Text) <> 0 Then
                txtEncog_Ancho.Text = "-" & txtEncog_Ancho.Text
                If Mid(txtEncog_Ancho.Text, 1, 2) <> "--" Then
                    txtEncog_Ancho.Text = Right(txtEncog_Ancho.Text, Len(txtEncog_Ancho.Text) - 1)
                End If
            Else
                Exit Sub
            End If
        Else
            KeyAscii = 0
            'txtEncog_Largo.Text = Right(txtEncog_Largo.Text, Len(txtEncog_Largo.Text) - 1)
        End If
        'Exit Sub
    Else
        SoloNumeros txtEncog_Ancho, KeyAscii, True, 4, 4
    End If
End Sub

Private Sub txtEncog_Ancho_LostFocus()
    If txtEncog_Ancho.Text = "" Or Trim(txtEncog_Ancho.Text) = "-" Then
        txtEncog_Ancho.Text = 0
    End If
End Sub

Private Sub txtEncog_Ancho_Vap_KeyPress(KeyAscii As Integer)
    If KeyAscii = 45 And Val(txtEncog_Ancho_Vap.Text) <> 0 Then
        If Mid(txtEncog_Ancho_Vap.Text, 1, 1) <> "-" Then
            If Val(txtEncog_Ancho_Vap.Text) <> 0 Then
                txtEncog_Ancho_Vap.Text = "-" & txtEncog_Ancho_Vap.Text
                If Mid(txtEncog_Ancho_Vap.Text, 1, 2) <> "--" Then
                    txtEncog_Ancho_Vap.Text = Right(txtEncog_Ancho_Vap.Text, Len(txtEncog_Ancho_Vap.Text) - 1)
                End If
            Else
                Exit Sub
            End If
        Else
            KeyAscii = 0
        End If
    Else
        SoloNumeros txtEncog_Ancho_Vap, KeyAscii, True, 4, 4
    End If
End Sub

Private Sub txtEncog_Ancho_Vap_LostFocus()
    If txtEncog_Ancho_Vap.Text = "" Or Trim(txtEncog_Ancho_Vap.Text) = "-" Then
        txtEncog_Ancho_Vap.Text = 0
    End If
End Sub

Private Sub txtEncog_Largo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 45 And Val(txtEncog_Largo.Text) <> 0 Then
        If Mid(txtEncog_Largo.Text, 1, 1) <> "-" Then
            If Val(txtEncog_Largo.Text) <> 0 Then
                txtEncog_Largo.Text = "-" & txtEncog_Largo.Text
                If Mid(txtEncog_Largo.Text, 1, 2) <> "--" Then
                    txtEncog_Largo.Text = Right(txtEncog_Largo.Text, Len(txtEncog_Largo.Text) - 1)
                End If
            Else
                Exit Sub
            End If
        Else
            KeyAscii = 0
            'txtEncog_Largo.Text = Right(txtEncog_Largo.Text, Len(txtEncog_Largo.Text) - 1)
        End If
        'Exit Sub
    Else
        SoloNumeros txtEncog_Largo, KeyAscii, True, 4, 4
    End If
    
End Sub

Private Sub txtEncog_Largo_LostFocus()
    If txtEncog_Largo.Text = "" Or Trim(txtEncog_Largo.Text) = "-" Then
        txtEncog_Largo.Text = 0
    End If
End Sub

Private Sub txtEncog_Largo_Vap_KeyPress(KeyAscii As Integer)
    If KeyAscii = 45 And Val(txtEncog_Largo_Vap.Text) <> 0 Then
        If Mid(txtEncog_Largo_Vap.Text, 1, 1) <> "-" Then
            If Val(txtEncog_Largo_Vap.Text) <> 0 Then
                txtEncog_Largo_Vap.Text = "-" & txtEncog_Largo_Vap.Text
                If Mid(txtEncog_Largo_Vap.Text, 1, 2) <> "--" Then
                    txtEncog_Largo_Vap.Text = Right(txtEncog_Largo_Vap.Text, Len(txtEncog_Largo_Vap.Text) - 1)
                End If
            Else
                Exit Sub
            End If
        Else
            KeyAscii = 0
        End If
    Else
        SoloNumeros txtEncog_Largo_Vap, KeyAscii, True, 4, 4
    End If
End Sub

Private Sub txtEncog_Largo_Vap_LostFocus()
    If txtEncog_Largo_Vap.Text = "" Or Trim(txtEncog_Largo_Vap.Text) = "-" Then
        txtEncog_Largo_Vap.Text = 0
    End If
End Sub



Private Sub txtGramaje_Acab_KeyPress(KeyAscii As Integer)
    SoloNumeros txtGramaje_Acab, KeyAscii, False, 0, 4
End Sub

Private Sub txtGramaje_Acab_LostFocus()
    If txtGramaje_Acab.Text = "" Then
        txtGramaje_Acab.Text = 0
    End If
End Sub

Private Sub TxtGramaje_Comercial_GotFocus()
SelectionText TxtGramaje_Comercial
End Sub

Private Sub TxtGramaje_Comercial_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtAncho_Comercial.SetFocus
Else
    Call SoloNumeros(TxtGramaje_Comercial, KeyAscii, True, 2)
End If
End Sub

Private Sub txtGramaje_Crudo_KeyPress(KeyAscii As Integer)
    SoloNumeros txtGramaje_Crudo, KeyAscii, False, 0, 4
End Sub

Private Sub txtGramaje_Crudo_LostFocus()
    If Trim(txtGramaje_Crudo.Text) = "" Then
        txtGramaje_Crudo.Text = 0
    End If
End Sub

Private Sub TxtMts_Twill_x_Hora_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(TxtMts_Twill_x_Hora, KeyAscii, True, 2)
End If
End Sub

Private Sub TxtMts_Twill_x_Hora_LostFocus()
If Trim(TxtMts_Twill_x_Hora.Text) = "" Then
    TxtMts_Twill_x_Hora.Text = 0
End If
End Sub

Private Sub txtNom_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(txtNom_Cliente) > 4 Then
            StrSQL = "SELECT Abr_Cliente FROM TG_CLIENTE WHERE Nom_Cliente LIKE '" & Trim(txtNom_Cliente.Text) & "%'"
            txtcliente.Text = DevuelveCampo(StrSQL, cConnect)
        Else
            MsgBox ("El Texto Ingresado debe contar con un mínimo de 5 caracteres")
        End If
    End If
End Sub

Private Sub txtNom_TemCli_KeyPress(KeyAscii As Integer)
    'Esta consulta es para obtener el Codigo de Cliente
    StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtcliente.Text) & "'"
    If KeyAscii = 13 Then
        If Len(txtNom_TemCli) > 4 Then
            'Esta consulta nos permite obtener el Matching entre Cliente y Temporada
            StrSQL = "SELECT Cod_TemCli FROM TG_TEMCLI WHERE Cod_Cliente='" & DevuelveCampo(StrSQL, cConnect) & "' AND Nom_TemCli LIKE '" & Trim(txtNom_TemCli.Text) & "%'"
            txttemporada.Text = DevuelveCampo(StrSQL, cConnect)
        Else
            MsgBox ("El Texto Ingresado debe contar con un mínimo de 5 caracteres")
        End If
    End If
End Sub

Private Sub txtNum_Aguja_KeyPress(KeyAscii As Integer)
    SoloNumeros txtNum_Aguja, KeyAscii, False, 0, 4
End Sub

Private Sub txtNum_Aguja_LostFocus()
    If txtNum_Aguja.Text = "" Then
        txtNum_Aguja.Text = 0
    End If
End Sub

Private Sub TxtNum_Alimentadores_KeyPress(KeyAscii As Integer)
    SoloNumeros txtNum_Alimentadores, KeyAscii, False, 0, 4
End Sub

Private Sub txtNum_Alimentadores_LostFocus()
    If txtNum_Alimentadores.Text = "" Then
        txtNum_Alimentadores.Text = 0
    End If
End Sub

Private Sub txtNum_Lavadas_KeyPress(KeyAscii As Integer)
    SoloNumeros txtNum_Lavadas, KeyAscii, False, 2, 6
End Sub

Private Sub txtNum_Lavadas_LostFocus()
    If txtNum_Lavadas.Text = "" Then
        txtNum_Lavadas.Text = 0
    End If
End Sub

Private Sub txtNum_Rpm_KeyPress(KeyAscii As Integer)
    SoloNumeros txtNum_Rpm, KeyAscii, False, 0, 4
End Sub

Private Sub txtNum_Rpm_LostFocus()
    If txtNum_Rpm.Text = "" Then
        txtNum_Rpm.Text = 0
    End If
End Sub


Private Sub txtPeso_KeyPress(KeyAscii As Integer)
    Call SoloNumeros(txtPeso, KeyAscii, True, 5, 4)
End Sub

Private Sub txtPeso_LostFocus()
    If Trim(txtPeso.Text) = "" Then
        txtPeso.Text = "0"
    End If
End Sub

Private Sub TxtRevirado_KeyPress(KeyAscii As Integer)
   ' Call SoloNumeros(TxtRevirado, KeyAscii, True, 2, 3)
End Sub

Private Sub TxtSufijo_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case Asc("A") To Asc("Z"), Asc("a") To Asc("z")
    Case Else: If KeyAscii = 8 Then Else KeyAscii = 0
End Select
End Sub

Private Sub txttemporada_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txttemporada.Text) = "" Then
            cmdBusTemporada_Click
        Else
            StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtcliente.Text) & "'"
            StrSQL = "SELECT Nom_TemCli FROM TG_TemCli WHERE Cod_Cliente='" & DevuelveCampo(StrSQL, cConnect) & "' AND Cod_TemCli='" & txttemporada.Text & "'"
            txtNom_TemCli.Text = DevuelveCampo(StrSQL, cConnect)
                       
            FunctBuscar.SetFocus
        End If
    End If
End Sub

Private Sub txtcod_tela_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtcod_tela.Text) = "" Then
            MsgBox ("Sirvase ingresar un codigo de Item")
        Else
            txtcod_tela.Text = CompletaCodigo(Trim(txtcod_tela.Text), 8, 2)
            
            'Esta consulta es para obtener el Codigo de Cliente
            StrSQL = "SELECT Des_Tela FROM TX_TELA WHERE Cod_Tela='" & txtcod_tela.Text & "'"
            txtdes_tela.Text = DevuelveCampo(StrSQL, cConnect)
        End If
        Call CARGA_GRID
    End If
End Sub

Private Sub txtfamilia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtgrupo.Text = ""
        If Trim(txtfamilia.Text) = "" Then
            cmdBusFamItem_Click
        Else
            If ValidaFamilia = False Then
                 Exit Sub
            Else
                StrSQL = "SELECT Des_FamTela FROM TX_FAMTELA WHERE Cod_FamTela='" & txtfamilia.Text & "'"
                txtdes_familia.Text = DevuelveCampo(StrSQL, cConnect)
                
                txtgrupo.Enabled = True
                cmdBusgrupo.Enabled = True
                
            End If
        End If
    End If
End Sub
Private Sub txtgrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtgrupo.Text) = "" Then
            cmdBusgrupo_Click
        Else
            If ValidaGrupo = False Then
                 Exit Sub
            Else
                StrSQL = "SELECT  Des_GruTela as Descripción FROM TX_GRUTELA WHERE Cod_FamTela='" & Trim(txtfamilia.Text) & "' AND Cod_GruTela='" & Trim(txtgrupo.Text) & "'"
                txtdes_famgruite = DevuelveCampo(StrSQL, cConnect)
                FunctBuscar.SetFocus
            End If
        End If
    End If
End Sub

Private Sub CARGA_GRID()
    Set rsgrid = New ADODB.Recordset
    rsgrid.ActiveConnection = cConnect
    rsgrid.CursorType = adOpenStatic
    rsgrid.CursorLocation = adUseClient
    rsgrid.LockType = adLockReadOnly

    'Esta cadena es para devolver el Codigo de Cliente
    StrSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtcliente.Text) & "'"

    'Esta cadena es la que nos devolvera los items segun la seleccion establecida
    StrSQL = "EXEC UP_SEL_TALLAS " & Opcion & ",'" & txtfamilia.Text & "','" & Right(txtgrupo.Text, 4) & "','" & txtcod_tela.Text & "','" & DevuelveCampo(StrSQL, cConnect) & "','" & txttemporada.Text & "'"
    rsgrid.Open StrSQL
    Set DGridLista.DataSource = rsgrid

    If rsgrid.RecordCount > 0 Then
        'HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Call CargaDatos
    Else
        'HabilitaMant Me.MantFunc1, "ADICIONAR"
        Call LIMPIAR_DATOS
    End If
End Sub

Public Function CompletaCodigo(CodOrigen As String, longcodfinal As Integer, PosfinalCod As Integer) As String
' CodOrigen     = Es el codigo que sera pasado por parametro
' LongCodFinal  = Es el tamaño del Codigo a devolver
' PosFinalCod   = Es la posicion de la 1era parte del codigo
    Dim Contador As Integer
    CompletaCodigo = Mid(CodOrigen, 1, PosfinalCod)
    For Contador = 1 To longcodfinal - Len(CodOrigen)
        CompletaCodigo = CompletaCodigo & "0"
    Next
    CompletaCodigo = CompletaCodigo & Right(CodOrigen, Len(CodOrigen) - PosfinalCod)
End Function
Public Function ValidaFamilia() As Boolean
    Dim RS As New ADODB.Recordset
    Dim opcmessage As Integer
    RS.ActiveConnection = cConnect
    RS.CursorType = adOpenStatic
    RS.CursorLocation = adUseClient
    RS.LockType = adLockReadOnly
    RS.Open "SELECT Cod_FamTela as Código, Des_FamTela as Descripción FROM TX_FAMTELA WHERE Cod_FamTela='" & Trim(txtfamilia.Text) & "'"
'    If Rs.EOF Then
'        opcmessage = MsgBox("La familia ingresada no existe, Desea Crearla?", vbInformation + vbYesNo)
'        If opcmessage = vbYes Then
'            Load frmMantFamTela
'            frmMantFamTela.Show 1
'
'        Else
'        ValidaFamilia = False
'        End If
'    Else
'        ValidaFamilia = True
 '   End If
    Set RS = Nothing
End Function

Public Function ValidaGrupo() As Boolean
    Dim RS As New ADODB.Recordset
    Dim opcmessage As Integer
    RS.ActiveConnection = cConnect
    RS.CursorType = adOpenStatic
    RS.CursorLocation = adUseClient
    RS.LockType = adLockReadOnly
    RS.Open "SELECT  Cod_GruTela, Des_GruTela FROM TX_GRUTELA WHERE Cod_FamTela='" & Trim(txtfamilia.Text) & "' AND Cod_GruTela='" & Trim(txtgrupo.Text) & "'"
'    If Rs.EOF Then
'        opcmessage = MsgBox("El Grupo ingresado no existe. Desea añadirlo", vbInformation + vbYesNo)
'        If opcmessage = vbYes Then
'            Load frmMantFamGruTela
'            frmMantFamGruTela.txtCod_FamTela = txtfamilia.Text
'            frmMantFamGruTela.CARGAR_DATOS
'            frmMantFamGruTela.Show 1
'        Else
'            ValidaGrupo = False
'        End If
'    Else
'        ValidaGrupo = True
'    End If
    Set RS = Nothing
End Function

Public Sub CargaCombos()
        
   
    'Llena Familia de Telas
    StrSQL = "SELECT des_famtela + space(100) + cod_famtela  FROM TX_FamTela"
    Call LlenaCombo(cboCod_FamTela, StrSQL, cConnect)
    
    'Llena Grupo
    'cboCod_GruTela
    
    'Llena Unida de MEdida
    StrSQL = "SELECT Des_UniMed + space(100) + Cod_UniMed  FROM TG_UniMed"
    Call LlenaCombo(cboCod_UniMed, StrSQL, cConnect)
    
    'Llena Unida de MEdida
    StrSQL = "SELECT Des_UniMed + space(100) + Cod_UniMed  FROM TG_UniMed"
    Call LlenaCombo(cboCod_UniMedcnf, StrSQL, cConnect)
    
    'LLena Galga
    StrSQL = "SELECT Des_Galga + space(100) + Cod_Galga  FROM TX_GALGA"
    Call LlenaCombo(cboCod_Galga, StrSQL, cConnect)
        
    'Llena Tipos de Ancho
    StrSQL = "SELECT Des_TipAncho +  SPACE(100) + Tip_Ancho FROM TG_TIPANC"
    Call LlenaCombo(cboTip_Ancho, StrSQL, cConnect)
    'cboTip_Ancho
    
    'Llema Tipos de Tela
    StrSQL = "SELECT Des_TipTela + space(100) + Cod_TipTela FROM TG_TIPTELA"
    Call LlenaCombo(cboCodTip_Tela, StrSQL, cConnect)
    
    'Llena Tipos de Raya
    StrSQL = "SELECT Des_TipRaya + space(100) + Cod_TipRaya FROM TX_TIPRAYA"
    Call LlenaCombo(cboCod_TipRaya, StrSQL, cConnect)
    
    'Llena Status Opeartividad Telas
    StrSQL = "SELECT Flg_Operatividad + space(1) + Des_Operatividad FROM TX_StatusOperatividad_Tela"
    Call LlenaCombo(cboFlg_Operatividad, StrSQL, cConnect)
    
    
    
'    'Combo Familia Item
'    Strsql = "SELECT des_famitem + space(100) + cod_famitem  FROM LG_FamIte"
'    Call LlenaCombo(cboCod_FamItem, Strsql, cCONNECT)
'
'    'Combo Unidad de Medida
'    Strsql = "SELECT Des_UniMed + space(100) + Cod_UniMed  FROM TG_UniMed"
'    Call LlenaCombo(cboCod_UniMed, Strsql, cCONNECT)
'
'    'Combo Clase de Item
'    Strsql = "SELECT des_claitem + space(100) + cod_claitem  FROM LG_Claitem"
'    Call LlenaCombo(cboCod_ClaItem, Strsql, cCONNECT)
'
'    'Combo Flag Estatus
'    Strsql = "SELECT des_status + space(100) + flg_status  FROM TG_StaDes"
'    'Strsql = "SELECT cod_famitem as Codigo, des_famitem as Descripcion FROM LG_FamIte"
'    Call LlenaCombo(cboFlg_Status, Strsql, cCONNECT)
'
'    'Combo Origen
'    Strsql = "SELECT des_origen + space(100) + cod_origen  FROM LG_Origen"
'    Call LlenaCombo(cboCod_Origen, Strsql, cCONNECT)
'
'    'Combo Motivo Preproduccion
'    Strsql = "SELECT des_motprepro + space(100) + cod_motprepro  FROM TG_MotPrePro"
'    Call LlenaCombo(cboCod_MotPrePro, Strsql, cCONNECT)
'
    
End Sub

Public Sub CargaDatos()

    If Not rsgrid.EOF Then
        If IsNull(rsgrid("Fec_Ult_Revision")) Then
            TxtSinRevision.Visible = True
            DTPUltRevision.Visible = False
        Else
            DTPUltRevision.Value = rsgrid("Fec_Ult_Revision")
            TxtSinRevision.Visible = False
            DTPUltRevision.Visible = True
        End If
        
        TxtAnchoLavado.Text = rsgrid("Ancho_Lavado")
        TxtRendimiento.Text = DevuelveCampo("select dbo.es_calcula_rendimiento_mts_por_kg_tela('" & rsgrid("Cod_Tela") & "')", cConnect)
        txtcodtela.Text = rsgrid("Cod_Tela")
        txtdestela.Text = Trim(rsgrid("Des_Tela"))
        txtGramaje_Acab.Text = Trim(rsgrid("Gramaje_Acab"))
        
        txtAncho_Acab.Text = Trim(rsgrid("Ancho_Acab"))
        txtAncho_Acab_Abierto.Text = Trim(rsgrid("Ancho_Acab")) * 2
        
        txtEncog_Ancho.Text = Trim(rsgrid("Encog_Ancho"))
        txtEncog_Largo.Text = Trim(rsgrid("Encog_Largo"))
        TxtGramDesLavado.Text = Trim(rsgrid("Gramaje_despuesLavado"))
        TxtMts_Twill_x_Hora.Text = Trim(rsgrid("Mts_Twill_x_Hora"))
        txtDes_Tela_Comercial.Text = IIf(IsNull(Trim(rsgrid("Des_Tela_Comercial"))), "", Trim(rsgrid("Des_Tela_Comercial")))
        
        If IIf(IsNull(rsgrid("grado_linea_doblez")), " ", rsgrid("grado_linea_doblez")) = "N" Or IIf(IsNull(rsgrid("grado_linea_doblez")), " ", rsgrid("grado_linea_doblez")) = " " Then
            OptNinguna.Value = True
        ElseIf rsgrid("grado_linea_doblez") = "M" Then
            OptMangas.Value = True
        ElseIf rsgrid("grado_linea_doblez") = "A" Then
            OptAmbas.Value = True
        End If
        
        If IIf(IsNull(rsgrid("inclinacion_trama")), "0", rsgrid("inclinacion_trama")) = "0" Or rsgrid("inclinacion_trama") = " " Then
            Opt0 = True
        ElseIf IIf(IsNull(rsgrid("inclinacion_trama")), "0", rsgrid("inclinacion_trama")) = "1" Then
            Opt1 = True
        End If
        
        
        
        'Esta validacion se efectua por que ya existe data con valores nulos
        If IsNull(rsgrid("Num_Lavadas")) Then
            txtNum_Lavadas.Text = "0"
        Else
            txtNum_Lavadas.Text = rsgrid("Num_Lavadas")
        End If
                       
        If IsNull(rsgrid("Encog_Ancho_Vap")) Then
            txtEncog_Ancho_Vap.Text = "0"
        Else
            txtEncog_Ancho_Vap.Text = rsgrid("Encog_Ancho_Vap")
        End If
        
        If IsNull(rsgrid("Encog_Largo_Vap")) Then
            txtEncog_Largo_Vap.Text = "0"
        Else
            txtEncog_Largo_Vap.Text = rsgrid("Encog_Largo_Vap")
        End If
        
        If IsNull(rsgrid("Rapport")) Then
            txtRapport.Text = ""
        Else
            txtRapport.Text = Trim(rsgrid("Rapport"))
        End If
        
        
        If IsNull(rsgrid("Comentario")) Then
            txtComentario.Text = ""
        Else
            txtComentario.Text = Trim(rsgrid("Comentario"))
        End If

        If IsNull(rsgrid("peso_kg")) Then
            txtPeso.Text = "0.00"
        Else
            txtPeso.Text = Trim(rsgrid("peso_kg"))
        End If
        
        'Fin de validacion y asignacion
        txtGramaje_Crudo.Text = CDbl(rsgrid("Gramaje_Crudo"))
        txtAncho_Crudo.Text = CDbl(rsgrid("Ancho_Crudo"))
        txtCod_CtaCont.Text = Trim(rsgrid("Cod_CtaCont"))
        txtdiametro.Text = Trim(rsgrid("diametro"))
        txtNum_Alimentadores.Text = Trim(rsgrid("Num_Alimentadores"))
        txtNum_Aguja.Text = Trim(rsgrid("Num_Aguja"))
        txtNum_Rpm.Text = Trim(rsgrid("Num_Rpm"))
        TxtRevirado.Text = rsgrid("Revirado")
        
        txtCod_Tel_Origen.Text = rsgrid("Cod_Tela_Desarrollo_Origen")
        StrSQL = "SELECT Des_Tela FROM TX_TELA WHERE Cod_Tela= '" & rsgrid("Cod_Tela_Desarrollo_Origen") & "'"
        txtDes_Tel_Origen.Text = DevuelveCampo(StrSQL, cConnect)
        TxtCod_OrdTra.Text = Trim(rsgrid("cod_ordtra_Tintoreria"))
        TxtCod_OrdTra_Tejeduria = Trim(rsgrid("cod_ordtra_Tejeduria_Tela"))
        
        TxtGram_Comercial.Text = rsgrid("gramaje_comercial")
        TxtAnc_Comercial.Text = rsgrid("ancho_comercial")
        TxtTipoCorte.Text = rsgrid("tipo_familia_tela_corte")
        TxtDes_TipoCorte.Text = DevuelveCampo("select rtrim(isnull(descripcion,'')) from Es_Tipo_Fam_Merma_Corte where tipo_familia_tela_corte = '" & rsgrid("tipo_familia_tela_corte") & "'", cConnect)

        '-- RMP --
        txtFactor_Ajuste_Explosion = rsgrid("Factor_Ajuste_Explosion")
        
        FillCombo
        
        StrSQL = "SELECT Des_Comb FROM TX_TELACOMB WHERE Cod_Tela = '" & rsgrid("Cod_Tela_Desarrollo_Origen") & "' and  Cod_Comb='" & rsgrid("Cod_Comb_Desarrollo_Origen") & "'"
        
        BuscaCombo FixNulos(rsgrid("Cod_Comb_Desarrollo_Origen"), vbString) & " " & FixNulos(DevuelveCampo(StrSQL, cConnect), vbString), 1, cboCombo
  
        Call BuscaCombo(rsgrid("Flg_Operatividad"), 1, cboFlg_Operatividad)
        
        Call BuscaCombo(rsgrid("Cod_FamTela"), 2, cboCod_FamTela)
        Call BuscaCombo(rsgrid("Cod_GruTela"), 2, cboCod_GruTela)
        Call BuscaCombo(rsgrid("Cod_UniMed"), 2, cboCod_UniMed)
        Call BuscaCombo(rsgrid("Cod_UniMedcnf"), 2, cboCod_UniMedcnf)
        Call BuscaCombo(rsgrid("Cod_Galga"), 2, cboCod_Galga)
        Call BuscaCombo(rsgrid("Tip_Ancho"), 2, cboTip_Ancho)
        Call BuscaCombo(rsgrid("Cod_TipTela"), 2, cboCodTip_Tela)
        Call BuscaCombo(rsgrid("Cod_TipRaya"), 2, cboCod_TipRaya)
        
        If Trim(rsgrid("Cod_TelaOriginal")) <> "" Then
            ChkTelaProcesoAdicional.Value = Checked
            ChkTelaProcesoAdicional_Click
            TxtCod_TelaOriginal.Text = Trim(rsgrid("Cod_TelaOriginal"))
            TxtDes_TelaOriginal.Text = DevuelveCampo("select des_tela from tx_tela where cod_tela ='" & TxtCod_TelaOriginal.Text & "'", cConnect)
            'TxtSufijo.Text = ""
        Else
            ChkTelaProcesoAdicional.Value = Unchecked
            TxtCod_TelaOriginal.Text = ""
            TxtDes_TelaOriginal.Text = ""
            TxtSufijo.Text = ""
        End If
        
        'Aqui se mostrara el textbox txtLongMalla si es que solo tiene 1 hilado en la composicion
        StrSQL = "SELECT COUNT(*) FROM TX_HilosTel WHERE Cod_Tela='" & rsgrid("Cod_Tela") & "'"
        If DevuelveCampo(StrSQL, cConnect) <> 1 Then
            txtLongMalla.Visible = False
            LblLongMalla.Visible = False
        Else
            txtLongMalla.Visible = True
            LblLongMalla.Visible = True
            StrSQL = "SELECT Long_Malla FROM TX_HilosTel WHERE Cod_Tela='" & rsgrid("Cod_Tela") & "'"
            txtLongMalla.Text = DevuelveCampo(StrSQL, cConnect)
        End If
        
        If rsgrid("Tipo_Medida") = "P" Then
            Me.OptPulgadas = True
        Else
            Me.OptCentimetros = True
        End If
        
        Me.txtColumnas = rsgrid("Num_Columnas")
        Me.txtCunSas = rsgrid("Num_CunSas")
        
        guarda_variables
    End If
End Sub

Public Sub LIMPIAR_DATOS()
    TxtAnchoLavado.Text = "0"
    txtcodtela.Text = ""
    txtdestela.Text = ""
    txtDes_Tela_Comercial.Text = ""
    txtGramaje_Acab.Text = "0"
    txtAncho_Acab.Text = "0"
    txtAncho_Acab_Abierto.Text = "0"
    txtComentario.Text = ""
    txtEncog_Ancho.Text = "0"
    txtEncog_Largo.Text = "0"
    txtNum_Lavadas.Text = "0"
    txtEncog_Ancho_Vap.Text = "0"
    txtEncog_Largo_Vap.Text = "0"
    txtGramaje_Crudo.Text = "0"
    txtAncho_Crudo.Text = "0"
    txtCod_CtaCont.Text = ""
    txtdiametro.Text = "0"
    txtNum_Alimentadores.Text = "0"
    txtNum_Aguja.Text = "0"
    txtNum_Rpm.Text = "0"
    txtRapport.Text = ""
    txtPeso.Text = "0.00"
    TxtMts_Twill_x_Hora = "0"
    '-- RMP --
    txtFactor_Ajuste_Explosion = 1
    
    TxtGram_Comercial.Text = 0
    TxtAnc_Comercial.Text = 0
    
    
    'cboFlg_Operatividad.ListIndex = 0
    
    TxtGramDesLavado.Text = "0"
    
    'Limpiamos el Grupo
    cboCod_GruTela.Clear  '.ListIndex = -1
    
    cboCod_UniMed.ListIndex = -1
    cboCod_UniMedcnf.ListIndex = -1
    
    If Opcion = 1 And Trim(txtfamilia.Text) <> "" And sTipo = "I" Then
        Call BuscaCombo(txtfamilia.Text, 2, cboCod_FamTela)
    Else
        cboCod_FamTela.ListIndex = -1
    End If
    
    txtCod_Tel_Origen = ""
    txtDes_Tel_Origen = ""
    
    cboCod_FamTela_Click
    'cboCod_FamTela.ListIndex = -1
        
    cboCod_Galga.ListIndex = -1
    cboTip_Ancho.ListIndex = -1
    cboCodTip_Tela.ListIndex = -1
    cboCod_TipRaya.ListIndex = -1
    cboFlg_Operatividad.ListIndex = -1
    cboCombo.ListIndex = -1
    OptNinguna = True
    Opt0 = True
    
    ChkTelaProcesoAdicional.Value = Unchecked
    TxtCod_TelaOriginal.Text = ""
    TxtDes_TelaOriginal.Text = ""
    TxtSufijo.Text = ""
    ChkTelaProcesoAdicional_Click
    TxtCod_OrdTra_Tejeduria.Text = ""
    TxtCod_OrdTra.Text = ""
    TxtTipoCorte.Text = ""
    TxtDes_TipoCorte = ""
    
    Me.txtColumnas.Text = "0"
    Me.txtCunSas = "0"
End Sub

Public Sub HABILITA_DATOS()
Dim vCOD_TIPFAMTELA As String, vFLG_TELAS As String

  vCOD_TIPFAMTELA = DevuelveCampo("SELECT cod_tipfamtela from tx_famtela where cod_famtela='" & Right(cboCod_FamTela, 2) & "'", cConnect)
  vFLG_TELAS = DevuelveCampo("select flg_telas  from tx_tipfam where cod_tipfamtela='" & vCOD_TIPFAMTELA & "'", cConnect)

  If vFLG_TELAS = "N" Then
        txtdestela.Enabled = False
  Else
        txtdestela.Enabled = True
  End If
    
    TxtAnchoLavado.Enabled = True
    txtGramaje_Acab.Enabled = True
    txtAncho_Acab.Enabled = True
    txtEncog_Ancho.Enabled = True
    txtEncog_Largo.Enabled = True
    TxtGram_Comercial.Enabled = True
    TxtAnc_Comercial.Enabled = True
    
    txtNum_Lavadas.Enabled = True
    txtEncog_Ancho_Vap.Enabled = True
    txtEncog_Largo_Vap.Enabled = True
    txtComentario.Enabled = True
    txtGramaje_Crudo.Enabled = True
    txtAncho_Crudo.Enabled = True
    'txtLong_Malla1.Enabled = True
    'txtLong_Malla2.Enabled = True
    txtCod_CtaCont.Enabled = True
    TxtGramDesLavado.Enabled = True
    txtdiametro.Enabled = True
    txtNum_Alimentadores.Enabled = True
    txtNum_Aguja.Enabled = True
    txtNum_Rpm.Enabled = True
    txtRapport.Enabled = True
    txtPeso.Enabled = True
    TxtRevirado.Enabled = True
    txtDes_Tela_Comercial.Enabled = True
        
'    If UCase(Mid(Me.cboCod_FamTela, 1, 2)) = "TW" Then
'        TxtMts_Twill_x_Hora.Enabled = True
'    Else
'        TxtMts_Twill_x_Hora.Enabled = False
'    End If
    
    'FraGradoDoblez.Enabled = True
    'FraInclinacion.Enabled = True
    
    '-- RMP --
    txtFactor_Ajuste_Explosion.Enabled = True
    'cboFlg_Operatividad.Locked = False
    
    If Opcion = 1 Then
        cboCod_FamTela.Enabled = False
    Else
        cboCod_FamTela.Enabled = True
    End If
    
    
    cboCod_GruTela.Enabled = True
    cboCod_UniMed.Enabled = True
    cboCod_UniMedcnf.Enabled = True
    cboCod_Galga.Enabled = True
    cboTip_Ancho.Enabled = True
    cboCodTip_Tela.Enabled = True
    cboCod_TipRaya.Enabled = True
    'cboFlg_Operatividad.Enabled = True
    cboCombo.Enabled = True
    txtCod_Tel_Origen.Enabled = True
    txtDes_Tel_Origen.Enabled = True
    TxtTipoCorte.Enabled = True
    TxtDes_TipoCorte.Enabled = True
        
    If sTipo = "I" Then
        ChkTelaProcesoAdicional.Enabled = True
        TxtCod_TelaOriginal.Enabled = True
        TxtDes_TelaOriginal.Enabled = True
        TxtSufijo.Enabled = True
    End If
    TxtCod_OrdTra_Tejeduria.Enabled = True
    TxtCod_OrdTra.Enabled = True
    
    Me.OptPulgadas.Enabled = True
    Me.OptCentimetros.Enabled = True
    Me.txtColumnas.Enabled = True
    Me.txtCunSas.Enabled = True
    
End Sub

Public Sub INHABILITA_DATOS()
    TxtAnchoLavado.Enabled = False
    txtGramaje_Acab.Enabled = False
    txtAncho_Acab.Enabled = False
    txtcodtela.Enabled = False
    txtdestela.Enabled = False
    txtDes_Tela_Comercial.Enabled = False
    txtGramaje_Acab.Enabled = False
    txtAncho_Acab.Enabled = False
    txtEncog_Ancho.Enabled = False
    txtEncog_Largo.Enabled = False
    txtNum_Lavadas.Enabled = False
    txtEncog_Ancho_Vap.Enabled = False
    txtEncog_Largo_Vap.Enabled = False
    txtGramaje_Crudo.Enabled = False
    txtAncho_Crudo.Enabled = False
    txtComentario.Enabled = False
    'txtLong_Malla1.Enabled = False
    'txtLong_Malla2.Enabled = False
    txtCod_CtaCont.Enabled = False
    txtCod_Tel_Origen.Enabled = False
    txtDes_Tel_Origen.Enabled = False
    'txtCod_TelaOriginal.Enabled = False
    'txtCod_TelaFinal.Enabled = False
    'txtAcu_Porcentaje.Enabled = False
    txtdiametro.Enabled = False
    txtNum_Alimentadores.Enabled = False
    txtNum_Aguja.Enabled = False
    txtNum_Rpm.Enabled = False
    txtRapport.Enabled = False
    txtPeso.Enabled = False
    TxtRevirado.Enabled = False
    txtFactor_Ajuste_Explosion.Enabled = False
    'FraGradoDoblez.Enabled = False
    'FraInclinacion.Enabled = False
    TxtMts_Twill_x_Hora.Enabled = False
    TxtGramDesLavado.Enabled = False
    '-- RMP --
    txtFactor_Ajuste_Explosion.Visible = False
    lblFactor_Ajuste_Explosion.Visible = False
    If sFlg_Factor_Ajuste_Explosion = "S" Then
        txtFactor_Ajuste_Explosion.Visible = True
        lblFactor_Ajuste_Explosion.Visible = True
    End If
    
    'cboFlg_Operatividad.Locked = True
    
    TxtGram_Comercial.Enabled = False
    TxtAnc_Comercial.Enabled = False
    cboCod_FamTela.Enabled = False
    cboCod_GruTela.Enabled = False
    cboCod_UniMed.Enabled = False
    cboCod_UniMedcnf.Enabled = False
    cboCod_Galga.Enabled = False
    cboTip_Ancho.Enabled = False
    cboCodTip_Tela.Enabled = False
    cboCod_TipRaya.Enabled = False
    'cboFlg_Operatividad.Enabled = False
    cboCombo.Enabled = False
    ChkTelaProcesoAdicional.Enabled = False
    TxtCod_TelaOriginal.Enabled = False
    TxtDes_TelaOriginal.Enabled = False
    TxtSufijo.Enabled = False
    TxtCod_OrdTra_Tejeduria.Enabled = False
    TxtCod_OrdTra.Enabled = False
    TxtTipoCorte.Enabled = False
    TxtDes_TipoCorte.Enabled = False
    
    Me.txtColumnas.Enabled = False
    Me.txtCunSas.Enabled = False
    
    Me.OptPulgadas.Enabled = False
    Me.OptCentimetros.Enabled = False
    
    
    
End Sub

Public Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True
    If sTipo <> "D" Then 'Es decir es "I" o "U"
        If Trim(txtdestela.Text) = "" Then
            Call MsgBox("La descripción no puede estar vacia. Sirvase verificar", vbCritical)
            txtdestela.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
        If Trim(cboCod_FamTela.Text) = "" Then
            Call MsgBox("La Familia no puede estar vacia. Sirvase verificar", vbCritical)
            cboCod_FamTela.SetFocus
            VALIDA_DATOS = False
            Exit Function
        End If
        
        StrSQL = "SELECT ISNULL(Flg_Grupo_Obl,'') FROM  TX_FAMTELA WHERE Cod_FamTela = '" & Right(Me.cboCod_FamTela.Text, 2) & "'"
        If DevuelveCampo(StrSQL, cConnect) = "S" And Trim(cboCod_GruTela.Text) = "" Then
            Call MsgBox("El Grupo de Tela no puede estar vacio. Sirvase verificar", vbCritical)
            VALIDA_DATOS = False
            Me.cboCod_GruTela.SetFocus
            Exit Function
        End If
        
        If Trim(cboCod_UniMed.Text) = "" Then
            Call MsgBox("La U.M de Tela no puede estar vacia. Sirvase verificar", vbCritical)
            VALIDA_DATOS = False
            cboCod_UniMed.SetFocus
            Exit Function
        End If
        If Trim(cboCod_UniMedcnf.Text) = "" Then
            Call MsgBox("La U.M. de Confección no puede estar vacia. Sirvase verificar", vbCritical)
            VALIDA_DATOS = False
            cboCod_UniMedcnf.SetFocus
            Exit Function
        End If
'        If Trim(cboCod_Galga.Text) = "" Then
'            Call MsgBox("El Tipo de Galga no puede estar vacio. Sirvase verificar", vbCritical)
'            VALIDA_DATOS = False
'            cboCod_Galga.SetFocus
'            Exit Function
'        End If
        If Trim(cboTip_Ancho.Text) = "" Then
            Call MsgBox("El Tipo de Ancho no puede estar vacio. Sirvase verificar", vbCritical)
            VALIDA_DATOS = False
            cboTip_Ancho.SetFocus
            Exit Function
        End If
        If Trim(cboCodTip_Tela.Text) = "" Then
            Call MsgBox("El Tipo de Tela no puede estar vacio. Sirvase verificar", vbCritical)
            VALIDA_DATOS = False
            cboCodTip_Tela.SetFocus
            Exit Function
        End If
'        If Right(cboCodTip_Tela, 1) = "L" Then
'            If Trim(cboCod_TipRaya.Text) = "" Then
'                Call MsgBox("El Tipo de Raya no puede estar vacio. Sirvase verificar", vbCritical)
'                VALIDA_DATOS = False
'                cboCodTip_Tela.SetFocus
'                Exit Function
'            End If
'        End If
        If Trim(Me.TxtGramDesLavado.Text) = "" Then
            Me.TxtGramDesLavado.Text = 0
        End If
                
        If sTipo = "I" And ChkTelaProcesoAdicional Then
            If Trim(TxtCod_TelaOriginal.Text) = "" Then
                Call MsgBox("El Codigo Tela Original no puede estar vacio. Sirvase verificar", vbCritical)
                VALIDA_DATOS = False
                TxtCod_TelaOriginal.SetFocus
                Exit Function
            End If
            If Trim(TxtSufijo.Text) = "" Then
                Call MsgBox("El Sufijo no puede estar vacio. Sirvase verificar", vbCritical)
                VALIDA_DATOS = False
                TxtSufijo.SetFocus
                Exit Function
            End If
        End If
        
'        If Trim(txtCod_Tel_Origen) <> "" Then
'            strSQL = "select count(*) from ti_ordtra_Tintoreria_ITEMS where cod_ordtra='" & Trim(TxtCod_OrdTra.Text) & "' AND COD_TELA ='" & txtCod_Tel_Origen.Text & "' and cod_comb ='" & Left(cboCombo.Text, 3) & "'"
'            If DevuelveCampo(strSQL, cCONNECT) = 0 Then
'                MsgBox "Desarrollo - Comb Tela no pertenece a la Partida, verifique", vbCritical, "Orden de Trabajo"
'                VALIDA_DATOS = False
'                TxtCod_OrdTra.SetFocus
'                SelectionText TxtCod_OrdTra
'                Exit Function
'            End If
'
'            strSQL = "select count(*) from tx_ordtra_TEJEDURIA where cod_ordtra='" & Trim(TxtCod_OrdTra_Tejeduria.Text) & "' and cod_tela ='" & txtCod_Tel_Origen & "'" 'and cod_comb='" & Left(cboCombo.Text, 3) & "'"
'            If DevuelveCampo(strSQL, cCONNECT) = 0 Then
'                MsgBox "Orden de Trabajo no valida, verifique telas", vbCritical, "Orden de Trabajo"
'                TxtCod_OrdTra_Tejeduria.SetFocus
'                SelectionText TxtCod_OrdTra_Tejeduria
'                VALIDA_DATOS = False
'                Exit Function
'            End If
'        End If
    End If
        
End Function

Public Function SALVAR_DATOS() As Boolean
    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Dim Strconsulta As String
    
    SALVAR_DATOS = False
    Con.ConnectionString = cConnect
    Con.Open
    
    Con.BeginTrans
       
    'Esta sentecia es para obtener el Codigo de Cliente
    Strconsulta = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtcliente.Text) & "'"

    'Esta es la sentencia que realizara el salvado de datos
    
    If OptNinguna = True Then
        doblez = "N"
    ElseIf OptMangas = True Then
        doblez = "M"
    Else
        doblez = "A"
    End If
    
    If Opt0 = True Then
        inclinacion = "0"
    Else
        inclinacion = 1
    End If
    
    If Trim(TxtGram_Comercial.Text) = "" Then TxtGram_Comercial.Text = 0
    If Trim(TxtGram_Comercial.Text) = "" Then TxtGram_Comercial.Text = 0
    If Trim(TxtAnchoLavado.Text) = "" Then TxtAnchoLavado.Text = 0
        
    StrSQL = "EXEC UP_MAN_TELAS " & _
    Opcion & ",'" & _
    sTipo & "','" & _
    txtcodtela.Text & "','" & _
    txtdestela.Text & "','" & _
    Right(cboCodTip_Tela, 1) & "'," & _
    txtGramaje_Acab.Text & "," & _
    txtAncho_Acab.Text & ",'" & _
    Right(cboCod_FamTela, 2) & "'," & _
    "0" & "," & _
    "0" & ",'" & _
    Right(cboCod_UniMed, 2) & "','" & _
    Right(cboCod_UniMedcnf, 2) & "'," & _
    txtEncog_Ancho.Text & "," & _
    txtEncog_Largo.Text & ",'" & _
    Right(cboCod_GruTela, 4) & "','" & _
    Right(cboCod_Galga, 2) & "'," & _
    txtGramaje_Crudo.Text & "," & _
    txtAncho_Crudo.Text & ",'" & _
    Right(cboTip_Ancho, 1) & "','" & _
    txtCod_CtaCont.Text & "','" & _
    TxtCod_TelaOriginal.Text & "','" & _
    "" & "',"
    
    'txtLong_Malla1.Text & "," & _
    'txtLong_Malla2.Text & ",'" & _

    StrSQL = StrSQL & _
    txtdiametro.Text & ",'" & _
    Right(cboCod_TipRaya, 2) & "'," & _
    txtNum_Alimentadores.Text & "," & _
    txtNum_Aguja.Text & "," & _
    txtNum_Rpm.Text & ",'" & _
    DevuelveCampo(Strconsulta, cConnect) & "','" & _
    txttemporada.Text & "'," & _
    txtNum_Lavadas.Text & "," & _
    txtEncog_Ancho_Vap.Text & "," & _
    txtEncog_Largo_Vap.Text & ",'" & _
    txtRapport.Text & "','" & _
    txtComentario.Text & "'," & txtPeso.Text & "," & _
    IIf(Trim(TxtRevirado) = "", 0, TxtRevirado.Text) & ",'" & vusu & "', '" & _
    Left(cboFlg_Operatividad, 1) & "', " & txtFactor_Ajuste_Explosion & ",'" & _
    doblez & "','" & inclinacion & "','" & ComputerName & "','" & vusu & "'," & _
    TxtGramDesLavado & "," & CDbl(TxtMts_Twill_x_Hora.Text) & ",'" & _
    TxtSufijo.Text & "','" & txtCod_Tel_Origen & "','" & Left(cboCombo.Text, 3) & "','" & _
    TxtCod_OrdTra & "','" & TxtCod_OrdTra_Tejeduria & "'," & CDbl(TxtGram_Comercial.Text) & "," & CDbl(TxtAnc_Comercial.Text) & ",'" & _
    Trim(TxtTipoCorte.Text) & "'," & CDbl(TxtAnchoLavado.Text) & ",'" & _
    IIf(OptCentimetros, "C", "P") & "'," & _
    Me.txtColumnas & "," & _
    Me.txtCunSas & " ,'" & txtDes_Tela_Comercial.Text & "' "

    '-- RMP --: agregué cboflg_operatividad y converti el proc. en funcion
    Con.Execute StrSQL
        
    Con.CommitTrans
    Dim amensaje As New clsMessages
    amensaje.CODIGO = CodeMsg.kMeSsaGe_INF_DATA_save
    Informa "", amensaje
    Call INHABILITA_DATOS
    'Call LIMPIAR_DATOS
    SALVAR_DATOS = True
    Exit Function
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler err, "Salvar_Datos"
End Function

Public Sub ELIMINAR_DATOS()
    Dim Con As New ADODB.Connection
    Dim Strconsulta As String
    On Error GoTo Eliminar_DatosErr
       
    Strconsulta = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtcliente.Text) & "'"

    'Esta consulta verifica si este registro posee registros relacionados
    StrSQL = "SELECT COD_CLIENTE FROM TX_TELATEMCLI WHERE Cod_Tela='" & txtcodtela.Text & "'"
    If DevuelveCampo(StrSQL, cConnect) <> "" Then
        MsgBox ("No se puede eliminar el Registro por que posee registros relacionados")
        Exit Sub
    End If
    
    Con.ConnectionString = cConnect
    Con.Open
    Con.BeginTrans
           
        StrSQL = "EXEC UP_MAN_TELAS " & _
        Opcion & ",'" & _
        sTipo & "','" & _
        txtcodtela.Text & "','" & _
        txtdestela.Text & "','" & _
        Right(cboCodTip_Tela, 1) & "'," & _
        txtGramaje_Acab.Text & "," & _
        txtAncho_Acab.Text & ",'" & _
        Right(cboCod_FamTela, 2) & "'," & _
        "0" & "," & _
        "0" & ",'" & _
        Right(cboCod_UniMed, 2) & "','" & _
        Right(cboCod_UniMedcnf, 2) & "'," & _
        txtEncog_Ancho.Text & "," & _
        txtEncog_Largo.Text & ",'" & _
        Right(cboCod_GruTela, 4) & "','" & _
        Right(cboCod_Galga, 2) & "'," & _
        txtGramaje_Crudo.Text & "," & _
        txtAncho_Crudo.Text & ",'" & _
        Right(cboTip_Ancho, 1) & "','" & _
        txtCod_CtaCont.Text & "','" & _
        "" & "','" & _
        "" & "'," & 0
        
        'txtLong_Malla1.Text & "," & _
        'txtLong_Malla2.Text & ",'" & _

        StrSQL = StrSQL & _
        txtdiametro.Text & ",'" & _
        Right(cboCod_TipRaya, 2) & "'," & _
        txtNum_Alimentadores.Text & "," & _
        txtNum_Aguja.Text & "," & _
        txtNum_Rpm.Text & ",'" & _
        DevuelveCampo(Strconsulta, cConnect) & "','" & _
        txttemporada.Text & "'," & _
        txtNum_Lavadas.Text & "," & _
        txtEncog_Ancho_Vap.Text & "," & _
        txtEncog_Largo_Vap.Text & ",'" & _
        txtRapport.Text & "','',0, 0, '" & vusu & "', '" & _
        Left(cboFlg_Operatividad, 1) & "',0,'','','','',0,0"
        '-- RMP --
        
        Con.Execute StrSQL
    
    Con.CommitTrans
    
    Dim amensaje As New clsMessages
    amensaje.CODIGO = CodeMsg.kMeSsaGe_INF_DATA_DELETE
    Informa "", amensaje

    LIMPIAR_DATOS
    'RECARGAR_DATOS
Exit Sub
Eliminar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler err, "Eliminar_Datos"
End Sub

Sub guarda_variables()
Dim Strconsulta As String
Strconsulta = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(txtcliente.Text) & "'"

If OptNinguna = True Then
    doblez = "N"
ElseIf OptMangas = True Then
    doblez = "M"
Else
    doblez = "A"
End If
        
If Opt0 = True Then
   inclinacion = "0"
Else
   inclinacion = 1
End If

With FrmCambiosTelas
   .Cod_tiptela = Right(cboCodTip_Tela, 1)
   .Cod_Famtela = Right(cboCod_FamTela, 2)
   .gramaje1_lav = 0
   .Ancho1_Lav = 0
   .cod_uniMed = Right(cboCod_UniMed, 2)
   .cod_uniMedCnf = Right(cboCod_UniMedcnf, 2)
   .Encog_Ancho = txtEncog_Ancho
   .Encog_Largo = txtEncog_Ancho
   .cod_gruTela = Right(cboCod_GruTela, 4)
   .cod_galga = Right(cboCod_Galga, 2)
   .gramaje_crudo = txtGramaje_Crudo
   .Ancho_crudo = txtAncho_Crudo
   .Tip_Ancho = Right(cboTip_Ancho, 1)
   .Cod_CtaCont = txtCod_CtaCont
   .cod_telaoriginal = ""
   .cod_telaFinal = ""
   .diametro = txtdiametro.Text
   .cod_tipRaya = Right(cboCod_TipRaya, 2)
   .num_alimentadores = txtNum_Alimentadores
   .num_aguja = txtNum_Aguja
   .num_rpm = txtNum_Rpm
   .cod_cliente = DevuelveCampo(Strconsulta, cConnect)
   .cod_temcli = txttemporada.Text
   .num_lavadas = txtNum_Lavadas
   .Encog_Ancho_Vap = txtEncog_Ancho
   .Encog_Largo_Vap = txtEncog_Largo
   .Rapport = txtRapport
   .comentario = txtComentario
   .peso = txtPeso
   .Grado_Doblez = doblez
   .inclinacion = inclinacion
   .flg_operatividad = Left(cboFlg_Operatividad, 1)
   .Opcion = Opcion
   .gramaje_despueslavado = Me.TxtGramDesLavado
   .sMts_Twill_x_Hora = CDbl(TxtMts_Twill_x_Hora)
   .Cod_Tela_Desarrollo_Origen = txtCod_Tel_Origen
   .Cod_Comb_Desarrollo_Origen = Left(cboCombo.Text, 3)
End With
End Sub

Sub NO_OPERATIVAS()
On Error GoTo hand
    Dim oo As Object
    Dim StrSQL As String
    Screen.MousePointer = 11
    
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\TelasNoOperativas.xlt"
    oo.Visible = True
    oo.run "Reporte", cConnect
    Screen.MousePointer = vbNormal
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler err, "Telas No Operativas"
    Screen.MousePointer = vbNormal
    Set oo = Nothing
End Sub

Sub Hoja_Ruta()
Dim sRuta As String
On Error GoTo hand
    Dim oo As Object
    Dim StrSQL As String
    Screen.MousePointer = 11
    
    sRuta = vRuta
    
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open sRuta & "\Hoja_Tecnica.xlt"
    oo.Visible = True
    oo.run "Reporte", txtcodtela.Text, cConnect, sRuta, "Ocultar"
    Screen.MousePointer = vbNormal
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler err, "Telas No Operativas"
    Screen.MousePointer = vbNormal
    Set oo = Nothing
End Sub

Sub CAMBIA_STATUS()
On Error GoTo ErrStatus

'If Val(DevuelveCampo("SELECT DBO.lg_devuelve_seg_Famite('" & vusu & "','" & rsgrid("Cod_FamTela") & "')", cCONNECT)) = 0 Then
'    MsgBox "Usuario no puede poner No Operativo en esta familia", vbCritical, "No Operativo"
'    Exit Sub
'End If

StrSQL = "TX_Cambio_Status_Tela '" & rsgrid("Cod_Tela") & "','" & vusu & "','" & ComputerName & "'"
Call ExecuteCommandSQL(cConnect, StrSQL)

Call CARGA_GRID

Exit Sub
ErrStatus:
    ErrorHandler err, "Cambia Status"
End Sub

Sub Modifica_Comercial()
Dim i As Integer
On Error GoTo errModifica

i = DGridLista.Row

StrSQL = "TX_ACTUALIZA_DATOS_GRAMAJE_ANCHO_COMERCIAL '" & rsgrid("Cod_Tela") & "'," & CDbl(TxtGramaje_Comercial.Text) & "," & CDbl(TxtAncho_Comercial.Text) & ",'" & vusu & "','" & ComputerName & "'"
ExecuteCommandSQL cConnect, StrSQL

FraComercial.Visible = False
FraLista.Enabled = True

Call CARGA_GRID
DGridLista.Row = i

Exit Sub
errModifica:
    MsgBox err.Description, vbCritical, "Modificacion Comercial"
End Sub

Public Sub Busca_FamiliaTipoCorte(Opcion As Integer)
Dim rstAux As ADODB.Recordset
On Error GoTo fin
Dim iCol As Long

    StrSQL = "SELECT tipo_familia_tela_corte as Codigo, Descripcion as Descripcion, Merma_1, Merma_2 FROM Es_Tipo_Fam_Merma_Corte where "
    
    Select Case Opcion
    Case 1: StrSQL = StrSQL & " tipo_familia_tela_corte like '%" & Trim(TxtTipoCorte) & "%'"
    Case 2: StrSQL = StrSQL & " Descripcion like '%" & Trim(TxtDes_TipoCorte.Text) & "%'"
    End Select
    
    With frmBusqGeneral3
        Set .oParent = Me
        .SQuery = StrSQL
        .Cargar_Datos
        .Caption = "Seleccionar Familia Tipo Corte"
        CODIGO = ".."
        'Set rstAux = .DGridLista.ADORecordset
        Set rstAux = .gexLista.ADORecordset
        .gexLista.Columns("Codigo").Width = 700
        .gexLista.Columns("Descripcion").Width = 3000
        .gexLista.Columns("Merma_1").Width = 700
        .gexLista.Columns("Merma_2").Width = 700
        
        If rstAux.RecordCount = 1 Then
            CODIGO = Trim(rstAux!CODIGO)
        End If
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If CODIGO <> "" And rstAux.RecordCount > 0 Then
            TxtTipoCorte = CODIGO
            TxtDes_TipoCorte = Descripcion
            MantFunc1.SetFocus
        End If
    End With
    CODIGO = "": Descripcion = ""
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    Set rstAux = Nothing
Exit Sub
fin:
On Error Resume Next
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Busca Familia Tipo Corte (" & Opcion & ")"
End Sub

Private Sub TxtTipoCorte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_FamiliaTipoCorte(1)
End If
End Sub
