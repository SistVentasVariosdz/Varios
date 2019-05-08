VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmAddSolicitudDesaColores 
   Caption         =   "Mantenimiento Solicitud Desarrollo Carta"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraDatos 
      Height          =   3255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7215
      Begin VB.TextBox TxtCorr_Carta 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TxtCod_Cliente 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox TxtDes_Cliente 
         Height          =   285
         Left            =   2870
         TabIndex        =   12
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox TxtNum_Carta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   600
         Width           =   5295
      End
      Begin VB.CommandButton CmdCliente 
         Caption         =   "..."
         Height          =   270
         Left            =   2450
         TabIndex        =   8
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton CmdTipoLuz 
         Caption         =   "..."
         Height          =   270
         Left            =   2445
         TabIndex        =   7
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox TxtDes_TipoLuz 
         Height          =   285
         Left            =   2865
         TabIndex        =   6
         Top             =   2400
         Width           =   3495
      End
      Begin VB.TextBox Txtcod_TipoLuz 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "Nuevo"
         Height          =   255
         Left            =   6390
         TabIndex        =   4
         Top             =   2420
         Width           =   735
      End
      Begin VB.OptionButton optHilo 
         Caption         =   "Hilo"
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   2870
         Width           =   855
      End
      Begin VB.OptionButton optTela 
         Caption         =   "Tela"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   2870
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPSolicitud 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   1680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   71565313
         CurrentDate     =   38210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Corr. Carta"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   315
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   1035
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Solicitud"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Num. Carta"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   2160
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   675
         Width           =   840
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Luz"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   2475
         Width           =   615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Carta"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   2870
         Width           =   735
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2160
      TabIndex        =   0
      Top             =   3360
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmAddSolicitudDesaColores.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "FrmAddSolicitudDesaColores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

