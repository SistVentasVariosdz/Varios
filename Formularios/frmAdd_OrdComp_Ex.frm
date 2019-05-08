VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdd_OrdComp_Ex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adicionar O/C"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   12330
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmCantidadxTela 
      Caption         =   "Ingrese la Cantidad Total Por Tela"
      Height          =   1335
      Left            =   5160
      TabIndex        =   73
      Top             =   2400
      Width           =   3015
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   1680
         TabIndex        =   76
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton CmdAnadir2 
         Caption         =   "Añadir"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   720
         Width           =   1365
      End
      Begin VB.TextBox Txt_CantidadXTela 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   600
         TabIndex        =   74
         Top             =   240
         Width           =   1770
      End
   End
   Begin VB.TextBox Txt_Unidades 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10080
      TabIndex        =   34
      Text            =   "0"
      Top             =   3040
      Width           =   810
   End
   Begin VB.Frame Frame1 
      Height          =   3990
      Left            =   60
      TabIndex        =   50
      Top             =   2520
      Width           =   12225
      Begin VB.CommandButton CmdAnadir 
         Height          =   495
         Left            =   120
         Picture         =   "frmAdd_OrdComp_Ex.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   1590
         Width           =   1365
      End
      Begin VB.CommandButton CmdEliminar 
         Height          =   495
         Left            =   1560
         Picture         =   "frmAdd_OrdComp_Ex.frx":05BE
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   1560
         Width           =   1245
      End
      Begin VB.CommandButton cmdManTela 
         Caption         =   "..."
         Height          =   240
         Left            =   4365
         TabIndex        =   67
         ToolTipText     =   "Agregar / Modificar Telas"
         Top             =   210
         Width           =   285
      End
      Begin VB.TextBox txtDes_Receta 
         Height          =   285
         Left            =   1785
         TabIndex        =   25
         Top             =   1215
         Width           =   2535
      End
      Begin VB.TextBox txtCod_Receta 
         Height          =   285
         Left            =   1050
         MaxLength       =   2
         TabIndex        =   24
         Top             =   1215
         Width           =   735
      End
      Begin VB.TextBox txtCod_TelaCliente 
         Height          =   285
         Left            =   6435
         TabIndex        =   32
         Top             =   1500
         Width           =   1290
      End
      Begin VB.TextBox txtPrecio 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10020
         TabIndex        =   35
         Text            =   "0"
         Top             =   870
         Width           =   825
      End
      Begin VB.TextBox txtCant_Pedida 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10020
         TabIndex        =   33
         Text            =   "0"
         Top             =   195
         Width           =   810
      End
      Begin VB.TextBox txtObs_Det 
         Height          =   705
         Left            =   8715
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         Top             =   1425
         Width           =   2955
      End
      Begin VB.TextBox txt_IGVDet 
         Height          =   285
         Left            =   5925
         TabIndex        =   29
         Top             =   495
         Width           =   510
      End
      Begin VB.TextBox txtDes_DstoDet 
         Height          =   285
         Left            =   6480
         TabIndex        =   28
         Top             =   210
         Width           =   2145
      End
      Begin VB.TextBox txtCod_DsctoDet 
         Height          =   285
         Left            =   5925
         MaxLength       =   2
         TabIndex        =   27
         Top             =   210
         Width           =   525
      End
      Begin VB.TextBox txtCod_Talla 
         Height          =   285
         Left            =   3615
         TabIndex        =   26
         Top             =   1560
         Width           =   735
      End
      Begin GridEX20.GridEX gexDetalle 
         Height          =   1680
         Left            =   120
         TabIndex        =   54
         Top             =   2175
         Width           =   11520
         _ExtentX        =   20320
         _ExtentY        =   2963
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         FormatStylesCount=   7
         FormatStyle(1)  =   "frmAdd_OrdComp_Ex.frx":0B8D
         FormatStyle(2)  =   "frmAdd_OrdComp_Ex.frx":0CC5
         FormatStyle(3)  =   "frmAdd_OrdComp_Ex.frx":0D75
         FormatStyle(4)  =   "frmAdd_OrdComp_Ex.frx":0E29
         FormatStyle(5)  =   "frmAdd_OrdComp_Ex.frx":0F01
         FormatStyle(6)  =   "frmAdd_OrdComp_Ex.frx":0FB9
         FormatStyle(7)  =   "frmAdd_OrdComp_Ex.frx":1099
         ImageCount      =   0
         PrinterProperties=   "frmAdd_OrdComp_Ex.frx":10B9
      End
      Begin VB.TextBox txtCod_Color 
         Height          =   285
         Left            =   1050
         MaxLength       =   6
         TabIndex        =   22
         Top             =   885
         Width           =   735
      End
      Begin VB.TextBox txtDes_Color 
         Height          =   285
         Left            =   1785
         TabIndex        =   23
         Top             =   885
         Width           =   2535
      End
      Begin VB.TextBox txtCod_Tela 
         Height          =   285
         Left            =   1050
         MaxLength       =   8
         TabIndex        =   18
         Top             =   210
         Width           =   1050
      End
      Begin VB.TextBox txtDes_Tela 
         Height          =   285
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   210
         Width           =   2220
      End
      Begin VB.TextBox txtCod_Comb 
         Height          =   285
         Left            =   1050
         MaxLength       =   3
         TabIndex        =   20
         Top             =   555
         Width           =   735
      End
      Begin VB.TextBox txtDes_Comb 
         Height          =   285
         Left            =   1785
         TabIndex        =   21
         Top             =   555
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtpEntregaI_Det 
         Height          =   315
         Left            =   6435
         TabIndex        =   30
         Top             =   810
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   71368707
         CurrentDate     =   37832
      End
      Begin MSComCtl2.DTPicker dtpEntregaF_Det 
         Height          =   315
         Left            =   6435
         TabIndex        =   31
         Top             =   1155
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   71368707
         CurrentDate     =   37832
      End
      Begin VB.Label Label25 
         Caption         =   "Unidades"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   8685
         TabIndex        =   68
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label Label24 
         Caption         =   "Receta"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   180
         TabIndex        =   66
         Top             =   1290
         Width           =   630
      End
      Begin VB.Label Label22 
         Caption         =   "Cod.Tela Cliente"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   4815
         TabIndex        =   64
         Top             =   1530
         Width           =   1455
      End
      Begin VB.Label Label21 
         Caption         =   "Precio ($)"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   8685
         TabIndex        =   63
         Top             =   915
         Width           =   915
      End
      Begin VB.Label Label20 
         Caption         =   "Cant.Pedida (KG)"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   8685
         TabIndex        =   62
         Top             =   255
         Width           =   1275
      End
      Begin VB.Label Label19 
         BackColor       =   &H80000013&
         Caption         =   "Observ:"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   8685
         TabIndex        =   61
         Top             =   1230
         Width           =   885
      End
      Begin VB.Label Label18 
         Caption         =   "%"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   6510
         TabIndex        =   60
         Top             =   570
         Width           =   270
      End
      Begin VB.Label Label17 
         BackColor       =   &H80000013&
         Caption         =   "Fec. Entrega Fin"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   4815
         TabIndex        =   59
         Top             =   1200
         Width           =   1530
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000013&
         Caption         =   "Fec. Entrega Inicio"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   4815
         TabIndex        =   58
         Top             =   870
         Width           =   1530
      End
      Begin VB.Label Label15 
         Caption         =   "I.G.V."
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   4845
         TabIndex        =   57
         Top             =   570
         Width           =   495
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000013&
         Caption         =   "Descuento"
         ForeColor       =   &H80000002&
         Height          =   210
         Left            =   4830
         TabIndex        =   56
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label13 
         Caption         =   "Talla"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2910
         TabIndex        =   55
         Top             =   1605
         Width           =   540
      End
      Begin VB.Label Label12 
         Caption         =   "Color"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   165
         TabIndex        =   53
         Top             =   975
         Width           =   690
      End
      Begin VB.Label Label11 
         Caption         =   "Tela"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   150
         TabIndex        =   52
         Top             =   285
         Width           =   690
      End
      Begin VB.Label Label31 
         Caption         =   "Combo"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   165
         TabIndex        =   51
         Top             =   645
         Width           =   900
      End
   End
   Begin VB.Frame fraOrdComp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2460
      Left            =   60
      TabIndex        =   37
      Top             =   30
      Width           =   12240
      Begin VB.ComboBox cmbViaTransporte_Cli 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   1920
         Width           =   3555
      End
      Begin VB.TextBox TxtPO 
         Height          =   285
         Left            =   1080
         MaxLength       =   60
         TabIndex        =   4
         Top             =   900
         Width           =   2715
      End
      Begin VB.ComboBox cmbPais 
         Height          =   315
         Left            =   6600
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1920
         Width           =   4035
      End
      Begin VB.CheckBox ChkCtrPeso 
         BackColor       =   &H80000018&
         Caption         =   "Control de Peso x Rollo"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   3960
         TabIndex        =   13
         Top             =   1440
         Width           =   2025
      End
      Begin VB.TextBox txtCod_CondVent 
         Height          =   285
         Left            =   1095
         TabIndex        =   5
         Top             =   1215
         Width           =   510
      End
      Begin VB.TextBox txtNom_Cliente 
         Height          =   285
         Left            =   1635
         TabIndex        =   1
         Top             =   285
         Width           =   2145
      End
      Begin VB.TextBox txtAbr_Cliente 
         Height          =   285
         Left            =   1095
         TabIndex        =   0
         Top             =   285
         Width           =   510
      End
      Begin VB.TextBox txtCod_LugEntr 
         Height          =   285
         Left            =   5040
         TabIndex        =   10
         Top             =   630
         Width           =   495
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   2040
      End
      Begin VB.TextBox txtCod_OrdComp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1650
         TabIndex        =   3
         Top             =   615
         Width           =   2145
      End
      Begin VB.TextBox txtSer_OrdComp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1095
         TabIndex        =   2
         Top             =   615
         Width           =   510
      End
      Begin VB.TextBox txtDes_CondVent 
         Height          =   285
         Left            =   1635
         TabIndex        =   6
         Top             =   1215
         Width           =   2145
      End
      Begin VB.TextBox txtCod_Descuento 
         Height          =   285
         Left            =   1095
         TabIndex        =   7
         Top             =   1545
         Width           =   510
      End
      Begin VB.TextBox txtDes_Descuento 
         Height          =   285
         Left            =   1635
         TabIndex        =   8
         Top             =   1545
         Width           =   2145
      End
      Begin VB.TextBox txtPorc_IGV 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7620
         TabIndex        =   38
         Top             =   240
         Width           =   435
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   495
         Left            =   9240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   1155
         Width           =   2880
      End
      Begin VB.TextBox txtDes_LugEntr 
         Height          =   285
         Left            =   5550
         TabIndex        =   11
         Top             =   630
         Width           =   3510
      End
      Begin VB.ComboBox cboTipoOC 
         Height          =   315
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   975
         Width           =   4035
      End
      Begin MSComCtl2.DTPicker dtpEntregaI 
         Height          =   315
         Left            =   10785
         TabIndex        =   15
         Top             =   210
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   71368707
         CurrentDate     =   37832
      End
      Begin MSComCtl2.DTPicker dtpEntregaF 
         Height          =   315
         Left            =   10785
         TabIndex        =   16
         Top             =   540
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   71368707
         CurrentDate     =   37832
      End
      Begin VB.Label Label26 
         BackColor       =   &H80000018&
         Caption         =   "Via Transporte"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label28 
         BackColor       =   &H80000018&
         Caption         =   "P.O."
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   75
         TabIndex        =   72
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label27 
         BackColor       =   &H80000018&
         Caption         =   "País Destino"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   5400
         TabIndex        =   71
         Top             =   2040
         Width           =   1140
      End
      Begin VB.Label Label23 
         BackColor       =   &H80000018&
         Caption         =   "%"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   8085
         TabIndex        =   65
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000018&
         Caption         =   "C. de Pago"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   75
         TabIndex        =   49
         Top             =   1260
         Width           =   900
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000018&
         Caption         =   "Fec. Entrega Inicio"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   9225
         TabIndex        =   48
         Top             =   255
         Width           =   1530
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000018&
         Caption         =   "Cliente"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   75
         TabIndex        =   47
         Top             =   345
         Width           =   630
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000018&
         Caption         =   "Lug. Entrega"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   3945
         TabIndex        =   46
         Top             =   675
         Width           =   1020
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000018&
         Caption         =   "Moneda"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   3945
         TabIndex        =   45
         Top             =   315
         Width           =   720
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000018&
         Caption         =   "Nro. O/C"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   75
         TabIndex        =   44
         Top             =   690
         Width           =   945
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000018&
         Caption         =   "Fec. Entrega Fin"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   9225
         TabIndex        =   43
         Top             =   585
         Width           =   1530
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000018&
         Caption         =   "Descuento"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   75
         TabIndex        =   42
         Top             =   1590
         Width           =   900
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000018&
         Caption         =   "I.G.V."
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   7155
         TabIndex        =   41
         Top             =   270
         Width           =   450
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000018&
         Caption         =   "Observ:"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   9225
         TabIndex        =   40
         Top             =   945
         Width           =   885
      End
      Begin VB.Label Label30 
         BackColor       =   &H80000018&
         Caption         =   "Tipo"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   3930
         TabIndex        =   39
         Top             =   1005
         Width           =   960
      End
   End
   Begin VB.TextBox txt_totalcarga 
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   77
      Text            =   "Text1"
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   4320
      TabIndex        =   80
      Top             =   6600
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmAdd_OrdComp_Ex.frx":1291
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Label lbl_Total_Detalle 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10200
      TabIndex        =   70
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label lbl_et_Total_Detalle 
      Caption         =   "Kilos restantes para la tela"
      Height          =   495
      Left            =   8160
      TabIndex        =   69
      Top             =   6600
      Width           =   1815
   End
End
Attribute VB_Name = "frmAdd_OrdComp_Ex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rstAux1 As ADODB.Recordset
Public rstAux2 As ADODB.Recordset
Public rstAux5 As ADODB.Recordset
Public CODIGO As String
Public Descripcion As String
Public TipoAdd As String
Dim strSQL As String
Dim Reg As ADODB.Recordset
Dim Flg_Existe_Tela As Boolean
Dim RegProceso As ADODB.Recordset
Dim sCod_ClaOrdComp As String
Dim sNum_Intentos As Integer
Dim sAccesoPrecio As Boolean
Dim bFilling As Boolean
Public sCadena_Fam As String
Dim cod_PaisX As String
Dim scod_Cliente_x As String
Dim scod_Tela_x  As String

Private Sub LimpiaCab()
    fraOrdComp.Enabled = True
    txtCod_CondVent.Text = ""
    txtDes_CondVent.Text = ""
    txtCod_Descuento.Text = ""
    txtDes_Descuento.Text = ""
    txtCod_LugEntr.Text = ""
    txtDes_LugEntr.Text = ""
    dtpEntregaI.Value = Date
    dtpEntregaF.Value = Date
    cboMoneda.ListIndex = -1
    BuscaCombo1 "C", 100, cboTipoOC
    txtObservaciones.Text = ""
'    Flg_ClientePropio = False
'    Flg_OrdenPropia = False
'    sFechaHora = ""
'    FillCarac

End Sub

Private Sub LimpiaDetalle()
    'txtCod_Tela.Text = ""
    'txtDes_Tela.Text = ""
    txtCod_Comb.Text = ""
    txtDes_Comb.Text = ""
    txtCod_Color.Text = ""
    txtDes_Color.Text = ""
    txtCod_Receta.Text = ""
    txtDes_Receta.Text = ""
    txtCod_Talla.Text = ""
    txtCod_DsctoDet.Text = Trim(txtCod_Descuento.Text)
    txtDes_DstoDet.Text = Trim(txtDes_Descuento.Text)
    txt_IGVDet.Text = txtPorc_IGV.Text
    dtpEntregaI_Det.Value = dtpEntregaI.Value
    dtpEntregaF_Det.Value = dtpEntregaF.Value
    txtCant_Pedida.Text = 0
    txtPrecio.Text = 0
    txtObs_Det.Text = ""
    txtCod_TelaCliente.Text = ""
    Txt_Unidades.Text = 0
End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cboTipoOC_Click()
Dim irow As Long
 If bFilling Then Exit Sub
 sCadena_Fam = ""
 If rstAux1.RecordCount > 0 Then rstAux1.MoveFirst
 With rstAux1
    .Find "Cod_Tipo_Orden_tinto = '" & Left(cboTipoOC, 2) & "'"
    If Not .EOF Then
    If RTrim(rstAux1!Cod_Proceso_Tinto) = "" Then
        sCadena_Fam = ""
        Frm_Procesos_Tinto_Ex.Show vbModal
        For irow = 0 To Frm_Procesos_Tinto_Ex.lstFam.ListCount - 1
            If Frm_Procesos_Tinto_Ex.lstFam.Selected(irow) Then _
            sCadena_Fam = sCadena_Fam & "." & Left(Frm_Procesos_Tinto_Ex.lstFam.List(irow), 2) & ".,"
        Next irow
        If sCadena_Fam <> "" Then sCadena_Fam = Left(sCadena_Fam, Len(sCadena_Fam) - 1)
        Unload Frm_Procesos_Tinto_Ex
        
    End If
    If sCadena_Fam = "" Then sCadena_Fam = "." & RTrim(rstAux1!Cod_Proceso_Tinto) & ".,"
    End If
    End With
 
 
 
    

End Sub

Private Sub cboTipoOC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Function BuscarSaldos() As Boolean
Dim i As Integer
Dim sRows As Integer, dPedido As Double, dTotalxTela As Double, dSumaCantidadTela As Double
On Error GoTo hand

dPedido = 0
dTotalxTela = 0

BuscarSaldos = False

If Not Reg.EOF Then
    sRows = Reg.RecordCount
    Reg.MoveFirst
    
    
        With Reg
            For i = 1 To sRows
            
                If txtCod_Tela = Trim(!Cod_Tela) Then
                
                    dPedido = dPedido + Trim(!Pedida)
                    dTotalxTela = dTotalxTela + Trim(!TotalxTela)
                
                    
                End If
                
                Reg.MoveNext
               
            Next
        End With
        
Reg.MoveFirst


End If

dSumaCantidadTela = dPedido + CDbl(IIf(txtCant_Pedida.Text = "", 0, txtCant_Pedida.Text))

If dTotalxTela > 0 Then

    If dTotalxTela < dSumaCantidadTela Then
        BuscarSaldos = True
    End If
    
Else
    BuscarSaldos = False
End If


    
Exit Function
hand:
    ErrorHandler Err, "SALVAR_CABECERA"
    Set gexDetalle.ADORecordset = Reg

End Function
Private Function BuscaDetalleTelaRepetida() As Boolean
Dim i As Integer
Dim sRows As Integer
On Error GoTo hand

BuscaDetalleTelaRepetida = False

If Not Reg.EOF Then
    sRows = Reg.RecordCount
    Reg.MoveFirst
    

        With Reg
            For i = 1 To sRows
            
                If txtCod_Tela = Trim(!Cod_Tela) Then
                    BuscaDetalleTelaRepetida = True
                End If
                Reg.MoveNext
                'Reg.Delete
                'Reg.MoveFirst
            Next
        End With
End If
    
Exit Function
hand:
    ErrorHandler Err, "SALVAR_CABECERA"
    Set gexDetalle.ADORecordset = Reg

End Function

Private Sub CmdAnadir_Click()
    If Trim(txtCod_Tela.Text) = "" Then
        MsgBox "Seleccione la Tela", vbCritical, Me.Caption
        txtCod_Tela.SetFocus
        Exit Sub
    End If

    If Trim(txtCod_Color.Text) = "" Then
        MsgBox "Seleccione el Color", vbCritical, Me.Caption
        txtCod_Color.SetFocus
        Exit Sub
    End If

    If Trim(txtCod_Descuento.Text) = "" Then
        MsgBox "Seleccione el Descuento", vbCritical, Me.Caption
        txtCod_Descuento.SetFocus
        Exit Sub
    End If

    If Trim(txtCant_Pedida.Text) = "" Or CDbl(txtCant_Pedida.Text) <= 0 Then
        MsgBox "Ingrese una Cantidad valida", vbCritical, Me.Caption
        txtCant_Pedida.SetFocus
        Exit Sub
    End If
    
'    If ValidaCargaItem = True Then
'        MsgBox "La cantidad a Ingresar sobrepasa el total de la orden establecida", vbCritical, Me.Caption
'        txtCant_Pedida.SetFocus
'        Exit Sub
'    End If
    
    If BuscarSaldos = True Then
        MsgBox "La cantidad a Ingresar sobrepasa el total de la orden establecida", vbCritical, Me.Caption
        txtCant_Pedida.SetFocus
        Exit Sub
    End If
        
    Call Busca_Tela_Repetida(Trim(txtCod_Tela)) '--> agregado
    ' --> If BuscaDetalleTelaRepetida = False And Trim(Txt_CantidadXTela.Text) = "" Then
    If Flg_Existe_Tela = False And Trim(Txt_CantidadXTela.Text) = "" Then
        MsgBox "Debe Ingresar la Cantidad Total Por Tela", vbInformation, "Informacion"
        FrmCantidadxTela.Left = 4440
        Txt_CantidadXTela.SetFocus
        Exit Sub
    End If
    
    lbl_Total_Detalle.Caption = CDbl(IIf(lbl_Total_Detalle.Caption = "", 0, lbl_Total_Detalle.Caption)) + CDbl(txtCant_Pedida.Text)
    
    CargaGrilla
    FrmCantidadxTela.Left = 12720
    Txt_CantidadXTela.Text = ""
    txtCod_Color.SetFocus
End Sub
Function ValidaCargaItem() As Boolean
Dim dCan_Pedida As Double, strSQL As String, SCod_Cliente_Tex As String, dCan_Item As Double

'SCod_Cliente_Tex = DevuelveCampo("select cod_cliente_tex from tx_cliente where abr_cliente='" & txtAbr_Cliente.Text & "'", cConnect)

'strSQL = "Select Isnull(Sum(Can_Pedida),0) From tx_ordcompitem_tinto where cod_cliente_tex='" & SCod_Cliente_Tex & "' and ser_ordcomp='" & txtSer_OrdComp.Text & "' and cod_ordcomp='" & txtCod_OrdComp.Text & "'"

'dCan_Pedida = DevuelveCampo(strSQL, cConnect)


dCan_Item = CDbl(IIf(lbl_Total_Detalle.Caption = "", 0, lbl_Total_Detalle.Caption)) + CDbl(txtCant_Pedida.Text)
txt_totalcarga.Text = CDbl(IIf(txt_totalcarga.Text = "", 0, txt_totalcarga.Text))

ValidaCargaItem = False

If CDbl(txt_totalcarga.Text) < dCan_Item Then
    ValidaCargaItem = True
End If

End Function

Private Sub CmdAnadir2_Click()
Dim dCantidadxTela As Double, dCant_Pedida As Double

    dCantidadxTela = IIf(Trim(Txt_CantidadXTela.Text) = "", 0, Txt_CantidadXTela.Text)
    dCant_Pedida = IIf(Trim(txtCant_Pedida.Text) = "", 0, txtCant_Pedida.Text)

    If dCantidadxTela < dCant_Pedida Then
        MsgBox "La cantidad a Ingresar sobrepasa el total de la orden establecida para esta tela", vbCritical, Me.Caption
        txtCant_Pedida.SetFocus
        Exit Sub
    Else
        Call CmdAnadir_Click
    End If

End Sub

Private Sub cmdCancelar_Click()
FrmCantidadxTela.Left = 12720
Txt_CantidadXTela.Text = ""
End Sub

Private Sub CmdEliminar_Click()
    If Reg.RecordCount = 0 Then
        MsgBox "No existen datos a Eliminar", vbCritical, Me.Caption
        Exit Sub
    End If
    
    Reg.Bookmark = gexDetalle.RowBookmark(gexDetalle.Row)
    Reg.Delete
    Set gexDetalle.ADORecordset = Reg
    ConfiguraGrid
    If Reg.RecordCount = 0 Then HabilitaCabecera
End Sub

Private Sub cmdManTela_Click()
Dim sCod_Cliente As String
    strSQL = "SELECT Cod_Cliente_Tex FROM TX_CLIENTE " & _
             "WHERE Abr_Cliente = '" & txtAbr_Cliente & "'"
    sCod_Cliente = DevuelveCampo(strSQL, cConnect)
    If Trim(sCod_Cliente) = "" Then
        MsgBox "Cliente no Valido", vbCritical + vbOKOnly, "Mant.Telas"
        Exit Sub
    End If
    frmManTxTela_Ex.sCod_Cliente = sCod_Cliente
    frmManTxTela_Ex.Show vbModal
End Sub





Private Sub dtpEntregaF_Det_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub dtpEntregaF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub dtpEntregaF_LostFocus()
    dtpEntregaF_Det.Value = dtpEntregaF.Value
End Sub

Private Sub dtpEntregaI_Det_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub dtpEntregaI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub dtpEntregaI_LostFocus()
    dtpEntregaI_Det.Value = dtpEntregaI.Value
End Sub

Private Sub Form_Load()
    txtAbr_Cliente.Text = ""
    txtNom_Cliente.Text = ""
    txtSer_OrdComp.Text = ""
    txtCod_OrdComp.Text = ""
    LimpiaCab
    
    FillMoneda
    FillTipoOC
    FillIGV
    FillPais
    FillViaTransporte
    GeneraRecorset
    ObtieneIGV
    txt_IGVDet.Text = txtPorc_IGV.Text
    sNum_Intentos = 0
    strSQL = "select count(*) from TI_Seg_Acesso_Precios where cod_usuario='" & vusu & "'"
    If DevuelveCampo(strSQL, cConnect) > 0 Then
        sAccesoPrecio = True
    Else
        sAccesoPrecio = False
    End If
    BuscaLugEntr (1)
    
    txtPrecio.Visible = sAccesoPrecio
    Label21.Visible = sAccesoPrecio
    FrmCantidadxTela.Left = 12720
    Txt_CantidadXTela.Text = ""
    
    txtSer_OrdComp.Text = DevuelveCampo("Select Ser_ordcomp From MParametros Where CodParametro='005'", cConnect)
    txtCod_OrdComp.Text = DevuelveCampo("Select Cod_ordcomp From MParametros Where CodParametro='005'", cConnect)
    
  
'    strSQL = "Select cod_cliente_tex from tx_cliente where abr_cliente='" & Trim(Me.txtAbr_Cliente.Text) & "'"
'    sCod_ClienteX = DevuelveCampo(strSQL, cConnect)


End Sub
Private Sub FillViaTransporte()

    strSQL = "SELECT idViaTransporteKey,NombreVia FROM Tx_MViaTransporte"
    
    Set rstAux5 = CargarRecordSetDesconectado(strSQL, cConnect)
    
    cmbViaTransporte_Cli.Clear
    With rstAux5
    If .RecordCount > 0 Then .MoveFirst
    Do Until .EOF
        cmbViaTransporte_Cli.AddItem !idViaTransporteKey & Space(5) & !NombreVia
        .MoveNext
    Loop

    End With

    BuscaCombo1 "C", 5, cmbViaTransporte_Cli
End Sub

Sub ObtieneIGV()
    strSQL = "Select porc_igv from tg_igv where ano='" & Year(Date) & "' and mes='" & Right("00" & Month(Date), 2) & "'"
    txtPorc_IGV.Text = DevuelveCampo(strSQL, cConnect)
End Sub

Private Sub FillMoneda()
Dim irow As Long
Dim rstAux As New ADODB.Recordset
    strSQL = "SELECT Cod_Moneda, Nom_Moneda, Flg_Principal FROM TG_Moneda"
    Set rstAux = CargarRecordSetDesconectado(strSQL, cConnect)
    
    cboMoneda.Clear
    With rstAux
    If .RecordCount > 0 Then .MoveFirst
    irow = 0
    Do Until .EOF
        cboMoneda.AddItem !Nom_Moneda & Space(100) & !Cod_Moneda
        If !Flg_Principal = "*" Then cboMoneda.ListIndex = irow
        irow = irow + 1
        .MoveNext
    Loop
    .Close
    End With
    Set rstAux = Nothing
End Sub

Private Sub FillTipoOC()
'Dim rstAux1 As New ADODB.Recordset
    'strSQL = "SELECT Cod_Tipo_Orden_tinto, Descripcion,Cod_Proceso_Tinto " & _
    '         "FROM Ti_Tipo_Orden_Tintoreria where Tip_Item='T' And Cod_Tipo_Orden_tinto='19'"
    strSQL = "EXEC Usp_Lista_TipoOrdenServicioExportacion"
    
    Set rstAux1 = CargarRecordSetDesconectado(strSQL, cConnect)
    
    cboTipoOC.Clear
    With rstAux1
    If .RecordCount > 0 Then .MoveFirst
    Do Until .EOF
        cboTipoOC.AddItem !Cod_Tipo_Orden_tinto & Space(5) & !Descripcion
        .MoveNext
    Loop
    '.Close
    End With
    'Set rstAux1 = Nothing
    BuscaCombo1 "C", 5, cboTipoOC
End Sub
Private Sub FillPais()
'Dim rstAux1 As New ADODB.Recordset
    strSQL = "select Cod_Pais,Descripcion from CN_PAISES"
    
    Set rstAux2 = CargarRecordSetDesconectado(strSQL, cConnect)
    
    cmbPais.Clear
    With rstAux2
    If .RecordCount > 0 Then .MoveFirst
    Do Until .EOF
        cmbPais.AddItem !Cod_Pais & Space(5) & !Descripcion
        .MoveNext
    Loop
    '.Close
    End With
    'Set rstAux1 = Nothing
    BuscaCombo1 "C", 5, cmbPais
End Sub
Private Sub FillIGV()
    strSQL = "SELECT Porc_IGV FROM TG_IGV " & _
             "WHERE ANO = YEAR(GETDATE()) " & _
             "AND MES = RIGHT('0' + CONVERT(VARCHAR, MONTH(GETDATE())), 2) "
    txtPorc_IGV.Text = DevuelveCampo(strSQL, cConnect)
End Sub

Public Sub BUSCA_CLIENTE(tipo As Integer)
    Select Case tipo
        Case 1:
                    strSQL = "EXEC TI_BUSCA_CLIENTE 1,'" & Trim(Me.txtAbr_Cliente.Text) & "','','" & vusu & "'"
                    Me.txtNom_Cliente.Text = Trim(DevuelveCampo(strSQL, cConnect))
                    If Trim(txtNom_Cliente.Text) <> "" Then
                        SendKeys "{TAB}"
                        SendKeys "{TAB}"
                    End If
        Case 2, 3:
                    Dim oTipo As New frmBusGeneral6
                    Dim Rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If tipo = 2 Then
                        oTipo.SQuery = "EXEC TI_BUSCA_CLIENTE 2,'','" & Trim(txtNom_Cliente.Text) & "','" & vusu & "'"
                    Else
                        oTipo.SQuery = "EXEC TI_BUSCA_CLIENTE 3,'','','" & vusu & "'"
                    End If
                    
                    oTipo.CARGAR_DATOS
                    oTipo.DGridLista.Columns(2).Width = 3500
                    oTipo.Show 1
                    If CODIGO <> "" Then
                         Me.txtAbr_Cliente.Text = Trim(CODIGO)
                         Me.txtNom_Cliente.Text = Trim(Descripcion)
'                         OptCliPend.SetFocus
                         CODIGO = "": Descripcion = ""
                         txtSer_OrdComp.SetFocus
                    End If
                    Set oTipo = Nothing
                    Set Rs = Nothing
    End Select
    scod_Cliente_x = DevuelveCampo("Select isnull(Cod_Cliente_TEx,'') from TX_Cliente where abr_Cliente = '" & txtAbr_Cliente.Text & "'", cConnect)
      cod_PaisX = DevuelveCampo("Select isnull(Cod_Pais,'') from tx_cliente where cod_Cliente_TEx = '" & scod_Cliente_x & "'", cConnect)
      If Trim(cod_PaisX) <> "" Then
            BuscaCombo cod_PaisX, 1, frmAdd_OrdComp_Ex.cmbPais
       End If
End Sub

Public Sub BUSCA_COMBO(tipo As Integer)
    Select Case tipo
        Case 1:
                    strSQL = "Select Des_Comb From Tx_TelaComb Where cod_tela='" & txtCod_Tela.Text & "' and cod_comb='" & Trim(txtCod_Comb.Text) & "'"
                    Me.txtDes_Comb.Text = Trim(DevuelveCampo(strSQL, cConnect))
                    If Trim(txtDes_Comb.Text) <> "" Then
                        SendKeys "{TAB}"
                        SendKeys "{TAB}"
                    End If
        Case 2, 3:
                    Dim oTipo As New frmBusGeneral6
                    Dim Rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If tipo = 2 Then
                        oTipo.SQuery = "Select Cod_Comb as Codigo, Des_Comb as Descripcion From Tx_TelaComb Where cod_tela='" & txtCod_Tela.Text & "' and des_comb like '%" & Trim(txtDes_Comb.Text) & "%' order by cod_comb"
                    Else
                        oTipo.SQuery = "Select Cod_Comb as Codigo, Des_Comb as Descripcion From Tx_TelaComb Where cod_tela='" & txtCod_Tela.Text & "' order by cod_comb"
                    End If
                    
                    oTipo.CARGAR_DATOS
                    oTipo.DGridLista.Columns(2).Width = 3500
                    oTipo.Show 1
                    If CODIGO <> "" Then
                         Me.txtCod_Comb.Text = Trim(CODIGO)
                         Me.txtDes_Comb.Text = Trim(Descripcion)
                         CODIGO = "": Descripcion = ""
                         txtCod_Color.SetFocus
                    End If
                    Set oTipo = Nothing
    End Select
End Sub

Public Sub BUSCA_COLOR(tipo As Integer)
    Select Case tipo
        Case 1:
            strSQL = "Select Des_Color From Lb_Color Where cod_color='" & txtCod_Color.Text & "' and isnull(cod_tipoReceta,'')<>''"
            Me.txtDes_Color.Text = Trim(DevuelveCampo(strSQL, cConnect))
            If Trim(txtDes_Color.Text) <> "" Then
                SendKeys "{TAB}"
                SendKeys "{TAB}"
            End If
        Case 2, 3:
            Dim oTipo As New frmBusGeneral6
            Dim Rs As New ADODB.Recordset
            Set oTipo.oParent = Me
            
            If tipo = 2 Then
                oTipo.SQuery = "Select Cod_Color as Codigo, Des_Color as Descripcion From lb_Color Where isnull(cod_tipoReceta,'')<>'' and des_color like '%" & Trim(txtDes_Color.Text) & "%' order by des_color"
            Else
                oTipo.SQuery = "Select Cod_Color as Codigo, Des_Color as Descripcion From Lb_Color where isnull(cod_tipoReceta,'')<>'' order by des_color"
            End If
            
            oTipo.CARGAR_DATOS
            oTipo.DGridLista.Columns(2).Width = 3500
            oTipo.Show 1
            If CODIGO <> "" Then
                 Me.txtCod_Color.Text = Trim(CODIGO)
                 Me.txtDes_Color.Text = Trim(Descripcion)
                 CODIGO = "": Descripcion = ""
                 txtCod_Receta.SetFocus
            End If
            Set oTipo = Nothing
    End Select
End Sub

Public Sub BUSCA_RECETA(tipo As Integer)
    Select Case tipo
        Case 1:
                    strSQL = "Select Descripcion From TI_Recetas_Tintoreria Where cod_color='" & txtCod_Color.Text & "' and cod_opcion='" & Trim(txtCod_Receta.Text) & "'"
                    Me.txtDes_Receta.Text = Trim(DevuelveCampo(strSQL, cConnect))
                    If Trim(txtDes_Receta.Text) <> "" Then
                        SendKeys "{TAB}"
                        SendKeys "{TAB}"
                    End If
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim Rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If tipo = 2 Then
                        oTipo.SQuery = "Select Cod_opcion as Codigo, Descripcion as Descripcion From TI_Recetas_Tintoreria Where cod_color='" & txtCod_Color.Text & "' and descripcion like '%" & Trim(txtDes_Receta.Text) & "%' order by descripcion"
                    Else
                        oTipo.SQuery = "Select Cod_opcion as Codigo, Descripcion as Descripcion From TI_Recetas_Tintoreria where cod_color='" & txtCod_Color.Text & "' order by descripcion"
                    End If
                    
                    oTipo.CARGAR_DATOS
                    oTipo.gexList.Columns(2).Width = 3500
                    oTipo.Show 1
                    If CODIGO <> "" Then
                         Me.txtCod_Receta.Text = Trim(CODIGO)
                         Me.txtDes_Receta.Text = Trim(Descripcion)
                         CODIGO = "": Descripcion = ""
                         txtCod_Talla.SetFocus
                    End If
                    Set oTipo = Nothing
    End Select
End Sub

Sub CargaDatosClientePropio()
Dim Rs_temp As New ADODB.Recordset
On Error GoTo hand
    Rs_temp.ActiveConnection = cConnect
    Rs_temp.CursorLocation = adUseClient
    strSQL = "EXEC TI_MUESTRA_LG_ORDCOMP_TINTORERIA '" & Trim(txtSer_OrdComp.Text) & "','" & Trim(txtCod_OrdComp.Text) & "'"
    Rs_temp.Open strSQL
    If Rs_temp.RecordCount Then
        Rs_temp.MoveFirst
        With Rs_temp
            sCod_ClaOrdComp = !Cod_ClaOrdComp
            txtDes_CondVent.Text = !Des_CondVent
            txtCod_CondVent.Text = !Cod_CondVent
            txtDes_CondVent.Text = !Des_CondVent
            txtCod_Descuento.Text = !cod_descuento
            txtDes_Descuento.Text = !Des_Descuento
            Call BuscaCombo1(!Cod_Moneda, 2, cboMoneda)
            txtCod_LugEntr.Text = !cod_lugentr
            txtDes_LugEntr.Text = !Des_LugEntr
            dtpEntregaI.Value = !fec_entrega_inicio
            dtpEntregaF.Value = !fec_entrega_fin
            Call BuscaCombo1(DevuelveCampo("select cod_tipo_orden_Tinto from Ti_Tipo_Orden_Tintoreria where where Tip_Item='T' AND  flg_default='*'", cConnect), 2, cboTipoOC)
        End With
        CargaGrillaClientePropio
    Else
        MsgBox "LA O.C. NO EXISTE", vbInformation, Me.Caption
        txtSer_OrdComp.Text = ""
        txtCod_OrdComp.Text = ""
    End If

    Set Rs_temp = Nothing
Exit Sub
hand:
    Set Rs_temp = Nothing
    ErrorHandler Err, Me.Caption
End Sub

Sub CargaDatosExistentes()
Dim Rs_temp As New ADODB.Recordset
Dim sCod_Cliente As String
On Error GoTo hand

    strSQL = "Select cod_cliente_tex from tx_cliente where abr_cliente='" & Trim(Me.txtAbr_Cliente.Text) & "'"
    sCod_Cliente = DevuelveCampo(strSQL, cConnect)
    
    Rs_temp.ActiveConnection = cConnect
    Rs_temp.CursorLocation = adUseClient
    strSQL = "EXEC TI_BUSCA_ORDCOMP_TINTO_EXISTENTE '" & sCod_Cliente & "','" & Trim(txtSer_OrdComp.Text) & "','" & Trim(txtCod_OrdComp.Text) & "'"
    Rs_temp.Open strSQL
    If Rs_temp.RecordCount Then
        Rs_temp.MoveFirst
        With Rs_temp
            sCod_ClaOrdComp = !Cod_ClaOrdComp
            txtDes_CondVent.Text = !Des_CondVent
            txtCod_CondVent.Text = !Cod_CondVent
            txtDes_CondVent.Text = !Des_CondVent
            txtCod_Descuento.Text = !cod_descuento
            txtDes_Descuento.Text = !Des_Descuento
            Call BuscaCombo1(!Cod_Moneda, 2, cboMoneda)
            txtCod_LugEntr.Text = !cod_lugentr
            txtDes_LugEntr.Text = !Des_LugEntr
            dtpEntregaI.Value = !fec_entrega_inicio
            dtpEntregaF.Value = !fec_entrega_fin
            Call BuscaCombo1(DevuelveCampo("select cod_tipo_orden_Tinto from Ti_Tipo_Orden_Tintoreria where  Tip_Item='T' AND flg_default='*'", cConnect), 2, cboTipoOC)
            sNum_Intentos = sNum_Intentos + 1
        End With
        DeshabilitaCabecera
    End If

    Set Rs_temp = Nothing
Exit Sub
hand:
    Set Rs_temp = Nothing
    ErrorHandler Err, Me.Caption
End Sub

Sub CargaGrillaClientePropio()
Dim Rs_tmp As New ADODB.Recordset
Dim i, j As Integer
On Error GoTo Err_CargaGrid

    strSQL = "EXEC TI_MUESTRA_LG_ORDCOMPITEM_TINTORERIA '" & Trim(txtSer_OrdComp.Text) & "','" & Trim(txtCod_OrdComp.Text) & "'"
    Rs_tmp.ActiveConnection = cConnect
    Rs_tmp.CursorLocation = adUseClient
    Rs_tmp.Open strSQL
    If Rs_tmp.RecordCount Then
        For i = 1 To Rs_tmp.RecordCount
            Reg.AddNew
            For j = 0 To Rs_tmp.Fields.Count - 1
                Reg.Fields(j).Value = Rs_tmp.Fields(j).Value
            Next
            Reg.Update
        Next
        CmdAnadir.Enabled = False
        CmdEliminar.Enabled = False
        DeshabilitaCabecera
    End If
    Set Rs_tmp = Nothing
    
    Set gexDetalle.ADORecordset = Reg
    ConfiguraGrid
Exit Sub
Err_CargaGrid:
    Set Rs_tmp = Nothing
    ErrorHandler Err, "Err_CargaGrid"
End Sub

Sub CargaGrilla()
On Error GoTo Err_CargaGrid
    Reg.AddNew
    Reg.Fields("cod_tela").Value = Trim(txtCod_Tela.Text)
    Reg.Fields("tela").Value = Trim(txtCod_Tela.Text) & " - " & Trim(txtDes_Tela.Text)
    Reg.Fields("cod_comb").Value = Trim(txtCod_Comb.Text)
    Reg.Fields("combinacion").Value = Trim(txtCod_Comb.Text) & " - " & Trim(txtDes_Comb.Text)
    Reg.Fields("cod_color").Value = Trim(txtCod_Color.Text)
    Reg.Fields("color").Value = Trim(txtCod_Color.Text) & " - " & Trim(txtDes_Color.Text)
    Reg.Fields("cod_receta").Value = Trim(txtCod_Receta.Text)
    Reg.Fields("Receta").Value = Trim(txtDes_Receta.Text)
    Reg.Fields("talla").Value = Trim(txtCod_Talla.Text)
    Reg.Fields("cod_descuento").Value = Trim(txtCod_DsctoDet.Text)
    Reg.Fields("descuento").Value = Trim(txtDes_DstoDet.Text)
    Reg.Fields("igv").Value = CDbl(txt_IGVDet.Text)
    Reg.Fields("Fec.Inicio").Value = dtpEntregaI_Det.Value
    Reg.Fields("Fec.Fin").Value = dtpEntregaF_Det.Value
    Reg.Fields("Pedida").Value = CDbl(txtCant_Pedida.Text)
    Reg.Fields("P.U.").Value = CDbl(txtPrecio.Text)
    Reg.Fields("cod_tela_cliente").Value = Trim(txtCod_TelaCliente.Text)
    Reg.Fields("observaciones").Value = txtObs_Det.Text
    Reg.Fields("Unidades").Value = CDbl(Txt_Unidades.Text)
    Reg.Fields("TotalxTela").Value = IIf(Trim(Txt_CantidadXTela.Text) = "", 0, Txt_CantidadXTela.Text)
    
    Reg.Update
    busca_Diferencia_Tela (Trim(txtCod_Tela.Text))
    Set gexDetalle.ADORecordset = Reg
    ConfiguraGrid
    DeshabilitaCabecera
    LimpiaDetalle
Exit Sub
Err_CargaGrid:
    ErrorHandler Err, "Err_CargaGrid"
End Sub
Sub busca_Diferencia_Tela(Scod_TelaX As String)
Dim REGX As ADODB.Recordset
Dim Total_Ingresado As Double
Dim total_Cargado As Double
Dim total_Diferencia As Double
Dim i As Integer
Set REGX = New ADODB.Recordset
Set REGX = gexDetalle.ADORecordset
total_Cargado = 0
Total_Ingresado = 0
total_Diferencia = 0
If REGX.RecordCount > 0 Then
REGX.MoveFirst
    For i = 1 To REGX.RecordCount
    'scod_Tela_x = gexDetalle.Value(gexDetalle.Columns("Cod_Tela").Index)
            If Trim(REGX.Fields("cod_tela").Value) = Trim(Scod_TelaX) Then
                If REGX.Fields("TotalxTela").Value <> "0" Then
                    Total_Ingresado = CDbl(REGX.Fields("TotalxTela").Value)
                End If
                total_Cargado = total_Cargado + CDbl(REGX.Fields("Pedida").Value)
            End If
            
            REGX.MoveNext
    Next
End If
lbl_Total_Detalle = Total_Ingresado - total_Cargado
End Sub

Sub Busca_Tela_Repetida(Scod_TelaX As String)
Dim REGX As ADODB.Recordset
Set REGX = New ADODB.Recordset
Set REGX = gexDetalle.ADORecordset
Dim i As Integer
Flg_Existe_Tela = False
If REGX.RecordCount > 0 Then
REGX.MoveFirst
    For i = 1 To REGX.RecordCount
    'scod_Tela_x = gexDetalle.Value(gexDetalle.Columns("Cod_Tela").Index)
            If Trim(REGX.Fields("cod_tela").Value) = Trim(Scod_TelaX) Then
            Flg_Existe_Tela = True
            Exit For
            End If
            
            REGX.MoveNext
    Next
End If
'lbl_Total_Detalle = Total_Ingresado - total_Cargado
End Sub

Sub DeshabilitaCabecera()
    txtAbr_Cliente.Enabled = False
    txtNom_Cliente.Enabled = False
    txtSer_OrdComp.Enabled = False
    txtCod_OrdComp.Enabled = False
    txtCod_CondVent.Enabled = False
    txtDes_CondVent.Enabled = False
    txtCod_Descuento.Enabled = False
    txtDes_Descuento.Enabled = False
    txtCod_LugEntr.Enabled = False
    txtDes_LugEntr.Enabled = False
    dtpEntregaI.Enabled = False
    dtpEntregaF.Enabled = False
    cboMoneda.Enabled = False
    'cboTipoOC.Enabled = False
    txtObservaciones.Enabled = False
End Sub

Sub HabilitaCabecera()
    txtAbr_Cliente.Enabled = True
    txtNom_Cliente.Enabled = True
    txtSer_OrdComp.Enabled = True
    txtCod_OrdComp.Enabled = True
    txtCod_CondVent.Enabled = True
    txtDes_CondVent.Enabled = True
    txtCod_Descuento.Enabled = True
    txtDes_Descuento.Enabled = True
    txtCod_LugEntr.Enabled = True
    txtDes_LugEntr.Enabled = True
    dtpEntregaI.Enabled = True
    dtpEntregaF.Enabled = True
    cboMoneda.Enabled = True
    cboTipoOC.Enabled = True
    txtObservaciones.Enabled = True
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "GRABAR"
            If VALIDA_DATOS Then
                DeshabilitaCabecera
                If SALVAR_CABECERA Then
                    sNum_Intentos = sNum_Intentos + 1
                    SALVAR_DETALLE
                End If
            End If
        Case "CANCELAR"
            Unload Me
    End Select
End Sub





Private Sub gexDetalle_Click()

scod_Tela_x = gexDetalle.Value(gexDetalle.Columns("Cod_Tela").Index)
'DGridLista.Value(DGridLista.Columns("Cod_Almacen").Index)
'MsgBox scod_Tela_x
busca_Diferencia_Tela (scod_Tela_x)
lbl_et_Total_Detalle.Caption = "Kilos restantes para la tela " & scod_Tela_x
End Sub

Private Sub txt_IGVDet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        dtpEntregaI_Det.SetFocus
    Else
        Call SoloNumeros(txt_IGVDet, KeyAscii, True, 2)
    End If
End Sub

Private Sub Txt_Unidades_GotFocus()
SelectionText Txt_Unidades
End Sub

Private Sub Txt_Unidades_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_Unidades.SetFocus
    Else
        Call SoloNumeros(Txt_Unidades, KeyAscii, True, 2)
    End If

End Sub

Private Sub txtAbr_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtAbr_Cliente.Text) = "" Then
            BUSCA_CLIENTE 3
        Else
            BUSCA_CLIENTE 1
        End If
    End If
End Sub

Private Sub txtCant_Pedida_GotFocus()
    SelectionText txtCant_Pedida
End Sub

Private Sub txtCod_Color_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_Color.Text) = "" Then
            BUSCA_COLOR 3
        Else
            BUSCA_COLOR 1
        End If
    End If
End Sub

Private Sub txtCod_Comb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_Tela.Text) = "" Then
            MsgBox "El Codigo de Tela no puede estar vacio, Verifique", vbCritical, Me.Caption
            txtCod_Tela.SetFocus
            Exit Sub
        End If
        
        If Trim(txtCod_Comb.Text) = "" Then
            BUSCA_COMBO 3
        Else
            BUSCA_COMBO 1
        End If
    End If
End Sub

Private Sub txtCod_DsctoDet_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then BuscaDescuento 1
End Sub

Private Sub txtCod_LugEntr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then BuscaLugEntr 1
End Sub

Private Sub txtCod_OrdComp_GotFocus()
    If Trim(txtSer_OrdComp.Text) = "" Then
        MsgBox "Ingrese el numero de Serie de la OC"
        txtSer_OrdComp.SetFocus
    Else
        txtSer_OrdComp = Format(txtSer_OrdComp, "000")
    End If
End Sub

Private Sub txtCod_Receta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_Color.Text) = "" Then
            MsgBox "Selecciones 1ero el Color", vbInformation, Me.Caption
            txtCod_Color.SetFocus
            Exit Sub
        End If
        
        If Trim(txtCod_Receta.Text) = "" Then
            BUSCA_RECETA 3
        Else
            BUSCA_RECETA 1
        End If
    End If
End Sub

Private Sub txtCod_Talla_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtCod_Tela_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim oTipo As New frmBuscaTela
    Dim sCod_Cliente As String
    
        Set oTipo.oParent = Me
        
        If Len(Trim(txtCod_Tela)) > 2 Then
'            Dim Temp As String
'            Temp = Trim(DevuelveCampo("Select dbo.uf_devuelvecodigo(6," & IIf(Trim(txtCod_Tela.Text) = "", 0, Mid(txtCod_Tela.Text, 3)) & ")", cCONNECT))
             txtCod_Tela = Left(txtCod_Tela, 2) & Format(Mid(txtCod_Tela, 3), "000000")
        End If
        
        strSQL = "Select cod_cliente_tex from tx_cliente where abr_cliente='" & Me.txtAbr_Cliente.Text & "'"
        sCod_Cliente = DevuelveCampo(strSQL, cConnect)
        
        oTipo.sCod_Cliente = sCod_Cliente
        oTipo.sCod_Tela = Trim(txtCod_Tela.Text)
        If Trim(txtCod_Tela.Text) = "" Then oTipo.ChkAllClient.Visible = False
        oTipo.Campo = 1
        oTipo.CARGAR_DATOS
        'oTipo.DGridLista.Columns(2).Width = 3500
        oTipo.Show 1
        If CODIGO <> "" Then
             Me.txtCod_Tela.Text = Trim(CODIGO)
             Me.txtDes_Tela.Text = Trim(Descripcion)
             CODIGO = "": Descripcion = ""
             SendKeys "{TAB}"
             SendKeys "{TAB}"
        End If
        Set oTipo = Nothing
End If
End Sub

Private Sub txtCod_TelaCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtDes_Color_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtDes_Color.Text) = "" Then
            BUSCA_COLOR 3
        Else
            BUSCA_COLOR 2
        End If
    End If
End Sub

Private Sub txtDes_Comb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_Tela.Text) = "" Then
            MsgBox "El Codigo de Tela no puede estar vacio, Verifique", vbCritical, Me.Caption
            txtCod_Tela.SetFocus
            Exit Sub
        End If
        
        If Trim(txtDes_Comb.Text) = "" Then
            BUSCA_COMBO 3
        Else
            BUSCA_COMBO 2
        End If
    End If
End Sub

Private Sub txtDes_DstoDet_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then BuscaDescuento 2
End Sub

Private Sub txtDes_LugEntr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then BuscaLugEntr 2
End Sub

Private Sub txtCant_Pedida_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtPrecio.Visible Then
            txtPrecio.SetFocus
        Else
            Txt_Unidades.SetFocus
        End If
    Else
        Call SoloNumeros(txtCant_Pedida, KeyAscii, True, 2)
    End If
End Sub

Private Sub txtCod_CondVent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then BuscaCondPago 1
End Sub

Private Sub Txtcod_Descuento_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then BuscaDescuento 1
End Sub

Private Sub TxtDes_CondVent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then BuscaCondPago 2
End Sub

Private Sub TxtDes_Descuento_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then BuscaDescuento 2
End Sub

Private Sub txtCod_OrdComp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCod_OrdComp_LostFocus()
If Trim(txtSer_OrdComp.Text) <> "" Then
    txtCod_OrdComp = Format(txtCod_OrdComp, "000000")
    
    'Limpiamos los datos
    LimpiaCab
    FillMoneda
    LimpiaDetalle
    sCod_ClaOrdComp = ""
    If Not Reg.EOF And Not Reg.BOF Then
        Reg.MoveFirst
        While Not Reg.EOF
            Reg.Delete
            Reg.MoveFirst
        Wend
    End If
    Set gexDetalle.ADORecordset = Reg
    ConfiguraGrid
    
    strSQL = "Select flg_ClientePropio from tx_cliente where abr_cliente='" & Trim(Me.txtAbr_Cliente.Text) & "'"
    If Trim(DevuelveCampo(strSQL, cConnect)) = "*" Then
        CargaDatosClientePropio
    Else
        CargaDatosExistentes
    End If
End If
End Sub


Private Sub txtDes_Receta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_Color.Text) = "" Then
            MsgBox "Selecciones 1ero el Color", vbInformation, Me.Caption
            txtCod_Color.SetFocus
            Exit Sub
        End If
    
        If Trim(txtDes_Receta.Text) = "" Then
            BUSCA_RECETA 3
        Else
            BUSCA_RECETA 2
        End If
    End If
End Sub

Private Sub txtNom_Cliente_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
        If Trim(txtNom_Cliente.Text) = "" Then
            BUSCA_CLIENTE 3
        Else
            BUSCA_CLIENTE 2
        End If
    End If
End Sub

Private Sub txtPrecio_GotFocus()
    SelectionText txtPrecio
End Sub

Private Sub TxtPrecio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        Call SoloNumeros(txtPrecio, KeyAscii, True, 3)
    End If
End Sub

Private Sub txtSer_OrdComp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtCod_OrdComp.SetFocus
    End If
End Sub

Private Sub BuscaCondPago(Opcion As Integer)
Dim rstAux As New ADODB.Recordset
On Error GoTo Fin
    strSQL = "SELECT Cod_CondVent, Des_CondVent " & _
             "FROM Lg_CondVent WHERE "
    txtCod_CondVent = Trim(txtCod_CondVent)
    txtDes_CondVent = Trim(txtDes_CondVent)
    Select Case Opcion
    Case 1: strSQL = strSQL & "Cod_CondVent like '%" & txtCod_CondVent & "%'"
    Case 2: strSQL = strSQL & "Des_CondVent like '%" & txtDes_CondVent & "%'"
    End Select
    txtCod_CondVent = ""
    txtDes_CondVent = ""
    With frmBusGeneral6
        Set .oParent = Me
        .SQuery = strSQL
        .CARGAR_DATOS
        CODIGO = ".."
        Set rstAux = .DGridLista.ADORecordset
        
        .DGridLista.Columns("Cod_CondVent").Caption = "Codigo"
        .DGridLista.Columns("Cod_CondVent").Width = 700
        .DGridLista.Columns("Des_CondVent").Caption = "Cond.Venta"
        .DGridLista.Columns("Des_CondVent").Width = 5000
        
        If rstAux.RecordCount > 1 Then
            rstAux.MoveFirst
            .Show vbModal
        End If
        
        If CODIGO <> "" And rstAux.RecordCount > 0 Then
            txtCod_CondVent = Trim(rstAux!Cod_CondVent)
            txtDes_CondVent = Trim(rstAux!Des_CondVent)
            SendKeys "{TAB}"
        End If
        SendKeys "{TAB}"
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Busqueda de Cliente (" & Opcion & ")"
End Sub

Private Sub BuscaDescuento(Opcion As Integer)
Dim rstAux As New ADODB.Recordset
On Error GoTo Fin
    strSQL = "SELECT Cod_Descuento, Des_Descuento, Porcentaje1 " & _
             "FROM LG_DSCTOS WHERE "
    txtCod_Descuento = Trim(txtCod_Descuento)
    txtDes_Descuento = Trim(txtDes_Descuento)
    Select Case Opcion
    Case 1: strSQL = strSQL & "Cod_Descuento like '%" & txtCod_Descuento & "%'"
    Case 2: strSQL = strSQL & "Des_Descuento like '%" & txtDes_Descuento & "%'"
    End Select
    txtCod_Descuento = ""
    txtDes_Descuento = ""
    txtDes_Descuento.Tag = ""
    With frmBusGeneral6
        Set .oParent = Me
        .SQuery = strSQL
        .CARGAR_DATOS
        
        .DGridLista.Columns("Cod_Descuento").Caption = "Codigo"
        .DGridLista.Columns("Cod_Descuento").Width = 700
        .DGridLista.Columns("Des_Descuento").Caption = "Descuento"
        .DGridLista.Columns("Des_Descuento").Width = 5000
        .DGridLista.Columns("Porcentaje1").Visible = False
        
        CODIGO = ".."
        Set rstAux = .DGridLista.ADORecordset
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If CODIGO <> "" And rstAux.RecordCount > 0 Then
            txtCod_Descuento = Trim(rstAux!cod_descuento)
            txtDes_Descuento = Trim(rstAux!Des_Descuento)
            txtCod_DsctoDet = Trim(rstAux!cod_descuento)
            txtDes_DstoDet = Trim(rstAux!Des_Descuento)
            txtDes_Descuento.Tag = rstAux!Porcentaje1
            SendKeys "{TAB}"
        End If
        SendKeys "{TAB}"
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Busqueda de Descuento (" & Opcion & ")"
End Sub

Public Sub BuscaLugEntr(Opcion As Integer)
Dim rstAux As New ADODB.Recordset
On Error GoTo Fin
Dim scod_Clientex As String
        strSQL = "Select cod_cliente_tex from tx_cliente where abr_cliente='" & Me.txtAbr_Cliente.Text & "'"
        scod_Clientex = DevuelveCampo(strSQL, cConnect)

    strSQL = "SELECT Cod_LugEntr, Des_LugEntr FROM LG_LUGENTR WHERE cod_cliente_tex = '" & scod_Clientex & "' and "
    
    txtCod_LugEntr = Trim(txtCod_LugEntr)
    txtDes_LugEntr = Trim(txtDes_LugEntr)
    Select Case Opcion
    Case 1: strSQL = strSQL & "Cod_LugEntr like '%" & txtCod_LugEntr & "%'"
    Case 2: strSQL = strSQL & "Des_LugEntr like '%" & txtDes_LugEntr & "%'"
    End Select
    txtCod_LugEntr = ""
    txtDes_LugEntr = ""
    With frmBusGeneral6
        Set .oParent = Me
        .SQuery = strSQL
        .CARGAR_DATOS
        CODIGO = ".."
        Set rstAux = .DGridLista.ADORecordset
        
        .DGridLista.Columns("Cod_LugEntr").Caption = "Codigo"
        .DGridLista.Columns("Cod_LugEntr").Width = 700
        .DGridLista.Columns("Des_LugEntr").Caption = "Lugar de Entrega"
        .DGridLista.Columns("Des_LugEntr").Width = 5000
        
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If CODIGO <> "" And rstAux.RecordCount > 0 Then
            txtCod_LugEntr = Trim(rstAux!cod_lugentr)
            txtDes_LugEntr = Trim(rstAux!Des_LugEntr)
            SendKeys "{TAB}"
        End If
        SendKeys "{TAB}"
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Busqueda de Lugar de Entrega (" & Opcion & ")"
End Sub

Private Sub FillOrdComp()
Dim vCod_Cliente As String
Dim sCod_Moneda As String
Dim rstOrdComp As ADODB.Recordset
    
    strSQL = "Select cod_cliente_tex from tx_cliente where abr_cliente='" & Me.txtAbr_Cliente.Text & "'"
    vCod_Cliente = DevuelveCampo(strSQL, cConnect)
    
    strSQL = "EXEC TI_BUSCA_ORDCOMP_TINTO '" & vCod_Cliente & "', '" & _
             txtSer_OrdComp & "', '" & txtCod_OrdComp & "'"
    Set rstOrdComp = CargarRecordSetDesconectado(strSQL, cConnect)
    With rstOrdComp
        If .RecordCount = 0 Then
            GoTo Search
        End If
        txtCod_CondVent = !Cod_CondVent
        txtCod_Descuento = !cod_descuento
        
        sCod_Moneda = !Cod_Moneda
        
        txtCod_LugEntr = !cod_lugentr
        dtpEntregaI = !fec_entrega_inicio
        dtpEntregaF = !fec_entrega_fin
        .Close
    End With
    Set rstOrdComp = Nothing
Search:
    BuscaCombo1 sCod_Moneda, 100, cboMoneda
    BuscaCondPago 1
    BuscaDescuento 1
    BuscaLugEntr 1
    dtpEntregaI.SetFocus
End Sub


Sub GeneraRecorset()
    Set Reg = New ADODB.Recordset
    Reg.ActiveConnection = Nothing
    Reg.CursorType = adOpenStatic
    Reg.CursorLocation = adUseClient
    
    Reg.Fields.Append "Cod_Tela", adVarChar, 8
    Reg.Fields.Append "Tela", adVarChar, 200
    Reg.Fields.Append "Cod_Comb", adVarChar, 3
    Reg.Fields.Append "Combinacion", adVarChar, 100
    Reg.Fields.Append "Cod_Color", adVarChar, 6
    Reg.Fields.Append "Color", adVarChar, 80
    Reg.Fields.Append "Cod_Receta", adVarChar, 3
    Reg.Fields.Append "Receta", adVarChar, 80
    Reg.Fields.Append "Talla", adVarChar, 10
    Reg.Fields.Append "Cod_Descuento", adVarChar, 3
    Reg.Fields.Append "Descuento", adVarChar, 80
    Reg.Fields.Append "IGV", adDouble
    Reg.Fields.Append "Fec.Inicio", adDate
    Reg.Fields.Append "Fec.Fin", adDate
    Reg.Fields.Append "P.U.", adDouble
    Reg.Fields.Append "Pedida", adDouble
    Reg.Fields.Append "Cod_Tela_Cliente", adVarChar, 20
    Reg.Fields.Append "Observaciones", adLongVarChar, 1000
    Reg.Fields.Append "Unidades", adDouble
    Reg.Fields.Append "TotalxTela", adDouble
    Reg.Open
    
    Set gexDetalle.ADORecordset = Reg
    ConfiguraGrid
End Sub

Sub ConfiguraGrid()
    gexDetalle.Columns("cod_tela").Visible = False
    gexDetalle.Columns("cod_comb").Visible = False
    gexDetalle.Columns("cod_color").Visible = False
    gexDetalle.Columns("cod_receta").Visible = False
    gexDetalle.Columns("cod_descuento").Visible = False
    
    gexDetalle.Columns("P.U.").Visible = sAccesoPrecio
End Sub

Function VALIDA_DATOS() As Boolean
Dim sCod_Cliente As String
    VALIDA_DATOS = True
    If Trim(txtAbr_Cliente.Text) = "" Then
        MsgBox "Tiene que Seleccionar el Cliente", vbCritical, "Valida O.C."
        VALIDA_DATOS = False
        Exit Function
    End If

    If Trim(txtSer_OrdComp) = "" Then
        MsgBox "Ingrese la Serie de la OC", vbCritical, "Valida O.C."
        VALIDA_DATOS = False
        Exit Function
    End If

    If Trim(txtCod_OrdComp.Text) = "" Then
        MsgBox "Ingrese el Codigo de OC", vbCritical, "Valida O.C."
        VALIDA_DATOS = False
        Exit Function
    End If
    
    If Trim(cmbPais.Text) = "" Then
        MsgBox "Debe Seleccionar un Pais ", vbCritical, "Valida O.C."
        VALIDA_DATOS = False
        Exit Function
    End If
    strSQL = "Select flg_ClientePropio from tx_cliente where abr_cliente='" & Trim(Me.txtAbr_Cliente.Text) & "'"
    If Trim(DevuelveCampo(strSQL, cConnect)) <> "*" And sNum_Intentos = 0 Then
        strSQL = "Select cod_cliente_tex from tx_cliente where abr_cliente='" & Trim(Me.txtAbr_Cliente.Text) & "'"
        sCod_Cliente = DevuelveCampo(strSQL, cConnect)
        
        strSQL = "Select count(*) from TX_OrdComp where Cod_Cliente_Tex='" & sCod_Cliente & "' and Ser_OrdComp='" & Trim(txtSer_OrdComp.Text) & "' and Cod_OrdComp='" & Trim(txtCod_OrdComp) & "'"
        
        If DevuelveCampo(strSQL, cConnect) > 0 Then
            strSQL = "Select isnull(cod_tipoOC_Tintoreria,'') as TipoOC from TX_OrdComp where Cod_Cliente_Tex='" & sCod_Cliente & "' and Ser_OrdComp='" & Trim(txtSer_OrdComp.Text) & "' and Cod_OrdComp='" & Trim(txtCod_OrdComp) & "'"
            If Trim(DevuelveCampo(strSQL, cConnect)) = "" Then
                MsgBox "La OC a sido creada solo para Tejeduria ", vbInformation, Me.Caption
            Else
                MsgBox "La OC ya Existe", vbInformation, Me.Caption
            End If
            
            txtSer_OrdComp.Enabled = True
            txtCod_OrdComp.Enabled = True
            VALIDA_DATOS = False
            txtSer_OrdComp.SetFocus
            Exit Function
        End If
    End If

    If Trim(txtCod_CondVent.Text) = "" Then
        MsgBox "Ingrese el Tipo de Pago", vbCritical, "Valida O.C."
        VALIDA_DATOS = False
        Exit Function
    End If

    If Trim(txtCod_Descuento.Text) = "" Then
        MsgBox "Ingrese el Descuento", vbCritical, "Valida O.C."
        VALIDA_DATOS = False
        Exit Function
    End If

    If Trim(txtCod_LugEntr.Text) = "" Then
        MsgBox "Ingrese el Lugar de Entrega", vbCritical, "Valida O.C."
        VALIDA_DATOS = False
        Exit Function
    End If

    If Trim(cboTipoOC.Text) = "" Then
        MsgBox "Seleccione el Tipo de OC", vbCritical, "Valida O.C."
        VALIDA_DATOS = False
        Exit Function
    End If
    
    If Reg.RecordCount = 0 Then
        MsgBox "La OC no tiene Detalle", vbCritical, "Valida O.C."
        VALIDA_DATOS = False
        Exit Function
    End If

End Function

Function SALVAR_CABECERA() As Boolean
On Error GoTo hand
    SALVAR_CABECERA = True
    strSQL = "EXEC TI_MAN_TX_ORDCOMP_EXP 'I','" & _
        DevuelveCampo("select cod_cliente_tex from tx_cliente where abr_cliente='" & txtAbr_Cliente.Text & "'", cConnect) & "','" & _
        txtSer_OrdComp.Text & "','" & _
        txtCod_OrdComp.Text & "','" & _
        txtCod_CondVent.Text & "','" & _
        txtCod_Descuento.Text & "'," & _
        txtPorc_IGV.Text & ",'" & _
        Trim(Right(cboMoneda, 5)) & "','" & _
        txtCod_LugEntr.Text & "','" & _
        Trim(txtObservaciones.Text) & "','" & _
        Trim(sCod_ClaOrdComp) & "','" & _
        dtpEntregaI.Value & "','" & _
        dtpEntregaF.Value & "','" & _
        Trim(Left(cboTipoOC, 2)) & "','" & _
        IIf(ChkCtrPeso.Value = 1, "S", "N") & "'," & _
        CDbl(gexDetalle.Value(gexDetalle.Columns("TotalxTela").Index)) & ",'" & _
        Left(cmbPais.Text, 4) & "','" & _
        Trim(TxtPO.Text) & "','" & _
        Left(cmbViaTransporte_Cli.Text, 2) & "'"

    Call ExecuteSQL(cConnect, strSQL)
    
Exit Function
hand:
    SALVAR_CABECERA = False
    ErrorHandler Err, "SALVAR_CABECERA"
End Function

Sub SALVAR_DETALLE()
Dim i As Integer
Dim sRows As Integer
On Error GoTo hand
    sRows = Reg.RecordCount
    Reg.MoveFirst
    With Reg
        For i = 1 To sRows
            strSQL = "EXEC TI_MAN_TX_ORDCOMPITEM_TINTO_EX 'I','" & _
            DevuelveCampo("select cod_cliente_tex from tx_cliente where abr_cliente='" & txtAbr_Cliente.Text & "'", cConnect) & "','" & _
            txtSer_OrdComp.Text & "','" & _
            txtCod_OrdComp.Text & "','','" & _
            Trim(!Cod_Tela) & "','" & _
            Trim(!cod_comb) & "','" & _
            Trim(!cod_Color) & "','" & _
            Trim(!cod_Receta) & "','" & _
            Trim(!talla) & "','" & _
            Trim(!cod_descuento) & "'," & _
            !Igv & ",'" & _
            .Fields("Fec.Inicio").Value & "','" & _
            .Fields("Fec.Fin").Value & "'," & _
            .Fields("P.U.").Value & "," & _
            !Pedida & ",'" & _
            Trim(!Cod_tela_cliente) & "','" & _
            Trim(!OBSERVACIONES) & "','" & .Fields("UNIDADES").Value & "','" & sCadena_Fam & "'," & .Fields("TotalxTela").Value & ""
            
            'MsgBox "Ok"
            Call ExecuteSQL(cConnect, strSQL)
            'Reg.MoveNext
            Reg.Delete
            Reg.MoveFirst
        Next
    End With
    
    Unload Me
Exit Sub
hand:
    ErrorHandler Err, "SALVAR_CABECERA"
    Set gexDetalle.ADORecordset = Reg
End Sub


Sub GeneraRecorset_Proceso()
    Set RegProceso = New ADODB.Recordset
    RegProceso.ActiveConnection = Nothing
    RegProceso.CursorType = adOpenStatic
    RegProceso.CursorLocation = adUseClient
    
    RegProceso.Fields.Append "Cod_Cliente_Tex", adVarChar, 5
    RegProceso.Fields.Append "Ser_OrdComp", adVarChar, 3
    RegProceso.Fields.Append "Cod_OrdComp", adVarChar, 6
    RegProceso.Fields.Append "CSec_OrdComp", adVarChar, 3
    RegProceso.Fields.Append "Cod_Proceso_Tinto", adVarChar, 2
    RegProceso.Open
    
    Set gexDetalle.ADORecordset = Reg
    ConfiguraGrid
End Sub


