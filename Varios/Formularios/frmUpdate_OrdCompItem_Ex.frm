VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "mantenimientos.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUpdate_OrdCompItem_Ex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle O.C."
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdManTela 
      Caption         =   "cmdManTela"
      Height          =   375
      Left            =   6360
      TabIndex        =   38
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detalle"
      Height          =   3045
      Left            =   45
      TabIndex        =   20
      Top             =   3000
      Width           =   8550
      Begin VB.Frame FrmCantidadxTela 
         Caption         =   "Ingrese la Cantidad Total Por Tela"
         Height          =   1335
         Left            =   2880
         TabIndex        =   40
         Top             =   600
         Width           =   3015
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
            TabIndex        =   43
            Top             =   240
            Width           =   1770
         End
         Begin VB.CommandButton CmdAnadir2 
            Caption         =   "Añadir"
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   720
            Width           =   1365
         End
         Begin VB.CommandButton CmdCancelar 
            Caption         =   "Cancelar"
            Height          =   495
            Left            =   1680
            TabIndex        =   41
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.TextBox Txt_Unidades 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7440
         TabIndex        =   37
         Text            =   "0"
         Top             =   1560
         Width           =   510
      End
      Begin VB.TextBox txtNumSec 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7875
         TabIndex        =   35
         Text            =   "0"
         Top             =   645
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.TextBox txtDes_Color 
         Height          =   285
         Left            =   1740
         TabIndex        =   5
         Top             =   915
         Width           =   2535
      End
      Begin VB.TextBox txtDes_Comb 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1725
         TabIndex        =   3
         Top             =   570
         Width           =   2535
      End
      Begin VB.TextBox txtCod_Comb 
         Enabled         =   0   'False
         Height          =   285
         Left            =   990
         MaxLength       =   3
         TabIndex        =   2
         Top             =   570
         Width           =   735
      End
      Begin VB.TextBox txtDes_Tela 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   210
         Width           =   2220
      End
      Begin VB.TextBox txtCod_Tela 
         Enabled         =   0   'False
         Height          =   285
         Left            =   990
         MaxLength       =   8
         TabIndex        =   0
         Top             =   210
         Width           =   1050
      End
      Begin VB.TextBox txtCod_Color 
         Height          =   285
         Left            =   990
         MaxLength       =   6
         TabIndex        =   4
         Top             =   915
         Width           =   735
      End
      Begin VB.TextBox txtCod_Talla 
         Enabled         =   0   'False
         Height          =   285
         Left            =   990
         TabIndex        =   8
         Top             =   1560
         Width           =   750
      End
      Begin VB.TextBox txtCod_DsctoDet 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5625
         TabIndex        =   10
         Top             =   225
         Width           =   525
      End
      Begin VB.TextBox txtDes_DstoDet 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6195
         TabIndex        =   11
         Top             =   225
         Width           =   2220
      End
      Begin VB.TextBox txt_IGVDet 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5625
         TabIndex        =   12
         Top             =   540
         Width           =   510
      End
      Begin VB.TextBox txtObs_Det 
         Enabled         =   0   'False
         Height          =   570
         Left            =   990
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   2265
         Width           =   7410
      End
      Begin VB.TextBox txtCant_Pedida 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   5610
         TabIndex        =   15
         Text            =   "0"
         Top             =   1590
         Width           =   690
      End
      Begin VB.TextBox txtPrecio 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   990
         TabIndex        =   16
         Text            =   "0"
         Top             =   1935
         Width           =   720
      End
      Begin VB.TextBox txtCod_TelaCliente 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3045
         TabIndex        =   9
         Top             =   1575
         Width           =   1230
      End
      Begin VB.TextBox txtCod_Receta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   990
         MaxLength       =   2
         TabIndex        =   6
         Top             =   1245
         Width           =   735
      End
      Begin VB.TextBox txtDes_Receta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1725
         TabIndex        =   7
         Top             =   1245
         Width           =   2550
      End
      Begin MSComCtl2.DTPicker dtpEntregaI_Det 
         Height          =   315
         Left            =   6135
         TabIndex        =   13
         Top             =   870
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   74514435
         CurrentDate     =   37832
      End
      Begin MSComCtl2.DTPicker dtpEntregaF_Det 
         Height          =   315
         Left            =   6135
         TabIndex        =   14
         Top             =   1215
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   74514435
         CurrentDate     =   37832
      End
      Begin VB.Label lbl_CanpedidaSinMod 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3120
         TabIndex        =   39
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Uni.Pedida"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   6480
         TabIndex        =   36
         Top             =   1680
         Width           =   780
      End
      Begin VB.Label Label31 
         Caption         =   "Combo"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   105
         TabIndex        =   34
         Top             =   645
         Width           =   900
      End
      Begin VB.Label Label11 
         Caption         =   "Tela"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   90
         TabIndex        =   33
         Top             =   300
         Width           =   690
      End
      Begin VB.Label Label12 
         Caption         =   "Color"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   105
         TabIndex        =   32
         Top             =   975
         Width           =   690
      End
      Begin VB.Label Label13 
         Caption         =   "Talla"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   1635
         Width           =   540
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000013&
         Caption         =   "Descuento"
         ForeColor       =   &H80000002&
         Height          =   210
         Left            =   4530
         TabIndex        =   30
         Top             =   255
         Width           =   900
      End
      Begin VB.Label Label15 
         Caption         =   "I.G.V."
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   4545
         TabIndex        =   29
         Top             =   615
         Width           =   495
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000013&
         Caption         =   "Fec. Entrega Inicio"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   4515
         TabIndex        =   28
         Top             =   930
         Width           =   1530
      End
      Begin VB.Label Label17 
         BackColor       =   &H80000013&
         Caption         =   "Fec. Entrega Fin"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   4515
         TabIndex        =   27
         Top             =   1260
         Width           =   1530
      End
      Begin VB.Label Label18 
         Caption         =   "%"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   6210
         TabIndex        =   26
         Top             =   615
         Width           =   270
      End
      Begin VB.Label Label19 
         BackColor       =   &H80000013&
         Caption         =   "Observ:"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   120
         TabIndex        =   25
         Top             =   2325
         Width           =   810
      End
      Begin VB.Label Label20 
         Caption         =   "Cant.Pedida"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   4530
         TabIndex        =   24
         Top             =   1620
         Width           =   915
      End
      Begin VB.Label Label21 
         Caption         =   "Precio"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   120
         TabIndex        =   23
         Top             =   1995
         Width           =   540
      End
      Begin VB.Label Label22 
         Caption         =   "Cod.Tela Cliente"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   1830
         TabIndex        =   22
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label Label24 
         Caption         =   "Receta"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1305
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2910
      Left            =   45
      TabIndex        =   18
      Top             =   75
      Width           =   8535
      Begin GridEX20.GridEX gexDetalle 
         Height          =   2670
         Left            =   90
         TabIndex        =   19
         Top             =   150
         Width           =   8340
         _ExtentX        =   14711
         _ExtentY        =   4710
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         FormatStylesCount=   7
         FormatStyle(1)  =   "frmUpdate_OrdCompItem_Ex.frx":0000
         FormatStyle(2)  =   "frmUpdate_OrdCompItem_Ex.frx":0138
         FormatStyle(3)  =   "frmUpdate_OrdCompItem_Ex.frx":01E8
         FormatStyle(4)  =   "frmUpdate_OrdCompItem_Ex.frx":029C
         FormatStyle(5)  =   "frmUpdate_OrdCompItem_Ex.frx":0374
         FormatStyle(6)  =   "frmUpdate_OrdCompItem_Ex.frx":042C
         FormatStyle(7)  =   "frmUpdate_OrdCompItem_Ex.frx":050C
         ImageCount      =   0
         PrinterProperties=   "frmUpdate_OrdCompItem_Ex.frx":052C
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   3120
      TabIndex        =   44
      Top             =   6240
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmUpdate_OrdCompItem_Ex.frx":0704
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmUpdate_OrdCompItem_Ex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SCod_Cliente_Tex As String
Public sser_ordcomp As String
Public scod_ordcomp As String
Public CODIGO, Descripcion As String, TipoAdd As String
Public sIGV_Cabecera As Double
Dim strSQL As String
Public sTipo As String
Dim Reg As ADODB.Recordset

Public Sub CARGA_GRID()
On Error GoTo err_Carga
    strSQL = "exec TI_SEL_ORDCOMPITEMDET_TINTO_VTAEXP '" & SCod_Cliente_Tex & "','" & sser_ordcomp & "','" & scod_ordcomp & "'"
    Set gexDetalle.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    ConfigurarGrid
Exit Sub
err_Carga:
    ErrorHandler err, "CARGA_GRID"
End Sub

Sub ConfigurarGrid()
    gexDetalle.Columns("cod_tela").Visible = False
    gexDetalle.Columns("cod_comb").Visible = False
    gexDetalle.Columns("cod_color").Visible = False
    gexDetalle.Columns("cod_opcion").Visible = False
    gexDetalle.Columns("Pre_Unitario").Visible = False

    gexDetalle.Columns("Sec_ordcomp").Caption = "Sec"
    gexDetalle.Columns("Cod_Tela_Cliente").Caption = "Cod.Tela Cliente"
    gexDetalle.Columns("cod_talla").Caption = "Talla"
    gexDetalle.Columns("porc_igv").Caption = "% IGV"
    
End Sub

Public Sub DeshabilitaCampos()
    Txtcod_Tela.Enabled = False
    TxtDes_Tela.Enabled = False
    txtCod_Color.Enabled = False
    txtDes_Color.Enabled = False
    txtCod_Comb.Enabled = False
    txtDes_Comb.Enabled = False
    txtCod_DsctoDet.Enabled = False
    txtDes_DstoDet.Enabled = False
    txtCod_Receta.Enabled = False
    txtDes_Receta.Enabled = False
    txtCod_Talla.Enabled = False
    txtCod_TelaCliente.Enabled = False
    txt_IGVDet.Enabled = False
    txtCant_Pedida.Enabled = False
    txtObs_Det.Enabled = False
    txtPrecio.Enabled = False
    dtpEntregaF_Det.Enabled = False
    dtpEntregaI_Det.Enabled = False
End Sub

Public Sub HabilitaCampos()
    txtCod_DsctoDet.Enabled = True
    txtDes_DstoDet.Enabled = True
    txtCod_Receta.Enabled = True
    txtDes_Receta.Enabled = True
    txtCod_Color.Enabled = True
    txtDes_Color.Enabled = True
    txtCod_TelaCliente.Enabled = True
    txt_IGVDet.Enabled = True
    txtCant_Pedida.Enabled = True
    txtObs_Det.Enabled = True
    txtPrecio.Enabled = True
    dtpEntregaF_Det.Enabled = True
    dtpEntregaI_Det.Enabled = True
End Sub

Sub LimpiaData()
    lbl_CanpedidaSinMod.Caption = ""
    Txtcod_Tela.Text = ""
    TxtDes_Tela.Text = ""
    txtCod_Color.Text = ""
    txtDes_Color.Text = ""
    txtCod_Comb.Text = ""
    txtDes_Comb.Text = ""
    strSQL = "SELECT cod_descuento FROM Tx_OrdComp where cod_cliente_tex='" & SCod_Cliente_Tex & "' and ser_ordcomp='" & sser_ordcomp & "' and cod_ordcomp='" & scod_ordcomp & "'"
    txtCod_DsctoDet.Text = DevuelveCampo(strSQL, cConnect)
    BuscaDescuento 1
    txtCod_Receta.Text = ""
    txtDes_Receta.Text = ""
    txtCod_Talla.Text = ""
    txtCod_TelaCliente.Text = ""
    strSQL = "SELECT porc_igv FROM Tx_OrdComp where cod_cliente_tex='" & SCod_Cliente_Tex & "' and ser_ordcomp='" & sser_ordcomp & "' and cod_ordcomp='" & scod_ordcomp & "'"
    txt_IGVDet.Text = DevuelveCampo(strSQL, cConnect)
    txtCant_Pedida = 0
    txtObs_Det.Text = ""
    txtPrecio = 0
    dtpEntregaF_Det.Value = Date
    dtpEntregaI_Det.Value = Date
End Sub

Private Sub CmdAnadir2_Click()
Dim dCantidadxTela As Double, dCant_Pedida As Double

    dCantidadxTela = IIf(Trim(Txt_CantidadXTela.Text) = "", 0, Txt_CantidadXTela.Text)
    dCant_Pedida = IIf(Trim(txtCant_Pedida.Text) = "", 0, txtCant_Pedida.Text)

    If dCantidadxTela < dCant_Pedida Then
        MsgBox "La cantidad a Ingresar sobrepasa el total de la orden establecida para esta tela", vbCritical, Me.Caption
        txtCant_Pedida.SetFocus
        Exit Sub
    Else
            If VALIDA_DATOS Then
                HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                DeshabilitaCampos
                SALVAR_DATOS
                CARGA_GRID
                sTipo = ""
                FrmCantidadxTela.Left = 12720
                Txt_CantidadXTela.Text = ""
            End If
    End If
End Sub

Private Sub cmdCancelar_Click()
FrmCantidadxTela.Left = 12720
Txt_CantidadXTela.Text = ""
End Sub

Private Sub Form_Load()
FrmCantidadxTela.Left = 12720
Txt_CantidadXTela.Text = ""
End Sub

Private Sub gexDetalle_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    txtNumSec.Text = gexDetalle.Value(gexDetalle.Columns("sec_ordcomp").Index)
    Txtcod_Tela.Text = Trim(gexDetalle.Value(gexDetalle.Columns("cod_tela").Index))
    TxtDes_Tela.Text = Trim(gexDetalle.Value(gexDetalle.Columns("tela").Index))
    txtCod_Color.Text = Trim(gexDetalle.Value(gexDetalle.Columns("cod_color").Index))
    txtDes_Color.Text = Trim(gexDetalle.Value(gexDetalle.Columns("color").Index))
    txtCod_Comb.Text = Trim(gexDetalle.Value(gexDetalle.Columns("cod_comb").Index))
    txtDes_Comb.Text = Trim(gexDetalle.Value(gexDetalle.Columns("combinacion").Index))
    txtCod_DsctoDet.Text = gexDetalle.Value(gexDetalle.Columns("cod_descuento").Index)
    txtDes_DstoDet.Text = Trim(gexDetalle.Value(gexDetalle.Columns("descuento").Index))
    txtCod_Receta.Text = gexDetalle.Value(gexDetalle.Columns("cod_opcion").Index)
    txtDes_Receta.Text = Trim(gexDetalle.Value(gexDetalle.Columns("receta").Index))
    txtCod_Talla.Text = Trim(gexDetalle.Value(gexDetalle.Columns("cod_talla").Index))
    txtCod_TelaCliente.Text = Trim(gexDetalle.Value(gexDetalle.Columns("cod_tela_cliente").Index))
    txt_IGVDet.Text = CDbl(gexDetalle.Value(gexDetalle.Columns("porc_igv").Index))
    txtCant_Pedida.Text = CDbl(gexDetalle.Value(gexDetalle.Columns("can_pedida").Index))
    lbl_CanpedidaSinMod.Caption = CDbl(gexDetalle.Value(gexDetalle.Columns("can_pedida").Index))
    
    If sTipo = "U" Then
        lbl_CanpedidaSinMod.Caption = CDbl(gexDetalle.Value(gexDetalle.Columns("can_pedida").Index))
    Else
        lbl_CanpedidaSinMod.Caption = 0
    End If
    
    txtPrecio.Text = CDbl(gexDetalle.Value(gexDetalle.Columns("pre_unitario").Index))
    txtObs_Det.Text = Trim(gexDetalle.Value(gexDetalle.Columns("Observaciones").Index))
    If Trim(CStr(gexDetalle.Value(gexDetalle.Columns("fec_entrega_inicio").Index))) <> "" Then dtpEntregaI_Det.Value = gexDetalle.Value(gexDetalle.Columns("fec_entrega_inicio").Index)
    If Trim(CStr(gexDetalle.Value(gexDetalle.Columns("fec_entrega_fin").Index))) <> "" Then dtpEntregaF_Det.Value = gexDetalle.Value(gexDetalle.Columns("fec_entrega_fin").Index)
End Sub


'Private Function BuscarSaldos() As Boolean
'Dim i As Integer
'Dim sRows As Integer, dPedido As Double, dTotalxTela As Double, dSumaCantidadTela As Double
'On Error GoTo hand
'
'dPedido = 0
'dTotalxTela = 0
'
'BuscarSaldos = False
'
'If Not Reg.EOF Then
'    sRows = Reg.RecordCount
'    Reg.MoveFirst
'
'
'        With Reg
'            For i = 1 To sRows
'
'                If txtCod_Tela = Trim(!Cod_Tela) Then
'
'                    dPedido = dPedido + Trim(!Pedida)
'                    dTotalxTela = dTotalxTela + Trim(!TotalxTela)
'
'
'                End If
'
'                Reg.MoveNext
'
'            Next
'        End With
'
'Reg.MoveFirst
'
'
'End If
'
'dSumaCantidadTela = dPedido + CDbl(IIf(txtCant_Pedida.Text = "", 0, txtCant_Pedida.Text))
'
'If dTotalxTela > 0 Then
'
'    If dTotalxTela < dSumaCantidadTela Then
'        BuscarSaldos = True
'    End If
'
'Else
'    BuscarSaldos = False
'End If
'
'
'
'Exit Function
'hand:
'    ErrorHandler err, "SALVAR_CABECERA"
'    Set gexDetalle.ADORecordset = Reg
'
'End Function
Private Function Buscar_Existe() As Boolean
    On Error GoTo SALTO_ERROR

    Dim oRs As New Recordset
    Dim dTotalxTela As Double, dCan_Pedida As Double, dCan_PedidaModif As Double

    strSQL = "EXEC Usp_Busca_Saldo '" & SCod_Cliente_Tex & "','" & sser_ordcomp & "','" & scod_ordcomp & "','" & Trim(Txtcod_Tela.Text) & "'"
    Set oRs = CargarRecordSetDesconectado(strSQL, cConnect)
    
    Buscar_Existe = False
    
    If oRs.RecordCount > 0 Then
    
        txtCant_Pedida.Text = IIf(Trim(txtCant_Pedida.Text) = "", 0, txtCant_Pedida.Text)
        lbl_CanpedidaSinMod.Caption = IIf(Trim(lbl_CanpedidaSinMod.Caption) = "", 0, lbl_CanpedidaSinMod.Caption)

        dTotalxTela = oRs.Fields("TotalxTela")
        dCan_Pedida = oRs.Fields("Can_Pedida")
        dCan_PedidaModif = dCan_Pedida - CDbl(lbl_CanpedidaSinMod.Caption) + CDbl(txtCant_Pedida.Text)
        
        If dTotalxTela = 0 Then
            Buscar_Existe = False
        Else
            
            Buscar_Existe = True
            
            
        End If

    End If

    Exit Function
SALTO_ERROR:
    MsgBox err.Description, vbCritical, Me.Caption
End Function
Private Function Buscar_Saldos() As Boolean
    On Error GoTo SALTO_ERROR

    Dim oRs As New Recordset
    Dim dTotalxTela As Double, dCan_Pedida As Double, dCan_PedidaModif As Double

    strSQL = "EXEC Usp_Busca_Saldo '" & SCod_Cliente_Tex & "','" & sser_ordcomp & "','" & scod_ordcomp & "','" & Trim(Txtcod_Tela.Text) & "'"
    Set oRs = CargarRecordSetDesconectado(strSQL, cConnect)
    
    Buscar_Saldos = False
    
    If oRs.RecordCount > 0 Then
    
        txtCant_Pedida.Text = IIf(Trim(txtCant_Pedida.Text) = "", 0, txtCant_Pedida.Text)
        lbl_CanpedidaSinMod.Caption = IIf(Trim(lbl_CanpedidaSinMod.Caption) = "", 0, lbl_CanpedidaSinMod.Caption)

        dTotalxTela = oRs.Fields("TotalxTela")
        dCan_Pedida = oRs.Fields("Can_Pedida")
        dCan_PedidaModif = dCan_Pedida - CDbl(lbl_CanpedidaSinMod.Caption) + CDbl(txtCant_Pedida.Text)
        
        If dTotalxTela = 0 Then
            Buscar_Saldos = False
        Else
            
            If dCan_PedidaModif > dTotalxTela Then
            
                Buscar_Saldos = True
            
            End If
            
        End If

    End If

    Exit Function
SALTO_ERROR:
    MsgBox err.Description, vbCritical, Me.Caption

End Function
Public Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ADICIONAR"
            sTipo = "I"
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            LimpiaData
            
            Txtcod_Tela.Enabled = True
            TxtDes_Tela.Enabled = True
            txtCod_Color.Enabled = True
            txtDes_Color.Enabled = True
            txtCod_Comb.Enabled = True
            txtDes_Comb.Enabled = True
            txtCod_Talla.Enabled = True
            HabilitaCampos
            gexDetalle.Enabled = False
        Case "MODIFICAR"
            If gexDetalle.RowCount = 0 Then Exit Sub
            sTipo = "U"
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            HabilitaCampos
            gexDetalle.Enabled = False
        Case "ELIMINAR"
            Dim vMessage As Variant
            If gexDetalle.RowCount = 0 Then Exit Sub
            vMessage = MsgBox("Esta seguro que desea el registro selecionado", vbYesNo, "Eliminar")
            If vMessage = vbYes Then
                sTipo = "D"
                SALVAR_DATOS
            End If
            CARGA_GRID
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Case "GRABAR"
            If VALIDA_DATOS Then
                HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                DeshabilitaCampos
                SALVAR_DATOS
                CARGA_GRID
                sTipo = ""
                'Call gexFamGru.Find(gexFamGru.Columns("codigo").Index, jgexEqual, vCodigo)
                FrmCantidadxTela.Left = 12720
                Txt_CantidadXTela.Text = ""
            End If
            gexDetalle.Enabled = True
        Case "DESHACER"
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            DeshabilitaCampos
            CARGA_GRID
            gexDetalle.Enabled = True
        Case "SALIR"
            Unload Me
    End Select
End Sub
Function ValidaCargaItem() As Boolean
Dim dCan_Pedida As Double, dCan_Total As Double, strSQL As String, dCan_Item As Double, Strsql2 As String

'SCod_Cliente_Tex = DevuelveCampo("select cod_cliente_tex from tx_cliente where abr_cliente='" & txtAbr_Cliente.Text & "'", cConnect)
If sTipo = "I" Then
    strSQL = "Select Isnull(Sum(Can_Pedida),0) From tx_ordcompitem_tinto where cod_cliente_tex='" & SCod_Cliente_Tex & "' and ser_ordcomp='" & sser_ordcomp & "' and cod_ordcomp='" & scod_ordcomp & "'"
Else
    strSQL = "Select Isnull(Sum(Can_Pedida),0) From tx_ordcompitem_tinto where cod_cliente_tex='" & SCod_Cliente_Tex & "' and ser_ordcomp='" & sser_ordcomp & "' and cod_ordcomp='" & scod_ordcomp & "' and Sec_OrdComp<>'" & Trim(txtNumSec.Text) & "'"
End If

Strsql2 = "Select Isnull(CantidadTotal,0) From tx_ordcomp where cod_cliente_tex='" & SCod_Cliente_Tex & "' and ser_ordcomp='" & sser_ordcomp & "' and cod_ordcomp='" & scod_ordcomp & "' "
'select CantidadTotal  from tx_ordcomp where

dCan_Pedida = DevuelveCampo(strSQL, cConnect)
dCan_Total = DevuelveCampo(Strsql2, cConnect)


dCan_Item = dCan_Pedida + CDbl(txtCant_Pedida.Text)

ValidaCargaItem = False

If dCan_Total < dCan_Item Then
    ValidaCargaItem = True
End If

End Function
Sub SALVAR_DATOS()
Dim i As Integer
On Error GoTo hand

Txt_CantidadXTela.Text = IIf(Trim(Txt_CantidadXTela.Text) = "", 0, Txt_CantidadXTela.Text)
            
            strSQL = "EXEC TI_MAN_TX_ORDCOMPITEM_TINTO_EX '" & sTipo & "','" & _
            SCod_Cliente_Tex & "','" & _
            sser_ordcomp & "','" & _
            scod_ordcomp & "','" & _
            txtNumSec.Text & "','" & _
            Trim(Txtcod_Tela.Text) & "','" & _
            Trim(txtCod_Comb.Text) & "','" & _
            Trim(txtCod_Color.Text) & "','" & _
            Trim(txtCod_Receta.Text) & "','" & _
            Trim(txtCod_Talla.Text) & "','" & _
            Trim(txtCod_DsctoDet.Text) & "'," & _
            txt_IGVDet.Text & ",'" & _
            dtpEntregaI_Det.Value & "','" & _
            dtpEntregaF_Det.Value & "'," & _
            txtPrecio.Text & "," & _
            txtCant_Pedida.Text & ",'" & _
            txtCod_TelaCliente.Text & "','" & _
            txtObs_Det.Text & "','" & Txt_Unidades & "',''," & Txt_CantidadXTela.Text & ""

            Call ExecuteSQL(cConnect, strSQL)
    
Exit Sub
hand:
    ErrorHandler err, "SALVAR_CABECERA"
End Sub
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
            
                If Txtcod_Tela = Trim(!Cod_Tela) Then
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
    ErrorHandler err, "SALVAR_CABECERA"
    Set gexDetalle.ADORecordset = Reg

End Function

Function VALIDA_DATOS() As Boolean

    VALIDA_DATOS = True
    If Trim(Txtcod_Tela.Text) = "" Then
        MsgBox "Seleccione la Tela", vbCritical, Me.Caption
        Txtcod_Tela.SetFocus
        VALIDA_DATOS = False
        Exit Function
    End If

    If Trim(txtCod_Color.Text) = "" Then
        MsgBox "Seleccione el Color", vbCritical, Me.Caption
        txtCod_Color.SetFocus
        VALIDA_DATOS = False
        Exit Function
    End If

    If Trim(txtCod_DsctoDet.Text) = "" Then
        MsgBox "Seleccione el Descuento", vbCritical, Me.Caption
        txtCod_DsctoDet.SetFocus
        VALIDA_DATOS = False
        Exit Function
    End If

    If Trim(txtCant_Pedida.Text) = "" Or CDbl(txtCant_Pedida.Text) <= 0 Then
        MsgBox "Ingrese una Cantidad valida", vbCritical, Me.Caption
        txtCant_Pedida.SetFocus
        VALIDA_DATOS = False
        Exit Function
    End If
    
'    If ValidaCargaItem = True Then
'        MsgBox "La cantidad a Ingresar sobrepasa el total de la Orden de Servicio", vbCritical, Me.Caption
'        txtCant_Pedida.SetFocus
'        VALIDA_DATOS = False
'        Exit Function
'
'    End If

    If Buscar_Existe = False And Trim(Txt_CantidadXTela.Text) = "" Then
        MsgBox "Debe Ingresar la Cantidad Total Por Tela", vbInformation, "Informacion"
        FrmCantidadxTela.Left = 4440
        Txt_CantidadXTela.SetFocus
        VALIDA_DATOS = False
        Exit Function
    End If
    
   If Buscar_Saldos = True Then
        MsgBox "La cantidad a Ingresar sobrepasa el total de la orden establecida para esta Tela", vbCritical, Me.Caption
        txtCant_Pedida.SetFocus
        VALIDA_DATOS = False
        Exit Function
    End If
    
End Function

Private Sub txt_IGVDet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        Call SoloNumeros(txt_IGVDet, KeyAscii, True, 2)
    End If
End Sub

Private Sub Txt_Unidades_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        Call SoloNumeros(Txt_Unidades, KeyAscii, True, 2)
    End If
End Sub

Private Sub txtCant_Pedida_GotFocus()
    SelectionText txtCant_Pedida
End Sub

Private Sub txtCant_Pedida_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        Call SoloNumeros(txtCant_Pedida, KeyAscii, True, 2)
    End If
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

Private Sub txtCod_Talla_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtObs_Det_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtPrecio_GotFocus()
    SelectionText txtPrecio
End Sub

Private Sub TxtPrecio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        Call SoloNumeros(txtPrecio, KeyAscii, True, 5)
    End If
End Sub

Private Sub txtcod_tela_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim oTipo As New frmBuscaTela
    
        Set oTipo.oParent = Me
        
        If Len(Trim(Txtcod_Tela)) > 2 Then
            Dim Temp As String
            Temp = Trim(DevuelveCampo("Select dbo.uf_devuelvecodigo(6," & IIf(Trim(Txtcod_Tela.Text) = "", 0, Mid(Txtcod_Tela.Text, 3)) & ")", cConnect))
             Txtcod_Tela = Left(Txtcod_Tela.Text, 2) & Temp
        End If

       
        oTipo.scod_Cliente = Me.SCod_Cliente_Tex
        oTipo.sCod_Tela = Trim(Txtcod_Tela.Text)
        If Trim(Txtcod_Tela.Text) = "" Then oTipo.ChkAllClient.Visible = False
        oTipo.Campo = 1
        oTipo.Cargar_Datos
        'oTipo.DGridLista.Columns(2).Width = 3500
        oTipo.Show 1
        If CODIGO <> "" Then
             Me.Txtcod_Tela.Text = Trim(CODIGO)
             Me.TxtDes_Tela.Text = Trim(Descripcion)
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

Private Sub txtCod_Comb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Txtcod_Tela.Text) = "" Then
            MsgBox "El Codigo de Tela no puede estar vacio, Verifique", vbCritical, Me.Caption
            Txtcod_Tela.SetFocus
            Exit Sub
        End If
        
        If Trim(txtCod_Comb.Text) = "" Then
            BUSCA_COMBO 3
        Else
            BUSCA_COMBO 1
        End If
    End If
End Sub

Private Sub txtDes_Comb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Txtcod_Tela.Text) = "" Then
            MsgBox "El Codigo de Tela no puede estar vacio, Verifique", vbCritical, Me.Caption
            Txtcod_Tela.SetFocus
            Exit Sub
        End If
        
        If Trim(txtDes_Comb.Text) = "" Then
            BUSCA_COMBO 3
        Else
            BUSCA_COMBO 2
        End If
    End If
End Sub

Private Sub txtCod_DsctoDet_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then BuscaDescuento 1
End Sub

Private Sub txtDes_DstoDet_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then BuscaDescuento 2
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

Public Sub BUSCA_COMBO(tipo As Integer)
    Select Case tipo
        Case 1:
                    strSQL = "Select Des_Comb From Tx_TelaComb Where cod_tela='" & Txtcod_Tela.Text & "' and cod_comb='" & Trim(txtCod_Comb.Text) & "'"
                    Me.txtDes_Comb.Text = Trim(DevuelveCampo(strSQL, cConnect))
                    If Trim(txtDes_Comb.Text) <> "" Then
                        SendKeys "{TAB}"
                        SendKeys "{TAB}"
                    End If
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If tipo = 2 Then
                        oTipo.SQuery = "Select Cod_Comb as Codigo, Des_Comb as Descripcion From Tx_TelaComb Where cod_tela='" & Txtcod_Tela.Text & "' and des_comb like '%" & Trim(txtDes_Comb.Text) & "%' order by cod_comb"
                    Else
                        oTipo.SQuery = "Select Cod_Comb as Codigo, Des_Comb as Descripcion From Tx_TelaComb Where cod_tela='" & Txtcod_Tela.Text & "' order by cod_comb"
                    End If
                    
                    oTipo.Cargar_Datos
                    oTipo.gexList.Columns(2).Width = 3500
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
                    strSQL = "Select Des_Color From Lb_Color Where cod_color='" & txtCod_Color.Text & "'"
                    Me.txtDes_Color.Text = Trim(DevuelveCampo(strSQL, cConnect))
                    If Trim(txtDes_Color.Text) <> "" Then
                        SendKeys "{TAB}"
                        SendKeys "{TAB}"
                    End If
        Case 2, 3:
                    Dim oTipo As New frmBusGeneral6
                    Dim rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If tipo = 2 Then
                        oTipo.SQuery = "Select Cod_Color as Codigo, Des_Color as Descripcion From lb_Color Where des_color like '%" & Trim(txtDes_Color.Text) & "%' order by des_color"
                    Else
                        oTipo.SQuery = "Select Cod_Color as Codigo, Des_Color as Descripcion From Lb_Color order by des_color"
                    End If
                    
                    oTipo.Cargar_Datos
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
                    Dim rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If tipo = 2 Then
                        oTipo.SQuery = "Select Cod_opcion as Codigo, Descripcion as Descripcion From TI_Recetas_Tintoreria Where cod_color='" & txtCod_Color.Text & "' and descripcion like '%" & Trim(txtDes_Receta.Text) & "%' order by descripcion"
                    Else
                        oTipo.SQuery = "Select Cod_opcion as Codigo, Descripcion as Descripcion From TI_Recetas_Tintoreria where cod_color='" & txtCod_Color.Text & "' order by descripcion"
                    End If
                    
                    oTipo.Cargar_Datos
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

Private Sub BuscaDescuento(Opcion As Integer)
Dim rstAux As New ADODB.Recordset
On Error GoTo Fin
    strSQL = "SELECT Cod_Descuento, Des_Descuento, Porcentaje1 " & _
             "FROM LG_DSCTOS WHERE "
    Select Case Opcion
    Case 1: strSQL = strSQL & "Cod_Descuento like '%" & txtCod_DsctoDet & "%'"
    Case 2: strSQL = strSQL & "Des_Descuento like '%" & txtDes_DstoDet & "%'"
    End Select
    txtCod_DsctoDet = ""
    txtDes_DstoDet = ""
    txtDes_DstoDet.Tag = ""
    With frmBusGeneral6
        Set .oParent = Me
        .SQuery = strSQL
        .Cargar_Datos
        
        .DGridLista.Columns("Cod_Descuento").Caption = "Codigo"
        .DGridLista.Columns("Cod_Descuento").Width = 700
        .DGridLista.Columns("Des_Descuento").Caption = "Descuento"
        .DGridLista.Columns("Des_Descuento").Width = 5000
        .DGridLista.Columns("Porcentaje1").Visible = False
        
        CODIGO = ".."
        Set rstAux = .DGridLista.ADORecordset
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If CODIGO <> "" And rstAux.RecordCount > 0 Then
            txtCod_DsctoDet = Trim(rstAux!cod_descuento)
            txtDes_DstoDet = Trim(rstAux!Des_Descuento)
            txtDes_DstoDet.Tag = rstAux!Porcentaje1
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
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Busqueda de Descuento (" & Opcion & ")"
End Sub

