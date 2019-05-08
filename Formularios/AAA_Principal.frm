VERSION 5.00
Begin VB.Form AAA_Principal 
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command25 
      Caption         =   "DesarrolloColores"
      Height          =   495
      Left            =   7320
      TabIndex        =   46
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton CmdInspeccionOrdenes 
      Caption         =   "Inspeccion Ordenes"
      Height          =   495
      Left            =   5520
      TabIndex        =   45
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton cmdCrecionColore 
      Caption         =   "Reporte de Creacion de Colores"
      Height          =   615
      Left            =   5520
      TabIndex        =   44
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Recepcion de Tela Cruda"
      Height          =   495
      Left            =   5520
      TabIndex        =   43
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cmdAutorizacionVentas 
      Caption         =   "Autorizacion Guias Prendas Facontex"
      Height          =   495
      Left            =   5520
      TabIndex        =   42
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdAuditoriaRollos 
      Caption         =   "Auditoria de Rollos"
      Height          =   615
      Left            =   5520
      TabIndex        =   41
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Mantenimiento SubProcesos"
      Height          =   495
      Left            =   5520
      TabIndex        =   40
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Mantenimiento Operario Proceso"
      Height          =   495
      Left            =   5520
      TabIndex        =   39
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Mant Fam Telas"
      Height          =   495
      Left            =   5520
      TabIndex        =   38
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Telas Facontex"
      Height          =   495
      Left            =   5520
      TabIndex        =   37
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Muestra Stocks Rango Fecha"
      Height          =   495
      Left            =   5520
      TabIndex        =   36
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdQycPenDes 
      Caption         =   "Quimicos Pendientes Despachos"
      Height          =   495
      Left            =   5520
      TabIndex        =   35
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdDocExportacion 
      Caption         =   "Documentos de Exportacion"
      Height          =   495
      Left            =   3720
      TabIndex        =   34
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdPReFactura 
      Caption         =   "Reporte de Prefactura"
      Height          =   495
      Left            =   3720
      TabIndex        =   33
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton cmdFacturarGuia 
      Caption         =   "Autorizar Guias"
      Height          =   495
      Left            =   3720
      TabIndex        =   32
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmDespachoPacking 
      Caption         =   "Despacho Packing"
      Height          =   495
      Left            =   3720
      TabIndex        =   31
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton Command18 
      Caption         =   "PACKING LIST  EXP TELAS"
      Height          =   495
      Left            =   3720
      TabIndex        =   30
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Proforma Factura Exportacion"
      Height          =   495
      Left            =   3720
      TabIndex        =   29
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Orden Servicio Exportacion"
      Height          =   495
      Left            =   3720
      TabIndex        =   28
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Cambio Modelo Talla"
      Height          =   495
      Left            =   3720
      TabIndex        =   27
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Ventas Diaria"
      Height          =   615
      Left            =   3720
      TabIndex        =   26
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdCierreCaja 
      Caption         =   "Cierre Caja"
      Height          =   495
      Left            =   3720
      TabIndex        =   25
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdAperturaCaja 
      Caption         =   "Apertura de Caja"
      Height          =   495
      Left            =   3720
      TabIndex        =   24
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Mant Tg Cliente"
      Height          =   495
      Left            =   3720
      TabIndex        =   23
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Asiga Serie Almacen"
      Height          =   615
      Left            =   1920
      TabIndex        =   22
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton Command11 
      Caption         =   "ModifiCar Correlativo Guia"
      Height          =   615
      Left            =   1920
      TabIndex        =   21
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Acceso Guia"
      Height          =   615
      Left            =   1920
      TabIndex        =   20
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Guias Remision Prendas"
      Height          =   615
      Left            =   1920
      TabIndex        =   19
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cmdVentasPrendas 
      Caption         =   "Ventas Prendas"
      Height          =   495
      Left            =   1920
      TabIndex        =   18
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton cmdVentaPrendas 
      Caption         =   "Punto Ventas Prendas"
      Height          =   495
      Left            =   1920
      TabIndex        =   16
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Etiqueta Prenda"
      Height          =   495
      Left            =   1920
      TabIndex        =   15
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   495
      Left            =   1920
      TabIndex        =   14
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   1920
      TabIndex        =   13
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdCorrigeDocumentos 
      Caption         =   "corrige Numeracion"
      Height          =   495
      Left            =   1920
      TabIndex        =   12
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdStockTelTen 
      Caption         =   "stock Tela acabada"
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdConsultaTelTel 
      Caption         =   "consulta mov tela acabada"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton CmdGuiasRemision 
      Caption         =   "Guias Remision Telas"
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Imprime Etiqueta Inventario"
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Muestra Aprobacion Quimicos"
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdPuntoVenta 
      Caption         =   "Punto Ventas Telas"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "FACTURA SS TEJIDO"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton CmdDiasplanta 
      Caption         =   "Dias Tela Planta"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdproduccionTermo 
      Caption         =   "Produccion Termofijado"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cambio Contraseña"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdBarras 
      Caption         =   "Barras"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdCtaCteRangos 
      Caption         =   "Cta Cte Rangos"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "AAA_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAperturaCaja_Click()
frmAperturaCaja.Show 1
End Sub

Private Sub cmdAuditoriaRollos_Click()
FrmShowAuditoriaTejeduria.Show 1
End Sub

Private Sub cmdAutorizacionVentas_Click()
frmShowGuiasxFact_Prendas.Show 1
End Sub

Private Sub cmdBarras_Click()
    FrmBarraTrabajador.Show 1
End Sub

Private Sub cmdCierreCaja_Click()
    FrmCierraCajas.Show 1
End Sub
Private Sub cmdConsultaTelTel_Click()
    FrmConsultaMovTelaTenida.Show 1
End Sub
Private Sub cmdCorrigeDocumentos_Click()
    frmCorrigeNumeracionDocumento.Show 1
End Sub

Private Sub cmdCrecionColore_Click()
FrmRptColoresCreados.Show 1
End Sub

Private Sub cmdCtaCteRangos_Click()
    FrmMuestraCtaCteClientesRangos.Show 1
End Sub
Private Sub CmdDiasplanta_Click()
    FrmMuestraTiempoPartidaPlanta.Show
End Sub

Private Sub cmdDocExportacion_Click()
'frmShowFactVentasEx.Show 1
End Sub

Private Sub cmDespachoPacking_Click()
FrmDespachoPackingEx.Show 1
End Sub

Private Sub cmdFacturarGuia_Click()
frmShowGuiasxFact_TelaTenida.Show 1
End Sub

Private Sub CmdGuiasRemision_Click()
    FrmGuiaRemision.Show 1
End Sub

Private Sub CmdInspeccionOrdenes_Click()
frmInspeccionOrdenesCompra.Show 1
End Sub

Private Sub cmdPReFactura_Click()
FrmRepPreFactura.Show 1
End Sub

Private Sub cmdproduccionTermo_Click()
    FrmProduccionTermofijado.Show 1
End Sub

Private Sub cmdPuntoVenta_Click()
    frmShowFactVentas.Show 1
End Sub

Private Sub cmdQycPenDes_Click()
FrmDatosPartidas.Show 1
End Sub

Private Sub cmdStockTelTen_Click()
    frmStocksTenido.Show 1
End Sub

Private Sub cmdVentaPrendas_Click()
frmAdicionaDocumVentasPrendas.Show
End Sub

Private Sub cmdVentasPrendas_Click()
frmShowFactVentasPrendas.Show 1
End Sub

Private Sub Command1_Click()
    FrmCambioClave.Show 1
End Sub

Private Sub Command10_Click()
FrmUsuarioCorrelativos.Show 1
End Sub

Private Sub Command11_Click()
frmCambiarUltimoCorrelativo.Show 1
End Sub

Private Sub Command12_Click()
Frm_mantenimiento_series_Por_Almacen.Show 1
End Sub

Private Sub Command13_Click()
frmMatTg_Cliente.Show 1
End Sub

Private Sub Command14_Click()
'Form2.Show 1
FrmMuestraVentaDiario.Show 1

End Sub

Private Sub Command15_Click()
FrmCambioModeloTalla.Show 1
End Sub

Private Sub Command16_Click()
frmShowTX_OrdComp_Ex.Show 1
End Sub

Private Sub Command17_Click()
FrmFacturaProforma.Show 1
End Sub

Private Sub Command18_Click()
Frm_ListaPackingList.Show 1
End Sub

Private Sub Command19_Click()
    FrmMuestraStocksFechas.Show 1
End Sub

Private Sub Command2_Click()
   'FrmDesbloqueTeclado.Show 1
   FrmFacturaGuiaTejido.Show 1
End Sub

Private Sub Command20_Click()
frmManTelas.Show 1
End Sub

Private Sub Command21_Click()
frmMantFamTela.Show 1
End Sub

Private Sub Command22_Click()
Frm_Mantenimiento_Operario_Proceso.Show 1
End Sub

Private Sub Command23_Click()
Frm_Mantenimiento_Subprocesos.Show 1
End Sub

Private Sub Command24_Click()
FrmCapturaMovimientoTejeduria.Show 1
End Sub

Private Sub Command25_Click()
FrmSolicitudDesaColoresLocal.Show 1
End Sub

Private Sub Command3_Click()
FrmEstadoLaboratorioItems.Show 1
End Sub
Private Sub Command4_Click()
FrmImprimeEtiquetaInventario.Show 1
End Sub
'

Private Sub Command6_Click()
FrmCalculaMerma.Show 1
End Sub

Private Sub Command7_Click()
FrmImprimeEtiquetasPrendas.Show 1
End Sub

Private Sub Command8_Click()
FrmAPrueba.Show
End Sub

Private Sub Command9_Click()
FrmGuiasRemisionPrendas.Show 1
End Sub

Private Sub Form_Load()

    cConnect = "Provider=SQLOLEDB.1;Persist Security Info=False;Initial Catalog=FERRETERIA;Data Source=pcleon;uid=soporte;pwd=soporte"
    cSEGURIDAD = "Provider=SQLOLEDB.1;Persist Security Info=False;Initial Catalog=SEGURIDAD;Data Source=vpcleon;uid=soporte;pwd=soporte"

    vper = "0001"
    vemp = "01"
    vemp1 = "01"
    vusu = "SISTEMAS"
    'vRuta = "C:\Program Files\Sistema Produccion"
    vRuta = "C:\Program Files (x86)\Sistema Produccion"
    'vRuta = "C:\Program Files\Gestion de Pedidos"
    
End Sub

Private Sub FrmMuestra_Click()

End Sub
