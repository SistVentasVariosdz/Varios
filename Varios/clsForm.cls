VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private mConnect As String
Private Declare Function PutFocus Lib "user32" _
   Alias "SetFocus" _
  (ByVal hwnd As Long) As Long
Public Property Get Parent() As Object
    Set Parent = oMDIParent
End Property
Public Property Let Parent(ByVal vnewvalue As Object)
    Set oMDIParent = vnewvalue
End Property
Public Property Get Language() As Integer
    Language = iLanguage
End Property
Public Property Let Language(ByVal vnewvalue As Integer)
    iLanguage = vnewvalue
End Property
Public Property Let Rutas(ByVal mRutas As Variant)
    vRuta = mRutas
End Property
Public Property Get Rutas()
    Rutas = vRuta
End Property
Public Property Let Cod_Empresa(ByVal mCod_Empresa As String)
    vemp = mCod_Empresa
    bCargaConexion = True
End Property
Public Property Let Cod_Perfil(ByVal mCod_Perfil As String)
    vper = mCod_Perfil
End Property
Public Property Let Cod_Opcion(ByVal mCod_Opcion As String)
    'gsCod_Opcion = mCod_Opcion
End Property
Public Property Let Key_Node(ByVal mKey_Node As String)
    'gsKey_Node = mKey_Node
End Property
Public Property Let Permisos(ByVal mPermisos As String)
    'gsPermisos = mPermisos
End Property
Public Property Let Nom_Empresa(ByVal vnewvalue As String)
    'gsNomEmpresa = vnewvalue
End Property
Public Property Let Nom_Aplicacion(ByVal vnewvalue As String)
    'gsNomAplicacion = vnewvalue
End Property
Public Property Let Nom_Opcion(ByVal vnewvalue As String)
    vNomFor = vnewvalue
End Property
Public Property Let UserName(ByVal vnewvalue As String)
    vusu = vnewvalue
End Property
Public Property Get UserName() As String
    UserName = vusu
End Property
Public Property Let ConnectEmpresa(ByVal vnewvalue As String)
    mConnect = vnewvalue
    cCONNECT = mConnect
End Property
Public Property Let ConnectSeguridad(ByVal vnewvalue As String)
    cSEGURIDAD = vnewvalue
End Property
Public Function GetForm(sFormName) As Object
    sFormName = Trim(UCase(sFormName))
    Select Case sFormName
        Case "FRMMUESTRACTACTECLIENTESRANGOS"
            Set GetForm = New FrmMuestraCtaCteClientesRangos
        Case "FRMBARRATRABAJADOR"
            Set GetForm = New FrmBarraTrabajador
        Case "FRMCAMBIOCLAVE"
            Set GetForm = New FrmCambioClave
         Case "FRMPRODUCCIONTERMOFIJADO"
            Set GetForm = New FrmProduccionTermofijado
         Case "FRMMUESTRATIEMPOPARTIDAPLANTA"
            Set GetForm = New FrmMuestraTiempoPartidaPlanta
         Case "FRMESTADOLABORATORIOITEMS"
            Set GetForm = New FrmEstadoLaboratorioItems
         Case "FRMSHOWFACTVENTAS"
           Set GetForm = New frmShowFactVentas
         Case "FRMGUIAREMISION"
           Set GetForm = New FrmGuiaRemision
         Case "FRMCONSULTAMOVTELATENIDA"
           Set GetForm = New FrmConsultaMovTelaTenida
         Case "FRMSTOCKSTENIDO"
           Set GetForm = New frmStocksTenido
         Case "FRMCORRIGENUMERACIONDOCUMENTO"
           Set GetForm = New frmCorrigeNumeracionDocumento
         Case "FRMCALCULAMERMA"
           Set GetForm = New FrmCalculaMerma
         Case "FRMIMPRIMEETIQUETASPRENDAS"
           Set GetForm = New FrmImprimeEtiquetasPrendas
         Case "FRMADICIONADOCUMVENTASPRENDAS"
           Set GetForm = New frmAdicionaDocumVentasPrendas
         Case "FRMUSUARIOCORRELATIVOS"
           Set GetForm = New FrmUsuarioCorrelativos
         Case "FRMCAMBIARULTIMOCORRELATIVO"
           Set GetForm = New frmCambiarUltimoCorrelativo
         Case "FRM_MANTENIMIENTO_SERIES_POR_ALMACEN"
           Set GetForm = New Frm_mantenimiento_series_Por_Almacen
         Case "FRMGUIASREMISIONPRENDAS"
           Set GetForm = New FrmGuiasRemisionPrendas
         Case "FRMSHOWFACTVENTASPRENDAS"
           Set GetForm = New frmShowFactVentasPrendas
         Case "FRMMATTG_CLIENTE"
           Set GetForm = New frmMatTg_Cliente
         Case "FRMAPERTURACAJA"
           Set GetForm = New frmAperturaCaja
         Case "FRMCIERRACAJAS"
           Set GetForm = New FrmCierraCajas
         Case "FRMMUESTRAVENTADIARIO"
           Set GetForm = New FrmMuestraVentaDiario
         Case "FRMCAMBIOMODELOTALLA"
            Set GetForm = New FrmCambioModeloTalla
         Case "FRMSHOWTX_ORDCOMP_EX"
            Set GetForm = New frmShowTX_OrdComp_Ex
         Case "FRMFACTURAPROFORMA"
            Set GetForm = New FrmFacturaProforma
         Case "FRM_LISTAPACKINGLIST"
             Set GetForm = New Frm_ListaPackingList
         Case "FRMDESPACHOPACKINGEX"
             Set GetForm = New FrmDespachoPackingEx
         Case "FRMDATOSPARTIDAS"
              Set GetForm = New FrmDatosPartidas
         Case "FRMMANTELAS"
              Set GetForm = New frmManTelas
         Case "FRMMANTFAMTELA"
              Set GetForm = New frmMantFamTela
         Case "FRM_MANTENIMIENTO_OPERARIO_PROCESO"
              Set GetForm = New Frm_Mantenimiento_Operario_Proceso
         Case "FRM_MANTENIMIENTO_SUBPROCESOS"
               Set GetForm = Frm_Mantenimiento_Subprocesos
         Case "FRMSHOWAUDITORIATEJEDURIA"
               Set GetForm = FrmShowAuditoriaTejeduria
         Case "FRMSHOWGUIASXFACT_PRENDAS"
               Set GetForm = frmShowGuiasxFact_Prendas
         Case "FRMCAPTURAMOVIMIENTOTEJEDURIA"
               Set GetForm = FrmCapturaMovimientoTejeduria
         Case "FRMRPTCOLORESCREADOS"
               Set GetForm = FrmRptColoresCreados
         Case "FRMINSPECCIONORDENESCOMPRA"
               Set GetForm = frmInspeccionOrdenesCompra
         Case "FRMSOLICITUDDESACOLORESLOCAL"
               Set GetForm = FrmSolicitudDesaColoresLocal
    End Select
End Function


