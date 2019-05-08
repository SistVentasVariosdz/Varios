VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_ListaPackingList 
   Caption         =   "Packing List Exportaciones"
   ClientHeight    =   7605
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   5775
      Left            =   10200
      TabIndex        =   12
      Top             =   1680
      Width           =   1575
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   3030
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   5345
         Custom          =   $"Frm_ListaPackingList.frx":0000
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1150
         ControlHeigth   =   500
         ControlSeparator=   0
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Height          =   1650
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   11625
      Begin VB.TextBox TxtPacking 
         Height          =   285
         Left            =   8280
         TabIndex        =   17
         Top             =   1275
         Width           =   2175
      End
      Begin VB.OptionButton OptPack 
         BackColor       =   &H0080FFFF&
         Caption         =   "Packing List"
         Height          =   255
         Left            =   6720
         TabIndex        =   16
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox TxtFacturaProforma 
         Height          =   285
         Left            =   8280
         TabIndex        =   14
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtAbr_Cliente 
         Height          =   315
         Left            =   825
         TabIndex        =   7
         Top             =   285
         Width           =   690
      End
      Begin VB.TextBox txtNom_Cliente 
         Height          =   315
         Left            =   1530
         TabIndex        =   6
         Top             =   285
         Width           =   4560
      End
      Begin VB.OptionButton OptOC 
         BackColor       =   &H0080FFFF&
         Caption         =   "Orden Compra"
         Height          =   255
         Left            =   6735
         TabIndex        =   8
         Top             =   180
         Width           =   1395
      End
      Begin VB.OptionButton OptRango 
         BackColor       =   &H0080FFFF&
         Caption         =   "Rango Fechas"
         Height          =   270
         Left            =   6735
         TabIndex        =   5
         Top             =   555
         Value           =   -1  'True
         Width           =   1440
      End
      Begin VB.TextBox txtSer_OrdComp 
         Height          =   285
         Left            =   8235
         TabIndex        =   10
         Top             =   135
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.TextBox txtCod_OrdComp 
         Height          =   285
         Left            =   8775
         TabIndex        =   3
         Top             =   135
         Visible         =   0   'False
         Width           =   1395
      End
      Begin MSComCtl2.DTPicker DTPInicio 
         Height          =   300
         Left            =   8250
         TabIndex        =   2
         Top             =   525
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         _Version        =   393216
         Format          =   70320129
         CurrentDate     =   37988
      End
      Begin FunctionsButtons.FunctButt FunctBuscar 
         Height          =   495
         Left            =   4920
         TabIndex        =   4
         Top             =   720
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~1~~0~Falso~Falso~&Buscar~"
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin MSComCtl2.DTPicker DTPFin 
         Height          =   300
         Left            =   9780
         TabIndex        =   9
         Top             =   525
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         _Version        =   393216
         Format          =   70320129
         CurrentDate     =   37988
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Factura Proforma"
         Height          =   195
         Left            =   6720
         TabIndex        =   13
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "A"
         Height          =   195
         Left            =   9570
         TabIndex        =   11
         Top             =   570
         Visible         =   0   'False
         Width           =   150
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   10186
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
      FormatStyle(1)  =   "Frm_ListaPackingList.frx":01ED
      FormatStyle(2)  =   "Frm_ListaPackingList.frx":0325
      FormatStyle(3)  =   "Frm_ListaPackingList.frx":03D5
      FormatStyle(4)  =   "Frm_ListaPackingList.frx":0489
      FormatStyle(5)  =   "Frm_ListaPackingList.frx":0561
      FormatStyle(6)  =   "Frm_ListaPackingList.frx":0619
      FormatStyle(7)  =   "Frm_ListaPackingList.frx":06F9
      ImageCount      =   0
      PrinterProperties=   "Frm_ListaPackingList.frx":0719
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   1200
      Top             =   7680
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "Frm_ListaPackingList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CODIGO As String
Public DESCRIPCION As String
Dim rsx As New ADODB.Recordset

Private Sub Form_Load()
OptOC_Click
OptOC.Value = True
End Sub

Private Sub FunctBuscar_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
'On Error GoTo xerror:
Select Case ActionName
Case "BUSCAR":
Call Busca_PackingList
End Select
Exit Sub
'xerror:
'       Errores err.Number
'      Exit Sub
End Sub

Sub Busca_PackingList()
Dim StrsqlX As String

    If OptOC.Value = True Then
        StrsqlX = "Exec TI_Muestra_PackingList_Expo '1','" & txtAbr_Cliente & "','" & txtSer_OrdComp & "','" & _
                txtCod_OrdComp & "',null,null,'',''"
    End If
    If OptRango = True Then
        StrsqlX = "Exec TI_Muestra_PackingList_Expo '2','" & txtAbr_Cliente & "','','','" & DTPInicio.Value & "','" & DTPFin.Value & "','',''"
    End If
    
    If Option1 = True Then
        StrsqlX = "Exec TI_Muestra_PackingList_Expo '3','" & txtAbr_Cliente & "','','',null,null,'" & TxtFacturaProforma & "',''"
    End If
    
    If OptPack = True Then
        StrsqlX = "Exec TI_Muestra_PackingList_Expo '4','" & txtAbr_Cliente & "','','',null,null,'','" & TxtPacking & "'"
    End If
        
 Set rsx = CargarRecordSetDesconectado(StrsqlX, cConnect)
 Set GridEX1.ADORecordset = rsx
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo xerror:
Select Case ActionName
    Case "ADICIONAR":
            'If txtAbr_Cliente.Text <> "" And txtNom_Cliente <> "" Then
            FrmPackingListExpo.frm_Opcion = "I"
            FrmPackingListExpo.txtAbr_Cliente = Me.txtAbr_Cliente
            FrmPackingListExpo.txtNom_Cliente = Me.txtNom_Cliente
            FrmPackingListExpo.Show 1
            'Busca_PackingList
           ' End If
    Case "MODIFICAR":
    'If txtAbr_Cliente.Text <> "" And txtNom_Cliente <> "" Then
    If rsx.RecordCount <= 0 Then Exit Sub
    If rsx.RecordCount > 0 Then
        If Trim(GridEX1.Value(GridEX1.Columns("Packing_List").Index)) = "" Then
        MsgBox "Debe Seleccionar un Packing List", vbCritical, "Mensaje"
        Exit Sub
    End If
    End If
    FrmPackingListExpo.frm_Opcion = "U"
    FrmPackingListExpo.cod_PackingList = GridEX1.Value(GridEX1.Columns("Packing_List").Index)
            FrmPackingListExpo.txtAbr_Cliente = GridEX1.Value(GridEX1.Columns("Abr_Cliente").Index)
            FrmPackingListExpo.txtNom_Cliente = GridEX1.Value(GridEX1.Columns("Cliente").Index)
            FrmPackingListExpo.TxtFacturaProforma = GridEX1.Value(GridEX1.Columns("Factura_Proforma").Index)
            FrmPackingListExpo.DTEmision = GridEX1.Value(GridEX1.Columns("Fecha_Emision").Index)
            If Mid(GridEX1.Value(GridEX1.Columns("Tipo_Traslado").Index), 1, 1) = "F" Then
                FrmPackingListExpo.OptFardo.Value = True
                FrmPackingListExpo.TxtNum_Fardos = GridEX1.Value(GridEX1.Columns("Fardos").Index)
                FrmPackingListExpo.tipo_Traslado = "F"
            Else
            FrmPackingListExpo.tipo_Traslado = "R"
                FrmPackingListExpo.OptRollo.Value = True
            End If
            FrmPackingListExpo.txtSer_OrdComp = GridEX1.Value(GridEX1.Columns("Ser_ORdComp").Index)
            FrmPackingListExpo.txtCod_OrdComp = GridEX1.Value(GridEX1.Columns("Cod_ORdComp").Index)
            'FrmPackingListExpo.Show 1
            FrmPackingListExpo.LblPackingList = GridEX1.Value(GridEX1.Columns("Packing_List").Index)
            FrmPackingListExpo.Deshabilita_Campos
            FrmPackingListExpo.LblPackingList.Caption = GridEX1.Value(GridEX1.Columns("Packing_List").Index)
            FrmPackingListExpo.cod_PackingList = GridEX1.Value(GridEX1.Columns("Packing_List").Index)
            If FrmPackingListExpo.tipo_Traslado = "F" Then
                'FrmPackingListExpo.CArgar_Combo
                FrmPackingListExpo.Cargar_Totales_x_FArdo
            End If
                FrmPackingListExpo.carga_GRidRollosAlmacen
                FrmPackingListExpo.Carga_GridRolloFardos
            
            FrmPackingListExpo.CmdAnadir.Enabled = True
            FrmPackingListExpo.CmdEliminar.Enabled = True
            FrmPackingListExpo.Show 1
            'Busca_PackingList
    'End If
            
    Case "ELIMINAR":
        If rsx.RecordCount > 0 Then
            If GridEX1.RowCount Then
                If Trim(GridEX1.Value(GridEX1.Columns("Packing_List").Index)) = "" Then
                    MsgBox "Debe Seleccionar un Packing List", vbCritical, "Mensaje"
                    Exit Sub
                 End If
            End If
        End If
            Dim vmsg As Variant
            Dim i As Integer
            Dim STRSQL As String
            Msg = (MsgBox("¿Desea eliminar el PAcking List Seleccionado?", vbQuestion + vbYesNo, "Confirmacion"))
            If Msg = vbYes Then
            STRSQL = "Exec Elimina_PAcking_List '" & GridEX1.Value(GridEX1.Columns("Abr_Cliente").Index) & "','" & _
                        GridEX1.Value(GridEX1.Columns("Ser_OrdComp").Index) & "','" & GridEX1.Value(GridEX1.Columns("Cod_OrdComp").Index) & "','" & _
                        GridEX1.Value(GridEX1.Columns("Factura_Proforma").Index) & "','" & GridEX1.Value(GridEX1.Columns("Packing_List").Index) & "','" & _
                        GridEX1.Value(GridEX1.Columns("Fecha_Emision").Index) & "','" & GridEX1.Value(GridEX1.Columns("Nro_Despacho").Index) & "'"
            i = ExecuteSQL(cConnect, STRSQL)
            If i = 1 Then
                MsgBox "PAcking List Eliminado correctamente", vbInformation, "Mensaje"
                Busca_PackingList
            Else
                MsgBox "Error se eliminaron " & i & " Registros", vbInformation, "Mensaje"
            End If
            End If
    Case "IMPRIMIR":
            IMPRIMIR_REPORTE
    Case "IMPVALORIZADO"
            IMPRIMIR_REPORTE_VALORIZADO
    Case "SALIR":
        Unload Me

End Select
Exit Sub
xerror:
       errores err.Number
    Exit Sub
End Sub

Sub IMPRIMIR_REPORTE_VALORIZADO()
On Error GoTo xerror:

Dim oo As Object
Dim RsCabX As New ADODB.Recordset
Dim RsDetaX As New ADODB.Recordset
Dim RsDetaKgX As New ADODB.Recordset
Dim StrsqlX1 As String
Dim StrsqlX2 As String
Dim StrsqlX3 As String
Dim StrsUsu As String
Dim Abr_ClienteX As String
Dim Ser_ORdCompX As String
Dim cod_ORdCompX As String
Dim ID_FActuraProformaX As String
Dim Nro_PAckingListX As String


Abr_ClienteX = GridEX1.Value(GridEX1.Columns("Abr_Cliente").Index)
Ser_ORdCompX = GridEX1.Value(GridEX1.Columns("Ser_ORdComp").Index)
cod_ORdCompX = GridEX1.Value(GridEX1.Columns("Cod_ORdComp").Index)
ID_FActuraProformaX = GridEX1.Value(GridEX1.Columns("Factura_Proforma").Index)
Nro_PAckingListX = GridEX1.Value(GridEX1.Columns("Packing_List").Index)

Set oo = CreateObject("Excel.application")
oo.Workbooks.Open vRuta & "\RPT_PackingList_Valorizado.xlt"
oo.Visible = True
oo.DisplayAlerts = False
Screen.MousePointer = 11

StrsqlX1 = "Exec TI_PackingList_Valorizado_Cabecera '" & Abr_ClienteX & "','" & Ser_ORdCompX & "','" & cod_ORdCompX & "','" & ID_FActuraProformaX & "','" & Nro_PAckingListX & "'"
StrsqlX2 = "Exec TI_PackingList_Valorizado_Detalle '" & Abr_ClienteX & "','" & Ser_ORdCompX & "','" & cod_ORdCompX & "','" & ID_FActuraProformaX & "','" & Nro_PAckingListX & "','" & vusu & "'"
StrsqlX3 = "Exec TI_PackingList_Valorizado_Detalle_Kilos '" & Abr_ClienteX & "','" & Ser_ORdCompX & "','" & cod_ORdCompX & "'"

StrsUsu = DevuelveCampo("select Nom_Usuario from seguridad..Seg_Usuarios where cod_Usuario = '" & vusu & "'", cConnect)
Set RsCabX = CargarRecordSetDesconectado(StrsqlX1, cConnect)
Set RsDetaX = CargarRecordSetDesconectado(StrsqlX2, cConnect)
Set RsDetaKgX = CargarRecordSetDesconectado(StrsqlX3, cConnect)
'Dim ab As Integer
'Dim ac As Integer
'ab = RsCabX.RecordCount
'ac = RsDetaX.RecordCount
oo.Run "Reporte", RsCabX, RsDetaX, RsDetaKgX, StrsUsu, cConnect
Set oo = Nothing
Screen.MousePointer = 0
Exit Sub
xerror:
Screen.MousePointer = 0
errores err.Number
Exit Sub
End Sub


Sub IMPRIMIR_REPORTE()
On Error GoTo xerror:

Dim oo As Object
Dim rscab As New ADODB.Recordset
Dim RsDeta1 As New ADODB.Recordset
Dim RsDeta2 As New ADODB.Recordset
Dim RsDeta3 As New ADODB.Recordset
Dim strSQL1 As String
Dim Strsql2 As String
Dim Strsql3 As String
Dim Strsql4 As String
Dim Abr_ClienteX As String
Dim Ser_ORdCompX As String
Dim cod_ORdCompX As String
Dim ID_FActuraProformaX As String
Dim Nro_PAckingListX As String


 Abr_ClienteX = GridEX1.Value(GridEX1.Columns("Abr_Cliente").Index)
Ser_ORdCompX = GridEX1.Value(GridEX1.Columns("Ser_ORdComp").Index)
cod_ORdCompX = GridEX1.Value(GridEX1.Columns("Cod_ORdComp").Index)
ID_FActuraProformaX = GridEX1.Value(GridEX1.Columns("Factura_Proforma").Index)
Nro_PAckingListX = GridEX1.Value(GridEX1.Columns("Packing_List").Index)
Set oo = CreateObject("Excel.application")
oo.Workbooks.Open vRuta & "\RPT_PackingList.xlt"
oo.Visible = True
oo.DisplayAlerts = False
Screen.MousePointer = 11
strSQL1 = "Exec Lista_Cabecera_PAckingList '" & Abr_ClienteX & "','" & Ser_ORdCompX & "','" & cod_ORdCompX & "','" & ID_FActuraProformaX & "','" & Nro_PAckingListX & "'"
Strsql2 = "Exec Lista_Deta_PackingList '" & Abr_ClienteX & "','" & Ser_ORdCompX & "','" & cod_ORdCompX & "','" & ID_FActuraProformaX & "','" & Nro_PAckingListX & "'"
Strsql3 = "Exec Lista_Deta2_PackingList '" & Abr_ClienteX & "','" & Ser_ORdCompX & "','" & cod_ORdCompX & "','" & ID_FActuraProformaX & "','" & Nro_PAckingListX & "'"
Strsql4 = "Exec Lista_Deta3_PackingList '" & Abr_ClienteX & "','" & Ser_ORdCompX & "','" & cod_ORdCompX & "','" & ID_FActuraProformaX & "','" & Nro_PAckingListX & "'"
Set rscab = CargarRecordSetDesconectado(strSQL1, cConnect)
Set RsDeta1 = CargarRecordSetDesconectado(Strsql2, cConnect)
Set RsDeta2 = CargarRecordSetDesconectado(Strsql3, cConnect)
Set RsDeta3 = CargarRecordSetDesconectado(Strsql4, cConnect)

oo.Run "Reporte", rscab, RsDeta1, RsDeta2, RsDeta3, cConnect
Set oo = Nothing
Screen.MousePointer = 0
    Exit Sub
xerror:
Screen.MousePointer = 0
    errores err.Number
    Exit Sub
End Sub

Private Sub Option1_Click()
txtSer_OrdComp.Visible = False
txtCod_OrdComp.Visible = False
Label2.Visible = False
DTPInicio.Visible = False
DTPFin.Visible = False
TxtFacturaProforma.Visible = True
TxtPacking.Visible = False
End Sub

Private Sub OptOC_Click()
txtSer_OrdComp.Visible = True
txtCod_OrdComp.Visible = True
Label2.Visible = False
DTPInicio.Visible = False
DTPFin.Visible = False
TxtFacturaProforma.Visible = False
TxtPacking.Visible = False
End Sub

Private Sub OptPack_Click()
txtSer_OrdComp.Visible = False
txtCod_OrdComp.Visible = False
Label2.Visible = False
DTPInicio.Visible = False
DTPFin.Visible = False
TxtFacturaProforma.Visible = False
TxtPacking.Visible = True
End Sub

Private Sub OptRango_Click()
txtSer_OrdComp.Visible = False
txtCod_OrdComp.Visible = False
Label2.Visible = True
DTPInicio.Visible = True
DTPFin.Visible = True
TxtFacturaProforma.Visible = False
TxtPacking.Visible = False
End Sub


Private Sub txtAbr_Cliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Trim(txtAbr_Cliente.Text) = "" Then
            Call BUSCA_CLIENTE(3)
        Else
            Call BUSCA_CLIENTE(1)
        End If
        txtAbr_Cliente.Text = UCase(txtAbr_Cliente.Text)
    End If
    
End Sub
Public Sub BUSCA_CLIENTE(Tipo As Integer)
Dim STRSQL As String
    Select Case Tipo
        Case 1:
                    STRSQL = "EXEC TI_BUSCA_CLIENTE 1,'" & Trim(Me.txtAbr_Cliente.Text) & "','','" & vusu & "'"
                    Me.txtNom_Cliente.Text = Trim(DevuelveCampo(STRSQL, cConnect))
                    'If Trim(txtNom_Cliente.Text) <> "" Then CARGA_GRID
        Case 2, 3:
                    Dim oTipo As New frmBusGeneral6
                    Dim rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If Tipo = 2 Then
                        oTipo.SQuery = "EXEC TI_BUSCA_CLIENTE 2,'','" & Trim(txtNom_Cliente.Text) & "','" & vusu & "'"
                    Else
                        oTipo.SQuery = "EXEC TI_BUSCA_CLIENTE 3,'','','" & vusu & "'"
                    End If
                    
                    oTipo.CARGAR_DATOS
                    oTipo.DGridLista.Columns(2).Width = 3500
                    oTipo.Show 1
                    If CODIGO <> "" Then
                         Me.txtAbr_Cliente.Text = UCase(Trim(CODIGO))
                         Me.txtNom_Cliente.Text = UCase(Trim(DESCRIPCION))
'                         OptCliPend.SetFocus
                         CODIGO = "": DESCRIPCION = ""
'                         CARGA_GRID
                    End If
                    Set oTipo = Nothing
                    Set rs = Nothing
    End Select
    If txtNom_Cliente.Text <> "" Then
    OptOC.SetFocus
    End If
End Sub

Private Sub txtCod_OrdComp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FunctBuscar.SetFocus
    Else
        Call SoloNumeros(txtCod_OrdComp, KeyAscii, False, 0, 6)
    End If
End Sub

Private Sub txtCod_OrdComp_LostFocus()
    txtCod_OrdComp.Text = Format(Trim(txtCod_OrdComp.Text), "000000")
End Sub

Private Sub TxtFacturaProforma_KeyPress(KeyAscii As Integer)
Call SoloNumeros(TxtFacturaProforma, KeyAscii, False, 0, 8)
End Sub

Private Sub TxtFacturaProforma_LostFocus()
TxtFacturaProforma.Text = Format(Trim(TxtFacturaProforma.Text), "00000000")
End Sub

Private Sub txtNom_Cliente_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        If Trim(txtNom_Cliente.Text) = "" Then
            Call BUSCA_CLIENTE(3)
        Else
            Call BUSCA_CLIENTE(2)
        End If
    End If
End Sub


Private Sub TxtPacking_LostFocus()
TxtPacking.Text = Format(Trim(TxtPacking.Text), "00000000")
End Sub

Private Sub txtSer_OrdComp_LostFocus()
    txtSer_OrdComp.Text = Format(Trim(txtSer_OrdComp.Text), "000")
End Sub

Private Sub txtSer_OrdComp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        txtCod_OrdComp.SetFocus
    Else
        Call SoloNumeros(txtSer_OrdComp, KeyAscii, False, 0, 3)
    End If
End Sub
