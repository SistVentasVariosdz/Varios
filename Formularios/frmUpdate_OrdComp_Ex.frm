VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUpdate_OrdComp_EX 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizar O.C."
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2970
      Left            =   60
      TabIndex        =   15
      Top             =   45
      Width           =   9405
      Begin VB.ComboBox cmbViaTransporte_Cli 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   2040
         Width           =   3555
      End
      Begin VB.TextBox TxtPO 
         Height          =   285
         Left            =   1080
         MaxLength       =   60
         TabIndex        =   8
         Top             =   1635
         Width           =   3435
      End
      Begin VB.ComboBox cmbPais 
         Height          =   315
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1680
         Width           =   3195
      End
      Begin VB.TextBox txt_totalcarga 
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox cboTipoOC 
         Height          =   315
         Left            =   6105
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   675
         Width           =   3165
      End
      Begin VB.TextBox txtDes_LugEntr 
         Height          =   285
         Left            =   1665
         TabIndex        =   5
         Top             =   915
         Width           =   2895
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   495
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   2400
         Width           =   7500
      End
      Begin VB.TextBox txtPorc_IGV 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         Enabled         =   0   'False
         Height          =   285
         Left            =   8670
         TabIndex        =   16
         Top             =   285
         Width           =   435
      End
      Begin VB.TextBox txtDes_Descuento 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   570
         Width           =   2865
      End
      Begin VB.TextBox txtCod_Descuento 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   570
         Width           =   510
      End
      Begin VB.TextBox txtDes_CondVent 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   2865
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   6105
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   285
         Width           =   2010
      End
      Begin VB.TextBox txtCod_LugEntr 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   915
         Width           =   495
      End
      Begin VB.TextBox txtCod_CondVent 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   510
      End
      Begin VB.CheckBox ChkCtrPeso 
         BackColor       =   &H80000000&
         Caption         =   "Control de Peso x Rollo"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   2025
      End
      Begin MSComCtl2.DTPicker dtpEntregaI 
         Height          =   315
         Left            =   6120
         TabIndex        =   11
         Top             =   1020
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   72024067
         CurrentDate     =   37832
      End
      Begin MSComCtl2.DTPicker dtpEntregaF 
         Height          =   315
         Left            =   6120
         TabIndex        =   12
         Top             =   1350
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   72024067
         CurrentDate     =   37832
      End
      Begin VB.Label Label12 
         Caption         =   "Via Transp"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label28 
         Caption         =   "P.O."
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   120
         TabIndex        =   29
         Top             =   1695
         Width           =   900
      End
      Begin VB.Label Label27 
         Caption         =   "País Destino"
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4680
         TabIndex        =   28
         Top             =   1680
         Width           =   1500
      End
      Begin VB.Label Label26 
         Caption         =   "Total Carga"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   2160
         TabIndex        =   27
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label Label30 
         BackColor       =   &H80000004&
         Caption         =   "Tipo"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   4680
         TabIndex        =   26
         Top             =   705
         Width           =   960
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000004&
         Caption         =   "Observ:"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   165
         TabIndex        =   25
         Top             =   2400
         Width           =   885
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000004&
         Caption         =   "I.G.V."
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   8175
         TabIndex        =   24
         Top             =   315
         Width           =   450
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000004&
         Caption         =   "Descuento"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   120
         TabIndex        =   23
         Top             =   615
         Width           =   900
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000004&
         Caption         =   "Fec. Entrega Fin"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   4710
         TabIndex        =   22
         Top             =   1395
         Width           =   1530
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000004&
         Caption         =   "Moneda"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   4695
         TabIndex        =   21
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000004&
         Caption         =   "Lug. Entrega"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1020
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000004&
         Caption         =   "Fec. Entrega Inicio"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   4710
         TabIndex        =   19
         Top             =   1065
         Width           =   1530
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000004&
         Caption         =   "C. de Pago"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   285
         Width           =   900
      End
      Begin VB.Label Label23 
         BackColor       =   &H80000004&
         Caption         =   "%"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   9135
         TabIndex        =   17
         Top             =   315
         Width           =   165
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   3600
      TabIndex        =   32
      Top             =   3240
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmUpdate_OrdComp_Ex.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmUpdate_OrdComp_EX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rstAux1 As ADODB.Recordset
Public CODIGO, descripcion As String, TipoAdd As String
Public sCod_Cliente_Tex As String
Public sSer_OrdComp As String
Public sCod_OrdComp As String
Dim strSQL As String
Dim sCod_ClaOrdComp As String

Private Sub Form_Load()
FillPais
FillViaTransporte
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            If VALIDA_DATOS Then
                SALVAR_DATOS
            End If
        Case "CANCELAR"
            Unload Me
    End Select
End Sub

Private Sub BuscaCondPago(opcion As Integer)
Dim rstAux As New ADODB.Recordset
On Error GoTo Fin
    strSQL = "SELECT Cod_CondVent, Des_CondVent " & _
             "FROM Lg_CondVent WHERE "
    txtCod_CondVent = Trim(txtCod_CondVent)
    txtDes_CondVent = Trim(txtDes_CondVent)
    Select Case opcion
    Case 1: strSQL = strSQL & "Cod_CondVent = '" & txtCod_CondVent & "'"
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
    "Busqueda de Cliente (" & opcion & ")"
End Sub

Private Sub BuscaDescuento(opcion As Integer)
Dim rstAux As New ADODB.Recordset
On Error GoTo Fin
    strSQL = "SELECT Cod_Descuento, Des_Descuento, Porcentaje1 " & _
             "FROM LG_DSCTOS WHERE "
    txtCod_Descuento = Trim(txtCod_Descuento)
    txtDes_Descuento = Trim(txtDes_Descuento)
    Select Case opcion
    Case 1: strSQL = strSQL & "Cod_Descuento = '" & txtCod_Descuento & "'"
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
    "Busqueda de Descuento (" & opcion & ")"
End Sub

Private Sub BuscaLugEntr(opcion As Integer)
Dim rstAux As New ADODB.Recordset
On Error GoTo Fin
    strSQL = "SELECT Cod_LugEntr, Des_LugEntr FROM LG_LUGENTR WHERE "
    
    txtCod_LugEntr = Trim(txtCod_LugEntr)
    txtDes_LugEntr = Trim(txtDes_LugEntr)
    Select Case opcion
    Case 1: strSQL = strSQL & "Cod_LugEntr = '" & txtCod_LugEntr & "'"
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
    "Busqueda de Lugar de Entrega (" & opcion & ")"
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

Private Sub txtCod_LugEntr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then BuscaLugEntr 1
End Sub

Private Sub txtDes_LugEntr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then BuscaLugEntr 2
End Sub

Public Sub Carga_Data()
Dim Reg As New ADODB.Recordset
On Error GoTo hand

    FillMoneda
    FillTipoOC
    
    
    Reg.ActiveConnection = cConnect
    Reg.CursorType = adOpenStatic
    Reg.CursorLocation = adUseClient
    
    strSQL = "select a.*,b.Cod_Pais+space(5)+b.Descripcion as Pais,ISNULL(RTRIM(c.idViaTransporteKey),'')+space(5)+ISNULL(RTRIM(c.NombreVia),'') as Via from tx_ordcomp a Inner Join cn_paises b On a.Cod_Pais=b.Cod_Pais " _
    & " Left Outer Join Tx_MViaTransporte c On c.idViaTransporteKey=a.idViaTransporte   where a.cod_cliente_tex='" & sCod_Cliente_Tex & "' and a.ser_ordcomp='" & sSer_OrdComp & "' and a.cod_ordcomp='" & sCod_OrdComp & "'"
    
    Reg.Open strSQL
    If Reg.RecordCount Then
        Reg.MoveFirst
        With Reg
            txt_totalcarga.Text = IIf(IsNull(Trim(!CantidadTotal)) = True, 0, Trim(!CantidadTotal))
            txtCod_CondVent.Text = Trim(!Cod_CondVent)
            BuscaCondPago 1
            txtCod_LugEntr.Text = Trim(!cod_lugentr)
            BuscaLugEntr 1
            txtCod_Descuento.Text = Trim(!cod_descuento)
            BuscaDescuento 1
            Call BuscaCombo1(Trim(!Cod_Moneda), 2, cboMoneda)
            Call BuscaCombo1(Trim(!cod_tipooc_tintoreria), 2, cboTipoOC)
            txtPorc_IGV.Text = CDbl(!porc_igv)
            dtpEntregaI.Value = !fec_entrega_inicio
            dtpEntregaF.Value = !fec_entrega_fin
            If Trim(!flg_crtpeso_x_rollo) = "S" Then
                ChkCtrPeso.Value = 1
            Else
                ChkCtrPeso.Value = 0
            End If
            sCod_ClaOrdComp = Trim(!Cod_ClaOrdComp)
            txtObservaciones.Text = Trim(!OBSERVACIONES)
            cmbPais.Text = !Pais
            TxtPO.Text = IIf(IsNull(!PO) = True, "", !PO)
            
            
            If Trim(!Via) = "" Then
                cmbViaTransporte_Cli.ListIndex = -1
            Else
                
                BuscaCombo1 Trim(!Via), 1, cmbViaTransporte_Cli
            End If
            
        End With
        Valida_Estado_TipoOC_Tintoreria
    End If
    
Exit Sub
hand:
    ErrorHandler Err, "Carga_Data"
End Sub
Private Sub FillViaTransporte()

    strSQL = "SELECT idViaTransporteKey,NombreVia FROM Tx_MViaTransporte"
    
    Set rstAux1 = CargarRecordSetDesconectado(strSQL, cConnect)
    
    cmbViaTransporte_Cli.Clear
    With rstAux1
    If .RecordCount > 0 Then .MoveFirst
    Do Until .EOF
        cmbViaTransporte_Cli.AddItem !idViaTransporteKey & Space(5) & !NombreVia
        .MoveNext
    Loop

    End With

    BuscaCombo1 "C", 5, cmbViaTransporte_Cli
End Sub

Private Sub FillMoneda()
Dim iRow As Long
Dim rstAux As New ADODB.Recordset
    strSQL = "SELECT Cod_Moneda, Nom_Moneda, Flg_Principal FROM TG_Moneda"
    Set rstAux = CargarRecordSetDesconectado(strSQL, cConnect)
    
    cboMoneda.Clear
    With rstAux
    If .RecordCount > 0 Then .MoveFirst
    iRow = 0
    Do Until .EOF
        cboMoneda.AddItem !Nom_Moneda & Space(100) & !Cod_Moneda
        If !Flg_Principal = "*" Then cboMoneda.ListIndex = iRow
        iRow = iRow + 1
        .MoveNext
    Loop
    .Close
    End With
    Set rstAux = Nothing
End Sub

Private Sub FillTipoOC()
Dim rstAux As New ADODB.Recordset
'    strSQL = "SELECT Cod_Tipo_Orden_tinto, Descripcion " & _
'             "FROM Ti_Tipo_Orden_Tintoreria where Tip_Item='T'"

    strSQL = "EXEC Usp_Lista_TipoOrdenServicioExportacion"
             
    Set rstAux = CargarRecordSetDesconectado(strSQL, cConnect)
    
    cboTipoOC.Clear
    With rstAux
    If .RecordCount > 0 Then .MoveFirst
    Do Until .EOF
        cboTipoOC.AddItem !descripcion & Space(100) & !Cod_Tipo_Orden_tinto
        .MoveNext
    Loop
    .Close
    End With
    Set rstAux = Nothing
    BuscaCombo1 "C", 100, cboTipoOC
End Sub

Sub Valida_Estado_TipoOC_Tintoreria()
    strSQL = "select count(*) from ti_ordtra_tintoreria where cod_cliente_tex='" & sCod_Cliente_Tex & "' and ser_ordcomp='" & sSer_OrdComp & "' and cod_ordcomp='" & sCod_OrdComp & "'"
    If DevuelveCampo(strSQL, cConnect) > 0 Then
        cboTipoOC.Enabled = False
        Exit Sub
    End If

    strSQL = "select count(*) from TI_Guias_Crudo_Recibidas_Tintoreria where cod_cliente_tex='" & sCod_Cliente_Tex & "' and ser_ordcomp='" & sSer_OrdComp & "' and cod_ordcomp='" & sCod_OrdComp & "'"
    If DevuelveCampo(strSQL, cConnect) > 0 Then
        cboTipoOC.Enabled = False
        Exit Sub
    End If

End Sub
Private Sub FillPais()

    strSQL = "select Cod_Pais,Descripcion from CN_PAISES"
    
    Set rstAux1 = CargarRecordSetDesconectado(strSQL, cConnect)
    
    cmbPais.Clear
    With rstAux1
    If .RecordCount > 0 Then .MoveFirst
    Do Until .EOF
        cmbPais.AddItem !Cod_Pais & Space(5) & !descripcion
        .MoveNext
    Loop
    
    End With
    
    BuscaCombo1 "C", 5, cmbPais
End Sub
Sub SALVAR_DATOS()
Dim i As Integer
On Error GoTo hand

    strSQL = "EXEC TI_MAN_TX_ORDCOMP_EXP 'U','" & _
        sCod_Cliente_Tex & "','" & _
        sSer_OrdComp & "','" & _
        sCod_OrdComp & "','" & _
        txtCod_CondVent.Text & "','" & _
        txtCod_Descuento.Text & "'," & _
        txtPorc_IGV.Text & ",'" & _
        Trim(Right(cboMoneda, 5)) & "','" & _
        txtCod_LugEntr.Text & "','" & _
        Trim(txtObservaciones.Text) & "','" & _
        sCod_ClaOrdComp & "','" & _
        dtpEntregaI.Value & "','" & _
        dtpEntregaF.Value & "','" & _
        Trim(Right(cboTipoOC.Text, 5)) & "','" & _
        IIf(ChkCtrPeso.Value = 1, "S", "N") & "'," & _
        CDbl(txt_totalcarga.Text) & ",'" & Left(cmbPais.Text, 3) & "','" & Trim(TxtPO.Text) & "','" & Left(cmbViaTransporte_Cli.Text, 2) & "'"
            
        Call ExecuteSQL(cConnect, strSQL)
    
    Unload Me
Exit Sub
hand:
    ErrorHandler Err, "SALVAR_DATOS"
End Sub

Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True
    
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
    
    If Trim(txt_totalcarga.Text) = "" Then
        MsgBox "Debe ingresar el total de la carga para esta Orden de Servicio", vbCritical, "Valida O.C."
        VALIDA_DATOS = False
        Exit Function
    End If
    
    If CDbl(txt_totalcarga.Text) = 0 Then
        MsgBox "Debe ingresar el total de la carga para esta Orden de Servicio", vbCritical, "Valida O.C."
        VALIDA_DATOS = False
        Exit Function
    End If
    
    
End Function

