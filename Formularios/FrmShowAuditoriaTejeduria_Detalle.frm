VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmShowAuditoriaTejeduria_Detalle 
   Caption         =   "Auditoria Tejeduria Detalle"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   9090
   StartUpPosition =   2  'CenterScreen
   Begin GridEX20.GridEX GridEX1 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   7858
      Version         =   "2.0"
      PreviewRowIndent=   0
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      FormatStylesCount=   7
      FormatStyle(1)  =   "FrmShowAuditoriaTejeduria_Detalle.frx":0000
      FormatStyle(2)  =   "FrmShowAuditoriaTejeduria_Detalle.frx":0138
      FormatStyle(3)  =   "FrmShowAuditoriaTejeduria_Detalle.frx":01E8
      FormatStyle(4)  =   "FrmShowAuditoriaTejeduria_Detalle.frx":029C
      FormatStyle(5)  =   "FrmShowAuditoriaTejeduria_Detalle.frx":0374
      FormatStyle(6)  =   "FrmShowAuditoriaTejeduria_Detalle.frx":042C
      FormatStyle(7)  =   "FrmShowAuditoriaTejeduria_Detalle.frx":050C
      ImageCount      =   0
      PrinterProperties=   "FrmShowAuditoriaTejeduria_Detalle.frx":052C
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   2955
      Left            =   7800
      TabIndex        =   1
      Top             =   0
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   5212
      Custom          =   $"FrmShowAuditoriaTejeduria_Detalle.frx":0704
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   120
   End
End
Attribute VB_Name = "FrmShowAuditoriaTejeduria_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Public sPrefijo_Maquina As String, sCodigo_Rollo As String, sOT As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ADICIONAR"
    Load FrmMante_CC_Tejeduria_Detalle
    Set FrmMante_CC_Tejeduria_Detalle.oParent = Me
    FrmMante_CC_Tejeduria_Detalle.sAccion = "I"
    FrmMante_CC_Tejeduria_Detalle.txtCod_Maquina.Text = sPrefijo_Maquina
    FrmMante_CC_Tejeduria_Detalle.txtDes_Maquina.Text = DevuelveCampo("select des_maquina_tejeduria from tx_maquinas_tejeduria where prefijo_maquina ='" & sPrefijo_Maquina & "'", cConnect)
    FrmMante_CC_Tejeduria_Detalle.TxtCodigo_Rollo.Text = sCodigo_Rollo
    FrmMante_CC_Tejeduria_Detalle.LblOT = sOT
    FrmMante_CC_Tejeduria_Detalle.Show vbModal
    Set FrmMante_CC_Tejeduria_Detalle = Nothing
    Call FunctButt1_ActionClick(0, 0, "SALIR")
Case "MODIFICAR"
    If GridEX1.RowCount = 0 Then Exit Sub
    Load FrmMante_CC_Tejeduria_Detalle
    Set FrmMante_CC_Tejeduria_Detalle.oParent = Me
    FrmMante_CC_Tejeduria_Detalle.sAccion = "U"
    FrmMante_CC_Tejeduria_Detalle.txtCod_Maquina.Text = sPrefijo_Maquina
    FrmMante_CC_Tejeduria_Detalle.txtDes_Maquina.Text = DevuelveCampo("select des_maquina_tejeduria from tx_maquinas_tejeduria where prefijo_maquina ='" & sPrefijo_Maquina & "'", cConnect)
    FrmMante_CC_Tejeduria_Detalle.TxtCodigo_Rollo.Text = sCodigo_Rollo
    FrmMante_CC_Tejeduria_Detalle.LblOT = sOT
    FrmMante_CC_Tejeduria_Detalle.TxtSecuencia.Text = GridEX1.Value(GridEX1.Columns("Secuencia").Index)
    FrmMante_CC_Tejeduria_Detalle.TxtCod_Motivo.Text = GridEX1.Value(GridEX1.Columns("cod_motivo").Index)
    FrmMante_CC_Tejeduria_Detalle.TxtDes_Motivo.Text = GridEX1.Value(GridEX1.Columns("des_motivo").Index)
    FrmMante_CC_Tejeduria_Detalle.txtCantidad.Text = GridEX1.Value(GridEX1.Columns("Cantidad").Index)
    
    If GridEX1.Value(GridEX1.Columns("flg_contar").Index) = "S" Then
        FrmMante_CC_Tejeduria_Detalle.ChkContar.Value = Checked
    Else
        FrmMante_CC_Tejeduria_Detalle.ChkContar.Value = Unchecked
    End If
    
    FrmMante_CC_Tejeduria_Detalle.Show vbModal
    Set FrmMante_CC_Tejeduria_Detalle = Nothing
Case "ELIMINAR"
    If GridEX1.RowCount = 0 Then Exit Sub
    Load FrmMante_CC_Tejeduria_Detalle
    Set FrmMante_CC_Tejeduria_Detalle.oParent = Me
    FrmMante_CC_Tejeduria_Detalle.sAccion = "D"
    FrmMante_CC_Tejeduria_Detalle.fraDatos.Enabled = False
    FrmMante_CC_Tejeduria_Detalle.txtCod_Maquina.Text = sPrefijo_Maquina
    FrmMante_CC_Tejeduria_Detalle.txtDes_Maquina.Text = DevuelveCampo("select des_maquina_tejeduria from tx_maquinas_tejeduria where prefijo_maquina ='" & sPrefijo_Maquina & "'", cConnect)
    FrmMante_CC_Tejeduria_Detalle.TxtCodigo_Rollo.Text = sCodigo_Rollo
    FrmMante_CC_Tejeduria_Detalle.LblOT = sOT
    FrmMante_CC_Tejeduria_Detalle.TxtSecuencia.Text = GridEX1.Value(GridEX1.Columns("Secuencia").Index)
    FrmMante_CC_Tejeduria_Detalle.TxtCod_Motivo.Text = GridEX1.Value(GridEX1.Columns("cod_motivo").Index)
    FrmMante_CC_Tejeduria_Detalle.TxtDes_Motivo.Text = GridEX1.Value(GridEX1.Columns("des_motivo").Index)
    FrmMante_CC_Tejeduria_Detalle.txtCantidad.Text = GridEX1.Value(GridEX1.Columns("Cantidad").Index)
    
    If GridEX1.Value(GridEX1.Columns("flg_contar").Index) = "S" Then
        FrmMante_CC_Tejeduria_Detalle.ChkContar.Value = Checked
    Else
        FrmMante_CC_Tejeduria_Detalle.ChkContar.Value = Unchecked
    End If
    
    FrmMante_CC_Tejeduria_Detalle.Show vbModal
    Set FrmMante_CC_Tejeduria_Detalle = Nothing
    
 Case "MASIVO"
    Load FrmCCTejDetalleAddMasivo
    Set FrmCCTejDetalleAddMasivo.oParent = Me
    FrmCCTejDetalleAddMasivo.sAccion = "I"
    FrmCCTejDetalleAddMasivo.txtCod_Maquina.Text = sPrefijo_Maquina
    FrmCCTejDetalleAddMasivo.txtDes_Maquina.Text = DevuelveCampo("select des_maquina_tejeduria from tx_maquinas_tejeduria where prefijo_maquina ='" & sPrefijo_Maquina & "'", cConnect)
    FrmCCTejDetalleAddMasivo.TxtCodigo_Rollo.Text = sCodigo_Rollo
    FrmCCTejDetalleAddMasivo.LblOT = sOT
    FrmCCTejDetalleAddMasivo.BUSCAR
    FrmCCTejDetalleAddMasivo.Show vbModal
    Set FrmCCTejDetalleAddMasivo = Nothing
    BUSCAR
    
Case "SALIR"
    Unload Me
End Select
End Sub

Sub BUSCAR()
strSQL = "CC_MUESTRA_AUDITORIA_TEJEDURIA_ROLLOS_DETALLE '" & sPrefijo_Maquina & "','" & sCodigo_Rollo & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)

GridEX1.Columns("Secuencia").Width = 900
GridEX1.Columns("cod_motivo").Width = 0
GridEX1.Columns("Des_motivo").Width = 3500
GridEX1.Columns("Cantidad").Width = 1200
GridEX1.Columns("Flg_contar").Width = 700

GridEX1.Columns("Secuencia").Caption = "Sec."
GridEX1.Columns("Des_motivo").Caption = "Motivo"
GridEX1.Columns("flg_contar").Caption = "Contar"

End Sub





