VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMante_CC_Tejeduria 
   Caption         =   "Auditoria de Tejeduria"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDatos 
      Height          =   4335
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   6495
      Begin VB.TextBox TxtTip_Tejedor 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Text            =   "O"
         Top             =   3120
         Width           =   300
      End
      Begin VB.TextBox TxtCod_Tejedor 
         Height          =   285
         Left            =   1800
         TabIndex        =   15
         Top             =   3120
         Width           =   780
      End
      Begin VB.TextBox TxtNom_Tejedor 
         Height          =   285
         Left            =   2640
         TabIndex        =   16
         Top             =   3120
         Width           =   3615
      End
      Begin VB.TextBox TxtDes_Restriccion 
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Top             =   2760
         Width           =   4335
      End
      Begin VB.TextBox TxtRestriccion 
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   2760
         Width           =   540
      End
      Begin VB.TextBox TxtTurno 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   1320
         Width           =   540
      End
      Begin VB.TextBox TxtObservaciones 
         Height          =   645
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   3480
         Width           =   4980
      End
      Begin VB.TextBox TxtMerma 
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   2400
         Width           =   660
      End
      Begin VB.TextBox TxtCod_Auditor 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox TxtNom_Auditor 
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox TxtTip_Auditor 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Text            =   "O"
         Top             =   960
         Width           =   255
      End
      Begin MSComCtl2.DTPicker DTPFecha 
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   72351745
         CurrentDate     =   38415
      End
      Begin VB.TextBox TxtCod_Calidad 
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   1680
         Width           =   540
      End
      Begin VB.TextBox TxtDes_Calidad 
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Top             =   1680
         Width           =   4335
      End
      Begin VB.TextBox TxtCodigo_Rollo 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   600
         Width           =   900
      End
      Begin VB.TextBox txtDes_Maquina 
         Height          =   285
         Left            =   2280
         TabIndex        =   1
         Top             =   240
         Width           =   3975
      End
      Begin VB.TextBox txtCod_Maquina 
         Height          =   285
         Left            =   1335
         TabIndex        =   0
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tejedor"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   3225
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Restriccion"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   2865
         Width           =   795
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Turno"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   1425
         Width           =   420
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   3585
         Width           =   1065
      End
      Begin VB.Label LabelKilos 
         AutoSize        =   -1  'True
         Caption         =   "Kgs. Rollo"
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
         Left            =   4320
         TabIndex        =   27
         Top             =   675
         Width           =   885
      End
      Begin VB.Label LblKilos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5400
         TabIndex        =   26
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Merma Kgs."
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   2505
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Auditor:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Auditoria"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   2130
         Width           =   1110
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Calidad"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1785
         Width           =   525
      End
      Begin VB.Label LblOT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "OT"
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
         Left            =   2640
         TabIndex        =   21
         Top             =   675
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo Rollo"
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
         Left            =   120
         TabIndex        =   20
         Top             =   675
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Máquina"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   19
         Top             =   345
         Width           =   855
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1560
      TabIndex        =   32
      Top             =   4440
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmMante_CC_Tejeduria.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "FrmMante_CC_Tejeduria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Public sAccion As String
Public CODIGO As String, Descripcion As String, TipoAdd As String

Public oParent As Object
Public vOk As Boolean
Public fila_seleccionada  As Long
Private Sub BUSCAMAQUINA(Opcion As Integer)
On Error GoTo Fin
Dim rstAux As ADODB.Recordset
    strSQL = "SELECT Prefijo_Maquina, Des_Maquina_Tejeduria FROM TX_MAQUINAS_TEJEDURIA WHERE "
    txtCod_Maquina = Trim(txtCod_Maquina)
    txtDes_Maquina = Trim(txtDes_Maquina)
    Select Case Opcion
    Case 1: strSQL = strSQL & "Prefijo_Maquina like '%" & txtCod_Maquina & "%'"
    Case 2: strSQL = strSQL & "Des_Maquina_Tejeduria like '%" & txtDes_Maquina & "%'"
    End Select
    txtCod_Maquina = ""
    txtDes_Maquina = ""
    fila_seleccionada = 0
    With frmBusqGeneral
        Set .oParent = Me
        .SQuery = strSQL
        .Cargar_Datos
        
        .gexList.Columns("Prefijo_Maquina").Caption = "Prefijo"
        .gexList.Columns("Prefijo_Maquina").Width = 1000
        .gexList.Columns("Des_Maquina_Tejeduria").Caption = "Máquina"
        .gexList.Columns("Des_Maquina_Tejeduria").Width = 5000
        
        CODIGO = ".."
        Set rstAux = .gexList.ADORecordset
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If fila_seleccionada > 0 And rstAux.RecordCount > 0 Then
        
            rstAux.AbsolutePosition = fila_seleccionada
            'txtCod = Trim(rstAux!cod)
            'txtDes = Trim(rstAux!Descripcion)
            txtCod_Maquina = Trim(rstAux!prefijo_maquina)
            txtDes_Maquina = Trim(rstAux!Des_Maquina_Tejeduria)
            TxtCodigo_Rollo.SetFocus
            
        Else
            SendKeys "{TAB}"
        End If
        
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
    "Búsqueda de Maquina (" & Opcion & ")"
End Sub

Private Sub ChkContar_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub DTPFecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub


Private Sub Form_Load()
Dim iHora As Integer
'If vemp = "07" Then
'    iHora = DevuelveCampo("SELECT DATEPART(HH,GETDATE())", cCONNECT)
'
'    If iHora >= 8 And iHora <= 16 Then
'        TxtTurno.Text = "1"
'    Else
'        TxtTurno.Text = "2"
'    End If
'
'    dtpFecha.Enabled = False
'
'End If
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    Call Grabar
Case "SALIR"
    vOk = False
    Unload Me
End Select
End Sub


Private Sub TxtCod_Auditor_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Busca_Auditor 2
End Sub

Private Sub TxtCod_Calidad_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then BUSCACALIDAD 1
End Sub

Private Sub txtCod_Maquina_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then BUSCAMAQUINA 1
End Sub

Private Sub TxtCod_Tejedor_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then BUSCATEJEDOR 2
End Sub

Private Sub TxtCodigo_Rollo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    TxtCodigo_Rollo.Text = UCase(TxtCodigo_Rollo)
    LblOT = DevuelveCampo("select cod_ordtra from tj_ordtra_tejeduria_rollos where prefijo_maquina ='" & Trim(txtCod_Maquina.Text) & "' and codigo_rollo='" & Trim(TxtCodigo_Rollo.Text) & "'", cConnect)
    LblKilos = DevuelveCampo("select kgs_producidos from tj_ordtra_tejeduria_rollos where prefijo_maquina ='" & Trim(txtCod_Maquina.Text) & "' and codigo_rollo='" & Trim(TxtCodigo_Rollo.Text) & "'", cConnect)
    If Trim(LblOT) = "" Then
        MsgBox "Rollo no existe", vbCritical
        TxtCodigo_Rollo.SetFocus
        Exit Sub
    Else
        TxtTip_Tejedor.Text = Trim(DevuelveCampo("select tip_trabajador_tejedor from tj_ordtra_tejeduria_rollos where prefijo_maquina ='" & Trim(txtCod_Maquina.Text) & "' and codigo_rollo='" & Trim(TxtCodigo_Rollo.Text) & "'", cConnect))
        TxtCod_Tejedor.Text = Trim(DevuelveCampo("select cod_trabajador_tejedor from tj_ordtra_tejeduria_rollos where prefijo_maquina ='" & Trim(txtCod_Maquina.Text) & "' and codigo_rollo='" & Trim(TxtCodigo_Rollo.Text) & "'", cConnect))
        Call BUSCATEJEDOR(2)
        TxtCod_Auditor.SetFocus
    End If
End If
End Sub

Private Sub TxtDes_Calidad_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then BUSCACALIDAD 2
End Sub

Private Sub txtDes_Maquina_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then BUSCAMAQUINA 2
End Sub


'Private Sub BUSCAROLLO(Opcion As Integer)
'On Error GoTo Fin
'Dim rstAux As ADODB.Recordset
'    strSQL = "Tj_muestra_Rollos_Maquina '" & Trim(txtCod_Maquina.Text) & "','" & Trim(TxtCodigo_Rollo.Text) & "'"
'
'    With frmBusqGeneral
'        Set .oParent = Me
'        .sQuery = strSQL
'        .Cargar_Datos
'
'
'        .DGridLista.Columns("codigo_Rollo").Width = 1200
'        .DGridLista.Columns("OT").Width = 1000
'        .DGridLista.Columns("Num_Secuencia").Width = 1000
'        .DGridLista.Columns("Num_Rollo").Width = 1200
'
'        Codigo = ".."
'        Set rstAux = .DGridLista.ADORecordset
'        If rstAux.RecordCount > 1 Then .Show vbModal
'
'        If Codigo <> "" And rstAux.RecordCount > 0 Then
'            TxtCodigo_Rollo.Text = Trim(rstAux!Codigo_Rollo)
'            LblOT = Trim(rstAux!OT)
'            'txtDes_Maquina = Trim(rstAux!Des_Maquina_Tejeduria)
'            TxtCod_Motivo.SetFocus
'        Else
'            SendKeys "{TAB}"
'        End If
'
'    End With
'    Unload frmBusqGeneral
'    Set frmBusqGeneral = Nothing
'    rstAux.Close
'    Set rstAux = Nothing
'Exit Sub
'Fin:
'On Error Resume Next
'    Unload frmBusqGeneral
'    Set frmBusqGeneral = Nothing
'    rstAux.Close
'    Set rstAux = Nothing
'    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
'    "Búsqueda de Rollo (" & Opcion & ")"
'End Sub

Sub Grabar()
On Error GoTo errGrabar

If Trim(TxtMerma.Text) = "" Then TxtMerma.Text = "0"

If Trim(TxtTurno.Text) = "" Then
    MsgBox "Debe ingresar turno", vbCritical
    TxtTurno.SetFocus
    Exit Sub
End If

strSQL = "CC_MAN_AUDITORIA_TEJEDURIA_CABECERA '" & sAccion & "','" & Trim(txtCod_Maquina.Text) & "','" & Trim(TxtCodigo_Rollo.Text) & "','" & TxtTip_Auditor.Text & "','" & Me.TxtCod_Auditor.Text & "','" & Me.DTPFecha.Value & "','" & Trim(TxtCod_Calidad.Text) & "'," & CDbl(TxtMerma.Text) & ",'" & TxtObservaciones.Text & "','" & Trim(TxtTurno.Text) & "','" & Trim(TxtRestriccion.Text) & "','" & TxtTip_Tejedor.Text & "','" & Trim(TxtCod_Tejedor.Text) & "','" & vusu & "'"

ExecuteCommandSQL cConnect, strSQL

Me.oParent.DTPFecha.Value = DTPFecha.Value
Me.oParent.CARGA_GRID
vOk = True

Unload Me

Exit Sub
errGrabar:
    vOk = False
    ErrorHandler err, "Grabar"
End Sub

Sub BUSCACALIDAD(tipo As Integer)
Dim oTipo As New frmBusqGeneral3
Dim rs As New ADODB.Recordset

Set oTipo.oParent = Me

If tipo = 1 Then
    oTipo.SQuery = "select cod_Calidad as Codigo, Des_Calidad as Descripcion from Tx_Calidad_Rollos where cod_Calidad like '%" & Trim(Me.TxtCod_Calidad.Text) & "%' and cod_calidad <>'0'"
ElseIf tipo = 2 Then
    oTipo.SQuery = "select cod_Calidad as Codigo, Des_Calidad as Descripcion from Tx_Calidad_Rollos where des_Calidad like '%" & Trim(Me.TxtDes_Calidad.Text) & "%' and cod_calidad <>'0'"
End If

oTipo.Caption = "Buscar Calidad"
oTipo.Cargar_Datos

oTipo.gexLista.Columns("Codigo").Width = 1400
oTipo.gexLista.Columns("Descripcion").Width = 5000

If oTipo.gexLista.RowCount > 1 Then
    oTipo.Show vbModal
Else
    CODIGO = oTipo.gexLista.Value(oTipo.gexLista.Columns("Codigo").Index)
    Descripcion = oTipo.gexLista.Value(oTipo.gexLista.Columns("Descripcion").Index)
End If

If Trim(CODIGO) <> "" Then
    TxtCod_Calidad.Text = CODIGO
    TxtDes_Calidad.Text = Descripcion
    CODIGO = "": Descripcion = "": TipoAdd = ""
    'dtpFecha.SetFocus
    TxtMerma.SetFocus
End If
Unload oTipo
Set oTipo = Nothing
Set rs = Nothing
End Sub

Private Sub TxtDes_Restriccion_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then BUSCARESTRICCION 2
TxtCod_Tejedor.SetFocus
End Sub

Private Sub TxtMerma_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(TxtMerma, KeyAscii, True, 2)
End If
End Sub

Private Sub TxtNom_Auditor_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Busca_Auditor 3
End Sub

Private Sub TxtNom_Tejedor_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then BUSCATEJEDOR 3
End Sub

Private Sub TxtObservaciones_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtRestriccion_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then BUSCARESTRICCION 1
TxtCod_Tejedor.SetFocus
End Sub

Private Sub TxtTip_Auditor_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Busca_Auditor 1
End Sub

'Sub Busca_Auditor(Tipo As Integer)
'Dim oTipo As New frmBusqGeneral3
'Dim rs As New ADODB.Recordset
'
'Set oTipo.oParent = Me
'
'If Tipo = 1 Then
'    oTipo.sQuery = "select tip_Auditor as Tipo, cod_Auditor as Codigo, Nom_Auditor as Descripcion from ti_cc_auditor a, cc_areas b where a.area_auditor = b.cod_area_cc and tip_Auditor='" & Trim(TxtTip_Auditor.Text) & "' and Flg_Tejeduria ='*' order by tip_auditor, cod_auditor"
'ElseIf Tipo = 2 Then
'    oTipo.sQuery = "select tip_Auditor as Tipo, cod_Auditor as Codigo, Nom_Auditor as Descripcion from ti_cc_auditor  a, cc_areas b where a.area_auditor = b.cod_area_cc and cod_Auditor like '%" & Trim(TxtCod_Auditor.Text) & "%'  and Flg_Tejeduria ='*' order by tip_auditor, cod_auditor"
'Else
'    oTipo.sQuery = "select tip_Auditor as Tipo, cod_Auditor as Codigo, Nom_Auditor as Descripcion from ti_cc_auditor  a, cc_areas b where a.area_auditor = b.cod_area_cc and nom_Auditor like '%" & Trim(TxtNom_Auditor.Text) & "%' and  Flg_Tejeduria ='*' order by tip_auditor, nom_auditor"
'End If
'
'oTipo.Caption = "Buscar Auditor"
'oTipo.Cargar_Datos
'
'oTipo.gexLista.Columns("Tipo").Width = 750
'oTipo.gexLista.Columns("Codigo").Width = 1200
'oTipo.gexLista.Columns("Descripcion").Width = 4400
'
'If oTipo.gexLista.RowCount > 1 Then
'    oTipo.Show vbModal
'Else
'    Codigo = oTipo.gexLista.Value(oTipo.gexLista.Columns("Tipo").Index)
'    Descripcion = oTipo.gexLista.Value(oTipo.gexLista.Columns("Codigo").Index)
'     TipoAdd = oTipo.gexLista.Value(oTipo.gexLista.Columns("Descripcion").Index)
'End If
'
'If Trim(Codigo) <> "" Then
'    TxtTip_Auditor.Text = Codigo
'    TxtCod_Auditor.Text = Descripcion
'    TxtNom_Auditor.Text = TipoAdd
'    Codigo = "": Descripcion = "": TipoAdd = ""
'    TxtTurno.SetFocus
'End If
'Unload oTipo
'Set oTipo = Nothing
'Set rs = Nothing
'End Sub

Private Sub TxtTip_Tejedor_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then BUSCATEJEDOR 1
End Sub

Private Sub TxtTurno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub


Sub BUSCARESTRICCION(tipo As Integer)
Dim oTipo As New frmBusqGeneral3
Dim rs As New ADODB.Recordset

Set oTipo.oParent = Me

If tipo = 1 Then
    oTipo.SQuery = "select Cod_Restriccion as Codigo, Des_Restriccion as Descripcion from CC_Tejeduria_Restriccion_Rollos where cod_restriccion like '%" & Trim(Me.TxtRestriccion.Text) & "%'"
ElseIf tipo = 2 Then
    oTipo.SQuery = "select Cod_Restriccion as Codigo, Des_Restriccion as Descripcion from CC_Tejeduria_Restriccion_Rollos where des_restriccion like '%" & Trim(Me.TxtDes_Restriccion.Text) & "%'"
End If

oTipo.Caption = "Buscar Restriccion"
oTipo.Cargar_Datos

oTipo.gexLista.Columns("Codigo").Width = 1400
oTipo.gexLista.Columns("Descripcion").Width = 5000

If oTipo.gexLista.RowCount > 1 Then
    oTipo.Show vbModal
Else
    CODIGO = oTipo.gexLista.Value(oTipo.gexLista.Columns("Codigo").Index)
    Descripcion = oTipo.gexLista.Value(oTipo.gexLista.Columns("Descripcion").Index)
End If

If Trim(CODIGO) <> "" Then
    TxtRestriccion.Text = CODIGO
    TxtDes_Restriccion.Text = Descripcion
    CODIGO = "": Descripcion = "": TipoAdd = ""
End If
Unload oTipo
Set oTipo = Nothing
Set rs = Nothing
End Sub


Sub BUSCATEJEDOR(tipo As Integer)
Dim oTipo As New frmBusqGeneral3
Dim rs As New ADODB.Recordset

Set oTipo.oParent = Me

oTipo.SQuery = "sm_muestra_tg_operario_tejedor '" & tipo & "','" & Trim(TxtTip_Tejedor.Text) & "','" & Trim(TxtCod_Tejedor.Text) & "','" & Trim(TxtNom_Tejedor.Text) & "','007'"

oTipo.Caption = "Buscar Tejedor"
oTipo.Cargar_Datos

oTipo.gexLista.Columns("Tipo").Width = 600
oTipo.gexLista.Columns("Codigo").Width = 1000
oTipo.gexLista.Columns("nombre").Width = 5000

If oTipo.gexLista.RowCount > 1 Then
    oTipo.Show vbModal
Else
    CODIGO = oTipo.gexLista.Value(oTipo.gexLista.Columns("Tipo").Index)
    Descripcion = oTipo.gexLista.Value(oTipo.gexLista.Columns("codigo").Index)
    TipoAdd = oTipo.gexLista.Value(oTipo.gexLista.Columns("nombre").Index)
End If

If Trim(CODIGO) <> "" Then
    TxtTip_Tejedor.Text = CODIGO
    TxtCod_Tejedor.Text = Descripcion
    TxtNom_Tejedor.Text = TipoAdd
    CODIGO = "": Descripcion = "": TipoAdd = ""
    TxtObservaciones.SetFocus
End If

Unload oTipo
Set oTipo = Nothing
Set rs = Nothing
End Sub

Sub Busca_Auditor(tipo As Integer)
Dim oTipo As New frmBusqGeneral3
Dim rs As New ADODB.Recordset

Set oTipo.oParent = Me

oTipo.SQuery = "sm_muestra_tg_operario_tejedor '" & tipo & "','" & Trim(TxtTip_Auditor.Text) & "','" & Trim(TxtCod_Auditor.Text) & "','" & Trim(TxtNom_Auditor.Text) & "','006'"

oTipo.Caption = "Buscar Auditor"
oTipo.Cargar_Datos

oTipo.gexLista.Columns("Tipo").Width = 600
oTipo.gexLista.Columns("Codigo").Width = 1000
oTipo.gexLista.Columns("nombre").Width = 5000

If oTipo.gexLista.RowCount > 1 Then
    oTipo.Show vbModal
Else
    CODIGO = oTipo.gexLista.Value(oTipo.gexLista.Columns("Tipo").Index)
    Descripcion = oTipo.gexLista.Value(oTipo.gexLista.Columns("codigo").Index)
    TipoAdd = oTipo.gexLista.Value(oTipo.gexLista.Columns("nombre").Index)
End If

If Trim(CODIGO) <> "" Then
    TxtTip_Auditor.Text = CODIGO
    TxtCod_Auditor.Text = Descripcion
    TxtNom_Auditor.Text = TipoAdd
    CODIGO = "": Descripcion = "": TipoAdd = ""
    TxtTurno.SetFocus
End If

Unload oTipo
Set oTipo = Nothing
Set rs = Nothing
End Sub

