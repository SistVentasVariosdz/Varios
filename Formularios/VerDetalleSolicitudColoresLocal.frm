VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form VerDetalleSolicitudColoresLocal 
   Caption         =   "Detalle de Solicitud"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   11835
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   700
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11775
      Begin VB.TextBox TxtDescripcion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   5175
      End
      Begin VB.TextBox TxtCorr_Carta 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Corr. Carta"
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
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   930
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   11655
      Begin GridEX20.GridEX GridEX1 
         Height          =   6345
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   11192
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "VerDetalleSolicitudColoresLocal.frx":0000
         Column(2)       =   "VerDetalleSolicitudColoresLocal.frx":00C8
         FormatStylesCount=   6
         FormatStyle(1)  =   "VerDetalleSolicitudColoresLocal.frx":016C
         FormatStyle(2)  =   "VerDetalleSolicitudColoresLocal.frx":02A4
         FormatStyle(3)  =   "VerDetalleSolicitudColoresLocal.frx":0354
         FormatStyle(4)  =   "VerDetalleSolicitudColoresLocal.frx":0408
         FormatStyle(5)  =   "VerDetalleSolicitudColoresLocal.frx":04E0
         FormatStyle(6)  =   "VerDetalleSolicitudColoresLocal.frx":0598
         ImageCount      =   0
         PrinterProperties=   "VerDetalleSolicitudColoresLocal.frx":0678
      End
   End
   Begin VB.Frame FraReLab 
      Caption         =   "Comentario Re - Lab"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4080
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox TxtReLab 
         Height          =   615
         Left            =   360
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   240
         Width           =   3975
      End
      Begin FunctionsButtons.FunctButt FunctButt3 
         Height          =   510
         Left            =   1200
         TabIndex        =   5
         Top             =   960
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"VerDetalleSolicitudColoresLocal.frx":0850
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
   End
   Begin VB.Frame FraAprobacion 
      Caption         =   "Comentario Aprobación Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4080
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox TxtComentario 
         Height          =   615
         Left            =   360
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   240
         Width           =   3975
      End
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   510
         Left            =   1200
         TabIndex        =   2
         Top             =   960
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   $"VerDetalleSolicitudColoresLocal.frx":08E6
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2880
      TabIndex        =   12
      Top             =   7560
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   900
      Custom          =   $"VerDetalleSolicitudColoresLocal.frx":097C
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "VerDetalleSolicitudColoresLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim StrSQL As String
Dim i As Integer

Public SCLIENTE As String
Public sTemporada As String
Public Add As Integer
Dim Mensaje As Variant

Sub CARGA_GRID()
StrSQL = "es_muestra_solicitudes_desarrollo_detalle_Local '" & TxtCorr_Carta.Text & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(StrSQL, cConnect)

GridEX1.Columns("sec").Width = 500
GridEX1.Columns("descripcion_color").Width = 2200
GridEX1.Columns("descripcion_fibra").Width = 2500
GridEX1.Columns("Fec_Asignacion").Width = 1450
GridEX1.Columns("COD_COLOR").Width = 900
GridEX1.Columns("nombre").Width = 1500
GridEX1.Columns("pc").Width = 0
GridEX1.Columns("cod_usuario").Width = 0
GridEX1.Columns("codigo_color_cliente").Width = 1800

GridEX1.Columns("nombre").Caption = "Nombre Color Tintoreria"
GridEX1.Columns("descripcion_color").Caption = "Descripcion Color"
GridEX1.Columns("descripcion_fibra").Caption = "Descripcion Fibra"
GridEX1.Columns("Fec_Asignacion").Caption = "Fec. Asignac."
GridEX1.Columns("Status").Width = 1400

GridEX1.Row = i
GridEX1.FrozenColumns = 3
End Sub

Private Sub Form_Load()
Dim sSeguridad  As String
sSeguridad = get_botones1(Me, vper, vemp, Me.Name)

'Me.FunctButt1.FunctionsUser = sSeguridad
End Sub


Public Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ADICIONAR"
    Load FrmAddDetSolicitudDesaColoresLocal
    FrmAddDetSolicitudDesaColoresLocal.sAccion = "I"
    FrmAddDetSolicitudDesaColoresLocal.TxtCorr_Carta.Text = TxtCorr_Carta.Text
    FrmAddDetSolicitudDesaColoresLocal.TxtDescripcion.Text = TxtDescripcion.Text
    i = GridEX1.Row
    FrmAddDetSolicitudDesaColoresLocal.Show 1
    If FrmAddDetSolicitudDesaColoresLocal.vOk = True Then
        CARGA_GRID
        GridEX1.Row = GridEX1.RowCount
        Call FunctButt1_ActionClick(0, 0, "ADICIONAR")
    End If
    Set FrmAddDetSolicitudDesaColoresLocal = Nothing
    
Case "MODIFICAR"
    If GridEX1.RowCount = 0 Then Exit Sub
    Load FrmAddDetSolicitudDesaColoresLocal
    
    FrmAddDetSolicitudDesaColoresLocal.sAccion = "U"
    FrmAddDetSolicitudDesaColoresLocal.TxtCorr_Carta.Text = TxtCorr_Carta.Text
    FrmAddDetSolicitudDesaColoresLocal.TxtSec = GridEX1.Value(GridEX1.Columns("Sec").Index)
    FrmAddDetSolicitudDesaColoresLocal.TxtDescripcion.Text = TxtDescripcion.Text
    FrmAddDetSolicitudDesaColoresLocal.TxtDes_Color.Text = GridEX1.Value(GridEX1.Columns("Descripcion_Color").Index)
    FrmAddDetSolicitudDesaColoresLocal.TxtDes_fibra.Text = GridEX1.Value(GridEX1.Columns("Descripcion_Fibra").Index)
    FrmAddDetSolicitudDesaColoresLocal.TxtCod_ColCli.Text = GridEX1.Value(GridEX1.Columns("Codigo_Color_Cliente").Index)
    FrmAddDetSolicitudDesaColoresLocal.TxtMat_Prima_Entregada.Text = GridEX1.Value(GridEX1.Columns("Mat_Prima_Entregada").Index)
    'FrmAddDetSolicitudDesaColoresLocal.TxtComentario.Text = GridEX1.Value(GridEX1.Columns("Comentario_ReLab").Index)
    i = GridEX1.Row
    FrmAddDetSolicitudDesaColores.Show 1
    Set FrmAddDetSolicitudDesaColores = Nothing
    
    CARGA_GRID
Case "ELIMINAR"
    If GridEX1.RowCount = 0 Then Exit Sub
    Load FrmAddDetSolicitudDesaColoresLocal
    FrmAddDetSolicitudDesaColoresLocal.sAccion = "D"
    FrmAddDetSolicitudDesaColoresLocal.TxtCorr_Carta.Text = TxtCorr_Carta.Text
    FrmAddDetSolicitudDesaColoresLocal.TxtSec = GridEX1.Value(GridEX1.Columns("Sec").Index)
    FrmAddDetSolicitudDesaColoresLocal.TxtDescripcion.Text = TxtDescripcion.Text
    FrmAddDetSolicitudDesaColoresLocal.TxtDes_Color.Text = GridEX1.Value(GridEX1.Columns("Descripcion_Color").Index)
    FrmAddDetSolicitudDesaColoresLocal.TxtDes_fibra.Text = GridEX1.Value(GridEX1.Columns("Descripcion_Fibra").Index)
    FrmAddDetSolicitudDesaColoresLocal.TxtCod_ColCli.Text = GridEX1.Value(GridEX1.Columns("Codigo_Color_Cliente").Index)
    FrmAddDetSolicitudDesaColoresLocal.TxtMat_Prima_Entregada.Text = GridEX1.Value(GridEX1.Columns("Mat_Prima_Entregada").Index)
    'FrmAddDetSolicitudDesaColoresLocal.TxtComentario.Text = GridEX1.Value(GridEX1.Columns("Comentario_ReLab").Index)
    FrmAddDetSolicitudDesaColoresLocal.FraDatos.Enabled = False
    i = GridEX1.Row
    FrmAddDetSolicitudDesaColoresLocal.Show 1
    Set FrmAddDetSolicitudDesaColoresLocal = Nothing
    CARGA_GRID
Case "IMPRIMIR"
    If GridEX1.RowCount = 0 Then Exit Sub
    Call Reporte
Case "ESTADO"
    If GridEX1.RowCount = 0 Then Exit Sub
    i = GridEX1.Row
    Call Cambia_Estado
    CARGA_GRID
Case "RELAB"
    If GridEX1.RowCount = 0 Then Exit Sub
    i = GridEX1.Row
    FraReLab.Visible = True
    TxtReLab.SetFocus
'Case "APROBAR"
'    If GridEX1.RowCount = 0 Then Exit Sub
'    i = GridEX1.Row
'    FraAprobacion.Visible = True
'    TxtComentario.SetFocus
Case "SALIR"
    Unload Me
End Select
End Sub

Sub Reporte()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Cadena
    Cadena = "es_muestra_solicitudes_desarrollo_detalle_Local '" & TxtCorr_Carta.Text & "'"
    
    Ruta = vRuta & "\RptSolDesaColores_Detalle_Local.xlt"
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.run "Reporte", Val(TxtCorr_Carta.Text), Trim(TxtDescripcion.Text), SCLIENTE, sTemporada, Cadena, cConnect
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler err, "GeneraReportes"
    Set oo = Nothing
End Sub

Sub Re_Lab()
On Error GoTo errRe_Lab
StrSQL = "es_envia_color_Local_a_re_lab " & Val(TxtCorr_Carta.Text) & "," & Val(GridEX1.Value(GridEX1.Columns("Sec").Index)) & ",'" & vusu & "','" & ComputerName & "','" & _
            Trim(TxtReLab.Text) & "'"
ExecuteSQL cConnect, StrSQL
MsgBox "Operación realizada con éxito", vbInformation, "Envio Re-Lab"
Exit Sub
errRe_Lab:
MsgBox "no se puede continuar", vbInformation, "Mensaje"
    
End Sub

Sub Aprobar_color()
On Error GoTo errRe_Lab
StrSQL = "es_Up_aprueba_Color " & Val(TxtCorr_Carta.Text) & "," & Val(GridEX1.Value(GridEX1.Columns("Sec").Index)) & ",'" & Trim(TxtComentario.Text) & "'"
ExecuteSQL cConnect, StrSQL
MsgBox "Operación realizada con éxito", vbInformation, "Envio Re-Lab"
Exit Sub
errRe_Lab:
    MsgBox "no se puede continuar", vbInformation, "Mensaje"
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo fin
    Select Case ActionName
    Case "ACEPTAR"
        StrSQL = "es_Cambia_Status_Color_Local " & Val(TxtCorr_Carta) & "," & _
                                    Val(GridEX1.Value(GridEX1.Columns("Sec").Index)) & ",'" & _
                                    vusu & "','" & ComputerName & "','" & TxtComentario.Text & "'"
                                    
        ExecuteSQL cConnect, StrSQL
        MsgBox "Cambio de estado realizado", vbInformation
        FraAprobacion.Visible = False
        CARGA_GRID
    Case "CANCELAR"
        FraAprobacion.Visible = False
    End Select
    TxtComentario.Text = ""
    GridEX1.Row = i

Exit Sub
fin:
MsgBox "No se puede Continuar", vbExclamation, "Mensaje"

End Sub

Private Sub FunctButt3_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    Call Re_Lab
    FraReLab.Visible = False
    CARGA_GRID
Case "CANCELAR"
    FraReLab.Visible = False
End Select
End Sub

Private Sub TxtComentario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub

Sub Cambia_Estado()
Dim Mensaje As Variant
On Error GoTo errCambioEstado
Mensaje = MsgBox("¿Esta seguro de cambiar estado?", vbYesNo)
If Mensaje = vbNo Then Exit Sub

If Mid(GridEX1.Value(GridEX1.Columns("Status").Index), 1, 1) = "D" Then
    Me.FraAprobacion.Visible = True
Else
    StrSQL = "es_Cambia_Status_Color_Local " & Val(TxtCorr_Carta) & "," & _
                                Val(GridEX1.Value(GridEX1.Columns("Sec").Index)) & ",'" & _
                                vusu & "','" & ComputerName & "','" & TxtComentario.Text & "'"
                                
    ExecuteSQL cConnect, StrSQL
    MsgBox "Cambio de estado realizado", vbInformation
End If
Exit Sub
errCambioEstado:
    MsgBox "Inconvenientes para en el cambio de estado", vbInformation + vbOKOnly, "Mensaje"
    
End Sub

