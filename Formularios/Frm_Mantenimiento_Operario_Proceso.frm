VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form Frm_Mantenimiento_Operario_Proceso 
   Caption         =   "Mantenimiento Operario Proceso"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12900
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   12900
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrmOperario 
      Height          =   1335
      Left            =   240
      TabIndex        =   9
      Top             =   4440
      Width           =   11415
      Begin VB.CheckBox chkSupervisor 
         Caption         =   "Supervisor"
         Height          =   375
         Left            =   10080
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "CANCELAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4920
         TabIndex        =   14
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "ACEPTAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3480
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtdes_Trabajador 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   4080
         TabIndex        =   12
         Top             =   240
         Width           =   7215
      End
      Begin VB.TextBox txtcod_trabajador 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   2640
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtTip_trabajador 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "TRABAJADOR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "NUEVO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "GUARDAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      TabIndex        =   7
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Frame famCabecera 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12795
      Begin VB.CheckBox chktodos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "TODOS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   11760
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "BUSCAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10440
         TabIndex        =   5
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtDes_Proceso 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3720
         TabIndex        =   2
         Top             =   240
         Width           =   6675
      End
      Begin VB.TextBox txtCod_Proceso 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   1905
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "CODIGO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   60
         TabIndex        =   4
         Top             =   330
         Width           =   615
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "DESCRIPCION"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2760
         TabIndex        =   3
         Top             =   360
         Width           =   1050
      End
   End
   Begin GridEX20.GridEX grxDatos 
      Height          =   5835
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   10292
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GroupByBoxVisible=   0   'False
      HeaderFontName  =   "Verdana"
      HeaderFontBold  =   -1  'True
      HeaderFontSize  =   6.75
      HeaderFontWeight=   700
      FontName        =   "Tahoma"
      ColumnHeaderHeight=   270
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      FormatStylesCount=   9
      FormatStyle(1)  =   "Frm_Mantenimiento_Operario_Proceso.frx":0000
      FormatStyle(2)  =   "Frm_Mantenimiento_Operario_Proceso.frx":0128
      FormatStyle(3)  =   "Frm_Mantenimiento_Operario_Proceso.frx":01D8
      FormatStyle(4)  =   "Frm_Mantenimiento_Operario_Proceso.frx":028C
      FormatStyle(5)  =   "Frm_Mantenimiento_Operario_Proceso.frx":0364
      FormatStyle(6)  =   "Frm_Mantenimiento_Operario_Proceso.frx":041C
      FormatStyle(7)  =   "Frm_Mantenimiento_Operario_Proceso.frx":04FC
      FormatStyle(8)  =   "Frm_Mantenimiento_Operario_Proceso.frx":05A8
      FormatStyle(9)  =   "Frm_Mantenimiento_Operario_Proceso.frx":0674
      ImageCount      =   0
      PrinterProperties=   "Frm_Mantenimiento_Operario_Proceso.frx":0740
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   2760
      Top             =   6600
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "Frm_Mantenimiento_Operario_Proceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public fila_seleccionada As Double
Private StrSQL As String
Public CODIGO As String
Public Descripcion As String

Private Sub CmdAceptar_Click()
Call AltaOperario
Call limpiarcajas
FrmOperario.Visible = False
End Sub
Sub limpiarcajas()
'txtTip_trabajador.Text = ""
txtcod_trabajador.Text = ""
txtdes_Trabajador.Text = ""
chkSupervisor.Value = False

End Sub

Private Sub AltaOperario()
    On Error GoTo Eliminar_DatosErr
    Dim sTipo As String
    Dim sflg_supervisor As String
    Dim sflg_activo As String

    If txtCod_Proceso.Text = "" Then
      MsgBox "Codigo de proceso no valido", vbCritical + vbOKOnly, "Mensaje"
      Exit Sub
    End If
    If txtTip_trabajador.Text = "" Then
      MsgBox "tipo de Trabajador no valido", vbCritical + vbOKOnly, "Mensaje"
      Exit Sub
    End If
    If txtcod_trabajador.Text = "" Then
      MsgBox "codigo de Trabajador no valido", vbCritical + vbOKOnly, "Mensaje"
      Exit Sub
    End If
     
        sTipo = "A"
        sflg_supervisor = "N"
        sflg_activo = "S"
        
        If chkSupervisor.Value = 1 Then
          sflg_supervisor = "S"
        End If
        
        StrSQL = " UP_MANT_TI_TINTORERIA_PROCESOOPERARIO '" & sTipo & _
                                                    "' ,'" & txtCod_Proceso.Text & _
                                                    "', '" & txtTip_trabajador.Text & _
                                                    "', '" & txtcod_trabajador & _
                                                    "', '" & sflg_supervisor & _
                                                    "', '" & sflg_activo & "'"
                    
  
    Call ExecuteSQL(cConnect, StrSQL)
    Call CmdBuscar_Click
    
    Exit Sub

Eliminar_DatosErr:
    ErrorHandler err, "guardar Datos"
End Sub

Private Sub CmdBuscar_Click()
    On Error GoTo Salvar_DatosErr
    Dim TODOS As String
    TODOS = "N"
    
    If chktodos.Value = 1 Then
     TODOS = "S"
    End If
    
    StrSQL = " SM_MUESTRA_OPERARIOS_PROCESOS '" & txtCod_Proceso.Text & "','" & TODOS & "' "
    Set grxDatos.ADORecordset = CargarRecordSetDesconectado(StrSQL, cConnect)
    Call Configurar_Grid
    
   Exit Sub
Salvar_DatosErr:
    ErrorHandler err, "Salvar_Datos"
    
End Sub
Private Sub cmdCancelar_Click()
    Call limpiarcajas
    FrmOperario.Visible = False
End Sub

Private Sub cmdGuardar_Click()
    Call GuardarDatos
End Sub

Private Sub cmdNuevo_Click()
    Call limpiarcajas
    FrmOperario.Visible = True
End Sub

Private Sub Form_Load()

    FrmOperario.Visible = False
    txtTip_trabajador.Text = "O"

End Sub

Sub Busca_Opcion(strCampo1 As String, strCampo2 As String, strTabla As String, txtCod As TextBox, txtDes As TextBox, Opcion As Integer)
On Error GoTo fin
Dim rstAux As ADODB.Recordset
    StrSQL = "Select " & strCampo1 & " AS Cod," & strCampo2 & " as Descripcion from " & strTabla

    txtCod = Trim(txtCod)
    txtDes = Trim(txtDes)
    Select Case Opcion
    Case 1: StrSQL = StrSQL & strCampo1 & " like '%" & txtCod & "%'"
    Case 2: StrSQL = StrSQL & strCampo2 & " like '%" & txtDes & "%'"
    End Select
    txtCod = ""
    txtDes = ""
    fila_seleccionada = 0
    
    With frmBusqGeneral
        Set .oParent = Me
        .SQuery = StrSQL
        .Cargar_Datos
        
        CODIGO = ".."
        Set rstAux = .gexList.ADORecordset
        'If rstAux.RecordCount > 1 Then
        .Show vbModal
        
        If fila_seleccionada > 0 And rstAux.RecordCount > 0 Then
            rstAux.AbsolutePosition = fila_seleccionada
            txtCod = Trim(rstAux!cod)
            txtDes = Trim(rstAux!Descripcion)
            'Select Case Opcion
            'Case 1: SendKeys "{TAB}": SendKeys "{TAB}"
            'Case 2: SendKeys "{TAB}"
            'End Select
        Else
            txtCod = ""
            txtDes = ""
            SendKeys "{TAB}"
        End If
        
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento (" & Opcion & ")"
End Sub


Private Sub txtCod_Proceso_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  
    Call Busca_Opcion("COD_PROCESO_TINTO", "DESCRIPCION", "TI_PROCESOS_TINTORERIA where ", txtCod_Proceso, txtDes_Proceso, 1)
    If Trim(txtCod_Proceso.Text) <> "" Then
       CmdBuscar.SetFocus
    Else
       txtCod_Proceso.SetFocus
    End If
    
  End If
End Sub

Private Sub txtcod_trabajador_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   Call buscaoperario("1")
End If
End Sub

Private Sub txtDes_Proceso_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
  
    Call Busca_Opcion("COD_PROCESO_TINTO", "DESCRIPCION", "TI_PROCESOS_TINTORERIA where ", txtCod_Proceso, txtDes_Proceso, 2)
    If Trim(txtDes_Proceso.Text) <> "" Then
       CmdBuscar.SetFocus
    Else
       txtDes_Proceso.SetFocus
    End If
    
  End If

End Sub
Public Sub buscaoperario(sOpcion As String)
On Error GoTo fin
Dim iCol As Long
Dim rstAux As New ADODB.Recordset
Dim Opcion As String
  StrSQL = "SM_MUESTRA_OPERARIOS '" & sOpcion & "','" & Trim(txtTip_trabajador.Text) & "','" & Trim(txtcod_trabajador.Text) & "','" & Trim(txtdes_Trabajador.Text) & "'"

    With frmBusqGeneralOperario
        Set .oParent = Me
        .SQuery = StrSQL
        .Cargar_Datos
        CODIGO = ".."
        Set rstAux = .DGridLista.ADORecordset
        
        .DGridLista.Columns("tipo").Caption = "Codigo"
        .DGridLista.Columns("tipo").Width = 900
        .DGridLista.Columns("Codigo").Caption = "Codigo"
        .DGridLista.Columns("Codigo").Width = 900
        .DGridLista.Columns("nombre").Caption = "Nombre"
        .DGridLista.Columns("nombre").Width = 1500
        
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If CODIGO <> "" And rstAux.RecordCount > 0 Then
            txtTip_trabajador.Text = Trim(rstAux!Tipo)
            txtcod_trabajador.Text = Trim(rstAux!CODIGO)
            txtdes_Trabajador.Text = Trim(rstAux!Nombre)
        End If
    End With
    Unload frmBusqGeneralOperario
    Set frmBusqGeneralOperario = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
fin:
On Error Resume Next
    Unload frmBusqGeneralOperario
    Set frmBusqGeneralOperario = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Operario(" & Opcion & ")"
End Sub
Private Sub GuardarDatos()
    On Error GoTo Eliminar_DatosErr
    Dim rstAux As New ADODB.Recordset
    Dim sTipo As String
    Dim sflg_supervisor As String
    Dim sflg_activo As String
    grxDatos.Update
    Set rstAux = grxDatos.ADORecordset
     sTipo = "U"
    rstAux.Update
    rstAux.MoveFirst
    Do While Not rstAux.EOF
        
        sflg_supervisor = "N"
        sflg_activo = "N"
        
        If rstAux!flg_supervisor = True Then
          sflg_supervisor = "S"
        End If
        If rstAux!flg_activo = True Then
          sflg_activo = "S"
        End If
       
        StrSQL = " UP_MANT_TI_TINTORERIA_PROCESOOPERARIO '" & sTipo & _
                                                    "' ,'" & rstAux!cod_proceso_tinto & _
                                                    "', '" & rstAux!tip_trabajador & _
                                                    "', '" & rstAux!cod_operario & _
                                                    "', '" & sflg_supervisor & _
                                                    "','" & sflg_activo & "'"
                    
        Call ExecuteSQL(cConnect, StrSQL)
    
        rstAux.MoveNext
    Loop
    
    Call CmdBuscar_Click
    'Call buscardatos
    Exit Sub

Eliminar_DatosErr:
    ErrorHandler err, "guardar Datos"
End Sub
Private Sub txtdes_Trabajador_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
      Call buscaoperario("2")
    End If

End Sub

Public Sub Configurar_Grid()
    Dim C As Integer
    With grxDatos
    
        For C = 1 To .Columns.Count
            With .Columns(C)
                .Caption = UCase(.Caption)
                .HeaderAlignment = jgexAlignCenter
                .TextAlignment = jgexAlignCenter
            End With
        Next C

        With .Columns("COD_PROCESO_TINTO")
            .Width = 800
            .TextAlignment = jgexAlignLeft
            .Caption = "COD"
        End With

        With .Columns("DES_PROCESO")
            .Width = 1500
            .TextAlignment = jgexAlignLeft
            .Visible = True
            .Caption = "PROCESO"
        End With
        With .Columns("TIP_TRABAJADOR")
            .Width = 800
            .TextAlignment = jgexAlignLeft
            .Caption = "TIPO"
            .Visible = True
        End With
         'COD_TELA
        With .Columns("COD_OPERARIO")
            .Width = 1500
            .TextAlignment = jgexAlignLeft
            .Caption = "COD. OPERARIO"
            .Visible = True
        End With

        With .Columns("DESCRIPCION")
            .Width = 3000
            .TextAlignment = jgexAlignLeft
            .Caption = "DES. OPERARIO"
            .Visible = True
        End With
        
        With .Columns("FLG_SUPERVISOR")
            .Width = 1500
            .TextAlignment = jgexAlignLeft
            .Caption = "SUPERVISOR"
            .Visible = True
        End With
        
        With .Columns("FLG_ACTIVO")
            .Width = 800
            .TextAlignment = jgexAlignLeft
            .Caption = "ACTIVO"
            .Visible = True
        End With

    End With
    Call setcolorcolumnas
End Sub

Private Sub setcolorcolumnas()

    grxDatos.Columns("FLG_SUPERVISOR").CellStyle = "cod1"
    grxDatos.Columns("FLG_ACTIVO").CellStyle = "cod2"
        
End Sub

Private Sub grxDatos_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
  Select Case ColIndex
    Case Is = grxDatos.Columns("flg_supervisor").Index
      Cancel = False
    Case Is = grxDatos.Columns("flg_activo").Index
      Cancel = False
    Case Else
      Cancel = True
  End Select
End Sub

