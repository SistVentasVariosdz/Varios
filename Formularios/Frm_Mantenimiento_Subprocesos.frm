VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "mantenimientos.ocx"
Begin VB.Form Frm_Mantenimiento_Subprocesos 
   Caption         =   "MANTENIMIENTO DE SUBPROCESOS"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12690
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   12690
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      Caption         =   "Datos"
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   0
      TabIndex        =   5
      Top             =   6000
      Width           =   12555
      Begin VB.TextBox txtCod_proceso_tinto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1710
         TabIndex        =   9
         Top             =   270
         Width           =   1665
      End
      Begin VB.TextBox txtSub_secuencia 
         Height          =   285
         Left            =   1710
         TabIndex        =   8
         Top             =   570
         Width           =   1665
      End
      Begin VB.TextBox txtSub_proceso_Descripcion 
         Height          =   285
         Left            =   4800
         TabIndex        =   7
         Top             =   570
         Width           =   7695
      End
      Begin VB.TextBox txtDes_proceso_tinto 
         Height          =   285
         Left            =   4800
         TabIndex        =   6
         Top             =   270
         Width           =   7665
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
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
         Left            =   120
         TabIndex        =   13
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "SUB SECUENCIA"
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
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1200
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
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
         Left            =   3480
         TabIndex        =   11
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "SUB PROCESOS"
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
         Left            =   3480
         TabIndex        =   10
         Top             =   600
         Width           =   1155
      End
   End
   Begin VB.Frame famCabecera 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12555
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
         Left            =   11160
         TabIndex        =   26
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtDes_proceso_tinto_bus 
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
         Left            =   3480
         TabIndex        =   2
         Top             =   240
         Width           =   7515
      End
      Begin VB.TextBox txtCod_proceso_tinto_bus 
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
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   1785
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "COD"
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
         Width           =   315
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "DES PROCESO"
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
         Left            =   2400
         TabIndex        =   3
         Top             =   360
         Width           =   1050
      End
   End
   Begin GridEX20.GridEX grxDatos 
      Height          =   5355
      Left            =   0
      TabIndex        =   14
      Top             =   660
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   9446
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      AllowEdit       =   0   'False
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
      FormatStyle(1)  =   "Frm_Mantenimiento_Subprocesos.frx":0000
      FormatStyle(2)  =   "Frm_Mantenimiento_Subprocesos.frx":0128
      FormatStyle(3)  =   "Frm_Mantenimiento_Subprocesos.frx":01D8
      FormatStyle(4)  =   "Frm_Mantenimiento_Subprocesos.frx":028C
      FormatStyle(5)  =   "Frm_Mantenimiento_Subprocesos.frx":0364
      FormatStyle(6)  =   "Frm_Mantenimiento_Subprocesos.frx":041C
      FormatStyle(7)  =   "Frm_Mantenimiento_Subprocesos.frx":04FC
      FormatStyle(8)  =   "Frm_Mantenimiento_Subprocesos.frx":05A8
      FormatStyle(9)  =   "Frm_Mantenimiento_Subprocesos.frx":0674
      ImageCount      =   0
      PrinterProperties=   "Frm_Mantenimiento_Subprocesos.frx":071C
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   630
      Left            =   9000
      TabIndex        =   15
      Top             =   6960
      Width           =   3630
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"Frm_Mantenimiento_Subprocesos.frx":08EC
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1305
         Left            =   -1260
         TabIndex        =   16
         Top             =   -2970
         Width           =   6975
         Begin VB.TextBox Txtcod_Fabrica 
            Height          =   285
            Left            =   1440
            MaxLength       =   3
            TabIndex        =   20
            Top             =   240
            Width           =   795
         End
         Begin VB.TextBox TxtNom_Fabrica 
            Height          =   285
            Left            =   2250
            TabIndex        =   19
            Top             =   255
            Width           =   4635
         End
         Begin VB.TextBox Txt_Planilla 
            Height          =   285
            Left            =   2250
            TabIndex        =   18
            Top             =   570
            Width           =   4635
         End
         Begin VB.TextBox Txt_Tipo 
            Height          =   285
            Left            =   1440
            TabIndex        =   17
            Top             =   570
            Width           =   795
         End
         Begin VB.Label Label2 
            Caption         =   "Inicio:"
            Height          =   210
            Left            =   840
            TabIndex        =   24
            Top             =   960
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fabrica:"
            Height          =   195
            Left            =   690
            TabIndex        =   23
            Top             =   270
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Trabajador:"
            Height          =   195
            Left            =   90
            TabIndex        =   22
            Top             =   570
            Width           =   1170
         End
         Begin VB.Label Label4 
            Caption         =   "Fin:"
            Height          =   195
            Left            =   3990
            TabIndex        =   21
            Top             =   990
            Width           =   435
         End
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   780
      Top             =   8610
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label Label18 
      Caption         =   "Label18"
      Height          =   255
      Left            =   2040
      TabIndex        =   25
      Top             =   8880
      Width           =   1935
   End
End
Attribute VB_Name = "Frm_Mantenimiento_Subprocesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CODIGO As String
'Public Descripcion As String
Private StrSQL As String
Private sTipo As String
Public fila_seleccionada As Double

Private Sub Form_Load()
  Call HABILITA_DATOS("N")
  HabilitaMant Me.MantFunc1, "ADICIONAR"
  'Label18.Caption = Get_User_Name
End Sub

Private Sub CmdBuscar_Click()
    Call buscardatos
End Sub

Private Sub grxDatos_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    If grxDatos.RowCount > 0 Then Call CARGA_DATOS
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Eliminar As Integer
    
    Select Case ActionName
        Case "ADICIONAR"
            sTipo = "I"
            LIMPIAR_DATOS
            HABILITA_DATOS ("N")
            txtSub_proceso_Descripcion.SetFocus
            
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            
        Case "MODIFICAR"
        
            sTipo = "U"
            HABILITA_DATOS ("M")
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            
        Case "ELIMINAR"
            If VALIDA_DATOS Then
                If MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?.", vbInformation + vbYesNo, "Mensaje") = vbNo Then Exit Sub
                Call ELIMINAR_DATOS
                Call LIMPIAR_DATOS
                Call buscardatos
            End If
            
        Case "GRABAR"
            If VALIDA_DATOS Then
                Call SALVAR_DATOS
                Call buscardatos
            End If
            
        Case "DESHACER"
            Call LIMPIAR_DATOS
            Call buscardatos
            
        Case "SALIR"
            sTipo = "V"
            Unload Me
    End Select
End Sub

Public Sub HABILITA_DATOS(Tipo As String)
    Select Case Tipo
    
        Case "N" 'nuevo
        
            txtCod_proceso_tinto.Enabled = False
            txtDes_proceso_tinto.Enabled = False
            txtSub_secuencia.Enabled = False
            txtSub_proceso_Descripcion.Enabled = True
            grxDatos.Enabled = False
        
        Case "M" 'cuando modifica
            
            txtCod_proceso_tinto.Enabled = False
            txtDes_proceso_tinto.Enabled = False
            txtSub_secuencia.Enabled = False
            txtSub_proceso_Descripcion.Enabled = True
            grxDatos.Enabled = False


        Case "D" 'cuando deshace
            
            txtCod_proceso_tinto.Enabled = False
            txtDes_proceso_tinto.Enabled = False
            txtSub_secuencia.Enabled = False
            txtSub_proceso_Descripcion.Enabled = False
            grxDatos.Enabled = True
    
    End Select
End Sub

Public Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True

    If Trim(txtSub_proceso_Descripcion.Text) = "" Then
        VALIDA_DATOS = False
        Call MsgBox("sirvase ingresar el nombre del proceso", vbInformation, "Mensaje")
        txtSub_proceso_Descripcion.SetFocus
        Exit Function
    End If
    
End Function

Public Sub LIMPIAR_DATOS()
    
    'txtCod_proceso_tinto.Text = ""
    'txtDes_proceso_tinto.Text = ""
    'txtSub_secuencia.Text = ""
    txtSub_proceso_Descripcion.Text = ""

End Sub
Private Sub buscardatos()
    On Error GoTo Salvar_DatosErr

    StrSQL = " UP_MAN_TI_SUBPROCESOS_TINTORERIA 'V','" & Trim(txtCod_proceso_tinto_bus.Text) & "','',''"
    Set grxDatos.ADORecordset = CargarRecordSetDesconectado(StrSQL, cConnect)
        
    Call CARGA_DATOS
    Call HABILITA_DATOS("D")
    Call Configurar_Grid
    
    If grxDatos.RowCount > 0 Then
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
    Else
        HabilitaMant Me.MantFunc1, "ADICIONAR"
    End If
    Exit Sub
    
Salvar_DatosErr:
    ErrorHandler err, "Salvar_Datos"
End Sub

Public Sub CARGA_DATOS()

        txtCod_proceso_tinto.Text = grxDatos.Value(grxDatos.Columns("COD_PROCESO_TINTO").Index)
        txtDes_proceso_tinto.Text = grxDatos.Value(grxDatos.Columns("DESCRIPCION_PROCESO").Index)
        txtSub_secuencia.Text = grxDatos.Value(grxDatos.Columns("SUBSECUENCIA").Index)
        txtSub_proceso_Descripcion.Text = grxDatos.Value(grxDatos.Columns("DES_SUBPROCESO").Index)
        
End Sub
Private Sub SALVAR_DATOS()
    On Error GoTo Salvar_DatosErr
    
    StrSQL = " UP_MAN_TI_SUBPROCESOS_TINTORERIA '" & sTipo & _
                                                "' ,'" & Trim(txtCod_proceso_tinto.Text) & _
                                                "', '" & Trim(txtSub_secuencia.Text) & _
                                                "','" & Trim(txtSub_proceso_Descripcion.Text) & "'"
                
    Call ExecuteSQL(cConnect, StrSQL)
    Call LIMPIAR_DATOS
    Call buscardatos
    
    Exit Sub
Salvar_DatosErr:
    ErrorHandler err, "Salvar_Datos"
End Sub

Private Sub ELIMINAR_DATOS()
    On Error GoTo Eliminar_DatosErr
    sTipo = "D"
    StrSQL = " UP_MAN_TI_SUBPROCESOS_TINTORERIA '" & sTipo & _
                                                "' ,'" & Trim(txtCod_proceso_tinto.Text) & _
                                                "', '" & Trim(txtSub_secuencia.Text) & _
                                                "','" & Trim(txtSub_proceso_Descripcion.Text) & "'"
                
    Call ExecuteSQL(cConnect, StrSQL)
    Call LIMPIAR_DATOS
    Call buscardatos
    
    Exit Sub
    
Eliminar_DatosErr:
    ErrorHandler err, "Eliminar_Datos"
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

        txtCod_proceso_tinto.Text = grxDatos.Value(grxDatos.Columns("COD_PROCESO_TINTO").Index)
        txtDes_proceso_tinto.Text = grxDatos.Value(grxDatos.Columns("DESCRIPCION_PROCESO").Index)
        txtSub_secuencia.Text = grxDatos.Value(grxDatos.Columns("SUBSECUENCIA").Index)
        txtSub_proceso_Descripcion.Text = grxDatos.Value(grxDatos.Columns("DES_SUBPROCESO").Index)
        
        With .Columns("COD_PROCESO_TINTO")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "COD"
        End With

        With .Columns("DESCRIPCION_PROCESO")
            .Width = 2500
            .TextAlignment = jgexAlignLeft
            .Visible = True
            .Caption = "DESCRIPCION_PROCESO"
        End With
        
        With .Columns("SUBSECUENCIA")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "SUBSECUENCIA"
            .Visible = True
        End With
         'COD_TELA
        With .Columns("DES_SUBPROCESO")
            .Width = 2000
            .TextAlignment = jgexAlignLeft
            .Caption = "DES_SUBPROCESO"
            .Visible = True
        End With
        
    End With
    Call setcolorcolumnas
End Sub
Private Sub setcolorcolumnas()
    grxDatos.Columns("DES_SUBPROCESO").CellStyle = "cod1"
End Sub
''''busqueda

'Private Sub txtBus_Num_Ruc_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then
'     txtBus_Nom_Cliente.Text = ""
'     Call buscardatos
'   End If
'End Sub

Private Sub txtDes_proceso_tinto_bus_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        Call Busca_Opcion("COD_PROCESO_TINTO", "DESCRIPCION", "TI_PROCESOS_TINTORERIA where ", txtCod_proceso_tinto_bus, txtDes_proceso_tinto_bus, 2)
        If Trim(txtDes_proceso_tinto_bus.Text) <> "" Then
           CmdBuscar.SetFocus
        Else
           txtDes_proceso_tinto_bus.SetFocus
        End If
  End If
End Sub

Private Sub txtCod_proceso_tinto_bus_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
  
    Call Busca_Opcion("COD_PROCESO_TINTO", "DESCRIPCION", "TI_PROCESOS_TINTORERIA where ", txtCod_proceso_tinto_bus, txtDes_proceso_tinto_bus, 1)
    If Trim(txtCod_proceso_tinto_bus.Text) <> "" Then
       CmdBuscar.SetFocus
    Else
       txtCod_proceso_tinto_bus.SetFocus
    End If
    
  End If
  
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


