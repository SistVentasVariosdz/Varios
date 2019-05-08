VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Begin VB.Form frmMatTg_Cliente 
   Caption         =   "Mantenimiento de Clientes"
   ClientHeight    =   9255
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13890
   LinkTopic       =   "Form1"
   ScaleHeight     =   9255
   ScaleWidth      =   13890
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame famCabecera 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   12555
      Begin VB.TextBox txtBus_Num_Ruc 
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
         TabIndex        =   22
         Top             =   240
         Width           =   2385
      End
      Begin VB.TextBox txtBus_Nom_Cliente 
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
         TabIndex        =   21
         Top             =   240
         Width           =   7395
      End
      Begin VB.Label Label17 
         BackColor       =   &H0080FFFF&
         Caption         =   "<ENTER> para buscar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   11520
         TabIndex        =   39
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "CLIENTE"
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
         Left            =   3000
         TabIndex        =   34
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "RUC"
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
         TabIndex        =   23
         Top             =   330
         Width           =   300
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      Caption         =   "Datos"
      ForeColor       =   &H80000008&
      Height          =   2625
      Left            =   0
      TabIndex        =   0
      Top             =   6000
      Width           =   12555
      Begin VB.TextBox txtNum_ruc 
         Height          =   285
         Left            =   4680
         TabIndex        =   2
         Top             =   270
         Width           =   2985
      End
      Begin VB.TextBox txtDireccion 
         Height          =   315
         Left            =   1710
         TabIndex        =   5
         Top             =   840
         Width           =   10815
      End
      Begin VB.TextBox txtDEs_cliente 
         Height          =   285
         Left            =   8160
         TabIndex        =   4
         Top             =   570
         Width           =   4335
      End
      Begin VB.TextBox txtNom_cliente 
         Height          =   285
         Left            =   1710
         TabIndex        =   3
         Top             =   570
         Width           =   5265
      End
      Begin VB.TextBox txtLugar_Entrega 
         Height          =   315
         Left            =   1710
         TabIndex        =   6
         Top             =   1200
         Width           =   10815
      End
      Begin VB.TextBox txtTelefono1 
         Height          =   315
         Left            =   1710
         TabIndex        =   7
         Top             =   1530
         Width           =   4545
      End
      Begin VB.TextBox txtTelefono2 
         Height          =   315
         Left            =   7320
         TabIndex        =   8
         Top             =   1530
         Width           =   5175
      End
      Begin VB.TextBox txtUrl 
         Height          =   315
         Left            =   1710
         TabIndex        =   13
         Top             =   2190
         Width           =   10815
      End
      Begin VB.TextBox txtCod_cliente 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1710
         TabIndex        =   1
         Top             =   270
         Width           =   2265
      End
      Begin VB.TextBox txtEmail 
         Height          =   315
         Left            =   7320
         TabIndex        =   11
         Top             =   1860
         Width           =   5175
      End
      Begin VB.TextBox txtFax 
         Height          =   315
         Left            =   1710
         TabIndex        =   9
         Top             =   1860
         Width           =   4545
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "EMAIL"
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
         Left            =   6840
         TabIndex        =   38
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "TELEFONO2"
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
         Left            =   6360
         TabIndex        =   37
         Top             =   1560
         Width           =   885
      End
      Begin VB.Label Label13 
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
         Left            =   7080
         TabIndex        =   36
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "RUC"
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
         Left            =   4200
         TabIndex        =   35
         Top             =   360
         Width           =   300
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DIRECCION"
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
         TabIndex        =   19
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "NOMBRE"
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
         TabIndex        =   18
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "LUGAR ENTREGA"
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
         TabIndex        =   17
         Top             =   1230
         Width           =   1275
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "URL"
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
         Left            =   150
         TabIndex        =   16
         Top             =   2280
         Width           =   300
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "TELEFONO1"
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
         Left            =   150
         TabIndex        =   15
         Top             =   1590
         Width           =   885
      End
      Begin VB.Label CODIGO 
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
         TabIndex        =   12
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "FAX"
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
         Left            =   150
         TabIndex        =   10
         Top             =   1950
         Width           =   315
      End
   End
   Begin GridEX20.GridEX grxDatos 
      Height          =   5355
      Left            =   0
      TabIndex        =   24
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
      FormatStyle(1)  =   "FrmMantTg_Cliente.frx":0000
      FormatStyle(2)  =   "FrmMantTg_Cliente.frx":0128
      FormatStyle(3)  =   "FrmMantTg_Cliente.frx":01D8
      FormatStyle(4)  =   "FrmMantTg_Cliente.frx":028C
      FormatStyle(5)  =   "FrmMantTg_Cliente.frx":0364
      FormatStyle(6)  =   "FrmMantTg_Cliente.frx":041C
      FormatStyle(7)  =   "FrmMantTg_Cliente.frx":04FC
      FormatStyle(8)  =   "FrmMantTg_Cliente.frx":05A8
      FormatStyle(9)  =   "FrmMantTg_Cliente.frx":0674
      ImageCount      =   0
      PrinterProperties=   "FrmMantTg_Cliente.frx":071C
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   630
      Left            =   8880
      TabIndex        =   14
      Top             =   8640
      Width           =   3630
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"FrmMantTg_Cliente.frx":08EC
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1305
         Left            =   -1260
         TabIndex        =   25
         Top             =   -2970
         Width           =   6975
         Begin VB.TextBox Txt_Tipo 
            Height          =   285
            Left            =   1440
            TabIndex        =   29
            Top             =   570
            Width           =   795
         End
         Begin VB.TextBox Txt_Planilla 
            Height          =   285
            Left            =   2250
            TabIndex        =   28
            Top             =   570
            Width           =   4635
         End
         Begin VB.TextBox TxtNom_Fabrica 
            Height          =   285
            Left            =   2250
            TabIndex        =   27
            Top             =   255
            Width           =   4635
         End
         Begin VB.TextBox Txtcod_Fabrica 
            Height          =   285
            Left            =   1440
            MaxLength       =   3
            TabIndex        =   26
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label4 
            Caption         =   "Fin:"
            Height          =   195
            Left            =   3990
            TabIndex        =   33
            Top             =   990
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Trabajador:"
            Height          =   195
            Left            =   90
            TabIndex        =   32
            Top             =   570
            Width           =   1170
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fabrica:"
            Height          =   195
            Left            =   690
            TabIndex        =   31
            Top             =   270
            Width           =   570
         End
         Begin VB.Label Label2 
            Caption         =   "Inicio:"
            Height          =   210
            Left            =   840
            TabIndex        =   30
            Top             =   960
            Width           =   420
         End
      End
   End
   Begin VB.Label Label18 
      Caption         =   "Label18"
      Height          =   255
      Left            =   2040
      TabIndex        =   40
      Top             =   8880
      Width           =   1935
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   780
      Top             =   8610
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmMatTg_Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private CODIGO As String
'Public Descripcion As String
Private StrSQL As String
Private sTipo As String

Private Sub Form_Load()
  Call HABILITA_DATOS("N")
  HabilitaMant Me.MantFunc1, "ADICIONAR"
  'Label18.Caption = Get_User_Name
End Sub

Private Sub cmdBuscar_Click()
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
            txtNum_ruc.SetFocus
            
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
        
            txtCod_cliente.Enabled = False
            txtNom_cliente.Enabled = True
            txtDEs_cliente.Enabled = True
            txtDireccion.Enabled = True
            txtLugar_Entrega.Enabled = True
            txtEmail.Enabled = True
            txtTelefono1.Enabled = True
            txtTelefono2.Enabled = True
            txtFax.Enabled = True
            txtUrl.Enabled = True
            txtNum_ruc.Enabled = True
            grxDatos.Enabled = False
        
        Case "M" 'cuando modifica
            
            txtCod_cliente.Enabled = False
            txtNom_cliente.Enabled = True
            txtDEs_cliente.Enabled = True
            txtDireccion.Enabled = True
            txtLugar_Entrega.Enabled = True
            txtEmail.Enabled = True
            txtTelefono1.Enabled = True
            txtTelefono2.Enabled = True
            txtFax.Enabled = True
            txtUrl.Enabled = True
            txtNum_ruc.Enabled = True
            grxDatos.Enabled = False


        Case "D" 'cuando deshace
            
            txtCod_cliente.Enabled = False
            txtNom_cliente.Enabled = False
            txtDEs_cliente.Enabled = False
            txtDireccion.Enabled = False
            txtLugar_Entrega.Enabled = False
            txtEmail.Enabled = False
            txtTelefono1.Enabled = False
            txtTelefono2.Enabled = False
            txtFax.Enabled = False
            txtUrl.Enabled = False
            txtNum_ruc.Enabled = False
            grxDatos.Enabled = True
    
    End Select
End Sub

Public Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True
                
            txtCod_cliente.Enabled = False
            txtNom_cliente.Enabled = False
            txtDEs_cliente.Enabled = False
            txtDireccion.Enabled = False
            txtLugar_Entrega.Enabled = False
            txtEmail.Enabled = False
            txtTelefono1.Enabled = False
            txtTelefono2.Enabled = False
            txtFax.Enabled = False
            txtUrl.Enabled = False
            txtNum_ruc.Enabled = False
            grxDatos.Enabled = True
                

    If Trim(txtNom_cliente.Text) = "" Then
        VALIDA_DATOS = False
        Call MsgBox("sirvase ingresar una el nombre del cliente", vbInformation, "Mensaje")
        txtNom_cliente.SetFocus
        Exit Function
    End If
    
End Function

Public Sub LIMPIAR_DATOS()
    
        txtCod_cliente.Text = "0"
        txtNom_cliente.Text = ""
        txtDEs_cliente.Text = ""
        txtDireccion.Text = ""
        txtLugar_Entrega.Text = ""
        txtEmail.Text = ""
        txtTelefono1.Text = ""
        txtTelefono2.Text = ""
        txtFax.Text = ""
        txtUrl.Text = ""
        txtNum_ruc.Text = ""

End Sub
Private Sub buscardatos()
    On Error GoTo Salvar_DatosErr

    StrSQL = " UP_MAN_CLIENTE_VENTA_PRENDAS 'V','','" & Trim(txtBus_Num_Ruc.Text) & "','" & Trim(txtBus_Nom_Cliente.Text) & "','','','','','','','',''"
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

        txtCod_cliente.Text = grxDatos.Value(grxDatos.Columns("COD_CLIENTE").Index)
        txtNom_cliente.Text = grxDatos.Value(grxDatos.Columns("NOM_CLIENTE").Index)
        txtDEs_cliente.Text = grxDatos.Value(grxDatos.Columns("DES_CLIENTE").Index)
        txtDireccion.Text = grxDatos.Value(grxDatos.Columns("DIRECCION").Index)
        txtLugar_Entrega.Text = grxDatos.Value(grxDatos.Columns("LUG_ENTREGA").Index)
        txtEmail.Text = grxDatos.Value(grxDatos.Columns("EMAIL").Index)
        txtTelefono1.Text = grxDatos.Value(grxDatos.Columns("TELEFONO1").Index)
        txtTelefono2.Text = grxDatos.Value(grxDatos.Columns("TELEFONO2").Index)
        txtFax.Text = grxDatos.Value(grxDatos.Columns("FAX").Index)
        txtUrl.Text = grxDatos.Value(grxDatos.Columns("URL").Index)
        txtNum_ruc.Text = grxDatos.Value(grxDatos.Columns("NUM_RUC").Index)
End Sub
Private Sub SALVAR_DATOS()
    On Error GoTo Salvar_DatosErr
    
    StrSQL = " UP_MAN_CLIENTE_VENTA_PRENDAS '" & sTipo & _
                                                "' ,'" & Trim(txtCod_cliente.Text) & _
                                                "', '" & Trim(txtNum_ruc.Text) & _
                                                "', '" & Trim(txtNom_cliente.Text) & _
                                                "', '" & Trim(txtDEs_cliente.Text) & _
                                                "', '" & Trim(txtDireccion.Text) & _
                                                "','" & Trim(txtLugar_Entrega.Text) & _
                                                "','" & Trim(txtEmail.Text) & _
                                                "','" & Trim(txtTelefono1.Text) & _
                                                "','" & Trim(txtTelefono2.Text) & _
                                                "','" & Trim(txtFax.Text) & _
                                                "','" & Trim(txtUrl.Text) & "'"
                
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
    StrSQL = " UP_MAN_CLIENTE_VENTA_PRENDAS '" & sTipo & _
                                                "' ,'" & Trim(txtCod_cliente.Text) & _
                                                "', '" & Trim(txtNum_ruc.Text) & _
                                                "', '" & Trim(txtNom_cliente.Text) & _
                                                "', '" & Trim(txtDEs_cliente.Text) & _
                                                "', '" & Trim(txtDireccion.Text) & _
                                                "','" & Trim(txtLugar_Entrega.Text) & _
                                                "','" & Trim(txtEmail.Text) & _
                                                "','" & Trim(txtTelefono1.Text) & _
                                                "','" & Trim(txtTelefono2.Text) & _
                                                "','" & Trim(txtFax.Text) & _
                                                "','" & Trim(txtUrl.Text) & "'"
                
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

        With .Columns("COD_CLIENTE")
            .Width = 800
            .TextAlignment = jgexAlignLeft
            .Caption = "COD"
        End With

        With .Columns("NUM_RUC")
            .Width = 1500
            .TextAlignment = jgexAlignLeft
            .Visible = True
            .Caption = "RUC"
        End With
        With .Columns("NOM_CLIENTE")
            .Width = 2500
            .TextAlignment = jgexAlignLeft
            .Caption = "CLIENTE"
            .Visible = True
        End With
         'COD_TELA
        With .Columns("DES_CLIENTE")
            .Width = 1500
            .TextAlignment = jgexAlignLeft
            .Caption = "DESCRIPCION"
            .Visible = True
        End With
           'COD_INTCOL
        With .Columns("DIRECCION")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "DIRECCION"
            .Visible = True
        End With
           'COD_TELA_EQUI
        With .Columns("LUG_ENTREGA")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "LUG"
            .Visible = True
        End With
        
        With .Columns("EMAIL")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "EMAIL"
            .Visible = True
        End With
           'des_color
        With .Columns("TELEFONO1")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "TELEFONO1"
            .Visible = True
        End With

        With .Columns("TELEFONO2")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "TELEFONO2"
            .Visible = True
        End With
           
        With .Columns("FAX")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "FAX"
            .Visible = True
        End With
        With .Columns("URL")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "URL"
            .Visible = True
        End With
    End With
    Call setcolorcolumnas
End Sub
Private Sub setcolorcolumnas()

    grxDatos.Columns("NOM_CLIENTE").CellStyle = "cod1"
    grxDatos.Columns("NUM_RUC").CellStyle = "cod2"
        
End Sub
''''busqueda
Private Sub txtBus_Num_Ruc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     txtBus_Nom_Cliente.Text = ""
     Call buscardatos
   End If
End Sub
Private Sub txtBus_Num_Ruc_Change()
     If Len(Trim(txtBus_Num_Ruc)) > 4 Then
        txtBus_Nom_Cliente.Text = ""
        Call buscardatos
     End If
End Sub
Private Sub txtBus_Nom_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtBus_Num_Ruc.Text = ""
        Call buscardatos
    End If
End Sub
Private Sub txtNum_Ruc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtNom_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtDes_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtLugar_Entrega_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtTelefono1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtTelefono2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub TXTFAX_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub TXTEMAIL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub TXTURL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

