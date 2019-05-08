VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Begin VB.Form FrmManPaisesCodigoIso 
   Caption         =   "Mantenimiento de Pais "
   ClientHeight    =   7965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   11550
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      Caption         =   "Datos"
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   0
      TabIndex        =   6
      Top             =   6000
      Width           =   11475
      Begin VB.TextBox txtCod_Pais 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   270
         Width           =   2265
      End
      Begin VB.TextBox txtDes_Pais_Iso 
         Height          =   315
         Left            =   4560
         TabIndex        =   11
         Top             =   850
         Width           =   6825
      End
      Begin VB.TextBox txtCod_Pais_Iso 
         Height          =   315
         Left            =   1080
         TabIndex        =   10
         Top             =   850
         Width           =   2265
      End
      Begin VB.TextBox txtCod_RTPS 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   570
         Width           =   2265
      End
      Begin VB.TextBox txtCod_Postal 
         Height          =   285
         Left            =   4560
         TabIndex        =   8
         Top             =   570
         Width           =   2535
      End
      Begin VB.TextBox txtDes_Pais 
         Height          =   285
         Left            =   4560
         TabIndex        =   7
         Top             =   240
         Width           =   6825
      End
      Begin VB.Label cODIGO 
         AutoSize        =   -1  'True
         Caption         =   "Codigo Pais"
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
         Left            =   90
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Codigo ISO"
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
         Top             =   950
         Width           =   825
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Codigo RTPS"
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
         Left            =   15
         TabIndex        =   16
         Top             =   630
         Width           =   930
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
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
         Left            =   3645
         TabIndex        =   15
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Codigo Postal"
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
         Left            =   3465
         TabIndex        =   14
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion Iso"
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
         Left            =   3360
         TabIndex        =   13
         Top             =   960
         Width           =   1125
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
      Width           =   11475
      Begin VB.TextBox txtDes_Pais_Bus 
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
         Left            =   4080
         TabIndex        =   2
         Top             =   240
         Width           =   5955
      End
      Begin VB.TextBox txtCod_Pais_Bus 
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
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   1785
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Codigo Pais"
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
         TabIndex        =   5
         Top             =   330
         Width           =   855
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripcion Pais"
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
         Left            =   2880
         TabIndex        =   4
         Top             =   360
         Width           =   1200
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
         Left            =   10560
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin GridEX20.GridEX grxDatos 
      Height          =   5355
      Left            =   0
      TabIndex        =   19
      Top             =   660
      Width           =   11475
      _ExtentX        =   20241
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
      FormatStyle(1)  =   "FrmManPaisesCodigoIso.frx":0000
      FormatStyle(2)  =   "FrmManPaisesCodigoIso.frx":0128
      FormatStyle(3)  =   "FrmManPaisesCodigoIso.frx":01D8
      FormatStyle(4)  =   "FrmManPaisesCodigoIso.frx":028C
      FormatStyle(5)  =   "FrmManPaisesCodigoIso.frx":0364
      FormatStyle(6)  =   "FrmManPaisesCodigoIso.frx":041C
      FormatStyle(7)  =   "FrmManPaisesCodigoIso.frx":04FC
      FormatStyle(8)  =   "FrmManPaisesCodigoIso.frx":05A8
      FormatStyle(9)  =   "FrmManPaisesCodigoIso.frx":0674
      ImageCount      =   0
      PrinterProperties=   "FrmManPaisesCodigoIso.frx":071C
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   630
      Left            =   7920
      TabIndex        =   20
      Top             =   7320
      Width           =   3630
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"FrmManPaisesCodigoIso.frx":08EC
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1305
         Left            =   -1260
         TabIndex        =   21
         Top             =   -2970
         Width           =   6975
         Begin VB.TextBox Txtcod_Fabrica 
            Height          =   285
            Left            =   1440
            MaxLength       =   3
            TabIndex        =   25
            Top             =   240
            Width           =   795
         End
         Begin VB.TextBox TxtNom_Fabrica 
            Height          =   285
            Left            =   2250
            TabIndex        =   24
            Top             =   255
            Width           =   4635
         End
         Begin VB.TextBox Txt_Planilla 
            Height          =   285
            Left            =   2250
            TabIndex        =   23
            Top             =   570
            Width           =   4635
         End
         Begin VB.TextBox Txt_Tipo 
            Height          =   285
            Left            =   1440
            TabIndex        =   22
            Top             =   570
            Width           =   795
         End
         Begin VB.Label Label2 
            Caption         =   "Inicio:"
            Height          =   210
            Left            =   840
            TabIndex        =   29
            Top             =   960
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fabrica:"
            Height          =   195
            Left            =   690
            TabIndex        =   28
            Top             =   270
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Trabajador:"
            Height          =   195
            Left            =   90
            TabIndex        =   27
            Top             =   570
            Width           =   1170
         End
         Begin VB.Label Label4 
            Caption         =   "Fin:"
            Height          =   195
            Left            =   3990
            TabIndex        =   26
            Top             =   990
            Width           =   435
         End
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   720
      Top             =   7440
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmManPaisesCodigoIso"
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
            txtNum_Ruc.SetFocus
            
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
            
            TxtCod_Pais.Enabled = True
            TxtDes_Pais.Enabled = True
            txtCod_RTPS.Enabled = True
            txtCod_Postal.Enabled = True
            txtCod_Pais_Iso.Enabled = True
            txtDes_Pais_Iso.Enabled = True
            grxDatos.Enabled = False
            
            txtCod_Pais_Bus.Enabled = False
            txtDes_Pais_Bus.Enabled = False
        
        Case "M" 'cuando modifica
            
            TxtCod_Pais.Enabled = False
            TxtDes_Pais.Enabled = False
            txtCod_RTPS.Enabled = True
            txtCod_Postal.Enabled = True
            txtCod_Pais_Iso.Enabled = True
            txtDes_Pais_Iso.Enabled = True
            grxDatos.Enabled = False
            txtCod_Pais_Bus.Enabled = False
            txtDes_Pais_Bus.Enabled = False
            
        Case "D" 'cuando deshace
            
            
            TxtCod_Pais.Enabled = False
            TxtDes_Pais.Enabled = False
            txtCod_RTPS.Enabled = False
            txtCod_Postal.Enabled = False
            txtCod_Pais_Iso.Enabled = False
            txtDes_Pais_Iso.Enabled = False
            grxDatos.Enabled = True
            txtCod_Pais_Bus.Enabled = False
            txtDes_Pais_Bus.Enabled = False

    End Select
End Sub

Public Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True
                
            TxtCod_Pais.Enabled = False
            TxtDes_Pais.Enabled = False
            txtCod_RTPS.Enabled = False
            txtCod_Postal.Enabled = False
            txtCod_Pais_Iso.Enabled = False
            txtDes_Pais_Iso.Enabled = False
            grxDatos.Enabled = True
            txtCod_Pais_Bus.Enabled = False
            txtDes_Pais_Bus.Enabled = False

    If Trim(TxtCod_Pais.Text) = "" Then
        VALIDA_DATOS = False
        MsgBox "sirvase ingresar un el Codigo del Pais", vbInformation + vbOKOnly, "Mensaje"
        TxtCod_Pais.SetFocus
        Exit Function
    End If
    If Trim(TxtDes_Pais.Text) = "" Then
        VALIDA_DATOS = False
        MsgBox "sirvase ingresar una Descripcion para el Pais", vbInformation + vbOKOnly, "Mensaje"
        TxtDes_Pais.SetFocus
        Exit Function
    End If
    
    
End Function

Public Sub LIMPIAR_DATOS()
    
        TxtCod_Pais.Text = ""
        TxtDes_Pais.Text = ""
        txtCod_RTPS.Text = ""
        txtCod_Postal.Text = ""
        txtCod_Pais_Iso.Text = ""
        txtDes_Pais_Iso.Text = ""

End Sub
Private Sub buscardatos()
    On Error GoTo Salvar_DatosErr
    
    StrSQL = " CN_PAISES_MAN 'V','2','" & Trim(txtCod_Pais_Bus.Text) & "','" & Trim(txtDes_Pais_Bus.Text) & "','','',''"
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

        TxtCod_Pais.Text = grxDatos.Value(grxDatos.Columns("cod_pais").Index)
        TxtDes_Pais.Text = grxDatos.Value(grxDatos.Columns("descripcion").Index)
        txtCod_RTPS.Text = grxDatos.Value(grxDatos.Columns("cod_rtps").Index)
        txtCod_Pais_Iso.Text = grxDatos.Value(grxDatos.Columns("cod_pais_iso").Index)
        txtDes_Pais_Iso.Text = grxDatos.Value(grxDatos.Columns("des_pais_iso").Index)
        txtCod_Postal.Text = grxDatos.Value(grxDatos.Columns("cod_postal").Index)
        
End Sub
Private Sub SALVAR_DATOS()
    On Error GoTo Salvar_DatosErr
    Dim bus As String
    bus = "0"
    StrSQL = " CN_PAISES_MAN '" & sTipo & _
                                                "' ,'" & bus & _
                                                "', '" & Trim(TxtCod_Pais.Text) & _
                                                "', '" & Trim(TxtDes_Pais.Text) & _
                                                "', '" & Trim(txtCod_RTPS.Text) & _
                                                "', '" & Trim(txtCod_Pais_Iso.Text) & _
                                                "','" & Trim(txtCod_Postal.Text) & "'"
                                                                
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
    Dim bus As String
    bus = "0"
    StrSQL = " CN_PAISES_MAN '" & sTipo & _
                                                "' ,'" & bus & _
                                                "', '" & Trim(TxtCod_Pais.Text) & _
                                                "', '" & Trim(TxtDes_Pais.Text) & _
                                                "', '" & Trim(txtCod_RTPS.Text) & _
                                                "', '" & Trim(txtCod_Pais_Iso.Text) & _
                                                "','" & Trim(txtCod_Postal.Text) & "'"
                
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
        
        With .Columns("COD_PAIS")
            .Width = 800
            .TextAlignment = jgexAlignLeft
            .Caption = "Codigo"
        End With

        With .Columns("DESCRIPCION")
            .Width = 2000
            .TextAlignment = jgexAlignLeft
            .Visible = True
            .Caption = "Des. Pais"
        End With

        With .Columns("cod_Rtps")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Visible = True
            .Caption = "Cod. RTPS"
        End With

        With .Columns("cod_pais_iso")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Visible = True
            .Caption = "Cod Pais ISO"
        End With

        With .Columns("des_pais_iso")
            .Width = 2000
            .TextAlignment = jgexAlignLeft
            .Visible = True
            .Caption = "Des. Pais ISO"
        End With

        With .Columns("cod_postal")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Visible = True
            .Caption = "Cod. Postal"
        End With


    End With
    Call setcolorcolumnas
End Sub
Private Sub setcolorcolumnas()

    grxDatos.Columns("cod_pais").CellStyle = "cod1"
    grxDatos.Columns("Des_pais").CellStyle = "cod2"
        
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




