VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCambiarUltimoCorrelativo 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Cambiar Correlativo Serie"
   ClientHeight    =   6135
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNuevo 
      Height          =   285
      Left            =   2040
      MaxLength       =   8
      TabIndex        =   17
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Tipo Almacen"
      Height          =   615
      Left            =   360
      TabIndex        =   14
      Top             =   1080
      Width           =   6855
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Textiles"
         Height          =   255
         Left            =   4200
         TabIndex        =   16
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Confecciones"
         Height          =   195
         Left            =   480
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox txtActual 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Top             =   2880
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo cmbSerie 
      Height          =   315
      Left            =   1080
      TabIndex        =   10
      Top             =   2400
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo cmbAlmacen 
      Height          =   315
      Left            =   1080
      TabIndex        =   9
      Top             =   1920
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   "DataCombo1"
   End
   Begin VB.TextBox txtUsuario 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   720
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Motivo"
      Height          =   1335
      Left            =   240
      TabIndex        =   6
      Top             =   4080
      Width           =   7215
      Begin VB.TextBox txtMotivo 
         Height          =   855
         Left            =   240
         MaxLength       =   240
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   6735
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   3840
      Top             =   5640
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Nuevo Correlativo:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Correlativo Actual:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Serie:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Almacen:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Usuario:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cambio de Correlativos en Nro de Serie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmCambiarUltimoCorrelativo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strsql_User As String
Dim strsql_Tipo_almacen As String
Dim strsql_Num_series As String
Dim strsql_Ultimo_correlativo As String
Dim strsql_Mant As String

Dim rsx_Usuario As New ADODB.Recordset
Dim rsx_Almacenes As New ADODB.Recordset
Dim rsx_Series As New ADODB.Recordset
Dim rsx_UltimoCorrelativo  As New ADODB.Recordset

Dim Tipo_AlmacenX As String
Dim Cod_UsuarioX As String
Dim Cod_almacenX As String
Dim Num_SerieX  As String
Dim Correlativo_ActualX As String
Dim Correlativo_NuevoX As String
Dim MotivoX As String

Dim nom_almacenX As String

Dim Res_MantX As Long


Private Sub cmbAlmacen_Change()

    cmbSerie.Text = ""
    Cod_almacenX = cmbAlmacen.BoundText
    nom_almacenX = cmbAlmacen.Text
    strsql_Num_series = "Exec LG_Modifica_Usuario_Correlativo '2','" + Tipo_AlmacenX + "','" + Cod_UsuarioX + "','" + Cod_almacenX + "','','','',''"
    Set rsx_Series = New ADODB.Recordset
    Set rsx_Series.DataSource = CargarRecordSetDesconectado(strsql_Num_series, cConnect)
    Set cmbSerie.RowSource = rsx_Series
    cmbSerie.ListField = rsx_Series.Fields(1).Name
    cmbSerie.BoundColumn = rsx_Series.Fields(1).Name
End Sub

Private Sub cmbSerie_Change()
If cmbSerie.BoundText <> "" Then
    Num_SerieX = cmbSerie.BoundText
    txtActual.Text = ""
    strsql_Ultimo_correlativo = "Exec LG_Modifica_Usuario_Correlativo '3','" + Tipo_AlmacenX + "','" + Cod_UsuarioX + "','" + Cod_almacenX + "','" + Num_SerieX + "','','',''"
    Set rsx_UltimoCorrelativo = New ADODB.Recordset
    Set rsx_UltimoCorrelativo.DataSource = CargarRecordSetDesconectado(strsql_Ultimo_correlativo, cConnect)
    txtActual.Text = rsx_UltimoCorrelativo.Fields(0).Value
End If
End Sub

Private Sub Command1_Click()
On Erro GoTo Hand:
Correlativo_ActualX = txtActual.Text
Correlativo_NuevoX = txtNuevo.Text
Cod_UsuarioX = vusu
MotivoX = txtMotivo.Text

If Correlativo_ActualX = "" Or Correlativo_NuevoX = "" Or Cod_UsuarioX = "" Or Len(Trim(txtMotivo.Text)) < 5 _
    Or Num_SerieX = "" Or Cod_almacenX = "" Or Tipo_AlmacenX = "" Then
    MsgBox "Hay datos vacios o la descripcion del motivo es muy corta, favor de verificar y reintentar", vbCritical, "Actualizar correlativo"
    Exit Sub
Else
    strsql_Mant = "Exec LG_Modifica_Usuario_Correlativo '4','" + Tipo_AlmacenX + "','" + Cod_UsuarioX + "','" + Cod_almacenX + "','" + Num_SerieX + "','" + Correlativo_ActualX + "','" + Correlativo_NuevoX + "','" + MotivoX + "'"
    Res_MantX = ExecuteSQL(cConnect, strsql_Mant)
    MsgBox "El Correlativo " + Correlativo_ActualX + " fue cambiado por el correlativo " + Correlativo_NuevoX + " para el Nro de Serie " + Num_SerieX + " en el almacen " + nom_almacenX, vbInformation, "Datos Actualizados"
    Call cmbSerie_Change
    txtMotivo.Text = ""
    txtNuevo.Text = ""
End If
    Exit Sub
Hand:
MsgBox Err.Description, vbCritical, "Actualizar Correlativo"
Exit Sub
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Cod_UsuarioX = vusu
    strsql_User = "Select NOm_Usuario from [Seguridad]..seg_Usuarios where cod_Usuario = '" + vusu + "'"
    Set rsx_Usuario = New ADODB.Recordset
    Set rsx_Usuario.DataSource = CargarRecordSetDesconectado(strsql_User, cConnect)
    txtUsuario = rsx_Usuario(0).Value
    txtMotivo.Text = ""
End Sub
Sub Carga_almacenes()
    strsql_Tipo_almacen = "Exec LG_Modifica_Usuario_Correlativo '1','" + Tipo_AlmacenX + "','" + Cod_UsuarioX + "','','','','',''"
    Set rsx_Almacenes = New ADODB.Recordset
    Set rsx_Almacenes.DataSource = CargarRecordSetDesconectado(strsql_Tipo_almacen, cConnect)
    Set cmbAlmacen.RowSource = rsx_Almacenes
    cmbAlmacen.ListField = rsx_Almacenes.Fields(1).Name
    cmbAlmacen.BoundColumn = rsx_Almacenes.Fields(0).Name
End Sub

Private Sub Option1_Click()
    Tipo_AlmacenX = "CF"
    cmbAlmacen.Text = ""
    txtActual.Text = ""
    Call Carga_almacenes
End Sub

Private Sub Option2_Click()
    Tipo_AlmacenX = "LG"
    cmbAlmacen.Text = ""
    txtActual.Text = ""
    Call Carga_almacenes
End Sub


Private Sub txtNuevo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
        'TxtCod_Motivo.SetFocus
    End If
End Sub

Private Sub txtNuevo_KeyPress(KeyAscii As Integer)
        Call SoloNumeros(txtNuevo, KeyAscii, False, 0, 8)
End Sub

Private Sub txtNuevo_LostFocus()
        txtNuevo = Format(txtNuevo, "00000000")
End Sub
