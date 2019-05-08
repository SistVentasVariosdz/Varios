VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmUsuarioCorrelativos 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Usuarios Correlativos Guias de Remision"
   ClientHeight    =   8070
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Usuarios Autorizados por Serie"
      Height          =   2415
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Width           =   7335
      Begin GridEX20.GridEX GrdUserSerie 
         Height          =   2055
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   3625
         Version         =   "2.0"
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         ColumnAutoResize=   -1  'True
         ReadOnly        =   -1  'True
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         BackColorBkg    =   12648384
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         FormatStylesCount=   7
         FormatStyle(1)  =   "FrmUsuarioCorrelativos.frx":0000
         FormatStyle(2)  =   "FrmUsuarioCorrelativos.frx":0138
         FormatStyle(3)  =   "FrmUsuarioCorrelativos.frx":01E8
         FormatStyle(4)  =   "FrmUsuarioCorrelativos.frx":029C
         FormatStyle(5)  =   "FrmUsuarioCorrelativos.frx":0374
         FormatStyle(6)  =   "FrmUsuarioCorrelativos.frx":042C
         FormatStyle(7)  =   "FrmUsuarioCorrelativos.frx":050C
         ImageCount      =   0
         PrinterProperties=   "FrmUsuarioCorrelativos.frx":052C
      End
   End
   Begin VB.CommandButton btnAdicionar 
      Caption         =   "Adicionar"
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton BtnEliminar 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Adicionar Usuario"
      Height          =   3015
      Left            =   240
      TabIndex        =   7
      Top             =   4680
      Width           =   7335
      Begin GridEX20.GridEX grdUserDisponibles 
         Height          =   2655
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   4683
         Version         =   "2.0"
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         ColumnAutoResize=   -1  'True
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         BackColorBkg    =   12648384
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         FormatStylesCount=   7
         FormatStyle(1)  =   "FrmUsuarioCorrelativos.frx":0704
         FormatStyle(2)  =   "FrmUsuarioCorrelativos.frx":083C
         FormatStyle(3)  =   "FrmUsuarioCorrelativos.frx":08EC
         FormatStyle(4)  =   "FrmUsuarioCorrelativos.frx":09A0
         FormatStyle(5)  =   "FrmUsuarioCorrelativos.frx":0A78
         FormatStyle(6)  =   "FrmUsuarioCorrelativos.frx":0B30
         FormatStyle(7)  =   "FrmUsuarioCorrelativos.frx":0C10
         ImageCount      =   0
         PrinterProperties=   "FrmUsuarioCorrelativos.frx":0C30
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Tipo Almacen"
      Height          =   735
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   6975
      Begin VB.OptionButton OptTextiles 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Textiles"
         Height          =   255
         Left            =   4320
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optConfecciones 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Confecciones"
         Height          =   195
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSDataListLib.DataCombo cmbSerie 
      Height          =   315
      Left            =   960
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo cmbAlmacen 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   "DataCombo1"
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   4800
      Top             =   7440
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Nro Serie"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Almacen:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "FrmUsuarioCorrelativos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tipo_almacen As String
'Dim Rsx_CF As ADODB.Recordset, Rsx_LG As ADODB.Recordset
Dim RsxSerie As ADODB.Recordset
Dim RsxUsuariosDisponibles As New ADODB.Recordset
Dim Cod_almacen As String, Num_serie As String

Dim Cod_UsuarioX As String, Nom_UsuarioX As String
Dim Cod_UsuarioY As String, Nom_UsuarioY As String
Dim mant_delete As String, mant_Insert As String
Dim CANT_DEL As Long
Dim cant_Ins As Long

Dim strsql_CF As String
Dim strsql_LG As String
Dim Strsql_Series As String

Dim GrdUserHabilitados As String
Dim GrdUserInhabilitados As String

Dim Rsx_CF As New ADODB.Recordset
Dim Rsx_LG As New ADODB.Recordset
Dim rsx_Series As New ADODB.Recordset
Dim rsx_UserHabilitados As New ADODB.Recordset
Dim rsx_UserInhabilitados As New ADODB.Recordset
Dim rsx_Delete As New ADODB.Recordset
Dim rsx_Insert As New ADODB.Recordset

Sub limpiar_Variables()
Cod_UsuarioX = ""
Cod_UsuarioY = ""
Nom_UsuarioX = ""
Nom_UsuarioY = ""
End Sub

Private Sub btnAdicionar_Click()
On Erro GoTo Hand:
    If Cod_UsuarioY <> "" And Nom_UsuarioY <> "" And Tipo_almacen <> "" And Cod_almacen <> "" And Num_serie <> "" Then
            Dim vMessage As Variant
            vMessage = (MsgBox("¿Desea otorgar el permiso a  " + Nom_UsuarioY + " sobre el numero de Serie " + Num_serie + " del almacen " + cmbAlmacen.BoundText + "?", vbQuestion + vbYesNo, sTit))
            If vMessage = vbYes Then
                mant_Insert = "Exec Lg_Mant_Nro_Serie_x_Usuarios '6','" + Tipo_almacen + "','" + Cod_almacen + "','" + Num_serie + "','" + Cod_UsuarioY + "',''"
                cant_Ins = ExecuteSQL(cConnect, mant_Insert)
                MsgBox "Usuario Agregado correctamente", vbInformation, "Agregar Usuario"
                cmbSerie_Change
            End If
            Exit Sub
    Else
        MsgBox "Debe seleccionar el usuario a agregar", vbInfoBackground, "Adicionar Usuario"
        Exit Sub
    End If
Hand:
    MsgBox Err.Description, vbCritical, "Adicionar Usuario"
    Exit Sub
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub grdUserDisponibles_Click()
Cod_UsuarioY = grdUserDisponibles.Value(1)
Nom_UsuarioY = grdUserDisponibles.Value(2)
End Sub

Private Sub optConfecciones_Click()
    cmbAlmacen.Text = ""
    Tipo_almacen = "CF"
    Set Rsx_CF = New ADODB.Recordset
    strsql_CF = "EXEC Lg_Mant_Nro_Serie_x_Usuarios '1','" + Tipo_almacen + "','','','',''"
    Set Rsx_CF.DataSource = CargarRecordSetDesconectado(strsql_CF, cConnect)
    Set cmbAlmacen.RowSource = Rsx_CF
    cmbAlmacen.ListField = Rsx_CF.Fields(1).Name
    cmbAlmacen.BoundColumn = Rsx_CF.Fields(0).Name
    limpiar_Variables
End Sub

Private Sub OptTextiles_Click()
    cmbAlmacen.Text = ""
    Tipo_almacen = "LG"
    Set Rsx_LG = New ADODB.Recordset
    strsql_LG = "EXEC Lg_Mant_Nro_Serie_x_Usuarios '1','" + Tipo_almacen + "','','','',''"
    Set Rsx_LG.DataSource = CargarRecordSetDesconectado(strsql_LG, cConnect)
    Set cmbAlmacen.RowSource = Rsx_LG
    cmbAlmacen.ListField = Rsx_LG.Fields(1).Name
    cmbAlmacen.BoundColumn = Rsx_LG.Fields(0).Name
    limpiar_Variables
End Sub

Private Sub BtnEliminar_Click()
On Error GoTo Hand
If Nom_UsuarioX <> "" And Cod_UsuarioX <> "" Then
    Dim vMessage As Variant
    vMessage = (MsgBox("¿Desea quitar el permiso a  " + Nom_UsuarioX + " sobre el numero de Serie " + Num_serie + " del almacen " + cmbAlmacen.BoundText + "?", vbQuestion + vbYesNo, sTit))

    If vMessage = vbYes Then
        mant_delete = "Exec Lg_Mant_Nro_Serie_x_Usuarios '5','" + Tipo_almacen + "','" + Cod_almacen + "','" + Num_serie + "','" + Cod_UsuarioX + "',''"
         CANT_DEL = ExecuteSQL(cConnect, mant_delete)
            MsgBox "Usuario eliminado correctamente", vbInactiveBorder, "Eliminar"
            cmbSerie_Change
    End If
    Exit Sub
Else
    MsgBox "No ha seleccionado ningun usuario de la lista", vbCritical, "Eliminar"
    Exit Sub
End If
Hand:
MsgBox Err.Description
    Exit Sub
End Sub

Private Sub cmbAlmacen_Change()
    cmbSerie.Text = ""
    Cod_almacen = cmbAlmacen.BoundText
    Strsql_Series = "Exec Lg_Mant_Nro_Serie_x_Usuarios '2','" + Tipo_almacen + "','" + Cod_almacen + "','','',''"
    Set rsx_Series = New ADODB.Recordset
    Set rsx_Series.DataSource = CargarRecordSetDesconectado(Strsql_Series, cConnect)
    Set cmbSerie.RowSource = rsx_Series
    cmbSerie.ListField = rsx_Series.Fields(0).Name
    cmbSerie.BoundColumn = rsx_Series.Fields(0).Name
End Sub

Private Sub cmbSerie_Change()
    Set GrdUserSerie.ADORecordset = Nothing
    Num_serie = cmbSerie.BoundText
    Set rsx_UserHabilitados = New ADODB.Recordset
    GrdUserHabilitados = "Exec Lg_Mant_Nro_Serie_x_Usuarios '3','" + Tipo_almacen + "','" + Cod_almacen + "','" + Num_serie + "','',''"
    Set rsx_UserHabilitados.DataSource = CargarRecordSetDesconectado(GrdUserHabilitados, cConnect)
    Set GrdUserSerie.ADORecordset = rsx_UserHabilitados
Call UsuariosDisponibles
End Sub

Private Sub GrdUserSerie_Click()
    Cod_UsuarioX = GrdUserSerie.Value(1)
    Nom_UsuarioX = GrdUserSerie.Value(2)
End Sub

Private Sub UsuariosDisponibles()
    Set rsx_UserInhabilitados = New ADODB.Recordset
    GrdUserInhabilitados = "Exec Lg_Mant_Nro_Serie_x_Usuarios '4','" + Tipo_almacen + "','" + Cod_almacen + "','" + Num_serie + "','',''"
    Set rsx_UserInhabilitados.DataSource = CargarRecordSetDesconectado(GrdUserInhabilitados, cConnect)
    Set grdUserDisponibles.ADORecordset = rsx_UserInhabilitados
End Sub


