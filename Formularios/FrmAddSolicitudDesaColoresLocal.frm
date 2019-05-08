VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmAddSolicitudDesaColoresLocal 
   Caption         =   "Adicionar Solicitud"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraDatos 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.TextBox TxtNum_Carta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox TxtCorr_Carta 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdCliente 
         Caption         =   "..."
         Height          =   270
         Left            =   2450
         TabIndex        =   4
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   600
         Width           =   5415
      End
      Begin VB.TextBox TxtDes_Cliente 
         Height          =   285
         Left            =   2870
         TabIndex        =   3
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox TxtCod_Cliente 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPSolicitud 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   1680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   68485121
         CurrentDate     =   38210
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Num. Carta"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   2160
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Corr. Carta"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   315
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   675
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Solicitud"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1035
         Width           =   480
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2280
      TabIndex        =   9
      Top             =   2760
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmAddSolicitudDesaColoresLocal.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "FrmAddSolicitudDesaColoresLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public CODIGO, Descripcion, TipoAdd As String
Public sAccion As String
Public vOk As Boolean
Public Num_carta As Integer
Dim StrSQL As String


Private Sub CmdCliente_Click()
Dim oTipo As New frmBusqGeneral
Dim RS As New ADODB.Recordset
Set oTipo.oParent = Me
oTipo.SQuery = "SELECT Abr_Cliente as Código, nom_cliente as Descripción FROM Tx_Cliente ORDER BY Abr_Cliente"
oTipo.Cargar_Datos
oTipo.Show 1
If CODIGO <> "" Then
    TxtCod_Cliente.Text = CODIGO
    TxtDes_Cliente.Text = Descripcion
    'txtCod_TemCli.Enabled = True
    'TxtDes_TemCli.Enabled = True
    'CmdTempCli.Enabled = True
    'txtCod_TemCli.SetFocus
    CODIGO = ""
End If
Set oTipo = Nothing
Set RS = Nothing
End Sub

'Private Sub CmdTempCli_Click()
'Call BUSCA_TEMPORADA(3)
'End Sub


Private Sub Command1_Click()
'Load FrmManteTemCli
'Set FrmManteTemCli.oParent = Me
'FrmManteTemCli.Show vbModal
'Set FrmManteTemCli = Nothing

End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    Call Grabar
Case "CANCELAR"
    vOk = False
    Unload Me
End Select
End Sub

Sub Grabar()
Dim fecha As String
On Error GoTo errGrabar
If DTPSolicitud.Value <> "NULL" Then
    fecha = "'" & DTPSolicitud.Value & "'"
Else
    fecha = "NULL"
End If

StrSQL = "SELECT Cod_Cliente_Tex FROM Tx_CLIENTE WHERE Abr_Cliente='" & Trim(TxtCod_Cliente.Text) & "'"

StrSQL = "es_up_man_lb_cartacol_dg_Local '" & sAccion & "','" & TxtCorr_Carta.Text & "','" & TxtDescripcion.Text & "','" & DevuelveCampo(StrSQL, cConnect) & "',''," & _
            fecha & "," & Val(TxtNum_Carta.Text)

ExecuteSQL cConnect, StrSQL

StrSQL = "SELECT corr_carta_local from tg_control"
Num_carta = DevuelveCampo(StrSQL, cConnect)

vOk = True
Unload Me
Exit Sub
errGrabar:
    vOk = False
    ErrorHandler err, "Grabar"
End Sub

Public Sub BUSCA_TEMPORADA(Tipo As Integer)
Dim SCOD_CLIENTE As String

'StrSQL = "SELECT cod_cliente_tex FROM Tx_cliente WHERE abr_cliente = '" & Trim(Me.TxtCod_Cliente.Text) & "'"
'SCOD_CLIENTE = DevuelveCampo(StrSQL, cConnect)
'
'    Select Case Tipo
'        Case 1:
'                    StrSQL = "SELECT nom_temcli FROM Tx_TemCli WHERE cod_cliente = '" & SCOD_CLIENTE & "' and cod_temcli='" & txtCod_TemCli.Text & "'"
'                    Me.TxtDes_TemCli.Text = Trim(DevuelveCampo(StrSQL, cConnect))
'                    DTPSolicitud.SetFocus
'        Case 2, 3:
'                    Dim oTipo As New frmBusqGeneral
'                    Dim RS As New ADODB.Recordset
'                    Set oTipo.oParent = Me
'
'                    If Tipo = 2 Then
'                        oTipo.SQuery = "SELECT cod_temcli AS 'Código', nom_temcli AS 'Descripción' FROM Tx_TemCli where cod_cliente = '" & SCOD_CLIENTE & "' and nom_temcli like '%" & Trim(TxtDes_TemCli.Text) & "%' order by des_temcli"
'                    Else
'                        oTipo.SQuery = "SELECT cod_temcli AS 'Código', nom_temcli AS 'Descripción' FROM Tx_TemCli where cod_cliente = '" & SCOD_CLIENTE & "' order by nom_temcli"
'                    End If
'
'                    oTipo.Cargar_Datos
'                    oTipo.Show 1
'                    If CODIGO <> "" Then
'                         Me.txtCod_TemCli.Text = Trim(CODIGO)
'                         Me.TxtDes_TemCli.Text = Trim(Descripcion)
'                         CODIGO = "": Descripcion = ""
'                         DTPSolicitud.SetFocus
'                    End If
'                    Set oTipo = Nothing
'                    Set RS = Nothing
'    End Select
    
End Sub

Private Sub TxtCod_Cliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Trim(TxtCod_Cliente.Text) = "" Then
            CmdCliente_Click
        Else
            StrSQL = "SELECT Nom_Cliente FROM Tx_CLIENTE WHERE Abr_Cliente LIKE '" & Trim(TxtCod_Cliente.Text) & "%'"
            TxtDes_Cliente.Text = DevuelveCampo(StrSQL, cConnect)
            'txtCod_TemCli.Enabled = True
            'TxtDes_TemCli.Enabled = True
            'CmdTempCli.Enabled = True
            'txtCod_TemCli.SetFocus
        End If
    End If
End Sub

'Private Sub txtCod_TemCli_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If Trim(txtCod_TemCli.Text) = "" Then
'        Call BUSCA_TEMPORADA(3)
'    Else
'        Call BUSCA_TEMPORADA(1)
'    End If
'End If
'End Sub

Private Sub txtDes_Cliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Len(TxtDes_Cliente) > 4 Then
            StrSQL = "SELECT Abr_Cliente FROM Tx_CLIENTE WHERE Nom_Cliente LIKE '" & Trim(TxtDes_Cliente.Text) & "%'"
            TxtCod_Cliente.Text = DevuelveCampo(StrSQL, cConnect)
            StrSQL = "SELECT Nom_Cliente FROM Tx_CLIENTE WHERE Abr_Cliente='" & Trim(TxtCod_Cliente.Text) & "'"
            TxtDes_Cliente.Text = DevuelveCampo(StrSQL, cConnect)
'            txtCod_TemCli.Enabled = True
'            TxtDes_TemCli.Enabled = True
'            CmdTempCli.Enabled = True
'            txtCod_TemCli.SetFocus
        Else
            MsgBox ("El Texto Ingresado debe contar con un mínimo de 5 caracteres")
            TxtDes_Cliente.SetFocus
        End If
    End If
End Sub

'Private Sub TxtDes_TemCli_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    Call BUSCA_TEMPORADA(2)
'End If
'End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtNum_Carta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub


