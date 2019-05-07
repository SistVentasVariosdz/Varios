VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "mantenimientos.ocx"
Begin VB.Form frmDatoCtaCliente 
   Caption         =   "Datos Beneficiario"
   ClientHeight    =   7185
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   11220
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   3825
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11190
      Begin GridEX20.GridEX gexList 
         Height          =   3480
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   10980
         _ExtentX        =   19368
         _ExtentY        =   6138
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         RecordNavigatorString=   "Registro:|de"
         HoldSortSettings=   -1  'True
         GridLineStyle   =   2
         ColumnAutoResize=   -1  'True
         HeaderStyle     =   3
         MethodHoldFields=   -1  'True
         ContScroll      =   -1  'True
         Options         =   8
         RecordsetType   =   1
         GroupByBoxInfoText=   ""
         AllowEdit       =   0   'False
         BorderStyle     =   2
         GroupByBoxVisible=   0   'False
         ImageCount      =   3
         ImagePicture1   =   "frmDatoCtaCliente.frx":0000
         ImagePicture2   =   "frmDatoCtaCliente.frx":0112
         ImagePicture3   =   "frmDatoCtaCliente.frx":042C
         RowHeaders      =   -1  'True
         DataMode        =   1
         HeaderFontName  =   "Tahoma"
         FontName        =   "Tahoma"
         GridLines       =   1
         BackColorBkg    =   -2147483624
         ColumnHeaderHeight=   285
         SortKeysCount   =   1
         SortKey(1)      =   "frmDatoCtaCliente.frx":0746
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmDatoCtaCliente.frx":07AE
         FormatStyle(2)  =   "frmDatoCtaCliente.frx":088E
         FormatStyle(3)  =   "frmDatoCtaCliente.frx":09B6
         FormatStyle(4)  =   "frmDatoCtaCliente.frx":0A66
         FormatStyle(5)  =   "frmDatoCtaCliente.frx":0B1A
         FormatStyle(6)  =   "frmDatoCtaCliente.frx":0BF2
         ImageCount      =   3
         ImagePicture(1) =   "frmDatoCtaCliente.frx":0CAA
         ImagePicture(2) =   "frmDatoCtaCliente.frx":0DBC
         ImagePicture(3) =   "frmDatoCtaCliente.frx":10D6
         PrinterProperties=   "frmDatoCtaCliente.frx":13F0
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   3840
      Width           =   11175
      Begin VB.ComboBox CmbEstado 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox Txt_Cod_Banco 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Txt_Nombre_Banco 
         Height          =   315
         Left            =   1920
         TabIndex        =   5
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox txt_nrocuenta 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txt_beneficiario 
         Height          =   285
         Left            =   1200
         MaxLength       =   150
         TabIndex        =   3
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox Txt_Direccion 
         Height          =   285
         Left            =   1200
         MaxLength       =   150
         TabIndex        =   2
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox Txt_Cod_Swift 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   1
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre Banco"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Estado :"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Banco"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "N° Cuenta"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Beneficiario"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Dirección"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Código Swift"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   1095
      End
   End
   Begin Mantenimientos.MantFunc MantFunc2 
      Height          =   540
      Left            =   3240
      TabIndex        =   17
      Top             =   6600
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmDatoCtaCliente.frx":15C0
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmDatoCtaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Public CODIGO As String
Public DESCRIPCION As String
Public Estado As String

Private Sub Form_Load()
    
    
    CmbEstado.AddItem "Activo", 0
    CmbEstado.AddItem "Inactivo", 1
    
    CARGA_GRID
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub gexList_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
CARGA_DATOS
End Sub

Private Sub Habilita()
Txt_Cod_Banco.Enabled = True
Txt_Nombre_Banco.Enabled = True
txt_nrocuenta.Enabled = True
End Sub
Private Sub DesHabilita()
'Txt_Cod_Banco.Enabled = False
'Txt_Nombre_Banco.Enabled = False
'txt_nrocuenta.Enabled = False

End Sub
Private Sub LIMPIA()
Txt_Cod_Banco.Text = ""
Txt_Nombre_Banco.Text = ""
txt_nrocuenta.Text = ""

txt_beneficiario.Text = ""
Txt_Direccion.Text = ""
Txt_Cod_Swift.Text = ""
CmbEstado.ListIndex = -1
End Sub
Private Sub MantFunc2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo hand
Select Case ActionName
    Case "ADICIONAR"
            HabilitaMant Me.MantFunc2, "GRABAR/DESHACER"
            LIMPIA
            Habilita
            Estado = "NUEVO"
    Case "MODIFICAR"

        HabilitaMant Me.MantFunc2, "GRABAR/DESHACER"
        Estado = "MODIFICAR"

    Case "ELIMINAR"
    
        DESHABILITA_DATOS
        LIMPIA
        DesHabilita
        
    Case "GRABAR"
         SALVAR_DATOS
         HabilitaMant Me.MantFunc2, "ADICIONAR/MODIFICAR/ELIMINAR"
         CARGA_GRID
         DesHabilita
         LIMPIA

    Case "DESHACER"
        HabilitaMant Me.MantFunc2, "ADICIONAR/MODIFICAR/ELIMINAR"
        LIMPIA
        DesHabilita
    Case "SALIR"
        Unload Me
End Select
Exit Sub
hand:
ErrorHandler err, "MantFunc1_ActionClick"

End Sub

Private Sub Txt_Cod_Banco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Txt_Cod_Banco.Text) = "" Then
            Call BUSCA_BANCO(3)
        Else
            Call BUSCA_BANCO(1)
        End If
    End If
End Sub

Private Sub Txt_Nombre_Banco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Txt_Nombre_Banco.Text) = "" Then
            Call BUSCA_BANCO(3)
        Else
            Call BUSCA_BANCO(2)
        End If
    End If
End Sub

Public Sub BUSCA_BANCO(Tipo As Integer)
Dim STRSQL As String

    Select Case Tipo
        Case 1:
                    STRSQL = "EXEC TI_BUSCA_BANCO 1,'" & Trim(Me.Txt_Cod_Banco.Text) & "',''"
                    Me.Txt_Nombre_Banco.Text = Trim(DevuelveCampo(STRSQL, cConnect))
                    
        Case 2, 3:
                    Dim oTipo As New frmBusGeneral6
                    Dim rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If Tipo = 2 Then
                        oTipo.SQuery = "EXEC TI_BUSCA_BANCO 2,'','" & Trim(Txt_Nombre_Banco.Text) & "'"
                    Else
                        oTipo.SQuery = "EXEC TI_BUSCA_BANCO 3,'',''"
                    End If
                    
                    oTipo.CARGAR_DATOS
                    oTipo.DGridLista.Columns(2).Width = 3500
                    oTipo.Show 1
                    If CODIGO <> "" Then
                         Me.Txt_Cod_Banco.Text = Trim(CODIGO)
                         Me.Txt_Nombre_Banco.Text = Trim(DESCRIPCION)
'                         OptCliPend.SetFocus
                         CODIGO = "": DESCRIPCION = ""
                         'CARGA_GRID
                    End If
                    Set oTipo = Nothing
                    Set rs = Nothing
    End Select
End Sub



Sub CARGA_GRID()
Dim Rs_Carga As New ADODB.Recordset
Dim sSQL As String

On Error GoTo Cargar_DatosErr
Rs_Carga.ActiveConnection = cConnect
Rs_Carga.CursorType = adOpenStatic
Rs_Carga.CursorLocation = adUseClient
Rs_Carga.LockType = adLockReadOnly

sSQL = "EXEC UP_SEL_CLIENTE_BANCO"

Rs_Carga.Open sSQL

Set gexList.ADORecordset = Rs_Carga
'ConfiguraGrid
Exit Sub
Cargar_DatosErr:
    Set Rs_Carga = Nothing
    ErrorHandler err, "CARGA_GRID"
End Sub

Sub CARGA_DATOS()
On Error GoTo Cargar_DatosErr

If gexList.RowCount = 0 Then
    Txt_Cod_Banco.Text = ""
    Txt_Nombre_Banco.Text = ""
    txt_nrocuenta.Text = ""
    txt_beneficiario.Text = ""
    Txt_Direccion.Text = ""
    Txt_Cod_Swift.Text = ""
    Exit Sub
End If


    Txt_Cod_Banco.Text = gexList.Value(gexList.Columns("cod_banco").Index)
    Txt_Nombre_Banco.Text = Trim(gexList.Value(gexList.Columns("Nom_Banco").Index))
    txt_nrocuenta.Text = Trim(gexList.Value(gexList.Columns("Nro_Cuenta").Index))
    txt_beneficiario.Text = Trim(gexList.Value(gexList.Columns("beneficiario").Index))
    Txt_Direccion.Text = Trim(gexList.Value(gexList.Columns("Direccion").Index))
    Txt_Cod_Swift.Text = Trim(gexList.Value(gexList.Columns("cod_swift").Index))



    If gexList.Value(gexList.Columns("Estado").Index) = "I" Then
        CmbEstado.ListIndex = 1
    Else
        CmbEstado.ListIndex = 0
    End If

Exit Sub
Cargar_DatosErr:
    ErrorHandler err, "CARGA_GRID"
End Sub
Sub DESHABILITA_DATOS()
Dim rsd As ADODB.Recordset
On Error GoTo Cargar_DatosErr

Set rsd = New ADODB.Recordset
rsd.ActiveConnection = cConnect
rsd.CursorLocation = adUseClient
rsd.CursorType = adOpenStatic

rsd.Open "exec UP_DEL_CLIENTE_BANCO '" & Trim(Txt_Cod_Banco.Text) & "','" & Trim(txt_nrocuenta.Text) & "','" & txt_beneficiario.Text & "'"

Exit Sub
Cargar_DatosErr:
    ErrorHandler err, "CARGA_GRID"
    Set rsd = Nothing
End Sub


Sub SALVAR_DATOS()
Dim rs As ADODB.Recordset
On Error GoTo Cargar_DatosErr

Set rs = New ADODB.Recordset
rs.ActiveConnection = cConnect
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic

rs.Open "exec UP_MAN_CLIENTE_BANCO 'U','" & Txt_Cod_Banco.Text & "','" & txt_nrocuenta.Text & "','" & txt_beneficiario.Text & "','" & Txt_Direccion.Text & "','" & Txt_Cod_Swift.Text & "','" & Mid(CmbEstado.Text, 1, 1) & "'"

Exit Sub
Cargar_DatosErr:
    ErrorHandler err, "CARGA_GRID"
    Set rs = Nothing
End Sub





