VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmShowTX_OrdComp_Ex 
   Caption         =   "Orden de Pedido de Exportación"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13425
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   13425
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Opc. Generales"
      Height          =   5685
      Left            =   10560
      TabIndex        =   9
      Top             =   1440
      Width           =   2805
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   510
         Left            =   -180
         ScaleHeight     =   510
         ScaleWidth      =   525
         TabIndex        =   15
         Top             =   6150
         Width           =   525
      End
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   5400
         Left            =   0
         TabIndex        =   35
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   9525
         Custom          =   $"frmShowTX_OrdComp_EX.frx":0000
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1200
         ControlHeigth   =   510
         ControlSeparator=   30
      End
      Begin FunctionsButtons.FunctButt FunctButt3 
         Height          =   5400
         Left            =   1320
         TabIndex        =   36
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   9525
         Custom          =   $"frmShowTX_OrdComp_EX.frx":03E6
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1200
         ControlHeigth   =   510
         ControlSeparator=   30
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6555
      Left            =   90
      TabIndex        =   7
      Top             =   1410
      Width           =   10470
      Begin TabDlg.SSTab SSTab1 
         Height          =   5295
         Left            =   0
         TabIndex        =   16
         Top             =   120
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   9340
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "General"
         TabPicture(0)   =   "frmShowTX_OrdComp_EX.frx":0772
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "GridEX1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "FrmCantidadxTela"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Tela"
         TabPicture(1)   =   "frmShowTX_OrdComp_EX.frx":078E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "GridEXTela"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Color"
         TabPicture(2)   =   "frmShowTX_OrdComp_EX.frx":07AA
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "GridEXColor"
         Tab(2).ControlCount=   1
         Begin VB.Frame FrmCantidadxTela 
            Caption         =   "Ingrese la Cantidad Total Por Tela"
            Height          =   1455
            Left            =   4920
            TabIndex        =   20
            Top             =   3000
            Visible         =   0   'False
            Width           =   4935
            Begin VB.TextBox Txt_CantidadXTela 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   840
               TabIndex        =   23
               Top             =   240
               Width           =   1770
            End
            Begin VB.CommandButton CmdAnadir2 
               Caption         =   "Añadir"
               Height          =   495
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   22
               Top             =   720
               Width           =   1365
            End
            Begin VB.CommandButton CmdCancelar 
               Caption         =   "Cancelar"
               Height          =   495
               Left            =   2520
               TabIndex        =   21
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label lblidclientekey 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   5040
               TabIndex        =   28
               Top             =   1560
               Width           =   1455
            End
            Begin VB.Label lblcodordcomp 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   5040
               TabIndex        =   27
               Top             =   1200
               Width           =   1215
            End
            Begin VB.Label lbl_serordcompMod 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   5040
               TabIndex        =   26
               Top             =   840
               Width           =   1335
            End
            Begin VB.Label lbl_codtelaMod 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   5040
               TabIndex        =   25
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label lbl_nombretela 
               Alignment       =   2  'Center
               Height          =   255
               Left            =   240
               TabIndex        =   24
               Top             =   1320
               Width           =   4455
            End
         End
         Begin GridEX20.GridEX GridEX1 
            Height          =   4755
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   9960
            _ExtentX        =   17568
            _ExtentY        =   8387
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            AllowEdit       =   0   'False
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            FormatStylesCount=   7
            FormatStyle(1)  =   "frmShowTX_OrdComp_EX.frx":07C6
            FormatStyle(2)  =   "frmShowTX_OrdComp_EX.frx":08FE
            FormatStyle(3)  =   "frmShowTX_OrdComp_EX.frx":09AE
            FormatStyle(4)  =   "frmShowTX_OrdComp_EX.frx":0A62
            FormatStyle(5)  =   "frmShowTX_OrdComp_EX.frx":0B3A
            FormatStyle(6)  =   "frmShowTX_OrdComp_EX.frx":0BF2
            FormatStyle(7)  =   "frmShowTX_OrdComp_EX.frx":0CD2
            ImageCount      =   0
            PrinterProperties=   "frmShowTX_OrdComp_EX.frx":0CF2
         End
         Begin GridEX20.GridEX GridEXTela 
            Height          =   4875
            Left            =   -74760
            TabIndex        =   18
            Top             =   360
            Width           =   9960
            _ExtentX        =   17568
            _ExtentY        =   8599
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            AllowEdit       =   0   'False
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            FormatStylesCount=   7
            FormatStyle(1)  =   "frmShowTX_OrdComp_EX.frx":0ECA
            FormatStyle(2)  =   "frmShowTX_OrdComp_EX.frx":1002
            FormatStyle(3)  =   "frmShowTX_OrdComp_EX.frx":10B2
            FormatStyle(4)  =   "frmShowTX_OrdComp_EX.frx":1166
            FormatStyle(5)  =   "frmShowTX_OrdComp_EX.frx":123E
            FormatStyle(6)  =   "frmShowTX_OrdComp_EX.frx":12F6
            FormatStyle(7)  =   "frmShowTX_OrdComp_EX.frx":13D6
            ImageCount      =   0
            PrinterProperties=   "frmShowTX_OrdComp_EX.frx":13F6
         End
         Begin GridEX20.GridEX GridEXColor 
            Height          =   4635
            Left            =   -74880
            TabIndex        =   19
            Top             =   480
            Width           =   10080
            _ExtentX        =   17780
            _ExtentY        =   8176
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            AllowEdit       =   0   'False
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            FormatStylesCount=   7
            FormatStyle(1)  =   "frmShowTX_OrdComp_EX.frx":15CE
            FormatStyle(2)  =   "frmShowTX_OrdComp_EX.frx":1706
            FormatStyle(3)  =   "frmShowTX_OrdComp_EX.frx":17B6
            FormatStyle(4)  =   "frmShowTX_OrdComp_EX.frx":186A
            FormatStyle(5)  =   "frmShowTX_OrdComp_EX.frx":1942
            FormatStyle(6)  =   "frmShowTX_OrdComp_EX.frx":19FA
            FormatStyle(7)  =   "frmShowTX_OrdComp_EX.frx":1ADA
            ImageCount      =   0
            PrinterProperties=   "frmShowTX_OrdComp_EX.frx":1AFA
         End
      End
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   510
         Left            =   120
         TabIndex        =   37
         Top             =   5760
         Width           =   9630
         _ExtentX        =   16986
         _ExtentY        =   900
         Custom          =   $"frmShowTX_OrdComp_EX.frx":1CD2
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   50
      End
      Begin VB.Label Label3 
         Caption         =   "Total Servicio"
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Label LBL_TOTALCARGA 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   375
         Left            =   1440
         TabIndex        =   33
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Label LBL_TOTAL_DETALLECARGADO 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   375
         Left            =   4560
         TabIndex        =   32
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Total Cargado"
         Height          =   375
         Left            =   3240
         TabIndex        =   31
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Label LBL_TOTALFALTANTE 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   375
         Left            =   7560
         TabIndex        =   30
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Faltante"
         Height          =   375
         Left            =   6600
         TabIndex        =   29
         Top             =   5520
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   1410
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13305
      Begin MSComCtl2.DTPicker DTPInicio 
         Height          =   300
         Left            =   6810
         TabIndex        =   12
         Top             =   1005
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         _Version        =   393216
         Format          =   74579969
         CurrentDate     =   37988
      End
      Begin VB.TextBox txtCod_OrdComp 
         Height          =   285
         Left            =   7335
         TabIndex        =   11
         Top             =   615
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox txtSer_OrdComp 
         Height          =   285
         Left            =   6795
         TabIndex        =   10
         Top             =   615
         Visible         =   0   'False
         Width           =   525
      End
      Begin FunctionsButtons.FunctButt FunctBuscar 
         Height          =   495
         Left            =   3720
         TabIndex        =   8
         Top             =   720
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~1~~0~Falso~Falso~&Buscar~"
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.OptionButton OptRango 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Rango Fechas"
         Height          =   270
         Left            =   5295
         TabIndex        =   6
         Top             =   1035
         Value           =   -1  'True
         Width           =   1440
      End
      Begin VB.OptionButton OptOC 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Orden Compra"
         Height          =   255
         Left            =   5295
         TabIndex        =   5
         Top             =   660
         Width           =   1395
      End
      Begin VB.OptionButton OptPendientes 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Pendientes"
         Height          =   255
         Left            =   5310
         TabIndex        =   4
         Top             =   285
         Width           =   1365
      End
      Begin VB.TextBox txtNom_Cliente 
         Height          =   315
         Left            =   1530
         TabIndex        =   3
         Top             =   285
         Width           =   3360
      End
      Begin VB.TextBox txtAbr_Cliente 
         Height          =   315
         Left            =   825
         TabIndex        =   2
         Top             =   285
         Width           =   690
      End
      Begin MSComCtl2.DTPicker DTPFin 
         Height          =   300
         Left            =   8340
         TabIndex        =   13
         Top             =   1005
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         _Version        =   393216
         Format          =   74579969
         CurrentDate     =   37988
      End
      Begin VB.Label Label2 
         Caption         =   "A"
         Height          =   195
         Left            =   8130
         TabIndex        =   14
         Top             =   1050
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cliente:"
         Height          =   210
         Left            =   165
         TabIndex        =   1
         Top             =   345
         Width           =   615
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   360
      Top             =   6675
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmShowTX_OrdComp_Ex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CODIGO As String
Public Descripcion As String
Public TipoAdd As String
Dim sSer_OrdComp_X As String
Dim sCod_OrdComp_X As String
Dim scod_cliente_tex_X As String
Dim strsql_x As String
Dim strSQL As String
Dim rscab As ADODB.Recordset
Dim rsx As ADODB.Recordset
Dim rsTela As ADODB.Recordset
'''111
Private Sub CmdAnadir2_Click()
Dim dCantidadxTela As Double, dCant_Pedida As Double

    'dCantidadxTela = IIf(Trim(Txt_CantidadXTela.Text) = "", 0, Txt_CantidadXTela.Text)
    'dCant_Pedida = IIf(Trim(txtCant_Pedida.Text) = "", 0, txtCant_Pedida.Text)
    'sSerie_Ordcomp = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Ser_OrdComp").Index)), "", GridEX1.Value(GridEX1.Columns("Ser_OrdComp").Index))
    'GridEXTela

        If Trim(Txt_CantidadXTela.Text) = "" Then
            MsgBox "Debe ingresar cantidad a editar", vbInformation, "Información"
            Exit Sub
        
        ElseIf CDbl(LBL_TOTAL_DETALLECARGADO.Caption) > CDbl(Txt_CantidadXTela.Text) Then
            MsgBox "La cantidad de Tela debe ser mayor a la cargada por colores", vbInformation, "Información"
            Exit Sub
        Else
        
            SALVAR_DATOS_TOTALXTELA
            BuscarTela (lblidclientekey.Caption)
            lbl_codtelaMod.Caption = ""
            lbl_serordcompMod.Caption = ""
            lblcodordcomp.Caption = ""
            lblidclientekey.Caption = ""
            FrmCantidadxTela.Visible = False
            SSTab1.Tab = 0
        End If
        Call CARGA_GRID
        
End Sub
'Private Sub CmdAnadir2_Click()
'Dim dCantidadxTela As Double, dCant_Pedida As Double
'
'    dCantidadxTela = IIf(Trim(Txt_CantidadXTela.Text) = "", 0, Txt_CantidadXTela.Text)
'    dCant_Pedida = IIf(Trim(txtCant_Pedida.Text) = "", 0, txtCant_Pedida.Text)
'
'    If dCantidadxTela < dCant_Pedida Then
'        MsgBox "La cantidad a Ingresar sobrepasa el total de la orden establecida para esta tela", vbCritical, Me.Caption
'        txtCant_Pedida.SetFocus
'        Exit Sub
'    Else
'            If VALIDA_DATOS Then
'                HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
'                DeshabilitaCampos
'                SALVAR_DATOS
'                CARGA_GRID
'                sTipo = ""
'                FrmCantidadxTela.Left = 12720
'                Txt_CantidadXTela.Text = ""
'            End If
'
'
'
'    End If
'End Sub

Sub SALVAR_DATOS_TOTALXTELA()
Dim i As Integer, strSQL As String
On Error GoTo hand

'Txt_CantidadXTela.Text = IIf(Trim(Txt_CantidadXTela.Text) = "", 0, Txt_CantidadXTela.Text)
            
            strSQL = "EXEC Usp_Upd_CantidaTotalxTela '" & lblidclientekey.Caption & "','" & _
            lbl_serordcompMod.Caption & "','" & _
            lblcodordcomp.Caption & "','" & _
            Trim(lbl_codtelaMod.Caption) & "'," & _
            Txt_CantidadXTela.Text & ""

            Call ExecuteSQL(cConnect, strSQL)
    
Exit Sub
hand:
    ErrorHandler err, "SALVAR_UPD_CANTIDADXTELA"
End Sub
Private Sub cmdCancelar_Click()
FrmCantidadxTela.Visible = False
End Sub

Private Sub Form_Load()
Me.DTPFin.Value = Date
Me.DTPInicio.Value = DateAdd("m", -1, Date)

Dim sSeguridad As String
sSeguridad = get_botones1(Me, vper, vemp1, Me.Name)
    
'Me.FunctButt1.FunctionsUser = sSeguridad
'Me.FunctButt2.FunctionsUser = sSeguridad
'Me.FunctButt3.FunctionsUser = sSeguridad
SSTab1.Tab = 0

End Sub
Sub muestraReqTelCru()
Dim SerOc As String, CodOc As String, codCliTex As String
Dim rstAux As ADODB.Recordset
On Error GoTo Fin


SerOc = GridEX1.Value(GridEX1.Columns("SER_ORDCOMP").Index)
CodOc = GridEX1.Value(GridEX1.Columns("COD_ORDCOMP").Index)
strSQL = "select cod_cliente_tex from tx_Cliente where abr_cliente='" & Trim(TxtAbr_Cliente.Text) & "'"
codCliTex = DevuelveCampo(strSQL, cConnect)


strSQL = "tj_muestra_requerimientos_cliente_textil_oc '" & codCliTex & "','" & SerOc & "','" & CodOc & "'"
   
    With frmBusqGeneral
        Set .oParent = Me
        .SQuery = strSQL
        .Cargar_Datos
        
        CODIGO = ".."
        Set rstAux = .gexList.ADORecordset
        .Show vbModal
  
        
'        If Codigo <> "" And rstAux.RecordCount > 0 Then
'            TxtCod_Cliente = Trim(rstAux!Codigo)
'            TxtDes_Cliente = Trim(rstAux!Descripcion)
'            TxtCod_Moneda.SetFocus
'        End If
        
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
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly
End Sub

Public Sub BUSCA_CLIENTE(tipo As Integer)
    Select Case tipo
        Case 1:
                    strSQL = "EXEC TI_BUSCA_CLIENTE 1,'" & Trim(Me.TxtAbr_Cliente.Text) & "','','" & vusu & "'"
                    Me.txtNom_Cliente.Text = Trim(DevuelveCampo(strSQL, cConnect))
                    If Trim(txtNom_Cliente.Text) <> "" Then CARGA_GRID
        Case 2, 3:
                    Dim oTipo As New frmBusGeneral6
                    Dim rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If tipo = 2 Then
                        oTipo.SQuery = "EXEC TI_BUSCA_CLIENTE 2,'','" & Trim(txtNom_Cliente.Text) & "','" & vusu & "'"
                    Else
                        oTipo.SQuery = "EXEC TI_BUSCA_CLIENTE 3,'','','" & vusu & "'"
                    End If
                    
                    oTipo.Cargar_Datos
                    oTipo.DGridLista.Columns(2).Width = 3500
                    oTipo.Show 1
                    If CODIGO <> "" Then
                         Me.TxtAbr_Cliente.Text = Trim(CODIGO)
                         Me.txtNom_Cliente.Text = Trim(Descripcion)
'                         OptCliPend.SetFocus
                         CODIGO = "": Descripcion = ""
                         CARGA_GRID
                    End If
                    Set oTipo = Nothing
                    Set rs = Nothing
    End Select
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub frmUpdate_OrdCompItem_Ex_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

End Sub

Private Sub FunctBuscar_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    If Trim(TxtAbr_Cliente.Text) = "" Then
        MsgBox "Seleccione el Cliente a Buscar", vbCritical, Me.Caption
        TxtAbr_Cliente.SetFocus
        Exit Sub
    End If
    CARGA_GRID
    
End Sub
Private Sub BuscarTela(ByVal vCod_Cliente_Tex As String)
On Error GoTo drDepurar

Dim sSQL As String, sSerie_Ordcomp As String, scod_ordcomp As String
Dim Rs_Lista As ADODB.Recordset
'Dim oGroup As GridEX20.JSGroup
'Dim oFormat As JSFormatStyle

sSerie_Ordcomp = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Ser_OrdComp").Index)), "", GridEX1.Value(GridEX1.Columns("Ser_OrdComp").Index))
scod_ordcomp = IIf(IsNull(GridEX1.Value(GridEX1.Columns("cod_OrdComp").Index)), "", GridEX1.Value(GridEX1.Columns("cod_OrdComp").Index))

sSQL = "EXEC Usp_Ver_TotalesPorTela '" & vCod_Cliente_Tex & "','" & Trim(sSerie_Ordcomp) & "','" & Trim(scod_ordcomp) & "'"

Set Rs_Lista = CargarRecordSetDesconectado(sSQL, cConnect)

'Set GridEXTela.ADORecordset = CargarRecordSetDesconectado(sSql, cConnect)
Set GridEXTela.ADORecordset = Rs_Lista
  

GridEXTela.Columns("Des_Tela").Caption = "Descripcion"
GridEXTela.Columns("Des_Tela").Width = 4500

GridEXTela.Columns("TotalxTela").Caption = "Cantidad"
GridEXTela.Columns("Ser_OrdComp").Caption = "Serie"
GridEXTela.Columns("Ser_OrdComp").Width = 800
GridEXTela.Columns("Cod_OrdComp").Caption = "Numero"
GridEXTela.Columns("Cod_OrdComp").Width = 1200
GridEXTela.Columns("Cod_Tela").Caption = "Codigo"
GridEXTela.Columns("COD_CLIENTE_TEX").Visible = False
'GridEXTela.Columns("Can_Pedida").Visible = False
GridEXTela.Columns("Can_Pedida").Caption = "Cargado"




Exit Sub
Resume
drDepurar:
  errores err.Number
End Sub
Sub CARGA_GRID()
Dim vOpcBusq As Integer
Dim vCod_Cliente_Tex As String, Strsql2 As String, Strsql3 As String
Dim vRowBookmark As Long
On Error GoTo hand
    If OptPendientes.Value Then
        vOpcBusq = 1
    ElseIf OptOC.Value Then
        vOpcBusq = 2
    Else
        vOpcBusq = 3
    End If
    
    
    strSQL = "select cod_cliente_tex from tx_Cliente where abr_cliente='" & Trim(TxtAbr_Cliente.Text) & "'"
    vCod_Cliente_Tex = DevuelveCampo(strSQL, cConnect)
    
    Strsql2 = "Select Isnull(CantidadTotal,0) From tx_ordcomp where cod_cliente_tex='" & vCod_Cliente_Tex & "' and ser_ordcomp='" & txtSer_OrdComp.Text & "' and cod_ordcomp='" & txtCod_OrdComp.Text & "'"
    LBL_TOTALCARGA.Caption = DevuelveCampo(Strsql2, cConnect)
    LBL_TOTALCARGA.Caption = IIf(LBL_TOTALCARGA.Caption = "", 0, LBL_TOTALCARGA.Caption)
    
    Strsql3 = "Select Isnull(sum(Can_Pedida),0) From tx_ordcompitem_tinto where cod_cliente_tex='" & vCod_Cliente_Tex & "' and ser_ordcomp='" & txtSer_OrdComp.Text & "' and cod_ordcomp='" & txtCod_OrdComp.Text & "'"
    LBL_TOTAL_DETALLECARGADO.Caption = DevuelveCampo(Strsql3, cConnect)
    LBL_TOTAL_DETALLECARGADO.Caption = IIf(LBL_TOTAL_DETALLECARGADO.Caption = "", 0, LBL_TOTAL_DETALLECARGADO.Caption)
    
    LBL_TOTALFALTANTE.Caption = CDbl(IIf(LBL_TOTALCARGA.Caption = "", 0, LBL_TOTALCARGA.Caption)) - CDbl(IIf(LBL_TOTAL_DETALLECARGADO.Caption = "", 0, LBL_TOTAL_DETALLECARGADO.Caption))
    
    strSQL = "EXEC TI_SEL_ORDCOMPITEM_TINTO_EXP " & vOpcBusq & ",'" & vCod_Cliente_Tex & "','" & Trim(txtSer_OrdComp.Text) & "','" & Trim(txtCod_OrdComp.Text) & "','" & DTPInicio.Value & "','" & DTPFin & "'"
    
    vRowBookmark = GridEX1.Row
    Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    ConfigurarGrid
    GridEX1.Row = vRowBookmark
    
   
Exit Sub
hand:
    ErrorHandler err, "Carga_Grid"
End Sub

Sub ConfigurarGrid()
    GridEX1.Columns("SEC_ORDCOMP").Caption = "Sec."

    GridEX1.Columns("SER_ORDCOMP").Visible = False
    GridEX1.Columns("Cod_Cliente_Tex").Visible = False
    GridEX1.Columns("COD_ORDCOMP").Visible = False
    GridEX1.Columns("COD_TELA").Visible = False
    'GridEX1.Columns("COD_comb").Visible = False
    GridEX1.Columns("cod_color").Visible = False
    GridEX1.Columns("COD_descuento").Visible = False
    strSQL = "select count(*) from TI_Seg_Acesso_Precios where cod_usuario='" & vusu & "'"
    If DevuelveCampo(strSQL, cConnect) > 0 Then
       GridEX1.Columns("p.u.").Visible = True
    Else
        GridEX1.Columns("p.u.").Visible = False
    End If
    
    GridEX1.Columns("SEC_ORDCOMP").Width = 700
    GridEX1.Columns("Color").Width = 2500
    GridEX1.Columns("Tela").Width = 3000
    GridEX1.Columns("oc").Width = 1200
    GridEX1.Columns("talla").Width = 700
    GridEX1.Columns("% igv").Width = 700
    GridEX1.Columns("PEDIDA").Width = 800
    GridEX1.Columns("DESPACHADA").Width = 1000
    GridEX1.Columns("Can_Despachada_Otros_Clientes").Width = 1000
    'GridEX1.Columns("Nro_Rollos_Despachados_Otros_Clientes").Width = 1000
        
    GridEX1.Columns("DEVUELTA").Width = 800
    GridEX1.Columns("P.U.").Width = 800
    
    GridEX1.FrozenColumns = 4
    
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim sFactPro As String



If GridEX1.RowCount = 0 Then Exit Sub

        Select Case ActionName
            Case "ADICIONAR"
                
            
                Load frmUpdate_OrdCompItem_Ex
                frmUpdate_OrdCompItem_Ex.cmdManTela.Enabled = True
                strSQL = "select cod_cliente_tex from tx_Cliente where abr_cliente='" & Trim(TxtAbr_Cliente.Text) & "'"
                frmUpdate_OrdCompItem_Ex.SCod_Cliente_Tex = DevuelveCampo(strSQL, cConnect)
                frmUpdate_OrdCompItem_Ex.sser_ordcomp = GridEX1.Value(GridEX1.Columns("SER_ORDCOMP").Index)
                frmUpdate_OrdCompItem_Ex.scod_ordcomp = GridEX1.Value(GridEX1.Columns("COD_ORDCOMP").Index)
                frmUpdate_OrdCompItem_Ex.CARGA_GRID
                Call frmUpdate_OrdCompItem_Ex.MantFunc1_ActionClick(1, 0, "ADICIONAR")
                'Call frmUpdate_OrdCompItem_Ex.MantFunc1_ActionClick(1, 0, frmUpdate_OrdCompItem_Ex.seg1)
                frmUpdate_OrdCompItem_Ex.Txtcod_Tela.TabIndex = 0
                strSQL = "select count(*) from TI_Seg_Acesso_Precios where cod_usuario='" & vusu & "'"
                If DevuelveCampo(strSQL, cConnect) = 0 Then
                    frmUpdate_OrdCompItem_Ex.txtPrecio.Visible = False
                    frmUpdate_OrdCompItem_Ex.Label21.Visible = False
                End If
                frmUpdate_OrdCompItem_Ex.Caption = "Detalle O.C. No: " & Trim(GridEX1.Value(GridEX1.Columns("oc").Index))
                frmUpdate_OrdCompItem_Ex.Show 1
                CARGA_GRID
                Set frmUpdate_OrdCompItem_Ex = Nothing
            Case "MODIFICAR"
                Load frmUpdate_OrdCompItem_Ex
                frmUpdate_OrdCompItem_Ex.cmdManTela.Enabled = True
                strSQL = "select cod_cliente_tex from tx_Cliente where abr_cliente='" & Trim(TxtAbr_Cliente.Text) & "'"
                frmUpdate_OrdCompItem_Ex.SCod_Cliente_Tex = DevuelveCampo(strSQL, cConnect)
                frmUpdate_OrdCompItem_Ex.sser_ordcomp = GridEX1.Value(GridEX1.Columns("SER_ORDCOMP").Index)
                frmUpdate_OrdCompItem_Ex.scod_ordcomp = GridEX1.Value(GridEX1.Columns("COD_ORDCOMP").Index)
                frmUpdate_OrdCompItem_Ex.CARGA_GRID
                Call frmUpdate_OrdCompItem_Ex.gexDetalle.Find(frmUpdate_OrdCompItem_Ex.gexDetalle.Columns("sec_ordcomp").Index, jgexContains, Me.GridEX1.Value(Me.GridEX1.Columns("sec_ordcomp").Index))
                Call frmUpdate_OrdCompItem_Ex.MantFunc1_ActionClick(2, 0, "MODIFICAR")
                frmUpdate_OrdCompItem_Ex.Txtcod_Tela.TabIndex = 0
                strSQL = "select count(*) from TI_Seg_Acesso_Precios where cod_usuario='" & vusu & "'"
                If DevuelveCampo(strSQL, cConnect) = 0 Then
                    frmUpdate_OrdCompItem_Ex.txtPrecio.Visible = False
                    frmUpdate_OrdCompItem_Ex.Label21.Visible = False
                End If
                frmUpdate_OrdCompItem_Ex.Caption = "Detalle O.C. No: " & Trim(GridEX1.Value(GridEX1.Columns("oc").Index))
                'frmUpdate_OrdCompItem_Ex.lbl_CanpedidaSinMod.Caption
                frmUpdate_OrdCompItem_Ex.sTipo = "U"
                frmUpdate_OrdCompItem_Ex.Show 1
                CARGA_GRID
            Case "ELIMINAR"
                Dim vMessage As Variant
                vMessage = MsgBox("Esta seguro que desea el registro selecionado", vbYesNo, "Eliminar")
                If vMessage = vbYes Then
                    ELIMINAR_DATOS
                End If
                CARGA_GRID
            Case "IMPRIMIR"
                    If GridEX1.RowCount = 0 Then Exit Sub
                    Reporte
            
            Case "DETALLECRUDO"
                Load frmDetalleCrudo_Ex
                strSQL = "select cod_cliente_tex from tx_Cliente where abr_cliente='" & Trim(TxtAbr_Cliente.Text) & "'"
                frmDetalleCrudo_Ex.SCod_Cliente_Tex = DevuelveCampo(strSQL, cConnect)
                frmDetalleCrudo_Ex.sser_ordcomp = GridEX1.Value(GridEX1.Columns("SER_ORDCOMP").Index)
                frmDetalleCrudo_Ex.scod_ordcomp = GridEX1.Value(GridEX1.Columns("COD_ORDCOMP").Index)
                frmDetalleCrudo_Ex.CARGA_GRID
                frmDetalleCrudo_Ex.Caption = "Detalle Crudo O.C.: " & Trim(GridEX1.Value(GridEX1.Columns("SER_ORDCOMP").Index)) & "-" & GridEX1.Value(GridEX1.Columns("COD_ORDCOMP").Index)
                frmDetalleCrudo_Ex.Show 1
                
                        
            Case "GUIASDESP"
                Load frmDetalleDespachos_Ex
                strSQL = "select cod_cliente_tex from tx_Cliente where abr_cliente='" & Trim(TxtAbr_Cliente.Text) & "'"
                frmDetalleDespachos_Ex.SCod_Cliente_Tex = DevuelveCampo(strSQL, cConnect)
                frmDetalleDespachos_Ex.sser_ordcomp = GridEX1.Value(GridEX1.Columns("SER_ORDCOMP").Index)
                frmDetalleDespachos_Ex.scod_ordcomp = GridEX1.Value(GridEX1.Columns("COD_ORDCOMP").Index)
                frmDetalleDespachos_Ex.ssec_ordcomp = Me.GridEX1.Value(Me.GridEX1.Columns("sec_ordcomp").Index)
                frmDetalleDespachos_Ex.CARGA_GRID
                frmDetalleDespachos_Ex.Caption = "Detalle Despachos O.C.: " & Trim(GridEX1.Value(GridEX1.Columns("SER_ORDCOMP").Index)) & "-" & GridEX1.Value(GridEX1.Columns("COD_ORDCOMP").Index) & " Sec.: " & Me.GridEX1.Value(Me.GridEX1.Columns("sec_ordcomp").Index)
                frmDetalleDespachos_Ex.Show 1
                
            Case "PARTIDAS"
                Load frmDetallePartidas_Ex
                strSQL = "select cod_cliente_tex from tx_Cliente where abr_cliente='" & Trim(TxtAbr_Cliente.Text) & "'"
                frmDetallePartidas_Ex.SCod_Cliente_Tex = DevuelveCampo(strSQL, cConnect)
                frmDetallePartidas_Ex.sser_ordcomp = GridEX1.Value(GridEX1.Columns("SER_ORDCOMP").Index)
                frmDetallePartidas_Ex.scod_ordcomp = GridEX1.Value(GridEX1.Columns("COD_ORDCOMP").Index)
                frmDetallePartidas_Ex.ssec_ordcomp = Me.GridEX1.Value(Me.GridEX1.Columns("sec_ordcomp").Index)
                frmDetallePartidas_Ex.CARGA_GRID
                frmDetallePartidas_Ex.Caption = "Detalle Partida O.C.: " & Trim(GridEX1.Value(GridEX1.Columns("SER_ORDCOMP").Index)) & "-" & GridEX1.Value(GridEX1.Columns("COD_ORDCOMP").Index) & " Sec.: " & Me.GridEX1.Value(Me.GridEX1.Columns("sec_ordcomp").Index)
                frmDetallePartidas_Ex.Show 1
          
            Case "CAMBIOESTADO"
                Dim vMessage2 As Variant
                vMessage2 = MsgBox("Esta seguro que desea cambiar el Estado de la O.C.", vbYesNo, "Cambio de Estado")
                If vMessage2 = vbYes Then
                    CAMBIARESTADO_OC_ITEM
                    
                End If
                CARGA_GRID
        End Select


End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ADICIONAROC"
            If Trim(TxtAbr_Cliente.Text) = "" Then
                MsgBox "Seleccione el Cliente a Buscar", vbCritical, Me.Caption
                TxtAbr_Cliente.SetFocus
                Exit Sub
            End If
            Load frmAdd_OrdComp_Ex
            frmAdd_OrdComp_Ex.TxtAbr_Cliente.Text = UCase(Me.TxtAbr_Cliente.Text)
            frmAdd_OrdComp_Ex.BUSCA_CLIENTE (1)
            frmAdd_OrdComp_Ex.BuscaLugEntr (1)
            
            frmAdd_OrdComp_Ex.Show 1
            CARGA_GRID
        Case "MODIFICAROC"
            If GridEX1.RowCount = 0 Then Exit Sub
            Load frmUpdate_OrdComp_EX
            strSQL = "select cod_cliente_tex from tx_Cliente where abr_cliente='" & Trim(TxtAbr_Cliente.Text) & "'"
            frmUpdate_OrdComp_EX.SCod_Cliente_Tex = DevuelveCampo(strSQL, cConnect)
            frmUpdate_OrdComp_EX.sser_ordcomp = GridEX1.Value(GridEX1.Columns("ser_ordcomp").Index)
            frmUpdate_OrdComp_EX.scod_ordcomp = GridEX1.Value(GridEX1.Columns("cod_ordcomp").Index)
            frmUpdate_OrdComp_EX.Caption = "Actulizar O.C. No: " & Trim(GridEX1.Value(GridEX1.Columns("oc").Index))
            frmUpdate_OrdComp_EX.Carga_Data
            frmUpdate_OrdComp_EX.Show 1
            CARGA_GRID
            
        Case "CAPTURAR"
            Load Frm_Caporden_Ex
            Frm_Caporden_Ex.sTipo = "1"
            Frm_Caporden_Ex.Show vbModal
            Set Frm_Caporden_Ex = Nothing
            
        Case "INKA"
            Load Frm_Caporden_Ex
            Frm_Caporden_Ex.sTipo = "2"
            Frm_Caporden_Ex.Show vbModal
            Set Frm_Caporden_Ex = Nothing
            
        Case "REVISION"
            If GridEX1.RowCount = 0 Then Exit Sub
                Revisar
        Case "CAMBIOESTADO"
            If GridEX1.RowCount = 0 Then Exit Sub
            Dim vMessage As Variant
            vMessage = MsgBox("Esta seguro que desea cambiar el Estado de la O.C.", vbYesNo, "Cambio de Estado")
            If vMessage = vbYes Then
                CAMBIARESTADO_OC
            End If
            CARGA_GRID
        Case "PROCESOS"
            If GridEX1.RowCount = 0 Then Exit Sub
                strSQL = "select cod_cliente_tex from tx_Cliente where abr_cliente='" & Trim(TxtAbr_Cliente.Text) & "'"
                Frm_Muestra_Procesos_Ex.SCod_Cliente_Tex = DevuelveCampo(strSQL, cConnect)
                Frm_Muestra_Procesos_Ex.Ser_OrdComp = GridEX1.Value(GridEX1.Columns("ser_ordcomp").Index)
                Frm_Muestra_Procesos_Ex.Cod_OrdComp = GridEX1.Value(GridEX1.Columns("cod_ordcomp").Index)
                Frm_Muestra_Procesos_Ex.Sec_OrdComp = GridEX1.Value(GridEX1.Columns("Sec_OrdComp").Index)
                Frm_Muestra_Procesos_Ex.Caption = "Mostrar O.C. No: " & Trim(GridEX1.Value(GridEX1.Columns("oc").Index))
                Frm_Muestra_Procesos_Ex.CARGA_GRID
                Frm_Muestra_Procesos_Ex.Show 1
                
                
        Case "PRECIO"
            If GridEX1.RowCount = 0 Then Exit Sub
            Load FrmShow_PrecioOCOtrosClientes_Ex
            FrmShow_PrecioOCOtrosClientes_Ex.scod_Cliente = GridEX1.Value(GridEX1.Columns("cod_cliente_tex").Index)
            FrmShow_PrecioOCOtrosClientes_Ex.sser_ordcomp = GridEX1.Value(GridEX1.Columns("ser_ordcomp").Index)
            FrmShow_PrecioOCOtrosClientes_Ex.scod_ordcomp = GridEX1.Value(GridEX1.Columns("cod_ordcomp").Index)
            FrmShow_PrecioOCOtrosClientes_Ex.ssec_ordcomp = GridEX1.Value(GridEX1.Columns("sec_ordcomp").Index)
            FrmShow_PrecioOCOtrosClientes_Ex.Txt_Tela = GridEX1.Value(GridEX1.Columns("Tela").Index)
            FrmShow_PrecioOCOtrosClientes_Ex.txt_talla = GridEX1.Value(GridEX1.Columns("talla").Index)
            FrmShow_PrecioOCOtrosClientes_Ex.txt_combo = GridEX1.Value(GridEX1.Columns("cod_comb").Index)
            FrmShow_PrecioOCOtrosClientes_Ex.txt_color = GridEX1.Value(GridEX1.Columns("Color").Index)
            FrmShow_PrecioOCOtrosClientes_Ex.LblOrden = GridEX1.Value(GridEX1.Columns("ser_ordcomp").Index) & "-" & GridEX1.Value(GridEX1.Columns("cod_ordcomp").Index)
            FrmShow_PrecioOCOtrosClientes_Ex.LblSecuencia = GridEX1.Value(GridEX1.Columns("sec_ordcomp").Index)
            FrmShow_PrecioOCOtrosClientes_Ex.LblCliente = DevuelveCampo("select abr_cliente  + '-' + nom_cliente from tx_cliente where cod_cliente_tex='" & GridEX1.Value(GridEX1.Columns("cod_cliente_tex").Index) & "'", cConnect)
            FrmShow_PrecioOCOtrosClientes_Ex.CARGA_GRID
            FrmShow_PrecioOCOtrosClientes_Ex.Show vbModal
            Set FrmShow_PrecioOCOtrosClientes_Ex = Nothing
        Case "IMPOCSERHIL"
            Call ReporteOCSERHIL
        Case "AVANTEXTIL"
            If GridEX1.RowCount = 0 Then Exit Sub
              Call ReporteAvanceTextil
        Case "ASIGNARPARTIDA"
            If GridEX1.RowCount = 0 Then Exit Sub
            Load FrmAsignaPartida
            FrmAsignaPartida.scod_Cliente = GridEX1.Value(GridEX1.Columns("cod_cliente_tex").Index)
            FrmAsignaPartida.sser_ordcomp = GridEX1.Value(GridEX1.Columns("ser_ordcomp").Index)
            FrmAsignaPartida.scod_ordcomp = GridEX1.Value(GridEX1.Columns("cod_ordcomp").Index)
            FrmAsignaPartida.ssec_ordcomp = GridEX1.Value(GridEX1.Columns("sec_ordcomp").Index)
            FrmAsignaPartida.Show vbModal
              
        Case "SALIR"
            Unload Me
    End Select
End Sub


Sub desautorizar_OC()
Dim i As Integer
Dim scod_Cliente As String
On Error GoTo hand

    strSQL = "select cod_cliente_tex from tx_Cliente where abr_cliente='" & Trim(TxtAbr_Cliente.Text) & "'"
    scod_Cliente = DevuelveCampo(strSQL, cConnect)
            
            strSQL = "EXEC HIL_MAN_ORDCOMP_DESAUTORIZAR '" & _
            scod_Cliente & "','" & _
            GridEX1.Value(GridEX1.Columns("ser_ordcomp").Index) & "','" & _
            GridEX1.Value(GridEX1.Columns("cod_ordcomp").Index) & "','" & vusu & "','" & ComputerName & "'"

            Call ExecuteSQL(cConnect, strSQL)
            
            MsgBox "Se Desautorizó correctamente"
    
Exit Sub
hand:
    ErrorHandler err, "CAMBIARESTADO_OC"
End Sub

Sub autorizar_OC()
Dim i As Integer
Dim scod_Cliente As String
On Error GoTo hand

    strSQL = "select cod_cliente_tex from tx_Cliente where abr_cliente='" & Trim(TxtAbr_Cliente.Text) & "'"
    scod_Cliente = DevuelveCampo(strSQL, cConnect)
            
            strSQL = "EXEC HIL_MAN_ORDCOMP_AUTORIZAR '" & _
            scod_Cliente & "','" & _
            GridEX1.Value(GridEX1.Columns("ser_ordcomp").Index) & "','" & _
            GridEX1.Value(GridEX1.Columns("cod_ordcomp").Index) & "','" & vusu & "','" & ComputerName & "'"

            Call ExecuteSQL(cConnect, strSQL)
            MsgBox "Se Autorizó correctamente"
    
Exit Sub
hand:
    ErrorHandler err, "CAMBIARESTADO_OC"
End Sub

Sub CAMBIARESTADO_OC()
Dim i As Integer
Dim scod_Cliente As String
On Error GoTo hand

    strSQL = "select cod_cliente_tex from tx_Cliente where abr_cliente='" & Trim(TxtAbr_Cliente.Text) & "'"
    scod_Cliente = DevuelveCampo(strSQL, cConnect)
            
            strSQL = "EXEC TI_UP_CAMBIA_ESTADO_ORDEN_COMPRA_TINTO '" & _
            scod_Cliente & "','" & _
            GridEX1.Value(GridEX1.Columns("ser_ordcomp").Index) & "','" & _
            GridEX1.Value(GridEX1.Columns("cod_ordcomp").Index) & "'"

            Call ExecuteSQL(cConnect, strSQL)
    
Exit Sub
hand:
    ErrorHandler err, "CAMBIARESTADO_OC"
End Sub

Sub Revisar()
Dim i As Integer
Dim scod_Cliente As String
On Error GoTo hand

    strSQL = "select cod_cliente_tex from tx_Cliente where abr_cliente='" & Trim(TxtAbr_Cliente.Text) & "'"
    scod_Cliente = DevuelveCampo(strSQL, cConnect)
            
            strSQL = "EXEC ti_revisa_status_colores_oc_confecciones '" & _
            scod_Cliente & "','" & _
            GridEX1.Value(GridEX1.Columns("ser_ordcomp").Index) & "','" & _
            GridEX1.Value(GridEX1.Columns("cod_ordcomp").Index) & "'"
           

            Call ExecuteSQL(cConnect, strSQL)
    
    MsgBox "REVISION TERMINADA OK."
    
Exit Sub
hand:
    ErrorHandler err, "Revisar Orden De Compra"
End Sub

Sub ELIMINAR_DATOS()
Dim i As Integer
Dim scod_Cliente As String
On Error GoTo hand

    strSQL = "select cod_cliente_tex from tx_Cliente where abr_cliente='" & Trim(TxtAbr_Cliente.Text) & "'"
    scod_Cliente = DevuelveCampo(strSQL, cConnect)
            
            strSQL = "EXEC TI_MAN_TX_ORDCOMPITEM_TINTO_EX 'D','" & _
            scod_Cliente & "','" & _
            GridEX1.Value(GridEX1.Columns("ser_ordcomp").Index) & "','" & _
            GridEX1.Value(GridEX1.Columns("cod_ordcomp").Index) & "','" & _
            GridEX1.Value(GridEX1.Columns("sec_ordcomp").Index) & "','','','','','','',0,null,null,0,0,'',''"

            Call ExecuteSQL(cConnect, strSQL)
    
Exit Sub
hand:
    MsgBox err.Description
    'ErrorHandler Err, "ElIMINAR_DATOS"
End Sub

Sub Revision()
Dim i As Integer
Dim scod_Cliente As String
On Error GoTo hand
           
    strSQL = "EXEC ti_revisa_status_colores_oc_confecciones '" & GridEX1.Value(GridEX1.Columns("ser_ordcomp").Index) & _
    "','" & GridEX1.Value(GridEX1.Columns("cod_ordcomp").Index) & "'"
    Call ExecuteSQL(cConnect, strSQL)
    
Exit Sub
hand:
    ErrorHandler err, "Capturar Datos"
End Sub



Private Sub FunctButt3_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo xerror:
            strSQL = "select cod_cliente_tex from tx_Cliente where abr_cliente='" & Trim(TxtAbr_Cliente.Text) & "'"
            scod_cliente_tex_X = DevuelveCampo(strSQL, cConnect)
            
  Select Case ActionName
              Dim vMessage As Variant, sIdFacturaProforma As String
        Case "AUTORIZAR"
            If GridEX1.RowCount = 0 Then Exit Sub
            vMessage = MsgBox("Esta seguro que desea autorizar la O.C.", vbYesNo, "Autorizar")
            If vMessage = vbYes Then
                Call autorizar_OC
                
            End If
            CARGA_GRID
        Case "DESAUTORIZAR"
            If GridEX1.RowCount = 0 Then Exit Sub
            vMessage = MsgBox("Esta seguro que desea desautorizar la O.C.", vbYesNo, "Autorizar")
            If vMessage = vbYes Then
                Call desautorizar_OC
                
            End If
            CARGA_GRID
        Case "VERREQTELCRU"
            Call muestraReqTelCru
            'tj_muestra_requerimientos_cliente_textil_oc '00008','999','999999'
            CARGA_GRID
            
        Case "Imp_OS"
            sSer_OrdComp_X = ""
            sCod_OrdComp_X = ""
            strSQL = "select cod_cliente_tex from tx_Cliente where abr_cliente='" & Trim(TxtAbr_Cliente.Text) & "'"
            scod_cliente_tex_X = DevuelveCampo(strSQL, cConnect)
            If GridEX1.RowCount <> 0 Then
                sSer_OrdComp_X = GridEX1.Value(GridEX1.Columns("ser_ordcomp").Index)
                sCod_OrdComp_X = GridEX1.Value(GridEX1.Columns("cod_ordcomp").Index)
                Call validar_Impresion_OS
            Else
                MsgBox "No se ha seleccionado ninguna Orden de compra", vbInformation, "Orden de Compra"
                Exit Sub
            End If
            
        Case "PENDIENTE"
            Load frmPendiente_Ex
            frmPendiente_Ex.CARGA_GRID
            frmPendiente_Ex.Show 1
        Case "FACTURAPROFORMA"
            sSer_OrdComp_X = ""
            sCod_OrdComp_X = ""
            
            sIdFacturaProforma = IIf(IsNull(GridEX1.Value(GridEX1.Columns("IdFacturaProforma").Index)) = True, "", GridEX1.Value(GridEX1.Columns("IdFacturaProforma").Index))
            sSer_OrdComp_X = GridEX1.Value(GridEX1.Columns("ser_ordcomp").Index)
            sCod_OrdComp_X = GridEX1.Value(GridEX1.Columns("cod_ordcomp").Index)
            
            If Trim(sIdFacturaProforma) <> "" Then
            
                MsgBox "Ya existe una Factura Proforma para la Orden de servicio de Exportación N°" + sSer_OrdComp_X + "-" + sCod_OrdComp_X, vbInformation, "Información"
                Exit Sub
            
                
            Else
                    If MsgBox("Desea Generar la Factura Proforma", vbQuestion + vbYesNo, "Pregunta") = vbYes Then
                        Call GeneraFacturaProforma
                    End If
            End If
            
        Case "MODTOTTELA"
                'lbl_nombretela.Caption = IIf(IsNull(GridEXTela.Value(GridEXTela.Columns("Des_Tela").Index)), "", GridEXTela.Value(GridEXTela.Columns("Des_Tela").Index))
                'Txt_CantidadXTela.Text = IIf(IsNull(GridEXTela.Value(GridEXTela.Columns("TotalxTela").Index)), "", GridEXTela.Value(GridEXTela.Columns("TotalxTela").Index))
                
                'lbl_codtelaMod.Caption = IIf(IsNull(GridEXTela.Value(GridEXTela.Columns("Cod_Tela").Index)), "", GridEXTela.Value(GridEXTela.Columns("Cod_Tela").Index))
                'lbl_serordcompMod.Caption = IIf(IsNull(GridEXTela.Value(GridEXTela.Columns("Ser_OrdComp").Index)), "", GridEXTela.Value(GridEXTela.Columns("Ser_OrdComp").Index))
                'lblcodordcomp.Caption = IIf(IsNull(GridEXTela.Value(GridEXTela.Columns("Cod_OrdComp").Index)), "", GridEXTela.Value(GridEXTela.Columns("Cod_OrdComp").Index))
                'lblidclientekey.Caption = IIf(IsNull(GridEXTela.Value(GridEXTela.Columns("cod_cliente_tex").Index)), "", GridEXTela.Value(GridEXTela.Columns("cod_cliente_tex").Index))
                
                lbl_nombretela.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("tela").Index)), "", GridEX1.Value(GridEX1.Columns("tela").Index))
                Txt_CantidadXTela.Text = IIf(IsNull(GridEX1.Value(GridEX1.Columns("TOTALXTELA").Index)), "", GridEX1.Value(GridEX1.Columns("TOTALXTELA").Index))
                
                lbl_codtelaMod.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Cod_Tela").Index)), "", GridEX1.Value(GridEX1.Columns("Cod_Tela").Index))
                lbl_serordcompMod.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Ser_OrdComp").Index)), "", GridEX1.Value(GridEX1.Columns("Ser_OrdComp").Index))
                lblcodordcomp.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("Cod_OrdComp").Index)), "", GridEX1.Value(GridEX1.Columns("Cod_OrdComp").Index))
                lblidclientekey.Caption = IIf(IsNull(GridEX1.Value(GridEX1.Columns("cod_cliente_tex").Index)), "", GridEX1.Value(GridEX1.Columns("cod_cliente_tex").Index))
                
                
                If Trim(lbl_nombretela.Caption) = "" Then
                
                    MsgBox "Debe seleccionar la tela a editar la cantidad", vbInformation, "Información"
                    Exit Sub
        
                    
                Else
                    FrmCantidadxTela.Visible = True
                End If
        
        
'                Load frmUpdate_OrdCompItem
'                frmUpdate_OrdCompItem.cmdManTela.Enabled = True
'                StrSQL = "select cod_cliente_tex from tx_Cliente where abr_cliente='" & Trim(TxtAbr_Cliente.Text) & "'"
'                frmUpdate_OrdCompItem.SCod_Cliente_Tex = DevuelveCampo(StrSQL, cConnect)
'                frmUpdate_OrdCompItem.sser_ordcomp = GridEX1.Value(GridEX1.Columns("SER_ORDCOMP").Index)
'                frmUpdate_OrdCompItem.scod_ordcomp = GridEX1.Value(GridEX1.Columns("COD_ORDCOMP").Index)
'                frmUpdate_OrdCompItem.CARGA_GRID
'                Call frmUpdate_OrdCompItem.gexDetalle.Find(frmUpdate_OrdCompItem.gexDetalle.Columns("sec_ordcomp").Index, jgexContains, Me.GridEX1.Value(Me.GridEX1.Columns("sec_ordcomp").Index))
'                Call frmUpdate_OrdCompItem.MantFunc1_ActionClick(2, 0, "MODIFICAR")
'                frmUpdate_OrdCompItem.TxtCod_Tela.TabIndex = 0
'
''                StrSQL = "select count(*) from TI_Seg_Acesso_Precios where cod_usuario='" & vusu & "'"
''                If DevuelveCampo(StrSQL, cConnect) = 0 Then
''                    frmUpdate_OrdCompItem.txtPrecio.Visible = False
''                    frmUpdate_OrdCompItem.Label21.Visible = False
''                End If
'
'                frmUpdate_OrdCompItem.Caption = "Detalle O.C. No: " & Trim(GridEX1.Value(GridEX1.Columns("oc").Index))
'                frmUpdate_OrdCompItem.Show 1
'                CARGA_GRID
'
            
        Case "SALIR"
            Unload Me
    Case "Reg_datos"
            sSer_OrdComp_X = ""
            sCod_OrdComp_X = ""
            sSer_OrdComp_X = GridEX1.Value(GridEX1.Columns("ser_ordcomp").Index)
            sCod_OrdComp_X = GridEX1.Value(GridEX1.Columns("cod_ordcomp").Index)
        If sSer_OrdComp_X = "" Or sCod_OrdComp_X = "" Then
            MsgBox "Debe seleccionar una orden de Compra", vbInformation, "Mensaje"
            Exit Sub
        End If
        If scod_cliente_tex_X = "" Then
            MsgBox "Debe seleccionar un cliente", vbInformation, "Mensaje"
            TxtAbr_Cliente.SetFocus
            Exit Sub
        End If
        
        With Frm_Confirmar_Cliente_Expo
            .Cod_Cliente_TexX = scod_cliente_tex_X
            .cod_ORdCompX = sCod_OrdComp_X
            .Ser_ORdCompX = sSer_OrdComp_X
            .Carga_CAmpos
            .Show 1
        End With
    Case "Reg_Fact"
            sSer_OrdComp_X = ""
            sCod_OrdComp_X = ""
            sSer_OrdComp_X = GridEX1.Value(GridEX1.Columns("ser_ordcomp").Index)
            sCod_OrdComp_X = GridEX1.Value(GridEX1.Columns("cod_ordcomp").Index)
        If sSer_OrdComp_X = "" Or sCod_OrdComp_X = "" Then
            MsgBox "Debe seleccionar una orden de Compra", vbInformation, "Mensaje"
            Exit Sub
        End If
        If scod_cliente_tex_X = "" Then
            MsgBox "Debe seleccionar un cliente", vbInformation, "Mensaje"
            TxtAbr_Cliente.SetFocus
            Exit Sub
        End If
        
        With Frm_Confirmar_Cliente_Facturar_Expo
            .Cod_Cliente_TexX = scod_cliente_tex_X
            .cod_ORdCompX = sCod_OrdComp_X
            .Ser_ORdCompX = sSer_OrdComp_X
            .Carga_CAmpos
            .Show 1
        End With
    
    
    
    End Select
Exit Sub
xerror:
ErrorHandler err, "Mensaje"
End Sub
Private Sub GeneraFacturaProforma()
Dim i As Integer
Dim scod_Cliente As String, sser_ordcomp As String, scod_ordcomp As String
On Error GoTo hand

            strSQL = "select cod_cliente_tex from tx_Cliente where abr_cliente='" & Trim(TxtAbr_Cliente.Text) & "'"
            scod_Cliente = DevuelveCampo(strSQL, cConnect)
            sser_ordcomp = GridEX1.Value(GridEX1.Columns("ser_ordcomp").Index)
            scod_ordcomp = GridEX1.Value(GridEX1.Columns("cod_ordcomp").Index)

            
            strSQL = "EXEC Usp_Genera_FacturaProforma '" & scod_Cliente & "','" & sser_ordcomp & "','" & scod_ordcomp & "'"

            Call ExecuteSQL(cConnect, strSQL)
            MsgBox "Se Genero la Factura Proforma correctamente"
    
Exit Sub
hand:
    ErrorHandler err, "CAMBIARESTADO_OC"
End Sub
Private Sub GridEX1_Click()
Dim strSQL As String, sidcliente As String
strSQL = "select cod_cliente_tex from tx_Cliente where abr_cliente='" & Trim(TxtAbr_Cliente.Text) & "'"
sidcliente = DevuelveCampo(strSQL, cConnect)
Call BuscarTela(sidcliente)
End Sub

Private Sub GridEXTela_Click()
On Error GoTo drDepurar

Dim sSQL As String, sSerie_Ordcomp As String, scod_ordcomp As String, sCod_Tela As String, sidcliente As String
Dim oGroup As GridEX20.JSGroup
Dim oFormat As JSFormatStyle

sSerie_Ordcomp = IIf(IsNull(GridEXTela.Value(GridEXTela.Columns("Ser_OrdComp").Index)), "", GridEXTela.Value(GridEXTela.Columns("Ser_OrdComp").Index))
scod_ordcomp = IIf(IsNull(GridEXTela.Value(GridEXTela.Columns("cod_OrdComp").Index)), "", GridEXTela.Value(GridEXTela.Columns("cod_OrdComp").Index))
sCod_Tela = IIf(IsNull(GridEXTela.Value(GridEXTela.Columns("Cod_Tela").Index)), "", GridEXTela.Value(GridEXTela.Columns("Cod_Tela").Index))
sidcliente = IIf(IsNull(GridEXTela.Value(GridEXTela.Columns("COD_CLIENTE_TEX").Index)), "", GridEXTela.Value(GridEXTela.Columns("COD_CLIENTE_TEX").Index))

LBL_TOTALCARGA.Caption = IIf(IsNull(GridEXTela.Value(GridEXTela.Columns("TotalxTela").Index)), "0", GridEXTela.Value(GridEXTela.Columns("TotalxTela").Index))
LBL_TOTAL_DETALLECARGADO.Caption = IIf(IsNull(GridEXTela.Value(GridEXTela.Columns("Can_Pedida").Index)), "0", GridEXTela.Value(GridEXTela.Columns("Can_Pedida").Index))

If IsNumeric(LBL_TOTALCARGA.Caption) And IsNumeric(LBL_TOTAL_DETALLECARGADO.Caption) Then
    LBL_TOTALFALTANTE.Caption = CDbl(LBL_TOTALCARGA.Caption) - CDbl(LBL_TOTAL_DETALLECARGADO.Caption)
Else
   LBL_TOTALFALTANTE.Caption = 0
End If

sSQL = "EXEC Usp_Ver_TotalesPorTelaPorColor '" & sidcliente & "','" & Trim(sSerie_Ordcomp) & "','" & Trim(scod_ordcomp) & "','" & sCod_Tela & "'"


Set GridEXColor.ADORecordset = CargarRecordSetDesconectado(sSQL, cConnect)
  

GridEXColor.Columns("Des_Color").Caption = "Descripcion"
GridEXColor.Columns("Des_Color").Width = 4500

GridEXColor.Columns("Pre_Unitario").Caption = "Precio"
GridEXColor.Columns("Pre_Unitario").Width = 900
GridEXColor.Columns("Can_Pedida").Caption = "Cantidad"
GridEXColor.Columns("Can_Pedida").Width = 900


Exit Sub
Resume
drDepurar:
  errores err.Number

End Sub

Private Sub OptOC_Click()
    txtSer_OrdComp.Visible = True
    txtCod_OrdComp.Visible = True
    txtSer_OrdComp.Text = ""
    txtCod_OrdComp.Text = ""
    
    DTPInicio.Visible = False
    DTPFin.Visible = False
    Label2.Visible = False
End Sub

Private Sub optPendientes_Click()
    txtSer_OrdComp.Visible = False
    txtCod_OrdComp.Visible = False
    
    DTPInicio.Visible = False
    DTPFin.Visible = False
    Label2.Visible = False
End Sub

Private Sub OptRango_Click()
    txtSer_OrdComp.Visible = False
    txtCod_OrdComp.Visible = False
    
    DTPInicio.Visible = True
    DTPFin.Visible = True
    Label2.Visible = True
    
    DTPInicio.Value = Date
    DTPFin.Value = DateAdd("m", 1, Format(DTPInicio.Value, "dd/mm/yyyy"))
End Sub

Private Sub txtAbr_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(TxtAbr_Cliente.Text) = "" Then
            Call BUSCA_CLIENTE(3)
        Else
            Call BUSCA_CLIENTE(1)
        End If
    End If
End Sub

Private Sub txtCod_OrdComp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FunctBuscar.SetFocus
    Else
        Call SoloNumeros(txtCod_OrdComp, KeyAscii, False, 0, 6)
    End If
End Sub

Private Sub txtCod_OrdComp_LostFocus()
    txtCod_OrdComp.Text = Format(Trim(txtCod_OrdComp.Text), "000000")
End Sub

Private Sub txtNom_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtNom_Cliente.Text) = "" Then
            Call BUSCA_CLIENTE(3)
        Else
            Call BUSCA_CLIENTE(2)
        End If
    End If
End Sub

Private Sub txtSer_OrdComp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCod_OrdComp.SetFocus
    Else
        Call SoloNumeros(txtSer_OrdComp, KeyAscii, False, 0, 3)
    End If
End Sub

Private Sub txtSer_OrdComp_LostFocus()
    txtSer_OrdComp.Text = Format(Trim(txtSer_OrdComp.Text), "000")
End Sub


Sub CAMBIARESTADO_OC_ITEM()
Dim i As Integer
Dim scod_Cliente As String
On Error GoTo hand

    strSQL = "select cod_cliente_tex from tx_Cliente where abr_cliente='" & Trim(TxtAbr_Cliente.Text) & "'"
    scod_Cliente = DevuelveCampo(strSQL, cConnect)
            
            strSQL = "EXEC TI_UP_CAMBIA_ESTADO_ORDEN_COMPRA_ITEM_TINTO '" & _
            scod_Cliente & "','" & _
            GridEX1.Value(GridEX1.Columns("ser_ordcomp").Index) & "','" & _
            GridEX1.Value(GridEX1.Columns("cod_ordcomp").Index) & "','" & _
            GridEX1.Value(GridEX1.Columns("SEC_ordcomp").Index) & "'"
            Call ExecuteSQL(cConnect, strSQL)
    
Exit Sub
hand:
    ErrorHandler err, "CAMBIARESTADO_OC_ITEM"
End Sub

Private Sub ReporteAvanceTextil()
On Error GoTo Fin

Dim oo As Object
    
    Screen.MousePointer = 11
    
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\AvanceTextil.XLT"
    oo.DisplayAlerts = False
    oo.Visible = True
    Dim vCod_Cliente_Tex  As String
    strSQL = "select cod_cliente_tex from tx_Cliente where abr_cliente='" & Trim(TxtAbr_Cliente.Text) & "'"
    vCod_Cliente_Tex = DevuelveCampo(strSQL, cConnect)

    oo.Run "Reporte", CStr(vCod_Cliente_Tex), CStr(GridEX1.Value(GridEX1.Columns("ser_ordcomp").Index)), CStr(GridEX1.Value(GridEX1.Columns("cod_ordcomp").Index)), cConnect
    Screen.MousePointer = vbNormal
    Set oo = Nothing
Exit Sub
Fin:
    Screen.MousePointer = vbNormal
    errores err.Number
End Sub


Private Sub ReporteOCSERHIL()
On Error GoTo Fin
    Dim oo As Object
    Dim strSQL As String
    Dim sempresa As String
    
    strSQL = "SELECT DES_EMPRESA FROM SEGURIDAD..SEG_EMPRESAS WHERE COD_EMPRESA='" & vemp & "'"
    sempresa = DevuelveCampo(strSQL, cConnect)
    
    Screen.MousePointer = 11
    
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open vRuta & "\OrdSerHilado.XLT"
    oo.DisplayAlerts = False
    oo.Visible = True
    Dim vCod_Cliente_Tex  As String
    strSQL = "select cod_cliente_tex from tx_Cliente where abr_cliente='" & Trim(TxtAbr_Cliente.Text) & "'"
    vCod_Cliente_Tex = DevuelveCampo(strSQL, cConnect)

    oo.Run "Reporte", CStr(vCod_Cliente_Tex), CStr(GridEX1.Value(GridEX1.Columns("ser_ordcomp").Index)), CStr(GridEX1.Value(GridEX1.Columns("cod_ordcomp").Index)), cConnect, sempresa
    Screen.MousePointer = vbNormal
    Set oo = Nothing
Exit Sub
Fin:
    Screen.MousePointer = vbNormal
    errores err.Number
End Sub
Private Sub Reporte()
On Error GoTo Fin

Dim oo As Object, vRutaLogo As Variant
    
    Screen.MousePointer = 11
    strSQL = "SELECT Ruta_Logo From SEGURIDAD..SEG_EMPRESAS " & _
             "WHERE Cod_Empresa = '" & vemp & "'"
    vRutaLogo = DevuelveCampo(strSQL, cConnect)
    vRutaLogo = CStr(IIf(IsNull(vRutaLogo), "", vRutaLogo))
          Set oo = CreateObject("excel.application")
          oo.Workbooks.Open vRuta & "\Rpt_RordenCompra.XLT"
          oo.DisplayAlerts = False
          oo.Visible = True
    
    oo.Run "REPORTE", GridEX1.ADORecordset, txtNom_Cliente.Text
    
    Screen.MousePointer = vbNormal
    'oo.Workbooks.Close
    Set oo = Nothing
Exit Sub
Fin:
    Screen.MousePointer = vbNormal
    errores err.Number
End Sub
Sub validar_Impresion_OS()
On Error GoTo xerror:
If scod_cliente_tex_X = "" Then
    MsgBox "No se ha ingresado un Cliente Textil ", vbInformation, "Orden de Compra"
    Exit Sub
End If

If sSer_OrdComp_X = "" Or sSer_OrdComp_X = "" Then
    MsgBox "No ha ingresado la OC", vbInformation, "Orden de compra"
    Exit Sub
End If

Call Carga_Reporte
Exit Sub
xerror:
   ErrorHandler err, "Orden de compra"
    Exit Sub
End Sub
Sub Carga_Reporte()
'On Error GoTo xError:
Dim oo As Object
Set oo = CreateObject("Excel.application")
Screen.MousePointer = 0
oo.Visible = True
oo.DisplayAlerts = False
oo.Workbooks.Open vRuta & "\RptOServicio_Tintoreria_Exportacion.xlt"
strsql_x = "Exec Cabecera_Orden_Servicio_Tinto_Exp '" & scod_cliente_tex_X & "','" & sSer_OrdComp_X & "','" & sCod_OrdComp_X & "'"
Set rscab = CargarRecordSetDesconectado(strsql_x, cConnect)
strsql_x = "Exec DETALLE_ORDEN_SERVICIO_TINTO '" & scod_cliente_tex_X & "','" & sSer_OrdComp_X & "','" & sCod_OrdComp_X & "'"
Set rsx = CargarRecordSetDesconectado(strsql_x, cConnect)
strsql_x = "Exec Detalle_Tela_Orden_Servicio '" & scod_cliente_tex_X & "','" & sSer_OrdComp_X & "','" & sCod_OrdComp_X & "'"
Set rsTela = CargarRecordSetDesconectado(strsql_x, cConnect)
'MsgBox rscab.RecordCount
oo.Run "Cabecera_Orden_Servicio", sSer_OrdComp_X, sCod_OrdComp_X, rscab, cConnect
oo.Run "Detalle_Orden_servicio", rsx, rsTela, cConnect, vemp1
oo.Run "Detalle_Tela", rsTela, cConnect
Set oo = Nothing
Screen.MousePointer = 0
Exit Sub
'xError:
'   Screen.MousePointer = 0
  '  ErrorHandler Err, "Reporte Orden de Servicio"
   ' Set oo = Nothing
End Sub
