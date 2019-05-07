VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMuestraStocksFechas 
   Caption         =   "Muestra Stock por Rango de Fechas"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14595
   LinkTopic       =   "Form1"
   ScaleHeight     =   9060
   ScaleWidth      =   14595
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   1605
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   14640
      Begin VB.ComboBox cboAlmacen 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   600
         Width           =   2955
      End
      Begin VB.OptionButton optTodos 
         BackColor       =   &H00FFC0C0&
         Caption         =   "STOCK GENERAL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3240
         TabIndex        =   12
         Top             =   120
         Value           =   -1  'True
         Width           =   2115
      End
      Begin VB.OptionButton optEstCli 
         BackColor       =   &H00FFC0C0&
         Caption         =   "STOCK POR ESTILO CLIENTE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6000
         TabIndex        =   11
         Top             =   120
         Width           =   3315
      End
      Begin VB.Frame FraOpciones 
         BackColor       =   &H00FFC0C0&
         Height          =   615
         Left            =   3240
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   6135
         Begin VB.TextBox TxtCodigo 
            Height          =   300
            Left            =   570
            TabIndex        =   9
            Top             =   195
            Width           =   1100
         End
         Begin VB.TextBox TxtDescripcion 
            Height          =   300
            Left            =   1830
            TabIndex        =   8
            Top             =   195
            Width           =   4065
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Dato"
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
            Left            =   105
            TabIndex        =   10
            Top             =   180
            Width           =   360
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         Height          =   615
         Left            =   9480
         TabIndex        =   3
         Top             =   360
         Width           =   3615
         Begin VB.CheckBox chkColor 
            BackColor       =   &H00FFC0C0&
            Caption         =   "MOSTRAR COLOR"
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
            Left            =   1800
            TabIndex        =   6
            Top             =   120
            Width           =   1695
         End
         Begin VB.CheckBox chkTalla 
            BackColor       =   &H00FFC0C0&
            Caption         =   "MOSTRAR TALLA"
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
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   1695
         End
         Begin VB.CheckBox chkImagen 
            BackColor       =   &H00FFC0C0&
            Caption         =   "MOSTRAR IMAGEN"
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
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   1815
         End
      End
      Begin FunctionsButtons.FunctButt fnbBuscar 
         Height          =   495
         Left            =   13200
         TabIndex        =   14
         Top             =   480
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
      Begin MSComCtl2.DTPicker DTPInicio 
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   503
         _Version        =   393216
         Format          =   71041025
         CurrentDate     =   38182
      End
      Begin MSComCtl2.DTPicker DTPHasta 
         Height          =   270
         Left            =   3240
         TabIndex        =   18
         Top             =   1200
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   476
         _Version        =   393216
         Format          =   71041025
         CurrentDate     =   38182
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "DESDE:"
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
         TabIndex        =   20
         Top             =   1035
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "HASTA:"
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
         Left            =   3240
         TabIndex        =   19
         Top             =   1005
         Width           =   585
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ALMACEN:"
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
         Left            =   135
         TabIndex        =   15
         Top             =   180
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdVerDetMov 
      Caption         =   "Ver Detalle Movimientos"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   8340
      Width           =   1695
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   11865
      TabIndex        =   1
      Top             =   8340
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmMuestraStocksFechas.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   6540
      Left            =   0
      TabIndex        =   16
      Top             =   1695
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   11536
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ContScroll      =   -1  'True
      AllowColumnDrag =   0   'False
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      BackColorBkg    =   12648384
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "FrmMuestraStocksFechas.frx":0145
      Column(2)       =   "FrmMuestraStocksFechas.frx":020D
      FormatStylesCount=   6
      FormatStyle(1)  =   "FrmMuestraStocksFechas.frx":02B1
      FormatStyle(2)  =   "FrmMuestraStocksFechas.frx":03E9
      FormatStyle(3)  =   "FrmMuestraStocksFechas.frx":0499
      FormatStyle(4)  =   "FrmMuestraStocksFechas.frx":054D
      FormatStyle(5)  =   "FrmMuestraStocksFechas.frx":0625
      FormatStyle(6)  =   "FrmMuestraStocksFechas.frx":06DD
      ImageCount      =   0
      PrinterProperties=   "FrmMuestraStocksFechas.frx":07BD
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   8745
      Top             =   8340
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmMuestraStocksFechas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public CODIGO As String, DESCRIPCION As String
Dim STRSQL As String, sTit_OP As String, rstAux As ADODB.Recordset
Public sOpcion  As String
Public scod_Estcli As String
Public sCod_Calidad As String
Private flg_color As Integer
Private flg_talla As String
Private flg_imagen As String

Private Sub cboAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub chkColor_Click()
Set GridEX1.ADORecordset = Nothing
End Sub

Private Sub chkTalla_Click()
Set GridEX1.ADORecordset = Nothing
End Sub

Private Sub cmdVerDetMov_Click()
On Error GoTo FIN
    If GridEX1.RowCount <= 0 Then Exit Sub
    
    Load FrmMuestraDetalleMovPrendas
    FrmMuestraDetalleMovPrendas.scod_almacen = Left(Trim(cboAlmacen), 2)
    FrmMuestraDetalleMovPrendas.scod_Estcli = RTrim(GridEX1.Value(GridEX1.Columns("cod_estcli").Index))
    FrmMuestraDetalleMovPrendas.scod_present = RTrim(GridEX1.Value(GridEX1.Columns("cod_present").Index))
    FrmMuestraDetalleMovPrendas.scod_talla = RTrim(GridEX1.Value(GridEX1.Columns("cod_talla").Index))
    FrmMuestraDetalleMovPrendas.muestradatos
    FrmMuestraDetalleMovPrendas.Show 1

Exit Sub
FIN:
MsgBox "Inconvenientes PAra mostrar detalle" + err.Description, vbInformation + vbOKOnly, "IMPORTANTE"

End Sub

Private Sub fnbBuscar_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    On Error GoTo SALTO_ERROR
    
    scod_Estcli = TxtCodigo.Text
    
    If cboAlmacen.ListIndex = -1 Then
        MsgBox "Se debe elegir un Almacen", vbOKOnly + vbExclamation, "Ver Saldo de Stocks"
    End If
    flg_color = "0"
    flg_talla = "N"
    
    If chkColor.Value = 1 Then
     flg_color = "1"
    End If
    
    If chkTalla.Value = 1 Then
     flg_talla = "S"
    End If
    
    'CF_MUESTRA_DETALLE_STOCKS_RANGO_FECHAS '69','','1','X','01/02/2015','25/02/2015'

    'STRSQL = "EXEC SM_MUESTRA_CF_STOCKS_ALMACEN_PRENDAS_TERMINADAS '" & sOpcion & "','" & Right(cboAlmacen, 2) & "' ,'" & Left(cboAlmacen, 2) & _
    '         "', '" & Txtcod_Fabrica & "', '" & txtCod_ordpro & "','" & scod_Estcli & "','" & sCod_Calidad & "','" & flg_color & "','" & flg_talla & "'"
             
    STRSQL = "EXEC CF_MUESTRA_DETALLE_STOCKS_RANGO_FECHAS '" & Left(cboAlmacen, 2) & _
             "','" & scod_Estcli & "','" & flg_color & "','" & flg_talla & "','" & DTPInicio & "','" & DTPHasta & "'"
    Set GridEX1.ADORecordset = CargarRecordSetDesconectado(STRSQL, cConnect)
    Call CONFIGURA_GRILLA

    If GridEX1.RowCount > 0 Then
        GridEX1.SetFocus
    Else
        cboAlmacen.SetFocus
    End If
    
    Exit Sub
SALTO_ERROR:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub
Private Sub CONFIGURA_GRILLA()
    On Error GoTo SALTO_ERROR
    Dim C As Integer
    With GridEX1
    
        For C = 1 To .Columns.Count
            .Columns(C).Visible = False
            .Columns(C).HeaderAlignment = jgexAlignCenter
            .Columns(C).TextAlignment = jgexAlignCenter
        Next C

'        With .Columns("NOM_CLIENTE")
'            .Visible = True
'            .Width = 1500
'            .Caption = "CLIENTE"
'            .TextAlignment = jgexAlignLeft
'        End With
'        With .Columns("NOM_TEMCLI")
'            .Visible = True
'            .Width = 1500
'            .Caption = "TEMP"
'            .TextAlignment = jgexAlignLeft
'        End With
'
'        With .Columns("COD_PURORD")
'            .Visible = True
'            .Width = 1000
'            .Caption = "PO"
'            .TextAlignment = jgexAlignLeft
'        End With
                        
        With .Columns("cod_estcli")
            .Visible = True
            .Width = 2500
            .Caption = "ESTILO"
            .TextAlignment = jgexAlignLeft
        End With
                        
        With .Columns("DES_ESTCLI")
            .Visible = True
            .Width = 2500
            .Caption = "ESTILO"
            .TextAlignment = jgexAlignLeft
        End With
        
'        With .Columns("Cod_OrdPro")
'            .Visible = True
'            .Width = 800
'            .Caption = "OP"
'            .TextAlignment = jgexAlignCenter
'        End With
        
        With .Columns("Cod_Present")
            .Visible = False
            .Width = 500
            .Caption = "COD_PRESENT"
            .TextAlignment = jgexAlignLeft
        End With
        
        'Presentacion
        With .Columns("DES_PRESENT")
            .Visible = True
            .Width = 2000
            .Caption = "PRESENTACION"
            .TextAlignment = jgexAlignLeft
        End With
        With .Columns("Cod_Talla")
            .Visible = True
            .Width = 600
            .Caption = "TALLA"
            .TextAlignment = jgexAlignCenter
        End With
        
        With .Columns("STK_INICIAL")
            .Visible = True
            .Width = 1300
            .Caption = "STK_INICIAL"
            .TextAlignment = jgexAlignLeft
        End With
        With .Columns("ENTRADAS")
            .Visible = True
            .Width = 1300
            .Caption = "ENTRADAS"
            .TextAlignment = jgexAlignLeft
        End With
        With .Columns("SALIDAS")
            .Visible = True
            .Width = 1300
            .Caption = "SALIDAS"
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("STK_FINAL")
            .Visible = True
            .Width = 1300
            .Caption = "STK_FINAL"
            .TextAlignment = jgexAlignLeft
        End With
        
        End With
        
    Exit Sub
    
SALTO_ERROR:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub
Private Sub FillAlmacen()
Dim rstAlm As ADODB.Recordset

    STRSQL = "SM_MUESTRA_ALMACEN_STOCKS_PRENDAS '" & vusu & "' "
    Set rstAlm = CargarRecordSetDesconectado(STRSQL, cConnect)
    rstAlm.MoveFirst
    cboAlmacen.Clear
    Do Until rstAlm.EOF
        cboAlmacen.AddItem rstAlm!Cod_almacen & " " & rstAlm!nom_almacen
        rstAlm.MoveNext
    Loop
    rstAlm.Close
    Set rstAlm = Nothing
End Sub

Private Sub Form_Load()
    STRSQL = "Select Top 1 Titulo_Abr_Orden From TG_Control"
    'sTit_OP = DevuelveCampo(STRSQL, cConnect)
    'lblCod_OrdTra = sTit_OP
    FillAlmacen
'    STRSQL = "Select count(Cod_Fabrica) From TG_FABRICA"
'    If DevuelveCampo(STRSQL, cConnect) = 1 Then BuscaFabrica 1
    sOpcion = "1"
    flg_color = "0"
    flg_talla = "N"
    flg_imagen = "N"
    DTPInicio = Format(Date - 7, "DD/mm/YYYY")
    DTPHasta = Format(Date, "DD/mm/YYYY")
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If Not oMDIParent Is Nothing Then oMDIParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim strCalidad As String

    Select Case ActionName
    Case "IMPRIMIR"
        Reporte
    Case "SALIR"
        Unload Me
    End Select
End Sub

Private Sub Reporte()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Usu As String

    flg_imagen = "N"
    If chkImagen.Value = 1 Then
     flg_imagen = "S"
    End If

    Ruta = vRuta & "\RptStocksPrendasRangoFechas.xlt"
    Screen.MousePointer = 11
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "reporte", GridEX1.ADORecordset, scod_Estcli, flg_color, flg_talla, flg_imagen, Format(DTPInicio, "DD/MM/YYYY"), Format(DTPHasta, "DD/MM/YYYY")
    Set oo = Nothing
    Screen.MousePointer = 0
Exit Sub
hand:
    Screen.MousePointer = 0
    ErrorHandler err, "Reporte"
    Set oo = Nothing
End Sub

Private Sub optEstCli_Click()
sOpcion = "3"
'fraNP.Visible = False
FraOpciones.Visible = True
FraOpciones.Width = 4200
TxtDescripcion.Visible = False
TxtCodigo.Width = 2500
Label2.Caption = "ESTILO CLIENTE"
TxtCodigo.Text = ""
TxtDescripcion.Text = ""
End Sub
Private Sub optTodos_Click()
    sOpcion = "1"
    'fraNP.Visible = False
    FraOpciones.Visible = False
    TxtCodigo.Text = ""
End Sub



Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_EstCli", "Des_EstCli", "Tg_EstCliTem where  ", TxtCodigo, TxtDescripcion, 1, Me)
    fnbBuscar.SetFocus
End If
End Sub
Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_EstCli", "Des_EstCli", "Tg_EstCliTem where ", TxtCodigo, TxtDescripcion, 2, Me)
   fnbBuscar.SetFocus
End If
End Sub


Public Sub Busca_Opcion(strCampo1 As String, strCampo2 As String, strTabla As String, txtCod As TextBox, txtDes As TextBox, Opcion As Integer, frmME As Form)

On Error GoTo FIN

Dim rstAux As ADODB.Recordset, STRSQL As String

    STRSQL = "Select DISTINCT " & strCampo1 & " AS Cod," & strCampo2 & " as Descripcion from " & strTabla

    txtCod = Trim(txtCod)
    txtDes = Trim(txtDes)
    
    
    Select Case Opcion
    Case 1: STRSQL = STRSQL & strCampo1 & " like '%" & txtCod & "%'"
    Case 2: STRSQL = STRSQL & strCampo2 & " like '%" & txtDes & "%'"
   
    
    End Select
    txtCod = ""
    txtDes = ""
    
    With frmBusqGeneral
        Set .oParent = frmME
        .SQuery = STRSQL
        .CARGAR_DATOS
        
        frmME.CODIGO = ""
        Set rstAux = .gexList.ADORecordset
        If rstAux.RecordCount > 1 Then
          .Show vbModal
        Else
          frmME.CODIGO = ".."
        End If
        
        If frmME.CODIGO <> "" And rstAux.RecordCount > 0 Then
            txtCod = frmME.CODIGO 'Trim(rstAux!Cod)
            txtDes = frmME.DESCRIPCION  'Trim(rstAux!Descripcion)
            
            If txtCod = ".." Or frmME.DESCRIPCION = "" Then
                txtCod = Trim(rstAux!cod)
                txtDes = Trim(rstAux!DESCRIPCION)
            End If
            
            Select Case Opcion
            Case 1: SendKeys "{TAB}": SendKeys "{TAB}"
            Case 2: SendKeys "{TAB}"
            End Select
            
        Else
            SendKeys "{TAB}"
        End If
        
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Resume
FIN:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento (" & Opcion & ")"
End Sub




