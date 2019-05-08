VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmStocksTenido 
   Caption         =   "Stocks Tela Acabada"
   ClientHeight    =   8280
   ClientLeft      =   0
   ClientTop       =   240
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   14715
   StartUpPosition =   3  'Windows Default
   Begin GridEX20.GridEX gexStock 
      Height          =   6015
      Left            =   90
      TabIndex        =   11
      Top             =   1635
      Width           =   14565
      _ExtentX        =   25691
      _ExtentY        =   10610
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      RowHeaders      =   -1  'True
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmStocksTenido.frx":0000
      FormatStyle(2)  =   "frmStocksTenido.frx":0138
      FormatStyle(3)  =   "frmStocksTenido.frx":01E8
      FormatStyle(4)  =   "frmStocksTenido.frx":029C
      FormatStyle(5)  =   "frmStocksTenido.frx":0374
      FormatStyle(6)  =   "frmStocksTenido.frx":042C
      FormatStyle(7)  =   "frmStocksTenido.frx":050C
      ImageCount      =   0
      PrinterProperties=   "frmStocksTenido.frx":052C
   End
   Begin VB.Frame FraMain 
      BackColor       =   &H00C0FFFF&
      Height          =   1635
      Left            =   90
      TabIndex        =   12
      Top             =   30
      Width           =   14565
      Begin VB.Frame FraRangoPartidas 
         BackColor       =   &H00C0FFFF&
         Height          =   750
         Left            =   6480
         TabIndex        =   27
         Top             =   840
         Width           =   6255
         Begin VB.TextBox txtPartidaInicio 
            Height          =   300
            Left            =   720
            MaxLength       =   5
            TabIndex        =   29
            Top             =   270
            Width           =   1245
         End
         Begin VB.TextBox txtPartidaFin 
            Height          =   300
            Left            =   2820
            MaxLength       =   5
            TabIndex        =   28
            Top             =   270
            Width           =   1260
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Partida Fin"
            Height          =   450
            Left            =   2040
            TabIndex        =   31
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Partida Inicio"
            Height          =   450
            Left            =   75
            TabIndex        =   30
            Top             =   195
            Width           =   600
         End
      End
      Begin VB.OptionButton optRangoPartidas 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Rango de Partidas"
         Height          =   210
         Left            =   2880
         TabIndex        =   26
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton optTodos 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Todos"
         Height          =   210
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton optCliente 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Cliente"
         Height          =   210
         Left            =   1050
         TabIndex        =   22
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton optPartida 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Partida"
         Height          =   210
         Left            =   1920
         TabIndex        =   21
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox cboAlmacen 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   405
         Width           =   5040
      End
      Begin FunctionsButtons.FunctButt fnbBuscar 
         Height          =   495
         Left            =   13200
         TabIndex        =   10
         Top             =   600
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
      Begin VB.Frame FraCliente 
         BackColor       =   &H00C0FFFF&
         Height          =   1410
         Left            =   6480
         TabIndex        =   13
         Top             =   165
         Width           =   6225
         Begin VB.OptionButton optGuia 
            BackColor       =   &H00C0FFFF&
            Caption         =   "&Guia"
            Height          =   210
            Left            =   165
            TabIndex        =   6
            Top             =   1095
            Width           =   1545
         End
         Begin VB.OptionButton optOC 
            BackColor       =   &H00C0FFFF&
            Caption         =   "&Orden de Compra"
            Height          =   210
            Left            =   165
            TabIndex        =   5
            Top             =   855
            Width           =   1545
         End
         Begin VB.OptionButton optAllCli 
            BackColor       =   &H00C0FFFF&
            Caption         =   "&Todos"
            Height          =   210
            Left            =   165
            TabIndex        =   4
            Top             =   615
            Width           =   1545
         End
         Begin VB.TextBox txtAbr_Cliente 
            Height          =   300
            Left            =   1005
            TabIndex        =   2
            Top             =   210
            Width           =   900
         End
         Begin VB.TextBox txtNom_Cliente 
            Height          =   300
            Left            =   1935
            TabIndex        =   3
            Top             =   210
            Width           =   3960
         End
         Begin VB.Frame FraOC 
            BackColor       =   &H00C0FFFF&
            Height          =   750
            Left            =   1830
            TabIndex        =   15
            Top             =   570
            Width           =   4230
            Begin VB.TextBox txtCod_OrdComp 
               Height          =   300
               Left            =   1380
               TabIndex        =   9
               Top             =   270
               Width           =   1140
            End
            Begin VB.TextBox txtSer_OrdComp 
               Height          =   300
               Left            =   840
               TabIndex        =   8
               Top             =   270
               Width           =   525
            End
            Begin VB.Label Label1 
               BackColor       =   &H00C0FFFF&
               Caption         =   "O/C"
               Height          =   210
               Left            =   315
               TabIndex        =   16
               Top             =   315
               Width           =   360
            End
         End
         Begin VB.Frame FraGuia 
            Height          =   750
            Left            =   1830
            TabIndex        =   17
            Top             =   570
            Width           =   4230
            Begin VB.TextBox txtNumero_Guia 
               Height          =   300
               Left            =   825
               TabIndex        =   7
               Top             =   270
               Width           =   1335
            End
            Begin VB.Label Label3 
               Caption         =   "Guia"
               Height          =   210
               Left            =   315
               TabIndex        =   18
               Top             =   315
               Width           =   360
            End
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Cliente"
            Height          =   225
            Left            =   165
            TabIndex        =   14
            Top             =   255
            Width           =   675
         End
      End
      Begin VB.Frame fraPartida 
         BackColor       =   &H00C0FFFF&
         Height          =   1380
         Left            =   6480
         TabIndex        =   19
         Top             =   120
         Width           =   6225
         Begin VB.TextBox txtCod_OrdTra_Tinto 
            Height          =   300
            Left            =   2580
            TabIndex        =   1
            Top             =   615
            Width           =   900
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Partida"
            Height          =   225
            Left            =   1740
            TabIndex        =   20
            Top             =   660
            Width           =   675
         End
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Almacen"
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   165
         Width           =   675
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   0
      TabIndex        =   25
      Top             =   7680
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   900
      Custom          =   $"frmStocksTenido.frx":0704
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1200
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   1005
      Top             =   6240
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmStocksTenido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CODIGO As String, Descripcion As String, TipoAdd As String
Dim StrSQL As String, rstAux As ADODB.Recordset, sTit As String, sErr As String
Dim scod_almacen As String, sOpcion As String, SCod_Cliente_Tex As String, _
    sser_ordcomp As String, scod_ordcomp As String, sNumero_Guia As String, _
    sNom_Cliente As String, sDes_Almacen As String, sCod_OrdTra_Tinto As String
    
Dim sCod_Cliente_Comercial As String, sNom_Cliente_Comercial As String

Private Sub cboAlmacen_Click()
 scod_almacen = Left(cboAlmacen, 2)
End Sub

Private Sub cboAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
    scod_almacen = Left(cboAlmacen, 2)
    
End Sub

Private Sub fnbBuscar_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo ErrBusq
    sTit = "Mostrar Stocks"
    scod_almacen = "": sOpcion = "": SCod_Cliente_Tex = ""
    sser_ordcomp = "": scod_ordcomp = "": sNumero_Guia = ""
    sNom_Cliente = "": sDes_Almacen = "": sCod_OrdTra_Tinto = ""
    
    scod_almacen = Left(cboAlmacen, 2)
    sDes_Almacen = Mid(cboAlmacen, 3)
    Select Case True
  
    Case optTodos
        sOpcion = "1"

    Case optcliente
        If TxtAbr_Cliente.Tag = "" Then
            MsgBox "Se debe especificar un Cliente", vbExclamation + vbOKOnly, sTit
            TxtAbr_Cliente.SetFocus
            Exit Sub
        End If
        SCod_Cliente_Tex = TxtAbr_Cliente.Tag
        sNom_Cliente = txtNom_Cliente
        Select Case True
        Case optAllCli
            sOpcion = "2"
        Case OptOC
            txtSer_OrdComp = Trim(txtSer_OrdComp)
            txtCod_OrdComp = Trim(txtCod_OrdComp)
            
            If Len(txtSer_OrdComp) <> 3 Or Len(txtCod_OrdComp) <> 6 Then
                MsgBox "Orden de Compra Invàlida", vbExclamation + vbOKOnly, sTit
                txtSer_OrdComp.SetFocus
                Exit Sub
            End If
            
            sOpcion = "3"
            sser_ordcomp = txtSer_OrdComp
            scod_ordcomp = txtCod_OrdComp
        Case optGuia
            txtNumero_Guia = Trim(txtNumero_Guia)
            
            If Len(txtNumero_Guia) <> 8 Then
                MsgBox "Orden de Compra Invàlida", vbExclamation + vbOKOnly, sTit
                txtNumero_Guia.SetFocus
                Exit Sub
            End If
            
            sOpcion = "4"
            sNumero_Guia = txtNumero_Guia
        End Select
    
    Case optPartida
        sOpcion = "5"
        sCod_OrdTra_Tinto = txtCod_OrdTra_Tinto
        
    Case optRangoPartidas
        sOpcion = "8"

    End Select
    
'    Case optRangoPartidas
'        sopcion = "8"
'    End Select
'
    Screen.MousePointer = 11
        
        StrSQL = "EXEC TI_SM_MUESTRA_STOCKS_ALMACEN_TELA_TENIDA_CLIENTES '" & _
                 scod_almacen & "', '" & sOpcion & "', '" & SCod_Cliente_Tex & _
                 "', '" & sser_ordcomp & "', '" & scod_ordcomp & "', '" & _
                 sNumero_Guia & "', '" & sCod_OrdTra_Tinto & "', '" & vusu & "','" & sCod_Cliente_Comercial & "','" & txtPartidaInicio.Text & "','" & txtPartidaFin.Text & "'"
        Set gexStock.ADORecordset = CargarRecordSetDesconectado(StrSQL, cConnect)
        
        gexStock.Columns("Cliente").Width = 1005
        gexStock.Columns("OC").Width = 930
        gexStock.Columns("Partida").Width = 630
        gexStock.Columns("Partida_Cliente").Width = 1200
        gexStock.Columns("Tela").Width = 1965
        gexStock.Columns("Comb").Width = 0
        gexStock.Columns("Color").Width = 1800
        gexStock.Columns("Talla").Width = 630
        gexStock.Columns("Calidad").Width = 290
        gexStock.Columns("Kilos_Ingresados").Width = 885
        gexStock.Columns("Kilos_Despachados").Width = 795
        gexStock.Columns("Uni_Ingresados").Width = 1095
        gexStock.Columns("Uni_Despachados").Width = 990
        gexStock.Columns("Nro_Rollos_Ingresados").Width = 1000
        gexStock.Columns("Nro_Rollos_Despachados").Width = 1000
        gexStock.Columns("Fec_1ER_Ingreso").Width = 1065
        gexStock.Columns("Fec_Ult_Ingreso").Width = 1065
        gexStock.Columns("Fec_1ER_Salida").Width = 1140
        gexStock.Columns("Fec_Ult_Salida").Width = 960
        gexStock.Columns("Kilos_Despachados_Reales").Width = 1080
        gexStock.Columns("Kgs_Comprometidos").Width = 1140
        gexStock.Columns("NP").Width = 660
        
        gexStock.Columns("Stock_Kilos").Width = 1000
        gexStock.Columns("stock_rollos").Width = 800
        gexStock.Columns("Uni_Stock").Width = 1000
        
        
        
        gexStock.Columns("cod_almacen").Visible = False
        gexStock.Columns("cod_cliente_tex").Visible = False
        gexStock.Columns("cod_ordtra").Visible = False
        gexStock.Columns("cod_tela").Visible = False
        gexStock.Columns("cod_comb").Visible = False
        gexStock.Columns("cod_color").Visible = False
        gexStock.Columns("cod_talla").Visible = False
        gexStock.Columns("cod_calidad").Visible = False
        
        'gexStock.Columns("Cliente").Caption = 1005
        'gexStock.Columns("OC").Caption = 930
        'gexStock.Columns("Partida").Caption = 630
        gexStock.Columns("Partida_Cliente").Caption = "Part.Cliente"
        'gexStock.Columns("Tela").Caption = ""
        gexStock.Columns("Comb").Caption = "Combo"
        gexStock.Columns("Color").Caption = "Color"
        gexStock.Columns("Talla").Caption = "Talla"
        gexStock.Columns("Calidad").Caption = "Cal"
        gexStock.Columns("Kilos_Ingresados").Caption = "Kgs.Ingr."
        gexStock.Columns("Kilos_Despachados").Caption = "Kgs.Desp"
        gexStock.Columns("Uni_Ingresados").Caption = "Und.Ingr."
        gexStock.Columns("Uni_Despachados").Caption = "Und.Desp"
        gexStock.Columns("Nro_Rollos_Ingresados").Caption = "Rollos Ingr."
        gexStock.Columns("Nro_Rollos_Despachados").Caption = "Rollos Desp"
        gexStock.Columns("Fec_1ER_Ingreso").Caption = "1er Ingreso"
        gexStock.Columns("Fec_Ult_Ingreso").Caption = "Ult.Ingreso"
        gexStock.Columns("Fec_1ER_Salida").Caption = "1era Salida"
        gexStock.Columns("Fec_Ult_Salida").Caption = "Ult.Salida"
        gexStock.Columns("Kilos_Despachados_Reales").Caption = "Kgs.Desp"
        gexStock.Columns("Kgs_Comprometidos").Caption = "Kgs.Comp"
        
        gexStock.Columns("Stock_Kilos").Caption = "Stock Kilos"
        gexStock.Columns("stock_rollos").Caption = "Stock Rollos"
        gexStock.Columns("Uni_Stock").Caption = "Stock Uni "

    Screen.MousePointer = 0
Exit Sub
ErrBusq:
    sErr = err.Description
    Screen.MousePointer = 0
    MsgBox sErr, vbCritical + vbOKOnly, sTit
End Sub

Private Sub fnbBuscar_GotFocus()
    fnbBuscar_ActionClick 0, 0, ""
End Sub

Private Sub Form_Load()
    FillAlmacen
    optTodos = True
    optAllCli = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If Not oParent Is Nothing Then oParent.DropWindowList Me.Tag
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
    Case "IMPRIMIR"
            Reporte_Excel

    Case "VERDET"
        If gexStock.RowCount = 0 Then Exit Sub
        If sOpcion = "0" Then Exit Sub
        Load frmStockTenDet
        frmStockTenDet.scod_almacen = gexStock.Value(gexStock.Columns("Cod_Almacen").Index)
        frmStockTenDet.sCod_OrdTra = gexStock.Value(gexStock.Columns("Cod_OrdTra").Index)
        frmStockTenDet.sCod_Tela = gexStock.Value(gexStock.Columns("Cod_Tela").Index)
        frmStockTenDet.sCod_Comb = gexStock.Value(gexStock.Columns("Cod_Comb").Index)
        frmStockTenDet.sCod_Color = gexStock.Value(gexStock.Columns("Cod_Color").Index)
        frmStockTenDet.scod_talla = gexStock.Value(gexStock.Columns("Cod_Talla").Index)
        frmStockTenDet.sCod_Calidad = gexStock.Value(gexStock.Columns("Cod_Calidad").Index)
        frmStockTenDet.MostrarDetalleTen
        frmStockTenDet.Show vbModal
        Set frmStockTenDet = Nothing
    Case "GIRADO"
        If gexStock.RowCount = 0 Then Exit Sub
        If sOpcion = "0" Then Exit Sub
        
'            Load FrmVerStockGirado_Oc
'            FrmVerStockGirado_Oc.vCod_Almacen = gexStock.Value(gexStock.Columns("Cod_Almacen").Index)
'            FrmVerStockGirado_Oc.vCod_OrdTra = gexStock.Value(gexStock.Columns("Cod_OrdTra").Index)
'            FrmVerStockGirado_Oc.vCod_Tela = gexStock.Value(gexStock.Columns("Cod_Tela").Index)
'            FrmVerStockGirado_Oc.vCod_Comb = gexStock.Value(gexStock.Columns("Cod_Comb").Index)
'            FrmVerStockGirado_Oc.vCod_Color = gexStock.Value(gexStock.Columns("Cod_Color").Index)
'            FrmVerStockGirado_Oc.vCod_Calidad = gexStock.Value(gexStock.Columns("Cod_Calidad").Index)
'            FrmVerStockGirado_Oc.CARGA_GRID
'            FrmVerStockGirado_Oc.Show vbModal
'            Set FrmVerStockGirado_Oc = Nothing
     
    Case "CONFIRMAR"
        If gexStock.RowCount = 0 Then Exit Sub
        'Confirmar
    Case "IMPRIMEETIROLLO"
        If gexStock.RowCount = 0 Then Exit Sub
         Call IMPRIMEETIQUETAROLLO
        
    Case "ADDROLLO"
        If gexStock.RowCount = 0 Then Exit Sub
        If sOpcion = "0" Then Exit Sub
        
            If gexStock.Value(gexStock.Columns("Cod_Almacen").Index) = "31" And scod_almacen = "31" Then
'                Load FrmAdicionaRollo
'                FrmAdicionaRollo.vCod_Almacen = gexStock.Value(gexStock.Columns("Cod_Almacen").Index)
'                FrmAdicionaRollo.vCod_OrdTra = gexStock.Value(gexStock.Columns("Cod_OrdTra").Index)
'                FrmAdicionaRollo.txtPartida.Text = gexStock.Value(gexStock.Columns("Cod_OrdTra").Index)
'
'                FrmAdicionaRollo.vCod_Tela = gexStock.Value(gexStock.Columns("Cod_Tela").Index)
'                FrmAdicionaRollo.txtCod_tela.Text = gexStock.Value(gexStock.Columns("Cod_Tela").Index)
'                FrmAdicionaRollo.txtdes_tela.Text = gexStock.Value(gexStock.Columns("Tela").Index)
'
'                FrmAdicionaRollo.vCod_Comb = gexStock.Value(gexStock.Columns("Cod_Comb").Index)
'
'                FrmAdicionaRollo.vCod_Color = gexStock.Value(gexStock.Columns("Cod_Color").Index)
'                FrmAdicionaRollo.txtcod_color.Text = gexStock.Value(gexStock.Columns("Cod_Color").Index)
'                FrmAdicionaRollo.txtdes_color.Text = gexStock.Value(gexStock.Columns("Color").Index)
'
'
'                FrmAdicionaRollo.txtRollos_ingres.Text = gexStock.Value(gexStock.Columns("Nro_Rollos_Ingresados").Index)
'                FrmAdicionaRollo.txtRollos_desp.Text = gexStock.Value(gexStock.Columns("Nro_Rollos_Despachados").Index)
'                FrmAdicionaRollo.vCod_Calidad = gexStock.Value(gexStock.Columns("Cod_Calidad").Index)
'                FrmAdicionaRollo.txtcalidad.Text = gexStock.Value(gexStock.Columns("Cod_Calidad").Index)
'
'                FrmAdicionaRollo.Txt_Cantidad.Text = 1
'
'                FrmAdicionaRollo.Show vbModal
'                Set FrmAdicionaRollo = Nothing
           
        End If
        
    Case "SALIR"
        Unload Me
    End Select
End Sub
Sub Reporte_Excel()
On Error GoTo ErrorImpresion
    Dim oo As Object
    Set oo = CreateObject("excel.application")
  

        oo.workbooks.Open vRuta & "\TI_STOCK_TELATENIDA.xlt"

    
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.run "REPORTE", gexStock.ADORecordset, cboAlmacen, cConnect
    Screen.MousePointer = vbNormal
    oo.Visible = True
    Set oo = Nothing
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte de Consulta de Mov. Tela Cruda " & err.Description, vbCritical, "Impresion"
End Sub
Private Sub OptVisibles()
    
    Fracliente.Visible = optcliente
    fraPartida.Visible = optPartida
    FraRangoPartidas.Visible = optRangoPartidas

    If Not Fracliente.Visible Then
        TxtAbr_Cliente.Tag = ""
        TxtAbr_Cliente = ""
        txtNom_Cliente = ""
    End If

    
    txtSer_OrdComp = ""
    txtCod_OrdComp = ""
    
    txtPartidaFin.Text = ""
    txtPartidaInicio.Text = ""
    
    txtNumero_Guia = ""
    
    txtCod_OrdTra_Tinto = ""
    
    FraOC.Visible = OptOC
    FraGuia.Visible = optGuia
    
End Sub

Private Sub optAllCli_Click()
    OptVisibles
End Sub

Private Sub optAllCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub optcliente_Click()
    OptVisibles
End Sub

Private Sub optCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub OptClienteComercial_Click()
    OptVisibles
End Sub

Private Sub optGuia_Click()
    OptVisibles
End Sub

Private Sub optGuia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub OptOC_Click()
    OptVisibles
End Sub

Private Sub optOC_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub optPartida_Click()
    OptVisibles
End Sub

Private Sub optPartida_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub optRangoPartidas_Click()
    OptVisibles
End Sub

Private Sub optTodos_Click()
    OptVisibles
End Sub

Public Sub BuscaCliente(Opcion As Integer)
'On Error GoTo Fin
'Dim iCol As Long
'
''    Flg_ClientePropio = False
''    txtCod_FamGrupo.TabIndex = 19
''    txtDes_FamGrupo.TabIndex = 20
''    txtCod_Tela_tejeduria.TabIndex = 21
''    txtCod_Tela_tejeduria.Enabled = False
'
'    txtAbr_Cliente = Trim(txtAbr_Cliente)
'    txtNom_Cliente = Trim(txtNom_Cliente)
'    strSQL = "SELECT Abr_Cliente, Nom_Cliente, Cod_Cliente_Tex, Res_Cliente, Num_Ruc, " & _
'    "Cod_Moneda, Peso_por_Rollo, Flg_ClientePropio FROM TX_CLIENTE WHERE "
'
'    Select Case Opcion
'    Case 1: strSQL = strSQL & "Abr_Cliente like '" & txtAbr_Cliente & IIf(txtAbr_Cliente = "", "%", "") & "'"
'    Case 2: strSQL = strSQL & "Nom_Cliente like '%" & txtNom_Cliente & "%'"
'    End Select
'    strSQL = strSQL & " ORDER BY Abr_Cliente"
'
'    txtAbr_Cliente = ""
'    txtAbr_Cliente.Tag = ""
'    txtNom_Cliente = ""
'
'    With frmBusqGeneral
'        Set .oParent = Me
'        .sQuery = strSQL
'        .Cargar_Datos
'        Codigo = ".."
'        Set rstAux = .DGridLista.ADORecordset
'
'        .DGridLista.Columns("Abr_Cliente").Caption = "Abrev."
'        .DGridLista.Columns("Abr_Cliente").Width = 700
'        .DGridLista.Columns("Nom_Cliente").Caption = "Nombre Cliente"
'        .DGridLista.Columns("Nom_Cliente").Width = 5000
'        For iCol = 3 To .DGridLista.Columns.Count
'            .DGridLista.Columns(iCol).Visible = False
'        Next iCol
'
'        If rstAux.RecordCount > 1 Then .Show vbModal
'
'        If Codigo <> "" And rstAux.RecordCount > 0 Then
'            txtAbr_Cliente.Tag = Trim(rstAux!Cod_Cliente_Tex)
'            txtAbr_Cliente = Trim(rstAux!Abr_Cliente)
'            txtNom_Cliente = Trim(rstAux!Nom_Cliente)
'        End If
'    End With
'    Unload frmBusqGeneral
'    Set frmBusqGeneral = Nothing
'    rstAux.Close
'    Set rstAux = Nothing
'Exit Sub
'Fin:
'On Error Resume Next
'    Unload frmBusqGeneral
'    Set frmBusqGeneral = Nothing
'    rstAux.Close
'    Set rstAux = Nothing
'    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
'    "Búsqueda de Cliente (" & Opcion & ")"
End Sub

Private Sub optTodos_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtAbr_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        'BuscaCliente 1
        SendKeys "{TAB}"
    End If
End Sub


Private Sub txtCod_OrdComp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        'txtCod_OrdComp = Format(txtCod_OrdComp, "000000")
        'BuscaOrdenCompra
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCod_OrdTra_Tinto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        txtCod_OrdTra_Tinto = Format(txtCod_OrdTra_Tinto, "00000")
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtNom_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        'BuscaCliente 2
        SendKeys "{TAB}"
    End If
End Sub



Private Sub txtNumero_Guia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        txtNumero_Guia = Format(txtNumero_Guia, "00000000")
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtPartidaInicio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Or KeyAscii = 0 Then
     txtPartidaFin.SetFocus
  End If
End Sub
Private Sub txtPartidaFin_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Or KeyAscii = 0 Then
     fnbBuscar.SetFocus
  End If
End Sub

Private Sub txtPartidaInicio_LostFocus()
    txtPartidaInicio.Text = Format(txtPartidaInicio.Text, "00000")
End Sub
Private Sub txtPartidaFin_LostFocus()
    txtPartidaFin.Text = Format(txtPartidaFin.Text, "00000")
End Sub

Private Sub txtSer_OrdComp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        'txtSer_OrdComp = Format(txtSer_OrdComp, "000")
        'BuscaOrdenCompra
        SendKeys "{TAB}"
    End If
End Sub

Private Sub FillAlmacen()
On Error GoTo fin
Dim sTit As String
    
    sTit = "Cargar Almacenes"
    
    StrSQL = "SELECT Cod_Almacen, Nom_Almacen FROM TX_ALMACEN " & _
             "WHERE  Tip_Item = 'T' " & _
             "AND    Tip_Presentacion = 'T'"
    
    Set rstAux = CargarRecordSetDesconectado(StrSQL, cConnect)
    cboAlmacen.Clear
    With rstAux
        If .RecordCount > 0 Then .MoveFirst
        Do Until .EOF
            cboAlmacen.AddItem !Cod_almacen & " " & !nom_almacen
            .MoveNext
        Loop
        .Close
    End With
    If cboAlmacen.ListCount > 0 Then cboAlmacen.ListIndex = 0
    Set rstAux = Nothing
Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, sTit
End Sub

'Public Sub BuscaOrdenCompra()
'On Error GoTo Fin
'Dim sErr As String
'
'    txtCod_OrdComp = Trim(txtCod_OrdComp)
'    txtSer_OrdComp = Trim(txtSer_OrdComp)
'
'    strSQL = "EXEC TI_SM_MUESTRA_ORDENES_COMPRA_TINTO_ABIERTAS '" & txtAbr_Cliente.Tag & "'"
'
'    txtCod_OrdComp = ""
'    txtSer_OrdComp = ""
'    With frmBusqGeneral
'        Set .oParent = Me
'        .sQuery = strSQL
'        .Cargar_Datos
'
'        .DGridLista.Columns("Ser_OrdComp").Caption = "Serie"
'        .DGridLista.Columns("Ser_OrdComp").Width = 500
'        .DGridLista.Columns("Cod_OrdComp").Caption = "Nro.Orden"
'        .DGridLista.Columns("Cod_OrdComp").Width = 5000
'        Set rstAux = .DGridLista.ADORecordset
'        If rstAux.RecordCount > 1 Then .Show vbModal
'        Codigo = ".."
'
'        If Codigo <> "" And rstAux.RecordCount > 0 Then
'            txtSer_OrdComp = rstAux!Ser_OrdComp
'            txtCod_OrdComp = rstAux!Cod_OrdComp
'        End If
'    End With
'    Unload frmBusqGeneral
'    rstAux.Close
'    Set rstAux = Nothing
'Exit Sub
'Fin:
'    sErr = Err.Description
'    On Error Resume Next
'    Unload frmBusqGeneral
'    Set frmBusqGeneral = Nothing
'    rstAux.Close
'    Set rstAux = Nothing
'
'    MsgBox sErr, vbCritical + vbOKOnly, "Busqueda Orden de Compra"
'End Sub


Private Sub Confirmar()
On Error GoTo ErrDet
    
    If gexStock.RowCount = 0 Then Exit Sub
    
    If MsgBox("Desea Confirmar Partida?", vbQuestion + vbYesNo, "Confirmacion Partida") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    
    StrSQL = "EXEC Ti_Captura_Tela_Cruda '" & gexStock.Value(gexStock.Columns("Cod_Almacen").Index) & "', '" & _
             gexStock.Value(gexStock.Columns("Cod_OrdTra").Index) & "', '" & gexStock.Value(gexStock.Columns("Cod_Tela").Index) & "', '" & gexStock.Value(gexStock.Columns("Cod_Comb").Index) & "','" & gexStock.Value(gexStock.Columns("Cod_Color").Index) & "','" & gexStock.Value(gexStock.Columns("Cod_Talla").Index) & _
             "','" & gexStock.Value(gexStock.Columns("Cod_Calidad").Index) & "','" & vusu & "','" & ComputerName & "'"
    ExecuteSQL cConnect, StrSQL
    
    Screen.MousePointer = 0
    
 
Exit Sub
ErrDet:
    sErr = err.Description
    Screen.MousePointer = 0
    MsgBox sErr, vbCritical + vbOKOnly, "Eliminar Item"
End Sub
Private Sub IMPRIMEETIQUETAROLLO()
On Error GoTo fin
Dim rsrollos As New ADODB.Recordset

If optRangoPartidas = False Then
    MsgBox "¡¡¡advertencia!!!" & Chr(13) & " La impresion de Rollos Solo se puede Realizar por la opcion Rango de partidas", vbInformation + vbOKOnly, "IMPORTANTE"
    Exit Sub
End If

StrSQL = "TI_MUESTRA_RANGO_ROLLO '" & txtPartidaInicio.Text & "','" & txtPartidaFin.Text & "'"
Set rsrollos = CargarRecordSetDesconectado(StrSQL, cConnect)
If rsrollos.RecordCount = 0 Then Exit Sub

With rsrollos
.MoveFirst
Do While Not .EOF
    Call Imprime_ZEBRA_Etiqueta(!cod_ordtra, !color, !Tela, !peso, !codigoRollo, !INV)
    .MoveNext
Loop

End With

Exit Sub
fin:
MsgBox err.Description, vbCritical + vbOKOnly, "Advertencia"
End Sub



Private Function Imprime_ZEBRA_Etiqueta(ByVal partida As String, ByVal color As String, ByVal Tela As String, ByVal KIlos As Double, codigoRollo, INV As String) As Boolean
On Error GoTo errx
Dim sSQL  As String, SBARRA As String, sEmpresa As String
Dim mRs As ADODB.Recordset
Dim oPrint As clsPrintFile

sEmpresa = "LA CASA DEL REACTIVO"

Printer.Print " "
Printer.Print "^XA"
Printer.Print "^PRC"
Printer.Print "^LH0,0^FS"
Printer.Print "^LL1261"
Printer.Print "^MD0"
Printer.Print "^MNY"

SBARRA = codigoRollo
    
Printer.Print "^FO590,30^A0N,50,25^CI13^FR^FD" & "Partida :"; RTrim(partida) & "^FS"

Printer.Print "^FO10,20^A0N,75,30^CI13^FR^FD" & RTrim(sEmpresa) & "^FS"
'Printer.Print "^FO630,30^A0N,50,25^CI13^FR^FD" & "Kilos :"; Format(Str(KIlos), "##.00") & "^FS"

Printer.Print "^BY3,3.0^FO180,100^BCN,100,N,N,N^FR^FD" & Trim(SBARRA) & "^FS"
Printer.Print "^FO20,210^A0N,35,25^CI13^FR^FD"; INV & "^FS"
Printer.Print "^FO330,210^A0N,35,25^CI13^FR^FD"; RTrim(codigoRollo) & "^FS"
Printer.Print "^FO12,255^A0N,35,25^CI13^FR^FD" & "Tela :"; Left(RTrim(Tela), 35) & "^FS"

Printer.Print "^FO455,255^A0N,35,25^CI13^FR^FD" & "Color :"; RTrim(color) & "^FS"

Printer.Print "^PQ1,0, 0, n"
Printer.Print "^XZ"
Printer.Print "^FX End of job"
Printer.Print "^XA"
Printer.Print "^IDR:ID*.*"
Printer.Print "^XZ"
Printer.EndDoc

Exit Function
errx:
    Close #1
    errores err.numer
End Function

