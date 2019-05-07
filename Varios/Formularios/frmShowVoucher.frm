VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmShowVoucher 
   Caption         =   "Voucher Contable"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDatos 
      Caption         =   "Reingrese Cuenta Contable"
      Height          =   3450
      Left            =   510
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   8310
      Begin VB.TextBox txtDes_TipDoc 
         Height          =   285
         Left            =   2010
         TabIndex        =   25
         Top             =   2145
         Width           =   1605
      End
      Begin VB.TextBox txtDoc_Sunat 
         Height          =   285
         Left            =   1485
         TabIndex        =   22
         Top             =   2145
         Width           =   465
      End
      Begin VB.TextBox txtNum_Docum 
         Height          =   285
         Left            =   5625
         TabIndex        =   21
         Top             =   2145
         Width           =   2385
      End
      Begin VB.TextBox txtSer_Docum 
         Height          =   285
         Left            =   4980
         TabIndex        =   20
         Top             =   2145
         Width           =   585
      End
      Begin VB.TextBox txtGlosa 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1485
         TabIndex        =   7
         Top             =   1770
         Width           =   6525
      End
      Begin VB.TextBox txtDebe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2895
         TabIndex        =   3
         Text            =   "0"
         Top             =   1350
         Width           =   1230
      End
      Begin VB.TextBox txtHaber 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4185
         TabIndex        =   4
         Text            =   "0"
         Top             =   1350
         Width           =   1230
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2895
         TabIndex        =   17
         Text            =   "DEBE"
         Top             =   1005
         Width           =   1230
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4170
         TabIndex        =   16
         Text            =   "HABER"
         Top             =   1005
         Width           =   1230
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6780
         TabIndex        =   15
         Text            =   "HABER"
         Top             =   1005
         Width           =   1230
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5505
         TabIndex        =   14
         Text            =   "DEBE"
         Top             =   1005
         Width           =   1230
      End
      Begin VB.TextBox txtHaberDol 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6780
         TabIndex        =   6
         Text            =   "0"
         Top             =   1350
         Width           =   1230
      End
      Begin VB.TextBox txtDebeDol 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5505
         TabIndex        =   5
         Text            =   "0"
         Top             =   1350
         Width           =   1230
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2895
         TabIndex        =   13
         Text            =   "SOLES"
         Top             =   690
         Width           =   2505
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5505
         TabIndex        =   12
         Text            =   "DOLARES"
         Top             =   690
         Width           =   2505
      End
      Begin VB.TextBox txtTipodeCambio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1485
         TabIndex        =   2
         Text            =   "0"
         Top             =   705
         Width           =   1230
      End
      Begin VB.TextBox txtCuenta 
         Height          =   285
         Left            =   1485
         TabIndex        =   0
         Top             =   315
         Width           =   1290
      End
      Begin VB.TextBox txtDes_Cuenta 
         Height          =   285
         Left            =   2895
         TabIndex        =   1
         Top             =   315
         Width           =   5115
      End
      Begin FunctionsButtons.FunctButt FunctButt3 
         Height          =   510
         Left            =   2880
         TabIndex        =   8
         Top             =   2715
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   ""
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label6 
         Caption         =   "Serie/Número"
         Height          =   285
         Left            =   3705
         TabIndex        =   24
         Top             =   2190
         Width           =   1170
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo Doc:"
         Height          =   285
         Left            =   270
         TabIndex        =   23
         Top             =   2190
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Glosa :"
         Height          =   195
         Left            =   255
         TabIndex        =   19
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Cambio :"
         Height          =   195
         Left            =   210
         TabIndex        =   18
         Top             =   780
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Contable"
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1290
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4410
      Left            =   45
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   60
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   7779
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      AllowEdit       =   0   'False
      RowHeaders      =   -1  'True
      DataMode        =   1
      ColumnHeaderHeight=   285
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmShowVoucher.frx":0000
      FormatStyle(2)  =   "frmShowVoucher.frx":0138
      FormatStyle(3)  =   "frmShowVoucher.frx":01E8
      FormatStyle(4)  =   "frmShowVoucher.frx":029C
      FormatStyle(5)  =   "frmShowVoucher.frx":0374
      FormatStyle(6)  =   "frmShowVoucher.frx":042C
      FormatStyle(7)  =   "frmShowVoucher.frx":050C
      ImageCount      =   0
      PrinterProperties=   "frmShowVoucher.frx":052C
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   3360
      Left            =   9120
      TabIndex        =   26
      Top             =   120
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   5927
      Custom          =   $"frmShowVoucher.frx":0704
      Orientacion     =   1
      Style           =   1
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1100
      ControlHeigth   =   490
      ControlSeparator=   80
   End
End
Attribute VB_Name = "frmShowVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCod_TipoDiario As String
Public sano As String
Public smes As String
Public lNum_Registro As Long
Public Num_Corre As String
Public dImporte As Double
Public codigo As String
Public Descripcion As String
Public TipoAdd As String
Public sFlg_Status As String
Public sAccion As String
Public sItem As String, sccta As String
Dim sEsCtaCte As String, sdebehaber As String, sEstado As String
Public sFec_Transaccion As String, smoneda As String
Public sSecuencia As Integer


Private Sub Form_Load()
  FunctButt1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name) & "/SALIR"
  
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "IMPRIMIR"
            Imprimir
        Case "GENERAR"
            Generar
        Case "REVERTIR"
            Revertir
        Case "ADICIONAR"
            sAccion = "I"
            sItem = ""
            txtSer_Docum.Enabled = True
            txtNum_Docum.Enabled = True
            txtDoc_Sunat.Enabled = True
            
            txtCuenta.Text = ""
            txtDes_Cuenta.Text = ""
            txtTipodeCambio.Text = 0
            txtDebe.Text = 0
            txtDebeDol.Text = 0
            txtHaber.Text = 0
            txtHaberDol.Text = 0
            txtGlosa.Text = ""
            txtDoc_Sunat.Text = ""
            txtSer_Docum.Text = ""
            txtNum_Docum.Text = ""
            txtCuenta.Enabled = True
            Me.fraDatos.Visible = True
        Case "MODIFCUENTA"
            If GridEX1.RowCount = 0 Then Exit Sub
            sAccion = "U"
            sItem = GridEX1.Value(GridEX1.Columns("item").Index)
            txtCuenta.Text = GridEX1.Value(GridEX1.Columns("CUENTA").Index)

            sEstado = DevuelveCampo("select DBO.CN_Status_CierreTipoDiario('" & sano & "','" & smes & "','" & sCod_TipoDiario & "','')", cConnect)
            
            If sEstado = "S" Then
                MsgBox "DIARIO /SUBDIARIO SE ENCUENTRA CERRADO. VERIFICAR CON EL USUARIO RESPONSABLE"
                Exit Sub
            End If
            
            sEsCtaCte = DevuelveCampo("select DBO.CN_EvaluaCuenta_Contable_Es_Cta_Cte('" & sano & "','" & txtCuenta.Text & "')", cConnect)
            If UCase(sEsCtaCte) = "S" Then
                txtSer_Docum.Enabled = False
                txtNum_Docum.Enabled = False
                txtDoc_Sunat.Enabled = False
            Else
                txtSer_Docum.Enabled = True
                txtNum_Docum.Enabled = True
                txtDoc_Sunat.Enabled = True
            End If
            
            txtDes_Cuenta.Text = GridEX1.Value(GridEX1.Columns("DESCRIPCION").Index)
            txtTipodeCambio.Text = GridEX1.Value(GridEX1.Columns("TIPCAM").Index)
            smoneda = GridEX1.Value(GridEX1.Columns("Cod_Moneda_Docum").Index)
            sdebehaber = GridEX1.Value(GridEX1.Columns("FLG_DEBE_HABER").Index)
            txtDebe.Text = 0
            txtDebeDol.Text = 0
            txtHaber.Text = 0
            txtHaberDol.Text = 0

            If GridEX1.Value(GridEX1.Columns("FLG_DEBE_HABER").Index) = "D" Then
                txtDebe.Text = GridEX1.Value(GridEX1.Columns("IMPORTE").Index)
                txtDebeDol.Text = GridEX1.Value(GridEX1.Columns("DOLARES").Index)
            Else
                txtHaber.Text = GridEX1.Value(GridEX1.Columns("IMPORTE").Index)
                txtHaberDol.Text = GridEX1.Value(GridEX1.Columns("DOLARES").Index)
            End If
            txtGlosa.Text = GridEX1.Value(GridEX1.Columns("DESCRIPCIO").Index)
            txtDoc_Sunat.Text = RTrim(GridEX1.Value(GridEX1.Columns("TIPO").Index))
            txtDes_TipDoc.Text = RTrim(GridEX1.Value(GridEX1.Columns("DES_TIPDOC").Index))
            txtSer_Docum.Text = RTrim(GridEX1.Value(GridEX1.Columns("SERIE").Index))
            txtNum_Docum.Text = RTrim(GridEX1.Value(GridEX1.Columns("NUMERO").Index))
            txtCuenta.Enabled = True
            Me.fraDatos.Visible = True
            txtCuenta.SetFocus
        Case "ELIMINAR"
            If GridEX1.RowCount = 0 Then Exit Sub
            sAccion = "D"
            sItem = GridEX1.Value(GridEX1.Columns("item").Index)
            txtCuenta.Text = GridEX1.Value(GridEX1.Columns("CUENTA").Index)
            txtDes_Cuenta.Text = GridEX1.Value(GridEX1.Columns("DESCRIPCION").Index)
            txtTipodeCambio.Text = GridEX1.Value(GridEX1.Columns("TIPCAM").Index)
            
            txtDebe.Text = 0
            txtDebeDol.Text = 0
            txtHaber.Text = 0
            txtHaberDol.Text = 0
            
            sEsCtaCte = DevuelveCampo("select DBO.CN_EvaluaCuenta_Contable_Es_Cta_Cte('" & sano & "','" & txtCuenta.Text & "')", cConnect)
            If UCase(sEsCtaCte) = "S" Then
                txtSer_Docum.Enabled = False
                txtNum_Docum.Enabled = False
                txtDoc_Sunat.Enabled = False
            Else
                txtSer_Docum.Enabled = True
                txtNum_Docum.Enabled = True
                txtDoc_Sunat.Enabled = True
            End If
            
            If GridEX1.Value(GridEX1.Columns("FLG_DEBE_HABER").Index) = "D" Then
                txtDebe.Text = GridEX1.Value(GridEX1.Columns("IMPORTE").Index)
                txtDebeDol.Text = GridEX1.Value(GridEX1.Columns("DOLARES").Index)
            Else
                txtHaber.Text = GridEX1.Value(GridEX1.Columns("IMPORTE").Index)
                txtHaberDol.Text = GridEX1.Value(GridEX1.Columns("DOLARES").Index)
            End If
            txtGlosa.Text = GridEX1.Value(GridEX1.Columns("DESCRIPCIO").Index)
            txtDoc_Sunat.Text = RTrim(GridEX1.Value(GridEX1.Columns("TIPO").Index))
            txtDes_TipDoc.Text = RTrim(GridEX1.Value(GridEX1.Columns("DES_TIPDOC").Index))
            txtSer_Docum.Text = RTrim(GridEX1.Value(GridEX1.Columns("SERIE").Index))
            txtNum_Docum.Text = RTrim(GridEX1.Value(GridEX1.Columns("NUMERO").Index))
            txtCuenta.Enabled = False
            Me.fraDatos.Visible = True
        Case "SALIR"
            Unload Me
    End Select
End Sub

Private Sub Imprimir()
Dim oo As Object
Dim Ruta As String
Dim sModoEjecutarTransaccion As String
Dim strSQL As String
On Error GoTo errReporte
Dim sEmpresa As String


sEmpresa = DevuelveCampo("SELECT des_empresa FROM seg_empresas WHERE Cod_Empresa ='" & vemp & "'", cSEGURIDAD)

If RTrim(sCod_TipoDiario) = "21" Or RTrim(sCod_TipoDiario) = "24" Or RTrim(sCod_TipoDiario) = "25" Then
    Ruta = vRuta & "\RptVoucherVentas.XLT"
Else
    Ruta = vRuta & "\RptVoucherVentas2.XLT"
End If
Set oo = CreateObject("excel.application")
oo.Workbooks.Open Ruta
oo.Visible = False
oo.displayalerts = False
oo.Run "Reporte", cConnect, sCod_TipoDiario, sano, smes, lNum_Registro, Num_Corre, sEmpresa
'oo.Close
Set oo = Nothing


Exit Sub
errReporte:
    MsgBox Err.Description, vbCritical, "Print Voucher Finanzas"
End Sub

Public Sub Buscar()

On Error GoTo dprDepurar

Dim sSQL As String

sSQL = "Fi_Muestra_Voucher_Contable '$','$','$',$"
sSQL = VBsprintf(sSQL, sCod_TipoDiario, sano, smes, lNum_Registro)

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cConnect)


GridEX1.ContinuousScroll = True

Exit Sub

dprDepurar:

errores Err.Number
  
End Sub

Public Sub Generar()

On Error GoTo dprDepurar
Dim sSQL As String
Dim vConfirma As Variant
Dim mRs As ADODB.Recordset

vConfirma = MsgBox("Confirma Voucher de Ventas ?", vbYesNo + vbQuestion, "REVERSION")
If vConfirma = vbNo Then Exit Sub


sSQL = "CN_GENERA_ASIENTO_VENTAS_LETRAS_X_COBRAR '$'"
sSQL = VBsprintf(sSQL, Num_Corre)

ExecuteCommandSQL cConnect, sSQL
Mensaje kMESSAGE_INF_PROCESS_SATISFACTO

Set mRs = GetDataSet(cConnect, "FI_Muestra_Datos_Transaccion '" & Num_Corre & "'")
If Not mRs Is Nothing Then
    sCod_TipoDiario = mRs!Cod_TipoDiario
    sano = mRs!Ano_Contable
    smes = mRs!Mes_Contable
    lNum_Registro = mRs!Num_Registro
    dImporte = mRs!Importe
    mRs.Close
End If
Set mRs = Nothing

Buscar

Exit Sub

dprDepurar:

errores Err.Number
  
End Sub


Public Sub Revertir()
On Error GoTo dprDepurar
Dim sSQL As String
Dim vConfirma As Variant

vConfirma = MsgBox("Confirma REVERSION DE MOVIMIENTO BANCARIO ?", vbYesNo + vbQuestion, "REVERSION")
If vConfirma = vbNo Then Exit Sub


sSQL = "CN_REVIERTE_ASIENTO_VENTAS '$'"
sSQL = VBsprintf(sSQL, Num_Corre)

ExecuteCommandSQL cConnect, sSQL
Mensaje kMESSAGE_INF_PROCESS_SATISFACTO

Buscar

Exit Sub

dprDepurar:

errores Err.Number
  
End Sub




Private Sub FunctButt3_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            GrabarCuenta
        Case "CANCELAR"
            Me.fraDatos.Visible = False
    End Select
End Sub

Private Sub txtCuenta_GotFocus()
    SelectionText txtCuenta
End Sub

Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
      txtDes_Cuenta.SetFocus
End If
End Sub

Private Sub txtCUENTA_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And txtCuenta <> "" Then
        If RTrim(txtCuenta.Text) = "" Then
            BUSCA_CUENTACONTABLE 3
        Else
            BUSCA_CUENTACONTABLE 1
        End If
    Else
        If KeyAscii = vbKeyReturn Then
            FunctButt3.SetFocus
        End If
    End If
    
End Sub


Private Sub BUSCA_CUENTACONTABLE(Tipo As Integer)
On Error GoTo errx
Dim strSQL  As String

    Select Case Tipo
    Case 1:
        strSQL = "SELECT COD_CTACONT as 'Código', DES_CTACONT as 'Descripción' " & _
                 "FROM CN_PLANCONTABLE  WHERE ANO = '" & sano & "' AND   COD_CTACONT like '" & Trim(txtCuenta.Text) & "%' ORDER BY COD_CTACONT"
                
    Case 2, 3:
            strSQL = "SELECT COD_CTACONT AS 'Código', " & _
            " DES_CTACONT as 'Descripción' " & _
            "FROM CN_PLANCONTABLE " & _
            "WHERE ANO = '" & sano & "' AND  DES_CTACONT LIKE '%" & Trim(Me.txtDes_Cuenta.Text) _
            & "%' AND DATALENGTH(RTRIM(COD_CTACONT )) = 8 ORDER BY 2"
    End Select
    
    With frmBusqGeneral3
        .Caption = "Buscar Cuenta"
        .sQuery = strSQL
        .Cargar_Datos
        Set .oParent = Me
        
        .gexLista.Columns("Código").Caption = "Código"
        .gexLista.Columns("Descripción").Caption = "Desc. Cuenta"
        
        .gexLista.Columns("Descripción").Width = 4800
                
        If .gexLista.RowCount > 1 Then
            .Show vbModal
        Else
            codigo = .gexLista.Value(.gexLista.Columns("Código").Index)
            Descripcion = .gexLista.Value(.gexLista.Columns("Descripción").Index)
        End If
            
        
        If .gexLista.RowCount > 0 And Not .bCancel Then
            txtCuenta = codigo
            txtDes_Cuenta = Descripcion
'            FunctButt3.SetFocus
            If sdebehaber = "D" Then
                If smoneda = "SOL" Then
                    txtDebe.SetFocus
                Else
                    txtDebeDol.SetFocus
                End If
            Else
            If smoneda = "SOL" Then
                    txtHaber.SetFocus
                Else
                    txtHaberDol.SetFocus
                End If
            
            End If
            
            codigo = "": Descripcion = ""
        End If
    End With
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    Exit Sub
    
errx:
    errores Err.Number
End Sub

Private Sub GrabarCuenta()
On Error GoTo errx
Dim sSQL As String
Dim sTipodeImpo As String
Dim xImporte As Double
Dim xImporteDolares As Double

If Val(txtDebe) + Val(txtDebeDol) > 0 Then
    sTipodeImpo = "1"
    xImporte = txtDebe
    xImporteDolares = txtDebeDol

Else
    sTipodeImpo = "2"
    xImporte = txtHaber
    xImporteDolares = txtHaberDol
End If

sSQL = "CN_CambiaCuenta_Movim '$','$','$','$','$','$',$ ,'$',$ ,'$',$,$,'$','$','$','$','$','$',$"
sSQL = VBsprintf(sSQL, sano, smes, sCod_TipoDiario, GridEX1.Value(GridEX1.Columns("comprob").Index), sItem, txtCuenta.Text, 0, sAccion, txtTipodeCambio.Text, sTipodeImpo, xImporte, xImporteDolares, txtGlosa.Text, txtDoc_Sunat.Text, txtSer_Docum.Text, txtNum_Docum.Text, Num_Corre, sFec_Transaccion, sSecuencia)

ExecuteCommandSQL cConnect, sSQL

Me.fraDatos.Visible = False
Buscar

Exit Sub
errx:
    errores Err.Number
End Sub


Private Sub txtDebe_GotFocus()
    SelectionText txtDebe
End Sub

Private Sub txtDebe_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
    txtGlosa.SetFocus
End If
If KeyCode = 40 Then
    
    txtHaber.SetFocus
End If
End Sub

Private Sub txtDebe_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        txtDebeDol.SetFocus
'        txtHaber.Text = 0
'        txtHaberDol.Text = 0
'    Else
'        txtHaber.Text = "0"
'
'    End If
    
    If KeyAscii = vbKeyReturn Then
        txtDebeDol.Text = Round(Val(txtDebe.Text) / Val(txtTipodeCambio.Text), 2)
    If txtDebe.Text = 0 Then
            txtHaber.SetFocus
    Else
            FunctButt3.SetFocus
    End If
    End If
End Sub

Private Sub txtDebeDol_GotFocus()
    SelectionText txtDebeDol
End Sub

Private Sub txtDebeDol_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
    txtHaber.SetFocus
End If
If KeyCode = 40 Then
    txtHaberDol.SetFocus
End If
End Sub

Private Sub txtdebedol_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        txtDebe.Text = Round(Val(txtDebeDol.Text) * Val(txtTipodeCambio.Text), 2)
'        txtHaberDol.SetFocus
'    End If

    If KeyAscii = vbKeyReturn Then
        txtDebe.Text = Round(Val(txtDebeDol.Text) * Val(txtTipodeCambio.Text), 2)
'        txtHaberDol.SetFocus
    If txtDebeDol.Text = 0 Then
            txtHaberDol.SetFocus
    Else
            FunctButt3.SetFocus
    End If
    End If
End Sub

Private Sub txtDoc_Sunat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Busca_TipoDocuSunat
    End If
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If txtDoc_Sunat.Enabled Then
            txtDoc_Sunat.SetFocus
        Else
            FunctButt3.SetFocus
        End If
    End If
End Sub

Private Sub txtHaber_GotFocus()
    SelectionText txtHaber
End Sub

Private Sub txtHaber_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
    txtDebe.SetFocus
End If

End Sub

Private Sub txtHaber_KeyPress(KeyAscii As Integer)
     
'    If KeyAscii = vbKeyReturn Then
'        FunctButt3.SetFocus
'        txtDebeDol.Text = 0
'        txtDebe.Text = 0
'    Else
'        txtDebe.Text = "0"
'    End If

        If KeyAscii = vbKeyReturn Then
        txtHaberDol.Text = Round(Val(txtHaber.Text) / Val(txtTipodeCambio.Text), 2)
        FunctButt3.SetFocus
    End If
End Sub

Private Sub txtHaberDol_GotFocus()
    SelectionText txtHaberDol
End Sub

Private Sub txtHaberDol_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
    txtDebeDol.SetFocus
End If
End Sub

Private Sub txtHaberDol_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        txtHaber.Text = Round(Val(txtHaberDol.Text) * Val(txtTipodeCambio.Text), 2)
'        txtDebe.Text = 0
'        txtDebeDol.Text = 0
'        FunctButt3.SetFocus
'    End If

    If KeyAscii = vbKeyReturn Then
        txtHaber.Text = Round(Val(txtHaberDol.Text) * Val(txtTipodeCambio.Text), 2)
        FunctButt3.SetFocus
    End If
End Sub

Sub Busca_TipoDocuSunat()
Dim oTipo As New frmBusqGeneral3
Dim RS As Object
Set RS = CreateObject("ADODB.Recordset")

Set oTipo.oParent = Me
oTipo.sQuery = "select Doc_Sunat as Codigo, cod_TipDoc as Tip_Doc, Des_TipDoc as Descripcion from cn_tiposdocum where Doc_Sunat LIKE '%" & Trim(txtDoc_Sunat.Text) & "%' order by Doc_Sunat"

oTipo.Caption = "Buscar Documento"
oTipo.Cargar_Datos

oTipo.gexLista.Columns("Codigo").Width = 900
oTipo.gexLista.Columns("Tip_Doc").Width = 900
oTipo.gexLista.Columns("Descripcion").Width = 3800

If oTipo.gexLista.RowCount > 1 Then
    oTipo.Show vbModal
Else
    codigo = oTipo.gexLista.Value(oTipo.gexLista.Columns("Codigo").Index)
    Descripcion = oTipo.gexLista.Value(oTipo.gexLista.Columns("tip_doc").Index)
    TipoAdd = oTipo.gexLista.Value(oTipo.gexLista.Columns("Descripcion").Index)
End If

If Trim(codigo) <> "" Then
    txtDoc_Sunat = codigo 'oTipo.gexLista.Value(oTipo.gexLista.Columns("Codigo").Index)
    txtDes_TipDoc.Text = TipoAdd
    codigo = "": Descripcion = "": TipoAdd = ""
    txtSer_Docum.SetFocus
Else
    txtDoc_Sunat.SetFocus
End If
Unload oTipo
Set oTipo = Nothing
Set RS = Nothing
End Sub

Private Sub txtDoc_Sunat_GotFocus()
SelectionText txtDoc_Sunat
End Sub

