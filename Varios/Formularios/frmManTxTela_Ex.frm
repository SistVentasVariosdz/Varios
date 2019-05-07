VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "mantenimientos.ocx"
Begin VB.Form frmManTxTela_Ex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar / Modificar Telas"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   10320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDetalle 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   90
      TabIndex        =   14
      Top             =   2970
      Width           =   10140
      Begin VB.TextBox txtCod_Tela 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1335
         TabIndex        =   0
         Top             =   270
         Width           =   1155
      End
      Begin VB.ComboBox cboTip_Ancho 
         Height          =   315
         Left            =   3975
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   960
         Width           =   1530
      End
      Begin VB.ComboBox cboUniMedCnf 
         Height          =   315
         Left            =   1335
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   1395
      End
      Begin VB.TextBox txtDes_grutela 
         Height          =   285
         Left            =   2235
         MaxLength       =   10
         TabIndex        =   7
         Top             =   1335
         Width           =   3660
      End
      Begin VB.TextBox txtCod_grutela 
         Height          =   285
         Left            =   1335
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1335
         Width           =   870
      End
      Begin VB.TextBox txtDes_TipTela 
         Height          =   285
         Left            =   2235
         MaxLength       =   10
         TabIndex        =   3
         Top             =   600
         Width           =   3660
      End
      Begin VB.TextBox txtDes_Tela 
         Height          =   285
         Left            =   2535
         TabIndex        =   1
         Top             =   270
         Width           =   3360
      End
      Begin VB.TextBox txtCod_TipTela 
         Height          =   285
         Left            =   1335
         MaxLength       =   10
         TabIndex        =   2
         Top             =   600
         Width           =   870
      End
      Begin VB.TextBox txtGramaje_Acab 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   7230
         TabIndex        =   8
         Text            =   "0"
         Top             =   270
         Width           =   750
      End
      Begin VB.TextBox txtAncho_Acab 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7230
         MaxLength       =   4
         TabIndex        =   9
         Text            =   "0"
         Top             =   630
         Width           =   750
      End
      Begin VB.TextBox txtEncog_Ancho 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7230
         TabIndex        =   11
         Text            =   "0"
         Top             =   1380
         Width           =   750
      End
      Begin VB.TextBox txtEncog_Largo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7230
         TabIndex        =   10
         Text            =   "0"
         Top             =   990
         Width           =   750
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Ancho :"
         Height          =   195
         Left            =   2880
         TabIndex        =   23
         Top             =   1005
         Width           =   915
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Unid. Medida :"
         Height          =   195
         Left            =   90
         TabIndex        =   22
         Top             =   1035
         Width           =   1035
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Grupo Tela :"
         Height          =   195
         Left            =   105
         TabIndex        =   21
         Top             =   1410
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Acab : "
         Height          =   195
         Left            =   6015
         TabIndex        =   20
         Top             =   705
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Gramaje Acab :"
         Height          =   195
         Left            =   6015
         TabIndex        =   19
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Tela :"
         Height          =   195
         Left            =   90
         TabIndex        =   18
         Top             =   660
         Width           =   990
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Tela :"
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   330
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Encog. Ancho :"
         Height          =   195
         Left            =   6015
         TabIndex        =   16
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "Encog. Largo :"
         Height          =   195
         Left            =   6015
         TabIndex        =   15
         Top             =   1065
         Width           =   1125
      End
   End
   Begin GridEX20.GridEX gexLista 
      Height          =   2490
      Left            =   165
      TabIndex        =   12
      Top             =   315
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   4392
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      BackColorBkg    =   -2147483624
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmManTxTela_Ex.frx":0000
      FormatStyle(2)  =   "frmManTxTela_Ex.frx":0138
      FormatStyle(3)  =   "frmManTxTela_Ex.frx":01E8
      FormatStyle(4)  =   "frmManTxTela_Ex.frx":029C
      FormatStyle(5)  =   "frmManTxTela_Ex.frx":0374
      FormatStyle(6)  =   "frmManTxTela_Ex.frx":042C
      FormatStyle(7)  =   "frmManTxTela_Ex.frx":050C
      ImageCount      =   0
      PrinterProperties=   "frmManTxTela_Ex.frx":052C
   End
   Begin VB.Frame FraLista 
      Caption         =   "Lista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   75
      TabIndex        =   13
      Top             =   90
      Width           =   10155
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   4080
      TabIndex        =   24
      Top             =   5040
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmManTxTela_Ex.frx":0704
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmManTxTela_Ex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs_Lista As ADODB.Recordset
Dim strSQL As String
Dim sTipo As String, rstAux As ADODB.Recordset
Public sCod_Cliente As String
Public Codigo As String, Descripcion As String

Private Sub FillUniMed()
    
    'strSQL = "SELECT Cod_UniMed, Des_UniMed FROM TG_UNIMED WHERE Cod_UniMed IN ('UN', 'MT')"
    strSQL = "SELECT Cod_UniMed, Des_UniMed FROM TG_UNIMED WHERE flg_telas = 'S'"
    
    cboUniMedCnf.Clear
    Set rstAux = CargarRecordSetDesconectado(strSQL, cConnect)
    With rstAux
    If .RecordCount > 0 Then .MoveFirst
    Do Until .EOF
        cboUniMedCnf.AddItem !Cod_UniMed & " " & !Des_UniMed
        .MoveNext
    Loop
    .Close
    End With
    Set rstAux = Nothing
    If cboUniMedCnf.ListCount > 0 Then cboUniMedCnf.ListIndex = 0
End Sub

Private Sub FillTipAncho()
    
    strSQL = "SELECT Tip_Ancho, Des_TipAncho FROM TG_TIPANC"
    
    cboTip_Ancho.Clear
    Set rstAux = CargarRecordSetDesconectado(strSQL, cConnect)
    With rstAux
    If .RecordCount > 0 Then .MoveFirst
    Do Until .EOF
        cboTip_Ancho.AddItem !Tip_Ancho & " " & !Des_TipAncho
        .MoveNext
    Loop
    .Close
    End With
    Set rstAux = Nothing
    If cboTip_Ancho.ListCount > 0 Then cboTip_Ancho.ListIndex = 0
    
End Sub


Private Sub cboTip_Ancho_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cboUniMedCnf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmdFirst_Click()
    If Not Rs_Lista.BOF Then
        Rs_Lista.MoveFirst
    End If
End Sub

Private Sub cmdLast_Click()
    If Not Rs_Lista.EOF Then
        Rs_Lista.MoveLast
    End If
End Sub

Private Sub cmdNext_Click()
    If Not Rs_Lista.EOF Then
        Rs_Lista.MoveNext
        If Rs_Lista.EOF Then
            Rs_Lista.MoveLast
        End If
    End If
End Sub

Private Sub cmdPrevious_Click()
    If Not Rs_Lista.BOF Then
        Rs_Lista.MovePrevious
        If Rs_Lista.BOF Then
            Rs_Lista.MoveFirst
        End If
    End If
End Sub

Sub LIMPIAR_DATOS()
    
    txtCod_Tela = ""
    txtDes_Tela = ""
    txtCod_TipTela = ""
    txtGramaje_Acab = 0
    txtAncho_Acab = 0
    cboUniMedCnf.ListIndex = 0
    txtEncog_Ancho = 0
    txtEncog_Largo = 0
    cboTip_Ancho.ListIndex = 0
    txtCod_grutela = ""
    
End Sub

Function VALIDA_DATOS() As Boolean
Dim NombreTabla As String
Dim CodigoTabla As String
    
    VALIDA_DATOS = False
    
    VALIDA_DATOS = True
    
End Function

Sub CARGA_DATOS()
On Error GoTo Fin
    If gexLista.RowCount = 0 Then Exit Sub
    
    Rs_Lista.AbsolutePosition = gexLista.RowIndex(gexLista.Row)
    
    If Not Rs_Lista.EOF And Not Rs_Lista.BOF Then
        
        txtCod_Tela = Rs_Lista!Codigo
        txtDes_Tela = Rs_Lista!Nombre
        txtCod_TipTela = Rs_Lista!Tipo_Tela
        BuscaTipoTela 1
        txtGramaje_Acab = Rs_Lista!Gramaje_Acab
        txtAncho_Acab = Rs_Lista!Ancho_Acab
        BuscaCombo Rs_Lista!Cod_UniMedCnf, 1, cboUniMedCnf
        txtEncog_Ancho = Rs_Lista!Encog_Ancho
        txtEncog_Largo = Rs_Lista!Encog_Largo
        BuscaCombo Rs_Lista!Tip_Ancho, 1, cboTip_Ancho
        txtCod_grutela = Rs_Lista!Cod_GruTela
        
    End If
Exit Sub
Fin:
End Sub

Sub HABILITA_DATOS()
    
    'txtCod_Tela.Enabled = True
    
    txtDes_Tela.Enabled = True
    txtCod_TipTela.Enabled = True
    txtGramaje_Acab.Enabled = True
    txtAncho_Acab.Enabled = True
    cboUniMedCnf.Enabled = True
    txtEncog_Ancho.Enabled = True
    txtEncog_Largo.Enabled = True
    cboTip_Ancho.Enabled = True
    txtCod_grutela.Enabled = True
    
    txtDes_TipTela.Enabled = True
    txtDes_grutela.Enabled = True
    
    txtDes_Tela.SetFocus
    
    
    
End Sub

Sub INHABILITA_DATOS()
    
    txtDes_Tela.Enabled = False
    txtCod_TipTela.Enabled = False
    txtGramaje_Acab.Enabled = False
    txtAncho_Acab.Enabled = False
    cboUniMedCnf.Enabled = False
    txtEncog_Ancho.Enabled = False
    txtEncog_Largo.Enabled = False
    cboTip_Ancho.Enabled = False
    txtCod_grutela.Enabled = False
    
    txtDes_TipTela.Enabled = False
    txtDes_grutela.Enabled = False
    
End Sub

Sub CARGA_GRID()
    
    'Esta cadena es para devolver el Codigo de Cliente
    strSQL = "EXEC TI_MUESTRA_TELAS_SERVICIO_POR_CLIENTE '" & sCod_Cliente & "'"
    Set Rs_Lista = CargarRecordSetDesconectado(strSQL, cConnect)
    
    Set gexLista.ADORecordset = Rs_Lista
    'Set DGridLista.DataSource = Rs_Lista
    'DGridLista.Refresh
    
    gexLista.Columns("Codigo").Width = 825
    gexLista.Columns("Nombre").Width = 1620
    gexLista.Columns("Tipo_Tela").Width = 480
    gexLista.Columns("Nombre_Tipo").Width = 645
    gexLista.Columns("Gramaje_Acab").Width = 1140
    gexLista.Columns("Ancho_Acab").Width = 1065
    gexLista.Columns("Cod_UniMedCnf").Width = 480
    gexLista.Columns("Encog_Ancho").Width = 1155
    gexLista.Columns("Encog_Largo").Width = 1095
    gexLista.Columns("Cod_GruTela").Width = 1080
    gexLista.Columns("Nombre_Grupo").Width = 1215
    gexLista.Columns("Tip_Ancho").Width = 345
    gexLista.Columns("Nombre_Ancho").Width = 1245
    
    If Rs_Lista.RecordCount > 0 Then
        gexLista.Enabled = True
        'DGridLista.Enabled = True
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR"
        Call CARGA_DATOS
    Else
        gexLista.Enabled = False
        'DGridLista.Enabled = False
        HabilitaMant Me.MantFunc1, "ADICIONAR"
        Call LIMPIAR_DATOS
    End If
End Sub

Sub SALVAR_DATOS()
    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Dim strSQL As String
    
    Con.ConnectionString = cConnect
    Con.Open
    
        Con.BeginTrans

        strSQL = "EXEC TI_UP_MAN_TX_TELA_SERVICIO '" & sTipo & _
        "', '" & sCod_Cliente & "', '" & txtCod_Tela & "', '" & _
        txtDes_Tela & "', '" & txtCod_TipTela & "', " & _
        txtGramaje_Acab & ", " & txtAncho_Acab & ", '" & _
        Left(cboUniMedCnf, 3) & "', " & txtEncog_Ancho & ", " & _
        txtEncog_Largo & ", '" & Left(cboTip_Ancho, 2) & "', '" & _
        txtCod_grutela & "'"
        
        Con.Execute strSQL

        Con.CommitTrans
        Dim amensaje As New clsMessages
        amensaje.Codigo = CodeMsg.kMESSAGE_INF_DATA_SAVE
        MsgBox amensaje.mTexto, vbInformation + vbOKOnly, "Guardar"
        'Informa "", amensaje
        
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub

Sub ELIMINAR_DATOS()
 Dim Con As New ADODB.Connection
    On Error GoTo Eliminar_DatosErr
   
'    Con.ConnectionString = cCONNECT
'    Con.Open
'    Con.BeginTrans
'
'        strSQL = "EXEC UP_MAN_ITEMPROV '" & _
'        sTipo & "','" & _
'        varCod_item & "','" & _
'        Trim(txtCod_Proveedor.Text) & "','" & _
'        Trim(txtCod_ItemProv.Text) & "'," & _
'        Trim(txtFac_EquiProv.Text) & ",'" & _
'        Trim(txtCod_UniMedProv.Text) & "',0,0"
'
'        Con.Execute strSQL
'
'    Con.CommitTrans
'
'    Dim amensaje As New clsMessages
'    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_DELETE
'    MsgBox amensaje.mTexto, vbInformation + vbOKOnly, "Guardar"
'    'Informa "", amensaje
    
Exit Sub
Eliminar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Eliminar_Datos"

End Sub

Private Sub Form_Load()
    'Call FormateaGrid(DGridLista)
    Call INHABILITA_DATOS
    'Me.MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    FillTipAncho
    FillUniMed
    CARGA_GRID
End Sub

Private Sub gexLista_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    Call CARGA_DATOS
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim ELIMINAR As Integer, irow As Long
    
    Select Case ActionName
        Case "ADICIONAR"
            sTipo = "I"
            LIMPIAR_DATOS
            HABILITA_DATOS
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            gexLista.Enabled = False
        Case "MODIFICAR"
            If gexLista.RowCount = 0 Then Exit Sub
            sTipo = "U"
            HABILITA_DATOS
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            gexLista.Enabled = False
        Case "ELIMINAR"
            ELIMINAR = MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?", vbInformation + vbYesNo, "Telas")
            If ELIMINAR = vbYes Then
                sTipo = "D"
                If VALIDA_DATOS Then
                    Call ELIMINAR_DATOS
                    Call CARGA_GRID
                    sTipo = ""
                End If
            End If
        Case "GRABAR"
            irow = gexLista.Row
            If VALIDA_DATOS Then
                Call SALVAR_DATOS
                Call CARGA_GRID
                Call INHABILITA_DATOS
                HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR"
                If sTipo = "U" Then gexLista.Row = irow
                gexLista.Enabled = True
                sTipo = ""
            End If
        Case "DESHACER"
            Call LIMPIAR_DATOS
            Call CARGA_DATOS
            Call INHABILITA_DATOS
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR"
            gexLista.Enabled = True
            sTipo = ""
         Case "SALIR"
            Unload Me
    End Select
End Sub

Private Sub txtAncho_Acab_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCod_grutela_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaGruTela 1
        SendKeys "{TAB}"
    End If
End Sub

Public Sub BuscaGruTela(opcion As Integer)
On Error GoTo Fin
Dim iCol As Long
    
'    Flg_ClientePropio = False
'    txtCod_FamGrupo.TabIndex = 19
'    txtDes_FamGrupo.TabIndex = 20
'    txtCod_Tela_tejeduria.TabIndex = 21
'    txtCod_Tela_tejeduria.Enabled = False
    
    txtCod_grutela = Trim(txtCod_TipTela)
    txtDes_grutela = Trim(txtDes_grutela)
    strSQL = "SELECT Cod_GruTela, Des_GruTela FROM TX_GRUTELA " & _
             "WHERE Cod_FamTela = 'SE' AND "
    
    Select Case opcion
    Case 1: strSQL = strSQL & "Cod_GruTela like '%" & txtCod_grutela & "%'"
    Case 2: strSQL = strSQL & "Cod_GruTela like '%" & txtCod_grutela & "%'"
    End Select
    strSQL = strSQL & " ORDER BY Des_GruTela"
    
    txtCod_grutela = ""
    txtDes_grutela = ""
    
    With frmBusGeneral6
        Set .oParent = Me
        .SQuery = strSQL
        .CARGAR_DATOS
        Codigo = ".."
        Set rstAux = .DGridLista.ADORecordset
        
        .DGridLista.Columns("Cod_GruTela").Caption = "Cod.Grupo"
        .DGridLista.Columns("Cod_GruTela").Width = 900
        .DGridLista.Columns("Des_GruTela").Caption = "Grupo Tela"
        .DGridLista.Columns("Des_GruTela").Width = 4000
        
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            txtCod_grutela = Trim(rstAux!Cod_GruTela)
            txtDes_grutela = Trim(rstAux!Des_GruTela)
        End If
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
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Grupo de Tela (" & opcion & ")"
End Sub

Private Sub txtCod_TipTela_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaTipoTela 1
        SendKeys "{TAB}"
    End If
End Sub

Public Sub BuscaTipoTela(opcion As Integer)
On Error GoTo Fin
Dim iCol As Long
    
'    Flg_ClientePropio = False
'    txtCod_FamGrupo.TabIndex = 19
'    txtDes_FamGrupo.TabIndex = 20
'    txtCod_Tela_tejeduria.TabIndex = 21
'    txtCod_Tela_tejeduria.Enabled = False
    
    txtCod_TipTela = Trim(txtCod_TipTela)
    txtDes_TipTela = Trim(txtDes_TipTela)
    strSQL = "SELECT Cod_TipTela, Des_TipTela FROM TG_TIPTELA WHERE "
    
    Select Case opcion
    Case 1: strSQL = strSQL & "Cod_TipTela like '%" & txtCod_TipTela & "%'"
    Case 2: strSQL = strSQL & "Des_TipTela like '%" & txtDes_TipTela & "%'"
    End Select
    strSQL = strSQL & " ORDER BY Des_TipTela"
    
    txtCod_TipTela = ""
    txtDes_TipTela = ""
    
    With frmBusGeneral6
        Set .oParent = Me
        .SQuery = strSQL
        .CARGAR_DATOS
        Codigo = ".."
        Set rstAux = .DGridLista.ADORecordset
        
        .DGridLista.Columns("Cod_TipTela").Caption = "Cod.Tipo"
        .DGridLista.Columns("Cod_TipTela").Width = 900
        .DGridLista.Columns("Des_TipTela").Caption = "Tipo de Tela"
        .DGridLista.Columns("Des_TipTela").Width = 4000
        
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            txtCod_TipTela = Trim(rstAux!Cod_TipTela)
            txtDes_TipTela = Trim(rstAux!Des_TipTela)
        End If
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
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Tipo de Tela (" & opcion & ")"
End Sub


Private Sub txtDes_grutela_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaGruTela 2
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtdes_tela_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDes_TipTela_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscaTipoTela 2
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtEncog_Ancho_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtEncog_Largo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtGramaje_Acab_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
