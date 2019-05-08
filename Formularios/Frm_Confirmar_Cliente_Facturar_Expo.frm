VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Begin VB.Form Frm_Confirmar_Cliente_Facturar_Expo 
   Caption         =   "Cliente a Facturar"
   ClientHeight    =   2910
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6435
   LinkTopic       =   "Form2"
   ScaleHeight     =   2910
   ScaleWidth      =   6435
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Cliente"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.TextBox txtAbr_Cliente 
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   690
      End
      Begin VB.TextBox txtNom_Cliente 
         Height          =   315
         Left            =   1725
         TabIndex        =   1
         Top             =   360
         Width           =   3360
      End
      Begin VB.Label Label2 
         Caption         =   "Ruc:"
         Height          =   210
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lbl_Ruc 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente:"
         Height          =   210
         Left            =   360
         TabIndex        =   3
         Top             =   420
         Width           =   615
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1680
      TabIndex        =   6
      Top             =   2280
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"Frm_Confirmar_Cliente_Facturar_Expo.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "Frm_Confirmar_Cliente_Facturar_Expo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flg_error As String
Public Ser_ORdCompX As String
Public Cod_Cliente_TexX As String
Public cod_ORdCompX  As String
Public CODIGO As String
Public descripcion As String
Dim rsx As ADODB.Recordset
Dim strsql As String

Private Sub txtAbr_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtAbr_Cliente.Text) = "" Then
            Call BUSCA_CLIENTE(3)
        Else
            Call BUSCA_CLIENTE(1)
        End If
    End If
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
Public Sub BUSCA_CLIENTE(Tipo As Integer)
    Dim oTipo As New frmBusGeneral6
    Dim rs As New ADODB.Recordset
    
    Select Case Tipo
        Case 1:
                    strsql = "EXEC TI_BUSCA_CLIENTE 1,'" & Trim(Me.txtAbr_Cliente.Text) & "','','" & vusu & "'"
                    'Me.TxtNom_Cliente.Text = Trim(DevuelveCampo(strsql, cConnect))
                    Set rs = CargarRecordSetDesconectado(strsql, cConnect)
                    
                    Me.txtNom_Cliente.Text = rs.Fields("Nom_Cliente").Value
                    Me.lbl_Ruc.Caption = rs.Fields("num_ruc").Value
                    
                    Set rs = Nothing
                    'If Trim(txtNom_Cliente.Text) <> "" Then CARGA_GRID
        Case 2, 3:
                    
                    Set oTipo.oParent = Me
                    
                    If Tipo = 2 Then
                        oTipo.SQuery = "EXEC TI_BUSCA_CLIENTE 2,'','" & Trim(txtNom_Cliente.Text) & "','" & vusu & "'"
                    Else
                        oTipo.SQuery = "EXEC TI_BUSCA_CLIENTE 3,'','','" & vusu & "'"
                    End If
                    
                    oTipo.CARGAR_DATOS
                    oTipo.DGridLista.Columns(2).Width = 3500
                    oTipo.Show 1
                    If CODIGO <> "" Then
                         Me.txtAbr_Cliente.Text = Trim(CODIGO)
                         Me.txtNom_Cliente.Text = Trim(descripcion)
                         Me.lbl_Ruc.Caption = Trim(DevuelveCampo("Select Num_Ruc From Tx_Cliente Where Abr_cliente='" & txtAbr_Cliente.Text & "'", cConnect))
'                         OptCliPend.SetFocus
                         CODIGO = "": descripcion = ""
                         'CARGA_GRID
                    End If
                    Set oTipo = Nothing
                    Set rs = Nothing
    End Select
    
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo xerror:
Select Case ActionName
        Case "Grabar"
        
        flg_error = "0"
            Validar_Campos
            If flg_error = "1" Then
                Exit Sub
            End If
            Grabar_Campos
        Case "Cancelar"
            Unload Me
End Select
Exit Sub
xerror:
errores err.Number
Exit Sub
End Sub

Sub Validar_Campos()

    
    If Trim(Me.txtAbr_Cliente) = "" Or Trim(Me.txtNom_Cliente) = "" Then
        MsgBox "No ha ingresado el cliente a facturar", vbInformation, "Validacion"
        flg_error = "1"
        txtAbr_Cliente.SetFocus
    End If
    
    
End Sub

Sub Grabar_Campos()
On Error GoTo xerror:
Dim filas As Integer
    StrsqlX = "EXec TX_Mant_Datos_Cliente_Facturar 'U','" & Ser_ORdCompX & "','" & cod_ORdCompX & "','" & Cod_Cliente_TexX & "','" & lbl_Ruc.Caption & "'"
    Call ExecuteSQL(cConnect, StrsqlX)
   ' If filas <> 1 Then
    '    MsgBox "Ocurrio un Problema informar a Sistemas"
   ' Else
   MsgBox "Se han registrado correctamente los datos del cliente a facturar", vbInformation, "Mensaje"
        Unload Me
   ' End If
    Exit Sub
xerror:
errores err.Number
Exit Sub
End Sub

Public Sub Carga_CAmpos()

    Set rsx = New ADODB.Recordset
    Set rsx = CargarRecordSetDesconectado("Exec TX_Mant_Datos_Cliente_Facturar 'L','" & Ser_ORdCompX & "','" & cod_ORdCompX & "','" & Cod_Cliente_TexX & "',''", cConnect)
    If Not rsx.EOF = True Then
    
    Me.txtAbr_Cliente = rsx.Fields("Abr_cliente").Value
    Me.txtNom_Cliente = rsx.Fields("Nom_Cliente").Value
    Me.lbl_Ruc = rsx.Fields("Num_Ruc").Value
    
    End If
End Sub
        

'Private Sub TxtCod_Pais_KeyPress(KeyAscii As Integer)
'  If KeyAscii = 13 Then
'        If Trim(TxtCod_Pais.Text) = "" Then
'            Call Me.BUSCA_Pais(3)
'        Else
'            Call Me.BUSCA_Pais(1)
'        End If
'    End If
'End Sub
'
'Private Sub TxtDes_Pais_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'        Call Me.BUSCA_Pais(2)
'    End If
'End Sub

'Public Sub BUSCA_Pais(Tipo As Integer)
'    Select Case Tipo
'        Case 1: 'select Cod_Pais,Descripcion from CN_PAISES
'                    strSQL = "SELECT Descripcion as 'Descripción' FROM CN_PAISES WHERE Cod_Pais = '" & Trim(TxtCod_Pais) & "'"
'                    Me.TxtCod_Pais.Text = Trim(DevuelveCampo(strSQL, cConnect))
'                    If Trim(TxtCod_Pais.Text) <> "" Then SendKeys "{TAB}", True
'        Case 2, 3:
'                    Dim oTipo As New frmBusqGeneral
'                    Dim rs As New ADODB.Recordset
'                    Set oTipo.oParent = Me
'
'                    If Tipo = 2 Then
'                        oTipo.sQuery = "SELECT Cod_Pais as 'Código', Descripcion as 'Descripción' FROM CN_PAISES WHERE Descripcion LIKE '%" & Trim(TxtDes_Pais.Text) & "%' ORDER BY Descripcion"
'                    Else
'                        oTipo.sQuery = "SELECT Cod_Pais as 'Código', Descripcion as 'Descripción' FROM CN_PAISES ORDER BY Descripcion"
'                    End If
'
'                    oTipo.CARGAR_DATOS
'                    oTipo.Show 1
'                    If Codigo <> "" Then
'                        Me.TxtCod_Pais = Trim(Codigo)
'                        Me.TxtDes_Pais.Text = Trim(descripcion)
'                        Codigo = "": descripcion = ""
'                        'MantFunc1.SetFocus
'                    End If
'                    Set oTipo = Nothing
'                    Set rs = Nothing
'    End Select
'
'End Sub
