VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Begin VB.Form Frm_Confirmar_Cliente_Expo 
   Caption         =   "Confirmar Datos del Cliente"
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
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.TextBox TxtCod_Pais 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox TxtDes_Pais 
         Height          =   285
         Left            =   2160
         TabIndex        =   8
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox TxtDireccion 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   1080
         Width           =   4455
      End
      Begin VB.TextBox TxtRazonSocial 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox TxtNIT 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "País:"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Direccion:"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Razon Social:"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "RUC/NIT :"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2160
      TabIndex        =   10
      Top             =   2400
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"Frm_Confirmar_Cliente_Expo.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "Frm_Confirmar_Cliente_Expo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flg_error As String
Public Ser_ORdCompX As String
Public Cod_Cliente_TexX As String
Public cod_ORdCompX  As String
Public Codigo As String
Public descripcion As String
Dim rsx As ADODB.Recordset
Dim strSQL As String

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
errores Err.Number
Exit Sub
End Sub

Sub Validar_Campos()
    If Trim(TxtNIT) = "" Then
        MsgBox "NO ha ingresado un RUC/NIT Valido ", vbInformation, "Validacion"
        flg_error = "1"
        TxtNIT.SetFocus
    End If
    
    If Trim(TxtCod_Pais) = "" Or Trim(TxtDes_Pais) = "" Then
        MsgBox "No ha ingresado un Pais", vbInformation, "Validacion"
        flg_error = "1"
        TxtCod_Pais.SetFocus
    End If
    
    If Trim(txtDireccion) = "" Then
        MsgBox "No ha ingresado una direccion", vbInformation, "Validacion"
        flg_error = "1"
        txtDireccion.SetFocus
    End If
    If Trim(TxtRazonSocial) = "" Then
        MsgBox "No ha ingresado una Razon Social", vbInformation, "Validacion"
        flg_error = "1"
        TxtRazonSocial.SetFocus
    End If
    
End Sub

Sub Grabar_Campos()
On Error GoTo xerror:
Dim filas As Integer
    StrsqlX = "EXec TX_Mant_Datos_Cliente_Confirmacion 'U','" & Ser_ORdCompX & "','" & cod_ORdCompX & "','" & Cod_Cliente_TexX & "','" & TxtCod_Pais & "','" & TxtNIT & "','" & TxtRazonSocial & "','" & txtDireccion & "'"
    Call ExecuteSQL(cConnect, StrsqlX)
   ' If filas <> 1 Then
    '    MsgBox "Ocurrio un Problema informar a Sistemas"
   ' Else
   MsgBox "Se han registrado correctamente los datos del cliente", vbInformation, "Mensaje"
        Unload Me
   ' End If
    Exit Sub
xerror:
errores Err.Number
Exit Sub
End Sub

Public Sub Carga_CAmpos()
    Set rsx = New ADODB.Recordset
    Set rsx = CargarRecordSetDesconectado("Exec TX_Mant_Datos_Cliente_Confirmacion 'L','" & Ser_ORdCompX & "','" & cod_ORdCompX & "','" & Cod_Cliente_TexX & "','','','',''", cConnect)
    TxtCod_Pais = rsx.Fields("Cod_Pais").Value
    TxtDes_Pais = rsx.Fields("Pais").Value
    TxtNIT = rsx.Fields("RUC_NIT").Value
    TxtRazonSocial = rsx.Fields("RAZON_SOCIAL").Value
    txtDireccion = rsx.Fields("DIRECCION").Value
End Sub
        

Private Sub TxtCod_Pais_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        If Trim(TxtCod_Pais.Text) = "" Then
            Call Me.BUSCA_Pais(3)
        Else
            Call Me.BUSCA_Pais(1)
        End If
    End If
End Sub

Private Sub TxtDes_Pais_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Call Me.BUSCA_Pais(2)
    End If
End Sub

Public Sub BUSCA_Pais(Tipo As Integer)
    Select Case Tipo
        Case 1: 'select Cod_Pais,Descripcion from CN_PAISES
                    strSQL = "SELECT Descripcion as 'Descripción' FROM CN_PAISES WHERE Cod_Pais = '" & Trim(TxtCod_Pais) & "'"
                    Me.TxtCod_Pais.Text = Trim(DevuelveCampo(strSQL, cConnect))
                    If Trim(TxtCod_Pais.Text) <> "" Then SendKeys "{TAB}", True
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If Tipo = 2 Then
                        oTipo.SQuery = "SELECT Cod_Pais as 'Código', Descripcion as 'Descripción' FROM CN_PAISES WHERE Descripcion LIKE '%" & Trim(TxtDes_Pais.Text) & "%' ORDER BY Descripcion"
                    Else
                        oTipo.SQuery = "SELECT Cod_Pais as 'Código', Descripcion as 'Descripción' FROM CN_PAISES ORDER BY Descripcion"
                    End If
                    
                    oTipo.CARGAR_DATOS
                    oTipo.Show 1
                    If Codigo <> "" Then
                        Me.TxtCod_Pais = Trim(Codigo)
                        Me.TxtDes_Pais.Text = Trim(descripcion)
                        Codigo = "": descripcion = ""
                        'MantFunc1.SetFocus
                    End If
                    Set oTipo = Nothing
                    Set rs = Nothing
    End Select
    
End Sub
