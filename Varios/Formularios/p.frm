VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Begin VB.Form FrmAddOCOtrosCliente 
   Caption         =   "Mantenimiento OC Otros Clientes"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   6825
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDatos 
      Height          =   2490
      Left            =   0
      TabIndex        =   10
      Top             =   -30
      Width           =   6825
      Begin VB.Frame FraMod 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   6495
         Begin VB.TextBox TxtNom_moneda 
            Height          =   285
            Left            =   2205
            TabIndex        =   3
            Top             =   0
            Width           =   4215
         End
         Begin VB.TextBox TxtCod_Moneda 
            Height          =   285
            Left            =   1440
            TabIndex        =   2
            Top             =   0
            Width           =   735
         End
         Begin VB.TextBox TxtPrecio 
            Height          =   285
            Left            =   1440
            TabIndex        =   4
            Top             =   345
            Width           =   735
         End
         Begin VB.TextBox TxtIGV 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3360
            TabIndex        =   5
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox TxtDes_Condicion 
            Height          =   285
            Left            =   2205
            TabIndex        =   7
            Top             =   675
            Width           =   4215
         End
         Begin VB.TextBox TxtCod_Condicion 
            Height          =   285
            Left            =   1440
            TabIndex        =   6
            Top             =   675
            Width           =   735
         End
         Begin VB.TextBox TxtDes_Descuento 
            Height          =   285
            Left            =   2220
            TabIndex        =   8
            Top             =   1020
            Width           =   4215
         End
         Begin VB.TextBox Txtcod_Descuento 
            Height          =   285
            Left            =   1410
            TabIndex        =   9
            Top             =   1020
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Precio"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   465
            Width           =   450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "IGV"
            Height          =   195
            Left            =   2640
            TabIndex        =   15
            Top             =   450
            Width           =   270
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Condicion Venta"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   810
            Width           =   1170
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Descuentos"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   1155
            Width           =   855
         End
      End
      Begin VB.TextBox TxtCod_Cliente 
         Height          =   285
         Left            =   1590
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox TxtDes_Cliente 
         Height          =   285
         Left            =   2355
         TabIndex        =   1
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Otro Cliente"
         Height          =   195
         Left            =   270
         TabIndex        =   11
         Top             =   465
         Width           =   825
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2520
      TabIndex        =   18
      Top             =   2640
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"p.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "FrmAddOCOtrosCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sAccion As String
Public sCod_Cliente As String, sSer_OrdComp As String, sCod_Ordcomp As String, sSec_Ordcomp As String
Dim strSQL As String
Public Codigo As String, Descripcion As String, TipoAdd As String
Dim sOtroCliente  As String

Sub BUSCA_CLIENTE(opcion As Integer)

On Error GoTo Fin

Dim rstAux As ADODB.Recordset
    If opcion = 1 Then
        strSQL = "Select Abr_Cliente as Codigo , Nom_Cliente as Descripcion from tx_cliente where abr_cliente like '%" & Trim(TxtCod_Cliente.Text) & "' order by abr_cliente"
    Else
        strSQL = "Select Abr_Cliente as Codigo , Nom_Cliente as Descripcion from tx_cliente where nom_cliente like '%" & Trim(TxtDes_Cliente.Text) & "%' order by nom_cliente"
    End If
    
    With frmBusqGeneral
        Set .oParent = Me
        .SQuery = strSQL
        .CARGAR_DATOS
        
        Codigo = ".."
        Set rstAux = .gexList.ADORecordset
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            TxtCod_Cliente = Trim(rstAux!Codigo)
            TxtDes_Cliente = Trim(rstAux!Descripcion)
            TxtCod_Moneda.SetFocus
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
    "Búsqueda de Cliente (" & opcion & ")"
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    Call Agregar_OC
Case "CANCELAR"
    Unload Me
End Select
End Sub

Private Sub TxtCod_Cliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BUSCA_CLIENTE(1)
End If
End Sub

Private Sub TxtCod_Condicion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_CondicionVenta(1)
End If
End Sub

Private Sub Txtcod_Descuento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Descuento(1)
End If
End Sub

Private Sub TxtCod_Moneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Moneda(1)
End If
End Sub

Private Sub TxtDes_Cliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BUSCA_CLIENTE(2)
End If
End Sub


Sub Agregar_OC()
On Error GoTo errAgregar

sOtroCliente = DevuelveCampo("select cod_cliente_tex from tx_cliente where abr_cliente ='" & Trim(TxtCod_Cliente) & "'", cConnect)
 
strSQL = "up_man_Ti_ordcompitem_otros '" & sAccion & "','" & sCod_Cliente & "','" & sSer_OrdComp & "','" & sCod_Ordcomp & "','" & sSec_Ordcomp & "','" & sOtroCliente & "','" & _
            TxtCod_Condicion & "','" & Txtcod_Descuento & "'," & CDbl(TxtIGV) & "," & CDbl(TxtPrecio) & ",'" & TxtCod_Moneda & "'"
Call ExecuteCommandSQL(cConnect, strSQL)
Unload Me
Exit Sub
errAgregar:
    MsgBox Err.Description, vbCritical, "Agregar OC otros Cliente"
End Sub

Sub Busca_Moneda(opcion As Integer)

On Error GoTo Fin

Dim rstAux As ADODB.Recordset
    If opcion = 1 Then
        strSQL = "Select cod_moneda as Codigo, nom_moneda as Descripcion from tg_moneda where cod_moneda like '%" & Trim(TxtCod_Moneda) & "%' order by cod_moneda"
    Else
        strSQL = "Select cod_moneda as Codigo, nom_moneda as Descripcion from tg_moneda where nom_moneda like '%" & Trim(TxtNom_moneda) & "%' order by nom_moneda"
    End If
    
    With frmBusqGeneral
        Set .oParent = Me
        .SQuery = strSQL
        .CARGAR_DATOS
        
        Codigo = ".."
        Set rstAux = .gexList.ADORecordset
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            TxtCod_Moneda = Trim(rstAux!Codigo)
            TxtNom_moneda = Trim(rstAux!Descripcion)
            TxtPrecio.SetFocus
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
    "Búsqueda de Moneda (" & opcion & ")"
End Sub

Private Sub TxtDes_Condicion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_CondicionVenta(2)
End If
End Sub

Private Sub TxtDes_Descuento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Descuento(2)
End If
End Sub

Private Sub TxtIGV_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(TxtPrecio, KeyAscii, True, 3)
End If
End Sub

Private Sub TxtNom_moneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Moneda(2)
End If
End Sub

Private Sub TxtPrecio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(TxtPrecio, KeyAscii, True, 3)
End If
End Sub

Sub Busca_CondicionVenta(opcion As Integer)

On Error GoTo Fin

Dim rstAux As ADODB.Recordset
    If opcion = 1 Then
        strSQL = "Select cod_condvent as Codigo, Des_CondVent as Descripcion from lg_condvent where cod_condvent like '%" & Trim(TxtCod_Condicion) & "%' order by cod_condvent"
    Else
        strSQL = "Select cod_condvent as Codigo, Des_CondVent as Descripcion from lg_condvent where des_condvent like '%" & Trim(TxtDes_Condicion) & "%' order by des_condvent"
    End If
    
    With frmBusqGeneral
        Set .oParent = Me
        .SQuery = strSQL
        .CARGAR_DATOS
        
        Codigo = ".."
        Set rstAux = .gexList.ADORecordset
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            TxtCod_Condicion = Trim(rstAux!Codigo)
            Me.TxtDes_Condicion = Trim(rstAux!Descripcion)
            Txtcod_Descuento.SetFocus
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
    "Búsqueda de Condicion Venta (" & opcion & ")"
End Sub

Sub Busca_Descuento(opcion As Integer)

On Error GoTo Fin

Dim rstAux As ADODB.Recordset
    If opcion = 1 Then
        strSQL = "Select cod_descuento as Codigo, Des_descuento as Descripcion from Lg_Dsctos where cod_descuento like '%" & Trim(Txtcod_Descuento) & "%' order by cod_descuento"
    Else
        strSQL = "Select cod_descuento as Codigo, Des_descuento as Descripcion from Lg_Dsctos where Des_descuento like '%" & Trim(TxtDes_Condicion) & "%' order by Des_descuento"
    End If
    
    With frmBusGeneral6
        Set .oParent = Me
        .SQuery = strSQL
        .CARGAR_DATOS
        
        Codigo = ".."
        Set rstAux = .DGridLista.ADORecordset
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            Txtcod_Descuento = Trim(rstAux!Codigo)
            TxtDes_Descuento = Trim(rstAux!Descripcion)
            FunctButt1.SetFocus
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
    "Búsqueda de Descuento (" & opcion & ")"
End Sub

