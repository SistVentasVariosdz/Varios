VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Begin VB.Form frmTransporte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transporte"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtReg_Transportista 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   5400
      TabIndex        =   4
      Top             =   4095
      Width           =   1455
   End
   Begin VB.TextBox txtPlaca 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3405
      MaxLength       =   20
      TabIndex        =   2
      Top             =   3720
      Width           =   960
   End
   Begin VB.TextBox TxtPeso 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   1785
      TabIndex        =   3
      Top             =   4095
      Width           =   1455
   End
   Begin VB.TextBox TxtMarca 
      Enabled         =   0   'False
      Height          =   315
      Left            =   885
      MaxLength       =   20
      TabIndex        =   1
      Top             =   3735
      Width           =   1770
   End
   Begin VB.TextBox TxtSecuencia 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1245
      TabIndex        =   9
      Top             =   3360
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7695
      Begin GridEX20.GridEX gexTrans 
         Height          =   2775
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4895
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         BackColorBkg    =   -2147483624
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         FormatStylesCount=   7
         FormatStyle(1)  =   "FrmTransporte.frx":0000
         FormatStyle(2)  =   "FrmTransporte.frx":0138
         FormatStyle(3)  =   "FrmTransporte.frx":01E8
         FormatStyle(4)  =   "FrmTransporte.frx":029C
         FormatStyle(5)  =   "FrmTransporte.frx":0374
         FormatStyle(6)  =   "FrmTransporte.frx":042C
         FormatStyle(7)  =   "FrmTransporte.frx":050C
         ImageCount      =   0
         PrinterProperties=   "FrmTransporte.frx":052C
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2880
      TabIndex        =   13
      Top             =   4560
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"FrmTransporte.frx":0704
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin VB.Label Label3 
      Caption         =   "Registro de Transp :"
      Height          =   255
      Left            =   3930
      TabIndex        =   12
      Top             =   4140
      Width           =   1440
   End
   Begin VB.Label Label8 
      Caption         =   "Placa :"
      Height          =   255
      Left            =   2745
      TabIndex        =   11
      Top             =   3750
      Width           =   585
   End
   Begin VB.Label Label7 
      Caption         =   "Kg"
      Height          =   255
      Left            =   3285
      TabIndex        =   10
      Top             =   4155
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "Peso Vehicular :"
      Height          =   255
      Left            =   225
      TabIndex        =   7
      Top             =   4170
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Marca :"
      Height          =   255
      Left            =   225
      TabIndex        =   6
      Top             =   3765
      Width           =   585
   End
   Begin VB.Label Label1 
      Caption         =   "Secuencia :"
      Height          =   255
      Left            =   225
      TabIndex        =   5
      Top             =   3375
      Width           =   975
   End
End
Attribute VB_Name = "frmTransporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrSql As String
Dim Estado As String
Dim sTipo As String
Dim cadena As String
Sub CARGA_GRID()
On Error GoTo hand

StrSql = "EXEC UP_MAN_LG_TRANSPORTE 'S'"

Set Me.gexTrans.ADORecordset = CargarRecordSetDesconectado(StrSql, cConnect)
'ConfigurarGrid

Exit Sub
hand:
ErrorHandler Err, "CARGA_GRID"
End Sub

Sub ConfigurarGrid()
        gexTrans.Columns("Sec_Transp").Visible = False
        
        gexTrans.Columns("Marca").Width = 2500
        gexTrans.Columns("Placa").Width = 1000
        gexTrans.Columns("Peso").Width = 2500
End Sub


Private Sub Form_Load()
    CARGA_GRID
End Sub

Sub CARGA_DATOS()
On Error GoTo hand
    TxtSecuencia.Text = gexTrans.Value(gexTrans.Columns("Sec_Transp").Index)
    TxtMarca.Text = Trim(gexTrans.Value(gexTrans.Columns("Marca").Index))
    txtPlaca.Text = Trim(gexTrans.Value(gexTrans.Columns("Placa").Index))
    TxtPeso.Text = Trim(gexTrans.Value(gexTrans.Columns("Peso").Index))
    txtReg_Transportista = Trim(gexTrans.Value(gexTrans.Columns("Reg_Transportista").Index))
Exit Sub
hand:
ErrorHandler Err, "CARGA_DATOS"
End Sub

Private Sub gexTrans_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    If gexTrans.RowCount = 0 Then Exit Sub
    CARGA_DATOS
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
On Error GoTo hand
Select Case ActionName
    Case "ADICIONAR"
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        LIMPIA
        HABILITA_CAMPOS True
        Estado = "NUEVO"
        sTipo = "I"
        TxtMarca.SetFocus
    Case "MODIFICAR"
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Estado = "MODIFICAR"
        sTipo = "U"
        HABILITA_CAMPOS True
        TxtMarca.SetFocus
    Case "ELIMINAR"
        sTipo = "D"
        SALVAR_DATOS
        LIMPIA
        CARGA_GRID
        HABILITA_CAMPOS False
        sTipo = ""
    Case "GRABAR"
        If Trim(Me.txtPlaca.Text) = "" Then MsgBox "Ingrese el numero de placa del auto", vbInformation: Exit Sub
'        If CmbTipItem = "" Then MsgBox "Seleccione un tipo de item", vbInformation: Exit Sub
        
        SALVAR_DATOS
        LIMPIA
        HABILITA_CAMPOS False
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        CARGA_GRID
        sTipo = ""
    Case "DESHACER"
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        LIMPIA
        CARGA_GRID
        HABILITA_CAMPOS False
        sTipo = ""
    Case "SALIR"
        Unload Me
End Select

Exit Sub
hand:
ErrorHandler Err, "MantFunc1_ActionClick"
End Sub

Sub HABILITA_CAMPOS(vEstado As Boolean)
    TxtMarca.Enabled = vEstado
    txtPlaca.Enabled = vEstado
    TxtPeso.Enabled = vEstado
    txtReg_Transportista.Enabled = vEstado
End Sub

Sub LIMPIA()
    TxtSecuencia = ""
    TxtMarca = ""
    txtPlaca = ""
    TxtPeso = "0.00"
    txtReg_Transportista = ""
End Sub

Sub SALVAR_DATOS()
On Error GoTo hand
    
    StrSql = "EXEC UP_MAN_LG_TRANSPORTE '" & sTipo & "','" & _
    TxtSecuencia.Text & "','" & Trim(Me.TxtMarca) & "','" & _
    Trim(txtPlaca) & "'," & Me.TxtPeso.Text & ", '" & _
    txtReg_Transportista & "'"
    
    Call ExecuteSQL(cConnect, StrSql)
    
Exit Sub
hand:
ErrorHandler Err, "SALVAR_DATOS"
End Sub

Private Sub TxtMarca_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub TxtPeso_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        Me.MantFunc1.SetFocus
    Case vbKeyEscape
        'Nada
    Case Else
        Call SoloNumeros(TxtPeso, KeyAscii, True, 2)
    End Select
End Sub

Private Sub txtPlaca_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtReg_Transportista_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
