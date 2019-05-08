VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmShowCN_DocumVoucher_Ventas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voucher Contable"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   11025
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDebe 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   8445
      TabIndex        =   4
      Text            =   "0"
      Top             =   5010
      Width           =   1230
   End
   Begin VB.TextBox txtHaber 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   9720
      TabIndex        =   3
      Text            =   "0"
      Top             =   5010
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
      Height          =   315
      Left            =   8460
      TabIndex        =   2
      Text            =   "Total Debe"
      Top             =   4590
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
      Height          =   315
      Left            =   9720
      TabIndex        =   1
      Text            =   "Total Haber"
      Top             =   4590
      Width           =   1230
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4410
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   7779
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      GroupByBoxVisible=   0   'False
      DataMode        =   1
      ColumnHeaderHeight=   285
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmShowCN_DocumVoucher_Ventas.frx":0000
      FormatStyle(2)  =   "frmShowCN_DocumVoucher_Ventas.frx":0138
      FormatStyle(3)  =   "frmShowCN_DocumVoucher_Ventas.frx":01E8
      FormatStyle(4)  =   "frmShowCN_DocumVoucher_Ventas.frx":029C
      FormatStyle(5)  =   "frmShowCN_DocumVoucher_Ventas.frx":0374
      FormatStyle(6)  =   "frmShowCN_DocumVoucher_Ventas.frx":042C
      FormatStyle(7)  =   "frmShowCN_DocumVoucher_Ventas.frx":050C
      ImageCount      =   0
      PrinterProperties=   "frmShowCN_DocumVoucher_Ventas.frx":052C
   End
End
Attribute VB_Name = "frmShowCN_DocumVoucher_Ventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sNum_Corre As String
Public sFlg_TipMondoc As String
Public oParent As Object
Public dTipoCambio As Double
Public sTipAnexo As String
Public sCod_Anexo As String
Public sSubdiario As String
Public sAno_Registro As String
Public sMes_Registro As String
Public sNum_Movimiento As String
Public sTipOpcion As String

Public Function Buscar() As Boolean
On Error GoTo errores
Dim sSQL As String
Dim vBookmark As Variant
sSQL = "SM_VOUCHER_CONTABLE '$' ,'$' , '$','$','$','$'"
sSQL = VBsprintf(sSQL, sNum_Corre, sTipOpcion, sSubdiario, sAno_Registro, sMes_Registro, sNum_Movimiento)

vBookmark = GridEX1.Row
GridEX1.ClearFields

Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sSQL, cConnect)

GridEX1.Row = vBookmark

If GridEX1.RowCount > 0 Then
    txtDebe = Format(GridEX1.Value(GridEX1.Columns("TOTAL_DEBE").Index), "###,##0.00")
    txtHaber = Format(GridEX1.Value(GridEX1.Columns("TOTAL_HABER").Index), "###,##0.00")
End If

GridEX1.Columns("IMPORTE").Format = "###,##0.00"

GridEX1.ContinuousScroll = True

GridEX1.FrozenColumns = 2
Exit Function

errores:
    ErrorHandler Err, "Busca Voucher"
End Function

Private Sub Form_Load()
    sTipOpcion = "3"
    sSubdiario = "41"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If txtDebe <> txtHaber Then
'        Cancel = 1
'    End If
End Sub

Private Sub GridEX1_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = True
End Sub

Sub Imprimir()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String
Dim Cadena1, Cadena2 As String

Cadena1 = "EXEC CN_Muestra_Cabecera_Vocuher_Contabilidad '" & sNum_Corre & "'"
Cadena2 = "EXEC CN_Muestra_Detalle_Vocuher_Contabilidad '" & sNum_Corre & "'"

    If DevuelveCampo("select cod_tipdoc from cn_docum where num_corre = '" & sNum_Corre & "'", cConnect) <> "AT" Then
        Ruta = vRuta & "\VoucherContabilidad.XLT"
    Else
        Ruta = vRuta & "\VoucherAnticipos.XLT"
    End If
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "reporte", Cadena1, Cadena2, cConnect
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler Err, "GeneraReportes"
    Set oo = Nothing
End Sub
