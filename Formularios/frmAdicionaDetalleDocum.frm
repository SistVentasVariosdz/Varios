VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{144A86C7-1AF0-44BA-9AA8-AF3AAF6043B8}#1.0#0"; "NumBox.ocx"
Begin VB.Form frmAdicionaDetalleDocum 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4995
   ClientLeft      =   345
   ClientTop       =   1020
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8295
   Begin VB.Frame Frame1 
      Height          =   4230
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8055
      Begin VB.TextBox txtCod_Producto 
         Height          =   285
         Left            =   2280
         MaxLength       =   25
         TabIndex        =   6
         Top             =   600
         Width           =   1200
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   645
         Left            =   2280
         MaxLength       =   53
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1005
         Width           =   5505
      End
      Begin VB.TextBox txtUnida_Medida 
         Height          =   285
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   4
         Top             =   1800
         Width           =   600
      End
      Begin VB.TextBox txtTip_Item 
         Height          =   285
         Left            =   2280
         MaxLength       =   1
         TabIndex        =   3
         Top             =   240
         Width           =   360
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   270
         Left            =   3600
         TabIndex        =   2
         Tag             =   "..."
         Top             =   600
         Visible         =   0   'False
         Width           =   300
      End
      Begin NumBoxProject.NumBox txtCantidad 
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Tag             =   "SET/VALID"
         Top             =   2235
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         TypeVal         =   2
         Mask            =   "9,999,999,999.99"
         Formato         =   "#,###,###,###.##"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   3
         Text            =   "0.00"
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   2
      End
      Begin NumBoxProject.NumBox txtImp_Unitario 
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Tag             =   "SET/VALID"
         Top             =   2595
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         TypeVal         =   2
         Mask            =   "9,999,999,999.9999"
         Formato         =   "#,###,###,###.####"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   3
         Text            =   "0.0000"
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   4
      End
      Begin NumBoxProject.NumBox txtImp_Total 
         Height          =   285
         Left            =   2280
         TabIndex        =   9
         Tag             =   "SET/VALID"
         Top             =   3015
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         TypeVal         =   2
         Mask            =   "9,999,999,999.99"
         Formato         =   "#,###,###,###.##"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   3
         Text            =   "0.00"
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   2
      End
      Begin NumBoxProject.NumBox txtPorc_Commision 
         Height          =   285
         Left            =   6600
         TabIndex        =   10
         Tag             =   "SET/VALID"
         Top             =   3000
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         TypeVal         =   2
         Mask            =   "9,999,999,999.99"
         Formato         =   "#,###,###,###.##"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   3
         Text            =   "0.00"
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   2
      End
      Begin NumBoxProject.NumBox txtCantUniAlter 
         Height          =   285
         Left            =   2265
         TabIndex        =   11
         Tag             =   "SET/VALID"
         Top             =   3510
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         TypeVal         =   2
         Mask            =   "9,999,999,999.99"
         Formato         =   "#,###,###,###.##"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   3
         Text            =   "0.00"
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   2
      End
      Begin VB.Label LblOtros 
         Caption         =   "Cantidad :"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2250
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "Importe Unitario :"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2610
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Total :"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   3030
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Codigo de Producto :"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   615
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "Descripcion :"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1005
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Unidad Medida :"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1815
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo Item (P/D) :"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   255
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Commision :"
         Height          =   255
         Left            =   5280
         TabIndex        =   13
         Top             =   3015
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Cantidad en Unidad Alternativa"
         Height          =   405
         Left            =   135
         TabIndex        =   12
         Top             =   3555
         Width           =   1515
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2760
      TabIndex        =   0
      Top             =   4320
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmAdicionaDetalleDocum.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmAdicionaDetalleDocum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CODIGO As String
Public Descripcion As String
Public strNum_Corre_Detalle As String, StrOption As String, IntSencuencia As Integer, strNum_Corre_Doc_Asig As String, _
        IntSencuencia_Doc_Asig As Integer
Dim StrSQL As String

Private Sub cmdBuscar_Click()
    'Load Frm_DetallaeItems
    'Frm_DetallaeItems.Show vbModal
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

On Error GoTo dprDepurar

Select Case ActionName

Case Is = "GRABAR"
  If MsgBox("Desea Grabar este Producto " & txtCod_Producto, vbYesNo, "AVISO") = vbYes Then
    Grabar
    Unload Me
  End If
Case Is = "CANCELAR"
  IntSencuencia = 0
  Unload Me
End Select

Exit Sub

dprDepurar:

errores err.Number

End Sub

Sub Grabar()

Dim RS As Object
Set RS = CreateObject("ADODB.Recordset")

StrSQL = "Ventas_Up_Man_Detalle '" & StrOption & "','" & strNum_Corre_Detalle & "'," & IntSencuencia & ",'" & txtTip_Item & "','" _
        & txtCod_Producto.Text & "','" & RTrim(Des_Apos(txtDescripcion.Text)) & "','" & txtUnida_Medida & "'," & TxtCantidad.Text & "," _
        & txtImp_Unitario.Text & "," & txtImp_Total.Text & "," & txtPorc_Commision.Text & ",'" & strNum_Corre_Doc_Asig & "'," & IntSencuencia_Doc_Asig & " ," & txtCantUniAlter.Text & ""

Set RS = CargarRecordSetDesconectado(StrSQL, cConnect)

'strSQL = "VENTAS_UP_MAN_DETALLE_ROLLO '" & strOption & "','" & strNum_Corre_Detalle & "'," & IntSencuencia & ",'" & txtTip_Item & "','" _
'        & txtCod_Producto.Text & "','" & Trim(txtCod_Producto.Text) & "','" & RTrim(Des_Apos(txtDescripcion.Text)) & "','" & txtUnida_Medida & "',0,0," & txtCantidad.Text & "," _
'        & txtImp_Unitario.Text & "," & txtImp_Total.Text & "," & txtPorc_Commision.Text & ",'" & strNum_Corre_Doc_Asig & "'," & IntSencuencia_Doc_Asig & " ," & txtCantUniAlter.Text & ",'','',''"

If Not (RS.EOF Or RS.BOF) Then
  IntSencuencia = RS!Secuencia
End If

End Sub

Private Sub txtCantidad_Change()
  txtImp_Total.Text = Format(TxtCantidad.Text * txtImp_Unitario.Text, "####.00")
  
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtCod_Producto_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtImp_Total_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtImp_Unitario_Change()
  txtImp_Total.Text = Format(TxtCantidad.Text * txtImp_Unitario.Text, "####.00")
End Sub

Private Sub txtImp_Unitario_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtPorc_Commision_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtTip_Item_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
  If KeyAscii = 13 Then
        If Trim(txtTip_Item.Text) = "P" Then
            cmdBuscar.Visible = True
        Else
            cmdBuscar.Visible = False
        End If
        
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtUnida_Medida_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
