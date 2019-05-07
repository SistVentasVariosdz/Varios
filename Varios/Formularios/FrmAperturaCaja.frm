VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Begin VB.Form frmAperturaCaja 
   Caption         =   "APERTURA CAJA"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "ACEPTAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "CANCELAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox txtFondoFijo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   840
      TabIndex        =   11
      Top             =   3840
      Width           =   3855
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "F O N D O   F I J O"
      Top             =   3480
      Width           =   5295
   End
   Begin VB.TextBox txtFechaApertura 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3000
      Width           =   3855
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "F E C H A   D E  A P E R T U R A"
      Top             =   2640
      Width           =   5295
   End
   Begin VB.TextBox txtcajaNro 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2040
      Width           =   5295
   End
   Begin VB.TextBox txtusuarioWindows 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox txtEstacion 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtTipoCambio 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox txtTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "A P E R T U R A  D E   D I A   V E N T A"
      Top             =   0
      Width           =   5295
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   240
      Top             =   4920
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   4320
      Width           =   5055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "USUARIO WINDOWS"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ESTACION"
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
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "TIPO CAMBIO DEL DIA"
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
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1740
   End
End
Attribute VB_Name = "frmAperturaCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CODIGO As String, Descripcion As String
Private strsql As String
Private cod_fabrica As String
Private cod_tienda As String
Private cod_caja As String
Private fecha_apertura  As Date
Private Sub cmdAceptar_Click()

    Call salvar_datos
    Call estadoCaja
End Sub
Private Sub salvar_datos()
On Error GoTo fin
Dim stit  As String
stit = "Advertencia"
Dim lon As Long

If validaApertura = True Then

    strsql = "CN_VENTAS_PERTURA_CAJA '" & cod_fabrica & _
    "', '" & cod_tienda & _
    "','" & cod_caja & "','" & Format(fecha_apertura, "dd/mm/yyyy") & _
    "'," & txtFondoFijo.Text & ",'" & ComputerName & _
    "','" & vusu & "','" & usuario_windows & "' "
    lon = ExecuteSQL(cConnect, strsql)
 
Call MsgBox("La Caja actual ha sido Aperturada con exito...! ", vbInformation + vbOKOnly, "Mensaje")

End If

Exit Sub
fin:
MsgBox Err.Description, vbCritical + vbOKOnly, stit
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call estadoCaja
End Sub
Private Sub estadoCaja()
Dim estado_caja As String
cmdAceptar.Visible = True

strsql = "EXEC SM_MUESTRA_ESTADO_CN_VENTAS_CAJAS_FECHA '" & ComputerName & "','" & usuario_windows & "'"
estado_caja = DevuelveCampo(strsql, cConnect)

Call datos_apertura

strsql = "EXEC SM_MUESTRA_ESTADO_CN_VENTAS_CAJAS_FECHA '" & ComputerName & "','" & usuario_windows & "'"
If DevuelveCampo(strsql, cConnect) = "P" Then
        
        Label3.Caption = "La caja no esta Apertura para la fecha actual...!"
        cmdAceptar.Enabled = True
Else
    If estado_caja = "A" Then
        Label3.Caption = "La Caja actual ya se Encuentra aperturada...!"
        cmdAceptar.Enabled = False
     
    ElseIf estado_caja = "C" Then
        Label3.Caption = "La Caja actual ya se Encuentra Cerrada...!"
        cmdAceptar.Enabled = False
      
    End If
End If

End Sub

Private Function validaApertura() As Boolean

validaApertura = True

If txtTipoCambio.Text = "" Or txtTipoCambio.Text = 0 Then
    Call MsgBox("¡¡¡ADVERTENCIA!!!" & Chr(13) & " No se puede Aperturar la caja...sirvase ingresar el tipo de cambio", vbInformation + vbOKOnly, "IMPORTANTE")
    validaApertura = False
    Exit Function
End If

If txtFondoFijo.Text = "" Or txtFondoFijo.Text = 0 Then
    Call MsgBox("¡¡¡ADVERTENCIA!!!" & Chr(13) & "No se puede Aperturar la caja...El Fondo Fijo no puede ser cero", vbInformation + vbOKOnly, "IMPORTANTE")
        validaApertura = False
    Exit Function
End If

'& Chr(13) & Chr(10)

End Function

Private Sub datos_apertura()
On Error GoTo fin

Dim rsdatos As New ADODB.Recordset
Dim stit As String
stit = "Advertencia"
strsql = "SM_MUESTRA_DATOS_APERTURA_CAJA '" & Now() & "','" & ComputerName & "','" & usuario_windows & "' "
Set rsdatos = CargarRecordSetDesconectado(strsql, cConnect)
'''SM_MUESTRA_ESTADO_CN_VENTAS_CAJAS_FECHA
With rsdatos

    txtTipoCambio.Text = rsdatos!TIPO_CAMBIO
    txtEstacion.Text = ComputerName
    txtcajaNro.Text = "Caja Nro " + !cod_caja
    txtFechaApertura.Text = !FECHA_APERTURA_TEXTO '!FECHA_APERTURA
    txtusuarioWindows.Text = usuario_windows
    txtFondoFijo.Text = Format(!IMP_APERTURA, "#######0.00")
    cod_fabrica = !cod_fabrica
    cod_tienda = !cod_tienda
    cod_caja = !cod_caja
    fecha_apertura = !fecha_apertura
    
End With

Exit Sub
fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, stit

End Sub
Private Sub txtFondoFijo_KeyPress(KeyAscii As Integer)

Call SoloNumeros(txtFondoFijo, KeyAscii, True, 4)

End Sub
