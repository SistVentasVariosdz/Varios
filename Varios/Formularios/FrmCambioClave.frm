VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Begin VB.Form FrmCambioClave 
   Caption         =   "CAMBIO DE CONTRASEÑA"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnGuardar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&GUARDAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Width           =   1245
   End
   Begin VB.CommandButton btnCancelar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&SALIR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   1245
   End
   Begin VB.TextBox txtconfirmapwd 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1920
      Width           =   3495
   End
   Begin VB.TextBox txtNuevopwd 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1200
      Width           =   3495
   End
   Begin VB.TextBox txtClave 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   240
      Top             =   2160
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label Label4 
      Caption         =   "CONFIRMAR CONTRASEÑA:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "CONTRASEÑA NUEVA:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "CONTRASEÑA ACTUAL:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "FrmCambioClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSQL  As String
Dim clave As String
Private Sub btnCancelar_Click()
Unload Me
End Sub

Private Sub btnGuardar_Click()

If UCase(Trim(Me.txtNuevopwd.Text)) = "" Or UCase(Trim(Me.txtconfirmapwd)) = "" Then
Call Aviso("Debe ingresar la contraseña Actual", 2)
Else
    If clave = UCase(Trim(txtClave.Text)) Then
    guardaClave
    Else
    Call Aviso("La Contraseña Actual no es correcta", 2)
    txtClave.Text = ""
    End If
End If
End Sub

Private Sub Form_Load()

strSQL = " seg_muestra_clave_usuario '" & vusu & "'"
clave = Trim(DevuelveCampo(strSQL, cSEGURIDAD))
  txtClave.Text = ""
  txtNuevopwd.Text = ""
  txtconfirmapwd.Text = ""
  
'Me.txtClave.Text = clave
End Sub

Private Sub guardaClave()
  If UCase(Trim(Me.txtNuevopwd.Text)) <> UCase(Trim(Me.txtconfirmapwd)) Then
   
  Call Aviso("la Contraseña nueva no son iguales confirmar", 1)
  
  txtNuevopwd.Text = ""
  txtconfirmapwd.Text = ""
  
  Else
  
         'If msgbox( Aviso("Se va cambiar su Contraseña", 4) = True Then
         If MsgBox("Esta Seguro de Realizar en cambio de Contraseña", vbOKCancel, "Confirmacion ") = vbOK Then
            sSql = " seg_mant_Clave_Usuario '" & vusu & "','" & UCase(Trim(Me.txtNuevopwd.Text)) & "' "
            ExecuteSQL cSEGURIDAD, sSql
            Call Aviso("Se Relizo el Cambio Correctamente", 2)
         End If
        
        Unload Me
 
 End If
End Sub



