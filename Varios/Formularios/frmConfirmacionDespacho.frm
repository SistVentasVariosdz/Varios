VERSION 5.00
Begin VB.Form frmConfirmacionDespacho 
   Caption         =   "Confirmación de despacho"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   LinkTopic       =   "frmConfirmacionDespacho"
   ScaleHeight     =   3120
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   3975
      Begin VB.CheckBox Check1 
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Despachar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmConfirmacionDespacho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public Cod_TipDoc As String
Public Serie As String
Public Nro_doc As String




Private Sub Command1_Click()
Dim cadena As String
Dim sql As String
If Check1.Value = 1 Then
cadena = "S"
'MsgBox (cadena)
Else
cadena = "N"
'MsgBox (cadena)
End If

   sql = " exec SP_ActualizarFlgDespaExten  '" & Cod_TipDoc & "', '" & Serie & "','" & Nro_doc & "','" & cadena & "' "
   Text1.Text = sql
   ExecuteSQL cCONNECT, sql
   
oParent.Buscar
End Sub
