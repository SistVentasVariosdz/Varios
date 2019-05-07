VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Begin VB.Form FrmCalculaMerma 
   Caption         =   "Calculos de Mermas"
   ClientHeight    =   3075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   9525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   10
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CALCULAR PORC. MERMA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   9
      Top             =   960
      Width           =   3015
   End
   Begin VB.CommandButton cmdCalcularPB 
      Caption         =   "CALCULAR PESO BRUTO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   8
      Top             =   360
      Width           =   3015
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "BORRAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   7
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmdCalcularPN 
      Caption         =   "CALCULAR PESO NETO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   6
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox txtPesoNeto 
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox txtMerma 
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtPesoBruto 
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   720
      Top             =   1920
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label Label3 
      Caption         =   "PESO NETO:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "PORCENTAJE MERMA %:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "PESO BRUTO:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "FrmCalculaMerma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalcularPB_Click()
   
   If txtPesoNeto.Text = "" Or txtMerma.Text = "" Then
        Call MsgBox("ingreso Valores Validos", vbInformation, "Mensaje")
        Exit Sub
   End If
    
    If (100# - txtMerma.Text) = 0 Then
        txtPesoBruto.Text = 0
    Else
        txtPesoBruto.Text = ((txtPesoNeto.Text) * 100#) / (100# - txtMerma.Text)
    End If

End Sub

Private Sub cmdCalcularPN_Click()

   If txtPesoBruto.Text = "" Or txtMerma.Text = "" Then
        Call MsgBox("ingreso Valores Validos", vbInformation, "Mensaje")
        Exit Sub
   End If
   

    txtPesoNeto.Text = (txtPesoBruto.Text) * ((100 - txtMerma.Text) / 100#)
End Sub

Private Sub CmdCancelar_Click()
    txtMerma.Text = 1
    txtPesoBruto.Text = 1
    txtPesoNeto.Text = 1
End Sub
Private Sub Command1_Click()

   If txtPesoBruto.Text = "" Or txtPesoNeto.Text = "" Then
        Call MsgBox("ingreso Valores Validos", vbInformation, "Mensaje")
        Exit Sub
   End If
   
   
If txtPesoBruto.Text = 0 Then
    txtMerma.Text = 0
Else
    txtMerma.Text = ((txtPesoBruto.Text - txtPesoNeto.Text) / txtPesoBruto.Text) * 100
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

    txtMerma.Text = 1
    txtPesoBruto.Text = 1
    txtPesoNeto.Text = 1

End Sub

Private Sub txtMerma_KeyPress(KeyAscii As Integer)
    Call SoloNumeros(txtMerma, KeyAscii, True, 4)
End Sub
Private Sub txtPesoBruto_KeyPress(KeyAscii As Integer)
    Call SoloNumeros(txtPesoBruto, KeyAscii, True, 4)
End Sub
Private Sub txtPesoNeto_KeyPress(KeyAscii As Integer)
    Call SoloNumeros(txtPesoNeto, KeyAscii, True, 4)
End Sub
