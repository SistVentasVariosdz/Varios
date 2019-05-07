VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7785
   Icon            =   "FrmCorrigeCorrelativoDoc.frx":0000
   ScaleHeight     =   3840
   ScaleWidth      =   7785
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   975
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   1455
      Left            =   1080
      TabIndex        =   1
      Top             =   1680
      Width           =   5055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub nombreEquipoRemoto()
'Dim tcpClient As Winsock
'Winsock1.RemoteHost = "pcleon"
'Winsock1.LocalIP = "192.168.1.7"
'Label2.Caption = Winsock1.LocalHostName
Dim cadena As String
Dim cadena2 As String
Dim cadena3 As String
Dim num_posicion As Integer

cadena = "BIXOLON SRP-270-PCLEON (desde PCLEON) en la sesión 2"

'Label2 = InStr(cadena, "(")
cadena2 = Mid(cadena, 1, CStr(InStr(cadena, "(")))

cadena2 = Trim(Mid(cadena, 1, CStr(InStr(cadena, "(")) - 1))
num_posicion = 2
cadena2 = Trim(Mid(cadena, 1, CStr(InStr(cadena, "(")) - 1))


'Label2.Caption = tcpClient.RemoteHost

End Sub

Private Sub Command2_Click()
Call nombreEquipoRemoto
End Sub
