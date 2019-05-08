VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{144A86C7-1AF0-44BA-9AA8-AF3AAF6043B8}#1.0#0"; "NumBox.ocx"
Begin VB.Form frmImporteFobDua 
   Caption         =   "ImporteFobDua"
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   2310
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6015
      Begin VB.TextBox txtDua 
         Height          =   285
         Left            =   2400
         MaxLength       =   25
         TabIndex        =   2
         Top             =   240
         Width           =   3120
      End
      Begin NumBoxProject.NumBox txtFec_Numeracion 
         Height          =   285
         Left            =   2400
         TabIndex        =   3
         Top             =   600
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         TypeVal         =   3
         Mask            =   "99/99/9999"
         Formato         =   "dd/MM/yyyy"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   2
         Text            =   ""
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   0
      End
      Begin NumBoxProject.NumBox txtFec_Embarque 
         Height          =   285
         Left            =   4800
         TabIndex        =   4
         Top             =   600
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         TypeVal         =   3
         Mask            =   "99/99/9999"
         Formato         =   "dd/MM/yyyy"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   2
         Text            =   ""
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   0
      End
      Begin NumBoxProject.NumBox txtImp_FOB_Dol_Dua 
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Tag             =   "SET/VALID"
         Top             =   945
         Width           =   1095
         _ExtentX        =   1931
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
      Begin VB.Label Label30 
         Caption         =   "Importe FOB DUA $:"
         Height          =   255
         Left            =   255
         TabIndex        =   9
         Top             =   945
         Width           =   1455
      End
      Begin VB.Label Label23 
         Caption         =   "Dua :"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   255
         Width           =   495
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Fec. Numeracion :"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   615
         Width           =   1305
      End
      Begin VB.Label Label25 
         Caption         =   "Fec  Embarque:"
         Height          =   255
         Left            =   3600
         TabIndex        =   6
         Top             =   615
         Width           =   1215
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1920
      TabIndex        =   0
      Top             =   1680
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmImporteFobDua.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmImporteFobDua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrSQL As String
Public numCorre As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "ACEPTAR"
        SALVAR_DATOS
        
    Case "CANCELAR"
        Unload Me
End Select
End Sub


Sub SALVAR_DATOS()
    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr

    Screen.MousePointer = vbHourglass

    Con.ConnectionString = cConnect
    Con.CommandTimeout = 10000
    Con.Open

        Con.BeginTrans

        StrSQL = "EXEC Ventas_Up_Man_Datos_Dua '" & _
        numCorre & "','" & _
        txtDua & "','" & _
        txtFec_Numeracion.Text & "','" & _
        txtFec_Embarque.Text & "','" & _
        txtImp_FOB_Dol_Dua.Text & "'"

        ExecuteCommandSQL cConnect, StrSQL

        Con.CommitTrans

        Screen.MousePointer = vbDefault
        MsgBox "Los datos fueron procesados con éxito.", vbInformation, "Mensaje"
        Unload Me
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    Screen.MousePointer = vbDefault
    ErrorHandler err, "Salvar_Datos"
End Sub
 
'Private Sub LimpiaDatos()
'txtCodBarra.Text = ""
'txtCod_Modulo.Text = ""
'TxtCod_Dimension.Text = ""
'txtPesoBruto.Text = ""
'txtPesoNeto.Text = ""
'End Sub

Private Sub txtDua_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtFec_Numeracion.SetFocus
    End If
End Sub

Private Sub txtFec_Embarque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtImp_FOB_Dol_Dua.SetFocus
    End If
End Sub

Private Sub txtFec_Numeracion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtFec_Embarque.SetFocus
    End If
End Sub

Private Sub txtImp_FOB_Dol_Dua_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.FunctButt1.SetFocus
    End If
End Sub
