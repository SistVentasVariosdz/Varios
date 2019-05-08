VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Begin VB.Form FrmCambiosTelas 
   Caption         =   "Modificación Tela"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6915
   LinkTopic       =   "Form2"
   ScaleHeight     =   1965
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1275
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6840
      Begin VB.TextBox TxtAncho 
         Height          =   330
         Left            =   5670
         TabIndex        =   2
         Top             =   735
         Width           =   1065
      End
      Begin VB.TextBox TxtGramaje 
         Height          =   330
         Left            =   1785
         TabIndex        =   1
         Top             =   735
         Width           =   1065
      End
      Begin VB.TextBox TxtCod_Tela 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1785
         TabIndex        =   4
         Top             =   315
         Width           =   1065
      End
      Begin VB.TextBox TxtDes_Tela 
         Enabled         =   0   'False
         Height          =   330
         Left            =   2940
         TabIndex        =   3
         Top             =   315
         Width           =   3795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Acabados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4095
         TabIndex        =   7
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Gramaje Acabados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   105
         TabIndex        =   6
         Top             =   840
         Width           =   1605
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   105
         TabIndex        =   5
         Top             =   420
         Width           =   390
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2400
      TabIndex        =   8
      Top             =   1440
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmCambiosTelas.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "FrmCambiosTelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String

Public Cod_tiptela As String
Public Cod_Famtela As String
Public gramaje1_lav As Integer
Public Ancho1_Lav As Double
Public cod_uniMed As String
Public cod_uniMedCnf As String
Public Encog_Ancho As Double
Public Encog_Largo As Double
Public cod_gruTela As String
Public cod_galga As String
Public gramaje_crudo As Integer
Public Ancho_crudo As Double
Public Tip_Ancho As String
Public Cod_CtaCont    As String
Public cod_telaoriginal As String
Public cod_telaFinal As String
Public diametro As Integer
Public cod_tipRaya As String
Public num_alimentadores As Integer
Public num_aguja As Integer
Public num_rpm As Integer
Public cod_cliente As String
Public cod_temcli As String
Public num_lavadas As Integer
Public Encog_Ancho_Vap As Double
Public Encog_Largo_Vap As Double
Public Rapport As String
Public comentario As String
Public peso As Double
Public Grado_Doblez As String
Public inclinacion As String
Public flg_operatividad As String
Public gramaje_despueslavado As Integer
Public Opcion As Integer
Public sMts_Twill_x_Hora As Double
Public Cod_Tela_Desarrollo_Origen As String
Public Cod_Comb_Desarrollo_Origen As String
'


Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    If VALIDA = False Then Exit Sub
    Call CONFIRMA_CAMBIOS
Case "CANCELAR"
    Unload Me
End Select
End Sub

Sub CONFIRMA_CAMBIOS()
Dim nro_solicitud As Integer
On Error GoTo ErrSalvarDatos
strSQL = "DEVUELVE_NRO_SOLICITUD '" & txtcod_tela & "','" & vusu & "'"

nro_solicitud = DevuelveCampo(strSQL, cConnect)
strSQL = "TX_CAMBIOSTELA '1'," & Opcion & "," & nro_solicitud & ",'" & txtcod_tela & "','" & vusu & "','" & ComputerName & "','" & Cod_tiptela & "'," & txtGramaje & "," & _
                         TxtAncho & ",'" & Cod_Famtela & "','" & gramaje1_lav & "','" & Ancho1_Lav & "','" & cod_uniMed & "','" & _
                         cod_uniMedCnf & "','" & Encog_Ancho & "','" & Encog_Largo & "','" & cod_gruTela & "','" & cod_galga & "','" & _
                         gramaje_crudo & "','" & Ancho_crudo & "','" & Tip_Ancho & "','" & Cod_CtaCont & "','" & cod_telaoriginal & "','" & _
                         cod_telaFinal & "','" & diametro & "','" & cod_tipRaya & "','" & num_alimentadores & "','" & num_aguja & "','" & _
                         num_rpm & "','" & cod_cliente & "','" & cod_temcli & "','" & num_lavadas & "','" & Encog_Ancho_Vap & "','" & _
                         Encog_Largo_Vap & "','" & Rapport & "','" & comentario & "','" & peso & "',0,'" & vusu & "','S',1,'" & Grado_Doblez & "','" & inclinacion & "',0,0," & _
                         gramaje_despueslavado & ",0," & sMts_Twill_x_Hora & ",'" & Cod_Tela_Desarrollo_Origen & "','" & Cod_Comb_Desarrollo_Origen & "'"
ExecuteCommandSQL cConnect, strSQL
Unload Me
Exit Sub
ErrSalvarDatos:
    ErrorHandler err, "APRUEBA_SOLICITUD"
End Sub

Private Sub TxtAncho_GotFocus()
SelectionText TxtAncho
End Sub

Private Sub TxtAncho_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtGramaje_GotFocus()
SelectionText txtGramaje
End Sub

Function VALIDA() As Boolean
If Trim(txtGramaje.Text) = "" Then
    MsgBox "Ingrese Nuevo Gramaje"
    VALIDA = False
    Exit Function
End If
If Trim(TxtAncho.Text) = "" Then
    MsgBox "Ingrese Nuevo Ancho"
    VALIDA = False
    Exit Function
End If
VALIDA = True
End Function

Private Sub TxtGramaje_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
