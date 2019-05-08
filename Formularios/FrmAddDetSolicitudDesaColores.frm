VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Begin VB.Form FrmAddDetSolicitudDesaColores 
   Caption         =   "Form1"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraDatos 
      Height          =   2775
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8055
      Begin VB.TextBox TxtCorr_Carta 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TxtDescripcion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   4455
      End
      Begin VB.TextBox TxtSec 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TxtDes_Color 
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Top             =   2280
         Width           =   5535
      End
      Begin VB.TextBox TxtDes_fibra 
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Top             =   1080
         Width           =   5535
      End
      Begin VB.TextBox TxtCod_ColCli 
         Height          =   285
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox TxtMat_Prima_Entregada 
         Height          =   405
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1440
         Width           =   5535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Corr. Carta"
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
         Left            =   240
         TabIndex        =   14
         Top             =   435
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sec."
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   330
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Des. Color"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Des. Fibra"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Codigo Color Cliente"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   2040
         Width           =   1425
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Mat. Prima Entregada"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   1530
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2760
      TabIndex        =   0
      Top             =   2880
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmAddDetSolicitudDesaColores.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "FrmAddDetSolicitudDesaColores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sAccion As String
Public CODIGO, Descripcion, TipoAdd As String
Public vOk As Boolean
Public oParent As Object
Dim StrSQL As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    Call Graba_Detalle
Case "CANCELAR"
    vOk = False
    Unload Me
End Select
End Sub

Sub Graba_Detalle()
On Error GoTo errGrabaDetalle

If Trim(TxtCod_ColCli.Text) = "." Or Trim(TxtCod_ColCli.Text) = "" Then
    MsgBox "Debe ingresar Codigo Color", vbInformation
    TxtCod_ColCli.SetFocus
    SelectionText TxtCod_ColCli
End If

StrSQL = "es_up_man_lb_cartacol_detalle '" & sAccion & "','" & TxtCorr_Carta.Text & "','" & TxtSec.Text & "','" & TxtDes_Color.Text & "','" & _
                 TxtDes_fibra.Text & "','" & TxtCod_ColCli.Text & "','" & vusu & "','" & ComputerName & "','" & Trim(TxtMat_Prima_Entregada.Text) & "','" & _
                 vusu & "','" & ComputerName & "'"
                 
ExecuteSQL cConnect, StrSQL
If sAccion = "I" Then
    oParent.vDes_Fibra = Trim(TxtDes_fibra.Text)
    oParent.vMatPrima = Trim(TxtMat_Prima_Entregada.Text)
End If

vOk = True
Unload Me
Exit Sub
errGrabaDetalle:
    ErrorHandler err, "Graba_Detalle"
End Sub

Private Sub TxtCod_ColCli_GotFocus()
SelectionText TxtCod_ColCli
End Sub

Private Sub TxtCod_ColCli_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    TxtDes_Color.SetFocus
End If
End Sub

Private Sub TxtDes_Color_GotFocus()
SelectionText TxtDes_Color
End Sub

Private Sub TxtDes_Color_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    FunctButt1.SetFocus
End If
End Sub

Private Sub TxtDes_fibra_GotFocus()
SelectionText TxtDes_fibra
End Sub

Private Sub TxtDes_fibra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    TxtMat_Prima_Entregada.SetFocus
End If
End Sub

Private Sub TxtMat_Prima_Entregada_GotFocus()
SelectionText TxtMat_Prima_Entregada
End Sub

Private Sub TxtMat_Prima_Entregada_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    TxtCod_ColCli.SetFocus
End If
End Sub



