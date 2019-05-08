VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Begin VB.Form FrmMante_CC_Tejeduria_Detalle 
   Caption         =   "Mantenimiento Detalle Auditoria Tejeduria Rollos"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDatos 
      Height          =   2415
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtCod_Maquina 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1455
         TabIndex        =   0
         Top             =   240
         Width           =   1140
      End
      Begin VB.TextBox txtDes_Maquina 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2625
         TabIndex        =   1
         Top             =   240
         Width           =   3750
      End
      Begin VB.TextBox TxtCodigo_Rollo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1455
         TabIndex        =   2
         Top             =   600
         Width           =   1140
      End
      Begin VB.TextBox TxtSecuencia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1455
         TabIndex        =   4
         Top             =   960
         Width           =   660
      End
      Begin VB.TextBox TxtCod_Motivo 
         Height          =   285
         Left            =   1455
         TabIndex        =   5
         Top             =   1560
         Width           =   780
      End
      Begin VB.TextBox TxtDes_Motivo 
         Height          =   285
         Left            =   2280
         TabIndex        =   6
         Top             =   1560
         Width           =   4095
      End
      Begin VB.TextBox TxtCantidad 
         Height          =   285
         Left            =   1455
         TabIndex        =   7
         Top             =   1920
         Width           =   780
      End
      Begin VB.CheckBox ChkContar 
         Caption         =   "Contar"
         Height          =   255
         Left            =   5400
         TabIndex        =   8
         Top             =   2025
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.Label LblUniMed 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   2520
         TabIndex        =   16
         Top             =   1995
         Width           =   45
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   6600
         Y1              =   1400
         Y2              =   1400
      End
      Begin VB.Label Label12 
         Caption         =   "Máquina"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   15
         Top             =   345
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo Rollo"
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
         Top             =   675
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "OT"
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
         Left            =   3480
         TabIndex        =   13
         Top             =   675
         Width           =   270
      End
      Begin VB.Label LblOT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3960
         TabIndex        =   3
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Secuencia"
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
         TabIndex        =   12
         Top             =   1065
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Motivo Defecto"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1665
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   2025
         Width           =   630
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2280
      TabIndex        =   17
      Top             =   2520
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmMante_CC_Tejeduria_Detalle.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "FrmMante_CC_Tejeduria_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Public CODIGO As String, Descripcion As String, TipoAdd As String

Public sAccion As String
Public oParent As Object

Sub BUSCAMOTIVO(tipo As Integer)
Dim oTipo As New frmBusqGeneral3
Dim rs As New ADODB.Recordset

Set oTipo.oParent = Me

'If Tipo = 1 Then
'    If Trim(TxtCod_Motivo.Text) = "" Then
'        oTipo.sQuery = "select cod_motivo as Codigo, Descripcion, Cod_UniMed As UniMed from CC_Confec_Motivos where cod_area_cc in ('hilan','tejed') order by descripcion"
'    Else
'        oTipo.sQuery = "select cod_motivo as Codigo, Descripcion, Cod_UniMed As UniMed from CC_Confec_Motivos where cod_motivo = '" & Trim(TxtCod_Motivo.Text) & "' and cod_area_cc in ('hilan','tejed')"
'    End If
'ElseIf Tipo = 2 Then
'    oTipo.sQuery = "select cod_motivo as Codigo, Descripcion, Cod_UniMed As UniMed  from CC_Confec_Motivos where descripcion  like '%" & Trim(TxtDes_Motivo.Text) & "%' and cod_area_cc in ('hilan','tejed') order by Descripcion"
'End If

If tipo = 1 Then
    If vemp = "07" Then
    oTipo.SQuery = "select cod_motivo as Codigo, A.Descripcion, cod_unimed as UniMed from CC_Confec_Motivos A, CC_AREAS B where cod_motivo like '%" & Trim(TxtCod_Motivo.Text) & "%' AND a.cod_area_cc = b.cod_area_cc and flg_tejeduria='*' AND B.Cod_Area_CC Not In ('COS','MAP') AND a.FLG_ESTADO='1'"
    Else
    oTipo.SQuery = "select cod_motivo as Codigo, A.Descripcion, cod_unimed as UniMed from CC_Confec_Motivos A, CC_AREAS B where cod_motivo like '%" & Trim(TxtCod_Motivo.Text) & "%' AND a.cod_area_cc = b.cod_area_cc and flg_tejeduria='*'"
    End If
ElseIf tipo = 2 Then
    If vemp = "07" Then
        oTipo.SQuery = "select cod_motivo as Codigo, A.Descripcion, cod_unimed as UniMed  from CC_Confec_Motivos A, CC_AREAS B where A.descripcion  like '%" & Trim(TxtDes_Motivo.Text) & "%' AND a.cod_area_cc = b.cod_area_cc and flg_tejeduria='*' AND B.Cod_Area_CC Not In ('COS','MAP') AND a.FLG_ESTADO='1' order by A.Descripcion"
    Else
        oTipo.SQuery = "select cod_motivo as Codigo, A.Descripcion, cod_unimed as UniMed  from CC_Confec_Motivos A, CC_AREAS B where A.descripcion  like '%" & Trim(TxtDes_Motivo.Text) & "%' AND a.cod_area_cc = b.cod_area_cc and flg_tejeduria='*' order by A.Descripcion"
    End If
End If

oTipo.Caption = "Buscar Motivo"
oTipo.Cargar_Datos

oTipo.gexLista.Columns("Codigo").Width = 1400
oTipo.gexLista.Columns("Descripcion").Width = 5000
oTipo.gexLista.Columns("UniMed").Width = 800

If oTipo.gexLista.RowCount > 1 Then
    oTipo.Show vbModal
Else
    CODIGO = oTipo.gexLista.Value(oTipo.gexLista.Columns("Codigo").Index)
    Descripcion = oTipo.gexLista.Value(oTipo.gexLista.Columns("Descripcion").Index)
    TipoAdd = oTipo.gexLista.Value(oTipo.gexLista.Columns("UniMed").Index)
End If

If Trim(CODIGO) <> "" Then
    TxtCod_Motivo.Text = CODIGO
    TxtDes_Motivo.Text = Descripcion
    LblUniMed = TipoAdd
    If UCase(LblUniMed) = "RL" Then
        TxtCantidad.Text = "1"
    End If
    CODIGO = "": Descripcion = "": TipoAdd = ""
    TxtCantidad.SetFocus
End If
Unload oTipo
Set oTipo = Nothing
Set rs = Nothing
End Sub

Private Sub ChkContar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    Call Grabar
Case "SALIR"
    Unload Me
End Select
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    Call SoloNumeros(TxtCantidad, KeyAscii, True, 2)
End If
End Sub

Private Sub TxtCantidad_LostFocus()
If Trim(TxtCantidad) = "" Then TxtCantidad.Text = "0"
End Sub

Private Sub TxtCod_Motivo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then BUSCAMOTIVO 1
End Sub

Private Sub TxtDes_Motivo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then BUSCAMOTIVO 2
End Sub

Sub Grabar()
On Error GoTo errGrabar

If Trim(TxtCantidad.Text) = "" Then TxtCantidad.Text = "0"

If UCase(Trim(LblUniMed)) = "RL" And CDbl(TxtCantidad.Text) <> 1 Then
    MsgBox "Cantidad para rollos debe ser igual a 1", vbCritical
    TxtCantidad.Text = "1"
    TxtCantidad.SetFocus
    Exit Sub
End If

strSQL = "CC_MAN_AUDITORIA_TEJEDURIA_Detalle '" & sAccion & "','" & Trim(txtCod_Maquina.Text) & "','" & Trim(TxtCodigo_Rollo.Text) & "','" & Val(TxtSecuencia.Text) & "','" & TxtCod_Motivo.Text & "'," & CDbl(TxtCantidad) & ",'" & IIf(ChkContar, "S", "N") & "'"
            
ExecuteCommandSQL cConnect, strSQL
Me.oParent.BUSCAR

If sAccion <> "I" Then
    Unload Me
Else
    TxtCod_Motivo.Text = ""
    TxtDes_Motivo.Text = ""
    TxtCantidad.Text = ""
    ChkContar.Value = Checked
    TxtCod_Motivo.SetFocus
End If

Exit Sub
errGrabar:
    ErrorHandler err, "Grabar"
End Sub
