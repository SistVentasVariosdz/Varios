VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmCCTejDetalleAddMasivo 
   Caption         =   "Adicionar Detalle Auditoria Tejeduria Rollos"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDatos 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.TextBox txtCod_Maquina 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1455
         TabIndex        =   5
         Top             =   240
         Width           =   1140
      End
      Begin VB.TextBox txtDes_Maquina 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2625
         TabIndex        =   4
         Top             =   240
         Width           =   3750
      End
      Begin VB.TextBox TxtCodigo_Rollo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1455
         TabIndex        =   3
         Top             =   600
         Width           =   1140
      End
      Begin VB.TextBox TxtSecuencia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1455
         TabIndex        =   2
         Top             =   960
         Width           =   660
      End
      Begin VB.CheckBox ChkContar 
         Caption         =   "Contar"
         Height          =   255
         Left            =   5520
         TabIndex        =   1
         Top             =   1080
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.Label LblUniMed 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   2520
         TabIndex        =   11
         Top             =   1995
         Width           =   45
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6720
         Y1              =   1395
         Y2              =   1395
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   1065
         Width           =   915
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   2280
      TabIndex        =   12
      Top             =   8400
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmCCTejDetalleAddMasivo.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin GridEX20.GridEX GridEX2 
      Height          =   6135
      Left            =   0
      TabIndex        =   13
      Top             =   2040
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   10821
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      HideSelection   =   2
      Options         =   8
      RecordsetType   =   1
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "FrmCCTejDetalleAddMasivo.frx":008A
      DataMode        =   1
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "FrmCCTejDetalleAddMasivo.frx":03A4
      Column(2)       =   "FrmCCTejDetalleAddMasivo.frx":046C
      FormatStylesCount=   6
      FormatStyle(1)  =   "FrmCCTejDetalleAddMasivo.frx":0510
      FormatStyle(2)  =   "FrmCCTejDetalleAddMasivo.frx":0648
      FormatStyle(3)  =   "FrmCCTejDetalleAddMasivo.frx":06F8
      FormatStyle(4)  =   "FrmCCTejDetalleAddMasivo.frx":07AC
      FormatStyle(5)  =   "FrmCCTejDetalleAddMasivo.frx":0884
      FormatStyle(6)  =   "FrmCCTejDetalleAddMasivo.frx":093C
      ImageCount      =   1
      ImagePicture(1) =   "FrmCCTejDetalleAddMasivo.frx":0A1C
      PrinterProperties=   "FrmCCTejDetalleAddMasivo.frx":0D36
   End
End
Attribute VB_Name = "FrmCCTejDetalleAddMasivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public sAccion As String
Public strSQL As String



Sub BUSCAR()
strSQL = "exec CC_AYUDA_MOTIVOS_NOTABLES_TEJEDURIA"
Set GridEX2.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
configura
End Sub


Sub configura()
GridEX2.Columns("codigo").Width = 1200
GridEX2.Columns("descripcion").Width = 3000
GridEX2.Columns("unimed").Width = 1200
GridEX2.Columns("cantidad").Width = 1200


End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "GRABAR"
    Call Grabar
Case "SALIR"
    Unload Me
End Select
End Sub

Private Sub GridEX2_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
If GridEX2.Columns("Cantidad").Index = ColIndex Then
    Cancel = False
Else
    Cancel = True
End If

End Sub

Sub Grabar()
On Error GoTo hand
Dim j As Integer

If GridEX2.RowCount = 0 Then Exit Sub

If GridEX2.RowCount > 0 Then

GridEX2.MoveFirst
For j = 1 To Me.GridEX2.RowCount

If GridEX2.Value(GridEX2.Columns("Cantidad").Index) > 0 Then

strSQL = "CC_MAN_AUDITORIA_TEJEDURIA_Detalle '" & sAccion & "','" & Trim(txtCod_Maquina.Text) & "','" & Trim(TxtCodigo_Rollo.Text) & "','" & Val(TxtSecuencia.Text) & "','" & Trim(GridEX2.Value(GridEX2.Columns("Codigo").Index)) & "'," & GridEX2.Value(GridEX2.Columns("Cantidad").Index) & ",'" & IIf(ChkContar, "S", "N") & "'"

 Call ExecuteSQL(cConnect, strSQL)
 
End If

            
GridEX2.MoveNext
Next
Unload Me
           
End If
 
Exit Sub
hand:
ErrorHandler err, "SALVAR_CABECERA"
End Sub


