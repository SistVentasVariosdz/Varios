VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmBusqGeneralOperario 
   Caption         =   "Busqueda"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   6870
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5460
      TabIndex        =   1
      Tag             =   "&Cancel"
      Top             =   3390
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4215
      TabIndex        =   0
      Tag             =   "&OK"
      Top             =   3390
      Width           =   1215
   End
   Begin GridEX20.GridEX DGridLista 
      Height          =   3300
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   5821
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigator =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      SelectionStyle  =   1
      AllowColumnDrag =   0   'False
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmBusqGeneralOperario.frx":0000
      FormatStyle(2)  =   "frmBusqGeneralOperario.frx":0138
      FormatStyle(3)  =   "frmBusqGeneralOperario.frx":01E8
      FormatStyle(4)  =   "frmBusqGeneralOperario.frx":029C
      FormatStyle(5)  =   "frmBusqGeneralOperario.frx":0374
      FormatStyle(6)  =   "frmBusqGeneralOperario.frx":042C
      FormatStyle(7)  =   "frmBusqGeneralOperario.frx":050C
      ImageCount      =   0
      PrinterProperties=   "frmBusqGeneralOperario.frx":052C
   End
End
Attribute VB_Name = "frmBusqGeneralOperario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public oParent As Object
Public sQuery As String, Campo As Integer
Dim Rs_Carga As New ADODB.Recordset

Sub Cargar_Datos()
On Error GoTo Cargar_DatosErr

Set DGridLista.ADORecordset = CargarRecordSetDesconectado(sQuery, cConnect)

If DGridLista.Columns.Count = 2 Then
    DGridLista.Columns(2).Width = 4000
End If

With oParent
    .codigo = ""
    .Descripcion = ""
    
'    If DGridLista.Columns.Count = 3 Then
'        .TipoAdd = ""
'    End If
    
End With
Exit Sub
Cargar_DatosErr:
    Set Rs_Carga = Nothing
    ErrorHandler Err, "Cargar_Datos"
End Sub

Private Sub DGridlista_DblClick()
    With oParent
        If DGridLista.RowCount > 0 Then DGridLista.ADORecordset.AbsolutePosition = DGridLista.RowIndex(DGridLista.Row)
        .codigo = DGridLista.Value(DGridLista.Columns(1).Index)
        
        If DGridLista.Columns.Count > 1 Then
            If IsNull(DGridLista.Value(DGridLista.Columns(2).Index)) Then
                .Descripcion = ""
            Else
                .Descripcion = DGridLista.Value(DGridLista.Columns(2).Index)
            End If
        End If
        
'        If DGridLista.Columns.Count = 3 Then
'            .TipoAdd = DGridLista.Value(DGridLista.Columns(3).Index)
'        End If
        
    End With

Unload Me
End Sub

Private Sub DGridlista_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        DGridlista_DblClick
    End If
End Sub

Private Sub DGridLista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DGridlista_DblClick
    Else
        If DGridLista.RowCount > 0 Then
            Dim rs As New ADODB.Recordset
            Set rs = DGridLista.ADORecordset
            If rs.RecordCount > 0 Then rs.MoveFirst
            Call BuscaCampo(rs, rs(IIf(DGridLista.Col = 0, 0, DGridLista.Col - 1)).Name, UCase(Chr(KeyAscii)))
            DGridLista.MoveToBookmark rs.Bookmark
            Set rs = Nothing
        End If
    End If

End Sub

Private Sub Form_Activate()
    DGridLista.SetFocus
End Sub

Private Sub Form_Load()
Call FormSet(Me)
   SetGeneralGridEX DGridLista, 0, 1
   Campo = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Rs_Carga = Nothing
End Sub
Public Sub cmdaceptar_Click()
    DGridlista_DblClick
End Sub
Private Sub cmdcancelar_Click()
oParent.codigo = ""
Unload Me
End Sub

