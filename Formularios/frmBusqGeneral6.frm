VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmBusqGeneral4 
   Caption         =   "Busqueda"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12975
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   12975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
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
      Left            =   11730
      TabIndex        =   1
      Tag             =   "&Cancel"
      Top             =   4455
      Width           =   1185
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
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
      Left            =   10470
      TabIndex        =   0
      Tag             =   "&OK"
      Top             =   4455
      Width           =   1185
   End
   Begin GridEX20.GridEX DGridLista 
      Height          =   4290
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   12870
      _ExtentX        =   22701
      _ExtentY        =   7567
      Version         =   "2.0"
      RecordNavigator =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmBusqGeneral6.frx":0000
      FormatStyle(2)  =   "frmBusqGeneral6.frx":0138
      FormatStyle(3)  =   "frmBusqGeneral6.frx":01E8
      FormatStyle(4)  =   "frmBusqGeneral6.frx":029C
      FormatStyle(5)  =   "frmBusqGeneral6.frx":0374
      FormatStyle(6)  =   "frmBusqGeneral6.frx":042C
      FormatStyle(7)  =   "frmBusqGeneral6.frx":050C
      ImageCount      =   0
      PrinterProperties=   "frmBusqGeneral6.frx":052C
   End
End
Attribute VB_Name = "frmBusqGeneral4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public oParent As Object
Public SQuery As String
Dim Rs_Carga As New ADODB.Recordset

Sub CARGAR_DATOS()
On Error GoTo Cargar_DatosErr

Set DGridLista.ADORecordset = CargarRecordSetDesconectado(SQuery, cConnect)

If DGridLista.Columns.Count = 2 Then
    DGridLista.Columns(2).Width = 4000
End If

With oParent
    .Codigo = ""
    .Descripcion = ""
    
    If DGridLista.Columns.Count = 3 Then
        .TipoAdd = ""
    End If
    
End With
Exit Sub
Cargar_DatosErr:
    Set Rs_Carga = Nothing
    ErrorHandler Err, "Cargar_Datos"
End Sub

Private Sub DGridlista_DblClick()
    With oParent
    
    If DGridLista.RowCount = 0 Then
        Unload Me
        Exit Sub
    End If
    DGridLista.ADORecordset.AbsolutePosition = DGridLista.Row
    .Codigo = DGridLista.Value(DGridLista.Columns(1).Index)
    
    If DGridLista.Columns.Count > 1 Then
        If IsNull(DGridLista.Value(DGridLista.Columns(2).Index)) Then
            .Descripcion = ""
        Else
            .Descripcion = DGridLista.Value(DGridLista.Columns(2).Index)
            .campo3 = DGridLista.Value(DGridLista.Columns(3).Index)
            .campo4 = DGridLista.Value(DGridLista.Columns(4).Index)
        End If
    End If
    
    If DGridLista.Columns.Count = 3 Then
        .TipoAdd = DGridLista.Value(DGridLista.Columns(3).Index)
    End If
        
    End With

Unload Me
End Sub

Private Sub DGridLista_KeyDown(KeyCode As Integer, Shift As Integer)
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
            Call BuscaCampo(rs, rs(0).Name, UCase(Chr(KeyAscii)))
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
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Rs_Carga = Nothing
End Sub
Private Sub cmdAceptar_Click()
    DGridlista_DblClick
End Sub
Private Sub CmdCancelar_Click()
oParent.Codigo = ""
Unload Me
End Sub



