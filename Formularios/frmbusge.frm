VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmBusqGeneral6 
   Caption         =   "Busqueda"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   9720
   StartUpPosition =   1  'CenterOwner
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
      Left            =   7200
      TabIndex        =   1
      Tag             =   "&OK"
      Top             =   3600
      Width           =   1215
   End
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
      Left            =   8400
      TabIndex        =   0
      Tag             =   "&Cancel"
      Top             =   3600
      Width           =   1215
   End
   Begin GridEX20.GridEX DGridLista 
      Height          =   3300
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   5821
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
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
      FormatStyle(1)  =   "frmbusge.frx":0000
      FormatStyle(2)  =   "frmbusge.frx":0138
      FormatStyle(3)  =   "frmbusge.frx":01E8
      FormatStyle(4)  =   "frmbusge.frx":029C
      FormatStyle(5)  =   "frmbusge.frx":0374
      FormatStyle(6)  =   "frmbusge.frx":042C
      FormatStyle(7)  =   "frmbusge.frx":050C
      ImageCount      =   0
      PrinterProperties=   "frmbusge.frx":052C
   End
End
Attribute VB_Name = "frmBusqGeneral6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public oParent As Object
Public SQuery As String
Dim Rs_Carga As New ADODB.Recordset
Public nomfor As String

Sub CARGAR_DATOS()
On Error GoTo Cargar_DatosErr

'Set DGridLista.ADORecordset = CargarRecordSetDesconectado(SQuery, cConnect)
Set DGridLista = CargarRecordSetDesconectado(SQuery, cConnect)

If DGridLista.Columns.Count = 2 Then
    DGridLista.Columns(2).Width = 4000
End If

With oParent
    .CODIGO = ""
    .Descripcion = ""
    
    If DGridLista.Columns.Count = 5 Then
        .TipoAdd = ""
    End If

    If DGridLista.Columns.Count = 4 Then
        .campo3 = ""
        .campo4 = ""
    End If

End With


Exit Sub
Cargar_DatosErr:
    Set Rs_Carga = Nothing
    ErrorHandler Err, "Cargar_Datos"
End Sub

Private Sub DGridlista_DblClick()
    With oParent
        If DGridLista.Row = 0 Then
            Exit Sub
        End If
        DGridLista.ADORecordset.AbsolutePosition = DGridLista.RowIndex(DGridLista.Row)
        .CODIGO = DGridLista.Value(DGridLista.Columns(1).Index)
        '.Descripcion = DGridLista.Value(DGridLista.Columns(2).Index)
        
        If DGridLista.Columns.Count > 1 Then
            If IsNull(DGridLista.Value(DGridLista.Columns(2).Index)) Then
                .Descripcion = ""
            Else
                .Descripcion = DGridLista.Value(DGridLista.Columns(2).Index)
            End If
        End If
        
        If DGridLista.Columns.Count = 3 Then
            .TipoAdd = DGridLista.Value(DGridLista.Columns(3).Index)
        End If
        
        If DGridLista.Columns.Count = 4 Then
            .campo3 = DGridLista.Value(DGridLista.Columns(3).Index)
            .campo4 = DGridLista.Value(DGridLista.Columns(4).Index)
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
            rs.MoveFirst
            Call BuscaCampo(rs, rs(0).Name, UCase(Chr(KeyAscii)))
            DGridLista.MoveToBookmark rs.Bookmark
    '        Call DGridLista.Find(1, jgexContains, UCase(Chr(KeyAscii)))
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
oParent.CODIGO = ""
Unload Me
End Sub

