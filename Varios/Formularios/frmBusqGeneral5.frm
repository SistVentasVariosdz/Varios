VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmBusqGeneral5 
   Caption         =   "Form1"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3900
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8145
      Begin GridEX20.GridEX gexLista 
         Height          =   3585
         Left            =   75
         TabIndex        =   0
         Top             =   210
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   6324
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MethodHoldFields=   -1  'True
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
         DataMode        =   1
         BackColorBkg    =   -2147483624
         ColumnHeaderHeight=   285
         FormatStylesCount=   7
         FormatStyle(1)  =   "frmBusqGeneral5.frx":0000
         FormatStyle(2)  =   "frmBusqGeneral5.frx":0138
         FormatStyle(3)  =   "frmBusqGeneral5.frx":01E8
         FormatStyle(4)  =   "frmBusqGeneral5.frx":029C
         FormatStyle(5)  =   "frmBusqGeneral5.frx":0374
         FormatStyle(6)  =   "frmBusqGeneral5.frx":042C
         FormatStyle(7)  =   "frmBusqGeneral5.frx":050C
         ImageCount      =   0
         PrinterProperties=   "frmBusqGeneral5.frx":052C
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
   End
End
Attribute VB_Name = "frmBusqGeneral5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SQuery As String, bCancel As Boolean
Public oParent As Object

Private Sub cmdAceptar_Click()
gexLista_DblClick
'Me.Hide
End Sub

Public Sub CARGAR_DATOS()
    bCancel = False
    Set gexLista.ADORecordset = CargarRecordSetDesconectado(SQuery, cConnect)
End Sub

Private Sub CmdCancelar_Click()
bCancel = True
Unload Me
End Sub

'Public Sub Form_Unload(Cancel As Integer)
'bCancel = True
'UnloadForm Me
'End Sub

Private Sub gexLista_DblClick()
With oParent
        'If gexLista.RowCount > 0 Then gexLista.ADORecordset.AbsolutePosition = gexLista.RowIndex(gexLista.Row)
        .Codigo = gexLista.Value(gexLista.Columns(1).Index)
        
        If gexLista.Columns.Count > 1 Then
            If IsNull(gexLista.Value(gexLista.Columns(2).Index)) Then
                .Descripcion = ""
            Else
                .Descripcion = gexLista.Value(gexLista.Columns(2).Index)
            End If
        End If
        
        If gexLista.Columns.Count >= 3 Then
            .TipoAdd = gexLista.Value(gexLista.Columns(3).Index)
        End If
        If gexLista.Columns.Count >= 4 Then
            .TipoAdd2 = gexLista.Value(gexLista.Columns(4).Index)
        End If
        If gexLista.Columns.Count >= 5 Then
            .TipoAdd3 = gexLista.Value(gexLista.Columns(5).Index)
        End If
    End With
If Me.Visible Then
    Me.Hide
End If
'Unload Me
End Sub

Private Sub gexLista_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        gexLista_DblClick
    End If
    If KeyCode = vbKeyEscape Then CmdCancelar_Click
End Sub
