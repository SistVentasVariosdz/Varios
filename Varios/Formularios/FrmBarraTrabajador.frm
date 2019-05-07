VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmBarraTrabajador 
   Caption         =   "CODIGO DE BARRAS"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox CheckTodos 
      BackColor       =   &H00C0C0C0&
      Caption         =   "TODOS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   9840
      TabIndex        =   12
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&IMPRIME SELECIONADOS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7920
      Width           =   1245
   End
   Begin VB.CommandButton cmdImprimeGrande 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&IMPRIMIR ACTUAL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7920
      Width           =   1245
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&IMPRIMIR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7920
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10860
      Begin VB.TextBox Txt_Tipo 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Txt_Planilla 
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&BUSCAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9720
         Picture         =   "FrmBarraTrabajador.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1005
      End
      Begin VB.TextBox txtTrabjador 
         Height          =   285
         Left            =   4800
         TabIndex        =   2
         Top             =   240
         Width           =   4605
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "TIPO PLANILLA:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1290
      End
      Begin VB.Label Label3 
         Caption         =   "TRABAJADOR:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3600
         TabIndex        =   4
         Tag             =   "Document Type"
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&SALIR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7920
      Width           =   1005
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   6735
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   11880
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "FrmBarraTrabajador.frx":0102
      Column(2)       =   "FrmBarraTrabajador.frx":01CA
      FormatStylesCount=   6
      FormatStyle(1)  =   "FrmBarraTrabajador.frx":026E
      FormatStyle(2)  =   "FrmBarraTrabajador.frx":03A6
      FormatStyle(3)  =   "FrmBarraTrabajador.frx":0456
      FormatStyle(4)  =   "FrmBarraTrabajador.frx":050A
      FormatStyle(5)  =   "FrmBarraTrabajador.frx":05E2
      FormatStyle(6)  =   "FrmBarraTrabajador.frx":069A
      ImageCount      =   0
      PrinterProperties=   "FrmBarraTrabajador.frx":077A
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   2640
      Top             =   7440
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmBarraTrabajador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim strSQL As String
Public codigo As String
Public Descripcion As String
Public sAccion As String
Public TipoAdd As String
Public sPeriodo As Integer

Private Sub CheckTodos_Click()
     If GridEX1.RowCount = 0 Then Exit Sub
    
    Dim Rs As New ADODB.Recordset
    Dim valor As Boolean
    Dim i As Long

    If CheckTodos.Value = Checked Then
        valor = True
    Else
        valor = False
    End If

    GridEX1.Update
    Set Rs = GridEX1.ADORecordset
    Rs.MoveFirst
    Do While Not Rs.EOF
        Rs("SEL") = valor
        Rs.MoveNext
    Loop
   
    Rs.MoveFirst
    Rs.Update
    Set GridEX1.ADORecordset = Rs

End Sub
Private Sub cmdImprimeGrande_Click()
If GridEX1.RowCount = 0 Then Exit Sub
 Call Imprime_ZEBRA_grande(GridEX1.Value(GridEX1.Columns("trabajador").Index), GridEX1.Value(GridEX1.Columns("codigo").Index), GridEX1.Value(GridEX1.Columns("cargo").Index), GridEX1.Value(GridEX1.Columns("area").Index))

End Sub

Private Sub cmdImprimir_Click()
If GridEX1.RowCount = 0 Then Exit Sub
 Call Imprime_ZEBRA(GridEX1.Value(GridEX1.Columns("trabajador").Index), GridEX1.Value(GridEX1.Columns("codigo").Index), GridEX1.Value(GridEX1.Columns("cargo").Index), GridEX1.Value(GridEX1.Columns("area").Index))
 
End Sub

Private Sub imprimetodos()
If GridEX1.RowCount = 0 Then Exit Sub

Dim Rs As New ADODB.Recordset
Dim i As Integer
i = 1
Set Rs = Nothing
GridEX1.Refresh
GridEX1.Update

Set Rs = GridEX1.ADORecordset

Rs.MoveFirst
Do While i <= Rs.RecordCount

 If Rs!sel = True Then
    Call Imprime_ZEBRA_grande(Rs!trabajador, Rs!codigo, Rs!cargo, Rs!area)
 End If
 
Rs.MoveNext
i = i + 1
Loop


End Sub

Private Function Imprime_ZEBRA_grande(trabajador As String, dni As String, cargo As String, area As String) As Boolean
On Error GoTo errx
Dim sSQL  As String, SBARRA As String, sempresa As String
Dim mRS As ADODB.Recordset
Dim oPrint As clsPrintFile

sempresa = "TEXTILES JOC SRL"
Printer.Print " "
Printer.Print "^XA"
Printer.Print "^PRC"
Printer.Print "^LH0,0^FS"
Printer.Print "^LL1261"
Printer.Print "^MD0"
Printer.Print "^MNY"

SBARRA = RTrim(dni)

Printer.Print "^FO50,15^A0N,100,60^CI13^FR^FD" & RTrim(sempresa) & "^FS"
Printer.Print "^FO20,180^A0N,75,55^CI13^FR^FD" & RTrim(trabajador) & "^FS"


Printer.Print "^PQ1,0, 0, n"
Printer.Print "^XZ"
Printer.Print "^FX End of job"
Printer.Print "^XA"
Printer.Print "^IDR:ID*.*"
Printer.Print "^XZ"
Printer.EndDoc


Exit Function
errx:
    Close #1
    errores Err.numer
End Function

Private Function Imprime_ZEBRA(trabajador As String, dni As String, cargo As String, area As String) As Boolean
On Error GoTo errx
Dim sSQL  As String, SBARRA As String, sempresa As String
Dim mRS As ADODB.Recordset
Dim oPrint As clsPrintFile

sempresa = "TEXTILES JOC SRL"
Printer.Print " "
Printer.Print "^XA"
Printer.Print "^PRC"
Printer.Print "^LH0,0^FS"
Printer.Print "^LL1261"
Printer.Print "^MD0"
Printer.Print "^MNY"

SBARRA = RTrim(dni)

Printer.Print "^FO50,05^A0N,35,25^CI13^FR^FD" & RTrim(sempresa) & "^FS"
Printer.Print "^FO20,45^A0N,25,25^CI13^FR^FD" & RTrim(trabajador) & "^FS"
Printer.Print "^FO20,80^A0N,25,28^CI13^FR^FD" & RTrim(cargo) & "^FS"

Printer.Print "^BY3,3.0^FO50,120^BCN,100,N,N,N^FR^FD" & Trim(SBARRA) & "^FS"

Printer.Print "^FO20,240^A0N,25,28^CI13^FR^FD" & RTrim(area) & "^FS"

Printer.Print "^PQ1,0, 0, n"
Printer.Print "^XZ"
Printer.Print "^FX End of job"
Printer.Print "^XA"
Printer.Print "^IDR:ID*.*"
Printer.Print "^XZ"
Printer.EndDoc


Exit Function
errx:
    Close #1
    errores Err.numer
End Function


Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Call imprimetodos
End Sub

'Printer.Print "^FO10,05^A0N,50,40^CI13^FR^FD" & "Importado Por: "; RTrim(sempresa) & "^FS"
'Printer.Print "^FO10,45^A0N,40,40^CI13^FR^FD" & "Rif:J-294891390" & "^FS"
'Printer.Print "^FO10,80^A0N,40,25^CI13^FR^FD" & "Jersey Viscoza Full Ly. 30/1" & "^FS"
'Printer.Print "^FO10,120^A0N,40,25^CI13^FR^FD" & "95% Viscoza 5% Spandex" & "^FS"
'
'Printer.Print "^BY3,3.0^FO295,100^BCN,100,N,N,N^FR^FD" & Trim(SBARRA) & "^FS"
'
'
'Printer.Print "^FO10,160^A0N,40,25^CI13^FR^FD" & "Partida: "; RTrim(partida) & "^FS"
'Printer.Print "^FO10,205^A0N,40,25^CI13^FR^FD" & "Color: "; RTrim(Color) & "^FS"
'Printer.Print "^FO10,245^A0N,40,25^CI13^FR^FD" & "Peso: "; Format(Str(KIlos), "##.00") & "^FS"
'Printer.Print "^FO250,270^A0N,40,40^CI13^FR^FD" & "HECHO EN PERU " & "^FS"

Private Sub Txt_Tipo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        If Trim(Me.Txt_Tipo.Text) = "" Then
            Call Me.BUSCA_TIPTRABAJADOR(3)
            
        Else
            Call Me.BUSCA_TIPTRABAJADOR(1)
        End If
        
        
    End If
End Sub
Private Sub Txt_Planilla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Trim(Me.Txt_Planilla.Text) = "" Then
            Call Me.BUSCA_TIPTRABAJADOR(3)
        Else
            Call Me.BUSCA_TIPTRABAJADOR(1)
        End If
    End If
End Sub

Public Sub BUSCA_TIPTRABAJADOR(Tipo As Integer)
On Error GoTo hand
    Select Case Tipo
        Case 1:
                    strSQL = "SELECT Descripcion as 'Descripción' FROM  Rh_Tipo_Planilla WHERE Tip_Planilla = '" & Trim(Me.Txt_Tipo.Text) & "' "
                    Me.Txt_Planilla.Text = Trim(DevuelveCampo(strSQL, cConnect))
                    
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim Rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If Tipo = 2 Then
                        oTipo.sQuery = "SELECT Tip_Planilla as 'Código', Descripción as 'Descripción' FROM Rh_Tipo_Planilla WHERE Descripción LIKE '%" & Trim(Me.Txt_Planilla.Text) & "%'  ORDER BY Descripción"
                    Else
                        oTipo.sQuery = "SELECT Tip_Planilla as 'Código', Descripción AS 'Descripción' FROM Rh_Tipo_Planilla  ORDER BY Descripción"
                    End If
                    
                    oTipo.Cargar_Datos
                    oTipo.Show 1
                    If codigo <> "" Then
                        Me.Txt_Tipo = Trim(codigo)
                        Me.Txt_Planilla = Trim(Descripcion)
                        'Me.TxtTip_Trabajador = Trim(Codigo)
                        
                    End If
                    Set oTipo = Nothing
                    Set Rs = Nothing
                    
    End Select
    codigo = ""
    Descripcion = ""
    'Txt_CorrePlanilla.SetFocus
Exit Sub
hand:
ErrorHandler Err, "BUSCA TIPO PLANILLA"
End Sub

Private Sub cmdBuscar_Click()

    GridEX1.ClearFields
    Call mostrar
    
End Sub

Sub mostrar()
    Dim strSQL As String
    Dim sCodCentroCosto As String
    
    On Error GoTo Fin
   
   
    strSQL = "EXEC RH_MUESTRA_TRABAJADOR_BARRA '" & Txt_Planilla & _
                                             "','" & Trim(txtTrabjador.Text) & "'"
    cadena = strSQL
    
    Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    Dim C As Integer
        

    With GridEX1
        
        .Columns("Codigo").Width = 1000
        .Columns("TRABAJADOR").Width = 4000
        .Columns("INGRESO").Width = 1000
        .Columns("DNI").Width = 1000
        
        For C = 1 To .Columns.Count
            .Columns(C).HeaderAlignment = jgexAlignCenter
            .Columns(C).TextAlignment = jgexAlignLeft
        Next C
        
        
        If .RowCount > 0 Then
            .Row = -1
            .Col = .Columns.Count - 1
        End If
        

        .SetFocus
    End With
    Exit Sub
Fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub

