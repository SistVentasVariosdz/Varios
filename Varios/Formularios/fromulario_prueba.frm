VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmImprimeEtiquetasPrendas 
   Caption         =   "IMPRESION DE ETIQUETAS DE PRENDAS"
   ClientHeight    =   8970
   ClientLeft      =   75
   ClientTop       =   375
   ClientWidth     =   12630
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   12630
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmcaja 
      Height          =   525
      Left            =   4200
      TabIndex        =   32
      Top             =   360
      Width           =   6015
      Begin VB.TextBox txt_numcaja 
         Height          =   285
         Left            =   1395
         TabIndex        =   33
         Top             =   165
         Width           =   4530
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Caja"
         Height          =   195
         Left            =   210
         TabIndex        =   34
         Top             =   210
         Width           =   315
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Elija"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   0
      TabIndex        =   4
      Tag             =   "Selection"
      Top             =   0
      Width           =   12615
      Begin VB.Frame frmPacking 
         Height          =   525
         Left            =   4200
         TabIndex        =   29
         Top             =   360
         Width           =   6015
         Begin VB.TextBox txt_numpacking 
            Height          =   285
            Left            =   1395
            TabIndex        =   30
            Top             =   165
            Width           =   4530
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Packing"
            Height          =   195
            Left            =   210
            TabIndex        =   31
            Top             =   210
            Width           =   585
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Caja"
         Height          =   255
         Index           =   5
         Left            =   9600
         TabIndex        =   28
         Top             =   120
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Packing"
         Height          =   255
         Index           =   4
         Left            =   8520
         TabIndex        =   27
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Op"
         Height          =   255
         Index           =   3
         Left            =   7800
         TabIndex        =   26
         Top             =   120
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Estilo Cliente"
         Height          =   255
         Index           =   2
         Left            =   6360
         TabIndex        =   25
         Top             =   120
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Po"
         Height          =   255
         Index           =   1
         Left            =   5520
         TabIndex        =   24
         Top             =   120
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Temporada"
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   23
         Top             =   120
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Frame fraEstCli 
         Height          =   525
         Left            =   4200
         TabIndex        =   19
         Top             =   360
         Width           =   6015
         Begin VB.TextBox txtCod_EstCli 
            Height          =   285
            Left            =   1395
            TabIndex        =   20
            Top             =   165
            Width           =   4530
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Estilo Cliente"
            Height          =   195
            Left            =   210
            TabIndex        =   21
            Top             =   210
            Width           =   900
         End
      End
      Begin VB.Frame fraPurOrd 
         Height          =   525
         Left            =   4200
         TabIndex        =   16
         Top             =   375
         Width           =   6015
         Begin VB.TextBox txtCod_PurOrd 
            Height          =   285
            Left            =   1380
            TabIndex        =   17
            Top             =   165
            Width           =   4530
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Purchase Order"
            Height          =   195
            Left            =   135
            TabIndex        =   18
            Top             =   225
            Width           =   1110
         End
      End
      Begin VB.Frame fraOP 
         Height          =   525
         Left            =   4200
         TabIndex        =   12
         Top             =   375
         Width           =   6015
         Begin VB.TextBox txtCod_Ordpro 
            Height          =   285
            Left            =   1350
            MaxLength       =   5
            TabIndex        =   14
            Top             =   180
            Width           =   795
         End
         Begin VB.TextBox txtDes_estpro 
            Height          =   285
            Left            =   2220
            TabIndex        =   13
            Top             =   180
            Width           =   3765
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "O/P"
            Height          =   195
            Left            =   150
            TabIndex        =   15
            Top             =   225
            Width           =   300
         End
      End
      Begin VB.Frame fraTemporada 
         Height          =   525
         Left            =   4200
         TabIndex        =   8
         Top             =   375
         Width           =   6015
         Begin VB.TextBox txtCod_TemCli 
            Height          =   285
            Left            =   1380
            MaxLength       =   3
            TabIndex        =   10
            Top             =   165
            Width           =   600
         End
         Begin VB.TextBox txtNom_TemCli 
            Height          =   285
            Left            =   2040
            TabIndex        =   9
            Top             =   165
            Width           =   3900
         End
         Begin VB.Label Label1 
            Caption         =   "Temporada"
            Height          =   180
            Left            =   150
            TabIndex        =   11
            Tag             =   "Season"
            Top             =   225
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "BUSCAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   10800
         TabIndex        =   7
         Tag             =   "Find"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtDesCliente 
         Height          =   285
         Left            =   1545
         TabIndex        =   6
         Top             =   360
         Width           =   2400
      End
      Begin VB.TextBox txtCodCliente 
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Top             =   360
         Width           =   690
      End
      Begin VB.Label lblCod_Cliente 
         Caption         =   "Cliente"
         Height          =   285
         Left            =   135
         TabIndex        =   22
         Tag             =   "Client"
         Top             =   390
         Width           =   765
      End
   End
   Begin VB.CheckBox CheckExpandir 
      Caption         =   "EXPANDIR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   11400
      TabIndex        =   2
      Top             =   1050
      Width           =   1185
   End
   Begin VB.CheckBox CheckTodos 
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
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   1050
      Width           =   1035
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   7185
      Left            =   0
      TabIndex        =   0
      Top             =   1290
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   12674
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GridLineStyle   =   2
      MethodHoldFields=   -1  'True
      Options         =   8
      RecordsetType   =   1
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "fromulario_prueba.frx":0000
      DataMode        =   1
      HeaderFontName  =   "Verdana"
      HeaderFontBold  =   -1  'True
      HeaderFontSize  =   6.75
      HeaderFontWeight=   700
      FontName        =   "Tahoma"
      ColumnHeaderHeight=   270
      ColumnsCount    =   2
      Column(1)       =   "fromulario_prueba.frx":031A
      Column(2)       =   "fromulario_prueba.frx":03E2
      FormatStylesCount=   8
      FormatStyle(1)  =   "fromulario_prueba.frx":0486
      FormatStyle(2)  =   "fromulario_prueba.frx":05AE
      FormatStyle(3)  =   "fromulario_prueba.frx":065E
      FormatStyle(4)  =   "fromulario_prueba.frx":0712
      FormatStyle(5)  =   "fromulario_prueba.frx":07EA
      FormatStyle(6)  =   "fromulario_prueba.frx":08A2
      FormatStyle(7)  =   "fromulario_prueba.frx":0982
      FormatStyle(8)  =   "fromulario_prueba.frx":0A7E
      ImageCount      =   1
      ImagePicture(1) =   "fromulario_prueba.frx":0B7E
      PrinterProperties=   "fromulario_prueba.frx":0E98
   End
   Begin VB.CommandButton cmdImprimetickets 
      Caption         =   "IMPRIMIR SELECCIONDAS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   3
      Top             =   8550
      Width           =   2925
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   2310
      Top             =   9960
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmImprimeEtiquetasPrendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private indice  As Integer
Private sqlstr  As String
Private sCod_Cliente As String
Public codigo As String
Public Descripcion As String

Private Sub CheckExpandir_Click()
    If GridEX1.RowCount = 0 Then Exit Sub
    With GridEX1
        Select Case CBool(CheckExpandir.Value)
            Case True: .ExpandAll
            Case False: .CollapseAll
        End Select
        If .RowCount > 0 Then
            .Row = -1
            .Col = .Columns.Count - 1
        End If
    End With
    
End Sub

Private Sub CheckTodos_Click()

    Dim valor As String
    Dim I As Long
    Dim rs As ADODB.Recordset
    If CheckTodos.Value = Checked Then
        valor = "1"
    Else
        valor = "0"
    End If
    GridEX1.MoveFirst
    Set rs = GridEX1.ADORecordset
    rs.MoveFirst
    Do While Not rs.EOF
        rs("sel") = valor
        rs.MoveNext
    Loop

    rs.MoveFirst
    Set GridEX1.ADORecordset = rs
    CONFIGURA_GRILLA
    

End Sub

Private Sub cmdBuscar_Click()
On Error GoTo SALTO_ERROR

Call muestraGrilla
Call CONFIGURA_GRILLA

    Exit Sub
SALTO_ERROR:
    MsgBox Err.Description, vbCritical, Me.Caption
    
End Sub
Private Sub muestraGrilla()

sqlstr = "EXEC CF_DETALLECAJA_IMPRESION_ETIQUETAS  '" & Trim(txtCodCliente.Text) & "','" & Trim(txtCod_PurOrd.Text) & "','" & Trim(txtCod_TemCli.Text) & "','" & Trim(txtCod_EstCli.Text) & "','" & Trim(txtCod_Ordpro.Text) & "','" & Trim(txt_numpacking.Text) & "','" & Trim(txt_numcaja.Text) & "','" & indice & "'"
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(sqlstr, cConnect)

End Sub

Private Sub cmdImprimetickets_Click()
Call IMPRIMIRZEBRA
End Sub
Private Sub IMPRIMIRZEBRA()
On Error GoTo SALTO_ERROR
Dim cantImprimir As Integer
Dim I As Integer
Dim sempresa  As String

sempresa = DevuelveCampo("SELECT Des_Empresa FROM seg_empresas WHERE Cod_Empresa='" & vemp & "'", cSEGURIDAD)

',CLIENTE1=   'CLIENTE : '+UPPER(C.ABR_CLIENTE)
',COD_PURORD1='PO      : '+UPPER(A.COD_PURORD)
',COD_ESTCLI1='ESTILO  : '+UPPER(A.COD_ESTCLI)
',COD_COLCLI1='COLOR   : '+UPPER(A.COD_COLCLI)
',COD_TALLA1= 'TALLA   : '+UPPER(LEFT(LTRIM(RTRIM(A.COD_TALLA))+'          ',10))
',DES_PRESENT1='PRESENT : '+UPPER(DES_PRESENT)


           Dim oRs As New ADODB.Recordset
           GridEX1.Update
           Set oRs = GridEX1.ADORecordset
        
           If GridEX1.RowCount > 0 Then
                oRs.MoveFirst
                Do While Not oRs.EOF
                    If CBool(oRs.Fields("SEL").Value) = True Then
                     
                     cantImprimir = oRs.Fields("NROIMPRIMIR").Value
                     I = 1
                     If cantImprimir > 0 Then
                        Do While I <= cantImprimir
                         
                            Call IMPRIMIRZPL(oRs.Fields("cliente1").Value, oRs.Fields("COD_PURORD1").Value, _
                            oRs.Fields("COD_ESTCLI1").Value, oRs.Fields("COD_COLCLI1").Value, _
                            oRs.Fields("COD_ORDPRO").Value, oRs.Fields("COD_ORDPRO1").Value, oRs.Fields("COD_TALLA").Value, oRs.Fields("COD_TALLA1").Value, _
                            oRs.Fields("COD_calidad").Value, oRs.Fields("COD_PRESENT").Value, oRs.Fields("DES_PRESENT1").Value, sempresa)
                           
                            I = I + 1
                        Loop
                        
                     End If
                     
                    End If
                    oRs.MoveNext
                Loop
           End If
    Call MsgBox("TERMINO DE IMPRIMIR", vbInformation, "Mensaje")
    Exit Sub
SALTO_ERROR:
    MsgBox Err.Description, vbCritical, Me.Caption
    
End Sub
Private Function IMPRIMIRZPL(cliente As String, poo As String, estilo As String, color As String, _
oop As String, oop1 As String, talla As String, TALLA1 As String, calidad As String, presentacion As String, DES_PRESET As String, sempresa As String) As Boolean

    On Error GoTo errx

    Dim sSQL  As String, SBARRA As String
    Dim oPrint As clsPrintFile
    Dim sRollo As String
    Dim sTela As String
    SBARRA = Trim(oop) + Trim(presentacion) + Trim(calidad) + talla

        Printer.Print " "
        Printer.Print "^XA"
        Printer.Print "^PRC"
        Printer.Print "^LH0,0^FS"
        Printer.Print "^LL1000"
        Printer.Print "^MD0"
        Printer.Print "^MNY"
        
'        Printer.Print "^FO25,35"
'        Printer.Print "^A0,18,25"
'        Printer.Print "^FDCLIENTE: " & Trim(cliente) & "^FS"
'
'        Printer.Print "^FO25,50"
'        Printer.Print "^A0,18,25"
'        Printer.Print "^FDPO        : " & Trim(poo) & "^FS"
'
'        Printer.Print "^FO25,65"
'        Printer.Print "^A0,18,25"
'        Printer.Print "^FDESTILO  : " & Trim(estilo) & "^FS"
'
'        Printer.Print "^FO25,80"
'        Printer.Print "^A0,18,25"
'        Printer.Print "^FDOP         : " & Trim(oop) & "^FS"
'
'        Printer.Print "^FO25,95"
'        Printer.Print "^A0,18,25"
'        Printer.Print "^FDPRESENT: " & Trim(DES_PRESET) & "^FS"
'
'        Printer.Print "^FO25,110"
'        Printer.Print "^A0,18,25"
'        Printer.Print "^FDTALLA     : " & Trim(talla) & "^FS"
'
'        Printer.Print "^FO120,125,^BY1"
'        Printer.Print "^BCN,80,Y,N,N^FR^FD" & RTrim(Trim(SBARRA)) & "^FS"
''********************
        Printer.Print "^FO25,35"
        Printer.Print "^A0,18,25"
        Printer.Print "^FD" & Trim(cliente) & "^FS"

        Printer.Print "^FO25,50"
        Printer.Print "^A0,18,25"
        Printer.Print "^FD" & Trim(poo) & "^FS"

        Printer.Print "^FO25,65"
        Printer.Print "^A0,18,25"
        Printer.Print "^FD" & Trim(estilo) & "^FS"

        Printer.Print "^FO25,80"
        Printer.Print "^A0,18,25"
        Printer.Print "^FD" & Trim(oop1) & "^FS"

        Printer.Print "^FO25,95"
        Printer.Print "^A0,18,25"
        Printer.Print "^FD" & Trim(DES_PRESET) & "^FS"

        Printer.Print "^FO25,110"
        Printer.Print "^A0,18,25"
        Printer.Print "^FD" & Trim(TALLA1) & "^FS"

        Printer.Print "^FO120,125,^BY1"
        Printer.Print "^BCN,80,Y,N,N^FR^FD" & RTrim(Trim(SBARRA)) & "^FS"


        Printer.Print "^XZ"
        Printer.Print "^FX End of job"
        Printer.Print "^XA"
        Printer.Print "^IDR:ID*.*"
        Printer.Print "^XZ"
        Printer.EndDoc
        
        
    Exit Function

errx:
    Close #1
    Errores Err.numer
End Function
'Private Function Imprime_ZEBRA_original(ByVal sCod_OrdTra As String, ByVal iNum_SecOrdTra As String, ByVal sRolloInicio As String, ByVal sRolloFin As String) As Boolean
'    On Error GoTo errx
'
'    Dim ssQl  As String, SBARRA As String, sempresa As String
'    Dim mRS As ADODB.Recordset
'    Dim oPrint As clsPrintFile
'
'    ssQl = "EXEC SM_MUESTRA_DATA_ROLLOS_A_IMPRIMIR '$',$,'$','$','$'"
'    ssQl = VBsprintf(ssQl, sCod_OrdTra, iNum_SecOrdTra, iSec_Maquina, sRolloInicio, sRolloFin)
'
'    Set mRS = GetDataSet(cConnect, ssQl)
'    sempresa = DevuelveCampo("SELECT Des_Empresa FROM seg_empresas WHERE Cod_Empresa='" & vemp & "'", cSEGURIDAD)
'
'    Dim sRollo As String
'    Dim sTela As String
'
'    Do While Not mRS.EOF
'        Printer.Print " "
'        Printer.Print "^XA"
'        Printer.Print "^PRC"
'        Printer.Print "^LH0,0^FS"
'        Printer.Print "^LL1261"
'        Printer.Print "^MD0"
'        Printer.Print "^MNY"
'
'        'SBARRA = Mid(mRS("CODIGO_ANTIGUO1").Value, 1, 2) & Mid(mRS("CODIGO_ANTIGUO1").Value, 3, 2) & Mid(mRS("CODIGO_ANTIGUO1").Value, 5, 4) & _
'        Mid(mRS("CODIGO_ANTIGUO2").Value, 1, 2) & Mid(mRS("CODIGO_ANTIGUO2").Value, 3, 2) & Trim(Mid(mRS("CODIGO_ANTIGUO2").Value, 5, 4)) & _
'        Trim(Mid(mRS("COD_FAMGRUPO").Value, 1, 4)) & Format(mRS("Long_Malla1").Value, "0.000") & Format(mRS("Long_Malla2").Value, "0.000") & _
'        Mid(mRS("Codigo_Rollo").Value, 1, 5) & Trim(Mid(mRS("Prefijo_Maquina").Value, 1, 2))
'
'        sempresa = LTrim(sempresa)
'        sRollo = Trim(mRS("Prefijo_Maquina").Value) & "-" & Trim(mRS("Codigo_Rollo").Value)
'        sTela = Trim(mRS("COD_TELA")) + ": " + UCase(Left(Trim(mRS("Des_tela").Value), 20))
'
'        'ECN: 01/08/2011 -
'        'SBARRA = Trim(Mid(mRS("Prefijo_Maquina").Value, 1, 2)) & Trim(Mid(mRS("Codigo_Rollo").Value, 1, 5)) + "00000000000000"
'        SBARRA = Trim(Mid(mRS("Prefijo_Maquina").Value, 1, 2)) & Trim(Mid(mRS("Codigo_Rollo").Value, 1, 5))
'
'
'        Printer.Print "^FO15,20^A0N,23,25^CI13^FR^FDROLLO : " & sRollo & "^FS"
'        Printer.Print "^FO15,50^A0N,23,25^CI13^FR^FDOT       : " & RTrim(sCod_OrdTra) & "^FS"
'        Printer.Print "^FO15,80^A0N,23,25^CI13^FR^FDGRUPO: " & Trim(mRS("Grupo")) & "^FS"
'
'        Printer.Print "^FO530,20^A0N,23,25^CI13^FR^FD" & Trim(sempresa) & "^FS"
'        Printer.Print "^FO680,50^A0N,23,25^CI13^FR^FD" & Trim(FixNulos(mRS("FEC_GENERACIONOT").Value, vbString)) & "^FS"
'
'        Printer.Print "^FO550,220^A0N,18,25^CI13^FR^FDLONG.MALLA 1: " & Format(mRS("Long_Malla1").Value, "0.000") & "^FS"
'        Printer.Print "^FO550,240^A0N,18,25^CI13^FR^FDLONG.MALLA 2: " & Format(mRS("Long_Malla2").Value, "0.000") & "^FS"
'
'        Printer.Print "^BY1,3.0^FO380,20^BCN,220,Y,N,Y^FR^FD" & RTrim(Trim(SBARRA)) & "^FS"
'
'
'        Printer.Print "^FO15,240^A0N,18,23^CI13^FR^FD" & Trim(sTela) & "^FS"
'
'        Printer.Print "^PQ1,0, 0, n"
'        Printer.Print "^XZ"
'        Printer.Print "^FX End of job"
'        Printer.Print "^XA"
'        Printer.Print "^IDR:ID*.*"
'        Printer.Print "^XZ"
'        Printer.EndDoc
'
'        mRS.MoveNext
'    Loop
'
'    mRS.Close
'    Set mRS = Nothing
'    Exit Function
'
'errx:
'    Close #1
'    Errores Err.numer
'End Function

Private Sub Form_Load()

    indice = 1
    Me.fraEstCli.Visible = False
    Me.fraOP.Visible = False
    Me.fraPurOrd.Visible = False
    Me.fraTemporada.Visible = True
    frmcaja.Visible = False
    frmPacking.Visible = False
    'Me.txtCod_TemCli.SetFocus

End Sub

Private Sub CONFIGURA_GRILLA()
    On Error GoTo SALTO_ERROR
    Dim C As Integer
            
    With GridEX1
        For C = 1 To .Columns.Count
            .Columns(C).Visible = False
            .Columns(C).HeaderAlignment = jgexAlignCenter
            .Columns(C).TextAlignment = jgexAlignCenter
        Next C
        
        With .Columns("NUM_CAJA")
            .Visible = False
            .Caption = "N° CAJA"
            .Width = 900
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("CLIENTE")
            .Visible = True
            .Caption = "CLIENTE"
            .Width = 1000
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("COD_PURORD")
            .Visible = True
            .Caption = "PO"
            .Width = 1500
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("COD_ESTCLI")
            .Visible = True
            .Caption = "ESTILO"
            .Width = 1500
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("COD_COLCLI")
            .Visible = True
            .Caption = "COLOR"
            .Width = 1000
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("COD_TALLA")
            .Visible = True
            .Caption = "TALLA"
            .Width = 800
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("COD_ORDPRO")
            .Visible = True
            .Caption = "OP"
            .Width = 800
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("DES_PRESENT")
            .Visible = True
            .Caption = "PRESENTACION"
            .Width = 1500
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("COD_CALIDAD")
            .Visible = True
            .Caption = "CALIDAD"
            .Width = 800
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("NUM_PRENDAS")
            .Visible = True
            .Caption = "Q PRENDAS"
            .Width = 1200
            .TextAlignment = jgexAlignRight
        End With
        
        With .Columns("NROIMPRIMIR")
            .Visible = True
            .Width = 1200
            .Caption = "Q IMPRIMIR"
            .TextAlignment = jgexAlignRight
        End With
        With .Columns("sel")
            .Visible = True
            .Width = 500
            .Caption = "SEL"
            .TextAlignment = jgexAlignLeft
        End With
        
        Dim oGroup01 As GridEX20.JSGroup
        Dim oGroup02 As GridEX20.JSGroup
        
        Dim colnumPRENDAS As JSColumn
        Dim colNroImprimirPRENDAS As JSColumn
    
        With GridEX1
            Set oGroup01 = .Groups.Add(.Columns("NUM_CAJA").Index, jgexSortAscending)
            'Set oGroup02 = .Groups.Add(.Columns("NUM_PACKING").Index, jgexSortAscending)
             
            .GroupFooterStyle = jgexTotalsGroupFooter
            
            Set colnumPRENDAS = .Columns("NUM_PRENDAS")
            With colnumPRENDAS
                .AggregateFunction = jgexSum
                .TotalRowPrefix = "T. PREND: "
            End With
            
            Set colNroImprimirPRENDAS = .Columns("NROIMPRIMIR")
            With colNroImprimirPRENDAS
                .AggregateFunction = jgexSum
                .TotalRowPrefix = "T. IMPRI: "
            End With
            If CheckExpandir.Value = Checked Then
                .DefaultGroupMode = jgexDGMExpanded
            Else
                .DefaultGroupMode = jgexDGMCollapsed
            End If
                
        End With
        
        
    End With
    Call setcolorcolumnas
    
    Exit Sub
    
SALTO_ERROR:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub
Private Sub setcolorcolumnas()

    GridEX1.Columns("NUM_PRENDAS").CellStyle = "AZUL"
    GridEX1.Columns("NROIMPRIMIR").CellStyle = "VERDE"
        
End Sub

Private Sub Option1_Click(Index As Integer)
    indice = Index + 1
    txt_numcaja.Text = ""
    txt_numpacking.Text = ""
    txtCod_EstCli.Text = ""
    txtNom_TemCli.Text = ""
    txtCod_TemCli.Text = ""
    txtCod_Ordpro.Text = ""
    txtCod_PurOrd.Text = ""
    txtDes_estpro.Text = ""
    Set GridEX1.ADORecordset = Nothing
    
   Select Case indice
   Case 1
       
        Me.fraEstCli.Visible = False
        Me.fraOP.Visible = False
        Me.fraPurOrd.Visible = False
        Me.fraTemporada.Visible = True
        frmcaja.Visible = False
        frmPacking.Visible = False
        Me.txtCod_TemCli.SetFocus
    
   Case 2
        Me.fraEstCli.Visible = False
        Me.fraOP.Visible = False
        Me.fraPurOrd.Visible = True
        Me.fraTemporada.Visible = False
        frmcaja.Visible = False
        frmPacking.Visible = False
        Me.txtCod_PurOrd.SetFocus
  
   Case 3
        Me.fraEstCli.Visible = True
        Me.fraOP.Visible = False
        Me.fraPurOrd.Visible = False
        Me.fraTemporada.Visible = False
        frmcaja.Visible = False
        frmPacking.Visible = False
        Me.txtCod_EstCli.SetFocus

   Case 4
        Me.fraEstCli.Visible = False
        Me.fraOP.Visible = True
        Me.fraPurOrd.Visible = False
        Me.fraTemporada.Visible = False
        frmcaja.Visible = False
        frmPacking.Visible = False
        txtCod_Ordpro.SetFocus
   Case 5
        Me.fraEstCli.Visible = False
        Me.fraOP.Visible = False
        Me.fraPurOrd.Visible = False
        Me.fraTemporada.Visible = False
        frmcaja.Visible = False
        frmPacking.Visible = True
        txt_numpacking.SetFocus
    
   Case 6
        Me.fraEstCli.Visible = False
        Me.fraOP.Visible = False
        Me.fraPurOrd.Visible = False
        Me.fraTemporada.Visible = False
        frmcaja.Visible = True
        frmPacking.Visible = False
        'txt_numpacking.SetFocus
   End Select
End Sub
Private Sub txt_numpacking_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdBuscar.SetFocus
    Else
        Call SoloNumeros(txt_numpacking, KeyAscii, False)
    End If
End Sub
Private Sub txt_numCAJA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdBuscar.SetFocus
        KeyAscii = 0
    Else
        Call SoloNumeros(txt_numcaja, KeyAscii, False)
    End If
End Sub

Private Sub txtCod_TemCli_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Trim(txtCodCliente.Text) <> "" Then
     Call Busca_Opcion("Cod_TemCli", "Nom_TemCli", "TG_TemCli where Cod_Cliente='" & Trim(txtCodCliente.Text) & "' and ", txtCod_TemCli, txtNom_TemCli, 1, Me)
   Else
     Call MsgBox("Ingrese El cliente", vbInformation, "Mensaje")
   End If
   
End If
End Sub
Private Sub txtNom_TemCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call Busca_Opcion("Cod_TemCli", "Nom_TemCli", "TG_TemCli where Cod_Cliente='" & Trim(txtCodCliente.Text) & "' and ", txtCod_TemCli, txtNom_TemCli, 2, Me)
    End If
End Sub
Private Sub txtCodCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Busca_Opcion("Cod_Cliente", "Nom_Cliente", "tg_cliente where  ", txtCodCliente, txtDesCliente, 1, Me)
End If
End Sub
Private Sub txtDesCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call Busca_Opcion("Cod_Cliente", "Nom_Cliente", "tg_cliente where  ", txtCodCliente, txtDesCliente, 1, Me)
End Sub

Public Sub Busca_Opcion(strCampo1 As String, strCampo2 As String, strTabla As String, txtCod As TextBox, txtDes As TextBox, Opcion As Integer, frmME As Form)
On Error GoTo Fin

Dim rstAux As ADODB.Recordset, strSQL As String

    strSQL = "Select " & strCampo1 & " AS Cod," & strCampo2 & " as Descripcion from " & strTabla

    txtCod = Trim(txtCod)
    txtDes = Trim(txtDes)
    
    
    Select Case Opcion
    Case 1: strSQL = strSQL & strCampo1 & " like '%" & txtCod & "%'"
    Case 2: strSQL = strSQL & strCampo2 & " like '%" & txtDes & "%'"
   
    
    End Select
    txtCod = ""
    txtDes = ""
    
    With frmBusqGeneral
        Set .oParent = frmME
        .sQuery = strSQL
        .CARGAR_DATOS
        
        frmME.codigo = ""
        Set rstAux = .gexList.ADORecordset
        If rstAux.RecordCount > 0 Then
          .Show vbModal
        Else
          frmME.codigo = ".."
        End If
        
        If frmME.codigo <> "" And rstAux.RecordCount > 0 Then
            txtCod = frmME.codigo 'Trim(rstAux!Cod)
            txtDes = frmME.Descripcion  'Trim(rstAux!Descripcion)
            
            txtCod = Trim(rstAux!Cod)
            txtDes = Trim(rstAux!Descripcion)
            
            Select Case Opcion
            Case 1: SendKeys "{TAB}": SendKeys "{TAB}"
            Case 2: SendKeys "{TAB}"
            End Select
        Else
            SendKeys "{TAB}"
        End If
        
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Resume
Fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento (" & Opcion & ")"
End Sub




