VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMuestraVentaDiario 
   BackColor       =   &H00FFC0C0&
   Caption         =   "VENTAS DEL DIA"
   ClientHeight    =   9075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15495
   LinkTopic       =   "Form1"
   ScaleHeight     =   9075
   ScaleWidth      =   15495
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15495
      Begin VB.Frame FraEstilo 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ESTILO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   11400
         TabIndex        =   24
         Top             =   120
         Width           =   2775
         Begin VB.TextBox txtEstCliBus 
            Height          =   285
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame fraorden 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ORDENAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   11400
         TabIndex        =   21
         Top             =   120
         Width           =   1695
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "ALFABETICO"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "% VENTAS"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.Frame fraColTal 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MOSTRAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   9120
         TabIndex        =   17
         Top             =   120
         Width           =   2205
         Begin VB.CheckBox chkImagen 
            BackColor       =   &H00FFC0C0&
            Caption         =   "IMAGEN"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1200
            TabIndex        =   20
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chkColor 
            BackColor       =   &H00FFC0C0&
            Caption         =   "COLOR"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chkTalla 
            BackColor       =   &H00FFC0C0&
            Caption         =   "TALLA"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.Frame FraTipos 
         BackColor       =   &H00FFC0C0&
         Caption         =   "TIPO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   7440
         TabIndex        =   14
         Top             =   120
         Width           =   1575
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "ESTILOS"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "DOCUMENTOS"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.CheckBox chkTodas 
         BackColor       =   &H00FFC0C0&
         Caption         =   "TODAS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6120
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox CboCaja 
         Height          =   315
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   2280
      End
      Begin FunctionsButtons.FunctButt FunctButt2 
         Height          =   495
         Left            =   14160
         TabIndex        =   2
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
         Custom          =   "0~0~BUSCAR~Verdadero~Verdadero~&Buscar~0~0~1~~0~Falso~Falso~&Buscar~"
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin MSComCtl2.DTPicker dtpDia 
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   69795841
         CurrentDate     =   41000
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   1920
         TabIndex        =   8
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   69795841
         CurrentDate     =   41000
      End
      Begin VB.Frame FraResDet 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MOSTRAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   9120
         TabIndex        =   11
         Top             =   120
         Width           =   1335
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "RESUMEN"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "DETALLE"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "FIN :"
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
         Left            =   1920
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "CAJA :"
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
         Left            =   3840
         TabIndex        =   7
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "INICIO :"
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
         TabIndex        =   3
         Top             =   195
         Width           =   645
      End
   End
   Begin GridEX20.GridEX GridEX2 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   12938
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      BackColorBkg    =   12648384
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      FormatStylesCount=   7
      FormatStyle(1)  =   "FrmMuestraVentaDiario.frx":0000
      FormatStyle(2)  =   "FrmMuestraVentaDiario.frx":0138
      FormatStyle(3)  =   "FrmMuestraVentaDiario.frx":01E8
      FormatStyle(4)  =   "FrmMuestraVentaDiario.frx":029C
      FormatStyle(5)  =   "FrmMuestraVentaDiario.frx":0374
      FormatStyle(6)  =   "FrmMuestraVentaDiario.frx":042C
      FormatStyle(7)  =   "FrmMuestraVentaDiario.frx":050C
      ImageCount      =   0
      PrinterProperties=   "FrmMuestraVentaDiario.frx":052C
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   630
      Left            =   12600
      TabIndex        =   4
      Top             =   8400
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1111
      Custom          =   "0~0~IMPRIMIR~Verdadero~Verdadero~&Imprimir~0~0~1~~0~Falso~Falso~&Imprimir~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1500
      ControlHeigth   =   600
      ControlSeparator=   110
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   1200
      Top             =   5640
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "FrmMuestraVentaDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cadena As String
Public StrSQL As String
Private indiceResDet As Integer
Private indiceTipo  As Integer
Private flg_todascajas As String
Private flg_imagen As String
Public orden As Integer

Private Sub chkColor_Click()
Set GridEX2.ADORecordset = Nothing
End Sub

Private Sub chkTalla_Click()
Set GridEX2.ADORecordset = Nothing
End Sub

Private Sub chkTodas_Click()

flg_todascajas = "N"
If chkTodas.Value = 1 Then
    flg_todascajas = "S"
End If

End Sub

Private Sub dtpDia_Change()
DTPicker1 = Format(dtpDia, "dd/mm/yyyy")
End Sub

Private Sub Form_Load()
Dim sSeguridad  As String
chkTodas.Value = 0
dtpDia = Format(Date, "dd/mm/yyyy")
DTPicker1 = Format(Date, "dd/mm/yyyy")
indice = 0
CARGA_GRID_RESUMEN
CARGACAJAS
indiceResDet = 0
indiceTipo = 0
habilitaTipo (0)
flg_todascajas = "N"
orden = 0
habilitaTipoestilo (0)
End Sub
Private Sub CARGACAJAS()
On Error GoTo fin
Dim sTit As String
2
    sTit = "MUESTRA CAJAS"
    StrSQL = "CN_MUESTRA_CAJAS_VENTAS '001','001','' "
    
    Set rstAux = CargarRecordSetDesconectado(StrSQL, cConnect)
    CboCaja.Clear
    With rstAux
        If .RecordCount > 0 Then .MoveFirst
        Do Until .EOF
            CboCaja.AddItem !cod_caja & " " & !DES_CAJA
            .MoveNext
        Loop
        .Close
    End With
    If CboCaja.ListCount > 0 Then CboCaja.ListIndex = 0
    Set rstAux = Nothing
Exit Sub
fin:
    MsgBox err.Description, vbCritical + vbOKOnly, sTit
End Sub

Sub CARGA_GRID_RESUMEN()
On Error GoTo fin
StrSQL = "exec CN_MUESTRA_VENTAS_DIARIO_RESUMEN '" & Format(dtpDia, "dd/mm/yyyy") & "','" & Format(DTPicker1, "dd/mm/yyyy") & "','" & Left(CboCaja, 2) & "','" & flg_todascajas & "'"

Set GridEX2.ADORecordset = CargarRecordSetDesconectado(StrSQL, cConnect)
Call configuraGrillaResumen

Exit Sub
fin:
MsgBox err.Description, vbInformation + vbOKOnly, "IMPORTANTE"

End Sub
Sub CARGA_GRID_DETALLE()
On Error GoTo fin

StrSQL = "exec CN_MUESTRA_VENTAS_DIARIO_DETALLE '" & Format(dtpDia, "dd/mm/yyyy") & "','" & Format(DTPicker1, "dd/mm/yyyy") & "','" & Left(CboCaja, 2) & "','" & flg_todascajas & "','" & Trim(txtEstCliBus.Text) & "'"
Set GridEX2.ADORecordset = CargarRecordSetDesconectado(StrSQL, cConnect)
Call configuraGrillaDetalle

Exit Sub
fin:
MsgBox err.Description, vbInformation + vbOKOnly, "IMPORTANTE"
End Sub
Sub CARGA_GRID_ESTILO()
On Error GoTo fin
Dim flg_color As String
Dim flg_talla  As String

flg_color = "N"
If chkColor.Value = 1 Then
 flg_color = "S"
End If

flg_talla = "N"
If chkTalla.Value = 1 Then
 flg_talla = "S"
End If


StrSQL = "exec CN_MUESTRA_VENTAS_ESTILO_FECHAS '" & Format(dtpDia, "dd/mm/yyyy") & "','" & Format(DTPicker1, "dd/mm/yyyy") & "','" & Left(CboCaja, 2) & "','" & flg_todascajas & "','" & flg_color & "','" & flg_talla & "'," & orden & " "
Set GridEX2.ADORecordset = CargarRecordSetDesconectado(StrSQL, cConnect)
Call configuraGrillaEstilo

Exit Sub
fin:
MsgBox err.Description, vbInformation + vbOKOnly, "IMPORTANTE"
End Sub


Private Sub configuraGrillaResumen()
On Error GoTo fin

    Dim C As Integer
    With GridEX2
    
        For C = 1 To .Columns.Count
            With .Columns(C)
                .Caption = UCase(.Caption)
                .HeaderAlignment = jgexAlignCenter
                .TextAlignment = jgexAlignCenter
                .Visible = False
            End With
        Next C

        With .Columns("doc")
            .Width = 1500
            .TextAlignment = jgexAlignLeft
            .Caption = "DOCUMENTO"
            .Visible = True
        End With
        With .Columns("FECHA")
            .Width = 1300
            .TextAlignment = jgexAlignLeft
            .Caption = "FECHA"
            .Visible = True
        End With
        
        With .Columns("NUM_RUC")
            .Width = 1500
            .TextAlignment = jgexAlignLeft
            .Caption = "RUC"
            .Visible = True
        End With
        
        With .Columns("CLIENTE")
            .Width = 2500
            .TextAlignment = jgexAlignLeft
            .Caption = "CLIENTE"
            .Visible = True
        End With
        
        With .Columns("TIPO_CAMBIO")
            .Width = 800
            .TextAlignment = jgexAlignRight
            .Caption = "TC"
            .Visible = True
        End With
        
        With .Columns("TOTAL_VALOR_VENTA")
            .Width = 1000
            .TextAlignment = jgexAlignRight
            .Caption = "IMP NETO"
            .Visible = True
        End With
        With .Columns("IGV")
            .Width = 1000
            .TextAlignment = jgexAlignRight
            .Caption = "IGV"
            .Visible = True
        End With
        With .Columns("TOTAL")
            .Width = 1000
            .TextAlignment = jgexAlignRight
            .Caption = "IMP TOTAL"
            .Visible = True
        End With
        
        With .Columns("DES_TIPDOC")
            .Width = 1500
            .TextAlignment = jgexAlignLeft
            .Caption = "TIPO DOC"
            .Visible = True
        End With
        With .Columns("COD_MONEDA")
            .Width = 800
            .TextAlignment = jgexAlignCenter
            .Caption = "MONEDA"
            .Visible = True
        End With
        With .Columns("RAZONSOCIAL")
            .Width = 1500
            .TextAlignment = jgexAlignLeft
            .Caption = "EMPRESA"
            .Visible = True
        End With
        With .Columns("NUM_RUC_EMPRESA")
            .Width = 1500
            .TextAlignment = jgexAlignLeft
            .Caption = "NUM RUC EMPRESA"
            .Visible = True
        End With
        
'        With .Columns("stock")
'            .Width = 1500
'            .TextAlignment = jgexAlignLeft
'            .Caption = "STOCK"
'            .Visible = True
'        End With
        
      End With
Exit Sub
fin:
MsgBox "Hubo errores en la busqueda " & err.Description, vbInformation + vbOKOnly, "IMPORTANTE"
End Sub

Private Sub configuraGrillaDetalle()
On Error GoTo fin

    Dim C As Integer
    With GridEX2
    
        For C = 1 To .Columns.Count
            With .Columns(C)
                .Caption = UCase(.Caption)
                .HeaderAlignment = jgexAlignCenter
                .TextAlignment = jgexAlignCenter
                .Visible = False
            End With
        Next C

        With .Columns("fecha")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "FECHA"
            .Visible = True
        End With
        
        With .Columns("documento")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "DOCUMENTO"
            .Visible = True
        End With
        
        With .Columns("DES_ANEXO")
            .Width = 1500
            .TextAlignment = jgexAlignLeft
            .Caption = "DES_ANEXO"
            .Visible = True
        End With
        
        With .Columns("NUM_RUC")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "RUC"
            .Visible = True
        End With
 
        With .Columns("CAJA")
            .Width = 800
            .TextAlignment = jgexAlignLeft
            .Caption = "CAJA"
            .Visible = True
        End With
        With .Columns("COD_ESTCLI")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "ESTILO"
            .Visible = True
        End With
        With .Columns("DESCRIPCION")
            .Width = 1500
            .TextAlignment = jgexAlignLeft
            .Caption = "DESCRIPCION"
            .Visible = True
        End With
  
        With .Columns("COD_COLCLI")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "COD_COLCLI"
            .Visible = True
        End With
        With .Columns("COD_TALLA")
            .Width = 800
            .TextAlignment = jgexAlignLeft
            .Caption = "TALLA"
            .Visible = True
        End With
        With .Columns("CANTIDAD")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "CANT"
            .Visible = True
        End With
     
        With .Columns("PRECIO_LISTA")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "PRE. LISTA"
            .Visible = True
        End With
        With .Columns("IMP_TOTAL_LISTA")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "PRE. LISTA"
            .Visible = True
        End With
          
        With .Columns("CANTIDAD")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "CANTIDAD"
            .Visible = True
        End With
        With .Columns("IMP_UNITARIO_VENTA")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "PRE. VENTA"
            .Visible = True
        End With
        With .Columns("IMP_TOTAL_VENTA")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "IMP. VENTA"
            .Visible = True
        End With
        With .Columns("DIFERENCIA")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "DIF."
            .Visible = True
        End With
      
      End With
Exit Sub
fin:
MsgBox err.Description, vbInformation + vbOKOnly, "IMPORTANTE"
End Sub

Private Sub configuraGrillaEstilo()
On Error GoTo fin

    Dim C As Integer
    With GridEX2
    
        For C = 1 To .Columns.Count
            With .Columns(C)
                .Caption = UCase(.Caption)
                .HeaderAlignment = jgexAlignCenter
                .TextAlignment = jgexAlignCenter
                .Visible = False
            End With
        Next C


        With .Columns("CAJA")
            .Width = 800
            .TextAlignment = jgexAlignLeft
            .Caption = "CAJA"
            .Visible = True
        End With
        With .Columns("COD_ESTCLI")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "ESTILO"
            .Visible = True
        End With
        With .Columns("DESCRIPCION")
            .Width = 1500
            .TextAlignment = jgexAlignLeft
            .Caption = "DESCRIPCION"
            .Visible = True
        End With
  
        With .Columns("COD_COLCLI")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "COD_COLCLI"
            .Visible = True
        End With
        With .Columns("COD_TALLA")
            .Width = 800
            .TextAlignment = jgexAlignLeft
            .Caption = "TALLA"
            .Visible = True
        End With
        With .Columns("CANTIDAD")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "CANT"
            .Visible = True
        End With
     
        With .Columns("PRECIO_LISTA")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "PRE. LISTA"
            .Visible = True
        End With
        With .Columns("IMP_TOTAL_LISTA")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "PRE. LISTA"
            .Visible = True
        End With
          
        With .Columns("CANTIDAD")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "CANTIDAD"
            .Visible = True
        End With
        With .Columns("IMP_UNITARIO_VENTA")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "PRE. VENTA"
            .Visible = True
        End With
        With .Columns("IMP_TOTAL_VENTA")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "IMP. VENTA"
            .Visible = True
        End With
        With .Columns("porc")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "PORC."
            .Visible = True
        End With
      
        With .Columns("STOCK")
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "STOCK"
            .Visible = True
        End With
      
      
      End With
Exit Sub
fin:
MsgBox err.Description, vbInformation + vbOKOnly, "IMPORTANTE"
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
If GridEX2.RowCount <= 0 Then Exit Sub
Call Reporte
End Sub

Private Sub FunctButt2_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "BUSCAR"
    
        If indiceTipo = 0 Then
            If indiceResDet = 0 Then
                CARGA_GRID_RESUMEN
            Else
                CARGA_GRID_DETALLE
            End If
        Else
                CARGA_GRID_ESTILO
        End If
End Select
End Sub

Sub Reporte()
On Error GoTo ErrorImpresion
Dim smes As String, sRutaLogo As String
Dim oo As Object, lvSql As String
Dim XRS As New ADODB.Recordset
Dim Ruta As String

flg_imagen = "N"
If chkImagen.Value = 1 Then
 flg_imagen = "S"
End If

    StrSQL = "SELECT Ruta_Logo = ISNULL(Ruta_Logo, '') From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
    sRutaLogo = DevuelveCampo(StrSQL, cConnect)
    
    If indiceTipo = 0 Then
        If indiceResDet = 0 Then
            cadena = "exec CN_MUESTRA_VENTAS_DIARIO_RESUMEN '" & Format(dtpDia, "dd/mm/yyyy") & "','" & Format(DTPicker1, "dd/mm/yyyy") & "','" & Left(CboCaja, 2) & "','" & flg_todascajas & "'"
            Ruta = vRuta & "\" & "Rpt_Registro_Ventas_Diaria_Resumen.xlt"
        Else
            Ruta = vRuta & "\" & "RptVentasDiariaDetalle.xlt"
        End If
    Else
            Ruta = vRuta & "\" & "RptVentasEstilo.xlt"
    End If

    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta

    oo.Visible = True
    oo.DisplayAlerts = False
   
    If indiceTipo = 0 Then
        If indiceResDet = 0 Then
            oo.Run "reporte", dtpDia, DTPicker1, cadena, cConnect, sRutaLogo
        Else
            oo.Run "reporte", GridEX2.ADORecordset, "MUESTRA DETALLE DE VENTAS POR DOCUMENTOS DEL DIA " & Format(dtpDia, "dd/mm/yyyy") & " HASTA EL DIA " & Format(DTPicker1, "DD/MM/YYYY")
        End If
    Else
            oo.Run "reporte", GridEX2.ADORecordset, "MUESTRA DETALLE DE VENTAS POR ESTILOS DEL DIA " & Format(dtpDia, "dd/mm/yyyy") & " HASTA EL DIA " & Format(DTPicker1, "DD/MM/YYYY"), flg_imagen
    End If
    Set oo = Nothing
    Exit Sub
    
ErrorImpresion:
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte de Venta diaria " & err.Description, vbInformation + vbOKOnly, "Impresion"
End Sub
Private Sub Option1_Click(Index As Integer)
   Set GridEX2.ADORecordset = Nothing
    indiceResDet = Index
    habilitaTipoestilo (indiceResDet)
End Sub
Private Sub Option2_Click(Index As Integer)
    Set GridEX2.ADORecordset = Nothing
    indiceTipo = Index
    habilitaTipo (indiceTipo)
End Sub
Private Sub habilitaTipo(tipo As Integer)
    FraResDet.Visible = False
    fraColTal.Visible = False
    fraorden.Visible = False
    limpiaCajas
    
    If tipo = 0 Then
        FraResDet.Visible = True
    Else
        fraColTal.Visible = True
        fraorden.Visible = True
    End If
    
    
End Sub
Private Sub habilitaTipoestilo(tipo As Integer)
    FraEstilo.Visible = False
    limpiaCajas

    If indiceTipo = 0 Then
        If tipo = 1 Then
              FraEstilo.Visible = True
        End If
    End If

End Sub
Private Sub Option3_Click(Index As Integer)
    orden = Index
    Set GridEX2.ADORecordset = Nothing
    limpiaCajas
End Sub
Private Sub limpiaCajas()
    txtEstCliBus.Text = ""
End Sub
