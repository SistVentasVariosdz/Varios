VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmCambioModeloTalla 
   BackColor       =   &H00FFC0C0&
   Caption         =   "CAMBIO DE MODELO Y TALLA"
   ClientHeight    =   9255
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   ScaleHeight     =   9255
   ScaleWidth      =   15960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&CANCELAR"
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   8760
      Width           =   1725
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "CAMBIO MODELO-TALLA CLIENTES"
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
      Left            =   7200
      TabIndex        =   49
      Top             =   120
      Width           =   3015
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "CAMBIO MODELO-TALLA INTERNO"
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
      Left            =   4200
      TabIndex        =   48
      Top             =   120
      Value           =   -1  'True
      Width           =   3015
   End
   Begin VB.TextBox txtDocReferencia 
      Height          =   285
      Left            =   11640
      TabIndex        =   46
      Top             =   120
      Width           =   2820
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Height          =   5295
      Left            =   8280
      TabIndex        =   44
      Top             =   3360
      Width           =   7695
      Begin GridEX20.GridEX grxDatos 
         Height          =   4995
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   8811
         Version         =   "2.0"
         RecordNavigator =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         GridLineStyle   =   2
         HideSelection   =   2
         MethodHoldFields=   -1  'True
         GroupByBoxInfoText=   "Arrastra la cabecera de una columna aquí para agruparlo por esa misma columna"
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         HeaderFontName  =   "Verdana"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   6.75
         HeaderFontWeight=   700
         FontName        =   "Tahoma"
         BackColorBkg    =   12648384
         ColumnHeaderHeight=   270
         IntProp1        =   0
         ColumnsCount    =   2
         Column(1)       =   "FrmCambioModeloTalla.frx":0000
         Column(2)       =   "FrmCambioModeloTalla.frx":00C8
         FormatStylesCount=   9
         FormatStyle(1)  =   "FrmCambioModeloTalla.frx":016C
         FormatStyle(2)  =   "FrmCambioModeloTalla.frx":0294
         FormatStyle(3)  =   "FrmCambioModeloTalla.frx":0344
         FormatStyle(4)  =   "FrmCambioModeloTalla.frx":03F8
         FormatStyle(5)  =   "FrmCambioModeloTalla.frx":04D0
         FormatStyle(6)  =   "FrmCambioModeloTalla.frx":0588
         FormatStyle(7)  =   "FrmCambioModeloTalla.frx":0668
         FormatStyle(8)  =   "FrmCambioModeloTalla.frx":06F8
         FormatStyle(9)  =   "FrmCambioModeloTalla.frx":0830
         ImageCount      =   0
         PrinterProperties=   "FrmCambioModeloTalla.frx":0944
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   5295
      Left            =   0
      TabIndex        =   42
      Top             =   3360
      Width           =   8175
      Begin GridEX20.GridEX GridEX1 
         Height          =   4905
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   8010
         _ExtentX        =   14129
         _ExtentY        =   8652
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         GridLineStyle   =   2
         MethodHoldFields=   -1  'True
         Options         =   8
         RecordsetType   =   1
         GroupByBoxVisible=   0   'False
         DataMode        =   1
         HeaderFontName  =   "Verdana"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   6.75
         HeaderFontWeight=   700
         FontName        =   "Tahoma"
         BackColorBkg    =   12648384
         ColumnHeaderHeight=   270
         FormatStylesCount=   7
         FormatStyle(1)  =   "FrmCambioModeloTalla.frx":0B1C
         FormatStyle(2)  =   "FrmCambioModeloTalla.frx":0C44
         FormatStyle(3)  =   "FrmCambioModeloTalla.frx":0CF4
         FormatStyle(4)  =   "FrmCambioModeloTalla.frx":0DA8
         FormatStyle(5)  =   "FrmCambioModeloTalla.frx":0E80
         FormatStyle(6)  =   "FrmCambioModeloTalla.frx":0F38
         FormatStyle(7)  =   "FrmCambioModeloTalla.frx":1018
         ImageCount      =   0
         PrinterProperties=   "FrmCambioModeloTalla.frx":1038
      End
   End
   Begin VB.CommandButton cmdRealizarCambios 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&REALIZAR CAMBIO"
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
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   8760
      Width           =   1725
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ENTRADAS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   8160
      TabIndex        =   21
      Top             =   480
      Width           =   7785
      Begin VB.TextBox TxtCod_EstcliEn 
         Height          =   285
         Left            =   840
         TabIndex        =   30
         Top             =   240
         Width           =   1500
      End
      Begin VB.TextBox txtDes_EstcliEn 
         Height          =   285
         Left            =   3120
         TabIndex        =   29
         Top             =   240
         Width           =   4605
      End
      Begin VB.TextBox txtOpEn 
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   28
         Top             =   600
         Width           =   1500
      End
      Begin VB.TextBox txtCod_PresentEn 
         Height          =   285
         Left            =   840
         TabIndex        =   27
         Top             =   960
         Width           =   1500
      End
      Begin VB.TextBox txtDes_PresentEn 
         Height          =   285
         Left            =   3120
         TabIndex        =   26
         Top             =   960
         Width           =   4605
      End
      Begin VB.TextBox txtCod_TallaEn 
         Height          =   285
         Left            =   840
         TabIndex        =   25
         Top             =   1320
         Width           =   1500
      End
      Begin VB.TextBox txtStocksActualEn 
         Height          =   285
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1680
         Width           =   1500
      End
      Begin VB.TextBox txtCantidadEn 
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   23
         Top             =   2040
         Width           =   1500
      End
      Begin VB.TextBox txtNewStockEn 
         Height          =   285
         Left            =   3135
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1680
         Width           =   1500
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFC0C0&
         Caption         =   "STK NVO:"
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
         Left            =   2415
         TabIndex        =   41
         Tag             =   "Document Type"
         Top             =   1800
         Width           =   795
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFC0C0&
         Caption         =   "CODIGO:"
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
         Left            =   120
         TabIndex        =   38
         Tag             =   "Document Type"
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MODELO:"
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
         Left            =   2400
         TabIndex        =   37
         Tag             =   "Document Type"
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFC0C0&
         Caption         =   "OP:"
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
         Left            =   120
         TabIndex        =   36
         Tag             =   "Document Type"
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFC0C0&
         Caption         =   "CODIGO:"
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
         Left            =   120
         TabIndex        =   35
         Tag             =   "Document Type"
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFC0C0&
         Caption         =   "COLOR:"
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
         Left            =   2400
         TabIndex        =   34
         Tag             =   "Document Type"
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFC0C0&
         Caption         =   "TALLA:"
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
         Left            =   120
         TabIndex        =   33
         Tag             =   "Document Type"
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "STK ACT:"
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
         Left            =   135
         TabIndex        =   32
         Tag             =   "Document Type"
         Top             =   1800
         Width           =   795
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "CANT:"
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
         Left            =   120
         TabIndex        =   31
         Tag             =   "Document Type"
         Top             =   2040
         Width           =   435
      End
   End
   Begin VB.ComboBox cboAlmacen 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   120
      Width           =   3315
   End
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&AGREGAR"
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
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1365
   End
   Begin VB.Frame FraSalida 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SALIDAS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   8145
      Begin VB.TextBox txtNewStocks 
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtCantidad 
         Height          =   285
         Left            =   840
         TabIndex        =   16
         Top             =   2040
         Width           =   1500
      End
      Begin VB.TextBox txtStocksActual 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1680
         Width           =   1500
      End
      Begin VB.TextBox txtTalla 
         Height          =   285
         Left            =   840
         TabIndex        =   12
         Top             =   1320
         Width           =   1500
      End
      Begin VB.TextBox txtDes_present 
         Height          =   285
         Left            =   3120
         TabIndex        =   9
         Top             =   960
         Width           =   4605
      End
      Begin VB.TextBox txtCod_present 
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Top             =   960
         Width           =   1500
      End
      Begin VB.TextBox txtOp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Top             =   600
         Width           =   1500
      End
      Begin VB.TextBox txtDes_Estcli 
         Height          =   285
         Left            =   3120
         TabIndex        =   2
         Top             =   240
         Width           =   4605
      End
      Begin VB.TextBox TxtCod_estcli 
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFC0C0&
         Caption         =   "STK NVO:"
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
         Left            =   2400
         TabIndex        =   40
         Tag             =   "Document Type"
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "CANT:"
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
         Left            =   120
         TabIndex        =   17
         Tag             =   "Document Type"
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "STK ACT:"
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
         Left            =   120
         TabIndex        =   15
         Tag             =   "Document Type"
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "TALLA:"
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
         Left            =   120
         TabIndex        =   13
         Tag             =   "Document Type"
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "COLOR:"
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
         Left            =   2400
         TabIndex        =   11
         Tag             =   "Document Type"
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "CODIGO:"
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
         Left            =   120
         TabIndex        =   10
         Tag             =   "Document Type"
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "OP:"
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
         Left            =   120
         TabIndex        =   7
         Tag             =   "Document Type"
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MODELO:"
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
         Index           =   0
         Left            =   2400
         TabIndex        =   5
         Tag             =   "Document Type"
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "CODIGO:"
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
         Left            =   120
         TabIndex        =   3
         Tag             =   "Document Type"
         Top             =   240
         Width           =   675
      End
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   2520
      Top             =   3120
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFC0C0&
      Caption         =   "DOC DE REFER:"
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
      Left            =   10440
      TabIndex        =   47
      Tag             =   "Document Type"
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ALMACEN:"
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
      Left            =   0
      TabIndex        =   20
      Top             =   165
      Width           =   750
   End
End
Attribute VB_Name = "FrmCambioModeloTalla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public CODIGO As String, Descripcion As String
Dim strSQL As String, sTit_OP As String, rstAux As ADODB.Recordset
Public sopcion  As String
Public scod_estcli As String
Public fila_seleccionada As Long
Dim bClickColSelec As Boolean
Private indice As Integer
Private Sub cmdAgregar_Click()
  If validaagregarCambioModeloTalla = True Then
      Call adicionaCambioEstilo
      Call edicionaGrillaEntrada
      Call limpiaControles
  End If
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub GridEX1_Click()
    Dim ColIndex As Long
    If GridEX1.RowCount > 0 Then
        ColIndex = GridEX1.Col
         If ColIndex = 0 Then Exit Sub
            If UCase(GridEX1.Columns(ColIndex).Key) = "ELI" Then
                bClickColSelec = True
                SendKeys "{ENTER}"
            End If
    End If
End Sub
'''*******************EVENTOS POR COLUMNA **********************************************************
Private Sub GridEX1_AfterColEdit(ByVal ColIndex As Integer)
  AfterColEdit_modelo (ColIndex)
End Sub
Sub AfterColEdit_modelo(ByVal ColIndex As Integer)
Dim sSQL As String
On Error GoTo Error_Handler
Dim oGroup As GridEX20.JSGroup

Select Case ColIndex
Case Is = GridEX1.Columns("ELI").Index
        Call EliminaModelo
End Select
Exit Sub
Resume
Error_Handler:
errores Err.Number
End Sub
'''************************************************************ELIMINA modelo****************************
Private Sub EliminaModelo()
On Error GoTo fin

    If GridEX1.RowCount = 0 Then Exit Sub
    
    Dim I As Integer
    Dim rstAux  As ADODB.Recordset
    Dim rxsalida As ADODB.Recordset
    
    grxDatos.Update
    Set rxsalida = grxDatos.ADORecordset
    
    GridEX1.Update
    Set rstAux = GridEX1.ADORecordset
    
    rstAux.MoveFirst
    I = 1
    Do While I <= rstAux.RecordCount
        If rstAux("ELI").Value = True Then
           rstAux.AbsolutePosition = GridEX1.RowIndex(GridEX1.Row)
           rstAux.Delete
          
           rxsalida.AbsolutePosition = GridEX1.RowIndex(GridEX1.Row)
           rxsalida.Delete
          
        Else
          rstAux("ELI") = 0
        End If
        rstAux.MoveNext
        I = I + 1
    Loop
    Set GridEX1.ADORecordset = rstAux
    Set grxDatos.ADORecordset = rxsalida
    
    Call configuragrillaEntradas
    Call configuragrillaSalidas
Exit Sub
fin:
MsgBox "Problemas al Eliminar El modelo" + Err.Description, vbInformation + vbOKOnly, "IMPORTANTE"
    
End Sub
Private Sub limpiaControles()

TxtCod_estcli.Text = ""
txtDes_Estcli.Text = ""

TxtCod_EstcliEn.Text = ""
txtDes_EstcliEn.Text = ""

txtOp.Text = ""
txtOpEn.Text = ""

txtCod_present.Text = ""
txtCod_PresentEn.Text = ""

txtDes_present.Text = ""
txtDes_PresentEn.Text = ""

txtTalla.Text = ""
txtCod_TallaEn.Text = ""

txtStocksActual.Text = ""
txtStocksActualEn.Text = ""

txtNewStocks.Text = ""
txtNewStockEn.Text = ""

txtCantidad.Text = 0
txtCantidadEn.Text = 0

txtStocksActual.Text = 0
txtStocksActualEn.Text = 0

txtNewStockEn.Text = 0
txtNewStocks.Text = 0


End Sub

Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Load()
    Call FillAlmacen
    Call InicioGrilla
    Call incioControles
    indice = 0
End Sub
Private Sub incioControles()
    txtCantidad.Text = 0
    txtCantidadEn.Text = 0
    txtStocksActual.Enabled = False
    txtStocksActualEn.Enabled = False
    txtNewStockEn.Enabled = False
    txtNewStocks.Enabled = False
End Sub
Private Function validaagregarCambioModeloTalla() As Boolean

validaagregarCambioModeloTalla = True

If TxtCod_estcli.Text = "" Or txtDes_Estcli.Text = "" Then
    validaagregarCambioModeloTalla = False
    MsgBox "Ingrese Un estilo Valido", vbInformation + vbOKOnly, "IMPORTANTE"
    Exit Function
End If

If TxtCod_EstcliEn.Text = "" Or txtDes_EstcliEn.Text = "" Then
    validaagregarCambioModeloTalla = False
    MsgBox "Ingrese Un estilo Valido", vbInformation + vbOKOnly, "IMPORTANTE"
    Exit Function
End If

If txtCod_present.Text = "" Or txtDes_present.Text = "" Then
    validaagregarCambioModeloTalla = False
    MsgBox "Ingrese Un color de salida Valido", vbInformation + vbOKOnly, "IMPORTANTE"
    Exit Function
End If

If txtCod_PresentEn.Text = "" Or txtDes_PresentEn.Text = "" Then
    validaagregarCambioModeloTalla = False
    MsgBox "Ingrese Un color de entrada Valido", vbInformation + vbOKOnly, "IMPORTANTE"
    Exit Function
End If

If txtTalla.Text = "" Then
    validaagregarCambioModeloTalla = False
    MsgBox "Ingrese una talla valida", vbInformation + vbOKOnly, "IMPORTANTE"
    Exit Function
End If

If txtCod_TallaEn.Text = "" Then
    validaagregarCambioModeloTalla = False
    MsgBox "Ingrese una talla valida", vbInformation + vbOKOnly, "IMPORTANTE"
    Exit Function
End If

If txtCantidad.Text = "" Then
    validaagregarCambioModeloTalla = False
    MsgBox "Ingrese ingrese una cantidad de salida valida", vbInformation + vbOKOnly, "IMPORTANTE"
    Exit Function
End If

If txtStocksActual.Text < txtCantidad + stkGrilla Then
    validaagregarCambioModeloTalla = False
    MsgBox "La cantidad a transferir no puede ser mayor que el stocks actual", vbInformation + vbOKOnly, "IMPORTANTE"
    Exit Function
End If

End Function
Private Function stkGrilla() As Integer
If GridEX1.RowCount <= 0 Then Exit Function

Dim rxsal As New ADODB.Recordset
Dim stk As Integer

GridEX1.Update
Set rxsal = GridEX1.ADORecordset
'rxsal.Update
Dim I As Integer
rxsal.MoveFirst
Do While Not rxsal.EOF
    If Trim(rxsal!cod_estcli) = Trim(TxtCod_estcli.Text) And Trim(rxsal!cod_ordpro) = Trim(txtOp.Text) And Trim(rxsal!cod_present) = Trim(txtCod_present.Text) And Trim(rxsal!cod_talla) = Trim(txtTalla.Text) Then
      stk = stk + Trim(rxsal!cantidad)
    End If
    rxsal.MoveNext
Loop

stkGrilla = stk

End Function

Private Sub FillAlmacen()
Dim rstAlm As ADODB.Recordset

    strSQL = "SM_MUESTRA_ALMACEN_CAMBIO_MODELO_TALLA '" & vusu & "' "
    Set rstAlm = CargarRecordSetDesconectado(strSQL, cConnect)
    rstAlm.MoveFirst
    cboAlmacen.Clear
    Do Until rstAlm.EOF
        cboAlmacen.AddItem rstAlm!Cod_almacen & " " & rstAlm!nom_almacen
        rstAlm.MoveNext
    Loop
    rstAlm.Close
    Set rstAlm = Nothing
    
End Sub
Private Sub InicioGrilla()
    strSQL = "SM_MUESTRA_ESTILO_CAMBIO_MODELO_TALLA '1','" & Left(cboAlmacen, 2) & "','','','','','',''"
    Set GridEX1.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    configuragrillaSalidas
    
    Set grxDatos.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)
    configuragrillaEntradas
    
End Sub
Private Sub adicionaCambioEstilo()
    Dim rsxSalidas As New ADODB.Recordset
    Dim nro As Integer
    On Error GoTo fin
    nro = GridEX1.RowCount
    '''salidas
    
    GridEX1.Update
    nro = nro + 1
    Set rsxSalidas = Nothing
    Set rsxSalidas = GridEX1.ADORecordset
    rsxSalidas.AddNew
    rsxSalidas!Numero = nro
    rsxSalidas!cod_estcli = TxtCod_estcli.Text
    rsxSalidas!des_estcli = txtDes_Estcli.Text
    rsxSalidas!cod_ordpro = txtOp.Text
    rsxSalidas!cod_present = txtCod_present.Text
    rsxSalidas!des_present = txtDes_present.Text
    rsxSalidas!cod_talla = txtTalla.Text
    rsxSalidas!cantidad = txtCantidad
    Set GridEX1.ADORecordset = Nothing
    Set GridEX1.ADORecordset = rsxSalidas
    GridEX1.Update
    Call configuragrillaSalidas
    
    Exit Sub
fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
    
End Sub
Private Sub edicionaGrillaEntrada()
On Error GoTo fin
    Dim rsxEntradas As New ADODB.Recordset
    Dim rsxTemp As New ADODB.Recordset
    Dim nro As Integer
    Dim I, j, k, l As Integer
    On Error GoTo fin
    grxDatos.Update
    
    nro = grxDatos.RowCount
    '''entradas
    grxDatos.Update
    nro = nro + 1
    grxDatos.Update
    Set rsxEntradas = Nothing
    Set rsxEntradas = grxDatos.ADORecordset
    Set grxDatos.ADORecordset = Nothing
    I = rsxEntradas.RecordCount
   
    rsxEntradas.AddNew
    j = rsxEntradas.RecordCount
    rsxEntradas!Numero = nro
    rsxEntradas!cod_estcli = TxtCod_EstcliEn.Text
    rsxEntradas!des_estcli = txtDes_EstcliEn.Text
    rsxEntradas!cod_ordpro = txtOpEn.Text
    rsxEntradas!cod_present = txtCod_PresentEn
    rsxEntradas!des_present = txtDes_PresentEn.Text
    rsxEntradas!cod_talla = txtCod_TallaEn.Text
    rsxEntradas!cantidad = txtCantidadEn.Text
    k = rsxEntradas.RecordCount
    Set grxDatos.ADORecordset = Nothing
    Set grxDatos.ADORecordset = rsxEntradas
    grxDatos.Update
    Call configuragrillaEntradas
    Exit Sub
fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub

Private Sub configuragrillaSalidas()
    Dim C As Integer
    Dim colTemp As JSColumn
    Dim fmtCon  As JSFmtCondition
    On Error GoTo fin
    With GridEX1
         For C = 1 To .Columns.Count
            .Columns(C).HeaderAlignment = jgexAlignCenter
            .Columns(C).TextAlignment = jgexAlignLeft
            .Columns(C).Visible = False
        Next C

        With .Columns("NUMERO")
             .Visible = True
             .Width = 500
             .Caption = "NRO"
             .TextAlignment = jgexAlignLeft
        End With

        With .Columns("cod_estcli")
             .Visible = True
             .Width = 1500
             .Caption = "Codigo"
             .TextAlignment = jgexAlignLeft
        End With

        With .Columns("des_estcli")
             .Visible = True
             .Width = 2000
             .Caption = "Modelo"
             .TextAlignment = jgexAlignLeft
        End With
        With .Columns("cod_ordpro")
             .Visible = True
             .Width = 800
             .Caption = "OP"
             .TextAlignment = jgexAlignLeft
        End With
        With .Columns("des_present")
             .Visible = True
             .Width = 1000
             .Caption = "Color"
             .TextAlignment = jgexAlignLeft
        End With
        With .Columns("cod_talla")
             .Visible = True
             .Width = 700
             .Caption = "Talla"
             .TextAlignment = jgexAlignLeft
        End With
        With .Columns("cantidad")
             .Visible = True
             .Width = 700
             .Caption = "Cantidad"
             .TextAlignment = jgexAlignLeft
        End With
        With .Columns("ELI")
             .Visible = True
             .Width = 700
             .Caption = "ELI"
             .TextAlignment = jgexAlignLeft
        End With

    End With

    Exit Sub
fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub

Private Sub configuragrillaEntradas()
    Dim C As Integer
    Dim colTemp As JSColumn
    Dim fmtCon  As JSFmtCondition
    On Error GoTo fin
    With grxDatos
         For C = 1 To .Columns.Count
            .Columns(C).HeaderAlignment = jgexAlignCenter
            .Columns(C).TextAlignment = jgexAlignLeft
            .Columns(C).Visible = False
        Next C
        With .Columns("NUMERO")
             .Visible = True
             .Width = 500
             .Caption = "NRO"
             .TextAlignment = jgexAlignLeft
        End With
        With .Columns("cod_estcli")
             .Visible = True
             .Width = 1500
             .Caption = "Codigo"
             .TextAlignment = jgexAlignLeft
        End With
        With .Columns("des_estcli")
             .Visible = True
             .Width = 2000
             .Caption = "Modelo"
             .TextAlignment = jgexAlignLeft
        End With
        With .Columns("cod_ordpro")
             .Visible = True
             .Width = 800
             .Caption = "OP"
             .TextAlignment = jgexAlignLeft
        End With
        With .Columns("des_present")
             .Visible = True
             .Width = 1000
             .Caption = "Color"
             .TextAlignment = jgexAlignLeft
        End With
        With .Columns("cod_talla")
             .Visible = True
             .Width = 700
             .Caption = "Talla"
             .TextAlignment = jgexAlignLeft
        End With
        With .Columns("cantidad")
             .Visible = True
             .Width = 700
             .Caption = "Cantidad"
             .TextAlignment = jgexAlignLeft
        End With
    End With

    Exit Sub
fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, Me.Caption
End Sub

Private Sub Option1_Click(Index As Integer)
indice = Index

End Sub

Private Sub txtCantidad_Change()
If txtCantidad.Text <> "" Then
  'txtCantidad.Text = 0
    txtCantidadEn.Text = Val(txtCantidad.Text)
    If txtStocksActual.Text = "" Then txtStocksActual.Text = 0
    If txtNewStocks.Text = "" Then txtNewStocks.Text = 0

    txtNewStocks.Text = Val(txtStocksActual.Text) - Val(txtCantidad.Text)
    txtNewStockEn.Text = Val(txtStocksActualEn) + Val(txtCantidadEn)

End If

End Sub

Private Sub txtCantidad_LostFocus()
If txtCantidad.Text = "" Then
  txtCantidad.Text = 0
  txtCantidadEn.Text = txtCantidad.Text
End If

End Sub

Private Sub TxtCod_EstcliEn_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  
        txtOpEn.Text = ""
        txtCod_PresentEn.Text = ""
        txtDes_PresentEn.Text = ""
        txtCod_TallaEn.Text = ""
        txtStocksActualEn.Text = ""
        txtNewStockEn.Text = ""
        
        Call Busca_OpcioneEstiloEntrada(2, Left(cboAlmacen, 2), TxtCod_EstcliEn.Text, txtDes_EstcliEn, txtOpEn.Text, txtCod_PresentEn.Text, txtDes_PresentEn.Text, txtCod_TallaEn.Text, TxtCod_EstcliEn, txtDes_EstcliEn)
        If Trim(txtDes_EstcliEn.Text) <> "" Then
            txtCod_PresentEn.SetFocus
        Else
            TxtCod_EstcliEn.SetFocus
        End If
  
  End If
End Sub


Private Sub txtDes_EstcliEn_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  
        txtOpEn.Text = ""
        txtCod_present.Text = ""
        txtDes_PresentEn.Text = ""
        txtCod_TallaEn.Text = ""
        txtStocksActualEn.Text = ""
        txtNewStockEn.Text = ""

        Call Busca_OpcioneEstiloEntrada(3, Left(cboAlmacen, 2), TxtCod_EstcliEn.Text, txtDes_EstcliEn, txtOpEn.Text, txtCod_PresentEn.Text, txtDes_PresentEn.Text, txtCod_TallaEn.Text, TxtCod_EstcliEn, txtDes_EstcliEn)
        If Trim(txtDes_EstcliEn.Text) <> "" Then
            txtCod_PresentEn.SetFocus
        Else
            txtDes_EstcliEn.SetFocus
        End If
        
  End If
End Sub
Private Sub txtCod_PresentEn_keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        txtCod_TallaEn.Text = ""
        txtStocksActualEn.Text = ""
        txtNewStockEn.Text = ""

        Call Busca_OpcioneEstiloEntrada(4, Left(cboAlmacen, 2), TxtCod_EstcliEn.Text, txtDes_EstcliEn, txtOpEn.Text, txtCod_PresentEn.Text, txtDes_PresentEn.Text, txtCod_TallaEn.Text, txtCod_PresentEn, txtDes_PresentEn)
        If Trim(txtCod_PresentEn.Text) <> "" Then
            txtCod_TallaEn.SetFocus
        Else
            txtCod_PresentEn.SetFocus
        End If
  End If
End Sub
Private Sub txtDes_PresentEn_keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        txtCod_TallaEn.Text = ""
        txtStocksActualEn.Text = ""
        txtNewStockEn.Text = ""

        Call Busca_OpcioneEstiloEntrada(5, Left(cboAlmacen, 2), TxtCod_EstcliEn.Text, txtDes_EstcliEn, txtOpEn.Text, txtCod_PresentEn.Text, txtDes_PresentEn.Text, txtCod_TallaEn.Text, txtCod_PresentEn, txtDes_PresentEn)
        If Trim(txtDes_PresentEn.Text) <> "" Then
            txtCod_TallaEn.SetFocus
        Else
            txtDes_PresentEn.SetFocus
        End If
        
  End If
End Sub
Private Sub txtcod_TallaEn_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        Call Busca_OpcioneEstiloEntrada(6, Left(cboAlmacen, 2), TxtCod_EstcliEn.Text, txtDes_EstcliEn, txtOpEn.Text, txtCod_PresentEn.Text, txtDes_PresentEn.Text, txtCod_TallaEn.Text, txtCod_TallaEn, txtCod_TallaEn)
        If Trim(txtCod_TallaEn.Text) <> "" Then
            txtCantidadEn.SetFocus
        Else
            txtCod_TallaEn.SetFocus
        End If
  End If
End Sub
'''''''busqueda de salidas
Private Sub TxtCod_estcli_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        txtOp.Text = ""
        txtCod_present.Text = ""
        txtDes_present.Text = ""
        txtTalla.Text = ""
        txtStocksActual.Text = ""
        txtNewStocks.Text = ""
        txtCantidad.Text = 0
        
        Call Busca_OpcioneEstilo(2, Left(cboAlmacen, 2), TxtCod_estcli.Text, txtDes_Estcli, txtOp.Text, txtCod_present.Text, txtDes_present.Text, txtTalla.Text, TxtCod_estcli, txtDes_Estcli)
        If Trim(TxtCod_estcli.Text) <> "" Then
            txtCod_present.SetFocus
        Else
            TxtCod_estcli.SetFocus
        End If
        

        
  End If
End Sub
Private Sub txtDes_Estcli_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  
        txtOp.Text = ""
        txtCod_present.Text = ""
        txtDes_present.Text = ""
        txtTalla.Text = ""
        txtStocksActual.Text = ""
        txtNewStocks.Text = ""
        txtCantidad.Text = 0
        
        Call Busca_OpcioneEstilo(3, Left(cboAlmacen, 2), TxtCod_estcli.Text, txtDes_Estcli, txtOp.Text, txtCod_present.Text, txtDes_present.Text, txtTalla.Text, TxtCod_estcli, txtDes_Estcli)
        If Trim(txtDes_Estcli.Text) <> "" Then
            txtCod_present.SetFocus
        Else
            txtDes_Estcli.SetFocus
        End If
        

  End If
End Sub
Private Sub txtCod_present_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  
        txtTalla.Text = ""
        txtStocksActual.Text = ""
        txtNewStocks.Text = ""
        txtCantidad.Text = 0

        Call Busca_OpcioneEstilo(4, Left(cboAlmacen, 2), TxtCod_estcli.Text, txtDes_Estcli, txtOp.Text, txtCod_present.Text, txtDes_present.Text, txtTalla.Text, txtCod_present, txtDes_present)
        If Trim(txtCod_present.Text) <> "" Then
            txtTalla.SetFocus
        Else
            txtCod_present.SetFocus
        End If
  End If
End Sub
Private Sub txtDes_present_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        Call Busca_OpcioneEstilo(5, Left(cboAlmacen, 2), TxtCod_estcli.Text, txtDes_Estcli, txtOp.Text, txtCod_present.Text, txtDes_present.Text, txtTalla.Text, txtCod_present, txtDes_present)
        If Trim(txtDes_present.Text) <> "" Then
            txtTalla.SetFocus
        Else
            txtDes_present.SetFocus
        End If
        
        txtTalla.Text = ""
        txtStocksActual.Text = ""
        txtNewStocks.Text = ""
        txtCantidad.Text = 0
  End If
End Sub
Private Sub txtTalla_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        Call Busca_OpcioneEstilo(6, Left(cboAlmacen, 2), TxtCod_estcli.Text, txtDes_Estcli, txtOp.Text, txtCod_present.Text, txtDes_present.Text, txtTalla.Text, txtTalla, txtTalla)
        If Trim(txtTalla.Text) <> "" Then
            'TxtCod_EstcliEn.SetFocus
            txtCantidad.SetFocus
        Else
            txtTalla.SetFocus
        End If
        txtCantidad.Text = 0
        txtCantidadEn.Text = 0
  End If
End Sub
Sub Busca_OpcioneEstiloEntrada(sopcion As String, sCod_Almacen As String, scod_estcli As String, sdes_estcli As String, scod_ordpro As String, scod_present As String, sdes_present As String, sCod_Talla As String, txtCod As TextBox, txtDes As TextBox)
On Error GoTo fin
Dim rstAux As ADODB.Recordset

    If sCod_Almacen = "" Then
        MsgBox "seleccione un almacen", vbInformation + vbOKOnly, "Importante"
        Exit Sub
    End If

    strSQL = " SM_MUESTRA_ESTILO_CAMBIO_MODELO_TALLA '" & sopcion & "', '" & sCod_Almacen & "','" & scod_estcli & "','" & sdes_estcli & "','" & scod_ordpro & "','" & scod_present & "','" & sdes_present & "','" & sCod_Talla & "' "
    fila_seleccionada = 0

    With frmBusqGeneral
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos

        CODIGO = ".."
        Set rstAux = .gexList.ADORecordset
        'If rstAux.RecordCount > 1 Then
        .Show vbModal

        If fila_seleccionada > 0 And rstAux.RecordCount > 0 Then

            rstAux.AbsolutePosition = fila_seleccionada
            txtCod = Trim(rstAux!Cod)
            txtDes = Trim(rstAux!Descripcion)
            If sopcion = 2 Or sopcion = 3 Then
                txtOpEn.Text = rstAux!op
            End If
            If sopcion = 4 Or sopcion = 5 Then
                txtCod_TallaEn.Text = rstAux!cod_talla
                txtStocksActualEn.Text = rstAux!STOCKS
                txtNewStockEn.Text = rstAux!STOCKS + txtCantidadEn.Text
            End If
        Else
            txtCod = ""
            txtDes = ""
            SendKeys "{TAB}"
        End If

    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento (" & sopcion & ")"
End Sub

Sub Busca_OpcioneEstilo(sopcion As String, sCod_Almacen As String, scod_estcli As String, sdes_estcli As String, scod_ordpro As String, scod_present As String, sdes_present As String, sCod_Talla As String, txtCod As TextBox, txtDes As TextBox)
On Error GoTo fin
Dim rstAux As ADODB.Recordset
    
    If sCod_Almacen = "" Then
        MsgBox "seleccione un almacen", vbInformation + vbOKOnly, "Importante"
        Exit Sub
    End If

    strSQL = " SM_MUESTRA_ESTILO_CAMBIO_MODELO_TALLA '" & sopcion & "', '" & sCod_Almacen & "','" & scod_estcli & "','" & sdes_estcli & "','" & scod_ordpro & "','" & scod_present & "','" & sdes_present & "','" & sCod_Talla & "' "

    fila_seleccionada = 0

    With frmBusqGeneral
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos

        CODIGO = ".."
        Set rstAux = .gexList.ADORecordset
        'If rstAux.RecordCount > 1 Then
        .Show vbModal

        If fila_seleccionada > 0 And rstAux.RecordCount > 0 Then

            rstAux.AbsolutePosition = fila_seleccionada
            txtCod = Trim(rstAux!Cod)
            txtDes = Trim(rstAux!Descripcion)
            If sopcion = 2 Or sopcion = 3 Then
                txtOp.Text = rstAux!op
            End If
            If sopcion = 4 Or sopcion = 5 Then
                txtTalla.Text = rstAux!cod_talla
                txtStocksActual.Text = rstAux!STOCKS
                txtNewStocks.Text = rstAux!STOCKS + txtCantidad.Text
            End If
        Else
            txtCod = ""
            txtDes = ""
            SendKeys "{TAB}"
        End If

    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento (" & sopcion & ")"
End Sub

'''''***********************************realiza movimiento de salida y entrada de modelos****************************
Private Function GuardaDetalleCambioModeloTalla() As Boolean
On Error GoTo ErrDetMov
Dim sErr As String, cntAux As New ADODB.Connection, stit As String, _
    sNum_MovStksal As String, snum_movstkent   As String
Dim rstAux As New ADODB.Recordset
Dim scod_tipmovent As String
Dim scod_tipmovSal As String
Dim sCod_Almacen As String
Dim scod_barra As String

If indice = 0 Then
    scod_tipmovSal = "SMT"
    scod_tipmovent = "EMT"
Else
    scod_tipmovSal = "SDV"
    scod_tipmovent = "EDV"
End If
stit = "Guarda Cambios"
sCod_Almacen = Left(cboAlmacen, 2)

' EXEC UP_LG_MOVSTK 'I','69','','24/01/2015','SISTEMAS','','','E41','','','00003','  ','','','','001','0','','',''
' EXEC UP_LG_MOVSTK 'I','69','','24/01/2015','SISTEMAS','','','SV7','','','00001','  ','','DF','','001','0','','',''

GuardaDetalleCambioModeloTalla = False
    
    If GridEX1.RowCount = 0 Then
        MsgBox "Se debe especificar al menos un detalle", vbExclamation + vbOKCancel, "IMPORTANTE"
        Exit Function
    End If
    
    stit = "Guardar Detalle de cambio de modelo"
    
    cntAux.Open cConnect
    cntAux.BeginTrans

    '''Salida por cambio o devolucion
    strSQL = "EXEC UP_LG_MOVSTK 'I','" & sCod_Almacen & "','','" & Format(Now(), "DD/MM/YYYY") & "','" & vusu & "','','','" & scod_tipmovSal & "','','','00001','  ','','" & Trim(txtDocReferencia.Text) & "','','001','0','','',''"
    Set rstAux = cntAux.Execute(strSQL, adExecuteNoRecords)
    sNum_MovStksal = rstAux!num_movstk
    rstAux.Close
    
  Set rstAux = GridEX1.ADORecordset
  With rstAux
        If .RecordCount > 0 Then .MoveFirst
        Do Until .EOF
    
    '''DETALLE SALIDA POR CAMBIO O DEVOLUCION
           scod_barra = .Fields("COD_ORDPRO").Value + Right("0000" + CStr(Trim(.Fields("cod_present").Value)), 4) + Right("0000" + Trim(.Fields("cod_talla").Value), 4)
           
           strSQL = " EXEC LG_UP_MAN_LG_MOVISTK_ITEM_PRENDAS_VENTAS_DIRECTA '" & sCod_Almacen & "'," & _
                                                 "'" & sNum_MovStksal & "'," & _
                                                 "'PR000001'," & _
                                                 "'001'," & _
                                                 "'000001'," & _
                                                 "'" & Trim(.Fields("cod_estcli").Value) & "'," & _
                                                 "'" & Trim(.Fields("COD_ORDPRO").Value) & "'," & _
                                                 "'" & Trim(.Fields("COD_PRESENT").Value) & "'," & _
                                                 "'" & Trim(.Fields("COD_TALLA").Value) & "'," & _
                                                 "'" & scod_barra & "'," & _
                                                 "'" & Trim(.Fields("CANTIDAD").Value) & "'," & _
                                                 "'I','','','' "
                                                             
             cntAux.Execute strSQL, adExecuteNoRecords
            .MoveNext
      Loop
    End With
    
    '''CABECERA MOVIMIENTO DE ENTRADA POR CAMBIO O DEVOLUCION
    strSQL = "EXEC UP_LG_MOVSTK 'I','" & sCod_Almacen & "','','" & Format(Now(), "DD/MM/YYYY") & "','" & vusu & "','','','" & scod_tipmovent & "','','','00001','  ','','" & Trim(txtDocReferencia.Text) & "','','001','0','','',''"
    Set rstAux = cntAux.Execute(strSQL, adExecuteNoRecords)
    snum_movstkent = rstAux!num_movstk
    rstAux.Close

    Set rstAux = grxDatos.ADORecordset
    With rstAux
        If .RecordCount > 0 Then .MoveFirst
        Do Until .EOF
    
    '''DETALLE MOVIMIENTO DE SALIDA DE ALMACEN
        scod_barra = .Fields("COD_ORDPRO").Value + Right("0000" + CStr(Trim(.Fields("cod_present").Value)), 4) + Right("0000" + Trim(.Fields("cod_talla").Value), 4)
    
        strSQL = " EXEC LG_UP_MAN_LG_MOVISTK_ITEM_PRENDAS_VENTAS_DIRECTA '" & sCod_Almacen & "'," & _
                                                     "'" & snum_movstkent & "'," & _
                                                     "'PR000001'," & _
                                                     "'001'," & _
                                                     "'000001'," & _
                                                     "'" & Trim(.Fields("cod_estcli").Value) & "'," & _
                                                     "'" & Trim(.Fields("COD_ORDPRO").Value) & "'," & _
                                                     "'" & Trim(.Fields("COD_PRESENT").Value) & "'," & _
                                                     "'" & Trim(.Fields("COD_TALLA").Value) & "'," & _
                                                     "'" & scod_barra & "'," & _
                                                     "'" & Trim(.Fields("CANTIDAD").Value) & "'," & _
                                                     "'I','','','' "
                                                                 
                 cntAux.Execute strSQL, adExecuteNoRecords
                .MoveNext

        Loop
     End With
    
    '''RELACIONA MOVIMIENTOS
    strSQL = "LG_RELACIONA_MOV_CAMBIO_MODELO_TALLA '" & sCod_Almacen & "','" & sNum_MovStksal & "','" & snum_movstkent & "','" & vusu & "'"
    cntAux.Execute strSQL, adExecuteNoRecords
    
    Set GridEX1.ADORecordset = Nothing
    Set grxDatos.ADORecordset = Nothing
    
    cntAux.CommitTrans
    cntAux.Close
    Set cntAux = Nothing
    GuardaDetalleCambioModeloTalla = True
    Set GridEX1.ADORecordset = Nothing
    
Exit Function
ErrDetMov:
    GuardaDetalleCambioModeloTalla = False
    sErr = Err.Description
    cntAux.RollbackTrans
    cntAux.Close
    Set cntAux = Nothing
    MsgBox sErr, vbCritical + vbOKOnly, stit
End Function

Private Sub cmdRealizarCambios_Click()
   If GridEX1.RowCount <= 0 Then Exit Sub
    If GuardaDetalleCambioModeloTalla = True Then
        MsgBox "Se realizo los cambios de modelos con exito", vbInformation + vbOKOnly, "AVISO"
    Else
        MsgBox "Inconvenientes para realizar el cambio, favor de revisar los datos", vbInformation + vbOKOnly, "AVISO"
    End If

End Sub

