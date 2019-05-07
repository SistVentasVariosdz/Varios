VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmVerRollosTejeduria 
   Caption         =   "Ver Rollos Tejeduria"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12975
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   12975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11280
      TabIndex        =   0
      Top             =   6600
      Width           =   1455
   End
   Begin GridEX20.GridEX grxDatos 
      Height          =   6555
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   11562
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GroupByBoxVisible=   0   'False
      HeaderFontName  =   "Verdana"
      HeaderFontBold  =   -1  'True
      HeaderFontSize  =   6.75
      HeaderFontWeight=   700
      FontName        =   "Tahoma"
      ColumnHeaderHeight=   270
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      FormatStylesCount=   9
      FormatStyle(1)  =   "FrmVerRollosTejeduria.frx":0000
      FormatStyle(2)  =   "FrmVerRollosTejeduria.frx":0128
      FormatStyle(3)  =   "FrmVerRollosTejeduria.frx":01D8
      FormatStyle(4)  =   "FrmVerRollosTejeduria.frx":028C
      FormatStyle(5)  =   "FrmVerRollosTejeduria.frx":0364
      FormatStyle(6)  =   "FrmVerRollosTejeduria.frx":041C
      FormatStyle(7)  =   "FrmVerRollosTejeduria.frx":04FC
      FormatStyle(8)  =   "FrmVerRollosTejeduria.frx":05A8
      FormatStyle(9)  =   "FrmVerRollosTejeduria.frx":0674
      ImageCount      =   0
      PrinterProperties=   "FrmVerRollosTejeduria.frx":0740
   End
End
Attribute VB_Name = "FrmVerRollosTejeduria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public scod_almacen As String
Public snum_movstk  As String
Public ordtra_tinto As String
Dim StrSQL As String
Public Sub muestrarollos()
On Error GoTo Salvar_DatosErr
    
   StrSQL = " EXEC Tj_SM_MUESTRA_MOV_TELA_CRUDA_ROLLOS '" & scod_almacen & "', '" & snum_movstk & "',''"
    Set grxDatos.ADORecordset = CargarRecordSetDesconectado(StrSQL, cConnect)
    Call Configurar_Grid
Exit Sub
Salvar_DatosErr:
ErrorHandler err, "Salvar_Datos"

End Sub
Public Sub Configurar_Grid()
    Dim C As Integer
    Dim colTemp As JSColumn
    Dim fmtCon  As JSFmtCondition
    
    With grxDatos
    
        For C = 1 To .Columns.Count
            With .Columns(C)
                .Caption = UCase(.Caption)
                .HeaderAlignment = jgexAlignCenter
                .TextAlignment = jgexAlignCenter
                .Visible = False
            End With
        Next C

        
        With .Columns("cod_almacen")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "COD"
        End With

        With .Columns("num_movstk")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "num_movstk"
        End With

        With .Columns("num_secuencia")
            .Visible = True
            .Width = 800
            .TextAlignment = jgexAlignLeft
            .Caption = "sec"
        End With

        With .Columns("cod_ordtra")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "OT"
        End With

        With .Columns("NUM_SECUENCIA_OT")
            .Visible = True
            .Width = 800
            .TextAlignment = jgexAlignLeft
            .Caption = "Sec Ot"
        End With

        With .Columns("num_rollo")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "Num Rollo"
        End With
        With .Columns("prefijo_maquina")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "Pre-Maquina"
        End With
        With .Columns("codigo_rollo")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "cod-Rollo"
        End With

        With .Columns("kgs_rollo")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "kgs Rollo"
        End With

        With .Columns("cod_calidad")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "Calidad"
        End With

        With .Columns("cod_calidad_auditoria")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "Calidad Aud."
        End With

        With .Columns("fec_Auditoria")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "Fec Aud."
        End With

        With .Columns("auditor")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "Fec Aud."
        End With

        With .Columns("observacion")
            .Visible = True
            .Width = 1000
            .TextAlignment = jgexAlignLeft
            .Caption = "Fec Aud."
        End With
        
    Dim oGroup01 As GridEX20.JSGroup
    Dim oGroup02 As GridEX20.JSGroup
    Dim valorcant   As JSColumn
    Dim valorStock   As JSColumn
    
      With grxDatos
            
        Set oGroup01 = .Groups.Add(.Columns("tela").Index, jgexSortAscending)
        .DefaultGroupMode = jgexDGMExpanded
        .BackColorRowGroup = RGB(239, 235, 222)
           
           .GroupFooterStyle = jgexTotalsGroupFooter
           Set valorcant = .Columns("kgs_rollo")
           
           With valorcant
               .AggregateFunction = jgexSum
               .TotalRowPrefix = "Total: "
               .TextAlignment = jgexAlignRight
           End With

        End With
    End With
    
    Call colorGrupo
    
End Sub

Private Sub grxDatos_RowFormat(RowBuffer As GridEX20.JSRowData)

Dim strGroupCaption As String

If grxDatos.RowCount = 0 Then Exit Sub

If RowBuffer.RowType = jgexRowTypeGroupHeader Then
    strGroupCaption = RTrim(RowBuffer.GroupCaption) & " (" & RowBuffer.RecordCount & " Rollos " & "" & ") "
    RowBuffer.GroupCaption = strGroupCaption
End If
'Call colorGrupo

'Dim fmtConDIA_Programado As JSFmtCondition
'If Buscando = 1 Then
'    Set fmtConDIA_Programado = GridEX1.FmtConditions.Add(GridEX1.Columns("MONTODESPACHO").Index, jgexEqual, 0)
'
'    With fmtConDIA_Programado.FormatStyle
'        .ForeColor = &H8000&
'        .FontSize = 8
'        .BackColor = &H80000018 'vbYellow
'    End With
'End If

End Sub

Private Sub colorGrupo()

Dim fmtCon As JSFmtCondition

Set fmtCon = grxDatos.FmtConditions.Add(grxDatos.Columns("tela").Index, jgexGreaterThan, 1)

With grxDatos.FmtConditions
        .ApplyGroupCondition = True
        .ShowGroupConditionCount = True
        .GroupConditionCountTitle = "Rollos"
        Set fmtCon = .GroupCondition
End With

fmtCon.SetCondition grxDatos.Columns("tela").Index, jgexGreaterThan, 1
fmtCon.FormatStyle.FontBold = True
fmtCon.FormatStyle.BackColor = &HFFFFC0   '&HC0FFC0    ' &HC0E0FF    ' '&HC0FFFF

End Sub


Private Sub cmdCancelar_Click()
Unload Me
End Sub


