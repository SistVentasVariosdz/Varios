VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Frm_mantenimiento_series_Por_Almacen 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Series por Almacen"
   ClientHeight    =   5790
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   495
      Left            =   4080
      TabIndex        =   7
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton BtnEliminar 
      Caption         =   "Eliminar"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Registrar"
      Height          =   1695
      Left            =   0
      TabIndex        =   2
      Top             =   3360
      Width           =   8055
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nro Serie Guia:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Almacenes:"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
   Begin GridEX20.GridEX GrdLista1 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   4683
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GroupByBoxVisible=   0   'False
      BackColorBkg    =   12648384
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      FormatStylesCount=   7
      FormatStyle(1)  =   "frm_Mantenimiento_Series_Por_Almacen.frx":0000
      FormatStyle(2)  =   "frm_Mantenimiento_Series_Por_Almacen.frx":0138
      FormatStyle(3)  =   "frm_Mantenimiento_Series_Por_Almacen.frx":01E8
      FormatStyle(4)  =   "frm_Mantenimiento_Series_Por_Almacen.frx":029C
      FormatStyle(5)  =   "frm_Mantenimiento_Series_Por_Almacen.frx":0374
      FormatStyle(6)  =   "frm_Mantenimiento_Series_Por_Almacen.frx":042C
      FormatStyle(7)  =   "frm_Mantenimiento_Series_Por_Almacen.frx":050C
      ImageCount      =   0
      PrinterProperties=   "frm_Mantenimiento_Series_Por_Almacen.frx":052C
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
   Begin VB.CommandButton btnRegistrar 
      Caption         =   "Registrar"
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   5160
      Width           =   1335
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   5520
      Top             =   5160
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Seleccione un Almacen:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Frm_mantenimiento_series_Por_Almacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrConexion As String
Dim Rsx As New ADODB.Recordset
Dim RSy As New ADODB.Recordset
Dim RSz As New ADODB.Recordset
Dim cn_x As New ADODB.Connection
Dim cont As Integer
Dim Cod_almacen, cod_AlmacenX, Nro_Serie, txtnro As Integer
Dim nom_almacen, mant, opcX, Nro_SerieX, cadX, ax, bx, cx As String

Private Sub BtnEliminar_Click()
mant = "DEL"
If Nro_Serie <> 0 Then
    If (MsgBox("¿Desea eliminar la guia nro: " & Nro_Serie & " del almacen: " & nom_almacen & "?", 36, "Confirmar")) = 6 Then
        If RSz.State <> 0 Then RSz.Close
            RSz.Open "Exec LG_Mant_Series_x_Facturas '" & mant & "','" & opcX & "','" & Cod_almacen & "','" & Nro_Serie & "'", cn_x
            MsgBox "El Registro fue eliminado Satisfactoriamente", vbInformation, "Mensaje de Confirmacion"
            Cod_almacen = ""
            nom_almacen = ""
            Nro_Serie = 0
            Combo1_Click
            Exit Sub
    End If
Else
MsgBox "Seleccione el registro a eliminar", vbCritical, "Validacion"
Exit Sub
End If
End Sub

Private Sub btnRegistrar_Click()
On Error GoTo Errorx
    mant = "INS"
If MaskEdBox1.Text <> "000" And Trim(MaskEdBox1.Text) <> "" Then
    cod_AlmacenX = DataCombo1.BoundText
    Nro_SerieX = MaskEdBox1.Text
    If Nro_SerieX = "000" Then
        MsgBox "Ingrese un Numero de Serie Valido", vbExclamation, "Validacion"
    Else
        If DataCombo1.Text = "" Then
        MsgBox "Seleccione una Almacen", vbExclamation, "Mensaje"
        DataCombo1.SetFocus
        Else
        If RSz.State <> 0 Then RSz.Close
        RSz.Open "Exec LG_Mant_Series_x_Facturas '" & mant & "','" & opcX & "','" & cod_AlmacenX & "','" & Nro_SerieX & "'", cn_x, adOpenStatic, adLockReadOnly
        MsgBox "Nro de serie " & Nro_SerieX & " Grabada satisfactoriamente en el almacen " & DataCombo1.Text, vbInformation, "Mensaje de confirmacion"
        Combo1_Click
        Exit Sub
        End If
    End If
Else
    MsgBox "Debe ingresar un Numero de Serie", vbExclamation, "Mensaje"
    Exit Sub
End If
Errorx:
    MsgBox Err.Number & " ;" & Err.Description, 16, "Mensaje"
    Exit Sub
'Resume Next
End Sub

Private Sub Combo1_Click()
If Combo1.ListIndex = 0 Then
    opcX = "LG"
Else
    opcX = "CF"
End If
If Rsx.State <> 0 Then Rsx.Close
    Rsx.Open "Exec LG_Listado_almacenes '1','" & opcX & "'", cn_x
    Set GrdLista1.ADORecordset = Rsx
    GrdLista1.Columns(1).Visible = False
    GrdLista1.Columns(2).AutoSize
    GrdLista1.Columns(3).AutoSize
If RSy.State <> 0 Then RSy.Close
    DataCombo1.Text = ""
    RSy.Open "Exec LG_Listado_almacenes '2','" & opcX & "'", cn_x
    Set DataCombo1.RowSource = RSy
    DataCombo1.ListField = RSy.Fields(1).Name
    DataCombo1.BoundColumn = RSy.Fields(0).Name
End Sub

Private Sub Command1_Click()
If Rsx.State <> 0 Then Rsx.Close
If RSy.State <> 0 Then RSy.Close
If RSz.State <> 0 Then RSz.Close
Unload Me
End Sub

' Programacion Precotex

Private Sub Form_Load()
StrConexion = cConnect  '"Provider=sqloledb; Data Source=SERVERDATA;Initial Catalog=ONLY_STAR;integrated security = SSPI"
cn_x.Open StrConexion
Rsx.CursorLocation = adUseClient
RSy.CursorLocation = adUseClient
RSz.CursorLocation = adUseClient
Combo1.Clear
Combo1.AddItem "Almacen: LG"
Combo1.ItemData(Combo1.NewIndex) = 0
Combo1.AddItem "Almacen: CF"
Combo1.ItemData(Combo1.NewIndex) = 1
MaskEdBox1.Mask = "###"
GrdLista1.AllowEdit = False
MaskEdBox1.Text = "000"
End Sub


Private Sub GrdLista1_Click()
Cod_almacen = GrdLista1.Value(1)
nom_almacen = GrdLista1.Value(2)
Nro_Serie = GrdLista1.Value(3)
End Sub


Private Sub MaskEdBox1_GotFocus()
MaskEdBox1.SetFocus
'MaskEdBox1.SelText
With MaskEdBox1
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub MaskEdBox1_LostFocus()
If Len(Trim(MaskEdBox1.Text)) = 1 Then
    MaskEdBox1.Text = "00" & Trim(MaskEdBox1.Text)
End If
If Len(Trim(MaskEdBox1.Text)) = 2 Then
    MaskEdBox1.Text = "0" & Trim(MaskEdBox1.Text)
End If
If Len(Trim(MaskEdBox1.Text)) = 0 Then
    MaskEdBox1.Text = "000" & Trim(MaskEdBox1.Text)
End If
End Sub
