VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form Frm_Muestra_Procesos_Ex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Muestra Proceso De la Orden De Compra"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   9660
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   8160
      TabIndex        =   1
      Top             =   0
      Width           =   1455
      Begin FunctionsButtons.FunctButt FunctButt1 
         Height          =   2310
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   4075
         Custom          =   $"Frm_Muestra_Procesos_Ex.frx":0000
         Orientacion     =   1
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6800
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      FormatStylesCount=   7
      FormatStyle(1)  =   "Frm_Muestra_Procesos_Ex.frx":012C
      FormatStyle(2)  =   "Frm_Muestra_Procesos_Ex.frx":0264
      FormatStyle(3)  =   "Frm_Muestra_Procesos_Ex.frx":0314
      FormatStyle(4)  =   "Frm_Muestra_Procesos_Ex.frx":03C8
      FormatStyle(5)  =   "Frm_Muestra_Procesos_Ex.frx":04A0
      FormatStyle(6)  =   "Frm_Muestra_Procesos_Ex.frx":0558
      FormatStyle(7)  =   "Frm_Muestra_Procesos_Ex.frx":0638
      ImageCount      =   0
      PrinterProperties=   "Frm_Muestra_Procesos_Ex.frx":0658
   End
End
Attribute VB_Name = "Frm_Muestra_Procesos_Ex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SCod_Cliente_Tex As String
Public Ser_OrdComp     As String
Public Cod_OrdComp     As String
Public Sec_OrdComp     As String
Dim StrSql As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
    Case "ADICIONAR"
        Frm_Añadir_Procesos_Ex.saccion = "I"
        Frm_Añadir_Procesos_Ex.SCod_Cliente_Tex = SCod_Cliente_Tex
        
        StrSql = "select Nom_Cliente from tx_Cliente where Cod_Cliente_Tex='" & Trim(SCod_Cliente_Tex) & "'"
        Frm_Añadir_Procesos_Ex.Txt_Cliente = DevuelveCampo(StrSql, cConnect)
        
        
        Frm_Añadir_Procesos_Ex.Ser_OrdComp = Ser_OrdComp
        Frm_Añadir_Procesos_Ex.Cod_OrdComp = Cod_OrdComp
        Frm_Añadir_Procesos_Ex.Sec_OrdComp = Sec_OrdComp
        
        Frm_Añadir_Procesos_Ex.Txt_Sec = Sec_OrdComp
        Frm_Añadir_Procesos_Ex.txt_Serie = Ser_OrdComp
        Frm_Añadir_Procesos_Ex.Txt_Numero = Cod_OrdComp
        
        Frm_Añadir_Procesos_Ex.Show 1
        
        CARGA_GRID
        
        
        
    
    Case "MODIFICAR"
        Frm_Añadir_Procesos_Ex.saccion = "U"
        Frm_Añadir_Procesos_Ex.SCod_Cliente_Tex = SCod_Cliente_Tex
        StrSql = "select Nom_Cliente from tx_Cliente where Cod_Cliente_Tex='" & Trim(SCod_Cliente_Tex) & "'"
        Frm_Añadir_Procesos_Ex.Txt_Cliente = DevuelveCampo(StrSql, cConnect)
        
        
        Frm_Añadir_Procesos_Ex.Ser_OrdComp = Ser_OrdComp
        Frm_Añadir_Procesos_Ex.Cod_OrdComp = Cod_OrdComp
        Frm_Añadir_Procesos_Ex.Sec_OrdComp = Sec_OrdComp
        
        Frm_Añadir_Procesos_Ex.Txt_Sec = Sec_OrdComp
        Frm_Añadir_Procesos_Ex.txt_Serie = Ser_OrdComp
        Frm_Añadir_Procesos_Ex.Txt_Numero = Cod_OrdComp
        
        Frm_Añadir_Procesos_Ex.txtCod_Proceso_Tinto = GridEX1.Value(GridEX1.Columns("Cod_Proceso_Tinto").Index)
        Frm_Añadir_Procesos_Ex.txtDes_Proceso_Tinto = GridEX1.Value(GridEX1.Columns("Proceso").Index)
        Frm_Añadir_Procesos_Ex.Txt_Observaciones = GridEX1.Value(GridEX1.Columns("Observacion").Index)
        Frm_Añadir_Procesos_Ex.Show 1
        CARGA_GRID

    
    Case "ELIMINAR"
        If GridEX1.RowCount = 0 Then Exit Sub
        ELIMINAR
        CARGA_GRID
    
    Case "SALIR"
        Unload Me
    

End Select
End Sub


Public Sub CARGA_GRID()
On Error GoTo err_Carga

StrSql = "exec Ti_Sm_Muestra_Procesos_OrdenCompra  '" & SCod_Cliente_Tex & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Sec_OrdComp & "'"
  
Set GridEX1.ADORecordset = CargarRecordSetDesconectado(StrSql, cConnect)

GridEX1.Columns("Cliente").Width = 2000
GridEX1.Columns("Ser_OrdComp").Width = 800
GridEX1.Columns("Cod_OrdComp").Width = 800
GridEX1.Columns("Sec_OrdComp").Width = 800
GridEX1.Columns("Proceso").Width = 1500
GridEX1.Columns("Fec_Creacion").Width = 1000
GridEX1.Columns("Observacion").Width = 1000

GridEX1.Columns("Ser_OrdComp").Caption = "Serie"
GridEX1.Columns("Cod_OrdComp").Caption = "Numero"
GridEX1.Columns("Sec_OrdComp").Caption = "Secuencia"

GridEX1.Columns("Cod_Proceso_Tinto").Visible = False


Exit Sub
err_Carga:
    ErrorHandler Err, "CARGA_GRID"
End Sub


Private Sub ELIMINAR()
On Error GoTo Fin
Dim sTit As String
Dim saccion As String

saccion = "D"
   sTit = "Eliminar EL Proceso De la Orden De Compra"
    
   If MsgBox("Desea Eliminar ?", vbQuestion + vbYesNo, sTit) = vbNo Then Exit Sub
      
      
    StrSql = "EXEC Ti_Up_Procesos_Ordencompra 'D','" & SCod_Cliente_Tex & "','" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Sec_OrdComp & "','" & GridEX1.Value(GridEX1.Columns("Cod_Proceso_Tinto").Index) & "',''"
    
    ExecuteSQL cConnect, StrSql
    
    
Exit Sub
Fin:
    MsgBox Err.Description, vbCritical + vbOKOnly, sTit
End Sub




