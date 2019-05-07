VERSION 5.00
Begin VB.Form Frm_Toolbar 
   Caption         =   "Form1"
   ClientHeight    =   570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3285
   LinkTopic       =   "Form1"
   ScaleHeight     =   570
   ScaleWidth      =   3285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImprimir 
      Height          =   315
      Left            =   0
      Picture         =   "Frm_Toolbar.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   345
   End
End
Attribute VB_Name = "Frm_Toolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function CambiarContenedor(ByRef oContenedor As Form)
    Dim retvalue
    SetParent Me.cmdImprimir.hwnd, oContenedor.hwnd
    Me.Caption = oContenedor.Caption
    'Mover
    Dim oCtrl As Control
    For Each oCtrl In oContenedor.Controls
        If TypeOf oCtrl Is GridEX20.GridEx Then
            If (TypeOf oCtrl.Container Is Frame And oCtrl.Container.Left < 1000) Or Not TypeOf oCtrl.Container Is Frame Then
                If oCtrl.Container.Name <> "" And Not TypeOf oCtrl.Container Is Form Then
                    cmdImprimir.Move oCtrl.Container.Left + oCtrl.Left, oCtrl.Container.Top + oCtrl.Top
                Else
                    Me.cmdImprimir.Move oCtrl.Left, oCtrl.Top
                End If
                Exit Function
            End If
        End If
    Next
    Unload Me
End Function

Public Sub cmdImprimir_Click()
On Error GoTo lblError
Dim oCtrl As Control
Dim oControl As Control
For Each oCtrl In Me.Controls ' oFrmContenedor.container.Controls
    For Each oControl In oCtrl.Container.Controls ' oFrmContenedor.container.Controls
        If TypeOf oControl Is GridEX20.GridEx Then
            If oControl.RowCount > 0 Then
                If oControl.Visible Then
                    Call Reporte(Me.Caption, oControl.ADORecordset)
                End If
            End If
        End If
    Next
Next
Exit Sub
lblError:
    MsgBox err.Description, vbCritical, "Mensaje del Sistema"
    Exit Sub
End Sub

Sub Reporte(ByVal sTitulo As String, ByVal adoRs As ADODB.Recordset)
On Error GoTo lblError
Dim oo As Object
    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open App.Path & "\RptLista.XLT"
    oo.Visible = True
    oo.DisplayAlerts = False
    oo.Run "Reporte", sTitulo, adoRs, vemp
    
    Exit Sub
lblError:
    Set oo = Nothing
    MsgBox err.Description
End Sub



