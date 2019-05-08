Attribute VB_Name = "ModConexion"


Public Declare Function Captura Lib "CommLib32.dll" () As Long

Public CadConn As New ADODB.Connection
Global cConnect As String

Option Explicit
Option Base 0
Dim sQuery As String


Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Const CB_FINDSTRING = &H14C
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

'Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Const BlockSize = 100000 'This size can be experimented with for
                         'performance and reliability.


Global Ruta_Logo_Empresa As String
Global Num_Ruc_Empresa As String
Global Direccion_Empresa As String
Global DSN_Empresa As String
Global DSN_Seguridad As String
Global Ruta0_Empresa As String
Global Fecha_Hora_Conexion As String

Public vusu As String
Public vemp As String
Public vper As String
Public vemp1 As String
Public vpas As String
Public vNomFor As String
Public vRuta As String

Public iLanguage As Integer
Public oMDIParent As Object

Public cn As ADODB.Connection
Public sDllName As String
Public oFormObjDLL As Object
Public objFormDLL As Object

' Constantes de la Aplicación
Public Const M_NUEVO As String = "NUEVO"
Public Const M_EDITAR As String = "EDITAR"
Public Const M_ELIMINAR As String = "ELIMINAR"

' Tipo de Dato Message
Type RegMessage
    StringParm As String
    Cancel As Boolean
End Type

' Variable Publica
Public Message As RegMessage


' Registro Detalle Pedido
Type RegDetallePedido
    IdArticulo As String * 8
    Nombre As String * 35
    precio As Currency
    cantidad As Integer
    Cancel As Boolean
    Accion As String * 6
End Type


Public Enum FIELD
  COL_NAME
  COL_TYPE
  COL_DESCRIPLARGA
  COL_DESCRIPTION
  COL_LENGTH
  COL_MIN
  COL_MAX
  COL_DEFAULT
  COL_DES_CORTA
  COL_DES_ABREVIADA
End Enum

Public Enum DataTypeEnum
 adBigInt = 20
 adBinary = 128
 adBoolean = 11
 adBSTR = 8
 adChar = 129
 adCurrency = 6
 adDate = 7
 adDBDate = 133
 adDBTime = 134
 adDBTimeStamp = 135
 adDecimal = 14
 adDouble = 5
 adEmpty = 0
 adError = 10
 adGUID = 72
 adIDispatch = 9
 adInteger = 3
 adIUnknown = 13
 adLongVarBinary = 205
 adLongVarChar = 201
 adLongVarWChar = 203
 adNumeric = 131
 adSingle = 4
 adSmallInt = 2
 adTinyInt = 16
 adUnsignedBigInt = 21
 adUnsignedInt = 19
 adUnsignedSmallInt = 18
 adUnsignedTinyInt = 17
 adUserDefined = 132
 adVarBinary = 204
 adVarChar = 200
 adVariant = 12
 adVarWChar = 202
 adWChar = 130
End Enum

Public Enum FieldEnum
  IName = 0
  iActualSize = 1
  iAttributes = 2
  iDefinedSize = 3
  iNumericScale = 4
  iOriginalValue = 5
  iPrecision = 6
  iType = 7
  iUnderlyingValue = 8
  iValue = 9
  iMaxEnumField = 9
End Enum


' Variable Detalle Pedido
Public DetPedido As RegDetallePedido, cServidor As String

Public cSEGURIDAD As String

Public bCargaConexion As Boolean

Public Const cCLASS_TG_PURORD As String = "Visuales.clsTG_PurOrd"
    
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpbuffer As String, ByVal nSize As Long) As Long
Declare Function GetcomputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long

Public Const kCHR_BOLD_IN As String = "E"
Public Const kCHR_BOLD_OUT As String = "F"

Public Const kKEY_SEPARATOR As String = "-+-"
'Estructura de cada elemento del arreglo para los mensajes
Public Type Message
    Tipo As TypeMsg
    Code As CodeMsg
    Description As String
    HelpID As Integer
    Tag As Variant
End Type
'Ubound del arreglo (mayor codemsg  -1)
Public Const kMESSAGE_COUNT = 123
Public aMessage(kMESSAGE_COUNT) As Message

'Constantes para Tipo de Mensajes
Public Enum TypeMsg
    kTYPEMSG_INFORMATION = 1
    kTYPEMSG_QUESTION = 2
    kTYPEMSG_WARNING = 3
    kTYPEMSG_ERROR = 4
    kTYPEMSG_FIELD = 5
End Enum

'Constantes publicas para Codigo de Mensajes
Public Enum CodeMsg
    kMESSAGE_ERR_NOTEMPTY = 401
    kMESSAGE_ERR_FOUND = 402
    kMESSAGE_ERR_NOTFOUND = 403
    kMESSAGE_ERR_USERCONNECTFAIL = 405
    kMESSAGE_ERR_CODIGO_YA_REGISTRADO = 406
    kMESSAGE_ERR_HA_OCURRIDO_IMPREVISTO = 407
    kMESSAGE_ERR_REGISTRO_TIENE_TRANSAC_RELACIONADAS = 408
    kMESSAGE_ERR_FILE_NOT_FOUND = 409
    kMESSAGE_ERR_PROCESS_INSATISFACT = 410
    kMESSAGE_ERR_LOTEST_CLOSED = 456
    kMESSAGE_ERR_STYLE_HAVE_MORE_ESTPRO = 458
    kMESSAGE_INF_PROCESS_SATISFACTO = 102
    kMESSAGE_INF_NO_INIT_SEARCH = 103
    kMESSAGE_INF_DATA_NOTFOUND = 411
    kMESSAGE_INF_FILE_PRINT_OK = 104
    kMESSAGE_INF_NEW_CODIGO = 106
    kMESSAGE_INF_DATA_SAVE = 107
    kMESSAGE_INF_DATA_DELETE = 108
    
    kMESSAGE_WAR_ENABLED_DELETED = 301
    kMESSAGE_WAR_CONFIR_CHANGES = 302
    
    kMESSAGE_ASK_PRINT_FILE = 201
    kMESSAGE_ASK_EXIT_SYSTEM = 202
    kMESSAGE_ASK_PROCESS = 203
    kMESSAGE_ASK_MAILING_FILE = 204
    kMESSAGE_ASK_DELETE_PURORD = 207
    kMESSAGE_ASK_DELETE_LOTEST = 208
    
    
    
    kMESSAGE_ERR_VALIDA_COD_CLIENTE = 413
    kMESSAGE_ERR_VALIDA_NOM_CLIENTE = 414
    kMESSAGE_ERR_VALIDA_DES_CLIENTE = 415
    kMESSAGE_ERR_VALIDA_DES_DIVISION = 416
    kMESSAGE_ERR_VALIDA_COD_DIVISION = 417
    kMESSAGE_ERR_VALIDA_DES_COLOR = 418
    kMESSAGE_ERR_VALIDA_COD_COLOR = 419
    kMESSAGE_ERR_VALIDA_COD_ESTCLI = 420
    kMESSAGE_ERR_VALIDA_NOM_ESTCLI = 421
    kMESSAGE_ERR_VALIDA_TIP_ESTCLI = 422
    kMESSAGE_ASK_NUEVO_ESTCLI = 205
    kMESSAGE_ASK_NUEVO_PURORD = 206
    
    kMESSAGE_ERR_VALIDA_COD_COMI = 423
    kMESSAGE_ERR_VALIDA_DES_COMI = 424
    
    kMESSAGE_ERR_VALIDA_COD_DESTINO = 425
    kMESSAGE_ERR_VALIDA_DES_DESTINO = 426
    kMESSAGE_ERR_VALIDA_COD_FABRICA = 427
    kMESSAGE_ERR_VALIDA_ABR_FABRICA = 428
    kMESSAGE_ERR_VALIDA_NOM_FABRICA = 429
    kMESSAGE_ERR_VALIDA_DES_FABRICA = 430
    kMESSAGE_ERR_VALIDA_DIR_FABRICA = 431
    kMESSAGE_ERR_VALIDA_TEL_FABRICA = 432
    
    kMESSAGE_ERR_VALIDA_COD_MONEDA = 433
    kMESSAGE_ERR_VALIDA_DES_MONEDA = 434
    
    kMESSAGE_ERR_VALIDA_COD_ORGANIZ = 435
    kMESSAGE_ERR_VALIDA_DES_ORGANIZ = 436
    
    kMESSAGE_ERR_VALIDA_COD_PAGO = 437
    kMESSAGE_ERR_VALIDA_DES_PAGO = 438
    
    kMESSAGE_ERR_VALIDA_COD_FABRES = 439
    kMESSAGE_ERR_VALIDA_COD_CLIRES = 440
    kMESSAGE_ERR_VALIDA_RES_CLIENTE = 441
    
    kMESSAGE_ERR_VALIDA_COD_TEMP = 442
    kMESSAGE_ERR_VALIDA_DES_TEMP = 443
     
    
    kMESSAGE_ERR_VALIDA_PORC_CLIENTE = 444
    
    kMESSAGE_ERR_VALIDA_COD_TIPEMB = 445
    kMESSAGE_ERR_VALIDA_DES_TIPEMB = 446
    
    kMESSAGE_ERR_VALIDA_COD_TIPPRE = 447
    kMESSAGE_ERR_VALIDA_DES_TIPPRE = 448
    
    kMESSAGE_ERR_VALIDA_COD_UM = 449
    kMESSAGE_ERR_VALIDA_DES_UM = 450
    
    kMESSAGE_ERR_VALIDA_ANO_MES = 451
    
    kMESSAGE_ERR_VALIDA_SERIE = 452
    
    kMESSAGE_ERR_VALIDA_FACTURA = 453
    kMESSAGE_ERR_ASIGN_STYLE_TEMCLI = 457
    kMESSAGE_ERR_EXIST_CLIENT = 459
    kMESSAGE_ERR_NOT_RIGHT_OPTION = 460
    kMESSAGE_ERR_INVALID_SELECC = 461

End Enum

Public Enum TypeMante
    kMANT_ADICIONAR = 0
    kMANT_MODIFICAR = 1
    kMANT_ELIMINAR = 2
    kMANT_CONSULTAR = 3
    kMANT_BUSCAR = 4
    kMANT_IMPRIMIR = 5
    kMANT_SALIR = 6
End Enum

Enum FIELDx
  COL_NAME
  COL_TYPE
  COL_DESCRIPLARGA
  COL_DESCRIPTION
  COL_LENGTH
  COL_MIN
  COL_MAX
  COL_DEFAULT
  COL_DES_CORTA
  COL_DES_ABREVIADA
End Enum
Public Const kMESSAGE_PROVEEDOR_PASADOS As String = ""
Public Const kMESSAGE_PROVEEDOR_PENDIENTES As String = ""

Public Function get_botones1(ByVal f As Form, ByVal Vcod_perfil As Variant, ByVal vcod_empresa As Variant, ByVal fname As Variant)
On Error GoTo hand
Dim Rs1 As ADODB.Recordset
Set Rs1 = New ADODB.Recordset
Dim sQuery As String
sQuery = "Sp_funciones3 '" & Vcod_perfil & "','" & vcod_empresa & "','" & fname & "'"
'RS1.ActiveConnection = cSEGURIDAD
Rs1.CursorType = adOpenStatic
Rs1.Open sQuery, cSEGURIDAD
Dim Scad As String
Dim ICONT As Integer
If Not (Rs1.BOF And Rs1.EOF) Then
    'For j = 1 To rs1.RecordCount
           ICONT = 1
           While Not Rs1.EOF
            If ICONT = 1 Then
                Scad = LTrim(RTrim(Rs1!nom_corto))
            Else
                Scad = Scad + "/" + LTrim(RTrim(Rs1!nom_corto))
            End If
            Rs1.MoveNext
            ICONT = ICONT + 1
           Wend
         'Boton_Enabled rs1!nom_corto, f
End If
get_botones1 = Scad
Exit Function
hand:
ErrorHandler Err, "get_botones1"
End Function

Public Sub errores(sCodigo As Long)
Dim oCode As CodeMsg
Dim oMessage As clsMessages
Dim sMess As String
Dim iPos As Integer

    Select Case sCodigo
        Case "9999"
            oCode = kMESSAGE_ERR_CODIGO_YA_REGISTRADO
            Set oMessage = New clsMessages
            oMessage.Codigo = oCode
            Call oMessage.ShowMesage(iLanguage)
        Case Else
            sMess = Err.Description
            iPos = InStr(1, sMess, "SERVER]", 1)
            If iPos > 0 Then
                sMess = Mid(sMess, iPos + 7)
            End If
            oCode = kMESSAGE_ERR_HA_OCURRIDO_IMPREVISTO
            Set oMessage = New clsMessages
            oMessage.Codigo = oCode
            oMessage.OptionalText = Chr(13) & sMess
            
            Call oMessage.ShowMesage(iLanguage)

    End Select

Set oMessage = Nothing
End Sub


Public Sub Aviso(Mensaje As String, Tipo As Integer)
    Select Case Tipo
        Case 1
            MsgBox Mensaje, vbExclamation, "Aviso"
        Case 2
            MsgBox Mensaje, vbInformation + vbMsgBoxRight, "Mensaje"
        Case 3
            MsgBox Mensaje, vbCritical, "Error Grave"
    End Select
End Sub


Public Function GetRecordset(ByVal Connect As String, ByVal sql As String) As Object 'ADOR.Recordset
  On Error GoTo ehGetRecordset
  Dim objADORs As Object ' New ADODB.RecordSet '
  Dim objAdoCn As Object ' New ADODB.Connection '
  
 
  Set objADORs = CreateObject("ADODB.Recordset") 'New ADODB.RecordSet '
  Set objAdoCn = CreateObject("ADODB.Connection") ' New ADODB.Connection  '
  objAdoCn.CursorLocation = 3
  objAdoCn.Open Connect
  objAdoCn.CommandTimeout = 900
  
  
  objADORs.Open sql, objAdoCn, 3, 4 ', 4  'adOpenStatic= 3 ,  adLockBatchOptimistic = 4  (orignal)  'cambio desde 24/07/2000 ' 1 adLockReadOnly , ' 4 adCmdStoredProc
  Set GetRecordset = objADORs
  Set GetRecordset.ActiveConnection = objAdoCn
  Set objADORs.ActiveConnection = Nothing
  objAdoCn.Close
  Set objAdoCn = Nothing
 
Exit Function
ehGetRecordset:
  Err.Raise Err.Number, Err.Source, Err.Description
  MsgBox Err.Description
End Function
Public Sub InitMessages()

End Sub

