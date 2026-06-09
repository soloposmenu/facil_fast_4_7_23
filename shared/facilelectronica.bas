Attribute VB_Name = "facilelectronica"

'---------------------------------------------------------------------------------------
' Module    : facilelectronica
' Author    : hsequeira
' Date      : 17/03/2024
' Purpose   : TODO LO NECESARIO PARA CONECTARSE AL API DE FACTURA FACIL
'---------------------------------------------------------------------------------------

Public mi_urlWebservice As String
Public mi_tokenEmpresa As String
Public mi_tokenPassword As String
Public mi_codigoSucursalEmisor As String


Public FE_Facil_ONLINE As Boolean      'INDICA SI LA FACTURA ELECTRONICA FUE ACEPTADA POR EL API

Public X_FF_Company As String
Public X_FF_API_Key As String
Public X_FF_Branch As String
Public mi_Ambiente As Integer


Public FE_QRData As String
Public FE_pdfURL As String
Public FE_CUFE As String
Public mensajesType As String
Public mensajesCode As String
Public mensajesMsg As String

Public FE_process_date As String                    'FECHA DEL PROTOCOLO
Public FE_digest_value As String                     '
Public FE_authorization_number As String        'PROTOCOLO DE AUTORIZACION
Public o_msg As String                                  'MSG para el cliente, viene dentro del SOAP OFF LINE


'VARIABLES CLIENTE.json
Public tipoClienteFE As String
'01:Contribuyente. 02:Consumidor final.
'03:Gobierno. 04:Extranjero.

Public tipoContribuyente As Integer
'1:Natural. 2:Jurídico
Public numeroRUC As String  '0-0-0
Public DVRUC  As String      ' ""
Public pais  As String          'SIEMPRE PA
Public razonSocial  As String
Public direccion  As String
Public provincia  As String
Public distrito As String
Public corregimiento As String
Public telefono1 As String
Public correoElectronico1 As String
Public telefono2 As String
Public correoElectronico2 As String
Public codigoUbicacion  As String      '":"1-1-1"
'INFO": 10MAR2024
Public paisExtranjero As String

'INFO: UPDATE 21JUL2024
'SE VALIDA EN GENERATEJSON
Public AllowFE_to_Continue As Boolean   'SI ES UN FALLO GRAVE NO DEJA CONTINUAR, SI ES UN FALLO ADMISIBLE, LE PERMITE CONTINUAR



Private rsFETrans As ADODB.Recordset
Private rsFEPagos As ADODB.Recordset
Private rsEPago As ADODB.Recordset
Private rsFEPropinas As ADODB.Recordset

Public mi_App_Name As String
Public mi_Panta As String     'PARA IDENTIFICAR SI VIENE DE PLU o PAGOS
'Private mi_Num_Factura As Long
Private local_total_factura As Single

Public cFE_LastErrorCode As Long


'~~~~~~~~~~~~~~~~~~~~~~~~~~~
'17DIC2024
Public o_mi_urlWebservice As String
Public o_urlWebservice As String
Public o_X_FF_Company As String
Public o_X_FF_API_Key As String
Public o_X_FF_Branch As String      'SUCURSAL 13JUL2025
Public o_mi_Ambiente As Integer
Public o_poscode As String            'CAJA 13JUL2025
'~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public lineWidth As Integer
Public cLogMode As String

Public nIDClienteBusqueda As Long


'Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Function GetFacilRegistryValues() As Boolean
Dim step As String

On Error GoTo GetFacilRegistryValues_Error

    step = "mi_urlWebservice "
    mi_urlWebservice = RegRead("HKCU\Software\SoloPosMenu\FE\FACIFACTURA\urlWebservice")
    
    step = "X_FF_Company "
    X_FF_Company = RegRead("HKCU\Software\SoloPosMenu\FE\FACIFACTURA\company")
    
    step = "X_FF_API_Key "
    X_FF_API_Key = RegRead("HKCU\Software\SoloPosMenu\FE\FACIFACTURA\api")
    
    step = "X_FF_Branch "
    X_FF_Branch = RegRead("HKCU\Software\SoloPosMenu\FE\FACIFACTURA\branch")
    
    step = "mi_Ambiente "
    mi_Ambiente = Int(RegRead("HKCU\Software\SoloPosMenu\FE\FACIFACTURA\ambiente"))
    
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~
    step = "o_mi_urlWebservice"
    o_mi_urlWebservice = RegRead("HKCU\Software\SoloPosMenu\FE\FACIFACTURA\o_urlWebservice")
    
    o_X_FF_Company = RegRead("HKCU\Software\SoloPosMenu\FE\FACIFACTURA\o_company")
    o_X_FF_API_Key = RegRead("HKCU\Software\SoloPosMenu\FE\FACIFACTURA\o_api")
    o_X_FF_Branch = RegRead("HKCU\Software\SoloPosMenu\FE\FACIFACTURA\o_branch")
    step = "o_mi_Ambiente "
    o_mi_Ambiente = Int(RegRead("HKCU\Software\SoloPosMenu\FE\FACIFACTURA\o_ambiente"))
    
    step = "o_poscode "
    o_poscode = RegRead("HKCU\Software\SoloPosMenu\FE\FACIFACTURA\o_poscod")
    
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error GoTo 0
Exit Function

GetFacilRegistryValues_Error:

    ShowMsg step & vbCrLf & "Error en GetFacilRegistryValues() " & vbCrLf & Err.Number & ": " & Err.description, vbYellow, vbRed

End Function

Public Function GenerateJson(nEnvironment As Integer, nFactura As Long, RSPAGOS As ADODB.Recordset, _
                                        lstPagos As MSHFlexGrid, Exonerado As Boolean, mi_total_factura As String, _
                                        Optional rsParamPropina As ADODB.Recordset) As Boolean
    Dim JsonString As String
    Dim header As String
    Dim receptor As String
    Dim items As String
    Dim payments As String
    Dim docType As String
    Dim INFO As String
'    Dim tipoEmpresa As String
'    Dim tipoPersona As String
    'Dim cPais As String
    Dim cClienteDir As String
    Dim future_ObtenerPagado As Single
    
        
    Dim nFreefile As Byte
    Dim a As String
    
   On Error GoTo GenerateJson_Error

    AllowFE_to_Continue = False

    '|||||||||||||||||||||||||||||||||||||
    Call Load_FE_Mesa
    '|||||||||||||||||||||||||||||||||||||
    
    If rsFETrans.RecordCount = 0 Then
        AllowFE_to_Continue = True
        GenerateJson = False
        Exit Function
    End If
    
    If tipoClienteFE = "" Then tipoClienteFE = "02"
    If tipoContribuyente = 0 Then tipoContribuyente = 1
    
    ' Set values for each section
'    tipoClienteFE = Left(lstTipoEmpresa.Text, 2)
'    tipoContribuyente = LTrim(RTrim(Str(Left(lstNat_Jurid, 1))))
    'cPais = paisExtranjero
    '||||||||||||||||||||||||||||||||||||||||||||||||||||||
    header = GenerateHeader(1, nEnvironment)
    '||||||||||||||||||||||||||||||||||||||||||||||||||||||
    'xxxreceptor = GenerateReceptor("02", "CONSUMIDOR FINAL", 1, "Dirección Cliente", "cliente@correo.com", 123123123)
    'receptor = GenerateReceptor("02", "CONSUMIDOR FINAL", 1, "", "", """0-0-0""")
    'receptor = GenerateReceptor("02", "HUMBERTO SEQUEIRA", 1, "VISTA DEL VALLE", "ventas@soloposmenu.com", """N-16-882""")
    'receptor = GenerateReceptor("02", "HUMBERTO SEQUEIRA", 1, "VISTA DEL VALLE", "ventas@soloposmenu.com", "N-16-882", "73")
    'receptor = GenerateReceptor(tipoEmpresa, "HUMBERTO SEQUEIRA", tipoPersona, "VISTA DEL VALLE", "ventas@soloposmenu.com", "N-16-882", "73")
    cClienteDir = corregimiento & "," & distrito & "," & provincia
    If cClienteDir = ",," Then cClienteDir = "ESCRIBIR DIRECCION"
    If razonSocial = "" Then razonSocial = "CONSUMIDOR FINAL"
    If numeroRUC = "" Then numeroRUC = "0-0-0"
    If correoElectronico1 = "" Then correoElectronico1 = "eMail"
    '||||||||||||||||||||||||||||||||||||||||||||||||||||||
    receptor = GenerateReceptor(tipoClienteFE, razonSocial, LTrim(RTrim(Str(tipoContribuyente))), cClienteDir, correoElectronico1, numeroRUC, DVRUC, paisExtranjero)
    '||||||||||||||||||||||||||||||||||||||||||||||||||||||
    
    'local_total_factura = CSng(mi_total_factura)
    '||||||||||||||||||||||||||||||||||||||||||||||||||||||
    items = GenerateItems(LTrim(RTrim(Str(tipoContribuyente))), Exonerado)
    '||||||||||||||||||||||||||||||||||||||||||||||||||||||
    'local_total_factura = Format(local_total_factura, "#.00")
    
    Set rsFEPagos = RSPAGOS.Clone(adLockOptimistic)
    
    'mi_total_factura = local_total_factura
    
    '|||||||||||||||||||||||||||||||||||||
    'INFO: UPDATE 9JUL2024
    '|||||||||||||||||||||||||||||||||||||
    mi_total_factura = Format(local_total_factura, "#.00")
    'payments = GeneratePayments(lstPagos, future_ObtenerPagado)
    
    'INFO: FIX RSPAGOS EN VEZ DE LISTPAGOS
    '10SEP2024
    If Left(mi_App_Name, 7) = "SoloMix" Then
        Rem payments = GeneratePayments(lstPagos, future_ObtenerPagado)
        '||||||||||||||||||||||||||||||||||||||||||||||||||||||
        payments = Generate_NEW_Payments(RSPAGOS, future_ObtenerPagado)
        '||||||||||||||||||||||||||||||||||||||||||||||||||||||
    Else
        '||||||||||||||||||||||||||||||||||||||||||||||||||||||
        payments = Generate_NEW_Payments(RSPAGOS, future_ObtenerPagado)
        '||||||||||||||||||||||||||||||||||||||||||||||||||||||
    End If
    
    docType = "01"
    'mi_total_factura = LTrim(RTrim(mi_total_factura))
    'mi_total_factura = LTrim(RTrim(CStr(local_total_factura)))
    mi_total_factura = LTrim(RTrim(CStr(mi_total_factura)))
    
    If Left(mi_App_Name, 7) = "SoloMix" Then
        On Error Resume Next
        rsParamPropina.MoveFirst
        On Error GoTo 0
    
        Rem If rsFEPropinas.EOF Then
        If rsParamPropina.EOF Then
            INFO = "<string>"
        Else
            Do While Not rsParamPropina.EOF
                INFO = INFO & rsParamPropina!DESCRIP_1 & Space(2) & Format(rsParamPropina!MONTO_1, "STANDARD") & " // "
                rsParamPropina.MoveNext
            Loop
        End If
    Else
        INFO = "<string>"
    End If
    
    If Exonerado Then
        INFO = INFO & "CLIENTE EXONERADO //"
    End If
    
    ' Build the complete JSON
    'jsonString = "{" & header & ",""document"": {""fd_number"": 169," & receptor & ",""items"": " & items & ",""payments"": " & payments & ",""type"": """ & docType & """,""info"": """ & info & """}}"
    'jsonString = "{" & header & ",""document"": {""fd_number"": " & nFactura & "," & receptor & ",""items"": " & items & ",""payments"": " & payments & ",""type"": """ & docType & """,""info"": """ & info & """}}"
    'INCLUIR future_ObtenerPagado
    'jsonString = "{" & header & ",""document"": {""fd_number"": " & nFactura & "," & receptor & ",""items"": " & items & ",""payments"": " & payments & ",""type"": """ & docType & """,""info"": """ & info & """}}"
    
    'INCLUIR mi_total_factura
    'mi_Num_Factura = nFactura
    'jsonString = "{" & header & ",""document"": {""fd_number"": " & nFactura & "," & receptor & ",""items"": " & items & ",""payments"": " & payments & ",""total"":""" & mi_total_factura & """,""type"": """ & docType & """,""info"": """ & info & """}}"
    Rem JsonString = "{" & header & ",""document"": {""fd_number"": " & nFactura & "," & receptor & ",""items"": " & items & ",""payments"": " & payments & ",""total"":""" & mi_total_factura & """,""type"": """ & docType & """,""info"": """ & INFO & """}}"
    'INFO: 12MARZO2026
    Rem JsonString = "{" & header & ",""document"": {""fd_number"": " & nFactura & ",""cajero"": " & nCajero & "," & receptor & ",""items"": " & items & ",""payments"": " & payments & ",""total"":""" & mi_total_factura & """,""type"": """ & docType & """,""info"": """ & INFO & """}}"
    Rem JsonString = "{" & header & ",""document"": {""fd_number"": " & nFactura & ",""Cajero"": " & npNumCaj & ",""Mesero"": " & nMesero & "," & receptor & ",""items"": " & items & ",""payments"": " & payments & ",""total"":""" & mi_total_factura & """,""type"": """ & docType & """,""info"": """ & INFO & """}}"
    JsonString = "{" & header & ",""document"": {""fd_number"": " & nFactura & ",""Cajero"": " & npNumCaj & ",""Mesero"": " & nMesero & ",""CajaID"":  """ & rs00!DESCRIP & """ ,""CajaDir"":  """ & rs00!direccion & """," & receptor & ",""items"": " & items & ",""payments"": " & payments & ",""total"":""" & mi_total_factura & """,""type"": """ & docType & """,""info"": """ & INFO & """}}"
    
'    Debug.Print header
'    Debug.Print receptor
'    Debug.Print items
'    Debug.Print payments
'    ' Display the generated JSON
    ' Debug.Print jsonString
    
    'lstLog.AddItem "TestRequest" & Time
    If TestRequest(JsonString) Then
        'ShowMsg "ok"
        GenerateJson = True
        '||||||||||||||||||||||||||||||||||||||||||||||||||||||
        Call StoreJSON_In_OS(nFactura, JsonString)
        '||||||||||||||||||||||||||||||||||||||||||||||||||||||
    Else
        GenerateJson = False
    End If
    'lstLog.AddItem "~~~~~~~~~~~~~~~~"

   On Error GoTo 0
   Exit Function

GenerateJson_Error:
    GenerateJson = False
    ShowMsg "Error " & Err.Number & " (" & Err.description & ") en GenerateJson de facilelectronica"
    EscribeLog "FE.GenerateJson: " & "Error " & Err.Number & " (" & Err.description & ") "
   Resume
End Function

'---------------------------------------------------------------------------------------
' Procedure : GenerateHeader
' Author    : hsequeira
' Date      : 07/03/2024
' Purpose   : GENERA EL ENCABEZADO DE FACIL FACTURA
' ID: 1
' environment = 1=Producción, 2=Pruebas
'---------------------------------------------------------------------------------------
'
Private Function GenerateHeader(id As Integer, environment As Integer) As String
    GenerateHeader = """header"": {""id"": " & id & ",""environment"": " & environment & "}"
End Function

'---------------------------------------------------------------------------------------
' Procedure : GenerateReceptor
' Author    : hsequeira
' Date      : 07/03/2024
' Purpose   : ENCABEZADO DE LA FACTURA (DOCUMENT)
'tipe =
'name
'rucType
'address
'email
'ruc
'---------------------------------------------------------------------------------------
'
Private Function GenerateReceptor(tipe As String, name As String, rucType As String, address As String, email As String, ruc As String, Optional dv As String, Optional country As String) As String
    'tipe = 01=Contribuyente, 02=Consumidor final, 03=Gobierno, 04=Extranjero
    Dim location As String
    location = "8-8-2"
    Select Case tipe
        Case "01"      'Contribuyente
            GenerateReceptor = """receptor"": {" & """type"": """ & tipe & """,""name"": """ & name & """,""ruc_type"": " & rucType & "," & """address"": """ & address & """,""email"": """ & email & """,""ruc"": """ & ruc & """,""location"": """ & location & """,""dv"": """ & dv & """}"
        Case "02"       'Consumidor FINAL
            If rucType = "1" Then  'PERSONA NATURAL
                GenerateReceptor = """receptor"": {" & """type"": """ & tipe & """,""name"": """ & name & """,""ruc_type"": " & rucType & "," & """address"": """ & address & """,""email"": """ & email & """,""ruc"": """ & ruc & """,""dv"": """ & dv & """}"
            Else                        'JURIDICO (EMPRESA, GOBIERNO)
                GenerateReceptor = """receptor"": {" & """type"": """ & tipe & """,""name"": """ & name & """,""ruc_type"": " & rucType & "," & """address"": """ & address & """,""email"": """ & email & """,""ruc"": """ & ruc & """,""dv"": """ & dv & """}"
            End If
        Case "03"       'Gobierno
            'GenerateReceptor = """receptor"": {" & """type"": """ & tipe & """,""name"": """ & name & """,""ruc_type"": " & rucType & "," & """address"": """ & address & """,""email"": """ & email & """,""ruc"": """ & ruc & """,""dv"": """ & dv & """}"
            GenerateReceptor = """receptor"": {" & """type"": """ & tipe & """,""name"": """ & name & """,""ruc_type"": " & rucType & "," & """address"": """ & address & """,""email"": """ & email & """,""ruc"": """ & ruc & """,""location"": """ & location & """,""dv"": """ & dv & """}"
        Case "04"       'Extranjero
            GenerateReceptor = """receptor"": {" & """type"": """ & tipe & """,""name"": """ & name & """,""ruc_type"": " & rucType & "," & """address"": """ & address & """,""email"": """ & email & """,""ruc"": """ & ruc & """,""country"": """ & country & """,    ""dv"": """ & dv & """}"
    End Select
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : GenerateItems
' Author    : hsequeira
' Date      : 22/03/2024
' Purpose   : GENERA LA LINEA DE ITEMS CON SU RESPECTIVO BRACKET DE TAXES
' Date      : 03/06/2024
' Purpose   : se incluye la cantidad para totalizar la factura
'---------------------------------------------------------------------------------------
'
Private Function GenerateItems(tipe As String, Exonerado As Boolean) As String
    'tipe = 01=Contribuyente, 02=Consumidor final, 03=Gobierno, 04=Extranjero
    Dim items As String
    Dim itemLine As Integer
    Dim nRegistros As Integer
    Dim cCodeTax As String
    Dim nLocalTotalFactura As Single
    Dim nLocalTax As Single
    Dim nTotalLinea As Single
    
    nRegistros = rsFETrans.RecordCount
    
    items = "["
    
    On Error Resume Next
    rsFETrans.MoveFirst
    On Error GoTo 0
    
    For itemLine = 1 To nRegistros
        ' Assuming taxes information remains the same for each item
        'Código del impuesto. Requerido para ITBMS y OTI. Valores posibles:
        'Para ITBMS: 00 = Exento, 01 = 7%, 02 = 10%, 03 = 15%
        'Para OTI: 00 = Cero, 01 = SUME911, 02 = Portabilidad númerica, 03 = Impuesto sobre seguro
        Dim itemTaxes As String
        'itemTaxes = "[{""type"": ""01"", ""amount"": 0.035, ""code"": ""01""}]"
        'itemTaxes = "[{""type"": rsFETrans!type, ""amount"": rsFETrans!amount, ""code"": rsFETrans!code}]"
        'itemTaxes = "[{""type"": & rsFETrans!type, ""amount"": & rsFETrans!amount, ""code"": & rsFETrans!code}]"
        If Exonerado Then
            cCodeTax = "00"
        Else
            Select Case rsFETrans!CON_TAX
                Case 10
                    cCodeTax = "02"
                Case 7
                    cCodeTax = "01"
                Case 0
                    cCodeTax = "00"
                Case Else
                    cCodeTax = "03"
            End Select
        End If
        ' If rsFETrans!amount = 0 Then cCodeTax = "00" Else cCodeTax = rsFETrans!code
        If Exonerado Then
            itemTaxes = "[{""type"": """ & rsFETrans!Type & """, ""amount"": " & 0# & ", ""code"": """ & cCodeTax & """}]"
            
        Else
            'INFO: PASAR EL VALOR DE DESCUENTO SIN REDONDEAR
            '2ABRIL2024
            'itemTaxes = "[{""type"": """ & rsFETrans!Type & """, ""amount"": " & Format(rsFETrans!amount, "STANDARD") & ", ""code"": """ & cCodeTax & """}]"
            'itemTaxes = "[{""type"": """ & rsFETrans!Type & """, ""amount"": " & Format(rsFETrans!amount, "0.###") & ", ""code"": """ & cCodeTax & """}]"
            itemTaxes = "[{""type"": """ & rsFETrans!Type & """, ""amount"": " & rsFETrans!amount & ", ""code"": """ & cCodeTax & """}]"
            
            'nLocalTotalFactura = nLocalTotalFactura + rsFETrans!amount
            nLocalTax = rsFETrans!amount
        End If

        Select Case tipe
            Case "2"   'Gobierno
                nTotalLinea = ((rsFETrans!price * rsFETrans!quantity) - (rsFETrans!DESCUENTO * rsFETrans!quantity)) + nLocalTax
                nLocalTotalFactura = nLocalTotalFactura + nTotalLinea
                items = items & GenerateItem(rsFETrans!LIN, rsFETrans!price, "und", rsFETrans!quantity, rsFETrans!description, itemTaxes, rsFETrans!DESCUENTO, rsFETrans!internal_code, tipe, "9010", rsFETrans!depto) & ","
                
            Case Else
                'nLocalTotalFactura = nLocalTotalFactura + (rsFETrans!price - rsFETrans!DESCUENTO)
                nTotalLinea = ((rsFETrans!price * rsFETrans!quantity) - (rsFETrans!DESCUENTO * rsFETrans!quantity)) + nLocalTax
                nLocalTotalFactura = nLocalTotalFactura + nTotalLinea
                items = items & GenerateItem(rsFETrans!LIN, rsFETrans!price, "und", rsFETrans!quantity, rsFETrans!description, itemTaxes, rsFETrans!DESCUENTO, rsFETrans!internal_code, tipe, "", rsFETrans!depto) & ","
        End Select
        'Debug.Print "Tot Linea: " & nTotalLinea
        'Debug.Print "Price: " & rsFETrans!price & " | tax: " & rsFETrans!amount
        rsFETrans.MoveNext
    Next itemLine
    
    ' Remove the trailing comma
    If Right(items, 1) = "," Then
        items = Left(items, Len(items) - 1)
    End If
    
    GenerateItems = items & "]"
    If local_total_factura <> nLocalTotalFactura Then local_total_factura = nLocalTotalFactura
    'ShowMsg "total factura: " & nLocalTotalFactura
End Function
'---------------------------------------------------------------------------------------
' Procedure : GenerateItem
' Author    : hsequeira
' Date      : 12/03/2024
' update: 27/04/2024
' Purpose   : Genera cada linea de Items de la factura
'---------------------------------------------------------------------------------------
'
Rem Private Function GenerateItem(Line As Integer, price As Single, mu As String, quantity As Integer, description As String, taxes As String, discount As Single, internalCode As Long, ctipe As String, Optional gns As String, Optional depto As Integer) As String
'INFO: 14MARZO2026: AGREGAR PLU
Private Function GenerateItem(Line As Integer, price As Single, mu As String, quantity As Integer, description As String, taxes As String, discount As Single, internalCode As Long, ctipe As String, Optional gns As String, Optional depto As Integer) As String
    Dim nLocalDiscount As Single
    
    nLocalDiscount = Format(discount, "0.0000")
    Select Case ctipe
        'Case "01"
        Case "2"
            GenerateItem = "{""line"": " & Line & ",""price"": " & price & ",""mu"": """ & mu & ""","
            Rem GenerateItem = GenerateItem & """quantity"": " & quantity & ",""description"": """ & description & ""","
            'INFO: ADDING DEPTO TO JASON
            'INFO: 3FEB2026
            Rem GenerateItem = GenerateItem & """quantity"": " & quantity & ",""description"": """ & description & ",""depto"": " & depto & ""","
            GenerateItem = GenerateItem & """quantity"": " & quantity & ",""description"": """ & description & """,""depto"": " & depto & ","
            
            'GenerateItem = GenerateItem & """taxes"": " & taxes & ",""discount"": " & discount & ", ""gns"": " & gns & ",""internal_code"": " & internalCode & "}"
            GenerateItem = GenerateItem & """taxes"": " & taxes & ",""discount"": " & nLocalDiscount & ", ""gns"": " & gns & ",""internal_code"": " & internalCode & "}"
            'nLocalDiscount
        Case Else
            GenerateItem = "{""line"": " & Line & ",""price"": " & price & ",""mu"": """ & mu & ""","
            Rem GenerateItem = GenerateItem & """quantity"": " & quantity & ",""description"": """ & description & ""","
            GenerateItem = GenerateItem & """quantity"": " & quantity & ",""description"": """ & description & """,""depto"": " & depto & ","
            'INFO: ADDING DEPTO TO JASON
            'INFO: 3FEB2026
            Rem GenerateItem = GenerateItem & """quantity"": " & quantity & ",""description"": """ & description & ",""depto"": " & depto & ""","
            Rem GenerateItem = GenerateItem & """quantity"": " & quantity & ",""description"": """ & description & ",depto"": " & depto & ""","
            Rem GenerateItem = GenerateItem & """quantity"": " & quantity & ",""description"": """ & description & """,""depto"": " & depto & ""","""
            Rem GenerateItem = "{" & appos & "quantity" & appos & ": quantity ," & appos & "description" & appos & ": " & appos & description & appos & "," & appos & "depto" & appos & ": depto}"
                
            
            'GenerateItem = GenerateItem & """taxes"": " & taxes & ",""discount"": " & discount & ",""internal_code"": " & internalCode & "}"
            GenerateItem = GenerateItem & """taxes"": " & taxes & ",""discount"": " & nLocalDiscount & ",""internal_code"": " & internalCode & "}"
        End Select
End Function

'---------------------------------------------------------------------------------------
' Procedure : Generate_NEW_Payments
' Author    : hsequeira
' Date      : 10/09/2024
' Purpose   : GENERA PAGOS DESDE EL RECORSET EN VEZ DEL MSHFlexGrid
'---------------------------------------------------------------------------------------
'
Private Function Generate_NEW_Payments(rsMisPagos As ADODB.Recordset, ByRef nPagosFuturo As Single) As String
Dim MasDeUnPago As Boolean
Dim cCadena As String
Dim coma As String
Dim nTipoPago As Integer
Dim cTextoTipoPago As String
Dim nValorPago As Single
Dim nMax As Integer
Dim nFixRows As Integer
Dim i As Integer
Dim UnSoloPago As Boolean
Dim nDiferenciaConFE As Single

rsMisPagos.MoveFirst

If rsMisPagos.RecordCount > 0 Then coma = ","

If rsMisPagos.RecordCount = 1 Then
    UnSoloPago = True
Else
    UnSoloPago = False
End If


If UnSoloPago Then
    Rem cCadena = cCadena & "{""type"": """ & rsMisPagos!CODIGO_FE & """,""amount"": " & local_total_factura & ",""description"": """ & rsMisPagos!DESCRIP_FE & """}" & coma
    cCadena = cCadena & "{""type"": """ & GetPagoElectronico(rsMisPagos!CODIGO) & """,""amount"": " & local_total_factura & ",""description"": """ & rsMisPagos!DESCRIP_FE & """}" & coma
    nPagosFuturo = nPagosFuturo + local_total_factura
Else
    Do While Not rsMisPagos.EOF
        Rem cCadena = cCadena & "{""type"": """ & rsMisPagos!CODIGO_FE & """,""amount"": " & rsMisPagos!MONTO_1 & ",""description"": """ & rsMisPagos!DESCRIP_FE & """}" & coma
        cCadena = cCadena & "{""type"": """ & GetPagoElectronico(rsMisPagos!CODIGO) & """,""amount"": " & rsMisPagos!MONTO_1 & ",""description"": """ & rsMisPagos!DESCRIP_FE & """}" & coma
        nPagosFuturo = nPagosFuturo + rsMisPagos!MONTO_1
        rsMisPagos.MoveNext
    Loop
End If

'INFO: 12NOV2024
' AJUSTE A LA FACTURA X LOS DECIMALES QUE CALCULA LA FE
If local_total_factura > nPagosFuturo Then
    Rem nDiferenciaConFE = local_total_factura - nPagosFuturo
    nDiferenciaConFE = Round(local_total_factura - nPagosFuturo, 5)
    'A LA ELECTRONIA DICE QUE LE FALTA PARA PODER COBRAR LA FACTURA, SE HACE UN PAGO ADICIONAL
    cCadena = cCadena & "{""type"": """ & "99" & """,""amount"": " & nDiferenciaConFE & ",""description"": """ & "FE . AJUSTE" & """}" & coma
    nPagosFuturo = local_total_factura
End If



If Right(cCadena, 1) = "," Then cCadena = Left(cCadena, Len(cCadena) - 1)

Generate_NEW_Payments = "[" & cCadena & "]"
'
'Set rsFEPropinas = New ADODB.Recordset
'rsFEPropinas.Fields.Append "description", adChar, 35
'rsFEPropinas.Fields.Append "amount", adSingle
'
'rsFEPropinas.CursorType = adOpenStatic
'rsFEPropinas.LockType = adLockOptimistic
'rsFEPropinas.Open


End Function

'---------------------------------------------------------------------------------------
' Procedure : GeneratePayments
' Author    : hsequeira
' Date      : 12/03/2024
' Purpose   : Guarda la Cadena con los pagos
' INFO: UPDATE 26MAY2024, SE PONE 1 A nMax, ANTES ERA 0, ENTONCES PONIA 2 PAGOS EN VEZ DE UNO
' INFO: REEMPLAZADO POR GENERATE NEW PAYMENTS
'---------------------------------------------------------------------------------------
'
'Private Function GeneratePayments(lstPagos As MSHFlexGrid, ByRef nPagosFuturo As Single) As String
'Dim MasDeUnPago As Boolean
'Dim cCadena As String
'Dim coma As String
'Dim nTipoPago As Integer
'Dim cTextoTipoPago As String
'Dim nValorPago As Single
'Dim nMax As Integer
'Dim nFixRows As Integer
'Dim i As Integer
'
'
'
'Set rsFEPagos = New ADODB.Recordset
''01=Crédito, 02=Efectivo, 03=Tarjeta de crédito, 04=Tarjeta de débito, 05=Tarjeta de fidelidad,
''06=Vale, 07=Tarjeta de regalo, 08=Transferencia o déposito bancario, 09=Cheque,
''10=Punto de venta, 99 = Otro
'    ' "[{""type"": 99, ""amount"": 61.85, ""description"": ""Medio de pago de prueba""}]"
'
'rsFEPagos.Fields.Append "type", adChar, 2
'rsFEPagos.Fields.Append "description", adChar, 35
'rsFEPagos.Fields.Append "amount", adSingle
'
'rsFEPagos.CursorType = adOpenStatic
'rsFEPagos.LockType = adLockOptimistic
'rsFEPagos.Open
'
'
'Set rsFEPropinas = New ADODB.Recordset
'rsFEPropinas.Fields.Append "description", adChar, 35
'rsFEPropinas.Fields.Append "amount", adSingle
'
'rsFEPropinas.CursorType = adOpenStatic
'rsFEPropinas.LockType = adLockOptimistic
'rsFEPropinas.Open
'
'If Left(mi_App_Name, 7) = "SoloMix" Then
''INFO: UPDATE 26MAY2024, SE PONE 1 A nMax, ANTES ERA 0, ENTONCES PONIA 2 PAGOS EN VEZ DE UNO
'    nMax = 1
'Else
'    If mi_Panta = "PAGOS" Then
'        nMax = 0
'    Else
'        nMax = 1
'    End If
'End If
'
'If lstPagos.Rows = 1 Then nFixRows = 0 Else nFixRows = 1
'
'For i = nMax To (lstPagos.Rows - nFixRows)
'    On Error Resume Next
'    lstPagos.row = i
'    On Error GoTo 0
'    lstPagos.col = 0
'    nTipoPago = lstPagos.Text
'    'INFO: 5JUN2007
'    lstPagos.col = 1
'    cTextoTipoPago = Trim(lstPagos.Text)
'
'    lstPagos.col = 2
'    nValorPago = Format(lstPagos.Text, "STANDARD")
'
'    If VerificaSI_es_Propina(cTextoTipoPago) Then
'        rsFEPropinas.AddNew
'
'        rsFEPropinas.Fields("description").value = cTextoTipoPago
'        rsFEPropinas.Fields("amount").value = nValorPago
'        rsFEPropinas.Update
'
'    Else
'        rsFEPagos.AddNew
'
'        rsFEPagos.Fields("type").value = GetPagoElectronico(nTipoPago)
'        rsFEPagos.Fields("description").value = cTextoTipoPago & " / " & cTextoTipoPago
'        rsFEPagos.Fields("amount").value = nValorPago
'        rsFEPagos.Update
'    End If
'    Sleep 50
'
'Next
'
'
''01=Crédito, 02=Efectivo, 03=Tarjeta de crédito, 04=Tarjeta de débito, 05=Tarjeta de fidelidad,
''06=Vale, 07=Tarjeta de regalo, 08=Transferencia o déposito bancario, 09=Cheque,
''10=Punto de venta, 99 = Otro
'    'GeneratePayments = "[{""type"": 99, ""amount"": 61.85, ""description"": ""Medio de pago de prueba""}]"
'
'    rsFEPagos.MoveFirst
'
'    If rsFEPagos.RecordCount > 0 Then coma = ","
'
'    Do While Not rsFEPagos.EOF
'        'cCadena = "{""type"": " & rsFEPagos!Type & ",""amount"": " & rsFEPagos!amount & ",""description"": " & rsFEPagos!description & "}" & coma
'        'cCadena = cCadena & "{""type"": """ & rsFEPagos!Type & """,""amount"": " & rsFEPagos!amount & ",""description"": """ & Left(rsFEPagos!description, 10) & """}" & coma
'        'GenerateItem = GenerateItem & """taxes"": " & taxes & ",""discount"": " & discount & ", ""gns"": " & gns & ",""internal_code"": " & internalCode & "}"
'        '''''''''''''''''''''''''''cCadena = cCadena & "{""type"": """ & rsFEPagos!Type & """,""amount"": " & rsFEPagos!amount & "}" & coma
'        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'        'INFO: UPDATE PARA CUMPLIR CON EL API DE FACTURA FACIL QUE REQUIERE HASTA 6 DECIMALES PARA CALCULAR EL MONTO TOTAL
'        '31MAY2024
'        'cCadena = cCadena & "{""type"": """ & rsFEPagos!Type & """,""amount"": " & local_total_factura & "}" & coma
'        '|||||||||||||||||||||||||||||||||||||
'        'INFO: UPDATE 9JUL2024
'        '|||||||||||||||||||||||||||||||||||||
'        'cCadena = cCadena & "{""type"": """ & rsFEPagos!Type & """,""amount"": " & Format(local_total_factura, "STANDARD") & "}" & coma
'        'cCadena = cCadena & "{""type"": """ & rsFEPagos!Type & """,""amount"": " & Format(local_total_factura, "STANDARD") & ",""description"": """ & rsFEPagos!description & """}" & coma
'        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'        'UPDATE" 21AGO2024
'        'SE ARREGLA EL MONTO PAGANDO PARA QUE NO INCLUYA COMAS
'        cCadena = cCadena & "{""type"": """ & rsFEPagos!Type & """,""amount"": " & local_total_factura & ",""description"": """ & rsFEPagos!description & """}" & coma
'        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'        nPagosFuturo = nPagosFuturo + rsFEPagos!amount
'        rsFEPagos.MoveNext
'    Loop
'
'    If Right(cCadena, 1) = "," Then cCadena = Left(cCadena, Len(cCadena) - 1)
'
'    'GeneratePayments = "[{""type"": 99, ""amount"": 61.85, ""description"": ""Medio de pago de prueba""}]"
'    ''GeneratePayments = "[" & cCadena & "], ""total"": " & TotalPagado
'    'GeneratePayments = "[" & cCadena & ", ""total"": " & TotalPagado & "]"
'    GeneratePayments = "[" & cCadena & "]"
'
'
'End Function

Private Function TestRequest(jsonContent As String, Optional TipoRequest As String) As Boolean
Dim response As String
Dim http As Object
Dim nFreefile As Byte
'INFO: TOTAL DE LA FACTURA
'19NOV2024
Dim gTot As String
Dim xmltotal_factura As String
Dim cTemp As String
Dim cTemp2 As String
Dim nCurrentValor As Single
Dim nCurrenCounter As Long

   On Error GoTo TestRequest_Error

    Call Clear_FE_Variables

    Set http = CreateObject("MSXML2.ServerXMLHTTP")

    Dim url As String
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'url = "https://backend-qa-api.facturafacil.com.pa/api/pac/reception_fe/detailed/"
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    url = mi_urlWebservice
    

    http.Open "POST", url, False
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'http.setRequestHeader "X-FF-Company", "8b6a0910-20c1-48af-a464-46c512a811b7"
    'http.setRequestHeader "X-FF-API-Key", "nHtzOvXZqjrI-XAoyqAnG9zexL9FX1hy6vIQ2rOTsu7aHBEoAPATt3r936Tiu0hYIqIKYIO8UGWMm0NubuEnzm3WY_9y5FNLQakU5hc7w0t5ttWDjfOXaL5UxdEOPVPX"
    'http.setRequestHeader "X-FF-Branch", "521bf0a9-e8aa-488f-b43f-a83122e534f0"
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    http.setRequestHeader "X-FF-Company", X_FF_Company
    http.setRequestHeader "X-FF-API-Key", X_FF_API_Key
    http.setRequestHeader "X-FF-Branch", X_FF_Branch
    http.setRequestHeader "Accept", "application/json"
    http.setRequestHeader "Content-Type", "application/json"

    ' Modify the JSON content as needed
    'Dim jsonContent As String
    'jsonContent = "{""header"":{""id"":1,""environment"":""2""},""document"":{""fd_number"":""36"",""receptor"":{""type"":""02"",""name"":""Nombre Cliente"",""ruc_type"":""1"",""address"":""Dirección Cliente"",""email"":""cliente@correo.com"",""ruc"":""123123123""},""items"":[{""line"":1,""price"":0.5,""mu"":""und"",""quantity"":1,""description"":""Producto de prueba"",""taxes"":[{""type"":""01"",""amount"":0.035,""code"":""01""}],""discount"":0,""internal_code"":""123123""}],""payments"":[{""type"":""99"",""amount"":""0.54"",""description"":""Medio de pago de prueba""}],""type"":""01"",""info"":""<string>""}}"
    'Debug.Print jsonContent
    'Form1.lstLog.AddItem "http.send jsonContent" & Time
    http.send jsonContent


    ' Extracting values from JSON
    Dim cufe As String
    Dim qrCodeData As String
    Dim pdfUrl As String
    Dim mensajes As String
    Dim mensajesType As String
    Dim mensajesCode As String
    Dim mensajesMsg As String
    
    ' Check if the request was successful (status code 200)
    If http.Status = 201 Then
        cFE_LastErrorCode = http.Status
        response = http.responseText
        'Form1.lstLog.AddItem "RESPONSE OK " & Time
    Else
        'Form1.lstLog.AddItem "FAIL http.send jsonContent" & Time
        
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        On Error Resume Next
        
        nFreefile = FreeFile()
        
        Open App.Path & "\" & "1.json" For Output As #nFreefile
        Print #nFreefile, jsonContent
        Close #nFreefile
        On Error GoTo 0

        
        response = "Error: " & http.Status & " - " & http.statusText & " + " & http.responseText
        
        nFreefile = FreeFile()
        
        Open App.Path & "\" & "1.res" For Output As #nFreefile
        Print #nFreefile, response
        Close #nFreefile
        On Error GoTo 0
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        
        mensajesType = ExtractValueFromJSON3(response, "type")
        mensajesCode = ExtractValueFromJSON3(response, "code")
        mensajesMsg = ExtractValueFromJSON3(response, "message")
        cFE_LastErrorCode = mensajesCode
        
        If TipoRequest = "NC" Then
            ShowMsg "ERROR EN NOTA DE CREDITO (TestRequest): " & vbCrLf & "Codigo: " & mensajesCode & vbCrLf & "Motivo Error: " & mensajesMsg, vbBlue, vbRed
            EscribeLog "ERROR EN NOTA DE CREDITO: " & mensajesCode & ". " & mensajesMsg
        Else
            ShowMsg "ERROR EN FACTURA (TestRequest): " & vbCrLf & "Codigo: " & mensajesCode & vbCrLf & "Motivo Error: " & mensajesMsg, vbBlue, vbRed
            EscribeLog "ERROR EN FACTURA: " & mensajesCode & ". " & mensajesMsg
        End If
        
        TestRequest = False
        Exit Function
    End If
    
    'Call AI_Examine_SOAP_Response(response)
    
    
    ''Debug.Print response




'xmlResponse = "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"">" & "<SOAP-ENV:Body><ns2:gTot>12345.67</ns2:gTot></SOAP-ENV:Body></SOAP-ENV:Envelope>"


'INFO: TOTAL DE LA FACTURA
'19NOV2024
'SE DEBE LIMPIAR CUANDO SE TIRA LA Z
gTot = ExtractGTot(response, "gTot")
On Error Resume Next
cTemp = RegRead("HKCU\Software\SoloSoftware\SoloMix\FE_Sales")
cTemp2 = RegRead("HKCU\Software\SoloSoftware\SoloMix\FE_Counter")
On Error GoTo 0
If Left(gTot, 6) = "Error:" Then
    EscribeLog "Error extracting gTot en Factura (" & nNUMTrans & ")"
Else
    Rem MsgBox "Extracted gTot value: " & gTot
    xmltotal_factura = ExtractGTot(gTot, "dVTot")
   
    If cTemp = "" Then
        nCurrentValor = 0
        nCurrenCounter = 0
    Else
        nCurrenCounter = CLng(cTemp2)
        cTemp = RegRead("HKCU\Software\SoloSoftware\SoloMix\FE_Sales")
        nCurrentValor = CSng(cTemp)
    End If
    Rem cTemp = CSng(cMonto)
    nCurrenCounter = nCurrenCounter + 1
    nCurrentValor = nCurrentValor + CSng(xmltotal_factura)
    
    RegWrite "HKCU\Software\SoloSoftware\SoloMix\FE_Counter", nCurrenCounter
    RegWrite "HKCU\Software\SoloSoftware\SoloMix\FE_Sales", nCurrentValor
End If


'Form1.lstLog.AddItem "ExtractValueFromJSON 2 y 3 " & Time
'Debug.Print response
cufe = ExtractValueFromJSON3(response, "cufe")
FE_CUFE = cufe
'Debug.Print "CUFE:" & cufe
qrCodeData = ExtractValueFromJSON3(response, "qr_code_data")
FE_QRData = qrCodeData
'Debug.Print qrCodeData
pdfUrl = ExtractValueFromJSON3(response, "pdf_url")
FE_pdfURL = pdfUrl

FE_process_date = ExtractValueFromJSON3(response, "process_date")
FE_digest_value = ExtractValueFromJSON3(response, "digest_value")
FE_authorization_number = ExtractValueFromJSON3(response, "authorization_number")


'Debug.Print "PDF URL: " & pdfUrl

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
mensajes = ExtractValueFromJSON2(response, "messages")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

mensajesType = ExtractValueFromJSON3(mensajes, "type")
mensajesCode = ExtractValueFromJSON3(mensajes, "code")
mensajesMsg = ExtractValueFromJSON3(mensajes, "message")

' Displaying extracted values
'MsgBox "CUFE: " & cufe & vbCrLf & "QR Code Data: " & qrCodeData & vbCrLf & "PDF URL: " & pdfUrl & "Mensaje: " & mensajes


'If FE_QRData = "" Then
'    ShowMsg "FE. FAIL. NO SE PUEDE OBTENER QR", vbBlue, vbCyan
'    ShowMsg "mensajesType : " & mensajesType & ", Code : " & mensajesCode & " .Msg: " & mensajesMsg, vbBlue, vbRed
'    EscribeLog "FE. FAIL. NO SE PUEDE OBTENER QR"
'    TestRequest = False
'    Exit Function
'Else
'    TestRequest = True
'End If


        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        On Error Resume Next
        
        nFreefile = FreeFile()
        
        Open App.Path & "\" & "1.json" For Output As #nFreefile
        Print #nFreefile, jsonContent
        Close #nFreefile
        On Error GoTo 0

        On Error Resume Next
        response = "Error: " & http.Status & " - " & http.statusText & " + " & http.responseText
        
        nFreefile = FreeFile()
        
        Open App.Path & "\" & "1.res" For Output As #nFreefile
        Print #nFreefile, response
        Close #nFreefile
        On Error GoTo 0
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||



'If mensajesType = "R" Then
If mensajesType = "N" And mensajesCode = "0260" Then
    '~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~
    'INFO: UPDATE 16JUL2024
    '~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~
    If Left(mi_App_Name, 7) = "SoloMix" Then
        EscribeLog "FE. OK. Factura: " & rs00!TRANS + 1 & " - CUFE: " & Right(FE_CUFE, 15) & ", PDF: " & Right(FE_QRData, 15)
    Else
        EscribeLog "FE. OK. Factura: " & nNUMTrans & " - CUFE: " & Right(FE_CUFE, 15) & ", PDF: " & Right(FE_QRData, 15)
    End If
    TestRequest = True
Else
    'Debug.Print jsonContent
    'Debug.Print "========================"
    'Debug.Print response
    'ShowMsg "Factura # " & Str(nNUMTrans) & ". FAIL. TestRequest : " & vbCrLf & "mensajesType : " & mensajesType & ", Code : " & mensajesCode & " .Msg: " & mensajesMsg, vbBlue, vbRed
    If Left(mi_App_Name, 7) = "SoloMix" Then
        ShowMsg "Factura # " & Str(rs00!TRANS + 1) & ". FAIL. TestRequest : " & vbCrLf & "mensajesType : " & mensajesType & ", Code : " & mensajesCode & " .Msg: " & mensajesMsg, vbBlue, vbRed
        EscribeLog "FE. FAIL. mensajesType : " & mensajesType & ", Code : " & mensajesCode & " .Msg: " & mensajesMsg
    Else
        ShowMsg "Factura # " & Str(nNUMTrans) & ". FAIL. TestRequest : " & vbCrLf & "mensajesType : " & mensajesType & ", Code : " & mensajesCode & " .Msg: " & mensajesMsg, vbBlue, vbRed
        EscribeLog "FE. FAIL. mensajesType : " & mensajesType & ", Code : " & mensajesCode & " .Msg: " & mensajesMsg
    End If
    
    If mensajesCode = "1002" Then       '1002= FACTURA DUPLICADA, INCREMENTAR EL CONTADOR e INTENTAR DE NUEVO
        
        If Left(mi_App_Name, 7) = "SoloMix" Then
            Call SOLOTrans("BEGIN")
            msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1"
            Call SOLOTrans("COMMIT")
        Else
            On Error Resume Next
            'INFO: 11JUN2012)
            #If FAST_POS = 1 Then
                'OPOS_Eval = Sys_Pos.Coptr1.JrnEmpty
            #Else
                'Call GetNewNumTrans
            #End If
            'GetNewNumTrans
            On Error GoTo 0
        End If
    End If
End If
    
'    Debug.Print "code===> " & mensajesCode
'    Debug.Print "msg===> " & mensajesMsg
'    Debug.Print "type ===> " & mensajesType


'Form1.lstLog.AddItem "END ! " & Time

   On Error GoTo 0
   Exit Function

TestRequest_Error:

    Select Case Err.Number
        Case -2147012889
            If FE_Facil_ONLINE Then
                ShowMsg "NO SE PUEDE CONECTAR AL SERVICIO DE " & vbCrLf & vbCrLf & "FACTURACION ELECTRONICA <<ON LINE>>", vbBlue, vbRed
            End If
        Case Else
            ShowMsg "Error " & Err.Number & " (" & Err.description & ") en TestRequest", vbBlue, vbCyan
    End Select
    Resume
    EscribeLog "FE.TestRequest: " & "Error " & Err.Number & " (" & Err.description & ") "
    
'    If FE_Facil_ONLINE Then
'        ShowMsg "Error " & Err.Number & " (" & Err.description & ") en TestRequest", vbBlue, vbCyan
'        EscribeLog "FE.TestRequest: " & "Error " & Err.Number & " (" & Err.description & ") "
'    Else
'        ShowMsg "Error. SISTEMA PASANDO A OFF-LINE en TestRequest", vbBlue, vbCyan
'        EscribeLog "FE.TestRequest. PASANDO A OFF-LINE "
'    End If
    Rem Resume Next
End Function

'---------------------------------------------------------------------------------------
' Procedure : ExtractValueFromJSON2
' Author    : hsequeira
' Date      : 17/03/2024
' Purpose   : EXTRAE EL VALOR BUSCADO DENTRO DE UN JSON USANDO UN METODO MAS SIMPLE
'---------------------------------------------------------------------------------------
'
Private Function ExtractValueFromJSON2(JsonString As String, Key As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim keyPos As Long
    
    ' Find the starting position of the key
   On Error GoTo ExtractValueFromJSON2_Error
    
    On Error Resume Next
    keyPos = InStr(JsonString, """" & Key & """")
    On Error GoTo 0
    
    If keyPos = 0 Then
        ExtractValueFromJSON2 = ""
        Exit Function
    End If
    
    'startPos = InStr(keyPos, jsonString, ":") + 2 ' Adding 2 to skip ": and opening double quotes
    startPos = InStr(keyPos, JsonString, "[")  ' Adding 2 to skip ": and opening double quotes
    
    ' Find the end position of the value
    If InStr(startPos, JsonString, "}") > 0 Then
        endPos = InStr(startPos, JsonString, "}")  ' Subtract 1 to exclude comma
    Else
        endPos = InStr(startPos, JsonString, "}")  ' Subtract 1 to exclude closing brace
    End If
    
    ' Extract the value
    ExtractValueFromJSON2 = Mid(JsonString, startPos, endPos - startPos + 1) & "]"

   On Error GoTo 0
   Exit Function

ExtractValueFromJSON2_Error:

    ShowMsg "Error " & Err.Number & " (" & Err.description & ") in ExtractValueFromJSON2", vbYellow, vbRed
End Function


'---------------------------------------------------------------------------------------
' Procedure : ExtractValueFromJSON3
' Author    : hsequeira
' Date      : 17/03/2024
' Purpose   : EXTRA EL VALOR BUSCADO DENTRO DE UN JSON GENERAL (SOAP)
'---------------------------------------------------------------------------------------
'
Private Function ExtractValueFromJSON3(JsonString As String, Key As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim keyPos As Long
    
    ' Find the starting position of the key
   On Error GoTo ExtractValueFromJSON3_Error
    On Error Resume Next
    keyPos = InStr(JsonString, """" & Key & """")
    On Error GoTo 0
    If keyPos = 0 Then
        ExtractValueFromJSON3 = ""
        Exit Function
    End If
    
    ' Adjust the starting position based on key position
    startPos = InStr(keyPos, JsonString, ":") + 1 ' Adjusted to include colon
    
    ' Adjust the starting position to handle different representations of values
    startPos = InStr(startPos, JsonString, """") + 1 ' Adjusted to skip opening quote
    
    ' Find the end position of the value
    endPos = InStr(startPos, JsonString, """") - 1 ' Adjusted to exclude closing quote
    
    ' Extract the value
    ExtractValueFromJSON3 = Mid(JsonString, startPos, endPos - startPos + 1)

   On Error GoTo 0
   Exit Function

ExtractValueFromJSON3_Error:

    ShowMsg "Error " & Err.Number & " (" & Err.description & ") en ExtractValueFromJSON3 ", vbYellow, vbRed
End Function

'---------------------------------------------------------------------------------------
' Procedure : Load_FE_Mesa
' Author    : hsequeira
' Date      : 22/03/2024
' Purpose   : CARGA LA MESA PARA LA FACTURA ELECTRONICA
' PARA OBTENER EL AMOUNT LINEAL, SE MULTIPLICA FE_DESCUENTO * CANT
' Date      : 3/06/2024
'---------------------------------------------------------------------------------------
'
Private Function Load_FE_Mesa() As Boolean
Dim cSQL As String
Set rsFETrans = New ADODB.Recordset

cSQL = "SELECT LIN, PRECIO_UNIT  as price, 'und' as mu, CANT AS quantity, PLU as internal_code,"
'cSQL = cSQL & " DESCRIP AS description, '01' as type, '01' as code,  ((PRECIO - FE_DESCUENTO) * (CON_TAX/100)) as amount, "
cSQL = cSQL & " DESCRIP AS description, '01' as type, '01' as code,  ((PRECIO - (FE_DESCUENTO * CANT)) * (CON_TAX/100)) as amount, "
cSQL = cSQL & " FE_DESCUENTO as descuento, CON_TAX "
'-------------------------------------------
''INFO: ADDING DEPTO 12FEB2026
cSQL = cSQL & " , DEPTO AS depto "
'-------------------------------------------
cSQL = cSQL & " FROM TMP_TRANS "
cSQL = cSQL & " WHERE MESA = " & nMesa
If nCta <> 0 Then
    cSQL = cSQL & " AND CUENTA = " & nCta
End If
cSQL = cSQL & " AND LEFT(TIPO,2) NOT IN ('VO','BV','DC','BE','EC') "
cSQL = cSQL & " ORDER BY LIN"

rsFETrans.Open cSQL, msConn, adOpenStatic, adLockOptimistic

End Function

''---------------------------------------------------------------------------------------
'' Procedure : ClearFE_Variables
'' Author    : hsequeira
'' Date      : 17/03/2024
'' Purpose   : LIMPIA TODAS LAS VARIABLES QUE AFECTAN LA FACTURA
''---------------------------------------------------------------------------------------
''
'Private Function ClearFE_Variables()
'
'End Function


Public Function OFF_Line_PlaceHolder(nFactura As String, mmFactura As Long)
''---------------------------------------------------------------------------------------
'' Procedure : OFF_Line_PlaceHolder
'' Author    : hsequeira
'' Date      : 17/03/2024
'' Purpose   : PLACEHOLDER ANTES DEL DESARROLLO
''---------------------------------------------------------------------------------------
'METODO OFF LINE
Rem ShowMsg "SE PROCEDE A HACER UNA FACTURA" & vbCrLf & "(OFF LINE)", vbBlue, vbCyan
Rem EscribeLog "SE PROCEDE A HACER FACTURA (OFF LINE) # (" & mmFactura & "), POR VALOR DE " & nFactura
'Debug.Print "SE PROCEDE A HACER FACTURA (OFF LINE) # (" & mmFactura & "), POR VALOR DE " & nFactura
End Function

Public Function GetPagoElectronico(nPago As Integer) As String
'OBTIENE EL ID DEL PAGO DE  FE ON LINE

On Error Resume Next
rsEPago.MoveFirst
rsEPago.Find "CODIGO =" & nPago
If Not rsEPago.EOF Then
    GetPagoElectronico = IIf(IsNull(rsEPago!ID_FE), "02", Left(rsEPago!ID_FE, 2))
    Rem GetPagoElectronico = IIf(IsNull(rsEPago!ID_FE), "02", rsEPago!ID_FE)
Else
    GetPagoElectronico = "02"
End If

End Function
Public Function GetPagoElectronico_OFFLINE(nPago As Integer) As String
'OBTIENE EL ID DEL PAGO DE  FE OFF LINE

On Error Resume Next
rsEPago.MoveFirst
rsEPago.Find "CODIGO =" & nPago
If Not rsEPago.EOF Then
    GetPagoElectronico_OFFLINE = IIf(IsNull(rsEPago!ID_FE), "01", Right(rsEPago!ID_FE, 2))
    Rem GetPagoElectronico = IIf(IsNull(rsEPago!ID_FE), "02", rsEPago!ID_FE)
Else
    'REGRESA EFECTIVO
    GetPagoElectronico_OFFLINE = "01"
End If

End Function
'-----------------------------------------------------------------------------------------
' Procedure : LoadPagoElectronico
' Author    : hsequeira
' Date      : 12/09/2023
' Purpose   : CARGA LAS FORMAS DE PAGO DE FACTURACION ELECTRONICA
'-----------------------------------------------------------------------------------------
'
Public Function LoadPagoElectronico() As Boolean

LoadPagoElectronico = True
   On Error GoTo LoadPagoElectronico_Error

Set rsEPago = New ADODB.Recordset
rsEPago.Open "SELECT CODIGO, DESCRIP, ID_FE FROM PAGOS ORDER BY CODIGO ", msConn, adOpenStatic, adLockOptimistic

   On Error GoTo 0
   Exit Function

LoadPagoElectronico_Error:

    ShowMsg "PRIMERO DEBE DEFINIR EL AMBIENTE DE LA FACTURACION ELECTRONICA " & vbCrLf & vbCrLf & "(FE.EXE)" & vbCrLf & " (LoadPagoElectronico) ", vbBlue, vbCyan
    LoadPagoElectronico = False

End Function

'---------------------------------------------------------------------------------------
' Procedure : CheckWebsite
' Author    : hsequeira
' Date      : 17/03/2024
' Purpose   : chatgpt
'---------------------------------------------------------------------------------------
'
Public Function Check_FE_Website(Optional cCallado As String, Optional nTimer As Integer) As Boolean
    Dim objHTTP As Object
    Dim strURL As String
    Dim cErrMsg As String
    
    ' Website URL to check

   On Error GoTo CheckWebsite_Error

    'strURL = "https://backend-qa-api.facturafacil.com.pa/"
    'INFO: WEB SITE DE SITIO DE VENTAS
    Rem strURL = "https://api.facturafacil.com.pa/"
    '///////////////////////////////////////////////////////////////////////////////////////////
    strURL = "https://api.facturafacil.com.pa/api/health/ready/"
    '///////////////////////////////////////////////////////////////////////////////////////////
    Rem strURL = "https://backend-qa-api.facturafacil.com.pa/api/health/ready/"
    'strURL = "http://httpstat.us/200"
    'strURL = "https://pos.facturafacil.com.pa"

    'strURL = "https://soloposASDASmenu.com"
    
    ' Create HTTP object
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Open HTTP connection
    objHTTP.Open "GET", strURL, False
    
    ' Send HTTP request
    objHTTP.send
    
    If nTimer = 0 Then
        Sleep 500
    Else
        Sleep nTimer
    End If
    
    ' Check if status is 200 (OK)
    'Debug.Print objHTTP.Status
    If objHTTP.Status = 200 Then
        'ShowMsg "FE WEB SERVICE <<OK>>", vbWhite, vbBlue
        Check_FE_Website = True
        'INFO: 13ENE2026. SI ESTAMOS OFF LINE QUEDATE OFF LINE
        If cLogMode = "SI" Then
            FE_Facil_ONLINE = False
        Else
            FE_Facil_ONLINE = True
        End If
'        X_FF_Company = "8b6a0910-20c1-48af-a464-46c512a811b7"
'        X_FF_API_Key = "nHtzOvXZqjrI-XAoyqAnG9zexL9FX1hy6vIQ2rOTsu7aHBEoAPATt3r936Tiu0hYIqIKYIO8UGWMm0NubuEnzm3WY_9y5FNLQakU5hc7w0t5ttWDjfOXaL5UxdEOPVPX"
'        X_FF_Branch = "521bf0a9-e8aa-488f-b43f-a83122e534f0"
        
    Else
        EscribeLog "FE WEB SERVICE IS DOWN"
        'ShowMsg "FE WEB SERVICE IS DOWN", vbYellow, vbRed
        Check_FE_Website = False
        FE_Facil_ONLINE = False
    End If
    
    ' Clean up
    Set objHTTP = Nothing

   On Error GoTo 0
   Exit Function

CheckWebsite_Error:

If cCallado = "" Then
    cErrMsg = "REVISE INTERNET " & vbCrLf & "Error ( " & Err.Number & ")" & vbCrLf & Err.description & ""
    
    'INFO: 21DIC2024
    Rem SI ESTA FUERA DE LINEA, SE APAGAN LOS MENSAJES, SOLO SE CAMBIA EL COLOR Y SE PROCEDE NORMALMENTE
    Rem ShowMsg cErrMsg, vbBlue, vbCyan
    
    EscribeLog "NO HAY INTERNET " & cErrMsg
Else
    Rem ~~~~~~ NADA ~~~~~~
End If
    Check_FE_Website = False
    FE_Facil_ONLINE = False

   On Error GoTo 0
   Exit Function

End Function


'---------------------------------------------------------------------------------------
' Procedure : FE_PreCheckRecordSet
' Author    : hsequeira
' Date      : 05/09/2023
' Purpose   : REVISA EL RECORDSET ANTES DE FACTURAR PARA:
' DESCUENTOS LINEALES
'---------------------------------------------------------------------------------------
'
Public Function FE_PreCheckRecordSet(cSQL As String) As Boolean
'Dim cSQL As String
Dim canalisis() As String
Dim nLineaQDaDescuento As Integer
Dim rsTABLA As ADODB.Recordset
Dim miCuenta As Integer

Set rsTABLA = New ADODB.Recordset
rsTABLA.Open cSQL, msConn, adOpenStatic, adLockReadOnly

On Error Resume Next
rsTABLA.MoveFirst
On Error GoTo 0

Do While Not rsTABLA.EOF
    If Left(rsTABLA!Tipo, 3) = "BDC" Then      'ES UN ITEM QUE RECIBE UN DESCUENTO
        canalisis = Split(rsTABLA!Tipo, " ")
        nLineaQDaDescuento = canalisis(1)
        nFEDescuento = GetDescuentoLinealFE(rsTABLA!precio_unit, rsTABLA!CANT, nLineaQDaDescuento)
        If IsEmpty(nCta) Then miCuenta = 0 Else miCuenta = nCta
        msConn.BeginTrans
        msConn.Execute "UPDATE TMP_TRANS SET FE_DESCUENTO = " & nFEDescuento & " WHERE MESA = " & nMesa & " AND CUENTA = " & miCuenta & " AND LIN = " & rsTABLA!LIN
        msConn.CommitTrans
    End If
        
    If Left(rsTABLA!Tipo, 2) = "DC" Then    'ES UN ITEM QUE DA UN DESCUENTO
    End If
    'Debug.Print rsTABLA!LIN & " - " & rsTABLA!TIPO
    rsTABLA.MoveNext
Loop
'Debug.Print "============================"
rsTABLA.Close
Set rsTABLA = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetDescuentoLinealFE
' Author    : hsequeira
' Date      : 06/09/2023
' Update    : 27/10/2023
' Purpose   : CALCULA (SI HAY EL DESCUENTO LINEAL) EL DESCUENTO << UNITARIO >>, NO EL DESCUENTO TOTAL DE LA LINEA
' CAMPO NUEVO DE LA DB
'---------------------------------------------------------------------------------------
'
Private Function GetDescuentoLinealFE(elPrecio As Single, laCant As Integer, laLinea As Integer) As Single
Dim cSQL As String
Dim rsLaTabla As ADODB.Recordset
Dim miCuenta As Integer

If IsEmpty(nCta) Then miCuenta = 0 Else miCuenta = nCta
cSQL = "SELECT * FROM TMP_TRANS WHERE MESA = " & nMesa & " AND CUENTA = " & miCuenta & " AND LIN = " & laLinea
Set rsLaTabla = New ADODB.Recordset
rsLaTabla.Open cSQL, msConn, adOpenStatic, adLockReadOnly
'GetDescuentoLinealFE = (elPrecio * laCant) * rsLaTabla!DESCUENTO
GetDescuentoLinealFE = elPrecio * rsLaTabla!DESCUENTO
rsLaTabla.Close
Set rsLaTabla = Nothing
End Function

Public Sub ClearCliente()
    tipoClienteFE = ""
    tipoContribuyente = 0
    numeroRUC = ""
    DVRUC = ""
    razonSocial = ""
    direccion = ""
    pais = ""
        provincia = ""
        codigoUbicacion = "8-8-1"
    distrito = ""
    corregimiento = ""
    correoElectronico1 = ""
    telefono1 = ""
    paisExtranjero = ""

End Sub

Public Function WebValidaRUC(cRUC As String, cTipoJuridicoNatural As String) As String
    Dim response As String
    Dim http As Object
    Dim cRUCRetorno As String
    Dim razonRetorno As String
    Dim url As String
    Dim cTipo As String
    Dim DVRetorno As String
    Dim cAmbiente As String
    Dim baseURL As String
    
    
    Select Case cTipoJuridicoNatural
        Case "1"
            cTipo = "N"
        Case Else
            cTipo = "J"
    End Select
    
    Select Case mi_Ambiente
        Case 1
            cAmbiente = "prod"
            baseURL = "https://api.facturafacil.com.pa/api/gateway/wp/"
        Case 2
            cAmbiente = "test"
            baseURL = "https://backend-qa-api.facturafacil.com.pa/api/gateway/wp/"
    End Select
    
    razonRetorno = ""
    'On Error GoTo TestRequest_Error

    Set http = CreateObject("MSXML2.ServerXMLHTTP")

    url = baseURL & cAmbiente & "/checkRUC/" & X_FF_API_Key & "/" & cRUC & "/" & cTipo
    
    Clipboard.SetText url
    
    http.Open "GET", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send

    ' Check if the request was successful (status code 200)
    If http.Status = 200 Then
        response = http.responseText
        ''Debug.Print http.responseText
        cRUCRetorno = ExtractValueFromJSON3(response, "ruc")
        Rem razonRetorno = ExtractValueFromJSON3(response, "razonSocial")
        'INFO: UPDATE 31ENE2026
        'REMOVE APOSTROFE
        razonRetorno = ExtractValueFromJSON3(response, "razonSocial")
        razonRetorno = Replace(razonRetorno, "'", "´")
        DVRetorno = ExtractValueFromJSON3(response, "dv")
        WebValidaRUC = cRUCRetorno & "|" & razonRetorno & "|" & DVRetorno
    Else
        'Form1.lstLog.AddItem "FAIL http.send jsonContent" & Time
        response = "Error: " & http.Status & " - " & http.statusText & " + " & http.responseText
        'ShowMsg "FACTURA ERRONEA: " & vbCrLf & "Motivo Error: " & http.Status & " - " & http.statusText & " + " & http.responseText, vbBlue, vbRed
        'TestRequest = False
        ' Debug.Print http.Status & " - " & http.statusText & " + " & http.responseText
        'Debug.Print response
        WebValidaRUC = razonRetorno
        Exit Function
    End If
End Function

Private Function VerificaSI_es_Propina(cTipoPago As String) As Boolean
    cPagoName = Left(cTipoPago, 7)
    Select Case cPagoName
        Case "PROPINA", "PROP SU", "PROP. S", "SERVICI", "SERV SU", "SERV. S", "SERV IN", "SERV. I", "TIP SUG"
            VerificaSI_es_Propina = True
        Case Else
            'Call EscribeLog("EL TEXTO DE PROPINA EN BASE DE DATOS ES INVALIDO, DEBE CAMBIARLO POR UNO VALIDO")
            VerificaSI_es_Propina = False
    End Select

End Function


'---------------------------------------------------------------------------------------
' Procedure : Send_FE_Mail
' Author    : hsequeira
' Date      : 03/11/2024
' Purpose   : ENVIA EL EMAIL DE NOTIFICACION DE LA NC
'---------------------------------------------------------------------------------------
'
Public Function Send_FE_Mail(cMailReceptor As String, cCUFE As String, cID As String) As Boolean
Dim xmlhttp As Object
Dim parametros  As String
Dim url As String
Dim startTicks As Long
Dim endTicks As Long

If FE_Facil_ONLINE Then

    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
    
    Select Case mi_Ambiente
        Case 1
            cAmbiente = "prod"
            baseURL = "https://api.facturafacil.com.pa/api/pac/reception_fe/"
        Case 2
            cAmbiente = "test"
            baseURL = "https://backend-qa-api.facturafacil.com.pa/api/pac/reception_fe/"
    End Select
    
    url = baseURL & cCUFE & "/send_email/"
    
    parametros = "{""email"": """ & cMailReceptor & """}"
    
    xmlhttp.Open "POST", url, False
    xmlhttp.setRequestHeader "accept", "application/json"
    xmlhttp.setRequestHeader "X-FF-Company", X_FF_Company
    xmlhttp.setRequestHeader "X-FF-API-Key", X_FF_API_Key
    xmlhttp.setRequestHeader "Content-Type", "application/json"
    
    Sleep 950
    xmlhttp.send parametros
    
    If xmlhttp.Status = 200 Then
        Send_FE_Mail = True
    Else
        Send_FE_Mail = False
    End If
Else
    Send_FE_Mail = True
End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : StoreJSON_In_OS
' Author    : hsequeira
' Date      : 22/04/2024
' Purpose   : CREA UN JSON CON LOS DATOS DE LA FACTURA ORIGINAL
'PARA EN EL CASO DE QUE SE HAGA LA NOTA DE CREDITO (N.C.), UNICAMENTE HAY QUE AGREGAR LO SIGUIENTE:
'
'============================
'A) N.C. SI ES DE UNA FACTURA EXISTENTE
'============================
'    "type": "04",     <==== SE INDICA 04 PARA INDICAR QUE EL JSON ES UNA NC DE UNA FACTURA EXISTENTE
'    "referred": {                                          <==== DATOS NECESARIOS PARA INDICAR LA FACTURA ORIGINAL
'        "fd_number": "[CUFE O NUMERO]",       <==== CUFE DE LA FACTUA, ESTA EN TABLA TRANSAC_FISCAL
'        "fd_date": "2024-04-17"                      <==== FECHA QUE SE ESTA HACIENDO LA N.C.
'    },                                                        <==== SE CIERRA EL BRACKET DE "referred"  "
'    "info": "Information of interest to be included in signed document"

'================================
'B) N.C. GENERICA, SIN REFERENCIA A FACTURA
'================================
'
'    "type": "06",                                          <==== NADA MAS SE INDICA 06, TODO LO QUE ES EN EL JSON COMPONE LA N.C.
'    "info": "Information of interest to be included in signed document"
'  }

'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
' Procedure : StoreJSON_In_OS
' Author    : hsequeira
' Date      : 03/11/2024
' Purpose   : GUARDA LA NOTA DE CREDITO (NC#####.TXT)  EN EL SISTEMA OPERATIVO
Rem EL ##### INDICA EL NUMERO DE FACTURA
'---------------------------------------------------------------------------------------
'
Public Sub StoreJSON_In_OS(nFactura As Long, cJSONCadena As String, Optional cNC As String)
Dim nFreefile As Byte

On Error Resume Next

nFreefile = FreeFile()

Open App.Path & "\HIST\" & cNC & LTrim(RTrim(Str(nFactura))) & ".txt" For Output As #nFreefile
Print #nFreefile, cJSONCadena
Close #nFreefile
On Error GoTo 0
End Sub


Public Function Public_TestRequest(jsonContent As String, Optional TipoRequest As String, Optional la_factura_fue_hecha_off_line As String) As Boolean

If la_factura_fue_hecha_off_line = "OFF-LINE" Then
    Public_TestRequest = TestRequest_OFFLINE(jsonContent, TipoRequest)
Else
    'INFO: SI VIENE EN BLANCO, FUE HECHA ON-LINE
    Public_TestRequest = TestRequest(jsonContent, TipoRequest)
End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : CHECK_NOTAS_DE_CREDITO_FE
' Author    : hsequeira
' Date      : 03/11/2024
' Purpose   : REVISA SI HAY VALORES EN NOTA DE CREDITO
Rem SI LO HAY , LO IMPRIME Y SI ES EL REPORTE Z LOS BORRA
'---------------------------------------------------------------------------------------
'
Public Function CHECK_NOTAS_DE_CREDITO_FE(bBorrar As Boolean, n1 As Long, n2 As Single) As Boolean

Dim nCurrenCounter As Long
Dim nCurrentValor As Single
Dim cTemp As String

cTemp = RegRead("HKCU\Software\SoloSoftware\SoloMix\NCCounter")
If cTemp = "" Then
    CHECK_NOTAS_DE_CREDITO_FE = False
Else
    CHECK_NOTAS_DE_CREDITO_FE = True
    n1 = CLng(cTemp)
    cTemp = RegRead("HKCU\Software\SoloSoftware\SoloMix\NCValor")
    n2 = CSng(cTemp)
    If bBorrar Then
        Rem SOLO SI ES REPRTE Z
        RegWrite "HKCU\Software\SoloSoftware\SoloMix\NCCounter", 0
        RegWrite "HKCU\Software\SoloSoftware\SoloMix\NCValor", 0
        RegWrite "HKCU\Software\SoloSoftware\SoloMix\NCFechaRZ", Format(Date, "YYYYMMDD") & Space(1) & Format(Time(), "HH:mm:ss")
    Else
    End If
End If
Rem cTemp = CSng(cMonto)
End Function


'Private Function AI_Examine_SOAP_Response(cSOAPResponse As String)
'
'' Import the MSXML2 library
'Dim doc As Object
'
'' Load the XML data
'Set doc = CreateObject("MSXML2.DOMDocument.6.0")
'doc.LoadXml cSOAPResponse ' Replace "xmlString" with the actual XML string
'
'' Get the root element
'Dim rootNode As Object
'Set rootNode = doc.documentElement
'
'' Define variables to store the values
'Dim cufe As String
'Dim qr_code_data As String
'Dim pdf_url As String
'Dim ctype As String
'Dim code As String
'Dim message As String
'
'' Extract cufe value
'cufe = rootNode.SelectSingleNode("dId").Text
'
'' Extract qr_code_data (assuming it's the base64 encoded QR code within dQRCode)
'qr_code_data = rootNode.SelectSingleNode("gNoFirm/dQRCode").Text
'
'' pdf_url is not directly available in the provided XML schema
'
'' Extract type (assuming it's iTipoEmision within gDGen)
'ctype = rootNode.SelectSingleNode("gDGen/iTipoEmision").Text
'
'' Extract code (assuming it's an error code, need more context to determine the exact element)
'' Without more information, assuming there's no error code element
'code = "" ' Modify this line if there's a specific error code element
'
'' Extract message (assuming it's associated with the code, need more context to determine the exact element)
'message = "" ' Modify this line if there's a specific message element
'
'' Display or store the extracted values
'Debug.Print "cufe: " & cufe
'Debug.Print "qr_code_data: " & qr_code_data
'Debug.Print "pdf_url: " & pdf_url ' This will be blank
'Debug.Print "type: " & ctype
'Debug.Print "code: " & code
'Debug.Print "message: " & message

' You can modify this code to store the values in variables or perform other operations


    ' Declare MSXML2.DOMDocument object for parsing XML
'    Dim xmlDoc As Object
'    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
'    xmlDoc.async = False
'    xmlDoc.validateOnParse = False
'
'    ' Replace this with your actual JSON/XML string
'    Dim jsonResponse As String
'    jsonResponse = cSOAPResponse
'
'    ' Extract XML data from JSON
'    Dim startPos As Long, endPos As Long
'    startPos = InStr(jsonResponse, "<SOAP-ENV:Envelope")
'    endPos = InStr(startPos, jsonResponse, "</SOAP-ENV:Envelope>") + Len("</SOAP-ENV:Envelope>")
'    Dim xmlData As String
'    xmlData = Mid(jsonResponse, startPos, endPos - startPos + 1)
'
'    ' Load XML data into XML DOM object
'    If Not xmlDoc.LoadXml(xmlData) Then
'        MsgBox "Failed to parse XML.", vbCritical
'        Exit Function
'    End If
'
'    ' Extract values
'    Dim cufe As String, qrCodeData As String, pdfUrl As String
'    Dim xmlQrCodeDataNode As Object
'    Dim xmlCufeNode As Object
'    Dim xmlMessageNode As Object
'    Dim xmlCodeNode As Object
'    Dim xmlTypeNode As Object
'
'    ' Get CUFE
'    Set xmlCufeNode = xmlDoc.SelectSingleNode("//ns2:dCUFE")
'    If Not xmlCufeNode Is Nothing Then
'        cufe = xmlCufeNode.Text
'    End If
'
'    ' Get QR Code Data
'    qrCodeData = GetJSONValue(jsonResponse, "qr_code_data")
'
'    ' Get PDF URL
'    pdfUrl = GetJSONValue(jsonResponse, "pdf_url")
'
'    ' Get type
'    Dim typeValue As String
'    typeValue = GetJSONValue(jsonResponse, "type")
'
'    ' Get code
'    Dim codeValue As String
'    codeValue = GetJSONValue(jsonResponse, "code")
'
'    ' Get message
'    Dim messageValue As String
'    messageValue = GetJSONValue(jsonResponse, "message")
'
'    ' Display values
'    MsgBox "CUFE: " & cufe & vbCrLf & '           "QR Code Data: " & qrCodeData & vbCrLf & '           "PDF URL: " & pdfUrl & vbCrLf & '           "Type: " & typeValue & vbCrLf & '           "Code: " & codeValue & vbCrLf & '           "Message: " & messageValue

'End Function

' Helper function to get a value from the JSON structure
'Private Function GetJSONValue(jsonStr As String, key As String) As String
'    Dim startPos As Long, endPos As Long, result As String
'    startPos = InStr(jsonStr, """" & key & """:""") + Len(key) + 4
'    endPos = InStr(startPos, jsonStr, """")
'    If startPos > 0 And endPos > 0 Then
'        result = Mid(jsonStr, startPos, endPos - startPos)
'    End If
'    GetJSONValue = result
'End Function



Rem dead end
    ' Define necessary objects and variables
'    Dim xmlDoc As Object
'    Dim soapXML As String
'    Dim cufe As String
'    Dim qr_code_data As String
'    Dim pdf_url As String
'    Dim code As String
'    Dim message As String
'
'    ' Assign the XML content
'    soapXML = cSOAPResponse
'
'    ' Load XML Document
'    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
'    xmlDoc.async = False
'    xmlDoc.LoadXml soapXML
'
'    ' Check if XML is loaded correctly
'    If xmlDoc.parseError.errorCode <> 0 Then
'        MsgBox "Error in loading XML: " & xmlDoc.parseError.reason
'        Debug.Print xmlDoc.parseError.reason
'        Debug.Print xmlDoc.parseError.errorCode
'        Exit Function
'    End If
'
'    ' Get the CUFE
'    cufe = xmlDoc.selectSingleNode("//ns2:dCUFE").Text
'
'    ' Extract code and message
'    code = xmlDoc.selectSingleNode("//ns2:dCodRes").Text
'    message = xmlDoc.selectSingleNode("//ns2:dMsgRes").Text
'
'    ' Print values
'    MsgBox "CUFE: " & cufe & vbCrLf & "Code: " & code & vbCrLf & "Message: " & message
'
'    ' Clean up
'    Set xmlDoc = Nothing





'Dim xmlDoc As Object
''Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
'Set xmlDoc = CreateObject("MSXML2.DOMDocument.4.0")
'
'xmlDoc.async = False
'xmlDoc.validateOnParse = False
'
'' Sample XML (use the entire XML content here, ensure the XML is valid and properly encoded)
'Dim xmlContent As String
'xmlContent = cSOAPResponse
'
'' Attempt to load the XML
'If xmlDoc.LoadXml(xmlContent) Then
'    MsgBox "XML loaded successfully."
'
'    ' Read the nodes you need
'    Dim cufe As String
'    Dim qr_code_data As String
'    Dim pdf_url As String
'    Dim ctype As String
'    Dim code As String, message As String
'
'    cufe = xmlDoc.SelectSingleNode("//dCUFE").Text
'    qr_code_data = xmlDoc.SelectSingleNode("//qr_code_data").Text
'    pdf_url = xmlDoc.SelectSingleNode("//pdf_url").Text
'    ctype = xmlDoc.SelectSingleNode("//type").Text
'    code = xmlDoc.SelectSingleNode("//dCodRes").Text
'    message = xmlDoc.SelectSingleNode("//dMsgRes").Text
'
'    ' Display the values
'    MsgBox "CUFE: " & cufe & vbCrLf & "QR Code Data: " & qr_code_data & vbCrLf & '           "PDF URL: " & pdf_url & vbCrLf & "Type: " & ctype & vbCrLf & '           "Code: " & code & vbCrLf & "Message: " & message
'Else
'    ' Display error information
'    Debug.Print "Error loading XML: " & xmlDoc.parseError.reason & "Error Code: " & xmlDoc.parseError.errorCode
'End If
















'Dim xmlDoc As Object
' Set xmlDoc = CreateObject("MSXML2.DOMDocument")
'Rem Set xmlDoc = CreateObject("MSXML2.ServerXMLHTTP")
'
'
'' Load the SOAP response into the XML document
'Dim soapResponse As String
'soapResponse = cSOAPResponse
'xmlDoc.LoadXml soapResponse
'
'If xmlDoc.parseError.errorCode <> 0 Then
'    MsgBox "Error parsing SOAP response: " & xmlDoc.parseError.reason
'    Exit Function
'End If
'
'' Get CUFE
'Dim cufe As String
'cufe = xmlDoc.SelectSingleNode("//ns2:dCUFE").Text
'Debug.Print "CUFE: " & cufe
'
'' Get qr_code_data
'Dim qrCodeData As String
'qrCodeData = xmlDoc.SelectSingleNode("//qr_code_data").Text
'Debug.Print "QR Code Data: " & qrCodeData
'
'' Get pdf_url (if available in the SOAP response)
'Dim pdfURL As String
'pdfURL = xmlDoc.SelectSingleNode("//pdf_url").Text
'Debug.Print "PDF URL: " & pdfURL
'
'' Get type (this would depend on the structure, placeholder here)
'Dim messageType As String
'messageType = xmlDoc.SelectSingleNode("//type").Text
'Debug.Print "Type: " & messageType
'
'' Get code
'Dim code As String
'code = xmlDoc.SelectSingleNode("//ns2:dCodRes").Text
'Debug.Print "Code: " & code
'
'' Get message
'Dim message As String
'message = xmlDoc.SelectSingleNode("//ns2:dMsgRes").Text
'Debug.Print "Message: " & message
'
'Set xmlDoc = Nothing

'End Function


Private Function ExtractGTot(xmlString As String, cLookUpValue As String) As String
    Dim gTotStartTag As String
    Dim gTotEndTag As String
    Dim startPos As Long
    Dim endPos As Long
    Dim gTotValue As String
    
    ' Define the start and end tags for gTot
    gTotStartTag = "<" & cLookUpValue & ">"
    gTotEndTag = "</" & cLookUpValue & ">"
    
    ' Find the starting position of the gTot value
    startPos = InStr(1, xmlString, gTotStartTag, vbTextCompare)
    If startPos > 0 Then
        ' Adjust startPos to point to the start of the value
        startPos = startPos + Len(gTotStartTag)
        
        ' Find the ending position of the gTot value
        endPos = InStr(startPos, xmlString, gTotEndTag, vbTextCompare)
        If endPos > 0 Then
            ' Extract the gTot value
            gTotValue = Mid(xmlString, startPos, endPos - startPos)
        Else
            gTotValue = "Error: End tag not found"
        End If
    Else
        gTotValue = "Error: Start tag not found"
    End If
    
    ' Return the extracted value
    ExtractGTot = gTotValue
End Function

Private Function ExtractFromSOAP(xmlString As String, cLookUpValue As String) As String
    
    Dim startTag As String
    Dim endTag As String
    Dim startPos As Long
    Dim endPos As Long
    Dim dFecProc As String
    
    ' Define the start and end tags for dFecProc
    startTag = "<ns2:" & cLookUpValue & ">"
    endTag = "</ns2:" & cLookUpValue & ">"

    ' Find the starting position of dFecProc
    startPos = InStr(xmlString, startTag)
    If startPos > 0 Then
        ' Adjust the start position to the beginning of the value
        startPos = startPos + Len(startTag)

        ' Find the ending position of dFecProc
        endPos = InStr(startPos, xmlString, endTag)
        If endPos > 0 Then
            ' Extract the value of dFecProc
            dFecProc = Mid(xmlString, startPos, endPos - startPos)

            ' Display the extracted value
            ExtractFromSOAP = dFecProc
        Else
            ExtractFromSOAP = ""
        End If
    Else
        ExtractFromSOAP = ""
    End If
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : GenerateJson_OFFLINE
' Author    : hsequeira
' Date      : 17/12/2024
' Purpose  : HACER LA FACTURAS FUERA DE LINEA
' BASADO EN LOCALHOST:
'---------------------------------------------------------------------------------------
'
Public Function GenerateJson_OFFLINE(nEnvironment As Integer, nFactura As Long, RSPAGOS As ADODB.Recordset, _
                                                        lstPagos As MSHFlexGrid, Exonerado As Boolean, mi_total_factura As String, _
                                                        Optional rsParamPropina As ADODB.Recordset) As Boolean
    Dim JsonString As String
    Dim header As String
    Dim receptor As String
    Dim items As String
    Dim payments As String
    Dim docType As String
    Dim INFO As String
    Dim cClienteDir As String
    Dim future_ObtenerPagado As Single
    
        
    Dim nFreefile As Byte
    Dim a As String
    
   On Error GoTo GenerateJson_Error

    AllowFE_to_Continue = False

    Call Load_FE_Mesa
    
    If rsFETrans.RecordCount = 0 Then
        AllowFE_to_Continue = True
        GenerateJson_OFFLINE = False
        Exit Function
    End If
    
    If tipoClienteFE = "" Then tipoClienteFE = "02"
    If tipoContribuyente = 0 Then tipoContribuyente = 1
    
    cClienteDir = corregimiento & "," & distrito & "," & provincia
    If cClienteDir = ",," Then cClienteDir = "ESCRIBIR DIRECCION"
    If razonSocial = "" Then razonSocial = "CONSUMIDOR FINAL"
    If numeroRUC = "" Then numeroRUC = "0-0-0"
    If correoElectronico1 = "" Then correoElectronico1 = "eMail"
    
    '(1)
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem header = GenerateHeader_OFFINE(o_X_FF_Company, o_X_FF_Branch, "001", "F", nFactura, razonSocial, "07", cClienteDir)
    header = GenerateHeader_OFFINE(o_X_FF_Company, o_X_FF_Branch, o_poscode, "F", nFactura, razonSocial, "07", cClienteDir)
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'receptor = GenerateReceptor(tipoClienteFE, razonSocial, LTrim(RTrim(Str(tipoContribuyente))), cClienteDir, correoElectronico1, numeroRUC, DVRUC, paisExtranjero)
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    '(2)
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~
    items = GenerateItems_OFFLINE(LTrim(RTrim(Str(tipoContribuyente))), Exonerado)
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    Set rsFEPagos = RSPAGOS.Clone(adLockOptimistic)
    
    '|||||||||||||||||||||||||||||||||||||
    'INFO: UPDATE 9JUL2024
    '|||||||||||||||||||||||||||||||||||||
    local_total_factura = Format(local_total_factura, "0.00") '' // REDONDEANDO A DOS DECIMALES  134.115
    mi_total_factura = Format(local_total_factura, "#.00")
    'payments = GeneratePayments(lstPagos, future_ObtenerPagado)
    
    'INFO: FIX RSPAGOS EN VEZ DE LISTPAGOS
    '10SEP2024
    If Left(mi_App_Name, 7) = "SoloMix" Then
        '(3a)
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem payments = GeneratePayments_OFFLINE(lstPagos, future_ObtenerPagado)
        payments = Generate_NEW_Payments_OFFLINE(RSPAGOS, future_ObtenerPagado)
        'INFO: UPDATE 25ENE2025. CAMBIANDO A NUEVO FORMATO
        Rem payments = Generate_NEW_Payments_OFFLINE(RSPAGOS, future_ObtenerPagado)
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Else
        '(3b)
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        payments = Generate_NEW_Payments_OFFLINE(RSPAGOS, future_ObtenerPagado)
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    End If
    
    docType = "01"
    mi_total_factura = LTrim(RTrim(CStr(mi_total_factura)))
    
    If Left(mi_App_Name, 7) = "SoloMix" Then
        On Error Resume Next
        rsParamPropina.MoveFirst
        On Error GoTo 0
    
        Rem If rsFEPropinas.EOF Then
        If rsParamPropina.EOF Then
            INFO = "<string>"
        Else
            Do While Not rsParamPropina.EOF
                INFO = INFO & rsParamPropina!DESCRIP_1 & Space(2) & Format(rsParamPropina!MONTO_1, "STANDARD") & " // "
                rsParamPropina.MoveNext
            Loop
        End If
    Else
        INFO = "<string>"
    End If
    
    If Exonerado Then
        INFO = INFO & "CLIENTE EXONERADO //"
    End If
    
    ' Build the complete JSON
    'INCLUIR mi_total_factura
    'jsonString = "{" & header & ",""document"": {""fd_number"": " & nFactura & "," & receptor & ",""items"": " & items & ",""payments"": " & payments & ",""total"":""" & mi_total_factura & """,""type"": """ & docType & """,""info"": """ & INFO & """}}"
    Rem jsonString = "{" & header & ",""items"": " & items & ",""payments"": " & payments & ",""total"":""" & mi_total_factura & """,""type"": """ & docType & """,""info"": """ & INFO & """}"
    JsonString = "{" & header & """items"": " & items & ",""payments"": " & payments & "}}"

    If TestRequest_OFFLINE(JsonString) Then
        'ShowMsg "ok"
        GenerateJson_OFFLINE = True
        Call StoreJSON_In_OS(nFactura, JsonString)
    Else
        GenerateJson_OFFLINE = False
    End If

   On Error GoTo 0
   Exit Function

GenerateJson_Error:
    GenerateJson_OFFLINE = False
    ShowMsg "Error " & Err.Number & " (" & Err.description & ") en GenerateJson_OFFLINE de facilelectronica"
    EscribeLog "FE.GenerateJson_OFFLINE: " & "Error " & Err.Number & " (" & Err.description & ") "
   Resume
End Function

'---------------------------------------------------------------------------------------
' Procedure : GenerateHeader_OFFINE
' Author    : hsequeira
' Date      : 17/12/2024
' Purpose   : GENERA EL ENCABEZADO DE FACIL FACTURA
'header = GenerateHeader_OFFINE(o_X_FF_Company, o_X_FF_Branch, o_X_FF_Branch, "F", nFactura, razonSocial, "07", cClienteDir)
'---------------------------------------------------------------------------------------
'
Private Function GenerateHeader_OFFINE(xCompa As String, xBranch As String, xPuntoVenta As String, xTipoDoc As String, xFactura As Long, _
                                                            xCustomer As String, xCustomerType As String, xDireccion As String) As String


'    Dim location As String
'    location = "8-8-2"
'    Select Case tipe
'        Case "01"      'Contribuyente
'            GenerateReceptor = """receptor"": {" & """type"": """ & tipe & """,""name"": """ & name & """,""ruc_type"": " & rucType & "," & """address"": """ & address & """,""email"": """ & email & """,""ruc"": """ & ruc & """,""location"": """ & location & """,""dv"": """ & dv & """}"
'        Case "02"       'Consumidor FINAL
'            If rucType = "1" Then  'PERSONA NATURAL
'                GenerateReceptor = """receptor"": {" & """type"": """ & tipe & """,""name"": """ & name & """,""ruc_type"": " & rucType & "," & """address"": """ & address & """,""email"": """ & email & """,""ruc"": """ & ruc & """,""dv"": """ & dv & """}"
'            Else                        'JURIDICO (EMPRESA, GOBIERNO)
'                GenerateReceptor = """receptor"": {" & """type"": """ & tipe & """,""name"": """ & name & """,""ruc_type"": " & rucType & "," & """address"": """ & address & """,""email"": """ & email & """,""ruc"": """ & ruc & """,""dv"": """ & dv & """}"
'            End If
'        Case "03"       'Gobierno
'            GenerateReceptor = """receptor"": {" & """type"": """ & tipe & """,""name"": """ & name & """,""ruc_type"": " & rucType & "," & """address"": """ & address & """,""email"": """ & email & """,""ruc"": """ & ruc & """,""location"": """ & location & """,""dv"": """ & dv & """}"
'        Case "04"       'Extranjero
'            GenerateReceptor = """receptor"": {" & """type"": """ & tipe & """,""name"": """ & name & """,""ruc_type"": " & rucType & "," & """address"": """ & address & """,""email"": """ & email & """,""ruc"": """ & ruc & """,""country"": """ & country & """,    ""dv"": """ & dv & """}"
'    End Select

    'INFO: UPDATE 12MAR2026
    'SE AGREGA CAJERO Y MESERO
    GenerateHeader_OFFINE = """fiscalDoc"": {" & """companyLicCod"": """ & xCompa & """,""branchCod"": """ & xBranch & """,""posCod"": """ & xPuntoVenta & _
            """,""docType"": """ & xTipoDoc & """,""docNumber"": """ & xFactura & """,""customerName"": """ & xCustomer & _
            """,""Cajero"": """ & (npNumCaj) & """,""Mesero"": """ & nMesero & _
            """,""CajaID"": """ & (rs00!DESCRIP) & """,""CajaDir"": """ & rs00!direccion & _
            """,""customerPhone"": """ & xCustomer & """,""customerRUC"": """ & xCustomer & """,""customerType"": """ & xCustomerType & """,""customerAddress"": """ & xDireccion & ""","
            Rem """,""customerPhone"": """ & xCustomer & """,""customerRUC"": """ & xCustomer & """,""customerType"": """ & xCustomerType & """,""customerAddress"": """ & xDireccion & """}"
End Function



'---------------------------------------------------------------------------------------
' Procedure : GenerateItems_OFFLINE
' Author    : hsequeira
' Date      : 22/03/2024
' Purpose   : GENERA LA LINEA DE ITEMS CON SU RESPECTIVO BRACKET DE TAXES
' Date      : 03/06/2024
' Purpose   : se incluye la cantidad para totalizar la factura
'---------------------------------------------------------------------------------------
'
Private Function GenerateItems_OFFLINE(tipe As String, Exonerado As Boolean) As String
    'tipe = 01=Contribuyente, 02=Consumidor final, 03=Gobierno, 04=Extranjero
    Dim items As String
    Dim itemLine As Integer
    Dim nRegistros As Integer
    Dim nCodeTax As Integer
    Dim nLocalTotalFactura As Single
    Dim nLocalTax As Single
    Dim nTotalLinea As Single
    
'    "items": [
'      {
'        "id": 1,
'        "qty": "14.0",
'        "price": "12.5",
'        "desc": "CARGO CONSUMIDOR FINAL",
'        "tax": "2",
'        "damt": "1.88",
'        "dperc": "15%"
'      }
'    ],
    
    
    
    nRegistros = rsFETrans.RecordCount
    
    items = "["
    
    On Error Resume Next
    rsFETrans.MoveFirst
    On Error GoTo 0
    
    For itemLine = 1 To nRegistros
        
        Dim itemTaxes As String
        
        If Exonerado Then
            nCodeTax = 0
        Else
            Select Case rsFETrans!CON_TAX
                Case 10
                    nCodeTax = 2
                Case 7
                    nCodeTax = 1
                Case 0
                    nCodeTax = 0
                Case Else
                    nCodeTax = 3
            End Select
        End If
        
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        '18FEB2025
        'INFO: CALCULO DE IMPUESTO
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        If Exonerado Then
            itemTaxes = "[{""type"": """ & rsFETrans!Type & """, ""amount"": " & 0# & ", ""code"": """ & cCodeTax & """}]"
            
        Else
            'INFO: PASAR EL VALOR DE DESCUENTO SIN REDONDEAR
            '2ABRIL2024
            'itemTaxes = "[{""type"": """ & rsFETrans!Type & """, ""amount"": " & Format(rsFETrans!amount, "STANDARD") & ", ""code"": """ & cCodeTax & """}]"
            'itemTaxes = "[{""type"": """ & rsFETrans!Type & """, ""amount"": " & Format(rsFETrans!amount, "0.###") & ", ""code"": """ & cCodeTax & """}]"
            itemTaxes = "[{""type"": """ & rsFETrans!Type & """, ""amount"": " & rsFETrans!amount & ", ""code"": """ & cCodeTax & """}]"
            
            'nLocalTotalFactura = nLocalTotalFactura + rsFETrans!amount
            nLocalTax = rsFETrans!amount
        End If
        
        Select Case tipe
            Case "2"   'Gobierno
                nTotalLinea = ((rsFETrans!price * rsFETrans!quantity) - (rsFETrans!DESCUENTO * rsFETrans!quantity)) + nLocalTax
                nLocalTotalFactura = nLocalTotalFactura + nTotalLinea
                
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                Rem items = items & GenerateItem_OFFLINE(rsFETrans!LIN, rsFETrans!price, "und", rsFETrans!quantity, rsFETrans!description, itemTaxes, rsFETrans!DESCUENTO, _
                                                                        rsFETrans!internal_code, tipe, "9010") & ","
                items = items & GenerateItem_OFFLINE(rsFETrans!LIN, rsFETrans!quantity, rsFETrans!price, rsFETrans!description, nCodeTax, rsFETrans!DESCUENTO, tipe) & ","
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Case Else
                nTotalLinea = ((rsFETrans!price * rsFETrans!quantity) - (rsFETrans!DESCUENTO * rsFETrans!quantity)) + nLocalTax
                nLocalTotalFactura = nLocalTotalFactura + nTotalLinea
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                items = items & GenerateItem_OFFLINE(rsFETrans!LIN, rsFETrans!quantity, rsFETrans!price, rsFETrans!description, nCodeTax, rsFETrans!DESCUENTO, tipe) & ","
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        End Select
        rsFETrans.MoveNext
    Next itemLine
    
    ' Remove the trailing comma
    If Right(items, 1) = "," Then
        items = Left(items, Len(items) - 1)
    End If
    
    GenerateItems_OFFLINE = items & "]"
    If local_total_factura <> nLocalTotalFactura Then local_total_factura = nLocalTotalFactura
End Function

'---------------------------------------------------------------------------------------
' Procedure : GenerateItem_OFFLINE
' Author    : hsequeira
' Date      : 12/03/2024
' update: 27/04/2024
' NEW  17DIC2024
' Purpose   : Genera cada linea de Items de la factura
'---------------------------------------------------------------------------------------
Private Function GenerateItem_OFFLINE(Line As Integer, quantity As Integer, price As Single, description As String, _
                                                        taxes As Integer, discount As Single, Optional ctipe As String, Optional internalCode As Long, _
                                                        Optional gns As String) As String
    Dim cLocalDiscount As String
    Dim cLocalPrice As String
    
    cLocalDiscount = Format(discount, "#.000")
    cLocalPrice = Format(price, "#.000")
    
    
    Rem nLocalDiscount = Format(discount, "0.0000")
    Rem nLocalDiscount = Format(discount, "0.00") ' SE REDONDEA A 2 DECIMALES
    Rem price = Format(price, "0.000") ' SE REDONDEA A 3 DECIMALES
    
    Select Case ctipe
        Case "2"
            Rem GenerateItem_OFFLINE = "{""id"": " & Line & ",""price"": """ & price & ""","
            GenerateItem_OFFLINE = "{""id"": " & Line & ",""price"": """ & cLocalPrice & ""","
            GenerateItem_OFFLINE = GenerateItem_OFFLINE & """qty"": """ & quantity & """,""desc"": """ & description & ""","
            GenerateItem_OFFLINE = GenerateItem_OFFLINE & """tax"": """ & taxes & """,""damt"": """ & cLocalDiscount & """}"
        Case Else
            Rem GenerateItem_OFFLINE = "{""id"": " & Line & ",""price"": """ & price & ""","
            GenerateItem_OFFLINE = "{""id"": " & Line & ",""price"": """ & cLocalPrice & ""","
            GenerateItem_OFFLINE = GenerateItem_OFFLINE & """qty"": """ & quantity & """,""desc"": """ & description & ""","
            GenerateItem_OFFLINE = GenerateItem_OFFLINE & """tax"": """ & taxes & """,""damt"": """ & cLocalDiscount & """}"
        End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : GeneratePayments_OFFLINE
' Author    : hsequeira
' Date      : 12/03/2024
' Purpose   : Guarda la Cadena con los pagos
' INFO: UPDATE 26MAY2024, SE PONE 1 A nMax, ANTES ERA 0, ENTONCES PONIA 2 PAGOS EN VEZ DE UNO
' 17dic2024
' off line
'---------------------------------------------------------------------------------------
'
Private Function GeneratePayments_OFFLINE(lstPagos As MSHFlexGrid, ByRef nPagosFuturo As Single) As String

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Rem  USAR Generate_NEW_Payments_OFFLINE
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim MasDeUnPago As Boolean
Dim cCadena As String
Dim coma As String
Dim nTipoPago As Integer
Dim cTextoTipoPago As String
Dim nValorPago As Single
Dim nMax As Integer
Dim nFixRows As Integer
Dim i As Integer
Dim cLocalPayment As String



Set rsFEPagos = New ADODB.Recordset

rsFEPagos.Fields.Append "type", adChar, 2
rsFEPagos.Fields.Append "description", adChar, 35
rsFEPagos.Fields.Append "amount", adSingle

rsFEPagos.CursorType = adOpenStatic
rsFEPagos.LockType = adLockOptimistic
rsFEPagos.Open


Set rsFEPropinas = New ADODB.Recordset
rsFEPropinas.Fields.Append "description", adChar, 35
rsFEPropinas.Fields.Append "amount", adSingle

rsFEPropinas.CursorType = adOpenStatic
rsFEPropinas.LockType = adLockOptimistic
rsFEPropinas.Open

If Left(mi_App_Name, 7) = "SoloMix" Then
'INFO: UPDATE 26MAY2024, SE PONE 1 A nMax, ANTES ERA 0, ENTONCES PONIA 2 PAGOS EN VEZ DE UNO
    nMax = 1
Else
    If mi_Panta = "PAGOS" Then
        nMax = 0
    Else
        nMax = 1
    End If
End If

If lstPagos.Rows = 1 Then nFixRows = 0 Else nFixRows = 1

For i = nMax To (lstPagos.Rows - nFixRows)
    On Error Resume Next
    lstPagos.Row = i
    On Error GoTo 0
    lstPagos.Col = 0
    nTipoPago = lstPagos.Text
    'INFO: 5JUN2007
    lstPagos.Col = 1
    cTextoTipoPago = Trim(lstPagos.Text)

    lstPagos.Col = 2
    nValorPago = Format(lstPagos.Text, "STANDARD")

    If VerificaSI_es_Propina(cTextoTipoPago) Then
        rsFEPropinas.AddNew

        rsFEPropinas.Fields("description").value = cTextoTipoPago
        rsFEPropinas.Fields("amount").value = nValorPago
        rsFEPropinas.Update

    Else
        rsFEPagos.AddNew

        rsFEPagos.Fields("type").value = GetPagoElectronico_OFFLINE(nTipoPago)
        rsFEPagos.Fields("description").value = cTextoTipoPago & " / " & cTextoTipoPago
        rsFEPagos.Fields("amount").value = nValorPago
        rsFEPagos.Update
    End If
    Sleep 50

Next

    rsFEPagos.MoveFirst

    If rsFEPagos.RecordCount > 0 Then coma = ","

    i = 1
    Do While Not rsFEPagos.EOF
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        'UPDATE" 21AGO2024
        'SE ARREGLA EL MONTO PAGANDO PARA QUE NO INCLUYA COMAS
        Rem cCadena = cCadena & "{""id"": " & i & ",""type"": """ & rsFEPagos!Type & """,""amt"": " & local_total_factura & ",""desc"": """ & rsFEPagos!description & """}" & coma
        cCadena = cCadena & "{""id"": " & i & ",""type"": """ & rsFEPagos!Type & """,""amt"": " & rsFEPagos!amount & ",""desc"": """ & rsFEPagos!description & """}" & coma
        
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        nPagosFuturo = nPagosFuturo + rsFEPagos!amount
        i = i + 1
        rsFEPagos.MoveNext
    Loop

    If Right(cCadena, 1) = "," Then cCadena = Left(cCadena, Len(cCadena) - 1)
    GeneratePayments_OFFLINE = "[" & cCadena & "]"
   
End Function


'---------------------------------------------------------------------------------------
' Procedure : Generate_NEW_Payments_OFFLINE
' Author    : hsequeira
' Date      : 10/09/2024
' Date      : 9/ene/2025
' Purpose   : GENERA PAGOS DESDE EL RECORSET EN VEZ DEL MSHFlexGrid
'---------------------------------------------------------------------------------------
'
Private Function Generate_NEW_Payments_OFFLINE(rsMisPagos As ADODB.Recordset, ByRef nPagosFuturo As Single) As String
Dim MasDeUnPago As Boolean
Dim cCadena As String
Dim coma As String
Dim nTipoPago As Integer
Dim cTextoTipoPago As String
Dim nValorPago As Single
Dim nMax As Integer
Dim nFixRows As Integer
Dim i As Integer
Dim UnSoloPago As Boolean
Dim nDiferenciaConFE As Single


Dim cLocalPayment As String

rsMisPagos.MoveFirst

If rsMisPagos.RecordCount > 0 Then coma = ","

If rsMisPagos.RecordCount = 1 Then
    UnSoloPago = True
Else
    UnSoloPago = False
End If

If UnSoloPago Then
    cLocalPayment = Format(local_total_factura, "#.00")
    Rem cCadena = cCadena & "{""id"": 1" & ", ""type"": """ & rsMisPagos!CODIGO_FE & """,""amt"": """ & local_total_factura & """,""desc"": """ & rsMisPagos!DESCRIP_FE & """}" & coma
    Rem cCadena = cCadena & "{""id"": 1" & ", ""type"": """ & rsMisPagos!CODIGO_FE & """,""amt"": """ & cLocalPayment & """,""desc"": """ & rsMisPagos!DESCRIP_FE & """}" & coma
    cCadena = cCadena & "{""id"": 1" & ", ""type"": """ & GetPagoElectronico_OFFLINE(rsMisPagos!CODIGO) & """,""amt"": """ & cLocalPayment & """,""desc"": """ & rsMisPagos!DESCRIP_FE & """}" & coma
    
    nPagosFuturo = nPagosFuturo + local_total_factura
Else
    i = 1
    Do While Not rsMisPagos.EOF
        cLocalPayment = Format(rsMisPagos!MONTO_1, "#.00")
        Rem cCadena = cCadena & "{""id"": " & i & ", ""type"": """ & rsMisPagos!CODIGO_FE & """,""amt"":  """ & rsMisPagos!MONTO_1 & """,""desc"": """ & rsMisPagos!DESCRIP_FE & """}" & coma
        Rem cCadena = cCadena & "{""id"": " & i & ", ""type"": """ & rsMisPagos!CODIGO_FE & """,""amt"":  """ & cLocalPayment & """,""desc"": """ & rsMisPagos!DESCRIP_FE & """}" & coma
        cCadena = cCadena & "{""id"": " & i & ", ""type"": """ & GetPagoElectronico_OFFLINE(rsMisPagos!CODIGO) & """,""amt"":  """ & cLocalPayment & """,""desc"": """ & rsMisPagos!DESCRIP_FE & """}" & coma
        nPagosFuturo = nPagosFuturo + rsMisPagos!MONTO_1
        i = i + 1
        rsMisPagos.MoveNext
    Loop
End If

'INFO: 12NOV2024
' AJUSTE A LA FACTURA X LOS DECIMALES QUE CALCULA LA FE
If local_total_factura > nPagosFuturo Then
    Rem nDiferenciaConFE = local_total_factura - nPagosFuturo
    nDiferenciaConFE = Round(local_total_factura - nPagosFuturo, 5)
    cLocalPayment = Format(nDiferenciaConFE, "#.00")
    'A LA ELECTRONIA DICE QUE LE FALTA PARA PODER COBRAR LA FACTURA, SE HACE UN PAGO ADICIONAL
    Rem cCadena = cCadena & "{""id"": """ & 1 & """type"": """ & "99" & """,""amt"":  """ & nDiferenciaConFE & """,""desc"": """ & "FE . AJUSTE" & """}" & coma
    cCadena = cCadena & "{""id"": " & i & ", ""type"": """ & "99" & """,""amt"":  """ & cLocalPayment & """,""desc"": """ & "FE . AJUSTE" & """}" & coma
    nPagosFuturo = local_total_factura
End If


If Right(cCadena, 1) = "," Then cCadena = Left(cCadena, Len(cCadena) - 1)

Generate_NEW_Payments_OFFLINE = "[" & cCadena & "]"

'Set rsFEPropinas = New ADODB.Recordset
'rsFEPropinas.Fields.Append "description", adChar, 35
'rsFEPropinas.Fields.Append "amount", adSingle
'
'rsFEPropinas.CursorType = adOpenStatic
'rsFEPropinas.LockType = adLockOptimistic
'rsFEPropinas.Open

End Function


'---------------------------------------------------------------------------------------
' Procedure : TestRequest_OFFLINE
' Author    : hsequeira
' Date      : 19/12/2024
' Purpose   : HACE EL REQUEST OFF LINE Y LE PASA EL CONTROL AL AGENTE
'---------------------------------------------------------------------------------------
'
Private Function TestRequest_OFFLINE(jsonContent As String, Optional TipoRequest As String) As Boolean
Dim response As String
Dim http As Object
'INFO: TOTAL DE LA FACTURA
'19NOV2024
Dim gTot As String
Dim xmltotal_factura As String
Dim cTemp As String
Dim cTemp2 As String
Dim nCurrentValor As Single
Dim nCurrenCounter As Long
Dim local_ambiente As String
Dim local_baseURL As String
Dim finalUrl  As String
Dim cLin As String

   On Error GoTo TestRequest_Error


    Call Clear_FE_Variables


    Dim url As String
    url = o_mi_urlWebservice
    If o_mi_Ambiente = 1 Then
        local_ambiente = "prod"
    Else
        local_ambiente = "test"
    End If
    local_baseURL = "{environment}/sendFileToProcess/{companyId}/{apiKey}"
    
    finalUrl = Replace(local_baseURL, "{environment}", local_ambiente)
    finalUrl = Replace(finalUrl, "{companyId}", o_X_FF_Company)
    finalUrl = Replace(finalUrl, "{apiKey}", o_X_FF_API_Key)
        
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    
    http.Open "POST", o_mi_urlWebservice & finalUrl, False
    
    http.setRequestHeader "Accept", "application/json"
    http.setRequestHeader "Content-Type", "application/json"

    cLin = "jsonContent"
    http.send jsonContent


    ' Extracting values from JSON
    Dim cufe As String
    Dim qrCodeData As String
    Dim pdfUrl As String
    Dim mensajes As String
    Dim mensajesType As String
    Dim mensajesCode As String
    Dim mensajesMsg As String
    
    ' Check if the request was successful (status code 200)
    If http.Status = 200 Then
        cFE_LastErrorCode = http.Status
        response = http.responseText
        'Form1.lstLog.AddItem "RESPONSE OK " & Time
    Else
        response = "Error: " & http.Status & " - " & http.statusText & " + " & http.responseText
        If Len(http.responseText) < 250 Then
            mensajesMsg = response
        Else
            mensajesType = ExtractValueFromJSON3(response, "type")
            mensajesCode = ExtractValueFromJSON3(response, "code")
            mensajesMsg = ExtractValueFromJSON3(response, "message")
            cFE_LastErrorCode = mensajesCode
        End If
        
        If TipoRequest = "NC" Then
            ShowMsg "ERROR EN NOTA DE CREDITO (TestRequest): " & vbCrLf & "Codigo: " & mensajesCode & vbCrLf & "Motivo Error: " & mensajesMsg, vbBlue, vbRed
            EscribeLog "ERROR EN NOTA DE CREDITO: " & mensajesCode & ". " & mensajesMsg
        Else
            ShowMsg "ERROR EN FACTURA (TestRequest): " & vbCrLf & "Codigo: " & mensajesCode & vbCrLf & "Motivo Error: " & mensajesMsg, vbBlue, vbRed
            EscribeLog "ERROR EN FACTURA: " & mensajesCode & ". " & mensajesMsg
        End If
        
        TestRequest_OFFLINE = False
        Exit Function
    End If


'INFO: TOTAL DE LA FACTURA
'19NOV2024
'SE DEBE LIMPIAR CUANDO SE TIRA LA Z
gTot = ExtractGTot(response, "gTot")
On Error Resume Next
cTemp = RegRead("HKCU\Software\SoloSoftware\SoloMix\FE_Sales")
cTemp2 = RegRead("HKCU\Software\SoloSoftware\SoloMix\FE_Counter")
On Error GoTo 0
If Left(gTot, 6) = "Error:" Then
    EscribeLog "Error extracting gTot en Factura (" & nNUMTrans & ")"
Else
    xmltotal_factura = ExtractGTot(gTot, "dVTot")
   
    If cTemp = "" Then
        nCurrentValor = 0
        nCurrenCounter = 0
    Else
        nCurrenCounter = CLng(cTemp2)
        cTemp = RegRead("HKCU\Software\SoloSoftware\SoloMix\FE_Sales")
        nCurrentValor = CSng(cTemp)
    End If
    nCurrenCounter = nCurrenCounter + 1
    nCurrentValor = nCurrentValor + CSng(xmltotal_factura)
    
    RegWrite "HKCU\Software\SoloSoftware\SoloMix\FE_Counter", nCurrenCounter
    RegWrite "HKCU\Software\SoloSoftware\SoloMix\FE_Sales", nCurrentValor
End If


cufe = ExtractValueFromJSON3(response, "cufe")
FE_CUFE = cufe
qrCodeData = ExtractValueFromJSON3(response, "qrContent")
FE_QRData = qrCodeData
pdfUrl = ExtractValueFromJSON3(response, "pdf_url")
FE_pdfURL = pdfUrl

o_msg = ExtractValueFromJSON3(response, "msg")

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
mensajes = ExtractValueFromJSON2(response, "msg")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

mensajesType = ExtractValueFromJSON3(mensajes, "type")
mensajesCode = ExtractValueFromJSON3(mensajes, "code")
mensajesMsg = ExtractValueFromJSON3(mensajes, "message")

Rem If mensajesType = "N" And mensajesCode = "0260" Then
If cufe <> "" Then
    '~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~
    'INFO: UPDATE 16JUL2024
    '~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~
    If Left(mi_App_Name, 7) = "SoloMix" Then
        EscribeLog "FE. OK. Factura: " & rs00!TRANS + 1 & " - CUFE: " & Right(FE_CUFE, 15) & ", PDF: " & Right(FE_QRData, 15)
    Else
        EscribeLog "FE. OK. Factura: " & nNUMTrans & " - CUFE: " & Right(FE_CUFE, 15) & ", PDF: " & Right(FE_QRData, 15)
    End If
    TestRequest_OFFLINE = True
Else
    If Left(mi_App_Name, 7) = "SoloMix" Then
        ShowMsg "Factura # " & Str(rs00!TRANS + 1) & ". FAIL. TestRequest : " & vbCrLf & "mensajesType : " & mensajesType & ", Code : " & mensajesCode & " .Msg: " & mensajesMsg, vbBlue, vbRed
        EscribeLog "FE. FAIL. mensajesType : " & mensajesType & ", Code : " & mensajesCode & " .Msg: " & mensajesMsg
    Else
        ShowMsg "Factura # " & Str(nNUMTrans) & ". FAIL. TestRequest : " & vbCrLf & "mensajesType : " & mensajesType & ", Code : " & mensajesCode & " .Msg: " & mensajesMsg, vbBlue, vbRed
        EscribeLog "FE. FAIL. mensajesType : " & mensajesType & ", Code : " & mensajesCode & " .Msg: " & mensajesMsg
    End If
    
'    If mensajesCode = "1002" Then       '1002= FACTURA DUPLICADA, INCREMENTAR EL CONTADOR e INTENTAR DE NUEVO
'
'        If Left(mi_App_Name, 7) = "SoloMix" Then
'            Call SOLOTrans("BEGIN")
'            msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1"
'            Call SOLOTrans("COMMIT")
'        Else
'            On Error Resume Next
'            'INFO: 11JUN2012)
'            #If FAST_POS = 1 Then
'            #Else
'            #End If
'            On Error GoTo 0
'        End If
'    End If
End If
    
   On Error GoTo 0
   Exit Function

TestRequest_Error:
    Rem If cLogMode = "SI" Then
        OLD_EscribeLog jsonContent
        OLD_EscribeLog response
    Rem End If

    ShowMsg "Error " & Err.Number & " (" & Err.description & ")" & ". Response ErrorCode:" & cFE_LastErrorCode & ", en TestRequest" & vbCrLf & "OFF-LINE", vbBlue, vbCyan
    EscribeLog "FE.TestRequest. OFF-LINE: " & "Error " & Err.Number & " (" & Err.description & ") "
    Resume Next
End Function

Private Function Clear_FE_Variables() As Boolean
Rem limpia las varables de FE
FE_CUFE = ""
FE_QRData = ""
FE_pdfURL = ""
FE_process_date = ""
FE_digest_value = ""
FE_authorization_number = ""
o_msg = ""
End Function

Public Function BusquedaCliente() As Long

ClienteBuscar.Show 1

End Function

Public Function GetTipoClienteDescrip(cTipoCliente As String) As String
'INFO: 20FEB2025
Select Case cTipoCliente
    Case "01"
        GetTipoClienteDescrip = "Contribuyente"
    Case "02"
        GetTipoClienteDescrip = "Consumidor Final"
    Case "03"
        GetTipoClienteDescrip = "Gobierno"
    Case "04"
        GetTipoClienteDescrip = "Extranjero"
    Case Else
        GetTipoClienteDescrip = "Consumidor Final"
End Select

End Function


'---------------------------------------------------------------------------------------
' Procedure : toJSON
' Author     : hsequeira
' Date        : 26/10/2021
' Date        :  3/07/2022 -> SE ESCLARECE PRESENTACION EN EL FILE JSON
' Purpose   : PASA UN RECORDSET A JSON
' 8DIC2021 : HEADER AND FOOTER
'---------------------------------------------------------------------------------------
'
Public Function toJSON(PassTblQry As String)
' EXPORT JSON FILE FROM TABLE OR QUERY
'Dim mydb As Database, rs As Recordset
'Dim VarField(255), VarFieldType(255)
'Dim fld As DAO.Field, cValorCampo As String
'Dim db As DAO.Database
'
'db.op
'Set Db = CurrentDb
   On Error GoTo toJSON_Error

'Set db = msConn

Dim cSQL As String
Dim rsTABLA As ADODB.Recordset
Dim TipoCampo As Integer
Dim BuscaCadenaError As Integer
Dim cConnString  As String
Dim JmsConn As ADODB.Connection

Set JmsConn = New ADODB.Connection


cConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cDataPath
cConnString = cConnString & ";Jet OLEDB:Database Password=master24"
Set JmsConn = New ADODB.Connection
JmsConn.Open cConnString


Set rsTABLA = New ADODB.Recordset

If Left(PassTblQry, 6) = "SELECT" Then
    cSQL = PassTblQry
    Rem fn = CurDir(dlgDialog.FileName) & "\QUERY " & Format(Now(), "YYYY-MM-DD HHMM") & ".json"  ' define export current folder query date/time
    fn = App.Path & "\QUERY " & Format(Now(), "YYYY-MM-DD HHMM") & ".json"  ' define export current folder query date/time
    
    'jsonTableName = PassTblQry   // info 21jun200
    jsonTableName = "sql_query"
Else
    cSQL = MakeSQL(PassTblQry)
    
    Rem fn = App.Path & "\" & PassTblQry & " " & Format(Now(), "YYYY-MM-DD HHMM") & ".json"   ' define export current folder query date/time
    'info: 08nov2025
    fn = App.Path & "\HOLD" & ".json"   ' define export current folder query date/time
End If

rsTABLA.Open cSQL, JmsConn, adOpenStatic, adLockReadOnly

'fn = CurrentProject.Path & "\" & PassTblQry & " " & Format(Now(), "YYYY-MM-DD HHMM") & ".json" ' define export current folder query date/time
Open fn For Output As #1    ' output to text file
'Recs = DCount("*", PassTblQry) ' record count
'Set rs = db.OpenRecordset("Select * from [" & PassTblQry & "]")
Nonulls = True ' set NoNulls = true to remove all null values within output ELSE set to false
fieldcount = 0
' Save field count, fieldnames, and type into array
For Each fld In rsTABLA.Fields
    fieldcount = fieldcount + 1
    '''VarField(fieldcount) = fld.Name
    'Debug.Print VarField(fieldcount)
    '''''''''''VarFieldType(fieldcount) = "TEXT"
    Select Case fld.Type
        Case 4, 5, 6, 7 ' fieldtype 4=long, 5=Currency, 6=Single, 7-Double
            '''''''''''VarFieldType(fieldcount) = "NUMBER"
    End Select
Next
Set fld = Nothing

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Print #1, "{ "                                                          ' HEADER JSON
' Print #1, Chr(34) & jsonTableName & Chr(34) & " : ["   ' Start; JSON; dataset
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Print #1, "[" 'Start; JSON; dataset

'Print #1, " : [" ' start JSON dataset
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'' Print #1, "[" ' start JSON dataset
' build JSON dataset from table/query data passed
Do While Not rsTABLA.EOF
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Print #1, Chr(9) & "{"  ' START JSON record
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' build JSON record from table/query record using fieldname and type arrays
    'For nCampoActual = 1 To fieldcount
    For nCampoActual = 0 To fieldcount - 1
        '''''''''''TipoCampo = VarFieldType(nCampoActual)
        'TipoCampo = VarFieldType(nCampoActual)
        TipoCampo = rsTABLA.Fields(nCampoActual).Type
        Select Case TipoCampo
            Case 2      'SMALL INT
                QuoteID = ""
            Case 3      'INTEGER
                QuoteID = ""
            Case 5      'DOUBLE
                QuoteID = ""
            Case 6      'CURRENCY
                QuoteID = Chr(34)
                'QuoteID = ""
            Case 11     'True o False // debe ser true o false en JSON
                QuoteID = ""
            Case 202    'STRING o VARCHAR
                QuoteID = Chr(34)
            Case Else
                QuoteID = Chr(34)
        End Select
        
        'If TipoCampo = 3 Or 11 Then QuoteID = "" Else QuoteID = Chr(34)     ' No quote for numbers
        'QuoteID = Chr(34) ' double quote for text
        '''''''''''If IsNull(rs(VarField(nCampoActual)).Value) Then  ' deal with null values
        '''''''''''    cValorCampo = "Null": QuoteID = ""   ' no quote for nulls
        '''''''''''    If Nonulls = True Then cValorCampo = "": QuoteID = Chr(34)                       ' null text to empty quotes
        '''''''''''    If Nonulls = True And TipoCampo = "NUMBER" Then cValorCampo = "0": QuoteID = ""      ' null number to zero without quotes
        '''''''''''    Else
        '''''''''''    cValorCampo = Trim(rs(VarField(nCampoActual)).Value)
        '''''''''''End If
        
        If IsNull(rsTABLA.Fields(nCampoActual).value) Then  ' deal with null values
            cValorCampo = "Null": QuoteID = ""   ' no quote for nulls
            If Nonulls = True Then cValorCampo = "": QuoteID = Chr(34)                       ' null text to empty quotes
            If Nonulls = True And TipoCampo = 3 Then cValorCampo = "0": QuoteID = ""      ' null number to zero without quotes
            Else
                Select Case TipoCampo
                    Case 11
                        Select Case rsTABLA.Fields(nCampoActual).value
                            Case True
                                cValorCampo = "true"
                            Case False
                                cValorCampo = "false"
                        End Select
                    Case Else
                        cValorCampo = Trim(rsTABLA.Fields(nCampoActual).value)
                End Select
        End If
        
        BuscaCadenaError = InStr(cValorCampo, "'")
        If BuscaCadenaError > 0 Then
            cValorCampo = Replace(cValorCampo, "'", "-")
        End If
        
        cValorCampo = Replace(cValorCampo, Chr(34), "'") ' replace double quote with single quote
        cValorCampo = Replace(cValorCampo, Chr(8), "")   ' remove backspace
        cValorCampo = Replace(cValorCampo, Chr(10), "")  ' remove line feed
        cValorCampo = Replace(cValorCampo, Chr(12), "")  ' remove form feed
        cValorCampo = Replace(cValorCampo, Chr(13), "")  ' remove carriage return
        cValorCampo = Replace(cValorCampo, Chr(9), "   ")  ' replace tab with spaces
        
        '''''''''''jsonRow = Chr(34) & VarField(nCampoActual) & Chr(34) & ":" & QuoteID & cValorCampo & QuoteID
        'Debug.Print TipoCampo & " - " & rsTabla.Fields(nCampoActual).Name & " - " & cValorCampo
        jsonRow = Chr(34) & LCase(rsTABLA.Fields(nCampoActual).name) & Chr(34) & ":" & QuoteID & cValorCampo & QuoteID
        If nCampoActual < fieldcount - 1 Then jsonRow = jsonRow & "," ' add comma if not last field
        
        Print #1, Chr(9) & Chr(9) & jsonRow
    
    Next nCampoActual
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Print #1, Chr(9) & "}";  ' END JSON record
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    rsTABLA.MoveNext
    If Not rsTABLA.EOF Then
        Print #1, "," ' add comma if not last record
        Else
        Print #1, ""
    End If
Loop

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Print #1, "],"  'FOOTER JSON FILE
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''Print #1, Chr(34) & "Info" & Chr(34) & " : {"
'''Print #1, Chr(34) & "Fecha" & Chr(34) & " : " & Chr(34) & Date & Chr(34) & ","
'''Print #1, Chr(34) & "Hora" & Chr(34) & " : " & Chr(34) & Time & Chr(34) & ","
'''Print #1, Chr(34) & "Autor" & Chr(34) & " : " & Chr(34) & App.EXEName & " - " & App.CompanyName & Chr(34)
'''Print #1, "}"
'''
'''Print #1, "}"  ' close JSON dataset
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Close #1

   On Error GoTo 0
   Exit Function

toJSON_Error:

    ShowMsg "Error " & Err.Number & vbCrLf & " (" & Err.description & ") in procedure toJSON "
    Resume Next

End Function


Private Function MakeSQL(cTabla As String) As String
Dim cSQL As String
Select Case cTabla
    Case "HIST_INVENT"
        cSQL = "SELECT A.ID, B.DESCRIP, A.BOD1_01, A.BOD2_01, A.COSTO_01, A.TOTAL_01, A.BOD1_02, A.BOD2_02, A.COSTO_02, A.TOTAL_02    ,A.BOD1_03 ,A.BOD2_03 ,A.COSTO_03    ,A.TOTAL_03    ,A.BOD1_04 ,A.BOD2_04 ,A.COSTO_04,A.TOTAL_04"
        cSQL = cSQL & ",A.BOD1_05, A.BOD2_05, A.COSTO_05, A.TOTAL_05, A.BOD1_06, A.BOD2_06, A.COSTO_06    ,A.TOTAL_06    ,A.BOD1_07 ,A.BOD2_07, A.COSTO_07, A.TOTAL_07, A.BOD1_08, A.BOD2_08, A.COSTO_08, A.TOTAL_08, A.BOD1_09 "
        cSQL = cSQL & ",A.BOD2_09 ,A.COSTO_09, A.TOTAL_09, A.BOD1_10, A.BOD2_10, A.COSTO_10, A.TOTAL_10, A.BOD1_11, A.BOD2_11, A.COSTO_11, A.TOTAL_11, A.BOD1_12, A.BOD2_12, A.COSTO_12, A.TOTAL_12"
        cSQL = cSQL & " FROM HIST_INVENT AS A, INVENT AS B "
        cSQL = cSQL & " WHERE A.ID = B.ID "
        MakeSQL = cSQL
    Case Else
        MakeSQL = "SELECT * FROM " & cTabla
End Select
End Function

Public Function GetLicenciaFE() As String
    Dim response As String
    Dim http As Object
    Dim url As String
    Dim cTipo As String
    Dim cAmbiente As String
    Dim baseURL As String
    Dim quota As String
    Dim quota_usage As String
    Dim end_date As String
    Dim folios_restantes As Double


    Select Case mi_Ambiente
        Case 1
            cAmbiente = "prod"
            baseURL = "https://api.facturafacil.com.pa/api/core/subscriptions/status/"
        Case 2
            cAmbiente = "test"
            baseURL = "https://backend-qa-api.facturafacil.com.pa/api/core/subscriptions/status/"
        Case Else
            GetLicenciaFE = ""
    End Select
    
    razonRetorno = ""
    'On Error GoTo TestRequest_Error

    Set http = CreateObject("MSXML2.ServerXMLHTTP")

    
    http.setOption 2, 13056
    http.Open "GET", baseURL, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "X-FF-Company", X_FF_Company
    http.setRequestHeader "X-FF-API-Key", X_FF_API_Key
    http.send

    ' Check if the request was successful (status code 200)
    If http.Status = 200 Then
        response = http.responseText
        ''Debug.Print http.responseText
        quota = ExtractValueFromJSON3(response, "quota")
        quota_usage = ExtractValueFromJSON3(response, "quota_usage ")
        end_date = ExtractValueFromJSON3(response, "end_date")
        folios_restantes = Val(quota) - Val(quota_usage)
        
        GetLicenciaFE = "Tiene Disponible " & vbCrLf & "(" & folios_restantes & ") de (" & quota & ") FOLIOS / FACTURAS " & vbCrLf & vbCrLf & "Hasta el " & end_date & " (año-mes-dia)"
    Else
        'Form1.lstLog.AddItem "FAIL http.send jsonContent" & Time
        response = "Error: " & http.Status & " - " & http.statusText & " + " & http.responseText
        'ShowMsg "FACTURA ERRONEA: " & vbCrLf & "Motivo Error: " & http.Status & " - " & http.statusText & " + " & http.responseText, vbBlue, vbRed
        'TestRequest = False
        ' Debug.Print http.Status & " - " & http.statusText & " + " & http.responseText
        'Debug.Print response
        
        GetLicenciaFE = "No se pudo obtener informacion de Quota de Folios/Facturas " & vbCrLf & "Contacte a SOLO POS MENU"
        Exit Function
    End If
End Function

Public Function MiDebug(cParam As String) As Boolean
Debug.Print "Time: " & Format(Time(), "Long Time") & " - " & cParam
End Function

