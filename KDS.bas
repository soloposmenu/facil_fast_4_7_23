Attribute VB_Name = "KDS"
Option Explicit
Public KDS_Activo As String
Public KDS_Path As String

Public Function KDS_File(Optional cTicketNumber As String) As Boolean
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Procedure : KDS_File
' Author    : hsequeira
' Date      : 16MAY2014
' Purpose   : ARMA SQL PARA LOS KDS
' Parametros: # DE TICKET
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim cSQL As String
Dim KDS_Orden As String
Dim cDestination As String
Dim KDSStation As String

If PLU.ChkToGo.value = vbChecked Then cDestination = "PARA LLEVAR" Else cDestination = "AQUI"

'KDS_CONSOLIDADO
KDSStation = "2"
cSQL = "SELECT A.CAJA, A.CAJERO, B.NOMBRE & ' ' & B.APELLIDO AS CAJERO_NOMBRE, "
cSQL = cSQL & " A.MESERO, '' AS MESERO_NOMBRE, "
cSQL = cSQL & " A.MESA, A.CUENTA, A.VALID, A.LIN, A.DESCRIP, A.CANT, "
cSQL = cSQL & " A.PRECIO_UNIT, A.PRECIO, A.CON_TAX, A.DESCUENTO, A.TIPO "
cSQL = cSQL & " FROM TMP_TRANS AS A, CAJEROS AS B "
cSQL = cSQL & " WHERE A.CAJERO = B.NUMERO "
cSQL = cSQL & " AND A.MESA = " & nMesa
cSQL = cSQL & " ORDER BY A.LIN"

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
KDS_Orden = KDS_Path & "KDS_CONSOLIDADO_" & cTicketNumber & ".xml"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Call KDS_Create(cSQL, KDS_Orden, cDestination, KDSStation, cTicketNumber)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'KDS_COCINA
KDSStation = "0"
cSQL = "SELECT A.CAJA, A.CAJERO, B.NOMBRE & ' ' & B.APELLIDO AS CAJERO_NOMBRE, "
cSQL = cSQL & " A.MESERO, '' AS MESERO_NOMBRE, "
cSQL = cSQL & " A.MESA, A.CUENTA, A.VALID, A.LIN, A.DESCRIP, A.CANT, "
cSQL = cSQL & " A.PRECIO_UNIT, A.PRECIO, A.CON_TAX, A.DESCUENTO, A.TIPO "
cSQL = cSQL & " FROM TMP_TRANS AS A, CAJEROS AS B "
cSQL = cSQL & " WHERE A.CAJERO = B.NUMERO "
cSQL = cSQL & " AND A.MESA = " & nMesa
cSQL = cSQL & " AND (A.IMPRESORA = 0 OR A.IMPRESORA = 2) "
cSQL = cSQL & " ORDER BY A.LIN"

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
KDS_Orden = KDS_Path & "KDS_COCINA_" & cTicketNumber & ".xml"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Call KDS_Create(cSQL, KDS_Orden, cDestination, KDSStation, cTicketNumber)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'KDS_BAR
KDSStation = "1"
cSQL = "SELECT A.CAJA, A.CAJERO, B.NOMBRE & ' ' & B.APELLIDO AS CAJERO_NOMBRE, "
cSQL = cSQL & " A.MESERO, '' AS MESERO_NOMBRE, "
cSQL = cSQL & " A.MESA, A.CUENTA, A.VALID, A.LIN, A.DESCRIP, A.CANT, "
cSQL = cSQL & " A.PRECIO_UNIT, A.PRECIO, A.CON_TAX, A.DESCUENTO, A.TIPO "
cSQL = cSQL & " FROM TMP_TRANS AS A, CAJEROS AS B "
cSQL = cSQL & " WHERE A.CAJERO = B.NUMERO "
cSQL = cSQL & " AND A.MESA = " & nMesa
cSQL = cSQL & " AND (A.IMPRESORA = 1 OR A.IMPRESORA = 2) "
cSQL = cSQL & " ORDER BY A.LIN"

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
KDS_Orden = KDS_Path & "KDS_BAR_" & cTicketNumber & ".xml"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Call KDS_Create(cSQL, KDS_Orden, cDestination, KDSStation, cTicketNumber)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


PLU.ChkToGo.value = vbUnchecked
KDS_File = True

End Function


'---------------------------------------------------------------------------------------
' Procedimiento : KDS_Create
' Autor       : hsequeira
' Fecha       : 28/05/2014
' Proposito   : CREA LOS DIFERENTES ARCHIVOS XML NECESARIOS
'---------------------------------------------------------------------------------------
'
Private Function KDS_Create(cSQL As String, KDS_Orden As String, cDestination As String, cKDSStation As String, cTicketNumber As String)
Dim rsProductos As ADODB.Recordset
Dim nFreefile As Byte
Dim cDoNotPrint As String

   On Error GoTo KDS_Create_Error

Set rsProductos = New ADODB.Recordset

cDoNotPrint = "BECEC-DC-BVOVO-"

rsProductos.Open cSQL, msConn, adOpenStatic, adLockReadOnly

If Not rsProductos.EOF Then

    nFreefile = FreeFile()
    Open KDS_Orden For Output As #nFreefile

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Print #nFreeFile, "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & "?>"
    Print #nFreefile, "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "Windows-1252" & Chr(34) & "?>"
    Print #nFreefile, "<Transaction>"
    Print #nFreefile, "<Order>"
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Print #nFreefile, "<ID>" & cTicketNumber & "</ID>"
    Print #nFreefile, "<PosTerminal>1</PosTerminal>"
    Print #nFreefile, "<TransType>1</TransType>"
    Print #nFreefile, "<OrderStatus>1</OrderStatus>"
    Print #nFreefile, "<OrderType></OrderType>"
    Print #nFreefile, "<ServerName></ServerName>"
    Print #nFreefile, "<Destination>" & cDestination & "</Destination>"
    Print #nFreefile, "<GuestTable></GuestTable>"
    Print #nFreefile, "<UserInfo></UserInfo>"
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    '~~~~~~~~~~~~~~~~~~
    Do While Not rsProductos.EOF
    '~~~~~~~~~~~~~~~~~~
Proximo_Registro:
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        'VALIDACION DE ITEMS QUE NO HAY QUE IMPRIMIR
'        If InStr(cDoNotPrint, Left(rsProductos!TIPO, 3)) > 0 Then
'            'MsgBox rsProductos!DESCRIP & " - " & rsProductos!TIPO
'            Print #nFreefile, "<Item>"
'            Print #nFreefile, "</Item>"
'            rsProductos.MoveNext
'            If rsProductos.EOF Then
'                Exit Do
'            Else
'                GoTo Proximo_Registro:
'            End If
'        End If
        
        Do While InStr(cDoNotPrint, Left(rsProductos!Tipo, 3)) > 0
            Print #nFreefile, "<Item>"
            Print #nFreefile, "</Item>"
            rsProductos.MoveNext
            If rsProductos.EOF Then
                GoTo Proximo_Registro_II:
            Else
                GoTo Proximo_Registro:
            End If
        Loop
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        
        Print #nFreefile, "<Item>"
        Print #nFreefile, "<ID>" & rsProductos!LIN & "</ID>"
        Print #nFreefile, "<TransType>1</TransType>"
        Print #nFreefile, "<Name>" & rsProductos!DESCRIP & "</Name>"
        Print #nFreefile, "<Category></Category>"
        Print #nFreefile, "<Quantity>" & rsProductos!CANT & "</Quantity>"
        Print #nFreefile, "<KDSStation>" & cKDSStation & "</KDSStation>"
        rsProductos.MoveNext
        If rsProductos.EOF Then
            Print #nFreefile, "</Item>"
            Exit Do
        Else
            '@ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~ @
            Do While Left(rsProductos!DESCRIP, 1) = "@"
            '@ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~ @
                If InStr(cDoNotPrint, Left(rsProductos!Tipo, 3)) > 0 Then
                    Print #nFreefile, "<Condiment>"
                    Print #nFreefile, "</Condiment>"
                Else
                    Print #nFreefile, "<Condiment>"
                    Print #nFreefile, "<ID>" & rsProductos!LIN & "</ID>"
                    Print #nFreefile, "<TransType>1</TransType>"
                    Print #nFreefile, "<Name>" & Mid(rsProductos!DESCRIP, 2, 40) & "</Name>"
                    Print #nFreefile, vbTab & "<Color BG=" & Chr(34) & "8" & Chr(34) & " FG=" & Chr(34) & "213" & Chr(34) & "></Color>"
                    Print #nFreefile, "<Action>-1</Action>"
                    Print #nFreefile, "</Condiment>"
                End If
                rsProductos.MoveNext
                If rsProductos.EOF Then Exit Do
            '@ ~~ @
            Loop
            '@ ~~ @
        End If
        Print #nFreefile, "</Item>"
    '~~~~
Proximo_Registro_II:
    Loop
    '~~~~
    
    Print #nFreefile, "</Order>"
    Print #nFreefile, "</Transaction>"

End If

Close #nFreefile

rsProductos.Close
Set rsProductos = Nothing
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

   On Error GoTo 0
   Exit Function

KDS_Create_Error:

    ShowMsg "No funciona la conexión al KDS." & vbCrLf & "Folder/Directorio del KDS: " & vbCrLf & KDS_Path & vbCrLf & "Error: " & Err.Number & " (" & Err.description & ")", vbYellow, vbRed
End Function
