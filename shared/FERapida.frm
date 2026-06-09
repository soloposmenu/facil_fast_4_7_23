VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FERapida 
   BackColor       =   &H00EA6E67&
   Caption         =   "FACTURA RAPIDA"
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8160
   Icon            =   "FERapida.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   8160
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNombre 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      MaxLength       =   40
      TabIndex        =   2
      Top             =   3435
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   9
      Top             =   7560
      Width           =   1815
   End
   Begin VB.CommandButton cmdAplicaFR 
      Caption         =   "GENERAR FACTURA RAPIDA"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   7560
      Width           =   3015
   End
   Begin VB.TextBox txteMail 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      MaxLength       =   40
      TabIndex        =   3
      Top             =   4080
      Width           =   4815
   End
   Begin VB.ComboBox ComboFR 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   6480
      Width           =   4455
   End
   Begin VB.ComboBox ComboFRTax 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "FERapida.frx":030A
      Left            =   3240
      List            =   "FERapida.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   5955
      Width           =   1095
   End
   Begin VB.TextBox txtFValor 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      MaxLength       =   10
      TabIndex        =   5
      Top             =   5400
      Width           =   2175
   End
   Begin VB.TextBox txtFDescrip 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   4740
      Width           =   4815
   End
   Begin VB.TextBox txtFRZ 
      BackColor       =   &H000000FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3240
      MaxLength       =   40
      TabIndex        =   1
      Text            =   "Consumidor Final"
      Top             =   3150
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.TextBox txtFRUC 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      MaxLength       =   20
      TabIndex        =   0
      Text            =   "0-0-0"
      Top             =   2760
      Width           =   3615
   End
   Begin MSComctlLib.ProgressBar ProgBar 
      Height          =   525
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   926
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label Label 
      BackColor       =   &H00EA6E67&
      Caption         =   "NOMBRE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Index           =   1
      Left            =   1560
      TabIndex        =   20
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
      Caption         =   "Sin Impuesto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   28
      Left            =   1680
      TabIndex        =   18
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label 
      BackColor       =   &H00EA6E67&
      Caption         =   "Correo Electronico"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Label lbmsg 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1935
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   7935
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EA6E67&
      Caption         =   "Tipo Pago"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Index           =   18
      Left            =   1320
      TabIndex        =   15
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EA6E67&
      Caption         =   "Tasa Impuesto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   375
      Index           =   19
      Left            =   480
      TabIndex        =   14
      Top             =   6045
      Width           =   2655
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EA6E67&
      Caption         =   "Valor de Factura"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   375
      Index           =   20
      Left            =   360
      TabIndex        =   13
      Top             =   5445
      Width           =   2775
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EA6E67&
      Caption         =   "Motivo de Factura"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   375
      Index           =   21
      Left            =   240
      TabIndex        =   12
      Top             =   4920
      Width           =   2895
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EA6E67&
      Caption         =   "Razon Social"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   375
      Index           =   22
      Left            =   720
      TabIndex        =   11
      Top             =   3240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EA6E67&
      Caption         =   "R.U.C."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   375
      Index           =   23
      Left            =   1920
      TabIndex        =   10
      Top             =   2880
      Width           =   1095
   End
End
Attribute VB_Name = "FERapida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Private Const PBM_SETBARCOLOR As Long = &H409
Private Const PBM_SETBKCOLOR As Long = &H2001
Private Const PROGBAR_DEF_COLOR = &HFF000000 '&H8000000D


'---------------------------------------------------------------------------------------
' Procedure : cmdAplicaFR_Click
' Author    : hsequeira
' Date      : 21/06/2024
' Purpose   : COMANDO QUE EJECUTA LA FACTURA RAPIDA
' SON VALORES SIMPLES, CON LA INTENCION DE HACER UNA FACTURA EN EL CASO QUE NO
' SE DESEE FACTURAR EN EL SISTEMA.
' EN ESTA VERSION, SOLO ES PARA CLIENTE FINAL
'---------------------------------------------------------------------------------------
'
Private Sub cmdAplicaFR_Click()
Dim cGenerateHeader As String, cGenerateReceptor As String, cGenerateItems As String, cGeneratePayments As String
Dim cGenerateHeader_OFFLINE As String, cGenerateReceptor_OFFLINE As String, cGenerateItems_OFFLINE As String, cGeneratePayments_OFFLINE As String
Dim cTipoPago As String
Dim miJSONRequest As String
Dim nPrice As Single
Dim nTax As Single
Dim nPayAmount As Single
Dim cTaxType As String
Dim cFormadePago() As String
Dim cFPago As String

   On Error GoTo cmdAplicaFR_Click_Error

If Me.txtFDescrip = "" Then
    ShowMsg "DEBE ESCRIBIR UN MOTIVO", vbYellow, vbRed
    Me.txtFDescrip.SetFocus
    Exit Sub
End If

If Me.txtFValor = "" Then
    ShowMsg "DEBE ESCRIBIR UN VALOR DE LA FACTURA", vbYellow, vbRed
    Me.txtFValor.SetFocus
    Exit Sub
End If

If Me.txtEmail = "" Then Me.txtEmail = "eMail"


nPrice = Me.txtFValor
nTax = nPrice * (Val(Me.ComboFRTax.Text) / 100)
nPayAmount = Format((nPrice + nTax), "STANDARD")
cTipoPago = Left(ComboFR.Text, 2)

cFormadePago = Split(ComboFR.Text, "~")
cFPago = cFormadePago(1) & " / " & cFormadePago(1)

If FE_Facil_ONLINE Then
    Select Case Me.ComboFRTax.Text
        Case "0"
            cTaxType = "00"
        Case "7"
            cTaxType = "01"
        Case "10"
            cTaxType = "02"
        Case "15"
            cTaxType = "03"
    End Select
Else
    Select Case Me.ComboFRTax.Text
        Case "0"
            cTaxType = "0"
        Case "7"
            cTaxType = "1"
        Case "10"
            cTaxType = "2"
        Case "15"
            cTaxType = "3"
    End Select
End If

If FE_Facil_ONLINE Then
    cGenerateHeader = "{" & """header"": {""id"": " & 1 & ",""environment"": " & mi_Ambiente & "}"
    cGenerateReceptor = ",""document"": {""fd_number"": " & 0 & "," & """receptor"": {" & """type"": """ & "02" & """,""name"": """ & Me.txtNombre & """,""ruc_type"": " & 1 & "," & """address"": """ & "CONSUMIDOR FINAL" & """,""email"": """ & txtEmail.Text & """,""ruc"": """ & txtFRUC.Text & """,""dv"": """ & "" & """},"
    cGenerateItems = """items"": [{""line"": 1,""price"": " & nPrice & ",""mu"": ""und"",""quantity"": 1,""description"": """ & txtFDescrip.Text & """,""taxes"": [{""type"": ""01"", ""amount"": " & nTax & ", ""code"": """ & cTaxType & """}],""discount"": 0,""internal_code"": """"}],"
    cGeneratePayments = """payments"": [{""type"": """ & cTipoPago & """,""amount"": " & nPayAmount & "}],""total"": " & nPayAmount & ",""type"": ""01"",""info"": ""<string>""}}"
    
    miJSONRequest = cGenerateHeader
    miJSONRequest = miJSONRequest & cGenerateReceptor
    miJSONRequest = miJSONRequest & cGenerateItems
    miJSONRequest = miJSONRequest & cGeneratePayments
Else
    Rem cGenerateHeader_OFFLINE = "{" & """fiscalDoc"": {""id"": " & 1 & ",""environment"": " & mi_Ambiente & "}"
    Rem Call GetNewNumTrans
    Call GetNewNumTrans_LOCAL
    cGenerateReceptor_OFFLINE = "{" & """fiscalDoc"": {" & """companyLicCod"": """ & o_X_FF_Company & """,""branchCod"": """ & o_X_FF_Branch & _
                                                """,""posCod"": " & """001""" & "," & """docType"": """ & "F" & _
                                                """,""docNumber"": """ & (nNUMTrans) & """,""customerName"": """ & Me.txtNombre & _
                                                """,""Cajero"": """ & (npNumCaj) & """,""Mesero"": """ & nMesero & _
                                                """,""customerPhone"": """ & "CONSUMIDOR FINAL" & """,""customerRUC"": """ & txtFRUC.Text & _
                                                """,""customerType"": """ & "07" & """, ""customerAddress"": """ & "ESCRIBIR DIRECCION" & ""","
                                                
                                                Rem // """,""customerType"": """ & "07" & """, ""customerAddress"": """ & "ESCRIBIR DIRECCION" & """},"
                                                
    cGenerateItems_OFFLINE = """items"": [{""id"": 1,""price"": " & nPrice & ",""qty"": ""1"",""desc"": """ & txtFDescrip.Text & """,""tax"": """ & cTaxType & _
                                            """,""damt"": " & 0 & "}],"
    cGeneratePayments_OFFLINE = """payments"": [{""id"": """ & "1" & """,""type"": """ & cTipoPago & """,""amt"": """ & nPayAmount & """,""desc"": """ & cFPago & """}]}}"
    
    miJSONRequest = cGenerateHeader_OFFLINE
    miJSONRequest = miJSONRequest & cGenerateReceptor_OFFLINE
    miJSONRequest = miJSONRequest & cGenerateItems_OFFLINE
    miJSONRequest = miJSONRequest & cGeneratePayments_OFFLINE

End If

If FE_Facil_ONLINE Then
    ProgBar.value = 30
    If Public_TestRequest(miJSONRequest) Then
        ProgBar.value = 50
        Sleep 150
        ProgBar.value = 85
        Sleep 200
        ProgBar.value = 100
        ShowMsg "FACTURA RAPIDA REALIZADA CORRECTAMENTE", vbYellow, vbBlue
        Me.txtFValor = ""
        Me.txtFDescrip = ""
        Me.cmdAplicaFR.Enabled = False
        ProgBar.value = 0
    Else
        ShowMsg "FALLO AL HACER FACTURA." & vbCrLf & "INTENTE HACER LA FACTURA DE NUEVO", vbBlue, vbCyan
        ProgBar.value = 0
    End If
Else
    If Public_TestRequest(miJSONRequest, "", "OFF-LINE") Then
        ProgBar.value = 50
        Sleep 150
        ProgBar.value = 85
        Sleep 200
        ProgBar.value = 100
        ShowMsg "FACTURA RAPIDA REALIZADA CORRECTAMENTE", vbYellow, vbBlue
        Me.txtFValor = ""
        Me.txtFDescrip = ""
        Me.cmdAplicaFR.Enabled = False
        ProgBar.value = 0
        
        Rem Call GetNewNumTrans
        Call GetNewNumTrans_LOCAL
    Else
        'FALLO EN FACTURA ELECTRONICA
        ShowMsg "FALLO AL HACER FACTURA." & vbCrLf & "INTENTE HACER LA FACTURA DE NUEVO", vbBlue, vbCyan
        ProgBar.value = 0
    End If
End If

   On Error GoTo 0
   Exit Sub

cmdAplicaFR_Click_Error:

    ShowMsg "Ocurrio un Error el los datos" & vbCrLf & "Revise:" & vbCrLf & Err.Number & " (" & Err.description & ")", vbYellow, vbRed

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim cmsg As String

txtFecIni = Format(Date, "SHORT DATE")
Call GetFacilRegistryValues

If mi_Ambiente = 2 Then Me.BackColor = vbCyan

cmsg = "Factura Electrónica Rápida UNICAMENTE entra a la Plataforma Web de la Facturación Electrónica." & vbCrLf & vbCrLf
cmsg = cmsg & "<<<< NO >>>> " & vbCrLf & vbCrLf
cmsg = cmsg & "va a Entra en la Facturación del Sistema de Ventas en su Local"
lbmsg.Caption = cmsg

ComboFRTax.AddItem "0", 0
ComboFRTax.AddItem "7", 1
ComboFRTax.AddItem "10", 2
ComboFRTax.AddItem "15", 3
ComboFRTax.ListIndex = 1

Call LoadComboPagos

Call ChangeColorOfProgBar
End Sub


 '{"header": {"id": 1,"environment": 2},"document": {"fd_number": 0,"receptor": {"type": "02","name": "CONSUMIDOR FINAL","ruc_type": 1,"address": "ESCRIBIR DIRECCION","email": "eMail","ruc": "0-0-0","dv": ""},"items": [{"line": 1,"price": 4.5,"mu": "und","quantity": 1,"description": "POSTRES TARASCOS","taxes": [{"type": "01", "amount": 0.315, "code": "01"}],"discount": 0,"internal_code": 440}],"payments": [{"type": "02","amount": 4.82}],"total":"4.82","type": "01","info": "<string>"}}


Private Sub LoadComboPagos()
Dim rsPPagos As ADODB.Recordset
Dim cSQL As String
Dim nIndice As Long, nSelected As Long

cSQL = "SELECT CODIGO, ID_FE, DESCRIP FROM PAGOS WHERE TIPO IN ('EF','CH','TJ','OT') ORDER BY DESCRIP"
Set rsPPagos = New ADODB.Recordset

rsPPagos.Open cSQL, msConn, adOpenStatic, adLockOptimistic
Do While Not rsPPagos.EOF
'    ComboPagos.AddItem rsPPagos!DESCRIP
'    ComboND.AddItem rsPPagos!DESCRIP
    ComboFR.AddItem rsPPagos!ID_FE & "~" & rsPPagos!DESCRIP
    'INFO: 7ENE2015
'    ComboNCRPago.AddItem rsPPagos!DESCRIP
    If rsPPagos!DESCRIP = "EFECTIVO" Then
        nSelected = nIndice
    End If
    nIndice = nIndice + 1
    rsPPagos.MoveNext
Loop
rsPPagos.Close
Set rsPPagos = Nothing
'ComboPagos.ListIndex = nSelected
'ComboND.ListIndex = nSelected
ComboFR.ListIndex = nSelected
'INFO: 7ENE2015
'ComboNCRPago.ListIndex = nSelected
End Sub
Private Sub txtFValor_GotFocus()
    txtFValor.SelStart = 0
    txtFValor.SelLength = Len(txtFValor.Text)
    cmdAplicaFR.Enabled = True
End Sub


Private Sub ChangeColorOfProgBar()
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    ''Call SendMessageLong(ProgBar.hwnd, PBM_SETBARCOLOR, 0&, ByVal 255)      'Red
    ''Call SendMessageLong(ProgBar.hwnd, PBM_SETBKCOLOR, 0&, ByVal 65535)    'yellow
If FE_Facil_ONLINE Then
    Call SendMessageLong(ProgBar.hwnd, PBM_SETBARCOLOR, 0&, ByVal 65280)       'green
    Call SendMessageLong(ProgBar.hwnd, PBM_SETBKCOLOR, 0&, ByVal 0)    'black
Else
    Call SendMessageLong(ProgBar.hwnd, PBM_SETBARCOLOR, 0&, ByVal 255)      'Red
    Call SendMessageLong(ProgBar.hwnd, PBM_SETBKCOLOR, 0&, ByVal 65535)    'yellow
End If
Rem OLD_EscribeLog "ChangeColorOfProgBar:" & fe_online
'Call VerificaURL2
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
End Sub

Private Sub GetNewNumTrans_LOCAL()

If Left(mi_App_Name, 7) = "SoloMix" Then
    msConn.BeginTrans
    msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1"
    msConn.CommitTrans
Else
    'INCREMENTA EL NUMERO DE TRANSACCION EN 1
    msConn.BeginTrans
    msConn.Execute "UPDATE ORGANIZACION SET TRANS_FAST = TRANS_FAST + 1"
    msConn.CommitTrans
    
    'LEE LA INFORMACION DE TRANSACCION
    rs00.Requery
    nNUMTrans = rs00!TRANS_FAST
End If

'=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/
'=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/=/
End Sub
