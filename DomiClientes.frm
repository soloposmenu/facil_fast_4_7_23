VERSION 5.00
Begin VB.Form DomiClientes 
   BackColor       =   &H00B39665&
   Caption         =   "CLIENTES DOMICILIO"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   9915
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdListaDomicilio 
      Caption         =   "&LISTA DE CLIENTES"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6000
      Picture         =   "DomiClientes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txtDir 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   1800
      MaxLength       =   150
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   3060
      Width           =   3735
   End
   Begin VB.TextBox txteMail 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   7
      Top             =   5160
      Width           =   3735
   End
   Begin VB.TextBox txtEmpresa 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   4
      Top             =   2520
      Width           =   3735
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000FFFF&
      Caption         =   "ASIGNAR CLIENTE EN PANTALLA"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8040
      Picture         =   "DomiClientes.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txtFax 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      MaxLength       =   10
      TabIndex        =   8
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&CANCELAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      Picture         =   "DomiClientes.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "GUARDAR DATOS"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      Picture         =   "DomiClientes.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox txtNotas 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5880
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   9
      ToolTipText     =   "LAS NOTAS SE REFLEJAN EN EL PEDIDO ACTUAL"
      Top             =   4200
      Width           =   3735
   End
   Begin VB.CommandButton cmdSeek 
      Caption         =   "&BUSCAR CLIENTE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3960
      Picture         =   "DomiClientes.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   360
      Width           =   1815
   End
   Begin VB.ComboBox cmbZona 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2520
      Width           =   3735
   End
   Begin VB.TextBox txtExt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      MaxLength       =   5
      TabIndex        =   1
      Top             =   720
      Width           =   900
   End
   Begin VB.TextBox txtApellido 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      MaxLength       =   25
      TabIndex        =   3
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox txtNombre 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      MaxLength       =   25
      TabIndex        =   2
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox txtNumber 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      MaxLength       =   10
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label LbMesa 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   7920
      TabIndex        =   24
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "FAX"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   11
      Left            =   5880
      TabIndex        =   22
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "E-MAIL / CORREO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Index           =   10
      Left            =   240
      TabIndex        =   21
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "NOTAS ENTREGA / COMENTARIOS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   9
      Left            =   5880
      TabIndex        =   20
      Top             =   3960
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "TELEFONO*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   19
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "DIRECCION"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   18
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "ZONA*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   4
      Left            =   5880
      TabIndex        =   17
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "EMPRESA*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   16
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "EXTENSION"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   15
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "APELLIDO*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   14
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "NOMBRE*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "DomiClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsDomiClientes As ADODB.Recordset
Private bCustomerAlreadySelected As Boolean
Public Function ShowDomi(cTel As String) As Boolean
Me.txtNumber = cTel
Me.txtNumber.BackColor = vbGreen
bCustomerAlreadySelected = True
If Me.Visible Then
Else
    Me.Show 1
End If
End Function
Private Sub cmdCancelar_Click()

'ShowMsg "AL CANCELAR YA NO PODRA ASIGNAR OTRO PEDIDO A ESTA MESA"

On Error Resume Next
rsDomiClientes.Close
Set rsDomiClientes = Nothing
On Error GoTo 0

Set DomiClientes = Nothing
Unload Me
End Sub
Private Function DataEntryCleanUp() As Boolean
With Me
    '.txtNumber = ""
    '.txtExt = ""
    '.txtNotas2 = rsDomiClientes!NOTA_ENTREGA2
    .txtNombre = ""
    .txtApellido = ""
    .txtEmpresa = "PERSONAL"
    .txtDir = ""
'    .txtDir1 = ""
'    .txtDir2 = ""
'    .txtDir3 = ""
    .txtEmail = ""
    .txtNotas = ""
    .txtFax = ""
    .cmbZona.ListIndex = 0
    cmdGuardar.Caption = "GUARDAR DATOS"
    cmdReturn.Enabled = False
End With

End Function
Private Sub cmdGuardar_Click()
Dim cSQL As String
Dim nAffected As Long

On Error GoTo ErrAdm:
If MsgBox("¿ DESEA GUARDAR LOS DATOS EN PANTALLA ?", vbQuestion + vbYesNoCancel, "BUSQUEDA DE CLIENTE") = vbYes Then

    '>>> INFO. SE PONE MAS ABAJO YA QUE HAY UNA TRANSACCION EN GETCLIENTES. (msConnDomi.BeginTrans)
   If cmdGuardar.Caption = "MODIFICAR CLIENTE" Then
        With Me
            cSQL = "UPDATE CLIENTES SET "
            'cSQL = cSQL & "NUMERO = " & .txtNumber & ","
            cSQL = cSQL & "NOMBRE = '" & .txtNombre & "',"
            cSQL = cSQL & "APELLIDO = '" & .txtApellido & "',"
            cSQL = cSQL & "EMPRESA = '" & .txtEmpresa & "',"
            cSQL = cSQL & "EXTENSION = '" & .txtExt & "',"
            
             cSQL = cSQL & "DIRECCION1 = '" & Left(.txtDir, 50) & "',"
             cSQL = cSQL & "DIRECCION2 = '" & Mid(.txtDir, 51, 50) & "',"
             cSQL = cSQL & "DIRECCION3 = '" & Mid(.txtDir, 101, 50) & "',"

'            cSQL = cSQL & "DIRECCION1 = '" & RemoveNULL(.txtDir1) & "',"
'            cSQL = cSQL & "DIRECCION2 = '" & RemoveNULL(.txtDir2) & "',"
'            cSQL = cSQL & "DIRECCION3 = '" & RemoveNULL(.txtDir3) & "',"

            cSQL = cSQL & "EMAIL = '" & .txtEmail & "',"
            cSQL = cSQL & "NOTA_ENTREGA1 = '" & RemoveNULL(.txtNotas) & "',"
            '.txtNotas2 = rsDomiClientes!NOTA_ENTREGA2
            cSQL = cSQL & "FAX = '" & .txtFax & "',"
            cSQL = cSQL & "ZONA = '" & Left(cmbZona.Text, 3) & "'"
            cSQL = cSQL & " WHERE TELEFONO = '" & .txtNumber & "'"
            
            'INFO: 15ENE2011 - REVISA SI CAMBIO LA EXTENSION
            If .txtExt.Tag = .txtExt.Text Then
                cSQL = cSQL & " AND EXTENSION = '" & .txtExt & "'"
            Else
                'SI LA EXTENSION CAMBIA NADA MAS ACTUALIZA EL NUMERO DE TELEFONO
            End If
        End With
    Else

        cSQL = "INSERT INTO CLIENTES (ID, TELEFONO, EXTENSION,"
        cSQL = cSQL & "NOMBRE, APELLIDO, EMPRESA, ZONA, DIRECCION1,"
        cSQL = cSQL & "DIRECCION2, DIRECCION3, EMAIL, FAX, NOTA_ENTREGA1 ) "
        cSQL = cSQL & " VALUES ("
        With Me
            cSQL = cSQL & GetCounter("CLIENTES") & ",'"
            cSQL = cSQL & .txtNumber & "','" & .txtExt & "','"
            cSQL = cSQL & .txtNombre & "','" & .txtApellido & "','"
            cSQL = cSQL & .txtEmpresa & "','"
            cSQL = cSQL & Left(.cmbZona.Text, 3) & "','"
            
            cSQL = cSQL & Left(.txtDir, 50) & "','"
            cSQL = cSQL & Mid(.txtDir, 51, 50) & "','"
            cSQL = cSQL & Mid(.txtDir, 101, 50) & "','"
            
'            cSQL = cSQL & RemoveNULL(.txtDir1) & "','"
'            cSQL = cSQL & RemoveNULL(.txtDir2) & "','"
'            cSQL = cSQL & RemoveNULL(.txtDir3) & "','"
            cSQL = cSQL & .txtEmail & "','" & .txtFax & "','"
            cSQL = cSQL & .txtNotas & "')"
        End With
    End If
    
    msConnDomi.BeginTrans
    msConnDomi.Execute cSQL, nAffected
    msConnDomi.CommitTrans
    
    cmdReturn.Enabled = True
    
    Call cmdReturn_Click

Else
    Me.txtNombre.SetFocus
End If
On Error GoTo 0
Exit Sub

ErrAdm:
    ShowMsg "ERROR : " & Err.Number & vbCrLf & Err.description
    EscribeLog "Ventas. ERROR. Domicilio Guardar Cliente. " & cSQL
    msConnDomi.RollbackTrans
End Sub
Private Function GetCounter(cTipoCounter As String) As Long
Dim rsCounter As New ADODB.Recordset

rsCounter.Open "SELECT ID FROM ID_COUNTER", msConnDomi, adOpenStatic

msConnDomi.BeginTrans
msConnDomi.Execute "UPDATE ID_COUNTER SET ID = ID + 1"
msConnDomi.CommitTrans

GetCounter = rsCounter!id
rsCounter.Close
Set rsCounter = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : cmdListaDomicilio_Click
' Author    : hsequeira
' Date      : 24/09/2012
' Purpose   : MUESTRA LA LISTA DE
'---------------------------------------------------------------------------------------
'
Private Sub cmdListaDomicilio_Click()

RegWrite "HKCU\Software\SoloSoftware\TelefonoDomicilio", txtNumber
ListaClientes.Show 1
End Sub

Private Sub cmdReturn_Click()
Dim cSQL As String

   On Error GoTo cmdReturn_Click_Error

Dim cSQL2 As String

If Not Validacion(txtNumber.Text) Then
    ShowMsg "EL NUMERO DE TELEFONO ES INCORRECTO o NO TIENE EL LARGO REQUERIDO." & vbCrLf & "NO PUEDEN EXISTIR CARACTERES ESPECIALES EN EL NUMERO DEL CLIENTE", vbYellow, vbRed
    'rsDomiClientes.Close
    txtNumber.SetFocus
    txtNumber.SelStart = 0
    txtNumber.SelLength = Len(txtNumber.Text)
    Exit Sub
End If

If bCustomerAlreadySelected Then
    cSQL = "UPDATE MESA_CLIENTE SET TELEFONO = '" & Me.txtNumber & "', "
    cSQL = cSQL & " EXTENSION = '" & Me.txtExt & "',"
    cSQL = cSQL & " NOTA_ENTREGA1 = '" & Me.txtNotas & "',"
    cSQL = cSQL & " FECHA = '" & Format(Date, "YYYYMMDD") & "',"
    cSQL = cSQL & " HORA = '" & Format(Time, "HH:MM") & "'"
    cSQL = cSQL & " WHERE MESA = " & nMesa
Else
    
    cSQL2 = "DELETE * FROM MESA_CLIENTE WHERE MESA = 1"
    
    cSQL = "INSERT INTO MESA_CLIENTE (MESA, TELEFONO, NOTA_ENTREGA1, EXTENSION, FECHA, HORA) "
    cSQL = cSQL & " VALUES (" & nMesa & ",'" & Me.txtNumber & "','"
    cSQL = cSQL & txtNotas.Text & "','" & txtExt.Text & "','"
    cSQL = cSQL & Format(Date, "YYYYMMDD") & "','" & Format(Time, "HH:MM") & "')"
End If

msConnDomi.BeginTrans
If cSQL2 = "" Then Else msConnDomi.Execute cSQL2
msConnDomi.Execute cSQL
msConnDomi.CommitTrans

'PLU.lbMensaje.WordWrap = True
'PLU.lbMensaje = Me.txtNombre & Me.txtApellido
Call cmdCancelar_Click

   On Error GoTo 0
   Exit Sub

cmdReturn_Click_Error:
    If Err.Number = -2147467259 Then
    Else
        ShowMsg "Error " & Err.Number & " (" & Err.description & ") en cmdReturn_Click de Formulario DomiClientes", vbYellow, vbRed
        EscribeLog "Ventas. ERROR. Domicilio Guardar Cliente. " & cSQL
        msConnDomi.RollbackTrans
    End If
    Call cmdCancelar_Click
    On Error GoTo 0
End Sub

Private Sub cmdSeek_Click()
Dim cSQL As String

If Not Validacion(txtNumber.Text) Then
    ShowMsg "EL NUMERO DE TELEFONO ES INCORRECTO o NO TIENE EL LARGO REQUERIDO." & vbCrLf & "NO PUEDEN EXISTIR CARACTERES ESPECIALES EN EL NUMERO DEL CLIENTE", vbYellow, vbRed
    'rsDomiClientes.Close
    txtNumber.SetFocus
    txtNumber.SelStart = 0
    txtNumber.SelLength = Len(txtNumber.Text)
    Exit Sub
End If

cSQL = "SELECT * FROM CLIENTES WHERE TELEFONO = '" & txtNumber.Text & "' AND EXTENSION = '" & txtExt.Text & "'"

If rsDomiClientes.State = adStateOpen Then rsDomiClientes.Close

rsDomiClientes.Open cSQL, msConnDomi, adOpenStatic, adLockOptimistic

If rsDomiClientes.EOF Then
    Select Case MsgBox("CLIENTE CON EL NUMERO: " & txtNumber.Text & " - " & txtExt.Text & vbCrLf & "NO EXISTE, DESEA CREARLO ?", vbQuestion + vbYesNoCancel, "BUSQUEDA DE CLIENTE")
        Case vbNo, vbCancel
            DoEvents
            rsDomiClientes.Close
            txtNumber.SetFocus
            txtNumber.SelStart = 0
            txtNumber.SelLength = Len(txtNumber.Text)
            Exit Sub
        Case vbYes
            cmdGuardar.Enabled = True
            Call DataEntryCleanUp
            txtNombre.SetFocus
            cmdGuardar.Caption = "CREAR CLIENTE"
            Exit Sub
    End Select
Else
    If rsDomiClientes.RecordCount > 1 Then
    End If
    With Me
        .txtNombre = rsDomiClientes!nombre
        .txtApellido = rsDomiClientes!APELLIDO
        .txtEmpresa = rsDomiClientes!empresa
        .txtExt = rsDomiClientes!EXTENSION
        'PARA VALIDAR SI CAMBIO LA EXTENSION
        .txtExt.Tag = rsDomiClientes!EXTENSION
        
        .txtDir = rsDomiClientes!DIRECCION1 & rsDomiClientes!DIRECCION2 & rsDomiClientes!DIRECCION3
        
''        If IsNull(rsDomiClientes!DIRECCION1) Then .txtDir1 = "" Else .txtDir1 = RemoveNULL(rsDomiClientes!DIRECCION1)
''        If IsNull(rsDomiClientes!DIRECCION2) Then .txtDir2 = "" Else .txtDir2 = RemoveNULL(rsDomiClientes!DIRECCION2)
''        If IsNull(rsDomiClientes!DIRECCION3) Then .txtDir3 = "" Else .txtDir3 = RemoveNULL(rsDomiClientes!DIRECCION3)
        
        On Error Resume Next
        .txtEmail = rsDomiClientes!email
        .txtNotas = rsDomiClientes!NOTA_ENTREGA1
        '.txtNotas2 = rsDomiClientes!NOTA_ENTREGA2
        .txtFax = rsDomiClientes!FAX
        Call FindInCombo(.cmbZona, Left(rsDomiClientes!ZONA, 3), True, False, False)
        On Error GoTo 0
        cmdGuardar.Enabled = True
        cmdGuardar.Caption = "MODIFICAR CLIENTE"
        cmdReturn.Enabled = True
    End With
End If
End Sub
Private Function Validacion(cNumCliente As String) As Boolean
Dim nLargo As String
Dim i As Byte

Validacion = True
nLargo = Len(cNumCliente)

If nLargo < 7 Then
    Validacion = False
    Exit Function
End If
For i = 1 To nLargo
    Select Case Mid(cNumCliente, i, 1)
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "-"
        Case Else
            Validacion = False
            Exit For
    End Select
Next

End Function
Private Sub Form_Load()
Dim rsZona As ADODB.Recordset

'''If App.EXEName = "SysMesas" Then
'''    Call MoveScreen(Me)
'''End If

DoEvents
Set rsDomiClientes = New ADODB.Recordset
Set rsZona = New ADODB.Recordset

rsZona.Open "SELECT ZONA, DESCRIPCION_CORTA FROM ZONAS", msConnDomi, adOpenStatic, adLockReadOnly
rsZona.MoveFirst
Do While Not rsZona.EOF
    cmbZona.AddItem rsZona!ZONA & "-" & rsZona!DESCRIPCION_CORTA
    rsZona.MoveNext
Loop
rsZona.Close
Set rsZona = Nothing

LbMesa.Caption = nMesa

End Sub

' Module:       modFindInCombo.bas
' Function:     FindInCombo
' By:           John Allan Lee
' Last Updated: 05/09/2000
' Access:       Public
' Type:         Function
' Description:  Selects a specified item in a ComboBox. Handles 'MatchCase' and 'Partial' Searching.
'               Will not alter position if string is not found.
' Notes:        None
' Example:      blnRetVal = FindInCombo(cboStateCode, "nc", False, False, True)
' Inputs:       Name                    Type        Optional    Default     Desc
'               cboComboBox             ComboBox    False       Nothing     ComboBox's Internal Name
'               strStringToFind         String      False       ""          String to find in the ComboBox
'               blnPartialMatch         Boolean     True        False       If 'True' will search to parttial match
'               blnMatchCase            Boolean     True        True        If 'True' will only match on exact case
'               blnShowErrors           Boolean     True        False       If 'True' will display errors
' Returns:      Boolean (True/False)
'___________________________________________________________________________________________________________________
Private Function FindInCombo(cboComboBox As ComboBox, strStringToFind As String, Optional blnPartialMatch As Boolean, Optional blnMatchCase As Boolean = True, Optional blnShowErrors As Boolean) As Boolean
' init the error handler
On Error GoTo err_FindInCombo
    ' set the function's return to true
    FindInCombo = True
    
    ' define variables
    Dim intTempIndex    As Integer
    Dim intIndex        As Integer
    
    ' if there are no items in the combobos then exit
    If cboComboBox.ListCount < 0 Then Exit Function
    
    ' store the current index on the combobox
    intTempIndex = cboComboBox.ListIndex
    
    ' move through the items inthe combobox
    For intIndex = 0 To cboComboBox.ListCount - 1
        ' set the combobox's listindex to the current intindex
        cboComboBox.ListIndex = intIndex
        ' if we want a partial match...
        If blnPartialMatch = True Then
            ' ...and we want the case to be the same...
            If blnMatchCase = True Then
                ' if a match is found then exit
                If InStr(cboComboBox.Text, strStringToFind) > 0 Then Exit Function
            Else
                '...but we do not want the case to match....
                ' if a match is found then exit
                If InStr(UCase(cboComboBox.Text), UCase(strStringToFind)) > 0 Then Exit Function
            End If
            
        Else
            ' if we do not want a partial match...
            ' ...and we want the case to be the same...
            If blnMatchCase = True Then
                'Default
                ' if a match is found then exit
                If strStringToFind = cboComboBox.Text Then Exit Function
            Else
                '...but we do not want the case to match....
                ' if a match is found then exit
                If UCase(strStringToFind) = UCase(cboComboBox.Text) Then Exit Function
            End If
            
        End If
    ' move to the next index
    Next intIndex
    
    ' set the combobox's index back to the original value
    cboComboBox.ListIndex = intTempIndex
    
    ' exit the function
    Exit Function
' error handler
err_FindInCombo:
    ' set the function's return to false
    FindInCombo = False
    ' send a message to the immediate window
    'Debug.Print Now & " | Function: FindInCombo | Error: #" & Err.Number & vbTab & Err.Description
    ' hard-breakpoint
    'Debug.Assert False
    ' if we wanted to see an error..
    If blnShowErrors = True Then
        MsgBox "Error: #" & Err.Number & vbTab & Err.description & vbCrLf & vbCrLf & Now, vbOKOnly + vbCritical, "Function: FindInCombo"
    End If
    ' clear the error object
    Err.Clear
    
End Function

Private Sub txtDir_GotFocus()
txtDir.BackColor = vbYellow
End Sub

Private Sub txtDir_LostFocus()
txtDir.BackColor = &HC0E0FF
End Sub
Private Sub txtExt_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then Call cmdSeek_Click

End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtApellido.SetFocus
    txtApellido.SelStart = 0
    txtApellido.SelLength = Len(txtApellido.Text)
End If
End Sub
Private Sub txtApellido_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtEmpresa.SetFocus
    txtEmpresa.SelStart = 0
    txtEmpresa.SelLength = Len(txtEmpresa.Text)
End If
End Sub
Private Sub txtEmpresa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmbZona.SetFocus
End Sub
Private Sub cmbZona_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtDir.SetFocus
End Sub
'Private Sub txtDir1_KeyPress(KeyAscii As Integer)
'If KeyAscii = 8 Then Exit Sub
'If KeyAscii = 13 Then txtDir2.SetFocus
'If Len(txtDir1.Text) = 44 Then txtDir2.SetFocus
'End Sub
'
'Private Sub txtDir2_KeyPress(KeyAscii As Integer)
'If KeyAscii = 8 Then Exit Sub
'If KeyAscii = 13 Then txtDir3.SetFocus
'If Len(txtDir2.Text) = 44 Then txtDir3.SetFocus
'End Sub
Private Sub txtDir_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtEmail.SetFocus
End Sub

Private Sub txteMail_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtFax.SetFocus
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtNotas.SetFocus
End Sub

Private Sub txtNotas_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then cmdGuardar.SetFocus
End Sub
Private Sub txtNumber_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call cmdSeek_Click
End Sub
Private Function RemoveNULL(cTexto As String) As String
Dim i As Integer

For i = 1 To Len(cTexto)
    Select Case Mid(cTexto, i, 1)
        Case Chr(0)
           Mid(cTexto, i, 1) = "-"
        Case Chr(10) + Chr(13)
           Mid(cTexto, i, 1) = "-"
        Case Chr(13) + Chr(10)
           Mid(cTexto, i, 1) = "-"
        Case Chr(10)
            Mid(cTexto, i, 1) = "-"
        Case Chr(13)
            Mid(cTexto, i, 1) = "-"
    End Select
Next
RemoveNULL = cTexto
End Function

