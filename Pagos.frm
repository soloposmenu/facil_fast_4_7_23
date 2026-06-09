VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form Pagos 
   BackColor       =   &H00B39665&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MODULO DE FACTURACION GENERAL"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8670
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Pagos.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   8670
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid ListaPagos 
      Height          =   1335
      Left            =   6360
      TabIndex        =   17
      Top             =   5760
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2355
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdCARGOS 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   715
      Index           =   0
      Left            =   240
      TabIndex        =   30
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdFPagos 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   715
      Index           =   0
      Left            =   240
      TabIndex        =   28
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CheckBox chkInfo 
      BackColor       =   &H00B39665&
      Caption         =   "Información del Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   2040
      TabIndex        =   26
      ToolTipText     =   "Incluir Informacion del Cliente"
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cdmBill 
      Height          =   615
      Index           =   0
      Left            =   5040
      Picture         =   "Pagos.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   25
      Tag             =   "1.00"
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cdmBill 
      Height          =   615
      Index           =   1
      Left            =   6840
      Picture         =   "Pagos.frx":172C
      Style           =   1  'Graphical
      TabIndex        =   24
      Tag             =   "5.00"
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cdmBill 
      Height          =   615
      Index           =   2
      Left            =   5040
      Picture         =   "Pagos.frx":2E0B
      Style           =   1  'Graphical
      TabIndex        =   23
      Tag             =   "10.00"
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cdmBill 
      Height          =   615
      Index           =   3
      Left            =   6840
      Picture         =   "Pagos.frx":449C
      Style           =   1  'Graphical
      TabIndex        =   22
      Tag             =   "20.00"
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdDescGlob 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Descuento Global"
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
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H0000FF00&
      Caption         =   "REGRESAR SIN FACTURAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   715
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4750
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00B39665&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Index           =   3
      Left            =   6120
      TabIndex        =   0
      Top             =   1800
      Width           =   2415
      Begin VB.CommandButton Clear 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1560
         Picture         =   "Pagos.frx":5B64
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3000
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   10
         Left            =   840
         TabIndex        =   11
         Top             =   3000
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   3000
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   9
         Left            =   1560
         TabIndex        =   10
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   8
         Left            =   840
         TabIndex        =   9
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   6
         Left            =   1560
         TabIndex        =   7
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   840
         TabIndex        =   6
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   1560
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   840
         TabIndex        =   3
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lbMonto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """B/."" #,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   6154
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   170
         Width           =   2175
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFClientes 
      Height          =   255
      Left            =   2760
      TabIndex        =   21
      Top             =   1440
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   393216
      Rows            =   0
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.Label Label4 
      BackColor       =   &H00008000&
      Caption         =   "Formas de Pago"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """B/."" #,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   6154
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   29
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      Height          =   2865
      Index           =   1
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   5940
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C00000&
      Caption         =   "CARGOS"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """B/."" #,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   6154
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   27
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   1905
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   5520
      Width           =   5940
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   19
      Top             =   840
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   1575
      Left            =   4920
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   20
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label lbPend 
      BackColor       =   &H00B39665&
      Caption         =   "Monto Pendiente"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """B/."" #,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   6154
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label lbFact 
      BackColor       =   &H00B39665&
      Caption         =   "Total Factura"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """B/."" #,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   6154
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "Pagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nfPase As Integer
Dim nMntOculto As String
Dim RSPAGOS As Recordset    'Pagos
Dim OrigSB As Single
Dim nFlagParciales As Integer
Dim RSPROPINAS As Recordset   'Propinas
Dim aInfoCliente(1) As String                                       '07 OCTUBRE 2011
Dim rsPropinaMarcadas As New ADODB.Recordset        '07 OCTUBRE 2011
Dim rsPagosMarcados As New ADODB.Recordset           '07 OCTUBRE 2011
Dim RSCARGOS As Recordset   'INFO 17032019
Dim nMontoRecibido As Single
Dim nTaxRecibido As Single

Private Sub chkInfo_Click()

'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
If FE_Facil_ONLINE Then
    If Pagos.chkInfo.value = 1 Then
        Cliente.Show 1
    End If
Else
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    
    ShowMsg "EN MODO (OFF-LINE)" & vbCrLf & "ES IMPOSIBLE" & vbCrLf & "GRABAR LA INFORMACION DEL CLIENTE EN LA FACTURA ELECTRONICA, UNICAMENTE EN EL TICKET IMPRESO", vbBlue, vbCyan
    
    'INFO: SEPTIEMBRE 2010. GUARDA e IMPRIME LA INFORMACION DEL CLIENTE
    If chkInfo.value = 1 Then
        cNombreCliente = InputBox("ESCRIBA EL NOMBRE DEL CLIENTE o EMPRESA", "ESCRIBIR NOMBRE DEL CLIENTE o EMPRESA", aInfoCliente(0))
        cRUCCliente = InputBox(" <<< ESCRIBA EL NUMERO DE CEDULA / R.U.C. >>>", "ESCRIBIR CEDULA o R.U.C. DEL CLIENTE", aInfoCliente(1))
        aInfoCliente(0) = cNombreCliente
        aInfoCliente(1) = cRUCCliente
    Else
        cNombreCliente = ""
        cRUCCliente = ""
        aInfoCliente(0) = ""
        aInfoCliente(1) = ""
    End If
    
    'INFO: 18SEP2012
    RegWrite "HKCU\Software\SoloSoftware\SoloMix\LastCliente", cNombreCliente
    RegWrite "HKCU\Software\SoloSoftware\SoloMix\LastRUC", cRUCCliente
End If
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||



''INFO: 07OCT2011. GUARDA e IMPRIME LA INFORMACION DEL CLIENTE
'If chkInfo.value = 1 Then
'    aInfoCliente(0) = InputBox("ESCRIBA EL NOMBRE DEL CLIENTE", "ESCRIBIR NOMBRE DEL CLIENTE", aInfoCliente(0))
'    aInfoCliente(1) = InputBox(" <<< ESCRIBA EL NUMERO DE CEDULA / R.U.C. >>>", "ESCRIBIR CEDULA o R.U.C. DEL CLIENTE", aInfoCliente(1))
'Else
'    aInfoCliente(0) = ""
'    aInfoCliente(1) = ""
'End If
'
''INFO: 18SEP2012
'RegWrite "HKCU\Software\SoloSoftware\SoloMix\LastCliente", cNombreCliente
'RegWrite "HKCU\Software\SoloSoftware\SoloMix\LastRUC", cRUCCliente

End Sub
Private Function PrepareRecordSets()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' DE PAGOS y PROPINAS MARCADAS
' INFO: 07OCT2011
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PAGOS:
'NO --->    ListaPagos.AddItem cmdFPagos(Index).Tag & Chr(9) & cmdFPagos(Index).Caption & Chr(9) & Format((lbMonto / nTasa) - (nCambio / nTasa), "STANDARD") & Chr(9) & Format(lbMonto, "STANDARD")
'ListaPagos.AddItem cmdFPagos(Index).Tag & Chr(9) & cmdFPagos(Index).Caption & Chr(9) & Format(lbMonto - nCambio, "STANDARD") & Chr(9) & Format(lbMonto, "STANDARD")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'PROPINAS
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'EN EFECTIVO
'ListaPagos.AddItem "20" & Chr(9) & "PROPINA EN EFECTIVO" & Chr(9) & Format(lbMonto, "STANDARD") & Chr(9) & Format(lbMonto, "STANDARD")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'EN CREDITO y OTROS
'ListaPagos.AddItem cmdPropina(Index).Tag & Chr(9) & PROPINA_DESCRIP & " " & cmdPropina(Index).Caption & Chr(9) & Format(lbMonto, "STANDARD") & Chr(9) & Format(lbMonto, "STANDARD")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

If rsPropinaMarcadas.State = adStateOpen Then
Else
    With rsPropinaMarcadas      'PROPINAS EN EFECTIVO y OTROS
        .Fields.Append "CODIGO", adInteger, , adFldUpdatable
        .Fields.Append "DESCRIP_1", adVarWChar, 60, adFldUpdatable
        .Fields.Append "MONTO_1", adSingle, , adFldUpdatable
        .Fields.Append "MONTO_2", adSingle, , adFldUpdatable
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
End If

If rsPagosMarcados.State = adStateOpen Then
Else
    With rsPagosMarcados        'PAGOS RECIBIDOS
        .Fields.Append "CODIGO", adInteger, , adFldUpdatable
        .Fields.Append "DESCRIP_1", adVarWChar, 60, adFldUpdatable
'        .Fields.Append "MONTO_1", adSingle, , adFldUpdatable        'lbMonto - nCambio
'        .Fields.Append "MONTO_2", adSingle, , adFldUpdatable
        '.Fields.Append "MONTO_1", adNumeric, , adFldUpdatable
        .Fields.Append "MONTO_1", adNumeric, , adFldUpdatable
        .Fields("MONTO_1").Precision = 10
        .Fields("MONTO_1").NumericScale = 5
        .Fields.Append "MONTO_2", adNumeric, , adFldUpdatable
        .Fields("MONTO_2").Precision = 10
        .Fields("MONTO_2").NumericScale = 5
        '.Fields.Append "MONTO_2", adNumeric, , adFldUpdatable
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        'INFO: UPDATE 10SEPT2024. SE VA A GUARDAR LA INFO DE LA FE ACA DE UNA VEZ
        .Fields.Append "CODIGO_FE", adVarWChar, 2, adFldUpdatable
        .Fields.Append "DESCRIP_FE", adVarWChar, 60, adFldUpdatable
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
End If
End Function

Private Function AddPago(cTipo As String, nCodigo As Integer, cDescrip As String, nMonto1 As Single, nMonto2 As Single) As Boolean
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' DE PAGOS y PROPINAS MARCADAS
' INFO: 07OCT2011
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Select Case cTipo
    Case "PAGO"
        With rsPagosMarcados
            .AddNew
            !CODIGO = nCodigo           'CODIGO EN DB
            !DESCRIP_1 = cDescrip       'DESCRIP EN DB
            !MONTO_1 = nMonto1
            !MONTO_2 = nMonto2
            
            '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
            '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
            'UPDATE 10SEP2024
            If FE_Facil_ONLINE Then
                !CODIGO_FE = GetPagoElectronico(nCodigo)
            Else
                !CODIGO_FE = GetPagoElectronico_OFFLINE(nCodigo)
            End If
            '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
            '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
            
            !DESCRIP_FE = cDescrip & " / " & cDescrip
            .Update
        End With
    Case "PROPINA"
        With rsPropinaMarcadas
            .AddNew
            !CODIGO = nCodigo
            !DESCRIP_1 = cDescrip
            !MONTO_1 = nMonto1
            !MONTO_2 = nMonto2
            .Update
        End With
    Case Else
End Select

End Function

Private Sub Actualizador(fSubtotal As Single, fTax As Single, Optional nImpuesto As Single)

Dim rsAcutalizacion As Recordset
Dim rsTrans As Recordset
Dim cSQL As String, cSQL2 As String
Dim ImpText As String
Dim MiValor As Currency
Dim nValorPago As Single
Dim nTipoPago As Integer
Dim i As Integer
Dim iError As Integer
Dim ccLERR As String

iError = 0

Set rsAcutalizacion = New Recordset

'Actualiza los valores de la factura
'''INCREMENTA EL NUMERO DE TRANSACCION EN 1
''msConn.BeginTrans
''msConn.Execute "UPDATE ORGANIZACION SET TRANS = TRANS + 1"
''msConn.CommitTrans

'AUMENTA E INCREMENTA LOS VALORES POR DEPARTAMENTO
'AUMENTA E INCREMENTA LOS VALORES POR PLATO (PLU)
'INFO: MULTICAJAS
'cSQL = "SELECT * FROM TMP_TRANS WHERE VALID AND CANT >= 0 AND  MESA = " & nMesa
'INFO: SQL SERVER (2005) DIC2009
cSQL = "SELECT * FROM TMP_TRANS "
cSQL = cSQL & "WHERE VALID = 1 AND CANT >= 0 AND  MESA = " & nMesa
cSQL = cSQL & " AND CAJERO = " & npNumCaj
cSQL = cSQL & " ORDER BY DEPTO, PLU"

rsAcutalizacion.Open cSQL, msConn, adOpenStatic, adLockReadOnly

msConn.BeginTrans

Do Until rsAcutalizacion.EOF
    
    If IsEmpty(rsAcutalizacion!precio) Then GoTo Proximo:
    
    MiValor = Format(rsAcutalizacion!precio, "#0.00")
    'OK Para DEPTO
    cSQL = "UPDATE DEPTO SET X_COUNT = X_COUNT + " & rsAcutalizacion!CANT
    cSQL = cSQL & " , Z_COUNT = Z_COUNT + " & rsAcutalizacion!CANT
    cSQL = cSQL & " , VALOR = VALOR + " & MiValor
    cSQL = cSQL & " , X_PERIOD_CNT = X_PERIOD_CNT + " & rsAcutalizacion!CANT
    cSQL = cSQL & " , Z_PERIOD_CNT = Z_PERIOD_CNT + " & rsAcutalizacion!CANT
    cSQL = cSQL & " , PERIOD_VAL = PERIOD_VAL + " & MiValor
    cSQL = cSQL & " WHERE CODIGO = " & rsAcutalizacion!depto
    msConn.Execute cSQL
    'OK para PLU
    cSQL = "UPDATE PLU SET X_COUNT = X_COUNT + " & rsAcutalizacion!CANT
    cSQL = cSQL & " , Z_COUNT = Z_COUNT + " & rsAcutalizacion!CANT
    cSQL = cSQL & " , VALOR = VALOR + " & MiValor
    cSQL = cSQL & " , X_PERIOD_CNT = X_PERIOD_CNT + " & rsAcutalizacion!CANT
    cSQL = cSQL & " , Z_PERIOD_CNT = Z_PERIOD_CNT + " & rsAcutalizacion!CANT
    cSQL = cSQL & " , PERIOD_VAL = PERIOD_VAL + " & MiValor
    cSQL = cSQL & " WHERE CODIGO = " & rsAcutalizacion!PLU
    msConn.Execute cSQL
    
    'OK Para CONTEND_02
    cSQL = "UPDATE CONTEND_02 SET X_COUNT = X_COUNT + " & rsAcutalizacion!CANT
    cSQL = cSQL & " , Z_COUNT = Z_COUNT + " & rsAcutalizacion!CANT
    cSQL = cSQL & " , VALOR = VALOR + " & MiValor
    cSQL = cSQL & " , X_PERIOD_CNT = X_PERIOD_CNT + " & rsAcutalizacion!CANT
    cSQL = cSQL & " , Z_PERIOD_CNT = Z_PERIOD_CNT + " & rsAcutalizacion!CANT
    cSQL = cSQL & " , PERIOD_VAL = PERIOD_VAL + " & MiValor
    cSQL = cSQL & " WHERE CODIGO = " & rsAcutalizacion!PLU & " AND "
    cSQL = cSQL & " CONTENEDOR = " & rsAcutalizacion!envase
    msConn.Execute cSQL
Proximo:
    rsAcutalizacion.MoveNext
Loop

msConn.CommitTrans

'///////////////////////////////////////////////////////////////
msConn.BeginTrans

'AUMENTA E INCREMENTA LOS VALORES POR CAJERO
cSQL = "UPDATE CAJEROS SET X_COUNT = X_COUNT + 1"
cSQL = cSQL & " , Z_COUNT = Z_COUNT + 1 "
cSQL = cSQL & " , VALOR = VALOR + " & Format(Label3, "#0.00")
cSQL = cSQL & " WHERE NUMERO = " & npNumCaj & " OR NUMERO = " & 999
'Pagos.Label3
msConn.Execute cSQL

'AUMENTA E INCREMENTA LOS VALORES POR MESEROS
cSQL = "UPDATE MESEROS SET X_COUNT = X_COUNT + 1"
cSQL = cSQL & " , Z_COUNT = Z_COUNT + 1 "
cSQL = cSQL & " , VALOR = VALOR + " & Format(Label3, "#0.00")
cSQL = cSQL & " WHERE NUMERO = " & nMesero & " OR NUMERO = " & 999
msConn.Execute cSQL

'cSQL = "UPDATE MESAS SET VALOR = VALOR + " & Format(Pagos.Label3, "#0.00")
cSQL = "UPDATE MESAS SET VALOR = VALOR + " & Format(Label3, "#0.00")
cSQL = cSQL & ", MESERO_ACTUAL = 0 WHERE NUMERO = " & nMesa & " OR "
cSQL = cSQL & " NUMERO = -99 "
msConn.Execute cSQL

msConn.CommitTrans

msConn.BeginTrans

For i = 0 To (ListaPagos.Rows - 1)
    On Error GoTo ErrAdm:
        ListaPagos.Row = i
        ListaPagos.Col = 0
        nTipoPago = ListaPagos.Text
        ListaPagos.Col = 2
    On Error GoTo 0
    nValorPago = Format(ListaPagos.Text, "STANDARD")
    
    cSQL = "UPDATE PAGOS SET X_COUNT = X_COUNT + 1"
    cSQL = cSQL & " , Z_COUNT = Z_COUNT + 1 "
    cSQL = cSQL & " , VALOR = VALOR + " & Format(nValorPago, "#0.00")
    cSQL = cSQL & " , X_PERIOD_CNT = X_PERIOD_CNT + 1"
    cSQL = cSQL & " , Z_PERIOD_CNT = Z_PERIOD_CNT + 1"
    cSQL = cSQL & " , PERIOD_VAL = PERIOD_VAL + " & Format(nValorPago, "#0.00")
    cSQL = cSQL & " WHERE CODIGO = " & nTipoPago & " OR CODIGO = " & 999
    msConn.Execute cSQL
    
    
    'SI ES UN PAGO COMPUESTO LO HACE AQUI. SI ES SIMPLE LO HACE EN LA PANTLLA DE PLU.
    'INFO: ENE2010 = ARGREGANDO LA MESA A TRANSAC_PAGO
    cSQL = "INSERT INTO TRANSAC_PAGO "
    cSQL = cSQL & " (NUM_TRANS,TIPO_PAGO,CAJERO,LIN,MONTO, MESA) VALUES ("
    cSQL = cSQL & nNUMTrans & "," & nTipoPago & "," & npNumCaj & ","
    cSQL = cSQL & (i + 1) & "," & Format(nValorPago, "#0.00") & "," & nMesa & ")"
    msConn.Execute cSQL

    On Error GoTo ErrAdm:
        ListaPagos.Col = 1
    On Error GoTo 0

    'SI HAY PROPINAS, MARCAR PARA PAGAR A MESEROS
    'If Mid(ListaPagos.Text, 1, 7) = "PROPINA" Then
    If Mid(ListaPagos.Text, 1, 5) = "CARGO" Then
        cSQL = "INSERT INTO TRANSAC_PROP "
        cSQL = cSQL & " (NUM_TRANS,MESERO,CAJERO,TIPO_PAGO,LIN,MONTO) VALUES ("
        cSQL = cSQL & nNUMTrans & "," & nMesero & "," & npNumCaj & "," & nTipoPago & ","
        cSQL = cSQL & i + 1 & "," & Format(nValorPago, "#0.00") & ")"
        
        msConn.Execute cSQL
    End If
Next

msConn.CommitTrans

ccLERR = "INSERT INTO TRANSAC_FISCAL."
'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
'INFO: 26OCT2011
'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
If cFullFiscal = "SI" Then
    msConn.BeginTrans
    
    cSQL = "INSERT INTO TRANSAC_FISCAL (FISCAL, DOC_SOLO, FECHA, HORA, SUB_TOTAL, DESCUENTO, ITBM)"
    cSQL = cSQL & " VALUES ("
    cSQL = cSQL & "'" & FISCAL_.GetLastFactura & "'," & nNUMTrans & ",'"
    cSQL = cSQL & Format(Date, "YYYYMMDD") & "','" & Format(Time, "HH:MM") & "',"
    'INFO: 24NOV2011
    cSQL = cSQL & FISCAL_.GetCurrentSubTotal & "," & 0 & "," & FISCAL_.GetCurrentITBMS & ")"
    'cSQL = cSQL & CSng(PLU.SubTot) & "," & 0 & "," & Round(nImpuesto, 2) & ")"
    
    msConn.Execute cSQL
    msConn.CommitTrans
Else
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||


    ArregloPDF_URL = Split(FE_pdfURL, "/")

    msConn.BeginTrans

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'INFO: PARA FACTURA ELECTRONICA. SE CAMBIA LA ESTRUCTURA DE TRANSAC_FISCAL
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    cSQL = "INSERT INTO TRANSAC_FISCAL (FISCAL, DOC_SOLO, FECHA, HORA, SUB_TOTAL, DESCUENTO, ITBM, ID_URL)"
    cSQL = cSQL & " VALUES ("
    cSQL = cSQL & "'" & FE_CUFE & "'," & nNUMTrans & ",'"
    cSQL = cSQL & Format(Date, "YYYYMMDD") & "','" & Format(Time, "HH:MM") & "',"

    If FE_Facil_ONLINE Then
        cSQL = cSQL & nMontoRecibido & "," & 0 & "," & (nTaxRecibido) & ",'" & ArregloPDF_URL(6) & "')"
    Else
        cSQL = cSQL & nMontoRecibido & "," & 0 & "," & (nTaxRecibido) & ",'" & "OFF-LINE" & "')"
    End If
    ''If nDescAplicado > 0# Then
        'cSQL = cSQL & nMontoRecibido & "," & 0 & "," & (nTaxRecibido) & ",'" & ArregloPDF_URL(6) & "')"
    ''Else
        ''cSQL = cSQL & lbNOTax & "," & 0 & "," & (lbTAX) & ",'" & ArregloPDF_URL(6) & "')"
    ''End If

    msConn.Execute cSQL

    msConn.CommitTrans

    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
End If
'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~

'INFO: 24OCT2017
If nTransDomicilio Then Call PutDomicilioInfo(nNUMTrans, Format(Label3, "#0.00"))
If nTransDomicilio Then Call UpdateDomiTrans(nNUMTrans, Format(Label3, "#0.00"))

nTransDomicilio = False

'INFO: 22DIC2014
On Error GoTo ErrAdm:

msConn.BeginTrans

'ANEXA LAS TRANSACCIONES AL ARCHIVO DE TRANSACCIONES

Set rsTrans = New Recordset

cSQL = "SELECT * FROM TMP_TRANS "
cSQL = cSQL & " WHERE MESA = " & nMesa
cSQL = cSQL & " AND CAJERO = " & npNumCaj
cSQL = cSQL & " ORDER BY LIN "

rsTrans.Open cSQL, msConn, adOpenStatic, adLockReadOnly
Dim MiFecha As String   'PARA CLIENTES

ccLERR = "INSERT INTO TRANSAC."
Do Until rsTrans.EOF
    cSQL2 = "INSERT INTO TRANSAC "
    cSQL2 = cSQL2 & "(NUM_TRANS,CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT,DEPTO,PLU,ENVASE,PRECIO_UNIT,PRECIO,"
    cSQL2 = cSQL2 & "FECHA,HORA,TIPO,DESCUENTO,CON_TAX, FECHA_TRANS, HORA_TRANS) VALUES ("
    cSQL2 = cSQL2 & "" & nNUMTrans & "," & rsTrans!CAJA & "," & rsTrans!CAJERO & "," & rsTrans!MESA & "," & rsTrans!MESERO & "," & GetBoolean(rsTrans!VALID) & "," & rsTrans!LIN & ",'"
    cSQL2 = cSQL2 & rsTrans!DESCRIP & "'," & rsTrans!CANT & "," & rsTrans!depto & "," & rsTrans!PLU & ","
    cSQL2 = cSQL2 & rsTrans!envase & "," & rsTrans!precio_unit & "," & rsTrans!precio & ",'" & rsTrans!FECHA & "','" & Time & "'"
    cSQL2 = cSQL2 & ",'" & rsTrans!Tipo & "'," & rsTrans!DESCUENTO
    cSQL2 = cSQL2 & "," & rsTrans!CON_TAX & ",'"
    'INFO: ENERO 2010
    cSQL2 = cSQL2 & Format(Date, "YYYYMMDD") & "','" & Time & "')"
    msConn.Execute cSQL2
    
    MiFecha = rsTrans!FECHA
    rsTrans.MoveNext
Loop

'BORRA REGISTROS DE LA TEMPORAL
ccLERR = "DELETE FROM TMP_TRANS."
cSQL = "DELETE FROM TMP_TRANS WHERE MESA = " & nMesa & " AND CAJERO = " & npNumCaj
msConn.Execute cSQL
msConn.CommitTrans

msConn.BeginTrans

''  msConn.Execute cSQL (NO ES NECESARIO)
''''''''''''''------------- msConnLoc.Execute cSQL
'ACTUALIZA MESAS
ccLERR = "UPDATE Mesas."
msConn.Execute "UPDATE Mesas SET ocupada = 0, MESERO_ACTUAL = 0 WHERE numero = " & nMesa
'CON PAGOS A CREDITO. INSERTA INFO. DEL GRID INVISIBLE

msConn.CommitTrans
msConn.BeginTrans

If MSHFClientes.Rows > 0 Then
    Dim nVal1 As Integer
    Dim nVal2 As Integer
    Dim nVal3 As Single
    Dim nVal4 As Single
    Dim rsCli As Recordset
    
    Set rsCli = New Recordset
    For i = 0 To (MSHFClientes.Rows - 1)
        On Error GoTo ErrAdm:
            MSHFClientes.Row = i
            MSHFClientes.Col = 0: nVal1 = MSHFClientes.Text 'CODIGO CLIENTE
            MSHFClientes.Col = 1: nVal2 = MSHFClientes.Text
            MSHFClientes.Col = 2: nVal3 = MSHFClientes.Text
            'NVAL3 ES LO QUE SE ESTA CARGANDO A LA CUENTA
        On Error GoTo 0
        rsCli.Open "SELECT * FROM CLIENTES WHERE CODIGO = " & nVal2, msConn, adOpenStatic, adLockOptimistic

        nVal4 = 0#
        If Not rsCli.EOF Then
            If rsCli!saldo < 0# Then
                If Abs(rsCli!saldo) > nVal3 Then
                    nVal4 = nVal3
                ElseIf Abs(rsCli!saldo) < nVal3 Then
                    nVal4 = Abs(rsCli!saldo)
                ElseIf Abs(rsCli!saldo) = nVal3 Then
                    nVal4 = Abs(rsCli!saldo)
                End If
            End If
        End If
        
        cSQL = "INSERT INTO TRANSAC_CLI "
        cSQL = cSQL & " (CODIGO_TP,CODIGO_CLI,NUM_TRANS,MONTO,FECHA,RECIBIDO) "
        cSQL = cSQL & " VALUES ("
        cSQL = cSQL & nVal1 & "," & nVal2 & "," & nNUMTrans & "," & nVal3 & ",'"
        cSQL = cSQL & MiFecha & "'," & nVal4 & ")"
        
        msConn.Execute cSQL
        
        cSQL = "UPDATE CLIENTES SET SALDO = SALDO + " & Format(nVal3, "#0.00") & " WHERE CODIGO = " & nVal2
        msConn.Execute cSQL
        rsCli.Close
    Next
End If

On Error GoTo ErrAdm:

msConn.CommitTrans
'msConn.BeginTrans

'MuestraProductosMarcados

'LA INSTRUCCION CommitTrans EJECUTA TODAS ESTAS ACTUALIZACIONES

'msConn.CommitTrans
ccLERR = "DELETE FROM TMP_PAR_PAGO."
msConn.BeginTrans
msConn.Execute "DELETE FROM TMP_PAR_PAGO WHERE MESA = " & nMesa
msConn.Execute "DELETE FROM TMP_PAR_PROP WHERE MESA = " & nMesa
msConn.Execute "DELETE FROM TMP_CLI WHERE MESA = " & nMesa
msConn.CommitTrans

rsAcutalizacion.Close
rsTrans.Close
Set rsAcutalizacion = Nothing
Set rsTrans = Nothing

On Error GoTo 0

Exit Sub

ErrAdm:

If iError < 4 Then
    ''EscribeLog ("Procedimiento de Captura de Error. Con Error menor que 4")
    iError = iError + 1
    Resume
Else
    EscribeLog "Error.Pagos.Actualizador." & ccLERR & Err.Number & " - " & Err.description
    Dim OBJERR As Error
    For Each OBJERR In msConn.Errors
        MsgBox OBJERR.description, vbCritical, "OCURRIO UN ERROR, ANOTE LOS DATOS EN PANTALLA"
    Next
    Exit Sub
End If
End Sub
Private Sub ImprFactura()
Dim i As Integer
Dim nMiSub As Single
Dim nCodigoPago As Integer
Dim sqltext As String
Dim LinTx As String
Dim rsCuenta As Recordset
Dim MiMatriz(0, 3) As String
Dim MiLen1, Milen2 As Integer
Dim n1 As Single
Dim n2 As Single
Dim nImp As Integer
Dim iSlip As Integer
Dim nEspacio As Integer
Dim nLinDetalle As Integer
Dim STATION_2PRINT As Integer
Dim LOCAL_ISC As Single
    
LOCAL_ISC = iISCTransaccion

nImp = 0: nEspacio = 0: nLinDetalle = 0
STATION_2PRINT = FptrSReceipt
For nImp = 0 To 1
    
    Set rsCuenta = New Recordset
    nMiSub = 0#
    sqltext = "SELECT * FROM TMP_TRANS WHERE MESA = " & nMesa & " AND CAJERO = " & npNumCaj & " ORDER BY LIN "
    rsCuenta.Open sqltext, msConn, adOpenStatic, adLockReadOnly
    
    If nImp = 0 Then
        'Sys_Pos.CoCash1.OpenDrawer     'INFO: 14FEB2017. USANDO DIRECTIO QUE NO ESPERA RESPUESTA DE LA GAVETA
        'Call Sys_Pos.Cocash1.DirectIO(DRW_DI_OPEN_DRAWER, 0, "")
        Call OpenGavetaDinero
        
        If cNombreCliente <> "" Then Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, "Cliente : " & cNombreCliente & Chr(&HD) & Chr(&HA)
    Else
        nEspacio = 0
        STATION_2PRINT = FPTR_S_JOURNAL
    End If
    'SUBTOTAL DE LOS PLATOS
    Do Until rsCuenta.EOF
        If SLIP_OK = True And nEspacio = 16 Then
            MiMatriz(0, 0) = FormatTexto(rsCuenta!DESCRIP, 35)
        Else
            MiMatriz(0, 0) = FormatTexto(rsCuenta!DESCRIP, 15)
        End If
        MiMatriz(0, 1) = Format(rsCuenta!CANT, "general number")
        MiMatriz(0, 2) = Format(rsCuenta!precio, "#,###.00")
        nMiSub = nMiSub + rsCuenta!precio
        MiLen1 = Len(MiMatriz(0, 1))
        Milen2 = Len(MiMatriz(0, 2))
        LinTx = MiMatriz(0, 0) & Space(5 - MiLen1) & MiMatriz(0, 1) & Space(10 - Milen2) & MiMatriz(0, 2)
        ''Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, LinTx & Chr(&HD) & Chr(&HA)
        nLinDetalle = nLinDetalle + 1
        rsCuenta.MoveNext
    Loop
    Milen2 = Len(Format(nMiSub, "CURRENCY"))
    'Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)
    If nImp = 0 Then Sys_Pos.Coptr1.PrintNormal PtrSReceipt, "------------------------------" & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, "SUB-TOTAL :" & Space(19 - Milen2) & Format(nMiSub, "CURRENCY") & Chr(&HD) & Chr(&HA)
    
    '~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
    '~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
    iLen = Len(Format(LOCAL_ISC, "STANDARD"))
    'INFO: JULIO 2010, NO SE PUEDE PONER 5%, YA QUE EL ITBMS ES CALCULADO x PRODUCTO
    'txtString = "ITBMS(5%) : " & Space(18 - iLen) & Format(LOCAL_ISC, "STANDARD")
    'txtString = "ITBMS        : " & Space(18 - iLen) & Format(LOCAL_ISC, "STANDARD")
    'INFO: 12SEP2013
    txtString = "ITBMS     : " & Space(18 - iLen) & Format(LOCAL_ISC, "STANDARD")
                    
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, txtString & Chr(&HD) & Chr(&HA)

    Call PutISC(Format(iISCTransaccion, "STANDARD"))

    iISCTransaccion = 0

    iLen = Len(Format(Label3, "STANDARD"))
    txtString = "TOTAL     : " & Space(18 - iLen) & Format(Label3, "STANDARD")
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, txtString & Chr(&HD) & Chr(&HA)
    '~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
    '~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
    
    If nImp = 0 Then Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)
    If nFlagParciales = 1 Then
        Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, "** Detalle de Pagos Parciales" & Chr(&HD) & Chr(&HA)
    End If
    
    For i = 0 To (ListaPagos.Rows - 1)
        ListaPagos.Row = i
        ListaPagos.Col = 0
        nCodigoPago = ListaPagos.Text
        ListaPagos.Col = 1
        MiMatriz(0, 0) = FormatTexto(ListaPagos.Text, 15)
        ListaPagos.Col = 2
        If nCodigoPago = 99 Then
            MiMatriz(0, 1) = Format(ListaPagos.Text * (-1), "##,##0.00")
        Else
            ListaPagos.Col = 2
            n1 = Format(ListaPagos.Text, "##,##0.00")
            ListaPagos.Col = 3
            n2 = Format(ListaPagos.Text, "##,##0.00")
            If n1 <> n2 Then
                MiMatriz(0, 1) = Format(n2, "##,##0.00")
            Else
                MiMatriz(0, 1) = Format(n1, "##,##0.00")
            End If
        End If
        Milen2 = Len(MiMatriz(0, 1))
        Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, MiMatriz(0, 0) & Space(15 - Milen2) & MiMatriz(0, 1) & Chr(&HD) & Chr(&HA)
    Next
    
    If nImp = 0 Then Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)
    
    Milen2 = Len(Format(nCambio, "CURRENCY"))
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, "CAMBIO    : " & Space(18 - Milen2) & Format(nCambio, "CURRENCY") & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Date & Space(2) & Time & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, "SERIAL:" & rs00!SERIAL & Chr(&HD) & Chr(&HA)
    'INFO: AGO2009
    'Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, "TRANS# " & nNumTrans & Chr(&HD) & Chr(&HA)
    
    '~|~|~|~|~|~|~|~|~|~|~|~|~|
    'INFO: FEB2009
    If nImp = 0 Then
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
                        'Call GetNewNumTrans
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        
        '~~~~~~~~~~~~~~~~~~~~~~
        'INFO: 6JUL2014
        If KDS_Activo = "SI" Then
            Call KDS_File(Trim(Str(nNUMTrans)))
        End If
        '~~~~~~~~~~~~~~~~~~~~~~
        
    End If
    '~|~|~|~|~|~|~|~|~|~|~|~|~|
    
    'Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, "TRANS# " & cCaja & " - " & nNUMTrans & Chr(&HD) & Chr(&HA)
    
    'Sys_Pos.Coptr1.PrintNormal PtrSReceipt, Chr(&H1B) & "!" & Chr(32) & Space(3) & "NOTA DE CREDITO" & vbCrLf
    'Sys_Pos.Coptr1.PrintNormal PtrSReceipt, Chr(&H1B) & Chr(&H46) & Space(6) & "FECHA/HORA: " & cFechaFactura & " - " & cHoraFactura & vbCrLf
    
    
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Chr(&H1B) & "!" & Chr(32) & "FACTURA # " & cCaja & " - " & nNUMTrans & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Chr(&H1B) & Chr(&H46) & "Cajero: " & cNomCaj & Chr(&HD) & Chr(&HA)
    Rem SE INVIERTE SUCURSAL y CAJA Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, "Caja: " & o_X_FF_Branch & " - Sucursal: " & o_poscode & Chr(&HD) & Chr(&HA)
    Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, "Caja: " & o_poscode & " - Sucursal: " & o_X_FF_Branch & Chr(&HD) & Chr(&HA)
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    
    If nImp = 0 Then

        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        Print2_OPOS_Dev Space(3)
        Rem txtString = "Consulte en https://dgi.mef.gob.pa/consulta usando el CUFE"
        txtString = "https://dgi-fep.mef.gob.pa/Consultas/FacturasPorCUFE/ usando el CUFE o Escanee el codigo QR     impreso en el comprobante"
        Print2_OPOS_Dev Space(nEspacio) & txtString
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        
        Print2_OPOS_Dev Space(3)
        
        Print2_OPOS_Dev Space(nEspacio) & "CUFE: " & FE_CUFE
        Print2_OPOS_Dev Space(3)
    End If
    
    If nImp = 0 Then
        If FE_QRData <> "" Then
            With Sys_Pos.Coptr1
                    .PrintBarCode PtrSReceipt, FE_QRData, PTR_BCS_QRCODE, .RecLineHeight * 0.6, .RecLineWidth * 0.6, PtrBcCenter, PtrBcTextNone&
            End With
        Else
            'INFO: UPDATE 21JUL2024
            If FE_Facil_ONLINE Then
                Print2_OPOS_Dev "<< FALLO AL IMPRIMIR QR >>"
                Print2_OPOS_Dev "<< FALLO AL IMPRIMIR QR >>"
                Print2_OPOS_Dev "VALOR QR: " & FE_QRData
                Print2_OPOS_Dev "<< FALLO AL IMPRIMIR QR >>"
                Print2_OPOS_Dev "<< FALLO AL IMPRIMIR QR >>"
            Else
                Print2_OPOS_Dev "<< FACTURA OFF LINE >>"
                Print2_OPOS_Dev "<< FACTURA OFF LINE >>"
                Print2_OPOS_Dev "VALOR QR: " & FE_QRData
                Print2_OPOS_Dev "<< FACTURA OFF LINE >>"
                Print2_OPOS_Dev "<< FACTURA OFF LINE >>"
            End If
            'Call Java_Factura
        End If
    End If
    
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    
    
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    'INFO: 28ABR2024
    If chkInfo.value = 1 Then
        
        If nImp = 0 Then
            'INFO:DATOS DEL CLIENTE
            
            Print2_OPOS_Dev Space(2)
            Print2_OPOS_Dev rs00!Mensaje
            Print2_OPOS_Dev "------------------------------"
            
            Rem txtString = "Protocolo de autorizacion: 0000155723374-2-202220240000000031165018, de 31/12/2024 09:29:37"
            If FE_Facil_ONLINE Then
                Print2_OPOS_Dev Space(3)
                txtString = "Protocolo de autorizacion: " & FE_authorization_number & " de " & FE_process_date
                Print2_OPOS_Dev Space(nEspacio) & txtString
            Else
                Print2_OPOS_Dev Space(3)
                txtString = o_msg
                Print2_OPOS_Dev Space(nEspacio) & txtString
            End If
            
            If FE_Facil_ONLINE Then
                Print2_OPOS_Dev Space(3)
                Print2_OPOS_Dev Space(nEspacio) & "RECEPTOR: " & GetTipoClienteDescrip(tipoClienteFE)
                Print2_OPOS_Dev Space(nEspacio) & "RUC: " & numeroRUC & " DV " & DVRUC
                Print2_OPOS_Dev Space(nEspacio) & "NOMBRE: " & razonSocial
                Print2_OPOS_Dev Space(nEspacio) & "EMAIL: " & LCase(correoElectronico1)
            Else
                Print2_OPOS_Dev Space(3)
                Print2_OPOS_Dev Space(nEspacio) & "RECEPTOR: Consumidor Final"
                Print2_OPOS_Dev Space(nEspacio) & "RUC: " & cRUCCliente
                Print2_OPOS_Dev Space(nEspacio) & "NOMBRE: " & cNombreCliente
                Rem Print2_OPOS_Dev Space(nEspacio) & "EMAIL: " & correoElectronico1
            End If
            
            Print2_OPOS_Dev Space(3)
            txtString = "Documento validado por Factura Facil S.A con RUC 155723374-2-2022, es Proveedor Autorizado Calificado, Resolucion No. 201-2167 de 14/03/2023."
            Print2_OPOS_Dev Space(nEspacio) & txtString
        End If
    Else
        If nImp = 0 Then
        
            Print2_OPOS_Dev Space(2)
            Print2_OPOS_Dev rs00!Mensaje
            Print2_OPOS_Dev "------------------------------"
            
            'INFO: CUMPLIR CON LA LEY
            If FE_Facil_ONLINE Then
                Print2_OPOS_Dev Space(3)
                txtString = "Protocolo de autorizacion: " & FE_authorization_number & " de " & FE_process_date
                Print2_OPOS_Dev Space(nEspacio) & txtString
            Else
                Print2_OPOS_Dev Space(3)
                txtString = o_msg
                Print2_OPOS_Dev Space(nEspacio) & txtString
            End If
            
            Print2_OPOS_Dev Space(3)
            txtString = "DOCUMENTO FACTURA "
            Print2_OPOS_Dev Space(nEspacio) & txtString
            txtString = "RECEPTOR: CONSUMIDOR FINAL"
            Print2_OPOS_Dev Space(nEspacio) & txtString
            
            'INFO: CUMPLIR CON LA LEY
            Print2_OPOS_Dev Space(3)
            txtString = "Documento validado por Factura Facil S.A con RUC 155723374-2-2022, es Proveedor Autorizado Calificado, Resolucion No. 201-2167 de 14/03/2023."
            Print2_OPOS_Dev Space(nEspacio) & txtString
            
        End If
        
    End If
    
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    
    If nImp = 0 Then
        Rem Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)
        Rem Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, rs00!Mensaje & Chr(&HD) & Chr(&HA)
        
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        'INFO: 22OCT2017.
        
        If nTransDomicilio Then
            
            Print2_OPOS_Dev String(30, "~")
            Print2_OPOS_Dev String(30, "~")
            
            aDomiInfo = GetDomicilioInfo()
            
            'INFO: 22NOV2017. Print2_OPOS_Dev "PEDIDO : " & nMesa & "     " & aDomiInfo(14, 0)
            Print2_OPOS_Dev "PEDIDO # " & cCaja & " - " & Right(Trim(Str(nNUMTrans)), 2)
            Print2_OPOS_Dev "Nombre  : " & aDomiInfo(3, 0) & Space(1) & aDomiInfo(4, 0)
            If aDomiInfo(6, 0) <> "" Then Print2_OPOS_Dev "Empresa : " & aDomiInfo(6, 0)
            Print2_OPOS_Dev "Telefono: " & FormatPhone(aDomiInfo(1, 0)) & IIf(aDomiInfo(2, 0) <> "", " (" & aDomiInfo(2, 0) & ")", Space(1))
            Print2_OPOS_Dev "ZONA: " & aDomiInfo(12, 0)
            
            Print2_OPOS_Dev "Dir: " & Left(aDomiInfo(8, 0), 25)
            If Mid(aDomiInfo(8, 0), 26, 25) <> "" Then Print2_OPOS_Dev "- " & Mid(aDomiInfo(8, 0), 26, 25)
            
            If aDomiInfo(9, 0) <> "" Then Print2_OPOS_Dev "- " & Left(aDomiInfo(9, 0), 25)
            If Mid(aDomiInfo(9, 0), 26, 25) <> "" Then Print2_OPOS_Dev "- " & Mid(aDomiInfo(9, 0), 26, 25)
            
            If aDomiInfo(10, 0) <> "" Then Print2_OPOS_Dev "- " & Left(aDomiInfo(10, 0), 25)
            If Mid(aDomiInfo(10, 0), 26, 25) <> "" Then Print2_OPOS_Dev "- " & Mid(aDomiInfo(10, 0), 26, 25)
            
            If Left(aDomiInfo(11, 0), 25) <> "" Then Print2_OPOS_Dev "INFO:" & Left(aDomiInfo(11, 0), 25)
            If Mid(aDomiInfo(11, 0), 26, 25) <> "" Then Print2_OPOS_Dev "- " & Mid(aDomiInfo(11, 0), 26, 25)
            'INFO: 23ENE2011
            If IsNull(aDomiInfo(15, 0)) Then
                'DO NOTHING, NO SE HA ASIGNADO MOTORIZADO
            Else
                Print2_OPOS_Dev "Motorizado: " & GetMotorizado(CLng(aDomiInfo(15, 0)))
            End If
            Print2_OPOS_Dev String(30, "~")
        End If
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        
        For i = 0 To 5
            'Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)
            Print2_OPOS_Dev Space(2)
        Next
        
        Sleep MUNBYNSeconds
        Sys_Pos.Coptr1.CutPaper 100
        ''Sys_Pos.Cocash1.ClaimDevice 5000
        'Sys_Pos.CoCash1.OpenDrawer     'INFO: 14FEB2017. USANDO DIRECTIO QUE NO ESPERA RESPUESTA DE LA GAVETA
        'Call Sys_Pos.Cocash1.DirectIO(DRW_DI_OPEN_DRAWER, 0, "")
        Call OpenGavetaDinero
        
    Else
        'Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, "------------------------------" & Chr(&HD) & Chr(&HA)
        Print2_OPOS_Dev "------------------------------"
    End If

    rsCuenta.Close
Next
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev Space(2)
'Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)
'Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Space(2) & Chr(&HD) & Chr(&HA)


nFlagParciales = 0
End Sub
Private Sub SetupPantalla()
    With ListaPagos
        .ColWidth(0) = 0: .ColWidth(1) = 2000: .ColWidth(2) = 800:
        .ColWidth(3) = 0
    End With
End Sub
Private Sub cdmBill_Click(Index As Integer)
nMntOculto = cdmBill(Index).Tag
lbMonto = Format(Val(nMntOculto), "STANDARD")
cmdFPagos_Click (0)
End Sub

Private Sub Clear_Click()
nfPase = 0
lbMonto = Format(0#, "STANDARD")
nMntOculto = ""
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdCARGOS_Click
' Author    : hsequeira
' Date      : 20/03/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdCARGOS_Click(Index As Integer)

On Error GoTo ErrAdm:

'If lbMonto < 0.01 Then Exit Sub

If PLU.PlatosMesa.Rows < 1 Then
    ShowMsg "NO HAY NADA MARCADO, FAVOR MARQUE PRODUCTOS", vbBlue, vbYellow
    cmdSalir_Click
    Exit Sub
End If

If ShowMsg("Desea realizar Cargo de" & vbCrLf & cmdCARGOS(Index).Caption & vbCrLf & vbCrLf & " por " & cmdCARGOS(Index).ToolTipText & "?", vbYellow, vbBlue, vbYesNo) = vbYes Then
    
    For i = 0 To cmdCARGOS.Count - 1
        cmdCARGOS(i).Enabled = False
    Next
    
    SBTot = SBTot + CSng(cmdCARGOS(Index).ToolTipText)
    'INFO: SI EL nTaxToRemove YA SE APLICO AQUI, ENTONCES NO SE VUELVE A APLICAR
    nTaxToRemove = 0
    
    Label2 = Format(SBTot, "CURRENCY")
    Label2.BackColor = vbYellow
    
    'ListaPagos.AddItem "20" & Chr(9) & cmdCARGOS(Index).Caption & Chr(9) &
    ListaPagos.AddItem cmdCARGOS(Index).Tag & Chr(9) & cmdCARGOS(Index).Caption & Chr(9) & Format(cmdCARGOS(Index).ToolTipText, "STANDARD") & Chr(9) & Format(cmdCARGOS(Index).ToolTipText, "STANDARD")
    
    'Call AddPago("PROPINA", cmdCARGOS(Index).Tag, PROPINA_DESCRIP & " " & cmdCARGOS(Index).Caption, cmdCARGOS(Index).ToolTipText, cmdCARGOS(Index).ToolTipText)
    'Call AddPago("PROPINA", "20", PROPINA_DESCRIP & " " & cmdCARGOS(Index).Caption,
    Call AddPago("PROPINA", cmdCARGOS(Index).Tag, PROPINA_DESCRIP & " " & cmdCARGOS(Index).Caption, cmdCARGOS(Index).ToolTipText, cmdCARGOS(Index).ToolTipText)
    ShowMsg "CARGO DE " & vbCrLf & cmdCARGOS(Index).Caption & vbCrLf & vbCrLf & " REALIZADO CON EXITO", vbBlue, vbYellow
Else
End If

Exit Sub

ErrAdm:
ShowMsg "ERROR EN PANTALLA DE CARGOS" & vbCrLf & Err.Number & " - " & Err.description, vbYellow, vbRed
'MsgBox Err.Description, vbCritical, Me.Name
'Resume
Exit Sub
End Sub

Private Sub cmdDescGlob_Click()
Dim nMiDesc As Integer
Dim nDescAplicado As Single

nDescAplicado = Format(lbMonto, "STANDARD")

If nDescAplicado < 0.01 Then
    MsgBox "NO PUEDE APLICAR ESE DESCUENTO", vbExclamation, BoxTit
    OKGlobal = 0
    Exit Sub
ElseIf OrigSB <> SBTot Then
    MsgBox "ES IMPOSIBLE APLICAR DESCUENTO GLOBAL", vbExclamation, BoxTit
    OKGlobal = 0
    Exit Sub
ElseIf nDescAplicado > SBTot Then
    MsgBox "ES IMPOSIBLE APLICAR ESTE DESCUENTO GLOBAL", vbExclamation, BoxTit
    OKGlobal = 0
    Exit Sub
End If
OKGlobal = 1
If OKGlobal = 1 Then
    OKGlobal = 0
    BoxPreg = "¿ DESEA APLICAR DESCUENTO DE " & Format(nDescAplicado, "CURRENCY") & "  ?"
    BoxResp = MsgBox(BoxPreg, vbQuestion + vbYesNo, BoxTit)
    If BoxResp = vbYes Then
        ListaPagos.AddItem 99 & Chr(9) & "DESC.GLOBAL" & Chr(9) & Format(nDescAplicado, "STANDARD") & Chr(9) & Format(nDescAplicado, "STANDARD")
        SBTot = SBTot - nDescAplicado
        Label3.BackColor = &HFFC0FF
        Label2.BackColor = &HFFC0FF
        Label3 = Format(SBTot, "CURRENCY")
        Label2 = Format(SBTot, "CURRENCY")
        'OrigSB = 1
    End If
    nfPase = 0
    lbMonto = Format(0#, "STANDARD")
    nMntOculto = ""
Else
    MsgBox "USTED NO ESTA AUTORIZADO PARA HACER DESCUENTOS", vbInformation, BoxTit
End If
End Sub
Private Sub cmdFPagos_Click(Index As Integer)
Rem BOTON DE FORMAS DE PAGO

Dim iFPagoCounter As Integer

On Error GoTo ErrAdm:
If lbMonto < 0.01 Then Exit Sub

'INFO: 13ENE2026
If lbMonto > 10000 Then
    ShowMsg "CANTIDAD A PAGAR ES INCORRECTA" & vbCrLf & vbCrLf & " $ " & lbMonto, vbYellow, vbRed
    Call Clear_Click
End If

If PLU.PlatosMesa.Rows < 1 Then
    ShowMsg "NO HAY NADA MARCADO, FAVOR MARQUE PRODUCTOS", vbBlue, vbYellow
    cmdSalir_Click
    Exit Sub
End If

RSPAGOS.MoveFirst
RSPAGOS.Find "CODIGO = " & cmdFPagos(Index).Tag
If Not RSPAGOS.EOF Then
    If RSPAGOS!CLIENTES = True Then
        MisClientes.Show 1
        If nCliNum = 0 Then
            'NO HACE NADA YA QUE NO SE MARCO UN CLIENTE
            Exit Sub
        Else
            MSHFClientes.AddItem cmdFPagos(Index).Tag & Chr(9) & nCliNum & Chr(9) & lbMonto
        End If
    End If
End If
On Error GoTo 0

'''''''''If SBTot < 0# Then
'''''''''    Label2 = Format(0#, "CURRENCY")
'''''''''    nCambio = SBTot * (-1)
'''''''''    SBTot = 0#
'''''''''Else
'''''''''    Label2 = Format(SBTot, "CURRENCY")
'''''''''End If
    
'INFO: UPDATE 07OCT2011
nfPase = 0
'INFO: 02NOV2011. CAMBIANDO DE SINGLE A CURRENCY
'SBTot = SBTot - CSng(lbMonto)
SBTot = SBTot - CCur(lbMonto)

If SBTot < 0 Then
    RSPROPINAS.MoveFirst
    RSPROPINAS.Find "CODIGO = " & cmdFPagos(Index).Tag
    If Not RSPROPINAS.EOF Then
        ShowMsg "NO puede cargar mas del SALDO DE LA CUENTA a esta tarjeta", vbRed, vbYellow
        'INFO: 16NOV2011. CAMBIANDO DE SINGLE A CURRENCY
        'SBTot = SBTot + CSng(lbMonto)
        SBTot = SBTot + CCur(lbMonto)
        
        '''ListaPagos.RemoveItem (ListaPagos.Rows)
        Exit Sub
    End If
    Label2 = Format(0#, "CURRENCY")
    nCambio = Format(SBTot * (-1), "#.00")
    ''''Debug.Print "1 - " & nCambio
    SBTot = 0#
Else
    Label2 = Format(SBTot, "CURRENCY")
End If
    
ListaPagos.AddItem cmdFPagos(Index).Tag & Chr(9) & cmdFPagos(Index).Caption & Chr(9) & Format(lbMonto - nCambio, "STANDARD") & Chr(9) & Format(lbMonto, "STANDARD")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'INFO: 07OCT2011
Call AddPago("PAGO", cmdFPagos(Index).Tag, cmdFPagos(Index).Caption, (lbMonto - nCambio), lbMonto)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

nMntOculto = ""
lbMonto = Format(0#, "STANDARD")

If SBTot = 0# Then
    'ImpresionFactura
    
    On Error Resume Next
    For iFPagoCounter = 0 To cmdFPagos.UBound
        cmdFPagos(iFPagoCounter).Enabled = False
    Next
    On Error GoTo 0
    cdmBill(0).Enabled = False
    cdmBill(1).Enabled = False
    cdmBill(2).Enabled = False
    cdmBill(3).Enabled = False
    cmdSalir.Enabled = False
    Clear.Enabled = False
    chkInfo.Enabled = False

    '~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
    '~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
    '07OCT2011
    If cFullFiscal = "SI" Then
        '~|~|~|~|~|~|~|~|~|~|~|~|~|
        'INFO: 26OCT2011
        Call GetNewNumTrans
        
        '~~~~~~~~~~~~~~~~~~~~~~
        'INFO: 6JUL2014
        If KDS_Activo = "SI" Then
            Call KDS_File(Trim(Str(nNUMTrans)))
        End If
        '~~~~~~~~~~~~~~~~~~~~~~
    
        '~|~|~|~|~|~|~|~|~|~|~|~|~|
        'Call FISCAL_FAST_ImprFactura(rsPagosMarcados, False, "EFECTIVO", 0#, aInfoCliente)
        'INFO: CARGOS. 20MAR2019
        Call FISCAL_FAST_ImprFactura(rsPagosMarcados, False, "EFECTIVO", 0#, aInfoCliente, rsPropinaMarcadas)
        
        'INFO: 24NOV2011
        Select Case nMarca
            Case 0      'STAR
                'INFO: 24NOV2011
                Call PutISC(FISCAL_.GetCurrentITBMS)
            Case 1      'BEMATECH
                'INFO: UPDATE 01DIC2011
                'Call PutISC(Format(iISCTransaccion, "STANDARD"))
                Call PutISC(FISCAL_.GetCurrentITBMS)
            Case Else
        End Select
        
    Else
        
        Call GetNewNumTrans
        
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        'NO HAY FISCAL. SE VA A SISTEMA VIEJO --> FACTURA ELECTRONICA
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        
        If FE_Facil_ONLINE Then
            'ProgBar.value = 30
            If GenerateJson(mi_Ambiente, nNUMTrans, rsPagosMarcados, ListaPagos, False, Format(Str(nMontoRecibido + nTaxRecibido), "#0.00")) Then
'                ProgBar.value = 85
'                ProgBar.value = 100
            Else
                If AllowFE_to_Continue Then
'                    ProgBar.value = 100
                Else
                    'Call CierreObjetos
                    ShowMsg "INTENTE HACER EL COBRO DE NUEVO", vbBlue, vbCyan
                    Call Check_FE_Website("SILENCIO")
                    cmdSalir.Enabled = True
                    Exit Sub
                End If
            End If
        Else
        
            '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
            '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
            If GenerateJson_OFFLINE(mi_Ambiente, nNUMTrans, rsPagosMarcados, ListaPagos, False, Format(Str(SubTot), "#0.00")) Then
            
            End If
            '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
            '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

            Rem Call OFF_Line_Dummy(Label3.Caption, nNUMTrans)
            Call ImprFactura
        End If
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        'INFO: UPDATE 21JUL2024
        If FE_Facil_ONLINE Then
            Call ImprFactura
        End If
        'INFO: UPDATE 21JUL2024
    End If
    '~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
    '~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
    
    'Call ImprFactura
    'EscribeLog ("Entrada a Actualizacion")
    
    '===---===---===---===---===---===---===
    If IMPRIME_TICKET = True Then Call ImprimeTicket(nNUMTrans)
    
    'If IMPRIME_PEDIDO = True Then Call ImprimeEnCocina(rs00!TRANS + 1)
    If IMPRIME_PEDIDO = True Then
        'INFO: 21DIC2014
        'INFO: ACTUALIZANDO PARA HACER UNA SOLA APERTURA A MESASPED, AHORRA 50% DE TIEMPO
        'Y HACER QUE EL PEDIDO SE EJECUTE EN UN SOLO COMMIT A MESASPED.
        MesasPED "OPEN"
        Call SendToPrintServer(nNUMTrans)
        MesasPED "CLOSE"
    End If
    
    If DUPLICA_PEDIDO = True Then Call DuplicaPedido(nNUMTrans)
    '===---===---===---===---===---===---===
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    
    Call Actualizador(nMontoRecibido, nTaxRecibido, iISCTransaccion)
    
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    
    If nCambio <> 0# Then
        Vuelto.Show 1
    End If
    nCambio = 0#

    'INFO: 20OCT2011
    cNombreCliente = ""
    
    On Error Resume Next
    For iFPagoCounter = 0 To cmdFPagos.UBound
        cmdFPagos(iFPagoCounter).Enabled = True
    Next
    On Error GoTo 0
    
    cdmBill(0).Enabled = True
    cdmBill(1).Enabled = True
    cdmBill(2).Enabled = True
    cdmBill(3).Enabled = True
    cmdSalir.Enabled = True
    Clear.Enabled = True
    chkInfo.Enabled = True
    
    Unload Me
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'INFO: 07OCT2011
    Set rsPropinaMarcadas = Nothing
    Set rsPagosMarcados = Nothing
    
    Sys_Pos.Timer1.Enabled = True
    
    Call MuestraProductosMarcados
    'INFO: VALIDANDO SI ES FISCAL, SE DESACTIVA EL 15JUN2012
    '~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||
    'INFO: 06OCT2011
    If cFullFiscal = "SI" Then
    Else
        Call PLU.ImprimeTitulos
    End If
    '~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||
    
    PLU.axPagado.Text = 0#
    Set Pagos = Nothing
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End If

Exit Sub

ErrAdm:
MsgBox Err.Number & " - " & Err.description, vbCritical, "ERROR EN PANTALLA DE PAGOS"
'MsgBox Err.Description, vbCritical, Me.Name
Exit Sub
End Sub
Private Sub cmdSalir_Click()
RSPROPINAS.Close
Set RSPROPINAS = Nothing

'INFO: RESTABLECIAENDO NCANTIDAD 13ENE2026
nCantidad = 1

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'INFO: 07OCT2011
If OrigSB = SBTot Then
    'NO HACER NADA
Else
   SBTot = OrigSB
End If
Sys_Pos.Timer1.Enabled = True

Set rsPropinaMarcadas = Nothing
Set rsPagosMarcados = Nothing

Unload Me
Set Pagos = Nothing
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End Sub


Private Sub Command2_Click(Index As Integer)
Dim cChar As String
Dim cCant As String

'INFO: UPDATE 13RNR2026
If nfPase = 0 Then
    nMntOculto = Command2(Index).Caption
Else
    nMntOculto = nMntOculto & Command2(Index).Caption
    cCant = Str(nCantidad)
    cCant = cCant & Command2(Index).Index
    Rem nCantidad = Val(cCant)
    'INFO: 15FEB2011. BAJANDO LA CANTIDAD DE 8 a 4
    'If Len(cCant) = 6 Then
    'INFO: 3OCT2017
    If Len(cCant) = 8 Then
        nCantidad = 1
        nfPase = 0
        lbMonto = Format(0#, "STANDARD")
        nMntOculto = ""
        
        ShowMsg "CANTIDAD/MONTO NO ES VALIDO", vbYellow, vbRed
        'Text1(2) = nCantidad / 100
        Exit Sub
    End If
End If

lbMonto = Format(Val(nMntOculto) / 100, "STANDARD")
nfPase = nfPase + 1


'INFO: 13ENE2026
If lbMonto > 10000 Then
    ShowMsg "CANTIDAD A PAGAR ES INCORRECTA" & vbCrLf & vbCrLf & " $ " & lbMonto, vbYellow, vbRed
    nCantidad = 1
    Call Clear_Click
End If

End Sub

Private Sub Form_Load()
nfPase = 0
OKGlobal = 0

'nMontoRecibido As Single
'nTaxRecibido As Single

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'INFO: 07OCT2011
Sys_Pos.Timer1.Enabled = False
aInfoCliente(0) = ""
aInfoCliente(1) = ""

'INFO: 18SEP2012
RegWrite "HKCU\Software\SoloSoftware\SoloMix\LastCliente", ""
RegWrite "HKCU\Software\SoloSoftware\SoloMix\LastRUC", ""

'INFO: RECORDSET PARA PAGOS y PROPINAS QUE SE ESTAN MARCANDO
Call PrepareRecordSets
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Label3 = Format(SBTot, "CURRENCY")
Label2 = Format(SBTot, "CURRENCY")
OrigSB = Format(SBTot, "CURRENCY")

nMntOculto = ""

CargaFormasPagoSimple RSPAGOS, Pagos
CargaCargos RSPROPINAS, Pagos
SetupPantalla

Dim rsParciales As Recordset
Dim rsParcPro As Recordset
Dim rsCliParc As Recordset
Dim lParc As Integer

Set rsParciales = New Recordset
rsParciales.Open "SELECT CAJERO,MESA,MESERO,TIPO_PAGO,LIN,MONTO FROM TMP_PAR_PAGO ", msConn, adOpenDynamic, adLockOptimistic

Set rsParcPro = New Recordset
rsParcPro.Open "SELECT CAJERO,MESA,MESERO,TIPO_PAGO,LIN,MONTO FROM TMP_PAR_PROP ", msConn, adOpenDynamic, adLockOptimistic

Set rsCliParc = New Recordset
rsCliParc.Open "SELECT * FROM TMP_CLI ", msConn, adOpenDynamic, adLockOptimistic

nFlagParciales = 0

Do Until rsParciales.EOF
    RSPAGOS.MoveFirst
    RSPAGOS.Find "CODIGO = " & rsParciales!TIPO_PAGO
    If RSPAGOS.EOF Then
        'SI ES FIN DE ARCHIVO LO UNICO QUE PUEDE SER ES UN DESCUENTO <<GLOBAL>>
        ListaPagos.AddItem rsParciales!TIPO_PAGO & Chr(9) & "DESC.GLOBAL" & Chr(9) & Format(rsParciales!MONTO, "STANDARD") & Chr(9) & Format(rsParciales!MONTO, "STANDARD")
    Else
        ListaPagos.AddItem rsParciales!TIPO_PAGO & Chr(9) & RSPAGOS!DESCRIP & Chr(9) & Format(rsParciales!MONTO, "STANDARD") & Chr(9) & Format(rsParciales!MONTO, "STANDARD")
    End If
    SBTot = SBTot - rsParciales!MONTO
    rsParciales.MoveNext
    nFlagParciales = 1
Loop

Label2 = Format(SBTot, "CURRENCY")

Do Until rsParcPro.EOF
    RSPAGOS.MoveFirst
    RSPAGOS.Find "CODIGO = " & rsParcPro!TIPO_PAGO
    ListaPagos.AddItem rsParcPro!TIPO_PAGO & Chr(9) & "PROPINA " & RSPAGOS!DESCRIP & Chr(9) & Format(rsParcPro!MONTO, "STANDARD") & Chr(9) & Format(rsParcPro!MONTO, "STANDARD")
    rsParcPro.MoveNext
    nFlagParciales = 1
Loop

Do Until rsCliParc.EOF
    MSHFClientes.AddItem rsCliParc!CODIGO_TP & Chr(9) & rsCliParc!CODIGO_CLI & Chr(9) & rsCliParc!MONTO
    rsCliParc.MoveNext
Loop

'INFO: 07OCT2011. LO HAGO AQUI x CUESTION DE TIEMPO, PASARLO A CargaFormasPagoSimple
Set RSPROPINAS = New ADODB.Recordset
sqltext = "SELECT * FROM pagos WHERE TIPO = 'TJ' OR TIPO = 'CH' ORDER BY CODIGO"
RSPROPINAS.Open sqltext, msConn, adOpenStatic, adLockOptimistic

End Sub
Private Sub Label2_Click()
    nMntOculto = Label2.Caption
    lbMonto = Format(nMntOculto, "STANDARD")
End Sub


''---------------------------------------------------------------------------------------
'' Procedure : PutDomicilioInfo
'' Author    : hsequeira
'' Date      : 18/01/2017
'' Purpose   : CONTABILIZA LAS VENTAS DE DOMICILIO
''---------------------------------------------------------------------------------------
''
'Private Function PutDomicilioInfo(dNUM_TRANS As Long, dNVALOR_PAGO As Single) As Boolean
'Dim cSQL As String
'Dim rsDOMI As ADODB.Recordset
'
'Set rsDOMI = New ADODB.Recordset
'cSQL = "SELECT SUM(CANT) AS dCANT, SUM(PRECIO) AS dVENTAS FROM TMP_TRANS WHERE DEPTO = " & nHayDomicilio
'rsDOMI.Open cSQL, msConn, adOpenStatic, adLockReadOnly
'
'If rsDOMI.EOF Then
'    rsDOMI.Close
'    Set rsDOMI = Nothing
'    PutDomicilioInfo = True
'    Exit Function
'End If
'
'If IsNull(rsDOMI!dCANT) Then
'    rsDOMI.Close
'    Set rsDOMI = Nothing
'    PutDomicilioInfo = True
'    Exit Function
'End If
'
'cSQL = "INSERT INTO DOMICILIO (FECHA, HORA, REPORTEZ, CANT, VALOR_DOMICILIO, VALOR_FACTURA, DOC_SOLO) "
'cSQL = cSQL & " VALUES ('"
'cSQL = cSQL & Format(Date, "YYYYMMDD") & "','" & Format(Time, "HH:MM") & "', 'N', " & rsDOMI!dCANT & "," & rsDOMI!dVENTAS & ","
'cSQL = cSQL & dNVALOR_PAGO & "," & dNUM_TRANS & ")"
'
'msConn.Execute cSQL
'
'rsDOMI.Close
'Set rsDOMI = Nothing
'PutDomicilioInfo = True
'Exit Function
'
'End Function

Public Function MuestraPanta_Pagos(fSubtotal As Single, fTax As Single)
nMontoRecibido = fSubtotal
nTaxRecibido = fTax
mi_Panta = "PAGOS"
Me.Show 1
mi_Panta = ""
End Function
