VERSION 5.00
Begin VB.Form ActHost 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DEPARTAMENTOS ABIERTOS"
   ClientHeight    =   6360
   ClientLeft      =   1980
   ClientTop       =   3735
   ClientWidth     =   5355
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   Icon            =   "ActHost.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOpenDept 
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
      Left            =   120
      MaxLength       =   20
      TabIndex        =   0
      ToolTipText     =   "Opcional: Escribir nombre corto de Producto"
      Top             =   1680
      Width           =   5055
   End
   Begin VB.CommandButton cmdPor 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      Picture         =   "ActHost.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdOpcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "RESTAURANTE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   635
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdOpcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "BARRA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   635
      Index           =   1
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdOpcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "OTRO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   635
      Index           =   2
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   1575
   End
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
      Left            =   4440
      Picture         =   "ActHost.frx":044E
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "00"
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
      Index           =   10
      Left            =   3720
      TabIndex        =   17
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "0"
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
      Index           =   0
      Left            =   3000
      TabIndex        =   16
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   9
      Left            =   4440
      TabIndex        =   15
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   3720
      TabIndex        =   14
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   3000
      TabIndex        =   13
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   4440
      TabIndex        =   12
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   3720
      TabIndex        =   11
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   3000
      TabIndex        =   10
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   4440
      TabIndex        =   9
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   3720
      TabIndex        =   8
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   3000
      TabIndex        =   7
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   635
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Opcional. Escriba Descripcion del Producto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   21
      Top             =   1320
      Width           =   5175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   4080
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      Height          =   2895
      Left            =   120
      Top             =   2400
      Width           =   1815
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
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   2160
      TabIndex        =   20
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      Height          =   3615
      Left            =   2040
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Anote el Precio de lo que desea cobrar y luego Seleccione un Departamento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Index           =   0
      Left            =   840
      TabIndex        =   19
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "ActHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private nfPase As Integer
Private nMontoOpen As Variant
Private Sub Clear_Click()
nfPase = 0
lbMonto = 0#
nMontoOpen = 0
nMulti = 1
Label2 = ""
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
'INFO: DESCRIPCION LIBRE EN DEPARTAMENTO ABIERTO
' SE UBICA ARRIBA DEL PROCEDIMIENTO PARA QUE GUARDE LA INFO
' QUE SE ESCRIBIO DEL DEPTO ABIERTO, AUNQUE TENGA VALOR CERO
'01OCT2012
'10DIC2012
Select Case Index
    Case 0  'COCINA/RESTAURANTE
        nDeptoAbiertoSeleccionado = 1
    Case 1  'BAR/BARRA
        nDeptoAbiertoSeleccionado = 2
    Case 2  'OTRO/PRECUENTA
        nDeptoAbiertoSeleccionado = 0
    Case Else
        nDeptoAbiertoSeleccionado = 0
End Select

If txtOpenDept.Text <> "" Then
    txtOpenDept.Text = UCase(txtOpenDept.Text)
    '10DIC2012. NO ESTABA VALIDANDO IMPRESION DE COCINA
    TXT_OPEN_DEPT = UCase(txtOpenDept.Text)
End If
'INFO: LO ESCRIBE EN EL REGISTRY SIEMPRE
RegWrite "HKCU\Software\SoloSoftware\SoloMix\DeptoAbiertoItem", txtOpenDept.Text

If Val(nMontoOpen) <= 0 Then
    cmdCancelar_Click
    Exit Sub
End If
ValOpenDept = Val(nMontoOpen * nMulti)

TXT_OPEN_DEPT = Space(1) + cmdOpcion(Index).Caption

cmdCancelar_Click

End Sub

Private Sub cmdPor_Click()
nMulti = lbMonto * 100
Label2 = nMulti
nfPase = 0
lbMonto = 0#
nMontoOpen = 0
End Sub

Private Sub Command2_Click(Index As Integer)
Dim cChar As String
If Len(nMontoOpen) > 5 Then
    'MsgBox "CANTIDAD ES INVALIDA", vbCritical, "Inserte una Cantidad Real"
    ShowMsg "MONTO ES INVALIDO" & vbCrLf & "Inserte un Monto Real"
    nfPase = 0
    nMontoOpen = 0
    lbMonto = Format(0#, "currency")
    Exit Sub
End If
If nfPase = 0 Then
    nMontoOpen = Command2(Index).Caption
Else
    nMontoOpen = nMontoOpen & Command2(Index).Caption
End If
lbMonto = Format(Val(nMontoOpen) / 100, "currency")
nfPase = nfPase + 1
End Sub

Private Sub Form_Load()
nMontoOpen = ""
nfPase = 0
nMulti = 1
End Sub

