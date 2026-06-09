VERSION 5.00
Begin VB.Form ActHost 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DEPARTAMENTOS ABIERTOS"
   ClientHeight    =   4935
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
   ScaleHeight     =   4935
   ScaleWidth      =   5355
   Begin VB.CommandButton Command1 
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
      Picture         =   "ActHost.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3960
      Width           =   735
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
      Picture         =   "ActHost.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1920
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
      TabIndex        =   17
      Top             =   1320
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
      TabIndex        =   16
      Top             =   2280
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
      TabIndex        =   15
      Top             =   3240
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
      Picture         =   "ActHost.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4080
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
      TabIndex        =   12
      Top             =   4080
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
      TabIndex        =   11
      Top             =   4080
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
      TabIndex        =   10
      Top             =   3360
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
      TabIndex        =   9
      Top             =   3360
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
      TabIndex        =   8
      Top             =   3360
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
      TabIndex        =   7
      Top             =   2640
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
      TabIndex        =   6
      Top             =   2640
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
      TabIndex        =   5
      Top             =   2640
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
      TabIndex        =   4
      Top             =   1920
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
      TabIndex        =   3
      Top             =   1920
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
      TabIndex        =   2
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancelar"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   1575
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
      TabIndex        =   19
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   2895
      Left            =   120
      Top             =   1200
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
      TabIndex        =   14
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   3615
      Left            =   2040
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Anote el Precio de lo que desea cobrar y luego Seleccione un Departamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   2160
      TabIndex        =   1
      Top             =   0
      Width           =   2895
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
'If Val(nMontoOpen) <= 0 Then
If Val(nMontoOpen) = 0 Then
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

Private Sub Command1_Click()
txtInfo = "Escriba Clave para Aplicar"
OkAnul = 0
AskClave.Show 1
If OkAnul = 1 Then
    lbMonto = Format((lbMonto * -1), "CURRENCY")
    nMontoOpen = (nMontoOpen * -1)
End If
OkAnul = 0
End Sub

Private Sub Command2_Click(Index As Integer)
Dim cChar As String
If nfPase = 0 Then
    nMontoOpen = Command2(Index).Caption
Else
    nMontoOpen = nMontoOpen & Command2(Index).Caption
End If
lbMonto = Format(Val(nMontoOpen) / 100, "CURRENCY")
nfPase = nfPase + 1
End Sub

Private Sub Form_Load()
nMontoOpen = ""
nfPase = 0
nMulti = 1
End Sub
