VERSION 5.00
Object = "{C53C04A0-763E-11D2-9A7B-0080C847EAF4}#1.0#0"; "axPanel.dll"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form PLU 
   BackColor       =   &H00B39665&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MODULO DE VENTAS"
   ClientHeight    =   8790
   ClientLeft      =   105
   ClientTop       =   495
   ClientWidth     =   14130
   ForeColor       =   &H00000000&
   Icon            =   "PLU.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   14130
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid ListaPagos 
      Height          =   735
      Left            =   1800
      TabIndex        =   39
      Top             =   8040
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorSel    =   65535
      BackColorBkg    =   255
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   12120
      Top             =   6600
   End
   Begin VB.CheckBox ChkToGo 
      BackColor       =   &H000000FF&
      Caption         =   "PARA LLEVAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   9240
      TabIndex        =   36
      Top             =   5355
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Platos del Menú"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   7275
      Index           =   1
      Left            =   1800
      TabIndex        =   5
      Top             =   60
      Width           =   7250
      Begin VB.CommandButton cmdPlus 
         BackColor       =   &H00C0C0C0&
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton Correccion 
      Caption         =   "Correción"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10320
      Picture         =   "PLU.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7995
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
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
      Left            =   9600
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   7995
      Width           =   735
   End
   Begin VB.CommandButton Command8 
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
      Height          =   615
      Index           =   9
      Left            =   11040
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7395
      Width           =   735
   End
   Begin VB.CommandButton Command8 
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
      Height          =   615
      Index           =   8
      Left            =   10320
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   7395
      Width           =   735
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid PlatosMesa 
      Height          =   3255
      Left            =   9120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   60
      Width           =   4975
      _ExtentX        =   8784
      _ExtentY        =   5741
      _Version        =   393216
      BackColor       =   0
      ForeColor       =   65280
      Rows            =   1
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      GridColor       =   16777215
      AllowBigSelection=   0   'False
      ScrollBars      =   2
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.CommandButton cmdAcepPago 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10440
      Picture         =   "PLU.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "PAGO EN EFECTIVO"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmbBill 
      Height          =   735
      Index           =   3
      Left            =   10065
      Picture         =   "PLU.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   33
      Tag             =   "20.00"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmbBill 
      Height          =   735
      Index           =   2
      Left            =   9150
      Picture         =   "PLU.frx":1B56
      Style           =   1  'Graphical
      TabIndex        =   32
      Tag             =   "10.00"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmbBill 
      Height          =   735
      Index           =   1
      Left            =   10065
      Picture         =   "PLU.frx":2B5C
      Style           =   1  'Graphical
      TabIndex        =   31
      Tag             =   "5.00"
      Top             =   3500
      Width           =   855
   End
   Begin VB.CommandButton cmbBill 
      Height          =   735
      Index           =   0
      Left            =   9150
      Picture         =   "PLU.frx":3B9F
      Style           =   1  'Graphical
      TabIndex        =   30
      Tag             =   "1.00"
      Top             =   3500
      Width           =   855
   End
   Begin VB.CommandButton cmpPago 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Picture         =   "PLU.frx":4C12
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "PANTALLA DE PAGOS"
      Top             =   6550
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   2
      Left            =   7920
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton Clear 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   25
      Top             =   7640
      Width           =   1575
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Anulación"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   575
      Index           =   1
      Left            =   4155
      Picture         =   "PLU.frx":50B1
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8125
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Descuen&to / Plato"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   575
      Index           =   0
      Left            =   4155
      TabIndex        =   23
      Top             =   7450
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
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
      Height          =   615
      Index           =   7
      Left            =   9600
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   7395
      Width           =   735
   End
   Begin VB.CommandButton Command8 
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
      Height          =   615
      Index           =   6
      Left            =   11040
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6795
      Width           =   735
   End
   Begin VB.CommandButton Command8 
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
      Height          =   615
      Index           =   5
      Left            =   10320
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6795
      Width           =   735
   End
   Begin VB.CommandButton Command8 
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
      Height          =   615
      Index           =   4
      Left            =   9600
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6795
      Width           =   735
   End
   Begin VB.CommandButton Command8 
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
      Height          =   615
      Index           =   3
      Left            =   11040
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6195
      Width           =   735
   End
   Begin VB.CommandButton Command8 
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
      Height          =   615
      Index           =   2
      Left            =   10320
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6195
      Width           =   735
   End
   Begin VB.CommandButton Command8 
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
      Height          =   615
      Index           =   1
      Left            =   9600
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6195
      Width           =   735
   End
   Begin VB.CommandButton cmdPLUNext 
      BackColor       =   &H0000C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      Picture         =   "PLU.frx":54F3
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7275
      Width           =   1095
   End
   Begin VB.CommandButton cmdPLUTop 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      Picture         =   "PLU.frx":67F5
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7275
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00008080&
      Caption         =   "Departamentos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6375
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   60
      Width           =   1815
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   50
         Picture         =   "PLU.frx":7AF7
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5520
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   930
         Picture         =   "PLU.frx":8DF9
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5520
         Width           =   855
      End
      Begin VB.CommandButton cmdDepto 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H0000FFFF&
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7900
      Width           =   1215
   End
   Begin AXPANELLibCtl.axPanel axPagado 
      Height          =   495
      Left            =   -360
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   7200
      Width           =   2175
      DigitsTotal     =   8
      DigitsAfterComma=   2
      ShowLeadingZeros=   0
      Text            =   "0"
      Style           =   0
      Thickness       =   3
      Object.Height          =   23
      BackgroundColor =   4210752
      ActiveColor     =   65280
      UnActiveColor   =   0
      BorderStyle     =   3
      OffsetLeft      =   0
      OffsetTop       =   3
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFClientes 
      Height          =   255
      Left            =   6360
      TabIndex        =   40
      Top             =   7200
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
   Begin MSComctlLib.ProgressBar ProgBar 
      Height          =   5415
      Left            =   13440
      TabIndex        =   41
      Top             =   3360
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   9551
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Orientation     =   1
   End
   Begin VB.Label lbNOTax 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   540
      Left            =   11520
      TabIndex        =   38
      Top             =   3380
      Width           =   1545
   End
   Begin VB.Label lbTAX 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   540
      Left            =   11520
      TabIndex        =   37
      Top             =   3960
      Width           =   1530
   End
   Begin VB.Image Imagen 
      Height          =   885
      Index           =   2
      Left            =   11760
      Picture         =   "PLU.frx":A0FB
      Stretch         =   -1  'True
      Top             =   5140
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Imagen 
      Height          =   885
      Index           =   1
      Left            =   11760
      Picture         =   "PLU.frx":BA45
      Stretch         =   -1  'True
      Top             =   5140
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image 
      Height          =   600
      Left            =   5520
      Picture         =   "PLU.frx":CAA0
      Stretch         =   -1  'True
      ToolTipText     =   "Repetir Ultima Factura"
      Top             =   7440
      Width           =   645
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   1665
      Left            =   9105
      Top             =   3435
      Width           =   1905
   End
   Begin VB.Label Hora 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   300
      Left            =   6360
      TabIndex        =   28
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "CANT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   8280
      TabIndex        =   27
      Top             =   7680
      Width           =   855
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   1095
      Left            =   6300
      Top             =   7560
      Width           =   3120
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      Height          =   1380
      Left            =   4080
      Top             =   7395
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   2535
      Left            =   9480
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Label SubTot 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   540
      Left            =   11040
      TabIndex        =   11
      Top             =   4520
      Width           =   2025
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "SUB-TOT"
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
      Height          =   255
      Index           =   0
      Left            =   12120
      TabIndex        =   0
      Top             =   7080
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "PLU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private num As Integer
Private numplu As Integer
Private nNLinSel As Integer
Private Arreg_Deptos(10) As Long
Private Arreg_Plu(33) As Integer
Private nPase As Integer 'Cantidad de Clicks a Cantidad
Private ElDepto As Long 'Es el Departamento Seleccionado
Private nGlobEnv As Long    'El envase seleccionado
Private TextEnv As String
Private rsTmpAco As New ADODB.Recordset
Private nAcoBookMark As Variant
Private Const nNUM_MAX_DEPTOS = 8
Private Const nNUM_MAX_PLU = 27
Private nMaxPLu As Integer
Private nMontoPagado As Currency
'*-----------PARA IMPRESOR GENERICO ------------------
Private lReturn As Long
Private lpcWritten As Long
Private lDoc As Long
Private sWrittenData As String
Private MyDocInfo As DOCINFO

'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Private Const PBM_SETBARCOLOR As Long = &H409
Private Const PBM_SETBKCOLOR As Long = &H2001
Private Const PROGBAR_DEF_COLOR = &HFF000000 '&H8000000D
Private cFEInitTime As String       'INICIO DE EMISION DE FACTURA
Private nFEPropinaEfectivo As Single

Private rsPropinaMarcadas As ADODB.Recordset        'DICIEMBRE DE 2014
Private rsPagosMarcados As ADODB.Recordset          'DICIEMBRE DE 2014
Private aInfoCliente(1) As String                                       'SEPTIEMBRE 2010
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

'Private Declare Function GetTickCount Lib "kernel32" () As Long
'-----------------------------------------------------------
Public Sub ImprimeTitulos()
'INFO: DESACTIVADO EL 15JUN2012
'DoEvents  'INFO: REMOVED 13SEP2013
Rem ENCABEZADO DE LA FACTURA
Dim ESC As String * 1
Dim lineWidth As Integer
Dim txtString As String

ESC = Chr(&H1B)

'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'INFO: 22NOV2017. CONSULTORES ALAMAR
'INFO: 27JUN2024. ESTE ES EL COMANDO QUE REALMENTE IMPRIME EL TITULO GUARDADO EN MEMORIA
If bLogo Then Sys_Pos.Coptr1.PrintNormal PtrSReceipt, ESC + "|1B"
Print2_OPOS_Dev Space(2)
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "BEGIN")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

txtString = "DGI"
Print2_OPOS_Dev Chr(&H1B) & "!" & Chr(32) & Space(11) & txtString

txtString = rs00!DESCRIP
Print2_OPOS_Dev GetSpace(txtString) & txtString
txtString = rs00!RAZ_SOC
Print2_OPOS_Dev GetSpace(txtString) & txtString
txtString = "RUC:" & rs00!ruc
Print2_OPOS_Dev GetSpace(txtString) & txtString
txtString = Mid$(rs00!direccion, 1, 25)
Print2_OPOS_Dev GetSpace(txtString) & txtString
Print2_OPOS_Dev Space(2)

txtString = "COMPROBANTE AUXILIAR DE FACTURA ELECTRONICA"
Print2_OPOS_Dev Chr(&H1B) & Chr(&H46) & GetSpace(txtString) & txtString
'Print2_OPOS_Dev GetSpace(txtString) & txtString
'txtString = "COMPROBANTE AUXILIAR"
'Print2_OPOS_Dev GetSpace(txtString) & txtString
Rem txtString = "FACTURA ELECTRONICA"
Rem Print2_OPOS_Dev GetSpace(txtString) & txtString
txtString = "FACTURA DE OPERACION INTERNA"
Print2_OPOS_Dev GetSpace(txtString) & txtString
Print2_OPOS_Dev Space(2)
Print2_OPOS_Dev "------------------------------"

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "END")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

End Sub
Private Sub AddOpenDeptItem()
'Agrega registros a TMP_TRANS desde un Departamento Abierto
Dim SOLO_FECHA As String
Dim cSQL As String
Dim cStr2Print As String    'INFO: 13SEP2013

ValOpenDept = 0
TXT_OPEN_DEPT = ""

ActHost.Show 1
If TXT_OPEN_DEPT = "" Then Exit Sub

CajLin = CajLin + 1
SOLO_FECHA = Format(Date, "YYYYMMDD")
    
cSQL = "INSERT INTO TMP_TRANS "
'cSQL = cSQL & "(CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT,DEPTO,PLU,ENVASE,PRECIO_UNIT,PRECIO,FECHA,HORA,TIPO,DESCUENTO,CUENTA,IMPRESO,IMPRESORA) VALUES ("
'INFO: 25NOV2014. VISTO EN SUSHI EXPRESS LOS ANDES x FUFO.
'SE PONE EL IMPUESTO EN CAMPO CON_TAX, YA QUE POR DISEÑO EN LA BASE DE DATOS, EL IMPUESTO
'DEFAULT LO PONE COMO 5 PORCIENTO. SE AGREGA AL SOLOINI.INI PorcentajeImpuesto
'QUE TIENE EL VALOR DEL IMPUESTO DEFAULT
cSQL = cSQL & "(CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT,DEPTO,PLU,ENVASE,PRECIO_UNIT,PRECIO,FECHA,HORA,TIPO,DESCUENTO,CUENTA,IMPRESO,IMPRESORA, CON_TAX) VALUES ("
cSQL = cSQL & cCaja & "," & npNumCaj & "," & nMesa & "," & nMesero & "," & -1 & "," & CajLin & ",'"
cSQL = cSQL & rs02!corto & TXT_OPEN_DEPT & "'," & nMulti & "," & rs02!CODIGO & "," & 0 & "," & 0 & ","
cSQL = cSQL & Format((ValOpenDept / nMulti) / 100, "#0.00") & "," & Format(ValOpenDept / 100, "#0.00") & ",'"
'cSQL = cSQL & SOLO_FECHA & "','" & Time & "','  '," & 0# & "," & 0 & ",0," & BARRA_01 & ")"
'INFO: 28OCT2016. BUFFALO WINGS. HAY QUE ARRGLAR PROCEDIMIENTO.
'cSQL = cSQL & SOLO_FECHA & "','" & Time & "','  '," & 0# & "," & 0 & ",0," & BARRA_01 & ","
cSQL = cSQL & SOLO_FECHA & "','" & Time & "','  '," & 0# & "," & 0 & ",0," & 1 & ","

'INFO: 25NOV2014.
cSQL = cSQL & GetFromINI("Facturacion", "PorcentajeImpuesto", App.Path & "\soloini.ini") & ")"


msConn.BeginTrans
msConn.Execute cSQL
msConn.CommitTrans

'~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||
'INFO: 26CT2011
If cFullFiscal = "SI" Then
    'INFO: NO HACER NADA, SE ESTA CAMBIANDO PARA HACERLO AL FINAL. CUANDO SE COBRA
Else
    'IMPRESION DE OPEN DEPT.
    cStr2Print = Format(Mid(rs02!corto & TXT_OPEN_DEPT, 1, 18), "!@@@@@@@@@@@@@@@@@@") + Space(1) & Format(nMulti, "@@@") & Space(1) & Format(Format((ValOpenDept / nMulti) / 100, "#0.00"), "@@@@@@@")
    'Sys_Pos.Coptr1.PrintTwoNormal FptrSJournalReceipt, cStr2Print, cStr2Print
    Print2_OPOS_Dev cStr2Print
End If
'~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||

Call MuestraProductosMarcados

nLineas = PlatosMesa.Rows - 1
PlatosMesa.TopRow = (PlatosMesa.Rows - 1)
nCantidad = 1: nPase = 0
nNLinSel = 0
Text1(2) = nCantidad

End Sub

'---------------------------------------------------------------------------------------
' Procedure : DescProducto
' Author    : hsequeira
' Date      : 27/03/2013
' Purpose   : DAR DESCUENTO A UN PRODUCTO
' HASTA 27MAR2013 EL REDONDEO DEL DESCUENTO ERA PARA ABAJO
' DE ESA FORMA SE EVITABA EL DESCUENTO.
' CAMBIANDO PARA QUE DE EL DESCUENTO REAL
'---------------------------------------------------------------------------------------
'
Private Sub DescProducto()
Dim MiDesc As Single
Dim nDescImpre As Single
Dim sqltext As String
Dim cStr2Print As String
Dim cSQL As String, cSQL2 As String

If PlatosMesa.Rows = 0 Then
    ShowMsg " No hay nada Marcado ", vbBlue, vbYellow
    Exit Sub
End If

If nCantidad > MAX_DESCUENTO Then
    ShowMsg "ES IMPOSIBLE DAR ESE DESCUENTO. INTENTE DAR UN PORCENTAJE MAS BAJO", vbBlue, vbYellow
    Clear_Click
    Exit Sub
End If

'nCantidad es el valor del Cuadro de Numeros de la Derecha Abajo
If nCantidad > 1 Then
    MiDesc = Format(nCantidad / 100, "STANDARD")
Else
    'nDesc01 es el Descuento Marcado
    MiDesc = Format(nDesc01 / 100, "STANDARD")
End If

MiDesc = Format(MiDesc, "STANDARD")

Dim rsFixTmpTrans As Recordset
Dim txto As String

Set rsFixTmpTrans = New Recordset
If nNLinSel <> 0 Then
    'INFO: CAMBIO PARA MULTICAJA
    txto = "SELECT * FROM TMP_TRANS "
    txto = txto & " WHERE LIN = " & nNLinSel & " AND MESA = " & nMesa & " AND CAJERO = " & npNumCaj
Else
    PlatosMesa.TopRow = (PlatosMesa.Rows - 1)
    PlatosMesa.Col = 0
    PlatosMesa.Row = (PlatosMesa.Rows - 1)
    txto = "SELECT * FROM TMP_TRANS "
    txto = txto & " WHERE LIN = " & Val(PlatosMesa.Text) & " AND MESA = " & nMesa & " AND CAJERO = " & npNumCaj
End If

rsFixTmpTrans.Open txto, msConn, adOpenStatic, adLockReadOnly

If rsFixTmpTrans.EOF = True Then
    rsFixTmpTrans.Close
    ShowMsg "Por Favor SELECCIONE un Producto", vbBlue, vbYellow
    nCantidad = 1
    Exit Sub
End If

If rsFixTmpTrans!CANT < 0 Then
    'Si la Cantidad es 0 entonces...
    ShowMsg "NO puede dar DESCUENTO a este Producto", vbBlue, vbYellow
    rsFixTmpTrans.Close
    nCantidad = 1
    Exit Sub
End If
    
If Mid(rsFixTmpTrans!DESCRIP, 1, 9) = "DESCUENTO" Then
    ShowMsg "NO puede dar DESCUENTO a un DESCUENTO", vbBlue, vbYellow
    rsFixTmpTrans.Close
    nCantidad = 1
    Exit Sub
End If
    
If Mid(rsFixTmpTrans!Tipo, 1, 1) = "B" Then
    ShowMsg "PRODUCTO YA FUE ANULADO/CORREGIDO/SE DIO DESCUENTO EN LA LINEA " & Val(Mid(rsFixTmpTrans!Tipo, 5, 2)), vbBlue, vbYellow
    rsFixTmpTrans.Close
    nCantidad = 1
    Exit Sub
End If
    
CajLin = CajLin + 1

'------------REVISION DE PAGOS PARCIALES-------------------
'--------------------------------

'Dim nTestDesc As Integer
'INFO: AL CALCULAR EL VALOR DE DESCUENTO x ESTANDAR, ESTE AUTOMATICAMENTE HACE EL REDONDEO
nDescImpre = Format(MiDesc * rsFixTmpTrans!precio * (-1), "STANDARD")
'nTestDesc = Val(Mid(nDescImpre, Len(nDescImpre) + 1, 1))

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' INFO: ELIMINANDO PROCESO DE REDONDEO DE DESCUENTO
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Proceso que quita los centavos del Descuento y los redondea al mas bajo
'y Asigna su valor a nDescImpre
'If nTestDesc = 0 Or nTestDesc = 5 Then
'ElseIf nTestDesc < 5 Then
'    nDescImpre = nDescImpre + (nTestDesc / 100)
'ElseIf nTestDesc > 5 And nTestDesc <= 9 Then
'    nDescImpre = nDescImpre + ((nTestDesc - 5) / 100)
'End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim SOLO_FECHA As String
SOLO_FECHA = Format(Date, "YYYYMMDD")

If nNLinSel <> 0 Then
    'INFO: AGREGANDO CON_TAX (FEB2010)
    cSQL = "INSERT INTO TMP_TRANS "
    cSQL = cSQL & "(CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT,DEPTO,"
    cSQL = cSQL & "PLU,ENVASE,PRECIO_UNIT,PRECIO,FECHA,HORA,TIPO,DESCUENTO,CUENTA,CON_TAX) "
    cSQL = cSQL & " VALUES ("
    cSQL = cSQL & "" & cCaja & "," & npNumCaj & "," & nMesa & "," & nMesero & "," & -1 & "," & CajLin & ","
    cSQL = cSQL & "'DESCUENTO : " & Format(MiDesc, "#.00") & "%'" & "," & 1 & "," & rsFixTmpTrans!depto & ","
    cSQL = cSQL & rsFixTmpTrans!PLU & "," & rsFixTmpTrans!envase & "," & nDescImpre & "," & nDescImpre & ","
    cSQL = cSQL & "'" & SOLO_FECHA & "'" & "," & "'" & Time & "'" & ",'DC-" & nNLinSel & "'," & MiDesc & "," & 0 & "," & rsFixTmpTrans!CON_TAX & ")"
    
    'INFO: MULTICAJA
    'INFO: 07OCT2011. SE GUARDABA EN TIPO LA MISMA LINEA DEL ITEM, AHORA SE PONE LA LINEA QUE ESTA DANDO EL DESCUENTO
    'cSQL2 = "UPDATE TMP_TRANS SET TIPO = 'BDC" & Str((CajLin)) & "' WHERE LIN = " & nNLinSel & " AND MESA = " & nMesa & " AND CAJERO = " & npNumCaj
    cSQL2 = "UPDATE TMP_TRANS SET TIPO = 'BDC" & Str((CajLin)) & "' WHERE LIN = " & nNLinSel & " AND MESA = " & nMesa & " AND CAJERO = " & npNumCaj
Else
    'INFO: AGREGANDO CON_TAX (FEB2010)
    cSQL = "INSERT INTO TMP_TRANS "
    cSQL = cSQL & "(CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT,DEPTO,"
    cSQL = cSQL & "PLU,ENVASE,PRECIO_UNIT,PRECIO,FECHA,HORA,TIPO,DESCUENTO,CUENTA,CON_TAX) VALUES ("
    cSQL = cSQL & "" & cCaja & "," & npNumCaj & "," & nMesa & "," & nMesero & "," & -1 & "," & Val(PlatosMesa.Text) + 1 & ","
    cSQL = cSQL & "'DESCUENTO : " & Format(MiDesc, "#.00") & "%'" & "," & 1 & "," & rsFixTmpTrans!depto & "," & rsFixTmpTrans!PLU & ","
    cSQL = cSQL & rsFixTmpTrans!envase & "," & nDescImpre & "," & nDescImpre & "," & "'" & SOLO_FECHA & "'" & "," & "'" & Time & "'"
    cSQL = cSQL & ",'DC-" & Val(PlatosMesa.Text) & "'," & MiDesc & "," & 0 & "," & rsFixTmpTrans!CON_TAX & ")"
        
    'INFO: MULTICAJA
    'cSQL2 = "UPDATE TMP_TRANS SET TIPO = 'BDC" & Str(Val(PlatosMesa.Text) + 1) & "' WHERE LIN = " & (Val(PlatosMesa.Text)) & " AND MESA = " & nMesa & " AND CAJERO = " & npNumCaj
    'INFO: 07OCT2011. SE GUARDABA EN TIPO LA MISMA LINEA DEL ITEM, AHORA SE PONE LA LINEA QUE ESTA DANDO EL DESCUENTO
    cSQL2 = "UPDATE TMP_TRANS SET TIPO = 'BDC" & Str(Val(PlatosMesa.Text) + 1) & "' WHERE LIN = " & (Val(PlatosMesa.Text)) & " AND MESA = " & nMesa & " AND CAJERO = " & npNumCaj
    CajLin = (Val(PlatosMesa.Text) + 1)
End If
    
msConn.BeginTrans
msConn.Execute cSQL
msConn.Execute cSQL2
msConn.CommitTrans


'INFO: SE ACTIVA EL 26SEP2017
EscribeLog "DESCUENTO x ITEM. " & RegRead("HKCU\Software\SoloSoftware\SoloMix\LastAuthorization") & ". " & rsFixTmpTrans!DESCRIP & ", del (" & MiDesc * 100 & " %), por Valor de: " & Format(nDescImpre, "CURRENCY")

'~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||
'INFO: SOLO IMPRIME LINEA SI ESTA LA IMPRESORA NO FISCAL PRESENTE
'20SEP2013
'~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||
If cFullFiscal = "SI" Then
Else
    ''''''''''ImprimeLineaDetalle 1, "DESCUENTO : " & Format(MiDesc, "#.00") & "%", Format(nDescImpre, "#.00")
    cStr2Print = Format("DESCUENTO : " & Format(MiDesc, "#.00") & "%", "!@@@@@@@@@@@@@@@@@@") + Space(1) & Format(1, "@@@") & Space(1) & Format(Format(nDescImpre, "####.00"), "@@@@@@@")
    'Sys_Pos.Coptr1.PrintTwoNormal FptrSJournalReceipt, cStr2Print, cStr2Print
    Print2_OPOS_Dev cStr2Print
End If

MuestraProductosMarcados

nLineas = PlatosMesa.Rows - 1
PlatosMesa.TopRow = (PlatosMesa.Rows - 1)
rsFixTmpTrans.Close

nCantidad = 1: nPase = 0
Text1(2) = nCantidad
End Sub
Public Sub SetupPantalla(Optional bOpcion As Boolean)
Dim i As Integer
Dim cHandler As String
'Formato de la Pantalla de Facturacion

On Error GoTo ErrCorreccion:

If bOpcion = True Then GoTo OnlyListaPagos:

With PlatosMesa
    For i = 0 To 17
        .ColWidth(i) = 0
    Next
    .ColWidth(0) = 340
    '.ColWidth(1) = 1500
    .ColWidth(1) = 2440
    .ColWidth(2) = 620      'CANT
    .ColWidth(3) = 0
    '.ColWidth(4) = 750:
    .ColWidth(4) = 1020:
    .ColAlignmentFixed(3) = flexAlignRightCenter
    .ColAlignmentFixed(4) = flexAlignRightCenter
    .ColAlignment(3) = flexAlignRightCenter
    .ColAlignment(4) = flexAlignRightCenter
End With

OnlyListaPagos:

With ListaPagos
    .Cols = 4
    .ColWidth(0) = 0: .ColWidth(1) = 2000: .ColWidth(2) = 800:
    .ColWidth(3) = 0
End With
Exit Sub

ErrCorreccion:
cHandler = Err.Number & " - " & Err.description & " EN Public Sub SetupPantalla(Optional bOpcion As Boolean)"
EscribeLog cHandler
ShowMsg cHandler, vbYellow, vbRed
ShowMsg "DESCARTE LA TRANSACCION Y CIERRE EL PROGRAMA, LUEGO VUELVA  A ENTRAR", vbYellow, vbRed
ShowMsg "DESCARTE LA TRANSACCION Y CIERRE EL PROGRAMA, LUEGO VUELVA  A ENTRAR", vbYellow, vbRed
On Error GoTo 0
End Sub
Private Sub QuitarPLUS()
Dim nNum As Integer, lNum As Integer
'OCULTA LOS PLU's DE LA PANTALLA
nNum = rs03.RecordCount

For lNum = 1 To nMaxPLu
    cmdPlus(1).Caption = ""
    If Not IsObject(cmdPlus(lNum)) Then
        Load cmdPlus(lNum)
    End If
    cmdPlus(lNum).Visible = False
Next
cmdPlus(0).Visible = True
End Sub
Private Sub MuestraPLU(ElDepto As Integer)
'VIENE DE HACER CLICK A LOS DEPARTAMENTOS
Dim MiTop As Long, MiLeft As Long, StayLeft As Long
Dim iTam As Integer
Dim cSQL As String

'Muestra los productos a Vender
Set rs03 = New Recordset
'Set rs04 = New Recordset
rs04.Close
'Busca PLUS del Depto
cSQL = "SELECT codigo, depto, descrip, corto, precio1, envases, IMPRESORA, CON_TAX "
cSQL = cSQL & " FROM PLU "
cSQL = cSQL & " WHERE depto = " & ElDepto
'INFO: 15FEB2011 (FUFO LO ENCONTRO PARA MAQUINA DEMO DE SUSHI EXPRESS
cSQL = cSQL & " AND DISPONIBLE = TRUE "
cSQL = cSQL & " ORDER BY CORTO"

rs03.Open cSQL, msConn, adOpenStatic, adLockReadOnly
'Busca Envases del Departamento
cSQL = "SELECT a.depto,a.contenedor,b.descrip "
cSQL = cSQL & " FROM contend_01 as a,contened as b "
cSQL = cSQL & " WHERE a.DEPTO = " & ElDepto
cSQL = cSQL & " AND a.contenedor = b.contenedor "
cSQL = cSQL & " ORDER BY a.depto,a.contenedor"

rs04.Open cSQL, msConn, adOpenStatic, adLockOptimistic
iTam = 0

MiTop = 240: StayLeft = 120
MiLeft = 0: numplu = 0

'Si No hay productos, quitar los que estan visibles
If rs03.EOF Then
    Dim lNum As Integer
    cmdPlus(0).Tag = ""
    cmdPlus(0).Caption = ""
    'For lNum = 0 To rs03.RecordCount - 1
    For lNum = 0 To nNUM_MAX_PLU
        If Not IsObject(cmdPlus(lNum)) Then
            Load cmdPlus(lNum)
        End If
        cmdPlus(lNum).Visible = False
    Next
    nMaxPLu = rs03.RecordCount - 1
    cmdPlus(0).Visible = True
    rs02.MoveFirst
    rs02.Find "CODIGO = " & ElDepto
    If Not rs02.EOF Then
        If rs02!ABIERTO = True Then
            'MsgBox "DEPARTAMENTO ABIERTO", vbCritical, BoxTit
            AddOpenDeptItem
        End If
    End If
    Exit Sub
End If

'SI HAY PRODUCTOS EN EL DEPARTAMENTO, LOS MUESTRO
'PRODUCTOS PARA LA PRIMERA PANTALLA
Do Until rs03.EOF
    If numplu < 1 Then
        cmdPlus(numplu).Caption = rs03!DESCRIP
        cmdPlus(numplu).Tag = rs03!CODIGO
        'Muestra los PLUs del primer departamento
    Else
        If Not IsObject(cmdPlus(numplu)) Then
           Load cmdPlus(numplu)
        End If
        cmdPlus(numplu).Visible = True
        cmdPlus(numplu).Caption = rs03!DESCRIP
        cmdPlus(numplu).Tag = rs03!CODIGO
        cmdPlus(numplu).Left = MiLeft + StayLeft
        cmdPlus(numplu).Top = MiTop
        StayLeft = 120
    End If
    numplu = numplu + 1
    MiLeft = MiLeft + 2400
    Select Case numplu
        Case 3, 6, 9, 12, 15, 18, 21, 24, 27
            MiTop = MiTop + 750
            MiLeft = 0
    End Select
    If numplu = nNUM_MAX_PLU Then Exit Do
    rs03.MoveNext
Loop
nMaxPLu = numplu
End Sub
Private Sub Quita_Subrallado(var As Integer)
Dim i As Integer

i = 0

For i = 0 To cmdDepto.Count - 1
    cmdDepto(i).BackColor = &HC0C0C0
Next
If var <> 67 Then
    cmdDepto(var).BackColor = &HFFFF80
End If
End Sub

Private Sub QuitarDeptos()
Dim nNum As Integer

For nNum = 1 To cmdDepto.Count - 1
    cmdDepto(1).Caption = ""
    cmdDepto(nNum).Visible = False
Next
End Sub

Private Sub Clear_Click()
nPase = 0
nCantidad = 1
Text1(2) = nCantidad
nMontoPagado = 0
axPagado.Text = Format(nMontoPagado, "####.00")

'INFO: FEB2010
'INFO: VERIFICA EL PRINT SERVER, SOLAMENTE SI ES LA CAJA SECUNDARIA
If nMesa = 1 Then
    'INFO: DO NOTHING. PRINT SERVER SHOULD BE RUNNING
Else
    If GetFromINI("Meseros", "PrintServerStatus", DATA_PATH & "soloini.ini") = "ON" Then
        'INFO: DO NOTHING, PRINTSERVER IS UP
    Else
        ShowMsg "LA CAJA PRINCIPAL NO ESTA FUNCIONANDO." & vbCrLf & "SALGA DEL PROGRAMA Y REVISE TODAS LAS IMPRESORAS DE LA CAJA PRINCIPAL.", vbYellow, vbRed
        Call EscribeLog("LA CAJA PRINCIPAL NO ESTA FUNCIONANDO. SALGA DEL PROGRAMA Y REVISE TODAS LAS IMPRESORAS DE LA CAJA PRINCIPAL.")
        Sys_Pos.Label2(2).ForeColor = vbYellow
        Sys_Pos.Label2(2).BackColor = vbRed
        Sys_Pos.Label2(2) = "Revise Caja Principal": Sys_Pos.Label2(2).Refresh
    End If
    
End If

End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : cmbBill_Click
' Autor         : hsequeira
' Fecha        : 12/09/2013
' Proposito   : IMPRIME LA FACTURA, EN LA FISCAL O EN OTRA IMPRESORA
' Upgrade     : 24ABR2024
' Proposito    : COMIENZA LA PROGRAMACION DE FACTURACION ELECTRONICA
'                   VAMOS A HACERLO MAS MODULAR
'---------------------------------------------------------------------------------------
'
Private Sub cmbBill_Click(Index As Integer)
Dim iLen As Integer
Dim txtString As String
Dim n2 As Variant
Dim cSQL As String
Dim nCopiaiISCTransaccion As Single
Dim nMiliseconds As Long
Dim OriginalnMiliseconds As Long

'Debug.Print Format(Time, "hh:mm:sss")

If PlatosMesa.Rows < 1 Then
    ShowMsg "NO HAY NADA MARCADO, FAVOR MARQUE PRODUCTOS", vbBlue, vbYellow
    Exit Sub
End If
If Index = 99 Then  'CUANDO EL MONTO NO ESTA EN LOS ICONOS DE BILLETE
    axPagado.Text = Format(nMontoPagado, "####.00")
Else
    nMontoPagado = nMontoPagado + cmbBill(Index).Tag
    axPagado.Text = Format(nMontoPagado, "####.00")
End If

If nMontoPagado < SubTot Then
    Exit Sub
End If

On Error Resume Next
''''''''''ImprimeLineaSubTot Space(2)
'Print2_OPOS_Dev  Space(2)
'*******************************************
'*******************************************

'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Set rsPropinaMarcadas = New ADODB.Recordset        'DICIEMBRE DE 2014
Set rsPagosMarcados = New ADODB.Recordset        'DICIEMBRE DE 2014

Call PrepareRecordSets
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||


'~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||
'~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||
'INFO: SI ENTRA AQUI, ES PORQUE YA COMIENZA EL PAGO DE LA FACTURA
If cFullFiscal = "SI" Then
    
    'INFO: NOV2011
    'DESHABILTANDO FRAME DE PRODUCTOS Y EL RESTO DE LOS BOTONES QUE
    'PUEDEN DESACTIVAR LA APLICACION.
    DoEvents
    Me.Frame2(1).BackColor = &HC0C0C0
    Me.Frame2(1).Enabled = False
    Me.cmbBill(0).Enabled = False
    Me.cmbBill(1).Enabled = False
    Me.cmbBill(2).Enabled = False
    Me.cmbBill(3).Enabled = False
    Me.cmdAcepPago.Enabled = False
    Me.cmpPago.Enabled = False
    Me.Command13(0).Enabled = False
    Me.Command13(1).Enabled = False
    Me.Correccion.Enabled = False

    Call GetNewNumTrans
    
    'Dim Tiempo1 As Long
    'Dim Tiempo2 As Long
    '''
    'Tiempo1 = GetTickCount()
    
    '~~~~~~~~~~~~~~~~~~~~~~
    'INFO: 28MAY2014
    If KDS_Activo = "SI" Then
        Call KDS_File(Trim(Str(nNUMTrans)))
    End If
    '~~~~~~~~~~~~~~~~~~~~~~
    
    If IMPRIME_TICKET = True Then
        Call FISCAL_FAST_Encabezado(Trim(Str(nNUMTrans)))
    Else
        Call FISCAL_FAST_Encabezado
    End If
        
        'Tiempo2 = GetTickCount()
        'Debug.Print "FISCAL_FAST_Encabezado " & vbTab & GetTickCount() - Tiempo1
        
    Call FISCAL_FAST_PrintArticulos
        'Tiempo2 = GetTickCount()
        'Debug.Print "FISCAL_FAST_PrintArticulos" & vbTab & GetTickCount() - Tiempo1
    Call FISCAL_FAST_IniciaCierreFactura(0#)
        'Tiempo2 = GetTickCount()
        'Debug.Print "FISCAL_FAST_IniciaCierreFactura" & vbTab & GetTickCount() - Tiempo1
    Call FISCAL_FAST_Pagos("SIMPLE", CSng(nMontoPagado))
        'Tiempo2 = GetTickCount()
        'Debug.Print "FISCAL_FAST_Pagos" & vbTab & GetTickCount() - Tiempo1
        'Debug.Print ">>>>>>>>>>>>>>>>>"
    
    Select Case nMarca
        Case 0      'STAR
            'INFO: 24JUN2011
            Call PutISC(FISCAL_.GetCurrentITBMS)
        Case 1      'BEMATECH
            'INFO: UPDATE 01DIC2011
            'Call PutISC(Format(iISCTransaccion, "STANDARD"))
            Call PutISC(FISCAL_.GetCurrentITBMS)
        Case Else
    End Select
    
'Debug.Print "=========================="
    
    nCopiaiISCTransaccion = iISCTransaccion
    iISCTransaccion = 0
    '~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||
    '~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||
Else

    'ListaPagos.AddItem "02" & Chr(9) & "EFECTIVO" & Chr(9) & Format(lbMonto - nCambio, "STANDARD") & Chr(9) & Format(lbMonto, "STANDARD")
    
    ListaPagos.AddItem "01" & Chr(9) & "EFECTIVO" & Chr(9) & PLU.SubTot.Caption & Chr(9) & PLU.SubTot.Caption
    
    On Error Resume Next
    
    Call AddPago("PAGO", "01", "EFECTIVO", SubTot, SubTot)
    
     Call GetNewNumTrans
    
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    If FE_Facil_ONLINE Then
        ProgBar.value = 30
        'If GenerateJson(mi_Ambiente, rs00!TRANS + 1, rsPagosMarcados, ListaPagos, bDescDiplomatico, Str(Label3)) Then
        'If GenerateJson(mi_Ambiente, rs00!TRANS + 1, rsPagosMarcados, ListaPagos, False, Format(Str(SubTot), "#0.00")) Then
        
        If GenerateJson(mi_Ambiente, nNUMTrans, rsPagosMarcados, ListaPagos, False, Format(Str(SubTot), "#0.00")) Then
        
            'ProgBar.value = 50
            ProgBar.value = 85
            ProgBar.value = 100
            Rem MiDebug "1cmdBill.ProgBar.value = 100"
            'Call CierreObjetos
            'MsgBox "OK  " & msConn.IsolationLevel
        Else
            If AllowFE_to_Continue Then
                ProgBar.value = 100
                Rem MiDebug "2cmdBill.ProgBar.value = 100"
            Else
                'MsgBox "FAIL  " & msConn.IsolationLevel
                Call CierreObjetos
                ShowMsg "INTENTE HACER EL COBRO DE NUEVO", vbBlue, vbCyan
                '||||||||||||||||||||||||||||||||
                'UPDATE 12ENE2025
                Call Clear_Click
                '||||||||||||||||||||||||||||||||
                Call Check_FE_Website
                Call ChangeColorOfProgBar
                Exit Sub
            End If
            'FALLO EN FACTURA ELECTRONICA
            Call Check_FE_Website
        End If
    Else
        'Call OFF_Line_Dummy(Label3.Caption, rs00!TRANS + 1)
        ProgBar.value = 50
        
        If GenerateJson_OFFLINE(mi_Ambiente, nNUMTrans, rsPagosMarcados, ListaPagos, False, Format(Str(SubTot), "#0.00")) Then
        
        End If
       
        Rem Call OFF_Line_Dummy(Label3.Caption, nNUMTrans)
        ProgBar.value = 85
        Call ImprFactura_____Falso
        ProgBar.value = 100
        Rem MiDebug "Call ImprFactura.ProgBar.value = 100"
        'Call Actualizador
        'Call Actualizador(lbNOTax, lbTAX, nCopiaiISCTransaccion)
    End If
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "BEGIN")
    Rem Debug.Print "OPOSTransactionPrint(Sys_Pos.Coptr1.name, BEGIN)"
    'COMIENZA IMPRESION DE FACTURA (26ENE2025) con Print2_OPOS_Dev
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

        Print2_OPOS_Dev "------------------------------"
        iLen = Len(Format(SubTot, "STANDARD"))
        txtString = "SUB TOTAL : " & Space(18 - iLen) & Format(SubTot - FormatCurrency(iISCTransaccion, 2), "STANDARD")
        Print2_OPOS_Dev txtString
        'Print2_OPOS_Dev  Space(2)
        
        iLen = Len(Format(iISCTransaccion, "STANDARD"))
        'INFO: JULIO 2010 NO SE PUEDE PONER EL 5%, EL ITBMS ES CALCULADO x PRODUCTO
        'txtString = "ITBMS(5%) : " & Space(18 - iLen) & Format(iISCTransaccion, "STANDARD")
        'INFO: 12SEP2013
        txtString = "ITBMS     : " & Space(18 - iLen) & Format(iISCTransaccion, "STANDARD")
        Print2_OPOS_Dev txtString
        
        Call PutISC(Format(iISCTransaccion, "STANDARD"))
        
        iLen = Len(Format(SubTot, "STANDARD"))
        txtString = "TOTAL     : " & Space(18 - iLen) & Format(SubTot, "STANDARD")
        Print2_OPOS_Dev txtString
        Print2_OPOS_Dev Space(2)
        'SubTot = SubTot + iISCTransaccion
        'SubTot = FormatCURRENCY(SubTot + iISCTransaccion, 2)
        iISCTransaccion = 0
        '*******************************************
        '*******************************************
        On Error GoTo 0
    
        iLen = Len(Format(nMontoPagado, "STANDARD"))
        txtString = "PAGADO    : " & Space(18 - iLen) & Format(nMontoPagado, "STANDARD")
        Print2_OPOS_Dev txtString
        
        iLen = Len(Format((nMontoPagado - SubTot), "STANDARD"))
        txtString = "CAMBIO    : " & Space(18 - iLen) & Format((nMontoPagado - SubTot), "STANDARD")
        Print2_OPOS_Dev txtString
        Print2_OPOS_Dev Space(2)
        
        Print2_OPOS_Dev Date & Space(2) & Time
        Print2_OPOS_Dev "SERIAL:" & rs00!SERIAL
        'INFO: AGO2009
        'Sys_Pos.Coptr1.PrintTwoNormal FptrSJournalReceipt , "TRANS# " & rs00!TRANS + 1, "TRANS# " & rs00!TRANS + 1
        
        'INFO: DIC2009
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
                'Call GetNewNumTrans
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    
        '~~~~~~~~~~~~~~~~~~~~~~
        'INFO: 28MAY2014
        If KDS_Activo = "SI" Then
            Call KDS_File(Trim(Str(nNUMTrans)))
        End If
        '~~~~~~~~~~~~~~~~~~~~~~
    
        'Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Chr(&H1B) & "!" & Chr(32) & "FACTURA # " & cCaja & " - " & nNUMTrans & Chr(&HD) & Chr(&HA)
        'Sys_Pos.Coptr1.PrintNormal STATION_2PRINT, Chr(&H1B) & Chr(&H46) & "Cajero : " & cNomCaj & Chr(&HD) & Chr(&HA)
        
        Print2_OPOS_Dev Chr(&H1B) & "!" & Chr(32) & "FACTURA # " & cCaja & " - " & nNUMTrans
        Print2_OPOS_Dev Chr(&H1B) & Chr(&H46) & "Cajero: " & cNomCaj
        Rem SE INVIERTEN SUCURSAL Y CAJA. Print2_OPOS_Dev "Caja: " & o_X_FF_Branch & " - Sucursal: " & o_poscode
        Print2_OPOS_Dev "Caja: " & o_poscode & " - Sucursal: " & o_X_FF_Branch
    
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
    
        If nImp = 0 Then
            If FE_QRData <> "" Then
                With Sys_Pos.Coptr1
                         Rem FE_QRData = "https://wa.me/p/9159792377440884/50769261093"
                        Rem .PrintBarCode PtrSReceipt, FE_QRData, PTR_BCS_UQRCODE, .RecLineHeight * 0.6, .RecLineWidth * 0.6, PtrBcCenter, PtrBcTextNone&
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
        txtString = "DOCUMENTO FACTURA"
        Print2_OPOS_Dev Space(nEspacio) & txtString
        txtString = "RECEPTOR: CONSUMIDOR FINAL"
        Print2_OPOS_Dev Space(nEspacio) & txtString
    
    
        'INFO: CUMPLIR CON LA LEY
        Print2_OPOS_Dev Space(3)
        txtString = "Documento validado por Factura Facil S.A con RUC 155723374-2-2022, es Proveedor Autorizado Calificado, Resolucion No. 201-2167 de 14/03/2023."
        Print2_OPOS_Dev Space(nEspacio) & txtString
        
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
   
    
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        'INFO: 22OCT2017.
        
        If nTransDomicilio Then
            
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
        
        'DoEvents - DESACTIVADO PARA OPOS
        ''Sys_Pos.Cocash1.ClaimDevice 5000
    
        'Sys_Pos.CoCash1.OpenDrawer     'INFO: 14FEB2017. USANDO DIRECTIO QUE NO ESPERA RESPUESTA DE LA GAVETA
        'Call Sys_Pos.Cocash1.DirectIO(DRW_DI_OPEN_DRAWER, 0, "")
        Rem MiDebug "Antes de OpenGavetaDinero"
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem 15ENE2026 Call OpenGavetaDinero
        Call OpenGavetaDinero
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem MiDebug "DESPUES de OpenGavetaDinero"
        'INFO:REDUCIENDO EL ESPACIO DE ABAJO DE 10 A 5
        '06JUL2024
        'For i = 1 To 10
        For i = 1 To 5
            Print2_OPOS_Dev Space(2)
        Next
        'DoEvents - DESACTIVADO MAY2014

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, "END")
    
    Rem Debug.Print "Call OPOSTransactionPrint(Sys_Pos.Coptr1.name, END)"
    'FIN TRANSACCION DE IMPRESION 26ENE2025
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    Sleep MUNBYNSeconds
    
    Sys_Pos.Coptr1.CutPaper 100
    
    If Sys_Pos.Coptr1.RecEmpty Then
        ShowMsg "¡¡ A T E N C I O N !!" & vbCrLf & vbCrLf & "SE ESTA ACABANDO EL PAPEL DE PRE-CUENTA / PEDIDOS, SE LE RECOMIENDA SALIR DEL PROGRAMA PARA CAMBIAR EL PAPEL Y LUEGO REGRESAR AL PROGRAMA DE VENTAS", vbRed, vbYellow
    End If
    
End If
'~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||

'===---===---===---===---===---===---===
If IMPRIME_TICKET = True Then Call ImprimeTicket(nNUMTrans)

'INFO: ENVIAR PEDIDO A LA COCINA
'If IMPRIME_PEDIDO = True Then Call ImprimeEnCocina(nNumTrans)
If IMPRIME_PEDIDO = True Then
    'INFO: 16DIC2014
    'INFO: ACTUALIZANDO PARA HACER UNA SOLA APERTURA A MESASPED, AHORRA 50% DE TIEMPO
    'Y HACER QUE EL PEDIDO SE EJECUTE EN UN SOLO COMMIT A MESASPED.
    MesasPED "OPEN"
    Call SendToPrintServer(nNUMTrans)
    MesasPED "CLOSE"
End If
If DUPLICA_PEDIDO = True Then Call DuplicaPedido(nNUMTrans)
'===---===---===---===---===---===---===

'===---===---===---===---===---===---===
'===---===---===---===---===---===---===
'INFO: DIC2009
'MOVIENDO EL VUELTO HASTA DESPUES DE LA ACTUALIZACION DE LAS TABLAS

On Error Resume Next
    nCambio = nMontoPagado
    n2 = SubTot
    nCambio = nCambio - n2
    If nCambio <> 0# Then
        Vuelto.Show 1
    End If
    'Debug.Print Format(Time, "hh:mm:sss")
    nMiliseconds = Format(Time(), "ssss")
    OriginalnMiliseconds = nMiliseconds
On Error GoTo 0

nCambio = 0#
nMontoPagado = 0#
'===---===---===---===---===---===---===
'===---===---===---===---===---===---===

'cSQL = "INSERT INTO TRANSAC_PAGO (NUM_TRANS,TIPO_PAGO,CAJERO,LIN,MONTO, MESA) VALUES ("
'cSQL = cSQL & (nNUMTrans) & "," & 1 & "," & npNumCaj & "," & 1 & "," & Format(SubTot, "#0.00") & "," & nMesa & ")"

'msConn.BeginTrans
'msConn.Execute cSQL
'msConn.CommitTrans

nMiliseconds = Format(Time(), "ssss") - OriginalnMiliseconds
'Debug.Print "== SQL ==  " & nMiliseconds & " ==========="

axPagado.Text = 0#

'Call Pagos.Actualizador(lbNOTax, lbTAX, nCopiaiISCTransaccion)

Call Actualizador(lbNOTax, lbTAX, nCopiaiISCTransaccion)

'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '23JUL2024
    'INFO: TRATA DE RECUPERAR LA CONEXION A INTERNET
    
    If cLogMode = "SI" Then
        '16FEB2025
        'INFO: SI ESTOY EN LOGMODE, ME QUEDO OFF LINE
        If Not FE_Facil_ONLINE Then
            FE_Facil_ONLINE = False
        Else
        End If
    Else
        Call Check_FE_Website("SILENCIO")
        Call ChangeColorOfProgBar
    End If
    
    ProgBar.value = 0
    Call CierreObjetos
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

'nMiliseconds = Format(Time(), "ssss") - OriginalnMiliseconds
'Debug.Print "== Pagos.Actualizador ==  " & nMiliseconds & " ==========="

nCopiaiISCTransaccion = 0

SBTot = 0#

'INFO: 24NOV2011
'HABILTANDO FRAME DE PRODUCTOS
Me.Frame2(1).Enabled = True
Me.Frame2(1).BackColor = vbBlack
Me.cmbBill(0).Enabled = True
Me.cmbBill(1).Enabled = True
Me.cmbBill(2).Enabled = True
Me.cmbBill(3).Enabled = True
Me.cmdAcepPago.Enabled = True
Me.cmpPago.Enabled = True
Me.Command13(0).Enabled = True
Me.Command13(1).Enabled = True
Me.Correccion.Enabled = True

Call MuestraProductosMarcados

nMiliseconds = Format(Time(), "ssss") - OriginalnMiliseconds
'Debug.Print "== MuestraProductosMarcados ==  " & nMiliseconds & " ==========="

Call Clear_Click

nMiliseconds = Format(Time(), "ssss") - OriginalnMiliseconds
'Debug.Print "== Clear_Click ==  " & nMiliseconds & " ==========="

'~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||
'INFO: 06OCT2011.
'INFO: 16FIC2014. MOVIENDO LA IMPRESION DE TITULOS DESPUES DEL DISPLAY
If cFullFiscal = "SI" Then
Else
    Call ImprimeTitulos
End If
'~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||

If IsObject(Pagos) = True Then Unload Pagos

'INFO: UPDATE 23JUL2024
ListaPagos.Clear
ListaPagos.Rows = 0

'INFO: DOMICILIOS MARCADOS DESDE OTRA APLICACION
'21ABR2018
'Call CheckPedidoDomicilio
nMiliseconds = Format(Time(), "ssss") - nMiliseconds
' Debug.Print "======= " & nMiliseconds & " ==========="

End Sub
Private Sub cmdAcepPago_Click()
Dim nCant As Long
On Error GoTo ErrAdm:
nCant = Val(Text1(2))
If nCant / 100 < Format(SubTot - axPagado.Text, "STANDARD") Then
    ShowMsg "EL MONTO QUE SE MARCO (" & Format(nCant / 100, "CURRENCY") & ")  DEBE SER MAYOR QUE EL TOTAL o SALDO DE LA CUENTA", vbRed, vbYellow
    Call Clear_Click
    Exit Sub
End If
If nCant > 1 Then
    nMontoPagado = nMontoPagado + (nCant / 100)
    cmbBill_Click (99)
End If
On Error GoTo 0
Exit Sub

ErrAdm:
End Sub
Private Sub cmdDepto_Click(Index As Integer)
Quita_Subrallado (Index)
TextEnv = ""
QuitarPLUS
ElDepto = Arreg_Deptos(Index)
MuestraPLU (Arreg_Deptos(Index))
nGlobEnv = 0: nNLinSel = 0


'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'Unload Acompanantes
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

'INFO: 22OCT2017. AGREGAR OPCION DE DOMICILIO.
If nHayDomicilio = ElDepto Then
    If ShowMsg("OPCION DE DOMICILIO" & vbCrLf & vbCrLf & "¿Desea Entrar a Capurar la Informacion del Cliente para DELIVERY?", vbYellow, vbBlue, vbYesNo) = vbYes Then
        nTransDomicilio = True
        DomiClientes.Show 1
    End If
End If

End Sub

Private Sub cmdDomicilio_Click()
On Error GoTo ErrAdm:
If Not Timer1.Enabled Then
    Timer1.Enabled = True
Else
    Timer1.Enabled = False
End If
On Error GoTo 0
Exit Sub

ErrAdm:
 EscribeLog "Error en Timer.cmdDomicilio de PLU"
End Sub

Private Sub cmdPLUNext_Click()
Dim MiTop As Long, MiLeft As Long

If cmdPlus(nNUM_MAX_PLU - 1).Visible = False Then
    'SI NO ESTA nNUM_MAX_PLU - 1 VISIBLE NO ES NECESARIO IR A
    'BUSCAR EL RESTO DE LOS PRODUCTOS
    Exit Sub
End If

numplu = 0
'Llamar proc. de limpiar PLUS anteriores
QuitarPLUS
If rs03.EOF Then
    'PARA EMPEZAR LA PROXIMA PANTALLA ME POSICIONO EN
    'EL PLU ANTERIOR COMO PRIMERO
    rs03.MovePrevious
    cmdPlus(0).Caption = rs03!DESCRIP
    cmdPlus(0).Tag = rs03!CODIGO
End If

MiTop = 240: StayLeft = 120
MiLeft = 0: numplu = 0

Do Until rs03.EOF
    If numplu < 1 Then
        cmdPlus(numplu).Caption = rs03!DESCRIP
        cmdPlus(numplu).Tag = rs03!CODIGO
        'Muestra los PLUs del primer departamento
    Else
        If Not IsObject(cmdPlus(numplu)) Then
           Load cmdPlus(numplu)
        End If
        cmdPlus(numplu).Visible = True
        cmdPlus(numplu).Caption = rs03!DESCRIP
        cmdPlus(numplu).Tag = rs03!CODIGO
        cmdPlus(numplu).Left = MiLeft + StayLeft
        cmdPlus(numplu).Top = MiTop
        StayLeft = 120
    End If
    numplu = numplu + 1
    MiLeft = MiLeft + 2400
    Select Case numplu
        Case 3, 6, 9, 12, 15, 18, 21, 24, 27
            MiTop = MiTop + 750
            MiLeft = 0
    End Select
    If numplu = nNUM_MAX_PLU Then Exit Do
    rs03.MoveNext
Loop
nMaxPLu = numplu
nNLinSel = 0
End Sub

Private Sub cmdPLUTop_Click()
Dim nNum As String
num = 0
'Tengo que saber quien se ve de primero para mostrar
'los 11 anteriores

If rs03.EOF Then
    If rs03.BOF Then Exit Sub
    rs03.MovePrevious
    cmdPlus(num).Caption = rs03!DESCRIP
    cmdPlus(numplu).Tag = rs03!CODIGO
End If
nNum = rs03!CODIGO

If (numplu - nNUM_MAX_PLU) <= 0 Then
    ' Desde el principio
    rs03.MoveFirst
End If
MiTop = 240: StayLeft = 120
MiLeft = 0: numplu = 0

Do Until rs03.EOF
    If numplu < 1 Then
        cmdPlus(numplu).Caption = rs03!DESCRIP
        cmdPlus(numplu).Tag = rs03!CODIGO
        'Muestra los PLUs del primer departamento
    Else
        If Not IsObject(cmdPlus(numplu)) Then
           Load cmdPlus(numplu)
        End If
        cmdPlus(numplu).Visible = True
        cmdPlus(numplu).Caption = rs03!DESCRIP
        cmdPlus(numplu).Tag = rs03!CODIGO
        cmdPlus(numplu).Left = MiLeft + StayLeft
        cmdPlus(numplu).Top = MiTop
        StayLeft = 120
    End If
    numplu = numplu + 1
    MiLeft = MiLeft + 2400
    Select Case numplu
        Case 3, 6, 9, 12, 15, 18, 21, 24, 27
            MiTop = MiTop + 750
            MiLeft = 0
    End Select
    If numplu = nNUM_MAX_PLU Then Exit Do
    rs03.MoveNext
Loop
nMaxPLu = numplu
nNLinSel = 0

End Sub
Private Sub cmdSalir_Click()
Dim localResp As Variant
Dim cTexto As String

'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
'INFO: SI LA FISCAL TIENE EL CONTROL NO HACER NADA
' 02OCT2012
If cFullFiscal = "SI" Then
Else
    'INFO: 14MAY2019
    'INFO. MUNBYN CANNOT HANDLE CASHDRAWER
    'If OPOS_DevName = "POSPrinter80" Or OPOS_DevName = "SRP-E300" Then
    If Left(OPOS_DevName, 12) = "POSPrinter80" Or Left(OPOS_DevName, 10) = "SRP-E300" Then
    Else
        If Sys_Pos.Cocash1.DrawerOpened = True Then
            ShowMsg "POR FAVOR CIERRE LA GAVETA DE DINERO ANTES DE CONTINUAR", vbBlue, vbYellow
            Beep
            Exit Sub
        End If
    End If
End If

'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~

If PlatosMesa.Rows > 0 Then
    cTexto = "TIENE PRODUCTOS PENDIENTES, ES IMPOSIBLE SALIR DEL PROGRAMA AHORA" & vbCrLf & vbCrLf
    cTexto = cTexto & "Debe Facturarlos o Anularlos. Despues Proceder a Cerrar la Factura"
    ShowMsg cTexto, vbYellow, vbRed
    Me.MousePointer = vbHourglass
    EscribeLog cTexto
    Me.MousePointer = vbDefault
    Exit Sub
    'localResp = MsgBox("TIENE PRODUCTOS PENDIENTES, DESEA SALIR ?", vbYesNo + vbQuestion, "¿ DESEA SALIR ?")
    'If localResp = vbNo Then
        'Exit Sub
    'End If
End If

'SI TODO ESTA BIEN, PERMITE SALIR DEL PROGRAMA
cmdSalir.Tag = 0

If rsTmpAco.State = adStateOpen Then rsTmpAco.Close
Set rsTmpAco = Nothing
Sys_Pos.Text1 = ""
EscribeLog "Salida a Pantalla de Login"
Unload Me
End Sub

Private Sub cmpPago_Click()
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'Unload Acompanantes
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

    Call ClearCliente
    Call Pagos.MuestraPanta_Pagos(lbNOTax, lbTAX)

'Pagos.Show 1
'Pagos.Height = 6030
'Pagos.Width = 8715
End Sub

Private Sub Command13_Click(Index As Integer)
Dim DescResp As Variant
'||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Dim cSQLFE As String
'||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

Select Case Index
    '~~~~~~~~~~~~~~~~~~~~~~~~~~
    Case 0  'DESCUENTO
    '~~~~~~~~~~~~~~~~~~~~~~~~~~
        'Descuento al ultimo producto de la lista
        If HaySeguridad Then
            txtInfo = "CLAVE PARA DESCUENTO"
            OkAnul = 0
            AskClave.Show 1
            If OkAnul = 1 Then
                Call DescProducto
                Call SetupPantalla
                
                'Unload Acompanantes
                
                '||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
                '||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
                If IsEmpty(nCta) Then miCuenta = 0 Else miCuenta = nCta
                cSQLFE = "SELECT * FROM TMP_TRANS WHERE MESA = " & nMesa & " AND CUENTA = " & miCuenta & " ORDER BY LIN"
                Call FE_PreCheckRecordSet(cSQLFE)
                '||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
                '||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

            End If
            OkAnul = 0
        Else
            OkAnul = 1
            Call DescProducto
            Call SetupPantalla
            '||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
            '||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
            If IsEmpty(nCta) Then miCuenta = 0 Else miCuenta = nCta
            cSQLFE = "SELECT * FROM TMP_TRANS WHERE MESA = " & nMesa & " AND CUENTA = " & miCuenta & " ORDER BY LIN"
            Call FE_PreCheckRecordSet(cSQLFE)
            '||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
            '||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        End If
    '~~~~~~~~~~~~~~~~~~~~~~~~~~
    Case 1  'ANULACION DE LINEA
    '~~~~~~~~~~~~~~~~~~~~~~~~~~
        If HaySeguridad Then
            txtInfo = "CLAVE PARA ANULACION"
            OkAnul = 0
            AskClave.Show 1
            If OkAnul = 1 Then
                BorraLin.Show 1
                Call SetupPantalla
                
                'Unload Acompanantes
            
            End If
            OkAnul = 0
        Else
            OkAnul = 1
            BorraLin.Show 1
            Call SetupPantalla
        End If
End Select
nNLinSel = 0

End Sub
Private Sub Command2_Click()
'AVANZA HACIA ABAJO
Dim num As Integer
num = 0
nNLinSel = 0

'Llamar proc. de limpiar deptos anteriores
Quita_Subrallado (67)
QuitarDeptos

If rs02.EOF = True Then
    'SI EOF, se mueve al anterior
    rs02.MovePrevious
    cmdDepto(0).Caption = rs02!corto
    cmdDepto(0).Tag = rs02!CODIGO
End If

Do Until rs02.EOF
    If num < 1 Then
        cmdDepto(num).Caption = rs02!corto
        Arreg_Deptos(num) = rs02!CODIGO
    Else
        If Not IsObject(cmdDepto(num)) Then
           Load cmdDepto(num)
        End If
        cmdDepto(num).Caption = rs02!corto
        Arreg_Deptos(num) = rs02!CODIGO
        cmdDepto(num).Left = 120
        cmdDepto(num).Top = cmdDepto(num - 1).Top + 660
        cmdDepto(num).Visible = True
    End If
    num = num + 1
    If num = nNUM_MAX_DEPTOS Then Exit Do
    rs02.MoveNext
Loop
End Sub
Private Sub Command3_Click()
'AVANZA HACIA ARRIBA
Dim nNum As Integer
num = 0
nNLinSel = 0

rs02.MoveFirst
rs02.Find "codigo = " & Arreg_Deptos(0)

If rs02.EOF Then
    'El PG-DOWN llego a la ultima pantalla
    rs02.MovePrevious
    'cmdDepto(num).Caption = rs02!CORTO
End If

rs02.Move -11
If rs02.BOF Then rs02.MoveFirst

Quita_Subrallado (67)
QuitarDeptos

'CARGANDO EL CODIGO DE LOS DEPARTAMENTO EN LOS BOTONES DISPONIBLES
Do Until rs02.EOF
    If num < 1 Then
        cmdDepto(num).Caption = rs02!corto
        Arreg_Deptos(num) = rs02!CODIGO
    Else
        If Not IsObject(cmdDepto(num)) Then
           Load cmdDepto(num)
        End If
        cmdDepto(num).Caption = rs02!corto
        Arreg_Deptos(num) = rs02!CODIGO
        cmdDepto(num).Left = 120
        cmdDepto(num).Top = cmdDepto(num - 1).Top + 660
        cmdDepto(num).Visible = True
    End If
    num = num + 1
    If num = nNUM_MAX_DEPTOS Then Exit Do
    rs02.MoveNext
Loop

End Sub
Private Sub cmdPlus_GotFocus(Index As Integer)
cmdPlus(Index).BackColor = &HFFFF00
End Sub
Private Sub cmdPlus_LostFocus(Index As Integer)
cmdPlus(Index).BackColor = &HC0C0C0
End Sub
'---------------------------------------------------------------------------------------
' Procedimiento : cmdPlus_Click
' Autor       : hsequeira
' Fecha       : 12/09/2013
' Proposito   : IMPRIME CADA PLU MARCADO EN LA IMPRESORA NO-FISCAL
'---------------------------------------------------------------------------------------
'
Private Sub cmdPlus_Click(Index As Integer)
Dim cSQL As String
'Dim Cad As Date
Dim nLineas As Long
Dim i As Integer
Dim SOLO_FECHA As String
Dim cStr2Print As String

If cmdPlus(Index).Tag = "" Then Beep: Exit Sub

'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
'INFO: HAY QUE AVERIGUAR ESTA VALIDACION
''If Sys_Pos.Cocash1.DrawerOpened = True Then
''    MsgBox "POR FAVOR CIERRE LA GAVETA DE DINERO ANTES DE CONTINUAR", vbExclamation, BoxTit
''    Beep
''    Exit Sub
''End If
'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~

'INFO: 3OCT2017
If HayDecimales Then
    If Me.Text1(2).Text > 10000 Then
        Beep
        ShowMsg "LA CANTIDAD MARCADA (" & nCantidad & ") ES INCORRECTA, FAVOR CORREGIR CON LA TECLA DE CLEAR" & vbCrLf & "NO PUEDEN MARCARSE MAS DE 100", vbRed, vbYellow
        Exit Sub
    End If
Else
    If nCantidad > 100 Then
        Beep
        ShowMsg "LA CANTIDAD MARCADA (" & nCantidad & ") ES INCORRECTA, FAVOR CORREGIR CON LA TECLA DE CLEAR" & vbCrLf & "NO PUEDEN MARCARSE MAS DE 100", vbRed, vbYellow
        Exit Sub
    End If
End If
'INFO: 16FEB2011
'If nCantidad > 100 Then
'INFO: 3OCT2017


On Error Resume Next
rs03.MoveFirst
On Error GoTo 0
rs03.Find "codigo = " + cmdPlus(Index).Tag
If rs03.EOF Then Beep: Exit Sub

i = 0
'* Cad = FormatDateTime(Now, vbGeneralDate)

nAcoTop = 240
CajLin = CajLin + 1
SOLO_FECHA = Format(Date, "YYYYMMDD")


' Call Acompanantes.LoadAcompanantes(rs03!CODIGO)
    
cSQL = "INSERT INTO TMP_TRANS "
cSQL = cSQL & "(CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT,DEPTO,PLU,ENVASE,PRECIO_UNIT,PRECIO,FECHA,HORA,TIPO,DESCUENTO,CUENTA,IMPRESO,IMPRESORA, CON_TAX) "
cSQL = cSQL & " VALUES ("
cSQL = cSQL & "" & cCaja & "," & npNumCaj & "," & nMesa & "," & nMesero & "," & -1 & "," & CajLin & "," & "'"
'INFO: 3OCT2017
If HayDecimales Then
    cSQL = cSQL & rs03!DESCRIP + TextEnv & "'" & "," & 1 & "," & rs03!depto & "," & rs03!CODIGO & ","
    cSQL = cSQL & 0 & "," & rs03!PRECIO1 & "," & rs03!PRECIO1 * Me.Text1(2) & "," & "'" & SOLO_FECHA & "'" & "," & "'" & Time & "'"
Else
    cSQL = cSQL & rs03!DESCRIP + TextEnv & "'" & "," & nCantidad & "," & rs03!depto & "," & rs03!CODIGO & ","
    cSQL = cSQL & 0 & "," & rs03!PRECIO1 & "," & rs03!PRECIO1 * nCantidad & "," & "'" & SOLO_FECHA & "'" & "," & "'" & Time & "'"
End If
cSQL = cSQL & ",'  '," & 0# & "," & 0 & ",0," & rs03!IMPRESORA
'INFO: AG02009
cSQL = cSQL & "," & rs03!CON_TAX & ")"

msConn.BeginTrans
msConn.Execute cSQL
msConn.CommitTrans

'~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||
'~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||
'INFO: 6OCT2011
If cFullFiscal = "SI" Then
    'INFO: NO HACER NADA, SE ESTA CAMBIANDO PARA HACERLO AL FINAL. CUANDO SE COBRA
Else
    'INFO: 3OCT2017
    If HayDecimales Then
        cStr2Print = Format(Mid(rs03!DESCRIP, 1, 18), "!@@@@@@@@@@@@@@@@@@") + Space(1) & Format(Me.Text1(2), "@@@") & Space(1) & Format(Format(rs03!PRECIO1 * Me.Text1(2), "####.00"), "@@@@@@@")
    Else
        cStr2Print = Format(Mid(rs03!DESCRIP, 1, 18), "!@@@@@@@@@@@@@@@@@@") + Space(1) & Format(nCantidad, "@@@") & Space(1) & Format(Format(rs03!PRECIO1 * nCantidad, "####.00"), "@@@@@@@")
    End If
    Print2_OPOS_Dev cStr2Print
    '-> Sys_Pos.Coptr1.PrintTwoNormal FptrSJournalReceipt, cStr2Print, cStr2Print
    
    If Sys_Pos.Coptr1.ResultCode = OposEIllegal Then
        'SI HAY PROBLEMAS DE IMPRESION
        'SE IMPRIME LA INFORMACION MINIMA POSIBLE
        'INFO: 3OCT2017
        If HayDecimales Then
            cStr2Print = Format(Mid(rs03!DESCRIP, 1, 12), "!@@@@@@@@@@@@") + Space(1) & Format(Me.Text1(2), "@@@") & Space(1) & Format(Format(rs03!PRECIO1 * Me.Text1(2), "####.00"), "@@@@@@@")
        Else
            cStr2Print = Format(Mid(rs03!DESCRIP, 1, 12), "!@@@@@@@@@@@@") + Space(1) & Format(nCantidad, "@@@") & Space(1) & Format(Format(rs03!PRECIO1 * nCantidad, "####.00"), "@@@@@@@")
        End If
        Print2_OPOS_Dev cStr2Print
        '-> Sys_Pos.Coptr1.PrintTwoNormal FptrSJournalReceipt, cStr2Print, cStr2Print
    End If
End If
'~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||
'~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||

Call MuestraProductosMarcados

If HayAcompanates(rs03!CODIGO) Then
    nPLUPassed = rs03!CODIGO
    nDeptoPassed = rs03!depto
    'Call Acompanantes.LoadAcompanantes
Else
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    'Unload Acompanantes
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
End If

On Error GoTo ErrAdm:

nLineas = PlatosMesa.Rows - 1
PlatosMesa.TopRow = (PlatosMesa.Rows - 1)
nCantidad = 1: nPase = 0: nNLinSel = 0: Text1(2) = nCantidad
On Error GoTo 0
Exit Sub

ErrAdm:
If Err.Number = 30009 Then
    ShowMsg "Selección de Producto. Error # (" & Err.Number & ") La Base de Datos esta INESTABLE, no se pueden realizar transacciones" & vbCrLf & "CONTACTE A SOLO POS MENU INMEDIATAMENTE", vbYellow, vbRed
Else
    ShowMsg "Selección de Producto. (" & Err.Number & ") - " & Err.description, vbYellow, vbRed
End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : HayAcompanates
' Author    : hsequeira
' Date      : 07/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function HayAcompanates(nPluID As Integer) As Boolean
HayAcompanates = False
On Error Resume Next
rsAcomPLU.MoveFirst
rsAcomPLU.Find "PLU_ID = " & nPluID
If rsAcomPLU.EOF Then HayAcompanates = False Else HayAcompanates = True
On Error GoTo 0
End Function

Private Sub Command8_Click(Index As Integer)
Dim cCant As String

On Error GoTo FixError:

If nPase = 0 Then
    nCantidad = Command8(Index).Index
Else
    cCant = Str(nCantidad)
    cCant = cCant & Command8(Index).Index
    nCantidad = Val(cCant)
    'INFO: 15FEB2011. BAJANDO LA CANTIDAD DE 8 a 4
    'If Len(cCant) = 6 Then
    'INFO: 3OCT2017
    If Len(cCant) = 8 Then
        ShowMsg "CANTIDAD/MONTO NO ES VALIDO, ESTABLECIENDO UNO (1)", vbYellow, vbRed
        nPase = 0
        nCantidad = 1
        Text1(2) = nCantidad
        'Text1(2) = nCantidad / 100
        Exit Sub
    End If
End If

On Error GoTo 0

If HayDecimales Then
    'INFO: 4OCT2017
    Text1(2) = Format(nCantidad / 100, "STANDARD")
Else
    Text1(2) = nCantidad
End If
nPase = nPase + 1
Exit Sub

FixError:
    nCantidad = 1
    Resume Next
End Sub

Private Sub Correccion_Click()
'------------------- CORRECCION / ERROR CORRECT ----------------
Dim rsFixTmpTrans As Recordset
Dim txto As String
Dim sqltext As String
Dim SSD As Single
Dim nTp  As Integer
Dim nn, i As Integer
Dim zz As Integer
Dim SOLO_FECHA As String
Dim nVeriCant As Integer
Dim cStr2Print As String
Dim cSQL As String

nNLinSel = 0: nTp = 0

Set rsFixTmpTrans = New Recordset
'INFO: PARA MULTICAJA
txto = "SELECT * FROM TMP_TRANS "
txto = txto & " WHERE LIN = " & CajLin & " AND MESA = " & nMesa & " AND CAJERO = " & npNumCaj
rsFixTmpTrans.Open txto, msConn, adOpenStatic, adLockReadOnly

If rsFixTmpTrans.EOF = True Then
    rsFixTmpTrans.Close
    Exit Sub
End If

If rsFixTmpTrans!CANT < 0 Then
    ShowMsg "¡ NO puede CORREGIR este Producto !", vbBlue, vbYellow
    rsFixTmpTrans.Close
    Exit Sub
End If

If Mid(rsFixTmpTrans!Tipo, 1, 1) = "B" Then
    ShowMsg "PRODUCTO YA FUE ANULADO/CORREGIDO/SE DIO DESCUENTO EN LA LINEA " & Val(Mid(rsFixTmpTrans!Tipo, 5, 2)), vbBlue, vbYellow
    rsFixTmpTrans.Close
    Exit Sub
End If

'---------------------------------------

nn = 0: i = 1: zz = 0
'Pregunta si hay un Numero en TIPO, si hay significa que tiene Desc
For i = i To 9
    nn = InStr(1, rsFixTmpTrans!Tipo, i)
    If nn <> 0 Then Exit For
Next
If nn <> 0 Then zz = Val(Mid(rsFixTmpTrans!Tipo, nn, 2))
'----------------------------------------

SSD = rsFixTmpTrans!precio * (-1)

SOLO_FECHA = Format(Date, "YYYYMMDD")
CajLin = CajLin + 1

'INSERTA LA LINEA DE CORRECCION
cSQL = "INSERT INTO TMP_TRANS "
cSQL = cSQL & "(CAJA,CAJERO,MESA,MESERO,VALID,LIN,DESCRIP,CANT,DEPTO,PLU,ENVASE,PRECIO_UNIT,PRECIO,FECHA,HORA,TIPO,DESCUENTO,CUENTA,IMPRESO,IMPRESORA, CON_TAX) VALUES ("
cSQL = cSQL & "" & cCaja & "," & npNumCaj & "," & nMesa & "," & nMesero & "," & -1 & "," & CajLin & "," & "'EC-"
cSQL = cSQL & rsFixTmpTrans!DESCRIP & "'" & "," & rsFixTmpTrans!CANT * (-1) & "," & rsFixTmpTrans!depto & "," & rsFixTmpTrans!PLU & ","
cSQL = cSQL & rsFixTmpTrans!envase & "," & rsFixTmpTrans!precio_unit * (-1) & "," & SSD & "," & "'" & SOLO_FECHA & "'" & "," & "'" & Time & "'"
cSQL = cSQL & ",'EC-" & CajLin - 1 & "'," & 0# & "," & rsFixTmpTrans!CUENTA & ",0," & rsFixTmpTrans!IMPRESORA & "," & rsFixTmpTrans!CON_TAX & ")"

'MARCA LA LINEA QUE SE ESTA CORRIGIENDO PARA QUE NO PUEDA SER CORREGIDA/ANULADA
'DE NUEVO
'INFO: MULTICAJA
sqltext = "UPDATE TMP_TRANS SET VALID = 0,TIPO = 'BEC" & Str(CajLin) & "' WHERE LIN = " & (CajLin - 1) & " AND MESA = " & nMesa & " AND CAJERO = " & npNumCaj


'INFO: SE ACTIVA EL 26SEP2017
EscribeLog "CORRECCION. " & RegRead("HKCU\Software\SoloSoftware\SoloMix\LastAuthorization") & ". " & rsFixTmpTrans!DESCRIP & ", por Valor de: " & Format(rsFixTmpTrans!precio, "CURRENCY")

'~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||
'INFO: SOLO IMPRIME LINEA SI ESTA LA 950 PRESENTE
'22OCT2011
'~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||
If cFullFiscal = "SI" Then
Else
    '''''''''''ImprimeLineaDetalle rsFixTmpTrans!cant * (-1), "'EC-" & rsFixTmpTrans!descrip, Format(SSD, "CURRENCY")
    cStr2Print = Format("EC-" & Mid(rsFixTmpTrans!DESCRIP, 1, 14), "!@@@@@@@@@@@@@@@@@") + Space(1) & Format(rsFixTmpTrans!CANT * (-1), "@@@@") & Space(1) & Format(Format(SSD, "####.00"), "@@@@@@@")
    'Sys_Pos.Coptr1.PrintTwoNormal FptrSJournalReceipt, cStr2Print, cStr2Print
    Print2_OPOS_Dev cStr2Print
    
    If Sys_Pos.Coptr1.ResultCode = OposEIllegal Then
        'SI HAY PROBLEMAS DE IMPRESION
        'SE IMPRIME LA INFORMACION MINIMA POSIBLE
        cStr2Print = Format("EC-" & Mid(rsFixTmpTrans!DESCRIP, 1, 8), "!@@@@@@@@@@@") + Space(1) & Format(rsFixTmpTrans!CANT * (-1), "@@@@") & Space(1) & Format(Format(SSD, "####.00"), "@@@@@@@")
        'Sys_Pos.Coptr1.PrintTwoNormal FptrSJournalReceipt, cStr2Print, cStr2Print
        Print2_OPOS_Dev cStr2Print
    End If
End If
'~~~~||||~~~~~~~~||||~~~~~~~~||||~~~~~~~~||||

msConn.BeginTrans
If zz > 0 Then
    'SQLTXT = "UPDATE TMP_TRANS SET TIPO = ' ' WHERE LIN = " & zz & " AND MESA = " & nMesa & " AND CAJERO = " & npNumCaj
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    'msConn.Execute "UPDATE TMP_TRANS SET TIPO = '  ' WHERE MESA = " & nMesa & " AND CAJERO = " & npNumCaj & " AND LIN = " & nLinDescuento
    'UPDATE: 26ABR2024
    SQLTXT = "UPDATE TMP_TRANS SET TIPO = '  ', FE_DESCUENTO = 0 "
    SQLTXT = SQLTXT & " WHERE MESA = " & nMesa
    SQLTXT = SQLTXT & " AND CAJERO = " & npNumCaj
    SQLTXT = SQLTXT & " AND LIN = " & zz
    'msConn.Execute cSQLFix
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    
End If
msConn.Execute cSQL
msConn.Execute sqltext
If zz > 0 Then
    msConn.Execute SQLTXT
End If
msConn.CommitTrans

Call MuestraProductosMarcados

nLineas = PlatosMesa.Rows - 1
PlatosMesa.TopRow = (PlatosMesa.Rows - 1)

rsFixTmpTrans.Close

nCantidad = 1: nPase = 0
Text1(2) = nCantidad

'Unload Acompanantes

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Rem If KeyCode = vbKeyF6 Then Call toJSON("TMP_TRANS")
Rem Call lbNOTax_DblClick
Rem  Call MANT_TMP_TRANS(True)

End Sub


Private Sub Form_Load()
Dim MiTop As Long, MiLeft As Long, StayLeft As Long
Dim iTam As Integer
Dim cSQL As String

Set rs01 = New Recordset
Set rs02 = New Recordset
Set rs03 = New Recordset
Set rs04 = New Recordset
Set rs05 = New Recordset
Set rs06 = New Recordset
Set rs07 = New Recordset
Set rs08 = New Recordset 'Para Precios de PLU con Envases
Set rs09 = New Recordset 'Para Precios de PLU con Envases


'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
'''07OCT2011
If cFullFiscal = "SI" Then
Else
    Call ImprimeTitulos
End If
'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~

'INFO: PREVIENE QUE USEN LA (X) PARA SALIR
cmdSalir.Tag = 100


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'28MAY2014
KDS_Activo = UCase(GetFromINI("Facturacion", "KDS", App.Path & "\soloini.ini"))
If KDS_Activo = "SI" Then
    KDS_Path = UCase(GetFromINI("Facturacion", "KDSDir", App.Path & "\soloini.ini"))
    ChkToGo.Visible = True
End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'----------------------------
'INFO: 24OCT2017
nTransDomicilio = False
'----------------------------

nCantidad = 1:  Text1(2) = nCantidad
num = 0: iTam = 0: CajLin = 0: nPase = 0:
nGlobEnv = 0: nCliNum = 0: nFlag = 0
OkAnul = 0: nNLinSel = 0   'Linea Seleccionada
OKProp = 0: OKDesc = 0: OKCancelar = 0

SetupPantalla

'Show    'Muestra la Pantalla de Facturacion

'Departamentos ordenados segun el menu en administracion
rs02.Open "SELECT codigo, corto, abierto FROM depto ORDER BY ORDEN", msConn, adOpenDynamic, adLockOptimistic
    If rs02.EOF Then
        ShowMsg "NO EXISTEN DEPARTAMENTOS. ES NECESARIO CREAR DEPARTAMENTOS DE VENTAS. EL PROGRAMA TERMINARA AHORA", vbYellow, vbRed
        Unload Me
        End
    End If

'PRODUCTOS del Primer Departamento cargado segun DEPTO!ORDEN

cSQL = "SELECT codigo, depto, descrip, corto, precio1, envases, IMPRESORA, CON_TAX"
cSQL = cSQL & " FROM PLU "
cSQL = cSQL & " WHERE depto = " & rs02!CODIGO
'INFO: 15FEB2011 (FUFO LO ENCONTRO PARA MAQUINA DEMO DE SUSHI EXPRESS
cSQL = cSQL & " AND DISPONIBLE = TRUE "
cSQL = cSQL & " ORDER BY CORTO"

rs03.Open cSQL, msConn, adOpenStatic, adLockOptimistic

'" WHERE depto = " & rs02!codigo & " ORDER BY descrip", msConn, adOpenStatic, adLockOptimistic
'TABLA:Contened_01. ENVASES DISPONIBLES DEL DEPARTAMENTO
rs04.Open "SELECT a.depto,a.contenedor,b.descrip " & " FROM contend_01 as a,contened as b " & " WHERE a.DEPTO = " & rs02!CODIGO & " AND a.contenedor = b.contenedor " & " ORDER BY a.depto,a.contenedor", msConn, adOpenStatic, adLockOptimistic

'Meseros
rs05.Open "SELECT numero,nombre,apellido " & " FROM meseros " & " WHERE NUMERO <> 999 ORDER BY NUMERO", msConn, adOpenStatic, adLockReadOnly
'If rs05.EOF Then
'    MsgBox "NO EXISTEN MESEROS/SALONEROS. ES NECESARIO CREARLOS. EL PROGRAMA TERMINARA AHORA", vbCritical, BoxTit
'    Unload Me
'    End
'End If
'SELECCIONA EL NUMERO DE CAJA
rs06.Open "SELECT caja_cod, descrip FROM cajas WHERE caja_cod = " & cCaja, msConn, adOpenStatic, adLockReadOnly

'SE DEFINE SI ES DEMO O PRODUCTO FINAL
If TipoApplicacion <> "" Then
    Command13(1).Enabled = False
    Correccion.Enabled = False
End If

If ON_LINE Then
    PLU.Caption = PLU.Caption + "." + rs00!DESCRIP + ".ON-LINE" + TipoApplicacion
Else
    PLU.Caption = PLU.Caption + "." + rs00!DESCRIP + ".OFF-LINE" + TipoApplicacion
End If

' Lo maximo que puede caber en el frame de Departamentos son 11
' botones. Indice es entonces 10

'CARGA LOS DATOS EN LOS BOTONES DEPARTAMENTALES DISPONIBLE
'Y SU VALOR EN EL ARREGLO ARREG_DEPTOS
Do Until rs02.EOF
    If num < 1 Then
        cmdDepto(num).Caption = rs02!corto
        Arreg_Deptos(num) = rs02!CODIGO
        ElDepto = rs02!CODIGO
    Else
        Load cmdDepto(num)
        cmdDepto(num).Caption = rs02!corto
        Arreg_Deptos(num) = rs02!CODIGO
        cmdDepto(num).Left = 120
        cmdDepto(num).Top = cmdDepto(num - 1).Top + 660
        cmdDepto(num).Visible = True
    End If
    num = num + 1
    
    If num = nNUM_MAX_DEPTOS Then Exit Do
    rs02.MoveNext
Loop

MiTop = 240: StayLeft = 120
MiLeft = 0: numplu = 0

'Muestra los PLUs(Botones) del primer departamento
For i = 1 To nNUM_MAX_PLU
    Load cmdPlus(i)
Next

Do Until rs03.EOF
    If numplu < 1 Then
        cmdPlus(numplu).Caption = rs03!DESCRIP
        cmdPlus(numplu).Tag = rs03!CODIGO
    Else
        cmdPlus(numplu).Visible = True
        cmdPlus(numplu).Caption = rs03!DESCRIP
        cmdPlus(numplu).Tag = rs03!CODIGO
        cmdPlus(numplu).Left = MiLeft + StayLeft
        cmdPlus(numplu).Top = MiTop
    End If
    numplu = numplu + 1
    MiLeft = MiLeft + 2400
    
    Select Case numplu
        'SE UBICA EN EL SIGUIENTE REGLON
        Case 3, 6, 9, 12, 15, 18, 21, 24, 27
            MiTop = MiTop + 750
            MiLeft = 0
    End Select
    If numplu = nNUM_MAX_PLU Then Exit Do
    rs03.MoveNext
Loop
nMaxPLu = numplu
If nMesa > 0 Then
    'PERMITE MARCAR PLU'S
    Frame2(1).Enabled = True
End If

If HayDecimales Then
    PLU.Label1(3).BackColor = vbCyan
End If

Call GetISC
Call MuestraProductosMarcados

'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

'Set rsPropinaMarcadas = New ADODB.Recordset        'DICIEMBRE DE 2014
'Set rsPagosMarcados = New ADODB.Recordset        'DICIEMBRE DE 2014

'Call PrepareRecordSets
Call ChangeColorOfProgBar
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

End Sub
Private Sub GetISC()

'INFO: UPDATE 24NOV2011
'(A) CHECK YEAR ON ISC
Call CheckUpdateYearOnISC(Year(Date))

On Error Resume Next
rsISC.Open "SELECT ISC FROM ISC", msConn, adOpenStatic, adLockOptimistic
iISC = rsISC!ISC
rsISC.Close
On Error GoTo 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If cmdSalir.Tag = 100 Then
    ShowMsg "¡ Favor utilize el boton SALIR !", vbYellow, vbRed
    Cancel = True
End If
End Sub
Private Sub Image_Click()
Dim cTituloCopia As String
Dim cFile As String


'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

If PlatosMesa.Rows > 0 Then
    ShowMsg "TIENE PRODUCTOS MARCADOS, NO ES EL MOMENTO DE IMPRIMIR COPIA DE FACTURA", vbYellow, vbRed
    Exit Sub
End If

'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
'''02OCT2012
cTituloCopia = UCase(GetFromINI("Fiscal", "TituloCopia", App.Path & "\soloini.ini"))
If cTituloCopia = "" Then cTituloCopia = "COPIA NO/FISCAL ULTIMA TRANS"
cTituloCopia = Left(cTituloCopia, 28)
cFile = DATA_PATH & "sent\LF_NO_F.TXT"
'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~

'''07OCT2011
If cFullFiscal = "SI" Then
    Select Case nMarca
        Case 0      'STAR
            'Call FISCAL_RepiteFactura
            '02OCT2012
            If RepiteFactura(cTituloCopia, cFile) Then GoPrintDOC (cFile)
        Case 1      'BEMATECH
            'Call FISCAL_RepiteFactura
            '02OCT2012
            If RepiteFactura(cTituloCopia, cFile) Then GoPrintDOC (cFile)
        Case Else
    End Select
Else
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    
   If RepiteFactura_FE(cTituloCopia, cFile) Then
        GoPrintDOC (cFile)
        Call ImprimeTitulos
    End If

'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
End If
'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~~~||||'~~
End Sub

Private Sub Imagen_Click(Index As Integer)
On Error GoTo ErrAdm:
Timer1.Enabled = False
Imagen(1).Visible = False
Imagen(2).Visible = False

If PlatosMesa.Rows = 0 Then
    ShowMsg "Cargando el Primer Pedido de Domicilio", vbBlue, vbYellow
Else
    ShowMsg "DEBE FINALIZAR LA ORDEN ACTUAL", vbYellow, vbRed
    Imagen(1).Visible = True
    Imagen(2).Visible = True
    Timer1.Enabled = True
End If
On Error GoTo 0
Exit Sub

ErrAdm:
EscribeLog "Error en Timer.Imagen_Click de PLU"
End Sub

Private Sub Label1_DblClick(Index As Integer)
Dim cLargo As String

    If HayDecimales Then
        'INFO: 11OCT2017
        Text1(2) = Format(InputBox("RECUERDE INTRODUCIR AMBOS DECIMALES. " & vbCrLf & "INTRODUZCA VALOR", "VALOR ...", 1#), "STANDARD")
        cLargo = Text1(2)
        If Len(cLargo) > 8 Then
            ShowMsg "LA CANTIDAD MARCADA NO ES VALIDA, SE ESTABLECERA EL VALOR DE UNO (1)", vbYellow, vbRed
            Text1(2) = 1
        End If
        'Text1(2) = nCantidad
    End If
End Sub

Private Sub lbNOTax_DblClick()

Call MANT_TMP_TRANS

End Sub

Private Sub PlatosMesa_Click()
If PlatosMesa.Rows = 0 Then
    ShowMsg "¡ DEBE MARCAR UN PLATO !", vbBlue, vbYellow
    Exit Sub
End If
nNLinSel = Val(PlatosMesa.Text)
PlatosMesa.Col = 0
End Sub

Private Sub SubTot_Click()
'nMntOculto = Label2.Caption
'lbMonto = Format(nMntOculto, "STANDARD")
Dim nCant As Single
Dim i As Integer

'INFO: 4OCT2017
If HayDecimales Then
    nCant = Format(SubTot.Caption, "STANDARD")
Else
    nCant = Val(Format(SubTot.Caption, "STANDARD"))
End If

If Format(Val(axPagado.Text), "###.00") > 0# Then
    ShowMsg "NO SE PUEDE APLICAR ESTE PAGO" & vbCrLf & vbCrLf & "INTENTE HACERLO UNA VEZ MAS", vbRed, vbYellow
    Call Clear_Click

    Exit Sub
End If
nMontoPagado = nMontoPagado + nCant
For i = 0 To 26
    If cmdPlus(i).BackColor = &HFFFF00 Then
        cmdPlus(i).BackColor = &HC0C0C0
    End If
Next

'INFO: AQUI SIEMPRE VA ENTRAR A cmbBill() como pago en efectivo
cmbBill_Click (99)
End Sub

Private Function CheckUpdateYearOnISC(nAnno As Integer) As Boolean
Dim rsAnnoISC As ADODB.Recordset
Dim nLocalError As Long
Dim cLocalError As String
Dim bUpdateISCYear As Boolean

Set rsAnnoISC = New ADODB.Recordset
On Error GoTo ErrAdm:

CheckUpdateYearOnISC = False

rsAnnoISC.Open "SELECT ISC_YEAR FROM ISC WHERE ISC_YEAR = " & nAnno, msConn, adOpenStatic, adLockOptimistic

If rsAnnoISC.RecordCount = 0 Then
    'INFO: AÑO NO SE ENCUENTRA, HAY QUE AGREGARLO
    bUpdateISCYear = True
End If

rsAnnoISC.Close
Set rsAnnoISC = Nothing

If bUpdateISCYear Then
    'PREPARA LA DATA DE ESTE AÑO
    msConn.BeginTrans
    msConn.Execute "INSERT INTO ISC(ISC_YEAR) VALUES (" & nAnno & ")"
    msConn.CommitTrans
End If

On Error GoTo 0
Exit Function

ErrAdm:
nLocalError = Err.Number
cLocalError = Err.description
Select Case nLocalError
    Case -2147217904
    
        EscribeLog "ERROR AÑO EN ISC, CAMPO (ISC_YEAR) NO ESTA DEFINIDO. " & nLocalError & " - " & cLocalError
        
        msConn.BeginTrans
        msConn.Execute "ALTER TABLE ISC ADD COLUMN ISC_YEAR INTEGER NULL"
        msConn.CommitTrans
        
        'ACTUALIZA LA DATA DEL AÑO ANTERIOR
        msConn.BeginTrans
        msConn.Execute "UPDATE ISC SET ISC_YEAR = " & Year(Date) - 1
        msConn.CommitTrans
        
        'PREPARA LA DATA DE ESTE AÑO
        msConn.BeginTrans
        msConn.Execute "INSERT INTO ISC(ISC_YEAR) VALUES (" & nAnno & ")"
        msConn.CommitTrans
        
        CheckUpdateYearOnISC = True
        
        EscribeLog "Actualización del Campo ISC.ISC_Year esta listo!"
    Case Else
        EscribeLog "ERROR AÑO EN ISC . " & nLocalError & " - " & cLocalError
End Select
'rsAnnoISC.Close
Set rsAnnoISC = Nothing
End Function

Private Function GoPrintDOC(cFile As String) As Boolean
Dim nFreefileTR As Integer
Dim a$

On Error GoTo ErrAdm:

Sys_Pos.Timer1.Enabled = False

nFreefileTR = FreeFile()
Open cFile For Input As #nFreefileTR
Do Until EOF(nFreefileTR)
    Line Input #nFreefileTR, a$
    'Debug.Print a$
    Print2_OPOS_Dev a$, True
Loop
Sleep MUNBYNSeconds
Sys_Pos.Coptr1.CutPaper 100
Close #nFreefileTR

Sys_Pos.Timer1.Enabled = True
On Error GoTo 0
Exit Function

ErrAdm:
    EscribeLog "Error Function GoPrintDOC"
    ShowMsg "Error Function GoPrintDOC", vbYellow, vbRed
    On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Procedure : CheckPedidoDomicilio
' Author    : hsequeira
' Date      : 21/04/2018
' Purpose   : REVISA SI HAY PEDIDOS DE DOMICILIO
'---------------------------------------------------------------------------------------
'
Private Function CheckPedidoDomicilio() As Boolean
nImagen = Val(InputBox("IMAGEN 1 O 2", "TITULO", 1))
cmdDomicilio.Picture = ImageList.ListImages(nImagen)
End Function

Private Sub Timer1_Timer()
On Error GoTo ErrAdm:
If Imagen(1).Visible = True Then
    Imagen(2).Visible = True
    Imagen(1).Visible = False
ElseIf Imagen(2).Visible = True Then
    Imagen(2).Visible = False
    Imagen(1).Visible = True
End If
On Error GoTo 0
Exit Sub

ErrAdm:
 EscribeLog "Error en Timer de PLU"
 On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ImprFactura
' Author    : hsequeira
' Date      : 24/04/2024
' Purpose  : IMPRIME EL FINAL DE LA FACTURA.
'               YA LOS ENCABEZADOS y EL DETALLE ESTAN IMPRESOS
'               ACA SOLO HACE EL CIERRE
'---------------------------------------------------------------------------------------
'
Private Sub ImprFactura_____Falso()
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
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


Private Function PrepareRecordSets()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' DE PAGOS y PROPINAS MARCADAS
' INFO: AGO 2010
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

   On Error GoTo PrepareRecordSets_Error

With rsPropinaMarcadas      'PROPINAS EN EFECTIVO y OTROS
    .Fields.Append "CODIGO", adInteger, , adFldUpdatable
    .Fields.Append "DESCRIP_1", adVarWChar, 60, adFldUpdatable
    .Fields.Append "MONTO_1", adSingle, , adFldUpdatable
    .Fields.Append "MONTO_2", adSingle, , adFldUpdatable
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open
End With

With rsPagosMarcados        'PAGOS RECIBIDOS
    .Fields.Append "CODIGO", adInteger, , adFldUpdatable
    .Fields.Append "DESCRIP_1", adVarWChar, 60, adFldUpdatable
    .Fields.Append "MONTO_1", adSingle, , adFldUpdatable        'lbMonto - nCambio
    .Fields.Append "MONTO_2", adSingle, , adFldUpdatable
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

   On Error GoTo 0
   Exit Function

PrepareRecordSets_Error:

    MsgBox "Error " & Err.Number & " (" & Err.description & ") en PrepareRecordSets de Form PLU"

End Function

Private Function AddPago(cTipo As String, nCodigo As Integer, cDescrip As String, nMonto1 As Single, nMonto2 As Single) As Boolean
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' DE PAGOS y PROPINAS MARCADAS
' INFO: AGO 2010
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

   On Error GoTo AddPago_Error

Select Case cTipo
    Case "PAGO"
        With rsPagosMarcados
            .AddNew
            !CODIGO = nCodigo
            !DESCRIP_1 = cDescrip
            !MONTO_1 = nMonto1
            !MONTO_2 = nMonto2
            
            '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
            '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
            'UPDATE 10SEP2024
            'OBTIENE EL CODIGO DE LA BASE DE DATOS, YA SEA ON LINE O OFF LINE
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

   On Error GoTo 0
   Exit Function

AddPago_Error:

    MsgBox "Error " & Err.Number & " (" & Err.description & ") en AddPago def Form PLU"

End Function

Private Sub RemoveRow_ListaPagos()
Dim i As Byte
On Error GoTo ErrAm:
''i = 1
For i = 1 To ListaPagos.Rows - 1
    ListaPagos.Row = i
    ListaPagos.RemoveItem i
Next
ListaPagos.Clear
On Error GoTo 0
Exit Sub

ErrAm:
'ShowMsg "Sub RemoveRow_ListaPagos" & vbCrLf & Err.Number & " " & Err.description, vbYellow, vbRed
End Sub

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
'cSQL = cSQL & " , VALOR = VALOR + " & Format(Label3, "#0.00")
cSQL = cSQL & " , VALOR = VALOR + " & Format(fSubtotal + fTax, "#0.00")
cSQL = cSQL & " WHERE NUMERO = " & npNumCaj & " OR NUMERO = " & 999
'Pagos.Label3
msConn.Execute cSQL

'AUMENTA E INCREMENTA LOS VALORES POR MESEROS
cSQL = "UPDATE MESEROS SET X_COUNT = X_COUNT + 1"
cSQL = cSQL & " , Z_COUNT = Z_COUNT + 1 "
'cSQL = cSQL & " , VALOR = VALOR + " & Format(Label3, "#0.00")
cSQL = cSQL & " , VALOR = VALOR + " & Format(fSubtotal + fTax, "#0.00")
cSQL = cSQL & " WHERE NUMERO = " & nMesero & " OR NUMERO = " & 999
msConn.Execute cSQL

'cSQL = "UPDATE MESAS SET VALOR = VALOR + " & Format(Pagos.Label3, "#0.00")
cSQL = "UPDATE MESAS SET VALOR = VALOR + " & Format(fSubtotal + fTax, "#0.00")
cSQL = cSQL & ", MESERO_ACTUAL = 0 WHERE NUMERO = " & nMesa & " OR "
cSQL = cSQL & " NUMERO = -99 "
msConn.Execute cSQL

msConn.CommitTrans

msConn.BeginTrans


'INFO: 26JUL2024
'FIX: EMPEZANDO COMO ESTA ANTES CON EL i en 0 en vez de uno
For i = 0 To (ListaPagos.Rows - 1)
'For i = 1 To (ListaPagos.Rows - 1)
    On Error GoTo ErrAdm:
        ListaPagos.Row = i
        ListaPagos.Col = 0
        nTipoPago = ListaPagos.Text
        ListaPagos.Col = 2
    On Error GoTo 0
    nValorPago = Format(ListaPagos.Text, "STANDARD")
    
    cSQL = "UPDATE PAGOS SET X_COUNT = X_COUNT + 1"
    cSQL = cSQL & " , Z_COUNT = Z_COUNT + 1 "
    'cSQL = cSQL & " , VALOR = VALOR + " & Format(nValorPago, "#0.00")
    cSQL = cSQL & " , VALOR = VALOR + " & Format(fSubtotal + fTax, "#0.00")
    
    cSQL = cSQL & " , X_PERIOD_CNT = X_PERIOD_CNT + 1"
    cSQL = cSQL & " , Z_PERIOD_CNT = Z_PERIOD_CNT + 1"
    'cSQL = cSQL & " , PERIOD_VAL = PERIOD_VAL + " & Format(nValorPago, "#0.00")
    cSQL = cSQL & " , PERIOD_VAL = PERIOD_VAL + " & Format(fSubtotal + fTax, "#0.00")
    cSQL = cSQL & " WHERE CODIGO = " & nTipoPago & " OR CODIGO = " & 999
    msConn.Execute cSQL
    
    
    'SI ES UN PAGO COMPUESTO LO HACE AQUI. SI ES SIMPLE LO HACE EN LA PANTLLA DE PLU.
    'INFO: ENE2010 = ARGREGANDO LA MESA A TRANSAC_PAGO
    cSQL = "INSERT INTO TRANSAC_PAGO "
    cSQL = cSQL & " (NUM_TRANS,TIPO_PAGO,CAJERO,LIN,MONTO, MESA) VALUES ("
    cSQL = cSQL & nNUMTrans & "," & nTipoPago & "," & npNumCaj & ","
    'cSQL = cSQL & (i + 1) & "," & Format(nValorPago, "#0.00") & "," & nMesa & ")"
    cSQL = cSQL & (i + 1) & "," & Format(fSubtotal + fTax, "#0.00") & "," & nMesa & ")"
    
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
        'cSQL = cSQL & i + 1 & "," & Format(nValorPago, "#0.00") & ")"
        cSQL = cSQL & i + 1 & "," & Format(fSubtotal + fTax, "#0.00")
        
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

    If FE_Facil_ONLINE Then
        cSQL = "INSERT INTO TRANSAC_FISCAL (FISCAL, DOC_SOLO, FECHA, HORA, SUB_TOTAL, DESCUENTO, ITBM, ID_URL)"
        cSQL = cSQL & " VALUES ("
        cSQL = cSQL & "'" & FE_CUFE & "'," & nNUMTrans & ",'"
        cSQL = cSQL & Format(Date, "YYYYMMDD") & "','" & Format(Time, "HH:MM") & "',"
        cSQL = cSQL & fSubtotal & "," & 0 & "," & (fTax) & ",'" & ArregloPDF_URL(6) & "')"
    Else
        cSQL = "INSERT INTO TRANSAC_FISCAL (FISCAL, DOC_SOLO, FECHA, HORA, SUB_TOTAL, DESCUENTO, ITBM, ID_URL)"
        cSQL = cSQL & " VALUES ("
        cSQL = cSQL & "'" & FE_CUFE & "'," & nNUMTrans & ",'"
        cSQL = cSQL & Format(Date, "YYYYMMDD") & "','" & Format(Time, "HH:MM") & "',"
        cSQL = cSQL & fSubtotal & "," & 0 & "," & (fTax) & ",'" & "OFF-LINE" & "')"
    End If
    
    '
    
    
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
'If nTransDomicilio Then Call PutDomicilioInfo(nNUMTrans, Format(Label3, "#0.00"))
If nTransDomicilio Then Call PutDomicilioInfo(nNUMTrans, Format(fSubtotal + fTax, "#0.00"))

'If nTransDomicilio Then Call UpdateDomiTrans(nNUMTrans, Format(Label3, "#0.00"))
If nTransDomicilio Then Call UpdateDomiTrans(nNUMTrans, Format(fSubtotal + fTax, "#0.00"))

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
    EscribeLog "Error PLU.Actualizador." & ccLERR & Err.Number & " - " & Err.description
    Dim OBJERR As Error
    For Each OBJERR In msConn.Errors
        ShowMsg "Error PLU.Actualizador. " & vbCrLf & "OCURRIO UN ERROR, ANOTE LOS DATOS EN PANTALLA" & vbCrLf & OBJERR.description, vbYellow, vbRed
    Next
    Exit Sub
End If
End Sub

Private Function CierreObjetos() As Boolean
Dim i As Byte

rsPagosMarcados.Close
rsPropinaMarcadas.Close
'Call RemoveRow_ListaPagos

On Error GoTo ErrAm:

ListaPagos.Rows = 0
'
'''i = 1
''INFO: UPDATE 23JUL2024. MOVIENDO i DE 1 A 0
'For i = 0 To ListaPagos.Rows - 1
'    ListaPagos.Row = i
'    ListaPagos.RemoveItem i
'    ListaPagos.RemoveItem
'Next
ListaPagos.Clear
On Error GoTo 0
Exit Function

ErrAm:
'ShowMsg "Sub RemoveRow_ListaPagos" & vbCrLf & Err.Number & " " & Err.description, vbYellow, vbRed
Resume Next
End Function

Private Function MANT_TMP_TRANS(Optional NoPassWord As Boolean) As Boolean
Dim nAfectados As Long
Dim cQuestion As String


'INFO: 13ENE2026
If PlatosMesa.Rows = 0 Then
    ShowMsg "NO HAY NADA MARCADO", vbRed, vbYellow
    Exit Function
End If

If NoPassWord Then
    cQuestion = "SE PONE REGISTROS EN" & vbCrLf & vbCrLf & "<< H O L D >>"
Else
    cQuestion = "¿ DESCARTAR TODO LO MARCADO ?"
End If

If NoPassWord Then
    ShowMsg cQuestion, vbYellow, vbBlue
            
    cSQL = "DELETE * FROM TMP_TRANS "
    cSQL = cSQL & " WHERE CAJA = '" & cCaja & "' AND CAJERO = " & npNumCaj
    cSQL = cSQL & " AND MESA = " & nMesa & " AND MESERO = " & nMesero
    
    msConn.BeginTrans
    msConn.Execute cSQL, nAfectados
    msConn.CommitTrans
    
    EscribeLog "CIERRE DE CUENTA. CON (" & nAfectados & ") PRODUCTOS DESCARTADOS "
    
    Call MuestraProductosMarcados
    
    'INFO: UPDATE 21JUL2024
    axPagado.Text = 0#
    ProgBar = 0
    'INFO: UPDATE 23JUL2024
    ListaPagos.Clear
    ListaPagos.Rows = 0
    'INFO: UPDATE 21JUL2024
    
    nLineas = PlatosMesa.Rows - 1
    'PlatosMesa.TopRow = (PlatosMesa.Rows - 1)
    nCantidad = 1: nPase = 0: nNLinSel = 0: Text1(2) = nCantidad
    
    Call Check_FE_Website
    
    'INFO: UPDATE 21JUL2024
    Call ChangeColorOfProgBar
    
    Rem ShowMsg "CIERRE DE CUENTA" & vbCrLf & "CON (" & nAfectados & ") PRODUCTOS DESCARTADOS ", vbYellow, vbBlue
    
    EscribeLog "Ventas. SE PONEN EN HOLD (" & nAfectados & ") ITEMS."
            
    'INFO:REDUCIENDO EL ESPACIO DE ABAJO DE 10 A 5
    '06JUL2024
    'For i = 1 To 10
    Print2_OPOS_Dev Space(2)
    Sys_Pos.Coptr1.PrintNormal PtrSReceipt, Chr(&H1B) & "!" & Chr(32) & Space(3) & "PEDIDO DESCARTADO" & vbCrLf
    For i = 1 To 6
        Print2_OPOS_Dev Space(2)
    Next
    'DoEvents - DESACTIVADO MAY2014
    
    Sleep MUNBYNSeconds
    Sys_Pos.Coptr1.CutPaper 100
    
    Sleep MUNBYNSeconds
    Call ImprimeTitulos
   
Else
    If ShowMsg(cQuestion, vbYellow, vbBlue, vbYesNo) = vbYes Then
        txtInfo = "DESCARTAR LO MARCADO"
        OkAnul = 0
        AskClave.Show 1
        If OkAnul = 1 Then
            cSQL = "DELETE * FROM TMP_TRANS "
            cSQL = cSQL & " WHERE CAJA = '" & cCaja & "' AND CAJERO = " & npNumCaj
            cSQL = cSQL & " AND MESA = " & nMesa & " AND MESERO = " & nMesero
            
            msConn.BeginTrans
            msConn.Execute cSQL, nAfectados
            msConn.CommitTrans
            
            EscribeLog "CIERRE DE CUENTA. CON (" & nAfectados & ") PRODUCTOS DESCARTADOS "
            
            Call MuestraProductosMarcados
            
            'INFO: UPDATE 21JUL2024
            axPagado.Text = 0#
            ProgBar = 0
            'INFO: UPDATE 23JUL2024
            ListaPagos.Clear
            ListaPagos.Rows = 0
            'INFO: UPDATE 21JUL2024
            
            nLineas = PlatosMesa.Rows - 1
            'PlatosMesa.TopRow = (PlatosMesa.Rows - 1)
            nCantidad = 1: nPase = 0: nNLinSel = 0: Text1(2) = nCantidad
            
            Call Check_FE_Website
            
            'INFO: UPDATE 21JUL2024
            Call ChangeColorOfProgBar
            
            ShowMsg "CIERRE DE CUENTA" & vbCrLf & "CON (" & nAfectados & ") PRODUCTOS DESCARTADOS ", vbYellow, vbBlue
            
            EscribeLog "Ventas. PEDIDO DESCARTADO de (" & nAfectados & ") ITEMS."
            
            'INFO:REDUCIENDO EL ESPACIO DE ABAJO DE 10 A 5
            '06JUL2024
            'For i = 1 To 10
            Print2_OPOS_Dev Space(2)
            Sys_Pos.Coptr1.PrintNormal PtrSReceipt, Chr(&H1B) & "!" & Chr(32) & Space(3) & "PEDIDO DESCARTADO" & vbCrLf
            For i = 1 To 6
                Print2_OPOS_Dev Space(2)
            Next
            'DoEvents - DESACTIVADO MAY2014
            
            Sleep MUNBYNSeconds
            Sys_Pos.Coptr1.CutPaper 100
            
            Sleep MUNBYNSeconds
            Call ImprimeTitulos
            
        Else
            ShowMsg "CODIGO DE AUTORIZACION ES INCORRECTO", vbYellow, vbRed
        End If
    Else
        ShowMsg "DESCARTE DE CUENTA FUE CANCELADO", vbBlue, vbYellow
    End If
End If
End Function


