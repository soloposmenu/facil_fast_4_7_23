VERSION 5.00
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form ListaClientes 
   BackColor       =   &H00B39665&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LISTADO DE CLIENTES DE DOMICILIO"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12450
   Icon            =   "ListaClientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   12450
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtApellido 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      MaxLength       =   12
      TabIndex        =   6
      Top             =   600
      Width           =   2655
   End
   Begin VB.TextBox txtNombre 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      MaxLength       =   12
      TabIndex        =   4
      Top             =   600
      Width           =   2655
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "REGRESAR"
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
      Left            =   5520
      TabIndex        =   3
      Top             =   5880
      Width           =   1335
   End
   Begin VB.TextBox txtTelefono 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      MaxLength       =   12
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
   Begin DDSharpGridOLEDB2.SGGrid DD_PEDDETALLE 
      Height          =   4575
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   12015
      _cx             =   21193
      _cy             =   8070
      DataMember      =   ""
      DataMode        =   1
      AutoFields      =   -1  'True
      Enabled         =   -1  'True
      GridBorderStyle =   1
      ScrollBars      =   3
      FlatScrollBars  =   0
      ScrollBarTrack  =   0   'False
      DataRowCount    =   2
      BeginProperty HeadingFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataColCount    =   2
      HeadingRowCount =   1
      HeadingColCount =   0
      TextAlignment   =   0
      WordWrap        =   0   'False
      Ellipsis        =   1
      HeadingBackColor=   12632256
      HeadingForeColor=   -2147483630
      HeadingTextAlignment=   0
      HeadingWordWrap =   0   'False
      HeadingEllipsis =   1
      GridLines       =   1
      HeadingGridLines=   2
      GridLinesColor  =   -2147483633
      HeadingGridLinesColor=   -2147483632
      EvenOddStyle    =   1
      ColorEven       =   -2147483628
      ColorOdd        =   14737632
      UserResizeAnimate=   1
      UserResizing    =   3
      RowHeightMin    =   400
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      UserDragging    =   2
      UserHiding      =   0
      CellPadding     =   15
      CellBkgStyle    =   1
      CellBackColor   =   -2147483643
      CellForeColor   =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   2
      FocusRectColor  =   0
      FocusRectLineWidth=   1
      TabKeyBehavior  =   0
      EnterKeyBehavior=   1
      NavigationWrapMode=   1
      SkipReadOnly    =   0   'False
      DefaultColWidth =   1219
      DefaultRowHeight=   255
      CellsBorderColor=   0
      CellsBorderVisible=   -1  'True
      RowNumbering    =   0   'False
      EqualRowHeight  =   0   'False
      EqualColWidth   =   0   'False
      HScrollHeight   =   0
      VScrollWidth    =   0
      Format          =   "General"
      Appearance      =   2
      FitLastColumn   =   0   'False
      SelectionMode   =   2
      MultiSelect     =   0
      AllowAddNew     =   0   'False
      AllowDelete     =   0   'False
      AllowEdit       =   0   'False
      ScrollBarTips   =   0
      CellTips        =   0
      CellTipsDelay   =   1000
      SpecialMode     =   0
      OutlineLines    =   1
      CacheAllRecords =   -1  'True
      ColumnClickSort =   -1  'True
      PreviewPaneColumn=   ""
      PreviewPaneType =   0
      PreviewPanePosition=   2
      PreviewPaneSize =   2000
      GroupIndentation=   241
      InactiveSelection=   1
      AutoScroll      =   -1  'True
      AutoResize      =   1
      AutoResizeHeadings=   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   0
      Caption         =   ""
      ScrollTipColumn =   ""
      MaxRows         =   4194304
      MaxColumns      =   8192
      NewRowPos       =   1
      CustomBkgDraw   =   0
      AutoGroup       =   0   'False
      GroupByBoxVisible=   0   'False
      GroupByBoxText  =   "Arrastre el Titulo de la columna aqui para agrupar por esa columna"
      AlphaBlendEnabled=   0   'False
      DragAlphaLevel  =   206
      AutoSearch      =   0
      AutoSearchDelay =   2000
      OutlineIcons    =   1
      CellTipsDisplayLength=   3000
      StylesCollection=   $"ListaClientes.frx":0442
      ColumnsCollection=   $"ListaClientes.frx":2271
      ValueItems      =   $"ListaClientes.frx":2BE7
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Escriba el APELLIDO"
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
      Left            =   7560
      TabIndex        =   7
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Escriba el NOMBRE"
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
      Left            =   3960
      TabIndex        =   5
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "Escriba el Numero de Telefono"
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
      TabIndex        =   2
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "ListaClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsClientes As ADODB.Recordset
Private cSQL As String
Private nFlagAllowChange As Boolean

Private Sub cmdRegresar_Click()
Dim cPhone As String
Dim cExt As String
Dim WshShell

On Error Resume Next
cPhone = DD_PEDDETALLE.ValueEx(DD_PEDDETALLE.Rows.Current.Key, 1)
cExt = DD_PEDDETALLE.ValueEx(DD_PEDDETALLE.Rows.Current.Key, 2)
On Error GoTo 0

If cPhone = "" Then
    'DO NOTHING
Else
    DomiClientes.txtExt = cExt
    DomiClientes.txtNumber = cPhone
    'INFO: ENVIALE ENTER A TXTNUMBER DE DOMICLIENTES
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.SendKeys "{ENTER}"
    'SendKeys "{ENTER}"

End If

Unload Me
End Sub

Private Sub Form_Load()
'INFO: UPDATE 24SEP2012
Dim cTempTelefono As String
Set rsClientes = New ADODB.Recordset

cTempTelefono = RegRead("HKCU\Software\SoloSoftware\TelefonoDomicilio")

If cTempTelefono = "" Then
Else
    'nFlagAllowChange = False
    txtTelefono.Text = cTempTelefono
    'nFlagAllowChange = True
    RegWrite "HKCU\Software\SoloSoftware\TelefonoDomicilio", ""
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsClientes = Nothing
End Sub

Private Sub txtApellido_Change()

If txtApellido = "" Then
    'NO HACE NADA
Else
    cSQL = "SELECT TELEFONO,EXTENSION, NOMBRE, APELLIDO, DIRECCION1"
    cSQL = cSQL & " FROM CLIENTES "
    cSQL = cSQL & " WHERE APELLIDO LIKE '%" & txtApellido & "%'"
    cSQL = cSQL & " ORDER BY APELLIDO, NOMBRE"
    
    rsClientes.Open cSQL, msConnDomi, adOpenStatic, adLockOptimistic
    
    DD_PEDDETALLE.DataMode = sgBound
    Set DD_PEDDETALLE.DataSource = rsClientes
    
    DD_PEDDETALLE.ReBind
    
    DD_PEDDETALLE.RowHeightMin = 400
    DD_PEDDETALLE.TextAlignment = sgAlignCenterCenter
    
    rsClientes.Close
End If
End Sub

Private Sub txtNombre_Change()

If txtNombre = "" Then
    'NO HACE NADA
Else
    cSQL = "SELECT TELEFONO,EXTENSION, NOMBRE, APELLIDO, DIRECCION1"
    cSQL = cSQL & " FROM CLIENTES "
    cSQL = cSQL & " WHERE NOMBRE LIKE '%" & txtNombre & "%'"
    cSQL = cSQL & " ORDER BY NOMBRE, APELLIDO"
    
    rsClientes.Open cSQL, msConnDomi, adOpenStatic, adLockOptimistic
    
    DD_PEDDETALLE.DataMode = sgBound
    Set DD_PEDDETALLE.DataSource = rsClientes
    
    DD_PEDDETALLE.ReBind
    
    DD_PEDDETALLE.RowHeightMin = 400
    DD_PEDDETALLE.TextAlignment = sgAlignCenterCenter
    
    rsClientes.Close
End If
End Sub

Private Sub txtTelefono_Change()
'Debug.Print "se activo el change:" & Time

'INFO: SI VIENE DE LA OTRA PANTALLA CON INFO, NO PERMITIR CAMBIO
'If Not nFlagAllowChange Then Exit Sub

If txtTelefono = "" Then
    'NO HACE NADA
Else
    cSQL = "SELECT TELEFONO,EXTENSION, NOMBRE, APELLIDO, DIRECCION1"
    cSQL = cSQL & " FROM CLIENTES "
    cSQL = cSQL & " WHERE TELEFONO LIKE '%" & txtTelefono & "%'"
    cSQL = cSQL & " ORDER BY TELEFONO, EXTENSION"
    
    rsClientes.Open cSQL, msConnDomi, adOpenStatic, adLockOptimistic
    
    DD_PEDDETALLE.DataMode = sgBound
    Set DD_PEDDETALLE.DataSource = rsClientes
    
    DD_PEDDETALLE.ReBind
    
    DD_PEDDETALLE.RowHeightMin = 400
    DD_PEDDETALLE.TextAlignment = sgAlignCenterCenter
    
    rsClientes.Close
End If
End Sub
