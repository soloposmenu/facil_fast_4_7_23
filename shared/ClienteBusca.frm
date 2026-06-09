VERSION 5.00
Object = "{43135020-B751-4DDD-AB4A-B6D8A342216E}#1.0#0"; "sg20o.ocx"
Begin VB.Form ClienteBuscar 
   BackColor       =   &H00B39665&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BUSQUEDA DE CLIENTES"
   ClientHeight    =   7725
   ClientLeft      =   17865
   ClientTop       =   3315
   ClientWidth     =   11940
   DrawMode        =   14  'Copy Pen
   Icon            =   "ClienteBusca.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   11940
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRUC 
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
      Left            =   7680
      MaxLength       =   12
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "SELECCIONAR CLIENTE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6840
      Width           =   4575
   End
   Begin VB.CommandButton cmdRegresar 
      BackColor       =   &H000000FF&
      Caption         =   "REGRESAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   2775
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
      Left            =   2400
      MaxLength       =   12
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin DDSharpGridOLEDB2.SGGrid DD_PEDDETALLE 
      Height          =   5895
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   11655
      _cx             =   20558
      _cy             =   10398
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
      TextAlignment   =   5
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
      FocusRectColor  =   255
      FocusRectLineWidth=   7
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
      AutoResize      =   0
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
      StylesCollection=   $"ClienteBusca.frx":0442
      ColumnsCollection=   $"ClienteBusca.frx":2271
      ValueItems      =   $"ClienteBusca.frx":2BE7
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "RUC"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Index           =   1
      Left            =   6600
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B39665&
      Caption         =   "NOMBRE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "ClienteBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsBusqueda As New ADODB.Recordset
Dim nCantidadMarcada As Integer
Dim nDepto As Long
Dim nPlu As Long
Dim cItemMarcado As String
Dim MiImpresora As Integer
Dim MiPrecio As Single
Dim MiTax As Single
Dim MiEnvase As Integer

Dim cDescripAcompa As String
Dim nIDAcompa As Integer
Dim rslocalAcompa As New ADODB.Recordset
Dim nAcoBookMark As Variant
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'INFO: MANEJO DE TAMAÑOS (30MAY2023)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim rsLocal As ADODB.Recordset
Dim rsClientesFE As ADODB.Recordset
Dim nContenedor As Long
Dim cContenedor As String


Private Sub cmdRegresar_Click()
nIDClienteBusqueda = 0
ShowMsg "SE REGRESA SIN AGREGAR CLIENTE", vbBlue, vbYellow
Unload Me
End Sub

Private Sub Command1_Click()
    On Error GoTo ErrAdm:
    ShowMsg "SE ASIGNA CLIENTE " & vbCrLf & vbCrLf & DD_PEDDETALLE.Rows.Current.Cells(2).value, vbYellow, vbBlue
    Unload Me
    On Error GoTo 0
    
ErrAdm:
    On Error GoTo 0
    Unload Me
End Sub

Private Sub DD_PEDDETALLE_Click()

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   On Error GoTo DD_PEDDETALLE_Click_Error

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

nIDClienteBusqueda = DD_PEDDETALLE.Rows.Current.Cells(0).value
   
   On Error GoTo 0
   Exit Sub

DD_PEDDETALLE_Click_Error:
    If Err.Number = -2147024809 Then
        ShowMsg "PRIMERO DEBE BUSCAR UN CLIENTE", vbYellow, vbRed
    Else
        ShowMsg "Error " & Err.Number & " (" & Err.description & ") in ClienteBuscar"
    End If
    txtNombre.SetFocus
End Sub

Private Sub DD_PEDDETALLE_OnInit()
DD_PEDDETALLE.VScrollWidth = DD_PEDDETALLE.ClientHeight / 10
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : hsequeira
' Date      : 11/05/2023
' Purpose   : CIERRA EL ITEM DE BUSQUEDA
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()


Set rsClientesFE = New ADODB.Recordset
cSQL = "SELECT ID, TIPO_CLIENTE, CONTRIBUYENTE, NOMBRE AS ESCONDIDO, CEDULA_RUC, "
cSQL = cSQL & " DV, RAZON_SOCIAL_NOMBRE, DIRECCION, "
cSQL = cSQL & " PAIS, PROVINCIA, DISTRITO, CORREGIMIENTO, "
cSQL = cSQL & " EMAIL, TELEFONO, ID_UBICACION, "
'INFO: 10MAR2024
cSQL = cSQL & " PAIS_COMPLETO "
cSQL = cSQL & " FROM CLIENTE_FE"
cSQL = cSQL & " ORDER BY NOMBRE  ASC "

rsClientesFE.Open cSQL, msConn, adOpenStatic, adLockOptimistic

Set rsLocal = New ADODB.Recordset
Set rsLocal = rsClientesFE.Clone(adLockOptimistic)

cItemMarcado = ""
AddEnBusqueda = False
End Sub
Private Sub txtNombre_Change()

On Error GoTo txtNombre_Change_Error

If txtNombre = "" Then
    'NO HACE NADA
    cItemMarcado = ""
    Set DD_PEDDETALLE.DataSource = Nothing
   
    DD_PEDDETALLE.ReBind
    
Else
    cSQL = "SELECT A.ID, A.CEDULA_RUC, A.RAZON_SOCIAL_NOMBRE  "
    cSQL = cSQL & " FROM CLIENTE_FE AS A "
    cSQL = cSQL & " WHERE A.RAZON_SOCIAL_NOMBRE LIKE '%" & txtNombre & "%' "
    cSQL = cSQL & " ORDER BY 3"
    
    rsLocal.Filter = "RAZON_SOCIAL_NOMBRE LIKE '%" & txtNombre & "%' "
    
    DD_PEDDETALLE.DataMode = sgUnbound
    DD_PEDDETALLE.DataRowCount = 0
    DD_PEDDETALLE.DataColCount = 11
    rsLocal.MoveFirst
    Do While Not rsLocal.EOF
        cData = rsLocal!id & "|" & rsLocal!CEDULA_RUC & "|" & _
                    rsLocal!RAZON_SOCIAL_NOMBRE & "|" & String(4, Chr(126))
        DD_PEDDETALLE.Rows.Add sgFormatCharSeparatedValue, cData, "|"
        rsLocal.MoveNext
    Loop
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    cItemMarcado = ""
    
    DD_PEDDETALLE.Columns(1).Caption = "ID"
    DD_PEDDETALLE.Columns(2).Caption = "CEDULA_RUC"
    DD_PEDDETALLE.Columns(3).Caption = "RAZON_SOCIAL_NOMBRE"
    
    
     On Error Resume Next
     With DD_PEDDETALLE
        .ColumnClickSort = False
        .EvenOddStyle = sgEvenOddRows
        .ColorEven = vbWhite
        .ColorOdd = &HE0E0E0
        .Columns(1).Width = 1200:        'ID
        .Columns(1).Style.WordWrap = True
        .Columns(2).Width = 2800:        'CEDULA RUC
        .Columns(2).Style.WordWrap = True
        .Columns(3).Width = 5000:     'RAZON_SOCIAL_NOMBRE

    End With
    On Error GoTo 0
    
    
    DD_PEDDETALLE.RowHeightMin = 650
    DD_PEDDETALLE.TextAlignment = sgAlignCenterCenter
    
End If

   On Error GoTo 0
   Exit Sub

txtNombre_Change_Error:
    If txtNombre.Text = "" Then
          DD_PEDDETALLE.DataRowCount = 0
    Else
        If rsLocal.EOF Then
              DD_PEDDETALLE.DataRowCount = 0
        Else
            ShowMsg "Error " & Err.Number & " (" & Err.description & ") in ClienteBuscar", vbYellow, vbRed
        End If
    End If

End Sub

Private Sub txtNombre_GotFocus()
txtNombre.BackColor = vbYellow
txtRUC.BackColor = vbWhite
End Sub

Private Sub txtRUC_Change()

On Error GoTo txtRUC_Change_Error

If txtRUC = "" Then
    'NO HACE NADA
    cItemMarcado = ""
    Set DD_PEDDETALLE.DataSource = Nothing
   
    DD_PEDDETALLE.ReBind
    
Else
    cSQL = "SELECT A.ID, A.CEDULA_RUC, A.RAZON_SOCIAL_NOMBRE  "
    cSQL = cSQL & " FROM CLIENTE_FE AS A "
    cSQL = cSQL & " WHERE A.RAZON_SOCIAL_NOMBRE LIKE '%" & txtRUC & "%' "
    cSQL = cSQL & " ORDER BY 3"
    
    rsLocal.Filter = "CEDULA_RUC LIKE '%" & txtRUC & "%' "
    
    DD_PEDDETALLE.DataMode = sgUnbound
    DD_PEDDETALLE.DataRowCount = 0
    DD_PEDDETALLE.DataColCount = 11
    rsLocal.MoveFirst
    Do While Not rsLocal.EOF
        cData = rsLocal!id & "|" & rsLocal!CEDULA_RUC & "|" & _
                    rsLocal!RAZON_SOCIAL_NOMBRE & "|" & String(4, Chr(126))
        DD_PEDDETALLE.Rows.Add sgFormatCharSeparatedValue, cData, "|"
        rsLocal.MoveNext
    Loop
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    cItemMarcado = ""
    
    DD_PEDDETALLE.Columns(1).Caption = "ID"
    DD_PEDDETALLE.Columns(2).Caption = "CEDULA_RUC"
    DD_PEDDETALLE.Columns(3).Caption = "RAZON_SOCIAL_NOMBRE"
    
    
     On Error Resume Next
     With DD_PEDDETALLE
        .ColumnClickSort = False
        .EvenOddStyle = sgEvenOddRows
        .ColorEven = vbWhite
        .ColorOdd = &HE0E0E0
        .Columns(1).Width = 1200:        'ID
        .Columns(1).Style.WordWrap = True
        .Columns(2).Width = 2800:        'CEDULA RUC
        .Columns(2).Style.WordWrap = True
        .Columns(3).Width = 5000:     'RAZON_SOCIAL_NOMBRE

    End With
    On Error GoTo 0
    
    
    DD_PEDDETALLE.RowHeightMin = 650
    DD_PEDDETALLE.TextAlignment = sgAlignCenterCenter
    
End If

   On Error GoTo 0
   Exit Sub

txtRUC_Change_Error:
    If txtRUC.Text = "" Then
          DD_PEDDETALLE.DataRowCount = 0
    Else
        If rsLocal.EOF Then
              DD_PEDDETALLE.DataRowCount = 0
        Else
            ShowMsg "Error " & Err.Number & " (" & Err.description & ") in ClienteBuscar", vbYellow, vbRed
        End If
    End If

End Sub

Private Sub txtRUC_GotFocus()
txtNombre.BackColor = vbWhite
txtRUC.BackColor = vbYellow
End Sub
