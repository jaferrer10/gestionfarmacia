VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Begin VB.Form frmProveedores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proveedores ..."
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7680
   Icon            =   "frmProveedores.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid dtgProveedores 
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Doble Click para seleccionar Proveedor"
      Top             =   1080
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   24
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "idproveedor"
         Caption         =   "Codigo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nombre"
         Caption         =   "Nombre Proveedor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
            ColumnWidth     =   989,858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   6134,74
         EndProperty
      EndProperty
   End
   Begin AIFCmp1.asxPowerButton cmdSeleccionar 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      Picture         =   "frmProveedores.frx":0442
      Caption         =   "&Seleccionar Proveedor"
      CaptionAlignment=   5
      CaptionOffsetX  =   -5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureAlignment=   3
   End
   Begin VB.Frame Frame1 
      Caption         =   "Para buscar "
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.TextBox txtNombre 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   50
         TabIndex        =   1
         Top             =   300
         Width           =   7215
      End
   End
   Begin AIFCmp1.asxPowerButton cmdSalir 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   3720
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      Picture         =   "frmProveedores.frx":0894
      Caption         =   "&Cancelar"
      CaptionAlignment=   5
      CaptionOffsetX  =   -10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureAlignment=   3
   End
End
Attribute VB_Name = "frmProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsProv As New ADODB.Recordset
Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdSeleccionar_Click()
VarPro = rsProv!nombre
vIdPro = rsProv!idproveedor
rsProv.Close
Unload Me
End Sub

Private Sub dtgProveedores_DblClick()
VarPro = rsProv!nombre
vIdPro = rsProv!idproveedor
rsProv.Close
Unload Me
End Sub
Private Sub dtgProveedores_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    VarPro = rsProv!nombre
    vIdPro = rsProv!idproveedor
    rsProv.Close
    Unload Me
End If
End Sub
Private Sub Form_Load()
Me.Top = 500
Me.Left = 500
rsProv.Open "select idproveedor,nombre from proveedores order by nombre", cn, adOpenDynamic, adLockReadOnly, adCmdText
Set dtgProveedores.DataSource = rsProv
dtgProveedores.Refresh

End Sub
Private Sub Form_Unload(cancel As Integer)
If rsProv.State = 1 Then
    rsProv.Close
    Set rsProv = Nothing
End If
End Sub

Private Sub txtNombre_Change()
If IsNumeric(txtNombre.Text) = True Then
    strSQL = "idproveedor = " & CStr(txtNombre.Text)
ElseIf Len(txtNombre.Text) > 0 Then
    strSQL = "nombre LIKE '" & txtNombre.Text & "%'"
End If
rsProv.Find strSQL, , adSearchForward, 1
If rsProv.EOF = True Then
    MsgBox "NO COINCIDE CON NINGUN PROVEEDOR, INTENTE NUEVAMENTE !", vbExclamation, "Atencion !"
    txtNombre.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If
End Sub
Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    dtgProveedores.SetFocus
End If
End Sub
Private Sub txtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    VarPro = rsProv!nombre
    vIdPro = rsProv!idproveedor
    rsProv.Close
    Unload Me
End If
End Sub
