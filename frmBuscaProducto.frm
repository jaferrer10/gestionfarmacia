VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBuscaProducto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda de Productos..."
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7830
   Icon            =   "frmBuscaProducto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid dtgListaPro 
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Hacer Doble Click para seleccionar"
      Top             =   720
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "id"
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
         DataField       =   "troquel"
         Caption         =   "Troquel"
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
      BeginProperty Column02 
         DataField       =   "descripcion"
         Caption         =   "Descripcion"
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
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1275,024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4710,047
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtDescripcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   1
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descripcion:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1320
   End
End
Attribute VB_Name = "frmBuscaProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsListaProductos As New ADODB.Recordset
Private Sub dtgListaPro_DblClick()
vProducto = rsListaProductos!descripcion
VarTroquel = rsListaProductos!troquel
Unload Me
End Sub

Private Sub dtgListaPro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    vProducto = rsListaProductos!descripcion
    VarTroquel = rsListaProductos!troquel
    Unload Me
End If
End Sub

Private Sub Form_Load()
Me.Top = 4150
Me.Left = 2250
txtDescripcion.Text = vProducto
If Len(txtDescripcion.Text) > 0 Then
    SendKeys "{home}+{end}"
End If
If rsListaProductos.State = 0 Then
    rsListaProductos.Open "select * from productos order by descripcion", cn, adOpenDynamic, adLockReadOnly, adCmdText
End If
Set dtgListaPro.DataSource = rsListaProductos
If Len(txtDescripcion.Text) > 0 And rsListaProductos.State = 1 Then
    rsListaProductos.Find "descripcion like '" & txtDescripcion.Text & "%'", , adSearchForward, 1
    If rsListaProductos.EOF = True Then
        rsListaProductos.MoveFirst
    End If
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
rsListaProductos.Close
End Sub
Private Sub txtDescripcion_Change()
If Len(txtDescripcion.Text) > 0 And rsListaProductos.State = 1 Then
    rsListaProductos.Find "descripcion LIKE '" & txtDescripcion.Text & "%'", , adSearchForward, 1
End If
End Sub
Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    dtgListaPro.SetFocus
End If
End Sub
