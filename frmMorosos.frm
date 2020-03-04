VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmMorosos 
   Caption         =   "LISTADO INFORMATIVO DE CLIENTES MOROSOS ..."
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15390
   Icon            =   "frmMorosos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   15390
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   13800
      Picture         =   "frmMorosos.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   855
      Left            =   13800
      Picture         =   "frmMorosos.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid dtgMorosos 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   11033
      _Version        =   393216
      AllowUpdate     =   0   'False
      ForeColor       =   192
      HeadLines       =   1
      RowHeight       =   19
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         AllowRowSizing  =   -1  'True
         AllowSizing     =   -1  'True
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMorosos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rptMorosos As New CrptMorosos
Private rsLisMorosos As New ADODB.Recordset

Private Sub cmdImprimir_Click()

rptMorosos.Database.SetDataSource rsLisMorosos

Set rptGeneral = rptMorosos ' Asigna el reporte al objeto reporte general utilizado
                           ' en el Form de la Vista Previa.

frmVistaPrevia.Show vbModal

Set rptFacturas = Nothing

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()

rsLisMorosos.Open "select cl.Nombre,cl.Apellido, cl.Telefono, ct.fecha, ct.importe from cuentascorrientes ct, clientes cl " & _
                " where ct.idcliente = cl.idcliente and " & _
                " (DateDiff('s', " & Date & ", ct.fecha))> '" & 30 & "' and importe > 0 order by fecha, apellido", cn, adOpenDynamic, adLockOptimistic, adCmdText
                
Set dtgMorosos.DataSource = rsLisMorosos
dtgMorosos.Refresh
 
End Sub

Private Sub Form_Unload(cancel As Integer)
rsLisMorosos.Close
End Sub
