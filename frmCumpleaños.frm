VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Begin VB.Form frmCumpleaños 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CLIENTES CON FESTEJOS ..."
   ClientHeight    =   6330
   ClientLeft      =   540
   ClientTop       =   735
   ClientWidth     =   9390
   Icon            =   "frmCumpleaños.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9390
   Begin VB.Frame FrameImp 
      Caption         =   "Impresion de Salutaciones"
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   9135
      Begin AIFCmp1.asxPowerButton cmdNoImp 
         Height          =   615
         Left            =   6840
         TabIndex        =   8
         Top             =   1560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         Picture         =   "frmCumpleaños.frx":0442
         Caption         =   "Cancelar"
         CaptionAlignment=   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureAlignment=   0
      End
      Begin AIFCmp1.asxPowerButton cmdImprime 
         Height          =   615
         Left            =   4680
         TabIndex        =   7
         Top             =   1560
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         Picture         =   "frmCumpleaños.frx":059C
         Caption         =   "Imprimir"
         CaptionAlignment=   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureAlignment=   0
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Todos los Clientes de la Lista ..."
         Height          =   495
         Left            =   4320
         TabIndex        =   6
         Top             =   360
         Width           =   3375
      End
      Begin VB.OptionButton opt1 
         Caption         =   "Registro Seleccionado de la Grilla..."
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   480
         Value           =   -1  'True
         Width           =   3255
      End
   End
   Begin AIFCmp1.asxPowerButton cmdSalir 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   6840
      TabIndex        =   3
      Top             =   5520
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      Picture         =   "frmCumpleaños.frx":0B68
      Caption         =   "&Cancelar"
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
      PictureOffsetX  =   10
   End
   Begin AIFCmp1.asxPowerButton cmdSalutacion 
      Height          =   615
      Left            =   4200
      TabIndex        =   2
      Top             =   5520
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      Picture         =   "frmCumpleaños.frx":0CC2
      Caption         =   "&Imprimir Salutación"
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
      PictureOffsetX  =   10
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista de Clientes"
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   9135
      Begin MSDataGridLib.DataGrid dtgLista 
         Height          =   3975
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7011
         _Version        =   393216
         AllowUpdate     =   0   'False
         ForeColor       =   16576
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   735
      Left            =   120
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   1296
      EndColor        =   65535
      FormatString    =   "                 CLIENTES CON FESTEJOS DE CUMPLEAÑOS"
      Orientation     =   0
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCumpleaños"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsCumples As New ADODB.Recordset
Private rsImp As New ADODB.Recordset
Private rptSaludo As New crptTarjeta

Private Sub cmdImprime_Click()
If opt1.Value = True Then
    'imprime solo el registro seleccionado
    rsImp.Open "select * from Clientes where idcliente = " & rsCumples!idcliente, cn, adOpenDynamic, adLockReadOnly, adCmdText
Else
    'imprime todos los clientes de la lista
    rsImp.Open "select * from Clientes order by apellido", cn, adOpenDynamic, adLockReadOnly, adCmdText
End If

rptSaludo.Database.SetDataSource rsImp
Set rptGeneral = rptSaludo ' Asigna el reporte al objeto reporte general utilizado
                           ' en el Form de la Vista Previa.
frmVistaPrevia.Show vbModal

Set rptSaludo = Nothing

rsImp.Close
Set rsImp = Nothing

End Sub
Private Sub cmdNoImp_Click()
FrameImp.Visible = False
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub cmdSalutacion_Click()
FrameImp.Visible = True
End Sub
Private Sub Form_Load()
Me.Top = 500
Me.Left = 3000
'establesco los dias desde y hasta se filtraran las fechas de nacimientos
'para tirar el listado de salutaciones a clientes
Dim DiasDesde As Integer
Dim diasHasta As Integer
If Day(Date) <= 3 Then
    DiasDesde = 1
Else
    DiasDesde = (Day(Date) - 2)
End If
diasHasta = (Day(Date) + 2)

rsCumples.Open "select * from clientes where Day(fechanac) >= " & DiasDesde & " and Day(fechanac) <= " & diasHasta & " and Month(fechanac) = " & Month(Date) & " order by fechanac, apellido", cn, adOpenDynamic, adLockReadOnly, adCmdText

FrameImp.Visible = False
Set dtgLista.DataSource = rsCumples
dtgLista.Refresh

End Sub

Private Sub Form_Unload(cancel As Integer)
rsCumples.Close
Set rsCumples = Nothing
End Sub
