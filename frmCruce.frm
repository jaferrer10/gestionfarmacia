VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{20976770-692B-4564-84B5-CCC822AA2B7A}#1.4#0"; "CmdBtnX5.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCruce 
   Caption         =   "Control y cruce de Resúmenes de Compra..."
   ClientHeight    =   9180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15150
   Icon            =   "frmCruce.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15150
   Begin VB.Frame frErrores 
      Height          =   1335
      Left            =   11760
      TabIndex        =   23
      Top             =   2760
      Width           =   3255
      Begin AIFCmp1.asxLabel lblErrores 
         Height          =   330
         Left            =   960
         TabIndex        =   24
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "asxLabel1"
         BorderStyle     =   2
         AutoSize        =   -1  'True
         WordWrap        =   -1  'True
         Alignment       =   2
         UseMnemonic     =   -1  'True
         MouseIcon       =   "frmCruce.frx":0442
      End
      Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
         Height          =   375
         Left            =   0
         Top             =   120
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         EndColor        =   255
         FormatString    =   "Registros con Error"
         Orientation     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   14895
      Begin VB.TextBox txtArchivo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   2520
         MaxLength       =   150
         TabIndex        =   13
         Top             =   1080
         Width           =   6135
      End
      Begin VB.ComboBox cbRubro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         ItemData        =   "frmCruce.frx":075C
         Left            =   11760
         List            =   "frmCruce.frx":0769
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Indica si la factura se debe o fue pagada"
         Top             =   360
         Visible         =   0   'False
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   375
         Left            =   10320
         TabIndex        =   10
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   132710401
         CurrentDate     =   42131
      End
      Begin CommandButtonXCtl.CommandButtonX cmdBuscar 
         Height          =   375
         Left            =   8760
         TabIndex        =   12
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         DropDownPicture =   "frmCruce.frx":0793
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmCruce.frx":0815
      End
      Begin MSDataListLib.DataCombo dtcProveedor 
         Height          =   555
         Left            =   2520
         TabIndex        =   14
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   979
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ForeColor       =   255
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin AIFCmp1.asxPowerButton cmdEjecutar 
         Height          =   495
         Left            =   2520
         TabIndex        =   15
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmCruce.frx":0DAF
         Caption         =   "Ejecutar"
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
         PictureOffsetX  =   10
      End
      Begin AIFCmp1.asxPowerButton cmdCancelar 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   4320
         TabIndex        =   16
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmCruce.frx":17C1
         Caption         =   "&Salir"
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
         PictureOffsetX  =   10
      End
      Begin MSComDlg.CommonDialog dlgLocalizarArchivo 
         Left            =   9480
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   375
         Left            =   12840
         TabIndex        =   21
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   132710401
         CurrentDate     =   42131
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12000
         TabIndex        =   22
         Top             =   1920
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Proveedor:"
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
         TabIndex        =   20
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Archivo de Drogueria:"
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
         TabIndex        =   19
         Top             =   1080
         Width           =   2280
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Rubro:"
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
         Left            =   11040
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Período:"
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
         Left            =   9360
         TabIndex        =   17
         Top             =   1920
         Visible         =   0   'False
         Width           =   885
      End
   End
   Begin VB.Frame frameDatos 
      Caption         =   "Compras que no fueron registradas"
      Height          =   6255
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   11535
      Begin VB.TextBox txtBusCpte 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         MaxLength       =   15
         TabIndex        =   5
         ToolTipText     =   "Ingrese número de Comprobante y luego presione Enter..."
         Top             =   320
         Width           =   3135
      End
      Begin MSDataGridLib.DataGrid dtgArchivo 
         Height          =   5295
         Left            =   120
         TabIndex        =   0
         Top             =   840
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   9340
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16761024
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "fecha"
            Caption         =   "Fecha"
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
            DataField       =   "fechavto"
            Caption         =   "FecVto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1034
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "numero"
            Caption         =   "Comprobante"
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
         BeginProperty Column03 
            DataField       =   "tipo"
            Caption         =   "Tipo"
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
         BeginProperty Column04 
            DataField       =   "importe"
            Caption         =   "Importe"
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
         BeginProperty Column05 
            DataField       =   "rubro"
            Caption         =   "Rubro"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1034
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "depositado"
            Caption         =   "Depositado"
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
         BeginProperty Column07 
            DataField       =   "estado"
            Caption         =   "Estado"
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
         BeginProperty Column08 
            DataField       =   "Observaciones"
            Caption         =   "Observaciones"
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
         BeginProperty Column09 
            DataField       =   "Usuario"
            Caption         =   "Usuario"
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
               ColumnWidth     =   1200,189
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1170,142
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   434,835
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   929,764
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   929,764
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   585,071
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   7590,048
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   884,976
            EndProperty
         EndProperty
      End
      Begin VB.Label LblTotal 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Total"
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
         Left            =   9360
         TabIndex        =   8
         Top             =   360
         Width           =   555
      End
      Begin VB.Label label10 
         AutoSize        =   -1  'True
         Caption         =   "Total de Registros:"
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
         Left            =   7320
         TabIndex        =   7
         Top             =   360
         Width           =   2010
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Buscar Comprobante:"
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
         TabIndex        =   4
         Top             =   360
         Width           =   2250
      End
   End
   Begin AIFCmp1.asxPowerButton cmdBorrar 
      Height          =   495
      Left            =   11880
      TabIndex        =   1
      Top             =   7680
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      Picture         =   "frmCruce.frx":21D3
      Caption         =   "&Borrar"
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
      PictureOffsetX  =   10
   End
   Begin AIFCmp1.asxPowerButton cmdImprimir 
      Height          =   495
      Left            =   11880
      TabIndex        =   2
      Top             =   8400
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      Picture         =   "frmCruce.frx":2BE5
      Caption         =   "&Imprimir"
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
      PictureOffsetX  =   10
   End
   Begin AIFCmp1.asxPowerButton cmdGrabar 
      Height          =   975
      Left            =   11880
      TabIndex        =   6
      Top             =   6480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1720
      Picture         =   "frmCruce.frx":317F
      Caption         =   "Grabar Facturas No Cargadas"
      CaptionAlignment=   5
      CaptionOffsetX  =   -10
      CaptionTextAlignment=   0
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
   Begin AIFCmp1.asxPowerButton cmdMarcaErr 
      Height          =   495
      Left            =   11880
      TabIndex        =   25
      ToolTipText     =   "Renombra el archivo de drogueria, agregando al final del nombre ""_Error"""
      Top             =   5760
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      Picture         =   "frmCruce.frx":5501
      Caption         =   "Marcar archivo con &Errores"
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
      PictureOffsetX  =   10
   End
   Begin AIFCmp1.asxPowerButton cmdMarcaPer 
      Height          =   495
      Left            =   11880
      TabIndex        =   26
      ToolTipText     =   "Renombra archivo de drogueria, agreando al final del nombre ""_ok"""
      Top             =   5040
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      Picture         =   "frmCruce.frx":5DDB
      Caption         =   "Marcar Archivo &Perfecto"
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
      PictureOffsetX  =   10
   End
End
Attribute VB_Name = "frmCruce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsArchivo As New ADODB.Recordset
Private rsProv As New ADODB.Recordset
Private rsCompras As New ADODB.Recordset
Private rsErrores As New ADODB.Recordset
Private vidPro As Integer
Private Carpeta As String
Private Linea As String, Cpte As String, vImp As String, vfec As String, vFecVto As String, vtipo As String, vus As String
Private rptErrores As New CrptErroresCruce
Private vRename As String
'Contador de errores
Private vRegErr As Long

Private Sub cmdBorrar_Click()
SioNo = MsgBox("ESTA SEGURO DE BORRAR EL REGISTRO SELECCIONADO ?", vbExclamation + vbYesNo, "Borrando registro...")

If SioNo = vbYes Then
    rsErrores.Delete
    rsErrores.Update
    dtgArchivo.Refresh
End If

End Sub

Private Sub cmdBuscar_Click()
    With dlgLocalizarArchivo
        .Filter = "Resumen (*.txt)|*.xls" ' Establece el filtro.
        .DialogTitle = "Seleccione la copia de seguridad a restaurar"
        .FileName = "*.txt"
        '.InitDir = App.path
        .DefaultExt = ".txt"
        .Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist + cdlOFNNoChangeDir
        .ShowOpen
        If Len(.FileName) > 0 Then
            If .FileName <> "*.txt" Then
                txtArchivo = .FileName  ' Presentar el nombre del archivo seleccionado.
            End If
        End If
    End With

End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdEjecutar_Click()

Err.Clear

On Error GoTo MuestraError

vRegErr = 0

If Len(txtArchivo.Text) = 0 Then
    MsgBox "DEBE BUSCAR UN ARCHIVO PARA EJECUTAR...!", vbCritical, "FALTA ARCHIVO...."
    cmdBuscar.SetFocus
    Exit Sub
End If

'========= vamos a extraer las fechas inicio y final del periodo del resumen ============
Dim vIni
Dim Vfin

frErrores.Visible = False

Call DefinoRecordsetNuevo

'abro el archivo seleccionado
Open txtArchivo.Text For Input As #1

Do Until EOF(1)
    Line Input #1, Linea
     'Toma la fecha
     vfec = Trim(Mid(Linea, 2, 11))
     'verifica que sea un dato de fecha real para verificar la linea
    If IsNumeric(Mid(vfec, 4, 2)) = True Then
        'grabo registro en el recorset temporal
        rsArchivo.AddNew
        rsArchivo!fecha = Mid(Linea, 2, 11)
        If Mid(Linea, 13, 1) = "A" Then 'Significa que es medicamento y el tamaño del numero de factura es diferente
            rsArchivo!numero = Trim((Mid(Linea, 19, 8)))
            rsArchivo!tipo = Trim(Mid(Linea, 28, 2))
            rsArchivo!importe = Trim(Mid(Linea, 39, 14))
        Else
            rsArchivo!numero = Trim((Mid(Linea, 12, 9)))
            rsArchivo!tipo = Trim(Mid(Linea, 22, 2))
            rsArchivo!importe = Trim(Mid(Linea, 28, 14))
        End If
    End If
Loop

Close #1

rsArchivo.Sort = "fecha"
rsArchivo.MoveFirst
vIni = rsArchivo!fecha
rsArchivo.MoveLast
Vfin = rsArchivo!fecha

'========================================================================================
'Abro tabla de compras según el proveedor seleccionado
If rsCompras.State = 1 Then
    rsCompras.Close
End If
rsCompras.Open "select * from facturascompras where idproveedor = " & dtcProveedor.BoundText & _
                " and fecha >= #" & Format(vIni, "mm/dd/yyyy") & "# and fecha <= #" & Format(Vfin, "mm/dd/yyyy") & _
                "# order by idcompra, fecha desc", cn, adOpenDynamic, adLockOptimistic, adCmdText

'Borro todo el contenido de la tabla de errores del cruce
cn.Execute ("delete from FacturasCruce")

'Abro tabla que almacenará los errores
If rsErrores.State = 1 Then
    rsErrores.Close
End If
rsErrores.Open "select * from FacturasCruce order by Fecha", cn, adOpenDynamic, adLockOptimistic, adCmdText

'abro el archivo seleccionado
Open txtArchivo.Text For Input As #1


'comienza en la linea 13 porque antes son todos titulos
Dim Contador As Integer

Contador = 0

'controlar la primer linea que toma para el proceso
Do Until EOF(1)
    Line Input #1, Linea
     'Toma la fecha
     vfec = Trim(Mid(Linea, 1, 11))
     
     'verifica que sea un dato de fecha real para verificar la linea
     If IsNumeric(Mid(vfec, 4, 2)) = True Then
         If Mid(Linea, 13, 1) = "A" Then 'Significa que es medicamento y el tamaño del numero de factura es diferente
             'Toma el numero de comprobante
            Cpte = LTrim((Mid(Linea, 19, 8)))
            'Tomo el tipo de comprobante
            vtipo = LTrim(Mid(Linea, 28, 2))
            'Toma el importe del comprobante leido
            vImp = Trim(Mid(Linea, 39, 13))
         Else
             'Toma el numero de comprobante
            Cpte = Trim((Mid(Linea, 12, 9)))
            'Tomo el tipo de comprobante
            vtipo = Trim(Mid(Linea, 22, 2))
            'Toma el importe del comprobante leido
            vImp = Trim(Mid(Linea, 28, 14))
         
         End If
         If Mid(vfec, 4, 2) >= 1 And Mid(vfec, 4, 2) <= 12 Then
             Call BuscaCpte
             Contador = Contador + 1
         End If
     
     End If
    
Loop

Close #1

MsgBox "PROCESO TERMINADO CON EXITO !" + Chr(13) & _
        "VERIFIQUE LAS OBSERVACIONES DE CADA REGISTRO...!", vbCritical, "FIN ..."
            
Set dtgArchivo.DataSource = rsErrores
dtgArchivo.Refresh

LblTotal.Caption = Contador

If vRegErr > 0 Then
    frErrores.Visible = True
    lblErrores.Caption = vRegErr
End If

Exit Sub

MuestraError:
    MsgBox Err.Description, vbCritical, "Error en ejecucion: Copie el error y muéstreselo al Programador..."
    Close #1
    
End Sub


Private Sub cmdGrabar_Click()
SioNo = MsgBox("ESTA SEGURO DE GRABAR EN EL SISTEMAS TODAS LAS FACTURAS QUE NO SE ENCUENTRAN CARGADAS ?", vbExclamation + vbYesNo, "GRABAR FACTURAS...")

If SioNo = vbNo Then
    Exit Sub
End If

'graba todas las facturas que tenga la observacion que no se encuentran cargadas en la tabla compra
If rsErrores.RecordCount = 0 Then
    MsgBox "NO HAY DATOS PARA GRABAR !!!", vbCritical, "ATENCION !!!"
    Exit Sub
End If
    
Dim RegGra As Integer

RegGra = 0

rsErrores.MoveFirst
Do While rsErrores.EOF = False
    If rsErrores!observaciones = "FACTURA NO CARGADA !!!" Then
        rsCompras.AddNew
        rsCompras!numero = rsErrores!numero
        rsCompras!tipo = rsErrores!tipo
        rsCompras!fecha = rsErrores!fecha
        rsCompras!fechavto = rsErrores!fecha
        rsCompras!importe = rsErrores!importe
        rsCompras!idproveedor = rsErrores!idproveedor
        rsCompras!rubro = rsErrores!rubro
        rsCompras!usuario = rsErrores!usuario
        rsCompras!Estado = "D"
        rsCompras.Update
        RegGra = RegGra + 1
        rsErrores.MoveNext
    End If
Loop
cn.Execute ("delete * from facturascruce where observaciones = 'FACTURA NO CARGADA !!!'")
rsErrores.Requery
dtgArchivo.Refresh

'Nuevo total de registros que quedan en la grilla
Contador = rsErrores.RecordCount
LblTotal.Caption = Contador

End Sub

Private Sub cmdImprimir_Click()
rptErrores.Database.SetDataSource rsErrores

rptErrores.txtArchivo.SetText txtArchivo.Text
rptErrores.txtProveedor.SetText dtcProveedor.Text

Set rptGeneral = rptErrores ' Asigna el reporte al objeto reporte general utilizado
                           ' en el Form de la Vista Previa.
                           
                           
frmVistaPrevia.Show vbModal

Set rptErrores = Nothing

End Sub


Private Sub cmdMarcaErr_Click()

On Error GoTo ErrorRenombrar

If frErrores.Visible = False Then
    MsgBox "EL PROCESO NO TUVO ERRORES....!", vbCritical, "NO HAY ERRORES..."
    Exit Sub
End If

vRename = ""

If Len(txtArchivo.Text) = 0 Then
    MsgBox "DEBE SELECCIONAR UN ARCHIVO PARA PODER MARCAR...", vbCritical, "ERROR DE DATOS..."
    cmdBuscar.SetFocus
    Exit Sub
End If

SioNo = MsgBox("Esta seguro de Marcar el archivo procesado con Error ??", vbInformation + vbYesNo, "ATENCION !!!")

If SioNo = vbYes Then
    vRename = (Mid(txtArchivo.Text, 1, (Len(txtArchivo.Text) - 4)))
    vRename = vRename + "_Error"
    vRename = vRename + (Mid(txtArchivo.Text, (Len(txtArchivo.Text) - 3), 4))
    Name txtArchivo.Text As vRename
End If

Call LimpiaVentana

Exit Sub

ErrorRenombrar:

MsgBox Err.Description + Chr(13) + "Error de ejecución, imprima pantalla y llame al Programador..."

End Sub

Private Sub cmdMarcaPer_Click()

On Error GoTo ErrorRenombrar

vRename = ""

If Len(txtArchivo.Text) = 0 Then
    MsgBox "DEBE SELECCIONAR UN ARCHIVO PARA PODER MARCAR...", vbCritical, "ERROR DE DATOS..."
    cmdBuscar.SetFocus
    Exit Sub
End If

SioNo = MsgBox("Esta seguro de Marcar el archivo procesado como Perfecto ??", vbInformation + vbYesNo, "ATENCION !!!")

If SioNo = vbYes Then
    vRename = (Mid(txtArchivo.Text, 1, (Len(txtArchivo.Text) - 4)))
    vRename = vRename + "_Ok"
    vRename = vRename + (Mid(txtArchivo.Text, (Len(txtArchivo.Text) - 3), 4))
    Name txtArchivo.Text As vRename
    'limpio ventana
    Call LimpiaVentana
    
End If

'Exit Sub

ErrorRenombrar:

MsgBox Err.Description + Chr(13) + "Error de ejecución, imprima pantalla y llame al Programador..."

End Sub

Private Sub dtcProveedor_LostFocus()
'cuando pierde el foco toma el id del proveedor
If Len(dtcProveedor.Text) > 0 Then
    vidPro = dtcProveedor.BoundText
End If
End Sub
Private Sub dtpDesde_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    dtpHasta.SetFocus
End If
End Sub
Private Sub dtpHasta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdEjecutar.SetFocus
End If
End Sub

Private Sub Form_Load()

Me.Width = 15390
Me.Height = 9750

frErrores.Visible = False

'fechas del periodo toma la fecha del dia
dtpDesde.Value = Date
dtpHasta.Value = Date

If rsProv.State = 1 Then
    rsProv.Close
End If

rsProv.Open "select * from proveedores order by nombre", cn, adOpenDynamic, adLockReadOnly, adCmdText

'llena el combo de proveedores
Set dtcProveedor.DataSource = rsProv
Set dtcProveedor.RowSource = rsProv
dtcProveedor.ListField = "Nombre"
dtcProveedor.BoundColumn = "idproveedor"
dtcProveedor.BoundText = 1


End Sub
Private Sub BuscaCpte()
'busca el comprobante del archivo de texto en la BD
If Len(Cpte) > 0 Then

SQL = "Numero like '%" & Cpte & "%'"

rsCompras.Find (SQL), , adSearchForward, 1

Dim vobs As String
Dim vest As String
Dim vvto As String
Dim vRubro As String
vobs = ""
vest = ""
vvto = ""

If rsCompras.EOF = False Then
    vRubro = rsCompras!rubro
    vus = rsCompras!usuario
    vest = rsCompras!Estado & ""
    vvto = rsCompras!fechavto
    
    Dim vImpCpra As String
    vImpCpra = rsCompras!importe
    
    Dim vImpTxt As String
    vImpTxt = CambiarPunto(CDbl(vImp))
    
    If vImpTxt <> vImpCpra Then
        'si no coiciden los montos graba el registro para auditar en una tabla que se muestra al final en grilla
        vobs = "No coincide los montos"
        vRegErr = vRegErr + 1
    End If
    
    If CDate(vfec) <> rsCompras!fecha Then
        'si no coiciden las fechas graba el registro para auditar en una tabla que se muestra al final en grilla
        vobs = vobs & " + No coinciden las Fechas"
        vRegErr = vRegErr + 1
    End If
    
    If vImpTxt = vImpCpra And CDate(vfec) = rsCompras!fecha Then
        vobs = "PERFECTO"
    End If
    
    
    rsErrores.AddNew
    rsErrores!idproveedor = vidPro
    rsErrores!fecha = vfec
    rsErrores!fechavto = vvto
    rsErrores!numero = Cpte
    rsErrores!tipo = vtipo
    rsErrores!rubro = vRubro
    rsErrores!importe = CDbl(vImpTxt)
    rsErrores!observaciones = vobs
    rsErrores!Estado = vest
    rsErrores!usuario = vUsu 'Esta variable de local
    rsErrores.Update
    
Else
    rsErrores.AddNew
    rsErrores!idproveedor = vidPro
    rsErrores!fecha = vfec
    rsErrores!numero = Cpte
    rsErrores!tipo = vtipo
    'rsErrores!rubro = cbRubro.Text
    rsErrores!importe = CDbl(vImp)
    rsErrores!observaciones = "FACTURA NO CARGADA !!!"
    rsErrores!Estado = vest
    rsErrores!usuario = vUsu 'esta variable es la global
    rsErrores.Update
    
    vRegErr = vRegErr + 1
    
End If
End If

End Sub

Private Sub Form_Unload(cancel As Integer)
If rsErrores.State = 1 Then
    rsErrores.Close
End If
If rsCompras.State = 1 Then
    rsCompras.Close
End If

End Sub


Private Sub txtBusCpte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(txtBusCpte.Text) > 0 Then
        rsErrores.Find ("numero like '%" & txtBusCpte.Text & "%'"), , adSearchForward, 1
        If rsErrores.EOF = True Then
            MsgBox "NO HAY COINCIDENCIAS, COMPROBANTE NO EXISTE...!", vbCritical, "Comprobante inexistente..."
            rsErrores.MoveFirst
        End If
    End If
End If
End Sub

Private Sub DefinoRecordsetNuevo()
  Set rsArchivo = Nothing
  With rsArchivo
    .Fields.Append "fecha", adDate
    .Fields.Append "tipo", adChar, 2
    .Fields.Append "numero", adChar, 10
    .Fields.Append "importe", adChar, 12
    
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .CursorLocation = adUseClient
    .Open
    .Sort = "fecha"
  End With
End Sub

Private Sub LimpiaVentana()

txtArchivo.Text = ""
LblTotal.Caption = "00"

If rsArchivo.State = 1 Then
    rsArchivo.Close
    Set rsArchivo = Nothing
End If
If rsCompras.State = 1 Then
    rsCompras.Close
    Set rsCompras = Nothing
End If
If rsErrores.State = 1 Then
    rsErrores.Close
    Set rsErrores = Nothing
End If
dtgArchivo.Refresh
'txtArchivo.SetFocus
frErrores.Visible = False

End Sub
Public Function CambiarPunto(numero As Variant) As Variant 'cambia el punto por coma
Dim pos As Integer
  If IsNull(numero) Then numero = 0
  pos = InStr(numero, ".")
  If pos = 0 Then
    CambiarPunto = numero
  Else
    CambiarPunto = Mid(numero, 1, pos - 1) & "," & Mid(numero, pos + 1, Len(numero))
  End If

End Function

