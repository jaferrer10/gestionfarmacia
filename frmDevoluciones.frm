VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmDevoluciones 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Registro de Devoluciones a Droguerias ..."
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13215
   Icon            =   "frmDevoluciones.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   13215
   Begin VB.TextBox txtImporte 
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
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   36
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Frame frmNc 
      Caption         =   "Registro Nota de Crédito"
      Height          =   1455
      Left            =   6600
      TabIndex        =   28
      Top             =   2280
      Width           =   6375
      Begin MSMask.MaskEdBox mskImpNc 
         Height          =   300
         Left            =   4680
         TabIndex        =   33
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         ClipMode        =   1
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtNc 
         Height          =   300
         Left            =   240
         TabIndex        =   30
         Top             =   480
         Width           =   2295
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx1 
         Height          =   240
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   423
         Caption         =   "N° Nota de Crédito"
      End
      Begin MSComCtl2.DTPicker dtpNc 
         Height          =   375
         Left            =   2760
         TabIndex        =   31
         Top             =   400
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
         Format          =   143654913
         CurrentDate     =   39393
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx2 
         Height          =   240
         Left            =   4680
         TabIndex        =   32
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   423
         Caption         =   "Importe"
      End
      Begin AIFCmp1.asxPowerButton cmdOkNc 
         Height          =   495
         Left            =   4680
         TabIndex        =   34
         ToolTipText     =   "Agrega Proveedor a la lista"
         Top             =   840
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
         BorderStyle     =   4
         Picture         =   "frmDevoluciones.frx":058A
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
      End
      Begin AIFCmp1.asxPowerButton cmdCancelNc 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   5400
         TabIndex        =   35
         ToolTipText     =   "Agrega Proveedor a la lista"
         Top             =   840
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
         BorderStyle     =   4
         Picture         =   "frmDevoluciones.frx":0B24
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
      End
   End
   Begin VB.Frame frArchivo 
      Caption         =   "Archivo de facturas de compras"
      Height          =   4695
      Left            =   6360
      TabIndex        =   24
      Top             =   840
      Width           =   6975
      Begin VB.TextBox txtBusFac 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   0
         ToolTipText     =   "Pruebe ingresar con codigo de barra"
         Top             =   320
         Width           =   2775
      End
      Begin AIFCmp1.asxPowerButton cmdSalefacturas 
         Height          =   495
         Left            =   6120
         TabIndex        =   26
         ToolTipText     =   "Agrega Proveedor a la lista"
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
         BorderStyle     =   4
         Picture         =   "frmDevoluciones.frx":10BE
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
      End
      Begin MSDataGridLib.DataGrid dtgArchivo 
         Height          =   3615
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   6376
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483638
         ForeColor       =   -2147483643
         HeadLines       =   1
         RowHeight       =   15
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
            Size            =   8.25
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin AIFCmp1.asxPowerButton cmdBusFac 
         Height          =   495
         Left            =   5520
         TabIndex        =   37
         ToolTipText     =   "Agrega Proveedor a la lista"
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
         BorderStyle     =   4
         Picture         =   "frmDevoluciones.frx":1658
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
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Busca Factura:"
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
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   1560
      End
   End
   Begin VB.Frame frListo 
      Height          =   735
      Left            =   5760
      TabIndex        =   19
      Top             =   3000
      Width           =   6495
      Begin MSComCtl2.DTPicker dtpDevo 
         Height          =   375
         Left            =   3120
         TabIndex        =   21
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   144441345
         CurrentDate     =   41857
      End
      Begin AIFCmp1.asxPowerButton cmdok 
         Height          =   495
         Left            =   4920
         TabIndex        =   22
         ToolTipText     =   "Agrega Proveedor a la lista"
         Top             =   120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
         BorderStyle     =   4
         Picture         =   "frmDevoluciones.frx":1BF2
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
      End
      Begin AIFCmp1.asxPowerButton cmdCancel 
         Height          =   495
         Left            =   5640
         TabIndex        =   23
         ToolTipText     =   "Agrega Proveedor a la lista"
         Top             =   120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
         BorderStyle     =   4
         Picture         =   "frmDevoluciones.frx":218C
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
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Ingrese Fecha Recibido:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.Frame frDevo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Archivo de Devoluciones"
      Height          =   4575
      Left            =   120
      TabIndex        =   16
      Top             =   3720
      Width           =   12975
      Begin VB.TextBox txtBuscpte 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   39
         ToolTipText     =   "Ingrese Número de comprobante y presione Enter para buscar..."
         Top             =   240
         Width           =   2775
      End
      Begin MSDataGridLib.DataGrid dtgDevo 
         Height          =   3615
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   6376
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   8438015
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
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "iddevolucion"
            Caption         =   "Cod. Dev"
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
         BeginProperty Column02 
            DataField       =   "comprobante"
            Caption         =   "comprobante"
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
            DataField       =   "Tipo"
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
            DataField       =   "Importe"
            Caption         =   "Importe"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "fechafactura"
            Caption         =   "F. Factura"
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
         BeginProperty Column06 
            DataField       =   "fecharegistro"
            Caption         =   "F. Registro"
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
            DataField       =   "fechadev"
            Caption         =   "F. Devoluc"
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
            DataField       =   "fechanc"
            Caption         =   "Fecha NC"
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
            DataField       =   "nombre"
            Caption         =   "Proveedor"
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
         BeginProperty Column10 
            DataField       =   "observaciones"
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
         BeginProperty Column11 
            DataField       =   "nc"
            Caption         =   "Nota de Credito"
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
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1395,213
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   390,047
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1035,213
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1244,976
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1184,882
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1260,284
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1244,976
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   3960
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   4860,284
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1934,929
            EndProperty
         EndProperty
      End
      Begin AIFCmp1.asxPowerButton cmdListo 
         Height          =   495
         Left            =   11160
         TabIndex        =   17
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmDevoluciones.frx":2726
         Caption         =   "&Recibido"
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
      Begin AIFCmp1.asxPowerButton cmdBorrar 
         Height          =   495
         Left            =   11160
         TabIndex        =   18
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmDevoluciones.frx":2CC0
         Caption         =   "&Eliminar"
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
      Begin AIFCmp1.asxPowerButton cmdnc 
         Height          =   495
         Left            =   11160
         TabIndex        =   27
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmDevoluciones.frx":325A
         Caption         =   "&Nota de Créd"
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Buscar Factura:"
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
         Left            =   240
         TabIndex        =   40
         Top             =   360
         Width           =   1635
      End
   End
   Begin VB.TextBox txtNumero 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      MaxLength       =   15
      TabIndex        =   2
      ToolTipText     =   "Pruebe ingresar con codigo de barra"
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox txtObservaciones 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      MaxLength       =   100
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2400
      Width           =   7215
   End
   Begin MSDataListLib.DataCombo cbTipo 
      Height          =   360
      Left            =   5160
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AIFCmp1.asxPowerButton cmdGrabar 
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Picture         =   "frmDevoluciones.frx":37F4
      Caption         =   "&Grabar"
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
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
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
      Format          =   143654913
      CurrentDate     =   39393
   End
   Begin MSDataListLib.DataCombo dtcProveedor 
      Height          =   555
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
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
   Begin AIFCmp1.asxPowerButton cmdCancelar 
      Height          =   495
      Left            =   4080
      TabIndex        =   7
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Picture         =   "frmDevoluciones.frx":4206
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
   Begin AIFCmp1.asxPowerButton cmdAgrPro 
      Height          =   495
      Left            =   5280
      TabIndex        =   15
      ToolTipText     =   "Agrega Proveedor a la lista"
      Top             =   720
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      BackColor       =   12632256
      BorderStyle     =   4
      Picture         =   "frmDevoluciones.frx":4C18
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
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nº Comprobante:"
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
      TabIndex        =   14
      Top             =   840
      Width           =   1785
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Importe:"
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
      TabIndex        =   13
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      TabIndex        =   12
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha:"
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
      TabIndex        =   11
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tipo Cpte:"
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
      Left            =   3960
      TabIndex        =   10
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones:"
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
      TabIndex        =   9
      Top             =   2400
      Width           =   1650
   End
End
Attribute VB_Name = "frmDevoluciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsTipFac As New ADODB.Recordset
Private rsProv As New ADODB.Recordset
Private rsDevo As New ADODB.Recordset
Private rsFacturas As New ADODB.Recordset
Private vIdDev As Integer

Private Sub asxPowerButton2_Click()
frListo.Visible = False
frDevo.Enabled = True
End Sub
Private Sub cbtipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtImporte.SetFocus
End If
End Sub

Private Sub cmdAgrPro_Click()
frArchivo.Visible = True
If rsFacturas.State = 1 Then
    rsFacturas.Close
    Set rsFacturas = Nothing
End If
rsFacturas.Open "select fecha,numero,tipo,importe,observaciones from FacturasCompras where idproveedor = " & dtcProveedor.BoundText & " order by fecha desc", cn, adOpenDynamic, adLockReadOnly, adCmdText
Set dtgArchivo.DataSource = rsFacturas
dtgArchivo.Refresh

txtBusFac.Text = ""
txtBusFac.SetFocus

End Sub

Private Sub cmdBorrar_Click()

If rsDevo.EOF = True Or rsDevo.BOF = True Then
    MsgBox "NO HAY REGISTROS PARA ELIMINAR ...", vbCritical, "NO HAY DATOS..."
    txtNumero.SetFocus
    Exit Sub
End If

SioNo = MsgBox("ESTA SEGURO DE ELMINAR EL REGISTRO SELECCIONADO ?", vbExclamation + vbYesNo, "ELIMINANDO DATOS...")

vIdDev = rsDevo!idDevolucion

If SioNo = vbYes Then
    cn.Execute "delete from devoluciones where idDevolucion = " & vIdDev
    rsDevo.Requery
    dtgDevo.Refresh
End If


End Sub

Private Sub cmdBusFac_Click()
If Len(txtBusFac.Text) > 0 Then
    rsFacturas.Find ("numero like %" & txtBusFac.Text & "%"), , adSearchForward
    If rsFacturas.EOF = True Then
        MsgBox "NO HAY COINCIDENCIAS, EL COMPROBANTE NO SE ENCUENTRA...", vbCritical, "BUSQUEDA SIN EXITO..."
        rsFacturas.MoveFirst
    End If
End If
End Sub

Private Sub cmdCancel_Click()
frListo.Visible = False
txtNumero.SetFocus
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdCancelNc_Click()
txtNc.Text = ""
mskImpNc.Text = 0
frmNc.Visible = False
End Sub

Private Sub cmdGrabar_Click()

If Len(txtNumero.Text) = 0 Then
    MsgBox "DEBE INGRESAR EL NUMERO DEL COMPROBANTE A DEVOLVER...", vbCritical, "FALTAN DATOS..."
    txtNumero.SetFocus
    Exit Sub
End If

If IsNumeric(txtImporte.Text) = False Then
    MsgBox "SOLO SE ADMINTEN DIGITOS Y DOS DECIMALES...!", vbCritical, "DATOS INCORRECTOS...."
    txtImporte.SetFocus
    With txtImporte
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End If

If Len(txtImporte.Text) = 0 Then
    MsgBox "DEBE INGERESAR EL IMPORTE DE LA DEVOLUCION...", vbCritical, "FALTAN DATOS..."
    txtImporte.SetFocus
    Exit Sub
End If

rsDevo.Find ("comprobante like " & Val(Trim(txtNumero.Text))), , adSearchForward, 1
If rsDevo.EOF = False Then
    MsgBox "ESTE COMPROBANTE YA FUE REGISTRADO PARA DEVOLUCION...!", vbCritical, "REGISTRO DUPLICADO ...!"
    txtNumero.SetFocus
    Exit Sub
End If

SioNo = MsgBox("ESTA SEGURO DE GRABAR TODOS LOS DATOS ?", vbExclamation + vbYesNo, "Grabando datos...")

If SioNo = vbYes Then
    rsDevo.AddNew
    rsDevo!idproveedor = dtcProveedor.BoundText
    rsDevo!Estado = "EN ESPERA"
    rsDevo!comprobante = txtNumero.Text
    rsDevo!tipo = cbTipo.BoundText
    rsDevo!importe = Val(txtImporte.Text)
    rsDevo!fechafactura = dtpFecha.Value
    rsDevo!observaciones = txtObservaciones.Text
    rsDevo!fecharegistro = Date
    rsDevo!usuario = vUsu
    rsDevo.Update
    rsDevo.Requery
    dtgDevo.Refresh
    txtNumero.Text = ""
    txtImporte.Text = 0
    txtObservaciones.Text = ""
    txtNumero.SetFocus
End If


End Sub

Private Sub cmdListo_Click()
If rsDevo.EOF Or rsDevo.BOF Then
    MsgBox "NO HAY DATOS DE DEVOLUCIONES ...!!!", vbCritical, "SIN DATOS ..."
    Exit Sub
End If

If rsDevo!Estado = "Resuelto" Then
    MsgBox "ESTE TRAMITE YA HA SIDO RESUELTO...!", vbCritical, "TRAMITE RESUELTO..."
    Exit Sub
End If

If rsDevo!Estado = "ESPERA NC" Then
    MsgBox "YA HA SIDO RECIBIDA LA DEVOLUCION, SE ENCUENTRA EN ESPERA DE NC...", vbCritical, "EN ESPERA DE NC..."
    Exit Sub
End If

Me.dtpDevo.Value = Date
frListo.Visible = True
vIdDev = rsDevo!idDevolucion
dtpDevo.SetFocus

End Sub

Private Sub cmdNc_Click()
If rsDevo!Estado = "Resuelto" Then
    MsgBox "ESTE TRAMITE YA HA SIDO RESUELTO...!", vbCritical, "TRAMITE RESUELTO..."
    Exit Sub
End If
frmNc.Visible = True
dtpNc.Value = Date
txtNc.SetFocus

End Sub

Private Sub cmdOk_Click()
SioNo = MsgBox("ESTA SEGURO DE REGISTRAR LA DEVOLUCION ???", vbExclamation + vbYesNo, "REGISTRANDO DEVOLUCION ...")

If SioNo = vbYes Then
    rsDevo!fechadev = dtpDevo.Value
    rsDevo!Estado = "ESPERA NC"
    rsDevo.Update
    rsDevo.Requery
    frDevo.Refresh
End If
txtNumero.SetFocus
frListo.Visible = False

End Sub

Private Sub cmdOkNc_Click()
SioNo = MsgBox("ESTA SEGURO DE REGISTRAR ESTA NOTA DE CREDITO ?", vbExclamation + vbYesNo, "GRABANDO DATOS...")

If SioNo = vbYes Then
    'buscar nc para no duplicar
    If rsFacturas.State = 1 Then
        rsFacturas.Close
        Set rsFacturas = Nothing
    End If
    
    rsFacturas.Open "select * from FacturasCompras where idproveedor = " & Val(dtcProveedor.BoundText) & " and numero like '" & txtNumero.Text & "' order by Numero", cn, adOpenDynamic, adLockOptimistic, adCmdText
    
    rsFacturas.Find ("numero like " & Trim(txtNc.Text)), , adSearchForward, 1
    If rsFacturas.EOF = False Then
        MsgBox "ESTA NOTA DE CREDITO YA ESTÁ REGISTRADA....!", vbCritical, "DATOS DUPLICADOS..."
        txtNc.SetFocus
        With txtNc
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        Exit Sub
    End If
    
    If Len(txtNc.Text) = 0 Then
        MsgBox "DEBE INGRESAR EL NUMERO DEL COMPROBANTE...!", vbCritical, "FALTAN DATOS..."
        txtNc.SetFocus
        With txtNc
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        Exit Sub
    End If
    
    If Len(mskImpNc.Text) = 0 Then
        MsgBox "DEBE INGRESAR EL IMPORTE DE LA NOTA DE CREDITO ...", vbCritical, "FALTAN DATOS..."
        mskImpNc.SetFocus
        With mskImpNc
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        Exit Sub
    End If
    
    'grabo los datos de la NC en el tabla compras y ya queda registrada y termina el tramite de devolucion
    rsFacturas.AddNew
    rsFacturas!idproveedor = rsDevo!idproveedor
    rsFacturas!fecha = dtpNc.Value
    rsFacturas!numero = txtNc.Text
    rsFacturas!tipo = "NC"
    rsFacturas!importe = Abs(Val(mskImpNc.Text)) * -1
    rsFacturas!rubro = "Medicamentos"
    rsFacturas!observaciones = txtObservaciones.Text
    rsFacturas!Estado = "D"
    rsFacturas!usuario = vUsu
    rsFacturas!fechavto = dtpNc.Value
    rsFacturas.Update
    rsFacturas.Requery

    rsDevo!Estado = "Resuelto"
    'graba la fecha del dia que es cuando se recibe la nc
    rsDevo!fechanc = Date
    rsDevo!nc = txtNc.Text
    
    rsDevo.Update
    rsDevo.Requery
    frDevo.Refresh

End If
txtNc.Text = ""
mskImpNc.Text = 0
frmNc.Visible = False
End Sub

Private Sub cmdSalefacturas_Click()
rsFacturas.Close
Set rsFacturas = Nothing
frArchivo.Visible = False
End Sub

Private Sub dtcProveedor_Change()
If rsDevo.State = 1 Then
    rsDevo.Close
End If

rsDevo.Open "select d.*, p.nombre from devoluciones d inner join proveedores p on d.idproveedor = p.idproveedor " & _
            "where d.idproveedor = " & dtcProveedor.BoundText & " order by fechafactura desc", cn, adOpenDynamic, adLockOptimistic, adCmdText

Set dtgDevo.DataSource = rsDevo
dtgDevo.Refresh

End Sub

Private Sub dtcProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtNumero.SetFocus
End If
End Sub
Private Sub dtgArchivo_DblClick()
frArchivo.Visible = False
txtNumero.Text = rsFacturas!numero
txtImporte.Text = rsFacturas!importe
dtpFecha.Value = rsFacturas!fecha
txtObservaciones.Text = rsFacturas!observaciones
txtImporte.SetFocus
With txtImporte
    .SelStart = 0
    .SelLength = Len(.Text)
End With


'rsFacturas.Close
'Set rsfactura = Nothing

End Sub

Private Sub dtgDevo_DblClick()
If rsDevo.EOF = True Or rsDevo.BOF = True Then
    MsgBox "NO HAY INFORMACION PARA DEVOLUCION...", vbCritical, "NO HAY INFORMACION...."
    txtNumero.SetFocus
    Exit Sub
End If

If rsDevo!Estado = "Resuelto" Then
    MsgBox "ESTE TRAMITE YA HA SIDO RESUELTO...!", vbCritical, "TRAMITE RESUELTO..."
    Exit Sub
End If

If rsDevo!Estado = "ESPERA NC" Then
    MsgBox "YA HA SIDO RECIBIDA LA DEVOLUCION, SE ENCUENTRA EN ESPERA DE NC...", vbCritical, "EN ESPERA DE NC..."
    Exit Sub
End If

frListo.Visible = True
dtpDevo.SetFocus

End Sub

Private Sub dtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cbTipo.SetFocus
End If
End Sub
Private Sub dtpNc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    mskImpNc.SetFocus
End If
End Sub
Private Sub Form_Load()
Me.Width = 13455
Me.Height = 9000

'frames de pantalla
frDevo.Enabled = True
frListo.Visible = False
frArchivo.Visible = False
frmNc.Visible = False
dtpFecha.Value = Date

'Tabla de registro de devoluciones

rsDevo.Open "select d.*, p.nombre from devoluciones d inner join proveedores p on d.idproveedor=p.idproveedor order by fechafactura", cn, adOpenDynamic, adLockOptimistic, adCmdText

Set dtgDevo.DataSource = rsDevo
dtgDevo.Refresh


'llena el combo de proveedores
rsProv.Open "select * from proveedores order by nombre", cn, adOpenDynamic, adLockReadOnly, adCmdText

Set dtcProveedor.DataSource = rsProv
Set dtcProveedor.RowSource = rsProv
dtcProveedor.ListField = "Nombre"
dtcProveedor.BoundColumn = "idproveedor"
dtcProveedor.BoundText = 1


'llena el combo de tipo de factura
rsTipFac.Open "select * from TipoFacturasCpras", cn, adOpenDynamic, adLockReadOnly, adCmdText
Set cbTipo.DataSource = rsTipFac
Set cbTipo.RowSource = rsTipFac
cbTipo.ListField = "Descripcion"
cbTipo.BoundColumn = "idTipo"
cbTipo.BoundText = 1


End Sub

Private Sub Form_Unload(cancel As Integer)
rsTipFac.Close
rsProv.Close
rsDevo.Close
If rsFacturas.State = 1 Then
    rsFacturas.Close
End If

End Sub
Private Sub mskImpNc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdOkNc.SetFocus
End If
ValidarDigitos mskImpNc, KeyAscii

End Sub
Private Sub txtBusCpte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(txtBusCpte.Text) > 0 Then
        rsDevo.Find ("comprobante like %" & txtBusCpte.Text & "%"), , adSearchForward, 1
        If rsDevo.EOF = True Then
            MsgBox "EL COMPROBANTE NO EXISTE !", vbCritical, "Buscando comprobante..."
            Exit Sub
        End If
    End If
End If
End Sub
Private Sub txtBusFac_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBusFac.SetFocus
End If
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtObservaciones.SetFocus
End If
ValidarDigitos txtImporte.Text, KeyAscii
End Sub

Private Sub txtNc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    dtpNc.SetFocus
End If
End Sub
Private Sub txtNumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    dtpFecha.SetFocus
End If
End Sub

Private Sub txtNumero_LostFocus()
If Len(txtNumero.Text) = 0 Then
    Exit Sub
End If
If rsFacturas.State = 1 Then
    rsFacturas.Close
    Set rsFacturas = Nothing
End If

rsFacturas.Open "select idproveedor, numero, fecha, importe, tipo, observaciones from FacturasCompras where idproveedor = " & Val(dtcProveedor.BoundText) & " and numero like '" & txtNumero.Text & "' order by Numero", cn, adOpenDynamic, adLockReadOnly, adCmdText

If rsFacturas.EOF = True Then
    MsgBox "EL COMPROBANTE NO ESTA REGISTRADO EN EL SISTEMA ...!!!", vbCritical, "ERROR ..."
    txtNumero.SetFocus
    With txtNumero
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Exit Sub
End If

txtImporte.Text = rsFacturas!importe
cbTipo.Text = rsFacturas!tipo
txtObservaciones.Text = rsFacturas!observaciones
txtImporte.SetFocus

End Sub

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdGrabar.SetFocus
End If
End Sub

