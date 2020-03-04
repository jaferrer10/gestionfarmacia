VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Begin VB.Form frmFacturacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturacion ..."
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15660
   Icon            =   "frmFacturacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8850
   ScaleWidth      =   15660
   Begin MSDataGridLib.DataGrid dtgProductos 
      Height          =   5535
      Left            =   1440
      TabIndex        =   17
      Top             =   3000
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9763
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   8421631
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
      ColumnCount     =   4
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
         DataField       =   "precio"
         Caption         =   "Precio"
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
         DataField       =   "fabricante"
         Caption         =   "Fabricante"
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
            Locked          =   -1  'True
            ColumnWidth     =   1349,858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4965,166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1950,236
         EndProperty
      EndProperty
   End
   Begin VB.Frame frameVenta 
      Caption         =   "Venta"
      Height          =   5535
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   15375
      Begin AIFCmp1.asxToolButton cmdDesc 
         Height          =   495
         Left            =   1920
         Top             =   3960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Picture         =   "frmFacturacion.frx":030A
         Caption         =   "&Desc (F3)"
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
         PictureAlignment=   3
      End
      Begin AIFCmp1.asxToolButton cmdFacturar 
         Height          =   495
         Left            =   240
         Top             =   3960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Picture         =   "frmFacturacion.frx":0624
         Caption         =   "&Facturar (F10)"
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
         PictureAlignment=   3
      End
      Begin AIFCmp1.asxPowerBanner lblTotVta 
         Height          =   375
         Left            =   13200
         Top             =   3960
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         EndColor        =   4259584
         FormatString    =   ""
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
      Begin MSDataGridLib.DataGrid dtgLista 
         Height          =   3495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   6165
         _Version        =   393216
         BackColor       =   16777088
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "producto"
            Caption         =   "Descripcion Producto"
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
            DataField       =   "cantidad"
            Caption         =   "Cantidad"
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
            DataField       =   "precio"
            Caption         =   "Precio"
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
            DataField       =   "descuento"
            Caption         =   "Desc%"
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
            DataField       =   "condicioniva"
            Caption         =   "CondicionIVA"
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
            DataField       =   "fabricante"
            Caption         =   "Fabricante"
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
         BeginProperty Column08 
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
               ColumnWidth     =   4860,284
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   840,189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   750,047
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   989,858
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1110,047
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
         EndProperty
      End
      Begin AIFCmp1.asxPowerBanner lbltotalDesc 
         Height          =   375
         Left            =   13200
         Top             =   4440
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         FormatString    =   ""
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
      Begin AIFCmp1.asxToolButton cmdAnular 
         Height          =   495
         Left            =   240
         Top             =   4680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Picture         =   "frmFacturacion.frx":077E
         Caption         =   "&Anular (F9)"
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
         PictureAlignment=   3
      End
      Begin AIFCmp1.asxToolButton cmdReimp 
         Height          =   495
         Left            =   1920
         Top             =   4680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Picture         =   "frmFacturacion.frx":0A98
         Caption         =   "&Reimpri (F4)"
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
         PictureAlignment=   3
      End
      Begin AIFCmp1.asxPowerBanner lblCobrar 
         Height          =   375
         Left            =   13200
         Top             =   4920
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         EndColor        =   255
         FormatString    =   ""
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
      Begin VB.Label Label5 
         Caption         =   "Importe a Cobrar:"
         Height          =   195
         Left            =   11640
         TabIndex        =   20
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total c/Descuento:"
         Height          =   195
         Left            =   11520
         TabIndex        =   14
         Top             =   4560
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Importe Venta:"
         Height          =   195
         Left            =   11520
         TabIndex        =   13
         Top             =   4080
         Width           =   1440
      End
   End
   Begin VB.Frame frameCon 
      Caption         =   "Condiciones"
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   15375
      Begin VB.TextBox txtCantidad 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   12360
         MaxLength       =   10
         TabIndex        =   2
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtProducto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   1
         Top             =   960
         Width           =   9615
      End
      Begin VB.ComboBox cbEntrega 
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
         ItemData        =   "frmFacturacion.frx":0BF2
         Left            =   3120
         List            =   "frmFacturacion.frx":0C02
         TabIndex        =   12
         Text            =   "Mostrador"
         Top             =   480
         Width           =   3135
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx4 
         Height          =   240
         Left            =   3120
         TabIndex        =   11
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   423
         Caption         =   "Forma Entrega"
      End
      Begin VB.ComboBox cbPago 
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
         ItemData        =   "frmFacturacion.frx":0C34
         Left            =   120
         List            =   "frmFacturacion.frx":0C44
         TabIndex        =   10
         Text            =   "Contado Efectivo"
         Top             =   480
         Width           =   2775
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx3 
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   423
         Caption         =   "Forma de Pago"
      End
      Begin AIFCmp1.asxToolButton cmdNo 
         Height          =   495
         Left            =   14520
         Top             =   840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         Picture         =   "frmFacturacion.frx":0C75
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
      Begin AIFCmp1.asxPowerButton cmdOk 
         Height          =   495
         Left            =   13680
         TabIndex        =   3
         Top             =   840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         Picture         =   "frmFacturacion.frx":1687
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cantidad:"
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
         Height          =   240
         Left            =   11160
         TabIndex        =   16
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Producto:"
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
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1005
      End
   End
   Begin VB.Frame frameDatos 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15375
      Begin MSDataListLib.DataCombo dtcTipo 
         Height          =   360
         Left            =   9000
         TabIndex        =   19
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         _Version        =   393216
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
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx5 
         Height          =   240
         Left            =   9000
         TabIndex        =   18
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   423
         Caption         =   "Tipo Cpte."
      End
      Begin VB.TextBox txtCliente 
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
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   8775
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx1 
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   423
         Caption         =   "Cliente"
      End
      Begin AIFCmp1.asxPowerBanner lblDireccion 
         Height          =   375
         Left            =   120
         Top             =   1005
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   661
         FormatString    =   "Direccion"
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
      Begin AIFCmp1.asxPowerBanner lblTelefono 
         Height          =   375
         Left            =   9000
         Top             =   1005
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   661
         FormatString    =   "Telefono"
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
      Begin AIFCmp1.asxToolButton cmdSalir 
         Height          =   495
         Left            =   12720
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmFacturacion.frx":2099
         Caption         =   "&Salir"
         CaptionAlignment=   5
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
      Begin AIFCmp1.asxToolButton cmdClientes 
         Height          =   495
         Left            =   10920
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmFacturacion.frx":2625
         Caption         =   "&Clientes"
         CaptionAlignment=   5
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
End
Attribute VB_Name = "frmFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsFacturador As New ADODB.Recordset
Private rsFacturacion As New ADODB.Recordset
Private rsProductos As New ADODB.Recordset
Private rsTipoCte As New ADODB.Recordset
Private rsClientes As New ADODB.Recordset
Private vNeto As Double
Private vDesc As Double
Private vBruto As Double

Private Sub cmdClientes_Click()
frmBuscaClientes.Show
rsClientes.Find "idcliente = " & vIdCliente, , adSearchForward, 1
If rsClientes.EOF = False Then
    txtCliente.Text = rsClientes!nombre + " " + rsClientes!apellido
    lblDireccion.FormatString = rsClientes!direccion
    lblTelefono.FormatString = rsClientes!telefono
End If
End Sub
Private Sub cmdDesc_Click()
Call Descuento
rsFacturador.Requery
Me.dtgLista.Refresh
End Sub
Private Sub cmdFacturar_Click()
Call FacturaVenta
End Sub
Private Sub cmdNo_Click()
txtProducto.Text = ""
txtProducto.SetFocus
End Sub
Private Sub cmdOk_Click()
    
KeyAscii = 0
If Len(txtProducto.Text) = 0 Then
    MsgBox "DEBE INGRESAR EL PRODUCTO A VENDER !", vbCritical, "ATENCION !"
    txtProducto.SetFocus
    Exit Sub
End If
If Len(txtCantidad.Text) = 0 Then
    MsgBox "DEBE INGRESAR LA CANTIDAD VENDIDA...!", vbCritical, "ATENCION !"
    txtCantidad.SetFocus
    Exit Sub
End If

rsFacturador.AddNew
rsFacturador!tipo = dtcTipo.BoundText
rsFacturador!fecha = Format(Date, "dd/mm/yyyy")
rsFacturador!Hora = Format(Time, "hh:mm")
rsFacturador!producto = txtProducto.Text
rsFacturador!precio = rsProductos!precio
rsFacturador!cantidad = txtCantidad.Text
rsFacturador!Descuento = 0
rsFacturador!importe = (rsProductos!precio * txtCantidad.Text)
rsFacturador!fabricante = rsProductos!fabricante
rsFacturador!troquel = rsProductos!troquel
rsFacturador!cliente = txtCliente.Text
rsFacturador.Update
dtgLista.Refresh
txtProducto.Text = ""
txtProducto.SetFocus

Call CalculaImportes

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub dtgLista_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If rsFacturador.RecordCount > 0 Then
        rsFacturador.Delete
        dtgLista.Refresh
        Call CalculaImportes
    End If
    txtProducto.SetFocus
End If
End Sub
Private Sub dtgProductos_Click()
    txtProducto.Text = rsProductos!descripcion
    txtCantidad.Text = 1
    SendKeys "{home}+{end}"
    dtgProductos.Visible = False
End Sub

Private Sub dtgProductos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtProducto.Text = rsProductos!descripcion
    txtCantidad.Text = 1
    txtCantidad.SetFocus
    SendKeys "{home}+{end}"
    
End If
If KeyAscii = 27 Then
    txtProducto.SetFocus
    SendKeys "{end}+{home}"
End If
dtgProductos.Visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF10 Then
    Call FacturaVenta
End If
If KeyCode = vbKeyF3 Then
    Call Descuento
    rsFacturador.Requery
    rsFacturador.Resync
    dtgProductos.Refresh
End If

If KeyCode = vbKeyF9 Then
    cn.Execute "delete from facturador"
    rsFacturador.Requery
    Set dtgLista.DataSource = rsFacturador
    dtgLista.Refresh
    lblTotVta.FormatString = "0.00"
    lbltotalDesc.FormatString = "0.00"
    lblCobrar.FormatString = "0.00"
End If

End Sub

Private Sub Form_Load()
Me.Top = 20
Me.Left = 1500

Me.KeyPreview = True 'para que el formulario procese primero las pulsaciones de teclas
dtgProductos.Visible = False

'limpia tabla del facturador
cn.Execute "delete from facturador"
rsFacturador.Open "select * from facturador", cn, adOpenDynamic, adLockOptimistic, adCmdText
Set dtgLista.DataSource = rsFacturador
lblTotVta.FormatString = "0.00"
lbltotalDesc.FormatString = "0.00"
lblCobrar.FormatString = "0.00"

'toma los datos de tipo de comprobantes
rsTipoCte.Open "select * from TipoCpte", cn, adOpenKeyset, adLockReadOnly, adCmdText
Set dtcTipo.DataSource = rsTipoCte
Set dtcTipo.RowSource = rsTipoCte
dtcTipo.ListField = "Tipo"
dtcTipo.BoundColumn = "idtipo"
dtcTipo.BoundText = 1

'ABRE TABLA CLIENTES
rsClientes.Open "select * from clientes order by apellido", cn, adOpenDynamic, adLockReadOnly, adCmdText
rsClientes.Find "nombre like " & "'CONSUMIDOR%'", , adSearchForward, 1
If rsClientes.EOF = False Then
    txtCliente.Text = rsClientes!nombre + " " + rsClientes!apellido
    lblDireccion.FormatString = rsClientes!direccion
    lblTelefono.FormatString = rsClientes!telefono
    vIdCliente = rsClientes!idcliente 'variable global
End If
End Sub
Private Sub FacturaVenta()
SioNo = MsgBox("IMPRIME EL COMPROBANTE DE VENTA ?", vbInformation + vbYesNo, "ATENCION !")
If SioNo = vbYes Then


End If
End Sub
Private Sub Form_Unload(cancel As Integer)
rsFacturador.Close
rsTipoCte.Close
rsClientes.Close
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdok.SetFocus
End If
If KeyAscii = 27 Then
    txtProducto.Text = ""
    txtProducto.SetFocus
End If
End Sub
Private Sub txtProducto_KeyDown(KeyCode As Integer, Shift As Integer)
If dtgProductos.Visible = False Then
    Exit Sub
End If
If KeyCode = 40 Then
    dtgProductos.SetFocus
End If
End Sub
Private Sub txtProducto_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 27 Then
    dtgProductos.Visible = False
    txtProducto.Text = ""
    txtProducto.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If
If KeyAscii = 13 Then
    KeyAscii = 0
    If rsProductos.State = 1 Then
        rsProductos.Close
        Set rsProductos = Nothing
    End If
    If IsNumeric(txtProducto.Text) = True Then
        'busca por troquel si es numerico
        rsProductos.Open "select troquel,descripcion,precio,fabricante from productos where " & _
                     "troquel ='" & txtProducto.Text & "' order by descripcion", cn, adOpenDynamic, adLockReadOnly, adCmdText
        If rsProductos.EOF = True Then
            MsgBox "EL PRODUCTO NO EXISTE ...!", vbCritical, "ATENCION !"
            txtProducto.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        Else
            txtProducto.Text = rsProductos!descripcion
            txtCantidad.Text = 1
            txtCantidad.SetFocus
            SendKeys "{home}+{end}"
        End If
    Else
        Dim Cadena As String
        Cadena = Replace(txtProducto.Text, " ", "%", 1)
        'busca por la descripcion con la grilla
        rsProductos.Open "select troquel,descripcion,precio,fabricante from productos where " & _
                     "descripcion like '" & Cadena & "%'" & " order by descripcion", cn, adOpenDynamic, adLockReadOnly, adCmdText
        If rsProductos.RecordCount = 0 Then
            MsgBox "NO EXISTE NADA CON LA DESCRIPCION INGRESADA ...!", vbExclamation, "Atencion !"
            Exit Sub
        End If
        dtgProductos.Visible = True
        Set dtgProductos.DataSource = rsProductos
        dtgProductos.SetFocus
    End If
End If
End Sub
Private Sub Descuento()
If rsFacturador.RecordCount = 0 Then
    MsgBox "NO HAY REGISTROS CARGADOS EN EL FACTURADOR PARA EL DESCUENTO..!", vbExclamation, "ATENCION !"
    Exit Sub
End If
If rsFacturador.EOF = False And rsFacturador.BOF = False Then
    vIdFactura = rsFacturador!idfactura
Else
    MsgBox "DEBE SELECCIONAR UN REGITRO DEL FACTURADOR PARA EFECTUAR UN DESCUENTO....!", vbExclamation, "ATENCIÓN !"
    Exit Sub
End If

If rsFacturador!precio = 0 Then
    MsgBox "EL PRODUCTO SELECCIONADO NO TIENE PRECIO PARA REALIZAR DESCUENTO ...", vbExclamation, "ATENCION !"
    txtProducto.SetFocus
    Exit Sub
End If

frmDescuento.Show vbModal

rsFacturador.Update
rsFacturador.Requery
dtgProductos.Refresh

Call CalculaImportes

End Sub
Private Sub CalculaImportes() 'calcula los totales del facturador
Dim reg As Integer
If rsFacturador.RecordCount = 0 Then
    lblTotVta.FormatString = "0.00"
    lbltotalDesc.FormatString = "0.00"
    lblCobrar.FormatString = "0.00"
    Exit Sub
End If
rsFacturador.Update
rsFacturador.Requery

reg = rsFacturador.RecordCount
vNeto = 0
vBruto = 0
vDesc = 0
rsFacturador.MoveFirst
For i = 1 To reg
    vNeto = vNeto + rsFacturador!precio   'acumulador compra
    vBruto = vBruto + rsFacturador!importe 'acumulador de compra con descuento
    vDesc = vNeto - vBruto 'total de descuentos
    rsFacturador.MoveNext
Next

lblTotVta.FormatString = vNeto
lbltotalDesc.FormatString = Round(vDesc, 2)
lblCobrar.FormatString = vBruto

End Sub
