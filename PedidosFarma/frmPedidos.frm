VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPedidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestion de Pedidos a Proveedores ..."
   ClientHeight    =   8940
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   14745
   HasDC           =   0   'False
   Icon            =   "frmPedidos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   15603.19
   ScaleMode       =   0  'User
   ScaleWidth      =   91782.56
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Listado del Pedido"
      Height          =   7935
      Left            =   120
      TabIndex        =   24
      Top             =   960
      Width           =   14535
      Begin AIFCmp1.asxPowerButton cmdDevo 
         Height          =   855
         Left            =   12120
         TabIndex        =   34
         Top             =   2760
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1508
         FocusStyle      =   1
         ShadowDkColor   =   0
         Picture         =   "frmPedidos.frx":030A
         Caption         =   "Devoluciones y NC"
         CaptionOffsetY  =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextColor       =   255
      End
      Begin VB.CheckBox ChkOrdEnv 
         Caption         =   "Ordena Enviados Alfabéticamente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11880
         TabIndex        =   32
         Top             =   4800
         Width           =   2415
      End
      Begin VB.CheckBox chkOrdenaPedido 
         Caption         =   "Ordena Pedido Alfabéticamente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12000
         TabIndex        =   31
         Top             =   1920
         Width           =   2415
      End
      Begin MSDataGridLib.DataGrid dtgMuestraProductos 
         Height          =   6735
         Left            =   2040
         TabIndex        =   9
         Top             =   960
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   11880
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648384
         HeadLines       =   2
         RowHeight       =   19
         RowDividerStyle =   4
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
            Caption         =   "PrecVta"
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
               ColumnWidth     =   1395,213
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4995,213
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   959,811
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dtgFaltantes 
         Height          =   2055
         Left            =   120
         TabIndex        =   13
         Top             =   5640
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   3625
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   8421631
         HeadLines       =   1
         RowHeight       =   15
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "FALTANTES"
         ColumnCount     =   6
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
            DataField       =   "troquel"
            Caption         =   "Troquel/Cod."
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
         BeginProperty Column03 
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
         BeginProperty Column04 
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
         BeginProperty Column05 
            DataField       =   "idproveedor"
            Caption         =   "Cod.Proveedor"
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
               ColumnWidth     =   1124,787
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1470,047
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   5564,977
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   599,811
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1170,142
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dtgEnviado 
         Height          =   1695
         Left            =   120
         TabIndex        =   10
         Top             =   3840
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   2990
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648384
         HeadLines       =   1
         RowHeight       =   15
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ENVIADOS"
         ColumnCount     =   6
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
            DataField       =   "troquel"
            Caption         =   "Troquel/Cod."
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
            Caption         =   "Descripcion del Producto"
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
         BeginProperty Column04 
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
         BeginProperty Column05 
            DataField       =   "idproveedor"
            Caption         =   "Cod.Proveedor"
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
               ColumnWidth     =   1154,835
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1500,095
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   5490,142
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   645,165
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   14,74
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox optFiltroFecha 
         Caption         =   "Filtra Fecha del Día"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11880
         TabIndex        =   18
         Top             =   4320
         Width           =   2415
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Ver todos los Faltantes ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10080
         TabIndex        =   17
         ToolTipText     =   "Visualiza todas las faltas sin diferenciar Proveedor"
         Top             =   6360
         Width           =   3135
      End
      Begin VB.OptionButton optIndividual 
         Caption         =   "Ver Faltantes del Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10080
         TabIndex        =   16
         Top             =   6000
         Value           =   -1  'True
         Width           =   3375
      End
      Begin VB.TextBox txtTroquel 
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
         Left            =   120
         MaxLength       =   20
         TabIndex        =   1
         Top             =   600
         Width           =   1815
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx5 
         Height          =   240
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   423
         Caption         =   "Cod. Troquel"
      End
      Begin VB.TextBox txtCantidad 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         MaxLength       =   6
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx4 
         Height          =   240
         Left            =   8520
         TabIndex        =   26
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   423
         Caption         =   "Cantidad"
      End
      Begin VB.TextBox txtDescripcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   6375
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx3 
         Height          =   240
         Left            =   2040
         TabIndex        =   25
         Top             =   240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   423
         Caption         =   "Descripcion del Producto"
      End
      Begin MSDataGridLib.DataGrid dtgPedido 
         Height          =   2655
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   4683
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   3
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
            Caption         =   "Descripción"
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
            DataField       =   "cantidad"
            Caption         =   "Cant"
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
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   6359,812
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1379,906
            EndProperty
         EndProperty
      End
      Begin AIFCmp1.asxPowerButton cmdBorrar 
         Height          =   495
         Left            =   10080
         TabIndex        =   6
         ToolTipText     =   "Borra el registro seleccionado del pedido"
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmPedidos.frx":08A4
         Caption         =   "&borrar"
         CaptionAlignment=   5
         CaptionOffsetX  =   -10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureAlignment=   3
         PictureOffsetX  =   15
         TextColor       =   255
      End
      Begin AIFCmp1.asxPowerButton cmdFalta 
         Height          =   495
         Left            =   10080
         TabIndex        =   4
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmPedidos.frx":12B6
         Caption         =   "&Falta"
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
      Begin AIFCmp1.asxPowerButton cmdEnviado 
         Height          =   495
         Left            =   10080
         TabIndex        =   14
         ToolTipText     =   "Pasa el registro seleccionado a los Enviados"
         Top             =   6840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmPedidos.frx":15D0
         Caption         =   "Enviado"
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
      Begin AIFCmp1.asxPowerButton cmdFaltaEnviado 
         Height          =   495
         Left            =   10080
         TabIndex        =   11
         ToolTipText     =   "Si vino facturado como faltante..."
         Top             =   4320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmPedidos.frx":1A22
         Caption         =   "&Falta"
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
      Begin AIFCmp1.asxPowerButton cmdNuevo 
         Height          =   495
         Left            =   10080
         TabIndex        =   5
         ToolTipText     =   "Agrega un producto nuevo al pedido"
         Top             =   2520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmPedidos.frx":1D3C
         Caption         =   "&Nuevo Pdto"
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
         PictureOffsetX  =   15
      End
      Begin AIFCmp1.asxPowerButton cmdBorraEnviado 
         Height          =   495
         Left            =   10080
         TabIndex        =   12
         ToolTipText     =   "Borra el registro seleccionado de Enviados"
         Top             =   4920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmPedidos.frx":274E
         Caption         =   "&borrar"
         CaptionAlignment=   5
         CaptionOffsetX  =   -10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureAlignment=   3
         PictureOffsetX  =   15
         TextColor       =   255
      End
      Begin AIFCmp1.asxPowerButton cmdBorrarFaltas 
         Height          =   495
         Left            =   12000
         TabIndex        =   15
         ToolTipText     =   "Borra el registro seleccionado de Faltantes"
         Top             =   6840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmPedidos.frx":3160
         Caption         =   "&borrar"
         CaptionAlignment=   5
         CaptionOffsetX  =   -10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureAlignment=   3
         PictureOffsetX  =   15
         TextColor       =   255
      End
      Begin AIFCmp1.asxPowerButton cmdImprimir 
         Height          =   495
         Left            =   12000
         TabIndex        =   7
         ToolTipText     =   "Imprime la lista del pedido"
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmPedidos.frx":3B72
         Caption         =   "&Imprimir Pedido"
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
         PictureOffsetX  =   10
      End
      Begin AIFCmp1.asxPowerButton cmdOk 
         Height          =   495
         Left            =   10080
         TabIndex        =   29
         Top             =   1310
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmPedidos.frx":4584
         Caption         =   "&Ok"
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
      Begin AIFCmp1.asxPowerButton cmdAgregar 
         Height          =   495
         Left            =   10080
         TabIndex        =   30
         ToolTipText     =   "Agrega el producto al pedido"
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmPedidos.frx":489E
         Caption         =   "&Pedir"
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
      Begin AIFCmp1.asxPowerButton cmdRefrescar 
         Height          =   495
         Left            =   12000
         TabIndex        =   33
         ToolTipText     =   "Agrega el producto al pedido"
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmPedidos.frx":4CF0
         Caption         =   "&Refrescar"
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
      Begin VB.Line Line10 
         X1              =   9720
         X2              =   10080
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line9 
         X1              =   11640
         X2              =   12000
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line8 
         X1              =   11640
         X2              =   12000
         Y1              =   7080
         Y2              =   7080
      End
      Begin VB.Line Line7 
         X1              =   9720
         X2              =   10080
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Line Line6 
         X1              =   9720
         X2              =   10080
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line5 
         X1              =   9720
         X2              =   10080
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line4 
         X1              =   9720
         X2              =   10080
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line3 
         X1              =   9720
         X2              =   10080
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line2 
         X1              =   9720
         X2              =   10080
         Y1              =   7080
         Y2              =   7080
      End
      Begin VB.Line Line1 
         X1              =   9720
         X2              =   10200
         Y1              =   5160
         Y2              =   5160
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14535
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx6 
         Height          =   240
         Left            =   2160
         TabIndex        =   28
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   423
         Caption         =   "Codigo"
      End
      Begin AIFCmp1.asxPowerButton cmdCambiaProveedor 
         Height          =   495
         Left            =   10080
         TabIndex        =   19
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         Picture         =   "frmPedidos.frx":5142
         Caption         =   "&Cambiar Proveedor"
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
      Begin AIFCmp1.asxPowerBanner lblNombreProveedor 
         Height          =   375
         Left            =   3480
         Top             =   360
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   661
         EndColor        =   65280
         FormatString    =   "Nombre Proveedor"
         Orientation     =   0
         BorderStyle     =   2
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
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx2 
         Height          =   240
         Left            =   3480
         TabIndex        =   23
         Top             =   120
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   423
         Caption         =   "Nombre Proveedor"
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   405
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   714
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
         Format          =   249888769
         CurrentDate     =   39294
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx1 
         Height          =   240
         Left            =   240
         TabIndex        =   22
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   423
         Caption         =   "Fecha"
      End
      Begin AIFCmp1.asxPowerButton cmdSalir 
         Height          =   495
         Left            =   12360
         TabIndex        =   20
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         Picture         =   "frmPedidos.frx":5594
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
      Begin AIFCmp1.asxPowerBanner lblCodigo 
         Height          =   375
         Left            =   2160
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         EndColor        =   65280
         FormatString    =   "Nombre Proveedor"
         Orientation     =   0
         BorderStyle     =   2
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
End
Attribute VB_Name = "frmPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsItems As New ADODB.Recordset
Private rsEnviado As New ADODB.Recordset
Private rsFaltantes As New ADODB.Recordset
Private rsProveedor As New ADODB.Recordset
Private rsProductos As New ADODB.Recordset
Private rsImpPedido As New ADODB.Recordset
Private rptPedido As New crptPedidoAngosto

Private Sub chkOrdenaPedido_Click()
If chkOrdenaPedido.Value = 1 Then
    'cambia el orden de la lista de pedidos
     rsItems.Sort = "descripcion"
Else
    rsItems.Sort = "fecha,descripcion"
    'Call RefrescaDatos
End If
End Sub
Private Sub ChkOrdEnv_Click()
If ChkOrdEnv.Value = 1 Then
    rsEnviado.Sort = "descripcion"
Else
    rsEnviado.Sort = "fecha,descripcion"
End If
End Sub
Private Sub cmdAgregar_Click()
If Len(txtDescripcion.Text) = 0 Then
    MsgBox "DEBE COMPLETAR LA DESCRIPCION DEL PRODUCTO ...", vbCritical, "ATENCION !"
    txtDescripcion.SetFocus
    Exit Sub
End If
If Len(txtCantidad) = 0 Or Val(txtCantidad.Text) = 0 Then
    MsgBox "LA CANTIDAD A PEDIR DEBE SER DISTINTA DE CERO ...!", vbExclamation, "ATENCION !"
    txtCantidad.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If
rsItems.AddNew
rsItems!fecha = Format(dtpFecha.Value, "dd/mm/yyyy")
rsItems!troquel = txtTroquel
rsItems!descripcion = txtDescripcion
rsItems!cantidad = txtCantidad
rsItems!Estado = 3
rsItems!idproveedor = vidPro
rsItems.Update
rsItems.Requery
dtgPedido.Refresh
txtTroquel.Text = ""
txtDescripcion.Text = ""
txtCantidad.Text = 1 'inicialisa siempre la cantidad en uno como defecto
dtgMuestraProductos.Visible = False
txtTroquel.SetFocus
End Sub

Private Sub cmdBorraEnviado_Click()
If rsEnviado.RecordCount = 0 Then
    MsgBox "NO HAY REGISTROS DE ENVIADOS PARA BORRAR ...!", vbExclamation, "ATENCION !"
    Exit Sub
End If
SioNo = MsgBox("ESTA SEGURO DE BORRAR ESTE PRODUCTO DE LA LISTA DE ENVIADOS ?", vbExclamation + vbYesNo, "ATENCION !!!")
If SioNo = vbYes Then
    rsEnviado.Delete
    rsEnviado.Update
    rsEnviado.Requery
    dtgEnviado.Refresh
End If
txtTroquel.SetFocus
End Sub

Private Sub cmdBorrar_Click()
If rsItems.RecordCount = 0 Then
    MsgBox "NO HAY REGISTROS PARA BORRAR ...!", vbCritical, "ATENCION !"
    Exit Sub
End If
SioNo = MsgBox("ESTA SEGURO DE BORRAR ESTE PRODUCTO DE LA LISTA DE PEDIDO ?", vbExclamation + vbYesNo, "ATENCION !!!")
If SioNo = vbYes Then
    rsItems.Delete
    rsItems.Update
    rsItems.Requery
    dtgPedido.Refresh
    txtCantidad.SetFocus
End If
txtTroquel.SetFocus
End Sub
Private Sub cmdBuscar_Click()
vProducto = Trim(txtDescripcion.Text)
frmBuscaProducto.Show vbModal
txtDescripcion.Text = vProducto
txtTroquel.Text = VarTroquel
If Len(txtDescripcion.Text) = 0 Then
    txtDescripcion.SetFocus
Else
    txtCantidad.SetFocus
    SendKeys "{home}+{end}"
End If
dtgMuestraProductos.Visible = False
End Sub

Private Sub cmdBorrarFaltas_Click()
If rsFaltantes.RecordCount = 0 Then
    MsgBox "NO HAY REGISTROS DE FALTAS PARA BORRAR...!", vbExclamation, "ATENCION !"
    Exit Sub
End If
SioNo = MsgBox("ESTA SEGURO DE BORRAR ESTE PRODUCTO DE LA LISTA DE FALTANTES ?", vbExclamation + vbYesNo, "ATENCION !!!")
If SioNo = vbYes Then
    rsFaltantes.Delete
    rsFaltantes.Update
    rsFaltantes.Requery
    dtgFaltantes.Refresh
End If
txtTroquel.SetFocus
End Sub

Private Sub cmdCambiaProveedor_Click()
frmProveedores.Show vbModal
lblNombreProveedor.FormatString = VarPro
lblCodigo.FormatString = vidPro
Call RefrescaDatos
txtTroquel.SetFocus
End Sub

Private Sub cmdDevo_Click()
frmDevoluciones.Show
End Sub

Private Sub cmdEnviado_Click()
If rsFaltantes.RecordCount = 0 Then
    MsgBox "No hay registros de faltantes ...!", vbCritical, "Atención !!!"
    Exit Sub
End If
rsFaltantes!Estado = 1
rsFaltantes!fecha = dtpFecha.Value
rsFaltantes!idproveedor = vidPro
rsFaltantes.Update
Call RefrescaDatos
End Sub
Private Sub cmdFalta_Click()
If rsItems.RecordCount = 0 Then
    MsgBox "NO HAY REGISTROS CARGADOS ...", vbExclamation, "Atención !!!"
    Exit Sub
End If
rsItems!Estado = 2
rsItems.Update
Call RefrescaDatos
End Sub

Private Sub cmdFaltaEnviado_Click()
If rsEnviado.RecordCount = 0 Then
    MsgBox "NO HAY REGISTROS CARGADOS ...", vbExclamation, "Atención !!!"
    Exit Sub
End If
rsEnviado!Estado = 2
rsEnviado.Update
Call RefrescaDatos
End Sub

Private Sub cmdImprimir_Click()

SioNo = MsgBox("(FIJESE QUE LA IMPRESORA CONTENGA PAPEL A4)" + Chr(13) & _
             "ESTA SEGURO DE IMPRIMIR EL PEDIDO ???", vbInformation + vbYesNo, "ATENCION !!!")

If SioNo = vbNo Then
    Exit Sub
End If

rsImpPedido.Open "SELECT Pedidos.Fecha, Pedidos.Troquel, Pedidos.Descripcion, Pedidos.cantidad, Proveedores.Nombre, Proveedores.Telefono" & _
                " FROM Pedidos INNER JOIN Proveedores ON Pedidos.idProveedor = Proveedores.idProveedor" & _
                " where pedidos.idproveedor = " & vidPro & " and pedidos.estado = " & 3, cn, adOpenDynamic, adLockReadOnly, adCmdText

rptPedido.Database.SetDataSource rsImpPedido

rptPedido.PrintOut

rsImpPedido.Close
Set rsImpPedido = Nothing

End Sub

Private Sub cmdNuevo_Click()
frmNuevoProducto.Show vbModal
rsItems.Requery
dtgPedido.Refresh
txtDescripcion.SetFocus
End Sub

Private Sub cmdok_Click()
If rsItems.RecordCount = 0 Then
    MsgBox "NO HAY REGISTROS EN EL PEDIDO ...!", vbExclamation, "ATENCION !"
    Exit Sub
End If
rsItems!Estado = 1
rsItems!fecha = dtpFecha.Value
rsItems.Update
Call RefrescaDatos
End Sub

Private Sub cmdRefrescar_Click()
Call RefrescaDatos
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub dtgMuestraProductos_DblClick()
txtTroquel.Text = rsProductos!troquel
txtDescripcion.Text = rsProductos!descripcion
rsProductos.Close
Set rsProductos = Nothing
dtgMuestraProductos.Visible = False
txtCantidad.SetFocus
SendKeys "{home}+{end}"
End Sub
Private Sub dtgMuestraProductos_GotFocus()
    KeyAscii = 0
End Sub
Private Sub dtgMuestraProductos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtTroquel.Text = rsProductos!troquel & ""
    txtDescripcion.Text = rsProductos!descripcion
    rsProductos.Close
    Set rsProductos = Nothing
    dtgMuestraProductos.Visible = False
    txtCantidad.SetFocus
    SendKeys "{home}+{end}"
End If
If KeyAscii = 27 Then
    dtgMuestraProductos.Visible = False
    txtTroquel.Text = ""
    rsProductos.Close
    Set rsProductos = Nothing
    txtTroquel.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub
Private Sub dtgMuestraProductos_LostFocus()
KeyAscii = 0
dtgMuestraProductos.Visible = False
If rsProductos.State = 1 Then
    rsProductos.Close
    Set rsProductos = Nothing
End If
End Sub
Private Sub dtpFecha_Change()
Call RefrescaDatos
End Sub
Private Sub Form_Load()
Me.Top = 20
Me.Left = 25

rsProveedor.Open "select * from proveedores order by idproveedor", cn, adOpenDynamic, adLockReadOnly
vidPro = rsProveedor!idproveedor

dtpFecha.Value = Date
VarPro = rsProveedor!nombre
lblNombreProveedor.FormatString = VarPro
lblCodigo.FormatString = vidPro

Call RefrescaDatos

dtgMuestraProductos.Visible = False
txtCantidad.Text = 1
End Sub
Private Sub Form_Unload(cancel As Integer)
If rsItems.State = 1 Then
    rsItems.Close
    Set rsItems = Nothing
End If
If rsFaltantes.State = 1 Then
    rsFaltantes.Close
    Set rsFaltantes = Nothing
End If
If rsEnviado.State = 1 Then
    rsEnviado.Close
    Set rsEnviado = Nothing
End If
If rsProveedor.State = 1 Then
    rsProveedor.Close
    Set rsProveedor = Nothing
End If
End Sub
Private Sub optFiltroFecha_Click()
    Call RefrescaDatos
End Sub
Private Sub optIndividual_Click()
OptTodos.Value = False
rsFaltantes.Close
rsFaltantes.Open "select * from pedidos where idproveedor = " & vidPro & " and Estado= " & 2 & " order by descripcion", cn, adOpenKeyset, adLockOptimistic, adCmdText
Set dtgFaltantes.DataSource = rsFaltantes
dtgFaltantes.Refresh
End Sub
Private Sub OptTodos_Click()
optIndividual.Value = False
rsFaltantes.Close
rsFaltantes.Open "select * from pedidos where Estado= " & 2 & " order by descripcion", cn, adOpenKeyset, adLockOptimistic, adCmdText
Set dtgFaltantes.DataSource = rsFaltantes
dtgFaltantes.Refresh
End Sub
Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
'Controla que solo se ingresen numeros
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        cmdAgregar.SetFocus
    Case 27
        txtTroquel.Text = ""
        txtDescripcion.Text = ""
        txtTroquel.SetFocus
        SendKeys "{home}+{end}"
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
KeyAscii = 0
If dtgMuestraProductos.Visible = False Then
    Exit Sub
End If
If KeyCode = 40 Then
    dtgMuestraProductos.SetFocus
End If
End Sub
Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 27 Then
    KeyAscii = 0
    dtgMuestraProductos.Visible = False
    txtTroquel.Text = ""
    txtDescripcion.Text = ""
    txtTroquel.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If

If KeyAscii = 13 Then
    KeyAscii = 0
    If Len(txtDescripcion.Text) = 0 Then
        Exit Sub
    End If
    Dim Cadena As String
    Cadena = Replace(txtDescripcion.Text, " ", "%", 1)
    If rsProductos.State = 1 Then
        rsProductos.Close
        Set rsProductos = Nothing
    End If
    rsProductos.Open "select troquel,descripcion,precio from productos where " & _
                     "descripcion like '" & Cadena & "%'" & " order by descripcion", cn, adOpenDynamic, adLockReadOnly, adCmdText
    If rsProductos.RecordCount = 0 Then
        MsgBox "NO EXISTE NINGUN PRODUCTO QUE COINCIDA ...!", vbExclamation, "BUSQUEDA ERRONEA !!!"
        rsProductos.Close
        Set rsProductos = Nothing
        dtgMuestraProductos.Visible = False
        txtTroquel.SetFocus
        SendKeys "{home}+{end}"
    Else
        dtgMuestraProductos.Visible = True
        Set dtgMuestraProductos.DataSource = rsProductos
        dtgMuestraProductos.Refresh
        dtgMuestraProductos.SetFocus
    End If
End If

End Sub
Private Sub txtTroquel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtDescripcion.SetFocus
End If
If KeyAscii = 27 Then
    txtDescripcion.Text = ""
    txtTroquel.Text = ""
    txtTroquel.SetFocus
    If dtgMuestraProductos.Visible = True Then
        dtgMuestraProductos.Visible = False
    End If
End If
End Sub
Private Sub txtTroquel_LostFocus()
Err.Clear
On Error GoTo Solucion
If Len(txtTroquel) = 0 Then Exit Sub
If rsProductos.State = 1 Then
    rsProductos.Close
    Set rsProductos = Nothing
    If Me.dtgMuestraProductos.Visible = True Then
        dtgMuestraProductos.Visible = False
    End If
End If
rsProductos.Open "Select troquel, descripcion, precio from productos " & _
                "where troquel = '" & txtTroquel.Text & "'", cn, adOpenDynamic, adLockReadOnly, adCmdText
If rsProductos.RecordCount = 0 Then
    MsgBox "EL NUMERO DE TROQUEL NO SE ENCUENTRA REGISTRADO !", vbCritical, "NO EXISTE EL PRODUCTO..."
    txtTroquel.SetFocus
    SendKeys "{home}+{end}"
Else
    txtDescripcion.Text = rsProductos!descripcion
    txtCantidad.SetFocus
    SendKeys "{home}+{end}"
End If
rsProductos.Close
Set rsProductos = Nothing
Exit Sub

Solucion:
   MsgBox Err.Number & "-" & Err.Description, vbInformation, "Error del Sistema..."

End Sub
Private Sub RefrescaDatos()
'refresca los datos de todos los recorset y data grip

'codigos de estado: 1= Enviado, 2= Faltante, 3=Sin pedir

'recorset para la lista de pedido sin pedir
If rsItems.State = 1 Then
    rsItems.Close
    Set rsItems = Nothing
End If
rsItems.Open "select * from pedidos where idproveedor = " & vidPro & " and Estado= " & 3 & " order by fecha,descripcion", cn, adOpenDynamic, adLockOptimistic, adCmdText
Set dtgPedido.DataSource = rsItems
dtgPedido.Refresh

'Recorset para los productos enviados sin problema
If rsEnviado.State = 1 Then
    rsEnviado.Close
    Set rsEnviado = Nothing
End If
rsEnviado.Open "select * from pedidos where idproveedor = " & vidPro & " and Estado= " & 1 & " order by fecha desc,descripcion", cn, adOpenDynamic, adLockOptimistic, adCmdText
Set dtgEnviado.DataSource = rsEnviado
dtgEnviado.Refresh

'Recorset para los productos faltantes
If rsFaltantes.State = 1 Then
    rsFaltantes.Close
    Set rsFaltantes = Nothing
End If

If optIndividual.Value = True Then
    rsFaltantes.Open "select * from pedidos where idproveedor = " & vidPro & " and Estado= " & 2 & " order by descripcion", cn, adOpenDynamic, adLockOptimistic, adCmdText
Else
    rsFaltantes.Open "select * from pedidos where Estado= " & 2 & " order by descripcion", cn, adOpenDynamic, adLockOptimistic, adCmdText
End If
Set dtgFaltantes.DataSource = rsFaltantes
dtgFaltantes.Refresh
End Sub
