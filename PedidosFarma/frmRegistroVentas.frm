VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRegistroVentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Ventas ..."
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12090
   Icon            =   "frmRegistroVentas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   12090
   Begin VB.Frame frameCaja 
      Caption         =   "Control de Caja"
      Height          =   3735
      Left            =   6360
      TabIndex        =   21
      Top             =   4560
      Visible         =   0   'False
      Width           =   3975
      Begin AIFCmp1.asxPowerButton cmdDarResultado 
         Height          =   300
         Left            =   1320
         TabIndex        =   28
         Top             =   2160
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "Calcular !"
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
      Begin VB.OptionButton optTT 
         Caption         =   "TT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         ToolTipText     =   "Caja Turno Tarde"
         Top             =   3360
         Width           =   855
      End
      Begin VB.OptionButton optTM 
         Caption         =   "TM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         ToolTipText     =   "Caja Turno Mañana"
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox txtCaja 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         Height          =   300
         Left            =   2640
         TabIndex        =   26
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtObs 
         Height          =   345
         Left            =   1800
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox txtResultado 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2640
         TabIndex        =   35
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtCredito 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         Height          =   300
         Left            =   2640
         TabIndex        =   25
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtInicio 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         Height          =   300
         Left            =   2640
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtExt 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         Height          =   300
         Left            =   2640
         TabIndex        =   24
         Top             =   720
         Width           =   1095
      End
      Begin AIFCmp1.asxPowerButton cmdOk 
         Height          =   495
         Left            =   1200
         TabIndex        =   29
         ToolTipText     =   "Queda registrado el Control de Caja"
         Top             =   3120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Picture         =   "frmRegistroVentas.frx":030A
         Caption         =   "G&rabar"
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
         PictureOffsetX  =   5
      End
      Begin AIFCmp1.asxPowerButton cmdCancel 
         Height          =   495
         Left            =   2520
         TabIndex        =   30
         Top             =   3120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Picture         =   "frmRegistroVentas.frx":08A4
         Caption         =   "Cancelar"
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
         PictureOffsetX  =   5
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Caja:"
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
         TabIndex        =   36
         Top             =   1800
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Inicio de Caja:"
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
         TabIndex        =   34
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Extracciones:"
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
         TabIndex        =   33
         Top             =   840
         Width           =   1410
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "A Crédito:"
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
         TabIndex        =   32
         Top             =   1320
         Width           =   1035
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Resultado:"
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
         TabIndex        =   31
         Top             =   2160
         Width           =   1140
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
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
         TabIndex        =   22
         Top             =   2640
         Width           =   1650
      End
   End
   Begin AIFCmp1.asxPowerButton cmdSalir 
      Cancel          =   -1  'True
      Height          =   1095
      Left            =   10440
      TabIndex        =   13
      ToolTipText     =   "Sale del modulo registro de ventas"
      Top             =   7200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1931
      FocusStyle      =   1
      BorderStyle     =   4
      Picture         =   "frmRegistroVentas.frx":0E3E
      Caption         =   "&Salir"
      CaptionAlignment=   7
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
   Begin VB.Frame Frame3 
      Caption         =   "Ingreso de Datos Nuevos"
      Height          =   2655
      Left            =   6360
      TabIndex        =   17
      Top             =   1800
      Width           =   5655
      Begin AIFCmp1.asxPowerButton cmdCaja 
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1085
         Picture         =   "frmRegistroVentas.frx":13CA
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
      Begin MSComCtl2.DTPicker dtpRegistro 
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   240
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
         Format          =   68878337
         CurrentDate     =   39309
      End
      Begin VB.TextBox txtTarde 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   3840
         MaxLength       =   8
         TabIndex        =   2
         ToolTipText     =   "Admite signo negativo"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtMañana 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3840
         MaxLength       =   8
         TabIndex        =   1
         ToolTipText     =   "Admite signo negativo"
         Top             =   720
         Width           =   1575
      End
      Begin AIFCmp1.asxPowerButton cmdGrabar 
         Height          =   495
         Left            =   2280
         TabIndex        =   3
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Picture         =   "frmRegistroVentas.frx":16E4
         Caption         =   "&Grabar"
         CaptionAlignment=   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         Height          =   495
         Left            =   3960
         TabIndex        =   4
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Picture         =   "frmRegistroVentas.frx":1C7E
         Caption         =   "&Cancelar"
         CaptionAlignment=   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureAlignment=   3
         PictureOffsetX  =   10
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Ventas del Día:"
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
         Width           =   1605
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Venta de la Tarde:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   1950
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Venta de la Mañana:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   2145
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtro por fecha"
      Height          =   1455
      Left            =   6360
      TabIndex        =   12
      Top             =   240
      Width           =   5655
      Begin AIFCmp1.asxPowerButton cmdVer 
         Height          =   420
         Left            =   3960
         TabIndex        =   9
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         Picture         =   "frmRegistroVentas.frx":2218
         Caption         =   "&Ver Datos"
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
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   375
         Left            =   3840
         TabIndex        =   8
         Top             =   360
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
         Format          =   68878337
         CurrentDate     =   39308
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   360
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
         Format          =   68878337
         CurrentDate     =   39308
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3000
         TabIndex        =   15
         Top             =   480
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Archivo de Ventas"
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      Begin MSDataGridLib.DataGrid dtgVentas 
         Height          =   7095
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   12515
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "Fecha"
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
            DataField       =   "mañana"
            Caption         =   "Mañana"
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
            DataField       =   "tarde"
            Caption         =   "Tarde"
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
            DataField       =   "total"
            Caption         =   "Total Diario"
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
               ColumnWidth     =   1454,74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1470,047
            EndProperty
            BeginProperty Column03 
            EndProperty
         EndProperty
      End
   End
   Begin AIFCmp1.asxPowerButton cmdInformes 
      Height          =   1095
      Left            =   10440
      TabIndex        =   11
      ToolTipText     =   "Informe de Ventas Anual"
      Top             =   6000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1931
      FocusStyle      =   1
      BorderStyle     =   4
      Picture         =   "frmRegistroVentas.frx":27B2
      Caption         =   "&Informe"
      CaptionAlignment=   7
      CaptionOffsetY  =   -5
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
   Begin AIFCmp1.asxFontLabel lblTotalMañana 
      Height          =   360
      Left            =   1560
      TabIndex        =   39
      Top             =   7920
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   635
      TextColor       =   12582912
      Caption         =   "Mañana"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
   End
   Begin AIFCmp1.asxFontLabel lblTotalTarde 
      Height          =   360
      Left            =   3120
      TabIndex        =   40
      Top             =   7920
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   635
      TextColor       =   33023
      Caption         =   "Mañana"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
   End
   Begin AIFCmp1.asxFontLabel lblTotal 
      Height          =   360
      Left            =   4680
      TabIndex        =   41
      Top             =   7920
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   635
      TextColor       =   49152
      Caption         =   "Mañana"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
   End
   Begin AIFCmp1.asxPowerButton cmdPuntoControl 
      Height          =   1095
      Left            =   10440
      TabIndex        =   42
      ToolTipText     =   "Informe de Ventas Anual"
      Top             =   4800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1931
      FocusStyle      =   1
      BorderStyle     =   4
      Picture         =   "frmRegistroVentas.frx":2ACC
      Caption         =   "&Pto.Ctrol"
      CaptionAlignment=   7
      CaptionOffsetY  =   -5
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
      Caption         =   "TOTALES:"
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
      TabIndex        =   16
      Top             =   7920
      Width           =   1110
   End
End
Attribute VB_Name = "frmRegistroVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private rptConsumo As New crptEntradas
Private rsVentas As New ADODB.Recordset
Private rsCajas As New ADODB.Recordset
Private vDesde As String
Private vHasta As String
Private vTotM, vTotT As Double
Private Bandera As Boolean
Private rpt_InfVentas As New crptVentasXfechas

Private Sub cmdCaja_Click()
If Len(txtMañana.Text) = 0 And Len(txtTarde.Text) = 0 Then
    MsgBox "DEBE INGRESAR EL TOTAL DE VENTA DEL TURNO QUE DESEA CONTROLAR", vbCritical, "ATENCION !!!"
    txtMañana.SetFocus
    Exit Sub
End If
frameCaja.Visible = True
If Len(txtMañana.Text) > 0 Then
    optTM.Value = True
    optTT.Value = False
Else
    optTT.Value = True
    optTM.Value = False
End If
txtInicio.Text = 0
txtExt.Text = 0
txtCredito.Text = 0
txtCaja.Text = 0
txtResultado.Text = 0
txtInicio.SetFocus
SendKeys "{home}+{end}"
End Sub

Private Sub cmdCancel_Click()
Err.Clear
On Error GoTo Solucion
frameCaja.Visible = False
'chekea si se activo la bandera
If Bandera = True Then
    Bandera = False
End If
If rsCajas.State = 1 Then
    rsCajas.Close
    Set rsCajas = Nothing
End If
Exit Sub
Solucion:
    MsgBox Err.Number & " - " & Err.Description, vbInformation, "Error del Sistema..."
    
End Sub

Private Sub cmdCancelar_Click()
txtMañana.Text = ""
txtTarde.Text = ""
If Time < 14 Then
    txtMañana.SetFocus
Else
    txtTarde.SetFocus
End If
vAgrega = True
End Sub

Private Sub cmdDarResultado_Click()
If Len(txtInicio.Text) = 0 Then
    MsgBox "DEBE INGRESAR EL INICIO DE CAJA !", vbCritical, "ATENCION !"
    txtInicio.SetFocus
    Exit Sub
End If
If Len(txtExt.Text) = 0 Then
    MsgBox "EL CAMPO EXTRACCIONES NO PUEDE ESTAR VACIO !!!", vbCritical, "ATENCION !"
    txtExt.SetFocus
    Exit Sub
End If
If Len(txtCredito.Text) = 0 Then
    MsgBox "EL CAMPO A CREDITO NO PUEDE ESTAR VACIO !!!", vbCritical, "ATENCION !"
    txtCredito.SetFocus
    Exit Sub
End If
If Len(txtCaja.Text) = 0 Then
    MsgBox "EL CAMPO CAJA NO PUEDE ESTAR VACIO !!!", vbCritical, "ATENCION !"
    txtCaja.SetFocus
    Exit Sub
End If
Err.Clear
On Error GoTo ErrProc
If optTM.Value = True Then
    txtResultado.Text = ((Val(txtInicio.Text) + Val(txtMañana.Text)) - (Val(txtExt.Text) + Val(txtCredito.Text)))
Else
    txtResultado.Text = ((Val(txtInicio.Text) + Val(txtTarde.Text)) - (Val(txtExt.Text) + Val(txtCredito.Text)))
End If

txtResultado.Text = Round((Val(txtCaja.Text)) - ((txtResultado.Text)), 2)
Bandera = True 'indica que el resultado ha sido procesado
Exit Sub
ErrProc:
  MsgBox Err.Number & " " & Err.Description, vbInformation, "Información"

End Sub

Private Sub cmdGrabar_Click()
If Len(txtMañana.Text) = 0 And Len(txtTarde.Text) = 0 Then
    MsgBox "DEBE HABER ALGUN DATO SOBRE LAS VENTAS PARA GRABAR ...", vbCritical, "ATENCION !"
    If Time < 14 Then
        txtMañana.SetFocus
    Else
        txtTarde.SetFocus
    End If
End If

'busca la fecha para sumar montos de misma fecha
strSQL = "Fecha = #" & Format(dtpRegistro.Value, "dd/mm/yyyy") & "#"
rsVentas.Find strSQL, , adSearchForward, 1
If rsVentas.EOF = True Then
    vAgrega = True
Else
    vAgrega = False
End If

If vAgrega = True Then
    rsVentas.AddNew
    rsVentas!fecha = dtpRegistro.Value
    rsVentas!mañana = 0
    rsVentas!tarde = 0
End If
If Len(txtMañana.Text) > 0 Then
    If IsNull(rsVentas!mañana) = True Then
        rsVentas!mañana = Round(Val(txtMañana.Text), 2)
    Else
        rsVentas!mañana = (rsVentas!mañana) + (Val(txtMañana.Text))
    End If
End If
If Len(txtTarde.Text) > 0 Then
    If IsNull(rsVentas!tarde) = True Then
        rsVentas!tarde = Round(Val(txtTarde.Text), 2)
    Else
        rsVentas!tarde = (rsVentas!tarde) + (Val(txtTarde.Text))
    End If
End If
rsVentas!total = (rsVentas!mañana) + (rsVentas!tarde)
rsVentas.Update
rsVentas.Requery
dtgVentas.Refresh

Call CalculaTotales

'termino el proceso blanqueando el ingreso de datos
txtMañana.Text = ""
txtTarde.Text = ""
If Time < 14 Then
    txtMañana.SetFocus
Else
    txtTarde.SetFocus
End If
vAgrega = True
End Sub

Private Sub cmdInformes_Click()

frmPideClave.Show vbModal
If TempNivel = 1 Then
    rpt_InfVentas.Database.SetDataSource rsVentas
    
    Set rptGeneral = rpt_InfVentas ' Asigna el reporte al objeto reporte general utilizado
                               ' en el Form de la Vista Previa.
    frmVistaPrevia.Show vbModal
    
    Set rpt_InfVentas = Nothing
End If
End Sub

Private Sub cmdOk_Click()
If Len(txtResultado.Text) = 0 Or Bandera = False Then
    MsgBox "DEBE CALCULAR EL RESULTADO DE CAJA PARA GRABAR !!!", vbExclamation, "ATENCION !!!"
    cmdDarResultado.SetFocus
    Exit Sub
End If
Err.Clear
On Error GoTo ErrProc

If Len(txtObs.Text) = 0 Then
    MsgBox "DEBE INGRESAR EN EL CAMPO OBSERVACIONES EL NOMBRE DE LA PERSONA" + Chr(13) & _
            "QUE HIZO EL RECUENTRO DE CAJA...", vbExclamation, "ATENCION !!!"
    txtObs.SetFocus
    Exit Sub
End If

SioNo = MsgBox("ESTA SEGURO DE REGISTRAR ESTE CONTROL DE CAJA ???", vbExclamation + vbYesNo, "ATENCION !!!")
If SioNo = vbYes Then
    rsCajas.Open "select * from ControlCajas", cn, adOpenDynamic, adLockOptimistic, adCmdText
    rsCajas.AddNew
    rsCajas!fecha = dtpRegistro.Value
    rsCajas!inicio = txtInicio.Text
    rsCajas!extracciones = txtExt.Text
    rsCajas!credito = txtCredito.Text
    rsCajas!caja = txtCaja.Text
    rsCajas!observaciones = txtObs.Text & ""
    rsCajas!resultado = Str(txtResultado.Text)
    If optTM.Value = True Then
        rsCajas!turno = "MAÑANA"
    Else
        rsCajas!turno = "TARDE"
    End If
    Bandera = False
    rsCajas.Update
    rsCajas.Close
    Set rsCajas = Nothing
    If txtResultado.Text > 0 Then
        SioNo = MsgBox("LA CAJA A TENIDO SOBRANTE, DESEA SUMARLO A LAS VENTAS ?", vbInformation + vbYesNo, "ATENCION SOBRANTES")
        If SioNo = vbYes Then
            If optTM.Value = True Then
                txtMañana.Text = (Val(txtMañana.Text) + (txtResultado.Text))
            Else
                txtTarde.Text = (Val(txtTarde.Text) + (txtResultado.Text))
            End If
        End If
    End If
    frameCaja.Visible = False
End If
Exit Sub
ErrProc:
  MsgBox Err.Description, vbInformation, "Información"
End Sub

Private Sub cmdPuntoControl_Click()
frmPideClave.Show vbModal
If TempNivel = 1 Then
    frmEjecutarPuntoControl.Show
Else
    MsgBox "NO ESTA AUTORIZADO PARA INGRESAR A PUNTOS DE CONTROL ...!", vbCritical, "SEGURIDAD DEL SISTEMA ..."
End If
End Sub

Private Sub cmdSalir_Click()
If Len(txtMañana.Text) > 0 Or Len(txtTarde.Text) > 0 Then
    SioNo = MsgBox("Esta seguro de salir de este módulo ??? Hay totales de Ventas," + Chr(13) & _
            "que no han sido grabados en la base de datos, y pueden perderse...", vbExclamation + vbYesNo, "ATENCION !")
    If SioNo = vbYes Then
        Unload Me
    End If
Else
    Unload Me
End If
End Sub

Private Sub cmdVer_Click()
rsVentas.Close
rsVentas.Open "select * from ventas where fecha between #" & Format(dtpDesde.Value, "mm/dd/yyyy") & "# and #" & Format(dtpHasta.Value, "mm/dd/yyyy") & "#", cn, adOpenDynamic, adLockOptimistic, adCmdText
Set dtgVentas.DataSource = rsVentas
dtgVentas.Refresh
Call CalculaTotales

End Sub
Private Sub dtpRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyAscii = 0
    txtTarde.SetFocus
End If
End Sub
Private Sub Form_Load()
Me.Top = 20
Me.Left = 25
Bandera = False
vAgrega = True

dtpDesde.Value = "01/" & Month(Date) & "/" & Year(Date)
If Month(Date) = 1 Or Month(Date) = 3 Or Month(Date) = 5 Or Month(Date) = 7 Or Month(Date) = 8 Or Month(Date) = 10 Or Month(Date) = 12 Then
    dtpHasta.Value = Format("31/" & Month(Date) & "/" & Year(Date), "dd/mm/yyyy")
End If
If Month(Date) = 4 Or Month(Date) = 6 Or Month(Date) = 9 Or Month(Date) = 11 Then
    dtpHasta.Value = Format("30/" & Month(Date) & "/" & Year(Date), "dd/mm/yyyy")
End If
If Month(Date) = 2 Then
    If Int(Year(Date) / 4) = (Year(Date) / 4) Then
        dtpHasta.Value = Format("29/" & Month(Date) & "/" & Year(Date), "dd/mm/yyyy")
    Else
        dtpHasta.Value = Format("28/" & Month(Date) & "/" & Year(Date), "dd/mm/yyyy")
    End If
End If

dtpRegistro.Value = Date
rsVentas.Open "select * from ventas where fecha between #" & Format(dtpDesde.Value, "mm/dd/yyyy") & "# and #" & Format(dtpHasta.Value, "mm/dd/yyyy") & "# order by fecha", cn, adOpenDynamic, adLockOptimistic, adCmdText
Set dtgVentas.DataSource = rsVentas
dtgVentas.Refresh

Call CalculaTotales

End Sub
Private Sub Form_Unload(cancel As Integer)
If rsVentas.State = 1 Then
    rsVentas.Close
    Set rsVentas = Nothing
End If
If rsCajas.State = 1 Then
    rsCajas.Close
    Set rsCajas = Nothing
End If
End Sub
Private Sub optTM_Click()
   optTT.Value = False
End Sub
Private Sub optTT_Click()
    optTM.Value = False
End Sub
Private Sub txtCaja_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        txtObs.SetFocus
    Case 44 'para que no acepte la entrada de la coma
        KeyAscii = 0
    Case 45
    Case 46
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub txtCredito_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        txtCaja.SetFocus
        SendKeys "{end}+{home}"
    Case 44 'para que no acepte la entrada de la coma
        KeyAscii = 0
    Case 45
    Case 46
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub txtExt_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        txtCredito.SetFocus
        SendKeys "{end}+{home}"
    Case 44 'para que no acepte la entrada de la coma
        KeyAscii = 0
    Case 45
    Case 46
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub txtInicio_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        txtExt.SetFocus
        SendKeys "{end}+{home}"
    Case 44 'para que no acepte la entrada de la coma
        KeyAscii = 0
    Case 45
    Case 46
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub txtMañana_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        cmdCaja.SetFocus
    Case 44 'para que no acepte la entrada de la coma
        KeyAscii = 0
    Case 45
    Case 46
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub txtTarde_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        cmdCaja.SetFocus
    Case 44 'para que no acepte la coma
        KeyAscii = 0
    Case 45
    Case 46
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub CalculaTotales()
'mostrando totales
vTotM = 0
vTotT = 0
If rsVentas.RecordCount > 0 Then
    rsVentas.MoveFirst
    Do While rsVentas.EOF = False
        If IsNull(rsVentas!mañana) = False Then
            vTotM = vTotM + rsVentas!mañana
        End If
        If IsNull(rsVentas!tarde) = False Then
            vTotT = vTotT + rsVentas!tarde
        End If
        rsVentas.MoveNext
    Loop
End If

lblTotalMañana.Caption = vTotM
lblTotalTarde.Caption = vTotT

lblTotal.Caption = (vTotM + vTotT)
End Sub
