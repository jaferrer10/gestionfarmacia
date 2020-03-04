VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCompras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Facturas de Compras a Proveedores ..."
   ClientHeight    =   8865
   ClientLeft      =   2700
   ClientTop       =   1695
   ClientWidth     =   13320
   Icon            =   "frmCompras.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   13320
   Begin VB.CheckBox chkPlazo 
      Caption         =   "Largo Plazo"
      Height          =   255
      Left            =   4200
      TabIndex        =   52
      Top             =   2520
      Width           =   1815
   End
   Begin MSDataListLib.DataCombo cbRubro 
      Height          =   360
      Left            =   4920
      TabIndex        =   51
      Top             =   1560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ForeColor       =   16384
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
   Begin VB.Frame frameSubtotal 
      Caption         =   "Subtotal del Rubro"
      Height          =   1215
      Left            =   6960
      TabIndex        =   41
      Top             =   2520
      Width           =   6255
      Begin AIFCmp1.asxPowerBanner lblRubro 
         Height          =   375
         Left            =   4320
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         EndColor        =   32768
         FormatString    =   "asxPowerBanner1"
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
      Begin AIFCmp1.asxPowerButton cmdSacarSub 
         Height          =   450
         Left            =   5520
         TabIndex        =   44
         ToolTipText     =   "Desaparece cuadro"
         Top             =   600
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   794
         BorderStyle     =   4
         Picture         =   "frmCompras.frx":0442
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
      Begin AIFCmp1.asxFontLabel lblimporte 
         Height          =   240
         Left            =   2040
         TabIndex        =   43
         Top             =   780
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   423
         TextColor       =   12582912
         Caption         =   "asxFontLabel1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSize        =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   375
         Left            =   840
         TabIndex        =   46
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   154861569
         CurrentDate     =   39392
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   375
         Left            =   2880
         TabIndex        =   47
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   154861569
         CurrentDate     =   39392
      End
      Begin AIFCmp1.asxPowerButton cmdVerSubTot 
         Height          =   450
         Left            =   4800
         TabIndex        =   50
         ToolTipText     =   "Desaparece cuadro"
         Top             =   600
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   794
         BorderStyle     =   4
         Picture         =   "frmCompras.frx":09DC
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
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   49
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   48
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Subototal Rubro $:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   780
         Width           =   1695
      End
   End
   Begin AIFCmp1.asxPowerButton cmdSubtotal 
      Height          =   495
      Left            =   11640
      TabIndex        =   40
      ToolTipText     =   "Calcula el SubTotal de un Rubro en Debe"
      Top             =   2040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "Ver Subtotal Rubro"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   12582912
   End
   Begin MSDataListLib.DataCombo cbTipo 
      Height          =   360
      Left            =   8400
      TabIndex        =   10
      Top             =   960
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
   Begin VB.Frame Frame1 
      Caption         =   "Facturas Cargadas"
      Height          =   615
      Left            =   240
      TabIndex        =   35
      Top             =   2880
      Width           =   1695
      Begin VB.Label lblCuenta 
         Alignment       =   2  'Center
         Caption         =   "Label9"
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
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txtDeposito 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      MaxLength       =   10
      TabIndex        =   13
      Top             =   2520
      Width           =   1575
   End
   Begin AIFCmp1.asxPowerButton cmdCalc 
      Height          =   375
      Left            =   7560
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Calcular"
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
   Begin VB.TextBox txtDesc 
      Height          =   285
      Left            =   6240
      TabIndex        =   6
      ToolTipText     =   "Ingrese un porcentaje a descontar"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CheckBox chkDesc 
      Caption         =   "Efectuar Descuento"
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Frame frameBusFac 
      Caption         =   "Buscando Factura"
      Height          =   855
      Left            =   8280
      TabIndex        =   29
      Top             =   2760
      Width           =   4695
      Begin AIFCmp1.asxPowerButton cmdBusFac 
         Height          =   450
         Left            =   3240
         TabIndex        =   32
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   794
         Picture         =   "frmCompras.frx":0F76
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
      Begin VB.TextBox txtNumFac 
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
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   31
         Top             =   340
         Width           =   1695
      End
      Begin AIFCmp1.asxPowerButton asxPowerButton1 
         Height          =   450
         Left            =   3960
         TabIndex        =   33
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   794
         Picture         =   "frmCompras.frx":1290
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nº Factura:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   990
      End
   End
   Begin VB.ComboBox cbEstado 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frmCompras.frx":13EA
      Left            =   10800
      List            =   "frmCompras.frx":13F7
      Style           =   2  'Dropdown List
      TabIndex        =   11
      ToolTipText     =   "Indica si la factura se debe o fue pagada"
      Top             =   960
      Width           =   2415
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
      Left            =   8880
      MaxLength       =   100
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   1560
      Width           =   4335
   End
   Begin VB.Frame frameDatos 
      Caption         =   "Archivo de Carga"
      Height          =   4935
      Left            =   240
      TabIndex        =   26
      Top             =   3720
      Width           =   11055
      Begin MSDataGridLib.DataGrid dtgArchivo 
         Height          =   4575
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   8070
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
   End
   Begin AIFCmp1.asxPowerButton cmdGrabar 
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Picture         =   "frmCompras.frx":1412
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
      Height          =   285
      Left            =   2400
      MaxLength       =   10
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
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
      Height          =   285
      Left            =   2400
      MaxLength       =   15
      TabIndex        =   3
      ToolTipText     =   "Pruebe ingresar con codigo de barra"
      Top             =   1560
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   960
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
      Format          =   154927105
      CurrentDate     =   39393
   End
   Begin MSDataListLib.DataCombo dtcProveedor 
      Height          =   555
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   6975
      _ExtentX        =   12303
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
      Left            =   4200
      TabIndex        =   14
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Picture         =   "frmCompras.frx":1E24
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
      PictureOffsetX  =   10
   End
   Begin AIFCmp1.asxPowerButton cmdBorrar 
      Height          =   495
      Left            =   11520
      TabIndex        =   17
      Top             =   4560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Picture         =   "frmCompras.frx":2836
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
   Begin AIFCmp1.asxPowerButton cmdModificar 
      Height          =   495
      Left            =   11520
      TabIndex        =   18
      Top             =   5280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Picture         =   "frmCompras.frx":3248
      Caption         =   "&Modificar"
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
      Left            =   11520
      TabIndex        =   19
      Top             =   6000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Picture         =   "frmCompras.frx":37E2
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
   Begin AIFCmp1.asxPowerButton cmdSalir 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   11520
      TabIndex        =   21
      Top             =   8160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Picture         =   "frmCompras.frx":41F4
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
   Begin AIFCmp1.asxPowerButton cmdCalcular 
      Height          =   495
      Left            =   11520
      TabIndex        =   16
      Top             =   3840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Picture         =   "frmCompras.frx":478E
      Caption         =   "&Calcular Dpto."
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
   Begin AIFCmp1.asxPowerButton cmdBuscar 
      Height          =   495
      Left            =   11520
      TabIndex        =   20
      Top             =   6720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Picture         =   "frmCompras.frx":4AA8
      Caption         =   "&Buscar"
      CaptionAlignment=   5
      CaptionOffsetX  =   -13
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
      Left            =   9480
      TabIndex        =   37
      ToolTipText     =   "Agrega Proveedor a la lista"
      Top             =   240
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      BorderStyle     =   4
      Picture         =   "frmCompras.frx":54BA
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
   Begin AIFCmp1.asxPowerButton cmdFiltro 
      Height          =   495
      Left            =   11520
      TabIndex        =   38
      Top             =   7440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Picture         =   "frmCompras.frx":5ECC
      Caption         =   "&Filtrar Creditos"
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
      PictureOffsetX  =   5
   End
   Begin MSComCtl2.DTPicker dtpVto 
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   960
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
      Format          =   154927105
      CurrentDate     =   39393
   End
   Begin AIFCmp1.asxPowerButton cmdPlazos 
      Height          =   495
      Left            =   9840
      TabIndex        =   53
      ToolTipText     =   "Calcula el SubTotal de un Rubro en Debe"
      Top             =   2040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Picture         =   "frmCompras.frx":6466
      Caption         =   "Ver Plazos   "
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
      TextColor       =   12582912
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Vto:"
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
      Left            =   4080
      TabIndex        =   45
      Top             =   960
      Width           =   1125
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
      Left            =   4080
      TabIndex        =   39
      Top             =   1560
      Width           =   705
   End
   Begin VB.Label lblDeposito 
      AutoSize        =   -1  'True
      Caption         =   "Deposito:"
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
      TabIndex        =   34
      Top             =   2520
      Width           =   1020
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Estado:"
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
      Left            =   9840
      TabIndex        =   28
      Top             =   960
      Width           =   810
   End
   Begin VB.Label Label6 
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
      Left            =   7200
      TabIndex        =   27
      Top             =   1560
      Width           =   1650
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Left            =   240
      TabIndex        =   25
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
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
      Left            =   7200
      TabIndex        =   24
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Left            =   240
      TabIndex        =   23
      Top             =   1560
      Width           =   1785
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Left            =   240
      TabIndex        =   22
      Top             =   960
      Width           =   720
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
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsProv As New ADODB.Recordset
Private rsCompras As New ADODB.Recordset
Private rsTipFac As New ADODB.Recordset
Private rsSTRubro As New ADODB.Recordset
Private rsRubros As New ADODB.Recordset
Private banderaCred As Byte
Private vIdc As Long
Private Sub asxPowerButton1_Click()
frameBusFac.Visible = False
End Sub
Private Sub cbEstado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtImporte.SetFocus
End If
End Sub
Private Sub cbEstado_LostFocus()
If cbEstado.Text = "Credito" Then
    cbRubro.Text = "Fragancias"
Else
    cbRubro.Text = "Perfumeria"
End If
End Sub
Private Sub cbRubro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdGrabar.SetFocus
End If
End Sub

Private Sub cbtipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtNumero.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub chkPlazo_Click()
'Recalcula fecha de vencimiento por cobro a largo plazo

If chkPlazo.Value = 0 Then
    chkPlazo.Value = False
    dtpVto.Value = dtpFecha.Value
    txtObservaciones.Text = ""
    Exit Sub
End If

vfecfact = dtpFecha.Value
vFecVto = dtpVto.Value
frmLargoPlazo.Show vbModal

If rtaLargoPlazo = True Then
    dtpVto.Value = vFecVto
    txtObservaciones.Text = "LARGO PLAZO"
End If

End Sub

Private Sub cmdAgrPro_Click()
frmGestionProveedores.Show vbModal
rsProv.Requery
dtcProveedor.BoundText = rsProv!idproveedor
'dtcProveedor.Refresh
End Sub
Private Sub cmdFiltro_Click()

rsCompras.Close
If banderaCred = 0 Then
    rsCompras.Open "select * from facturascompras where idproveedor = " & dtcProveedor.BoundText & _
                    " and estado = 'C' order by idcompra desc", cn, adOpenDynamic, adLockOptimistic, adCmdText
    cmdFiltro.Caption = "Sacar Filtro"
    banderaCred = 1
Else
    rsCompras.Open "select * from facturascompras where idproveedor = " & dtcProveedor.BoundText & _
                    " order by idcompra desc", cn, adOpenDynamic, adLockOptimistic, adCmdText
    cmdFiltro.Caption = "Filtra Creditos"
    banderaCred = 0
End If

Set dtgarchivo.DataSource = rsCompras
dtgarchivo.Refresh

End Sub

Private Sub cmdPlazos_Click()
frmVerLargoPlazos.Show vbModal
End Sub

Private Sub cmdSacarSub_Click()
frameSubtotal.Visible = False
End Sub

Private Sub cmdSubtotal_Click()

If rsSTRubro.State = 1 Then
    rsSTRubro.Close
End If

frameBusFac.Visible = False
frameSubtotal.Visible = True
lblRubro.Visible = False

'Toma el primer y ultimo dia del mes en curso para las fecha de inicio y fin
dtpDesde.Value = "01/" & Month(Date) & "/" & Year(Date)

If Month(Date) = 1 Or Month(Date) = 3 Or Month(Date) = 5 Or Month(Date) = 7 Or Month(Date) = 8 Or Month(Date) = 10 Or Month(Date) = 12 Then
    dtpHasta.Value = Format("31/" & Month(Date) & "/" & Year(Date), "dd/mm/yyyy")
End If
If Month(Date) = 4 Or Month(Date) = 6 Or Month(Date) = 9 Or Month(Date) = 11 Then
    dtpHasta.Value = Format("30/" & Month(Date) & "/" & Year(Date), "dd/mm/yyyy")
End If
If Month(Date) = 2 Then
    dtpHasta.Value = Format("28/" & Month(Date) & "/" & Year(Date), "dd/mm/yyyy")
End If

Dim vTot As Double
vTot = 0
lblimporte.Caption = vTot

End Sub

Private Sub chkDesc_Click()
If Len(txtImporte.Text) = 0 Then
    MsgBox "DEBE INGRESAR UN IMPORTE PARA CALCULAR DESCUENTO...", vbExclamation, "ATENCION !"
    txtImporte.SetFocus
    chkDesc.Value = 0
    Exit Sub
End If
If IsNumeric(txtImporte.Text) = False Then
    MsgBox "LOS DATOS INGRESADOS EN EL CAMPO IMPORTE NO SON NUMERICOS !!!!", vbCritical, "ATENCION !"
    chkDesc.Value = 0
    txtImporte.SetFocus
    Exit Sub
End If
If chkDesc.Value = 1 Then
    txtDesc.Visible = True
    cmdCalc.Visible = True
    txtDesc.SetFocus
    SendKeys "{end}+{home}"
Else
    txtDesc.Visible = False
    cmdCalc.Visible = False
End If
End Sub
Private Sub cmdBorrar_Click()
If TempNivel = 0 Then
    MsgBox "NO POSEE IDENTIFICACION DE USUARIO PARA OPERAR...!", vbCritical, "Atención !"
    Exit Sub
End If
If rsCompras.RecordCount = 0 Then
    MsgBox "NO HAY REGISTROS PARA ELIMINAR !!!", vbExclamation, "ATENCION !"
    Exit Sub
End If
SioNo = MsgBox("ESTA SEGURO DE BORRAR ESTE REGISTRO ?", vbExclamation + vbYesNo, "ATENCION !")
If SioNo = vbYes Then
    rsCompras.Delete
    rsCompras.Update
    dtgarchivo.Refresh
End If
txtNumero.SetFocus

End Sub
Private Sub cmdBuscar_Click()
frameSubtotal.Visible = False
frameBusFac.Visible = True
txtNumFac.Text = ""
txtNumFac.SetFocus
End Sub
Private Sub cmdBusFac_Click()
If Len(txtNumFac.Text) = 0 Then
    MsgBox "DEBE INGRESAR DATOS EN EL CAMPO DE BUSQUEDA...", vbExclamation, "Atención !"
    txtNumero.SetFocus
    Exit Sub
End If
'If IsNumeric(txtNumFac.Text) = False Then
'    MsgBox "SOLO SE ADMITEN DIGITOS EN EL CAMPO BUSQUEDA....", vbCritical, "ATENCION !"
'    txtNumFac.SetFocus
'    SendKeys "{home}+{end}"
'    Exit Sub
'End If
rsCompras.Find "numero = " & Trim(txtNumFac.Text), , adSearchForward, 1
If rsCompras.EOF = True Then
    MsgBox "EL NUMERO DE COMPROBANTE SOLICITADO NO EXISTE ...!", vbCritical, "ATENCION !"
    rsCompras.MoveFirst
End If
frameBusFac.Visible = False
End Sub
Private Sub cmdCalc_Click()
Dim Descuento As Double
Dim Resul As Double
Dim vobser As String
If Len(txtDesc.Text) = 0 Or IsNumeric(txtDesc.Text) = False Then
    MsgBox "EL CAMPO PORCENTAJE DE DESCUENTO NO CONTIENE UN NUMERAL, VERIFIQUE...", vbCritical, "ATENCION !"
    txtDesc.SetFocus
    Exit Sub
End If
vobser = "Se ha descontado el " & txtDesc.Text & "% del Importe original de la factura de $" & txtImporte.Text
Descuento = (txtImporte * txtDesc) / 100
Resul = Round(txtImporte - Descuento, 2)
txtImporte = Resul
txtObservaciones.Text = vobser
chkDesc.Value = 0
txtDesc.Visible = False
cmdCalc.Visible = False
cmdGrabar.SetFocus
End Sub
Private Sub cmdCalcular_Click()
If TempNivel = 0 Then
    MsgBox "NO POSEE IDENTIFICACION DE USUARIO PARA OPERAR...!", vbCritical, "Atención !"
    Exit Sub
End If
vidPro = dtcProveedor.BoundText
frmCalculoDeposito.Show vbModal
rsCompras.Requery
Me.dtgarchivo.Refresh
End Sub
Private Sub cmdCancelar_Click()
txtNumero.Text = ""
txtImporte.Text = ""
txtObservaciones.Text = ""
cbTipo.Text = "A"
cbEstado.Text = "Debe"
txtDeposito.Text = ""
txtNumero.SetFocus
lblDeposito.Visible = False
txtDeposito.Visible = False
dtcProveedor.Enabled = True
cbRubro.Text = "Perfumeria"
vAgrega = True
End Sub
Private Sub cmdGrabar_Click()
If TempNivel = 0 Then
    MsgBox "NO POSEE IDENTIFICACION DE USUARIO PARA OPERAR...!", vbCritical, "Atención !"
    Exit Sub
End If
If Len(txtNumero.Text) = 0 Or Len(txtImporte.Text) = 0 Then
    MsgBox "FALTAN DATOS PARA GRABAR ...", vbInformation, "ATENCION !"
    txtNumero.SetFocus
    Exit Sub
End If
Err.Clear
On erro GoTo Solucion
'busca si hay duplicado de factura al agregar registro
If vAgrega = True Then
    rsCompras.Find "numero = " & Trim(txtNumero.Text), , adSearchForward, 1
    If rsCompras.EOF = False Then
        MsgBox "EL NUMERO DEL COMPROBANTE YA SE ENCUENTRA REGISTRADO, CONTROLE !", vbCritical, "ATENCION !"
        txtNumero.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
End If
If vAgrega = True Then
    rsCompras.AddNew
    lblCuenta.Caption = lblCuenta.Caption + 1 'cuentador de fact cargadas
End If
rsCompras!fecha = dtpFecha.Value
rsCompras!numero = txtNumero.Text
rsCompras!idproveedor = dtcProveedor.BoundText
If Len(txtDeposito.Text) = 0 Then
    rsCompras!depositado = Null
Else
    rsCompras!depositado = Str(txtDeposito.Text)
End If
rsCompras!tipo = cbTipo.Text
If cbTipo.Text = "NC" Then
    'grabo con signo negativo por ser nota de credito
    rsCompras!importe = Abs(txtImporte.Text) * -1
Else
    'grabo con signo positivo
    rsCompras!importe = Str(txtImporte.Text)
End If
rsCompras!observaciones = txtObservaciones.Text
If cbEstado.Text = "Debe" Then
    rsCompras!Estado = "D"
End If
If cbEstado.Text = "Pagado" Then
    rsCompras!Estado = "P"
End If
If cbEstado.Text = "Credito" Then
    rsCompras!Estado = "C"
End If
rsCompras!usuario = vUsu
rsCompras!rubro = cbRubro.Text
rsCompras!idRubro = cbRubro.BoundText
rsCompras!fechavto = dtpVto.Value

rsCompras.Update
rsCompras.Requery
dtgarchivo.Refresh
txtNumero.Text = ""
txtImporte.Text = ""
txtObservaciones.Text = ""
cbTipo.Text = "A"
cbEstado.Text = "Debe"
txtDeposito.Text = ""
cbRubro.Text = "Perfumeria"
txtNumero.SetFocus
dtcProveedor.Enabled = True
lblDeposito.Visible = False
txtDeposito.Visible = False
txtDesc.Text = ""
If vAgrega = False Then
    dtgarchivo.Bookmark = vIdc
    'rsCompras.Move (vIdc), 1
    'dtgArchivo.Refresh
End If
vAgrega = True

Exit Sub

Solucion:
   MsgBox Err.Number & "-" & Err.Description, vbInformation, "Error del Sistema ..."
   
End Sub

Private Sub cmdImprimir_Click()
If TempNivel = 0 Then
    MsgBox "NO POSEE IDENTIFICACION DE USUARIO PARA OPERAR...!", vbCritical, "Atención !"
    Exit Sub
End If
frmImprimirFacturas.Show vbModal
End Sub

Private Sub cmdModificar_Click()
If TempNivel = 0 Then
    MsgBox "NO POSEE IDENTIFICACION DE USUARIO PARA OPERAR...!", vbCritical, "Atención !"
    Exit Sub
End If
If rsCompras.RecordCount = 0 Then
    MsgBox "NO HAY REGISTROS PARA MODIFICAR ...", vbExclamation, "ATENCION !"
    Exit Sub
End If
vIdc = dtgarchivo.Bookmark
'vIdc = rsCompras!idcompra
lblDeposito.Visible = True
txtDeposito.Visible = True
dtcProveedor.BoundText = rsCompras!idproveedor
dtcProveedor.Enabled = False 'deshabilito el combo porque no se puede modificar el proveedor
dtpFecha.Value = rsCompras!fecha
txtNumero.Text = rsCompras!numero
txtImporte.Text = rsCompras!importe
'cbRubro.Text = rsCompras!rubro
cbRubro.BoundText = rsCompras!idRubro

If IsNull(rsCompras!fechavto) = True Then
    dtpVto.Value = dtpFecha.Value
Else
    dtpVto.Value = rsCompras!fechavto
End If
If IsNull(rsCompras!depositado) = True Then
    txtDeposito.Text = ""
Else
    txtDeposito.Text = rsCompras!depositado
End If
cbTipo.Text = rsCompras!tipo
If rsCompras!Estado = "D" Then
    cbEstado.Text = "Debe"
End If
If rsCompras!Estado = "P" Then
    cbEstado.Text = "Pagado"
End If
If rsCompras!Estado = "C" Then
    cbEstado.Text = "Credito"
End If
txtObservaciones.Text = rsCompras!observaciones
txtImporte.SetFocus
SendKeys "{home}+{end}"
vAgrega = False
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdVerSubTot_Click()
rsSTRubro.Open "select * from facturascompras where idproveedor = " & dtcProveedor.BoundText & " and estado = 'D' and rubro = '" & cbRubro.Text & "'" & _
                " and Fecha between #" & Format(dtpDesde.Value, "mm/dd/yyyy") & _
                "# and #" & Format(dtpHasta.Value, "mm/dd/yyyy") & "#", cn, adOpenDynamic, adLockReadOnly, adCmdText

If rsSTRubro.RecordCount > 0 Then
    rsSTRubro.MoveFirst
Else
    MsgBox "NO HAY INFORMACION PARA PARA TOTALES...!", vbCritical, "ATENCION !!!"
    frameSubtotal.Visible = False
    Exit Sub
End If
Do While rsSTRubro.EOF = False
    vTot = vTot + rsSTRubro!importe
    rsSTRubro.MoveNext
Loop
rsSTRubro.Close
lblimporte.Caption = vTot
lblRubro.Visible = True
lblRubro.FormatString = cbRubro.Text

End Sub

Private Sub dtcProveedor_Change()
If rsCompras.State = 1 Then
    rsCompras.Close
    Set rsCompras = Nothing
    
    rsCompras.Open "select * from facturascompras where idproveedor = " & dtcProveedor.BoundText & _
                " order by idcompra desc", cn, adOpenDynamic, adLockOptimistic, adCmdText

    Set dtgarchivo.DataSource = rsCompras
    dtgarchivo.Refresh
    lblCuenta.Caption = 0
End If
End Sub

Private Sub dtcProveedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    dtpFecha.SetFocus
End If
End Sub
Private Sub dtpDesde_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    dtpHasta.SetFocus
End If
End Sub

Private Sub dtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtNumero.SetFocus
    dtpVto.Value = dtpFecha.Value
End If
End Sub

Private Sub dtpHasta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdVerSubTot.SetFocus
End If

End Sub

Private Sub dtpVto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cbRubro.SetFocus
End If
End Sub

Private Sub Form_Load()
Me.Top = 50
Me.Left = 0
frameSubtotal.Visible = False
lblCuenta.Caption = 0
vAgrega = True
dtpFecha.Value = Date - 1
dtpVto.Value = dtpFecha.Value
banderaCred = 0

rsProv.Open "select * from proveedores order by nombre", cn, adOpenDynamic, adLockReadOnly, adCmdText

'llena el combo de proveedores
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

'llena el combo de tipo de Rubros
rsRubros.Open "select idrubro, rubro from Rubros order by rubro", cn, adOpenDynamic, adLockReadOnly, adCmdText
Set cbRubro.DataSource = rsRubros
Set cbRubro.RowSource = rsRubros
cbRubro.ListField = "Rubro"
cbRubro.BoundColumn = "idRubro"
'cbrubro.BoundText = 1

'llena el grid de facturas de compras
rsCompras.Open "select * from facturascompras where idproveedor = " & dtcProveedor.BoundText & _
                " order by idcompra desc", cn, adOpenDynamic, adLockOptimistic, adCmdText

Set dtgarchivo.DataSource = rsCompras
dtgarchivo.Refresh

cbTipo.Text = "A"
cbEstado.Text = "Debe"
cbRubro.Text = "Perfumeria"
frameBusFac.Visible = False
txtDesc.Visible = False
cmdCalc.Visible = False
lblDeposito.Visible = False
txtDeposito.Visible = False
End Sub
Private Sub Form_Unload(cancel As Integer)
If rsProv.State = 1 Then
    rsProv.Close
    Set rsProv = Nothing
End If
If rsCompras.State = 1 Then
    rsCompras.Close
    Set rsCompras = Nothing
End If
If rsTipFac.State = 1 Then
    rsTipFac.Close
    Set rsTipFac = Nothing
End If

If rsRubros.State = 1 Then
    rsRubros.Close
    Set rsRubros = Nothing
End If

vUsu = ""
TempNivel = 0
End Sub
Private Sub txtDeposito_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        cbRubro.SetFocus
    Case 44
        KeyAscii = 0
    Case 45
    Case 46
        KeyAscii = 44
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub txtDesc_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        If Len(txtDesc.Text) = 0 Or IsNumeric(txtDesc.Text) = False Then
            MsgBox "EL CAMPO PORCENTAJE DE DESCUENTO NO CONTIENE UN NUMERAL, VERIFIQUE...", vbCritical, "ATENCION !"
            txtDesc.SetFocus
            Exit Sub
        End If
        
        If (txtDesc.Text) > 0 Then
            Me.cbRubro.Text = "Medicamentos"
        End If
        cmdCalc.SetFocus
    Case 44
        KeyAscii = 0
    Case 45
    Case 46
        KeyAscii = 44
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub txtImporte_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        If dtcProveedor.BoundText = "2" Then
            Me.cbRubro.Text = "Medicamentos"
        End If
        KeyAscii = 0
        dtpVto.SetFocus
    Case 44
        KeyAscii = 0
    Case 45
    Case 46
        KeyAscii = 44
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub txtImporte_LostFocus()
If dtcProveedor.BoundText = "2" Then
    Me.cbRubro.Text = "Medicamentos"
End If
If IsNumeric(txtImporte.Text) = False Then
    If Len(txtImporte.Text) = 0 Then
        Exit Sub
    End If
    MsgBox "SOLO SE ACEPTAN IMPORTES NUMERICOS....!", vbCritical, "ATENCION !"
    txtImporte.SetFocus
End If
End Sub
Private Sub txtNumero_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        txtImporte.SetFocus
        If Len(txtImporte.Text) > 0 Then
            SendKeys "{end}+{home}"
        End If
    Case 44
        KeyAscii = 0
    Case 45
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub txtNumero_LostFocus()
    'este chek se hizo para controlar que no se carguen facturas que no
    'correspondan a Suiza
    If dtcProveedor.BoundText = 1 Then
        Dim raya As String
        raya = InStr(txtNumero.Text, "-")
        If raya > 0 Then
            MsgBox "ES POSIBLE QUE LA FACTURA NO CORRESPONDA A ESTE PROVEEDOR, VERIFIQUE !", vbCritical, "ATENCION !!!!"
        End If
    End If

End Sub

Private Sub txtNumFac_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBusFac.SetFocus
End If
End Sub
Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdGrabar.SetFocus
End If
End Sub
