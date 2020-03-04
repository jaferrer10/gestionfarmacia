VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmGestionClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestion de Clientes ..."
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14160
   Icon            =   "frmGestionClientes.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9240
   ScaleWidth      =   14160
   Begin TabDlg.SSTab sstClientes 
      DragIcon        =   "frmGestionClientes.frx":0442
      Height          =   9015
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   15901
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Archivo de Clientes"
      TabPicture(0)   =   "frmGestionClientes.frx":6C94
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frameLista"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frameDatos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Cuenta Corriente"
      TabPicture(1)   =   "frmGestionClientes.frx":76A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameCuentas"
      Tab(1).Control(1)=   "frameRegistro"
      Tab(1).Control(2)=   "lblNomCliente"
      Tab(1).Control(3)=   "lblDeuda"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Registro de Presión"
      TabPicture(2)   =   "frmGestionClientes.frx":80B8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Image1"
      Tab(2).Control(1)=   "lblCtePres"
      Tab(2).Control(2)=   "frameListaPresion"
      Tab(2).Control(3)=   "framePresion"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Pesaje"
      TabPicture(3)   =   "frmGestionClientes.frx":8652
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblNombres"
      Tab(3).Control(1)=   "Frame2"
      Tab(3).Control(2)=   "Frame3"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Medicación"
      TabPicture(4)   =   "frmGestionClientes.frx":9064
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "asxPowerBanner1"
      Tab(4).Control(1)=   "Frame4"
      Tab(4).Control(2)=   "Frame5"
      Tab(4).ControlCount=   3
      Begin VB.Frame Frame5 
         Caption         =   "Datos del Registro"
         Height          =   2415
         Left            =   -74760
         TabIndex        =   97
         Top             =   6000
         Width           =   13455
         Begin VB.TextBox Text4 
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
            Left            =   240
            MaxLength       =   50
            TabIndex        =   101
            Top             =   1440
            Width           =   5055
         End
         Begin VB.TextBox Text3 
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
            Left            =   4680
            MaxLength       =   6
            TabIndex        =   100
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox Text2 
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
            Left            =   3240
            MaxLength       =   6
            TabIndex        =   99
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox Text1 
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
            Left            =   6240
            MaxLength       =   6
            TabIndex        =   98
            Top             =   480
            Width           =   1215
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   375
            Left            =   2040
            TabIndex        =   102
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx20 
            Height          =   240
            Left            =   2040
            TabIndex        =   103
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   423
            Caption         =   "Hora"
         End
         Begin AIFCmp1.asxPowerButton asxPowerButton6 
            Height          =   495
            Left            =   11880
            TabIndex        =   104
            Top             =   1200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Picture         =   "frmGestionClientes.frx":95FE
            Caption         =   "Grabar"
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
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx21 
            Height          =   240
            Left            =   240
            TabIndex        =   105
            Top             =   1080
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   423
            Caption         =   "Observaciones"
         End
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx22 
            Height          =   240
            Left            =   3240
            TabIndex        =   106
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   423
            Caption         =   "Peso"
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   240
            TabIndex        =   107
            Top             =   480
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
            Format          =   135397377
            CurrentDate     =   39317
         End
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx23 
            Height          =   240
            Left            =   240
            TabIndex        =   108
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   423
            Caption         =   "Fecha"
         End
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx24 
            Height          =   240
            Left            =   4680
            TabIndex        =   109
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   423
            Caption         =   "% Grasa Corp"
         End
         Begin AIFCmp1.asxPowerButton asxPowerButton7 
            Height          =   495
            Left            =   11880
            TabIndex        =   110
            Top             =   1800
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Picture         =   "frmGestionClientes.frx":A010
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
            PictureOffsetX  =   10
         End
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx25 
            Height          =   240
            Left            =   6240
            TabIndex        =   111
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   423
            Caption         =   "% Liquido"
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Archivo de Pesajes"
         Height          =   4575
         Left            =   -74760
         TabIndex        =   91
         Top             =   1320
         Width           =   13455
         Begin AIFCmp1.asxPowerButton asxPowerButton1 
            Height          =   735
            Left            =   11040
            TabIndex        =   92
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   1296
            Picture         =   "frmGestionClientes.frx":AA22
            Caption         =   "Agregar Registro"
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
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   3975
            Left            =   120
            TabIndex        =   93
            Top             =   360
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   7011
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   49152
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
               DataField       =   "Hora"
               Caption         =   "Hora"
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
               DataField       =   "peso"
               Caption         =   "Peso"
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
               DataField       =   "grasa"
               Caption         =   "%Grasa"
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
               DataField       =   "liquido"
               Caption         =   "%Liquido"
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   1230,236
               EndProperty
               BeginProperty Column01 
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1184,882
               EndProperty
               BeginProperty Column04 
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   5385,26
               EndProperty
            EndProperty
         End
         Begin AIFCmp1.asxPowerButton asxPowerButton3 
            Height          =   735
            Left            =   11040
            TabIndex        =   94
            Top             =   1320
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   1296
            Picture         =   "frmGestionClientes.frx":B434
            Caption         =   "Modificar datos"
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
         Begin AIFCmp1.asxPowerButton asxPowerButton4 
            Height          =   735
            Left            =   11040
            TabIndex        =   95
            Top             =   2280
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   1296
            Picture         =   "frmGestionClientes.frx":BE46
            Caption         =   "Eliminar Registro"
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
         Begin AIFCmp1.asxPowerButton asxPowerButton5 
            Height          =   735
            Left            =   11040
            TabIndex        =   96
            Top             =   3240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   1296
            Picture         =   "frmGestionClientes.frx":C858
            Caption         =   "Imprimir Registros"
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
      End
      Begin VB.Frame Frame3 
         Caption         =   "Archivo de Pesajes"
         Height          =   4575
         Left            =   -74760
         TabIndex        =   73
         Top             =   1620
         Width           =   13215
         Begin AIFCmp1.asxPowerButton cmdAgregaP 
            Height          =   735
            Left            =   10560
            TabIndex        =   49
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   1296
            Picture         =   "frmGestionClientes.frx":D26A
            Caption         =   "Agregar Registro"
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
         Begin MSDataGridLib.DataGrid dtgPesaje 
            Height          =   3975
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   7011
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
               DataField       =   "Hora"
               Caption         =   "Hora"
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
               DataField       =   "peso"
               Caption         =   "Peso"
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
               DataField       =   "grasa"
               Caption         =   "%Grasa"
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
               DataField       =   "liquido"
               Caption         =   "%Liquido"
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   1230,236
               EndProperty
               BeginProperty Column01 
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1184,882
               EndProperty
               BeginProperty Column04 
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   5385,26
               EndProperty
            EndProperty
         End
         Begin AIFCmp1.asxPowerButton cmdModiP 
            Height          =   735
            Left            =   10560
            TabIndex        =   50
            Top             =   1320
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   1296
            Picture         =   "frmGestionClientes.frx":DC7C
            Caption         =   "Modificar datos"
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
         Begin AIFCmp1.asxPowerButton cmdElimP 
            Height          =   735
            Left            =   10560
            TabIndex        =   78
            Top             =   2280
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   1296
            Picture         =   "frmGestionClientes.frx":E68E
            Caption         =   "Eliminar Registro"
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
         Begin AIFCmp1.asxPowerButton cmdImpPeso 
            Height          =   735
            Left            =   10560
            TabIndex        =   79
            Top             =   3240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   1296
            Picture         =   "frmGestionClientes.frx":F0A0
            Caption         =   "Imprimir Registros"
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
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos del Registro"
         Height          =   2055
         Left            =   -74760
         TabIndex        =   67
         Top             =   6300
         Width           =   10095
         Begin VB.TextBox txtLiquido 
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
            Left            =   6240
            MaxLength       =   6
            TabIndex        =   55
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtPeso 
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
            Left            =   3240
            MaxLength       =   6
            TabIndex        =   53
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtGrasa 
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
            Left            =   4680
            MaxLength       =   6
            TabIndex        =   54
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtObserP 
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
            Left            =   240
            MaxLength       =   50
            TabIndex        =   56
            Top             =   1440
            Width           =   5055
         End
         Begin MSMask.MaskEdBox MskHoraP 
            Height          =   375
            Left            =   2040
            TabIndex        =   52
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx13 
            Height          =   240
            Left            =   2040
            TabIndex        =   68
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   423
            Caption         =   "Hora"
         End
         Begin AIFCmp1.asxPowerButton cmdGrabarP 
            Height          =   495
            Left            =   8160
            TabIndex        =   57
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Picture         =   "frmGestionClientes.frx":FAB2
            Caption         =   "Grabar"
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
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx14 
            Height          =   240
            Left            =   240
            TabIndex        =   69
            Top             =   1080
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   423
            Caption         =   "Observaciones"
         End
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx15 
            Height          =   240
            Left            =   3240
            TabIndex        =   70
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   423
            Caption         =   "Peso"
         End
         Begin MSComCtl2.DTPicker dtpFechaP 
            Height          =   375
            Left            =   240
            TabIndex        =   51
            Top             =   480
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
            Format          =   137428993
            CurrentDate     =   39317
         End
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx16 
            Height          =   240
            Left            =   240
            TabIndex        =   71
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   423
            Caption         =   "Fecha"
         End
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx17 
            Height          =   240
            Left            =   4680
            TabIndex        =   72
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   423
            Caption         =   "% Grasa Corp"
         End
         Begin AIFCmp1.asxPowerButton cmdCancelarP 
            Height          =   495
            Left            =   8160
            TabIndex        =   58
            Top             =   1200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Picture         =   "frmGestionClientes.frx":104C4
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
            PictureOffsetX  =   10
         End
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx18 
            Height          =   240
            Left            =   6240
            TabIndex        =   74
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   423
            Caption         =   "% Liquido"
         End
      End
      Begin AIFCmp1.asxPowerBanner lblDeuda 
         Height          =   615
         Left            =   -67680
         Top             =   1200
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1085
         EndColor        =   16744703
         FormatString    =   "Debe = $"
         Orientation     =   0
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
      Begin VB.Frame framePresion 
         Caption         =   "Datos del Registro"
         Height          =   2175
         Left            =   -74880
         TabIndex        =   61
         Top             =   6540
         Width           =   10095
         Begin MSMask.MaskEdBox mskHora 
            Height          =   375
            Left            =   2040
            TabIndex        =   42
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx12 
            Height          =   240
            Left            =   2040
            TabIndex        =   66
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   423
            Caption         =   "Hora"
         End
         Begin AIFCmp1.asxPowerButton cmdGrabaPresion 
            Height          =   495
            Left            =   8160
            TabIndex        =   46
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Picture         =   "frmGestionClientes.frx":10ED6
            Caption         =   "Grabar"
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
         Begin VB.TextBox txtObserPre 
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
            Left            =   240
            MaxLength       =   50
            TabIndex        =   45
            Top             =   1440
            Width           =   5055
         End
         Begin VB.TextBox txtBaja 
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
            Left            =   5280
            MaxLength       =   6
            TabIndex        =   44
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtAlta 
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
            Left            =   3600
            MaxLength       =   6
            TabIndex        =   43
            Top             =   480
            Width           =   1335
         End
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx11 
            Height          =   240
            Left            =   240
            TabIndex        =   65
            Top             =   1080
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   423
            Caption         =   "Observaciones"
         End
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx9 
            Height          =   240
            Left            =   3600
            TabIndex        =   63
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   423
            Caption         =   "Presion Alta"
         End
         Begin MSComCtl2.DTPicker dtpPresion 
            Height          =   375
            Left            =   240
            TabIndex        =   41
            Top             =   480
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
            Format          =   137232385
            CurrentDate     =   39317
         End
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx8 
            Height          =   240
            Left            =   240
            TabIndex        =   62
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   423
            Caption         =   "Fecha"
         End
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx10 
            Height          =   240
            Left            =   5280
            TabIndex        =   64
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   423
            Caption         =   "Presion Baja"
         End
         Begin AIFCmp1.asxPowerButton cmdCancelaPresion 
            Height          =   495
            Left            =   8160
            TabIndex        =   47
            Top             =   1320
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Picture         =   "frmGestionClientes.frx":13258
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
            PictureOffsetX  =   10
         End
      End
      Begin VB.Frame frameListaPresion 
         Caption         =   "Archivo de Presion"
         Height          =   4455
         Left            =   -74880
         TabIndex        =   39
         Top             =   2040
         Width           =   13215
         Begin AIFCmp1.asxPowerButton cmdAgrPre 
            Height          =   735
            Left            =   10560
            TabIndex        =   59
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   1296
            Picture         =   "frmGestionClientes.frx":13C6A
            Caption         =   "Agregar Registro"
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
         Begin MSDataGridLib.DataGrid dtgPresion 
            Height          =   3975
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   7011
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
            ColumnCount     =   5
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
               DataField       =   "Hora"
               Caption         =   "Hora"
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
               DataField       =   "alta"
               Caption         =   "Alta"
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
               DataField       =   "baja"
               Caption         =   "Baja"
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   1230,236
               EndProperty
               BeginProperty Column01 
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1184,882
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   5385,26
               EndProperty
            EndProperty
         End
         Begin AIFCmp1.asxPowerButton cmdModPre 
            Height          =   735
            Left            =   10560
            TabIndex        =   60
            Top             =   1320
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   1296
            Picture         =   "frmGestionClientes.frx":1467C
            Caption         =   "Modificar datos"
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
         Begin AIFCmp1.asxPowerButton cmdEliPre 
            Height          =   735
            Left            =   10560
            TabIndex        =   76
            Top             =   2280
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   1296
            Picture         =   "frmGestionClientes.frx":1508E
            Caption         =   "Eliminar Registro"
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
         Begin AIFCmp1.asxPowerButton cmdImpPresion 
            Height          =   735
            Left            =   10560
            TabIndex        =   77
            Top             =   3240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   1296
            Picture         =   "frmGestionClientes.frx":15AA0
            Caption         =   "Imprimir Registros"
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
      End
      Begin AIFCmp1.asxPowerBanner lblNomCliente 
         Height          =   615
         Left            =   -74760
         Top             =   1200
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   1085
         FormatString    =   "asxPowerBanner1"
         Orientation     =   0
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
      Begin VB.Frame frameRegistro 
         Caption         =   "Datos del Registro"
         Height          =   2655
         Left            =   -74880
         TabIndex        =   31
         Top             =   6240
         Width           =   13455
         Begin VB.TextBox txtBarras 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   375
            Left            =   1920
            MaxLength       =   20
            TabIndex        =   82
            ToolTipText     =   "Pase el código de Barras del Producto"
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtConcepto 
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
            Left            =   3960
            TabIndex        =   83
            Top             =   600
            Width           =   6375
         End
         Begin MSMask.MaskEdBox txtPrecio 
            Height          =   375
            Left            =   1800
            TabIndex        =   85
            Top             =   1320
            Width           =   1095
            _ExtentX        =   1931
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
            Format          =   "$#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtObser 
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
            TabIndex        =   88
            Top             =   2160
            Width           =   8775
         End
         Begin VB.TextBox txtDescuento 
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
            Left            =   3240
            TabIndex        =   86
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox txtCantidad 
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
            TabIndex        =   84
            Top             =   1320
            Width           =   1455
         End
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx7 
            Height          =   240
            Left            =   120
            TabIndex        =   38
            Top             =   1800
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   423
            Caption         =   "Observaciones"
         End
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx6 
            Height          =   240
            Left            =   4560
            TabIndex        =   37
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   423
            Caption         =   "Importe"
         End
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx5 
            Height          =   240
            Left            =   3240
            TabIndex        =   36
            Top             =   1080
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   423
            Caption         =   "Desc.%"
         End
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx4 
            Height          =   240
            Left            =   1800
            TabIndex        =   35
            Top             =   1080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   423
            Caption         =   "Precio"
         End
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx3 
            Height          =   240
            Left            =   120
            TabIndex        =   34
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   423
            Caption         =   "Cant."
         End
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx2 
            Height          =   240
            Left            =   3960
            TabIndex        =   33
            Top             =   360
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   423
            Caption         =   "Concepto"
         End
         Begin MSComCtl2.DTPicker dtpRegistro 
            Height          =   375
            Left            =   120
            TabIndex        =   81
            Top             =   600
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
            Format          =   137232385
            CurrentDate     =   39311
         End
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx1 
            Height          =   240
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   423
            Caption         =   "Fecha"
         End
         Begin AIFCmp1.asxPowerButton cmdNo 
            Height          =   615
            Left            =   11040
            TabIndex        =   90
            Top             =   1800
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1085
            Picture         =   "frmGestionClientes.frx":164B2
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
            PictureOffsetX  =   10
         End
         Begin AIFCmp1.asxPowerButton cmdGuardar 
            Height          =   615
            Left            =   11040
            TabIndex        =   89
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1085
            Picture         =   "frmGestionClientes.frx":16EC4
            Caption         =   "Grabar"
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
            PictureAlignment=   6
         End
         Begin MSMask.MaskEdBox txtImporte 
            Height          =   375
            Left            =   4560
            TabIndex        =   87
            Top             =   1320
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "$#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx19 
            Height          =   240
            Left            =   1920
            TabIndex        =   80
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   423
            Caption         =   "Cod. Barras"
         End
      End
      Begin VB.Frame frameCuentas 
         Caption         =   "Archivo Cuenta Corriente"
         Height          =   4095
         Left            =   -74880
         TabIndex        =   30
         Top             =   2040
         Width           =   13455
         Begin AIFCmp1.asxPowerButton cmdNuevo 
            Height          =   615
            Left            =   11040
            TabIndex        =   16
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1085
            Picture         =   "frmGestionClientes.frx":173B0
            Caption         =   "Agregar"
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
         Begin MSDataGridLib.DataGrid dtgListaCta 
            Height          =   3615
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   6376
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
            ColumnCount     =   8
            BeginProperty Column00 
               DataField       =   "idcliente"
               Caption         =   "IdCliente"
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
            BeginProperty Column02 
               DataField       =   "concepto"
               Caption         =   "Concepto"
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
            BeginProperty Column05 
               DataField       =   "descuento"
               Caption         =   "Desc.%"
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
               DataField       =   "Importe"
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
            BeginProperty Column07 
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
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   4
               BeginProperty Column00 
                  ColumnWidth     =   734,74
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1275,024
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   4619,906
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   764,787
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   945,071
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   780,095
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1184,882
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   6224,882
               EndProperty
            EndProperty
         End
         Begin AIFCmp1.asxPowerButton cmdCambiar 
            Height          =   615
            Left            =   11040
            TabIndex        =   17
            Top             =   1080
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1085
            Picture         =   "frmGestionClientes.frx":17DC2
            Caption         =   "Modificar"
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
         Begin AIFCmp1.asxPowerButton cmdBorraRegCta 
            Height          =   615
            Left            =   11040
            TabIndex        =   18
            Top             =   1800
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1085
            Picture         =   "frmGestionClientes.frx":187D4
            Caption         =   "Eliminar"
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
         Begin AIFCmp1.asxPowerButton cmdBorraTodo 
            Height          =   615
            Left            =   11040
            TabIndex        =   19
            Top             =   3240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1085
            Picture         =   "frmGestionClientes.frx":191E6
            Caption         =   "Borrar Cuenta"
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
         Begin AIFCmp1.asxPowerButton cmdImprimirCta 
            Height          =   615
            Left            =   11040
            TabIndex        =   75
            Top             =   2520
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1085
            Picture         =   "frmGestionClientes.frx":19638
            Caption         =   "Imprimir Cuenta"
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
      End
      Begin VB.Frame frameDatos 
         Caption         =   "Datos del Cliente "
         Height          =   2895
         Left            =   120
         TabIndex        =   23
         Top             =   5640
         Width           =   13695
         Begin VB.TextBox txtOs 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   6840
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   113
            Top             =   1320
            Width           =   3975
         End
         Begin VB.TextBox txtNombre 
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
            Left            =   1440
            MaxLength       =   30
            TabIndex        =   7
            Top             =   330
            Width           =   3975
         End
         Begin VB.TextBox txtApellido 
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
            Left            =   6840
            MaxLength       =   30
            TabIndex        =   8
            Top             =   360
            Width           =   3975
         End
         Begin VB.TextBox txtTelefono 
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
            Left            =   1440
            MaxLength       =   30
            TabIndex        =   9
            Top             =   840
            Width           =   3975
         End
         Begin VB.TextBox txtDireccion 
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
            Left            =   6840
            MaxLength       =   50
            TabIndex        =   10
            Top             =   840
            Width           =   3975
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
            Height          =   855
            Left            =   1920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   1800
            Width           =   4815
         End
         Begin MSComCtl2.DTPicker dtpNac 
            Height          =   375
            Left            =   2640
            TabIndex        =   11
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
            Format          =   137232385
            CurrentDate     =   39303
         End
         Begin AIFCmp1.asxPowerButton cmdCancelar 
            Height          =   495
            Left            =   11760
            TabIndex        =   14
            Top             =   2040
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            Picture         =   "frmGestionClientes.frx":1A04A
            Caption         =   "&Cancelar"
            CaptionAlignment=   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PictureAlignment=   3
            PictureOffsetX  =   10
         End
         Begin AIFCmp1.asxPowerButton cmdGrabar 
            Height          =   495
            Left            =   11760
            TabIndex        =   13
            Top             =   1320
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            Picture         =   "frmGestionClientes.frx":1AA5C
            Caption         =   "&Gabar "
            CaptionAlignment=   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PictureAlignment=   3
            PictureOffsetX  =   10
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Datos Obra Social:"
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
            Left            =   4800
            TabIndex        =   112
            Top             =   1320
            Width           =   1980
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
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
            TabIndex        =   29
            Top             =   360
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Apellidos:"
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
            Left            =   5640
            TabIndex        =   28
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Telefonos:"
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
            TabIndex        =   27
            Top             =   960
            Width           =   1125
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Direccion:"
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
            Left            =   5640
            TabIndex        =   26
            Top             =   960
            Width           =   1065
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Nacimiento:"
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
            Top             =   1320
            Width           =   2280
         End
         Begin VB.Label Label7 
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
            Left            =   240
            TabIndex        =   24
            Top             =   1800
            Width           =   1650
         End
      End
      Begin VB.Frame frameLista 
         Caption         =   "Listado de Clientes"
         Height          =   3975
         Left            =   120
         TabIndex        =   22
         Top             =   1560
         Width           =   13695
         Begin MSDataGridLib.DataGrid dtgLista 
            Height          =   3615
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   6376
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16761024
            HeadLines       =   2
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
            ColumnCount     =   8
            BeginProperty Column00 
               DataField       =   "idcliente"
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
               DataField       =   "apellido"
               Caption         =   "Apellidos"
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
               DataField       =   "nombre"
               Caption         =   "Nombres"
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
               DataField       =   "telefono"
               Caption         =   "Telefono"
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
               DataField       =   "direccion"
               Caption         =   "Direccion"
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
               DataField       =   "fechanac"
               Caption         =   "Nacimiento"
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
               DataField       =   "ObraSocial"
               Caption         =   "Datos Obra Social"
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
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               BeginProperty Column00 
                  ColumnWidth     =   599,811
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2039,811
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   2190,047
               EndProperty
               BeginProperty Column03 
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   3284,788
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1200,189
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   5369,953
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   7694,93
               EndProperty
            EndProperty
         End
         Begin AIFCmp1.asxPowerButton cmdEliminar 
            Height          =   495
            Left            =   11880
            TabIndex        =   6
            Top             =   1920
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Picture         =   "frmGestionClientes.frx":1B46E
            Caption         =   "&Eliminar"
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
         Begin AIFCmp1.asxPowerButton cmdModificar 
            Height          =   495
            Left            =   11880
            TabIndex        =   5
            Top             =   1200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Picture         =   "frmGestionClientes.frx":1BE80
            PictureDown     =   "frmGestionClientes.frx":1C892
            Caption         =   "&Modificar"
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
         Begin AIFCmp1.asxPowerButton cmdAgregar 
            Height          =   495
            Left            =   11880
            TabIndex        =   4
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Picture         =   "frmGestionClientes.frx":1D2A4
            Caption         =   "&Agregar"
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
      End
      Begin VB.Frame Frame1 
         Caption         =   "Busqueda de Cliente"
         Height          =   855
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   13695
         Begin VB.TextBox txtCliente 
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
            Left            =   2520
            MaxLength       =   30
            TabIndex        =   0
            Top             =   240
            Width           =   6735
         End
         Begin AIFCmp1.asxPowerButton asxPowerButton2 
            Cancel          =   -1  'True
            Height          =   570
            Left            =   11880
            TabIndex        =   2
            Top             =   195
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   1005
            Picture         =   "frmGestionClientes.frx":1DCB6
            Caption         =   "&Salir"
            CaptionAlignment=   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PictureAlignment=   3
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Apellido del Cliente:"
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
            TabIndex        =   21
            Top             =   360
            Width           =   2115
         End
      End
      Begin AIFCmp1.asxPowerBanner lblCtePres 
         Height          =   615
         Left            =   -74880
         Top             =   1200
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   1085
         EndColor        =   16744448
         FormatString    =   "asxPowerBanner1"
         Orientation     =   0
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
      Begin AIFCmp1.asxPowerBanner lblNombres 
         Height          =   615
         Left            =   -74760
         Top             =   780
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   1085
         EndColor        =   16744448
         FormatString    =   "asxPowerBanner1"
         Orientation     =   0
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
      Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
         Height          =   615
         Left            =   -74760
         Top             =   600
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   1085
         EndColor        =   32768
         Orientation     =   0
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
      Begin VB.Image Image1 
         Height          =   720
         Left            =   -66960
         Picture         =   "frmGestionClientes.frx":1E242
         Top             =   1140
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmGestionClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsBuscaCod As New ADODB.Recordset
Private rsClientes As New ADODB.Recordset
Private rsCuentas As New ADODB.Recordset
Private rsPresion As New ADODB.Recordset
Private rsPesaje As New ADODB.Recordset
Private rsMorosos As New ADODB.Recordset
Private rptPresiones As New crptPresiones
Private rptPesos As New crptPesajes
Private Sub asxPowerButton2_Click()
Unload Me
End Sub

Private Sub cmdAgregaP_Click()
vAgrega = True
Frame2.Visible = True
Frame3.Enabled = False
dtpFechaP.Value = Date
MskHoraP.Text = Format(Time, "hh:mm")
txtPeso.SetFocus
End Sub
Private Sub cmdAgregar_Click()
vAgrega = True
frameLista.Enabled = False
FrameDatos.Enabled = True
Call BlanqueaCampos
txtNombre.SetFocus
End Sub
Private Sub cmdAgrPre_Click()
framePresion.Visible = True
frameListaPresion.Enabled = False
dtpPresion.Value = Date
mskHora.Text = FormatDateTime(Time, vbShortTime)
mskHora.SetFocus
vAgrega = True
End Sub
Private Sub cmdBorraRegCta_Click()
If rsCuentas.RecordCount = 0 Then
    MsgBox "NO HAY REGISTROS PARA ELIMINAR ..!", vbInformation, "ATENCION !"
    Exit Sub
End If
If TempNivel = 1 Then
    SioNo = MsgBox("ESTA SEGURO DE BORRAR ESTE REGISTRO DE LA CUENTA ???", vbInformation + vbYesNo, "ATENCION !")
    If SioNo = vbYes Then
        rsCuentas.Delete
        rsCuentas.Requery
        Me.dtgListaCta.Refresh
        Call CalculaDeuda
    End If
Else
    MsgBox "SU NIVEL DE AUTORIZACION NO LE PERMITE ELIMINAR REGISTROS...", vbExclamation, "Seguridad del Sistema..."
End If
End Sub
Private Sub cmdBorraTodo_Click()
If TempNivel = 1 Then
    frmBorraCuenta.Show vbModal
    rsCuentas.Requery
    Me.dtgListaCta.Refresh
    Call CalculaDeuda
Else
    MsgBox "SU NIVEL DE AUTORIZACON NO LE PERMITE BORRAR CUENTAS...", vbExclamation, "Seguridad del Sistema..."
End If
End Sub
Private Sub cmdCambiar_Click()
If rsCuentas.RecordCount = 0 Then
    MsgBox "NO HAY DATOS PARA MODIFICAR !", vbExclamation, "ATENCION !"
    Exit Sub
End If
If TempNivel = 1 Then
    vAgrega = False
    frameCuentas.Enabled = False
    frameRegistro.Visible = True
    dtpRegistro.Value = rsCuentas!fecha
    txtConcepto.Text = rsCuentas!concepto
    txtCantidad.Text = rsCuentas!cantidad
    txtPrecio.Text = rsCuentas!precio
    txtDescuento.Text = rsCuentas!Descuento
    txtImporte.Text = rsCuentas!Importe
    txtObser.Text = rsCuentas!observaciones
    txtConcepto.SetFocus
    SendKeys "{home}+{end}"
Else
    MsgBox "SU NIVEL DE AUTORIZACON NO LE PERMITE MODIFICAR REGISTROS...", vbExclamation, "Seguridad del Sistema..."
End If
End Sub
Private Sub cmdCancelaPresion_Click()
txtAlta.Text = ""
txtBaja.Text = ""
txtObserPre.Text = ""
framePresion.Visible = False
frameListaPresion.Enabled = True
End Sub
Private Sub cmdCancelar_Click()
frameLista.Enabled = True
Call TomaDatos
FrameDatos.Enabled = False
End Sub

Private Sub cmdCancelarP_Click()
Frame2.Visible = False
Frame3.Enabled = True
dtgPesaje.SetFocus
End Sub

Private Sub cmdEliminar_Click()
If TempNivel > 1 Or TempNivel = 0 Then
    MsgBox "NO POSEE EL NIVEL DE AUTORIZACION PARA REALIZAR ESTA FUNCION...", vbExclamation, "Atencion !!!"
    Exit Sub
End If
If rsClientes!idcliente = 3 Then
    MsgBox "Cliente del Sistema, imposible borrar !", vbCritical, "ATENCION !"
    Exit Sub
End If
If rsClientes.RecordCount = 0 Then
    MsgBox "NO HAY DATOS PARA ELIMINAR ...!", vbCritical, "ATENCION !"
    dtgLista.SetFocus
    Exit Sub
End If
SioNo = MsgBox("ESTA SEGURO DE BORRAR ESTE CLIENTE ???", vbExclamation + vbYesNo, "ATENCION !!!")
If SioNo = vbYes Then
    rsClientes.Delete
    rsClientes.Update
    dtgLista.Refresh
    dtgLista.SetFocus
End If
End Sub
Private Sub cmdElimP_Click()
SioNo = MsgBox("ESTA SEGURO DE ELIMINAR ESTE REGISTRO DE PESAJE ?", vbExclamation + vbYesNo, "ATENCION !")
If SioNo = vbYes Then
    rsPesaje.Delete
    rsPesaje.Requery
    dtgPesaje.Refresh
End If
End Sub
Private Sub cmdEliPre_Click()
If rsPresion.RecordCount = 0 Then
    MsgBox "NO HAY REGISTROS PARA SER ELIMINADOS ...", vbCritical, "ATENCION !"
    Exit Sub
End If
SioNo = MsgBox("ESTA SEGURO DE ELIMINAR ESTE REGISTRO ???", vbExclamation + vbYesNo, "ATENCION !")
If SioNo = vbYes Then
    rsPresion.Delete
    rsPresion.Update
    dtgPresion.Refresh
End If
End Sub
Private Sub cmdGrabaPresion_Click()
If Len(txtAlta.Text) = 0 Or Len(txtBaja.Text) = 0 Or mskHora.Text = "__:__" Then
    MsgBox "FALTAN DATOS SOBRE LA TOMA DE PRESION ...!", vbExclamation, "ATENCION !"
    txtAlta.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If
Err.Clear
On Error GoTo Solucion
If vAgrega = True Then
    rsPresion.AddNew
End If
rsPresion!idcliente = rsClientes!idcliente
rsPresion!fecha = dtpPresion.Value
rsPresion!Hora = mskHora.Text
rsPresion!alta = txtAlta.Text
rsPresion!baja = txtBaja.Text
rsPresion!observaciones = txtObserPre.Text
rsPresion.Update
rsPresion.Requery
dtgPresion.Refresh
txtAlta.Text = ""
txtBaja.Text = ""
txtObserPre.Text = ""
framePresion.Visible = False
frameListaPresion.Enabled = True
Exit Sub
Solucion:
   MsgBox Err.Number & "-" & Err.Description, vbInformation, "Error del Sistema..."

End Sub

Private Sub cmdGrabar_Click()
If Len(txtNombre.Text) = 0 Then
    MsgBox "DEBE INGRESAR EL NOMBRE DEL CLIENTE ...", vbCritical, "ATENCION !"
    txtNombre.SetFocus
    Exit Sub
End If
If Len(txtApellido.Text) = 0 Then
    MsgBox "DEBE INGRESAR EL APELLIDO DEL CLIENTE ...", vbCritical, "ATENCION !"
    txtApellido.SetFocus
    Exit Sub
End If
SioNo = MsgBox("ESTA SEGURO DE GRABAR LOS DATOS DEL CLIENTE ?", vbInformation + vbYesNo, "ATENCION !")
If SioNo = vbYes Then
    If vAgrega = True Then
        rsClientes.AddNew
    End If
    rsClientes!nombre = txtNombre.Text
    rsClientes!apellido = txtApellido.Text
    rsClientes!telefono = txtTelefono.Text
    rsClientes!direccion = txtDireccion.Text
    rsClientes!fechanac = dtpNac.Value
    rsClientes!obrasocial = txtOs.Text
    rsClientes!observaciones = txtObservaciones.Text
    rsClientes.Update
    dtgLista.Refresh
End If
frameLista.Enabled = True
FrameDatos.Enabled = False
dtgLista.SetFocus
vAgrega = True
End Sub

Private Sub cmdGrabarP_Click()
Err.Clear
On Error GoTo Solucion
If vAgrega = True Then
    rsPesaje.AddNew
End If
rsPesaje!idcliente = rsClientes!idcliente
rsPesaje!fecha = dtpFechaP.Value
rsPesaje!Hora = MskHoraP.Text
rsPesaje!peso = Str(txtPeso.Text)
rsPesaje!grasa = Str(txtGrasa.Text)
rsPesaje!liquido = Str(txtLiquido.Text)
rsPesaje!observaciones = txtObserP.Text
rsPesaje.Update
dtgPesaje.Refresh
txtPeso.Text = ""
txtGrasa.Text = ""
txtLiquido.Text = ""
txtObserP.Text = ""
Frame2.Visible = False
Frame3.Enabled = True
Exit Sub
Solucion:
   MsgBox Err.Number & "-" & Err.Description, vbInformation, "Error del Sistema..."

End Sub

Private Sub cmdGuardar_Click()
If Len(txtConcepto.Text) = 0 Then
    MsgBox "DEBE INGRESAR EL CONCEPTO ...!", vbExclamation, "ATENCION !"
    txtConcepto.SetFocus
    Exit Sub
End If
If Len(txtCantidad.Text) = 0 Then
    MsgBox "DEBE INGRESAR LA CANTIDAD ...!", vbExclamation, "ATENCION !"
    txtCantidad.SetFocus
    Exit Sub
End If
If Len(txtPrecio.Text) = 0 Then
    MsgBox "Debe ingresar el precio del producto, se acepta cero...", vbExclamation, "Atención !"
    txtPrecio.SetFocus
    Exit Sub
End If
Err.Clear
On Error GoTo Solucion
If vAgrega = True Then
    rsCuentas.AddNew
End If
rsCuentas!fecha = dtpRegistro.Value
rsCuentas!idcliente = rsClientes!idcliente
rsCuentas!concepto = txtConcepto.Text
rsCuentas!cantidad = txtCantidad.Text
rsCuentas!precio = Str(txtPrecio.Text)
If Len(txtDescuento.Text) = 0 Then
    rsCuentas!Descuento = 0
Else
    rsCuentas!Descuento = txtDescuento.Text
End If
rsCuentas!Importe = Str(txtImporte.Text)
rsCuentas!observaciones = txtObser.Text
rsCuentas.Update
rsCuentas.Requery
dtgListaCta.Refresh
txtConcepto.Text = ""
txtCantidad.Text = ""
txtPrecio.Text = ""
txtDescuento.Text = ""
txtImporte.Text = ""
txtObser.Text = ""
frameRegistro.Visible = False
frameCuentas.Enabled = True
dtgListaCta.SetFocus
Call CalculaDeuda
Exit Sub
Solucion:
   MsgBox Err.Number & "-" & Err.Description, vbInformation, "Error del Sistema..."
   
End Sub

Private Sub cmdImpPeso_Click()
rptPesos.Database.SetDataSource rsPesaje
rptPesos.Text13.SetText rsClientes!nombre
rptPesos.Text14.SetText rsClientes!apellido
Set rptGeneral = rptPesos
frmVistaPrevia.Show
End Sub

Private Sub cmdImpPresion_Click()
rptPresiones.Database.SetDataSource rsPresion

rptPresiones.Text13.SetText rsClientes!nombre
rptPresiones.Text14.SetText rsClientes!apellido

Set rptGeneral = rptPresiones
frmVistaPrevia.Show
End Sub

Private Sub cmdImprimirCta_Click()
frmImprimeCtaCte.Show vbModal
End Sub
Private Sub cmdModificar_Click()
If TempNivel > 2 Then
    MsgBox "NO POSEE EL NIVEL DE AUTORIZACION PARA REALIZAR ESTA FUNCION...", vbExclamation, "Atencion !!!"
    Exit Sub
End If
If rsClientes.RecordCount = 0 Then
    MsgBox "NO HAY DATOS PARA MODIFICAR ...!", vbCritical, "ATENCION !"
    dtgLista.SetFocus
    Exit Sub
End If
frameLista.Enabled = False
FrameDatos.Enabled = True
vAgrega = False
Call TomaDatos
txtNombre.SetFocus
SendKeys "{home}+{end}"
End Sub
Private Sub cmdModiP_Click()
If rsPesaje.RecordCount = 0 Then
    MsgBox "NO HAY INFORMACION PARA MODIFICAR ...!"
    Exit Sub
End If
Frame2.Visible = True
Frame3.Enabled = False
dtpFechaP.Value = rsPesaje!fecha
MskHoraP.Text = Format(rsPesaje!Hora, "hh:mm")
txtPeso.Text = rsPesaje!peso
txtGrasa.Text = rsPesaje!grasa
txtLiquido.Text = rsPesaje!liquido
txtObserP.Text = rsPesaje!observaciones
txtPeso.SetFocus
vAgrega = False
End Sub
Private Sub cmdModPre_Click()
vAgrega = False
framePresion.Visible = True
frameListaPresion.Enabled = False
dtpPresion.Value = rsPresion!fecha
mskHora.Text = FormatDateTime(rsPresion!Hora, vbShortTime)
txtAlta.Text = rsPresion!alta
txtBaja.Text = rsPresion!baja
txtObserPre.Text = rsPresion!observaciones
mskHora.SetFocus
SendKeys "{home}+{end}"
End Sub
Private Sub cmdNo_Click()
txtConcepto.Text = ""
txtCantidad.Text = ""
txtPrecio.Text = ""
txtDescuento.Text = ""
txtImporte.Text = ""
txtObser.Text = ""
frameRegistro.Visible = False
frameCuentas.Enabled = True
dtgListaCta.SetFocus
End Sub
Private Sub cmdNuevo_Click()
frameCuentas.Enabled = False
frameRegistro.Visible = True
vAgrega = True
dtpRegistro.Value = Date
txtBarras.Text = ""
txtBarras.SetFocus
End Sub
Private Sub dtgLista_GotFocus()
dtgLista.Refresh
End Sub
Private Sub dtgLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
FrameDatos.Enabled = True
If frameLista.Enabled = True Then
    If rsClientes.EOF = False Then
        Call TomaDatos
    End If
End If
FrameDatos.Enabled = False
End Sub
Private Sub dtpNac_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtOs.SetFocus
End If
End Sub
Private Sub Form_Load()
Me.Top = 20
Me.Left = 50
FrameDatos.Enabled = False
frameLista.Enabled = True
'solapa cuenta corriente
frameRegistro.Visible = False
frameCuentas.Enabled = True
'solapa de registro de presion
frameListaPresion.Enabled = True
framePresion.Visible = False
If rsClientes.State = 1 Then
    rsClientes.Close
    Set rsClientes = Nothing
End If
rsClientes.Open "select * from clientes order by apellido", cn, adOpenDynamic, adLockOptimistic, adCmdText
Set dtgLista.DataSource = rsClientes
Call TomaDatos
sstClientes.Tab = 0

'controla que no haya morosos
rsMorosos.Open "select cl.Nombre,cl.Apellido, cl.Telefono, ct.fecha, ct.importe, ct.fecha from cuentascorrientes ct, clientes cl " & _
                " where ct.idcliente = cl.idcliente and " & _
                " (DateDiff('s', " & Date & ", ct.fecha))> '" & 30 & "' order by fecha", cn, adOpenDynamic, adLockOptimistic, adCmdText
                
If rsMorosos.RecordCount > 0 Then
    frmMorosos.Show vbModal
End If
rsMorosos.Close

'rsCuentas.MoveFirst
'Dim Moroso As Integer
'Dim vCteMoroso As Integer
'Do While Not rsCuentas.EOF
'    Moroso = DateDiff(DateInterval.Day, Date, rsCuentas!fecha)
'    If Moroso > 30 Then
'        vCteMoroso = rsCuentas!idcliente
'
'    End If
'Loop

End Sub
Private Sub BlanqueaCampos()
txtNombre.Text = ""
txtApellido.Text = ""
txtTelefono.Text = ""
txtDireccion.Text = ""
txtOs.Text = ""
txtObservaciones.Text = ""

End Sub
Private Sub Form_Unload(cancel As Integer)
If rsClientes.State = 1 Then
    rsClientes.Close
    Set rsClientes = Nothing
End If
If rsCuentas.State = 1 Then
    rsCuentas.Close
    Set rsCuentas = Nothing
End If
If rsPresion.State = 1 Then
    rsPresion.Close
    Set rsPresion = Nothing
End If
If rsPesaje.State = 1 Then
    rsPesaje.Close
    Set rsPesaje = Nothing
End If
vUsu = ""
TempNivel = 0
Exit Sub
End Sub
Private Sub mskHora_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtAlta.SetFocus
End If
End Sub
Private Sub MskHoraP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtPeso.SetFocus
End If
End Sub
Private Sub sstClientes_Click(PreviousTab As Integer)
'Solapa de cuentas corrientes
If sstClientes.Tab = 1 Then
    frameLista.Enabled = True
    vIdCliente = rsClientes!idcliente
    lblNomCliente.FormatString = rsClientes!nombre + " " + rsClientes!apellido
    If rsCuentas.State = 1 Then
        rsCuentas.Close
        Set rsCuentas = Nothing
    End If
    rsCuentas.Open "select * from cuentascorrientes where idcliente = " & rsClientes!idcliente & " order by fecha", cn, adOpenDynamic, adLockOptimistic, adCmdText
    Set dtgListaCta.DataSource = rsCuentas
    dtgListaCta.Refresh
    frameRegistro.Visible = False
    Me.frameCuentas.Enabled = True
    Call CalculaDeuda
End If
'solapa de mediciones de presion
If sstClientes.Tab = 2 Then
    lblCtePres.FormatString = rsClientes!nombre + " " + rsClientes!apellido
    If rsPresion.State = 1 Then
        rsPresion.Close
        Set rsPresion = Nothing
    End If
    rsPresion.Open "select * from presion where idcliente = " & rsClientes!idcliente & " order by fecha,hora", cn, adOpenDynamic, adLockOptimistic, adCmdText
    Set dtgPresion.DataSource = rsPresion
    dtgPresion.Refresh
End If
'solpa de mediciones de pesaje
If sstClientes.Tab = 3 Then
    Frame2.Visible = False
    lblNombres.FormatString = rsClientes!nombre + " " + rsClientes!apellido
    If rsPesaje.State = 1 Then
        rsPesaje.Close
        Set rsPesaje = Nothing
    End If
    rsPesaje.Open "select * from pesaje where idcliente = " & rsClientes!idcliente & " order by fecha,hora", cn, adOpenDynamic, adLockOptimistic, adCmdText
    Set dtgPesaje.DataSource = rsPesaje
    dtgPesaje.Refresh
End If

End Sub
Private Sub txtAlta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtBaja.SetFocus
End If
End Sub
Private Sub txtApellido_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    KeyAscii = 0
    txtTelefono.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub
Private Sub txtBaja_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtObserPre.SetFocus
End If
End Sub

Private Sub txtBarras_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtConcepto.SetFocus
End If
End Sub

Private Sub txtBarras_LostFocus()
If Len(txtBarras) = 0 Then Exit Sub
rsBuscaCod.Open "Select troquel, descripcion, precio from productos " & _
                "where troquel = '" & txtBarras.Text & "'", cn, adOpenDynamic, adLockReadOnly, adCmdText
If rsBuscaCod.RecordCount = 0 Then
    MsgBox "EL CODIGO DE BARRAS NO SE ENCUENTRA REGISTRADO !", vbCritical, "NO EXISTE EL PRODUCTO..."
    txtBarras.SetFocus
    SendKeys "{home}+{end}"
Else
    txtConcepto.Text = rsBuscaCod!descripcion
    txtCantidad.Text = 1
    txtCantidad.SetFocus
    SendKeys "{home}+{end}"
End If
rsBuscaCod.Close
Set rsBuscaCod = Nothing

End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        txtPrecio.SetFocus
        SendKeys "{end}+{home}"
    Case 44
        KeyAscii = 0
    Case 46
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub

Private Sub txtCliente_Change()
If Len(txtCliente.Text) = 0 Then Exit Sub
rsClientes.Find "apellido like '" & txtCliente.Text & "%'", , adSearchForward, 1
If rsClientes.EOF Then
    rsClientes.MoveFirst
End If
End Sub
Private Sub txtCliente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    Me.dtgLista.SetFocus
End If
End Sub
Private Sub txtCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    dtgLista.SetFocus
End If
End Sub
Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCantidad.SetFocus
End If
End Sub
Private Sub txtDescuento_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        txtObser.SetFocus
        SendKeys "{end}+{home}"
    Case 44
    Case 46
        KeyAscii = 44
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub

Private Sub txtDescuento_LostFocus()
If Len(txtDescuento.Text) = 0 Then Exit Sub
Dim subT, vpre, vDes As Single
vpre = CDbl(txtImporte.Text)
vDes = CDbl(txtDescuento.Text)
subT = (vpre * vDes) / 100

txtImporte.Text = Round((vpre - subT), 2)

End Sub
Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    dtpNac.SetFocus
End If
End Sub
Private Sub txtGrasa_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        txtLiquido.SetFocus
        SendKeys "{end}+{home}"
    Case 44
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
        KeyAscii = 0
        txtObser.SetFocus
        SendKeys "{end}+{home}"
    Case 44
        KeyAscii = 0
    Case 46
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub txtLiquido_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        txtObserP.SetFocus
        SendKeys "{end}+{home}"
    Case 44
    Case 45
    Case 46
        KeyAscii = 44
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    KeyAscii = 0
    txtApellido.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub
Private Sub txtObser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdGuardar.SetFocus
End If
End Sub
Private Sub txtObserP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdGrabarP.SetFocus
End If
End Sub

Private Sub txtObserPre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdGrabaPresion.SetFocus
End If
End Sub

Private Sub txtOs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtObservaciones.SetFocus
End If
End Sub

Private Sub txtPeso_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        txtGrasa.SetFocus
        SendKeys "{end}+{home}"
    Case 44
    Case 45
    Case 46
        KeyAscii = 44
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        txtDescuento.SetFocus
        SendKeys "{end}+{home}"
    Case 44
        KeyAscii = 0
    Case 45
    Case 46
        KeyAscii = 44
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub txtPrecio_LostFocus()
If Len(txtPrecio.Text) = 0 Or Len(txtCantidad.Text) = 0 Then Exit Sub
txtImporte.Text = CDbl(txtCantidad.Text) * CDbl(txtPrecio.Text)
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtDireccion.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub CalculaDeuda()
    Dim TotalDeuda As Double
    If rsCuentas.RecordCount = 0 Then
        lblDeuda.FormatString = " DEUDA = $ " & Round(TotalDeuda, 2)
        TotalDeuda = 0
        Exit Sub
    End If
    TotalDeuda = 0
    rsCuentas.MoveFirst
    Do While rsCuentas.EOF = False
        TotalDeuda = TotalDeuda + rsCuentas!Importe
        rsCuentas.MoveNext
    Loop
    lblDeuda.FormatString = " DEUDA = $ " & Round(TotalDeuda, 2)
End Sub

Private Sub TomaDatos()

If rsClientes.RecordCount = 0 Then
    txtNombre.Text = ""
    txtApellido.Text = ""
    txtTelefono.Text = ""
    txtDireccion.Text = ""
    txtOs.Text = ""
    txtObservaciones.Text = ""
    Exit Sub
End If

txtNombre.Text = rsClientes!nombre
txtApellido.Text = rsClientes!apellido
txtTelefono.Text = rsClientes!telefono
txtDireccion.Text = rsClientes!direccion
dtpNac.Value = rsClientes!fechanac
txtOs.Text = rsClientes!obrasocial & ""
txtObservaciones.Text = rsClientes!observaciones
End Sub
