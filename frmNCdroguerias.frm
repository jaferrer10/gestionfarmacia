VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Begin VB.Form frmNCdroguerias 
   Caption         =   "Registro de Notas de Creditos a Droguerias ..."
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10710
   Icon            =   "frmNCdroguerias.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10710
   StartUpPosition =   1  'CenterOwner
   Begin AIFCmp1.asxPowerButton cmdBus 
      Height          =   375
      Left            =   8760
      TabIndex        =   28
      Top             =   5520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "Busque"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   32768
   End
   Begin VB.TextBox txtBus 
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
      Height          =   285
      Left            =   8760
      MaxLength       =   20
      TabIndex        =   26
      Top             =   5160
      Width           =   1815
   End
   Begin AIFCmp1.asxLineHeaderEx lblBus 
      Height          =   240
      Left            =   8760
      TabIndex        =   25
      Top             =   4920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   423
      Caption         =   "Ingrese Codigo"
   End
   Begin AIFCmp1.asxPowerButton cmdSalir 
      Height          =   495
      Left            =   9000
      TabIndex        =   9
      Top             =   6240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Picture         =   "frmNCdroguerias.frx":6852
      Caption         =   "&Salir"
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
   Begin VB.Frame FrameBotones 
      Caption         =   "Operaciones"
      Height          =   4695
      Left            =   8760
      TabIndex        =   3
      Top             =   120
      Width           =   1815
      Begin AIFCmp1.asxPowerButton cmdAgregar 
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BorderStyle     =   5
         Picture         =   "frmNCdroguerias.frx":7264
         Caption         =   "&Agregar"
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
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmNCdroguerias.frx":77FE
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
      Begin AIFCmp1.asxPowerButton cmdBorrar 
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmNCdroguerias.frx":8210
         Caption         =   "Bo&rrar"
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
      Begin AIFCmp1.asxPowerButton cmdBuscar 
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmNCdroguerias.frx":87AA
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
      Begin AIFCmp1.asxPowerButton cmdFiltro 
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   3240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmNCdroguerias.frx":91BC
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
      Begin AIFCmp1.asxPowerButton cmdImp 
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmNCdroguerias.frx":9BCE
         Caption         =   "&Informe"
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
   End
   Begin VB.Frame frameEdicion 
      Caption         =   "Edicion"
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   8535
      Begin MSDataListLib.DataCombo dtcOs 
         Height          =   315
         Left            =   3600
         TabIndex        =   30
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcDrog 
         Height          =   315
         Left            =   6120
         TabIndex        =   24
         ToolTipText     =   "Seleccione el proveedor al cual se acreditara la NC"
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin AIFCmp1.asxPowerButton cmdOk 
         Height          =   375
         Left            =   6240
         TabIndex        =   22
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Picture         =   "frmNCdroguerias.frx":A5E0
         Caption         =   "Ok"
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
      Begin VB.TextBox txtObs 
         Height          =   285
         Left            =   1560
         MaxLength       =   49
         TabIndex        =   21
         Top             =   1320
         Width           =   4455
      End
      Begin VB.TextBox txtImp 
         Height          =   285
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   20
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtPer 
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   19
         ToolTipText     =   "Indique el periodo del NC con mes y año mm/aaaa"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtCod 
         Height          =   285
         Left            =   7200
         MaxLength       =   20
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtRes 
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   17
         ToolTipText     =   "Ingrese el nº de Resumen al cual se descuenta"
         Top             =   360
         Width           =   855
      End
      Begin AIFCmp1.asxPowerButton cmdCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   7320
         TabIndex        =   23
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Picture         =   "frmNCdroguerias.frx":AFF2
         Caption         =   "Cancel"
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Drogueria Asignada:"
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
         Left            =   4320
         TabIndex        =   16
         Top             =   840
         Width           =   1740
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
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
         TabIndex        =   15
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Periodo:"
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
         TabIndex        =   14
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Obra Social:"
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
         Left            =   2400
         TabIndex        =   13
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Codigo NC:"
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
         Left            =   6120
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
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
         Left            =   2400
         TabIndex        =   11
         Top             =   840
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Resumen:"
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
         TabIndex        =   10
         Top             =   360
         Width           =   1125
      End
   End
   Begin VB.Frame frameDatos 
      Caption         =   "Archivo"
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      Begin VB.Frame FrameFiltro 
         Caption         =   "Filtro de Datos"
         Height          =   1455
         Left            =   4800
         TabIndex        =   31
         Top             =   2520
         Width           =   3495
         Begin AIFCmp1.asxPowerButton cmdfil 
            Height          =   375
            Left            =   1320
            TabIndex        =   34
            Top             =   960
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            Caption         =   "Ver Filtro"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextColor       =   32768
         End
         Begin MSDataListLib.DataCombo dtcFiltro 
            Height          =   315
            Left            =   240
            TabIndex        =   33
            Top             =   480
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx1 
            Height          =   240
            Left            =   240
            TabIndex        =   32
            Top             =   240
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   423
            Caption         =   "Seleccione Obra Social"
         End
         Begin AIFCmp1.asxPowerButton cmdCancelFil 
            Height          =   375
            Left            =   2400
            TabIndex        =   35
            Top             =   960
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            Caption         =   "&Cancelar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextColor       =   255
         End
      End
      Begin MSDataGridLib.DataGrid dtgNC 
         Height          =   4335
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   7646
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "obsocial"
            Caption         =   "Ob.Social"
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
         BeginProperty Column01 
            DataField       =   "codigo"
            Caption         =   "Codigo"
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
            DataField       =   "periodo"
            Caption         =   "Periodo"
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
         BeginProperty Column03 
            DataField       =   "importe"
            Caption         =   "Importe"
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
         BeginProperty Column04 
            DataField       =   "resumen"
            Caption         =   "Resumen"
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
         BeginProperty Column05 
            DataField       =   "proveedor"
            Caption         =   "Proveedor"
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
            DataField       =   "observaciones"
            Caption         =   "Observaciones"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
               ColumnWidth     =   1305,071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1184,882
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   989,858
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1019,906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   780,095
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2700,284
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   4484,977
            EndProperty
         EndProperty
      End
   End
   Begin AIFCmp1.asxPowerButton cmdNoBusca 
      Height          =   375
      Left            =   9720
      TabIndex        =   29
      Top             =   5520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   192
   End
End
Attribute VB_Name = "frmNCdroguerias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsNC As New ADODB.Recordset
Private rsPro As New ADODB.Recordset
Private rsOs As New ADODB.Recordset
Private vidNc As Integer
Private Sub cmdAgregar_Click()
frameEdicion.Enabled = True
frameDatos.Enabled = False
FrameBotones.Enabled = False
vAgrega = True 'indica que el programa debe usar el addnew
txtRes.Text = ""
txtCod.Enabled = True
txtCod.Text = ""
txtPer.Text = ""
txtImp.Text = ""
txtObs.Text = ""
txtRes.SetFocus
End Sub
Private Sub cmdBorrar_Click()
If rsNC.RecordCount = 0 Then
    MsgBox "NO HAY REGISTROS PARA ELIMINAR !!! ", vbCritical, "ATENCION !"
    Exit Sub
End If
SioNo = MsgBox("ESTA SEGURO DE ELIMINAR ESTE REGISTRO DE NOTA DE CREDITO ??", vbExclamation + vbYesNo, "ELIMINANDO REGISTRO...")
If SioNo = vbYes Then
    Dim vNc 'variable solo para borrar el registro ya que el comando delete no funciona con la relacion
    vNc = rsNC!idnc
    cn.Execute "delete from ncdroguerias where idnc = " & vNc
    rsNC.Requery
    dtgNC.Refresh
    dtgNC.SetFocus
    Call TomaDatos
End If
End Sub
Private Sub cmdBus_Click()
lblBus.Visible = True
txtBus.Visible = True
cmdBus.Visible = True
cmdNoBusca.Visible = True
If Len(txtBus.Text) > 0 Then
    rsNC.Find "codigo = " & txtBus.Text, , adSearchForward, 1
    If rsNC.EOF = True Then
        MsgBox "EL CODIGO INGRESADO NO EXISTE...", vbExclamation, "RESULTADO DE LA BUSQUEDA ..."
    End If
End If
lblBus.Visible = False
txtBus.Visible = False
cmdBus.Visible = False
cmdNoBusca.Visible = False
dtgNC.SetFocus
End Sub
Private Sub cmdBuscar_Click()
lblBus.Visible = True
txtBus.Visible = True
cmdBus.Visible = True
cmdNoBusca.Visible = True
txtBus.SetFocus
SendKeys "{home}+{end}"
End Sub
Private Sub cmdCancel_Click()
frameEdicion.Enabled = False
frameDatos.Enabled = True
FrameBotones.Enabled = True
Me.dtgNC.SetFocus
End Sub

Private Sub cmdCancelFil_Click()
FrameFiltro.Visible = False
End Sub

Private Sub cmdfil_Click()
rsNC.Close
Set rsNC = Nothing
'abrimos nuevamente con el nuevo filtro
rsNC.Open "select o.nombre as ObSocial, n.codigo, n.periodo, n.importe, n.resumen, p.nombre as Proveedor, n.observaciones, n.Drogueria, n.idnc, n.osocial " & _
            "from (ncdroguerias n " & _
            "inner join obrasociales o on n.osocial = o.idos) " & _
            "inner join Proveedores p on n.drogueria = p.idproveedor " & _
            " where n.osocial = " & dtcFiltro.BoundText & _
            " order by n.drogueria,n.periodo,n.resumen", cn, adOpenDynamic, adLockOptimistic, adCmdText

If rsNC.RecordCount = 0 Then
    MsgBox "NO HAY INFORMACION PARA EL FILTRO SELECCIONADO...", vbExclamation, "Resultado filtro..."
    rsNC.Close
    rsNC.Open "select o.nombre as ObSocial, n.codigo, n.periodo, n.importe, n.resumen, p.nombre as Proveedor, n.observaciones, n.Drogueria, n.idnc, n.osocial " & _
            "from (ncdroguerias n " & _
            "inner join obrasociales o on n.osocial = o.idos) " & _
            "inner join Proveedores p on n.drogueria = p.idproveedor " & _
            " order by n.drogueria,n.periodo,n.resumen", cn, adOpenDynamic, adLockOptimistic, adCmdText
Else
    cmdFiltro.Caption = "&Sacar Filtro"
    cmdFiltro.TextColor = &HFF&
End If
Set dtgNC.DataSource = rsNC
dtgNC.Refresh
FrameFiltro.Visible = False
dtgNC.SetFocus
Call TomaDatos
End Sub

Private Sub cmdFiltro_Click()
If cmdFiltro.Caption = "&Filtrar Creditos" Then
    FrameFiltro.Visible = True
    dtcFiltro.SetFocus
Else
    rsNC.Close
    Set rsNC = Nothing
    rsNC.Open "select o.nombre as ObSocial, n.codigo, n.periodo, n.importe, n.resumen, p.nombre as Proveedor, n.observaciones, n.Drogueria, n.idnc, n.osocial " & _
            "from (ncdroguerias n " & _
            "inner join obrasociales o on n.osocial = o.idos) " & _
            "inner join Proveedores p on n.drogueria = p.idproveedor " & _
            " order by n.drogueria,n.periodo,n.resumen", cn, adOpenDynamic, adLockOptimistic, adCmdText

    Set dtgNC.DataSource = rsNC
    dtgNC.Refresh
    cmdFiltro.TextColor = &H80000008
    cmdFiltro.Caption = "&Filtrar Creditos"
End If

Call TomaDatos

End Sub

Private Sub cmdImp_Click()
frmInfNotasCreditos.Show vbModal
End Sub
Private Sub cmdModificar_Click()
If rsNC.RecordCount = 0 Then
    MsgBox "NO HAY REGISTROS PARA MODIFICAR...!", vbInformation, "ATENCION !"
    Exit Sub
End If
vAgrega = False
'Call TomaDatos
frameEdicion.Enabled = True
frameDatos.Enabled = False
FrameBotones.Enabled = False
txtCod.Enabled = False
txtRes.SetFocus
'SendKeys "{home}+{end}"
End Sub
Private Sub cmdNoBusca_Click()
lblBus.Visible = False
txtBus.Visible = False
cmdBus.Visible = False
cmdNoBusca.Visible = False
End Sub
Private Sub cmdok_Click()
Err.Clear
On Error GoTo Solucion
If Len(txtCod.Text) = 0 Then
    MsgBox "Debe ingresar el código del comprobante de la Nota de Credito...!", vbExclamation, "Atención !"
    txtCod.SetFocus
    Exit Sub
End If
If Len(dtcOs.Text) = 0 Or Len(txtImp.Text) = 0 Or Len(dtcDrog.Text) = 0 Then
    MsgBox "FALTAN INGRESAR DATOS PARA GRABAR EL REGISTRO ....!", vbCritical, "ATENCION !"
    txtRes.SetFocus
    Exit Sub
End If
If vAgrega = True Then
    rsNC.Find "codigo = " & txtCod.Text, , adSearchForward, 1
    If rsNC.EOF = False Then
        MsgBox "EL CODIGO DE NOTA DE CREDITO YA ESTA REGISTRADO, VERIFIQUE...", vbCritical, "DUPLICADO !"
        txtRes.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    'usamos el comando insert into sobre la tabla porque joden las relaciones
    strSQL = "insert into NCdroguerias (osocial, codigo, periodo, importe, resumen, drogueria, observaciones) " & _
                "values (" & CInt(dtcOs.BoundText) & ",'" & txtCod.Text & "','" & txtPer.Text & "','" & CDbl(txtImp.Text) & "','" & _
                CInt(txtRes.Text) & "'," & CInt(dtcDrog.BoundText) & ",'" & txtObs.Text & " ')"
    
    cn.Execute strSQL
Else
    vidNc = rsNC!idnc
    
    'ejecutar update para modificar los datos de la tabla
    strSQL = "update NCdroguerias set osocial = " & CInt(dtcOs.BoundText) & ", periodo = '" & txtPer.Text & "', importe = '" & CDbl(txtImp.Text) & _
    "', resumen = '" & CInt(txtRes.Text) & "', drogueria = " & CInt(dtcDrog.BoundText) & ", observaciones = '" & txtObs.Text & " ' where idnc = " & vidNc
    
    cn.Execute strSQL
End If

rsNC.Requery
dtgNC.Refresh
vAgrega = False
frameEdicion.Enabled = False
frameDatos.Enabled = True
FrameBotones.Enabled = True
dtgNC.SetFocus
Call TomaDatos
Exit Sub
Solucion:
   MsgBox Err.Number & "-" & Err.Description, vbInformation, "Error del Sistema..."
   Err.Clear
   On Error Resume Next
End Sub
Private Sub cmdSalir_Click()
rsOs.Close
rsPro.Close
rsNC.Close
Unload Me
End Sub
Private Sub dtcDrog_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtObs.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub
Private Sub dtcFiltro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdfil.SetFocus
End If
End Sub
Private Sub dtcOs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If txtCod.Enabled = False Then
        txtPer.SetFocus
    Else
        txtCod.SetFocus
    End If
    SendKeys "{home}+{end}"
End If
End Sub
Private Sub dtgNC_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If rsNC.EOF = False And rsNC.RecordCount > 0 Then
    Call TomaDatos
End If
End Sub
Private Sub Form_Load()

lblBus.Visible = False
txtBus.Visible = False
cmdBus.Visible = False
cmdNoBusca.Visible = False

FrameFiltro.Visible = False
frameEdicion.Enabled = False
FrameBotones.Enabled = True
frameDatos.Enabled = True
vAgrega = False

'tabla de proveedores para llenar el combo
rsPro.Open "select * from proveedores order by nombre", cn, adOpenDynamic, adLockReadOnly, adCmdText

Set dtcDrog.DataSource = rsPro
Set dtcDrog.RowSource = rsPro
dtcDrog.BoundColumn = "idproveedor"
dtcDrog.ListField = "nombre"
rsPro.MoveFirst
dtcDrog.BoundText = rsPro!idproveedor

'tabla con todo el registro de las notas de creditos obras sociales
'rsNC.Open "select o.nombre as ObSocial, n.codigo, n.periodo, n.importe, n.resumen, p.nombre as Proveedor, n.observaciones, n.Drogueria, n.idnc, n.osocial " & _
'            "from (ncdroguerias n " & _
'            "inner join obrasociales o on n.osocial = o.idos) " & _
'            "inner join Proveedores p on n.drogueria = p.idproveedor " & _
'            "order by n.periodo desc", cn, adOpenDynamic, adLockOptimistic, adCmdText

rsNC.Open "select o.nombre as ObSocial, n.codigo, n.periodo, n.importe, n.resumen, p.nombre as Proveedor, n.observaciones, n.Drogueria, n.idnc, n.osocial " & _
            "from ncdroguerias n, obrasociales o, proveedores p " & _
            "where n.osocial = o.idos and n.drogueria = p.idproveedor order by n.periodo desc", cn, adOpenDynamic, adLockOptimistic, adCmdText

'grilla de datos
rsNC.MoveFirst
Set dtgNC.DataSource = rsNC
dtgNC.Refresh

'Tabla de todas las obras sociales para llenar el combo
rsOs.Open "select * from obrasociales order by nombre", cn, adOpenDynamic, adLockReadOnly, adCmdText
Set dtcOs.DataSource = rsOs
Set dtcOs.RowSource = rsOs
dtcOs.BoundColumn = "idos"
dtcOs.ListField = "nombre"

'Combo de obras sociales para filtrar los datos de la grilla
Set dtcFiltro.DataSource = rsOs
Set dtcFiltro.RowSource = rsOs
dtcFiltro.BoundColumn = "idos"
dtcFiltro.ListField = "nombre"
Dim reg
rsOs.MoveFirst
reg = rsOs!idos
dtcFiltro.BoundText = reg

If rsNC.RecordCount > 0 Then
    Call TomaDatos
End If
End Sub
Private Sub txtBus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBus.SetFocus
End If
End Sub
Private Sub txtCod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtPer.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub
Private Sub txtImp_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        dtcDrog.SetFocus
    Case 44
        KeyAscii = 0
    Case 45
    Case 46
        KeyAscii = 44
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub txtImp_LostFocus()
If IsNumeric(txtImp.Text) = False And Len(txtImp.Text) > 0 Then
    MsgBox "NO SE ADMITEN CARACTERES !!", vbCritical, "ATENCION !"
    txtImp.SetFocus
End If
End Sub
Private Sub txtObs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdok.SetFocus
End If
End Sub
Private Sub txtPer_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtImp.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub
Private Sub txtRes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    dtcOs.SetFocus
End If
End Sub
Private Sub TomaDatos()
If rsNC.RecordCount = 0 Then
    txtRes.Text = ""
    txtCod.Text = ""
    txtImp.Text = ""
    txtObs.Text = ""
    txtPer.Text = ""
    Exit Sub
End If
txtRes.Text = rsNC!resumen
txtCod.Text = rsNC!codigo
txtPer.Text = rsNC!periodo
txtImp.Text = rsNC!importe
dtcDrog.BoundText = rsNC!drogueria
dtcOs.BoundText = rsNC!osocial
txtObs.Text = rsNC!observaciones
End Sub
