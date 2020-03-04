VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmConsCajas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Archivo de Registros de Control de Cajas ..."
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9405
   Icon            =   "frmConsCajas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8715
   ScaleWidth      =   9405
   Begin VB.Frame frameFechas 
      Caption         =   "Rango de Fechas para filtrar Cuenta Corriente"
      Height          =   1455
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   7215
      Begin AIFCmp1.asxPowerButton cmdVer 
         Height          =   405
         Left            =   3240
         TabIndex        =   9
         Top             =   960
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   714
         Picture         =   "frmConsCajas.frx":058A
         Caption         =   "Ver"
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
         PictureAlignment=   0
         PictureOffsetX  =   5
         PictureOffsetY  =   5
      End
      Begin MSComCtl2.DTPicker dtpIni 
         Height          =   375
         Left            =   1800
         TabIndex        =   5
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
         CurrentDate     =   39392
      End
      Begin MSComCtl2.DTPicker dtpFin 
         Height          =   375
         Left            =   4680
         TabIndex        =   6
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
         CurrentDate     =   39392
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
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
         Left            =   840
         TabIndex        =   8
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label2 
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
         Height          =   240
         Left            =   3720
         TabIndex        =   7
         Top             =   480
         Width           =   690
      End
   End
   Begin VB.OptionButton optFechas 
      Caption         =   "Mostrar según Rango de Fechas"
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
      Left            =   4200
      TabIndex        =   3
      Top             =   120
      Width           =   3495
   End
   Begin VB.OptionButton OptTodos 
      Caption         =   "Mostrar Todos los Registros"
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
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Registros"
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   9135
      Begin AIFCmp1.asxPowerButton cmdImprimir 
         Height          =   615
         Left            =   7440
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         Picture         =   "frmConsCajas.frx":08A4
         Caption         =   "&Imprimir"
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
      Begin AIFCmp1.asxPowerButton cmdEliminar 
         Height          =   615
         Left            =   7440
         TabIndex        =   10
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         Picture         =   "frmConsCajas.frx":09FE
         Caption         =   "&Eliminar"
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
      End
      Begin MSDataGridLib.DataGrid dtgRegistros 
         Height          =   6015
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   10610
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
         ColumnCount     =   8
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
            DataField       =   "Inicio"
            Caption         =   "Inicio"
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
            DataField       =   "extracciones"
            Caption         =   "Extracciones"
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
            DataField       =   "credito"
            Caption         =   "Credito"
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
            DataField       =   "caja"
            Caption         =   "Caja"
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
            DataField       =   "resultado"
            Caption         =   "Resultado"
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
            DataField       =   "turno"
            Caption         =   "Turno"
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
               ColumnWidth     =   1349,858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1005,165
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1035,213
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1049,953
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1124,787
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1124,787
            EndProperty
            BeginProperty Column07 
            EndProperty
         EndProperty
      End
   End
   Begin AIFCmp1.asxPowerButton cmdSalir 
      Cancel          =   -1  'True
      Height          =   1095
      Left            =   7680
      TabIndex        =   12
      ToolTipText     =   "Sale del modulo registro de ventas"
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1931
      FocusStyle      =   1
      BorderStyle     =   4
      Picture         =   "frmConsCajas.frx":0E50
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
End
Attribute VB_Name = "frmConsCajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsRegCajas As New ADODB.Recordset
Private rpt_Cajas As New CrptControlCajas

Private Sub cmdEliminar_Click()
If rsRegCajas.RecordCount = 0 Then
    MsgBox "NO REGISTROS PARA ELIMINAR !", vbCritical, "ATENCION !"
    Exit Sub
End If
frmPideClave.Show vbModal
If TempNivel = 1 Then
    SioNo = MsgBox("ESTA SEGURO DE ELEMINAR ESTE REGISTRO DE CAJA ?", vbExclamation + vbYesNo, "ATENCION !")
    If SioNo = vbYes Then
        rsRegCajas.Delete
        rsRegCajas.Update
        dtgRegistros.Refresh
    End If
Else
    MsgBox "SU NIVEL DE AUTORIZACON NO LE PERMITE ELIMINAR REGISTROS...", vbExclamation, "Seguridad del Sistema..."
End If
End Sub

Private Sub cmdImprimir_Click()

rpt_Cajas.Database.SetDataSource rsRegCajas

Set rptGeneral = rpt_Cajas ' Asigna el reporte al objeto reporte general utilizado
                           ' en el Form de la Vista Previa.
frmVistaPrevia.Show vbModal

Set rpt_Cajas = Nothing

End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub cmdVer_Click()
rsRegCajas.Close
Set rsRegCajas = Nothing
rsRegCajas.Open "select * from controlcajas where fecha between #" & _
                Format(dtpIni.Value, "mm/dd/yyyy") & "# and #" & Format(dtpFin.Value, "mm/dd/yyyy") & "# order by fecha desc", cn, adOpenDynamic, adLockOptimistic, adCmdText
                
Set dtgRegistros.DataSource = rsRegCajas
dtgRegistros.Refresh

End Sub
Private Sub Form_Load()
Me.Top = 150
Me.Left = 0
rsRegCajas.Open "select * from controlcajas order by fecha desc", cn, adOpenDynamic, adLockOptimistic, adCmdText
Set dtgRegistros.DataSource = rsRegCajas
dtgRegistros.Refresh
dtpIni.Value = Date
dtpFin.Value = Date
frameFechas.Visible = False
End Sub
Private Sub Form_Unload(cancel As Integer)
If rsRegCajas.State = 1 Then
    rsRegCajas.Clone
    Set rsRegCajas = Nothing
End If
End Sub
Private Sub optFechas_Click()
optFechas.Value = True
frameFechas.Visible = True
End Sub
Private Sub OptTodos_Click()
OptTodos.Value = True
frameFechas.Visible = False
rsRegCajas.Close
Set rsRegCajas = Nothing
rsRegCajas.Open "select * from controlcajas order by fecha desc", cn, adOpenDynamic, adLockOptimistic, adCmdText
Set dtgRegistros.DataSource = rsRegCajas
dtgRegistros.Refresh

End Sub
