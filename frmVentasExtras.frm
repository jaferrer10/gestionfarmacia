VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVentasExtras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Ventas o Ingresos Extraordinarios ..."
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13920
   Icon            =   "frmVentasExtras.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   13920
   Begin VB.Frame frameFiltro 
      Caption         =   "Filtra datos del archivo"
      Height          =   2055
      Left            =   8640
      TabIndex        =   21
      Top             =   6240
      Width           =   5175
      Begin AIFCmp1.asxPowerButton cmdSinfiltro 
         Height          =   375
         Left            =   3600
         TabIndex        =   22
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Sacar Filtro"
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
      Begin AIFCmp1.asxPowerButton cmdFiltrar 
         Height          =   375
         Left            =   2280
         TabIndex        =   23
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Filtrar Datos"
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
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   375
         Left            =   2040
         TabIndex        =   24
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
         CalendarForeColor=   32768
         CalendarTitleForeColor=   32768
         Format          =   172883969
         CurrentDate     =   39451
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   375
         Left            =   2040
         TabIndex        =   25
         Top             =   840
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
         Format          =   174063617
         CurrentDate     =   39451
      End
      Begin VB.Label Label7 
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
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label8 
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
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   690
      End
   End
   Begin VB.Frame frameDatos 
      Caption         =   "Datos"
      Height          =   6015
      Left            =   8640
      TabIndex        =   8
      Top             =   120
      Width           =   5175
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
         Height          =   405
         Left            =   240
         MaxLength       =   150
         TabIndex        =   18
         Top             =   3360
         Width           =   4815
      End
      Begin VB.TextBox txtImporte 
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
         MaxLength       =   15
         TabIndex        =   17
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox txtDescripcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         MaxLength       =   100
         TabIndex        =   16
         Top             =   1680
         Width           =   4815
      End
      Begin MSDataListLib.DataCombo dtcConceptos 
         Bindings        =   "frmVentasExtras.frx":0442
         DataSource      =   "adoConceptos"
         Height          =   360
         Left            =   1560
         TabIndex        =   15
         Top             =   890
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         ForeColor       =   16711680
         Text            =   "DataCombo1"
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
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   375
         Left            =   1560
         TabIndex        =   14
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
         Format          =   174063617
         CurrentDate     =   39451
      End
      Begin AIFCmp1.asxPowerButton cmdGrabar 
         Height          =   495
         Left            =   1920
         TabIndex        =   19
         Top             =   4920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Picture         =   "frmVentasExtras.frx":045D
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
      End
      Begin AIFCmp1.asxPowerButton cmdCancelar 
         Height          =   495
         Left            =   3600
         TabIndex        =   20
         Top             =   4920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Picture         =   "frmVentasExtras.frx":08AF
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
      End
      Begin AIFCmp1.asxPowerButton cmdAgrCon 
         Height          =   495
         Left            =   4320
         TabIndex        =   28
         ToolTipText     =   "Agrega Conceptos Nuevos"
         Top             =   840
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
         BorderStyle     =   4
         Picture         =   "frmVentasExtras.frx":0D01
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
      Begin VB.Label Label5 
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
         TabIndex        =   13
         Top             =   3120
         Width           =   1650
      End
      Begin VB.Label Label4 
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
         TabIndex        =   12
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion:"
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
         TabIndex        =   11
         Top             =   1440
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Concepto:"
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
         TabIndex        =   10
         Top             =   960
         Width           =   1065
      End
      Begin VB.Label Label1 
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
         TabIndex        =   9
         Top             =   480
         Width           =   720
      End
   End
   Begin VB.Frame FrameArchivo 
      Caption         =   "Archivo"
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin AIFCmp1.asxLabel lblTotal 
         Height          =   390
         Left            =   6840
         TabIndex        =   3
         Top             =   6600
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Total"
         BorderStyle     =   1
         WordWrap        =   -1  'True
         Alignment       =   1
         UseMnemonic     =   -1  'True
         MouseIcon       =   "frmVentasExtras.frx":0E5B
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx1 
         Height          =   240
         Left            =   4200
         TabIndex        =   2
         Top             =   6600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   423
         Caption         =   "Total Importe"
      End
      Begin MSDataGridLib.DataGrid dtgVentas 
         Height          =   6255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   11033
         _Version        =   393216
         AllowUpdate     =   0   'False
         ForeColor       =   16711680
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "idextra"
            Caption         =   "Cod"
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
            DataField       =   "cd"
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
            DataField       =   "vd"
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
         BeginProperty Column06 
            DataField       =   "procesado"
            Caption         =   "Procesado"
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
               ColumnWidth     =   599,811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1230,236
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1830,047
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2849,953
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   5114,835
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   870,236
            EndProperty
         EndProperty
      End
      Begin AIFCmp1.asxPowerButton cmdAgregar 
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   7320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Picture         =   "frmVentasExtras.frx":1175
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
      Begin AIFCmp1.asxPowerButton cmdModificar 
         Height          =   495
         Left            =   1680
         TabIndex        =   5
         Top             =   7320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Picture         =   "frmVentasExtras.frx":170F
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
      End
      Begin AIFCmp1.asxPowerButton cmdEliminar 
         Height          =   495
         Left            =   3120
         TabIndex        =   6
         Top             =   7320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Picture         =   "frmVentasExtras.frx":1B61
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
      End
      Begin AIFCmp1.asxPowerButton cmdSalir 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   6960
         TabIndex        =   7
         Top             =   7320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Picture         =   "frmVentasExtras.frx":1FB3
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
         PictureOffsetX  =   10
      End
   End
End
Attribute VB_Name = "frmVentasExtras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsExtras As New ADODB.Recordset
Private rsConceptos As New ADODB.Recordset
Private VidExtra As Integer
Private Sub cmdAgrCon_Click()
frmAgregaConcepExt.Show vbModal
rsConceptos.Requery
dtcConceptos.Refresh
End Sub
Private Sub cmdAgregar_Click()
frameDatos.Visible = True
FrameArchivo.Enabled = False
dtpFecha.Value = Date
vAgrega = True
End Sub
Private Sub cmdCancelar_Click()
txtDescripcion.Text = ""
txtImporte.Text = ""
frameDatos.Visible = False
FrameArchivo.Enabled = True
End Sub
Private Sub cmdEliminar_Click()
Err.Clear
On Error GoTo SolucionEliminar
SioNo = MsgBox("ESTA SEGURO DE ELIMINAR ESTE REGISTRO ?", vbExclamation + vbYesNo, "ATENCION !")
If SioNo = vbYes Then
    VidExtra = rsExtras!idextra
    cn.Execute "delete from ventasextraordinarias where idextra = " & VidExtra
    rsExtras.Requery
    rsExtras.Update
    dtgVentas.Refresh
    Call CalculaTotal
End If
Exit Sub

SolucionEliminar:
    If Err.Number = 3021 Then
        MsgBox "POR FAVOR, SELECCIONE UN REGISTRO EN LA GRILLA PARA ESTA FUNCION...", vbCritical, "ATENCION !"
    Else
       MsgBox Err.Number & "-" & Err.Description, vbInformation, "Error del Sistema..."
    End If
   
End Sub
Private Sub cmdFiltrar_Click()
rsExtras.Close
Set rsExtras = Nothing
strSQL = "SELECT VentasExtraordinarias.idExtra, VentasExtraordinarias.Fecha, VentasExtraordinarias.concepto, ConceptosExtraordinarios.Descripcion as cd, VentasExtraordinarias.Descripcion as vd, VentasExtraordinarias.Importe, VentasExtraordinarias.Observaciones " & _
         "FROM VentasExtraordinarias INNER JOIN ConceptosExtraordinarios ON VentasExtraordinarias.Concepto = ConceptosExtraordinarios.idConcepto " & _
         "where VentasExtraordinarias.fecha between #" & Format(dtpDesde.Value, "mm/dd/yyyy") & "# and #" & Format(dtpHasta.Value, "mm/dd/yyyy") & "#"
rsExtras.Open strSQL, cn, adOpenDynamic, adLockOptimistic, adCmdText
Set dtgVentas.DataSource = rsExtras
dtgVentas.Refresh
Call CalculaTotal
End Sub
Private Sub cmdGrabar_Click()
If Len(dtcConceptos.Text) = 0 Then
    MsgBox "DEBE SELECCIONAR ALGUN CONCEPTO PARA ESTE INGRESO....", vbInformation, "ATENCION !"
    dtcConceptos.SetFocus
    Exit Sub
End If
If Len(txtDescripcion.Text) = 0 Then
    MsgBox "DEBE INGRESAR ALGUNA DESCRIPCION PARA ESTE REGISTRO...", vbInformation, "ATENCION !"
    txtDescripcion.SetFocus
    Exit Sub
End If
If Len(txtImporte.Text) = 0 Then
    MsgBox "DEBE INGRESAR ALGUN IMPORTE PARA ESTE REGISTRO ...", vbInformation, "ATENCION !"
    txtImporte.SetFocus
    Exit Sub
End If
SioNo = MsgBox("ESTA SEGURO GRABAR TODOS LOS DATOS INGRESADOS ?", vbExclamation + vbYesNo, "ATENCION !")
If SioNo = vbYes Then
    If vAgrega = True Then
        rsExtras.AddNew
    End If
    rsExtras!fecha = dtpFecha.Value
    rsExtras!concepto = dtcConceptos.BoundText
    rsExtras!vd = txtDescripcion.Text
    rsExtras!importe = txtImporte.Text
    rsExtras!observaciones = txtObser.Text
    rsExtras.Update
    rsExtras.Requery
    'blanqueo campos
    txtDescripcion.Text = ""
    txtObser.Text = ""
    txtImporte.Text = 0
    
    frameDatos.Visible = False
    FrameArchivo.Enabled = True
    Call CalculaTotal
    dtgVentas.SetFocus
End If
vAgrega = True
End Sub
Private Sub cmdModificar_Click()
Err.Clear
On Error GoTo SolucionModi
If rsExtras.RecordCount = 0 Then
    MsgBox "NO HAY INFORMACION PARA MODIFICAR ...!", vbInformation, "ATENCION !"
    Exit Sub
End If
frameDatos.Visible = True
FrameArchivo.Enabled = False
vAgrega = False
txtDescripcion.Text = rsExtras!vd
dtpFecha.Value = Format(rsExtras!fecha, "dd/mm/yyyy")
txtImporte.Text = rsExtras!importe
dtcConceptos.BoundText = rsExtras!concepto
txtObser.Text = rsExtras!observaciones
txtImporte.SetFocus
txtDescripcion.SetFocus
SendKeys "{home}+{end}"
Exit Sub

SolucionModi:
    If Err.Number = 3021 Then
        MsgBox "POR FAVOR, SELECCIONE UN REGISTRO EN LA GRILLA PARA ESTA FUNCION...", vbCritical, "ATENCION !"
    Else
       MsgBox Err.Number & "-" & Err.Description, vbInformation, "Error del Sistema..."
    End If
    frameDatos.Enabled = False
    FrameArchivo.Enabled = True

End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub cmdSinfiltro_Click()
rsExtras.Close
Set rsExtras = Nothing
strSQL = "SELECT VentasExtraordinarias.idExtra, VentasExtraordinarias.Fecha, VentasExtraordinarias.concepto, ConceptosExtraordinarios.Descripcion as cd, VentasExtraordinarias.Descripcion as vd, VentasExtraordinarias.Importe, VentasExtraordinarias.Observaciones " & _
         "FROM VentasExtraordinarias INNER JOIN ConceptosExtraordinarios ON VentasExtraordinarias.Concepto = ConceptosExtraordinarios.idConcepto"

rsExtras.Open strSQL, cn, adOpenDynamic, adLockOptimistic, adCmdText
Set dtgVentas.DataSource = rsExtras
dtgVentas.Refresh
Call CalculaTotal
End Sub
Private Sub dtcConceptos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtDescripcion.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub Form_Load()
Me.Top = 300
Me.Left = 0
strSQL = "SELECT VentasExtraordinarias.idExtra, VentasExtraordinarias.Fecha, VentasExtraordinarias.concepto, ConceptosExtraordinarios.Descripcion as cd, VentasExtraordinarias.Descripcion as vd, VentasExtraordinarias.Importe, VentasExtraordinarias.Observaciones " & _
         "FROM VentasExtraordinarias INNER JOIN ConceptosExtraordinarios ON VentasExtraordinarias.Concepto = ConceptosExtraordinarios.idConcepto order by VentasExtraordinarias.Fecha desc"

rsExtras.Open strSQL, cn, adOpenDynamic, adLockOptimistic, adCmdText
Set dtgVentas.DataSource = rsExtras
dtgVentas.Refresh

frameDatos.Visible = False
'llena el combo conceptos
rsConceptos.Open "select * from conceptosextraordinarios order by Descripcion", cn, adOpenDynamic, adLockReadOnly, adCmdText
Set dtcConceptos.RowSource = rsConceptos
dtcConceptos.BoundColumn = "idconcepto"
dtcConceptos.BoundText = rsConceptos!descripcion
dtcConceptos.ListField = "Descripcion"

FrameArchivo.Enabled = True

Call CalculaTotal

'pone el 1º y utlimo dia del mes actual en los campos fechas para filtros
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
End Sub
Private Sub Form_Unload(cancel As Integer)
rsExtras.Close
Set rsExtras = Nothing
rsConceptos.Close
Set rsConceptos = Nothing
End Sub
Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtImporte.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub
Private Sub txtImporte_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        txtObser.SetFocus
        SendKeys "{home}+{end}"
    Case 44 'para que no acepte la entrada de la coma
        KeyAscii = 0
    Case 45
    Case 46
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub CalculaTotal()
If rsExtras.RecordCount = 0 Then Exit Sub
LblTotal.Caption = 0
rsExtras.MoveFirst
Do While rsExtras.EOF = False
    LblTotal.Caption = LblTotal.Caption + rsExtras!importe
    rsExtras.MoveNext
Loop
rsExtras.MoveFirst
End Sub
Private Sub txtObser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdGrabar.SetFocus
End If
End Sub
