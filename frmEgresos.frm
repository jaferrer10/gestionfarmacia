VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEgresos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Egresos de Dinero ..."
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11985
   Icon            =   "frmEgresos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   11985
   Begin VB.Frame Frame2 
      Caption         =   "Archivo de Egresos"
      Height          =   5175
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   11775
      Begin VB.Frame frameFiltro 
         Caption         =   "Filtro de datos"
         Height          =   2415
         Left            =   9480
         TabIndex        =   17
         Top             =   2640
         Width           =   2175
         Begin AIFCmp1.asxPowerButton cmdSi 
            Height          =   495
            Left            =   360
            TabIndex        =   22
            Top             =   1800
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   873
            Picture         =   "frmEgresos.frx":030A
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
         Begin MSComCtl2.DTPicker dtpDesde 
            Height          =   375
            Left            =   240
            TabIndex        =   20
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
            Format          =   128843777
            CurrentDate     =   39468
         End
         Begin MSComCtl2.DTPicker dtpHasta 
            Height          =   375
            Left            =   240
            TabIndex        =   21
            Top             =   1200
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
            Format          =   128843777
            CurrentDate     =   39468
         End
         Begin AIFCmp1.asxPowerButton cmdNo 
            Height          =   495
            Left            =   1080
            TabIndex        =   23
            Top             =   1800
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   873
            Picture         =   "frmEgresos.frx":075C
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
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
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
            TabIndex        =   19
            Top             =   960
            Width           =   630
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
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
            TabIndex        =   18
            Top             =   240
            Width           =   705
         End
      End
      Begin MSDataGridLib.DataGrid dtgArchivo 
         Height          =   4695
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   8281
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
            Weight          =   400
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
            DataField       =   "DE"
            Caption         =   "Descripcion Egreso"
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
            DataField       =   "DC"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
               ColumnWidth     =   1110,047
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3404,977
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2954,835
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1454,74
            EndProperty
         EndProperty
      End
      Begin AIFCmp1.asxPowerButton cmdModificar 
         Height          =   495
         Left            =   9840
         TabIndex        =   8
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Picture         =   "frmEgresos.frx":08B6
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
      End
      Begin AIFCmp1.asxPowerButton cmdEliminar 
         Height          =   495
         Left            =   9840
         TabIndex        =   9
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Picture         =   "frmEgresos.frx":0D08
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
      End
      Begin AIFCmp1.asxPowerButton cmdFiltro 
         Height          =   495
         Left            =   9840
         TabIndex        =   10
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Picture         =   "frmEgresos.frx":115A
         Caption         =   "Filtrar"
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
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Egreso"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      Begin AIFCmp1.asxPowerBanner lblFiltro 
         Height          =   375
         Left            =   9720
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         FormatString    =   "Filtro Activo"
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
         TextColor       =   16744576
      End
      Begin AIFCmp1.asxPowerButton asxPowerButton1 
         Height          =   495
         Left            =   9000
         TabIndex        =   25
         ToolTipText     =   "Filtra los datos por el Concepto seleccionado"
         Top             =   200
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
         BorderStyle     =   4
         Picture         =   "frmEgresos.frx":15AC
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
      Begin AIFCmp1.asxPowerButton cmdAgrCon 
         Height          =   495
         Left            =   8160
         TabIndex        =   24
         ToolTipText     =   "Agrega Conceptos Nuevos"
         Top             =   180
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
         BorderStyle     =   4
         Picture         =   "frmEgresos.frx":1B46
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1560
         TabIndex        =   16
         Top             =   1440
         Width           =   1455
      End
      Begin AIFCmp1.asxPowerButton cmdok 
         Height          =   495
         Left            =   1560
         TabIndex        =   4
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Picture         =   "frmEgresos.frx":20E0
         Caption         =   "&Grabar"
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
      Begin MSDataListLib.DataCombo dtcConcepto 
         Height          =   360
         Left            =   4080
         TabIndex        =   2
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin VB.TextBox txtDescripcion 
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   3
         Top             =   840
         Width           =   6615
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   375
         Left            =   960
         TabIndex        =   1
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
         Format          =   128974849
         CurrentDate     =   39468
      End
      Begin AIFCmp1.asxPowerButton cmdSalir 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   9840
         TabIndex        =   6
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Picture         =   "frmEgresos.frx":267A
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
      Begin AIFCmp1.asxPowerButton cmdCancelar 
         Height          =   495
         Left            =   3240
         TabIndex        =   5
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Picture         =   "frmEgresos.frx":2C06
         Caption         =   "&Cancelar"
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
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
         TabIndex        =   15
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
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
         Left            =   2880
         TabIndex        =   14
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
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
         Top             =   960
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Top             =   360
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmEgresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsEgresos As New ADODB.Recordset
Private rsConceptos As New ADODB.Recordset

Private Sub asxPowerButton1_Click()
Err.Clear
On Error GoTo ErrorEgresos

If Len(dtcConcepto.Text) = 0 Then
    MsgBox "DEBE SELECCIONAR UN CONCEPTO PARA USAR EL FILTRO...!", vbCritical, "ATENCION !"
    dtcConcepto.SetFocus
    Exit Sub
End If
    
If lblFiltro.Visible = False Then

    lblFiltro.Visible = True
    rsEgresos.Close
    Set rsEgresos = Nothing
    rsEgresos.Open "select egresos.idegreso,egresos.Fecha,egresos.Descripcion as DE,egresos.concepto,conceptosegresos.Descripcion as DC,egresos.Importe from Egresos inner join ConceptosEgresos on egresos.concepto=conceptosegresos.idconeg " & _
                    "where egresos.concepto = " & dtcConcepto.BoundText & " order by fecha desc", cn, adOpenDynamic, adLockOptimistic, adCmdText
    
    Set dtgArchivo.DataSource = rsEgresos
    dtgArchivo.Refresh
    
Else
    lblFiltro.Visible = False
    rsEgresos.Close
    Set rsEgresos = Nothing
    rsEgresos.Open "select egresos.idegreso,egresos.Fecha,egresos.Descripcion as DE,egresos.concepto,conceptosegresos.Descripcion as DC,egresos.Importe from Egresos inner join ConceptosEgresos on egresos.concepto=conceptosegresos.idconeg " & _
                   "order by fecha desc, concepto", cn, adOpenDynamic, adLockOptimistic, adCmdText
    Set dtgArchivo.DataSource = rsEgresos
    dtgArchivo.Refresh
    
End If
Exit Sub
ErrorEgresos:
     MsgBox Err.Number & " " & Err.Description, vbInformation, "ATENCION - ERROR ..."
End Sub

Private Sub cmdAgrCon_Click()
frmAgregaConcep.Show vbModal
rsConceptos.Requery
dtcConcepto.Refresh
End Sub
Private Sub cmdCancelar_Click()
vAgrega = True
txtDescripcion.Text = ""
txtImporte.Text = ""
dtcConcepto.SetFocus
End Sub
Private Sub cmdEliminar_Click()
SioNo = MsgBox("ESTA SEGURO DE ELIMIAR EL REGISTRO SELECCIONADO ?", vbExclamation + vbYesNo, "ATENCION !")
If SioNo = vbYes Then
    videg = rsEgresos!idegreso
    If rsEgresos.RecordCount = 0 Then
        MsgBox "NO HAY REGISTROS EN EL ARCHIVO PARA ELMINAR ...!", vbCritical, "ATENCION !"
        Exit Sub
    End If
    cn.Execute "delete from egresos where idegreso = " & videg
    rsEgresos.Requery
    dtgArchivo.Refresh
End If
End Sub
Private Sub cmdFiltro_Click()
FrameFiltro.Visible = True
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
Private Sub cmdModificar_Click()
vAgrega = False
dtpFecha.Value = rsEgresos!fecha
dtcConcepto.BoundText = rsEgresos!concepto
txtDescripcion.Text = rsEgresos!DE
txtImporte.Text = rsEgresos!Importe
txtDescripcion.SetFocus
End Sub
Private Sub cmdNo_Click()
FrameFiltro.Visible = False
rsEgresos.Close
Set rsEgresos = Nothing
rsEgresos.Open "select egresos.idegreso,egresos.Fecha,egresos.Descripcion as DE,egresos.concepto,conceptosegresos.Descripcion as DC,egresos.Importe from Egresos inner join ConceptosEgresos on egresos.concepto=conceptosegresos.idconeg " & _
                " order by fecha desc, concepto", cn, adOpenDynamic, adLockOptimistic, adCmdText
Set dtgArchivo.DataSource = rsEgresos
dtgArchivo.Refresh
End Sub
Private Sub cmdOk_Click()
If Len(dtcConcepto.Text) = 0 Or Len(txtDescripcion.Text) = 0 Or Len(txtImporte.Text) = 0 Then
    MsgBox "HAY CAMPOS VACIOS, DEBE INTRODUCIR DATOS PARA GRABAR ...!", vbExclamation, "ATENCION !"
    txtDescripcion.SetFocus
    Exit Sub
End If
If vAgrega = True Then
    rsEgresos.AddNew
End If
rsEgresos!fecha = dtpFecha.Value
rsEgresos!DE = txtDescripcion.Text
rsEgresos!concepto = dtcConcepto.BoundText
rsEgresos!Importe = txtImporte.Text
rsEgresos.Update
rsEgresos.Requery
dtgArchivo.Refresh
txtDescripcion.Text = ""
txtImporte.Text = ""
txtDescripcion.SetFocus
vAgrega = True
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdSi_Click()
rsEgresos.Close
Set rsEgresos = Nothing
rsEgresos.Open "select egresos.idegreso,egresos.Fecha,egresos.Descripcion as DE,egresos.concepto,conceptosegresos.Descripcion as DC,egresos.Importe from Egresos inner join ConceptosEgresos on egresos.concepto=conceptosegresos.idconeg " & _
                "where fecha between #" & Format(dtpDesde.Value, "mm/dd/yyyy") & "# and #" & Format(dtpHasta.Value, "mm/dd/yyyy") & "# order by fecha", cn, adOpenDynamic, adLockOptimistic, adCmdText

Set dtgArchivo.DataSource = rsEgresos
dtgArchivo.Refresh

End Sub

Private Sub dtcConcepto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtDescripcion.SetFocus
End If
End Sub
Private Sub dtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    dtcConcepto.SetFocus
End If
End Sub
Private Sub Form_Load()
Me.Top = 200
Me.Left = 0
FrameFiltro.Visible = False
lblFiltro.Visible = False

dtpFecha.Value = Date

rsEgresos.Open "select egresos.idegreso,egresos.Fecha,egresos.Descripcion as DE,egresos.concepto,conceptosegresos.Descripcion as DC,egresos.Importe from Egresos inner join ConceptosEgresos on egresos.concepto=conceptosegresos.idconeg " & _
                "order by fecha desc, concepto", cn, adOpenDynamic, adLockOptimistic, adCmdText

Set dtgArchivo.DataSource = rsEgresos
dtgArchivo.Refresh

'llena el combo de conceptos
rsConceptos.Open "select * from conceptosegresos order by descripcion", cn, adOpenDynamic, adLockReadOnly, adCmdText
Set dtcConcepto.DataSource = rsConceptos
Set dtcConcepto.RowSource = rsConceptos
dtcConcepto.ListField = "Descripcion"
dtcConcepto.BoundColumn = "idconeg"

vAgrega = True
End Sub
Private Sub Form_Unload(cancel As Integer)
If rsEgresos.State = True Then
    rsEgresos.Close
End If
Set rsEgresos = Nothing
rsConceptos.Close
Set rsConceptos = Nothing
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        cmdOk.SetFocus
    Case 44 'para que no acepte la entrada de la coma
        KeyAscii = 0
    Case 45
    Case 46
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtImporte.SetFocus
End If
End Sub
