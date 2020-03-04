VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDepuCpras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Depuracion de Archivo de Compras ..."
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9495
   Icon            =   "frmDepuCpras.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   9495
   Begin VB.Frame frameOp 
      Caption         =   "Opciones de Depuracion de COMPRAS"
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9255
      Begin VB.OptionButton optTodo 
         Caption         =   "Borrar archivo completo ..."
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
         Left            =   4800
         TabIndex        =   13
         Top             =   420
         Width           =   3615
      End
      Begin VB.OptionButton optRango 
         Caption         =   "Borrar según rango de Fechas ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   480
         Value           =   -1  'True
         Width           =   3855
      End
      Begin VB.Frame frameFechas 
         Caption         =   "Rango de Fechas para filtrar Cuenta Corriente"
         Height          =   1095
         Left            =   600
         TabIndex        =   6
         Top             =   960
         Width           =   7215
         Begin MSComCtl2.DTPicker dtpIni 
            Height          =   375
            Left            =   1320
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
            Format          =   68747265
            CurrentDate     =   39392
         End
         Begin MSComCtl2.DTPicker dtpFin 
            Height          =   375
            Left            =   4200
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
            Format          =   68747265
            CurrentDate     =   39392
         End
         Begin AIFCmp1.asxPowerButton cmdVer 
            Height          =   405
            Left            =   6000
            TabIndex        =   9
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   714
            Picture         =   "frmDepuCpras.frx":0442
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
            Left            =   360
            TabIndex        =   11
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
            Left            =   3240
            TabIndex        =   10
            Top             =   480
            Width           =   690
         End
      End
      Begin AIFCmp1.asxPowerButton cmdSalir 
         Cancel          =   -1  'True
         Height          =   975
         Left            =   8040
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1720
         Picture         =   "frmDepuCpras.frx":075C
         Caption         =   "&Salir"
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
      End
   End
   Begin VB.Frame frameDatos 
      Caption         =   "Archivo de Pedidos a ser Depurado"
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   9255
      Begin AIFCmp1.asxPowerButton cmdCancelar 
         Height          =   975
         Left            =   7800
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1720
         Picture         =   "frmDepuCpras.frx":0CE8
         Caption         =   "&Cancelar"
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
      End
      Begin MSDataGridLib.DataGrid dtgArchivo 
         Height          =   5055
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   8916
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   255
         ColumnHeaders   =   -1  'True
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
         BeginProperty Column02 
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
         BeginProperty Column03 
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
            DataField       =   "depositado"
            Caption         =   "Deposito"
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
         BeginProperty Column07 
            DataField       =   "observaciones"
            Caption         =   "Observacionesf"
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
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   1214,929
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2505,26
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   390,047
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1830,047
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   929,764
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   870,236
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   599,811
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   6630,236
            EndProperty
         EndProperty
      End
      Begin AIFCmp1.asxPowerButton cmdBorrar 
         Height          =   975
         Left            =   7800
         TabIndex        =   3
         ToolTipText     =   "Elimina todos los registros visualizados en pantalla"
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1720
         Picture         =   "frmDepuCpras.frx":113A
         Caption         =   "Borrar Registros"
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
      End
   End
End
Attribute VB_Name = "frmDepuCpras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsFCpras As New ADODB.Recordset
Private Sub cmdBorrar_Click()
Err.Clear
On Error GoTo Solucion
SioNo = MsgBox("ESTA SEGURO DE ELIMINAR ESTA INFORMACION DEL SISTEMA ?", vbExclamation + vbYesNo, "ATENCION !")
If SioNo = vbNo Then
    Exit Sub
End If
'ejecuta comando sql
If optRango.Value = True Then
    cn.Execute "delete from FacturasCompras where fecha between #" & dtpIni.Value & "# and #" & dtpFin.Value & "#"
Else
    cn.Execute "delete * from FacturasCompras"
End If
rsFCpras.Requery
rsFCpras.Close
Set rsFCpras = Nothing
FrameDatos.Visible = False
frameFechas.Enabled = True

Exit Sub
Solucion:
MsgBox Err.Number & " - " & Err.Description + Chr(13) & _
        "NO SE PUDEDE SEGUIR CON EL PROCEDIMIENTO...", vbCritical, "ERROR DEL SISTEMA ...!"

End Sub
Private Sub cmdCancelar_Click()
FrameDatos.Visible = False
frameFechas.Enabled = True
rsFCpras.Close
Set rsFCpras = Nothing
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub cmdVer_Click()
If rsFCpras.State = 1 Then
    rsFCpras.Close
    Set rsFCpras = Nothing
End If
rsFCpras.Open "SELECT FacturasCompras.Fecha, Proveedores.Nombre, FacturasCompras.Numero, FacturasCompras.Tipo, FacturasCompras.Importe, FacturasCompras.Depositado, FacturasCompras.Observaciones, FacturasCompras.Estado " & _
              "FROM FacturasCompras INNER JOIN Proveedores ON FacturasCompras.idProveedor = Proveedores.idProveedor " & _
              "where fecha between #" & Format(dtpIni.Value, "mm/dd/yyyy") & "# and #" & Format(dtpFin.Value, "mm/dd/yyyy") & "# order by fecha", cn, adOpenDynamic, adLockOptimistic, adCmdText

If rsFCpras.RecordCount = 0 Then
    MsgBox "NO HAY DATOS EN EL PERIODO DE FECHAS INGRESADO...!", vbExclamation, "ATENCION !"
    dtpIni.SetFocus
    Exit Sub
End If
Set dtgArchivo.DataSource = rsFCpras
dtgArchivo.Refresh

FrameDatos.Visible = True
frameFechas.Enabled = False
End Sub
Private Sub Form_Load()
FrameDatos.Visible = False
Me.Top = 300
Me.Left = 200
If Month(Date) = 1 Then 'como es el primer mes del año, debe referirse al dic del año anterior
    dtpIni.Value = Format("01/12/" & (Year(Date)) - 1, "dd/mm/yyyy")
    dtpFin.Value = Format("31/12/" & (Year(Date)) - 1, "dd/mm/yyyy")
Else
    dtpIni.Value = Format("01/01" & "/" & Year(Date), "dd/mm/yyyy")
    If (Month(Date) - 1) = 1 Or (Month(Date) - 1) = 3 Or (Month(Date) - 1) = 5 Or (Month(Date) - 1) = 7 Or (Month(Date) - 1) = 8 Or (Month(Date) - 1) = 10 Or (Month(Date) - 1) = 12 Then
        dtpFin.Value = Format("31/" & (Month(Date) - 1) & "/" & Year(Date), "dd/mm/yyyy")
    End If
    If (Month(Date) - 1) = 4 Or (Month(Date) - 1) = 6 Or (Month(Date) - 1) = 9 Or (Month(Date) - 1) = 11 Then
        dtpFin.Value = Format("30/" & (Month(Date) - 1) & "/" & Year(Date), "dd/mm/yyyy")
    End If
    If (Month(Date) - 1) = 2 Then
        dtpFin.Value = Format("28/" & (Month(Date) - 1) & "/" & Year(Date), "dd/mm/yyyy")
    End If
End If
End Sub
Private Sub Form_Unload(cancel As Integer)
If rsFCpras.State = 1 Then
    rsFCpras.Close
    Set rsFCpras = Nothing
End If
End Sub
Private Sub optRango_Click()
optTodo.Value = False
frameFechas.Visible = True
frameFechas.Enabled = True
If FrameDatos.Visible = True Then
    FrameDatos.Visible = False
End If
End Sub
Private Sub optTodo_Click()
optRango.Value = False
frameFechas.Visible = False
If rsFCpras.State = 1 Then
    rsFCpras.Close
    Set rsFCpras = Nothing
End If

rsFCpras.Open "SELECT FacturasCompras.Fecha, Proveedores.Nombre, FacturasCompras.Numero, FacturasCompras.Tipo, FacturasCompras.Importe, FacturasCompras.Depositado, FacturasCompras.Observaciones, FacturasCompras.Estado " & _
                "FROM FacturasCompras INNER JOIN Proveedores ON FacturasCompras.idProveedor=Proveedores.idProveedor", cn, adOpenDynamic, adLockOptimistic, adCmdText

If rsFCpras.RecordCount = 0 Then
    MsgBox "NO HAY INFORMACION PARA MOSTRAR...!", vbExclamation, "ATENCION !"
    optTodo.SetFocus
    Exit Sub
End If
Set dtgArchivo.DataSource = rsFCpras
FrameDatos.Visible = True
dtgArchivo.Refresh
End Sub
