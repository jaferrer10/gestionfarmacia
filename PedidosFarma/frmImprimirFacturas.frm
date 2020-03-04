VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmImprimirFacturas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresion de Archivo de Facturas de Proveedores ..."
   ClientHeight    =   5055
   ClientLeft      =   4350
   ClientTop       =   3360
   ClientWidth     =   7455
   Icon            =   "frmImprimirFacturas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7455
   Begin VB.Frame frameOpc 
      Caption         =   "Opciones de Periodo"
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7215
      Begin VB.OptionButton optTodo 
         Caption         =   "Todos los datos del Archivo"
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton optFechas 
         Caption         =   "Establecer Rango de Fecha"
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame frameFechas 
      Caption         =   "Rango de Fechas para filtrar Cuenta Corriente"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   7215
      Begin VB.OptionButton Option2 
         Caption         =   "Usar Filtro por Tipo de Factura:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   2040
         Width           =   3735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "No usar filtro de Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   1440
         Value           =   -1  'True
         Width           =   3495
      End
      Begin MSDataListLib.DataCombo cbtipo 
         Height          =   360
         Left            =   4680
         TabIndex        =   10
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
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
      Begin MSComCtl2.DTPicker dtpIni 
         Height          =   375
         Left            =   1800
         TabIndex        =   1
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
         Format          =   153157633
         CurrentDate     =   39392
      End
      Begin MSComCtl2.DTPicker dtpFin 
         Height          =   375
         Left            =   4680
         TabIndex        =   2
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
         Format          =   153157633
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
         Index           =   0
         Left            =   840
         TabIndex        =   4
         Top             =   720
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
         TabIndex        =   3
         Top             =   720
         Width           =   690
      End
   End
   Begin AIFCmp1.asxPowerButton cmdImprimir 
      Height          =   735
      Left            =   3480
      TabIndex        =   5
      Top             =   4080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      Picture         =   "frmImprimirFacturas.frx":27A2
      Caption         =   "&Imprimir"
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
   Begin AIFCmp1.asxPowerButton cmdCancelar 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   5520
      TabIndex        =   9
      Top             =   4080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      Picture         =   "frmImprimirFacturas.frx":307C
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
      PictureOffsetX  =   5
   End
End
Attribute VB_Name = "frmImprimirFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsFact As New ADODB.Recordset
Private rsTipFac As New ADODB.Recordset
Private rptFacturas As New CrptFacturasCompras
Private Sub cbtipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdImprimir.SetFocus
End If
End Sub
Private Sub cmdCancelar_Click()
Unload Me
End Sub
Private Sub cmdImprimir_Click()
Err.Clear
On erro GoTo SolucionErr
If rsFact.State = 1 Then
    rsFact.Close
    Set rsFact = Nothing
End If
If optTodo.Value = True Then
    rsFact.Open "select f.idproveedor, p.nombre, p.telefono, fechavto, numero, tipo, importe, depositado, f.observaciones, estado " & _
                "from facturascompras f inner join Proveedores p " & _
                "on f.idproveedor = p.idproveedor " & _
                "where f.idproveedor = " & frmCompras.dtcProveedor.BoundText & _
                " order by fechavto desc, numero", cn, adOpenDynamic, adLockReadOnly, adCmdText
Else
    If Option2.Value = True Then
        rsFact.Open "select f.idproveedor, p.nombre, p.telefono, fechavto, numero, tipo, importe, depositado, f.observaciones, estado " & _
                "from facturascompras f inner join Proveedores p " & _
                "on f.idproveedor = p.idproveedor " & _
                "WHERE f.idproveedor = " & frmCompras.dtcProveedor.BoundText & _
                " and Fechavto between #" & Format(dtpIni.Value, "mm/dd/yyyy") & _
                "# and #" & Format(dtpFin.Value, "mm/dd/yyyy") & "# and tipo = '" & cbTipo.Text & "' order by fechavto desc, numero", cn, adOpenDynamic, adLockReadOnly, adCmdText
    Else
        rsFact.Open "select f.idproveedor, p.nombre, p.telefono, fechavto, numero, tipo, importe, depositado, f.observaciones, estado " & _
                "from facturascompras f inner join Proveedores p " & _
                "on f.idproveedor = p.idproveedor " & _
                "WHERE f.idproveedor = " & frmCompras.dtcProveedor.BoundText & _
                " and Fechavto between #" & Format(dtpIni.Value, "mm/dd/yyyy") & _
                "# and #" & Format(dtpFin.Value, "mm/dd/yyyy") & "# order by fechavto desc, numero", cn, adOpenDynamic, adLockReadOnly, adCmdText
    End If
End If
If rsFact.RecordCount = 0 Then
    MsgBox "NO HAY IMFORMACION PARA LA IMPRESION...", vbInformation, "ATENCION !"
    rsFact.Close
    Set rsFact = Nothing
    Exit Sub
End If

rptFacturas.Database.SetDataSource rsFact
rptFacturas.txtDesde.SetText dtpIni.Value
rptFacturas.txtHasta.SetText dtpFin.Value

Set rptGeneral = rptFacturas ' Asigna el reporte al objeto reporte general utilizado
                          ' en el Form de la Vista Previa.

frmVistaPrevia.Show vbModal

Set rptFacturas = Nothing

rsFact.Close
Set rsFact = Nothing

Exit Sub

SolucionErr:
     MsgBox Err.Number & " " & Err.Description, vbInformation, "ATENCION - ERROR ..."

End Sub
Private Sub dtpFin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Option1.SetFocus
End If
End Sub
Private Sub dtpIni_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    dtpFin.SetFocus
End If
End Sub
Private Sub Form_Load()
Me.Top = 2000
Me.Left = 4000

If Month(Date) = 1 Then
    dtpIni.Value = Format("01/12" & "/" & Str(Year(Date) - 1), "dd/mm/yyyy")
Else
    dtpIni.Value = Format("01/" & Str((Month(Date)) - 1) & "/" & Str(Year(Date)), "dd/mm/yyyy")
End If
If (Month(Date)) = 1 Then
    dtpFin.Value = Format("31/12" & "/" & Year(Date), "dd/mm/yyyy")
ElseIf (Month(Date) - 1) = 3 Or (Month(Date) - 1) = 5 Or (Month(Date) - 1) = 7 Or (Month(Date) - 1) = 8 Or (Month(Date) - 1) = 10 Or (Month(Date) - 1) = 12 Then
    dtpFin.Value = Format("31/" & (Month(Date) - 1) & "/" & Year(Date), "dd/mm/yyyy")
ElseIf (Month(Date) - 1) = 4 Or (Month(Date) - 1) = 6 Or (Month(Date) - 1) = 9 Or (Month(Date) - 1) = 11 Then
    dtpFin.Value = Format("30/" & (Month(Date) - 1) & "/" & Year(Date), "dd/mm/yyyy")
ElseIf (Month(Date) - 1) = 2 Then
    dtpFin.Value = Format("28/" & (Month(Date) - 1) & "/" & Year(Date), "dd/mm/yyyy")
End If
frameFechas.Visible = False

'llena el combo de tipo de factura
rsTipFac.Open "select * from TipoFacturasCpras", cn, adOpenDynamic, adLockReadOnly, adCmdText
Set cbTipo.DataSource = rsTipFac
Set cbTipo.RowSource = rsTipFac
cbTipo.ListField = "Descripcion"
cbTipo.BoundColumn = "idTipo"
cbTipo.BoundText = 1

Option1.Value = True
Option2.Value = False
cbTipo.Enabled = False
End Sub
Private Sub Form_Unload(cancel As Integer)
If rsFact.State = 1 Then
    rsFact.Close
    Set rsFact = Nothing
End If
If rsTipFac.State = 1 Then
    rsTipFac.Close
    Set rsTipFac = Nothing
End If
End Sub
Private Sub optFechas_Click()
optTodo.Value = False
frameFechas.Visible = True
End Sub
Private Sub Option1_Click()
Option2.Value = False
cbTipo.Enabled = False
End Sub
Private Sub Option1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Option2.SetFocus
End If
End Sub
Private Sub Option2_Click()
Option1.Value = False
cbTipo.Enabled = True
End Sub

Private Sub Option2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cbTipo.SetFocus
End If
End Sub

Private Sub optTodo_Click()
optFechas.Value = False
frameFechas.Visible = False
End Sub
