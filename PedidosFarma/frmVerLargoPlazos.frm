VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVerLargoPlazos 
   Caption         =   "Facturas de Largo Plazos..."
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9270
   Icon            =   "frmVerLargoPlazos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9270
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameDatos 
      Height          =   4335
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   9015
      Begin MSDataGridLib.DataGrid dtgarchivo 
         Height          =   3975
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   7011
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   12632256
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
            Caption         =   "Fecha Vto"
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
         BeginProperty Column06 
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
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1005,165
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   434,835
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   945,071
            EndProperty
            BeginProperty Column05 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1874,835
            EndProperty
         EndProperty
      End
      Begin AIFCmp1.asxPowerButton cmdVerSubTot 
         Height          =   690
         Left            =   8040
         TabIndex        =   9
         ToolTipText     =   "Concluye con la vista de datos..."
         Top             =   3240
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   1217
         BorderStyle     =   4
         Picture         =   "frmVerLargoPlazos.frx":08CA
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
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      Begin VB.Frame frameTotales 
         Height          =   900
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   4650
         Begin VB.TextBox txtotal 
            Alignment       =   1  'Right Justify
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
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   500
            Width           =   1215
         End
         Begin VB.TextBox txtReg 
            Alignment       =   1  'Right Justify
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
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   150
            Width           =   1215
         End
         Begin VB.Label lblTotal 
            AutoSize        =   -1  'True
            Caption         =   "Total de Importes:"
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
            Left            =   50
            TabIndex        =   15
            Top             =   500
            Width           =   1905
         End
         Begin VB.Label lblReg 
            AutoSize        =   -1  'True
            Caption         =   "Total de registros encontrados:"
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
            Left            =   50
            TabIndex        =   13
            Top             =   150
            Width           =   3270
         End
      End
      Begin VB.CheckBox ChkTodo 
         Caption         =   "Ver todos ..."
         Height          =   255
         Left            =   6000
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpIni 
         Height          =   375
         Left            =   1200
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
         Format          =   143523841
         CurrentDate     =   39392
      End
      Begin MSComCtl2.DTPicker dtpFin 
         Height          =   375
         Left            =   4080
         TabIndex        =   2
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
         Format          =   143523841
         CurrentDate     =   39392
      End
      Begin AIFCmp1.asxPowerButton cmdImprimir 
         Height          =   735
         Left            =   5040
         TabIndex        =   6
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Picture         =   "frmVerLargoPlazos.frx":0E64
         Caption         =   "&Ver      "
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
         Height          =   735
         Left            =   7080
         TabIndex        =   7
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Picture         =   "frmVerLargoPlazos.frx":173E
         Caption         =   "&Cancelar     "
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
         Left            =   3120
         TabIndex        =   4
         Top             =   360
         Width           =   690
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
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmVerLargoPlazos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsFact As New ADODB.Recordset

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdImprimir_Click()
On Error GoTo VerError
If ChkTodo.Value = 1 Then
    rsFact.Open "select fecha, fechavto, numero, tipo, importe, f.observaciones, estado " & _
                "from facturascompras f inner join Proveedores p " & _
                "on f.idproveedor = p.idproveedor " & _
                "where f.idproveedor = " & frmCompras.dtcProveedor.BoundText & " and fechavto <> fecha " & _
                " order by fechavto desc, numero", cn, adOpenDynamic, adLockReadOnly, adCmdText
Else
    rsFact.Open "select fecha, fechavto, numero, tipo, importe, f.observaciones, estado " & _
            "from facturascompras f inner join Proveedores p " & _
            "on f.idproveedor = p.idproveedor " & _
            "WHERE f.idproveedor = " & frmCompras.dtcProveedor.BoundText & _
            " and Fechavto between #" & Format(dtpIni.Value, "mm/dd/yyyy") & _
            "# and #" & Format(dtpFin.Value, "mm/dd/yyyy") & "# and fechavto <> fecha order by fechavto desc, numero", cn, adOpenDynamic, adLockReadOnly, adCmdText
End If
If rsFact.RecordCount = 0 Then
    MsgBox "NO HAY IMFORMACION PARA LA IMPRESION...", vbInformation, "ATENCION !"
    rsFact.Close
    Set rsFact = Nothing
    Exit Sub
End If

Frame1.Enabled = False

frameTotales.Visible = True
txtReg.Text = rsFact.RecordCount

'hace la sumatoria de importes del rango de facturas
Dim vTot As Double
vTot = 0
rsFact.MoveFirst
Do While rsFact.EOF = False
    vTot = vTot + rsFact!importe
    rsFact.MoveNext
Loop
txtotal.Text = vTot

frameDatos.Visible = True

Set dtgarchivo.DataSource = rsFact
dtgarchivo.Refresh
Exit Sub

VerError:
    MsgBox Err.Description + Chr(13) + "Seleccione nuevamente el proveedor.", vbCritical, "Atención ...!"
    Exit Sub

End Sub

Private Sub cmdVerSubTot_Click()
Frame1.Enabled = True
frameDatos.Visible = False
rsFact.Close
Set rsFact = Nothing

End Sub

Private Sub Form_Load()
frameTotales.Visible = False
frameDatos.Visible = False
Frame1.Enabled = True

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

End Sub

Private Sub Form_Unload(cancel As Integer)
If rsFact.State = 1 Then
    rsFact.Close
End If

End Sub
