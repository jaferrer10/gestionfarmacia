VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Begin VB.Form frmFacturasImpagas 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aviso de Facturas Impagas a vencer de Proveedores ..."
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12600
   Icon            =   "frmFacturasImpagas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   12600
   Begin VB.Frame frameTotales 
      Caption         =   "Informacion de Totales"
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   5640
      Width           =   12375
      Begin AIFCmp1.asxPowerBanner lblProveedor 
         Height          =   375
         Left            =   1440
         Top             =   300
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   661
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
      Begin AIFCmp1.asxLabel asxLabel1 
         Height          =   225
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Proveedor :"
         AutoSize        =   -1  'True
         UseMnemonic     =   -1  'True
         MouseIcon       =   "frmFacturasImpagas.frx":6852
      End
      Begin AIFCmp1.asxLabel asxLabel2 
         Height          =   225
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Total Deuda:"
         AutoSize        =   -1  'True
         UseMnemonic     =   -1  'True
         MouseIcon       =   "frmFacturasImpagas.frx":6B6C
      End
      Begin AIFCmp1.asxPowerBanner lblDeuda 
         Height          =   375
         Left            =   1440
         Top             =   840
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   661
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
         TextColor       =   255
      End
   End
   Begin VB.Frame FrameDatos 
      Caption         =   "Archivo Facturas de Compras"
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12375
      Begin AIFCmp1.asxPowerButton cmdPagar 
         Height          =   495
         Left            =   10920
         TabIndex        =   3
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&Pagar"
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
      Begin MSDataGridLib.DataGrid dtgImpagas 
         Height          =   4815
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   8493
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "estado"
            Caption         =   "Estado"
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
            DataField       =   "nombre"
            Caption         =   "Nombre Proveedor"
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
            DataField       =   "fecha"
            Caption         =   "Fecha"
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
            DataField       =   "numero"
            Caption         =   "Nº Factura"
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
            DataField       =   "tipo"
            Caption         =   "Tipo"
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
         BeginProperty Column06 
            DataField       =   "depositado"
            Caption         =   "Depositado"
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
         BeginProperty Column07 
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
         BeginProperty Column08 
            DataField       =   "idproveedor"
            Caption         =   "idProveedor"
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
         BeginProperty Column09 
            DataField       =   "usuario"
            Caption         =   "Usuario"
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
               ColumnWidth     =   929,764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2624,882
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1275,024
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1244,976
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   434,835
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   824,882
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   599,811
            EndProperty
            BeginProperty Column09 
            EndProperty
         EndProperty
      End
      Begin AIFCmp1.asxPowerButton cmdImprimir 
         Height          =   495
         Left            =   10920
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&Imprimir"
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
      Begin AIFCmp1.asxPowerButton cmdSalir 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   10920
         TabIndex        =   5
         Top             =   1800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&Salir"
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
Attribute VB_Name = "frmFacturasImpagas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsImpagas As New Recordset
Private rsTotales As New Recordset
Private vDeuda As Double
Private vidPro As Integer
Private vidCpra As Integer
Private Sub cmdPagar_Click()
vidCpra = rsImpagas!idcompra
cn.Execute "update facturascompras set Estado = 'P' where idcompra = " & vidCpra
rsImpagas.Requery
Call CalculaTotales
MsgBox "SE HA CAMBIADO EL ESTADO DE AL FACTURA A PAGADA !", vbInformation, "Proceso exitoso"

End Sub
Private Sub cmdSalir_Click()
rsImpagas.Close
Unload Me
End Sub
Private Sub dtgImpagas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call CalculaTotales
End Sub
Private Sub Form_Load()

'Tabla de facturas de compras
rsImpagas.Open "SELECT idcompra, FacturasCompras.idProveedor, Proveedores.Nombre, FacturasCompras.Fecha, FacturasCompras.Numero, FacturasCompras.Tipo, FacturasCompras.Importe, FacturasCompras.Depositado, FacturasCompras.Observaciones, FacturasCompras.Estado, FacturasCompras.Usuario " & _
                "FROM FacturasCompras INNER JOIN Proveedores ON FacturasCompras.idProveedor = Proveedores.idProveedor" & _
                " where Estado = 'D' order by facturascompras.idProveedor, fecha ", cn, adOpenDynamic, adLockOptimistic, adCmdText
rsImpagas.MoveFirst
Set dtgImpagas.DataSource = rsImpagas
dtgImpagas.Refresh

lblProveedor.FormatString = rsImpagas!nombre
vDeuda = 0
vidPro = rsImpagas!idproveedor
rsImpagas.MoveFirst
Call CalculaTotales
End Sub
Private Sub CalculaTotales()
If rsImpagas.EOF = False Then
    If rsTotales.State = 1 Then
        rsTotales.Close
    End If
    lblProveedor.FormatString = rsImpagas!nombre
    vidPro = rsImpagas!idproveedor
    rsTotales.Open "select idproveedor, sum(importe) as Total from facturascompras where idproveedor = " & vidPro & " and Estado = 'D' group by idproveedor", cn, adOpenDynamic, adLockReadOnly, adCmdText
    vDeuda = rsTotales!total
    lblDeuda.FormatString = vDeuda
    rsTotales.Close
End If
End Sub
