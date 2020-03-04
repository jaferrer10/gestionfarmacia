VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Begin VB.Form frmImportarDatos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar Datos del Sistema de Farmacia ..."
   ClientHeight    =   4710
   ClientLeft      =   3735
   ClientTop       =   2115
   ClientWidth     =   8325
   Icon            =   "frmImportarDatos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   8325
   Begin MSAdodcLib.Adodc AdoDBF 
      Height          =   330
      Left            =   360
      Top             =   240
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=TablasIPG;Mode=Read"
      OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=TablasIPG;Mode=Read"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "TablasIpg"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin AIFCmp1.asxPowerBanner lblCartel 
      Height          =   615
      Left            =   1080
      Top             =   2160
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1085
      FormatString    =   "AGUARDE UNOS MINUTOS - PROCESANDO"
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
   Begin AIFCmp1.asxPowerButton cmdCancelar 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   4320
      TabIndex        =   0
      Top             =   3240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      Picture         =   "frmImportarDatos.frx":0742
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
   End
   Begin AIFCmp1.asxPowerButton cmdAceptar 
      Height          =   735
      Left            =   2160
      TabIndex        =   1
      Top             =   3240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      Picture         =   "frmImportarDatos.frx":0B94
      Caption         =   "Importar datos"
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
   Begin VB.Image Image3 
      Height          =   720
      Left            =   5640
      Picture         =   "frmImportarDatos.frx":0FE6
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   840
   End
   Begin VB.Image Image2 
      Height          =   840
      Left            =   1800
      Picture         =   "frmImportarDatos.frx":1428
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   1080
      Left            =   3360
      Picture         =   "frmImportarDatos.frx":186A
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1680
   End
End
Attribute VB_Name = "frmImportarDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsProductos As New ADODB.Recordset
Private rsTabla As New ADODB.Recordset
Private cnIpg As New Connection

Private Sub cmdAceptar_Click()

lblCartel.Visible = True
cmdCancelar.Enabled = False
cmdAceptar.Enabled = False

On Error GoTo HayError

Me.MousePointer = 11

'Abre la tabla del sistema ipg atravez del origen de datos TablasIPG

cnIpg.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=TablasIPG;Mode=Read"

rsTabla.Open "select * from f10t01", cnIpg, adOpenDynamic, adLockOptimistic, adCmdText

'borra todo el contenido de la tabla productos
cn.Execute "delete from productos"

rsProductos.Open "select * from productos", cn, adOpenDynamic, adLockOptimistic, adCmdText

rsTabla.MoveFirst
Do While rsTabla.EOF = False
    rsProductos.AddNew
    rsProductos!descripcion = rsTabla!nomprd
    rsProductos!troquel = rsTabla!codbar
    rsProductos!precio = rsTabla!prcvta
    rsTabla.MoveNext
Loop
rsTabla.Close
Set rsTabla = Nothing

lblCartel.FormatString = "PROCESO TERMINADO !"
cmdCancelar.Enabled = True
cmdCancelar.Caption = "&Salir"
Me.MousePointer = 1

HayError:
If Err.Number Then
    MsgBox "Se ha producido el siguiente error:" & vbCrLf & _
            Err.Number & ", " & Err.Description, vbCritical, "Error del Sistema ..."
End If
cmdCancelar.Enabled = True
If rsTabla.State = 1 Then
    rsTabla.Close
End If
If rsProductos.State = 1 Then
    rsProductos.Close
End If
On Error Resume Next

End Sub
Private Sub cmdCancelar_Click()
Unload Me
End Sub
Private Sub Form_Load()
Me.Top = 800
Me.Left = 1500
lblCartel.Visible = False
End Sub
