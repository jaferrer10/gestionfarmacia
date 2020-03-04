VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPrestamosFarmacias 
   Caption         =   "Prestamos de Productos entre Farmacias ..."
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15000
   Icon            =   "frmPrestamosFarmacias.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7830
   ScaleWidth      =   15000
   Begin VB.Frame Frame2 
      Caption         =   "Archivo de Prestamos"
      Height          =   5295
      Left            =   120
      TabIndex        =   18
      Top             =   2400
      Width           =   14775
      Begin AIFCmp1.asxPowerButton cmdBorrar 
         Height          =   495
         Left            =   13200
         TabIndex        =   19
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Picture         =   "frmPrestamosFarmacias.frx":058A
         Caption         =   "&Borrar"
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
         TextColor       =   192
      End
      Begin MSDataGridLib.DataGrid dtgArchivo 
         Height          =   4935
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   8705
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   49344
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
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "farpresta"
            Caption         =   "PRESTADOR"
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
            DataField       =   "fardeudora"
            Caption         =   "DEUDOR"
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
            DataField       =   "Estado"
            Caption         =   "ESTADO"
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
            DataField       =   "devolucion"
            Caption         =   "Devolución"
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
            DataField       =   "descripcion"
            Caption         =   "Descripción Producto"
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
            DataField       =   "cantidad"
            Caption         =   "Cant"
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
            DataField       =   "contacto"
            Caption         =   "Contacto"
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
            DataField       =   "codbarra"
            Caption         =   "CodBarra"
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
            BeginProperty Column00 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1844,787
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2190,047
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1365,165
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2115,213
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   3435,024
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   480,189
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   2055,118
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   6795,213
            EndProperty
         EndProperty
      End
      Begin AIFCmp1.asxPowerButton cmdDevol 
         Height          =   495
         Left            =   13200
         TabIndex        =   21
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Picture         =   "frmPrestamosFarmacias.frx":0B24
         Caption         =   "&Devolver"
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
         TextColor       =   32768
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Prestamo"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14775
      Begin VB.TextBox txtObser 
         Height          =   285
         Left            =   1200
         TabIndex        =   17
         Top             =   1920
         Width           =   6735
      End
      Begin AIFCmp1.asxPowerButton cmdPrestar 
         Height          =   405
         Left            =   8280
         TabIndex        =   14
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   714
         Picture         =   "frmPrestamosFarmacias.frx":10BE
         PictureDown     =   "frmPrestamosFarmacias.frx":1658
         Caption         =   "&Prestar"
         CaptionAlignment=   5
         CaptionTextAlignment=   1
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
         TextColor       =   255
      End
      Begin VB.TextBox txtContacto 
         Height          =   285
         Left            =   4200
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtCantidad 
         Height          =   285
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   11
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtDescripcion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4200
         MaxLength       =   50
         TabIndex        =   9
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   1
         Top             =   960
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   320
         Left            =   8280
         TabIndex        =   7
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   21430273
         CurrentDate     =   40333
      End
      Begin MSDataListLib.DataCombo dtcPrestador 
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   "DataCombo1"
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx3 
         Height          =   240
         Left            =   8280
         TabIndex        =   4
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   423
         Caption         =   "Fecha Prestamo"
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx2 
         Height          =   240
         Left            =   4200
         TabIndex        =   3
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   423
         Caption         =   "DEUDOR"
      End
      Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx1 
         Height          =   240
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   423
         Caption         =   "PRESTADOR"
      End
      Begin MSDataListLib.DataCombo dtcDeudor 
         Height          =   315
         Left            =   4200
         TabIndex        =   6
         Top             =   480
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   "DataCombo1"
      End
      Begin AIFCmp1.asxPowerButton cmdSalir 
         Cancel          =   -1  'True
         Height          =   405
         Left            =   9840
         TabIndex        =   15
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   714
         Picture         =   "frmPrestamosFarmacias.frx":1BF2
         Caption         =   "&Salir"
         CaptionAlignment=   5
         CaptionTextAlignment=   1
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
      Begin AIFCmp1.asxPowerButton cmdAgrPro 
         Height          =   495
         Left            =   8160
         TabIndex        =   22
         ToolTipText     =   "Agrega Producto no registrado"
         Top             =   840
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
         BorderStyle     =   4
         Picture         =   "frmPrestamosFarmacias.frx":218C
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Observac:"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Contacto:"
         Height          =   195
         Left            =   3240
         TabIndex        =   12
         Top             =   1440
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cod.Barra:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmPrestamosFarmacias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsPrestamos As New ADODB.Recordset
Private rsFarmacias As New ADODB.Recordset
Private rsProductos As New ADODB.Recordset
Private idPres As Integer

Private Sub cmdAgrPro_Click()
frmNuevoProductoPrestamo.Show vbModal

txtCodigo.SetFocus

End Sub

Private Sub cmdBorrar_Click()
SioNo = MsgBox("ESTA SEGURO DE ELIMINAR EL REGISTRO SELECCIONADO ?", vbInformation + vbYesNo, "ELIMINANDO REGISTRO...")
If SioNo = vbYes Then
    If rsPrestamos.EOF = False Then
        idPres = rsPrestamos!idprestamo
        cn.Execute "delete from prestamosinterfarmacias where idprestamo = " & idPres
        rsPrestamos.Requery
        dtgArchivo.Refresh
    Else
        MsgBox "NO HAY REGISTROS PARA BORRAR !", vbExclamation, "ATENCION !"
    End If
End If
End Sub
Private Sub cmdDevol_Click()
SioNo = MsgBox("ESTA SEGURO DE DEVOLVER EL PRODUCTO ?", vbInformation + vbYesNo, "DEVOLVIENDO PRODUCTO...")
If SioNo = vbYes Then
    idPres = rsPrestamos!idprestamo
    cn.Execute "update prestamosinterfarmacias set estado = 'DEVUELTO', devolucion = #" & (Format(dtpFecha.Value, "mm/dd/yyyy")) & " " & (Format(Time, "hh:mm")) & "# where idprestamo = " & idPres
    rsPrestamos.Update
    rsPrestamos.Requery
    dtgArchivo.Refresh
End If
End Sub
Private Sub cmdPrestar_Click()
If Len(txtCantidad.Text) = 0 Then
    MsgBox "Debe colocar la cantidad prestada...", vbExclamation, "Atención !"
    txtCantidad.SetFocus
    Exit Sub
End If
If Len(txtDescripcion.Text) = 0 Then
    MsgBox "Falta la descripción del producto ...", vbCritical, "Atención !"
    txtDescripcion.SetFocus
    Exit Sub
End If

If dtcPrestador.Text = dtcDeudor.Text Then
    MsgBox "EL PRESTADOR NO PUEDE SER IGUAL AL DEUDOR, VERIFIQUE ...", vbCritical, "ATENCION !"
    dtcDeudor.SetFocus
    Exit Sub
End If

'agrega el registro del prestamo
strSQL = "insert into prestamosinterfarmacias (FarPresta,FarDeudora,codbarra,fecha,cantidad,contacto,estado,observaciones) values ('" & dtcPrestador.Text & "','" & dtcDeudor.Text & "','" & txtCodigo.Text & "',#" & (Format(dtpFecha.Value, "mm/dd/yyyy")) & " " & (Format(Time, "hh:mm")) & "#," & txtCantidad.Text & ",'" & txtContacto.Text & "','DEBE'" & ",'" & txtObser.Text & "')"

cn.Execute strSQL

rsPrestamos.Requery
rsPrestamos.Update

Set dtgArchivo.DataSource = rsPrestamos
Set dtcPrestador.RowSource = rsFarmacias
dtgArchivo.Refresh

txtCodigo.Text = ""
txtDescripcion.Text = ""
txtCantidad.Text = ""
'txtContacto.Text = ""
txtObser.Text = ""

txtCodigo.SetFocus
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub dtcDeudor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    txtCodigo.SetFocus
End If
End Sub
Private Sub dtcPrestador_Change()
rsFarmacias.Find "idfarmacia = " & dtcPrestador.BoundText, , adSearchForward, 1
txtContacto.Text = rsFarmacias!contacto
End Sub
Private Sub dtcPrestador_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    dtcDeudor.SetFocus
End If
End Sub
Private Sub Form_Load()

Me.Width = 15120
Me.Height = 8340
Me.Top = 200
Me.Left = 100

rsPrestamos.Open "SELECT PrestamosInterFarmacias.Fecha, PrestamosInterFarmacias.idPrestamo, PrestamosInterFarmacias.FarPresta, PrestamosInterFarmacias.FarDeudora, PrestamosInterFarmacias.codBarra, Productos.Descripcion, PrestamosInterFarmacias.Cantidad, PrestamosInterFarmacias.Contacto, PrestamosInterFarmacias.Estado,devolucion, PrestamosInterFarmacias.Observaciones " & _
                    " FROM PrestamosInterFarmacias INNER JOIN Productos ON PrestamosInterFarmacias.codBarra = Productos.troquel " & _
                    " order by fecha desc", cn, adOpenDynamic, adLockOptimistic, adCmdText
                 
rsFarmacias.Open "select * from farmacias order by nombre", cn, adOpenDynamic, adCmdText

dtpFecha.Value = Date
txtCantidad.Text = 1

'Llena el combo de farmacia prestadora
Set dtcPrestador.DataSource = rsFarmacias
Set dtcPrestador.RowSource = rsFarmacias
dtcPrestador.ListField = "Nombre"
dtcPrestador.BoundColumn = "idfarmacia"
dtcPrestador.BoundText = 1
txtContacto.Text = rsFarmacias!contacto

Set dtcDeudor.DataSource = rsFarmacias
Set dtcDeudor.RowSource = rsFarmacias
dtcDeudor.ListField = "Nombre"
dtcDeudor.BoundColumn = "idfarmacia"
dtcDeudor.BoundText = 2

Set dtgArchivo.DataSource = rsPrestamos
dtgArchivo.Refresh

End Sub
Private Sub Form_Unload(cancel As Integer)
rsFarmacias.Close
If rsPrestamos.State = 1 Then
    rsPrestamos.Close
End If
End Sub
Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtContacto.SetFocus
End If
End Sub
Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtCantidad.SetFocus
End If
End Sub
Private Sub txtCodigo_LostFocus()
If Len(txtCodigo.Text) > 0 Then
    If rsProductos.State = 1 Then
        rsProductos.Close
    Else
        rsProductos.Open "select troquel,descripcion from productos where troquel = '" & txtCodigo.Text & "'", cn, adOpenDynamic, adLockReadOnly, adCmdText
        If rsProductos.RecordCount = 0 Then
            MsgBox "Código ingresado no existe en el archivo...!", vbInformation, "Código inexistente..."
            frmNuevoProductoPrestamo.Show vbModal
            txtCodigo.SetFocus
            SendKeys "{end}+{home}"
        Else
            txtDescripcion.Text = rsProductos!descripcion
            txtCantidad.Text = 1
            txtCantidad.SetFocus
            SendKeys "{end}+{home}"
        End If
        rsProductos.Close
    End If
End If
End Sub
Private Sub txtContacto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtObser.SetFocus
End If
End Sub
Private Sub txtObser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdPrestar.SetFocus
End If
End Sub
