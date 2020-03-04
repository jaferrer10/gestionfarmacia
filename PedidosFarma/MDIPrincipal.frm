VERSION 5.00
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "Gestion de Pedidos - V.20 - 082017"
   ClientHeight    =   8985
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11190
   Icon            =   "MDIPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin AIFCmp1.asxToolbar asxToolbar1 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   1508
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonCount     =   9
      SolidChecked    =   -1  'True
      ButtonCaption1  =   "Pedir"
      ButtonDescription1=   "Agrega items a un pedido"
      ButtonKey1      =   "WINDW01G"
      ButtonPicture1  =   "MDIPrincipal.frx":014A
      ButtonPictureOver1=   "MDIPrincipal.frx":0D9C
      ButtonToolTipText1=   "Agrega items a un pedido"
      ButtonCaption2  =   "Ventas"
      ButtonDescription2=   "Registro de ventas diarias"
      ButtonKey2      =   "DOLLR00F"
      ButtonPicture2  =   "MDIPrincipal.frx":19EE
      ButtonToolTipText2=   "Registro de Ventas"
      ButtonCaption3  =   "Compras"
      ButtonDescription3=   "Carga de facturas de compras"
      ButtonKey3      =   "CRDFLE06"
      ButtonPicture3  =   "MDIPrincipal.frx":2640
      ButtonToolTipText3=   "Registro de compras"
      ButtonCaption4  =   "Egresos/Gastos"
      ButtonKey4      =   "WATER"
      ButtonPicture4  =   "MDIPrincipal.frx":3292
      ButtonToolTipText4=   "Registro de todo tipo de Egresos de Dinero"
      ButtonCaption5  =   "Clientes"
      ButtonDescription5=   "Gestion de clientes de farmacia"
      ButtonKey5      =   "fcabin02"
      ButtonPicture5  =   "MDIPrincipal.frx":3EE4
      ButtonToolTipText5=   "Gestion de clientes"
      ButtonCaption6  =   "Ctrl Precios"
      ButtonKey6      =   "BANK00A"
      ButtonPicture6  =   "MDIPrincipal.frx":4B36
      ButtonToolTipText6=   "Control de precios de productos"
      ButtonCaption7  =   "Agenda    "
      ButtonDescription7=   "Agenda de clientes y personas varias"
      ButtonKey7      =   "Avant Browser"
      ButtonPicture7  =   "MDIPrincipal.frx":5788
      ButtonToolTipText7=   "Agenda de personas"
      ButtonCaption8  =   "Facturación"
      ButtonDescription8=   "Ejecuta facturacion"
      ButtonKey8      =   "DOLLR03F"
      ButtonPicture8  =   "MDIPrincipal.frx":72DA
      ButtonPictureOver8=   "MDIPrincipal.frx":7F2C
      ButtonToolTipText8=   "Facturacion"
      ButtonCaption9  =   "Salir"
      ButtonKey9      =   "SCDCNCLL"
      ButtonPicture9  =   "MDIPrincipal.frx":8B7E
      ButtonToolTipText9=   "Sale del Sistema"
   End
   Begin VB.Menu mnArc 
      Caption         =   "&Archivos"
      Begin VB.Menu mnAr_cruce 
         Caption         =   "Control y cruce de resúmenes de compra"
      End
      Begin VB.Menu mnDepPed 
         Caption         =   "Depuracion Archivo de Pedidos"
      End
      Begin VB.Menu mnDepFact 
         Caption         =   "Depuracion Archivo de Facturas Proveedores"
      End
      Begin VB.Menu mnArCtrlCajas 
         Caption         =   "Archivo de Control de Cajas"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnArNC 
         Caption         =   "Nota de Creditos a Proveedores"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnArcProv 
         Caption         =   "Proveedores"
      End
      Begin VB.Menu mnFacImp 
         Caption         =   "Facturas Impagas a Proveedores"
      End
      Begin VB.Menu mnUsuarios 
         Caption         =   "Usuarios"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu mnCarga 
      Caption         =   "Carga de Datos"
      Begin VB.Menu mnCarPed 
         Caption         =   "Pedidos"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnPedCaja 
         Caption         =   "Gestión de Caja Diaria"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnCarFAc 
         Caption         =   "Facturas de Proveedores"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnCargaExtra 
         Caption         =   "Ventas Extraordinarias"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnCarVtas 
         Caption         =   "Ventas Diarias"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnCarEg 
         Caption         =   "Egresos"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnCarPre 
         Caption         =   "Control de Precios"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnCarCli 
         Caption         =   "Clientes"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnCarPto 
         Caption         =   "Punto de Control de Cajas"
      End
      Begin VB.Menu mnCarTroquel 
         Caption         =   "Troqueles disponibles"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnPresFar 
         Caption         =   "Prestamos entre Farmacias"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnDevo 
         Caption         =   "Devoluciones registro"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnInformes 
      Caption         =   "Informes"
      Begin VB.Menu mnInfVtas 
         Caption         =   "Informe de Ventas"
      End
      Begin VB.Menu mninfVE 
         Caption         =   "Informe de Ventas Extraordinarias"
      End
      Begin VB.Menu mnInfEg 
         Caption         =   "Informe de Egresos"
      End
      Begin VB.Menu mnInfGcias 
         Caption         =   "Informe de Ganancias detallado"
      End
      Begin VB.Menu mnInfGciasM 
         Caption         =   "Informe de Ganancias Mensuales"
      End
      Begin VB.Menu mnInfCpras 
         Caption         =   "Listado de Facturas de Compras"
      End
   End
   Begin VB.Menu mnSistemas 
      Caption         =   "Sistemas"
      Begin VB.Menu mnSisBackup 
         Caption         =   "Copias de Seguridad"
      End
      Begin VB.Menu mnSisRestore 
         Caption         =   "Restauracion de Copias de Seguridad"
      End
      Begin VB.Menu mnSisImportar 
         Caption         =   "Importar datos al Sistema"
      End
      Begin VB.Menu mnRef 
         Caption         =   "Corregir Referencias"
      End
   End
   Begin VB.Menu mnSalir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsCumples As New ADODB.Recordset
Private Sub asxToolbar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
Select Case ButtonIndex
    Case Is = 1
            frmPideClave.Show vbModal
            If TempNivel = 0 Then
                Exit Sub
            Else
                frmPedidos.Show
            End If
    Case Is = 2
            frmRegistroVentas.Show
    Case Is = 3
            frmPideClave.Show vbModal
            If TempNivel = 0 Then
                Exit Sub
            Else
                frmCompras.Show
            End If
    Case Is = 4
            frmEgresos.Show
    Case Is = 5
            frmPideClave.Show vbModal
            If TempNivel = 0 Then
                Exit Sub
            Else
                frmGestionClientes.Show
            End If
    Case Is = 6
            frmControlPrecios.Show
    Case Is = 7
            frmAgenda.Show
    Case Is = 8
            frmFacturacion.Show
    Case Is = 9
            SioNo = MsgBox("ESTA SEGURO DE SALIR DEL SISTEMA ???", vbInformation + vbYesNo, "Salida del Sistema...")
            If SioNo = vbYes Then End
End Select
End Sub
Private Sub MDIForm_Load()
cbd = "serena24"
cn.CursorLocation = adUseClient 'permite utilizar el recorset con eventos en la grilla
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.path & "\Pedidos.mdb;Mode=ReadWrite;Persist Security Info=False;Jet OLEDB:Database Password= " & cbd
'cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.path & "\Pedidos.mdb;Mode=Share Deny Read|Share Deny Write;Persist Security Info=False; Jet OLEDB:Database Password= " & cbd

'establesco los dias desde y hasta se filtraran las fechas de nacimientos
'para tirar el listado de salutaciones a clientes
Dim DiasDesde As Integer
Dim diasHasta As Integer
If Day(Date) <= 3 Then
    DiasDesde = 1
Else
    DiasDesde = (Day(Date) - 2)
End If
diasHasta = (Day(Date) + 2)

rsCumples.Open "select * from clientes where Day(fechanac) >= " & DiasDesde & " and Day(fechanac) <= " & diasHasta & " and Month(fechanac) = " & Month(Date) & " order by fechanac, apellido", cn, adOpenDynamic, adLockReadOnly, adCmdText
'--------------------------------------------------------------------------
If rsCumples.RecordCount > 0 Then
    frmCumpleaños.Show
End If
rsCumples.Close
Set rsCumples = Nothing
End Sub

Private Sub mnAr_cruce_Click()
frmPideClave.Show vbModal
If TempNivel > 0 Then
    frmCruce.Show
End If
End Sub
Private Sub mnArcProv_Click()
frmGestionProveedores.Show vbModal
End Sub
Private Sub mnArCtrlCajas_Click()
frmConsCajas.Show
End Sub
Private Sub mnArNC_Click()
frmNCdroguerias.Show vbModal
End Sub
Private Sub mnCarCli_Click()
frmPideClave.Show vbModal
If TempNivel = 0 Then
    Exit Sub
Else
    frmGestionClientes.Show
End If
End Sub
Private Sub mnCarEg_Click()
frmEgresos.Show
End Sub
Private Sub mnCarFAc_Click()
frmPideClave.Show vbModal
If TempNivel = 0 Then
    Exit Sub
Else
    frmCompras.Show
End If
End Sub
Private Sub mnCargaExtra_Click()
frmVentasExtras.Show
End Sub
Private Sub mnCarPed_Click()
frmPedidos.Show
End Sub
Private Sub mnCarPre_Click()
frmControlPrecios.Show
End Sub
Private Sub mnCarPto_Click()
frmPuntoControl.Show
End Sub
Private Sub mnCarTroquel_Click()
frmTroqueles.Show
End Sub
Private Sub mnCarVtas_Click()
frmRegistroVentas.Show
End Sub
Private Sub mnDepFact_Click()
frmDepuCpras.Show
End Sub
Private Sub mnDepPed_Click()
frmDepuPedidos.Show
End Sub

Private Sub mnDevo_Click()
frmPideClave.Show vbModal
If TempNivel = 0 Then
    Exit Sub
Else
    frmDevoluciones.Show
End If
End Sub
Private Sub mnFacImp_Click()
frmFacturasImpagas.Show
End Sub
Private Sub mnInfCpras_Click()
frmListadoFacturas.Show
End Sub
Private Sub mnInfEg_Click()
frmInfEgresos.Show
End Sub
Private Sub mnInfGcias_Click()
frmInfGanancias.Show
End Sub
Private Sub mnInfGciasM_Click()
frmInfGciasMensual.Show
End Sub
Private Sub mninfVE_Click()
frmInfVentasExtras.Show
End Sub
Private Sub mnInfVtas_Click()
frmInfVentasMensuales.Show
End Sub
Private Sub mnPedCaja_Click()
frmGestionCaja.Show
End Sub
Private Sub mnPresFar_Click()
frmPrestamosFarmacias.Show
End Sub

Private Sub mnRef_Click()
frmReferencias.Show vbModal
End Sub

Private Sub mnSalir_Click()
 SioNo = MsgBox("ESTA SEGURO DE SALIR DEL SISTEMA ???", vbInformation + vbYesNo, "Salida del Sistema...")
            If SioNo = vbYes Then End
End Sub
Private Sub mnSisBackup_Click()
FrmUtilidadesBaseDatosRealizarCS.Show
End Sub
Private Sub mnSisImportar_Click()
frmImportarDatos.Show
End Sub
Private Sub mnSisRestore_Click()
FrmUtilidadesBaseDatosRestaurarCS.Show
End Sub
Private Sub mnUsuarios_Click()
frmUsuarios.Show
End Sub
