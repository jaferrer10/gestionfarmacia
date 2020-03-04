VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmVistaPrevia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vista Previa"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6885
   Icon            =   "FrmVistaPrevia.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4935
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   569
      Left            =   3360
      MouseIcon       =   "FrmVistaPrevia.frx":0442
      Picture         =   "FrmVistaPrevia.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   1397
   End
   Begin CRVIEWERLibCtl.CRViewer crvwVistaPrevia 
      Height          =   5775
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   7575
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmVistaPrevia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    With crvwVistaPrevia
        .ReportSource = rptGeneral
        .Zoom 100
        .ViewReport
    End With
End Sub
Private Sub Form_Resize()
    crvwVistaPrevia.Top = 0
    crvwVistaPrevia.Left = 0
    
    cmdSalir.Top = ScaleHeight - 15 - cmdSalir.Height
    cmdSalir.Left = ScaleWidth - 15 - cmdSalir.Width
    If ScaleHeight = 0 Or ScaleWidth = 0 Then Exit Sub
    crvwVistaPrevia.Height = ScaleHeight - cmdSalir.Height - 30
    crvwVistaPrevia.Width = ScaleWidth - 15
End Sub
