VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCliente 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Clienti"
   ClientHeight    =   7290
   ClientLeft      =   3435
   ClientTop       =   1695
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   8745
   Begin MSComctlLib.ImageList ilToolBar 
      Left            =   7440
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":015C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":06A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":1128
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":166C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":1BB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":20F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":2770
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":2DEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":3468
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":38BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":3D10
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":4164
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbCommand 
      Align           =   2  'Align Bottom
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   6660
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   1111
      ButtonWidth     =   1164
      ButtonHeight    =   953
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ilToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuovo"
            Key             =   "N"
            Object.ToolTipText     =   "Nuovo Reclamo"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "prova"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salva"
            Key             =   "S"
            Object.ToolTipText     =   "Salva Reclamo Corrente"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Annulla"
            Key             =   "A"
            Object.ToolTipText     =   "Annulla modifiche su reclamo corrente"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Chiudi"
            Key             =   "C"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cerca"
            Key             =   "G"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stampa"
            Key             =   "P"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "phRight"
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Esci"
            Key             =   "X"
            Object.ToolTipText     =   "Esci"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Shape shpRisposta 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   1455
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   5295
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Capitolato Cliente"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   630
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   3930
   End
   Begin VB.Image iLogo 
      Height          =   450
      Left            =   0
      Picture         =   "frmCliente.frx":45B8
      Top             =   0
      Width           =   450
   End
End
Attribute VB_Name = "frmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mKey As String, rsATC As Recordset
Private fNuovoRecord As Boolean, sStato As String
Private mStatus As Integer
Private Enum CMDS
    cNuovo = 1
    cSalva = 3
    cAnnulla = 5
'    cModifica = 7
    cChiudi = 7
    cCerca = 9
    cStampa = 11
    cEsci = 13
End Enum
Private Enum Status
   cNew = 0
   cMod = 1
   cExists = 2
   cNotExists = 3
End Enum
Dim rsComponente As Recordset, rsN As Recordset
Dim fS As Boolean, dTot As Double, dC As Double, dF As Double, dP As Double
Dim fSalva As Boolean
Dim sSe As String
Public Property Let Status(vValue As Integer)
 Dim iPos As Integer, s As String
   s = Me.Key
   iPos = InStr(s, ",")
   If iPos <> 0 Then
      s = Mid$(s, iPos + 2, 5)
   End If
  
   
   
   If vValue >= cNew And vValue <= cNotExists Then
      mStatus = vValue
   End If
   With tbCommand
    If sFlag = "ATC" Then ' controllo che non possano modificare i dati dal MKG
        Select Case s
            Case "FATTA"
'        If s = "FATTA" Then
'            MsgBox "Fatta"
                .Buttons(cNuovo).Enabled = False
                .Buttons(cSalva).Enabled = False
                .Buttons(cAnnulla).Enabled = False
'                .Buttons(cModifica).Enabled = False
                .Buttons(cChiudi).Enabled = False
                .Buttons(cStampa).Enabled = True
                .Buttons(cCerca).Enabled = True
                .Buttons(cEsci).Enabled = True
            Case "TRASF"
                .Buttons(cNuovo).Enabled = False
                .Buttons(cSalva).Enabled = False
                .Buttons(cAnnulla).Enabled = False
'                .Buttons(cModifica).Enabled = False
                .Buttons(cChiudi).Enabled = False
                .Buttons(cStampa).Enabled = True
                .Buttons(cCerca).Enabled = True
                .Buttons(cEsci).Enabled = True
            
            Case Else
              
                Select Case mStatus
                   Case cNew, cMod
                        .Buttons(cNuovo).Enabled = False
                        .Buttons(cSalva).Enabled = True
                        .Buttons(cAnnulla).Enabled = True
'                        .Buttons(cModifica).Enabled = False
                        .Buttons(cChiudi).Enabled = False
                        .Buttons(cStampa).Enabled = True
                        .Buttons(cCerca).Enabled = True
                        .Buttons(cEsci).Enabled = False
                   Case cExists
                        .Buttons(cNuovo).Enabled = False
                        .Buttons(cSalva).Enabled = False
                        .Buttons(cAnnulla).Enabled = False
'                        .Buttons(cModifica).Enabled = True
                        .Buttons(cChiudi).Enabled = Not (sStato = "C")
                        .Buttons(cStampa).Enabled = True
                        .Buttons(cCerca).Enabled = True
                        .Buttons(cEsci).Enabled = True
                        fS = True
                   Case cNotExists
                        .Buttons(cNuovo).Enabled = True
                        .Buttons(cSalva).Enabled = False
                        .Buttons(cAnnulla).Enabled = False
'                        .Buttons(cModifica).Enabled = False
                        .Buttons(cChiudi).Enabled = False
                        .Buttons(cCerca).Enabled = True
                        .Buttons(cStampa).Enabled = True
                        .Buttons(cEsci).Enabled = True
                        fS = False
                End Select
        End Select
      
    Else
      Select Case mStatus
         Case cNew, cMod
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
'            .Buttons(cModifica).Enabled = False
            .Buttons(cChiudi).Enabled = False
            .Buttons(cCerca).Enabled = True
            .Buttons(cStampa).Enabled = True
            .Buttons(cEsci).Enabled = True
         Case cExists
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
'            .Buttons(cModifica).Enabled = False
            .Buttons(cChiudi).Enabled = False
            .Buttons(cCerca).Enabled = True
            .Buttons(cStampa).Enabled = True
            .Buttons(cEsci).Enabled = True
         Case cNotExists
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
'            .Buttons(cModifica).Enabled = False
            .Buttons(cChiudi).Enabled = False
            .Buttons(cCerca).Enabled = True
            .Buttons(cStampa).Enabled = True
            .Buttons(cEsci).Enabled = True
      End Select
    End If
    
    
    
   End With
'Private Sub NuovoReclamo()
'End Sub
End Property
Public Property Get Status() As Integer
   Status = mStatus
End Property

Public Property Let Key(vValue As String)
    mKey = Trim$(vValue)
'    CercaID
'    CercaComponenti
End Property
Public Property Get Key() As String
    Key = mKey
End Property

Private Sub Form_Resize()
   Dim b As Button, l As Long, idx As Integer
   If Me.WindowState <> vbMinimized Then
      l = shpRisposta.Left * 2 + shpRisposta.Width + Me.Width - Me.ScaleWidth
      If Me.Width < l Then
         Me.Width = l
      End If
      l = 0
      Me.Refresh
      For Each b In tbCommand.Buttons
         If b.Caption <> "phRight" Then
            l = l + b.Width
         Else
            idx = b.Index
         End If
      Next
      Set b = tbCommand.Buttons(idx)
      b.Width = Me.ScaleWidth - l
      Set b = Nothing
   End If
End Sub

Private Sub tbCommand_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim iRes As Integer, rsTmp As Recordset, sqlc As String
    Select Case Button.Key
        Case "X"
            Unload Me
        Case "N"
'            fNuovoRecord = True
'            CercaIDN
'            Me.Status = cNew
        Case "S"
'            Controllo
'            If fSalva Then
'                SalvaMP
'                Me.Status = cExists
'            End If
            
        Case "A"
'            If fNuovoRecord Then
'                CancellaCampi
'                Me.Status = cNotExists
'            Else
'                CercaComponenti
'                Me.Status = cExists
'            End If
        Case "M"
'            fNuovoRecord = False
'            Campi True
'            Me.Status = cMod
        Case "C"
'            iRes = MsgBox("Vuoi chiudere questa produzione?", _
'            vbYesNo + vbDefaultButton2 + vbQuestion)
'            If iRes = vbYes Then
'                sqlc = "Update Formule set Caricata = 'P' " & _
'                    "Where SeId = " & txtSeId.Text & " and " & _
'                    "Caricata = 'S'"
'                db.Execute (sqlc)
'                sqlc = "Update SE set Qta = '" & txtSFT.Text & "' " & _
'                        "where SeId = " & txtSeId.Text
'                db.Execute (sqlc)
'
'                sStato = "C"
'                Me.Status = cExists
'                frmTV.ChangedTV = True
'            End If
        Case "P"
 '           fNuovoRecord = False
 '           Me.PrintForm
 '           Me.Status = cExists
        Case "G"
'            CercaIdC
        Case "T"
'            iRes = MsgBox("Vuoi trasferire questa produzione?", _
'            vbYesNo + vbDefaultButton2 + vbQuestion)
'            If iRes = vbYes Then
'                sqlc = "Update Formule set Caricata = 'T' " & _
'                    "Where SeId = " & txtSeId.Text & " and " & _
'                    "Caricata = 'F'"
'                db.Execute (sqlc)
'                Me.Status = cExists
'                frmTV.ChangedTV = True
'            End If
'            Load frmExportMFG
'            frmExportMFG.Key = txtSeId.Text
'            frmExportMFG.Show
            
    End Select
End Sub


