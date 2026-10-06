VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSpostamento 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Spostamento"
   ClientHeight    =   7035
   ClientLeft      =   3075
   ClientTop       =   3540
   ClientWidth     =   9255
   Icon            =   "frmReclamiVenditori.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   9255
   Begin VB.CommandButton cmdTrova 
      Caption         =   "&Trova"
      Height          =   375
      Left            =   4080
      TabIndex        =   21
      Top             =   2040
      Width           =   1815
   End
   Begin VB.ComboBox cmbLunghezza 
      Height          =   315
      Left            =   7680
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cmbAltezza 
      Height          =   315
      Left            =   6240
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cmbFinitura 
      Height          =   315
      Left            =   4800
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cmbSpessore 
      Height          =   315
      Left            =   3360
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cmbArticolo 
      Height          =   315
      Left            =   240
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   1440
      Width           =   2535
   End
   Begin VB.ListBox lstODL 
      Height          =   2985
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   1335
   End
   Begin MSComctlLib.Toolbar tbCommand 
      Align           =   2  'Align Bottom
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   6405
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   1111
      ButtonWidth     =   1164
      ButtonHeight    =   953
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ilToolBarV"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuovo"
            Key             =   "N"
            Object.ToolTipText     =   "Nuovo Reclamo"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Preleva"
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
            Object.Visible         =   0   'False
            Caption         =   "Modifica"
            Key             =   "M"
            Object.ToolTipText     =   "Modifica il Reclamo Corrente"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Chiudi"
            Key             =   "C"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "phRight"
            Style           =   4
            Object.Width           =   500
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Esci"
            Key             =   "X"
            Object.ToolTipText     =   "Esci"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilToolBarV 
      Left            =   8160
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReclamiVenditori.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReclamiVenditori.frx":0E26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReclamiVenditori.frx":136A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReclamiVenditori.frx":18AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReclamiVenditori.frx":1DF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReclamiVenditori.frx":2336
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape2 
      Height          =   1455
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   9015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qta a scaffale"
      Height          =   195
      Left            =   360
      TabIndex        =   34
      Top             =   2640
      Width           =   990
   End
   Begin VB.Label lId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Id"
      Height          =   195
      Left            =   3120
      TabIndex        =   33
      Top             =   3840
      Width           =   135
   End
   Begin VB.Label lIdP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pallette"
      Height          =   195
      Left            =   2280
      TabIndex        =   32
      Top             =   4800
      Width           =   525
   End
   Begin VB.Label lColonna 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Colonna"
      Height          =   195
      Left            =   5640
      TabIndex        =   31
      Top             =   4800
      Width           =   585
   End
   Begin VB.Label lRiga 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Riga"
      Height          =   195
      Left            =   4800
      TabIndex        =   30
      Top             =   4800
      Width           =   330
   End
   Begin VB.Label lSc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scaffale"
      Height          =   195
      Left            =   3480
      TabIndex        =   29
      Top             =   4800
      Width           =   585
   End
   Begin VB.Label lCodArt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CodArt"
      Height          =   195
      Left            =   6840
      TabIndex        =   28
      Top             =   3840
      Width           =   480
   End
   Begin VB.Label lblId 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2760
      TabIndex        =   27
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label lblIdP 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      TabIndex        =   26
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label lblColonna 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5520
      TabIndex        =   25
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label lblRiga 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4440
      TabIndex        =   24
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label lblScaffale 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3360
      TabIndex        =   23
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label lblCodArt 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6360
      TabIndex        =   22
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lunghezza"
      Height          =   195
      Left            =   7920
      TabIndex        =   20
      Top             =   1200
      Width           =   780
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Altezza"
      Height          =   195
      Left            =   6600
      TabIndex        =   19
      Top             =   1200
      Width           =   510
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Finitura"
      Height          =   195
      Left            =   5160
      TabIndex        =   18
      Top             =   1200
      Width           =   510
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spessore"
      Height          =   195
      Left            =   3600
      TabIndex        =   17
      Top             =   1200
      Width           =   660
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Articolo"
      Height          =   195
      Left            =   1200
      TabIndex        =   16
      Top             =   1200
      Width           =   525
   End
   Begin VB.Shape Shape1 
      Height          =   1695
      Left            =   1680
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   7095
   End
   Begin VB.Label lQta 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qta"
      Height          =   195
      Left            =   8160
      TabIndex        =   10
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label lProdotto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prodotto"
      Height          =   195
      Left            =   4680
      TabIndex        =   9
      Top             =   3840
      Width           =   600
   End
   Begin VB.Label lCliente 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente "
      Height          =   195
      Left            =   7320
      TabIndex        =   8
      Top             =   4560
      Width           =   525
   End
   Begin VB.Label lblQta 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   7920
      TabIndex        =   7
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label lblProdotto 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label lblCliente 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   6480
      TabIndex        =   5
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label lblODL 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label lOdl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ODL"
      Height          =   195
      Left            =   2040
      TabIndex        =   2
      Top             =   3840
      Width           =   330
   End
   Begin VB.Image iLogo 
      Height          =   450
      Left            =   0
      Picture         =   "frmReclamiVenditori.frx":287A
      Top             =   0
      Width           =   450
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Spostamento materiali"
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
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   4845
   End
End
Attribute VB_Name = "frmSpostamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const LB_FINDSTRING = &H18F
Private sSearch As String
Private mKey As String, rsReclamoVenditore As Recordset
Private fNuovoRecord As Boolean, sStato As String
Private Enum CMDS
   cNuovo = 1
   cSalva = 3
   cAnnulla = 5
   cModifica = 7
   cChiudi = 9
   cEsci = 11
End Enum
Dim rsTemp As Recordset

Private Function CheckCampi() As Boolean
   Dim fExit As Boolean
End Function





Private Sub SalvaRec()
   Dim sCod As Variant
   If fNuovoRecord Then
      rsReclamoVenditore.AddNew
   Else
      rsReclamoVenditore.Edit
   End If
   
'   rsReclamoVenditore("ODL") = lblODL.Caption
'   rsReclamoVenditore("DataReclamo") = mskDataReclamo.Text
'   rsReclamoVenditore("QtaReclamata") = txtQtaReclamata.Text
'   rsReclamoVenditore("CausaReclamo") = txtCausaReclamo.Text
'   rsReclamoVenditore("Venditore") = txtVenditore.Text
'   rsReclamoVenditore("NumeroReclamo") = txtNumeroReclamo.Text
'   rsReclamoVenditore("Stato") = "N"
'   rsReclamoVenditore.Update
'   rsReclamoVenditore.MoveFirst
'   CercaReclamoVenditore lblODL.Caption
'     disabilito i campi
'   Campi False
   fNuovoRecord = False
End Sub

Private Sub cmbAltezza_Click()
ListLunghezza
cmbLunghezza.Visible = True
'cmbAltezza.Locked = True
End Sub

Private Sub cmbArticolo_Click()
If Not fNuovoRecord Then
    ListSpessore
    cmbSpessore.Visible = True
'    cmbSpessore.Locked = True
End If
End Sub

Private Sub cmbFinitura_Click()
ListAltezza
cmbAltezza.Visible = True
'cmbAltezza.Locked = True
End Sub

Private Sub cmbSpessore_Click()
ListFinitura
cmbFinitura.Visible = True
'cmbFinitura.Locked = True
cmdTrova.Visible = True
End Sub

Private Sub cmdTrova_Click()
TrovoQta
End Sub

Private Sub Form_Load()
   Dim btn As Button
   Dim sqlc As String, rsODL As Recordset
    NascondiLbl False
'    ListArticolo
'    ListSpessore
'    ListFinitura
'    ListAltezza
'    ListLunghezza

    Combo False
   Me.MousePointer = vbArrowHourglass
   Me.Show
   Me.Refresh
   ShowWait "Attendere !!" & vbCrLf & vbCrLf & "Caricamento in corso"
'   sqlc = "Select *   from [InvioVenditore] order by 1"
   

'   Set rsODL = db.OpenRecordset(sqlc, dbOpenSnapshot)
'   Do While Not rsODL.EOF
'      lstODL.AddItem rsODL(0)
'      rsODL.MoveNext
'   Loop
'   rsODL.Close
'   Set rsODL = Nothing
'   lstODL.ListIndex = -1
'   lstODL.SetFocus
   
'   Campi False
    'Nascondo tutti i pulsanti
'    cmdNew.Visible = False
'    cmdSalva.Visible = False
'    cmdAnnulla.Visible = False
'    cmdModifica.Visible = False
   For Each btn In tbCommand.Buttons
     If btn.Key <> "X" Then
        If btn.Key = "N" Then
            btn.Enabled = True
        Else
            btn.Enabled = False
        End If
        
'        btn.Enabled = False
         
     End If
   Next
'   tbCommand.Buttons(cNuovo).Enabled = False
   'tbCommand.Buttons(cNuovo).Enabled = True
   Me.MousePointer = vbNormal
    cmdTrova.Visible = False
    
   HideWait

End Sub

Private Sub Form_Resize()
   Dim b As Button, l As Long, idx As Integer
   If Me.WindowState <> vbMinimized Then
'      l = shpRisposta.Left + shpRisposta.Width + Me.Width - Me.ScaleWidth
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

Private Sub Form_Unload(Cancel As Integer)
'    Load frmVenditori
'    frmVenditori.Show
'   rsReclamoVenditore.Close
   Set rsReclamoVenditore = Nothing
End Sub

Private Sub NuovoReclamo()
   tbCommand.Buttons(cNuovo).Enabled = True
   tbCommand.Buttons(cSalva).Enabled = False
   tbCommand.Buttons(cAnnulla).Enabled = False
   tbCommand.Buttons(cModifica).Enabled = False
End Sub
Private Sub CancellaCampi()
'   mskDataReclamo.Text = "__/__/__"
'   txtQtaReclamata.Text = ""
'   txtCausaReclamo.Text = ""
'   txtVenditore.Text = ""
'   txtNumeroReclamo.Text = ""
End Sub


Private Sub lstODL_DblClick()
Dim sTmp As String, sqlc As String
Dim iPos As Integer, sL As String
iPos = 0

sTmp = lstODL.List(lstODL.ListIndex)
iPos = InStr(sTmp, "-")
'sL = Len(sTmp)
sTmp = Mid$(sTmp, iPos + 1, 9999)

   
   
   
   CancellaCampi
   lblIdP.Caption = sTmp
'   sqlc = "Select *   from [InvioVenditore] where ODL='" & sTmp & "'"
   
       sqlc = "SELECT DISTINCT Pallette.Qta, Pallette.IdP, ODL.Id, Odl.Odl, " & _
                    "ODL.CodiceArticolo, ODL.Articolo, ODL.Spessore, " & _
                    "ODL.Altezza, ODL.Lunghezza, ODL.Finitura, ODL.Cliente, " & _
                    "Legami.Stato, Legami.DataE, Legami.IdS, Scaffali.Scaffale, Scaffali.Riga, Scaffali.Colonna " & _
                "FROM (ODL INNER JOIN (Legami INNER JOIN Pallette ON " & _
                    "Legami.IdP = Pallette.IdP) ON ODL.ID = Pallette.Id) " & _
                    "INNER JOIN Scaffali ON Legami.IdS = Scaffali.IdS " & _
            "WHERE Legami.Stato='E' AND Articolo='" & cmbArticolo.Text & "'" & _
                " AND Spessore=" & cmbSpessore.Text & " AND Altezza=" & _
                cmbAltezza.Text & "AND Finitura='" & cmbFinitura.Text & "'" & _
                "AND Lunghezza=" & cmbLunghezza.Text & "AND pallette.IdP=" & sTmp

   
   
   
   Set rsTemp = db.OpenRecordset(sqlc, dbOpenSnapshot)
   If Not rsTemp.EOF Then
        If Not IsNull(rsTemp("Odl")) Then
            lblODL.Caption = rsTemp("Odl")
        Else
            lblODL.Caption = ""
        End If
        If Not IsNull(rsTemp("Cliente")) Then
            lblCliente.Caption = rsTemp("Cliente")
        Else
            lblCliente.Caption = ""
        
        End If
        If Not IsNull(rsTemp("Articolo")) Then
            lblProdotto.Caption = rsTemp("Articolo")
        Else
            lblProdotto.Caption = ""
        End If
        If Not IsNull(rsTemp("Qta")) Then
            lblQta.Caption = rsTemp("Qta")
        Else
            lblQta.Caption = ""
        End If
        If Not IsNull(rsTemp("CodiceArticolo")) Then
            lblCodArt.Caption = rsTemp("CodiceArticolo")
        Else
            lblCodArt.Caption = ""
        End If
        If Not IsNull(rsTemp("Scaffale")) Then
            lblScaffale.Caption = rsTemp("Scaffale")
        Else
            lblScaffale.Caption = ""
        End If
        If Not IsNull(rsTemp("Riga")) Then
            lblRiga.Caption = rsTemp("Riga")
        Else
            lblRiga.Caption = ""
        End If
        If Not IsNull(rsTemp("Colonna")) Then
            lblColonna.Caption = rsTemp("Colonna")
        Else
            lblColonna.Caption = ""
        End If
        If Not IsNull(rsTemp("Id")) Then
            lblId.Caption = rsTemp("Id")
        Else
            lblId.Caption = ""
        End If
   End If
   
   NascondiLbl True

'Controllare che significa


'    If CercaReclamoVenditore(sTmp) Then
'       FoundODL
'    Else
'       NuovoReclamo
'    End If
End Sub

Private Sub lstODL_KeyPress(KeyAscii As Integer)
'
'  Sarebbe anche meglio decidere quali tasti vanno ad accrescere la "Stringa"
'     (ad es. solo numeri ed il . od anche altri ??)
'
'  Inoltre potrebbe essere utile (per l'operatore) associare al tasto INVIO
'      l 'equivalente del DblClick sull'elemento
'
   If KeyAscii = vbKeyReturn Then
      lstODL_DblClick
      Exit Sub
   End If
   If KeyAscii = vbKeyBack Then
       If Len(sSearch) >= 1 Then
           sSearch = Left(sSearch, Len(sSearch) - 1)
       End If
   Else
      sSearch = sSearch & Chr$(KeyAscii)
   End If
   lstODL.ListIndex = SendMessage(lstODL.hwnd, LB_FINDSTRING, -1, ByVal CStr(sSearch))
   KeyAscii = 0
End Sub


Private Sub tbCommand_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim iRes As Integer, rsTmp As Recordset, sqlc As String
   Select Case Button.Key
      Case "X"
         Unload Me
      Case "N"
         fNuovoRecord = True
'         Campi True
         CancellaCampi
         ListArticolo
         cmbArticolo.Visible = True
         With tbCommand
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = True
            .Buttons(cAnnulla).Enabled = True
            .Buttons(cModifica).Enabled = False
'            .Buttons(cChiudi).Enabled = False
            .Buttons(cEsci).Enabled = False
            lstODL.Enabled = False
         End With
            fNuovoRecord = False
      Case "S" 'Prelievo
'         If CheckCampi Then
'            SalvaRec
         iRes = MsgBox("Vuoi prelevare la pallette corrente", _
               vbYesNo + vbDefaultButton2 + vbQuestion)
         If iRes = vbYes Then
            Prelievo
         End If
            With tbCommand
               .Buttons(cNuovo).Enabled = False
               .Buttons(cSalva).Enabled = False
               .Buttons(cAnnulla).Enabled = False
               .Buttons(cModifica).Enabled = True
   '            .Buttons(cChiudi).Enabled = (sStato = "V")
               .Buttons(cEsci).Enabled = True
               lstODL.Enabled = True
            End With
'         End If
      Case "A"
'         If fNuovoRecord Then
            CancellaCampi
            NascondiLbl False
            cmdTrova.Visible = False
            Combo False
'         Else
'            FoundODL
'         End If
'            NascondiLbl True
'         Campi False
            lstODL.Clear
         With tbCommand
            .Buttons(cNuovo).Enabled = True   '  Dipende se arrivo da Mod o Nuovo
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = Not fNuovoRecord '  Dipende se arrivo da Mod o Nuovo
'            .Buttons(cChiudi).Enabled = (sStato = "V")
            .Buttons(cEsci).Enabled = True
            lstODL.Enabled = True
         End With
      Case "M"
         fNuovoRecord = False
'         Campi True
         With tbCommand
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = True
            .Buttons(cAnnulla).Enabled = True
            .Buttons(cModifica).Enabled = False
'            .Buttons(cChiudi).Enabled = False
            .Buttons(cEsci).Enabled = False
            lstODL.Enabled = False
         End With
'      Case "C"
'         iRes = MsgBox("Vuoi chiudere il corrente reclamo", _
'               vbYesNo + vbDefaultButton2 + vbQuestion)
'         If iRes = vbYes Then
'            sqlc = "Update ReclamoVenditore set Stato = 'N' " & _
'                     "Where ODL = '" & Me.Key
'            db.Execute (sqlc)
'            tbCommand.Buttons(cChiudi).Enabled = False
'         End If
   End Select
End Sub

Private Sub ListArticolo()
Dim rsArticolo As Recordset, sqlc As String, sMsg As String
Dim i As Integer
'ScLiberi

    sqlc = "SELECT DISTINCT ODL.Articolo, Legami.Stato " & _
            "FROM ODL INNER JOIN (Legami INNER JOIN " & _
                "Pallette ON Legami.IdP = Pallette.IdP) " & _
                "ON ODL.ID = Pallette.Id " & _
            "WHERE Legami.Stato='E'"

            
    Set rsArticolo = db.OpenRecordset(sqlc, dbOpenSnapshot)
    
    cmbArticolo.Clear
    
    Do While Not rsArticolo.EOF
        cmbArticolo.AddItem rsArticolo("Articolo")
        rsArticolo.MoveNext
    Loop
    
    rsArticolo.Close
    Set rsArticolo = Nothing
    
    cmbArticolo.ListIndex = 0
    
End Sub
Private Sub ListSpessore()
Dim rsSpessore As Recordset, sqlc As String, sMsg As String
Dim i As Integer
'ScLiberi
    
    sqlc = "SELECT distinct ODL.CodiceArticolo, ODL.Articolo, " & _
                "ODL.Spessore, ODL.Altezza, ODL.Lunghezza, " & _
                "ODL.Finitura, Legami.Stato " & _
            "FROM ODL INNER JOIN (Legami INNER " & _
                "JOIN Pallette ON Legami.IdP = Pallette.IdP) " & _
                "ON ODL.ID = Pallette.Id " & _
            "WHERE Legami.Stato='E' AND Articolo='" & cmbArticolo.Text & "'"
    

    Set rsSpessore = db.OpenRecordset(sqlc, dbOpenSnapshot)
    
    cmbSpessore.Clear
    
    Do While Not rsSpessore.EOF
        cmbSpessore.AddItem rsSpessore("Spessore")
        rsSpessore.MoveNext
    Loop
    
    rsSpessore.Close
    Set rsSpessore = Nothing
    
    cmbSpessore.ListIndex = 0

End Sub
Private Sub ListFinitura()
Dim rsFinitura As Recordset, sqlc As String, sMsg As String
Dim i As Integer
'ScLiberi
    sqlc = "SELECT distinct ODL.CodiceArticolo, ODL.Articolo, " & _
                "ODL.Spessore, ODL.Altezza, ODL.Lunghezza, " & _
                "ODL.Finitura, Legami.Stato " & _
            "FROM ODL INNER JOIN (Legami INNER " & _
                "JOIN Pallette ON Legami.IdP = Pallette.IdP) " & _
                "ON ODL.ID = Pallette.Id " & _
            "WHERE Legami.Stato='E' AND Articolo='" & cmbArticolo.Text & "'" & _
                " AND Spessore=" & cmbSpessore.Text
    
            

    Set rsFinitura = db.OpenRecordset(sqlc, dbOpenSnapshot)
    
    cmbFinitura.Clear
    
    Do While Not rsFinitura.EOF
        If Not IsNull(rsFinitura("Finitura")) Then
            cmbFinitura.AddItem rsFinitura("Finitura")
        Else
            cmbFinitura.AddItem "0"
        End If
        
        rsFinitura.MoveNext
        
    Loop
    
    rsFinitura.Close
    Set rsFinitura = Nothing
    
    cmbFinitura.ListIndex = 0

End Sub
Private Sub ListAltezza()
Dim rsAltezza As Recordset, sqlc As String, sMsg As String
Dim i As Integer
'ScLiberi
    sqlc = "SELECT distinct ODL.CodiceArticolo, ODL.Articolo, " & _
                "ODL.Spessore, ODL.Altezza, ODL.Lunghezza, " & _
                "ODL.Finitura, Legami.Stato " & _
            "FROM ODL INNER JOIN (Legami INNER " & _
                "JOIN Pallette ON Legami.IdP = Pallette.IdP) " & _
                "ON ODL.ID = Pallette.Id " & _
            "WHERE Legami.Stato='E' AND Articolo='" & cmbArticolo.Text & "'" & _
                " AND Spessore=" & cmbSpessore.Text & " AND finitura='" & _
                cmbFinitura.Text & "'"
    
            

    Set rsAltezza = db.OpenRecordset(sqlc, dbOpenSnapshot)
    
    cmbAltezza.Clear
    
    Do While Not rsAltezza.EOF
        If Not IsNull(rsAltezza("Altezza")) Then
            cmbAltezza.AddItem rsAltezza("Altezza")
        Else
            cmbAltezza.AddItem "0"
        End If
        
        rsAltezza.MoveNext
    Loop
    
    rsAltezza.Close
    Set rsAltezza = Nothing
    
    cmbAltezza.ListIndex = 0

End Sub
Private Sub ListLunghezza()
Dim rsLunghezza As Recordset, sqlc As String, sMsg As String
Dim i As Integer
'ScLiberi
    sqlc = "SELECT distinct ODL.CodiceArticolo, ODL.Articolo, " & _
                "ODL.Spessore, ODL.Altezza, ODL.Lunghezza, " & _
                "ODL.Finitura, Legami.Stato " & _
            "FROM ODL INNER JOIN (Legami INNER " & _
                "JOIN Pallette ON Legami.IdP = Pallette.IdP) " & _
                "ON ODL.ID = Pallette.Id " & _
            "WHERE Legami.Stato='E' AND Articolo='" & cmbArticolo.Text & "'" & _
                " AND Spessore=" & cmbSpessore.Text & " AND Altezza=" & _
                cmbAltezza.Text & "AND Finitura='" & cmbFinitura.Text & "'"
    
            

    Set rsLunghezza = db.OpenRecordset(sqlc, dbOpenSnapshot)
    
    cmbLunghezza.Clear
    
    Do While Not rsLunghezza.EOF
        cmbLunghezza.AddItem rsLunghezza("Lunghezza")
        rsLunghezza.MoveNext
    Loop
    
    rsLunghezza.Close
    Set rsLunghezza = Nothing
    
    cmbLunghezza.ListIndex = 0

End Sub

Private Sub ScLiberi()
Dim sqlc As String, rsT As Recordset, sT As String
Dim fP As Boolean
sT = ""
fP = False
sT = "Where not (Ids="
sqlc = "Select IdS " & _
        "From Legami " & _
        "Where Stato='E'"
        
Set rsT = db.OpenRecordset(sqlc, dbOpenSnapshot)
    If Not rsT.EOF Then
        Do While Not rsT.EOF
            If Not fP Then
                sT = sT & rsT(0)
                fP = True
            Else
                sT = sT & " or IdS=" & rsT(0)
            End If
            
            rsT.MoveNext
        Loop
        
    End If

sT = sT & ")"
'sSc = sT

rsT.Close
Set rsT = Nothing
    

End Sub

Private Sub Combo(Flag As Boolean)

cmbArticolo.Visible = Flag
cmbSpessore.Visible = Flag
cmbFinitura.Visible = Flag
cmbAltezza.Visible = Flag
cmbLunghezza.Visible = Flag



End Sub
Private Sub TrovoQta()
Dim rsQta As Recordset, sqlc As String, sMsg As String
Dim i As Integer
'ScLiberi
    sqlc = "SELECT distinct Pallette.Qta, Pallette.IdP, ODL.CodiceArticolo, ODL.Articolo, " & _
                "ODL.Spessore, ODL.Altezza, ODL.Lunghezza, " & _
                "ODL.Finitura, Legami.Stato " & _
            "FROM ODL INNER JOIN (Legami INNER " & _
                "JOIN Pallette ON Legami.IdP = Pallette.IdP) " & _
                "ON ODL.ID = Pallette.Id " & _
            "WHERE Legami.Stato='E' AND Articolo='" & cmbArticolo.Text & "'" & _
                " AND Spessore=" & cmbSpessore.Text & " AND Altezza=" & _
                cmbAltezza.Text & "AND Finitura='" & cmbFinitura.Text & "'" & _
                "AND Lunghezza=" & cmbLunghezza.Text
    
            

    Set rsQta = db.OpenRecordset(sqlc, dbOpenSnapshot)
    
    Do While Not rsQta.EOF
        lstODL.AddItem rsQta(0) & "-" & rsQta(1)
        rsQta.MoveNext
    Loop
    rsQta.Close
    Set rsQta = Nothing
    lstODL.ListIndex = -1
    lstODL.Enabled = True
    lstODL.SetFocus
    
    


End Sub
Private Sub NascondiLbl(Flag As Boolean)

    lblODL.Visible = Flag
    lblCliente.Visible = Flag
    lblProdotto.Visible = Flag
    lblQta.Visible = Flag
    lblCodArt.Visible = Flag
    lblScaffale.Visible = Flag
    lblRiga.Visible = Flag
    lblColonna.Visible = Flag
    lblId.Visible = Flag
    lblIdP.Visible = Flag
    lOdl.Visible = Flag
    lCliente.Visible = Flag
    lProdotto.Visible = Flag
    lQta.Visible = Flag
    lCodArt.Visible = Flag
    lSc.Visible = Flag
    lRiga.Visible = Flag
    lColonna.Visible = Flag
    lIdP.Visible = Flag
    lId.Visible = Flag
    
    
    
End Sub
Private Sub Prelievo()
Dim sT As String, rsT As Recordset, sqlc As String
Dim sL As String, sD As String

sL = rsTemp("IdS")
sD = Format(rsTemp("DataE"), "mm/dd/yyyy hh.mm.ss")
sT = rsTemp("IdP")

   
       sqlc = "SELECT IdP, IdS, Stato, DataE, DataU " & _
                "From Legami " & _
                "WHERE IdP=" & sT & " AND IdS=" & sL & _
                " AND DataE=#" & sD & "#"

   
   
   
    Set rsTemp = db.OpenRecordset(sqlc, dbOpenDynaset)
    If Not rsTemp.EOF Then
        rsTemp.Edit
            rsTemp("Stato") = "U"
            rsTemp("DataU") = Now
        rsTemp.Update
    End If
      

End Sub
