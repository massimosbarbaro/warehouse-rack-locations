VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Main"
   ClientHeight    =   5175
   ClientLeft      =   5040
   ClientTop       =   3405
   ClientWidth     =   4950
   ForeColor       =   &H00C0C0C0&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   4950
   Begin VB.CommandButton cmdScarico 
      Caption         =   "&Scarico"
      Height          =   615
      Left            =   1800
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdEsci 
      Caption         =   "Esci"
      Height          =   615
      Left            =   4080
      Picture         =   "frmMain.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmdStatistiche 
      Caption         =   "&Aggiorna"
      Height          =   615
      Left            =   1800
      TabIndex        =   2
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdSegnalazioni 
      Caption         =   "&Import"
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdReclami 
      Caption         =   "&Carico"
      Height          =   615
      Left            =   1800
      MaskColor       =   &H00FF0000&
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Image iLogo 
      Height          =   450
      Left            =   120
      Picture         =   "frmMain.frx":0E14
      Top             =   0
      Width           =   450
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Scaffalature"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   540
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   2280
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   3015
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Menu mnuReclami 
      Caption         =   "Reclami"
      Visible         =   0   'False
      Begin VB.Menu mnuR 
         Caption         =   "Inserimento"
         Index           =   0
      End
      Begin VB.Menu mnuR 
         Caption         =   "-"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuR 
         Caption         =   "Spostamento"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuR 
         Caption         =   "-"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuR 
         Caption         =   "Posta"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuR 
         Caption         =   "Visualizza MKG"
         Index           =   5
      End
      Begin VB.Menu mnuR 
         Caption         =   "aTC/mKG"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuR 
         Caption         =   "Marketing"
         Index           =   7
      End
   End
   Begin VB.Menu mnuStatistiche 
      Caption         =   "Statistiche"
      Visible         =   0   'False
      Begin VB.Menu mnuS 
         Caption         =   "Importa ODL"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEsci_Click()
   Unload Me
End Sub

Private Sub cmdReclami_Click()
'    Load frmVenditori
'    frmVenditori.Show
   Screen.MousePointer = vbArrowHourglass
    
    Load frmID
    frmID.Key = "ATC"
    frmID.Show
   Screen.MousePointer = vbNormal


End Sub

Private Sub cmdReclami_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Select Case sFlag
'      Case "ATC"
'         mnuR(0).Visible = True
'         mnuR(2).Visible = True
'         mnuR(5).Visible = False
'         mnuR(7).Visible = False
'      Case "MKG"
'         mnuR(0).Visible = False
'         mnuR(5).Visible = True
'         mnuR(7).Visible = False
'      Case "ATC/MKG"
'         mnuR(0).Visible = True
'         mnuR(5).Visible = True
'         mnuR(7).Visible = False
'      Case "VIEW"
'         mnuR(0).Visible = False
'         mnuR(5).Visible = False
'         mnuR(7).Visible = True
'      Case Else
'         MsgBox "Questo messaggio non esiste !!"
'
'   End Select
'   PopupMenu mnuReclami, vbPopupMenuCenterAlign
End Sub


Private Sub cmdScarico_Click()
   Screen.MousePointer = vbArrowHourglass
        
        Load frmSpostamento
  '      frmSpostamento.Key = "ATC"
        frmSpostamento.Show
   Screen.MousePointer = vbNormal

End Sub

Private Sub cmdSegnalazioni_Click()
'   Load frmImportMFG
'   frmImportMFG.Show vbModal

End Sub

Private Sub cmdStatistiche_Click()
   Screen.MousePointer = vbArrowHourglass
        Load frmExportMFG
        frmExportMFG.Key = "Aggiorna"
        frmExportMFG.Show
   
   Screen.MousePointer = vbNormal
End Sub

Private Sub cmdStatistiche_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  PopupMenu mnuStatistiche, vbPopupMenuCenterAlign
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim i As Integer, res As Integer, nRs As Integer
    res = MsgBox("Esco dall'Applicazione ?", vbQuestion + vbYesNo + vbDefaultButton1, _
        "scaffali - Uscita")
    If res = vbYes Then
'
'   Check su altre forms caricate
'
        For i = Forms.Count - 1 To 0 Step -1
            If Forms(i).Name <> Me.Name Then
    '            Set f = Forms(i)
                Unload Forms(i)
            End If
        Next i
    '
    '   Controllo rs aperti !!!  (Ben Fatto)
    '
        nRs = db.Recordsets.Count
        For i = 1 To nRs
            db.Recordsets(0).Close
        Next i
    '
    '   Chiusura DB
    '
        If fDBOk Then
            db.Close
        End If
        Set db = Nothing
    Else
        Cancel = True
    End If
End Sub

Private Sub mnuR_Click(Index As Integer)
   Select Case Index
      Case 0
        Load frmID
        frmID.Key = "ATC"
        frmID.Show
       
'         Load frmTV
'         frmTV.Param = "ATC"
'         frmTV.Show
      Case 2
        Load frmSpostamento
  '      frmSpostamento.Key = "ATC"
        frmSpostamento.Show

      Case 4
'         Load frmPop3
'         frmPop3.Show vbModal
      Case 5
'         Load frmTV
'         frmTV.Param = "MKG"
'         frmTV.Show
'      Case 6
'         Load frmTV
'         frmTV.Param = "ATC/MKG"
'         frmTV.Show
      Case 7
'         Load frmTV
'         frmTV.Param = "VIEW"
'         frmTV.Show
   
   End Select
End Sub

Private Sub mnuS_Click(Index As Integer)
Select Case Index
    Case 0
'        Load frmGrafici
'        frmGrafici.Show
        Load frmExportMFG
        frmExportMFG.Key = "Aggiorna"
        frmExportMFG.Show

    Case 1
'        Load frmIQ
'        frmIQ.Show
End Select
End Sub
