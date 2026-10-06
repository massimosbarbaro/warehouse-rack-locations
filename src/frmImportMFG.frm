VERSION 5.00
Begin VB.Form frmExportMFG 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Export MFG"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   Icon            =   "frmImportMFG.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCLP 
      Caption         =   "Import"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton cmdEsci 
      Cancel          =   -1  'True
      Caption         =   "Esci"
      Default         =   -1  'True
      Height          =   615
      Left            =   6840
      Picture         =   "frmImportMFG.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   615
   End
   Begin VB.TextBox txtFileName 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Text            =   "C:\Programmi\reclami\txt\resi282.txt"
      Top             =   1320
      Width           =   5415
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   415
      Left            =   1440
      TabIndex        =   3
      Top             =   3480
      Width           =   4575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   1695
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   5655
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Importazione dati da MFG"
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
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   5760
   End
   Begin VB.Image iLogo 
      Height          =   450
      Left            =   240
      Picture         =   "frmImportMFG.frx":0E14
      Top             =   120
      Width           =   450
   End
   Begin VB.Label lblLog 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   2175
      Left            =   960
      TabIndex        =   4
      Top             =   2880
      Width           =   5655
   End
End
Attribute VB_Name = "frmExportMFG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mKey As String, rsATC As Recordset
Private Declare Function OpenProcess Lib "kernel32" _
   (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
   ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" _
   (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" _
   (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" _
   (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Const INFINITE = &HFFFF
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const STILL_ACTIVE = 259
Private Const PROCESS_ALL_ACCESS = &H1F0FFF

'  Private ImpMfg(0 To 19, 1 To 2) As Variant
Private Const dbTempName As String = "$$Temp.mdb"
Private dbTemp As Database
Private ImpMfg(0 To 1) As String
Private ImpCLP(0 To 5) As String
Dim sFinale As String
Dim fClp As Boolean
Dim sqlc As String, rsTemp As Recordset, i As Integer
Private Function IsActive(hProg) As Long
   Dim hProc, RetVal As Long
   Const PROCESS_QUERY_INFORMATION = &H400
   Const STILL_ACTIVE = 259
   hProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, hProg)
   If hProc <> 0 Then
      GetExitCodeProcess hProc, RetVal
   End If
   IsActive = (RetVal = STILL_ACTIVE)
   CloseHandle hProc
End Function

Private Sub cmdEsci_Click()
   Unload Me
End Sub

Private Sub cmdImport_Click()
   Dim nFile As Integer, sFileName As String
   Dim sqlc As String, rsImport As Recordset, rsMFG As Recordset
   Dim sLine As String, i As Integer, sTmp As Variant
   Dim nFields As Integer, nRecs As Integer
   Dim sFields(0 To 19) As String
'   *************************************************************************
   Dim hProg, hProc, RetVal As Long
   Dim sCmd As String, sPrg As String, sFile As String, sLabel As String
   
'   Const PROCESS_ALL_ACCESS = 0
   Const PROCESS_ALL_ACCESS = &H1F0FFF
   cmdEsci.Enabled = False
   Me.Enabled = False
      
'  trasporto il file txt dal server sul pc
   sCmd = "ftp -s:" & App.Path & "\" & "wowomt.scr"
'   sCmd = "ftp -s:" & App.Path & "\" & "ftp.txt" ' originale
'   sCmd = "ftp -s:" & App.Path & "\" & "prova.txt"
  hProg = Shell(sCmd, vbNormalFocus)
  hProc = OpenProcess(PROCESS_ALL_ACCESS, False, hProg)
   If hProc <> 0 Then
      RetVal = WaitForSingleObject(hProc, INFINITE)
      CloseHandle hProc
   End If






' Parte lo script di reflection per creare il file Caricare la CIM
   sPrg = "c:\programmi\r1win\r1win.exe"
   sFile = App.Path & "\" & "CaricoCIM.rbs"
   sCmd = sPrg & " " & "/rbs" & " " & sFile
'   sCmd = "calc.exe"       'prova di funzionamento sincrono
   
   hProg = Shell(sCmd, vbNormalFocus)
   hProc = OpenProcess(PROCESS_ALL_ACCESS, False, hProg)
   If hProc <> 0 Then
      RetVal = WaitForSingleObject(hProc, INFINITE)
      CloseHandle hProc
   End If
'  MsgBox "Fine scarico da MFG"
 
 '   sCmd = App.Path & "\" & "teln.bat"
   

' Parte lo script di reflection per creare crearmi la
' for each per l'ID di mescola

    sPrg = "c:\programmi\r1win\r1win.exe"
    sFile = App.Path & "\" & "NoteFile.rbs"
    sCmd = sPrg & " " & "/rbs" & " " & sFile
    '   sCmd = "calc.exe"       'prova di funzionamento sincrono
    
    hProg = Shell(sCmd, vbNormalFocus)
    hProc = OpenProcess(PROCESS_ALL_ACCESS, False, hProg)
    If hProc <> 0 Then
        RetVal = WaitForSingleObject(hProc, INFINITE)
        CloseHandle hProc
    End If

'  trasporto il file con le note dal server sul pc
   sCmd = "ftp -s:" & App.Path & "\" & "note.scr"
'   sCmd = "ftp -s:" & App.Path & "\" & "ftp.scr" ' originale
'   sCmd = "ftp -s:" & App.Path & "\" & "ftp.txt" ' originale
'   sCmd = "ftp -s:" & App.Path & "\" & "prova.txt"
   hProg = Shell(sCmd, vbNormalFocus)
   hProc = OpenProcess(PROCESS_ALL_ACCESS, False, hProg)
   If hProc <> 0 Then
      RetVal = WaitForSingleObject(hProc, INFINITE)
      CloseHandle hProc
   End If
   
   Do While IsActive(hProg)
      DoEvents
   Loop
   hProc = OpenProcess(PROCESS_ALL_ACCESS, False, hProg)
   If hProc <> 0 Then
      RetVal = WaitForSingleObject(hProc, INFINITE)
      CloseHandle hProc
  End If
 ' MsgBox "Fine trasporto FTP"
     
  Me.Enabled = True
  Me.SetFocus
'   ******************************************************************************
   Screen.MousePointer = vbArrowHourglass
   
'   cmdImport.Enabled = False
   Me.Enabled = False
'
'  Apertura Db Temporaneo
'
   Set dbTemp = dao.OpenDatabase(App.Path & "\" & dbTempName)

   lblLog.Caption = ""
   lblLog.Visible = True
   With lblProgress
      .Caption = ""
      .Top = lblLog.Top + lblLog.Height - (.Height * 1.5)
      .Visible = True
   End With
   Me.Refresh

   nRecs = 0

   lblLog.Caption = "E' in corso " & vbCrLf & _
                  " la cancellazione della tabella temporanea"
   DoEvents

   sqlc = "Delete from [$$ImportMFG]"
   dbTemp.Execute (sqlc)

   sqlc = "Select * from [$$ImportMFG]"
   Set rsImport = dbTemp.OpenRecordset(sqlc, dbOpenDynaset)
   nFields = rsImport.Fields.Count

   lblLog.Caption = lblLog.Caption & vbCrLf & _
      "Stiamo caricando una nuova tabella "
   DoEvents

   sFileName = txtFileName.Text
   nFile = FreeFile
   Open sFileName For Input As nFile
   Do While Not EOF(nFile) '
      Line Input #nFile, sLine
      nRecs = nRecs + 1
      If (nRecs Mod 100) = 0 Then
         DoEvents
         lblProgress.Caption = "Caricamento rec. #" & Format$(nRecs, "####0")
      End If
  If Not IsNull(sLine) And sLine <> "" Then
      For i = 0 To nFields - 1
         sFields(i) = Trim$(Mid$(sLine, (i * 9 + 1), 8))
      Next i
      rsImport.AddNew
      For i = 0 To nFields - 1
         sTmp = Trim$(sFields(i))
         Select Case rsImport.Fields(ImpMfg(i)).Type
            Case dbText '
               If sTmp <> "" Then
                  rsImport(ImpMfg(i)) = Left$(sTmp, rsImport(ImpMfg(i)).Size)
               Else
                  rsImport(ImpMfg(i)) = Null
               End If
            Case dbDate
               If sTmp = "" Then
                  rsImport(ImpMfg(i)) = Null
               Else
                  rsImport(ImpMfg(i)) = sTmp
               End If
            Case Else
               If sTmp = "" Then
                  rsImport(ImpMfg(i)) = Null
               Else
                  rsImport(ImpMfg(i)) = Val(sTmp)
               End If
         End Select
      Next i
      rsImport.Update
    End If
    
   Loop

   Close nFile
   rsImport.Close
   Set rsImport = Nothing
   lblProgress.Visible = False

   lblLog.Caption = vbCrLf & _
         "Fine caricamento tabella temporanea: " & _
         Str$(nRecs) & " Records"
   DoEvents
   sLabel = "Caricati " & Str$(nRecs) & " Records"

'  Caricamento tabella MFG con i soli ODL Distinti e non nulli
   lblLog.Caption = ""
   With lblLog
'      .Caption = lblLog.Caption & vbCrLf & _
         "Cancellazione  vecchi dati "
      .ForeColor = &H80000007
      .FontSize = 10
      .Alignment = 0
   End With
   DoEvents
'Comincio il confronto dei dati
' rsImport è per noi il rsNew

   
   sqlc = "SELECT [$$ImportMFG].ID, [$$ImportMFG].Note " & _
            "From [$$ImportMFG] " & _
            "WHERE ((([$$ImportMFG].Note) Like 'seid*'))"

   Set rsImport = dbTemp.OpenRecordset(sqlc, dbOpenSnapshot)
        nRecs = 0
        With lblProgress
            .Caption = ""
            .Top = lblLog.Top + (.Height * 0.5)
            .Visible = True
        End With
        nFields = rsImport.Fields.Count
    
    Do While Not rsImport.EOF
        sqlc = "SELECT * FROM [Note] " & _
                "where Id='" & rsImport("Id") & "'"
        Set rsMFG = db.OpenRecordset(sqlc, dbOpenDynaset)
        
        On Error Resume Next
        
      If Not rsMFG.EOF Then
         rsMFG.Edit
      Else
         rsMFG.AddNew
         rsMFG("Id") = rsImport("Id")
      End If
' Questa parte serve solo a caricare la barra ---------------------
        nRecs = nRecs + 1
        If (nRecs Mod 100) = 0 Then
           lblProgress.Caption = "Caricamento rec. #" & Format$(nRecs, "####0")
           DoEvents
        End If
' Qui si fa il confronto tra i dati per la verifica degli ODL

      For i = 0 To nFields - 1
         If rsMFG.Fields(i).Name <> "Id" Then
            rsMFG(i) = rsImport(i)
         End If
      Next i
        rsMFG.Update
' Controllo la possibilità di errori di duplicatura e li segnalo
        If Err.Number <> 0 Then
           Select Case Err.Number
              Case 3022
                 MsgBox "Duplicato ID: " & rsMFG("ODL")
              Case Else
                 MsgBox "Errore: " & Str$(Err.Number) & vbCrLf & _
                       Err.Description
           End Select
           Err.Clear
        End If
        rsMFG.Close
        rsImport.MoveNext
   Loop
   lblProgress.Visible = False
   With lblLog
'      .Caption = lblLog.Caption & vbCrLf &
'      .Caption = vbCrLf & _
       "Caricati " & Str$(rsMFG.RecordCount) & _
               " ODL distinti in tabella MFG"
      .ForeColor = &HFF0000
      .FontSize = 12
      .Alignment = 2
   End With
   DoEvents
   sLabel = sLabel & vbCrLf & _
       " di cui  " & Str$(rsImport.RecordCount) & _
               " ODL  distinti"

   rsImport.Close
  
   Set rsImport = Nothing
   Set rsMFG = Nothing
'  Chiusura db temporaneo
   dbTemp.Close
   lblLog.Caption = vbCrLf & vbCrLf & sLabel & vbCrLf & vbCrLf & _
            "Fine esecuzione"
                                                                                                                                                                                                                                                               
   Screen.MousePointer = vbNormal
   Me.Enabled = True
   cmdEsci.Enabled = True
'
End Sub

Private Sub Form_Load()
   Dim sqlc As String, rsTemp As Recordset, i As Integer
    fClp = False
'   txtFileName.Text = App.Path & "\" & "WoNoteCLP.txt"

   lblProgress.Visible = False
   lblLog.Visible = False
   
   'sqlc = "Select CampoMFG from CorrImportMFG order by CampoImport"
   'Set rsTemp = db.OpenRecordset(sqlc, dbOpenSnapshot)
'   i = 0
   'Do While Not rsTemp.EOF
   '   ImpMfg(i) = rsTemp(0)
   '   rsTemp.MoveNext
   '   i = i + 1
   'Loop
   'rsTemp.Close
   'Set rsTemp = Nothing
End Sub
Public Property Let Key(vValue As String)
    mKey = Trim$(vValue)
'If Me.Key <> "Aggiorna" Then
'    fClp = False
'   txtFileName.Text = App.Path & "\" & "WoNoteCLP.txt"
'     sqlc = "Select CampoMFG from CorrImportMFG order by CampoImport"
'   Set rsTemp = db.OpenRecordset(sqlc, dbOpenSnapshot)
'   i = 0
'   Do While Not rsTemp.EOF
'      ImpMfg(i) = rsTemp(0)
'      rsTemp.MoveNext
'      i = i + 1
'   Loop
'   rsTemp.Close
'   Set rsTemp = Nothing
'
'    CercaID
' Else
'    fClp = True
'    txtFileName.Text = App.Path & "\" & "OdlCLP.txt"
'
'    'Carico i dati per la tabella generale degli ordini
'    sqlc = "Select CampoMFG from CorrImportCLP order by CampoImport"
'
'    Set rsTemp = db.OpenRecordset(sqlc, dbOpenSnapshot)
'    i = 0
'    Do While Not rsTemp.EOF
'        ImpCLP(i) = rsTemp(0)
'        rsTemp.MoveNext
'        i = i + 1
'    Loop
'
'    rsTemp.Close
'    Set rsTemp = Nothing
'
'
' '   CercaComponenti
' End If
Campi
End Property
Public Property Get Key() As String
    Key = mKey
End Property
Private Sub CercaID()

Dim sqlc As String, rsT As Recordset
Dim sMescola, sQta
Dim sCim As String
sMescola = ""
sQta = ""


sqlc = "SELECT Formule.Mescola, SE.SeId, SE.Qta " & _
        "FROM SE INNER JOIN Formule ON " & _
            "SE.SeId = Formule.SeID " & _
        "GROUP BY Formule.Mescola, SE.SeId, SE.Qta " & _
        "HAVING SE.SeId=" & Me.Key
        
Set rsT = db.OpenRecordset(sqlc, dbOpenSnapshot)
    If Not rsT.EOF Then
        sMescola = rsT("Mescola")
        sQta = rsT("Qta")
        
    End If
    
    sFinale = "@@batchload wowomt.p " & vbCrLf & _
                """"" " & """""" & vbCrLf & _
                """" & sMescola & """ - ""346""" & vbCrLf & _
                sQta & " - - - - ""R"" - - - - - - - - ""SeID" & Me.Key & """ ""No""" & vbCrLf & _
                """No""" & vbCrLf & _
                """""" & vbCrLf & _
                "- ""1113""" & vbCrLf & _
                "@@end"
     
        MsgBox sFinale
        sCim = App.Path & "\" & "wowomt.cim"
        
        Open sCim For Output As #1   ' Open file for output.
        Print #1, sFinale
        Close #1   ' Close file.

     
End Sub


Private Sub cmdCLP_Click()
   Dim nFile As Integer, sFileName As String
   Dim sqlc As String, rsImport As Recordset, rsMFG As Recordset
   Dim sLine As String, i As Integer, sTmp As Variant
   Dim nFields As Integer, nRecs As Integer
   Dim sFields(0 To 19) As String
'   *************************************************************************
   Dim hProg, hProc, RetVal As Long
   Dim sCmd As String, sPrg As String, sFile As String, sLabel As String
   
'   Const PROCESS_ALL_ACCESS = 0
'   Const PROCESS_ALL_ACCESS = &H1F0FFF
'   cmdEsci.Enabled = False
'   Me.Enabled = False
'
'
'' Parte lo script di reflection per creare crearmi la
'' for each per lo scarico dei dati
'
'    sPrg = "c:\programmi\r1win\r1win.exe"
'    sFile = App.Path & "\" & "OdlCLP.rbs"
'    sCmd = sPrg & " " & "/rbs" & " " & sFile
'    '   sCmd = "calc.exe"       'prova di funzionamento sincrono
'
'    hProg = Shell(sCmd, vbNormalFocus)
'    hProc = OpenProcess(PROCESS_ALL_ACCESS, False, hProg)
'    If hProc <> 0 Then
'        RetVal = WaitForSingleObject(hProc, INFINITE)
'        CloseHandle hProc
'    End If
'
''  trasporto il file con le note dal server sul pc
'   sCmd = "ftp -s:" & App.Path & "\" & "OdlCLP.scr"
''   sCmd = "ftp -s:" & App.Path & "\" & "ftp.scr" ' originale
''   sCmd = "ftp -s:" & App.Path & "\" & "ftp.txt" ' originale
''   sCmd = "ftp -s:" & App.Path & "\" & "prova.txt"
'   hProg = Shell(sCmd, vbNormalFocus)
'   hProc = OpenProcess(PROCESS_ALL_ACCESS, False, hProg)
'   If hProc <> 0 Then
'      RetVal = WaitForSingleObject(hProc, INFINITE)
'      CloseHandle hProc
'   End If
'
'   Do While IsActive(hProg)
'      DoEvents
'   Loop
'   hProc = OpenProcess(PROCESS_ALL_ACCESS, False, hProg)
'   If hProc <> 0 Then
'      RetVal = WaitForSingleObject(hProc, INFINITE)
'      CloseHandle hProc
'  End If
' ' MsgBox "Fine trasporto FTP"
'
'  Me.Enabled = True
'  Me.SetFocus
'   ******************************************************************************
   Screen.MousePointer = vbArrowHourglass
   
'   cmdImport.Enabled = False
   Me.Enabled = False
'
'  Apertura Db Temporaneo
'
'   Set dbTemp = DAO.OpenDatabase(App.Path & "\" & dbTempName)
   Set dbTemp = dao.OpenDatabase("x:\imballo\imballo.mde")

   lblLog.Caption = ""
   lblLog.Visible = True
   With lblProgress
      .Caption = ""
      .Top = lblLog.Top + lblLog.Height - (.Height * 1.5)
      .Visible = True
   End With
   Me.Refresh

   nRecs = 0

 '  lblLog.Caption = "E' in corso " & vbCrLf & _
                  " la cancellazione della tabella temporanea"
 '  DoEvents

'   sqlc = "Delete from [$$ImportCLP]"
'   dbTemp.Execute (sqlc)

'   sqlc = "Select * from [ODL]"
'   Set rsImport = dbTemp.OpenRecordset(sqlc, dbOpenDynaset)
'   nFields = rsImport.Fields.Count
'
'   lblLog.Caption = lblLog.Caption & vbCrLf & _
'      "Stiamo caricando una nuova tabella "
'   DoEvents
'
'   sFileName = txtFileName.Text
'   nFile = FreeFile
'   Open sFileName For Input As nFile
'   Do While Not EOF(nFile) '
'      Line Input #nFile, sLine
'      nRecs = nRecs + 1
'      If (nRecs Mod 100) = 0 Then
'         DoEvents
'         lblProgress.Caption = "Caricamento rec. #" & Format$(nRecs, "####0")
'      End If
'  If Not IsNull(sLine) And sLine <> "" Then
'      For i = 0 To nFields - 1
'         sFields(i) = Trim$(Mid$(sLine, (i * 13 + 1), 13))
'      Next i
'      rsImport.AddNew
'      For i = 0 To nFields - 1
'         sTmp = Trim$(sFields(i))
'         Select Case rsImport.Fields(ImpCLP(i)).Type
'            Case dbText '
'               If sTmp <> "" Then
'                  rsImport(ImpCLP(i)) = Left$(sTmp, rsImport(ImpCLP(i)).Size)
'               Else
'                  rsImport(ImpCLP(i)) = Null
'               End If
'            Case dbDate
'               If sTmp = "" Then
'                  rsImport(ImpCLP(i)) = Null
'               Else
'                  rsImport(ImpCLP(i)) = sTmp
'               End If
'            Case Else
'               If sTmp = "" Then
'                  rsImport(ImpCLP(i)) = Null
'               Else
'                  rsImport(ImpCLP(i)) = Val(sTmp)
'               End If
'         End Select
'      Next i
'      rsImport.Update
'    End If
'
'   Loop
'
'   Close nFile
'   rsImport.Close
'   Set rsImport = Nothing
'   lblProgress.Visible = False
'
'   lblLog.Caption = vbCrLf & _
'         "Fine caricamento tabella temporanea: " & _
'         Str$(nRecs) & " Records"
'   DoEvents
'   sLabel = "Caricati " & Str$(nRecs) & " Records"

'  Caricamento tabella MFG con i soli ODL Distinti e non nulli
'   lblLog.Caption = ""
'   With lblLog
''      .Caption = lblLog.Caption & vbCrLf & _
'         "Cancellazione  vecchi dati "
'      .ForeColor = &H80000007
'      .FontSize = 10
'      .Alignment = 0
'   End With
'   DoEvents
'Comincio il confronto dei dati
' rsImport è per noi il rsNew

   
'   sqlc = "SELECT [$$ImportMFG].ID, [$$ImportMFG].Note " & _
            "From [$$ImportMFG] " & _
            "WHERE ((([$$ImportMFG].Note) Like 'seid*'))"
    
    sqlc = "SELECT * From [ODL] "
            

   Set rsImport = dbTemp.OpenRecordset(sqlc, dbOpenSnapshot)
        nRecs = 0
        With lblProgress
            .Caption = ""
            .Top = lblLog.Top + (.Height * 0.5)
            .Visible = True
        End With
        nFields = rsImport.Fields.Count
    
    Do While Not rsImport.EOF
        sqlc = "SELECT * FROM [ODL] " & _
                "where Id=" & rsImport("Id")
        Set rsMFG = db.OpenRecordset(sqlc, dbOpenDynaset)
        
        On Error Resume Next
        
      If Not rsMFG.EOF Then
         rsMFG.Edit
      Else
         rsMFG.AddNew
         rsMFG("Id") = rsImport("Id")
      End If
' Questa parte serve solo a caricare la barra ---------------------
        nRecs = nRecs + 1
        If (nRecs Mod 100) = 0 Then
           lblProgress.Caption = "Caricamento rec. #" & Format$(nRecs, "####0")
           DoEvents
        End If
' Qui si fa il confronto tra i dati per la verifica degli ODL

      For i = 0 To nFields - 1
         If rsMFG.Fields(i).Name <> "Id" Then
            rsMFG(i) = rsImport(i)
         End If
      Next i
        rsMFG.Update
' Controllo la possibilità di errori di duplicatura e li segnalo
        If Err.Number <> 0 Then
           Select Case Err.Number
              Case 3022
'                 MsgBox "Duplicato ID: " & rsMFG("ODL")
              Case Else
'                 MsgBox "Errore: " & Str$(Err.Number) & vbCrLf & _
                       Err.Description
           End Select
           Err.Clear
        End If
        rsMFG.Close
        rsImport.MoveNext
   Loop
   lblProgress.Visible = False
   With lblLog
'      .Caption = lblLog.Caption & vbCrLf &
'      .Caption = vbCrLf & _
       "Caricati " & Str$(rsMFG.RecordCount) & _
               " ODL distinti in tabella MFG"
      .ForeColor = &HFF0000
      .FontSize = 12
      .Alignment = 2
   End With
   DoEvents
   sLabel = sLabel & vbCrLf & _
       " di cui  " & Str$(rsImport.RecordCount) & _
               " ODL  distinti"

   rsImport.Close
  
   Set rsImport = Nothing
   Set rsMFG = Nothing
'  Chiusura db temporaneo
   dbTemp.Close
   lblLog.Caption = vbCrLf & vbCrLf & sLabel & vbCrLf & vbCrLf & _
            "Fine esecuzione"
                                                                                                                                                                                                                                                               
   Screen.MousePointer = vbNormal
   Me.Enabled = True
   cmdEsci.Enabled = True
'
End Sub

Private Sub Campi()
'If fClp Then
    
'    cmdImport.Visible = False
    cmdCLP.Visible = True
'    cmdCLP.Left = cmdImport.Left
'    cmdCLP.Top = cmdImport.Top
    
'Else
'    cmdCLP.Visible = False
'    cmdImport.Visible = True
'End If

End Sub
