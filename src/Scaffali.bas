Attribute VB_Name = "mScaffali"
Option Explicit
Public Declare Function GetPrivateProfileString Lib "kernel32" _
   Alias "GetPrivateProfileStringA" _
   (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
   ByVal lpDefault As String, ByVal lpReturnedString As String, _
   ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" _
   Alias "WritePrivateProfileStringA" _
   (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
   ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
   lParam As Any) As Long

Public db As Database, wrkCurrent As Workspace
'   Public IdContainer As Long, rsSearch As Recordset
'   Public PesiCheck(0 To 25) As Integer, frm As Form, asCriteri() As String
'Public rsTabsGrid As Recordset, sDbPath As String
Public sDbPath As String, DbName As String  '   potrebbero non servire
'
'   Seguono dei parametri che possono essere utili
'   public parIdSpuntaStrada As String
'   Public parIdClienteVuoti As String, parIdTerminalVuoti As String
'   Public parKmMinimi As Integer
Public fSelectField As Boolean, fDBOk As Boolean, sFlag As String

Public Enum cColors
   cGray = &H8000000F
   cWhite = &H80000005
   cLightRed = &H8080FF
End Enum

Public Enum FormStatus
   fDisplay = 1
   fInsert = 2
   fModify = 3
   fSearch = 4
End Enum

Public Enum Menu
   cmnuApri = 0
'   cmnuNuovo = 1
   cmnuSetPrinter = 2
   cmnuEsci = 4
    
   cmnuCliente = 0
   cmnuCondizPagam = 1
   cmnuTreno = 2
   cmnuFornitoreTreno = 3
   cmnuViaggioNave = 4
   cmnuNave = 5
   cmnuPorto = 6
   cmnuTransitPorto = 7
   cmnuTraspStrada = 8
   cmnuSpedizioniere = 9
   cmnuTerminalInterno = 11
   cmnuTerminalVuoti = 12
   cmnuDogana = 13
   cmnuTipoContainer = 14
   cmnuServizio = 15
   cmnuCompagnia = 16
   cmnuCodiciPorto = 17
   cmnuCVuoti = 18
   cmnuDistanze = 19
    
   cmnuSpStrada = 0
   cmnuSpIntl = 1
   cmnuSpMsc = 2
    
   cmnuGPieniN = 0
   cmnuGPieniC = 1
   
   cmnuGVuoti = 1
   cmnuGUVuoti = 3
   
   cmnuSTCompleto = 0
   cmnuSTCliente = 1
   
   cmnuOpzClear = 0
   cmnuOpzTreno = 2
   cmnuOpzUpdate = 4
   cmnuOpzSelect = 6
   
   cmnuSetTreno = 0
   cmnuuRecalcTreno = 1
   cmnuExport = 3
   cmnuCompact = 5
   cmnuRepair = 6
   cmnuPVuoti = 8
   cmnuStatistiche = 10
    
End Enum

Public Function DataUsa(DataI As String) As Date
   Dim gg As Integer, mm As Integer, aa  As Integer
   gg = Val(Mid$(DataI, 1, 2))
   mm = Val(Mid$(DataI, 4, 2))
   aa = Val(Mid$(DataI, 7, 4))
   DataUsa = DateSerial(aa, mm, gg)
'   DataUsa = "#" & Format$(DataTemp, "mm/dd/Yyyy") & "#"
End Function

Sub Main()
   Dim sCmdLine As String
   Screen.MousePointer = vbArrowHourglass
   sDbPath = App.Path & "\"
   sCmdLine = Command()
   If Len(sCmdLine) > 0 Then
       sDbPath = sDbPath & sCmdLine
   Else
       sDbPath = sDbPath & "scaffali.mdb"
   End If
   
   If Not LeggiIni(sFlag) Then
      MsgBox "Errore - Esco"
      End
   End If
   
   Load frmWait
   
   Load frmSplash
'    frmSplash.Show
'    frmSplash.Refresh
   Set wrkCurrent = DBEngine.Workspaces(0)
   fDBOk = DbOpen(sDbPath)
   
'    Set frmMain = New frmMain
   Load frmMain
   
'   Unload frmSplash
'   frmSplash.WindowState = vbMaximized
'   frmSplash.Enabled = False
    
   If fDBOk Then
       SetAllParameters
   End If
   Screen.MousePointer = vbNormal
   frmMain.Show
End Sub
Sub SetAllParameters()
'   parPreview = SetParameters("Preview")
'   parAliquotaIva = SetParameters("Aliquota Iva")
'   parIdClienteVuoti = SetParameters("Id Cliente Vuoti")
'   parIdTerminalVuoti = SetParameters("Id Terminal Vuoti")
'   parKmMinimi = Val(SetParameters("KmMinimi"))
'   parDirTreniExp = SetParameters("Dir TreniExp")
End Sub
Sub ShowWait(msg As String)
   Dim lMeH As Long, lMeW As Long
'   With Forms("frmWait")
'     lMeH = Forms("frmWait").Height
'      lMeW = Forms("frmWait").Width
'   End With
   lMeH = Forms(0).Height
   lMeW = Forms(0).Width
   With Screen.ActiveForm
'        Forms("frmWait").Top = .Top + (.Height - lMeH) / 2
'        Forms("frmWait").Left = .Left + (.Width - lMeW) / 2
       Forms(0).Top = .Top + (.Height - lMeH) / 2
       Forms(0).Left = .Left + (.Width - lMeW) / 2
   End With
   frmWait.lblMessage.Caption = msg
   frmWait.Show
   frmWait.Refresh
End Sub
Sub HideWait()
   frmWait.Hide
End Sub
Public Sub CenterForm(f As Form)
   f.Top = (Screen.Height - f.Height) / 2
   f.Left = (Screen.Width - f.Width) / 2
End Sub
Public Function DbOpen(DbName As String)
   Dim sCheck As String
   On Error GoTo DbOpenError
   Set db = wrkCurrent.OpenDatabase(DbName, False, False)
   On Error Resume Next
   sCheck = db.TableDefs("$$CheckReclami").Name
   If Err.Number <> 0 Then
       DbOpen = False
       MsgBox "Database NON corretto !!", vbCritical + vbOKOnly, "Apertura DataBase"
   Else
       DbOpen = True
   End If
   Exit Function
    
DbOpenError:
   Dim sMsg As String
   sMsg = "Errore nell'apertura del Database" & vbCrLf & _
       DbName & vbCrLf & vbCrLf & _
       "Numero Errore: " & Err.Number & vbCrLf & Err.Description
   MsgBox sMsg
   DbOpen = False
End Function
Public Sub DbCompact()
   Dim DbPath As String, DbName As String, dbTemp As String
   Dim msg As String, res As Integer
   msg = "Procedo con la Compattazione ? " & vbCrLf & _
       "(Il Database deve essere aperto su questa sola macchina !!)"
   res = MsgBox(msg, vbYesNo + vbQuestion + vbDefaultButton2, "")
   If res = vbYes Then
       DbPath = App.Path & "\"
       DbName = db.Name
       Screen.MousePointer = vbArrowHourglass
       ShowWait "Attendere ...." & vbCrLf & _
           "Compattamento DataBase in Corso"
       dbTemp = DbPath & "$$$Temp"
       db.Close
       Name DbName As dbTemp
       DBEngine.CompactDatabase dbTemp, DbName
       Kill (dbTemp)
       DbOpen (DbName)
       HideWait
       Screen.MousePointer = vbDefault
   End If
End Sub
Public Sub DbRepair()
   Dim DbName As String
   Dim msg As String, res As Integer
   msg = "Procedo con il Recupero ? " & vbCrLf & _
       "(Il Database deve essere aperto su questa sola macchina !!)"
   res = MsgBox(msg, vbYesNo + vbQuestion + vbDefaultButton2, "")
   If res = vbYes Then
       DbName = db.Name
       Screen.MousePointer = vbArrowHourglass
       ShowWait "Attendere ...." & vbCrLf & _
           "Recupero DataBase in Corso"
       db.Close
       DBEngine.RepairDatabase DbName
       DbOpen (DbName)
       HideWait
       Screen.MousePointer = vbDefault
   End If
End Sub


Public Function SetParameters(ParName As String)
   Dim sqlc As String, rsTemp As Recordset, sReturn As Variant
   sqlc = "select Valore from Parametri where Parametro = '" & ParName & "'"
   Set rsTemp = db.OpenRecordset(sqlc, dbOpenSnapshot)
   If Not rsTemp.EOF Then
       sReturn = rsTemp(0)
   Else
       sReturn = ""
   End If
   rsTemp.Close
   Set rsTemp = Nothing
   SetParameters = sReturn
End Function

Public Sub SelectText(Flag As Boolean)
   Dim ctl As Control
   Set ctl = Screen.ActiveControl
   If Not (ctl Is Nothing) Then
      If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Then
         With ctl
            .SelStart = 0
            If Flag Then
               .SelLength = Len(.Text)
            Else
               .SelLength = 0
            End If
         End With
      End If
       
      If TypeOf ctl Is MaskEdBox Then
         With ctl
            .SelStart = 0
            If Flag Then
'               .SelLength = Len(.Text)
               .SelLength = .MaxLength
            Else
               .SelLength = 0
            End If
         End With
      End If
   End If
End Sub

Public Sub xxmain()
   Dim sDbName As String
   sDbName = "C:\Programmi\reclami\costi.mdb"
'    sDbName = "C:\Programmi\reclami\ReclamiVB.mdb"
   Set db = dao.DBEngine.Workspaces(0).OpenDatabase(sDbName)
   Load frmMain
   frmMain.Show
End Sub
Sub ShadeForm(frm As Form)
' , TheColor As Integer, R As Integer, G As Integer, B As Integer
    Dim dS, DW, SM, SH
    Dim i As Integer
    Const Inside_Solid = 6
    Const Copy_Pen = 13
   Const maxColor As Integer = 224
   Const minColor As Integer = 0
   Dim iRealeR As Integer, iRealeG As Integer, iRealeB As Integer
   
    dS = frm.DrawStyle                  'save 'em
    DW = frm.DrawWidth
    SM = frm.ScaleMode
    SH = frm.ScaleHeight
    frm.DrawStyle = vbInsideSolid
    frm.DrawWidth = 22
    frm.ScaleMode = vbPixels
    frm.ScaleHeight = maxColor - minColor + 1
'    For i = 0 To 255
'         Select Case TheColor
'            Case 0
'               frm.Line (0, i)-(frm.Width, i + 1), _
'                  RGB(255 - i, G, B), B 'red
'            Case 1
'               frm.Line (0, i)-(frm.Width, i + 1), _
'                  RGB(R, 255 - i, B), B 'green
'            Case 2
'               frm.Line (0, i)-(frm.Width, i + 1), _
'                  RGB(R, G, 255 - i), B 'blue
'            Case Else
'               MsgBox "Internal Error in Color Selection!"
'         End Select
'    Next i
'    For i = 0 To maxRed
'      frm.Line (0, i)-(frm.Width, i + 1), RGB(maxRed - i, G, B), B
'    Next i
    frm.DrawStyle = vbInsideSolid
    frm.DrawWidth = 50
    frm.ScaleMode = vbPixels
    frm.ScaleHeight = maxColor - minColor + 1
    For i = 0 To frm.ScaleHeight - 1
'Originale
'      frm.Line (0, i)-(frm.Width, i + 1), RGB(maxColor - i, maxColor - i, maxColor - i), B
'      frm.Line (0, i)-(frm.Width, i + 1), RGB(maxColor - i, maxColor - i, 224), B
 
'      If i < 20 Then
'         iRealeR = maxColor - i
'      Else
'         iRealeR = (maxColor - i) + 20
'      End If
'
'      If i < 154 Then
'            iRealeG = maxColor - i
'      Else
''         iRealeG = (maxColor - i) + 154
'      End If
'
'       If i < 216 Then
'            iRealeG = maxColor - i
' '     Else
'         iRealeG = (maxColor - i) + 216
'      End If
      
'      iReale = maxColor - i
   
   If i < 224 Then
      iRealeR = maxColor - i
      iRealeG = maxColor - i
      iRealeB = maxColor - i
   Else
      iRealeR = (maxColor - i) + 20
      iRealeG = maxColor - i + 154
      iRealeB = maxColor - i + 216
   End If
      
      frm.Line (0, i)-(frm.Width, i + 1), RGB(iRealeR, iRealeG, iRealeB), B
      
    
    Next i
    frm.DrawStyle = dS          'restore the settings
    frm.DrawWidth = DW
    frm.ScaleHeight = SH        'must be restored before ScaleMode
    frm.ScaleMode = SM
End Sub

Private Function LeggiIni(sParam As String) As Boolean
    Dim sTxtLine As String, nFile As Integer
    Dim sFile As String, fRes As Boolean
    
    fRes = True
 ' Modifica inserimento ini
    sFile = "c:\" + "SCAFFALI" + ".ini"
    
    nFile = FreeFile
    On Error Resume Next
    
    Open sFile For Input As nFile
    
    If Err.Number <> 0 Then
        MsgBox "Errore nella lettura del file INI - Esco dall'Applicazione !"
        fRes = False
    End If
    
    If fRes Then
        Do While Not EOF(nFile)
            Line Input #1, sTxtLine
            sTxtLine = Trim$(UCase(sTxtLine))
            Select Case sTxtLine
                Case "ATC"
                    sParam = "ATC"
                Case "MKG"
                    sParam = "MKG"
                Case "ATC/MKG"
                    sParam = "ATC/MKG"
                Case "VIEW" ' modifica per inserire anche il visualizza
                    sParam = "VIEW"
                Case Else
                    MsgBox "Opzione non riconosciuta: " & sTxtLine
                    fRes = False
            End Select
        Loop
    End If
    
    Close #nFile

   LeggiIni = fRes
End Function

Public Function CalcolaScarto(dS As Double, dC As Double) As Double
                                'dS è lo scarto, dC è il costo
   Dim dPrimo As Double, dDenom As Double, dTot As Double
   
   dPrimo = dS / 100
   dDenom = 1 - dPrimo
   
   dTot = Format((dPrimo * dC) / dDenom, "#0.00")
'   dTot = Format(dTot, "0,00")
   
   CalcolaScarto = dTot

End Function

Public Function ValID(Valore As String)

Dim iPos As Integer, s As String
   s = Valore
   iPos = InStr(s, ",")
   If iPos <> 0 Then
      s = Left$(s, iPos - 1)
   End If
    ValID = s

End Function


