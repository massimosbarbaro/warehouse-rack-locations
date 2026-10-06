VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8805
   ClientLeft      =   2415
   ClientTop       =   1140
   ClientWidth     =   10200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraImages 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   8130
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   9660
      Begin VB.Image imglogo 
         Height          =   1050
         Index           =   3
         Left            =   7080
         Top             =   240
         Width           =   1215
      End
      Begin VB.Image imglogo 
         Height          =   1530
         Index           =   1
         Left            =   720
         Top             =   240
         Width           =   1695
      End
      Begin VB.Image imglogo 
         Height          =   1290
         Index           =   2
         Left            =   7560
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lbltitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scaffalature"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C49030&
         Height          =   1095
         Left            =   2040
         TabIndex        =   2
         Top             =   6840
         Width           =   5025
      End
      Begin VB.Image imglogo 
         Height          =   1530
         Index           =   0
         Left            =   1920
         Top             =   480
         Width           =   3060
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7800
         TabIndex        =   1
         Top             =   5640
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private iImg As Integer

Private Sub Form_Load()
    
    
    iImg = -1
'    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
'    lblProductName.Caption = App.Title
    Me.MousePointer = vbArrowHourglass
    
  
  
    imgLogo(0).Picture = LoadPicture(App.Path & "\sbarbaro.jpg")
    imgLogo(1).Picture = LoadPicture(App.Path & "\SE.jpg")
    imgLogo(2).Picture = LoadPicture(App.Path & "\Bicgo1g.gif ")
    imgLogo(3).Picture = LoadPicture(App.Path & "\perifrag.gif")
    With Me
      .WindowState = vbMaximized
      .Enabled = False
      .Show
   End With
   With fraImages
     .Top = (Me.Height - .Height) \ 2
     .Left = (Me.Width - .Width) \ 2
   End With
   ShadeForm Me

   Me.Refresh
End Sub

