VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "My FLV Player & Converter by Adilou"
   ClientHeight    =   4830
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4845
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4830
   ScaleWidth      =   4845
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox dragpic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   -120
      ScaleHeight     =   1305
      ScaleWidth      =   4905
      TabIndex        =   10
      Top             =   1080
      Width           =   4935
      Begin VB.Label draglabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Drag && Drop files here"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   840
         TabIndex        =   11
         Top             =   480
         Width           =   2940
      End
   End
   Begin VB.PictureBox restorepic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   320
      Left            =   0
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":0CCA
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   9
      Top             =   50
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   720
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4440
      Picture         =   "Form1.frx":20BC
      ScaleHeight     =   345
      ScaleWidth      =   315
      TabIndex        =   6
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   350
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   3120
   End
   Begin VB.PictureBox resizepic 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4320
      MouseIcon       =   "Form1.frx":2776
      MousePointer    =   8  'Size NW SE
      Picture         =   "Form1.frx":2A80
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   5
      Top             =   4080
      Width           =   495
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   -240
      ScaleHeight     =   705
      ScaleWidth      =   5145
      TabIndex        =   4
      Top             =   4080
      Width           =   5175
      Begin VB.PictureBox fullscreen 
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   3000
         Picture         =   "Form1.frx":2C9E
         ScaleHeight     =   300
         ScaleWidth      =   1500
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox OpenFile 
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   360
         Picture         =   "Form1.frx":4450
         ScaleHeight     =   300
         ScaleWidth      =   1260
         TabIndex        =   7
         Top             =   120
         Width           =   1260
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact: adilou89@hotmail.com"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   480
         Width           =   3135
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -240
      Picture         =   "Form1.frx":5842
      ScaleHeight     =   345
      ScaleWidth      =   55770
      TabIndex        =   0
      Top             =   0
      Width           =   55800
      Begin VB.Label status 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3550
         TabIndex        =   3
         Top             =   105
         Width           =   855
      End
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '==========================================='
    '           My FLV Player by Adilou         '
    '        Contact: adilou89@hotmail.com      '
    '                                           '
    '           Based on FLV Player 3.7         '
    '           http://www.jeroenwijering.com/  '
    '===========================================
Private Declare Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As Long, ByVal uFlags As Long, dwItem1 As Any, dwItem2 As Any)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public WithEvents ctrl As VBControlExtender
Attribute ctrl.VB_VarHelpID = -1

Private Const SHCNE_ASSOCCHANGED = &H8000000
Private Const SHCNF_IDLIST = &H0&

Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1

Const HTCAPTION = 2
Dim tempfile As String
Dim currentheight As Long
Dim currentwidth As Long
Dim p As Integer

Private Sub Form_Load()
On Error Resume Next

'Assoiciate with FLV files
Association ".flv", "image", App.Path & "\" & App.EXEName & ".exe"

'Load height & width from ressources
Form1.Width = (320 * Screen.TwipsPerPixelX)
Form1.Height = (260 * Screen.TwipsPerPixelY) + 375 + Picture3.Height

'Move the resize pic
resizepic.Left = Form1.Width - resizepic.Width
resizepic.Top = Form1.Height - resizepic.Height


Form1.BackColor = &H0&
EnableDragDrop Me.hwnd  'Activation du Drag'n'Drop

 Dim Buf As String * 128       'Pour le dossier temporaire
 Dim Value As Integer          '
Value = GetTempPath(128, Buf)  '
tempdir = Left(Buf, Value)     '


Randomize Timer 'Pour obtenir des nombres vraiment aléatoires
tempfile = tempdir & Int(Rnd * 100000) + 187  'on génère un nom de fichier aléatoire
extractress 101, tempfile  'Extraction de la ressource

'Ajout dynamique du conrôle. Consulter ce tuto
'http://www.vbfrance.com/tutoriaux/AJOUTER-CONTROLE-OCX-DYNAMIQUEMENT-PLEINE-EXECUTION-LATE-BINDING_361.aspx
Set ctrl = Form1.Controls.Add("ShockwaveFlash.ShockwaveFlash", "flashplayer")
ctrl.Visible = True

ctrl.Width = 0
ctrl.Height = 0

Call ctrl.LoadMovie(0, tempfile)
ctrl.BackgroundColor = &H0&
ctrl.Menu = False

dragpic.Top = 375
dragpic.Left = 0
dragpic.Height = Form1.Height - Picture3.Height - Picture3.Height
dragpic.Width = Form1.Width
resizepic.Visible = False
'Check the command line
If Not lignecommande("") = "" Then Call GotADrop(lignecommande(""))
End Sub


'To extract the ressource
Public Sub extractress(ress As Integer, nom_fich As String)
Dim tab_ani() As Byte
Open nom_fich For Binary Access Write As 1
tab_ani = LoadResData(ress, "CUSTOM")
ReDim Preserve tab_ani(UBound(tab_ani))
Put 1, , tab_ani
Close 1
End Sub

'When a file is droped into the form
Public Sub GotADrop(ByVal flvfile As String)
    flvfile = tempfile & "?file=" & flvfile
    flvfile = Replace(flvfile, "\", "/")
    
    Form1.Controls.Remove "flashplayer"  'We remove the control
    'And we add another
    Set ctrl = Form1.Controls.Add("ShockwaveFlash.ShockwaveFlash", "flashplayer")
    ctrl.Visible = True
    ctrl.BackgroundColor = &H0&
    ctrl.Top = 375
    'We set the flashvars
    ctrl.FlashVars = "autostart=true"
    
    'We load the flash player
    Call ctrl.LoadMovie(0, flvfile)
    
    ctrl.Height = Form1.Height - Picture3.Height
    ctrl.Width = Form1.Width
    
    fullscreen.Visible = True
    Form1.Width = (320 * Screen.TwipsPerPixelX)
    Form1.Height = (240 * Screen.TwipsPerPixelY) + 375 + Picture3.Height
    dragpic.Visible = False
    resizepic.Visible = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
Static bReEntry As Boolean
If bReEntry Then Exit Sub
bReEntry = True

resizepic.Left = Width - resizepic.Width
resizepic.Top = Height - resizepic.Height
Picture3.Top = Height - Picture3.Height
Picture4.Left = Width - Picture4.Width

  ctrl.Height = Form1.Height - Picture3.Height - 375
  ctrl.Width = Form1.Width


bReEntry = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DisableDragDrop Me.hwnd 'Désactivation du Drag'n'Drop
    
    If tempfile <> "" Then Kill tempfile
End Sub

Private Sub mpeg_Click()
Form2.Show
End Sub

Private Sub ctrl_ObjectEvent(info As EventInfo)
    Dim i As Long
    Dim nbArgs As Long
    Dim msg As String
    If info.Name = "FSCommand" Then
    nbArgs = info.EventParameters.Count
    msg = msg & "Evenement : " & info.Name & vbCrLf
    If CStr(info.EventParameters(i).Value) = "play" Then status.Caption = "- Playing"
    If CStr(info.EventParameters(i).Value) = "pause" Then status.Caption = "- Paused"
    For i = 0 To nbArgs - 1
        msg = msg & "Argument n. " & CStr(i) & " name = " & _
            CStr(info.EventParameters(i).Name & _
            " valeur = " & CStr(info.EventParameters(i).Value)) & _
            vbCrLf
    Next i
    End If
End Sub




Private Sub fullscreen_Click()
'Save the current height and width
currentheight = Form1.Height
currentwidth = Form1.Width

 largeur% = Screen.Width \ Screen.TwipsPerPixelX
 hauteur% = Screen.Height \ Screen.TwipsPerPixelY
 Form1.Top = 0
 Form1.Left = 0
 Form1.Width = (Str$(largeur%) * Screen.TwipsPerPixelX)
 Form1.Height = (Str$(hauteur%) * Screen.TwipsPerPixelY)
 Picture1.Visible = False
 Picture3.Visible = False
 ctrl.Top = 0
 ctrl.Width = Form1.Width
 ctrl.Height = Form1.Height
 
 resizepic.Visible = False
 restorepic.Left = Width - Picture4.Width - 1450
 restorepic.Visible = True
End Sub

Private Sub Label5_Click()
Call fullscreen_Click
End Sub



Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ValRetour As Long
  Call ReleaseCapture 'on appelle l'api ici
  ValRetour = SendMessage(Form1.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub resizepic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
p = 1
End Sub

Private Sub resizepic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If p = 1 Then
Width = X + Width
Height = Y + Height
End If
End Sub

Private Sub resizepic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
p = 0
End Sub


Private Sub Picture4_Click()
 Call Form_Unload(0)
 End
End Sub

Private Sub OpenFile_Click()
With dlg
.DefaultExt = "flv"
.Filter = "Flash video file (*.flv)|*.flv"
.Flags = cdlOFNExplorer
End With
dlg.ShowOpen
If dlg.FileName = "" Then Exit Sub
GotADrop (dlg.FileName)
End Sub

Private Sub restorepic_Click()
  Form1.Width = currentwidth
  Form1.Height = currentheight
  ctrl.Width = currentwidth
  ctrl.Height = currentheight - Picture3.Height
  Picture1.Visible = True
  Picture3.Visible = True
  resizepic.Visible = True
  Picture4.Left = Width - Picture4.Width
  ctrl.Top = 375
  restorepic.Visible = False
End Sub


Private Sub status_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ValRetour As Long
  Call ReleaseCapture 'on appelle l'api ici
  ValRetour = SendMessage(Form1.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub status_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
p = 1
End Sub

Private Sub status_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If p = 1 Then
Width = X + Width
Height = Y + Height
End If
End Sub


Private Sub Timer1_Timer()
wid = (320 * Screen.TwipsPerPixelX) + (Form1.Width - Form1.ScaleWidth)
Heig = (260 * Screen.TwipsPerPixelY) + 375 + Picture3.Height

If Me.Width <= wid Then
        Me.Width = wid
        ctrl.Width = wid
       End If


    If Me.Height <= Heig Then
        Me.Height = Heig
        ctrl.Height = Form1.Height - Picture3.Height - 375
    End If

  ctrl.Height = Form1.Height - Picture3.Height - 375
  ctrl.Width = Form1.Width
  
End Sub
