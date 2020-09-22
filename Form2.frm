VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Converting Options"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   3855
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlg 
      Left            =   3360
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox flvfile 
      Height          =   375
      Left            =   480
      TabIndex        =   21
      Top             =   4320
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Audio"
      Height          =   1095
      Left            =   240
      TabIndex        =   14
      Top             =   2400
      Width           =   3375
      Begin VB.ComboBox audiobitrate 
         Height          =   315
         ItemData        =   "Form2.frx":0CCA
         Left            =   840
         List            =   "Form2.frx":0CE6
         TabIndex        =   4
         Text            =   "32"
         Top             =   300
         Width           =   975
      End
      Begin VB.TextBox echt 
         Height          =   285
         Left            =   2040
         TabIndex        =   5
         Text            =   "44100"
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Bitrate"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Taux d'échantillonnage:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Hz"
         Height          =   255
         Left            =   2880
         TabIndex        =   16
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "kbit/s"
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Video"
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.OptionButton Option2 
         Caption         =   "AVI"
         Height          =   255
         Left            =   1680
         TabIndex        =   20
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "MPEG"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox videobitrate 
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Text            =   "400"
         Top             =   1620
         Width           =   735
      End
      Begin VB.TextBox hauteur 
         Height          =   285
         Left            =   2640
         TabIndex        =   2
         Text            =   "240"
         Top             =   1140
         Width           =   615
      End
      Begin VB.TextBox largeur 
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Text            =   "320"
         Top             =   1140
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "kbit/s"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Bitrate:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Height"
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Width"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Resolution:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Format:"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label Lab 
      Caption         =   """"
      Height          =   135
      Left            =   120
      TabIndex        =   22
      Top             =   3720
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim err As String
If Option1.Value = True Then
With dlg
.DefaultExt = "mpg"
.Filter = "fichier vidéo (*.mpg)|*.mpg"
.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
End With
dlg.ShowSave
If dlg.FileName = "" Then Exit Sub
ShellAndWaitForTermination App.Path + "\ffmpeg.exe -i " + Lab.Caption & _
flvfile.Text & Lab.Caption + " -b " + videobitrate.Text + " -s " & largeur.Text & _
"x" & hauteur.Text & " -ab " & audiobitrate.Text & " -ar " & echt.Text + " -y" & " " & _
Lab.Caption & dlg.FileName & Lab.Caption, vbHide

Else
With dlg
.DefaultExt = "avi"
.Filter = "fichier vidéo (*.avi)|*.avi"
.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNFileMustExist
End With
dlg.ShowSave
If dlg.FileName = "" Then Exit Sub
ShellAndWaitForTermination App.Path + "\ffmpeg.exe -i " + Lab.Caption & _
flvfile.Text & Lab.Caption + " -vcodec msmpeg4" + " -b " + videobitrate.Text & _
" -s " & largeur.Text & "x" & hauteur.Text & " -ab " & audiobitrate.Text & _
" -ar " & echt.Text + " -y" & " " & Lab.Caption & dlg.FileName & Lab.Caption, vbHide
End If

If dlg.FileName <> "" Then
 If FileLen(dlg.FileName) = "0" Then
 MsgBox "Une erreur inconnue est survenue", vbCritical
 Exit Sub
 Else
 MsgBox "Conversion terminée", vbInformation
 Form2.Hide
 End If
End If
End Sub

Private Sub Command2_Click()
Form2.Hide
End Sub
