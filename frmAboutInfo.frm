VERSION 5.00
Begin VB.Form frmAboutInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About Program - By Gerard Mc Donnell 99/2000"
   ClientHeight    =   3900
   ClientLeft      =   2940
   ClientTop       =   3150
   ClientWidth     =   6615
   ClipControls    =   0   'False
   Icon            =   "frmAboutInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtAboutProgram 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmAboutInfo.frx":0442
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "frmAboutInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    frmAboutInfo.Visible = False
    frmMain.Visible = True
End Sub

Private Sub Form_Load()
    NL = Chr(13) & Chr(10)
End Sub
