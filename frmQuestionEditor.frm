VERSION 5.00
Begin VB.Form frmQuestionEditor 
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1680
      TabIndex        =   10
      Top             =   4725
      Width           =   4110
   End
   Begin VB.TextBox txtAnsC 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1575
      TabIndex        =   9
      Text            =   "Text5"
      Top             =   2940
      Width           =   4110
   End
   Begin VB.TextBox txtAnsB 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1575
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   1995
      Width           =   4110
   End
   Begin VB.TextBox txtAnsD 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1575
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   3675
      Width           =   4110
   End
   Begin VB.TextBox txtAnsA 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1575
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   1050
      Width           =   4110
   End
   Begin VB.TextBox txtQuestion 
      Height          =   540
      Left            =   1470
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   105
      Width           =   5895
   End
   Begin VB.Label Label5 
      Caption         =   "Question"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   0
      TabIndex        =   4
      Top             =   210
      Width           =   1275
   End
   Begin VB.Label Label4 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   315
      TabIndex        =   3
      Top             =   3675
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   315
      TabIndex        =   2
      Top             =   2835
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   315
      TabIndex        =   1
      Top             =   2100
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   315
      TabIndex        =   0
      Top             =   1155
      Width           =   1275
   End
End
Attribute VB_Name = "frmQuestionEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
If txtQuestion <> "" And txtAnsA <> "" And txtAnsB <> "" And txtAnsC <> "" And txtAnsD <> "" Then

    filenum = FreeFile
    Open dataFileLoc For Append As #filenum
    Write #filenum, txtquestiom, txtAnsA, txtAnsB, txtAnsC, txtAnsD
    Close
Else
'fields contain null strings
'dont save
MsgBox "You must enter data in all Fields to save a new question!", vbInformation, "Enter Data"
End If
End Sub

Private Sub Form_Load()
txtQuestion = ""
txtAnsA = ""
txtAnsB = ""
txtAnsC = ""
txtAnsD = ""

''GET DATA FILE LOCATION FROM REG
dataFileLoc = GetSetting(appname, "Data", "Data File Loc", "")

End Sub
