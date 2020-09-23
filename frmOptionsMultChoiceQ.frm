VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptionsMultChoiceQ 
   Caption         =   "Configuration"
   ClientHeight    =   6615
   ClientLeft      =   2130
   ClientTop       =   1980
   ClientWidth     =   7905
   ControlBox      =   0   'False
   Icon            =   "frmOptionsMultChoiceQ.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7905
   Begin VB.CheckBox chkQuizBackGround 
      Caption         =   "Enable Quiz Background"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3480
      Width           =   3255
   End
   Begin VB.CheckBox chkButtonType 
      Caption         =   "Enable Buttons Title to Letters"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3480
      Width           =   3375
   End
   Begin VB.TextBox txtMainPicture 
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
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3000
      Width           =   7335
   End
   Begin VB.TextBox txtCorrectAnsPic 
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
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Text            =   "Text2"
      ToolTipText     =   "Select the picture file you want when a correct answer occurs."
      Top             =   2280
      Width           =   7335
   End
   Begin VB.TextBox txtWrongAnsPic 
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
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Text            =   "Text1"
      ToolTipText     =   "Select the picture file you want to appera when a wrong answer occurs."
      Top             =   1560
      Width           =   7335
   End
   Begin VB.CheckBox chkWrongAdv 
      Caption         =   "Advance to next Question if the wrong answer is selected."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Allows you to keep going until you get the correct answer if enabled, then the next question will appear."
      Top             =   4200
      Width           =   3375
   End
   Begin VB.CheckBox chkRandomQuestions 
      Caption         =   "Random Questions"
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Makes the questions appear in a roandom order each time."
      Top             =   4320
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Configure:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   7575
      Begin VB.TextBox txtDataFileLoc 
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
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   600
         Width           =   7335
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5520
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblQuizSkin 
         Caption         =   "Set Quiz Background"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label lblCorrectAnsPic 
         Caption         =   "Correct Answer Picture"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label lblWrongAnsPic 
         Caption         =   "Wrong Answer Picture"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label lblDataFileLocation 
         BackColor       =   &H80000004&
         Caption         =   "Set Data File Location"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
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
      Left            =   3240
      TabIndex        =   0
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label lblConfigurationStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   5040
      Width           =   7575
   End
End
Attribute VB_Name = "frmOptionsMultChoiceQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dataFileLoc As String, randomQuestValue As String
Dim appname As String

Public WrongAnsPic As String, CorrectAnsPic As String
Public QuizBackGround As String, ButtonType As String
Public QuizBackGroundEnable As String


Private Sub chkButtonType_Click()
    If chkRandomQuestions.Value = vbChecked Then
        'MsgBox "true"
        ButtonType = "True"
    ElseIf chkRandomQuestions.Value = vbUnchecked Then
        'MsgBox "false"
        ButtonType = "False"
    End If
    
    SaveSetting appname, "Data", "Button Type", ButtonType

End Sub

Private Sub chkButtonType_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblConfigurationStatus.Caption = "Click this to Make the buttons appear ans numbers (1-4) or letters (A-D)."

End Sub

Private Sub chkQuizBackGround_Click()
    If chkQuizBackGround.Value = vbChecked Then
        ''txtMainPicture.Enabled = True
        QuizBackGroundEnable = "True"
    Else
        QuizBackGroundEnable = "False"
        txtMainPicture.Enabled = False
        ''QuizBackGround = ""

    End If
    
    SaveSetting appname, "Data", "Enable Main Skin", QuizBackGroundEnable

    
End Sub

Private Sub chkQuizBackGround_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        lblConfigurationStatus.Caption = "Enable/Disable the Quiz background.  Puts the form back to the windows default colour."

End Sub

Public Sub chkRandomQuestions_Click()

    If chkRandomQuestions.Value = vbChecked Then
        'MsgBox "true"
        randomQuestValue = "True"
    ElseIf chkRandomQuestions.Value = vbUnchecked Then
        'MsgBox "false"
        randomQuestValue = "False"
    End If
    
    SaveSetting appname, "Data", "Random Quests", randomQuestValue

End Sub

Private Sub chkRandomQuestions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        lblConfigurationStatus.Caption = "Load Random Questions when the form laods.  Questions will appear in no particular but will be random."

End Sub

''''ALLOW USER TO ADVANCE TO NEXT QUESTION IF THEY PICK
'''WRONG ANSWER.
Private Sub chkWrongAdv_Click()
    If chkWrongAdv.Value = vbChecked Then
        
        AdvanceWrong = "True"
    ElseIf chkWrongAdv.Value = vbUnchecked Then
        
        AdvanceWrong = "False"
    End If
    
End Sub

Private Sub chkWrongAdv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        lblConfigurationStatus.Caption = "Allows you to move on to the next question if you pick the wrongs answer."

End Sub

Private Sub cmdExit_Click()
    SaveSetting appname, "Data", "Advance Wrong Ans", AdvanceWrong
    
    If QuizBackGround <> "" Then
        SaveSetting appname, "Data", "Main Form Skin", QuizBackGround
        
        frmMain.Picture = LoadPicture()
        frmMain.Picture = LoadPicture(QuizBackGround)

    End If


    frmOptionsMultChoiceQ.Visible = False
End Sub

Private Sub Form_Load()
''REG SETTINGS FOR PROGRAM NAME
appname = App.ProductName
formname = App.ProductName + "BETA Ver " + Str(App.Major & App.Minor) & "." & Str(App.Revision)

End Sub

Private Sub txtCorrectAnsPic_click()


On Error GoTo cancel
CommonDialog1.filename = CorrectAnsPic
CommonDialog1.CancelError = True
CommonDialog1.DialogTitle = "Select Picture to Use For Correct Ans"
CommonDialog1.Flags = &H4
CommonDialog1.Filter = "Graphic File (*.ICO)|*.ICO|All|*.*"
CommonDialog1.ShowOpen
If Dir(CommonDialog1.filename) <> "" Then
    
    CorrectAnsPic = CommonDialog1.filename
    txtCorrectAnsPic.Text = CorrectAnsPic
    
    SaveSetting appname, "Data", "Correct Ans Pic", CorrectAnsPic

Else
    MsgBox "No such program dude", vbCritical
End If
cancel:

End Sub

Private Sub txtCorrectAnsPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblConfigurationStatus.Caption = "Select the picture to appear when you get a Correct Ansnwer."

End Sub

Public Sub txtDatafileLoc_click()


On Error GoTo cancel
''CommonDialog1.filename = dataFileLoc
CommonDialog1.CancelError = True
CommonDialog1.DialogTitle = "Select the Data File to Use"
CommonDialog1.Flags = &H4
CommonDialog1.Filter = "Data File (*.TXT)|*.TXT|All|*.*"
CommonDialog1.ShowOpen
If Dir(CommonDialog1.filename) <> "" Then
    
    dataFileLoc = CommonDialog1.filename
    txtDatafileLoc.Text = dataFileLoc
    ''''SAVE THE DATA FILE LOCATION TO REG
    SaveSetting appname, "Data", "Data File Loc", dataFileLoc


Else
    MsgBox "No such program dude", vbCritical
End If
cancel:

End Sub

'''ALLOW USER TO CAHNGE THE BACKGROUND PICTURE PROPERTY OF THE _
'''MAIN FROM

Private Sub txtDataFileLoc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblConfigurationStatus.Caption = "Set the datafile location which contians the questions you want to use."
End Sub

Private Sub txtMainPicture_click()

On Error GoTo cancel
CommonDialog1.filename = QuizBackGround
CommonDialog1.CancelError = True
CommonDialog1.DialogTitle = "Select Picture to Use For Quiz Background"
CommonDialog1.Flags = &H4
CommonDialog1.Filter = "Graphic File (*.BMP;*.JPG;*.GIF)|*.BMP;*.bmp;*.JPG;*.jpg;*.GIF;*.gif| All|*.*"
CommonDialog1.ShowOpen

'''''Text (*.txt)|*.txt|Pictures (*.bmp;*.ico)|*.bmp;*.ico


If Dir(CommonDialog1.filename) <> "" Then
    
    QuizBackGround = CommonDialog1.filename
    txtMainPicture.Text = QuizBackGround
    
    If QuizBackGround <> "" Then
        GoTo cancel
    End If
    
    SaveSetting appname, "Data", "Main Form Skin", QuizBackGround

    
    
Else
    MsgBox "No such program dude", vbCritical
End If
cancel:

End Sub

Private Sub txtMainPicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        lblConfigurationStatus.Caption = "Set the background appearance for the Quiz.  Can be a bitmap, jpg, or gif file."

End Sub

''WrongAnsPic As String, CorrectAnsPic As String

Private Sub txtWrongAnsPic_click()


On Error GoTo cancel
CommonDialog1.filename = WrongAnsPic
CommonDialog1.CancelError = True
CommonDialog1.DialogTitle = "Select Picture to Use For Wrong Ans"
CommonDialog1.Flags = &H4
CommonDialog1.Filter = "Graphic File (*.ICO)|*.ICO|All|*.*"
CommonDialog1.ShowOpen
If Dir(CommonDialog1.filename) <> "" Then
    
    WrongAnsPic = CommonDialog1.filename
    txtWrongAnsPic.Text = WrongAnsPic
    
    SaveSetting appname, "Data", "Wrong Ans Pic", WrongAnsPic

Else
    MsgBox "No such program dude", vbCritical
End If
cancel:

End Sub

Private Sub txtWrongAnsPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblConfigurationStatus.Caption = "Select the picture to appear when you get a wrong answer."

End Sub
