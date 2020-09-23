VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "App.ProductName+ ""V"" + Str(App.Revision)"
   ClientHeight    =   7170
   ClientLeft      =   2115
   ClientTop       =   1440
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExercise-control aray q2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   6795
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "jkjkjkjkjkjkj"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   6240
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAns 
      Appearance      =   0  'Flat
      Caption         =   "1"
      DownPicture     =   "frmExercise-control aray q2.frx":0442
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   240
      MouseIcon       =   "frmExercise-control aray q2.frx":0884
      TabIndex        =   4
      Top             =   2040
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000014&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   7080
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   3
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3840
      Top             =   2640
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5160
      TabIndex        =   12
      Top             =   4560
      Width           =   2655
   End
   Begin VB.Label lblTimer 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      Top             =   6240
      Width           =   2655
   End
   Begin VB.Label lblAnswerStat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   5160
      TabIndex        =   10
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label lblRandomQuests 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Label lblQuestNum 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5160
      TabIndex        =   8
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label lblQuestion 
      BackStyle       =   0  'Transparent
      Caption         =   "question appears here"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label lblAns 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   2040
      Width           =   6255
   End
   Begin VB.Menu filemenu 
      Caption         =   "&File"
      Begin VB.Menu mnuNewGame 
         Caption         =   "New Game"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Options 
      Caption         =   "&Options"
      Begin VB.Menu crap 
         Caption         =   "Transparency"
         Begin VB.Menu mnu_Disable_Trans 
            Caption         =   "Disable Transparecy"
         End
         Begin VB.Menu mnuEnable_Trans 
            Caption         =   "Enable Transparecy"
         End
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Change Fonts Type"
         WindowList      =   -1  'True
         Begin VB.Menu mnuButtonsFont 
            Caption         =   "Buttons Font"
         End
         Begin VB.Menu mnuAnswersFont 
            Caption         =   "Answers Font"
         End
         Begin VB.Menu mnuQuestionFont 
            Caption         =   "Question Font"
         End
      End
      Begin VB.Menu fontclour 
         Caption         =   "Change Font Colour"
         Begin VB.Menu mnuFont_Black 
            Caption         =   "Font Black"
         End
         Begin VB.Menu mnuFont_White 
            Caption         =   "Font White"
         End
      End
      Begin VB.Menu mnuSetDataFileLoc 
         Caption         =   "Configure"
      End
      Begin VB.Menu mnuUseQeditor 
         Caption         =   "Add New Questions"
      End
   End
   Begin VB.Menu aboutmenu 
      Caption         =   "&About"
      Begin VB.Menu mnuWebsite 
         Caption         =   "&Gerrys Web Site"
      End
      Begin VB.Menu mnuVersionInfo 
         Caption         =   "&Version Info"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim correctAns() As Integer
Dim question() As String, ans1() As String, ans2() As String, ans3() As String, ans4() As String
Public num As Integer
Dim questCount As Integer
Public numCorrect As Integer
Public numWrong As Integer
Public numTries As Integer
Dim starttime As Single, clearStatus As Integer
Dim totquest As Integer

Public randomQuestValue As String

Public NL


Public Sub cmdAns_Click(Index As Integer)
Dim choice As Integer
choice = cmdAns(Index).Index

StatusBar1.SimpleText = "Press a button to play."

'start the timer when button presssed only
If numTries = 0 Then
    starttime = Timer
    Timer1.Enabled = True
End If

numTries = numTries + 1


If choice = correctAns(num) Then
    numCorrect = numCorrect + 1
    Set Picture1.Picture = LoadPicture(CorrectAnsPic)
    
    StatusBar1.SimpleText = "Correct Answer"
    
    lblAnswerStat.Caption = "CORRECT ANSWER:"
    Beep
    Beep
    
    Picture1.BackColor = &H80000005
    lblAnswerStat.BackColor = &H80000005
    
    lblStatus.Caption = "You got " & numCorrect & " answers right." & Chr(13) & "You got " & numWrong & " answers wrong." & Chr(13) & "Your attempts were " & numTries & Chr(13) & questCount & " questions available."
    
    
    '''DISPLAY A MESSAGE INDICATING IF RANDOM QUESTIONS ARE TO BE USED
    If randomQuestValue = "True" Then
        lblRandomQuests.Caption = "Random Questions Enabled"
    Else
        lblRandomQuests.Caption = "Random Questions Disabled"
    End If
      
    num = num + 1
    
    If num <> questCount Then
    
        ''ACTIVATE RANDOM QUESTION LOADING
        If randomQuestValue = "True" Then
            randomQuestion (totquest)
        End If

        Call display_question(num)

    Else
    
    End If

Else
        Set Picture1.Picture = LoadPicture(WrongAnsPic)
        numWrong = numWrong + 1
        lblStatus.Caption = "You got " & numCorrect & " answers right." & Chr(13) & "You got " & numWrong & " answers wrong." & Chr(13) & "Your attempts were " & numTries & Chr(13) & questCount & " questions available."
        
        StatusBar1.SimpleText = "Wrong Answer!"
        
        Picture1.BackColor = &HFF&
        lblAnswerStat.BackColor = &HFF&

        
        lblAnswerStat.Caption = "WRONG ANSWER DUDE."
        Beep
        Beep
   
        ''MOVE TO NEXT QUESTION IF ANSWER IS WRONG
        
        If AdvanceWrong = "True" Then
            ''CALL PROCEDURE
            num = num + 1
            
            Call Advance_if_Wrong(num)

            If num <= questCount Then
                
                ''ACTIVATE RANDOM QUESTION LOADING
                If randomQuestValue = "True" Then
                    randomQuestion (totquest)
                Else
                    Call display_question(num)
                End If

            End If
            
            Call Advance_if_Wrong(num)
            
        Else
            
            
        End If
        
        Picture1.BackColor = &HFF&
        lblAnswerStat.BackColor = &HFF&
            
End If


End Sub


Private Sub cmdClear_Click()
If clearStatus = False Then
    For i = 1 To 4
        cmdAns(i).Visible = False
        lblAns(i).Visible = False
    Next i
lblQuestion.Visible = False
End If
clearStatus = False
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdRandomQuestion_Click()
    randomQuestion (totquest)
End Sub

Public Sub cmdStart_Click()

    lblQuestion.Visible = True
    cmdAns(1).Visible = True
    lblAns(1).Visible = True
    
    For i = 2 To 4
        lblAns(i).Visible = True
        cmdAns(i).Visible = True
    Next i
End Sub

Private Sub Form_Unload(cancel As Integer)
    
    SaveSetting appname, "Data", "Form Font", frmMain.FontName
    SaveSetting appname, "Data", "Form Font", frmMain.FontName
    
    End

End Sub


Private Sub mnu_Disable_Trans_Click()
    Call Form_Trans_Controls_Disable
End Sub

Private Sub mnu_Transparency_Click()
    Call Form_Trans_Controls
End Sub

Private Sub mnuAbout_Click()
    frmAboutInfo.Visible = True
End Sub

Private Sub mnuAnswersFont_Click()
Dim fntAnsFontNme As String
Dim fntAnsFontSze As String

CommonDialog1.Flags = cdlCFScreenFonts
CommonDialog1.ShowFont

fntAnsFontNme = CommonDialog1.FontName
fntAnsFontSze = CommonDialog1.FontSize

For i = 1 To 4
lblAns(i).FontName = fntAnsFontNme
lblAns(i).FontSize = fntAnsFontSze
Next i

''reg settibgs
SaveSetting appname, "Data", "Answers Font", fntAnsFontNme
SaveSetting appname, "Data", "Answers Font Size", fntAnsFontSze

End Sub

Private Sub mnuButtonsFont_Click()
Dim fntButtonsFontNme As String
Dim fntButtonsFontSze As String

CommonDialog1.Flags = cdlCFScreenFonts
CommonDialog1.ShowFont

fntButtonsFontNme = CommonDialog1.FontName
fntButtonsFontSze = CommonDialog1.FontSize

For i = 1 To 4
cmdAns(i).FontName = fntButtonsFontNme
cmdAns(i).FontSize = fntButtonsFontSze
Next i

''save reg button font name
SaveSetting appname, "Data", "Buttons Font", fntButtonsFontNme
SaveSetting appname, "Data", "Buttons Font Size", fntButtonsFontSze


End Sub


Private Sub mnuEnable_Trans_Click()
    Call Form_Trans_Controls
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFont_Black_Click()
    Call Form_Font_Black
End Sub

Private Sub mnuFont_White_Click()
    Call Form_Font_White
End Sub

Private Sub mnuNewGame_Click()
    
    Timer1.Enabled = False
    lblTimer.Caption = "READY"
    num = 1
    Call display_question(num)
    
    cmdAns(1).Enabled = True
    cmdAns(2).Enabled = True
    cmdAns(3).Enabled = True
    cmdAns(4).Enabled = True

    
    numCorrect = 0
    numTries = 0
    numWrong = 0
       
    lblStatus.Caption = "You got " & numCorrect & " answers right." & Chr(13) & "You got " & numWrong & " answers wrong." & Chr(13) & "Your attempts were " & numTries & Chr(13) & questCount & " questions available."

    
End Sub


Private Sub mnuQuestionFont_Click()
Dim fntQuestionFontNme As String
Dim fntQuestionFontSze As String

CommonDialog1.Flags = cdlCFScreenFonts
CommonDialog1.ShowFont

fntQuestionFontNme = CommonDialog1.FontName
fntQuestionFontSze = CommonDialog1.FontSize

lblQuestion.FontName = fntQuestionFontNme
lblQuestion.FontSize = fntQuestionFontSze

''save reg questions font name
SaveSetting appname, "Data", "Questions Font", fntQuestionFontNme
SaveSetting appname, "Data", "Questions Font Size", fntQuestionFontSze

End Sub

Private Sub mnuSetDataFileLoc_Click()
    frmOptionsMultChoiceQ.Visible = True
End Sub

Private Sub mnuUseQeditor_Click()
    frmQuestionEditor.Visible = True
End Sub

Private Sub mnuVersionInfo_Click()
    frmAbout.Visible = True
End Sub

Private Sub mnuWebsite_click()

response = MsgBox("Do you want to go to see gerrys Web site?", vbYesNo, "Gerrys Web Site")

If response = vbYes Then
    Call CallWebSite
End If

End Sub

Private Sub CallWebSite()

    Shell "start http://go.to/wdwtbam"
      
End Sub

Public Sub Form_Load()

NL = Chr(13) & Chr(10)

num = 1
Dim clearStatus As Boolean

''''DISPLAT SPLASG SCREEN FRO 5 SECS
Dim PauseTime, Start, Finish, TotalTime
If vbYes = vbYes Then
    PauseTime = 1   ' Set duration.
    Start = Timer   ' Set start time.
    Do While Timer < Start + PauseTime
        frmSplash.Visible = True
        DoEvents    ' Yield to other processes.
        
        ''UPDATE THE PROGRESS BAR
        Call Update_ProgressBar
        ''GET APPNAME AND DATFILE LOCATION
        Call Get_Appname_Datafile
        
        Call Update_ProgressBar

        ''LOAD QUESTIONS
        Call Load_Question
        
        Call Update_ProgressBar

        Call Count_Questions

        
    Loop
    Finish = Timer  ' Set end time.
    TotalTime = Finish - Start  ' Calculate total time.
Else
    End
End If




If cmdClear.Value = True Then
    clearStatus = True
End If


lblAnswerStat.Caption = "ready"
lblStatus.Caption = "PRESS Start to start Quiz."
lblTimer.Caption = "ready"
lblRandomQuests.Caption = "Ready"
Timer1.Enabled = False


cmdAns(1).Visible = False
lblQuestion.Visible = False
lblAns(1).Visible = False

For i = 2 To 4
    Load cmdAns(i)
    Load lblAns(i)
    cmdAns(i).Top = cmdAns(i - 1).Top + cmdAns(1).Height + 300
    
    lblAns(i).Top = lblAns(i - 1).Top + cmdAns(1).Height + 300
    lblAns(i).Visible = False
    cmdAns(i).Visible = False
    
    cmdAns(i).Caption = i
    lblAns(i).Caption = ""

Next i

Call Update_ProgressBar

Call Load_Question
Call display_question(num)



frmMain.Caption = formname
'Prog.Caption = GetSetting(AppName, "Data", "Program", "")

''LOAD REG SAVES SETTINGS FOR ANSWERS FONT SETTINGS
fntAnsFontNme = GetSetting(appname, "Data", "Answers Font", "")
fntAnsFontSze = GetSetting(appname, "Data", "Answers Font Size", "")

''LOAD FOR BUTTONS FONT AND SIZE
fntButtonsFontNme = GetSetting(appname, "Data", "Buttons Font", "")
fntButtonsFontSze = GetSetting(appname, "Data", "Buttons Font Size", "")

''LOAD SETTINGS FOR QUESTIONS FONT AND SIZE
fntQuestionFontNme = GetSetting(appname, "Data", "Questions Font", "")
fntQuestionFontSze = GetSetting(appname, "Data", "Questions Font Size", "")

''SEE IF RANDOM QUESTIONS ARE WANTED
randomQuestValue = GetSetting(appname, "Data", "Random Quests", "")

''SEE IF ADVANCE TO NEXT QUESTION IS WANTED
AdvanceWrong = GetSetting(appname, "Data", "Advance Wrong Ans", "")



''CORRECT ANS AND WRONG ANS PICTURE FILES
WrongAnsPic = GetSetting(appname, "Data", "Wrong Ans Pic", "")
CorrectAnsPic = GetSetting(appname, "Data", "Correct Ans Pic", "")

frmOptionsMultChoiceQ.txtCorrectAnsPic.Text = CorrectAnsPic
frmOptionsMultChoiceQ.txtWrongAnsPic.Text = WrongAnsPic



''CHECK TO SEE IF SETTING IS SAVED FOR PICTURE FILES
If WrongAnsPic = "" Then
    Picture1.Picture = LoadPicture()
'Else
End If


If CorrectAnsPic = "" Then
    Picture1.Picture = LoadPicture()
End If

Call Update_ProgressBar

''''''''''''''''''''''''''''''''''''''''''
''''LOAD FORM BACKGROUND
QuizBackGround = GetSetting(appname, "Data", "Main Form Skin", "")
QuizBackGroundEnable = GetSetting(appname, "Data", "Enable Main Skin", "")

frmOptionsMultChoiceQ.txtMainPicture = QuizBackGround

If QuizBackGroundEnable = "False" Then
    frmMain.Picture = LoadPicture()
    ''QuizBackGround = ""
    frmOptionsMultChoiceQ.chkQuizBackGround.Value = 0
    ''frmOptionsMultChoiceQ.txtMainPicture = QuizBackGround

ElseIf QuizBackGroundEnable = "True" Then

    frmOptionsMultChoiceQ.chkQuizBackGround.Value = 1
    frmMain.Picture = LoadPicture(QuizBackGround)
    frmOptionsMultChoiceQ.txtMainPicture = QuizBackGround
    
'''''SEE IF TRANSPARENCY IS WANTED

    TransState = GetSetting(appname, "Data", "TransState", "")

'''ENABLE TRANSPARENT CONTROLS ON FORM
  If TransState = "True" Then
     Call Form_Trans_Controls
  Else
     Call Form_Trans_Controls_Disable
  End If
    
End If
'''''''''''''''''''''''''''''''''''''''''''''''''
Call Update_ProgressBar


'''CHECK BUTTON TYPE IF TRUE THEN USE LETTERS, FALSE NUMBERS
ButtonType = GetSetting(appname, "Data", "Button Type", "")

If ButtonType = "True" Then
        
        cmdAns(1).Caption = "A"
        cmdAns(2).Caption = "B"
        cmdAns(3).Caption = "C"
        cmdAns(4).Caption = "D"
        frmOptionsMultChoiceQ.chkButtonType.Value = 1

Else
    ''leave a alone
        cmdAns(1).Caption = "1"
        cmdAns(2).Caption = "2"
        cmdAns(3).Caption = "3"
        cmdAns(4).Caption = "4"
        frmOptionsMultChoiceQ.chkButtonType.Value = 0

End If



frmOptionsMultChoiceQ.txtDataFileLoc.Text = dataFileLoc


''IF THE SETTING ISNT SAVED
If fntAnsFontNme = "" Then
    fntAnsFontNme = "Arial"
End If
If fntAnsFontSze = "" Then
    fntAnsFontSze = "12"
End If


''IF THE SETTING ISNT SAVED
If fntButtonsFontNme = "" Then
    fntButtonsFontNme = "Arial"
End If
If fntButtonsFontSze = "" Then
    fntButtonsFontSze = "12"
End If

''IF NOT SAVED
If fntQuestionFontNme = "" Then
    fntQuestionFontNme = "Arial"
End If
If fntQuestionFontSze = "" Then
    fntQuestionFontSze = "12"
End If


''''''''ACTIVATE THE SETTINGS AS LOADED FROM THE REGISTRY


''SET FONT SETTINGS FROM REG FOR ANSWERS
lblQuestion.FontName = fntQuestionFontNme
lblQuestion.FontSize = fntQuestionFontSze

For i = 1 To 4
    lblAns(i).FontName = fntAnsFontNme
    lblAns(i).FontSize = Val(fntAnsFontSze)
    
    cmdAns(i).FontName = fntButtonsFontNme
    cmdAns(i).FontSize = Val(fntButtonsFontSze)
    
Next i

''ACTIVATE RANDOM QUESTION LOADING
If randomQuestValue = "True" Then
    randomQuestion (totquest)
    frmOptionsMultChoiceQ.chkRandomQuestions.Value = 1
End If

''PUT THE CHECK MARK IF ADVANCE TO NEXT QUESTION IS PICKED
If AdvanceWrong = "True" Then
    ''CALL PROCEDURE
    Advance_if_Wrong (num)
    frmOptionsMultChoiceQ.chkWrongAdv.Value = 1
End If

''SET THE DATA FILE T O APPEAR IN TXT BOC CAPTION

frmOptionsMultChoiceQ.txtDataFileLoc = dataFileLoc


''STATUS BAR DISPLAYS
StatusBar1.SimpleText = "Click start to play."

''''''ENABLE DISABLE TRANSPARENY



End Sub

Public Sub Load_Question()
questCount = 0

Open dataFileLoc For Input As #1



    Do While Not EOF(1)
        questCount = questCount + 1
        ReDim Preserve question(1 To questCount)
        ReDim Preserve ans1(1 To questCount)
        ReDim Preserve ans2(1 To questCount)
        ReDim Preserve ans3(1 To questCount)
        ReDim Preserve ans4(1 To questCount)
        ReDim Preserve correctAns(1 To questCount)
        
        Input #1, question(questCount), ans1(questCount), ans2(questCount), ans3(questCount), ans4(questCount)
        Input #1, correctAns(questCount)
        
    Loop

Close #1

End Sub



Public Sub display_question(ByRef num As Integer)

If num = (questCount + 1) Then

    MsgBox "You Have Reached the End the Quiz." & Chr(13) & "You got " & numCorrect & " answers right." & Chr(13) & "You got " & numWrong & " answers wrong." & Chr(13) & "Your attempts were " & numTries, vbInformation
    ''code to clear form
    cmdAns(1).Enabled = False
    cmdAns(2).Enabled = False
    cmdAns(3).Enabled = False
    cmdAns(4).Enabled = False

Else
    lblAns(1).Caption = ans1(num)
    lblAns(2).Caption = ans2(num)
    lblAns(3).Caption = ans3(num)
    lblAns(4).Caption = ans4(num)
    lblQuestion.Caption = question(num)
    lblQuestNum.Caption = "Quest No." & num
    
    Picture1.BackColor = &H80000005
    lblAnswerStat.BackColor = &H80000005
End If

End Sub

Private Sub pp_Click()

End Sub

Private Sub Timer1_Timer()
    
    Dim elapsedtime As Single
    Dim temp As Single
    temp = Timer
    elapsedtime = temp - starttime
    lblTimer.Caption = "Elapsed Time : " & Format(elapsedtime, "#####")

End Sub

Public Function randomQuestion(ByVal totquest As Integer) As Integer

Dim MyValue
Randomize Timer  ' Initialize random-number generator.

MyValue = Int((totquest * Rnd) + 1)    ' Generate random value between 1 and 6.
'MsgBox MyValue
 num = MyValue
display_question (MyValue)
End Function

Public Sub Get_DataFile()

    On Error GoTo cancel
    'frmOptionsMultChoiceQ.Visible = True
    CommonDialog1.Filter = "Data File (*.TXT)|*.TXT|All|*.*"
    
    CommonDialog1.CancelError = True
    CommonDialog1.DialogTitle = "Select the Data File to Use:"
    CommonDialog1.Flags = &H4
    CommonDialog1.ShowOpen
    
    If Dir(CommonDialog1.filename) <> "" Then
    
    dataFileLoc = CommonDialog1.filename
    SaveSetting appname, "Data", "Data File Loc", dataFileLoc

    MsgBox dataFileLoc
    Else
    MsgBox "No such program dude", vbCritical
    End If

cancel:

End Sub

Public Sub Get_Appname_Datafile()
''CALL PROCDURE TO GET APP NAME AND DATAFILE LOCATION

''REG SETTINGS FOR PROGRAM NAME
appname = App.ProductName
formname = App.ProductName & " BETA V " & Str(App.Major) & "." & Str(App.Minor) & "." & (App.Revision)


''GET DATA FILE LOCATION FROM REG
dataFileLoc = GetSetting(appname, "Data", "Data File Loc", "")

If dataFileLoc = "" Then
    ''MsgBox "You will now be asked to set the datafile location for this program.  The datafiel contains the questions which are being used in this quiz.", vbInformation, "Set Data File Location"
    ''MsgBox "ERROR:  Data File Path not specified, you must set the data file to use.  The data file should be in the directory as this program e.g; c:\program files\gerrys muliple choice quests\.  YOu only have to do this once i.e when you start the program for the first time." & Chr(13) & Chr(13) & "NOTE: You MUST set data file location.", vbExclamation, "ERROR: Datafile not Specified"
    ''Call Get_DataFile
    '' SET DATA FILE TO DEFAULT LOCATION
    
    MsgBox "Thanks for using this great product, since its your first use im gonna set up the questions and answers data file for you.  Setting up data file for first use", vbInformation, "First use: Loading Datafile"
    dataFileLoc = App.Path & "\questionsandanswers.txt"
    
    SaveSetting appname, "Data", "Data File Loc", dataFileLoc
    
    frmOptionsMultChoiceQ.txtDataFileLoc.Text = dataFileLoc
End If

frmOptionsMultChoiceQ.txtDataFileLoc.Text = dataFileLoc

End Sub


Public Sub Count_Questions()
''COUNT NUMBER OF QUESTIONS

Open dataFileLoc For Input As #2

totquest = 0

    Do While Not EOF(2)
        totquest = totquest + 1
        Input #2, temp1, temp2, temp3, temp4, temp5, ans
    Loop
    
Close #2

End Sub


Public Sub Advance_if_Wrong(ByVal num As Integer)
    
    Call display_question(num)

End Sub
