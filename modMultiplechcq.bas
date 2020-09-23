Attribute VB_Name = "Module1"
Public dataFileLoc As String, datafile As String
Public appname As String, formname As String
Public AdvanceWrong As String
Public WrongAnsPic As String, CorrectAnsPic As String
Public QuizBackGround As String, ButtonType As String

Public QuizBackGroundEnable As String

Public TransState As String

'''''PROGRESS BAR CONTROL
Public Sub Update_ProgressBar()
        Dim Counter As Integer
        Dim Workarea(500) As String
        frmSplash.ProgressBar1.Min = LBound(Workarea)
        frmSplash.ProgressBar1.Max = UBound(Workarea)
        frmSplash.ProgressBar1.Visible = True

        'Set the Progress's Value to Min.
        frmSplash.ProgressBar1.Value = frmSplash.ProgressBar1.Min

        'Loop through the array.
        For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        frmSplash.ProgressBar1.Value = Counter

        Next Counter
        'ProgressBar1.Visible = False
        frmSplash.ProgressBar1.Value = frmSplash.ProgressBar1.Min
''MsgBox appname
End Sub


Public Sub Form_Font_White()

    '''ENABLE TRANSPARENT CONTROLS ON FORM
    For X = 1 To 4
        frmMain.lblAns(X).ForeColor = vbWhite
    Next X
    
    frmMain.lblQuestion.ForeColor = vbWhite
    
    
    frmMain.lblQuestNum.ForeColor = vbWhite
    
    
    frmMain.lblAnswerStat.ForeColor = vbWhite
    
    
    frmMain.lblStatus.ForeColor = vbWhite
    
    frmMain.lblRandomQuests.ForeColor = vbWhite
    
    frmMain.lblTimer.ForeColor = vbWhite

End Sub


Public Sub Form_Font_Black()

    '''ENABLE TRANSPARENT CONTROLS ON FORM
    For X = 1 To 4
        frmMain.lblAns(X).ForeColor = vbBlack
    Next X
    
    frmMain.lblQuestion.ForeColor = vbBlack
    
    
    frmMain.lblQuestNum.ForeColor = vbBlack
    
    
    frmMain.lblAnswerStat.ForeColor = vbBlack
    
    
    frmMain.lblStatus.ForeColor = vbBlack
    
    frmMain.lblRandomQuests.ForeColor = vbBlack
    
    frmMain.lblTimer.ForeColor = vbBlack

End Sub


Public Sub Form_Trans_Controls()

    '''ENABLE TRANSPARENT CONTROLS ON FORM
    For X = 1 To 4
        frmMain.lblAns(X).BackStyle = 0
    Next X
    
    frmMain.lblQuestion.BackStyle = 0
    
    
    frmMain.lblQuestNum.BackStyle = 0
    frmMain.lblQuestNum.BorderStyle = 0
    
    
    frmMain.lblAnswerStat.BackStyle = 0
    frmMain.lblAnswerStat.BorderStyle = 0
    
    
    frmMain.lblStatus.BackStyle = 0
    frmMain.lblStatus.BorderStyle = 0
    
    frmMain.lblRandomQuests.BackStyle = 0
    frmMain.lblRandomQuests.BorderStyle = 0
    
    frmMain.lblTimer.BackStyle = 0
    frmMain.lblTimer.BorderStyle = 0
    
    TransState = "True"
    SaveSetting appname, "Data", "TransState", TransState


End Sub

Public Sub Form_Trans_Controls_Disable()

    '''ENABLE TRANSPARENT CONTROLS ON FORM
    For X = 1 To 4
        frmMain.lblAns(X).BackStyle = 1
    Next X
    
    frmMain.lblQuestion.BackStyle = 1
    
    
    frmMain.lblQuestNum.BackStyle = 1
    frmMain.lblQuestNum.BorderStyle = 1
    
    
    frmMain.lblAnswerStat.BackStyle = 1
    frmMain.lblAnswerStat.BorderStyle = 1
    
    
    frmMain.lblStatus.BackStyle = 1
    frmMain.lblStatus.BorderStyle = 1
    
    frmMain.lblRandomQuests.BackStyle = 1
    frmMain.lblRandomQuests.BorderStyle = 1
    
    frmMain.lblTimer.BackStyle = 1
    frmMain.lblTimer.BorderStyle = 1
    
    TransState = "False"
    SaveSetting appname, "Data", "TransState", TransState


End Sub

