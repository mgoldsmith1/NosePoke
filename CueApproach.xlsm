Dim CueTime() As Single
Dim FoodRewardTime() As Single
Dim NosepokeTime() As Single
Dim WheelTurnTime() As Single
Dim LeverPressTime() As Single
Dim Address() As Variant

Sub FindNumberofFiles()
     
    'Erases any previously formated results
    'Worksheets(1).Range("H1")ClearContents
    
    NumofFiles = 0
    ReDim Address(100)

    Worksheets(1).Activate
    With Worksheets(1).Range("a1:a25500")
        
        Set c = .Find("MSN:", LookIn:=xlValues)
        If Not c Is Nothing Then
            Address(0) = c.Address
            Do
                c.Interior.Pattern = xlPatternGray50
                Set c = .FindNext(c)
                NumofFiles = NumofFiles + 1
                Address(NumofFiles) = c.Address
          
            Loop While Not c Is Nothing And c.Address <> Address(0)
        End If
    End With
    Debug.Print NumofFiles; Address(0); Address(1); Address(2)
    Call ChooseStudy(NumofFiles)
    
End Sub

Sub ChooseStudy(NumofFiles)
    ThisFile = 0
    Debug.Print NumofFiles; Address(0); Address(1); Address(2)
        Do
            Range(Address(ThisFile)).Select
            ProgramRan = ActiveCell.Offset(0, 1)
            
            If ProgramRan = "Cue" Or ProgramRan = "WISESA5" Then
                Call YokeInfoForCueApproach(ThisFile, NumofFiles, ProgramRan)
            End If
            If ProgramRan = "NEWYOK~1" Or ProgramRan = "GoodYok_yoked" Then
                Call Yokeratdata(ThisFile, NumofFiles, ProgramRan)
            End If
            If ProgramRan = "WISESA5" Or ProgramRan = "WISESA6" Or ProgramRan = "WISESA6" Then
                Call WISESA5(ThisFile, NumofFiles, ProgramRan)
            End If
            
            If ProgramRan = "DRUG_NEW_B1" Or ProgramRan = "DRUG_NEW_B3" Or ProgramRan = "DRUG_NEW_B5" Or ProgramRan = "DRUG_NEW_B7" Then
                Call WISESA5(ThisFile, NumofFiles, ProgramRan)
            End If
            If ProgramRan = "DRUGRAT" Or ProgramRan = "SORAYA" Then
                Call Yokeratdata(ThisFile, NumofFiles, ProgramRan)
            End If
            ThisFile = ThisFile + 1
        Loop While ThisFile < NumofFiles
         ActiveSheet.Columns("N:N").Insert Shift:=xlToRight
         ActiveSheet.Columns("U:U").Insert Shift:=xlToRight
         ActiveSheet.Columns("AB:AB").Insert Shift:=xlToRight
         ActiveSheet.Columns("AI:AI").Insert Shift:=xlToRight
         ActiveSheet.Columns("AP:AP").Insert Shift:=xlToRight
         ActiveSheet.Columns("AW:AW").Insert Shift:=xlToRight
         ActiveSheet.Columns("BD:BD").Insert Shift:=xlToRight
         ActiveSheet.Columns("BK:BK").Insert Shift:=xlToRight
         ActiveSheet.Columns("BR:BR").Insert Shift:=xlToRight
         ActiveSheet.Columns("BY:BY").Insert Shift:=xlToRight
         'Worksheets(1).Range("H1:AAA1").Columns.AutoFit
         Application.ScreenUpdating = True
         
End Sub

Sub YokeInfoForCueApproach(ThisFile, NumofFiles, ProgramRan)
    If ThisFile = NumofFiles - 1 Then
        EndOfFile = "A20000"
    Else
        Range(Address(ThisFile + 1)).Select
        EndOfFile = ActiveCell.Offset(-9, 0).Address
    End If
    
    Range(Address(ThisFile)).Select
    SearchArea = Address(ThisFile) + ":" + EndOfFile
    subnum = ActiveCell.Offset(-6, 1)
    testcell = ActiveCell.Offset(-8, 4)
    Debug.Print ActiveCell.Offset(-6, 1), ActiveCell.Offset(-8, 4)
    Debug.Print testcell
    If Val(testcell) = Empty Then 'Val(testcell) = 0 Then
        Filedate$ = ActiveCell.Offset(-8, 2) '+ Trim$(Str$(ActiveCell.Offset(-8, 3)))
    Else
        Filedate$ = ActiveCell.Offset(-8, 2) + ActiveCell.Offset(-8, 3) + Str$(ActiveCell.Offset(-8, 4))
   
    End If
    
  
   '------------------------------------------------------------
   '------------ SearchArea begins at the "MSN" offset loc: -----------
   '------------------------------------------------------------
    With Worksheets(1).Range(SearchArea)

    
        '----------------------------------------
        '------------ Total Actions: ---------------
        '----------------------------------------
    
        'Find Total Cue
        Set c = .Find("A:", LookIn:=xlValues)
        cFound = c.Address
        Range(c.Address).Select
        totNumCue = ActiveCell.Offset(0, 1)
        
        'Find total FoodReward / Pellet
        Set c = .Find("B:", LookIn:=xlValues)
        EFound = c.Address
        Range(c.Address).Select
        totFoodReward = ActiveCell.Offset(0, 1)
        
        'Find Total Nosepoke
        Set c = .Find("F:", LookIn:=xlValues)
        cFound = c.Address
        Range(c.Address).Select
        totNosepoke = ActiveCell.Offset(0, 1)
        
        'Find Total WheelTurn
        Set c = .Find("J:", LookIn:=xlValues)
        cFound = c.Address
        Range(c.Address).Select
        totWheelTurn = ActiveCell.Offset(0, 1)
        
        'Find Total Lever Press
        Set c = .Find("M:", LookIn:=xlValues)
        cFound = c.Address
        Range(c.Address).Select
        totLeverPress = ActiveCell.Offset(0, 1)
        
        '----------------------------------------
        '------------ Timestamps: ----------------
        '----------------------------------------
        
        'Find Total Time for Cue ---cumalative
        Set c = .Find("D:", LookIn:=xlValues)
        IFound = c.Address
        ReDim CueTime(totNumCue + 2)
        Range(c.Address).Select
        'Call gatherdata(totNumCue, CueTime)
        Row = 1
        Col = 2
        
        For x = 1 To totNumCue + 1
            CueTime(x) = ActiveCell.Offset(Row, Col) '(ActiveCell.Offset(Row, Col) / 10 / 60)
            
            If Col >= 7 Then
                Col = 2
                Row = Row + 1
                CueTime(x) = ActiveCell.Offset(Row, Col)
                
                If Col >= 1 And CueTime(x) >= 0 Then
                    Col = Col + 1
                End If
            
            Else
            
                If Col >= 0 And CueTime(x) > -1 Then
                    Col = Col + 1
                End If
            
            End If
        Next
        
        'Find Total Time for FoodReward Time / Pellet Time ---cumalative
        Set c = .Find("E:", LookIn:=xlValues)
        RFound = c.Address
        ReDim FoodRewardTime(totFoodReward + 2)
        Range(c.Address).Select
        Row = 1
        Col = 2
        For x = 1 To totFoodReward + 1
            FoodRewardTime(x) = ActiveCell.Offset(Row, Col) '(ActiveCell.Offset(Row, Col) / 10 / 60)
            If Col >= 7 Then
                Col = 2
                Row = Row + 1
                FoodRewardTime(x) = ActiveCell.Offset(Row, Col)
                
                If Col >= 1 And FoodRewardTime(x) >= 0 Then
                     Col = Col + 1
                End If
            
            Else
            
                If Col >= 0 And FoodRewardTime(x) > -1 Then
                    Col = Col + 1
                End If
            
            End If
        Next
        
        'Find Total Time for Nosepoke ---cumalative
        Set c = .Find("P:", LookIn:=xlValues)
        RFound = c.Address
        ReDim NosepokeTime(totNosepoke + 2)
        Range(c.Address).Select
        Row = 1
        Col = 2
        For x = 1 To totNosepoke + 1
            NosepokeTime(x) = ActiveCell.Offset(Row, Col) '(ActiveCell.Offset(Row, Col) / 10 / 60)
            
            If Col >= 7 Then
                Col = 2
                Row = Row + 1
                NosepokeTime(x) = ActiveCell.Offset(Row, Col)
                
                If Col >= 1 And NosepokeTime(x) >= 0 Then
                     Col = Col + 1
                End If
            
            Else
            
                If Col >= 0 And NosepokeTime(x) > -1 Then
                    Col = Col + 1
                End If
            
            End If
            
        Next
        
        'Find Total Time for WheelTurn --- cumalative
        Set c = .Find("L:", LookIn:=xlValues)
        RFound = c.Address
        ReDim WheelTurnTime(totWheelTurn + 2)
        Range(c.Address).Select
        Row = 1
        Col = 2
        For x = 1 To totWheelTurn + 1
           WheelTurnTime(x) = ActiveCell.Offset(Row, Col) ' (ActiveCell.Offset(Row, Col) / 10 / 60)
           If Col >= 7 Then
                Col = 2
                Row = Row + 1
                WheelTurnTime(x) = ActiveCell.Offset(Row, Col)
                
                If Col >= 1 And WheelTurnTime(x) >= 0 Then
                    Col = Col + 1
                End If
            
            Else
            
                If Col >= 0 And WheelTurnTime(x) > -1 Then
                    Col = Col + 1
                End If
            
            End If
        Next
        
        'Find Total Time for LeverPress --- cumalative
        Set c = .Find("N:", LookIn:=xlValues)
        RFound = c.Address
        ReDim LeverPressTime(totLeverPress + 2)
        Range(c.Address).Select
        Row = 1
        Col = 2
        For x = 1 To totLeverPress + 1
            LeverPressTime(x) = ActiveCell.Offset(Row, Col) '(ActiveCell.Offset(Row, Col) / 10 / 60)
            If Col >= 7 Then
                Col = 2
                Row = Row + 1
                LeverPressTime(x) = ActiveCell.Offset(Row, Col)
                
                If Col >= 1 And LeverPressTime(x) >= 0 Then
                    Col = Col + 1
                End If
            
            Else
            
                If Col >= 0 And LeverPressTime(x) > -1 Then   'WheelTurnTime(x) > -1 Then  <-- This could be the problem
                    Col = Col + 1
                End If
            
            End If
        Next
        
        
    End With
    'Print Data collected onto active worksheet
    'Print Totals--Inj,Active,Inactive
     'Print Data collected onto active worksheet
    'Print Totals--Inj,Active,Inactive
    'determine file nmber to know where to place data
    

    Range("H1").Select
    
    Call PrintTitlesForYokeInfoCueApproach(ThisFile, subnum, Filedate$, totNumCue, totFoodReward, totNosepoke, totWheelTurn, totLeverPress)
     
    'Print Time of Cues
    Range("I4").Select
    Title1$ = "Cue Time"
    Title2$ = "Cue Total"
    Call ExcelPrintYokeInfoForCueApproach(ThisFile, Title1$, Title2$, totNumCue, CueTime, ProgramRan)
    
    
    'Print Time of FoodReward / Pellets
    Range("J4").Select
    Title1$ = "FoodReward Time"
    Title2$ = "FoodReward Total"
    Call ExcelPrintYokeInfoForCueApproach(ThisFile, Title1$, Title2$, totFoodReward, FoodRewardTime, ProgramRan)
    
    Range("K4").Select
    Title1$ = "Nosepoke Time"
    Title2$ = "Nosepoke Total"
    Call ExcelPrintYokeInfoForCueApproach(ThisFile, Title1$, Title2$, totNosepoke, NosepokeTime, ProgramRan)
    
    Range("L4").Select
    Title1$ = "WheelTurn Time"
    Title2$ = "WheelTurn Total"
    Call ExcelPrintYokeInfoForCueApproach(ThisFile, Title1$, Title2$, totWheelTurn, WheelTurnTime, ProgramRan)
    
    Range("M4").Select
    Title1$ = "LeverPress Time"
    Title2$ = "LeverPress Total"
    Call ExcelPrintYokeInfoForCueApproach(ThisFile, Title1$, Title2$, totLeverPress, LeverPressTime, ProgramRan)
    
         
    'Print Time of InActive Responses
  '  Range("M4").Select
    'ActiveCell.Offset(0, ThisFile * 6).Select
    'ActiveCell.Offset(0, 0) = "InAct Time"
  '  If NONACTIVERESPCOLLECTED = False Then
 '       ActiveCell.Offset(0, ThisFile * 6).Select
'        ActiveCell.Offset(0, 0) = "InAct Time"
 '       ActiveCell.Offset(1, 0) = "Not Collected"
'    Else
          '  TitleA$ = "InAct Time"
         '   Call ExcelPrintInj(ThisFile, TitleA$, totinactive, InActiveTime, ProgramRan)
    'End If

  '  Row = 1
 '   Col = 2
  '  For x = 2 To 1000
        'the totnum of resp from d(1) does not
    '    'match up with time of responses. Fixed in Wisesa6
   '     If ActiveTime(x) = -1 Then
  '         Exit For
 '       End If
       ' CumalativeActTime = ActiveTime(x) ' + CumalativeActTime
   '     ActiveCell.Offset(x, Col) = "-" 'ActiveTime(x)  'CumalativeActTime

 End Sub

Sub Yokeratdata(ThisFile, NumofFiles, ProgramRan)
    If ThisFile = NumofFiles - 1 Then
        EndOfFile = "A20000"
    Else
        Range(Address(ThisFile + 1)).Select
        EndOfFile = ActiveCell.Offset(-9, 0).Address
    End If
    
    Range(Address(ThisFile)).Select
    SearchArea = Address(ThisFile) + ":" + EndOfFile
    subnum = ActiveCell.Offset(-6, 1)
    testcell = ActiveCell.Offset(-8, 4)
    Debug.Print ActiveCell.Offset(-6, 1), ActiveCell.Offset(-8, 4)
    Debug.Print testcell
    If Val(testcell) = 0 Then
        Filedate$ = ActiveCell.Offset(-8, 2) + Trim$(Str$(ActiveCell.Offset(-8, 3)))
    Else
        Filedate$ = ActiveCell.Offset(-8, 2) + ActiveCell.Offset(-8, 3) + Str$(ActiveCell.Offset(-8, 4))
    End If
    

    With Worksheets(1).Range(SearchArea)
        'Find Total Injection
        Set c = .Find("C:", LookIn:=xlValues)
        cFound = c.Address
        Range(c.Address).Select
        TotNumInj = ActiveCell.Offset(1, 2)
        
        'Find total active presses
        Set c = .Find("D:", LookIn:=xlValues)
        EFound = c.Address
        Range(c.Address).Select
        Totactive = ActiveCell.Offset(1, 2)
        
        'Find total Inactive presses
        Set c = .Find("E:", LookIn:=xlValues)
        EFound = c.Address
        Range(c.Address).Select
        totinactive = ActiveCell.Offset(1, 2)
        Debug.Print totinactive
        
        'Find Total Time for Injections---cumalative
        Set c = .Find("I:", LookIn:=xlValues)
        IFound = c.Address
        ReDim InjTime(TotNumInj)
        Range(c.Address).Select
        'Call gatherdata(TotNumInj, InjTime)
        Row = 1
        Col = 2
        For x = 1 To TotNumInj
            InjTime(x) = ActiveCell.Offset(Row, Col)
        If Col >= 6 Then
                Col = 2
                Row = Row + 1
            Else
                Col = Col + 1
            End If
        Next
        
        
    'Find Total Time for Active Responses---cumalative
        Set c = .Find("R:", LookIn:=xlValues)
        RFound = c.Address
        ReDim ActiveTime(Totactive)
        Range(c.Address).Select
        Row = 1
        Col = 2
        For x = 1 To Totactive
            ActiveTime(x) = (ActiveCell.Offset(Row, Col) / 10 / 60)
            If Col >= 6 Then
                Col = 2
                Row = Row + 1
            Else
                Col = Col + 1
            End If
        Next
        
        'Find Total Time for InActive Responses---cumalative
        Set c = .Find("S:", LookIn:=xlValues)
        SFound = c.Address
        ReDim InActiveTime(totinactive)
        Range(c.Address).Select
        If ActiveCell.Offset(1, 0) = "T:" Then
        
            Dim Msg, Style, Title, Response, MyString
            Msg = "Non Active resp. Not compiled in MPC"    ' Define message.
            Style = vbOKOnly    ' Define buttons.
            Title = "Non-Active resp info"
            Response = MsgBox(Msg, Style, Title)
            NONACTIVERESPCOLLECTED = False
        Else
            Row = 1
            Col = 2
            For x = 1 To totinactive
                InActiveTime(x) = ActiveCell.Offset(Row, Col) / 10 / 60
                If Col >= 6 Then
                    Col = 2
                    Row = Row + 1
                Else
                    Col = Col + 1
                End If
            Next
            NONACTIVERESPCOLLECTED = True
        End If
        
    End With
    'Print Data collected onto active worksheet
    'Print Totals--Inj,Active,Inactive
     'Print Data collected onto active worksheet
    'Print Totals--Inj,Active,Inactive
    'determine file nmber to know where to place data
    Range("O1").Select
    Call PrintTitles(ThisFile, subnum, Filedate$, TotNumInj, Totactive, totinactive)
   
    'Print Time of Injections
    Range("P4").Select
    TitleA$ = "Inj Time"
    Call ExcelPrintInj(ThisFile, TitleA$, TotNumInj, InjTime, ProgramRan)
     
    'Print Time of Active Responses
    Range("Q4").Select
    Title1$ = "Act Time"
    Title2$ = "Act IRT"
    Call ExcelPrintAct(ThisFile, Title1$, Title2$, Totactive, ActiveTime, ProgramRan)
         
    'Print Time of InActive Responses
    Range("R4").Select
    'ActiveCell.Offset(0, ThisFile * 6).Select
    'ActiveCell.Offset(0, 0) = "InAct Time"
    If NONACTIVERESPCOLLECTED = False Then
        ActiveCell.Offset(0, ThisFile * 6).Select
        ActiveCell.Offset(0, 0) = "InAct Time"
        ActiveCell.Offset(1, 0) = "Not Collected"
    Else
            TitleA$ = "InAct Time"
            Call ExcelPrintInj(ThisFile, TitleA$, totinactive, InActiveTime, ProgramRan)
    End If
 End Sub

Sub WISESA5(ThisFile, NumofFiles, ProgramRan)
    If ThisFile = NumofFiles - 1 Then
        EndOfFile = "A20000"
    Else
        Range(Address(ThisFile + 1)).Select
        EndOfFile = ActiveCell.Offset(-9, 0).Address
    End If
    Range(Address(ThisFile)).Select
    SearchArea = Address(ThisFile) + ":" + EndOfFile
    subnum = ActiveCell.Offset(-6, 1)
     
     If ActiveCell.Offset(-8, 4) <> Empty Then
        Filedate$ = (ActiveCell.Offset(-8, 2)) + (ActiveCell.Offset(-8, 3)) + Str$(ActiveCell.Offset(-8, 4))
     Else
        Filedate$ = (ActiveCell.Offset(-8, 2)) + Str$(ActiveCell.Offset(-8, 3))
    End If
        
     With Worksheets(1).Range(SearchArea)
        'Find Total Injection
        Set c = .Find("C:", LookIn:=xlValues)
        cFound = c.Address
        Range(c.Address).Select
        TotNumInj = ActiveCell.Offset(1, 2)
        
        'Find total active presses
        Set c = .Find("D:", LookIn:=xlValues)
        EFound = c.Address
        Range(c.Address).Select
        Totactive = ActiveCell.Offset(1, 2)
        
        'Find total Inactive presses
        Set c = .Find("E:", LookIn:=xlValues)
        EFound = c.Address
        Range(c.Address).Select
        totinactive = ActiveCell.Offset(1, 2)
        Debug.Print totinactive
        
        'Find Total Time for Injections---cumalative
        'For some reason the time is kept in seconds for rft,
        'tenths of seconds for Active,Inactive IRTs. In both cases
        'the timer is then divided by 60. Before I print I'm
        'going to multiply by 60 to get in back into seconds of 10th of seconds
        
        Set c = .Find("H:", LookIn:=xlValues)
        IFound = c.Address
        ReDim InjTime(TotNumInj)
        Range(c.Address).Select
        Row = 1
        Col = 2
        For x = 1 To TotNumInj
            InjTime(x) = ActiveCell.Offset(Row, Col)
            If Col >= 6 Then
                Col = 2
                Row = Row + 1
            Else
                Col = Col + 1
            End If
        Next
        
        
    'Find Total Time for Active Responses---cumalative
        Set c = .Find("I:", LookIn:=xlValues)
        RFound = c.Address
        ReDim ActiveTime(Totactive)
        Range(c.Address).Select
        Row = 1
        Col = 2
        For x = 1 To Totactive
            ActiveTime(x) = ActiveCell.Offset(Row, Col)
            If Col >= 6 Then
                Col = 2
                Row = Row + 1
            Else
                Col = Col + 1
            End If
        Next
        
        'Find Total Time for InActive Responses---cumalative
        Set c = .Find("N:", LookIn:=xlValues)
        SFound = c.Address
        ReDim InActiveTime(totinactive)
        Range(c.Address).Select
        Row = 1
        Col = 2
        For x = 1 To totinactive
            InActiveTime(x) = ActiveCell.Offset(Row, Col)
            If Col >= 6 Then
                Col = 2
                Row = Row + 1
            Else
                Col = Col + 1
            End If
        Next
    End With
    'Print Data collected onto active worksheet
    'Print Totals--Inj,Active,Inactive
    'determine file nmber to know where to place data
    
    Range("S1").Select
    Call PrintTitles(ThisFile, subnum, Filedate$, TotNumInj, Totactive, totinactive)
   
    
    'Print Time of Injections
    Range("T4").Select
    Title$ = "Inj Time"
    Call ExcelPrintInj(ThisFile, Title$, TotNumInj, InjTime, ProgramRan)
   
   'Print Time of Active Responses
    Range("U4").Select
    Title1$ = "Act Time"
    Title2$ = "Act IRT"
    Call ExcelPrintAct(ThisFile, Title1$, Title2$, Totactive, ActiveTime, PROGRAMRUN)
    
    'Print Time of InActive Responses
    Range("V4").Select
     Title1$ = "InAct Time"
     Title2$ = "InAct IRT"
    Call ExcelPrintInAct(ThisFile, Title1$, Title2$, totinactive, InActiveTime, PROGRAMRUN)
    
    
End Sub

Sub EventLine()
    If ActiveCell.Value <> "Sub#" Then
            Dim Msg, Style, Title, Response, MyString
            Msg = "Place active cell on  sub#  title cell"    ' Define message.
            Style = vbOKOnly    ' Define buttons.
            Title = "Can't produce recording"
            Response = MsgBox(Msg, Style, Title)
            Exit Sub
    End If
    ActiveCell.Offset(1, 0).Select
    AnimalNum$ = ActiveCell
    'Filedate has slashes.Have to get rid of them
     Filedate$ = ActiveCell.Offset(1, 1)
    Searchstring = Filedate$
    SearchChar = "/"
    Mypos = 1

    Mypos = InStr(1, Searchstring, SearchChar, 1)
    FileMinusMonth$ = Right(Filedate$, Len(Searchstring) - Mypos)
    MonthwithSlash$ = Left(Searchstring, Mypos)
    Mypos = InStr(1, MonthwithSlash$, SearchChar, 1)
    FileMonth$ = Left(MonthwithSlash$, Mypos - 1)
    Mypos = InStr(1, FileMinusMonth$, SearchChar, 1)
    FileDay$ = Left(FileMinusMonth$, Mypos - 1)
    FileYear$ = Right(Searchstring, 4)
   Filedate$ = FileMonth + "-" + FileDay$ + "-" + FileYear$
   
   
    'Gatherdata for Recorder
    'Timenfo(1,X) = injectionTimes
    'Timenfo(2,X) = Act. lever time
    'Timenfo(3,X) = InAct. lever times
    Dim ViInfo(3, 501)
    Dim TimeInfo(3, 501)
    ActiveCell.Offset(3, 0).Select
    thistime = ActiveCell
    x = 0
    Do While thistime > 0
        TimeInfo(1, x) = thistime * 60
        x = x + 1
        thistime = ActiveCell.Offset(x, 0)
    Loop
    
    ActiveCell.Offset(0, 1).Select
    thistime = ActiveCell
    x = 0
    Do While thistime > 0
        TimeInfo(2, x) = thistime * 60
        x = x + 1
        thistime = ActiveCell.Offset(x, 0)
    Loop
    
    ActiveCell.Offset(0, 1).Select
    thistime = ActiveCell
    x = 0
    If thistime = "Not Collected" Then
    '......
    Else
    
        Do While thistime > 0
            TimeInfo(3, x) = thistime * 60
            x = x + 1
            thistime = ActiveCell.Offset(x, 0)
        Loop
        End If
    ExcelWork$ = Filedate$ + "(" + Trim$(AnimalNum$) + ")" + ".xls"
    Set newBook = Workbooks.Add
    With newBook
        .Title = AnimalNum$ + "Event"
        .Subject = "Sales"
        .SaveAs Filename:=ExcelWork$
    End With
    Worksheets(1).Name = "Recorder"
    Set MyDoc = Workbooks(ExcelWork$).Worksheets("Recorder")
    MyDoc.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.SplitRow = 6.84615384615385
    ActiveWindow.FreezePanes = True
    ActiveWindow.TabRatio = 0.697
    
    
    
    Dim NumOfinjPerTime(30)
    IntervalTime = 60
    NumPixRsp = 1
    Timerdivider = 42
    'Timerdivider = 10
    IntervalTime = 60
    ActiveWindow.TabRatio = 0.697
    NumPixRft = 4
    Respreset = 500
    ResetYnum = NumPixRsp * Respreset
    StartX = 30
    EndX = 30
    OriginalX = StartX
    StartY = 600 * NumPixRsp
    EndY = 600 * NumPixRsp
    OriginalY = 600 * NumPixRsp
    
    StartIncorrY = StartY
    
    'Incorrect responses
    starteventoneY = OriginalY + 60
    StarteventOneX = OriginalX
    EndeventOneX = 0
    
    StarteventRftX = OriginalX
    StarteventRftY = starteventoneY + 15
    
    

        
    StarteventtwoY = StarteventRftY + 15
    StarteventtwoX = OriginalX
    EndeventTwoX = 0
    StarteventCorrX = OriginalX
    StarteventCorrY = StarteventRftY + 15
                
    
    
    x = 0
    RespType = 1
    Do
         If TimeInfo(2, x) = 0 Then
            Exit Do
         End If
         Timers = TimeInfo(2, x)
         Debug.Print x, Timers
          If Val(Timers) > 0 Then
            EndX = Timers / Timerdivider + OriginalX
            EndeventTwoX = Timers / Timerdivider + OriginalX
            TotTime = Val(Timers)
            Call drawLine(MyDoc, StartX, EndX, StartY, EndY, OriginalY, 0, 0, 225)
            StartX = EndX
        End If
        TotRsp = TotRsp + 1
        LastRespTime = Val(Timers)
        EndY = StartY - NumPixRsp
        'Call drawLine(MyDoc, StartX, EndX, StartY, EndY, OriginalY, 0, 0, 225)
        TotResp = 1 + TotResp
        StartY = EndY
        EndeventCorrX = EndeventTwoX
        EndeventCorrY = StarteventCorrY
        Call drawLine(MyDoc, StarteventCorrX, EndeventCorrX, StarteventCorrY, EndeventCorrY, OriginalY, 0, 0, 225)
        StarteventCorrX = EndeventCorrX
        EndeventCorrY = StarteventCorrY + NumPixRft
        Call drawLine(MyDoc, StarteventCorrX, EndeventCorrX, StarteventCorrY, EndeventCorrY, OriginalY, 0, 0, 225)
        EndeventCorrY = StarteventCorrY
        x = x + 1
    Loop
    
    x = 0
    RespType = 2
    Do
        If TimeInfo(1, x) = 0 Then
            Exit Do
         End If
         Timers = TimeInfo(1, x)
         Debug.Print x, Timers
          If Val(Timers) > 0 Then
            EndX = Timers / Timerdivider + OriginalX
            EndeventFiveX = Timers / Timerdivider + OriginalX
            TotTime = Val(Timers)
            'Call drawLine(MyDoc, StartX, EndX, StartY, EndY, OriginalY, 0, 225, 0)
            StartX = EndX
        End If
        'rft has occured
        EndY = StartY + NumPixRft
        EndX = StartX + NumPixRft
            'Call drawLine(MyDoc, StartX, EndX, StartY, EndY, OriginalY, 250, 0, 0)
            'for rft blip I do not want the coordinates to cahneg so I set them back to startx,y numbers
        EndY = StartY
        EndX = StartX
        TotRft = TotRft + 1
             'Determine the number of injections that occur in a given timeframe.(i.e. # of injections per hour)
            'Array NumOfinjPerTime contains injections By time,Var IntervalTime will be the alotted time to look at.
            'Debug.Print TimerForGroupInj, IntervalTime
            'InjIndex = Fix(Val((TimerForGroupInj / 60) / IntervalTime))
            'Debug.Print InjIndex
            'NumOfinjPerTime(InjIndex) = NumOfinjPerTime(InjIndex) + 1
            'Draw rft event Line
            'StarteventRftY
        EndeventRftX = EndeventFiveX
        EndeventRftY = StarteventRftY
        Call drawLine(MyDoc, StarteventRftX, EndeventRftX, StarteventRftY, EndeventRftY, OriginalY, 250, 0, 0)
        StarteventRftX = EndeventRftX
        EndeventRftY = StarteventRftY + NumPixRft
        Call drawLine(MyDoc, StarteventRftX, EndeventRftX, StarteventRftY, EndeventRftY, OriginalY, 250, 0, 0)
        EndeventRftY = StarteventRftY
        x = x + 1
    Loop
        
        
     x = 0
    RespType = 51
    Do
        If TimeInfo(3, x) = 0 Then
            Exit Do
         End If
         Timers = TimeInfo(3, x)
         Debug.Print x, Timers
          If Val(Timers) > 0 Then
            EndX = Timers / Timerdivider + OriginalX
            EndeventOneX = Timers / Timerdivider + OriginalX
            TotTime = Val(Timers)
            'Call drawLine(MyDoc, StartX, EndX, StartY, EndY, OriginalY, 0, 225, 0)
            StartX = EndX
        End If
        'Incorrect Resp EventPen 51
        EndIncorrY = StartIncorrY - NumPixRsp
        Inactive = Inactive + 1
        endeventoneY = starteventoneY
        EndeventOneX = Val(Timers) / Timerdivider + OriginalX
        Call drawLine(MyDoc, StarteventOneX, EndeventOneX, starteventoneY, endeventoneY, OriginalY, 0, 225, 0)
        StartIncorrY = EndIncorrY
        Call drawLine(MyDoc, StarteventOneX, EndeventOneX, starteventoneY, endeventoneY, OriginalY, 0, 225, 0)
        StarteventOneX = EndeventOneX
        endeventoneY = starteventoneY + NumPixRft
        Call drawLine(MyDoc, StarteventOneX, EndeventOneX, starteventoneY, endeventoneY, OriginalY, 0, 225, 0)
        endeventoneY = starteventoneY
         x = x + 1
    Loop
               
              
    'Draw TimeLine for recorder
    'Below Draws a line
    
   With MyDoc.Shapes.AddLine(OriginalX, OriginalY + 30, EndX, OriginalY + 30).Line
       .ForeColor.RGB = RGB(redd, greenn, Bluee)
   End With
    EachPixTime = 1 * Timerdivider
    NumOfPixPerHalfHourInSecs = (30 * 60) / EachPixTime
    'Finish lines to end of session
    Call drawLine(MyDoc, OriginalX, (NumOfPixPerHalfHourInSecs * 12 + OriginalX), OriginalY + 30, OriginalY + 30, OriginalY + 30, 10, 50, 10)
    Call drawLine(MyDoc, OriginalX, (NumOfPixPerHalfHourInSecs * 12 + OriginalX), starteventoneY, endeventoneY, OriginalY, 0, 225, 0)
    Call drawLine(MyDoc, StarteventCorrX, (NumOfPixPerHalfHourInSecs * 12 + OriginalX), StarteventCorrY, EndeventCorrY, OriginalY, 0, 0, 225)
    Call drawLine(MyDoc, StarteventRftX, (NumOfPixPerHalfHourInSecs * 12 + OriginalX), StarteventRftY, EndeventRftY, OriginalY, 225, 0, 0)
    For x = 1 To 12
        SlashX = OriginalX + (NumOfPixPerHalfHourInSecs * x)
        With MyDoc.Shapes.AddLine(SlashX, OriginalY + 34, SlashX, OriginalY + 26).Line
                    .ForeColor.RGB = RGB(redd, greenn, Bluee)
        End With
        Application.CommandBars("Drawing").Visible = True
       Call drawLabels(ExcelWork$, x, SlashX, OriginalY)
    Next x
    ActiveWindow.DisplayGridlines = False
    ActiveCell.FormulaR1C1 = "Date"
    ActiveCell.Offset(1, 0) = "Tot Rfts"
    ActiveCell.Offset(0, 1) = Filedate$
    ActiveCell.Offset(2, 0) = "Tot Rsp"
    'ActiveCell.Offset(3, 0) = "Rate(Rsp/min)"
    ActiveCell.Offset(4, 0) = "Tot Inactive"
    ActiveCell.Offset(2, 1) = TotRsp
    ActiveCell.Offset(1, 1) = TotRft
    ActiveCell.Offset(4, 1) = Inactive
     'If TotResp > 0 Then
     '   ActiveCell.Offset(3, 1) = Format(TotResp / ((TotTime / 10) / 60), "###.###")
     '   TotRate = TotResp / (TotTime / 10)
    'Else
    '    ActiveCell.Offset(3, 1) = 0
    '    TotRate = 0
    'End If
     'ActiveCell.Offset(5, 0) = "Time"
     'ActiveCell.Offset(5, 1) = Format((TotTime / 10) / 60, "###.#")
     'ActiveCell.Offset(6, 0) = "Vi"
     'ActiveCell.Offset(6, 1) = ""
     'ActiveCell.Offset(1, 2) = "% CRF"
     'ActiveCell.Offset(1, 3) = Format(ProbCRF / 100, "###.#")
     ActiveCell.Offset(7, 3) = AnimalNumber$
     'Print Resp,vi dur etc onto sheet
    'ActiveCell.Offset(1, 4) = "Vi Rsp"
   ' ActiveCell.Offset(2, 4) = "Vi Dur"
   ' ActiveCell.Offset(3, 4) = "Ex Vi Rsp"
   ' ActiveCell.Offset(4, 4) = "Ex Vi Dur"
   ' For x = 1 To ExVi
   '     ActiveCell.Offset(1, 5 + x) = ViInfo(0, x)
   '     ActiveCell.Offset(2, 5 + x) = ViInfo(4, x) / 10
   '     ActiveCell.Offset(3, 5 + x) = ViInfo(1, x)
   '    ActiveCell.Offset(4, 5 + x) = ViInfo(5, x) / 10
   ' Next x
    'MyDoc.Range("AE11:AW11").Name = "DatToCopy"
    Range("A1:A7").Select
        With Selection
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .ShrinkToFit = False
            .MergeCells = False
        End With
        Range("B1:B7").Select
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .IndentLevel = 0
            .ShrinkToFit = False
            .MergeCells = False
        End With
       
         
        Range("E1:E5").Select
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .IndentLevel = 0
            .ShrinkToFit = False
            .MergeCells = False
        End With
        If NumPixRsp = 1 Then
            MyDoc.Range("A45").Name = "CR"
        ElseIf NumPixRsp = 2 Then
            MyDoc.Range("A86").Name = "CR"
        ElseIf NumPixRsp = 3 Then
            MyDoc.Range("A127").Name = "CR"
        ElseIf NumPixRsp = 4 Then
            MyDoc.Range("A168").Name = "CR"
        ElseIf NumPixRsp = 5 Then
            MyDoc.Range("A209").Name = "CR"
        End If
        whichone$ = "InActive Resp"
        yoffset = 300
        Call drawLegend(MyDoc, OriginalX, OriginalX + 80, OriginalY - 295, OriginalY - 295, OriginalY, 0, 225, 0, whichone$, yoffset)
        
        whichone$ = "RFT"
        yoffset = 275
        Call drawLegend(MyDoc, OriginalX, OriginalX + 80, OriginalY - 270, OriginalY - 270, OriginalY, 225, 0, 0, whichone$, yoffset)
        
        whichone$ = "Active Resp"
        yoffset = 250
        Call drawLegend(MyDoc, OriginalX, OriginalX + 80, OriginalY - 245, OriginalY - 245, OriginalY, 0, 0, 225, whichone$, yoffset)
        
        Application.CommandBars("Drawing").Visible = False
End Sub
Sub drawLine(MyDoc, StartX, EndX, StartY, EndY, OriginalY, redd, greenn, Bluee)
    'Below Draws a line
    With MyDoc.Shapes.AddLine(StartX, StartY, EndX, EndY).Line
           .ForeColor.RGB = RGB(redd, greenn, Bluee)
    End With

End Sub
Sub drawLabels(ExcelWork$, x, SlashX, OriginalY)
'
' draw Macro
' Macro recorded 3/8/00 by Eric Thorndike
     Windows(ExcelWork$).Activate
    ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, (SlashX - 10), OriginalY + 40, 0# _
        , 0#).Select
    Selection.ShapeRange(1).TextFrame.AutoSize = msoTrue
    Selection.Characters.Text = Str$(x * 30)
    With Selection.Characters(Start:=1, Length:=2).Font
        .Name = "Geneva"
        .FontStyle = "Regular"
        .Size = 9
        .StrikeThrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
End Sub


Sub drawLegend(MyDoc, StartX, EndX, StartY, EndY, OriginalY, redd, greenn, Bluee, whichone$, yoffset)
    'Below Draws a line
    With MyDoc.Shapes.AddLine(StartX, StartY, EndX, EndY).Line
           .ForeColor.RGB = RGB(redd, greenn, Bluee)
    End With

ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, (OriginalX + 110), OriginalY - (yoffset), 0# _
        , 0#).Select
    Selection.ShapeRange(1).TextFrame.AutoSize = msoTrue
    Selection.Characters.Text = whichone$
    With Selection.Characters(Start:=1, Length:=2).Font
        .Name = "Geneva"
        .FontStyle = "Regular"
        .Size = 9
        .StrikeThrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With


End Sub

Sub ExcelPrintInj(ThisFile, Title$, TotNum, InjTime, ProgramRan)

    ActiveCell.Offset(0, ThisFile * 6).Select
    ActiveCell.Offset(0, 0) = Title$
    
    Row = 1
    Col = 0
    
    For x = 1 To TotNum
        If ProgramRan = "WISESA6" Then
            CumalativeInjTime = ((InjTime(x) / 10) * 60) + CumalativeInjTime
      '  ElseIf ProgramRan = "Cue" Then
       '     CumalativeInjTime = ((InjTime(x) / 10) * 60) + CumalativeInjTime
        ElseIf ProgramRan = "DRUGRAT" Then
            CumalativeInjTime = (InjTime(x) * 60) + CumalativeInjTime
        Else
            CumalativeInjTime = InjTime(x) + CumalativeInjTime
        End If
        ActiveCell.Offset(Row, Col) = CumalativeInjTime
        ActiveCell.Offset(Row, Col).NumberFormat = "0.00"
        Row = Row + 1
    Next


End Sub
Sub ExcelPrintAct(ThisFile, Title1$, Title2$, TotNum, ActiveTime, PROGRAMRUN)
    ActiveCell.Offset(0, ThisFile * 6).Select
    ActiveCell.Offset(0, 0) = Title1$
    ActiveCell.Offset(0, 2) = Title2$
    Row = 1
    Col = 0
    For x = 1 To TotNum
        'the totnum of resp from d(1) does not
    '    'match up with time of responses. Fixed in Wisesa6
        If ActiveTime(x) = -1 Then
           Exit For
        End If
        CumalativeActTime = ActiveTime(x) + CumalativeActTime
        ActiveCell.Offset(Row, Col) = CumalativeActTime
        ActiveCell.Offset(Row, Col).NumberFormat = "0.00"
        ActiveCell.Offset(Row, Col + 2) = ActiveTime(x)
        ActiveCell.Offset(Row, Col + 2).NumberFormat = "0.00"
        Row = Row + 1
    Next
 
End Sub

Sub ExcelPrintYokeInfoForCueApproach(ThisFile, Title1$, Title2$, TotNum, ActiveTime, PROGRAMRUN)

    ActiveCell.Offset(0, ThisFile * 6).Select
    ActiveCell.Offset(0, 0) = Title1$
    'ActiveCell.Offset(0, 2) = Title2$
  
    Row = 1
    Col = 0
    For x = 1 To TotNum + 1
        'the totnum of resp from d(1) does not
    '    'match up with time of responses. Fixed in Wisesa6
        If ActiveTime(x) = -1 Then
           Exit For
        End If
       ' CumalativeActTime = ActiveTime(x) ' + CumalativeActTime
        ActiveCell.Offset(Row, Col) = ActiveTime(x)  'CumalativeActTime
        ActiveCell.Offset(Row, Col).NumberFormat = "0.00"
        
        
        ' CumalativeActTime = ActiveTime(x) + CumalativeActTime
       ' ActiveCell.Offset(Row, Col) = CumalativeActTime
       ' ActiveCell.Offset(Row, Col).NumberFormat = "0.00"
      '  ActiveCell.Offset(Row, Col + 2) = ActiveTime(x)
      '  ActiveCell.Offset(Row, Col + 2).NumberFormat = "0.00"
        Row = Row + 1
    Next
 
End Sub
Sub PrintTitlesForYokeInfoCueApproach(ThisFile, subnum, Filedate$, totNumCue, totFoodReward, totNosepoke, totWheelTurn, totLeverPress)

    ActiveCell.Offset(0, ThisFile * 6).Select
    ActiveCell.Offset(0, 0) = "Sub#"
    ActiveCell.Offset(1, 0) = subnum
    ActiveCell.Offset(2, 0) = "Date"
    ActiveCell.Offset(2, 1) = Filedate$
    
    ActiveCell.Offset(0, 1) = "Tot NumCue"
    ActiveCell.Offset(1, 1) = totNumCue
    ActiveCell.Offset(0, 2) = "Tot FoodReward"
    ActiveCell.Offset(1, 2) = totFoodReward
    ActiveCell.Offset(0, 3) = "Tot Nosepoke"
    ActiveCell.Offset(1, 3) = totNosepoke
    ActiveCell.Offset(0, 4) = "Tot Wheel Turn"
    ActiveCell.Offset(1, 4) = totWheelTurn
    ActiveCell.Offset(0, 5) = "Tot LeverPress"
    ActiveCell.Offset(1, 5) = totLeverPress
    
End Sub
Sub PrintTitles(ThisFile, subnum, Filedate$, TotNumInj, Totactive, totinactive)


    ActiveCell.Offset(0, ThisFile * 6).Select
    ActiveCell.Offset(0, 0) = "Sub#"
    ActiveCell.Offset(1, 0) = subnum
    ActiveCell.Offset(2, 0) = "Date"
    ActiveCell.Offset(2, 1) = Filedate$
    
    ActiveCell.Offset(0, 1) = "Tot Inj"
    ActiveCell.Offset(1, 1) = TotNumInj
    ActiveCell.Offset(0, 2) = "Tot Ac"
    ActiveCell.Offset(1, 2) = Totactive
    ActiveCell.Offset(0, 3) = "Tot InAc"
    ActiveCell.Offset(1, 3) = totinactive
End Sub

Sub gatherdata(TotNum, InjTime)
        Row = 1
        Col = 2
        For x = 1 To TotNum
            InjTime(x) = ActiveCell.Offset(Row, Col)
            If Col >= 6 Then
                Col = 2
                Row = Row + 1
            Else
                Col = Col + 1
            End If
        Next
End Sub
Sub ExcelPrintInAct(ThisFile, Title1$, Title2$, TotNum, ActiveTime, PROGRAMRUN)
    ActiveCell.Offset(0, ThisFile * 6).Select
    ActiveCell.Offset(0, 0) = Title1$
    ActiveCell.Offset(0, 2) = Title2$
    Row = 1
    Col = 0
    For x = 1 To TotNum
        'the totnum of resp from d(1) does not
    '    'match up with time of responses. Fixed in Wisesa6
        If ActiveTime(x) = -1 Then
           Exit For
        End If
        CumalativeActTime = ActiveTime(x) + CumalativeActTime
        ActiveCell.Offset(Row, Col) = CumalativeActTime
        ActiveCell.Offset(Row, Col).NumberFormat = "0.00"
        ActiveCell.Offset(Row, Col + 2) = ActiveTime(x)
        ActiveCell.Offset(Row, Col + 2).NumberFormat = "0.00"
        Row = Row + 1
    Next
 
End Sub








