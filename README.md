Option Explicit

Sub Funnel_Model()

' This is the main macro that will call all other macros.
' Whichever sheet you're on, just run this macro and the appropriate macro will be triggered.
    ' If you're on a sheet that shouldn't have a macro, then nothing will be triggered!

' This macro will...
' firstly assign a number of variables that are almost always used in all other macros. These include...
'   -   Assign the current and Assumptions worksheets.
'   -   Assign the start time and month serived from the Assumptions tab (used in some calculations)
'   -   Assign the used range.
'       -   If a range is selected, then the subsequent macro will perform the actions on that section.
'       -   If only a single cell is selected, then the subsequent macro will perform the actions for the entire sheet.
'   -   Insert references into the Assumptions tab.
'       -   The macros work by finding the references on the Assumptions tab.
'       -   The references are unique values derived from the Assumptions and will be used in MATCH (lookup) formulas.
'   -   Determine which worksheet is active and therefore which macro will need to be called.


Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.ReferenceStyle = xlR1C1

' Assigning the current and Assumptions worksheets
    Dim ws, ass As Worksheet
    Set ws = ActiveSheet
    Set ass = Worksheets("Assumptions")

' Assigning the start time and month serived from the Assumptions tab (used in some calculations)
    Dim start_time, start_month As Date
    start_time = Now()
    start_month = ass.Cells(1, 2)
    start_month = DateSerial(Year(start_month), Month(start_month), 1)

' Assigning the used range
    Dim FR, LR, FC, LC, actual_fr, actual_fc, actual_lr, actual_lc As Long
    Dim rng As Range
    ' get range of entire worksheet
        If ws.Name = "Summary" Then
            actual_fr = 4
            actual_fc = 6  'Application.WorksheetFunction.Match(ass.Cells(1, 2), Rows(1), 0)
            actual_lr = Cells(1048576, 1).End(xlUp).Row
            actual_lc = Cells(2, 16384).End(xlToLeft).Column
        Else
            actual_fr = 3
            actual_fc = Application.WorksheetFunction.Match(ass.Cells(1, 2), Rows(2), 0)
            actual_lr = Cells(1048576, 1).End(xlUp).Row
            actual_lc = Cells(2, 16384).End(xlToLeft).Column
        End If
    
    ' get range of selected cells
        FR = Selection.Cells(1).Row
        FC = Selection.Cells(1).Column
        LR = Selection.Cells(Selection.Cells.Count).Row
        LC = Selection.Cells(Selection.Cells.Count).Column
    
    ' determine whether to calculate selected range or entire worksheet
        If Selection.CountLarge = 1 Then
            FR = actual_fr
            FC = actual_fc
            LR = actual_lr
            LC = actual_lc
        Else
            FR = Application.Max(FR, actual_fr)
            FC = Application.Max(FC, actual_fc)
            LR = Application.Min(LR, actual_lr)
            LC = Application.Min(LC, actual_lc)
        End If
    
        
        Set rng = Range(Cells(FR, FC), Cells(LR, LC))


' the summary sheet macro is very complex and will take a very long time to complete (normally and overnight process).
' This will verify that you really intend on doing this before any further action is taken.
    
    If ws.Name = "Summary" And FR = actual_fr And FC = actual_fc And LR = actual_lr And LC = actual_lc Then
    
        If MsgBox("Are you sure you want to run the macro on the entire summary sheet?" & vbCrLf & vbCrLf & "It will take a very long time!", vbYesNo) = vbNo Then GoTo TheEnd
    
    End If


' insert references into assumptions tab
    Dim ass_LR, ass_lc As Long
    Dim ass_section, ass_ref As Range
    
    ' Assumptions Last Row and Last Column
    ass_LR = ass.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    ass_lc = ass.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column + 2
    
    ' Set the range on the Assumptions worksheet where the calculations will be placed.
    Set ass_section = Range(ass.Cells(3, ass_lc), ass.Cells(ass_LR, ass_lc))
    Set ass_ref = Range(ass.Cells(3, ass_lc + 1), ass.Cells(ass_LR, ass_lc + 1))
    
    With ass_section
        .FormulaR1C1 = "=IF(R[1]C1=""Geo"",RC1,R[-1]C)"
        .Value = .Value
    End With
    
    With ass_ref
        .FormulaR1C1 = "=RC" & ass_lc & "&RC1&RC2&RC3&RC4&RC5&RC6&RC7&RC8&RC9"
        .Value = .Value
    End With


' Determine which worksheet is active and therefore which macro will need to be called and pass relevant fields (calculated ranges, assumption ranges etc)
    Select Case ws.Name
        Case "pipe_create"
            Call fm_1_pipe_create(rng, ass_lc + 1)
        
        Case "pipe_transposed"
            Call fm_2_pipe_transposed(rng, ass_lc + 1, FC)
        
        Case "pipe_create_won"
            Call fm_3_pipe_create_won(rng, ass_lc + 1)
        
        Case "pipe_create_opps"
            Call fm_4_pipe_create_opps(rng, ass_lc + 1)
        
        Case "pre_q_outflow"
            Call fm_5_pre_q_outflow(rng, ass_lc + 1)
        
        Case "pre_q_inflow"
            Call fm_5_pre_q_inflow(rng, ass_lc + 1)
        
        Case "in_q_outflow"
            Call fm_5_in_q_outflow(rng, ass_lc + 1)
        
        Case "in_q_inflow"
            Call fm_5_in_q_inflow(rng, ass_lc + 1)
        
        Case "adjusted_open_pipe"
            Call fm_5_adjusted_open_pipe(rng, ass_lc + 1)
        
        Case "existing_pipe_won"
            Call fm_5_existing_pipe_won(rng, ass_lc + 1)
        
        Case "marketing_opps"
            Call fm_7_marketing_opps(rng, ass_lc + 1)
        
        Case "marketing_qls"
            Call fm_8_marketing_qls(rng, ass_lc + 1)
            
        Case "Summary"
            Call fm_9_summary(rng, start_month)
        
        Case "Pipe Create Summary"
            Call fm_x_pipe_balance_summary
        
    End Select
    

' cleanup (remove references on Assumptions tab)
    ass_section.ClearContents
    ass_ref.ClearContents


TheEnd:
    Application.Calculate
    If Environ$("Username") <> "matbu" Then Application.ReferenceStyle = xlA1
    Application.ScreenUpdating = True
    MsgBox "Completed in " & Format(Now() - start_time, "HH:mm:ss") & "!"
End Sub

Private Sub fm_1_pipe_create(rng As Range, ass_lc)

' This macro will calculate the pipe create
' All columns on this sheet represent the month the pipe was created.
' This does this by multiplying the headcount by the tarets by product spluts by source splits by deal type splits (i.e. percentage of percentage of percentage etc)

' 1.    Each part of the calculation uses Excels builtin MATCH function to find specific rows on the Assumptions sheet (it's faster to it this way)!
' 2.    It will then surround the MATCH formula with the rest of the calculation as a text string that includes the MATCH-ed row references.
' 3.    It then pastes values to convert the text string formula into an actual formula on the worksheet.

' For the pipe create calculation, these are then concatenated into a very long formula to capture all of the percentages on the assumptions tab.


Dim hc_calc, productivity_calc, prod_calc, source_calc, deal_calc As String
Dim calc As String

' setup individual calculations to find the references in the assumptions sheet and add the text/string calculations
    hc_calc = """=Assumptions!R""&MATCH(""Headcount""&RC2&RC3&RC4&R2C,Assumptions!C" & ass_lc & ",0)&""C10"
    productivity_calc = "*(Assumptions!R""&MATCH(""Productivity""&RC2&RC3&RC4&YEAR(R2C)&""-Q""&CEILING.MATH(MONTH(R2C)/3,1),Assumptions!C" & ass_lc & ",0)&""C10/3)"
    prod_calc = "*Assumptions!R""&MATCH(""Product Split""&RC2&RC4&RC5,Assumptions!C" & ass_lc & ",0)&""C10"
    source_calc = "*Assumptions!R""&MATCH(""Source Split""&RC1&RC4&RC5&RC6&YEAR(R2C)&""-Q""&CEILING.MATH(MONTH(R2C)/3,1),Assumptions!C" & ass_lc & ",0)&""C10"
    deal_calc = "*Assumptions!R""&MATCH(""Deal Type""&RC5&RC7&YEAR(R2C)&""-Q""&CEILING.MATH(MONTH(R2C)/3,1),Assumptions!C" & ass_lc & ",0)&""C10"""


' merge calculations
    calc = "=IFERROR(" & hc_calc & productivity_calc & prod_calc & source_calc & deal_calc & ",0)"


' paste values to convert the text/string formulas into actual formulas
    With rng
        .FormulaR1C1 = calc
        .Value = .Value
    End With


End Sub


Private Sub fm_2_pipe_transposed(rng As Range, ass_lc, FC)

' This macro will create the calculations on the pipe create tansposed worksheet.

' Each column in the previous tab (pipe_create) represents the month when the pipe was created.
' We need to switch this so that each column represents the month the pipe will close.
' The sales cycle on the Assumptions tab will dictate the percentage of pipe create that we expect to close in the current month, or 1 month from now, or 2 months etc. up to 12 months out.
' Therefore each cell on the pipe_create tab will need to be multiplied against each of the 13 Sales Cycle percentages on the Assumption sheet.

' The layout of the worksheet is as follows...
' Each of the pipe_transposed columns represents the month when the pipe will close.
' Each of the rows represents how long ago the pipe was created (in month, 1 month [ago], 2 months [ago] etc.)
' For example, cell S5 will show a percentage of the pipe created 2 months ago (in July) and is expected to close in September.

' The macro...
' 1. The Assumption references to the Sales Cycle were created during the main macro.
' 2. the pipe_transposed macro will insert the necessary references on the pipe_create tab.
' 3. The calculation works the same as the pipe_create calculations by using the MATCH formula to find the appropiate row references. Then surround the reference with a text/string of the calculation.
' 4. However, we can't have pipe created in the past, so the calculation will start with an IF statement.
' 5. This IF statement will determine if the first characters in column H are numbers. If yes, then it will return the number and if not, it will return 0.
' 6. It will use this number to find the relative reference of the column in the pipe_create tab (same column (i.e. 0), or minus 1, 2, 3, ... columns to represent 1, 2, 3, ... months ago).
' 7. It will use this relative column referece along with the absolute row reference from the MATCH-ed pipe_create references (from point 2), and multiply them by the MATCH-ed references on the Assumptions sheet (from point 1).


Dim create_ws As Worksheet
Dim create_rng As Range
Dim create_LR, create_LC As Long
Dim create_calc, ass_calc, calc As String

' insert the necessary references on the pipe_create tab
    Set create_ws = Worksheets("pipe_create")
    create_LR = create_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    create_LC = create_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column + 2
    Set create_rng = Range(create_ws.Cells(3, create_LC), create_ws.Cells(create_LR, create_LC))
    With create_rng
        .FormulaR1C1 = "=RC1&RC2&RC3&RC4&RC5&RC6&RC7"
        .Value = .Value
    End With

' setup individual calculations
    calc = "=IF(COLUMN()-IF(ISNUMBER(LEFT(RC8,2)*1),LEFT(RC8,2),0)>=10,""=pipe_create!R""&MATCH(RC1&RC2&RC3&RC4&RC5&RC6&RC7,pipe_create!C" & create_LC & ",0)&""C""&COLUMN()-IF(ISNUMBER(LEFT(RC8,2)*1),LEFT(RC8,2),0)&""*Assumptions!R""&MATCH(""Sales Cycle""&RC4&RC5&RC7&RC8,Assumptions!C" & ass_lc & ",0)&""C10"")"

' paste values to convert the text/string formulas into actual formulas
    With rng
        .FormulaR1C1 = calc
        .Value = .Value
    End With


create_rng.ClearContents

End Sub

Private Sub fm_3_pipe_create_won(rng As Range, ass_lc)

' This macro will calculate how much of the pipe we expect to see Won in each month.
' All columns on this sheet represent the month the pipe is expected to close.

' The calculation is a fairly simply SUMIFS to return the same column on the pipe_transposed sheet, based on the breakouts in columns A-G.
' This will then be multiplied by the appropirate Win Rate on the Assumptions sheet.

Dim calc As String

' setup individual calculations
    calc = "=IFERROR(""=SUMIFS(pipe_transposed!C,pipe_transposed!C1,RC1,pipe_transposed!C2,RC2,pipe_transposed!C3,RC3,pipe_transposed!C4,RC4,pipe_transposed!C5,RC5,pipe_transposed!C6,RC6,pipe_transposed!C7,RC7)*Assumptions!R""&MATCH(""Win Rates""&RC1&RC5&RC6&RC7&YEAR(R2C)&""-Q""&CEILING.MATH(MONTH(R2C)/3,1),Assumptions!C" & ass_lc & ",0)&""C10"",0)"

' paste values to convert the text/string formulas into actual formulas
    With rng
        .FormulaR1C1 = calc
        .Value = .Value
    End With

End Sub


Private Sub fm_4_pipe_create_opps(rng, ass_lc)

' This macro will create the calculations for the pipe_create_opps
' Each column in this sheet shows the quantity of opportunities to be created each month.
' The quantity of opportunities is determined by dividing the amount of pipe created by the expected Average Selling Price (ASP).
'
' The macro will...
' 1. Insert the necessary references in the pipe_create sheet for find the appropriate pipe created.
' 2. It will use the Assumption references, created in the main macro, to find the ASP.
' 3. Create the calculation in the standard way by dividing point 1 by point 2.

Dim pipe_ws As Worksheet
Dim pipe_rng As Range
Dim pipe_LR, pipe_LC As Long
Dim calc As String

' insert reference on existing pipe tab
    Set pipe_ws = Worksheets("pipe_create")
    pipe_LR = pipe_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    pipe_LC = pipe_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column + 2
    Set pipe_rng = Range(pipe_ws.Cells(3, pipe_LC), pipe_ws.Cells(pipe_LR, pipe_LC))
    With pipe_rng
        .FormulaR1C1 = "=RC1&RC2&RC3&RC4&RC5&RC6&RC7"
        .Value = .Value
    End With


' paste values to convert the text/string formulas into actual formulas

'This calculation can be turned on for all non LATAM territories
calc = "=IFERROR(""=IFERROR(pipe_create!R"" & MATCH(RC1&RC2&RC3&RC4&RC5&RC6&RC7, pipe_create!C" & pipe_LC & ", 0) & ""C/(Assumptions!R"" & MATCH(""ASP"" & RC1 & RC4 & RC5 & RC6 & RC7, Assumptions!C" & ass_lc & ", 0) & ""C10),0)"",0)"

'This calculation can be turned on for all LATAM territories
'calc = "=IFERROR(""=IFERROR(pipe_create!R"" & MATCH(RC1&RC2&RC3&RC4&RC5&RC6&RC7, pipe_create!C" & pipe_LC & ", 0) & ""C/(Assumptions!R"" & MATCH(""ASP"" & RC1 & RC4 & RC5 & RC6 & RC7, Assumptions!C" & ass_lc & ", 0) & ""C10*Assumptions!R5148C45),0)"",0)"


With rng
    .FormulaR1C1 = calc
    .Value = .Value
End With

End Sub

Private Sub fm_5_pre_q_outflow(rng, ass_lc)

' this is the macro for Pre Quarter outflow, refering to Existing Open Pipe that will be "pushed" out prior to the quarter starting.
' This is pipe that was originally expected to close in a certain period but, for whatever reason, will now close in a later period.
' This macro specifically refers to pipe that will be pushed before the start of a quarter, rather than pipe that is pushed in the same quarter (we have different rates for each).

' The macro...
' 1. Insert the necessary references in the existing_pipe sheet to find the appropriate existing open pipe.
' 2. It will use the Assumption references, created in the main macro, to find the Pre-Q Push Rate.
'       The Assumption Push rates are grouped into four.
'       The amount of Existing pipe to be pushed out. <-- This is what is used in this sheet!
'       Then the percentage of the pushed pipe that will "land" in each month (3, 6, 9 months away) <-- this will be used in the next macro!
' 3. The calculation first uses an IF to determine if the month column is within the current quarter. If it is, then no calculation will happen.
' 4. Then, using the references from the existing_pipe and Assumptions it will multiply the Existing Open Pipe, by the Push rate to find out how much will be pushed from the current month (i.e. how much will be lost)

Dim pipe_ws As Worksheet
Dim pipe_rng As Range
Dim pipe_LR, pipe_LC As Long
Dim calc As String

' insert reference on existing pipe tab
    Set pipe_ws = Worksheets("existing_pipe")
    pipe_LR = pipe_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    pipe_LC = pipe_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column + 2
    Set pipe_rng = Range(pipe_ws.Cells(3, pipe_LC), pipe_ws.Cells(pipe_LR, pipe_LC))
    With pipe_rng
        .FormulaR1C1 = "=RC1&RC2&RC3&RC4&RC5&RC6&RC7&RC8"
        .Value = .Value
    End With

' setup individual calculations
    calc = "=IFERROR(""=IF(YEAR(R2C)&""""-Q""""&CEILING.MATH(MONTH(R2C)/3,1)>YEAR(Assumptions!R1C2)&""""-Q""""&CEILING.MATH(MONTH(Assumptions!R1C2)/3,1),-existing_pipe!R""&MATCH(RC1&RC2&RC3&RC4&RC5&RC6&RC7&""Open"",existing_pipe!C" & pipe_LC & ",0)&""C*Assumptions!R""&MATCH(""Pre-Q Push Rate""&YEAR(R2C)&""-Q""&CEILING.MATH(MONTH(R2C)/3,1)&""Pre-Q push rate"",Assumptions!C" & ass_lc & ",0)&""C10)"",0)"

' paste values to convert the text/string formulas into actual formulas
    With rng
        .FormulaR1C1 = calc
        .Value = .Value
    End With


pipe_rng.ClearContents
End Sub

Private Sub fm_5_pre_q_inflow(rng, ass_lc)

' this is the macro for Pre Quarter inflow, refering to Existing Open Pipe that will be "pushed" in prior to the quarter starting.
' This is the second part of the Pre Quarter Outflow and captures the amounts that were pushed out (lost) from previous periods.
' This macro specifically refers to pipe that will be pushed before the start of a quarter, rather than pipe that is pushed in the same quarter (we have different rates for each).

' The macro...
' 1. Insert the necessary references in the pre_q_outflow sheet to find the appropriate amounts that were previously pushed out.
' 2. It will use the Assumption references, created in the main macro, to find the Pre-Q Push Rates.
'       The Assumption Push rates are grouped into four.
'       The amount of Existing pipe to be pushed out.
'       Then the percentage of the pushed pipe that will "land" in each month (3, 6, 9 months away) <-- This is what is used in this sheet!
'    This is one of the longest calculations in the model, but it's essentially the same thing repeated three times for 3, 6, and 9 months.
' Concetrating on the first part of the calculation...
' 1. "=IF(COLUMN()-3>=10" will determine if the current column minus three is greater than 10. 10 being the first column in the pre_q_outflow sheet that could have numbers.
' 2. It will then MATCH the row on the pre_q_outflow sheet and use column -3 to get the ABS/positive amount that was previously pushed out.
' 3. It will then MATCH the row on the Assumptions tab to multiply it by the "Pre-Q Push Rate" "pushed out 3 months" rate.
' 4. This exact calculation is repeated two more times, except for -6 months and -9 months.
' 5. The three parts of the text/string calculation are concatenated together and pasted as values (proper calculations).
' 6. If ever one of the IF conditions returns FALSE, then it will return the double quotes and nothing will be concatenated into the text/string formula.

Dim pre_ws As Worksheet
Dim pre_rng As Range
Dim pre_LR, pre_LC As Long
Dim calc As String

' insert reference on pre-Q-outflow pipe tab
    Set pre_ws = Worksheets("pre_q_outflow")
    pre_LR = pre_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    pre_LC = pre_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column + 2
    Set pre_rng = Range(pre_ws.Cells(3, pre_LC), pre_ws.Cells(pre_LR, pre_LC))
    With pre_rng
        .FormulaR1C1 = "=RC1&RC2&RC3&RC4&RC5&RC6&RC7"
        .Value = .Value
    End With

' setup individual calculations
    calc = "=IF(COLUMN()-3>=10,""=(ABS(pre_q_outflow!R""&MATCH(RC1&RC2&RC3&RC4&RC5&RC6&RC7,pre_q_outflow!C" & pre_LC & ",0)&""C[-3])*Assumptions!R""&MATCH(""Pre-Q Push Rate""&YEAR(R2C)&""-Q""&CEILING.MATH(MONTH(R2C)/3,1)&""pushed out 3 months"",Assumptions!C" & ass_lc & ",0)&""C10)"")&IF(COLUMN()-6>=10,""+(ABS(pre_q_outflow!R""&MATCH(RC1&RC2&RC3&RC4&RC5&RC6&RC7,pre_q_outflow!C" & pre_LC & ",0)&""C[-6])*Assumptions!R""&MATCH(""Pre-Q Push Rate""&YEAR(R2C)&""-Q""&CEILING.MATH(MONTH(R2C)/3,1)&""pushed out 6 months"",Assumptions!C" & ass_lc & ",0)&""C10)"","""")&IF(COLUMN()-9>=10,""+(ABS(pre_q_outflow!R""&MATCH(RC1&RC2&RC3&RC4&RC5&RC6&RC7,pre_q_outflow!C" & pre_LC & ",0)&""C[-9])*Assumptions!R""&MATCH(""Pre-Q Push Rate""&YEAR(R2C)&""-Q""&CEILING.MATH(MONTH(R2C)/3,1)&""pushed out 9 months"",Assumptions!C" & ass_lc & ",0)&""C10)"","""")"

' paste values to convert the text/string formulas into actual formulas
    With rng
        .FormulaR1C1 = calc
        .Value = .Value
    End With


pre_rng.ClearContents
End Sub

Private Sub fm_5_in_q_outflow(rng, ass_lc)

' this is the macro for In Quarter outflow, refering to Existing Open Pipe that will be "pushed" out after the quarter commences.
' The logic is much the same as the Pre Quarter push however, we first need to factor for the pipe that has already been pushed around in the Pre Quarter calculations!
' Therefore, we'll need to get the Existing Open pipe each month, then subtract any pipe that was pushed out, and then add back in any pipe that was pushed in. All before we can start pushing it around some more!

' The macro...
' 1. Insert the necessary references in the existing_pipe, pre_q_outflow, and pre_q_inflow sheets to find the appropriate amounts that were previously pushed out.
' 2. It will use the Assumption references, created in the main macro, to find the In-Q Push Rates.
'       The Assumption Push rates are also grouped into four.
'       The amount of Existing pipe to be pushed out.  <-- This is what is used in this sheet!
'       Then the percentage of the pushed pipe that will "land" in each month (3, 6, 9 months away)
' The calculation...
' 1. Will MATCH the appropriate row and current column on the existing_pipe
' 2. Then MATCH the appropriate row and current column on the pre_q_outflow and add this negative amount to point 1 (i.e. subtract it!)
' 3. Then MATCH the appropriate row and current column on the pre_q_inflow and add this positive amount to point 2.
' 4. Then MATCH the "In-Q Push Rate" "In-Q push rate" in the Assumptions sheet and multiply it by Point 3.
' 5. The text/string calculation is then and pasted as values (proper calculations).

Dim calc As String

' insert reference on pre-Q-inflow pipe tab
    Dim pre_in_ws As Worksheet
    Dim pre_in_rng As Range
    Dim pre_in_LR, pre_in_LC As Long
    
    Set pre_in_ws = Worksheets("pre_q_inflow")
    pre_in_LR = pre_in_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    pre_in_LC = pre_in_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column + 2
    Set pre_in_rng = Range(pre_in_ws.Cells(3, pre_in_LC), pre_in_ws.Cells(pre_in_LR, pre_in_LC))
    With pre_in_rng
        .FormulaR1C1 = "=RC1&RC2&RC3&RC4&RC5&RC6&RC7"
        .Value = .Value
    End With

' insert reference on pre-Q-outflow pipe tab
    Dim pre_out_ws As Worksheet
    Dim pre_out_rng As Range
    Dim pre_out_LR, pre_out_LC As Long

    Set pre_out_ws = Worksheets("pre_q_outflow")
    pre_out_LR = pre_out_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    pre_out_LC = pre_out_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column + 2
    Set pre_out_rng = Range(pre_out_ws.Cells(3, pre_out_LC), pre_out_ws.Cells(pre_out_LR, pre_out_LC))
    With pre_out_rng
        .FormulaR1C1 = "=RC1&RC2&RC3&RC4&RC5&RC6&RC7"
        .Value = .Value
    End With

' insert reference on existing pipe tab
    Dim pipe_ws As Worksheet
    Dim pipe_rng As Range
    Dim pipe_LR, pipe_LC As Long

    Set pipe_ws = Worksheets("existing_pipe")
    pipe_LR = pipe_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    pipe_LC = pipe_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column + 2
    Set pipe_rng = Range(pipe_ws.Cells(3, pipe_LC), pipe_ws.Cells(pipe_LR, pipe_LC))
    With pipe_rng
        .FormulaR1C1 = "=RC1&RC2&RC3&RC4&RC5&RC6&RC7&RC8"
        .Value = .Value
    End With


' setup individual calculations
     calc = "=""=-(existing_pipe!R""&MATCH(RC1&RC2&RC3&RC4&RC5&RC6&RC7&""Open"",existing_pipe!C" & pipe_LC & ",0)&""C+pre_q_outflow!R""&MATCH(RC1&RC2&RC3&RC4&RC5&RC6&RC7,pre_q_outflow!C" & pre_out_LC & ",0)&""C+pre_q_inflow!R""&MATCH(RC1&RC2&RC3&RC4&RC5&RC6&RC7,pre_q_inflow!C" & pre_in_LC & ",0)&""C)*Assumptions!R""&MATCH(""In-Q Push Rate""&YEAR(R2C)&""-Q""&CEILING.MATH(MONTH(R2C)/3,1)&""In-Q push rate"",Assumptions!C" & ass_lc & ",0)&""C10"""

' paste values to convert the text/string formulas into actual formulas
    With rng
        .FormulaR1C1 = calc
        .Value = .Value
    End With


pre_out_rng.ClearContents
pre_in_rng.ClearContents
pipe_rng.ClearContents
End Sub

Private Sub fm_5_in_q_inflow(rng, ass_lc)

' this is the macro for In Quarter inflow, refering to Existing Open Pipe that will be "pushed" in after the quarter commences.
' This is the second part of the In Quarter Outflow and captures the amounts that were pushed out (lost) from previous periods.

' The macro...
' 1. Insert the necessary references in the in_q_outflow sheet to find the appropriate amounts that were previously pushed out.
' 2. It will use the Assumption references, created in the main macro, to find the In-Q Push Rates.
'       The Assumption Push rates are grouped into four.
'       The amount of Existing pipe to be pushed out.
'       Then the percentage of the pushed pipe that will "land" in each month (3, 6, 9 months away) <-- This is what is used in this sheet!
'    This is one of the longest calculations in the model but again, it's essentially the same thing repeated three times for 3, 6, and 9 months.
' Concetrating on the first part of the calculation...
' 1. "=IF(COLUMN()-3>=10" will determine if the current column minus three is greater than 10. 10 being the first column in the in_q_outflow sheet that could have numbers.
' 2. It will then MATCH the row on the in_q_outflow sheet and use column -3 to get the ABS/positive amount that was previously pushed out.
' 3. It will then MATCH the row on the Assumptions tab to multiply it by the "In-Q Push Rate" "pushed out 3 months" rate.
' 4. This exact calculation is repeated two more times, except for -6 months and -9 months.
' 5. The three parts of the text/string calculation are concatenated together and pasted as values (proper calculations).
' 6. If ever one of the IF conditions returns FALSE, then it will return the double quotes and nothing will be concatenated into the text/string formula.

Dim in_out_ws As Worksheet
Dim in_out_rng As Range
Dim in_out_LR, in_out_LC As Long
Dim calc As String

' insert reference on pre-Q-outflow pipe tab
    Set in_out_ws = Worksheets("in_q_outflow")
    in_out_LR = in_out_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    in_out_LC = in_out_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column + 2
    Set in_out_rng = Range(in_out_ws.Cells(3, in_out_LC), in_out_ws.Cells(in_out_LR, in_out_LC))
    With in_out_rng
        .FormulaR1C1 = "=RC1&RC2&RC3&RC4&RC5&RC6&RC7"
        .Value = .Value
    End With

' setup individual calculations
    calc = "=IF(COLUMN()-3>=10,""=(ABS(in_q_outflow!R""&MATCH(RC1&RC2&RC3&RC4&RC5&RC6&RC7,in_q_outflow!C" & in_out_LC & ",0)&""C[-3])*Assumptions!R""&MATCH(""In-Q Push Rate""&YEAR(R2C)&""-Q""&CEILING.MATH(MONTH(R2C)/3,1)&""pushed out 3 months"",Assumptions!C" & ass_lc & ",0)&""C10)"")&IF(COLUMN()-6>=10,""+(ABS(in_q_outflow!R""&MATCH(RC1&RC2&RC3&RC4&RC5&RC6&RC7,in_q_outflow!C" & in_out_LC & ",0)&""C[-6])*Assumptions!R""&MATCH(""In-Q Push Rate""&YEAR(R2C)&""-Q""&CEILING.MATH(MONTH(R2C)/3,1)&""pushed out 6 months"",Assumptions!C" & ass_lc & ",0)&""C10)"","""")&IF(COLUMN()-9>=10,""+(ABS(in_q_outflow!R""&MATCH(RC1&RC2&RC3&RC4&RC5&RC6&RC7,in_q_outflow!C" & in_out_LC & ",0)&""C[-9])*Assumptions!R""&MATCH(""In-Q Push Rate""&YEAR(R2C)&""-Q""&CEILING.MATH(MONTH(R2C)/3,1)&""pushed out 9 months"",Assumptions!C" & ass_lc & ",0)&""C10)"","""")"

' paste values to convert the text/string formulas into actual formulas
    With rng
        .FormulaR1C1 = calc
        .Value = .Value
    End With


in_out_rng.ClearContents
End Sub

Private Sub fm_5_adjusted_open_pipe(rng, ass_lc)

' this is the macro for Adjusted Open Pipe, refering to Existing Open Pipe after all the pre-quarter/in-quarter pushing.
' i.e. This is the final amount of Open Pipe we would expect to see that currently exists.
' The logic is pretty simple: Existing Open Pipe minus everything that was pushed out plus everything that was pushed in!

' The macro...
' 1. Insert the necessary references in the existing_pipe, pre_q_outflow, pre_q_inflow, in_q_outflow, and in_q_inflow sheets.
' 2. It will then MATCH the row on the existing_pipe sheet and use the current column to get the Existing Open Pipe.
' 3. It will then MATCH the row on the pre_q_outflow sheet and use the current column to get the negative value of what was pushed out prior to the start of the quarter.
' 4. It will then MATCH the row on the pre_q_inflow sheet and use the current column to get the positive amount of what was pushed in from point 3.
' 5. It will then MATCH the row on the in_q_outflow sheet and use the current column to get the negative value of what was pushed out after the quarter commenced.
' 6. It will then MATCH the row on the in_q_inflow sheet and use the current column to get the positive amount of what was pushed in from point 5.
' 7. It will then paste values to convert the text/string formulas into actual formulas

Dim calc As String

' insert reference on pre-Q-inflow pipe tab
    Dim pre_in_ws As Worksheet
    Dim pre_in_rng As Range
    Dim pre_in_LR, pre_in_LC As Long
    
    Set pre_in_ws = Worksheets("pre_q_inflow")
    pre_in_LR = pre_in_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    pre_in_LC = pre_in_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column + 2
    Set pre_in_rng = Range(pre_in_ws.Cells(3, pre_in_LC), pre_in_ws.Cells(pre_in_LR, pre_in_LC))
    With pre_in_rng
        .FormulaR1C1 = "=RC1&RC2&RC3&RC4&RC5&RC6&RC7"
        .Value = .Value
    End With

' insert reference on pre-Q-outflow pipe tab
    Dim pre_out_ws As Worksheet
    Dim pre_out_rng As Range
    Dim pre_out_LR, pre_out_LC As Long

    Set pre_out_ws = Worksheets("pre_q_outflow")
    pre_out_LR = pre_out_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    pre_out_LC = pre_out_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column + 2
    Set pre_out_rng = Range(pre_out_ws.Cells(3, pre_out_LC), pre_out_ws.Cells(pre_out_LR, pre_out_LC))
    With pre_out_rng
        .FormulaR1C1 = "=RC1&RC2&RC3&RC4&RC5&RC6&RC7"
        .Value = .Value
    End With

' insert reference on in-Q-inflow pipe tab
    Dim in_in_ws As Worksheet
    Dim in_in_rng As Range
    Dim in_in_LR, in_in_LC As Long
    
    Set in_in_ws = Worksheets("in_q_inflow")
    in_in_LR = in_in_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    in_in_LC = in_in_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column + 2
    Set in_in_rng = Range(in_in_ws.Cells(3, in_in_LC), in_in_ws.Cells(in_in_LR, in_in_LC))
    With in_in_rng
        .FormulaR1C1 = "=RC1&RC2&RC3&RC4&RC5&RC6&RC7"
        .Value = .Value
    End With

' insert reference on in-Q-outflow pipe tab
    Dim in_out_ws As Worksheet
    Dim in_out_rng As Range
    Dim in_out_LR, in_out_LC As Long

    Set in_out_ws = Worksheets("in_q_outflow")
    in_out_LR = in_out_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    in_out_LC = in_out_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column + 2
    Set in_out_rng = Range(in_out_ws.Cells(3, in_out_LC), in_out_ws.Cells(in_out_LR, in_out_LC))
    With in_out_rng
        .FormulaR1C1 = "=RC1&RC2&RC3&RC4&RC5&RC6&RC7"
        .Value = .Value
    End With
    
' insert reference on existing pipe tab
    Dim pipe_ws As Worksheet
    Dim pipe_rng As Range
    Dim pipe_LR, pipe_LC As Long

    Set pipe_ws = Worksheets("existing_pipe")
    pipe_LR = pipe_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    pipe_LC = pipe_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column + 2
    Set pipe_rng = Range(pipe_ws.Cells(3, pipe_LC), pipe_ws.Cells(pipe_LR, pipe_LC))
    With pipe_rng
        .FormulaR1C1 = "=RC1&RC2&RC3&RC4&RC5&RC6&RC7&RC8"
        .Value = .Value
    End With




' setup individual calculations
     calc = "=""=SUM(existing_pipe!R""&MATCH(RC1&RC2&RC3&RC4&RC5&RC6&RC7&""Open"",existing_pipe!C" & pipe_LC & ",0)&""C,pre_q_outflow!R""&MATCH(RC1&RC2&RC3&RC4&RC5&RC6&RC7,pre_q_outflow!C" & pre_out_LC & ",0)&""C,pre_q_inflow!R""&MATCH(RC1&RC2&RC3&RC4&RC5&RC6&RC7,pre_q_inflow!C" & pre_in_LC & ",0)&""C,in_q_outflow!R""&MATCH(RC1&RC2&RC3&RC4&RC5&RC6&RC7,in_q_outflow!C" & in_out_LC & ",0)&""C,in_q_inflow!R""&MATCH(RC1&RC2&RC3&RC4&RC5&RC6&RC7,in_q_inflow!C" & in_in_LC & ",0)&""C)"""

' paste values to convert the text/string formulas into actual formulas
    With rng
        .FormulaR1C1 = calc
        .Value = .Value
    End With


pre_out_rng.ClearContents
pre_in_rng.ClearContents
in_out_rng.ClearContents
in_in_rng.ClearContents
pipe_rng.ClearContents
End Sub


Private Sub fm_5_existing_pipe_won(rng, ass_lc)

' this is the macro for Existing Pipe Won.
' This will take the adjuted Open Pipe (after all the pushing) and apply the appropriate Win rates. Then add anything that's already set to Won.

' The macro...
' 1. Insert the necessary references in the adjusted_open_pipe and existing_pipe sheets.
' 2. It will use the Assumption references, created in the main macro, to find the Win Rates
' 3. It will then MATCH the row on the existing_pipe sheet and use the current column to get the Existing Won Pipe.
' 4. It will then MATCH the row on the adjusted_open_pipe sheet and use the current column to get the Open Pipe after all the pushing.
' 5. It will then multiply point 2 by point 4 and add point 3.
' 6. It will then paste values to convert the text/string formulas into actual formulas

Dim calc As String

' insert reference on adjusted_open_pipe tab
    Dim adj_ws As Worksheet
    Dim adj_rng As Range
    Dim adj_LR, adj_LC As Long
    
    Set adj_ws = Worksheets("adjusted_open_pipe")
    adj_LR = adj_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    adj_LC = adj_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column + 2
    Set adj_rng = Range(adj_ws.Cells(3, adj_LC), adj_ws.Cells(adj_LR, adj_LC))
    With adj_rng
        .FormulaR1C1 = "=RC1&RC2&RC3&RC4&RC5&RC6&RC7"
        .Value = .Value
    End With

    
' insert reference on existing pipe tab
    Dim pipe_ws As Worksheet
    Dim pipe_rng As Range
    Dim pipe_LR, pipe_LC As Long

    Set pipe_ws = Worksheets("existing_pipe")
    pipe_LR = pipe_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    pipe_LC = pipe_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column + 2
    Set pipe_rng = Range(pipe_ws.Cells(3, pipe_LC), pipe_ws.Cells(pipe_LR, pipe_LC))
    With pipe_rng
        .FormulaR1C1 = "=RC1&RC2&RC3&RC4&RC5&RC6&RC7&RC8"
        .Value = .Value
    End With




' setup individual calculations
     calc = "=IFERROR(""=SUM(existing_pipe!R""&MATCH(RC1&RC2&RC3&RC4&RC5&RC6&RC7&""Closed Won"",existing_pipe!C" & pipe_LC & ",0)&""C,existing_pipe!R""&MATCH(RC1&RC2&RC3&RC4&RC5&RC6&RC7&""Closed"",existing_pipe!C" & pipe_LC & ",0)&""C,adjusted_open_pipe!R""&MATCH(RC1&RC2&RC3&RC4&RC5&RC6&RC7,adjusted_open_pipe!C" & adj_LC & ",0)&""C)*Assumptions!R""&MATCH(""Win Rates""&RC1&RC5&RC6&RC7&YEAR(R2C)&""-Q""&CEILING.MATH(MONTH(R2C)/3,1),Assumptions!C" & ass_lc & ",0)&""C10"",0)"

' paste values to convert the text/string formulas into actual formulas
    With rng
        .FormulaR1C1 = calc
        .Value = .Value
    End With


adj_rng.ClearContents
pipe_rng.ClearContents
End Sub

Private Sub fm_7_marketing_opps(rng, ass_lc)

' this is the macro for Marketing opps.
' Each column represents the number of opportunities we expect to by created from Marketing activities, broken out into P1/P2 (Priority 1/Priority 2).

' The macro...
' 1. Insert the necessary references in the pipe_create_opps.
' 2. It will use the Assumption references, created in the main macro, to MATCH the P1 vs P2 split
' 3. It will then MATCH the row on the pipe_create_opps sheet and use the current column to get the Created Opps Sourced by Marketing.
' 4. It will then multiply point 2 by point 3.
' 6. It will then paste values to convert the text/string formulas into actual formulas


Dim opp_ws As Worksheet
Dim opp_rng As Range
Dim opp_LR, opp_LC As Long
Dim calc As String

' insert reference on pipe_create_opps tab
    Set opp_ws = Worksheets("pipe_create_opps")
    opp_LR = opp_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    opp_LC = opp_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column + 2
    Set opp_rng = Range(opp_ws.Cells(3, opp_LC), opp_ws.Cells(opp_LR, opp_LC))
    With opp_rng
        .FormulaR1C1 = "=RC1&RC2&RC3&RC4&RC5&RC6&RC7"
        .Value = .Value
    End With

' setup individual calculations
    calc = "=IFERROR(""=pipe_create_opps!R""&MATCH(RC1&RC2&RC3&RC4&RC5&RC6&RC7,pipe_create_opps!C" & opp_LC & ",0)&""C*Assumptions!R""&MATCH(""P1 vs P2 split""&RC4&RC5&""Q""&CEILING.MATH(MONTH(R2C)/3,1)&RC8,Assumptions!C" & ass_lc & ",0)&""C10"",0)"

' paste values to convert the text/string formulas into actual formulas
    With rng
        .FormulaR1C1 = calc
        .Value = .Value
    End With


opp_rng.ClearContents
End Sub

Private Sub fm_8_marketing_qls(rng, ass_lc)

' this is the macro for Marketing Qualified Leads.
' Each column represents the number of Leads are needed in order to created the number of opps in the marketing_opps sheet.

' The macro...
' 1. Insert the necessary references in the marketing_opps sheet.
' 2. It will use the Assumption references, created in the main macro, to MATCH the QL to Opp conversion
' 3. It will then MATCH the row on the marketing_opps sheet and use the current column to get the P1 or P2 Opps Sourced by Marketing.
' 4. It will then divide point 2 by point 3 to get the total number of leads required.
' 6. It will then paste values to convert the text/string formulas into actual formulas


Dim opp_ws As Worksheet
Dim opp_rng As Range
Dim opp_LR, opp_LC As Long
Dim calc As String

' insert reference on marketing_opps tab
    Set opp_ws = Worksheets("marketing_opps")
    opp_LR = opp_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    opp_LC = opp_ws.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column + 2
    Set opp_rng = Range(opp_ws.Cells(3, opp_LC), opp_ws.Cells(opp_LR, opp_LC))
    With opp_rng
        .FormulaR1C1 = "=RC1&RC2&RC3&RC4&RC5&RC6&RC7&RC8"
        .Value = .Value
    End With

' setup individual calculations
    calc = "=IFERROR(""=IFERROR(marketing_opps!R""&MATCH(RC1&RC2&RC3&RC4&RC5&RC6&RC7&RC8,marketing_opps!C" & opp_LC & ",0)&""C/Assumptions!R""&MATCH(""QL to Opp conversion""&RC4&RC5&""Q""&CEILING.MATH(MONTH(R2C)/3,1)&RC8,Assumptions!C" & ass_lc & ",0)&""C10,0)"",0)"

' paste values to convert the text/string formulas into actual formulas
    With rng
        .FormulaR1C1 = calc
        .Value = .Value
    End With


opp_rng.ClearContents
End Sub
            
            

Private Sub fm_9_summary(rng As Range, start_month)

'Be careful running the macro on the entire sheet as it will take a long time to run.
'
'Everything is grouped into ranges of about 30 rows by 15 columns.
'The macro takes care of everything that follows the format of these ranges.
'    i.e. Dmitri has added additional summaries at the bottom of the sheet that aren't included in this macro.
'Never insert rows or columns on this worksheet.
'
'Calculation Types...
'
'Within each group there are three main types of calculation...
'Never Changing Calculations:
'    These will never change, regardless of which group they're a part of.
'    They work vertically to return a sum or percentage.
'    Because these calculations never change, they're loaded into an array in VBA for quick access.
'
'Semi Changing Calculations:
'    These will never change, regarless of which group they're a part of but are more locallized to column.
'    They work horizontally to return a sum of the cells to the right.
'    These have "WW" in the column header.
'
'
'Individual calculations:
'    These calculations are broken out into three sub groups:
'        Aggregated calculations:
'            These are calculations that work virtically to aggregate from the groups lower on the sheet.
'            For example, the "All Products" group will SUM all the values from the individual product groups below it.
'            These are quite simple because they use relative references and simply sum the appropriate ranges below.
'
'        Quarter/Fiscal Calculations:
'            These are calculations that work horizontally to aggregate from the groups to the left.
'            These are quite simple because they use relative references and simply sum the appropriate ranges to the left.
'            These will have the Quarter or Fiscal number in the header.
'
'        Indiviaul PreSummary references:
'            These are the lowest form of calculations on the sheet and pull in the individual values from the PreSummary. All other calculations on this sheet are sourced from these.
'
'
'
'
'After assigning variables the macro will.
'1. Insert the Never Changing Calculations arrays in the last column for use later. This is so a simple index/match will return the correct formula.
'2. insert FALSE into the used range. The used range is either the selected cells or the entire worksheet.
'3. Loop through the used range to check if the column or row headers are populated.
'    If they're blank, then no calculations are required and the FALSE values will be individually removed.
'
'-- Start inserting calculations
'-- Never Changing and Semi Changing Calculations:
'4. The macro will use .SpecialCells(xlCellTypeConstants) to select all cells that should end up with calculations.
'5. an IF(NOT(ISERROR())) condition is used to see if the row header (column E) is included in the last column (point nr. 1).
'6. If True, then the Never Changing Calculations will be index/matched into the current cell based on the value in column E.
'7. If False, then the formula will look for "WW" in the header.
'8. If True, then the formula will look at the sub header to determine which formula to insert. There's three possible alternatives but they all SUM values from the right of the current cell using relative references.
'9. If False, then the formula will return an error.
'
'-- Individual calculations
'-- Aggregated calculations:
'10. The macro will now use .SpecialCells(xlCellTypeFormulas, 1) to find formulas that have returned an error and should be replaced with proper formulas.
'11. Based on the values in Columns A, B, and/or, C; the formula will determine whether it should aggregate all sources, all products, OR all customers.
'    The rows have been very carefully laid out so that the groups below are always in the same position in relation to each other.
'    i.e. New Customer is always immediately below All Customers. Tosca is always immediately below All Products. Marketing Sourced is always immediatly below All Sources. etc.
'    With this logic, we can use the relative references to simply aggregate from the groups below.
'12. The macro will use an IF statement to determine if all sources, all products, or all customers appears in columns A, B, or C and will insert the appropriate formula.
'13. If False, it will return an error.
'
'-- Quarter/Fiscal calculations:
'14. The macro will still use .SpecialCells(xlCellTypeFormulas, 1) to select all errors to update with proper formulas.
'15. The inserted formula will check if "Q" or "FY" appears in the header and will insert the appropriate header.
'16. If False then it will return an error.
'
'-- Indiviaul PreSummary calculations:
'17. The macro will insert the row number into the last column on the PreSummary sheet. This is so we can use MINIFS to return a row reference with multiple criteria (without needing a slower array formula)
'18. It will then insert Closed, Open, Closed Won in the last column of the PreSummary based on the values in column E. This is so we can reference the Existing Opportunities by stage.
'19. The macro will then use .SpecialCells(xlCellTypeFormulas, 1) to select all errors (these should be all remaining cells) and insert the following formula...
'
'    The formula will first check if the header date is in the past by comparing it to the date in the Assumptions tab.
'    If the date is in the past, it will then check for the word "existing" in column E. This will determine if we want to bring in the actual pipe that happened in the past.
'    If true, then the formula will reference the pivot table in the "past opps" worksheet to find the Open, Closed, or Closed Won opportunities.
'    If False, then the formula will check in the PreSummary for the existance of the matching
'        Region (based on the header),
'        Source, Product, Customer Type, and row header (column E)
'    If found, then it will return the row number based on the MINIFS statement and the column based on matching the column date.
'
'20. Finally the macro will loop through the individual cells again and paste the text/string formulas into actual formulas, and then remove all other temporaty references in the last columns.


Dim calc As String

' assign variables relating to the summary tab
    Dim summary As Worksheet
    Dim FR, LR, FC, LC, summary_LC As Long
    Dim summary_rng As Range
    Set summary = ActiveSheet
    FR = rng.Cells(1).Row
    FC = rng.Cells(1).Column
    LR = rng.Cells(rng.Cells.Count).Row
    LC = rng.Cells(rng.Cells.Count).Column
    summary_LC = Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column + 2

' assign variables relating to the pre-summary tab.
    Dim preSummary As Worksheet
    Dim pre_rng As Range
    Dim pre_LR, pre_LC As Long
    Set preSummary = Worksheets("PreSummary")
    pre_LR = preSummary.Cells(1048576, 1).End(xlUp).Row
    pre_LC = preSummary.Cells(2, 16384).End(xlToLeft).Column + 2


' setup arrays for never changing calculations
    Dim never_changing_str, never_changing_calc, semi_changing_str, semi_changing_calc As Variant
    Dim i As Long



    never_changing_str = Array("ASP (Blended)", _
                               "ASP (individual product)", _
                               "Existing Pipe (for close in period)", _
                               "Adjusted Total Open Pipe", _
                               "Adjusted Existing Total Pipe", _
                               "Total Pipe (new, existing, pushed)", _
                               "Bookings (NACV)", _
                               "Win Rate (Blended)", _
                               "From new pipe", _
                               "From existing/push pipe")
                               
    never_changing_calc = Array("=""=IFERROR(R[-6]C/R[-3]C,0)""", _
                                "=""=IFERROR(R[-7]C/R[-3]C,0)""", _
                                "=""=SUM(R[1]C:R[3]C)""", _
                                "=""=SUM(R[-5]C:R[-1]C)""", _
                                "=""=SUM(R[-8]C,R[-7]C,R[-1]C)""", _
                                "=""=SUM(R[-19]C,R[-2]C)""", _
                                "=""=SUM(R[1]C:R[2]C)""", _
                                "=""=IFERROR(R[-4]C/R[-6]C,0)""", _
                                "=""=IFERROR(R[-4]C/R[-26]C,0)""", _
                                "=""=IFERROR(R[-4]C/R[-10]C,0)""")


' insert never changing calculations into the last column on the summary tab
    For i = 0 To UBound(never_changing_str)
        Cells(i + 1, summary_LC) = never_changing_str(i)
        Cells(i + 1, summary_LC + 1) = never_changing_calc(i)
    Next i


On Error GoTo Skip1
' find range that should be calculated and set everything to FALSE
    Range(Cells(FR, FC), Cells(LR, LC)) = False


' remove FALSE for cells that don't have values in the column or row headers (and shouldn't have calculations)
    For i = FR To LR
        If IsEmpty(Cells(i, 5)) Then Range(Cells(i, FC), Cells(i, LC)).ClearContents
    Next i
    For i = FC To LC
        If IsEmpty(Cells(2, i)) Or Cells(2, i) = "<blank>" Then Range(Cells(FR, i), Cells(LR, i)).ClearContents
    Next i

 
' Select the FALSE range that should be calculated with never changing or semi changing calculations
' If true then insert the formula, if false then return an error.
    Set summary_rng = Range(Cells(FR, FC), Cells(LR, LC)).SpecialCells(xlCellTypeConstants)
    With summary_rng
        .FormulaR1C1 = "=IF(NOT(ISERROR(MATCH(RC5,C" & summary_LC & ",0))),INDEX(C" & summary_LC + 1 & ",MATCH(RC5,C" & summary_LC & ",0)),IF(R2C=""WW"",IF(R3C=""Total"",""=SUM(RC[1]:RC[2])"",IF(R3C=""Core IT"",""=SUM(RC[2]:RC[14])-RC[1]"",IF(R3C=""Devops"",""=SUM(RC[6],RC[11],RC[13])""))),0))"
    End With



' Select the range with errors that should be calculated with aggregated calculations (all sources, all products, all customers)
' If true then insert the formula, if false then return an error.
    Set summary_rng = Range(Cells(FR, FC), Cells(LR, LC)).SpecialCells(xlCellTypeFormulas, 1)
    With summary_rng
        .Formula2R1C1 = "=IF(RC3=""All Customers"",""=SUM(R[32]C,R[64]C,R[96]C)"",IF(RC2=""All Products"",""=SUM(R[128]C,R[256]C,R[384]C,R[512]C,R[640]C,R[768]C,R[896]C,R[1024]C,R[1152]C,R[1280]C)"",IF(RC1=""All Sources"",""=SUM(R[1408]C,R[2816]C,R[4224]C,R[5632]C)"",0)))"
    End With
    

' Select the range with errors that should be calculated with Quarter or Fiscal calculations
' If true then insert the formula, if false then return an error.
    Set summary_rng = Range(Cells(FR, FC), Cells(LR, LC)).SpecialCells(xlCellTypeFormulas, 1)
    With summary_rng
        .Formula2R1C1 = "=IF(MID(R1C,6,1)=""Q"",""=SUM(RC[-17],RC[-34],RC[-51])"",IF(LEFT(R1C,2)=""FY"",""=SUM(RC[-221],RC[-153],RC[-85],RC[-17])"",0))"
    End With



' insert row references into the pre summary.
' These will be used later with a MINIFS lookup-type formula.
    Set pre_rng = Range(preSummary.Cells(3, pre_LC), preSummary.Cells(pre_LR, pre_LC))
    With pre_rng
        .FormulaR1C1 = "=ROW()"
        .Value = .Value
    End With
    
' insert range on summary to help with pivot references (converting the Existing pipe rows into their appropriate Open, Closed, Closed Won values for lookups)
        With Range(Cells(20, summary_LC), Cells(LR, summary_LC))
            .Formula2R1C1 = "=IF(RC5=""Existing Pipe (Closed)"",""Closed"",IF(RC5=""Existing Pipe (Open)"",""Open"",IF(OR(RC5=""Existing Pipe (Won)"",RC5=""Bookings (NACV) from existing/push pipe""),""Closed Won"")))"
            .Value = .Value
        End With
    
' select the range with errors (this should be everything else) that should be calculated from PreSummary
' If the column date is in the past and the row refers to existing pipe, then calculate from the "Past Opps" pivot table.
' otherwise MATCH the row reference from the PreSummary using the MINIFS formula and MATCH the column using the header date to return the PreSummary Reference.
    
    Set summary_rng = Range(Cells(FR, FC), Cells(LR, LC)).SpecialCells(xlCellTypeFormulas, 1)
    calc = "=IF(R1C<Assumptions!R1C2,IF(ISNUMBER(FIND(""existing"",LOWER(RC5))),""=IFERROR(GETPIVOTDATA(""""Product_NACV"""",'past opps'!R3C1,""""Region"""",SUBSTITUTE(R2C&""""-""""&R3C,""""APAC-Core IT"""",""""APAC""""),""""Product"""",RC2,""""source"""",RC1,""""Stage"""",""""""&RC" & summary_LC & "&"""""",""""deal type"""",RC3,""""Close_Month"""",R1C),0)"",0),IF(MINIFS(PreSummary!C" & pre_LC & ",PreSummary!C2,SUBSTITUTE(R2C&""-""&R3C,""APAC-Core IT"",""APAC""),PreSummary!C6,RC1,PreSummary!C5,RC2,PreSummary!C7,RC3,PreSummary!C9,RC5)>0,""=PreSummary!R""&MINIFS(PreSummary!C" & pre_LC & ",PreSummary!C2,SUBSTITUTE(R2C&""-""&R3C,""APAC-Core IT"",""APAC""),PreSummary!C6,RC1,PreSummary!C5,RC2,PreSummary!C7,RC3,PreSummary!C9,RC5)&""C""&MATCH(R1C,PreSummary!R2,0),0))"
    summary_rng = calc
    

' remove reference column from PreSummary
    pre_rng.ClearContents


Skip1:
' transform formula strings to true formulas
    For i = FR To LR
        If Not IsEmpty(Cells(i, 5)) Then
            With Range(Cells(i, FC), Cells(i, LC))
                .Value = .Value
            End With
        End If
    Next i


' Remove reference from summary
    Range(Cells(1, summary_LC), Cells(1, summary_LC + 1)).EntireColumn.ClearContents

End Sub

Private Sub fm_x_pipe_balance_summary()
Dim LR As Long
Dim rng As Range

LR = Cells(1048576, 1).End(xlUp).Row
Set rng = Range(Cells(2, 7), Cells(20, 7))

With rng
    .Formula2R1C1 = "=""=SUMIFS(pipe_transposed!C""&MATCH(RC2,pipe_transposed!R2,0)+IF(RC4=""in month"",0,LEFT(RC4,2))&"",pipe_transposed!C2,RC1,pipe_transposed!C8,RC4)"""
    .Value = .Value
End With


End Sub








