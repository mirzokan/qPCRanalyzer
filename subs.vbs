Option Explicit

Sub amp_normalization(ws As Worksheet, i As Integer, g as Integer)
 'Reference Data to Amplitude Normalization Sheet
    Dim categories As Integer, datas As Integer 'Ubounds
    Dim ci As Integer, di As Integer, ddi As Integer 'Counters
    Dim fj As String

    With ws_ampnorm.Cells
        .Validation.Delete
    End With

    'Referencing Starts
        categories = row_count(ws.Cells(1,1))
        For ci=0 To (categories-1)
            datas = col_count(ws.Cells(ci+1,1))-1
            For di=0 To datas-1
                ' Sample Type
                rb_ampnorm_base.Offset(0,i).Value = "=Data_Import!" &  rb_import_base.Offset(0,ws.Cells(ci+1,di+2).Value-1).Address(False,False)
                ' Concentration
                rb_ampnorm_base.Offset(1,i).Value = "=Data_Import!" &  rb_import_base.Offset(1,ws.Cells(ci+1,di+2).Value-1).Address(False,False)
                ' Sample Name
                rb_ampnorm_base.Offset(2,i).Value = "=Data_Import!" &  rb_import_base.Offset(2,ws.Cells(ci+1,di+2).Value-1).Address(False,False)
                ' Normalization On/Off
                If rb_ampnorm_base.Offset(0,i).Value = "Blank" Then
                    rb_ampnorm_base.Offset(3,i).Value = "0"
                Else
                    rb_ampnorm_base.Offset(3,i).Value = "1"
                End If
                ' Index
                rb_ampnorm_base.Offset(14,i).Value = "=Data_Import!" &  rb_import_base.Offset(-1,ws.Cells(ci+1,di+2).Value-1).Address(False,False)
                ' Group
                rb_ampnorm_base.Offset(15,i).Value = g+1
                ' Coordinate
                rb_ampnorm_base.Offset(16,i).Value = "=Data_Import!" &  rb_import_base.Offset(3,ws.Cells(ci+1,di+2).Value-1).Address(False,False)
                ' First Cycle, Formula =((Data_Import!B16-MIN(Data_Import!B16:Data_Import!B47))+2)*B5
                rb_ampnorm_base.Offset(17, i).Value = "=((Data_Import!" & rb_import_base.Offset(4, ws.Cells(ci + 1, di + 2).Value - 1).Address(False, False) & "-MIN(Data_Import!" & rb_import_base.Offset(4, ws.Cells(ci + 1, di + 2).Value - 1).Address & ":Data_Import!" & rb_import_base.Offset(4 + (num_cycles - 1), ws.Cells(ci + 1, di + 2).Value - 1).Address & "))+2)*" & rb_ampnorm_base.Offset(4, i).Address
                ' Fill Down
                ws_ampnorm.Range(rb_ampnorm_base.Offset(17,i), rb_ampnorm_base.Offset(17+(num_cycles.Value-1),i)).FillDown
                 'Normalization Multiplier, Formula =IF(E49 = 1, (MIN(MAX(E6:E45)-MIN(E6:E45),MAX(F6:F45)-MIN(E6:E45),MAX(G6:G45)-MIN(E6:E45))/(MAX(E6:E45)-MIN(E6:E45))),1)
                    fj = "=IF(" & rb_ampnorm_base.Offset(3,i).Address & "=1,(MIN("
                    For ddi=0 to (datas-1)
                        fj = fj & "MAX(Data_Import!" &  rb_import_base.Offset(4,ws.Cells(ci+1,ddi+2).Value-1).Address & ":Data_Import!" &  rb_import_base.Offset(4+(num_cycles.Value-1),ws.Cells(ci+1,ddi+2).Value-1).Address & ")-MIN(Data_Import!" &  rb_import_base.Offset(4,ws.Cells(ci+1,ddi+2).Value-1).Address & ":Data_Import!" &  rb_import_base.Offset(4+(num_cycles.Value-1),ws.Cells(ci+1,ddi+2).Value-1).Address & ")"
                        If ddi < (datas-1) Then
                            fj = fj & ","
                        End If
                    Next ddi
                    fj = fj & ")/(MAX(Data_Import!" &  rb_import_base.Offset(4,ws.Cells(ci+1,di+2).Value-1).Address & ":Data_Import!" &  rb_import_base.Offset(4+(num_cycles.Value-1),ws.Cells(ci+1,di+2).Value-1).Address &  ")-MIN(Data_Import!" &  rb_import_base.Offset(4,ws.Cells(ci+1,di+2).Value-1).Address & ":Data_Import!" &  rb_import_base.Offset(4+(num_cycles.Value-1),ws.Cells(ci+1,di+2).Value-1).Address & "))),1)"
                    rb_ampnorm_base.Offset(4,i).Formula = fj

                ' Y2, Array Formula =IF((MIN(IF((E53:E92-AmpNorm Standard Curve'!B1)>0,(E53:E92))))=0,"X",(MIN(IF((E53:E92-AmpNorm Standard Curve'!B1)>0,(E53:E92)))))
                rb_ampnorm_base.Offset(5,i).FormulaArray = "=IF((MIN(IF((" & rb_ampnorm_base.Offset(17,i).Address & ":" & rb_ampnorm_base.Offset(17+(num_cycles.Value-1),i).Address & "-Data_Review!" & rthreshold.Address & ")>0,(" & rb_ampnorm_base.Offset(17,i).Address & ":" & rb_ampnorm_base.Offset(17+(num_cycles.Value-1),i).Address & "))))=0,""X"",(MIN(IF((" & rb_ampnorm_base.Offset(17,i).Address & ":" & rb_ampnorm_base.Offset(17+(num_cycles.Value-1),i).Address & "-Data_Review!" & rthreshold.Address & ")>0,(" & rb_ampnorm_base.Offset(17,i).Address & ":" & rb_ampnorm_base.Offset(17+(num_cycles.Value-1),i).Address & ")))))"
                ' Y1, Formula =OFFSET(INDEX(E53:E92,MATCH(E96,E53:E92,0)),-1,0)
                rb_ampnorm_base.Offset(6,i).Value = "=OFFSET(INDEX(" & rb_ampnorm_base.Offset(17,i).Address & ":" & rb_ampnorm_base.Offset(17+(num_cycles.Value-1),i).Address & ",MATCH(" & rb_ampnorm_base.Offset(5,i).Address & "," & rb_ampnorm_base.Offset(17,i).Address & ":" & rb_ampnorm_base.Offset(17+(num_cycles.Value-1),i).Address & ",0)),-1,0)"
                ' X2, Formula =OFFSET($A$18,MATCH(E96,E53:E92,0)-1,0)
                rb_ampnorm_base.Offset(7,i).Value = "=OFFSET($A$18,MATCH(" & rb_ampnorm_base.Offset(5,i).Address & "," & rb_ampnorm_base.Offset(17,i).Address & ":" & rb_ampnorm_base.Offset(17+(num_cycles.Value-1),i).Address & ",0)-1,0)"
                ' X1, Formula =OFFSET($A$18,MATCH(E97,E53:E92,0)-1,0)
                rb_ampnorm_base.Offset(8,i).Value = "=OFFSET($A$18,MATCH(" & rb_ampnorm_base.Offset(6,i).Address & "," & rb_ampnorm_base.Offset(17,i).Address & ":" & rb_ampnorm_base.Offset(17+(num_cycles.Value-1),i).Address & ",0)-1,0)"
                ' Slope, Formula =(E96-E97)/(E98-E99)
                rb_ampnorm_base.Offset(9,i).Value = "=(" & rb_ampnorm_base.Offset(5,i).Address & "-" & rb_ampnorm_base.Offset(6,i).Address & ")/(" & rb_ampnorm_base.Offset(7,i).Address & "-" &  rb_ampnorm_base.Offset(8,i).Address & ")"
                ' b, Formula =E96-(E100*E98)
                rb_ampnorm_base.Offset(10,i).Value = "=" & rb_ampnorm_base.Offset(5,i).Address & "-(" & rb_ampnorm_base.Offset(9,i).Address & "*" & rb_ampnorm_base.Offset(7,i).Address & ")"               
                ' CT, Formula =IFERROR(($Y$4-E101)/E100,"Infinity")
                rb_ampnorm_base.Offset(11,i).Value = "=IFERROR((Data_Review!" & rthreshold.Address & "-" & rb_ampnorm_base.Offset(10,i).Address & ")/" & rb_ampnorm_base.Offset(9,i).Address & ",""Infinity"")"                            
                i=i+1 'Next Column
            Next di
            ' Average for group, Formula =IFERROR(AVERAGE(B12:C12),"No TC")
            rb_ampnorm_base.Offset(12,i-1).Value = "=IFERROR(AVERAGE(" & rb_ampnorm_base.Offset(11,(i-datas)).Address & ":" & rb_ampnorm_base.Offset(11,i-1).Address & "),""No TC"")"
            ' STDEV for group, Formula =IFERROR(STDEV(B12:C12),"No TC")
            rb_ampnorm_base.Offset(13,i-1).Value = "=IFERROR(STDEV(" & rb_ampnorm_base.Offset(11,(i-datas)).Address & ":" & rb_ampnorm_base.Offset(11,i-1).Address & "),""N/A"")"
            g = g +1
        Next ci
End Sub

Sub categorize_data ()
    Dim i As Integer 'Counter for 1 to 96 data wells
    Dim cc As Integer 'Counter for categories in groups
    Dim matchfound As Boolean
    Dim current_group As Worksheet
    Dim current_name As Range
    Dim ex_groups As Integer
    Dim ex_entries As Integer

    
    i = 0

    For i = 0 To 95 'Cycle through columns on the main sheet
        matchfound = False

     'Determine grouping sheet
        If rb_sampletype.Offset(0, i).Value <> "Empty" Then
            Select Case rb_sampletype.Offset(0, i).Value
                Case "Blank"
                    Set current_group = ws_blanks
                    Set current_name = rb_samplename
                Case "Standard"
                    Set current_group = ws_standards
                    Set current_name = rb_concentration
                Case "Unknown"
                    Set current_group = ws_unknowns
                    Set current_name = rb_samplename
            End Select
            

            'Check how many groups already exist in the sheet
            ex_groups = row_count(current_group.Range("A1"))
            'Check if this is the first entry
            If ex_groups = 0 Then
               current_name.Offset(0, i).Copy Destination:=current_group.Cells(1, 1)
               current_group.Cells(1, 2).Value = i+1
            Else
                'If not, cycle through existing categories in the group to find a match
                cc = 1
                For cc = 1 To ex_groups 'Cycle through existing categories in group sheets
                    ex_entries = 0
                    If current_name.Offset(0, i).Value = current_group.Cells(cc, 1) Then
                        'Count number of entries in a group
                        ex_entries = col_count(current_group.Range("A"&cc))-1
                        current_group.Cells(cc, ex_entries + 2).Value = i+1
                        matchfound = True
                        Exit For
                    End If
                Next cc
                    
                'If match not found, make new group
                If matchfound = False Then
                    current_name.Offset(0, i).Copy Destination:=current_group.Cells(ex_groups + 1, 1)
                    current_group.Cells(ex_groups + 1, 2).Value = i+1
                End If
            End If
        End If 'If not empty
    Next i

    'Sort Blanks
    ws_blanks.UsedRange.Sort key1:=ws_blanks.Cells(1,1), order1:=xlAscending, header:=xlNo
    'Sort Standards
    ws_standards.UsedRange.Sort key1:=ws_standards.Cells(1,1), order1:=xlAscending, header:=xlNo
    'Sort Unknowns
    If sort_unknowns = "Yes" Then
        ws_unknowns.UsedRange.Sort key1:=ws_unknowns.Cells(1,1), order1:=xlAscending, header:=xlNo
    End If
End Sub

Sub cell_changer(target As Range)
    Select Case target.Value
        Case "Blank"
            With target.Offset(1, 0) 'Concentration
                .Value = 0 
                .NumberFormat = "0"
                .Font.Name = "Calibri"
                .Font.Size = 9
                .HorizontalAlignment = xlCenter
            End With
            Call lock_cell(target.Offset(1, 0), 1) 'Concentration
            target.Offset(2, 0).HorizontalAlignment = xlCenter 'Sample Name'
            If target.Offset(2, 0).Value = "N/A" Then 'Sample Name'
                target.Offset(2, 0).Value = vbNullString 'Sample Name'
            End If
        Case "Standard"
            With target.Offset(1, 0) 'Concentration
                .Value = 0 
                .NumberFormat = "0.00E+00"
                .Font.Name = "Calibri"
                .Font.Size = 9
                .HorizontalAlignment = xlCenter
            End With
            Call lock_cell(target.Offset(1, 0), 0) 'Concentration
            Call lock_cell(target.Offset(2, 0), 1) 'Sample Name'
            target.Offset(2, 0).HorizontalAlignment = xlCenter
            If target.Offset(1, 0).Value = "N/A" Then 'Concentration
                target.Offset(1, 0).Value = vbNullString 'Concentration
            End If
        Case "Unknown"
            Call lock_cell(target.Offset(2, 0), 0) 'Sample Name'
            Call lock_cell(target.Offset(1, 0), 1) 'Concentration
            target.Offset(1, 0).Value = "N/A" 'Concentration
            If target.Offset(2, 0).Value = "N/A" Then 'Sample Name'
                target.Offset(2, 0).Value = vbNullString 'Sample Name'
            End If
        Case "Empty"
            target.Offset(1, 0).Value = "N/A" 'Concentration
            target.Offset(2, 0).Value = "N/A" 'Sample Name'
            Call lock_cell(target.Offset(1, 0), 1) 'Concentration
            Call lock_cell(target.Offset(2, 0), 1) 'Sample Name'
    End Select
End Sub

Sub chart_spinner (chartname As String, ws As Worksheet)
    Dim datas As Integer
    Dim subrange As Range
    Dim c As Integer
    Dim mychart As Chart

    Application.ScreenUpdating = False

    Set mychart = ws.ChartObjects(chartname).chart
    ' spincounter.Value

    datas = Application.WorksheetFunction.CountA(rb_ampnorm_groups)
    Set subrange = Range(rb_ampnorm_groups.Cells(1,1), rb_ampnorm_groups.Cells(1,1).Offset(0,datas-1))
    Call clear_chart(chartname, ws)

    For c=0 To subrange.Cells.Count - 1  
        If subrange.Cells(1, c+1).Value = spincounter.Value Then
            With mychart
                With .Axes(xlCategory, xlPrimary)
                    .MaximumScale = num_cycles.Value+1
                    .MinimumScale = 1
                End With

                With .SeriesCollection.NewSeries
                    .XValues = ws_ampnorm.Range(ws_ampnorm.Range("A18"),ws_ampnorm.Range("A18").Offset(num_cycles.Value,0))
                    .Values = ws_ampnorm.Range(ws_ampnorm.Range("B18").Offset(0,c),ws_ampnorm.Range("B18").Offset(num_cycles.Value,c))
                    .Name = ws_ampnorm.Range("B3").Offset(0,c)
                    .MarkerStyle = 0
                End With
            End With
        End If            
    Next c

    Application.ScreenUpdating = True
End Sub

Function check_blanks(base As Range) As Boolean
    Dim i As Integer

    For i = 1 To 96
        If Len(Trim(base.Offset(0, i))) = 0 Then
            check_blanks = False
            Exit Function
        Else
            check_blanks = True
        End If
    Next i
End Function

Function check_validation(r As String) As Boolean
 '   Returns True if every cell in Range r uses Data Validation
    Dim ws As Worksheet
    Dim x As Variant
    Set ws = ActiveWorkbook.ActiveSheet

    On Error Resume Next
    x = ws.Range(r).Validation.Type
    If Err.Number = 0 Then check_validation = True Else check_validation = False
End Function

Sub clear_chart(chartname As String, ws As Worksheet)
    Dim n As Integer
    Dim mychart As Chart

    Set mychart = ws.ChartObjects(chartname).chart  
        With mychart
            For n=.SeriesCollection.Count To 1 Step -1 
                If (.SeriesCollection(n).Name) <> "Threshold" Then
                    .SeriesCollection(n).Delete
                End If
            Next n
        End With    
End Sub

Function count_data(rbase As Range, total_cells As Integer) As Integer
 'Counts cells in Sample Type row that are not empty    
    Dim i As Integer
    count_data = 0
    For i=0 to total_cells-1
        If rbase.Offset(0, i).Value <> "Empty" Then
            count_data = count_data + 1
        End If
    Next i
End Function

Function count_incategory (rbase As Worksheet) As Integer
    Dim categories As Integer, i As Integer

    count_incategory = 0
    categories = row_count(rbase.Cells(1,1))
    For i=0 To categories-1 
        count_incategory = count_incategory + (col_count(rbase.Cells(1,1).Offset(i,0))-1)
    Next i    
End Function

Function count_cycles(target As Range) As Integer
 'Counts number of cycles in the input file
    count_cycles = Range(target, target.End(xlDown)).Count
End Function

Sub fix_emptycolumns()
    Dim i As Integer

    For i = 1 To 96
        If Len(Trim(rb_firstcycle.Offset(0, i))) = 0 Then
            rb_sampletype.Offset(0, i).Value = "Empty"
        End If
    Next i
End Sub

Sub populate_charts (chartname As String, ws As Worksheet, cats As String)
    'Chart name, Worsheet, and What Category to populate with
    Dim n As Integer, b As Integer 'Counters
    Dim mychart As Chart
    Dim cgroup As Integer
    Dim rr As Integer, rg As Integer, rb As Integer
    cgroup = 0

    Set mychart = ws.ChartObjects(chartname).chart  

    Call clear_chart(chartname, ws)
    For b=0 To total_samples.Value-1
            If (ws_ampnorm.Range("B1").Offset(0,b) = cats) Then
            'Check group

                cgroup = rb_ampnorm_base.Offset(15,b)
                Call color_convert(cgroup, rr, rg, rb)

                With mychart

                    With .Axes(xlCategory, xlPrimary)
                        .MaximumScale = num_cycles.Value+1
                        .MinimumScale = 1
                    End With

                    With .SeriesCollection.NewSeries
                        .XValues = ws_ampnorm.Range(ws_ampnorm.Range("A18"),ws_ampnorm.Range("A18").Offset(num_cycles.Value,0))
                        .Values = ws_ampnorm.Range(ws_ampnorm.Range("B18").Offset(0,b),ws_ampnorm.Range("B18").Offset(num_cycles.Value,b))
                        .Name = ws_ampnorm.Range("B3").Offset(0,b)
                        .MarkerStyle = 0
                        .Format.Line.ForeColor.RGB = RGB(rr, rg, rb)
                    End With
                End With
        End If
    Next b
End Sub

Sub reset_styles()
    Dim cws As Worksheet
    On Error GoTo ErrHandler:
 '============================Conditional formatting
 ' Delete all format conditions on the sheet
    With ws_dataimport.Cells
        .FormatConditions.Delete
    End With

    'Styles to sample Type
    With trigger_range1
        'Standard
        .FormatConditions.Add Type:=xlTextString, String:="Standard", TextOperator:=xlContains
        
        With .FormatConditions(1).Font
            .ThemeColor = xlThemeColorLight2
            .TintAndShade = -0.249946592608417
        End With

        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.399945066682943
        End With
            
        'Blank
        .FormatConditions.Add Type:=xlTextString, String:="Blank", TextOperator:=xlContains
        With .FormatConditions(2).Font
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = -0.499984740745262
        End With

        With .FormatConditions(2).Interior
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = 0.399945066682943
        End With
        
        'Unknown
        .FormatConditions.Add Type:=xlTextString, String:="Unknown", TextOperator:=xlContains            
        With .FormatConditions(3).Font
         .ThemeColor = xlThemeColorAccent3
         .TintAndShade = -0.499984740745262
        End With
        
        With .FormatConditions(3).Interior
         .PatternColorIndex = xlAutomatic
         .ThemeColor = xlThemeColorAccent3
         .TintAndShade = 0.399945066682943
        End With
        
        'Empty
        .FormatConditions.Add Type:=xlTextString, String:="Empty", TextOperator:=xlContains        
        With .FormatConditions(4).Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.499984740745262
        End With
        
        With .FormatConditions(4).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.14996795556505
        End With
    End With
    
    With Range(trigger_range2, trigger_range3)
        
        'N/A
        .FormatConditions.Add Type:=xlTextString, String:="N/A", TextOperator:=xlContains
        With .FormatConditions(1).Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.499984740745262
        End With
            
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.14996795556505
        End With
    End With

 '============================Data Validation
    Set cws = ActiveSheet 

    Worksheets(ws_dataimport.Name).Activate
    With ws_dataimport.Cells
        .Validation.Delete
    End With

    ws_dataimport.Range("A11:CS11").Validation.Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Operator:=xlEqual, Formula1:="=False"

    trigger_range1.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= xlBetween, Formula1:="=hidden_options!$A$2:$A$5"

    sort_unknowns.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= xlBetween, Formula1:="=hidden_options!$B$2:$B$3"
    force_display.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= xlBetween, Formula1:="=hidden_options!$B$2:$B$3"
    Worksheets(cws.Name).Activate
    Exit Sub 
    
    ErrHandler:
    MsgBox "Sheet failed to initialize properly." & vbCrLf &  "Please enable Macros and click the Reset button."
End Sub

Sub reset_silent ()
 'Procedure resets and re-initializes the entire sheet
    Dim cur_ws As Worksheet
    Set cur_ws = ActiveSheet 
    Call Initialize_vars

    'Clear Sheets
    ws_blanks.Cells.Clear
    ws_standards.Cells.Clear
    ws_unknowns.Cells.Clear 
    'Call clear_right_down(ws_ampnorm.Range("B1"))
    Call clear_right(ws_ampnorm.Range("B1:B16"))
    Call clear_right_down(rb_ampnorm)
    Call clear_right_down(rb_ampnorm_stcurve)
    Call clear_right_down(rb_ampnorm_unknowns)
    Call clear_helper


    'Clear Charts
    Call clear_chart("ch_blanks", ws_ampnorm_datareview)
    Call clear_chart("ch_blanks_l", ws_ampnorm_datareview)
    Call clear_chart("ch_standards", ws_ampnorm_datareview)
    Call clear_chart("ch_standards_l", ws_ampnorm_datareview)
    Call clear_chart("ch_unknowns", ws_ampnorm_datareview)
    Call clear_chart("ch_unknowns_l", ws_ampnorm_datareview)
    Call clear_chart("ch_group", ws_ampnorm_datareview)
    Call clear_chart("ch_group_l", ws_ampnorm_datareview)

    Worksheets(cur_ws.Name).Activate
End Sub

Sub toggle_normalization_all ()
    Dim datas As Integer
    Dim subrange As Range
    Dim c As Integer

    Application.ScreenUpdating = False
    datas = Application.WorksheetFunction.CountA(rb_ampnorm_groups)
    Set subrange = Range(rb_ampnorm_groups.Cells(1,1).Offset(-12,0), rb_ampnorm_groups.Cells(1,1).Offset(-12,datas-1))

    For c=0 To subrange.Cells.Count - 1  
        If subrange.Cells(1, c+1).Value = 1 Then
            subrange.Cells(1, c+1).Value = 0
        Else
            If subrange.Cells(1, c+1).Offset(-3,0).Value = "Blank" Then
                subrange.Cells(1, c+1).Value = 0
            Else
                subrange.Cells(1, c+1).Value = 1
            End If
        End If            
    Next c
    Application.ScreenUpdating = True
End Sub

Sub st_curve()
    Dim i As Integer
    Dim cats as Integer, datas As Integer, lastd As Integer
    Dim refa As String, refs As String, refn As String

    ' Standards
        cats = row_count(ws_standards.Range("A1"))

    For i=0 To cats-1
        datas = col_count(ws_standards.Cells(1,1).Offset(i,0))-1
        lastd = ws_standards.Cells(1,1).Offset(i,datas).Value

        Call frame_cell_thin(rb_ampnorm_stcurve.Offset(i,0))
        'Concentration
        rb_ampnorm_stcurve.Offset(i,0).Value = "=hidden_group_standards!" & ws_standards.Cells(i+1,1).Address
        
        With rb_ampnorm_stcurve.Offset(i,1)
        '# Molecules Formula =A38*(Data_Import!B5*1E-6)*6.022E23
        .Value = "=" & rb_ampnorm_stcurve.Offset(i,0).Address(False,False) & "*(Data_Import!"& sample_volume.Address &"*1E-6)*6.022E23"
        .NumberFormat = "0.00E+00"
        .HorizontalAlignment = xlCenter
        End With
        Call frame_cell_thin(rb_ampnorm_stcurve.Offset(i,1))

        With rb_ampnorm_stcurve.Offset(i,2)
        'Log # Molecules Formula =log10(B38)
        .Value = "=log10(" & rb_ampnorm_stcurve.Offset(i,1).Address(False,False) & ")"
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlCenter
        End With
        Call frame_cell_thin(rb_ampnorm_stcurve.Offset(i,2))

        refa = "=analysis_ampnorm!" & rb_ampnorm_base.Offset(12,WorksheetFunction.Match(lastd, Range(rb_ampnorm_base.Offset(14,0),rb_ampnorm_base.Offset(14,95)),0)-1).Address
        With rb_ampnorm_stcurve.Offset(i,3)
            'Average TC, Formula =analysis_ampnorm!C13
            .Value = refa
            .NumberFormat = "0.0"
            .HorizontalAlignment = xlCenter
        End With
        Call frame_cell_thin(rb_ampnorm_stcurve.Offset(i,3))

        refs = "=analysis_ampnorm!" & rb_ampnorm_base.Offset(13,WorksheetFunction.Match(lastd, Range(rb_ampnorm_base.Offset(14,0),rb_ampnorm_base.Offset(14,95)),0)-1).Address
        With rb_ampnorm_stcurve.Offset(i,4)
            'Error , Formula =analysis_ampnorm!C14
             .Value = refs
             .NumberFormat = "0.0"
            .HorizontalAlignment = xlCenter
        End With
        Call frame_cell_thin(rb_ampnorm_stcurve.Offset(i,4))
    Next i

    'Chart'
    Call clear_chart("ch_ampnorm_stcurve", ws_ampnorm_stcurve)
    With ws_ampnorm_stcurve.ChartObjects("ch_ampnorm_stcurve").chart
        ' With .Axes(xlCategory, xlPrimary)
        '     .MaximumScale = num_cycles.Value+1
        '     .MinimumScale = 1
        ' End With

        With .SeriesCollection.NewSeries
            .XValues = Range(rb_ampnorm_stcurve.Offset(0,2), rb_ampnorm_stcurve.Offset(cats-1,2))
            .Values = Range(rb_ampnorm_stcurve.Offset(0,3), rb_ampnorm_stcurve.Offset(cats-1,3))
            .Name = "Standard Curve"
            .Trendlines.Add
            .HasErrorBars = True
            .ErrorBar Direction:=xlY, Include:=xlBoth, Type:=xlCustom, Amount:=Range(rb_ampnorm_stcurve.Offset(0,4), rb_ampnorm_stcurve.Offset(cats-1,4)), MinusValues:=Range(rb_ampnorm_stcurve.Offset(0,4), rb_ampnorm_stcurve.Offset(cats-1,4))
            .ErrorBar Direction:=xlX, Include:=xlErrorBarIncludeNone, Type:=xlCustom, Amount:=Range(rb_ampnorm_stcurve.Offset(0,4), rb_ampnorm_stcurve.Offset(cats-1,4)), MinusValues:=Range(rb_ampnorm_stcurve.Offset(0,4), rb_ampnorm_stcurve.Offset(cats-1,4))
        End With

        With .Axes(xlCategory, xlPrimary)
            .MaximumScale = rb_ampnorm_stcurve.Offset(cats-1,2).Value+0.5
            .MinimumScale = 0.5
        End With

        'Blanks
        cats = row_count(ws_blanks.Range("A1"))
        For i=0 To cats-1
            datas = col_count(ws_blanks.Cells(1,1).Offset(i,0))-1
            lastd = ws_blanks.Cells(1,1).Offset(i,datas).Value
            refa = "analysis_ampnorm!" & rb_ampnorm_base.Offset(12,WorksheetFunction.Match(lastd, Range(rb_ampnorm_base.Offset(14,0),rb_ampnorm_base.Offset(14,95)),0)-1).Address
            refs = "=analysis_ampnorm!" & rb_ampnorm_base.Offset(13,WorksheetFunction.Match(lastd, Range(rb_ampnorm_base.Offset(14,0),rb_ampnorm_base.Offset(14,95)),0)-1).Address
            refn = "=analysis_ampnorm!" & rb_ampnorm_base.Offset(2,WorksheetFunction.Match(lastd, Range(rb_ampnorm_base.Offset(14,0),rb_ampnorm_base.Offset(14,95)),0)-1).Address

            With .SeriesCollection.NewSeries
                .XValues = Array(0, 30)
                .Values = "=(" & refa & "," & refa & ")" 
                .Name = refn
                .MarkerStyle = 0
                .Format.Line.Visible = msoTrue
                .Format.Line.ForeColor.RGB = RGB(rand_i(20,150), rand_i(20,150), rand_i(20,150))
                .Format.Line.DashStyle = msoLineSysDot
            End With
        Next i

        'Linest =INDEX(LINEST(D38:D42, C38:C42),1,1)
        cats = row_count(ws_standards.Range("A1"))
        ws_ampnorm_stcurve.Range("A36").Value = "=INDEX(LINEST(" & rb_ampnorm_stcurve.Offset(0,3).Address(False,False) & ":" & rb_ampnorm_stcurve.Offset(cats-1,3).Address(False,False) & ", " & rb_ampnorm_stcurve.Offset(0,2).Address(False,False) & ":" & rb_ampnorm_stcurve.Offset(cats-1,2).Address(False,False) & "),1,1)"
        ws_ampnorm_stcurve.Range("B36").Value = "=INDEX(LINEST(" & rb_ampnorm_stcurve.Offset(0,3).Address(False,False) & ":" & rb_ampnorm_stcurve.Offset(cats-1,3).Address(False,False) & ", " & rb_ampnorm_stcurve.Offset(0,2).Address(False,False) & ":" & rb_ampnorm_stcurve.Offset(cats-1,2).Address(False,False) & "),1,2)"
        ws_ampnorm_stcurve.Range("C36").Value = "=INDEX(LINEST(" & rb_ampnorm_stcurve.Offset(0,3).Address(False,False) & ":" & rb_ampnorm_stcurve.Offset(cats-1,3).Address(False,False) & ", " & rb_ampnorm_stcurve.Offset(0,2).Address(False,False) & ":" & rb_ampnorm_stcurve.Offset(cats-1,2).Address(False,False) & ",TRUE, TRUE),3,1)"
    End With
End Sub

Sub color_convert (i As Integer, r As Integer, g As Integer, b As Integer)
    Select Case i
        Case 1 
        r=46 
        g=66 
        b=126
        Case 2 
        r=0 
        g=142 
        b=23
        Case 3 
        r=154 
        g=143 
        b=104
        Case 4 
        r=39 
        g=52 
        b=73
        Case 5 
        r=159 
        g=19 
        b=109
        Case 6 
        r=169 
        g=74 
        b=39
        Case 7 
        r=53 
        g=23 
        b=161
        Case 8 
        r=167 
        g=73 
        b=91
        Case 9 
        r=33 
        g=192 
        b=162
        Case 10 
        r=142 
        g=86 
        b=14
        Case 11 
        r=183 
        g=107 
        b=160
        Case 12 
        r=87 
        g=30 
        b=63
        Case 13 
        r=132 
        g=172 
        b=106
        Case 14 
        r=20 
        g=169 
        b=197
        Case 15 
        r=71 
        g=157 
        b=52
        Case 16 
        r=95 
        g=55 
        b=178
        Case 17 
        r=126 
        g=81 
        b=85
        Case 18 
        r=174 
        g=121 
        b=188
        Case 19 
        r=108 
        g=73 
        b=176
        Case 20 
        r=59 
        g=59 
        b=45
        Case 21 
        r=76 
        g=10 
        b=152
        Case 22 
        r=170 
        g=63 
        b=145
        Case 23 
        r=87 
        g=165 
        b=152
        Case 24 
        r=93 
        g=24 
        b=89
        Case 25 
        r=74 
        g=145 
        b=70
        Case 26 
        r=162 
        g=38 
        b=124
        Case 27 
        r=32 
        g=83 
        b=104
        Case 28 
        r=31 
        g=197 
        b=130
        Case 29 
        r=173 
        g=58 
        b=76
        Case 30 
        r=56 
        g=174 
        b=25
        Case 31 
        r=123 
        g=46 
        b=190
        Case 32 
        r=82 
        g=85 
        b=111
        Case 33 
        r=61 
        g=121 
        b=51
        Case 34 
        r=167 
        g=23 
        b=120
        Case 35 
        r=20 
        g=57 
        b=47
        Case 36 
        r=191 
        g=171 
        b=95
        Case 37 
        r=103 
        g=156 
        b=170
        Case 38 
        r=83 
        g=47 
        b=73
        Case 39 
        r=88 
        g=39 
        b=129
        Case 40 
        r=28 
        g=49 
        b=143
        Case 41 
        r=106 
        g=44 
        b=200
        Case 42 
        r=100 
        g=10 
        b=92
        Case 43 
        r=65 
        g=153 
        b=181
        Case 44 
        r=166 
        g=43 
        b=31
        Case 45 
        r=11 
        g=38 
        b=83
        Case 46 
        r=198 
        g=172 
        b=43
        Case 47 
        r=114 
        g=120 
        b=170
        Case 48 
        r=59 
        g=15 
        b=169
        Case 49 
        r=142 
        g=150 
        b=104
        Case 50 
        r=188 
        g=50 
        b=38
        Case 51 
        r=117 
        g=48 
        b=50
        Case 52 
        r=79 
        g=119 
        b=76
        Case 53 
        r=110 
        g=91 
        b=120
        Case 54 
        r=19 
        g=114 
        b=50
        Case 55 
        r=72 
        g=19 
        b=190
        Case 56 
        r=170 
        g=88 
        b=164
        Case 57 
        r=198 
        g=59 
        b=44
        Case 58 
        r=142 
        g=13 
        b=75
        Case 59 
        r=62 
        g=138 
        b=200
        Case 60 
        r=65 
        g=86 
        b=117
        Case 61 
        r=159 
        g=98 
        b=33
        Case 62 
        r=59 
        g=40 
        b=41
        Case 63 
        r=170 
        g=88 
        b=149
        Case 64 
        r=135 
        g=87 
        b=192
        Case 65 
        r=133 
        g=108 
        b=53
        Case 66 
        r=121 
        g=152 
        b=98
        Case 67 
        r=182 
        g=64 
        b=138
        Case 68 
        r=180 
        g=63 
        b=81
        Case 69 
        r=12 
        g=66 
        b=198
        Case 70 
        r=155 
        g=120 
        b=125
        Case 71 
        r=61 
        g=107 
        b=161
        Case 72 
        r=173 
        g=27 
        b=114
        Case 73 
        r=89 
        g=24 
        b=97
        Case 74 
        r=143 
        g=105 
        b=88
        Case 75 
        r=17 
        g=155 
        b=25
        Case 76 
        r=111 
        g=127 
        b=28
        Case 77 
        r=45 
        g=64 
        b=14
        Case 78 
        r=39 
        g=90 
        b=115
        Case 79 
        r=67 
        g=134 
        b=41
        Case 80 
        r=87 
        g=21 
        b=70
        Case 81 
        r=125 
        g=108 
        b=123
        Case 82 
        r=72 
        g=183 
        b=157
        Case 83 
        r=119 
        g=92 
        b=150
        Case 84 
        r=41 
        g=21 
        b=57
        Case 85 
        r=154 
        g=191 
        b=76
        Case 86 
        r=55 
        g=143 
        b=126
        Case 87 
        r=161 
        g=102 
        b=50
        Case 88 
        r=125 
        g=183 
        b=191
        Case 89 
        r=134 
        g=12 
        b=137
        Case 90 
        r=114 
        g=131 
        b=135
        Case 91 
        r=24 
        g=66 
        b=122
        Case 92 
        r=164 
        g=17 
        b=21
        Case 93 
        r=117 
        g=70 
        b=89
        Case 94 
        r=107 
        g=150 
        b=97
        Case 95 
        r=71 
        g=24 
        b=89
        Case 96 
        r=67 
        g=148 
        b=59
        Case Else 
        r=0
        g=0
        b=0    
    End Select
End Sub 

Sub populate_unknowns()
    Dim i As Integer
    Dim cats as Integer, datas As Integer, lastd As Integer
    Dim refa As String, refs As String, refn As String

    ' Standards
        cats = row_count(ws_unknowns.Range("A1"))

    For i=0 To cats-1
        datas = col_count(ws_unknowns.Cells(1,1).Offset(i,0))-1
        lastd = ws_unknowns.Cells(1,1).Offset(i,datas).Value

        'Sample Number
        With rb_ampnorm_unknowns.Offset(0,i)
        .Value = i+1
        .HorizontalAlignment = xlCenter
        End With
        Call frame_cell_thin(rb_ampnorm_unknowns.Offset(0,i))

        'Sample Name
        With rb_ampnorm_unknowns.Offset(1,i)
        .Interior.Color = RGB(196,215,155)
        .Value = "=hidden_group_unknowns!" & ws_unknowns.Cells(i+1,1).Address
        .HorizontalAlignment = xlCenter
        End With
        Call frame_cell_thin(rb_ampnorm_unknowns.Offset(1,i))

        'Average TC, Formula =analysis_ampnorm!C13
        With rb_ampnorm_unknowns.Offset(2,i)
            .Value = "=analysis_ampnorm!" & rb_ampnorm_base.Offset(12,WorksheetFunction.Match(lastd, Range(rb_ampnorm_base.Offset(14,0),rb_ampnorm_base.Offset(14,95)),0)-1).Address
            .NumberFormat = "0.0"
            .HorizontalAlignment = xlCenter
        End With
        Call frame_cell_thin(rb_ampnorm_unknowns.Offset(2,i))
            
        '# Molecules, Formula =IF(D78="No TC", "X", (10^(((D78-Standard_Curve!$B$36)/Standard_Curve!$A$36))))
        With rb_ampnorm_unknowns.Offset(3,i)
            .Value = "=IF("& rb_ampnorm_unknowns.Offset(2,i).Address(False,False) &"=""No TC"", ""X"",(10^((("& rb_ampnorm_unknowns.Offset(2,i).Address(False,False) &"-Standard_Curve!$B$36)/Standard_Curve!$A$36))))"
            .NumberFormat = "0.00E+00"
            .HorizontalAlignment = xlCenter
        End With
        Call frame_cell_thin(rb_ampnorm_unknowns.Offset(3,i))

        'Percent Error, Formula = IF(D78="N/A", "X",analysis_ampnorm!E14/analysis_ampnorm!E13)
        With rb_ampnorm_unknowns.Offset(6,i)
            .Value = "=IF("& rb_ampnorm_unknowns.Offset(2,i).Address(False,False) &"=""N/A"", ""X"",analysis_ampnorm!" & rb_ampnorm_base.Offset(13,WorksheetFunction.Match(lastd, Range(rb_ampnorm_base.Offset(14,0),rb_ampnorm_base.Offset(14,95)),0)-1).Address & "/analysis_ampnorm!" & rb_ampnorm_base.Offset(12,WorksheetFunction.Match(lastd, Range(rb_ampnorm_base.Offset(14,0),rb_ampnorm_base.Offset(14,95)),0)-1).Address & ")"
            .NumberFormat = "0.00%"
            .HorizontalAlignment = xlCenter
        End With
        Call frame_cell_thin(rb_ampnorm_unknowns.Offset(6,i))

        'Concentration, Formula =IF(D78="No TC", "X",C3/((Data_Import!B5*1E-6)*6.022E23))
        With rb_ampnorm_unknowns.Offset(4,i)
            .Value = "=IF("& rb_ampnorm_unknowns.Offset(2,i).Address(False,False) &"=""No TC"", ""X""," & rb_ampnorm_unknowns.Offset(3,i).Address(False,False) & "/((Data_Import!"& sample_volume.Address &"*1E-6)*6.022E23))"
            .NumberFormat = "0.00E+00"
            .HorizontalAlignment = xlCenter
        End With
        Call frame_cell_thin(rb_ampnorm_unknowns.Offset(4,i))

        'Concentration Error, Formula =IF(D78="No TC","X",C4*C6)
        With rb_ampnorm_unknowns.Offset(5,i)
            .Value = "=IF("& rb_ampnorm_unknowns.Offset(2,i).Address(False,False) &"=""N/A"",""X""," & rb_ampnorm_unknowns.Offset(4,i).Address(False,False) & "*" & rb_ampnorm_unknowns.Offset(6,i).Address(False,False) & ")"
            .NumberFormat = "0.00E+00"
            .HorizontalAlignment = xlCenter
        End With
        Call frame_cell_thin(rb_ampnorm_unknowns.Offset(5,i))
    Next i

    'Chart'
    Call clear_chart("ch_unknowns_results", ws_ampnorm_unknowns)
    With ws_ampnorm_unknowns.ChartObjects("ch_unknowns_results").chart

        With .SeriesCollection.NewSeries
            .Values = Range(rb_ampnorm_unknowns.Offset(4,0), rb_ampnorm_unknowns.Offset(4,cats-1))
            .HasErrorBars = True
            .ErrorBar Direction:=xlY, Include:=xlBoth, Type:=xlCustom, Amount:=Range(rb_ampnorm_unknowns.Offset(5,0), rb_ampnorm_unknowns.Offset(5,cats-1)), MinusValues:=Range(rb_ampnorm_unknowns.Offset(5,0), rb_ampnorm_unknowns.Offset(5,cats-1))
        End With
        
 
    End With
    ws_ampnorm_unknowns.Cells.EntireColumn.AutoFit
End Sub

Sub transpose_paste ()
    On Error GoTo ErrHandler:
    ws_helper.Range("A17").PasteSpecial Transpose:=True
    Exit Sub
    ErrHandler:
    MsgBox  "The Clipboard is empty!"& vbCrLf & "Please copy the range to be Transposed first!"
    Call Initialize_vars
End Sub

Sub standards_generator ()
    Dim  rep As Integer,  i As Integer, j As Integer, c As Integer, z As Integer, stps As Integer
    Dim sn As Double, en As Double, mult As Double


    sn = ws_helper.Range("B2").Value
    en = ws_helper.Range("C2").Value
    mult = ws_helper.Range("D2").Value
    rep = ws_helper.Range("E2").Value
    c = 0
    z = 0

    stps = Round(Application.WorksheetFunction.Log((en/sn),mult))
    
    If stps < 1 Then
        MsgBox "Your Multiplier Step does not make sense!"
        Exit Sub
    End If
    
    For i=0 To stps
        For j=0 To rep-1
            ws_helper.Range("A3").Offset(0,c).Value = sn*(mult)^i
            c = c +1
        Next j        
    Next i 
End Sub

Sub unknown_generator ()
    Dim labl As String, zs As String
    Dim sn As Integer, en As Integer, rep As Integer, i As Integer, j As Integer, c As Integer, z As Integer


    labl = ws_helper.Range("B7").Value
    sn = ws_helper.Range("C7").Value
    en = ws_helper.Range("D7").Value
    rep = ws_helper.Range("E7").Value
    c = 0
    z = 0

    For i=sn To en
        zs = ""
        If ws_helper.Range("F7").Value = "Yes" Then
            If Len(CStr(i)) < Len(CStr(en)) Then
                For z=0 To Len(CStr(en))-Len(CStr(i))-1
                    zs= zs & "0"
                Next z
            End If
        End If 
        For j=0 To rep-1
            ws_helper.Range("A8").Offset(0,c).Value = labl & " " & zs & i
            c = c +1
        Next j        
    Next i 
End Sub

Sub clear_helper ()
    Call clear_right(ws_helper.Range("A3"))
    Call clear_right(ws_helper.Range("A8"))
    Call clear_right_down(ws_helper.Range("A17"))
End Sub