Option Explicit

'========================== Global variable declaration
'Sheets
Public ws_dataimport As Worksheet
Public ws_options As Worksheet
Public ws_blanks As Worksheet
Public ws_standards As Worksheet
Public ws_unknowns As Worksheet
Public ws_ampnorm As Worksheet
Public ws_ampnorm_datareview As Worksheet
Public ws_ampnorm_stcurve As Worksheet
Public ws_ampnorm_unknowns As Worksheet
Public ws_helper As Worksheet


'Ranges
Public rb_import As Range
Public rb_sampletype As Range
Public rb_concentration As Range
Public rb_samplename  As Range
Public rb_coordinate As Range
Public rb_firstcycle As Range

Public rb_ampnorm As Range
Public rb_ampnorm_base As Range
Public rb_ampnorm_groups As Range
Public rb_import_base As Range
Public rb_ampnorm_stcurve As Range
Public rb_ampnorm_unknowns As Range

Public sample_volume As Range
Public sort_unknowns As Range
Public force_display As Range
Public rthreshold As Range
Public spincounter As Range
Public trigger_range1 As Range
Public trigger_range2 As Range
Public trigger_range3 As Range

'Printed Variables
Public num_cycles As Range
Public max_samples As Range
Public total_samples As Range
Public total_groups As Range



'========================== Interface Level

Sub Initialize_vars()
    On Error Goto ErrHandler:
    'Initialize sheets
    Set ws_dataimport = ThisWorkbook.Sheets("Data_Import")
    Set ws_options = ThisWorkbook.Sheets("hidden_options")
    Set ws_blanks = ThisWorkbook.Sheets("hidden_group_blanks")
    Set ws_standards = ThisWorkbook.Sheets("hidden_group_standards")
    Set ws_unknowns = ThisWorkbook.Sheets("hidden_group_unknowns")
    Set ws_ampnorm = ThisWorkbook.Sheets("analysis_ampnorm")
    Set ws_ampnorm_datareview = ThisWorkbook.Sheets("Data_Review")
    Set ws_ampnorm_stcurve = ThisWorkbook.Sheets("Standard_Curve")
    Set ws_ampnorm_unknowns = ThisWorkbook.Sheets("Unknowns")
    Set ws_helper = ThisWorkbook.Sheets("Import_Helper")

    'Initialize Ranges
    Set rb_import = ws_dataimport.Range("A11")
    Set rb_sampletype = ws_dataimport.Range("B12")
    Set rb_concentration = ws_dataimport.Range("B13")
    Set rb_samplename  = ws_dataimport.Range("B14")
    Set rb_coordinate = ws_dataimport.Range("B15")
    Set rb_firstcycle = ws_dataimport.Range("B16")

    Set rb_ampnorm = ws_ampnorm.Range("A17")
    Set rb_ampnorm_base = ws_ampnorm.Range("B1")
    Set rb_ampnorm_groups = ws_ampnorm.Range("B16:CS16")
    Set rb_import_base = ws_dataimport.Range("B12")
    Set rb_ampnorm_stcurve = ws_ampnorm_stcurve.Range("A38")
    Set rb_ampnorm_unknowns = ws_ampnorm_unknowns.Range("B1")


    Set sample_volume = ws_dataimport.Range("B5")
    Set sort_unknowns = ws_dataimport.Range("B6")
    Set force_display = ws_dataimport.Range("B7")
    Set rthreshold = ws_ampnorm_datareview.Range("B1")
    Set spincounter = ws_ampnorm_datareview.Range("R8")
    Set trigger_range1 = ws_dataimport.Range("B12:CS12")
    Set trigger_range2 = ws_dataimport.Range("B13:CS13")
    Set trigger_range3 = ws_dataimport.Range("B14:CS14")

    Set num_cycles = ws_options.Range("A15") 
    Set max_samples = ws_options.Range("B15") 
    Set total_samples = ws_options.Range("C15")
    Set total_groups = ws_options.Range("D15")
   

    'Reset Styles
    Call reset_styles

    'Establish Default Variables
    num_cycles.Value = 0
    max_samples.Value = 96
    total_samples.Value = 0
    total_groups.Value = 0
    spincounter.Value = 1
    Exit Sub

    ErrHandler:
    MsgBox "Sheet failed to initialize properly." & vbCrLf &  "Please enable Macros and click the Reset button."
End Sub

Sub Reset()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    Call reset_silent


    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True 

    MsgBox "Reset Complete!"
End Sub

Private Sub Test()
    'Call Initialize_vars 
End Sub

Sub Data_import()
    'Variables
    Dim i As Integer, g as Integer
    Dim cur_ws As Worksheet

    On Error GoTo ErrHandler:
    
    Set cur_ws = ActiveSheet 

    'Initialize the sheet
    Call Initialize_vars
    Call reset_silent

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False 

    'Set columns that don't contain any data to Empty
    Call fix_emptycolumns

    'Check for blanks in Сoncentrations fields
    If check_blanks(rb_concentration) = False Then
        MsgBox ("Please don't leave any fields in the Concentration row as blank.")
        Exit Sub
    End If

    'Check for blanks in Sample Name fields
    If check_blanks(rb_samplename) = False Then
        MsgBox ("Please don't leave any fields in the Sample Name row as blank.")
        Exit Sub
    End If

    total_samples.Value = count_data(rb_sampletype, 96)
    num_cycles.Value = count_cycles(rb_firstcycle)

    'Sort data into groups
    Call categorize_data
    
    total_groups.Value = row_count(ws_blanks.Cells(1,1)) + row_count(ws_standards.Cells(1,1)) + row_count(ws_unknowns.Cells(1,1)) 

    'Copy Cycles to Amplitude Normalization Sheet
    ws_dataimport.Range(rb_coordinate.Offset(0,-1), rb_coordinate.Offset(num_cycles.Value,-1)).Copy 
    rb_ampnorm.PasteSpecial xlPasteValues

    'Reference Data to Amplitude Normalization Sheet
    i = 0
    g = 0
    Call amp_normalization(ws_blanks, i,g)
    Call amp_normalization(ws_standards, i,g)
    Call amp_normalization(ws_unknowns, i, g)
    i = 0
    g = 0

    Call populate_charts("ch_blanks", ws_ampnorm_datareview, "Blank")
    Call populate_charts("ch_blanks_l", ws_ampnorm_datareview, "Blank")
    Call populate_charts("ch_standards", ws_ampnorm_datareview, "Standard")
    Call populate_charts("ch_standards", ws_ampnorm_stcurve, "Standard")
    Call populate_charts("ch_standards_l", ws_ampnorm_datareview, "Standard")
    If force_display.Value = "Yes" Then
        Call populate_charts("ch_unknowns", ws_ampnorm_datareview, "Unknown")
        Call populate_charts("ch_unknowns_l", ws_ampnorm_datareview, "Unknown")
    End If
    Call chart_spinner("ch_group", ws_ampnorm_datareview)
    Call chart_spinner("ch_group_l", ws_ampnorm_datareview)

    With spincounter.Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="1", Formula2:="=hidden_options!D15"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    'AmpNorm Standard Cuve 
    Call st_curve()
    Call populate_unknowns()

    Worksheets(cur_ws.Name).Activate

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True 

    MsgBox "Analysis Complete"
    Exit Sub

    ErrHandler:
    MsgBox "An Error has occured." & vbCrLf & "Please try Resetting the sheet."
End Sub

Sub Jump_group()
    Call chart_spinner("ch_group", ws_ampnorm_datareview)
    Call chart_spinner("ch_group_l", ws_ampnorm_datareview)
End Sub
