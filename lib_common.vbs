Option Explicit

Function col_char(iCol As Integer) As String
 'Function returns the letter associated with a particular column number
      col_char = Split(Cells(1, iCol).Address, "$")(1)
End Function

Function col_count(base_cell As Range) As Integer
 'Count number of used cells in a row (starts from column 1)
    Dim row_num As Integer
    Dim ws As Worksheet

    Set ws = base_cell.Parent
    row_num = base_cell.row
    col_count = Application.WorksheetFunction.CountA(ws.Range(row_num & ":" & row_num)) 
End Function

Function row_count(base_cell As Range) As Integer
 'Count number of used cells in a row (starts from column 1)
    Dim col_num As Integer
    Dim col_let As String
    Dim ws As Worksheet

    Set ws = base_cell.Parent
    col_num = base_cell.column
    col_let = col_char(col_num)
    row_count = Application.WorksheetFunction.CountA(ws.Range(col_let & ":" & col_let)) 
End Function

Sub clear_right_down (rbase As Range)
 'Deletes everything in a contiguous region to right and down of a base cell
    Dim ws As Worksheet
    Dim base_add As String

    Set ws = rbase.Parent
    
    With ws
        .Range(rbase, .Cells(rbase.End(xlDown).Row,rbase.End(xltoRight).Column)).Clear
    End With
End Sub

Sub clear_right (rbase As Range)
 'Deletes everything in a contiguous region to right and down of a base cell
    Dim ws As Worksheet
    Dim base_add As String

    Set ws = rbase.Parent

    With ws
        .Range(rbase, .Cells(rbase.Row,rbase.End(xltoRight).Column)).Clear
    End With
End Sub

Sub lock_cell(target As Range, state As Boolean)
 'Sub uses Validation False Formula to prevent cell editing. 1 - Locked, 0 - Unlocked
 'The cell value can still be changed programatically, or by copy/pasting, so it is not fool proof.
    If state = True Then
        With target
            With .Validation
                .Delete
                .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Operator:=xlEqual, Formula1:="=False"
            End With
        End With
    ElseIf state = False Then
        target.Validation.Delete
    End If
End Sub

Function rand_i (low As Integer, up As Integer)
    rand_i = Int((up - low + 1) * Rnd + low)
End Function

Sub frame_cell_thin(r As Range)
    With r
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone

        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
End Sub

Function metric_prefix(num_val As Double, decimals As Integer) As String
 ' Funtion takes in a number and returns a number with a metrix prefix
     Dim exponent As Integer
     Dim num_holder As Double
     Dim revalue As Double
     Dim prefix As String
     Dim sign As Integer
 
    sign = 1
    num_holder = num_val
    
    If num_val = 0 Then
        metric_prefix = "0"
        Exit Function
    ElseIf num_val < 0 Then
        sign = -1
        num_val = Abs(num_val)
        num_holder = num_val
    End If
 
    If num_val >= 1 Then
        exponent = 0
        Do While num_holder >= 1000
            num_holder = num_holder / 1000
            exponent = exponent + 1
        Loop
        revalue = num_val / (1000 ^ exponent) * sign
    ElseIf num_val < 1 Then
        Do While num_holder < 1
             num_holder = num_holder * 1000
             exponent = exponent - 1
        Loop
        revalue = num_val * (1000 ^ -exponent) * sign
    End If

    revalue = Round(revalue, Abs(decimals))

    Select Case exponent
        Case 0
            prefix = ""
        Case 1
            prefix = "k"
        Case 2
            prefix = "M"
        Case 3
            prefix = "G"
        Case 4
            prefix = "T"
        Case 5
            prefix = "P"
        Case 6
            prefix = "E"
        Case 7
            prefix = "Z"
        Case 8
            prefix = "Y"
        Case Is > 8
            prefix = "humongo"
        Case -1
            prefix = "m"
        Case -2
            prefix = "µ"
        Case -3
            prefix = "n"
        Case -4
            prefix = "p"
        Case -5
            prefix = "f"
        Case -6
            prefix = "a"
        Case -7
            prefix = "z"
        Case -8
            prefix = "y"
        Case Is < -8
            prefix = "insignifico"
    End Select
    metric_prefix = revalue & " " & prefix
End Function

Function log10(inpt As Variant) As Variant
 'Function calculates Log base 10
    log10 = Log(inpt) / Log(10)
End Function

'==================================
' File import functions
'================================== 
Sub validate_filepath(filepath_range As Range)
 'Function corrects folder path if entered incorrectly
    If Right(filepath_range.Value, 1) <> "\" Then
        filepath_range.Value = filepath_range.Value & "\"
    End If
End Sub

Function list_files(file_path As String) As Variant
 'Function gets the list of files in a folder
  
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim file_num As Long
    Dim file_list() As Variant
        
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(file_path)

    'Loop through the Files collection
    file_num = 0
    ReDim file_list(objFolder.Files.Count - 1)
    
    For Each objFile In objFolder.Files
        file_list(file_num) = objFile.Name
        file_num = file_num + 1
    Next
     
     'Clean up!
    Set objFolder = Nothing
    Set objFile = Nothing
    Set objFSO = Nothing
  
    list_files = file_list
End Function


Public Function BASE64SHA1(ByVal sTextToHash As String)

    Dim asc As Object
    Dim enc As Object
    Dim TextToHash() As Byte
    Dim SharedSecretKey() As Byte
    Dim bytes() As Byte
    Const cutoff As Integer = 5

    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA1")

    TextToHash = asc.GetBytes_4(sTextToHash)
    SharedSecretKey = asc.GetBytes_4(sTextToHash)
    enc.Key = SharedSecretKey

    bytes = enc.ComputeHash_2((TextToHash))
    BASE64SHA1 = EncodeBase64(bytes)
    BASE64SHA1 = Left(BASE64SHA1, cutoff)

    Set asc = Nothing
    Set enc = Nothing

End Function

Private Function EncodeBase64(ByRef arrData() As Byte) As String

    Dim objXML As Object
    Dim objNode As Object

    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")

    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.text

    Set objNode = Nothing
    Set objXML = Nothing

End Function
