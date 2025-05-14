' VBScript to convert specific CSV columns to numbers and sum them

Dim fso, file, line, columns, sum1, sum2
Dim columnIndex1, columnIndex2
Dim filePath

' === Configuration ===
filePath = "C:\bots\Sample VB\sample_sales.csv" ' Change this to your actual file path
columnIndex1 = 1 ' First column to sum (0-based index)
columnIndex2 = 2 ' Second column to sum (0-based index)

Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile(filePath, 1)

sum1 = 0
sum2 = 0

' Skip header
If Not file.AtEndOfStream Then
    file.ReadLine
End If

' Read and process each line
Do While Not file.AtEndOfStream
    line = file.ReadLine
    columns = Split(line, ",")

    ' Convert and sum column values
    If IsNumeric(columns(columnIndex1)) Then
        sum1 = sum1 + CDbl(columns(columnIndex1))
    End If

    If IsNumeric(columns(columnIndex2)) Then
        sum2 = sum2 + CDbl(columns(columnIndex2))
    End If
Loop

file.Close

' Output results
WScript.Echo "Sum of Column " & (columnIndex1 + 1) & ": " & sum1
WScript.Echo "Sum of Column " & (columnIndex2 + 1) & ": " & sum2
WScript.Echo "Grand Total: " & (sum1 + sum2)
