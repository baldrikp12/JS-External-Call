' Author: Kenneth Baldridge
' E-Mail: Kenneth.p.Baldridge@boeing.com
' 4/21/2014
'
'   This program is designed to:
' 	1) Export specific cells to a Data.txt file following the format:
'		First line tells the program the number of dimensions, i.e. X, Y and Z
'       Second line tells the program how many points to evaluate
'       The following lines defines X, Y, Z coordinates of data point(s), 
'		separated by a space (one point per line).
'	2) Pass the Data.txt file through an external program, cbnd.exe, that will calculate 
'		the smallest enclosing sphere then spit out the center point and radius of said 
'		sphere to a Results.txt file.
'	3) Read the Results.txt file and process the lines containing the X, Y, and Z coordinates 
'		and Radius of the smallest enclosing spehere and putting them in the excel file.
'	4) Clean up by deleting the Data.txt and Results.txt files.

Sub SmallestEcnlosingSphere()
   
   'Declares the location of the File holding
   'our exported cells and the results from cbnd.exe as a variable.
    Dim workspace As String
    workspace = "C:\Test\" 'Update this and make sure the cbnd.exe program is in here too.
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Part 1) Export Data From Excel to text file Data.txt.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'The Excel created file to pass into cbnd.exe.
    Dim In_file As String
    In_file = workspace & "Data.txt"
   
   'Opens the file for writing to.
    Open In_file For Append As #1
   
   'i will traverse rows.
    Dim i As Integer
    
    'j will traverse columns.
    Dim j As Integer
   
   '3 dimensional points.
   '9 points to be used.
    Print #1, "3"
    Print #1, "9"
   
    'B:J
    For i = 2 To 10
        
        '7:9
        For j = 7 To 9
        
            'Grabs data from cell and prints it to text file.
            'Semicolon indicated printing on same line.
            Print #1, Cells(j, i).Value & " ";
         
        Next j
        
        'No semicolon at end means move cursor to new line after printing.
        Print #1, ""
      
    Next i
    
    'Closes the connection to the opened file
    Close #1
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Part 2) Pass Data.txt through cbnd.exe program and create Results.txt.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'The cbnd.exe output file.
    Dim Out_file As String
    Out_file = workspace & "Results.txt"
    
    'Runs the command "cbnd <FileToReadData.txt >FiletoPutResults.txt".
    Call Shell("cmd.exe /S /C" & workspace & "cbnd.exe <" & In_file & " >" & Out_file)
    
    'Waits a few seconds allowing for the Results.txt file to get created first
    'before trying to find and open it for reading.
    Application.Wait Now + TimeSerial(0, 0, 5)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Part 3) Read Results.txt file and process appropriate lines.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Opens the newly created file with the results.
    Open Out_file For Input As #2
   
    'Used to keep track of what line in the file we are currently on.
    Dim k As Integer
    k = 1
    
    'Used to hold the contents of the current line.
    Dim textline As String
   
   'Keep reading until end of file.
    Do Until EOF(2)
   
        'Assigns the current line to textline
        Line Input #2, textline
      
        'Check what k currently is then go to appropriate case
        'and perform action: Assign data to cell.
        Select Case k
            Case 2 'Line 2 (X-coordinate)
                Cells(67, 1).Value = textline 'Update me
            Case 3 'Line 3 (Y-coordinate)
                Cells(68, 1).Value = textline 'Update me
            Case 4 'Line 4 (Z-coordinate)
                Cells(69, 1).Value = textline 'Update me
            Case 6 'Line 6 (Radius)
                Cells(70, 1).Value = textline 'Update me
        End Select
        
        k = k + 1
      
    Loop
   
    'Closes connection to second opened file.
    Close #2
   
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Part 4) Delete created text files.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Deletes the file with data.
    Killfile In_file
    
    'Deletes the file with results.
    Killfile Out_file
    
End Sub

Function Killfile$(file As String)

    'Checks if directory exists then kills passed file.
    If Len(Dir$(file)) > 0 Then
        Kill file
        Killfiles = "File Deleted"
    End If
        
End Function
   


