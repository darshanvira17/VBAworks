Sub TextFile()
'
'TextFile Generator Macro
'
Dim f As Integer '-----------------For .txt file------------------'
'--Dim f2 As Integer '-----------------For ftpCommands.txt script------------------'
'---Dim f3 As Integer '-----------------For .bat to run the ftp.Commands.txt script------------------'

Dim OutPut As Integer '-------for MsgBoX
Dim Path As String '---------main folder which will have all the files being created by this macro-------'
Dim TextFile As String '------comnplete path to .txt------'
'---Dim ftpCommFile As String '-----comnplete path to .txt file containing ftp commands------'
'---Dim batchFile As String '-----comnplete path to .bat file to execute ftpCommFile------'
Dim cellValue As Variant, a As Integer, b As Integer, c As Integer, d As Integer
For c = 1 To Range("A1:B8").Rows.Count
        For d = 1 To Range("A1:B8").Columns.Count
            If (IsEmpty(Range("A1:B8").Cells(c, d).Value)) Then
                MsgBox "Error: Missing Data!"
                Exit Sub
            End If
        Next d
Next c
If (Range("B9").Value = "N" Or Range("B9").Value = "n") Then
    MsgBox "Value in cell B9 is 'N'/'n' so No data written to the text file!"
    Exit Sub
ElseIf (Range("B9").Value = "Y" Or Range("B9").Value = "y") Then
    On Error GoTo Err_Handler
        f = FreeFile
        Path = "Path\to\your\directory\end\with\a\slash\or\add\one\before\Data\in\next\line\"
        TextFile = Path & "Data" & Format(DateTime.Now, "yyyymmddhhmm") & ".txt"
        'If you want to FTP the text file to some server -> ftpCommFile = Path & "ftpCommands.txt"
        'Batch File to run the FTP file which will execute FTP commands file -> batchFile = Path & "runftp.bat"
        '-------------Writing Data to .txt file--------------'
    
        Open TextFile For Output As #f
        For a = 1 To Range("A1:B8").Rows.Count
            For b = 1 To Range("A1:B8").Columns.Count
                cellValue = Range("A1:B8").Cells(a, b).Value
                If b = Range("A1:B8").Columns.Count Then
                    Print #f, cellValue
                Else
                    Print #f, cellValue,
                End If
            Next b
        Next a
        Close #f
        '------------------After the Text file is generated, clear the values inserted and ready for next use!----------------'
        Set rng1 = Worksheets("Sheet1").Range("A1:B8")
        rng1.ClearContents
    
        '---------------Writing FTP commands to the ftp commands file-----------'
        '---Open ftpCommFile For Output As #f2
        '----Print #f2, "open SERVERNAME"
        '----Print #f2, "USERNAME"
        '----Print #f2, "PASSWORD"
        '----Print #f2, "cd /Path/to/directory/on/Server"
        '----Print #f2, "ascii"
        '----Print #f2, "send " & TextFile
        '----Print #f2, "bye"
        '----Close #f2

        '--------Creating the batch file to execute ftp and send .txt to the server-----------'
        '-----Open batchFile For Output As #f3
        '-----Print #f3, "ftp -s:" & ftpCommFile
        '-----Print #f3, "Echo ""Complete!"" > " & Path & "ftpDone.out"
        '-----Close #f3
        '----invoke the shell to run the batch file to run the ftp command and send .txt over to the server----'
        '-----Shell (batchFile), vbHide
        '---------------------------Wait for completion----------------------'
        'Do While Dir(Path & "ftpDone.out") = ""
        '   DoEvents
        'Loop
        '------------------------- Clean up files------------------------------'
        '----If Dir(batchFile) <> "" Then Kill (batchFile)
        '----If Dir(Path & "ftpDone.out") <> "" Then Kill (Path & "ftpDone.out")
        '----If Dir(ftpCommFile) <> "" Then Kill (ftpCommFile)
        '----If Dir(TextFile) <> "" Then Kill (TextFile)
    
        '---display the success report-----'
        OutPut = MsgBox("Success! Text File is Generated!", vbOKOnly)
        Exit Sub
Err_Handler:
        OutPut = MsgBox("Error : " & Err.Number & vbCrLf & "Description : " & Err.Description, vbCritical)
        Exit Sub
Else
OutPut = MsgBox("Error : Please Enter Y/y or N/n in B9 cell.")
End If
End Sub
