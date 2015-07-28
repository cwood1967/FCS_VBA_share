Attribute VB_Name = "modReadFile"

Public Sub readDataFile()
    
    Dim datafile As String
    datafile = "c:\temp\xdata.dat"
    
    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim Stream As TextStream
    Set Stream = fso.OpenTextFile(datafile, ForReading)
    
    Dim inum As Integer
    inum = CLng(Stream.ReadLine)
    
    Dim data_array()
    ReDim data_array(inum)
    Dim iline As Integer
    iline = 0
    
    While Not Stream.AtEndOfStream
        data_array(iline) = Stream.ReadLine
        iline = iline + 1
            
    Wend
    
    Stream.Close
    
    For iline = 0 To inum - 1
        
    Next
End Sub

