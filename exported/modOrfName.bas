Attribute VB_Name = "modOrfName"

Sub test()
    OrfFilename = "U:\cjw\OME\Jan_09_2008-Plate3\000_AutoScan.mdb\OrfName.txt"
    readOrfNames
End Sub

Public Function readOrfNames()

    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim Stream As TextStream
    
    Set Stream = fso.OpenTextFile(OrfFilename, ForReading)
    
    Dim n  As Integer
    Dim orfnames() As String
    Dim tmpnames
    ReDim tmpnames(100)
    
    Do While Not Stream.AtEndOfStream
        tmpstr = Stream.ReadLine
        If (Strings.Trim(tmpstr) <> "") Then
            
            tmpnames(n) = tmpstr
            n = n + 1
            
        End If
        
    Loop
    
    Stream.Close
    Set fso = Nothing
    
    ReDim orfnames(0 To n - 1)
    
    For i = 0 To n - 1
        orfnames(i) = tmpnames(i)
        Debug.Print orfnames(i)
    Next
    
    readOrfNames = orfnames
End Function

