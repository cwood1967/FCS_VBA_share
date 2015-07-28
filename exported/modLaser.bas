Attribute VB_Name = "modLaser"
Sub lasersOff()
    
    Dim lasers As CpLasers
    
    Dim i As Integer
    
    Dim count As Integer
    
    Set lasers = Lsm5.Hardware.CpLasers
    
    count = lasers.count
    
    
    For i = 0 To count - 1
        
        If (lasers.Select(i)) Then
            lasers.State = eLaserOff
        End If
        
    Next
    
End Sub
