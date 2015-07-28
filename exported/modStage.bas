Attribute VB_Name = "modStage"

Function moveStage(dX As Double, dY As Double)


    '' the LSM software returns passes everything in micrometers
    Dim stages As CpStages
    
    Set stages = Lsm5.Hardware.CpStages
    
    Dim cx As Double
    Dim cy As Double
    
    cx = stages.PositionX
    cy = stages.PositionY
    
    Debug.Print "current: " & cx & " , " & cy
    
    Dim newx As Double
    Dim newy As Double
    
    newx = cx + dX
    newy = cy + dY
    
    Debug.Print "new: " & newx & " , " & newy
    
    stages.PositionX = newx
    stages.PositionY = newy
    
    Debug.Print "new set: " & stages.PositionX & " , " & stages.PositionY
End Function


Function teststage()
    moveStage Lsm5.DsRecording.FrameWidth, 0
    
    MsgBox Lsm5.DsRecording.FrameWidth & " " & Lsm5.DsRecording.FrameHeight
    
End Function
