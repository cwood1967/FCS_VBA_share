Attribute VB_Name = "modPump"

Public Sub test()
    pumpoffsetx = (30683.75 - 20838.75) - 500
    pumpoffsety = (-16979 - (-42138)) + 0
    
    stageWellOriginX = Lsm5.Hardware.CpStages.PositionX
    stageWellOriginY = Lsm5.Hardware.CpStages.PositionY
    
    AddWater
End Sub

Sub ptest()

    MsgBox Lsm5.Hardware.CpStages.PositionX
    MsgBox Lsm5.Hardware.CpStages.PositionY
End Sub

Public Sub AddWater()
    Dim currentz As Double
    Dim currentx As Double
    Dim currenty As Double
    
    
    currentz = Lsm5.Hardware.CpFocus.position
    currentx = Lsm5.Hardware.CpStages.PositionX
    currenty = Lsm5.Hardware.CpStages.PositionY
    
    Lsm5.Hardware.CpFocus.MoveToLoadPosition
    
    While Lsm5.Hardware.CpFocus.IsBusy
        Sleep 250
    Wend
    
    
    px = waterx
    py = watery
    
    'px = stageWellOriginX + pumpoffsetx
    'py = stageWellOriginY - pumpoffsety
    
    
    
    Lsm5.Hardware.CpStages.PositionX = px
    Lsm5.Hardware.CpStages.PositionY = py
    
    While Lsm5.Hardware.CpStages.IsBusy
        Sleep 250
    Wend
    
    
    sendtrigger
    Sleep 1000
    
    
    Lsm5.Hardware.CpStages.PositionY = currenty
    
    While Lsm5.Hardware.CpStages.IsBusy
        Sleep 500
    Wend
    
    Lsm5.Hardware.CpStages.PositionX = currentx
    
    While Lsm5.Hardware.CpStages.IsBusy
        Sleep 500
    Wend
    
    Lsm5.Hardware.CpFocus.position = currentz
    
    While Lsm5.Hardware.CpFocus.IsBusy
        Sleep 250
    Wend
    
End Sub

Sub sendtrigger()

    Lsm5.Hardware.CpScancontrol.SendTriggerOut 0
    DoEvents
    Sleep 1500
    Lsm5.Hardware.CpScancontrol.SendTriggerOut 0
    Sleep 1000
    
End Sub



