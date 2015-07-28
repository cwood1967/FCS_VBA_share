Attribute VB_Name = "Module4"
Public Sub testscan1()
    'Lsm5.NewScanWindow
    'Lsm5.StartScan
Dim x As Double
Dim y As Double

Lsm5.Hardware.CpStages.MarkMoveTo 0
'r = Lsm5.Hardware.CpStages.MarkGet(1, x, y)

'    While Lsm5.Hardware.CpScancontrol.IsGrabbing
'        Sleep 100
'    Wend
            
End Sub

Sub mytest()

    Dim rec As DsRecording
    Dim i As Integer
    
    Set rec = Lsm5.DsRecording
    rec.Sample0X = 25
    rec.Sample0Y = -30
    
    rec.zoomx = 5
    rec.zoomy = 5
    rec.SamplesPerLine = 500
    rec.LinesPerFrame = 500
    
    Dim dc As DsDataChannel
    
    Set dc = rec.TrackObjectByIndex(0, i).DataChannelObjectByIndex(0, i)
    
    dc.BitsPerSample = 12
    
    Lsm5.StartScan
    While Lsm5.Hardware.CpScancontrol.IsGrabbing
    
        Sleep 50
        
    Wend

End Sub

Public Sub aotftest()
    
    Dim d As DsRecordingDoc
    Dim power As Double
    Dim Options As AimScanControllerOptions
    Dim imagePower As Double
    Lsm5.NewScanWindow
    
    
    
    
    Dim bp As AimFcsBeamPathParameters
    'Set bp = FcsController.BeamPathParameters
    Dim ap
    
    'ap = bp.AttenuatorPower(0)
    'MsgBox ap
    'bp.AttenuatorPower(0) = 0.005
    power = 0.38
        
    Set Options = Lsm5.ExternalDsObject.ScanController
    imagePower = Options.AotfDriverPower
    
    Options.AotfDriverPower = power
    
    Set d = Lsm5.StartScan
    
    While Lsm5.Hardware.CpScancontrol.IsGrabbing
        Sleep 100
    Wend
    'Set Options = Nothing
    
    power = 0.13
    Options.AotfDriverPower = power
    
    Dim d2 As DsRecordingDoc
    
    Lsm5.NewScanWindow
    Set d2 = Lsm5.StartScan
    
    
    While Lsm5.Hardware.CpScancontrol.IsGrabbing
        Sleep 100
    Wend
    
End Sub

