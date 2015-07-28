Attribute VB_Name = "modFocusFCS"
Sub ltest()

    zpoints = 10
    ztime = 0.1
    Dim z As Double
    fname = "e:\chris\ltest"
    For i = 0 To 300
    
        z = setupFCS(0, 0)
        takeFCS 0, 0, z, i
        Debug.Print "Number " & i & "######################################################"
    Next
End Sub

Function setupFCS(xf, yf) As Double
    
    Dim FcsController As AimFcsController
    Dim HardwareControl As AimFcsHardwareControl
    Dim fcspos As AimFcsSamplePositionParameters
    Dim Progress As AimProgress
    Dim DocumentList As ConfoCorDisplay.ConfoCorDisplay
    Dim Document As ConfoCorDisplay.FcsDocument
    Dim DataList As AimFcsData
    Dim DataSet As AimFcsDataSet
    Dim bp As AimFcsBeamPathParameters
    
    'Dim FileWriter As AimFcsFileWrite
    'Dim CorrelationArraySize As Long
    'Dim CorrelationTime() As Double
    'Dim Correlation() As Double
    
    Set FcsController = Lsm5.ExternalDsObject.FcsController
    Set DocumentList = Lsm5.ExternalDsObject.ConfoCorDisplay
    Set fcspos = FcsController.SamplePositionParameters
    Set Progress = FcsController.Progress
    'Set bp = FcsController.BeamPathParameters
    'Dim ap
    
    'ap = bp.AttenuatorPower(2)
    
    Set HardwareControl = FcsController.HardwareControl
     
     time1 = Now
    FcsController.Options.StageIndex = eFcsStageScannerXY
    
    
     'MsgBox FcsController.HardwareInformation.StageIndexFromIdentifier(eFcsStageScannerXY)
    'HardwareControl.MicroscopeMode = eFcsMicroscopeModeFCS
    
    'DocumentList.MicroscopeToFcsMode
    
    Dim imagePower As Double
    
    'Dim power As Double
    'Dim Options As AimScanControllerOptions
    
    'power = 0.38
    
    'Set Options = Lsm5.ExternalDsObject.ScanController
    'imagePower = Options.AotfDriverPower
    'Options.AotfDriverPower = Power
    'ap = bp.AttenuatorPower(0)
    Dim pz As Double
    
    
    Dim s1 As Double
    Dim s2 As Double
    Dim s3 As Double
    
    stx = Lsm5.Hardware.CpStages.PositionX
    sty = Lsm5.Hardware.CpStages.PositionY
    
    HardwareControl.GetStagePosition s1, s2
    HardwareControl.GetFocusPosition s3
    
    FcsController.AcquisitionParameters.MeasurementTime = ztime
    FcsController.AcquisitionParameters.MeasurementRepeat = 1
    
    p1 = fcspos.PositionX(0)
    p2 = fcspos.PositionY(0)
    p3 = fcspos.PositionZ(0)
    
    'fcspos.SamplePositionMode = eFcsSamplePositionModeCurrent
    'Debug.Print fcspos.Dump
    'Lsm5.DsRecording.Sample0X = xf
    'Lsm5.DsRecording.Sample0Y = yf
    
    FcsController.Options.StageIndex = eFcsStageScannerXY
    fcspos.SamplePositionMode = eFcsSamplePositionModeSequential
    
    fcspos.PositionX(0) = xf * 0.000001
    fcspos.PositionY(0) = yf * 0.000001
    'Debug.Print fcspos.Dump
    'fcspos.SamplePositionMode = eFcsSamplePositionModeList
    'fcspos.SamplePositionMode = eFcsSamplePositionModeCurrent
    
    Debug.Print "stage **** " & s1 & " "; s2 & " " & s3
    
    
    Debug.Print "position **** " & p1 & " " & p2 & " " & p3
    
     
    'Debug.Print fcspos.SamplePositionMode
    'fcspos.SamplePositionMode = eFcsDataSamplePositionModeList
    'fcspos.SamplePositionMode = eFcsSamplePositionModeCurrent
    
    ''''''Lsm5.DsRecording.Sample0X
    'fcspos.PositionListSize = 1

    
    While HardwareControl.IsStageBusy
        Sleep 100
    Wend
      
    'FcsController.StopAcquisition
    Dim na As Long
    
    'na = FcsController.GetCountRateImageResultArraySize
    'MsgBox na
    Dim gdata As Double
    Dim xd(0 To 9) As Double
    
    fcspos.PositionZ(0) = s3
    'Debug.Print fcspos.Dump
    Dim fstart As Double
    Dim nmeas As Double
    Dim fthick As Double
    
    fthick = zthickness * 0.000001
    'fcspos.SamplePositionMode = eFcsSamplePositionModeCurrent
    'fthick = 0.000008
    'Debug.Print fcspos.Dump
    'nmeas = 10
    
    nmeas = zpoints
    fdel = fthick / nmeas
    
    Dim forig As Double
    
    forig = s3
    fstart = s3 - 0.5 * nmeas * fdel
    
    If fstart * 1000000# < (slipPosition + 2) Then
        fstart = (0.000001) * (slipPosition + 2)
    End If
    
    Debug.Print "Start: " & " " & fstart
    Dim fposz As Double
    
    'Dim c1cr(0 To 9)
    Set Document = DocumentList.NewDocument(False)
    Debug.Print Second(Now) - Second(time1) & "  *************************"
    Set DataList = Document.DataList
    
    fmx = Lsm5.Hardware.CpStages.PositionX
    fmy = Lsm5.Hardware.CpStages.PositionY
    Debug.Print "Stage right before measure (1)" & fmx & " " & fmy
    
    FcsController.HardwareControl.MoveToMeasurementPosition (0)
    
    fmx = Lsm5.Hardware.CpStages.PositionX
    fmy = Lsm5.Hardware.CpStages.PositionY
    Debug.Print "Stage right before measure (2)" & fmx & " " & fmy
    
    Dim num_meas As Long
    num_meas = 0
    For ifcs = 0 To nmeas - 1
'        DoEvents
        fposz = fstart + fdel * ifcs
        
        If (fposz * 1000000#) < slipPosition + 2 Then
            Exit For
        End If
        
        fmz = Lsm5.Hardware.CpFocus.position
        HardwareControl.GetFocusPosition s3
        'Debug.Print "Stop 1"
        'Debug.Print "Focus Position before " & fmz & " " & s3
        
        fcspos.PositionZ(0) = fposz
        
        FcsController.HardwareControl.MoveToMeasurementPosition (0)
        'FcsController.HardwareControl.MoveFocusToPosition (0)
        Lsm5.Hardware.CpFocus.position = fposz * 1000000#
        
        ''take this out for now, to see what happpenss
        While HardwareControl.IsFocusBusy
            'HardwareControl.GetFocusPosition s3
            'Debug.Print "F Busy " & fcspos.PositionZ(0) & " " & s3
            'Debug.Print "LSM Focus Before " & Lsm5.Hardware.CpFocus.position
            Sleep 20
            DoEvents
        Wend
    
        'Debug.Print "Postion: " & ifcs & " " & fposz
        'Debug.Print "LSM Focus Before " & Lsm5.Hardware.CpFocus.position
        
        'fmx = Lsm5.Hardware.CpStages.PositionX
        'fmy = Lsm5.Hardware.CpStages.PositionY
        'Debug.Print "Stage right before measure " & fmx & " " & fmy
        
        'Debug.Print "scan position before " & fcspos.PositionX(0) & " " & fcspos.PositionY(0)
        'Debug.Print "scan position before " & Lsm5.DsRecording.Sample0X & " " & Lsm5.DsRecording.Sample0Y
               
        FcsController.StartMeasurement DataList
        
        
        'Debug.Print MyTime

    
        Debug.Print "Scan Started  " & Lsm5.Hardware.CpFocus.position
        
        
        While FcsController.IsAcquisitionRunning
            Sleep 100
            'DoEvents
            'Debug.Print "In While loop " & s3 & " " & fcspos.PositionZ(0)
            'Debug.Print I & "  " & Lsm5.Hardware.CpFocus.Position
            i = i + 1
        Wend
    
        While (Not Progress.Ready)
            Sleep 200
            'DoEvents
            'Debug.Print "Progress Busy"
        Wend
         DoEvents
        'Debug.Print "Scan complete"
        'fmx = Lsm5.Hardware.CpStages.PositionX
        'fmy = Lsm5.Hardware.CpStages.PositionY
        'Debug.Print "Stage right after measure " & fmx & " " & fmy
        'fmz = Lsm5.Hardware.CpFocus.position
     
        'Debug.Print "scan position after" & fcspos.PositionX(0) & " " & fcspos.PositionY(0)
        'Debug.Print "scan position after " & Lsm5.DsRecording.Sample0X & " " & Lsm5.DsRecording.Sample0Y
        'HardwareControl.GetStagePosition s1, s2
        'HardwareControl.GetFocusPosition s3
        'Debug.Print "Focus Position After" & fmz & " " & s3
        'Debug.Print "LSM Focus After " & Lsm5.Hardware.CpFocus.position
        
        'Debug.Print "Stages: " & s1 & " "; s2 & " " & s3
        'Debug.Print "Scan Positions " & fcspos.PositionX(0) & "  " & fcspos.PositionY(0)
        'Debug.Print Now & " *** " & Time
        num_meas = num_meas + 1
        FcsController.StopAcquisitionAndWait
        Sleep 100
        DoEvents
        Debug.Print "*************end " & ifcs & " *******************"
    Next
    
    'fcspos.PositionZ(0) = forig
    'Lsm5.Hardware.CpFocus.Position = forig * 1000000#
    Sleep 100
    
    Debug.Print fcspos.PositionZ(0)
    
    'Dim I As Integer
    
    Dim ar1 As Double
    Dim ar2 As Double
    
    Dim fmax As Long
    Dim imax As Integer
    fmax = 0
    imax = 0
    Dim av As Long
    Dim acr As Double
    Dim nsets As Long
    
    nsets = DataList.DataSets
    nstep = nsets / num_meas
    
    For i = 0 To num_meas - 1
        Set DataSet = Document.DataList.DataSet(nstep * i)
        DataSet.GetAverageCountRate av, acr
        DataSet.GetDataArray eFcsDataTypeCountRate, 1, ar1, ar2
        If acr > fmax Then
            fmax = acr
            imax = i
        End If
        Debug.Print i & " " & 1000000# * (fstart + fdel * i) & " " & av & " " & acr
        Set DataSet = Nothing
        DoEvents
    Next
    Debug.Print fmax & " " & imax
    
    fcspos.PositionZ(0) = fstart + fdel * imax
    Lsm5.Hardware.CpFocus.position = (fstart + fdel * imax) * 1000000#
    
    Debug.Print "After Stage " & Lsm5.Hardware.CpStages.PositionX & " " & Lsm5.Hardware.CpStages.PositionY
    
    Lsm5.Hardware.CpStages.PositionX = stx
    Lsm5.Hardware.CpStages.PositionY = sty
    
    While Lsm5.Hardware.CpStages.IsBusy
        Sleep 100
        DoEvents
    Wend
    
    'Options.AotfDriverPower = 1
    Sleep 200
    'Set Options = Nothing
    
    
    
    Debug.Print "After Stage " & Lsm5.Hardware.CpStages.PositionX & " " & Lsm5.Hardware.CpStages.PositionY
    
    Debug.Print "**** max position *** == "; fstart + fdel * imax
    Document.CloseWindow
    Set FcsController = Nothing
    Set Progress = Nothing
    Set fcspos = Nothing
    Set DocumentList = Nothing
    Set Document = Nothing
    Set HardwareControl = Nothing
    
    time2 = Now
    'Debug.Print Second(time2 - Time1)
    setupFCS = CDbl(1000000# * (fstart + fdel * imax))
    DoEvents
End Function


Function checkFCS(xf, yf) As Double
    
    Dim FcsController As AimFcsController
    Dim HardwareControl As AimFcsHardwareControl
    Dim fcspos As AimFcsSamplePositionParameters
    Dim Progress As AimProgress
    Dim DocumentList As ConfoCorDisplay.ConfoCorDisplay
    Dim Document As ConfoCorDisplay.FcsDocument
    Dim DataList As AimFcsData
    Dim DataSet As AimFcsDataSet
    Dim bp As AimFcsBeamPathParameters
    
    'Dim FileWriter As AimFcsFileWrite
    'Dim CorrelationArraySize As Long
    'Dim CorrelationTime() As Double
    'Dim Correlation() As Double
    
    Set FcsController = Lsm5.ExternalDsObject.FcsController
    Set DocumentList = Lsm5.ExternalDsObject.ConfoCorDisplay
    Set fcspos = FcsController.SamplePositionParameters
    Set Progress = FcsController.Progress
    'Set bp = FcsController.BeamPathParameters
    'Dim ap
    
    'ap = bp.AttenuatorPower(2)
    
    Set HardwareControl = FcsController.HardwareControl
     
    
    'DocumentList.MicroscopeToFcsMode
    
    Dim imagePower As Double
    
   
    Dim pz As Double
  
    
    FcsController.AcquisitionParameters.MeasurementTime = 0.2
    FcsController.AcquisitionParameters.MeasurementRepeat = 1
    
    p1 = fcspos.PositionX(0)
    p2 = fcspos.PositionY(0)
    p3 = fcspos.PositionZ(0)
    
    
    FcsController.Options.StageIndex = eFcsStageScannerXY
    fcspos.SamplePositionMode = eFcsSamplePositionModeSequential
    
    fcspos.PositionX(0) = xf * 0.000001
    fcspos.PositionY(0) = yf * 0.000001
    
    While HardwareControl.IsStageBusy
        Sleep 100
    Wend
      
    'FcsController.StopAcquisition
    Dim na As Long
    
    'na = FcsController.GetCountRateImageResultArraySize
    'MsgBox na
    Dim gdata As Double
    Dim xd(0 To 9) As Double
    
    fcspos.PositionZ(0) = s3
    'Debug.Print fcspos.Dump
    Dim fstart As Double
    Dim nmeas As Double
    Dim fthick As Double
    
    fthick = zthickness * 0.000001
    'fcspos.SamplePositionMode = eFcsSamplePositionModeCurrent
    'fthick = 0.000008
    'Debug.Print fcspos.Dump
    'nmeas = 10
    
    nmeas = zpoints
    fdel = fthick / nmeas
    
    Dim forig As Double
    
    forig = s3
    fstart = s3 - 0.5 * nmeas * fdel
    
    If fstart * 1000000# < (slipPosition + 2) Then
        fstart = (0.000001) * (slipPosition + 2)
    End If
    
    Debug.Print "Start: " & " " & fstart
    Dim fposz As Double
    
    'Dim c1cr(0 To 9)
    Set Document = DocumentList.NewDocument(False)
    Debug.Print Second(Now) - Second(time1) & "  *************************"
    Set DataList = Document.DataList
    
    fmx = Lsm5.Hardware.CpStages.PositionX
    fmy = Lsm5.Hardware.CpStages.PositionY
    Debug.Print "Stage right before measure (1)" & fmx & " " & fmy
    
    FcsController.HardwareControl.MoveToMeasurementPosition (0)
    
    fmx = Lsm5.Hardware.CpStages.PositionX
    fmy = Lsm5.Hardware.CpStages.PositionY
    Debug.Print "Stage right before measure (2)" & fmx & " " & fmy
    
    Dim num_meas As Long
    num_meas = 0
    For ifcs = 0 To nmeas - 1
'        DoEvents
        fposz = fstart + fdel * ifcs
        
        If (fposz * 1000000#) < slipPosition + 2 Then
            Exit For
        End If
        
        fmz = Lsm5.Hardware.CpFocus.position
        HardwareControl.GetFocusPosition s3
        'Debug.Print "Stop 1"
        'Debug.Print "Focus Position before " & fmz & " " & s3
        
        fcspos.PositionZ(0) = fposz
        
        FcsController.HardwareControl.MoveToMeasurementPosition (0)
        'FcsController.HardwareControl.MoveFocusToPosition (0)
        Lsm5.Hardware.CpFocus.position = fposz * 1000000#
        
        ''take this out for now, to see what happpenss
        While HardwareControl.IsFocusBusy
            'HardwareControl.GetFocusPosition s3
            'Debug.Print "F Busy " & fcspos.PositionZ(0) & " " & s3
            'Debug.Print "LSM Focus Before " & Lsm5.Hardware.CpFocus.position
            Sleep 20
            DoEvents
        Wend
    
        'Debug.Print "Postion: " & ifcs & " " & fposz
        'Debug.Print "LSM Focus Before " & Lsm5.Hardware.CpFocus.position
        
        'fmx = Lsm5.Hardware.CpStages.PositionX
        'fmy = Lsm5.Hardware.CpStages.PositionY
        'Debug.Print "Stage right before measure " & fmx & " " & fmy
        
        'Debug.Print "scan position before " & fcspos.PositionX(0) & " " & fcspos.PositionY(0)
        'Debug.Print "scan position before " & Lsm5.DsRecording.Sample0X & " " & Lsm5.DsRecording.Sample0Y
               
        FcsController.StartMeasurement DataList
        
        
        'Debug.Print MyTime

    
        Debug.Print "Scan Started  " & Lsm5.Hardware.CpFocus.position
        
        
        While FcsController.IsAcquisitionRunning
            Sleep 100
            'DoEvents
            'Debug.Print "In While loop " & s3 & " " & fcspos.PositionZ(0)
            'Debug.Print I & "  " & Lsm5.Hardware.CpFocus.Position
            i = i + 1
        Wend
    
        While (Not Progress.Ready)
            Sleep 200
            'DoEvents
            'Debug.Print "Progress Busy"
        Wend
         DoEvents
        'Debug.Print "Scan complete"
        'fmx = Lsm5.Hardware.CpStages.PositionX
        'fmy = Lsm5.Hardware.CpStages.PositionY
        'Debug.Print "Stage right after measure " & fmx & " " & fmy
        'fmz = Lsm5.Hardware.CpFocus.position
     
        'Debug.Print "scan position after" & fcspos.PositionX(0) & " " & fcspos.PositionY(0)
        'Debug.Print "scan position after " & Lsm5.DsRecording.Sample0X & " " & Lsm5.DsRecording.Sample0Y
        'HardwareControl.GetStagePosition s1, s2
        'HardwareControl.GetFocusPosition s3
        'Debug.Print "Focus Position After" & fmz & " " & s3
        'Debug.Print "LSM Focus After " & Lsm5.Hardware.CpFocus.position
        
        'Debug.Print "Stages: " & s1 & " "; s2 & " " & s3
        'Debug.Print "Scan Positions " & fcspos.PositionX(0) & "  " & fcspos.PositionY(0)
        'Debug.Print Now & " *** " & Time
        num_meas = num_meas + 1
        FcsController.StopAcquisitionAndWait
        Sleep 100
        DoEvents
        Debug.Print "*************end " & ifcs & " *******************"
    Next
    
    'fcspos.PositionZ(0) = forig
    'Lsm5.Hardware.CpFocus.Position = forig * 1000000#
    Sleep 100
    
    Debug.Print fcspos.PositionZ(0)
    
    'Dim I As Integer
    
    Dim ar1 As Double
    Dim ar2 As Double
    
    Dim fmax As Long
    Dim imax As Integer
    fmax = 0
    imax = 0
    Dim av As Long
    Dim acr As Double
    Dim nsets As Long
    
    nsets = DataList.DataSets
    nstep = nsets / num_meas
    
    For i = 0 To num_meas - 1
        Set DataSet = Document.DataList.DataSet(nstep * i)
        DataSet.GetAverageCountRate av, acr
        DataSet.GetDataArray eFcsDataTypeCountRate, 1, ar1, ar2
        If acr > fmax Then
            fmax = acr
            imax = i
        End If
        Debug.Print i & " " & 1000000# * (fstart + fdel * i) & " " & av & " " & acr
        Set DataSet = Nothing
        DoEvents
    Next
    Debug.Print fmax & " " & imax
    
    fcspos.PositionZ(0) = fstart + fdel * imax
    Lsm5.Hardware.CpFocus.position = (fstart + fdel * imax) * 1000000#
    
    Debug.Print "After Stage " & Lsm5.Hardware.CpStages.PositionX & " " & Lsm5.Hardware.CpStages.PositionY
    
    Lsm5.Hardware.CpStages.PositionX = stx
    Lsm5.Hardware.CpStages.PositionY = sty
    
    While Lsm5.Hardware.CpStages.IsBusy
        Sleep 100
        DoEvents
    Wend
    
    'Options.AotfDriverPower = 1
    Sleep 200
    'Set Options = Nothing
    
    
    
    Debug.Print "After Stage " & Lsm5.Hardware.CpStages.PositionX & " " & Lsm5.Hardware.CpStages.PositionY
    
    Debug.Print "**** max position *** == "; fstart + fdel * imax
    Document.CloseWindow
    Set FcsController = Nothing
    Set Progress = Nothing
    Set fcspos = Nothing
    Set DocumentList = Nothing
    Set Document = Nothing
    Set HardwareControl = Nothing
    
    time2 = Now
    'Debug.Print Second(time2 - Time1)
    setupFCS = CDbl(1000000# * (fstart + fdel * imax))
    DoEvents
End Function


Function takeSpotZ()

    Dim d As DsRecordingDoc
    Lsm5.DsRecording.ScanMode = "point"
    Lsm5.DsRecording.Sample0X = -18.21
    Lsm5.DsRecording.Sample0Y = 35.75
    Lsm5.DsRecording.TimeSeries = False
    Lsm5.DsRecording.SamplesPerLine = 1
    
    'Lsm5.DsRecording.zoomx = 35
    'Lsm5.DsRecording.zoomy = 35

    
    
    
    Dim c1 As Long
    Dim t1 As Long
    Dim z1 As Long
    Dim y1 As Long
    Dim x1 As Long
    'd.GetHotSpot c1, t1, z1, y1, x1
    Dim del As Double
    del = 0.5
    swidth = 8#
    Lsm5.Hardware.CpFocus.position = Lsm5.Hardware.CpFocus.position - swidth * 0.5
    For j = 0 To 9
        Lsm5.Hardware.CpFocus.position = -23 + del * j
        
        While Lsm5.Hardware.CpFocus.IsBusy
            Sleep 20
        Wend
        
        Set d = Lsm5.StartScan
        'Sleep 500
        
        While Lsm5.Hardware.CpScancontrol.IsScanning
            'Debug.Print "Scanning"
            Sleep 300
        Wend
        
        Lsm5.StopScan
    
        x = d.ScanLine(0, 0, 0, 0, 1, 12)
    
        u = UBound(x)
        Dim total
        total = 0
        For i = 0 To u
            total = total + x(i)
        
        Next
    
        m = total / (u + 1)
        Debug.Print j & " " & m & " " & u & " " & Lsm5.Hardware.CpFocus.position
        Debug.Print "-----"
    Next
        
        
        
End Function

Public Function MyTime() As String
  MyTime = Format(Now, "dd-MMM-yyyy HH:nn:ss") & "." & VBA.Right(Format(Timer, "#0.000"), 3)
End Function

Public Sub fff()
    
    
    'Set icp = Lsm5.ExternalCpObject
    Dim t As AimRealtimeReceiver
    
    Set t = Lsm5.ExternalCpObject.pHardwareObjects.pRealTimeReceivers.pItem(CVar(0))
    t.ReleaseFcsBuffer
    
    icp.AimRealtimeReceiver.ReleaseFcsBuffer
    
End Sub
