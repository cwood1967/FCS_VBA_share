Attribute VB_Name = "Module5"
Function xFCS()
    Dim FcsController As AimFcsController
    Dim HardwareControl As AimFcsHardwareControl
    Dim fcspos As AimFcsSamplePositionParameters
    Dim Progress As AimProgress
    Dim DocumentList As ConfoCorDisplay.ConfoCorDisplay
    Dim Document As ConfoCorDisplay.FcsDocument
    Dim DataList As AimFcsData
    Dim idata As IAimFcsDataList
    Dim DataSet As AimFcsDataSet
    Dim DataSet2 As AimFcsDataSet
    Dim FileWriter As AimFcsFileWrite
    Dim CorrelationArraySize As Long
    Dim CorrelationTime() As Double
    Dim Correlation() As Double
    Dim fcsOptions As AimFcsControllerOptions
    
    xf = 100
    yf = 100
    zf = 100
    inum = 3
    
    fname = "c:\chris\fcstest"
    Set FcsController = Lsm5.ExternalDsObject.FcsController
    Set DocumentList = Lsm5.ExternalDsObject.ConfoCorDisplay
    Set fcspos = FcsController.SamplePositionParameters
    
    Set fcsOptions = FcsController.Options
    
    
    fcsOptions.RawDataPath = fname & "\" & "raw_file" & CStr(inum) & "__"

    FcsController.AcquisitionParameters.MeasurementTime = 2
    FcsController.AcquisitionParameters.MeasurementRepeat = 2
    
    'MsgBox fcspos.PositionZ(0)
    ' Here comes an Example on how to access other interfaces of the fcs controller
    ' either assign to a varibale
    
    
    Set HardwareControl = FcsController.HardwareControl
    
     FcsController.Options.StageIndex = eFcsStageScannerXY
     
    'FcsController.Initialize HardwareControl
    'MsgBox FcsController.HardwareControl.MicroscopeMode
    'HardwareControl.MicroscopeMode = eFcsMicroscopeModeFCS
    
    Set Document = DocumentList.NewDocument(True)
    Set DataList = Document.DataList
    
    'DocumentList.MicroscopeToFcsMode
    
    Dim pz As Double
    
    'FcsController.HardwareControl.GetFocusPosition (pz)
    'MsgBox   z
    'HardwareControl.MicroscopeMode = 3
    'HardwareControl.MicroscopeMode = 4
    ' or do the same by usage of the correponding method
    
   ' FcsController.HardwareControl.MicroscopeMode = 4 'eFcsMicroscopeModeFCS
    
    ' Create new document
    
    
    
    ' Acqiusition
    'fcspos.SamplePositionMode = eFcsSamplePositionModeCurrent
    
    
    fcspos.SamplePositionMode = eFcsSamplePositionModeSequential
    fcspos.PositionListSize = 1
    Dim s1 As Double
    Dim s2 As Double
    Dim s3 As Double
    
    Dim stx
    Dim sty
    
    stx = Lsm5.Hardware.CpStages.PositionX
    sty = Lsm5.Hardware.CpStages.PositionY
    
    Debug.Print "Stage Positions " & " " & stx & " " & sty
    
    HardwareControl.GetStagePosition s1, s2
    HardwareControl.GetFocusPosition s3
    
    p1 = fcspos.PositionX(0)
    p2 = fcspos.PositionY(0)
    p3 = fcspos.PositionZ(0)
    
    Debug.Print "****" & s1 & " " & stx
    Debug.Print "****" & s2 & " " & sty
    Debug.Print "Current Z position 1****" & s3
    
    'Debug.Print fcspos.SamplePositionMode
    'fcspos.SamplePositionMode = eFcsDataSamplePositionModeList
    fcspos.SamplePositionMode = eFcsSamplePositionModeSequential
    fcspos.PositionListSize = 1
    
    fcspos.PositionX(0) = xf * 0.000001
    fcspos.PositionY(0) = yf * 0.000001

    Debug.Print "x position 0**** " & fcspos.PositionX(0)
    Debug.Print "y position 0****" & fcspos.PositionY(0)
    Debug.Print "Current Z position 2 0****" & fcspos.PositionZ(0)
    
    fcspos.PositionZ(0) = s3
    
    HardwareControl.MoveToMeasurementPosition (0)

    Lsm5.Hardware.CpFocus.position = zf
    
    While HardwareControl.IsFocusBusy
        HardwareControl.GetFocusPosition s3
        Debug.Print "F Busy " & fcspos.PositionZ(0) & " " & s3
        Debug.Print "LSM Focus Before " & Lsm5.Hardware.CpFocus.position
        Sleep 20
    Wend
        
    Debug.Print "Current Z position 3 0****" & fcspos.PositionZ(0)
    Debug.Print "LSM Z Position ***  " & Lsm5.Hardware.CpFocus.position
    fcsOptions.SaveRawData = True
    FcsController.StartMeasurement DataList
    
    Set Progress = FcsController.Progress
    
    While (Not Progress.Ready)
        DoEvents
    Wend
    
    ' Do something with the data of the first data set
    
    Debug.Print "x pos 1**** " & fcspos.PositionX(0)
    Debug.Print "y pos 1**** " & fcspos.PositionY(0)
    Debug.Print "Z Position After 1**** " & fcspos.PositionZ(0)
    'MsgBox CStr(DataList.DataSets) + " data sets acquired"
    
    
    Set DataSet = DataList.DataSet(0)
    'Set DataSet2 = DataList.DataSet(1)
    
    'DataSet.SetRawDataFileName False, "fcs1_" & CStr(inum) & ".raw"
    'DataSet2.SetRawDataFileName False, "fcs2_" & CStr(inum) & ".raw"
    CorrelationArraySize = DataSet.DataArraySize(eFcsDataTypeCorrelation)
    
    ReDim CorrelationTime(CorrelationArraySize) As Double
    ReDim Correlation(CorrelationArraySize) As Double
    
    DataSet.GetDataArray eFcsDataTypeCorrelation, CorrelationArraySize, CorrelationTime(0), Correlation(0)
    
    'MsgBox "Correlation at " + Format(CorrelationTime(0), "0.00000000 seconds") + " is " + Format(Correlation(0), "0.00000000")
    
    ' Write to file
    
    Set FileWriter = New AimFcsFileWrite
    FileWriter.Source = DataList
    FileWriter.FileName = fname & "\" & "fcs_" & CStr(inum) & ".fcs"
    FileWriter.FileWriteType = eFcsFileWriteTypeAll
    FileWriter.Format = eFcsFileFormatConfoCor3
    FileWriter.Format = eFcsFileFormatConfoCor3WithRawData
    FileWriter.Start ' use start for asynchron writing or "Run" for sysnchron writing
                     ' there is no need to check the progress in the latter case
    
    Set Progress = FileWriter
    While (Not Progress.Ready)
        DoEvents
    Wend
    
'    FileWriter.FileName = fname & "\" & "fcs_" & CStr(inum) & ".raw"
'    FileWriter.FileWriteType = eFcsFileWriteTypeAll
'    FileWriter.Format = eFcsFileFormatRawConfoCor3
'     FileWriter.Start
'    'FileWriter.Format = eFcsFileFormatConfoCor3WithRawData
    
    Set Progress = FileWriter
    While (Not Progress.Ready)
        DoEvents
    Wend
    
    'remove the document again
    Document.CloseWindow
    
    Lsm5.Hardware.CpStages.PositionX = stx
    Lsm5.Hardware.CpStages.PositionY = sty
    
    FcsController.HardwareControl.MicroscopeMode = 1 'eFcsMicroscopeModeLSM
    Sleep 100
    
End Function

Sub trtest()
Open "c:\temp\wellpositions.dat" For Append As #4
            Print #4, "Position " & i & " " & j; ""
            Print #4, "output : "; Lsm5.Hardware.CpStages.PositionY & " " & Lsm5.Hardware.CpStages.PositionY
            Close #4
End Sub
