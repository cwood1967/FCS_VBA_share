Attribute VB_Name = "Module3"
Function test_1()

    Dim f As Object
    Dim g As ConfoCorDisplay.ConfoCorDisplay
    Dim gd As ConfoCorDisplay.FcsDocument
    
    Dim cpo As Object
    Dim hext As Hardware
    
    
    Set f = Lsm5.ExternalCpObject.pHardwareObjects.pFcsController

    Set hext = cpo.pHardwareObjects
    
    'Set f = hext.pFcsController
    
     'Set g = New ConfoCorDisplay.ConfoCorDisplay
     'Set f = CreateObject("FCSController.AimFcsController")
     
     Dim hc As AimFcsHardwareControl
     Set hc = f.HardwareControl
    
    
     f.Initialize hc
     MsgBox f.IsAcquisitionRunning
     g.MicroscopeToFcsMode
     'g.SetFcsController Lsm5Vba.Fcs
     
     Set gd = g.NewDocument(1)
     
    
     
    
End Function

Function aotftest()

    Dim sc As AimScanController
    
    'Set sc = Lsm5.ExternalCpObject.pHardwareObjects.pScanController
    Set sc = Lsm5.ExternalDsObject.ScanController
    MsgBox sc.AutoFocusFastMode
    MsgBox sc.AotfDriverPower
    
    
End Function

Function cctest()
   
    Dim gd As ConfoCorDisplay.FcsDocument
    Dim g As ConfoCorDisplay.ConfoCorDisplay
  
    Set ff = Application
  'On Error GoTo Errhand
  
  Set g = New ConfoCorDisplay.ConfoCorDisplay
  g.CloseAllWindows 0
  
   Dim cpo As Object
   
   Dim f As CP.FcsController
   'Dim fo As Object
   'Dim ho As Object
   Set cpo = Lsm5.ExternalCpObject
   'Set af = cpo.pHardwareObjects.pRealtimeControllers
   'Set ho = cpo.Document
   'Set rt1 = af.Item(0)
   'Set fc1 = rt1.FcsJob
   Dim a(1) As Long
   a(0) = CLng(0)
   Dim al As Long
   al = CLng(1)
   'fc1.StartMeasure al, CLng(0), CLng(0)
   
   Set f = cpo.pHardwareObjects.pFcsController
   'MsgBox f.DetectorPinholeDiameter(0)
   'MsgBox f.BeamSplitterFilter(1)
   
   'Dim e As AimFcsController
   'Set e = CreateObject("FcsController.AimFcsController")
   ' Set e = New AimFcsController
  
  'Dim z As Double
  'Dim hi As AimFcsHardwareControl
  'Set hi = e.HardwareControl

   ' Set hi = e.HardwareInformation
    
   ' Set gd = g.NewDocument(1)
  'g.MicroscopeToFcsMode
  Dim u As Object
  'Set u = Nothing
  Dim c As AimFcsCalibration
  
  Set c = e.Calibration
  'MsgBox f.DetectorPinholeDiameter(0)
  'e.StartMeasurement c
  Dim ho As AimFcsControllerOptions
  
   Set ho = e.Options
  'MsgBox Lsm5Vba.Lsm5.Info.IsFCS
   
   Set fo = f
  g.SetFcsController f
  
  
  'f.Kinetics = False
  f.StartMeasure
  

  Dim px As Double
  Dim py As Double
  
  'Dim z As Double
  
  'Set r = cpo.AimRealtimeControllerFcsJob
   
   Exit Function
   
Errhand:
   MsgBox "-err- " & err.Description
End Function
Sub Macro1()
    '**************************************
    'Recorded: 3/9/2007
    '**************************************
    Dim RecordingDoc As DsRecordingDoc
    Dim Recording As DsRecording
    Dim Track As DsTrack
    Dim Laser As DsLaser
    Dim DetectionChannel As DsDetectionChannel
    Dim IlluminationChannel As DsIlluminationChannel
    Dim DataChannel As DsDataChannel
    Dim BeamSplitter As DsBeamSplitter
    Dim Timers As DsTimers
    Dim Markers As DsMarkers
    Dim Success As Integer
    Set Recording = Lsm5.DsRecording
     
    
     
    Set RecordingDoc = Nothing
    Set Recording = Nothing
    Set Track = Nothing
    Set Laser = Nothing
    Set DetectionChannel = Nothing
    Set IlluminationChannel = Nothing
    Set DataChannel = Nothing
    Set BeamSplitter = Nothing
    Set Timers = Nothing
    Set Markers = Nothing
    '************* End ********************
End Sub

Function stest()
    Dim a As DsRecordingDoc
    
    
    Dim c As CpServos
    
    MsgBox Lsm5.Hardware.CpFocus.position
    Lsm5.Hardware.CpFocus.position = -3.2
    Lsm5.Hardware.CpFocus.MoveRelative -0.5
    MsgBox Lsm5.DsRecording.Sample0Z
    Set c = Lsm5.Hardware.CpServos
    Debug.Print c.count
    count = c.count
    Set a = Lsm5.StartScan
    For i = 0 To count
        c.Select (i)
        Debug.Print c.name
        Debug.Print i & " " & c.name & " " & c.Exist(c.name)
    Next
    Do While a.IsBusy
        Sleep (10)
        
            c.Select (3)
            Debug.Print i & " " & Lsm5.Hardware.CpServos.name
        
    Loop
End Function

Function doctest()

    Dim xd As DsRecordingDoc
    
    Set xd = Lsm5.NewScanWindow
    
    Lsm5.StartScan
    
    Sleep 5000
    
    FindCells
End Function

Sub justimage()
    
    
    Dim xs
    Dim ys
    Dim spacing
    
    Dim dr As DsRecordingDoc
    Set dr = Lsm5.DsRecordingActiveDocObject
    
    ifilename = "zoomf_" & CStr(WellNumber) & "_" & CStr(ScanNumber) & "_" & CStr(i)
    j = dr.SaveToDatabase(fname & "\" & dbname, ifilename)
    CurrentImage = fname & "\" & ifilename & ".lsm"
    
    spacing = getSpacing
    ys = 1000000# * dr.Recording.LineSpacing
    xs = 1000000# * dr.Recording.SampleSpacing
    
    xp = dr.Recording.Sample0X
    yp = dr.Recording.Sample0Y
    
    Dim xc
    Dim yc
    Dim ac
    
    process_closeup Lsm5.DsRecordingActiveDocObject, ac, xc, yc
    
    nx = Lsm5.DsRecording.SamplesPerLine
    ny = Lsm5.DsRecording.LinesPerFrame
        
    xf = xp + (xc - nx / 2#) * xs
    yf = yp - (-yc + ny / 2#) * ys
    
    ixf = CInt(xc)
    iyf = CInt(yc)
    ll = 10
    circ1 = dr.VectorOverlay.AddSimpleDrawingElement(eDrawingModeClosedPolyLine, _
                ixf, iyf - ll, ixf, iyf + ll)
    circ1 = dr.VectorOverlay.AddSimpleDrawingElement(eDrawingModeClosedPolyLine, _
                ixf - ll, iyf, ixf + ll, iyf)
                
    UserForm1.txtfcsx.Text = xf
    UserForm1.txtfcsy.Text = yf
    
        nv = dr.VectorOverlay.GetNumberDrawingElements
    For ii = 0 To nv - 1
        dr.VectorOverlay.ElementColor(ii) = RGB(255, 255, 0)
    Next

    
End Sub
    
