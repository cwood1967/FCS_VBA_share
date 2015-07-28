Attribute VB_Name = "modtest"

'Private Declare Function Hello Lib "u:\cjw\testdll\hello.dll" (ByVal n As Long, ByVal m As Long, ByRef x As Long) As Long



Function dtest()


     Dim dimx As Long
    Dim dimy As Long
    Dim doc As DsRecordingDoc
    Set doc = Lsm5.DsRecordingActiveDocObject
    dimx = doc.GetDimensionX
    dimy = doc.GetDimensionY
    
    Lsm5.StopScan
    
    
    Dim spl As Long
    Dim bpp As Long
    Dim j As Long
    
    npixels = dimx * dimy
    
    
    
    ReDim ximage(0 To dimx - 1)
    ReDim qimage(0 To dimx - 1, dimy - 1)
    
    t1 = Now
    For d = 0 To 0
    
    f = (doc.GetSubregion(0, 0, 0, 0, 0, 1, 1, 1, 1, 512, 512, 1, 1, 12))
    
    Dim fd() As Long
    ReDim fd(0 To dimx * dimy - 1)
    
    For j = 0 To dimx * dimy - 1
        'ximage(j) = doc.ScanLine(0, 0, 0, j, spl, bpp)
        fd(j) = CLng(f(j))
    Next

    
    'ximage = CLng(ximage)
    
'    For i = 0 To dimx - 1
'        For j = 0 To dimy - 1
'            qimage(i, j) = ximage(i)(j)
'        Next
'    Next
    
    n = 200
    m = 200
    

    Dim y As Long
   
    
'    For i = 0 To n - 1
'        For j = 0 To m / 2
'            x(j, i) = 200#
'        Next
'    Next
    
    y = Hello(dimx, dimy, fd(0))
    
    'MsgBox y & " " & x(0, 0) & " " & x(2, 2)
    
    Next
    t2 = Now
End Function


Function htest()

    Dim r As DsRecording
    Dim s As CpServos
    
    Set r = Lsm5.DsRecording
    Set s = Lsm5.Hardware.CpServos
    
    r.Sample0X = 32
    
    Debug.Print "Sample Position" & " " & r.Sample0X
    
    Lsm5.StartScan
    Sleep 500
    Lsm5.StopScan
    
    Debug.Print "Sample Position" & " " & r.Sample0X
    
    sc = s.count
    
    'r.Sample0X = 14#
    
    Debug.Print "Sample Position" & " " & r.Sample0X
    For i = 0 To sc - 1
        s.Select (i)
        'Debug.Print s.name & " " & s.Position
    Next
    
End Function

Sub testsettrack()
    zoomaotfpower = 0.5
    pmtaotfpower = 1#
    
    setTrack "zoomscan"
End Sub

Sub testfcs()
    fcspinhole = 70
    setaotf 0.38
    laser488 = 0.02
    laser561 = 0.04
    setaotf (0.001)
    
    fname = "c:\temp"
    WellNumber = 1
    ScanNumber = 1
    Dim tlist(0 To 2)
    tlist(0) = "a0"
    tlist(1) = "a1"
    tlist(2) = "a2"
    OrfNamesList = tlist
     
    Dim z As Double
    z = Lsm5.Hardware.CpFocus.position
    takeFCS -2.32, -10.58, z, 0
End Sub
Sub deltracks()
    Dim t As DsTrack
    Dim trackexists As Boolean
    Dim b As Boolean
    Dim suc As Integer
    Dim i As Integer
    
    count = Lsm5.DsRecording.TrackCount

    trackexists = False
    
    For i = count - 1 To 0 Step -1
        
        tracknames = "Ratio1,Bleach1,Lambda"
        Set t = Lsm5.DsRecording.TrackObjectByIndex(i, suc)
        nt = t.name
        'Lsm5.DsRecording.TrackRemove i
        If InStr(1, tracknames, nt, vbTextCompare) = 0 Then
            Lsm5.DsRecording.TrackRemove i
            't.LoadConfigurationSetting "overviewscan"
        End If
    Next
    
    Lsm5.DsRecording.TrackAddNew "NewTrack"
    
    Set t = Lsm5.DsRecording.TrackObjectByName("NewTrack", suc)
    
    t.LoadConfigurationSetting ("zoomscan")
    bs = t.BeamSplitterCount
    Dim bso As DsBeamSplitter
    Dim dc As DsDataChannel
    dcn = t.DataChannelCount
    For i = 0 To dcn - 1
        Set dc = t.DataChannelObjectByIndex(i, suc)
        nm = dc.name
        If InStr(1, nm, "APD", vbTextCompare) > 0 Then
            MsgBox nm & "  " & dc.Acquire
        End If
    Next
    
    Set t = Nothing
    
    Exit Sub
    If Not trackexists Then
        b = Lsm5.DsRecording.TrackAddNew("new track")
        Set t = Lsm5.DsRecording.TrackObjectByName("new track", suc)
    Else
        Set t = Lsm5.DsRecording.TrackObjectByName(trackname, suc)
    
    End If
    
    b = t.LoadConfigurationSetting(trackname)
    
    count = Lsm5.DsRecording.TrackCount
    
    'just to make sure
    
    For i = 0 To count - 1
        Set t = Lsm5.DsRecording.TrackObjectByIndex(i, suc)
        tname = t.name
        If StrComp(tname, trackname, vbTextCompare) = 0 Then
            t.Acquire = True
        Else
            t.Acquire = False
        End If
            
    Next
    
    Sleep 1000
End Sub


Sub rtest()

    x = Lsm5.DsRecording.FrameHeight
    y = Lsm5.DsRecording.FrameWidth
    
End Sub


Sub tcurrent()
    
    Dim lasers As CpLasers
    
    Set lasers = Lsm5.Hardware.CpLasers
    
    n = lasers.count
    
    For i = 0 To n - 1
        lasers.Select ("Argon/2")
        MsgBox lasers.name & " " & lasers.tubecurrent
    Next
    
End Sub

Sub testsetfcslaser()

    level = 0.02
    Dim fcs As AimFcsController
    Dim beampath As AimFcsBeamPathParameters
    Dim hinfo As AimFcsHardwareInformation
    
    Set fcs = Lsm5.ExternalDsObject.FcsController
    Set hinfo = fcs.HardwareInformation
    
    nlasers = hinfo.GetNumberAttenuators
    
    Set beampath = fcs.BeamPathParameters
    
    For i = 0 To nlasers - 1
        Debug.Print i

        If beampath.AttenuatorOn(i) Then

            ap = beampath.AttenuatorPower(i)
            Debug.Print "     " & ap
            'MsgBox i & " " & ap
            'beampath.AttenuatorPower(i) = 0.0321
        End If
    Next
        
    beampath.AttenuatorPower(3) = 0.0421
    beampath.AttenuatorPower(5) = 0.0121
    'rdiam = 0.000001 * diam
    'MsgBox beampath.PinholeDiameter(0)
    'beampath.PinholeDiameter(0) = rdiam
    
    Set beampath = Nothing
    Set fcs = Nothing
    
    Sleep 200
End Sub

Sub testidl()
    
    Dim x As conIDL
    Set x = New conIDL
    z = x.sendCloseUp("")
    
End Sub
