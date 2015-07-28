Attribute VB_Name = "modSetFocus"




Sub zautofocus()

'    focusaotfpower = 1
'    pmtaotfpower = 1
'    cx = Lsm5.Hardware.CpStages.PositionX
'    cy = Lsm5.Hardware.CpStages.PositionY
'
'    Lsm5.Hardware.CpStages.PositionX = cx
'    Lsm5.Hardware.CpStages.PositionY = cy - 0
    '' use sample0Z to control where the first z position is
        
    
    
    setTrack "focus"
    
    setaotf 1#
    Sleep 500
    
    stacksize = 210
    
    z = Lsm5.Hardware.CpFocus.position
    zstart = stacksize / 2
    abs_zstart = z - zstart
    Lsm5.DsRecording.Sample0Z = zstart
    
    
    Debug.Print "ZZZZZZZZZZZZZZZZ123 " & z
    Debug.Print "ABS Start " & abs_zstart
    Lsm5.DsRecording.ScanMode = "zscan"
    
    Lsm5.DsRecording.SpecialScanMode = "OnTheFly"
    'Lsm5.DsRecording.SpecialScanMode = "FocusStep"
    'Lsm5.DsRecording.FramesPerStack = 500
    
    sp = Lsm5.DsRecording.FrameSpacing
    Lsm5.DsRecording.FrameSpacing = 1
    sp = Lsm5.DsRecording.FrameSpacing
    Lsm5.DsRecording.FramesPerStack = CInt(stacksize / sp)
   
    Debug.Print Lsm5.DsRecording.FramesPerStack
    Dim d As DsRecordingDoc
    t = Lsm5.DsRecording.StackDepth
    Set d = Lsm5.NewScanWindow
    
    Lsm5.StartScan
    
    
    While Lsm5.Hardware.CpScancontrol.IsGrabbing
        
      
'        c = Lsm5.Hardware.CpFocus.Position
'        Debug.Print c
'
            
        DoEvents
        Sleep 20
    Wend
    DoEvents
    'setaotf 0.1
    z = Lsm5.Hardware.CpFocus.position
    Debug.Print z
    m = getLineMax(d, 0)
    Moveto = abs_zstart + m * sp + 1.5
    
    zfocuspostion = Moveto
    Debug.Print "Max z " & m
    Debug.Print "Moveto " & Moveto
    Lsm5.Hardware.CpFocus.position = Moveto
    On Error GoTo err
    Open zlogfile For Append As #2
    Print #2, "*****Cover slip located:"
    Print #2, abs_zstart + m * sp
    Close #2
    reset_frame
err:
    DoEvents
    d.CloseAllWindows
    Lsm5.DsRecording.ScanMode = "Plane"
    Exit Sub
    
End Sub

Function zfocusfcs(zname As String)

'    cx = Lsm5.Hardware.CpStages.PositionX
'    cy = Lsm5.Hardware.CpStages.PositionY
'
'    Lsm5.Hardware.CpStages.PositionX = cx
'    Lsm5.Hardware.CpStages.PositionY = cy - 0
    '' use sample0Z to control where the first z position is
    
    Dim z As Double
    Dim z0  As Double
    
    setaotf zoomaotfpower
    Sleep 200
    
    
    ''this should already be done
    ''setTrack "zoomscan"
    stacksize = 20
    
    z = Lsm5.Hardware.CpFocus.position
    z0 = z
    On Error GoTo ErrorHand
    zstart = stacksize / 2
    abs_zstart = z - zstart
    Lsm5.Hardware.CpFocus.position = abs_zstart
    
    While Lsm5.Hardware.CpFocus.IsBusy
        Sleep 200
    Wend
    
    Lsm5.DsRecording.Sample0Z = 0 'abs_zstart
    
    
    Debug.Print "ZZZZZZZZZZZZZZZZ123 " & z
    Debug.Print "ABS Start " & abs_zstart
    Lsm5.DsRecording.ScanMode = "zscan"
    
    'Lsm5.DsRecording.SpecialScanMode = "OnTheFly"
    Lsm5.DsRecording.SpecialScanMode = "FocusStep"
    'Lsm5.DsRecording.FramesPerStack = 500
    
    sp = Lsm5.DsRecording.FrameSpacing
    Lsm5.DsRecording.FrameSpacing = 1
    sp = Lsm5.DsRecording.FrameSpacing
    Lsm5.DsRecording.FramesPerStack = CInt(stacksize / sp)
   
    Dim fps As Integer
    fps = CInt(stacksize / sp)
    
    Debug.Print Lsm5.DsRecording.FramesPerStack
    Dim d As DsRecordingDoc
    t = Lsm5.DsRecording.StackDepth
    
    Lsm5.CloseAllImageWindows False
    
    zerotest = 0
    zcount = 0
    
    While (zerotest < 1) And (zcount < 3)
    
        Set d = Lsm5.NewScanWindow
        
        Lsm5.StartScan
        
        
        While Lsm5.Hardware.CpScancontrol.IsGrabbing
            
          
    '        c = Lsm5.Hardware.CpFocus.Position
    '        Debug.Print c
     
            DoEvents
            Sleep 20
        Wend
        
        While Lsm5.Hardware.CpFocus.IsBusy
            Sleep 200
        Wend
        
        Sleep 200
        DoEvents
        
        z = Lsm5.Hardware.CpFocus.position
        'Debug.Print z
        m = getLineMax(d, 1)
        Debug.Print "this time " & m
        If m < fps - 1 Then
            zerotest = 1
        Else
            zcount = zcount + 1
        End If
        
    Wend
    
    If m >= fps - 1 Then m = 10
    
    Debug.Print "Z pos " & m
    circ1 = d.VectorOverlay.AddSimpleDrawingElement(eDrawingModeClosedPolyLine, _
                 0, m, 512, m)
                 
    'For ii = 0 To nv - 1
        d.VectorOverlay.ElementColor(0) = RGB(255, 255, 0)
    'Next
                 
    Moveto = abs_zstart + m * sp
    
    'zfocuspostion = Moveto
    Debug.Print "Max z " & m
    Debug.Print "Moveto " & Moveto
'    g = Trim(CStr(Moveto))
'    g = CDbl(Strings.Left(g, 5))
    
    Lsm5.Hardware.CpFocus.position = Moveto
    
    While Lsm5.Hardware.CpFocus.IsBusy
        Debug.Print Lsm5.Hardware.CpFocus.position
        Sleep 200
    Wend

    'Open zlogfile For Append As #2
    'Print #2, "*****Cover slip located:"
    'Print #2, abs_zstart + m * sp
    'Close #2
    DoEvents
    z = Lsm5.Hardware.CpFocus.position
    'd.CloseAllWindows
    
    If StrComp(zname, "none") <> 0 Then
        j = d.SaveToDatabase(fname & "\" & dbname, zname)
    End If
    
    Sleep 500
    reset_frame
    zfocusfcs = Moveto
    'Stop
    'd.CloseAllWindows
    
    Exit Function
ErrorHand:
    
    Lsm5.Hardware.CpFocus.position = Moveto
    
    While Lsm5.Hardware.CpFocus.IsBusy
        Debug.Print Lsm5.Hardware.CpFocus.position
        Sleep 200
    Wend
    
    reset_frame
    zfocusfcs = z0
        
End Function
Function getLineMax(d As DsRecordingDoc, stype As Integer) As Long
    
    On Error GoTo ErrorHand
    Dim x As Long
    Dim z As Long
    
    x = d.GetDimensionX
    z = d.GetDimensionZ
    
    Dim xline
    Dim thistotal
    
    Dim ztotal()
    Dim maxz
    ReDim ztotal(z)
    For k = 0 To z - 1
        xline = d.ScanLine(0, 0, k, 0, 0, 0)
        'xline = d.ScanLine(0, 0, 0, k, 0, 0)
        thistotal = 0
        For i = 0 To x - 1
            thistotal = thistotal + xline(i)
        Next
        ztotal(k) = thistotal
        
        
        If stype = 0 Then
            If (thistotal / x) > 2000 Then
                maxz = thistotal
                kmax = k
            End If
        End If
        
        If stype = 1 Then
           If thistotal >= maxz Then
               maxz = thistotal
               kmax = k
           End If
        End If
    Next
    
    getLineMax = kmax
    
    
    Exit Function
    
ErrorHand:
    
    getLineMax = z / 2
End Function


Sub testfcsfocus()
    Lsm5.CloseAllImageWindows False
    
    v = zfocusfcs("none")
    
End Sub



