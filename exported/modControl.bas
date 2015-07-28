Attribute VB_Name = "modControl"
Dim well_center_x As Double
Dim well_center_y As Double

''CAUTION, ATTENTION
''  cpstage.positionX returns the negative of what the aim gui displays!!!!!!

Function scan_over_plate()

'' May 1, 2007  need to change the path of the laser to move to the closest well,
'' not go all the way back to the top.

    Dim nx As Integer
    Dim ny As Integer
    
    nx = xWellCountX
    ny = xWellCountY
    
    
    '' global variables already defined
    '' xFrameWidth, xFrameHeight, xWellWidth, xWellHeight, xWellRadius,
    '' nWellSpacingX, nWellSpacingY, xWellCountY, stageWellOriginX, stageWellOriginY
    
    Dim xc As Double
    Dim yc As Double
    
    xc = stageWellOriginX
    yc = stageWellOriginY
    
    del = 1000 * nWellSpacingX
    
    fnum = 0
    'fname = "c:\chris\May01_fcs"
    Dim xrowcenter As Double
    Dim yrowcenter As Double
    
    For i = 0 To nx - 1
        Lsm5.Hardware.CpStages.PositionX = xc - del * i
        xrowcenter = xc - del * i
        While Lsm5.Hardware.CpStages.IsBusy
            Sleep 200
        Wend
            
        'Sleep 2000
        
        If (i Mod 2) = 0 Then
            jstart = 0
            jstop = ny - 1
            jstep = 1
        Else
            jstart = ny - 1
            jstop = 0
            jstep = -1
        End If
        
        For j = jstart To jstop Step jstep
        
            
            Lsm5.Hardware.CpStages.PositionY = yc + del * j
            yrowcenter = yc + del * j
            While Lsm5.Hardware.CpStages.IsBusy
                Sleep 200
            Wend
            
'            If j = 0 And i > 0 Then
'                Sleep 5000
'
'            Else
'                Sleep 2000
'            End If
                
            Open "c:\temp\wellpositions.dat" For Append As #4
            Print #4, "Position " & i & " " & j; ""
            Print #4, "output : "; Lsm5.Hardware.CpStages.PositionY & " " & Lsm5.Hardware.CpStages.PositionY
            Close #4
            Debug.Print "output : " & xc - del * i & " " & yc + del * j
            'ScanNewOverview
            
            '''right here create the new folder and database
             fname = basefname & "\" & OrfNamesList(WellNumber) & "_Well_" & WellNumber & ".mdb"
            
            dbname = OrfNamesList(WellNumber) & "_Well_" & WellNumber & ".mdb"
            
            Lsm5.NewDatabase fname
            DoEvents
            scan_in_well xrowcenter, yrowcenter, CInt(gScanX), CInt(gScanY)
            
            Lsm5.CloseAllDatabaseWindows
            DoEvents
            'takeFCS 0.0000013, 0.0000024, 0, fnum
            Sleep 1500
            fnum = fnum + 1
            WellNumber = WellNumber + 1
            
            DoEvents
            
            If killall = 1 Then End
        Next
        
        AddWater
    Next
    
End Function


Function scan_in_well(wellx As Double, welly As Double, num_scans_x As Integer, _
            num_scans_y As Integer)
 
 
    '' the variable fname needs to change here,
    '' a new folder needs to be created for each different well
    
    
    well_center_x = wellx
    well_center_y = welly
    
    Dim startx As Double
    Dim starty As Double
    
    Dim framesperwell As Integer
    Dim framesperwell_x As Integer
    Dim framesperwell_y As Integer
    
    framesperwell_x = CInt(xWellWidth / (xFrameWidth * 0.001))
    Debug.Print xWellWidth & " " & xFrameWidth & " " & "Widths"
    framesperwell_y = framesperwell_x
    
    ''how many frames are in each well
    framesperwell = framesperwell_x * framesperwell_y
        
    '' calculatae the starting positions for x and y
    
    If num_scans_x > 0 Then
        startx = well_center_x - 1000 * xWellWidth * (num_scans_x - 1) / 2# / (num_scans_x + 1)
        'startx = well_center_x - 1000 * (xWellWidth / 2# + xWellWidth / (num_scans_x + 1))
    Else
        startx = well_center_x - 1000 * xWellWidth * (framesperwell_x - 1) / 2# / (framesperwell_x + 1)
        'startx = well_center_x - 1000 * (xWellWidth / 2# + xWellWidth / (framesperwell_x + 1))
        num_scans_x = framesperwell_x
    End If
    
    
    If num_scans_y > 0 Then
        starty = well_center_y - 1000 * xWellWidth * (num_scans_y - 1) / 2# / (num_scans_y + 1)
        'starty = well_center_y - 1000 * (xWellHeight / 2# + xWellHeight / (num_scans_y + 1))
    Else
        starty = well_center_y - 1000 * xWellWidth * (framesperwell_y - 1) / 2# / (framesperwell_y + 1)
        'starty = well_center_y - 1000 * (xWellHeight / 2# + xWellHeight / (framesperwell_y + 1))
        num_scans_y = framesperwell_y
    End If
   
    Debug.Print "Dims " & framesperwell & " " & framesperwell_x & " " & framesperwell_y
    Debug.Print "Starts " & startx & " " & starty
    
    Dim i As Integer
    Dim j As Integer
    
    Dim xcurrent As Double
    Dim ycurrent As Double
    Dim xdel As Double
    Dim ydel As Double
    
    xcurrent = startx
    ycurrent = starty
    
    xdel = (xWellWidth * 1000) / (num_scans_x + 1)
    ydel = (xWellWidth * 1000) / (num_scans_y + 1)
    
    Lsm5.CloseAllImageWindows 0
    Dim odoc As DsRecordingDoc
    
    Debug.Print "Dels " & xdel & " " & ydel
    Debug.Print "Stage Position" & Lsm5.Hardware.CpStages.PositionX
    Debug.Print "Stage Position" & Lsm5.Hardware.CpStages.PositionY
    ScanNumber = 0
    
    '' set this to zero at the start of every well
    gNumberOfCells = 0
    For i = 0 To num_scans_x - 1
        xcurrent = startx + xdel * i
        Debug.Print i & " " & xcurrent & " " & xcurrent - wellx
        Lsm5.Hardware.CpStages.PositionX = xcurrent
        While Lsm5.Hardware.CpStages.IsBusy
            Sleep 100
        Wend
        
        For j = 0 To num_scans_y - 1
            ycurrent = starty + ydel * j
            Debug.Print "    * " & j & " " & ycurrent & " " & ycurrent - welly
            Lsm5.Hardware.CpStages.PositionY = ycurrent
            
            'Debug.Print i & " " & j & "Stage Position x" & Lsm5.Hardware.CpStages.PositionX
            'Debug.Print i & " " & j & "Stage Position Y" & Lsm5.Hardware.CpStages.PositionY
    
            
            Sleep 500
            Debug.Print i & " " & j & "Stage Position x" & Lsm5.Hardware.CpStages.PositionX
            Debug.Print i & " " & j & "Stage Position Y" & Lsm5.Hardware.CpStages.PositionY
            '' set scanoverview to a documment or something
            
'' uncommet
            setTrack "overviewscan"
            Set odoc = ScanNewOverview
            
            Dim spacing
            spacing = getSpacing
            xspace = spacing(0) * 1000000#
            yspace = spacing(1) * 1000000#
    
   
            xpixels = odoc.Recording.SamplesPerLine
            ypixels = odoc.Recording.LinesPerFrame
            dzoom = getZoom
''' uncommet
            
            On Error GoTo FixErr
            
            ''odoc.NeverAgainScanToTheImage
            
            '' fname name is the current directory

            odoc.SaveToDatabase fname & "\" & dbname, OrfNamesList(WellNumber) & "_wellscan_" & _
                    thisday & "_" & WellNumber & "_" & ScanNumber
            CurrentImage = fname & "\" & OrfNamesList(WellNumber) & "_wellscan_" _
                            & thisday & "_" & WellNumber & "_" & ScanNumber & ".lsm"
            
            Sleep 100
            'odoc.CloseAllWindows
            
            DoEvents
            
            If killall = 1 Then End
            FindCells
''''uncomment
            DoEvents
            If gNumberOfCells > 9 Then
                DoEvents
                Exit For
            End If
            
            
            ScanNumber = ScanNumber + 1
FixErr:
            If killall = 1 Then End
        Next
        
        If gNumberOfCells > 9 Then
                DoEvents
                Exit For
        End If
        
        'Lsm5.Hardware.CpStages.PositionX = wellx
        'Lsm5.Hardware.CpStages.PositionY = welly
        
        'While Lsm5.Hardware.CpStages.IsBusy
        '    Sleep 100
        'Wend
    Next
End Function

Function test_control()
    xFrameHeight = Lsm5.DsRecording.FrameHeight
    xFrameWidth = Lsm5.DsRecording.FrameWidth
    
    xWellWidth = 0.707 * 6.32 * 1000
    xWellHeight = 0.707 * 6.32 * 1000
    
    Dim xp As Double
    Dim yp As Double
    
    xp = Lsm5.Hardware.CpStages.PositionX
    yp = Lsm5.Hardware.CpStages.PositionY
    
    Debug.Print "stage positions" & " " & xp & " " & yp
    scan_in_well xp, yp, 2, 2
End Function

Function test_plate_scan()
    'nWellSpacingX = 9000
    'nWellSpacingY = 9000
    
    Open "c:\temp\wellpositions.dat" For Output As #4
    Print #4, "Position of stage in well center"
    
    Close #4
        
    pumpoffsetx = (30683.75 - 20838.75) - 500
    pumpoffsety = (-16979 - (-42138)) - 200
    
    reset_frame
    Sleep 200
    DoEvents
    
    xFrameHeight = Lsm5.DsRecording.FrameHeight
    xFrameWidth = Lsm5.DsRecording.FrameWidth
    
    If xFrameHeight = -1 Or xFrameWidth = -1 Then
        xFrameHeight = 255#
        xFrameWidth = 255#
    End If
        'Lsm5.StartScan
    
    While Lsm5.Hardware.CpScancontrol.IsGrabbing
        Sleep 200
    Wend
    
    'xFrameHeight = Lsm5.DsRecording.FrameHeight
    'xFrameWidth = Lsm5.DsRecording.FrameWidth
    
    'xWellWidth = 0.707 * 6.32 * 1000
    'xWellHeight = 0.707 * 6.32 * 1000
        
    If marked = 0 Then
        stageWellOriginX = Lsm5.Hardware.CpStages.PositionX
        stageWellOriginY = Lsm5.Hardware.CpStages.PositionY
    End If
    
    'xWellCountX = 2
    'xWellCountY = 3
    zlogfile = fname & "\" & "focus.log"
    
    Open zlogfile For Output As #2
    Print #2, "Log file for focus positions"
    Print #2, "---------------------------------"
    Close #2
    scan_over_plate
End Function

Function closeall()
    Lsm5.CloseAllImageWindows 0
End Function
