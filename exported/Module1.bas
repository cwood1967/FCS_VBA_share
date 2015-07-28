Attribute VB_Name = "Module1"
'Option Explicit
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim xpoints
Dim ypoints
Dim areas
Public fname
Public basefname
Public dbname
Public basedbname As String

Public xspace
Public yspace
Public xpixels
Public ypixels
Public dzoom

Sub Main()
    
    fcsaotfpower = 0.38
    zoomaotfpower = 0.38
    pmtaotfpower = 1#
    
    UserForm1.Show
    
    On Error GoTo ErrHandler
    
    Exit Sub
    
ErrHandler:
    MsgBox err.Description
End Sub

Function newdb()
    
    basedbname = Trim(UserForm1.TextBox3.Text)
    
    Dim r As String
    r = Strings.Right(basedbname, 4)
    
    If StrComp(r, ".mdb", vbTextCompare) <> 0 Then
        basedbname = basedbname & ".mdb"
    End If
    
    dbname = getdname(UserForm1.basedir, basedbname)
    basefname = UserForm1.basedir & "\" & dbname
    fname = UserForm1.basedir & "\" & dbname
    'fname = "c:\chris\newtest9.mdb"
    'dbname = getdbname(fname)
    
    
    Lsm5.NewDatabase fname
    
End Function

Function welldb()
    
End Function
'' this
Function getdname(basedir As String, dbbasename As String)
    Dim fso As Scripting.FileSystemObject
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim bf As Folder
    
    bl = Strings.Left(basedir, 2)
    
    If Not fso.DriveExists(bl) Then
        MsgBox "drive can't be accessed"
        
        End
        'Err.Raise -999, , "Drive can't be accessed"
    End If
    If Not fso.FolderExists(basedir) Then
        fso.CreateFolder basedir
    End If
    
    Set bf = fso.GetFolder(basedir)
    
    Dim bfolders As Folders
    Set bfolders = bf.SubFolders

    Dim name As String
    Dim bsf As Folder
    Dim sprefix As String
    Dim nprefix As Long
    Dim pmax As Long
    
    pmax = -1
    
    For Each bsf In bfolders
        
        name = bsf.name
        
        lname = Len(name)
        isadb = InStr(1, name, dbbasename, vbTextCompare)
        
        
        If isadb > 0 Then
            
            
            sprefix = VBA.Strings.Left(name, 3)
            
            If IsNumeric(sprefix) Then
                nprefix = CLng(sprefix)
                
                If nprefix > pmax Then
                    pmax = nprefix
                End If
                
            End If
            
        End If
            
    Next
        
    Dim np As Long
    np = pmax + 1
    snp = Trim(Str(np))
    l1 = Len(snp)
    zp = 3 - l1
    
    getdname = String(zp, "0") & snp & "_" & basedbname


End Function
Function ScanNewOverview()
    
    Dim tlsm As DsTrack
    Dim d As DsRecordingDoc
    Dim t As DsTrack
    Dim overviewexists As Boolean
    Dim b As Boolean
    Dim suc As Integer
    Dim i As Integer
    
    setaotf 1
    
    Sleep 500
    
    count = Lsm5.DsRecording.TrackCount
    
    setTrack "focus"
    
    
    zautofocus
    DoEvents
    Sleep 500
    reset_frame
    
    setaotf pmtaotfpower
    setTrack "overviewscan"
    Sleep 1000
    
    Set d = Lsm5.StartScan
    Set scanned = Lsm5.StartScan

        Do While Lsm5.Hardware.CpScancontrol.IsGrabbing
            DoEvents
            Sleep 1000
        Loop
    Dim j

    'Stop
    
    Set d = Lsm5.DsRecordingActiveDocObject
    
    'j = scanned.SaveToDatabase(fname & "\" & dbname, "overview")
    
    'MsgBox d.Recording.TrackObjectByIndex(0, suc).Name
    
    Set ScanNewOverview = d
End Function


''' this set track written, nov. 22 2007, to delete tracks rather than create new ones
Function setTrack(trackname As String)

Dim t As DsTrack
    Dim trackexists As Boolean
    Dim b As Boolean
    Dim suc As Integer
    Dim i As Integer
    
    count = Lsm5.DsRecording.TrackCount

    trackexists = False
    
    '''start at the highest number, so removing tracks does not reorder lower ones
    
    '' common names of the internal tracks
    
     tracknames = "Ratio1,Bleach1,Lambda"
     
    '' go through all of the tracks and delete all that are not in tracknames
    For i = count - 1 To 0 Step -1
       
        Set t = Lsm5.DsRecording.TrackObjectByIndex(i, suc)
        nt = t.name
        'Lsm5.DsRecording.TrackRemove i
        If InStr(1, tracknames, nt, vbTextCompare) = 0 Then
            Lsm5.DsRecording.TrackRemove i
            
        End If
    Next
    
    '''Create a new track
    Lsm5.DsRecording.TrackAddNew "NewTrack"
    
    Set t = Lsm5.DsRecording.TrackObjectByName("NewTrack", suc)
    
    ''' load a saved configuration into the track
    t.LoadConfigurationSetting (trackname)
    
    '''go through to datachannel to see if any are using apds and set the aotf accordingly
    bs = t.BeamSplitterCount
    Dim bso As DsBeamSplitter
    Dim dc As DsDataChannel
    dcn = t.DataChannelCount
    Dim bapd As Integer
    bapd = 0
    'setaotf zoomaotfpower
    For i = 0 To dcn - 1
        Set dc = t.DataChannelObjectByIndex(i, suc)
        nm = dc.name
        If InStr(1, nm, "APD", vbTextCompare) > 0 Then
            If dc.Acquire Then
                If zoomaotfpower > 0.38 Then
                    zoomaotfpower = 0.38
                    'MsgBox "the aoft was set to 10%"
                End If
                setaotf zoomaotfpower
                Sleep 200
                bapd = 1
            End If
        End If
    Next
    
    If bapd = 0 Then
        setaotf pmtaotfpower
    End If
    
    'getAotf
    Set t = Nothing
    Set dc = Nothing
    Set bs = Nothing
    Exit Function
End Function


''' this one deprecated nov 22, 2007 because apd do not work in multitrack mode.
''' adding additional tracks causes AIM to be multitrack
Function oldsetTrack(trackname As String) As Boolean

    Dim t As DsTrack
    Dim trackexists As Boolean
    Dim b As Boolean
    Dim suc As Integer
    Dim i As Integer
    
    count = Lsm5.DsRecording.TrackCount

    trackexists = False
    
    For i = 0 To count - 1
        
        Set t = Lsm5.DsRecording.TrackObjectByIndex(i, suc)
        tname = t.name
        If StrComp(tname, trackname, vbTextCompare) = 0 Then
            t.Acquire = True
            trackexists = True
        Else
            t.Acquire = False
        End If
            
    Next
    Set t = Nothing
    
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

End Function

Sub GoProcess()
    
    Dim d As DsRecordingDoc
    Set d = Lsm5.DsRecordingActiveDocObject
    
    d.NeverAgainScanToTheImage
    d.SaveToDatabase fname & "\" & dbname, "wellscan_" & WellNumber & "_" & ScanNumber
    CurrentImage = fname & "\" & "wellscan_" & WellNumber & "_" & ScanNumber & ".lsm"
    Set d = Nothing
    
    FindCells
End Sub

Function FindCells()
    Dim loc
    Dim npoints
    Dim i
    
    On Error GoTo ErrorHand
    setaotf zoomaotfpower
    Sleep 200
    'setaotf 0.38
    xStageX0 = Lsm5.Hardware.CpStages.PositionX
    xStageY0 = Lsm5.Hardware.CpStages.PositionY
    xStageZ0 = Lsm5.Hardware.CpFocus.position
    
    '' get image returns coordinates, but it also writes data to the idl object in the
    '' form object
    
    ''07-23-2007 do this now so the document can be closed
    ''spacing is returning in meters, convert to micro meters
'    Dim spacing
'    spacing = getSpacing
'
'    xspace = spacing(0) * 1000000#
'    yspace = spacing(1) * 1000000#
'
'    Dim d As DsRecordingDoc
'    Set d = Lsm5.DsRecordingActiveDocObject
'    xpixels = d.Recording.SamplesPerLine
'    ypixels = d.Recording.LinesPerFrame
'    dzoom = getZoom
    ''
    
    '' this gets the image data
    loc = getImage
    If Not IsArray(loc) Then
        FindCells = -1
        Exit Function
    End If
    
    Dim nb
    nb = UBound(loc)
    npoints = nb / 2 ' integer division
    
    ReDim xpoints(npoints)
    ReDim ypoints(npoints)
    
    '' now use
    
    xpoints = UserForm1.oIDL.get_xCoords
    ypoints = UserForm1.oIDL.get_yCoords
    areas = UserForm1.oIDL.getAreas
    
    UserForm1.oIDL.closeidl
    
'    For i = 0 To npoints
'        xpoints(i) = loc(i)
'        ypoints(i) = loc(i + (npoints + 1) / 2)
'    Next
    
        
    'd.CloseAllWindows
    'Stop
    If areas(0) > 0 Then
        If UserForm1.cbCloseUp Then
        
            testzoom
            
        End If
    End If
ErrorHand:

End Function

Function testzoom()
    Dim d As DsRecordingDoc
'    Dim xylog As String
'
'    xylog = fname & "\" & OrfNamesList(WellNumber) & "_" & WellNumber _
'                    & "_" & CStr(ScanNumber) & "_xylog.txt"
'
'    Open xylog For Output As #4
'
'    Print 4, "FCS position " & OrfNamesList(WellNumber) & "_" & WellNumber _
'                    & "_" & CStr(ScanNumber)
'    Close #4
        
        
    setfcspinhole imagepinhole
    
    Set d = Lsm5.DsRecordingActiveDocObject
    
    Dim xdim
    Dim ydim
     
    ''xdim and ydim are the dimensions of the overview image (these should be the same???)
    '' x and y spacing xspace and yspace are defined in scope of the module
    
    ''xdim and ydim come back as meters
    xdim = xpixels 'd.Recording.SamplesPerLine
    ydim = ypixels 'd.Recording.LinesPerFrame
    
    setTrack "zoomscan"
    
    Dim ozoom
    ozoom = dzoom
    
    Debug.Print "************************************"
    'Debug.Print d.Recording.zoomx
    Debug.Print Lsm5.DsRecording.SampleSpacing
    Debug.Print Lsm5.DsRecording.LineSpacing
    Debug.Print Lsm5.DsRecording.SamplesPerLine
    Debug.Print Lsm5.DsRecording.LinesPerFrame
    Debug.Print "************************************"
    
    Lsm5.DsRecording.SamplesPerLine = 512
    Lsm5.DsRecording.LinesPerFrame = 512
    
    'Lsm5.DsRecording.zoomx  8
    'Lsm5.DsRecording.zoomy = 8
    Lsm5.Options.StatusDisplayDetectorGain = True
    Lsm5.Options.StatusDisplayPosition = True
    Lsm5.Hardware.CpScancontrol.GetScanState
  
    'xspace = Lsm5.DsRecording.SampleSpacing
    'yspace = Lsm5.DsRecording.LineSpacing
    Debug.Print "************************************"
    Dim xspot
    Dim yspot
    
    Dim xop
    Dim yop
    
    Dim ux
    
    ux = UBound(xpoints)
    Dim vnum
    vnum = 12
    If (ux < vnum) Then
        lk = ux
    Else
        lk = vnum
    End If
    
    '' the document d has already been closed
    'd.CloseAllWindows
    
    Dim nd As DsRecordingDoc
    Dim afterfcs As DsRecordingDoc
    '' this loop goes through each of the points and take a close up
    zoomzero = Lsm5.Hardware.CpFocus.position
    
    Dim ifilename As String
    Dim zmax1
    
    zmax = Lsm5.Hardware.CpFocus.position
    For i = 0 To lk
    
        'Lsm5.NewScanWindow
        Sleep 500
        'Debug.Print xdim & " " & ydim
        xop = xdim / 2
        yop = ydim / 2
        
        '' xspot converts from pixels to micor meters
        xspot = (xpoints(i) - xop) * xspace
        yspot = -(-ypoints(i) + yop) * yspace
        
        Debug.Print xspot & " " & yspot
        
        If areas(i) = -1 And lk = 0 Then Exit For
        slen = 3 * Sqr(areas(i))
        
        
        ''09/05/2007
        ''changed this to 512/slen
        Lsm5.DsRecording.zoomx = ozoom(0) * 512# / slen
        If Lsm5.DsRecording.zoomx > 18 Then Lsm5.DsRecording.zoomx = 18
        Lsm5.DsRecording.zoomy = ozoom(1) * 512# / slen
        If Lsm5.DsRecording.zoomy > 18 Then Lsm5.DsRecording.zoomy = 18
        'Lsm5.DsRecording.SampleSpacing = (xspace) * slen / 512
        'Lsm5.DsRecording.LineSpacing = (yspace) * slen / 512
        Lsm5.DsRecording.Sample0X = xspot
        Lsm5.DsRecording.Sample0Y = yspot
        
        Dim z_xdim
        Dim z_ydim
        
        Dim xsp As Double
        Dim ysp As Double
        
        z_xdim = Lsm5.DsRecording.SampleSpacing * 1000000#
        z_ydim = Lsm5.DsRecording.LineSpacing * 1000000#
        
        xsp = Lsm5.DsRecording.SampleSpacing
        ysp = Lsm5.DsRecording.LineSpacing
        
        'Debug.Print "zoom " & nd.Recording.zoomx
        Debug.Print "x space " & Lsm5.DsRecording.SampleSpacing
        Debug.Print "y space " & Lsm5.DsRecording.LineSpacing
        Debug.Print Lsm5.DsRecording.SamplesPerLine
        Debug.Print Lsm5.DsRecording.LinesPerFrame
        Debug.Print "zoom " & Lsm5.DsRecording.zoomx
        
        Lsm5.Hardware.CpFocus.position = zoomzero
        
        
        'setaotf 0.38
        'Sleep 100
        
        setTrack "zoomscan"
        Sleep 300
        
        'setaotf zoomaotfpower
        Sleep 100
        
        Dim zname As String
        zname = OrfNamesList(WellNumber) & "_zfocus_" & CStr(WellNumber) _
                    & "_" & CStr(ScanNumber) & "_" & CStr(i)
                    
        zmax = zfocusfcs(zname)
                
                
        Lsm5.Hardware.CpFocus.position = zmax
        
        While Lsm5.Hardware.CpFocus.IsBusy
            Sleep 100
        Wend
        
        Sleep 100
        
        Lsm5.CloseAllImageWindows 0
        
        Lsm5.NewScanWindow
        Sleep 200
        
        Lsm5.StartScan
    
        Do While Lsm5.Hardware.CpScancontrol.IsGrabbing
            DoEvents
            Sleep 400
        Loop
        
        
        Sleep 100
        
        setaotf zoomaotfpower
        
        Set nd = Lsm5.DsRecordingActiveDocObject
        nd.SetTitle xspot & " " & yspot
        
        Dim ac
        'Dim xc
        'Dim yc
        
        Lsm5.DsRecording.Sample0X = xspot
        Lsm5.DsRecording.Sample0Y = yspot
        
        DoEvents
        
        
        ifilename = OrfNamesList(WellNumber) & "_zoomf_" & thisday & "_" & CStr(WellNumber) _
                    & "_" & CStr(ScanNumber) & "_" & CStr(i)
        
        
        zname = OrfNamesList(WellNumber) & "_zfocus_" & thisday & "_" & CStr(WellNumber)
        j = nd.SaveToDatabase(fname & "\" & dbname, ifilename)
        
        CurrentImage = fname & "\" & ifilename & ".lsm"
        
        
        process_closeup nd, ac, xc, yc
        
'        Dim xylog As String
'        xylog = fname & "\" & OrfNamesList(WellNumber) & "_" & WellNumber _
'                    & "_" & CStr(ScanNumber) & "_" & CStr(i) & "_xylog.txt"
                    
        Open xylog For Append As #4
        
        Dim tubecurrent As Double
        
        Lsm5.Hardware.CpLasers.Select ("Argon/2")
        tubecurrent = Lsm5.Hardware.CpLasers.tubecurrent
        
        Print #4, OrfNamesList(WellNumber) & " , " & _
                  WellNumber & " , " _
                  ; ScanNumber & " , " & _
                  CStr(i) & " , " & _
                  xc & " , " & yc & " , " & CStr(tubecurrent)
                  
        Close #4
        
        after_name = "cr_ " & ifilename
        j = nd.SaveToDatabase(fname & "\" & dbname, after_name)
        
        'j = nd.SaveToDatabase(fname & "\" & dbname, ifilename)
        'Lsm5.SaveRecording
         'j = nd.SaveToDatabase(fname & "\" & dbname, ifilename)
        'CurrentImage = fname & "\" & ifilename & ".lsm"
        'Debug.Print j & " " & fname & "\" & dbname, "zoomf" & CStr(i)
        
        nd.CloseAllWindows
        
        DoEvents
        If ac <> -1 Then
            Dim nx
            Dim ny
        
            nx = Lsm5.DsRecording.SamplesPerLine
            ny = Lsm5.DsRecording.LinesPerFrame
            
            xf = xspot + (xc - nx / 2#) * xsp
            yf = yspot - (-yc + ny / 2#) * ysp
            zf = xStageZ0
            'zf = Lsm5.DsRecording.Sample0Z
    '
            '' this is all going in micrometer, must be converted to meters in fcs
            
            
            If UserForm1.cbFCS Then
            
                'Lsm5.DsRecording.Sample0X = xf
                'Lsm5.DsRecording.Sample0Y = yf
                
                'setaotf 0.38
                'Sleep 100
                'zmax = zfocusfcs
                'zmax = zfocuspostion
                'If i = 0 Then
                 '   zmax = setupFCS(xf, yf)
                
                'End If
                DoEvents
                Open zlogfile For Append As #2
                Print #2, zmax
                Close #2
                DoEvents
                
                'If i = 1 Then
                '    zmax1 = zmax
                'End If
                Lsm5.Hardware.CpFocus.position = zmax
                
                While Lsm5.Hardware.CpFocus.IsBusy
                    Sleep 100
                Wend
                
                r1 = Lsm5.DsRecording.Sample0X
                r2 = Lsm5.DsRecording.Sample0Y
                
                'setaotf 0.25 'fcsaotfpower
                Sleep 100
                
                takeFCS xf, yf, zmax, i
                gNumberOfCells = gNumberOfCells + 1
                setaotf zoomaotfpower
                'Sleep 100
                
                Lsm5.DsRecording.Sample0X = xf
                Lsm5.DsRecording.Sample0Y = yf
                DoEvents
'                Lsm5.NewScanWindow
'                Lsm5.StartScan
'                Do While Lsm5.Hardware.CpScancontrol.IsGrabbing
'                    DoEvents
'                    Sleep 400
'                Loop
                
              '  DoEvents
'                Set afterfcs = Lsm5.DsRecordingActiveDocObject
                   
                ifilename = "after_zoomf_" & CStr(WellNumber) & "_" & CStr(ScanNumber) & "_" & CStr(i)
        
        '''******notes by CJW 11/12/2007
                ''' the current name of the data base is  ---- fname & "\" & dbname ---
                ''' so the current folder is fname, by changing fname, the directory will change
                ''' fname is defined in the top of module 1
                
'                j = afterfcs.SaveToDatabase(fname & "\" & dbname, ifilename)
'                afterfcs.CloseAllWindows
                Sleep 100
                DoEvents
                
            End If
            
            'j = nd.SaveToDatabase(fname & "\" & dbname, "zoomf" & CStr(i))
            'Debug.Print j & " " & fname & "\" & dbname, "zoomf" & CStr(i)
            
            'nd.EnableReuse 1
            
        Else
            Debug.Print "no regions!!!!!!!!"
        End If
        Debug.Print "************************************"
        
        UserForm1.lblCells.Caption = CStr(gNumberOfCells)
        
        If gNumberOfCells > 9 Then
            Exit For
        End If
         Sleep 500
    Next
    
    Set nd = Nothing
    Set afterfcs = Nothing
    
    Lsm5.DsRecording.zoomx = 1
    Lsm5.DsRecording.zoomy = 1
    Lsm5.DsRecording.Sample0X = 0
    Lsm5.DsRecording.Sample0Y = 0
    Lsm5.DsRecording.LinesPerFrame = 512
    Lsm5.DsRecording.SamplesPerLine = 512
    
End Function
Function prtest()

    MsgBox Lsm5.Tools.GetSettingKey
    Dim doc As DsRecordingDoc
    'Set doc = Lsm5.DsRecordingActiveDocObject
    
    'process_closeup doc
    
End Function


Function process_closeup(doc As DsRecordingDoc, ac, xc, yc)

    Dim dimx As Long
    Dim dimy As Long
    
    'doc = Lsm5.DsRecordingActiveDocObject
    dimx = doc.GetDimensionX
    dimy = doc.GetDimensionY
    
    Lsm5.StopScan
    
    Sleep 200
    
    Dim spl As Long
    Dim bpp As Long
    Dim j As Long
    
    npixels = dimx * dimy
    
    ReDim ximage(0 To dimx - 1)
    ReDim qimage(0 To dimx - 1, dimy - 1)
    
    'For i = 0 To dimx - 1
    
'    For j = 0 To dimx - 1
'        ximage(j) = doc.ScanLine(0, 0, 0, j, spl, bpp)
'    Next
'
'
'    For i = 0 To dimx - 1
'        For j = 0 To dimy - 1
'            qimage(i, j) = ximage(i)(j)
'        Next
'
'    Next

    'doc.CloseAllWindows
    'r = doc.Export(eExportTiff, "c:\temp\test.tif", False, False, 0, 0, True, 0, 0, 0)
    'qimagefilename = "c:\temp\test.tif"
    
    'Set UserForm1.oIDL = New conIDL
    UserForm1.oIDL.sendCloseUp CurrentImage
    
    'Dim xc
    'Dim yc
    'Dim ac
    
    xc = UserForm1.oIDL.get_xCloseUp
    yc = UserForm1.oIDL.get_yCloseUp
    ac = UserForm1.oIDL.getAreas_CloseUp
    
    'UserForm1.oIDL.closeidl
    'np = UBound(xc)
    ll = 5
    

        ix = xc
        iy = yc

        circ1 = doc.VectorOverlay.AddSimpleDrawingElement(eDrawingModeClosedPolyLine, _
                 ix, iy - ll, ix, iy + ll)
        circ1 = doc.VectorOverlay.AddSimpleDrawingElement(eDrawingModeClosedPolyLine, _
                 ix - ll, iy, ix + ll, iy)




    nv = doc.VectorOverlay.GetNumberDrawingElements
    
    For ii = 0 To nv - 1
        doc.VectorOverlay.ElementColor(ii) = RGB(255, 255, 0)
    Next

    Sleep 1000
    
End Function

Sub testpinhole()
    setfcspinhole CInt(56)
End Sub



Function setfcspinhole(diam As Integer)

    
    Dim fcs As AimFcsController
    Dim beampath As AimFcsBeamPathParameters
    
    Set fcs = Lsm5.ExternalDsObject.FcsController
    Set beampath = fcs.BeamPathParameters
    Dim rdiam As Double
    
    rdiam = 0.000001 * diam
    'MsgBox beampath.PinholeDiameter(0)
    beampath.PinholeDiameter(0) = rdiam
    
    Set beampath = Nothing
    Set fcs = Nothing
    
    Sleep 200
End Function

Function setfcslaser(level As Double)

    
    Dim fcs As AimFcsController
    Dim beampath As AimFcsBeamPathParameters
    Dim hinfo As AimFcsHardwareInformation
    
    Set fcs = Lsm5.ExternalDsObject.FcsController
    Set hinfo = fcs.HardwareInformation
    
    nlasers = hinfo.GetNumberAttenuators
    
    Set beampath = fcs.BeamPathParameters
    
'    For i = 0 To nlasers - 1
'        If beampath.AttenuatorOn(i) Then
'            ap = beampath.AttenuatorPower(i)
'
'            'MsgBox i & " " & ap
'            beampath.AttenuatorPower(i) = level
'        End If
'    Next
    
    
    beampath.AttenuatorPower(3) = laser488
    
    beampath.AttenuatorPower(5) = laser561
    
    'rdiam = 0.000001 * diam
    'MsgBox beampath.PinholeDiameter(0)
    'beampath.PinholeDiameter(0) = rdiam
    
    Set beampath = Nothing
    Set fcs = Nothing
    
    Sleep 200
End Function

Function takeFCS(xf, yf, zf, inum)
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
    
    If inum = 0 Then
        setfcslaser 0.014
    End If
    
    setfcspinhole fcspinhole
    
    Set FcsController = Lsm5.ExternalDsObject.FcsController
    Set DocumentList = Lsm5.ExternalDsObject.ConfoCorDisplay
    Set fcspos = FcsController.SamplePositionParameters
    
    Dim power As Double
    Dim Options As AimScanControllerOptions
    
    Dim imagePower As Double
    Dim bp As AimFcsBeamPathParameters
    'Set bp = FcsController.BeamPathParameters

    
    Set fcsOptions = FcsController.Options
    fcsOptions.SaveRawData = True
    
    fcsOptions.RawDataPath = fname & "\" & OrfNamesList(WellNumber) & _
        "_raw_file_" & CStr(WellNumber) & "_" & CStr(ScanNumber) _
        & "_" & CStr(inum) & "__"

    FcsController.AcquisitionParameters.MeasurementTime = 10
    FcsController.AcquisitionParameters.MeasurementRepeat = 1
    
    'MsgBox fcspos.PositionZ(0)
    ' Here comes an Example on how to access other interfaces of the fcs controller
    ' either assign to a varibale
    
    
    Set HardwareControl = FcsController.HardwareControl
    
    ''Do this, suggestion from Zeiss
    ''cjw 07/19/2010
    
    'fcspos.SamplePositionMode = eFcsSamplePositionModeSequential
    Dim index As Long
    index = FcsController.HardwareInformation.StageIndexFromIdentifier(eFcsStageScannerXY)
    
    FcsController.AcquisitionParameters.SequentialIllumination = True
    f1 = FcsController.AcquisitionParameters.SequentialIlluminationPeriod
    ''use the following line to control switching frequency
    FcsController.AcquisitionParameters.SequentialIlluminationPeriod = 0.0001
    ''cjw change this, using index from above
    ''FcsController.Options.StageIndex = eFcsStageScannerXY
     FcsController.Options.StageIndex = index
     fcspos.SamplePositionMode = eFcsDataSamplePositionModeList
     
     ppp = FcsController.Options.StageIndex
     pau = FcsController.Options.SequentialIlluminationPause
     
    'FcsController.Initialize HardwareControl
    'MsgBox FcsController.HardwareControl.MicroscopeMode
    'HardwareControl.MicroscopeMode = eFcsMicroscopeModeFCS
    
    Set Document = DocumentList.NewDocument(True)
    Set DataList = Document.DataList
    
    Dim chkDataList As AimFcsData
    Dim chkDoc As FcsDocument
    Dim chkdataSet As AimFcsDataSet
    
    Set chkDoc = DocumentList.NewDocument(False)
    Set chkDataList = chkDoc.DataList
    
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
    
    'Debug.Print "****" & s1 & " " & stx
    'Debug.Print "****" & s2 & " " & sty
    'Debug.Print "Current Z position 1****" & s3
    
    'Debug.Print fcspos.SamplePositionMode
    fcspos.SamplePositionMode = eFcsDataSamplePositionModeList
    'fcspos.SamplePositionMode = eFcsSamplePositionModeSequential
    fcspos.PositionListSize = 1
    
    fcspos.PositionX(0) = xf * 0.000001
    fcspos.PositionY(0) = yf * 0.000001

    'Debug.Print "x position 0**** " & fcspos.PositionX(0)
    'Debug.Print "y position 0****" & fcspos.PositionY(0)
    'Debug.Print "Current Z position 2 0****" & fcspos.PositionZ(0)
    
    fcspos.PositionZ(0) = s3
    
    HardwareControl.MoveToMeasurementPosition (0)

    Lsm5.Hardware.CpFocus.position = zf
    
    While HardwareControl.IsFocusBusy
        'HardwareControl.GetFocusPosition s3
        'Debug.Print "F Busy " & fcspos.PositionZ(0) & " " & s3
        'Debug.Print "LSM Focus Before " & Lsm5.Hardware.CpFocus.position
        Sleep 20
    Wend
        

    setfcslaser 0.02
'    UserForm1.lblLaser.Caption = CStr(laspow)
    DoEvents
    
    
    Set chkDataList = Nothing
    Set chkdataSet = Nothing
    Set chkDoc = Nothing
    
    FcsController.AcquisitionParameters.MeasurementTime = 10
    FcsController.AcquisitionParameters.MeasurementRepeat = 1
    currentaotf = getAotf
    FcsController.StartMeasurement DataList
    
    Set Progress = FcsController.Progress
    
    While (Not Progress.Ready)
        DoEvents
        'Debug.Print Lsm5.Hardware.CpStages.PositionX
        'Debug.Print Lsm5.Hardware.CpStages.PositionY
    Wend
    
    ' Do something with the data of the first data set
    
    'Debug.Print "x pos 1**** " & fcspos.PositionX(0)
    'Debug.Print "y pos 1**** " & fcspos.PositionY(0)
    
    'Debug.Print "Z Position After 1**** " & fcspos.PositionZ(0)
    'MsgBox CStr(DataList.DataSets) + " data sets acquired"
    DoEvents
    FcsController.StopAcquisitionAndWait
    DoEvents
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
    DataList.Comment = "Aotf: " & currentaotf
    
    Set FileWriter = New AimFcsFileWrite
    FileWriter.Source = DataList
    'FileWriter.FileName = fname & "\" & "fcs_" & CStr(inum) & ".fcs"
    
    fcsname = fname & "\" & OrfNamesList(WellNumber) & "_fcs_" & thisday & "_" & CStr(WellNumber) _
        & "_" & CStr(ScanNumber) _
        & "_" & CStr(inum) & ".fcs"
        
    'fcsname = "c:\temp\temp.fcs"
    FileWriter.FileName = fcsname
    FileWriter.FileWriteType = eFcsFileWriteTypeAll
    'FileWriter.Format = eFcsFileFormatConfoCor3
    FileWriter.Format = eFcsFileFormatConfoCor3WithRawData
    FileWriter.Start ' use start for asynchron writing or "Run" for sysnchron writing
                     ' there is no need to check the progress in the latter case
    
    Set Progress = FileWriter
    While (Not Progress.Ready)
        DoEvents
    Wend
    
    'remove the document again
    Document.CloseWindow
    
    Lsm5.Hardware.CpStages.PositionX = stx
    Lsm5.Hardware.CpStages.PositionY = sty
    
        
    'Set Options = Lsm5.ExternalDsObject.ScanController
    'Options.AotfDriverPower = imagePower
    
    'Set Options = Nothing
    
    FcsController.HardwareControl.MicroscopeMode = 1 'eFcsMicroscopeModeLSM
    Sleep 100
    
    setfcslaser 0.02
End Function

Function fcstest(xf, yf, zf, inum)
    Dim FcsController As AimFcsController
    Dim HardwareControl As AimFcsHardwareControl
    Dim fcspos As AimFcsSamplePositionParameters
    Dim Progress As AimProgress
    Dim DocumentList As ConfoCorDisplay.ConfoCorDisplay
    Dim Document As ConfoCorDisplay.FcsDocument
    Dim DataList As AimFcsData
    Dim DataSet As AimFcsDataSet
    Dim FileWriter As AimFcsFileWrite
    Dim CorrelationArraySize As Long
    Dim CorrelationTime() As Double
    Dim Correlation() As Double
    
    Set FcsController = Lsm5.ExternalDsObject.FcsController
    Set DocumentList = Lsm5.ExternalDsObject.ConfoCorDisplay
    Set fcspos = FcsController.SamplePositionParameters
    
    'MsgBox fcspos.PositionZ(0)
    ' Here comes an Example on how to access other interfaces of the fcs controller
    ' either assign to a varibale
    
    
    Set HardwareControl = FcsController
     
    'FcsController.Initialize HardwareControl
    'MsgBox FcsController.HardwareControl.MicroscopeMode
    'HardwareControl.MicroscopeMode = eFcsMicroscopeModeFCS
    
   
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
    
    Dim s1 As Double
    Dim s2 As Double
    Dim s3 As Double
    
    HardwareControl.GetStagePosition s1, s2
    HardwareControl.GetFocusPosition s3
    FcsController.Options.StageIndex = eFcsStageScannerXY
    FcsController.AcquisitionParameters.MeasurementTime = 0.01
    p1 = fcspos.PositionX(0)
    p2 = fcspos.PositionY(0)
    p3 = fcspos.PositionZ(0)
    
    Debug.Print "****" & s1
    Debug.Print "****" & s2
    Debug.Print "****" & s3
    
    Debug.Print fcspos.SamplePositionMode
    fcspos.SamplePositionMode = eFcsDataSamplePositionModeList
    
    fcspos.PositionListSize = 1

    'fcspos.PositionX(0) = xf * 0.000001
    'fcspos.PositionY(0) = yf * 0.000001

    Dim delx As Double
    delx = 0.000004
    dely = 0.000004
    

    FcsController.StartScanCountRateImage p1, p2, p3 - delx, p1, _
        p2, p3 + delx, 1, 1, 10
    
    'Set Progress = FcsController.Progress
    
    i = 0
    While FcsController.IsAcquisitionRunning
        Sleep 1000
        'Debug.Print s3 & " " & fcspos.PositionZ(0)
        Debug.Print i & "  " & Lsm5.Hardware.CpFocus.position
        i = i + 1
    Wend
    
    'Sleep 2000
        
    FcsController.StopAcquisition
    Dim na As Long
    
    na = FcsController.GetCountRateImageResultArraySize
    MsgBox na
    Dim gdata As Double
    Dim xd(0 To 9) As Double
    
    For i = 0 To 9
        'FcsController.GetCountRateImageData 0, i, 1, gdata
        
        
        xd(i) = gdata
        Debug.Print i & " " & xd(i)
    Next
    
    'FcsController.GetCountRateImageData 0, 0, 10, dd
    Sleep 2000
    Debug.Print "0****" & fcspos.PositionX(0)
    Debug.Print "0****" & fcspos.PositionY(0)
    Debug.Print "0****" & fcspos.PositionZ(0)
    
    Set Document = DocumentList.NewDocument(True)
    Set DataList = Document.DataList
    
    fcspos.PositionZ(0) = s3
    Debug.Print "0****" & fcspos.PositionZ(0)
'    fcspos.PositionX(1) = 15
'    fcspos.PositionY(1) = -1
'    fcspos.PositionX(2) = 20
'    fcspos.PositionY(2) = 3
    
    
    
    
'    FcsController.StartMeasurement DataList
'
'
'    Set Progress = FcsController.Progress
'
'    While (Not Progress.Ready)
'        DoEvents
'    Wend
'
'    Set DataSet = DataList.DataSet(0)
'    Dim ar1 As Double
'    Dim ar2 As Double
'
'    DataSet.GetDataArray eFcsDataTypeCountRate, 1, ar1, ar2
'    ' Do something with the data of the first data set
'
'    Debug.Print "1****" & fcspos.PositionX(0)
'    Debug.Print "1****" & fcspos.PositionY(0)
'    Debug.Print "1****" & fcspos.PositionZ(0)
    'MsgBox CStr(DataList.DataSets) + " data sets acquired"
    
    Sleep 1000
    
 
    
    
End Function
Sub setaotf(power As Double)

    Dim Options As AimScanControllerOptions
    
    Set Options = Lsm5.ExternalDsObject.ScanController
    
    Options.AotfDriverPower = power
    
    Sleep 200
    Set Options = Nothing
    
End Sub


Function stepaotf(step As Integer)

    Dim aotf As Double
    Dim power As Double
    
    aotf = getAotf
    
    If step * step <> 1 Then
        setaotf 0.13
        Exit Function
    End If
    
    If step = -1 Then
        Select Case aotf
            Case 0.13
                power = 0.13
            Case 0.25
                power = 0.13
            Case 0.38
                power = 0.25
            Case 1#
                power = 0.38
        End Select
    Else
        Select Case aotf
            Case 0.13
                power = 0.25
            Case 0.25
                power = 0.38
            Case 0.38
                power = 1#
            Case 1#
                power = 1#
        End Select
    End If
    
    setaotf power
    
End Function
Function getAotf()
    
    Dim power As Double
    Dim Options As AimScanControllerOptions
    
    Set Options = Lsm5.ExternalDsObject.ScanController
    
    power = Options.AotfDriverPower
    Sleep 200
    
    'Sleep 100
    Set Options = Nothing
    getAotf = power
End Function
Sub reset_frame()
    Dim ds As DsRecording
    Set ds = Lsm5.DsRecording
    ds.SpecialScanMode = "NoSpecialMode"
    ds.ScanMode = "Plane"
    ds.SamplesPerLine = 512
    ds.LinesPerFrame = 512
    ds.FramesPerStack = 1
    ds.StacksPerRecord = 1
    Set ds = Nothing
End Sub

Function runtest()
    

    fcstest 0, 0, 0, 1
End Function

Sub maxtest()
    
    Dim d As DsRecordingDoc
    Set d = Lsm5.DsRecordingActiveDocObject
    
    m = getLineMax(d)
    
    circ1 = d.VectorOverlay.AddSimpleDrawingElement(eDrawingModeClosedPolyLine, _
                 0, m, 512, m)
    
    
End Sub
