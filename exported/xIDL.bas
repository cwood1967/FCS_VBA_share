Attribute VB_Name = "xIDL"
Function getSpacing() As Variant
    Dim d As DsRecordingDoc
    
    'Dim xspace
    'Dim yspace
    Set d = Lsm5.DsRecordingActiveDocObject
    
    '' xspace and yspace are returned in meters
    
    xspace = d.Recording.SampleSpacing
    yspace = d.Recording.LineSpacing
   
    'd.Export eExportTiff, "c:\test.tif", False, False, 0, 0, True, 0, 0, 0
    
    Set d = Nothing
    getSpacing = Array(xspace, yspace)
End Function

Function getZoom() As Variant

    Dim d As DsRecordingDoc
    
    Dim zoomx
    Dim zoomy
    
    Set d = Lsm5.DsRecordingActiveDocObject
    zoomx = d.Recording.zoomx
    zoomy = d.Recording.zoomy

    getZoom = Array(zoomx, zoomy)
    
End Function
Function getImage() As Variant
    
    Dim lt As Lsm5Tools
    
    Dim d As DsRecordingDoc
    
   
    Dim ximage()
    Dim qimage()
    
    If Lsm5.DsRecordingActiveDocObject Is Nothing Then
        MsgBox "Scan a new image"
        getImage = -1
        Exit Function
    End If
        
        
'    Set d = Lsm5.DsRecordingActiveDocObject

'    d.NeverAgainScanToTheImage
    'rj = d.SaveToDatabase(fname & "\" & dbname, "overviewscan")
    
    Dim suc As Integer
'    Dim tr As DsTrack
    
'    Set tr = d.Recording.TrackObjectByIndex(0, suc)
    
'    If Not tr Is Nothing Then
'        Dim bs1 As DsBeamSplitter
'        Set bs1 = tr.BeamSplitterObjectByIndex(0, suc)
'
'        Lsm5.Hardware.CpServos.Select (4)
'    '    MsgBox Lsm5.Hardware.CpServos.Summary
'    '    MsgBox bs1.Name
'    '    MsgBox bs1.Filter & " " & bs1.FilterSet
'        'tr.SaveConfigurationSetting "FromVBA"
'
'    End If
    
'    Dim dimx As Long
'    Dim dimy As Long
    
'    dimx = d.GetDimensionX
'    dimy = d.GetDimensionY
    
    Dim spl As Long
    Dim bpp As Long
    Dim j As Long
    
'    npixels = dimx * dimy
    
    'ReDim ximage(0 To dimx - 1)
    'ReDim qimage(0 To dimx - 1, dimy - 1)
    
    'For i = 0 To dimx - 1
    
'    For j = 0 To dimx - 1
'        ximage(j) = d.ScanLine(0, 0, 0, j, spl, bpp)
'    Next
'
'
'    For i = 0 To dimx - 1
'        For j = 0 To dimy - 1
'            qimage(i, j) = ximage(i)(j)
'        Next
'    Next
    
    
    'Dim oc As conIDL
    'Set oc = New conIDL
    
    Dim loc
    Dim params
    ''cjw
    '' i am not doing anything with this, so I will comment it out
    ''09/19/2007
    'params = getLSMparams(d)
    
    ''CJW
    ''09/19/2007
    ''the image should have already been saved, but not closed
    ''so don't do this
    ''rj = d.SaveToDatabase(fname & "\" & dbname, "overviewscan")
    
    
    'Use true for a planar image,
    'and 0,0,0 for channel 2
    ' and 1,1,1 for transmitted?
    
    
    '' 09/19/2007
    ''CJW
    '' not going to export this any longer because the image is already being saved
    '' somewhere else
    '' under a variable called CurrentImage defined in rootInfo
    ''
    ''b = d.Export(eExportTiff, "c:\temp\test000.tif", False, False, 0, 0, True, 0, 0, 0)
    
    
    
    'd.CloseAllWindows
    Lsm5.CloseAllDialogs
    'Lsm5.CloseAllImageWindows 0
    'Lsm5.CloseAllDatabaseWindows
    'Set d = Nothing
    
    'loc = UserForm1.oIDL.sendImage(params, qimage)
    ''loc = UserForm1.oIDL.sendImage(params, "c:\temp\test.tif")
    loc = UserForm1.oIDL.sendImage(params, CurrentImage)
    
    Dim areas
    Dim xc
    Dim yc
    
    areas = UserForm1.oIDL.getAreas
    xc = UserForm1.oIDL.get_xCoords
    yc = UserForm1.oIDL.get_yCoords
    
    loc = areas
    'Exit Function
    Dim n
    'n = UBound(loc, 1)
    n = UBound(xc, 1)
    'Debug.Print n
    Dim ix
    Dim iy
    

'    d.VectorOverlay.RemoveAllDrawingElements
'
'    If n > 6 Then
'        nlimit = 6
'    Else
'        nlimit = n
'    End If
'
'    nlimit = n
'
'
'    For i = 0 To nlimit
'
'        ll = Sqr(areas(i))
'        ix = xc(i)
'        iy = yc(i)
'        'Debug.Print ix & " " & iy
'        'circ = d.VectorOverlay.AddSimpleDrawingElement(eDrawingModeCircle, _
'        '    ix - ll, iy - ll, ix + ll, iy + ll)
'        circ1 = d.VectorOverlay.AddSimpleDrawingElement(eDrawingModeClosedPolyLine, _
'                 ix - ll, iy - ll, ix + ll, iy - ll)
'        circ2 = d.VectorOverlay.AddSimpleDrawingElement(eDrawingModeClosedPolyLine, _
'                 ix + ll, iy - ll, ix + ll, iy + ll)
'        circ3 = d.VectorOverlay.AddSimpleDrawingElement(eDrawingModeClosedPolyLine, _
'                 ix + ll, iy + ll, ix - ll, iy + ll)
'        circ4 = d.VectorOverlay.AddSimpleDrawingElement(eDrawingModeClosedPolyLine, _
'                 ix - ll, iy + ll, ix - ll, iy - ll)
'        'd.VectorOverlay.ElementLineWidth(i) = 2
'        'd.+VectorOverlay.ElementColor(i) = RGB(250, 200, 14)
'        'd.VectorOverlay.ElementColor(i) = 2 ^ 23 + 2 ^ 22 + 2 ^ 21 + 2 ^ 20 + 2 ^ 19 + 2 ^ 18 + 2 ^ 17 + 2 ^ 7
'    Next
'
'    nv = d.VectorOverlay.GetNumberDrawingElements
'    For ii = 0 To nv - 1
'        d.VectorOverlay.ElementColor(ii) = RGB(255, 0, 0)
'    Next
    
    'd.CloseAllWindows
    'rj = d.SaveToDatabase(fname & "\" & dbname, "overviewscan")
    'Set d = Nothing
    getImage = loc
    
   ' Next
    'MsgBox ximage(200)(100)
End Function

Function getLSMparams(d As DsRecordingDoc)

    Dim rec As DsRecording
    Dim params(0 To 50) As Variant
    For si = 0 To 50
        params(si) = "blank"
    Next
    
    Set rec = d.Recording
    
    params(0) = CStr(rec.FrameWidth)
    params(1) = CStr(rec.LinesPerFrame)
    params(2) = CStr(rec.SamplesPerLine)
    params(3) = CStr(rec.Sample0X)
    params(4) = CStr(rec.Sample0Y)
    params(5) = CStr(rec.LineSpacing)
    params(6) = CStr(rec.SampleSpacing)
    params(7) = CStr(rec.zoomx)
    params(8) = CStr(rec.zoomy)
    params(9) = CStr(rec.name)
    
    nl = rec.LaserCount
    params(10) = CStr(nl)
    
    Dim suc As Integer
    Dim Laser As DsLaser
    
    indexnum = 11
    '' don't allow more than 4 lasers
    '' lasers fill numbers 11,12, 13,14, 15,16, 17,18 -- leave 19 for extra
    If nl > 4 Then nl = 4
    
    For i = 0 To nl - 1
        Set Laser = rec.LaserObjectByIndex(i, suc)
        params(indexnum) = CStr(Laser.name)
        indexnum = indexnum + 1
        params(indexnum) = CStr(Laser.power)
        indexnum = indexnum + 1
    Next
    
    params(19) = CStr(0)
    getLSMparams = params
End Function
