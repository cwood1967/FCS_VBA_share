VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "conIDL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Dim idl As COM_IDL_connectLib.COM_IDL_connect
'Dim WithEvents idl As COM_IDL_connectLib.COM_IDL_connect


Public a
Private areas()
Private xCoords()
Private yCoords()
Private xclose
Private yclose
Private aclose

Public Event idlevent(s As String)


Public Function get_xCoords()
    get_xCoords = xCoords
End Function

Public Function get_yCoords()
    get_yCoords = yCoords
End Function

Public Function getAreas()
    getAreas = areas
End Function

Public Function get_xCloseUp()
    get_xCloseUp = xclose
End Function

Public Function get_yCloseUp()
    get_yCloseUp = yclose
End Function

Public Function getAreas_CloseUp()
    getAreas_CloseUp = aclose
End Function

Public Function getmes()
    idl.ExecuteString "print, (*(*(myptr.q)))"
    getmes = idl.GetProcessName
End Function

Public Function sendImage(ByVal params, ByVal imagefile)
    
    Sleep 100
    
    Open "c:\temp\currentimage.dat" For Output As #1
    Print #1, CurrentImage
    Print #1, OverviewChannel - 1
    Close #1
    
    Sleep 100
    idlpath = "C:\Program Files\ITT\IDL\IDL81\bin\bin.x86\idlrt"
'    idlpath = "C:\Program Files\ITT\IDL706\bin\bin.x86\idlrt"
'    idlpath = "C:\Program Files\ITT\IDL64\bin\bin.x86\idlrt"
    idlsav = "c:\temp\idlfindyeast.sav"
    RunCmd idlpath & " -novm " & idlsav, vbNormalFocus
     readDataFile
     sendImage = areas
    Exit Function
'    'Set idl = Nothing
'
'    Set idl = New COM_IDL_connectLib.COM_IDL_connect
'
'    idl.CreateObject 0, 0, 0
'    'MsgBox idl.GetProcessName
'    'idl.ExecuteString "heap_gc"
'     idl.ExecuteString "pref_set, 'idl_gr_win_renderer', 1"
'
'    ''does the following line cause the code to crash?
'
'    'idl.ExecuteString ".reset"
'    'idl.SetIDLVariable "ximage", image
'    idl.SetIDLVariable "ximagefile", imagefile
'    idl.ExecuteString "ximage = read_tiff(ximagefile)"
'
'    idl.SetIDLVariable "params", params
'    'idl.ExecuteString "print, max(ximage)"
'
'    idl.ExecuteString "cd , 'U:\cjw\Worm_tracking'"
'
'    idl.ExecuteString ".compile 'U:\cjw\Worm_tracking\cjw_morph_yeast.pro'"
'
'    idl.ExecuteString ".compile 'U:\cjw\Worm_tracking\cjw_overview_info.pro'"
'
'    idl.ExecuteString "cjw_morph_yeast, ximage, timage = timage, mimage = mimage"
'    'idl.ExecuteString "w_tracking, mimage, params, myptr, res"
'    idl.ExecuteString "r= cjw_overview_info(mimage, xarea = xarea, xc = xc, yc =yc)"
'    Debug.Print "i am outta here"
'
'    'va = idl.GetIDLVariable("(*(*myptr.coords))")
'
'    readDataFile
''    xCoords = idl.GetIDLVariable("xc")
''    yCoords = idl.GetIDLVariable("yc")
''    areas = idl.GetIDLVariable("xarea")
'    Dim nh
'
'
'
'    'idl.ExecuteString "device, decomposed=0"
'
'    a = 14
'    sendImage = areas
End Function
Public Sub readCloseUpFile()

    
    Open "c:\temp\currentimage.dat" For Output As #1
    Print #1, CurrentImage
    Close #1
    
    Dim datafile As String
    datafile = "c:\temp\celldata.dat"
    
    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim Stream As TextStream
    Set Stream = fso.OpenTextFile(datafile, ForReading)
    
    Dim inum As Integer
    inum = CLng(Stream.ReadLine)
    
    If inum <> 1 Then
        xclose = -1
        yclose = -1
        aclose = -1
        Exit Sub
    End If
    
    Dim data_array()
    
    sline = Stream.ReadLine
    Stream.Close
    aclose = CDbl(Mid(sline, 1, 13))
    If aclose = -1 Then
        xclose = -1
        yclose = -1
    Else
        xclose = CDbl(Mid(sline, 14, 13))
        yclose = CDbl(Mid(sline, 27, 13))
    End If
    
    Set fso = Nothing

End Sub

Public Sub readDataFile()
    
    Dim datafile As String
    datafile = "c:\temp\xdata.dat"
    
    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim inum As Integer
    Dim data_array()
    Dim Stream As TextStream
    If fso.FileExists(datafile) Then
        Set Stream = fso.OpenTextFile(datafile, ForReading)
    Else
        inum = 1
        ReDim data_array(inum)
        ReDim xCoords(0 To inum - 1)
        ReDim yCoords(0 To inum - 1)
        ReDim areas(0 To inum - 1)
        areas(0) = -1
        xCoords(0) = -1
        yCoords(0) = -1
        Exit Sub
    End If
    
   
    inum = CLng(Stream.ReadLine)
    
    If inum = 0 Then inum = 1
    If inum = -1 Then inum = 1
    ReDim data_array(inum)
    ReDim xCoords(0 To inum - 1)
    ReDim yCoords(0 To inum - 1)
    ReDim areas(0 To inum - 1)
    
    Dim iline As Integer
    iline = 0
    
    While Not Stream.AtEndOfStream
        data_array(iline) = Stream.ReadLine
        iline = iline + 1
                
    Wend
    
    Stream.Close
    
    For iline = 0 To inum - 1
        sline = data_array(iline)
        areas(iline) = CDbl(Mid(sline, 1, 13))
        If areas(iline) = -1 Then Exit For
        xCoords(iline) = CDbl(Mid(sline, 14, 13))
        yCoords(iline) = CDbl(Mid(sline, 27, 13))
    Next
    
    Set fso = Nothing
End Sub






Public Function sendCloseUp(ByVal imagefilename)
   '' use this function to call an idl function that will send the cliose up image to idl
   '' and find the center of mass of the cells in the image
   '' ......
   '' the idl  - program cjw_overview_info.pro -
        
        
    Open "c:\temp\currentimage.dat" For Output As #1
    Print #1, CurrentImage
    Print #1, ZoomChannel - 1
    Close #1
    
    Sleep 100
        idlpath = "C:\Program Files\ITT\IDL\IDL81\bin\bin.x86\idlrt"
        'idlpath = "C:\Program Files\ITT\IDL706\bin\bin.x86\idlrt"
        'idlpath = "C:\Program Files\ITT\IDL64\bin\bin.x86\idlrt"
        idlsav = "c:\temp\idlCloseUp.sav"
        RunCmd idlpath & " " & idlsav, vbNormalFocus
    
        readCloseUpFile
        
        sendCloseUp = areas
    Exit Function
    Sleep 100
    Set idl = Nothing
    
    'Set idl = New COM_IDL_connectLib.COM_IDL_connect
    
    'idl.CreateObject 0, 0, 0
   
    'idl.ExecuteString "heap_gc"
    
    ''rsi support claims the following line may be to blame for the crashing problems
    '' we will see if this is the case
    
    'idl.ExecuteString ".reset"   ''does this cause the crash?
    
'    idl.SetIDLVariable "ximagefilename", imagefilename
'
'    idl.ExecuteString "ximage = read_tiff(ximagefilename)"
'
'    'idl.ExecuteString "tvscl, ximage"
'    idl.ExecuteString "!path = !path + '; u:\cjw\Worm_tracking'"
'    idl.ExecuteString ".compile 'u:\cjw\Worm_tracking\cjw_get_regions.pro'"
'    idl.ExecuteString "r= cjw_get_regions(ximage,8, area=area, xc = xc, yc =yc)"
'    idl.ExecuteString "help, xc"
'
'    xclose = idl.GetIDLVariable("xc")
'    yclose = idl.GetIDLVariable("yc")
'    aclose = idl.GetIDLVariable("area")

   
   
   
End Function
Public Sub closeidl()
    On Error Resume Next
    'idl.DestroyObject
End Sub


Private Sub Class_Initialize()

End Sub

Public Sub idl_OnIDLOutput(ByVal Output As String)
    If InStr(1, Output, "compiled", vbTextCompare) = 0 Then
        'Debug.Print Output
        RaiseEvent idlevent(Output)
    End If
'
'    If InStr(1, Output, "raise_event_100", vbTextCompare) > 0 Then
'        Debug.Print "Yes"
'        MsgBox "hey I got it"
'        idl.ExecuteString "print, (*(*myptr.q))"
'
'    End If
'    a = a + 1
'
'
    'Debug.Print Output
'    RaiseEvent idlevent(Output)
End Sub

Private Sub idl_OnMouseUp(ByVal Button As Long, ByVal KeyState As Long, ByVal x As Long, ByVal y As Long, ForwardToIdl As Long)
    idl.ExecuteString "print, (*(*(myptr.q)))"
End Sub
