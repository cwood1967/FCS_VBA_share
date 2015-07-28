VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12615
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents oIDL As conIDL
Attribute oIDL.VB_VarHelpID = -1
Public basedir
Public waterset As Integer

Private Sub Button1_Click()

End Sub

Private Sub cblasersoff_Click()

End Sub

Private Sub cmbOVChannel_Change()
    OverviewChannel = CInt(cmbOVChannel.Value)
End Sub

Private Sub cmbZoomChannel_Change()
    ZoomChannel = CInt(cmbZoomChannel.Value)
End Sub

Private Sub cmdfcs_Click()
    Dim x
    Dim y
    Dim z
    
    fname = Trim(TextBox2.Text)
    x = CDbl(txtfcsx.Text)
    y = CDbl(txtfcsy.Text)
    z = Lsm5.Hardware.CpFocus.position
    takeFCS x, y, z, 0
    
End Sub

Private Sub cmdFolder_Click()
    'getImage
    basedir = BrowseFolder
    TextBox2.Text = basedir
End Sub

Private Sub cmdFind_Click()
    delsav
    
    ztime = Trim(txtztime.Text)
    zthickness = Trim(txtzthickness.Text)
    zpoints = Trim(txtzpoints.Text)
    
    basedir = Trim(TextBox2.Text)
    If basedir = "" Then
        basedir = BrowseFolder
        TextBox2.Text = basedir
    End If
    newdb
    Set oIDL = New conIDL
    WellNumber = 0
    ScanNumber = 0
    'justimage
    GoProcess
    MsgBox "done done done"
End Sub

Private Sub cmdmarkOrigin_Click()
    stageWellOriginX = Lsm5.Hardware.CpStages.PositionX
    stageWellOriginY = Lsm5.Hardware.CpStages.PositionY
    marked = 1
End Sub

Private Sub cmdOrfBrowse_Click()
    Dim name As String
    
    g = SaveFileNameBox(name, "Text Files", "*.txt", False)
    txtorffile.Text = name
    
    cmdOrfBrowse.ControlTipText = name
    
    
    
End Sub

Public Function SaveFileNameBox(name As String, FileType As String, FileExtension As String, bSave As Boolean) As Boolean
    On Error GoTo cend

    Dim filebox As OPENFILENAME  ' open file dialog structure
    Dim result As Long           ' result of opening the dialog
    
    With filebox
        .lStructSize = Len(filebox)
            .hwndOwner = 0 'Me.hWnd
        .hInstance = 0
        If FileType = "" Then
            .lpstrFilter = "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar
        Else
            .lpstrFilter = FileType + " (" + FileExtension + ")" + vbNullChar + FileExtension + vbNullChar & vbNullChar
        End If
                
        .nMaxCustomFilter = 0
        .nFilterIndex = 1
        .lpstrFile = name + Space(256) & vbNullChar
        .nMaxFile = Len(.lpstrFile)
        .lpstrFileTitle = Space(256) & vbNullChar
        .nMaxFileTitle = Len(.lpstrFileTitle)
        
        If bSave = False Then
            If CurrentDir <> "" Then
                .lpstrInitialDir = CurrentDir
            End If
        End If
        
        
        .lpstrDefExt = FileExtension

        
        If bSave Then
            .lpstrTitle = "Save" & vbNullChar
        Else
            .lpstrTitle = "Open" & vbNullChar
        End If
        
        If bSave Then
            .flags = OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY
        Else
            .flags = 0
        End If
        .nFileOffset = 0
        .nFileExtension = 0
        .lCustData = 0
        .lpfnHook = 0
    End With
    
    If bSave Then
        result = GetSaveFileName(filebox)
    Else
        result = GetOpenFileName(filebox)
    End If
    
    CurrentDir = VBA.Left(filebox.lpstrFile, InStrRev(filebox.lpstrFile, "\"))
    
    SaveFileNameBox = result <> 0
    
    If result <> 0 Then
        name = VBA.Left(filebox.lpstrFile, InStr(filebox.lpstrFile, vbNullChar) - 1)
    End If
cend:
End Function

Private Sub cmdScanNew_Click()
    ScanNewOverview
End Sub

Private Sub ImageList1_Enter()

End Sub



Private Sub cmdTestPlate_Click()
    'MsgBox OverviewChannel & " " & ZoomChannel
    
    If waterset = 0 Then
        MsgBox "please mark the water position", vbInformation
        Exit Sub
    End If
    
    killall = 0
    delsav
    xWellRadius = 0.5 * txtwelldiam.Text
    xWellHeight = 2# * 0.707 * xWellRadius
    xWellWidth = xWellHeight
    nWellSpacingX = txtwellspacing
    nWellsSpacingY = nWellSpacingX
    xWellCountX = txtnumwellsx
    xWellCountY = txtnumwellsy
    
    zoomaotfpower = txtzoomaotf.Text
    fcsaotfpower = txtfcsaotf.Text
    pmtaotfpower = txtpmtaotf.Text
    
    laser488 = CDbl(txt488.Text) * 0.01
    laser561 = CDbl(txt561.Text) * 0.01
    imagepinhole = CInt(tbImagePinhole.Text)
    fcspinhole = CInt(tbfcspinhole.Text)
    
    gScanX = txtscanx.Text
    gScanY = txtscany.Text
    

    ztime = Trim(txtztime.Text)
    zthickness = Trim(txtzthickness.Text)
    zpoints = Trim(txtzpoints.Text)
    
    Dim ofiles
    
    If txtorffile <> "" Then
        OrfFilename = txtorffile.Text
        ofiles = readOrfNames
    Else
    
        Dim tmp(0 To 47)
        For i = 0 To 47
            tmp(i) = "Y"
        Next
        ofiles = tmp
    End If
     
    OrfNamesList = ofiles
    
    
    basedir = Trim(TextBox2.Text)
    If basedir = "" Then
        basedir = BrowseFolder
        TextBox2.Text = basedir
    End If
    
    newdb
    
    xylog = fname & "\" & "FCS_XY_Positions.txt"

    Open xylog For Output As #4
    Print #4, "## Orf, Well, Overview Scan, Image #########"
    Close #4

    Set oIDL = New conIDL
    WellNumber = 0
    ScanNumber = 0
    test_plate_scan
    
    Lsm5.Hardware.CpFocus.MoveToLoadPosition
    
    If cblasersoff Then
        lasersOff
    End If
    
    MsgBox "done done done"
End Sub

Private Sub delsav()
    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists("c:\temp\thres.sav") Then
        fso.DeleteFile "c:\temp\thres.sav"
    End If
    
    If fso.FileExists("c:\temp\xdata.dat") Then
        fso.DeleteFile "c:\temp\xdata.dat"
    End If
    
End Sub



Private Sub CommandButton4_Click()
    Dim res As Integer
    Dim ctext As String
    
    ctext = "fconfig01"
    ctext = Trim(cbConfigList.Text)
    'res = SaveFluorConfig(Trim(TextBox1.Text))
    
    res = SaveFluorConfig(ctext)
    
    If res = 1 Then
        MsgBox Trim(ctext) & " was saved as a configuration"
    Else
        MsgBox Trim(ctext) & "failed to saved as a configuration"
    End If
End Sub

Private Sub CommandButton5_Click()
    Set oIDL = New conIDL
    
      fname = TextBox2.Text
    justimage
End Sub


Private Sub cmdTestFocus_Click()
    'fname = "fcstest"
    fname = TextBox2.Text
    ztime = CDbl(txtztime.Text)
    zthickness = CDbl(txtzthickness.Text)
    zpoints = CLng(txtzpoints.Text)
    
    Dim x()
    Dim y()
    Dim z
    
    
    num = 4
    ReDim x(0 To num - 1)
    ReDim y(0 To num - 1)
    ReDim z(0 To num - 1)
    '
    'x1 = 35.22
    'y1 = 7.93
    
    z = setupFCS(CDbl(txtfcsx.Text), CDbl(txtfcsy.Text))
    'setupFCS x1, y1
    takeFCS CVar(txtfcsx.Text), CVar(txtfcsy.Text), z, 2
    Exit Sub
'    x(0) = 41.32
'    y(0) = 37.79
'
'    x(1) = 9.68
'    y(1) = 11.87
'
'    x(2) = 48.79
'    y(2) = -49.66
'
'    x(3) = 21.11
'    y(3) = -89.65
'
''    x(4) = 0
''    y(4) = 0
'    Dim i As Integer
'
'    For i = 0 To num - 1
'
'        z(i) = setupFCS(x(i), y(i))
'        takeFCS x(i), y(i), z(i), i
'
'    Next
    'x1 = -40.07
    'y1 = 75.73
    'setupFCS x1, y1
    
    '' maybe use a fcs at (0,0) to reset the stage
    
    'takeFCS txtfcsx.Text, txtfcsy.Text, 0, 12
End Sub


Private Sub CommandButton6_Click()
    Dim g As conIDL
    Set g = New conIDL
    'g.readDataFile
    
    
    j = g.sendCloseUp("c:\temp\test.tif")

End Sub

Private Sub CommandButton7_Click()
    waterx = Lsm5.Hardware.CpStages.PositionX
    watery = Lsm5.Hardware.CpStages.PositionY
    
    Label19.Caption = waterx
    Label20.Caption = watery
    
    waterset = 1
End Sub

Private Sub CommandButton10_Click()

    Open "c:\temp\watermark.dat" For Input As #1
    Input #1, waterx
    Input #1, watery
    
    Close #1
    Label19.Caption = waterx
    Label20.Caption = watery
    
    waterset = 1
    
End Sub

Private Sub CommandButton8_Click()
    killall = 1
    MsgBox "hello"
    End
End Sub

Private Sub CommandButton9_Click()
    AddWater
End Sub

Private Sub Label11_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub oIDL_idlevent(s As String)
   'MsgBox s
    Debug.Print s
End Sub

Private Sub txt_Change()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub txtfcsy_Change()

End Sub

Private Sub UserForm_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)

End Sub

Private Sub UserForm_Initialize()

    Dim f1
    Dim s1
    
    f1 = Lsm5.Tools.RegExistKey("UI\Settings\zoomscan")
    s1 = Lsm5.Tools.RegExistKey("UI\Settings\overviewscan")
    
    If Not f1 Then
        MsgBox "Create a configuration called zoomscan for scanning detected cells"
        Unload Me
    End If
    
    If Not s1 Then
        MsgBox "Create a configuration called overviewscan for the overview scan"
        Unload Me
    End If
    
    marked = 0
    
    txtzthickness.Value = 8
    txtzpoints.Value = 10
    txtztime = 0.1
    
    txtwellspacing.Value = 9
    txtwelldiam.Value = 6
'    cbConfigList.AddItem "overviewscan"
'    cbConfigList.AddItem "zoomscan"
'    cbConfigList.Text = "zoomscan"
    
    txtpmtaotf.Text = pmtaotfpower
    txtzoomaotf.Text = zoomaotfpower
    txtfcsaotf.Text = fcsaotfpower
    
    d1 = Now
    thismonth = Month(d1)
    thisday = Day(d1)
    thisyear = Year(d1)
    TextBox2.Text = "d:\hc-fcs\" & Format(Now, "mmm_dd_yyyy")
    thisday = Format(Now, "mmm_dd_yyyy")
    'TextBox2.Text = "e:\chris\sept11"
    
    tbfcspinhole.Text = "70"
    tbImagePinhole.Text = "130"
    
    getconfig "overviewscan"
    nc = Lsm5.DsRecording.NumberOfChannels
    
    For i = 1 To nc
        cmbOVChannel.AddItem i
    Next
    
    cmbOVChannel.Value = 1
    
    getconfig "zoomscan"
    nc = Lsm5.DsRecording.NumberOfChannels
    
    For i = 1 To nc
        cmbZoomChannel.AddItem i
    Next
    
    OverviewChannel = 1
    ZoomChannel = 1
    cmbZoomChannel.Value = 1
    waterset = 0
End Sub

Private Sub UserForm_Terminate()

    On Error Resume Next
    If Not oIDL Is Nothing Then
        oIDL.closeidl
        Set oIDL = Nothing
    End If
    
End Sub
