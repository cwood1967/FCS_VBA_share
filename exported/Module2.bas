Attribute VB_Name = "Module2"
Sub getconfig(configname As String)
    
    Dim doc As DsRecordingDoc
    Dim rec As DsRecording
    Dim tr As DsTrack
    
    Dim suc As Integer
    
    'Set Doc = Lsm5.MakeNewImageDocument(512, 512, 1, 1, 1, 1, 1)
    Set rec = Lsm5.DsRecording
    
    c = rec.TrackCount
    
    For i = 0 To c - 1
    
        rec.TrackRemove i
        
    Next
    
    rec.TrackAddNew configname
    Set tr = rec.TrackObjectByName(configname, suc)
    tr.DataChannelRemove (0)
    b = tr.LoadConfigurationSetting(configname)
    'b = tr.LoadConfigurationSetting("how is this")
    'Set tr = rec.TrackObjectByName("cjw488", suc)
    
    'MsgBox c & " " & b & " " & tr.name
End Sub

Function SaveFluorConfig(cname As String) As Integer
    Dim b As Boolean
    
    Dim r As DsRecording
    Dim t As DsTrack
    Dim suc As Integer
    Dim xsuc As Integer
    
    
    If Lsm5.DsRecordingActiveDocObject Is Nothing Then
        MsgBox "You must have an open image"
        SaveFluorConfig = 0
        Exit Function
    End If
    
    Set r = Lsm5.DsRecordingActiveDocObject.Recording
    'Set r = Lsm5.DsRecording
    'Count = Lsm5.DsRecordingActiveDocObject.Recording.TrackCount
    count = r.TrackCount
    
    '' active track is called 'Track'
    
    'Set t = r.TrackObjectByName("Track", suc)
    Set t = r.TrackObjectByIndex(0, suc)
    
    
    't.LoadSaveConfiguration
    Debug.Print , t.name
    
    suc = t.SaveConfigurationSetting(cname)
     
    'suc = Lsm5.Tools.SaveConfigurationSetting(t, cname)
    
    Debug.Print count
    
    For i = 0 To count - 1
        Set t = r.TrackObjectByIndex(i, xsuc)
        Debug.Print t.name
    Next
    
    If suc Then
        SaveFluorConfig = 1
    Else
        SaveFluorConfig = 0
    End If
            
        
End Function

Sub texttracks()

    Dim r As DsRecording
    Dim t As DsTrack
    
    Set r = Lsm5.DsRecording
    
    c = r.TrackCount
    
    Dim i As Integer
    Dim suc As Integer
    Debug.Print '*************************'
    Debug.Print c
    For i = 0 To c - 1
        Set t = r.TrackObjectByIndex(i, suc)
        Debug.Print t.name
        Set t = Nothing
        'r.TrackRemove 0
    Next
    
'    For i = 0 To c - 1
'        r.TrackRemove 0
'        Debug.Print r.TrackCount
'
'    Next
    
    'r.TrackAddNew "newtrack" & CStr(i)
    'Lsm5.StartScan
    'Set t = r.TrackObjectByName("cjw488", suc)
    
    't.LoadConfigurationSetting "j1234"
    f = "UI\Settings\j1234"
    MsgBox Lsm5.Tools.RegStringValue(f, "TrackName")
    
    
End Sub
