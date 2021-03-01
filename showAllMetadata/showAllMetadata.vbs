'
' MediaMonkey Script
'
' NAME: showAllMetadata
'
' AUTHOR: C. Seeling
' DATE : 2020-05-22
'
' ENTRY Point is showPanel
'
' NECESSARY: exiftool.exe has to be placed in same folder!

' Global Variables
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

RootPath=Replace(Script.ScriptPath,"showAllMetadata.vbs","")
ExePath=RootPath&"exiftool.exe"
ArgfilePath=RootPath&"exiftool_Argfile"
InfoFile = RootPath&"exiftoolinfo.txt"

' Register Events
Script.RegisterEvent SDB, "OnChangedSelection", "showAllMetadata"

Call cleanUp()

Sub showAllMetadata
    If SDB.Objects("AllTagsPanelObj").Common.Visible Then
        allMD = getAllMetadata
        ' CreateLabel.Caption = allMD
        writeDoc formatHTML(allMD)
    End If
End Sub

Function getAllMetadata()
    If SDB.SelectedSongList.Count>0 Then
        AudioPath = SDB.SelectedSongList.Item(0).path
        cmd=ExePath&" -api filter='s/\n/\<br\>/g' -charset filename=utf8 -@ exiftool_Argfile -textOut! """&RootPath&"exiftoolinfo%c.txt"""
        deleteAFile(InfoFile)

        ArgfileContent=AudioPath&chr(10)&"-sort"&chr(10)&"-tab"&chr(10)&"-unknown"&chr(10)&"-quiet"
        ' some arguments
        ' -k
        ' -api filter='s/\n/\\n/g'  ''' output newlines as "\n"
        ' -charset ID3=Latin
        ' -charset UTF8
        ' -quiet
        ' -sort
        ' -tab
        ' -unknown
        writeFile ArgfileContent, ArgfilePath
        CreateObject("WScript.Shell").run cmd, 0, True 'RUNS HIDDEN and waits for execution to finish

     ''' Different Alternatives to run external program
        ' command="example.exe -arg """&path&""""

        ' CreateObject("WScript.Shell").Run cmd, 1, True
        
        ' Dim wsh : Set wsh = CreateObject("WScript.Shell")
        ' Call wsh.Run(command,1,True)

        ' wsh.Run cmd,0,True  'RUNS HIDDEN
        ' wsh.Run cmd,1,True 'RUNS NOT HIDDEN
        '  SEE: https://www.vbsedit.com/html/6f28899c-d653-4555-8a59-49640b0e32ea.asp
        '  object.Run(strCommand, [intWindowStyle], [bWaitOnReturn]) 

        ' wsh.Exec(command)
        ' outPut = wsh.Exec(command).Stdout.ReadAll()  ' ADDITIONALLY READ OUTPUT
        ' Set wsh = Nothing

        ' Read Infofile
        getAllMetadata = readFile(InfoFile)
        deleteAFile(InfoFile)
    End If
End Function

Sub deleteAFile(filespec)
    If fso.FileExists(filespec) Then
        On Error Resume Next
        fso.DeleteFile(filespec)
    End If
End Sub

Sub cleanUp()
    ' Delete Old Files
    counter=1
    filename = RootPath&"exiftoolinfo"&counter&".txt"
    finished=False
    While fso.FileExists(filename)
        On Error Resume Next
        fso.DeleteFile(filename)
        counter=counter+1
        filename = RootPath&"exiftoolinfo"&counter&".txt"
    Wend
End Sub

Function writeFile(text,outfile)
    If fso.FileExists(outfile) Then
        deleteAFile(outfile)
    End If
    'Dim objStream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.CharSet = "utf-8"
    objStream.Open
    On Error Resume Next
    objStream.WriteText text
    objStream.SaveToFile outfile, 2
    ' clean up
    objStream.Close
    Set objStream = Nothing
End Function

Function readFile(infile)
    readFile=""
    Dim objStr
    Set objStr = CreateObject("ADODB.Stream")
    objStr.CharSet = "utf-8"
    objStr.Open
    On Error Resume Next
    objStr.LoadFromFile(infile)
    ' Set result = objStr.ReadText
    ' Set readFile = result& vbcrlf
    readFile = objStr.ReadText
    ' clean up
    objStr.Close
    Set objStr = Nothing
End Function

' ENTRY Procname
Sub showPanel
    '  SDB.Objects: https://www.mediamonkey.com/wiki/index.php?title=ISDBApplication::Objects
    If SDB.Objects("AllTagsPanelObj") Is Nothing Then
        Set Pnl=SDB.UI.NewDockablePersistentPanel("AllTagsPanel")
        ' Set Pnl=SDB.UI.NewDockablePanel
        Pnl.Caption="All Metadata (via ExifTool)"
        Pnl.Common.Width=250
        Pnl.DockedTo=1

        ' Adding HTML Parser/Browser
        Set Sxp=SDB.UI.NewActiveX(Pnl,"Shell.Explorer")
        Sxp.Common.ClientWidth = Pnl.Common.ClientWidth      ' ZvezdanD's trick #1 for 
        Sxp.Common.ClientHeight = Pnl.Common.ClientHeight    ' borderless dockable panel
        Sxp.Common.Anchors = 15                              ' 
        Sxp.Common.ControlName="AllTagsPanelSXP"

        SDB.Objects("AllTagsPanelObj")=Pnl
        'SDB.Objects("PLS").Common.ChildControl("Lyrics").Common.Visible=True
        SDB.Objects("AllTagsPanelObj").Common.Visible=True
    Else
        SDB.Objects("AllTagsPanelObj").Common.Visible=True
    End If
End Sub

Sub hidePanel
    If not SDB.Objects("AllTagsPanelObj") Is Nothing Then
        SDB.Objects("AllTagsPanelObj").Common.Visible=False
    End If
End Sub

Function formatHTML(ExifInfo)
    tab=chr(9)
    css="caption{font-weight: bold;font-size:large}"&_
        "th { text-align: right; }"&_
        "table{width:100%;/*border: 1px solid black;border-collapse: collapse;*/}"&_
        "th {color: white;background-color: gray;}"&_
        "th, td {border-bottom: 1px solid #ddd;}"&_
        "tr:hover {background-color: #f5f5f5;}"
        '"table{width:100%;border: 1px solid black;border-collapse: collapse;}"&_
    header="<html><head><style>"&css&"</style></head><body>"
    footer="</body>"
    
    ' format as table
    content="<table>" ' <caption>All Metadata (via ExifTool)</caption>
    text=Split(ExifInfo,chr(10))
    for each line in text
        content=content&"<tr>"
        lineSplit=Split(line,tab)
        first=True
        for each part in lineSplit
            If first Then
                content=content&"<th>"&part&"</th>"
                first=False
            Else
                content=content&"<td>"&part&"</td>"
            End If
            ' tmp = Replace(part,"\n","<br>")
        next
        content=content&"</tr>"
    next
    content=content&"</table>"

    formatHTML=header&content&footer
End Function

Sub writeDoc(content)
    If not (SDB.Objects("AllTagsPanelObj") Is Nothing) Then
        Set Doc=SDB.Objects("AllTagsPanelObj").Common.ChildControl("AllTagsPanelSXP").Interf.Document
        Doc.Write content
        Doc.Close
    End If
End Sub

'''' for DEBUGGING ''''
' Dim testForm1 : Set testForm1 = SDB.UI.NewForm
' testForm1.Common.SetRect 100, 100, 540, 190
' testForm1.Caption = "test..."
' testForm1.StayOnTop = True
' testForm1.FormPosition = 1
' 'testForm1.FormPosition = 4
'
' Set Btn2 = SDB.UI.NewButton(testForm1)
' Btn2.Caption = "reload"
' Btn2.Common.SetRect 10, 10, 100, 20
' Script.RegisterEvent Btn2, "OnClick", "reload"
'
' Set Btn3 = SDB.UI.NewButton(testForm1)
' Btn3.Caption = "showPanel"
' Btn3.Common.SetRect 10, 30, 100, 20
' Script.RegisterEvent Btn3, "OnClick", "showPanel"
'
' Set Btn3 = SDB.UI.NewButton(testForm1)
' Btn3.Caption = "hidePanel"
' Btn3.Common.SetRect 10, 50, 100, 20
' Script.RegisterEvent Btn3, "OnClick", "hidePanel"
'
' Set CreateLabel = SDB.UI.NewLabel(testForm1)
' CreateLabel.Common.SetRect 120, 50, 100, 20
' CreateLabel.Caption = "pCaption for Label"
'
' Set Btn1 = SDB.UI.NewButton(testForm1)
' Btn1.Caption = "cleanUp"
' Btn1.Common.SetRect 120, 10, 100, 20
' Script.RegisterEvent Btn1, "OnClick", "cleanUp"
'
' Set Edt1 = SDB.UI.NewEdit(testForm1)
' Edt1.Text = ""
' Edt1.Common.SetRect 120, 30, 500, 20
'
' Sub showForm()
'     ' Form1.ShowModal
'     testForm1.Common.Visible = True
'     testForm1.Common.BringToFront()
' End Sub
'
' Sub reload()
'     Script.Reload(Script.ScriptPath)
' End Sub
' SDB.MessageBox "Hallo",mtError,Array(mbOk)
'''' for DEBUGGING Ende''''