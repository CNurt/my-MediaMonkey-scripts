'
' MediaMonkey Script
'
' NAME: TagToolbox
'
' AUTHOR: C. Seeling
' DATE : 2020-09-29
'
' ENTRY Point is showFormTlBx


''''''''''''''''''''''''''''
'' Global Variables ''
''''''''''''''''''''''

Btny = 20
Btnx = 90

' various genres or lanuage keywords for dropdown list
Keywords = Array(_
	"= Languages =",_
	"_Instrumental",_
	"_mul",_
	"_unknown",_
	"Chinese Cantonese",_
	"French",_
	"",_
    "= Genres =",_
	"_cover",_
	"_mp3.320")
	"_flac",_
	"_HQcopy",_

'''''''''''''''''''''''''''
'' GUI Tag Toolbox ''
'''''''''''''''''''''''''''

Dim FormTlBx : Set FormTlBx = SDB.UI.NewForm
	FormTlBx.Common.SetRect 0, 0, 193, 250
	FormTlBx.Caption = "Tag Toolbox"
	FormTlBx.StayOnTop = True
	FormTlBx.FormPosition = 4

Set PnlTag=SDB.UI.NewTranspPanel(FormTlBx)
    PnlTag.Common.SetRect 0,0,210, 90
    PnlTag.Common.ControlName="PanelTags"
	Set BtnGerman = SDB.UI.NewButton(PnlTag)
		BtnGerman.Caption = "German"
		BtnGerman.Common.SetRect 0, 0, Btnx, Btny
		Script.RegisterEvent BtnGerman, "OnClick", "LangGerman"
	Set BtnEnglish = SDB.UI.NewButton(PnlTag)
		BtnEnglish.Caption = "English"
		BtnEnglish.Common.SetRect 0, 20, Btnx, Btny
		Script.RegisterEvent BtnEnglish, "OnClick", "LangEnglish"
	Set BtnChineseMandarin = SDB.UI.NewButton(PnlTag)
		BtnChineseMandarin.Caption = "Chinese Mandarin"
		BtnChineseMandarin.Common.SetRect 90, 0, Btnx, Btny
		Script.RegisterEvent BtnChineseMandarin, "OnClick", "LangChineseMandarin"
	Set BtnJapanese = SDB.UI.NewButton(PnlTag)
		BtnJapanese.Caption = "Japanese"
		BtnJapanese.Common.SetRect 90, 20, Btnx, Btny
		Script.RegisterEvent BtnJapanese, "OnClick", "LangJapanese"
	Set DDetc = SDB.UI.NewDropDown(PnlTag)
		DDetc.Common.SetRect 0, 40, Btnx*2, Btny
		For Each keyword In  Keywords
			DDetc.AddItem keyword
		Next
		'OnSelect
		'OnChange
		Script.RegisterEvent DDetc, "OnChange", "DDetcOnChange"
		Script.RegisterEvent DDetc, "OnSelect", "DDetcOnChange"
	Set LblDescription = SDB.UI.NewLabel(PnlTag)
		LblDescription.Caption = "for multiple tags delimiter = ;"
		LblDescription.Common.SetRect 5, DDetc.Common.Top+DDetc.Common.Height+2, 180, 20
		LblDescription.Multiline = True	

Set PnlCat=SDB.UI.NewTranspPanel(FormTlBx)
    PnlCat.Common.SetRect 0,PnlTag.Common.Height,Btnx-20,40
    PnlCat.Common.ControlName="PanelCategory"
	Set RBCat=SDB.UI.NewRadioButton(PnlCat)
	    RBCat.Common.SetRect 0,0,60,20
	    RBCat.Caption="Genre"
	    RBCat.Checked=False
	    RBCat.Common.ControlName="RBCatGenre"
	Set RBCat=SDB.UI.NewRadioButton(PnlCat)
	    RBCat.Common.SetRect 0,20,60,20
	    RBCat.Caption="Language"
	    RBCat.Checked=False
	    RBCat.Common.ControlName="RBCatLang"

Set PnlAction=SDB.UI.NewTranspPanel(FormTlBx)
    PnlAction.Common.SetRect PnlCat.Common.Width+10,PnlCat.Common.Top,Btnx+50,100
    'PnlAction.Common.SetRect 0,300,100,100
    PnlAction.Common.ControlName="PanelAction"
	Set BtnAdd = SDB.UI.NewButton(PnlAction)
		BtnAdd.Caption = "Add"
		BtnAdd.Common.SetRect 0, 0, Btnx-40, Btny
		Script.RegisterEvent BtnAdd, "OnClick", "BtnAddClick"
	Set BtnAssign = SDB.UI.NewButton(PnlAction)
		BtnAssign.Caption = "Assign"
		BtnAssign.Common.SetRect 0, 20, Btnx-40, Btny
		Script.RegisterEvent BtnAssign, "OnClick", "BtnAssignClick"
		Set CBInstant = SDB.UI.NewCheckBox(PnlAction)
			CBInstant.Caption = "instant"
			CBInstant.Common.SetRect BtnAdd.Common.Left+BtnAdd.Common.Width, 0, 30, 20
			CBInstant.Checked = False
	Set BtnRemove = SDB.UI.NewButton(PnlAction)
		BtnRemove.Caption = "Remove"
		BtnRemove.Common.SetRect 0, 50, Btnx-40, Btny
		Script.RegisterEvent BtnRemove, "OnClick", "BtnRemoveClick"
		Set CBRemove = SDB.UI.NewCheckBox(PnlAction)
			CBRemove.Caption = "!"
			CBRemove.Common.SetRect BtnRemove.Common.Left+BtnRemove.Common.Width, BtnRemove.Common.Top, 30, 20
			CBRemove.Checked = False
		Set CBSubstr = SDB.UI.NewCheckBox(PnlAction)
			CBSubstr.Caption = "substr"
			CBSubstr.Common.SetRect BtnRemove.Common.Left, BtnRemove.Common.Top+BtnRemove.Common.Height, 50, 20
			CBSubstr.Checked = False
Set BtncopyLangToGenre = SDB.UI.NewButton(FormTlBx)
	BtncopyLangToGenre.Caption = "copy LANGUAGE --> Genre"
	BtncopyLangToGenre.Common.SetRect 5, PnlAction.Common.Top+PnlAction.Common.Height, 170, 20
	Script.RegisterEvent BtncopyLangToGenre, "OnClick", "copyLangToGenre"
		
''''''''''''''''''
'' Subroutines ''
''''''''''''''''''

Sub showFormTlBx()
    FormTlBx.Common.Visible = True
    FormTlBx.Common.BringToFront()
End Sub

Sub LangGerman()
    DDetc.Text = "German"
	Call TextChange
	If CBInstant.Checked Then
		Call BtnAddClick
	End If
End Sub

Sub LangEnglish()
    DDetc.Text = "English"
	Call TextChange
	If CBInstant.Checked Then
		Call BtnAddClick
	End If
End Sub

Sub LangChineseMandarin()
    DDetc.Text = "Chinese Mandarin"
	Call TextChange
	If CBInstant.Checked Then
		Call BtnAddClick
	End If
End Sub

Sub LangJapanese()
    DDetc.Text = "Japanese"
	Call TextChange
	If CBInstant.Checked Then
		Call BtnAddClick
	End If
End Sub

Sub DDetcOnChange(ctrl)
	Call TextChange
End Sub
Sub TextChange
	CBRemove.Checked = False
End Sub

sub BtnAddClick
	TagName = getTagname
    If TagnameChoosen(Tagname) Then
        Select Case Tagname
            Case "Genre"
                call addToTag(TagName,DDetc.Text,"","")
            Case "Lang"
                call addToTag_LANG(DDetc.Text,"","")
            Case Else	pass = NULL
        End Select
    	
    End If
end Sub

sub BtnAssignClick
	TagName = getTagname
	If TagnameChoosen(Tagname) Then
    	call assignToTag(TagName,DDetc.Text,"","")
    End If
end Sub

Sub BtnRemoveClick
	TagName = getTagname
	If TagnameChoosen(Tagname) Then
		If Not CBRemove.Checked Then
			SDB.MessageBox "PROTECTED!" & vbCrLf & vbCrLf & "Checkbox '!' next to the Remove button needs to be checked!", mtInformation, Array(mbOk)
            Else
                Select Case Tagname
                    Case "Genre"
                        Call removeFromTag(TagName, CBSubstr.Checked)
                    Case "Lang"
                        Call removeFromTag(TagName, CBSubstr.Checked)
                        Call removeLangFromGenre(CBSubstr.Checked)
                    Case Else	writeTag = False
                End Select
			End If	
	End If
End Sub


Sub addToTag(Tagname,values,prefix,suffix)
    if len(values) > 0 Then
        Set list = SDB.SelectedSongList
        For i = 0 to list.count - 1
            Set objSongData = list.Item(i)
            Keywords = Split(values,";")
            tagvalues = readTag(objSongData,Tagname)
            ' genre = objSongData.Genre
            newKeywords = ""
                for each kword in Keywords
                    kword = trim(kword)
                    ' if InStr(1,tagvalues,kword,1) = 0 then
                    if not findKeywordInList(kword,tagvalues) then
                        if len(newKeywords) > 0 then
                            newKeywords = newKeywords & ";"
                        end if
                        newKeywords = newKeywords & prefix & kword & suffix
                    end if
                Next
            if len(newKeywords) > 0 then
                if len(tagvalues) > 0 and right(trim(tagvalues),1) <> ";" then
                    tagvalues = tagvalues & ";"
                end if
                ' objSongData.Genre = tagvalues & newKeywords
                writeTag objSongData,Tagname,tagvalues & newKeywords
            end if
        Next
        list.UpdateAll
    end if
End Sub


Sub addToTag_LANG(values,prefix,suffix)
    Tagname = "Lang"
    if len(values) > 0 Then
        Set list = SDB.SelectedSongList
        For i = 0 to list.count - 1
            Set objSongData = list.Item(i)
            Keywords = Split(values,";")
            tagvalues = readTag(objSongData,Tagname)
            ' genre = objSongData.Genre
            newKeywords = ""
                for each kword in Keywords
                    kword = trim(kword)
                    ' if InStr(1,tagvalues,kword,1) = 0 then
                    if not findKeywordInList(kword,tagvalues) then
                        if len(newKeywords) > 0 then
                            newKeywords = newKeywords & ";"
                        end if
                        newKeywords = newKeywords & prefix & kword & suffix
                    end if
                Next
            if len(newKeywords) > 0 then
                if len(tagvalues) > 0 and right(trim(tagvalues),1) <> ";" then
                    tagvalues = tagvalues & ";"
                end if
                ' objSongData.Genre = tagvalues & newKeywords
                writeTag objSongData,Tagname,tagvalues & newKeywords
            end if
        Next
        list.UpdateAll
        Call copyLangToGenre
    end if
End Sub

Sub assignToTag(Tagname,values,prefix,suffix)
    if len(values) > 0 Then
        Set list = SDB.SelectedSongList
        For i = 0 to list.count - 1
            Set objSongData = list.Item(i)
            Keywords = Split(values,";")
            newKeywords = ""
                for each kword in Keywords
                    kword = trim(kword)
                    if len(newKeywords) > 0 then
                        newKeywords = newKeywords & ";"
                    end if
                    newKeywords = newKeywords & prefix & kword & suffix
                Next
            if len(newKeywords) > 0 then
                writeTag objSongData,Tagname,newKeywords
            end if
        Next
        list.UpdateAll
    end if
End Sub

Sub removeFromTag(Tagname,removeSubstring)
    Set list = SDB.SelectedSongList
    str = DDetc.Text
    ' https://www.mediamonkey.com/wiki/index.php?title=ISDBApplication::MessageBox
    ' answer = SDB.MessageBox( "remove: " & str, mtConfirmation, Array(mbYes,mbNo) )
    ' SDB.MessageBox "your answwer: " & answer, mtInformation, Array(mbOk)
    ' if len(str) > 0 and answer = 6 then
    if len(str) > 0 then
        For i = 0 to list.count - 1
            Set objSongData = list.Item(i)
            ' genres = Split(objSongData.Genre,";")
            genres = Split(readTag(objSongData,Tagname),";")
            values = Split(str,";")
            newKeywords = ""
            changed = False
            for each genre in genres
                genre = trim(genre)
                remove = False
                for each keyword in values
                    keyword=trim(keyword)
                    if len(keyword) > 0 then
                        if (not removeSubstring and genre = keyword) or ((removeSubstring) and InStr(1,genre,keyword,1) > 0) then
                            remove = True
                        end if
                    end if
                next
                if not remove then
                    if len(newKeywords) > 0 then
                        newKeywords = newKeywords & ";"
                    end if
                    newKeywords = newKeywords & genre
                else
                    changed = True
                end if
            next 
            if changed then
                writeTag objSongData,Tagname,newKeywords
            end if
        Next
    end if
    list.UpdateAll
end sub

Sub removeLangFromGenre(removeSubstring)
    Tagname = "Genre"
    Set list = SDB.SelectedSongList
    str = DDetc.Text
    ' https://www.mediamonkey.com/wiki/index.php?title=ISDBApplication::MessageBox
    ' answer = SDB.MessageBox( "remove: " & str, mtConfirmation, Array(mbYes,mbNo) )
    ' SDB.MessageBox "your answwer: " & answer, mtInformation, Array(mbOk)
    ' if len(str) > 0 and answer = 6 then
    if len(str) > 0 then
        For i = 0 to list.count - 1
            Set objSongData = list.Item(i)
            ' genres = Split(objSongData.Genre,";")
            genres = Split(readTag(objSongData,Tagname),";")
            values = Split(str,";")
            newKeywords = ""
            changed = False
            for each genre in genres
                genre = trim(genre)
                remove = False
                for each keyword in values
                    keyword=trim(keyword)
                    if len(keyword) > 0 then
                        if (not removeSubstring and genre = "["&keyword&"]") then
                            remove = True
                        else
                            If ( removeSubstring and (InStr(1,genre,keyword,1)>0 and Left(genre,1)="[" and Right(genre,1)="]") ) Then remove = True
                        end if
                    end if
                next
                if not remove then
                    if len(newKeywords) > 0 then
                        newKeywords = newKeywords & ";"
                    end if
                    newKeywords = newKeywords & genre
                else
                    changed = True
                end if
            next 
            if changed then
                writeTag objSongData,Tagname,newKeywords
            end if
        Next
    end if
    list.UpdateAll
end sub

Sub copyLangToGenre
    Set list = SDB.SelectedSongList
    For i = 0 to list.count - 1
        Set objSongData = list.Item(i)
        langs = Split(objSongData.Custom1,";")
        ' genres = Split(objSongData.Genre,";")
        genre = objSongData.Genre
        newGenres = ""
        ' for each g in genres
            for each l in langs
                l = trim(l)
                ' SDB.MessageBox "g = " &  trim(genre) & " ; l = " & trim(l), mtError, Array(mbOK)
                if InStr(1,genre,l,1) = 0 then
                ' if not foundTagInList(l,genre) then
                    if len(newGenres) > 0 then
                        newGenres = newGenres & ";"
                    end if
                    newGenres = newGenres & "[" & l & "]"
                end if
            next
        ' next 
        if len(newGenres) > 0 then
            ' SDB.MessageBox "len(newGenres) = " & len(newGenres) , mtError, Array(mbOK) ''' for DEBUGGING
            if len(genre) > 0 and right(trim(genre),1) <> ";" then
                objSongData.Genre = objSongData.Genre & ";"
            end if
            objSongData.Genre = objSongData.Genre & newGenres
        end If
        ' SDB.MessageBox "new Genre tag = " & objSongData.Genre , mtError, Array(mbOK)
    Next
    list.UpdateAll
End Sub

''''''''''''''''''
'' Functions ''
''''''''''''''''''

Function findKeywordInList(kword,listt)
    list = Split(listt,";")
    kword = trim(kword)
    if len(kword) > 0 then
        found = False
        for each l in list
            if kword = trim(l) then
                found = True
            end if
        next
    end if
    findKeywordInList = found
end Function

Function readTag(objSongData,Tagname)
	Select Case Tagname
	    Case "Genre"	readTag = objSongData.Genre
	    Case "Lang"	readTag = objSongData.Custom1
	    Case Else	readTag = ""
	End Select
End Function

Function writeTag(objSongData,Tagname,value)
	Select Case Tagname
	    Case "Genre"
	    	objSongData.Genre = value
	    	writeTag = True
	    Case "Lang"
	    	objSongData.Custom1 = value
	    	writeTag = True
	    Case Else	writeTag = False
	End Select
End Function


Function TagnameChoosen(Tagname)
	If Tagname = "" Then
		TagnameChoosen = False
		SDB.MessageBox "Choose Tag on which to apply the action! (Genre or Language)", mtInformation, Array(mbOk)
	Else
		TagnameChoosen = True
	End If
End Function

Function getTagname
	getTagname = ""
	If PnlCat.Common.ChildControl("RBCatGenre").Checked Then
		getTagname = "Genre"
	Else
		If PnlCat.Common.ChildControl("RBCatLang").Checked Then getTagname = "Lang"
	End If
End Function


''' DEBUGGING '''
'DEVELOPMENT SECTION below'

' Set BtnReloadScript = SDB.UI.NewButton(FormTlBx)
'	 BtnReloadScript.Caption = "reloadScript()"
'	 'BtnReloadScript.Common.SetRect 10, 80, 100, 20
'	 BtnReloadScript.Common.SetRect FormTlBx.Common.ClientWidth-100, FormTlBx.Common.ClientHeight-20, 100, 20
'	 BtnReloadScript.Common.Anchors = 4+8   ' The button is always in a constant distance from Bottom Right corner. ' https://www.mediamonkey.com/wiki/index.php?title=ISDBUICommon::Anchors
'	 Script.RegisterEvent BtnReloadScript, "OnClick", "reloadScript"
'
' Sub reloadScript()
'     Set FormTlBx = SDB.Objects("SyncTheSyncFormTlBx")
'     If Not (FormTlBx Is Nothing) Then
'         Script.UnregisterEvents FormTlBx
'         FormTlBx.Common.Visible = False
'         FormTlBx.Common.ControlName = ""
'         Set FormTlBx = Nothing  
'     End If
'     Script.UnRegisterAllEvents()
'     Script.Reload(Script.ScriptPath)
'     ' SDB.ScriptControl.UnRegisterAllEvents()
'     ' SDB.ScriptControl.Reload(SDB.ScriptControl.ScriptPath)
'     SDB.RefreshScriptItems()
'     Set SDB.Objects("SyncTheSyncFormTlBx") = Nothing
' End Sub
