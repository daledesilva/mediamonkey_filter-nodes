' MediaMonkey Script

' NAME: FilterNodes
' Version: 1
' Author: Dale de Silva
' Website: www.oiltinman.com
' Date last edited: 16/03/2008

' INSTALL: Copy to Scripts\Auto\

' FILES THAT SHOULD BE PRESENT UPON A FRESH INSTALL:
' FilterNodes.vbs

Option Explicit

'Global Variables
Dim sblnNewNode : sblnNewNode = True
Dim sarrName()
ReDim sarrName(0)
Dim sarrID()
ReDim sarrID(0)

Sub firstRun()
	
	SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"startupCheck") = True
	SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"hideLibrary") = True
	SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"keywords") = True
	SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"altIcons") = True
	SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"expPlaylist") = True
	SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"expMainNodes") = True
	
	OnStartup()
End Sub
 
Sub OnStartup()
  Dim Tree : Set Tree = SDB.MainTree

  'Dim podcasts
  'Set podcasts = Tree.Node_Playlists.NextSiblingNode
  'podcasts.Caption = "Subscriptions"

  If SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"hideLibrary") = True Then
  	Tree.Node_Library.Visible = False
  End If
  
  'Expand Playlist Node
  If SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"expPlaylist") = True Then
  	Tree.Node_Playlists.Expanded = True
  End If
  
  'Change Icons
  If SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"altIcons") = True Then
  	Tree.Node_NowPlaying.IconIndex = 63
  	Tree.Node_Library.IconIndex = 14
  	
  	If SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"hideLibrary") = False Then
  		' could add a loop in here to change all the library subnodes to other icons too.. but what about their subnodes?
  	End If
  End If
  
  'calculate Nodes
  If SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"startupCheck") = True Then
  	calculateNodes
  End If
  
  'Create Options Sheet
  Dim aSheet : aSheet = SDB.UI.AddOptionSheet( "Filter Nodes Settings", Script.ScriptPath, "InitSheet", "", 0)
  
  initNodes

End Sub




Sub calculateNodes()
	Dim SQL, Iter, Name, ID, ContentType
	
	'Erase Nodes In Ini
	SDB.IniFile.DeleteSection("FilterNodes-Nodes_MM"&Round(SDB.VersionHi))
	
	SQL = "SELECT Filters.ID, Filters.Name, Filters.ContentType, Filters.Pos FROM Filters ORDER BY Filters.Pos"
	Set Iter = SDB.Database.OpenSQL(SQL)
	
	'Look for Main Node Filters
	Do
		Name = Iter.StringByName("Name")
		ID = Iter.StringByName("ID")
		ContentType = Iter.StringByName("ContentType")
		
		If InStr(Name,"--") = 0 And ContentType<>2 Then
			storeNameID Name,ID
		End If
		
		Iter.Next
	Loop While Not Iter.EOF
	
	'Have to run SQL again because I don't know how to set Iter back to start
	Set Iter = SDB.Database.OpenSQL(SQL)
	
	'Look for Sub Node Filters
	Do
		Name = Iter.StringByName("Name")
		ID = Iter.StringByName("ID")
		ContentType = Iter.StringByName("ContentType")
		
		If InStr(Name,"--") > 0 Then
			insertNameID Name,ID
		End If
		
		Iter.Next
	Loop While Not Iter.EOF
	
	writeNodesToIni
		
End Sub



Sub storeNameID(Name,ID)
	ReDim Preserve sarrName(UBound(sarrName)+1)
	ReDim Preserve sarrID(UBound(sarrID)+1)
	sarrName(UBound(sarrName)) = Name
	sarrID(UBound(sarrID)) = ID
End Sub





Sub insertNameID(Name,ID)
	Dim k

	For k = 1 To UBound(sarrName)
		If InStr(Name,sarrName(k)) = 1 Then
			
			splice sarrName,k,Name
			splice sarrID,k,ID
			
			Exit Sub
		End If	
	Next

End Sub





Sub splice(ByRef OrigArray,ByVal Slot, ByVal NewValue)
	Dim k
	Dim NewArray()
	
	'populate items before slot
	For k=1 To Slot
		ReDim Preserve NewArray(k)
		NewArray(k) = OrigArray(k)
	Next
	
	'slot in New Value
	ReDim Preserve NewArray(UBound(NewArray)+1)
	NewArray(UBound(NewArray)) = NewValue
	
	' populate the rest
	For k=Slot+2 To UBound(OrigArray)+1
		ReDim Preserve NewArray(UBound(NewArray)+1)
		NewArray(k) = OrigArray(k-1)
	Next
	
	'copy data back over original array
	ReDim OrigArray(UBound(NewArray))
	For k=0 To UBound(NewArray)
		OrigArray(k) = NewArray(k)
	Next
End Sub




Sub writeNodesToIni()
	Dim k
	
	SDB.IniFile.StringValue("FilterNodes-Nodes_MM"&Round(SDB.VersionHi),"NodeTotal") = UBound(sarrName)
	For k = 1 To UBound(sarrName)
		SDB.IniFile.StringValue("FilterNodes-Nodes_MM"&Round(SDB.VersionHi),"Node"&k) = sarrName(k)
		SDB.IniFile.StringValue("FilterNodes-Nodes_MM"&Round(SDB.VersionHi),"Node"&k&"ID") = sarrID(k)
	Next
	Erase sarrName
	Erase sarrID
End Sub




Sub initNodes()
	Dim k
	
	Dim Tree : Set Tree = SDB.MainTree
	Dim aMainNode, aSubNode, oldMainNode, oldSubNode, firstSubLoop, Name, divPos
	Dim firstLoop : firstLoop = True
	
	For k=1 To SDB.IniFile.StringValue("FilterNodes-Nodes_MM"&Round(SDB.VersionHi),"NodeTotal")
		Name = SDB.IniFile.StringValue("FilterNodes-Nodes_MM"&Round(SDB.VersionHi),"Node"&k)
		
		divPos = InStr(Name,"--")
		If divPos = 0 Then
			Set aMainNode = Tree.CreateNode
			aMainNode.IconIndex = 14
			aMainNode.Caption = Name
			If firstLoop Then
				Tree.AddNode Tree.Node_NowPlaying, aMainNode, 1
				firstLoop = False
			Else
				Tree.AddNode oldMainNode, aMainNode, 1
			End If
			aMainNode.UseScript = Script.ScriptPath
			aMainNode.OnFillTracksFunct = "fillMainNode"
			firstSubLoop = True
			
			Set oldMainNode = aMainNode
		Else
			Name = Right(Name,Len(Name)-(divPos+2))
			Name = LTrim(Name)
			
			Set aSubNode = Tree.CreateNode
			aSubNode.IconIndex = 21
			aSubNode.Caption = Name
			If firstSubLoop Then
				Tree.AddNode oldMainNode,aSubNode,2
				firstSubLoop = False
			Else
				Tree.AddNode oldSubNode,aSubNode,1
			End If
			aSubNode.UseScript = Script.ScriptPath
			aSubNode.OnFillTracksFunct = "fillSubNode"
			
			Set oldSubNode = aSubNode
			
			If SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"expMainNodes") = True Then
				oldMainNode.Expanded = True
			End If
			
		End If
		
	Next

End Sub




Sub fillMainNode(theNode)
	
	'This is an attempt to error protect the Tree.CurrentNode = Sel Line, but it treats it like a local variable for some reason
	'If sblnNewNode Then
	'	MsgBox(sblnNewNode)
	'	sblnNewNode = False
	'	MsgBox(sblnNewNode)
	'Else
	'	MsgBox("Second")
	'	sblnNewNode = True
	'	Exit Sub
	'End If
	
	
	Dim Tree : Set Tree = SDB.MainTree
	Dim Trcks : Set Trcks = SDB.MainTracksWindow
	Dim query, k, Sel, SQL, Order
	
	For k=1 To SDB.IniFile.StringValue("FilterNodes-Nodes_MM"&Round(SDB.VersionHi),"NodeTotal")
		If theNode.Caption = SDB.IniFile.StringValue("FilterNodes-Nodes_MM"&Round(SDB.VersionHi),"Node"&k) Then
			'Change filter & reselect node (this causes an infinite loop if a dialog appears during)
			Set Sel = Tree.CurrentNode
  			SDB.Database.ActiveFilterID = SDB.IniFile.StringValue("FilterNodes-Nodes_MM"&Round(SDB.VersionHi),"Node"&k&"ID")
  			
  			query = SDB.Database.GetFilterQuery(SDB.IniFile.StringValue("FilterNodes-Nodes_MM"&Round(SDB.VersionHi),"Node"&k&"ID"))
  			
  			'Annoying need for hack
			If SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"keywords") Then
				If theNode.Caption = "Albums" Then
					'If the user is trying to only show actual albums - this helps them by modifying the sql (as filter criteria doesn't currently allow it)
					SQL = "(" & query & ") AND Songs.Album<>''"
					Order = " ORDER BY Songs.AlbumArtist, Songs.Album, CAST(DiscNumber AS INTEGER), CAST(TrackNumber AS INTEGER), Songs.SongTitle"
				Else
	  				SQL = query
	  				If theNode.Caption = "Non-Albums" Then
	  					Order = " ORDER BY Songs.Artist, Songs.SongTitle"
	  				Else
	  					Order = ""
	  				End If
	  			End If
	  		Else
	  			SQL = query
	  			Order = ""
	  		End If
  			
  			If SQL = "" Then
  				Trcks.AddTracksFromQuery(Order)
  			Else
				Trcks.AddTracksFromQuery("WHERE " & SQL & Order)
			End If
				
			Trcks.FinishAdding
			
			Tree.CurrentNode = Sel
  			SDB.ProcessMessages
  			
			Exit Sub
			
		End If
	Next

End Sub




Sub fillSubNode(theNode)
	
	'If sblnNewNode Then
	'	MsgBox("First")
	'	sblnNewNode = False
	'Else
	'	MsgBox("Second")
	'	sblnNewNode = True
	'	Exit Sub
	'End If
	
	Dim Tree : Set Tree = SDB.MainTree
	Dim Trcks : Set Trcks = SDB.MainTracksWindow
	Dim theParent, theIniName, theCaption, SQL, Order
	Dim query, k, Sel
	
	Set theParent = Tree.ParentNode(theNode)
	For k=1 To SDB.IniFile.StringValue("FilterNodes-Nodes_MM"&Round(SDB.VersionHi),"NodeTotal")
		theIniName = SDB.IniFile.StringValue("FilterNodes-Nodes_MM"&Round(SDB.VersionHi),"Node"&k)
		If InStr(theIniName,theParent.Caption) = 1 Then
			If Right(theIniName,Len(theNode.Caption)) = theNode.Caption Then
				theIniName = Left(theIniName,Len(theIniName)-Len(theNode.Caption))
				theIniName = Right(theIniName,Len(theIniName)-Len(theParent.Caption))
				theIniName = LTrim(theIniName)
				theIniName = RTrim(theIniName)
				If theIniName = "--" Then
					'Change filter & reselect node (this causes an infinite loop if a dialog appears during)
					Set Sel = Tree.CurrentNode
					SDB.Database.ActiveFilterID = SDB.IniFile.StringValue("FilterNodes-Nodes_MM"&Round(SDB.VersionHi),"Node"&k&"ID")
  			
		  			query = SDB.Database.GetFilterQuery(SDB.IniFile.StringValue("FilterNodes-Nodes_MM"&Round(SDB.VersionHi),"Node"&k&"ID"))
  					
  					'Annoying need for hack
  					If SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"keywords") = True Then
	  					If theNode.Caption = "Albums" Then
							'If the user is trying to only show actual albums - this helps them by modifying the sql (as filter criteria doesn't currently allow it)
							SQL = "(" & query & ") AND Songs.Album<>''"
							Order = " ORDER BY Songs.AlbumArtist, Songs.Album, CAST(DiscNumber AS INTEGER), CAST(TrackNumber AS INTEGER), Songs.SongTitle"
						Else
			  				SQL = query
			  				If theNode.Caption = "Non-Albums" Then
			  					Order = " ORDER BY Songs.Artist, Songs.SongTitle"
			  				Else
			  					Order = ""
			  				End If
			  			End If
			  		Else
			  			SQL = query
			  			Order = ""
			  		End If
		  			
		  			If SQL = "" Then
		  				Trcks.AddTracksFromQuery(Order)
		  			Else
						Trcks.AddTracksFromQuery("WHERE " & SQL & Order)
					End If
					Trcks.FinishAdding
					
					'setting the selection back at this point doesn't run the fillSubNode function for some reason
					'which is handy for this situaion but keep an eye on it because if it starts doing it in the furtue it will cause an infinite loop
					' it already causes an infinite loop if a msgbox pops up here (eg, error with query)
		  			Tree.CurrentNode = Sel
		  			SDB.ProcessMessages
		  			
					Exit Sub
					
				End If
			End If
		End If
	Next
End Sub












'Option Sheet Subs'
Sub ChangeHideLibrarySetting(Control)
  Dim Tree : Set Tree = SDB.MainTree

  If Control.Common.ControlName = "hideLibrarySetting1" Then
    SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"hideLibrary") = False
    Tree.Node_Library.Visible = True
  Else
    SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"hideLibrary") = True
    Tree.Node_Library.Visible = False
  End If
End Sub

Sub ChangeExpPlaylistSetting(Control)
  Dim Tree : Set Tree = SDB.MainTree

  If Control.Common.ControlName = "expPlaylistSetting1" Then
    SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"expPlaylist") = True
    Tree.Node_Playlists.Expanded = True
  Else
    SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"expPlaylist") = False
    Tree.Node_Playlists.Expanded = False
  End If
End Sub

Sub ChangeExpMainNodesSetting(Control)
  If Control.Common.ControlName = "expMainNodesSetting1" Then
    SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"expMainNodes") = True
  Else
    SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"expMainNodes") = False
  End If
End Sub

Sub ChangeAltIconsSetting(Control)
  Dim Tree : Set Tree = SDB.MainTree

  If Control.Common.ControlName = "altIconsSetting1" Then
    SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"altIcons") = True
  Else
    SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"altIcons") = False
  End If
End Sub

Sub ChangeKeywordsSetting(Control)
  If Control.Common.ControlName = "keywordsSetting1" Then
    SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"keywords") = True
  Else
    SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"keywords") = False
  End If
End Sub




' Draw Option Sheet - PAGE 2'
Sub InitSheet(Sheet)
  Dim aLabel, aRadio, aPanel, tempText
	  
  
  'library visibility'
  Set aPanel = SDB.UI.NewTranspPanel(Sheet)
  aPanel.Common.SetRect 5,4,462,67
  aPanel.Common.ControlName = "hideLibrarySettingPanel"
  
  Set aLabel = SDB.UI.NewLabel(aPanel)
  aLabel.Common.SetRect 9,4,400,17
  aLabel.Caption = "Please note that hiding your Library Node does NOT prevent you from using" + vbCRLF + "'More From Same' links. They will still work as expected." + vbCRLF + "Should your Library Node by visible?"
  
  Set aRadio = SDB.UI.NewRadioButton(aPanel)
  aRadio.Caption = "Of course, it's so useful"
  aRadio.Common.SetRect 10,45,200,20
  aRadio.Common.ControlName = "hideLibrarySetting1"
  If SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"hideLibrary") = False Then
    aRadio.Checked = True
  End If
  Script.RegisterEvent aRadio.Common, "OnClick", "ChangeHideLibrarySetting"
  
  Set aRadio = SDB.UI.NewRadioButton(aPanel)
  aRadio.Caption = "Nah, just show me what I need"
  aRadio.Common.SetRect 210,45,200,20
  aRadio.Common.ControlName = "hideLibrarySetting2"
  If SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"hideLibrary") = True Then
    aRadio.Checked = True
  End If
  Script.RegisterEvent aRadio.Common, "OnClick", "ChangeHideLibrarySetting"
  
  
  
  'expand playlist'
  Set aPanel = SDB.UI.NewTranspPanel(Sheet)
  aPanel.Common.SetRect 5,69,462,47
  aPanel.Common.ControlName = "expPlaylistSettingPanel"
  
  Set aLabel = SDB.UI.NewLabel(aPanel)
  aLabel.Common.SetRect 9,8,400,17
  aLabel.Caption = "When MediaMonkey starts, should the playlist node be expanded?"
  
  Set aRadio = SDB.UI.NewRadioButton(aPanel)
  aRadio.Caption = "Yeah, I use them all the time"
  aRadio.Common.SetRect 10,25,200,20
  aRadio.Common.ControlName = "expPlaylistSetting1"
  If SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"expPlaylist") = True Then
    aRadio.Checked = True
  End If
  Script.RegisterEvent aRadio.Common, "OnClick", "ChangeExpPlaylistSetting"
  
  Set aRadio = SDB.UI.NewRadioButton(aPanel)
  aRadio.Caption = "Nope, when I want it, I'll expand it myself"
  aRadio.Common.SetRect 210,25,300,20
  aRadio.Common.ControlName = "expPlaylistSetting2"
  If SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"expPlaylist") = False Then
    aRadio.Checked = True
  End If
  Script.RegisterEvent aRadio.Common, "OnClick", "ChangeExpPlaylistSetting"
  
  
  
  'expand Main Nodes'
  Set aPanel = SDB.UI.NewTranspPanel(Sheet)
  aPanel.Common.SetRect 5,114,462,47
  aPanel.Common.ControlName = "expMainNodesSettingPanel"
  
  Set aLabel = SDB.UI.NewLabel(aPanel)
  aLabel.Common.SetRect 9,8,400,17
  aLabel.Caption = "When MediaMonkey starts, should the your Main Filter Nodes be expanded (needs restart)?"
  
  Set aRadio = SDB.UI.NewRadioButton(aPanel)
  aRadio.Caption = "Yeppers!"
  aRadio.Common.SetRect 10,25,200,20
  aRadio.Common.ControlName = "expMainNodesSetting1"
  If SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"expMainNodes") = True Then
    aRadio.Checked = True
  End If
  Script.RegisterEvent aRadio.Common, "OnClick", "ChangeExpMainNodesSetting"
  
  Set aRadio = SDB.UI.NewRadioButton(aPanel)
  aRadio.Caption = "Nope, I don't use the subnodes often"
  aRadio.Common.SetRect 210,25,200,20
  aRadio.Common.ControlName = "expMainNodesSetting2"
  If SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"expMainNodes") = False Then
    aRadio.Checked = True
  End If
  Script.RegisterEvent aRadio.Common, "OnClick", "ChangeExpMainNodesSetting"
  
  
  
  'alt Icons'
  Set aPanel = SDB.UI.NewTranspPanel(Sheet)
  aPanel.Common.SetRect 5,159,462,47
  aPanel.Common.ControlName = "altIconsSettingPanel"
  
  Set aLabel = SDB.UI.NewLabel(aPanel)
  aLabel.Common.SetRect 9,8,400,17
  aLabel.Caption = "Do you want to use alternate icons on your Music Nodes (needs restart)?"
  
  Set aRadio = SDB.UI.NewRadioButton(aPanel)
  aRadio.Caption = "Yaha, I like subdued icons"
  aRadio.Common.SetRect 10,25,200,20
  aRadio.Common.ControlName = "altIconsSetting1"
  If SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"altIcons") = True Then
    aRadio.Checked = True
  End If
  Script.RegisterEvent aRadio.Common, "OnClick", "ChangeAltIconsSetting"
  
  Set aRadio = SDB.UI.NewRadioButton(aPanel)
  aRadio.Caption = "Nope, I'm used to the default ones"
  aRadio.Common.SetRect 210,25,200,20
  aRadio.Common.ControlName = "altIconsSetting2"
  If SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"altIcons") = False Then
    aRadio.Checked = True
  End If
  Script.RegisterEvent aRadio.Common, "OnClick", "ChangeAltIconsSetting"
  
  
  
  'keywords'
  Set aPanel = SDB.UI.NewTranspPanel(Sheet)
  aPanel.Common.SetRect 5,204,462,87
  aPanel.Common.ControlName = "keywordsPanel"
  
  Set aLabel = SDB.UI.NewLabel(aPanel)
  aLabel.Common.SetRect 9,8,400,17
  aLabel.Caption = "Some querys can be convoluted to set up in a filter." + vbCRLF + "If you put certain keywords into the name of the filter," + vbCRLF + "Filter Nodes will filter the tracks using extra criteria (see list below)." + vbCRLF + "Would you like keywords to be active?"
  
  Set aRadio = SDB.UI.NewRadioButton(aPanel)
  aRadio.Caption = "Sure Do!"
  aRadio.Common.SetRect 10,65,200,20
  aRadio.Common.ControlName = "keywordsSetting1"
  If SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"keywords") = True Then
    aRadio.Checked = True
  End If
  Script.RegisterEvent aRadio.Common, "OnClick", "ChangeKeywordsSetting"
  
  Set aRadio = SDB.UI.NewRadioButton(aPanel)
  aRadio.Caption = "Nah, it's conflicting with my names"
  aRadio.Common.SetRect 210,65,200,20
  aRadio.Common.ControlName = "keywordsSetting2"
  If SDB.IniFile.BoolValue("FilterNodes_MM"&Round(SDB.VersionHi),"keywords") = False Then
    aRadio.Checked = True
  End If
  Script.RegisterEvent aRadio.Common, "OnClick", "ChangeKeywordsSetting"
  
  Set aPanel = SDB.UI.NewPanel(Sheet)
  aPanel.Common.SetRect 5,291,462,88
  'aPanel.Common.ControlName = "keywordsPanel"
  
  Set aLabel = SDB.UI.NewLabel(aPanel)
  aLabel.Common.SetRect 9,8,450,80
  aLabel.Caption = "Keywords:"
  
  Set aLabel = SDB.UI.NewLabel(aPanel)
  aLabel.Common.SetRect 15,28,450,80
  tempText =                     "'Albums'"
  tempText = tempText + vbCRLF + "'Non-Albums'"
  aLabel.Caption = tempText
  
  Set aLabel = SDB.UI.NewLabel(aPanel)
  aLabel.Common.SetRect 105,28,450,80
  tempText = 					 "-  All tracks must have an album name."
  tempText = tempText + vbCRLF + "-  Default sort order will be Artist then Song Title."
  aLabel.Caption = tempText
  
  
  
  
  
End Sub