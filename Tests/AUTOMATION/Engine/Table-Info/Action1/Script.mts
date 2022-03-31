Option Explicit
	'Set default timeout
	Setting("DefaultTimeOut")= 20000


	'Declarations Variables
	Dim iSyncTime : iSyncTime = 180
	Dim iGUIOverloadSyncTime : iGUIOverloadSyncTime = 3    

	Dim pOR
	Dim oRepository, oTest
	Dim aLoadOR
	Dim iLoadOR

	Public gORLookup,aORLookUp,gStack,gStackList

	Public gALMProject,gALMUsername 																	'V5.3

	'Declarations Arrays
	Public  aOR (), aORDepth(5)
	Public aGUIStack()


	'Get ALM project and username
	gALMProject = qcutil.QCConnection.ProjectName
	gALMUsername = qcutil.QCConnection.UserName

	'Constants
	'Object Repositories
	Dim sLoadOR
	
	Dim sCommand


	'Load libraries directly from ALM
	'Global Relative Paths
	createobject("QuickTest.Application").Folders.Add "[QualityCenter\Resources] Resources\AAF"

	'Default OR to be loaded
	sLoadOR="Object Repositories\FM.tsr,Object Repositories\CM.tsr,Object Repositories\TechGUI.tsr,Object Repositories\Inventory.tsr,Object Repositories\AdminGUI.tsr,Object Repositories\FMWeb.tsr"
	
	ReDim aOR(10,0)
	aORDepth(0) = 1
	
	
	ReDim aGlobalVar(0)
	ReDim aGUIStack(10,0)
	ReDim aFunctions(0,0)
	ReDim aObjectPreFix(12)
	'ReDim aPerfTime(4,0)

	'Initialize aObjectPreFix
	aObjectPreFix(0) = ":BTN~:CHK~:GRP~:LST~:MNU~:OPT~:STS~:TAB~:TBL~:TXT~:VIS~:WND~"
	aObjectPreFix(1) = ":JavaButton:JavaCheckBox:JavaObject:JavaRadioButton:JavaStaticText:" 'BTN
	aObjectPreFix(2) = ":JavaCheckBox:" 'chk
	aObjectPreFix(3) = ":JavaObject:" 'grp
	aObjectPreFix(4) = ":JavaList:" 'lst
	aObjectPreFix(5) = ":JavaMenu:" 'mnu
	aObjectPreFix(6) = ":JavaMenu:" 'opt
	aObjectPreFix(7) = ":JavaStaticText:JavaObject:" 'sts
	aObjectPreFix(8) = ":JavaTab:" 'tab
	aObjectPreFix(9) = ":JavaTable:" 'tbl
	aObjectPreFix(10) = ":JavaStaticText:JavaEdit:JavaObject:" 'txt
	aObjectPreFix(11) = ":JavaObject:" 'vis
	aObjectPreFix(12) = ":JavaDialog:JavaWindow:JavaInternalFrame:" 'wnd
	
	
	

	'Load Object Repositories into OR and into aOR
	aLoadOR = split(sLoadOR,",")
	RepositoriesCollection.RemoveAll
	Set oRepository = CreateObject("Mercury.ObjectRepositoryUtil")
	For iLoadOR = 0 to ubound(aLoadOR)
			pOR = Pathfinder.Locate(aLoadOR(iLoadOR) ) 
			oRepository.Load pOR
			call  fFrame_LoadAllObjectsProperties ("")
			RepositoriesCollection.Add pOR
	Next


'Get the object from the user
sCommand = inputbox ("Please enter the object name", "Object Checker")

Call fFrame_GUIStackORLookUp (sCommand)

sCommand =  fFrame_BuildCall(sCommand)



'Check if the methods exist on the object

'Sync
Dim dSyncTime
dSyncTime = dateadd("s",iSyncTime,Now)
Do while datediff("s",now,dSyncTime) > 0
		execute "bSync = cint(" & sCommand & ".GetROProperty(" & chr(34) & "enabled" & chr(34) & "))"
		If bSync =1 Then
				Exit do
		End If
Loop

'Check time to see if upper limit has been reached
If  bSync = 0  Then
        reporter.ReportEvent micDone,"Table Info" ,"Table Info : Fatal Error, sync error"
		ExitTest
end if

'Methods
Dim sDemo

On error resume next

'Check GetCellText
execute "sDemo = " & sCommand & ".Object.getCellText(0,0," & chr(34) & "," & chr(34) & ")"
If err.number <> 0 Then
	reporter.ReportEvent micFail,"Table Info" ,"Table Info [CellText] : Not available"
Else
	reporter.ReportEvent micPass,"Table Info" ,"Table Info [CellText] : Available"

End If
err.clear

'Check GetCellIconNames
execute "sDemo = " & sCommand & ".Object.getCellIconNames(0,0," & chr(34) & "," & chr(34) & ")"
If err.number <> 0 Then
	reporter.ReportEvent micFail,"Table Info" ,"Table Info [CellIconName] : Not available"
else
	reporter.ReportEvent micPass,"Table Info" ,"Table Info [CellIconName] : Available"
End If
err.clear


'Check TableMultipleHeaders
execute "sDemo = " & sCommand & ".Object.getALMIndexOfColumnInMultipleHeaderTable(sGroupHeaderString, " & chr(34) & "~#" & chr(34) & ", false)"
If err.number <> 0 Then
	reporter.ReportEvent micFail,"Table Info" ,"Table Info [MultipleHeaders] : Not available"
else
	reporter.ReportEvent micPass,"Table Info" ,"Table Info [MultipleHeaders] : Available"
End If
err.clear










Public function fFrame_GUIStackORLookUp (byref sCommandObject)

	'Perf Timer
	'fFrame_StartPerfTimer "GUI Stack OR Lookup [" & sCommandObject & "]", "AAF", 2

	Dim iCommandLoop, iStackLoop,iComLoop
	Dim sGUIStack,sCall,sType,sIndexSearch
	Dim aCommandOcc(),aPLookUp,aCommandObject

	'Initialise sGUIStack
	sGUIStack = "~"

	'Append # to end of GUIStack string
	'Check if no object has been passed in, from SendKeys
	If sCommandObject = "--~--" then
		'SendKeys command so dont append ::# ... just use #
		sGUIStack = ucase(mid(sGUIStack,1,instr(1,sGUIStack,"::")-1)) & "#"
	else
		sGUIStack = ucase(sGUIStack & sCommandObject & "#")
	End if

			
	'Simple look up to just find the current object ... and not the GUI stack history of where we should be
	aORLookUp = fFrame_ObjectRepoIndex("::" & ucase(sCommandObject) & "#")

	If aORLookUp(0) = 1 Then
			'Warning Simple look up passed
			reporter.ReportEvent micPass,"GUIStackORLookUp" ,"GUIStackORLookUp [" & sCommandObject & "] : Object [" & sCommandObject & "] has been located in the Object Repository" 
			fFrame_GUIStackORLookUp = 0
	elseif aORLookUp(0) = 0 then
			'Error message that object cannot be found within the OR
			reporter.ReportEvent micFail,"GUIStackORLookUp" ,"GUIStackORLookUp [" & sCommandObject & "] : Fatal Error, Object is not within the Oject Repository"
			ExitTest
	elseif aORLookUp(0) > 1 then
			'Error message that object cannot be found within the OR
			reporter.ReportEvent micFail,"GUIStackORLookUp" ,"GUIStackORLookUp [" & sCommandObject & "] : Fatal Error, Object is not unique within the Oject Repository"
			ExitTest					 
	End If

	'Set global index for other fucntions to get index into current Object, aORLookUp(1) will give the index to the first object/command found.
	gORLookup =  aORLookUp(1)
	fFrame_GUIStackORLookUp = 1
	'Perf Timer
	'fFrame_EndPerfTimer "GUI Stack OR Lookup [" & sCommandObject & "]"
End function

public function fFrame_ObjectRepoIndex (byref sCommandObject)

	'Perf Timer
	'fFrame_StartPerfTimer "Object Repository Index [" & sCommandObject & "]", "AAF", 2	

	Dim iStart,iCount,iConfirmLoop
	Dim aCommandOcc(),aConfirmed(1)

	Dim sPreFix,iPreFixIndex

	ReDim aCommandOcc(0)

					'Simple look up
					iStart = 1
					iCount = 0
					'Search for number of matches for a given object from QC command
					Do while instr(iStart,ucase(aOR(0,0)), ucase(sCommandObject) ) > 0
							iCount = iCount + 1
							redim preserve  aCommandOcc(iCount)
							aCommandOcc(iCount) = ubound(split(mid(ucase(aOR(0,0)),1,instr(iStart,ucase(aOR(0,0)),ucase(sCommandObject)) ),"~"))
							iStart = instr(iStart,ucase(aOR(0,0)),ucase(sCommandObject) ) + 1				
					Loop
					aCommandOcc(0) = iCount

					'If more than one object found whilst using Object name to identify the object within aOR then try 
					' to filter using the object type expected. The expected object type is determined by the 3 letter prefix of the object.
					If iCount > 1 Then
							'Multiple objects matching name
							aConfirmed(0)= 0
	
							'Loop through the found objects that match by name
							For iConfirmLoop = 1 to iCount
									'If the pre-fix shown in the name matches the required object type as defined in the aObjectPreFix array 
									If instr(1,aObjectPreFix(ubound(split(mid(ucase(aObjectPreFix(0)),1, instr(1,ucase(aObjectPreFix(0)),ucase(":" & mid(sCommandObject,instrrev(sCommandObject,"::") + 2,3) & "~"))),":"))),aOR(2,aCommandOcc(iConfirmLoop))) > 0  Then
											If  aConfirmed(0)= 0 Then
													'Set aConfirmed to hold the filtered data
													aConfirmed(0) = 1
													aConfirmed(1) = aCommandOcc(iConfirmLoop)
											else
													aConfirmed(0) = -1
													'Exit for as there are still more than 1 result even after applying object type filtering
													Exit for
											End If
									End If
							Next
	
							If  aConfirmed(0) = -1 Then
									'Return original array
									fFrame_ObjectRepoIndex = aCommandOcc
							else
									'Return object filtered array
									fFrame_ObjectRepoIndex = aConfirmed
							End If
					else
							'Only 1 object matched by name
							fFrame_ObjectRepoIndex = aCommandOcc
					End If
	'Perf Timer
	'fFrame_EndPerfTimer "Object Repository Index [" & sCommandObject & "]"
end function


Public Function fFrame_LoadAllObjectsProperties(byref oRoot) 
	'Perf Timer
	'fFrame_StartPerfTimer "Load Object Repository", "AAF", 2

'The following function recursively enumerates all the test objects directly under 
'a specified parent object. For each test object, a message box opens containing the 
'test object's name, properties, and property values. 
Dim oCollection,oPropertiesCollection,oProperty,oTestObject
Dim iObjectLoop,iPropertiesLoop,iORIndex,iIndexNameLoop
Dim sIndexName


	Set oCollection = oRepository.GetChildren(oRoot) 

	For iObjectLoop = 0 To oCollection.Count - 1 
			Set oTestObject = oCollection.Item(iObjectLoop) 

			'Set the depth level
			aORDepth(aORDepth(0)) = oRepository.GetLogicalName(oTestObject)

			'Generate Index name
			sIndexName = ""
			For iIndexNameLoop = 1 to aORDepth(0)
					If len(sIndexName) = 0 Then
							sIndexName = "~" & aORDepth(iIndexNameLoop)		
					else
							sIndexName = sIndexName & "::" & aORDepth(iIndexNameLoop)
					End If
			Next
			sIndexName = sIndexName & "#" 

			'Get index into aOR
			If instr(1,aOR(0,0),sIndexName) = 0 Then
					' Object not found in aOR, so create new reference
					'Update array size
					redim preserve aOR(10,ubound(aOR,2)+1)
	
					'Set Virtual ref
					aOR(0,0) = aOR(0,0) & sIndexName
					'Set next spare lcoation
					aOR(1,0) = aOR(1,0) + 1
					iORIndex = aOR(1,0)
			else
					'Get index into aOR array
					iORIndex = ubound(split(mid(aOR(0,0),1,instr(1,aOR(0,0),sIndexName) ),"~"))
			end if



			'Load common values
  			aOR(1,iORIndex) = aORDepth(aORDepth(0))


			'Load the Parents into the array
			For iIndexNameLoop = 1 to aORDepth(0)
					aOR(5+iIndexNameLoop ,iORIndex) = aORDepth(iIndexNameLoop)
			Next

			aORDepth(0) = aORDepth(0) + 1

			'Stack error handling
			If  aORDepth(0) > 5 Then
					reporter.ReportEvent micFail,"LoadAllObjectsProperties" ,"LoadAllObjectsProperties : Fatal Error, Stack handling error within the Object Repository loading"
					ExitTest
			End If


            Set oPropertiesCollection = oTestObject.GetTOProperties() 
            For iPropertiesLoop = 0 To oPropertiesCollection.Count - 1 
					Set oProperty = oPropertiesCollection.Item(iPropertiesLoop) 

					'Get Object properties
					select case  (oProperty.Name)
							Case "to_class","class","nativeclass" 
									aOR(2,iORIndex) = oProperty.Value
							Case "label","title","text","attached text"
									 aOR(3,iORIndex) = oProperty.Value
							Case "guistack" 
									aOR(4,iORIndex) = oProperty.Value
							Case "next" 
									aOR(5,iORIndex) = oProperty.Value
					end select
            Next 


			'Check if the current object has any children. If so then get the properties
			If  oRepository.GetChildren(oTestObject).count > 0  Then call fFrame_LoadAllObjectsProperties (oTestObject) 
			aORDepth(0) = aORDepth(0) - 1
			aORDepth(aORDepth(0)) = ""
	Next 

	'Perf Timer
	'fFrame_EndPerfTimer "Load Object Repository"
	
End Function


Function fFrame_BuildCall (byref sCommandObject)

	Dim iCommandLoop, iStackLoop,iComLoop,iObjectStatus
	Dim sGUIStack,sCall,sType,sIndexSearch
	Dim aCommandOcc(),aPLookUp
	Dim bObjectStatus

	ReDim aCommandOcc(0)
	
	'Generate Function call
	sCall = ""
	sIndexSearch = ""
	'Loop through each parent level for a QTP command upto the current max of 5
	For iComLoop = 6 to 10
			If len(aOR(iComLoop, aORLookUp(1))) > 0 Then
					'Look up the P1 to P5 parents for the command object
					aPLookUp= fFrame_ObjectRepoIndex("~" & sIndexSearch & aOR(iComLoop, aORLookUp(1)) & "#")
					If aPLookup(0) = 1 Then
							sType = aOR(2,aPLookUp(1))		
'							select case aOR(2,aPLookUp(1))
'									Case "JavaWindow"
'											 sCall = "" 
'									Case "JavaDialog"
'											 sCall = "" 
'									Case "JavaList"
'											 sCall = "" 
'									Case "JavaButton"
'											 sCall = "" 
'							end select

							'Java Objects all seem to follow the same naming convention. eg the class is the calling structure
							If len(sCall) = 0 then
									sCall = aOR(2,aPLookUp(1)) & "(" & chr(34) & aOR(1,aPLookUp(1)) & chr(34) & ")"
							else
									sCall = sCall & "." & aOR(2,aPLookUp(1)) & "(" & chr(34) & aOR(1,aPLookUp(1)) & chr(34) & ")"
							end if 
							sIndexSearch = sIndexSearch & aOR(iComLoop, aORLookUp(1)) & "::"

							'If Menu item is being used then only use the main app, as the rest of the command line isn't needed.
							'This is confirmed for. DCS, AdminGUI and CM.							
							'V4.0 - Update the check for a menu object to check the object class rather than the first 3 letters of the object name (mnu), as this did not work where the object name started with opt
							'If  (aOR(6, aORLookUp(1)) = "wndAltéaDepartureControl" or aOR(6, aORLookUp(1)) = "wndAltéaAdministration" or aOR(6, aORLookUp(1)) = "wndCMAltéaDepartureControl") and (instr(1,ucase(aOR(2, aORLookUp(1))),"MENU") > 0) Then	'V4.0
							If  (instr(1,ucase(aOR(2, aORLookUp(1))),"MENU") > 0) Then	'V4.0
									Exit for
							End If

					else
							'Error message that object cannot be found within the OR
							reporter.ReportEvent micFail,"BuildCall" ,"BuildCall [" & sCommandObject & "] Fatal Error,  P" & iComLoop-5 & " parent that is missing/duplicated within the OR"
							ExitTest					 
					End If


					'Search for the object description within the aOR

					'Use the found Object's type to generate the next part of the command structure

			End If

	Next


	'set error handling off.... as there might not be a pop up box.
	If ucase(mid(sCommandObject,1,3)) <> "OPT"  Then 'Check if this is an optional menu, typically from a right mouse click. If it is then ignore check for existing, as it will neve exist.
		'Check if expected  message box appears
		On error resume next
		Execute "iObjectStatus = " & sCall & ".getroproperty(" & chr(34) & "enabled" & chr(34) & ")"
		If  iObjectStatus = 0 Then
				'Update Results with warning
				reporter.ReportEvent micWarning,"BuildCall" ,"BuildCall [" & sCommandObject & "] WARNING,  The expected object was not enabled. This is likely due to Message box being displayed"
	
				'Set default timeout
				Setting("DefaultTimeOut")= 100
	
				'No pop up window expected to be open ... so process as an unexpected window.
				If javawindow("wndAltéaDepartureControl").JavaDialog("wndGenericMessageBox").JavaButton("btnCancel").Exist = true then
						 javawindow("wndAltéaDepartureControl").JavaDialog("wndGenericMessageBox").JavaButton("btnCancel").Click
						reporter.ReportEvent micWarning,"BuildCall" ,"BuildCall [" & sCommandObject & "] WARNING,  The expected object was not enabled. Message box was closed using 'Cancel' "
				elseIf javawindow("wndAltéaDepartureControl").JavaDialog("wndGenericMessageBox").JavaButton("btnOk").Exist = true  then
						 javawindow("wndAltéaDepartureControl").JavaDialog("wndGenericMessageBox").JavaButton("btnOk").Click
						reporter.ReportEvent micWarning,"BuildCall" ,"BuildCall [" & sCommandObject & "] WARNING,  The expected object was not enabled. Message box was closed using 'Ok' "
				end if
	
				'Set default timeout
				Setting("DefaultTimeOut")= 20000
		End If
	
		'Reset the error checking
		On error goto 0
	End If


	'Return the built call string
	 fFrame_BuildCall = sCall

End Function
