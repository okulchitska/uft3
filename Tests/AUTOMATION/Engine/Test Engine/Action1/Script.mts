
Option Explicit
'################################################################
'#
'#		Main Processing Engine for the Amadeus Automation Framework
'#      V6.1.0
'################################################################




	'Set the report to capture all events
	reporter.Filter = rfEnableAll


	'Set default timeoutf
	Setting("DefaultTimeOut")= 20000


	'Declarations Variables
	Dim iSyncTime : iSyncTime = 180
	Dim iGUIOverloadSyncTime : iGUIOverloadSyncTime = 5   
	Dim iGUIOverloadStablityTime : iGUIOverloadStablityTime = 3   									'V6.1 - Add default stability time
	Dim iGUIOverloadSyncTimeDefault : iGUIOverloadSyncTimeDefault = 5     'V5.5
	Dim iGUIOverloadStabilityTimeDefault : iGUIOverloadStabilityTimeDefault = 3						'V6.1 - Add default stability time
	
	Dim gCallStack : gCallStack = "" 			'V5.7
	Public gUseLiteralValues : gUseLiteralValues = true 'V5.7
	
	Dim gVDDPHeader   	'V6.0
	Dim gDDPParams		'V6.03 - Global var used for checking of duplicate VDDP parameters


	Dim pOR
	Dim oRepository, oTest
	Dim aLoadOR
	Dim aORNav
	Dim aSKNav
	Dim iLoadOR
	Dim bRunningFromQC : bRunningFromQC = true
	Dim iParamLoop
	Dim iSpaceLoop
	Dim iExecLoop,iPreReqCommandCheck
	Dim iParamReTries, iMaxParamReTries : iMaxParamReTries = 20
	Dim bReLoop : bReLoop = false
	Dim oTestCasePlanAttachmentF,oTestCasePlanAttachmentList

	Dim gRunType, gScreenshot
	Dim oQCConn,oTestSetF,oTestSetObj,oTestF,oTestSet,oTestSetAttachmentF,oTestSetAttachmentList,oCurrentTSTest
	Dim oTestCaseF,oTestCaseObj, oStepF,oStepList,oFileSystem, oTestCase,oTestSetTestCaseObj
	Dim oTestCaseAttachmentF,oTestCaseAttachmentList, oParamValueF, oParamList,oParam,oCurrentTestSet 
	Dim iDDPRow, iCurr,iTestCaseAttachmentLoop, iTestSetID, iTestSetAttachmentLoop, iNoSheets,iStepLoop,iTestCaseCount
	Dim iTestCaseId, iCurrentSteps, iStepOffset
	Dim sTestCaseAttachmentType,sTestSetAttachmentType

	Public bPreRun : bPreRun = True
	Public gORLookup,aORLookUp,gStack,gStackList

	Public gCapture: gCapture = true
	Public bDDPPresent : bDDPPresent = false
	Public gDDPUpdate : gDDPUpdate = True
	Public gDDPObject,  gDDPExcel, gDDPExcelWorkbook
	Public gDDPCol : gDDPCol = 0
	Public gJFEEnv : gJFEEnv = "" 'Stores the env part of the folder structure for JFE
	Public gCaptureLogFiles : gCaptureLogFiles = true
	Public gStep,gAction,gDDPRow																		'v2.1
	Public gControlCodes																				'V4.5
	Public gALMProject,gALMUsername 																	'V5.3
	Public gReportObject, gReportParam
	Public goPONStart,goPONEnd,gPONEndFileName															'V6.0
	Public gDDPLocation : gDDPLocation = ""																'V6.0
	Public bPONStart,bPONEnd																			'V6.0
	
	'Debug print
	Dim bDebugPrint : 	bDebugPrint = true
	
	'Performance output level
	Dim gPerfLevel : gPerfLevel = 0

	'Logic Commands
	Dim sLogicCommands: sLogicCommands =  "~LOOP#~ENDLOOP#~ELSE#~ENDIF#"

	'Declarations Arrays
	Public aExec(), aOR (),  aGUIOverload(1,2), aParam(), aSteps(),aActions(),aActionDetails,aORDepth(5),aObjectPreFix()
	Public aGUIStack()
	Public aFunctions()  'Function Definitions
	Public aPerfTime()   'V2.9
	Public aReportDetails()	'V5.0
	
	'PON file downloads
	Public sPONDirPath: sPONDirPath = "C:\AAF\PON\"
	
	'Define values to be replaced pre and post calls
	Dim gParamReplaceString												'V5.7
	gParamReplaceString = "`" & chr(10) & chr(13) & ",();^+-"			'V5.7
	Dim gParamReplace()													'V5.7
	Dim iParamReplaceLoop												'V5.7
	ReDim gParamReplace(len(gParamReplaceString)+2)						'V5.7
	
	'Populate the aParamReplace array
	gParamReplace(0) = len(gParamReplaceString)+2						'V5.7
	gParamReplace(1) = "||'"											'V5.7
	gParamReplace(2) = "||`"											'V5.7
	For iParamReplaceLoop = 3 To gParamReplace(0)						'V5.7
		gParamReplace(iParamReplaceLoop) = mid(gParamReplaceString,iParamReplaceLoop-2,1)		'V5.7
	Next																'V5.7



	'Get ALM project and username
	gALMProject = "AAF_Beta"
	gALMUsername = qcutil.QCConnection.UserName
	
	'#################################################################################################
	Dim xxTestCaseId,xxTestSetId,xxTestCaseIdInTestSet,xxLoadDevLib
	If gALMProject = "AAF_Beta" and gALMUsername = "andy.hatchett" Then
		'AH Test Data
			xxTestCaseId = 237
			xxTestSetId = 201
			xxTestCaseIdInTestSet = 135
			xxLoadDevLib = "Functions\SR Dev.qfl"
	ElseIf gALMProject = "AAF_Beta" and gALMUsername = "stuart.richmond" Then
		'SR Test Data
			xxTestCaseId = 1369
			xxTestSetId = 1
			xxTestCaseIdInTestSet = 563
			

	End if
	'#################################################################################################


	'Constants
	'Object Repositories
'	Dim sLoadOR : sLoadOR="Object Repositories\FM.tsr,Object Repositories\CM.tsr,Object Repositories\TechGUI.tsr,Object Repositories\Inventory.tsr,Object Repositories\AdminGUI.tsr,Object Repositories\FMWeb.tsr"
	Dim sLoadOR

	'Test Vars
	Dim sTestName,iTestId

	'Setup initial Perf Timings
	ReDim aPerfTime(4,1)	
	
	
	'Check if running in Altea_dcs
	If ucase(gALMProject) = "ALTEA_DCS" then 'or  ucase(gALMProject) = "AUTOMATION" Then
		
		Dim oALMConn,oRootResourceFolder,oRootResourceFolderFactory,oRootResourceFolderFilter,oAAFFolder
		Dim oFunctionsFolder,oFunctionsFilter,oResourceFolder,oResourceFactory,oResourceFilter,oResource   
		Dim sResourceFolder,iResourceLoop


		'Open User settings connection
		Set oALMConn = qcutil.QCConnection.UserSettings
		oALMConn.Open("AAF-ReleaseManager")
		

		'Get AAF resource folder object
		Set oRootResourceFolder = QCUtil.QCConnection.QCResourceFolderFactory.Root
		Set oRootResourceFolderFactory= oRootResourceFolder.QCResourceFolderFactory
		Set oRootResourceFolderFilter = oRootResourceFolderFactory.Filter
		oRootResourceFolderFilter("RFO_NAME") = "'AAF'"
		
		'Check that the \AAF is a unique folder
		If oRootResourceFolderFilter.NewList.count = 0 Then
			reporter.ReportEvent micFail,"MainEngine","MainEngine - Fatal Error, 'AAF' resource folder does not exist."
			sFrame_ExitTest()
		elseIf oRootResourceFolderFilter.NewList.count > 1 Then
			reporter.ReportEvent micFail,"MainEngine","MainEngine - Fatal Error, Mulitple 'AAF' resource folders exist."
			sFrame_ExitTest()
		End If


		'Set the object to the \AAF folder
		Set oAAFFolder = oRootResourceFolderFilter.NewList(1)

		Call ReleaseManager ("Functions")
		Call ReleaseManager ("Object Repositories")



		
	else
		'Load libraries directly from ALM
		'Global Relative Paths
		createobject("QuickTest.Application").Folders.Add "[QualityCenter\Resources] Resources\AAF"
	
		'Write our status
		reporter.ReportEvent micDone,"Engine" ,"Engine: All Engine resources being loaded from ALM."
	
	
		'Load Function Libraries
'		If gALMProject = "AAF_Beta" then
'			Executefile xxLoadDevLib
'		End if
		Executefile "Test\TestFrame.qfl"
		Executefile "Functions\General.qfl"
		Executefile "Functions\Command.qfl"
		Executefile "Functions\CommandSup.qfl"
		
		
		'Export resource file from QC, refresh Pause.exe
		If (fFrame_QCGetResource("Object Repositories", "AAFPause.exe","c:\AAF\Object Repositories")) = 0 then 
				reporter.ReportEvent micFail,"MainEngine","MainEngine - Fatal Error, Download from QC of Pause.exe failed"
				sFrame_ExitTest()
		end if

		'Load the Object Repository Navigation array
		'Export resource file from QC
		If (fFrame_QCGetResource("Object Repositories","ORNav.xls","c:\AAF\Object Repositories")) = 0 then 
				reporter.ReportEvent micFail,"MainEngine","MainEngine - Fatal Error, Download from QC of ORNav failed"
				sFrame_ExitTest()
		end if
	
		'Default OR to be loaded
		sLoadOR="Object Repositories\FM.tsr,Object Repositories\CM.tsr,Object Repositories\TechGUI.tsr,Object Repositories\Inventory.tsr,Object Repositories\AdminGUI.tsr,Object Repositories\FMWeb.tsr,Object Repositories\AmadeusDeviceSim.tsr"
		
		
	End If



	'Import ORNav Sheet 1 from local drive into array (Object Navigation)
	If (fFrame_ExcelLoad ("c:\AAF\Object Repositories\ORNav.xls", "", aORNav,"C1",true,1,1)) = 0 then 
			reporter.ReportEvent micFail,"MainEngine","MainEngine - Fatal Error, ORNav Import failed."
			sFrame_ExitTest()
	end if


	'Import ORNav Sheet 2 from local drive into array (SendKeys Navigation)
	If (fFrame_ExcelLoad ("c:\AAF\Object Repositories\ORNav.xls", "", aSKNav,"C1",true,2,1)) = 0 then 
			reporter.ReportEvent micFail,"MainEngine","MainEngine - Fatal Error, ORNav Import failed."
			sFrame_ExitTest()
	end if


	'Load Environment Variables
	'Initialize Global Array
	call sFrame_InitializeOR
	Call sFrame_InitializeGlobalVar
	Call sCommand_FunctionDefinitions

	'Determine if Test is valid
	If  fGeneral_ExecutionFromQC = 0  then
		reporter.ReportEvent micWarning,"PRE-Check: Check 1","PRE-REQ failed due to test not being stored within QC"
		bRunningFromQC = False
	end if

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

	'Default aSteps
	ReDim aSteps(4,0)
	

	
	If bRunningFromQC Then

		'Get Run Type from QC  (TestSet or Single)
        Set oQCConn = QCUtil.QCConnection 'Set  QC connection
        gRunType = oQCConn.UserSettings.value("RunType")

		'Get current active test case in Test Set
		Set oTest  = QCUtil.CurrentTest
		sTestName = oTest.Name
		iTestId = oTest.ID

	End if
	
	'Connect to QC
	Set oQCConn = QCUtil.QCConnection 
	Set oCurrentTestSet = QCUtil.CurrentTestSet

	'Set global var for Run Type
	gRunType =   oQCConn.UserSettings.value("RunType")


	'End Perf Timer
	'fFrame_EndPerfTimer "AAF Startup"	


'****************************************
If gALMProject = "AAF_Beta" then
	gRunType = "TestCase" '"TestCase"  '"TestSet"
'	gRunType = "TestSet" '"TestCase"  '"TestSet"
End if
'****************************************
		'grab the RunType from QC to see if a complete test set is being run. Or just a single test case
		If  gRunType  = "TestSet" Then

			'Start Perf Timer
			'fFrame_StartPerfTimer "Load TestSet Attachments", "AAF", 0	

'****************************************
		If gALMProject = "AAF_Beta" then
			 iTestSetID = xxTestSetId '201 '101
		else
'****************************************
			 iTestSetID = oCurrentTestSet.ID   '101
		End if
'****************************************
			Set oTestSetF =  oQCConn.TestSetFactory
			Set oTestSetObj = oTestSetF.Item(iTestSetID)
			Set oTestF = oTestSetObj.TSTestFactory
			Set oTestSet = oTestF.NewList("")

			'TestSet attachments
			Set oTestSetAttachmentF = oTestSetObj.Attachments
			Set oTestSetAttachmentList = oTestSetAttachmentF.newlist("")

			'Loop through all attachements
			iDDPRow = 0
			For iTestSetAttachmentLoop = 1 to oTestSetAttachmentList.count
					'Check to determine if any meet the naming convention for parameters to drive the test set 
                    sTestSetAttachmentType = ucase(mid(oTestSetAttachmentList.item( iTestSetAttachmentLoop).name,instrrev(oTestSetAttachmentList.item( iTestSetAttachmentLoop).name,".")))

					'Check for PONStart
					If  ucase(mid(replace(oTestSetAttachmentList.item( iTestSetAttachmentLoop).name,"CYCLE" & "_" & iTestSetID & "_",""),1,8)) = "PONSTART" and (sTestSetAttachmentType = ".XLS" or sTestSetAttachmentType = ".XLSX" or sTestSetAttachmentType = ".XLSM") Then	'V6.0
						'Set the object for use later
						set goPONStart = oTestSetAttachmentList.item( iTestSetAttachmentLoop)																																													'V6.0
					elseIf  ucase(mid(replace(oTestSetAttachmentList.item( iTestSetAttachmentLoop).name,"CYCLE" & "_" & iTestSetID & "_",""),1,6)) = "PONEND" and (sTestSetAttachmentType = ".XLS" or sTestSetAttachmentType = ".XLSX" or sTestSetAttachmentType = ".XLSM") Then   'V6.0
						'Set the object for use later																																																							'V6.0
						set goPONEnd = oTestSetAttachmentList.item( iTestSetAttachmentLoop)																																														'V6.0
					elseIf  ucase(mid(replace(oTestSetAttachmentList.item( iTestSetAttachmentLoop).name,"CYCLE" & "_" & iTestSetID & "_",""),1,3)) = "DDP" and (sTestSetAttachmentType = ".XLS" or sTestSetAttachmentType = ".XLSX" or sTestSetAttachmentType = ".XLSM") Then
							'This is a data driven parameter list and is an excel file#
							If iDDPRow = 0 then
									set gDDPObject = oTestSetAttachmentList.item( iTestSetAttachmentLoop)
									 gDDPObject.load true, ""
	
									'Load aTestSetDPP into aParam
									iNoSheets =  fFrame_ExcelLoad(gDDPObject.filename, "", aParam, "R1",false,1,1) 
									If iNoSheets = 0 Then
											reporter.ReportEvent micFail,"MainEngine" ,"MainEngine: Fatal Error, ExcelLoad [" & gDDPObject.filename & "] - Failed"
											sFrame_ExitTest()
									End If

									iDDPRow = ubound(aParam,1) 'Number of rows of data
									bDDPPresent = true
									gDDPLocation = "Test Set"     'V6.0
									
									If  gDDPUpdate = true Then
											'Open the DDP for global access
											gDDPCol = fFrame_OpenDDP
									else
											'Delete temp attachment downloaded to client machine
											If fFrame_FileDelete (gDDPObject.filename) = 0 then
												reporter.ReportEvent micWarning,"MainEngine" ,"MainEngine: File [" & oTestSetAttachmentList.item( iTestSetAttachmentLoop).filename & "] Could not be deleted"
											end if
									end if
							else
									'Function call failed
									reporter.ReportEvent micFail,"MainEngine" ,"MainEngine: Fatal Error,TestSet Data Driven Parameter Loading. Multiple DDP_*.XLS files found attached to the TestSet. Please ensure only 1 Data Driven Parameter spreadsheet is attached to the testset." 							
									sFrame_ExitTest()
							End if
					End If
			Next

			'End Perf Timer
			'fFrame_EndPerfTimer "Load TestSet Attachments"

			if iDDPRow > 0 then 'Check if a DDP spreadsheet has been loaded from the testset
					'DDP loaded from testset, so iterate through all test cases for each row
					' of data

					'Start Perf Timer
					'fFrame_StartPerfTimer "Load Test Steps", "AAF", 0

					'Initialise aSteps
					ReDim aSteps(4,0)

					'Loop through all the test cases within the test set
					iTestCaseCount = 0
					For Each oTestCase in oTestSet
						'Increase Test case count
						iTestCaseCount = iTestCaseCount + 1

						'Steps
						Set oTestCaseF =  oQCConn.TestFactory
						Set oTestCaseObj = oTestCaseF.Item(oTestCase.TestId)
						Set oStepF = oTestCaseObj.DesignStepFactory
						Set oStepList = oStepF.NewList("")

						'Check and load Pre-Reqs
						iCurr = aSteps(1,0)
						fFrame_SpreadsheetSteps oTestCaseObj, aSteps, "PRE-REQ",oTestCase.Id,iCurr,iTestCaseCount & "-" & oTestSet.Count
					
						'Re-size for the number of steps
						iCurr = aSteps(1,0)
						ReDim preserve aSteps(4, iCurr + oStepList.Count)

						'Load Steps
						iStepOffset = 0
						For iStepLoop = 1 to oStepList.Count

							'Set offset to 1 if the step order starts at 0
							If iStepLoop = 1 and oStepList(iStepLoop).order = 0 Then	'V2.1 10/3/2015
							    iStepOffset = 1											'V2.1 10/3/2015
							End If                             							'V2.1 10/3/2015

							'Set Virtual ref
							aSteps(0,0) =aSteps(0,0) & "~"  &  fGeneral_ClearHTML(oStepList(iStepLoop).StepName) & "#"
							'Set next spare location
							aSteps(1,0) = aSteps(1,0) + 1
							aSteps(0,aSteps(1,0) ) = oTestCase.Id & "~" & oStepList(iStepLoop).order + iStepOffset 'V2.1 10/3/2015
							aSteps(1,aSteps(1,0) ) = fGeneral_ClearHTML(oStepList(iStepLoop).StepDescription)
							aSteps(2,aSteps(1,0) ) = fGeneral_ClearHTML(oStepList(iStepLoop).StepExpectedResult)
							aSteps(3,aSteps(1,0) ) =  oTestCaseObj.Name 
							aSteps(4,aSteps(1,0)) = iTestCaseCount & "-" & oTestSet.Count
						Next

						'Check and load Post-Reqs
						iCurr = aSteps(1,0)
						fFrame_SpreadsheetSteps oTestCaseObj, aSteps, "POST-REQ",oTestCase.Id,iCurr,iTestCaseCount & "-" & oTestSet.Count

					Next
					
					'End Perf Timer
					'fFrame_EndPerfTimer "Load Test Steps"
					
					'Run the parse and execute
					fFrame_Run			'V6.0
			else
				'No DDP attached to the test set
				'Check if each test case has a DDP it it does then iterate through each data row for
				'that test case... then move onto the next test case.
				'If no DDP attached to the test case, then check if manual parameters have been defined

				'Loop through all the test cases within the test set
				iTestCaseCount = 0
				For Each oTestCase in oTestSet

					'Start Perf Timer
					'fFrame_StartPerfTimer "Load Test Steps", "AAF", 0	

					'Reset the Parameters per test case
					ReDim aParam(2,1 )
					bDDPPresent = false

					'Increase Test case count
					iTestCaseCount = iTestCaseCount + 1

					'Steps
					Set oTestCaseF =  oQCConn.TestFactory
					Set oTestCaseObj = oTestCaseF.Item(oTestCase.TestId)
					Set oStepF = oTestCaseObj.DesignStepFactory
					Set oStepList = oStepF.NewList("")
				
					'Re-set steps
					ReDim aSteps(4,0)

					'Check and load Pre-Reqs
					fFrame_SpreadsheetSteps oTestCaseObj, aSteps, "PRE-REQ",oTestCase.Id,0, iTestCaseCount & "-" & oTestSet.Count
				
					'Re-size for the number of params
					iCurrentSteps = aSteps(1,0) 
					ReDim Preserve aSteps(4,iCurrentSteps + oStepList.Count)

					'Load Steps
					iStepOffset = 0
					For iStepLoop = 1 to oStepList.Count
							'Set offset to 1 if the step order starts at 0
							If iStepLoop = 1 and oStepList(iStepLoop).order = 0 Then	'V2.1 10/3/2015
							    iStepOffset = 1											'V2.1 10/3/2015
							End If                             							'V2.1 10/3/2015


							'Set Virtual ref
							aSteps(0,0) =aSteps(0,0) & "~"  &  fGeneral_ClearHTML(oStepList(iStepLoop).StepName) & "#"
							'Set next spare location
							aSteps(1,0) = aSteps(1,0) + 1
							aSteps(0,aSteps(1,0) ) = oTestCase.Id & "~" & oStepList(iStepLoop).order + iStepOffset 'V2.1 10/3/2015
							aSteps(1,aSteps(1,0) ) = fGeneral_ClearHTML(oStepList(iStepLoop).StepDescription)
							aSteps(2,aSteps(1,0) ) = fGeneral_ClearHTML(oStepList(iStepLoop).StepExpectedResult)
							aSteps(3,aSteps(1,0) ) =  oTestCaseObj.Name 
							aSteps(4,aSteps(1,0)) = iTestCaseCount & "-" & oTestSet.Count
					Next
					
					'Check and load Post-Reqs
					iCurrentSteps = aSteps(1,0) 
					fFrame_SpreadsheetSteps oTestCaseObj, aSteps, "POST-REQ",oTestCase.Id,iCurrentSteps, iTestCaseCount & "-" & oTestSet.Count
					
					'Attachments
					'Connect to test in test set attachment factory
					set oTestCaseAttachmentF = oTestCase.Attachments    'oTestCaseObj
					Set oTestCaseAttachmentList = oTestCaseAttachmentF.newlist("")
					
					
					'End Perf Timer
					'fFrame_EndPerfTimer "Load Test Steps"
					
					'Start Perf Timer
					'fFrame_StartPerfTimer "Load Attachments [Test Case in TestLab]", "AAF", 0	

					'Loop through all attachments attached to the test case WITHIN the testlab
					iNoSheets = 0
					For iTestCaseAttachmentLoop = 1 to oTestCaseAttachmentList.count
							'Check to determine if any meet the naming convention for parameters to drive the test set 
							sTestCaseAttachmentType = ucase(mid(oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,instrrev(oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,".")))
							
							If isempty(goPONStart) or isempty(goPONEnd) Then						'V6.0
								'Check for PONStart
								If  ucase(mid(oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,instr(instr(1,oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name ,"_")+1,oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,"_")+1,8))  = "PONSTART" and (sTestCaseAttachmentType = ".XLS" or sTestCaseAttachmentType= ".XLSX" or sTestCaseAttachmentType = ".XLSM") Then	'V6.0
									'Set the object for use later
									set goPONStart = oTestCaseAttachmentList.item(iTestCaseAttachmentLoop)																																													'V6.0                      																																																								'V6.0
								'Check for PONEnd
								elseIf  ucase(mid(oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,instr(instr(1,oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name ,"_")+1,oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,"_")+1,6))  = "PONEND" and (sTestCaseAttachmentType = ".XLS" or sTestCaseAttachmentType = ".XLSX" or sTestCaseAttachmentType = ".XLSM") Then   'V6.0
									'Set the object for use later																																																							'V6.0
									set goPONEnd = oTestCaseAttachmentList.item(iTestCaseAttachmentLoop)									'V6.0
								End if            																							'V6.0
							End if                  																						'V6.0
							
							If ucase(mid(oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,instr(instr(1,oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name ,"_")+1,oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,"_")+1,3))  = "DDP"  and (sTestCaseAttachmentType = ".XLS" or sTestCaseAttachmentType = ".XLSX" or sTestCaseAttachmentType = ".XLSM") Then

								if iNoSheets = 0 then
										'This is a data driven parameter list and is an excel file#
										Set gDDPObject = oTestCaseAttachmentList.item( iTestCaseAttachmentLoop)
										gDDPObject.load true, ""
										iNoSheets =  fFrame_ExcelLoad(gDDPObject.filename, "", aParam, "R1",false,1,1)
										If iNoSheets = 0 Then
												reporter.ReportEvent micFail,"MainEngine" ,"MainEngine: Fatal Error, ExcelLoad [" & gDDPObject.filename & "] - Failed"
												sFrame_ExitTest()
										End If
										bDDPPresent = true
										gDDPLocation = "Test Case within Test set"  'V6.0
		
										If  gDDPUpdate = true Then
												'Open the DDP for global access
												gDDPCol = fFrame_OpenDDP
										else
												'Delete temp attachement downloaded to client machine
												If fFrame_FileDelete (gDDPObject.filename) = 0 then
													reporter.ReportEvent micWarning,"MainEngine" ,"MainEngine: File [" & gDDPObject.filename & "] Could not be deleted"
												end if
										end if
								else
									'Function call failed
									reporter.ReportEvent micFail,"MainEngine","MainEngine - Fatal Error, TestSet Data Driven Parameter Loading Multiple DDP_*.XLS files found attached to the Test Case within the test lab. Please ensure only 1 Data Driven Parameter spreadsheet is attached to the testset."
									sFrame_ExitTest()
								end if
							End If
					Next
					
					'End Perf Timer
					'fFrame_EndPerfTimer "Load Attachments [Test Case in TestLab]"

					'Start Perf Timer
					'fFrame_StartPerfTimer "Load Attachments [Test Case in TestPlan]", "AAF", 0	

					'Connect to test in test plan attachment factory
					set oTestCasePlanAttachmentF = oTestCaseObj.Attachments    
					Set oTestCasePlanAttachmentList = oTestCasePlanAttachmentF.newlist("")

					'Loop through all attachments attached to the test case WITHIN the test plan
					For iTestCaseAttachmentLoop = 1 to oTestCasePlanAttachmentList.count
							'Check to determine if any meet the naming convention for parameters to drive the test set 
							sTestCaseAttachmentType = ucase(mid(oTestCasePlanAttachmentList.item( iTestCaseAttachmentLoop).name,instrrev(oTestCasePlanAttachmentList.item( iTestCaseAttachmentLoop).name,".")))
							
							
							If isempty(goPONStart) or isempty(goPONEnd) Then						'V6.0
								'Check for PONStart
								If  ucase(mid(oTestCasePlanAttachmentList.item( iTestCaseAttachmentLoop).name,instr(instr(1,oTestCasePlanAttachmentList.item( iTestCaseAttachmentLoop).name ,"_")+1,oTestCasePlanAttachmentList.item( iTestCaseAttachmentLoop).name,"_")+1,8))  = "PONSTART" and (sTestCaseAttachmentType = ".XLS" or sTestCaseAttachmentType= ".XLSX" or sTestCaseAttachmentType = ".XLSM") Then	'V6.0
									'Set the object for use later
									set goPONStart = oTestCasePlanAttachmentList.item( iTestCaseAttachmentLoop)																																													'V6.0                      																																																								'V6.0
								'Check for PONEnd
								elseIf  ucase(mid(oTestCasePlanAttachmentList.item( iTestCaseAttachmentLoop).name,instr(instr(1,oTestCasePlanAttachmentList.item( iTestCaseAttachmentLoop).name ,"_")+1,oTestCasePlanAttachmentList.item( iTestCaseAttachmentLoop).name,"_")+1,6))  = "PONEND" and (sTestCaseAttachmentType = ".XLS" or sTestCaseAttachmentType = ".XLSX" or sTestCaseAttachmentType = ".XLSM") Then   'V6.0
									'Set the object for use later																																																							'V6.0
									set goPONEnd = oTestCasePlanAttachmentList.item( iTestCaseAttachmentLoop)									'V6.0
								End if            																							'V6.0
							End if                  																						'V6.0
							
							
							
							
							
							
							If ucase(mid(oTestCasePlanAttachmentList.item( iTestCaseAttachmentLoop).name,instr(instr(1,oTestCasePlanAttachmentList.item( iTestCaseAttachmentLoop).name ,"_")+1,oTestCasePlanAttachmentList.item( iTestCaseAttachmentLoop).name,"_")+1,3))  = "DDP"  and (sTestCaseAttachmentType = ".XLS" or sTestCaseAttachmentType = ".XLSX" or sTestCaseAttachmentType = ".XLSM") Then

									if iNoSheets = 0 then
										'This is a data driven parameter list and is an excel file#
										Set gDDPObject = oTestCasePlanAttachmentList.item( iTestCaseAttachmentLoop)
										gDDPObject.load true, ""
										iNoSheets =  fFrame_ExcelLoad(gDDPObject.filename, "", aParam, "R1",false,1,1)
										If iNoSheets = 0 Then
												reporter.ReportEvent micFail,"MainEngine" ,"MainEngine: Fatal Error, ExcelLoad [" & gDDPObject.filename & "] - Failed"
												sFrame_ExitTest()
										End If
										bDDPPresent = true
										gDDPLocation = "Test Case" 	'V6.0

			
										If  gDDPUpdate = true Then
												'Open the DDP for global access
												gDDPCol = fFrame_OpenDDP
										else
												'Delete temp attachement downloaded to client machine
												if fFrame_FileDelete (gDDPObject.filename) = 0 then
													reporter.ReportEvent micWarning,"MainEngine" ,"MainEngine: File [" & gDDPObject.filename & "] Could not be deleted"
												end if
										end if
								else
										reporter.ReportEvent micFail,"MainEngine" ,"MainEngine: Fatal Error,TestSet Data Driven Parameter Loading. Multiple DDP_*.XLS files found attached to the TestCase within Test Plan. Please ensure only 1 Data Driven Parameter spreadsheet is attached to the testcase." 																	
										sFrame_ExitTest()
								end if
							End If
					Next

					'End Perf Timer
					'fFrame_EndPerfTimer "Load Attachments [Test Case in TestPlan]"



					'Params
					'If no DDP spreadsheets loaded from Test case then check the manual parameters
					if iNoSheets = 0 then 
						'Start Perf Timer
						'fFrame_StartPerfTimer "Load Attachments [Manual Parameters]", "AAF", 0

						'Check manual parameters
						'Connect to Paramter Factory
						Set oParamValueF = oTestCase.ParameterValueFactory
						Set oParamList =  oParamValueF.newlist("")
	
						'If the test case has parameters
						If  oParamList.Count > 0 Then

								'Size the parameter  storage array
								ReDim aParam(2,oParamList.Count )
								gDDPLocation = "Manual Parameters"   'V6.0
			
								For each oParam in oParamList
										If bDebugPrint Then : print oParam.Name
										If bDebugPrint Then : print oParam.DefaultValue
										If bDebugPrint Then : print oParam.ActualValue

										'Set Virtual ref
										aParam(0,0) =aParam(0,0) & "~"  &  fGeneral_ClearHTML( ucase(Trim(oParam.Name))) & "#"
	
										'Set next spare lcoation
										aParam(1,0) = aParam(1,0) + 1
										aParam(1,aParam(1,0)) =  fGeneral_ClearHTML(Trim(oParam.Name)) 
										aParam(2,aParam(1,0)) = fGeneral_ClearHTML(Trim(oParam.ActualValue) )
								Next
						else
								'Size the parameter  storage array
								ReDim aParam(2,1) 					'V4.0

								'No Parameters set count to 0
								aParam(1,0) = 0
						end if
						
						
						'End Perf Timer
						'fFrame_EndPerfTimer "Load Attachments [Manual Parameters]"

					end if
					
					'Run the parse and execute
					fFrame_Run			'V6.0

				Next
			end if
		else
			'Running individual test cases

			'Start Perf Timer
			'fFrame_StartPerfTimer "Load Test Steps", "AAF", 0	

			'Initialise aSteps
			ReDim aSteps(4,0)

			'Load Steps
			'Get current active test case in Test Set
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If gALMProject = "AAF_Beta" then
	Set oTestCaseF =  oQCConn.TestFactory
	Set oTestCaseObj = oTestCaseF.Item(xxTestCaseId)  '237 '109   '216	'160
	Set oStepList  = oTestCaseObj.DesignStepFactory.NewList("")
	iTestCaseId = xxTestCaseIdInTestSet  '117  '126
	fFrame_SpreadsheetSteps oTestCaseObj, aSteps, "PRE-REQ",iTestCaseId,0,"1-1"
else
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			Set oStepList  = QCUtil.CurrentTest.DesignStepFactory.NewList("")
			iTestCaseId = qcutil.CurrentTestSetTest.ID
			fFrame_SpreadsheetSteps qcutil.CurrentTest, aSteps, "PRE-REQ",iTestCaseId,0,"1-1"
End if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
			'Re-size for the number of params
			iCurrentSteps = aSteps(1,0) 
			ReDim Preserve aSteps(4,iCurrentSteps + oStepList.Count)
		
			'Load Steps
			iStepOffset = 0
			For iStepLoop = 1 to oStepList.Count
				'Set offset to 1 if the step order starts at 0
				If iStepLoop = 1 and oStepList(iStepLoop).order = 0 Then	'V2.1 10/3/2015
				    iStepOffset = 1											'V2.1 10/3/2015
				End If                             							'V2.1 10/3/2015
							
							
							
				'Set Virtual ref
				aSteps(0,0) =aSteps(0,0) & "~"  &  fGeneral_ClearHTML(oStepList(iStepLoop).StepName) & "#"
				'Set next spare lcoation
				aSteps(1,0) = aSteps(1,0) + 1
				aSteps(0, aSteps(1,0)) = iTestCaseId & "~" & oStepList(iStepLoop).order + iStepOffset 'V2.1 10/3/2015
				aSteps(1,aSteps(1,0)) = fGeneral_ClearHTML(oStepList(iStepLoop).StepDescription)
				aSteps(2,aSteps(1,0)) = fGeneral_ClearHTML(oStepList(iStepLoop).StepExpectedResult)
				aSteps(3,aSteps(1,0) ) =  QCUtil.CurrentTest.Name 
                aSteps(4,aSteps(1,0)) = "1-1"
			Next

			'Check and load Post-Reqs
			iCurrentSteps = aSteps(1,0)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			If gALMProject = "AAF_Beta" then
				fFrame_SpreadsheetSteps oTestCaseObj, aSteps, "POST-REQ",iTestCaseId,iCurrentSteps,"1-1"
			else
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				fFrame_SpreadsheetSteps qcutil.CurrentTest, aSteps, "POST-REQ",iTestCaseId,iCurrentSteps,"1-1"       
			End if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

			'End Perf Timer
			'fFrame_EndPerfTimer "Load Test Steps"

			'Start Perf Timer
			'fFrame_StartPerfTimer "Load Attachments [Test Case in TestLab]", "AAF", 0	
					
			'Attachments
			'Connect to attachment factory
'#Get both the attachments to the test plan and the test case in the test set
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If gALMProject = "AAF_Beta" then
	Set oTestSetF =  oQCConn.TestSetFactory
	Set oTestSetObj = oTestSetF.Item(xxTestSetId)
	Set oTestF = oTestSetObj.TSTestFactory
	Set oTestSetTestCaseObj = oTestF.item(xxTestCaseIdInTestSet) '117  '126
	
	set oTestCaseAttachmentF = oTestSetTestCaseObj.Attachments 'check this works !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
	Set oTestCaseAttachmentList = oTestCaseAttachmentF.newlist("")
Else
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		set oTestCaseAttachmentF = QCUtil.CurrentTestSetTest.Attachments 
		Set oTestCaseAttachmentList = oTestCaseAttachmentF.newlist("")
End if	
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

 			
			'Loop through all attachments to the test case in the test set
			iNoSheets = 0
			For iTestCaseAttachmentLoop = 1 to oTestCaseAttachmentList.count
					'Check to determine if any meet the naming convention for parameters to drive the test set 
					sTestCaseAttachmentType = ucase(mid(oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,instrrev(oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,".")))
					
					
					If isempty(goPONStart) or isempty(goPONEnd) Then						'V6.0
						'Check for PONStart
						If  ucase(mid(oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,instr(instr(1,oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name ,"_")+1,oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,"_")+1,8))  = "PONSTART" and (sTestCaseAttachmentType = ".XLS" or sTestCaseAttachmentType= ".XLSX" or sTestCaseAttachmentType = ".XLSM") Then	'V6.0
							'Set the object for use later
							set goPONStart = oTestCaseAttachmentList.item( iTestCaseAttachmentLoop)																																													'V6.0                      																																																								'V6.0
						'Check for PONEnd
						elseIf  ucase(mid(oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,instr(instr(1,oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name ,"_")+1,oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,"_")+1,6))  = "PONEND" and (sTestCaseAttachmentType = ".XLS" or sTestCaseAttachmentType = ".XLSX" or sTestCaseAttachmentType = ".XLSM") Then   'V6.0
							'Set the object for use later																																																							'V6.0
							set goPONEnd = oTestCaseAttachmentList.item( iTestCaseAttachmentLoop)									'V6.0
						End if            																							'V6.0
					End if                  																						'V6.0


					If ucase(mid(oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,instr(instr(1,oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name ,"_")+1,oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,"_")+1,3))  = "DDP"  and (sTestCaseAttachmentType = ".XLS" or sTestCaseAttachmentType = ".XLSX" or sTestCaseAttachmentType = ".XLSM") Then
							if iNoSheets = 0 then
								'This is a data driven parameter list and is an excel file#
								Set gDDPObject = oTestCaseAttachmentList.item( iTestCaseAttachmentLoop)
								gDDPObject.load true, ""
								iNoSheets =  fFrame_ExcelLoad(gDDPObject.filename, "", aParam, "R1",false,1,1)
								If iNoSheets = 0 Then
										reporter.ReportEvent micFail,"MainEngine" ,"MainEngine: Fatal Error, ExcelLoad [" & gDDPObject.filename & "] - Failed"
										sFrame_ExitTest()
								End If
								bDDPPresent = true
								gDDPLocation = "Test Case within Test set"			'V6.0

								If  gDDPUpdate = true Then
										'Open the DDP for global access
										gDDPCol = fFrame_OpenDDP
								else
										'No DDP updates required, so delete the local file
										'Delete temp attachement downloaded to client machine
										If fFrame_FileDelete (gDDPObject.filename) = 0 then
											reporter.ReportEvent micWarning,"MainEngine" ,"MainEngine: File [" & gDDPObject.filename & "] Could not be deleted"
										end if
								end if

						else
								reporter.ReportEvent micFail,"MainEngine" ,"MainEngine: Fatal Error,TestSet Data Driven Parameter Loading. Multiple DDP_*.XLS files found attached to the TestCase in test lab. Please ensure only 1 Data Driven Parameter spreadsheet is attached to the testcase." 															
								sFrame_ExitTest()
						end if
					End If
			Next


			'End Perf Timer
			'fFrame_EndPerfTimer "Load Attachments [Test Case in TestLab]"
					
			'Start Perf Timer
			'fFrame_StartPerfTimer "Load Attachments [Test Case in TestPlan]", "AAF", 0	
			
			set oTestCaseAttachmentF = nothing
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If gALMProject = "AAF_Beta" then
	set oTestCaseAttachmentF = oTestCaseObj.Attachments 'check this works !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
	Set oTestCaseAttachmentList = oTestCaseAttachmentF.newlist("")
else
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Set oTestCaseAttachmentF = qcutil.CurrentTest.Attachments
	Set oTestCaseAttachmentList = oTestCaseAttachmentF.newlist("")
End if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
			'Loop through all attachments to the test case in the test plan
			For iTestCaseAttachmentLoop = 1 to oTestCaseAttachmentList.count
					'Check to determine if any meet the naming convention for parameters to drive the test set 
					sTestCaseAttachmentType = ucase(mid(oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,instrrev(oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,".")))
					
					
					If isempty(goPONStart) or isempty(goPONEnd) Then						'V6.0
						'Check for PONStart
						If  ucase(mid(oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,instr(instr(1,oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name ,"_")+1,oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,"_")+1,8))  = "PONSTART" and (sTestCaseAttachmentType = ".XLS" or sTestCaseAttachmentType= ".XLSX" or sTestCaseAttachmentType = ".XLSM") Then	'V6.0
							'Set the object for use later
							set goPONStart = oTestCaseAttachmentList.item( iTestCaseAttachmentLoop)																																													'V6.0                      																																																								'V6.0
						'Check for PONEnd
						elseIf  ucase(mid(oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,instr(instr(1,oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name ,"_")+1,oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,"_")+1,6))  = "PONEND" and (sTestCaseAttachmentType = ".XLS" or sTestCaseAttachmentType = ".XLSX" or sTestCaseAttachmentType = ".XLSM") Then   'V6.0
							'Set the object for use later																																																							'V6.0
							set goPONEnd = oTestCaseAttachmentList.item( iTestCaseAttachmentLoop)									'V6.0
						End if            																							'V6.0
					End if                  																						'V6.0
									
					
					
					
					
					
					
					If ucase(mid(oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,instr(instr(1,oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name ,"_")+1,oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,"_")+1,3))  = "DDP"  and (sTestCaseAttachmentType = ".XLS" or sTestCaseAttachmentType = ".XLSX" or sTestCaseAttachmentType = ".XLSM") Then
						if iNoSheets = 0 then
							'This is a data driven parameter list and is an excel file#
							Set gDDPObject = oTestCaseAttachmentList.item( iTestCaseAttachmentLoop)
							gDDPObject.load true, ""
							iNoSheets =  fFrame_ExcelLoad(gDDPObject.filename, "", aParam, "R1",false,1,1)
							If iNoSheets = 0 Then
									reporter.ReportEvent micFail,"MainEngine" ,"MainEngine: Fatal Error, ExcelLoad [" & gDDPObject.filename & "] - Failed"
									sFrame_ExitTest()
							End If
							bDDPPresent = true
							gDDPLocation = "Test Case"			'V6.0
							
								If  gDDPUpdate = true Then
										'Open the DDP for global access
										gDDPCol = fFrame_OpenDDP
								else
										'Delete temp attachement downloaded to client machine
										If fFrame_FileDelete (gDDPObject.filename) = 0 then
												reporter.ReportEvent micWarning,"MainEngine" ,"MainEngine: File [" & gDDPObject.filename & "] Could not be deleted"
										end if
								end if
						else
							reporter.ReportEvent micFail,"MainEngine" ,"MainEngine: Fatal Error,TestSet Data Driven Parameter Loading. Multiple DDP_*.XLS files found attached to the TestCase in test plan. Please ensure only 1 Data Driven Parameter spreadsheet is attached to the testcase." 							
							sFrame_ExitTest()
						end if
					End If
			Next

			'End Perf Timer
			'fFrame_EndPerfTimer "Load Attachments [Test Case in TestPlan]"
					
			'Params
			'If no DDP spreadsheets loaded from Test case then check the manual parameters
			if iNoSheets = 0 then 
				'Start Perf Timer
				'fFrame_StartPerfTimer "Load Attachments [Manual Parameters]", "AAF", 0

				'Check manual parameters
				'Connect to Paramter Factory

			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			If gALMProject = "AAF_Beta" then
				Set oCurrentTSTest= oTestSetTestCaseObj
			else
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				Set oCurrentTSTest= QCUtil.CurrentTestSetTest
			End if
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				If oCurrentTSTest.HasSteps Then

					'If the test case has parameters
					If  oCurrentTSTest.Params.Count > 0 Then

						'Re-size for the number of params
						ReDim aParam(2,oCurrentTSTest.Params.Count)
						gDDPLocation = "Manual Parameters"				'V6.0
						
						'Load
						For iParamLoop = 0  to oCurrentTSTest.Params.Count-1
								'Set Virtual ref
								aParam(0,0) =aParam(0,0) & "~"  &   fGeneral_ClearHTML(ucase(Trim(oCurrentTSTest.Params.ParamName(iParamLoop)))) & "#"
								'Set next spare lcoation
								aParam(1,0) = aParam(1,0) + 1
								aParam(1, iParamLoop+1) =  fGeneral_ClearHTML(Trim(oCurrentTSTest.Params.ParamName(iParamLoop)) )
								aParam(2, iParamLoop+1) =  fGeneral_ClearHTML(Trim(oCurrentTSTest.Params.ParamValue(iParamLoop))) 
						Next
					else
						'Size the parameter  storage array
						ReDim aParam(2,1 )        'V4.0

						'No Parameters set count to 0
						aParam(1,0) = 0
					end if
				end if 'No test steps
				'End Perf Timer
				'fFrame_EndPerfTimer "Load Attachments [Manual Parameters]"
			end if
					
			'Run the parse and execute
			fFrame_Run			'V6.0

		End if


		'Start Perf Timer
		'fFrame_StartPerfTimer "Update DDP to ALM","AAF",0

		'Save & Close DDP
		sFrame_CloseDDP

		'End Perf Timer
		'fFrame_EndPerfTimer "Update DDP to ALM"


		'Write out Perf Timings
		If ubound(aPerfTime,2) > 0 Then
			fFrame_WritePerfTiming()
		End if


'#####################################################################
'# Function    ReleaseManager
'#
'# description:	 Controls the downloading of resources from ALM.
'#					These include the Object Repositories, Function libraries,
'#					AAF-Pause, ORNav
'#
'# inputs:		Resource folder in ALM to be processed
'#					  
'# return value:	N/A
'#
'# author:		Stuart Richmond
'#
'# date written: 01/02/2016 - Engine V5.3
'#
'#####################################################################
'#	Change History
'#	Date		Author			Version			Change Made
'#  9Nov16		S Richmond		V6.1.0			AAF-577 : Update ReleaseManager to only write out to log when ALM version is out of date
'#####################################################################
 Function ReleaseManager(byval sResourceFolder)
 	
 Dim bDownloadResource
	
	
		sLoadOR = ""
	
		'Set the object to the resource folder
		Set oFunctionsFolder= oAAFFolder.QCResourceFolderFactory
		Set oFunctionsFilter = oFunctionsFolder.Filter
		oFunctionsFilter("RFO_NAME") = "'" & sResourceFolder & "'"
		
		'Check that the \AAF\sResourceFolder is a unique folder
		If oFunctionsFilter.NewList.count = 0 Then
			reporter.ReportEvent micFail,"MainEngine","MainEngine - Fatal Error, 'AAF\" & sResourceFolder & "' resource folder does not exist."
			ExitTest()
		elseIf oFunctionsFilter.NewList.count > 1 Then
			reporter.ReportEvent micFail,"MainEngine","MainEngine - Fatal Error, Mulitple 'AAF\" & sResourceFolder & "' resource folders exist."
			ExitTest
		End If
		
		
		
		Set oResourceFolder = oFunctionsFilter.NewList(1)
		Set oResourceFactory = oResourceFolder.QCResourceFactory
		Set oResourceFilter = oResourceFactory.Filter
		
		'oResourceFilter("RSC_NAME") = "'Command.qfl'"
		Set oResource = oResourceFilter.NewList
		If oResource.Count = 0 Then
			reporter.ReportEvent micFail,"MainEngine","MainEngine - Fatal Error, 'AAF\" & sResourceFolder & "' contains no resources."
			ExitTest
		End If
		
		'Loop through all the resources in the given folder
		For iResourceLoop = 1 To oResource.Count
'			If bDebugPrint Then : print oResourceFilter.NewList(iResourceLoop).field("RSC_NAME") & "    " & oResourceFilter.NewList(iResourceLoop).field("RSC_VTS") & "    DateDiff: " & dateDiff("s",oResourceFilter.NewList(iResourceLoop).field("RSC_VTS"),qcutil.QCConnection.ServerTime)
			
			'Default
			bDownloadResource = false
			
			'Check if the object exists within the Release Manager
			If oALMConn.value(oResourceFilter.NewList(iResourceLoop).field("RSC_FILE_NAME")) <> "" Then
				'ALM version is later version than the local version
				If datediff("s",oResourceFilter.NewList(iResourceLoop).field("RSC_VTS"),oALMConn.value(oResourceFilter.NewList(iResourceLoop).field("RSC_FILE_NAME"))) < 0 Then
					'Download Object
					bDownloadResource = true

					'Write out status
					reporter.ReportEvent micDone,"Engine" ,"Engine: Test Resource [" & oResourceFilter.NewList(iResourceLoop).field("RSC_NAME") & "] Is out of date and will be downloaded from ALM." 'V6.1.0 


				End If

			else
				'No Local version
				'Download Object
				bDownloadResource = true
				'Write out status
				reporter.ReportEvent micDone,"Engine" ,"Engine: Test Resource [" & oResourceFilter.NewList(iResourceLoop).field("RSC_NAME") & "] Does not exist locally so will be downloaded from ALM."
			End If
			
			'Build up the object repositories list
			If sResourceFolder = "Object Repositories" Then
				If right(oResourceFilter.NewList(iResourceLoop).field("RSC_FILE_NAME"),4) = ".tsr" Then
						sLoadOR=sLoadOR & ",C:\AAF\Object Repositories\" & oResourceFilter.NewList(iResourceLoop).field("RSC_FILE_NAME")
				End if
			End if

			If bDownloadResource = true Then
				On error resume next
				err.clear

				'Download resource
				oResourceFilter.NewList(iResourceLoop).DownloadResource "c:\AAF\" & sResourceFolder, True	
				
				'Set ReleaseManager Date time
				oALMConn.value(oResourceFilter.NewList(iResourceLoop).field("RSC_FILE_NAME")) = qcutil.QCConnection.ServerTime
				oALMConn.post
				
				'If bDebugPrint Then : print oResourceFilter.NewList(iResourceLoop).field("RSC_NAME") & " : DOWNLOADED FROM ALM"

				If err.number <> 0 then 
						reporter.ReportEvent micFail,"MainEngine","MainEngine - Fatal Error, Download from QC of " & oResourceFilter.NewList(iResourceLoop).field("RSC_NAME") & " failed"
						ExitTest
				end if
				
				On error resume next
			else
				'Write out status
				reporter.ReportEvent micDone,"Engine" ,"Engine: Test Resource [" & oResourceFilter.NewList(iResourceLoop).field("RSC_NAME") & "] Local version is the current version."
			End if
			
			'Switch depending on folder being processed
			If sResourceFolder = "Functions" Then
				Executefile "C:\AAF\Functions\" & oResourceFilter.NewList(iResourceLoop).field("RSC_FILE_NAME")				
			End If
			
			
		Next

		If sResourceFolder = "Object Repositories" Then
			sLoadOR = mid(sLoadOR,2)
		End if

		'Close the release manager connection
		oALMConn.close

End Function

'#########################################################################################################################################################################


Public function fFrame_LoadParse( byref aExec,byref aSteps, byref aActions, byref sTestName  )
	'Variables
	Dim  iStepLoop, iSwitchLoop, iCharCount,  iCharLoop, iSearchStart,iCurrSize, iCleanLoop, iExecArraySize ,iTestCaseId,iStepId
	Dim  iLogicLoopCheck,iEndIf,iIfThen,iElse,iLogicLoop,iLogicEndLoop
	Dim sSearchTerm,sParameterCheck 
	Dim sParam1,sParam2

	Dim iParamRows,iParamCols,sVDDPHeader
	'Action delimiters
	Dim sActionStart : sActionStart = "["
	Dim sActionEnd :  sActionEnd = "]"

	Dim aActionDetails,aStepIndex
	
	Dim bRunStatus: bRunStatus = True

	'Load & Parse
	'Set Exec Array size
	ReDim aExec(10,aSteps(1,0))
	ReDim aReportDetails(2,aSteps(1,0))
	aExec(1,0) = 0
	
	'Default bPONStart & bPONEnd				'V6.0
	bPONStart = false			'V6.0
	bPONEnd = false				'V6.0
	Dim oFileSystem 				'V6.0


	'Start Perf Timer
	'fFrame_StartPerfTimer "Parse", "AAF", 0	

	'Grab the DDP parameter string before virtual parameters are added to the string -This is used in fFrame_Run to check for duplicates 	'V6.0.3
	gDDPParams = aParam(0,0)																												'V6.0.3

	'Set gUseLiterlValues from DDP
	If Len(fFrame_GetParamData(2,"UseLiterals",false)) > 0 and  ucase(fFrame_GetParamData(2,"UseLiterals",false)) = "TRUE" Then			'V5.0
			gUseLiteralValues = true																						'V5.0	
	ElseIf Len(fFrame_GetParamData(2,"UseLiterals",false)) > 0 and  ucase(fFrame_GetParamData(2,"UseLiterals",false)) = "FALSE" Then	'V5.0
			gUseLiteralValues = false																						'V5.0
	else																													'V5.0
		'Default to true, if no value set in DDP																			'V5.0
		gUseLiteralValues = true																							'V5.0
	End If                         																							'V5.0

	'Process Steps
	'Loop through all steps
	For iStepLoop = 1 to aSteps(1,0)
			'Clear actions array for use
			ReDim aActions(0)
			aActions(0) = 0
			'Split Test id & Step Id
			iStepId = split(aSteps(0,iStepLoop),"~")(1)

			'Check for inconsistencies in the Start and End markers
			For iSwitchLoop = 1 to 2 '1 = Description 2 = Expected results
					iCharCount = 0
					For iCharLoop = 1 to len(aSteps(iSwitchLoop,iStepLoop))
							select case mid(aSteps(iSwitchLoop,iStepLoop),iCharLoop,1)
									Case sActionStart iCharCount = iCharCount + 1
									Case sActionEnd iCharCount = iCharCount - 1
							end select
							If iCharCount > 1 or iCharCount < 0 Then
									'Incorrect nmber of action start markers to action end markers or Nested Action commands
									reporter.ReportEvent micFail,"PRE-Check: Step ( " & iStepId & ")" ,"Unmatched number of Start markers '" & sActionStart & "' compared with number of End markers '" & sActionEnd & "'  or nested actions found in  (" & aSteps(iSwitchLoop,iStepLoop) & ")"
									sFrame_ExitTest()
							End If
					Next
			Next

			'Extract Automation from each step description
			iSearchStart = 1
			Do while  instr(iSearchStart,aSteps(1,iStepLoop),sActionStart) > 0  
							If instr( instr(1,aSteps(1,iStepLoop),sActionStart)+1,aSteps(1,iStepLoop),sActionEnd) > 0 Then
		
									'Expand Array
									iCurrSize = aActions(0) + 1
									ReDim Preserve aActions( iCurrSize)
									aActions(0) = iCurrSize

									'Exrtract the action out from the description and store into the aActions array
									aActions(aActions(0)) = mid(aSteps(1,iStepLoop),instr(iSearchStart,aSteps(1,iStepLoop),sActionStart), instr( instr(iSearchStart,aSteps(1,iStepLoop),sActionStart)+1,aSteps(1,iStepLoop),sActionEnd) - instr(iSearchStart,aSteps(1,iStepLoop),sActionStart) +1)

									'Set the start for the next search of actions with the description
									iSearchStart = instr( instr(iSearchStart,aSteps(1,iStepLoop),sActionStart)+1,aSteps(1,iStepLoop),sActionEnd) +1

									'If the current action is a Start or End action then ignore it, as it is not to be executed
									aStepIndex = split(aSteps(4,iStepLoop),"-")
									If (instr(1,ucase(aActions(aActions(0))),"^START") > 0)  and (cint(aStepIndex(0)) > 1)  and   (cint(aStepIndex(1))  > 1)          Then    'V6.1.0
											aActions(aActions(0)) = ""																		    
											aActions(0) = aActions(0) - 1
											iCurrSize = aActions(0)
											ReDim Preserve aActions( iCurrSize)
									End If
									If (instr(1,ucase(aActions(aActions(0))),"^END") > 0)  and (cint(aStepIndex(0)) < cint(aStepIndex(1)))  and  (cint(aStepIndex(1))  > 1)          Then  'V6.1.0
											aActions(aActions(0)) = ""
											aActions(0) = aActions(0) - 1
											iCurrSize = aActions(0)
											ReDim Preserve aActions( iCurrSize)
									End If
									If instr(1,ucase(aActions(aActions(0))),"PONSTART") > 0 Then		'V6.0
											aActions(aActions(0)) = ""			'V6.0
											aActions(0) = aActions(0) - 1		'V6.0
											iCurrSize = aActions(0)				'V6.0
											ReDim Preserve aActions( iCurrSize)	'V6.0
											bPONStart = true					'V6.0
									End If                 						'V6.0
									If instr(1,ucase(aActions(aActions(0))),"PONEND") > 0 Then	'V6.0
											aActions(aActions(0)) = ""			'V6.0
											aActions(0) = aActions(0) - 1		'V6.0
											iCurrSize = aActions(0)				'V6.0
											ReDim Preserve aActions( iCurrSize)	'V6.0
											bPONEnd = true						'V6.0
									End If             							'V6.0

							End If
			Loop

			'Extract Automation from each expected result
			iSearchStart = 1
			Do while  instr(iSearchStart,aSteps(2,iStepLoop),sActionStart) > 0 

							'Expand array
							If instr( instr(1,aSteps(2,iStepLoop),sActionStart)+1,aSteps(2,iStepLoop),sActionEnd) > 0 Then
									iCurrSize = aActions(0) + 1
									ReDim Preserve aActions( iCurrSize)
									aActions(0) = iCurrSize

									'Exrtract the action out from the descriptiona nd store into the aActions array
									aActions(aActions(0)) = mid(aSteps(2,iStepLoop),instr(iSearchStart,aSteps(2,iStepLoop),sActionStart), instr( instr(iSearchStart,aSteps(2,iStepLoop),sActionStart)+1,aSteps(2,iStepLoop),sActionEnd) - instr(iSearchStart,aSteps(2,iStepLoop),sActionStart) +1)

									'Set the start for the next search of actions with the description
									iSearchStart = instr( instr(iSearchStart,aSteps(2,iStepLoop),sActionStart)+1,aSteps(2,iStepLoop),sActionEnd) +1

									'If the current action is a Start or End action then ignore it, as it is not to be executed
									aStepIndex = split(aSteps(4,iStepLoop),"-")
									If (instr(1,ucase(aActions(aActions(0))),"^START") > 0)  and (cint(aStepIndex(0)) > 1)  and   (cint(aStepIndex(1))  > 1)          Then    'V6.1.0
											aActions(aActions(0)) = ""
											aActions(0) = aActions(0) - 1
											iCurrSize = aActions(0)
											ReDim Preserve aActions( iCurrSize)
									End If
									If (instr(1,ucase(aActions(aActions(0))),"^END") > 0)  and (cint(aStepIndex(0)) < cint(aStepIndex(1)))  and  (cint(aStepIndex(1))  > 1)          Then   'V6.1.0
											aActions(aActions(0)) = ""
											aActions(0) = aActions(0) - 1
											iCurrSize = aActions(0)
											ReDim Preserve aActions( iCurrSize)
									End If
									If instr(1,ucase(aActions(aActions(0))),"PONSTART") > 0 Then	'V6.0
											aActions(aActions(0)) = ""			'V6.0
											aActions(0) = aActions(0) - 1		'V6.0
											iCurrSize = aActions(0)				'V6.0
											ReDim Preserve aActions( iCurrSize)	'V6.0
											gPONStart = true					'V6.0
									End If                       				'V6.0
									If instr(1,ucase(aActions(aActions(0))),"PONEND") > 0 Then  'V6.0
											aActions(aActions(0)) = ""			'V6.0	
											aActions(0) = aActions(0) - 1		'V6.0
											iCurrSize = aActions(0)				'V6.0
											ReDim Preserve aActions( iCurrSize)	'V6.0
											gPONEnd = false						'V6.0
									End If           							'V6.0

							End If
			Loop

			'Sort automation for each step and load aActions array
			If  aActions(0) > 0 Then
				fGeneral_ArraySort aActions

				'clean and check for duplicate indices
				For iCleanLoop = 1 to aActions(0)
						aActions(iCleanLoop) = replace(replace(aActions(iCleanLoop),sActionEnd,""),sActionStart,"")
						If iCleanLoop > 1 Then
								If trim(mid(aActions(iCleanLoop),1,instr(1,aActions(iCleanLoop)," ")))  = trim(mid(aActions(iCleanLoop-1),1,instr(1,aActions(iCleanLoop-1)," "))) Then
										'Duplicate index found
										reporter.ReportEvent micFail,"PRE-Check: Step (" & split(aSteps(0,iStepLoop),"~")(1) & ")" ,"Duplicate automation indices have been used in the Description/Expected results for (" & split(aSteps(0,iStepLoop),"~")(1) & ")"
										bRunStatus = false
								End If
						End If

						'Write automation into aExec array
						iExecArraySize = aExec(1,0) + 1
						ReDim Preserve aExec(10,iExecArraySize)
						ReDim Preserve aReportDetails(2,iExecArraySize)
						
						'Remove any accidental double or more spaces in the action definition
						Do while instr(1,aActions(iCleanLoop),"  ") > 0  
									aActions(iCleanLoop) = replace(aActions(iCleanLoop),"  "," ")
						Loop


						'Parse Parameters check that the parameters names exists
						'Parse Parameters into aActions(iCleanLoop), loop until all <<< >>> replaced with values from  fFrame_GetParamData
						sParameterCheck = aActions(iCleanLoop)
						
						'replace <<<V- and >>> with ¬¬¬V- 											'V6.0
						Do while ubound(split(sParameterCheck,"<<<V-")) > 0   						'V6.0
								If (instr((instr(1,sParameterCheck,"<<<V-") + 5) ,sParameterCheck,">>>"))  < (instr(1,sParameterCheck,"<<<V-") + 5) Then	'V6.0
										'Paramter Parsing Failed
										reporter.ReportEvent micFail,"PRE-Check: Step (" & split(aSteps(0,iStepLoop),"~")(1) & ")" ,"Paramter Parsing Failed (" & split(aSteps(0,iStepLoop),"~")(1) & ") Action being processed for VDDP [" & sParameterCheck & "]"   'V6.0
										bRunStatus = false			'V6.0
								End If                               'V6.0
								'sParameterCheck = replace(sParameterCheck,"<<<V-" & (mid(sParameterCheck, (instr(1,sParameterCheck,"<<<V-") + 5) , (instr((instr(1,sParameterCheck,"<<<") + 3) ,sParameterCheck,">>>"))  - (instr(1,sParameterCheck,"<<<") + 3) ) )  & ">>>", fFrame_GetParamData(1,(mid(sParameterCheck, (instr(1,sParameterCheck,"<<<") + 3) , (instr((instr(1,sParameterCheck,"<<<") + 3) ,sParameterCheck,">>>"))  - (instr(1,sParameterCheck,"<<<") + 3) ) ),true))


								'Dim a,b,c,d,e,f,g,h
								'a = instr(1,sParameterCheck,"<<<V-")-1
								'b = instr(a,sParameterCheck,">>>")
								'c = mid(sParameterCheck,1,a)
								'd = mid(sParameterCheck,b)
								'e = replace(d,">>>","¬@¬",1,1)
								'f = mid(sParameterCheck,a+6,b-a-6)								
								'h = c & "@¬@" & f & e
								
								sVDDPHeader = "V-" & mid(sParameterCheck,(instr(1,sParameterCheck,"<<<V-")-1)+6,(instr((instr(1,sParameterCheck,"<<<V-")-1),sParameterCheck,">>>"))-(instr(1,sParameterCheck,"<<<V-")-1)-6)  'V6.0
								sParameterCheck =  mid(sParameterCheck,1,instr(1,sParameterCheck,"<<<V-")-1) & "@¬@" & mid(sParameterCheck,(instr(1,sParameterCheck,"<<<V-")-1)+6,(instr((instr(1,sParameterCheck,"<<<V-")-1),sParameterCheck,">>>"))-(instr(1,sParameterCheck,"<<<V-")-1)-6) & replace((mid(sParameterCheck,(instr((instr(1,sParameterCheck,"<<<V-")-1),sParameterCheck,">>>")))),">>>","¬@¬",1,1)  'V6.0
								
								'Load aParams with the Virtual DDP values
								iParamRows = ubound(aParam,1)				'V6.0
								iParamCols = ubound(aParam,2)				'V6.0
								ReDim preserve aParam(iParamRows, iParamCols + 1)	'V6.0
								
								
								aParam(0,0) = aParam(0,0) & "~" & sVDDPHeader & "#"		'V6.0
								aParam(1,iParamCols+1) = sVDDPHeader					'V6.0
								gVDDPHeader = gVDDPheader & "~" & sVDDPHeader & "#"		'V6.0
								
						loop
						
						
						Do while ubound(split(sParameterCheck,"<<<")) > 0   
								If (instr((instr(1,sParameterCheck,"<<<") + 3) ,sParameterCheck,">>>"))  < (instr(1,sParameterCheck,"<<<") + 3) Then
										'Paramter Parsing Failed
										reporter.ReportEvent micFail,"PRE-Check: Step (" & split(aSteps(0,iStepLoop),"~")(1) & ")" ,"Paramter Parsing Failed (" & split(aSteps(0,iStepLoop),"~")(1) & ") Action being processed [" & sParameterCheck & "]" 
										bRunStatus = false
								End If
								sParameterCheck = replace(sParameterCheck,"<<<" & (mid(sParameterCheck, (instr(1,sParameterCheck,"<<<") + 3) , (instr((instr(1,sParameterCheck,"<<<") + 3) ,sParameterCheck,">>>"))  - (instr(1,sParameterCheck,"<<<") + 3) ) )  & ">>>", fFrame_GetParamData(1,(mid(sParameterCheck, (instr(1,sParameterCheck,"<<<") + 3) , (instr((instr(1,sParameterCheck,"<<<") + 3) ,sParameterCheck,">>>"))  - (instr(1,sParameterCheck,"<<<") + 3) ) ),true))
						loop
						'Check if any errors occurred during the parameter lookup
						If instr(1,sParameterCheck, "~ERROR#") Then
										'Paramter Parsing Failed
										reporter.ReportEvent micFail,"PRE-Check: Step ( " & split(aSteps(0,iStepLoop),"~")(1) &")" ,"Paramter Parsing Failed (" &split(aSteps(0,iStepLoop),"~")(1) & ") Action being processed [" & sParameterCheck & "]" 
										bRunStatus = false
						End If

						'Parse Dynamic Parameters
						Do while instr(1,aActions(iCleanLoop),"{{{") > 0  
								If (instr((instr(1,aActions(iCleanLoop),"{{{") + 3) ,aActions(iCleanLoop),"}}}"))  < (instr(1,aActions(iCleanLoop),"{{{") + 3) Then
										'Paramter Parsing Failed
										reporter.ReportEvent micFail,"PRE-Check: Step (" & split(aSteps(0,iStepLoop),"~")(1) & ")" ,"Dynamic Paramter Parsing Failed (" & split(aSteps(0,iStepLoop),"~")(1) & ") Action being processed [" & aActions(iCleanLoop) & "]" 
										bRunStatus = false
								End If
								aActions(iCleanLoop) = replace(aActions(iCleanLoop),"{{{" & (mid(aActions(iCleanLoop), (instr(1,aActions(iCleanLoop),"{{{") + 3) , (instr((instr(1,aActions(iCleanLoop),"{{{") + 3) ,aActions(iCleanLoop),"}}}"))  - (instr(1,aActions(iCleanLoop),"{{{") + 3) ) )  & "}}}", fFrame_GetDynamicParamData((mid(aActions(iCleanLoop), (instr(1,aActions(iCleanLoop),"{{{") + 3) , (instr((instr(1,aActions(iCleanLoop),"{{{") + 3) ,aActions(iCleanLoop),"}}}"))  - (instr(1,aActions(iCleanLoop),"{{{") + 3) ) )))
						loop
						'Check if any errors occurred during the parameter lookup
						If instr(1,aActions(iCleanLoop), "~ERROR#") Then
										'Paramter Parsing Failed
										reporter.ReportEvent micFail,"PRE-Check: Step ( " & split(aSteps(0,iStepLoop),"~")(1) &")" ,"Dynamic Paramter Parsing Failed (" & split(aSteps(0,iStepLoop),"~")(1)  & ") Action being processed [" & aActions(iCleanLoop) & "]" 
										bRunStatus = false
						End If

						'Split the Action by space to get index into the action details
						aActionDetails = split(aActions(iCleanLoop)," ")
						If ubound(aActionDetails) = 0 Then
							reporter.ReportEvent micFail,"PRE-Check: Step (" & split(aSteps(0,iStepLoop),"~")(1) &")" ,"Command formatting error Step(" & split(aSteps(0,iStepLoop),"~")(1)  & ") Action being processed [" & aActions(iCleanLoop) & "] Minimum number of items within a command are: Index and Command" 
							sFrame_ExitTest()
						End If

						aExec(1,0) = iExecArraySize
						aExec(1,iExecArraySize) = aSteps(4,iStepLoop) 	'Test number within test set 1-14 (this is test case 1 of the 14 in the test set)
						aExec(2,iExecArraySize) = aSteps(0,iStepLoop) '  split(aSteps(0,iStepLoop),"~")(1) 	''Test id : Step Number
						aExec(10,iExecArraySize) = aSteps(3,iStepLoop) 	'Test Case name

						'Set the Step last actions value - used for determining if we have performed all of the actions within a step
						If  iCleanLoop > 0  and  iCleanLoop = aActions(0) then
								aExec(2,0) = aExec(2,0) & "#" & iExecArraySize & "#" 
						end if

						If  iCleanLoop > 1Then
								If  ucase(trim( aExec(2,iExecArraySize-1))) <> ucase(trim(aExec(2,iExecArraySize))) Then
										aExec(2,0) = aExec(2,0) & "#" & iExecArraySize - 1 & "#" 
								end if
						End If

						aExec(3,iExecArraySize) =  	 aActionDetails(1)  'Action
						aExec(9,iExecArraySize) = aActionDetails(0) 'Action number  


						'Check that action is a valid function
						If fFrame_GetFunctionIndex(ucase(aExec(3,iExecArraySize))) < 1 Then
								reporter.ReportEvent micFail,"PRE-Check:  Step ( " & split(aSteps(0,iStepLoop),"~")(1) &")" ,"Action [" & aExec(3,iExecArraySize) & "] has not been defined" 
								bRunStatus = false
						else					
								'Check if  the function is expecting an object to be passed to it
								If aFunctions(2,fFrame_GetFunctionIndex(ucase(aExec(3,iExecArraySize)))) > 0 then
										'Object exists
										aExec(5,iExecArraySize) =  	 aActionDetails(2)   'Object
										aReportDetails(1,iExecArraySize) =  	 aActionDetails(2)   'Object
										sSearchTerm = aActionDetails(2)
								else
										sSearchTerm = aActionDetails(1)
								End if
						End If

						'Check 



'# -----------------Literal values						
						

						'Save report details
						If instr(aActions(iCleanLoop),"^")  > 0       Then			'V5.0
								'Extract the parameters
								aReportDetails(2,iExecArraySize) =  trim(mid(aActions(iCleanLoop),instr(1,aActions(iCleanLoop) ,sSearchTerm)+ len(sSearchTerm) + 1,instr(1,aActions(iCleanLoop),"^") - (instr(1,aActions(iCleanLoop) ,sSearchTerm)+ len(sSearchTerm) + 1)))	'V5.0
						else															'V5.0
								'No Control code
								aReportDetails(2,iExecArraySize) = trim(mid(aActions(iCleanLoop),instr(1,aActions(iCleanLoop),sSearchTerm)+len(sSearchTerm)+1 ))       'Params	'V5.0
						end if                 																																	'V5.0
						
						

						If gUseLiteralValues = false Then																															'V5.0			
							'Check for || within commands, as this indicates that the following , is real and should not be used as a parameter seperator
							If instr(1,aActions(iCleanLoop),"||,") > 0 Then																								
								aActions(iCleanLoop) = replace(aActions(iCleanLoop),"||,","~#~")	
							End If
	
							If ucase(aExec(3,iExecArraySize)) = "SETTABLE" or  ucase(aExec(3,iExecArraySize)) = "GETTABLE" or  ucase(aExec(3,iExecArraySize)) = "SELECTROW" Then   'V3.8
								aActions(iCleanLoop) = replace(aActions(iCleanLoop),"||(","~£$~")	
								aActions(iCleanLoop) = replace(aActions(iCleanLoop),"||)","~$£~")	
							End If
						else																											'V5.0
							'New literal values being processed
							aActions(iCleanLoop) = fFrame_PreCallParam(aActions(iCleanLoop))											'V5.0
						
						End if                          																				'V5.0

						'Extract Control codes and parameters
						If instr(aActions(iCleanLoop),"^")  > 0       Then
								'Control code
								aExec(8,iExecArraySize) = trim(mid(aActions(iCleanLoop),instr(aActions(iCleanLoop),"^") ))
								'Extract the parameters
								aExec(6,iExecArraySize) =  trim(mid(aActions(iCleanLoop),instr(1,aActions(iCleanLoop) ,sSearchTerm)+ len(sSearchTerm) + 1,instr(1,aActions(iCleanLoop),"^") - (instr(1,aActions(iCleanLoop) ,sSearchTerm)+ len(sSearchTerm) + 1)))
						else
								'No Control code
								aExec(6,iExecArraySize) = trim(mid(aActions(iCleanLoop),instr(1,aActions(iCleanLoop),sSearchTerm)+len(sSearchTerm)+1 ))       'Params
						end if
						
						
						'Pause Processing
						If ucase(aExec(3,iExecArraySize)) = "PAUSE" and gUseLiteralValues = false Then				'V5.0
							sParam1 = mid(aExec(6,iExecArraySize),1,instr(1,aExec(6,iExecArraySize),","))
							sParam2 = mid(aExec(6,iExecArraySize),instr(1,aExec(6,iExecArraySize),",") + 1)	

							'Replace the comma and CRs
							sParam2 = replace(sParam2,",","~#~")
							sParam2 = replace (sParam2,chr(13),"~¬~")
							sParam2 = replace (sParam2,chr(10),"~§~")
							
							aExec(6,iExecArraySize) = sParam1 & sParam2
						End If
						
						



				Next   'iCleanLoop
			End If
	Next ' iStepLoop


	iIfThen= 0
	iEndIf = 0
	iElse = 0
	iLogicEndLoop = 0
	iLogicLoop = 0
	'Logic Command checks
	For iLogicLoopCheck = 1 to aExec(1,0)
			'If Then
			If fFrame_ControlCodes("^IFTHEN",aExec(8,iLogicLoopCheck)) = 1  Then
					iIfThen = iIFThen + 1
			End If

			'EndIf
			If ucase(aExec(3,iLogicLoopCheck)) = "ENDIF"  Then
					iEndIf = iEndIf + 1
			End If
			'Else
			If ucase(aExec(3,iLogicLoopCheck)) = "ELSE"  Then
					iElse = iElse + 1
			End If

			'Loop
			If ucase(aExec(3,iLogicLoopCheck)) = "LOOP"  Then
					iLogicLoop = iLogicLoop + 1
			End If

			'EndLoop
			If ucase(aExec(3,iLogicLoopCheck)) = "ENDLOOP"  Then
					iLogicEndLoop = iLogicEndLoop + 1
			End If

	Next

	'Check Logic structures
	If iIfThen <> iEndIf Then
			reporter.ReportEvent micFail,"PRE-Check:  Logic Command Structure " ,"[IfThen] do not match [EndIf]" 
			bRunStatus = false
	End If
	If iLogicLoop <> iLogicEndLoop Then
			reporter.ReportEvent micFail,"PRE-Check:  Logic Command Structure " ,"[Loop] do not match [EndLoop]"
			bRunStatus = false			
	End If
	If iElse > iIfThen Then
			reporter.ReportEvent micFail,"PRE-Check:  Logic Command Structure " ,"[Else] do not match [IfThen]" 
			bRunStatus = false
	End If

	' Check whether there is a PONEnd and no PONStart - Cannot have PONEnd without PONStart - Stop test
	If not(bPONStart) and bPONEnd Then    																																		'V6.0.1
			reporter.ReportEvent micFail,"PRE-Check: VDDP PON Commands" ,"VDDP PONStart[" & bPONStart & "] & PONEnd[" & bPONEnd & "] - Cannot have PONEnd without PONStart" 	'V6.0.1
			bRunStatus = false	'V6.0
			
	else
		'Valid PON commands found, now download PONStart & PONEnd files
		If bPONStart Then
			'PonStart
			On error resume next
			goPONStart.load true,""
			if err.number <> 0 Then
				reporter.ReportEvent micFail,"PRE-Check: VDDP" ,"PONStart file cannot be downloaded"
				bRunStatus = false			
			else
	 
				Set oFileSystem = CreateObject("Scripting.FileSystemObject") 
	
				'Check folder exists if not then create it
				If Not oFileSystem.FolderExists(sPONDirPath) Then
					oFileSystem.CreateFolder sPONDirPath
				End If
				
				'Delete all files and folders from the PON directory
				oFileSystem.DeleteFolder sPONDirPath & "*.*"					'V6.0
				fFrame_FileDelete(sPONDirPath & "*.*")							'V6.0
				
				'Copy file from temp to sPONDirPath
				oFileSystem.CopyFile goPONStart.filename , sPONDirPath & "PONStart.xls" , TRUE
				
				'Delete temp file
				if fFrame_FileDelete(goPONStart.filename) = 0 then
					reporter.ReportEvent micFail,"PRE-Check: VDDP" ,"Temp PONStart file cannot be deleted [" & goPONStart.filename & "]"
				End if
			End If		
			On error goto 0
		End if

		If bPONEnd Then
			'PONEnd
			On error resume next
			goPONEnd.load true,""
			gPONEndFileName = ""
			if err.number <> 0 Then
				reporter.ReportEvent micDone,"PRE-Check: VDDP" ,"PONEnd file cannot be downloaded, only rollback file will be processed!"
			else
	 
				Set oFileSystem = CreateObject("Scripting.FileSystemObject") 
	
				'Check folder exists if not then create it
				If Not oFileSystem.FolderExists(sPONDirPath) Then
					oFileSystem.CreateFolder sPONDirPath
				End If
	
				'Copy file from temp to sPONDirPath
				oFileSystem.CopyFile goPONEnd.filename , sPONDirPath & "PONEnd.xls" , TRUE
				gPONEndFileName = "PONEnd.xls"
				'Delete temp file
				if fFrame_FileDelete(goPONEnd.filename) = 0 then
					reporter.ReportEvent micFail,"PRE-Check: VDDP" ,"Temp PONEnd file cannot be deleted [" & goPONEnd.filename & "]"
				End if
			End If		
			On error goto 0
		End if

		'Copy DDP
		If  gDDPUpdate = false Then
			'DDP has been deleted, download it again
			On error resume next
			gDDPObject.load true,""
			if err.number <> 0 Then
				reporter.ReportEvent micFail,"PRE-Check: VDDP" ,"DDP file cannot be downloaded for PON"
				sFrame_ExitTest()				
			End if
		End if
		
		'DDP in temp folder	
		Set oFileSystem = CreateObject("Scripting.FileSystemObject") 

		'Check folder exists if not then create it
		If Not oFileSystem.FolderExists(sPONDirPath) Then
			oFileSystem.CreateFolder sPONDirPath
		End If

		'Copy file from temp to sPONDirPath
		If bDDPPresent = true Then																					'V6.0.3			
				On error resume next																				'V6.0.1
				oFileSystem.CopyFile gDDPObject.filename , sPONDirPath & "DDP.xls" , TRUE
				if err.number <> 0 Then																				'V6.0.1
					reporter.ReportEvent micFail,"PRE-Check: VDDP" ,"DDP file cannot be downloaded for PON"			'V6.0.1
					sFrame_ExitTest()																				'V6.0.1
				End if																								'V6.0.1
				On error goto 0																						'V6.0.1
		End If 																										'V6.0.3
	End If       				'V6.0


	'If the pre-checks/Loads etc have failed then stop the run
	If bRunStatus = false Then sFrame_ExitTest()
	
	'Start Perf Timer
	'fFrame_EndPerfTimer "Parse"	
end  function


Public function fFrame_Execute ( byref aExec, byval bDebugPrint, byref aFunctions  )

	'Variables
	dim iActionStatus, iFunctionParameters,  iPadLoop, iIterationLoop, iIterationMax, iPrevIteration,iStepChange, iActionResponseLoop,iNumParams
	Dim iSearchLoop,iPrevStep
	dim sParams,sLocalAction,sLocalObject,sLocalParam,sLocalControlCodes,sActionDetails,sCaptureCommand,sActionStatus,sActionResponse
	Dim aLocalControlCodes,aPostProc,aActionStatus
	Dim bIfFlagProcess : bIfFlagProcess = true
	Dim bJumpToTopOfStep : bJumpToTopOfStep = false

	Dim oTestSetF,oTestSetObj,oTestF,oTestObj,oRun,oStep,oStepList,oCreateFolder,oFolder

	Dim aTestDetails, aPrevTestDetails,aNextTestDetails
	Dim sStepStatus : sStepStatus = "No Run"
	Dim sTestStatus : sTestStatus = "No Run"
	Dim iLocalStart,iLocalEnd,iLocalStep								'V3.1

	Dim iDynLoop,iStart,iDynStart,iEnd				'V3.8
	Dim sStart,sEnd,sDyn							'V3.8
	
	Dim iNested 									'V4.2
	Dim iCommandCounter								'V4.7
	



	'Start Perf Timer
	'fFrame_StartPerfTimer "Execute", "AAF", 0	

	'Check if DDP spreadsheet has been processed
	If bDDPPresent = true Then
			'If DDP has been processed set the iteration to the number of rows in the spreadsheet
			iIterationMax =   ubound(aParam,1)
	else
			'If no DDP then only iterate once
			iIterationMax = 2
	End If

	'Create Test Set Object
	Set oTestSetF =  oQCConn.TestSetFactory

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	If gALMProject = "AAF_Beta" then
		Set oTestSetObj = oTestSetF.Item(xxTestSetId)
	else
		Set oTestSetObj = oTestSetF.Item(qcutil.CurrentTestSet.id)
	End if
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			
	'Iteration Loop
	For iIterationLoop = 2 to iIterationMax
					'If first Column in data sheet is empty then no more data
					'Dummy data should be used in first column if required
			
					If len(fFrame_GetParamData(iIterationLoop, 1,false)) > 0  or bDDPPresent = false Then
						'Execution Loop
						For iExecLoop = 1 to aExec(1,0)
						
								'Start Perf Timer
								'fFrame_StartPerfTimer "Pre-command:" & aExec(3,iExecLoop), "AAF", 0
								
								'Set the action status to fail
								iActionStatus = 0
					
								'Parse for current iteration loop, with data in aParam
								sLocalAction = ""   'Parse Action name
								sLocalObject = ""   'Parse object
								sLocalParam = ""   'Parse Parameters
								sLocalControlCodes =""   'Parse Control Codes

								sLocalAction = aExec(3,iExecLoop)   'Action name
								
								'Set Global call stack monitor
								gCallStack = sLocalAction				'V5.0
								

								sLocalObject = fFrame_ParseString (aExec(5,iExecLoop),  iIterationLoop, false)   'Parse object
								sLocalParam = fFrame_ParseString (aExec(6,iExecLoop),  iIterationLoop, false)   'Parse Parameters
								sLocalControlCodes = fFrame_ParseString (aExec(8,iExecLoop),  iIterationLoop, false)   'Parse Control Codes
								
								'Report details
								gReportObject = replace(replace(fFrame_ParseString (aReportDetails(1,iExecLoop),  iIterationLoop, true),chr(13),""),chr(10),"")  'Parse object		'V5.0
								gReportParam =  replace(replace(fFrame_ParseString (aReportDetails(2,iExecLoop),  iIterationLoop, true),chr(13),""),chr(10),"")   'Parse Parameters	'V5.0
								
								
								
								
								gControlCodes = sLocalControlCodes   'set global control codes											'V4.5
					
								'Get the passed in Parameters
								sParams = sLocalParam
					
								'Check if action exists in the list of functions that we have coded up. NULL if not found
								If instr(1,aFunctions(0,0),"~" & ucase(sLocalAction) & "#") = 0 Then
										'Function call failed
										reporter.ReportEvent micFail,"Iteration[" &  iIterationLoop -1 & "]  Step Name: " & aExec(2,iExecLoop) & " Action " & aExec(9,iExecLoop) & " [" & sLocalAction & "]" , "Action " & aExec(9,iExecLoop) & " [" & sLocalAction & "] - FAILED.  Parameter Look up failed to find action" 							
										sFrame_ExitTest()
								else
									'Get the expected number of parameters for the function
									iFunctionParameters = aFunctions(3,fFrame_GetFunctionIndex(ucase(sLocalAction))) + aFunctions(4,fFrame_GetFunctionIndex(ucase(sLocalAction)))
									
									iNumParams =  ubound(split(sParams,",")) + 1

									If iNumParams < aFunctions(3,fFrame_GetFunctionIndex(ucase(sLocalAction)))  Then
											'Function call failed
											reporter.ReportEvent micFail,"Iteration[" &  iIterationLoop -1 & "]  Step Name: " & aExec(2,iExecLoop) & " Action " & aExec(9,iExecLoop) & " [" & sLocalAction & "]" , "Action " & aExec(9,iExecLoop) & " [" & sLocalAction & "] - FAILED.  Action defined in QC has less mandatory parameters [" & ubound(split(gReportParam,","))+ 1 & "] than expected [" & aFunctions(3,fFrame_GetFunctionIndex(ucase(sLocalAction))) & "]." 							
											sFrame_ExitTest()
									elseIf iNumParams < iFunctionParameters then
											'Pad out with dummy values
											For iPadLoop =  ubound(split(sParams,","))+2 to iFunctionParameters
													If len(sParams) = 0 Then
															sParams = "**~**" 
													else
															sParams = sParams & ",**~**"
													End If
											Next
									elseif ubound(split(sParams,","))+1 > iFunctionParameters then
											'Function call failed
											reporter.ReportEvent micFail,"Iteration[" &  iIterationLoop -1 & "]  Step Name: " & aExec(2,iExecLoop) & " Action " & aExec(9,iExecLoop) & " [" & sLocalAction & "]" , "Action " & aExec(9,iExecLoop) & " [" & sLocalAction & "] - FAILED.  Action defined in QC has more parameters [" & ubound(split(gReportParam,","))+ 1 & "] than expected [" & iFunctionParameters & "]." 							
											sFrame_ExitTest()
									end if
								End If
					
								'Format the sParams string ready for the function call
								If len(sParams) > 0 Then
										sParams = chr(34) & replace(sParams,",",chr(34) & "," & chr(34)) & chr(34)
								end if
					
								'Check if we also need to send a object to the function
								if aFunctions(2,fFrame_GetFunctionIndex(ucase(sLocalAction))) > 0  Then
										'Command needs the object to also be sent to the function
										If len(sParams) > 0  Then
												sParams = chr(34) & sLocalObject & chr(34)  & "," & sParams
										else
												sParams = chr(34) & sLocalObject & chr(34) 
										end if
								end if

								'First loop through the commands, create a new run
								aTestDetails = split(aExec(2, iExecLoop),"~")
								aPrevTestDetails = split(aExec(2, iExecLoop-1),"~")
								If iExecLoop < aExec(1,0) Then
										aNextTestDetails = split(aExec(2, iExecLoop+1),"~")
								End If

								'Check if we are moving from pre-reqs to actual test steps
								iStepChange = ubound(aTestDetails) - ubound(aPrevTestDetails)
								if  iStepChange = -1 or ( iExecLoop = 1 and  ubound(aTestDetails) = 1 and bJumpToTopOfStep = false) or  ((aPrevTestDetails(0) <> aTestDetails(0)) and iStepChange = 0)  or ((iIterationLoop <>  iPrevIteration) and ubound(aTestDetails) = 1) then   'V4.3
										'Get object from Test Set
										Set oTestF = oTestSetObj.TSTestFactory
										Set oTestObj  = oTestF.item(int(aTestDetails(0)))

										'create new run
										oTestObj.status =  "Not Completed"
										oTestObj.post
										set oRun =  oTestObj.RunFactory.additem(null)
										oRun. status = "Not Completed"
										oRun.name = "QTP: Iteration " &  iIterationLoop -1 & "  [" & gDDPLocation & "]"     'V6.0
										oRun.Post()
										oRun.copydesignsteps()
										oRun.post()
										set oStepList = oRun.StepFactory.NewList("") 
										iPrevStep = 0	
										bJumpToTopOfStep = false										
								end if




								'Call action functions
								sActionStatus = ""
								iActionStatus = ""
								sActionDetails = ""							'V4.3
								If  instr(1,sLogicCommands,"~" & UCASE(sLocalAction) & "#") = 0  and  bIfFlagProcess = true Then
									'Command  is not a Control LOGIC command so it should be executed

									'Set Global Vars for enhanced debug
									gStep = int(aTestDetails(1))
									gAction = int(aExec(9,iExecLoop))
									gDDPRow = iIterationLoop - 1							'V3.8
									call sFrame_SetGlobalVar("gv_DDPRow", gDDPRow)			'V3.8

									'Replace Dynamic Variables								'V3.8
									For iDynLoop = 1 To 10									'V3.8				
										
										If instr(1,sParams,"+ {") then						'V3.8
											iStart = instr(1,sParams,"+ {") -1				'V3.8
											iDynStart = iStart + 4							'V3.8
										
										elseif instr(1,sParams,"+{") Then					'V3.8
											iStart = instr(1,sParams,"+{") - 1				'V3.8
											iDynStart = iStart + 3							'V3.8
										else												'V3.8
											Exit for										'V3.8	
										End if
										
										iEnd = instr(iStart,sParams,"}")					'V3.8
										sStart = mid(sParams,1,iStart)						'V3.8
										sDyn = mid(sParams,iDynStart,iEnd-iDynStart)		'V3.8
										sEnd = mid(sParams,iEnd+1)							'V3.8			
										
										sParams = sStart & fFrame_GetGlobalVar(sDyn) & sEnd	'V3.8

									Next													'V3.8


									'Replace the comma and CRs
								If gUseLiteralValues = false Then																															'V5.0			
									sParams = replace(sParams,"~#~",",")
								End if
								If instr(1,sParams,chr(10))>0 or instr(1,sParams,chr(13))>0  Then
									sParams = replace(sParams,chr(10),"")
									sParams = replace(sParams,chr(13),"")
									reporter.ReportEvent micWarning,"Iteration[" &  iIterationLoop -1 & "]  Test[" & aExec(10,iExecLoop)   &  "] Step[" & gStep & "] Action [" & gAction & "]", "Step[" & gStep & "] Action [" & gAction & "] - Parameters contains a CR or LF (Probably from ALM line wrapping). This has been removed." 
								End If
									
									
									'OptionalNavPathTime
									If sLocalAction = "OptionalNavPathTime" Then			'V4.7	
										'Set Command counter = 0
										iCommandCounter = 0									'V4.7
									elseIf iCommandCounter < 999 Then						'V4.7
											If iCommandCounter = 1 Then						'V4.7
												'Return iGUIOverloadSyncTime to default
												 iGUIOverloadSyncTime =  iGUIOverloadSyncTimeDefault	'V4.7	
												 iGUIOverloadStablityTime = iGUIOverloadStabilityTimeDefault 'V6.0
												 iCommandCounter = 999									'V4.7
											End If                                						'V4.7
											iCommandCounter = iCommandCounter + 1							'V4.7
									End If                                       						'V4.7
									
									
									
									If bDebugPrint Then :print iExecLoop & ":  fCommand_" & sLocalAction & "(" &  sParams & ")"	
									'End Perf Timer
									'fFrame_EndPerfTimer "Pre-command:" & aExec(3,iExecLoop)
									
									'Perf Timer
									'fFrame_StartPerfTimer "Execute:" & sLocalAction & " [Iteration:" & gDDPRow & " ,Step:" & gStep & " ,Command:" & gAction & "]",  "AAF", 0
																		
									'Execute the command
									Execute "sActionResponse = fCommand_" & sLocalAction & "(" &  sParams & ")"
									
									'ReSet Global call stack monitor
									gCallStack = ""				'V5.0


									'Perf Timer
									'fFrame_EndPerfTimer  "Execute:" & sLocalAction & " [Iteration:" & gDDPRow & " ,Step:" & gStep & " ,Command:" & gAction & "]"

									'Start Perf Timer
									'fFrame_StartPerfTimer "Post-command:" & aExec(3,iExecLoop), "AAF", 0	
								else																				'Test
										sActionResponse = "1"															'Test
										'If a nested IFTHEN
										If (fFrame_ControlCodes("^IFTHEN",sLocalControlCodes) > 0) Then		  	'V4.2
											iNested = iNested + 1			  	'V4.2
										End if                             	'V4.2
										if (fFrame_ControlCodes("^IGNORE",sLocalControlCodes) = 0) then			'V4.3
											sLocalControlCodes = ""														'Test
										else																	'V4.3
											sLocalControlCodes = "^IGNORE"										'V4.3
										End If
								End If
								If bDebugPrint Then : print sActionResponse 


'								aActionStatus = split(sActionResponse,":")
								aActionStatus = split(sActionResponse,":#~")
								iActionStatus = int(aActionStatus(0))
								If ubound(aActionStatus) > 0 Then
									For iActionResponseLoop = 1 to ubound(aActionStatus)
										If   iActionResponseLoop = 1 Then
											sActionStatus = aActionStatus(iActionResponseLoop)
										else
											sActionStatus = sActionStatus & "," & aActionStatus(iActionResponseLoop)
										End If
									Next
								End If

					
								'Split test details
								aTestDetails = split(aExec(2, iExecLoop),"~")

								'If Logic control then set status to passed
'								If  instr(1,sLogicCommands,"~" & UCASE(sLocalAction) & "#") > 0  or  bIfFlagProcess = false Then
'										iActionStatus = 1
'										sLocalControlCodes = ""
'								end if

								
								'Check if the status of this command should be ignored
								If (fFrame_ControlCodes("^IGNORE",sLocalControlCodes) = 0 and  instr(1,sLogicCommands,"~" & UCASE(sLocalAction) & "#") = 0  and (fFrame_ControlCodes("^IFTHEN",sLocalControlCodes) = 0)) Then    'V4.3
									'Action Status processing
									If   instr(sLogicCommands,"~" & sLocalAction & "#") > 0 Then 'Hard code the logic commands to always pass
											If  ubound(aTestDetails) = 1  Then
													'Action Details
													sActionDetails ="N/A"					'V4.6
											else
													'Pre/Post Req Action N/A
													reporter.ReportEvent micPass,"Iteration[" &  iIterationLoop -1 & "]  Test[" & aExec(10,iExecLoop)   &  "]  Step[ " & aTestDetails(1) & "] Action " & aExec(9,iExecLoop) & " [" & sLocalAction &" " &  gReportObject & "]" , "Action " & aExec(9,iExecLoop) & " [" & sLocalAction & " "  & gReportObject & "] - N/A" 
											end if
											If sStepStatus = "No Run" Then
													sStepStatus = "N/A"
											End If
											If sTestStatus = "No Run" Then
													sTestStatus = "N/A"
											End If
									elseif bIfFlagProcess = false then
											If  ubound(aTestDetails) = 1  Then
													'Action Details
													sActionDetails = "N/A"				'V4.6
											else
													'Pre/Post Req Action N/A
													reporter.ReportEvent micPass,"Iteration[" &  iIterationLoop -1 & "]  Test[" & aExec(10,iExecLoop)   &  "]  Step[ " & aTestDetails(1) & "] Action " & aExec(9,iExecLoop) & " [" & sLocalAction &" " &  gReportObject & "]" , "Action " & aExec(9,iExecLoop) & " [" & sLocalAction & " "  & gReportObject & "] - N/A" 
											end if
											If sStepStatus = "No Run" Then
													sStepStatus = "N/A"
											End If
											If sTestStatus = "No Run" Then
													sTestStatus = "N/A"
											End If
									elseIf   ((iActionStatus = 0 or iActionStatus = -1)and len(sLocalControlCodes) = 0) or  ((iActionStatus = 0 or iActionStatus = -1) and  fFrame_ControlCodes("^FAIL",sLocalControlCodes) = 0)  or  (iActionStatus = 1 and fFrame_ControlCodes("^FAIL",sLocalControlCodes) = 1)  Then										'Action Failed
											If  ubound(aTestDetails) =1  Then
													'Action Details
													sActionDetails = "FAILED"			'V4.6
											else
												'Pre/Post Req Action Failed
												reporter.ReportEvent micFail,"Iteration[" &  iIterationLoop -1 & "]  Test[" &  aExec(10,iExecLoop)  &  "] Step[ " & aTestDetails(1) & "] Action " & aExec(9,iExecLoop) & " [" & sLocalAction &" " &  gReportObject & "]" , "Action " & aExec(9,iExecLoop) & " [" & sLocalAction & " "  & gReportObject & "] - FAILED" 
											end if
											'StepStatus
											sStepStatus = "Failed"
						
											'Update Test  Status
											sTestStatus = "Failed"
									else 
											If  ubound(aTestDetails) = 1  Then
													'Action Details
													sActionDetails = "PASSED"			'V4.6

											else
													'Pre/Post Req Action Passed
													reporter.ReportEvent micPass,"Iteration[" &  iIterationLoop -1 & "]  Test[" & aExec(10,iExecLoop)   &  "] Step[ " & aTestDetails(1) & "] Action " & aExec(9,iExecLoop) & " [" & sLocalAction &" " &  gReportObject & "]" , "Action " & aExec(9,iExecLoop) & " [" & sLocalAction & " "  & gReportObject & "] - PASSED" 
											end if
											If sStepStatus <> "Failed" Then
													sStepStatus = "Passed"
											End If
											If sTestStatus <> "Failed" Then
													sTestStatus = "Passed"
											End If
									End If
								End if

								'Update the actions/functions status within the step.
								If iPrevStep <>  int(aTestDetails(1)) Then
										Set oStep = oStepList.item(aTestDetails(1))
										iPrevStep = int(aTestDetails(1))
								end if
								
								 oStep.Field("ST_ACTUAL") = oStep.Field("ST_ACTUAL") & chr(10) & sActionDetails & " - Action " & aExec(9,iExecLoop) & " [" & sLocalAction & " "  & gReportObject & "]  Parameters [" & gReportParam & "] " 	'V4.6

								 'Additional info passed back from function
								 If (fFrame_ControlCodes("^IGNORE",sLocalControlCodes) > 0 ) Then  
											  oStep.Field("ST_ACTUAL") = replace(oStep.Field("ST_ACTUAL")," - Action " & aExec(9,iExecLoop),"IGNORED - Action " & aExec(9,iExecLoop))
								 ElseIf (fFrame_ControlCodes("^IFTHEN",sLocalControlCodes) > 0) or ( instr(1,sLogicCommands,"~" & UCASE(sLocalAction) & "#") > 0)Then							'V4.3
								 		If bIfFlagProcess = true Then                                                   'V4.3
											 If fFrame_ControlCodes("^IFTHEN",sLocalControlCodes) > 0 Then
											 	oStep.Field("ST_ACTUAL") = replace(oStep.Field("ST_ACTUAL"),chr(10) & " - Action " & aExec(9,iExecLoop),chr(10) & "CONTROL LOGIC - Action " & aExec(9,iExecLoop))  & "  - [Actual Data " &  sActionStatus & " ]" 		'V4.3
											else
											 	oStep.Field("ST_ACTUAL") = replace(oStep.Field("ST_ACTUAL"),chr(10) & " - Action " & aExec(9,iExecLoop),chr(10) & "CONTROL LOGIC - Action " & aExec(9,iExecLoop)) 'V4.3
											 End if
										else
											 oStep.Field("ST_ACTUAL") = replace(oStep.Field("ST_ACTUAL"),chr(10) & " - Action " & aExec(9,iExecLoop),chr(10) & "N/A - Action " & aExec(9,iExecLoop))		'V4.3
										End if
								 elseIf len(sActionStatus) > 0  Then
											 oStep.Field("ST_ACTUAL") = oStep.Field("ST_ACTUAL") & " - [Actual Data " &  sActionStatus & " ]"       'V4.3 Moved into the elseif
								 End If
								 ' oStep.post()   																V4.6
	
																
					
								' If the last action in the step then check the Step Status
								If instr(1,aExec(2,0) ,"#" & iExecLoop & "#") > 0 Then
										If  ubound(aTestDetails) = 1  Then
											Set oStep = oStepList.item(aTestDetails(1))
											 oStep.Status = sStepStatus 
											 oStep.post() 
											sStepStatus = "No Run"
											sActionDetails = ""
										else
											'Step Passed
											If sStepStatus = "Passed" Then
												reporter.ReportEvent micPass,"Iteration[" &  iIterationLoop -1 & "]  Test[" &  aExec(10,iExecLoop)  &  "] Step[ " & aTestDetails(1) & "]"  & " - PASSED" , " PRE/POST Req [ " & aExec(2,iExecLoop) & "] - PASSED" 
											else
												reporter.ReportEvent micFail,"Iteration[" &  iIterationLoop -1 & "]  Test[" &  aExec(10,iExecLoop)  &  "] Step[ " & aTestDetails(1) & "]"  & " - FAILED" , " PRE/POST Req [ " & aExec(2,iExecLoop) & "] - FAILED" 
											End If

										End If
								End if

								'Set Test Status within the test set
								If    iExecLoop < aExec(1,0) Then
										If ( (aTestDetails(0) <> split(aExec(2, iExecLoop+1),"~")(0)) or ((ubound(aNextTestDetails) - ubound(aTestDetails) <> 0) and iExecLoop > 1 )) Then
												If  ubound(aTestDetails) = 1  Then
														'Steps
														oTestObj.status = sTestStatus
														oTestObj.post
														oRun. status = sTestStatus
														oRun.Post

														'Update DDP
														sFrame_UpdateDDP iIterationLoop, gDDPCol , sTestStatus

														'If Test failed then grab required log files
														If sTestStatus = "Failed" then 
																fFrame_LoadLogs(oRun)
														End if

												else
														'pre/Post reqs
														If sStepStatus = "Passed" or sStepStatus = "N/A" Then
																reporter.ReportEvent micPass,"Iteration[" &  iIterationLoop -1 & "]  Test[" &  aExec(10,iExecLoop)  &  "] - " & sTestStatus , " PRE/POST Req [ " & aExec(2,iExecLoop) & "] - " & sTestStatus 
														else
																reporter.ReportEvent micFail,"Iteration[" &  iIterationLoop -1 & "]  Test[" &  aExec(10,iExecLoop)  &  "] - " & sTestStatus , " PRE/POST Req [ " & aExec(2,iExecLoop) & "] - " & sTestStatus 
														end if
												end if
												sStepStatus = "No Run"
												sTestStatus = "No Run"
										End If
								elseif  (iExecLoop=  aExec(1,0)) then
										If  ubound(aTestDetails) = 1  Then
												'Steps
												oTestObj.status = sTestStatus
												oTestObj.post
												oRun. status = sTestStatus
												oRun.Post

												'Update DDP
												sFrame_UpdateDDP  iIterationLoop, gDDPCol , sTestStatus

												'If Test failed then grab required log files
												If sTestStatus = "Failed" then 
														fFrame_LoadLogs(oRun)
												End if


										else
												'pre/Post reqs
												If sStepStatus = "Passed" or sStepStatus = "N/A" Then
														reporter.ReportEvent micPass,"Iteration[" &  iIterationLoop -1 & "]  Test[" &  aExec(10,iExecLoop)  &  "] - " & sTestStatus , " PRE/POST Req [ " & aExec(2,iExecLoop) & "] - " & sTestStatus 
												else
														reporter.ReportEvent micFail,"Iteration[" &  iIterationLoop -1 & "]  Test[" &  aExec(10,iExecLoop)  &  "] - " & sTestStatus , " PRE/POST Req [ " & aExec(2,iExecLoop) & "] - " & sTestStatus 
												end if
										end if
										 sStepStatus = "No Run"
										 sTestStatus = "No Run"
								End If
					
								'If critical error eg. ActionStatus = -1 then exit run
								If  iActionStatus = -1 Then
										reporter.ReportEvent micFail,"Iteration[" &  iIterationLoop -1 & "]  Test: " & aExec(10,iExecLoop)  & " - CRITICAL FAILURE"  , " FAILED - Test stopped" 
										'If Test failed then grab required log files
										fFrame_LoadLogs(oRun)
										sFrame_ExitTest()
								End If


								'Control Code processing
								If (fFrame_ControlCodes("^IFTHEN",sLocalControlCodes) > 0) and (fFrame_ControlCodes("^IGNORE",sLocalControlCodes) = 0 ) Then
										'IFTHEN :  If the command = 1 then process the following commands up to an 'ELSE' or 'ENDIF'. The same as IFComapre, but can be used on any command
										
										'Take the result from the compare function and check result

										If   ((iActionStatus = 0 or iActionStatus = -1)and len(sLocalControlCodes) = 0) or  ((iActionStatus = 0 or iActionStatus = -1) and  fFrame_ControlCodes("^FAIL",sLocalControlCodes) = 0)  or  (iActionStatus = 1 and fFrame_ControlCodes("^FAIL",sLocalControlCodes) = 1)  Then										'Action Failed
												'False - Jump excution to 'ELSE' + 1
												 bIfFlagProcess = false
										else
												'True - Process next commands
												bIfFlagProcess =true
										End If
								End if
								If (fFrame_ControlCodes("^CAPTUREON",sLocalControlCodes) > 0)  and (fFrame_ControlCodes("^IGNORE",sLocalControlCodes) = 0 ) Then
											'Set Global Capture to ON - Any fails will do a screen grab
											gCapture = true
								end if
								If (fFrame_ControlCodes("^CAPTUREOFF",sLocalControlCodes) > 0)  and (fFrame_ControlCodes("^IGNORE",sLocalControlCodes) = 0 ) Then
											'Set Global Capture to ON - Any fails will do a screen grab
											gCapture = false
								end if
								If (((fFrame_ControlCodes("^CAPTURE",sLocalControlCodes) > 0) or (gCapture = true and  iActionStatus <> 1) or _
								 ( (fFrame_ControlCodes("^CAPTURETRUE",sLocalControlCodes) > 0)  and iActionStatus = 1) or _
								  ( (fFrame_ControlCodes("^CAPTUREFALSE",sLocalControlCodes) > 0)  and iActionStatus <> 1)))   and (fFrame_ControlCodes("^IGNORE",sLocalControlCodes) = 0 ) Then
										'IFTHEN :  If the command = 1 then process the following commands up to an 'ELSE' or 'ENDIF'. The same as IFComapre, but can be used on any command
										
											'Check that an object has been used
											If sLocalObject = "" Then
													reporter.ReportEvent micDone,"Iteration[" &  iIterationLoop -1 & "]  Test[" &  aExec(10,iExecLoop)  &  "] - Step[ " & aTestDetails(1) & "] " & sTestStatus , " Capture cannot be executed as no Object has been provided" 
											else	
												'Get the command text to strip off the last objects back to a window, then capture the window
												sCaptureCommand = fFrame_BuildCall (sLocalObject) 
	
												'Create the folder if it doesn't exist
												Set oCreateFolder = CreateObject("Scripting.FileSystemObject")
												On Error Resume Next
												Set oFolder = oCreateFolder.CreateFolder("c:\AAF")
												On Error Goto 0
	
												'Execute the capture on window
												Execute mid(sCaptureCommand , 1, instr(instrRev(sCaptureCommand,chr(34) & "wnd",len(sCaptureCommand)),sCaptureCommand,chr(34)&").") + 1) & ".capturebitmap " & chr(34) & "c:\AAF\Step" & gStep & "_Action" & gAction & "_AAF_Capture.png" & chr(34) &", true"
	
												'Upload attachment to Step within test run
												 call fFrame_UploadAttachment (oStepList.item(aTestDetails(1)),  "c:\AAF\Step" & gStep & "_Action" & gAction & "_AAF_Capture.png", True) 
											 End if
								End If

								If (fFrame_ControlCodes("^DDP-UpdateOn",sLocalControlCodes) > 0)  and (fFrame_ControlCodes("^IGNORE",sLocalControlCodes) = 0 ) Then
											'Set Global DDP_Update to ON - Any processing of DDPs will be marked with 

											If  gDDPUpdate = false Then
													'If  DDP not open for updating
													gDDPUpdate = true

													'Open DDP
													fFrame_OpenDDP 								
											End If
											gDDPUpdate = true
								end if

								If (fFrame_ControlCodes("^DDP-UpdateOff",sLocalControlCodes) > 0)  and (fFrame_ControlCodes("^IGNORE",sLocalControlCodes) = 0 ) Then
											'Set Global DDP_Update to ON - Any processing of DDPs will be marked with 
											gDDPUpdate = false

											'Close DDP logging
											sFrame_CloseDDP 
								end if

								If (fFrame_ControlCodes("^LogFilesOn",sLocalControlCodes) > 0)  and (fFrame_ControlCodes("^IGNORE",sLocalControlCodes) = 0 ) Then
											'Set Global Capture to ON - Any fails will do a screen grab
											gCaptureLogFiles = true
								end if

								If (fFrame_ControlCodes("^LogFilesOff",sLocalControlCodes) > 0)  and (fFrame_ControlCodes("^IGNORE",sLocalControlCodes) = 0 ) Then
											'Set Global Capture to ON - Any fails will do a screen grab
											gCaptureLogFiles = false
								end if

								'Logic Control processing
								If  (instr(1,sLogicCommands,"~" & UCASE(sLocalAction) & "#") > 0)  and (fFrame_ControlCodes("^IGNORE",sLocalControlCodes) = 0 )  Then    
										'Command  Control LOGIC command so it requires post processing

										'Switch on command
										select case ucase(sLocalAction)
												Case "LOOP"
													If bIfFlagProcess = true Then            	'V4.2
														'Break down the params
														sParams = replace(sParams,chr(34),"")											'V3.1
														aPostProc = split(sParams,",")

														If  fFrame_GetGlobalVar(aPostProc(0)) = ""  or  fFrame_GetGlobalVar(aPostProc(0)) = "LOOPENDED" Then
																'Check if start/end/step exist as script variables
																'Start
																If ubound(aPostProc) > 0 Then													'V3.1						
																	iLocalStart = fFrame_GetGlobalVar(aPostProc(1))								'V3.1
																	if len(iLocalStart) = 0 then iLocalStart = aPostProc(1)						'V3.1
																End If																			'V3.1
																'End
																If ubound(aPostProc) > 1 Then													'V3.1
																	iLocalEnd = fFrame_GetGlobalVar(aPostProc(2))								'V3.1
																	if len(iLocalEnd) = 0 then iLocalEnd = aPostProc(2)							'V3.1
																End If																			'V3.1
																'Step
																If ubound(aPostProc) > 2 Then													'V3.1
																	If aPostproc(3) <> "**~**" Then												'V3.1
																	    iLocalStep = fFrame_GetGlobalVar(aPostProc(3))								'V3.1
																	    if len(iLocalStep) = 0 then iLocalStep = aPostProc(3)						'V3.1  
																	else																		'V3.1
																		iLocalStep = 1																'V3.1
																	End if																		'V3.1
																End If																			'V3.1
																
																'Set global var to hold loop variable
																If isnumeric(trim(aPostProc(0))) = true Then																																																																			'V5.0
																	reporter.ReportEvent micFail,"Iteration[" &  iIterationLoop -1 & "]  Step Name: " & aExec(2,iExecLoop) & " Action " & aExec(9,iExecLoop) & " [" & sLocalAction & "]" , "Action " & aExec(9,iExecLoop) & " [" & sLocalAction & "] - FAILED.  Loop script variables cannot be numeric [" & aPostProc(0) & "]." 		'V5.0				
																	sFrame_ExitTest()																																																																									'V5.0
																End If                                  																																																																				'V5.0
															
																call sFrame_SetGlobalVar (aPostProc(0),iLocalStart)									'V3.1
																call sFrame_SetGlobalVar (aPostProc(0) & "_End",iLocalEnd)							'V3.1
																call sFrame_SetGlobalVar (aPostProc(0) & "_Top",iExecLoop)					
																call sFrame_SetGlobalVar (aPostProc(0) & "_Start",iLocalStart)						'V3.1
																
																'Insert Loop count into ALm log
'																oStep.Field("ST_ACTUAL") = replace(oStep.Field("ST_ACTUAL") ,"] - Control","] [" & iLocalStart & "] - Control")    'V4.3
'																oStep.Field("ST_ACTUAL") = oStep.Field("ST_ACTUAL") & chr(10) & mid(oStep.Field("ST_ACTUAL"),1,instrrev(oStep.Field("ST_ACTUAL")," - Control")) & " [" & iLocalStart & "] - Control Logic"
																oStep.Field("ST_ACTUAL") = oStep.Field("ST_ACTUAL") & chr(10) &  "[" & aPostProc(0) & " = " & iLocalStart & "] - Control Logic"
																'oStep.post		V4.6																									'V4.3
																
																'Step 
																If ubound(aPostProc) > 2 Then
																	call sFrame_SetGlobalVar (aPostProc(0) & "_Step",iLocalStep)					'V3.1
																End if
														else
																'Increase the loop count
'																call sFrame_SetGlobalVar (aPostProc(0), int(replace(fFrame_GetGlobalVar(aPostProc(0)),chr(34),"")) + int(aPostProc(0) & "_Step"))  
																call sFrame_SetGlobalVar (aPostProc(0), int(replace(fFrame_GetGlobalVar(aPostProc(0)),chr(34),"")) + iLocalStep) 									'V3.1 
																
																'Insert Loop count into ALm log
'																oStep.Field("ST_ACTUAL") = oStep.Field("ST_ACTUAL") & mid(oStep.Field("ST_ACTUAL"),1,instrrev(oStep.Field("ST_ACTUAL")," - Control")) & " [" & fFrame_GetGlobalVar(aPostProc(0)) & "] - Control Logic"
																oStep.Field("ST_ACTUAL") = oStep.Field("ST_ACTUAL") & chr(10) & " [" & aPostProc(0) & " = " & fFrame_GetGlobalVar(aPostProc(0)) & "] - Control Logic"
																'oStep.post			'V4.6																								'V4.3
														End If
													End if                     	'V4.2
												Case "ENDLOOP"
'																'Compare loop count to end value
'																If int(replace(fFrame_GetGlobalVar(aPostProc(0)),chr(34),"")) < int(replace(fFrame_GetGlobalVar(aPostProc(0) & "_End"),chr(34),"")) then
'																		'Jump back to 'Loop'
'																		iExecLoop =  fFrame_GetGlobalVar(aPostProc(0) & "_Top") -1 
														If bIfFlagProcess = true Then             	'V4.2

																'Break down the params
																sParams = replace(sParams,chr(34),"")											'V3.1
																aPostProc = split(sParams,",")

																'Compare loop count to end value
																If int(replace(fFrame_GetGlobalVar(aPostProc(0)),chr(34),"")) < int(replace(fFrame_GetGlobalVar(aPostProc(0) & "_End"),chr(34),"")) and int(replace(fFrame_GetGlobalVar(aPostProc(0) & "_End"),chr(34),"")) > int(replace(fFrame_GetGlobalVar(aPostProc(0) & "_Start"),chr(34),""))then
																		'Jump back to 'Loop'
																		iExecLoop =  fFrame_GetGlobalVar(aPostProc(0) & "_Top") -1 
																		'Set booolean flag to prevent a new run being started
																		If iExecLoop = 0 Then							'V4.3
																			bJumpToTopOfStep = true						'V4.3
																		End If                                          'V4.3
																'Compare loop count to end value
																elseIf int(replace(fFrame_GetGlobalVar(aPostProc(0)),chr(34),"")) > int(replace(fFrame_GetGlobalVar(aPostProc(0) & "_End"),chr(34),"")) and int(replace(fFrame_GetGlobalVar(aPostProc(0) & "_End"),chr(34),"")) < int(replace(fFrame_GetGlobalVar(aPostProc(0) & "_Start"),chr(34),"")) then
																		'Jump back to 'Loop'
																		iExecLoop =  fFrame_GetGlobalVar(aPostProc(0) & "_Top") -1 
																		'Set booolean flag to prevent a new run being started
																		If iExecLoop = 0 Then							'V4.3
																			bJumpToTopOfStep = true						'V4.3
																		End If                                          'V4.3
																else
																		'Reset  the loop vars back to 0 
																		call sFrame_SetGlobalVar (aPostProc(0),"LOOPENDED")
																		call sFrame_SetGlobalVar (aPostProc(0) & "_End",0)
																		call sFrame_SetGlobalVar (aPostProc(0) & "_Top",0)
																end if
														End if              	'V4.2
												Case "ELSE"
														If iNested = 0 Then            	'V4.2
																If bIfFlagProcess = True Then
																		bIfFlagProcess = false
																else
																		bIfFlagProcess = true
																End If
														End if            	'V4.2
												Case "ENDIF"
														If iNested = 0 Then       	'V4.2
																'Reset bIfFlag
																bIfFlagProcess = true
														Else  	'V4.2
															iNested = iNested -1      	'V4.2
														End if       	'V4.2
												end select
								end if		

								iPrevIteration = iIterationLoop
																	
								'End Perf Timer
								'fFrame_EndPerfTimer "Post-command:" & aExec(3,iExecLoop)
						Next
				else
					'No data detected in column A of DDP
					reporter.ReportEvent micDone,"Iteration[" &  iIterationLoop -1 & "] - NO DATA" , "No data detected in Col A of DDP for Excel Row [" & iIterationLoop & "]" 							

				
				End if
	Next
	
	'End Perf Timer
	'fFrame_EndPerfTimer "Execute"
End Function
