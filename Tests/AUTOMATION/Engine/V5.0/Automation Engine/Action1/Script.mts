Option Explicit
'################################################################
'#
'#		Main Processing Engine for the Amadeus Automation Framework
'#      V5.7
'################################################################

	'Set the report to capture all events
	reporter.Filter = rfEnableAll

	'Set default timeout
	Setting("DefaultTimeOut")= 20000


	'Declarations Variables
	Dim iSyncTime : iSyncTime = 180
	Dim iGUIOverloadSyncTime : iGUIOverloadSyncTime = 5   
	Dim iGUIOverloadSyncTimeDefault : iGUIOverloadSyncTimeDefault = 5     'V5.5
	
	Dim gCallStack : gCallStack = "" 			'V5.7
	Public gUseLiteralValues : gUseLiteralValues = true 'V5.7

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

	Public gCapture: gCapture = false
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
	gALMProject = qcutil.QCConnection.ProjectName
	gALMUsername = qcutil.QCConnection.UserName

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
'		Executefile xxLoadDevLib
		Executefile "Functions\Frame.qfl"
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
		sLoadOR="Object Repositories\FM.tsr,Object Repositories\CM.tsr,Object Repositories\TechGUI.tsr,Object Repositories\Inventory.tsr,Object Repositories\AdminGUI.tsr,Object Repositories\FMWeb.tsr"
		
		
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
'gRunType = "TestCase" '"TestCase"  '"TestSet"
'****************************************
		'grab the RunType from QC to see if a complete test set is being run. Or just a single test case
		If  gRunType  = "TestSet" Then

			'Start Perf Timer
			'fFrame_StartPerfTimer "Load TestSet Attachments", "AAF", 0	

'****************************************
'			 iTestSetID = xxTestSetId '201 '101
'****************************************
			 iTestSetID = oCurrentTestSet.ID   '101
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
					If  ucase(mid(replace(oTestSetAttachmentList.item( iTestSetAttachmentLoop).name,"CYCLE" & "_" & iTestSetID & "_",""),1,3)) = "DDP" and (sTestSetAttachmentType = ".XLS" or sTestSetAttachmentType = ".XLSX" or sTestSetAttachmentType = ".XLSM") Then
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
					
					'Parse the steps
					fFrame_LoadParse  aExec,aSteps,  aActions, sTestName
					
					'Execute all steps
					fFrame_Execute  aExec, bDebugPrint, aFunctions
					
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
			
								For each oParam in oParamList
										print oParam.Name
										print oParam.DefaultValue
										print oParam.ActualValue

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
					
					'Parse the steps
					fFrame_LoadParse  aExec,aSteps,  aActions, sTestName
					
					'Execute all steps
					fFrame_Execute  aExec, bDebugPrint, aFunctions

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
'Set oTestCaseF =  oQCConn.TestFactory
'Set oTestCaseObj = oTestCaseF.Item(xxTestCaseId)  '237 '109   '216	'160
'Set oStepList  = oTestCaseObj.DesignStepFactory.NewList("")
'iTestCaseId = xxTestCaseIdInTestSet  '117  '126
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			Set oStepList  = QCUtil.CurrentTest.DesignStepFactory.NewList("")
			iTestCaseId = qcutil.CurrentTestSetTest.ID

			'Check and load Pre-Reqs
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~			
			fFrame_SpreadsheetSteps qcutil.CurrentTest, aSteps, "PRE-REQ",iTestCaseId,0,"1-1"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~			
'			fFrame_SpreadsheetSteps oTestCaseObj, aSteps, "PRE-REQ",iTestCaseId,0,"1-1"
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
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~			
			fFrame_SpreadsheetSteps qcutil.CurrentTest, aSteps, "POST-REQ",iTestCaseId,iCurrentSteps,"1-1"       
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'			fFrame_SpreadsheetSteps oTestCaseObj, aSteps, "POST-REQ",iTestCaseId,iCurrentSteps,"1-1"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

			'End Perf Timer
			'fFrame_EndPerfTimer "Load Test Steps"

			'Start Perf Timer
			'fFrame_StartPerfTimer "Load Attachments [Test Case in TestLab]", "AAF", 0	
					
			'Attachments
			'Connect to attachment factory
'#Get both the attachments to the test plan and the test case in the test set
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Set oTestSetF =  oQCConn.TestSetFactory
'Set oTestSetObj = oTestSetF.Item(xxTestSetId)
'Set oTestF = oTestSetObj.TSTestFactory
'Set oTestSetTestCaseObj = oTestF.item(xxTestCaseIdInTestSet) '117  '126

'set oTestCaseAttachmentF = oTestSetTestCaseObj.Attachments 'check this works !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Set oTestCaseAttachmentList = oTestCaseAttachmentF.newlist("")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

		set oTestCaseAttachmentF = QCUtil.CurrentTestSetTest.Attachments 
		Set oTestCaseAttachmentList = oTestCaseAttachmentF.newlist("")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 			
			'Loop through all attachments to the test case in the test set
			iNoSheets = 0
			For iTestCaseAttachmentLoop = 1 to oTestCaseAttachmentList.count
					'Check to determine if any meet the naming convention for parameters to drive the test set 
					sTestCaseAttachmentType = ucase(mid(oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,instrrev(oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,".")))
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
'set oTestCaseAttachmentF = oTestCaseObj.Attachments 'check this works !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Set oTestCaseAttachmentList = oTestCaseAttachmentF.newlist("")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			Set oTestCaseAttachmentF = qcutil.CurrentTest.Attachments
			Set oTestCaseAttachmentList = oTestCaseAttachmentF.newlist("")
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			
			'Loop through all attachments to the test case in the test plan
			For iTestCaseAttachmentLoop = 1 to oTestCaseAttachmentList.count
					'Check to determine if any meet the naming convention for parameters to drive the test set 
					sTestCaseAttachmentType = ucase(mid(oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,instrrev(oTestCaseAttachmentList.item( iTestCaseAttachmentLoop).name,".")))
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
			'Set oCurrentTSTest= oTestSetTestCaseObj
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			Set oCurrentTSTest= QCUtil.CurrentTestSetTest
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				If oCurrentTSTest.HasSteps Then

					'If the test case has parameters
					If  oCurrentTSTest.Params.Count > 0 Then

						'Re-size for the number of params
						ReDim aParam(2,oCurrentTSTest.Params.Count)
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

					
			'Parse the steps
			fFrame_LoadParse  aExec,aSteps,  aActions, sTestName
			
			'Execute all steps
			fFrame_Execute  aExec, bDebugPrint, aFunctions
			

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
'#
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
'			print oResourceFilter.NewList(iResourceLoop).field("RSC_NAME") & "    " & oResourceFilter.NewList(iResourceLoop).field("RSC_VTS") & "    DateDiff: " & dateDiff("s",oResourceFilter.NewList(iResourceLoop).field("RSC_VTS"),qcutil.QCConnection.ServerTime)
			
			'Default
			bDownloadResource = false
			
			'Check if the object exists within the Release Manager
			If oALMConn.value(oResourceFilter.NewList(iResourceLoop).field("RSC_FILE_NAME")) <> "" Then
				'ALM version is later version than the local version
				If datediff("s",oResourceFilter.NewList(iResourceLoop).field("RSC_VTS"),oALMConn.value(oResourceFilter.NewList(iResourceLoop).field("RSC_FILE_NAME"))) < 0 Then
					'Download Object
					bDownloadResource = true
				End If
				'Write out status
				reporter.ReportEvent micDone,"Engine" ,"Engine: Test Resource [" & oResourceFilter.NewList(iResourceLoop).field("RSC_NAME") & "] Is out of date and will be downloaded from ALM."

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
				
				'print oResourceFilter.NewList(iResourceLoop).field("RSC_NAME") & " : DOWNLOADED FROM ALM"

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


'###################################


