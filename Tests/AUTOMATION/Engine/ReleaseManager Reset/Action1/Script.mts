		
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
			reporter.ReportEvent micFail,"Release Manager","Release Manager - Fatal Error, 'AAF' resource folder does not exist."
			ExitTest
		elseIf oRootResourceFolderFilter.NewList.count > 1 Then
			reporter.ReportEvent micFail,"Release Manager","Release Manager - Fatal Error, Mulitple 'AAF' resource folders exist."
			ExitTest
		End If


		'Set the object to the \AAF folder
		Set oAAFFolder = oRootResourceFolderFilter.NewList(1)
		
		Call ClearRMValues("Functions")
		Call ClearRMValues("Object Repositories")
		
		
		

'########################## start the function here
Function ClearRMValues (byval sResourceFolder)

Dim bDownloadResource

	
		'Set the object to the resource folder
		Set oFunctionsFolder= oAAFFolder.QCResourceFolderFactory
		Set oFunctionsFilter = oFunctionsFolder.Filter
		oFunctionsFilter("RFO_NAME") = "'" & sResourceFolder & "'"
		
		'Check that the \AAF\sResourceFolder is a unique folder
		If oFunctionsFilter.NewList.count = 0 Then
			reporter.ReportEvent micFail,"Release Manager","Release Manager - Fatal Error, 'AAF\" & sResourceFolder & "' resource folder does not exist."
			ExitTest
		elseIf oFunctionsFilter.NewList.count > 1 Then
			reporter.ReportEvent micFail,"Release Manager","Release Manager - Fatal Error, Mulitple 'AAF\" & sResourceFolder & "' resource folders exist."
			ExitTest
		End If
		
		
		
		Set oResourceFolder = oFunctionsFilter.NewList(1)
		Set oResourceFactory = oResourceFolder.QCResourceFactory
		Set oResourceFilter = oResourceFactory.Filter
		
		'oResourceFilter("RSC_NAME") = "'Command.qfl'"
		Set oResource = oResourceFilter.NewList
		If oResource.Count = 0 Then
			reporter.ReportEvent micFail,"Release Manager","Release Manager - Fatal Error, 'AAF\" & sResourceFolder & "' contains no resources."
			ExitTest
		End If
		
		'Loop through all the resources in the given folder
		For iResourceLoop = 1 To oResource.Count
			
				'Set ReleaseManager Date time
				oALMConn.value(oResourceFilter.NewList(iResourceLoop).field("RSC_FILE_NAME")) = ""
				oALMConn.post
				
				print oResourceFilter.NewList(iResourceLoop).field("RSC_NAME") & " : Data Cleared"
				reporter.ReportEvent micPass,"Release Manager", "Release Manager - " & oResourceFilter.NewList(iResourceLoop).field("RSC_NAME") & " - Reset"
			
		Next

		'Close the release manager connection
		oALMConn.close


End Function
		



