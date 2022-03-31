Option Explicit
'################################################################
'#
'#		Object Repo export 
'#
'################################################################

	'Set the report to capture all events
	reporter.Filter = rfEnableAll

	'Set default timeout
	Setting("DefaultTimeOut")= 20000


	'Declarations Variables
	Dim pOR
	Dim oRepository, oTest
	Dim aLoadOR
	Dim aORNav
	Dim iLoadOR, iArrayIndex
	Public gORLookup,aORLookUp,gStack,gStackList
    Dim oExcel,oSheet,iColumns,iRow,oFile, oWorkbook
	Dim  sFileName,sSheetName

	'Debug print
	Dim bDebugPrint : 	bDebugPrint = true
 


	'Declarations Arrays
	Public aOR (),  aGUIOverload(1,2), aORDepth(5)

	'Constants

	'Object Repositories
	Dim sLoadOR : sLoadOR="Object Repositories\FM.tsr,Object Repositories\CM.tsr,Object Repositories\TechGUI.tsr,Object Repositories\Inventory.tsr,Object Repositories\AdminGUI.tsr"

	'Global Relative Paths
	createobject("QuickTest.Application").Folders.Add "[QualityCenter\Resources] Resources\AAF"

	'Load Function Libraries
	Executefile "Functions\Frame.qfl"
	Executefile "Functions\General.qfl"



	'Load the Object Repository Navigation array
	'Export resource file from QC
	If (fFrame_QCGetResource("Object Repositories","ORNav.xls","c:\AAF\Object Repositories\ORNav")) = 0 then exitrun
	'Import file from local drive into array
	If (fFrame_ExcelLoad ("c:\AAF\Object Repositories\ORNav\ORNav.xls", "", aORNav,"C1",true,1,1)) = 0 then exitrun
	'Delete local copy of resource file
	if (fFrame_FileDelete ("c:\AAF\Object Repositories\ORNav\ORNav.xls")) = 0 then exitrun

	'Load Environment Variables

	'Initialize Global Array
	call sFrame_InitializeOR
	Call sFrame_InitializeGlobalVar

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



	sFileName = inputbox ("Please enter xls filename and location. eg. C:\temp\OR.xls","File Export","C:\Temp\ObjectRepo.xls")

	If len(sFileName) = 0 Then
		msgbox "No File Created"
		exittest
	End If








	sSheetName = "OR Export "

    Set oExcel = CreateObject("Excel.Application")
     oExcel.Visible =false

	 If  fFrame_FileExists(sFileName) = true Then
			sFileName = replace(sFileName,".xls","") & "_" &year(now) & month(now) & day(now) & "_" & hour(now) & minute(now) & second(now) & ".xls"
	end if

	oExcel.Workbooks.Add
    oExcel.ActiveWorkbook.SaveAs (sFileName)
    Set oWorkbook = oExcel.Workbooks.Open(sFileName)
        
     'Add a sheet
     oWorkbook.Worksheets.Add
        
     'set sheet active
    set oSheet = oWorkbook.ActiveSheet

     'rename sheet
     On error resume next
          oSheet.Name = sSheetName
          If Err.Number <> 0  Then
               oSheet.Name = sSheetName & "_" & hour(now) & minute(now) & second(now)
          end if
     On error goto 0

	'Headers
	oSheet.cells(1,1).value = "Name"
	oSheet.cells(1,2).value = "Type"
	oSheet.cells(1,3).value = "Label"
	oSheet.cells(1,5).value = "Tree P0"
	oSheet.cells(1,6).value = "Tree P1"
	oSheet.cells(1,7).value = "Tree P2"
	oSheet.cells(1,8).value = "Tree P3"
	oSheet.cells(1,9).value = "Tree P4"



	'Loop through aOR
	For iArrayIndex = 1 to aOR(1,0)

			oSheet.cells( iArrayIndex +1,1).value =aOR(1, iArrayIndex)
			oSheet.cells( iArrayIndex +1,2).value = aOR(2, iArrayIndex)
			oSheet.cells( iArrayIndex +1,3).value =aOR(3, iArrayIndex)
			oSheet.cells( iArrayIndex +1,5).value = aOR(6, iArrayIndex)
			oSheet.cells( iArrayIndex +1,6).value = aOR(7, iArrayIndex)
			oSheet.cells( iArrayIndex +1,7).value =aOR(8, iArrayIndex)
			oSheet.cells( iArrayIndex +1,8).value =aOR(9, iArrayIndex)
			oSheet.cells( iArrayIndex +1,9).value = aOR(10, iArrayIndex)
	Next
                        
     'format sheet
     oExcel.columns.autofit
     oExcel.columns.autofilter

     oWorkbook.Save
     oWorkbook.Close 'new line
     oExcel.Quit
    Set oExcel = Nothing


	msgbox ("Object Repository export completed. The file [" & sFileName & "] has been created")









