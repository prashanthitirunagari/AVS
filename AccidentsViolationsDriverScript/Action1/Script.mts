'################################################################################################################################
'	TEST SCRIPT NAME				:  AccidentsViolationsDriverScript
'
' 	TEST SCRIPT DESCRIPTION		 :  This test script runs all policies specified in test  data request file and adds special conditions to the test data
'
' 	PARAMETERS					  :  (None)
'       
' 	RETURNS						:  (None)
'
' 	ERRORS						 :  (None)
'
'	AUTHOR						 : PrashanthiNandagiri
'
'	ORIGINAL DATE				  : SEP 15 2008
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		     											 R E V I S I O N    H I S T O R Y
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	REVISED DATE	: 	     REVISED BY		 :										CHANGE DESCRIPTION
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	05-May-09						Prashanthi Nandagiri								Added condition to verify and accept the transaction when only MVR condition is updated from 1 point to 2 points
'   29-May-09						Prashanthi Nandagiri								Added condition to accept the transaction if it is not  a Z4AV scenario
'################################################################################################################################

Option Explicit

'Variable declarations used in driver file
Dim rc, ColumnName, NoOfAmmends, EffectiveDateOption
Dim dicPolicyRecord
Dim  ErrorFlag, ErrorMessage, Cnt,  InformationSource,SheetName
Dim TestCaseIteration, TestCaseCount, FunctionIteration
Dim TestCaseColumnName, FunctionName, FunctionParameter,DrvCount, DrvName,DrvFlag,icount
Dim arrVar,arrVar1

On Error Resume Next

'Intialize global variables
InitializeVariables
DrvFlag=0
'Set Environment variables
rc = AV_SetEnvirnomentValues
If rc <> micPass Then
	ExitAction (micFail)
End If

'Intialize generic variables
IntializeAVSVariables
IntializeRepViewVariables

'Intialize Reporter
InitializeReporter
If Err.Number <> 0 Then
	Msgbox "Initialization scripts failed." & Err.Description & vbCrLf & "Test execution terminated", vbExclamation, gbApplicationName
	CloseReporter
	ExitAction (micFail)
End If


'Get Testcase ID Column name
TestCaseColumnName = "TestCaseID"
Err.Clear


'Open database
rc=OpenDatabase
'Error handling
If rc<>micPass Then
	Msgbox "Unable to establish database connection. Test execution terminated"
	CloseReporter
	ExitAction (rc)
End If


''This Statement decides type of driver script has to run
If gbApplicableScript = "Complex Scenarios" Then
	gbTestLabName = "ComplexScenario"
Else
	gbTestLabName = "SimpleScenario"
End If

If gbExecutionFlow = "Fetch Policies and Execute"  OR gbExecutionFlow = "Fetch Policies" Then
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible= True
	Set objWorkbook = objExcel.Workbooks.Open(gbServerPath & "Applications\" & gbApplicationName & "\Testware\" & "AVS Get Policies.xls")
	Set objWorksheet = objExcel.Worksheets("Sheet1")
	objWorksheet.Activate
	objExcel.run "Sheet1.FetchPoliciesAndDriverName",gbServerPath & "Applications\" & gbApplicationName & "\Testware\" & "AVSTestware.xls",gbDBUserName, gbDBUserPassword,gbAppRegion,gbTestLabName
	If objWorksheet.Var <> "" Then
		Msgbox "Unable to fetch policies from DB2 since " & objWorksheet.Var
        ExitAction (micFail)
	End If
	objWorkbook.Close("AVS Get Policies.xls")
	If  gbExecutionFlow = "Fetch Policies" Then
		rc = ImportDataRequestSheet
		If rc = micPass Then rc= AV_ReservePolicies()
		If rc = micPass Then AV_ComplexScenarioAmendments()
		DataTable.ExportSheet gbServerPath & "Applications\" & gbApplicationName & "\Testware\" & "AVSTestware.xls",gbTestLabName

        CloseBrowsers		
        ShowHTMLReport
		ExitAction (rc)
	End If
End If

'Connection to DB2		
		Set AV_Connection = CreateObject("ADODB.Connection")
		AV_Connection.Open gbDB2ConnectionSring
		'Check for DB2 connection establishment
		If Err.Number <> micPass and AV_Connection.State =0 Then
			Msgbox "Unable to establish DB2 database connection. Test execution terminated"
			CloseReporter
			ExitAction (rc)
		End If


'Import the master testware from server
rc = ImportDataRequestSheet
'Error Handling
If rc <> micPass Then
	Msgbox "Unable to import testware from " & gbServerPath & "Applications\" & gbApplicationName & "\Testware\" & gbApplicationName & "Testware.xls" & vbCrLf & Err.Description
	CloseReporter
	ExitAction (micFail)
End If

If  gbTestLabName = "SimpleScenario" Then
	If rc = micPass Then rc= AV_ReservePolicies()
ElseIf gbTestLabName = "ComplexScenario" Then
	If gbExecutionFlow = "PLCS CLAIMS Segregation" Then
		rc = AV_TestDataCreation()
		DataTable.ExportSheet gbServerPath & "Applications\" & gbApplicationName & "\Testware\" & "AVSTestware.xls",gbTestLabName
		ExitAction(rc)
	Elseif gbExecutionFlow = "Prepare MVR Data"  Then
        If rc = micPass Then rc = PrepareMVRData()
        DataTable.ExportSheet gbServerPath & "Applications\" & gbApplicationName & "\Testware\" & "AVSTestware.xls",gbTestLabName
		ExitAction(rc)
	End If
End If


	'Get test case count
	TestCaseCount  = Datatable.GetSheet(gbTestLabName).GetRowCount
	
	'Iterate thru the test cases
	For TestCaseIteration =1 To TestCaseCount
		ErrorFlag=False
		ErrorMessage="Unable to execute test case."
		Err.Clear
		
		gbTestCaseStatus=micPass
		
		'Get test case details
		Datatable.GetSheet(gbTestLabName).SetCurrentRow TestCaseIteration
		gbCurrTestCaseName = Datatable.Value(TestCaseColumnName, gbTestLabName)
		'Verify test case is present 
		If Ltrim(rtrim(gbCurrTestCaseName)) = vbNullString Then Exit For

		gbCurrTestCaseDesc = vbNullString
		gbCurrTestCaseDesc = gbCurrTestCaseDesc & "<BR><b>" & Trim(Datatable.Value("TestCaseDescription", gbTestLabName))
		gbTestCaseType =  Trim(Datatable.GetSheet(gbTestLabName).GetParameter("TestCaseType"))
		gbPolicyNum = Trim(Datatable.GetSheet(gbTestLabName).GetParameter("PolicyNo").Value)
		gbPolicyNumber = gbPolicyNum
		gbDriverName = Trim(Datatable.GetSheet(gbTestLabName).GetParameter("DriverFullName").Value)
        arrVar=Split(gbDriverName,";")
		If Ubound(arrVar) = 1 Then
			gbDriverFirstName = arrVar(0)
			gbDriverLastName = arrVar(1)
			gbDriverName = gbDriverFirstName&" "&gbDriverLastName
		End If
		If gbTestLabName = "ComplexScenario" Then
			gbLossDate= Trim(Datatable.GetSheet(gbTestLabName).GetParameter("LossDate").Value)
			gbLossNumber = Trim(Datatable.GetSheet(gbTestLabName).GetParameter("LossNumber").Value)
			gbLossNumber = Replace(gbLossNumber,"-","")
			arrVar1 = Split(gbLossNumber,";")
			If Ubound(arrVar1)=1 Then
				gbClaimsLossNumber = arrVar1(0)
				gbPLCSLossNumber = arrVar1(1)
			End If
			gbMVRLastName = Trim(Datatable.GetSheet(gbTestLabName).GetParameter("MVRDriverLastName").Value)
			gbLicenseNumber = Trim(Datatable.GetSheet(gbTestLabName).GetParameter("LicenseNumber").Value)
		End If
		
		gbApplicableState = Trim(Datatable.GetSheet(gbTestLabName).GetParameter("State").Value)
		If instr(gbApplicableState,"NJ")>0 Then
			gbAbbreviatedStateCode = "NJ"
		Else
			gbAbbreviatedStateCode = gbApplicableState
		End If
		
		EffectiveDateOption = "Any Date"

   If  gbPolicyNum <> VBNullString Then

        'Condition to check which functionality sheet needs to be selected
		If InStr(1,gbCurrTestCaseName,"FG") > 0  Then
			SheetName = "Forgiveness"
		ElseIf InStr(1,gbCurrTestCaseName,"DUP") > 0  Then
			SheetName = "PLCS Dup Same Day Test Data"
		ElseIf InStr(1,gbCurrTestCaseName,"DIFF") > 0  Then
			SheetName = "PLCS Dup Diff Day Test Data"
		ElseIf InStr(1,gbCurrTestCaseName,"PLCS") > 0  Then
			SheetName = "PLCS"
		ElseIf InStr(1,gbCurrTestCaseName,"Z4AV") > 0  Then
			SheetName = "Z4AV"
		Else
			SheetName = "Rating Hierarchy"
		End If
	
		Datatable.GetSheet(SheetName).SetCurrentRow AV_GetTestScriptRow(SheetName)
		'Error handling
		If Err.Number <> micPass Then
			ErrorFlag = True
			ErrorMessage = ErrorMessage & ". Test case details missing in "&SheetName&" sheet of TestWare"
			Err.Clear
			
		End If

	  
		'Close browsers
		  CloseBrowsers
	
		'Start Test case report
		StartTestCaseReport
		
		'Invoke and Login to RepView application
		If ErrorFlag <> True Then
			 rc = AV_SetupRepView()
			'Error handling
			If rc <> micPass Then
				ErrorFlag = True
				ErrorMessage = ErrorMessage & ". Unable to open the RepView application"
			End If
		'End If
	
		'Navigate the applicationn till Ammendments
	   If rc = micPass Then rc = AV_NavigateToAmendments((EffectiveDateOption))
		  'Error handling
			If rc <> micPass Then
				ErrorFlag = True
				ErrorMessage = ErrorMessage & ". Unable to Navigate to Ammendments Page"
			End If
		End If

	  If gbTestCaseType = "Amendment" and gbApplicableScript = "Simple Scenarios" Then
		If rc = micPass Then rc = AV_Ammendments(SheetName)
	  End If		
		
		If ErrorFlag <> True Then
			If SheetName = "Forgiveness" Then
				'Sets the pre-requisites for forgiveness
				If rc = micPass Then rc= AV_FGPreReqDatesVerification( )
			End if
			NoOfAmmends = Datatable.GetSheet(SheetName).GetParameter("Counter")
			If rc = micPass Then rc = AV_DataAssignedToDic(SheetName,NoOfAmmends)
		   If rc=micPass Then rc =AV_AddAccidentOrViolation()
		End if
		
	'Error Handling	
	If ErrorFlag <> True Then
		'Clicks on "OK" button on Accidents Summary Screen
		If rc = micPass Then rc = ClickOnAccidentViolationButton()
		'Select a Policy Tab
		If rc=micPass Then rc = SelectRepviewTab("PolicyQuoteSummary Tabs",  "Policy")
		If rc=micPass Then rc = WaitForWindow(gbBrowserName,gbApplicationHomePage,"PolicyQuoteSummary Tabs","")

		If (SheetName = "Z4AV" and gbClaimsOrPLCSFlag= True) or (gbSimpleSourceFlag = False) Then
			'Do nothing i.e. Accept transaction should not run
        Else
			If rc=micPass Then rc =AcceptTransaction(frPolicyAmendment,"Cancel","")
		End if
		If  (SheetName = "Z4AV" and gbMVRFlag = True) Then
			If rc=micPass Then rc =AcceptTransaction(frPolicyAmendment,"Cancel","")
		End if

		If  (MVRUpdateFlag = True) and Instr(1,gbApplicableState,"CA") > 0 and (gbSimpleSourceFlag = False) and SheetName <> "Z4AV" Then
			If rc=micPass Then rc =AcceptTransaction(frPolicyAmendment,"Cancel","")
		End if
		
		If SheetName <> "Z4AV" Then
			If gbTestLabName = "ComplexScenario" Then
				If gbSimpleSourceFlag = True Then
					'Call Dairy function
					If rc=micPass Then rc =Emulator_diaryRun(gbCurrTestCaseName,"APS",  gbAppRegion,  gbAppUserName,  gbAppUserPassword, gbPolicyNum ,"SCHEDULE;AUS;RENEWAL")
				Else
					'Call Dairy function
					If rc=micPass Then rc =Emulator_diaryRun(gbCurrTestCaseName,"APS",  gbAppRegion,  gbAppUserName,  gbAppUserPassword, gbPolicyNum ,"AUS;RENEWAL")
				End If
			Elseif gbTestLabName = "SimpleScenario" Then
				'Call Dairy function for Simple Scenario only
				If rc=micPass Then rc =Emulator_diaryRun(gbCurrTestCaseName,"APS",  gbAppRegion,  gbAppUserName,  gbAppUserPassword, gbPolicyNum ,"SCHEDULE;MVR;SURCHARGE;AUS;RENEWAL")
				'If rc=micPass Then rc =Emulator_diaryRun(gbCurrTestCaseName,"APS",  gbAppRegion,  gbAppUserName,  gbAppUserPassword, gbPolicyNum ,"SCHEDULE;RENEWAL")
		   End If
		
			'Verify the Source points for acci and viol after renewal
			If rc=micPass Then rc= AV_VerificationAfterRenewal()
	
		'Dairy run for Z4AV MVR conditions
		Elseif SheetName = "Z4AV" Then
			' Verify in the backend for Z4AV test cases
			If rc=micPass Then rc =AV_FetchDriverRecordPointsFromDB2()
		End If  'end of Z4AV condition
		
		'Remove all the keys from the dictionalry object
		 gbdicAVtmpObject.RemoveAll
	
		'Logout of the application
		ReportStep 1,"End the session","",""
		If rc=micPass Then rc = AV_EndSession()
	
		'Close Browsers
		CloseBrowsers
		
	 Else
		HandleRC micFail
		If DrvFlag <> 0  Then
			ReportStep 1,"Pre-Requiste Issue","The Policy should have atleast one driver with no Accidents and Violations added","The Policy Number <b class=""highlight"">" &  gbPolicyNum & " </b> doesn't have such drivers."
		Else
			ReportStep 1, gbCurrTestCaseName, "Execute the test case", ErrorMessage
		End If

	End If
		
		'Complete test case leve report
		EndTestCaseReport

	Else 'Else if gbPolicyNum = VBNULLString
       CloseBrowsers
		StartTestCaseReport
		HandleRC micFail
		ReportStep 1, gbCurrTestCaseName, "Policy should fetch for the testcase ID", "Policy didnot fetch for testcaseID:" & gbCurrTestCaseName
        EndTestCaseReport
	End if 'End if gbPolicyNum = VBNULLString
   
	Next
'End If 		'end of Driver script 

'Release policies allocated in TDM for the current user  --- removed as releasing policies will make TDMTempTables and System (DEb2) tables inconsistent
 'rc = AV_ReleasePolicies()

'Export the Runtime Table
DataTable.ExportSheet gbServerPath & "Applications\" & gbApplicationName & "\Testware\" & "AVSTestware.xls",gbTestLabName


''Close Browsers
CloseBrowsers

'Close Reporter
CloseReporter

'Close data base
CloseDatabase

'Closes DB2 data base
CloseDB2Database

'Opens the HTML report
ShowHTMLReport



