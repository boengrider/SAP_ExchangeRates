Class RateUpload_v1

	Private oFSO
	Private oWSH
	Private oNET
	Private oSES ' Session should be obtained from SapLauncher.GetSession method
	Private strUserName ' System user name e.g. a293793
	Private strComputerName ' System name e.g. SKSENEW128
	
	' ============== Constructor & Destructor ==================
	Private Sub Class_Initialize
	
		Set oFSO = CreateObject("scripting.filesystemobject")
		Set oWSH = CreateObject("wscript.shell")
		Set oNET = CreateObject("wscript.network")
		oSES = Null 
		strUserName = oNET.UserName
		strComputerName = oNET.ComputerName

	End Sub
	
	Private Sub Class_Terminate
	
	End Sub
	
	' ============ P U B L I C  &  P R I V A T E   M E T H O D S  &   S U B R O U T I N E S =========



	' --------- UploadRates
	Public Function UploadRates(strFiles,strExRateType,boolDoNotNEX) ' strFiles is comma delimited list of files to upload,ex rate type ie YHR2, preserve session. Do not cal /NEX
	
		Dim validfrom,SAPfile,i,ratetype,filename
		i = 0
		ratetype = UCase(strExRateType)
	
		For Each SAPfile In Split(strFiles,",")
		
			If oFSO.FileExists(SAPfile) Then 
				filename = oFSO.GetFileName(SAPfile) ' Returns 20200630.txt 
				validfrom = "" ' Clear
				validfrom = Mid(filename,7,2) & "." & Mid(filename,5,2) & "." & Mid(filename,1,4) ' SAP compatible date format DD.MM.YYYY
				KillPopups(oSES)
				oSES.findById("wnd[0]/tbar[0]/okcd").text = "/NZTC_ZCURR_UPLOAD"
				oSES.findById("wnd[0]").sendVKey 0 ' ENTER
				KillPopups(oSES)
				oSES.findById("wnd[0]/usr/txtP_FILE").text = SAPfile
				oSES.findById("wnd[0]/usr/txtP_KURST").text = ratetype
				oSES.findById("wnd[0]/usr/ctxtP_GDATU").text = validfrom
				oSES.findById("wnd[0]").sendVKey 8
				KillPopups(oSES)
				oSES.findById("wnd[0]").sendVKey 0
				KillPopups(oSES)
		
				Do While oSES.Children.Count > 1
					oSES.findById("wnd[0]").sendVKey 0
				Loop
				i = i + 1
				WScript.Sleep 2000 ' Wait a bit
			End If 	
		Next
		
		If Not boolDoNotNEX Then
			oSES.findById("wnd[0]/tbar[0]/okcd").text = "/NEX" ' Close transaction window
			oSES.findById("wnd[0]").sendVKey 0
		End If 
	
		UploadRates = i ' Return the number of uploaded files or 0 if error occured   

	End Function 
	
	
End Class 




Class RateCheck_v1
	
	
	Private oFSO
	Private oWSH
	Private oSES
	Private oFile
	Private oTempFile        ' temp file to hold data from sap TCURR
	Private strHomeCurrency  ' Home currency e.g. CZK
	Private numFilesVerified ' Number of successfully verified files
	Private strTempFilePath  ' Absolute path to the temp file
	Private strTempFileName  ' temp file name
	Private strIncompleteEntries ' Input line count and sap entries count don't match
	Private strMissingEntries ' There is no rate for this day
	Private strInvalidEntries ' These entries line count match, but lines differ i.e rates are invalid
	Private missingEntries 
	Private incompleteEntries
	Private invalidEntries
	
	
	' ============== Constructor & Destructor ==================
	Private Sub Class_Initialize
	
		Set oFSO = CreateObject("scripting.filesystemobject")
		Set oWSH = CreateObject("wscript.shell")
		numFilesVerified = 0
		missingEntries = 0
		incompleteEntries = 0
		invalidEntries = 0
		oSES = Null 
		strMissingEntries = Null
		strIncompleteEntries = Null
		strInvalidEntries = Null
		strTempFilePath = Null
		strTempFileName = Null
		strHomeCurrency = Null

	End Sub
	
	Private Sub Class_Terminate
	
	End Sub
	
	' ============ P U B L I C  &  P R I V A T E   M E T H O D S  &   S U B R O U T I N E S ===========
	
	Public Function Init(str_home_curr)
	
		strHomeCurrency = str_home_curr
		
	End Function 
	
	' ------------ CreateGUID
	
	Private Function CreateGUID
  		Dim TypeLib
  		Set TypeLib = CreateObject("Scriptlet.TypeLib")
  		CreateGUID = Mid(TypeLib.Guid, 2, 36)
	End Function



	' --------- CheckRates
	Public Function CheckRates(strFiles,strExRateType) ' strFiles is comma delimited list of files to check
	
		Dim SAPfile,files
		
		files = Split(strFiles,",") ' Split files and use the first one to determine where to put temp file
		
		
		If strHomeCurrency = Null Or strHomeCurrency = "" Then
			CheckRates = -1 ' ERROR, home currency not set
		End If 
		
		strTempFileName = CreateGUID ' Create a temp file name
		
		strTempFilePath = oFSO.GetParentFolderName(files(0)) ' Get a temp file location
		
	
		For Each SAPfile In Split(strFiles,",")
			
			If oFSO.FileExists(SAPfile) Then 
				Check SAPfile,strExRateType
			End If 
			
		Next
		
		 
		oSES.findById("wnd[0]/tbar[0]/okcd").text = "/NEX" ' Close transaction window
		oSES.findById("wnd[0]").sendVKey 0
		
		If oFSO.FileExists(strTempFilePath & "\" & strTempFileName & ".txt") Then 
		
			oFSO.DeleteFile strTempFilePath & "\" & strTempFileName & ".txt"
			
		End If 
		CheckRates = numFilesVerified ' Returns number of successfully verified files.
		

	End Function 
	
	
	
	
	
	
	
	
	Private Sub Check(strFile,strType) ' Private sub to check files. Call within for loop
	
		Dim lines,filename,gdatu,line,i,sapline,fileline,column,sapentries,j,saplinetrimmed,filelinetrimmed
		i = 0
		lines = 0
		
		If Not oFSO.FileExists(strFile) Then
			Exit Sub 
		End If 
		 
		 Set oFile = oFSO.OpenTextFile(strFile,1,False) ' Open file containing uploaded rates for reading
		 
		 Do While Not oFile.AtEndOfStream
		 	oFile.ReadLine
		 Loop
		 	
		 lines = oFile.Line - 1
		 oFile.Close
		 	
		 	
		filename = oFSO.GetBaseName(strFile) ' Returns 20200630 
		gdatu = 99999999 - filename
		KillPopups(oSES)
		oSES.findById("wnd[0]/tbar[0]/okcd").text = "/nse17"
		oSES.findById("wnd[0]").sendVKey 0 ' ENTER
		KillPopups(oSES)
		oSES.findById("wnd[0]/usr/ctxtDD02V-TABNAME").text = "TCURR"
		oSES.findById("wnd[0]").sendVKey 0 ' ENTER
		KillPopups(oSES)
		' FIELDS
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-FSELECT[1,1]").text = LCase(strType)
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-FSELECT[1,2]").text = ""
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-FSELECT[1,3]").text = ""
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-FSELECT[1,4]").text = gdatu
		' FLAGS
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-SHOWFLAG[2,0]").text = ""  ' CLIENT
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-SHOWFLAG[2,1]").text = ""  ' KURST
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-SHOWFLAG[2,2]").text = "X" ' FCURR
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-SHOWFLAG[2,3]").text = "X" ' TCURR
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-SHOWFLAG[2,4]").text = ""  ' GDATU
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-SHOWFLAG[2,5]").text = "X" ' UKURS
		oSES.findById("wnd[0]").sendVKey 8
		KillPopups(oSES)
		
		If oSES.findById("wnd[0]/sbar/pane[0]").text <> "" Or oSES.findById("wnd[0]/sbar/pane[0]").text = "No values selected in the specified area" Then
			
			missingEntries = missingEntries + 1 
			strMissingEntries = strMissingEntries & " " & strFile
			numFilesVerified = numFilesVerified + 1
			Exit Sub 
		
		End If 
			
		sapentries = oSES.findById("wnd[0]/usr/lbl[23,3]").text ' Number of entries. Compare this to the input file line count
		
		If CInt(sapentries) <> lines Then
		  
			incompleteEntries = incompleteEntries + 1 
			strIncompleteEntries = strIncompleteEntries & " " & strFile
			numFilesVerified = numFilesVerified + 1
			Exit Sub 
			
		End If 
		
		' Continue with complete entries. Generate output file from SAP
		oSES.findById("wnd[0]/mbar/menu[5]/menu[5]/menu[2]/menu[1]").select
		oSES.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
		oSES.findById("wnd[1]").sendVKey 0
		KillPopups(oSES)
		oSES.findById("wnd[1]/usr/ctxtDY_PATH").text = strTempFilePath 'directory
		oSES.findById("wnd[1]/usr/ctxtDY_FILENAME").text = strTempFileName & ".txt"
		oSES.findById("wnd[1]").sendVKey 11
		KillPopups(oSES)
	
		Set oFile = oFSO.OpenTextFile(strFile,1,False) ' Open input file for reading
		Set oTempFile = oFSO.OpenTextFile(strTempFilePath & "\" & strTempFileName & ".txt") ' Open sap generated file for reading
		j = 0
		Do While j < 9
			oTempFile.SkipLine
			j = j + 1
		Loop 
		
		Do While Not oFile.AtEndOfStream
		
			
			sapline = Split(oTempFile.ReadLine,vbCrLf)
			fileline = Split(oFile.ReadLine,vbCrLf)
			column = Split(fileline(0),vbTab)
			saplinetrimmed = Replace((Trim(sapline(0))),vbTab,"")
			filelinetrimmed = Replace((Trim(column(0) & column(1) & column(2))),vbTab,"")
			
			If Replace(saplinetrimmed," ","") <> filelinetrimmed Then
			
				invalidEntries = invalidEntries + 1
				numFilesVerified = numFilesVerified + 1
				strInvalidEntries = strInvalidEntries & " " & strFile
				oFile.Close
				oTempFile.Close
				Exit Sub 
		
			End If 
			
			
		Loop
		
		
		numFilesVerified = numFilesVerified + 1
		oTempFile.Close ' Close the temp file
		oFile.Close     ' Close the rate file
		
	End Sub  	
	
	
				
	Public Function KillPopups(ByRef objSession)
		Do While objSession.Children.Count > 1
			If InStr(objSession.ActiveWindow.Text, "System Message") > 0 Then
				objSession.ActiveWindow.sendVKey 12
			ElseIf InStr(objSession.ActiveWindow.Text, "Information") > 0 And InStr(objSession.ActiveWindow.PopupDialogText, "Exchange rate adjusted to system settings") > 0 Then
				objSession.ActiveWindow.sendVKey 0
			ElseIf InStr(objSession.ActiveWindow.Text, "Copyright") > 0 Then
				objSession.ActiveWindow.sendVKey 0
			ElseIf InStr(objSession.ActiveWindow.Text, "License Information for Multiple Logon") > 0 Then
				objSession.ActiveWindow.findById("usr/radMULTI_LOGON_OPT2").select
				objSession.ActiveWindow.sendVKey 0
			'ElseIF   'Insert next type of popup windows which you want to kill
			Else
				Exit Do
			End If
		Loop
	End Function 
	
		

	' ================= P R O P E R T I E S ====================
	Public Property Let SAPSession(s)
		Set oSES = s
	End Property 	
	
	Public Property Get FilesVerified
		FilesVerified = numFilesVerified
	End Property 
	
	Public Property Get GetMissingEntries
		GetMissingEntries = strMissingEntries
	End Property  
	
	Public Property Get GetIncompleteEntries
		GetIncompleteEntries = strIncompleteEntries
	End Property 
	
	Public Property Get GetNumMissingEntries
		GetNumMissingEntries = missingEntries
	End Property
	
	Public Property Get GetNumIncompleteEntries
		GetNumIncompleteEntries = incompleteEntries
	End Property 
	
	Public Property Get GetNumInvalidEntries
		GetNumInvalidEntries = invalidEntries
	End Property 
	
	Public Property Get GetInvalidEntries
		GetInvalidEntries = strInvalidEntries
	End Property 
	
End Class 
