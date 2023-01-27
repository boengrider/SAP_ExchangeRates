Class TCDCalendar
	'v3
	'Utilizes a sharepoint list as s source of TCDs
	'Private member variables 
	Private oRX__ 
	Private oCON__
	Private oRST__
	Private oDICT__ ' Dictionary holding closing days key=day value=name of holiday
	Private strConnectionString__
	Private strList__
	Private daySunday__
	Private daySaturday__
	Private oFirstNTCD__
	
	
	Private Sub Class_Initialize
		
		Set oRX__ = New RegExp
		oRX__.Pattern = "(?:list)=\w*"
		oRX__.Multiline = False
		oRX__.IgnoreCase = True
		oRX__.Global = True
		Set oDICT__ = CreateObject("scripting.dictionary")
		Set oRST__ = CreateObject("adodb.recordset")
		Set oCON__ = CreateObject("adodb.connection")
		strConnectionString__ = ""
		strList__ = ""
		daySaturday__ = 7
		daySunday__ = 1
		
	End Sub 
	
    
	Public Function Init(sConnectionString)
	
	    Dim matches : Set matches = oRX__.Execute(sConnectionString)
	    oRX__.Pattern = "list="
	    strList__ = oRX__.Replace(matches.Item(0),"")
		strConnectionString__ = sConnectionString
		oCON__.ConnectionString = strConnectionString__
		
	
		If Not strList__ = "" And Not IsNull(strConnectionString__) And Not strConnectionString__ = "" Then
		   
		    oCON__.Open
	    	oRST__.Open "SELECT Title, TCD FROM [" & strList__ & "]", oCON__, 3, 3
	        oRST__.MoveFirst
	        
	        Do While Not oRST__.EOF
	        
	        	oDICT__.Add oRST__.Fields("TCD").Value, oRST__.Fields("Title").Value
	        	oRST__.MoveNext
	        	
	        Loop
	        
	    Else
	    
	    	Init = -1
	    	
	    End If
	    	    	
    End Function 
    
    
    Private Function FindFirstNonTcdDate(D)
    
    	oFirstNTCD__ = D
    	    
        If (Weekday(D) = daySaturday__ Or Weekday(D) = daySunday__) Or oDICT__.Exists(D) Then
        	FindFirstNonTcdDate(D - 1) ' Recursive call
        Else
        	oFirstNTCD__ = D
        	Exit Function
        End If 
        
    End Function 
	
	Public Property Get connectionString
	
		connectionString = strConnectionString__
		
	End Property 
	
	Public Property Get list
	   
	    list = strList__
  		
	End Property
	
	Public Property Get firstNTCD(D)
	
		FindFirstNonTcdDate(D)
	
		firstNTCD = oFirstNTCD__
		
	End Property 
	
			
End Class
