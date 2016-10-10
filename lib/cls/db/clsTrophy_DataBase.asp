<% 


class Trophy_DataBase

	Private updateQueryCount
	Private viewQueryCount
	Private datasourceObject
	Private debugMessageFormat
	Private autoCloseRecordset
	Private preparedQuery
	Private commandObject
	Public objRs
	Public debugMode
	Public lastQuery

	Public Property Let datasource(d)
		Set datasourceObject = d
	End Property
	
	Public Property Let messageFormat(f)
		debugMessageFormat = f
	End Property
	
	Public Property Let closeRecordsetBefore(r)
		autoCloseRecordset = CBool(r)
	End Property
	
	Public Property Get isRecordsetAvailable
		Dim r
		if objRs.State <>  0 then
			r = false
		else
			r = true
		end if
		isRecordsetAvailable = r
	End Property
	
	Public Property Get command
		Set command = commandObject
	End Property
	
	Public Property Get query
		query = preparedQuery
	End Property
	
	Public Property Let query(q)
		Set commandObject.ActiveConnection = datasourceObject.connection
		preparedQuery = q
		commandObject.CommandText = preparedQuery
		commandObject.Prepared = true
		commandObject.CommandType = adCmdText
	End Property
	
	Public Sub Class_Initialize()
		Call createCommandObject()
		Call createRecordsetObject()
    	debugMessageFormat = "<hr>SQL:<b>%QUERY</b><hr>"
		debugMode = false
		autoCloseRecordset = true
	End Sub
	
	Public Sub class_terminate()
		Set objRs = Nothing
		Set commandObject = Nothing
	End Sub
	
	Private Function createRecordsetObject()
		if Not isObject(objRs) then
			Set objRs = Server.CreateObject("Adodb.Recordset")
		end if
    End Function
	
	Private Function createCommandObject()
		if Not isObject(commandObject) then
			Set commandObject = Server.CreateObject("Adodb.Command")
		end if
	End Function
	
	Private Function sqlOutput()
		echo replace(debugMessageFormat,"%QUERY",lastQuery)
	End Function

	Private Function checkDatasource()
		if Not datasourceObject.isOpen then
			Call datasourceObject.Open()
		end if
	End Function
	
	Private Function runQueryOnRs(ByRef rs,q)
	   Call checkDatasource()
	   if autoCloseRecordset then
	   		if Not isRecordsetAvailable then
				objRs.Close
			end if
	   end if
	   viewQueryCount = viewQueryCount+1
	   lastQuery = q
	   if debugMode then
	   		Call sqlOutput()
	   end if
	   objRs.CursorLocation = 3
       objRs.CursorType = 3	
	   objRs.ActiveConnection = datasourceObject.connection
       objRs.open lastQuery
	End Function
	
	Public Function open()
		datasourceObject.Open
	End Function
	
	Public Function close()
		datasourceObject.Close
	End Function
	
	Private Function addCommandParam(name,ptype,direction,maxsize,value)
		Call commandObject.Parameters.Append(commandObject.CreateParameter(name,pType,adParamInput,maxsize,value))
	End Function
	
	Private Function addCommandParam_Input(name,ptype,maxsize,value)
		Call addCommandParam(name,ptype,adParamInput,maxsize,value)
	End Function
	
	Public Function addString(name,value)
		Call addCommandParam_Input(name,adVarChar,255,value)
	End Function
	
	Public Function addInteger(name,value)
		Call addCommandParam_Input(name,adInteger,-1,value)
	End Function
	
	Public Function addDate(name,value)
		Call addCommandParam_Input(name,adDate,-1,value)
	End Function
	
	Public Function addBoolean(name,value)
		Call addCommandParam_Input(name,adBoolean,-1,value)
	End Function
	
	Public Function addDouble(name,value)
		Call addCommandParam_Input(name,adDouble,-1,value)
	End Function
	
	Public Function addText(name,value)
		Call addCommandParam_Input(name,adLongVarChar,16,value)
	End Function
	
	Public Function addBinary(name,value)
		Call addCommandParam_Input(name,adLongVarBinary,16,value)
	End Function
	
	Public Function execQuery(ByVal query)
		Call checkDatasource()
		lastQuery = query
		updateQueryCount = updateQueryCount+1
		execQuery = runQuery(query)
	End Function
	
	Public Function execute(q)
		Call runQueryOnRs(objRs,q)
	End Function
	
	Public Function executePreparedQuery()
		Set objRs = commandObject.Execute
	End Function
	
	Public Function runQuery(ByVal Q)
	   Call runQueryOnRs(objRs,q)
	End Function
	
	Public Function runQueryP(ByVal q,ByVal params)
		
	End Function

	Public Function returnQuery(ByVal q)
	   Set rs = Server.CreateObject("Adodb.Recordset")
	   Call runQueryOnRs(rs,q)
	   Set returnQuery = rs
	End Function

	Public Function runQueryForUpdate(ByVal Q)
       viewQueryCount = viewQueryCount+1
	   lastQuery = Q
	   if debugMode then
	   		Call sqlOutput()
	   end if
	   Call checkDatasource()
       objRs.open Q, datasourceObject.connection, 1,3
	End Function
	
	Public Function setPageParameters(currentPage,pageSize)
		Call setPage(currentPage)
		Call setPageSize(pageSize)
	End Function
	
	Public Function setPageSize(ByVal size)
       objRs.PageSize = size
	   objRs.CacheSize = size
	End Function

	Public Function setPage(ByVal page)
       if Not objRs.EOF then
	   	objRs.AbsolutePage = page
	   end if
	End Function
	
	Public Function fetch(ByVal start,ByVal numRows)
       fetch = objRs.GetRows(numRows,start)
	End Function
	
	Public Function cloneRecordset()
		Set cloneRecordset = objRs.Clone(-1)
	End Function
	
	Public Function getLastID()
       dim ID
	   Call runQuery("SELECT @@IDENTITY AS ID")
	   ID = objRs("ID")
	   objRs.Close()
	   getlastID = ID
	End Function
	
End Class 
%>