<%


class Trophy_DataSource

	Public connection
	Private name
	Private url
	
	Public Property Get dataSourceName
		dataSourceName = name
	End Property
	
	Public Property Get dataSourceurl 
		dataSourceurl = url
	End Property
	
	Public Property Let dataSourceName(n)
		name = n
	End Property
	
	Public Property Let dataSourceurl(u) 
		url = u
	End Property
	
	Public Property Get isOpen
		Dim r
		if connection.State <> 0 then
			r = true
		else
			r = false
		end if
		isOpen = r
	End Property
	
	Private Function createConnectionObject()
		Set connection = CreateObject("Adodb.Connection")
	End Function
	
	Public Sub Class_Initialize()
		Call createConnectionObject()
	End Sub
	
	Public Sub Class_Terminate()
		Set connection = Nothing
	End Sub
	
	Public Function open()
		connection.open url
	End Function
	
	Public Function close()
		if isObject(connection) then
			if isOpen then
				connection.close
			end if
		end if
	End Function
	
end class
%>