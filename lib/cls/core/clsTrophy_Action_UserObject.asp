<%


class Trophy_Action_UserObject
	
	Private objectName
	Private objectPassAllProperties
	Private objectProperties
	Private objectInstanceName
	
	Public Property Get name
		name = objectName
	End Property
	
	Public Property Let name(n)
		objectName = n
	End Property
	
	Public Property Get passAllProperties
		passAllProperties = objectPassAllProperties
	End Property
	
	Public Property Let passAllProperties(p)
		objectPassAllProperties = CBool(p)
	End Property
	
	Public Property Get properties
		Set properties = objectProperties
	End Property
	
	Public Property Get instanceName
		instanceName = objectInstanceName
	End Property
	
	Public Property Let instanceName(i)
		objectInstanceName = i
	End Property
	
	Public Function setProperty(key,value)
		if Not objectProperties.Exists(key) then
			Call objectProperties.Add(key,value)
		else
			objectProperties.Item(key) = value
		end if
	End Function
	
	Public Function getProperty(key)
		if objectProperties.Exists(key) then
			value = objectProperties.Item(key)
		else
			value = Null
		end if
		getProperty = value
	End Function
	
	Public Sub class_initialize()
		passAllProperties = false
		Set objectProperties = newDictionary()
	End Sub
	
end class




%>