<%


class Trophy_Page

	Private startupFunctions
	Private shutdownFunctions
	
	Public Property Get hasValidationFields
		Dim r
		if isObject(Session(CONFIG_SESSION_VALIDATOR_KEY))  then
			r = true
		else
			r = false
		end if
		hasValidationFields = r
	End Property
	
	Public Property Get validationFields
		Set validationFields = getValidationFields()
	End Property
	
	Public Sub Class_Initialize()
		Set startupFunctions = newDictionary()
		Set shutdownFunctions = newDictionary()
		isDefault = False
	End Sub
	
	Public Function addStartupFunction(startupFunction)
		if Not startupFunctions.Exists(startupFunction) then
			Call startupFunctions.Add(startupFunction,startupFunctions.Count)
		end if
	End Function
	
	Public Function addShutdownFunction(shutdownFunction)
		if Not shutdownFunctions.Exists(shutdownFunction) then
			Call shutdownFunctions.Add(shutdownFunction,shutdownFunctions.Count)
		end if
	End Function
	
	Public Function runStartup()
		For Each funcName in startupFunctions
			Call Execute(funcName & "()")
		Next
	End Function
	
	Public Function runShutdown()
		For Each funcName in shutdownFunctions
			Call Execute(funcName & "()")
		Next
	End Function
	
	Private Function getValidationFields()
		Set fieldDict = Session(CONFIG_SESSION_VALIDATOR_KEY)
		Call Session.Contents.Remove(CONFIG_SESSION_VALIDATOR_KEY)
		Set getValidationFields = fieldDict
	End Function
	
end class


%>