<%


class Trophy_Page

	Private startupFunctions
	Private shutdownFunctions
	
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

end class
%>