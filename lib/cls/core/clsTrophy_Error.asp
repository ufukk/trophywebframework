<%


class Trophy_Error

	Private reporters
	Private lastError

	Public Sub Class_Initialize()
		Set reporters = newDictionary()
	End Sub
	
	Private Function notifyReporters(message)
		For Each reporter in reporters
			command = "Call reporter." & reporters.Item(reporter) & "(message)"
			Call Execute(command)
		Next
	End Function
	
	Public Function trigger(file,message)
		Call notifyReporters(message)
		Call Err.Raise(vbObjectError,file, message)
	End Function
		
	Public Function hookReporter(reporterObject,method)
		Call reporters.Add(reporterObject,method)
	End Function
		
end class
%>