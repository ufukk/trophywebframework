<%
if objTrophy.isController = false then
	if objTrophy.selectedAction.isClass then
		classCmd = "Set objAc = new " & objTrophy.selectedAction.className
		Execute(classCmd)
		Call runActionObject(objAc)
		For Each objectName in objTrophy.selectedAction.userObjects
			className = objTrophy.config.userObjects.Item(objectName).className
			instanceName = objTrophy.selectedAction.userObjects.Item(objectName).instanceName
			classCmd = "Set objAc." & instanceName & "= new " & className
			instanceCmd = "Set objTmp = objAc." & instanceName
			Execute(classCmd)
			Execute(instanceCmd)
			if Not isNull(objTrophy.config.userObjects.Item(objectName).dbObjectName) then
				dbObjectInstanceName = objTrophy.config.userObjects.Item(objectName).dbObjectName
				dbObjectCmd = "Set objTmp." & dbObjectInstanceName & " = objTrophy.getDBInstance(" & "objTrophy.config.userObjects.Item(objectName).datasource" & ")"
				Execute(dbObjectCmd)
			end if
			if objTrophy.selectedAction.userObjects.Item(objectName).passAllProperties = true then
				Call setObjectProperties(objTmp,objTrophy.selectedAction.properties)
			end if
			Call setObjectRequestProperties(objTmp,objTrophy.selectedAction.userObjects.Item(objectName).properties)
		Next
		'run action object
		method = objTrophy.selectedAction.method
		Execute("Call objAc." & method & "()")
	end if
	if actionFound = true then
		For Each template in objTrophy.selectedAction.templates
			Call objTemplate.displayFile(template)
		Next
	end if
end if
%>