<%


class Trophy_Action

	Public actionType
	Public isDefault
	Public page
	Private pageClassName
	Private pageMethod
	Private startupFunctions
	Private shutdownFunctions
	Private actionValidator
	Private actionTemplates
	Private actionProperties
	Private actionUserObjects
	
	Public Property Get validator
		validator = actionValidator
	End Property
	
	Public Property Let validator(v)
		actionValidator = v
	End Property
	
	Public Property Get className
		className = pageClassName
	End Property
	
	Public Property Let className(n)
		pageClassName = n
	End Property
	
	Public Property Get method
		method = pageMethod
	End Property
	
	Public Property Let method(m)
		pageMethod = m
	End Property
	
	Public Property Get isClass
		if Not isNull(className) then
			isClass = true
		else
			isClass = false
		End if
	End Property
	
	Public Property Get templates
		Set templates = actionTemplates
	End Property
	
	Public Property Get properties
		Set properties = actionProperties
	End Property
	
	Public Property Get userObjects
		Set userObjects = actionUserObjects
	End Property
	
	Public Property Get userObject(name)
		if objTrophy.config.userObjects.Exists(name) then
			Set userObject = objTrophy.config.userObjects.Item(name)
		else
			userObject = Null
		end if
	End Property
	
	Public Sub Class_Initialize()
		Set startupFunctions = newDictionary()
		Set shutdownFunctions = newDictionary()
		Set actionTemplates = newDictionary()
		Set actionProperties = newDictionary()
		Set actionUserObjects = newDictionary()
		actionValidator = Null
		pageClassName = Null
		isDefault = False
	End Sub
	
	Public Function addProperty(name,var)
		if Not actionProperties.Exists(name) then
			Call actionProperties.Add(name,var)
		end if
	End Function
	
	Public Function getUserObjectProperties(name)
		Set getUserObjectProperties = userObject(name).properties
	End Function
	
	Public Function addUserObject(obj)
		if Not actionUserObjects.Exists(obj.name) then
			Call actionUserObjects.Add(obj.name,obj)
		end if
	End Function
	
	''
	' Runs validator
	Public Function runValidator()
		if Not isNull(actionValidator) then 
			Set validatorObject = new Trophy_FormValidator_RequestHandler
			validatorObject.validationFile = actionValidator
			Call validatorObject.validate()
		end if
	End Function
	
	''
	' Registers a new startup function 
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
	
	Public Function addTemplate(template)
		if Not actionTemplates.Exists(template) then
			Call actionTemplates.Add(template,templates.Count)
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