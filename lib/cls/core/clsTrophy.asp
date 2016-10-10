<%


class Trophy

	Public requestAction
	Public config
	Public page
	Public action
	Public selectedAction
	Private rootPath
	Private isControllerPage
	Private loggerObjects
	Private datasourceObjects
	
	Public Property Get loggers(f)
		Set loggers = loggerObjects(f)
	End Property
	
	Public Property Get datasources(d)
		Set datasources = datasourceObjects(d)
	End Property
	
	Public Property Get realPath(p)
		realPath = rootPath & "/" & p
	End Property
	
	Public Property Get path
		path = rootPath
	End Property
	
	Public Property Let path(p)
		rootPath = path 
	End Property
	
	Public Property Let isController(c)
		isControllerPage = CBool(c)
	End Property
	
	Public Property Get isController
		isController = isControllerPage
	End Property
	
	Private Function getDefaultAction()
		For Each actionType in config.actions
			Set actionObject = config.actions.Item(actionType)
			if actionObject.isDefault = true then
				Set getDefaultAction = actionObject
				Exit Function
			end if
		Next
		getDefaultAction = Null
	End Function
	
	Private Function getLoggers()
		For Each loggerName in config.loggers
			Call loggerObjects.Add(loggerName,config.loggers(loggerName))		
		Next
		config.loggers.RemoveAll
	End Function
	
	Private Function getDatasources()
		For Each datasourceName in config.datasources
			Call datasourceObjects.Add(datasourceName,config.datasources(datasourceName))		
		Next
		config.datasources.RemoveAll
	End Function
	
	Private Function registerLogger(name,path,maxsize,format)
		Set lobj = new Trophy_Logger
		lobj.loggerName = name
		lobj.logfilePath = path
		lobj.maxSize = maxsize
		lobj.format = format
		Call loggerObjects.Add(name,lobj)
	End Function
	
	Private Sub setDefaultPath()
		if Not isEmpty(TROPHY_PATH) then
			rootPath = TROPHY_PATH
		else
			rootPath = Server.MapPath(DEFAULT_TROPHY_PATH)
		end if
	End Sub
	
	Private Sub createInstances()
		Call createConfigInstance()
		Call createRequestInstance()
		Call createPageInstance()
		Call createErrorInstance()
		Call createFsInstance()
		Call getLoggers()
		Call getDatasources()
	End Sub
	
	Private Sub createRequestInstance()
		if Not IsObject(req) then
			Set req = new Trophy_RequestHandler
			req.requestFilters = config.request_Filters
		end if
	End Sub
	
	Private Sub createConfigInstance()
		if Not IsObject(config) then
			Set config = new Trophy_Configuration
			config.configFile = path & DEFAULT_CFG_FILE
			if isController = true then
				config.isController = true
			end if
			Call config.load()
		end if
	End Sub
	
	Private Sub createPageInstance()
		if Not IsObject(page) then
			Set page = new Trophy_Page
			Call addPageStartup()
		end if
	End Sub
	
	Public Sub createTemplateInstance()
		if Not IsObject(objTemplate) then
			Set objTemplate = new Trophy_Template
		end if
	End Sub
	
	Private Sub createErrorInstance()
		if Not IsObject(objError) then
			Set objError = new Trophy_Error
		end if
	End Sub
	
	Private Sub createFsInstance()
		if Not IsObject(objFs) then
			Set objFs = newFileSystem()
		end if
	End Sub
	
	Private Function addPageStartup()
		startupFunctions = config.startup_Functions
		For i=0 to UBound(startupFunctions)
			Call page.addStartupFunction(startupFunctions(i))
		Next
	End Function
	
	Public Sub class_initialize()
		isController = false
		Set loggerObjects = newDictionary()
		Set datasourceObjects = newDictionary()
		Call setDefaultPath()
	End Sub
	
	Public Function loadInstance()
		Call createInstances()
	End Function
	
	Public Function stopPage()
		response.end
		terminate()
	End Function
	
	Public Function getDBInstance(datasourceName)
		Set dobj = new Trophy_Database
		dobj.datasource = datasourceObjects.Item(datasourceName)
		Set getDBInstance = dobj
	End Function
	
	Public Sub findAction()
		requestAction = req(getControllerParameter())
		if requestAction = "" then
			Set selectedAction = getDefaultAction()
		else
			if Not config.actions.Exists(requestAction) then
				Set objError = new Trophy_Error
				Call objError.trigger("action-finder","Could not find the action: " & requestAction & " in file: " & DEFAULT_CFG_FILE)
			else
				Set selectedAction = config.actions.Item(requestAction)
			end if
		end if
	End Sub
	
	Public Function terminate()
		For Each item in loggerObjects
			loggerObjects.Item(item).closeFile()
		Next
		For Each item in datasourceObjects
			datasourceObjects.Item(item).close()
		Next
		Set datasourceObjects = Nothing
		Set loggerObjects = Nothing
		Set userObjects = Nothing
	End Function
	
end class
%>