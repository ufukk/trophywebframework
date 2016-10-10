<%


class Trophy_Configuration

	Private startupFunctions()
	Private shutdownFunctions()
	Private requestFilters
	Private autoLoadFiles
	Private startupFunctions_Started
	Private shutdownFunctions_Started
	Private requestFilters_Started
	Private autoload_Started
	Private templateResources
	Private defaultIndex
	Private element
	Private singlePageMode
	Private parsed
	Private objFs
	Private configFilePath
	Private isControllerPage
	Public loggers
	Public datasources
	Public userObjects
	Public parameters
	Public actions
	
	Public Default Property Get parameter(k)
		parameter = parameters(k)
	End Property
	
	Public Property Get request_Filters()
		request_Filters = requestFilters
	End Property
	
	Public Property Get startup_Functions()
		startup_Functions = startupFunctions
	End Property
	
	Public Property Get shutdown_Functions()
		shutdown_Functions = shutdownFunctions
	End Property
	
	Public Property Get defaultActionIndex
		defaultActionIndex = defaultIndex
	End Property
	
	Public Property Get isController
		 isController = isControllerPage
	End Property
	
	Public Property Let isController(c)
		isControllerPage = CBool(c)
	End Property
	
	Public Property Get resources(f) 
		resources = templateResources.Item(f)
	End Property
	
	Public Property Get filestoload
		filestoload = autoloadfiles
	End Property
	
	Public Property Get configFile
		configFile = configFilePath
	End Property
	
	Public Property Let configFile(c)
		 configFilePath = c
	End Property
	
	Public Sub Class_Initialize()
		startupFunctions_Started = False
		requestFilters_Started = False
		singlePageMode = False
		autoLoad_Started = False
		isControllerPage = False
		if Not IsObject(objFs) then
			Set objFs = newDictionary()
		end if
		Set actions = newDictionary()
		Set parameters = newDictionary()
		Set templateResources = newDictionary()
		Set loggers = newDictionary()
		Set datasources = newDictionary()
		Set userObjects = newDictionary()
		requestFilters = array("")
		autoLoadFiles = array("")
		defaultIndex = 0
		configurationInited = True
		Set objFs = Nothing
	End Sub
	
	Private Function addAction(atype,actionObject)
		if Not actions.Exists(atype) then
			Call actions.Add(atype,actionObject)
		else
			Call raiseError(vbObjectError,"duplicate action-type: " & atype)
		end if
	End Function
	
	Private Function getElementFunction(tag)
		func = replace(tag,"-","")
		prefix = "handleSingleElement_"
		cmd = prefix & func
		getElementFunction = cmd
	End Function
	
	Private Function parseDocument()
		Set xmlObj = Server.CreateObject("Microsoft.XMLDOM")
		xmlSrc = getConfigurationContent()
		Call xmlObj.loadXML(xmlSrc)
		if xmlObj.parseerror.ErrorCode = 0 then
			Call handleElements_Page(xmlObj.getElementsByTagName("page"))
			Call handleElements_Actions(xmlObj.getElementsByTagName("actions").item(0))
			if Not isControllerPage then
				Call handleElements_GlobalParameters(xmlObj.getElementsByTagName("global-parameters"))
				Call handleElements_Templates(xmlObj.getElementsByTagName("templates").item(0))
				Call handleElements_Loggers(xmlObj.getElementsByTagName("loggers").item(0).getElementsByTagName("logger"))
				Call handleElements_DataSources(xmlObj.getElementsByTagName("datasources").item(0).getElementsByTagName("datasource"))
				Call handleElements_UserObjects(xmlObj.getElementsByTagName("userobjects").item(0).getElementsByTagName("userobject"))
			end if
		else
			raiseError vbObjectError,"Could not parse config file: " & xmlObj.parseError.reason & " file: <b>" & DEFAULT_CFG_FILE & "</b> Line Number: <b>" & xmlObj.parseError.Line & "</b>"
			response.end
		end if
		parsed = True
	End Function
	
	Private Function handleElements_Page(nodes)
		Dim elements,element,c,tagName
		Set elements = nodes(0).childNodes
		c = 0
		For c=0 to elements.length-1
			Set element = elements.item(c)
			elementType = element.getAttribute("type")
			cmd = "Call " & getElementFunction(elementType) & "(element)"
			Call Execute(cmd)
		Next
	End Function
	
	Private Function handleElements_Actions(parentNode)
		Dim elements,element,c,tagName
		Set elements = parentNode.getElementsByTagName("action")
		c = 0
		For c=0 to elements.length-1
			Set element = elements.item(c)
			tagName = element.tagName
			cmd = "Call " & getElementFunction(element.tagName) & "(element)"
			Call Execute(cmd)
		Next
	End Function
	
	Private Function handleElements_GlobalParameters(nodes)
		Dim elements,a
		Set elements = nodes(0).childNodes
		a = 0
		For a=0 to elements.length-1
			Set element = elements.item(a)
			parameterName = element.getAttribute("name")
			parameterValue = element.getAttribute("value")
			if Not parameters.Exists(tagName) then
				Call parameters.Add(parameterName,parameterValue)
			end if
		Next
	End Function
	
	Private Function handleElements_Templates(nodes)
		Dim elements,element,c,tagName
		Set elements = nodes.childNodes
		c = 0
		For c=0 to elements.length-1
			Set element = elements.item(c)
			tagName = element.tagName
			cmd = "Call " & getElementFunction(element.tagName) & "(element)"
			Call Execute(cmd)
		Next
	End Function
	
	Private Function handleElements_Loggers(nodes)
		For i=0 to nodes.length-1
			Set lobj = new Trophy_Logger
			lobj.loggerName = nodes.item(i).getAttribute("name")
			lobj.logfilePath = nodes.item(i).getAttribute("path")
			lobj.maxSize = nodes.item(i).getAttribute("maxsize")
			lobj.format = nodes.item(i).getAttribute("format")
			Call loggers.Add(nodes.item(i).getAttribute("name"),lobj)
		Next
	End Function
	
	Private Function handleElements_DataSources(nodes)
		For i=0 to nodes.length-1
			Set dobj = new Trophy_Datasource
			dobj.dataSourceName = nodes.item(i).getAttribute("name")
			dobj.dataSourceUrl = nodes.item(i).getAttribute("url")
			Call datasources.Add(nodes.item(i).getAttribute("name"),dobj)
		Next
	End Function
	
	Private Function handleElements_UserObjects(nodes)
		For i=0 to nodes.length-1
			Set uobj = new Trophy_UserObject
			uobj.name = nodes(i).getAttribute("name")
			uobj.className = nodes(i).getAttribute("className")
			uobj.dataSource = nodes(i).getAttribute("datasource")
			uobj.dbObjectName = nodes(i).getAttribute("dbobjectName")
			Call userObjects.Add(uobj.name,uobj)
		Next
	End Function
	
	Private Function handleSingleElement_Action(element)
		atype = element.getAttribute("type")
		page =  element.getAttribute("page") 
		className = element.getAttribute("className")
		method = element.getAttribute("method")
		if isNull(method) then
			method = "processRequest"
		end if
		Set aobj = new Trophy_Action
		aobj.actionType = atype
		aobj.page = page
		if className <> "" then
			aobj.className = className
			aobj.method = method
		end if
		if element.getAttribute("default") = "true" then
			aobj.isDefault = True
		end if
		if Not isNull(element.getAttribute("validator")) then
			aobj.validator = element.getAttribute("validator")
		end if
		Set actionStartupFunctions = element.getElementsByTagName("startup-function")
		Set actionShutdownFunctions = element.getElementsByTagName("startup-function")
		For i=0 to actionStartupFunctions.length-1
			Call aobj.addStartupFunction(actionStartupFunctions.item(i).Text)
		Next
		For i=0 to actionShutdownFunctions.length-1
			Call aobj.addshutdownFunction(actionShutdownFunctions.item(i).Text)
		Next
		if Not isController then
			if element.getElementsByTagName("action-templates").length > 0 then
				Call handleElements_ActionTemplates(aobj,element.getElementsByTagName("action-templates").item(0).getElementsByTagName("action-template"))
			end if
			if element.getElementsByTagName("action-properties").length > 0  then
				Call handleElements_ActionProperties(aobj,element.getElementsByTagName("action-properties").item(0).getElementsByTagName("action-property"))
			end if
			if element.getElementsByTagName("action-userobjects").length > 0 then
				Call handleElements_actionUserObjects(aobj,element.getElementsByTagName("action-userobjects").item(0).getElementsByTagName("action-userobject"))
			end if
		end if
		Call addAction(atype,aobj)
	End Function
	
	Private Function handleElements_ActionTemplates(action,nodes)
		For i=0 to nodes.length-1
			Call action.addTemplate(nodes(i).getAttribute("path"))
		Next
	End Function
	
	Private Function handleElements_ActionProperties(action,nodes)
		For i=0 to nodes.length-1
			if isNull(nodes(i).getAttribute("parameter")) then
				var = nodes(i).getAttribute("name")
			else
				var = nodes(i).getAttribute("parameter")
			end if
			Call action.addProperty(nodes(i).getAttribute("name"),var)
		Next
	End Function
	
	Private Function handleElements_actionUserObjects(action,nodes) 
		Dim uobj
		For i=0 to nodes.length-1
			Set element = nodes(i)
			Set uobj = new Trophy_Action_UserObject
			uobj.name = nodes(i).getAttribute("name")
			uobj.instanceName = nodes(i).getAttribute("instanceName")
			passAllProperties = nodes(i).getAttribute("passAllProperties")
			if passAllProperties <> "" then
				uobj.passAllProperties = nodes(i).getAttribute("passAllProperties")
			end if
			if element.getElementsByTagName("action-userobject-properties").length > 0 then
				Set propNodes = element.getElementsByTagName("action-userobject-properties").item(0).getElementsByTagName("action-userobject-property")
				For z=0 to propNodes.length-1
					name = propNodes(z).getAttribute("name")
					if isNull(propNodes(z).getAttribute("parameter")) then
						var = name
					else
						var = propNodes(z).getAttribute("parameter")
					end if
					Call uobj.setProperty(name,var)
				Next
			end if
			Call action.addUserObject(uobj)
		Next
	End Function
 
	Private Function handleSingleElement_startupFunction(element)
		value = element.getAttribute("name")
		if InStr(value,"(")=1 then
			funcName = value
		else
			funcName = "pageStartup_" & value
		end if
		dim arrayIndex
		if Not startupFunctions_Started then
			arrayIndex = 0
			startupFunctions_Started = True
		else
			arrayIndex = UBound(startupFunctions)+1
		end if
		Redim Preserve startupFunctions(arrayIndex)
		startupFunctions(arrayIndex) = funcName
	End Function
	
	Private Function handleSingleElement_shutDownFunction(element)
		value = element.getAttribute("name")
		if InStr(value,"(")=1 then
			funcName = value
		else
			funcName = "pageShutdown_" & value
		end if
		arrayIndex = 0
		if Not shutdownFunctions_Started then
			arrayIndex = 0
			shutdownFunctions_Started = True
		else
			arrayIndex = UBound(shutdownFunctions)+1
		end if
		Redim Preserve shutdownFunctions(arrayIndex)
		shutdownFunctions(arrayIndex) = funcName
	End Function
	
	Private Function handleSingleElement_requestFilter(element)
		value = element.getAttribute("name")
		if InStr(value,"(")=1 then
			funcName = value
		else
			funcName = "requestFilter_" & value
		end if
		if Not requestFilters_Started then
			arrayIndex = 0
			requestFilters_Started = True
		else
			arrayIndex = UBound(requestFilters)+1
		end if
		Redim Preserve requestFilters(arrayIndex)
		requestFilters(arrayIndex) = funcName
	End Function
	
	Private Function handleSingleElement_autoLoadFile(element)
		value = element.getAttribute("path")
		if Not autoLoadFiles_Started then
			arrayIndex = 0
			autoLoadFiles_Started = True
		else
			arrayIndex = UBound(autoLoadFiles)+1
		end if
		Redim Preserve autoLoadFiles(arrayIndex)
		autoLoadFiles(arrayIndex) = value
	End Function
	
	Private Function handleSingleElement_resource(element)
		if Not templateResources.Exists(element.getAttribute("name")) then
			Dim path
			if element.getAttribute("relative") = "false" then
				path = element.getAttribute("path")
			else
				path = element.getAttribute("path")
			end if
			Call templateResources.Add(element.getAttribute("name"),path)
		end if
	End Function
	
	Public Function load()
		Call parseDocument()
	End Function
	
end class

%>