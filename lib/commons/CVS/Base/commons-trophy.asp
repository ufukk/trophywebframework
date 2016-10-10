<!--#include file="commons-string.asp"-->
<!--#include file="commons-date.asp"-->
<%
'Ufuk KOCOLU 11.04.2004 16:06:00

Public Function getUserObjects(dir)
	Dim obj
	On Error Resume Next
	if Not IsObject(objFs) then
		Set objFs = Server.CreateObject("Scripting.FileSystemObject")
	end if
	Set dict = newDictionary()
	Set d = objFs.GetFolder(dir)
	For Each f In d.Files
		if InStr(f.Name,"cls") = 1 And InStr(f.name,".asp")=Len(f.Name)-3 then
			class_name = Mid(f.name,4,Len(f.Name)-7)
			Err.Clear
			Execute("Set obj = new " & class_name)
			if Err.Number = 0 then
				Call dict.Add(class_name,obj)
			end if
		end if
	Next
	Set getUserObjects = dict
End Function

Public Function getTags(s)
	getTags = "[!" & s & "]"
End Function
	
Public Function stripTags(ByVal s)
	s = Mid(s,3)
	s = Mid(s,1,Len(s)-1)
	stripTags = s
End Function

Public Function loadVariableFiles(ByRef otmp,ByRef arr)
	For i=0 to UBound(arr)
		if arr(i) <> "" then
			path = arr(i)
			Call otmp.loadVariableFile(path)
		end if 
	Next
End Function 

Public Function getControllerScript()
	if CONTROLLER_SCRIPT = "" then
		CONTROLLER_SCRIPT = DEFAULT_CONTROLLER_SCRIPT
	end if
	getControllerScript = CONTROLLER_SCRIPT
End Function

Public Function getControllerParameter()
	if ACTION_PARAMETER = "" then
		ACTION_PARAMETER = DEFAULT_ACTION_PARAMETER
	end if
	getControllerParameter = ACTION_PARAMETER
End Function

Public Function getTrophyLink(action)
	getTrophyLink = getControllerScript() & "?" & getControllerParameter() & "=" & action
End Function

Public Sub invokeMethod(ByRef object,method,args)
	Dim methodCmd,argsStr
	argsStr = ""
	For i=0 to UBound(args)
		argsStr = argsStr & "args(" & i & ")"
		if i <> UBound(args) then
			argsStr = argsStr & ","
		end if
	Next
	methodCmd = "Call object." & method & "(" & argsStr & ")"
	Execute(methodCmd)
End Sub

Public Sub setObjectProperty(ByRef object,ByVal name,ByVal value)
	Dim propCmd
	propCmd = "object." & name & " = value"
	Execute(propCmd)
End Sub

Public Sub setObjectProperties(ByRef object,ByRef properties)
	For Each key in properties
		Call setObjectProperty(object,key,properties.Item(key))
	Next
End Sub

Public Sub setRequestProperty(ByRef object,name,requestName)
	propertyCmd = "object." & name & "=" &  "req(""" & requestName & """)"
	Execute(propertyCmd)	
End Sub

Public Sub setObjectRequestProperties(ByRef object,ByRef properties)
	For Each key in properties
		Call setRequestProperty(object,key,properties.Item(key))
	Next
End Sub

Public Sub setActionProperties(ByRef actionObject)
	For Each name in objTrophy.selectedAction.properties
		var = objTrophy.selectedAction.properties.Item(name)
		if Not isNull(req(name)) then
			Call setRequestProperty(actionObject,name,var)
		end if
	Next
End Sub

Public Sub runActionObject(ByRef object)
	Call setActionProperties(object)
End Sub

Public Function getConfigurationContent()
	Dim cfile,datemodified,applicationTime,refreshConfig
	if applicationCustomConfigKey <> "" And Not isNull(applicationCustomConfigKey) then
		applicationConfigKey = applicationCustomConfigKey
	end if
	if applicationCustomTimeKey <> "" And Not isNull(applicationCustomTimeKey) then
		applicationTimeKey = applicationCustomTimeKey
	end if
	Set cfile = objFs.GetFile(Server.MapPath(DEFAULT_CFG_FILE))
	datemodified = getTimeStamp(cfile.DateLastModified)
	Set cfile = Nothing
	applicationTime = Application(applicationTimeKey)
	refreshConfig = true
	if applicationAlwaysRefreshConfig <> True then
		if applicationTime <> "" then
			if isNumeric(applicationTime) then
				if applicationTime >= datemodified then
					refreshConfig = false
				end if
			end if
		end if
	else
		refreshConfig = true
	end if
	refreshConfig = true
	if refreshConfig then
		Call saveConfigurationFile(DEFAULT_CFG_FILE)
	end if
	getConfigurationContent = Application(applicationConfigKey)
End Function

Public Sub saveConfigurationFile(configFilePath)
	Set textStream = objFs.OpenTextFile(Server.MapPath(DEFAULT_CFG_FILE))
	textContent = textStream.ReadAll
	Application(applicationConfigKey) = textContent
	Application(applicationTimeKey) = getTimeStamp(Now)
End Sub

Public Function getPostData(key)
	getPostData = Session(CONFIG_SESSION_REQUEST_KEY)(key)
End Function
%>