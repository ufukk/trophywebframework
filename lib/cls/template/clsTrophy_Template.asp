<%
Public Const OPEN_TAG = "[!"
Public Const CLOSE_TAG = "]"

class Trophy_Template

	Private resource
	Private filename
	Private filepath
	Private variables
	Private containerObjects
	Private lastContainerID
	Private blockObjects
	Private validBlocks
	Private aliases
	Private objFs
	Private objTxt
	Private objRxp
	Private rawContent
	Private markContent
	Private compiledContent
	Private rxp
	Private limitGlobalAccess
	Private templateKeywords
	Public caseSensitive
	Public debugMode
	
	Public Sub Class_Initialize()
		Set variables = newDictionary()
		Set loop_variables = newDictionary()
		Set blockObjects = newDictionary()
		Set containerObjects = newDictionary()
		Set aliases = newDictionary()
		if Not IsObject(objFs) then
			Set objFs = newFileSystem()
		end if
		Set objRxp = new Regexp
		validBlocks = array("if","loop","container","literal")
		templateKeywords = split(TROPHY_TEMPLATE_KEYWORDS,",")
		blockIndex = 0
		Set rxp = newRegexp()
		debugMode = False
		caseSensitive = False
		limitGlobalAccess = false
	End Sub
	
	Public Sub Class_Terminate()
		Set blockObjects = Nothing
		Set variables = Nothing
		Set Errors = Nothing
		Set objFs = Nothing
		Set objRxp = Nothing
	End Sub
	
	Public Property Let TemplateFile(tname)
		filepath = getRealFilePath(tname)
		Call assign("__FILE__",filepath)
	End Property
	
	Public Property Get globalAccess
		globalAccess = limitGlobalAccess
	End Property
	
	Public Property Let globalAccess(a)
		if a = true or a = false then
			limitGlobalAccess = a
		end if
	End Property
	
	Public Default Property Get vars(k)
		vars = variables.Item(k)
	End Property

	Public Function getResourcePath(resource_)
		Dim resource_path,path
		resource_path = Null
		path = objTrophy.realPath(objTrophy.config.resources(resource_))
		if Len(path) > 0 then
			resource_path = path
		end if
		getResourcePath = resource_path
	End Function
	
	Private Function getRealFilePath(tname)
		typ = "file"
		if InStr(tname,":") > 0 then
			if Not objFs.FileExists(tname) then
				tmpArr = split(tname,":")
				typ = "resource"
			end if
		end if	
		if typ = "resource" then
			getRealFilePath = getResourceTemplate(Trim(tmpArr(0)),Trim(tmpArr(1)))
			Exit Function
		elseif typ = "file" then
			getRealFilePath = getFileTemplate(tname)
			Exit Function
		end if
	End Function
	
	Private Function getResourceTemplate(resource_,file_)
		resource_path = getResourcePath(resource_)
		if Len(resource_path) > 0 then
			getResourceTemplate = resource_path & FILE_SEPARATOR & file_
			Exit Function
		end if
	End Function
	
	Private Function getFileTemplate(file_)
		getFileTemplate = file_
	End Function
	
	'static
	Private Function getAttribute(str,attr,isDynamic)
		Set rxp = newRegexp
		rxp.Pattern = "" & attr & "\=([" & join(ARG_SEPARATORS,"") & "])" & "([^]]*)\1"
		Set results = rxp.Execute(str)
		if results.Count = 0 then
			getAttribute = Null
			Exit Function
		else
			Set result = results(0)
			separator = result.SubMatches(0)
			returnValue = result.SubMatches(1)
			if InStr(returnValue,separator) > 0 then
				values = split(returnValue,separator)
				returnValue  = values(0)
			end if
			'replace dynamic variables
			if isDynamic then
				Set innerRxp = newRegexp()
				innerRxp.Pattern = "\{([^}]*)\}"
				Set attributeVars = innerRxp.Execute(returnValue)
				For Each attributeVar in attributeVars
					variable = attributeVar.SubMatches(0)
					attributeValue = setVariable(getTags(variable))
					attributeVarName = Trim(Mid(attributeVar,2,Len(attributeVar)-2))
					if Not IsInteger(attributeValue) then
						attributeValue = fixSlashesAndQuote(attributeValue)
					end if
					if Len(attributeValue)=0 then
						attributeValue = """"""
					end if
					returnValue = replace(returnValue,attributeVar,attributeValue)
				Next
			end if
			getAttribute = returnValue
			Set innerRxp = Nothing
			Set rxp = Nothing
			Exit Function
		end if
	End Function
	
	Private Function getAttributeMap(str)
		Dim rObject,resultDict,value
		Set resultDict = newDictionary()
		Set rObject = newRegexp()
		rObject.Pattern = "([a-zA-Z0-9]*)\=(""|`)"
		Set attrList = rObject.Execute(str)
		For Each attr in attrList
			value = getAttribute(str,attr.SubMatches(0),true)
			Call resultDict.Add(attr.SubMatches(0),value)
		Next
		Set getAttributeMap = resultDict
	End Function
	
	Public Function getVariable(key)
		getVariable = variables.Item(key)
	End Function
	
	Public Function getVariablesByParent(parent)
		Set resultDict = newDictionary()
		For Each variable in variables
			if InStr(variable,parent & ".") > 0 then
				varName = Mid(variable,Len(parent)+2)
				if InStr(varName,".") = 0 then
					Call resultDict.Add(varName,variables.Item(variable))
				end if
			end if
		Next
		Set getVariablesByParent = resultDict
	End Function
	
	Public Function isAssigned(key)
		isAssigned = variables.Exists(key)
	End Function
	
	Public Function assign(key,val)
		if validateVariableName(key) then
			Call remove(key)
			Call variables.Add(key,val)
		else
			Call objError.trigger("template","use of reserved keyword in " & key)
		end if
	End Function
	
	Private Function validateVariableName(variableName)
		Dim r
		r = true
		For i=0 to UBound(templateKeywords)
			if InStr(variableName,templateKeywords(i)) = 1 then
				r = false
			end if
		Next
		validateVariableName = r
	End Function
	
	Private Function assignBlockObject(id,btype,header,body,footer)
		if blockObjects.Exists(id) then
			Call blockObjects.remove(id)
		end if
		Set bO = new Trophy_TemplateBlock
		Call bO.loadBlock(id,btype,header,body,footer)
		Call blockObjects.Add(id,bO)
	End Function
	
	Private Function addContainerObject(id,ByRef obj)
		if Not containerObjects.Exists(id) then
			Call containerObjects.Add(id,obj)
		end if	
	End Function
	
	Public Function assignByRef(key,ByRef val)
		Call remove(key)
		Call variables.Add(key,val)
	End Function
	
	Public Function assignTemplate(key,tmpfile)
		Set objT = new Trophy_Template
		Call assign(key,objT.fetchFile(tmpfile))
	End Function

	Public Function removeAll
		Call variables.removeAll()
	End Function
	
	Public Function remove(key)
		if variables.Exists(key) then
			Call variables.Remove(key)
		end if
	End Function
	
	Public Function removeAlias(key)
		if aliases.Exists(key) then
			Call aliases.Remove(key)
		end if
	End Function
	
	Public Function assignAlias(key,val)
		Call removeAlias(val)
		Call aliases.Add(val,key)
	End Function
	
	Private Function loadVariables(target,path)
		Dim varFile,fileContent,vars,varObj,var_name,var_value
		realFilePath = getRealFilePath(path)
		if Not objFs.FileExists(realFilePath) then
			Call objError.trigger("template","variable file not found: " & realFilePath)
		end if
		Set varFile = objFs.OpenTextFile(realFilePath)
		fileContent = varFile.ReadAll
		Set varFile = Nothing
		rxp.Pattern = "([a-zA-Z0-9_.]*)( ?)\=( ?)" & """" & "([^" & """" & "]*)" & """"
		Set vars = rxp.Execute(fileContent)
		if target="resource" then
			Set rDict = newDictionary()
		end if
		For Each varObj in vars
			if target="this" then
				Call assign(varObj.SubMatches(0),varObj.SubMatches(3))
			else
				Call rDict.Add(varObj.SubMatches(0),varObj.SubMatches(3))
			end if
		Next
		if target="resource" then
			Set loadVariables = rDict
			Exit Function
		end if
	End Function
	
	Public Function loadVariableFile(path)
		Call loadVariables("this",path)
	End Function
	
	Public Function getResourceFile(path)
		Set getResourceFile = loadVariables("resource",path)
	End Function
	
	Public Function readAll(filepath,par)
		Dim result,content
		if Not objFs.FileExists(filepath) then
			Call objError.trigger("template","template file not found: " & filepath)
		end if
		Set objTxt = objFs.OpenTextFile(filepath)
		content = objTxt.ReadAll
		Set objTxt = Nothing
		objRxp.Global = True
		objRxp.MultiLine = True
		objRxp.IgnoreCase = True 
		objRxp.Pattern = INCLUDE_PATTERN
		Set results = objRxp.Execute(content)
		For Each result In results
			str = result.Value
			path = getRealFilePath(getAttribute(str,"file",false))
			content = replace(content,str,readAll(path,1))
			a = a+1
		Next
		if par = 0 then
			compiledContent = content
		end if
		readAll = content
	End Function
	
	Public Function markBlocks(ByRef content) 
		Dim openIndex,closeIndex,orgContent
		markPattern = ".*(\[\!(" & join(validBlocks,"|") & ")([^]]*)\])((.|(\n))*?)(\[\!\/(\2)\])"
		Set rexp = newRegexp()
		rexp.multiline = false
		rexp.Pattern = markPattern
		Set blockHeaders = rexp.Execute(content)
		if blockHeaders.Count > 0 then
			Set blockHeader = blockHeaders(0)
			blockType = blockHeader.SubMatches(1)
			nextBlockHeader = InStrRev(blockHeader.Value,OPEN_TAG & blockType)
			innerBlock = Mid(blockHeader.Value,nextBlockHeader)
			variableName = saveBlockObject(blockType,innerBlock)
			content = replace(content,innerBlock,OPEN_TAG & variableName & CLOSE_TAG)
			Call markBlocks(content)
		end if
		Set rexp = Nothing
	End Function
	
	Private Function saveBlockObject(btype,rawcontent)
			Set rxp = newRegexp()
			markPattern = ".*(\[\!" & btype & "([^]]*)\])((.|(\n))*?)" & "(\[\!\/" & btype & "\])"
			rxp.Pattern = markPattern
			Set blockParts = rxp.Execute(rawcontent)
			if blockParts.Count > 0 then
				Set matchObj = blockParts(0) 
				blockBody = matchObj.SubMatches(2)
				Call markBlocks(blockBody)
				t = blockObjects.Count+1
				variableName = TROPHY_TEMPLATE_BLOCK_PREFIX & t
				Call assignBlockObject(variableName,btype,matchObj.SubMatches(0),blockBody,matchObj.SubMatches(5))
			end if
			saveBlockObject = variableName
	End Function
	
	Public Function parse(ByRef content)
		Dim val2
		if Len(content) > 0 then
			objRxp.Global = True
			if CaseSensitive then
				objRxp.IgnoreCase = False
			else
				objRxp.IgnoreCase = True
			end if
			objRxp.Pattern = VAR_PATTERN
			Set results = objRxp.Execute(content)
			For Each trpDirection in results
				Set rxp = new Regexp
				rxp.Global = True
				rxp.IgnoreCase = True
				val = setVariable(trpDirection.Value)
				content = replace(content,trpDirection.Value,val)
			Next
		end if
	End Function
	
	Private Function getFunction(trpDirection) 
		Set frs = rxp.Execute(trpDirection)
		rxp.MultiLine = True
		Dim tmpArr,argstr
		'arguments
		rxp.Pattern = "\[\![a-zA-Z0-9_]*\((.*)\)\]"
		argstr = rxp.Replace(trpDirection,"$1")
		rxp.Pattern = "[a-zA-Z0-9_]*\(.*\)"
		Set fRs = rxp.Execute(argstr)
		For Each fR in fRs
			if fR.FirstIndex = 0 then
				firstIndex = 1
			else
				firstIndex = fR.FirstIndex -1
			end if
			if Mid(argstr,firstIndex,1) <> " " then
				argstr = replace(argstr,fR,"""" & fixSlashes(getFunction(getTags(fR))) & """")
			end if
		Next 
		rxp.Pattern = "\[\!([a-zA-Z0-9_]*).*\]"
		'function name
		functionName = rxp.Replace(trpDirection,"$1")
		if limitGlobalAccess = true then
			functionName = "templateFunction_" & functionName
		end if
		'only one argument
		Set parameters = getParameterList(argstr)
		For Each parameter in parameters
			parameterValue = parameters.Item(parameter)
			if InStr(parameterValue,"""") <> 1 then
				if Not IsNumeric(parameterValue) then
					'dynamic variable
					parameterValue = fixSlashesAndQuote(setVariable(getTags(parameterValue)))
				else
					parameterValue = CInt(parameterValue)
				end if
			end if
			argList = argList & parameterValue
			argList = argList & ","
			j = j+1
		Next
		argList = str_unpush(argList)
		cmd = "val = " & functionName & "(" & argList & ")"
		Execute(cmd)
		getFunction = val
	End Function
	
	Public Function setVariable(trpDirection)
		Dim tmps,fresults
		trpStr = trpDirection
		strippedValue = stripTags(trpStr)
		if aliases.Exists(strippedValue) then
			trpDirection = getTags(aliases.Item(stripTags(trpStr)))
		end if
		if InStr(strippedValue,"{") = 1 And InStr(strippedValue,"}") = Len(strippedValue) then
			trpDirection = getTags(setVariable(getTags(str_unshift(str_unpush(strippedValue)))))
		end if
		Set rxp = new Regexp
		rxp.Global = True
		rxp.IgnoreCase = True
		rxp.Pattern = FUNCTION_PATTERN
		Set fresults = rxp.Execute(trpStr)
		isFunction = False
		x = trpDirection
		var_name = stripTags(x)
		if fresults.Count > 0 then
			if fresults(0).FirstIndex = 0 then
				isFunction = True
			end if
		end if
		if isFunction=True then
			wTags = Trim(stripTags(trpDirection))
			functionName = Mid(wTags,1,InStr(wTags,"(")-1)
			if variables.Exists(functionName) then
				'todo: array support
				val = getFunction(trpDirection)
			else
				val = getFunction(trpDirection)
			end if
		elseif blockObjects.Exists(var_name) then
			val = getBlockVariable(var_name)
		else	
			if InStr(trpDirection," ") > 0 then
				prefix = Mid(trpDirection,2,InStr(trpDirection," "))
				prefix = Mid(prefix,2,Len(prefix)-2)
				prefix = Trim(prefix)
				if prefix <> "loop" then
					key = Mid(trpDirection,Len(prefix)+2)
					key = Mid(key,3,Len(key)-2)
					key = Mid(key,1,Len(key)-1)
					if prefix <> "global" then
						val = getExternalVariable(prefix,key)
					else
						val = eval(key)
					end if
				end if
				else
					key = Mid(trpDirection,3)
					key = Mid(key,1,Len(key)-1)
					val = variables.Item(key)	
				end if
		end if
		setVariable = val
	End Function
	
	Private Function getBlockVariable(varname)
		Set blockObject = blockObjects.Item(varname)
		blockType = blockObject.bType
		Execute("getBlockVariable = getBlockVariable_" & blockType & "(blockObject)")
	End Function
	
	Private Function getBlockVariable_Literal(blockObject)
		getBlockVariable_Literal = blockObject.body
	End Function
	
	Private Function getBlockVariable_Container(blockObject)
		Dim containerObj
		containerType = getAttribute(blockObject.Header,"type",false)
		containerID = getAttribute(blockObject.Header,"id",false)
		containerGeneratorCmd = "Set containerObj = new Trophy_TemplateContainer_" & containerType
		Execute(containerGeneratorCmd)
		containerObj.ID = containerID
		containerObj.templateObj = Me
		'call container starter method
		containerObj.starter
		lastContainerID = containerID
		Call containerObj.setPropertyMap(getAttributeMap(blockObject.Header))
		Call containerObjects.Add(containerID,containerObj)
		body = containerObj.header & blockObject.Body & containerObj.footer
		Call parse(body)
		'call container finisher method
		containerObj.finisher
		getBlockVariable_Container = body
	End Function
	
	Private Function getBlockVariable_If(blockObject)
		header = blockObject.Header
		body = blockObject.Body
		condition = getAttribute(header,"condition",true)
		if Not IsNull(condition) then
			elseIndex = InStr(blockObject.body,"[!else]")
			if elseIndex <> 0 then
				trueBody = Mid(blockObject.body,1,elseIndex-1)
				falseBody = Mid(blockObject.body,elseIndex)
			else
				trueBody = blockObject.body
				falseBody = ""
			end if
			'condition result, true or false
			conditionResult = Eval(condition)
			if conditionResult then
				body = trueBody
			else
				body = falseBody
			end if
			'parse block body and return result
			Call parse(body)
			getBlockVariable_If = body
			Exit Function
		else
			Call objError.trigger("template","no condition attribute in if string: " & htmlEncode(header))
		end if
	End Function
	
	Private Function getBlockVariable_Loop(blockObject)
		Dim name,var,r,o
		header = blockObject.Header
		str = blockObject.Content
		name = getAttribute(blockObject.Header,"name",true)
		var = getAttribute(blockObject.Header,"var",true)
		vtype = getAttribute(blockObject.Header,"type",false)
		if IsNull(name) then
			Call objError.trigger(filepath,"Missing attribute: name, " & header)
		end if
		if IsObject(variables.Item(var)) then
			if IsNull(vtype) then
				vtype = TypeName(variables.Item(var))
			end if
			Set loop_var = variables.Item(var)
		else
			loop_var = variables.Item(var)
		end if
		Set r = new Regexp
		r.Global = True
		r.Multiline = true
		r.IgnoreCase = True
		r.Pattern = "\[\!loop ([^]]*)\]"
		Set mtc = r.Execute(str)
		Set s = mtc(0)
		body = Mid(str,Len(s)+1)
		body = Mid(body,1,Len(body)-8)
		tmp = body
		loopContent = ""
		if IsArray(loop_var) then
			loopContent = getLoopOutput_array(body,name,loop_var)
		elseif IsObject(loop_var) And vtype = "Dictionary" then
			loopContent = getLoopOutput_dictionary(body,name,loop_var)
		elseif IsObject(loop_var) And vtype = "Recordset" then
			limit = getAttribute(str,"limit",true)
			page = getAttribute(str,"page",true)
			autoclose = getAttribute(str,"autoclose",true)
			if IsNull(limit) then
				limit = -1
			else
				limit = CInt(limit)
			end if	
			
			if IsNull(page) then
				page = 1
			else
				page = CInt(page)
			end if	
			if IsNull(autoclose) then
				autoclose = False
			else
				if autoclose="True" Or autoclose="1" then
					autoclose = True
				else
					autoclose = False
				end if
			end if
			loopContent = getLoopOutput_recordset(body,name,loop_var,limit,page,autoclose)
		elseif vtype = "for" then
			startNumber = getAttribute(str,"start",true)
			endNumber = getAttribute(str,"end",true)
			stepNumber = getAttribute(str,"step",true)
			if IsNull(startNumber) then
				startNumber = 0
			else
				startNumber = CInt(startNumber)
			end if
			if IsNull(endNumber) then
				endNumber = 10
			else
				endNumber = CInt(endNumber)
			end if
			if IsNull(stepNumber) then
				stepNumber = 1
			else
				stepNumber = CInt(stepNumber)
			end if
			loopContent = getLoopOutput_for(body,name,startNumber,endNumber,stepNumber)
		end if
		getBlockVariable_Loop = loopContent
	End Function
	
	Private Function getLoopOutput_array(ByRef tmp,name,array_var)
		Dim o
		o = ""
		For i = 0 To UBound(array_var)
			Call assign("#" & name,i)
			Call assign("$" & name,array_var(i))  
			tmpS = tmp
			Call parse(tmpS)
			o = o & tmpS
		Next
		getLoopOutput_array = o
	End Function
	
	Private Function getLoopOutput_dictionary(ByRef tmp,name,dict_var)
		Dim o
		o = ""
		For Each k in dict_var
			Call assign("#" & name,k)
			Call assign("$" & name,dict_var.Item(k))
			tmpS = tmp
			Call parse(tmpS)
			o = o & tmpS
		Next
		getLoopOutput_dictionary = o
	End Function
	
	Private Function getLoopOutput_recordset(ByRef tmp,name,ByRef rs,limit,page,autoclose)
		Dim counter,rsindex,o
		o = ""
		if limit=-1 then
			counter = 1
		else
			counter = limit
		end if
		rsindex = 0
		z = 0
		Do While Not rs.EOF And rsindex < counter
			Call assign("#" & name,z)
			tmpS = tmp
			Call parse(tmpS)
			if limit > -1 then
				rsindex = rsindex+1
			end if
			z = z+1
			o = o & tmpS
			rs.movenext
		Loop
		if autoclose = True then
			rs.Close
		end if
		getLoopOutput_recordset = o
	End Function
	
	Private Function getLoopOutput_for(tmp,name,startNumber,endNumber,stepNumber)
		Dim p,o
		p = 0
		o = ""
		For p=startNumber to endNumber
			Call assign("#" & name,p)
			tmpS = tmp
			Call parse(tmpS)
			o = o & tmpS
		Next
		getLoopOutput_for = o
	End Function
	
	Private Function getExternalVariable(typ,key) 
		Dim val,nVal
		Select Case typ
			Case "assign"
				tkey = getAttribute(key,"key",true)
				tval = getAttribute(key,"value",true)
				tvalstr = getTags(tval)
				nVal = setVariable(tvalstr)
				if nVal = "" then
					nVal = tval
				end if
				Call assign(tkey,nval)
			Case "this"
				val = eval(key)
			Case "Call"
				val = eval(key)	
			Case "trophy"
				val = eval("objTrophy(" & key & ")")
			Case "request"
				val = getRequestVariable(key)
			Case "session"
				val = getSessionVariable(key)
			Case "cookie"
				val = getCookieVariable(key)
			Case "element"
				val = getElementVariable(key)
			Case Else
				'collection or object
				if variables.Exists(typ) then
					if isObject(variables.Item(typ)) then
						Set obj = variables.Item(typ)	
						if InStr(key,".") = 1 then
							eStr = Trim(typ) & Trim(key)
							val = eval("obj" & key)
							if IsNull(val) then
								val = ""
							end if
						else
							'collection ex:rs
							val = obj(key)
						end if
					end if
				end if
		End Select
		getExternalVariable = val
	End Function
	
	Private Function getRequestVariable(key)
		getRequestVariable = req(key)
	End Function 
	
	Public Function getSessionVariable(key)
		Dim val
		if key = "id" then
			val = Session.SessionID
		else
			val = Session(key)
		end if
		getSessionVariable = val
	End Function
	
	Public Function getCookieVariable(key)
		getCookieVariable = Request.Cookies(key).Item
	End Function
	
	Public Function getServerVariable(key)
		getServerVariable = Request.ServerVariables(key)
	End Function
	
	Private Function getElementVariable(key) 
		if Not containerObjects.Exists(lastContainerID) then
			Call objError.trigger("template","no mathing container objects found")
		end if
		elementType = getAttribute(key,"type",false)
		reference = getAttribute(key,"ref",true)
		if reference <> "" then
			getElementVariable = containerObjects.Item(lastContainerID).addElementByRef(elementType,variables.Item(reference))
			Exit Function
		else
			getElementVariable = containerObjects.Item(lastContainerID).addElement(elementType,getAttributeMap(key))
			Exit Function
		end if
	End Function
	
	Private Function Math(v1)
		argstr = v1
		Set rxp = new Regexp
		rxp.Global = True
		rxp.IgnoreCase = True
		rxp.Multiline = True
		rxp.Pattern = "[a-zA-Z_0-9 ." & """" & "]*[\+\-\*\/\(\)]{1}"
		Set rp = new Regexp
		rp.Global = True
		rp.IgnoreCase = True
		rp.Multiline = True
		Set results = rxp.Execute(v1)
		For Each result in results
			rp.Pattern = "[a-zA-Z_ ." & """" & "]{2,}"
			If rp.Test(result) then
				rp.pattern = "[\+\*\-\/]"
				var = Trim(rp.Replace(result,""))
				argstr = replace(argstr,var,setVariable(getTags(var)))
			end if
		Next
		cmd = "val=" & argstr
		Execute(cmd)
		Math = val
	End Function
	
	Public Function debugBlocks()
		For Each key in blockObjects
			echo key & "=" & blockObjects.Item(key) & "<hr>"
		Next
	End Function
	
	Public Function Read()
		Call markBlocks(compiledContent)
		Call parse(compiledContent)
	End Function
	
	Public Function fetchString(str)
		compiledContent = str
		Call read()
		fetchString = compiledContent
	End Function
	
	Public Function displayString(str)
		echo fetchString(str)
	End Function
	
	Public Function getOutput()
		Call readAll(filepath,0)
		Call Read()
		getOutput = compiledContent
	End Function
	
	Public Function display()
		Call getOutput()
		echo compiledContent
	End Function
	
	Public Function fetchFile(f)
		TemplateFile = f
		fetchFile = getOutput()
	End Function
	
	Public Function displayTFile(resource,tmp)
		Call displayFile(resource & ":" & tmp)
	End Function 
	
	Public Function fetchTFile(resource,tmp)
		fetchTFile = fetchFile(resource & ":" & tmp)
	End Function
	
	Public Function displayFile(f) 
		TemplateFile = f
		Call getOutput()
		echo compiledContent
	End Function
	
	Public Function toString()
		Dim o,valueStr,key
		o = "<table border='0' style='font-family:Tahoma;font-size:10px'><tr><td><h3>VARIABLES</h3></td></tr></table>"
		For Each key in variables
			if Not IsObject(variables.Item(key)) And Not IsNull(variables.Item(key)) And Not IsArray(variables.Item(key)) then
				valueStr = htmlEncode(variables.Item(key))
			elseif IsArray(variables.Item(key)) then
				valueArray = variables.Item(key)
				valueStr = "Array, Size <b>" & UBound(valueArray) & "</b>" 
				for j=0 to UBound(variables.Item(key))
					valueElement = valueArray(j) 
					if Not IsObject(valueElement) then
						aV = valueElement
					else
						aV = "Object: " & TypeName(valueElement)
					end if
					o = o & "<table border='0'><tr><td>" & aV & "</td></tr></table>"
				Next
			else
				valueStr = TypeName(variables.Item(key))
			end if
			o = o & "<table border='0'><tr><td>Name:</td><td>" & key & "</td></tr>"
			o = o & "<tr><td>Value</td><td>" & valueStr & "</td></tr>"
			o = o & "</table>"
		Next
		toString = o
	End Function
	
end Class
%>