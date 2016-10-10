<%


class Trophy_FormValidator_XMLReader

	Private validationXML
	
	Public Property Get file
		file = validationXML
	End Property
	
	Public Property Let file(f)
		validationXML = f
	End Property
	
	Private Function addField(ByRef formValidatorObject,ByRef node)
		Dim client
		if isNull(node.getAttribute("client")) then
			client = true
		else
			client = CBool(node.getAttribute("client"))
		end if
		if isNull(node.getAttribute("server")) then
			serverVal = true
		else
			serverVal = CBool(node.getAttribute("server"))
		end if
		Call formValidatorObject.addField(node.getAttribute("name"),node.getAttribute("rule"),node.getAttribute("message"),client,serverVal)
	End Function
	
	Public Sub class_initialize()
	End Sub
	
	Public Sub class_terminate()
	End Sub
	
	Private Function loadXML()
		Dim parserObj,formObject,filePath
		Set parserObj = newXMLParser()
		filePath = objTrophy.realPath("WEB-INF/validator/" & file)
		Call parserObj.load(filePath)
		if parserObj.parseError.errorCode = 0 then
			Set formValidatorObject = new Trophy_FormValidator
			Set forms = parserObj.getElementsByTagName("form")
			if forms.length > 1 then
				For i=0 to forms.length-1
					if forms.item(i).getAttribute("name") = formName then
						Set formObject = forms.item(i)
					end if
				Next
			else
				Set formObject = forms.item(0)
			end if
			formValidatorObject.form = formObject.getAttribute("name")
			if Not isNull(formObject.getAttribute("redirect")) then
				redirectPage = formObject.getAttribute("redirect")
			else
				redirectPage = req("HTTP_REFERER")
			end if
			formValidatorObject.redirect = redirectPage
			Set fields = formObject.getElementsByTagName("field")
			For x=0 to fields.length-1
				Call addField(formValidatorObject,fields(x))
			Next
			Set parserObj = Nothing
			Set forms = Nothing
			Set formObject = Nothing
		else
			Call objError.trigger("form-validator","cannot load form-validator: " & parserObj.parseError.reason & " file " & filePath)
			Exit Function
		end if
		Set loadXML = formValidatorObject
	End Function
	
	Public Function getFormValidatorInstance(vfile)
		validationXML = vfile
		Set getFormValidatorInstance = loadXML()
	End Function

end class

%>