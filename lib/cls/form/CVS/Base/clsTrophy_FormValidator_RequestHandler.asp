<%


class Trophy_FormValidator_RequestHandler
	
	Private formValidatorObject
	
	Public Property Get field(n)
		field = formValidatorObject.field(n)
	End Property
	
	Public Property Get message(n)
		message = formValidatorObject.field(n).message
	End Property
	
	Public Property Let validationFile(f)
		Call loadXMLReader(f)	
	End Property
	
	Private Function getFieldValue(fieldName)
		getFieldValue = req(fieldName)
	End Function
	
	Private Function loadXMLReader(path)
		Dim xmlReader
		Set xmlReader = new Trophy_FormValidator_XMLReader
		Set formValidatorObject = xmlReader.getFormValidatorInstance(path)
		Set xmlReader = Nothing
	End Function
	
	Private Function recordFields(fieldDict)
		Call Session.Contents.Remove(CONFIG_SESSION_REQUEST_KEY)
		Call Session.Contents.Remove(CONFIG_SESSION_VALIDATOR_KEY)
		Set dict = newDictionary()
		For Each item in req.iterator()
			Call dict.Add(item,req(item))
		Next
		Set Session(CONFIG_SESSION_REQUEST_KEY) = dict
		Set Session(CONFIG_SESSION_VALIDATOR_KEY) = fieldDict
	End Function
	
	Public Sub class_initialize()
		Set formFields = newDictionary()
		Set fieldMessages = newDictionary()
	End Sub
	
	Public Sub class_terminate()
		Set formFields = Nothing
		Set fieldMessages = Nothing
	End Sub
	
	Public Function validate() 
		Dim validated
		validated = true
		Set validatefields = formValidatorObject.fields
		Set fieldDict = newDictionary()
		For Each fieldObject in validateFields
			if fieldObject.serverValidate then
				if Not fieldObject.validate then
					fieldName = fieldObject.Name
					validated = false
					if Not fieldDict.Exists(fieldName) then
						Call fieldDict.Add(fieldName,fieldObject.message)
					end if
				end if
			end if
		Next
		if validated = false then
			if formValidatorObject.redirect <> "" then
				redirectPage = formValidatorObject.redirect
			else
				redirectPage = req("HTTP_REFERER")
			end if
			Call recordFields(fieldDict)
			Call redirect(redirectPage)
			Call objTrophy.stopPage()
		end if
	End Function
end class

%>