<%


class Trophy_TemplateContainer_Form

	Private containerID
	Private formObject
	Private templateObject
	Private labels
	
	Public Property Get ID
		ID = containerID
	End Property
	
	Public Property Let ID(i)
		ID = i
	End Property
	
	Public Property Get templateObj
		Set templateObj = templateObject
	End Property
	
	Public Property Let templateObj(o)
		Set templateObject = o
	End Property
	
	Public Property Get header
		header = formObject.header
	End Property
	
	Public Property Get footer
		footer = formObject.footer
	End Property
	
	Public Property Get testFormObject
		testFormObject = formObject.validator.form
	End Property
	
	Private Function setFormFieldProperties(ByRef element,attributeMap)
		Dim validation,message,client
		For Each name in attributeMap
			value = attributeMap.Item(name)
			Select Case name
				Case "type"
				Case "ftype"
					element.fType = value
				Case "name"
					element.name = value
				Case "title" 
					element.title = value
				Case "value"
					element.value = stripQuotes(value)
				Case "validation"
					validation = value
				Case "message"
					message = value
				Case "client"
					client = CBool(value)
				Case Else
					Call element.setAttribute(name,value)
			End Select
		Next
		if validation <> "" then
			Call formObject.setFieldValidation(element.name,validation,message,client)
		end if
	End Function
	
	Private Function setNonValidatedField(ByRef obj)
		if req.isNonValidated then
			if obj.name <> "" then
				obj.value = req.getNonValidatedValue(obj.name)
			end if
		end if
		For Each fieldName in req.getValidatorMessages()
			if fieldName = obj.name then
				Call obj.setAttribute("class","nonvalidated_field")
			end if
		Next
	End Function
	
	Private Function setNonValidatedLabel(ByRef attributeMap)
		For Each fieldName in req.getValidatorMessages()
			if fieldName = attributeMap.Item("for") then
				attributeMap.Item("class") = "nonvalidated_label"
				attributeMap.Item("msg") = req.getValidatorMessages().Item(fieldName)
			end if
		Next
	End Function
	
	Private Function getElementFieldOutput(ByRef obj)
		Call setNonValidatedField(obj)
		getElementFieldOutput = obj.toString()
	End Function
	
	Private Function getElementLabelOutput(attributeMap)
		Dim o
		o = ""
		Call setNonValidatedLabel(attributeMap)
		For Each name in attributeMap
			if Not inArray(array("msg","value"),name) then
				o = o & name & "=" & addQuotes(attributeMap.Item(name)) & " "
			end if
		Next
		o = str_unpush(o)
		o = "<label " & o & ">" & attributeMap.Item("value") & "</label>"
		if attributeMap.Exists("msg") then
			o = "<div class=""nonvalidated_msg"">" & attributeMap.Item("msg") & "</div>" & o
		end if
		getElementLabelOutput = o
	End Function
	
	Public Function setPropertyMap(map)
		Dim name,value
		For Each name in map
			value = map.Item(name)
			Call setProperty(name,value)
		Next
	End Function
	
	Public Function setProperty(name,value)
		Select Case name
			Case "name"
				formObject.name = value  
			Case "action"
				formObject.action = value
			Case "method"
				formObject.method = value
			Case "validate"
				formObject.validate = CBool(value)
			Case "validationfile"
				Set xmlValidatorObject = new Trophy_FormValidator_XMLReader
				formObject.validator = xmlValidatorObject.getFormValidatorInstance(value)
				Set xmlValidatorObject = Nothing
			Case "enctype"
				formObject.formEncType = value
			Case Else
		End Select
	End Function
	
	Public Sub class_Initialize()
		Set formObject = new Trophy_Form
		Set labels = newDictionary()
	End Sub
	
	Public Sub class_terminate()
		Set formObject = Nothing
	End Sub
	
	Public Function addElement(etype,attributeMap)
		Select Case etype
			Case "field"
				Set obj = new Trophy_FormField	
				Call setFormFieldProperties(obj,attributeMap)
				addElement = getElementFieldOutput(obj)
				Exit Function
			Case "label"
				Call labels.Add(attributeMap.Item("for"),attributeMap.Item("value"))
				addElement = getElementLabelOutput(attributeMap)
			Case Else
		End Select	
	End Function
	
	Public Function addElementByRef(etype,ByRef obj)
		Select Case etype
			Case "field"
				Call formObject.addFieldByRef(obj)
				addElementByRef = getElementOutput(obj)
				Exit Function
			Case Else
		End Select
	End Function 
	
	Public Function setNonValidatedProperties(ByRef field)
		
	End Function
	
	Public Function starter() 
	End Function

	Public Function finisher()
		Call templateObject.assign("formObject",formObject)
		if isObject(formObject.validator) then
			Call templateObject.assign("formValidatorObject_fields",formObject.validator.fields)
		end if
		Call templateObject.assign("formValidatorObject",formObject.validator)
	End Function
		
end class
%>