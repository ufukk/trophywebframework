<%

class Trophy_FormField

	Private fieldAttributes
	Private fieldOptions 
	Private fieldType
	Private fieldTitle
	Private fieldName
	Private fieldValue
	Private fieldValidation
	Private fieldTemplate
	Public autoGenerateID
	Public setDefaults
	
	Public Property Get title 
		title = fieldTitle 
	End Property
	
	Public Property Let title(t) 
		fieldTitle = t
	End Property
	
	Public Property Get name
		name = fieldName
	End Property
	
	Public Property Let name(n)
		fieldName = n
	End Property
	
	Public Property Get fType
		fType = fieldType
	End Property
	
	Public Property Let fType(t)
		fieldType = t
	End Property
	
	Public Property Get value
		value = fieldValue
	End Property
	
	Public Property Let value(v)
		fieldValue = v
	End Property
	
	Public Property Get template
		template = fieldTemplate
	End Property
	
	Public Property Let template(t)
		fieldTemplate = t
	End Property
	
	Public Property Get validation
		validation = replace(fieldValidation,"%VALUE","frm." & name)
	End Property
	
	Public Property Let validation(v)
		fieldValidation = v
	End Property 
	
	Public Property Get validate
		Dim r
		if fieldValidation <> "" then
			r = true
		else
			r = false
		end if
		validate = r
	End Property
	
	Private Function setDefaultAttributes()
		if isObject(objTemplate) then
			Set attrDict = objTemplate.getVariablesByParent("form.fields")
			Call setAttributesFromDictionary(attrDict)
			Set attrDict = objTemplate.getVariablesByParent("form.fields." & fieldType)
			Call setAttributesFromDictionary(attrDict)
		end if
	End Function
	
	Private Function setAttributesFromDictionary(dict)
		For Each attr in dict
			Call setAttribute(attr,dict.Item(attr))
		Next
	End Function
	
	Private Function getStandartOutput()
		attributestring = getAttributeString()
		output = "<input type=" & """" & fieldType & """" & " name=" & """" & fieldName & """"
		output = " " & output & " value=" & """" & fieldValue & """"
		output = output & attributestring
		output = output & ">"
		getStandartOutput = output
	End Function
	
	Private Function getAttributeString()
		Dim output
		if setDefaults = True then
			Call setDefaultAttributes()
		end if
		For Each key in fieldAttributes
			if key <> "selected" And key <> fieldName then
				output = output & " " & key & "=" & """" & fieldAttributes.Item(key) & """"
			end if
		Next
		if Not fieldAttributes.Exists("id") And autoGenerateID then
			output = output & " " & "id" & "=" & """" & fieldAttributes.Item(name) & """"
		end if
		getAttributeString = output
	End Function
	
	Public Sub class_Initialize()
		Set fieldAttributes = newDictionary()
		Set fieldOptions = newDictionary()
		autoGenerateID = True
		setDefaults = True
		fieldTemplate = Null
	End Sub
	
	Public Sub class_Terminate()
		Set fieldAttributes = Nothing
		Set fieldOptions = Nothing
	End Sub
	
	Public Function setAttribute(key,value)
		Call removeAttribute(key)
		Call FieldAttributes.Add(key,value)
	End Function
	
	Public Function removeAttribute(key)
		if fieldAttributes.Exists(key) then 
			Call fieldAttributes.Remove(key)
		end if
	End Function
	
	Public Function getAttribute(key)
		Dim result
		if fieldAttributes.Exists(key) then
			result = fieldAttributes.Item(key)
		else
			result = Null
		end if
		getAttribute = result
	End Function
	
	Public Function addOption(key,value)
		Call removeOption(key)
		Call fieldOptions.Add(CStr(key),CStr(value))
	End Function
	
	Public Function removeOption(key)
		if fieldOptions.Exists(key) then
			Call FieldOptions.Remove(key)
		end if
	End Function
	
	Public Default Function toString()
		Dim output
		output = ""
		attributestring = getAttributeString()
		Select Case FieldType
			Case "hidden"
				output = getStandartOutput()
			Case "textfield"
				output = getStandartOutput()
			Case "password"
				output = getStandartOutput()
			Case "radiobutton"
				ReDim output(FieldOptions.Count)
				Dim count
				count = 0
				For Each soption in fieldOptions
					output(count) = getStandartOutput()
				Next 
			Case "checkbox"
				attributestring = getAttributeString()
				output = "<input type=" & """" & fieldType & """" & " name=" & """" & fieldName & """"
				output = " " & output & " value=" & """" & fieldValue & """"
				output = output & attributestring
				if getAttribute("checked") = True then
					output = output & " checked "
				end if
				output = output & ">"
			Case "selectbox"
				Dim optionStr,sstr
				For Each optionkey in FieldOptions
					if CStr(getAttribute("selected")) = CStr(optionkey) then
						sstr = " selected"
					else
						sstr = ""
					end if
					optionStr = optionStr & "<option value=" & """" & optionkey & """"
					optionStr = optionStr & sstr
					optionStr = optionStr & ">" & FieldOptions.Item(optionkey) & "</option>"
				Next
				output = "<select name=" & """" & fieldName & """" & " "
				output = output & attributestring
				output = output & ">"
				output = output & optionStr
				output = output & "</select>"
			Case "textarea"
				output = "<textarea name=" & """" & fieldName & """"
				output = output & attributestring
				output = output & ">" & fieldValue & "</textarea>"
			Case "file"
				output = getStandartOutput()
			Case Else
				output = getStandartOutput()
		End Select
		toString = output
	End Function

end class










%>