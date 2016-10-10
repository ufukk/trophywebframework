<%

class Trophy_Form

	Private formFields
	Private fieldIds
	Private formCustoms
	Private formLastError
	Private hasValidation
	Private fieldIndex
	Private formName
	Private formAction
	Private formMethod
	Private validatorObject
	Public formSubmitLabel
	Public formEncType
	Public formTemplate
	Public fieldTemplate
	Public validationTemplate
	Public formTools
	
	Public Sub Class_Initialize() 
		Set formFields = newDictionary()
		Set fieldIds = newDictionary()
		Set formCustoms = newDictionary()
		Set formTools = new Trophy_FormTools
		hasValidation = False
		fieldIndex = 0
	End Sub
	
	Public Property Get name
		name = formName
	end Property
	
	Public Property Let name(n)
		formName = n
	End Property
	
	Public Property Get action
		action = formAction
	End Property
	
	Public Property Let action(a)
		formAction = a
	End Property
	
	Public Property Get method
		method = formMethod
	End Property
	
	Public Property Let method(m)
		if m = "GET" or m = "POST" then
			formMethod = m
		end if
	End Property
	
	Public Property Get onSubmit
		if validate then
			onSubmit = "return cF_" & name & "(this)"
			Exit Property
		end if
	End Property
	
	Public Property Get LastError
		LastError = FormLastError
	End Property
	
	Public Property Get validate
		validate = hasValidation
	End Property
	
	Public Property Get validator
		if isObject(validatorObject) then
			Set validator = validatorObject
		end if
	End Property
	
	Public Property Let validator(v)
		if isObject(v) then
			Set validatorObject = v
			hasValidation = true
		end if
	End Property
	
	Public Property Let validate(v)
		if v = true then
			hasValidation = true
			if Not isObject(validatorObject) then
				Set validatorObject = new Trophy_FormValidator
				validatorObject.form = formName
			end if
		elseif v = false then
			hasValidation = false
		end if
	End Property
	
	Public Property Let validationFile(f)
		Set objFormR = new Trophy_FormValidator_XMLReader
		Set validatorObject = objFormR.getFormValidatorInstance(f)
		validate = true
		Set objFormR = Nothing
	End Property
	
	Public Property Get fields
		Set fields = formFields
	End Property
	
	Public Default Property Get field(name)
		Set field = formFields.Item(getFieldIdByName(name))
	End Property
	
	Public Property Get submitLabel
		submitLabel = formSubmitLabel
	End Property
	
	Public Property Let submitLabel(s)
		formSubmitLabel = s
	End Property
	
	Public Property Get header
		Dim r
		r = "<FORM name=""" & name & """ action=""" & action & """ method=""" & method & """"
		'if method = "POST" then
		'	r = r & " enctype=""multipart/form-data"""
		'end if
		if validate then
			r = r & " onSubmit=""return cF_" & name & "(this)"""
		end if
		if formEncType <> "" then
			r = r & " enctype=""" & formEncType & """"
		end if
		r = r & ">"
		header = r
	End Property
	
	Public Property Get footer
		footer = "</FORM>"
	End Property
	
	Private Function getFieldIdByName(name)
		getFieldIdByName = fieldIds.Item(name)
	End Function
	
	Private Function addFieldId(name)
		Call fieldIds.Add(name,fieldIndex)
		fieldIndex = fieldIndex+1
	End Function
	
	Private Function addFieldObject(fieldObject)
		Call formFields.Add(fieldIndex,fieldObject)
		Call addFieldId(fieldObject.name)
	End Function
	
	Private Function setTemplateDefaults()
		if Len(formTemplate) < 1 then
			formTemplate = objTrophy.Config.parameters("default-form-template")
		end if
		if Len(fieldTemplate) < 1 then
			fieldTemplate = objTrophy.Config.parameters("default-field-template")
		end if
		if Len(validationTemplate) < 1 then
			validationTemplate = objTrophy.Config.parameters("default-form-validation-template")
		end if
		if Len(formSubmitLabel) < 1 then
			formSubmitLabel = objTrophy.Config.parameters("default-submit-label")
		end if
	End Function
	
	Public Function addField(type_,title,name,value)
		Set fieldObject = new Trophy_FormField
		fieldObject.fType = type_
		fieldObject.name = name
		fieldObject.title = title
		fieldObject.value = value
		Call addFieldObject(fieldObject)
	End Function
	
	Public Function addFieldByRef(ByRef fieldObject)
		Call addFieldObject(fieldObject)
	End Function
	
	Public Function addTextField(title,name,value)
		Call addField("textfield",title,name,value)
	End Function
	
	Public Function addPassword(title,name,value)
		Call addField("password",title,name,value)
	End Function
	
	Public Function addHidden(name,value)
		Call addField("hidden","",name,value)
	End Function
	
	Public Function addSelectBox(title,name,keys,values,selected)
		Call addField("selectbox",title,name,"")
		For i=0 to UBound(keys)
			Call addFieldOption(name,keys(i),values(i))
		Next	
		Call setFieldAttribute(name,"selected",selected)
	End Function
	
	Public Function addCheckBox(title,name,value,ischecked)
		Call addField("checkbox",title,name,value)
		if ischecked then
			Call setFieldAttribute(name,"checked",True)
		end if
	End Function
	
	Public Function addRadioButton(title,name,values,selected)
		Call addField("radio",title,name,"")
		For i=0 to UBound(values)
			Call addFieldOption(name,values(i),"")
		Next
		if Not IsNull(selected) then
			Call setFieldAttribute(name,"selected",selected)
		end if
	End Function
	
	Public Function addTextArea(title,name,value,rows,cols)
		Call addField("textarea",title,name,value)
		Call setFieldAttribute(name,"rows",rows)
		Call setFieldAttribute(name,"cols",cols)
	End Function
	
	Public Function addFile(title,name,value)
		Call addField("file",title,name,value)
	End Function
	
	Public Function addString(fname,value)
		Call formFields.Add(fieldIndex,value)
		Call addFieldId(fname)
	End Function
	
	Public Function removeField(name)
		if formFields.Exists(getFieldIdByName(name)) then
			Call formFields.Remove(getFieldIdByName(name))
		End if
	End Function
	
	Public Function getFieldTitle(name)
		getFieldTitle = field(name).title
	End Function
	
	Public Function setFieldAttribute(name,key,value)
		Call field(name).setAttribute(key,value)
	End Function
	
	Public Function addFieldOption(name,key,value)
		Call field(name).addOption(key,value)
	End Function
	
	Public Function setFieldValue(f,v)
		field(f).value = v
	End Function
	
	Public Function setFieldTitle(f,t)
		field(f).title = t
	End Function
	
	Public Function setFieldTemplate(f,t)
		field(f).template = t
	End Function
	
	Public Function setFieldValidation(n,r,m,c)
		Call validatorObject.addField(n,r,m,c)
	End Function
	
	Public Function setFieldAttributes(f,attributes)
		attributeList = split(attributes,";")
		For i=0 to UBound(attributeList)
			Call field(f).setAttribute(Mid(attributeList(i),1,InStr(attributeList(i),"=")-1),Mid(attributeList(i),InStr(attributeList(i),"=")+1))
		Next
	End Function
	
	Private Function compile()
		Dim output
		output = ""
		Dim foutput
		Call setTemplateDefaults()
		For Each f in formFields
			if isObject(formFields.Item(f)) then
				Set fieldObject = formFields.Item(f)
				if fieldObject.fType = "hidden" then
					foutput = fieldObject.toString()
				else
					Call objTemplate.assignByRef("fieldobj",fieldObject)
					if IsNull(fieldObject.template) then
						templateFile = fieldTemplate
					else
						templateFile = fieldObject.template
					end if
					foutput = objTemplate.fetchFile(templateFile)
				end if
				Set fieldObject = Nothing
			else
				foutput = formFields.Item(f)
			end if
			output = output & foutput
		Next
		Call objTemplate.assignByRef("form",Me)
		Call objTemplate.assign("fields",formFields)
		if isObject(validatorObject) then
			Call objTemplate.assignByRef("formValidatorObject_fields",validatorObject.fields)
		end if
		if FormMethod = "POST" then
			Call objTemplate.assign("ENCTYPE_STR"," enctype=" & """" & "multipart/form-data" & """")
		end if
		Call objTemplate.assign("FIELDS",output)
		compile = objTemplate.fetchFile(formTemplate)
	End Function
	
	Public Function fetch()
		fetch = compile()
	End Function
	
	Public Function display()
		echo compile()
	End Function

end class




%>