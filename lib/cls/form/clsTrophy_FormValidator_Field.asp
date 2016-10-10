<%


class Trophy_FormValidator_Field

	Private fieldName
	Private fieldRule
	Private fieldMessage
	Private formName
	Private clientValidation
	Private serverValidation
	
	Public Property Get name
		name = fieldName
	End Property
	
	Public Property Let name(n)
		fieldName = n
	End Property
	
	Public Property Get rule
		rule = fieldRule
	End Property
	
	Public Property Let rule(r)
		fieldRule = r
	End Property
	
	Public Property Get message
		message = fieldMessage
	End Property
	
	Public Property Let message(m)
		fieldMessage = m
	End Property
	
	Public Property Get form
		form = formName
	End Property
	
	Public Property Let form(f)
		formName = f
	End Property
	
	Public Property Get client
		client = clientValidation
	End Property
	
	Public Property Let client(c)
		clientValidation = CBool(c)
	End Property
	
	Public Property Get serverValidate
		serverValidate = serverValidation
	End Property
	
	Public Property Let serverValidate(c)
		serverValidation = CBool(c)
	End Property
	
	Public Property Get jsCode
		jsCode = parseRuleForJS()
	End Property
	
	Private Function getFieldValue(f)
		getFieldValue = req(f)
	End Function
	
	Private Function parseRule()
		Dim rxp
		Set rxp = newRegexp()
		rxp.Pattern = "\{fields\.([a-zA-Z0-9_.]*)\}"
		str = rxp.Replace(rule,"getFieldValue(""$1"")")
		Set rxp = Nothing
		parseRule = replace(str,"field(this)","getFieldValue(" & addQuotes(fieldName) & ")")
	End Function
	
	Private Function parseRuleForJS()
		Dim rxp
		Set rxp = newRegexp()
		rxp.Pattern = "\{fields\.([a-zA-Z0-9_.]*)\}"
		str = rxp.Replace(rule,"frm.$1.value")
		Set rxp = Nothing
		parseRuleForJS = replace(str,"field(this)",formName & "." & fieldName & ".value")
	End Function
	
	Private Function getValidationCommand()
		getValidationCommand = parseRule()
	End Function
	
	Private Function minimumLength(value,length)
		minimumLength = Len(value) >= length
	End Function
	
	Private Function regex(value,pattern)
		Dim rxp,r,results,result
		Set rxp = newRegexp()
		rxp.Pattern = pattern
		Set results = rxp.Execute(value)
		r = false
		For Each result in results
			if result <> value then
				r = false
			else
				r = true
			end if
		Next
		Set rxp = Nothing
		regex = r
	End Function
	
	Private Function isAlphaNumeric(val) 
		isAlphaNumeric = regex(val,"([a-zA-Z0-9]{2,255})")
	End Function
	
	Private Function isAlpha(val) 
		isAlpha = regex(val,"([a-zA-Z]{2,255})")
	End Function
	
	Private Function isValidEmail(email)
		isValidEmail = regex(email,"[a-zA-Z0-9._-]{1,255}@[a-zA-Z0-9_.-]{1,255}.[a-zA-Z0-9]{1,10}")
	End Function
	
	Private Function equals(val1,val2)
		r = false
		if val1=val2 then
			r = true
		end if
		equals = r
	End Function
	
	Public Sub class_initialize()
		client = true
	End Sub

	Public Function validate()
		Dim validationResult
		validationCommand = "validationResult = " & getValidationCommand()
		Execute(validationCommand)
		validate = validationResult
	End Function

end class


%>