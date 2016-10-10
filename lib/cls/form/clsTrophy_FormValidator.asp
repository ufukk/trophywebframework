<%


class Trophy_FormValidator

	Private fieldObjects
	Private redirectPage
	Private formName
	
	Public Property Get redirect
		redirect = redirectPage
	End Property
	
	Public Property Let redirect(r)
		redirectPage = r
	End Property
	
	Public Property Get form
		form = formName
	End Property
	
	Public Property Let form(f)
		formName = f
	End Property
	
	Public Default Property Get field(n)
		Set field = fieldObjects.Item(n)
	End Property
	
	Public Property Get fields
		Set fields = fieldObjects
	End Property
	
	Private Function addFieldObject(fieldObject)
		Call fieldObjects.Add(fieldObject.name,fieldObject)
	End Function
	
	Public Sub class_initialize()
		Set fieldObjects = newDictionary()
	End Sub
	
	Public Function addField(name,rule,message,client,serverVal)
		Set fieldObject = new Trophy_FormValidator_Field
		fieldObject.name = name
		fieldObject.rule = rule
		fieldObject.message = message
		fieldObject.form = formName
		fieldObject.client = client
		fieldObject.serverValidate = serverVal
		Call fieldObjects.Add(fieldObject,"")
	End Function
	
	Public Function removeField(name)
		Call fieldObjects.remove(name)
	End Function
	
	  
end class



%>