<%


class Person

	Public primary
	Public tab
	Public personTest_id
	Public fname
	Public lname
	Public email
	Public objDb
	Public objectName
	
	Public Property Get test_id
		test_id = personTest_id
	End Property
	
	Public Property Let test_id(t)
		if t <> "" then
			personTest_id = CInt(t)
		end if
	End Property
	
	Public Sub class_initialize()
		objDb = Null
	End Sub
	
	Public Function list()
		Call objDb.runQuery("SELECT * FROM test_table")
	End Function
	
	Public Function addnew()
		Set builderObj = new Trophy_SQLBuilder
		builderObj.tableName = "test_table"
		builderObj.primaryField = "test_id"
		builderObj.propertyFields = objTrophy.selectedAction.userObjects.Item(objectName).properties
		Set excluded = newDictionary()
		Call excluded.Add("test_id","")
		builderObj.excludedFields = excluded
		builderObj.objDb = objDb
		builderObj.userObject = Me
		builderObj.processType = "INSERT"
		builderObj.executeUpdate()
		Set builderObj = Nothing
	End Function
	
	Public Function update()
		Set builderObj = new Trophy_SQLBuilder
		builderObj.tableName = "test_table"
		builderObj.primaryField = "test_id"
		builderObj.propertyFields = objTrophy.selectedAction.userObjects.Item(objectName).properties
		builderObj.objDb = objDb
		builderObj.userObject = Me
		builderObj.processType = "UPDATE"
		builderObj.executeUpdate()
		Set builderObj = Nothing
	End Function
	
	Public Function getSelectedRecord()
		Set builderObj = new Trophy_SQLBuilder
		builderObj.tableName = "test_table"
		builderObj.primaryField = "test_id"
		builderObj.propertyFields = objTrophy.selectedAction.userObjects.Item(objectName).properties
		builderObj.objDb = objDb
		builderObj.userObject = Me
		Call builderObj.getByPrimaryKey()
	End Function
	
end class


%>