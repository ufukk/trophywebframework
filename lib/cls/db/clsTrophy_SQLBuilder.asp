<%


class Trophy_SQLBuilder

	Private sqlPropertyFields
	Private sqlExcludedFields
	Private sqlTableName
	Private sqlPrimaryField
	Private sqlProcessType
	Private sqlUserObject
	Private sqlObjDb
	
	Public Property Get propertyFields
		Set propertyFields = sqlPropertyFields
	End Property
	
	Public Property Let propertyFields(p)
		Set sqlPropertyFields = p
	End Property
	
	Public Property Get excludedFields
		Set excludedFields = sqlExcludedFields
	End Property
	
	Public Property Let excludedFields(e)
		Set sqlExcludedFields = e
	End Property
	
	Public Property Get tableName
		tableName = sqlTableName
	End Property
	
	Public Property Let tableName(t)
		sqlTableName = t
	End Property
	
	Public Property Get primaryField
		primaryField = sqlPrimaryField
	End Property
	
	Public Property Let primaryField(p)
		if Not isArray(p) then
			p = array(p)
		end if
		sqlPrimaryField = p
	End Property
	
	Public Property Get processType
		processType = sqlProcessType
	End Property
	
	Public Property Let processType(p)
		sqlProcessType = p
	End Property
	
	Public Property Get userObject
		Set userObject = sqluserObject
	End Property
	
	Public Property Let userObject(ByRef o)
		Set sqlUserObject = o
	End Property
	
	Public Property Get objDb
		Set objDb = sqlObjDb
	End Property
	
	Public Property Let objDb(ByRef o)
		Set sqlObjDb = o
	End Property
	
	Public Sub class_initialize()
		tableName = Null
		Set sqlExcludedFields = newDictionary()
	End Sub
	
	Private Function getFieldValue(key)
		Dim fieldValue,propCmd
		propertyCmd = "fieldValue = userObject." & key
		Execute(propertyCmd)
		getFieldValue = fieldValue
	End Function
	
	Private Function getFieldValueQuotes(key)
		fieldValue = getFieldValue(key)
		if Not isNumeric(fieldValue) then
			fieldValue = addSingleQuotes(fieldValue)
		end if
		getFieldValueQuotes = fieldValue
	End Function
	
	Private Function runUpdateQuery()
		Dim updateSql
		Select Case processType
			Case "UPDATE"
				updateSql = "SELECT * FROM " & tableName & " WHERE "
				For i=0 to UBound(sqlPrimaryField) 
					updateSql = updateSql & " " & sqlPrimaryField(i) & "=" & getFieldValueQuotes(sqlPrimaryField(i))
					if i <> UBound(sqlPrimaryField) then
						updateSql = updateSql & " AND "
					end if
				Next
				Call objDb.runQueryForUpdate(updateSql)
			Case "INSERT"
				updateSql = " SELECT * FROM [" & tableName & "] WHERE 0=1"
				Call objDb.runQueryForUpdate(updateSql)
				objDb.objRs.AddNew
		End Select
	End Function

	Private Function assignRecordset()
		For Each key in propertyFields
			if Not excludedFields.Exists(key) And Not inArray(sqlPrimaryField,key) then
				objDb.objRs(key) = getFieldValue(key)
			end if
		Next
	End Function
	
	Public Function executeUpdate()
		Call runUpdateQuery()
		Call assignRecordset()
		objDb.objRs.Update
		objDb.objRs.Close
	End Function
	
	Public Function getByPrimaryKey()
		Dim fieldValue
		fieldValue = getFieldValue(primaryField)
		if Not TypeName(fieldValue) = "Integer" then
			fieldValue = addSingleQuotes(fieldValue)
		end if
		selectQuery = "SELECT * FROM " & tableName & " WHERE " & primaryField & " = " & fieldValue
		Call objDb.runQuery(selectQuery)
	End Function
	
end class




%>