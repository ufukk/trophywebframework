<%


class Trophy_FormGenerator_Recordset

	Private objDb
	Private recordsetObject
	Private formObject
	Private foreignKeys
	Private options
	Private tableName
	
	Public Property Get recordset
		Set recordset = recordsetObject
	End Property
	
	Public Property Let recordset(r)
		Set recordsetObject = r
	End Property
	
	Public Property Let dbObject(ByRef db)
		Set objDb = db
	End Property
	
	Public Property Get table
		table = tableName
	End Property
	
	Public Property Let table(t)
		tableName = t
	End Property
	
	Private Function createFieldFromForeignKey(foreignTable,fieldName,formFieldType,selectedValue)
		Dim formfieldObject,rs
		valuefieldName = fieldName
		Set rs = objDb.returnQuery("SELECT " & fieldName & " FROM " & foreignTable & " ORDER BY " & fieldName)
		Select Case formFieldType
			Case "selectbox"
				Set formfieldObject = formToolsObject.getRSAsComboBox("selectbox",rs,fieldName,valuefieldName,fieldName,fieldName,selectedValue,"")
				Call options.Add(fieldName,formfieldObject)
			Case Else
		End Select
	End Function
	
	Private Function findForeignKeys()
		Set rs = objDb.returnQuery("EXEC sp_fkeys @fktable_name='" & tableName & "'")
		Do While Not rs.EOF
			Call createFieldFromForeignKey(rs("PKTABLE_NAME"),rs("PKCOLUMN_NAME"),"selectbox","")
			rs.movenext
		Loop
	End Function
	
	Private Function createForm()
	
	End Function
	
	Public Sub class_initialize()
		Set options = newDictionary()
	End Sub
	

end class


%>