<%


class Trophy_FormTools


	Public Sub Class_Initialize()
	End Sub
	
	Public Function getRSAsComboBox(ByVal ftype,ByRef rs,keyfield,valuefield,name,title,selected,fitem)
		Set objF = new Trophy_FormField
		objF.FType = ftype
		objF.Name = name
		objF.Title = title
		if Not IsNull(fitem) then
			Call objF.addOption("",fitem)
		end if
		Do While Not rs.EOF
			Call objF.addOption(CStr(rs(keyfield)),CStr(rs(valuefield)))
			rs.movenext
		Loop
		if Not IsNull(selected) then
			Call objF.setAttribute("selected",selected)
		end if
		Set getRSAsComboBox = objF
	End Function
	
	Public Function getDayBox(ftype,name,selected,onchange)
		Dim F,i
		Set F = new Trophy_Field
		F.FieldType = ftype
		F.FieldName = name
		Call F.setAttribute("onChange",onchange)
		For i = 1 to 31
			Call F.addOption(i,i)
		Next
		if selected <> "" then
			Call F.setAttribute("selected",selected)
		end if
		Call F.setAttribute("leftMargin","3")
		Call F.setAttribute("width","60")
		Set getDayBox = F
	End Function
	
	Public Function getMonthBox(ftype,name,selected,onchange)
		Dim F,i
		Set F = new Trophy_Field
		F.FieldType = ftype
		F.FieldName = name
		Call F.setAttribute("onChange",onchange)
		For i = 1 to 12
			Call F.addOption(i,MonthName(i))
		Next
		if selected <> "" then
			Call F.setAttribute("selected",selected)
		end if
		Call F.setAttribute("leftMargin","3")
		Call F.setAttribute("width","80")
		Set getMonthBox = F
	End Function
	
	Public Function getYearBox(ftype,name,selected,onchange)
		Dim F,i
		Set F = new Trophy_Field
		F.FieldType = ftype
		F.FieldName = name
		Call F.setAttribute("onChange",onchange)
		For i = Year(Now)-3 to Year(Now)
			Call F.addOption(i,i)
		Next
		if selected <> "" then
			Call F.setAttribute("selected",selected)
		end if
		Call F.setAttribute("leftMargin","3")
		Call F.setAttribute("width","80")
		Set getYearBox = F
	End Function
	
	Public Function getFormByRs(ByRef rs,ByRef fdict,ByRef titles,name,method,action,onsubmit,enctype)
		Dim oF,f,fieldtype,fObj
		if isNull(fdict) then
			Set fdict = newDictionary()
		end if
		if isNull(titles) then
			Set titles = newDictionary()
		end if
		Set oF = new Trophy_Form
		oF.FormName = name
		oF.FormMethod = method
		oF.FormAction = action
		if Not IsNull(onsubmit) then
			oF.FormOnSubmit = onsubmit
		end if
		if Not IsNull(enctype) then
			oF.FormEncType = enctype
		end if
		For Each f in rs.Fields
			fObj = 0
			if fdict.Exists(f.Name) then
				if IsObject(fdict.Item(f.Name)) then
					if TypeName(fdict.Item(f.Name))="Trophy_Field" then
						Set fObj = fdict.Item(f.Name)
					end if			
				else
					fieldtype = fdict.Item(f.Name)
				end if
			else
				'primary key,identity
				if f.Attributes = 16 then
					fieldtype = "hidden"
				else
					'determine form field type by field data type
					dtype = f.Type
					defaultValue = ""
					Select Case dtype
						Case adInteger
							fieldtype = "textfield"
							defaultValue = 0
						Case adBoolean
							fieldtype = "checkbox"
							defaultValue = 0
						Case adVarChar
							fieldtype = "textfield"
						Case adChar
							fieldtype = "textfield"
						Case adWChar
							fieldtype = "textfield"
						Case adLongVarChar
							fieldtype = "textfield"
						Case adVarWChar
							fieldtype = "textfield"
						Case adLongVarWChar
							fieldtype = "textarea"
						Case Else
							fieldtype = "textfield"
					End Select
				end if
			end if
			if titles.Exists(f.Name) then
				title = titles.Item(f.Name) 
			else
				title = f.Name
			end if
			if Not rs.EOF then
				value = f.Value
			else
				value = defaultValue
			end if
			if fieldtype = "checkbox" then
				value = "1"
			end if
			if Not IsObject(fObj) then
				Call oF.addField(fieldtype,title,f.Name,value)
			else
				Call oF.addFieldByRef(fObj)
			end if
		Next
		Set getFormByRs = oF
	End Function

end class


%>