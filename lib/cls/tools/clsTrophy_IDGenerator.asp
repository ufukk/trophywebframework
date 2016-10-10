<%


class Trophy_IDGenerator

	Private genUseIntegers
	Private genUseChars
	Private genUseLowerCase
	Private genUseUpperCase
	Private genUseSpecialChars
	Private genUseDate
	Private genUseMD5
	Private genLength
	Private charArray
	Private intArray
	Private specArray
	
	Public Property Get useIntegers
		useIntegers = genUseIntegers
	End Property
	
	Public Property Let useIntegers(u)
		genUseIntegers = CBool(u)
	End Property
	
	Public Property Get useChars
		useChars = genUseChars
	End Property
	
	Public Property Let useChars(u)
		genUseChars = CBool(u)
	End Property
	
	Public Property Get useLowerCase
		useLowerCase = genUseLowerCase
	End Property
	
	Public Property Let useLowerCase(u)
		genUseLowerCase = CBool(u)
	End Property
	
	Public Property Get useUpperCase
		useUpperCase = genUseUpperCase
	End Property
	
	Public Property Let useUpperCase(u)
		genUseUpperCase = CBool(u)
	End Property
	
	Public Property Get useSpecialChars
		useSpecialChars = genUseSpecialChars
	End Property
	
	Public Property Let useSpecialChars(u)
		genUseSpecialChars = CBool(u)
	End Property
	
	Public Property Get length
		length = genLength
	End Property
	
	Public Property Let length(u)
		genLength = CInt(u)
	End Property

	Public Property Get useMD5
		useMD5 = genUseMD5
	End Property
	
	Public Property Let useMD5(u)
		genUseMD5 = CBool(u)
	End Property
	
	Public Property Get useDate
		useDate = genUseDate
	End Property
	
	Public Property Let useDate(u)
		genUseDate = u
	End Property
	
	Public Sub class_initialize()
		'set defaults
		useMD5 = true
		useChars = true
		useIntegers = true
		useUpperCase = true
		useLowerCase = true
		useSpecialChars = false
		length = 20
		useDate = True
		Call setArrays()
	End Sub
	
	Private Function setArrays()
		charArray = "abcdefghijklmnoprqstvywxz"
		intArray = "0123456789"
		specArray = "_-.;,"
	End Function
	
	Private Function getMD5(str)
		Dim objMD5
		Set objMD5 = New MD5
		objMD5.Text = str
		getMD5 = objMD5.HEXMD5
	End Function
	
	Private Function getString()
		Dim rand,totalArray,selected
		if useIntegers then
			totalArray = totalArray & intArray
		end if
		if useChars then
			totalArray = totalArray & charArray
		end if
		if useSpecialChars then
			totalArray = totalArray & specArray
		end if
		arrayLength = Len(totalArray)-1
		Randomize
		For i=0 to length
			rand = Int(arrayLength*Rnd)
			if rand = 0 then
				rand = 1
			end if
			selectedChar = Mid(totalArray,rand,1)
			selected = selected & selectedChar
		Next 
		getString = selected
	End Function
	
	Public Function generate()
		Dim str
		str = getString()
		if useMD5 then
			if useDate then
				str = str & CStr(Now)
			end if
			str = getMD5(str)
		end if
		if useUpperCase then
			str = UCase(str)
		end if
		generate = str
	End Function
	
end class


%>