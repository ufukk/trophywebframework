<%


class Vector

	Private stackObject
	Private lastIndex
	Public currentIndex
	Public tempStack
	Public t
	
	Private Function addToStack(element)
		lastIndex = lastIndex+1
		Call stackObject.Add(lastIndex,element)
	End Function
	
	Public Sub class_initialize()
		Set stackObject = newDictionary()
		currentIndex = Null
	End Sub
	
	Public Sub class_terminate()
		Set stackObject = Nothing
	End Sub
	
	Public Function add(element)
		Call addToStack(element)
	End Function
	
	Public Function getElement(index)
		z =0
		For Each key in stackObject
			if index=z then
				if isObject(stackObject.Item(key)) then
					Set getElement = stackObject.Item(key)
				else
					getElement = stackObject.Item(key)
				end if
				Exit Function
			end if
			z = z+1
		Next
	End Function
	
	Public Function remove(index)
		z = 0
		For Each key in stackObject
			if z = index then
				Call stackObject.Remove(key)
				Exit Function
			end if
			z = z+1
		Next
	End Function
	
	Public Function removeAll()
		Call stackObject.removeAll()
		lastIndex = 0
	End Function
	
	Public Function sort(sortType)
		if Not isObject(tempStack) then
			Set tempStack = newDictionary()
		end if
	End Function
	
end class


%>