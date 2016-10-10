<%


class Trophy_BusinessTools

	
	Public Sub Class_Initialize()
	End Sub
	
	Public Function callUserFunction(ByRef req,ByRef parr,ByRef narr,cl,op)
		Dim obj,coll,o,cmd,r
		Execute("Set obj = new " & cl)
		o = ""
		Set coll = req.Iterator
		Dim val
		For Each key In coll
			if Not inArray(narr,CStr(key)) then
				val = req(key)
				if Not isInteger(val) then
					val = """" & val & """"
				end if
				o = o & val & ","
			end if
		Next
		For a=0 To UBound(parr)
				o = o & """" &  parr(a) & """" & ","
		Next
		o = str_unpush(o)
		cmd = "r = obj." & op & "(" & o & ")"
		Call Execute(cmd)
		callUserFunction = r
	End Function

end class



%>