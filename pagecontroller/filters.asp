<%

Public Function requestFilter_fixInjection(f)
	Dim output
	f = replace(f,"'","''")
	requestFilter_fixInjection = f
End Function

Public Function requestFilter_fixCSS(f)
	f = replace(f,"<","&lt;")
	f = replace(f,">","&gt;")
	requestFilter_fixCSS = f
End Function

Public Function requestFilter_void(f)
	requestFilter_void = f
End Function
%>