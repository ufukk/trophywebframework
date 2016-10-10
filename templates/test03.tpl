

[!if condition="1 > 0"]
	[!loop name="LOOP_1" type="for" start="0" end="30"]
		[!#LOOP_1]
		[!if condition="{#LOOP_1} > 15"]
			[!#LOOP_1]aaa
		[!/if]
	[!/loop]
[!/if]
	
	[!if condition=`{VAR_3} <> ""`]
		asdasdad
	[!/if]
	
	[!if condition=`{VAR_4} <> ""`]
		asdasdad
	[!/if]
	
	
	[!if condition=`{VAR_5} <> ""`]
		asdasdad
	[!else]
		qqqqq
	[!/if]
	
	[!VAR_3]