
<script language="javascript">
	/*Auto Generated*/
	function cF_[!formObject .name](frm) {
	[!loop name="VALIDATOR_LOOP" var="formValidatorObject_fields"] [!if condition="{#VALIDATOR_LOOP .client} = true"]
		if(frm.[!#VALIDATOR_LOOP .name]) {
			if (![!#VALIDATOR_LOOP .jsCode]) {
				alert('[!#VALIDATOR_LOOP .message]');
				frm.[!#VALIDATOR_LOOP .name].focus();
				return false;
			}
		}
		[!/if] [!/loop] 
		return true;
	}
	/*Auto Generated End*/
		
	function isInteger(val) {
		 return (parseInt(val)); 
	}
	
	function minimumLength(val,len) {
		return (val.length >= len);
	}
	
	function betweenLength(val,minlen,maxlen) {
		return ( (val.length >= minlen) && (val.length <= maxlen) );
	}
	
	function equals(v1,v2) {
		return (v1==v2);
	}
	
	function isAlphaNumeric(val) {
		return regex(val,"([a-zA-Z0-9]{2,255})");
	}
	
	function isAlpha(val) {
		return regex(val,"([a-zA-Z]{2,255})");
	}
	
	function isValidEmail(val) {
		return regex(val,"[a-zA-Z0-9._-]{1,255}@[a-zA-Z0-9_.-]{1,255}.[a-zA-Z0-9]{1,10}");
	}
	
	function regex(val,pattern) {
		var regexObj = eval("/"+pattern+"/gi");
		result = regexObj.exec(val);
		try {
			if(result[0]!=val) {
				return false;
			} else {
				return true;
			}
		} catch(e) {
			return false;
		}
	}
</script>