<%
' Copyright 2000 Jason Abbott (webcal@webott.com)
' Last updated 4/7/2000

'    0 id
'    1 name
'    2 description
'    3 parent
'    4 url
'    5 target frame
'    6 hidden
'    7 depth in tree
'    8 number of children

' generate HTML for menu strHTML
' returns string----------------------------------------------------------
Function showItem(ByRef x, ByRef arMenu)
	dim strHTML		' HTML returned

	if arMenu(x,4) <> "" then
		strHTML = "<a href='" & arMenu(x,4) _
			& "' target='"
		if arMenu(x,5) <> "" then
			strHTML = strHTML & arMenu(x,5)
		else
			' if no frame target was specified then
			' make _top the target
			strHTML = strHTML & "_top"
		end if
		strHTML = strHTML & "' "

		' if a description was entered for the
		' strHTML then generate the javascript to
		' display it in the status bar
		if arMenu(x,2) <> "" then
			strHTML = strHTML & showStatus(arMenu(x,2))
		else
			strHTML = strHTML & showStatus(arMenu(x,1))
		end if
		strHTML = strHTML & ">" & arMenu(x,1) & "</a>"
	else
		strHTML = arMenu(x,1)
	end if
	
	showItem = strHTML & "</nobr>"
End Function
%>