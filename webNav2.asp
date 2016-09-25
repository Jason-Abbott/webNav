<% Option Explicit %>
<% Response.Buffer = True %>
<html>
<head>
<!--#include file="include/webNav2_themes.inc"-->
<!--#include file="data/webNav2_data.inc"-->
<!--#include file="data/webNav2_cache.inc"-->
<!--#include file="include/show_status.inc"-->
<%
' Copyright 2000 Jason Abbott (jason@webott.com)
' Last updated 4/7/2000

dim strExpand		' lists expanded folders by index number
dim strHTML			' menu written to client
dim intRow			' track row
dim x				' loop counter
dim arMenu			' array of menu items, as follows
' 0 id
' 1 name
' 2 description
' 3 parent
' 4 url
' 5 target frame
' 6 hidden
' 7 depth in tree
' 8 number of children

arMenu = getMenu()

if Request.QueryString("strExpand") <> "" then
	strExpand = Request.QueryString("strExpand")
else
	strExpand = ""
end if

for x = 0 to UBound(arMenu,1)
	' display item only if it is not hidden
	if arMenu(x,6) = 0 then
	
		' track current item even after x is
		' possibly incremented to skip children
		intRow = x
		strHTML = strHTML & "<nobr>"

		' insert a blank space for every unit of tree depth		
		for d = 1 to (arMenu(x,7))
			strHTML = strHTML & "<img src='./images/blank_" _
				& strStyle & ".gif'>"
		next
		strHTML = strHTML & showItem(x,arMenu)

		if arMenu(x,8) > 0 then
			' display plus/minus graphic
			strHTML = strHTML & "<a name='" & x & "'>" _
				& "<a href='webNav2.asp?strExpand="
	 		if InStr(strExpand, "(" & x & ")") then
				' if the current item is selected for expansion then
				' display option to collapse
				strHTML = strHTML & Replace(strExpand, "(" & x & ")", "") _
					& "' " & showStatus("collapse " & arMenu(x,1)) _
					& "><img src=""./images/minus_"
			else
				' show option to expand and skip all children
				strHTML = strHTML & strExpand & "(" & x & ")" _
					& "#" & x & "' " & showStatus("strExpand " & arMenu(x,1)) _
					& "><img src=""./images/plus_"
				x = x + arMenu(x,8)	
			end if
			strHTML = strHTML & strStyle & ".gif"" border=0></a></a>"
		else
			strHTML = strHTML & "<img src=""./images/blank_" _
				& strStyle & ".gif"">"
		end if
		strHTML = strHTML & vbCrLf & item & vbCrLf & "<br>" & VbCrLf

		if x < UBound(arMenu,1) then
			' if items remain then adjust font size appropriately		
	 		if arMenu(thisx, 7) = 0 AND arMenu(x + 1, 7) > 0 then
				' if the current item is root (depth=0) but the next is a
				' child then set font size to 1
				strHTML = strHTML & "<font size=1>"
			elseif arMenu(thisx, 7) > 0 AND arMenu(x + 1, 7) = 0 then
				' if the current item is a child (depth>0) but the next
				' is at the root level then restore default font size
				strHTML = strHTML & "</font>"
			end if
		end if
	else
		' if the item should be hidden then also skip children
		x = x + arMenu(x,8)
	end if
next
%>
</head>
<body bgarColor="#<%=arColor(10)%>" text="#<%=arColor(7)%>" link="#<%=arColor(8)%>" vlink="#<%=arColor(8)%>" alink="#<%=arColor(9)%>">

<font face="Arial, Helvetica" size=2>
<%=strHTML%>
<p>
</font>
</table>
</body>
</html>