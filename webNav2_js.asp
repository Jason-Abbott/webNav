<% Option Explicit %>
<% Response.Buffer = True %>

<!-- 
Copyright 1999 Jason Abbott (jason@webott.com)
Last updated 07/30/1999
-->

<!--#include file="webNav2_themes.inc"-->
<SCRIPT LANGUAGE="JavaScript">
<!--

// go through each division and hide children

function init() {
	var allDIVs = document.all.tags("DIV");
	for (i=0; i<allDIVs.length; i++) {
		if (allDIVs(i).className == "child") {
			allDIVs(i).strStyle.display = "none";
		}
	}
}

// cache graphics for fast display

plus = new Image(); plus.src = "./images/plus_<%=strStyle%>.gif";
minus = new Image(); minus.src = "./images/minus_<%=strStyle%>.gif";
blank = new Image(); blank.src = "./images/blank_<%=strStyle%>.gif";

// change strStyle visibility and image

function expandIt(el) {
	whichEl = eval(el + "_child");
	if (whichEl.strStyle.display == "none") {
		whichEl.strStyle.display = "block";
		eval("document." + el + "_img.src=minus.src");
	} else {
		whichEl.strStyle.display = "none";
		eval("document." + el + "_img.src=plus.src");
	}
   window.event.cancelBubble = true;
}

onload = init;
//-->
</SCRIPT>
</head>
<body bgarColor="#<%=arColor(10)%>" text="#<%=arColor(7)%>" link="#<%=arColor(8)%>" vlink="#<%=arColor(8)%>" alink="#<%=arColor(9)%>">
<font face="Arial, Helvetica" size=2>
<div class="root">

<!--#include file="data/webNav2_array.inc"-->
<!--#include file="show_status.inc"-->
<%
' the above include provides the array menuItem
' with these contents:
' 0 id
' 1 name
' 2 description
' 3 parent
' 4 url
' 5 target frame
' 6 hidden
' 7 depth in tree
' 8 number of children

dim item, divList, z

for x = 0 to UBound(menuItem,1)

' display item only if it is not hidden

	if menuItem(x,6) = 0 then
		response.write "<nobr>"
		for d = 1 to menuItem(x,7)
			response.write "<img src='./images/blank_" _
				& strStyle & ".gif'>"
		next
%>
<!--#include file="webNav2_item.inc"-->
<%
' JavaScript doesn't accept variable names that begin
' with a number so prepend "webNav"

		if menuItem(x,8) > 0 then
			response.write "<a name='" & x & "'><a href='#" & x & "' onClick=" _
				& """expandIt('webNav" & menuItem(x,0) & "'); return true;"">" _
				& "<img name='webNav" & menuItem(x,0) & "_img' src='./images/plus_" _
				& strStyle & ".gif' border=0></a></a>" & VbCrLf _
				& item & VbCrLf & "<br>" & VbCrLf & "<div id='webNav" _
				& menuItem(x,0) & "_child' class='child'>" & VbCrLf
				
' figure out where this new division should end
' add the number of children to the current row
				
			divList = divList & "(" & x + menuItem(x,8) & ") "
		else
			response.write "<img src='./images/blank_" _
				& strStyle & ".gif'>" & VbCrLf & item & "<br>" & VbCrLf
		end if
	else
	
' if the item should be hidden then also skip children

		x = x + menuItem(x,8)
	end if
	
' write closing division tag for all ocurrences of the
' present row in divList
	
	for each z in Filter(Split(divList), "(" & x & ")")
		response.write "</div>" & VbCrLf & VbCrLf
	next
next
%>
</div>
<p>
<font face="Tahoma, Arial, Helvetica" size=2>
<b>
<img src="./images/blank_<%=strStyle%>.gif">
<a href="/jason/webNav.html" target="_top"
<%=ShowStatus("Return to webNav home")%>>Home</a><br>
<img src="./images/blank_<%=strStyle%>.gif">
<a href="webNav2_admin.asp" target="body"
<%=ShowStatus("Edit the menu")%>>Admin</a><br>
<img src="./images/blank_<%=strStyle%>.gif">
<a href="webNav2.asp"
<%=ShowStatus("Switch to non-JavaScript version")%>>Legacy</a>
</b>
</font>
</body>
</html>