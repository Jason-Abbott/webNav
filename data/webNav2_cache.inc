<%
' Copyright 2000 Jason Abbott (jason@webott.com)
' Last updated 4/7/2000

' creates this array
' 0 id
' 1 name
' 2 description
' 3 parent
' 4 url
' 5 target frame
' 6 hidden
' 7 depth in tree (replaces order field)
' 8 number of children (added)
' returns array-----------------------------------------------------------
Function getMenu()
	dim strQuery
	dim oRS				' recordset object
	dim intCount		' ZERO-BASED count of menu items
	dim x				' loop counter
	dim intDepth		' item's depth in tree
	dim arTemp			' temporary array
	dim arMenu			' array of menu items
	dim intRow			' tracks menu row globally
	dim strUsed
	
	' create menu array if it doesn't exist
	if Not IsArray(Application(unique & "Menu")) then
		Set oRS = Server.CreateObject("ADODB.RecordSet")
		strQuery = "SELECT * FROM menu_items ORDER BY item_order"
		oRS.Open strQuery, strDSN, adOpenStatic, adLockReadOnly, adCmdText
		intCount = CInt(oRS.RecordCount - 1)
		arTemp = oRS.GetRows
		oRS.Close
		Set oRS = nothing
		Redim arMenu(intCount,8)
		intDepth = 0
		intRow = -1
		strUsed = ""
		' we prime the subroutine by displaying all
		' root items
		Call orderItems(0,"")
		Application.Lock
		Application(unique & "Menu") = arMenu
		Application.Unlock
	else
		' if the Session array already exists then
		' assign it to the arMenu array
		arMenu = Application(unique & "Menu")
	end if
	getMenu = arMenu	
End Function

' this calls itself until the array is complete
' updates array-----------------------------------------------------------
Sub orderItems(intPid,intProw)
	' these need to be local so they can be used recursively
	
	dim i			' loop counter
	dim j			' loop counter
	dim intRowSub	' current menu row
	dim intOldRow	' track old row
	dim intProwSub	' parent item's row
	dim intPidSub	' parent's id
	
	intProwSub = intProw
	
	for i = 0 to intCount
		if intPidSub = arTemp(i,3) then
			' this item is a child
			intRowSub = intRow + 1
			intOldRow = intRow		

			' transfer the contents of the temporary
			' array to the new one
			for j = 0 to 6
				arMenu(intRow,j) = arTemp(i,j)
			next
			arMenu(intRow,7) = intDepth
			arMenu(intRow,8) = 0

			intDepth = intDepth + 1
			Call orderItems(arTemp(i,0),intRow)
			intDepth = intDepth - 1

			' each item must indicate not only the number of
			' immediate children but all progeny, so we add
			' one to the immediate parent row with each run
			' through and then we add the progeny of the current
			' row to the parent row, so it has a total
			if IsNumeric(intProwSub) then
				arMenu(intProwSub,8) = arMenu(intProwSub,8) + arMenu(intOldRow,8) + 1
			end if
		end if
	next
End Sub

%>