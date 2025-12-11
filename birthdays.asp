
<TR>
<TD valign=top bgcolor="<% =strForumCellColor %>" colspan="<%=sGetColspan(7, 6)%>" >
<%     
today = date()
currday = day(date())  
currmonth = month(date())  
curryear = year(date()) 


strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID," & strMemberTablePrefix & "MEMBERS.M_NAME," & strMemberTablePrefix & "MEMBERS.M_HIDEAGE," & strMemberTablePrefix & "MEMBERS.M_BIRTHDATE"
strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_BIRTHDATE LIKE '____" & currmonth & currday &"'" 
strSql = strSql & "ORDER BY (" & strMemberTablePrefix & "MEMBERS.MEMBER_ID) "

set rsBirthday = my_Conn.Execute (StrSql)

If rsBIRTHDAY.Eof OR rsBIRTHDAY.Bof Then
	response.write "<font face="& strForumCellColor &" size=" & strFooterFontSize & ">No Birthdays Today</font>" 
Else
	btoday = false
	do until rsBIRTHDAY.eof 
	strBirthdate = rsBirthday("M_BIRTHDATE")  & "000001"
	strBirthdate = strToDate (strBirthdate)
	Age = GetAge(strBirthdate)
	If rsBirthday("M_HIDEAGE") = 0 then

		if btoday = false then
			Response.Write "<font face="& strForumCellColor &" size=" & strFooterFontSize & ">Todays Birthdays: </font>"
			btoday = true
		else
			'do nothing
		end if
		if strUseExtendedProfile then
			response.write "<a href='pop_profile.asp?mode=display&id=" & rsBIRTHDAY("Member_ID") & "'><font face="& strForumCellColor &" size=" & strFooterFontSize & ">" & rsBIRTHDAY("M_NAME") & "</a>" & "(" & age & "), </font>"
		else		
			response.write "<a href=""JavaScript:openWindow2('pop_profile.asp?mode=display&id=" & rsBIRTHDAY("Member_ID") & "')""><font face="& strForumCellColor &" size=" & strFooterFontSize & ">" & rsBIRTHDAY("M_NAME") & "</a>" & "(" & age & "), </font>"
		end if
			 
	else
	'do nothing
	End if

	rsBIRTHDAY.movenext
	loop
End if

rsBIRTHDAY.close
set rsBIRTHDAY = Nothing
%> 
</TD>
</TR>
