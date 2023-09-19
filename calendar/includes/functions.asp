<%
on error resume next
	strVersion = "4.0"

	ConnStr = "Provider=sqloledb;Data Source="&msdbserver&";Initial Catalog="&msdb&";User Id="&msdbid&";Password="&msdbpwd
  
	Set Conn = Server.CreateObject("ADODB.Connection")
	Call ConnOpen(Conn)

	Set rsCommon = Server.CreateObject("ADODB.Recordset")
		
	Call getTableRecordset(""&msdbprefix&"settings",rsCommon)
	If Not rsCommon.Eof Then
	    strSiteTitle = DBDecode(rsCommon("site_title"))
	    strDomainname = DBDecode(rsCommon("domain_name"))
	    blnLetUsers = rsCommon("letusers")
	    intDelDays = rsCommon("delete_days")
	    strAnnouncements = DBDecode(rsCommon("announcements"))
	End If
	Call closeRecordset(rsCommon)
	Call ConnClose(Conn)

	If Request.Cookies("EZCalAdmin")("adminID") <> "" Then
	    Call getMyInfo(Request.Cookies("EZCalAdmin")("adminID"))
	End If

	Sub getMyInfo(lMemberID)

		blnSchedule = false
		blnEvents = false
		blnSettings = false
		blnAdminRights = false
		blnARights = false
	    blnPurge = false

		If Session("loggedin") = "" Then

			Session("loggedin") = "yes"

			Set Conn = Server.CreateObject("ADODB.Connection")
			Call ConnOpen(Conn)

			Set rsCommon = Server.CreateObject("ADODB.Recordset")
			strSQL = "SELECT * FROM "&msdbprefix&"admin WHERE adminID = "&lMemberID

			Call getTextRecordset(strSQL,rsCommon)
			If Not rsCommon.EOF Then
				blnSchedule = rsCommon("schedule")
				blnEvents = rsCommon("events")
				blnSettings = rsCommon("settings")
				blnAdminRights = rsCommon("admin_rights")
				blnARights = rsCommon("arights")
	            blnPurge = rsCommon("purge")
			End If
			Call closeRecordset(rsCommon)
			Call ConnClose(Conn)

			Session("blnSchedule") = blnSchedule
			Session("blnEvents") = blnEvents
			Session("blnSettings") = blnSettings
			Session("blnAdminRights") = blnAdminRights
			Session("blnARights") = blnARights
            Session("blnPurge") = blnPurge

		Else

			blnSchedule = Session("blnSchedule")
			blnEvents = Session("blnEvents")
			blnSettings = Session("blnSettings")
			blnAdminRights = Session("blnAdminRights")
			blnARights = Session("blnARights")
            blnPurge = Session("blnPurge")

		End If

	End Sub

	Function getResponse(sURL)
		Dim strTemp
		strTemp = ""

		Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		xmlhttp.SetOption(2) = (xmlhttp.GetOption(2) - SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS)
		xmlhttp.Open "GET", sURL, false
		xmlhttp.Send(NULL)
		If xmlhttp.readyState = 4 Then
			If xmlhttp.status <> 200 Then
			    strTemp = "<span style=""color:#FF0000"">Error "&xmlhttp.status&" - "&xmlhttp.statusText&"</span><br />"
			Else
			    strTemp = xmlhttp.ResponseText
			End If
		End If
		Set xmlhttp = Nothing

		getResponse = strTemp

	End Function

	Function schedCheck(datCurr)
	    Dim strTemp: strTemp = ""
		schedCheck = False

		Set Conn = Server.CreateObject("ADODB.Connection")
	    Call ConnOpen(Conn)
		
		Set rsTemp = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM "&msdbprefix&"calendar WHERE schDate = '"&datCurr&"'"
		Call getTextRecordset(strSQL,rsTemp)
		If not rsTemp.EOF Then
		    schedCheck = True
		End If
		Call closeRecordset(rsTemp)
	    Call ConnClose(Conn)
	
	End Function
	
	Function GetDaysInMonth(iMonth, iYear) 
	    dTemp = DateAdd("d", -1, DateSerial(iYear, iMonth + 1, 1))
	    GetDaysInMonth = Day(dTemp)
	End Function

	Function GetWeekdayMonthStartsOn(dAnyDayInTheMonth)
	    dTemp = DateAdd("d", -(Day(dAnyDayInTheMonth) - 1), dAnyDayInTheMonth)
	    GetWeekdayMonthStartsOn = WeekDay(dTemp)
	End Function

	Function SubtractOneMonth(dDate)
	    SubtractOneMonth = DateAdd("m", -1, dDate)
	End Function

	Function AddOneMonth(dDate)
	    AddOneMonth = DateAdd("m", 1, dDate)
	End Function
	
	
	Sub selectAllEvents(iEvID)
	    Dim rsTemp

	    If iEvID = "" Then iEvID = 0

		Set Conn = Server.CreateObject("ADODB.Connection")
		Call ConnOpen(Conn)

		Set rsTemp = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM "&msdbprefix&"calendar ORDER BY schDate asc"
		
		Call getTextRecordset(strSQL,rsTemp)
		If Not rsTemp.EOF Then
			Response.Write "<select id=""schedid"" name=""schedid"" required>"&vbcrlf
			Response.Write "                        <option value="""">Select an event</option>"&vbcrlf
			Do While Not rsTemp.EOF
				strEvent = DBDecode(rsTemp("event"))
				If Len(strEvent) > 15 Then strEvent = left(strEvent,15)&"..."

	            If Cint(rsTemp("schedID")) = Cint(iEvID) Then
	                Response.Write "                        <option value="""&rsTemp("schedID")&""" title="""&DBDecode(rsTemp("event"))&""" selected>"&rsTemp("schDate")&" - "&strEvent&"</option>"&vbcrlf
	            Else	
				    Response.Write "                        <option value="""&rsTemp("schedID")&""" title="""&DBDecode(rsTemp("event"))&""">"&rsTemp("schDate")&" - "&strEvent&"</option>"&vbcrlf
	           End If
				rsTemp.MoveNext
				If rsTemp.EOF Then Exit Do
			Loop
			Response.Write "                    </select>"&vbcrlf
		Else
		    Response.Write "<select id=""schedid"" name=""schedid""><option value=""0"">No Current Events</option></select>"&vbcrlf
		End If
		Call closeRecordset(rsTemp)
		Call ConnClose(Conn)
		
	End Sub
	
	Function getMessage(sMsg)
		Dim rsTemp, strTemp
		strTemp = ""
		
		Set Conn = Server.CreateObject("ADODB.Connection")
	    Call ConnOpen(Conn)

		Set rsTemp= Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT message FROM "&msdbprefix&"messages WHERE msg = '"&Trim(sMsg)&"'"
		
		Call getTextRecordset(strSQL,rsTemp)
		If Not rsTemp.EOF Then
		    strTemp = DBDecode(rsTemp("message"))
		Else
		    strTemp = sMsg
		End If
		Call closeRecordset(rsTemp)
		Call ConnClose(Conn)
		
		getMessage = strTemp
		 
	End Function

	Function msgTrans(sMsg)

		Dim strTemp: strTemp = ""

		Select  Case sMsg
			case "logch"
				strTemp = "Login info updated:"
			case "evsch"
				strTemp = "Event scheduled:"
			case "evmod"
				strTemp = "Event info changed:"
			case "setch"
				strTemp = "Settings updated:"
			case "evdel"
				strTemp = "Event deleted:"
			case "regs"
				strTemp = "Registration successfull:"
			case "notad"
				strTemp = "No Date:"
			case "ps"
			    strTemp = "Purge successfull:"
			case "ant"
				strTemp = "Admin name taken:"
			case "das"
				strTemp = "Deleted an Admin:"
			case "adad"
				strTemp = "Added an Admin:"
			case "nadmin"
				strTemp = "Change Main Admin:"
			case "cpwds"
				strTemp = "Changed password:"
			case "siu"
				strTemp = "Site Info:"
			case "nar"
			    strTemp = "No Admin rights:"
			case "car"
			    strTemp = "Change Admin rights:"
			case "mus"
			    strTemp = "Message updated:"
	        case else
	            strTemp = "TBD:"
		End Select

		msgTrans = strTemp

	End Function

	Sub displayFancyMsg(sText)
%>
<div style="display:none">
	<a id="textmsg" href="#displaymsg">Message</a>
	<div id="displaymsg" style="text-align:left;width:500px;">
		<div class="left_menu_block">
			<div class="left_menu_top"><h2>Message</h2></div>
			<div class="left_menu_center" align="center" style="padding-left:0px;"><span><%= sText %></span></div>
			<div class="left_menu_bottom"></div>
		</div>
	</div>
</div>
<%
	End Sub

	Sub getRegistrants(lSchedID)
	    Dim rsTemp
%>
	  <div style="display:none">
      <a id="textnrmsg" href="#displaynrmsg">Message</a>
      <div id="displaynrmsg" style="background-color:#000000;color:#FFFFFF;text-align:center;width:400px;">
        <h4><a class="first" target="_blank" href="printpage.asp?schedid=<%= lSchedID %>">Printable page</a></h4>
        <div style="float:left;position:relative;display:block;width:95%;padding-bottom:5px;">
          <div style="float:left;position:relative;display:inline;width:50%"><strong>Name</strong></div>
          <div style="float:right;position:relative;display:inline;width:50%;"><strong>Info</strong></div>
        </div>
        <hr style="font-weight:bold;width:99%" />
<%
	Set Conn = Server.CreateObject("ADODB.Connection")
	Call ConnOpen(Conn)

	Set rsTemp = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM registration WHERE schedID = "&lSchedID&" ORDER BY reg_name asc"
	
	Call getTextRecordset(strSQL,rsTemp)
	If Not rsTemp.EOF Then
		Do While Not rsTemp.EOF
			Response.Write "        <div style=""float:left;position:relative;display:block;width:95%;padding-bottom:5px;"">"&vbcrlf
			Response.write "          <div style=""float:left;position:relative;display:inline;width:50%;text-align:left;"">"&DBDecode(rsTemp("reg_name"))&"</div>"&vbcrlf
			Response.write "          <div style=""float:right;position:relative;display:inline;width:50%;text-align:left;""><pre class=""pre"">"&DBDecode(rsTemp("add_info"))&"</pre></div>"&vbcrlf
			Response.Write "        </div>"&vbcrlf
			Response.Write "        <hr style=""font-weight:bold;width:99%"" />"&vbcrlf
			rsTemp.MoveNext
			If rsTemp.EOF Then Exit Do
		Loop
	Else
	    Response.Write "        <div style=""float:left;position:relative;display:block;width:99%;padding-bottom:5px;"">Oops! You forgot to select something!</div>"&vbcrlf
	    Response.Write "        <hr style=""font-weight:bold;width:99%"" />"&vbcrlf
	End If
	Call closeRecordset(rsTemp)
	Call ConnClose(Conn)  
%> 
        <div class="clear"></div>           
      </div>
		</div>
<%
	End Sub

	Function DBEncode(DBvalue)
		Dim fieldvalue: fieldvalue = Trim(DBvalue)
		
		If fieldvalue <> "" AND Not IsNull(fieldvalue) Then
		
			Set encodeRegExp = New RegExp 
			encodeRegExp.Pattern = "((delete)*(select)*(update)*(into)*(drop)*(insert)*(declare)*(xp_)*(union)*)"
			encodeRegExp.IgnoreCase = True
			encodeRegExp.Global = True
			Set Matches = encodeRegExp.Execute(fieldvalue)
			For Each Match In Matches
			    fieldvalue = Replace(fieldvalue,Match.Value,StrReverse(Match.Value))
			Next
			fieldvalue=replace(fieldvalue,"'","''")

		End If
		
		DBEncode = fieldvalue
		
	End Function

	Function DBDecode(DBvalue)
		Dim fieldvalue: fieldvalue = Trim(DBvalue)
		
		If fieldvalue <> "" AND ( NOT IsNull(fieldvalue) ) Then
		
			Set encodeRegExp = New RegExp 
			encodeRegExp.Pattern = "((eteled)*(tceles)*(etadpu)*(otni)*(pord)*(tresni)*(eralced)*(_px)*(noinu)*)"
			encodeRegExp.IgnoreCase = True
			encodeRegExp.Global = True
			Set Matches = encodeRegExp.Execute(fieldvalue)
			For Each Match In Matches
			    fieldvalue = Replace(fieldvalue,Match.Value,StrReverse(Match.Value))
			Next
			fieldvalue = replace(fieldvalue,"''","'")

		End If
		
		DBDecode = fieldvalue
		
	End Function
	
	Function checkInt(iVal)
		Dim intTemp: intTemp = 0
		
		If iVal <> "" Then
			If Not IsNumeric(iVal) Then
			    Call displayFancyMsg("Input was not a number!")
			Else
			    intTemp = iVal		
			End If
		Else
			Call displayFancyMsg(txtInputWasEmpty&"Input was empty!")
		End If
		
		checkInt = intTemp
		
	End Function

	Sub trace(strDebugString)
	    Response.Write "Debug: "&strDebugString&"<br>"
	End Sub

	Sub catch(sText,sText2)

		If Err.Number <> 0 then
		    trace(sText&" - "&err.description)
		Else
		    trace(sText&" - no error")
		End If

		If sText2 <> "" Then
		    trace(sText&" - "&sText2)
		End If
				
		on error goto 0

	End Sub
%>
