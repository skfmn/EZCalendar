<!-- #include file="../includes/general_includes.asp"-->
<%
	strCookies = Request.Cookies("EZCalAdmin")("name")
	
	If strCookies = "" Then

		Response.Redirect "admin_login.asp"
  
	End If

	If Not blnARights Then

	    Response.Cookies("msg") = "nar"
	    Response.Redirect "admin.asp"

	End If

    msg = ""
    msg = Trim(Request.Cookies("msg"))

	If msg <> "" Then
		Call displayFancyMsg(getMessage(msg))
        Response.Cookies("msg") = ""
	End If

	If Trim(Request.QueryString("as")) = "y" Then

		lngMemberID = checkint(Trim(Request.Form("memberid")))
		strRights = Trim(Request.Form("rights"))
	
		Set Conn = Server.CreateObject("ADODB.Connection")
		Call ConnOpen(Conn)

		Set rsCommon = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT name FROM "&msdbprefix&"admin WHERE adminID = "&lngMemberID
	
		Call getTextRecordset(strSQL,rsCommon)
		If Not rsCommon.EOF Then
	        strTempName = rsCommon("name")
		End If
		Call closeRecordset(rsCommon)
	
		If InStr(strRights,"schedule") > 0 Then
		    Call getExecuteQuery("UPDATE "&msdbprefix&"admin SET schedule = 'True' WHERE adminID = "&lngMemberID)
		Else
		    Call getExecuteQuery("UPDATE "&msdbprefix&"admin SET schedule = 'False' WHERE adminID = "&lngMemberID)
		End If
	
		If InStr(strRights,"events") > 0 Then
		    Call getExecuteQuery("UPDATE "&msdbprefix&"admin SET events = 'True' WHERE adminID = "&lngMemberID)
		Else
		    Call getExecuteQuery("UPDATE "&msdbprefix&"admin SET events = 'False' WHERE adminID = "&lngMemberID)
		End If
	
		If InStr(strRights,"settings") > 0 Then
		    Call getExecuteQuery("UPDATE "&msdbprefix&"admin SET settings = 'True' WHERE adminID = "&lngMemberID)
		Else
		    Call getExecuteQuery("UPDATE "&msdbprefix&"admin SET settings = 'False' WHERE adminID = "&lngMemberID)
		End If
	
		If InStr(strRights,"admins_rights") > 0 Then
		    Call getExecuteQuery("UPDATE "&msdbprefix&"admin SET admins_rights = 'True' WHERE adminID = "&lngMemberID)
		Else
		    Call getExecuteQuery("UPDATE "&msdbprefix&"admin SET admins_rights = 'False' WHERE adminID = "&lngMemberID)
		End If

		If InStr(strRights,"arights") > 0 Then
		    Call getExecuteQuery("UPDATE "&msdbprefix&"admin SET arights = 'True' WHERE adminID = "&lngMemberID)
		Else
		    Call getExecuteQuery("UPDATE "&msdbprefix&"admin SET arights = 'False' WHERE adminID = "&lngMemberID)
		End If

		If InStr(strRights,"purge") > 0 Then
		    Call getExecuteQuery("UPDATE "&msdbprefix&"admin SET purge = 'True' WHERE adminID = "&lngMemberID)
		Else
		    Call getExecuteQuery("UPDATE "&msdbprefix&"admin SET purge = 'False' WHERE adminID = "&lngMemberID)
		End If

		Call ConnClose(Conn)

	    Response.Cookies("msg") = "car"
        Response.Redirect "arights.asp?id="&lngMemberID

    End If


	lngMemberID = checkint(Trim(Request.QueryString("id")))
	Set Conn = Server.CreateObject("ADODB.Connection")
	Call ConnOpen(Conn)

	Set rsCommon = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT name FROM "&msdbprefix&"admin WHERE adminID = "&lngMemberID 
	
	Call getTextRecordset(strSQL,rsCommon)
	If NOT rsCommon.EOF Then
	    strName = DBDecode(rsCommon("name"))
	End If
	Call closeRecordset(rsCommon)
%>
<!-- #include file="../includes/header.asp"-->
<div id="main" class="container">
  <header>
    <h2>Manage Rights for <%= strName %></h2>
  </header>
	<div class="row uniform">
		<div class="-4u 4u 12u(medium)">
<%
	Set rsCommon = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM "&msdbprefix&"admin WHERE adminID = "&lngMemberID 
	
	Call getTextRecordset(strSQL,rsCommon)
	If NOT rsCommon.EOF Then
%>
	    <form method="post" name="rights" id="rights" action="arights.asp?as=y" >
		  <input type="hidden" name="memberid" value="<%= lngMemberID %>" >
      <div class="row uniform">
        <div class="12u 12u$(small)">
<%
	intCount = 0
	For Each x in rsCommon.Fields
	    strChecked = ""
	    strRight = ""
	    
		If x.name <> "adminID" AND x.name <> "name" AND x.name <> "pwd" AND x.name <> "salt" Then	
			If x.name = "schedule" Then
				strRight = "Schedule"
				If rsCommon("schedule") = "True" Then strChecked = "checked"
			End If
			If x.name = "events" Then
				strRight = "Events"
				If rsCommon("events") = "True" Then strChecked = "checked"
			End If
			If x.name = "settings" Then
				strRight = "Settings"
				If rsCommon("settings") = "True" Then strChecked = "checked"
			End If
			If x.name = "admin_rights" Then
				strRight = "Admin"
				If rsCommon("admin_rights") = "True" Then strChecked = "checked"
			End If
			If x.name = "arights" Then
				strRight = "Rights"
				If rsCommon("arights") = "True" Then strChecked = "checked"
			End If
			If x.name = "purge" Then
				strRight = "Purge"
				If rsCommon("purge") = "True" Then strChecked = "checked"
			End If
%>				
				  <div class="12u 12u$(small)">
						<input type="checkbox" id="<%= strRight %>" name="rights" value="<%= x.name %>" <%= strChecked %> >
						<label for="<%= strRight %>"><%= strRight %></label>
				  </div>
<%
	    End If
    Next
%>
		      <div class="12u 12u$(small)">
            <input type="submit" name="submit" value="Submit" />
		      </div>
        </div>
      </div>
		  </form>
<%
	Else
%>
      <div class="table-wrapper">
	      <table>
	        <tr>
	          <td style="width:75%;text-align:left"><span>That person is not an Admin.</span></td>
		      </tr>
		    </table>
      </div>
  <%
	End If
    Call closeRecordset(rsCommon)
	Call ConnClose(Conn)
%>
    </div>
  </div>
</div>
<!-- #include file="../includes/footer.asp"-->