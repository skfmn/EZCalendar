<!--#include file="../includes/general_includes.asp"-->
<%
on error resume next
	strCookies = Request.Cookies("EZCalAdmin")("name")

	If strCookies = "" Then
		Response.Redirect "admin_login.asp"
	End If

	If Not blnEvents Then
	    Response.Cookies("msg") = "nar"
	    Response.Redirect "admin.asp"
	End If

    msg = ""
    msg = Trim(Request.Cookies("msg"))

	If msg <> "" Then
		Call displayFancyMsg(getMessage(msg))
        Response.Cookies("msg") = ""
	End If

	lngSchedID = ""
	If Trim(Request.Form("schedid")) <> "" Then  lngSchedID = checkint(Trim(Request.Form("schedid")))

    If Trim(Request.QueryString("eid")) <> "" Then  lngSchedID = checkint(Trim(Request.QueryString("eid")))

    If Trim(Request.Cookies("eid")) <> "" Then
	    lngSchedID = checkint(Trim(Request.Cookies("eid")))
	    Response.Cookies("eid") = ""
	End If

    If Trim(Request.Form("edit")) <> "" Then

        lngSchedID = Request.Form("edit")
        strEventTitle = DBEncode(Request.Form("eventname"))
        strEventText = DBEncode(Request.Form("comments"))

        If Request.Form("registration") = "yes" Then
            blnAllowReg = True
        Else
            blnAllowReg = False
        End if

		Set Conn = Server.CreateObject("ADODB.Connection")
		Call ConnOpen(Conn)
		strSQL = "UPDATE "&msdbprefix&"calendar SET allow_reg = '"&blnAllowReg&"', event = '"&strEventTitle&"', text = '"&strEventText&"' WHERE schedID = "&lngSchedID
        Call getExecuteQuery(strSQL)
        Call ConnClose(Conn)

        Response.Cookies("msg") = "evmod"
        Response.Cookies("eid") = lngSchedID
        Response.Redirect "admin_view.asp"

    End If

    If Trim(Request.QueryString("delete")) = "yes" Then

		Set Conn = Server.CreateObject("ADODB.Connection")
		Call ConnOpen(Conn)

		strSQL = "DELETE FROM "&msdbprefix&"calendar WHERE schedID = "&lngSchedID
        Call getExecuteQuery(strSQL)

		strSQL = "DELETE FROM "&msdbprefix&"registration WHERE schedID = "&lngSchedID
        Call getExecuteQuery(strSQL)

        Call ConnClose(Conn)

        Response.Cookies("msg") = "evdel"
        Response.Cookies("eid") = lngSchedID
        Response.Redirect "admin_view.asp"

    End If

    If Request.QueryString("p") = "s" Then

	    If Not blnARights Then

            Response.Cookies("msg") = "nar"
            Response.Redirect "admin_view.asp"

        Else

            datDelDate = DateAdd("d",-Cint(intDelDays),Date)

            strSQl = "DELETE FROM "&msdbprefix&"calendar WHERE schDate < '"&datDelDate&"'"
		    Set Conn = Server.CreateObject("ADODB.Connection")
		    Call ConnOpen(Conn)
            Call getExecuteQuery(strSQL)
            Call ConnClose(Conn)

            Response.Cookies("msg") = "ps"
            Response.Redirect "admin_view.asp"

	    End If

    End If

%>
<!-- #include file="../includes/header.asp"-->
    <div id="main" class="container">
        <div class="row">
            <div class="-3u 6u$">
                <header>
                    <h2>Manage Events</h2>
                </header>
            </div>
        </div>
        <form action="admin_view.asp" id="selectevent" name="selectevent" method="post">
            <div class="row">
                <div class="-3u 6u$ 12u(medium)">
                    <div class="select-wrapper">
                        <%Call selectAllEvents(lngSchedID)%>
                    </div>
                </div>
                <div class="-3u 6u$ 12u(medium)" style="padding-top: 10px;">
                    <input type="submit" class="button fit" value="Select an Event" />
                </div>
            </div>
        </form>
        <%
	    If lngSchedID <> "" Then

		    Set Conn = Server.CreateObject("ADODB.Connection")
		    Call ConnOpen(Conn)

		    Set rsCommon = Server.CreateObject("ADODB.Recordset")
		    strSQL = "SELECT * FROM "&msdbprefix&"calendar WHERE schedID = "&lngSchedID
		    Call getTextRecordset(strSQL, rsCommon)
		    If Not rsCommon.EOF Then
			    datDate = rsCommon("schDate")
			    strEventTitle = DBDecode(rsCommon("event"))
			    blnAllowReg = rsCommon("allow_reg")
			    strEventText = DBDecode(rsCommon("text"))
		    End If
		    Call closeRecordset(rsCommon)
		    Call ConnClose(Conn)

            If Cint(intDelDays) <> 0 Then
                strPurgeText = "Purge Events older than "&intDelDays&" days"
            Else
                strPurgeText = "Purge all past events"
            End If
        %>
        <header style="text-align: center;">
            <h4><%= datDate %></h4>
        </header>
        <form action="admin_view.asp" method="post" id="event" name="event">
            <input type="hidden" name="edit" value="<%= lngSchedID %>" />
            <div class="row">
                <div class="-3u 6u 12u(medium)" style="padding-bottom: 30px;">
                    <label for="eventname" style="margin-bottom: -2px;">Event Name:</label>
                    <input type="text" id="eventname" name="eventname" value="<%= strEventTitle %>">
                </div>
                <div class="-3u 6u 12u(medium)" style="padding-bottom: 10px;">
                    <h5>Allow Registration:&nbsp;&nbsp;<a class="picimg fancybox.ajax" href="view_registration.asp?eid=<%= lngSchedID %>">View Registrations</a></h5>
                    <% If blnAllowReg Then %>
                    <input type="radio" id="regyes" name="registration" value="yes" checked>
                    <label for="regyes">Yes</label>
                    <input type="radio" id="regno" name="registration" value="no">
                    <label for="regno">No</label>
                    <% Else %>
                    <input type="radio" id="regyes" name="registration" value="yes">
                    <label for="regyes">Yes</label>
                    <input type="radio" id="regno" name="registration" value="no" checked>
                    <label for="regno">No</label>
                    <% End If %>
                </div>
                <div class="-3u 6u 12u$medium)" style="padding-bottom: 10px;">
                    <label for="comments" style="margin-bottom: -2px;">Comments:</label>
                    <textarea id="comments" name="comments" cols="50" rows="10"><%= strEventText %></textarea>
                </div>
                <div class="-3u 3u 12u$(medium)" style="text-align: center">
                    <input class="button fit" type="button" onclick="return confirmSubmit('Are you SURE you want to delete this event?','admin_view.asp?delete=yes&eid=<%= lngSchedID %>');" value="Delete" />
                </div>
                <div class="3u$ 12u$(medium)" style="text-align: center">
                    <input class="button fit" type="submit" value="Edit" />
                </div>
            </div>
        </form>
        <div class="row">
            <div class="-3u 6u$ 12u$(medium)">
                <input type="button" onclick="return confirmSubmit('WARNING!!\n Are you sure you want to <%= LCase(strPurgeText) %>?\n This cannot be undone!','admin_view.asp?p=s')" class="button fit" value="<%= strPurgeText %> " />
            </div>
        </div>
        <% End If %>
    </div>
    <div style="display: none;">
        <form action="view_registration.asp" method="post" id="register">
            <header style="text-align: center;">
                <h2>Registrants</h2>
            </header>
        <%
	If lngSchedID <> "" Then

		Set Conn = Server.CreateObject("ADODB.Connection")
	    Call ConnOpen(Conn)

		Set rsCommon = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT "&msdbprefix&"registration.*, "&msdbprefix&"calendar.* FROM "&msdbprefix&"registration INNER JOIN "&msdbprefix&"calendar ON "&msdbprefix&"registration.schedID = "&msdbprefix&"calendar.schedID  WHERE "&msdbprefix&"calendar.schedID = "&Cint(lngSchedID)

		Call getTextRecordset(strSQL,rsCommon)
		If Not rsCommon.EOF Then
			Do While Not rsCommon.EOF

				intRegID = rsCommon("regID")
				strName = DBDecode(rsCommon("reg_name"))
				strAddInfo = DBDecode(rsCommon("add_info"))
				strEventName = DBDecode(rsCommon("event"))
				datDate = rsCommon("schDate")
        %>
            <h4 style="text-align: center;"><%= DBDecode(strEventName) %>&nbsp;&nbsp;<%= datDate %></h4>
            <input type="hidden" name="rid" value="<%= intRegID %>" />
            <input type="hidden" name="edit" value="yes" />
            <input type="hidden" name="eid" value="<%= lngSchedID %>" />
            <div class="row">
                <div class="-3u 6u 12u(medium)" style="padding-bottom: 30px;">
                    <label for="rname" style="margin-bottom: -3px;">Name</label>
                    <input type="text" id="rname" name="rname" value="<%= strName %>" size="20" />
                </div>
                <div class="-3u 6u 12u(medium)" style="padding-bottom: 10px;">
                    <label for="addinfo" style="margin-bottom: -3px;">Additional info:</label>
                    <textarea id="addinfo" name="addinfo"><%= strAddInfo %></textarea>
                </div>
                <div class="-3u 3u 12u(medium)">
                    <input type="submit" class="button fit" name="submit" value="Edit" />
                </div>
                <div class="3u$ 12u(medium)">
                    <a class="button fit" name="register" href="view_registration.asp?cancel=yes&rid=<%= intRegID %>">Cancel</a>
                </div>
            </div>
        <%
				rsCommon.MoveNext
				If rsCommon.EOF Then
	               Exit Do
	            Else
        %><hr class="major" />
        <%
	            End If
			Loop
		Else
        %>
        <div style="padding-bottom: 30px; width: 200px;">No Registrants</div>
        <%
		End If
		Call closeRecordset(rsCommon)
		Call ConnClose(Conn)
	End If
        %>
        </form>
    </div>
<!-- #include file="../includes/footer.asp"-->