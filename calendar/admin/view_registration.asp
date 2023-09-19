<!--#include file="../includes/general_includes.asp"-->
<html>
<head>
  <link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
  <link type="text/css" rel="stylesheet" href="/calendar/assets/css/jquery.fancybox.css" />
  <link type="text/css" rel="stylesheet" href="/calendar/assets/css/main.css" />
</head>
<body>
<div id="main" class="container" style="max-width:600px;">
  <header style="text-align:center;"><h2>Registrants</h2></header>
<%
    lngSchedID = ""
	If Trim(Request.QueryString("eid")) <> "" Then 
	    lngSchedID = checkint(Trim(Request.QueryString("eid")))
	End If

	If Trim(Request.Form("eid")) <> "" then
	    lngSchedID = checkint(Trim(Request.Form("eid")))
	End If

	If Trim(Request.QueryString("cancel")) = "yes" Then

        intRegID = ""
	    If Trim(Request.QueryString("rid")) <> "" Then intRegID = checkint(Trim(Request.QueryString("rid")))

		Set Conn = Server.CreateObject("ADODB.Connection")
		Call ConnOpen(Conn)
		
		strSQL = "DELETE FROM "&msdbprefix&"registration WHERE regID = "&intRegID
		Call getExecuteQuery(strSQL)
		
		Call ConnClose(Conn)
		
		Response.Redirect "admin_view.asp"
		 
	End If

	If Trim(Request.Form("edit")) = "yes" Then

        intRegID = ""
	    If Trim(Request.Form("rid")) <> "" Then intRegID = checkint(Trim(Request.Form("rid")))

        strName = DBEncode(Request.Form("rname"))
        strAddInfo = DBEncode(Request.Form("addinfo"))

		Set Conn = Server.CreateObject("ADODB.Connection")
		Call ConnOpen(Conn)
		
		strSQL = "UPDATE "&msdbprefix&"registration SET reg_name = '"&strName&"', add_info = '"&strAddInfo&"'  WHERE regID = "&intRegID
		Call getExecuteQuery(strSQL)
		
		Call ConnClose(Conn)

	    Response.Cookies("eid") = lngSchedID
		Response.Redirect "admin_view.asp"

	End If

	If lngSchedID <> "" Then

		Set Conn = Server.CreateObject("ADODB.Connection")
	    Call ConnOpen(Conn)

	    Set rsCommon = Server.CreateObject("ADODB.Recordset")
	    strSQL = "SELECT allow_reg FROM "&msdbprefix&"calendar WHERE schedID = "&lngSchedID

	    Call getTextRecordset(strSQL,rsCommon)
		If Not rsCommon.EOF Then
	        blnAllowReg = rsCommon("allow_reg")
		End If
		Call closeRecordset(rsCommon)

	    If blnAllowReg = "True" Then 

			Set rsCommon = Server.CreateObject("ADODB.Recordset")
			strSQL = "SELECT "&msdbprefix&"registration.*, "&msdbprefix&"calendar.* FROM "&msdbprefix&"registration INNER JOIN "&msdbprefix&"calendar ON "&msdbprefix&"registration.schedID = "&msdbprefix&"calendar.schedID  WHERE "&msdbprefix&"calendar.schedID = "&lngSchedID
		
			Call getTextRecordset(strSQL,rsCommon)
			If Not rsCommon.EOF Then

				strEventTitle = DBDecode(rsCommon("event"))
			    datDate = rsCommon("schDate")
%>
<h4 style="text-align:center;"><%= DBDecode(strEventTitle) %><br /><%= datDate %></h4>	
<%
				Do While Not rsCommon.EOF

					intRegID = rsCommon("regID")
					strName = DBDecode(rsCommon("reg_name"))
					strAddInfo = DBDecode(rsCommon("add_info"))
	            
%>
  <form action="view_registration.asp" method="post" name="register" id="register" >
	  <input type="hidden" name="rid" value="<%= intRegID %>" />
	  <input type="hidden" name="edit" value="yes" />
	  <input type="hidden" name="eid" value="<%= lngSchedID %>" />
<div class="row">
    <div class="-3u 6u 12u(medium)" style="padding-bottom:30px;">
      <label for="rname" style="margin-bottom:-3px;">Name</label>
      <input type="text" id="rname" name="rname" value="<%= strName %>" size="20" />
    </div>
    <div class="-3u 6u 12u(medium)" style="padding-bottom:10px;">
      <label for="addinfo" style="margin-bottom:-3px;">Additional info:</label>
      <textarea id="addinfo" name="addinfo"><%= strAddInfo %></textarea>
    </div>
    <div class="-3u 3u 12u(medium)">
      <input type="submit" class="button fit" name="register" value="Edit" />
    </div>
    <div class="3u$ 12u(medium)">
      <a class="button fit"  name="register" href="view_registration.asp?cancel=yes&rid=<%= intRegID %>" >Cancel</a>
    </div>
  </div>
  </form>
<%
					rsCommon.MoveNext
				    If rsCommon.EOF Then 
	                    Exit Do
					Else
%><hr class="major" /><%
					End If
				Loop
		    Else
%>
  <div class="row">
    <div class="-3u 8u$ 12u(medium)" style="padding-bottom:30px;">There are no registrants for this event!</div>
  </div>
<%
			End If
			Call closeRecordset(rsCommon)
			
	    Else
%>
  <div class="row">
    <div class="-3u 8u$ 12u(medium)" style="padding-bottom:30px;">Registration is disabled for this event!</div>
  </div>
<%
	    End If
	    Call ConnClose(Conn)
	End If	
%>
</div>
</body>
</html>