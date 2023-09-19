<%
		Response.Cookies("EZCalAdmin").Expires = date-1
		Session.Abandon
		
		Response.Redirect "admin_login.asp"
%>