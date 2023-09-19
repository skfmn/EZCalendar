<!--#include file="../includes/general_includes.asp"-->
<%
on error resume next
	strCookies = Request.Cookies("EZCalAdmin")("name")

	If strCookies = "" Then
		Response.Redirect "admin_login.asp"
	End If

	If Not blnSettings Then
	    Response.Cookies("msg") = "nar"
	    Response.Redirect "admin.asp"
	End If

    msg = ""
    msg = Trim(Request.Cookies("msg"))

	If msg <> "" Then
		Call displayFancyMsg(getMessage(msg))
        Response.Cookies("msg") = ""
	End If

   If Request.Form("chmstgs") <> "" Then

        strSiteTitle = DBEncode(Request.Form("sitename"))
        strDomainname = DBEncode(Request.Form("domainname"))

        Set Conn = Server.CreateObject("ADODB.Connection")
        Call ConnOpen(Conn)
        strSQL = "UPDATE " & msdbprefix & "settings SET site_title = '"&strSiteTitle&"', domain_name = '"&strDomainname&"'"
        Call getExecutequery(strSQL)
        Call ConnClose(Conn)

        Response.Cookies("msg") = "siu"
        Response.Redirect "admin_settings.asp"

    End If

    If Request.Form("chmsg") <> "" Then

        Set Conn = Server.CreateObject("ADODB.Connection")
        Call ConnOpen(Conn)

        For Each i in Request.Form
	        If left(i,8) = "messages" Then

		        strFormValue = Replace(i,left(i,9),"")
		        strFormValue = Replace(strFormValue,right(i,1),"")
                strFormMsg = Request.Form(i)

                strSQL = "UPDATE " & msdbprefix & "messages SET message = '"&DBEncode(strFormMsg)&"' WHERE msg = '"&strFormValue&"'"
                Call getExecuteQuery(strSQL)

	        End If
        Next

        Call ConnClose(Conn)

        Response.Cookies("msg") = "mus"
        Response.Redirect "admin_settings.asp"

    End If

    If Request.Form("change") = "yes" Then

        blnLetusers = Request.Form("letusers")
        intDays = checkint(Request.Form("ndays"))
        strAnnounce = DBEncode(Request.Form("announce"))

        Set Conn = Server.CreateObject("ADODB.Connection")
        Call ConnOpen(Conn)

        strSQL = "UPDATE "&msdbprefix&"settings SET letusers = "&blnLetusers&", delete_days = "&intDays&", announcements = '"&strAnnounce&"'"
        Call getExecuteQuery(strSQL)
        Call ConnClose(Conn)

        Response.Cookies("msg") = "setch"
        Response.Redirect "admin_settings.asp"

    End If
%>
<!-- #include file="../includes/header.asp"-->
<div id="main" class="container">
    <div class="row">
        <div class="12u$">
            <header>
                <h2>Manage Settings</h2>
            </header>
        </div>
    </div>
    <div class="row">
        <div class="6u 12u$(medium)">
            <h3>Messages</h3>
<%
    Set Conn = Server.CreateObject("ADODB.Connection")
    Call ConnOpen(Conn)

    Set rsCommon = Server.CreateObject("ADODB.Recordset")
    strSQL = "SELECT * FROM "&msdbprefix&"messages"

    Call getTextRecordset(strSQL, rsCommon)
    If Not rsCommon.EOF Then
%>
            <div class="table-wrapper">
                <form action="admin_settings.asp" method="post">
                    <input type="hidden" name="chmsg" value="y" />
                    <table>
                        <tbody>
<%
        Do While Not rsCommon.EOF
            strTempMsg = ""
            strTempMsg = rsCommon("msg")
%>
                            <tr>
                                <td style="width: 30%;">
                                    <%= msgTrans(strTempMsg) %>
                                </td>
                                <td style="width: 70%;">
                                    <input type="text" name="messages[<%=strTempMsg %>]" value="<%= DBDecode(rsCommon("message")) %>" />
                                </td>
                            </tr>
<%
            rsCommon.MoveNext
            If rsCommon.EOF Then Exit Do
        Loop
%>
                            <tfoot>
                                <tr>
                                    <td colspan="2">
                                        <input type="submit" value="Save Admin Messages" class="button fit" />
                                    </td>
                                </tr>
                            </tfoot>
                        </tbody>
                    </table>
                </form>
            </div>
<%
    End If
    Call closeRecordset(rsCommon)
    Call ConnClose(Conn)
%>
        </div>
        <div class="6u 12u(medium)">

            <h3>Site Settings</h3>
            <form action="admin_settings.asp" method="post">
                <input type="hidden" name="chmstgs" value="y" />
                <div class="row">
                    <div class="4u 12u$(medium)">
                        <label for="sitename">Site Name</label>
                        <input type="text" id="sitename" name="sitename" value="<%= strSiteTitle %>" />
                    </div>
                    <div class="4u 12u$(medium)">
                        <label for="domainname">Domain Name</label>
                        <input type="text" id="domainname" name="domainname" value="<%= strDomainname %>" />
                    </div>
                    <div class="4u$ 12u$(medium)">
                        <label for="submit">&nbsp;</label>
                        <input class="button fit" type="submit" name="submit" value="Save Settings" style="vertical-align: bottom;" />
                    </div>
                </div>
            </form>

            <h3>Calendar Options</h3>
            <form action="admin_settings.asp" id="allowusers" name="allowusers" method="post">
                <input type="hidden" name="change" value="yes" />
                <h4>Allow Users To Schedule and Delete Events?</h4>
                <h5 style="color: #F00;">*Should only be enabled in a trusted/secure Environment*<br />
                    Consider adding an Admin instead!
                </h5>
                <% If blnLetUsers Then %>
                <input type="radio" id="letyes" name="letusers" value="1" checked />
                <label for="letyes">Yes</label>
                <input type="radio" id="letno" name="letusers" value="0" />
                <label for="letno">No</label>
                <% Else %>
                <input type="radio" id="letyes" name="letusers" value="1" checked />
                <label for="letyes">Yes</label>
                <input type="radio" id="letno" name="letusers" value="0" checked />
                <label for="letno">No</label>
                <% End If %>
                <h4>Delete Events older than</h4>
                <input type="text" id="ndays" name="ndays" value="<%= intDelDays %>" style="width: 75px;" />
                <label for="ndays">Days: 0 = All past events.</label>
                <label for="announce">Announcements</label>
                <textarea id="announce" name="announce" rows="5"><%= strAnnouncements %></textarea>
                <input class="button fit" type="submit" value="Edit Settings" style="margin-top:10px;" />
            </form>
        </div>
    </div>
</div>
<!-- #include file="../includes/footer.asp"-->