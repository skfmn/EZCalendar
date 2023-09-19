<!--#include file="../includes/general_includes.asp"-->
<%
	strCookies = Request.Cookies("EZCalAdmin")("name")
	
	If strCookies = "" Then

		Response.Redirect "admin_login.asp"
  
	End If

    msg = ""
    msg = Trim(Request.Cookies("msg"))

	If msg <> "" Then
		Call displayFancyMsg(getMessage(msg))
        Response.Cookies("msg") = ""
	End If
	
    'You can delete this after the folder is gone.'''''''''''''''''''''''''''''''''''''''''
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(Server.MapPath(strDir&"install")) Then
		fso.DeleteFolder(Server.MapPath(strDir&"install"))
	End If
	Set fso = Nothing
    'You can delete this after the folder is gone.'''''''''''''''''''''''''''''''''''''''

%>
<!-- #include file="../includes/header.asp"-->
<div id="main" class="container">
    <header>
        <h1 style="text-align: center;font-size:30px">EZCalendar</h1>
        <h4 style="text-align: center;">Choose an Option below</h4>
    </header>
  <div class="row">
         <div class="-3u 3u 12u$(medium)">
            <ul class="alt">
                <li>
                    <a class="button fit" href="admin_schedule.asp">
                        <span>Schedule Event</span>
                    </a>
                </li>
            </ul>
        </div>
        <div class="3u$ 12u$(medium)">
            <ul class="alt">
                <li>
                    <a class="button fit" href="admin_view.asp">
                        <span>Manage Events</span>
                    </a>
                </li>
            </ul>
        </div>
        <div class="-3u 3u 12u$(medium)">
            <ul class="alt">
                <li>
                    <a class="button fit" href="admin_settings.asp">
                        <span>Manage Settings</span>
                    </a>
                </li>
            </ul>
        </div>
        <div class="3u$ 12u$(medium)">
            <ul class="alt">
                <li>
                    <a class="button fit" href="admin_manage.asp">
                        <span>Manage Admins</span>
                    </a>
                </li>
            </ul>
        </div>
        <div class="-3u 6u 12u$(medium)">
		    <%= getResponse("http://www.aspjunction.com/gnews.asp?calv="& strVersion&"") %>
        </div>	
    </div> 
</div>
<!-- #include file="../includes/footer.asp"-->