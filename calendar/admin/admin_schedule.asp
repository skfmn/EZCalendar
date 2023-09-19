<!--#include file="../includes/general_includes.asp"-->
<%
	strCookies = Request.Cookies("EZCalAdmin")("name")

	If strCookies = "" Then

		Response.Redirect "admin_login.asp"

	End If

	If Not blnSchedule Then
	    Response.Cookies("msg") = "nar"
	    Response.Redirect "admin.asp"
	End If

    msg = ""
    msg = Trim(Request.Cookies("msg"))

	If msg <> "" Then
		Call displayFancyMsg(getMessage(msg))
        Response.Cookies("msg") = ""
	End If

    If Trim(Request.Form("schedule")) = "now" Then

        strSchedDate = Trim(Request.Form("mcount")) & "/" & Trim(Request.Form("dcount")) & "/" & Trim(Request.Form("ycount"))

        If Not IsDate(strSchedDate) Then

		    Response.Cookies("msg") = "notad"
            Response.Redirect "admin_schedule.asp"

        Else

            If Trim(Request.Form("registration")) = "yes" Then
                blnAllowReg = 1
            Else
                blnAllowReg = 0
            End If

            strSchedDate = CDate(strSchedDate)
            strSQL = "INSERT INTO "&msdbprefix&"calendar([allow_reg],[schDate],[event],[text]) VALUES ("&blnAllowReg&",'"&strSchedDate&"','"&DBEncode(Request.Form("event_name"))&"','"&DBEncode(Request.Form("comments"))&"')"

            Set Conn = Server.CreateObject("ADODB.Connection")
            Call ConnOpen(Conn)
            Call getExecuteQuery(strSQL)
            Call ConnClose(Conn)

            Response.Cookies("msg") = "evsch"
            Response.Redirect "admin_schedule.asp"

        End If

    End If

%>
<!-- #include file="../includes/header.asp"-->
<div id="main" class="container">
    <div class="row">
        <div class="-3u 6u$ 12u$(medium)">
            <header>
                <h2>Schedule an Event</h2>
            </header>
        </div>
    </div>
     <script type="text/javascript" src="../assets/js/articles.js"></script>
    <form action="admin_schedule.asp" method="post" name="schedule" onsubmit="return validateSched();">
        <input type="hidden" name="schedule" value="now" />
        <div class="row">
            <div class="-3u 6u 12u(medium)" style="padding-bottom: 30px;">
                <label for="event_name" style="margin-bottom: -2px;"><strong>Event Name</strong></label>
                <input type="text" id="event_name" name="event_name" />
            </div>
            <div class="-3u 6u 12u(medium)">
                <h5 style="margin-bottom: -2px;">Select Date</h5>
            </div>
             <div class="-3u 2u 12u$(medium)">
                <div class="select-wrapper">
                    <select id="ycount" name="ycount" onchange="setFocusArea();">
                    <option value="0">Select Year</option>
                    <%
                    Dim intYear: intYear = Year(Date)
                    For ycount = intYear to intYear+10
                        Response.Write " <option value="""&ycount&""">" & ycount & "</option>" & vbcrlf
                    Next
                    %>
                    </select>
                </div>
            </div>
            <div class="2u 12u$(medium)">
                <div class="select-wrapper">
                    <span id="mFocus">
                    <select name="mcount">
                        <option>Select Month</option>
                    </select>
                    </span>
                </div>
            </div>
            <div class="2u$ 12u$(medium)">
                <div class="select-wrapper">
                    <span id="Focus">
                    <select name="dcount">
                        <option>Select Day</option>
                    </select>
                    </span>
                </div>
            </div>
            <div class="-3u 2u 12u(medium)" style="padding-top: 30px;">
                <h5 style="margin-bottom: -3px;"><strong>Allow Registration</strong></h5>
            </div>
            <div class="2u$ 12u(medium)" style="padding-top: 30px;">
                <input type="radio" id="regyes" name="registration" value="yes">
                <label for="regyes">Yes</label>
                <input type="radio" id="regno" name="registration" value="no" checked>
                <label for="regno">No</label>
            </div>
            <div class="-3u 6u 12u(medium)" style="padding-bottom: 10px;">
                <label for="comments" style="margin-bottom: -2px;"><strong>Comments</strong></label>
                <textarea id="comments" name="comments" cols="50" rows="10"></textarea>
            </div>
            <div class="-3u 6u 12u(medium)">
                <a class="button fit" onclick="document.schedule.submit()">Submit</a>
            </div>
        </div>
    </form>
</div>
<!-- #include file="../includes/footer.asp"-->