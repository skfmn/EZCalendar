<!-- #include file="includes/general_includes.asp"-->
<html>
<head>
    <link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
    <link type="text/css" rel="stylesheet" href="/calendar/assets/css/jquery.fancybox.css" />
    <link type="text/css" rel="stylesheet" href="/calendar/assets/css/main.css" />
    <script language="javascript" type="text/javascript">
        function validateSched() {
            with (window.document.schedule) {
                if (event_name.value == "") {
                    alert('Please enter an Event Name!');
                    event_name.focus();
                    return false;
                }
                if (mcount.selectedIndex == 0) {
                    alert('Please select a month!');
                    mcount.focus();
                    return false;
                }
                if (ycount.selectedIndex == 0) {
                    alert('Please select a year!');
                    ycount.focus();
                    return false;
                }
                if (comments.value == "") {
                    alert('Please enter a comment for the event!');
                    comments.focus();
                    return false;
                }
                return true;
            }
        }

        function confirmSubmit(imsg, ihref) {
            var smsg = confirm(imsg);
            if (smsg == true) {
                window.location = ihref;
            } else {
                return false;
            }
        }

        function focusarea(iMonthID, iNum) {
            this.Num = iNum;
            this.monthID = iMonthID;
        }

        function getFocusArea() {

            var dyear = document.schedule.ycount.value;
            if (leapYear(dyear)) {
                aFocusArea[2] = new focusarea(2, 29);
            } else {
                aFocusArea[2] = new focusarea(2, 28);
            }

            var sSelect = '<select id="dcount" name="dcount">';
            var selectID = document.schedule.mcount.value;

            for (var i = 1; i < aFocusArea[selectID].Num + 1; i++) {
                sSelect = sSelect + '<option>' + [i] + '</option>';
            }

            sSelect = sSelect + '</select>';
            document.getElementById('Focus').innerHTML = "";
            document.getElementById('Focus').innerHTML = sSelect;
        }

        function setFocusArea() {

            var mSelect = '<select id="mcount" name="mcount" onchange="getFocusArea();">';
            mSelect = mSelect + '<option value="0">Select Month</option>';

            for (var x = 1; x <= 12; x++) {
                mSelect = mSelect + '<option value="' + x + '">' + x + '</option>';
            }

            mSelect = mSelect + '</select>';
            document.getElementById('mFocus').innerHTML = "";
            document.getElementById('mFocus').innerHTML = mSelect;

            var sSelect = '<select name="dcount">';
            sSelect = sSelect + '<option>Select Day</option>';
            sSelect = sSelect + '</select>';
            document.getElementById('Focus').innerHTML = "";
            document.getElementById('Focus').innerHTML = sSelect;

        }

        function leapYear(lyear) {
            return ((lyear % 4 == 0) && (lyear % 100 != 0)) || (lyear % 400 == 0);
        }

        var aFocusArea = new Array;
        aFocusArea[1] = new focusarea(1, 31);
        aFocusArea[2] = new focusarea(2, 29);
        aFocusArea[3] = new focusarea(3, 31);
        aFocusArea[4] = new focusarea(4, 30);
        aFocusArea[5] = new focusarea(5, 31);
        aFocusArea[6] = new focusarea(6, 30);
        aFocusArea[7] = new focusarea(7, 31);
        aFocusArea[8] = new focusarea(8, 31);
        aFocusArea[9] = new focusarea(9, 30);
        aFocusArea[10] = new focusarea(10, 31);
        aFocusArea[11] = new focusarea(11, 30);
        aFocusArea[12] = new focusarea(12, 31);
    </script>
</head>
<body>
<%
    If Trim(Request.QueryString("delete")) = "yes" Then

        lngSchedID = checkint(Trim(Request.QueryString("schedid")))

        Set Conn = Server.CreateObject("ADODB.Connection")
        Call ConnOpen(Conn)

        strSQL = "DELETE FROM "&msdbprefix&"calendar WHERE schedID = "&lngSchedID
        Call getExecuteQuery(strSQL)

        strSQL = "DELETE FROM "&msdbprefix&"registration WHERE schedID = "&lngSchedID
        Call getExecuteQuery(strSQL)

        Call ConnClose(Conn)

        Response.Cookies("msg") = "evdel"
        Response.Redirect "calendar.asp"

    End If

    If Trim(Request.QueryString("schedule")) = "now" Then

        strSchedDate = Trim(Request.Form("mcount")) & "/" & Trim(Request.Form("dcount")) & "/" & Trim(Request.Form("ycount"))

        If Not IsDate(strSchedDate) Then

            Response.Redirect "calendar.asp?msg=notad"

        Else

            If Trim(Request.Form("registration")) = "yes" Then
                blnAllowReg = 1
            Else
                blnAllowReg = 0
            End If

            strSchedDate = CDate(strSchedDate)

            strSQL = "INSERT INTO "&msdbprefix&"calendar ([allow_reg],[schDate],[event],[text]) VALUES ("&blnAllowReg&",'"&strSchedDate&"','"&DBEncode(Trim(Request.Form("event_name")))&"','"&DBEncode(Trim(Request.Form("comments")))&"')"

            Set Conn = Server.CreateObject("ADODB.Connection")
            Call ConnOpen(Conn)
            Call getExecuteQuery(strSQL)
            Call ConnClose(Conn)

            Response.Cookies("msg") = "evsch"
            Response.Redirect "calendar.asp"

        End If
    End If

    If Trim(Request.QueryString("process")) = "reg" Then

        lngSchedID = checkint(Trim(Request.QueryString("schedid")))
        strName = DBEncode(Trim(Request.Form("rname")))
        strAddInfo = DBEncode(Trim(Request.Form("addinfo")))
        datDate = Cdate(Trim(Request.QueryString("date")))

        strSQL = "INSERT INTO "&msdbprefix&"registration(schedID,reg_name,add_info) values("&lngSchedID&",'"&strName&"','"&strAddInfo&"')"
        Set Conn = Server.CreateObject("ADODB.Connection")
        Call ConnOpen(Conn)
        Call getExecuteQuery(strSQL)
        Call ConnClose(Conn)

        Response.Cookies("msg") = "regs"
        Response.Redirect "calendar.asp?date="&datDate

    End If

	If Trim(Request.QueryString("reg")) = "yes" Then

	    lngSchedID = checkint(Trim(Request.QueryString("schedid")))

		Set Conn = Server.CreateObject("ADODB.Connection")
		Set rsCommon = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM "&msdbprefix&"calendar WHERE SchedID = "&lngSchedID
		Call ConnOpen(Conn)
		Call getTextRecordset(strSQL,rsCommon)
		If Not rsCommon.EOF Then
		    strEventsStr = rsCommon("event")
			datDate = rsCommon("schDate")
		End If
		Call closeRecordset(rsCommon)
		Call ConnClose(Conn)

%>
    <div id="main" class="container" style="max-width: 1200px;">
        <header style="text-align: center;">
            <h2>Register</h2>
        </header>
        <h4 style="text-align: center;"><%= DBDecode(strEventsStr) %>&nbsp;&nbsp;<%= datDate %></h4>
        <form action="calendar_view.asp?process=reg&schedid=<%= lngSchedID %>&date=<%= datDate %>" method="post" name="register" id="register">
            <div class="row">
                <div class="-3u 6u 12u(medium)" style="padding-bottom: 30px;">
                    <label for="rname" style="margin-bottom: -3px;">Name</label>
                    <input type="text" id="rname" name="rname" size="20" />
                </div>
                <div class="-3u 6u 12u(medium)" style="padding-bottom: 10px;">
                    <label for="addinfo" style="margin-bottom: -3px;">Additional info:</label>
                    <textarea id="addinfo" name="addinfo"></textarea>
                </div>
                <div class="-3u 6u 12u(medium)">
                    <input type="button" class="button fit" name="register" value="Register" onclick="document.register.submit()" />
                </div>
            </div>
        </form>
    </div>
<%
	End If

	If Request.QueryString("view") = "yes" Then

	    If Trim(Request.QueryString("sdate")) <> "" Then datDate = Trim(Request.QueryString("sdate"))

		If Trim(Request.QueryString("sdate")) <> "" Then
		    strSQL = "SELECT * FROM "&msdbprefix&"calendar WHERE schDate = '"&CDate(Trim(Request.QueryString("sdate")))&"' ORDER BY schedID Asc"
		ElseIf Trim(Request.QueryString("schedid")) <> "" Then
		    strSQL = "SELECT * FROM "&msdbprefix&"calendar WHERE schedID = "&checkint(Trim(Request.QueryString("schedid")))&" ORDER BY schedID Asc"
		Else
		    strSQL = "SELECT * FROM "&msdbprefix&"calendar WHERE schDate = '"&date&"' ORDER BY schedID Asc"
		End If

		Set Conn = Server.CreateObject("ADODB.Connection")
	    Call ConnOpen(Conn)

		Set rsCommon2 = Server.CreateObject("ADODB.Recordset")

		Call getTextRecordset(strSQL,rsCommon2)
		If Not rsCommon2.EOF Then
%>
    <div id="view">
        <h3 style="text-align: center;">Events Scheduled For<br />
            <%= datDate %></h3>
<%
	        Do While Not rsCommon2.EOF

	            intSchedID = rsCommon2("schedID")
				datDate = rsCommon2("schDate")
				strEventName = DBDecode(rsCommon2("event"))
				strComments = DBDecode(rsCommon2("text"))

                Set rsTemp = Server.CreateObject("ADODB.Recordset")
	            strSQL = "SELECT * FROM "&msdbprefix&"registration WHERE schedID = "&intSchedID
				Call getTextRecordset(strSQL,rsTemp)
				If Not rsTemp.EOF Then
					intRCount = rsTemp.RecordCount
				Else
					intRCount = 0
				End If
				Call closeRecordset(rsTemp)

%>
        <div id="view_title">
            <h2><%= strEventName %></h2>
        </div>
        <div id="view_comments">
            <pre class="pre"><%= strComments %></pre>
        </div>
        <br />
        <br />
        <% If rsCommon2("allow_reg") Then %>
        <div align="center">
            <a class="button picimg fancybox.ajax" href="calendar_view.asp?reg=yes&schedid=<%= intSchedID %>">Register</a>
            <br />
            <%
			If intRCount = 0 Then
	            Response.Write "No one has registered. Be the first!"
	        ElseIf intRCount = 1 Then
	            Response.Write "1 person has already registered!"
	        Else
	            Response.Write intRCount&" people have already registered!"
	        End If
            %>
        </div>
        <% End If %>
        <% If blnLetUsers Then %>
        <br />
        <br />
        <div align="center">
            <a class="button" style="cursor: pointer;" onclick="return confirmSubmit('Are you SURE you want to delete this event?','calendar_view.asp?delete=yes&schedid=<%= intSchedID %>')">Delete</a>
        </div>
        <% End If %>
    </div>
    <hr class="major" />
<%
	            rsCommon2.MoveNext
	            If rsCommon2.EOF Then Exit Do
	        Loop
		End If
		Call closeRecordset(rsCommon2)
		Call ConnClose(Conn)
	End If

	If Trim(Request.QueryString("sched")) = "yes" Then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Call ConnOpen(Conn)
%>
    <div id="main" class="container" style="max-width: 1200px;">
        <div class="row">
            <div class="-3u 6u$">
                <header>
                    <h2>Schedule an Event</h2>
                </header>
            </div>
        </div>
        <form action="calendar_view.asp?schedule=now&date=<%= date %>" name="schedule" method="post">
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
                    <input type="radio" id="regno" name="registration" value="no">
                    <label for="regno">No</label>
                </div>
                <div class="-3u 6u 12u(medium)" style="padding-bottom: 10px;">
                    <label for="comments" style="margin-bottom: -2px;"><strong>Comments</strong></label>
                    <textarea id="comments" name="comments" cols="50" rows="10"></textarea>
                </div>
                <div class="-3u 6u 12u(medium)">
                    <a class="button" onclick="document.schedule.submit()">Submit</a>
                </div>
            </div>
        </form>
    </div>
    <%
    Call ConnClose(Conn)
  End If
    %>
</body>
</html>