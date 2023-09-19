<!-- #include file="includes/general_includes.asp"-->
<%
on error resume next
    msg = ""
    msg = Trim(Request.Cookies("msg"))

	If msg <> "" Then
		Call displayFancyMsg(getMessage(msg))
        Response.Cookies("msg") = ""
	End If
	
    If IsDate(Trim(Request.QueryString("sDate"))) Then
	    datSDate = CDate(Trim(Request.QueryString("sDate")))
	ElseIf IsDate(Trim(Request.QueryString("date"))) Then
	    datSDate = CDate(Trim(Request.QueryString("date")))
	Else
	    datSDate = date
	End If
	 
	If IsDate(Trim(Request.QueryString("date"))) Then
	  dDate = CDate(Trim(Request.QueryString("date")))
	Else
	  dDate = Date
	End If

    iDIM = GetDaysInMonth(Month(dDate), Year(dDate))
    iDOW = GetWeekdayMonthStartsOn(dDate)
%>
<html>
<head>
  <title>EZCalendar</title>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
  <link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
  <link type="text/css" rel="stylesheet" href="/calendar/assets/css/jquery.fancybox.css" />
  <link type="text/css" rel="stylesheet" href="/calendar/assets/css/main.css" />
</head>
<body>
<div id="main" class="container">
  <header style="text-align:center;"><h2>Calendar</h2></header>
  <div class="row">
    <div class="-3u 1u 12u(medium)">
      <a class="button" href="calendar.asp?date=<%= SubtractOneMonth(dDate) %>"><i class="fa fa-arrow-left" style="margin-top:15px;"></i></a>
    </div>
    <div class="-1u 2u 12u(medium)" style="text-align:center;">
      <%= MonthName(Month(dDate)) & "  " & Year(dDate) %>
    </div>
    <div class="-1u 1u$ 12u(medium)">
      <a class="button" href="calendar.asp?date=<%= AddOneMonth(dDate) %>"><i class="fa fa-arrow-right" style="margin-top:15px;"></i></a>
    </div>
    <div class="-3u 6u 12u(medium)" style="padding-top:10px;">
      <div class="table-wrapper">
        <table class="alt">
          <thead>
            <tr>
              <th>S</th>
              <th>M</th>
              <th>T</th>
              <th>W</th>
              <th>T</th>
              <th>F</th>
              <th>S</th>
            </tr>
          </thead>
          <tbody>
<%

	If iDOW <> 1 Then
		Response.Write "        <tr>" & vbCrLf
		iPosition = 1
		Do While iPosition < iDOW
			Response.Write "          <td>&nbsp;</td>" & vbCrLf
			iPosition = iPosition + 1
		Loop
	End If

	iCurrent = 1
	iPosition = iDOW

	Do While iCurrent <= iDIM
	
		If iPosition = 1 Then
		    Response.Write "        <tr>" & vbCrLf
		End If
		
		Set Conn = Server.CreateObject("ADODB.Connection")
	    Call ConnOpen(Conn)

		sDate = CDate(Month(dDate) & "/" & iCurrent & "/" & Year(dDate))	
		blnDBDate = schedCheck(sDate)

		If blnDBDate Then
			strEventsStr = ""
			intSchedID = 0
			Set Conn = Server.CreateObject("ADODB.Connection")
	        Call ConnOpen(Conn)

			Set rsCommon = Server.CreateObject("ADODB.Recordset")
			strSQL = "SELECT * FROM "&msdbprefix&"calendar WHERE schDate = '"&sDate&"' ORDER BY SchedID ASC"

			Call getTextRecordset(strSQL,rsCommon)
			If not rsCommon.EOF Then
			    strEvent = DBDecode(rsCommon("event"))
			    intSchedID = rsCommon("schedID")
			End If
			Call closeRecordset(rsCommon)
	        Call ConnClose(Conn)

		End If
		Call ConnClose(Conn)
		
	    datSDate = Cdate(DateSerial(year(dDate),month(dDate),iCurrent))

		If blnDBDate AND datSDate = date Then
	        'highlight events scheduled for today
			If intSchedID <> 0 Then
				Response.Write "          <td class=""picimg fancybox.ajax"" style=""background-color:#676767;cursor:pointer;position:relative;height:75px;width:75px !important;"" href=""calendar_view.asp?view=yes&sdate="&datSDate&"""><canvas width=""96"" height=""75"" style=""position:absolute;top:0;left:0;"" id=""currevent""></canvas><span style=""position:absolute;top:1;left:1;color:#0000ff"">"&iCurrent&"</span></td>" & vbCrLf
			Else
				Response.Write "          <td style=""position:relative;height:75px;width:75px;""><span style=""position:absolute;top:1;left:1;"">"&iCurrent&"</span></td>" & vbCrLf
			End If

		ElseIf blnDBDate Then

			'highlight scheduled dates
			If intSchedID <> 0 Then
			    Response.Write "          <td class=""picimg fancybox.ajax"" style=""background-color:#FFFF00;cursor:pointer;position:relative;height:75px;width:75px;"" href=""calendar_view.asp?view=yes&sdate="&datSDate&"""><span style=""position:absolute;top:1;left:1;"">"&iCurrent&"</span></td>" & vbcrLf
			Else
			    Response.Write "          <td style=""position:relative;height:75px;width:75px;""><span style=""position:absolute;top:1;left:1;"">"&iCurrent&"</span></td>" & vbCrLf
			End If
			
		ElseIf Cdate(DateSerial(year(dDate),month(dDate),iCurrent)) = date Then
		
			'highlight todays date
			Response.Write "          <td style=""background-color:#676767;color:#FFFFFF;position:relative;height:75px;width:75px;""><span style=""position:absolute;top:1;left:1;"">"&iCurrent&"</span></td>" & vbCrLf
		
		Else
		
			'Rest of the days in the month 
		  Response.Write "          <td style=""position:relative;height:75px;width:75px;""><span style=""position:absolute;top:1;left:1;"">"&iCurrent&"</span></td>" & vbCrLf
			
		End If
		 
		' If we're at the endof a row then write </tr>
		If iPosition = 7 Then
			Response.Write "        </tr>" & vbCrLf
			iPosition = 0
		End If
		
		' Increment variables
		iCurrent = iCurrent + 1
		iPosition = iPosition + 1			
    Loop

	' Write spacer cells at end of last row if month doesn't end on a Saturday.
	If iPosition <> 1 Then
		Do While iPosition <= 7
			Response.Write "          <td>&nbsp;</td>" & vbCrLf
			iPosition = iPosition + 1
		Loop
        Response.Write "        </tr>" & vbCrLf
	End If	
%>
          </tbody>
        </table>
      </div>
    </div>
	  <div class="-3u 6u$ 12u$(medium)">
		  Legend<br />
		  <img src="assets/images/noevent-today.jpg" width="25" /> Today <img src="assets/images/event-any.jpg" width="25" /> Event <img src="assets/images/event-today.jpg" width="25" /> Event Today
	  </div>
<%If blnLetUsers Then%>
    <div class="-3u 6u 12u(medium)">
    <a class="button picimg fancybox.ajax" href="calendar_view.asp?sched=yes&date=<%= date %>">Schedule Event</a>
    <br /><br />
    </div>
<% End If %>
    <div class="-3u 6u 12u(medium)">
      <header><h3>Announcements</h3></header>
      <pre><%= strAnnouncements %></pre>
    </div>
    <br /><br />
  </div>
  <div class="-3u 6u 12u(medium)" style="text-align:center;">
    <!-- REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE LICENSE AGREEMENT -->
    <br /><br />
    <span style="font-size:16px">Powered by <a style="font-size:16px" href="http://www.aspjunction.com">EZCalendar</a> Copyright &copy; 2003 - <%= year(now) %> </span>
    <!-- REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE LICENSE AGREEMENT -->
  </div>
</div>
<script
  src="https://code.jquery.com/jquery-1.12.4.min.js"
  integrity="sha256-ZosEbRLbNQzLpnKIkEdrPv7lOy9C27hHQ+Xp8a4MxAQ="
  crossorigin="anonymous"></script>
<script language="javascript" type="text/javascript" src="assets/js/jquery.fancybox.js" ></script>
	
<script language="JavaScript">
	$(document).ready(function(){
		$(".picimg").fancybox({ maxWidth: 1200 });
		$("#textmsg").fancybox({
			afterClose : function() {
				location.href='calendar.asp';
			}
		});
		$("#textmsg").trigger('click');
	});

    var canvasElement = document.getElementById("currevent");
    var context = canvasElement.getContext("2d");

    // the triangle
    context.beginPath();
    context.moveTo(0, 0);
    context.lineTo(0, 75);
    context.lineTo(96, 75);
    context.closePath();

    // the outline
    context.lineWidth = 0;
    context.strokeStyle = '#FFFF00';
    context.stroke();

    // the fill color
    context.fillStyle = "#FFFF00";
    context.fill();
</script>
</body>
</html>