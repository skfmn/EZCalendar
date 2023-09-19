<%
on error resume next
    Sub trace(strText)
        Response.Write "Debug: "&strText&"<br />"&vbcrlf
    End Sub
	
    Sub catch(sText,sText2)

        If Err.Number <> 0 then
            Call trace(sText&" - "&err.description)
        Else
            Call trace(sText&" - no error")
        End If

        If sText2 <> "" Then
            Call trace(sText&" - "&sText2)
        End If
				
        on error goto 0	
    End Sub

	Function DBEncode(DBvalue)
		Dim fieldvalue
		fieldvalue = Trim(DBvalue)

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
%>
<!DOCTYPE HTML>
<html>
<head>
<title>Install</title>
<link type="text/css" rel="stylesheet" href="../assets/css/main.css" />
</head>
<body>
  <div id="main" class="container" align="center" style="margin-top:-75px;">
    <div class="row 50%">
      <div class="12u 12u$(medium)">
        <header><h2>EZCalendar Installation</h2></header>
      </div>
    </div>
  </div>
<% If Trim(Request.QueryString("step")) = "one" Then %>
  <div id="main" class="container" align="center" style="margin-top:-100px;">
    <div class="row 50%">
      <div class="12u 12u$(medium)">
        <form action="install.asp?setsql=y" method="post">
        <header>
          <h2>MSSQL Database</h2>
        </header>
        <div class="row">
          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="svrname" style="text-align:left;">Server Host Name or IP Address
              <input type="text" name="svrname" required>
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="dbname" style="text-align:left;">Database Name
              <input type="text" name="dbname" required>
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="dbid" style="text-align:left;">Database Login
              <input type="text" name="dbid" required>
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="dbpwd" style="text-align:left;">Database Password
              <input type="password" name="dbpwd" required>
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="dbprefix" style="text-align:left;">Table Prefix
              <input type="text" name="dbprefix" value="ezcal_" required>
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="12u 12u$(medium)">
           <input class="button" type="submit" name="submit" value="Continue">
          </div>
        </div>
        </form>     
      </div>
    </div>
  </div>
<% 
  ElseIf Request.QueryString("setsql") = "y" Then
%>
  <div id="main" class="container" align="center">
    <div class="row 50%">
      <div class="12u 12u$(medium)">
<%

    msdbserver = Trim(Request.Form("svrname"))
    msdb = Trim(Request.Form("dbname"))
    msdbid = Trim(Request.Form("dbid"))
    msdbpwd = Trim(Request.Form("dbpwd"))
    msdbprefix = Trim(Request.Form("dbprefix"))

    Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Open "Provider=sqloledb;Data Source="&msdbserver&";Initial Catalog="&msdb&";User Id="&msdbid&";Password="&msdbpwd

    Response.Write "Creating Database Tables<br /><br />"
    Response.Write "Creating admin table...<br />"
    Response.Flush
	
    Conn.Execute "CREATE TABLE "&msdbprefix&"admin " & _
    "([adminID] [numeric](10, 0) IDENTITY (1, 1) CONSTRAINT [PK_"&msdbprefix&"admin] PRIMARY KEY," & _
    "[name] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    "[pwd] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
    "[salt] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
    "[schedule]  [bit] NULL ," & _
    "[events]  [bit] NULL ," & _
    "[settings]  [bit] NULL ," & _
    "[admin_rights]  [bit] NULL ," & _
    "[arights]  [bit] NULL, " & _
    "[purge]  [bit] NULL)"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0	 

    Response.Write "Populating admin table...<br /><br />"
    Response.Flush
	
    Conn.Execute "INSERT INTO "&msdbprefix&"admin ([name],[pwd],[salt],[schedule],[events],[settings],[admin_rights],[arights],[purge]) VALUES ('admin','EB36FB0C1F1A92A838AA1ECAAD4AB6E3B5257103','833D1','True','True','True','True','True','True')"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

    Response.Write "Creating settings table...<br />"
    Response.Flush
	
    Conn.Execute "CREATE TABLE "&msdbprefix&"settings " & _
    "([site_title] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
    "[domain_name] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
    "[letusers] [bit] NULL ," & _
    "[delete_days] [numeric] (10, 0)  NULL," & _
    "[announcements] [nvarchar] (MAX) COLLATE SQL_Latin1_General_CP1_CI_AS NULL)"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

    Response.Write "Populating settings table...<br /><br />"
    Response.Flush
	
    Conn.Execute "INSERT INTO "&msdbprefix&"settings ([letusers],[delete_days],[announcements]) VALUES (0,30,'Announcements go here!')"


    Response.Write "Creating Messages table...<br />"
    Response.Flush
				  
    Conn.Execute "CREATE TABLE "&msdbprefix&"messages " & _ 
    "([messageID] [numeric] IDENTITY (1, 1) CONSTRAINT [PK_"&msdbprefix&"messages] PRIMARY KEY," & _
    "[msg] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[message] [nvarchar] (MAX) COLLATE SQL_Latin1_General_CP1_CI_AS NULL)"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0
  
    Response.Write "Populating Messages table...<br /><br />"
    Response.Flush
	
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('logch','Your login information was successfully changed.')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('evsch','Event was successfully scheduled.')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('evmod','The event information was successfully changed.')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('setch','The settings were successfully etadpud.')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('evdel','The event was successfully deleted.')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('regs','Registration was successful!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('ps','Purge of past events was successful!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('notad','You did not enter a date!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('nadmin','You can not change Main Admins info!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('das','You successfully eteledd the Admin.')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('adad','You have successfully added an Admin.')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('ant','Admin name taken!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('cpwds','Password changed successfully!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('siu','Site Info etadpud!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('nar','You do not have sufficient rights!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('car','You have successfully modified Admin Rights.')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('mus','Messages etadpud successfully!')"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0
		   
    Response.Write "Creating Registration table...<br /><br />"
    Response.Flush
				  
    Conn.Execute "CREATE TABLE "&msdbprefix&"registration " & _ 
    "([regID] [numeric] IDENTITY (1, 1) CONSTRAINT [PK_"&msdbprefix&"registration] PRIMARY KEY," & _
    "[schedID] [numeric] (10, 0) NULL ," & _
    "[reg_name] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
    "[add_info] [nvarchar] (MAX) COLLATE SQL_Latin1_General_CP1_CI_AS NULL)"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0
		
    Response.Write "Creating Calendar table...<br />"
    Response.Flush
				  
    Conn.Execute "CREATE TABLE "&msdbprefix&"calendar " & _ 
    "([schedID] [numeric] IDENTITY (1, 1) CONSTRAINT [PK_"&msdbprefix&"calendar] PRIMARY KEY," & _
    "[allow_reg] [bit] NULL, " & _
    "[schDate] [smalldatetime] NULL, " & _
    "[event] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[text] [nvarchar] (MAX) COLLATE SQL_Latin1_General_CP1_CI_AS NULL)"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0
		
    Response.Write "Creating database tables...Complete!<br />"
    Response.Flush
						
    Response.Write "<br /><br />"
%>
      </div>
    </div>
  </div>
  <div id="main" class="container" align="center">
    <div class="row 50%">
      <div class="12u 12u$(medium)">
        <form action="install.asp?step=two" method="post">
        <input type="hidden" name="msdbserver" value="<%= msdbserver %>">
        <input type="hidden" name="msdb" value="<%= msdb %>">
        <input type="hidden" name="msdbid" value="<%= msdbid %>">
        <input type="hidden" name="msdbpwd" value="<%= msdbpwd %>">
        <input type="hidden" name="msdbprefix" value="<%= msdbprefix %>">
        <header>
          <h3><span class="first">You have successfully installed the MSSQL Database!</span></h3>
        </header>
        <div class="row">
          <div class="12u 12u$(medium)">
           <input class="button" type="submit" name="submit" value="Continue">
          </div>
        </div>
        </form>     
      </div>
    </div>
  </div>
<%  
		Conn.Close: Set Conn = Nothing

  ElseIf Request.QueryString("step") = "two" Then  
%>
  <div id="main" class="container" align="center">
    <div class="row 50%">
      <div class="12u 12u$(medium)">
        <form action="install.asp?step=three" method="post">
        <input type="hidden" name="msdbserver" value="<%= Trim(Request.Form("msdbserver")) %>">
        <input type="hidden" name="msdb" value="<%= Trim(Request.Form("msdb")) %>">
        <input type="hidden" name="msdbid" value="<%= Trim(Request.Form("msdbid")) %>">
        <input type="hidden" name="msdbpwd" value="<%= Trim(Request.Form("msdbpwd")) %>">
        <input type="hidden" name="msdbprefix" value="<%= Trim(Request.Form("msdbprefix")) %>">
        <input type="hidden" name="PhyPath" value="<%= strPhysPath %>" />
        <header>
          <h2>Path Settings</h2>
        </header>
        <div class="row">
          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="dbid" style="text-align:left;">Base Directory
              <input type="text" name="bdir" value="<%= Request.ServerVariables("APPL_PHYSICAL_PATH") %>" />
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="dir" style="text-align:left;">EZCalendar Directory
              <input type="text" name="dir" value="/calendar/" size="40" />
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>
          <div class="12u 12u$(medium)">
           <input class="button" type="submit" name="submit" value="Continue">
          </div>
        </div>
        </form>
      </div>
    </div>
  </div>
<%
    ElseIf Request.QueryString("step") = "three" Then 

        strPageFileName = Server.MapPath("../includes/config.asp")

        Set objPageFileFSO = CreateObject("Scripting.FileSystemObject")

        If objPageFileFSO.FileExists(strPageFileName) Then
        Set objPageFileTs = objPageFileFSO.OpenTextFile(strPageFileName, 2)
        Else
        Set objPageFileTs = objPageFileFSO.CreateTextFile(strPageFileName)
        End If

        strPageEntry = Chr(60) & Chr(37) & vbcrlf & _
        "baseDir=""" & Trim(Request.Form("bdir")) & """" & vbcrlf & _
        "strDir=""" & Trim(Request.Form("dir")) & """" & vbcrlf & _
        "msdbprefix=""" & Trim(Request.Form("msdbprefix")) & """" & vbcrlf & _
        "msdbserver=""" & Trim(Request.Form("msdbserver")) & """" & vbcrlf & _
        "msdb=""" & Trim(Request.Form("msdb")) & """" & vbcrlf & _
        "msdbid=""" & Trim(Request.Form("msdbid") )& """" & vbcrlf & _
        "msdbpwd=""" & Trim(Request.Form("msdbpwd")) & """" & vbcrlf & _
        Chr(37) & Chr(62)
				 
        objPageFileTs.WriteLine strPageEntry
  
        objPageFileTs.Close

        Response.Redirect "install.asp?step=four"

    ElseIf Request.QueryString("step") = "four" Then 
%>
  <div id="main" class="container" style="margin-top:-100px;">
    <div class="row">
      <div class="12u 12u$(medium)" style="text-align:center;">
        <form action="install.asp?step=five" method="post">
        <header>
          <h2>Other stuff</h2>
        </header>
        <div class="row">

          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="sitetitle" style="text-align:left;">Site title
              <input type="text" name="sitetitle" />
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="-4u 4u 12u$(medium)" style="padding-bottom:20px;">
            <label for="domainname" style="text-align:left;">Domain name
              <input type="text" name="domainname" value="<%= Request.ServerVariables("SERVER_NAME") %>" />
            </label>
          </div>
          <div class="4u 1u$"><span></span></div>

          <div class="12u 12u$(medium)">
           <input class="button" type="submit" name="submit" value="Continue">
          </div>
        </div>
        </form>      
      </div>
    </div>
  </div>
<%
  ElseIf Request("step") = "five" Then
    %><!-- #include file="../includes/config.asp"--><%
    Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Open "Provider=sqloledb;Data Source="&msdbserver&";Initial Catalog="&msdb&";User Id="&msdbid&";Password="&msdbpwd

    Conn.Execute "UPDATE "&msdbprefix&"settings SET site_title = '"&DBEncode(Request.Form("sitetitle"))&"', domain_name = '"&DBEncode(Request.Form("domainname"))&"'"

    Conn.Close: Set Conn = Nothing

    Response.Redirect "install.asp?step=done"

  ElseIf Request("step") = "done" Then
%>
  <div id="main" class="container">
    <div class="row">
      <div class="12u 12u$(medium)" style="text-align:center;">
        <span class="first">
          Success!
          <br />
          You have successfully configured EZCalendar!
          <br />
          The next step is to change your password.
          <br />
          Click on the link below and login to admin.
          <br />
          Click on "Password" in the left options menu and change your password.
          <br /><br />
          <a class="button" href="../admin/admin_login.asp">Login</a>
        </span>
      </div>
    </div>
  </div>
<% Else %>
  <div id="main" class="container" style="margin-top:-75px;">
    <div class="row">
      <div class="12u 12u$(medium)" style="text-align:center;">
        <span class="first">
	      You are about to install EZCalendar.
	      <br />
	      Please follow the instructions carefully!
	      <br /><br />
	      <input class="button" type="button" onclick="parent.location='install.asp?step=one'" value="Continue">
	      </span>      
      </div>
    </div>
  </div>
<% End If %>
<br />
</body>
</html>