<%
	Response.ExpiresAbsolute = Now() - 2
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "No-Store"
	Response.AddHeader "If-Modified-Since",now
	Response.AddHeader "Last-Modified",now
	Response.Expires = 0

	Dim Conn
	Dim ConnStr
	Dim strSQL
	Dim rsCommon 
	Dim strEventsStr
	Dim strAddInfo
	Dim strSchedDate
	Dim blnDBDate
	Dim schedid
	Dim lngSchedID
	Dim iMonth
	Dim iYear
	Dim dTemp
	Dim dDate
	Dim iDIM
	Dim iDOW
	Dim sDate
	Dim datCurr
	Dim strName
	Dim dAnyDayInTheMonth
	Dim iPosition
	Dim iCurrent
	Dim datSDate
	Dim blnMCheck
	Dim mcount
	Dim dcount
	Dim ycount
	Dim datDate
	Dim datEnd
	Dim datTemp
	Dim strLetUsers
	Dim strVersion
	Dim msg
	Dim blnAllowReg
	Dim action
	Dim intDelDays
	Dim strAnnouncements
	Dim strUserName
	Dim strEventTitle
	Dim strEventText
	Dim strSiteTitle
	Dim strDomainname
	Dim strPurgeText
	Dim datDelDate
	Dim strChecked
	Dim strRight

	Dim blnSchedule
	Dim blnEvents
	Dim blnSettings
	Dim blnAdminRights
	Dim blnARights
	Dim blnPurge
	
	blnLetUsers = false
	
%>
