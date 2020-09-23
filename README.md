<div align="center">

## Complete ASP Calendar


</div>

### Description

Displays a dynamic calendar in html using asp
 
### More Info
 
You can input the month or the year, but neither is required

Just copy and paste the whole thing

a calendar


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matt Pantana](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matt-pantana.md)
**Level**          |Intermediate
**User Rating**    |4.3 (17 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__4-9.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matt-pantana-complete-asp-calendar__4-6668/archive/master.zip)

### API Declarations

none: distribute freely


### Source Code

```
<%
response.expires=0
dim CurMonth
dim CurDay
dim CurYear
dim NumOfDays
dim CurCell
dim onDay
dim FoundFirst
CurMonth = request.querystring("cmonth")
CurYear = request.querystring("cyear")
CurDay = request.querystring("cday")
if CurMonth = "" then curMonth = month(date)
if CurYear = "" then CurYear = year(date)
if CurDay = "" then CurDay = day(date)
FirstDay = weekday(CurMonth & "/01/" & CurYear)
cmonth = CurMonth
cyear = CurYear
NumOfDays = getlastday(cmonth,cyear)
FoundFirst = false
curcell = 1
onDay = 0
 function GetLastDay( tmonth, tyear )
 tmonth = tmonth + 1
 if tmonth > 12 Then
 tmonth = tmonth - 12
 tyear = tyear + 1
 End if
 Dim x
 x = DateAdd("d", -1, tmonth & "/01/" & tyear)
 GetLastDay = Day( x )
 End function
rows = 5
if firstday >= 5 and numofdays = 31 then
 rows = 6
end if
if firstday >= 6 and numofdays = 30 then
 rows = 6
end if
 function DayOf()
 if foundFirst then
	 onDay = OnDay + 1
	 if onDay > NumofDays then
	 DayOf = ""
	 else
	 DayOf = onDay
	 end if
	else
	 if curcell = Firstday or firstday = 1 then foundfirst = true
	 if foundFirst then
	 onDay = OnDay + 1
		if onDay > NumofDays then
		DayOf = ""
		else
		DayOf = onDay
	 end if
	 else
	 DayOf = ""
	 end if
	end if
  curcell = curcell + 1
	 if (OnDay + 1) = int(CurDay) and int(CurMonth) = int(month(date)) and int(CurYear) = year(date) then
	 bgcolor = "yellow"
	 else
	 bgcolor = ""
	 end if
 end function
%>
<html>
<head>
<title>Calendar</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<script language="VBScript">
Sub cmonth_onchange
frm.submit
end sub
Sub cyear_onchange
frm.submit
end sub
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000">
<table width="350" border="0" cellspacing="0" cellpadding="0">
 <form name=frm method=get action=calendar.asp>
 <tr>
  <td width="37"> </td>
  <td width="137"> </td>
  <td width="144"> </td>
  <td width="32"> </td>
 </tr>
 <tr>
  <td width="37"> </td>
  <td width="137">
  <select name="cmonth">
   <% for i = 1 to 12 %>
   <option value="<%=i%>" <%if int(curmonth) = i then response.write("Selected")%>><%=monthname(i)%></option>
   <% next %>
  </select>
  </td>
  <td width="144">
  <select name="cyear">
   <% for i = 2050 to 1980 step -1 %>
   <option value="<%=i%>" <%if int(curyear) = i then response.write("Selected")%>><%=i%></option>
   <% next %>
  </select>
  </td>
  <td width="32"> </td>
 </tr>
 <tr>
  <td width="37"> </td>
  <td width="137"> </td>
  <td width="144"> </td>
  <td width="32"> </td>
 </tr>
 </form>
</table>
<table width="350" border="0" cellspacing="0" cellpadding="0">
 <tr bgcolor="#666666">
 <td>
  <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">S</font></b></font></div>
 </td>
 <td>
  <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">M</font></b></font></div>
 </td>
 <td>
  <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">T</font></b></font></div>
 </td>
 <td>
  <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">W</font></b></font></div>
 </td>
 <td>
  <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">T</font></b></font></div>
 </td>
 <td>
  <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">F</font></b></font></div>
 </td>
 <td>
  <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">S</font></b></font></div>
 </td>
 </tr>
 <% for i = 1 to rows %>
 <tr>
 <td bgcolor="<%=bgcolor%>">
  <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"><%=DayOf%></font></div>
 </td>
 <td bgcolor="<%=bgcolor%>">
  <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"><%=DayOf%></font></div>
 </td>
 <td bgcolor="<%=bgcolor%>">
  <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"><%=DayOf%></font></div>
 </td>
 <td bgcolor="<%=bgcolor%>">
  <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"><%=DayOf%></font></div>
 </td>
 <td bgcolor="<%=bgcolor%>">
  <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"><%=DayOf%></font></div>
 </td>
 <td bgcolor="<%=bgcolor%>">
  <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"><%=DayOf%></font></div>
 </td>
 <td bgcolor="<%=bgcolor%>">
  <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"><%=DayOf%></font></div>
 </td>
 </tr>
 <% next %>
</table>
</body>
</html>
```

