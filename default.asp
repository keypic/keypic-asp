<%

'  Copyright 2010-2011  Keypic LLC  (email : info@keypic.com)
'
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License, version 2, as 
'    published by the Free Software Foundation.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

%>

<% Response.clear %>

<!--#include file="include/keypic.asp"-->

<%
'on error resume next


'**************************************************************************
'* Check fields for post it inside my db
'**************************************************************************
name = request("name")
email = request("email")
message = request("message")

if request.servervariables("REQUEST_METHOD") = "POST" then
	if name<>"" and email<>"" and message<>"" then
		response.write "<font color=""red"">" & show_results() & "</font><br />"
		response.write session("img") & "<br />" & vbCrLf
		response.write "<a href="""">reload</a>"
		response.end
	else
		response.write "<font color=""red"">Complete all the fields</font><br />" & vbCrLf
	end if
end if

%>

<form method="post" action="http://<%=Request.ServerVariables("HTTP_HOST")%>">
Name: <br />
<input type="text" name="name" value="<%=name%>" /> <br />
Email: <br />
<input type="text" name="email" value="<%=email%>" /> <br />
Message: <br />
<textarea name="message" rows="5" cols="30"><%=message%></textarea> <br />
<%=session("img")%> <br />
<input type="hidden" name="token" value="<%=session("token")%>" /> <br />
<input type="submit" value="Send"> <br />
</form>
