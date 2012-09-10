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
<!--#include file="include/keypic.asp"-->
<%
'on error resume next
Response.clear
Set kp = New Keypic


kp.SetFormID("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")
'kp.setDebug("true")

'**************************************************************************
'* Check fields for post it inside my db
'**************************************************************************
name = request("name")
email = request("email")
message = request("message")
Token = request("Token")

if request.servervariables("REQUEST_METHOD") = "POST" then
	if name<>"" and email<>"" and message<>"" then
		spamPercentage = kp.isSpam(Token, email, name, message, "")
		color = "red"
		if(IsNumeric(spamPercentage)) then
			if(spamPercentage >= 70) then
				color = "red"
			elseif(spamPercentage >= 40) then
				color = "orange"
			else
				color = "green"
			end if
			response.write "<font color=""" & color & """>This message has " & spamPercentage & "% of spam probability</font><br />"			
		else
			response.write "<font color=""" & color & """>There was some problem with this request</font><br />"
		end if

		response.write kp.getImage("") & "<br />" & vbCrLf
		response.write "<a href="""">reload</a>"
		response.end
	else
		response.write "<font color=""" & color & """>Complete all the fields</font><br />" & vbCrLf
	end if
end if

%>

<form method="post" action="">
Name: <br />
<input type="text" name="name" value="<%=name%>" /> <br />
Email: <br />
<input type="text" name="email" value="<%=email%>" /> <br />
Message: <br />
<textarea name="message" rows="5" cols="30"><%=message%></textarea> <br />
<input type="hidden" name="token" value="<%=kp.getToken(Token, email, name, message, "")%>" /> <br />
<%=kp.getImage("")%> <br />
<input type="submit" value="Send"> <br />
</form>

<%

'Response.Write "Default parameters class<br />"
'Response.Write "Token: " & kp.Token & "<br />"
'Response.Write "Host: " & kp.Host & "<br />"
'Response.Write "FormID: " & kp.FormID & "<br />"
'Response.Write "Version: " & kp.Version & "<br />"
'Response.Write "UserAgent: " & kp.UserAgent & "<br />"

'Response.Write "***********************************************************<br />"
'Response.Write "* checkFormID <br />"
'Response.Write "***********************************************************<br />"
'
'Response.Write "checkFormID: " & kp.checkFormID("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx") & "<br />"


%>
