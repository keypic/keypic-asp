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


session("token") = request("token")
session("host") = "http://ws.keypic.com/"
session("FormID") = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
'session("RequestType") = "RequestNewToken"
session("quantity") = "1"


if request.servervariables("REQUEST_METHOD") = "GET" then
	if session("token")="" then
		session("token") = spamtest("RequestNewToken")
	end if
end if


'********************************************************************
'* TEST FUNCTION
'********************************************************************
function spamtest(RequestType)

	if session("token")<>"" then
		strPost = strPost & "token=" & session("token") & "&"
	end if

	strPost = strPost & "FormID=" & session("FormID")  & "&" 'Customer ID
	strPost = strPost & "RequestType=" & RequestType  & "&"
	strPost = strPost & "ResponseType=4&" ' XML
	strPost = strPost & "ServerName=" & request.servervariables("SERVER_NAME")  & "&"
	strPost = strPost & "Quantity=" & session("quantity")  & "&"
	strPost = strPost & "ClientIP=" & request.servervariables("REMOTE_ADDR") & "&"
	strPost = strPost & "ClientUserAgent=" & request.servervariables("HTTP_USER_AGENT") & "&"
	strPost = strPost & "ClientAccept=" & request.servervariables("HTTP_ACCEPT") & "&"
	strPost = strPost & "ClientAcceptEncoding=" & request.servervariables("HTTP_ACCEPT_ENCODING") & "&"
	strPost = strPost & "ClientAcceptLanguage=" & request.servervariables("HTTP_ACCEPT_LANGUAGE") & "&"
	strPost = strPost & "ClientAcceptCharset=" & request.servervariables("HTTP_ACCEPT_CHARSET") & "&"
	strPost = strPost & "ClientUserAgent=" & request.servervariables("HTTP_USER_AGENT") & "&"
	strPost = strPost & "ClientHttpReferer=" & request.servervariables("HTTP_REFERER") & "&"
	strPost = strPost & "ClientUsername=" & ClientUsername & "&"
	strPost = strPost & "ClientEmailAddress=" & ClientEmailAddress & "&"
	strPost = strPost & "ClientMessage=" & ClientMessage & "&"
	strPost = strPost & "ClientFingerprint=" & ClientFingerprint & "&"

	set http = Server.CreateObject("Microsoft.XMLHTTP") ' CreateObject("MSXML2.ServerXMLHTTP") or "Microsoft.XMLHTTP"
	http.open "POST", session("host"), false
	http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" & vbCrLf 'multipart/form-data
	http.setRequestHeader "User-Agent", "ASP class / Version 0.5" & vbCrLf 
	'http.setRequestHeader "Content-Type", "multipart/form-data; boundary=" + Boundary

	if err.number = 0 then

		http.send strPost

		http_status = cint(http.status)

		if http_status = 200 then
			set xml = Server.CreateObject("MSXML2.DOMDocument") '("Microsoft.XMLDOM") or ("MSXML2.DOMDocument") or ("MSXML2.DOMDocument.3.0") or ("Msxml2.DOMDocument.4.0") or ("Msxml2.DOMDocument.5.0 ") or ("Msxml2.DOMDocument.6.0")
			xml.async = false
			xml.loadxml(http.responsetext)
			'response.write "http.responsetext = " & http.responsetext

			if xml.parseerror.errorcode <> 0 then
				response.write "XML Error: " & xml.parseerror.errorcode
			else
				dim xml_status
				set xml_status = xml.documentElement.selectSingleNode("//root/status")
				status = xml_status.text

				select case RequestType
					case "RequestNewToken"
						if status="new_token" then
							dim xml_token
							set xml_token = xml.documentElement.selectSingleNode("//root/Token")
							session("token") = xml_token.text
						end if
					case "RequestValidation"
						if status="response" then
							dim xml_spam
							set xml_spam = xml.documentElement.selectSingleNode("//root/spam")
							session("spam_percentage") = xml_spam.text
						end if
				end select
			end if

		else
			response.write "Error: " & status
		end if

		if RequestType = "RequestNewToken" and session("token")<>"" and status="new_token" then
			'response.write "token=" & token
			session("img") = "<a href=""" & session("host") & "?RequestType=getClick&amp;Token=" & session("token") & """ target=""_blank""><img src=""" & session("host") & "?RequestType=getImage&amp;Token=" & session("token") & """ alt="""" /></a>"
			spamtest = session("token")
		else
			'response.write "status=" & status
			spamtest = status
		end if

	else
		'response.write "err.number: " & err.number
		'In case the WASF is unreachable 
		spamtest = ""
	end if
		
end function


'********************************************************************
'* SHOW RESULTS FUNCTION
'********************************************************************
function show_results()
	if request.servervariables("REQUEST_METHOD") = "POST" then

			spamtest("RequestValidation")

			show_results = "This message has " & session("spam_percentage") & "% of spam probability<br />" & vbCrLf
	end if
end function


'********************************************************************
'* Version 1.0 Must be like this
'********************************************************************
'class Keypic
'
'	public sub setVersion(string version) ' return void
'	public sub setUserAgent(string UserAgent) ' return void
'	public sub setFormID(string FormID) ' return void
'	public sub checkFormID(string FormID) ' return bool
'	public sub setDebug(bool Debug) ' return void
'	private sub sendRequest(array fields) ' return array
'	public sub getToken(string Token, string ClientEmailAddress = "", string ClientUsername = "", string ClientMessage = "", string ClientFingerprint = "") ' return array
'	public sub getImage(string WeightHeight = null, string Debug = null) ' return string
'	public sub getiFrame(string WeightHeight = null) ' return String
'	public sub isSpam(string Token, string ClientEmailAddress = "", string ClientUsername = "", string ClientMessage = "", string ClientFingerprint = "") ' return int
'	public sub reportSpam(string Token) ' return bool
'
'end class
'
'Set kp = New Keypic
'
'kp.SetVersion = ""
'kp.SetUserAgent = ""
'kp.setFormID = ""


%>
