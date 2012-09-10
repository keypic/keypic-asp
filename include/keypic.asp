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

Response.Charset="UTF-8"
Session.CodePage=65001

class Keypic

	private prToken
	Public Property Get Token
		Token = prToken
	End Property
	Public Property Let Token(strToken)
		prToken = strToken
	End Property

	private prHost
	Public Property Get Host
		Host = prHost
	End Property
	Public Property Let Host(strHost)
		prHost = strHost
	End Property

	private prFormID
	Public Property Get FormID
		FormID = prFormID
	End Property
	Public Property Let FormID(strFormID)
		prFormID = strFormID
	End Property

	private prVersion
	Public Property Get Version
		Version = prVersion
	End Property
	Public Property Let Version(strVersion)
		prVersion = strVersion
	End Property

	private prUserAgent
	Public Property Get UserAgent
		UserAgent = prUserAgent
	End Property
	Public Property Let UserAgent(strUserAgent)
		prUserAgent = strUserAgent
	End Property

	private prDebug
	Public Property Get Debug
		Debug = prDebug
	End Property
	Public Property Let Debug(strDebug)
		prDebug = strDebug
	End Property

	Private Sub Class_Initialize()
		Token = ""
		Host = "http://ws.keypic.com/"
		FormID = ""
		Version = "1.0"
		UserAgent = "ASP class / Version " & Version
		Debug = ""
	End Sub

	Private Sub Class_Terminate()
	End Sub

	public function setVersion(strVersion) ' return void
		prVersion = strVersion
	end function

	public function setUserAgent(strUserAgent) ' return void
		prUserAgent = strUserAgent
	end function

	public function setFormID(strFormID) ' return void
		prFormID = strFormID
	end function

	public function checkFormID(FormID) ' return bool
		Set fields = Server.CreateObject("Scripting.Dictionary")
		fields.add "RequestType", "checkFormID"
		fields.add "ResponseType", "4" ' xml
		fields.add "FormID", FormID

'		checkFormID = sendRequest(fields) ' TODO: Finish it

		set xml = Server.CreateObject("MSXML2.DOMDocument")
		xml.async = false
		xml.loadxml(sendRequest(fields))

		dim xml_status
		set xml_status = xml.documentElement.selectSingleNode("//root/status")
		status = xml_status.text

		if status = "response" then
			set xml_report = xml.documentElement.selectSingleNode("//root/report")
			if xml_report.text = "OK" then
				checkFormID = true
			end if
		elseif status = "error" then
'			set xml_report = xml.documentElement.selectSingleNode("//root/error")
'			spamPercentage = xml_spam.text
			checkFormID = false
		end if

	end function

	public function setDebug(strDebug) ' return void
		prDebug = strDebug
	end function

	private function sendRequest(fields) ' return array

		for each key in fields.Keys
			strPost = strPost & key & "=" & fields(key) & "&"
		next
'response.write strPost
		set http = Server.CreateObject("Microsoft.XMLHTTP") ' CreateObject("MSXML2.ServerXMLHTTP") or "Microsoft.XMLHTTP"
		http.open "POST", Host, false
		http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" & vbCrLf 'multipart/form-data
		http.setRequestHeader "User-Agent", UserAgent & vbCrLf 
		'http.setRequestHeader "Content-Type", "multipart/form-data; boundary=" + Boundary

		if err.number = 0 then

			http.send strPost

			http_status = cint(http.status)

			if http_status = 200 then
				set xml = Server.CreateObject("MSXML2.DOMDocument") '("Microsoft.XMLDOM") or ("MSXML2.DOMDocument") or ("MSXML2.DOMDocument.3.0") or ("Msxml2.DOMDocument.4.0") or ("Msxml2.DOMDocument.5.0 ") or ("Msxml2.DOMDocument.6.0")
				xml.async = false
				xml.loadxml(http.responsetext)

				if xml.parseerror.errorcode <> 0 then
					sendRequest = "XML Error: " & xml.parseerror.errorcode
				'else
					
				end if

			else
				sendRequest = "HTTP Error number: " & http_status
			end if
		else
			'response.write "err.number: " & err.number
			'In case the WASF is unreachable 
			sendRequest = "err.number: " & err.number
		end if

		sendRequest = http.responsetext

	end function

	public function getToken(localToken, ClientEmailAddress, ClientUsername, ClientMessage, ClientFingerprint) ' return array
		Token = localToken
		Set fields = Server.CreateObject("Scripting.Dictionary")

		if Token <> "" then
			fields.add "token", Token
		end if

		fields.add "FormID", FormID
		fields.add "RequestType", "RequestNewToken"
		fields.add "ResponseType", 4
		fields.add "ServerName", request.servervariables("SERVER_NAME")
		fields.add "Quantity", 1
		fields.add "ClientIP", request.servervariables("REMOTE_ADDR")
		fields.add "ClientUserAgent", request.servervariables("HTTP_USER_AGENT")
		fields.add "ClientAccept", request.servervariables("HTTP_ACCEPT")
		fields.add "ClientAcceptEncoding", request.servervariables("HTTP_ACCEPT_ENCODING")
		fields.add "ClientAcceptLanguage", request.servervariables("HTTP_ACCEPT_LANGUAGE")
		fields.add "ClientAcceptCharset", request.servervariables("HTTP_ACCEPT_CHARSET")
		fields.add "ClientHttpReferer", request.servervariables("HTTP_REFERER")
		fields.add "ClientUsername", ClientUsername
		fields.add "ClientEmailAddress", ClientEmailAddress
		fields.add "ClientMessage", ClientMessage
		fields.add "ClientFingerprint", ClientFingerprint

		set xml = Server.CreateObject("MSXML2.DOMDocument")
		xml.async = false
		xml.loadxml(sendRequest(fields))

		dim xml_status
		set xml_status = xml.documentElement.selectSingleNode("//root/status")
		status = xml_status.text

		if status = "new_token" then
			dim xml_token
			set xml_token = xml.documentElement.selectSingleNode("//root/Token")
			Token = xml_token.text
		elseif status = "error" then
			getToken = "error"
		end if

		getToken = Token
	end function

	public function getImage(WeightHeight) ' return string
		getImage = "<a href=""" & Host & "?RequestType=getClick&amp;Token=" & Token & """ target=""_blank""><img src=""" & Host & "?RequestType=getImage&amp;WeightHeight=" & WeightHeight & "&amp;Debug=" & Debug & "&amp;Token=" & Token & """ alt=""Form protected by Keypic"" /></a>"
	end function

	public function getiFrame(WeightHeight) ' return String
	end function

	public function isSpam(localToken, ClientEmailAddress, ClientUsername, ClientMessage, ClientFingerprint) ' return int
		Token = localToken

		Set fields = Server.CreateObject("Scripting.Dictionary")

		if Token <> "" then
			fields.add "token", Token
		end if

		fields.add "FormID", FormID
		fields.add "RequestType", "RequestValidation"
		fields.add "ResponseType", 4
		fields.add "ServerName", request.servervariables("SERVER_NAME")
		fields.add "Quantity", 1
		fields.add "ClientIP", request.servervariables("REMOTE_ADDR")
		fields.add "ClientUserAgent", request.servervariables("HTTP_USER_AGENT")
		fields.add "ClientAccept", request.servervariables("HTTP_ACCEPT")
		fields.add "ClientAcceptEncoding", request.servervariables("HTTP_ACCEPT_ENCODING")
		fields.add "ClientAcceptLanguage", request.servervariables("HTTP_ACCEPT_LANGUAGE")
		fields.add "ClientAcceptCharset", request.servervariables("HTTP_ACCEPT_CHARSET")
		fields.add "ClientHttpReferer", request.servervariables("HTTP_REFERER")
		fields.add "ClientUsername", ClientUsername
		fields.add "ClientEmailAddress", ClientEmailAddress
		fields.add "ClientMessage", ClientMessage
		fields.add "ClientFingerprint", ClientFingerprint


		set xml = Server.CreateObject("MSXML2.DOMDocument")
		xml.async = false
		xml.loadxml(sendRequest(fields))

		dim xml_status
		set xml_status = xml.documentElement.selectSingleNode("//root/status")
		status = xml_status.text

		dim xml_spam
		if status = "response" then
			set xml_spam = xml.documentElement.selectSingleNode("//root/spam")
			spamPercentage = xml_spam.text
		elseif status = "error" then
			set xml_spam = xml.documentElement.selectSingleNode("//root/error")
			spamPercentage = xml_spam.text
		end if

		isSpam = spamPercentage
	end function

	public function reportSpam(Token) ' return bool
		if Token = "" then
			reportSpam "error"
		end if

		if FormID = "" then
			reportSpam "error"
		end if

		Set fields = Server.CreateObject("Scripting.Dictionary")

		fields.add "Token", Token
		fields.add "FormID", FormID
		fields.add "RequestType", "ReportSpam"
		fields.add "ResponseType", 4

		set xml = Server.CreateObject("MSXML2.DOMDocument")
		xml.async = false
		xml.loadxml(sendRequest(fields))

		dim xml_status
		set xml_status = xml.documentElement.selectSingleNode("//root/status")
		status = xml_status.text

' TODO: finish it!
'		if status = "response" then
'			set xml_spam = xml.documentElement.selectSingleNode("//root/spam")
'			spamPercentage = xml_spam.text
'		elseif status = "error" then
'			set xml_spam = xml.documentElement.selectSingleNode("//root/error")
'			spamPercentage = xml_spam.text
'		end if
		reportSpam = "TODO"
	end function

end class

%>
