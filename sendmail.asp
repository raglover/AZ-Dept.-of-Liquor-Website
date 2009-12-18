<script language="VBScript" RUNAT="Server">
Dim hdwc_UploadRequest
Sub HDWCaptchaValidation
  If (Request.QueryString("hdwtest") = "captchainstalled") Then 
    Response.Write "Captcha verification code installed. Note: Version for uploads." 
    Response.End 
  End If
  If (Request.ServerVariables("REQUEST_METHOD") = "POST") OR (Request.QueryString.Count >= 4) then
      Dim SessionCAPTCHA, CheckCAPTCHA, HDWCaptchaBack, HDWname
  	SessionCAPTCHA = Trim(CStr(Session("HDWCAPTCHA")))
  	Session("CAPTCHA") = ""
  	HDWCaptchaBack = "contact.asp?hdwmsg=invalid"
  	If (HDWCaptchaBack = "") Then HDWCaptchaBack = Request.QueryString("hdwfail")
  	If (Len(SessionCAPTCHA) < 1) OR ((Request.Cookies("hdcaptcha") = "") AND (Request.QueryString("hdcaptcha") = "")) OR ((SessionCAPTCHA <> CStr(Request.Cookies("hdcaptcha"))) AND (SessionCAPTCHA <> CStr(Request.QueryString("hdcaptcha")))) Then
       If (InStr(1,Request.ServerVariables("CONTENT_TYPE"), "multipart/form-data", 1) > 0) Then
          Dim byteCount, RequestBin, keys, i, ssst
          Set hdwc_UploadRequest = CreateObject("Scripting.Dictionary")
          byteCount = Request.TotalBytes
          RequestBin = Request.BinaryRead(byteCount)
          Buildhdwc_UploadRequest  RequestBin
          keys = hdwc_UploadRequest.Keys
          For i = 0 To hdwc_UploadRequest.Count -1
             on error resume next
             ssst = CStr(hdwc_UploadRequest.Item(keys(i)).Item("Value"))
             Response.Cookies("hdw" & keys(i)) = ssst
          Next
          For Each HDWname in Request.Cookies
              On Error Resume Next
              If (Left(HDWname,3) = "hdw")  AND (hdwc_UploadRequest.Item(Right(HDWname,Len(HDWname)-3)).Item("Value") = "") Then
                   If (Err.Description = "") Then Response.Cookies(HDWname) = ""
              End If
          Next
       Else
          For Each HDWname in Request.QueryString
              Response.Cookies("hdw" & HDWname) = Request.QueryString(HDWname)
          Next
          For Each HDWname in Request.Form
              Response.Cookies("hdw" & HDWname) = Request.Form(HDWname)
          Next
          For Each HDWname in Request.Cookies
              On Error Resume Next
              If (Left(HDWname,3) = "hdw") AND (Request.QueryString(Right(HDWname,Len(HDWname)-3)) = "") AND (Request.Form(Right(HDWname,Len(HDWname)-3)) = "") Then
                   If (Err.Description = "") Then Response.Cookies(HDWname) = ""
              End If
          Next
       End If
  	    Response.Redirect HDWCaptchaBack
  	End If
      For Each HDWname in Request.Cookies
         If (Left(HDWname,3) = "hdw") Then
             Response.Cookies(HDWname) = ""
         End If
      Next
  End If
End Sub

Sub Buildhdwc_UploadRequest(RequestBin)
	Dim PosBeg, PosEnd, boundary, boundaryPos, Pos, Name, PosFile, PosBound, FileName, i, Value, gd, tmp
	Dim ContentType
	PosBeg = 1
	PosEnd = InstrB(PosBeg,RequestBin,hdwc_getByteString(chr(13)))
	boundary = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
	boundaryPos = InstrB(1,RequestBin,boundary)
	Do until (boundaryPos=InstrB(RequestBin,boundary & hdwc_getByteString("--")))
		Dim UploadControl
		Set UploadControl = CreateObject("Scripting.Dictionary")
		Pos = InstrB(BoundaryPos,RequestBin,hdwc_getByteString("Content-Disposition"))
		Pos = InstrB(Pos,RequestBin,hdwc_getByteString("name="))
		PosBeg = Pos+6
		PosEnd = InstrB(PosBeg,RequestBin,hdwc_getByteString(chr(34)))
		Name = hdwc_getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
		PosFile = InstrB(BoundaryPos,RequestBin,hdwc_getByteString("filename="))
		PosBound = InstrB(PosEnd,RequestBin,boundary)
		If  PosFile<>0 AND (PosFile<PosBound) Then
			PosBeg = PosFile + 10
			PosEnd =  InstrB(PosBeg,RequestBin,hdwc_getByteString(chr(34)))
			FileName = hdwc_getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			UploadControl.Add "FileName", CleanFileName(FileName)
			Pos = InstrB(PosEnd,RequestBin,hdwc_getByteString("Content-Type:"))
			If Pos = 0 Then
			  PosEnd = PosEnd+1
			Else
			  PosBeg = Pos+14
			  PosEnd = InstrB(PosBeg,RequestBin,hdwc_getByteString(chr(13)))
			  ContentType = hdwc_getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			  UploadControl.Add "ContentType",ContentType
			End If
			PosBeg = PosEnd+4
			tmp = PosBeg-1
			PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
			Value = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
			Set gd = CreateObject("ADODB.Stream")
            gd.Type = 1
            gd.Open
            gd.Write RequestBin
            gd.Flush
            gd.Position=tmp
            Set globalbuffer = CreateObject("ADODB.Stream")
            globalbuffer.Type = 1
            globalbuffer.Open
            On Error Resume Next
            globalbuffer.Write gd.Read(PosEnd-PosBeg)
            globalbuffer.Flush
            globalbuffer.Position=0
			UploadControl.Add "Value" , globalbuffer
		Else
			Pos = InstrB(Pos,RequestBin,hdwc_getByteString(chr(13)))
			PosBeg = Pos+4
			PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
			Value = hdwc_getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			UploadControl.Add "Value" , Value
		End If
	If Not hdwc_UploadRequest.Exists(name) Then
     	    hdwc_UploadRequest.Add name, UploadControl
     	Else
     	    hdwc_UploadRequest.Item(name).Item("Value") = hdwc_UploadRequest.Item(name).Item("Value") + ";"+Value
     	End If
		BoundaryPos=InstrB(BoundaryPos+LenB(boundary),RequestBin,boundary)
	Loop
End Sub

Function hdwc_getByteString(StringStr)
 Dim i, char
 For i = 1 to Len(StringStr)
 	char = Mid(StringStr,i,1)
	hdwc_getByteString = hdwc_getByteString & chrB(AscB(char))
 Next
End Function

Function hdwc_getString(StringBin)
 Dim intCount
 hdwc_getString =""
 For intCount = 1 to LenB(StringBin)
	hdwc_getString = hdwc_getString & chr(AscB(MidB(StringBin,intCount,1)))
 Next
End Function

Function CleanFileName (fname)
  CleanFileName = Right(fname, Len(fname) - InStrRev(fname,""))
End Function

</script>
<% HDWCaptchaValidation %>
<%


'---------------------------------------------------------------------------------------------------
'FORM MAIL SCRIPT
'----------------
'usage:
'<form ACTION="sendmail.asp" ...>
'
'hidden fields:
'	redirect	- the url to redirect to when the mail has been sent (REQUIRED)
'	mailto		- the email address of the recipient (separate multiple recipients with commas)  (REQUIRED)
'	cc			- the email address of the cc recipient (separate multiple recipients with commas) (OPTIONAL)
'	bcc			- the email address of the bcc recipient (separate multiple recipients with commas) (OPTIONAL)
'	mailfrom	- the email address of the sender  (REQUIRED)
'	subject		- the subject line of the email  (REQUIRED)
'	message		- the message to include in the email above the field values. not used when a template is being used. (OPTIONAL)
'	template	- specifies a text or html file to use as the email template, relative to the location of the sendmail script. (e.g. ../email.txt)
'				  A template should reference form fields like this: [$Field Name$]
'	html		- if this has the value "yes", the email will be sent as an html email. only used if a template is supplied.
'	testmode	- if this is set to "yes", the email contents will be written to the screen instead of being emailed.
'---------------------------------------------------------------------------------------------------
dim pde : set pde = createobject("scripting.dictionary")
'---------------------------------------------------------------------------------------------------
'PREDEFINED ADDRESSES for the "mailto" hidden field
'if you don't want to reveal email addresses in hidden fields, use a token word instead and specify
'below which email address it applies to. e.g. <input type="hidden" name="mailto" value="%stratdepartment%">
'ALSO, in the same way, you can use %mailfrom% to hide the originating email address
pde.add "%contactform%", "myemail@someaddress.com"
pde.add "%salesenquiry%", "anotheremail@someaddress.com"
'---------------------------------------------------------------------------------------------------

function getTextFromFile(path)
	dim fso, f, txt
	set fso = createobject("Scripting.FileSystemObject")
	if not fso.fileexists(path) then
		getTextFromFile = ""
		exit function
	end if
	set f = fso.opentextfile(path,1)
	if f.atendofstream then txt = "" else txt = f.readall
	f.close
	set f = nothing
	set fso = nothing
	getTextFromFile = txt
end function

dim redir, mailto, mailfrom, subject, item, body, cc, bcc, message, html, template, usetemplate, testmode
redir = request.form("redirect")
mailto = request.form("mailto")
if pde.exists(mailto) then mailto = pde(mailto)
cc = request.form("cc")
bcc = request.form("bcc")
mailfrom = request.form("mailfrom")
if mailfrom = "" then mailfrom = pde("%mailfrom%")
subject = request.form("subject")
message = request.form("message")
template = request.form("template")
testmode = lcase(request.form("testmode"))="yes"

if len(template) > 0 then template = getTextFromFile(server.mappath(template))
if len(template) > 0 then usetemplate = true else usetemplate = false
%>
<!-- 
    METADATA 
    TYPE="typelib" 
    UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  
    NAME="CDO for Windows 2000 Library" 
--> 
<%
Dim SMTPServer, SMTPPort, SendUsingMethod
SMTPServer = "azmtaprd01.az.gov"
SMTPPort = 25
SendUsingMethod = 2		

Const SendMail__cdoSMTPServer = "http://schemas.microsoft.com/cdo/configuration/smtpserver"
Const SendMail__cdoSendUsingMethod = "http://schemas.microsoft.com/cdo/configuration/sendusing"  
Const SendMail__cdoSMTPPort = "http://schemas.microsoft.com/cdo/configuration/smtpserverport"

dim cdoConfig : set cdoConfig = server.CreateObject("CDO.Configuration")  
 
With cdoConfig.Fields  
	.Item(SendMail__cdoSendUsingMethod) = SendUsingMethod
	.Item(SendMail__cdoSMTPServer) = SMTPServer
	.Item(SendMail__cdoSMTPPort) = SMTPPort
	.Update
End With

dim msg : set msg = server.createobject("CDO.Message")

Set msg.Configuration = cdoConfig

msg.subject = subject
msg.to = mailto
msg.from = mailfrom
if len(cc) > 0 then msg.cc = cc
if len(bcc) > 0 then msg.bcc = bcc

if not usetemplate then
	body = body & message & vbcrlf & vbcrlf
else
	body = template
end if
for each item in request.form
	select case item
		case "redirect", "mailto", "cc", "bcc", "subject", "message", "template", "html", "testmode"
		case else
			if not usetemplate then
				if item <> "mailfrom" then body = body & item & ": " & request.form(item) & vbcrlf & vbcrlf
			else
				body = replace(body, "[$" & item & "$]", replace(request.form(item),vbcrlf,"<br>"))
			end if
	end select
next

if usetemplate then 'remove any leftover placeholders
	dim rx : set rx = new regexp
	rx.pattern = "\[\$.*\$\]"
	rx.global = true
	body = rx.replace(body, "")
end if

if usetemplate and lcase(request.form("html")) = "yes" then
	msg.htmlbody = body
else
	msg.textbody = body
end if
if testmode then
	if lcase(request.form("html")) = "yes" then
		response.write "<pre>" & vbcrlf
		response.write "Mail to: " & mailto & vbcrlf
		response.write "Mail from: " & mailfrom & vbcrlf
		if len(cc) > 0 then response.write "Cc: " & cc & vbcrlf
		if len(bcc) > 0 then response.write "Bcc: " & bcc & vbcrlf
		response.write "Subject: " & subject & vbcrlf & string(80,"-") & "</pre>"
		response.write body
	else
		response.write "<html><head><title>Sendmail.asp Test Mode</title></head><body><pre>" & vbcrlf
		response.write "Mail to: " & mailto & vbcrlf
		response.write "Mail from: " & mailfrom & vbcrlf
		if len(cc) > 0 then response.write "Cc: " & cc & vbcrlf
		if len(bcc) > 0 then response.write "Bcc: " & bcc & vbcrlf
		response.write "Subject: " & subject & vbcrlf & vbcrlf
		response.write string(80,"-") & vbcrlf & vbcrlf & "<span style=""color:blue;"">"
		response.write body & "</span>" & vbcrlf & vbcrlf
		response.write string(80,"-") & vbcrlf & "**END OF EMAIL**</pre></body></html>"
	end if
else
	msg.send
	response.redirect redir
end if
set msg = nothing
%>
