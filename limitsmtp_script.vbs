
'------------------------------------------------------------------
' Global variables and settings
'------------------------------------------------------------------

'General
Public obApp
Public domain_buffer
Public Const ipslocalhost = "0.0.0.0"  'separated by #
Public Const user = "Administrator"
Public Const pw = "yourpw"
Public Const logspath = "C:\hMailServer\Logs\"   'ends with a backslash
Public Const write_log_active = true

'User and Domain outgoing limitation
Public Const outgoingstore = "C:\hMailServer\Events\outboundstore.txt"
Public Const outgoingexceptions = "C:\hMailServer\Events\outboundexceptions.txt"
Public Const outgoingstoreavg = "C:\hMailServer\Events\outboundstoreavg.txt"
Public Const max_emails_per_user = 100
Public Const max_emails_per_domain = 300
Public Const warning_factor = 0.8
Public Const server_average_days = 20     ' 0 will deactivate
Public Const server_average_threshold_factor = 10
Public Const warning_factor_avg = 0.6
Public Const msg_admin_warning = True
Public Const msg_admin_passed = True
Public Const msg_user_warning = True
Public Const msg_user_passed = True
Public Const msg_from = "Your Name <yourmail@yourdomain.com>"
Public Const msg_fromaddress = "yourmail@yourdomain.com"



'------------------------------------------------------------------
' Hmailserver Eventhandlers
'------------------------------------------------------------------

Sub OnAcceptMessage(oClient, oMessage)
	Result.Value = 0
	Set obApp = CreateObject("hMailServer.Application")
	Call obApp.Authenticate(user, pw)
	
	If has_client_authenticated(oClient) Then
		write_log ("  User has authenticated. User " & oCLient.username & ", Client " & oClient.IPAddress)
		if not check_outgoing_limitations(oClient, oMessage) Then
			Result.Message = "Your account/Mailserver has passed SMTP outgoing limits."
			Result.Value = 2
		End if
	End if
End Sub



'------------------------------------------------------------------
' SMTP limit outgoing emails of domain and user 
'------------------------------------------------------------------

function check_outgoing_limitations(oClient, oMessage)
	check_outgoing_limitations = true
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim fs , f
	Set fs = CreateObject("scripting.filesystemobject")
	Dim idt
	Dim content
	Dim ln
	Dim arr
	Dim usern
	Dim usernadd
	Dim usernnr
	Dim usernnrmax
	Dim domn
	Dim domnadd
	Dim domnnr
	Dim domnnrmax
	Dim reason
	Dim rcptscnt
	Dim dayamounts(200)
	For i = 0 To 200
		dayamounts(i) = 0
	Next 
	Dim pos
	Dim avg
	Dim minday
	minday = 999999
	Dim toindex
	Dim excptn
	
	write_log("  SMTP outgoing limitations")
	
	If oclient.username <> "" Then
		If instr(1,oclient.username,"@") = 0 Then
			usern = oclient.username & "@" & obApp.Settings.DefaultDomain
			domn = "@" & obApp.Settings.DefaultDomain
		Else
			usern = oclient.username
			domn = Mid(oclient.username,InStr(1,oclient.username,"@"))
		End If
	ElseIf is_local_domain(omessage.fromaddress) then
		usern = omessage.fromaddress
		domn = Mid(omessage.fromaddress,InStr(1,omessage.fromaddress,"@"))
	Else
		usern = "local"
		domn = "@local"
	End If
	content = "# SMTP outgoing storage" & nl & nl
	usernadd = true
	domnadd = true
	usernnr = 1
	domnnr = 1
	usernnrmax = max_emails_per_user
	domnnrmax = max_emails_per_domain
	idt = CLng(Date())
	rcptscnt = omessage.Recipients.count
	write_log("   Number of recipients " & rcptscnt)
	
	write_log("   Reading exceptions file " & outgoingexceptions)
	If fs.FileExists(outgoingexceptions) Then
		Set f = fs.OpenTextFile(outgoingexceptions, ForReading)
		Do While Not f.AtEndOfStream
			ln = f.ReadLine
			If ln <> "" And Mid(ln,1,1) <> "#" And Len(ln) > 3 Then
				arr = Split(ln,Chr(9))
				If UBound(arr) = 1 Then
					If arr(0) = usern Then
						usernnrmax = CLng(arr(1))
						write_log ("    new user limit " & ln)
					End if
					If arr(0) = domn Then
						domnnrmax = CLng(arr(1))
						write_log ("    new domain limit " & ln)
					End if
				Else
					write_log ("    cannot process line " & Mid(ln,1,25))
				End If
			ElseIf Len(ln) > 5 And f.Line > 4 + 1 then
				write_log ("    skipping line " & Mid(ln,1,25))
			End If
		Loop
	Else
		Set f = fs.OpenTextFile(outgoingexceptions, ForWriting, true)
		f.Write("# Outgoing limitation exceptions tab / chr(9) separated" & nl)
		f.Write("# Examples (without # at the beginning)" & nl)
		f.Write("# @yourdomain.com	10000" & nl)
		f.Write("# address@yourdomain.com	5000" & nl & nl)
		f.Close 
	End If
	
	write_log("   Reading storage file " & outgoingstore)
	If fs.FileExists(outgoingstore) Then
		Set f = fs.OpenTextFile(outgoingstore, ForReading)
		Do While Not f.AtEndOfStream
			ln = f.ReadLine
			If ln <> "" And Mid(ln,1,1) <> "#" And Len(ln) > 5 Then
				arr = Split(ln," ")
				If UBound(arr) > 1 Then
					If minday > CLng(arr(0)) Then
						minday = CLng(arr(0))
					End If
				End If
				If UBound(arr) = 2 Or UBound(arr) = 3 Then
					If CLng(arr(0)) = idt And arr(2) = usern Then
						usernnr = CLng(arr(1)) + rcptscnt
						usernadd = False
						write_log ("    adding to line " & ln)
						If usernnr > usernnrmax Then
							If UBound(arr) = 3 Then
								If arr(3) = "X" then
									write_log ("    deny already sent")
								Else
									write_log ("    sending deny")
									outgoing_limitations_send_message oClient, oMessage, false, usernnr, usernnrmax, false
								End if
							Else
								write_log ("    sending deny")
								outgoing_limitations_send_message oClient, oMessage, false, usernnr, usernnrmax, false
							End If
							content = content & arr(0) & " " & usernnr & " " & arr(2) & " X" & nl
						ElseIf usernnr > usernnrmax * warning_factor then
							If UBound(arr) = 3 Then
								If arr(3) = "W" then
									write_log ("    warning already sent")
								Else
									write_log ("    sending warning")
									outgoing_limitations_send_message oClient, oMessage, true, usernnr, usernnrmax, false
								End if
							Else
								write_log ("    sending warning")
								outgoing_limitations_send_message oClient, oMessage, true, usernnr, usernnrmax, false
							End If
							content = content & arr(0) & " " & usernnr & " " & arr(2) & " W" & nl
						Else
							content = content & arr(0) & " " & usernnr & " " & arr(2) & nl
						End if
					elseIf CLng(arr(0)) = idt And arr(2) = domn Then
						domnnr = CLng(arr(1)) + rcptscnt
						domnadd = false
						write_log ("    adding to line " & ln)
						If domnnr > domnnrmax Then
							If UBound(arr) = 3 Then
								If arr(3) = "X" then
									write_log ("    deny already sent")
								Else
									write_log ("    sending deny")
									outgoing_limitations_send_message oClient, oMessage, false, domnnr, domnnrmax, true
								End if
							Else
								write_log ("    sending deny")
								outgoing_limitations_send_message oClient, oMessage, false, domnnr, domnnrmax, true
							End If
							content = content & arr(0) & " " & domnnr & " " & arr(2) & " X" & nl
						ElseIf domnnr > domnnrmax * warning_factor then
							If UBound(arr) = 3 Then
								If arr(3) = "W" then
									write_log ("    warning already sent")
								Else
									write_log ("    sending warning")
									outgoing_limitations_send_message oClient, oMessage, true, domnnr, domnnrmax, true
								End if
							Else
								write_log ("    sending warning")
								outgoing_limitations_send_message oClient, oMessage, true, domnnr, domnnrmax, true
							End If
							content = content & arr(0) & " " & domnnr & " " & arr(2) & " W" & nl
						Else
							content = content & arr(0) & " " & domnnr & " " & arr(2) & nl
						End if
					ElseIf CLng(arr(0)) < idt - server_average_days Then
						write_log ("    deleting line " & ln)
					Else
						content = content & arr(0) & " " & arr(1) & " " & arr(2) & nl
						'write_log ("    copying line " & ln)
					End If
					If Mid(arr(2),1,1) <> "@" Then
						pos = idt - CLng(arr(0))
						dayamounts(pos) = dayamounts(pos) + CLng(arr(1))
					End If
				Else
					write_log ("    cannot process line " & Mid(ln,1,25))
				End If
			ElseIf Len(ln) > 5 And f.Line > 1 + 1 then
				write_log ("    skipping line " & Mid(ln,1,25))
			End If
		Loop
		f.Close
		If usernadd Then
			content = content & idt & " " & usernnr & " " & usern & nl
		End If
		If domnadd Then
			content = content & idt & " " & domnnr & " " & domn & nl
		End If
		Set f = fs.OpenTextFile(outgoingstore, ForWriting, true)
		f.Write(content)
		f.Close 
	Else
		content = content & idt & " " & usernnr & " " & usern & nl
		content = content & idt & " " & domnnr & " " & domn & nl
		Set f = fs.OpenTextFile(outgoingstore, ForWriting, true)
		f.Write(content)
		f.Close 
	End If
	
	toindex = idt - minday
	avg = CDbl(0)
	If toindex >=5 then
		For i = 1 To toindex
			avg = avg + CDbl(dayamounts(i))
		Next
		avg = CDbl(avg) / CDbl(toindex)
		write_log("   Statistic calculation over " & server_average_days & " days")
		write_log("     todays amount " & dayamounts(0) & "   average " & avg & "   maximum " & avg * server_average_threshold_factor)
		write_log("   Checking statistics")
	End If
	If toindex < 5 then
		write_log("   Statistic calculation is only done over at least 5 days. Available days: " & toindex)
	ElseIf avg < 5 Then
		write_log("     average below 5 mails per day, ignoring average statistic")
	ElseIf dayamounts(0) > avg * server_average_threshold_factor Then
		write_log("     todays amount has passed limit of " & avg * server_average_threshold_factor)
		outgoing_limitations_avg_send_admin dayamounts(0),avg * server_average_threshold_factor,false
		check_outgoing_limitations = False
	ElseIf dayamounts(0) > avg * server_average_threshold_factor * warning_factor_avg Then
		write_log("     todays amount has passed warning level of " & avg * server_average_threshold_factor * warning_factor_avg)
		outgoing_limitations_avg_send_admin dayamounts(0),avg * server_average_threshold_factor,true
		check_outgoing_limitations = False
	Else
		write_log("     within limits")
	End If
	
	write_log("   Checking limits")
	If usernnrmax < usernnr Then
		check_outgoing_limitations = false
		write_log("     max of user passed!")
	ElseIf domnnrmax < domnnr Then
		check_outgoing_limitations = false
		write_log("     max of domain passed!")
	Else
		write_log("     within limits")
	End If
	
	excptn = false
	If oMessage.FromAddress = emailadmin Then
		excptn = true
	Else
		For k = 0 To oMessage.recipients.count - 1
			If oMessage.recipients(k).OriginalAddress = emailadmin Then
				excptn = True
			End If
		Next
	End If
	If excptn = True Then
		write_log("   Mail from/to admin -> passes lock")
		check_outgoing_limitations = true
	End if
End function

Sub outgoing_limitations_avg_send_admin(nr, max, iswarning)
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim fs , f
	Set fs = CreateObject("scripting.filesystemobject")
	Dim txt
	Dim tmp
	Dim str
	Dim out
	Dim snd
	
	If iswarning Then
		tmp = "Warning: Todays outgoing emails will reach lock soon"
		
		txt = "Hello " & emailadmin & nl & nl
		txt = txt & "todays outgoing email will reach avg limit soon." & nl & nl
		txt = txt & "Current amount is " & nr & nl
		txt = txt & "Limit is " & max & nl & nl
		txt = txt & "Regards" & nl
		txt = txt & emailadmin
		
		str = "W" & CLng(Date())
	Else
		tmp = "Locked: Todays outgoing emails have passed avg limit"
		
		txt = "Hello " & emailadmin & nl & nl
		txt = txt & "todays outgoing email have passed avg limit." & nl & nl
		txt = txt & "Current amount is " & nr & nl
		txt = txt & "Limit is " & max & nl & nl
		txt = txt & "Regards" & nl
		txt = txt & emailadmin
		
		str = "X" & CLng(Date())
	End If
	
	snd = true
	If fs.FileExists(outgoingstoreavg) Then
		Set f = fs.OpenTextFile(outgoingstoreavg,ForReading)
		out = f.ReadAll
		f.Close
		If out = str Then
			snd = false
		End If
	End If
	
	If snd then
		Set nMessage = CreateObject("hMailServer.Message")
		nMessage.From = msg_from
		nMessage.FromAddress = msg_fromaddress
		nMessage.AddRecipient emailadmin, emailadmin
		nMessage.Subject = tmp
		nMessage.Body = txt
		nMessage.Save
		
		Set f = fs.OpenTextFile(outgoingstoreavg,ForWriting,True)
		f.Write(str)
		f.Close
	End If
End Sub

Sub outgoing_limitations_send_message(oClient, oMessage, iswarning, nr, max, isdomain)
	Dim txt
	Dim tmp
	If oclient.username <> "" then
		tmp = oclient.username
	Else
		tmp = oMessage.FromAddress
	End If
	If iswarning Then
		txt = "Hello " & tmp & nl & nl
		txt = txt & "you will soon reach your account limits." & nl & nl
		txt = txt & "Current amount is " & nr & nl
		txt = txt & "Limit is " & max & nl & nl
		If isdomain Then
			txt = txt & "This is a limit of the your domain." & nl & nl
		Else
			txt = txt & "This is a limit of the your account." & nl & nl
		End If
		txt = txt & "Regards" & nl
		txt = txt & emailadmin
		
		If msg_admin_warning then
		End If
		If msg_user_warning then
			Set nMessage = CreateObject("hMailServer.Message")
			nMessage.From = msg_from
			nMessage.FromAddress = msg_fromaddress
			nMessage.AddRecipient tmp, tmp
			nMessage.Subject = "Warning: Account limits will be reached soon"
			nMessage.Body = txt
			nMessage.Save
		End if
	Else
		txt = "Hello " & tmp & nl & nl
		txt = txt & "you have passed your account limits." & nl & nl
		txt = txt & "Current amount is " & nr & nl
		txt = txt & "Limit is " & max & nl & nl
		If isdomain Then
			txt = txt & "This is a limit of the your domain." & nl & nl
		Else
			txt = txt & "This is a limit of the your account." & nl & nl
		End If
		txt = txt & "Regards" & nl
		txt = txt & emailadmin
		
		If msg_admin_passed then
			Set nMessage = CreateObject("hMailServer.Message")
			nMessage.From = msg_from
			nMessage.FromAddress = msg_fromaddress
			nMessage.AddRecipient emailadmin, emailadmin
			nMessage.Subject = "Locked: Account limits passed"
			nMessage.Body = txt
			nMessage.Save
		End If
		If msg_user_passed then
			Set nMessage = CreateObject("hMailServer.Message")
			nMessage.From = msg_from
			nMessage.FromAddress = msg_fromaddress
			nMessage.AddRecipient tmp, tmp
			nMessage.Subject = "Locked: Account limits passed"
			nMessage.Body = txt
			nMessage.Save
		End if
	End If
End Sub

'------------------------------------------------------------------
' General functions of all scripts
'------------------------------------------------------------------

Sub write_log(txt)
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim fs
	Dim f
	
	If write_log_active then
		Set fs = CreateObject("scripting.filesystemobject")
		
		Dim fn
		Dim tmp 
		fn = logspath & "hmailserver_event_" & get_date & ".log"
		Set f = fs.opentextfile(fn, ForAppending, true)
		tmp = """" & FormatDateTime(Date + time,0) & """" & Chr(9) & """" & txt & """" & nl
		f.Write(tmp)
		f.Close
	End if
End Sub

Function get_date
	Dim tmp
	Dim erg
	tmp = Year(Date)
	erg = CStr(tmp)
	
	If Month(Date) < 10 Then
		tmp = "0" & Month(Date)
	Else
		tmp = Month(Date)
	End If
	erg = erg & "-" & tmp
	
	If day(Date) < 10 Then
		tmp = "0" & day(Date)
	Else
		tmp = day(Date)
	End If
	erg = erg & "-" & tmp
	
	get_date = erg
End Function

Function nl
	nl = Chr(13) & Chr(10)
End function

Function is_local_domain(domain_or_email)
	is_local_domain = False
	Dim domain
	Dim doms
	Dim alss
	Dim i
	Dim j
	
	If InStr(1,"  " & domain_or_email,"@") > 0 Then
		domain = Mid(domain_or_email, InStr(1,domain_or_email,"@") + 1)
	Else
		domain = domain_or_email
	End If
	
	If domain_buffer = "" then
		i = 0
		Set doms = obapp.Domains
		Do While i <= doms.Count - 1
			Set dom = doms.Item(i)
			domain_buffer = domain_buffer & "#" & dom.Name
			j = 0
			Set alss = dom.DomainAliases
			Do While j <= alss.Count - 1
				Set als = alss.item(j)
				domain_buffer = domain_buffer & "#" & als.AliasName
				j = j + 1
			Loop
			i = i + 1
		Loop
	End If
	
	If InStr(1, "  " & domain_buffer, domain) > 0 Then
		is_local_domain = True
	End If
End Function

Function has_client_authenticated(oclient)
	has_client_authenticated = false
	If oCLient.username <> "" Or InStr(1,"  " & ipslocalhost, oClient.IPAddress) > 0 Then
		has_client_authenticated = true
	End if
End Function

