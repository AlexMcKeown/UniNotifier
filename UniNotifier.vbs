Main 'Calls function"

Function Main()
hitClick = false
uni_username = "uni_username" 
uni_password = "password"

steam_username = "steamUI"
steam_password = "passwordXD"

'CSIT113
TimeA = "https://solss.uow.edu.au/sid/sols_tutorial_enrolment.confirm_transfer?p_student_number=6133216&p_sub_inst_id_pri=305059&p_sub_inst_id_ori=305059&p_tut_id=223775&p_session_id=&p_cs=775286886296590054"
TimeB = "https://solss.uow.edu.au/sid/sols_tutorial_enrolment.confirm_transfer p_student_number=6133216&p_sub_inst_id_pri=305059&p_sub_inst_id_ori=305059&p_tut_id=223773&p_session_id=&p_cs=7752868864065423826"

set webbrowser = createobject("internetexplorer.application") 'Set up Browser"
webbrowser.statusbar = false
webbrowser.toolbar = false
webbrowser.visible = true

'First location -> UNI Login"
webbrowser.navigate("https://solss.uow.edu.au/sid/sols_login_ctl.login?p_message=SOLS_AUTH_ERR10")
wscript.sleep(1000) 'Wait for page to load (1 second)'
'Inputing user credentials"
webbrowser.document.all.item("p_username").value = uni_username 
webbrowser.document.all.item("p_password").value = uni_password
wscript.sleep(100) 'Wait for for input to load in"

Set oInupts = webbrowser.document.getElementsByTagName("input")'Look for the Login button as it does not have an id
For Each elm In oInupts
	If elm.Value = "Login" Then 'If the button which is an element has the Value == to 'Login' Then click on it 
	elm.click
	Exit For 'Exit for loop
	End If
Next 
wscript.sleep(900) 'Wait for browser to update
'Next location -> Tutorial Enrolment
webbrowser.navigate("https://solss.uow.edu.au/sid/sols_tutorial_enrolment.display_tutorial_timetable?p_sub_inst_id_pri=305015&p_sub_inst_id_ori=305015&p_type_id=10&p_tut_id=224930&p_student_number=6133216&p_session_id=&p_screen=tut_transf&p_cs=26064326383672067739")

DO WHILE webbrowser.busy 'While the web browser is loading wait for 0.1 Second
wscript.sleep(100)
LOOP 'End of Loop

canTransfer = False 'If the Html button attribute is True that means you can transfer tutorial times | false if it is hidden aka unable to transfer tutorial times '
alternate = False 'Swap Pages

	DO WHILE canTransfer = False 'If html button attribute is false we wait until we can see it thus making it true
		wscript.sleep(1000) 'Wait for the webpage to load and to act a bit human like
		 	If canTransfer = False Then
				wscript.sleep(30000) 'Every 3 seconds it will either look for Time A or Time B
				If alternate = False Then 'Time A'
					webbrowser.navigate(TimeA)
					'If we are directed to TimeA we should find a elm.Value = "Confirm Transfer'
					'If available we should be able to transfer to new subject time"
					Set oInupts = webbrowser.document.getElementsByTagName("input")'Look for the Login button as it does not have an id
					For Each elm In oInupts
						If elm.Value = "Confirm Transfer" Then 'If the button which is an element has the Value == to 'Login' Then click on it 
						canTransfer = True
						Exit For 'Exit for loop
						End If
					Next 
					alternate = true
				Else 'Time B
	        		webbrowser.navigate(TimeB)
					alternate = False
				End If
			 End If
	LOOP 'End of Loop"

'Jump to steam
webbrowser.navigate("https://store.steampowered.com/login/?redir=&redir_ssl=1&snr=1_4_4__global-header")
wscript.sleep(100) 

'Inputing user credentials"
webbrowser.document.all.item("username").value = steam_username 
webbrowser.document.all.item("password").value = steam_password
wscript.sleep(100) 
Set oInupts = webbrowser.document.getElementsByTagName("button")'Look for the Login button as it does not have an id
For Each elm In oInupts
	If elm.Value = "submit" Then 'If the button which is an element has the Value == to 'Login' Then click on it 
	elm.click
	Exit For 'Exit for loop
	End If
Next 
'Sends Steam Authentication Message to my phone and then i can go transfer to the course

MsgBox "Operation Complete" 'Reports findings on computer
End Function ' Ends function Main






