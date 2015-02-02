<% OPTION EXPLICIT %>
<!-- #include file="include/connectdatabase.asp" -->
<!-- #include file="include/functions.asp" -->
<!-- #include file="include/index.asp" -->
<!-- #include file="include/beheer.asp" -->

<!-- #include file="include/indexform.asp" -->

<!-- #include file="include/indexheader.asp" -->
<%
IF Session("AantalFoutiveLogins")="" THEN
	Session("AantalFoutiveLogins")=0
END IF
Response.Buffer = True 'Buffers the content so Response.Redirect will work
dim Action
Action = request.QueryString("Action")
If Action = "" THEN
	Action = request.Form("Action")
End IF
IF Request.ServerVariables("REQUEST_METHOD") = "POST" THEN
	IF Action = "Login" THEN
		dim bLoginFout
		bLoginFout = false
		IF Request.form("remberme") ="true" THEN
			Response.Cookies("EmailAdresCookie") = Request.Form("email")
			Response.Cookies("RememberMeCookie") = "True"
			Response.Cookies("EmailAdresCookie").expires = Date() + 60
			Response.Cookies("RememberMeCookie").expires = Date() + 60
		Else
			Response.Cookies("RememberMeCookie") = "" 
			Response.Cookies("EmailAdresCookie") = "" 
		End If
		dim intCheckLogin
		intCheckLogin = CheckLogin(Replace(request.form("email"),"'","''"),Replace(request.form("password"),"'","''"))
		IF intCheckLogin = 0 THEN
			bLoginFout = true
			ToonLoginForm()
		ELSE
			IF intCheckLogin = 1 THEN
				'Als er gereageerd wordt op een link in een notificvatie dan naar het verslag gaan:
				IF request.form("LeesVerslagID")<>"" THEN
					response.redirect("home.asp?action=LeesVerslag&VerslagID="&request.form("LeesVerslagID"))
				ELSE
					IF isContactpersoon() OR isProfessional() THEN
						response.redirect("home.asp")
					END IF
					IF isBeheerder() THEN
						IF GetAantalClientenVanSchriftje() = 0 THEN
							response.redirect("aandeslag.asp")
						ELSE
							response.redirect("beheer.asp")
						END IF
					END IF
				END IF
			ELSE
				KiesSchriftjeForm()
			END IF
		END IF
	END IF
	IF Action ="KiesSchriftje" THEN
		IF CheckLoginOpEenScriftje(Replace(request.form("email"),"'","''"),Replace(request.form("password"),"'","''"),request.form("SchriftjeID")) = true Then
			'Als er gereageerd wordt op een link in een notificvatie dan naar het verslag gaan:
			IF request.form("LeesVerslagID")<>"" THEN
				response.redirect("home.asp?action=LeesVerslag&VerslagID="&request.form("LeesVerslagID"))
			ELSE
				IF isContactpersoon() OR isProfessional() THEN
					response.redirect("home.asp")
				END IF
				IF isBeheerder() THEN
					IF GetAantalClientenVanSchriftje() = 0 THEN
						response.redirect("aandeslag.asp")
					ELSE
						response.redirect("beheer.asp")
					END IF
				END IF
			END IF
		ELSE
			ToonLoginForm()
		END IF
	END IF
Else
	ToonLoginForm()
End If %>
<!-- #include file="include/indexfooter.asp" -->
