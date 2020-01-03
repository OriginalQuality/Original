<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>

<% 
vsoli = Request.Form("Solicitante")
vtele = Request.Form("Telefone")
vemai = Request.Form("Email")

' SAF 18/03/2010 as 12:29hs. 
vdest = "contato@oqcss.com"
'--------------------------------------

vassu = Request.Form("Assunto")
vdesc = Request.Form("Descricao")

If vsoli = "" OR vemai = "" OR vdesc = "" Then
	Response.redirect("email_erro.htm")
	Response.End
End If 

            //---------------------------
			// E-Mail para o SUPORTE.
            //---------------------------
			Set VS_Email_Suporte        = Server.CreateObject("Persits.MailSender")

			VS_Email_Suporte.Host       = "smtp.oqcss.com" 'Utilizar sempre esse endereco'
			VS_Email_Suporte.Username   = "contato@oqcss.com" 
			VS_Email_Suporte.Password   = "@Contato1" 
			VS_Email_Suporte.Port       = 587
			VS_Email_Suporte.MailFrom   = "contato@oqcss.com"


			VS_Email_Suporte.From       = VS_EmailSuporte  //' Endereco de EMail de quem está enviando o email.

			'  - Suporte da Original Quality.
			vsTitulo					= "Original Quality"
			VS_Email_Suporte.FromName   = "Original Quality - Corporativo"
			VS_Email_Suporte.AddAddress  vdest, "Original Quality "  //' Endereco de Email de quem vai receber o email.
			VS_Email_Suporte.Subject    =  "Fale Conosco - " & vassu

			// Monta o Corpo do EMAIL.
			MsgBody = "<!DOCTYPE HTML PUBLIC ""-//IETF//DTD HTML//EN"">" & NL 
			MsgBody = MsgBody & "<html>" 
			MsgBody = MsgBody & "<head>" 
			MsgBody = MsgBody & "<meta http-equiv=""Content-Type""" 
			MsgBody = MsgBody & "content=""text/html; charset=iso-8859-1"">" 
			MsgBody = MsgBody & "<title>Fale Conosco - Solicitação enviada pelo Cliente - Original Quality</title>"
			MsgBody = MsgBody & "<style type=""text/css"">"
			MsgBody = MsgBody & "<!--"
			MsgBody = MsgBody & ".style1 {"
			MsgBody = MsgBody &	"font-family: Arial, Helvetica, sans-serif;"
			MsgBody = MsgBody &	"font-size: 9px;"
			MsgBody = MsgBody & "}"
			MsgBody = MsgBody & "-->"
			MsgBody = MsgBody & "</style>"
			MsgBody = MsgBody & "</head>" 
			MsgBody = MsgBody & "<Body>"
			MsgBody = MsgBody & " <p>Solicitante</p>"
			MsgBody = MsgBody & " <p><strong>" & vsoli & "</strong></p>"
			MsgBody = MsgBody & " <p><strong>" & vtele & "</strong></p>"
			MsgBody = MsgBody & " <p><strong>" & vemai & "</strong></p><br>"
			MsgBody = MsgBody & " <p>A <strong>Original Quality</strong>, recebeu está solicitação pelo nosso Site Corporativo, sobre o assunto:</p>"
			MsgBody = MsgBody & " <p>( " & vassu & " ). </p>"
			MsgBody = MsgBody & " <p>Mais detalhes sobre este assunto, segue:</p>"
			MsgBody = MsgBody & " <p>( " & vdesc & " ). </p><br>"
			MsgBody = MsgBody & " <p>Análisar e retornar ao Cliente o mais breve possível!</p>"
			MsgBody = MsgBody & " <p>Obrigado,</p>"
			MsgBody = MsgBody & " <p><strong>Original Quality. Para ter qualidade, tem que ser original.</strong><br>"
			MsgBody = MsgBody & "   <span class=""style1"">Departamento Comercial <br>"
			MsgBody = MsgBody & "Telefones.......: 55 (11) 5611-4336<br>"
			MsgBody = MsgBody & "e-Mail..........: contato@oqcss.com<br>"
			MsgBody = MsgBody & "Loja Virtual ...: <a href=""http://www.oqcss.com"">http://www.oqcss.com</a><br>"
			MsgBody = MsgBody & "Site Corporativo: <a href=""http://www.oqcss.com"">http://www.oqcss.com</a><br>"
			MsgBody = MsgBody & "Facebook .......: <a href=""https://www.facebook.com/OriginalQualityCSS"">https://www.facebook.com/OriginalQualityCSS</a><br>"
			MsgBody = MsgBody & "Instagram ......: <a href=""https://www.instagram.com/originalquality_1"">https://www.instagram.com/originalquality_1</a><br>"
			MsgBody = MsgBody & "LinkeDIN .......: <a href=""http://linkedin.com/in/OriginalQuality"">http://linkedin.com/in/OriginalQuality</a><br>"
			MsgBody = MsgBody & "</span></p>"
			MsgBody = MsgBody & "</Body>" 
			MsgBody = MsgBody & "</html>" 

			VS_Email_Suporte.Body       = MsgBody
			VS_Email_Suporte.IsHTML     = True
			on error resume next
			VS_Email_Suporte.Send

			If Err <> 0 Then
				Set VS_Email_Suporte        = Nothing
				Response.redirect("email_erro02.htm")
				Response.End
			Else
			    Response.write "<font color='blue'><b>Mensagem enviada com sucesso para : </b></font> " & " [" & 	vdest & "] "
				Response.Write("<br>")
			End If

			Set VS_Email_Suporte        = Nothing


            //---------------------------
			// E-Mail para o CLIENTE.
            //---------------------------

			Set VS_Email_Suporte        = Server.CreateObject("Persits.MailSender")

			VS_Email_Suporte.Host       = "smtp.oqcss.com" 'Utilizar sempre esse endereco'
			VS_Email_Suporte.Username   = "contato@oqcss.com" 
			VS_Email_Suporte.Password   = "@Contato1" 
			VS_Email_Suporte.Port       = 587
			VS_Email_Suporte.MailFrom   = "contato@oqcss.com"


			
			VS_Email_Suporte.From       = "contato@oqcss.com" //' Endereco de EMail de quem está enviando o email.

			'  - Suporte da Original Quality.
			vsTitulo					= "Original Quality"
			VS_Email_Suporte.FromName   = "Original Quality - Corporativo"
			VS_Email_Suporte.AddAddress  vemai, "<" + vsoli + ">"  		//' Endereco de Email de quem vai receber o email.
			VS_Email_Suporte.Subject    =  "Fale Conosco - Confirmação de Recebimento (" + vassu + ")"

			// Monta o Corpo do EMAIL.
			MsgBody = "<!DOCTYPE HTML PUBLIC ""-//IETF//DTD HTML//EN"">" & NL 
			MsgBody = MsgBody & "<html>" 
			MsgBody = MsgBody & "<head>" 
			MsgBody = MsgBody & "<meta http-equiv=""Content-Type""" 
			MsgBody = MsgBody & "content=""text/html; charset=iso-8859-1"">" 
			MsgBody = MsgBody & "<title>Fale Conosco - Confirmação de recebimento de email comercial - Original Quality</title>"
			MsgBody = MsgBody & "<style type=""text/css"">"
			MsgBody = MsgBody & "<!--"
			MsgBody = MsgBody & ".style1 {"
			MsgBody = MsgBody &	"font-family: Arial, Helvetica, sans-serif;"
			MsgBody = MsgBody &	"font-size: 9px;"
			MsgBody = MsgBody & "}"
			MsgBody = MsgBody & "-->"
			MsgBody = MsgBody & "</style>"
			MsgBody = MsgBody & "</head>" 
			MsgBody = MsgBody & "<Body>"
			MsgBody = MsgBody & " <p>Att.</p>"
			MsgBody = MsgBody & " <p><strong>" & vsoli & "</strong></p>"
			MsgBody = MsgBody & " <p>A <strong>Original Quality</strong>, agradece o seu interesse, sobre o assunto:</p>"
			MsgBody = MsgBody & " <p>( " & vassu & " ). </p>"
			MsgBody = MsgBody & " <p>Sua solicitação está sendo analisada pelos nossos&nbsp;"
			MsgBody = MsgBody & "analistas comerciais.</p>"
			MsgBody = MsgBody & " <p>Em breve entraremos em contato  !</p>"
			MsgBody = MsgBody & " <p>Obrigado,</p>"
			MsgBody = MsgBody & " <p><strong>Original Quality. Para ter qualidade, tem que ser original.</strong><br>"
			MsgBody = MsgBody & "   <span class=""style1"">Departamento Comercial <br>"
			MsgBody = MsgBody & "Telefones.......: 55 (11) 5611-4336<br>"
			MsgBody = MsgBody & "e-Mail..........: contato@oqcss.com<br>"
			MsgBody = MsgBody & "Loja Virtual ...: <a href=""http://www.oqcss.com"">http://www.oqcss.com</a><br>"
			MsgBody = MsgBody & "Site Corporativo: <a href=""http://www.oqcss.com"">http://www.oqcss.com</a><br>"
			MsgBody = MsgBody & "Facebook .......: <a href=""https://www.facebook.com/OriginalQualityCSS"">https://www.facebook.com/OriginalQualityCSS</a><br>"
			MsgBody = MsgBody & "Instagram ......: <a href=""https://www.instagram.com/originalquality_1"">https://www.instagram.com/originalquality_1</a><br>"
			MsgBody = MsgBody & "LinkeDIN .......: <a href=""http://linkedin.com/in/OriginalQuality"">http://linkedin.com/in/OriginalQuality</a><br>"
			MsgBody = MsgBody & "</span></p>"
			MsgBody = MsgBody & "</Body>" 
			MsgBody = MsgBody & "</html>" 

			VS_Email_Suporte.Body       = MsgBody
			VS_Email_Suporte.IsHTML     = True
			on error resume next
			VS_Email_Suporte.Send

			If Err <> 0 Then
				Set VS_Email_Suporte        = Nothing
				Response.redirect("email_erro02.htm")
				Response.End
			Else
			    Response.write "<font color='blue'><b>Mensagem enviada com sucesso para : </b></font> " & " [" & 	vdest & "] "
				Response.Write("<br>")
			End If

			Set VS_Email_Suporte        = Nothing

Response.redirect("email_erro01.htm")
Response.End
%>