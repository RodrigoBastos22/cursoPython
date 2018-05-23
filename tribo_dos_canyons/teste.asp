<%

Dim sConnection

Dim oConnection

Dim oRS

sConnection = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=localhost; DATABASE=; UID=;PASSWORD=< Put your MySQL password here >; OPTION=3"

Set oConnection = Server.CreateObject("ADODB.Connection")

oConnection.Open(sConnection)

Set oRS = oConnection.Execute("SELECT * FROM Table1")

While Not oRS.EOF

Response.Write oRS ("Column1") & " " & oRS ("Column2") & "

"

oRS.MoveNext

Wend

oRS.Close

Set oRS = Nothing

oConnection.Close

Set oConnection = Nothing

%>

<H3>Enviando um e-mail com o componente AspEmail</H3>
<%

Dim Mail ' objeto Email

Dim strFromName ' nome do remetente

Dim strFromEmail ' endereço de Email do remetente

Dim strToEmail ' endereço do destinatario

Dim strSubject, strBody 'corpo da mensagem

Dim strThisPage ' o endereco do seu site

Dim strReferringPage ' a referencia URL

Dim bValidInput ' variável Booleana usada na validação

Dim strhost ' nome do servidor

' Retorna o nome do arquivo de script e a url da pagina

strThisPage = Request.ServerVariables("SCRIPT_NAME")

strReferringPage = Request.ServerVariables("HTTP_REFERER")

'define os valores iniciais dos parametros usados em nossa mensagem

strhost = "smtp.gmail.com.br" ‘altere o nome para o seu servidor de mensagens

strFromName = Trim(Request.Form("txtFromName"))

strFromEmail = Trim(Request.Form("txtFromEmail"))

strToEmail = Trim(Request.Form("txtToEmail"))

strSubject = "Site sobre Visual Basic"

strBody = Trim(Request.Form("txtMessage"))

'monta o corpo da mensagem

strBody = ""

strBody = strBody & "Achei um site que tem tudo sobre Visual Basic , dê uma olha em :" & vbCrLf

strBody = strBody & vbCrLf

strBody = strBody & " http://www.geocities.com/SiliconValley/Bay/3994 " & vbCrLf

' validacao dos dados

bValida_Entrada = True

bValida_Entrada = bValida_Entrada And strFromName <> ""

bValida_Entrada = bValida_Entrada And Valida_Email(strFromEmail)

bValida_Entrada = bValida_Entrada And Valida_Email(strToEmail)

'Se o e-mail é valido envia a mensagem

If bValida_Entrada Then

Set Mail = Server.CreateObject("Persits.MailSender")

Mail.Host = strHost

Mail.From = strFromEmail

Mail.FromName = strFromName

Mail.AddAddress strToEmail

Mail.Subject = strSubject

Mail.Body = strBody

on error resume next

Mail.Send

mensagem_erro = ""

if err <> 0 then

mensagem_erro = "Ocorreu o seguinte erro durante o envio do e-mail: " & Err.description

end if

Set Mail = Nothing

on error goto 0

' exibe mensagem de agradecimento

%>

<P><b>Sua mensagem foi enviada. Obrigado por ter visitado nosso site , volte sempre !</P></b>

<%

Else

If "http://" & Request.ServerVariables("HTTP_HOST") & strThisPage = strReferringPage Then

Response.Write "Ocorreu um erro . Verifique suas informações: " & "<BR>" & vbCrLf

End If

' exibe o formulario...

Exibe_Formulario strThisPage, strFromName, strFromEmail, strToEmail, strBody

End If

%>

<%

'verifica se o e-mail é valido

Function Valida_Email(strEmail)

Dim bIsValid

bIsValid = True

If Len(strEmail) < 5 Then

bIsValid = False

Else

If Instr(1, strEmail, " ") <> 0 Then

bIsValid = False

Else

If InStr(1, strEmail, "@", 1) < 2 Then

bIsValid = False

Else

If InStrRev(strEmail, ".") < InStr(1, strEmail, "@", 1) + 2 Then

bIsValid = False

End If

End If

End If

End If

Valida_Email = bIsValid

End Function

%>

<%

Sub Exibe_Formulario(strPageName, strFromName, strFromEmail, strToEmail, strBody)

%>

<html>

<body bgcolor=aqua>

<FORM ACTION="<%= strPageName %>" METHOD="post" name=frmReferral>

<TABLE BORDER="0">

<TR>

<TD VALIGN="top" ALIGN="right"><STRONG>Seu Nome:</STRONG></TD>

<TD><INPUT TYPE="text" NAME="txtFromName" VALUE="<%= strFromName %>" SIZE="30"></TD>

</TR>

<TR>

<TD VALIGN="top" ALIGN="right"><STRONG>E-mail do Remetente :</STRONG></TD>

<TD><INPUT TYPE="text" NAME="txtFromEmail" VALUE="<%= strFromEmail %>" SIZE="50"></TD>

</TR>

<TR>

<TD VALIGN="top" ALIGN="right"><STRONG>E-mail do destinatário:</STRONG></TD>

<TD><INPUT TYPE="text" NAME="txtToEmail" VALUE="<%= strToEmail %>" SIZE="50"></TD>

</TR>

<TR>

<TD VALIGN="top" ALIGN="right"><STRONG>Mensagem:</STRONG></TD>

<TD><TEXTAREA NAME="txtMessage" COLS="50" ROWS="5" WRAP="virtual" READONLY><%= strBody %></TEXTAREA></TR>

<TR>

<TD></TD>

<TD><INPUT TYPE="reset" VALUE="Limpar Formulário" name=rstReferral>&nbsp;&nbsp;<INPUT TYPE="submit" VALUE="Enviar E-mail" name=subReferral></TD>

</TR>

</TABLE>

</FORM>

</body>

</html>

<%

End Sub

%>
