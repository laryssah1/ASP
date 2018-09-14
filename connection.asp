<!DOCTYPE html>
<html>
<body>

<%	'Cria conexão'
	Set objconn = Server.CreateObject("ADODB.Connection")
	'Seta a base de dados'
	objconn.open = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=x; DATABASE=y; UID=user;PASSWORD=passw; PORT= z;"
	'Recebe a query'
	instrucao = "UPDATE x  INNER JOIN y ON x.a = y.a set x.status = 0 WHERE z NOT BETWEEN CURDATE()-INTERVAL 90 DAY AND CURDATE();"
	'Executa a query'
	set dados = objconn.execute(instrucao)

	response.write "OK!"
%>
</body>
</html>


