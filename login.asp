<!--#include file="assets/ConectaDB/conectaDB.asp"-->
<%
Cuser  = replace(request("emailCli"),"'","")
Csenha = replace(request("password"),"'","")

set bd=Server.CreateObject("ADODB.RecordSet")

strQuery = "SELECT * FROM usuarios u inner join franqueados f on  f.idfranqueado = u.idfranqueado " &_
          " where emailUsuario= '" & Cuser &"'" &_   
		  " and senhaUsu_10 =   '" & Csenha  &"'" &_
		  " and f.idfranqueado = " & session("idfranqueado") &";"
 response.write strQuery 
 response.end
'id_usuario	NomeUsuario	emailUsuario	senhaUsu_10	perfil_CK	idfranqueado	idGrupo	idfranqueado	NomeFranqueados	CgcFranqueado	dataIncFranq	statusFranq
'1	           Gilberto	skwcontadores@gmail.com	master12	SA	999	1	999	Curso Aprovar	99999999999999	2012-10-01 00:00:00.000	A

'bd.Open strQuery, strConexao  

'bd.CursorLocation = 2
'bd.CursorType = 0
'bd.LockType = 3
'bd.Open strQuery, objConn,,, &H0001

'response.write "nome " & bd("NomeUsuario")& "<br>"
'response.write "usu " & bd("usuario")     & "<br>"
'response.write "essenha " & bd("es_senha")& "<br>"
'response.end

set objConn = server.CreateObject("ADODB.Connection")
objConn.open strConexao
set bd=Server.CreateObject("ADODB.RecordSet")
on error resume next
bd.Open strQuery, strConexao

IF Err.Number <> 0 Then  
 call msg_modal_ret("Erro na consulta Login<br> "  & ADOErro(Err.Number)) 
End if
	
if bd.eof then
	%>
   <script> alert("Usuário não encontrado!");</script>
	<!--include FILE="Dbase/conectaDB_fechar.ASP"-->
	<%
	session("acesso_liberado") = ""
	response.redirect "erro.asp"
 else
	session("backgroundColor") = bd("backgroundColor")
	session("TextoColor")      = bd("TextoColor")

	'response.write "textoColor"& session("TextoColor") 
	'response.write "backgroundColor"& session("backgroundColor") 
	'response.End
%>
<style type="text/css">
* { margin:0;    padding:0; }
<%
htmlBackColor = "background:#333333;"
'response.write "backgroundColor::::::::::::::::::::::::" & session("backgroundColor")
BackColor = "background:#"&session("backgroundColor")&";"
textoColor = "color: #" &session("TextoColor")&";"

if session("backgroundColor") = "" then
 BackColor = "background:#474747;"
end if 
%>

html { <%=htmlBackColor%> }
body {
    margin:4px auto;
    width:100%;
    /*height:600px;*/
    <%=BackColor%>
    overflow:hidden;
}

.form label { 
	margin-left: 10px; 
	<%=textoColor%> 
    margin-top:0px;
    font-size:14px;
	}

</style>
<%

 
   session("acesso_liberado") = "*OK*"
   session("userPerfil")      = bd("perfil_CK")
   session("nomeUsuario")     = bd("NomeUsuario")
   session("$llogin0")        = bd("emailUsuario") ' o email e o login do usuario
   session("idFranqueado")    = bd("IDFranqueado") ' id franqueado
   session("NomeFranqueados") = bd("NomeFranqueados")
   session("id_usuario")      = bd("id_usuario")



 'response.write " acesso= "& session("acesso_liberado") 
 'response.write " perfil= "& session("userPerfil") 
 'response.write " nomeusu= "& session("emailUsuario") 
 'response.end

  dim strAlterarSenha, strLogin
  
  'strAlterarSenha = bd("AlterarSenha_CK")
   strLogin = bd("emailUsuario")
   'response.write "strLogin::::::" &  strLogin	& "<br><br>"
   'response.write " perfil= "& session("userPerfil") 
   'response.end
	
		%>
		<!--include FILE="dbase/conectaDB_fechar.ASP"-->
		<%   
		   if session("userPerfil") = "A"   then
			' email encontrado no cadastro
				session("$llogin0") =  bd("emailUsuario")
				session("conectadousu") = "@#@%A%@#@"
			%><script>
				  document.location.href = "principal.asp";
			</script><%
		   elseif session("userPerfil") = "SA"   then
			' email encontrado no cadastro
				session("$llogin0") =  bd("emailUsuario")
				session("conectadousu") = "@#@S%A%@#@"
			%><script>
				  document.location.href = "principal.asp";
			</script><%	
           elseif session("userPerfil") = "AL"  then
			' email encontrado no cadastro aluno
				session("$llogin0") =  bd("emailUsuario")
				session("conectadousu") = "@#@A%luno%@#@"
			%><script>
				  document.location.href = "principal.asp";
			</script><%
 				
		   elseif session("userPerfil") = "PR"  then
			' email encontrado no cadastro
				session("$llogin0") =  bd("emailUsuario")
				session("conectadousu") = "@#@Prof@#@"
			%><script>
				  document.location.href = "principal.asp";
			</script><%		
		   else
			   response.redirect "desconectar.asp"
		   end if

  

end if 



%>