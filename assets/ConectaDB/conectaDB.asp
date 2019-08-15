
<% 

' response.write Request.ServerVariables("SERVER_NAME")   & "<br>"

Set pc = CreateObject("Wscript.Network") 
    response.write "nome computador: " & pc.ComputerName 
	Nomecomputador = pc.ComputerName 
	
Set pc = nothing 

    'response.write " (currently connected as " 
    'response.write ucase(Request.ServerVariables("SERVER_NAME")) 
    'response.write " on " 
    'response.write "local_addr:::" & Request.ServerVariables("LOCAL_ADDR") & ")" 
	'response.end 

	
	
     NomeComputador = "WIN5110"

if  NomeComputador = "WIN5110" then   ' trabalho
'Provider=sqloledb;Data Source=SQL5016,1433;Initial Catalog=DB_A0CCEF_skwcontadores;User Id=DB_A0CCEF_skwcontadores_admin;Password=MASTER12GILBERTO;

			' Variáveis com os valores de sua base de dados. 
			' trabalho
			strDataSoure = "BRRJ-OLIVEIGC1\SQLTESTE" 'colocar aqui a localização de sua base de dados 
			strDataBase ="quiz" 'colocar aqui o nome do banco 
			strUser = "sa" 'colocar aqui o nome do usuário 
			strPWD = "Za9988!!@@" 'colocar aqui a senha 

			'Geramos a query SQL que irá acessar os dados na base de dados 
			'Conforme alteração 1 
			strQuery = "select * from cargos; " 

			'Montamos a string de conexão 
			strConexao = "Provider=sqloledb;Data Source=sql5016.smarterasp.net,1433;Initial Catalog=DB_A0CCEF_skwcontadores;User Id=DB_A0CCEF_skwcontadores_admin;Password=MASTER12GILBERTO;"
 
            'response.write   strConexao
			'response.end
			
			'Criando um objeto de conexão com a base de dados e executamos a Query 
			set objConn = server.CreateObject("ADODB.Connection")
			objConn.open strConexao 
			session("strConexao") = strConexao  
				
				'response.write "Conexao sql server "

			set bd=Server.CreateObject("ADODB.RecordSet")
			strQuery = " SELECT count(*) as contador FROM usuarios " 
					   
			response.write("teste_strqry: " & strQuery) 
			response.end

			msg = strQuery 
			'response.end
			'bd.Open strQuery, strConexao

			bd.CursorLocation = 2
			bd.CursorType = 0
			bd.LockType = 3
			bd.Open strQuery, objConn,,, &H0001

           If not bd.EOF Then 
		      count = bd("contador")
			  response.write "total: " & count & "<br>"
           end if  


				
END IF 		

	

if 1=2 then
 'response.write "00000000000000000000000000"
 ' Variáveis com os valores de sua base de dados. 
			'strDataSoure = "187.17.103.131" 'colocar aqui a localização de sua base de dados 
			strDataSoure = "187.17.103.131" 'colocar aqui a localização de sua base de dados 
			strDataBase ="quiz" 'colocar aqui o nome do banco 
			strUser = "sopecascar_1" 'colocar aqui o nome do usuário 
			strPWD = "master12" 'colocar aqui a senha 

			'Geramos a query SQL que irá acessar os dados na base de dados 
			'Conforme alteração 1 
			strQuery = "select * from cargos; " 

			'Montamos a string de conexão 
			strConexao = "driver=MySQL ODBC 3.51 Driver;SERVER=" & strDataSoure 
			strConexao = strConexao & "; DATABASE=" & strDataBase 
			strConexao = strConexao & ";UID="& strUser 
			strConexao = strConexao & " ;PWD="& strPWD 
			'Criando um objeto de conexão com a base de dados e executamos a Query 
			'set objConn = server.CreateObject("ADODB.Connection") 
			'objConn.open strConexao 
			
			

			Set objConn = Server.CreateObject("ADODB.Connection")
			 objConn.ConnectionString = "DSN=quiz2"
			 objConn.Open 'strConexao
end if

if  NomeComputador = "BRRJ-OLIVEIGC1" then   ' trabalho
			' Variáveis com os valores de sua base de dados. 
			' trabalho
			strDataSoure = "BRRJ-OLIVEIGC1\SQLTESTE" 'colocar aqui a localização de sua base de dados 
			strDataBase ="quiz" 'colocar aqui o nome do banco 
			strUser = "sa" 'colocar aqui o nome do usuário 
			strPWD = "Za9988!!@@" 'colocar aqui a senha 

			'Geramos a query SQL que irá acessar os dados na base de dados 
			'Conforme alteração 1 
			strQuery = "select * from cargos; " 

			'Montamos a string de conexão 
			strConexao = "Provider=SQLOLEDB.1;SERVER=" & strDataSoure 
			strConexao = strConexao & "; DATABASE=" & strDataBase 
			strConexao = strConexao & ";UID="& strUser 
			strConexao = strConexao & " ;PWD="& strPWD 
            'response.write   strConexao
			'response.end
			
			'Criando um objeto de conexão com a base de dados e executamos a Query 
			set objConn = server.CreateObject("ADODB.Connection")
			objConn.open strConexao 
			session("strConexao") = strConexao  
				
				'response.write "Conexao sql server "
END IF 		

if NomeComputador = "IIS2101"  then    ' DatabaseMArt
' Variáveis com os valores de sua base de dados. 
			strDataSoure = "sql2102.shared-servers.com,1087" 'colocar aqui a localização de sua base de dados 
			strDataBase ="quiz" 'colocar aqui o nome do banco 
			strUser = "gilberto" 'colocar aqui o nome do usuário 
			strPWD = "master12projetoq1968" 'colocar aqui a senha 

			'Geramos a query SQL que irá acessar os dados na base de dados 
			'Conforme alteração 1 
			strQuery = "select * from cargos; " 

			'Montamos a string de conexão 
			strConexao = "Provider=SQLOLEDB.1;SERVER=" & strDataSoure 
			strConexao = strConexao & "; DATABASE=" & strDataBase 
			strConexao = strConexao & ";UID="& strUser 
			strConexao = strConexao & " ;PWD="& strPWD 

			'Criando um objeto de conexão com a base de dados e executamos a Query 
			set objConn = server.CreateObject("ADODB.Connection") 
			objConn.open strConexao 
			session("strConexao") = strConexao  
				
				'response.write "Conexao sql server "
			
end if 

if  NomeComputador = "www.sopecascar.dominiotemporario.com" or NomeComputador="WHW0124" then	 ' uolhost			
			
			'response.write "Conexao My Sql "
			 set objConn = server.CreateObject("ADODB.Connection")
			 objConn.ConnectionString="driver=MySQL ODBC 3.51 Driver;server=187.17.103.131;uid=sopecascar_1;pwd=master12;database=sopecascar_1"
			 objConn.Open
			 
			 
end if 

if  NomeComputador = "GCO" then   ' Casa NetBook
			'response.write "GCO-NetBook"
			' Variáveis com os valores de sua base de dados. 
			' trabalho
			'strDataSoure = "BRRJ-OLIVEIGC1\SQLTESTE" 'colocar aqui a localização de sua base de dados 
		    ' netbook 
			strDataSoure = "GCO" 'colocar aqui a localização de sua base de dados 
			strDataBase ="quiz" 'colocar aqui o nome do banco 
			strUser = "sa" 'colocar aqui o nome do usuário 
			strPWD = "master12" 'colocar aqui a senha 

			'Geramos a query SQL que irá acessar os dados na base de dados 
			'Conforme alteração 1 
			strQuery = "select * from cargos; " 

			'Montamos a string de conexão 
			strConexao = "Provider=SQLOLEDB.1;SERVER=" & strDataSoure 
			strConexao = strConexao & "; DATABASE=" & strDataBase 
			strConexao = strConexao & ";UID="& strUser 
			strConexao = strConexao & " ;PWD="& strPWD 
            'response.write   strConexao
			'response.end
			
			'Criando um objeto de conexão com a base de dados e executamos a Query 
			set objConn = server.CreateObject("ADODB.Connection")
			objConn.open strConexao 
			session("strConexao") = strConexao  
				
				'response.write "Conexao sql server "
END IF 	


if  NomeComputador = "DESKTOP-2VF64SQ" then   ' netbook Sony-VAIO
			'response.write "DESKTOP-2VF64SQ"
			' Variáveis com os valores de sua base de dados. 
			' trabalho
			'strDataSoure = "BRRJ-OLIVEIGC1\SQLTESTE" 'colocar aqui a localização de sua base de dados 
		    ' netbook 
			' Provider=SQLNCLI10.1;Data Source=(local);Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=climb4acure;
			' Provider=SQLOLEDB.1;Data Source=GILBERTO-VAIO;User ID=sa;Initial Catalog=quiz;Persist Security Info=False;
			strDataSoure = "DESKTOP-2VF64SQ" 'colocar aqui a localização de sua base de dados 
			strDataBase ="quiz" 'colocar aqui o nome do banco 
			strUser = "sa" 'colocar aqui o nome do usuário 
			strPWD = "master12" 'colocar aqui a senha 

			'Geramos a query SQL que irá acessar os dados na base de dados 
			'Conforme alteração 1 
			strQuery = "select * from cargos; " 
	       
			'Montamos a string de conexão 
			strConexao = "Provider=SQLOLEDB.1;Data Source=" & strDataSoure 
			strConexao = strConexao & "; DATABASE=" & strDataBase 
			strConexao = strConexao & "; UID="& strUser 
			strConexao = strConexao & "; PWD="& strPWD   
			'strConexao = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=quiz;Data Source=GILBERTO-VAIO\SqlExpress/"

			response.write  strConexao
			response.end
			
			'Criando um objeto de conexão com a base de dados e executamos a Query 
			set objConn = server.CreateObject("ADODB.Connection")
			objConn.open strConexao 
			session("strConexao") = strConexao  
			'response.write "Conexao sql server "
END IF 	
 

%>