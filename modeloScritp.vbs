Option Explicit ' Forces us to declare all variables

Dim app         ' application
Dim project     ' Project object
Dim sasProgram  ' Code object (SAS program)
Dim n           ' counter

Dim ObjFso
Dim ObjFile

Set ObjFso = CreateObject("Scripting.FileSystemObject")
Set ObjFile = ObjFso.OpenTextFile("[local_arquivo_sas]",1)
 Dim MyVar
'MyVar = MsgBox (ObjFile.ReadAll, 65, "MsgBox Example")
Dim conteudo
conteudo = ObjFile.ReadAll

' Use SASEGObjectModel.Application.4.2 for EG 4.2
Set app = CreateObject("SASEGObjectModel.Application.5.1")
' Set to your metadata profile name, or "Null Provider" for just Local server
app.SetActiveProfile("sas2")
' Create a new project
Set project = app.New 
' add a new code object to the project
Set sasProgram = project.CodeCollection.Add
 
' set the results types, overriding app defaults
sasProgram.UseApplicationOptions = False
sasProgram.GenListing = True
sasProgram.GenSasReport = False
  
' Set the server (by Name) and text for the code
sasProgram.Server = "SGERAIS"
sasProgram.Text = conteudo '"options notes; ODS _ALL_ CLOSE; libname cluster odbc noprompt = 'server=10.207.44.120,1433;DRIVER=SQL Server;uid=sa;pwd=l7l4l1L2@@;Trusted_Connection=yes;Database=MAPFRE;schema=dbo;' INSERTBUFF=15000 dbcommit=100000 ;    %macro consulta (prod, cia, tipo_dv , tip_nivel);  proc sql ; CONNECT TO ODBC as con1 (DATASRC=CORPP0_STB authdomain=CORPP0_STB); create table tb_a2000020&prod. (compress = yes reuse = yes) as SELECT * FROM CONNECTION TO con1 (  select *            from  			a2000020 a 		    where a.cod_ramo = &prod. and a.cod_cia = &cia. and a.tip_nivel = &tip_nivel.  ;) ; DISCONNECT FROM con1;  create table work.&tipo_dv. as select * from tb_a2000020&prod. where tip_nivel = &tip_nivel. order by cod_cia, num_poliza, num_riesgo, num_spto;  quit;  proc transpose data=work.&tipo_dv. out=work.&tipo_dv._&prod.;  by notsorted cod_cia notsorted num_poliza notsorted num_spto notsorted num_riesgo notsorted tip_nivel notsorted cod_ramo  ;  var val_campo; id cod_campo;  run;   proc sql;   drop table work.tb_a2000020&prod.; drop table work.&tipo_dv.;   quit;    %mend;   %macro rodar (prod, tipo, tip_nivel);  proc contents data=work.&tipo._&prod.     memtype=DATA      out=COLUNAS_Apolice     nodetails      noprint;  run;  proc sql;  create table work.formatacao as select  a.name, b.campo_formatado LENGTH=10000, b.num_ordem from  work.COLUNAS_Apolice a left join sgerais.campo_formatado b on (a.name = b.campo)  where b.cod_ramo = &prod. and b.tip_nivel = &tip_nivel.   order by b.num_ordem asc ;     quit;   PROC SQL;    CREATE TABLE WORK.QUERY_FOR_FORMATACAO AS     SELECT t1.NAME,            /* num_ordem */             (IFN(missing(t1.num_ordem), 999, t1.num_ordem, 999)) AS num_ordem,            /* campo_formatado */             (ifc(missing(t1.campo_formatado ), t1.name, t1.campo_formatado , t1.name)) LENGTH=10000 AS campo_formatado, 			',' as calculation 	       FROM WORK.FORMATACAO t1       ORDER BY num_ordem; QUIT;   PROC SQL;    CREATE TABLE WORK.coluna_formatada_final AS     SELECT t1.NAME,            t1.num_ordem,            /* Calculation */             (CATS(t1.campo_formatado,t1.calculation)) LENGTH=10000 AS Calculation       FROM WORK.QUERY_FOR_FORMATACAO t1 ; QUIT;    PROC SQL;  		select count(*) 		into :count 		from work.coluna_formatada_final order by num_ordem;  		select Calculation LENGTH=10000    		   into :col1 - :col&SysMaxLong 		from work.coluna_formatada_final order by num_ordem;  	create table work.&tipo._A_&prod. as  		select   		%DO i=1 %TO &count; 			&&col&i 			 		%end; 		cod_cia  from WORK.&tipo._&prod.;  drop table work.formatacao;  drop table WORK.QUERY_FOR_FORMATACAO; drop table work.COLUNAS_APOLICE;  QUIT;    proc stdize data=work.&tipo._A_&prod. reponly missing=0 out=work.&tipo._&prod._wW; run;  PROC SQL;  CREATE TABLE WORK.&tipo._&prod._R AS SELECT * FROM WORK.&tipo._&prod._wW   ORDER BY        			   NUM_POLIZA,                NUM_RIESGO,                NUM_SPTO;  DROP TABLE work.&tipo._&prod._wW; drop table work.&tipo._A_&prod.; DROP TABLE WORK.COLUNA_FORMATADA_FINAL; DROP TABLE WORK.&tipo._&prod.;  QUIT;  proc sql;   drop table sgerais.&tipo.&prod.;  create table sgerais.&tipo.&prod. as select * from work.&tipo._&prod._r;  drop table cluster.&tipo._&prod.;  create table cluster.&tipo._&prod. as select * from work.&tipo._&prod._r;   drop table work.&tipo._&prod._r;  quit;   %mend;  %consulta(503,1,dvapolice,1); %rodar(503, dvapolice, 1); %consulta(503,1,dvrisco,2); %rodar(503, dvrisco, 2);   "
  
 
' Run the code
sasProgram.Run
' Save the log file to LOCAL disk
sasProgram.Log.SaveAs "[local_log_sas]"
 
' Filter through the results and save just the LISTING type
For n=0 to (sasProgram.Results.Count -1)
' Listing type is 7
If sasProgram.Results.Item(n).Type = 7 Then
' Save the listing file to LOCAL disk
sasProgram.Results.Item(n).SaveAs "[local_log_sas_lst]"
End If
Next
app.Quit
