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
sasProgram.Text = conteudo 
  
 
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
