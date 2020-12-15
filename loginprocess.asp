<%@LANGUAGE="JAVASCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/elab_conn.asp" -->

<%
	var guestId=Request.Form("guestId");
	var guestPassword=Request.Form("guestPassword");

	var rsadmin = Server.CreateObject("ADODB.RecordSet")	
	var sqlstmt = Server.CreateObject("ADODB.Command");
	
	sqlstmt.ActiveConnection = conn_str;
	sqlstmt.CommandText = "SELECT * FROM tblGuest Where guestId=? AND guestPassword=?";
	sqlstmt.Prepared = true;
    sqlstmt.Parameters.Append(sqlstmt.CreateParameter("param1", 202, 1, -1, guestId)); // adVarWChar
    sqlstmt.Parameters.Append(sqlstmt.CreateParameter("param2", 202, 1, -1, guestPassword)); // adVarWChar
	rsadmin = sqlstmt.Execute();
	
	if(rsadmin.EOF){Response.Write("<script language='javascript'>alert('Nama atau katau laluan salah!'); history.go(-1);</script>")}
	else{Session("username")=rsadmin.Fields.item("guestName").value Response.Redirect("viewResult.asp")}
	rsadmin.Close();

%>
		
