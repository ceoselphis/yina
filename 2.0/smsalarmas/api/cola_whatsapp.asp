<!--#include file="../include/scriptdb/parametros.inc"-->
<%
' GET /smsalarmas/api/cola_whatsapp.asp   (sin token)
' Mismos queries de JermSoft5/example3.js:
'   SELECT id_salida, movil, sms FROM sSMS_BsalidaSpeedWhatsApp WHERE status = 0
'   UPDATE sSMS_BsalidaSpeedWhatsApp SET status = 1 WHERE id_salida = <id>
' Devuelve los mensajes en JSON y los deja en status = 1.

Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType = "application/json; charset=utf-8"
Response.Expires = -1
Response.AddHeader "Cache-Control", "no-store, no-cache"

Function colaJson(s)
    Dim t : t = CStr(s)
    t = Replace(t, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, Chr(13), "\r")
    t = Replace(t, Chr(10), "\n")
    t = Replace(t, Chr(9), "\t")
    colaJson = t
End Function

Dim conn, rs, sb, first, idSalida
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open DB_CONNECTIONSTRING, DB_USERNAME, DB_PASSWORD
conn.CursorLocation = adUseClient

Set rs = conn.Execute("SELECT id_salida, movil, sms FROM sSMS_BsalidaSpeedWhatsApp WHERE status = 0")

sb = "{""ok"":true,""items"":["
first = True
Do While Not rs.EOF
    idSalida = CStr(rs("id_salida"))
    conn.Execute "UPDATE sSMS_BsalidaSpeedWhatsApp SET status = 1 WHERE id_salida = " & idSalida

    If Not first Then sb = sb & ","
    first = False
    sb = sb & "{""id_salida"":" & idSalida & _
              ",""movil"":""" & colaJson(Trim(rs("movil") & "")) & """" & _
              ",""sms"":""" & colaJson(rs("sms") & "") & """}"
    rs.MoveNext
Loop
sb = sb & "]}"

rs.Close : conn.Close
Response.Write sb
%>
