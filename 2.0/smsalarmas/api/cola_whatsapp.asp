<!--#include file="../include/scriptdb/conectdb.asp"-->
<%
' GET /smsalarmas/api/cola_whatsapp.asp   (sin token)
' Mismos queries de JermSoft5/example3.js:
'   SELECT id_salida, movil, sms FROM sSMS_BsalidaSpeedWhatsApp WHERE status = 0
'   UPDATE sSMS_BsalidaSpeedWhatsApp SET status = 1 WHERE id_salida = <id>
' Devuelve los mensajes en JSON y los deja en status = 1.
' Si algo falla devuelve {"ok":false,"paso":"...","error":"..."} en vez de un 500 mudo.

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

Sub colaFalla(paso)
    Response.Write "{""ok"":false,""paso"":""" & paso & """,""error"":""" & colaJson(Err.Number & " - " & Err.Description) & """}"
    Response.End
End Sub

Dim conn, rs, sb, first, idSalida

On Error Resume Next
Response.ContentType = "application/json; charset=utf-8"

' Misma conexion que usa todo el sistema (include/scriptdb/conectdb.asp)
Dim cmd
Set cmd = conectdb()
Set conn = cmd.ActiveConnection
If Err.Number <> 0 Or conn Is Nothing Then Call colaFalla("conectar")
If conn.State <> 1 Then Call colaFalla("conectar (conexion cerrada)")

Set rs = conn.Execute("SELECT id_salida, movil, sms FROM sSMS_BsalidaSpeedWhatsApp WHERE status = 0")
If Err.Number <> 0 Then Call colaFalla("select")

sb = "{""ok"":true,""items"":["
first = True
Do While Not rs.EOF
    idSalida = CStr(rs("id_salida"))
    conn.Execute "UPDATE sSMS_BsalidaSpeedWhatsApp SET status = 1 WHERE id_salida = " & idSalida
    If Err.Number <> 0 Then Call colaFalla("update id_salida " & idSalida)

    If Not first Then sb = sb & ","
    first = False
    sb = sb & "{""id_salida"":" & idSalida & _
              ",""movil"":""" & colaJson(Trim(rs("movil") & "")) & """" & _
              ",""sms"":""" & colaJson(rs("sms") & "") & """}"
    rs.MoveNext
    If Err.Number <> 0 Then Call colaFalla("leyendo filas")
Loop
sb = sb & "]}"

rs.Close
conn.Close
Response.Write sb
%>
