Option Explicit
'Se tienen que crear antes la variables que asignarlas
'Aqui estoy creando y asignando las variables de la url de la descarga y la de donde se va a guardar el archvio
Dim url, destination
url = "https://docs.google.com/spreadsheets/d/1FTBfdTbsC7q062MtgGDkdRQB8ADklxpa/edit?usp=drive_link&ouid=107319238878537577637&rtpof=true&sd=true"
destination = "C:\Users\gerar\OneDrive\Escritorio\MacroExcel\NecesarioMemPrac.xlsx"

'En este tramo de codigo lo que hago es crear dos tipos de opjetos
'1- el de XMLHTTP sirve para realizar las peticiones http
'2- el de ADODB para hacer lectura y escritura de archivos(para asi poder abrirlos )
Dim XMLHTTP, ADODB
Set XMLHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
Set ADODB = CreateObject("ADODB.Stream")

XMLHTTP.Open "GET", url, False
XMLHTTP.Send

ADODB.Open
ADODB.Type = 1 ' Binary
ADODB.Write XMLHTTP.ResponseBody
ADODB.Position = 0

ADODB.SaveToFile destination, 2 ' Overwrite

ADODB.Close
Set XMLHTTP = Nothing
Set ADODB = Nothing




