Option Explicit

Dim StartTime, EndTime
StartTime = Timer()
'Se tienen que crear antes la variables que asignarlas
'Aqui estoy creando y asignando las variables de la url de la descarga y la de donde se va a guardar el archvio
Dim url, destination
url = "url"
destination = "destination"

'En este tramo de codigo lo que hago es crear dos tipos de opjetos
'1- el de XMLHTTP sirve para realizar las peticiones http
'2- el de ADODB para hacer lectura y escritura de archivos(para asi poder abrirlos y escribirlos en el destino que queramos)
Dim XMLHTTP, ADODB
Set XMLHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
Set ADODB = CreateObject("ADODB.Stream")

'En el sigueinte bloque de codigo lo que se hace es abrir la url con el método "GET" para así obetener el archivo del enlace
'que se ha puesto anteriormente en la variable y se envia para que el otro objeto pueda leerlo y escribirlo
XMLHTTP.Open "GET", url, False
XMLHTTP.Send

If XMLHTTP.Status <> 200 Then
    MsgBox "Error en la descarga. Código de estado: " & XMLHTTP.Status
    WScript.Quit
End If

WScript.Sleep 1000

' Una vez enviado el archivo este objeto se encarga de recibirlo, abrirlo y guardarlo en el destino que nosotros hallamos elegido
with ADODB
    .type = 1 
    .open
    .write XMLHTTP.responseBody
    .savetofile destination, 2 '//overwrite
end with

'Por último lo que hacemos es setear estos objetos en nulos para que así se limpie la memoria asignándolos a nada
Set XMLHTTP = Nothing
Set ADODB = Nothing


Dim excelApp, workbook, worksheet
Dim filePath, destinatarioColumn, notaColumn, statusColumn, lastRow, destinatario
Dim i

i = 2
' Ruta del archivo Excel
filePath = "path"

' Columnas en Excel
destinatarioColumn = 4 
notaColumn = 2 
statusColumn = 5 

' En esta aprte de abre el excel y se abre tanto como la hojas como se asigna la variable lastRow a la ultima fila de la hoja
Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False 
Set workbook = excelApp.Workbooks.Open(filePath)
Set worksheet = workbook.Worksheets(1) 
lastRow = worksheet.Cells(worksheet.Rows.Count, destinatarioColumn).End(-4162).Row 

'' creación del objeto para enviar un correo
Dim objMessage
Set objMessage = CreateObject("CDO.Message")

' Configurar los detalles del servidor SMTP
objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 ' Uso de SMTP
objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "server" ' Especifica tu servidor SMTP
objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 0 ' Puerto SMTP (ajústalo según tu configuración)
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 ' Autenticación SMTP
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "user" ' Tu nombre de usuario SMTP
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "pass" ' Tu contraseña SMTP
objMessage.Configuration.Fields.Update

For i = 2 To lastRow
    Dim nota, status

    nota = worksheet.Cells(i, notaColumn).Value


    If nota < 5 Then
        status = "Suspenso"
    Else
        status = "Aprobado"
    End If

    worksheet.Cells(i, statusColumn).Value = status

    destinatario = worksheet.Cells(i, destinatarioColumn).Value
    objMessage.From = "from" 
    objMessage.To = destinatario
    objMessage.Subject = "Status examen"
    objMessage.TextBody = "Hola, su examen ha sido un: " & status
    objMessage.Send

    If i = 10 Then
        Exit For
    End If



    worksheet.Cells(i, statusColumn).Value = status
Next
workbook.Save
workbook.Close
EndTime = Timer()
MsgBox("Seconds to 2 decimal places: " & FormatNumber(EndTime - StartTime, 2))
MsgBox("Acabe")