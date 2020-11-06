Sub EnviarEmail()
'
' Declaramos variables
'
Dim OutlookApp As Outlook.Application
Dim MItem As Outlook.MailItem
Dim cell As Range
Dim Asunto As String
Dim Correo As String
Dim Destinatario As String
Dim Registros As String
Dim Valor As String
Dim Msg As String
Dim FechaEnvio As String
Dim cliente As String
Dim ciudad As String
Dim cuenta As String
Dim Gic As String
Dim ext As String
Dim conv As String
Dim arch As String
Dim cc As String
Dim path As String
Dim sbdy As String
Dim ar As String
    '
    Set OutlookApp = New Outlook.Application
    '
   ' logo = ActiveWorkbook.path & "\Encabezado.jpg"
    logo = "D:\ENVÍO CORREOS\Encabezado.jpg"
    'Recorremos la columna EMAIL

    Set l1 = ThisWorkbook 'libro
    '
    
   lar = Sheets("Hoja1").Range("A" & Rows.Count).End(xlUp).Row
   
   
   
 
    
    
    largo = Sheets("Hoja1").Range("A" & Rows.Count).End(xlUp).Row
   For Each cell In Range("C9:C" & largo)  '
        If cell.Value <> "" Then 'si condicional
        'Asignamos valor a las variables
        '
       
         Destinatario = cell.Offset(0, -1).Value
        Correo = cell.Value
        cc = cell.Offset(0, 1).Value
        Registros = cell.Offset(0, 2).Value
        Valor = Format(cell.Offset(0, 3).Value, "$#,##0")
        Cantidad = cell.Offset(0, 4).Value
        Dia = cell.Offset(0, 5).Value
        FechaEnvio = Format(Now, "MMMM dd") & " de " & Format(Now, "yyyy")
        
        'FechaEnvio = Format(cell.Offset(0, 4).Value, "dd/mmm/yyyy")
        convenio = cell.Offset(0, -2).Value
        cliente = cell.Offset(0, -1).Value
        
        If Sheets("Hoja1").Range("b6") = "SI" Then
        Asunto = "INP01" & " " & cliente & " CUENTA DE COBRO LIBRANZA"
        Else
         Asunto = "CORREO_SEGURO - " & Left(convenio, 3) & "01 " & cliente & " CUENTA DE COBRO LIBRANZA"
         End If
        
        'Cuerpo del mensaje
        '
        
        imagen = "<Div> <IMG SRC=""" & logo & """ &  ><br><br></Div>"
        Fecha = "<Div><p><FONT COLOR= ""#043263 "" FONT FACE=Arial size=2.5>Bogota, " & FechaEnvio & "</FONT></p><br></Div>"
        Destinatarioo = "<Div><p><FONT COLOR= ""#043263 "" FONT FACE=Arial size=2.5>Señores: <B><br><br>" & Destinatario & " </B></FONT></p></Div>"
        City = "<Div><p><FONT COLOR= ""#043263 "" FONT FACE=Arial size=2.5>Ciudad</FONT> </p><br></Div>"
        Cuerpoo = "<Div><p><FONT COLOR= ""#043263 "" FONT FACE=Arial size=2.5>Referencia:<B> Envío archivo cuenta de cobro y novedades de Libranzas.</B><br><br>Nos permitimos remitir el archivo de la referencia, correspondiente al cobro de cuotas de créditos de libranzas formalizados y activos en BBVA Colombia, así como las novedades respectivas (Suspensiones, actualizaciones, soportes de libranzas, vistos buenos, entre otros), para su registro en nómina</FONT> </p><br></Div>"
        Cantidad = "<Div><p><FONT COLOR= ""#043263 "" FONT FACE=Arial size=2.5>Cantidad de registros:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B> " & Cantidad & "</B></FONT> </p></Div>"
        Valor = "<Div><p><FONT COLOR= ""#043263 "" FONT FACE=Arial size=2.5>Valor de cuotas a Cancelar:&nbsp;&nbsp;&nbsp;&nbsp;<B>" & Valor & "</B></FONT> </p></Div>"
        Pago = "<Div><p><FONT COLOR= ""#043263 "" FONT FACE=Arial size=2.5>Fecha de Pago:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B>" & Dia & "</B></FONT> </p><br></Div>"
        importatnt = "<Div><p><FONT COLOR= ""#2DCCCD "" FONT FACE=Arial size=2.5>** NOTA IMPORTANTE **</FONT> </p></Div>"
        importanttt = "<Div><p><FONT COLOR= ""#2DCCCD "" FONT FACE=Arial size=2.5> Recuerda que para aplicar los pagos, debemos contar con los recursos disponibles en la cuenta asignada al convenio, así como, con el archivo que detalla los clientes, créditos y valor a pagar. Una vez cuente con esta información, debe informar a la Oficina Gestora del Convenio, con copia al siguiente correo del área operativa: LIBRANZAS.PAGOS@BBVA.COM </FONT></p></Div>"
        aclaracion = "<Div><p><FONT COLOR= ""#043263 "" FONT FACE=Arial size=2.5>Atentamente<br><br> REPORTES LIBRANZAS<br>BBVA COLOMBIA<br><br>Cualquier aclaración o información adicional con gusto le atenderemos, por favor comunicarse con el Área de Reportes de Libranzas BBVA Colombia, en Bogotá al Tel:(057)(091)4379310 23091<br></FONT></p><br></Div>"
        
        
        
        'Msg = "Bogotá, " & FechaEnvio & vbNewLine
        'Msg = Destinatario & vbNewLine & vbNewLine
        'Msg = ciudad & vbNewLine & vbNewLine
        'Msg = Msg & "Referencia: Envío archivo cuenta de cobro y novedades de Libranzas." & vbNewLine
        'Msg = Msg & "Nos permitimos remitir el archivo de la referencia, correspondiente al cobro de cuotas de créditos de libranzas formalizados y activos en BBVA Colombia, así como las novedades respectivas (Suspensiones, actualizaciones, soportes de libranzas, vistos buenos, entre otros), para su registro en nómina" & vbNewLine & vbNewLine
        'Msg = Msg & "**NOTA IMPORTANTE**" & vbNewLine & vbNewLine
        'Msg = Msg & "Recuerda que para aplicar los pagos, debemos contar con los recursos disponibles en la cuenta asignada al convenio, así como, con el archivo que detalla los clientes, créditos y valor a pagar. Una vez cuente con esta información, debe informar a la Oficina Gestora del Convenio, con copia" & vbNewLine
        'Msg = Msg & "al siguiente correo del área operativa: LIBRANZAS.PAGOS@BBVA.COM" & cuenta & vbNewLine & vbNewLine
        'Msg = Msg & "Atentamente:" & vbNewLine
        'Msg = Msg & "REPORTES LIBRANZAS"
        'Msg = Msg & "BBVA COLOMBIA" & vbNewLine & vbNewLine
        'Msg = Msg & "Cualquier aclaración o información adicional con gusto le atenderemos, por favor comunicarse con el Área de Reportes de Libranzas BBVA Colombia, en Bogotá al Tel:(057)(091)4379310 23035" & vbNewLine
        'Msg = Msg & "Para devoluciones de Sobrantes por favor realizar la solicitud por medio de cualquiera de nuestras oficinas"
        
        'sbdy = Fecha & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine
        sbdy = imagen & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine
        
        
        
        
        path = ThisWorkbook.path & "\"
        ruta = path & convenio & "\"
          ChDir ruta
        arch = Dir(ruta & "*.*")
        
        If arch <> "" Then
        
     Set MItem = OutlookApp.CreateItem(olMailItem)
        With MItem
        
            .To = Correo
            If cc <> " " Then
            .cc = cc
            Else
            End If
            .Subject = Asunto
            '.Body = Msg
            .Display
            .HTMLBody = (imagen) & (Fecha) & (Destinatarioo) & (City) & (Cuerpoo) & (Cantidad) & (Valor) & (Pago) & (importatnt) & (importanttt) & (aclaracion)
            
            Do While arch <> ""
            .Attachments.Add ruta & "\" & arch
            arch = Dir()
            Loop
            .Send
DETALLE = Now()
             cell.Offset(0, 6).Value = "Enviado " & DETALLE
              End With
Else

   cell.Offset(0, 6).Value = "NO Enviado"
    
End If

'Else



 End If
 
Next


MsgBox "proceso terminado"
End Sub
