'********************************************************
'Objetivo: Programa para convertir una tabla MySQL a HTML
'   Fecha:03/Jun/2022
'   Autor: edcruces99
'
'  220527: Creación del programa inicial
'  220603: Actualización incluir lectura parámetros
'********************************************************

'Crear Objeto
Set FSO = CreateObject("Scripting.FileSystemObject")

'Crear Array
Set arrCampos = CreateObject("System.Collections.ArrayList")
Set arrBDD =  CreateObject("System.Collections.ArrayList")

'Leer archivo entrada
Filename = WScript.Arguments.Item(0)

'Leer nombre archivo sin extensión
arrFilename=Split(Filename,".")
for each x in arrFilename
    FilenameJustName=x
	Exit For
next

'Abrir archivo
Set f = fso.OpenTextFile(filename)

k = 1
bEncontrado = 0
Do Until f.AtEndOfStream
  linea = f.ReadLine
  
  If InStr(linea,"ENGINE=") > 0 Then
     Exit Do
  End If

  If InStr(linea,"-- Base de datos:") > 0 Then
        For i = 1 to Len(linea)
          If Mid(linea, i, 1) = "`" Then
             i=i+1
             While Mid(linea, i, 1) <> "`"
                strC = strC & Mid(linea, i, 1)
                i=i+1
             Wend
             arrBDD.Add strC
             strC=""
             Exit For
          End If
        Next
     linea = f.ReadLine
  End If
  
  If InStr(linea,"CREATE TABLE") > 0 Then
     bEncontrado = 1
     linea = f.ReadLine
  End If
     
  If bEncontrado = 1 Then
     If InStr(linea,"`") > 0 Then
        'MsgBox k & " " & linea
                           
        'Leer entre caracteres comillitas >> `R_LOCK` tinyint(1) UNSIGNED NOT NULL COMMENT 'R_Lock',
        For i = 1 to Len(linea)
          If Mid(linea, i, 1) = "`" Then
             i=i+1
             While Mid(linea, i, 1) <> "`"
                strC = strC & Mid(linea, i, 1)
                i=i+1
             Wend
             arrCampos.Add strC
             'MsgBox strC
             strC=""
             Exit For
          End If
        Next       
     End If
  End If
        
  k=k+1
Loop

f.Close

'Archivo de resultado
Set f = FSO.OpenTextFile(FilenameJustName & "_out.sql" ,2 , True)

'Agregar Nombre BDD
f.WriteLine
For Each campo In arrBDD
    f.WriteLine "BDD:" & trim(campo)
Next

c = 1
'Script: Validacion de Campos------------------------------------
f.WriteLine
For Each campo In arrCampos
    f.WriteLine "$" & campo & " = trim($_POST['" & campo & "']);"
    c=c+1
Next
f.WriteLine "-------------------------------------"
f.WriteLine ""

'Script Instruccion INSERT--------------------------
f.WriteLine
f.WriteLine "      $sql = ""INSERT INTO TABLA ("" ."

For Each campo In arrCampos
    f.WriteLine "                " & Chr(34) & Chr(44) & campo & Chr(32) & Chr(34) & " ."
Next 

f.Write "                " & Chr(34) & ") VALUES ( "
For n = 1 To c-1
    f.Write ","
    f.Write "?"
Next
f.Write ") " & Chr(34) & ";"
f.WriteLine
f.WriteLine "      $q = $pdo->prepare($sql);"
f.WriteLine "      "
f.WriteLine "      "

f.Write "      if ($q->execute(array('0',"
For Each campo In arrCampos
    f.Write "$" & campo & ","
Next
    f.Write "))) QUITAR ULTIMA COMA"

'Script Instruccion <th>--------------------------
f.WriteLine ""
f.WriteLine "-------------------------------------"
For Each campo In arrCampos
    f.WriteLine "<th>" & campo & "</th>"
Next
f.WriteLine "-------------------------------------"
f.WriteLine ""

'Script Instruccion <A.>--------------------------
For Each campo In arrCampos
    f.WriteLine Chr(39) & ",A." & campo & " " & Chr(39) & " " & "."
Next
f.WriteLine "-------------------------------------"
f.WriteLine ""

'Script Instruccion <td>--------------------------
For Each campo In arrCampos
    f.WriteLine "echo " & Chr(39) & "<td>" & Chr(39) & ". $row[" & Chr(39) & campo & Chr(39) & "] . " & Chr(39) & "</td>" & Chr(39) & ";"
Next
f.WriteLine "-------------------------------------"
f.WriteLine ""

'Script Instruccion HTML--------------------------
For Each campo In arrCampos
    f.Write "$" & campo & ","
    If campo = "USUARIO" Then
       Exit For
    End If
Next
f.write "$id_usuario,$fecha_local_db,$fecha_local_db)))" & vbCrLf
f.WriteLine "          $r_actualizado=1;"
f.WriteLine "      else"
f.WriteLine "          $r_actualizado=0;"
f.WriteLine "                           "
f.WriteLine "      //Último rec ingresado"
f.WriteLine "      $stmt = $pdo->query(" & Chr(34) & "SELECT LAST_INSERT_ID()" & Chr(34) & ");"
f.WriteLine "      $last_id = $stmt->fetchColumn();"
f.WriteLine
f.WriteLine
f.WriteLine "-------------------------------------"
f.WriteLine ""

'Script HTML--------------------------------------------------------
f.WriteLine "         <div class=" & Chr(34) & "row" & Chr(34) & ">"
f.WriteLine "           <form action=" & Chr(34) & "insert_administradores.php" & Chr(34) & " method=" & Chr(34) & "post" & Chr(34) & ">"

For Each campo In arrCampos
    f.WriteLine "                  <div class=" & Chr(34) & "form-group row" & Chr(34) & ">"
    f.WriteLine "                    <label for=" & Chr(34) & campo & Chr(34) & " class=" & Chr(34) & "col-sm-2 col-form-label" & Chr(34) & ">" & campo & ":</label>"
    f.WriteLine "                    <div class=" & Chr(34) & "col-sm-8" & Chr(34) & ">"
    f.WriteLine "                      <input type=" & Chr(34) & "text" & Chr(34) & " class=" & Chr(34) & "form-control" & Chr(34) & " id=" & Chr(34) & campo & Chr(34) & " name=" & Chr(34) & campo & Chr(34) & " placeholder=" & Chr(34) & "Capturar campo" & Chr(34) & " maxlength=" & Chr(34) & "100" & Chr(34) & " size=" & Chr(34) & "100" & Chr(34) & " required>"
    f.WriteLine "                    </div>"
    f.WriteLine "                 </div>"
Next

    f.WriteLine "                 <button type=" & Chr(34) & "reset" & Chr(34) & " class=" & Chr(34) & "btn btn-secondary" & Chr(34) & ">Limpiar</button>"
    f.WriteLine "                 <a href=" & Chr(34) & "administradores.php" & Chr(34) & " class=" & Chr(34) & "btn btn-info active" & Chr(34) & " role=" & Chr(34) & "button" & Chr(34) & ">Regresar</a>"
    f.WriteLine "                 <button type=" & Chr(34) & "submit" & Chr(34) & " class=" & Chr(34) & "btn btn-success" & Chr(34) & ">Guardar</button>"
    f.WriteLine "            </form>"

'Script Instruccion <td>--------------------------
f.WriteLine "-------------------------------------"
         f.WriteLine ""
For Each campo In arrCampos
         f.WriteLine "<div class=" & Chr(34) & "form-row" & Chr(34) & ">"
		 f.WriteLine "<div class=" & Chr(34) & "form-group col-md-12" & Chr(34) & ">"
         f.WriteLine "<label for=" & Chr(34) & "input" & campo & Chr(34) & ">Nombre del " & campo & "</label>"
         f.WriteLine "<input type=" & Chr(34) & "text" & Chr(34) & " class=" & Chr(34) & "form-control" & Chr(34) & " name=" & Chr(34) & "input" & campo & Chr(34) & " id=" & Chr(34) & "input" & campo & Chr(34) & " placeholder=" & Chr(34) & "Indique el " & campo & Chr(34) & " required>"
         f.WriteLine "</div>"
         f.WriteLine "</div>"
         f.WriteLine ""
Next
f.WriteLine "-------------------------------------"
f.WriteLine ""

f.Close
Set FSO= Nothing
Wscript.Quit

'Chr(34) = "
'Chr(39) = '
'Chr(32) = espacio
'Chr(44) = ,
