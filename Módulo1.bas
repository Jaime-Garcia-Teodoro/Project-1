Attribute VB_Name = "M�dulo1"
Sub Macro_total()
    Limpiar_Excel_Avance
    Nombres_anuncios
    Base_tot
    Base_recuerda
    Bases_anuncios
    Recuerdo
    Pregunta_2
End Sub


Sub Limpiar_Excel_Avance()
    Dim wsDestino As Worksheet

    ' Define la hoja de destino que deseas limpiar
    Set wsDestino = ThisWorkbook.Sheets("Excel avance") ' Cambia el nombre seg�n sea necesario

    ' Borra solo el contenido del rango C4:AF70
    wsDestino.Range("C4:AF70").ClearContents
End Sub


Sub Nombres_anuncios()
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim colOrigen As Long
    Dim colDestino As Long
    Dim i As Integer
    Dim filaOrigen As Long
    Dim filaDestino As Long

    ' Define las hojas de origen y destino
    Set wsOrigen = ThisWorkbook.Sheets("Hoja1")   ' Cambia "Hoja1" por el nombre de tu hoja de origen
    Set wsDestino = ThisWorkbook.Sheets("Excel avance")  ' Cambia "Excel avance" por el nombre de tu hoja de destino

    ' Configuraci�n inicial para las filas y columnas de origen y destino
    filaOrigen = 10      ' Fila donde est�n los nombres de los anuncios en la hoja de origen
    filaDestino = 3     ' Fila donde quieres comenzar a pegar en la hoja de destino
    colDestino = 3       ' Comienza en la columna C de la hoja de destino

    ' Bucle para copiar los primeros 12 nombres de anuncios
    For i = 2 To 13
        colOrigen = i   ' Columna de origen en la fila de anuncio (va de 2 a 13)

        ' Copia el nombre del anuncio solo si no est� vac�o
        If wsOrigen.Cells(filaOrigen, colOrigen).Value <> "" Then
            ' Pega el nombre del anuncio en la fila de destino y en la columna con espacio de 2 columnas entre cada nombre
            wsDestino.Cells(filaDestino, colDestino).Value = wsOrigen.Cells(filaOrigen, colOrigen).Value
            ' Avanza a la siguiente posici�n de columna en destino (salta 2 columnas cada vez)
            colDestino = colDestino + 2
        End If
    Next i
End Sub


Sub Base_tot() 'Obtenemos las bases de cada anuncio

    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim colDestino As Integer
    Dim k As Integer
    Dim filaDestino As Integer

    ' Define las hojas de origen y destino
    Set wsOrigen = ThisWorkbook.Sheets("Hoja1")   ' Cambia "Hoja1" por el nombre de tu hoja de origen
    Set wsDestino = ThisWorkbook.Sheets("Excel avance")  ' Cambia "Excel avance" por el nombre de tu hoja de destino

    ' Encuentra la �ltima fila con datos en la hoja de origen
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).Row
    
    ' Define la fila destino fija y la primera columna de destino
    filaDestino = 6
    colDestino = 3  ' Comienza en la columna C

    ' Recorre las filas de la hoja de origen para buscar la pregunta espec�fica
    For i = 1 To ultimaFila
        ' Verifica si la celda contiene la pregunta "Registros"
        If wsOrigen.Cells(i, 1).Value = "Registros" Then
                ' Recorre las columnas
                For k = 2 To 13
                    ' Copia el valor de la columna en la hoja de destino en las columnas.
                    wsDestino.Cells(filaDestino, colDestino).Value = wsOrigen.Cells(i, k).Value
                    ' Salta dos columnas para la pr�xima posici�n
                    colDestino = colDestino + 2
                Next k
            Exit For
        End If
    Next i
End Sub


Sub Base_recuerda()
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim colDestino As Integer
    Dim k As Integer
    Dim filaDestino As Integer

    ' Define las hojas de origen y destino
    Set wsOrigen = ThisWorkbook.Sheets("Hoja1")   ' Cambia "Hoja1" por el nombre de tu hoja de origen
    Set wsDestino = ThisWorkbook.Sheets("Excel avance")  ' Cambia "Excel avance" por el nombre de tu hoja de destino

    ' Encuentra la �ltima fila con datos en la hoja de origen
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).Row
    
    ' Define la fila destino fija y la primera columna de destino
    filaDestino = 6
    colDestino = 4  ' Comienza en la columna D

    ' Recorre las filas de la hoja de origen para buscar la pregunta espec�fica
    For i = 1 To ultimaFila
        ' Verifica si la celda contiene la pregunta "Registros: Recuerda"
        If wsOrigen.Cells(i, 1).Value = "Registros: Recuerda" Then
                ' Recorre las columnas
                For k = 2 To 13
                    ' Copia el valor de la columna en la hoja de destino en las columnas
                    wsDestino.Cells(filaDestino, colDestino).Value = wsOrigen.Cells(i, k).Value
                    ' Salta dos columnas para la pr�xima posici�n
                    colDestino = colDestino + 2
                Next k
            Exit For
        End If
    Next i
End Sub

Sub Bases_anuncios()
'Simplemente pone los nombres de "Base total" y "Base recuerda" encima de las bases que hemos buscado antes

    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim filaDestino As Long
    Dim colDestino As Long
    Dim colOrigen As Long
    Dim filaOrigen As Long
    
    Set wsOrigen = ThisWorkbook.Sheets("Hoja1")   ' Cambia "Hoja1" por el nombre de tu hoja de origen
    Set wsDestino = ThisWorkbook.Sheets("Excel avance")  ' Cambia "Excel avance" por el nombre de tu hoja de destino
    
    filaDestino = 5      ' Fila en la hoja de destino donde quieres empezar a pegar
    colDestino = 3       ' Comienza en la columna C de la hoja de destino
    filaOrigen = 10      ' Fila en la hoja de origen donde est�n los nombres de campa�a

    ' Recorre las primeras 12 columnas de la hoja de origen
    For colOrigen = 2 To 13
        ' Verifica si hay un nombre de campa�a en la columna actual de la hoja de origen
        If wsOrigen.Cells(filaOrigen, colOrigen).Value <> "" Then
            ' Copia "Base total" en la primera columna del par en la hoja de destino
            wsDestino.Cells(filaDestino, colDestino).Value = "Base total"
            ' Copia "Base recuerda" en la segunda columna del par en la hoja de destino
            wsDestino.Cells(filaDestino, colDestino + 1).Value = "Base recuerda"
            ' Avanza dos columnas en la hoja de destino para el siguiente par
            colDestino = colDestino + 2
        End If
    Next colOrigen
End Sub


Sub Recuerdo()
'Primero se crean las variables que vamos a utilizar

'Se crean las variables tanto de origen (donde cogeremos los datos), como de destino (donde se pegar�n)
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet

'Se crean el resto de variables a utilizar
    Dim ultimaFila As Long
    Dim i As Long
    Dim colDestino As Long
    Dim k As Long
    Dim filaDestino As Long
    Dim j As Long

    ' Define las hojas de origen y destino
    Set wsOrigen = ThisWorkbook.Sheets("Hoja1") 'Cambia "Hoja1" por el nombre de tu hoja de origen
    Set wsDestino = ThisWorkbook.Sheets("Excel avance") 'Cambia "Excel avance" por el nombre de tu hoja de destino

    ' Encuentra la �ltima fila con datos en la hoja de origen.
    ' Esto te va a servir para que busque en todas las filas a la hora de buscar una pregunta
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).Row
    
    ' Define la primera fila de destino y la primera columna de destino.
    ' Lo iremos modificando para que no se sobreescriban.
    filaDestino = 8 ' Comienza en la fila 8
    colDestino = 3  ' Comienza en la columna C

    ' Recorre las filas de la hoja de origen para buscar la pregunta espec�fica
    For i = 1 To ultimaFila
        
        ' Verifica si la celda contiene la pregunta "RECUERDO ANUNCIO"
        If wsOrigen.Cells(i, 1).Value = "RECUERDO ANUNCIO" Then
            For j = i To ultimaFila
                
                ' Cuando encuentre la pregunta, desde esa fila hasta la �ltima busca "SI"
                If wsOrigen.Cells(j, 1).Value = "SI" Then
                    ' Luego recorre las columnas de origen para obtener el dato de todos los anuncios
                    For k = 2 To 13
                        
                        ' Copia el valor de la columna en la hoja de destino en las columnas
                        wsDestino.Cells(filaDestino, colDestino).Value = wsOrigen.Cells(j, k).Value
                        ' Salta a la siguiente columna de la hoja de destino para el pr�ximo anuncio
                        colDestino = colDestino + 1
                    Next k
                End If
            Next j
        End If
    Next i
End Sub


Sub Pregunta_2()
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim colDestino As Long
    Dim k As Long
    Dim filaDestino As Long
    Dim j As Long

    ' Define las hojas de origen y destino
    Set wsOrigen = ThisWorkbook.Sheets("Hoja1")   ' Cambia "Hoja1" por el nombre de tu hoja de origen
    Set wsDestino = ThisWorkbook.Sheets("Excel avance")  ' Cambia "Excel avance" por el nombre de tu hoja de destino

    ' Encuentra la �ltima fila con datos en la hoja de origen
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).Row
    
    ' Define la fila destino fija y la primera columna de destino
    filaDestino = 10
    colDestino = 4  ' Comienza en la columna D

    ' Recorre las filas de la hoja de origen para buscar la pregunta espec�fica
    For i = 1 To ultimaFila
        ' Verifica si la celda contiene la pregunta "Pregunta 2"
        If wsOrigen.Cells(i, 1).Value = "Pregunta 2" Then
            For j = i To ultimaFila
                ' Busca "Media"
                If wsOrigen.Cells(j, 1).Value = "Media" Then
                    ' Recorre las columnas
                    For k = 2 To 13
                        ' Copia el valor de la columna en la hoja de destino en las columnas
                        wsDestino.Cells(filaDestino, colDestino).Value = wsOrigen.Cells(j, k).Value
                        ' Salta dos columnas para la pr�xima posici�n
                        colDestino = colDestino + 2
                    Next k
                Exit For
                    
                End If
            Next j
        End If
    Next i
End Sub

