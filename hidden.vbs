Option Explicit

Dim strURL, strTempFile
strURL = "https://github.com/jockop77/fff/raw/main/Best_Gits.zip" ' Замените URL на нужный
strTempFile = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%temp%") & "\Best_Gits.zip" ' Имя файла

Dim objXMLHTTP, objStream

' Создаем объект для загрузки файла
Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
objXMLHTTP.open "GET", strURL, False
objXMLHTTP.send

' Проверка статуса ответа
If objXMLHTTP.Status = 200 Then
    ' Сохраняем загруженный файл
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 1 ' 1 = Binary
    objStream.Open
    objStream.Write objXMLHTTP.responseBody
    objStream.SaveToFile strTempFile, 2 ' 2 = Overwrite
    objStream.Close
Else
    ' Можно записать ошибку в файл вместо вывода сообщения
    Dim fso, errorLog
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Путь к лог-файлу
    errorLog = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%temp%") & "\error_log.txt"
    
    ' Записываем статус ошибки в лог
    Dim logStream
    Set logStream = fso.OpenTextFile(errorLog, 8, True) ' 8 = Append mode
    logStream.WriteLine "Ошибка загрузки файла. Код статуса: " & objXMLHTTP.Status & " - " & Now
    logStream.Close
    
    ' Освобождаем объекты
    Set logStream = Nothing
    Set fso = Nothing
End If

' Освобождаем объекты
Set objStream = Nothing
Set objXMLHTTP = Nothing
