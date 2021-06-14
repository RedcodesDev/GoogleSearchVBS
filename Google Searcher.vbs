REM Define Objects
set shell = createobject("wscript.shell")
set voice = createobject("Sapi.SpVoice")


REM Language Selection

selector = msgbox("Do you want to set the Language to German?", vbYesNoCancel, "Google Search")

If selector = 6 Then

REM Language = German

REM Create Inputbox
i = inputbox("Was moechtest du auf Google suchen?","Google Search","Suche")

REM Check if User Canceld the Inputbox
If not IsEmpty(i) Then

REM Open URL
shell.run "https://www.google.de/search?q=" + Replace(i," ","%20")

REM Speak Audio
voice.Speak("Hier sind die ergebnisse fuer: " + i)

End If

Else If selector = 7 Then

REM Language = English

REM Create Inputbox
i = inputbox("What do you want to search on Google?","Google Search","Search query")

REM Check if User Canceld the Inputbox
If not IsEmpty(i) Then

REM Open URL
shell.run "https://www.google.com/search?q=" + Replace(i," ","%20")

REM Speak Audio
voice.Speak("Here are the Results for: " + i)

End If
End If
End If
