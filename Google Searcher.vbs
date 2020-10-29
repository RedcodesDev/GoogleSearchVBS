REM Definde Objects
set Suche = createobject("wscript.shell")
set voice = createobject("Sapi.SpVoice")

REM Create Inputbox
i= inputbox("What do you want to search on Google?","Google Search","Enter something")


REM Open URL
Suche.run "https://www.google.com/search?q=" + Replace(i," ","%20")

REM Speak Audio
voice.Speak("Here are the Results for: " + i)