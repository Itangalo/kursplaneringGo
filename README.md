# Kursplanering Go
Grundläggande integration mellan Google Spreadsheet och Google Calendar, för kursplanering med mera.

Än så länge i beta-läge, men fungerar.

Videoguider:

* Hel videoserie som visar det mesta av verktyget i 2–3 minuter korta videor: https://www.youtube.com/watch?v=sffYfpw2ue4&list=PL5sq4vtv3sTHAHfNNbRwSZNl5rpi7GrsA
* Video som visar hur du kommer igång: https://youtu.be/2OIdw5jeItY
* Video som visar hur du flyttar runt planering exempelvis för att lektioner blir inställda en dag: https://youtu.be/DWazLndmKJQ

Använd det här kalkylbladet för att komma igång: https://docs.google.com/spreadsheets/d/1Wu83SaHQnSZ4u9xn9eYrBVxRXNW_CC-KSlE3z1pOB4o/edit?usp=sharing

Steg för att använda:
1) Skapa en egen kopia genom arkivmenyn, skapa kopia.
2) Godkänn behörigheter för skriptet.

### Funktioner

* Kan skapa en Google-kalender och lägga in markerade händelser (lektioner) med angiven startid, sluttider, rubrik och beskrivning.
* Kan uppdatera markerade händelser (lektioner) i Google-kalendern för att avspegla ändringar i kalkylbladet.
* Kan läsa in Google-kalender och lägga in data för händelser (lektioner) i kalkylbladet. Uppdaterar befintliga händelser om de redan finns. **Nytt för 1.3 beta.**
* Kan radera Google-kalendern (om man verkligen vill).
