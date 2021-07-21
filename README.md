# Kursplanering Go
Grundläggande integration mellan Google Spreadsheet och Google Calendar, för kursplanering med mera.

Än så länge i beta-läge, men fungerar.

Videoguider:

* 16 minuter lång video som visar i princip allt som finns i version 1.4 beta: https://youtu.be/sOTENhmi3X0
* Hel videoserie som visar det mesta av verktyget i 2–3 minuter korta videor, version 1.2 beta: https://www.youtube.com/watch?v=sffYfpw2ue4&list=PL5sq4vtv3sTHAHfNNbRwSZNl5rpi7GrsA
* Video som visar hur du kommer igång: https://youtu.be/2OIdw5jeItY
* Video som visar hur du flyttar runt planering exempelvis för att lektioner blir inställda en dag: https://youtu.be/DWazLndmKJQ

Använd det här kalkylbladet för att komma igång: https://docs.google.com/spreadsheets/d/1Wu83SaHQnSZ4u9xn9eYrBVxRXNW_CC-KSlE3z1pOB4o/copy

Steg för att använda:
1) Skapa en egen kopia (genom arkivmenyn, skapa kopia, om det inte dyker upp automatiskt).
2) Godkänn behörigheter för skriptet. Du får frågor om detta första gången du kör kommandon från Kursplanering Go-menyn.

## Funktioner

* Kan skapa en Google-kalender och lägga in markerade händelser (lektioner) med angiven startid, sluttider, rubrik och beskrivning.
* Kan uppdatera markerade händelser (lektioner) i Google-kalendern för att avspegla ändringar i kalkylbladet.
* Kan läsa in Google-kalender och lägga in data för händelser (lektioner) i kalkylbladet. Uppdaterar befintliga händelser om de redan finns. **Nytt för 1.3 beta.**
* Kan radera Google-kalendern (om man verkligen vill).

## Vad gör funktionerna i menyn?

* **Skapa uppdatera valda kalenderhändelser**: Markera raderna för ett antal lektioner och välj det här i menyn för att skapa lektioner i Google-kalendern (eller uppdatera, om lektionerna redan fanns). Kalender skapas automatiskt, om det inte redan fanns en. Den får namnet som är angivet i cell B1.
* **Radera valda kalenderhändelser**: Markera raderna för en eller flera lektioner och välj det här i menyn för att ta bort lektionerna från Google-kalendern. Informationen finns fortfarande kvar i kalkylbladet, så du kan lägga tillbaka lektionerna igen senare om du vill.
* **Läs in kalenderhändelser**: Detta läser in alla kalenderhändelser (lektioner) ett år framåt och bakåt i tiden, för kalendern med det ID som står i cell B2. De lektioner som redan finns i kalkylbladet (med lektions-ID) kommer att uppdateras med information från kalendern, och övriga lektioner kommer att dyka upp som nya rader. Det går fint att ändra ordning på raderna i kalkylbladet efteråt. *En bugg i version 1.4 gör att du måste ha minst en rad med lektioner i kalkylbladet för att köra det här kommandot. Det räcker att skriva vad som helst på översta raden, och ta bort den raden efter import från Google Calendar.*
* **Radera hela kalendern**: Detta tar bort Google-kalendern (och raderar alla lektions-ID:n). All information finns fortfarande kvar i kalkylbladet, så du kan skapa en ny kalender med samma innehåll om du skulle ångra dig. (Den kalendern kommer dock att få nytt ID och ny länk att dela med exempelvis elever, så även om du skapar en kalender med samma innehåll behöver du dela en ny länk.)
* **Om Kursplanering Go**: Detta visar en länk till projektsidan för Kursplanering Go.

## Frågor och svar

**Spelar det någon roll om jag raderar rader eller flyttar runt information i kalkylbladet?**
Nej, i princip inte. Skriptet förväntar sig att information för lektionerna finns i vissa kolumner, men så länge informationen finns där är det lungt. (Den som vill joxa med skriptet kan ändra vilka kolumner som används i de översta raderna i skriptet.)

**Vad gör "infoga rubriker i beskrivningar"?**
Cell F1 används för att markera om rubrikerna för de kolumner som används för beskrivningar ska läggas in i händelserna i Google-kalendern. Om F1 är 0 (eller något annat värde som tolkas som falskt) utelämnas rubrikerna. Annars kommer varje del av beskrivningen föregås av tillhörande rubrik.

**Kan jag ändra till andra rubriker för beskrivningar?**
Japp. Det är bara att redigera i kalkylbladet.

**Kan jag lägga till nya kolumner för beskrivningar, eller ta bort?**
Japp. Cellerna D1 och D2 anger i vilken kolumn beskrivningarna börjar och slutar, så uppdatera där om du lägger till fler kolumner. De kolumner som är tomma kommer inte att läggas in i beskrivningar i Google-kalendern, men det kan vara bra att ta bort såna som du inte använder ändå.

**Kan jag ändra kalenderns namn i efterhand?**
Inte genom skriptet, men du kan redigera kalendern genom Google Calendar.

**Det verkar lite krångligt med formler för att skriva in tider för lektioner. Kan jag göra det manuellt?**
Japp. Att hantera formler och andra saker i kalkylblad kan spara dig massor av jobb, men skriptet bryr sig bara om att det finns åtminstone starttid, sluttid och rubrik för varje lektion. Hur den information skapas spelar ingen roll.

**Är det farligt att experimentera med skriptet?**
Nej. Om du är orolig, skapa en ny kopia av kalkylbladet och börja om från början. Det värsta som kan hända är att du får en rad extra kalendrar i Google Calendar (som kan tas bort för hand) och att din nya kopia av kalkylbladet blir så rörig att du inte kan använda det längre.

**Men är det inte farligt att skriptet kan hantera alla mina kalkylblad och alla mina kalendrar?**
Så länge du bara gör ändringar i kalkybladet, och inte i skriptet, kommer bara det aktiva kalkylbladet och kalendrarna de synkar mot påverkas. Om du går in i skriptet och ändrar alltför mycket kommer det att kunna ändra även andra kalkylblad och kalendrar, men för att få till något sånt krävs det att du anstränger dig rätt ordentligt. Att skriptet ber om tillåtelse att hantera *alla* kalkylblad och kalendrar beror på att det är där Google dragit gränsen för den här typen av skript – allt eller inget.
