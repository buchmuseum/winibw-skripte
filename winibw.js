//(Declarations)
	
//End of (Declarations)


function __FehlerRegistrierung(erg) {
// In Feld 4720 werden verschiedene Sachverhalte gespeichert, die einer späteren Nachbearbeitung bedürfen
// Grundsätzlich ist die Struktur: 4720 [Fehlertyp]Feldname Feldinhalt_allegro-ID
// Die allegro-ID wird benötigt, um zu einem späteren Zeitpunkt im angegebnen Feld an Hand einer Konkordanz den Feldinhalt durch die PICA-IDN eines Normsatzes zu ersetzen
// Wird die Funktion von der Funktion Suche aus aufgerufen, dann war dort die Normdaten-Suche in der GND nicht erfolgreich und sie soll in allegro fortgesetzt werden
// erg hat dann einen Wert zwischen 1 und 4

if (!application.activeWindow.title) 
	{
	application.messageBox("Diese Funktion steht momentan nicht zur Verfügung!","Um die Funktion nutzen zu können\nmus im Bearbeitungsmodus zunächst das Feld \nund die zu registrierende Zeichenkette eingegeben werden.\nBeispiel: 5590 Meier, Hans","");
	}
else
	{
	var tag = application.activeWindow.title.tag;			//merke den tag
	application.activeWindow.title.startOfField(false);		//gehe zum Start des Feldes	
	application.activeWindow.title.wordRight(1,false);		//gehe ein Word nach rechts, um den tag auszuschließen
	application.activeWindow.title.endOfField(true);		//Selektiere von der Position alles bis zum Ende
	var content = application.activeWindow.title.Selection;         //merke den Suchbegriff
//	application.activeWindow.title.endOfField(false);		//hebe die Markierung auf
	if (erg < 1 && erg > 4)
		{
		erg = __Pruef("Eine Frage","Welcher Fehlertyp soll registriert werden?\n 1 = allegro-Personennormsatz noch nicht in GND\n 2 = allegro-Körperschaftssatz noch nicht in GND\n 3 = allegro-Schlagwortsatz noch nicht in GND\n 4 = andere Fehler","1,2,3,4,5","1");
		}
	if (erg == 1)
		{
		typ = "d-3a";
		ind = "PER ";
		}
	else if (erg == 2)
		{
		typ = "d-3a";
		ind = "KOE ";
		}
	else if (erg == 3)
		{
		typ = "d-3a";
		ind = "SSW ";
		}
	else if (erg == 4) 
		{
		// weitere Fehler sind noch nicht definiert
		search = typ = "d-?";
		application.messageBox("Geht nicht!","Weitere Fehlertypen sind noch nicht definiert!","");
		}
	if (erg > 0 && erg < 4)
		{
		// application.messageBox("",ind + content,"")
		application.activeWindow.clipboard = ind + content;   //indikator + content in die Zwischenablage
		var oShell = new ActiveXObject("Shell.Application"); 
		var commandtoRun = "p:\\\\AC\\dbsm\\BATCH\\IBW_a99_bib.cmd";  
		oShell.ShellExecute(commandtoRun,"","","open","1"); 
		application.messageBox("Suche in der allegro Datenbank ist abgeschlossen!","Weiter?","");
		content = application.activeWindow.clipboard;
		if (content != "nix")
			{			
			application.activeWindow.title.insertText(content.substring(0,content.indexOf("_")) + "\n4720 |" + typ + "|" + tag + " " + content);
			}
		else
			{
			application.messageBox("Schade","Nichts gefunden oder abgebrochen","");
			}
		}
	}	
}
function __getAbteilung() {

	return (__getProfVal("Abteilung erfassen","abteilung","Bitte geben Sie Ihre Abteilung (zB. FE, ERW, IE) ein."));
	
}
function __getProfVal(boxtit,valname,prompTxt,art) {

var sect="dnbUser";
var value = application.getProfileString(sect, valname, "");

  if (!value || art == "korr") {
  value = __dnbPrompter(boxtit,prompTxt,value);
  
    if (value != null) {
    application.writeProfileString(sect, valname, value);
    }
  }

	
}

function maskeEinfuegen(titel) {
	
	var suchePlus = false;
	
	//Datenmaske einfügen:
    var startP = application.activeWindow.title.selStart;
	application.activeWindow.title.insertText(titel);
	
	// Bearbeiterkuerzel einfuegen **
	application.activeWindow.title.startOfBuffer (false);
	suchePlus = application.activeWindow.title.find("**", false, false, true); 
	while (suchePlus == true) { 
	application.activeWindow.title.insertText(__getProfVal("Kürzel erfassen","kuerzel","Bitte geben Sie Ihr Bearbeiterkürzel ein."));
	application.activeWindow.title.startOfBuffer (false);
	suchePlus = application.activeWindow.title.find("**", false, false, true); 
	}

	// Standort einfügen ##
	suchePlus = false;
	application.activeWindow.title.startOfBuffer (false);
	suchePlus = application.activeWindow.title.find("##", false, false, true); 
	while (suchePlus == true) { 
	application.activeWindow.title.insertText(__getProfVal("Standort erfassen","standort","Bitte geben Sie Ihren Standort ein."));
	//	application.messageBox("",suchePlus,"");
	application.activeWindow.title.startOfBuffer (false);
	suchePlus = application.activeWindow.title.find("##", false, false, true); 
	} 

	// Abteilung einfügen ||
	suchePlus = false;
	application.activeWindow.title.startOfBuffer (false);
	suchePlus = application.activeWindow.title.find("||", false, false, true); 
	while (suchePlus == true) { 
	application.activeWindow.title.insertText("|" +__getAbteilung() + "|");
	//	application.messageBox("",suchePlus,"");
	application.activeWindow.title.startOfBuffer (false);
	suchePlus = application.activeWindow.title.find("||", false, false, true); 
	} 	
	
	// Text aus der Zwischenablage einfuegen !!
/*	application.activeWindow.title.startOfBuffer (false);
	suchePlus = application.activeWindow.title.find("!!", false, false, true); 
	idn = application.activeWindow.clipboard
	var strConfirm = "Zeichenfolge >" + idn + "< aus Zwischenablage einfügen?"
	if (__dnbConfirm("Bestätigung",strConfirm)) {
		application.activeWindow.title.insertText(idn);
	}
*/
		
	application.activeWindow.title.setSelection(startP, startP, false);
	suchePlus = application.activeWindow.title.find("++", false, false, true);

	if (suchePlus == true){
		//Entfernen der Plusse, der Cursor bleibt hier stehen:
		application.activeWindow.title.deleteSelection();
	}
	
}
function __Prompter(ttl,txt,dflt) {
	// Place your function code here
	// Übername der internen Funktion aus den DNB-Standardfunktionen mit der gleichen Funktionalität
var __Prompter = utility.newPrompter();
var msg;
	
msg = __Prompter.prompt(ttl,txt,dflt,null,null);
if (msg == 1) msg = __Prompter.getEditValue();
else msg = null;

return msg;
	
}
function __Pruef(boxTit,strTxt,werte,dflt) {
// Übername der internen Funktion aus den DNB-Standardfunktionen mit der gleichen Funktionalität
var erg = "";
var antw = true;
wertePrf = werte.toLowerCase();
do 
	{
	if (antw == false) 
		{
		boxTxt = "FALSCHE EINGABE! (" + erg + ")\n"  + "Nur " + werte + " eingeben!\n\n" + strTxt;
		} 
	else 
		{
		boxTxt = strTxt;
		}
	erg = __Prompter(boxTit,boxTxt,dflt);
	if (erg == null) 
		{
		break;
		} 
	else 
		{
		erg = erg.toLowerCase();
		} 
	if (wertePrf.indexOf(erg) > -1) 
		{ 
		antw = true
		} 
	else 
		{
		antw = false 
		}
	} while(antw != true);
return erg;	
	
}
function DatenmaskeEinfuegen(maskenNr) {
	var theFileInput = utility.newFileInput();
	var thePrompter = utility.newPrompter();
	var antwort, dasKommando = "", kommandoTitel, kommandoNorm;
	var theLine;
	var titel;
	var fileName = "\\datenmasken\\maske" + maskenNr + ".txt";
	var fileNameBinDir = "\\defaults\\datenmasken\\maske" + maskenNr + ".txt";

	//auf welchem Schirm befinden wir uns?
	screenNr = application.activeWindow.Variable("scr");

	//Kommandos zum Eingeben von Titeln und Normdaten
	kommandoTitel = "\\inv 1"
	kommandoNorm = "\\inv 2"

	// Datenmaskendatei im Verzeichnis profiles\<user>\datenmasken oeffnen:
	if (!theFileInput.openSpecial("ProfD", fileName)) {
		if (!theFileInput.openSpecial("BinDir", fileNameBinDir)) {
			application.messageBox("Fehler", "Datei " + maskenNr + " wurde nicht gefunden.", "error-icon");
			return;
		}
	}
	for (titel = ""; !theFileInput.isEOF(); ) {
		titel += theFileInput.readLine() + "\n"
	}
	theFileInput.close();

	var editing = (application.activeWindow.title != null);

	if ((Datentyp = titel.substr(0,4)) == "0500")
		dasKommando = kommandoTitel
	else if ((Datentyp = titel.substr(0,3)) == "005")
		dasKommando = kommandoNorm
	else if (!editing) {
		//wenn weder 0500 noch 005 vorkommt, muss er Benutzer nun entscheiden:
		antwort = thePrompter.select("Auswahl", "Leider konnte die WinIBW nicht erkennen," +
			"ob die Datenmaske für Titel oder Normdaten verwendet werden soll.\n" +
			"Bitte wählen Sie aus:", "Titeldaten\nNormdaten");

		if (!antwort) {
			// Benutzer hat den Dialog abgebrochen:
			return;
		}
		if (antwort == "Titeldaten")
			dasKommando = kommandoTitel
		else if (antwort == "Normdaten")
			dasKommando = kommandoNorm
	}

	if (dasKommando == "") {
		// The data is inserted in the edit window at the cursor position.
		// The "++" is not removed (as in maskeEinfuegen), because the data already present
		// might contain this.
		application.activeWindow.title.endOfBuffer(false);
		maskeEinfuegen(titel);
		return;
	}

	//wenn editing = true, dann wird das Kommando in neuem Fenster ausgeführt
	application.activeWindow.command(dasKommando, editing);

	// Eingeben oder Abbruch, falls kein titleedit vorliegt:
	if (application.activeWindow.title) {
	    maskeEinfuegen(titel);
	}
	else {
		application.messageBox("Fehler", "Datenmaske kann nicht eingefügt werden!", "error-icon");
		return;
	}	
}
function DBSMAaa() {
		application.activeWindow.command("e", false);
	// wenn nicht möglich, Meldung und Ende
	if (!application.activeWindow.title) application.messageBox("Fehler!","Die Maske konnte nicht aufgerufen werden!\nEventuell sind Sie nicht eingelogt\noder haben Sie keine Bearbeitungsrechte?","");
	else
		{
		application.activeWindow.simulateIBWKey("FE");
		// Vorgabewerte setzen / abfragen
		var erg;
		var eArt;
		var sect = "dnbUser";
		var strKuerzel = application.getProfileString(sect, "kuerzel", "");
		var strAbteilung = application.getProfileString(sect, "abteilung", "");
		var jetzt = new Date();
		var jahr = jetzt.getFullYear();
		var erg = __Pruef("Eine Frage","Welche Erwerbungsart?\n 1 = Kauf \n 2 = Tausch \n 3 = Geschenk","1,2,3","2")
		if (erg == 1) eArt = "ka";
		if (erg == 2) eArt = "ta";
		if (erg == 3) eArt = "ge";

		// Datenmaske einfügen und mit Vorgabewerten ausfüllen
		DatenmaskeEinfuegen("DBSMAaa");
		// Bearbeiter-Name in 4700
		application.activeWindow.title.findTag("4700", 0, false, true, false);
		application.activeWindow.title.endOfField(false);	
		application.activeWindow.title.insertText("|" + strAbteilung + "|" + strKuerzel + "* ");
		if (eArt == "ta") {
			application.activeWindow.title.insertText("DNB-Tausch")
			}
		// Erwerbungsart in 8510
		application.activeWindow.title.findTag("8510", 0, false, true, false);
		application.activeWindow.title.endOfField(false);	
		application.activeWindow.title.insertText(eArt);
		// aktuelles Jahr in 8598
//		application.activeWindow.title.findTag("8598", 0, false, true, false);
//		application.activeWindow.title.endOfField(false);	
//		application.activeWindow.title.insertText(jahr);

		// zum Startpunkt der Erfassung navigieren
		application.activeWindow.title.findTag("0600", 0, false, true, false);
		application.activeWindow.title.endOfField(false);	
		}	
}
function DBSMAam() {
		application.activeWindow.command("e", false);
	// wenn nicht möglich, Meldung und Ende
	if (!application.activeWindow.title) application.messageBox("Fehler!","Die Maske konnte nicht aufgerufen werden!\nEventuell sind Sie nicht eingelogt\noder haben Sie keine Bearbeitungsrechte?","");
	else
		{
		application.activeWindow.simulateIBWKey("FE");
		// Vorgabewerte setzen / abfragen
		var sect = "dnbUser";
		var strKuerzel = application.getProfileString(sect, "kuerzel", "");
		var strAbteilung = application.getProfileString(sect, "abteilung", "");
		var jetzt = new Date();
		var monat = jetzt.getMonth() + 1
		var datum = jetzt.getDate() + "." + monat + "." +  jetzt.getFullYear()

		// Datenmaske einfügen und mit Vorgabewerten ausfüllen
		DatenmaskeEinfuegen("DBSMAam");
		// Bearbeiter-Name in 4700
		application.activeWindow.title.findTag("4700", 0, false, true, false);
		application.activeWindow.title.endOfField(false);	
		application.activeWindow.title.insertText("|" + strAbteilung + "|" + strKuerzel + "*Erasmus Amsterdam, bestellt am " + datum);

		// zum Startpunkt der Erfassung navigieren
		application.activeWindow.title.findTag("1100", 0, false, true, false);
		application.activeWindow.title.endOfField(false);	
		}
}
function DBSMAbschluss() {

if (!application.activeWindow.Variable("scr") || application.activeWindow.Variable("scr") == "FI")
	{
	if (application.activeWindow.Variable("scr") == "FI") application.messageBox("Diese Funktion steht momentan nicht zur Verfügung!","Um die Funktion nutzen zu können\nmus ein Titel des DBSM-Bestandes ausgewählt werden.","");
	else application.messageBox("Diese Funktion steht momentan nicht zur Verfügung!","Um die Funktion nutzen zu können\nmüssen Sie sich erst einloggen\nund einen Titel des DBSM-Bestandes auswählen.","");
	return
	}
if (application.activeWindow.Variable("scr") != "MI") application.activeWindow.simulateIBWKey("F7");	//Bearbeiten ein
else application.activeWindow.title.startOfBuffer;
if (!application.activeWindow.title.find("!!DBSM",true,false,false))
	{
	application.activeWindow.simulateIBWKey("FE");
	application.messageBox("Diese Funktion steht momentan nicht zur Verfügung!","Die Funktion kann nur auf einen Titel des DBSM-Bestandes angewendet werden.","");
	return
	}
var jetzt = new Date();												// aktuelles Datum auf jetzt
var jahr = jetzt.getFullYear();										// aktuelless Jahr aus jetzt auf jahr
meldung = ""
if (!application.activeWindow.title.findTag("7100", 1, false, true, false))			// wenn es nur eine 7100 gibt
	{
	// notation und signatur finden und merken
	application.activeWindow.title.find("!!DBSM", true, false, false);			// Suche den richtigen Datensatz
	application.activeWindow.title.find(" ; ", true, false, false);				// gehe zum Start der Notation
	application.activeWindow.title.endOfField(true);						// markiere bis zum Ende der Zeile
	notation = application.activeWindow.title.selection						// setze diesen Text auf Notation
	notation = notation.substring(3,notation.length)						// " ; " vorn abschneiden
	application.activeWindow.title.startOfField(false);						// gehe an den Anfang der Zeile
	application.activeWindow.title.lineUp(1, true)							// Markiere die Zeile darüber						
	signatur = application.activeWindow.title.selection						// setze diese auf Signatur
	signatur = signatur.substring(5,signatur.length - 1)						// schneide vorne die Feldnummer und hinten den Zeilenumbruch ab
	//signatur = signatur.substring(5,signatur.indexOf(" @") - 2)						// schneide vorne die Feldnummer und hinten " @" bzw. den Zeilenumbruch ab
	// in Feld 7100 "@ k" austragen
	if (application.activeWindow.title.find(" @ k", true, false, false) != "")		// wenn "@ k" vorhanden
		{
//		application.activeWindow.title.endOfField(true);					// markiere bis zum Ende der Zeile
		application.activeWindow.title.deleteSelection();						// lösche Markierung
		meldung = meldung + "\n- in Feld 7100 '@ k' ausgetragen"
		}
	// eventuell ESK austragen
	if (notation != signatur)										// wenn Notation 
		{
		if (notation.indexOf("ESK") > -1)
			{
			application.activeWindow.title.find("(ESK", false, false, false);
			application.activeWindow.title.endOfField(true);
			application.activeWindow.title.insertText(signatur)
			meldung = meldung + "\n- in 7109 " + notation + " durch " + signatur + " überschrieben."
			}
		}
	// Feld 8598 für Jahresstatistik füllen
	if (!application.activeWindow.title.findTag("8598", 0, false, true, false))		// wenn 8598 noch nicht vorhanden
		{
		application.activeWindow.title.endOfField(false);						// gehe ans Ende der Zeile
		application.activeWindow.title.insertText("\n8598 " + jahr);				// füge Zeilenumbruch und aktuelles Jahr in 8598 ein
		meldung = meldung + "\n- '8598 " + jahr + "' eingefügt"
		}
	// Satzart ändern
	// application.activeWindow.title.startOfBuffer()
	// application.activeWindow.title.wordRight()
	// application.activeWindow.title.endOfField(true)
	// art = application.activeWindow.title.selection
	// if (art.substring(1,1) != "b" && art.substring(2,1) == "a")
	//	{
	//	art1 = art.substring(0,2)
	//	application.activeWindow.title.insertText(art1)
	//	meldung = meldung + "\n- Satzart von " + art + " auf " + art1 + " geändert."
	//	}
	// Meldung ausgeben, was verändert wurde
	if (meldung > "") 
		{
		var prompt = utility.newPrompter();
		prompt.setDebug(true);
		if (prompt.confirmEx("Alles in Ordnung?","Folgendes wurde gemacht:\n "+meldung+"\nSoll gespeichert werden?", "Yes", "No", "", "", false) != 1) application.activeWindow.simulateIBWKey("FR");
		}
		else application.messageBox("Alles in Ordnung?","Es wurden keine Veränderungen im Exemplarsatz vorgenommen.","")
	}
	else application.messageBox("Achtung!","Der automatisierte Abschluss funktioniert nur, wenn nur ein DBSM-Exemplarsatz vorhanden ist.\nBitte die notwendigen Änderungen von Hand vornehmen.","")

}
function DBSMArtikel() {
	application.activeWindow.command("e", false);
	// wenn nicht möglich, Meldung und Ende
	if (!application.activeWindow.title) application.messageBox("Fehler!","Die Maske konnte nicht aufgerufen werden!\nEventuell sind Sie nicht eingelogt\noder haben Sie keine Bearbeitungsrechte?","");
	else
		{
		application.activeWindow.simulateIBWKey("FE");
		// Vorgabewerte setzen / abfragen
		var sect = "dnbUser";
		var strKuerzel = application.getProfileString(sect, "kuerzel", "");
		var strAbteilung = application.getProfileString(sect, "abteilung", "");
		var jetzt = new Date();

		// Datenmaske einfügen und mit Vorgabewerten ausfüllen
		DatenmaskeEinfuegen("DBSMArtikel");
		// Bearbeiter-Name in 4700
		application.activeWindow.title.findTag("4700", 0, false, true, false);
		application.activeWindow.title.endOfField(false);	
		application.activeWindow.title.insertText("|" + strAbteilung + "|" + strKuerzel + "*L4; ");

		// zum Startpunkt der Erfassung navigieren
		application.activeWindow.title.findTag("1100", 0, false, true, false);
		application.activeWindow.title.endOfField(false);
		}
}
function DBSMGNDKopie() {
// Funktion übernimmt Daten aus einem kiz-Satz in einen piz-Satz
// IDN auslesen
if (!application.activeWindow.title) 
	{
	application.messageBox("Diese Funktion steht momentan nicht zur Verfügung!","Um die Funktion nutzen zu können\nmuss im Bearbeitungsmodus zunächst die IDN im Feld 510\nmarkiert werden werden.","");
	}
else
	{
	var texta = new Array();								//Array für mehr als einen Sprachencode definieren
	var n = 0											//Array-Zähler auf 0 setzen
	var WinID1 = application.activeWindow.windowID;					//1. Fenster-ID merken
	var IDN = application.activeWindow.title.selection;				//markierte IDN merken
	application.activeWindow.command("f IDN " + IDN, false);			//Suche ausführen	
	var WinID2 = application.activeWindow.windowID;					//2. Fenster merken
	application.activeWindow.simulateIBWKey("F7");					//Bearbeitungsmodus ein
	text = application.activeWindow.title.findTag("043",0,false,true);	//Feld 043 suchen und ohne Feldbezeichnung auf text merken
	text2 = ""											//text2 für Ausgabetext definieren
	if (text.indexOf(";") > 0) 								//wenn in text ";" --> mehr als ein Ländercode
		{
		while (text.indexOf(";") > 0) 						//solange noch weiterer Ländercode vorhanden
			{											//Starte Schleife
			texta[n] = text.substring(0,text.indexOf(";"));				//ersten Ländercode auf Array n 
			text = text.substring(text.indexOf(";")+1, text.length);		//Rest auf text merken
														//Bundesland abschneiden
			if (texta[n].indexOf("-") != texta[n].lastIndexOf("-")) texta[n] = texta[n].substring(0, texta[n].lastIndexOf("-"));
			if (n > 0) 										//wenn es nicht der erste durchlauf ist
				{
				if (texta[n] != texta[n-1]) text2 = texta[n] + ";" + text2;		//Ergebnis[n] nur vor text2 pappen, wenn Ergebnis[n] nicht gleich Ergebnis[n-1]
				}
			else											//sonst
				{
				text2 = texta[n] + ";";								//Ergebnis auf text2 merken
				}
			n = n+1										//Array-Zähler 1 hochsetzen
			}
		}
	if (text.indexOf("-") != text.lastIndexOf("-")) text = text.substring(0, text.lastIndexOf("-"));	//in text Bundesland abschneiden
	text2 = text2 + text														//text an text2 anhängen
	var test = "";
	var i = 0;
	do	{
		text1 = application.activeWindow.title.findTag("551",i,true,true);	//Feld 551 suchen
		if (text1 == "") test = "Ende";							//wenn nicht gefunden --> Ende
														//sonst testen ob "orta" vorhanden, 
		if (text1.indexOf("orta") > 0) 							//wenn ja 
			{
			text1 = text1.substring(0,text1.length - 1) + "w"				//String als text1 merken
			test = "Ende"										//und Ende
			}
		else 												//sosnt
			{
				i = i+1										//occurrence um 1 erhöhen
				if (i > 9) test = "Ende"							//wenn occurrence > 9 --> Ende
			}
		}while (test != "Ende")
	application.activeWindow.simulateIBWKey("FE");					// Bearbeitungsmodus beenden (ESC)
	application.activeWindow.closeWindow();						//Fenster schließen
	application.activateWindow(WinID1);							//zurück zum Ausgangsfenster
	application.activeWindow.title.endOfBuffer(false);				//ans Ende der Aufnahme
	application.activeWindow.title.insertText("\n043 "+text2+"\n"+text1);		//text und text1 einfügen
	//application.messageBox("",text,"")^
	}	
}
function DBSMGndNbmDow() {
	// Funktion startet den Normdaten-Downloade nach allegro-HANS
	// IDN auslesen
	var strIDN = application.activeWindow.Variable("P3GPP");
	// wenn erfolgreich, Explorer starten und Satz im Portal anzeigen
	if (!strIDN) 
		{
		application.messageBox("Diese Funktion steht momentan nicht zur Verfügung!","Es muss erst ein Normsatz ausgwählt sein.","");
		}
	else
		{
		application.activeWindow.command("dow mrc", false);
		var oShell = new ActiveXObject("Shell.Application"); 
		var commandtoRun = "p:\\\\AC\\ha2\\ndwl-ibw.bat";  
		oShell.ShellExecute(commandtoRun,"","","open","1"); 
		}
}
function DBSMLink(search, searchTerm, vTag) {
//	application.messageBox("Gesucht wird mit folgender Suchfrage:", searchTerm,"");

// Die Funktion soll generell beim Herstellen von Links in 3XXX, 42XX, 51XX, 53XX, 559X, 6710 und 680X verwendet werden
// Es wird: 
// - wenn noch nicht erfasst, der Suchstring abgefragt,
// - je nach Feldnummer, eine Suchfrage gebildet und die "Schärfe" der Suche eingestellt (nur Normdaten eines Typs, nur Level 1 usw.), 
// - die eigentliche Suche durchgeführt
// Wurde nichts gefunden, wird die Suche solange wiederholt, bis ein sinnvolles Suchergebnis vorliegt oder die Suche abgebrochen wurde,
// Dabei kann:
// - die "Schärfe" der Suche zurückgesetzt werden (alle Normdaten, auch Level unter 1 usw.),
// - die Suchfrage von Hand angepasst werden,
// - abgebrochen und in allegro weiter gesucht werden oder
// - die Funktion ganz ohne Reaktion beendet werden
// Wurde mehr als ein Ergebnis gefunden, wird die Funktion für die Auswahl des richtigen Satzes unterbrochen und kann mit Scrip/weiter fortgesetzt werden
// Wurden so viele Treffer gefunden, dass das set nicht mehr in der Kurztitelanzeige angezeigt wird, so wird mit einer Hinweis-Meldung abgebrochen
// Wurde genau ein Satz gefunden, oder die Funktion mit Script/weiter fortgesetzt, so erfolgt noch eine Kontroll-Frage "Richtiger Satz?" 
// Wird diese bestätigt, so wird in Normsätzen:
// - je nach Verknüpfungsfeld wenn noch nicht vorhanden das Teilbestandskennzeichen "f", "s" oder "g" gesetzt
// - wenn noch nicht vorhanden das Benutzungskennzeichen "o" gesetzt
//*-----------------------------------------------------------------------------------------------------*/

if (!searchTerm && !application.activeWindow.title) 
	{
	application.messageBox("Diese Funktion steht momentan nicht zur Verfügung!","Um die Funktion nutzen zu können\nmuss im Bearbeitungsmodus zunächst das Feld \nund die zu suchende Zeichenkette eingegeben werden.\nBeispiel: 5590 Meier, Hans","");
	}
else
	{
	if (!search) var search = "nix"
	var strIDN = ""
	var strMeldung = ""
	var strEinleitung = ""
	var strMarkierung = ""
	var winID1
	var winID2
            if (!searchTerm) 
                { 
	    var tag = application.activeWindow.title.tag;			//merke den tag
                var rest = ""
	    application.activeWindow.title.startOfField(false);		//gehe zum Start des Feldes
	    application.activeWindow.title.endOfField(true);		//Selektiere von der Position alles bis zum Ende
                var searchTerm = application.activeWindow.title.Selection;	//markiere den Suchbegriff
                searchTerm = searchTerm.substring(5,searchTerm.length);     //merke den Suchbegriff 
                }
            else var tag = vTag
            if (searchTerm.indexOf("$") > 0) 
                {
                rest = searchTerm.substring(searchTerm.indexOf("$"), searchTerm.length);
                searchTerm = searchTerm.substring(0, searchTerm.indexOf("$"));
                }
            content = searchTerm;							//Suchbegriff als content für eventuelle Suche in allegro merken
 
	// wenn der Suchbegriff leer ist (ganz leer oder ein Buchstabe oder nur einleitende Formel (In: / Rezension zu: / [Label]))
	if (searchTerm.length < 2 || searchTerm.length == 1 + searchTerm.indexOf(":") || searchTerm.length == 2 + searchTerm.indexOf(":") || searchTerm.length == 1 + searchTerm.indexOf("]")|| searchTerm.length == 2 + searchTerm.indexOf("]"))
		{
		erg = __Prompter("Wonach soll gesucht werden?","Geben Sie den Suchbegriff ein!")
		// einleitende Formel und Leerzeichen
		if (searchTerm.length == 2 + searchTerm.indexOf(":") || searchTerm.length == 2 + searchTerm.indexOf("]"))
			{
			var searchTerm = searchTerm + erg;
			}
		// einleitende Formel ohne Leerzeichen
		else if (searchTerm.length == 1 + searchTerm.indexOf(":") || searchTerm.length == 1 + searchTerm.indexOf("]"))
			{
			var searchTerm = searchTerm + " " + erg;
			}
		// keine einleitende Formel
		else
			{
			var strMarkierung = "0"
			var searchTerm = erg;
			}
		content = erg
		}
//application.messageBox("","searchTerm: " + searchTerm + "\nsearch: " + search + "\ntag: " + tag,"")
 //          if (!tag) searchTerm = '"'+searchTerm+'"'	// Anführungsstriche vor und hinter das Suchwort
 //          else searchTerm = '"'+searchTerm+'?"'	// Anführungsstriche vor und hinter das Suchwort und Joker dran
             searchTerm = searchTerm+"?"
	// wenn es ein 30XX-Feld ist, suche in per
	if (tag.substring(0,2) == "30") 
		{
		search = "f per ";
		searchTerm = searchTerm + " and rec n and tbs f";
		erg = "1"
		}
	// wenn es ein 31XX-Feld ist, suche in kor
	if (tag.substring(0,2) == "31") 
		{
		search = "f kor ";
		searchTerm = searchTerm + " and rec n and tbs f";
		erg = "1"
		}	// wenn es ein In-Vermerk ist, suche in tit
	else if (tag == "4241" || tag == "4261" || tag == "4262") 
		{
              	search = "f tst ";
//		strEinleitung = searchTerm.substring(1,searchTerm.indexOf(":")+2);
//		searchTerm = '"' + searchTerm.substring(searchTerm.indexOf(":")+2,searchTerm.length);
                        //  geändert wegen "Enthalten in" statt "In:" 
                        if (searchTerm.indexOf(":") > 0)
                            {
		    strEinleitung = searchTerm.substring(1,searchTerm.indexOf(":"));
		    searchTerm = '"' + searchTerm.substring(searchTerm.indexOf(":")+2,searchTerm.length);
                            }
                            else
                            {
                            strEinleitung = "Enthalten in"
                            }
		erg = __Pruef("Eine Frage","Wonach soll gesucht werden?\n 1 = Zeitschrift / Serie \n 2 = Monografie","1,2","1");
		if (erg == 1) searchTerm = searchTerm + " and bbg a#v?";
		else if (erg == 2) searchTerm = searchTerm + " and bbg (aa? or af? or ac?)";
		else if (!erg) search = "nix";
		}
	// wenn es eine Systemstelle ist, suche in syp
	else if (tag.substring(0,3) == "532" || tag == "6710") 
		{
		search = "f syf ";
		searchTerm = searchTerm + " or syp " + searchTerm + " or syw " + searchTerm
		if (tag.substring(0,3) == "532") searchTerm = searchTerm + " and bbg tk?";
		else searchTerm = searchTerm + " and bbg tq?";
		}

	// wenn es ein Gestaltungsmerkmal-Feld oder ein Inhaltserschließungsfeld ist, frage nach, wo gesucht werden soll
	else if (tag.substring(0,3) == "559" || tag.substring(0,3) == "680" || tag.substring(0,2) == "51")
		{
                        if (vTag == "5591") erg = 4
                        else if (vTag == "6800") erg = 4
		else erg = __Pruef("Eine Frage","Wonach soll gesucht werden?\n 1 = Person\n 2 = Körperschaft\n 3 = Sachschlagwort\n 4 = Geografikum\n 5 = Werk-Schlagwort","1,2,3,4,5","1");
		if (erg == 1) search = "f per ";
		else if (erg == 2) search = "f kor ";
		else if (erg == 3 || erg == 4 || erg == 5) search = "f an ";
		else if (!erg) search = "nix";
		// bei Inhaltserschließung nur Level 1
		if (tag.substring(0,2) == "51")
			{
			if (erg == 1) searchTerm = searchTerm + " and bbg tp1";
			else if (erg == 2) searchTerm = searchTerm + " and bbg (tb1 or tf1)";
			else if (erg == 3) searchTerm = searchTerm + " and bbg ts1";
			else if (erg == 4) searchTerm = searchTerm + " and bbg tg1";
			else if (erg == 5) searchTerm = searchTerm + " and bbg tu1";
			}
		// bei Gestaltungsmerkmalen alle Level
		else			
			{
			if (tag.substring(3,4) == "9")
				{
				strEinleitung = searchTerm.substring(1,searchTerm.indexOf("]")+2);
				searchTerm = '"'+searchTerm.substring(2+searchTerm.indexOf("]"),searchTerm.length);
				}
			if (erg == 1) searchTerm = searchTerm + " and bbg (tp? or tn?)";
			else if (erg == 2) searchTerm = searchTerm + " and bbg (tb? or tf?)";
			else if (erg == 3) searchTerm = searchTerm + " and bbg ts?";
			else if (erg == 4) searchTerm = searchTerm + " and bbg tg?";
			else if (erg == 5) searchTerm = searchTerm + " and bbg tu?";
			}
		}
           
//	application.messageBox("Gesucht wird mit folgender Suchfrage:",search + searchTerm,"");
 
	// wenn eine sinnvolle Suchfrage gebildet wurde	
	if (search != "nix") 
		{
		winID1 = application.activeWindow.windowID;
		strCommand = search + searchTerm;
		Erfolg: do
			{
//application.messageBox("Gesucht wird mit folgender Suchfrage:",search + searchTerm,"")
			application.activeWindow.command(strCommand,true);
			winID2 = application.activeWindow.windowID;
			status  = application.activeWindow.Status;
			scr = application.activeWindow.Variable("scr");
			count = application.activeWindow.Variable("P3GSZ");
			idn = application.activeWindow.Variable("P3GPP");
                                    if (vTag) tag = vTag
			// application.messageBox("", status + " " + scr + " " + count + " " + idn,"");
			if (status == "NOHITS")
				{
				erg1 = __Pruef("Nichts gefunden!","Gesucht wurde mit folgender Suchfrage: " + strCommand + "\n 1 = Suche unspezifisch wiederholen\n 2 = in Allegro weitersuchen\n 3 = Suchfrage von Hand kooriegieren\n4 = Suche beenden","1,2,3,4","1");
				if (erg1 == 1 || erg1 == 3)
					{ 
					if (erg1 == 3) strCommand = __Prompter("","",strCommand);
					if (erg1 == 1) strCommand = strCommand.substring(0,strCommand.indexOf(" and")) + " and bbg t*";
					application.activeWindow.closeWindow();
					}
				else if (erg1 == 2) 
					{
					strMeldung = "allegro";
					break Erfolg;
					}
				else if (erg1 == 4 || erg1 < 1) 
					{
					strMeldung = "ende";
					break Erfolg;
					}
				}
			if (count != "1" && scr == "7A" || scr == "GN") 
				{
 				if (tag.substring(0,3) == "532" || tag == "6710")
					{
					application.activeWindow.command("sor kls", false);
//					application.activeWindow.simulateIBWKey("FR");
					}
				if (scr == "GN") 
					{
					application.activeWindow.simulateIBWKey("FR");
					count = application.activeWindow.Variable("P3GSZ");
					}
				var prompt = utility.newPrompter();
				prompt.setDebug(true);
				if (prompt.confirmEx("Mehr als ein Satz gefunden!", "Ihre Suche ergab " + count + " Treffer.\nWählen Sie den richtigen Satz aus und klicken dann auf [Weiter]!", "OK", "Cancel", "", "", false) != 0)
					{
					strMeldung = "ende";
					}
				else
					{
					application.pauseScript();
					}
				}
			}
		while (status == "NOHITS")
	
		// Erfogsfall
		if (strMeldung != "ende" && strMeldung != "allegro")
			{
			if (application.activeWindow.Variable("scr") == "7A") application.activeWindow.simulateIBWKey("FR");
			var prompt = utility.newPrompter();
			prompt.setDebug(true);
			if (prompt.confirmEx("Gefunden!", "Ihre Suche ergab den angezeigten Treffer.\nSoll wirklich mit diesem Datensatz verknüpft werden?", "YES", "NO", "", "", false) == 0)
				{
				application.activeWindow.command("k",false);
				if (application.activeWindow.status != "ERROR")
					{
					test = application.activeWindow.title.findTag("005",0,true,true);
					if (test != "")
						{
						// Nutzungskenzeichen
						var kz = "";
						kz = application.activeWindow.title.findTag("012",0,true,true);
						if (kz != "")
							{
							application.activeWindow.title.startOfField(false);		//gehe zum Start des Feldes	
							application.activeWindow.title.endOfField(true);		//Selektiere von der Position alles bis zum Ende
							var kz = application.activeWindow.title.Selection;		//merke den Suchbegriff
							// application.messageBox("", kz,"");
							if (kz.indexOf("o") == -1) 
								{
								application.activeWindow.title.endOfField(false);
								application.activeWindow.title.insertText(";o");
								}
							}
						else
							{
							application.activeWindow.title.endOfField(false);
							application.activeWindow.title.insertText("\n012 o");
							}
						// Teilbestandskenzeichen
						application.activeWindow.title.findTag("011",0,true,true);
						application.activeWindow.title.startOfField(false);		//gehe zum Start des Feldes	
						application.activeWindow.title.endOfField(true);		//Selektiere von der Position alles bis zum Ende
						var kz = application.activeWindow.title.Selection;		//merke den Suchbegriff
				    		if (tag.substring(0,1) == "3") 
				    			{
							if (kz.indexOf("f") == -1) 
								{
								application.activeWindow.title.endOfField(false);
								application.activeWindow.title.insertText(";f");
								}
				    			}
					    	else if (tag.substring(0,2) == "51") 
					    		{
							if (kz.indexOf("s") == -1) 
								{
								application.activeWindow.title.endOfField(false);
								application.activeWindow.title.insertText(";s");
								}
				    			}
					    	else if (tag.substring(0,3) == "559" || tag.substring(0,2) == "68") 
					    		{
							if (kz.indexOf("g") == -1) 
								{
								application.activeWindow.title.endOfField(false);
								application.activeWindow.title.insertText(";g");
								}
				    			}
				    		}
					application.activeWindow.simulateIBWKey("FR");
				    	strIDN = application.activeWindow.Variable("P3GPP");
			    		}
		    		else
		    			{
					application.activeWindow.simulateIBWKey("FE");
				    	strIDN = application.activeWindow.Variable("P3GPP");
			    		}
				application.closeWindow(winID2);
                                                if (vTag) return strIDN
//				application.activeWindow.closeWindow();
				application.activeWindow.title.startOfField(false);		//gehe zum Start des Feldes	
				application.activeWindow.title.charRight(4,false);		//gehe vier zeichen nach rechts, um den tag auszuschließen
				if (strMarkierung != "0") application.activeWindow.title.endOfField(true);		//Selektiere von der Position alles bis zum Ende
				application.activeWindow.title.insertText(" " + strEinleitung + "!" + strIDN + "!" + rest);
				}
			else
				{
				strMeldung = "ende"
				}				
			}
			// nichts gefunden
			if (strMeldung == "allegro")
				{
				application.closeWindow(winID2);
				// in allegro nachschauen, ob dort ein relevanter Normsatz existiert und dessen Ansetzung + IDN im Fehlerfeld festhalten
				var prompt = utility.newPrompter();
				prompt.setDebug(true);
				if (prompt.confirmEx("Nichts gefunden!", "Soll in allego-BIB weiter gesucht werden?", "Yes", "No", "", "", false) == 0) 
					{
					__FehlerRegistrierung(erg, content);
					}
				}
			if (strMeldung == "ende")
				{
				application.activeWindow.closeWindow();
				}
		}
	}

}
function DBSMListen() {
	if (!application.activeWindow.Variable("scr") || application.activeWindow.Variable("scr") == "FI")
	{
	if (application.activeWindow.Variable("scr") == "FI") application.messageBox("Diese Funktion steht momentan nicht zur Verfügung!","Um die Funktion nutzen zu können\nmuss eine Ergebnismenge oder \nein einzelner Titel ausgewählt werden.","");
	else application.messageBox("Diese Funktion steht momentan nicht zur Verfügung!","Um die Funktion nutzen zu können\nmüssen Sie sich erst einloggen\nund eine Ergebnismenge bilden.","");
	return
	}
else
	{
            var Typ, titel, ea, datum, nr, sig, kommentar, teil, wert, liferant;
	var wert = ""
	var kommentar = ""
            var lieferant = ""
	// Warnung, wenn Ergebnismenge über 50
	if (application.activeWindow.Variable("P3GSZ") > 50)
		{
		var prompt = utility.newPrompter();
		prompt.setDebug(true);
		if (prompt.confirmEx("ACHTUNG","Die Ergebnismenge hat "+application.activeWindow.Variable("P3GSZ")+" Treffer!\nWirklich ausgeben?", "Yes", "No", "", "", false) != 0) return
		}
	// Setze Variablen und gehe zum Anfang
	application.activeWindow.command("\\TOO 1 D", false);	// falls es eine Ergebnsmenge ist, gehe zum ersten Titel und zeige diesen an
	var line = new Array()						// Array für die Zeilen
	var string = ""					                        // String für die Zellinhalte
	line[0] = "Nr."							// Vorbesetzten der ersten Spalte der Kopfzeile
	var m = 1; 								// Zähler für line[m]
	meldung = ""							// Staus zum Beenden der äußeren Schleife
	such = ""								// Suchstring innerhalb der Aufnahme
	such1 = ""								// exakter Suchstring innerhalb der Aufnahme
            bestand = "Geben Sie den Bestand ein!\nM/KD = Sammlung Künstlerische Drucke,\nF/Klemm = Klemmsammlung\nS = Studiensammlung\nusw.\nWenn nichts eingegeben wird,\nwird das jeweils erste DBSM-Exemplar ausgegeben."
            erg = ""
	// wähle den richtigen Listentyp
	erg1 = __Pruef("Welche Liste soll erstellt werden?","1 = Leihliste\n 2 = Zugangsbuch\n 3 = Buchbinder\n4 = Restaurierung\n5 = Kurzliste \n6 = Eigene Liste erstellen","1,2,3,4,5,6","1");
	// Liste für Leihvertrag
	if (erg1 == 1)
		{
                        Typ = "LL"
		erg = "4062;4019"
		such = "LV"
		such1 = "-Vormerkung"
		line[0] = line[0] + "\tSignatur\tTitel\tKommentar\tWert\tLeihnehmer\tDauer"
		}
	// Zugangsbuch
	else if (erg1 == 2) 
		{
                        Typ = "ZB"
		such = "DBSM/"
		such1 = __Prompter("Welche Exemplare soll die Liste enthalten?",bestand)
		line[0] = line[0] + "\tDatum\tTitel / Jahr\tLieferant\tErwerbungsart\tWert / Preis\tSignatur\tKommentar"
		}
	// Liste für Buchbinder / Liste für Restaurierung
	else if (erg1 == 3 || erg1 == 4) 
		{
                        Typ = "BB"
		if (erg1 == 3) such = "Bubi"
                        if (erg1 == 4) such = "Restaurierung"
		such1 = "-Vormerkung"
		line[0] = line[0] + "\tSignatur\tTitel\tKommentar\tDatum Rückgabe"
		}
	// wenn kein Listentyp gewählt wurde, soll selbst eine Liste zusammengestellt werden 
	// Zugangsbuch
	else if (erg1 == 5) 
		{
                        Typ = "KU"
		such = "DBSM/"
		such1 = __Prompter("Welche Exemplare soll die Liste enthalten?",bestand)
		line[0] = line[0] + "\tSignatur\tKurztitel"
		}
	else if (erg1 == 6) 
		{
                        Typ = "XX"
		erg = __Prompter("Welche Felder soll die Liste enthalten?","Geben Sie die Feldnummern getrennt durch Semikolon ein!")
		if (!erg || erg == "") return
		if (erg.indexOf("710") > -1)
			{
			such = "DBSM/"
			such1 = __Prompter("Welche Exemplare soll die Liste enthalten?",bestand)
			}
		}
	else
		{
		return
		}
	merk = erg	// Kette merken
	// merk und erg enthalten die vollständige Kette der Felder für die Liste
	// erg wird in den beiden inneren Schleifen (für die Kopfzeile und für die satzweise Abarbeitung) aufgefressen und nach der Schleife aus merk wieder neu gebildet

	// Kopfzeile bilden
            while (erg != "")								// solange Feldliste noch Inhalt hat
                {
	    if (erg.indexOf(";") != -1)						// wenn es mehre Felder in Feldliste gibt:
		{ 
		tag = erg.substring(0,erg.indexOf(";"))				// Spalte wird erstes Feld aus Feldliste
		erg = erg.substring(erg.indexOf(";")+1,erg.length)		// Feldliste wird Feldliste minus erstes Feld
		}
	    else										// sonst:
		{
		tag = erg									// Spalte wird einziges Feld aus Feldliste
		erg = ""									// Feldliste wird leer
		}
	    line[0] = line[0] + "\t" + tag						// neue Spalte an Kopfzeile anfügen
	    }
	erg = merk							// gemerkte Kette neu laden
            application.activeWindow.clipboard = line[0]	                        // Kopfzeile an Zwischenablage übergeben
	// einzelne Sätze abarbeiten
	do		// Ziel: je Satz die Feldinhalte mit Tabulator getrennt in eine eigene Zeile
		{
		application.activeWindow.command("k", false);
		if (application.activeWindow.status == "ERROR")
			{
			application.messageBox("Diese Funktion steht momentan nicht zur Verfügung!","Um die Funktion nutzen zu können\nmüssen Sie Bearbeitungsrechet für die Ergebnismenge haben.","");
			return
			}
		if (Typ != "ZB") line[m] = m;
                        // Titel Verlag und Jahr wird immer gebraucht
                        titel = application.activeWindow.title.findTag("4000",0,false,true)
                        // Wenn es ein *f-Satz ist, muss der Titel noch aus dem *c-Satz nachgeladen werden
                        if (titel.indexOf("!") > -1)
                            {
                            idn = titel.substring(titel.indexOf("!")+1,titel.length -1)
                            application.activeWindow.command("f idn " + idn,true);
		    application.activeWindow.command("k", false);
                            titel = application.activeWindow.title.findTag("4000",0,false,true)
                            application.activeWindow.simulateIBWKey("FE");
		    application.activeWindow.closeWindow();
                            }
                        // Titel des Bandes anhängen
                        test = application.activeWindow.title.findTag("4004",0,false,true)
                        if (test != "") 
                            {
                            titel = titel + ". - " + test.substring(1,test.length)
                            test = ""
                            }
                        // Entstehungsangabe anhängen 
                        test = application.activeWindow.title.findTag("4046",0,false,true)
                        if (test == "") test = application.activeWindow.title.findTag("4030",0,false,true)
                        if (test == "") test = application.activeWindow.title.findTag("4045",0,false,true)
                        if (test != "") 
                            {
                            titel = titel + ". - " + test
                            test = ""
                            }
                        // Entstehungszeit anhängen
                        test = application.activeWindow.title.findTag("1110",0,false,true)
                        if (test != "") 
                            {
                            if (test.indexOf("*") > 0) titel = titel + ", " + test.substring(0,test.indexOf("*"))
                            else titel = titel + ", " + test.substring(1,test.indexOf("$"))
                            }
                        else
                            {
                            test = application.activeWindow.title.findTag("1100",0,false,true)
                            if (test.indexOf("$n") > 0) titel = titel + ", " + test.substring(test.indexOf("$n")+2,test.length)
                            else titel = titel + ", " + test
                            }
                        // Titel ist fertig, Test kann gelöscht werden 
                        test = ""
		if (such != "")
			{
			// zunächst richtiges Exemplar suchen und bearbeiten
			test = application.activeWindow.title.find(such,false, false, false)
                                    // die Feldnummer ermitteln
                                    tagnr = application.activeWindow.title.Tag
                                    // wenn mit such etwas gefunden wurde, ist tagnr 7109 oder 4821, sonst ist hier schluss
                                    if (tagnr != "7109" && tagnr != "4821" ) 
                                        {
                                        application.messageBox("Abbruch","Bei mindestens einem Titel fehlte das Feld 7109 oder das Feld 4821.","")
                                        application.activeWindow.simulateIBWKey("FE");									// Bearbeitungsmodus beenden (Esc)
                                        return
                                        }
                                    // im Exemplarsatz an den Anfang gehen 
                                    while (tagnr.substring(0,3) != "700")
                                        {
                                        application.activeWindow.title.lineUp(1)
                                        tagnr = application.activeWindow.title.Tag
                                        }
                                    application.activeWindow.title.findTag2("700",0, false, true);
			if (such != "" && test != "")
				{
                                                // richtige 4821 finden
                                                // --> für Zugangsbuch (hier steht "DBSM" in such, deshalb $z hier explizit angegeben)
                                                if (Typ == "ZB") 
                                                    {
                                                    application.activeWindow.title.find("$zErwerbung",false, false, false);	//falls es mehr als ein 4821 gibt, gehe zu dem mit "Erwerbung"
                                                    }
                                                // --> für Leihliste / Buchbinder / Restaurierung (hier steht $z in such)
                                                else if (Typ == "BB" || Typ == "LL") test = application.activeWindow.title.find("$z" + such ,false, false, false);	//falls es mehr als ein 4821 gibt, gehe zu dem mit dem Inhalt von such
				// --> bei Listen mit Vormerkung, die Vormerkung entfernen
				if (such1 == "-Vormerkung")
                                                        {
			                    test = application.activeWindow.title.find(such1,false, false, false)	//falls es mehr als ein 4821 gibt, gehe zu dem mit "-Vormerkung"
				        if (test > "") application.activeWindow.title.deleteSelection()			//"-Vormerkung" löschen
                                                        }
				// zugehörige Exemplardaten auswerten
				application.activeWindow.title.startOfField(false);			//gehe zum Start des Feldes	
				application.activeWindow.title.endOfField(true);		            //markiere Feld 4821
                                                test = application.activeWindow.title.Selection;                            //merke Feld 4821 in test
    				application.activeWindow.title.startOfField(false);			//gehe zum Start des Feldes
 				if (test.indexOf("$l") > -1)
					{
					lieferant = test.substring(test.indexOf("$l") + 2, test.length)
                                                            if (lieferant.indexOf("$") > 0) lieferant = lieferant.substring(0,lieferant.indexOf("$"))
					}
 				if (test.indexOf("$t") > -1)
					{
					teil = test.substring(test.indexOf("$t") + 2, test.length)
                                                            if (teil.indexOf("$") > 0) teil = teil.substring(0,teil.indexOf("$"))
					}
 				if (test.indexOf("$K") > -1)
					{
					kommentar = test.substring(test.indexOf("$K") + 2, test.length)
                                                            if (kommentar.indexOf("$") > 0) kommentar = kommentar.substring(0,kommentar.indexOf("$"))
					}
 				if (test.indexOf("$D") > -1)
 					{
					sJahr = test.substring(test.indexOf("$D") + 2, test.indexOf("$D") + 6)
				            sMonat = test.substring(test.indexOf("$D") + 7, test.indexOf("$D") + 9)
					sTag = test.substring(test.indexOf("$D") + 10, test.indexOf("$D") + 12)
                                                            datum = sTag + "." + sMonat + "." + sJahr
					}
 				if (test.indexOf("$E") > -1)
 					{
					sJahr = test.substring(test.indexOf("$E") + 2, test.indexOf("$E") + 6)
					sMonat = test.substring(test.indexOf("$E") + 7, test.indexOf("$E") + 9)
					sTag = test.substring(test.indexOf("$E") + 10, test.indexOf("$E") + 12)
                                                            datum = datum + "-" + sTag + "." + sMonat + "." + sJahr
					}
				if (test.indexOf("$z") > -1)
					{
					zweck = test.substring(test.indexOf("$z") + 2, test.length)
                                                            if (zweck.indexOf("$") > 0) zweck = zweck.substring(0,zweck.indexOf("$"))
					}
				if (test.indexOf("$w") > -1)
					{
					wert = test.substring(test.indexOf("$w") + 2, test.length)
                                                            if (wert.indexOf("$") > 0) wert = wert.substring(0,wert.indexOf("$"))
					}
				if (test.indexOf("$q") > -1)
					{
					quelle = test.substring(test.indexOf("$q") + 2, test.length)
                                                            if (quelle.indexOf("$") > 0) quelle = quelle.substring(0,quelle.indexOf("$"))
					}
                                                test = ""
                                                if (Typ == "ZB" || Typ == "BB" || Typ == "LL" || Typ == "KU")
                                                    {
                                                    sig =  application.activeWindow.title.findTag("7100",0,false,true)
                                                    if (sig.indexOf("@") > -1) sig = sig.substring(0,sig.indexOf("@"))
                                                    if (Typ == "ZB")
                                                        {                                                        
                                                        ea = application.activeWindow.title.findTag("8510",0,false,true)
                                                        nr = application.activeWindow.title.findTag("8598",0,false,true)
                                                        line[m] = nr.substring(nr.indexOf("/")+1,nr.length)
                                                        if (datum) string = string + datum;
                                                        string = string + "\t" + titel
                                                        string = string + "\t"
                                                        if (lieferant) string = string + lieferant;    
                                                        string = string + "\t" + ea.substring(ea.indexOf("%")+1,ea.length-1)
                                                        string = string + "\t"
                                                        if (wert) string = string + wert;
                                                        string = string + "\t" + sig
                                                        string = string + "\t"
                                                        if (kommentar) string = string + kommentar;
                                                        if (teil) string = string + teil;
                                                        }
                                                    else
                                                        {
                                                        string = sig
                                                        string = string + "\t" + titel
                                                        if (teil) string = string + ", " + teil;
                                                        if (Typ != "KU")
                                                            {
                                                            string = string + "\t"
                                                            if (kommentar) string = string + kommentar;
                                                            }
                                                        }
                                                    if (Typ == "LL")
                                                        {
                                                        string = string + "\t"
                                                        if (wert) string = string + wert;
                                                        string = string + "\t" + zweck.substring(4, zweck.length)
                                                        string = string + "\t"
                                                        if (datum) string = string + datum;
                                                        }
                                                    ea = ""
                                                    datum = ""
                                                    nr = ""
                                                    sig = ""
                                                    titel = ""
                                                    kommentar = ""
                                                    teil = ""
                                                    wert = ""
                                                    liferant = ""
                                                    }
                                                line[m] = line[m] + "\t" + string;
                                                string = ""
	                                }
			}
		// jetzt alle Tags abarbeiten
                        while (erg != "")
			{
			if (erg.indexOf(";") != -1)
				{ 
				tag = erg.substring(0,erg.indexOf(";"))
				erg = erg.substring(erg.indexOf(";")+1,erg.length)
//				application.messageBox("Kette der Feldnummern", erg + "\n" + tag,"");
				}
			else
				{
				tag = erg
				erg = ""
				}
			if (tag != "7100")
				{ 
				string = application.activeWindow.title.findTag(tag,0,false,true)
                                                if (tag=="4019" || tag=="4062") string = string.substring(0,string.indexOf("$"))
                                                }
			else
				{ 
				application.activeWindow.title.find(such+such1,false, false, false)
				application.activeWindow.title.lineUp(3, false);
				application.activeWindow.title.startOfField(false);		//gehe zum Start des Feldes	
				string = application.activeWindow.title.find(tag,false,false,false)
				}
			if (string == "") string = "xxx Fehler: " + tag + " nicht vorhanden!";
			line[m] = line[m] + "\t" + string;
//			application.messageBox("Zwischenergebnis", line[m] + "\n" + string,"");
                                    string = ""
			}
		if (Typ == "ZB" || Typ == "BB" || Typ == "LL") 
                            {
                            application.activeWindow.simulateIBWKey("FR");				// Bearbeitungsmodus beenden (Enter)
                            if (application.activeWindow.status == "ERROR")
			{
			application.messageBox("Diese Funktion kann nicht fortgesetzt werden!","Der Datensatz konnte nicht gespeichert werden.","");
			return
			}
                            }
		else application.activeWindow.simulateIBWKey("FE");									// Bearbeitungsmodus beenden (Esc)
		test = application.activeWindow.simulateIBWKey("F1");									// zum nächsten Satz gehen
//		application.activeWindow.clipboard = application.activeWindow.clipboard + "\n" + line[m]		// Zeile an Zwischenablage anhängen
		erg = merk															// Kette neu laden
		m = m + 1															// Zeilenzähler um 1 hochsetzen
		if (application.activeWindow.messages.item(0) == "Dies ist der letzte Titel") meldung = "ende"	// falls der letzte Satz erreicht ist, Schleife beenden
		}
	while (meldung != "ende")
	}
            m = m - 1
            DBSMdruck(line, m)
//	var oShell = new ActiveXObject("Shell.Application"); 
//	var commandtoRun = "V:\\06_DBSM\\02_Erschließung\\01_allgemein\\WinIBW\\Starteinstellung\\word-start.cmd";  
//	oShell.ShellExecute(commandtoRun,"","","open","1"); 
//	application.messageBox("Ergebnis", application.activeWindow.clipboard ,"");	
}
function DBSMPortal() {
// Voraussetzung: Eine Ergebnismenge bzw. ein einzelner Datensatz ist geladen
// Ziel: Alle im Satz vorhandenen externen Links werden zusammengetragen und zur Auswahl angeboten. Nach Auswahl wird der Internet-Browser gestartet und der externe Link ausgeführt. 
// Existiert kein externer Link, so wird der Datensatz im Portal angezeigt.

// Setze Vorgabewerte 
	var adresse = new Array()									// Array für die Aufrufadressen
	var prae = "http://d-nb.info/" 								// Präfix für Aufrufadressen vorbelegen
	var url = "" 												// URL vorbelegen						
	var sprache = ""											// Sprache vorbelegen
	var format = ""												// Format vorbelegen
	var art = ""												// Art (nach ONIX-Codelist) vorbelegen
	var herkunft = ""											// Herkunft vorbelegen
	var code = ""												// Code vorbelegen
	var strIDN = application.activeWindow.Variable("P3GPP");	// IDN auslesen
	if (!strIDN) 												// wenn nicht erfolgreich, abbrechen
		{
		application.messageBox("Diese Funktion steht momentan nicht zur Verfügung!","Es muss erst ein Titel ausgwählt sein.","");
		}
	else														// sonst 
		{
		application.activeWindow.command("k", false);									// gehe in Bearbeitungsmodus
		var test = application.activeWindow.title.findTag("005",0,true,false,false);	// an Hand von 005 testen ob Normsatz
		if (test) 																		// wenn Normsatz  
			{
			if (test.indexOf("Tk") > 0 || test.indexOf("Tq") > 0) 								// wenn Tq- / Tk-Satz
				{
				prae = "http://d-nb.info/dnbn/"														// Präfix für DNB-Normsatz merken
				}
			else																				// sonst
				{
				prae = application.activeWindow.title.findTag("006",0,false,false,false);			// steht die Aufrufadresse vollständig in Feld 006
				strIDN = ""																			// und strIDN wird nicht mehr gebraucht
				}
			}
											// Test ob Felder für externen Link vorhanden
		var test1 = application.activeWindow.title.findTag("4083",0,true,true,false);
		if (!test1) var test1 = application.activeWindow.title.findTag("4085",0,true,true,false);
		if (!test1) var test1 = application.activeWindow.title.findTag("4715",0,true,true,false);
		if (!test1) var test1 = application.activeWindow.title.findTag("670",0,true,true,false);
		test2 = test1.substring(0,4)
		if (!test1)		// wenn weder 4083 noch 4085 noch 4715 noch 670 vorhanen, sofort Explorer starten und Satz im Portal anzeigen
			{
			application.activeWindow.simulateIBWKey("FE");		// Bearbeitungsmodus beenden (ESC)
			application.shellExecute(prae + strIDN);			// Explorer starten
			}
		else			// Sonst Array Adresse[n] mit weiteren Adressen und ask[n] mit entsprechenden Aufrufen befüllen 
			{
			ask = "Welche Seite soll aufgerufen werden?"			// ask mit Frage vorbelegen
			adresse[0] = prae + strIDN;					// Adresse[0] mit Portalaufruf des Datensatzes belegen
			n = 0								// Zähler für Adresszeilen und Aufrufe auf 0 setzen
			n1 = 0 								// Zähler für erste innere Abfrage mit 0 vorbesetzen
			n2 = 0 								// Zähler für zweite innere Abfrage mit 0 vorbesetzen
			n3 = 0 								// Zähler für dritte innere Abfrage mit 0 vorbesetzen
			ask = ask + "\n " + n + ": dieser Datensatz im Portal"	// erste Aufruf-Zeile in ask festlegen
			count = "0,"							// Zähler für Zeile im Input-Feld (erg) mit 0 vorbelegen
			Schleife01: while (test1)					// solange weitere Felder mit externen Links vorhanden sind
				{
				if (n > 0)						// wenn es nicht der erste Durchlauf der Schleife ist
					{
											// auf weiteres Feld 4083 testen
					test1 = application.activeWindow.title.findTag("4083",n,true,true); 
					if (!test1) 					// wenn nicht (test1 nicht besetzt)
						{					// prüfen, ob letztes mal Feld 4083 berabeite,
						if (test2 == "4083") n1 = 0 		// wenn ja, internen Zähler wieder auf 0 setzen
						else n1 = n1 + 1			// sonst internen Zähler um 1 erhöhen 
											// auf weiteres Feld 4085 testen
						test1 = application.activeWindow.title.findTag("4085",n1,true,true);
						}
					if (!test1)  					// wenn test1 noch nicht besetzt 
						{					// prüfen, ob letztes mal Feld 4083 oder 4085 berabeite,
						if (test2 == "4083" || test2 == "4085") n2 = 0	// wenn ja, internen Zähler wieder auf 0 setzen
						else n2 = n2 + 1 			// sonst internen Zähler um 1 erhöhen
											// auf weiteres Feld 4715 testen
						test1 = application.activeWindow.title.findTag("4715",n2,true,true);
						}
					if (!test1)  					// wenn test1 noch nicht besetzt 
						{					// auf weiteres Feld 670 testen
						n3 = n3 + 1 				// sonst internen Zähler um 1 erhöhen
											// auf weiteres Feld 670 testen
						test1 = application.activeWindow.title.findTag("670",n2,true,true);
						if (!test1) break Schleife01;
						}
					test2 = test1.substring(0,4)			// in test2 Feldnummer für Vergleich in nächster Runde merken				
					}
									// alle Startpunkte der Teilfelder wieder auf nicht vorhanden (-1) zurück setzen
				u = -1
				a = -1
				b = -1
				c = -1
				d = -1
				e = -1
				x = -1
				z = -1
				u = test1.indexOf("=A ")				// Startpunkt für url aus Teilfeld A ermitteln
				if (u == -1) u = test1.indexOf("=u ")		// wenn TF A nicht vorhanden, Startpunkt für url aus TF u ermitteln 
				if (u == -1) u = test1.indexOf("$u")-1		// wenn u noch nicht definiert (Normsatz), Startpunkt für url aus TF u ermitteln 
				a = test1.indexOf("=a ")				// Startpunkt für sprache aus Teilfeld a ermitteln
				b = test1.indexOf("=b ")				// Startpunkt für format aus Teilfeld b ermitteln
				c = test1.indexOf("=c ")				// Startpunkt für art aus Teilfeld c ermitteln
				d = test1.indexOf("=d ")				// Startpunkt für herkunft aus Teilfeld d ermitteln
				e = test1.indexOf("=e ")				// Startpunkt für code aus Teilfeld e ermitteln
				x = test1.indexOf("=x ")				// Startpunkt für ??? aus Teilfeld x ermitteln
				z = test1.indexOf("=z ")				// Startpunkt für ??? aus Teilfeld z ermitteln
				if (u != -1 && a != -1) url = test1.substring(u+3,a)	// wenn TF a vorhanden, steht url zwischen u und a
				else if (u > -1 && b != -1) url = test1.substring(u+3,b)	// wenn TF b vorhanden, steht url zwischen u und b
				else if (u > -1 && c != -1) url = test1.substring(u+3,c)	// wenn TF c vorhanden, steht url zwischen u und c
				else if (u > -1 && d != -1) url = test1.substring(u+3,d)	// wenn TF d vorhanden, steht url zwischen u und d
				else if (u > -1 && e != -1) url = test1.substring(u+3,e)	// wenn TF e vorhanden, steht url zwischen u und e
				else if (u > -1 && x != -1) url = test1.substring(u+3,x)	// wenn TF x vorhanden, steht url zwischen u und x
				else if (u > -1 && z != -1) url = test1.substring(u+3,z)	// wenn TF z vorhanden, steht url zwischen u und z
				else if (u > -1) url = test1.substring(u+3)			// sonst steht url zwischen u und Ende
				if (!url)		 					// wenn noch keine URL gefunden, steht auch keine drin
					{
					n = n+1
					continue; 						// nächste Runde in Schleife starten
					}
				if (a != -1 && b != -1) sprache = test1.substring(a+3,b)		// wenn TF b vorhanden, steht sprache zwischen a und b
				else if (a != -1 && c != -1) sprache = test1.substring(a+3,c)	// wenn TF c vorhanden, steht sprache zwischen a und c
				else if (a != -1 && d != -1) sprache = test1.substring(a+3,d)	// wenn TF d vorhanden, steht sprache zwischen a und d
				else if (a != -1 && e != -1) sprache = test1.substring(a+3,e)	// wenn TF e vorhanden, steht sprache zwischen a und e
				else if (a != -1) sprache = test1.substring(a+3)			// sonst steht sprache zwischen a und Ende
				if (b != -1 && c != -1) format = test1.substring(b+3,c)		// wenn TF c vorhanden, steht format zwischen b und c
				else if (b != -1 && d != -1) format = test1.substring(b+3,d)	// wenn TF d vorhanden, steht format zwischen b und d
				else if (b != -1 && e != -1) format = test1.substring(b+3,e)	// wenn TF e vorhanden, steht format zwischen b und e
				else if (b != -1) format = test1.substring(b+3)			// sonst steht format zwischen b und Ende
				if (test1.substring(0,4) == "4083") art = "34"			// wenn Feld 4083, ist art "34" (=Archivserver)
				else if (c != -1 && d != -1) art = test1.substring(c+3,d)		// wenn TF d vorhanden, steht art zwischen c und d
				else if (c != -1 && e != -1) art = test1.substring(c+3,e)		// wenn TF d vorhanden, steht art zwischen c und e
				else if (c != -1) art = test1.substring(c+3)				// sonst steht art zwischen c und Ende
				if (d != -1 && e != -1) herkunft = test1.substring(d+3,e)		// wenn TF e, steht herkunft zwischen d und e
				else if (d != -1) herkunft = test1.substring(d+3)			// sonst steht herkunft zwischen d und Ende
				if (e != -1) code = test1.substring(e+3)				// code steht zwschen e und Ende
				n = n+1									// Zähler für Adresszeilen und Aufrufe um 1 erhöhen
				count = count + n + ","						// Zähler für Zeile im Input-Feld (erg) mit n belegen
				if (url == "$") 							// wenn in url nur "$", ist es ein DN-interner Aufruf
					{
					adresse[n] = prae + strIDN + "/" + art			// also Adresse[n] aus Präfix, IDN und Art bilden
					if (art == "04") ask = ask + "\n " + n + ": Inhaltsverzeichnis im Portal"		// wenn Art = 04, ist es das Inhaltsverzeichnis
					else if (art == "34") ask = ask + "\n " + n + ": Seite im Archivserver der DNB"	// wenn Art = 34, ist es der Archivserver
					else ask = ask + "\n " + n + ": " + adresse[n]		// sonst die Adresse in Aufruf-Zeile ausgeben
					}
				else if (url.substring(0,4) == "http" || url.substring(0,3) == "www")	// wenn in url eine HTTP-Adresse steht
					{
					adresse[n] = url						// mit dieser Adresse[n] belegen
					if (test1.substring(0,3) == "670") ask = ask + "\n " + n + ": Quelle - " + adresse[n]		// und diese in der Aufrufzeile ausgeben
					else ask = ask + "\n " + n + ": " + adresse[n]		// und diese in der Aufrufzeile ausgeben
						}
				else									// sonst ist es wahrscheinlich ein Portalaufruf
					{
					adresse[n] = "http://portal.dnb.de/" + url			// dann Adresse[n] mit Portal-Präfix und url belegen
					ask = ask + "\n " + n + ": " + adresse[n]			// und diese in der Aufrufzeile ausgeben
					}
				}
			application.activeWindow.simulateIBWKey("FE");			// Bearbeitungsmodus beenden (ESC)
			erg = __Pruef("Eine Frage",ask,count,"0");			// Frage mit Input-Feld (erg) ausgeben
			if (erg) application.shellExecute(adresse[erg]); 	// und Explorer mit gewählter Adresse starten
			}

		}
}
function DBSMSucheErsetze() {
var line = new Array()						// Array für die Zeilen
line[0] = "Nr\tIDN\tFehler"							// Vorbesetzten der ersten Spalte der Kopfzeile
var m = 0; 								// Zähler für line[m]
var tag = "";
var search = "";
var movetext = "";
var addtext = "";
var ask = ""
var text = ""
var stop = -1
count = 1;																	//count mit 1 vorbelegen
count = application.activeWindow.Variable("P3GSZ");							//Anzahl der Treffer auf count
count1 = application.activeWindow.Variable("P3LNR");						//aktueller Satz auf count1
if (count1 != 1) var count = count - count1									//wenn es nicht der erste Satz ist, count auf Anzahl der restlichen Sätze verkleinern
if (count != 1)																//wenn count nicht 1 ist abfragen ob alle Sätze oder nur der aktuelle
	{
	if (count == 0) var count = 1;
	else
		{
		var prompt = utility.newPrompter();
		prompt.setDebug(true);
		ask = prompt.confirmEx("ACHTUNG","Die Ergebnismenge hat noch "+count+" Treffer!\nFunktion auf die restliche Ergebnismenge anwenden?", "Yes", "No", "", "", false);
		if (ask != 0) count = 1;
		}
	} 
if (application.activeWindow.Variable("title") == "Kurzanzeige") application.activeWindow.simulateIBWKey("FR");

tag = __Prompter("Eine Frage","In welchem Feld soll gesucht werden?","","");
if (tag == null) return
search = __Prompter("Eine Frage","Wonach soll gesucht werden?","","");
if (search == null) return
movetext = __Prompter("Eine Frage","Womit soll ersetzt werden?","","");
if (movetext == null) return
if (movetext == "") addtext = __Prompter("Eine Frage","Sie haben keinen Text zum Ersetzen eingegeben.\nWas soll statt dessen erfasst werden?","","");
if (addtext == null) return
//if (addtext == "" && movetext == "") return;

if (movetext != "")
	{
	if (tag != "") ask = "1. Ersetze [" + search + "] durch [" + movetext + "] in Feld [" + tag + "]."
	else if (search != "") ask = "2. Ersetze im gesamten Datensatz [" + search + "] durch [" + movetext + "]."
            else ask = "7. Ergänze Feld [" + movetext + "]."
	}
else
	{
	if (tag != "") ask = "3. Ergänze [" + addtext + "] wenn [" + search + "] in Feld [" + tag + "] vorkommt."
	else if (search != "" && addtext != "") ask = "4. Ergänze [" + addtext + "] wenn [" + search + "] im Datensatz vorkommt."
	else if (addtext != "") ask = "5. Ergänze [" + addtext + "] in allen Datensätzen."
            else ask = "6. Lösche [" + search + "] in allen Datensätzen."
            if (addtext.substring(0,4) == "6710") search = __Prompter("Eine Frage","Zu welcher Signaturengruppe soll 6710 ergänzt werden?","","");
	}
var __ask = utility.newPrompter();
erg = __ask.confirm("Letzte Frage","Soll folgende Aktion wirklich gestartet werden?\n" + ask);			// Frage mit Input-Feld (erg) ausgeben
if (erg == false) return;
ask = ask.substring(0,1)


a = 1;
while (a <= count)											//vom ersten bis zum letzten Satz
	{
	application.activeWindow.simulateIBWKey("F7");						//Bearbeiten ein
	test = "";
	if (ask == 1 || ask == 2)
		{
		if (ask == 1)
			{
			vor = "";
			rest = "";
			n = 0
			test = application.activeWindow.title.findTag(tag,n,false,true,true);		//suche Feld 
			while (test != "")																//solange etwas gefunden wurde
				{
				test = application.activeWindow.title.findTag(tag,n,false,true,true);		//suche Feld 
				if (test.indexOf(search) > -1)												//wenn darin Suchbegriff gefunden
					{
				            if (search=="")
                                                                {
                                                                vor = test
                                                                }
                                                           else
                                                                {                              
					    vor = test.substring(0,test.indexOf(search))								//Text vor Suchbegriff merken
					    rest = test.substring(test.indexOf(search) + search.length,test.length)		//Text nach Suchbegriff merken
                                                                }
					text = vor + movetext + rest												//Text mit Ersetze-Text neu zusammenbauen   
					if (tag == "8598")															//wenn es eine Zugangsnummer in 8598 ist
						{
						vor = text.substring(0,text.indexOf("/"))
						rest = text.substring(text.indexOf("/")+1,text.length)
						while (rest.length < 5) rest = "0" + rest
						text = vor + "/" + rest
						}
					if (test != "") application.activeWindow.title.insertText(text);
					test1 = "ok"                                  
					}
				n = n+1
				}
                                    if(test1) test = test1
			}
	            else
			{
			text = movetext
			test = application.activeWindow.title.find(search,false,false,false);		//suche Feld 
			while (test == true)
				{
				application.activeWindow.title.insertText(text);
				test = ""
				test = application.activeWindow.title.find(search,false,false,false);		//suche Feld
				}
			test = "ok"                                  
			}
		}
	else 
		{
		if (ask == 3)
			{
                                    if (addtext == " @ m")
                                        {
                                        application.activeWindow.title.findTag(tag,0,false,true,true);		//suche Feld
                                        application.activeWindow.title.endOfField()                                          //gehe ans Ende des Feldes
			    test = "ok"                                  
                                        }
                                    else
                                        {
			    application.activeWindow.title.findTag(tag,0,false,true,true);		//suche Feld 
			    test = application.activeWindow.title.find(search,false,true,false);		//suche Feld 
                                        }
			}
		if (ask == 6 || ask == 4) test = application.activeWindow.title.find(search,false,false,false);		//suche zu ersetzenden / löschenden Text 
		else test = "ok"
		if (addtext == " @ m") text = addtext
                        else text = "\n" + addtext
//application.messageBox("",test,"")
		if (ask == 5) 
			{
			if (text.substring(1,5) == "6710")
				{
				n = 0
				while (test != "")
					{
					test = ""
					test1 = ""
					if (text.substring(1,5) == "6710")
						{
						test = application.activeWindow.title.findTag("7100",n,false,true);		//suche nächstes Feld 7100
//						test1 = application.activeWindow.title.find(search,false,true,false);		//Signaturrumpf
						application.activeWindow.title.endOfField()                                                      //gehe ans Ende des Feldes
//						if (test != "" && test1 != "") application.activeWindow.title.insertText(text);         //wenn Feld 7100 Signaturrumpf enthält, füge 6710 ein
						if (test.substring(0,search.length) == search) application.activeWindow.title.insertText(text);         //wenn Feld 7100 Signaturrumpf enthält, füge 6710 ein
						n = n+1
//application.messageBox("Frage", "ask = " + ask + "\nsearch = " + search + "\ntext = " + text + "\ntest = " + test + "\ntest1 = " + test.substring(0,search.length), "");
						}
					}
                                                }
			else
				{
//application.messageBox("Frage", "ask = " + ask + "\nsearch = " + search + "\ntext = " + text + "\ntest = " + test, "");

				if (addtext != " @ m") 
                                                    {
                                                    application.activeWindow.title.endOfBuffer()
                                                    application.activeWindow.title.insertText(text);
                                                    }
				}
			test = "ok"                                  
			}
		}
	if (test != "") 
		{
		if (ask != 1 && ask != 2 && ask != 5)
			{
                                    if (search.substring(0,4)=="6800")
                                       {
                                       application.activeWindow.title.findTag("6710",0,false,true);
                                       application.activeWindow.title.endOfField()
                                       }
			if (text != "\n") application.activeWindow.title.insertText(text);
			else if (search.substring(5,6) == "!") application.activeWindow.title.deleteLine();
			else application.activeWindow.title.deleteSelection();
//application.messageBox("Frage", "ask = " + ask + "\nsearch = " + search + "\ntext = " + text + "\ntest = " + test, "");
			}
		application.activeWindow.simulateIBWKey("FR");								//speichern
		if (application.activeWindow.Status != "OK")								//wenn nicht erfolgreich
			{
//application.messageBox("",m + "\n" +  application.activeWindow.Variable("P3GPP") + "\n" +  application.activeWindow.messages.item(0),"")
			m = m+1
			line[m] = m + "\t" +  application.activeWindow.Variable("P3GPP") + "\t" +  application.activeWindow.messages.item(0)
			application.activeWindow.simulateIBWKey("FE");								//Bearbeiten abbrechen
			}
		}
	else application.activeWindow.simulateIBWKey("FE");								//Bearbeiten abbrechen
	if (count > 1) 							//wenn Ergebnismenge größer 1
		{
		if (stop == -1)
			{
			stop = __Pruef("Wie soll's weiter gehen?","Geben Sie die Anzahl der Datensätze ein\ndie im nächsten Schritt bearbeitet werden sollen!\n 0 = Abbruch \n beliebige Zahl = nächste n Sätze \n x = gesamte Ergebnismenge","1,2,3,4,5,6,7,8,9,10,20,30,40,50,60,70,80,90,100,x","1");
			}
		if (stop == null) 
			{
			DBSMdruck(line, m)
			return
			}
		if (stop == "x") stop = count
		application.activeWindow.simulateIBWKey("F1");							
		}
	stop=stop-1
	if (stop == 0)
		{
		DBSMdruck(line, m)
		return
		}
	a=a+1
	}
}
function DBSMSysHiera() {
// Voraussetzung: Eine Ergebnismenge mit Tq- oder Tk-Sätzen ist geladen
// Ziel: in allen Sätzen der Ergebnismenge ab der aktuellen Position bis zum Ende werden alle Über- und Untergeordneten in 553-Felder eingetragen
// Die Funktion kann auch zur Aktualisierung eines einzelnen Satzes verwendet werden.
//meldung = "IDN - P3GPP: " + application.activeWindow.variable("P3GPP");
//meldung = meldung + "\nAnzahl - P3GSZ: " + application.activeWindow.variable("P3GSZ");
//meldung = meldung + "\nSatznr - P3LNR: " + application.activeWindow.variable("P3LNR");
//meldung = meldung + "\nSatznr - P3GTI: " + application.activeWindow.variable("P3GTI");
//meldung = meldung + "\nSchirm - scr: " + application.activeWindow.Variable("scr");
//application.messageBox("",meldung,"")
//return

if (!application.activeWindow.Variable("scr") || application.activeWindow.Variable("scr") == "FI" || application.activeWindow.Variable("scr") == "GN")
	{
	if (application.activeWindow.Variable("scr") == "FI") application.messageBox("Diese Funktion steht momentan nicht zur Verfügung!","Um die Funktion nutzen zu können\nmuss eine Ergebnismenge oder \nein einzelner Titel ausgewählt werden.","");
	else if (application.activeWindow.Variable("scr") == "GN") application.messageBox("Diese Funktion steht momentan nicht zur Verfügung!","Um die Funktion nutzen zu können\nmuss eine Ergebnismenge oder \nein einzelner Titel in der Voll-\nbzw. Kurzanzeige angezeigt sein.","");
	else if (!application.activeWindow.Variable("scr")) application.messageBox("Diese Funktion steht momentan nicht zur Verfügung!","Um die Funktion nutzen zu können\nmüssen Sie sich erst einloggen\nund eine Ergebnismenge bilden.","");
	return
	}
application.activeWindow.simulateIBWKey("F7");										//Bearbeiten ein
if (!application.activeWindow.title.find("Tk",true,true,false) && !application.activeWindow.title.find("Tq",true,true,false))
	{
	application.activeWindow.simulateIBWKey("F1");
	application.messageBox("Diese Funktion steht momentan nicht zur Verfügung!","Die Funktion kann nur auf eine Ergebnismenge von Tk-/Tq-Sätzen oder \neinen einzelnen Tk-/Tq-Satz angewendet werden.","");
	return
	}

// Setze Vorgabewerte 
var stufe = new Array();													// Array für die Stufen oberhalb
var querry = new Array();													// Array für die Suchabfragen, die aus den Stufen gebildet werden
var idn = new Array();														// Array für die IDNs, die bei der übergeordneten Stufe eingesammelt werden
var idn1 = new Array();														// Array für die IDNs, die bei der untergeordneten Stufe eingesammelt werden
var WinID1 = application.activeWindow.windowID										// Fenster-ID des aktuellen Festers merken
var klass = ""															// Klassifikation (Teilfeld h)
var notation = ""															// Notation (Teilfeld a)
querry[0] = ""															// Suchfrage 0 vorbelegen
stufe[0] = ""															// Stufe 0 vorbelegen
merkklass = ""															// Klassifikation des jeweils vorhergehenden Satzes
StrIDN = application.activeWindow.variable("P3GPP");									// IDN speichern (um später Zirkelverweise zu vermeiden)
count = 1;																// count mit 1 vorbelegen
count = application.activeWindow.Variable("P3GSZ");									// Anzahl der Treffer auf count
count1 = application.activeWindow.Variable("P3GTI");									// aktueller Satz auf count1
if (count1 != 1) var count = count - count1										// wenn es nicht der erste Satz ist, count auf Anzahl der restlichen Sätze verkleinern
if (count != 1)															// wenn count nicht 1 ist abfragen, ob alle Sätze oder nur der aktuelle berarbeitet werden sollen
	{
	var prompt = utility.newPrompter();
	prompt.setDebug(true);
	if (prompt.confirmEx("ACHTUNG","Die Ergebnismenge hat noch "+count+" Treffer!\nFunktion auf die restliche Ergebnismenge anwenden?", "Yes", "No", "", "", false) != 0) count = 1
	} 
a = 1

// Schleife
while (a <= count)																	//vom aktuellen bis zum letzten Satz
	{
	// 1. Vorbereitung
	// 1.1 neue Klassifikation finden
	var x = 1																//Zähler für Untergeordnete auf 1
	StrIDN = application.activeWindow.variable("P3GPP");									// IDN speichern (um später Zirkelverweise zu vermeiden)
	application.activeWindow.simulateIBWKey("F7");										//Bearbeiten ein
	klass = application.activeWindow.title.findTag("153",0,false,true,true);					//suche Feld 153
	klass = klass.substring(klass.indexOf("$x")+2,klass.length)								//schneide die Sorierform der Notation heraus und merke sie auf klass
	// 1.2 vorhandene 553-Verknüpfungen löschen
	test = application.activeWindow.title.findTag("553",0,true,true);							//teste ob 553 enthalten ist
            n = 0
	while (test)															//solange dies der Fall ist, lösche alle 553
		{
		if (test.indexOf("$4nu") > -1) 
                            {
                            application.activeWindow.title.deleteLine()					//lösche aktuelle Zeile, wenn diese ein $4nueb oder $4nunt enthält
                            }
                            else 
                            {
                            n = n + 1
                            }
            	test = application.activeWindow.title.findTag("553",n,true,true)							//suche nächste 553
		}
            n = 0
	// 1.3 Klassifikationsstufe mit letztem Satz vergleichen (wenn identisch: Übergeordnete gleich aus letztem Satz übernehmen und weiter bei 3.3.)
	if (klass.lastIndexOf(".")> -1) neuklass = klass.substring(0,klass.lastIndexOf("."))				//wenn es mindestens ein "." als Trenner gibt, speichere alles bis zum letzten "." auf neuklass
	else if (klass.lastIndexOf("-")> -1) neuklass = klass.substring(0,klass.lastIndexOf("-"))				//oder wenn es mindestens ein "-" als Trenner gibt, speichere alles bis zum letzten "-" auf neuklass
	else neuklass = klass															//sonst speicher klass auf neuklass
	if (neuklass == merkklass)														//wenn Klassifikationsstufe identisch mit letztem Satz
		{
		for(var i = 1; i <= n; i++)														//für Übergeordnete alle Suchfragen querry[0] bis querry[n-1] durcharbeiten
			{
			application.activeWindow.title.endOfBuffer(false);											//ans Ende des Satzes gehen
			if (idn1[i] != StrIDN) application.activeWindow.title.insertText("553 $b" + i + "!" + idn1[i] + "!$4nueb\n");	//dort 553 mit IDN bilden
			}
		}
	else																	//sonst (wenn Klassifikationsstufe nicht identisch mit letztem Satz)
		{
	// 2. Vorgabewerte neu setzen
		idn1.splice(0, n);														//Arry idn1 aufräumen (das sind die IDN der Übergeordneten des letzten Satzes)
		var n = 0																//Zähler für Querry (Übergeordnete) auf 0
		merkklass = klass.substring(0,klass.lastIndexOf("."))
		if (klass.indexOf("-") > -1)													//wenn Klassifikation als Trenner "-" enthält
			{
			notation = klass.substring(klass.lastIndexOf("-")+1,klass.length)							//merke alles nach dem letzten Trenner auf notation
			klass = klass.substring(0,klass.lastIndexOf("-"))									//und den Anfang bis zum letzten Trenner auf klass
			}
	// 3. alle über- und untergeordneten Klassifikations-Stufen neu bilden
	// 3.1. übergeordnete Stufen finden und in querry[n] einsammeln
	// 3.1.1. Klassifikation abarbeiten
		application.activeWindow.title.endOfBuffer(false);									//gehe ans Ende des Satzes
		rest = klass															//belege Rest mit der gesamten Klassifikation vor
		do																	//solange rest einen "-" als Trenner enthält
			{
			if (rest.indexOf("-") > -1) stufe[n] = rest.substring(0,rest.indexOf("-"))					//wenn rest einen "-" als Trenner enthält, nächste Stufe aus rest herausschneiden
			else stufe[n] = rest														//sonst nächste Stufe gleich rest
			if (querry[n-1]) querry[n] = querry[n-1] + "-" + stufe[n] 								//wenn es schon eine Suchfrage gab, neue Suchfrage aus querry[n-1] und stufe[n] bilden
			else querry[n] = stufe[n]													//sonst Suchfrage aus Stufe[n] bilden 
			rest = rest.substring(rest.indexOf("-")+1,rest.length);								//rest um die letzte Stufe verkürzen
			n=n+1																	//Zähler n eins hochsetzen
			}while (rest.indexOf("-")> -1);
		if (querry[n-1] && querry[n-1] != rest) 
			{
			querry[n] = querry[n-1] + "-" + rest
			n=n+1
			}

	// 3.1.2. Notation abarbeiten
		if (notation != "")														//wenn Notaion vorhanden
			{
			rest = notation															//belege Rest mit der gesamten Notation vor
			do	{																//solange rest einen "." als Trenner enthält
				if (rest.indexOf(".") > -1) stufe[n] = rest.substring(0,rest.indexOf("."))					//wenn rest einen "-" als Trenner enthält, nächste Stufe aus rest herausschneiden
				else stufe[n] = rest														//sonst nächste Stufe gleich rest
				querry[n] = querry[n-1] + "." + stufe[n] 											//die Suchfrage als querry[n] speichern
				rest = rest.substring(rest.indexOf(".")+1,rest.length);								//rest um die letzte Stufe verkürzen
				n=n+1																	//Zähler n eins hochsetzen
				}while (rest.indexOf(".")> -1);
			querry[n] = querry[n-1] + "." + rest										//letzte Suchfrage mit rest bilden
			}
	// 3.2. 553 für Übergeordnete neu bilden
		for(var i = 1; i <= n; i++)											//für Übergeordnete alle Suchfragen querry[0] bis querry[n-1] durcharbeiten
			{
			application.activeWindow.command("f syf " + querry[i-1], false);					//Suche ausführen
			var gefunden = application.activeWindow.Status								//Status abfragen
			if (gefunden == "NOHITS") 											//wenn nichts gefunden
				{														//verschiebe den Trenner "-" um eine Stelle nach hinten
				newquerry = querry[i-1].substring(0,querry[i-1].indexOf(".")) + "-" + querry[i-1].substring(querry[i-1].indexOf(".")+1,querry[i-1].length)
				application.activeWindow.command("f syf " + newquerry , false);				//Suche erneut ausführen ausführen
				var gefunden = application.activeWindow.Status							//Status abfragen
				if (gefunden == "NOHITS") 											//wenn nichts gefunden
					{														//verschiebe den Trenner "-" um eine Stelle nach hinten
					newquerry = newquerry.substring(0,newquerry.indexOf(".")) + "-" + newquerry.substring(newquerry.indexOf(".")+1,newquerry.length)
					application.activeWindow.command("f syf " + newquerry , false);				//Suche erneut ausführen ausführen
					var gefunden = application.activeWindow.Status							//Status abfragen
					}
				}
			if (gefunden == "OK")											//wenn etwas gefunden wurde
				{
				idn1[i] = application.activeWindow.variable("P3GPP");											//IDN speichern
				application.activeWindow.closeWindow();													//Fenster schließen
				application.activateWindow(WinID1);														//zum ursprünglichen Fenster zurück
				application.activeWindow.title.endOfBuffer(false);											//ans Ende des Satzes gehen
				if (idn1[i] != StrIDN) application.activeWindow.title.insertText("553 $b" + i + "!" + idn1[i] + "!$4nueb\n");	//dort 553 mit IDN bilden
				}
			}
		}
	// 3.3. 553 für Untergeordnete neu bilden
	if (klass != "" && notation != "") find = klass + "-" + notation									//wenn sowohl Klassifikation als auch Notation besetzt, diese zusammen als Suchbegriff definieren
	else if (klass != "") find = klass 														//sonst, wenn Klassifikation besetzt, diese als Suchbegriff definieren
	else find = notation																//sonst (wenn also Notation und nicht Ksassifikation besetzt, Notation als Suchbegriff definieren
	application.activeWindow.command("f syf " + find + ".####", true);								//jetzt mit find und .#### die Untergeordneten suchen
	var gefunden = application.activeWindow.Status												//Status abfragen
	if (gefunden == "NOHITS") application.activeWindow.command("f syf " + find + "-####", true);				//wenn nicht gefunden, nochmal mit find und -#### suchen
	var gefunden = application.activeWindow.Status												//Status abfragen
	if (gefunden == "NOHITS") 															//wenn noch nicht gefunden
		{
		if (find.indexOf(".") > -1)															//und wenn in find ein . als Trenner ist
			{
			newquerry = find.substring(0,find.lastIndexOf(".")) + "-" + find.substring(find.lastIndexOf(".")+1,find.length)		//bilde neue Suchfrage mit - als Trenner
			application.activeWindow.command("f syf " + newquerry + ".####", false);								//Suche erneut ausführen
			var gefunden = application.activeWindow.Status													//Status abfragen
			if (gefunden == "NOHITS") application.activeWindow.command("f syf " + newquerry + "-####", false);				//wenn noch nicht gefunden, nochmal mit -#### suchen
			}
		else
			{
			if (find.indexOf("-") > -1) 
				{
				newquerry = find.substring(0,find.lastIndexOf("-")) + "." + find.substring(find.lastIndexOf("-")+1,find.length)
				application.activeWindow.command("f syf " + newquerry + ".####", true);				//und mit querry[n] die Untergeordneten erneut suchen
				}
			var gefunden = application.activeWindow.Status								//Status abfragen
			if (gefunden == "NOHITS") application.activeWindow.command("f syf " + newquerry + "-####", false);			//Suche ausführen	
			var gefunden = application.activeWindow.Status								//Status abfragen
			}
		}
	if (gefunden != "NOHITS")											//wenn etwas gefunden wurde
		{
		if (application.activeWindow.Variable("scr") == "GN") application.activeWindow.simulateIBWKey("FR");
		count1 = application.activeWindow.Variable("P3GSZ");							//Anzahl der Treffer auf count1
		application.activeWindow.command("sor kls", false);							//Ergebnismenge sortieren
		application.activeWindow.command("\\TOO D", false);							//ersten Satz anzeigen
		while (x <= count1)												//vom ersten bis zum letzten Satz
			{
			idn[x] = application.activeWindow.Variable("P3GPP");							//IDN speichern
			application.activeWindow.simulateIBWKey("F1");								//zum nächsten Satz
			x=x+1															//Zähler eins hochsetzen
			}
		application.activeWindow.closeWindow();									//Fenster schließen
		application.activateWindow(WinID1);										//zum ursprünglichen Fenster zurück
		for(i in idn)													//mit allen IDNs je eine 553 für Untergeordneten bilden
			{
			if (idn[i] != StrIDN) application.activeWindow.title.insertText("553 !" + idn[i] + "!$4nunt\n");
			}
		}
	//4. Endbearbeitung
//application.messageBox("","vorher: "+application.windows.count,"")
	application.activateWindow(WinID1);										//zum ursprünglichen Fenster zurück
	var sect = "dnbUser";
	var strAbteilung = application.getProfileString(sect, "abteilung", "");
	if (strAbteilung == "BSM")
		{
		test = "";
		test = application.activeWindow.title.findTag("012",0,true,true);					//testen ob 012 schon vorhanden
		if (test == "") application.activeWindow.title.insertText("012 o");				//wenn nicht diese mit "o" belegen
		}
	application.activeWindow.simulateIBWKey("FR");								//speichern
	if (application.activeWindow.Status != "OK")								//wenn nicht erfolgreich
		{
		application.messageBox("FEHLER", "konnte nicht speichern", "");					//Fehlerhinweis ausgeben
		application.activeWindow.simulateIBWKey("FE");								//Funktion abbrechen
		}
	if (count > 1)													//wenn Ergebnismenge größer 1
		{
		application.activeWindow.simulateIBWKey("F1");								//nächsten Satz aufrufen
//application.messageBox("","nachher: "+application.windows.count,"")
		stufe.splice(0, n);												//Arry stufe aufräumen	
		querry.splice(0, n);												//Arry querry aufräumen
		idn.splice(0, 999);													//Arry IDN aufräumen
		}
	a=a+1
//	if (application.windows.count > 1)
//		{
//		for (var i = 2; i <= application.windows.count; i++)
//			{
//			application.activateWindow(i)
//			application.messageBox("",application.activeWindow.codedData()+"\n"+application.activeWindow.materialCode()+"\n"+application.activeWindow.noviceMode(),"")
//			application.closeWindow(i)
//			}
//		}
	}
}
function DBSMSysSort() {
// Voraussetzung: Eine Ergebnismenge mit Tq- oder Tk-Sätzen ist geladen
// Ziel: in allen Sätzen der Ergebnismenge ab der aktuellen Position bis zum Ende werden alle Sortierformen in $x der 153 neu gebildet
// Die Funktion kann auch zur Aktualisierung eines einzelnen Satzes verwendet werden.

if (!application.activeWindow.Variable("scr") || application.activeWindow.Variable("scr") == "FI")
	{
	if (application.activeWindow.Variable("scr") == "FI") application.messageBox("Diese Funktion steht momentan nicht zur Verfügung!","Um die Funktion nutzen zu können\nmus eine Ergebnismenge oder \nein einzelner Titel ausgewählt werden.","");
	else application.messageBox("Diese Funktion steht momentan nicht zur Verfügung!","Um die Funktion nutzen zu können\nmüssen Sie sich erst einloggen\nund eine Ergebnismenge bilden.","");
	return
	}
application.activeWindow.simulateIBWKey("F7");										//Bearbeiten ein
if (!application.activeWindow.title.find("Tk",true,true,false) && !application.activeWindow.title.find("Tq",true,true,false))
	{
	application.activeWindow.simulateIBWKey("FE");
	application.messageBox("Diese Funktion steht momentan nicht zur Verfügung!","Die Funktion kann nur auf eine Ergebnismenge von Tk-/Tq-Sätzen oder \neinen einzelnen Tk-/Tq-Satz angewendet werden.","");
	return
	}


count = 1;																	//count mit 1 vorbelegen
count = application.activeWindow.Variable("P3GSZ");							//Anzahl der Treffer auf count
count1 = application.activeWindow.Variable("P3GTI");						//aktueller Satz auf count1
if (count1 != 1) var count = count - count1									//wenn es nicht der erste Satz ist, count auf Anzahl der restlichen Sätze verkleinern
if (count != 1)																//wenn count nicht 1 ist abfragen ob alle Sätze oder nur der aktuelle
	{
	if (count == 0) var count = 1;
	else
		{
		var prompt = utility.newPrompter();
		prompt.setDebug(true);
		if (prompt.confirmEx("ACHTUNG","Die Ergebnismenge hat noch "+count+" Treffer!\nFunktion auf die restliche Ergebnismenge anwenden?", "Yes", "No", "", "", false) != 0) count = 1
		}
	} 
a = 1;
while (a <= count)											//vom ersten bis zum letzten Satz
	{
	application.activeWindow.command("\\TOO D", false);						//aktiviere den ersten Titel
	application.activeWindow.simulateIBWKey("F7");							//Bearbeiten ein
	feld153 = application.activeWindow.title.findTag("153",0,false,true,true);		//suche Feld 153
	if (feld153.indexOf("$x") > -1)									//wenn $x bereits vorhanden
		{
		application.activeWindow.title.find("$x",false,false,false);				//gehe zu $x
		application.activeWindow.title.endOfField(true);						//und markiere bis zum Ende des Feldes
		}
	else														//sonst
		{
		application.activeWindow.title.endOfField(false);						//gehe zum Ende des Feldes
		}
	klass = feld153.substring(feld153.indexOf("$h")+2,feld153.indexOf("$j"))		//schneide die Kurzbezeichnung der Klassifikation heraus und merke sie auf klass
	klass = klass.replace(/\W/g, "-");									//ersetze alle Sonderzeichen durch Bindestrich
	klass = klass + "-";											//füge am Ende Bindestich ein
	klass = klass.replace(/\-\-/g, "-");								//ersetze zwei Ztriche durch einen 

	notation = feld153.substring(0,feld153.indexOf("$"))						//schneide die Notation heraus und merke sie auf notation
	notation = notation.replace(/\W/g, ".");								//ersetze alle Sonderzeichen durch Punkt
	notation = notation.replace(/\.\./g, ".");							//ersetze zwei Punkte durch einen Punkt

	notation = klass + notation										//füge Notation und Klasse zusammen
															//setze Vorgabewerte 
	stufe = ""
	sort = ""
	test = 1
	
	while (notation != "")											//solange noch eine Notationsstufe vorhanden ist
		{
//		application.messageBox("vorher","Stufe: "+stufe+"\nNotation: "+notation,"")
		if (notation.indexOf(".") > -1)									//wenn Notatin noch einen Trenner enthält
			{
			stufe = notation.substring(0,notation.indexOf(".")+1)						//belege Stufe mit dem ersten Teil der Notation bis zum Trenner
			notation = notation.substring(notation.indexOf(".")+1,notation.length)			//belege Notation mit dem Rest der Notation ab dem Trenner neu
			if (stufe.indexOf("-") > -1)
				{
				notation = stufe.substring(stufe.indexOf("-")+1,stufe.length) + notation		//belege Notation mit dem Rest der Notation ab dem Trenner neu
				stufe = stufe.substring(0,stufe.indexOf("-")+1)							//belege Stufe mit dem ersten Teil der Notation bis zum Trenner
				}
			}
		else															//sonst
			{
			stufe = notation													//belege Stufe mit Notation 
			notation = ""													//mache Notation leer
			if (stufe.indexOf("-") > -1)
				{
				notation = stufe.substring(stufe.indexOf("-")+1,stufe.length)			//belege Notation mit dem Rest der Notation ab dem Trenner neu
				stufe = stufe.substring(0,stufe.indexOf("-")+1)							//belege Stufe mit dem ersten Teil der Notation bis zum Trenner
				}
			}
			
		if (stufe == "-") stufe = ""													//wenn Stufe "-" (oberste Stufe) mache Stufe leer
		if (stufe == "Fg.") stufe = ""													//wenn Stufe "Fg." (= Formalgruppe ESK) mache Stufe leer
		if (stufe.charAt(0) < "A")														//wenn erstes Zeichen von Stufe eine Ziffer ist			
			{
			zeichen = stufe.substring(stufe.length-1,stufe.length);
			if (zeichen == "." || zeichen == "-") stufe = stufe.substring(0,stufe.length -1)
			else zeichen = ""
			if (stufe.length == 4) sort = sort + stufe + zeichen								//wenn Stufe 4 Zeichen fülle mit 1 Null rechtsbündig auf 
			if (stufe.length == 3) sort = sort + "0" + stufe + zeichen								//wenn Stufe 3 Zeichen fülle mit 2 Nullen rechtsbündig auf
			if (stufe.length == 2) sort = sort + "00" + stufe + zeichen								//wenn Stufe 2 Zeichen fülle mit 3 Nullen rechtsbündig auf
			if (stufe.length == 1) sort = sort + "000" + stufe + zeichen								//wenn Stufe 1 Zeichen fülle mit 3 Nullen rechtsbündig auf
			}
		else																			//sonst
			{
			zeichen = stufe.substring(stufe.length-1,stufe.length);
			if (zeichen != "." && zeichen != "-") zeichen = "";
			else stufe = stufe.substring(0,stufe.length-1);
			if (stufe.length > 4) sort = sort + stufe.substring(0,4).toLowerCase() + zeichen;							//wandle in Kleinbuchstaben und schneide nach vier Zeichen ab
			if (stufe.length == 4) sort = sort + stufe.toLowerCase() + zeichen;							//wandle in Kleinbuchstaben und schneide nach vier Zeichen ab
			if (stufe.length == 3) sort = sort + stufe.toLowerCase() + "0" + zeichen;							//wandle in Kleinbuchstaben und schneide nach vier Zeichen ab
			if (stufe.length == 2) sort = sort + stufe.toLowerCase() + "00" + zeichen;							//wandle in Kleinbuchstaben und schneide nach vier Zeichen ab
			if (stufe.length == 1) sort = sort + stufe.toLowerCase() + "000" + zeichen;							//wandle in Kleinbuchstaben und schneide nach vier Zeichen ab
			}
		test = test + 1
		}
	//sort nachbehandeln
	if (sort.lastIndexOf(".") == sort.length -1) sort = sort.substring(0,sort.length -1)	//Punkt am Ende entfernen
	if (sort.lastIndexOf("-") == sort.length -1) sort = sort.substring(0,sort.length -1)	//Bindestrich am Ende entfernen
	sort = sort.replace(/\-\./g, "-");														//Strich-Punkt durch einen Strich ersetzen
	sort = sort.replace(/\-\-/g, "-");														//zwei Striche durch einen ersetzen
	sort = sort.replace(/\.\./g, ".");														//zwei Punkte durch einen ersetzen
	sort = sort.replace(/\.i000$/g, ".0001");													//ersetze römische Zahl I
	sort = sort.replace(/\.ii00$/g, ".0002");													//ersetze römische Zahl II
	sort = sort.replace(/\.iii0$/g, ".0003");												//ersetze römische Zahl III
	sort = sort.replace(/\.iv00$/g, ".0004");													//ersetze römische Zahl IV
	sort = sort.replace(/\.v000$/g, ".0005");													//ersetze römische Zahl V
	sort = sort.replace(/\.vi00$/g, ".0006");													//ersetze römische Zahl VI
	sort = sort.replace(/\.vii0$/g, ".0007");												//ersetze römische Zahl VII
	sort = sort.replace(/\.viii$/g, ".0008");												//ersetze römische Zahl VIII
	sort = sort.replace(/\.ix00$/g, ".0009");													//ersetze römische Zahl IX
	sort = sort.replace(/\.x000$/g, ".0010");													//ersetze römische Zahl X
	sort = sort.replace(/\.xi00$/g, ".0011");													//ersetze römische Zahl XI
	sort = sort.replace(/\.xii0$/g, ".0012");												//ersetze römische Zahl XII
	sort = sort.replace(/\.xiii$/g, ".0013");												//ersetze römische Zahl XIII
	sort = sort.replace(/\.xiv0$/g, ".0014");												//ersetze römische Zahl XIV
	sort = sort.replace(/\.xv00$/g, ".0015");													//ersetze römische Zahl XV
	sort = sort.replace(/\.xvi0$/g, ".0016");												//ersetze römische Zahl XVI
	sort = sort.replace(/\.xvii$/g, ".0017");												//ersetze römische Zahl XVII
	sort = sort.replace(/\.xviii$/g, ".0018");												//ersetze römische Zahl XVIII
	sort = sort.replace(/\.xix0$/g, ".0019");												//ersetze römische Zahl XIX
	sort = sort.replace(/\.xx00$/g, ".0020");													//ersetze römische Zahl XX
	sort = sort.replace(/\.i000\./g, ".0001.");												//ersetze römische Zahl I.
	sort = sort.replace(/\.ii00\./g, ".0002.");												//ersetze römische Zahl II.
	sort = sort.replace(/\.iii0\./g, ".0003.");												//ersetze römische Zahl III.
	sort = sort.replace(/\.iv00\./g, ".0004.");												//ersetze römische Zahl IV.
	sort = sort.replace(/\.v000\./g, ".0005.");												//ersetze römische Zahl V.
	sort = sort.replace(/\.vi00\./g, ".0006.");												//ersetze römische Zahl VI.
	sort = sort.replace(/\.vii0\./g, ".0007.");												//ersetze römische Zahl VII.
	sort = sort.replace(/\.viii\./g, ".0008.");												//ersetze römische Zahl VIII.
	sort = sort.replace(/\.ix00\./g, ".0009.");												//ersetze römische Zahl IX.
	sort = sort.replace(/\.x000\./g, ".0010.");												//ersetze römische Zahl X.
	sort = sort.replace(/\.xi00\./g, ".0011.");												//ersetze römische Zahl XI.
	sort = sort.replace(/\.xii0\./g, ".0012.");												//ersetze römische Zahl XII.
	sort = sort.replace(/\.xiii\./g, ".0013.");												//ersetze römische Zahl XIII.
	sort = sort.replace(/\.xiv0\./g, ".0014.");												//ersetze römische Zahl XIV.
	sort = sort.replace(/\.xv00\./g, ".0015.");												//ersetze römische Zahl XV.
	sort = sort.replace(/\.xvi0\./g, ".0016.");												//ersetze römische Zahl XVI.
	sort = sort.replace(/\.xvii\./g, ".0017.");												//ersetze römische Zahl XVII.
	sort = sort.replace(/\.xviii\./g, ".0018.");											//ersetze römische Zahl XVIII.
	sort = sort.replace(/\.xix0\./g, ".0019.");												//ersetze römische Zahl XIX.
	sort = sort.replace(/\.xx00\./g, ".0020.");												//ersetze römische Zahl XX.

	application.activeWindow.title.insertText("$x" + sort);							//neues $x einfügen
	application.activeWindow.simulateIBWKey("FR");											//speichern
	if (application.activeWindow.Status != "OK")											//wenn nicht erfolgreich
		{
		application.messageBox("FEHLER", "konnte nicht speichern", "");							//Fehlermeldung und
		application.activeWindow.simulateIBWKey("FE");											//Bearbeiten abbrechen
		}
	if (count > 1) application.activeWindow.simulateIBWKey("F1");							//wenn Ergebnismenge größer 1 gehe zum nächsten Satz
	a=a+1																					//Satzzähler um 1 erhöhen
	}	
	
}
function FehlerReg() {
	// Aufruf der Fehlerregistrierung unabhängig von der Normdaten-Suche (wird vor allem interessant, wenn auch andere Fehler als noch nicht integrierte Normdaten registriert werden sollen
	if (!application.activeWindow.title) 
		{
		application.messageBox("Diese Funktion steht momentan nicht zur Verfügung!","Um die Funktion nutzen zu können\nmus im Bearbeitungsmodus zunächst das Feld \nund die zu registrierende Zeichenkette eingegeben werden.\nBeispiel: 5590 Meier, Hans","");
		}
	else
		{
		
		
		application.activeWindow.title.startOfField(false);		//gehe zum Start des Feldes	
		application.activeWindow.title.endOfField(true);		//Selektiere von der Position alles bis zum Ende
		var content = application.activeWindow.title.Selection;	//markiere den Suchbegriff
      	content = content.substring(5,content.length);     		//merke den Suchbegriff (wordRight hatte den Nachteil, dass bei leerem Feld auf das nächste Feld gesprungen wird)
		// wenn der Suchbegriff leer ist (ganz leer oder ein Buchstabe oder nur der In-Vermerk)
		if (content.length < 2)
			{
			content = __Prompter("Wonach soll gesucht werden?","Geben Sie den Suchbegriff ein!")
			}
		erg = __Pruef("Eine Frage","Welcher Fehlertyp soll registriert werden?\n 1 = allegro-Personennormsatz noch nicht in GND\n 2 = allegro-Körperschaftssatz noch nicht in GND\n 3 = allegro-Schlagwortsatz noch nicht in GND\n 4 = andere Fehler","1,2,3,4,5","1")
		__FehlerRegistrierung(erg, content)
	}
}
function test1() {
	// Send the command ""\sca \ter PER meienreis,johann"" to the system and display the data in the same window
	application.activeWindow.command("\\sca \\ter PER meienreis,johann", false);
	// Send the command ""\sca \ter PER meiendorff,margareta von"" to the system and display the data in the same window
	application.activeWindow.command("\\sca \\ter PER meiendorff,margareta von", false);
}
function test5() {
//	application.activeWindow.command("sc per meier", false);
//	application.activeWindow.command("\\sca PER meier,agnes", false);
	application.activeWindow.command("\\sca \\ter PER meienreis, walter", false);
//	application.activeWindow.command("\\sca \\ter PER meier,", true);

}
function dbsmtest() {
	var xulFeatures = "centerscreen, chrome, close, titlebar,modal=no,dependent=yes, dialog=yes";
	open_xul_dialog("chrome://ibw/content/xul/dnb_einstellungen_dialog.xul", xulFeatures);

//chrome://u://test.xul", xulFeatures);	
}

function open_xul_dialog(theUrl, theFeatures, theArguments)
{
	// try to get the window-watcher
	//var ww    = Components.classes["@mozilla.org/embedcomp/window-watcher;1"]
            //                     .getService(Components.interfaces.nsIWindowWatcher);

	if (!ww) {
		// no chance, give up
		return false;
	}

	// let's try to get a valid parent
	var theParent = ww.activeWindow;

	var features = null;

	if (theFeatures != null) {
		features = theFeatures;
	} else {
		// you may choose to remove some of the features
		// you may also want to specify width=xxx and/or height=xxx
		features = "centerscreen,chrome,close,titlebar,resizable,modal,dialog=yes";
	}

	// it doesn't matter, if we don't have a parent
	// we just use the active window, whether its null or not
	ww.openWindow(theParent, theUrl, "", features, theArguments);
}

function GNDBezf() {
// Überprüft ob Bezf wechselseitig vorhanden, wenn nicht wird 500 im Zielsatz ergänzt
var texta = new Array();							//Array für Änderungsprotokoll definieren
var x = 0								//Array-Zähler auf 0 setzen
if (application.activeWindow.variable("scr") == "7A") application.activeWindow.simulateIBWKey("FR")   //wenn Kurzliste, zeige den aktuellen Satz an
count = 1;								// count mit 1 vorbelegen
count = application.activeWindow.Variable("P3GSZ");			// Anzahl der Treffer auf count
count1 = application.activeWindow.Variable("P3GTI");			// aktueller Satz auf count1
if (count1 != 1) var count = count - count1				// wenn es nicht der erste Satz ist, count auf Anzahl der restlichen Sätze verkleinern
if (count != 1)								// wenn count nicht 1 ist abfragen, ob alle Sätze oder nur der aktuelle berarbeitet werden sollen
	{
	var prompt = utility.newPrompter();
	prompt.setDebug(true);
	if (prompt.confirmEx("ACHTUNG","Die Ergebnismenge hat noch "+count+" Treffer!\nFunktion auf die restliche Ergebnismenge anwenden?", "Yes", "No", "", "", false) != 0) count = 1
	} 
a = 1									//Satzzähler auf 1 setzen
while (a <= count)							//solange Satzzähler kleiner/gleich Ergebnismenge
	{
            onabort="break"
            ID1 = application.activeWindow.variable("P3GPP");			// IDN speichern
            application.activeWindow.simulateIBWKey("F7");	                                    //Bearbeiten ein
	n = 0									//Feldzähler auf 0 setzen
	$g = application.activeWindow.title.findTag("375",0,false,false,false)	//Geschlecht aus Feld 375 auf $g merken
	test = application.activeWindow.title.findTag("500",n,false,true);		//Feld 500 auf test merken
	while (test != false)							//solange neues Feld 500 gefunden
                        {
		n = n+1									//Feldzähler um 1 erhöhen
                        if (test.indexOf("!") != -1 && test.indexOf("bezf") != -1)                                  //wenn IDN vorhanden, also verknüpft ist
                            {
                            ID2 = test.substring(test.indexOf("!")+1,test.lastIndexOf("!"))		    //IDN aus Feld 500 als ID2 merken
                            $Z = ""                                                                                                  //$Z vorbelegen
                            if (test.indexOf("$Z") != -1) $Z = test.substring(test.indexOf("$Z"),test.indexOf("$v"))                          //$Z auf $Z
	                $v = test.substring(test.indexOf("$v")+1,test.length)			    //$v aus Feld 500 als $v merken
		    strCommand = "f IDN "+ID2						    //Such-Zeichenkette auf strCommand merken
		    application.activeWindow.command(strCommand,true);			    //nach verknüpfter IDN suchen
                            application.activeWindow.simulateIBWKey("F7");	                            //Bearbeiten ein
    		    test1 = application.activeWindow.title.find(ID1,false,false,false);	    //im gefundenen Satz die IDN des Ausgangssatz suchen
		    if (test1==false)							    //wenn nicht gefunden, gibt es keine $4bezf-Relation
			{
                        	$g1 = application.activeWindow.title.findTag("375",0,false,false,false)	//Geschlecht aus Feld 375 auf $g merken
			$v1 = ""									//$v1 vorbelegen
			$v2 = ""									//$v2 vorbelegen
			if ($v.indexOf("Stief") != -1) $v2 = "Stief" 				//wenn $v "Stief" enthält "Stief" auf $v2 merken
			if ($v.indexOf("ater") != -1 && $g == "m") $v1 = "Sohn"			//wenn $v "ater" enthält und Geschlecht "m", "Sohn" auf $v1 merken
			if ($v.indexOf("ater") != -1 && $g == "f") $v1 = "Tochter"			//wenn $v "ater" enthält und Geschlecht "w", "Tochter" auf $v1 merken
			if ($v.indexOf("utter") != -1 && $g == "m") $v1 = "Sohn"			//wenn $v "utter" enthält und Geschlecht "m", "Sohn" auf $v1 merken
			if ($v.indexOf("utter") != -1 && $g == "f") $v1 = "Tochter"			//wenn $v "utter" enthält und Geschlecht "w", "Tochter" auf $v1 merken
			if ($v.indexOf("ohn") != -1 && $g == "m") $v1 = "Vater"			//wenn $v "ohn" enthält und Geschlecht "m", "Vater" auf $v1 merken
			if ($v.indexOf("ohn") != -1 && $g == "f") $v1 = "Mutter"			//wenn $v "ohn" enthält und Geschlecht "w", "Mutter" auf $v1 merken
			if ($v.indexOf("ochter") != -1 && $g == "m") $v1 = "Vater"			//wenn $v "ochter" enthält und Geschlecht "m", "Vater" auf $v1 merken
			if ($v.indexOf("ochter") != -1 && $g == "f") $v1 = "Mutter"		            //wenn $v "ochter" enthält und Geschlecht "w", "Mutter" auf $v1 merken
			if ($v.indexOf("Ehe") != -1 && $g == "f") $v1 = "Ehefrau"			//wenn $v "Ehe" enthält und Geschlecht "w", "Ehefrau" auf $v1 merken
			if ($v.indexOf("Ehe") != -1 && $g == "m") $v1 = "Ehemann"		//wenn $v "Ehe" enthält und Geschlecht "m", "Ehemann" auf $v1 merken
			if ($v2 == "Stief") $v1 = $v2 + $v1.toLowerCase()				//wenn $v2 = "Stief" schreibe $v1 klein und füge davor "Stief" ein
			if ($v1 != "") $v1 = "$v"+$v1						//wenn $v1 überhaupt einen Inhalt hat, füge davor Teilfeldkennzeichen ein
			application.activeWindow.title.endOfBuffer(false);				//gehe ans Ende der Aufnahme
//application.messageBox("","ID1 = "+ID1+"\nv1 = "+$v1+"\nZ = "+$Z,"")
                                    if ($g1 == "f" && ($v1 == "Sohn" || $v1 == "Tochter"))
                                        {
			    application.activeWindow.simulateIBWKey("FE");				//bereche Bearbeitung ab
			    texta[x] = ID2+"\t"+ID1+"\t"+$g1+"\t"+$v1+"\t"+$Z+"\t"+"nicht gespeichert -> Mutter";				//Protokollzeile auf Array x 
                                        }
                                    else
                                        { 
//			    application.activeWindow.title.insertText("\n500 !"+ID1+"!$4bezf"+$v1+$Z);	//füge Feld 500 mit ID1 und $v1 ein
			    application.activeWindow.simulateIBWKey("FE");				//bereche Bearbeitung ab
//			    application.activeWindow.simulateIBWKey("FR");				//speichere den Datensatz
			    texta[x] = ID2+"\t"+ID1+"\t"+$g1+"\t"+$v1+"\t"+$Z+"\t"+application.activeWindow.Status;				//Protokollzeile auf Array x 
                                        }
			x = x + 1								//Array-Zähler um 1 erhöhen
			}
                             else
                                {
              	        application.activeWindow.simulateIBWKey("FE");				//abbrechen
                                }
                            application.activeWindow.closeWindow();					//schließe das Fenster
                            }
		test = application.activeWindow.title.findTag("500",n,false,true);		//suche nächstes Feld 500
		}
            application.activeWindow.simulateIBWKey("FE");				// abbrechen
	if (count > 1) application.activeWindow.simulateIBWKey("F1");		//wenn Ergebnismenge größer 1 gehe zum nächsten Satz
	a=a+1									//Satzzähler um 1 erhöhen
            if (a == 250) a = count 
	}	
	
while (x >= 0)                                                                                                      // Endbehandlung: Protokoll in Zwischenablage
	{
	application.activeWindow.clipboard = application.activeWindow.clipboard + "\n" + texta[x]		// Zeile an Zwischenablage anhängen
	x = x - 1
	}
}
function DBSMneu() {
	// setze Benutzerdaten
	var sect = "dnbUser";
	var strKuerzel = application.getProfileString(sect, "kuerzel", "");
	var strAbteilung = application.getProfileString(sect, "abteilung", "");
	// wenn Benuzerdaten noch nicht individualisiert, Meldung und Ende
	if (strKuerzel == "xxx") 
		{
		application.messageBox("Fehler!","Die Maske konnte nicht aufgerufen werden!\nBitte noch unter =Optionen / EinstellungenDNBBenutzerprofil das eigene Bearbeiter-Kürzel eintragen!","");
		}
	else
		{
Erfolg: do
{
		// Vorgabewerte setzen / abfragen
		var erg;
		var eArt;
		var SatzArt;
		var ITyp;
		var MTyp;
		var DTyp;
		var OGattung;
		var FormA1;
		var FormA2;
		var FormF1;
		var FormF2;
		var FormE1;
		var FormE2;
		var Form
		var DTMaterial = "";
		var ArtInhalt
		var FormV;
		var FormO;
		var Slg
                        var searchTerm 
                        var search
                        var VPer = "..."
                        var F1100 
                        var F1500 
                        var F1505 
		var F2105 = ""
                        var F3000 = ""
                        var F3100 = ""
		var F4030 
		var F4046 
                        var F4019
                        var F4070
                        var F4060 
		var F4105
                        var F4201 
                        var F4821
                        var F5320
                        var F510X
		var F5590 
		var F5591 
		var F5592 
		var F6710
                        var F6800
		var F7100 = ""
		var F7109
                        var F8100 
		var F8598 
                        var F8510
		var jetzt = new Date();
		var jahr = jetzt.getFullYear();
                        var jhr
                        var Status
                        var abbruch = "n"
                        var num
		// Was soll erfasst werden?
                        // spezielle Wünsche
                        if (strKuerzel == "Len" || "MaF")
                            {
                            SatzArt = "Aaa"
                            Slg = "M"
                            ITyp = "Text"
                            MTyp = "ohne Hilfsmittel zu benutzen"
                            DTyp = "Band"
                            var erg = __Pruef("Eine Frage","Was soll erfasst werden?\n 1 = StB\n 2 = IBST A-Format\n 3 = IBST B-Format\n 4 = IBST C-Format\n 5 = Sonstiges","1,2,3,4,5","1")
                            if (erg == null) return;
                            if (erg != 5)
                                {
                                Status = "kurz"
                                eArt = "ge";
                                }
                            if (erg == 1) 
                                {
                                F8598 = "[Z-StB]"
                                F2105 = "04,P01-s-63"
                                var ask = "Welche Signatur?"
                                F7100 = "StB JJJJ - Land, nnn"
                                
                                application.activeWindow.command("k", false);
                                if (application.activeWindow.title.findTag("7109",0,false,true).indexOf("StB") > 0) 
                                    {
                                    F6710 =  application.activeWindow.title.findTag("6710",0,false,true);
                                    F6710 = F6710.substring(0,F6710.indexOf("$"))
                                    F4821 =  application.activeWindow.title.findTag("4821",0,false,true);
                                    F4821 = F4821.substring(0,F4821.indexOf("$w")+2) + "??? EUR"  + F4821.substring(F4821.indexOf("$z"),F4821.length); 
                                    sig =  application.activeWindow.title.findTag("7109",0,false,true);
                                    sig = sig.substring(sig.lastIndexOf("!!")+5,sig.length)
                                    num = parseInt(sig.substring(sig.lastIndexOf(",")+1,sig.length), 10)
                                    num = num + 1
                                    sig = sig.substring(0,sig.lastIndexOf(",")+1) + " " + num.toString()
                                    F7100 =  sig
                                    F7109 = "!!DBSM/M/StB!! ; " + sig
                                    F8598 = application.activeWindow.title.findTag("8598",0,false,true);
                                    num = parseInt(F8598.substring(F8598.lastIndexOf("/")+1,F8598.length), 10)
                                    num = num + 1
                                    snum = num.toString() 
                                    do
                                        {
                                        snum = "0"+ snum
                                        } while(snum.length < 5)
                                    F8598 = F8598.substring(0,F8598.indexOf("/")+1) + snum
                                    application.activeWindow.simulateIBWKey("FE");
                                    num = 0
                                    }
                                Erfolg1: do 
                                    {
                                    var erg = __Prompter("Eine Frage",ask,sig)
                                    if (!erg) break Erfolg1;
                                    if (erg == sig) break Erfolg1;
                                    F7100 = erg
                                    F7109 = "!!DBSM/M/StB!! ; " + erg
                                    jhr = erg.substring(erg.indexOf(" ")+1,erg.indexOf("-")-1)
                                    if (erg.length < 4) {
                                        ask = "Jahr nicht exakt: " + jhr
                                        return
                                        } 
                                    var land = erg.substring(erg.indexOf("-")+2,14)
                                    land = land.replace(/ä/g,"ae").replace(/ö/g,"oe").replace(/ü/g,"ue").replace(/Ä/g,"Ae").replace(/Ö/g,"Oe").replace(/Ü/g,"Ue").replace(/ß/g,"ss");
                                    land = land.substring(0,3)+"0"
                                    application.activeWindow.command("f syf dbsm-stsl-buch-stb0.0004." + jhr + "." + land, true);
                                    F6710 = "!"+application.activeWindow.Variable("P3GPP")+"!"
                                    var gefunden = application.activeWindow.Status								//Status abfragen
    	                        if (gefunden == "NOHITS")
                                        {
                                        var sig = erg
                                        var ask = "Zu dieser Signatur wurde kein Bestandssatz (Tq) gefunden.\nBitte Signatur korrigieren oder abbrechen und \nzunächst Bestandssatz anlegen!"
                                        }
        		            application.activeWindow.closeWindow();	
                                    } while (gefunden == "NOHITS")	

                                break Erfolg;
                                }
                            else if (erg == 2 || erg == 3 || erg == 4) 
                                {
                                F2105 = "04,P01-s-62"
                                if (erg == 2)
                                    {
                                    F7100 = "IBST A "
                                    F6710 = "!1032375000!"
                                    } 
                                if (erg == 3)
                                    {
                                    F7100 = "IBST B "
                                    F6710 = "!1032374918!"
                                    } 
                                if (erg == 4)
                                    {
                                    F7100 = "IBST C "
                                    F6710 = "!103237487X!"
                                    } 
                                break Erfolg;
                                }
                            }
                        if (strKuerzel == "Sto")
                            {
                            var erg = __Pruef("Eine Frage","Was möchtest Du tun?\n 1 = Fachliteratur Be erfassen\n 2 = Sonstiges erfassen\n 3 = Scheemann bauen","1,2,3","1")
                            if (erg == null) return;
                            if (erg != 2) Status = "kurz"
                            if (erg == 1) 
                                {
                                SatzArt = "Aaa"
                                Slg = "F"
                                ITyp = "Text"
                                MTyp = "ohne Hilfsmittel zu benutzen"
                                DTyp = "Band"
                                F2105 = "04,P01-f-21"
                                F4030 = " "
                                F7109 = "(ESK Sdt)"
                                F8598 = "[Z-Klemm]"
                                var ask = "Welche Signatur?"
                                F7100 = "Be nnn, [nnn]"
                                Erfolg2: do 
                                    {
                                    var erg = __Prompter("Eine Frage",ask,F7100)
                                    if (!erg) break Erfolg2;
                                    F7100 = erg
                                    erg = erg.substring(erg.indexOf(" ")+1,erg.indexOf(","))
                                    if (erg.length < 4) {while (erg.length != 4) {erg = "0"+erg}}
                                    application.activeWindow.command("f syf dbsm-fb00-klem.b000.e000." + erg, true);
                                    F6710 = "!"+application.activeWindow.Variable("P3GPP")+"!"
                                    var gefunden = application.activeWindow.Status								//Status abfragen
    	                        if (gefunden == "NOHITS")
                                        {
                                        var ask = "Zu dieser Signatur wurde kein Bestandssatz (Tq) gefunden.\nBitte Signatur korrigieren oder abbrechen und \nzunächst Bestandssatz anlegen!"
                                        }
    	                        if (F7100.indexOf("n") > 0)
                                        {
                                        var ask = "Die Signatur ist nicht vollständig ausgefüllt.\nBitte zunächst exakte Signatur erfassen!"
                                        gefunden = "NOHITS"
                                        }
        		            application.activeWindow.closeWindow();	
                                    } while (gefunden == "NOHITS")	
		            //Erwerbungsart festlegen
                                    if(!eArt && SatzArt.indexOf("l") != 1)
                                        {
                                        var erg = __Pruef("Eine Frage","Welche Erwerbungsart?\n 1 = Kauf \n 2 = Tausch \n 3 = Geschenk\n 4 = Stiftung\n 5 = Depositum\n 6 = Leihgabe\n 7 = Fund\n 8 = Altbestand\n 9 = Pflicht","1,2,3,4,5,6,7,8,9","3")
                                        if (!erg) break Erfolg;
		                if (erg == 1) eArt = "ka";
		                if (erg == 2) eArt = "ta";
		                if (erg == 3) eArt = "ge";
		                if (erg == 4) eArt = "st";
		                if (erg == 5) eArt = "de";
		                if (erg == 6) eArt = "lg";
		                if (erg == 7) eArt = "fu";
		                if (erg == 8) eArt = "ab";
		                if (erg == 9) eArt = "pf";
                                        }
//                                     var day = jetzt.getDate();
//                                     vday = "0" + day
//                                     if (vday.length > 2) vday = vday.substring(1,vday.length)
//                                     var month = jetzt.getMonth() + 1;
//                                     month = "0" + month
//                                     if (month.length > 2) month = month.substring(1,month.length)
//                                     var vjahr = "#" + jahr
//                                     vjahr = vjahr.substring(3,vjahr.length)
//                                     Erfolg3: do
//                                         {
//                                         var com = "f lnr zklemm" + jahr + "* ser " + vday + "-" + month + "-" + vjahr;
//                                         application.activeWindow.command(com, false);
//                                         if (application.activeWindow.Status == "NOHITS") 
//                                             {
//                                             day = day - 1
//                                             vday = "0" + day
//                                             if (vday.length > 2) vday = vday.substring(1,vday.length)
//                                             }
//                                         else
//                                             {
//                                             application.activeWindow.command("k", false);
//                                             F8598 = application.activeWindow.title.findTag("8598",0,false,true);
//                                             num = parseInt(F8598.substring(F8598.lastIndexOf("0")+1,F8598.length))
//                                             num = num + 1
//                                             F8598 = F8598.substring(0,F8598.lastIndexOf("0")+1) + num.toString()
//                                             application.messageBox("",F8598,"")
//                                             application.activeWindow.simulateIBWKey("FE");
//                                             break Erfolg3
//                                             }
//                                         } while (day > 0)
                                break Erfolg;
                                }
                            else if (erg == 3)
                                {
                                application.messageBox("So ein Quatsch!","Das geht doch nicht ohne Schnee!!!","")
                                application.activeWindow.simulateIBWKey("FE");
                                return
                                }
                             var erg = 0
                             }
                        if (strKuerzel == "Lo")
                            {
                            var erg = __Pruef("Eine Frage","Sollen die Inhalte vom aktuellen Datensatz übernommen werden?\n 1 = pauschal übernehmen\n 2 = übernehmen, aber einzel abfragen\n 3 = nicht übernehmen","1,2,3","1")
                            if (erg == null) return;
                            if (erg == 1 || erg == 2)
                                {
                                if (erg == 1) abbruch = "j"
//                                application.activeWindow.simulateIBWKey("FE");
                                application.activeWindow.command("k", false);
                                F3100 =  application.activeWindow.title.findTag("3100",0,false,true);
                                F3000 =  application.activeWindow.title.findTag("3000",0,false,true);
                                F3000 =  application.activeWindow.title.findTag("3010",0,false,true);
                                F3000 =  application.activeWindow.title.findTag("3019",0,false,true);
                                if (F3000.indexOf("$BHrst") > 0) F3000 = F3000.substring(0,F3000.indexOf("$BHrst")-1) 
                                VPer = application.activeWindow.title.findTag("4000",0,false,true);
                                if  (VPer.indexOf("Papiermacher") > 0) VPer = VPer.substring(VPer.indexOf("Papiermacher")+13,VPer.length - 1);
                                else VPer = "..."
                                F4046 =  application.activeWindow.title.findTag("4030",0,false,true);
                                F4046 =  application.activeWindow.title.findTag("4046",0,false,true);
                                if (F4046.indexOf("[") == 0) F4046 = F4046.substring(1,F4046.length-1);
                                F4105 =  application.activeWindow.title.findTag("4105",0,false,true);
                                F4201 =  application.activeWindow.title.findTag("4201",0,false,true);
                                F510X = application.activeWindow.title.findTag("5100",0,false,true);
                                F510X =  F510X + "\n5101 " + application.activeWindow.title.findTag("5101",0,false,true);
                                F510X =  F510X + "\n5102 " + application.activeWindow.title.findTag("5102",0,false,true);
                                F5320 =  application.activeWindow.title.findTag("5320",0,false,true);
                                F5320 =  F5320 + "\n5320 " + application.activeWindow.title.findTag("5320",1,false,true);
                                F5320 =  F5320 + "\n5320 " + application.activeWindow.title.findTag("5320",2,false,true);
		        F5590 =  application.activeWindow.title.findTag("5590",1,false,true);
            	        F5591 =  application.activeWindow.title.findTag("5591",1,false,true);
		        F5592 =  application.activeWindow.title.findTag("5592",1,false,true);
                                F6710 =  application.activeWindow.title.findTag("6710",0,false,true);
 //                               F6800 =  application.activeWindow.title.findTag("6800",1,false,true);
                                F7100 =  application.activeWindow.title.findTag("7100",0,false,true);
                                F7100 = F7100.substring(0,F7100.lastIndexOf("/")) 
//                                num = parseInt(F7100.substring(F7100.lastIndexOf("/")+1,F7100.length))
//                                num = num + 1
//                                F7100 = F7100.substring(0,F7100.lastIndexOf("/")+1) + num.toString() + "/"
                                F7100 = application.activeWindow.title.findTag("7100 ",0,false,true);
                                F8100 = application.activeWindow.title.findTag("8100",0,false,true);
                                num = parseInt(F8100.substring(F8100.lastIndexOf("-")+1,F8100.length), 10)
                                num = num + 1
                                num = num.toString()
                                if (num.length == 7) {F8100 = F8100.substring(0,F8100.length-7) + num}
                                else if (num.length == 6) {F8100 = F8100.substring(0,F8100.length-6) + num}
                                else if (num.length == 5) {F8100 = F8100.substring(0,F8100.length-5) + num}
                                else if (num.length == 4) {F8100 = F8100.substring(0,F8100.length-4) + num}
                                else if (num.length == 3) {F8100 = F8100.substring(0,F8100.length-3) + num}
                                else if (num.length == 2) {F8100 = F8100.substring(0,F8100.length-2) + num}
                                else if (num.length == 1) {F8100 = F8100.substring(0,F8100.length-1) + num}
                                F8510 = application.activeWindow.title.findTag("8510",0,false,true);
                                application.activeWindow.simulateIBWKey("FE");
 //                               application.activeWindow.command("e", false);
                                num = 0
                                } 
                            var erg = 6
                            }
		else var erg = __Pruef("Eine Frage","Was soll erfasst werden?\n 1 = Publikation Studiensammlung \n 2 = Fachliteratur selbständig \n 3 = Fachliteratur unselbständig\n 4 = Handschrift\n 5 = Bild / Grafik\n 6 = WZ-Beleg / Papierprobe \n 7 = Buntpapier\n 8 = Brief / Archivalie\n 9 = Konvolut / Sammlung\n 10 = museales Objekt","1,2,3,4,5,6,7,8,9,10","")
		//Satzart festlegen
		if (erg == null) return;
 		if (erg == 1 || erg == 2) SatzArt = "ABSO";
		if (erg == 3) SatzArt = "Alxo";
		if (erg == 4) SatzArt = "H";
		if (erg == 5 || erg == 6 || erg == 7) 
			{
			SatzArt = "P";
			if (erg == 5) OGattung = "Bild"
			if (erg == 6) OGattung = "Wasserzeichen-Beleg"
			if (erg == 7) OGattung = "Buntpapier"
			}
		if (erg == 8) SatzArt = "D";
		if (erg == 9) SatzArt = "Qd";
		if (erg == 10) SatzArt = "X";
		//Sammlung festlegen
		if (erg == 1) Slg = "M";
		if (erg == 2) Slg = "F";
		if (erg == 4 || erg == 5 || erg == 6 || erg == 7 || erg == 8  || erg == 10) Slg = "S";
		var erg = 0
                        // bei Alxo vorsorglich Werte des letzten Titels einsammeln
		if (SatzArt == "Alxo") 
			{
                                    F1505 = "$erda"
                                    application.activeWindow.command("k", false);
                                    if (application.activeWindow.title) 
                                        {
                                        test = application.activeWindow.title.findTag("0500",0,false,true) 
                                        if(test ==  "Alxo") 
                                            {
                                            F1100 = application.activeWindow.title.findTag("1100",0,false,true);
                                            F1500 = application.activeWindow.title.findTag("1500",0,false,true);
                                            F4070 =  application.activeWindow.title.findTag("4070",0,false,true);
                                            F4070 = F4070.substring(0,F4070.indexOf("/p"))                                            
                                            F4241 =  application.activeWindow.title.findTag("4241",0,false,true);
                                            }
                                        test = ""
                                        application.activeWindow.simulateIBWKey("FE");
                                        }
                                    }
		//Publikationsart festlegen
		if (SatzArt == "ABSO") 
			{
			var erg = __Pruef("Eine Frage","Welche Publikationsart?\n 1 = Printpublikation \n 2 = AV-Medium \n 3 = elektronische Publikation auf Datenträger\n 4 = Online-Publikation","1,2,3,4","1")
			if (!erg) return;
			if (erg == 1) SatzArt = "A";
			if (erg == 2) SatzArt = "B";
			if (erg == 3) SatzArt = "S";
			if (erg == 4) SatzArt = "O";
			var erg = 0
			}
		//Inhaltstyp festlegen
		if (SatzArt == "A" || SatzArt == "Alxo" || SatzArt == "H" || SatzArt == "D") ITyp = "Text$btxt";
		if (SatzArt == "P") ITyp = "unbewegtes Bild$bsti";
		if (SatzArt == "X") ITyp = "dreidimensionale Form";
		if (SatzArt == "B" || SatzArt == "Qd" || SatzArt == "O" || SatzArt == "S") 
			{
			var erg = __Pruef("Eine Frage","Welcher Inhaltstyp?\n 1 = Text \n 2 = unbewegtes Bild \n 3 = zweidimensionales bewegtes Bild \n 4 = dreidimensionale Form \n 5 = kartografisches Bild\n 6 = gesprochenes Wort\n 7 = Computerdaten\n 8 = aufgeführte Musik\n ACHTUNG! Weitere Inhaltstypen siehe Vorgabewert-Tabelle!","1,2,3,4,5,6,7,8","1")
			if (!erg) break Erfolg;
			if (erg == 1) ITyp = "Text$btxt";
			if (erg == 1) FormF1 = "f1-text";
			if (erg == 2) ITyp = "unbewegtes Bild$bsti";
			if (erg == 3) ITyp = "zweidimensionales bewegtes Bild$btdi";
			if (erg == 4) ITyp = "dreidimensionale Form$btdf";
			if (erg == 5) ITyp = "kartografisches Bild$bcri";
			if (erg == 6) ITyp = "gesprochenes Wort$bspw";
			if (erg == 7) ITyp = "Computerdaten$bcod";
			if (erg == 8) ITyp = "aufgeführte Musik$bprm";
			var erg = 0
			}
		//Medientyp festlegen
		if (SatzArt == "A" || SatzArt == "Alxo" || SatzArt == "H" || SatzArt == "D" || SatzArt == "P" || SatzArt == "X") MTyp = "ohne Hilfsmittel zu benutzen$bn";
		if (SatzArt == "O" || SatzArt == "S") MTyp = "Computermedien$bc";
		if (SatzArt == "B" || SatzArt ==  "Qd" || SatzArt == "X")
			{
			if (SatzArt == "B") var erg = __Pruef("Eine Frage","Welcher Medientyp?\n 1 = audio \n 2 = video \n 3 = projizierbar \n 4 = stereografisch \n 5 = Mikroform \n 6 = mikroskopisch","1,2,3,4,5,6","1")
			if (SatzArt == "Qd") var erg = __Pruef("Eine Frage","Welcher Medientyp?\n 1 = audio \n 2 = video \n 3 = projizierbar \n 4 = stereografisch \n 5 = Mikroform \n 6 = mikroskopisch\n 7 = Computermedien\n 8 = ohne Hilfsmittel zu benutzen\n 9 = Sonstige \n 10 = nicht spezifiziert","1,2,3,4,5,6,7,8,9,10","8")
			if (SatzArt == "X") var erg = __Pruef("Eine Frage","Welcher Medientyp?\n 8 = ohne Hilfsmittel zu benutzen\n 9 = Sonstige \n 10 = nicht spezifiziert","8,9,10","8")
			if (!erg) break Erfolg;
			if (erg == 1) MTyp = "audio$bs";
			if (erg == 2) MTyp = "video$bv";
			if (erg == 3) MTyp = "projizierbar$bg";
			if (erg == 4) MTyp = "stereografisch$be";
			if (erg == 5) MTyp = "Mikroform$bh";
			if (erg == 6) MTyp = "mikroskopisch$bp";
			if (erg == 7) MTyp = "Computermedien$bc";
			if (erg == 8) MTyp = "ohne Hilfsmittel zu benutzen$bn";
			if (erg == 9) MTyp = "Sonstige$bx";
			if (erg == 10) MTyp = "nicht spezifiziert$bz";
			var erg = 0
			}
		//Datenträgertyp festlegen
		// wenn Medientyp "audio"
		if (MTyp == "audio$bs") 
			{
			var erg = __Pruef("Eine Frage","Welcher Datenträgertyp?\n 1 = Audiocartridge \n 2 = Audiodisk \n 3 = Audiokassette \n 4 = Notenrolle\n 5 = Phonographenzylinder\n 6 = Tonbandspule\n 7 = Tonspurspule\n 8 = Sonstige","1,2,3,4,5,6,7,8,9","5")
			if (!erg) break Erfolg;
			if (erg == 1) 
				{
				DTyp = "Audiocartridge$bsg";
				DTMaterial = "To-sonst" 
				}
			if (erg == 2) 
				{
				DTyp = "Audiodisk$bsd";
				FormF2 = "f2-schei"
				var erg = 0
				var erg = __Pruef("Eine Frage","Welches Trägermaterial?\n 1 = Audio-CD \n 2 = Audio-DVD \n 3 = Schallplatte","1,2,3","1")
				if (!erg) break Erfolg;
				if (erg == 1) DTMaterial = "To-cdda" 
				if (erg == 2) DTMaterial = "To-dvda"
				if (erg == 3) DTMaterial = "To-scha"
				}
			if (erg == 3) 
				{
				DTyp = "Audiokassette$bss";
				DTMaterial = "To-tonks" 
				}
			if (erg == 4) 
				{
				DTyp = "Notenrolle$bsq";
				DTMaterial = "To-rolle" 
				}
			if (erg == 5) 
				{
				DTyp = "Phonographenzylinder$bse";
				DTMaterial = "To-zyl" 
				}
			if (erg == 6) 
				{
				DTyp = "Tonbandspule$bst";
				DTMaterial = "To-tonbd" 
				}
			if (erg == 7) 
				{
				DTyp = "Tonspurspule$bsi";
				DTMaterial = "To-tonspur" 
				}
			if (erg == 8) 
				{
				DTyp = "Sonstige Tonträger$bsz";
				var erg = 0
				var erg = __Pruef("Eine Frage","Welches Trägermaterial?\n 1 = Schallplatte \n 2 = sonstige","1,2","1")
				if (!erg) break Erfolg;
				if (erg == 1) DTMaterial = "To-scha" 
				if (erg == 2) DTMaterial = "To-sonst"
				}
			}
		// wenn Medientyp "video"
		if (MTyp == "video$bv") 
			{
			var erg = __Pruef("Eine Frage","Welcher Datenträgertyp?\n 1 = Videobandspule \n 2 = Videocartridge \n 3 = Videodisk \n 4 = Videokassette\n 5 = Sonstige","1,2,3,4,5","4")
			if (!erg) break Erfolg;
			if (erg == 1) 
				{
				DTyp = "Videobandspule$bvr";
				DTMaterial = "BT-anfi" 
				}
			if (erg == 2) 
				{
				DTyp = "Videocartridge$bvc";
				DTMaterial = "BT-modul" 
				}
			if (erg == 3) 
				{
				DTyp = "Videodisk$bvd";
				FormF2 = "f2-schei"
				var erg = 0
				var erg = __Pruef("Eine Frage","Welches Trägermaterial?\n 1 = Blu-ray-Disc \n 2 = Video-DVD","1,2","1")
                                                if (!erg) break Erfolg;
				if (erg == 1) DTMaterial = "BT-bray" 
				if (erg == 2) DTMaterial = "BT-dvdv"
				}
			if (erg == 4) 
				{
				DTyp = "Videokassette$bvf";
				DTMaterial = "BT-vika" 
				}
			if (erg == 5) 
				{
				DTyp = "Sonstige Videodatenträger$bvz";
				DTMaterial = "BT-sonst" 
				}
			}
		// wenn Medientyp "projizierbar"
		if (MTyp == "projizierbar$bg") 
			{
			var erg = __Pruef("Eine Frage","Welcher Datenträgertyp?\n 1 = Dia \n 2 = Filmdose \n 3 = Filmkassette \n 4 = Filmrolle\n 5 = Filmspule\n 6 = Filmstreifen\n 7 = Filmstreifen für Einzelbildvorführung\n 8 = Filmstreifen-Cartridge\n 9 = Overheadfolie\ 10 = Sonstige","1,2,3,4,5,6,7,8,9","1")
			if (!erg) break Erfolg;
			if (erg == 1) 
				{
				DTyp = "Dia$bgs";
				var erg = 0
				var erg = __Pruef("Eine Frage","Welches Trägermaterial?\n 1 = Dia-Positiv \n 2 = Dia-Negativ","1,2","1")
				if (!erg) break Erfolg;
				if (erg == 1) DTMaterial = "TBH-fotop" 
				if (erg == 2) DTMaterial = "TBH-foton"
				}
			if (erg == 2 || erg == 3 || erg == 4 || erg == 5 || erg == 6 || erg == 8) DTMaterial = "BT-anfi"
			if (erg == 2) DTyp = "Filmdose$bmc";
			if (erg == 3) DTyp = "Filmkassette$bmf";
			if (erg == 4) DTyp = "Filmrolle$bmo";
			if (erg == 5) DTyp = "Filmspule$bmr";
			if (erg == 6) DTyp = "Filmstreifen$bgf";
			if (erg == 7) 
				{
				DTyp = "Filmstreifen für Einzelbildvorführung$bgd";
				var erg = 0
				var erg = __Pruef("Eine Frage","Welches Trägermaterial?\n 1 = Positiv-Film \n 2 = Negativ-Film","1,2","1")
				if (!erg) break Erfolg;
				if (erg == 1) DTMaterial = "TBH-fotop" 
				if (erg == 2) DTMaterial = "TBH-foton"
				}
			if (erg == 8) DTyp = "Filmstreifen-Cartridge$bgc";
			if (erg == 9) 
				{
				DTyp = "Overheadfolie$bgt";
				DTMaterial = "TBH-arbtrans";
				}
			if (erg == 10) 
				{
				DTyp = "Sonstige projizierbare Bilder$bmz";
				DTMaterial = "TBH-sonst"
				}
			}
		// wenn Medientyp "stenografisch"
		if (MTyp == "stereografisch$be") 
			{
			var erg = __Pruef("Eine Frage","Welcher Datenträgertyp?\n 1 =  Stereobild\n 2 =  Stereografische Disk\n 3 = Sonstige","1,2,3","1")
                                    if (!erg) break Erfolg;
			if (erg == 1) DTyp = "Stereobild$beh";
			if (erg == 2) DTyp = "Stereografische Disk$bes";
			if (erg == 2) DTyp = "Sonstige stereografische Datenträger$bez";
			}
		// wenn Medientyp "Mikroform"
		if (MTyp == "Mikroform$bh") 
			{
			var erg = __Pruef("Eine Frage","Welcher Datenträgertyp?\n 1 = Lichtundurchlässiger Mikrofiche \n 2 = Mikrofiche \n 3 = Mikrofichekassette \n 4 = Mikrofilm-Cartridge\n 5 = Mikrofilmkassette\n 6 = Mikrofilmlochkarte\n 7 = Mikrofilmrolle\n 8 = Mikrofilmspule\n 9 = Mikrofilmstreifen\n 10 = Sonstige","1,2,3,4,5,6,7,8,9,10","2")
			if (!erg) break Erfolg;
			if (erg == 1) 
				{
				DTyp = "Lichtundurchlässiger Mikrofiche$bha";
				DTMaterial = "Mi-ckop-lud"
				}
			if (erg == 2) 
				{
				DTyp = "Mikrofiche$bhe";
				DTMaterial = "Mi-ckop"
				}
			if (erg == 3) 
				{
				DTyp = "Mikrofichekassette$bhf";
				DTMaterial = "Mi-ckop-kass"
				}
			if (erg == 4) 
				{
				DTyp = "Mikrofilm-Cartridge$bhb";
				DTMaterial = "Mi-lkop-car"
				}
			if (erg == 5) 
				{
				DTyp = "Mikrofilmkassette$bhc";
				DTMaterial = "Mi-lkop-kass"
				}
			if (erg == 6) 
				{
				DTyp = "Mikrofilmlochkarte$bhg";
				DTMaterial = "Mi-lkop-karte"
				}
			if (erg == 7) 
				{
				DTyp = "Mikrofilmrolle$bhj";
				DTMaterial = "Mi-lkop"
				}
			if (erg == 8) 
				{
				DTyp = "Mikrofilmspule$bhd";
				DTMaterial = "Mi-lkop-spule"
				}
			if (erg == 9) 
				{
				DTyp = "Mikrofilmstreifen$bhh";
				DTMaterial = "Mi-lkop-streifen"
				}
			if (erg == 10) 
				{
				DTyp = "Sonstige Mikroformen$bhz";
				DTMaterial = "Mi-sonst"
				}
			}
		// wenn Medientyp "mikroskopisch"
		if (MTyp == "mikroskopisch$bp") DTyp = "Objektträger$bpp";
		// wenn Medientyp "Computermedien"
		if (MTyp == "Computermedien$bc") 
			{
			if (SatzArt == "O" ) 
				{
				DTyp = "Online-Ressource$bcr";
				DTMaterial = "O-cofz"
				}
			if (SatzArt == "S") 
				{
				var erg = __Pruef("Eine Frage","Welcher Datenträgertyp?\n 1 = Computerchip-Cartridge \n 2 = Computerdisk \n 3 = Computerdisk-Cartridge \n 4 = Magnetbandcartridge \n 5 = Magnetbandkassette \n 6 = Magnetbandspule\n 7 = Speicherkarte\n 8 = Sonstige","1,2,3,4,5,6,7,8","1")
				if (!erg) break Erfolg;
				if (erg == 1) 
					{
					DTyp = "Computerchip-Cartridge$bcb";
					DTMaterial = "Da-ccart"
					}
				if (erg == 2) 
					{
					DTyp = "Computerdisk$bcd";
					FormF2 = "f2-schei"
					var erg = 0
					var erg = __Pruef("Eine Frage","Welches Trägermaterial?\n 1 = Diskette \n 2 = CD-ROM\n 3 = DVD-ROM","1,2,3","1")
                                                            if (!erg) break Erfolg;
					if (erg == 1) DTMaterial = "Da-disk" 
					if (erg == 2) DTMaterial = "Da-crom"
					if (erg == 3) DTMaterial = "Da-dvdr"
					}
				if (erg == 3) 
					{
					DTyp = "Computerdisk-Cartridge$bcb";
					DTMaterial = "Da-dcart"
					}
				if (erg == 4) 
					{
					DTyp = "Magnetbandcartridge$bca";
					DTMaterial = "Da-datbndcart"
					}
				if (erg == 5) 
					{
					DTyp = "Magnetbandkassette$bhc";
					DTMaterial = "Da-datbndkass"
					}
				if (erg == 6) 
					{
					DTyp = "Magnetbandspule$bch";
					DTMaterial = "Da-datbndspule"
					}
				if (erg == 7) 
					{
					DTyp = "Speicherkarte$bck";
					DTMaterial = "Da-karte"
					}
				if (erg == 8) 
					{
					DTyp = "Sonstige Computermedien$bcz";
					DTMaterial = "Da-sonst"
					}
				}
			}
		//Originalität
		if (Slg == "S" || Slg == "M") 
		    {
		    if (OGattung == "Wasserzeichen-Beleg") var erg = __Pruef("Eine Frage","Noch eine Frage zur Originalität:\n 1 = Original \n 2 =  Kopie bzw. Pause","1,2","1");
		    else var erg = __Pruef("Eine Frage","Noch eine Frage zur Originalität:\n 1 = Original \n 2 =  Kopie / Replik / Faksimile / Nachbildung  (originalgetreue Nachbildung einer Vorlage)\n 3 = Modell   (verkleinerte Nachbildung eines Originals)","1,2,3","1");
		    if (!erg) break Erfolg;
		    if (erg == 1) 
                                {
                                FormO = "o-org"
                                if (OGattung == "Wasserzeichen-Beleg") F5592 = ""
                                }
		    if (erg == 2) FormO = "o-kopie"
		    if (erg == 3) FormO = "o-modell"
		    var erg = 0
		    }
		// wenn Medientyp "Computermedien"
		if (MTyp == "ohne Hilfsmittel zu benutzen$bn") 
			{
			if (SatzArt == "A" || SatzArt == "Alxo" || SatzArt == "H") DTyp = "Band$bnc";
			if (SatzArt == "D" || SatzArt == "P") DTyp = "Blatt$bnb";
			if (SatzArt == "X" || SatzArt == "Qd") 
				{
				var erg = __Pruef("Eine Frage","Welcher Datenträgertyp?\n 1 = Band \n 2 = Blatt \n 3 = Flipchart \n 4 = Gegenstand \n 5 = Karte \n 6 = Rolle\n 7 = Sonstige","1,2,3,4,5,6,7","4")
				if (!erg) break Erfolg;
				if (erg == 1) DTyp = "Band$bnc";
				if (erg == 2) DTyp = "Blatt$bnb";
				if (erg == 3) DTyp = "Flipchart$bnn";
				if (erg == 4) DTyp = "Gegenstand$bnr";
				if (erg == 5) DTyp = "Karte$bno";
				if (erg == 6) DTyp = "Rolle$bna";
				if (erg == 7) DTyp = "Sonstige Datenträger, die ohne Hilfsmittel zu benutzen sind$bnz";
				}
			//Trägermaterial 1130 festlegen
			var erg = 0
			if (Slg != "F" && SatzArt != "Alxo" && (DTyp == "Band$bnc" || DTyp == "Blatt$bnb" || DTyp == "Karte$bno" || DTyp == "Rolle$bna")) 
				{
				if (OGattung == "Wasserzeichen-Beleg") var erg = __Pruef("Eine Frage","Welches Trägermaterial hat das Original-Papier?\n 1 = Papier (unspezifisch) \n 2 = handgeschöpftes Papier (unspezifisch)\n 3 = handgeschöpftes gegittertes Papier \n 4 = handgeschöpftes geripptes Papier \n 5 = handgeschöpftes Velin-Papier\n 6 = handgeschöpftes Zeilen-Papier\n 7 = maschinell gefertigtes Papier (unspezifisch)\n 8 = maschinell gefertigtes geripptes Papier\n 9 = maschinell gefertigtes gegittertes Papier\n 10 = maschinell gefertigtes Velin-Papier\n 11 = maschinell gefertigtes Zeilen-Papier","1,2,3,4,5,6,7,8,9,10,11","4")
				else var erg = __Pruef("Eine Frage","Welches Trägermaterial?\n 1 = Papier (unspezifisch) \n 2 = handgeschöpftes Papier (unspezifisch)\n 3 = handgeschöpftes gegittertes Papier \n 4 = handgeschöpftes geripptes Papier \n 5 = handgeschöpftes Velin-Papier\n 6 = handgeschöpftes Zeilen-Papier\n 7 = maschinell gefertigtes Papier (unspezifisch)\n 8 = maschinell gefertigtes geripptes Papier\n 9 = maschinell gefertigtes gegittertes Papier\n 10 = maschinell gefertigtes Velin-Papier\n 11 = maschinell gefertigtes Zeilen-Papier\n 12 = transparente Kunststoff-Folie (außer Arbeitstransparent)\n 13 = Pergament\n 14 = Papyrus\n 15 = Sonstige (z. B. Leder, Holz, PVC, Pappe, Metall usw.)","1,2,3,4,5,6,7,8,9,10,11,12,13,14,15","1")
				if (!erg) break Erfolg;
                                                if (OGattung == "Wasserzeichen-Beleg" && (erg == 2 || erg == 3 || erg == 4 || erg == 5 || erg == 6)) F4201 = "Handbüttenpapier ";
                                                if (OGattung == "Wasserzeichen-Beleg" && (erg == 7 || erg == 8 || erg == 9 || erg == 10 || erg == 11)) F4201 = "maschinell gefertigtes Papier ";
                                                if (OGattung == "Wasserzeichen-Beleg" && (erg == 4 || erg == 9)) F4201 = F4201 + "gerippt; Position des Wasserzeichens: ...";
                                                if (OGattung == "Wasserzeichen-Beleg" && (erg == 5 || erg == 10)) F4201 = F4201 + "ungerippt; Position des Wasserzeichens: ...";
                                                if (OGattung == "Wasserzeichen-Beleg" && (erg == 3 || erg == 8)) F4201 = F4201 + "gegittert; Position des Wasserzeichens: ...";
                                                if (OGattung == "Wasserzeichen-Beleg" && (erg == 6 || erg == 11)) F4201 = F4201 + "Zeilen-Papier; Position des Wasserzeichens: ...";
				if (erg == 1) DTMaterial = "TB-papier";
				if (erg == 2) DTMaterial = "TB-papier-h";
				if (erg == 3) DTMaterial = "TB-papier-hg";
				if (erg == 4) DTMaterial = "TB-papier-hr";
				if (erg == 5) DTMaterial = "TB-papier-hv";
				if (erg == 6) DTMaterial = "TB-papier-hz";
				if (erg == 7) DTMaterial = "TB-papier-m";
				if (erg == 8) DTMaterial = "TB-papier-mg";
				if (erg == 9) DTMaterial = "TB-papier-mr";
				if (erg == 10) DTMaterial = "TB-papier-mv";
				if (erg == 11) DTMaterial = "TB-papier-mz";
				if (erg == 12) DTMaterial = "TB-folie";
				if (erg == 13) DTMaterial = "TB-perg";
				if (erg == 14) DTMaterial = "TB-papy";
				if (erg == 15) DTMaterial = "TB-sonst";
				var erg = 0
                                                if (OGattung == "Wasserzeichen-Beleg")  
                                                    {
                                                    if (OGattung == "Wasserzeichen-Beleg" && FormO == "o-kopie")  
                                                        {
                                                        var erg = __Pruef("Eine Frage","Um welchen Kopie-Typ handelt es sich?\n 1 = Handpause auf Transparentpapier\n 2 = Handpause auf Kunststofffolie\n 3 = Xerokopie\n 4 = Fotografie / Ausdruck eines Digitalbildes\n5 = Lichtpause","1,2,3,4,5","1")
				        if (!erg) break Erfolg;
				        if (erg == 2) DTMaterial = DTMaterial + ";TB-folie";
				        if (erg == 1 || erg == 2) 
                                                            {
                                                            F5592 = "Handpause";
                                                            F4019 = "Wasserzeichenpause";
                                                            F4060 = "1 Wasserzeichenpause";
                                                            if (erg == 1) F4060 = F4060 + " auf Transparentpapier";
                                                            if (erg == 2) F4060 = F4060 + " auf Zeichenfolie";
                                                            }
				        if (erg == 3) 
                                                            {
                                                            F5592 = "!041903838!";
                                                            F4019 = "Wasserzeichenreproduktion";
                                                            F4060 = "1 Xerokopie einer Wasserzeichenreproduktion";
                                                            }
				        if (erg == 4) 
                                                            {
                                                            F4019 = "Wasserzeichenreproduktion";
                                                            F4060 = "1 Ausdruck einer Durchlichtabbildung"
                                                            F5592 = "Fotografie";
                                                            }
				        if (erg == 5) 
                                                            {
                                                            F5592 = "!041675932!";
                                                            F4060 = "1 Lichtpause";
                                                            F4019 = "Wasserzeichenpause";
                                                            }
				        var erg = 0
                                                        }
                                                    else
                                                        {
                                                        F4019 = "Originalpapier"
                                                        if (DTMaterial.indexOf("-h") > 0) F5590 = "!042474868!"  
                                                        }
                                                    }
				}
			if (DTyp == "Flipchart$bnn") DTMaterial = "TB-papier";
			if (DTyp == "Gegenstand$bnr")
				{
				var erg = __Pruef("Eine Frage","Hat der Gegenstand einen Inhalt?\n 1 = Text \n 2 = Bild \n 3 = kein Inhalt","1,2,3","3")
				if (!erg) break Erfolg;
				if (erg == 1) FormF1 = "f1-text";
				if (erg == 2) FormF1 = "f1-bild";
				if (erg == 3) FormF1 = "";
				if (erg == 3) FormA2 = "";
				if (erg == 3) FormE1 = "";
				if (erg == 3) DTMaterial = "";
				var erg = 0
				}
			if ((DTyp == "Gegenstand$bnr" && FormF1 != "") || DTyp == "Sonstige Datenträger, die ohne Hilfsmittel zu benutzen sind$bnz") 
				{
				var erg = __Pruef("Eine Frage","Welches Trägermaterial?\n 1 = Tontafel \n 2 = Wachstafel \n 3 = Foto-Glasplatte \n 4 = Sonstige (z. B. Stein, Leder, Holz, PVC, Pappe, Metall usw.)","1,2,3,4","4")
				if (!erg) break Erfolg;
				if (erg == 1) DTMaterial = "TB-ton";
				if (erg == 2) DTMaterial = "TB-wachs";
				if (erg == 3) DTMaterial = "TB-fotog";
				if (erg == 4) DTMaterial = "TB-sonst";
				}
			var erg = 0
			}
		//Art des Inhalts 1131 festlegen
		if (OGattung == "Wasserzeichen-Beleg" || OGattung == "Bild" || OGattung == "Buntpapier") ArtInhalt = "!041454197!"
		if (DTMaterial == "To-scha") ArtInhalt = "!040520323!"
		//Formangaben 1132 festlegen
		if (SatzArt != "O" || SatzArt != "S") FormA1 = "a1-analog"
		if (SatzArt == "O" || SatzArt == "S") FormA1 = "a1-digital"
		if (SatzArt == "D") 
			{
			var erg = __Pruef("Eine Frage","Was liegt vor?\n 1 = Manuskript \n 2 = Typuskript \n 3 = digitaler 'Brief' (E-Mail/SMS/...)","1,2,3","1")
			if (!erg) break Erfolg;
			if (erg == 1) FormA2 = "a2-hand"
			if (erg == 2) FormA2 = "a2-masch"
			if (erg == 3) FormA1 = "a1-digital" 
			}
		if (Slg != "F" && SatzArt == "A") 
			{
			var erg = __Pruef("Eine Frage","Welches Druckverfahren liegt vor?\n 1 = Hochdruck \n 2 = Tiefdruck \n 3 = Flachdruck \n 4 = Durchdruck  (z. B. Siebdruck ...)\n 5 = Nonimpact-Druck  (z. B. Tintenstrahl-Druck ...)\n 6 = Prägedruck\n 7 = Sonstige Druckverfahren / Druck  nicht spezifiziert","1,2,3,4,5,6,7","7")
			if (!erg) break Erfolg;
			if (erg == 1) FormA2 = "a2-druck-h"
			if (erg == 2) FormA2 = "a2-druck-t"
			if (erg == 3) FormA2 = "a2-druck-f" 
			if (erg == 4) FormA2 = "a2-druck-d"
			if (erg == 5) FormA2 = "a2-druck-n"
			if (erg == 6) FormA2 = "a2-druck-p"
			if (erg == 7) FormA2 = "a2-druck"
			var erg = 0
			}
		if (SatzArt == "A" || SatzArt == "H" || SatzArt == "D") FormF1 ="f1-text"
		if (MTyp == "audio$bs") FormF1 ="f1-ton";
		if (MTyp == "video$bv") FormF1 ="f1-film";
		if (SatzArt == "P") FormF1 = "f1-bild"
		if (DTyp == "Band$bnc") FormF2 = "f2-kodex"
		if (DTyp == "Blatt$bnb") FormF2 = "f2-blatt"
		if (DTyp == "Rolle$bna") FormF2 = "f2-rolle"
		if (DTyp == "Gegenstand$bnr") FormF2 = "f2-3d"
		if (!FormF2) 
			{
			var erg = __Pruef("Eine Frage","Welche äußere Form liegt vor?\n 1 = Blatt \n 2 = Kodex \n 3 = Leporello \n 4 = Rolle\n 5 = sonstige 2D-Form \n 6 = sonstige 3D-Form","1,2,3,4,5,6","6")
			if (erg == 1) FormF2 = "f2-blatt"
			if (erg == 2) FormF2 = "f2-kodex"
			if (erg == 3) FormF2 = "f2-lepo" 
			if (erg == 4) FormF2 = "f2-rolle"
			if (erg == 5) FormF2 = "f2-2d"
			if (erg == 6) FormF2 = "f2-3d"
			var erg = 0
			}
		//Erscheinungsweise und Veröffentlichungsart festlegen
		if (SatzArt == "Qd") FormV = "v-cont"
		if (SatzArt == "A" || SatzArt == "B" || SatzArt == "O" || SatzArt == "S") 
			{
			FormE2 = "e2-se"
			}
		else
			{
			if (SatzArt == "Alxo") FormE2 = "e2-un"
			if (SatzArt != "Alxo") FormE2 = "e2-uv"
			}
		if (SatzArt != "Alxo" && SatzArt != "Qd" && SatzArt != "X") 
		    {
                            if (OGattung == "Wasserzeichen-Beleg")
                                {
		        var erg = __Pruef("Eine Frage","Ganzer Bogen oder Blatt / Fragment?\n 1 = Ganzer Bogen \n 2 = Halbbogen / Blatt\n3 = Fragment","1,2,3","3")
		        if (!erg) break Erfolg;
		        if (erg == 1) 
                                    {
                                    FormV = "v-ganz";
                                    F4060 = "1 unbeschnittener Originalbogen"
                                    }
		        if (erg == 2 || erg == 3) 
                                    {
                                    FormV = "v-frag";
                                    if (erg == 2) F4060 = " 1 halber Originalbogen = 1 Blatt"
                                    if (erg == 3 && !F4060) F4060 = "1 Fragment = Blatt"
                                    if (erg == 3) F4019 = "Wasserzeichen-Fragment"
                                    }
		        SatzArt = SatzArt + "a";
                                }
                            else
                                {
		        var erg = __Pruef("Eine Frage","Welche Erscheinungsweise?\n 1 = einteilig (Monographie) \n 2 = mehrteilig mit eigenem Titel \n 3 = mehrteilig ohne eigenen Titel \n4 = unselbständig","1,2,3,4","1")
		        if (!erg) break Erfolg;
		        if (erg == 1) 
			{
			FormV = "v-ganz";
			if (SatzArt == "A" || SatzArt == "S" || SatzArt == "O") FormE1 = "e1-ae";
			SatzArt = SatzArt + "a";
			}
		        if (erg == 2 || erg == 3) 
			{
			FormV = "v-teil";
			if (SatzArt == "A" || SatzArt == "S" || SatzArt == "O") FormE1 = "e1-am";
			if (erg == 2) SatzArt = SatzArt + "F";
			if (erg == 3) SatzArt = SatzArt + "f";
			}
		        if (erg == 4) 
			{
			FormV = "v-teil";
			FormE1 = "e1-uv";
			SatzArt = SatzArt + "l";
			}
                                }
		        var erg = 0
		    }
		// noch undefinierte Typen
		if (!FormV) 
			{
			if (SatzArt == "Alxo") 
				{
				FormV = ""
				}
			else
				{
				var erg = __Pruef("Eine Frage","Ist das Objekt\n 1 = ein ungeteiltes Ganzes \n 2 = ein Teil einer größeren Einheit \n 3 = ein Fragment?","1,2,3","1")                            
				if (!erg) break Erfolg;
				if (erg == 1)  FormV = "v-ganz";
				if (erg == 1)  SatzArt = SatzArt  + "a";
				if (erg == 2)  FormV = "v-teil";
				if (erg == 2)  SatzArt = SatzArt  + "F";
				if (erg == 3)  FormV = "v-frag";
				if (erg == 3)  SatzArt = SatzArt  + "a";
				}
			}
			if (!FormA1) FormA1 = ""
			if (!FormA2) FormA2 = ""
			if (!FormE1) FormE1 = ""
			if (!FormE2) FormE2 = ""
			if (!FormF1) FormF1 = ""
			if (!FormF2) FormF2 = ""
			if (!FormO) FormO = ""
			if (!ArtInhalt) ArtInhalt = ""
			if (!OGattung) OGattung = ""        
                        //Gestaltungsmerkmale  
                        if (Slg != "F" && SatzArt.indexOf("l") != 1)
                            {
                            //Objektgattung
                            if (OGattung == "Bild") searchTerm = "Bildliche Darstellung";
                            if (OGattung == "Buntpapier") searchTerm = "Buntpapier";
                            if (OGattung == "Wasserzeichen-Beleg") F5590 = "!0040648273!";
                            if (SatzArt.indexOf("D") == 0) searchTerm = "Brief";
                            if (SatzArt.indexOf("H") == 0) searchTerm = "Handschrift";
                            if (!F5590) searchTerm = __Prompter("Eine Frage","Welche Objektgattung?",searchTerm,"");
                            if (searchTerm != "" && searchTerm != null) 
                                {
                                if (!F4019) F4019 = searchTerm;
                                if (searchTerm == "Buntpapier") F5590 = "!040090809!";
                                else if (searchTerm == "Brief") F5590 = "!040082407!";
                                else if (searchTerm == "Handschrift") F5590 = "!040232875!";
                                else if (searchTerm == "Bildliche Darstellung") F5590 = "!041454197!";
                                else
                                    {
                                    searchTerm = searchTerm + "? bbg ts?";
                                    search = "f sw ";
                                    vTag = "5590"
                                    F5590 = "!"+DBSMLink(search, searchTerm, vTag)+"!";
                                    }
                                }
                           if (OGattung == "Wasserzeichen-Beleg" && abbruch == "j") break Erfolg;

                            //Entstehungsort
                            var vOrt = "Verlagsort";
                            if (Slg == "S") vOrt = "Entstehungsort";
                            searchTerm = "";
                            if (F5591) searchTerm = F5591 
                            searchTerm = __Prompter("Eine Frage","Welcher " + vOrt + "?",searchTerm,"");
                            if (searchTerm != "" && searchTerm != null && searchTerm.indexOf("!") != 0) 
                                {
                                if (Slg == "S") F4046 = searchTerm
                                else F4030 = searchTerm
                                searchTerm1 = searchTerm + "? bbg tg?";
                                search = "f sw ";
                                vTag = "5591"
                                F5591 = "!"+DBSMLink(search, searchTerm1, vTag)+"!";
                                if (OGattung == "Wasserzeichen-Beleg") 
                                    {
                                    searchTerm = "Papiermühle " + searchTerm
                                    if (F3100) searchTerm = F3100 
                                    searchTerm = __Prompter("Eine Frage","Welche Papiermühle?",searchTerm,"");
                                    if (searchTerm != "" && searchTerm != null)
                                        {
                                        vTag = "3100"
                                        searchTerm + "? bbg tb?";
                                        F3100 = "!"+DBSMLink(search, searchTerm, vTag)+"!";
                                        }
                                    }
                                }
                            //Verwendungsort
                            if (OGattung == "Wasserzeichen-Beleg") 
                                {
                                searchTerm = "";
                                if (F6800) searchTerm = F6800 
                                searchTerm = __Prompter("Eine Frage","Welcher Verwendungsort?",searchTerm,"");
                                if (searchTerm != "" && searchTerm != null && searchTerm.indexOf("!") != 0) 
                                    {
                                    searchTerm = searchTerm + "? bbg tg?";
                                    search = "f sw ";
                                    vTag = "6800"
                                    F6800 = "!"+DBSMLink(search, searchTerm, vTag)+"!";
                                    }
                                searchTerm = "";
                                if (F3000) searchTerm = F3000 
                                searchTerm = __Prompter("Eine Frage","Welcher Papiermacher?",searchTerm,"");
                                if (searchTerm != "" && searchTerm != null && searchTerm.indexOf("!") != 0) 
                                    {
                                    if (searchTerm.indexOf(",") > 0) VPer = searchTerm.substring(searchTerm.indexOf(",")+1,searchTerm.length) + " " +  searchTerm.substring(0, searchTerm.indexOf(","));
                                    else VPer = searchTerm
                                    searchTerm = searchTerm + "? rp papiermacher";
                                    search = "f sw ";
                                    vTag = "3000"
                                    F3000 = "!"+DBSMLink(search, searchTerm, vTag)+"!";
                                    }
                                }
                            }
                        //Systematik
                        searchTerm = ""
                        if (SatzArt == "Alxo" && strKuerzel != "MaF") searchTerm = "esk?"
                        if (SatzArt.indexOf("X") == 0) searchTerm = "og?"
                        if (OGattung == "Buntpapier") searchTerm = "bu?"
                        if (searchTerm)
                            {
                            search = "f syf "
                            vTag = "5320"
                            F5320 = "!"+DBSMLink(search, searchTerm, vTag)+"!";
                            searchTerm = ""
                            }
                        
                        if (OGattung == "Wasserzeichen-Beleg") 
                            {
                            if (F5320) searchTerm = F5320
                            else  searchTerm = "wz?"
                            searchTerm = __Prompter("Eine Frage","Geben Sie ein Suchwort für die WZ-Klassifikation ein.\n(Beispiele: Adler, Buchstabe, Krone ...)\nMit 'wz?' wird die gesamte Klassifikation zum Navigieren aufgeschlagen.",searchTerm,"")
                            }
                        if (searchTerm && searchTerm.indexOf("!") != 0)
                            {
                            if (searchTerm == "wz?") search = "f syf "
                            else 
                                {
                                search = "f syf wz? and syw "
                                searchTerm = searchTerm + "?"
                                }
                            vTag = "5320"
                            F5320 = "!"+DBSMLink(search, searchTerm, vTag)+"!";
                            }

		//Sammlungszugehörigkeit festlegen
		if(SatzArt.indexOf("l") != 1)
		    {
		    var F6710
		    if (Slg == "F")
	                    {
		        F2105 = "04,P01-f-21"
		        F8598 = "[Z-Klemm]";
		        var erg = __Pruef("Eine Frage","Welche Signaturgruppe?\n 1 = A Bibliographie und Allgemeines \n 2 = B Beschreibstoffe und Bedruckstoffe\n 3 = C Schrift\n 4 = D Buchdruck\n 5 = E Buchhandel\n 6 = F Bucheinband\n 7 = G Buchwissenschaft, Medienwissenschaft (einschließlich Sozialgeschichte des Lesens)\n 8 = H Bibliophilie\n 9 = J Originalgraphische Hochdruckverfahren (Holzschnitt u.ä.)\n10 = K Originalgraphische Tiefdruckverfahren (Kupferstich, Schabkunst, Stahlstich, Radierung, Aquatinta u.a.)\n11 = L Originalgraphische Flachdruckverfahren (Lithographie)\n12 = M Fotografie\n13 = N Reproduktionsverfahren : chemische, photomechanische u. a. Druckverfahren\n14 = O Bibliothekswesen\n15 = P Museums- und Ausstellungswesen\n16 = Q Pressewesen\n17 = R Reklame, Plakate, Gebrauchsgraphik\n18 = S Kultur- und Weltgeschichte, Historische Hilfswissenschaften\n19 = T Kunst und Kunstgewerbe\n20 = andere Sammlung","1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20","1")
                                if (!erg) break Erfolg;
		        if (erg == 1) 
                                    {
                                    F7100 = "A ";
                                    searchTerm = "dbsm-fb00-klem.a000.####";
                                    }
		        if (erg == 2) 
                                    {
                                    F7100 = "B ";
                                    searchTerm = "dbsm-fb00-klem.b000.####";
                                    }
		        if (erg == 3)
                                    {
                                    F7100 = "C ";
                                    searchTerm = "dbsm-fb00-klem.c000.####";
                                    }
		        if (erg == 4)
                                    {
                                    F7100 = "D ";
                                    searchTerm = "dbsm-fb00-klem.d000.####";
                                    }
		        if (erg == 5)
                                    {
                                    F7100 = "E ";
                                    searchTerm = "dbsm-fb00-klem.e000.####";
                                    }
		        if (erg == 6)
                                    {
                                    F7100 = "F ";
                                    searchTerm = "dbsm-fb00-klem.f000.####";
                                    }
		        if (erg == 7)
                                    {
                                    F7100 = "G ";
                                    searchTerm = "dbsm-fb00-klem.g000.####";
                                    }
		        if (erg == 8)
                                    {
                                    F7100 = "H ";
                                    searchTerm = "dbsm-fb00-klem.h000.####";
                                    }
		        if (erg == 9)
                                    {
                                    F7100 = "J ";
                                    searchTerm = "dbsm-fb00-klem.j000.####";
                                    }
		        if (erg == 10)
                                    {
                                    F7100 = "K ";
                                    searchTerm = "dbsm-fb00-klem.k000.####";
                                    }
		        if (erg == 11)
                                    {
                                    F7100 = "L ";
                                    searchTerm = "dbsm-fb00-klem.l000.####";
                                    }
		        if (erg == 12)
                                    {
                                    F7100 = "M ";
                                    searchTerm = "dbsm-fb00-klem.m000.####";
                                    }
		        if (erg == 13)
                                    {
                                    F7100 = "N ";
                                    searchTerm = "dbsm-fb00-klem.n000.####";
                                    }
		        if (erg == 14)
                                    {
                                    F7100 = "O ";
                                    searchTerm = "dbsm-fb00-klem.o000.####";
                                    }
		        if (erg == 15)
                                    {
                                    F7100 = "P ";
                                    searchTerm = "dbsm-fb00-klem.p000.####";
                                    }
		        if (erg == 16)
                                    {
                                    F7100 = "Q ";
                                    searchTerm = "dbsm-fb00-klem.q000.####";
                                    }
		        if (erg == 17)
                                    {
                                    F7100 = "R ";
                                    searchTerm = "dbsm-fb00-klem.r000.####";
                                    }
		        if (erg == 18)
                                    {
                                    F7100 = "S ";
                                    searchTerm = "dbsm-fb00-klem.s000.####";
                                    }
		        if (erg == 19)
                                    {
                                    F7100 = "T ";
                                    searchTerm = "dbsm-fb00-klem.t000.####";
                                    }
		        if (erg == 20)
                                    {
                                    F7100 = " ";
                                    F6710 = " ";
                                    }
                                else
                                    {
                                    search = "f syf "
                                    }
                           vTag = "6710"
                           F6710 = "!"+DBSMLink(search, searchTerm, vTag)+"!"
                           }
                           if (Slg == "M") 
                                {
                                var erg = __Pruef("Eine Frage","Welche Signaturgruppe?\n 1 = XI A Studiensammlungen A-Formate \n 2 = XI B Studiensammlungen B-Formate\n 3 = Stiftung Buchkunst\n4 = andere Sammlung","1,2,3,4","1")
                                if (!erg) break Erfolg;
		        if (erg == 1 || erg == 2)
                                    {
                                    F2105 = "04,P01-s-51"
                                    F8598 = "[Z-Klemm]";
		            if (erg == 1) 
                                        {
                                        F6710 = "!1080775625!";
                                        F7100 = "XI A ";
                                        }
                                        else
                                        {
                                        F6710 = "!108077565X!";
                                        F7100 = "XI B ";
                                        }
                                    }
            	        if (erg == 3) 
                                    {
                                    eArt = "ge";
                                    F2105 = "04,P01-s-63"
                                    F6710 = "";
                                    F7100 = "StB ";
                                    F8598 = "[Z-StB]";
                                    }
            	        if (erg == 4) 
                                    {
                                    var erg = __Pruef("Eine Frage","Welche Signaturgruppe?\n 1 = Klemm I: Handschriften\n 2 = Klemm II: Inkunabeln\n 3 = Klemm III: Drucke von 1501 bis 1560\n4 = Klemm IV: Drucke von 1561 bis 1800\n5 = Klemm V: Drucke von 1801 bis 2003\n6 = Klemm VI: Faksimiles\n7 = Klemm VII: Künstlerbücher, originalgraphische Mappenwerke\n8 = Klemm VIII: Einbände\n9 = Klemm IX: Kalender\n10 = Klemm X: digitae Medien\n11 = andere Sammlung: ","1,2,3,4,5,6,7,8,9,10,11","1")
                                    if (!erg) break Erfolg;
                                    if (erg != 11) F8598 = "[Z-Klemm]";
		            if (erg == 1)
                                        {
                                        F2105 = "04,P01-s-21"
                                        F6710 = "!103242012X!";
                                        F7100 = "I ";
                                        }
		            if (erg == 2)
                                        {
                                        F2105 = "04,P01-s-22"
                                        F6710 = "!1032420219!";
                                        F7100 = "II ";
                                        }
		            if (erg == 3)
                                        {
                                        F2105 = "04,P01-s-23"
                                        F6710 = "!1032420235!";
                                        F7100 = "III ";
                                        }
		            if (erg == 4)
                                        {
                                        F2105 = "04,P01-s-24"
                                        F6710 = "!1032420278!";
                                        F7100 = "IV ";
                                        }
		            if (erg == 5)
                                        {
                                        F2105 = "04,P01-s-25"
                                        F6710 = "!1032420286!";
                                        F7100 = "V ";
                                        }
		            if (erg == 6)
                                        {
                                        F2105 = "04,P01-s-31"
                                        F6710 = "!1032420294!";
                                        F7100 = "VI ";
                                        }
		            if (erg == 7)
                                        {
                                        F2105 = "04,P01-s-32"
                                        F6710 = "!1032420375!";
                                        F7100 = "VII ";
                                        }
		            if (erg == 8)
                                        {
                                        F2105 = "04,P01-s-33"
                                        F6710 = "!1032420448!";
                                        F7100 = "VIII ";
                                        }
		            if (erg == 9)
                                        {
                                        F2105 = "04,P01-s-34"
                                        F6710 = "!1032420480!";
                                        F7100 = "IX ";
                                        }
		            if (erg == 10)
                                        {
                                        F2105 = "04,P01-s-35"
                                        F6710 = "!1032420529!";
                                        F7100 = "X ";
                                        }
		            if (erg == 11)
                                        {
                                        F8598 = "[Z-???]";
                                        F2105 = "04,P01-s-??"
                                        F6710 = " ";
                                        F7100 = " ";
                                        }
                                    }
                                }
                            if (Slg == "S") 
                                {
                                if (OGattung != "Wasserzeichen-Beleg") var erg = __Pruef("Eine Frage","Welche Signaturgruppe?\n 1 = laufende Erwerbungen \n 2 = geschlossene Sammungen \n 3 = andere Sammlung","1,2,3","1")
                                else erg = 4
                                if (!erg) break Erfolg;
		        if (erg == 1) 
                                    {
 		            if (OGattung == "Bild")  
                                        {
                                        F4105 = "!1060610671!"
                                        F7100 = jahr + "/Bl/"
                                        F8598 = "[Z-GS]"
                                        var erg = __Pruef("Eine Frage","Welches Format?\n 1 = A-Format\n 2 = B-Format\n 3 = C-Format\n4 = D-Format\n 5 = E-Format\n 6 = Sonstiges","1,2,3,4,5,6","1")
                                        if (!erg) break Erfolg;
	                            if (erg == 1) F6710 = "!1032381809!";
                                        if (erg == 2) F6710 = "!1032381825!";
	                            if (erg == 3) F6710 = "!103238185X!";
	                            if (erg == 4) F6710 = "!1032381876!";
	                            if (erg == 5) F6710 = "!1032381906!";
	                            if (erg == 6) F6710 = "! !";
                                        }
                                    if (SatzArt.indexOf("D") == 0 || SatzArt.indexOf("H") == 0 || SatzArt.indexOf("L") == 0)
                                        {
                                        F4105 = "!1059146037!"
                                        F6710 = " "
                                        F7100 = jahr + "/Arch/"
                                        F8598 = "[Z-Arch]"
                                        }
		            if (OGattung == "Buntpapier")  
                                        {
                                        F4105 = "!106337698X!"
                                        F8598 = "[Z-BS]"
                                        var erg = __Pruef("Eine Frage","Welches Format?\n 1 = A-Format\n 2 = B-Format\n 3 = Sonstiges","1,2,3","2")
                                        if (!erg) break Erfolg;
	                            if (erg == 1) 
                                           {
                                           F6710 = "!1077045972!";
                                           F7100 = "BS EB " + jahr + " A"
                                           }
	                             if (erg == 2) 
                                            {
                                            F6710 = "!1077046030!";
                                            F7100 = "BS EB " + jahr + " B"
                                            }
	                              if (erg == 3) F6710 = " ";
                                          }
                                    if (SatzArt.indexOf("X") == 0)
                                          {
                                          F4105 = "!1032100486!"
                                          F6710 = "!108050477X!";
                                          F7100 = "KHS EO " + jahr + "/"
                                          F8598 = "[Z-KHS]"
                                          var erg = __Pruef("Eine Frage","Welches Format?\n 1 = RT 35\n 2 = RT 50\n 3 = Sonstiges","1,2,3","1")
                                          if (!erg) break Erfolg;
	                              if (erg == 1) F7109 = "UG2, RT 35"
	                              if (erg == 2) F7109 = "UG2, RT 50"
	                              if (erg == 3) F7109 = ""
                                          }
                                    var erg = 0
                                    }
		        if (erg == 2)
                                   {
                                   if (OGattung == "Buntpapier") 
                                       {
                                       F4105 = "!106337698X!"
                                       F6710 = " ";
                                       F7100 = "BS GS/"
                                       F8598 = "[Z-BS]"
                                       }
                                   if (OGattung == "Bild") 
                                       {
                                       F4105 = "!1060610671!"                                   
                                       F6710 = " ";
                                       F7100 = "GS/"
                                       F8598 = "[Z-GS]"
                                       }
                                    if (SatzArt.indexOf("D") == 0 || SatzArt.indexOf("H") == 0 || SatzArt.indexOf("L") == 0)
                                        {
                                        F4105 = "!1059146037!"
                                        F6710 = " "
                                        F7100 = jahr + "/GS "
                                        F8598 = "[Z-Arch]"
                                        }
                                    if (SatzArt.indexOf("X") == 0)
                                        {
                                        F4105 = "!1032100486!"
                                        F6710 = " ";
                                        F7100 = "KHS GS "
                                        F8598 = "[Z-KHS]"
                                        }
                                    }
                                if (erg == 3) F6710 = "! !";
                                if (erg == 4)
                                    {
                                     F4105 = "!1048061809!"
                                     search = "f syw "
                                     if (!6710) searchTerm = F4046 + " or syw " + VPer.substring(VPer.lastIndexOf(" ")+1,VPer.length)
                                     else  searchTerm = __Prompter("Eine Frage","Geben Sie ein Suchwort für die Bestandsgruppe (Feld 6710) ein!",F6710,"")
                                     vTag = "6710"
                                     if (searchTerm && searchTerm != null && searchTerm.indexOf("!") != 0) F6710 = "!"+DBSMLink(search, searchTerm, vTag)+"!"
                                     F7100 = "WZ II, ???/?/?/?"
                                     F8100 = "WZ-II-???-0000???"
                                     F8598 = "[Z-WZ]"
                                    } 
                                }
                            }
		//Erwerbungsart festlegen
                        if(!eArt && SatzArt.indexOf("l") != 1)
                            {
                            var erg = __Pruef("Eine Frage","Welche Erwerbungsart?\n 1 = Kauf \n 2 = Tausch \n 3 = Geschenk\n 4 = Stiftung\n 5 = Depositum\n 6 = Leihgabe\n 7 = Fund\n 8 = Altbestand\n 9 = Pflicht","1,2,3,4,5,6,7,8,9","1")
                            if (!erg) break Erfolg;
		    if (erg == 1) eArt = "ka";
		    if (erg == 2) eArt = "ta";
		    if (erg == 3) eArt = "ge";
		    if (erg == 4) eArt = "st";
		    if (erg == 5) eArt = "de";
		    if (erg == 6) eArt = "lg";
		    if (erg == 7) eArt = "fu";
		    if (erg == 8) eArt = "ab";
		    if (erg == 9) eArt = "pf";
                            }
	break Erfolg;
	}
while (status == "NOHITS")

//application.messageBox("","eArt: " + eArt + "\nSatzArt: " + SatzArt + "\nITyp: " + ITyp + "\nMTyp: " + MTyp + "\nDTyp: " + DTyp + "\nOGattung: " + OGattung + "\nFormA1: " + FormA1 + "\nFormA2: " + FormA2 + "\nFormF1: " + FormF1 + "\nFormF2: " + FormF2  + "\nFormE1: " + FormE1 + "\nFormE2: " + FormE2  + "\nFormV: " + FormV + "\nFormO: " + FormO  + "\nDTMaterial: " + DTMaterial + "\nArtInhalt: " + ArtInhalt + "\nSlg: " + Slg,"")


	application.activeWindow.command("e", false);
	// wenn Editiermodus nicht möglich, Meldung und Ende
	if (!application.activeWindow.title) 
		{
		application.messageBox("Fehler!","Die Maske konnte nicht aufgerufen werden!\nEventuell sind Sie nicht eingelogt\noder haben Sie keine Bearbeitungsrechte?","");
		}
		
                        // Datenausgabe	
		// 0500
		application.activeWindow.title.insertText("\n0500 " + SatzArt);

		// 0501
		application.activeWindow.title.insertText("\n0501 ");
		if(ITyp) application.activeWindow.title.insertText(ITyp);

		// 0502
		application.activeWindow.title.insertText("\n0502 ");
		if(MTyp) application.activeWindow.title.insertText(MTyp);

		// 0503
		application.activeWindow.title.insertText("\n0503 ");
            	if(DTyp) application.activeWindow.title.insertText(DTyp);

		// 0600
		application.activeWindow.title.insertText("\n0600 yy");
                        if (SatzArt == "Alxo") application.activeWindow.title.insertText(";at");

		// 1100 / 1110
                        if (jhr) jahr = jhr;
                        if (OGattung == "Wasserzeichen-Beleg" || strKuerzel == "Rue" || strKuerzel != "Sto") jahr = "";
                        if (F1100) jahr = F1100;
		application.activeWindow.title.insertText("\n1100 " + jahr);
                        if (Slg == "S") application.activeWindow.title.insertText("\n1110 *$4ezth");

		// 1130
		if (strKuerzel != "Len" && DTMaterial && SatzArt != "Alxo" && Slg != "F") application.activeWindow.title.insertText("\n1130 " + DTMaterial);

		// 1131
		if (ArtInhalt) application.activeWindow.title.insertText("\n1131 " + ArtInhalt);

		// 1132
                        if (SatzArt != "Alxo" && Slg != "F") 
                            {
		    if (FormA1) Form = FormA1;
		    if (FormA2) Form = Form + ";" + FormA2;
		    if (FormF1) Form = Form + ";" + FormF1; 
		    if (FormF2) Form = Form + ";" + FormF2;
		    if (FormE1) Form = Form + ";" + FormE1; 
		    if (FormE2) Form = Form + ";" + FormE2;
		    if (FormO) Form = Form + ";" + FormO;
		    if (FormV) Form = Form + ";" + FormV;
		    if (Form)
                                {
                                if (Form.indexOf(";") == 0) Form = Form.substring(1,Form.length-1);
		        application.activeWindow.title.insertText("\n1132 " + Form);
                                }
                            }

		// Sprachcode in 1500
  		if (Status != "kurz") application.activeWindow.title.insertText("\n1500 /1" );
                        if (F1500) application.activeWindow.title.insertText(F1500.substring(2,F1500.length));
                        if (Status != "kurz" && SatzArt != "Alxo" && FormF1 != "f1-text" && ITyp != "gesprochenes Wort$bspw") application.activeWindow.title.insertText("zxx");
                            
                        // RDA-Code
                        if (F1505) application.activeWindow.title.insertText("\n1505 " + F1505);

		// Ländercode in 1700   
		 if (Status != "kurz" && SatzArt != "Alxo" && OGattung != "Wasserzeichen-Beleg") application.activeWindow.title.insertText("\n1700 /1" );

                        // ISBN in 2000
		if (SatzArt.indexOf("A") == 0 && SatzArt != "Alxo") application.activeWindow.title.insertText("\n2000 ");

		// Pseudoheftnummer in 2105
                        if (Slg != "S" && SatzArt != "Alxo") application.activeWindow.title.insertText("\n2105 ");
		if (Slg != "S" && F2105 != "") application.activeWindow.title.insertText(F2105);

		// bg. Nachweis in 2035
		if (Status != "kurz" && Slg != "F" && SatzArt != "Alxo" && OGattung != "Wasserzeichen-Beleg") application.activeWindow.title.insertText("\n2035 [...]" );

                        // geistiger Schöpfer in 3000
                        if (Status != "kurz") 
                            {
                            if (OGattung != "Wasserzeichen-Beleg")
                                { 
		        application.activeWindow.title.insertText("\n3000 ");
                                if (SatzArt.indexOf("A") == 0 || SatzArt.indexOf("D") == 0) application.activeWindow.title.insertText(" $BVerfasser$4aut");
                                if (OGattung == "Bild" || OGattung == "Buntpapier") application.activeWindow.title.insertText(" $BKünstler$4art");
		        if (SatzArt.indexOf("X") == 0) application.activeWindow.title.insertText(" $Bgeistiger Schöpfer$4cre");
//                                if (SatzArt.indexOf("?") == 0) application.activeWindow.title.insertText(" $??$4??");
                                }
                            else
                                {
                                application.activeWindow.title.insertText("\n3010 "+ F3000);
                                if (F3000.indexOf("$4ppm") == -1) application.activeWindow.title.insertText("$BPapiermacher$4ppm");
                                application.activeWindow.title.insertText("\n3100 "+  F3100);
                                if (F3100.indexOf("$4cre") == -1) application.activeWindow.title.insertText("$Bgeistiger Schöpfer$4cre");
                                }
                            }

                        // Titel / Verantwortlichkeit in 4000
		if (F7100.indexOf("StB") == -1) application.activeWindow.title.insertText("\n4000  / ");
                        if (OGattung == "Wasserzeichen-Beleg") 
                            {
                            application.activeWindow.title.insertText("[Papiermühle ");
                            if (F4046) application.activeWindow.title.insertText(F4046);
                            else application.activeWindow.title.insertText("...");
                            application.activeWindow.title.insertText("; Papiermacher " + VPer + "]");
                            }
                        // Objektbezeichnung in 4019
		if (Status != "kurz" && SatzArt.indexOf("l") != 1 && Slg != "F") 
                                {
                                application.activeWindow.title.insertText("\n4019 ");
                                if (F4019) application.activeWindow.title.insertText(F4019);
                                application.activeWindow.title.insertText("$Bobja");
                                }
                                
                        // unselbständige 
                        if (SatzArt.indexOf("l") == 1)
                            {
                            application.activeWindow.title.insertText("\n0551 1$bi");                            
//                            application.activeWindow.title.insertText("\n1505 $erda");
                            if (!F4070) {ask = "Welche Zeitschrift?\n"; vorgabe = "1"}
                            else {ask = "Welche Zeitschrift?\n 0 = weiter mit aktueller ZS \n"; vorgabe = "0"}
                            if(strKuerzel == "Rue" && SatzArt.indexOf("A") == 0)
                                {
                                var erg = __Pruef("Eine Frage",ask + "1 = Boekenwereld\n 2 = Bookcollector\n 3 = Jaarboek\n 4 = andere","0,1,2,3,4",vorgabe)
                                if (!erg) return;;
		        if (erg == 1) F4241 = "Enthalten in!010846476!";
		        if (erg == 2) F4241 = "Enthalten in!010722246!";
		        if (erg == 3) F4241 = "Enthalten in!018404243!";
		        if (erg == 4) F4241 = " ";
		        }
                            if(strKuerzel == "Sta" && SatzArt.indexOf("A") == 0)
                                {
                                var erg = __Pruef("Eine Frage",ask + "1 = Börsenblatt\n 2 = Anzeiger  \n 3 = Handbuch / Verband Deutscher Antiquare\n 4 = Schweizer Buchhandel\n5 = The book collector \n6 = andere","0,1,2,3,4,5,6",vorgabe)
                                if (!erg) return;
		        if (erg == 1) F4241 = "Enthalten in!024179426!";
		        if (erg == 2) F4241 = "Enthalten in!015075915!";
		        if (erg == 3) F4241 = "Enthalten in!027219720!";
		        if (erg == 4) F4241 = "Enthalten in!012621358!";
		        if (erg == 5) F4241 = "Enthalten in!010722246!";
		        if (erg == 6) F4241 = "";
		        }
                            if(strKuerzel == "WH" && SatzArt.indexOf("A") == 0)
                                {
                                var erg = __Pruef("Eine Frage",ask + "1 = Deutscher Drucker\n 2 = Druckspiegel\n 3 = Bindereport\n 4 = Kultur & Technik \n5 = andere","0,1,2,3,4,5",vorgabe)
                                if (!erg) return;
		        if (erg == 1) F4241 = "Enthalten in!012630349!";
		        if (erg == 2) F4241 = "Enthalten in!011448857!";
		        if (erg == 3) F4241 = "Enthalten in!010084061!";
		        if (erg == 4) F4241 = "Enthalten in!010956107!";

		        if (erg == 5) F4241 = "";
		        }
                            if(strKuerzel == "Sdt" && SatzArt.indexOf("A") == 0)
                                {
                                var erg = __Pruef("Eine Frage",ask + "1 = Wochenblatt für Papierfabrikation\n 2 = Paper History\n 3 = sph-Kontakte\n 4 = The Quarterly\n 5 = Restauro\n 6 = andere","0,1,2,3,4,5,6",vorgabe)
                                if (!erg) return
		        if (erg == 1) F4241 = "Enthalten in!012686905!";
		        if (erg == 2) F4241 = "Enthalten in!020478526!";
		        if (erg == 3) F4241 = "Enthalten in!010730931!";
		        if (erg == 4) F4241 = "Enthalten in!017630932!";
		        if (erg == 5) F4241 = "Enthalten in!011768479!";
		        if (erg == 6) F4241 = "";
		        }
                            if(strKuerzel == "MaF" && SatzArt.indexOf("A") == 0)
                                {
                                var erg = __Pruef("Eine Frage",ask + "1 = Ars Scribendi \n 2 = Art & Métíers du Lívre \n 3 = Einbandforschung \n 4 = Exlibriskunst und Graphik \n 5 = Kultur & Technik \n 6 = Leipziger Blätter  \n 7 = Librarium  \n 8 = Marginalien  \n 9 = Mitteldeutsches Jahrbuch für Kultur und Geschichte  \n10 = Page \n11 = Rundbrief / Meister der Einbandkunst \n 12 = Weltkunst \n13 = andere","0,1,2,3,4,5,6,7,8,9,10,11,12,13",vorgabe)
                                if (!erg) return
		        if (erg == 1) F4241 = "Enthalten in!018102700!";
		        if (erg == 2) F4241 = "Enthalten in!015697932!";
		        if (erg == 3) F4241 = "Enthalten in!018986617!";
		        if (erg == 4) F4241 = "Enthalten in!1034634127!";
		        if (erg == 5) F4241 = "Enthalten in!010956107!";
		        if (erg == 6) F4241 = "Enthalten in!010603581!";
		        if (erg == 7) F4241 = "Enthalten in!010050132!";
		        if (erg == 8) F4241 = "Enthalten in!010690395!";
		        if (erg == 9) F4241 = "Enthalten in!017434629!";
		        if (erg == 10) F4241 = "Enthalten in!011792868!";
		        if (erg == 11) F4241 = "Enthalten in!012657921!";
		        if (erg == 12) F4241 = "Enthalten in!012657921!";
		        if (erg == 13) F4241 = "";
		        }
                            if (erg == 0) 
                                {
                                test = __Prompter("Eine Frage","Geben Sie die Seitenzahlen ein!")
                                application.activeWindow.title.insertText("\n4070 " + F4070 + "/p" + test);
                                }
                            else application.activeWindow.title.insertText("\n4070 /v/a/b/p");                                
                            if (!F4241 || F4241 == "")
                                {
                                application.activeWindow.title.insertText("\n4241 Enthalten in" );
                                }
                            else
                                {
                                application.activeWindow.title.insertText("\n4241 " + F4241 );
                                }
                            }
                            
                        // Verlagsangabe in 4030
                        if (Slg != "S" && SatzArt.indexOf("l") != 1) application.activeWindow.title.insertText("\n4030 ");
                        if (F4030) application.activeWindow.title.insertText(F4030);

                        // Enstehungsangabe in 4046
                        if (Slg == "S" && SatzArt.indexOf("l") != 1) application.activeWindow.title.insertText("\n4046 ");
                        if (F4046) application.activeWindow.title.insertText("["+F4046+"]");

                        // Umfang / Ill. / Maße / Beigaben
                        if (Status != "kurz" && SatzArt.indexOf("l") != 1) application.activeWindow.title.insertText("\n4060 ");
                        if (OGattung == "Buntpapier") application.activeWindow.title.insertText("1 Blatt");
                        if (F4060) application.activeWindow.title.insertText(F4060);
                        if (Status != "kurz" && SatzArt.indexOf("X") != 0 && SatzArt.indexOf("P") != 0) application.activeWindow.title.insertText("\n4061 Illustrationen");
                        if (OGattung == "Wasserzeichen-Beleg") 
                            {
                            application.activeWindow.title.insertText("\n4062 WZ-Motiv ?: ... mm, ... mm$b...$h...$4mwza");
                            if (DTMaterial.lastIndexOf("r") == 11) application.activeWindow.title.insertText("\n4062 WZ-Motiv ?: Abstand zwischen ? Kettlinien ... mm$b...$4mwza");
                            }
                        else if (Status != "kurz" && Slg != "F" && SatzArt.indexOf("l") != 1) 
                            {
                            if (FormF2 != "f2-3d") application.activeWindow.title.insertText("\n4062 x cm$b $h $4");
                            if (FormF2 == "f2-3d") application.activeWindow.title.insertText("\n4062 x x cm$b $h $t $4");
                            }
                        if (Status != "kurz" && Slg == "F") application.activeWindow.title.insertText("\n4062 ");
     
                        // Sammlung in 4105
                        if (F4105 && Slg != "F" && SatzArt.indexOf("l") != 1) application.activeWindow.title.insertText("\n4105 " + F4105);
                            
                        // Anmerkungen 4201
                        if (F4201) application.activeWindow.title.insertText("\n4201 " + F4201);

                        // Inhaltserschließung
                        F510X
                        if (F510X) application.activeWindow.title.insertText("\n5100 " + F510X);

                        // Systematik
                        if (F5320) application.activeWindow.title.insertText("\n5320 " + F5320);
                            
                        // Gestaltungsmerkmale
                        if (Status != "kurz" && Slg != "F" && SatzArt.indexOf("l") != 1) 
                            {
                            application.activeWindow.title.insertText("\n5590 [Objektgattung]\n5590 ");
                            if (F5590) application.activeWindow.title.insertText(F5590);
                            application.activeWindow.title.insertText("\n5591 [Entstehungsort]\n5591 ");
                            if (F5591) application.activeWindow.title.insertText(F5591);
                            if (F5592) application.activeWindow.title.insertText("\n5592 [Technik]\n5592 ");
                            if (F5592) application.activeWindow.title.insertText(F5592);
                            }

		// Bearbeiter-Name in 4700
		application.activeWindow.title.insertText("\n4700 |" + strAbteilung + "|" + strKuerzel + "; ");
		if (eArt == "ta") 
                                    {
			application.activeWindow.title.insertText("DNB-Tausch")
			}

                        // Exemplarebene
                        if(SatzArt.indexOf("l") == 1) return

                        if(SatzArt.indexOf("A") == -1) application.activeWindow.title.insertText("\n7001 x"); 
                        if(SatzArt.indexOf("A") == 0) application.activeWindow.title.insertText("\n7002 x"); 
//                        if(F7100.indexOf("StB") == -1) application.activeWindow.title.insertText("\n7001 x"); 
//                        if(F7100.indexOf("StB") == 0) application.activeWindow.title.insertText("\n7002 x"); 
                            
                        // Erwerbungsdaten in 4821 
		if(!F4821) 
                            {
                            if (strKuerzel != "Lo") application.activeWindow.title.insertText("\n4821 $l$q$w??? EUR$zErwerbung$D" + jahr +"-XX-XX$K");
                            if (strKuerzel == "Len") application.activeWindow.title.insertText("Stiftung Buchkunst");
                            }
                        else
                            {
                            application.activeWindow.title.insertText("\n4821 " + F4821);
                            }
                            
		// Sammlung in 6710
		if (F6710) application.activeWindow.title.insertText("\n6710 " + F6710 );
                        if (F6710 && strKuerzel != "Lo") application.activeWindow.title.insertText("$l")
                        if (strKuerzel == "Sto") application.activeWindow.title.insertText(F7100.substring(F7100.indexOf("[")+1,F7100.indexOf("]")));
                        if (F7100.indexOf("StB") == 0) application.activeWindow.title.insertText(F7100.substring(F7100.indexOf(",")+2,F7100.length));

                        // Verwendungsort in 6800
                        if (strKuerzel == "Lo") application.activeWindow.title.insertText("\n6800 [Verwendungsort]\n6800 XXX");
                        if (F6800) application.activeWindow.title.insertText("\n6800 [XXX]\n6800 " + F6800);

                        // Signatur in 7100 / 7109
		if (F7100) 
                            {
                            application.activeWindow.title.insertText("\n7100 " + F7100);
                            if (SatzArt.indexOf("A") == 0) application.activeWindow.title.insertText(" @ k");
                            if (Slg == "S" && F7100.indexOf("@ m") == -1) application.activeWindow.title.insertText(" @ m");

                            if (!F7109) 
                                {
                                F7109 = F7100.substring(0,F7100.indexOf("@")-1);
                                if (Status != "kurz" && Slg == "F") F7109 = "(ESK )";
                                if (Slg == "F") Slg = "F/Klemm";
                                if (Slg == "M") Slg = "M/Klemm";
                                application.activeWindow.title.insertText("\n7109 !!DBSM/" + Slg +"!! ; " + F7109);
                                }
                            else
                                {
                                application.activeWindow.title.insertText("\n7109 " + F7109);
                                }
                            }
		// Erwerbungsart in 8510
		if (eArt) application.activeWindow.title.insertText("\n8510 %" + eArt);
                        if (F8510) application.activeWindow.title.insertText("\n8510 "+F8510)

		// AKZ in 8100
		if (F8100 ) application.activeWindow.title.insertText("\n8100 " + F8100);
                        if(F7100.indexOf("StB") == 0) application.activeWindow.title.insertText("\n8100 ");
		// Zugangsnummer in 8598
		if (F8598 && SatzArt != "Qd") 
                            {
                            if(F7100.indexOf("StB") == -1) application.activeWindow.title.insertText("\n8598 " + F8598 + " " + jahr + "/00000");
                            if(F7100.indexOf("StB") == 0) application.activeWindow.title.insertText("\n8598 " + F8598);
                            }

		// zum Startpunkt der Erfassung navigieren
		application.activeWindow.title.findTag("0500", 0, false, true, false);
		application.activeWindow.title.endOfField(false);	
	           
    }

}

function DBSMdoklink() {
 //      var oShell = new ActiveXObject("Shell.Application"); 
 //      var commandtoRun = "C:\\WINDOWS\\explorer.exe";  
 //      oShell.ShellExecute(commandtoRun,"B:\\Bestand\\","dnb","open","1"); 
    if (!application.activeWindow.Variable("scr") || application.activeWindow.Variable("scr") == "FI")
		{
		if (application.activeWindow.Variable("scr") == "FI") application.messageBox("Diese Funktion steht momentan nicht zur Verfügung!","Um die Funktion nutzen zu können\nmus ein Datensatz ausgewählt werden.","");
		else application.messageBox("Diese Funktion steht momentan nicht zur Verfügung!","Um die Funktion nutzen zu können\nmüssen Sie sich erst einloggen\nund einen Datensatz auswählen.","");
		return
		}
	application.activeWindow.command("k", false);
	var siga = new Array();						//Array für Exemplarliste definieren
	var x = 0								//Array-Zähler auf 0 setzen
	var n = 0                                                                                           //Schleifen-Zähler auf 0 setzen
	var slink = ""
	// alle vorhandenen Exemplarsätze ermitteln und einzeln bearbeiten
	test = application.activeWindow.title.findTag("7100",n,false,true,true);	//Feld 7100 auf test merken
	Erfolg1: do
		{
		// überprüfen ob der Exemplarsatz eine 7109 mit !!DBSM... hat
		application.activeWindow.title.LineDown(1,false);						//eine Zeile nach unten
		application.activeWindow.title.StartOfField(false);						//am den Anfang der Zeile
		application.activeWindow.title.endOfField(true);						//markiere die Zeile
		var test1 = application.activeWindow.title.Selection;					//merke Zeile auf test1
		z = 1
		//suche die zugehörige 6710
		Erfolg2: do
			{
//		            var erg = __Pruef("Prüfung","Beginn innere Schleife\ntest1 =" +  test1 + "\nsiga[0] =" +  siga[0] + "\nsiga[1] =" +  siga[1] ,"1","1")
//		            if (!erg) return;
			if (test1.substring(7,11) == "DBSM")										//wenn es 7109 !!DBSM ... ist 
				{
				slink = application.activeWindow.title.findTag("6710",x,false,true,false);                                        //suche die zugehörige 6710
				if (slink && slink != null)                                                                                       //wenn es die gibt, starte die eigentliche Bearbeitung
					{
					slink = slink.substring(12,slink.indexOf(": "));
					// 6710 $8 "Ausstellung" und "LS" übergehen
				m = x
					while (slink.indexOf(".Ausstellung") > 0 || slink.indexOf(".LS") > 0)
						{
						m = m + 1
						slink = application.activeWindow.title.findTag("6710",m,false,true,false);                                        //suche die zugehörige 6710
						slink = slink.substring(12,slink.indexOf(": "));
 						}
					// aus 6710 $8 den richtigen Präfix ermitteln
					if (slink.indexOf("Boe-Archiv") > 0) {slink = "dnb-dbsm-"}
					else if(slink.indexOf("Archiv/Boe-GR") > 0) {slink = "dnb-dbsm-"};
					else if(slink.indexOf("Archiv/Boe-Med") > 0) {slink = "dnb-dbsm-"};
					else if(slink.indexOf("StSlg Archiv") > 0) {slink = "dnb-dbsm-archiv-"};
					else if (slink.indexOf("StSlg.PS BS") > 0) {slink = "dnb-dbsm-ps-"}
					else if (slink.indexOf("FB Klemm") > 0) {slink = "dnb-dbsm-f-"}
					else if (slink.indexOf("StSlg.Buch Klemm") > 0) {slink = "dnb-dbsm-klemm-"}
					else if (slink.indexOf("KHS") > 0) {slink = "dnb-dbsm-"}
					else if (slink.indexOf("FB Boe") > 0) {slink = "dnb-dbsm-"}
					else if (slink.indexOf("StSlg GS") > 0) {slink = "dnb-dbsm-"}
					else if (slink.indexOf("GS 1.BB") > 0) {slink = "dnb-dbsm-boe-bl-"}
					else if (slink.indexOf("Buch Boe M") > 0) {slink = "dnb-dbsm-"}
					else if (slink.indexOf("BoeHA") > 0) {slink = "dnb-dbsm-"}
					else 
						{
						application.messageBox("","Für [" + slink + "] ist kein gültiger Pfadname hinterlegt.","");
						application.activeWindow.simulateIBWKey("FE");
						return;
						}
 					// aus 7100 den richtigen Signatur-Dateinamen bilden
					ssig  = test;
					if (ssig.indexOf("@") > 0) ssig = ssig.substring(0,ssig.indexOf("@")-1);
					ssig = ssig.replace('/4°', '');															// Format am Ende abschneiden
					ssig = ssig.replace('/2°', '');															// Format am Ende abschneiden
					ssig = ssig.replace(/\,/g, "-");														//Komma durch einen Strich ersetzen
					ssig = ssig.replace(/\ /g, "-");														//Leerzeichen durch einen Strich ersetzen
					ssig = ssig.replace(/\//g, "-");														//Schrägstrich durch einen Strich ersetzen
					ssig = ssig.replace(/\./g, "-");														//Punkt durch einen Strich ersetzen
					ssig = ssig.replace(/\;/g, "-");														//Semikolon durch einen Strich ersetzen
					ssig = ssig.replace(/\:/g, "-");														//Doppelpunkt durch einen Strich ersetzen
					ssig = ssig.replace(/\-\-/g, "-");														//zwei Striche durch einen ersetzen
					ssig = ssig.replace(/\-\-/g, "-");														//zwei Striche durch einen ersetzen
					ssig = ssig.toLowerCase();																//alles klein
					ssig = ssig.replace(/ö/g, 'oe');
					ssig = ssig.replace(/ä/g, 'ae');
					ssig = ssig.replace(/ü/g, 'ue');
					ssig = ssig.replace(/\[/g, "");														//eckige Klammern vernichten
					ssig = ssig.replace(/\]/g, "");														//eckige Klammern vernichten
					slink = slink + ssig;
					test2 = application.activeWindow.title.find(slink,true,false,false);
					siga[x] = new Array(2)
					siga[x][0] = slink
					if (test2 != false) siga[x][1] = "ok";
					else siga[x][1] = "neu";
					break Erfolg2;
					}
				else
					{
					siga[x] = new Array(2);
					siga[x][0] = slink + " Der Exemplarsatz muss erst in 6710 mit einem Bestand verknüpft werden!";
					siga[x][1] = "Fehler";
					break Erfolg2;
					}
				// zur nächsten Zeile und schauen, ob diese eine 7109 mit DBSM ist
				application.activeWindow.title.LineDown(1,false);                             	//eine Zeile nach unten
				application.activeWindow.title.StartOfField(false);                            	//am den Anfang der Zeile
				application.activeWindow.title.endOfField(true);                               	//markiere die Zeile
				var test1 = application.activeWindow.title.Selection;                        	//merke Zeile auf test1
				z = z + 1																		//innerer Zähler um 1 erhöhen
				if (z > 5) break Erfolg2														//wenn Zähler größer als 5 (fünf Zeilen nach unten gesucht) Zwangsausstieg aus der Schleife
				}
                                    else
                                        {
		                break Erfolg2;
                                        }
			} while (test1.substring(0,4) != "8100")										//wenn Feld 8100 erreicht, Ausstieg aus der Schleife 
		if (siga[x]) x = x + 1
		n = n + 1
		test = application.activeWindow.title.findTag("7100",n,false,true,true);	//Feld 7100 auf test merken
		}  while (test)
		//wenn mehr als ein DBSM-Exemplar gefunden wurde, stapel diese in ein Array, zeige sie an und frage, für welches Exemplar die Aktion gestartet werden soll
		if (siga[1]) 
			{
			string = ""
			x = x - 1
			do 
				{
				if (siga[x][1] == "ok") string = " öffnen\n" + string
                                                else if (siga[x][1] == "Fehler") string = "\n" + string
				else string = " anlegen\n" + string
				string = (x+1).toString() + " = " + siga[x][0] + string
				x = x - 1
				} while (x >= 0)
			var erg = __Pruef("Es gibt mehr als ein DBSM-Exemplarsatz","Welche Aktion soll gestartet werden?:\n" + string,"1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20","1")
			if (!erg) return;
			slink = siga[erg-1][0];
			swas = siga[erg-1][1]
			x = erg-1
			}
		else 
			{
			slink = siga[0][0]
			swas = siga[0][1]
			x = 0
			}
                        // eigentliche Aktion
		if (swas == "ok")
			{
			slink =  "/e,b:" + "\\" + "bestand" + "\\" + slink;
			var oShell = new ActiveXObject("Shell.Application"); 
			var commandtoRun = "C:\\WINDOWS\\explorer.exe";  
			oShell.ShellExecute(commandtoRun,slink,"","open","1"); 
			application.activeWindow.simulateIBWKey("FE");
			return
			}
		else if (swas == "neu")
			{
			var oShell = new ActiveXObject("Shell.Application"); 
			var commandtoRun = "B:\\neu.cmd";  
			oShell.ShellExecute(commandtoRun,slink,"","open","1"); 
			var erg = __Pruef("Eine Frage","Soll der Pfad im Exemplarsatz " + siga[x] + " erfasst werden? (j/n)","j,n","n")
			if (!erg) return;
			if (erg == "j")
				{
				application.activeWindow.title.findTag("7100",x,false,true,false);
				application.activeWindow.title.endOfField(false);
				application.activeWindow.title.insertText("\n8598 " + "[DOC] " + slink);
				application.activeWindow.simulateIBWKey("FR");
				IDN = application.activeWindow.variable("P3GPP")
				application.activeWindow.clipboard = slink + "IDN" + IDN	// Zeile an Zwischenablage anhängen
				var oShell = new ActiveXObject("Shell.Application"); 
				var commandtoRun = "B:\\#Uebersicht-Vollstaendigkeit-Bestand.xlsm";  
				oShell.ShellExecute(commandtoRun,"","","open","0"); 
				}
			else application.activeWindow.simulateIBWKey("FE");
			}
		else if (swas == "Fehler")
			{
			application.messageBox(swas,slink,"")
			application.activeWindow.simulateIBWKey("FE");
			}
}
function DBSMdruck(line, m) {
        // übergibt in line[] gesammelte Zeilen an die Zwischenablage und startet danach Word. Dort wird die Zwischenablage eingefügt und in eine Tabelle gewandelt.
        // jede line[m] muss die folgende Form haben: "spalte1\tSpalte2\t...\tSpalteN"
        application.activeWindow.clipboard = ""    // Zwischenablage erstmal leeren
        n = m                    // n ist der Zähler für das Leeren des Arrays
        line.reverse()            // Array umkehren, weil in line[0] die Überschrift steht und hier von hinten nach vorn ausgegeben wird.
        while (m > -1)            // für alle line[m] bis line[0]
            {
//application.messageBox("","line " + m + " ist: " + line[m],"")
		application.activeWindow.clipboard = application.activeWindow.clipboard + "\n" + line[m]		// Zeile an Zwischenablage anhängen
		m = m - 1											// Zeilenzähler um 1 zurücksetzen
	}
	var oShell = new ActiveXObject("Shell.Application"); 
	var commandtoRun = "V:\\06_DBSM\\02_Erschließung\\01_allgemein\\WinIBW\\Starteinstellung\\word-start.cmd";  
	oShell.ShellExecute(commandtoRun,"","","open","1"); 
        while(n > 0)
            {          
            line.shift();
            n = n - 1
            }
}
function DBSMVerwendungsortVerschieben() {
var line = new Array()						// Array für die Zeilen
var n = __Pruef("Eine Frage","Wieviele Datensätze?","1,2,3,4,5,6,7,8,9,10,20,30,40,50,60,70,80,90,100","10")
m = 0
while(n > 0)
    {
    application.activeWindow.command("k", false);
    if (application.activeWindow.title.find("5592 [Verwendungsort]"))
        {
            strIDN = application.activeWindow.title.findTag("5592",1,false,true,true);
            application.activeWindow.title.deleteLine()
            application.activeWindow.title.findTag("5592",0,false,true,false);
            application.activeWindow.title.deleteLine()
            application.activeWindow.title.endOfBuffer()
	application.activeWindow.title.insertText("\n6800 [Verwendungsort]\n6800 " + strIDN.substring(0,strIDN.length))
            application.activeWindow.simulateIBWKey("FR");						//speichern
	if (application.activeWindow.Status != "OK")								//wenn nicht erfolgreich
	    {
	    m = m+1
	    line[m] = m + "\t" +  application.activeWindow.Variable("P3GPP") + "\t" +  application.activeWindow.messages.item(0)
	    application.activeWindow.simulateIBWKey("FE");								//Bearbeiten abbrechen
	    }
        }
    else application.activeWindow.simulateIBWKey("FE");
    n = n - 1
    application.activeWindow.simulateIBWKey("F1");
    }

DBSMdruck(line, m)
}
function StatCount() {
// Vorgabewerte abfragen
// Welche Statistik
erg = __Pruef("Eine Frage","Welche Statistik?\n 1 = Zugangsstatistik\n 2 = Statistik FE \n 3 = Statistik ... ","1,2,3","1");
if (erg == 1) searchTerm = "statzug";
else if (erg == 2) searchTerm = "statfoe";
else if (erg == 3) searchTerm = "statnfz";
else if (!erg) return;
// Welche Nutzerkennung
erg = __Pruef("Eine Frage","Welche Nutzerkennung?\n 1 = 1130 bis 1180\n 2 = 1230 bis 1280 \n 3 = 1330 bis 1380 \n 4 = Gesamtstatistik","1,2,3,4","4");
if (erg == 1) nutzerKennung = "11[345678]#";
else if (erg == 2) nutzerKennung = "12[345678]#";
else if (erg == 3) nutzerKennung = "13[345678]#";
else if (erg == 4) nutzerKennung = "1[12345][0123456789]#";
else if (!erg) return;
// Welcher Zeitraum
dat = __Prompter("Für welchen Zeitraum?","Geben Sie ein Datum ein!","20##-##-##")	

// Auswertung
summe = 1
for (var i = 1; i <= 10; i++)	
	{
	strCommand = "f DKE "+searchTerm+"###"+dat+"###"+nutzerKennung+"###"+i+"*"		//Such-Zeichenkette auf strCommand merken 
	application.activeWindow.command(strCommand,true);								//nach Such-Zeichenkette suchen
	var gefunden = application.activeWindow.Status									//Status abfragen
	if (gefunden == "NOHITS")
		{
		// für Testzwecke nächste Zeile aktivieren
		//application.messageBox(gefunden, "Für "+strCommand+" wurde nichts gefunden. Es bleibt bei  "+summe+".","");
		}
	else
		{
		if ( application.activeWindow.Variable("scr") != "8A") application.activeWindow.simulateIBWKey("FR");	//Wenn Bildschirm nicht schon in Vollanzeige ist, in Kurz- bzw. Vollanzeige wechseln
		count = application.activeWindow.Variable("P3GSZ");														//Anzahl der Treffer auf count merken
		// für Testzwecke nächste Zeile aktivieren
		//application.messageBox("Zwischen-Auswertung", summe+" wird um "+i+" * "+count+" erhöht.","");
		summe = parseInt(summe) + parseInt((count * i));														//Anzahl der Treffer mit i multiplizieren und zu Summe addieren
		application.activeWindow.closeWindow();																	//Fenster schließen
		}
	application.activeWindow.closeWindow();																	//Fenster schließen
	}

// Ausgabe des Ergebnisses
// Texte für Ausgabe anpassen
if (searchTerm == "statzug") searchTerm = "Zugänge.";
else if (searchTerm == "statfoe") searchTerm = "von FE bearbeitete Datensätze.";
else if (searchTerm == "...") searchTerm = "...";

if (nutzerKennung == "11[345678]#") nutzerKennung = "die Nutzerkennungen 1130 bis 1180";
else if (nutzerKennung == "12[345678]#") nutzerKennung = "die Nutzerkennungen 1230 bis 1280";
else if (nutzerKennung == "13[345678]#") nutzerKennung = "die Nutzerkennungen 1330 bis 1380";
else if (nutzerKennung == "1[12345][0123456789]#") nutzerKennung = "alle Nutzerkennungen";
// Ergebnis ausgeben
application.activeWindow.clipboard = summe				// Summe in Zwischenablage kopieren
application.messageBox("Auswertung", "Für  "+nutzerKennung+"\n gab es im Zeitraum "+dat+"\n"+summe+" "+ searchTerm+"\n Das Ergebnis wurde in die Zwischenablage kopiert.","");
}
function DBSMDruckerSuchen() {
	// Place your function code here
var line = new Array()						// Array für die Zeilen
m = 0
var n = __Pruef("Eine Frage","Wieviele Datensätze?\nx = gesamte Ergebnismenge","1,2,3,4,5,6,7,8,9,10,20,30,40,50,60,70,80,90,100,x","10")
if (n == "x") n = application.activeWindow.Variable("P3GSZ");
if ( application.activeWindow.Variable("scr") != "8A") application.activeWindow.simulateIBWKey("FR");
while(n > 0)
    {
    application.activeWindow.command("k", false);

    Test =  application.activeWindow.title.findTag("4201",0,false,true);                                      // Anmerkung auf Test merken
    Test = Test.substring(Test.indexOf("Drucker"),Test.length)                                                 // in Test alles vor "Drucker" abschneiden
    Test = Test.substring(Test.indexOf(":")+2,Test.length)                                                        //  in Test alles vor ":" abschneiden
//    if (Test.indexOf(". -") > 0)                                                                                                 // wenn es in Test ". -" gibt
//        {Test = Test.substring(0,Test.indexOf(". -"))}                                                                        //  in Test alles ab ". -" abschneiden
//    else                                                                                                                               // sonst
//        {Test = Test.substring(0,Test.indexOf(" "))}                                                                          //  in Test alles ab " " abschneiden
    Test = Test.substring(0, Test.indexOf(" "))
    Test = "3110 " + Test                                                                                                      //  in Test "3110 " ergänzen
    application.activeWindow.title.startOfBuffer()                                                                     // an den Anfang des Datensatzes gehen
    Test1 = application.activeWindow.title.find(Test)                                                                // im Datensatz nach Test suchen
    if (Test1)                                                                                                                        // wenn erfolgreich
        {
        application.activeWindow.title.find("$B")                                                                               // zu $B navigieren
        application.activeWindow.title.endOfField(true)                                                                     // bis zum Ende der Zeile markieren
        application.activeWindow.title.insertText("$BDrucker$4prt")                                                   // Markierung durch "$BDrucker$4prt" ersetzen
        application.activeWindow.simulateIBWKey("FR");						//speichern
	if (application.activeWindow.Status != "OK")						//wenn nicht erfolgreich Fehler merken und Bearbeiten abbrechen
	    {
	    m = m+1
	    line[m] = m + "\t" +  application.activeWindow.Variable("P3GPP") + "\t" +  application.activeWindow.messages.item(0)
	    application.activeWindow.simulateIBWKey("FE");					
	    }
        }
    else 
        {
        application.activeWindow.simulateIBWKey("FE");
        m = m+1
        line[m] = m + "\t" +  application.activeWindow.Variable("P3GPP") + "\tDrucker in 4201 und 3110 nicht gleich"
        application.activeWindow.simulateIBWKey("FE");								//Bearbeiten abbrechen
        }
    n = n - 1
    if ( application.activeWindow.Variable("scr") != "8A") application.activeWindow.simulateIBWKey("FR");
    application.activeWindow.simulateIBWKey("F1");
    }
DBSMdruck(line, m)

}
function DBSMDatumErgaenzen() {
	// Place your function code here
		// Place your function code here
var line = new Array()						// Array für die Zeilen
m = 0
var n = __Pruef("Eine Frage","Wieviele Datensätze?\nx = gesamte Ergebnismenge","1,2,3,4,5,6,7,8,9,10,20,30,40,50,60,70,80,90,100,x","10")
if (n == "x") n = application.activeWindow.Variable("P3GSZ");
if ( application.activeWindow.Variable("scr") != "8A") application.activeWindow.simulateIBWKey("FR");
while(n > 0)
    {
    application.activeWindow.command("k", false);

    Test =  application.activeWindow.title.findTag("4070",0,false,true);                                      // Anmerkung auf Test merken
    if (Test.indexOf("/b") > 0)                                                                                                 // wenn es in Test ". -" gibt
        {
        Test = Test.substring(Test.indexOf("/b")+2,Test.length)                                                 // in Test alles vor "Drucker" abschneiden
        if (Test.indexOf("/") > 0)                                                                                                 // wenn es in Test ". -" gibt
            {Test = Test.substring(0,Test.indexOf("/"))}                                                                        //  in Test alles ab ". -" abschneiden
        Test = "yy;at\n1100 " + Test                                                                                                      //  in Test "3110 " ergänzen
        application.activeWindow.title.startOfBuffer()                                                                     // an den Anfang des Datensatzes gehen
        application.activeWindow.title.find("yy")                                                                // im Datensatz nach Test suchen
        application.activeWindow.title.insertText(Test)                                                   // Markierung durch "$BDrucker$4prt" ersetzen
        application.activeWindow.simulateIBWKey("FR");						//speichern
	if (application.activeWindow.Status != "OK")						//wenn nicht erfolgreich Fehler merken und Bearbeiten abbrechen
	    {
	    m = m+1
	    line[m] = m + "\t" +  application.activeWindow.Variable("P3GPP") + "\t" +  application.activeWindow.messages.item(0)
	    application.activeWindow.simulateIBWKey("FE");					
	    }
        }
    else
        {
        Test1 =  application.activeWindow.title.findTag("1100",0,false,true);                                      // Anmerkung auf Test merken
        if (!Test1)
            {
            Test = "yy;at\n1100 XXXX"
            application.activeWindow.title.startOfBuffer()                                                                     // an den Anfang des Datensatzes gehen
            application.activeWindow.title.find("yy")                                                                // im Datensatz nach Test suchen
            application.activeWindow.title.insertText(Test)                                                   // Markierung durch "$BDrucker$4prt" ersetzen
            application.activeWindow.simulateIBWKey("FR");						//speichern
	if (application.activeWindow.Status != "OK")						//wenn nicht erfolgreich Fehler merken und Bearbeiten abbrechen
	    {
	    m = m+1
	    line[m] = m + "\t" +  application.activeWindow.Variable("P3GPP") + "\t" +  application.activeWindow.messages.item(0)
	    application.activeWindow.simulateIBWKey("FE");					
	    }            
            }
        }
    n = n - 1
    if ( application.activeWindow.Variable("scr") != "8A") application.activeWindow.simulateIBWKey("FR");
    application.activeWindow.simulateIBWKey("F1");    
    }
DBSMdruck(line, m)


	
}
function DBSMEJergaenzen() {
	// Place your function code here
	
	
}
