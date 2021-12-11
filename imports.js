// ja
function Datenimport() {

	var input = utility.newFileInput();
	var opened = input.openViaGUI("Eingabedatei wählen", "D:\\", "testTp.txt", "*.dat", "Textdateien");

	if (!opened) {
		application.messageBox("Fehler", "Kann Input nicht lesen", "error-icon");
		input.close();
		return;
	}
	var log = utility.newFileOutput();
	opened = log.create("D:\\za-batch.log");

	if (!opened) {
		application.messageBox("Fehler", "Kann Input nicht lesen", "error-icon");
		log.close();
		return;
	}
	var record;
var recordcount = 0;

	while ((record = _readRecord(input)) != null) {
		application.activeWindow.command("\\inv 1"); //neues eingabefenster titeldaten
                        //application.activeWindow.command("\\inv 2"); //neues eingabefenster normdaten

		application.activeWindow.title.insertText(record);
		//application.messageBox("", "", "");
		application.activeWindow.simulateIBWKey("FR"); // Enter
		recordcount++;
        if (application.activeWindow.status != "OK") {

            log.write("ERROR, ");
            log.write(recordcount);
			application.activeWindow.simulateIBWKey("FE"); // Escape

        } else {
            log.write(application.activeWindow.variable("P3GPP")); //idn
        }

        if (application.activeWindow.messages.count > 0) {
			for (i = 0; i < application.activeWindow.messages.count; i++) {
				log.write(",\x22");
				log.write(application.activeWindow.messages.item(i).text);
				log.write("\x22");

        }
	}

        log.write("\n");
	}

	input.close();
	log.close();
}

function teststring() {	
	var input = utility.newFileInput();
	var opened = input.openViaGUI("Eingabedatei wählen", "D:\\", "testTp.txt", "*.dat", "Textdateien");
	if (!opened){
		application.messageBox("Fehler", "Kann Input nicht lesen", "error-icon");
		return;
	}
	var log = utility.newFileOutput();
	opened = log.create("D:\\batchlog.txt");

	if (!opened) {
		input.close();
		application.messageBox("Fehler", "Logdatei nicht angelegt", "error-icon");
		return;
	}
    var record;
	while ((record = _readRecord(input)) != null) {
                        
	var recordlines = record.split("\n");
	application.messageBox("Info", "ganz record " + record + "." + "lines" + recordlines.length , "error-icon");
	for (var i = 1; i < recordlines.length; i++) {
			if (i == 1) {
				application.messageBox("Info", "nur idn " + recordlines[i], "error-icon");
				application.activeWindow.command("f idn " + recordlines[i], false);
			} else if (i > 1) {
				var feld = recordlines[i].split("\t");
				application.messageBox("Info", "datensatzzeilen " + recordlines[i], "error-icon");
				//application.activeWindow.command("k", false); //Bearbeiten ein
				//__addTag(feld[0], feld[1], true);
			}
		}

    }
}

function BatchChange() {
	var input = utility.newFileInput();
	var opened = input.openViaGUI("Eingabedatei wählen", "D:\\", "testTp.txt", "*.dat", "Textdateien");
	if (!opened)
		return;
	var log = utility.newFileOutput();
	opened = log.create("D:\\logTp.txt");
	if (!opened) {
		input.close();
		return;
	}

	var record;
	while ((record = _readRecord(input)) != null) {
                        application.messageBox("Info", "ganz record " + record, "error-icon");
		var recordlines = record.split("\n");
		for (var i = 1; i <= (recordlines.length - 1); i++) {
			if (i == 1) {
                //application.messageBox("Info", recordlines[i], "error-icon");
				application.activeWindow.command("f idn " + recordlines[i], false);
				var idn = recordlines[i];
			} else if (i > 1) {
				application.messageBox("Info", "länge " + recordlines[i].length, "error-icon");
				var feld = recordlines[i].split("\t");
				application.messageBox("Info","feld 1 " + feld[0] + "feld 2 " + feld[1], "error-icon");
				application.activeWindow.command("k", false); //Bearbeiten ein
				__addTag(feld[0],feld[1],true);
			}
		}

		application.activeWindow.simulateIBWKey("FR"); // Enter
		//status = application.activeWindow.status;
		//idn = application.activeWindow.variable("P3GPP");
        
        if (application.activeWindow.status != "OK") {
            log.write("ERROR, \x22");
            log.write(record);
            log.write("\x22");
			application.activeWindow.simulateIBWKey("FE"); // Escape
		
        } else {
            log.write(idn);
        }
		
        if (application.activeWindow.messages.count > 0) {
            log.write(",\x22");
            log.write(application.activeWindow.messages.item(0));
            log.write("\x22");
        }
        
        log.write("\n");
	}
	input.close();
	log.close();

}


function _readRecord(input) {

	// Leerzeilen überlesen:

	var line;
	while ((line = input.readLine()) == "\t") {
		if (input.isEOF())
			break;
	}
	
    if (input.isEOF())
		return null;

	var record = "";
	while (line != "\t" && !input.isEOF()) {
		if (record.length > 1)
			record += "\n";

        record += line;
		line = input.readLine();
}


	return record;
}

function __addTag(tag, content, update) {
	// /* Die Funktion geht an die mögliche Position des angegebenen Feldes ("tag") und prüft,
	// ob das Feld bereits vorhanden ist.
	// Wenn es noch nicht vorhanden ist, wird es in einer neuen Zeile mit Inhalt "content" erzeugt.
	// Wenn es bereits vorhanden ist, und das Update ist erwünscht (=true), wird das bestehende Feld überschrieben. (ACHTUNG: Es werden keine Feldwiederholungen behandelt!)

	// Übersicht der Parameter:
	// tag = Die hinzuzufügende Feldbeschreibung, z.B. "0599"
	// content = der im Feld zu ergänzende Inhalt (das einleitende Leerzeichen wird automatisch ergänzt!) z.B. "f"
	// update = true, wenn ein vorhandenes Feld überschrieben werden soll, ansonsten false

	// 2019-09-12 : Marcel Gruss:
	// 2020-08-31 : Christian Baumann: Einen Bug bereinigt. Zwischen 2 geschützte Zeilen muss man zuerst ein '\n'
	// schreiben, um dann den Tag einfügen zu können.
	//  */

	__geheZuKat(tag, "", true);
	content = " " + content;
	var strTag;
	if (!(strTag = application.activeWindow.title.findTag(tag, 0, true, true, false))) {
		application.activeWindow.title.endOfField(false);
		application.activeWindow.title.insertText("\n")
		application.activeWindow.title.insertText(tag + content);
	} else {
		if (update) {
			application.activeWindow.title.startOfField(false);
			application.activeWindow.title.endOfField(true);
			application.activeWindow.title.deleteSelection();
			application.activeWindow.title.insertText(tag + content);
		}
	}
}

function __geheZuKat(kat, ind, append) {
	// __geheZuKat(kat,ind,append)

	// Die Funktion geht in einem Datensatz an die Stelle, an der eine bestimmte neue Kategorie/Indikator der
	// Reihenfolge nach eingefügt werden würde. Übergeben wird als Parameter kat die einzufügende Kategorie
	// und als ind der Indikator.

	// 'kat = übergebene einzufügende Kategorie
	// 'ind = übergebener Indikator
	// 'append = true -> ans Ende eines vorhandenen Felds (das erste Vorkommen oder, wenn nicht vorhanden, genau ein Feld davor ans Ende), sonst: Anfang des ersten Feldes oder dort, wo es stehen müsste
	// '-> append bei noch nicht vorhandenem, einzufügendem Feld immer auf false setzen
	// 'kat_ind = Wert der Kategorie + Indikator
	// 'ta_kat = geprüfte Kategorie der TA (Schleife)
	// 'ta_kat_ind = geprüfte Kategorie der TA + Indikator

	// Historie:
	// 2010-01-09 Stefan Grund		: erstellt
	// 2010-09-18 Bernd			: Definitionen ergaenzt

	var ta_kat_ind; // Indikator des Feldes der TA, in dem die richtige Position gesucht wird
	var kat_ind; //
	var ta_kat; //Feld der TA, in dem die richtige Position gesucht wird (pro durgegangener Zeile)
	//kat -> Übergebenes Feld, dessen Postion gesucht werden soll
	//ind -> Übergebener Indikator des übergebenen Feldes, dessen Postion gesucht werden soll

	application.activeWindow.title.startOfBuffer(false);

	do {

		application.activeWindow.title.lineDown(1, false);
		ta_kat = application.activeWindow.title.tag;
		//das gesuchte Feld wurde gefunden, Indikator ist vorhanden
		if (ta_kat == kat && ind != "") {
			kat_ind = parseInt(kat) + parseInt(ind.charCodeAt(0));
			ta_kat_ind = parseInt(ta_kat) + parseInt(application.activeWindow.title.currentField.substr(5, 1).charCodeAt(0));
			//Prüfung: gesuchte kat ungleich Kat der Zeile oder gesuchte Kat + Indikator größer gleich Kat + Indikator der Zeile
			if (ta_kat != kat || ta_kat_ind >= kat_ind) {
				break;
			}
		}
	} while (ta_kat <= kat && ta_kat != ""); //solange Kat der Zeile kleiner gleich gesuchter kat ist und man nicht am Ende eines Datensatzes ist (zu erkennen daran, dass keine Feldbezeichnung vorhanden ist)

	application.activeWindow.title.startOfField(false);

	// Cursor ist jetzt entweder im gesuchten Feld, falls vorhanden, oder im nächsthöheren Feld, falls nicht vorhanden
	if (append == true) {
		//wenn Feld noch nicht vorhanden, ist der Cursor jetzt am Anfang des nächsthöheren Feldes -> muss eins hoch
		if (ta_kat > kat || ta_kat_ind > kat_ind || ta_kat == "") {
			application.activeWindow.title.lineUp(1, false)
		}
		application.activeWindow.title.endOfField(false);
	}
	return application.activeWindow.title.currentField;
}