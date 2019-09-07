VERSION 0.01
WAS MACHT BC_CONVERT?
	bc_convert überwacht das Verzeichnis C:\temp\barcode.

	Wenn dort eine *.docx oder *.xlsx Datei abgelegt wird, werden Spalten, die mit
	{1}, {2} ... {9} anfangen, derart bearbeitet, dass der Inhalt jener Spalte 
	in ein 	Barcode Bild umgewandelt wird. Die neue Datei wird unter 
	*-bc.docx oder *-bc.xlsx abgelegt. 

MUSEUMPLUS BERICHTE	
	Damit sinnvolle Barcodes entstehen, braucht man entsprechende Berichte aus
	MuseumPlus, die hier in Zukunft gelistet werden sollen (wenn sie existieren).
	

START DES PROGRAMMS
	Damit die Umwandlung funktioniert, muss man das Programm starten. Es läuft 
	normalerweise so lange, bis man es wieder ausmacht (x oben rechts klickt)

	Es wird wohl hier liegen:
	
		z.B. 
		P:\Mengel\BC\64Win7\bc_convert.exe
		P:\Mengel\BC\64Win10\bc_convert.exe

	Man muss eine "Version" des Programms wählen, die zum Prozessor und Betriebssystem
	passt. Ich nehme mal an, dass ich keine Version für 32 bit Prozessoren mehr mache.
	Und dass wir im Augenblick nur mit Windows 7 und demnächst mit Windows 10 arbeiten. 

    Wenn das Programm läuft, sieht man ein schwarzes Fenster, in dem manchmal Status-
    nachrichten erscheinen.
    
"Kompiliert" mit pyinstaller --onefile bc_concert.py.

