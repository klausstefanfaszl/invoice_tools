Das invoice_tool soll jetzt noch um eine funktion "mailto" erweitert mit folgenden Funktionen erweitert werden:

- Es gibt dafür wieder eine zusätzliche xml-config-Dattei invoice_tool_mailto_config.xml 
- Es werden Mail von einer in der config-Datei angegebenen Mailadresse versendet.
- Der text der mail ist Standardisiert und steht als html-Text in der config datei
- In der Konfig-Datei steht das Verzeichnis in dem sich die zu versendenden PDF-Rechnungen befinden.
  Die Dateinamen enthalten eine Rechnungnummer
- Es gibt einen optionalen Parameter --rg-nr    oder -r  mit dem kann eine einzelne Rechnung-Nr angegeben werden
- Es gibt einen optionalen Parameter --min-r-nr  mit dem kann die kleinste zu versendeden Rechnungsnumer  angegeben werden
- Es gibt einen optionalen Parameter --max-r-nr  mit dem kann die kgrößste zu versendeden Rechnungsnumer  angegeben werden
- Die Mailadresse kann auf verschiednenen Wegen angegeben sein, wobei das Verfahren über eine optionalen Parameter --adress-modus [0-2]  oder -a
   angegeben werden kann:
   (1) Sie steht in Rechnung hinter  einem Text "email:" als aus der PDF-extrahierbarere Text
   (2) Sie wird aus einer Datenbank in der Config-Datei angegebenen Mysql-Datenbank (DB_name, Accout, PW,  Port) über einen 
       configurierbaren in der Config-Datei hinterlegten SQL-Befehl ermittelt, wobei bis zu 2 SQL-Befehle angegeben werden können
	   und der zweite verwendet wird, wenn der erste ein leeres Ergebnis liefert
   (0) es Wird (1) und (2) versucht
  Der Standardwert des parameers ist 0 
  
 - Es gibt wieder --dry-run der nur simuliert und anzeigt wohin gesendet 
 - Es gibt wieder einen --debug-modus  mit genauer Anzeige was passiert
 - es gibt einen --noconfirm  Modus bei dem nach der Generierung der Mails alle sofort versendet werden
   Standardmäßig muus man noch einmal das Versenden der mails noch einmal bestätigen
   
   
   
Anmerkung:
Liefert die SQL-Befehle mehr als 1 Spalte zurück sollte immer in der ersten Spalte eine Rechnungsnummer stehen, die als Zahl zu interpretieren ist,
d.h. führende Nullen sind zu ignorieren.

	  
   
  
    