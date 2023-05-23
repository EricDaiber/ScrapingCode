*Vorwort:*
Alle verwendeten URLs wurden mit der Datei: "robots.txt" überprüft und erlauben das Web Crawler/Scraping verwendet wird.

------------------------------------------

                                >SCRAPING AKTIEN<
Mein Projekt besteht aus zwei Codes, welche entwickelt wurden, um Aktienkursen und Aktiennamen von den angegebenen URLs zu scrapen und diese in die jeweiligen Excel-Dateien einzutragen.

Der erste Code, "Scraping Aktien.ipynb", erstellt dabei eine eigene Excel-Datei und erstellt innerhalb dieser eine Tabelle. Abschließend werden die Daten in die erstellte Tabelle eingetragen.

Der zweite Code, "Scraping ROI Berechnung.ipynb", baut auf dem ersten Code auf und verwendet eine bereits erstellte Excel-Datei, in der Vergangenheitsdaten eingetragen sind. Der Code hat dabei die Aufgabe die aktuellen Aktienkurse zu scrapen und diese in die Excel-Datei einzutragen. Mit den Daten erfolgt dann die ROI-Berechnung.

Allgemein beginnen beide Codes damit Module wie beispielsweise:

-request => HTML-Inhalte von Webseiten abzurufen
-BeatifulSoup => HTML-Struktur der Webseite wiederzugeben und nötige Daten zu extrahieren
-openpyxl => gescrapte Daten in die Excel-Dateien einzutragen

zu importieren.

Anschließend werden die URLs mit dem "request" Modul überprüft. Dabei wird speziell darauf geachtet, dass der Statuscode als 200 ausgegeben wird. Der Statuscode 200 zeigt an, dass der Zugriff auf alle Informationen der URL/Webseite erfolgreich ist und das Scraping durchgeführt werden kann.

.......
                            -1. Scraping Aktien.ipynb-
Der Code "Scraping Aktien.ipynb" startet zu beginn eine Excel-Datei mit dem Modul "openpyxl". Im Anschluss werden die Spaltenbreite, Zellenschattierung und Rahmenlinien innerhalb der Excel formatiert.

Um eine Tabelle zu beginnen werden Spaltenüberschriften für die Namen und Kurse der Aktien zugewiesen. Mithilfe der Module "request" und "BeautifulSoup" werden die HTML Inhalte aus den URLs ausgelesen und die Aktiennamen werden ausgefiltert. Abschließend werden die ausgefilterten Namen zur Verifizierung ausgegeben und in die angegebenen Zellen eingetragen.

Wie zuvor wird Mithilfe der Module "request" und "BeautifulSoup" die HTML Inhalte erneut aus den URLs ausgelesen. In diesem Fall ist es das Ziel die aktuellen Aktienurse auszufiltern. 

Vor der Eintragung in die Excel wird geprüft, ob die ausgefilterten Daten tatsächlich Zahlen sind, um sicherzustellen, dass es sich um die gewünschten Aktienkurse handelt.

Falls diese Überprüfung erfolgreich ist, werden die Aktienkurse in die Excel übertragen und die Excel wird unter dem Namen 'Aktuelle Daten.xlsx' gespeichert. 

.......
                            -2. Scraping ROI Berechnung.ipynb-
(Im Vergleich zum zuvor genannten Code verwendet dieser Code das Modul "datetime", um das aktuelle Datum zu erfassen.)

Der Code "Scraping ROI Berechnung.ipynb" versucht zu Beginn mithilfe des Moduls "openpyxl.load" die Excel-Datei "ROI Berechnung.xlsx" zu finden und zu laden. Wichtig ist hierbei, dass die Datei sich im selben Ordner wie der Code befindet. Wenn dem so nicht ist, wird nach einem alternativen Dateipfad gefragt. 

Nach der Eingabe des Dateipfad wird dieser überprüft. Falls es wieder scheitert, wird eine Fehlermeldung ausgegeben und der User wird darauf aufgefordert den Code neu zu starten. (Eine Überlegung war es, eine Schleife zu integrieren, die weiterhin nach einem Dateipfad fragt. Wenn der User jedoch die bereits die benötigte Datei noch herunterladen muss, dann könnte er auch einfach den Code neu starten.)

Im nächsten Schritt wird das entsprechende Excel-Blatt mit dem Namen "Top 3 Aktien" ausgewählt. Mit Hilfe der Module "request" und "BeautifulSoup" werden die HTML-Inhalte aus den URLs ausgelesen, um die aktuellen Aktienkurse auszufiltern. Vor der Eintragung in die Excel wird geprüft, ob die ausgefilterten Daten Zahlen sind, um sicherzustellen, dass es sich tatsächlich um die Aktienkurse handelt. 

Falls diese Überprüfung erfolgreich ist, werden die Aktienkurse in die Excel übertragen. Anhand dieser Daten wird dann eine "Return-on-Investment" Berechnung durchgeführt.

Mit dem Modul "datetime.datetime.now().date()" wird das aktuelle Datum erfasst und on die Excel-Zelle "D13" eingetragen.

Abschließend wird versucht, die Excel zu speichern. Bei einem Fehlschlag, wird eine entsprechende Fehlermeldung ausgegeben.
