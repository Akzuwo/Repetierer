
![](https://raw.githubusercontent.com/srpnt3/Repetierer/master/public/images/banner.png)

# Repetierer

## Deutsch
Repetierer wurde entwickelt, damit Lehrpersonen zufällig eine Person aus einer Excel-Liste auswählen können.
Die ausgewählte Person kann direkt im Programm bewertet werden.
Diese Note wird danach wieder in die Excel-Datei geschrieben.

### Wie funktioniert die zufällige Auswahl?
Die Auswahl ist nicht für alle Personen immer gleich wahrscheinlich.
Das Programm schaut, wie oft eine Person bereits bewertet wurde.
Wer noch wenige oder gar keine Noten hat, wird wahrscheinlicher ausgewählt.
Wer schon öfter dran war, wird weniger wahrscheinlich ausgewählt.

Man kann sich das wie Lose in einem Topf vorstellen:
Eine Person ohne Note bekommt viele Lose in den Topf.
Eine Person mit einer Note bekommt weniger Lose.
Eine Person mit noch mehr Noten bekommt noch weniger Lose.
Dann zieht das Programm zufällig ein Los.
Je mehr Lose eine Person im Topf hat, desto größer ist ihre Chance, gezogen zu werden.

Wenn eine Person also schon einmal vorgekommen ist und danach eine Note erhält, hat sie beim nächsten Mal weniger Lose im Topf.
Sie kann immer noch ausgewählt werden, aber die Chance ist kleiner als vorher.
So werden mit der Zeit eher Personen ausgewählt, die noch nicht oder seltener bewertet wurden.

Genauer gesagt nutzt das Programm diese Formel:
`e^(6 - Anzahl der Noten) - 1`

Ein einfaches Beispiel:
Eine Person mit 0 Noten hat ungefähr `402` Lose.
Eine Person mit 1 Note hat ungefähr `147` Lose.
Nach der ersten Note hat diese Person also nur noch etwa ein Drittel so viele Lose wie vorher.
Die genaue Wahrscheinlichkeit hängt aber immer davon ab, wie viele Lose alle anderen Personen gerade haben.

Personen mit 6 Noten werden nicht mehr zufällig ausgewählt.

### Verwendung
Die Benutzeroberfläche ist aktuell auf Deutsch.
Zusätzlich zum Programm gibt es im Repository und/oder in den Releases eine Excel-Vorlage.
Diese Datei kann als Vorlage verwendet und an die eigenen Bedürfnisse angepasst werden.
Aktuell werden bis zu 25 Personen, maximal 6 Noten pro Person und beliebig viele Klassen/Fächer unterstützt.

### Neue Funktionen
- **Sitzungsprotokoll exportieren:** Im Menü kann die aktuelle Sitzung als CSV oder PDF exportiert werden. Das Protokoll enthält Zeitpunkt, Datum, Klasse, Person, Aktion, Note und Excel-Status.
- **Klassenliste importieren:** Eine bestehende Excel- oder CSV-Datei kann importiert und als neue Repetierer-Datei gespeichert werden. Das Programm übernimmt bis zu 25 eindeutige Namen und legt die nötige Excel-Struktur automatisch an.
- **Letzte Aktion rückgängig machen:** Die letzte Note oder Joker-Nutzung der aktuellen Sitzung kann rückgängig gemacht werden, solange der passende Excel-Eintrag noch unverändert ist. Falls Excel beim Speichern gesperrt war, wird der offene lokale Eintrag entfernt.
- **Wiederholen:** Eine versehentlich rückgängig gemachte Aktion kann direkt wiederhergestellt werden.
- **Abwesenheiten:** Personen können für den aktuellen Tag als abwesend markiert werden. Diese Personen werden bei der zufälligen Auswahl und bei der freiwilligen Auswahl übersprungen, ohne dass die Excel-Datei verändert wird.
- **Wiggersche Regel:** Wenn eine Person ihren Joker setzt, fällt ihre Auswahlwahrscheinlichkeit für eine begrenzte Zeit stark ab. Standardmäßig gilt dieser Malus 2 Stunden und reduziert das Gewicht auf 5%. Die Person kann weiterhin ausgewählt werden, aber es wird sehr unwahrscheinlich. Die Regel ist als Balancing-Maßnahme gedacht, damit ein Joker nicht sofort wieder durch eine neue Ziehung entwertet wird.

## English
Repetierer was created to enable teachers to test students by selecting a random individual from a list contained in an Excel file.
The teacher then has the ability to grade the student directly inside the application.
This grade will then be written in the Excel file containing the list of the students.
The algorithm's probability to select a specific student is based on the amount of times the student has already been tested and thus graded.

## Selection Probability
The selection is not equally likely for everyone all the time.
The program checks how many grades each student already has.
Students with fewer grades are more likely to be selected.
Students who have already been tested are less likely to be selected again.

You can imagine it like tickets in a bowl:
A student with no grade gets many tickets.
A student with one grade gets fewer tickets.
A student with more grades gets even fewer tickets.
The program then randomly draws one ticket.
The more tickets a student has, the higher their chance of being selected.

So, if a student has already appeared once and receives a grade, they will have fewer tickets next time.
They can still be selected, but the chance is lower than before.
This helps the program pick students who have not been tested yet, or who have been tested less often.

More precisely, the program uses this formula:
`e^(6 - number of grades) - 1`

Simple example:
A student with 0 grades has about `402` tickets.
A student with 1 grade has about `147` tickets.
After the first grade, that student therefore has only about one third as many tickets as before.
The exact probability still depends on how many tickets all other students currently have.

Students with 6 grades are no longer included in the random selection.

## Usage
(Note that the UI is currently in German, but feel free to add a translation as this project is Open-Source)
In addition to downloading the program there is also an Excel template file located in the repository and/or the releases.
Use said file as a template and edit it according to Your needs.
It currently supports up to 25 students (rows), a maximum six grades per student (columns) and infinite classes/subjects (individual worksheets).
The program itself is pretty self-explanatory.

## New Features
- **Session protocol export:** Export the current session as CSV or PDF, including time, date, class, person, action, grade and Excel write status.
- **Class list import:** Import an existing Excel or CSV list and save it as a new Repetierer workbook. Up to 25 unique names are converted into the required format.
- **Undo last action:** Undo the most recent grade or joker action from the current session as long as the corresponding Excel entry still matches.
- **Redo:** Restore an action that was undone by accident.
- **Absences:** Mark students as absent for the current day. They are skipped during random and voluntary selection without changing the Excel file.
- **Wiggersche Regel:** When a student uses a joker, their selection probability is heavily reduced for a limited time. By default, the penalty lasts 2 hours and reduces their weight to 5%. They can still be selected, but it becomes very unlikely.

![](https://raw.githubusercontent.com/srpnt3/Repetierer/master/public/images/preview.png)

## Future Development
I do not plan to continue this project any further, since I will soon finish the school for which I developed the application.
Of course the project is Open-Source so anyone can continue development.
