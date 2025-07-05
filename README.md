Twee2Docx è uno script che converte file .twee (da Twine) in documenti .docx e rimescola i capitoli in modo casuale ma mantenendo il corretto funzionamento dei rimandi interni.

⚠️ Attenzione: al momento della stesura di questo manuale, lo script non è stato ancora testato in fase di pubblicazione. Si consiglia di effettuare prove approfondite prima di usare il documento finale.

1. Scopo del Programma
Twee2Docx nasce per trasformare un file .twee in un documento Word .docx, riorganizzando i capitoli numerati in un ordine casuale, senza rompere i collegamenti tra di essi. È pensato per autori di librigame da stampare.

, cercando di mantenere i capitoli alla minima  distanza possibile, nei limiti del parametro --distanza

3. Requisiti
Python, python-docx, networkx

pip install python-docx
pip install networkx

Testato su Windows 11, al momento non ho modo di provare altri os.

4. Come Eseguire lo Script
Posizionati nella cartella dove si trovano il file twee2docx.py e il tuo .twee. Poi esegui:

python twee2docx.py [opzioni]

--nomefile <nome>
Cosa fa: indica il nome del file .twee da elaborare (senza l’estensione).
Default: se il parametro è omesso usa il primo .twee trovato nella cartella.

--distanza <numero>
Cosa fa: ove possibile imposta la distanza minima (in capitoli) tra un nodo e i suoi collegamenti. Se fallisce dà un warning a console
Default: se il parametro è omesso, 5

--lock <numeri_dei_capitoli>
Cosa fa: blocca uno o più capitoli,
Formato: lista separata da virgole o spazi, ad esempio: --lock 1 3 10 o --lock 1,3,10
Default: se il parametro è omesso, viene bloccato solo il primo capitolo.

--debug
Cosa fa: attiva la modalità debug.
Dopo ogni capitolo verrà inserita una riga in piccolo e corsivo con l’ID originale e i rimandi interni.
Default: disattivato

--no-ottimizza
Cosa fa: disattiva la modalità di rimescolamento capitoli a zone.

Esempi:

**python twee2docx.py --nomefile storia**

Risultato:
* rimescola i capitoli ma tiene bloccato il primo capitolo (default)
* Imposta distanza minima tra capitoli a 5 (default)
* Crea il file storia.docx con il keep together dei capitoli

python twee2docx.py --nomefile storia --distanza 15 --lock 1, 45, 102

* rimescola i capitoli ma tiene bloccati i capitoli 1, 45 e 102
* Imposta distanza minima tra capitoli a 15
* Crea il file storia.docx con il keep together dei capitoli

**python twee2docx.py --nomefile storia --debug**

Dopo ogni capitolo aggiunge in piccolo la numerazione originali di capitolo e link


