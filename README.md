Twee2Docx è uno script che converte file .twee (da Twine) in documenti .docx e rimescola i capitoli in modo casuale ma mantenendo il corretto funzionamento dei rimandi interni.

⚠️ Attenzione: al momento della stesura di questo manuale, lo script non è stato ancora testato in fase di pubblicazione. Si consiglia di effettuare prove approfondite prima di usare il documento finale.

1. Scopo del Programma
Twee2Docx nasce per trasformare un file .twee in un documento Word .docx, riorganizzando i capitoli numerati in un ordine casuale, senza rompere i collegamenti tra di essi. È pensato per autori di librigame da stampare.

2. Come Funziona
Il processo si compone di tre fasi principali:

1. Analisi del file .twee
Lo script legge il file, ignorando le intestazioni tecniche di Twine.

2. Rinumerazione controllata
I capitoli vengono rinumerati casualmente, ma seguendo alcune regole:

Casualità: i nuovi numeri non sono sequenziali.

Vincoli opzionali: puoi bloccare alcuni capitoli o stabilire una distanza minima tra quelli collegati.

I rimandi interni vengono aggiornati in automatico in base alla nuova numerazione.

3. Esportazione in .docx
Viene creato un file .docx con i capitoli rimescolati e numerati.
È applicata una formattazione che cerca di evitare la divisione di un capitolo tra due pagine.

3. Requisiti
Per usare lo script è necessario avere installato Python e la libreria python-docx.

Installa la libreria con: pip install python-docx

4. Come Eseguire lo Script
Posizionati nella cartella dove si trovano il file ricalcolo.py e il tuo .twee. Poi esegui:

python twee2docx.py [opzioni]

5. Opzioni Disponibili

--nomefile <nome>
Cosa fa: indica il nome del file .twee da elaborare (senza l’estensione).
Default: se non specificato, usa il primo .twee trovato nella cartella.

--distanza <numero>
Cosa fa: imposta la distanza minima (in capitoli) tra un nodo e i suoi collegamenti.

Utile per: evitare che un capitolo collegato compaia troppo vicino.
Default: 5

--lock <numeri_dei_capitoli>
Cosa fa: blocca uno o più capitoli,

Formato: lista separata da virgole o spazi, ad esempio: --lock 1 3 10 o --lock 1,3,10
Default: se omesso, viene bloccato solo il primo capitolo.

--debug
Cosa fa: attiva la modalità debug.

Dopo ogni capitolo verrà inserita una riga in piccolo e corsivo con l’ID originale e i rimandi interni.

Default: disattivato

6. Esempi Pratici

Esecuzione base:
python twee2docx.py --nomefile storia
Blocca solo il primo capitolo

Imposta distanza minima a 5

Crea il file storia.docx

Bloccare capitoli e aumentare la distanza:
python twee2docx.py --nomefile storia --distanza 15 --lock 1, 45, 102
Imposta distanza minima a 15

Blocca i capitoli 1, 45, 102

Attivare il debug:
python twee2docx.py --nomefile storia --debug
Aggiunge informazioni tecniche in piccolo dopo ogni capitolo

Tutto insieme:
python ricalcolo.py --nomefile storia --distanza 10 --lock 1 50 --debug
Imposta distanza 10

Blocca i capitoli 1 e 50
Attiva debug
Output: storia.docx
