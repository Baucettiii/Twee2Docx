Twee2Docx è uno script che converte file .twee (da Twine) in documenti .docx e rimescola i capitoli in modo casuale ma mantenendo il corretto funzionamento dei rimandi interni.È pensato per autori di librigame da stampare.

⚠️ Attenzione: al momento della stesura di questo manuale, lo script non è stato ancora testato in fase di pubblicazione. Si consiglia di effettuare prove approfondite prima di usare il documento finale.

3. Requisiti
Python
python-docx
networkx

pip install python-docx
pip install networkx

Testato su Windows 11, al momento non ho modo di provare altri os.

4. Come Eseguire lo Script
Posizionati nella cartella dove si trovano il file twee2docx.py e il tuo .twee. Poi esegui:

python twee2docx.py [opzioni]

--nomefile <nome>
    Cosa fa: indica il nome del file .twee da elaborare (senza l’estensione).
    Default: se il parametro è omesso usa il primo .twee trovato nella cartella.

--inizio <numero>
    Cosa fa: permette di iniziare la nunmerazione da un numero custom.
    Utile in caso il libro sia diviso in parti.
    Default: se il parametro è omesso, la rifinitura è disattivata

--lock <numeri_dei_capitoli>
    Cosa fa: blocca uno o più capitoli e ne impedisce in rimescolamento.
    Formato: lista separata da virgole o spazi, ad esempio: --lock 1 3 10 o --lock 1,3,10
    Default: se il parametro è omesso, viene bloccato solo il primo capitolo.

--debug
    Cosa fa: attiva la modalità debug.
    Dopo ogni capitolo inserisce una riga con numero di paragrafo e rimandi originali.
    Default: disattivato

--distanza <numero>
    Cosa fa: imposta la distanza minima (in capitoli) tra un nodo e i suoi collegamenti. Se fallisce dà un warning a console.
    Default: se il parametro è omesso, 5

--distanza-max <numero>
    Cosa fa: Imposta la distanza massima desiderata. Se un link supera questo valore, lo script considera il risultato "sbagliato" e attiva le fasi successive (--correzione e --rifinitura) per cercare di risolverlo.
    Default: se il parametro è omesso, il controllo è disattivato

--correzione <numero>
    Cosa fa: esegue un numero di pass di correzione mirata, dove lo script identifica i link che violano --distanza-max e cerca attivamente lo scambio migliore per risolverli.
    Default: se il parametro è omesso, il default è 1

--rifinitura <numero>
    Cosa fa: dopo la correzione mirata, esegue migliaia di scambi casuali per compattare ulteriormente la distanza media di tutti i link. 
    Default: se il parametro è omesso, il default è 1

--tentativi <numero>
    Cosa fa: genera più draft iniziali e sceglie il migliore come punto di partenza per le fasi di correzione e rifinitura. Utile se il primo draft è particolarmente sfortunata.
    Default: se il parametro è omesso, il default è 1


Esempi:

python twee2docx.py --distanza 12 --distanza-max 30 --correzione 60 --tentativi 10 --rifinitura 100
