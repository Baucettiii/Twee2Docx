# Twee2Docx


**Twee2Docx** è uno script che converte file `.twee` (da Twine 2.10 al momento) in documenti `.docx` e rimescola i capitoli in modo casuale, mantenendo intatti i rimandi tra i paragrafi. È pensato per autori di librigame da stampare. Il risultato può variare sensibilmente in base al numero di capitoli.

| :exclamation: ATTENZIONE          |
|:---------------------------|
| Al momento della stesura di questo readme, lo script non è stato ancora testato in produzione. Si consiglia di effettuare prove approfondite prima di usare il documento finale. Testato su Windows 11, al momento non ho modo di provare altri os.      |

## Prerequisiti

`Python`  
`python-docx`  
`networkx`  

```
pip install python-docx
pip install networkx
```

### Come Eseguire lo Script

(Il file `twee2docx.py` e il tuo `.twee` devono essere nella stessa cartella e python deve essere nel path)
```
python twee2docx.py [opzioni]
```

`--nomefile <nome_twee_senza_estensione>` : indica il nome del file .twee da elaborare (senza l’estensione). Default: se il parametro è omesso usa il primo `.twee` trovato nella cartella.

`--inizio <numero_capitolo>` : permette di iniziare la nunmerazione da un numero custom. Utile in caso il libro sia diviso in parti. Default: se il parametro è omesso, la rifinitura è disattivata

`--lock <numeri_dei_capitoli>` : blocca uno o più capitoli e ne impedisce in rimescolamento. Formato: lista separata da virgole o spazi, ad esempio: `--lock 1 3 10` o `--lock 1,3,10`. Default: se il parametro è omesso, viene bloccato solo il primo capitolo.

`--debug` : dopo ogni capitolo inserisce una riga con numero di paragrafo e rimandi originali. Default: disattivato

`--distanza <numero>` : imposta la distanza minima (in capitoli) tra un nodo e i suoi collegamenti. Se fallisce dà un warning a console. Default: se il parametro è omesso, `5`

`--distanza-max <numero>`      : imposta la distanza massima desiderata. Se un link supera questo valore, lo script considera il risultato "sbagliato" e attiva le fasi successive (`--correzione` e `--rifinitura`) per cercare di risolverlo. Default: se il parametro è omesso, il controllo è disattivato

`--correzione <numero>`        : esegue un numero di pass di correzione mirata, dove lo script identifica i link che violano `--distanza-max` e cerca attivamente lo scambio migliore per risolverli. Default: se il parametro è omesso, il default è `1`

`--rifinitura <numero>`        : dopo la correzione mirata, esegue migliaia di scambi casuali per compattare ulteriormente la distanza media di tutti i link. Default: se il parametro è omesso, il default è `1`

`--tentativi <numero>`         : genera più draft iniziali e sceglie il migliore come punto di partenza per le fasi di correzione e rifinitura. Utile se il primo draft è particolarmente sfortunata. Default: se il parametro è omesso, il default è `1`

### Considerazioni
Velocità di esecuzione: dipende dal numero dei capitoli, dai parametri e dalla velocità del processore. Iniziate con parametri relativamente bassi e alzateli per gradi per trovare il vostro sweet spot (esempi di tempi di esecuzione qua sotto)

Risultati in termini di posizionamento: come accennato, non sono ancora riuscito a testarlo seriamente sul vero flusso di un librogioco. Le medie sono accettabili ma ci sono ancora dei picchi che non mi soddisfano ma non sono sicuro se siano fisiologici o se siano migliorabili. Work in progress.

Proposta di metodo di lavoro ibrido ideale (da testare): dividere il libro in parti con 2 o tre capitoli di connessioni tra le parti, bloccare quei capitoli ed effettuare una rinunmerazione per parte.

### Esempio di script e statistiche:
```
python twee2docx.py --nomefile test42 --distanza 12 --distanza-max 30 --correzione 100 --tentativi 200 --rifinitura 200
```

Esecuzione 200 capitoli su un Intel P9

Rinumerazione completata.  
- Statistica finale: Distanza media: 29.16, Distanza massima: 164  
- Tempo di esecuzione totale: 21.93 secondi.
