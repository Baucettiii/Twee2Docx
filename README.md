# Twee2Docx

Twee2Docx è uno script che converte file .twee (da Twine) in documenti .docx e rimescola i capitoli in modo casuale mantenendo i rimandi dei capitoli. È pensato per autori di librigame da stampare. Il risultato varia parecchi in relazione al numero di capitoli. s

| :exclamation: ATTENZIONE          |
|:---------------------------|
| Al momento della stesura di questo readme, lo script non è stato ancora testato in produzione. Si consiglia di effettuare prove approfondite prima di usare il documento finale. Testato su Windows 11, al momento non ho modo di provare altri os.      |

## Prerequisiti

Python  
python-docx  
networkx  

```
pip install python-docx
pip install networkx
```

## Come Eseguire lo Script

(Il file twee2docx.py e il tuo .twee devono essere nella stessa cartella e python deve essere nel path)
```
python twee2docx.py [opzioni]
```

***--nomefile <nome_twee_senza_estensione>*** : indica il nome del file .twee da elaborare (senza l’estensione). Default: se il parametro è omesso usa il primo .twee trovato nella cartella.

***--inizio <numero_capitolo>*** : permette di iniziare la nunmerazione da un numero custom. Utile in caso il libro sia diviso in parti. Default: se il parametro è omesso, la rifinitura è disattivata

***--lock <numeri_dei_capitoli>*** : blocca uno o più capitoli e ne impedisce in rimescolamento. Formato: lista separata da virgole o spazi, ad esempio: --lock 1 3 10 o --lock 1,3,10. Default: se il parametro è omesso, viene bloccato solo il primo capitolo.

***--debug*** : dopo ogni capitolo inserisce una riga con numero di paragrafo e rimandi originali. Default: disattivato

***--distanza <numero>*** : imposta la distanza minima (in capitoli) tra un nodo e i suoi collegamenti. Se fallisce dà un warning a console. Default: se il parametro è omesso, 5

***--distanza-max <numero>***      : imposta la distanza massima desiderata. Se un link supera questo valore, lo script considera il risultato "sbagliato" e attiva le fasi successive (--correzione e --rifinitura) per cercare di risolverlo. Default: se il parametro è omesso, il controllo è disattivato

***--correzione <numero>***        : esegue un numero di pass di correzione mirata, dove lo script identifica i link che violano --distanza-max e cerca attivamente lo scambio migliore per risolverli. Default: se il parametro è omesso, il default è 1

***--rifinitura <numero>***        : dopo la correzione mirata, esegue migliaia di scambi casuali per compattare ulteriormente la distanza media di tutti i link. Default: se il parametro è omesso, il default è 1

***--tentativi <numero>***         : genera più draft iniziali e sceglie il migliore come punto di partenza per le fasi di correzione e rifinitura. Utile se il primo draft è particolarmente sfortunata. Default: se il parametro è omesso, il default è 1

## Velocità di esecuzione
Dipende dal numero dei capitoli, dai parametri e dalla velocità del processore. Iniziate con parametri relativamente bassi e alzateli per gradi per trovare il vostro sweet spot.

Esempio:
```
python twee2docx.py --distanza 12 --distanza-max 30 --correzione 100 --tentativi 200 --rifinitura 200
```

200 capitoli su un Intel P9

Rinumerazione completata.  
Statistica finale (dopo tutte le fasi): Distanza media: 29.16, Distanza massima: 164  
Tempo di esecuzione totale: 21.93 secondi.
