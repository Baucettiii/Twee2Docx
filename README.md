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

`--nomefile <nome_twee_senza_estensione>` : indica il nome del file .twee da elaborare (senza l’estensione). Default: usa il primo `.twee` trovato nella cartella.

`--inizio <numero_capitolo>` : permette di iniziare la nunmerazione da un numero custom. Utile in caso il libro sia diviso in parti. Default: `1`

`--lock <numeri_dei_capitoli>` : blocca uno o più capitoli e ne impedisce in rimescolamento. Formato: lista separata da virgole o spazi, ad esempio: `--lock 1 3 10` o `--lock 1,3,10`. Default: `1`

`--debug` : dopo ogni capitolo inserisce una riga con numero di paragrafo e rimandi originali. Default: disattivato

`--distanza-min <numero>` : imposta la distanza minima (in capitoli) tra un nodo e i suoi collegamenti. Default: `10`

`--correzione <numero>`        : esegue un numero di pass e cerca attivamente lo scambio migliore per risolvere le violazioni di --ditanza-min. Default: `20`

### Considerazioni

Risultati in termini di posizionamento: come accennato, non sono ancora riuscito a testarlo seriamente sul vero flusso di un librogioco. Le medie sono accettabili ma ci sono ancora dei picchi che non mi soddisfano. Non sono sicuro se siano fisiologici o se siano migliorabili. Work in progress.

Proposta di metodo di lavoro ibrido ideale (da testare): dividere il libro in parti con 2 o tre capitoli di connessioni tra le parti, bloccare quei capitoli ed effettuare una rinunmerazione per parte.

### Esempio di script e statistiche:
```
python twee2docx.py --nomefile test42 --distanza-min 12 --correzione 30
```

