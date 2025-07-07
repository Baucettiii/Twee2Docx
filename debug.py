import random
import time
import argparse
import math

# --- Strutture Dati per la Storia ---

class Capitolo:
    """Rappresenta un singolo capitolo (passage) con i suoi dati."""
    def __init__(self, id, contenuto, tipo="standard"):
        self.id = id
        self.contenuto = contenuto
        self.tipo = tipo  # es. "inizio", "tronco_a", "finale", "riunificazione"
        self.rimandi = [] # Lista di ID a cui questo capitolo punta
        self.pos_x = random.randint(50, 1200)
        self.pos_y = random.randint(50, 800)

    def aggiungi_rimando(self, id_destinazione):
        """Aggiunge un link a un altro capitolo."""
        if id_destinazione not in self.rimandi:
            self.rimandi.append(id_destinazione)

    def __repr__(self):
        return f"Capitolo(id={self.id}, tipo='{self.tipo}', rimandi={self.rimandi})"

    def to_twee(self):
        """Converte il capitolo nel formato stringa per il file .twee."""
        intestazione_capitolo = f':: {self.id} [{self.tipo}] {{"position":"{self.pos_x},{self.pos_y}","size":"100,100"}}'
        rimandi_str = "\n".join([f"[[{dest}]]" for dest in self.rimandi])
        return f"{intestazione_capitolo}\n{self.contenuto}\n{rimandi_str}"

# --- Funzione Principale di Generazione ---

def genera_librogame(
    num_capitoli: int,
    nome_file: str,
    verbose: bool,
    num_tronchi: int,
    num_riunificazioni: int,
    num_finali: int
):
    """
    Genera un file .twee con una struttura a librogame, con tronchi narrativi,
    punti di riunificazione e finali configurabili.
    """
    start_time = time.time()
    
    if verbose:
        print("--- Avvio Generatore di Storie a Librogame ---")
        print(f"Parametri: {num_capitoli} capitoli, {num_tronchi} tronchi, {num_riunificazioni} riunificazioni, {num_finali} finali")
        print(f"Output su '{nome_file}'")
        print("-" * 50)

    # 1. VALIDAZIONE E DEFINIZIONE DELLA STRUTTURA
    if num_tronchi <= 0:
        print("❌ Errore: Il numero di tronchi deve essere maggiore di zero.")
        return

    capitoli_speciali = 1 + num_riunificazioni + num_finali # Inizio + Riunificazioni + Finali
    capitoli_per_storia = num_capitoli - capitoli_speciali
    
    if capitoli_per_storia < num_tronchi * 5: # Assicuriamo almeno 5 capitoli per tronco
        print(f"❌ Errore: Pochi capitoli per la struttura richiesta. Con {num_capitoli} capitoli totali e {capitoli_speciali} speciali, non è possibile creare {num_tronchi} tronchi di dimensioni adeguate.")
        print("   Prova ad aumentare il numero di capitoli o a ridurre il numero di tronchi/finali/riunificazioni.")
        return

    capitoli_per_tronco = math.floor(capitoli_per_storia / num_tronchi)

    if verbose:
        print("[INFO] Partizionamento capitoli:")
        print(f"  - Tronchi narrativi: {num_tronchi} da {capitoli_per_tronco} capitoli ciascuno")
        print(f"  - Nodi speciali: 1 Inizio, {num_riunificazioni} Riunificazione, {num_finali} Finali")
        print("-" * 50)

    # 2. GENERAZIONE NODI (IN MEMORIA)
    parole_casuali = ["Aereo", "Albero", "Anello", "Ape", "Auto", "Balena", "Barca", "Bicchiere", "Borsa", "Bottiglia", "Cactus", "Calzino", "Camera", "Campana", "Cane", "Cappello", "Casa", "Castello", "Cavallo", "Chiave", "Chitarra", "Cielo", "Cintura", "Città", "Computer", "Coniglio", "Coppa", "Cuore", "Delfino", "Diamante", "Dinosauro", "Elefante", "Farfalla", "Fiume", "Fiore", "Foglia", "Forchetta", "Foresta", "Formica", "Fotografia", "Fungo", "Gabbia", "Gatto", "Ghiaccio", "Giornale", "Gomma", "Guanto", "Guerriero", "Imbuto", "Isola", "Lampada", "Leone", "Lettera", "Libro", "Limone", "Lupo", "Maglietta", "Mano", "Mappa", "Mare", "Martello", "Medusa", "Mela", "Microfono", "Mongolfiera", "Montagna", "Mostro", "Mucca", "Nave", "Nuvola", "Occhiali", "Ombrello", "Orologio", "Orso", "Palla", "Panda", "Pane", "Pecora", "Pinguino", "Ponte", "Porta", "Pozzo", "Quadro", "Ragno", "Razzo", "Riccio", "Robot", "Ruota", "Scarpa", "Scatola", "Scimmia", "Serpente", "Sirena", "Sole", "Spada", "Specchio", "Stella", "Stivale", "Tavolo", "Tazza", "Telefono", "Tigre", "Treno", "Uccello", "Uovo", "Vampiro", "Vaso", "Violino", "Volpe", "Zaino", "Zebra"]
    
    storia = {} # Dizionario di oggetti Capitolo, per un facile accesso tramite ID
    current_id = 1

    # Creazione Inizio
    storia[current_id] = Capitolo(current_id, "L'inizio del Viaggio", "inizio")
    id_inizio = current_id
    current_id += 1

    # Creazione Tronchi Narrativi
    partizioni = {}
    for i in range(num_tronchi):
        nome_tronco = f"tronco_{chr(65 + i)}"
        start_id = current_id
        end_id = start_id + capitoli_per_tronco - 1
        partizioni[nome_tronco] = list(range(start_id, end_id + 1))
        for j in range(capitoli_per_tronco):
            storia[current_id] = Capitolo(current_id, random.choice(parole_casuali), nome_tronco)
            current_id += 1

    # Creazione Nodi di Riunificazione
    id_riunificazioni = list(range(current_id, current_id + num_riunificazioni))
    for id_nodo in id_riunificazioni:
        storia[id_nodo] = Capitolo(id_nodo, "Un punto d'incontro", "riunificazione")
        current_id += 1
        
    # Creazione Finali
    id_finali = list(range(current_id, current_id + num_finali))
    for id_nodo in id_finali:
        storia[id_nodo] = Capitolo(id_nodo, "La Fine", "finale")
        current_id += 1

    # 3. COLLEGAMENTO DEI NODI (LINKING)
    if verbose: print("[INFO] Collegamento dei capitoli secondo la struttura...")

    for nome_tronco, ids in partizioni.items():
        storia[id_inizio].aggiungi_rimando(ids[0])

    for nome_tronco, ids in partizioni.items():
        for i, id_cap in enumerate(ids):
            if i == len(ids) - 1:
                if id_riunificazioni:
                    storia[id_cap].aggiungi_rimando(random.choice(id_riunificazioni))
                elif id_finali:
                    storia[id_cap].aggiungi_rimando(random.choice(id_finali))
                continue

            storia[id_cap].aggiungi_rimando(ids[i+1])
            scelte_desiderate = random.choices([1, 2, 3], weights=[2, 90, 8], k=1)[0]
            rimandi_aggiuntivi_necessari = scelte_desiderate - 1
            if rimandi_aggiuntivi_necessari <= 0: continue
            pool_destinazioni_aggiuntive = ids[i+2:]
            if not pool_destinazioni_aggiuntive: continue
            num_link_da_aggiungere = min(rimandi_aggiuntivi_necessari, len(pool_destinazioni_aggiuntive))
            destinazioni_scelte = random.sample(pool_destinazioni_aggiuntive, num_link_da_aggiungere)
            for dest in destinazioni_scelte:
                storia[id_cap].aggiungi_rimando(dest)

    altri_tronchi = list(partizioni.keys())
    for i, (nome_tronco, ids) in enumerate(partizioni.items()):
        for id_cap in ids:
            if random.random() < 0.05:
                tronco_dest = random.choice(altri_tronchi[i:] + altri_tronchi[:i])
                if tronco_dest != nome_tronco:
                    id_dest = random.choice(partizioni[tronco_dest])
                    storia[id_cap].aggiungi_rimando(id_dest)
                    if verbose: print(f"  - Cross-link creato: {id_cap} -> {id_dest}")

    if id_finali:
        for id_cap in id_riunificazioni:
            storia[id_cap].aggiungi_rimando(random.choice(id_finali))
            if random.random() < 0.5:
                storia[id_cap].aggiungi_rimando(random.choice(id_finali))

    # 4. CALCOLO STATISTICHE DISTANZE
    distanze = []
    for capitolo in storia.values():
        for id_dest in capitolo.rimandi:
            dist = abs(id_dest - capitolo.id)
            distanze.append(dist)
    
    dist_min, dist_max, dist_media = 0, 0, 0.0
    if distanze:
        dist_min = min(distanze)
        dist_max = max(distanze)
        dist_media = sum(distanze) / len(distanze)

    # 5. SCRITTURA DEL FILE
    if verbose: print("[INFO] Scrittura del file .twee...")
    
    intestazione_twee = f""":: StoryTitle\nLibrogame Generato\n\n:: StoryData\n{{\n  "ifid": "{str(random.randint(1000,9999))}-{str(random.randint(1000,9999))}",\n  "format": "Harlowe",\n  "format-version": "3.3.9",\n  "start": "1"\n}}\n\n"""
    
    contenuto_finale = [intestazione_twee]
    for id_cap in sorted(storia.keys()):
        contenuto_finale.append(storia[id_cap].to_twee())

    try:
        with open(nome_file, 'w', encoding='utf-8') as f:
            f.write("\n\n".join(contenuto_finale))
    except IOError as e:
        print(f"❌ Errore durante la scrittura del file: {e}")
        return

    # 6. REPORT FINALE
    end_time = time.time()
    durata = end_time - start_time
    
    print("-" * 50)
    print(f"✅ Successo! File '{nome_file}' creato con {len(storia)} capitoli.")
    print("\n--- REPORT FINALE ---")
    print(f"File generato:          {nome_file}")
    print(f"Capitoli totali creati: {len(storia)}")
    print(f"Struttura utilizzata:   {num_tronchi} tronchi narrativi")
    print(f"                        {num_riunificazioni} punto/i di riunificazione")
    print(f"                        {num_finali} finali")
    
    print("\n--- Statistiche Distanze Rimandi ---")
    print("Nota: non essendoci un'ottimizzazione, non esiste uno stato 'prima' e 'dopo'.")
    print("Le statistiche descrivono la struttura finale generata:")
    print(f"Distanza Minima:        {dist_min}")
    print(f"Distanza Massima:       {dist_max}")
    print(f"Distanza Media:         {dist_media:.2f}")

    print(f"\nTempo di esecuzione:    {durata:.4f} secondi")
    print("-" * 50)


def main():
    """Funzione principale che gestisce i parametri da riga di comando."""
    parser = argparse.ArgumentParser(
        description="Genera un file .twee con una struttura a librogame.",
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument(
        '--capitoli', type=int, default=100,
        help='Il numero totale di capitoli da generare.\nDefault: 100'
    )
    parser.add_argument(
        '--tronchi', type=int, default=3,
        help='Il numero di tronchi narrativi principali.\nDefault: 3'
    )
    parser.add_argument(
        '--riunificazioni', type=int, default=1,
        help='Il numero di punti di riunificazione.\nDefault: 1'
    )
    parser.add_argument(
        '--finali', type=int, default=3,
        help='Il numero di capitoli finali.\nDefault: 3'
    )
    parser.add_argument(
        '-v', '--verbose', action='store_true',
        help='Attiva la modalità verbosa per visualizzare i dettagli.'
    )
    args = parser.parse_args()
    
    NOME_FILE_OUTPUT = "librogame.twee"
    
    genera_librogame(
        args.capitoli,
        NOME_FILE_OUTPUT,
        args.verbose,
        args.tronchi,
        args.riunificazioni,
        args.finali
    )

if __name__ == '__main__':
    main()
