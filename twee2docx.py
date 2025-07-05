# -*- coding: utf-8 -*-

import re
import docx
import argparse
import glob
import os
import random
from docx.shared import Pt

# Prova a importare networkx e avvisa l'utente se manca
try:
    import networkx as nx
    from networkx.algorithms import community
except ImportError:
    print("ERRORE: La libreria 'networkx' non è installata.")
    print("Per favore, installala eseguendo il comando: pip install networkx")
    exit()

def parse_twee_file(file_path):
    """
    Analizza un file di testo in formato Twee e estrae i passaggi (capitoli).
    """
    print(f"Inizio analisi del file: {file_path}")
    passages = []
    current_passage = None
    ignoring_metadata_block = False

    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            for line in f:
                match = re.match(r'^::\s*([^[{]+)', line)
                if match:
                    original_id = match.group(1).strip()
                    if original_id in ["StoryTitle", "StoryData"]:
                        ignoring_metadata_block = True
                        if current_passage:
                            current_passage['content'] = current_passage['content'].strip()
                            passages.append(current_passage)
                            current_passage = None
                        continue
                    
                    ignoring_metadata_block = False
                    if current_passage:
                        current_passage['content'] = current_passage['content'].strip()
                        passages.append(current_passage)

                    current_passage = {
                        'original_id': original_id,
                        'title': original_id,
                        'content': ''
                    }
                elif not ignoring_metadata_block and current_passage:
                    current_passage['content'] += line
        
        if not ignoring_metadata_block and current_passage:
            current_passage['content'] = current_passage['content'].strip()
            passages.append(current_passage)

    except FileNotFoundError:
        print(f"Errore: File non trovato a questo percorso: {file_path}")
        return []
    
    print(f"Analisi completata. Trovati {len(passages)} passaggi validi.")
    return passages

def _place_passages_in_zone(passages_in_zone, available_ids, distance, id_map, links_graph):
    """
    Funzione ausiliaria per posizionare un gruppo di passaggi in un blocco di ID disponibili.
    """
    processing_queue = list(passages_in_zone)
    random.shuffle(processing_queue)
    deferred_passages = []

    while processing_queue:
        passage_to_place = processing_queue.pop(0)
        original_id = passage_to_place['original_id']

        placed_neighbor_new_ids = {id_map[dest_id] for dest_id in links_graph.get(original_id, []) if dest_id in id_map}
        for src_id, linked_ids in links_graph.items():
            if original_id in linked_ids and src_id in id_map:
                placed_neighbor_new_ids.add(id_map[src_id])

        valid_positions = [pid for pid in available_ids if all(abs(pid - nid) >= distance for nid in placed_neighbor_new_ids)]

        if valid_positions:
            cost_map = [{'id': pos, 'cost': sum(abs(pos - nid) for nid in placed_neighbor_new_ids)} for pos in valid_positions]
            best_fit_id = min(cost_map, key=lambda x: x['cost'])['id']
            
            available_ids.remove(best_fit_id)
            id_map[original_id] = best_fit_id

            if deferred_passages:
                processing_queue.extend(deferred_passages)
                deferred_passages = []
        else:
            deferred_passages.append(passage_to_place)

        if len(deferred_passages) == len(processing_queue) and deferred_passages:
            print("--- Rilevato stallo: impossibile soddisfare i vincoli per i capitoli rimanenti. Forzatura del posizionamento. ---")
            for stuck_passage in deferred_passages:
                stuck_original_id = stuck_passage['original_id']
                if not available_ids:
                    print(f"Errore critico: non ci sono più ID disponibili per il capitolo '{stuck_original_id}'.")
                    continue
                
                forced_id = available_ids.pop(0)
                id_map[stuck_original_id] = forced_id
                
                violating_neighbors_details = [f"capitolo {nid}" for nid in placed_neighbor_new_ids if abs(forced_id - nid) < distance]
                warning_message = (f"Attenzione: non è stato possibile soddisfare il vincolo di distanza per il capitolo {stuck_original_id}. "
                                   f"Assegnato al primo posto disponibile: {forced_id}")
                if violating_neighbors_details:
                    warning_message += f", che viola la distanza minima con: {', '.join(violating_neighbors_details)}"
                print(warning_message)
            break

def _order_zones_intelligently(communities, G):
    """
    Ordina le zone (community) in base alla loro interconnessione per minimizzare
    la distanza tra i blocchi di numeri che verranno loro assegnati.
    """
    if len(communities) <= 1:
        return communities

    # Crea un "meta-grafo" dove ogni nodo è una zona
    meta_graph = nx.Graph()
    community_map = {node: i for i, com in enumerate(communities) for node in com}
    
    for u, v in G.edges():
        if u in community_map and v in community_map:
            c1, c2 = community_map[u], community_map[v]
            if c1 != c2:
                if meta_graph.has_edge(c1, c2):
                    meta_graph[c1][c2]['weight'] += 1
                else:
                    meta_graph.add_edge(c1, c2, weight=1)

    if not meta_graph.edges():
        # Se le zone non sono connesse, un ordine casuale va bene
        random.shuffle(communities)
        return communities

    # Trova un percorso che massimizza il peso degli archi attraversati (approccio greedy)
    path = []
    # Inizia dall'arco con il peso maggiore
    start_edge = max(meta_graph.edges(data=True), key=lambda x: x[2]['weight'])
    path.extend([start_edge[0], start_edge[1]])
    
    remaining_nodes = set(meta_graph.nodes()) - set(path)
    
    while remaining_nodes:
        best_candidate, best_score, insert_at_end = None, -1, True
        
        # Cerca il miglior nodo da attaccare all'inizio o alla fine del percorso
        for node in remaining_nodes:
            # Punteggio per attaccarlo all'inizio
            score_start = meta_graph.get_edge_data(node, path[0], default={'weight': 0})['weight']
            # Punteggio per attaccarlo alla fine
            score_end = meta_graph.get_edge_data(node, path[-1], default={'weight': 0})['weight']
            
            if score_start > best_score:
                best_score, best_candidate, insert_at_end = score_start, node, False
            if score_end > best_score:
                best_score, best_candidate, insert_at_end = score_end, node, True

        if best_candidate:
            if insert_at_end:
                path.append(best_candidate)
            else:
                path.insert(0, best_candidate)
            remaining_nodes.remove(best_candidate)
        else:
            # Se ci sono nodi isolati, aggiungili alla fine
            path.extend(list(remaining_nodes))
            break
            
    return [communities[i] for i in path]

def _perform_one_renumbering_attempt(passages, distance, locked_ids, start_number, optimize):
    """Esegue un singolo tentativo completo di rinumerazione."""
    id_map = {}
    passage_dict = {p['original_id']: p for p in passages}
    available_new_ids = list(range(start_number, start_number + len(passages)))
    
    links_graph = {}
    all_passage_ids = set(p['original_id'] for p in passages)
    G = nx.Graph()
    G.add_nodes_from(all_passage_ids)
    for p in passages:
        links_graph[p['original_id']] = re.findall(r'\[\[.*?(\d+).*?\]\]', p['content'])
        for dest_id in links_graph[p['original_id']]:
            if dest_id in all_passage_ids:
                G.add_edge(p['original_id'], dest_id)

    for locked_id in locked_ids:
        if locked_id in passage_dict and available_new_ids:
            new_id = available_new_ids.pop(0)
            id_map[locked_id] = new_id
            print(f"Capitolo bloccato '{locked_id}' -> assegnato al nuovo ID: {new_id}")

    passages_to_renumber = [p for p in passages if p['original_id'] not in locked_ids]

    if optimize:
        print("Modalità ottimizzazione ATTIVA: identificazione e ordinamento delle macro-zone...")
        subgraph_nodes = [p['original_id'] for p in passages_to_renumber]
        subgraph = G.subgraph(subgraph_nodes)
        
        communities = list(community.greedy_modularity_communities(subgraph))
        ordered_zones = _order_zones_intelligently(communities, G)
        print(f"Trovate e ordinate {len(ordered_zones)} macro-zone.")
        
        for i, zone in enumerate(ordered_zones):
            zone_passages = [passage_dict[pid] for pid in zone]
            zone_size = len(zone_passages)
            zone_available_ids = available_new_ids[:zone_size]
            available_new_ids = available_new_ids[zone_size:]
            
            if not zone_available_ids: continue
            print(f"  - Posizionamento Zona {i+1} ({zone_size} capitoli) nel blocco [{zone_available_ids[0]}...{zone_available_ids[-1]}]")
            _place_passages_in_zone(zone_passages, zone_available_ids, distance, id_map, links_graph)
    else:
        print("Modalità ottimizzazione DISATTIVATA: posizionamento casuale semplice.")
        _place_passages_in_zone(passages_to_renumber, available_new_ids, distance, id_map, links_graph)
        
    return id_map, links_graph

def renumber_passages(passages, distance, locked_ids, start_number, optimize, max_dist, attempts):
    """
    Funzione principale che gestisce i tentativi di rinumerazione.
    """
    best_result = None
    best_avg_dist = float('inf')

    for attempt in range(attempts):
        print(f"\n--- Inizio Tentativo {attempt + 1} di {attempts} ---")
        id_map, links_graph = _perform_one_renumbering_attempt(passages, distance, locked_ids, start_number, optimize)
        
        # Calcola le statistiche per questo tentativo
        total_links, total_distance, current_max_dist = 0, 0, 0
        violating_link = None
        for source_id, linked_ids in links_graph.items():
            if source_id not in id_map: continue
            for dest_id in linked_ids:
                if dest_id in id_map:
                    dist = abs(id_map[source_id] - id_map[dest_id])
                    total_distance += dist
                    total_links += 1
                    if dist > current_max_dist:
                        current_max_dist = dist
                        violating_link = f"da '{source_id}' (ora {id_map[source_id]}) a '{dest_id}' (ora {id_map[dest_id]})"
        
        current_avg_dist = (total_distance / total_links) if total_links > 0 else 0
        print(f"Risultato tentativo {attempt + 1}: Distanza media: {current_avg_dist:.2f}, Distanza massima: {current_max_dist}")

        if max_dist is None or current_max_dist <= max_dist:
            print(f"Trovata una soluzione valida che rispetta la distanza massima ({max_dist}).")
            best_result = (id_map, links_graph)
            break # Successo!

        # Salva il risultato migliore trovato finora (quello con la media più bassa)
        if current_avg_dist < best_avg_dist:
            best_avg_dist = current_avg_dist
            best_result = (id_map, links_graph)
            
        if attempt < attempts - 1:
            print(f"La distanza massima {current_max_dist} viola la soglia di {max_dist}. Riprovo...")
    
    if best_result is None:
        print("ERRORE: Nessuna soluzione generata.")
        return []

    if max_dist is not None and current_max_dist > max_dist:
        print(f"\nATTENZIONE: Nessun tentativo è riuscito a rispettare la distanza massima di {max_dist}. Verrà usato il risultato migliore ottenuto.")

    final_id_map, final_links_graph = best_result
    
    # Stampa statistiche finali
    total_links, total_distance = 0, 0
    for source_id, linked_ids in final_links_graph.items():
        if source_id not in final_id_map: continue
        for dest_id in linked_ids:
            if dest_id in final_id_map:
                total_distance += abs(final_id_map[source_id] - final_id_map[dest_id])
                total_links += 1
    if total_links > 0:
        print(f"\n--- Statistica finale: Distanza media dei rimandi: {total_distance / total_links:.2f} ---")

    # Aggiornamento finale del contenuto
    print("Aggiornamento dei link nei capitoli...")
    updated_passages = []
    for p in passages:
        new_id = final_id_map.get(p['original_id'])
        if new_id is None: continue
        original_links = re.findall(r'\[\[([^\]]+)\]\]', p['content'])
        p['original_links_text'] = ", ".join(original_links) if original_links else "Nessuno"
        new_content = p['content']
        for link_text in original_links:
            old_link_id_match = re.search(r'(\d+)', link_text)
            if old_link_id_match and old_link_id_match.group(1) in final_id_map:
                old_id, new_id_val = old_link_id_match.group(1), final_id_map[old_link_id_match.group(1)]
                updated_link_text = re.sub(r'\b' + re.escape(old_id) + r'\b', str(new_id_val), link_text)
                new_content = new_content.replace(f'[[{link_text}]]', f'[[{updated_link_text}]]')
        p['new_id'], p['content'] = new_id, new_content
        updated_passages.append(p)
        
    print("Rinumerazione completata.")
    return updated_passages

def export_to_docx(passages, output_filename, debug_mode=False):
    """
    Esporta i passaggi elaborati in un file .docx.
    """
    print(f"Inizio esportazione nel file DOCX: {output_filename}")
    doc = docx.Document()
    
    sorted_passages = sorted(passages, key=lambda p: p['new_id'])

    for passage in sorted_passages:
        heading = doc.add_heading(f"Capitolo {passage['new_id']}", level=1)
        heading.paragraph_format.keep_with_next = True
        heading.paragraph_format.keep_together = True
        
        p = doc.add_paragraph()
        p.paragraph_format.keep_together = True
        
        content_parts = re.split(r'(\[\[.*?\]\])', passage['content'])
        for part in content_parts:
            if part.startswith('[['):
                link_text = part.strip('[]')
                number_match = re.search(r'(\d+)', link_text)
                if number_match:
                    number = number_match.group(1)
                    before_number, after_number = link_text.split(number, 1)
                    p.add_run(before_number)
                    p.add_run(number).bold = True
                    p.add_run(after_number)
                else:
                    p.add_run(link_text)
            else:
                p.add_run(part)

        if debug_mode:
            debug_p = doc.add_paragraph()
            run = debug_p.add_run(f"(Debug: ID Originale: {passage['original_id']}, Rimandi Originali: [{passage.get('original_links_text', 'Nessuno')}])")
            run.italic = True
            run.font.size = Pt(8)

    try:
        doc.save(output_filename)
        print(f"Successo! File '{output_filename}' creato correttamente.")
    except Exception as e:
        print(f"Errore durante il salvataggio del file DOCX: {e}")

def get_input_file(cli_arg):
    """
    Determina il file di input.
    """
    if cli_arg:
        filename = f"{cli_arg}.twee"
        if not os.path.exists(filename):
            print(f"Errore: Il file specificato '{filename}' non esiste.")
            return None
        return filename
    else:
        print("Nessun file specificato, cerco un file .twee nella directory...")
        twee_files = glob.glob('*.twee')
        if not twee_files:
            print("Errore: Nessun file .twee trovato nella directory.")
            return None
        print(f"Trovato file: {twee_files[0]}")
        return twee_files[0]

# --- ESECUZIONE PRINCIPALE ---
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Processa un file Twee, lo rinumera casualmente e lo esporta in DOCX.", formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('--nomefile', type=str, help="Il nome del file .twee da processare (senza estensione).")
    parser.add_argument('--inizio', type=int, default=1, help="Numero da cui far partire la rinumerazione dei capitoli.\nDefault: 1.")
    parser.add_argument('--distanza', type=int, default=5, help="La distanza MINIMA tra capitoli collegati.\nDefault: 5.")
    parser.add_argument('--distanza-max', type=int, help="La distanza MASSIMA accettabile per un singolo rimando.\nSe un link la supera, lo script ritenterà (vedi --tentativi).")
    parser.add_argument('--tentativi', type=int, default=1, help="Numero di volte che lo script deve tentare di trovare una soluzione\nche rispetti --distanza-max.\nDefault: 1.")
    parser.add_argument('--lock', type=str, default='', help="Lista di ID da bloccare, separati da virgola o spazio.\n(es. '1,21,5' o '1 21 5').")
    parser.add_argument('--ottimizza', action=argparse.BooleanOptionalAction, default=True, help="Attiva (default) o disattiva l'ottimizzazione basata su macro-zone.\nUsa --no-ottimizza per disattivarla.")
    parser.add_argument('--debug', action='store_true', help="Attiva le informazioni di debug nel file DOCX.")
    
    args = parser.parse_args()

    locked_ids = [item.strip() for item in args.lock.replace(',', ' ').split() if item.strip()]

    input_file = get_input_file(args.nomefile)
    
    if input_file:
        output_file = os.path.splitext(input_file)[0] + '.docx'
        raw_passages = parse_twee_file(input_file)

        if raw_passages:
            final_passages = renumber_passages(
                passages=raw_passages, 
                distance=args.distanza, 
                locked_ids=locked_ids, 
                start_number=args.inizio,
                optimize=args.ottimizza,
                max_dist=args.distanza_max,
                attempts=args.tentativi
            )
            export_to_docx(final_passages, output_file, debug_mode=args.debug)
