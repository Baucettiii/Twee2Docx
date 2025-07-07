# -*- coding: utf-8 -*-

import re
import docx
import argparse
import glob
import os
import random
import time
import sys
import math
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

def _print_stats_report(title, id_map, links):
    """Stampa una tabella formattata con le statistiche del layout."""
    avg, max_d, min_d = _calculate_layout_stats(id_map, links)
    print(f"\n--- {title} ---")
    print(f"{'Statistica':<20} | {'Valore':>10}")
    print("-" * 33)
    print(f"{'Distanza Media':<20} | {avg:>10.2f}")
    print(f"{'Distanza Massima':<20} | {max_d:>10}")
    print(f"{'Distanza Minima':<20} | {min_d:>10}")
    print("-" * 33)

def _calculate_layout_stats(id_map, links):
    """
    Calcola le statistiche complete (media, massima, minima) di un layout.
    """
    if not links:
        return 0, 0, 0
        
    distances = []
    for link in links:
        source_id = link['source']
        dest_id = link['dest']
        if source_id in id_map and dest_id in id_map:
            distances.append(abs(id_map[dest_id] - id_map[source_id]))
            
    if not distances:
        return 0, 0, 0

    avg_dist = sum(distances) / len(distances)
    max_dist = max(distances)
    min_dist = min(distances)
    
    return avg_dist, max_dist, min_dist

def _order_zones_intelligently(communities, G):
    """
    Ordina le zone (community) in base alla loro interconnessione.
    """
    if len(communities) <= 1: return communities
    meta_graph = nx.Graph()
    community_map = {node: i for i, com in enumerate(communities) for node in com}
    for u, v in G.edges():
        if u in community_map and v in community_map:
            c1, c2 = community_map[u], community_map[v]
            if c1 != c2:
                if meta_graph.has_edge(c1, c2): meta_graph[c1][c2]['weight'] += 1
                else: meta_graph.add_edge(c1, c2, weight=1)
    if not meta_graph.edges(): random.shuffle(communities); return communities
    
    try:
        start_edge = max(meta_graph.edges(data=True), key=lambda x: x[2]['weight'])
        path = [start_edge[0], start_edge[1]]
    except ValueError: 
        random.shuffle(communities)
        return communities

    remaining_nodes = set(meta_graph.nodes()) - set(path)
    while remaining_nodes:
        best_candidate, best_score, insert_at_end = None, -1, True
        for node in remaining_nodes:
            score_start = meta_graph.get_edge_data(node, path[0], default={'weight': 0})['weight']
            score_end = meta_graph.get_edge_data(node, path[-1], default={'weight': 0})['weight']
            if score_start > best_score: best_score, best_candidate, insert_at_end = score_start, node, False
            if score_end > best_score: best_score, best_candidate, insert_at_end = score_end, node, True
        if best_candidate:
            if insert_at_end: path.append(best_candidate)
            else: path.insert(0, best_candidate)
            remaining_nodes.remove(best_candidate)
        else: path.extend(list(remaining_nodes)); break
    return [communities[i] for i in path]

def _fix_min_dist_violations(id_map, links, non_locked_ids, min_dist, passes):
    """
    Fase di Correzione Robusta: itera per risolvere le violazioni della distanza minima.
    """
    if not passes or min_dist <= 1:
        return id_map

    print(f"\n--- Inizio Fase di Correzione Distanza Minima ({passes} passate) ---")
    
    for p in range(passes):
        violating_links = [link for link in links if abs(id_map.get(link['source'], 0) - id_map.get(link['dest'], 0)) < min_dist]
        
        if not violating_links:
            print(f"  -> Passata {p+1}/{passes}: Nessuna violazione trovata. Correzione completata.")
            break
            
        print(f"  -> Passata {p+1}/{passes}: Trovate {len(violating_links)} violazioni. Tento di risolverle...", end='\r')
        
        corrections_made = 0
        for link in list(violating_links):
            node1, node2 = link['source'], link['dest']
            
            node_to_move = None
            if node1 in non_locked_ids and node2 not in non_locked_ids: node_to_move = node1
            elif node2 in non_locked_ids and node1 not in non_locked_ids: node_to_move = node2
            elif node1 in non_locked_ids and node2 in non_locked_ids: node_to_move = random.choice([node1, node2])
            else: continue 

            for swap_candidate_id in random.sample(non_locked_ids, len(non_locked_ids)):
                if swap_candidate_id == node_to_move: continue

                id_map_after_swap = id_map.copy()
                id_map_after_swap[node_to_move], id_map_after_swap[swap_candidate_id] = id_map[swap_candidate_id], id_map[node_to_move]
                
                is_valid = True
                all_neighbors_of_node_to_move = [l['dest'] for l in links if l['source'] == node_to_move] + [l['source'] for l in links if l['dest'] == node_to_move]
                for neighbor in all_neighbors_of_node_to_move:
                    if neighbor in id_map_after_swap and abs(id_map_after_swap[node_to_move] - id_map_after_swap[neighbor]) < min_dist:
                        is_valid = False; break
                if not is_valid: continue
                
                all_neighbors_of_swap_candidate = [l['dest'] for l in links if l['source'] == swap_candidate_id] + [l['source'] for l in links if l['dest'] == swap_candidate_id]
                for neighbor in all_neighbors_of_swap_candidate:
                    if neighbor in id_map_after_swap and abs(id_map_after_swap[swap_candidate_id] - id_map_after_swap[neighbor]) < min_dist:
                        is_valid = False; break
                
                if is_valid:
                    id_map = id_map_after_swap
                    corrections_made += 1
                    break 
        
        if corrections_made == 0 and p > 0:
            print(f"  -> Passata {p+1}/{passes}: Nessuna ulteriore correzione possibile.                  ", end='\n')
            break
    
    print()
    return id_map

def renumber_passages_hybrid(passages, min_dist, locked_ids, start_number, correction_passes):
    """
    Motore di rinumerazione ibrido definitivo: Bozza strategica + Correzione robusta.
    """
    if not passages: return [], {}

    print("\n--- Inizio Ottimizzazione Ibrida ---")
    
    # FASE 1: Analisi Strutturale e Generazione della Bozza Iniziale OTTIMALE
    print("Fase 1: Analisi della struttura e generazione della bozza iniziale...")
    links, all_passage_ids = [], {p['original_id'] for p in passages}
    G = nx.Graph(); G.add_nodes_from(all_passage_ids)
    for p in passages:
        for dest_id in re.findall(r'\[\[.*?(\d+).*?\]\]', p['content']):
            if dest_id in all_passage_ids:
                links.append({'source': p['original_id'], 'dest': dest_id})
                G.add_edge(p['original_id'], dest_id)
    
    non_locked_ids = [p['original_id'] for p in passages if p['original_id'] not in locked_ids]
    subgraph = G.subgraph(non_locked_ids)
    communities = list(community.greedy_modularity_communities(subgraph))
    ordered_zones = _order_zones_intelligently(communities, G)
    
    initial_id_map = {}
    available_new_ids = list(range(start_number, start_number + len(passages)))
    
    for locked_id in locked_ids:
        if locked_id in G.nodes and available_new_ids:
            initial_id_map[locked_id] = available_new_ids.pop(0)
    
    for zone in ordered_zones:
        zone_nodes = sorted(list(zone), key=lambda x: int(x))
        for node in zone_nodes:
            if available_new_ids:
                initial_id_map[node] = available_new_ids.pop(0)

    initial_stats = _calculate_layout_stats(initial_id_map, links)
    _print_stats_report("Statistiche Bozza Iniziale", initial_id_map, links)
    
    final_id_map = initial_id_map.copy()

    # FASE 2: Correzione Robusta
    final_id_map = _fix_min_dist_violations(final_id_map, links, non_locked_ids, min_dist, correction_passes)
    
    final_stats = _calculate_layout_stats(final_id_map, links)
    stats_summary = {"before": initial_stats, "after": final_stats}

    print("Aggiornamento dei link nei capitoli...")
    updated_passages = []
    for p in passages:
        new_id = final_id_map.get(p['original_id'])
        if new_id is None: continue
        
        original_links_text = re.findall(r'\[\[([^\]]+)\]\]', p['content'])
        p['original_links_text'] = ", ".join(original_links_text) if original_links_text else "Nessuno"

        new_content = p['content']
        for link_text in original_links_text:
            old_link_id_match = re.search(r'(\d+)', link_text)
            if old_link_id_match and old_link_id_match.group(1) in final_id_map:
                old_id = old_link_id_match.group(1)
                new_id_val = final_id_map[old_id]
                updated_link_text = re.sub(r'\b' + re.escape(old_id) + r'\b', str(new_id_val), link_text)
                new_content = new_content.replace(f'[[{link_text}]]', f'[[{updated_link_text}]]')
        
        p['new_id'] = new_id
        p['content'] = new_content
        updated_passages.append(p)
        
    print("Rinumerazione completata.")
    return updated_passages, stats_summary

def export_to_docx(passages, output_filename, debug_mode, script_start_time, stats):
    """Esporta i passaggi elaborati in un file .docx e stampa le statistiche finali."""
    print(f"Inizio esportazione nel file DOCX: {output_filename}")
    doc = docx.Document()
    sorted_passages = sorted(passages, key=lambda p: p['new_id'])
    for passage in sorted_passages:
        heading = doc.add_heading(f"Capitolo {passage['new_id']}", level=1)
        heading.paragraph_format.keep_with_next = True; heading.paragraph_format.keep_together = True
        p = doc.add_paragraph(); p.paragraph_format.keep_together = True
        content_parts = re.split(r'(\[\[.*?\]\])', passage['content'])
        for part in content_parts:
            if part.startswith('[['):
                link_text = part.strip('[]')
                number_match = re.search(r'(\d+)', link_text)
                if number_match:
                    number = number_match.group(1)
                    before_number, after_number = link_text.split(number, 1)
                    p.add_run(before_number); p.add_run(number).bold = True; p.add_run(after_number)
                else: p.add_run(link_text)
            else: p.add_run(part)
        if debug_mode:
            debug_p = doc.add_paragraph()
            run = debug_p.add_run(f"(Debug: ID Originale: {passage['original_id']}, Rimandi Originali: [{passage.get('original_links_text', 'Nessuno')}])")
            run.italic = True; run.font.size = Pt(8)
    try:
        doc.save(output_filename)
        print(f"Successo! File '{output_filename}' creato correttamente.")
        
        before_avg, before_max, before_min = stats['before']
        after_avg, after_max, after_min = stats['after']
        
        print("\n--- Report Statistiche Ottimizzazione (Finale) ---")
        print(f"{'Statistica':<20} | {'Bozza Iniziale':>15} | {'Risultato Finale':>18}")
        print("-" * 60)
        print(f"{'Distanza Media':<20} | {before_avg:>15.2f} | {after_avg:>18.2f}")
        print(f"{'Distanza Massima':<20} | {before_max:>15} | {after_max:>18}")
        print(f"{'Distanza Minima':<20} | {before_min:>15} | {after_min:>18}")
        print("-" * 60)

        script_end_time = time.time()
        execution_time = script_end_time - script_start_time
        print(f"\nTempo di esecuzione totale: {execution_time:.2f} secondi.")
        sys.stdout.flush()
    except Exception as e:
        print(f"Errore durante il salvataggio del file DOCX: {e}")

def get_input_file(cli_arg):
    """Determina il file di input."""
    if cli_arg:
        filename = f"{cli_arg}.twee"
        if not os.path.exists(filename): print(f"Errore: File '{filename}' non trovato."); return None
        return filename
    else:
        print("Nessun file specificato, cerco un file .twee..."); twee_files = glob.glob('*.twee')
        if not twee_files: print("Errore: Nessun file .twee trovato."); return None
        print(f"Trovato file: {twee_files[0]}"); return twee_files[0]

def main():
    """Funzione principale che orchestra l'esecuzione dello script."""
    script_start_time = time.time()
    
    parser = argparse.ArgumentParser(description="Processa un file Twee, lo rinumera e lo esporta.", formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('--nomefile', type=str, help="Nome del file .twee da processare (senza estensione).")
    parser.add_argument('--inizio', type=int, default=1, help="Numero di partenza per la rinumerazione.\nDefault: 1.")
    parser.add_argument('--distanza-min', type=int, default=10, help="Distanza MINIMA tra capitoli collegati.\nDefault: 10.")
    parser.add_argument('--correzione', type=int, default=20, help="Numero di passate per risolvere le violazioni di distanza minima.\nDefault: 20.")
    parser.add_argument('--lock', type=str, default='', help="Lista di ID da bloccare, separati da virgola o spazio.\n(es. '1,21,5' o '1 21 5').")
    parser.add_argument('--debug', action='store_true', help="Attiva le informazioni di debug nel file DOCX.")
    
    args = parser.parse_args()
    locked_ids = [item.strip() for item in args.lock.replace(',', ' ').split() if item.strip()]
    input_file = get_input_file(args.nomefile)
    
    if input_file:
        output_file = os.path.splitext(input_file)[0] + '.docx'
        raw_passages = parse_twee_file(input_file)
        if raw_passages:
            final_passages, stats = renumber_passages_hybrid(
                passages=raw_passages, 
                min_dist=args.distanza_min,
                locked_ids=locked_ids, 
                start_number=args.inizio,
                correction_passes=args.correzione
            )
            if final_passages:
                export_to_docx(final_passages, output_file, args.debug, script_start_time, stats)

# --- ESECUZIONE PRINCIPALE ---
if __name__ == "__main__":
    main()
