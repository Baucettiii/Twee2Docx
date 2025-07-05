# -*- coding: utf-8 -*-

import re
import docx
import argparse
import glob
import os
import random
from docx.shared import Pt

def parse_twee_file(file_path):
    """
    Analizza un file di testo in formato Twee e estrae i passaggi (capitoli).
    Il titolo del passaggio è l'ID stesso, e tutto il testo seguente è il corpo.
    Ignora i blocchi di metadati come StoryTitle e StoryData.
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

def renumber_passages(passages, distance, locked_ids):
    """
    Rinumera i passaggi con una logica casuale, rispettando i vincoli di
    blocco e distanza minima tra capitoli collegati.
    """
    print("Inizio rinumerazione casuale con vincoli...")
    if not passages:
        return []

    # 1. Inizializzazione
    total_passages = len(passages)
    id_map = {}
    available_new_ids = list(range(1, total_passages + 1))

    # 2. Costruzione del grafo dei collegamenti per una facile consultazione
    links_graph = {}
    for p in passages:
        links_graph[p['original_id']] = re.findall(r'\[\[.*?(\d+).*?\]\]', p['content'])

    # 3. Gestione dei capitoli bloccati
    if not locked_ids and passages:
        locked_ids = [passages[0]['original_id']]
        print(f"Nessun capitolo bloccato specificato. Blocco il primo capitolo di default: ID {locked_ids[0]}")
    
    print(f"Capitoli da bloccare: {locked_ids}")

    locked_passages = [p for p in passages if p['original_id'] in locked_ids]
    passages_to_renumber = [p for p in passages if p['original_id'] not in locked_ids]

    # 4. Posizionamento dei capitoli bloccati per primi
    for locked_p in sorted(locked_passages, key=lambda p: [i for i, pa in enumerate(passages) if pa['original_id'] == p['original_id']][0]):
        if available_new_ids:
            new_id = available_new_ids.pop(0)
            id_map[locked_p['original_id']] = new_id
            print(f"Capitolo bloccato '{locked_p['original_id']}' -> assegnato al nuovo ID: {new_id}")

    # 5. Posizionamento dei capitoli non bloccati
    random.shuffle(passages_to_renumber)
    
    for passage_to_place in passages_to_renumber:
        original_id = passage_to_place['original_id']
        placed_neighbor_new_ids = set()
        for dest_id in links_graph.get(original_id, []):
            if dest_id in id_map:
                placed_neighbor_new_ids.add(id_map[dest_id])
        for src_id, linked_ids in links_graph.items():
            if original_id in linked_ids and src_id in id_map:
                placed_neighbor_new_ids.add(id_map[src_id])

        best_fit_id = None
        for i, potential_new_id in enumerate(available_new_ids):
            is_valid = all(abs(potential_new_id - neighbor_id) >= distance for neighbor_id in placed_neighbor_new_ids)
            if is_valid:
                best_fit_id = potential_new_id
                available_new_ids.pop(i)
                break
        
        if best_fit_id is None:
            if not available_new_ids:
                print(f"Errore critico: non ci sono più ID disponibili per il capitolo '{original_id}'.")
                continue
            best_fit_id = available_new_ids.pop(0)
            print(f"Attenzione: non è stato possibile soddisfare il vincolo di distanza per il capitolo '{original_id}'. Assegnato al primo posto disponibile: {best_fit_id}")
            
        id_map[original_id] = best_fit_id

    # 6. Aggiornamento finale del contenuto dei passaggi
    print("Aggiornamento dei link nei capitoli...")
    updated_passages = []
    for p in passages:
        new_id = id_map.get(p['original_id'])
        if new_id is None: continue

        # Salva i link originali per il debug prima di modificarli
        original_links = re.findall(r'\[\[([^\]]+)\]\]', p['content'])
        p['original_links_text'] = ", ".join(original_links) if original_links else "Nessuno"

        new_content = p['content']
        links_in_content = re.findall(r'\[\[([^\]]+)\]\]', new_content)
        for link_text in links_in_content:
            old_link_id_match = re.search(r'(\d+)', link_text)
            if old_link_id_match:
                old_link_id = old_link_id_match.group(1)
                if old_link_id in id_map:
                    new_link_id = id_map[old_link_id]
                    updated_link_text = re.sub(r'\b' + re.escape(old_link_id) + r'\b', str(new_link_id), link_text)
                    new_content = new_content.replace(f'[[{link_text}]]', f'[[{updated_link_text}]]')

        p['new_id'] = new_id
        p['content'] = new_content
        updated_passages.append(p)
        
    print("Rinumerazione completata.")
    return updated_passages

def export_to_docx(passages, output_filename, debug_mode=False):
    """
    Esporta i passaggi elaborati in un file .docx.
    Se debug_mode è True, aggiunge informazioni di debug dopo ogni capitolo.
    """
    print(f"Inizio esportazione nel file DOCX: {output_filename}")
    doc = docx.Document()
    
    sorted_passages = sorted(passages, key=lambda p: p['new_id'])

    for passage in sorted_passages:
        # Aggiunge il titolo e imposta la formattazione per non dividerlo dal paragrafo successivo
        heading = doc.add_heading(f"Capitolo {passage['new_id']}", level=1)
        heading_format = heading.paragraph_format
        heading_format.keep_with_next = True
        heading_format.keep_together = True
        
        # Aggiunge il contenuto del capitolo
        content_parts = re.split(r'(\[\[.*?\]\])', passage['content'])
        p = doc.add_paragraph()
        p.paragraph_format.keep_together = True
        
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

        # Se la modalità debug è attiva, aggiunge le informazioni
        if debug_mode:
            debug_p = doc.add_paragraph()
            original_links_info = passage.get('original_links_text', 'Nessuno')
            debug_text = f"(Debug: ID Originale: {passage['original_id']}, Rimandi Originali: [{original_links_info}])"
            
            run = debug_p.add_run(debug_text)
            run.italic = True
            font = run.font
            font.size = Pt(8) # Imposta la dimensione del font a 8pt

        # MODIFICA: La riga che aggiungeva un paragrafo vuoto è stata rimossa.
        # doc.add_paragraph('')

    try:
        doc.save(output_filename)
        print(f"Successo! File '{output_filename}' creato correttamente.")
    except Exception as e:
        print(f"Errore durante il salvataggio del file DOCX: {e}")

def get_input_file(cli_arg):
    """
    Determina il file di input. Se specificato, usa quello.
    Altrimenti, cerca il primo file .twee nella directory.
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
    parser = argparse.ArgumentParser(description="Processa un file Twee, lo rinumera casualmente e lo esporta in DOCX.")
    parser.add_argument('--nomefile', type=str, help="Il nome del file .twee da processare (senza estensione).")
    parser.add_argument('--distanza', type=int, default=5, help="La distanza minima tra capitoli collegati (default: 5).")
    parser.add_argument('--lock', type=str, default='', help="Lista di ID da bloccare, separati da virgola o spazio (es. '1,21,5' o '1 21 5').")
    parser.add_argument('--debug', action='store_true', help="Attiva le informazioni di debug nel file DOCX.")
    args = parser.parse_args()

    locked_ids = []
    if args.lock:
        id_string = args.lock.replace(',', ' ')
        locked_ids = [item.strip() for item in id_string.split() if item.strip()]

    input_file = get_input_file(args.nomefile)
    
    if input_file:
        output_file = os.path.splitext(input_file)[0] + '.docx'
        raw_passages = parse_twee_file(input_file)

        if raw_passages:
            final_passages = renumber_passages(raw_passages, args.distanza, locked_ids)
            export_to_docx(final_passages, output_file, debug_mode=args.debug)
