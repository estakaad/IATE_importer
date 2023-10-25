import pandas as pd
import json
import re
from openpyxl import load_workbook

def excel_to_json(excel_filename, concepts_output_filename, sources_output_filename, dataset):
    df = pd.read_excel(excel_filename)
    unique_refs = set(df['T_TERM_REF1'].dropna().tolist() +
                      df['T_TERM_REF2'].dropna().tolist() +
                      df['L_DEF_REF1'].dropna().tolist() +
                      df['L_NOTE_REF1'].dropna().tolist())

    sources_list = []
    for ref in unique_refs:
        source = {"type": "DOCUMENT", "name": ref, "description": ref, "isPublic": True}
        if ref.startswith("Ekspert"):
            source["type"] = 'PERSON'
            source["name"] = 'Ekspert'
            source["description"] = ref.replace('Ekspert: ', '')
            source["isPublic"] = False
        sources_list.append(source)

    all_concepts = []
    for e_id, group in df.groupby('E_ID'):
        concept_dict = {
            "datasetCode": dataset,
            "domains": [],
            "definitions": [],
            "notes": [],
            "words": [],
            "conceptIds": [e_id]
        }
        added_definitions = set()  # Track added definitions to avoid duplicates

        for idx, row in group.iterrows():
            domain = {"code": row['E_DOMAIN1'], "origin": "lenoch"}
            if domain not in concept_dict["domains"]:
                concept_dict["domains"].append(domain)

            if pd.notna(row['T_TERM']):
                lexeme_source_links = []

                if pd.notna(row['T_TERM_REF1']):
                    lexeme_source_links.append({"sourceId": None, "value": row['T_TERM_REF1']})
                if pd.notna(row['T_TERM_REF2']):
                    lexeme_source_links.append({"sourceId": None, "value": row['T_TERM_REF2']})

                if row['L_LANG'] == 'en':
                    lang = 'eng'
                elif row['L_LANG'] == 'et':
                    lang = 'est'
                else:
                    lang = row['L_LANG']

                word = {
                    "value": row['T_TERM'],
                    "lang": lang,
                    "lexemeSourceLinks": lexeme_source_links
                }

                concept_dict["words"].append(word)

            definition_value = row['L_DEF']
            if pd.notna(definition_value) and definition_value not in added_definitions:
                definition = {
                    "value": definition_value,
                    "lang": lang,
                    "sourceLinks": [{"sourceId": None, "value": row['L_DEF_REF1']}]
                }
                concept_dict["definitions"].append(definition)
                added_definitions.add(definition_value)

            if pd.notna(row['L_NOTE']):
                lexeme_note = {"value": row['L_NOTE'], "sourceLinks": [{"sourceId": None, "value": row['L_NOTE_REF1']}]}
                concept_dict["words"].append(lexeme_note)

        all_concepts.append(concept_dict)

    with open(concepts_output_filename, 'w', encoding='utf-8') as f:
        f.write(json.dumps(all_concepts, indent=4, ensure_ascii=False))

    with open(sources_output_filename, 'w', encoding='utf-8') as f:
        f.write(json.dumps(sources_list, indent=4, ensure_ascii=False))

