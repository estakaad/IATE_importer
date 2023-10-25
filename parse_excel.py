import pandas as pd
import json
from openpyxl import load_workbook


def excel_to_json(excel_filename, concepts_output_filename, sources_output_filename):
    df = pd.read_excel(excel_filename)

    # Prepare a list to store unique references for sources
    unique_refs = set(df['T_TERM_REF1'].dropna().tolist() +
                      df['T_TERM_REF2'].dropna().tolist() +
                      df['L_DEF_REF1'].dropna().tolist() +
                      df['L_NOTE_REF1'].dropna().tolist())

    sources_list = [{"source": ref} for ref in unique_refs]

    # Initialize an empty list to store all concepts
    all_concepts = []

    # Group the DataFrame by 'E_ID' and process each group to construct a concept
    for e_id, group in df.groupby('E_ID'):
        concept_dict = {
            "datasetCode": "esttest",
            "domains": [],
            "definitions": [],
            "notes": [],
            "words": [],
            "conceptIds": [e_id]
        }

        for idx, row in group.iterrows():
            domain = {
                "code": row['E_DOMAIN1'],
                "origin": "lenoch"
            }
            if domain not in concept_dict["domains"]:
                concept_dict["domains"].append(domain)

            word = {
                "value": row['T_TERM'],
                "lexemeSourceLinks": [
                    {
                        "value": row['T_TERM_REF1']
                    },
                    {
                        "value": row['T_TERM_REF2']
                    }
                ]
            }
            concept_dict["words"].append(word)

            definition = {
                "value": row['L_DEF'],
                "sourceLinks": [
                    {
                        "value": row['L_DEF_REF1']
                    }
                ]
            }
            concept_dict["definitions"].append(definition)

            note = {
                "value": row['L_NOTE'],
                "sourceLinks": [
                    {
                        "value": row['L_NOTE_REF1']
                    }
                ]
            }
            concept_dict["notes"].append(note)

        all_concepts.append(concept_dict)

    # Writing concept_json to a file
    with open(concepts_output_filename, 'w', encoding='utf-8') as f:
        f.write(json.dumps(all_concepts, indent=4, ensure_ascii=False))

    # Writing sources_json to a file
    with open(sources_output_filename, 'w', encoding='utf-8') as f:
        f.write(json.dumps(sources_list, indent=4, ensure_ascii=False))
