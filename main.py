import parse_excel
import concepts_requests

iate_excel = 'files/input/Kalandus_28-2-2023.xlsx'
iate_json_concepts = 'files/output/concepts.json'
iate_json_sources_without_ids = 'files/output/sources.json'
iate_json_sources_with_ids = ''
iate_json_concepts_with_word_ids_and_sources_ids = ''
iate_ids_of_added_concepts = ''

# Parse Excel
parse_excel.excel_to_json(iate_excel, iate_json_concepts, iate_json_sources_without_ids)

# Import sources and add their ID-s to sources

#iate_json_sources_with_ids = concepts_requests.import_sources(iate_json_sources_without_ids)

# Get word ID-s and add them to concepts and add IDs of sources also to concepts

#iate_json_concepts_with_word_ids_and_sources_ids

# Import concepts