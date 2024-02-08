import pandas as pd

input_file = 'mitre.xlsx'
df = pd.read_excel(input_file)

input_file_relationships = pd.read_excel(input_file, "relationships")
input_file_mitigations = pd.read_excel(input_file, "mitigations")

targ_dic = {}
src_dic = {}
subtech_to_detection_relationship = {}
subtech_to_source_name_relationship = {}

for index, row in input_file_relationships.iterrows():
    targ_dic[row['STIX ID']] = row['target ref']
    src_dic[row['STIX ID']] = row['source ref']
    subtech_to_detection_relationship[row['STIX ID']] = row['mapping description']
    subtech_to_source_name_relationship[row['STIX ID']] = row['source name']
    # targ_src_dic[row['target ref']] = row['source ref']
    # subtech_to_detection_relationship[row['target ref']] = row['mapping description']
    # subtech_to_source_name_relationship[row['target ref']] = row['source name']

formatted_df = pd.DataFrame(columns=['Tactics', 'Technique', 'Technique Description', 'Technique Detection', 'Sub-Technique', 'Sub-Technique Description','Sub-Technique Detection','Mitre Detection', 'Mitigation', 'Platforms'])

miti_dic_desc = {}
miti_dic_id = {}
miti_dic_name = {}
miti_desc = " "
miti_id = " "
miti_name = " "

for index, row in input_file_mitigations.iterrows():
    miti_dic_name[row['STIX ID']] = row['name']
    miti_dic_id[row['STIX ID']] = row['ID']
    miti_dic_desc[row['STIX ID']] = row['description']

description = " "
detection = " "

for index, row in df.iterrows():
    technique_id = row['ID']
    technique_name = row['name']
    sub_description = row['description']
    is_sub_technique = row['is sub-technique']  
    sub_detection = row['detection']
    STIX_ID = row['STIX ID']
    platforms = row['platforms']

    # if(str(subtech_to_detection_relationship[STIX_ID]) == ""):
    #     subtech_to_detection_relationship[STIX_ID] = " "

    list_STIX_Relationship = []
    for key, value in targ_dic.items():
        if(value == STIX_ID):
            list_STIX_Relationship.append(key)

    list_Src = []
    list_detc_details = []
    for i in range(len(list_STIX_Relationship)):
        list_Src.append(src_dic[list_STIX_Relationship[i]])
        list_detc_details.append(subtech_to_source_name_relationship[list_STIX_Relationship[i]] + "\n" + subtech_to_detection_relationship[list_STIX_Relationship[i]] +"\n\n")
    #     subtech_to_detection_relationship[row['STIX ID']] = row['mapping description']
    # subtech_to_source_name_relationship[row['STIX ID']] = row['source name']

    list_miti_details = []
    for i in range(len(list_Src)):
        if list_Src[i] in miti_dic_id:
            list_miti_details.append(miti_dic_id[list_Src[i]] + "\n"+miti_dic_name[list_Src[i]] + "\n" + miti_dic_desc[list_Src[i]] +"\n\n")

    to_str_dec = '\n'.join(list_detc_details)
    to_str_miti = '\n'.join(list_miti_details)
    

    # if(targ_src_dic[STIX_ID] + 'id' in miti_dic_id):
    #     miti_desc = miti_dic_desc[targ_src_dic[STIX_ID] + 'description']
    #     miti_id = miti_dic_id[targ_src_dic[STIX_ID] + 'id']
    #     miti_name = miti_dic_name[targ_src_dic[STIX_ID] + 'name']    	 

    if('.' not in technique_id):      	 
        description = row['description']
        detection = row['detection']

    if is_sub_technique:   	 
        
        formatted_df = formatted_df._append({'Tactics': row['tactics'],
        'Technique': technique_name.split(':')[0] + "\n"+ technique_id.split('.')[0],                  	 
        'Technique Description': description,
        'Technique Detection': detection,
        'Sub-Technique': technique_name.split(':')[-1] + "\n" +technique_id,
        'Sub-Technique Description': sub_description,
        'Sub-Technique Detection': sub_detection,        
        'Mitre Detection': to_str_dec,
        'Mitigation': to_str_miti,
        'Platforms': platforms},
        ignore_index=True)
    else:
        formatted_df = formatted_df._append({'Tactics': row['tactics'],
        'Technique': technique_name.split(':')[0] + "\n"+ technique_id.split('.')[0],                  	 
        'Technique Description': description,
        'Technique Detection': detection,
        'Sub-Technique': ' ',
        'Sub-Technique Description': ' ',
        'Sub-Technique Detection': ' ',
        'Mitre Detection': to_str_dec,
        'Mitigation': to_str_miti,
        'Platforms': platforms},
        ignore_index=True)                  


formatted_df.sort_values(by='Tactics', inplace=True)

output_file = 'file.xlsx'
formatted_df.to_excel(output_file, index=False)

print(f"Formatted data saved to {output_file}")
