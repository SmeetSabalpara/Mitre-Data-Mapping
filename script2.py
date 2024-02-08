import pandas as pd

input_file = 'file.xlsx'
df = pd.read_excel(input_file)

pending_rows = []

for index, row in df.iterrows():

    if(',' in row['Tactics']):   	 
        current_list_tactic = str(row['Tactics']).split(',')
        current_list_tactic = [x.strip() for x in current_list_tactic]

        for i in range(len(current_list_tactic)):

            if(i==0):

                df.at[index, 'Tactics'] = current_list_tactic[i]    
                
            else:
                row_copy = row.copy()
                row_copy['Tactics'] = current_list_tactic[i]       	 
                pending_rows.append(row_copy)
       	 
df = df._append(pending_rows,  ignore_index=True)

df.sort_values(by='Tactics', inplace=True)

output_file = 'file.xlsx'
df.to_excel(output_file, index=False)

print(f"Formatted data saved to {output_file}")
