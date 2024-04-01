import pandas as pd

# Load the Excel file into a pandas DataFrame
file_path = '/Users/spencer/Desktop/Roll Call/Actual Data/Excel Files/118_Session_2_All_Data.xlsx'
df = pd.read_excel(file_path)

# Just to verify, we'll print out the column names to ensure we have the correct column identifier for "Member Name"
df.columns.tolist()

# Define the list of names we are interested in
members_to_filter = [
    'Mast', 'Johnson (LA)', 'McCarthy', 'Scalise', 'Emmer', 'Stefanik',
    'Diaz-Balart', 'Graves (MO)', 'McCaul', 'Buchanan', 'Wagner (MO)',
    'Issa (CA)', 'Perry (PA)', 'Wilson (SC)', 'Smith (NJ)'
]

# Filter the DataFrame for the names in your list
filtered_df = df[df['Member Name'].isin(members_to_filter)]

# Pivot the table so that each roll call is a single row, and each member's name is a column
# We use a lambda function to handle potential issues with duplicate entries by joining votes with a comma if they exist
pivoted_df = filtered_df.pivot_table(
    index='Roll Call Number', 
    columns='Member Name', 
    values='Vote', 
    aggfunc=lambda x: ', '.join(str(v) for v in x)
)

# Reset index to make 'Roll Call Number' a column again if needed
pivoted_df.reset_index(inplace=True)

# Write the pivoted DataFrame to a new Excel file
output_path = '/Users/spencer/Desktop/Roll Call/Actual Data/Excel Files/118_Session_2_Pivoted.xlsx'
pivoted_df.to_excel(output_path, index=False)

output_path
