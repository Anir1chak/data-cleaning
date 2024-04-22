import pandas as pd

# Load data from anon_data2.xlsx
data2 = pd.read_excel("anon_data2.xlsx")

# Keep only the specified columns
columns_to_keep = ['Team Name', 'Candidate Type', "Candidate's Name", "Candidate's Email", "Candidate's Mobile"]
data2_filtered = data2[columns_to_keep]

# Rename the columns to remove special characters
data2_filtered.columns = ['Team Name', 'Candidate Type', 'Candidate Name', 'Candidate Email', 'Candidate Mobile']

# Save the filtered data to a new Excel file named output.xlsx
data2_filtered.to_excel("output.xlsx", index=False)
#import pandas as pd

# Load data from anon_data1.xlsx
data1 = pd.read_excel("./anon_data1.xlsx")

# Keep only the specified columns
columns_to_keep = ['Team Name', 'Leader Name', 'Leader Email', 'Leader Phone', 'Members']
data1_filtered = data1[columns_to_keep]

# Rename the columns as requested
data1_filtered.columns = ['Team Name', 'Candidate Name', 'Candidate Email', 'Candidate Mobile', 'Members']

# Add a new column for Candidate Type with value "leader" for all entries
data1_filtered['Candidate Type'] = 'leader'

# Save the modified data to a new Excel file named output2.xlsx
data1_filtered.to_excel("output2.xlsx", index=False)


# Load data from output2.xlsx
data_output2 = pd.read_excel("./output2.xlsx")

# Function to extract additional emails from the Members column
def extract_emails(row):
    emails = row['Members'].split(',')
    if len(emails) <= 1:
        return []
    return emails[1:]  # Exclude the first email

# Apply the function to extract additional emails
data_output2['Additional Emails'] = data_output2.apply(extract_emails, axis=1)

# Create a new dataframe for output3.xlsx
output3_data = pd.DataFrame(columns=['Team Name', 'Candidate Name', 'Candidate Email', 'Candidate Mobile', 'Candidate Type'])

# Iterate through each row to generate entries for output3_data
for index, row in data_output2.iterrows():
    team_name = row['Team Name']
    candidate_email = row['Candidate Email']
    candidate_type = 'member'  # Candidate Type is 'member' for all entries in output3.xlsx

    # Iterate through additional emails and create entries for output3_data
    for email in row['Additional Emails']:
        output3_data = pd.concat([output3_data, pd.DataFrame({
            'Team Name': [team_name],
            'Candidate Name': ['NA'],
            'Candidate Email': [email.strip()],  # Remove leading/trailing whitespaces
            'Candidate Mobile': ['NA'],  # Use 'NA' for all mobile numbers
            'Candidate Type': [candidate_type]
        })], ignore_index=True)

# Save output3_data to a new Excel file named output3.xlsx
output3_data.to_excel("output3.xlsx", index=False)


# Load data from output.xlsx, output2.xlsx, output3.xlsx
data_output = pd.read_excel("output.xlsx")
data_output2 = pd.read_excel("output2.xlsx")
data_output3 = pd.read_excel("output3.xlsx")

# Delete the 'Members' column from output2.xlsx
data_output2.drop(columns=['Members'], inplace=True)

# Concatenate the three dataframes into one
final_data = pd.concat([data_output, data_output2, data_output3], ignore_index=True)

# Save the final dataframe to a new Excel file named final.xlsx
final_data.to_excel("final.xlsx", index=False)


# Read the Excel file into a pandas DataFrame
df = pd.read_excel('final.xlsx')

# Save the DataFrame to a CSV file
df.to_csv('final.csv', index=False)

print("Conversion successful: final.xlsx -> final.csv")
