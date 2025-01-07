import pandas as pd

# Function to allocate SPOCs based on the round-robin logic with shifting patterns
def allocate_spocs(file_path):
    # Load the Excel file
    data = pd.read_excel(file_path, sheet_name='Sheet1')

    # Team members
    team = ['A Team', 'B Team', 'C Team', 'D Team']

    # Initialize variables
    spoc_allocation = []
    team_size = len(team)
    start_shift = 0  # Variable to shift the starting index for each new company

    # Dictionary to track the last assigned member for each company
    last_assigned = {}

    for company in data['Company']:
        if company not in last_assigned:
            # Shift the starting pattern for new companies
            start_index = start_shift % team_size
            start_shift += 1
        else:
            # Otherwise, continue from the next member
            start_index = (last_assigned[company] + 1) % team_size

        # Assign the next SPOC
        spoc = team[start_index]
        spoc_allocation.append(spoc)

        # Update the last assigned index for the company
        last_assigned[company] = start_index

    # Add the SPOC column to the dataframe
    data['SPOC'] = spoc_allocation

    # Save the updated dataframe back to an Excel file
    output_path = file_path.replace('.xlsx', '_updated.xlsx')
    data.to_excel(output_path, index=False)
    return output_path

# Main script
if __name__ == "__main__":
    file_path = input("Enter the path of the Excel file: ").strip()
    try:
        output_path = allocate_spocs(file_path)
        print(f"SPOC allocation completed. Updated file saved at: {output_path}")
    except Exception as e:
        print(f"An error occurred: {e}")
