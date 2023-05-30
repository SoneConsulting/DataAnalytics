import pandas as pd
import numpy as np

# Load the sheets into dataframes
xls_file = '/Users/oskarsoderbom/Sone Consulting AB/Final_scraped_data_analysisv3.xlsx'

raw_data_df = pd.read_excel(xls_file, sheet_name='Raw Data')
skills_df = pd.read_excel(xls_file, sheet_name='Skills De_Duped')

# Convert 'Budget_clean', 'Minimum budget' and 'Maximum budget' to float
raw_data_df['Budget_clean2'] = pd.to_numeric(raw_data_df['Budget_clean2'], errors='coerce')
raw_data_df['Minimum budget'] = pd.to_numeric(raw_data_df['Minimum budget'], errors='coerce')
raw_data_df['Maximum budget'] = pd.to_numeric(raw_data_df['Maximum budget'], errors='coerce')

# Prepare a dictionary to store the skills counts and other statistics
skills_stats = {
    'Skills': [],
    'Number of Projects': [],
    'Number of Fixed Price Projects': [],
    'Number of Hourly Projects': [],
    '% Fixed Price': [],
    '% Hourly': [],
    'Number With budget available': [],
    '% with Budget Available': [],
    'Number Fixed priced': [],
    'Number Hourly Priced': [],
    'Low Fixed Price': [],
    'High Fixed Price': [],
    'Median Fixed Price': [],
    'Hourly Range': [],
    'Number Entry Level': [],
    'Number Intermediate': [],
    'Number Expert': [],
    '% Entry': [],
    '% Intermediate': [],
    '% Expert': []
}

# Loop over each skill in 'Skills De_Duped'
for skill in skills_df['Skills']:  # Only the first 10 skills for testing
    skills_stats['Skills'].append(skill)
    
    # Filter rows with this skill
    skill_rows = raw_data_df[raw_data_df['skills'].str.contains(skill, case=False, na=False, regex=False)]
    
    # Count the total number of occurrences
    total_projects = len(skill_rows)
    skills_stats['Number of Projects'].append(total_projects)
    
    # Calculate the most common hourly range
    hourly_rows = skill_rows[skill_rows['JobType_Clean'] == 'Hourly']
    most_common_hourly_range = hourly_rows['Budget range only'].mode().iloc[0] if not hourly_rows['Budget range only'].mode().empty else np.nan
    skills_stats['Hourly Range'].append(most_common_hourly_range)
    
    # Calculate the Low, High and Median Fixed Price
    fixed_price_rows = skill_rows[skill_rows['JobType_Clean'].isin(['Fixed', 'Fixed-price'])]
    skills_stats['Low Fixed Price'].append(fixed_price_rows['Minimum budget'].min())
    skills_stats['High Fixed Price'].append(fixed_price_rows['Maximum budget'].max())
    skills_stats['Median Fixed Price'].append(fixed_price_rows['Budget_clean2'].median())
    
    # Count the number of Fixed Price Projects and Hourly Projects
    skills_stats['Number of Fixed Price Projects'].append(len(fixed_price_rows))
    skills_stats['Number of Hourly Projects'].append(len(hourly_rows))
    
    # Calculate % Fixed Price and % Hourly
    skills_stats['% Fixed Price'].append(len(fixed_price_rows) / total_projects * 100 if total_projects != 0 else 0)
    skills_stats['% Hourly'].append(len(hourly_rows) / total_projects * 100 if total_projects != 0 else 0)
    
    # Calculate Number With budget available and % with Budget Available
    # Calculate Number With budget available and % with Budget Available
    with_budget_rows = skill_rows[skill_rows['Minimum budget'] != "#N/A"]
    skills_stats['Number With budget available'].append(len(with_budget_rows))
    skills_stats['% with Budget Available'].append(len(with_budget_rows) / total_projects * 100 if total_projects != 0 else 0)

    # Calculate Number Fixed priced and Number Hourly Priced
    fixed_with_budget_rows = fixed_price_rows[fixed_price_rows['Minimum budget'] != "#N/A"]
    hourly_with_budget_rows = hourly_rows[hourly_rows['Minimum budget'] != "#N/A"]
    skills_stats['Number Fixed priced'].append(len(fixed_with_budget_rows))
    skills_stats['Number Hourly Priced'].append(len(hourly_with_budget_rows))

    
    # Count the number of occurrences in each contractorTier
    skills_stats['Number Entry Level'].append((skill_rows['contractorTier'] == 'Entry level').sum())
    skills_stats['Number Intermediate'].append((skill_rows['contractorTier'] == 'Intermediate').sum())
    skills_stats['Number Expert'].append((skill_rows['contractorTier'] == 'Expert').sum())
    
    # Calculate % Entry, % Intermediate and % Expert
    skills_stats['% Entry'].append((skill_rows['contractorTier'] == 'Entry level').sum() / total_projects * 100 if total_projects != 0 else 0)
    skills_stats['% Intermediate'].append((skill_rows['contractorTier'] == 'Intermediate').sum() / total_projects * 100 if total_projects != 0 else 0)
    skills_stats['% Expert'].append((skill_rows['contractorTier'] == 'Expert').sum() / total_projects * 100 if total_projects != 0 else 0)

# Convert the dictionary to a dataframe
skills_stats_df = pd.DataFrame(skills_stats)

# Write the dataframe to a new Excel file
skills_stats_df.to_excel('output_file7.xlsx', index=False)
