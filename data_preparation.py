import pandas as pd
from constants import NAME_MAPPING, HR_TEAMS, JOB_FAMILIES

def data_preparation_processor(raw_df):
    # Rename the columns using the mapping
    raw_df = raw_df.rename(columns=NAME_MAPPING)

    # Fill missing values with an empty string
    raw_df.fillna('', inplace=True)

    # Replace NaN values with an empty string
    raw_df['job_position'] = raw_df['job_position'].fillna('').astype(str)

    # Sort by location and full name
    raw_df = raw_df.sort_values(by=['location', 'full_name'])

    # Convert job_position column to lowercase and remove leading/trailing spaces

    # Add columns for different job position patterns
    raw_df['Is_partner'] = raw_df['job_position'].str.contains('partner', case=False).astype(int)
    raw_df['Is_Head'] = raw_df['job_position'].str.contains('head', case=False).astype(int)
    raw_df['Is_officer'] = raw_df['job_position'].str.contains('officer', case=False).astype(int)
    raw_df['Is_manager'] = raw_df['job_position'].str.contains('manager', case=False).astype(int)
    raw_df['Is_team_lead'] = raw_df['job_position'].str.contains('team lead|team leader', case=False, regex=True).astype(int)
    raw_df['Is_vice_president'] = raw_df['job_position'].str.contains('vice president|vp', case=False, regex=True).astype(int)
    raw_df['Is_acc_exec'] = raw_df['job_position'].str.contains('account executive', case=False).astype(int)
    raw_df['Is_associate'] = raw_df['job_position'].str.contains('associate', case=False).astype(int)
    raw_df['Is_admin'] = raw_df['job_position'].str.contains('admin|administrator', case=False, regex=True).astype(int)
    raw_df['Is_lead'] = raw_df['job_position'].str.contains('lead', case=False).astype(int)
    raw_df['Is_senior'] = raw_df['job_position'].str.contains('senior|Sr.', case=False).astype(int)
    raw_df['Is_director'] = raw_df['job_position'].str.contains('director', case=False).astype(int)
    raw_df['Is_coordinator'] = raw_df['job_position'].str.contains('coordinator', case=False).astype(int)
    raw_df['Is_designer'] = raw_df['job_position'].str.contains('designer', case=False).astype(int)
    raw_df['Is_junior'] = raw_df['job_position'].str.contains('junior', case=False).astype(int)
    raw_df['Is_analyst'] = raw_df['job_position'].str.contains('analyst', case=False).astype(int)
    raw_df['Is_COO'] = raw_df['job_position'].str.contains('COO', case=False).astype(int)
    raw_df['Is_cleaning'] = raw_df['job_position'].str.contains('cleaning', case=False).astype(int)

    # Drop rows where Is_cleaning is 1 from raw_df
    raw_df.drop(raw_df[raw_df['Is_cleaning'] == 1].index, inplace=True)

    raw_df.to_excel(r'.\test_data\processed_empl_data.xlsx')

    # Create separate dataframes for each HR team
    team_dfs = [raw_df[raw_df['hr_team'] == team] for team in HR_TEAMS]

    # Create separate dataframes for each job family
    job_family_dfs = [raw_df[raw_df['job_family'] == family] for family in JOB_FAMILIES]

    return team_dfs, job_family_dfs








