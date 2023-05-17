

def data_preparation_processor(raw_df):


    # Define the mapping from old names to new names
    name_mapping = {
        'Employee Name': 'full_name',
        'Company/Display Name': 'company_name',
        'Job Position/Display Name': 'job_position',
        'Hr Team/Display Name': 'hr_team',
        'Work Location/Display Name': 'location',
        'Job Position - New Nomenclature':'new_job_position',
        'Job Family':'job_family',
        'Image': 'image_path'}

    # fill missing values with empty string
    raw_df.fillna('', inplace=True)
    # Rename the columns using the mapping
    raw_df = raw_df.rename(columns=name_mapping)
    # Replace NaN values with an empty string
    raw_df['job_position'] = raw_df['job_position'].fillna('').astype(str)

    # Sort by location and full name
    raw_df = raw_df.sort_values(by=['location', 'full_name'])

    # Convert job_position column to lowercase and remove leading/trailing spaces

    # Add column for job positions containing 'Partner'
    raw_df['Is_partner'] = raw_df['job_position'].str.contains('partner', case=False).astype(int)

    # Add column for job positions containing 'head'
    raw_df['Is_Head'] = raw_df['job_position'].str.contains('head', case=False).astype(int)

    # Add column for job positions containing 'officer'
    raw_df['Is_officer'] = raw_df['job_position'].str.contains('officer', case=False).astype(int)

    # Add column for job positions containing 'manager'
    raw_df['Is_manager'] = raw_df['job_position'].str.contains('manager', case=False).astype(int)

    # Add column for job positions containing 'team lead' or 'team leader'
    raw_df['Is_team_lead'] = raw_df['job_position'].str.contains('team lead|team leader', case=False,
                                                                 regex=True).astype(int)

    # Add column for job positions containing 'vice president' or 'vp'
    raw_df['Is_vice_president'] = raw_df['job_position'].str.contains('vice president|vp', case=False,
                                                                      regex=True).astype(int)
    # Add column for job positions containing 'head'
    raw_df['Is_acc_exec'] = raw_df['job_position'].str.contains('account executive', case=False).astype(int)

    # Add column for job positions containing 'associate'
    raw_df['Is_associate'] = raw_df['job_position'].str.contains('associate', case=False).astype(int)

    # Add column for job positions containing 'admin' or 'administrator'
    raw_df['Is_admin'] = raw_df['job_position'].str.contains('admin|administrator', case=False, regex=True).astype(int)

    # Add column for job positions containing 'lead'
    raw_df['Is_lead'] = raw_df['job_position'].str.contains('lead', case=False).astype(int)

    # Add column for job positions containing 'senior'
    raw_df['Is_senior'] = raw_df['job_position'].str.contains('senior|Sr.', case=False).astype(int)

    # Add column for job positions containing 'director'
    raw_df['Is_director'] = raw_df['job_position'].str.contains('director', case=False).astype(int)

    # Add column for job positions containing 'coordinator' # for graphic design team
    raw_df['Is_coordinator'] = raw_df['job_position'].str.contains('coordinator', case=False).astype(int)

    # Add column for job positions containing 'designer' # for graphic design team
    raw_df['Is_designer'] = raw_df['job_position'].str.contains('designer', case=False).astype(int)
    # Add column for job positions containing 'junior' # for graphic design team
    raw_df['Is_junior'] = raw_df['job_position'].str.contains('junior', case=False).astype(int)

    # Add column for job positions containing 'analyst'
    raw_df['Is_analyst'] = raw_df['job_position'].str.contains('analyst', case=False).astype(int)

    raw_df['Is_COO'] = raw_df['job_position'].str.contains('COO', case=False).astype(int)

    # Add column for job positions containing 'cleaning lady'
    raw_df['Is_cleaning'] = raw_df['job_position'].str.contains('cleaning', case=False).astype(int)



    # Print number of rows where Is_cleaning is 1
    num_cleaning_rows = len(raw_df[raw_df['Is_cleaning'] == 1])
    print(f"There are {num_cleaning_rows} rows with Is_cleaning = 1")

    # Drop rows where Is_cleaning is 1 from raw_df
    raw_df.drop(raw_df[raw_df['Is_cleaning'] == 1].index, inplace=True)

    raw_df.to_excel(r'.\test_data\processed_empl_data.xlsx')
    # create separate dataframes for each HR team
    hr_teams = ["McK Team", "Office Admin team", "IK Team", "BCG ME Team", "Business Translation Team", "IT Team",
                "Finance Team", "BCG Europe Team", "Graphic Design Team", "BCG Africa Team", "Marketing team",
                "BCG KT support", "Operations Team", "Data Analytics Services", "Business Development Team",
                "HR Team", "Research Team", "Executive Management", "IKT/Finance"]

    # create a list of dataframes for each HR team
    team_dfs = [raw_df[raw_df['hr_team'] == team] for team in hr_teams]

    job_families=['Business Research','Office Management','Business Translation','IT','Finance','Graphic Design','Marketing','Operations','Data Analytics','Business Development'
                ,'Human Capital','Executive Management','Sales Representative']
    job_family_dfs = [raw_df[raw_df['job_family'] == family] for family in job_families]

    return team_dfs , job_family_dfs







