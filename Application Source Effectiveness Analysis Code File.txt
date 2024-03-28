import pandas as pd
import numpy as np

# IMPORT APPROPRIATE FILES FROM SAMPLE DATA
# These Excel files are copy and pasted from the data provided from case study.
recruiting_activity_data = pd.read_excel('RecruitingActivityData.xlsx', header=1)
offer_response_data = pd.read_excel('OfferResponseData.xlsx', header=0)
pd.set_option('display.max_columns', None)

# LEFT OUTER JOIN recruiting activity data with offer response data on Candidate ID Number
df = pd.merge(recruiting_activity_data, offer_response_data, on='Candidate ID Number', how='left')

# DATA CLEANING
"""
When trying to explore the data using value counts, the column names appear to have some trailing spaces.
To avoid confusion when calling columns of the dataframe later, we are trimming spaces from column names here. 
"""
# Removing leading and trailing spaces from column names
df = df.rename(columns=lambda x: x.strip())
"""
In checking for data consistency, the below view showed that "In-House Interview" was not named consistently.
recruiting_activity_data['Furthest Recruiting Stage Reached'].value_counts()
"""
df = df.replace('In-house Interview', 'In-House Interview')

# RENAMING COLUMNS FOR FUTURE EASE
new_col_names = {
    "Furthest Recruiting Stage Reached": "Stage",
    "Degree": "Degree 1",
    "Degree.1": "Degree 2",
    "Degree.2": "Degree 3",
    "Degree.3": "Degree 4",
    "School": "School 1",
    "School.1": "School 2",
    "School.2": "School 3",
    "School.3": "School 4",
    "Major": "Major 1",
    "Major.1": "Major 2",
    "Major.2": "Major 3",
    "Major.3": "Major 4"
}

df = df.rename(columns=new_col_names)

# INCLUDE YEAR COLUMN FOR YEAR-OVER-YEAR ANALYSIS
df['Date of Application'] = pd.to_datetime(df['Date of Application'])
df['Year of Application'] = df['Date of Application'].dt.year

# Assign Highest Stage Rank so that each candidate will only be given credit for their Furthest Application Stage.
conditions = [
    df['Stage'] == 'Offer Sent',
    df['Stage'] == 'In-House Interview',
    df['Stage'] == 'Phone Screen',
    df['Stage'] == 'New Application'
]

choices = [
    4,
    3,
    2,
    1,
]

df['Highest Stage Rank'] = np.select(conditions, choices)

# Include a column for number of applications submitted by each Candidate ID.
apps_per_candidate = df['Candidate ID Number'].value_counts()
df = pd.merge(df, apps_per_candidate, on='Candidate ID Number', how='outer')
df = df.rename(columns={'count': 'Number of Applications'})

# Create table with Candidate ID as the Primary Key to provide a Candidate Level information.
unique_candidates_df = df.groupby([
    'Candidate ID Number',
    'Year of Application',
    'Department',
    'Candidate Type',
    'Application Source',
    'Stage',
    'Offer Decision',
    'Number of Applications'
], as_index=False, dropna=False)['Highest Stage Rank'].max()

# Create binary columns of data for ease of comparison
unique_candidates_df['Accepted Offer'] = np.where(unique_candidates_df['Offer Decision'] == 'Offer Accepted', True,
                                                  False)
unique_candidates_df['In-House Interview Reached'] = np.where(unique_candidates_df['Highest Stage Rank'] > 2, True,
                                                              False)
unique_candidates_df['Experienced Candidate'] = np.where(unique_candidates_df['Candidate Type'] == 'Experienced', True,
                                                         False)

# Export to Excel for Tableau Visualization
unique_candidates_df.to_excel('Question 3 Data File for Visualization.xlsx')
