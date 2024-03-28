import pandas as pd
import numpy as np
from scipy import stats

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

"""
df['Year of Application'].value_counts() show that the 3 years of data included are 2016, 2017, 2018
"""

# FOCUSING ONLY ON APPLICANTS WITH APPLICATION SOURCES OF "CAREER FAIR" OR "CAMPUS EVENTS" AND CANDIDATES THAT HAVE
# REACHED "IN-HOUSE INTERVIEW"

cf_ce = df[(df['Application Source'] == 'Campus Event') | (df['Application Source'] == 'Career Fair')]

# Assign Highest Stage Rank so that each candidate will only be given credit for their Furthest Application Stage.
conditions = [
    cf_ce['Stage'] == 'Offer Sent',
    cf_ce['Stage'] == 'In-House Interview',
    cf_ce['Stage'] == 'Phone Screen',
    cf_ce['Stage'] == 'New Application'
]

choices = [
    4,
    3,
    2,
    1,
]

cf_ce['Highest Stage Rank'] = np.select(conditions, choices)

unique_candidates_cf_ce = cf_ce.groupby(['Candidate ID Number', 'Year of Application'], as_index=False)[
    'Highest Stage Rank'].max()

# Create binary column for In House Interview Status (True/False)
unique_candidates_cf_ce['Reached In-House Interview'] = np.where(unique_candidates_cf_ce['Highest Stage Rank'] > 2,
                                                                 True, False)

# Create 2-year subsets of data and perform group by function to identify UNIQUE candidate counts.
cf_ce_2016_2017 = unique_candidates_cf_ce[(unique_candidates_cf_ce['Year of Application'] != 2018)]
cf_ce_2017_2018 = unique_candidates_cf_ce[(unique_candidates_cf_ce['Year of Application'] != 2016)]
cf_ce_2016_2018 = unique_candidates_cf_ce[(unique_candidates_cf_ce['Year of Application'] != 2017)]

InHouse_2016_2017_matrix = pd.crosstab(cf_ce_2016_2017['Year of Application'],
                                       cf_ce_2016_2017['Reached In-House Interview'])
chi2_1617, p_1617, dof_1617, expected_1617 = stats.chi2_contingency(InHouse_2016_2017_matrix)
# P-value = 0.93

InHouse_2017_2018_matrix = pd.crosstab(cf_ce_2017_2018['Year of Application'],
                                       cf_ce_2017_2018['Reached In-House Interview'])
chi2_1718, p_1718, dof_1718, expected_1718 = stats.chi2_contingency(InHouse_2017_2018_matrix)
# P-value = 0.046

InHouse_2016_2018_matrix = pd.crosstab(cf_ce_2016_2018['Year of Application'],
                                       cf_ce_2016_2018['Reached In-House Interview'])
chi2_1618, p_1618, dof_1618, expected_1618 = stats.chi2_contingency(InHouse_2016_2018_matrix)
# P-value = 0.066

# Print Outputs
nl = "\n"
print(InHouse_2016_2017_matrix, nl)
print(f"Chi2 value= {chi2_1617}{nl}p-value= {p_1617}{nl}Degrees of freedom= {dof_1617}{nl}")

print(InHouse_2017_2018_matrix, nl)
print(f"Chi2 value= {chi2_1718}{nl}p-value= {p_1718}{nl}Degrees of freedom= {dof_1718}{nl}")

print(InHouse_2016_2018_matrix, nl)
print(f"Chi2 value= {chi2_1618}{nl}p-value= {p_1618}{nl}Degrees of freedom= {dof_1618}{nl}")
