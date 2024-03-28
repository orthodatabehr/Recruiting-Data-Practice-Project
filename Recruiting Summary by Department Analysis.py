import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

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


# ASSIGN DEGREE VALUES BASED ON HIERARCHY FROM LEGEND
# PHD = 1
# MASTERS + JD = 2
# BACHELORS = 3

def degree_rank(input_col):
    """
    This function will assign ranks based on degree per hierarchy defined.
    PhD = 1, Masters (including JD) =2, Bachelors =3, N/A or blank = 4
    """
    output_col = input_col + ' Rank'
    df[output_col] = np.select(
        [
            df[input_col] == 'Bachelors',
            df[input_col] == 'JD',
            df[input_col] == 'Masters',
            df[input_col] == 'PhD'
        ]
        , [
            3,
            2,
            2,
            1
        ],
        default=4
    )


degree_rank('Degree 1')
degree_rank('Degree 2')
degree_rank('Degree 3')
degree_rank('Degree 4')

# ASSIGN A HIGHEST DEGREE EARNED PER CANDIDATE
df['Highest Degree Rank'] = df[['Degree 1 Rank', 'Degree 2 Rank', 'Degree 3 Rank', 'Degree 4 Rank']].min(axis=1)
df['Highest Degree'] = np.select(
    [
        df['Highest Degree Rank'] == 1,
        df['Highest Degree Rank'] == 2,
        df['Highest Degree Rank'] == 3
    ],
    [
        'PhD',
        'Masters (including JD)',
        'Bachelors'
    ],
    default='No Degree'
)

# Assign Highest Stage Rank and Highest Stage so that each candidate will only be given credit for their Furthest
# Application Stage.
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

conditions = [
    df['Highest Stage Rank'] == 4,
    df['Highest Stage Rank'] == 3,
    df['Highest Stage Rank'] == 2,
    df['Highest Stage Rank'] == 1
]

choices = [
    'Offer Sent',
    'In-House Interview',
    'Phone Screen',
    'New Application'
]

df['Highest Stage'] = np.select(conditions, choices)

unique_candidates_df = df.groupby([
    'Candidate ID Number',
    'Highest Degree',
    'Department',
    'Highest Stage',
    'Offer Decision'
], as_index=False, dropna=False)['Highest Stage Rank'].max()


# SUMMARIZE RECRUITING FUNNEL BY DEPARTMENT AND HIGHEST EDUCATION ACHIEVED
def recruiting_funnel(highest_deg, dept):
    """
    This function will create a recruiting funnel summary for each combination of department and highest degree
    achieved.
    Highest Degree Achieved = 'PhD', 'Masters/JD', 'Bachelors'
    Departments = Engineering, Sales, Operations, Product, Finance, IT
    """
    dpt_deg = unique_candidates_df[(unique_candidates_df['Highest Degree'] == highest_deg)
                                   & (unique_candidates_df['Department'] == dept)]
    summ = dpt_deg.groupby('Highest Stage', as_index=False, dropna=False)['Candidate ID Number'].nunique()
    app_col_name = highest_deg + ' Applicants for ' + dept + ' Department'
    summ = summ.rename(columns={'Candidate ID Number': app_col_name})
    # Add accepted offers row to account for applicants who have accepted offers AND have reached "offer sent" stage.
    summ.loc[summ.shape[0]] = ['Offer Accepted',
                               dpt_deg[(dpt_deg['Offer Decision'] == 'Offer Accepted')
                                       & (dpt_deg['Highest Stage Rank'] == 4)]['Candidate ID Number'].nunique()]
    summ = summ.sort_values(by=app_col_name, ascending=False)
    summ = summ.reset_index(drop=True)
    conv_rate_col_name = dept + ' Conversion Rate for ' + highest_deg + ' Applicants'
    for i in range(len(summ) - 1):
        summ.loc[i + 1, conv_rate_col_name] = round((summ.loc[i + 1, app_col_name] / float(
            summ.loc[i, app_col_name]) * 100.0), 0)

    return summ


# Loop through all department and highest degree combinations to provide a recruiting funnel for each.
"""
This current output will create a separate excel file and PDF table for each dept and highest degree combination.
Each Excel file has 2 sheets, first with the recruiting funnel and the second with all candidate information for that 
department and degree combination. 
"""
departments = unique_candidates_df['Department'].unique()
highest_degrees = unique_candidates_df['Highest Degree'].unique()

for dpt in departments:
    for deg in highest_degrees:
        base_data = unique_candidates_df[(unique_candidates_df['Highest Degree'] == deg)
                                         & (unique_candidates_df['Department'] == dpt)]
        output = recruiting_funnel(deg, dpt)
        plt.axis('off')
        table = plt.table(
            cellText=output.values
            , colLabels=output.columns
            , loc='center'
            , cellLoc='center'
        )
        plt.savefig(f'{dpt} {deg} Output.pdf')
        with pd.ExcelWriter(f'{dpt} {deg} Funnel.xlsx') as writer:
            output.to_excel(writer, sheet_name='Recruiting Funnel')
            base_data.to_excel(writer, sheet_name='Base Data')

# For applicants where offers are sent, we need to pair that with the department of the position they were offered
# and check the 'Offer Decision' column to see if they accepted and add this to our output dataframe print(
# stage_summ.shape[0])

# OVERALL RECRUITING FUNNEL SUMMARY
stage_summ = pd.DataFrame(df.groupby("Highest Stage", as_index=False, dropna=False)["Candidate ID Number"].nunique())
stage_summ = stage_summ.rename(columns={'Candidate ID Number': 'Applicants'})
stage_summ.loc[stage_summ.shape[0]] = ['Offer Accepted',
                                       unique_candidates_df[(unique_candidates_df['Offer Decision'] == 'Offer Accepted')
                                                            & (unique_candidates_df['Highest Stage Rank'] == 4)][
                                           'Candidate ID Number'].nunique()]
stage_summ = stage_summ.sort_values(by='Applicants', ascending=False)
stage_summ = stage_summ.reset_index(drop=True)

for index in range(len(stage_summ) - 1):
    stage_summ.loc[index + 1, 'Conversion Rate'] = round((stage_summ.loc[index + 1, 'Applicants'] / float(
        stage_summ.loc[index, 'Applicants']) * 100.0), 0)

stage_summ.to_excel('Final Output.xlsx')
