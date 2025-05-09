import pandas as pd
import matplotlib.pyplot as plt

# Path to xlsx file
file_path = 'expense_may.xlsx'

# Creating the DataFrame for data extraction
try:
    df = pd.read_excel(file_path, header=3)
except FileNotFoundError:
    print(f"File {file_path} not found.")
    exit(1)
except Exception as e:
    print(f"Error loading excel file: {e}")
    exit(1)

# constants (relevant columns)
SUM_OF_TRANSACTION = "סכום\nחיוב"
BRANCH_OF_TRANSACTION = "ענף"

# Branch names in Hebrew
BRANCHES = [
    "מזון ומשקאות",
    "מסעדות",
    "תקשורת ומחשבים",
    "ריהוט ובית",
    "אנרגיה",
    "פנאי ובילוי",
    "אופנה",
    "ציוד ומשרד"
]
# Branch names in English
BRANCHES_EN = [
    "Food and Beverages",
    "Restaurants",
    "Communication and Computers",
    "Furniture and Home",
    "Energy",
    "Leisure and Entertainment",
    "Fashion",
    "Equipment and Office"
]

# Creating the dictionary
branch_translation_map = dict(zip(BRANCHES, BRANCHES_EN))
# Replacing the Hebrew branch names with English ones
df[BRANCH_OF_TRANSACTION] = df[BRANCH_OF_TRANSACTION].map(branch_translation_map)

# Creating a DataFrame for the relevant columns
df = df[[SUM_OF_TRANSACTION, BRANCH_OF_TRANSACTION]]
# Renaming the columns
df.columns = ['sum', 'branch']
# Dropping rows with NaN values
df = df.dropna(subset=['sum'])
# Dropping rows with empty strings
df = df[df['sum'] != '']

# Converting the 'sum' column to numeric
df['sum'] = pd.to_numeric(df['sum'], errors='coerce')
# Enpty branch to "others"
df['branch'] = df['branch'].fillna('others')
# Grouping the DataFrame by branch and summing the 'sum' column
df_grouped = df.groupby('branch')['sum'].sum().reset_index()
# Sorting the DataFrame by 'sum' in descending order
df_grouped = df_grouped.sort_values(by='sum', ascending=False)
print(df_grouped)

# Make plot
if not df_grouped.empty:
    plt.figure(figsize=(12, 9))  # Adjusted figure size for potentially more labels

    # Use the 'branch' column for labels and 'sum' column for sizes
    patches, texts, autotexts = plt.pie(
        df_grouped['sum'],
        labels=df_grouped['branch'],
        autopct='%1.1f%%',
        startangle=140,
        wedgeprops={'edgecolor': 'white'}
        # pctdistance can be adjusted if percentages overlap labels (e.g., pctdistance=0.85)
    )

    plt.title('Distribution of Totals by Branch', fontsize=16)
    plt.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
    plt.tight_layout() # Adjust plot to ensure everything fits without overlapping

    # Adjust font sizes for readability
    for text in texts:
        text.set_fontsize(10) # Branch labels
    for autotext in autotexts:
        autotext.set_fontsize(9) # Percentage labels
        autotext.set_color('white') # Set percentage text color to white for contrast

    plt.show()
else:
    print("No data available to plot a pie chart.")
