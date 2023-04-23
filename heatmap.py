import datetime
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# Assuming your data is stored in a DataFrame called "df"
# You can specify the index and column names if needed
# Here, we assume the index is the "name" column and the columns are the keywords
# Load the Excel file and select the relevant data
# name = input("Enter the filename from \"outputs\" folder: ")
df = pd.read_excel('keyword_count_sample.xlsx')
# df = df.drop(df.columns[[0, 1]], axis=1, inplace=True)
df = df.drop(columns=['Technology areas'])
df = df.set_index('Technologies')
# df = df.transpose()

plt.figure(figsize=(15, 10))

# "coolwarm"
# "viridis"
# "magma"
# "inferno"
# "plasma"
# "PuBu"
# "GnBu"
# "BuGn"
# "YlGnBu"
# "YlOrRd"
# "RdYlBu"
# "RdYlGn"
# "gist_earth"
# "terrain"
# "ocean"
# "gist_stern"

# cmap1 = 'YlGnBu'
# cmap2 = sns.color_palette("Blues")
# Generate the heatmap using Seaborn
sns.heatmap(df, annot=False, fmt="d", cmap='YlOrRd',
            linewidths=.5, linecolor='white', cbar=True, square=False)
# ax = sns.heatmap(df, cmap='YlGnBu'),
# ax.set(xlabel="", ylabel="")
# ax.xaxis.tick_top()

# # Add labels and title
plt.title('Keyword counts per publisher')
# plt.ylabel('Publishers or Filenames')
# plt.xlabel('Keywords')
plt.tight_layout()

# # Show the plot
plt.show()

# now = datetime.datetime.now()
# name = 'heatmap_'+now.strftime("%y%m%d%H%M%S")+'.pdf'
# Save the figure to a PDF file
# plt.savefig('outputs/heatmaps/'+name)

# Save the figure to an Excel file (requires the XlsxWriter library)
# writer = pd.ExcelWriter('heatmap.xlsx', engine='xlsxwriter')
# df.to_excel(writer, sheet_name='Sheet1')
# workbook = writer.book
# worksheet = writer.sheets['Sheet1']
# chart = workbook.add_chart({'type': 'heatmap'})
# chart.add_series({
#     'values': '=Sheet1!$B$2:$E$5',
# })
# worksheet.insert_chart('G2', chart)
# writer.save()
