import pyodbc
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# Database connection
db_path = r'C:\DBPATH\accdb'
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=' + db_path + ';'
)
conn = pyodbc.connect(conn_str)

# Fetch data
query = 'SELECT * FROM Automobile'
df = pd.read_sql(query, conn)
conn.close()

# Group by 'Make' and count the number of automobiles
automake_count = df['AutoMake'].value_counts()

# Combine smaller slices into 'Other'
threshold = 0.05 * automake_count.sum()
other = automake_count[automake_count < threshold].sum()
automake_count = automake_count[automake_count >= threshold]
automake_count['Other'] = other

# Create an interactive bar chart
bar_fig = px.bar(automake_count, x=automake_count.index, y=automake_count.values,
                 labels={'x': 'Automobile Make', 'y': 'Number of Automobiles'},
                 title='Automobile Make Report')

# Add annotations to the bar chart
bar_fig.update_traces(text=automake_count.values, textposition='outside')

# Create an interactive pie chart
pie_fig = px.pie(automake_count, values=automake_count.values, names=automake_count.index,
                 title='Automobile Make Distribution')

# Add annotations to the pie chart
pie_fig.update_traces(textinfo='percent+label')

# Show the figures
bar_fig.show()
pie_fig.show()
