"""
tutorial (link: https://towardsdatascience.com/stylize-and-automate-your-excel-files-with-python-cab49b57b25a)
"""

import pandas as pd
import xlwings as xw
import matplotlib.pyplot as plt
 
FOLDER_PATH = r"path_to_save_folder" # e.g. r"C:\Users\Name\Downloads"

# Import CSV file using one of the two methods below
#df = pd.read_csv(r"path_to_csv\fruit_and_veg_sales.csv")
df = pd.read_csv(r"https://raw.githubusercontent.com/Nishan-Pradhan/xlwings_dashboard/master/fruit_and_veg_sales.csv")

# Explained in previous tutorial (link: https://towardsdatascience.com/stylize-and-automate-your-excel-files-with-python-cab49b57b25a)
wb = xw.Book()
sht = wb.sheets["Sheet1"]
sht.name = "fruit_and_veg_sales"
sht.range("A1").options(index=False).value = df
wb.sheets.add('Dashboard')
sht_dashboard = wb.sheets('Dashboard')

# ===========================# Data Manipulation # ==========================

# Check column names
print(df.columns)

# Pivot by Item and show Total Profit
pv_total_profit = pd.pivot_table(df,index='Item',values='Total Profit ($)',aggfunc='sum')

# Pivot data by Item and show Quantity Sold
pv_quantity_sold = pd.pivot_table(df,index='Item',values='Quantity Sold',aggfunc='sum')

# Correct Date Sold data type
print(df.dtypes)
df["Date Sold"] = pd.to_datetime(df["Date Sold"], format='%d/%m/%Y')

# Group by Date Sold in months
gb_date_sold = df.groupby(df["Date Sold"].dt.to_period('m')).sum()[["Quantity Sold",'Total Revenue ($)',
       'Total Cost ($)',"Total Profit ($)"]]
gb_date_sold.index = gb_date_sold.index.to_series().astype(str)

# Group by Date Sold, sort by Total Revenue, show top 8 rows
gb_top_revenue = (df.groupby(df["Date Sold"])
 .sum()
 .sort_values('Total Revenue ($)',ascending=False)
 .head(8)
 )[["Quantity Sold",'Total Revenue ($)',
       'Total Cost ($)',"Total Profit ($)"]]

#  ============================ # Format Dashboard # =========================

#background
sht_dashboard.range('A1:Z1000').color = (198,224,180)

#A:B column width
sht_dashboard.range('A:B').column_width = 2.22

#title
sht_dashboard.range('B2').value = 'Sales Dashboard'
sht_dashboard.range('B2').api.Font.Name = 'Arial'
sht_dashboard.range('B2').api.Font.Size = 48
sht_dashboard.range('B2').api.Font.Bold = True
sht_dashboard.range('B2').api.Font.Color = 0x000000
sht_dashboard.range('B2').row_height = 61.2

# Underline Title
sht_dashboard.range('B2:W2').api.Borders(9).Weight = 4
sht_dashboard.range('B2:W2').api.Borders(9).Color = 0x00B050

# Subtitle
sht_dashboard.range('M2').value = 'Total Profit Per Item Chart'
sht_dashboard.range('M2').api.Font.Name = 'Arial'
sht_dashboard.range('M2').api.Font.Size = 20
sht_dashboard.range('M2').api.Font.Bold = True
sht_dashboard.range('M2').api.Font.Color = 0x000000

# Line dividing Title and Subtitle
sht_dashboard.range('L2').api.Borders(7).Weight = 3
sht_dashboard.range('L2').api.Borders(7).Color = 0x00B050
sht_dashboard.range('L2').api.Borders(7).LineStyle = -4115

# Function which formats a dataframe.
def create_formatted_summary(header_cell,title,df_summary,color):
    """
    

    Parameters
    ----------
    header_cell : Str
        Provide top left cell location where you want to place a dataframe. e.g. 'B2'
        
    title : Str
        Specify what title you want this block to have. e.g. 'Pivot Title'
        
    df_summary : DataFrame
        Provide the DataFrame you wish to place in Excel. 
        
    color : Str
        Provide name of a color. e.g. 'blue' etc.
        Check the function for dictionary of colors (colors). 
        More colors can be added to this dictionary, 
        just add RGB tuples of a lighter and darker shade of the same colour
        

    Returns
    -------
    None. This function just formats Excel.

    """
    
    # Dictionary of colors, [(Darker color),(Lighter color)]
    colors = {"purple":[(112,48,160),(161,98,208)],
              "blue":[(0,112,192),(155,194,230)],
              "green":[(0,176,80),(169,208,142)],
              "yellow":[(255,192,0),(255,217,102)]}
    
    # Set column width of first summary column
    sht_dashboard.range(header_cell).column_width = 1.5
    
    # Assign row and column location to variables
    row, col = sht_dashboard.range(header_cell).row, sht_dashboard.range(header_cell).column
    
    # Format title of DataFrame Summary
    summary_title_range = sht_dashboard.range(row,col)
    summary_title_range.value = title
    summary_title_range.api.Font.Size = 14
    summary_title_range.row_height = 32.5
    summary_title_range.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
    summary_title_range.api.Font.Color = 0xFFFFFF
    summary_title_range.api.Font.Bold = True
    sht_dashboard.range((row,col),
                        (row,col+len(df_summary.columns)+1)).color = colors[color][0] #Darker color
    
    # Format headers of DataFrame Summary 
    summary_header_range = sht_dashboard.range(row+1,col+1)
    summary_header_range.value = df_summary
    summary_header_range = summary_header_range.expand('right')
    summary_header_range.api.Font.Size = 11
    summary_header_range.api.Font.Bold = True
    sht_dashboard.range((row+1,col),
                        (row+1,col+len(df_summary.columns)+1)).color = colors[color][1] #Darker color
    sht_dashboard.range((row+1,col+1),
                        (row+len(df_summary),col+len(df_summary.columns)+1)).autofit()
    
    for num in range(1,len(df_summary)+2,2):
        sht_dashboard.range((row+num,col),
                    (row+num,col+len(df_summary.columns)+1)).color = colors[color][1] 
        
    # Find last row of DataFrame Summary
    last_row = sht_dashboard.range(row+1,col+1).expand('down').last_cell.row
    side_border_range = sht_dashboard.range((row+1,col),(last_row,col))
    
    # Add dashed border with color to left of DataFrame Summary
    sht_dashboard.range(side_border_range).api.Borders(7).Weight = 3
    sht_dashboard.range(side_border_range).api.Borders(7).Color = xw.utils.rgb_to_int(colors[color][1])
    sht_dashboard.range(side_border_range).api.Borders(7).LineStyle = -4115
    
    
# Runs the function and creates each section of our summary sheet
create_formatted_summary('B5','Total Profit per Item', pv_total_profit, 'green')
create_formatted_summary('B17','Total Items Sold', pv_quantity_sold, 'purple')
create_formatted_summary('F17','Sales by Month', gb_date_sold, 'blue')
create_formatted_summary('F5','Top 5 Days by Revenue ', gb_top_revenue, 'yellow')

# Makes a chart using Matplotlib
fig, ax = plt.subplots(figsize=(6,3))
pv_total_profit.plot(color='g',kind='bar',ax=ax)

# Add Chart to Dashboard Sheet
sht_dashboard.pictures.add(fig,name='ItemsChart',
                           left=sht_dashboard.range("M5").left,
                           top=sht_dashboard.range("M5").top,
                           update = True)

#=====================# BONUS - ADD LOGO TO YOUR DASHBOARD #===================

import requests 

# Get image from the web
image_url = r"https://github.com/Nishan-Pradhan/xlwings_dashboard/blob/master/pie_logo.png?raw=true"

# Open the url image.
r = requests.get(image_url, stream = True)

image_path = rf"{FOLDER_PATH}\logo.png"

# Saves image to image_path above
file = open(image_path, "wb")
file.write(r.content)
file.close()

# Adds image to Excel Dashboard
logo = sht_dashboard.pictures.add(image=image_path,
                           name='PC_3',
                           left=sht_dashboard.range("J2").left,
                           top=sht_dashboard.range("J2").top+5,
                           update=True)

# Resizes image
logo.width = 54
logo.height = 54

#==============================================================================

# Save your Excel file
wb.save(rf"{FOLDER_PATH}\fruit_and_veg_dashboard.xlsx")

print("Completed saving: fruit and veg dashboard")

