import pandas as pd
#loading the dataset
df = pd.read_excel("uncleaned bike sales dataset.xlsx", sheet_name="Bike Sales")

#removing extra spaces
df.columns= df.columns.str.strip()
df['Country'] = df['Country'].str.strip()

#removing middle extra spaces
df['Country'] = df['Country'].replace({"United  States": "United States"})

#replacing gender values
df['Customer_Gender'] = df['Customer_Gender'].replace({"M": "Male"})
df['Customer_Gender'] = df['Customer_Gender'].replace({"F": "Female"})

# handling with missing values
    #Changing 0 as Null value in Unit_Cost
df.loc[df['Unit_Cost'] == 0, 'Unit_Cost'] = pd.NA
    #Finding the median value for every product_Description group for Unit_Cost
Unit_cost_median = df.groupby('Product_Description')['Unit_Cost'].median()
    #Replacing Null values by median values in Unit_Cost
df['Unit_Cost']=df['Unit_Cost'].fillna(df['Product_Description'].map(Unit_cost_median))

    #Changing 0 as Null value in Unit_Cost
df.loc[df['Unit_Price'] == 0, 'Unit_Price'] = pd.NA
    #Finding the median value for every product_Description group for Unit_Price
Unit_price_median = df.groupby('Product_Description')['Unit_Price'].median()
    #Replacing Null values by median values in Unit_Price
df['Unit_Price']=df['Unit_Price'].fillna(df['Product_Description'].map(Unit_price_median))
    #Fixing Cost, Revenue and Profit
df['Cost']=df['Order_Quantity']*df['Unit_Cost']
df['Revenue']=df['Order_Quantity']*df['Unit_Price']
df['Profit']=df['Revenue']-df['Cost']
#Removing Duplicates
df=df.drop_duplicates()
#Exporting as Excel File
Filename="Cleaned_Bike_Sales_Dataset.xlsx"
df.to_excel(Filename,index=False,engine="openpyxl")

