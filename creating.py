import pandas as pd
import os
import time
from xlrd import XLRDError

start_time = time.time()

# list of paths to ebay files
ebay_files = []

# searching all excel files in the folder
for root, dirs, files in os.walk(r'D:\Projects\shopContent\ebay'):
	ebay_files.extend([os.path.join(root, file) for file in files if file.endswith('.xlsx')])
	dirs.clear()

# creating dataframe
ebay_df = pd.DataFrame()

# appending tables from all source ebay files to one dataframe skipping first 2 rows
print("Creating ebay dataframe!")
for file in ebay_files:
    try:
        ebay_df = ebay_df.append(pd.read_excel(file, sheet_name="Listings", skiprows=2))
    except XLRDError:
        print(f"No sheet named \'Listings\' in file - {file}")

# create dataframe from csv file
print("Creating shopify dataframe!")
shopify_df = pd.read_csv(r'D:\Projects\shopContent\shopify\shopify.csv', sep=',', encoding="utf-8", header=0)

# replace '||' symbols to ', ' in column 'C:Season'
print("Replacing '||' symbols in ebay dataframe!")
ebay_df['C:Season'] = ebay_df['C:Season'].str.replace("\|\|", ', ')

# enable only 'Custom Label (SKU)', 'C:Brand', 'C:Type', 'C:Season' columns in dataframe
print("Excluding columns in ebay dataframe!")
ebay_df = ebay_df[['Custom Label (SKU)', 'C:Brand', 'C:Type', 'C:Season']]

# export ebay_df and shopify_df to excel files
print("Export ebay and shopify dataframes to xlsx!")
ebay_df.to_excel(r'D:\Projects\shopContent\ebay\ebay.xlsx', index=False, header=True, encoding="utf-8")
shopify_df.to_excel(r'D:\Projects\shopContent\shopify\shopify.xlsx', index=False, header=True, encoding="utf-8")

# rename columns name in ebay dataframe
print("Renaming columns in ebay dataframe!")
ebay_df.rename(columns={'Custom Label (SKU)': 'Variant SKU', 'C:Brand': 'Vendor',
						'C:Type': 'Type', 'C:Season': 'Tags'}, inplace=True)

# exclude columns 'Vendor', 'Type', 'Tags' in shopify dataframe
print("Excluding columns in shopify dataframe!")
shopify_df = shopify_df[['Handle', 'Title', 'Body (HTML)', 'Published',
	   'Option1 Name', 'Option1 Value', 'Option2 Name', 'Option2 Value',
	   'Option3 Name', 'Option3 Value', 'Variant SKU', 'Variant Grams',
       'Variant Inventory Tracker', 'Variant Inventory Qty',
       'Variant Inventory Policy', 'Variant Fulfillment Service',
       'Variant Price', 'Variant Compare At Price',
       'Variant Requires Shipping', 'Variant Taxable', 'Variant Barcode',
       'Image Src', 'Image Position', 'Image Alt Text', 'Gift Card',
       'SEO Title', 'SEO Description',
       'Google Shopping / Google Product Category', 'Google Shopping / Gender',
       'Google Shopping / Age Group', 'Google Shopping / MPN',
       'Google Shopping / AdWords Grouping',
       'Google Shopping / AdWords Labels', 'Google Shopping / Condition',
       'Google Shopping / Custom Product', 'Google Shopping / Custom Label 0',
       'Google Shopping / Custom Label 1', 'Google Shopping / Custom Label 2',
       'Google Shopping / Custom Label 3', 'Google Shopping / Custom Label 4',
       'Variant Image', 'Variant Weight Unit', 'Variant Tax Code',
       'Cost per item']]

# replace unnecessary characters with blank in ebay dataframe
print("Replacing unnecessary symbols in ebay dataframe!")
ebay_df['Variant SKU'] = ebay_df['Variant SKU'].str.replace("-", '')
ebay_df['Variant SKU'] = ebay_df['Variant SKU'].str.replace("A", '')
ebay_df['Variant SKU'] = ebay_df['Variant SKU'].str.replace("B", '')
ebay_df['Variant SKU'] = ebay_df['Variant SKU'].str[:6]

# replace unnecessary characters with blank in shopify dataframe
print("Replacing unnecessary symbols in shopify dataframe!")
shopify_df['Variant SKU'] = shopify_df['Variant SKU'].str.replace("-", '')
shopify_df['Variant SKU'] = shopify_df['Variant SKU'].str.replace("\'", '')
shopify_df['Variant SKU'] = shopify_df['Variant SKU'].str.replace("A", '')
shopify_df['Variant SKU'] = shopify_df['Variant SKU'].str.replace("B", '')
shopify_df['Variant SKU'] = shopify_df['Variant SKU'].str[:6]

# delete rows-duplicates in ebay dataframe
print("Deleting duplicates in ebay dataframe!")
ebay_df = ebay_df.drop_duplicates(subset=['Variant SKU'], keep='first')

# left join shopify_df to ebay_df using column 'Variant SKU'
print('Joining shopify_df and ebay_df')
join_ebay_shopify_df = pd.merge(shopify_df, ebay_df, on='Variant SKU', how='left')

# set blank value in cell where 'Variant SKU' is null
print("Setting blank value in cell where 'Variant SKU' is null")
for index, row in join_ebay_shopify_df.iterrows():
    if row.isnull()['Variant SKU']:
        join_ebay_shopify_df.at[index, 'Vendor'] = ''
        join_ebay_shopify_df.at[index, 'Type'] = ''
        join_ebay_shopify_df.at[index, 'Tags'] = ''

# export join dataframe to excel file
print("Export final dataframe to xlsx!")
join_ebay_shopify_df.to_excel(r'D:\Projects\shopContent\final.xlsx', index=False, header=True, encoding="utf-8")

# time spent for execution
end_time = time.time()
print(f"\nTime spent: {end_time-start_time}")
