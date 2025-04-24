import pandas as pd
import streamlit as st
import os
import numpy as np
from datetime import datetime
import io
import openpyxl
import numpy as np

from streamlit_option_menu import option_menu



with st.sidebar:
    selected = option_menu(
        menu_title = None,
        options = ['MIB Product List', 'PO Creater'],


    )

if selected == 'MIB Product List':


    st.title('MIB Product List Excel')
    uploaded_file = st.file_uploader('Upload XLSX here:', type = ['xlsx'])


    if uploaded_file:
        productlist_df_dict = pd.read_excel(uploaded_file, sheet_name = None)

        colordf = productlist_df_dict['colors']
        productlist_df = productlist_df_dict['Current'].iloc[3:, :]



        productlist_df.columns = productlist_df.iloc[0]

        productlist_df = productlist_df[1:].reset_index(drop = True)

        Quickbooks_newproducts = pd.DataFrame(columns =  ['Product/service name', 'Category',	'Item type',	'SKU',	'Sales description',	'Sales price/rate', 'Income account',	'Purchase description',	'Purchase cost',	'Expense account',	'Quantity on hand',  'Quantity as-of date',	'Reorder point',	'Inventory asset account'])


        smallproductlist = productlist_df[productlist_df['NEW FOR SEASON'].notna()][['1X','2X','3X','4X','5X', '6X', '7X', '8X','NEW FOR SEASON', 'SKU', 'Description', 'Retail', 'Category', 'Country of Origin', 'hts code', 'Height', 'Width', 'Length', 'Weight']]

        sizes = ['1X','2X', '3X', '4X', '5X', '6X', '7X', '8X']
        newitems = []
        itemsku = []
        category = []
        purchasecosts = []
        sales = []
        totalnamelist = []
        countrylist = []
        heightlist = []
        widthlist = []
        lengthlist = []
        weightist = []
        htslist = []
        for index, row in smallproductlist.iterrows():
            sku = smallproductlist.loc[index, 'SKU']
            cat = 'Clothing:' +smallproductlist.loc[index, 'Category']
            sale = smallproductlist.loc[index, 'Retail']
            for j in sizes:
                purchasecost = smallproductlist.loc[index, j]
                if not pd.isna(smallproductlist.loc[index, j]):
                    for k in [item.strip() for item in smallproductlist.loc[index, 'NEW FOR SEASON'].split(",")]: 

                        hts = smallproductlist.loc[index, 'hts code']
                        htslist.append(hts)
                        height= smallproductlist.loc[index, 'Height']
                        heightlist.append(height)
                        width = smallproductlist.loc[index, 'Width']
                        widthlist.append(width)
                        length = smallproductlist.loc[index, 'Length']
                        lengthlist.append(length)
                        weight = smallproductlist.loc[index, 'Weight']
                        weightist.append(weight)
                        newitem =smallproductlist.loc[index, 'Description'] +' '+ str(j)+  ' '+ str(k)
                        totalnames =smallproductlist.loc[index, 'Description']
                        country = smallproductlist.loc[index, 'Country of Origin']
                        totalnamelist.append(totalnames)
                        sku_2 = str(sku)+'-'+str(j)+'-'+str(k)
                        newitems.append(newitem)
                        itemsku.append(sku_2)
                        category.append(cat)
                        purchasecosts.append(purchasecost)
                        sales.append(sale)
                        countrylist.append(country)
        

        Quickbooks_newproducts['Product/service name'] = newitems
        Quickbooks_newproducts['SKU'] = itemsku
        Quickbooks_newproducts['Category'] = category
        Quickbooks_newproducts['Purchase cost'] = purchasecosts
        Quickbooks_newproducts['Sales price/rate'] = sales


        Quickbooks_newproducts['Item type'] = 'Inventory'
        #Quickbooks_newproducts['Purchase description'] = 'INCOME'
        Quickbooks_newproducts['Income account'] = '4010 Sales:Merchandise'
        Quickbooks_newproducts['Expense account'] = '5005 Cost of Good Sold:COGS - Finished Goods'
        Quickbooks_newproducts['Quantity on hand'] = 0
        #Quickbooks_newproducts['Reorder point'] = np.nan
        Quickbooks_newproducts['Inventory asset account'] = '1450 Inventory:Finished Goods'
        Quickbooks_newproducts['Quantity as-of date'] = datetime.today().strftime("%m/%d/%Y")






        ####Second doctument 

        stringsl = 'stlye	desc    	color	color_desc	sizes	size_desc	coo	UPC	sku_ref		price	HEIGHT	WIDTH	LENGTH	WEIGHT'

        SKU_Upload_seconddoc = pd.DataFrame(columns = stringsl.split())

        def convert_df_to_csv(df):
        # Use StringIO to write to a string buffer (instead of a file)
            csv_buffer = io.StringIO()
            df.to_csv(csv_buffer, index=False)
            return csv_buffer.getvalue()

        csv_data_qb = convert_df_to_csv(Quickbooks_newproducts)


        ####nbd start 
        colornamelist = []

        for color in [i.split('-')[2] for i in itemsku]:
            if colordf[['seasons color','Unnamed: 2']][colordf['seasons color'].str.lower() == color.lower()]['Unnamed: 2'].empty:
                colorname = 'missing'
            else:
                temp = colordf[['seasons color','Unnamed: 2']][colordf['seasons color'].str.lower() == color.lower()]['Unnamed: 2']
                
                colorname = temp.iloc[0]

            colornamelist.append(colorname)


        
        SKU_Upload_seconddoc['style']= itemsku
        SKU_Upload_seconddoc['desc'] = totalnamelist
        SKU_Upload_seconddoc['color'] = [i.split('-')[2] for i in itemsku]
        SKU_Upload_seconddoc['color_desc'] =  colornamelist
        SKU_Upload_seconddoc['sizes'] = [i.split('-')[1] for i in itemsku]
        SKU_Upload_seconddoc['size_desc']= [i.split('-')[1] for i in itemsku]
        SKU_Upload_seconddoc['coo'] = countrylist
        SKU_Upload_seconddoc['UPC']= itemsku
        SKU_Upload_seconddoc['sku_ref']= itemsku
        SKU_Upload_seconddoc['price'] =sales
        SKU_Upload_seconddoc['hts code']=  htslist
        SKU_Upload_seconddoc['HEIGHT'] =heightlist
        SKU_Upload_seconddoc['WIDTH'] =widthlist
        SKU_Upload_seconddoc['LENGTH'] =lengthlist
        SKU_Upload_seconddoc['WEIGHT'] =weightist


        SKU_Upload_seconddoc = SKU_Upload_seconddoc[['style', 'desc', 'color', 'color_desc', 'sizes', 'size_desc', 'coo',
        'UPC', 'sku_ref',  'hts code', 'price', 'HEIGHT', 'WIDTH', 'LENGTH', 'WEIGHT']]

        

        csv_data_seconddoc = convert_df_to_csv(SKU_Upload_seconddoc)

        #### Shopfiy start


        shopify = pd.DataFrame(columns = [
        "Handle", "Title", "Body (HTML)", "Vendor", "Product Category", "Type", "Tags", "Published",
        "Option1 Name", "Option1 Value", "Option1 Linked To", "Option2 Name", "Option2 Value", "Option2 Linked To",
        "Option3 Name", "Option3 Value", "Option3 Linked To", "Variant SKU", "Variant Grams",
        "Variant Inventory Tracker", "Variant Inventory Policy", "Variant Fulfillment Service",
        "Variant Price", "Variant Compare At Price", "Variant Requires Shipping", "Variant Taxable",
        "Variant Barcode", "Image Src", "Image Position", "Image Alt Text", "Gift Card",
        "SEO Title", "SEO Description", "Google Shopping / Google Product Category",
        "Google Shopping / Gender", "Google Shopping / Age Group", "Google Shopping / MPN",
        "Google Shopping / Condition", "Google Shopping / Custom Product",
        "Google Shopping / Custom Label 0", "Google Shopping / Custom Label 1",
        "Google Shopping / Custom Label 2", "Google Shopping / Custom Label 3",
        "Google Shopping / Custom Label 4", "Variant Image", "Variant Weight Unit",
        "Variant Tax Code", "Cost per item", "Included / United States", "Price / United States",
        "Compare At Price / United States", "Included / Canada", "Price / Canada",
        "Compare At Price / Canada", "Included / International", "Price / International",
        "Compare At Price / International", "Status"
        ])


        shopify['Title'] = totalnamelist
        shopify['Option1 Value'] = [i.split('-')[1] for i in itemsku]
        shopify['Option2 Value'] = colornamelist
        shopify['Variant Price'] = sales
        shopify['Variant SKU'] = itemsku
        shopify['Variant Barcode'] = itemsku
        shopify['Cost per item'] = purchasecosts


        shopify['Variant Grams'] = 0.181436948
        shopify['Variant Inventory Tracker'] = 'shopify'
        shopify['Variant Inventory Policy'] = 'deny'
        shopify['Variant Fulfillment Service'] = 'nbd-vacaville'
        shopify['Variant Requires Shipping'] = 'TRUE'
        shopify['Variant Taxable'] = 'TRUE'
        shopify['Variant Weight Unit'] = 'lb'
        shopify['Handle']= shopify['Title'].str.replace(" ", "-").str.lower()

        shopify['Vendor'] = 'MIB Clothing'
        shopify['Product Category'] = 'Apparel & Accessories > Clothing'
        shopify['Type'] = [i.split(':',4)[1] for i in category]
        shopify['Option1 Name'] = 'Size'
        shopify['Option2 Name'] = 'Color'
        shopify['Gift Card'] = 'FALSE'
        shopify['Included / United States'] = 'TRUE'
        shopify['Included / Canada'] = 'TRUE'
        shopify['Included / International']= 'TRUE'
        shopify['Status']= 'draft'

        shopify.loc[shopify.duplicated(subset=['Handle']), ['Vendor', 'Product Category','Type',
                                                            'Option1 Name', 'Option2 Name',
                                                        'Gift Card', 'Included / United States',
                                                            'Included / Canada', 'Included / International',
                                                        'Status']] = ""



        csv_data_shopify = convert_df_to_csv(shopify)
        st.download_button(label ='Download Quickbooks CSV', data = csv_data_qb, file_name = 'Quickbooks_NewProductImport.csv', mime ='text/csv' )
        st.download_button(label ='Download NBD CSV', data = csv_data_seconddoc, file_name = 'NBDImport.csv', mime ='text/csv' )
        st.download_button(label ='Download Shopify CSV', data = csv_data_shopify, file_name = 'ShopifyImports.csv', mime ='text/csv' )

                
if selected == 'PO Creater':


    st.title('PO Creater')
    uploaded_file_prodlist = st.file_uploader('Upload Product List Here:', type = ['xlsx'])
    uploaded_file_OG = st.file_uploader('Upload Orange and Green Report Here:', type = ['xlsx'])

    if uploaded_file_prodlist and uploaded_file_OG:

        productlist_df_dict = pd.read_excel(uploaded_file_prodlist, sheet_name = 'Current')
        prodlist_colors = pd.read_excel(uploaded_file_prodlist,  sheet_name = 'colors', header= None)
        ogdf = pd.read_excel(uploaded_file_OG, sheet_name = None)

        newheaders= prodlistdf.iloc[3]
        prodlistdf = prodlistdf[4:]
        prodlistdf.columns = list(newheaders)

        prodlist_colors.columns =prodlist_colors.iloc[1]
        prodlist_colors = prodlist_colors.iloc[2:]



        potemplate = pd.read_excel('data/PO Template.xlsx')



        ogtableliststart = list(ogdf[ogdf[0] == 'PARENT'].index)
        ogtablelistend = list(ogdf[ogdf[0] == 'Grand Total'].index)

        counter = 0
        ogtables = []
        ogvenders = {}
        for i in ogtableliststart:
            ogtables.append(ogdf.iloc[i,1])
            globals()[f'ogtable_{ogdf.iloc[i,1]}'] = ogdf.iloc[ogtableliststart[counter]:int(ogtablelistend[counter])+1]
            globals()[f'PO_{ogdf.iloc[i,1]}'] =potemplate
            globals()[f'PO_{ogdf.iloc[i,1]}'].iloc[13,11] = ogdf.iloc[i,1]
            globals()[f'ogsubtable_{ogdf.iloc[i,1]}'] = globals()[f'ogtable_{ogdf.iloc[i,1]}'].reset_index().iloc[4:]
            globals()[f'ogsubtable_{ogdf.iloc[i,1]}'].columns = list(globals()[f'ogtable_{ogdf.iloc[i,1]}'].reset_index().iloc[3])
            counter+=1

        for j in ogtables:
            globals()[f'sizelist_{j}'] =[]
            ogsubtable = globals()[f'ogsubtable_{j}']
            ind =list(globals()[f'ogsubtable_{j}'].columns)
            for i in list(range(2,9)):
                colname = ind[i]
                if ogsubtable[colname].sum() >0:
                    globals()[f'sizelist_{j}'].append(ind[i])




        po_dict = {}
        for i in ogtables:
            

            prodlistskuinf = prodlistdf[prodlistdf['SKU']==i]
            
            sizecostlist =[]
            [sizecostlist.append(x) for x in list(prodlistskuinf[globals()[f'sizelist_{i}']].reset_index(drop = True).iloc[0]) if str(x) != 'nan']
            
            
            prodlistsizecost =  pd.DataFrame(prodlistskuinf[globals()[f'sizelist_{i}']].reset_index(
                drop = True).iloc[0])
            prodlistsizecost = prodlistsizecost[prodlistsizecost[0].notnull()]
            
            for j in set(list(prodlistsizecost[0])):
                tempcolnames = ['Row Labels'] +list(prodlistsizecost[prodlistsizecost[0]==j].index)
                temppotable =globals()[f'ogsubtable_{i}'].loc[:globals()[f'ogsubtable_{i}'][globals()[f'ogsubtable_{i}']['Row Labels']== 'Grand Total'
                ].index[0]-1][tempcolnames]
                temppotable['unit cost'] =j
                missinglist = []
                for x in globals()[f'sizelist_{i}']:
                    if x not in list(prodlistsizecost[prodlistsizecost[0] == j].index):
                        missinglist.append(x)
                tempfullcolnames = (['Row Labels'] +missinglist+list(prodlistsizecost[prodlistsizecost[0]==j].index)+ ['unit cost'])
                globals()[f'potable_{i}_{j}'] = temppotable.reindex(columns = tempfullcolnames)

                po_dict[f'potable_{i}_{j}'] = globals()[f'potable_{i}_{j}']


        for i in list(po_dict.keys()):   
            if str(f"po_subtable_{i.split('_')[1]}") in globals():
                del globals()[f"po_subtable_{i.split('_')[1]}"]


        
        costlistforpo = []
        for i in list(po_dict.keys()):   
            if str(f"po_subtable_{i.split('_')[1]}") in globals():
                globals()[f"po_subtable_{i.split('_')[1]}"] = pd.concat([globals()[f"po_subtable_{i.split('_')[1]}"],
                                        po_dict[i].reindex(columns =('Row Labels','2X','3X','4X','5X','6X','7X','8X', 'unit cost') )
                                                            ] )
                globals()[f"po_subtable_{i.split('_')[1]}"].reset_index(drop =True, inplace =True)
            else:
                globals()[f"po_subtable_{i.split('_')[1]}"] = po_dict[i].reindex(columns =('Row Labels','2X','3X','4X','5X','6X','7X','8X', 'unit cost'
        ) )
                costlistforpo.append(i.split('_')[2])





        PO_dataframes = []
        PO_dataframes_names = []
        for i in ogtables:
            
            try:
                globals()[f'po_subtable_{i}']
            
                globals()[f'po_subtable_{i}'] = globals()[f'po_subtable_{i}'].applymap(lambda x: round(x) if isinstance(x, float) and not pd.isna(x) else x)
                globals()[f'po_subtable_{i}']['Total'] = globals()[f'po_subtable_{i}'][['2X', '3X', '4X', '5X', '6X', '7X', '8X']].sum(axis= 1)
                globals()[f'po_subtable_{i}']['Total Cost'] = globals()[f'po_subtable_{i}']['Total']* globals()[f'po_subtable_{i}']['unit cost']
                rowtotal = (list(globals()[f'po_subtable_{i}'][['2X', '3X', '4X', '5X', '6X', '7X', '8X']][0:len(globals()[f'po_subtable_{i}'])-1].sum()
                            ))
                rowtotal.append(globals()[f'po_subtable_{i}'][['2X', '3X', '4X', '5X', '6X', '7X', '8X']][0:len(globals()[f'po_subtable_{i}'])-1].sum().sum()
                            )
                rowtotal.append('Total Cost')
                rowtotal.append(globals()[f'po_subtable_{i}']['Total Cost'].sum())
                rowtotal.insert(0, '')
                rowtotaldf = pd.DataFrame(rowtotal).transpose()
                rowtotaldf.columns = globals()[f'po_subtable_{i}'].columns
                globals()[f'po_subtable_{i}'] = pd.concat([globals()[f'po_subtable_{i}'],rowtotaldf], ignore_index = True)
            
                vendercolor = []
                mibcolor = []
                for row, item in globals()[f'po_subtable_{i}'].iloc[:len(globals()[f'po_subtable_{i}'])-1].iterrows():
                    try: 
                        tet =prodlist_colors[prodlist_colors['FOR SKU'].str.lower() == globals()[f'po_subtable_{i}'].iloc[row]['Row Labels'].lower()].reset_index(drop = True).iloc[0]
                        vendercolor.append(tet['Pantone/Vendor Code'])
                        mibcolor.append(tet['Full Name'].upper())
                    except IndexError:
                        mibcolor.append('missing')
                        vendercolor.append('missing')
                
                    
                
                vendercolor.append('')
                mibcolor.append('')
            
            
                globals()[f'po_subtable_{i}']['MIB Color'] = mibcolor
                globals()[f'po_subtable_{i}']['Vendor Color'] = vendercolor
                globals()[f'po_subtable_{i}']['For SKU'] = globals()[f'po_subtable_{i}']['Row Labels']
            
                
                
                
                globals()[f'po_subtable_{i}_long'] = globals()[f'po_subtable_{i}'][[
                
                    'MIB Color', 
                    'Vendor Color',
                    'For SKU',
                    '2X',
                    '3X',
                    '4X',
                    '5X',
                    '6X',
                    '7X',
                    '8X',
                    'Total',
                    'unit cost',
                    'Total Cost'
                ]].iloc[:-1]
            
            
                globals()[f'po_subtable_{i}_long']['Unnamed: 15'] = ''
                globals()[f'po_subtable_{i}_long']['Unnamed: 16'] = ''
                globals()[f'po_subtable_{i}_long']['Unnamed: 14'] = ''
                globals()[f'po_subtable_{i}_long']['Unnamed: 0'] = ''
                globals()[f'po_subtable_{i}_long']['Unnamed: 1'] = ''
                
                
                
                
                df1 =globals()[f'PO_{i}'].iloc[:16]
                df1.iloc[10,2] = date.today().strftime("%m/%d/%Y")
                df1.iloc[10,4] = (date.today()+timedelta(days=60)).strftime("%m/%d/%Y")
                df1.iloc[10,14] = (date.today()+timedelta(days=60)).strftime("%m/%d/%Y")
                
                df1.iloc[13,3] = i
                df1.iloc[14,3] = prodlistdf[prodlistdf['SKU']==i]['Description']

                df2 = globals()[f'po_subtable_{i}_long'][[
                        'Unnamed: 0',
                        'Unnamed: 1',
                        'MIB Color', 
                        'Vendor Color',
                        'For SKU',
                        '2X',
                        '3X',
                        '4X',
                        '5X',
                        '6X',
                        '7X',
                        '8X',
                        'Total',
                        'unit cost',
                        'Total Cost'
                ]]
                

                
                df3 =globals()[f'PO_{i}'].iloc[globals()[f'PO_{i}'][
                    globals()[f'PO_{i}']['Unnamed: 2'
                    ]=='''ALL SIZES MUST BE TO "MAKING IT BIG SPECS"'''].index[0]:]
                df2.columns = df1.columns
                df3.columns = df1.columns

                for j in range(5,12):
                    df3 = df3.reset_index(drop=True)
                    df3.iloc[1,j] = df2.iloc[:,j].sum()

                totalcost= df2.iloc[:,14].sum()
                df1.iloc[14,14] = totalcost
                df3.iloc[1,14] =totalcost
                df1.iloc[3,3] = prodlistdf[['SKU', ' Vendor']][prodlistdf['SKU']== i].iloc[0,1]
                
                globals()[f'po_output_{i}']= pd.concat([df1, df2, df3],ignore_index=True)
                PO_dataframes.append(globals()[f'po_output_{i}'])
                PO_dataframes_names.append(f'po_output_{i}')


            except KeyError:
                pass




        
        def convert_dfs_to_xlsx(dfs_dict):
            """
            Takes a dictionary of DataFrames and returns an in-memory Excel file (as bytes).
            Each key becomes a sheet name.
            
            Args:
                dfs_dict (dict): { "SheetName1": df1, "SheetName2": df2, ... }
            
            Returns:
                bytes: Excel file as byte stream.
            """
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                for sheet_name, df in dfs_dict.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
            output.seek(0)
            return output.getvalue()
        
        po_output_dict = {}

        for i in ogtables:
            po_output_dict[i] = globals()[f'ogtable_{i}']


        po_output = convert_dfs_to_xlsx(po_output_dict)
        st.download_button(label ='Download PO Output XLSX', data = po_output, file_name = 'PO_Output.xlsx')
        




            
                    
                    
                        




        





