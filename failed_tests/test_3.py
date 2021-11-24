import sys
import pandas as pd
import numpy as np
from tkinter import *
from tkinter import filedialog
import pyodbc

def browse_files():
       global filename
       filename = filedialog.askopenfilename(initialdir = "/Downloads", title = "Selecciona un archivo de MDM:", filetypes = ((".xlsm Files","*.xlsm*"),("all files","*.*")))
       if filename:
              l1 = Label(window, text = "File path: " + filename).pack()
       else:
              print('No seleccionaste ningún archivo.')
       window.destroy()
def export_file():
    export_file_path = filedialog.asksaveasfilename(defaultextension = '.xlsx',initialdir = "/Desktop", title = "Guardar archivo como:", filetypes = (("Excel Files","*.xlsx*"),("CSV Files","*.csv*"),("all files","*.*")))
    data.to_excel(export_file_path, index = False, header=True)
    window2.destroy()
def validate_data(df:pd.DataFrame, q:pd.DataFrame):
       global data

       #Merge dataframes
       df['key'] = df['Stock Keeping Unit'].astype(str)+'|'+df['Vendor Ownership ID'].str.rstrip('-USA')+'|'+df['Vendor DCs.POV ID'].astype(str)
       data = pd.merge(left = df, right = q, on = 'key', how = 'left')


       # Package level validations
       package_level_1 = data['Package Level'].isin(['Package Level 1'])
       package_level_2 = data['Package Level'].isin(['Package Level 2'])
       package_level_3 = data['Package Level'].isin(['Package Level 3'])

       # Volume calculations
       each_volume = data['Dimensions - Each.Shipping Height']*data['Dimensions - Each.Shipping Length']*data['Dimensions - Each.Shipping Width']
       case_volume = data['Dimensions - Case.Shipping Height']*data['Dimensions - Case.Shipping Length']*data['Dimensions - Case.Shipping Width']
       inner_volume = data['Dimensions - Inner Pack.Shipping Height']*data['Dimensions - Inner Pack.Shipping Length']*data['Dimensions - Inner Pack.Shipping Width']

       # Volume validations
       case_volume_validation = case_volume >= each_volume*data['Quantity of Eaches in Package - Case']
       inner_volume_validation = inner_volume >= each_volume*data['Quantity of Eaches in Package - Inner Pack']

       # Weight validations
       case_weight_validation = data['Weights - Case.Weight'] >= data['Weights - Each.Weight']*data['Quantity of Eaches in Package - Case']
       inner_weight_validation = data['Weights - Inner Pack.Weight'] >= data['Weights - Each.Weight']*data['Quantity of Eaches in Package - Inner Pack']

       # GTIN validations
       each_gtin = data['Item-Level GTIN'] == data['Package Level GTIN - Each']
       inner_gtin = data['Package Level GTIN - Inner Pack'] != data['Package Level GTIN - Each']
       case_gtin = data['Package Level GTIN - Case'] != data['Package Level GTIN - Each']
       case_inner_gtin = data['Package Level GTIN - Case'] != data['Package Level GTIN - Inner Pack']

       # Package 1 validations
       max_cases_pallet_layer_inv = data['Maximum Cases per Pallet Layer'].isin([0,1,'NaN',''])
       max_cases_pallet_layer = ~max_cases_pallet_layer_inv
       max_pallets_inv = data['Maximum Cases per Pallet Layer'].isin([0,1,'NaN',''])
       max_pallets = ~max_pallets_inv
       qty_per_case_1 = data['Quantity of Eaches in Package - Each'] == data['Quantity per Case'] #When True, Ok
       ship_round_qty_1 = data['Quantity of Eaches in Package - Each'] == data['Ship Round Quantity'] #When True, Ok
       store_order_pack_1 = data['Quantity of Eaches in Package - Each'] == data['Store Order Pack'] #When True, Ok

       # Package 2 validations
       qty_per_case_2 = data['Quantity of Eaches in Package - Case'] == data['Quantity per Case'] #When True, Ok
       ship_round_qty_2 = data['Quantity of Eaches in Package - Case'] == data['Ship Round Quantity'] #When True, Ok
       store_order_pack_2 = data['Quantity of Eaches in Package - Case'] == data['Store Order Pack'] #When True, Ok
       # Validate Inner fields are empty

       # Package 3 validations
       qty_per_case_2 = data['Quantity of Eaches in Package - Case'] == data['Quantity per Case'] #When True, Ok
       ship_round_qty_3 = data['Quantity of Eaches in Package - Inner Pack'] == data['Ship Round Quantity'] #When True, Ok
       store_order_pack_3 = data['Quantity of Eaches in Package - Inner Pack'] == data['Store Order Pack'] #When True, Ok

# Main

if __name__=="__main__":

	window = Tk()
	window.title("Data Governance: Step 2")
	window.geometry("200x100")
	label_file_explorer = Label(window, text = "Data Governance: Step 2", width = 100, height = 4).pack()
	button_explore = Button(window, text = "Buscar archivo", command = browse_files).pack()
	window.mainloop()

	print('Reading MDM file...')
	try:
			mdm = pd.read_excel(f'{filename}', index_col = False,sheet_name='Entities',header=1,dtype=object,na_filter= False)
			skus_list = list(mdm['Stock Keeping Unit'])
			skus_list_string = str(skus_list)
			skus = skus_list_string[1:-1]
			print('MDM file read successfully')
	except:
			raise Exception('Error at reading MDM file. Please check file and try again. Finishing execution.')
	
	print('Connecting to TORRENT...')
	try:
			conn = pyodbc.connect(Driver='{SQL Server}',
						Server='TORRENT',
						Database='db_product_integrity',
						Trusted_Connection='yes')
			print('Connected successfully to TORRENT')
			sql = ("""SELECT DISTINCT 
                                   CAST(A.sku AS VARCHAR) + '|' +  CAST(V.AP_REF AS VARCHAR) + '|' +  A.vendor_id AS [key]
                                   , CAST(A.sku AS VARCHAR) + '|' +  CAST(V.AP_REF AS VARCHAR)  [sec_key]
                                   , A.SKU
                                   , COALESCE(A.INNER_QTY, 0) INNER_QTY
                                   , COALESCE(A.OUTER_QTY, 0) OUTER_QTY
                                   , CASE WHEN COALESCE(A.INNER_QTY, 0) > 1 THEN 'Package Level 3'
                                          WHEN COALESCE(A.OUTER_QTY, 0) > 1 THEN 'Package Level 2'
                                          ELSE 'Package Level 1' END [Package Level]
                                   , V.AP_REF, A.vendor_id
                                   , ISNULL(B.rank, 999) VENDOR_RANK
                            FROM [db_product_integrity].[prod].[tb_sku_warehouse_status_vendor_upload] A
                            LEFT JOIN db_product_integrity.prod.tb_az_vendors V
                            ON A.vendor_id = V.VENDOR 
                            LEFT JOIN db_product_integrity.prod.tb_item_vendor_rank B
                            ON A.sku = B.item
                            AND A.vendor_id = b.vendor_id
                            WHERE  A.INTL_CODE = 'USA' AND
                            A.WHSE IN(10,11,20,22,33,55,66,77,88,99,91) AND 
                            A.SKU IN({})
                            order by A.sku, VENDOR_RANK""").format(skus)
			query = pd.read_sql(sql,conn)
			print('Query run successful. Closing connection...')
			conn.close()
			print('Closed connection to TORRENT successfully')
	except:
			raise Exception('Error connecting to Torrent. Finishing execution.')
	
	# Data validation
	print('Data validation...')
	validate_data(mdm,query)
	#data_integrity(data)
	print('Data is valid. Exporting file...')

	# File export
	# window2 = Tk()
	# window2.title("Data Governance: Step 2")
	# window2.geometry("200x100")
	# label_file_explorer = Label(window2, text = "Data Governance: Step 2", width = 100, height = 4).pack()
	# button_explore = Button(window2, text = "Exportar archivo", command = export_file).pack()
	# window2.mainloop()
print("Program end")
