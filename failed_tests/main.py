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
# def maximum_cases_pallet_layer(value):
     #if not value or value in [0,1,'None']:
         #return 7
     #else:
         #return value
# def pallet_layer_maximum(value):
     #if not value or value in [0,1,'None']:
         #return 4
     #else:
         #return value		
#def data_integrity(df: pd.DataFrame) -> bool:
	
	#Check if data is empty
	#if df.empty:
		#print('File is empty. No items found. Finishing execution')
		#return False

	# Check for unique keys
	#if pd.Series(df['key']).is_unique:
		#pass
	#else:
		#raise Exception('Key is not unique. Check data')

	#Check for nulls
	#if df.isnull().values.any():
		#raise Exception('Null values found. Check data')
def validate_data(df:pd.DataFrame, q:pd.DataFrame):
       global data

       #Merge dataframes
       df['key'] = df['Stock Keeping Unit'].astype(str)+'|'+df['Vendor Ownership ID'].str.rstrip('-USA')+'|'+df['Vendor DCs.POV ID'].astype(str)
       data = pd.merge(left = df, right = q, on = 'key', how = 'left')

       each_merch_dimensions = data[['Dimensions - Each.Merchandising Height',
       'Dimensions - Each.Merchandising Length',
       'Dimensions - Each.Merchandising Width']]

       each_ship_dimensions = data[['Dimensions - Each.Shipping Height',
       'Dimensions - Each.Shipping Length',
       'Dimensions - Each.Shipping Width',
       'Dimensions - Each.UOM',
       'Quantity of Eaches in Package - Each']]

       each_weights = data[['Weights - Each.Weight',
       'Weights - Each.UOM']]

       case_merch_dimensions = data[['Dimensions - Case.Merchandising Height',
       'Dimensions - Case.Merchandising Length',
       'Dimensions - Case.Merchandising Width']]

       case_ship_dimensions = data[[
       'Dimensions - Case.Shipping Height',
       'Dimensions - Case.Shipping Length',
       'Dimensions - Case.Shipping Width',
       'Dimensions - Case.UOM',
       'Quantity of Eaches in Package - Case']]
              
       case_weights = data[[
       'Weights - Case.Weight',
       'Weights - Case.Weight UOM']]
              
       inner_merch_dimensions = data[[
       'Dimensions - Inner Pack.Merchandising Height',
       'Dimensions - Inner Pack.Merchandising Length',
       'Dimensions - Inner Pack.Merchandising Width']]

       inner_ship_dimensions = data[[
       'Dimensions - Inner Pack.Shipping Height',
       'Dimensions - Inner Pack.Shipping Length',
       'Dimensions - Inner Pack.Shipping Width',
       'Dimensions - Inner Pack.UOM',
       'Quantity of Eaches in Package - Inner Pack']]

       inner_weights = data[['Weights - Inner Pack.Weight',
       'Weights - Inner Pack.UOM']]

       # Empty field validations

       each_merch_isempty = each_merch_dimensions.dropna(0,'all').empty
       each_ship_isempty = each_ship_dimensions.dropna(0,'all').empty
       each_weights_isempty = each_weights.dropna(0,'all').empty
       each_gtin_isempty = data['Package Level GTIN - Each'].dropna(0,inplace = False, how='all').empty
       item_gtin_isempty = data['Item-Level GTIN'].dropna(0,inplace = False, how='all').empty
       case_merch_isempty = case_merch_dimensions.dropna(0,'all').empty
       case_ship_isempty = case_ship_dimensions.dropna(0,'all').empty
       case_weights_isempty = case_weights.dropna(0,'all').empty
       case_gtin_isempty = data['Package Level GTIN - Case'].dropna(0,inplace = False, how='all').empty
       inner_merch_isempty = inner_merch_dimensions.dropna(0,'all').empty
       inner_ship_isempty = inner_ship_dimensions.dropna(0,'all').empty
       inner_weights_isempty = inner_weights.dropna(0,'all').empty
       inner_gtin_isempty = data['Package Level GTIN - Inner Pack'].dropna(0,inplace = False, how='all').empty

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


       #Empty Each fields
       conditions_each_fields = [
              package_level_1 & (each_merch_isempty == False) & (each_ship_isempty == False) & (each_weights_isempty == False) & (each_gtin_isempty == False) & (item_gtin_isempty == False),
              package_level_1 & item_gtin_isempty,
              package_level_1 & (each_gtin_isempty | each_merch_isempty | each_ship_isempty | each_weights_isempty),
              package_level_2 & (each_merch_isempty == False) & (each_ship_isempty == False) & (each_weights_isempty == False) & (each_gtin_isempty == False) & (item_gtin_isempty == False),
              package_level_2 & item_gtin_isempty,
              package_level_2 & (each_gtin_isempty | each_merch_isempty | each_ship_isempty | each_weights_isempty),
              package_level_3 & (each_merch_isempty == False) & (each_ship_isempty == False) & (each_weights_isempty == False) & (each_gtin_isempty == False) & (item_gtin_isempty == False),
              package_level_3 & item_gtin_isempty,
              package_level_3 & (each_gtin_isempty | each_merch_isempty | each_ship_isempty | each_weights_isempty)
       ]

       options_each_fields = [
              'Package Level 1: Ok. All Each fields have data',
              'Package Level 1: Item-GTIN is empty. Check',
              'Package Level 1: There are empty Each fields. Check',
              'Package Level 2: All Each fields have data',
              'Package Level 2: Item-GTIN is empty. Check',
              'Package Level 2: There are empty Each fields. Check',
              'Package Level 3: All Each fields have data',
              'Package Level 3: Item-GTIN is empty. Check',
              'Package Level 3: There are empty Each fields. Check',
       ]
       data['Empty Each'] = np.select(conditions_each_fields,options_each_fields,default='No Package Level Data: Check for empty Each fields')

       #Empty Case fields
       conditions_case_fields = [
              package_level_1 & ((case_merch_isempty == False) | (case_ship_isempty == False) | (case_weights_isempty == False) | (case_gtin_isempty == False)),
              package_level_1 & case_gtin_isempty & case_merch_isempty & case_ship_isempty & case_weights_isempty,
              package_level_2 & (case_merch_isempty == False) & (case_ship_isempty == False) & (case_weights_isempty == False) & (case_gtin_isempty == False),
              package_level_2 & (case_gtin_isempty | case_merch_isempty | case_ship_isempty | case_weights_isempty),
              package_level_3 & (case_merch_isempty == False) & (case_ship_isempty == False) & (case_weights_isempty == False) & (case_gtin_isempty == False),
              package_level_3 & (case_gtin_isempty | case_merch_isempty | case_ship_isempty | case_weights_isempty)
       ]

       options_case_fields = [
              'Package Level 1: Some Case fields have data. Check',
              'Package Level 1: Ok. Case fields are empty',
              'Package Level 2: Ok. All Case fields have data',
              'Package Level 2: Some Case fields are empty. Check',
              'Package Level 3: Ok. All Case fields have data',
              'Package Level 3: Some Case fields are empty. Check'
       ]
       data['Empty Case'] = np.select(conditions_case_fields,options_case_fields,default='No Package Level Data: Check for empty Case fields')

       #Empty Inner Pack fields
       conditions_inner_fields = [
              package_level_1 & ((inner_merch_isempty == False) | (inner_ship_isempty == False) | (inner_weights_isempty == False) | (inner_gtin_isempty == False)),
              package_level_1 & inner_gtin_isempty & inner_merch_isempty & inner_ship_isempty & inner_weights_isempty,
              package_level_2 & ((inner_merch_isempty == False) | (inner_ship_isempty == False) | (inner_weights_isempty == False) | (inner_gtin_isempty == False)),
              package_level_2 & inner_gtin_isempty & inner_merch_isempty & inner_ship_isempty & inner_weights_isempty,
              package_level_3 & (inner_merch_isempty == False) & (inner_ship_isempty == False) & (inner_weights_isempty == False) & (inner_gtin_isempty == False),
              package_level_3 & (inner_gtin_isempty | inner_merch_isempty | inner_ship_isempty | inner_weights_isempty)
       ]

       options_inner_fields = [
              'Package Level 1: Some Inner Pack fields have data. Check',
              'Package Level 1: Ok. Inner Pack fields are empty',
              'Package Level 2: Some Inner Pack fields have data. Check',
              'Package Level 2: Ok. Inner Pack fields are empty',
              'Package Level 3: Ok. All Inner Pack fields have data',
              'Package Level 3: Some Inner Pack fields are empty. Check'
       ]
       data['Empty Inner Pack'] = np.select(conditions_inner_fields,options_inner_fields,default='No Package Level Data: Check for empty Inner Pack fields')

       # Quantities validations
       conditions_validations = [
              package_level_1 & max_cases_pallet_layer & max_pallets & qty_per_case_1 & ship_round_qty_1 & store_order_pack_1,
              package_level_1 & max_cases_pallet_layer & max_pallets & ((qty_per_case_1 & ship_round_qty_1 & store_order_pack_1) == False),
              package_level_1 & ((max_cases_pallet_layer & max_pallets) == False) & qty_per_case_1 & ship_round_qty_1 & store_order_pack_1,
              package_level_1 & ((max_cases_pallet_layer & max_pallets) == False) & ((qty_per_case_1 & ship_round_qty_1 & store_order_pack_1) == False),
              package_level_2 & qty_per_case_2 & ship_round_qty_2 & store_order_pack_2,
              package_level_2 & qty_per_case_2 & ship_round_qty_2 & (store_order_pack_2 == False),
              package_level_2 & qty_per_case_2 & (ship_round_qty_2 == False) & store_order_pack_2,
              package_level_2 & qty_per_case_2 & (ship_round_qty_2 == False) & (store_order_pack_2 == False),
              package_level_2 & (qty_per_case_2 == False) & ship_round_qty_2 & store_order_pack_2,
              package_level_2 & (qty_per_case_2 == False) & ship_round_qty_2 & (store_order_pack_2 == False),
              package_level_2 & (qty_per_case_2 == False) & (ship_round_qty_2 == False) & store_order_pack_2,
              package_level_2 & (qty_per_case_2 == False) & (ship_round_qty_2 == False) & (store_order_pack_2 == False),
              package_level_3 & qty_per_case_2 & ship_round_qty_3 & store_order_pack_3,
              package_level_3 & qty_per_case_2 & ship_round_qty_3 & (store_order_pack_3 == False),
              package_level_3 & qty_per_case_2 & (ship_round_qty_3 == False) & store_order_pack_3,
              package_level_3 & qty_per_case_2 & (ship_round_qty_3 == False) & (store_order_pack_3 == False),
              package_level_3 & (qty_per_case_2 == False) & ship_round_qty_3 & store_order_pack_3,
              package_level_3 & (qty_per_case_2 == False) & ship_round_qty_3 & (store_order_pack_3 == False),
              package_level_3 & (qty_per_case_2 == False) & (ship_round_qty_3 == False) & store_order_pack_3,
              package_level_3 & (qty_per_case_2 == False) & (ship_round_qty_3 == False) & (store_order_pack_3 == False),
       ]
       options_validations = [
              'Package Level 1: Ok',
              'Package Level 1: Check Qty per Case, Ship Round Qty & Store Order Pack = "Quantity of Eaches in Package - Each"',
              'Package Level 1: Check Max Cases per Pallet Layer = 7 and/or Pallet Layer Max = 4',
              'Package Level 1: Check Qty per Case, Ship Round Qty & Store Order Pack = "Quantity of Eaches in Package - Each", Check Max Cases per Pallet Layer = 7 and/or Pallet Layer Max = 4',
              'Package Level 2: Ok',
              'Package Level 2: Check Store Order Pack = "Quantity of Eaches in Package - Case"',
              'Package Level 2: Check Ship Round Qty = "Quantity of Eaches in Package - Case"',
              'Package Level 2: Check Ship Round Qty and Store Order Pack = "Quantity of Eaches in Package - Case"',
              'Package Level 2: Check Qty per Case = "Quantity of Eaches in Package - Case"',
              'Package Level 2: Check Qty per Case and Store Order Pack = "Quantity of Eaches in Package - Case"',
              'Package Level 2: Check Qty per Case and Ship Round Qty = "Quantity of Eaches in Package - Case"',
              'Package Level 2: Check Qty per Case, Ship Round Qty and Store Order Pack = "Quantity of Eaches in Package - Case"',
              'Package Level 3: Ok',
              'Package Level 3: Store Order Pack = "Quantity of Eaches in Package - Inner Pack',
              'Package Level 3: Ship Round Qty = "Quantity of Eaches in Package - Inner Pack',
              'Package Level 3: Ship Round Qty and Store Order Pack = "Quantity of Eaches in Package - Inner Pack',
              'Package Level 3: Check Qty per Case = "Quantity of Eaches in Package - Case"',
              'Package Level 3: Check Qty per Case = "Quantity of Eaches in Package - Case" and Store Order Pack = "Quantity of Eaches in Package - Inner Pack',
              'Package Level 3: Check Qty per Case = "Quantity of Eaches in Package - Case" and Ship Round Qty = "Quantity of Eaches in Package - Inner Pack',
              'Package Level 3: Check Qty per Case = "Quantity of Eaches in Package - Case", Ship Round Qty and Store Order Pack = "Quantity of Eaches in Package - Inner Pack'
       ]
       data['Quantities Validations'] = np.select(conditions_validations,options_validations, default='No Package Level data')

       #GTIN validations
       conditions_gtin = [
              package_level_1 & each_gtin,
              package_level_1 & (each_gtin == False),
              package_level_2 & each_gtin & case_gtin,
              package_level_2 & each_gtin & (case_gtin == False),
              package_level_2 & (each_gtin == False) & case_gtin,
              package_level_3 & each_gtin & case_gtin & inner_gtin & case_inner_gtin,
              package_level_3 & each_gtin & case_gtin & inner_gtin & (case_inner_gtin == False),
              package_level_3 & each_gtin & case_gtin & (inner_gtin == False) & case_inner_gtin,
              package_level_3 & each_gtin & case_gtin & (inner_gtin == False) & (case_inner_gtin == False),
              package_level_3 & each_gtin & (case_gtin == False) & inner_gtin & case_inner_gtin,
              package_level_3 & each_gtin & (case_gtin == False) & inner_gtin & (case_inner_gtin == False),
              package_level_3 & each_gtin & (case_gtin == False) & (inner_gtin == False) & case_inner_gtin,
              package_level_3 & each_gtin & (case_gtin == False) & (inner_gtin == False) & (case_inner_gtin == False),
              package_level_3 & (each_gtin == False) & case_gtin & inner_gtin & case_inner_gtin,
              package_level_3 & (each_gtin == False) & case_gtin & inner_gtin & (case_inner_gtin == False),
              package_level_3 & (each_gtin == False) & case_gtin & (inner_gtin == False) & case_inner_gtin,
              package_level_3 & (each_gtin == False) & case_gtin & (inner_gtin == False) & (case_inner_gtin == False),
              package_level_3 & (each_gtin == False) & (case_gtin == False) & inner_gtin & case_inner_gtin,
              package_level_3 & (each_gtin == False) & (case_gtin == False) & inner_gtin & (case_inner_gtin == False),
              package_level_3 & (each_gtin == False) & (case_gtin == False) & (inner_gtin == False) & case_inner_gtin,
              package_level_3 & (each_gtin == False) & (case_gtin == False) & (inner_gtin == False) & (case_inner_gtin == False)
       ]
       options_gtin = [
              'Pck Lvl 1: Ok',
              'Check Each GTIN =! Item GTIN',
              'Pck Lvl 2: Ok',
              'Check Case GTIN = Each GTIN',
              'Check Each GTIN =! Item GTIN',
              'Pck Lvl 3: Ok',
              'Check Inner Pack GTIN = Case GTIN',
              'Check Inner Pack GTIN = Each GTIN',
              'Check Inner Pack GTIN = Each & Case GTINs',
              'Check Case GTIN = Each GTIN',
              'Check Case GTIN = Each & Inner GTINs',
              'Check Case & Inner GTINs = Each GTIN',
              'Check Each, Case and Inner GTINs are all the same',
              'Check Each GTIN =! Item GTIN',
              'Check Each GTIN =! Item GTIN & Inner Pack GTIN = Case GTIN',
              'Check Each GTIN =! Item GTIN & Inner Pack GTIN = Each GTIN',
              'Check Each GTIN =! Item GTIN & Inner Pack GTIN = Each & Case GTINs',
              'Check Each GTIN =! Item GTIN & Case GTIN = Each GTIN',
              'Check Each GTIN =! Item GTIN & Case GTIN = Each & Inner GTINs',
              'Check Each GTIN =! Item GTIN & Case & Inner GTINs = Each GTIN',
              'Check Each GTIN =! Item GTIN & Each, Case and Inner GTINs are all the same'
       ]
       data['GTIN flag'] = np.select(conditions_gtin,options_gtin, default = 'No Package Level data. Check GTIN fields for possible errors')

       #Weight validations
       conditions_weight = [
              package_level_1,
              package_level_2 & case_weight_validation,
              package_level_2 & (case_weight_validation == False),
              package_level_3 & inner_weight_validation,
              package_level_3 & (inner_weight_validation == False)
       ]
       options_weight = [
              'Package lvl 1: Ok',
              'Package lvl 2: Ok',
              'Package lvl 2: Check',
              'Package lvl 3: Ok',
              'Package lvl 3: Check'
       ]
       data['Weight flag'] = np.select(conditions_weight,options_weight, default = 'No Package Level data. Check weights for possible errors')

       #Volume validations
       conditions_volume = [
              package_level_1,
              package_level_2 & case_volume_validation,
              package_level_2 & (case_volume_validation == False),
              package_level_3 & inner_volume_validation,
              package_level_3 & (inner_volume_validation == False)
       ]
       options_volume = [
              'Package lvl 1: Ok',
              'Package lvl 2: Ok',
              'Package lvl 2: Check',
              'Package lvl 3: Ok',
              'Package lvl 3: Check'
       ]
       data['Volume flag'] = np.select(conditions_volume,options_volume, default = 'No Package Level data. Check volumes for possible errors')
       return
		
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
			mdm = pd.read_excel(f'{filename}', index_col = False,sheet_name='Entities',header=1,dtype=object)
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
			sql = ("""
					SELECT DISTINCT 
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
					FROM [db_product_integrity].[prod].[tb_sku_warehouse_status_vendor] A
					LEFT JOIN db_product_integrity.prod.tb_az_vendors V
					ON A.vendor_id = V.VENDOR 
					LEFT JOIN db_product_integrity.prod.tb_item_vendor_rank B
					ON A.sku = B.item
						AND A.vendor_id = b.vendor_id
					WHERE  A.INTL_CODE = 'USA' AND
						A.SKU IN ({})
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
	window2 = Tk()
	window2.title("Data Governance: Step 2")
	window2.geometry("200x100")
	label_file_explorer = Label(window2, text = "Data Governance: Step 2", width = 100, height = 4).pack()
	button_explore = Button(window2, text = "Exportar archivo", command = export_file).pack()
	window2.mainloop()
sys.exit(0)