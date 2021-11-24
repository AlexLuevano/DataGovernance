from pydantic import (BaseModel, Field,ValidationError)
import pandas as pd
import numpy as np
from tkinter import *
from tkinter import filedialog
import openpyxl
import pyodbc
from typing import Optional,List
import sys
import csv

def browse_files():
       global filename
       filename = filedialog.askopenfilename(initialdir = "/Downloads", title = "Selecciona un archivo de MDM:", filetypes = ((".csv Files","*.csv*"),("all files","*.*")))
       if filename:
              l1 = Label(window, text = "File path: " + filename).pack()
       else:
              print('No seleccionaste ningún archivo.')
       window.destroy()
def export_file():

    export_file_path = filedialog.asksaveasfilename(defaultextension = '.xlsx',initialdir = "/Desktop", title = "Guardar archivo como:", filetypes = (("Excel Files","*.xlsx*"),("CSV Files","*.csv*"),("all files","*.*")))
    data.to_excel(export_file_path, index = False, header=True)
    window2.destroy()
def load_items(df):  
       raw_values = list(df.values)
       headers = list(df.columns)
       dict_raw_items = [dict(zip(headers,row)) for row in raw_values]

       items: List[Item] = [Item(**item) for item in dict_raw_items]
              #items = [Item.parse_obj(sku) for sku in dict_raw_items]
       return items
def colclass_parser(mdm):
       cols_orig = list(mdm.columns)
       cols = list(mdm.columns)
       for i in range(len(cols)):
              cols[i] = cols[i].lower()
              cols[i] = cols[i].replace(' ','_')

       cols_list = list(zip(cols,cols_orig))
       for mac in cols_list:
              print(f"'{mac[0]}': str = Field(alias = '{mac[1]}')")

class Item(BaseModel):

    action: str = Field(alias = 'Action', default=None)
    type: str = Field(alias = 'Type')
    id: str = Field(alias = 'ID')
    name: str = Field(alias = 'Name', default='_EMPTY')
    parts_classification: str = Field(alias = 'Parts Classification')
    brand_aaiaid: str = Field(alias = 'Brand AAIAID')
    part_number: str = Field(alias = 'Part Number')
    subbrand_aaiaid: str = Field(alias = 'SubBrand AAIAID')
    vendor_ownership_id: str = Field(alias = 'Vendor Ownership ID')
    apref: Optional[int] = Field(default = None, alias = 'APREF')
    buyer_group: Optional[str] = Field(default = None, alias = 'Buyer Group')
    stock_keeping_unit: Optional[int] = Field(default = None, alias = 'Stock Keeping Unit')
    vendor_dcs_action: Optional[str] = Field(default = None, alias = 'Vendor DCs.Action')
    vendor_dcs_dcs: Optional[str] = Field(default = None, alias = "Vendor DCs.DC's")
    vendor_dcs_pov_id: Optional[str] = Field(default = None, alias = 'Vendor DCs.POV ID')
    product_integrity_reference: Optional[str] = Field(default = None, alias = 'Product Integrity Reference')
    marketing_flag: Optional[str] = Field(default = None, alias = 'Marketing Flag')
    core_vendor_id: Optional[str] = Field(default = None, alias = 'Core Vendor Id')
    recall_vendor_id: Optional[str] = Field(default = None, alias = 'Recall Vendor Id')
    warranty_vendor_id: Optional[str] = Field(default = None, alias = 'Warranty Vendor Id')
    autozone_part_number: Optional[str] = Field(default = None, alias = 'AutoZone Part number')
    line_code: Optional[str] = Field(default = None, alias = 'Line Code')
    alternate_part_number: Optional[str] = Field(default = None, alias = 'Alternate Part Number')
    alternate_line_code: Optional[str] = Field(default = None, alias = 'Alternate Line Code')
    major_department: Optional[int] = Field(default = None, alias = 'Major Department')
    minor_department: Optional[int] = Field(default = None, alias = 'Minor Department')
    quantity_per_application: Optional[int] = Field(default = None, alias = 'Quantity per Application')
    quantity_per_application_uom: Optional[str] = Field(default = None, alias = 'Quantity per Application.UOM')
    quantity_per_application_qualifier: Optional[str] = Field(default = None, alias = 'Quantity per Application Qualifier')
    store_order_pack: Optional[str] = Field(default = None, alias = 'Store Order Pack')
    uom_for_edi_ordering: Optional[str] = Field(default = None, alias = 'UOM for EDI Ordering')
    store_credit_flag: Optional[str] = Field(default = None, alias = 'Store Credit Flag')
    oil_gal: Optional[int] = Field(default = None, alias = 'Oil gal')
    employee_discount: Optional[str] = Field(default = None, alias = 'Employee Discount')
    warranty: Optional[str] = Field(default = None, alias = 'Warranty')
    warranty_months: Optional[str] = Field(default = None, alias = 'Warranty Months')
    quantity_force_flag: Optional[str] = Field(default = None, alias = 'Quantity Force Flag')
    custom_id: Optional[str] = Field(default = None, alias = 'Custom Id')
    schedule_b: Optional[str] = Field(default = None, alias = 'Harmonized Tariff Code (Schedule B)')
    michigan_flag: Optional[str] = Field(default = None, alias = 'Michigan Flag')
    country_of_origin: Optional[str] = Field(default = None, alias = 'Country of Origin (Primary)')
    item_level_gtin: Optional[str] = Field(default = None, alias = 'Item-Level GTIN')
    item_level_gtin_qualifier: Optional[str] = Field(default = None, alias = 'Item-Level GTIN Qualifier')
    dimensions_each_action: Optional[str] = Field(default = None, alias = 'Dimensions - Each.Action')
    dimensions_each_merchandising_height: Optional[str] = Field(default = None, alias = 'Dimensions - Each.Merchandising Height')
    dimensions_each_merchandising_length: Optional[str] = Field(default = None, alias = 'Dimensions - Each.Merchandising Length')
    dimensions_each_merchandising_width: Optional[str] = Field(default = None, alias = 'Dimensions - Each.Merchandising Width')
    dimensions_each_shipping_height: Optional[str] = Field(default = None, alias = 'Dimensions - Each.Shipping Height')
    dimensions_each_shipping_length: Optional[str] = Field(default = None, alias = 'Dimensions - Each.Shipping Length')
    dimensions_each_shipping_width: Optional[str] = Field(default = None, alias = 'Dimensions - Each.Shipping Width')
    dimensions_each_uom: Optional[str] = Field(default = None, alias = 'Dimensions - Each.UOM')
    package_uom_each: Optional[str] = Field(default = None, alias = 'Package UOM - Each')
    package_level_gtin_each: Optional[str] = Field(default = None, alias = 'Package Level GTIN - Each')
    quantity_of_eaches_in_package_each: Optional[str] = Field(default = None, alias = 'Quantity of Eaches in Package - Each')
    weights_each_action: Optional[str] = Field(default = None, alias = 'Weights - Each.Action')
    weights_each_dimensional_weight: Optional[str] = Field(default = None, alias = 'Weights - Each.Dimensional Weight')
    weights_each_uom: Optional[str] = Field(default = None, alias = 'Weights - Each.UOM')
    weights_each_weight: Optional[str] = Field(default = None, alias = 'Weights - Each.Weight')
    dimensions_case_action: Optional[str] = Field(default = None, alias = 'Dimensions - Case.Action')
    dimensions_case_merchandising_height: Optional[str] = Field(default = None, alias = 'Dimensions - Case.Merchandising Height')
    dimensions_case_merchandising_length: Optional[str] = Field(default = None, alias = 'Dimensions - Case.Merchandising Length')
    dimensions_case_merchandising_width: Optional[str] = Field(default = None, alias = 'Dimensions - Case.Merchandising Width')
    dimensions_case_shipping_height: Optional[str] = Field(default = None, alias = 'Dimensions - Case.Shipping Height')
    dimensions_case_shipping_length: Optional[str] = Field(default = None, alias = 'Dimensions - Case.Shipping Length')
    dimensions_case_shipping_width: Optional[str] = Field(default = None, alias = 'Dimensions - Case.Shipping Width')
    dimensions_case_uom: Optional[str] = Field(default = None, alias = 'Dimensions - Case.UOM')
    package_uom_case: Optional[str] = Field(default = None, alias = 'Package UOM - Case')
    package_level_gtin_case: Optional[str] = Field(default = None, alias = 'Package Level GTIN - Case')
    quantity_of_eaches_in_package_case: Optional[str] = Field(default = None, alias = 'Quantity of Eaches in Package - Case')
    weights_case_action: Optional[str] = Field(default = None, alias = 'Weights - Case.Action')
    weights_case_dimensional_weight: Optional[str] = Field(default = None, alias = 'Weights - Case.Dimensional Weight')
    weights_case_weight: Optional[str] = Field(default = None, alias = 'Weights - Case.Weight')
    weights_case_weight_uom: Optional[str] = Field(default = None, alias = 'Weights - Case.Weight UOM')
    dimensions_inner_pack_action: Optional[str] = Field(default = None, alias = 'Dimensions - Inner Pack.Action')
    dimensions_inner_pack_merchandising_height: Optional[str] = Field(default = None, alias = 'Dimensions - Inner Pack.Merchandising Height')
    dimensions_inner_pack_merchandising_length: Optional[str] = Field(default = None, alias = 'Dimensions - Inner Pack.Merchandising Length')
    dimensions_inner_pack_merchandising_width: Optional[str] = Field(default = None, alias = 'Dimensions - Inner Pack.Merchandising Width')
    dimensions_inner_pack_shipping_height: Optional[str] = Field(default = None, alias = 'Dimensions - Inner Pack.Shipping Height')
    dimensions_inner_pack_shipping_length: Optional[str] = Field(default = None, alias = 'Dimensions - Inner Pack.Shipping Length')
    dimensions_inner_pack_shipping_width: Optional[str] = Field(default = None, alias = 'Dimensions - Inner Pack.Shipping Width')
    dimensions_inner_pack_uom: Optional[str] = Field(default = None, alias = 'Dimensions - Inner Pack.UOM')
    package_uom_inner_pack: Optional[str] = Field(default = None, alias = 'Package UOM - Inner Pack')
    package_level_gtin_inner_pack: Optional[str] = Field(default = None, alias = 'Package Level GTIN - Inner Pack')
    quantity_of_eaches_in_package_inner_pack: Optional[str] = Field(default = None, alias = 'Quantity of Eaches in Package - Inner Pack')
    weights_inner_pack_action: Optional[str] = Field(default = None, alias = 'Weights - Inner Pack.Action')
    weights_inner_pack_dimensional_weight: Optional[str] = Field(default = None, alias = 'Weights - Inner Pack.Dimensional Weight')
    weights_inner_pack_uom: Optional[str] = Field(default = None, alias = 'Weights - Inner Pack.UOM')
    weights_inner_pack_weight: Optional[str] = Field(default = None, alias = 'Weights - Inner Pack.Weight')
    ship_round_quantity: Optional[str] = Field(default = None, alias = 'Ship Round Quantity')
    quantity_per_case: Optional[str] = Field(default = None, alias = 'Quantity per Case')
    pallet_layer_maximum: Optional[str] = Field(default = None, alias = 'Pallet Layer Maximum')
    maximum_cases_per_pallet_layer: Optional[str] = Field(default = None, alias = 'Maximum Cases per Pallet Layer')
    pricing_action: Optional[str] = Field(default = None, alias = 'Pricing.Action')
    pricing_currency_code: Optional[str] = Field(default = None, alias = 'Pricing.Currency Code')
    pricing_expiration_date: Optional[str] = Field(default = None, alias = 'Pricing.Expiration Date')
    pricing_price: Optional[float] = Field(default = None, alias = 'Pricing.Price')
    pricing_price_uom: Optional[str] = Field(default = None, alias = 'Pricing.Price.UOM')
    pricing_price_break_quantity: Optional[str] = Field(default = None, alias = 'Pricing.Price Break Quantity')
    pricing_price_break_quantity_uom: Optional[str] = Field(default = None, alias = 'Pricing.Price Break Quantity.UOM')
    pricing_price_break_quantity_uom2: Optional[str] = Field(default = None, alias = 'Pricing.Price Break Quantity UOM')
    pricing_price_sheet_level_effective_date: Optional[str] = Field(default = None, alias = 'Pricing.Price Sheet Level Effective Date')
    pricing_price_sheet_number: Optional[str] = Field(default = None, alias = 'Pricing.Price Sheet Number')
    pricing_price_type: Optional[str] = Field(default = None, alias = 'Pricing.Price Type')
    pricing_price_type_description: Optional[str] = Field(default = None, alias = 'Pricing.Price Type Description')
    pricing_price_uom: Optional[str] = Field(default = None, alias = 'Pricing.Price UOM')
    retail_cost: Optional[float] = Field(default = None, alias = 'Retail Cost')
    core_retail_price: Optional[float] = Field(default = None, alias = 'Core Retail Price')
    product_description_long: Optional[str] = Field(default = None, alias = 'Product Description – Long')
    assign_az_upc: Optional[str] = Field(default = None, alias = 'Assign AZ UPC')
    assign_sku: Optional[str] = Field(default = None, alias = 'Assign SKU')
    supply_chain_analyst: Optional[str] = Field(default = None, alias = 'Supply Chain Analyst')
    contains_electronic_components: Optional[str] = Field(default = None, alias = 'Contains Electronic Components?')
    does_it_require_sds: Optional[str] = Field(default = None, alias = 'Does it require SDS?')
    is_item_a_chemical: Optional[str] = Field(default = None, alias = 'Is item a chemical?')
    is_or_contains_a_bulb: Optional[str] = Field(default = None, alias = 'Is or contains a bulb?')
    is_or_contains_a_battery: Optional[str] = Field(default = None, alias = 'Is or contains a battery?')

# Main

window = Tk()
window.title("Data Governance")
window.geometry("200x100")
label_file_explorer = Label(window, text = "Data Governance", width = 100, height = 4).pack()
button_explore = Button(window, text = "Buscar archivo", command = browse_files).pack()
window.mainloop()

print('Reading MDM file...')
try:
        mdm = pd.read_csv(f'{filename}', index_col = False,na_filter= False)
        skus_list = list(mdm['Stock Keeping Unit'])
        skus_list_string = str(skus_list)
        skus = skus_list_string[1:-1]
        print('MDM file read successfully')
except:
        raise Exception('Error at reading MDM file. Please check file and try again. Finishing execution.')
print('Connecting to Aurora...')
try:
       conn = pyodbc.connect(Driver='{SQL Server}',
                            Server='pr-aurora1-rs01',
                            Database='db_product_integrity',
                            Trusted_Connection='yes')
       print('Connected successfully to Aurora')
except:
       raise Exception ('Connection to Aurora failed. Ending execution.')
sql = ("""
       SELECT DISTINCT 
       CAST(A.sku AS VARCHAR) + '|' +  CAST(V.AP_REF AS VARCHAR) + '|' +  A.vendor_id AS [key2]
       , CAST(A.sku AS VARCHAR) + '|' +  CAST(V.AP_REF AS VARCHAR)  [key]
       , A.sku
       , V.AP_REF
       , A.vendor_id
       , A.warehouses, ISNULL(B.rank, 999) VENDOR_RANK
       , CASE WHEN A.vendor_id != A.core_return_vendor OR A.vendor_id != A.recall_vendor OR A.vendor_id != A.warranty_vendor THEN 'Yes' ELSE 'No' END [Diff Contracts]
       , LEN(A.warehouses) L
       FROM db_product_integrity.prod.tb_sku_warehouses A WITH(NOLOCK)
       LEFT JOIN db_product_integrity.prod.tb_az_vendors V WITH(NOLOCK)
       ON A.vendor_id = V.VENDOR 
       LEFT JOIN db_product_integrity.prod.tb_item_vendor_rank B WITH(NOLOCK)
       ON A.sku = B.item
              AND A.vendor_id = b.vendor_id
       WHERE  
       SKU IN({})
       ORDER BY A.sku, VENDOR_RANK, L DESC""").format(skus)
query = pd.read_sql(sql,conn)
print('Query run successful. Closing connection...')
conn.close()
print('Closed connection to Aurora successfully')

# Dictionaries creation

items = load_items(mdm)
print(items[0].buyer_group)
