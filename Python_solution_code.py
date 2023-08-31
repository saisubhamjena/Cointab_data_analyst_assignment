
import pandas as pd
import numpy as np


# Import all Excel Files

order=pd.read_excel("Company X - Order Report.xlsx")
pincode=pd.read_excel("Company X - Pincode Zones.xlsx")
sku=pd.read_excel("Company X - SKU Master.xlsx")
invoice=pd.read_excel("Courier Company - Invoice.xlsx")
rate=pd.read_excel("Courier Company - Rates.xlsx")

# Check duplicate

print('pincode',pincode.duplicated().sum())
print('sku',sku.duplicated().sum())
print('invoice',invoice.duplicated().sum())

# Delete duplicate values

pincode.drop_duplicates(inplace=True)
pincode.reset_index(drop=True,inplace=True)

sku.drop_duplicates(inplace=True)
sku.reset_index(drop=True,inplace=True)

# Capitalizing the Zone

pincode["Zone"]=pincode["Zone"].str.capitalize()
invoice["Zone"]=invoice["Zone"].str.capitalize()

# Renaming Coulmn
order.rename(columns={'ExternOrderNo': 'Order ID'}, inplace=True)

# Creating new calculation table

calculation= pd.DataFrame(invoice["Order ID"])
calculation["AWB Cade"] = invoice["AWB Code"]
calculation["Total Weight as per Courier Company (KG)"] = invoice["Charged Weight"]
calculation["Delivery Zone charged by courier company"] = invoice["Zone"]
calculation["Invoice Amount (Rs.)"] = invoice["Billing Amount (Rs.)"]
calculation["Customer Pincode"] = invoice["Customer Pincode"]
calculation["Type of Shipment"] = invoice["Type of Shipment"]

# merge sku with corrseponding weight
order=pd.merge(order,sku, on ='SKU', how ='left')

# Calculating total weight by quantity
order["Total Weight as per X (KG)"]=order["Order Qty"]*order["Weight (g)"]/1000

# deleting the extra row
order.drop('Weight (g)',axis=1,inplace=True)

# Calculating total weight groug by Order ID
Order_weight=order.groupby('Order ID')['Total Weight as per X (KG)'].sum()

# Converting to datafrme
Order_weight=Order_weight.reset_index()

# Merge with calculation table
calculation=pd.merge(calculation,Order_weight, on ='Order ID', how ='left')

# Merge pincode with zone
calculation=pd.merge(calculation,pincode, on ='Customer Pincode', how ='left')

# Deleting extra column
calculation.drop('Customer Pincode',axis=1,inplace=True)
calculation.drop('Warehouse Pincode',axis=1,inplace=True)

# Merge rate table
calculation=pd.merge(calculation,rate, on ='Zone', how ='left')

# calculating weight slab as per X
calculation['Aditional slab']= (calculation['Total Weight as per X (KG)']/calculation['Weight Slabs']).apply(np.ceil)-1
calculation['Weight slab as per X (KG)']= (calculation['Aditional slab']+1) * calculation['Weight Slabs']

# Function to calculate Expected charges as per X

def total_cost(row):
    if row['Type of Shipment'] == "Forward charges":
        return row['Forward Fixed Charge']+row['Aditional slab']*row['Forward Additional Weight Slab Charge']
    else:
        return row['Forward Fixed Charge']+row['Aditional slab']*row['Forward Additional Weight Slab Charge'] + row['RTO Fixed Charge']+row['Aditional slab']*row['RTO Additional Weight Slab Charge']
    
    
calculation['Expected Charge as per X (Rs.)'] = calculation.apply(total_cost,axis=1)

# Rename Column 
calculation.rename(columns={'Zone': 'Delivery Zone as per X'}, inplace=True)

# Deleting extra column
calculation.drop('Weight Slabs',axis=1,inplace=True)
calculation.drop('Forward Fixed Charge',axis=1,inplace=True)
calculation.drop('Forward Additional Weight Slab Charge',axis=1,inplace=True)
calculation.drop('RTO Fixed Charge',axis=1,inplace=True)
calculation.drop('RTO Additional Weight Slab Charge',axis=1,inplace=True)
calculation.drop('Aditional slab',axis=1,inplace=True)

# Rename Column to match the column name
rate.rename(columns={'Zone': 'Delivery Zone charged by courier company'}, inplace=True)

calculation=pd.merge(calculation,rate, on ='Delivery Zone charged by courier company', how ='left')

# calculating weight slab charged by courier company
calculation['Aditional slab']= (calculation['Total Weight as per Courier Company (KG)']/calculation['Weight Slabs']).apply(np.ceil)-1
calculation['Weight slab charged by Courier Company (KG)']= (calculation['Aditional slab']+1) * calculation['Weight Slabs']

# Function to calculate Charges Billed by Courier Company

def total_cost(row):
    if row['Type of Shipment'] == "Forward charges":
        return row['Forward Fixed Charge']+row['Aditional slab']*row['Forward Additional Weight Slab Charge']
    else:
        return row['Forward Fixed Charge']+row['Aditional slab']*row['Forward Additional Weight Slab Charge'] + row['RTO Fixed Charge']+row['Aditional slab']*row['RTO Additional Weight Slab Charge']
    

calculation['Charges Billed by Courier Company (Rs.)'] = calculation.apply(total_cost,axis=1)

# Deleting extra column
calculation.drop('Weight Slabs',axis=1,inplace=True)
calculation.drop('Forward Fixed Charge',axis=1,inplace=True)
calculation.drop('Forward Additional Weight Slab Charge',axis=1,inplace=True)
calculation.drop('RTO Fixed Charge',axis=1,inplace=True)
calculation.drop('RTO Additional Weight Slab Charge',axis=1,inplace=True)
calculation.drop('Aditional slab',axis=1,inplace=True)
calculation.drop('Type of Shipment',axis=1,inplace=True)

# calculating Difference Between Expected Charges and Billed Charges

calculation['Difference Between Expected Charges and Billed Charges (Rs.)']= calculation['Expected Charge as per X (Rs.)']-calculation['Charges Billed by Courier Company (Rs.)']

# Creating final calculation table

calculation_table= calculation[['Order ID', 'AWB Cade','Total Weight as per X (KG)', 'Weight slab as per X (KG)', 
                                'Total Weight as per Courier Company (KG)','Weight slab charged by Courier Company (KG)', 
                                'Delivery Zone as per X', 'Delivery Zone charged by courier company', 
                                'Expected Charge as per X (Rs.)', 'Charges Billed by Courier Company (Rs.)',
                                'Difference Between Expected Charges and Billed Charges (Rs.)']]

# Count values for summary table

charged_correctly_count = (calculation['Difference Between Expected Charges and Billed Charges (Rs.)'] == 0).sum()
over_charged_count = (calculation['Difference Between Expected Charges and Billed Charges (Rs.)'] < 0).sum()
under_charged_count = (calculation['Difference Between Expected Charges and Billed Charges (Rs.)'] > 0).sum()


# sum values for summary table

total_invoice_amount = calculation[calculation['Difference Between Expected Charges and Billed Charges (Rs.)']==0]['Invoice Amount (Rs.)'].sum()
total_overcharging_amount = calculation[calculation['Difference Between Expected Charges and Billed Charges (Rs.)']<0]['Difference Between Expected Charges and Billed Charges (Rs.)'].sum()
total_undercharging_amount = calculation[calculation['Difference Between Expected Charges and Billed Charges (Rs.)']>0]['Difference Between Expected Charges and Billed Charges (Rs.)'].sum()

# Creating summary table

data= { 'Count' :[charged_correctly_count,over_charged_count,under_charged_count],
       'Amount (Rs.)' : [total_invoice_amount,total_overcharging_amount,total_undercharging_amount]}

index=['Total orders where X has been correctly charged',
       'Total Orders where X has been overcharged',
       'Total Orders where X has been undercharged']

summary_table= pd.DataFrame(data,index=index)

# export to excel file

file_name=pd.ExcelWriter('Cointab_assignment_result.xlsx')

summary_table.to_excel(file_name, sheet_name='Summary')
calculation_table.to_excel(file_name, sheet_name='Calculation',index=False)

file_name.save()




