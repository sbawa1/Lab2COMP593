from sys import argv,exit
import os,re
from datetime import date
import pandas as ps


def get_sales_csv():
	if len(argv) >= 2:
		sales_csv = argv[1]
		if os.path.isfile(sales_csv):
			return sales_csv
		else:
			print("Error: CSV file does not exist")
			exit("Script execution aborted")
	else:
		print("Error: No CSV file path provided")
		exit("Script execution aborted")


def get_order_dir(sales_csv):
	sales_dir = os.path.dirname(sales_csv)
	todays_date = date.today().isoformat()
	order_dir_name = 'Orders_' + todays_date
	order_dir = os.path.join(sales_dir, order_dir_name)
	if not os.path.exists(order_dir):
		os.makedirs(order_dir)
	return order_dir

def split_sales_into_orders(sales_csv, order_dir):
	sales_df = ps.read_csv(sales_csv)
	sales_df.insert(7, 'TOTAL PRICE', sales_df['ITEM QUANTITY'] * sales_df['ITEM PRICE'])
	sales_df.drop(columns=['ADDRESS', 'CITY', 'STATE', 'POSTAL CODE', 'COUNTRY'], inplace=True)
	for order_id, order_df in sales_df.groupby('ORDER ID'):
		order_df.drop(columns=['ORDER ID'], inplace=True)
		order_df.sort_values(by='ITEM NUMBER', inplace=True)
		grand_total = order_df['TOTAL PRICE'].sum()
		grand_total_df = ps.DataFrame({'ITEM PRICE':['GRAND TOTAL'], 'TOTAL PRICE': [grand_total]})
		order_df = ps.concat([order_df, grand_total_df])
		customer_name = order_df['CUSTOMER NAME'].values[0]
		customer_name =re.sub(r'\W', '', customer_name)
		order_file_name = 'Order' + str(order_id) + '_' + customer_name + '.xlsx'
		order_file_path = os.path.join(order_dir, order_file_name)
		sheet_name = 'Order #' + str(order_id)
		order_df.to_excel(order_file_path, index=False, sheet_name=sheet_name)

sales_csv = get_sales_csv()
order_dir = get_order_dir(sales_csv)
split_sales_into_orders(sales_csv, order_dir)