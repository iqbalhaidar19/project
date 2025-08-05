import pandas as pd
import numpy as np

def clean_and_process_data(sales_filepath, categories_filepath):
    
    # LOADING DATA
    print("Read file data...")
    df = pd.read_csv(sales_filepath)
    df_categories = pd.read_csv(categories_filepath)
    
    # CLEANING DATA
    print("Initiate the data cleaning process...")
    
    # DEBUGGING A MISFORMATTED DATE
    # Convert to datetime and create a ‘mask’ for problematic rows (into NaT)
    tanggal_konversi = pd.to_datetime(df['Tanggal Transaksi'], errors='coerce')
    tanggal_bermasalah_mask = tanggal_konversi.isnull()
    
    # Check if there are any problematic rows
    if tanggal_bermasalah_mask.any():
        print("\n!!! FOUND ROWS WITH PROBLEMATIC DATE FORMAT !!!")
        print("Please check and correct the following lines in the raw_sales_data.csv file:")
        # Display rows from the ORIGINAL DataFrame whose dates are problematic
        print(df[tanggal_bermasalah_mask][['No. Invoice', 'Tanggal Transaksi']])
        print("="*60)

    # Remove duplicates
    df.drop_duplicates(inplace=True)
    
    # Standardize column names
    df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')
    
    # Missing Value
    df['nama_pelanggan'].fillna('Tidak Diketahui', inplace=True)
    df['nama_produk'].fillna('Produk Tidak Dikenal', inplace=True)
    df['status_pembayaran'].fillna('Status Tidak Jelas', inplace=True)
    
    # Standardize nama pelanggan
    df['nama_pelanggan'] = df['nama_pelanggan'].str.title()
    
    # Standardize nama produk
    print("Standardize the capitalization of Nama Produk...")
    df['nama_produk'] = df['nama_produk'].str.strip() 
    df['nama_produk'] = df['nama_produk'].str.title() 
    
    # Correction for acronyms that have changed
    df['nama_produk'] = df['nama_produk'].str.replace('Led', 'LED')
    df['nama_produk'] = df['nama_produk'].str.replace('Hd', 'HD')
    df['nama_produk'] = df['nama_produk'].str.replace('Gb', 'GB')
    df['nama_produk'] = df['nama_produk'].str.replace('Tb', 'TB')
    df['nama_produk'] = df['nama_produk'].str.replace('All-In-One', 'All-in-One')
    
    # Fix typos in the product
    product_corrections = {
        'Laptp Gaming': 'Laptop Gaming', 'Mose wireless': 'Mouse Wireless', 'Mouse wireless': 'Mouse Wireless',
        'keybord mechanical': 'Keyboard Mechanical', 'keyboard mechanical': 'Keyboard Mechanical'
    }
    df['nama_produk'] = df['nama_produk'].replace(product_corrections)

    # Standardize and convert harga satuan to numeric
    df['harga_satuan'] = df['harga_satuan'].astype(str).str.replace(r'[^0-9]', '', regex=True)
    df['harga_satuan'] = pd.to_numeric(df['harga_satuan'], errors='coerce').fillna(0)

    # Standardize status pembayaran
    status_mapping = {
        'LUNAS': 'LUNAS', 'Paid': 'LUNAS', 'lunas': 'LUNAS', 'completed': 'LUNAS',
        'Belum Lunas': 'BELUM LUNAS', 'Pending': 'BELUM LUNAS', 'Unpaid': 'BELUM LUNAS'
    }
    df['status_pembayaran'] = df['status_pembayaran'].str.strip().map(status_mapping).fillna('Tidak Diketahui')
    
    # Date conversion
    df['tanggal_transaksi'] = pd.to_datetime(df['tanggal_transaksi'], errors='coerce')
    df['tanggal_transaksi'] = df['tanggal_transaksi'].dt.date
    
    # Standardize No. Invoice
    df['no._invoice'] = df['no._invoice'].astype(str).str.upper().str.replace('-', '/')
    df.loc[df['no._invoice'].str.startswith('2024/'), 'no._invoice'] = 'INV/' + df['no._invoice']
    df.loc[df['no._invoice'].str.startswith('TRX/'), 'no._invoice'] = 'INV/2024/' + df['no._invoice'].str.split('/').str[-1]
    
    # Create a total_harga column
    df['total_harga'] = df['jumlah'] * df['harga_satuan']
    
    print("Add category information...")
    
    df_complete = pd.merge(df, df_categories, left_on='nama_produk', right_on='Nama Produk', how='left')
    df_complete['Kategori'].fillna('Lainnya', inplace=True)
    df_complete.drop(columns=['Nama Produk'], inplace=True)

    # Calculate commission
    conditions = [
        (df_complete['Kategori'] == 'Komputer') & (df_complete['total_harga'] > 10000000),
        (df_complete['Kategori'].isin(['Aksesoris Komputer', 'Audio']))
    ]
    choices = [ df_complete['total_harga'] * 0.05, df_complete['total_harga'] * 0.10 ]
    df_complete['komisi'] = np.select(conditions, choices, default=0)

    print("Data cleaning and processing completed.")
    return df_complete

def generate_excel_report(df_final, output_filename):
    """
    Multi-sheet Excel report with summaries and graphs.
    """
    print(f"Create an Excel report: {output_filename}...")
    # Create a sales summary per category
    summary = df_final.groupby('Kategori').agg(
        total_penjualan=('total_harga', 'sum'),
        total_komisi=('komisi', 'sum'),
        jumlah_transaksi=('no._invoice', 'count')
    ).reset_index()

    with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
        # Write data to each sheet
        df_final.to_excel(writer, sheet_name='Data Lengkap & Bersih', index=False)
        summary.to_excel(writer, sheet_name='Ringkasan Penjualan', index=False)
        
        # Get workbook and worksheet objects for formatting
        workbook  = writer.book
        worksheet1 = writer.sheets['Data Lengkap & Bersih']
        worksheet2 = writer.sheets['Ringkasan Penjualan']
        
        # Rupiah format and thousands separator with comma
        format_rupiah = workbook.add_format({'num_format': 'Rp #,##0'})
        
        header = df_final.columns.tolist()
        harga_satuan_idx = header.index('harga_satuan')
        total_harga_idx = header.index('total_harga')
        komisi_idx = header.index('komisi')
        
        # set_column(index_beginning, index_end, column_width, format)
        worksheet1.set_column(harga_satuan_idx, harga_satuan_idx, 15, format_rupiah)
        worksheet1.set_column(total_harga_idx, total_harga_idx, 18, format_rupiah)
        worksheet1.set_column(komisi_idx, komisi_idx, 15, format_rupiah)
        
        # Apply Formatting to Columns in Sheet 'Ringkasan Penjualan'
        worksheet2.set_column('B:C', 18, format_rupiah)
        
        # Creating a graph
        chart = workbook.add_chart({'type': 'column'})
        
        (max_row, _) = summary.shape
        chart.add_series({
            'name':       '=Ringkasan Penjualan!$B$1',
            'categories': ['Ringkasan Penjualan', 1, 0, max_row, 0],
            'values':     ['Ringkasan Penjualan', 1, 1, max_row, 1],
        })
        
        chart.set_title({'name': 'Total Penjualan per Kategori'})
        chart.set_x_axis({'name': 'Kategori Produk'})
        chart.set_y_axis({'name': 'Total Penjualan'})
        chart.set_legend({'position': 'none'})
        
        worksheet2.insert_chart('E2', chart)
        
    print("Report successfully created!")

# RUNNING ALL PROCESSES
if __name__ == "__main__":
    SALES_DATA_FILE = 'raw_sales_data.csv'
    CATEGORIES_DATA_FILE = 'product_categories.csv'
    OUTPUT_REPORT_FILE = 'Sales_Report.xlsx'
    
    processed_data = clean_and_process_data(SALES_DATA_FILE, CATEGORIES_DATA_FILE)
    generate_excel_report(processed_data, OUTPUT_REPORT_FILE)