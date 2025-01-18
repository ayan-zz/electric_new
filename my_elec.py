import os
import pdfplumber
import pandas as pd
import numpy as np
import re
import warnings

warnings.filterwarnings("ignore")

# Replace 'folder_path' with the path to your folder containing PDF files
folder_path = r'C:\\New folder\\dec2024'  
data_dicts = []

for filename in os.listdir(folder_path):
    if filename.endswith('.pdf'):
        pdf_file_path = os.path.join(folder_path, filename)
        invoice = filename.split('.pdf')

        with pdfplumber.open(pdf_file_path) as pdf:
            for page_number in range(len(pdf.pages)):
                page = pdf.pages[page_number]
                page_text = page.extract_text()

                data = re.findall(r'Connected Load\s*:\s*(\d+\.\d+)[\s\S]*?Outstanding amount\(Rs.\)#([\d.]+)', page_text)
                consumer_data = data[0] if data else ("", "")

                data1 = re.findall(r'ConsumerId\s*:\s*(\d+)[\s\S]*?', page_text)
                consumer_data1 = data1[0] if data1 else ""

                data2 = re.findall(r'LPSC Charge (\d+\.\d+)[\s\S]*?', page_text)
                consumer_data2 = data2[0] if data2 else "0"

                data3 = re.findall(r'Government Subsidy (-\d+\.\d+)[\s\S]*?', page_text)
                consumer_data3 = data3[0] if data3 else "0"

                data4 = re.findall(r'Total Charges \(Rs\) ([0-9.]+)[\s\S]*?', page_text)
                consumer_data4 = data4[0] if data4 else "0"

                data5 = re.findall(r'Installation No\s*:\s*(\d+)[\s\S]*?', page_text)
                consumer_data5 = data5[0] if data5 else "NA"

                data6 = re.findall(r'(?i)Bill Month\s*:\s*([A-Z][a-z]+[,.\s]*\d{4})', page_text)
                consumer_data6 = data6[0] if data6 else "NA"

                meter_reading_pattern = r'Meter No Time Previous Present Multiplying Unit Max Demand\s+Reading ' \
                                       r'Reading Factor\(MF\) consumed \(KVA\)\s+([A-Z0-9]+)\s+[A-Z]+ ' \
                                       r'([0-9.]+) ([0-9.]+) ([0-9.]+) ([0-9.]+)'
                matches = re.findall(meter_reading_pattern, page_text)
                for match in matches:
                    meter_no, previous_reading, present_reading, multi_fact, units_cons = match

                    if units_cons == '0':
                        pattern = r'Fixed/Demand Charge ([\d.]+)[\s\S]*?Total Charges \(Rs\) ([0-9.]+)[\s\S]*?' \
                                  r'Rental Charge\s*([\d.]+)'
                        charges_data = re.findall(pattern, page_text)
                        charges_data = charges_data[0] if charges_data else ("", "", "")

                        data_dicts.append({
                            "InvoiceNo": invoice[0],
                            "Installation_No": consumer_data5,
                            "ConsumerId": consumer_data1,
                            "Bill_Month": consumer_data6,
                            "Connected_Load": consumer_data[0],
                            "Outstanding_amount": consumer_data[1],
                            "Meter_No": meter_no,
                            "Previous_Reading": previous_reading,
                            "Present_Reading": present_reading,
                            "MF": multi_fact,
                            "Units_Cons": units_cons,
                            "Fixed/Demand Charge": charges_data[0],
                            "Rental_Charge": charges_data[2],
                            "LPSC": consumer_data2,
                            "Govt Subsidy": consumer_data3,
                            "Total_Charge": consumer_data4
                        })

                    else:
                        pattern2 = r'Energy Charge ([\d.]+)[\s\S]*?Total Charges \(Rs\) ([0-9.]+)[\s\S]*?' \
                                   r'Fixed/Demand Charge ([\d.]+)[\s\S]*?Rental Charge\s*([\d.]+)'
                        charges_data2 = re.findall(pattern2, page_text)
                        charges_data2 = charges_data2[0] if charges_data2 else ("", "", "", "")

                        data_dicts.append({
                            "InvoiceNo": invoice[0],
                            "Installation_No": consumer_data5,
                            "ConsumerId": consumer_data1,
                            "Bill_Month": consumer_data6,
                            "Connected_Load": consumer_data[0],
                            "Outstanding_amount": consumer_data[1],
                            "Meter_No": meter_no,
                            "Previous_Reading": previous_reading,
                            "Present_Reading": present_reading,
                            "MF": multi_fact,
                            "Units_Cons": units_cons,
                            "EnergyCharge": charges_data2[0],
                            "Fixed/Demand Charge2": charges_data2[2],
                            "Rental_Charge": charges_data2[3],
                            "LPSC": consumer_data2,
                            "Govt Subsidy": consumer_data3,
                            "Total_Charge": consumer_data4
                        })

combined_df = pd.DataFrame(data_dicts)

numeric_columns = ['InvoiceNo', 'Installation_No', 'ConsumerId',
                   'Connected_Load', 'Outstanding_amount', 'Previous_Reading',
                   'Present_Reading', 'MF', 'Units_Cons', 'Fixed/Demand Charge',
                   'Rental_Charge', 'LPSC', 'Govt Subsidy', 'Total_Charge', 'EnergyCharge',
                   'Fixed/Demand Charge2']
combined_df[numeric_columns] = combined_df[numeric_columns].apply(pd.to_numeric, errors="coerce")

df = combined_df.copy()
df['Fixed/Demand Charge'] = df['Fixed/Demand Charge'].fillna(df['Fixed/Demand Charge2'])
df = df.drop(columns=['Fixed/Demand Charge2'])
df['Rental_Charge'] = df['Rental_Charge'].fillna(25.0)

# Add new mapping to a 'scheme' column
new_dict = new_dict={213127978: "Natukjaykrishnapur", 213127933: "Beurgram", 213127980: "Chaksadi",
                     213069106: "Joynagar", 213127979: "Lakshmanpur", 213127981: "Dihrgagram", 
                     213073810: "Shyamsundarpur", 213073812: "Radhakantapur", 213073813: "Barasimulia",
                     213073814: "Isabpur", 213073811: "Dehipalsa", 213039142: "Guchati",
                     213107071: "Mohanpur", 213107072: "Raniganj", 213107069: "Jakrasirsha",
                     213103031: "Mamudpur", 213103032: "Tenthua", 212009260: "Tangaguria",
                     212134540: "Dasagram", 212134538: "Sabong", 212130389: "Jorhut",
                     212143465: "Maligram", 212144160: "Naya",  212130376: "Bhusulia",
                     212144161: "Tangur", 212129733: "Angua", 212129740: "Bejda",
                     212140614: "Zinsahar",  212172076: "Kusumda",  212172077: "Rampura",
                     212172085: "Sautia", 212009809: "Karkoda", 214012745: "Kantapal",
                     200325904: "Kespur", 212140603: "Mukshedpur", 213229548: "Narayanchak",
                     212009684: "Palaskhanda",222247278: "Sibracharmanpur",
                     200470432: "Kuldiha", 213017470: "Patharghata", 203055483: "Renjura",
                     212140730: "Shaymsundarpur", 212130380: "Makorda", 202809084: "Rambhadrapur",
                     212140715: "Jamda"}

df['scheme'] = df['ConsumerId'].map(new_dict)

# Save to Excel file
filename = 'df_excel_2024_new.xlsx'
df.to_excel(filename, index=False)

print(f"Data saved to {filename}")
