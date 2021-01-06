from docxtpl import DocxTemplate
import sys, os
import click
import csv
import datetime
from PyPDF2 import PdfFileMerger


def search_table(table, search_val):
    for row_number, row_data in enumerate(table):
        if search_val in table[row_number]:
            return True, row_number
    return False, None

def extract_from_csv(path_csv):
    # import pdb; pdb.set_trace()
    with open(path_csv) as csvDataFile:
        data = [row for row in csv.reader(csvDataFile)]

        # import pdb; pdb.set_trace()
  
        search_val = input("Search for: ")
        search_result = search_table(data,search_val)

        if search_result[0]:
            # print(search_result[1])
            print("")
        else:
            print("Not found.")
      

        # i=int(input("row number: "))
        i = search_result[1]

        rn_number = data[i][1]
        event_number = data[i][2]
        event_title = data[i][3]
        mpl_status = data[i][4]
        trades = data [i][5]
        date_sent_to_trades = data[i][6]
        rfq_due_date = data[i][7]
        lundy_markup = data[i][8]
        date_created = datetime.datetime.now().strftime("%Y-%m-%d")
        #field_change = input("field change (Yes/No): ")
        field_change = " "
        #schedule_impact_days = input("schedule impact days (#): ")
        schedule_impact_days = "0"
        total_amount = data[i][9]
        date_sent_to_owner = data[i][10]
        date_approved_by_owner = data[i][11]
        co_number = data[i][13]
        printed_date_today = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        trades_with_pricing = data[i][14]
        # import pdb; pdb.set_trace()

        return rn_number, event_number, event_title, mpl_status, trades, date_sent_to_trades, rfq_due_date, lundy_markup, date_created, field_change, schedule_impact_days, total_amount, date_sent_to_owner, date_approved_by_owner, co_number, printed_date_today, trades_with_pricing

def fill_doc():
    csv_data = extract_from_csv('revision.csv')
    # import pdb; pdb.set_trace()
    
    data = {
        
        "rn_number": csv_data[0],
        "event_number": csv_data[1],
        "event_title": csv_data[2],
        "mpl_status": csv_data[3],
        "trades": csv_data[4],
        "date_sent_to_trades": csv_data[5],
        "rfq_due_date": csv_data[6],
        "lundy_markup": csv_data[7],
        "date_created": csv_data[8],
        "field_change": csv_data[9],
        "schedule_impact_days": csv_data[10],
        "total_amount": csv_data[11],
        "date_sent_to_owner": csv_data[12],
        "date_approved_by_owner": csv_data[13],
        "co_number": csv_data[14],
        "printed_date_today": csv_data[15],
        "trades_with_pricing": csv_data[16]
        # "change_source_issued_date": csv_data[9]
    }
   
    doc = DocxTemplate("templates/pco_template.docx")
    context = data
    doc.render(context)
    # import pdb; pdb.set_trace()
    doc.save("output/" + data["rn_number"] + " -  PCO-" + data["event_number"] + "-" + data["event_title"] + ".docx")
   
    doc1 = DocxTemplate("templates/pricing_template.docx")
    context = data
    doc1.render(context)
    # import pdb; pdb.set_trace()
    doc1.save("output/" + data["rn_number"] + " -  pricing detail.docx")
   
    print("------------------------------")
    print("Aloha! PCO cover page has been created! - Paulette :)")
    print("------------------------------")
    print(csv_data[0])
    print(csv_data[1] + " " + csv_data[3] + " - " + csv_data[2])
    print("sent to: "+ csv_data[4])
    print("pricing: "+ csv_data[16])
    print(csv_data[11])
    print("------------------------------\n")

if __name__ == '__main__':
      fill_doc()
   