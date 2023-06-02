import os
import openai
import xlsxwriter
import csv
import openpyxl


openai.api_key = "your openai key"
model_engine = "text-davinci-003"
prompt = input("Topic of spreadsheet:")

response = openai.Completion.create(
  engine=model_engine,
  prompt=prompt,
  max_tokens=1000,
  top_p=1,
  stop=None,
  temperature=0.6
)

r = response.choices[0].text
print("A:"+r)
print("Do you want to save this file?/n Say yes or no")

c = input()
if c=="yes":
    with open("spreadsheet.txt", 'wt') as f:
        print(r, file=f)
    input_file = 'spreadsheet.txt'
    output_file = 'spreadsheet.xlsx'
    wb = openpyxl.Workbook()
    ws = wb.worksheets[0]

    with open(input_file, 'r') as data:
        reader = csv.reader(data, delimiter='\t')
        for row in reader:
            if row != "|":
                ws.append(row)

        wb.save(output_file)