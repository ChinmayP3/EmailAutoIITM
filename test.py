import csv

first = True
with open("dataFinal.csv", mode ='r')as file:
    csvFile = csv.reader(file)
    for lines in csvFile:
        if(first):
            first = False
            continue
        company = lines[0].upper()
        company = lines[0].upper().strip()

        mail = lines[1]
        print("Mail will sent to "+company)