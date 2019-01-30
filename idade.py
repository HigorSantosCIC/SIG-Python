import datetime
from openpyxl import load_workbook
wb2 = load_workbook('Pessoas.xlsx')
sheet = wb2.active
born = sheet['E']
age = sheet['F']
imc = sheet['G']
h = sheet['I']
w = sheet['J']
i = 0
ux = list()
for c in born[1:]:
    today = datetime.datetime.today()
    y,m,d = [int(x) for x in c.value.split('-')]
    ux.append(today.year - y - ((today.month, today.day) < (m, d)))

for e in age[1:]:
    e.value = ux[0]
    ux.pop(0)

ux = list()
ui = list()
for hx in h[1:]:
    ux.append(int(hx.value) / 100)
for hy in w[1:]:
    ui.append(int(hy.value))
i = 0
for hx in ux:
    ui[i] = round(ui[i] / (hx * hx), 2)
    i = i + 1

for ix in imc[1:]:
    ix.value = ui[0]
    ui.pop(0)

wb2.save("Pessoas.xlsx")   
    