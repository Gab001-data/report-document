from docxtpl import DocxTemplate, InlineImage
from datetime import datetime as dt
from random import randint
import matplotlib.pyplot as plt

# import the template

doc = DocxTemplate('Report_templates/reportTmpl.docx')

#generate random sales data
sales = []
for idx, x in enumerate(range(9)):
    cPU = randint(0,9)+1
    unit_sold = randint(0,9)+5
    sale = {'name':f"Item {idx+1}", 'cPu':cPU, 'nUnits':unit_sold, 'revenue': cPU*unit_sold}
    sales.append(sale)

# get top 3 performing products

top3 =[item['name'] for item in sorted(sales, key=lambda x: x['revenue'], reverse=True)[:3] ]
print(top3)

# plot sales revenue

# Extract item names and revenues
items = [d["name"] for d in sales]
revenues = [d["revenue"] for d in sales]

# Plot bar chart
plt.bar(items, revenues)

# Add labels and title
plt.xlabel("Item")
plt.ylabel("Revenue")
plt.title("Revenue per Item")

plt.savefig('images/sales_visuals.png', dpi=300, bbox_inches='tight')

context = {
        'reportDtStr' : dt.strftime(dt.today(),'%Y-%m-%d'),
        'salesTblRows' : sales,
        'topItemsRows' : top3,
        'trendImg' : InlineImage(doc, 'images/sales_visuals.png')
}

doc.render(context)
doc.save('sale_report.docx')
