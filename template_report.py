from docxtpl import DocxTemplate, InlineImage
from datetime import datetime as dt

doc = DocxTemplate("Report_templates/inviteTmpl.docx")
today = dt.strftime(dt.now(),'%Y-%m-%d') 

# create context data
context = {
        'todayStr': today,
        'recipientName': 'John Adah',
        'evntDtStr': '2025-08-30',
        'venueStr' : 'Eleganza hotel',
        'senderName' : 'Gabriel',
        'bannerImg': InlineImage(doc,'images/party_banner_0.png')
}

doc.render(context)

doc.save('invite_doc.docx')