from docxtpl import DocxTemplate, InlineImage, RichText

doc = DocxTemplate('Report_templates/horizontal_merge_custom_tpl.docx')

doc.render({})

doc.save('output/horizontal_merger_report.docx')