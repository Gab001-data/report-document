from docxtpl import DocxTemplate

tpl = DocxTemplate("Report_templates/dynamic_table_tpl.docx")

context = {
    "col_labels": ["fruit", "vegetable", "stone", "thing"],
    "tbl_contents": [
        {"label": "yellow", "cols": ["banana", "capsicum", "pyrite", "taxi"]},
        {"label": "red", "cols": ["apple", "tomato", "cinnabar", "doubledecker"]},
        {"label": "green", "cols": ["guava", "cucumber", "aventurine", "card"]},
    ],
}

tpl.render(context)
tpl.save("output/dynamic_table.docx")
