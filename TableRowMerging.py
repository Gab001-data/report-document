from docx import Document

# Sample data
risks = [
    {
        "id": "R1",
        "description": "Non-adherence to approved Preventive Maintenance Plan (PMP)...",
        "rating": "VH(16)",
        "sites": [
            {
                "site": "Nairobi",
                "findings": [
                    {"finding": "No preventive maintenance...", "ref": "8", "opinion": "Inadequate design"},
                    {"finding": "Lack of equipment register...", "ref": "9", "opinion": "Appropriate design but operating ineffectively"},
                ]
            },
            {
                "site": "Head Office",
                "findings": [
                    {"finding": "No proper training schedule...", "ref": "10", "opinion": "Appropriate design but operating ineffectively"},
                    {"finding": "Underutilisation of SAP maintenance modules", "ref": "11", "opinion": ""},
                ]
            }
        ]
    }
]

doc = Document()
table = doc.add_table(rows=1, cols=7)
table.style = 'Light Grid'
hdr = table.rows[0].cells
hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text, hdr[4].text, hdr[5].text, hdr[6].text = \
    ["Risk ID", "Description", "Rating", "Site", "Finding", "Ref", "Opinion"]

for risk in risks:
    risk_start = len(table.rows)  # remember the first row index for this risk
    is_start_of_risk = True
    for site in risk["sites"]:
        site_start = len(table.rows)  # remember the first row index for this site
        is_start_of_site = True
        for f in site["findings"]:
            row = table.add_row().cells
            row[0].text = risk["id"] if is_start_of_risk else ''
            row[1].text = risk["description"] if is_start_of_risk else ''
            row[2].text = risk["rating"] if is_start_of_risk else ''
            row[3].text = site["site"] if is_start_of_site else ''
            row[4].text = f["finding"]
            row[5].text = f["ref"]
            row[6].text = f["opinion"]
            is_start_of_risk = False
            is_start_of_site = False

        # merge Site column cells for this site
        site_rows = table.rows[site_start:len(table.rows)]
        if len(site_rows) > 1:
            site_rows[0].cells[3].merge(site_rows[-1].cells[3])

    # merge Risk ID / Description / Rating cells across all rows of this risk
    risk_rows = table.rows[risk_start:len(table.rows)]
    if len(risk_rows) > 1:
        risk_rows[0].cells[0].merge(risk_rows[-1].cells[0])
        risk_rows[0].cells[1].merge(risk_rows[-1].cells[1])
        risk_rows[0].cells[2].merge(risk_rows[-1].cells[2])

doc.save("output/risk_report_python_docx_merged.docx")
