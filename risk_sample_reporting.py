from docxtpl import DocxTemplate

# Paths
template_path = "Report_templates/Risk_tmp.docx"
output_docx = "output/Risk_report_rendered.docx"

# Sample context with risks -> sites -> findings
context = {
    "risks": [
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
        },
        {
            "id": "R2",
            "description": "Incompleteness of the PMP may result in rejected claims by the Insurance company in case of breakdowns.",
            "rating": "VH(16)",
            "sites": [
                {
                    "site": "Head Office",
                    "findings": [
                        {"finding": "Missing procurement files", "ref": "13", "opinion": "Appropriate design but operating ineffectively"},
                        {"finding": "Failure to execute disposals of obsolete, unused and damaged items", "ref": "14", "opinion": ""},
                    ]
                }
            ]
        }
    ]
}

# Load and render the template
tpl = DocxTemplate(template_path)
tpl.render(context)
tpl.save(output_docx)

print(f"✅ Report generated: {output_docx}")
