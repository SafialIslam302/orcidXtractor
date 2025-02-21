import os
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
import json
import csv
import pandas as pd
from typing import List, Any


def create_pdf(output_file_name: str, orcid_res: Any) -> None:
    """
    Generates a PDF report for the given ORCID data.

    Args:
        output_file_name (str): The name of the output PDF file.
        orcid_res (Any): The ORCID data object.
    """
    # Ensure the "Result" directory exists
    result_dir = os.path.join(os.getcwd(), "Result")
    os.makedirs(result_dir, exist_ok=True)

    # Enforce saving inside the "Result" directory
    output_file = os.path.join(result_dir, os.path.basename(output_file_name))

    doc = SimpleDocTemplate(output_file, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []

    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Title'],
        fontSize=18,
        alignment=1,
        spaceAfter=20,
        textColor=colors.darkblue
    )

    heading_style = ParagraphStyle(
        'HeadingStyle',
        parent=styles['Heading2'],
        fontSize=14,
        spaceAfter=10,
        textColor=colors.darkgreen
    )

    body_style = ParagraphStyle(
        'BodyStyle',
        parent=styles['BodyText'],
        fontSize=12,
        spaceAfter=10,
        textColor=colors.black
    )

    title = Paragraph("ORCID Information", title_style)
    story.append(title)

    orcid_id = Paragraph(f"<b>ORCID:</b> "f"{orcid_res.orcid if orcid_res.orcid else 'Not Available'}", body_style)
    story.append(orcid_id)

    name = Paragraph(
        f"<b>Name:</b> {orcid_res.given_name if orcid_res.given_name else ''} "
        f"{orcid_res.family_name if orcid_res.family_name else ''}".strip() or "Not Available",
        body_style
    )
    story.append(name)
    story.append(Spacer(1, 20))

    last_modified_date = Paragraph(
        f"<b>Last Modified Date:</b> " f"{orcid_res.last_modify_date if orcid_res.last_modify_date else 'Not Available'}",
        body_style)
    story.append(last_modified_date)
    story.append(Spacer(1, 20))

    works_title = Paragraph("Works", heading_style)
    story.append(works_title)

    for i, work in enumerate(orcid_res.publications):
        work_title = Paragraph(f"<b>{i + 1}. Title:</b> {work.title if work.title else 'No Information found'}",
                               body_style)
        story.append(work_title)

        try:
            doi_url = list(work.url.values())[0] if work.url else "No DOI URL found"
        except (AttributeError, KeyError):
            doi_url = "No DOI URL found"
        doi_paragraph = Paragraph(f"<b>DOI URL:</b> {doi_url}", body_style)
        story.append(doi_paragraph)

        pub_year = Paragraph(f"<b>Publication Year:</b> {work.publicationyear if work.publicationyear else 'Unknown'}",
                             body_style)
        story.append(pub_year)
        pub_type = Paragraph(
            f"<b>Publication Type:</b> {work.publicationtype if hasattr(work, 'publicationtype') else 'Unknown'}",
            body_style)
        story.append(pub_type)

        citation_value = getattr(work, 'citation_value', None)
        if citation_value:
            citation_paragraph = Paragraph(f"<b>Citation:</b> {citation_value}", body_style)
            story.append(citation_paragraph)
        else:
            story.append(Paragraph("<b>Citation:</b> No Citation Found", body_style))

        story.append(Spacer(1, 10))

    story.append(Spacer(1, 20))

    education_title = Paragraph("Education and Qualifications", heading_style)
    story.append(education_title)

    for k, employment in enumerate(orcid_res.employments):
        if employment is None:
            continue

        department_name = employment.get('department-name')
        if department_name is not None:
            department_paragraph = Paragraph(f"<b>Department:</b> {department_name}", body_style)
            story.append(department_paragraph)

        role_title = employment.get('role-title')
        if role_title is not None:
            role_paragraph = Paragraph(f"<b>Role:</b> {role_title}", body_style)
            story.append(role_paragraph)

        organization = employment.get('organization', {})
        org_name = organization.get('name') if organization else None
        if org_name is not None:
            org_paragraph = Paragraph(f"<b>Organization:</b> {org_name}", body_style)
            story.append(org_paragraph)

        address = organization.get('address', {}) if organization else {}
        city = address.get('city') if address else None
        if city is not None:
            city_paragraph = Paragraph(f"<b>Address:</b> {city}", body_style)
            story.append(city_paragraph)

        start_date = employment.get('start-date', {})
        start_year = start_date.get('year', {}).get('value') if start_date else None
        if start_year is not None:
            start_year_paragraph = Paragraph(f"<b>Employment Start Year:</b> {start_year}", body_style)
            story.append(start_year_paragraph)
        story.append(Spacer(1, 10))
    story.append(Spacer(1, 20))

    footer = Paragraph("Generated by ORCID PDF Generator", ParagraphStyle(
        'FooterStyle',
        parent=styles['Normal'],
        fontSize=10,
        alignment=1,
        textColor=colors.grey
    ))
    story.append(Spacer(1, 20))
    story.append(footer)

    doc.build(story)


def create_txt(output_file_name: str, orcid_res: Any) -> None:
    """
    Creates a TXT report for a given ORCID record.

    Args:
        output_file_name: The directory to save the TXT file in.
        orcid_res (Any): The ORCID data object.
    """
    publication_number = len(orcid_res.publications)
    employment_number = len(orcid_res.employments)
    education_number = len(orcid_res.educations)

    # Ensure the "Result" directory exists
    result_dir = os.path.join(os.getcwd(), "Result")
    os.makedirs(result_dir, exist_ok=True)

    # Enforce saving inside the "Result" directory
    output_file = os.path.join(result_dir, os.path.basename(output_file_name))

    with open(output_file, 'w', encoding='utf-8') as output:
        output.writelines("ORCID: " + (orcid_res.orcid if orcid_res.orcid else '') + "\n")
        output.writelines("Name: " + (f"{orcid_res.given_name} {orcid_res.family_name}".strip() if orcid_res.given_name or orcid_res.family_name else '') + "\n")
        output.writelines(
            "Last Modified Date: " + (
                str(orcid_res.last_modify_date) if orcid_res.last_modify_date else 'Not Available') + "\n")
        output.writelines("\n")

        output.writelines("Number of Works: " + str(publication_number) + "\n")

        for i in range(publication_number):
            output.writelines("\n")
            output.writelines("Work Details No: " + str(i + 1) + "\n")
            try:
                if orcid_res.publications[i].title is None:
                    output.writelines("No Information found")
                else:
                    output.writelines("Paper title: " + orcid_res.publications[i].title + "\n")
                    try:
                        data = orcid_res.publications[i].url
                        doi_url = list(data.values())
                        output.writelines("Paper URL: " + doi_url[0] + "\n")
                    except (AttributeError, KeyError):
                        output.writelines("No Paper URL found.\n")

                publication_year = orcid_res.publications[i].publicationyear if orcid_res.publications[
                    i].publicationyear else "Unknown"
                publication_type = orcid_res.publications[i].publicationtype if getattr(orcid_res.publications[i],
                                                                                        'publicationtype',
                                                                                        None) else "Unknown"
                output.writelines("Publication Year: " + publication_year + "\n")
                output.writelines("Publication Type: " + publication_type + "\n")

                citation_value = getattr(orcid_res.publications[i], 'citation_value', None)
                if citation_value:
                    output.writelines("Citation: " + citation_value + "\n")
                else:
                    output.writelines("No Citation Found\n")
            except (ValueError, KeyError):
                output.writelines("No details found for work\n")

        output.writelines("\n\n")

        for k in range(employment_number):
            if orcid_res.employments[k] is None:
                continue

            if orcid_res.employments[k].get('end-date') is None:
                output.writelines("Current Employee\n")
            else:
                output.writelines("Past Employee\n")

            output.writelines(
                "Department name: " + str(orcid_res.employments[k].get('department-name', 'Unknown Department')) + "\n")
            output.writelines("Role: " + str(orcid_res.employments[k].get('role-title', 'Unknown Role')) + "\n")
            output.writelines("Organization: " + str(
                orcid_res.employments[k].get('organization', {}).get('name', 'Unknown Organization')) + "\n")
            output.writelines("Address: " + str(
                orcid_res.employments[k].get('organization', {}).get('address', {}).get('city',
                                                                                        'Unknown City')) + "\n")

            if orcid_res.employments[k].get('start-date') is not None:
                employment_start_year = orcid_res.employments[k].get('start-date', {}).get('year', {}).get('value',
                                                                                                           'Unknown Year')
                output.writelines("Employment Start Year: " + employment_start_year + "\n")

        output.writelines("\n\n")
        output.writelines("Number of Education: " + str(education_number) + "\n")

        for edu_l in range(education_number):
            if orcid_res.educations[edu_l] is None:
                continue

            output.writelines("Education Details No: " + str(edu_l + 1) + "\n")
            output.writelines(
                "Education role: " + str(orcid_res.educations[edu_l].get('role-title', 'Unknown Degree')) + "\n")
            output.writelines("Education organization: " + str(
                orcid_res.educations[edu_l].get('organization', {}).get('name', 'Unknown Institution')) + "\n")

            if orcid_res.educations[edu_l].get('start-date') is not None:
                education_start_year = orcid_res.educations[edu_l].get('start-date', {}).get('year', {}).get('value', 'Unknown Year')
                output.writelines("Education Start Year: " + education_start_year + "\n")

            if orcid_res.educations[edu_l].get('end-date') is not None:
                education_end_year = orcid_res.educations[edu_l].get('end-date', {}).get('year', {}).get('value', 'Unknown Year')
                output.writelines("Education End Year: " + education_end_year + "\n")
            else:
                output.writelines("Education End Year: Present\n")


def create_report(orcid_data: List[Any], report_type: str) -> None:
    """
    Generates a CSV or Excel report for the given ORCID data.

    Args:
        orcid_data (List[Any]): A list of ORCID data objects.
        report_type (str): The type of report to generate ('csv' or 'excel').
    """
    report_rows = []
    fieldnames = [
        'ORCID', 'Name', 'Employments', 'Education', 'Publication Number',
        'Work Title', 'Work DOI URL', 'Work Publication Year', 'Work Publication Type', 'Work Citation',
        'Last Modified Date'
    ]

    for orcid_res in orcid_data:
        orcid_info = {
            'ORCID': orcid_res.orcid if orcid_res.orcid else '',
            'Name': f"{orcid_res.given_name} {orcid_res.family_name}".strip() if orcid_res.given_name or orcid_res.family_name else '',
            'Employments': ' | '.join([
                f"{emp.get('department-name', 'Unknown Department')}, "
                f"{emp.get('role-title', 'Unknown Role')}, "
                f"{emp.get('organization', {}).get('name', 'Unknown Organization')}, "
                f"{emp.get('organization', {}).get('address', {}).get('city', 'Unknown City')}, "
                f"{emp.get('start-date')['year']['value'] if emp.get('start-date') else 'Unknown Year'}"
                for emp in orcid_res.employments[:3]
            ]) if orcid_res.employments else '',
            'Education': ' | '.join([
                f"{edu.get('role-title', 'Unknown Degree')}, "
                f"{edu.get('organization', {}).get('name', 'Unknown Institution')}, "
                f"{edu.get('start-date')['year']['value'] if edu.get('start-date') else 'Unknown Year'}, "
                f"{edu.get('end-date')['year']['value'] if edu.get('end-date') else 'Present'}"
                for edu in orcid_res.educations if edu
            ]) if orcid_res.educations else '',
            'Publication Number': len(orcid_res.publications),
            'Last Modified Date': orcid_res.last_modify_date if orcid_res.last_modify_date else ''
        }

        if orcid_res.publications:
            for idx, work in enumerate(orcid_res.publications):
                work_info = {
                    'Work Title': work.title if work.title else 'No Information found',
                    'Work DOI URL': list(work.url.values())[0] if work.url else 'No DOI URL found',
                    'Work Publication Year': work.publicationyear if work.publicationyear else 'Unknown',
                    'Work Publication Type': work.publicationtype if hasattr(work, 'publicationtype') else 'Unknown',
                    'Work Citation': getattr(work, 'citation_value', 'No Citation Found')
                }

                if idx == 0:
                    report_rows.append({**orcid_info, **work_info})
                else:
                    report_rows.append({
                        'ORCID': '',
                        'Name': '',
                        'Employments': '',
                        'Education': '',
                        'Publication Number': '',
                        'Last Modified Date': orcid_res.last_modify_date
                        if orcid_res.last_modify_date else '',
                        **work_info
                    })
        else:
            report_rows.append({**orcid_info, 'Work Title': '',
                                'Work DOI URL': '',
                                'Work Publication Year': '',
                                'Work Publication Type': '',
                                'Work Citation': ''})

    # Determine the script's directory
    current_dir = os.getcwd()
    result_dir = os.path.join(current_dir, "Result")

    if not os.path.exists(result_dir):
        os.makedirs(result_dir)

    if report_type == 'csv':
        output_file = os.path.join(result_dir, 'orcid_report.csv')
        with open(output_file, 'w', newline='', encoding='utf-8') as file:
            writer = csv.DictWriter(file, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(report_rows)
    elif report_type == 'excel':
        output_file = os.path.join(result_dir, 'orcid_report.xlsx')
        df = pd.DataFrame(report_rows)
        df.to_excel(output_file, index=False)


def create_json(output_file_name: str, orcid_res: Any) -> None:
    """
    Generates a JSON report for the given ORCID data.

    Args:
        output_file_name (str): The name of the output JSON file.
        orcid_res (Any): The ORCID data object.
    """
    data = {
        "orcid": orcid_res.orcid,
        "given_name": orcid_res.given_name,
        "family_name": orcid_res.family_name,
        "last_modify_date": orcid_res.last_modify_date,
        "publications": [
            {
                "title": work.title,
                "url": list(work.url.values())[0]
                if work.url else None, "publicationyear": work.publicationyear
                if work.publicationyear else None, "publicationtype": work.publicationtype
                if work.publicationtype else None, "citation_value": work.citation_value
                if work.citation_value else None
            }
            for work in orcid_res.publications
        ],
        "employments": [
            {
                "department-name": emp.get('department-name'),
                "role-title": emp.get('role-title'),
                "organization": emp.get('organization',
                                        {}).get('name'),
                "address": emp.get('organization', {}).
                get('address', {}).get('city'),
                "start-date": (emp.get('start-date')
                               or {}).get('year', {}).get('value')
            }
            for emp in orcid_res.employments
        ],
        "educations": [
            {
                "role-title": edu.get('role-title'),
                "organization": edu.get('organization',
                                        {}).get('name'),
                "start-date": (edu.get('start-date') or {}).get('year', {}).get('value'),
                "end-date": (edu.get('end-date') or {}).get('year', {}).get('value')
            }
            for edu in orcid_res.educations
        ]
    }

    # Ensure the "Result" directory exists
    result_dir = os.path.join(os.getcwd(), "Result")
    os.makedirs(result_dir, exist_ok=True)

    # Enforce saving inside the "Result" directory
    output_file = os.path.join(result_dir, os.path.basename(output_file_name))

    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
