import os
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
import json
import csv
import pandas as pd
from typing import List, Any


def safe_get(dictionary, keys, default="N/A"):
    """
    Safely get a nested dictionary value.

    Args:
        dictionary (dict or None): The dictionary to extract from.
        keys (list): A list of keys to traverse.
        default (str): Default value if any key is missing or None.

    Returns:
        The extracted value or default.
    """
    if dictionary is None:
        return default

    for key in keys:
        dictionary = dictionary.get(key) if isinstance(dictionary, dict) else None
        if dictionary is None:
            return default

    return dictionary


def extract_peer_reviews(peer_reviews):
    """
    Extracts peer review details from ORCID data.

    Args:
        peer_reviews (list): List of peer review data.

    Returns:
        list: A list of dictionaries containing peer review details.
    """
    peer_review_entries = []

    for review in peer_reviews:
        peer_review_groups = review.get("peer-review-group", [])

        for group in peer_review_groups:
            summaries = group.get("peer-review-summary", [])

            for summary in summaries:
                # Extract external IDs safely
                external_ids = safe_get(summary, ["external-ids", "external-id"], [])
                external_id_values = [safe_get(eid, ["external-id-value"]) for eid in external_ids]
                external_id_str = ", ".join(external_id_values) if external_id_values else "N/A"

                # Extract completion date safely
                year = safe_get(summary, ["completion-date", "year", "value"])
                month = safe_get(summary, ["completion-date", "month", "value"])
                day = safe_get(summary, ["completion-date", "day", "value"])

                peer_review_entry = {
                    'Peer Review Source': safe_get(summary, ["source", "source-name", "value"]),
                    'Peer Review External ID': external_id_str,
                    'Peer Review Completion Date': f"{year}-{month}-{day}",
                    'Peer Review Organization': safe_get(summary, ["convening-organization", "name"])
                }

                peer_review_entries.append(peer_review_entry)

    return peer_review_entries


def extract_funding_info(fundings):
    """
    Extracts funding details from ORCID data.

    Args:
        fundings (list): List of funding data.

    Returns:
        list: A list of dictionaries containing funding details.
    """
    funding_entries = []

    for funding in fundings:
        funding_summaries = funding.get("funding-summary", [])

        for summary in funding_summaries:
            # Extract external IDs safely
            external_ids = safe_get(summary, ["external-ids", "external-id"], [])
            external_id_values = [safe_get(eid, ["external-id-value"]) for eid in external_ids]
            grant_number = external_id_values[0] if external_id_values else "N/A"

            # Extract grant URL safely
            grant_url = safe_get(external_ids[0], ["external-id-url", "value"]) if external_ids else "N/A"

            funding_entry = {
                'Source': safe_get(summary, ["source", "source-name", "value"]),
                "Title": safe_get(summary, ["title", "title", "value"]),
                'Type': safe_get(summary, ["type"]),
                'Grant Number': grant_number,
                'Grant URL': grant_url,
                'Start Year': safe_get(summary, ["start-date", "year", "value"]),
                'End Year': safe_get(summary, ["end-date", "year", "value"], "Present"),
                'Organization': safe_get(summary, ["organization", "name"]),
                'Organization City': safe_get(summary, ["organization", "address", "city"])
            }

            funding_entries.append(funding_entry)

    return funding_entries


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
        f"{orcid_res.family_name if orcid_res.family_name else ''}".strip() or "Not Available", body_style
    )
    story.append(name)
    story.append(Spacer(1, 20))

    last_modified_date = Paragraph(f"<b>Last Modified Date:</b> " f"{orcid_res.last_modify_date if orcid_res.last_modify_date else 'Not Available'}", body_style)
    story.append(last_modified_date)
    story.append(Spacer(1, 20))

    works_title = Paragraph("Works", heading_style)
    story.append(works_title)

    for i, work in enumerate(orcid_res.publications):
        work_title = Paragraph(f"<b>{i + 1}. Title:</b> {work.title if work.title else 'No Information found'}", body_style)
        story.append(work_title)

        try:
            doi_url = list(work.url.values())[0] if work.url else "No DOI URL found"
        except (AttributeError, KeyError):
            doi_url = "No DOI URL found"
        doi_paragraph = Paragraph(f"<b>DOI URL:</b> {doi_url}", body_style)
        story.append(doi_paragraph)

        pub_year = Paragraph(f"<b>Publication Year:</b> {work.publicationyear if work.publicationyear else 'Unknown'}", body_style)
        story.append(pub_year)
        pub_type = Paragraph(f"<b>Publication Type:</b> {work.publicationtype if hasattr(work, 'publicationtype') else 'Unknown'}", body_style)
        story.append(pub_type)

        citation_value = getattr(work, 'citation_value', None)
        if citation_value:
            citation_paragraph = Paragraph(f"<b>Citation:</b> {citation_value}", body_style)
            story.append(citation_paragraph)
        else:
            story.append(Paragraph("<b>Citation:</b> No Citation Found", body_style))

        story.append(Spacer(1, 10))

    story.append(Spacer(1, 20))

    education_title = Paragraph("Employment", heading_style)
    story.append(education_title)

    # Iterate over the 'summaries' which is inside the 'employments' list
    for k, employment_group in enumerate(orcid_res.employments):
        if employment_group is None:
            continue

        # Access the 'summaries' key which contains the actual employment details
        summaries = employment_group.get('summaries', [])

        for employment in summaries:
            employment_summary = employment.get('employment-summary', {})

            department_name = employment_summary.get('department-name')
            if department_name is not None:
                department_paragraph = Paragraph(f"<b>Department:</b> {department_name}", body_style)
                story.append(department_paragraph)

            role_title = employment_summary.get('role-title')
            if role_title is not None:
                role_paragraph = Paragraph(f"<b>Role:</b> {role_title}", body_style)
                story.append(role_paragraph)

            organization = employment_summary.get('organization', {})
            org_name = organization.get('name') if organization else None
            if org_name is not None:
                org_paragraph = Paragraph(f"<b>Organization:</b> {org_name}", body_style)
                story.append(org_paragraph)

            address = organization.get('address', {}) if organization else {}
            city = address.get('city') if address else None
            if city is not None:
                city_paragraph = Paragraph(f"<b>Address:</b> {city}", body_style)
                story.append(city_paragraph)

            start_date = employment_summary.get('start-date', {})
            start_year = start_date.get('year', {}).get('value') if start_date else None
            if start_year is not None:
                start_year_paragraph = Paragraph(f"<b>Employment Start Year:</b> {start_year}", body_style)
                story.append(start_year_paragraph)
            story.append(Spacer(1, 10))
    story.append(Spacer(1, 20))  # Final spacing

    # Education and Qualifications Section
    education_title = Paragraph("Education and Qualifications", heading_style)
    story.append(education_title)

    education_number = len(orcid_res.educations)
    story.append(Paragraph(f"<b>Number of Education:</b> {education_number}", body_style))

    # Loop over the education entries
    for edu_l in range(education_number):
        if orcid_res.educations[edu_l] is None:
            continue

        # Access the 'summaries' key inside each education entry
        education_group = orcid_res.educations[edu_l]
        summaries = education_group.get('summaries', [])

        for edu in summaries:
            education_summary = edu.get('education-summary', {})

            story.append(Paragraph(f"<b>Education Details No:</b> {edu_l + 1}", body_style))

            # Get education role/title
            role_title = education_summary.get('role-title', 'Unknown Degree')
            story.append(Paragraph(f"<b>Education Role:</b> {role_title}", body_style))

            department_name = education_summary.get('department-name')
            if department_name is not None:
                department_paragraph = Paragraph(f"<b>Department:</b> {department_name}", body_style)
                story.append(department_paragraph)

            # Get education organization name
            organization = education_summary.get('organization', {})
            org_name = organization.get('name', 'Unknown Institution')
            story.append(Paragraph(f"<b>Education Organization:</b> {org_name}", body_style))

            # Get education start year
            start_date = education_summary.get('start-date', {})
            start_year = start_date.get('year', {}).get('value', 'Unknown Year') if start_date else None
            story.append(Paragraph(f"<b>Education Start Year:</b> {start_year}", body_style))

            # Get education end year
            end_date = education_summary.get('end-date', {})
            end_year = end_date.get('year', {}).get('value', 'Unknown Year') if end_date else None
            if end_year:
                story.append(Paragraph(f"<b>Education End Year:</b> {end_year}", body_style))
            else:
                story.append(Paragraph("<b>Education End Year:</b> Present", body_style))
            story.append(Spacer(1, 10))
    story.append(Spacer(1, 20))

    # Peer Review Section
    peer_review_title = Paragraph("Peer Reviews", heading_style)
    story.append(peer_review_title)

    for review in orcid_res.peer_reviews:
        peer_review_groups = review.get("peer-review-group", [])
        for group in peer_review_groups:
            summaries = group.get("peer-review-summary", [])
            for summary in summaries:
                source_name = summary.get("source", {}).get("source-name", {}).get("value", "N/A")
                external_id_value = (
                    summary.get("external-ids", {})
                    .get("external-id", [{}])[0]
                    .get("external-id-value", "N/A")
                )
                completion_year = (
                    summary.get("completion-date", {})
                    .get("year", {})
                    .get("value", "N/A")
                )
                organization_name = summary.get("convening-organization", {}).get("name", "N/A")

                # Add the formatted paragraphs to the story
                source_paragraph = Paragraph(f"<b>Source Name:</b> {source_name}", body_style)
                story.append(source_paragraph)

                external_id_paragraph = Paragraph(f"<b>External ID Value:</b> {external_id_value}", body_style)
                story.append(external_id_paragraph)

                completion_year_paragraph = Paragraph(f"<b>Completion Year:</b> {completion_year}", body_style)
                story.append(completion_year_paragraph)

                organization_paragraph = Paragraph(f"<b>Organization Name:</b> {organization_name}", body_style)
                story.append(organization_paragraph)

                # Add some space for separation between peer review items
                story.append(Spacer(1, 10))

    # Funding Section
    funding_title = Paragraph("Funding Information", heading_style)
    story.append(funding_title)

    for funding in orcid_res.fundings:
        funding_summaries = funding.get("funding-summary", [])
        for summary in funding_summaries:
            # Use if conditions to safely access each attribute
            funding_source = summary.get("source", {}).get("source-name", {}).get("value") if summary.get("source",{}).get("source-name") else "N/A"
            funding_title = summary.get("title", {}).get("title", {}).get("value") if summary.get("title") else "N/A"
            funding_type = summary.get("type") if summary.get("type") else "N/A"

            # Grant Number and URL
            funding_grant_number = "N/A"
            # Ensure 'external-ids' is not None and has the necessary structure
            external_ids = summary.get("external-ids")
            if external_ids is not None:
                external_id_list = external_ids.get("external-id", [])
                if isinstance(external_id_list, list) and len(external_id_list) > 0:
                    funding_grant_number = (
                        external_id_list[0].get("external-id-value")
                        if external_id_list[0].get("external-id-value") is not None
                        else "N/A"
                    )

            funding_grant_url = "N/A"
            # Check if 'external-ids' and 'external-id' have 'external-id-url'
            if external_ids and isinstance(external_ids, list) and len(external_ids) > 0:
                funding_grant_url = (
                    external_ids[0].get("external-id-url", {}).get("value")
                    if external_ids[0].get("external-id-url") is not None
                    else "N/A"
                )

            funding_start_year = (
                summary.get("start-date") and summary.get("start-date").get("year", {}).get("value", "N/A")
                if summary.get("start-date") is not None else "N/A"
            )

            funding_end_year = (
                summary.get("end-date", {}).get("year", {}).get("value")
                if summary.get("end-date") and summary["end-date"].get("year") is not None
                else "Present"
            )
            funding_organization = (
                summary.get("organization", {}).get("name") if summary.get("organization") is not None else "N/A"
            )
            funding_organization_city = (
                summary.get("organization", {}).get("address", {}).get("city") if summary.get("organization", {}).get("address") is not None else "N/A"
            )

            story.append(Paragraph(f"<b>Source:</b> {funding_source}", body_style))
            story.append(Paragraph(f"<b>Title:</b> {funding_title}", body_style))
            story.append(Paragraph(f"<b>Type:</b> {funding_type}", body_style))
            story.append(Paragraph(f"<b>Grant Number:</b> {funding_grant_number}", body_style))
            story.append(Paragraph(f"<b>Grant URL:</b> {funding_grant_url}", body_style))
            story.append(Paragraph(f"<b>Start Year:</b> {funding_start_year}", body_style))
            story.append(Paragraph(f"<b>End Year:</b> {funding_end_year}", body_style))
            story.append(Paragraph(f"<b>Organization:</b> {funding_organization}", body_style))
            story.append(Paragraph(f"<b>Organization City:</b> {funding_organization_city}", body_style))
            story.append(Spacer(1, 10))

    footer = Paragraph("Generated by OrcidXtract Generator", ParagraphStyle(
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

    peer_review_entries = extract_peer_reviews(orcid_res.peer_reviews)
    funding_entries = extract_funding_info(orcid_res.fundings)

    # Ensure the "Result" directory exists
    result_dir = os.path.join(os.getcwd(), "Result")
    os.makedirs(result_dir, exist_ok=True)

    # Enforce saving inside the "Result" directory
    output_file = os.path.join(result_dir, os.path.basename(output_file_name))

    with open(output_file, 'w', encoding='utf-8') as output:
        output.writelines("ORCID: " + (orcid_res.orcid if orcid_res.orcid else '') + "\n")
        output.writelines("Name: " + (f"{orcid_res.given_name} {orcid_res.family_name}".strip() if orcid_res.given_name or orcid_res.family_name else '') + "\n")
        output.writelines("Last Modified Date: " + (str(orcid_res.last_modify_date) if orcid_res.last_modify_date else 'Not Available') + "\n")
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

                publication_year = orcid_res.publications[i].publicationyear if orcid_res.publications[i].publicationyear else "Unknown"
                publication_type = orcid_res.publications[i].publicationtype if getattr(orcid_res.publications[i], 'publicationtype', None) else "Unknown"
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
        output.writelines("\nEmployment:\n")
        for k in range(employment_number):
            if orcid_res.employments[k] is None:
                continue

            # Access the summaries field to loop through the employment details
            summaries = orcid_res.employments[k].get('summaries', [])
            for employment in summaries:
                employment_summary = employment.get('employment-summary', {})

                # Extract and write department name, role, and organization
                department_name = employment_summary.get('department-name', 'Unknown Department')
                output.writelines("Department name: " + str(department_name) + "\n")

                role_title = employment_summary.get('role-title', 'Unknown Role')
                output.writelines("Role: " + str(role_title) + "\n")

                organization = employment_summary.get('organization', {})
                org_name = organization.get('name', 'Unknown Organization')
                output.writelines("Organization: " + str(org_name) + "\n")

                # Extract and write address (city)
                address = organization.get('address', {})
                city = address.get('city', 'Unknown City')
                output.writelines("Address: " + str(city) + "\n")

                # Extract and write start year
                start_date = employment_summary.get('start-date', {})
                start_year = start_date.get('year', {}).get('value', 'Unknown Year')
                output.writelines("Employment Start Year: " + start_year + "\n")

                output.writelines("\n")
        output.writelines("\n\n")

        output.writelines("Education and Qualifications: " + str(education_number) + "\n")
        for edu_l in range(education_number):
            if orcid_res.educations[edu_l] is None:
                continue

            # Access the summaries field to loop through the education details
            summaries = orcid_res.educations[edu_l].get('summaries', [])
            for edu in summaries:
                education_summary = edu.get('education-summary', {})

                output.writelines("Education Details No: " + str(edu_l + 1) + "\n")

                # Extract and write education role/title
                role_title = education_summary.get('role-title', 'Unknown Degree')
                output.writelines("Education role: " + str(role_title) + "\n")

                department_name = education_summary.get('department-name', 'Unknown Department')
                output.writelines("Department name: " + str(department_name) + "\n")

                # Extract and write education organization name
                organization = education_summary.get('organization', {})
                org_name = organization.get('name', 'Unknown Institution')
                output.writelines("Education organization: " + str(org_name) + "\n")

                # Extract and write start year
                start_date = education_summary.get('start-date', {})
                start_year = start_date.get('year', {}).get('value', 'Unknown Year')
                output.writelines("Education Start Year: " + start_year + "\n")

                # Extract and write end year
                end_date = education_summary.get('end-date', {})
                end_year = end_date.get('year', {}).get('value', 'Unknown Year') if end_date else None
                if end_year:
                    output.writelines("Education End Year: " + end_year + "\n")
                else:
                    output.writelines("Education End Year: Present\n")
                output.writelines("\n")
        output.writelines("\n\n")

        # Write Peer Review Information
        output.writelines("\nPeer Review Information:\n")
        if peer_review_entries:
            for entry in peer_review_entries:
                entry_str = "\n".join([f"{key}: {value}" for key, value in entry.items()])
                output.writelines(entry_str + "\n\n")
        else:
            output.writelines("No peer reviews found.\n")

        # Write Funding Information
        output.writelines("\nFunding Information:\n")
        if funding_entries:
            for entry in funding_entries:
                entry_str = "\n".join([f"{key}: {value}" for key, value in entry.items()])
                output.writelines(entry_str + "\n\n")
        else:
            output.writelines("No funding information found.\n")


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
                "department-name": (
                    emp.get('summaries', [{}])[0].get('employment-summary', {}).get('department-name', 'N/A')
                ),
                "role-title": (
                    emp.get('summaries', [{}])[0].get('employment-summary', {}).get('role-title', 'N/A')
                ),
                "organization": (
                    emp.get('summaries', [{}])[0].get('employment-summary', {}).get('organization', {}).get('name', 'N/A')
                ),
                "address": (
                    emp.get('summaries', [{}])[0].get('employment-summary', {}).get('organization', {}).get('address', {}).get('city', 'N/A')
                ),
                "start-date": (
                    (emp.get('summaries', [{}])[0].get('employment-summary', {}).get('start-date', {}) or {}).get('year', {}).get('value', 'N/A')
                )
            }
            for emp in orcid_res.employments if emp is not None
        ],
        "educations": [
            {
                "department-name": (
                    edu.get('summaries', [{}])[0].get('education-summary', {}).get('department-name', 'N/A')
                ),
                "role-title": (
                    edu.get('summaries', [{}])[0].get('education-summary', {}).get('role-title', 'Unknown Degree')
                ),
                "organization": (
                    edu.get('summaries', [{}])[0].get('education-summary', {}).get('organization', {}).get('name', 'Unknown Institution')
                ),
                "start-date": (
                    (edu.get('summaries', [{}])[0].get('education-summary', {}).get('start-date', {}) or {}).get('year', {}).get('value', 'Unknown Year')
                ),
                "end-date": (
                    (edu.get('summaries', [{}])[0].get('education-summary', {}).get('end-date', {}) or {}).get('year', {}).get('value', 'Present')
                    if edu.get('summaries', [{}])[0].get('education-summary', {}).get('end-date') else 'Present'
                )
            }
            for edu in orcid_res.educations if edu is not None
        ],
        "peer_reviews": [
            {
                "source_name": (
                    summary.get("source", {}).get("source-name", {}).get("value", "N/A")
                ),
                "external_id_value": (
                    summary.get("external-ids", {})
                    .get("external-id", [{}])[0]
                    .get("external-id-value", "N/A")
                ),
                "completion_year": (
                    summary.get("completion-date", {})
                    .get("year", {})
                    .get("value", "N/A")
                ),
                "organization_name": (
                    summary.get("convening-organization", {}).get("name", "N/A")
                )
            }
            for review in orcid_res.peer_reviews
            for peer_review_group in review.get("peer-review-group", [])
            for summary in peer_review_group.get("peer-review-summary", [])
        ],
        "fundings": [
            {
                "Source": (
                    summary.get("source", {}).get("source-name", {}).get("value")
                    if summary.get("source", {}).get("source-name") is not None
                    else "N/A"
                ),
                "Title": (
                    summary.get("title", {}).get("title", {}).get("value")
                    if summary.get("title") is not None
                    else "N/A"
                ),
                "Type": (
                    summary.get("type") if summary.get("type") is not None else "N/A"
                ),
                "Grant Number": (
                    summary.get("external-ids", {}).get("external-id", [{}])[0].get("external-id-value")
                    if summary.get("external-ids") is not None and summary.get("external-ids", {}).get("external-id", [{}])[0].get("external-id-value") is not None
                    else "N/A"
                ),
                "Grant URL": (
                    summary.get("external-ids", {}).get("external-id", [{}])[0].get("external-id-url", {}).get("value")
                    if summary.get("external-ids") is not None and summary.get("external-ids", {}).get("external-id", [{}])[0].get("external-id-url") is not None
                    else "N/A"
                ),
                "Start Year": (
                    summary.get("start-date") and summary.get("start-date").get("year", {}).get("value", "N/A")
                    if summary.get("start-date") is not None else "N/A"
                ),
                "End Year": (
                    summary.get("end-date") and summary.get("end-date").get("year", {}).get("value", "Present")
                    if summary.get("end-date") is not None else "Present"
                ),
                "Organization": (
                    summary.get("organization", {}).get("name")
                    if summary.get("organization") is not None
                    else "N/A"
                ),
                "Organization City": (
                    summary.get("organization", {})
                    .get("address", {})
                    .get("city")
                    if summary.get("organization", {}).get("address") is not None
                    else "N/A"
                )
            }
            for funding in orcid_res.fundings
            for summary in funding.get("funding-summary", [])
        ]
    }

    # Ensure the "Result" directory exists
    result_dir = os.path.join(os.getcwd(), "Result")
    os.makedirs(result_dir, exist_ok=True)

    # Enforce saving inside the "Result" directory
    output_file = os.path.join(result_dir, os.path.basename(output_file_name))

    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)


def create_report(orcid_data: List[Any], report_type: str) -> None:
    """
    Generates a CSV or Excel report for the given ORCID data.

    Args:
        orcid_data (List[Any]): A list of ORCID data objects.
        report_type (str): The type of report to generate ('csv' or 'excel').
    """
    report_rows = []

    fieldnames = [
        'ORCID', 'Name', 'Last Modified Date', 'Employments', 'Education', 'Publication Number',
        'Work Title', 'Work DOI URL', 'Work Publication Year', 'Work Publication Type', 'Work Citation',
        'Peer Review Source', 'Peer Review External ID', 'Peer Review Completion Year', 'Peer Review Organization',
        'Funding Source', 'Funding Title', 'Funding Type', 'Grant Number', 'Grant URL',
        'Funding Start Year', 'Funding End Year', 'Funding Organization', 'Funding Organization City'
    ]

    for orcid_res in orcid_data:
        # Extracting ORCID info
        orcid_info = {
            'ORCID': orcid_res.orcid if orcid_res.orcid else '',
            'Name': f"{orcid_res.given_name} {orcid_res.family_name}".strip() if orcid_res.given_name or orcid_res.family_name else '',
            'Last Modified Date': orcid_res.last_modify_date if orcid_res.last_modify_date else '',
            'Employments': ' | '.join([f"{emp.get('department-name', 'Unknown Department')}, "
                                       f"{emp.get('role-title', 'Unknown Role')}, "
                                       f"{emp.get('organization', {}).get('name', 'Unknown Organization')}, "
                                       f"{emp.get('organization', {}).get('address', {}).get('city', 'Unknown City')}, "
                                       f"{emp.get('start-date')['year']['value'] if emp.get('start-date') else 'Unknown Year'}"
                                       for emp in orcid_res.employments[:3]]) if orcid_res.employments else '',
            'Education': ' | '.join([f"{edu.get('role-title', 'Unknown Degree')}, "
                                     f"{edu.get('organization', {}).get('name', 'Unknown Institution')}, "
                                     f"{edu.get('start-date')['year']['value'] if edu.get('start-date') else 'Unknown Year'}, "
                                     f"{edu.get('end-date')['year']['value'] if edu.get('end-date') else 'Present'}"
                                     for edu in orcid_res.educations if edu]) if orcid_res.educations else '',
            'Publication Number': len(orcid_res.publications)
        }

        # Extract funding info
        funding_info_list = extract_funding_info(orcid_res.fundings) if hasattr(orcid_res, 'fundings') else []

        # Determine the maximum count to handle works, peer reviews, and fundings
        max_count = max(len(orcid_res.publications), len(orcid_res.peer_reviews), len(funding_info_list))

        for i in range(max_count):
            # Initialize dictionaries to hold individual work, review, and funding info
            work_info = {'Work Title': '', 'Work DOI URL': '', 'Work Publication Year': '', 'Work Publication Type': '',
                         'Work Citation': ''}
            review_info = {'Peer Review Source': '', 'Peer Review External ID': '', 'Peer Review Completion Year': '',
                           'Peer Review Organization': ''}
            funding_info = {'Funding Source': '', 'Funding Title': '', 'Funding Type': '', 'Grant Number': '',
                            'Grant URL': '',
                            'Funding Start Year': '', 'Funding End Year': '', 'Funding Organization': '',
                            'Funding Organization City': ''}

            # Add Work Info if available
            if i < len(orcid_res.publications):
                work = orcid_res.publications[i]
                work_info = {
                    'Work Title': work.title if work.title else 'No Information found',
                    'Work DOI URL': list(work.url.values())[0] if work.url else 'No DOI URL found',
                    'Work Publication Year': work.publicationyear if work.publicationyear else 'Unknown',
                    'Work Publication Type': work.publicationtype if hasattr(work, 'publicationtype') else 'Unknown',
                    'Work Citation': getattr(work, 'citation_value', 'No Citation Found')
                }

            # Add Peer Review Info if available (and handle multiple peer-review-groups)
            if i < len(orcid_res.peer_reviews):
                review = orcid_res.peer_reviews[i]
                peer_review_groups = review.get("peer-review-group", [])
                peer_review_sources = []
                peer_review_external_ids = []
                peer_review_completions = []
                peer_review_organizations = []

                for group in peer_review_groups:
                    summaries = group.get("peer-review-summary", [])
                    for summary in summaries:
                        peer_review_sources.append(summary.get("source", {}).get("source-name", {}).get("value", "N/A"))
                        peer_review_external_ids.append(' | '.join([ext.get("external-id-value", "N/A")
                                                                   for ext in summary.get("external-ids", {}).get("external-id", [])]))
                        peer_review_completions.append(str(summary.get("completion-date", {}).get("year", {}).get("value", "N/A")))
                        peer_review_organizations.append(summary.get("convening-organization", {}).get("name", "N/A"))

                review_info = {
                    'Peer Review Source': ' | '.join(peer_review_sources) if peer_review_sources else "N/A",
                    'Peer Review External ID': ' | '.join(peer_review_external_ids) if peer_review_external_ids else "N/A",
                    'Peer Review Completion Year': ' | '.join(peer_review_completions) if peer_review_completions else "N/A",
                    'Peer Review Organization': ' | '.join(peer_review_organizations) if peer_review_organizations else "N/A"
                }

            # Add Funding Info if available
            if i < len(funding_info_list):
                funding_info = {
                    'Funding Source': funding_info_list[i]['Source'],
                    'Funding Title': funding_info_list[i]['Title'],
                    'Funding Type': funding_info_list[i]['Type'],
                    'Grant Number': funding_info_list[i]['Grant Number'],
                    'Grant URL': funding_info_list[i]['Grant URL'],
                    'Funding Start Year': funding_info_list[i]['Start Year'],
                    'Funding End Year': funding_info_list[i]['End Year'],
                    'Funding Organization': funding_info_list[i]['Organization'],
                    'Funding Organization City': funding_info_list[i]['Organization City']
                }

            # Append the first row for ORCID info with Work, Peer Review, and Funding (on first loop)
            if i == 0:
                report_rows.append({**orcid_info, **work_info, **review_info, **funding_info})
            else:
                # For subsequent rows, leave ORCID ID and Name blank and only append work, review, and funding data
                report_rows.append({**{
                    'ORCID': '', 'Name': '', 'Last Modified Date': '', 'Employments': '', 'Education': '',
                    'Publication Number': ''
                }, **work_info, **review_info, **funding_info})

    # Determine the script's directory
    current_dir = os.getcwd()
    result_dir = os.path.join(current_dir, "Result")

    if not os.path.exists(result_dir):
        os.makedirs(result_dir)

    # Write to CSV
    if report_type == 'csv':
        output_file = os.path.join(result_dir, 'orcid_report.csv')
        with open(output_file, 'w', newline='', encoding='utf-8') as file:
            writer = csv.DictWriter(file, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(report_rows)

    # Write to Excel
    elif report_type == 'excel':
        output_file_excel = os.path.join(result_dir, 'orcid_report.xlsx')
        df = pd.DataFrame(report_rows)
        df.to_excel(output_file_excel, index=False)
