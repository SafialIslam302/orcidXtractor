import csv
import json
import shutil
import unittest
import os
from typing import List
from unittest.mock import MagicMock

import pandas as pd

from .report_generator import create_txt, create_pdf, create_json, create_report
# from report_generator import create_txt, create_pdf, create_json, create_report


class TestCreateTxt(unittest.TestCase):

    def setUp(self):
        """Set up a mock ORCID response object before each test."""
        self.orcid_res = MagicMock()
        self.orcid_res.orcid = "0000-0001-2345-6789"
        self.orcid_res.family_name = "Doe"
        self.orcid_res.given_name = "John"
        self.orcid_res.last_modify_date = "2024-01-01"
        self.orcid_res.publications = [
            MagicMock(title="Research Paper 1", url={"DOI": "https://doi.org/10.1234/example"},
                      publicationyear="2020", publicationtype="Journal", citation_value="10 citations"),
            MagicMock(title=None, url={}, publicationyear=None, publicationtype=None)
        ]
        self.orcid_res.employments = [
            {"role-title": "Researcher", "organization": {"name": "XYZ University", "address": {"city": "New York"}},
             "start-date": {"year": {"value": "2015"}}, "end-date": None}
        ]
        self.orcid_res.educations = [
            {"role-title": "PhD", "organization": {"name": "ABC University"}, "start-date": {"year": {"value": "2010"}}}
        ]

        # Set up the output directory
        self.output_dir = os.path.join(os.getcwd(), "Result")
        os.makedirs(self.output_dir, exist_ok=True)
        self.output_file_name = os.path.join(self.output_dir, f"{self.orcid_res.orcid}.txt")

    def test_create_txt_file_exists(self):
        """Test if the TXT file is created successfully."""
        create_txt(self.output_file_name, self.orcid_res)
        self.assertTrue(os.path.exists(self.output_file_name))

    def test_create_txt_content(self):
        """Test if the TXT file contains expected content."""
        create_txt(self.output_file_name, self.orcid_res)
        with open(self.output_file_name, "r", encoding="utf-8") as f:
            content = f.read()

        self.assertIn("ORCID: 0000-0001-2345-6789", content)
        self.assertIn("Name: John Doe", content)
        self.assertIn("Last Modified Date: 2024-01-01", content)
        self.assertIn("Research Paper 1", content)
        self.assertIn("XYZ University", content)
        self.assertIn("ABC University", content)
        self.assertIn("PhD", content)

    def test_create_txt_missing_data(self):
        """Test if the function handles missing data correctly."""
        # Simulating missing data
        self.orcid_res.family_name = None
        self.orcid_res.given_name = None
        self.orcid_res.publications = []
        self.orcid_res.employments = []
        self.orcid_res.educations = []

        create_txt(self.output_file_name, self.orcid_res)
        with open(self.output_file_name, "r", encoding="utf-8") as f:
            content = f.read()

        self.assertIn("Name: ", content)  # Should be empty but still present
        self.assertIn("Number of Works: 0", content)
        self.assertIn("Number of Education: 0", content)

    def tearDown(self):
        """Clean up test files after each test."""
        if os.path.exists(self.output_file_name):
            os.remove(self.output_file_name)
        if os.path.exists(self.output_dir):
            shutil.rmtree(self.output_dir)


class TestCreatePdf(unittest.TestCase):

    def setUp(self):
        """Set up a mock ORCID data object for testing."""
        self.orcid_res = MagicMock()
        self.orcid_res.orcid = "0000-0001-2345-6789"
        self.orcid_res.given_name = "John"
        self.orcid_res.family_name = "Doe"
        self.orcid_res.last_modify_date = "2024-02-10"

        # Mock Publications
        publication = MagicMock()
        publication.title = "Research Paper 1"
        publication.url = {"doi": "https://doi.org/10.1234/example"}
        publication.publicationyear = "2023"
        publication.publicationtype = "Journal Article"
        publication.citation_value = "Doe, J. (2023). Research Paper 1."
        self.orcid_res.publications = [publication]

        # Mock Employments
        employment = {
            "department-name": "Computer Science",
            "role-title": "Professor",
            "organization": {
                "name": "Tech University",
                "address": {"city": "New York"}
            },
            "start-date": {"year": {"value": "2015"}}
        }
        self.orcid_res.employments = [employment]

        # Set up the output directory
        self.output_dir = os.path.join(os.getcwd(), "Result")
        os.makedirs(self.output_dir, exist_ok=True)
        self.output_file_name = os.path.join(self.output_dir, "test_orcid_report.pdf")

    def tearDown(self):
        """Clean up the test file after running tests."""
        if os.path.exists(self.output_file_name):
            os.remove(self.output_file_name)
        if os.path.exists(self.output_dir):
            shutil.rmtree(self.output_dir)

    def test_create_pdf_file_exists(self):
        """Test if the PDF file is successfully created."""
        create_pdf(self.output_file_name, self.orcid_res)
        self.assertTrue(os.path.exists(self.output_file_name))

    def test_create_pdf_with_no_publications(self):
        """Test PDF generation when no publications exist."""
        self.orcid_res.publications = []
        create_pdf(self.output_file_name, self.orcid_res)
        self.assertTrue(os.path.exists(self.output_file_name))

    def test_create_pdf_with_no_employment(self):
        """Test PDF generation when no employment details exist."""
        self.orcid_res.employments = []
        create_pdf(self.output_file_name, self.orcid_res)
        self.assertTrue(os.path.exists(self.output_file_name))

    def test_create_pdf_with_missing_fields(self):
        """Test handling missing fields in ORCID data."""
        self.orcid_res.orcid = None
        self.orcid_res.given_name = None
        self.orcid_res.family_name = None
        self.orcid_res.last_modify_date = None
        create_pdf(self.output_file_name, self.orcid_res)
        self.assertTrue(os.path.exists(self.output_file_name))

    def test_create_pdf_with_empty_orcid_object(self):
        """Test handling an empty ORCID object."""
        empty_orcid_res = MagicMock()
        empty_orcid_res.orcid = None
        empty_orcid_res.given_name = None
        empty_orcid_res.family_name = None
        empty_orcid_res.last_modify_date = None
        empty_orcid_res.publications = []
        empty_orcid_res.employments = []

        create_pdf(self.output_file_name, empty_orcid_res)
        self.assertTrue(os.path.exists(self.output_file_name))


class TestCreateJson(unittest.TestCase):

    def setUp(self):
        """Set up a mock ORCID data object for testing."""
        self.orcid_res = MagicMock()
        self.orcid_res.orcid = "0000-0001-2345-6789"
        self.orcid_res.given_name = "John"
        self.orcid_res.family_name = "Doe"
        self.orcid_res.last_modify_date = "2024-02-10"

        # Mock Publications
        publication = MagicMock()
        publication.title = "Research Paper 1"
        publication.url = {"doi": "https://doi.org/10.1234/example"}
        publication.publicationyear = "2023"
        publication.publicationtype = "Journal Article"
        publication.citation_value = "Doe, J. (2023). Research Paper 1."
        self.orcid_res.publications = [publication]

        # Mock Employments
        employment = {
            "department-name": "Computer Science",
            "role-title": "Professor",
            "organization": {
                "name": "Tech University",
                "address": {"city": "New York"}
            },
            "start-date": {"year": {"value": "2015"}}
        }
        self.orcid_res.employments = [employment]

        # Mock Educations
        education = {
            "role-title": "PhD Student",
            "organization": {"name": "Tech University"},
            "start-date": {"year": {"value": "2010"}},
            "end-date": {"year": {"value": "2014"}}
        }
        self.orcid_res.educations = [education]

        # Set up the output directory
        self.output_dir = os.path.join(os.getcwd(), "Result")
        os.makedirs(self.output_dir, exist_ok=True)
        self.output_file_name = os.path.join(self.output_dir, "test_orcid_report.json")

    def tearDown(self):
        """Clean up the test file and directory after running tests."""
        if os.path.exists(self.output_file_name):
            os.remove(self.output_file_name)
        if os.path.exists(self.output_dir):
            shutil.rmtree(self.output_dir)

    def test_create_json_file_exists(self):
        """Test if the JSON file is successfully created."""
        create_json(self.output_file_name, self.orcid_res)
        self.assertTrue(os.path.exists(self.output_file_name))

    def test_create_json_with_valid_data(self):
        """Test if JSON file contains correct structured data."""
        create_json(self.output_file_name, self.orcid_res)
        with open(self.output_file_name, 'r', encoding='utf-8') as f:
            data = json.load(f)

        self.assertEqual(data["orcid"], "0000-0001-2345-6789")
        self.assertEqual(data["given_name"], "John")
        self.assertEqual(data["family_name"], "Doe")
        self.assertEqual(data["publications"][0]["title"], "Research Paper 1")
        self.assertEqual(data["publications"][0]["url"], "https://doi.org/10.1234/example")
        self.assertEqual(data["employments"][0]["organization"], "Tech University")
        self.assertEqual(data["educations"][0]["role-title"], "PhD Student")

    def test_create_json_with_no_publications(self):
        """Test JSON generation when no publications exist."""
        self.orcid_res.publications = []
        create_json(self.output_file_name, self.orcid_res)
        with open(self.output_file_name, 'r', encoding='utf-8') as f:
            data = json.load(f)
        self.assertEqual(data["publications"], [])

    def test_create_json_with_no_employment(self):
        """Test JSON generation when no employment details exist."""
        self.orcid_res.employments = []
        create_json(self.output_file_name, self.orcid_res)
        with open(self.output_file_name, 'r', encoding='utf-8') as f:
            data = json.load(f)
        self.assertEqual(data["employments"], [])

    def test_create_json_with_no_education(self):
        """Test JSON generation when no education details exist."""
        self.orcid_res.educations = []
        create_json(self.output_file_name, self.orcid_res)
        with open(self.output_file_name, 'r', encoding='utf-8') as f:
            data = json.load(f)
        self.assertEqual(data["educations"], [])

    def test_create_json_with_missing_fields(self):
        """Test handling missing fields in ORCID data."""
        self.orcid_res.orcid = None
        self.orcid_res.given_name = None
        self.orcid_res.family_name = None
        self.orcid_res.last_modify_date = None
        create_json(self.output_file_name, self.orcid_res)
        with open(self.output_file_name, 'r', encoding='utf-8') as f:
            data = json.load(f)

        self.assertIsNone(data["orcid"])
        self.assertIsNone(data["given_name"])
        self.assertIsNone(data["family_name"])
        self.assertIsNone(data["last_modify_date"])

    def test_create_json_with_empty_orcid_object(self):
        """Test handling an empty ORCID object."""
        empty_orcid_res = MagicMock()
        empty_orcid_res.orcid = None
        empty_orcid_res.given_name = None
        empty_orcid_res.family_name = None
        empty_orcid_res.last_modify_date = None
        empty_orcid_res.publications = []
        empty_orcid_res.employments = []
        empty_orcid_res.educations = []

        create_json(self.output_file_name, empty_orcid_res)
        with open(self.output_file_name, 'r', encoding='utf-8') as f:
            data = json.load(f)

        self.assertIsNone(data["orcid"])
        self.assertEqual(data["publications"], [])
        self.assertEqual(data["employments"], [])
        self.assertEqual(data["educations"], [])


class TestCreateReport(unittest.TestCase):

    def setUp(self):
        """Set up mock ORCID data for testing."""
        self.orcid_res_1 = MagicMock()
        self.orcid_res_1.orcid = "0000-0001-2345-6789"
        self.orcid_res_1.given_name = "John"
        self.orcid_res_1.family_name = "Doe"
        self.orcid_res_1.last_modify_date = "2024-02-10"

        # Mock Publications
        publication1 = MagicMock()
        publication1.title = "Research Paper 1"
        publication1.url = {"doi": "https://doi.org/10.1234/example1"}
        publication1.publicationyear = "2023"
        publication1.publicationtype = "Journal Article"
        publication1.citation_value = "Doe, J. (2023). Research Paper 1."

        publication2 = MagicMock()
        publication2.title = "Research Paper 2"
        publication2.url = {"doi": "https://doi.org/10.1234/example2"}
        publication2.publicationyear = "2021"
        publication2.publicationtype = "Conference Paper"
        publication2.citation_value = "Doe, J. (2021). Research Paper 2."

        self.orcid_res_1.publications = [publication1, publication2]

        # Mock Employments
        employment = {
            "department-name": "Computer Science",
            "role-title": "Professor",
            "organization": {"name": "Tech University", "address": {"city": "New York"}},
            "start-date": {"year": {"value": "2015"}}
        }
        self.orcid_res_1.employments = [employment]

        # Mock Educations
        education = {
            "role-title": "PhD Student",
            "organization": {"name": "Tech University"},
            "start-date": {"year": {"value": "2010"}},
            "end-date": {"year": {"value": "2014"}}
        }
        self.orcid_res_1.educations = [education]

        self.orcid_res_2 = MagicMock()
        self.orcid_res_2.orcid = "0000-0002-9876-5432"
        self.orcid_res_2.given_name = "Alice"
        self.orcid_res_2.family_name = "Smith"
        self.orcid_res_2.last_modify_date = "2024-01-15"
        self.orcid_res_2.publications = []
        self.orcid_res_2.employments = []
        self.orcid_res_2.educations = []

        self.orcid_data: List[MagicMock] = [self.orcid_res_1, self.orcid_res_2]

        self.result_dir = os.path.join(os.getcwd(), "Result")
        os.makedirs(self.result_dir, exist_ok=True)

    def tearDown(self):
        """Clean up test files."""
        csv_file = os.path.join(self.result_dir, "orcid_report.csv")
        excel_file = os.path.join(self.result_dir, "orcid_report.xlsx")

        if os.path.exists(csv_file):
            os.remove(csv_file)
        if os.path.exists(excel_file):
            os.remove(excel_file)
        if os.path.exists(self.result_dir):
            shutil.rmtree(self.result_dir)

    def test_create_csv_report(self):
        """Test if the CSV report is generated correctly."""
        create_report(self.orcid_data, "csv")
        output_file = os.path.join(self.result_dir, "orcid_report.csv")

        self.assertTrue(os.path.exists(output_file))

        with open(output_file, 'r', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            rows = list(reader)

        self.assertEqual(len(rows), 3)  # 2 for John (2 publications), 1 for Alice
        self.assertEqual(rows[0]['ORCID'], "0000-0001-2345-6789")
        self.assertEqual(rows[0]['Name'], "John Doe")
        self.assertEqual(rows[0]['Work Title'], "Research Paper 1")
        self.assertEqual(rows[1]['Work Title'], "Research Paper 2")
        self.assertEqual(rows[2]['ORCID'], "0000-0002-9876-5432")  # Alice's ORCID

    def test_create_excel_report(self):
        """Test if the Excel report is generated correctly."""
        create_report(self.orcid_data, "excel")
        output_file = os.path.join(self.result_dir, "orcid_report.xlsx")

        self.assertTrue(os.path.exists(output_file))

        df = pd.read_excel(output_file)
        self.assertEqual(len(df), 3)  # 2 for John (2 publications), 1 for Alice
        self.assertEqual(df.iloc[0]['ORCID'], "0000-0001-2345-6789")
        self.assertEqual(df.iloc[0]['Name'], "John Doe")
        self.assertEqual(df.iloc[0]['Work Title'], "Research Paper 1")
        self.assertEqual(df.iloc[1]['Work Title'], "Research Paper 2")
        self.assertEqual(df.iloc[2]['ORCID'], "0000-0002-9876-5432")  # Alice's ORCID

    def test_create_csv_with_empty_orcid_list(self):
        """Test CSV generation when no ORCID data is provided."""
        create_report([], "csv")
        output_file = os.path.join(self.result_dir, "orcid_report.csv")

        self.assertTrue(os.path.exists(output_file))

        with open(output_file, 'r', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            rows = list(reader)

        self.assertEqual(len(rows), 0)  # Should be empty since no ORCID data was provided

    def test_create_excel_with_empty_orcid_list(self):
        """Test Excel generation when no ORCID data is provided."""
        create_report([], "excel")
        output_file = os.path.join(self.result_dir, "orcid_report.xlsx")

        self.assertTrue(os.path.exists(output_file))

        df = pd.read_excel(output_file)
        self.assertEqual(len(df), 0)  # Should be empty

    def test_create_csv_with_no_publications(self):
        """Test CSV generation when ORCID has no publications."""
        create_report([self.orcid_res_2], "csv")
        output_file = os.path.join(self.result_dir, "orcid_report.csv")

        self.assertTrue(os.path.exists(output_file))

        with open(output_file, 'r', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            rows = list(reader)

        self.assertEqual(len(rows), 1)  # One row for Alice with no publications
        self.assertEqual(rows[0]['ORCID'], "0000-0002-9876-5432")
        self.assertEqual(rows[0]['Work Title'], '')

    def test_create_excel_with_no_publications(self):
        """Test Excel generation when ORCID has no publications."""
        create_report([self.orcid_res_2], "excel")
        output_file = os.path.join(self.result_dir, "orcid_report.xlsx")

        self.assertTrue(os.path.exists(output_file))

        df = pd.read_excel(output_file)
        self.assertEqual(len(df), 1)  # One row for Alice with no publications
        self.assertEqual(df.iloc[0]['ORCID'], "0000-0002-9876-5432")
        self.assertTrue(pd.isna(df.iloc[0]['Work Title']) or df.iloc[0]['Work Title'] == '')


if __name__ == "__main__":
    unittest.main()
