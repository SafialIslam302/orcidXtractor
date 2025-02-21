import argparse
import os
import sys

from orcid_extractor import get_orcid_data
from report_generator import create_txt, create_pdf, create_json, create_report

# from .orcid_extractor import get_orcid_data
# from .report_generator import create_txt, create_pdf, create_json, create_report


def parse_arguments() -> argparse.Namespace:
    """
    Parses command-line arguments.

    Returns:
        argparse.Namespace: An object containing the parsed arguments.
    """
    parser = argparse.ArgumentParser(
        description="Extract ORCID information and generate reports.\n"
                    "You can either provide an input file using --inputfile OR "
                    "pass ORCID IDs (--orcid_ids) directly as arguments (e.g., python main.py 0000-0001-...).",
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument('--inputfile', help="Path to the input file containing ORCID IDs.")
    parser.add_argument('--orcid_ids', nargs='*', help="List of ORCID IDs (e.g., 0000-0001-...).")
    parser.add_argument('--output-format', nargs='+', choices=['txt', 'pdf', 'json'],
                        help="Specify one or more output formats (txt, pdf, json).\nExample: --output-format txt pdf json")
    parser.add_argument('--report', choices=['csv', 'excel'],
                        help="Specify if you want to generate a CSV or Excel report.\nExample: --report csv")
    return parser.parse_args()


def main() -> None:
    """
    Main function to orchestrate ORCID data extraction and report generation.
    """
    args = parse_arguments()

    if args.inputfile:
        with open(args.inputfile) as f:
            orcid_ids = [line.rstrip() for line in f]
    elif args.orcid_ids:
        orcid_ids = args.orcid_ids
    else:
        print("Error: Either --inputfile or a list of ORCID IDs (--orcid_ids) must be provided.")
        print("\nFor more information, please use: orcidxtract --help")
        sys.exit(1)

    # Get the directory of the script
    current_dir = os.getcwd()
    result_dir = os.path.join(current_dir, "Result")

    if not os.path.exists(result_dir):
        os.makedirs(result_dir)

    if not args.output_format and not args.report:
        print("--output-format or --report is missing. For help, run: orcidxtract --help")
        sys.exit(1)

    orcid_data = []
    for orcid_id in orcid_ids:
        orcid_res = get_orcid_data(orcid_id)
        if orcid_res is not None:
            orcid_data.append(orcid_res)

            if args.output_format:
                print(f"Start processing ORCID ID: {orcid_res.orcid}")

            if 'txt' in args.output_format:
                output_file_name = os.path.join(result_dir, f"{orcid_id}.txt")
                create_txt(output_file_name, orcid_res)

            if 'pdf' in args.output_format:
                output_file_name = os.path.join(result_dir, f"{orcid_id}.pdf")
                create_pdf(output_file_name, orcid_res)

            if 'json' in args.output_format:
                output_file_name = os.path.join(result_dir, f"{orcid_id}.json")
                create_json(output_file_name, orcid_res)

            print(f"End processing ORCID ID: {orcid_res.orcid}")

    if args.report:
        print("Start creating report.")
        create_report(orcid_data, args.report)
        print("End creating report.")


if __name__ == "__main__":
    main()
