# Add this function at the very bottom of obe/mapper.py

def main_cli():
    """CLI Entry point for obe-mapper command."""
    import argparse

    parser = argparse.ArgumentParser(
        description="CO-WK / CO-PO mapping for NBA Outcome-Based Education"
    )
    sub = parser.add_subparsers(dest="command")

    p1 = sub.add_parser("wk", help="Generate CO + WK mapping from a syllabus PDF")
    p1.add_argument("pdf", help="Path to syllabus PDF")
    p1.add_argument("--model", required=True, help="Gemini model name")
    p1.add_argument("--api-key", required=True, help="Gemini API key")
    p1.add_argument("-o", "--output", default=DEFAULT_INPUT_EXCEL)

    p2 = sub.add_parser("po", help="Generate CO-PO mapping from Excel")
    p2.add_argument("excel", nargs="?", default=DEFAULT_INPUT_EXCEL)
    p2.add_argument("--model", required=True, help="Gemini model name")
    p2.add_argument("--api-key", required=True, help="Gemini API key")
    p2.add_argument("-o", "--output", default=DEFAULT_OUTPUT_EXCEL)
    p2.add_argument("--sheet", default=None)

    args = parser.parse_args()

    if args.command == "wk":
        generate_co_wk_excel(
            pdf_path=args.pdf,
            model_name=args.model,
            api_key=args.api_key,
            output_excel_path=args.output,
        )
    elif args.command == "po":
        generate_co_po_mapping(
            model_name=args.model,
            api_key=args.api_key,
            input_excel=args.excel,
            output_excel=args.output,
            sheet_name=args.sheet,
        )
    else:
        parser.print_help()
