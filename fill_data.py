#!/usr/bin/env python3
"""
Script to populate an existing Google Docs template with data.

This script builds a full, styled “Company Performance Report” in a Google Doc.
It can:
- Use a service account JSON for authentication
- Work with an existing template document
- Optionally copy the template before populating it
- Optionally send a Slack notification with the generated document link
"""

import argparse
import sys
import os
from datetime import datetime
import pandas as pd

from docbuilder.docs_client import GoogleDocsClient
from docbuilder.styles import TextStyle, ParagraphStyle, Colors, FontFamilies, FontSizes, ParagraphStyleType
from docbuilder.elements import create_simple_table
from slack_alerts.alerting import send_notification

# Configuration defaults (can be overridden by env vars or CLI args)
DEFAULT_SERVICE_ACCOUNT_FILE = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "service.json")
DEFAULT_TEMPLATE_ID = os.getenv("DOC_TEMPLATE_ID", "1leLEu59zq0syC20PHZiFJ0cIGa8V-ciQxNXsqRvktA4")


def parse_args(argv=None):
    """Parse command-line arguments for the report generator."""
    parser = argparse.ArgumentParser(
        description="Generate a formatted performance report in Google Docs using a template."
    )
    parser.add_argument(
        "-s",
        "--service-account-file",
        default=DEFAULT_SERVICE_ACCOUNT_FILE,
        help="Path to the Google service account JSON file (default: %(default)s)",
    )
    parser.add_argument(
        "-t",
        "--template-id",
        default=DEFAULT_TEMPLATE_ID,
        help="Google Docs template ID to use (default: value from DOC_TEMPLATE_ID env var or built-in default).",
    )
    parser.add_argument(
        "-d",
        "--document-title",
        default=None,
        help="Optional title to set on the target document.",
    )
    parser.add_argument(
        "-c",
        "--copy-template",
        action="store_true",
        help="Copy the template before populating it (recommended for generating new reports).",
    )
    parser.add_argument(
        "--slack-channel",
        action="append",
        dest="slack_channels",
        help="Slack channel to notify with the generated document link. "
             "Can be provided multiple times for multiple channels.",
    )
    parser.add_argument(
        "--slack-endpoint",
        default=None,
        help="Optional override for the Slack endpoint defined in slack_alerts.alerting.",
    )

    return parser.parse_args(argv)

def create_dummy_data():
    """Create dummy data for tables and paragraphs."""
    
    # Employee data table
    employee_data = [
        ["Employee ID", "Name", "Department", "Position", "Salary", "Start Date"],
        ["EMP001", "John Smith", "Engineering", "Senior Developer", "$85,000", "2020-01-15"],
        ["EMP002", "Sarah Johnson", "Marketing", "Marketing Manager", "$72,000", "2019-06-01"],
        ["EMP003", "Michael Brown", "Sales", "Sales Representative", "$65,000", "2021-03-10"],
        ["EMP004", "Emily Davis", "HR", "HR Specialist", "$58,000", "2018-11-20"],
        ["EMP005", "David Wilson", "Finance", "Financial Analyst", "$70,000", "2020-08-05"],
        ["EMP006", "Lisa Anderson", "Engineering", "DevOps Engineer", "$82,000", "2021-01-12"],
        ["EMP007", "Robert Taylor", "Marketing", "Content Creator", "$55,000", "2022-04-18"],
        ["EMP008", "Jennifer White", "Sales", "Account Manager", "$68,000", "2019-12-03"]
    ]
    
    # Sales performance table
    sales_data = [
        ["Quarter", "Product Category", "Units Sold", "Revenue", "Growth %"],
        ["Q1 2024", "Software Licenses", "1,250", "$125,000", "+15%"],
        ["Q1 2024", "Consulting Services", "85", "$340,000", "+22%"],
        ["Q1 2024", "Training Programs", "45", "$67,500", "+8%"],
        ["Q2 2024", "Software Licenses", "1,420", "$142,000", "+13.6%"],
        ["Q2 2024", "Consulting Services", "92", "$368,000", "+8.2%"],
        ["Q2 2024", "Training Programs", "52", "$78,000", "+15.6%"],
        ["Q3 2024", "Software Licenses", "1,680", "$168,000", "+18.3%"],
        ["Q3 2024", "Consulting Services", "105", "$420,000", "+14.1%"],
        ["Q3 2024", "Training Programs", "58", "$87,000", "+11.5%"]
    ]
    
    # Financial summary table
    financial_data = [
        ["Metric", "Q1 2024", "Q2 2024", "Q3 2024", "YTD Total"],
        ["Total Revenue", "$532,500", "$588,000", "$675,000", "$1,795,500"],
        ["Operating Expenses", "$425,000", "$445,000", "$485,000", "$1,355,000"],
        ["Net Income", "$107,500", "$143,000", "$190,000", "$440,500"],
        ["Profit Margin", "20.2%", "24.3%", "28.1%", "24.5%"],
        ["Customer Acquisition", "125", "148", "172", "445"],
        ["Customer Retention", "92%", "94%", "96%", "94%"]
    ]
    
    employee_df = pd.DataFrame(employee_data[1:], columns=employee_data[0])
    sales_df = pd.DataFrame(sales_data[1:], columns=sales_data[0])
    financial_df = pd.DataFrame(financial_data[1:], columns=financial_data[0])
    
    return employee_df, sales_df, financial_df


def populate_document(
    service_account_file: str = DEFAULT_SERVICE_ACCOUNT_FILE,
    template_id: str = DEFAULT_TEMPLATE_ID,
    document_title: str | None = None,
    copy_template: bool = False,
    slack_channels: list[str] | None = None,
    slack_endpoint: str | None = None,
):
    """
    Populate a Google Docs document with report data.

    Returns:
        tuple[int, str | None]: (exit_code, final_document_id)
    """

    print("🚀 Starting document population...")

    try:
        # Initialize client
        print("🔐 Initializing Google Docs client...")
        client = GoogleDocsClient(service_account_file)

        # Decide which document to write to
        target_document_id = template_id

        if copy_template:
            print("📄 Copying template document before populating...")
            copy_title = document_title or f"Auto-generated Report {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            copied = client.copy_document(template_id, copy_title)
            target_document_id = copied.get("id") or copied.get("documentId", template_id)
            print(f"✅ Created copy with ID: {target_document_id}")

        # Optionally update the document title
        if document_title:
            try:
                client.set_document_title(target_document_id, document_title)
                print(f"📝 Set document title to: {document_title}")
            except Exception as e:
                print(f"⚠️ Could not set document title: {e}")

        print(f"📄 Working with document ID: {target_document_id}")

        # Get current document to find insertion point
        document = client.get_document(target_document_id)
        end_index = document['body']['content'][-1]['endIndex'] - 1
        print(f"📍 Document end index: {end_index}")
        
        # Create styles
        title_style = TextStyle(
            font_family=FontFamilies.ARIAL,
            font_size=FontSizes.XXLARGE,
            bold=True,
            foreground_color=Colors.BLUE
        )
        
        heading_style = TextStyle(
            font_family=FontFamilies.ARIAL,
            font_size=FontSizes.XLARGE,
            bold=True,
            foreground_color=Colors.DARK_GRAY
        )
        
        header_cell_style = TextStyle(
            bold=True,
            foreground_color=Colors.WHITE,
            background_color=Colors.BLUE,
            font_family=FontFamilies.ARIAL
        )
        
        data_cell_style = TextStyle(
            font_family=FontFamilies.ARIAL,
            font_size=FontSizes.NORMAL
        )
        
        intro_style = TextStyle(
            font_family=FontFamilies.ARIAL,
            font_size=FontSizes.MEDIUM,
            foreground_color=Colors.DARK_GRAY
        )
        
        # Add main title
        print("📝 Adding main title...")
        client.append_text(target_document_id, "\n\n")
        client.add_title(target_document_id, "Company Performance Report 2024", style=title_style)
        
        # Add introduction paragraph
        print("📝 Adding introduction...")
        intro_text = (
            "This comprehensive report provides an overview of our company's performance "
            "across multiple departments and key metrics for 2024. The data presented below "
            "includes employee information, sales performance, and financial summaries that "
            "demonstrate our continued growth and success."
        )
        client.append_text(target_document_id, f"\n{intro_text}\n\n", intro_style)
        
        # Get dummy data
        employee_data, sales_data, financial_data = create_dummy_data()
        
        # Add Employee Information Section
        print("👥 Adding employee information section...")
        client.add_heading(target_document_id, "Employee Directory", level=1, style=heading_style)
        
        employee_intro = (
            "Our team consists of talented professionals across various departments. "
            "Below is a summary of our current workforce:"
        )
        client.append_text(target_document_id, f"\n{employee_intro}\n\n", intro_style)
        
        # Insert employee table
        print("📊 Inserting employee table...")
        client.insert_table_from_dataframe(
            target_document_id,
            employee_data, 
            header_style=header_cell_style, 
            cell_style=data_cell_style
        )
        
        # Add some analysis text
        employee_analysis = (
            "\nKey insights from our employee data:\n"
            "• We have 8 employees across 4 departments\n"
            "• Average tenure is approximately 3.2 years\n"
            "• Engineering department has the highest average salary\n"
            "• Strong representation across technical and business functions\n"
        )
        client.append_text(target_document_id, employee_analysis, intro_style)
        
        # Add Sales Performance Section
        print("💰 Adding sales performance section...")
        client.add_heading(target_document_id, "\n\nSales Performance Analysis", level=1, style=heading_style)
        
        sales_intro = (
            "Our sales team has delivered exceptional results across all product categories. "
            "The quarterly breakdown shows consistent growth patterns:"
        )
        client.append_text(target_document_id, f"\n{sales_intro}\n\n", intro_style)
        
        # Insert sales table
        print("📊 Inserting sales performance table...")
        client.insert_table_from_dataframe(
            target_document_id,
            sales_data, 
            header_style=header_cell_style, 
            cell_style=data_cell_style
        )
        
        # Add sales analysis
        sales_analysis = (
            "\nSales Performance Highlights:\n"
            "• Software Licenses show strongest unit growth (+18.3% in Q3)\n"
            "• Consulting Services maintain high revenue per unit\n"
            "• Training Programs demonstrate steady improvement\n"
            "• Overall revenue growth trend is positive across all quarters\n"
        )
        client.append_text(target_document_id, sales_analysis, intro_style)
        
        # Add Financial Summary Section
        print("💼 Adding financial summary section...")
        client.add_heading(target_document_id, "\n\nFinancial Summary", level=1, style=heading_style)
        
        financial_intro = (
            "Our financial performance demonstrates strong profitability and sustainable growth. "
            "Key financial metrics are summarized below:"
        )
        client.append_text(target_document_id, f"\n{financial_intro}\n\n", intro_style)
        
        # Insert financial table
        print("📊 Inserting financial summary table...")
        client.insert_table_from_dataframe(
            target_document_id,
            financial_data, 
            header_style=header_cell_style, 
            cell_style=data_cell_style
        )
        
        # Add financial analysis
        financial_analysis = (
            "\nFinancial Performance Highlights:\n"
            "• YTD revenue of $1.79M exceeds projections by 12%\n"
            "• Profit margin improved from 20.2% to 28.1% over three quarters\n"
            "• Customer acquisition rate increased by 37.6%\n"
            "• Customer retention improved to 96% in Q3\n"
            "• Operating efficiency gains evident in expense management\n"
        )
        client.append_text(target_document_id, financial_analysis, intro_style)
        
        # Add conclusion section
        print("📋 Adding conclusion...")
        client.add_heading(target_document_id, "\n\nConclusion & Next Steps", level=1, style=heading_style)
        
        conclusion_text = (
            "Based on the comprehensive analysis above, our company is positioned for continued "
            "success in 2024 and beyond. Key strategic initiatives for the upcoming quarter include:\n\n"
            "1. Expanding our engineering team to support increased software license demand\n"
            "2. Enhancing our consulting services portfolio with new specialized offerings\n"
            "3. Developing advanced training programs to meet market demand\n"
            "4. Implementing customer success programs to maintain high retention rates\n"
            "5. Optimizing operational processes to maintain profit margin growth\n\n"
            "This data-driven approach ensures we remain competitive while delivering value "
            "to our customers and stakeholders."
        )
        client.append_text(target_document_id, f"\n{conclusion_text}\n", intro_style)
        
        # Add page break for clean separation
        print("📄 Adding page break...")
        client.insert_page_break(target_document_id)
        
        # Add appendix with additional data
        print("📎 Adding appendix...")
        client.add_heading(target_document_id, "Appendix: Additional Information", level=1, style=heading_style)
        
        appendix_text = (
            "Document generated automatically using Google Docs API Client Library.\n"
            f"Generated on: {__import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            "Data source: Internal company systems\n"
            "Report frequency: Quarterly\n"
            "Next update: End of Q4 2024\n"
        )
        client.append_text(target_document_id, f"\n{appendix_text}", intro_style)
        
        print("\n" + "="*60)
        print("✅ Document population completed successfully!")
        doc_url = f"https://docs.google.com/document/d/{target_document_id}"
        print(f"📄 Document URL: {doc_url}")
        print("\n📊 Data inserted:")
        print("   • Employee directory (8 employees)")
        print("   • Sales performance data (9 quarters)")
        print("   • Financial summary (6 key metrics)")
        print("   • Analysis and insights")
        print("   • Conclusion and next steps")
        print("   • Appendix information")
        print("\n🎨 Formatting applied:")
        print("   • Title and heading styles")
        print("   • Styled tables with headers")
        print("   • Consistent text formatting")
        print("   • Color-coded elements")
        print("   • Professional layout")
        # Optional Slack notification
        if slack_channels:
            print("\n📣 Sending Slack notifications...")
            slack_kwargs = {}
            if slack_endpoint:
                slack_kwargs["slack_endpoint"] = slack_endpoint
            try:
                statuses = send_notification(slack_channels, doc_url, **slack_kwargs)
                print(f"Slack responses: {statuses}")
            except Exception as e:
                print(f"⚠️ Failed to send Slack notifications: {e}")

    except Exception as e:
        print(f"❌ Error: {e}")
        print("\nTroubleshooting tips:")
        print("• Ensure the service account has edit access to the document")
        print("• Verify the document ID is correct")
        print("• Check internet connectivity")
        print("• Make sure Google Docs and Drive APIs are enabled")
        return 1, None

    return 0, target_document_id


if __name__ == "__main__":
    args = parse_args()
    exit_code, _doc_id = populate_document(
        service_account_file=args.service_account_file,
        template_id=args.template_id,
        document_title=args.document_title,
        copy_template=args.copy_template,
        slack_channels=args.slack_channels,
        slack_endpoint=args.slack_endpoint,
    )
    sys.exit(exit_code)
