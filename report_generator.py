import pandas as pd
import json
from fpdf import FPDF
from datetime import datetime
from collections import defaultdict
import os

class ReportGenerator:
    """
    Comprehensive report generator for art instruction processing
    Generates reports in Excel, PDF, Text, and JSON formats
    """
    
    def __init__(self):
        self.timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    def generate_all_reports(self, report_data, output_folder, timestamp, sales_order_filter=None):
        """
        Generate all report formats (Excel, PDF, Text, JSON)
        
        Args:
            report_data (list): List of dictionaries containing processing results
            output_folder (str): Path to output folder
            timestamp (str): Timestamp for file naming
            sales_order_filter (str): Sales order filter used (if any)
        """
        print(f"Generating comprehensive reports with {len(report_data)} records...")
        
        # Generate each report format
        self.generate_excel_report(report_data, output_folder, timestamp, sales_order_filter)
        self.generate_pdf_report(report_data, output_folder, timestamp, sales_order_filter)
        self.generate_text_report(report_data, output_folder, timestamp, sales_order_filter)
        self.generate_json_report(report_data, output_folder, timestamp, sales_order_filter)
        
        print("All reports generated successfully!")
    
    def generate_excel_report(self, report_data, output_folder, timestamp, sales_order_filter=None):
        """
        Generate Excel report with original data plus execution status columns
        """
        if not report_data:
            print("No data to generate Excel report")
            return
            
        try:
            # Convert report data to DataFrame
            df = pd.DataFrame(report_data)
            
            # Ensure all required columns exist
            required_columns = [
                'Document Number', 'LOGO', 'VENDOR STYLE', 'COLOR', 'SUBCATEGORY', 
                'Quantity', 'Customer/Vendor Name', 'Due Date', 'DueDateStatus',
                'OPERATIONAL CODE', 'List of Operation Codes', 'LOGO POSITION',
                'STITCH COUNT', 'NOTES', 'FILE NAME', 'Execution Status', 'Error Message'
            ]
            
            for col in required_columns:
                if col not in df.columns:
                    df[col] = ''
            
            # Reorder columns to match input format with execution columns at the end
            input_columns = [col for col in required_columns if col not in ['Execution Status', 'Error Message']]
            final_columns = input_columns + ['Execution Status', 'Error Message']
            df = df[final_columns]
            
            # Generate filename
            filter_suffix = f"_SO_{sales_order_filter}" if sales_order_filter else ""
            filename = f"Art_Instructions_Report_{timestamp}{filter_suffix}.xlsx"
            filepath = os.path.join(output_folder, filename)
            
            # Save to Excel with formatting
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Execution Report', index=False)
                
                # Get the workbook and worksheet
                workbook = writer.book
                worksheet = writer.sheets['Execution Report']
                
                # Apply formatting
                from openpyxl.styles import PatternFill, Font
                
                # Header formatting
                header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                header_font = Font(color="FFFFFF", bold=True)
                
                for cell in worksheet[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                
                # Status column formatting
                success_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                failed_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                
                execution_status_col = None
                for idx, cell in enumerate(worksheet[1]):
                    if cell.value == 'Execution Status':
                        execution_status_col = idx + 1
                        break
                
                if execution_status_col:
                    for row in range(2, worksheet.max_row + 1):
                        status_cell = worksheet.cell(row=row, column=execution_status_col)
                        if status_cell.value == 'SUCCESS':
                            status_cell.fill = success_fill
                        elif status_cell.value == 'FAILED':
                            status_cell.fill = failed_fill
                
                # Auto-adjust column widths
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)  # Max width of 50
                    worksheet.column_dimensions[column_letter].width = adjusted_width
            
            print(f"Excel report generated: {filename}")
            
        except Exception as e:
            print(f"Error generating Excel report: {e}")
    
    def generate_pdf_report(self, report_data, output_folder, timestamp, sales_order_filter=None):
        """
        Generate PDF report organized by sales order with item-level details
        """
        if not report_data:
            print("No data to generate PDF report")
            return
            
        try:
            # Group data by sales order
            sales_orders = defaultdict(list)
            for record in report_data:
                so_number = record.get('Document Number', 'Unknown')
                sales_orders[so_number].append(record)
            
            # Generate filename
            filter_suffix = f"_SO_{sales_order_filter}" if sales_order_filter else ""
            filename = f"Art_Instructions_Report_{timestamp}{filter_suffix}.pdf"
            filepath = os.path.join(output_folder, filename)
            
            # Create PDF
            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            
            # Title page
            pdf.add_page()
            pdf.set_font('Arial', 'B', 16)
            pdf.cell(0, 10, 'Art Instructions Processing Report', ln=True, align='C')
            pdf.ln(5)
            
            pdf.set_font('Arial', '', 12)
            pdf.cell(0, 8, f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}', ln=True, align='C')
            
            if sales_order_filter:
                pdf.cell(0, 8, f'Filtered by Sales Order: {sales_order_filter}', ln=True, align='C')
            
            pdf.ln(10)
            
            # Summary statistics
            total_records = len(report_data)
            success_count = sum(1 for record in report_data if record.get('Execution Status') == 'SUCCESS')
            failed_count = total_records - success_count
            success_rate = (success_count / total_records * 100) if total_records > 0 else 0
            
            pdf.set_font('Arial', 'B', 14)
            pdf.cell(0, 8, 'Summary Statistics', ln=True)
            pdf.ln(2)
            
            pdf.set_font('Arial', '', 10)
            pdf.cell(0, 6, f'Total Records Processed: {total_records}', ln=True)
            pdf.cell(0, 6, f'Successful: {success_count}', ln=True)
            pdf.cell(0, 6, f'Failed: {failed_count}', ln=True)
            pdf.cell(0, 6, f'Success Rate: {success_rate:.1f}%', ln=True)
            pdf.cell(0, 6, f'Total Sales Orders: {len(sales_orders)}', ln=True)
            
            pdf.ln(10)
            
            # Detailed report by sales order
            pdf.set_font('Arial', 'B', 14)
            pdf.cell(0, 8, 'Detailed Report by Sales Order', ln=True)
            pdf.ln(5)
            
            for so_number, items in sales_orders.items():
                # Check if we need a new page
                if pdf.get_y() > 250:
                    pdf.add_page()
                
                # Sales Order header
                pdf.set_font('Arial', 'B', 12)
                pdf.cell(0, 8, f'Sales Order: {so_number}', ln=True)
                pdf.ln(2)
                
                # SO summary
                so_total = len(items)
                so_success = sum(1 for item in items if item.get('Execution Status') == 'SUCCESS')
                so_failed = so_total - so_success
                
                pdf.set_font('Arial', '', 10)
                pdf.cell(0, 5, f'Items: {so_total} | Successful: {so_success} | Failed: {so_failed}', ln=True)
                pdf.ln(3)
                
                # Customer info (from first item)
                if items:
                    customer_name = items[0].get('Customer/Vendor Name', 'N/A')
                    due_date = items[0].get('Due Date', 'N/A')
                    pdf.cell(0, 5, f'Customer: {customer_name}', ln=True)
                    pdf.cell(0, 5, f'Due Date: {due_date}', ln=True)
                    pdf.ln(3)
                
                # Items table header
                pdf.set_font('Arial', 'B', 8)
                col_widths = [15, 20, 25, 20, 20, 15, 30, 45]  # Adjust as needed
                headers = ['Logo', 'Style', 'Color', 'Description', 'Qty', 'Op Code', 'Status', 'Error Message']
                
                for i, header in enumerate(headers):
                    pdf.cell(col_widths[i], 6, header, 1, 0, 'C')
                pdf.ln()
                
                # Items data
                pdf.set_font('Arial', '', 7)
                for item in items:
                    # Check if we need a new page
                    if pdf.get_y() > 270:
                        pdf.add_page()
                        # Repeat header on new page
                        pdf.set_font('Arial', 'B', 8)
                        for i, header in enumerate(headers):
                            pdf.cell(col_widths[i], 6, header, 1, 0, 'C')
                        pdf.ln()
                        pdf.set_font('Arial', '', 7)
                    
                    values = [
                        str(item.get('LOGO', ''))[:12],  # Truncate long values
                        str(item.get('VENDOR STYLE', ''))[:18],
                        str(item.get('COLOR', ''))[:22],
                        str(item.get('SUBCATEGORY', ''))[:18],
                        str(item.get('Quantity', ''))[:12],
                        str(item.get('OPERATIONAL CODE', ''))[:12],
                        str(item.get('Execution Status', ''))[:8],
                        str(item.get('Error Message', ''))[:40]
                    ]
                    
                    for i, value in enumerate(values):
                        pdf.cell(col_widths[i], 5, value, 1, 0, 'L')
                    pdf.ln()
                
                pdf.ln(5)  # Space between sales orders
            
            pdf.output(filepath)
            print(f"PDF report generated: {filename}")
            
        except Exception as e:
            print(f"Error generating PDF report: {e}")
    
    def generate_text_report(self, report_data, output_folder, timestamp, sales_order_filter=None):
        """
        Generate text report with detailed breakdown by sales order and items
        """
        if not report_data:
            print("No data to generate text report")
            return
            
        try:
            # Generate filename
            filter_suffix = f"_SO_{sales_order_filter}" if sales_order_filter else ""
            filename = f"Art_Instructions_Report_{timestamp}{filter_suffix}.txt"
            filepath = os.path.join(output_folder, filename)
            
            # Group data by sales order
            sales_orders = defaultdict(list)
            for record in report_data:
                so_number = record.get('Document Number', 'Unknown')
                sales_orders[so_number].append(record)
            
            with open(filepath, 'w', encoding='utf-8') as f:
                # Header
                f.write("="*80 + "\n")
                f.write("ART INSTRUCTIONS PROCESSING REPORT\n")
                f.write("="*80 + "\n")
                f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                
                if sales_order_filter:
                    f.write(f"Filtered by Sales Order: {sales_order_filter}\n")
                
                f.write("\n")
                
                # Summary statistics
                total_records = len(report_data)
                success_count = sum(1 for record in report_data if record.get('Execution Status') == 'SUCCESS')
                failed_count = total_records - success_count
                success_rate = (success_count / total_records * 100) if total_records > 0 else 0
                
                f.write("SUMMARY STATISTICS\n")
                f.write("-" * 30 + "\n")
                f.write(f"Total Records Processed: {total_records}\n")
                f.write(f"Successful: {success_count}\n")
                f.write(f"Failed: {failed_count}\n")
                f.write(f"Success Rate: {success_rate:.1f}%\n")
                f.write(f"Total Sales Orders: {len(sales_orders)}\n")
                f.write("\n")
                
                # Error summary (if any failures)
                if failed_count > 0:
                    f.write("ERROR SUMMARY\n")
                    f.write("-" * 30 + "\n")
                    
                    error_counts = defaultdict(int)
                    for record in report_data:
                        if record.get('Execution Status') == 'FAILED':
                            error_msg = record.get('Error Message', 'Unknown error')
                            error_counts[error_msg] += 1
                    
                    for error, count in sorted(error_counts.items(), key=lambda x: x[1], reverse=True):
                        f.write(f"  {error}: {count} occurrences\n")
                    f.write("\n")
                
                # Detailed report by sales order
                f.write("DETAILED REPORT BY SALES ORDER\n")
                f.write("="*50 + "\n\n")
                
                for so_number, items in sorted(sales_orders.items()):
                    f.write(f"SALES ORDER: {so_number}\n")
                    f.write("-" * 40 + "\n")
                    
                    # SO summary
                    so_total = len(items)
                    so_success = sum(1 for item in items if item.get('Execution Status') == 'SUCCESS')
                    so_failed = so_total - so_success
                    
                    f.write(f"Items: {so_total} | Successful: {so_success} | Failed: {so_failed}\n")
                    
                    # Customer info (from first item)
                    if items:
                        customer_name = items[0].get('Customer/Vendor Name', 'N/A')
                        due_date = items[0].get('Due Date', 'N/A')
                        f.write(f"Customer: {customer_name}\n")
                        f.write(f"Due Date: {due_date}\n")
                    
                    f.write("\n")
                    
                    # Items details
                    f.write("ITEMS:\n")
                    for i, item in enumerate(items, 1):
                        status = item.get('Execution Status', 'UNKNOWN')
                        status_symbol = "✓" if status == 'SUCCESS' else "✗"
                        
                        f.write(f"  {i}. {status_symbol} Logo: {item.get('LOGO', 'N/A')} | ")
                        f.write(f"Style: {item.get('VENDOR STYLE', 'N/A')} | ")
                        f.write(f"Color: {item.get('COLOR', 'N/A')} | ")
                        f.write(f"Qty: {item.get('Quantity', 'N/A')}\n")
                        f.write(f"     Status: {status}")
                        
                        if status == 'FAILED':
                            error_msg = item.get('Error Message', 'No error message')
                            f.write(f" | Error: {error_msg}")
                        
                        f.write("\n")
                    
                    f.write("\n" + "="*50 + "\n\n")
            
            print(f"Text report generated: {filename}")
            
        except Exception as e:
            print(f"Error generating text report: {e}")
    
    def generate_json_report(self, report_data, output_folder, timestamp, sales_order_filter=None):
        """
        Generate JSON report with structured data for programmatic access
        """
        if not report_data:
            print("No data to generate JSON report")
            return
            
        try:
            # Generate filename
            filter_suffix = f"_SO_{sales_order_filter}" if sales_order_filter else ""
            filename = f"Art_Instructions_Report_{timestamp}{filter_suffix}.json"
            filepath = os.path.join(output_folder, filename)
            
            # Prepare structured data
            report_structure = {
                "metadata": {
                    "generated_at": datetime.now().isoformat(),
                    "total_records": len(report_data),
                    "success_count": sum(1 for record in report_data if record.get('Execution Status') == 'SUCCESS'),
                    "failed_count": sum(1 for record in report_data if record.get('Execution Status') == 'FAILED'),
                    "sales_order_filter": sales_order_filter,
                    "generator_version": "1.0"
                },
                "summary": {
                    "success_rate": 0,
                    "total_sales_orders": 0,
                    "error_breakdown": {}
                },
                "sales_orders": {},
                "raw_data": report_data
            }
            
            # Calculate summary metrics
            total_records = len(report_data)
            if total_records > 0:
                success_count = report_structure["metadata"]["success_count"]
                report_structure["summary"]["success_rate"] = round((success_count / total_records) * 100, 2)
            
            # Group by sales orders for structured view
            sales_orders = defaultdict(list)
            for record in report_data:
                so_number = record.get('Document Number', 'Unknown')
                sales_orders[so_number].append(record)
            
            report_structure["summary"]["total_sales_orders"] = len(sales_orders)
            
            # Error breakdown
            error_counts = defaultdict(int)
            for record in report_data:
                if record.get('Execution Status') == 'FAILED':
                    error_msg = record.get('Error Message', 'Unknown error')
                    error_counts[error_msg] += 1
            
            report_structure["summary"]["error_breakdown"] = dict(error_counts)
            
            # Structured sales orders data
            for so_number, items in sales_orders.items():
                so_success = sum(1 for item in items if item.get('Execution Status') == 'SUCCESS')
                so_failed = len(items) - so_success
                
                report_structure["sales_orders"][so_number] = {
                    "total_items": len(items),
                    "successful_items": so_success,
                    "failed_items": so_failed,
                    "success_rate": round((so_success / len(items)) * 100, 2) if items else 0,
                    "customer_name": items[0].get('Customer/Vendor Name', 'N/A') if items else 'N/A',
                    "due_date": items[0].get('Due Date', 'N/A') if items else 'N/A',
                    "items": items
                }
            
            # Save JSON file
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(report_structure, f, indent=2, ensure_ascii=False, default=str)
            
            print(f"JSON report generated: {filename}")
            
        except Exception as e:
            print(f"Error generating JSON report: {e}")
    
    def get_error_statistics(self, report_data):
        """
        Get detailed error statistics for debugging purposes
        """
        error_stats = {
            "total_errors": 0,
            "error_types": defaultdict(int),
            "errors_by_sales_order": defaultdict(list),
            "most_common_errors": []
        }
        
        for record in report_data:
            if record.get('Execution Status') == 'FAILED':
                error_stats["total_errors"] += 1
                error_msg = record.get('Error Message', 'Unknown error')
                error_stats["error_types"][error_msg] += 1
                
                so_number = record.get('Document Number', 'Unknown')
                error_stats["errors_by_sales_order"][so_number].append({
                    "logo": record.get('LOGO', 'N/A'),
                    "error": error_msg,
                    "style": record.get('VENDOR STYLE', 'N/A')
                })
        
        # Sort errors by frequency
        error_stats["most_common_errors"] = sorted(
            error_stats["error_types"].items(), 
            key=lambda x: x[1], 
            reverse=True
        )
        
        return dict(error_stats)