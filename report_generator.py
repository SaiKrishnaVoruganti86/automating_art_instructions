import pandas as pd
from fpdf import FPDF
from datetime import datetime
from collections import defaultdict
import os

class ReportGenerator:
    """
    Comprehensive report generator for art instruction processing
    Generates reports in Excel, PDF
    """
    
    def __init__(self):
        self.timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    def preprocess_report_data(self, report_data):
        """
        Preprocess report data to handle special cases:
        - Convert execution status to N/A for Invalid Logo SKU errors
        - Convert execution status to Not Approved for "Status: Not Approved" errors
        """
        processed_data = []
        for record in report_data:
            processed_record = record.copy()
            error_msg = record.get('Error Message', '')
            
            # Check if error message contains "Invalid Logo SKU:"
            if 'Invalid Logo SKU:' in error_msg and error_msg.strip().endswith('""'):
                processed_record['Execution Status'] = 'N/A'
            # Check if error message is "Status: Not Approved"
            elif error_msg.strip() == "Status: Not Approved":
                processed_record['Execution Status'] = 'Not Approved'
            
            processed_data.append(processed_record)
        
        return processed_data
    
    def create_overview_data(self, report_data):
        """
        Create overview data grouped by document number with completion status
        """
        # Group data by document number
        sales_orders = defaultdict(list)
        for record in report_data:
            so_number = record.get('Document Number', 'Unknown')
            sales_orders[so_number].append(record)
        
        overview_data = []
        for so_number, items in sales_orders.items():
            # Calculate counts for this sales order
            so_total = len(items)
            so_success = sum(1 for item in items if item.get('Execution Status') == 'SUCCESS')
            so_failed = sum(1 for item in items if item.get('Execution Status') == 'FAILED')
            so_na = sum(1 for item in items if item.get('Execution Status') == 'N/A')
            so_not_approved = sum(1 for item in items if item.get('Execution Status') == 'Not Approved')
            
            # Calculate success rate (include N/A as success)
            so_success_rate = ((so_success + so_na) / so_total * 100) if so_total > 0 else 0
            
            # Determine completion status based on success rate
            # FULLY SUCCESS: All items are either SUCCESS or N/A (100% success rate)
            # TOTAL FAILED: No items are SUCCESS or N/A (0% success rate) 
            # PARTIAL SUCCESS: Mix of success/failure (1-99% success rate)
            if so_success_rate == 100:
                completion_status = "FULLY SUCCESS"
            elif so_success_rate == 0:
                completion_status = "TOTAL FAILED"
            else:
                completion_status = "PARTIAL SUCCESS"
            
            # Get customer name and due date from first item
            customer_name = items[0].get('Customer/Vendor Name', 'N/A') if items else 'N/A'
            due_date = items[0].get('Due Date', 'N/A') if items else 'N/A'
            
            overview_data.append({
                'Document Number': so_number,
                'Customer/Vendor Name': customer_name,
                'Due Date': due_date,
                'Total Items': so_total,
                'Success': so_success,
                'Failed': so_failed,
                'N/A': so_na,
                'Not Approved': so_not_approved,
                'Success Rate (%)': round(so_success_rate, 1),
                'Completion Status': completion_status
            })
        
        # Sort by document number
        overview_data.sort(key=lambda x: x['Document Number'])
        
        return overview_data
    
    def generate_all_reports(self, report_data, output_folder, timestamp, sales_order_filter=None):
        """
        Generate all report formats (Excel, PDF)
        
        Args:
            report_data (list): List of dictionaries containing processing results
            output_folder (str): Path to output folder
            timestamp (str): Timestamp for file naming
            sales_order_filter (str): Sales order filter used (if any)
        """
        print(f"Generating comprehensive reports with {len(report_data)} records...")
        
        # Preprocess data to handle Invalid Logo SKU cases
        processed_data = self.preprocess_report_data(report_data)
        
        # Generate each report format
        self.generate_detailed_excel_report(processed_data, output_folder, timestamp, sales_order_filter)
        self.generate_overview_excel_report(processed_data, output_folder, timestamp, sales_order_filter)
        self.generate_pdf_report(processed_data, output_folder, timestamp, sales_order_filter)
        
        print("All reports generated successfully!")
    
    def generate_detailed_excel_report(self, report_data, output_folder, timestamp, sales_order_filter=None):
        """
        Generate detailed Excel report with all data including execution status columns
        """
        if not report_data:
            print("No data to generate detailed Excel report")
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
            filename = f"Art_Instructions_Detailed_Report_{timestamp}{filter_suffix}.xlsx"
            filepath = os.path.join(output_folder, filename)
            
            # Save to Excel with formatting
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Full Detailed Report', index=False)
                
                # Get the workbook and worksheet
                workbook = writer.book
                worksheet = writer.sheets['Full Detailed Report']
                
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
                na_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")  # Grey for N/A
                not_approved_fill = PatternFill(start_color="FFE4B5", end_color="FFE4B5", fill_type="solid")  # Light orange for Not Approved
                
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
                        elif status_cell.value == 'N/A':
                            status_cell.fill = na_fill
                        elif status_cell.value == 'Not Approved':
                            status_cell.fill = not_approved_fill
                
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
            
            print(f"Detailed Excel report generated: {filename}")
            
        except Exception as e:
            print(f"Error generating detailed Excel report: {e}")
    
    def generate_overview_excel_report(self, report_data, output_folder, timestamp, sales_order_filter=None):
        """
        Generate overview Excel report with only Document Number and Completion Status
        """
        if not report_data:
            print("No data to generate overview Excel report")
            return
            
        try:
            # Create overview data grouped by document number
            overview_data = self.create_simple_overview_data(report_data)
            overview_df = pd.DataFrame(overview_data)
            
            # Generate filename
            filter_suffix = f"_SO_{sales_order_filter}" if sales_order_filter else ""
            filename = f"Art_Instructions_Overview_Report_{timestamp}{filter_suffix}.xlsx"
            filepath = os.path.join(output_folder, filename)
            
            # Save to Excel with formatting
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                overview_df.to_excel(writer, sheet_name='Overview Report', index=False)
                
                # Get the workbook and worksheet
                workbook = writer.book
                worksheet = writer.sheets['Overview Report']
                
                # Apply formatting
                from openpyxl.styles import PatternFill, Font
                
                # Header formatting
                header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                header_font = Font(color="FFFFFF", bold=True)
                
                for cell in worksheet[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                
                # Status column formatting
                fully_success_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # Green
                partial_success_fill = PatternFill(start_color="000080", end_color="000080", fill_type="solid")  # Dark Blue
                partial_success_font = Font(color="FFFFFF")  # White text for dark blue background
                total_failed_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # Red
                
                # Find completion status column
                completion_status_col = None
                for idx, cell in enumerate(worksheet[1]):
                    if cell.value == 'Completion Status':
                        completion_status_col = idx + 1
                        break
                
                if completion_status_col:
                    for row in range(2, worksheet.max_row + 1):
                        status_cell = worksheet.cell(row=row, column=completion_status_col)
                        if status_cell.value == 'FULLY SUCCESS':
                            status_cell.fill = fully_success_fill
                        elif status_cell.value == 'PARTIAL SUCCESS':
                            status_cell.fill = partial_success_fill
                            status_cell.font = partial_success_font
                        elif status_cell.value == 'TOTAL FAILED':
                            status_cell.fill = total_failed_fill
                
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
                    adjusted_width = min(max_length + 2, 30)  # Max width of 30 for overview
                    worksheet.column_dimensions[column_letter].width = adjusted_width
            
            print(f"Overview Excel report generated: {filename}")
            
        except Exception as e:
            print(f"Error generating overview Excel report: {e}")
    
    def create_simple_overview_data(self, report_data):
        """
        Create simple overview data with only Document Number and Completion Status
        """
        # Group data by document number
        sales_orders = defaultdict(list)
        for record in report_data:
            so_number = record.get('Document Number', 'Unknown')
            sales_orders[so_number].append(record)
        
        overview_data = []
        for so_number, items in sales_orders.items():
            # Calculate counts for this sales order
            so_total = len(items)
            so_success = sum(1 for item in items if item.get('Execution Status') == 'SUCCESS')
            so_na = sum(1 for item in items if item.get('Execution Status') == 'N/A')
            
            # Calculate success rate (include N/A as success)
            so_success_rate = ((so_success + so_na) / so_total * 100) if so_total > 0 else 0
            
            # Determine completion status based on success rate
            # FULLY SUCCESS: All items are either SUCCESS or N/A (100% success rate)
            # TOTAL FAILED: No items are SUCCESS or N/A (0% success rate) 
            # PARTIAL SUCCESS: Mix of success/failure (1-99% success rate)
            if so_success_rate == 100:
                completion_status = "FULLY SUCCESS"
            elif so_success_rate == 0:
                completion_status = "TOTAL FAILED"
            else:
                completion_status = "PARTIAL SUCCESS"
            
            overview_data.append({
                'Document Number': so_number,
                'Completion Status': completion_status
            })
        
        # Sort by document number
        overview_data.sort(key=lambda x: x['Document Number'])
        
        return overview_data
    
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
            
            # Summary statistics (updated to include N/A and Not Approved counts)
            total_records = len(report_data)
            success_count = sum(1 for record in report_data if record.get('Execution Status') == 'SUCCESS')
            failed_count = sum(1 for record in report_data if record.get('Execution Status') == 'FAILED')
            na_count = sum(1 for record in report_data if record.get('Execution Status') == 'N/A')
            not_approved_count = sum(1 for record in report_data if record.get('Execution Status') == 'Not Approved')
            # Include N/A as success for success rate calculation
            success_rate = ((success_count + na_count) / total_records * 100) if total_records > 0 else 0
            
            pdf.set_font('Arial', 'B', 14)
            pdf.cell(0, 8, 'Summary Statistics', ln=True)
            pdf.ln(2)
            
            pdf.set_font('Arial', '', 10)
            pdf.cell(0, 6, f'Total Records Processed: {total_records}', ln=True)
            pdf.cell(0, 6, f'Successful: {success_count}', ln=True)
            pdf.cell(0, 6, f'Failed: {failed_count}', ln=True)
            pdf.cell(0, 6, f'N/A (Invalid Logo SKU): {na_count}', ln=True)
            pdf.cell(0, 6, f'Not Approved: {not_approved_count}', ln=True)
            pdf.cell(0, 6, f'Success Rate: {success_rate:.1f}% (includes N/A as success)', ln=True)
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
                
                # SO summary (updated to include N/A, Not Approved counts and success rate)
                so_total = len(items)
                so_success = sum(1 for item in items if item.get('Execution Status') == 'SUCCESS')
                so_failed = sum(1 for item in items if item.get('Execution Status') == 'FAILED')
                so_na = sum(1 for item in items if item.get('Execution Status') == 'N/A')
                so_not_approved = sum(1 for item in items if item.get('Execution Status') == 'Not Approved')
                # Include N/A as success for success rate calculation
                so_success_rate = ((so_success + so_na) / so_total * 100) if so_total > 0 else 0
                
                # Determine completion status and color
                if so_success_rate == 100:
                    completion_status = "FULLY SUCCESS"
                    status_color = (0, 128, 0)  # Green
                elif so_success_rate == 0:
                    completion_status = "TOTAL FAILED"
                    status_color = (255, 0, 0)  # Red
                else:
                    completion_status = "PARTIAL SUCCESS"
                    status_color = (0, 0, 139)  # Dark Blue
                
                pdf.set_font('Arial', '', 10)
                pdf.cell(0, 5, f'Items: {so_total} | Success: {so_success} | Failed: {so_failed} | N/A: {so_na} | Not Approved: {so_not_approved}', ln=True)
                pdf.cell(0, 5, f'Success Rate: {so_success_rate:.1f}% (includes N/A as success)', ln=True)
                
                # Add completion status with color
                pdf.set_text_color(status_color[0], status_color[1], status_color[2])
                pdf.set_font('Arial', 'B', 10)
                pdf.cell(0, 5, f'Completion Status: {completion_status}', ln=True)
                pdf.set_text_color(0, 0, 0)  # Reset to black
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

    
    def get_error_statistics(self, report_data):
        """
        Get detailed error statistics for debugging purposes
        Updated to handle N/A and Not Approved status separately
        """
        error_stats = {
            "total_errors": 0,
            "total_na": 0,
            "total_not_approved": 0,
            "error_types": defaultdict(int),
            "errors_by_sales_order": defaultdict(list),
            "na_by_sales_order": defaultdict(list),
            "not_approved_by_sales_order": defaultdict(list),
            "most_common_errors": []
        }
        
        for record in report_data:
            status = record.get('Execution Status')
            so_number = record.get('Document Number', 'Unknown')
            
            if status == 'FAILED':
                error_stats["total_errors"] += 1
                error_msg = record.get('Error Message', 'Unknown error')
                error_stats["error_types"][error_msg] += 1
                
                error_stats["errors_by_sales_order"][so_number].append({
                    "logo": record.get('LOGO', 'N/A'),
                    "error": error_msg,
                    "style": record.get('VENDOR STYLE', 'N/A')
                })
            elif status == 'N/A':
                error_stats["total_na"] += 1
                error_stats["na_by_sales_order"][so_number].append({
                    "logo": record.get('LOGO', 'N/A'),
                    "error": record.get('Error Message', 'Invalid Logo SKU'),
                    "style": record.get('VENDOR STYLE', 'N/A')
                })
            elif status == 'Not Approved':
                error_stats["total_not_approved"] += 1
                error_stats["not_approved_by_sales_order"][so_number].append({
                    "logo": record.get('LOGO', 'N/A'),
                    "error": record.get('Error Message', 'Status: Not Approved'),
                    "style": record.get('VENDOR STYLE', 'N/A')
                })
        
        # Sort errors by frequency
        error_stats["most_common_errors"] = sorted(
            error_stats["error_types"].items(), 
            key=lambda x: x[1], 
            reverse=True
        )
        
        return dict(error_stats)