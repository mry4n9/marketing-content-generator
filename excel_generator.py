# excel_generator.py
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

def _apply_styles(ws, header_fill_color="000000", header_font_color="FFFFFF"):
    """Applies common styling to a worksheet."""
    header_font = Font(color=header_font_color, bold=True)
    header_fill = PatternFill(start_color=header_fill_color, end_color=header_fill_color, fill_type="solid")
    center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    general_alignment = Alignment(horizontal="left", vertical="top", wrap_text=True) # Left align for most content
    
    thin_border_side = Side(style='thin')
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    for row_idx, row in enumerate(ws.iter_rows()):
        for cell in row:
            cell.border = cell_border
            if row_idx == 0:  # Header row
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = center_alignment
            else: # Data rows
                cell.alignment = general_alignment
                if cell.column_letter == 'A': # Center align 'Version #' or first column
                    cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)


    # Adjust column widths (simple auto-fit based on header or sample content)
    for col_idx, column_cells in enumerate(ws.columns):
        max_length = 0
        column_letter = openpyxl.utils.get_column_letter(col_idx + 1)
        
        # Check header length
        if ws.cell(row=1, column=col_idx + 1).value:
            max_length = len(str(ws.cell(row=1, column=col_idx + 1).value))

        # Check a few data cells for length
        for i in range(1, min(6, ws.max_row + 1)): # Check up to 5 data rows
             cell_value = ws.cell(row=i, column=col_idx + 1).value
             if cell_value:
                cell_len = len(str(cell_value).split('\n')[0]) # Consider first line for width
                max_length = max(max_length, cell_len)
        
        adjusted_width = (max_length + 5)
        if column_letter in ['D', 'E', 'F', 'G', 'H'] and ws.title in ["Email", "LinkedIn", "Facebook"]: # Wider for body/text fields
             adjusted_width = max(adjusted_width, 40) # Ensure body text columns are wider
        ws.column_dimensions[column_letter].width = min(adjusted_width, 60) # Cap max width


def create_excel_workbook(all_content_data, scraped_info, company_name_for_file):
    """Creates an Excel workbook with all generated content and styling."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Remove default sheet

    # --- Email Sheet ---
    if "email" in all_content_data and all_content_data["email"]:
        ws_email = wb.create_sheet(title="Email")
        email_headers = ["Version #", "Objective", "Headline", "Subject Line", "Body", "CTA"]
        ws_email.append(email_headers)
        for item in all_content_data["email"]:
            ws_email.append([
                item.get("Version #"), item.get("Objective"), item.get("Headline"),
                item.get("SubjectLine"), item.get("Body"), item.get("CTA")
            ])
        _apply_styles(ws_email)

    # --- LinkedIn Sheet ---
    if "linkedin" in all_content_data and all_content_data["linkedin"]:
        ws_linkedin = wb.create_sheet(title="LinkedIn")
        linkedin_headers = ["Version #", "Ad Name", "Objective", "Introductory Text", "Image Copy", "Headline", "Destination", "CTA Button"]
        ws_linkedin.append(linkedin_headers)
        for item in all_content_data["linkedin"]:
            ws_linkedin.append([
                item.get("Version #"), item.get("AdName"), item.get("Objective"), item.get("IntroductoryText"),
                item.get("ImageCopy"), item.get("Headline"), item.get("Destination"), item.get("CTAButton")
            ])
        _apply_styles(ws_linkedin)

    # --- Facebook Sheet ---
    if "facebook" in all_content_data and all_content_data["facebook"]:
        ws_facebook = wb.create_sheet(title="Facebook")
        facebook_headers = ["Version #", "Ad Name", "Objective", "Primary Text", "Image Copy", "Headline", "Link Description", "Destination", "CTA Button"]
        ws_facebook.append(facebook_headers)
        for item in all_content_data["facebook"]:
            ws_facebook.append([
                item.get("Version #"), item.get("AdName"), item.get("Objective"), item.get("PrimaryText"),
                item.get("ImageCopy"), item.get("Headline"), item.get("LinkDescription"),
                item.get("Destination"), item.get("CTAButton")
            ])
        _apply_styles(ws_facebook)

    # --- Google Search Sheet ---
    if "google_search" in all_content_data and all_content_data["google_search"]:
        gs_data = all_content_data["google_search"]
        ws_gs = wb.create_sheet(title="Google Search")
        gs_headers = ["Type", "Text"] # Simplified for varying numbers
        ws_gs.append(gs_headers)
        for headline in gs_data.get("headlines", []):
            ws_gs.append(["Headline", headline])
        for description in gs_data.get("descriptions", []):
            ws_gs.append(["Description", description])
        _apply_styles(ws_gs)
        # Custom column widths for Google Search
        ws_gs.column_dimensions['A'].width = 20
        ws_gs.column_dimensions['B'].width = 50


    # --- Google Display Sheet ---
    if "google_display" in all_content_data and all_content_data["google_display"]:
        gd_data = all_content_data["google_display"]
        ws_gd = wb.create_sheet(title="Google Display")
        gd_headers = ["Type", "Text"]
        ws_gd.append(gd_headers)
        for headline in gd_data.get("headlines", []):
            ws_gd.append(["Headline", headline])
        for description in gd_data.get("descriptions", []):
            ws_gd.append(["Description", description])
        _apply_styles(ws_gd)
        # Custom column widths for Google Display
        ws_gd.column_dimensions['A'].width = 20
        ws_gd.column_dimensions['B'].width = 50


    # --- Reasoning Sheet ---
    ws_reasoning = wb.create_sheet(title="Reasoning")
    ws_reasoning.append(["Scraped Client Data & Generation Reasoning"])
    ws_reasoning.merge_cells('A1:B1') # Merge for a title cell
    
    # Apply header style to the merged title cell
    title_cell = ws_reasoning['A1']
    title_cell.font = Font(color="FFFFFF", bold=True, size=14)
    title_cell.fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    title_cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))


    current_row = 2
    scraped_data_map = {
        "Company Name": scraped_info.get('company_name'),
        "Tagline": scraped_info.get('tagline'),
        "Mission Statement": scraped_info.get('mission_statement'),
        "Industry": scraped_info.get('industry'),
        "Products/Services": ", ".join(scraped_info.get('products_services', [])) if scraped_info.get('products_services') else 'N/A',
        "USPs/Value Proposition": scraped_info.get('usps_value_proposition'),
        "Target Audience": scraped_info.get('target_audience'),
        "Tone of Voice": scraped_info.get('tone_of_voice'),
        "CTAs from Website": ", ".join(scraped_info.get('ctas', [])) if scraped_info.get('ctas') else 'N/A'
    }

    for key, value in scraped_data_map.items():
        ws_reasoning.cell(row=current_row, column=1, value=key).font = Font(bold=True)
        ws_reasoning.cell(row=current_row, column=2, value=str(value))
        # Apply styles to these cells
        for col in [1,2]:
            cell = ws_reasoning.cell(row=current_row, column=col)
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        current_row += 1
    
    ws_reasoning.cell(row=current_row, column=1, value="").font = Font(bold=True) # Spacer row
    current_row +=1

    ws_reasoning.cell(row=current_row, column=1, value="Reasoning for Content Generation:").font = Font(bold=True, size=12)
    ws_reasoning.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
    merged_reasoning_header = ws_reasoning.cell(row=current_row, column=1)
    merged_reasoning_header.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    merged_reasoning_header.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    current_row += 1
    
    reasoning_text_cell = ws_reasoning.cell(row=current_row, column=1, value=all_content_data.get("reasoning_text", "Reasoning not available."))
    ws_reasoning.merge_cells(start_row=current_row, start_column=1, end_row=current_row + 5, end_column=2) # Merge for text block
    reasoning_text_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    # Apply border to the merged reasoning text cell block
    # This requires iterating through the cells in the merged range if openpyxl doesn't do it automatically for the top-left cell
    # For simplicity, styling the top-left cell of a merged range usually applies to the whole merged area visually.
    # To be explicit for borders:
    merged_range_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    for r in range(current_row, current_row + 6): # Iterate rows in merged cell
        for c_idx in [1,2]: # Iterate columns in merged cell
            # Check if cell is part of the merged range before styling (safer)
            # However, styling the top-left cell (reasoning_text_cell) is often sufficient for visual aspects.
            # For borders, it's better to apply to the range or ensure the top-left cell's border covers it.
            # The border on reasoning_text_cell should visually apply.
            pass # Border is applied to reasoning_text_cell, which is the top-left of the merge.

    ws_reasoning.column_dimensions['A'].width = 30
    ws_reasoning.column_dimensions['B'].width = 70
    ws_reasoning.row_dimensions[current_row].height = 100 # Make reasoning text area taller


    # Save to a BytesIO object
    excel_bytes = io.BytesIO()
    wb.save(excel_bytes)
    excel_bytes.seek(0)
    return excel_bytes