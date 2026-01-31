def create_detailed_summary_table(prs, entities_data, date_range):
    """Slide 8: Detailed Summary Table - Matching template exactly"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Slide title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = f"Summary of Alerts\n({date_range})"
    title_frame.paragraphs[0].font.size = Pt(24)
    title_frame.paragraphs[0].font.bold = True
    title_frame.paragraphs[0].font.color.rgb = COLORS['accent']
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # Build detailed table matching template exactly
    table_data = [[
        'S.no', 'Entity', 'Total Incidents count ', 'Closed Incident', 'WIP Incidents',
        'Severity wise incidents count', '', 'Count of False Positive Alerts',
        'Count of True Positive Alerts', 'Alert Category ', 'Remarks'
    ]]
    
    s_no = 1
    for entity in ENTITY_ORDER:
        entity_info = entities_data.get(entity, {})
        total = entity_info.get('total', 0)
        all_severities = entity_info.get('all_severities', {})
        closed_dict = entity_info.get('closed', {})
        wip_dict = entity_info.get('WIP', {})
        
        # Calculate totals
        total_closed = sum(closed_dict.values())
        total_wip = sum(wip_dict.values())
        
        # Entity name mapping for Slide 8
        if entity == 'ENT':
            entity_name = 'Enterprise'
        elif entity == 'MRO':
            entity_name = 'HIAL- MRO'
        else:
            entity_name = entity
        
        # Always generate 4 severity rows (Critical, High, Medium, Low)
        for idx, severity in enumerate(SEVERITY_ORDER):
            sev_count = all_severities.get(severity, 0)
            
            # First row for each entity (idx == 0, Critical)
            if idx == 0:
                # FP = total_closed for first row
                fp_count = f"{total_closed:02d}"
                tp_count = "00"
                category = "PUP & Legitimate Non malicious processes"
                remarks = "Malicious exe file accessed  & Legitimate processes triggered."
                
                row = [
                    str(s_no),
                    entity_name,
                    f"{total:02d}",
                    f"{total_closed:02d}",
                    f"{total_wip:02d}",
                    severity,
                    f"{sev_count:02d}",
                    fp_count,
                    tp_count,
                    category,
                    remarks
                ]
            else:
                # Rows 2-4 (High, Medium, Low) - mostly blank
                row = [
                    '',  # S.no blank
                    '',  # Entity blank
                    '',  # Total blank
                    '',  # Closed blank
                    '',  # WIP blank
                    severity,  # Severity name
                    f"{sev_count:02d}",  # Severity count
                    '',  # FP blank
                    '',  # TP blank
                    '',  # Category blank
                    ''   # Remarks blank
                ]
            
            table_data.append(row)
        
        s_no += 1
    
    # Add table with smaller fonts for more data
    table_placeholder = slide.shapes.add_table(
        len(table_data), 11,
        Inches(0.3), Inches(1.0), Inches(12.7), Inches(6)
    )
    table = table_placeholder.table
    
    # Set column widths
    col_widths = [Inches(0.4), Inches(0.8), Inches(1.0), Inches(1.0), Inches(1.0),
                  Inches(1.2), Inches(0.5), Inches(1.0), Inches(1.0),
                  Inches(1.5), Inches(2.3)]
    
    for col_idx, width in enumerate(col_widths):
        table.columns[col_idx].width = width
    
    # Populate table
    for r_idx, row_data in enumerate(table_data):
        for c_idx, cell_value in enumerate(row_data):
            cell = table.cell(r_idx, c_idx)
            cell.text = str(cell_value)
            
            # Format cell
            cell.text_frame.paragraphs[0].font.size = Pt(8)
            cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            
            # Header row
            if r_idx == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = COLORS['header']
                cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
                cell.text_frame.paragraphs[0].font.bold = True
