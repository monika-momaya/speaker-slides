import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from core.pptxtagparser import TemplateTagReport, KNOWN_SINGLE_TAGS

def build_excel_template_bytes(report: TemplateTagReport) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "SlideData"
    columns = []
    for tag in KNOWN_SINGLE_TAGS:
        if tag in report.found_single_tags:
            columns.append(tag)
    for slot in range(1, report.max_speaker_slot + 1):
        columns += [f"SPEAKER_NAME_{slot}", f"SPEAKER_TITLE_{slot}", f"SPEAKER_COMPANY_{slot}", f"SPEAKER_PHOTO_FILENAME_{slot}"]
    if report.max_speaker_slot > 0:
        columns.append("SPEAKER_ROLE_1")
    columns.append("Require Single Slide")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for col_idx, col_name in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        ws.column_dimensions[cell.column_letter].width = max(18, len(col_name) + 2)
    notes_ws = wb.create_sheet("Instructions")
    notes = [
        "Each row = one speaker.",
        "Leave Require Single Slide blank for one speaker per slide.",
        "Use the same label, e.g. Panel 1, for rows that belong on the same slide.",
        "Fill session-level fields across grouped rows as needed.",
        "Role labels (Moderator, Chair, Keynote...): fill SPEAKER_ROLE_1 (or PLACEHOLDER_1) on each speaker's own row.",
    ]
    for i, line in enumerate(notes, start=1):
        notes_ws.cell(row=i, column=1, value=line)
    notes_ws.column_dimensions["A"].width = 100
    # Lists sheet: feeds the Role dropdown. Edit freely - add roles in column A.
    # Column B = label used when several consecutive speakers share the role
    # (e.g. Panelist -> "Panelists"). Leave B blank for roles that always get their own label.
    lists_ws = wb.create_sheet("Lists")
    lists_ws.cell(row=1, column=1, value="Role (dropdown)").font = Font(bold=True)
    lists_ws.cell(row=1, column=2, value="Shared label (when several in a row)").font = Font(bold=True)
    for i, (role, shared) in enumerate([("Chair", None), ("Moderator", None), ("Keynote", None), ("Panelist", "Panelists")], start=2):
        lists_ws.cell(row=i, column=1, value=role)
        if shared:
            lists_ws.cell(row=i, column=2, value=shared)
    lists_ws.column_dimensions["A"].width = 24
    lists_ws.column_dimensions["B"].width = 38
    if "SPEAKER_ROLE_1" in columns:
        col_letter = ws.cell(row=1, column=columns.index("SPEAKER_ROLE_1") + 1).column_letter
        dv = DataValidation(type="list", formula1="=Lists!$A$2:$A$30", allow_blank=True, showErrorMessage=False)
        dv.promptTitle = "Role"
        dv.prompt = "Pick a role, or type your own."
        ws.add_data_validation(dv)
        dv.add(f"{col_letter}2:{col_letter}500")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
