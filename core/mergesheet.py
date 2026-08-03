from typing import List, Dict
import openpyxl
SINGLE_FIELD_KEYS = ["SESSION_NAME", "HALL_NAME", "DATE", "MAIN_SESSION_DETAILS", "SPEAKER_SESSION_DETAILS", "PLACEHOLDER_1", "PLACEHOLDER_2"]

def read_merge_sheet(file_obj, max_speaker_slot: int):
    wb = openpyxl.load_workbook(file_obj, data_only=True)
    ws = wb["SlideData"] if "SlideData" in wb.sheetnames else wb.active
    headers = [c.value for c in ws[1]]
    header_index = {h: i for i, h in enumerate(headers) if h}
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row is None or all(v is None or str(v).strip() == "" for v in row):
            continue
        row_dict = {}
        for h, idx in header_index.items():
            row_dict[h] = row[idx] if idx < len(row) else None
        rows.append(row_dict)
    groups: List[Dict] = []
    group_index_by_label: Dict[str, int] = {}
    for row in rows:
        single_slide_label = row.get("Require Single Slide")
        single_slide_label = str(single_slide_label).strip() if single_slide_label else ""
        speaker_entry = {
            "name": row.get("SPEAKER_NAME_1") or None,
            "title": row.get("SPEAKER_TITLE_1") or "",
            "company": row.get("SPEAKER_COMPANY_1") or "",
            "photo_key": row.get("SPEAKER_PHOTO_FILENAME_1") or None,
        }
        if single_slide_label:
            if single_slide_label in group_index_by_label:
                gi = group_index_by_label[single_slide_label]
                groups[gi]["speakers"].append(speaker_entry)
            else:
                new_group = {k: row.get(k, "") for k in SINGLE_FIELD_KEYS}
                new_group["speakers"] = [speaker_entry]
                new_group["_label"] = single_slide_label
                groups.append(new_group)
                group_index_by_label[single_slide_label] = len(groups) - 1
        else:
            new_group = {k: row.get(k, "") for k in SINGLE_FIELD_KEYS}
            new_group["speakers"] = [speaker_entry]
            new_group["_label"] = None
            groups.append(new_group)
    overflow_warnings = []
    for g in groups:
        if len(g["speakers"]) > max_speaker_slot and max_speaker_slot > 0:
            label = g.get("_label") or "(individual slide)"
            overflow_warnings.append(f"Group '{label}' has {len(g['speakers'])} speakers but the template only supports {max_speaker_slot} speaker slot(s). Extra speakers will be dropped.")
            g["speakers"] = g["speakers"][:max_speaker_slot]
    return groups, overflow_warnings
