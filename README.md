# Soft Slides

Streamlit app that turns a tagged PowerPoint template + an Excel sheet + speaker photos into a ready-to-use deck of speaker/session slides.

## How the template works
The template `.pptx` holds your **layouts** (they are removed from the final deck):

- **Slide 1** - the *individual* layout (one speaker). Used for every standalone row, and for each speaker's own slide after a panel.
- **Slides 2, 3, 4...** - optional *panel* layouts with room for several speakers (e.g. 3, 5, 9 slots). For each panel the app picks the **smallest layout that fits** the group, followed by one individual slide per speaker.

Type tags inside text boxes (any shape works for the photo placeholder - circle, rounded rectangle, rectangle):

| Tag | Meaning |
|---|---|
| `<<SESSION_NAME>>`, `<<HALL_NAME>>`, `<<DATE>>` | Session-level text |
| `<<MAIN_SESSION_DETAILS>>`, `<<SPEAKER_SESSION_DETAILS>>` | Session / talk details |
| `<<PLACEHOLDER_1>>`, `<<PLACEHOLDER_2>>` | Free text (per row on individual slides) |
| `<<SPEAKER_NAME_n>>`, `<<SPEAKER_TITLE_n>>`, `<<SPEAKER_COMPANY_n>>` | Speaker n text |
| `<<SPEAKER_ROLE_n>>` | Role label above photo n (Chair, Moderator, Keynote, Panelists...) |
| `<<SPEAKER_PHOTO_n>>` | Photo placeholder n - the photo takes this shape's outline |

## Workflow
1. Upload the template. The app detects the tags and speaker slots.
2. Download the **Excel data template**, fill it (one row per speaker), upload it back.
3. Upload speaker photos. Photos are matched to the `SPEAKER_PHOTO_FILENAME_1` column by file name (extension not needed). The same photo may be used on many rows.
4. Generate and download the deck.

## Excel notes
- **Require Single Slide**: rows sharing a label (e.g. `Panel 1`) form one panel slide, followed by one individual slide per speaker. Leave blank for a standalone slide.
- **SPEAKER_ROLE_1** has a dropdown fed by the **Lists** sheet (add your own roles in column A). Column B is the label used when consecutive speakers share a role - by default *Panelist* becomes **Panelists**, shown once above the first panelist (drag it wherever you like). Leave column B blank for roles that always get their own label.
- Keep rows in stage order: a shared label is applied to *consecutive* speakers with the same role.
- Unused speaker slots on a larger panel layout are cleaned up automatically.

## Run locally
    pip install -r requirements.txt
    streamlit run app.py

The sidebar shows the running **core build**; if it shows a red "OLD" warning, the `core/` folder was not fully replaced - replace it and restart the app.
