# CV Profile Picture Extractor

Tools to extract likely profile photos from CVs (`.pdf` and `.docx`), crop them to balanced squares, and save the results.

## Run Without Streamlit (Desktop App)

```powershell
pip install -r requirements-desktop.txt
python cv_profile_extractor_tk.py
```

This opens a native desktop window (Tkinter) with:
- `Add CV Files`
- `Extract Profile Pictures`
- `Save This PNG`
- `Save All PNGs`
- `Save ZIP`

## Optional Streamlit Version

```powershell
pip install -r requirements.txt
streamlit run cv_profile_extractor_app.py
```

## Features

- PDF and DOCX image extraction
- Face-aware profile-image selection
- Square face-centered crop with margin for nicer framing
- Save individual PNGs or all outputs as ZIP
