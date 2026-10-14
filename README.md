# Verbo Magico V4

A Streamlit web app that accepts a Spanish infinitive and produces a formatted Excel conjugation workbook.

## Workbook sheets

- Quick reference: infinitive, gerund, participle and five core tenses
- Full conjugation: simple and compound indicative forms, subjunctive, commands and non-finite forms
- All forms: every form returned by the conjugation API
- Sources: hidden source list

## Uploading to GitHub

1. Download and unzip this package.
2. In the GitHub repository, delete the old `streamlit_app.py`, `generator.py`, and `requirements.txt`.
3. Upload the three replacement files from this package. The README is optional.
4. Commit the changes.
5. Streamlit Community Cloud should redeploy automatically. If it does not, open **Manage app** and choose **Reboot app**.

The app requires outbound access to the Verbecc HTTP API. No API key is required.
