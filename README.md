# Recruitment form autofill

This public repository contains only reusable blank templates and fictional examples for a recruitment-form autofill skill. It contains no candidate records or supporting documents.

Do not commit completed workbooks, attachments, resumes, certificates, photos, or other personal data. Keep local profile data and uploaded files outside this repository; the included `.gitignore` helps prevent common accidental additions.

Before any public release, review the staged file list and run:

```powershell
python scripts/prepublish_check.py .
python -m unittest discover -v
git status --short
```

The checker rejects completed personal workbooks, local profile configuration, common attachment formats, and workbooks outside the published blank-template and fictional-example locations. It does not print file contents.
