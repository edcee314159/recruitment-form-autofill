# Recruitment form autofill

This public repository contains only reusable blank templates and fictional examples for a recruitment-form autofill skill. It contains no candidate records or supporting documents.

Do not commit completed workbooks, attachments, resumes, certificates, photos, or other personal data. Keep local profile data and uploaded files outside this repository; the included `.gitignore` helps prevent common accidental additions.

Start with [快速开始](docs/快速开始.md). Python 3.10+ and the dependencies in `requirements.txt` are required. The installable skill includes its own blank template and [setup instructions](skill/recruitment-form-autofill/references/setup.md). Browser filling also requires an available browser-control tool and the separate `profile-use` companion.

Before any public release, review the staged file list and run:

```powershell
python scripts/prepublish_check.py .
python -m unittest discover -v
git status --short
```

The checker compares all populated cell values and types against approved public fingerprints, rather than trusting a “blank” or “fictional” label. It rejects changed template/example values, added sheets or cells, comments, hyperlinks, local profile configuration, common attachment formats, and workbooks outside the approved locations. It does not print file contents. Review staged files as well; this check is not a general personal-data scanner for every file format.

The optional template generator and its tests currently use Node.js with Codex's bundled `@oai/artifact-tool`. Normal installation and profile use require only Python; the published workbooks are already generated.
