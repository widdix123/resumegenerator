#!/bin/bash
# Double-click to rebuild every resume from resume.md.
cd "$(dirname "$0")" || exit 1
set -e
echo "Building resumes from resume.md ..."
python3 scripts/create_designed_resume.py
python3 scripts/create_ats_resume.py
python3 scripts/create_designed_docx.py
echo
echo "Done:"
echo "  Abhishek_Gupta_Resume_Designed.pdf   <- send to people (the nice one)"
echo "  Abhishek_Gupta_Resume_ATS.pdf/.docx  <- upload to job portals"
echo "  Abhishek_Gupta_Resume_Designed.docx  <- only if someone insists on Word"
open Abhishek_Gupta_Resume_Designed.pdf
