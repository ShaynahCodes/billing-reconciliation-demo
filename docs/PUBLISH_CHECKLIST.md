# Before publishing this repository

- [ ] Verify the uploaded folder contains only the files in the README tree. **Do not upload the original employer ZIP.**
- [ ] Inspect `config.json`, both demo CSVs, and `TEST_DATA/demo_ship_to_master.json`. Confirm every identity and address is appropriate for public release.
- [ ] Search the `.pyw`, markdown, TXT, CSV, and JSON files for original employer/client names, internal email addresses, real tracking numbers, private addresses, and other confidential information.
- [ ] Verify `.gitignore` is in the root of the Git repository. Confirm no `data/`, `logs/`, exported PDF files, or `.env` files appear in GitHub's staged changes.
- [ ] On Windows, run `py -m unittest discover -s tests -v` and check all five tests pass.
- [ ] On Windows, open the app, browse to both included fictional CSVs, validate, generate PDFs, and review the output. Confirm the bottom action buttons are visible at your display scale.
- [ ] Do not send the Outlook demo draft. Demo recipients use reserved `example.com` addresses.
- [ ] Create a new **private** repository first; review the files and then change visibility to **public** if everything is verified.

**Source control warning:** `.gitignore` prevents untracked files from being added by default. It cannot remove confidential files already committed in Git history. Always initialize a fresh public repository using only the sanitized files.
