# Student Batch Mailer

Electron app for matching student feedback files to roster entries and mailing them through Outlook.

## Platform support
- The distributed build targets macOS on Apple Silicon (`--platform=darwin --arch=arm64`), i.e., M1/M2/M3 machines. Intel Macs require rebuilding with `npm run dist -- --arch=x64` or similar.
- Microsoft Outlook for macOS must be installed and allowed to run AppleScript, because sending relies on `osascript outlook.scpt`.
- Uploads are cached under `~/Library/Application Support/Student Batch Mailer/upload-cache/` for the duration of each session so pathless drag-and-drop files can be attached.

### Quick self-checks
- `uname -m` should print `arm64` to use the prebuilt binary. If it prints `x86_64`, rebuild with `npm run dist -- --arch=x64`.
- `osascript -e 'id of application "Microsoft Outlook"'` should print the Outlook bundle id (e.g., `com.microsoft.Outlook`) to confirm Outlook is installed and AppleScript-accessible.

## Sample data
To create the optional sample set run:

```
python3 scripts/create_sample_set.py your.gmail+alias@gmail.com
```

- Generates 60 PDFs in `sample-set/feedback-files/`
- Generates a matching `sample-set/student-sampleset.xlsx`
- Emails are based on Gmail `+` addressing so all messages land in your inbox
