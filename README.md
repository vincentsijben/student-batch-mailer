# student-batch-mailer
a tool to quickly send files to multiple students, matching filenames with student names

Electron app for matching student feedback files to roster entries and mailing them through Outlook.

## Sample data
To create the optional sample set run:

```
python3 scripts/create_sample_set.py your.gmail+alias@gmail.com
```

- Generates 60 PDFs in `sample-set/feedback-files/`
- Generates a matching `sample-set/student-sampleset.xlsx`
- Emails are based on Gmail `+` addressing so all messages land in your inbox
