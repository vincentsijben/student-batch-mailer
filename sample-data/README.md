# Sample data

Ready-made files to try out Student Batch Mailer without preparing your own data.

| What | File(s) |
| --- | --- |
| Student roster (Excel) | [`student-roster.xlsx`](./student-roster.xlsx) |
| Feedback files to drag & drop | [`feedback-files/`](./feedback-files/) (8 PDFs) |

## How to try it

1. Start the app (`pnpm run start`, or open the built `Student Batch Mailer.app`).
2. **Step 1 – Student roster:** drop [`student-roster.xlsx`](./student-roster.xlsx) on the roster box.
3. **Step 2 – Feedback files:** drag all PDFs from [`feedback-files/`](./feedback-files/) (or the whole folder) onto the files box.
   Each file gets a badge showing which student it matched (e.g. `feedback-emma-anderson.pdf` → *Emma Anderson*).
   All 8 should match. Use the **×** next to a file to remove it, or **Clear all** to start over.
4. **Step 3 – Message:** pick a template or write a subject/body. `{{firstname}}` etc. are filled in per student.
5. **Step 4 – Review & send:** click rows to (de)select students, then press **Send N emails**.

> **About the email addresses:** the roster uses Gmail `+` aliases
> (`vinzzz81+emmaanderson@gmail.com`, `vinzzz81+liambennett@gmail.com`, …), so
> every test mail lands in the `vinzzz81@gmail.com` inbox. To receive them
> yourself, open `student-roster.xlsx` and replace `vinzzz81` with your own Gmail
> local part.

## Roster format

The Excel sheet needs three columns; header names may be English or Dutch:

| firstname / voornaam | lastname / achternaam | email |
| --- | --- | --- |
| Emma | Anderson | vinzzz81+emmaanderson@gmail.com |

After sending, a timestamped log is written automatically — use **Open sent logs folder** in the app to see it.

## Files

- [`feedback-emma-anderson.pdf`](./feedback-files/feedback-emma-anderson.pdf)
- [`feedback-liam-bennett.pdf`](./feedback-files/feedback-liam-bennett.pdf)
- [`feedback-olivia-carter.pdf`](./feedback-files/feedback-olivia-carter.pdf)
- [`feedback-noah-diaz.pdf`](./feedback-files/feedback-noah-diaz.pdf)
- [`feedback-sophia-foster.pdf`](./feedback-files/feedback-sophia-foster.pdf)
- [`feedback-lucas-garcia.pdf`](./feedback-files/feedback-lucas-garcia.pdf)
- [`feedback-mia-patel.pdf`](./feedback-files/feedback-mia-patel.pdf)
- [`feedback-ethan-reynolds.pdf`](./feedback-files/feedback-ethan-reynolds.pdf)

Need a bigger set (60 students) that mails to your own Gmail? See
[`scripts/create_sample_set.py`](../scripts/create_sample_set.py).
