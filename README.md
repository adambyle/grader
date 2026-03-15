# Grader

A lightweight vibe-coded grading tool for Moodle-exported CSV files. Import a class roster, define reusable feedback, and export a completed grade sheet.

---

> **Testing wanted:** The automatic late-submission detection reads the *Last modified (submission)* date column from Moodle exports. If you have a real or anonymized CSV that includes completion dates in that column, please try the late policy feature and report any issues with how days-late are calculated. The sample CSV included in this repo (`sample_input.csv`) can be used if you don't have one handy.

---

## Workflow

**1. Import a CSV**

Click *Import CSV* and select the Moodle grade export file. The app reads student names, emails, and max grades. Submission dates are parsed automatically when present.

**2. Define reusable feedback**

Use the left panel to create feedback items — things you'll apply to multiple students, like "improper formatting" or "missing step 4". Each item has a point value (negative for deductions, positive for bonuses) and a text description.

Drag items by the ⠿ handle to reorder them. Drag an item directly onto a row in the student table to apply it to that student without opening their detail pane.

**3. Grade submissions**

Click a student's row to open their detail pane on the right. Check the boxes next to feedback items to apply them — the grade updates live. You can also:

- Edit the point value or text of an applied item just for that student without affecting the global definition. A ↺ button appears to revert back.
- Add one-off notes under *One-off notes* for feedback that doesn't belong in the reusable list.
- Click *✓ Perfect* to mark a submission as full credit with the configured perfect auto-text (e.g. "Great work!").
- Lock the grade manually with the 🔒 button if you need to override the computed total.

To apply feedback to multiple students at once, Ctrl+click or Shift+click rows to build a selection. The right panel switches to a group view where you can apply or remove feedback items across all selected students in one step.

**4. Configure late policy**

Expand the *Late policy* section in the left panel. Choose a policy type:

- **% per day** — deducts a percentage of the max grade for each day late
- **Points per day** — deducts a flat number of points per day
- **Zero if late** — any late submission receives zero
- **None** — no late penalty (default)

Set a deadline. If submission dates are present in the CSV, days-late values are computed automatically. Otherwise, you can enter days late manually per student in their detail pane. The late penalty entry appears automatically in the feedback summary and the exported CSV — no manual step required. You can waive the penalty for individual students with the *waive* checkbox.

**5. Export**

Click *Export CSV* to download a file ready to import back into Moodle. The Feedback comments column is formatted to match Moodle's expected input.

## Saving and loading

The app saves to your browser's local storage after every change, so a session survives an accidental page refresh. For more permanent storage, use *Save* (or Ctrl+S) to write a `.json` project file to disk. Once saved, subsequent changes auto-save to the same file. Use *Load* to reopen a saved project.

## Tips

- Press **Enter** in the feedback item list to create the next item without clicking the + button.
- Press **Escape** to unfocus any input field.
- The *Feedback* column in the student table shows a one-line summary of applied feedback. It updates live as you grade.
- Feedback items can be highlighted — click the count badge on any item in the left panel to highlight every student row that item is applied to.
- The left and right panels are resizable. Drag the thin bar between panels to adjust.
