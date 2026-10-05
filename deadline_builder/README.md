# CDEF deadline summary

Open **CDEF Status Summary.xlsx** and start with **Summary**. It contains counts by contractor, a filterable CDEF register, a date schedule, deadline entries needing review, and source references.

The report uses one row per CDEF from visible source rows. Hidden rows are excluded from deadline selection, status/contact lookup, Review and Source deadlines. Save the source workbook after changing row visibility so the script reads the updated state.

Only visible contractor-list deadlines contribute to the default summary. Synergy scope takes priority where both Synergy worksheets give a visible deadline. The Data export never adds a CDEF or supplies a summary deadline, even when a CDEF is absent from the contractor lists. It can supply missing subject/contact/status details for an existing selected CDEF, and its visible deadline entries remain in Source deadlines as references. A visible copy on another contractor sheet remains eligible. Blank deadlines and Minimax are excluded. Non-date notes appear in Review.

The optional `--policy all` explicitly includes every visible source deadline, including Data records and duplicates; it is not used by the default command below.

The script recognizes `Deadline` and `Correction Deadline` at any column/header row, including bilingual labels such as `Correction Deadline / 纠正截止日期`, `No. / 编号` and `Subject / 主题`. Other deadline fields such as `Deadline2` and `Construction Deadline` are not substituted.

CREC numeric 9.2 and 9.3 mean 20 and 30 September 2026, as confirmed. The ten `9..15` entries are interpreted as 15 September 2026 and annotated.

Change the blue **Summary!B3** date to recalculate deadline status. The source export status is dated 31 August 2026; overdue counts indicate deadline timing, not verified completion.

After updating the source workbook, refresh the report from this folder:

```powershell
python build_summary.py
```

Requires Python and `openpyxl`. Rebuilding replaces **CDEF Status Summary.xlsx**. Close it in Excel before rebuilding; save any manual edits separately. The source **CDEF Deadline.xlsx** is read only by the script.
