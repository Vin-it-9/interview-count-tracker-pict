
# Campus Interview Appearance Tracker

A high performance Java application that processes multiple Excel files to track how many times each student has appeared for campus interviews across different companies.

## Features

- ✅ Processes multiple Excel (.xlsx) files simultaneously using parallel processing
- ✅ Handles flexible column names (Name, Full Name, Student Name, Candidate Name, Email, etc.)
- ✅ Smart header detection - finds headers even if they're not in row 1
- ✅ Two-pass processing to prevent duplicate entries
- ✅ Handles files with or without email columns
- ✅ Generates professional Excel report with formatting
- ✅ Thread-safe concurrent processing for maximum speed
- ✅ Detailed progress tracking and performance metrics


## Project Structure

```
interview-tracker/
├── src/main/java/
│   ├── Main.java
│   ├── model/
│   │   └── StudentRecord.java
│   ├── processor/
│   │   ├── ExcelFileProcessor.java
│   │   └── SheetProcessor.java
│   ├── service/
│   │   └── InterviewTrackerService.java
│   ├── util/
│   │   ├── ExcelUtils.java
│   │   └── PerformanceMonitor.java
│   └── writer/
│       └── ReportGenerator.java
└── pom.xml
```


### Step 3: Check the Output

The application generates `Interview_Appearance_Report.xlsx` in the project root with:
- Rank column
- Student names
- Email addresses
- Total interview count
- Professional formatting with headers and borders

## How It Works

### Phase 1: Process Files with Email Columns
- Reads all Excel files that have email columns
- Builds a name-to-email mapping
- Tracks each student's appearances

### Phase 2: Process Files without Email Columns
- Reads files that only have name columns
- Looks up existing emails from Phase 1
- Only generates new emails if name is completely new
- Prevents duplicate entries

### Smart Features

**Flexible Column Detection:**
- Automatically detects columns like: Name, Full Name, Student Name, Candidate Name, Applicant Name
- Finds Email, E-mail, Email Address, Mail columns
- Searches up to 20 rows to find the header row

**Name Normalization:**
- "Vineet Shinde" and "VINEET SHINDE" are treated as the same person
- Handles extra spaces and special characters
- Ensures no duplicates across files

**Multi-Sheet Support:**
- Processes all sheets in each Excel file
- Handles empty sheets gracefully
- Skips duplicate entries within the same file

## Performance

- Uses configurable thread pool (default: CPU cores × 2)
- Processes 20+ files in seconds
- Memory efficient with streaming
- Thread-safe concurrent operations

## Output Format

The Excel report includes:
- **Rank**: Position based on interview count
- **Name**: Student's full name
- **Email**: Student's email address (real or generated)
- **Total_Count**: Number of interviews attended

Students are sorted by interview count (highest to lowest).

## Example Output

```

════════════════════════════════════════════════════════════
     CAMPUS INTERVIEW APPEARANCE TRACKER v2.0
════════════════════════════════════════════════════════════

📁 Folder: input_files
📊 Found 13 Excel file(s)
🔄 Processing with 16 threads

────────────────────────────────────────────────────────────

🔄 PHASE 1: Processing files with email columns...

──────────────────────────────────────────────────────────────────────
PHASE 1 RESULTS
──────────────────────────────────────────────────────────────────────
✓ Altizon Shortlist.xlsx: 28 students from 4 sheet(s)
✓ PICT - Interview shortlist.xlsx: 28 students from 1 sheet(s)
✓ PICT - Shortlists..xlsx: 52 students from 1 sheet(s)
✓ PICT shortlist (1).xlsx: 30 students from 1 sheet(s)
✓ PICT Shortlist for Interviews.xlsx: 66 students from 1 sheet(s)
✓ SAP BASIS Round-1_Responses_Shortlisted_For_GD.xlsx: 73 students from 1 sheet(s)
✓ Shortlist for ProcDNA PICT.xlsx: 87 students from 2 sheet(s)
✓ Shortlisted ACi PICT (1).xlsx: 70 students from 2 sheet(s)
✓ Shortlists-Final.xlsx: 41 students from 2 sheet(s)

🔄 PHASE 2: Processing files without email columns...

──────────────────────────────────────────────────────────────────────
PHASE 2 RESULTS
──────────────────────────────────────────────────────────────────────
✓ 1st_ShortList.xlsx: 55 students from 1 sheet(s)
✓ Sell.Do.xlsx: 75 students from 1 sheet(s)

════════════════════════════════════════════════════════════
FINAL REPORT (Sorted: Highest → Lowest Interview Count)
════════════════════════════════════════════════════════════
Rank Name                         Email                    Count
────────────────────────────────────────────────────────────
1    Dhananjay Suresh Boob        dhananjayboob@gmail.com  5
2    Manaswi Gaurishankar Hire... manaswi.1703@gmail.com   4
3    Swaraj Santosh Nalawade      swaraj.nalawade04@gma... 4
4    Pranav Vijay Hire            pranavhire0990@gmail.com 4
5    Girish Jaykumar Kale         girishjaykumarkale@gm... 4
6    Pushkar Santosh Yewale       pushkaryewale1817@gma... 4
7    Rutuja Annasaheb Gangawane   rutujagangawane21@gma... 4
8    Jagdish Dilip Bainade        www.jdb9860249117@gma... 3
9    Ayush Sanjay Agrawal         ayushsagrawal7249@gma... 3
10   Anshul Purandare             anshul.purandare@gmai... 3
... and 423 more
════════════════════════════════════════════════════════════

✓ Report exported to: Interview_Appearance_Report.xlsx
  Total unique students: 433

════════════════════════════════════════════════════════════
PERFORMANCE SUMMARY
════════════════════════════════════════════════════════════
Files Processed:    11
Students Processed: 605
Time Elapsed:       2.643 seconds
Throughput:         5.50 files/sec
════════════════════════════════════════════════════════════
Name-to-Email mappings created: 454

```
