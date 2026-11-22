# RUAS Student Attendance & Learning Dashboard

A browser-only dashboard for the Directorate / departments of Ramaiah University of Applied Sciences. Upload an anonymised Excel file of attendance and marks, and the dashboard shows:

- **★ Fast learners, ▲ slow learners, ◆ counselling and ⚑ parent notification**, based on adjustable rules
- **Attendance vs score scatter plot** with the decision thresholds drawn as lines, so anyone can read who needs what
- **Decision mix** by program, department, semester or course
- **Why attendance matters**: average score in each attendance band
- **Faculty view**: every class's attendance, score, pass % and number of students needing support, compared with the overall average
- **Action lists**, a per-student detail view with course-by-course bars, a draft parent letter, and **export to Excel**

Data never leaves the browser.

## Input format
Use `RUAS_Attendance_Anonymised_Data.xlsx` as the template. In the `Attendance_Data` sheet, each row is one student in one theory course:

Academic Year · Faculty · Dept · Program · Name of Student · Registration · Sem · Course (Theory) · Lab · Theory Attendance % · Lab Attendance % · CIE Mark · SEE Mark · Quiz Mark

The `Settings` sheet holds the maximum marks, weights and thresholds. The dashboard reads them, and you can also change them in the dashboard's **Decision rules** panel. CSV files with the same headers also work.

## Publish on GitHub Pages
Upload all files (there are no folders) to a public repository, then go to **Settings → Pages** and set *Deploy from a branch* → `main` → `/ (root)`. The **Use sample data** button works once the site is published.

## Important
The thresholds are example values. Set them to RUAS policy. The decisions are suggestions: mentors and the HoD review each case before contacting students or parents. Use anonymised data for training sessions.
