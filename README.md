# Harvard Citizenship Tutor Ranker

A static GitHub Pages website for ranking Harvard IOP Citizenship Tutoring Program tutors by seniority from a Google Forms Excel export.

## What it does

1. Uploads an `.xlsx`, `.xls`, or `.csv` file downloaded from Google Forms.
2. Reads the spreadsheet in the browser using SheetJS.
3. Detects relevant Google Form columns even when the column order changes.
4. Calculates seniority fields.
5. Sorts tutors by the Citizenship Program ranking rules.
6. Shows a preview table.
7. Downloads a ranked Excel workbook.

The app is browser-only. It does not upload spreadsheet data to any server.

## Files

```text
citizenship-tutor-ranker/
├── index.html
├── style.css
├── script.js
└── README.md
```

## How to publish on GitHub Pages

1. Create a new GitHub repository, for example `citizenship-tutor-ranker`.
2. Upload `index.html`, `style.css`, `script.js`, and `README.md` to the repository root.
3. Go to **Settings** > **Pages**.
4. Under **Build and deployment**, choose:
   - Source: **Deploy from a branch**
   - Branch: **main**
   - Folder: **/root**
5. Save.
6. GitHub will provide a public website URL after deployment.

## Required Google Form fields

The ranking needs at minimum:

- Forename
- Surname
- Class Year

The app can still use the rest of the ranking system when these fields are present:

- Birthdate
- Language(s) Other than English
- Current Board Member
- Current Board Position
- Former Board Member
- Former Highest Position
- Years Ago Selected
- Returning Member

## Ranking rules

1. Director? TRUE first
2. Secretary? TRUE first
3. Current Board Member? TRUE first
4. Former Director? TRUE first
5. Former Secretary? TRUE first
6. Former Board Member? TRUE first
7. Years since Director, largest to smallest
8. Years since Secretary, largest to smallest
9. Years since Board Member, largest to smallest
10. Returning Tutor? TRUE first
11. Class Year, smallest to largest
12. Number of LOTEs, largest to smallest
13. Spanish? TRUE first
14. Mandarin? TRUE first
15. Haitian Creole? TRUE first
16. French? TRUE first
17. Cantonese? TRUE first
18. Vietnamese? TRUE first
19. Arabic? TRUE first
20. Russian? TRUE first
21. Italian? TRUE first
22. Hindi? TRUE first
23. Cape Verdean Creole? TRUE first
24. Korean? TRUE first
25. German? TRUE first
26. Telugu? TRUE first
27. Amharic? TRUE first
28. Tagalog? TRUE first
29. Japanese? TRUE first
30. Urdu? TRUE first
31. Albanian? TRUE first
32. Persian? TRUE first
33. Thai? TRUE first
34. Punjabi? TRUE first
35. Yoruba? TRUE first
36. Swahili? TRUE first
37. Mongolian? TRUE first
38. Latin? TRUE first
39. Birthdate, oldest to youngest
40. Surname, A to Z
41. Forename, A to Z

## Notes

- The form should use `Italian`, not `Italy`.
- The form should use `Persian`; the code treats Persian as the Farsi/Persian priority category.
- If Google Forms creates duplicate columns because of conditional sections, the app uses the first non-empty matching value.
- The generated workbook preserves the original data and adds calculated ranking columns at the front.
