# Harvard Citizenship Tutor Ranker and Matcher

A static GitHub Pages website for the Harvard IOP Citizenship Tutoring Program. It ranks tutors from a Google Form Excel export and then generates optimized tutor-tutee matches from an internal tutee Excel export.

## What it does

1. Upload the tutor membership spreadsheet downloaded from Google Forms.
2. Rank tutors by program seniority.
3. Preview the ranked tutor table in the browser.
4. Download a ranked tutor Excel file.
5. Upload the internal tutee spreadsheet.
6. Generate optimized matches using ranked tutors.
7. Preview the match table in the browser.
8. Download a final Excel workbook with multiple sheets:
   - `Generated Matches`
   - `Manual Review`
   - `Tutee Priority Rankings`
   - `Tutor Seniority Rankings`
   - `Tutor Workload`

## Privacy

The site is browser-only. Spreadsheet contents are processed locally in the user's browser and are not uploaded to a server.

## Hosting with GitHub Pages

1. Create a new GitHub repository.
2. Upload these files to the repository root:
   - `index.html`
   - `style.css`
   - `script.js`
   - `README.md`
3. Open the repository's **Settings**.
4. Go to **Pages**.
5. Under **Build and deployment**, choose **Deploy from a branch**.
6. Choose the `main` branch and `/root` folder.
7. Save.
8. GitHub will provide a public site link after deployment.

## Tutor seniority rules

The tutor ranking follows this order:

1. Director
2. Secretary
3. Current board member
4. Former director
5. Former secretary
6. Former board member
7. Years since director, largest to smallest
8. Years since secretary, largest to smallest
9. Years since board member, largest to smallest
10. Returning tutor
11. Class year, smallest to largest
12. Number of LOTEs, largest to smallest
13. Language priority order:
    - Spanish
    - Mandarin
    - Haitian Creole
    - French
    - Cantonese
    - Vietnamese
    - Arabic
    - Russian
    - Italian
    - Hindi
    - Cape Verdean Creole
    - Korean
    - German
    - Telugu
    - Amharic
    - Tagalog
    - Japanese
    - Urdu
    - Albanian
    - Persian
    - Thai
    - Punjabi
    - Yoruba
    - Swahili
    - Mongolian
    - Latin
14. Birthdate, oldest to youngest
15. Surname A-Z
16. Forename A-Z

Note: Persian is used as the form label and seniority language label.

## Matching rules

The matching system implements the following hierarchy:

1. Returning tutor continuity is processed first. If a tutor says they are returning to work with a specific returning tutee, they are an instant match provided that they share at least one availability slot.
2. Availability is a hard rule for every tutor assignment.
3. Every tutee session must include at least one tutor who speaks the tutee's native language. If this is not possible, the session is flagged for manual review.
4. New tutees are prioritized before returning tutees after continuity matches.
5. Naturalization stage determines the base number of tutors needed per session.
6. Returning tutees receive half as many tutors per session as new tutees, rounded up.
7. The requested number of weekly meetings creates separate weekly session rows.
8. In-person or Zoom preference is optimized after hard rules.
9. Tutor workload is capped by the configurable maximum sessions per tutor per week.
10. All unresolved cases are placed in the `Manual Review` sheet.

## Configurable controls

Before generating matches, the director can change:

- Maximum sessions per tutor per week
- Default tutors/session for N-400 stage
- Default tutors/session for interview stage
- Default tutors/session for intensive stage

## Spreadsheet expectations

The code searches for headers flexibly rather than by column position. This helps with Google Forms exports and duplicated branch columns.

Important tutor fields include:

- Forename
- Surname
- Email or Harvard Email
- Class Year
- Language(s) Other than English
- Current Board Member
- Current board position
- Former Board Member
- Former highest position
- Years ago selected for position
- Returning Member
- Returning tutee name, when applicable
- Availability
- In-person/Zoom preference
- Birthdate

Important tutee fields include:

- Forename
- Surname
- Email
- Phone Number
- Associated Community Partner
- New or Returning
- Native Language
- Tutee availability by day
- Requested meetings per week
- Naturalization stage
- In-person/Zoom preference

## Notes for future maintainers

The most important logic is in `script.js`:

- `prepareTutor()` normalizes tutor spreadsheet rows.
- `compareTutors()` applies the seniority ranking.
- `prepareTutee()` normalizes tutee spreadsheet rows.
- `runMatching()` generates the final matches.
- `bestCandidates()` and `scoreCandidate()` choose tutors for each tutee-session.

The app uses SheetJS from a CDN to read and write Excel files.
