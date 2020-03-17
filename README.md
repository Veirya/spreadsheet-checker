Spreadsheet Checker, by Giovanni "Veirya" Oliver (2020)
  :::::::::::::::::::::::::::::::::::::::::::::::::
  Simple script that looks at a Google Sheets sheet of FFXIV FC member data,
    then compares it to information from the Lodestone. Information considered
    to be of interest is then printed out for the user to see. Current iteration
    only works for the FC "FullMetal Alliance" and a hard-coded sheet.
  Information considered of interest by the script:
    - New/Old Members
    - Name Changes
    - Rank Changes
    - Possibile Promotion Candidates (FMA New Recruits only)
    - Lack of Date for a Member in Spreadsheet.

The code used to read data from the Google Sheets API was taken from Google's provided quickstart example
code, and adjusted to fit my use. The code to start the Google API client and verify it was entirely
taken from the same file.
