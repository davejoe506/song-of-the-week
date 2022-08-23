# Song of the Week (SOTW) Competition

At my previous employer, to add some levity between otherwise very serious work responsibilities, several colleagues and I started a song of the week competition. Every week, each of us would add a song to our collaborative Spotify playlist. On Fridays, one of us would send a survey asking each participant to vote on the three best songs submitted that week. After everyone would cast their ballots, the votes would be tallied and that week's champion would be crowned. It is a fun little tradition that started (or started being documented, at the very least) on July 27, 2017. While participants have come and gone and almost all of us have since departed the employer that brought us together, the tradition continues to this day.

At some point I became the administrator of the competition and in my early days as admin, I would manually count the votes and update the results spreadsheet. As I became more adept at using Python, I realized that this would be a perfect task to automate, so I created a script and have uploaded it here.

## Rules of the Game
1) Since Spotify is the repository for the playlist, participants can only submit songs that exist on Spotify.
2) Participants cannot submit a song that has previously been submitted to the Spotify playlist.
3) Participants vote for the best, second-best, and third-best song each week (irrespective of their personal projections of points tabulations and/or personal agendas). First place = 3 points, second place = 2 points, and third place = 1 point.
4) Participants cannot vote for their own submission (honor system applies).
5) In the event of a tie, a tiebreaker has been coded up in the Python script. The tiebreaker methodology is as follows: each participant who ties is assigned an equal probability out of 999 to win (with the administrator collecting 1/1,000 as an administrative tax). Then, a number is randomly drawn out of 1,000, and whoever's assigned probability corresponds with the drawn number is declared the winner.
6) The winner of each week's competition gets to submit a victory song to the playlist that all participants must listen to.

## Requirements to Run Code
The following is needed to run the script. I have not included them in the repository (unless otherwise stated) so as not to expose any potentially sensitive information, but current participants in the program should have enough information to recreate them in case they need to take over as SOTW admin.
- Spotify API credentials
- Google Sheets API credentials
- Google Forms survey
- Google Sheets results spreadsheet
- Excel results spreadsheet

## General Steps to Follow as SOTW Admin
1) Clone this repository to your computer (if not done already).
2) Create Google Forms survey for this week's SOTW competition.
3) After creating Google Forms survey, click within Forms to create Google Sheets spreadsheet for responses.
4) Make sure Python IDE file directory is pointing to folder with sotw.py.
5) Change date in filename of SOTW Excel results file to be date of competition.
6) Run # IMPORT MODULES # portion of sotw.py.
7) Update/run # CONSIDER UPDATING THESE VARIABLES PRIOR TO EACH SOTW # portion of sotw.py.
8) Run sotw.py through creation of sotw_sn and check if Google Forms & Spotify song names match. If not, reconcile in either Google Forms or script.
9) Continue running sotw.py until right before # CREATE TOTAL POINTS ROW TO APPEND TO ALL-TIME RESULTS EXCEL SHEET portion and make sure results make sense by looking at printed text in iPython console and also looking at sotw_points DataFrame.
10) Run rest of sotw.py.
11) Open updated Excel results file and sort "All-Time Wins", "All-Time Points", "Points Per Submission", & "Percentage of Available Total Points" tables in "All-Time Results" sheet.
12) Write SOTW results email and attach SOTW Excel results file.
13) Add winning song to "NERA SOTW CHAMPS" Spotify playlist.
14) Create duplicate of Google Forms survey for next week and close this week's survey.

## Extending This Code
Some ideas to extend this code:
- Incorporate Google Forms API (released since creation of script) to create weekly Google Forms survey within Python.
- Make script more responsive/dynamic where there is less than full set of participants (10) in competition.
- Figure out Python commands for sorting portions of loaded/imported Excel spreadsheet.
