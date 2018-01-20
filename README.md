# SpaetzleCrawler

Gets the tips from all users of one league from the Spätzles Liga and saves it into a excel file. The excel file must be specific for the league and on the first worksheet must be for the tip entry. Note: The template for the excel file is currently not open source.

## Run the crawler
Note: All row/column indizes based on the worksheet for tip entries.
1. Compile with Roslyn compiler/.NET 4.7
2. Create a textfile called `config.txt` in the directory of the exe file (mostly `bin\Debug` or `bin\Release`)
3. Insert the full file path of the excel file in the first line in `config.txt`, like `C:\Users\<User>\Documents\league.xlsx`
4. Enter the matches of the current Bundesliga matchday into columns `B` (home team) and `E` (away team)
    * All previously matchdays needed at least one result in column `C` (home goals)
    * The current matchday is the first matchday with no results in column `C`
5. Start the crawler
6. Enter the thread url of the current tipping thread
    * The userlist is stored in row `2` from column `L` to the first not merged cell in the row
7. The crawler tries to parse the tips:
    * Download all pages in the thread, mostly 5 oder 6
    * Getting post data (URL, username and content without quotes)
    * Parse the matches between the users
    * Parse the tips from all users
    * Note if parsing is not successfully the crawler will write the post url to the console and log file
8. The crawler save the matches between the users and their tips
    * The user for the matches will be stored in column `G` and `J`
    * The usertips will be stored from column `L` upwards, one match per row
9. See log file for troubleshooting
