# SpaetzleCrawler

Gets the tips from all users of one league from the Spätzles Liga and saves it into a excel file. The excel file must be specific for the league and on the first worksheet must be for the tip entry.

Note: The template for the excel file is currently not open source.

## Requirements
- .NET Framework 4.7
- Microsoft Excel

## Run the crawler
Note: All row/column indizes based on the worksheet for tip entries.
1. Compile the source or download a binary
2. Create a textfile called `config.txt` in the same directory like the `SpätzleCrawler.exe` file
3. Insert the configurations for the leages described below
4. Enter the matches of the current Bundesliga matchday into columns `B` (home team) and `E` (away team)
    * All previously matchdays needed at least one result in column `C` (home goals)
    * The current matchday is the first matchday with no results in column `C`
5. Start the crawler (no worry, it uses a command line)
6. Enter the thread url of the current tipping thread
    * The userlist is stored in row `2` from column `L` to the first not merged cell in the row
7. The crawler tries to parse the tips:
    * Download all pages in the thread, mostly 5 oder 6
    * Getting post data (URL, username and content without quotes)
    * Parse the matches between the users if one post with all usernames is found
    * Parse the tips from all users
    * Note if parsing is not successfully the crawler will write the post url to the console and log file
8. The crawler save the matches between the users and their tips
    * The user for the matches will be stored in column `G` and `J`
    * The usertips will be stored from column `L` upwards, one match per row
9. See log file for troubleshooting

## Config file
Multiple league support was introduced in Version 2.0. With Version 1.0 only one leage can be crawled and config.txt stored only the file path to the excel file in the first line.

If you want to crawl one ore more leagues in V2.0, you need for each leage an own excel file. Comments to the end of the line can be started with `#`.

Each league is described as own line with league name and the config file in the following format:
```
League name = C:\Path\To\Excel\File.xlsx
```

Sample config.txt:
```
# Sample Crawler Config
Liga 1=%USERPROFILE%\Dokumente\liga1.xlsx
Liga 2=C:\Users\User\Dokumente\liga2.xlsx
```
