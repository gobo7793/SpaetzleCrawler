#region License
// SpätzleCrawler
// Copyright (C) 2018 Gerald Siegert
//  
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//  
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//  
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.
#endregion

using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using SimpleLogger;

namespace SpätzleCrawler
{
    /// <summary>
    /// Handler for Excel Files
    /// Works only if MS Excel is installed on the local machine.
    /// </summary>
    public class ExcelHandler
    {
        #region Constants

        public const int MatchdayCol = 2;
        public const int FirstMatchdayRow = 2;
        public const int MaxMatchdayRow = 205;
        public const int MatchdayRowCount = 12;

        public const int MatchdayTeam1Col = 2;
        public const int MatchdayTeam1ResultCol = 3;
        public const int MatchdayTeam2ResultCol = 4;
        public const int MatchdayTeam2Col = 5;

        public const int FirstUserCol = 12;
        public const int MaxUserCol = 71;
        public const int UserColCount = 3;
        public const int UserListRow = 2;

        public const int UsermatchUser1Col = 7;
        public const int UsermatchUser2Col = 10;

        public const int RealMatchesPerMatchday = 9;

        #endregion

        #region Constructor

        /// <summary>
        /// Creates a new excel handler and opens <see cref="Settings.TargetFileName"/>
        /// </summary>
        private ExcelHandler()
        {
            OpenFile(Settings.TargetFileName);
        }

        /// <summary>
        /// Finalizer, close excel and release the resources
        /// </summary>
        ~ExcelHandler()
        {
            if(CouldUse)
            {
#if !DEBUG
                //if(ExcelApp != null)
                //{
                //    ExcelApp.DisplayAlerts = false;
                //    ExcelApp.Quit();
                //}
#endif
            }

            ExcelApp = null;
            Workbook = null;
            Worksheet = null;
        }

        #endregion

        #region Properties

        private static ExcelHandler _Handler;

        /// <summary>
        /// The Excel Handler with the open <see cref="Settings.TargetFileName"/>
        /// </summary>
        public static ExcelHandler Handler => _Handler ?? (_Handler = new ExcelHandler());

        /// <summary>
        /// The current Excel Application
        /// </summary>
        public Application ExcelApp { get; set; }

        /// <summary>
        /// The current Excel Workbook
        /// </summary>
        public Workbook Workbook { get; set; }

        /// <summary>
        /// The current Excel Worksheet
        /// </summary>
        public Worksheet Worksheet { get; set; }

        /// <summary>
        /// Indicates if Excel can be used
        /// </summary>
        public bool CanUse => ExcelApp != null && Workbook != null && Worksheet != null;

        /// <summary>
        /// Indicates if Excel propably could be used
        /// </summary>
        public bool CouldUse => ExcelApp != null || Workbook != null || Worksheet != null;

        /// <summary>
        /// Row number for the next matchday
        /// </summary>
        public int NextMatchdayRow { get; set; }

        #endregion

        #region Main Methods

        /// <summary>
        /// Open the given file
        /// </summary>
        public bool OpenFile(string filename)
        {
            try
            {
                SimpleLog.Info($"Open Excel file \"{filename}\".");
                ExcelApp = new Application
                {
                    Visible = true,
                    DisplayAlerts = false,
                };
                Workbook = ExcelApp.Workbooks.Open(filename, Editable: true);
                Worksheet = (Worksheet)Workbook.Sheets[1];

                return true;
            }
            catch(Exception e)
            {
                SimpleLog.Error("Excel export cannot be used. Is MS Excel installed on your machine?");
                SimpleLog.Error(e.ToString());
            }
            return false;
        }

        /// <summary>
        /// Saves the file to the fiven file name and returns true, if file saved succesfully.
        /// </summary>
        /// <returns>True if file saved successfully</returns>
        public bool Close()
        {
            if(!CanUse) return false;

            try
            {
                SimpleLog.Info($"Save Excel file \"{Workbook.FullName}\".");
                Workbook.Save();
                //ExcelApp.Quit();
                SimpleLog.Info($"Excel file \"{Workbook.FullName}\" saved.");
                return true;
            }
            catch(Exception e)
            {
                SimpleLog.Log($"Failed to write file \"{Workbook.FullName}\".");
                SimpleLog.Error(e.ToString());
            }
            return false;
        }

        #endregion

        #region Read Data

        /// <summary>
        /// Reads and returns the participating users list
        /// </summary>
        /// <returns>The participating usernames</returns>
        public List<User> ReadUserList()
        {
            SimpleLog.Info("Reading userlist...");

            int col = FirstUserCol;

            var users = new List<User>();
            while(col < MaxUserCol)
            {
                var cell = (Range)Worksheet.Cells[UserListRow, col];
                if(!cell.MergeCells)
                    break;
                var cellVal = (string)cell.Value;
                if(!String.IsNullOrWhiteSpace(cellVal))
                {
                    var user = new User
                    {
                        Name = cellVal,
                        UserCol = col,
                    };
                    users.Add(user);
                }

                SimpleLog.Info($"Read Excel cell [{UserListRow},{col}]: {cellVal}");

                col += UserColCount;
            }

            return users;
        }

        /// <summary>
        /// Reads an returns the matches of the next matchday
        /// </summary>
        /// <returns></returns>
        public List<FootballMatch> GetNextMatchdayMatches()
        {
            SimpleLog.Info("Reading next matchday matches...");

            var nextNo = ReadMatchdayNo();

            // search current matchday row
            int matchdayRow;
            for(matchdayRow = FirstMatchdayRow; matchdayRow < MaxMatchdayRow; matchdayRow += MatchdayRowCount)
            {
                var resValue = (double?)((Range)Worksheet.Cells[matchdayRow, MatchdayCol]).Value;
                SimpleLog.Info($"Read Excel cell [{matchdayRow},{MatchdayCol}]: {resValue}");
                if(resValue.HasValue && Math.Abs(resValue.Value - nextNo) < 0.1)
                {
                    break;
                }
            }

            // Get teams
            var matches = new List<FootballMatch>();
            for(int i = matchdayRow + 1; i <= matchdayRow + RealMatchesPerMatchday; i++)
            {
                var team1 = (string)((Range)Worksheet.Cells[i, MatchdayTeam1Col]).Value;
                var team2 = (string)((Range)Worksheet.Cells[i, MatchdayTeam2Col]).Value;
                SimpleLog.Info($"Read Excel cell [{i},{MatchdayTeam1Col}]: {team1}");
                SimpleLog.Info($"Read Excel cell [{i},{MatchdayTeam2Col}]: {team2}");
                matches.Add(new FootballMatch { TeamA = team1, TeamB = team2 });
            }

            return matches;
        }

        /// <summary>
        /// Reads and returns the number of the next not yet started matchday.
        /// Returns -1 if no matchday found.
        /// Also saves the matchday on <see cref="Settings.NextMatchdayNo"/>.
        /// </summary>
        /// <returns>The next matchday number</returns>
        public int ReadMatchdayNo()
        {
            SimpleLog.Info("Reading next matchday number...");
            var matchdayNo = -1;
            NextMatchdayRow = FirstMatchdayRow;

            int weekRow = FirstMatchdayRow;

            while(weekRow < MaxMatchdayRow)
            {

                bool isEmpty = false;
                for(int i = weekRow + 1; i <= weekRow + RealMatchesPerMatchday; i++)
                {
                    var resValue = (double?)((Range)Worksheet.Cells[i, MatchdayTeam1ResultCol]).Value;
                    SimpleLog.Info($"Read Excel cell [{i},{MatchdayTeam1ResultCol}]: {resValue}");
                    isEmpty = !resValue.HasValue;

                    if(!isEmpty)
                        break; // break if result cells are not empty
                }

                // if result cells are all empty, return matchday number
                if(isEmpty)
                {
                    var weekValue = (double)((Range)Worksheet.Cells[weekRow, MatchdayCol]).Value;
                    SimpleLog.Info($"Read Excel cell [{weekRow},{MatchdayCol}]: {weekValue}");
                    matchdayNo = Convert.ToInt32(weekValue);
                    NextMatchdayRow = weekRow;
                    break;
                }

                // check next matchday if result cells are not empty
                weekRow += MatchdayRowCount;
            }

            Settings.NextMatchdayNo = matchdayNo;
            return matchdayNo;
        }

        #endregion

        #region Write Data

        /// <summary>
        /// Writes the usermatches to the excel file and returns true on matches to write
        /// </summary>
        /// <param name="usermatches">The usermatches</param>
        /// <returns>True on matches to write</returns>
        public bool WriteUsermatches(List<Usermatch> usermatches)
        {
            SimpleLog.Info($"Write {usermatches.Count} usermatches...");

            for(int i = 0; i < usermatches.Count; i++)
            {
                var cellRow = NextMatchdayRow + 1 + i;
                Worksheet.Cells[cellRow, UsermatchUser1Col] = usermatches[i].UserA.Name;
                Worksheet.Cells[cellRow, UsermatchUser2Col] = usermatches[i].UserB.Name;
                SimpleLog.Info($"Write Excel cell [{cellRow},{UsermatchUser1Col}]: {usermatches[i].UserA.Name}");
                SimpleLog.Info($"Write Excel cell [{cellRow},{UsermatchUser2Col}]: {usermatches[i].UserB.Name}");
            }

            return usermatches.Any();
        }

        /// <summary>
        /// Writes the tips of the user to the excel file and returns true on tips to write
        /// </summary>
        /// <param name="userlist">The userlist</param>
        /// <returns>True on tips to write</returns>
        public bool WriteUsertips(List<User> userlist)
        {
            SimpleLog.Info($"Write {userlist.SelectMany(x => x.Tips).Count()} usertips...");

            for(int row = NextMatchdayRow + 1; row <= NextMatchdayRow + RealMatchesPerMatchday; row++)
            {
                // get match in row
                var team1Name = (string)((Range)Worksheet.Cells[row, MatchdayTeam1Col]).Value;
                if(String.IsNullOrWhiteSpace(team1Name))
                    continue;
                foreach(var user in userlist)
                {
                    var match = user.Tips.FirstOrDefault(m => m.TeamA == team1Name);
                    if(match == null)
                        continue;

                    // write tip
                    Worksheet.Cells[row, user.UserCol] = match.ResultA;
                    Worksheet.Cells[row, user.UserCol + 1] = match.ResultB;
                    SimpleLog.Info($"Write Excel cell [{row},{user.UserCol}]: {match.ResultA}");
                    SimpleLog.Info($"Write Excel cell [{row},{user.UserCol + 1}]: {match.ResultB}");
                }
            }

            return userlist.SelectMany(x => x.Tips).Any();
        }

        #endregion

    }
}
