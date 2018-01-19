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
using System.IO;
using System.Linq;
using System.Net.Mime;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
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

        public const int FirstUserCol = 12;
        public const int MaxUserCol = 71;
        public const int UserColCount = 3;

        #endregion

        #region Constructor

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
                ExcelApp?.DisplayAlerts = false;
                ExcelApp?.Quit();
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
#if DEBUG
                    Visible = true,
#else
                    DisplayAlerts = false
#endif
                };
                Workbook = ExcelApp.Workbooks.Open(filename);
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
        /// <param name="fileName">File name for the file</param>
        /// <returns>True if file saved successfully</returns>
        public bool Close(string fileName)
        {
            if(!CanUse) return false;

            try
            {
                SimpleLog.Info($"Save Excel file \"{Workbook.FullName}\".");
                Workbook.Save();
                ExcelApp.Quit();
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

        #region Read/Write Data

        /// <summary>
        /// Reads and returns the participating users list
        /// </summary>
        /// <returns>The participating usernames</returns>
        public List<string> ReadUserList()
        {
            int row = 2;
            int col = FirstUserCol;

            var users = new List<string>();
            while(col < MaxUserCol)
            {
                var cell = (Range)Worksheet.Cells[row, col];
                if(!cell.MergeCells)
                    break;
                var cellVal = (string)cell.Value;
                if(!String.IsNullOrWhiteSpace(cellVal))
                    users.Add(cellVal);

                SimpleLog.Info($"Read Excel cell [{row},{col}]: {cellVal}");

                col += UserColCount;
            }

            return users;
        }

        /// <summary>
        /// Reads an returns the matches of the next matchday
        /// </summary>
        /// <returns></returns>
        public List<(string, string)> GetNextMatchdayMatches()
        {
            var nextNo = GetNextMatchdayNo();

            int team1Col = 2;
            int team2Col = 5;

            // search current matchday row
            int matchdayRow;
            for(matchdayRow = FirstMatchdayRow; matchdayRow < MaxMatchdayRow; matchdayRow += MatchdayRowCount)
            {
                var resValue = (double?)((Range)Worksheet.Cells[matchdayRow, MatchdayCol]).Value;
                if(resValue.HasValue && Math.Abs(resValue.Value - nextNo) < 0.1)
                {
                    break;
                }
            }

            // Get teams
            var matches = new List<(string, string)>();
            for(int i = matchdayRow + 1; i <= matchdayRow + 9; i++)
            {
                var team1 = (string)((Range)Worksheet.Cells[i, team1Col]).Value;
                var team2 = (string)((Range)Worksheet.Cells[i, team2Col]).Value;
                matches.Add((team1, team2));
            }

            return matches;
        }

        /// <summary>
        /// Reads and returns the number of the next not yet started matchday.
        /// Returns -1 if no matchday found.
        /// </summary>
        /// <returns>The next matchday number</returns>
        public int GetNextMatchdayNo()
        {
            int resCol = 3;

            int weekRow = FirstMatchdayRow;

            while(weekRow < MaxMatchdayRow)
            {

                bool isEmpty = false;
                for(int i = weekRow + 1; i <= weekRow + 9; i++)
                {
                    isEmpty = false;
                    var resValue = (double?)((Range)Worksheet.Cells[i, resCol]).Value;
                    isEmpty = !resValue.HasValue;

                    if(!isEmpty)
                        break; // break if result cells are not empty
                }

                // if result cells are all empty, return matchday number
                if(isEmpty)
                {
                    var weekValue = (double)((Range)Worksheet.Cells[weekRow, MatchdayCol]).Value;
                    return Convert.ToInt32(weekValue);
                }

                // check next matchday if result cells are not empty
                weekRow += MatchdayRowCount;
            }
            return -1;
        }

        #endregion

    }
}
