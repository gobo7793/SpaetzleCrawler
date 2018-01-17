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

        #region Constructor

        /// <summary>
        /// Finalizer, close excel and release the resources
        /// </summary>
        ~ExcelHandler()
        {
            if(CouldUse)
            {
#if !DEBUG
                ExcelApp.DisplayAlerts = false;
                ExcelApp.Quit();
#endif
            }

            ExcelApp = null;
            Workbook = null;
            Worksheet = null;
        }

        #endregion

        #region Properties

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
            int col = 'L' - 'A';

            var users = new List<string>();
            while(col < 71)
            {
                var cell = (Range)Worksheet.Cells[row, col];
                if(!cell.MergeCells)
                    break;
                var cellVal = (string)cell.Value;
                users.Add(cellVal);

                SimpleLog.Info($"Read Excel cell [{row},{col}]: {cellVal}");

                col += 3;
            }

            return users;
        }

        /// <summary>
        /// Reads and returns the number of the next not yet started matchday.
        /// Returns -1 if no matchday found.
        /// </summary>
        /// <returns>The next matchday number</returns>
        public int GetNextMatchday()
        {
            int resCol = 3;

            int weekRow = 2;
            int weekCol = 2;

            var matchdayRegex = new Regex(@"\d{1,2}");
            while(weekRow < 205)
            {

                bool isEmpty = false;
                for(int i = weekRow + 1; i <= weekRow + 9; i++)
                {
                    var resValue = (string)((Range)Worksheet.Cells[i, resCol]).Value;
                    isEmpty = String.IsNullOrWhiteSpace(resValue);

                    if(!isEmpty)
                        break; // break if result cells are not empty
                }

                // if result cells are all empty, return matchday number
                if(isEmpty)
                {
                    var weekValue = (string)((Range)Worksheet.Cells[weekRow, weekCol]).Value;
                    var day = matchdayRegex.Match(weekValue).Value;
                    return Int32.Parse(day);
                }

                // check next matchday if result cells are not empty
                weekRow += 12;
            }
            return -1;
        }

        #endregion

    }
}
