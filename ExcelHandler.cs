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
        /// Creates a new Instance for a new File and initialize it
        /// </summary>
        public ExcelHandler()
        {
            Init();
        }

        /// <summary>
        /// Finalizer, close excel and release the resources
        /// </summary>
        ~ExcelHandler()
        {
            if(CanUse)
            {
                ExcelApp.DisplayAlerts = false;
                ExcelApp.Quit();
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

        #endregion

        #region Main Methods

        /// <summary>
        /// Initialize a new file
        /// </summary>
        public void Init()
        {
            try
            {
                ExcelApp = new Application();
#if DEBUG
                ExcelApp.Visible = true;
#endif
                ExcelApp.DisplayAlerts = false;
                Workbook = ExcelApp.Workbooks.Add(1);
                Worksheet = (Worksheet)Workbook.Sheets[1];
            }
            catch(Exception e)
            {
                SimpleLog.Error("Excel export cannot be used. Is MS Excel installed on your machine?");
                SimpleLog.Error(e.ToString());
            }
        }

        /// <summary>
        /// Saves the file to the fiven file name and returns true, if file saved succesfully.
        /// </summary>
        /// <param name="fileName">File name for the file</param>
        /// <returns>True if file saved successfully</returns>
        public bool Close(string fileName)
        {
            if(!CanUse) return false;

            if(!fileName.EndsWith(".xlsx")) fileName += ".xlsx";
            FileInfo file = new FileInfo(fileName);
            bool retVal = false;

            try
            {
                SimpleLog.Info($"Save Excel file \"{file.FullName}\".");
                Workbook.SaveAs(Filename: file.FullName);
                ExcelApp.Quit();
                retVal = true;
                SimpleLog.Info($"Excel file \"{file.FullName}\" saved.");
            }
            catch(Exception e)
            {
                SimpleLog.Log($"Failed to write file \"{file.FullName}\".");
                SimpleLog.Error(e.ToString());
            }
            return retVal;
        }

        #endregion

        #region Write Data



        #endregion

    }
}
