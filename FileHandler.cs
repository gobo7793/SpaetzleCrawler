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
using System.IO;
using SimpleLogger;

namespace SpätzleCrawler
{
    /// <summary>
    /// File handler class for reading the config file
    /// </summary>
    public class FileHandler
    {
        /// <summary>
        /// Reads the config file and reads the target xlsx filename in the first line of the file
        /// </summary>
        /// <param name="filename">The config filename</param>
        /// <returns>The target xlsx filename</returns>
        public static string ReadConfig(string filename)
        {
            try
            {
                string targetFile;
                using(var sr = new StreamReader(filename))
                {
                    targetFile = sr.ReadLine();
                    if(String.IsNullOrWhiteSpace(targetFile) || File.Exists(targetFile))
                        throw new InvalidDataException("Target xlsx filename could not be readed or file not found.");
                }
                return targetFile;
            }
            catch(Exception e)
            {
                SimpleLog.Error("Cannot read config file. Wrong filename or format?");
                SimpleLog.Error(e.ToString());
            }
            return null;
        }
    }
}