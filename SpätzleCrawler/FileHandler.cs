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
using System.Text.RegularExpressions;
using SimpleLogger;

namespace SpätzleCrawler
{
    /// <summary>
    /// File handler class for reading the config file
    /// </summary>
    public class FileHandler
    {
        /// <summary>
        /// Reads the config file and saves the league settings
        /// </summary>
        /// <param name="filename">The config filename</param>
        /// <param name="settings">The league settings</param>
        /// <returns>True if settings loaded</returns>
        /// <remarks>
        /// Line Format: "League Name=C:\Path\to\file.xlsx"
        /// Comments starting with #
        /// </remarks>
        public static bool ReadConfig(string filename, out List<Settings> settings)
        {
            settings = new List<Settings>();
            var settingRegex = new Regex(@"([^=]*)\s*=\s*(.*)");
            try
            {
                using(var sr = new StreamReader(filename))
                {
                    while(!sr.EndOfStream)
                    {
                        var line = sr.ReadLine();
                        if(String.IsNullOrWhiteSpace(line))
                            continue;

                        // Comment handling
                        var relevantLineContent = line;
                        if(line.Contains("#"))
                        {
                            var lineSplitted = line.Split('#');
                            if(String.IsNullOrWhiteSpace(lineSplitted[0]))
                                continue;

                            relevantLineContent = lineSplitted[0];
                        }

                        // read settings
                        var lineMatch = settingRegex.Match(relevantLineContent);
                        if(!lineMatch.Success)
                            throw new InvalidDataException("Config file format cannot be readed.");

                        var targetFile = Environment.ExpandEnvironmentVariables(lineMatch.Groups[2].Value.Trim());
                        var fi = new FileInfo(targetFile);
                        if(!fi.Exists)
                            throw new InvalidDataException($"Target xlsx file \"{targetFile}\" not found.");

                        var setting = new Settings
                        {
                            LeagueName = lineMatch.Groups[1].Value.Trim(),
                            TargetFileName = fi.FullName,
                        };
                        settings.Add(setting);
                    }
                }
                return true;
            }
            catch(Exception e)
            {
                SimpleLog.Error("Cannot read config file. Wrong filename or format?");
                SimpleLog.Error(e.ToString());
            }
            return false;
        }
    }
}