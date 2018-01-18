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

namespace SpätzleCrawler
{
    /// <summary>
    /// Class to storing the settings
    /// </summary>
    public class Settings
    {
        /// <summary>
        /// Filename of the configuration file
        /// </summary>
        public static string ConfigFileName { get; set; } = "config.txt";

        /// <summary>
        /// File name of the target xlsx file
        /// </summary>
        public static string TargetFileName { get; set; }
    }
}