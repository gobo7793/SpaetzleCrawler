﻿#region License
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
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using SimpleLogger;

namespace SpätzleCrawler
{
    class Program
    {
        public static async void Main(string[] args)
        {
            try
            {
                // read necessary data
                var t = new Task<List<string>>(ReadUserList);
                t.Start();
                Console.Write("URL der Tippabgabe: ");
                var url = Console.ReadLine();
                t.Wait();
                var users = t.Result;

                // getting tips


                // saving all

            }
            catch(Exception e)
            {
                Console.WriteLine(e);
                SimpleLog.Log(e);
            }

        }

        /// <summary>
        /// Reads the userlist
        /// </summary>
        /// <returns>The userlist</returns>
        public static List<string> ReadUserList()
        {
            FileHandler.ReadConfig(Settings.ConfigFileName);
            var excel = new ExcelHandler();
            excel.OpenFile(Settings.TargetFileName);
            return excel.ReadUserList();
        }
    }
}
