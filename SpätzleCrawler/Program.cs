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
using System.Threading.Tasks;
using SimpleLogger;

namespace SpätzleCrawler
{
    class Program
    {
        public static void Main(string[] args)
        {
            SimpleLog.SetLogFile("logs", writeText: true);

            // read necessary data
            var isConfigReaded = FileHandler.ReadConfig(Settings.ConfigFileName, out List<Settings> settings);
            if(!isConfigReaded)
            {
                var msg = "Error while reading config file. See log file for details.";
                Console.WriteLine(msg);
                SimpleLog.Log(msg);
            }
            else
            {
                foreach(var setting in settings)
                {
                    try
                    {
                        ParseLeague(setting);
                    }
                    catch(Exception e)
                    {
                        Console.WriteLine(e);
                        SimpleLog.Log(e);
                    }
                }
            }

            Console.WriteLine("Finished! Press l to view logs, c to close all excel instances or any other key to exit.");
            var key = Console.ReadKey();
            if(key.KeyChar == 'l' || key.KeyChar == 'L')
                SimpleLog.ShowLogFile();
            if(key.KeyChar == 'c' || key.KeyChar == 'C')
                foreach(var s in settings)
                    s.Excel.CloseFile();
            SimpleLog.StopLogging();
        }

        public static void ParseLeague(Settings settings)
        {
            Console.WriteLine($"Parsing {settings.LeagueName}.");
            SimpleLog.Info($"Parsing {settings.LeagueName}.");

            settings.Excel = new ExcelHandler();
            var readUserTask = Task.Run(() =>
            {
                settings.Excel.OpenFile(settings.TargetFileName);
                return settings.Excel.ReadUserList();
            });
            Console.Write("URL current matchday thread: ");
            settings.TipThreadUrl = Console.ReadLine();
            readUserTask.Wait();
            var users = readUserTask.Result;
            if(String.IsNullOrWhiteSpace(settings.TipThreadUrl))
            {
                settings.Excel.CloseFile();
                return; // cancel execution
            }
            var matches = settings.Excel.GetNextMatchdayMatches();

            SimpleLog.Info($"{users.Count} Users found. Use {settings.TipThreadUrl} to get tips.");

            // getting tips
            var posts = Crawler.GetPosts(settings);
            SimpleLog.Info($"{posts.Count} posts found.");
            var usermatches = Parser.ParsePosts(matches, users, posts);
            SimpleLog.Info("All data parsed.");

            // saving all
            if(settings.Excel.WriteUsermatches(usermatches))
                Console.WriteLine("Matches between users writed.");
            else
                Console.WriteLine("No matches between users writed.");

            if(settings.Excel.WriteUsertips(users))
                Console.WriteLine("Tips from users writed.");
            else
                Console.WriteLine("No tips from users writed.");
            settings.Excel.SaveFile();

            Console.WriteLine();
        }
    }
}
