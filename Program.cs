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
using System.Threading.Tasks;
using SimpleLogger;

namespace SpätzleCrawler
{
    class Program
    {
        public static void Main(string[] args)
        {
            SimpleLog.SetLogFile("logs", writeText: true);

            try
            {
                // read necessary data
                bool isConfigReaded = FileHandler.ReadConfig(Settings.ConfigFileName);
                if(!isConfigReaded)
                    throw new InvalidOperationException("Error while reading config file.");

                var readUserTask = Task.Run(() => ExcelHandler.Handler.ReadUserList());
                Console.Write("URL current thread: ");
                Settings.TipThreadUrl = Console.ReadLine();
                readUserTask.Wait();
                var users = readUserTask.Result;
                var matches = ExcelHandler.Handler.GetNextMatchdayMatches();

                SimpleLog.Info($"{users.Count} Users found. Use {Settings.TipThreadUrl} to get tips.");

                // getting tips
                var posts = Crawler.GetPosts();
                SimpleLog.Info($"{posts.Count} posts found.");
                var usermatches = Parser.ParsePosts(matches, users, posts);
                SimpleLog.Info("All data parsed.");

                // saving all
                if(ExcelHandler.Handler.WriteUsermatches(usermatches))
                    Console.WriteLine("Matches between users writed.");
                else
                    Console.WriteLine("No matches between users writed.");

                if(ExcelHandler.Handler.WriteUsertips(users))
                    Console.WriteLine("Tips from users writed.");
                else
                    Console.WriteLine("No tips from users writed.");
                ExcelHandler.Handler.Close();

                Console.WriteLine("Finished! Press any key to exit.");
            }
            catch(Exception e)
            {
                Console.WriteLine(e);
                SimpleLog.Log(e);
            }

            Console.ReadKey();
        }
    }
}
