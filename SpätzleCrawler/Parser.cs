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
using System.Text.RegularExpressions;
using SimpleLogger;

namespace SpätzleCrawler
{
    /// <summary>
    /// Parser to getting the results of the tips
    /// </summary>
    public class Parser
    {

        /// <summary>
        /// The real matches
        /// </summary>
        public List<FootballMatch> Matches { get; set; }

        /// <summary>
        /// The Userlist
        /// </summary>
        public List<User> Userlist { get; set; }

        /// <summary>
        /// The usermatches
        /// </summary>
        public List<Usermatch> Usermatches { get; } = new List<Usermatch>();

        /// <summary>
        /// The thread posts
        /// </summary>
        public List<(string Url, string Username, string Content)> Posts { get; set; }

        /// <summary>
        /// Helper regex for line splitting
        /// </summary>
        private Regex _LineSplitterRegex => new Regex(@"\r\n|\r|\n");

        /// <summary>
        /// Parses the match tips
        /// </summary>
        public void ParseMatchTips()
        {
            SimpleLog.Info("Parse match tips from user...");

            var matchTipRegex = new Regex(@"(\d+)(?::|-)(\d+)");
            var nicks = Userlist.Select(u => u.Name).ToArray();
            foreach(var post in Posts)
            {
                if(!nicks.Contains(post.Username))
                    continue;

                SimpleLog.Info($"User {post.Username} writed post {post.Url}.");
                var user = Userlist.First(u => u.Name == post.Username);
                if(user.Tips.Count == User.TippingGamesCount)
                {
                    SimpleLog.Warning($"User {post.Username} has already {User.TippingGamesCount} tips!");
                    Console.WriteLine($"User {post.Username} has already {User.TippingGamesCount} tips. Skipping reading, please check the post {post.Url}.");
                    continue;
                }

                // try parse tips
                user.Tips.Clear();
                var lines = _LineSplitterRegex.Split(post.Content);
                foreach(var line in lines)
                {
                    var realMatch = Matches.FirstOrDefault(m => line.ToLower().Contains(m.TeamA.ToLower()) && line.ToLower().Contains(m.TeamB.ToLower()));
                    var tipRegexMatches = matchTipRegex.Matches(line);
                    if(tipRegexMatches.Count != 2 || realMatch == null)
                        continue;

                    // save tips
                    var footballMatch = new FootballMatch
                    {
                        TeamA = realMatch.TeamA,
                        TeamB = realMatch.TeamB,
                        ResultA = Int32.Parse(tipRegexMatches[1].Groups[1].Value),
                        ResultB = Int32.Parse(tipRegexMatches[1].Groups[2].Value),
                    };
                    user.Tips.Add(footballMatch);
                }

                if(user.Tips.Count != User.TippingGamesCount)
                {
                    SimpleLog.Warning("Tips could not be readed!");
                    Console.WriteLine($"Tips could not be readed. User {post.Username} in post {post.Url}. Do it yourself.");

                    user.Tips.Clear();
                    break;
                }

                SimpleLog.Info("Tips readed!");
                Console.WriteLine($"Tips from user {post.Username} in post {post.Url} readed.");
            }

            // Output not tipped
            foreach(var user in Userlist)
            {
                if(user.Tips.Any())
                    continue;

                SimpleLog.Info($"User {user.Name} obviously not tipped.");
                Console.WriteLine($"User {user.Name} obviously not tipped.");
            }
        }

        /// <summary>
        /// Fills <see cref="Usermatches"/> and returns true on success
        /// </summary>
        /// <returns>True if matches succesfully filled</returns>
        public bool GetUserMatches()
        {
            SimpleLog.Info("Parse usermatches...");

            // search post
            string content = String.Empty;
            string url = String.Empty;
            foreach(var post in Posts)
            {
                var postText = post.Content.ToLower();
                bool matches = Userlist.Select(u => u.Name).All(u => postText.Contains(u.ToLower()));

                if(matches)
                {
                    content = post.Content;
                    url = post.Url;
                    SimpleLog.Info($"Post with usermatches: {url}");
                    break;
                }
            }

            // escape if not found
            var userMatchRegex = new Regex(@"^([^\s]+)\s*(?:\((?:A|N)\))?\s*:\s([^\s]+)\s*(?:\((?:A|N)\))?$", RegexOptions.Multiline);
            var regexMatches = userMatchRegex.Matches(content);
            if(string.IsNullOrWhiteSpace(content) || regexMatches.Count == 0)
            {
                SimpleLog.Warning("Could not find post with usermatches!");
                Console.WriteLine("Could not find post with usermatches. Please insert it manually.");
                return false;
            }

            // parse
            Usermatches.Clear();
            for(int i = 0; i < regexMatches.Count; i++)
            {
                var user1 = Userlist.FirstOrDefault(u => u.Name == regexMatches[i].Groups[1].Value);
                var user2 = Userlist.FirstOrDefault(u => u.Name == regexMatches[i].Groups[2].Value);

                if(user1 != null && user2 != null)
                    Usermatches.Add(new Usermatch { UserA = user1, UserB = user2 });
                else if(Usermatches.Count == Userlist.Count / 2)
                    return true;
                else
                {
                    SimpleLog.Warning("Could not parse usermatches!");
                    Console.WriteLine($"Could not parse usermatches from {url}. Do it yourself.");
                    Usermatches.Clear();
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Parses the match tips from the user posts and returns the parsed usermatches
        /// </summary>
        /// <param name="matches">The matches to parse</param>
        /// <param name="userlist">The userlist</param>
        /// <param name="posts">The postlist</param>
        /// <returns></returns>
        public static List<Usermatch> ParsePosts(List<FootballMatch> matches, List<User> userlist, List<(string Url, string Username, string Content)> posts)
        {
            var parser = new Parser
            {
                Matches = matches,
                Userlist = userlist,
                Posts = posts,
            };

            parser.GetUserMatches();
            parser.ParseMatchTips();

            return parser.Usermatches;
        }
    }
}