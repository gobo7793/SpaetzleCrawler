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
        public List<(string, string)> Matches { get; set; }

        /// <summary>
        /// The Userlist
        /// </summary>
        public List<User> Userlist { get; set; }

        /// <summary>
        /// The usermatches
        /// </summary>
        public List<(User, User)> Usermatches { get; } = new List<(User, User)>();

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

                // try parse tips
                var lines = _LineSplitterRegex.Split(post.Content);
                bool mustBreak = false;
                foreach(var fullLine in lines)
                {
                    var secLine = fullLine.Split(new[] {"Uhr"}, StringSplitOptions.None)[1];
                    var realMatch = Matches.FirstOrDefault(m => secLine.Contains(m.Item1) && secLine.Contains(m.Item2));
                    var tipRegexMatch = matchTipRegex.Match(secLine);
                    if(!tipRegexMatch.Success || realMatch.Equals(default((string, string))))
                    {
                        SimpleLog.Warning("Tips could not be readed!");
                        Console.WriteLine($"Tips could not be readed. User {post.Username} in post {post.Url}. Do it yourself.");

                        user.Tips.Clear();
                        mustBreak = true;
                        break;
                    }

                    // save tips
                    var footballMatch = new FootballMatch
                    {
                        TeamA = realMatch.Item1,
                        TeamB = realMatch.Item2,
                        ResultA = Int32.Parse(tipRegexMatch.Groups[1].Value),
                        ResultB = Int32.Parse(tipRegexMatch.Groups[2].Value),
                    };
                    user.Tips.Add(footballMatch);
                }

                if(mustBreak) break;

                SimpleLog.Warning("Tips readed!");
                Console.WriteLine($"Tips from user {post.Username} in post {post.Url} readed.");
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
            var userMatchRegex = new Regex(@"([^\s]+)\s*(?:\((?:A|N)\))?\s*:\s([^\s]+)\s*(?:\((?:A|N)\))?");
            var regexMatches = userMatchRegex.Matches(content);
            if(string.IsNullOrWhiteSpace(content) || regexMatches.Count == 0)
            {
                SimpleLog.Warning("Could not find post with usermatches!");
                Console.WriteLine("Could not find post with usermatches. Please insert it manually.");
                return false;
            }

            // parse
            for(int i = 0; i < regexMatches.Count; i++)
            {
                var user1 = Userlist.FirstOrDefault(u => u.Name == regexMatches[i].Groups[1].Value);
                var user2 = Userlist.FirstOrDefault(u => u.Name == regexMatches[i].Groups[2].Value);

                if(user1 != null && user2 != null)
                    Usermatches.Add((user1, user2));
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
        public static List<(User, User)> ParsePosts(List<(string, string)> matches, List<User> userlist, List<(string Url, string Username, string Content)> posts)
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