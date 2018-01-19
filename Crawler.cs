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
using HtmlAgilityPack;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using SimpleLogger;

namespace SpätzleCrawler
{
    public class Crawler
    {
        /// <summary>
        /// Thread pages for tips
        /// </summary>
        public List<HtmlDocument> Pages { get; } = new List<HtmlDocument>();

        /// <summary>
        /// All posts in the thread
        /// </summary>
        public List<(string Url, string Username, string Content)> Posts { get; } = new List<(string, string, string)>();

        /// <summary>
        /// The TLD
        /// </summary>
        private string Domain { get; set; }

        /// <summary>
        /// Downloads the source of the whole tipping thread and returns the founded page count
        /// </summary>
        /// <param name="threadUrl">Thread URL</param>
        /// <returns>Founded pages count</returns>
        public int ReadPages(string threadUrl)
        {
            SimpleLog.Info("Read pages...");
            var uri = new Uri(threadUrl);
            Domain = $"{uri.Scheme}://{uri.Host}";

            var web = new HtmlWeb();
            var mainPage = web.Load(threadUrl);

            // Get page list
            var pager = mainPage.DocumentNode.Descendants("div").First(d => d.GetClasses().Contains("pager"));
            var pageUrls = pager.Descendants("li").Where(d => d.GetClasses().Contains("page"));

            // read pages
            foreach(var urlTag in pageUrls)
            {
                var path = urlTag.FirstChild.GetAttributeValue("href", String.Empty);
                var url = Domain + path;

                SimpleLog.Info($"Found page with url {url}.");

                Pages.Add(web.Load(url));
            }

            SimpleLog.Info($"{Pages.Count} pages found.");
            return Pages.Count;
        }

        /// <summary>
        /// Reads all posts and returns the founded posts count
        /// </summary>
        /// <returns>Founded posts count</returns>
        public int ReadPosts()
        {
            SimpleLog.Info("Read all posts...");

            foreach(var page in Pages)
            {
                var postSectionOuter = page.DocumentNode.Descendants("div").First(d => d.Id.Contains("postList"));
                var postSectionInner = postSectionOuter.Descendants("div").First(d => d.GetClasses().Contains("items"));
                //for(int i = 0; i < postSectionInner.ChildNodes.Count; i++)
                foreach(var postNode in postSectionInner.ChildNodes)
                {
                    if(postNode.Name != "div")
                        continue;

                    // read url
                    var postUrlSpan = postNode.Descendants("span").First(s => s.GetClasses().Contains("link-zum-post"));
                    var urlPath = postUrlSpan.FirstChild.GetAttributeValue("href", String.Empty);
                    var url = Domain + urlPath;

                    // read username
                    var usernameTag = postNode.Descendants("a").First(a => a.GetClasses().Contains("forum-user"));
                    var username = usernameTag.InnerText;
                    SimpleLog.Info($"Found post: Username: {username}, URL: {url}");

                    // read content
                    var contentTag = postNode.Descendants("div").First(d => d.GetClasses().Contains("forum-post-data"));
                    var removeTag = contentTag.ChildNodes.FirstOrDefault(c => c.Name == "div");
                    if(removeTag != null)
                        contentTag.RemoveChild(removeTag);

                    Posts.Add((url, username, contentTag.InnerText.Trim()));
                }
            }

            return Posts.Count;
        }
    }
}