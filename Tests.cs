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
using NUnit.Framework;

namespace SpätzleCrawler
{
    // For own test execution: setup tests!
    public class Tests
    {
        [Test]
        public void TestExcelReadUserList()
        {
            var filename = ""; // insert file here

            var excel = new ExcelHandler();
            excel.OpenFile(filename);
            var users = excel.ReadUserList();

            Assert.AreEqual(20, users.Count);
            Assert.AreEqual("", users[0]); // insert expected username here
            Assert.AreEqual("", users[7]); // insert expected username here
            Assert.AreEqual("", users[19]); // insert expected username here
        }

        [Test]
        public void TestExcelGetNextMatchday()
        {
            var filename = ""; // insert file here

            var excel = new ExcelHandler();
            excel.OpenFile(filename);
            var matchday = excel.GetNextMatchday();

            Assert.AreEqual(19, matchday); // insert expected matchday here
        }

        [Test]
        public void TestCrawlerReadPages()
        {
            var url =""; // insert url here

            var crawler = new Crawler();
            var pageCnt = crawler.ReadPages(url);

            Assert.AreEqual(0, pageCnt); // insert expected page count here

        }

        [Test]
        public void TestCrawlerReadPosts()
        {
            var url = ""; // insert url here

            var crawler = new Crawler();
            crawler.ReadPages(url);
            var postCnt = crawler.ReadPosts();

            Assert.AreEqual(0, postCnt); // insert expected post count here

        }
    }
}