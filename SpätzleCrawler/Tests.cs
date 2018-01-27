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

using NUnit.Framework;

namespace SpätzleCrawler
{
    // For own test execution: setup tests!
    public class Tests
    {
        [Test]
        public void TestExcelReadUserList()
        {
            var settings = new Settings { TargetFileName = "" }; // insert file here

            var excel = new ExcelHandler();
            excel.OpenFile(settings.TargetFileName);
            var users = excel.ReadUserList();

            Assert.AreEqual(20, users.Count);
            Assert.AreEqual("", users[0].Name); // insert expected username here
            Assert.AreEqual("", users[7].Name); // insert expected username here
            Assert.AreEqual("", users[19].Name); // insert expected username here
        }

        [Test]
        public void TestExcelGetNextMatchday()
        {
            var settings = new Settings { TargetFileName = "" }; // insert file here

            var excel = new ExcelHandler();
            excel.OpenFile(settings.TargetFileName);
            var matchday = excel.ReadMatchdayNo();

            Assert.AreEqual(19, matchday); // insert expected matchday here
        }

        [Test]
        public void TestExcelGetNextMatchdayMatches()
        {
            var settings = new Settings { TargetFileName = "" }; // insert file here

            var excel = new ExcelHandler();
            excel.OpenFile(settings.TargetFileName);
            var matchday = excel.GetNextMatchdayMatches();

            Assert.AreEqual(9, matchday.Count);
            Assert.AreEqual(new FootballMatch { TeamA = "", TeamB = "" }, matchday[1]); // insert expected match here
            Assert.AreEqual(new FootballMatch { TeamA = "", TeamB = "" }, matchday[8]); // insert expected match here
        }

        [Test]
        public void TestCrawlerReadPages()
        {
            var settings = new Settings { TipThreadUrl = "" }; // insert url here

            var crawler = new Crawler { Settings = settings };
            var pageCnt = crawler.ReadPages();

            Assert.AreEqual(0, pageCnt); // insert expected page count here

        }

        [Test]
        public void TestCrawlerReadPosts()
        {
            var settings = new Settings { TipThreadUrl = "" }; // insert url here

            var crawler = new Crawler { Settings = settings };
            crawler.ReadPages();
            var postCnt = crawler.ReadPosts();

            Assert.AreEqual(0, postCnt); // insert expected post count here

        }
    }
}