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
    // For own test execution: setup tests for own excel file!
    // It's based on the file for 2017-18 part 2 league 4!
    public class Tests
    {
        [Test]
        public void TestExcelReadUserList()
        {
            var filename = @"E:\Dokumente\VS Projects\SpaetzleCrawler\spaetzle2018-1.xlsx";

            var excel = new ExcelHandler();
            excel.OpenFile(filename);
            var users = excel.ReadUserList();

            Assert.AreEqual(20, users.Count);
            Assert.AreEqual("Valbatorix", users[0]);
            Assert.AreEqual("lars-gutsein", users[7]);
            Assert.AreEqual("Frei", users[19]);
        }

        [Test]
        public void TestExcelGetNextMatchday()
        {
            var filename = @"E:\Dokumente\VS Projects\SpaetzleCrawler\spaetzle2018-1.xlsx";

            var excel = new ExcelHandler();
            excel.OpenFile(filename);
            var matchday = excel.GetNextMatchday();

            Assert.AreEqual(19, matchday);
        }
    }
}