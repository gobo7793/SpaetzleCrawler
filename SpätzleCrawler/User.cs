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

using System.Collections.Generic;
using System.Diagnostics;

namespace SpätzleCrawler
{
    /// <summary>
    /// Represents an user who tips a <see cref="FootballMatch"/>
    /// </summary>
    [DebuggerDisplay("{" + nameof(Name) + "}")]
    public class User
    {
        /// <summary>
        /// Count of tipping games
        /// </summary>
        public const int TippingGamesCount = 9;

        /// <summary>
        /// Usernick
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Tips of the user
        /// </summary>
        public List<FootballMatch> Tips { get; } = new List<FootballMatch>(9);

        /// <summary>
        /// Base column of user in the excel file
        /// </summary>
        public int UserCol { get; set; }
    }

    /// <summary>
    /// Represents a match between 2 <see cref="User"/>
    /// </summary>
    [DebuggerDisplay("{" + nameof(UserA) + "}-{" + nameof(UserB) + "}")]
    public class Usermatch
    {
        /// <summary>
        /// The first <see cref="User"/>
        /// </summary>
        public User UserA { get; set; }

        /// <summary>
        /// The second <see cref="User"/>
        /// </summary>
        public User UserB { get; set; }
    }
}
