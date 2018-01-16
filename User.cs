using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpätzleCrawler
{
    /// <summary>
    /// Represents an user who tips a <see cref="FootballMatch"/>
    /// </summary>
    class User
    {
        /// <summary>
        /// Usernick
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Tips of the user
        /// </summary>
        public List<FootballMatch> Tips { get; } = new List<FootballMatch>(9);
    }
}
