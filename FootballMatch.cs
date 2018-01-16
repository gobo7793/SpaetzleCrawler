using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpätzleCrawler
{
    /// <summary>
    /// Represents a football match
    /// </summary>
    class FootballMatch
    {
        /// <summary>
        /// Home team
        /// </summary>
        public string TeamA { get; set; }

        /// <summary>
        /// Away team
        /// </summary>
        public string TeamB { get; set; }

        /// <summary>
        /// Goals home team
        /// </summary>
        public int ResultA { get; set; }

        /// <summary>
        /// Goals away team
        /// </summary>
        public int ResultB { get; set; }
    }
}
