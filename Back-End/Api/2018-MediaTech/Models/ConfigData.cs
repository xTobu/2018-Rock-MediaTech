using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace _2018_MediaTech.Models
{
    public class ConfigData
    {
        /// <summary>
        /// {type <int>,price<int>}
        /// </summary>
        public Dictionary<int, int> TicketType = new Dictionary<int, int>
        {
            // 早鳥 
            {0,6000},
            // 一般
            {1,8000},
            // 學生
            {2,2000} 
        };
    }
}