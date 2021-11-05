using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Agent2._0
{
    class tblCampaign
    {

        public tblCampaign(int id, DateTime fromdate, DateTime todate, string logpath)
        {
            this.Id = id;
            this.FromDate = fromdate;
            this.ToDate = todate;
            this.LogPath = logpath;
        }

        public int Id { get; set; }
        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }
        public string LogPath { get; set; }
    }
}
