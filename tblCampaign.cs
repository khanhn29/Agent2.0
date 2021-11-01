using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Agent2._0
{
    class tblCampaign
    {

        public tblCampaign(int id, DateTime createdate, DateTime lastupdate, string logpath)
        {
            this.Id = id;
            this.CreateDate = createdate;
            this.LastUpdate = lastupdate;
            this.LogPath = logpath;
        }

        public int Id { get; set; }
        public DateTime CreateDate { get; set; }
        public DateTime LastUpdate { get; set; }
        public string LogPath { get; set; }
    }
}
